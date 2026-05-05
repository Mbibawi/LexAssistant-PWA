//@ts-nocheck
// DocumentContext.ts
// Bridges Office.js (Word / Excel) with the Lex Assistant prompt pipeline.
// Requires: office.js loaded in the host page (Office Add-in context).
// Works transparently as a no-op when running as a standalone PWA.


// ─────────────────────────────────────────────────────────────────────────────
import { Folders } from "./onedrive";

const hosts = {
    word: 'Word' as const,
    excel: 'Excel' as const,
    standalone: 'Standalone' as const
};


export class DocumentContext {
    private host: typeof hosts[keyof typeof hosts] = hosts.standalone;
    private isReady = false;
    private readJson:<T>(path:string)=>Promise<T | null>;
    private writeJson:(path:string,data:any)=>Promise<void>;

    // OneDrive path for cross-app context file (matches your Folders convention)
    private readonly CONTEXT_FILE = '_LexContext/OfficeShareSession.json';

    // Max characters of full document body sent to Claude (token guard)
    private readonly BODY_CHAR_LIMIT = 12_000;



    // ── Initialisation ──────────────────────────────────────────────────────────

    /**
     * Call once at app startup, before any other method.
     * Resolves when Office.js is ready (or immediately when standalone).
     */
    async init() {
        if (typeof Office === 'undefined') {
            this.host = hosts.standalone;
            this.isReady = true;
            return this.host;
        }

        return new Promise((resolve) => {
            Office.onReady((info) => {
                if (info.host === Office.HostType.Word) {
                    this.host = hosts.word;
                } else if (info.host === Office.HostType.Excel) {
                    this.host = hosts.excel;
                } else {
                    this.host = hosts.standalone;
                }
                this.isReady = true;
                resolve(this.host);
            });
        });
    }

    getHost() {
        return this.host;
    }

    isOfficeContext(): boolean {
        return [hosts.word, hosts.excel].includes( this.host);
    }

    // ── Selection ───────────────────────────────────────────────────────────────

    /**
     * Read whatever the user has selected in the active document.
     * Returns null if nothing is selected or context is Standalone.
     */
    async getSelection(): Promise<DocumentSelection | null> {
        if (!this.isReady) throw new Error('DocumentContext.init() not called');

        if (this.host === hosts.word) return this._getWordSelection();
        if (this.host === hosts.excel) return this._getExcelSelection();
        return null;
    }

    private _getWordSelection(): Promise<DocumentSelection | null> {
        return new Promise((resolve) => {
            // Get plain text first
            Office.context.document.getSelectedDataAsync(
                Office.CoercionType.Text,
                (textResult) => {
                    if (
                        textResult.status !== Office.AsyncResultStatus.Succeeded ||
                        !textResult.value?.trim()
                    ) {
                        resolve(null);
                        return;
                    }
                    const text = textResult.value.trim();

                    // Also get OOXML for structure-aware operations (tracked changes etc.)
                    Office.context.document.getSelectedDataAsync(
                        Office.CoercionType.Ooxml,
                        (xmlResult) => {
                            resolve({
                                text,
                                ooxml:
                                    xmlResult.status === Office.AsyncResultStatus.Succeeded
                                        ? xmlResult.value
                                        : undefined,
                            });
                        }
                    );
                }
            );
        });
    }

    private _getExcelSelection(): Promise<DocumentSelection | null> {
        return new Promise((resolve) => {
            Office.context.document.getSelectedDataAsync(
                Office.CoercionType.Text,
                (result) => {
                    if (result.status !== Office.AsyncResultStatus.Succeeded) {
                        resolve(null);
                        return;
                    }
                    // For Excel, also capture the range address
                    Excel.run(async (ctx) => {
                        const range = ctx.workbook.getSelectedRange();
                        range.load('address,values');
                        await ctx.sync();
                        resolve({
                            text: result.value ?? '',
                            rangeAddress: range.address,
                        });
                    }).catch(() => {
                        resolve({ text: result.value ?? '' });
                    });
                }
            );
        });
    }

    // ── Document snapshot ────────────────────────────────────────────────────────

    /**
     * Build a DocumentSnapshot — what you pass to your Claude prompt builder.
     * Controls how much text is sent; never sends the whole document blindly.
     *
     * @param includeFullBody  If true, includes trimmed body text up to BODY_CHAR_LIMIT.
     *                         Default false — only headings + selection are included.
     */
    async getSnapshot(includeFullBody = false): Promise<DocumentSnapshot> {
        const base: DocumentSnapshot = {
            host: this.host,
            title: await this._getDocumentTitle(),
            timestamp: new Date().toISOString(),
        };

        if (this.host === hosts.word) {
            base.headings = await this._getWordHeadings();
            base.selection = (await this._getWordSelection()) ?? undefined;
            if (includeFullBody) {
                base.fullText = await this._getWordBodyText();
            }
        }

        if (this.host === hosts.excel) {
            base.sheetSummary = await this._getExcelSheetSummary();
            base.selection = (await this._getExcelSelection()) ?? undefined;
        }

        return base;
    }

    /**
     * Formats a snapshot as a compact context block to prepend to Claude prompts.
     * Keeps token cost predictable — never dumps the full document unless asked.
     */
    buildPromptContext(snapshot: DocumentSnapshot): string {
        const lines: string[] = [
            `[Document: ${snapshot.title} | Host: ${snapshot.host} | ${snapshot.timestamp}]`,
        ];

        if (snapshot.headings?.length) {
            lines.push('\n## Document Structure (Headings)');
            lines.push(snapshot.headings.join('\n'));
        }

        if (snapshot.sheetSummary) {
            lines.push(`\n## Active Sheet: ${snapshot.sheetSummary}`);
        }

        if (snapshot.selection?.text) {
            lines.push('\n## Current Selection');
            lines.push(snapshot.selection.text);
            if (snapshot.selection.rangeAddress) {
                lines.push(`(Range: ${snapshot.selection.rangeAddress})`);
            }
        }

        if (snapshot.fullText) {
            lines.push('\n## Document Body (excerpt)');
            lines.push(snapshot.fullText);
        }

        return lines.join('\n');
    }

    // ── Write back ───────────────────────────────────────────────────────────────

    /**
     * Insert Claude's output back into the document.
     * In Word: lands as a tracked change by default.
     * In Excel: writes to the selected cell(s).
     */
    async writeBack(text: string, options: WriteBackOptions = {}): Promise<void> {
        const { asTrackedChange = true, replaceSelection = true } = options;

        if (this.host === hosts.word) {
            await this._writeWordTrackedChange(text, asTrackedChange, replaceSelection);
        } else if (this.host === hosts.excel) {
            await this._writeExcelSelection(text, replaceSelection);
        }
        // Standalone: no-op (caller displays text in its own UI)
    }

    private async _writeWordTrackedChange(
        text: string,
        tracked: boolean,
        replace: boolean
    ): Promise<void> {
        await Word.run(async (context) => {
            if (tracked) {
                // Enable tracked changes so Claude's edits appear as revisions
                context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            }

            const selection = context.document.getSelection();
            selection.insertText(
                text,
                replace ? Word.InsertLocation.replace : Word.InsertLocation.after
            );

            await context.sync();

            if (tracked) {
                // Return to default after write so user edits aren't auto-tracked
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                await context.sync();
            }
        });
    }

    private async _writeExcelSelection(
        text: string,
        replace: boolean
    ): Promise<void> {
        return new Promise((resolve, reject) => {
            Office.context.document.setSelectedDataAsync(
                text,
                { coercionType: Office.CoercionType.Text },
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        reject(new Error(result.error.message));
                    } else {
                        resolve();
                    }
                }
            );
        });
    }

    // ── Cross-app context via OneDrive ───────────────────────────────────────────
    // Uses your existing OneDrive Graph API layer (graphFetch / uploadFile pattern)

    /**
     * Save current snapshot to OneDrive so the Excel (or Word) instance
     * of Lex Assistant can pick it up.
     *
     * @param graphFetch  Your existing authenticated Graph API fetch wrapper
     * @param history     Current conversation turns to persist
     */
    async saveSharedContext(history: CrossAppContext['sharedHistory'] = []
    ): Promise<void> {
        const snapshot = await this.getSnapshot(false);

        let existing: CrossAppContext | null = null;
        try {
            existing = await this.readJson<CrossAppContext | null>(this.CONTEXT_FILE);
        } catch {
            // File doesn't exist yet — that's fine
        }

        const sessionId = existing?.sessionId ?? crypto.randomUUID();
        const updated: CrossAppContext = {
            sessionId,
            wordSnapshot: this.host === hosts.word ? snapshot : existing?.wordSnapshot,
            excelSnapshot:
                this.host === hosts.excel ? snapshot : existing?.excelSnapshot,
            sharedHistory: history,
            updatedAt: new Date().toISOString(),
        };
        await this.writeJson(this.CONTEXT_FILE, JSON.stringify(updated, null, 2));
    }

    /**
     * Load the shared cross-app context from OneDrive.
     * Call this on startup in Excel to pick up context written by Word, or vice versa.
     */
    async loadSharedContext(
        graphFetch: (path: string, options?: RequestInit) => Promise<Response>
    ): Promise<CrossAppContext | null> {
        try {
            const res = await graphFetch(
                `/me/drive/root:/${this.CONTEXT_FILE}:/content`
            );
            if (!res.ok) return null;
            return await res.json();
        } catch {
            return null;
        }
    }

    // ── Private helpers ──────────────────────────────────────────────────────────

    private _getDocumentTitle(): Promise<string> {
        if (this.host === hosts.word) {
            return Word.run(async (ctx) => {
                const props = ctx.document.properties;
                props.load('title');
                await ctx.sync();
                return props.title || 'Untitled';
            }).catch(() => 'Untitled');
        }

        if (this.host === hosts.excel) {
            return Excel.run(async (ctx) => {
                const props = ctx.workbook.properties;
                props.load('title');
                await ctx.sync();
                return props.title || 'Untitled Workbook';
            }).catch(() => 'Untitled Workbook');
        }

        return Promise.resolve('Standalone');
    }

    private _getWordHeadings(): Promise<string[]> {
        return Word.run(async (ctx) => {
            const headings = ctx.document.body.paragraphs;
            headings.load('text,style');
            await ctx.sync();

            return headings.items
                .filter((p) =>
                    ['Heading 1', 'Heading 2', 'Heading 3'].includes(p.style)
                )
                .map((p) => {
                    const indent =
                        p.style === 'Heading 1'
                            ? ''
                            : p.style === 'Heading 2'
                                ? '  '
                                : '    ';
                    return `${indent}${p.text.trim()}`;
                });
        }).catch(() => []);
    }

    private _getWordBodyText(): Promise<string> {
        return Word.run(async (ctx) => {
            const body = ctx.document.body;
            body.load('text');
            await ctx.sync();
            const text = body.text ?? '';
            // Hard cap: never send more than BODY_CHAR_LIMIT to Claude
            return text.length > this.BODY_CHAR_LIMIT
                ? text.slice(0, this.BODY_CHAR_LIMIT) +
                `\n\n[... document truncated at ${this.BODY_CHAR_LIMIT} chars ...]`
                : text;
        }).catch(() => '');
    }

    private _getExcelSheetSummary(): Promise<string> {
        return Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange(true);
            sheet.load('name');
            usedRange.load('address');
            await ctx.sync();
            return `${sheet.name} (used range: ${usedRange.address})`;
        }).catch(() => 'Unknown sheet');
    }
}
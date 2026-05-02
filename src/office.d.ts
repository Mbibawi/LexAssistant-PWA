// src/office.d.ts
// NO imports, NO exports — this must be a pure ambient script file

declare namespace Office {
    // ...
}

declare namespace Word {
    // ...
}

declare namespace Excel {
    // ...
}

type HostType = 'Word' | 'Excel' | 'Standalone';

type DocumentSelection = {
    text: string;
    ooxml?: string;
};

type DocumentSnapshot = {
    host: HostType;
    title: string;
    timestamp: string;
    headings?: string[];
    selectedText?: string;
    selection?: {
        start: number;
        end: number;
    };
    fullBody?: string;
    rangeAddress?: string;
};

type WriteBackOptions = {
    onSuccess?: () => void;
    onError?: (error: Error) => void;
    writeBackMode?: WriteBackMode;
}

type CrossAppContext = {
    sessionId?: string;
    wordSnapshot?: DocumentSnapshot;
    excelSnapshot?: DocumentSnapshot;
    host?: HostType;
    title?: string;
    timestamp?: string;
    headings?: string[];
    selectedText?: string;
    selection?: {
        start: number;
        end: number;
    };
    fullBody?: string;
    rangeAddress?: string;
    sharedHistory?: LexMessage[];
    updatedAt?: string;
    version?: string;
}
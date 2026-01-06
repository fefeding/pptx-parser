import { BaseElement } from './BaseElement';
import type { NotesSlideResult, NotesMasterResult } from '../core/types';
export declare class NotesMasterElement extends BaseElement {
    type: 'notesMaster';
    elements: BaseElement[];
    placeholders: any[];
    background?: {
        type: 'color' | 'image' | 'none';
        value?: string;
        relId?: string;
    };
    constructor(id: string, elements?: BaseElement[], placeholders?: any[], props?: any);
    static fromResult(result: NotesMasterResult): NotesMasterElement;
    toHTML(): string;
    private getBackgroundStyle;
}
export declare class NotesSlideElement extends BaseElement {
    type: 'notesSlide';
    text?: string;
    masterRef?: string;
    master?: NotesMasterElement;
    slideId?: string;
    constructor(id: string, rect: {
        x: number;
        y: number;
        width: number;
        height: number;
    }, props?: any);
    static fromResult(result: NotesSlideResult): NotesSlideElement;
    toHTML(): string;
    setMaster(master: NotesMasterElement): void;
}

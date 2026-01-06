import JSZip from 'jszip';
import type { NotesMasterResult, NotesSlideResult } from './types';
export declare function parseAllNotesMasters(zip: JSZip): Promise<NotesMasterResult[]>;
export declare function parseAllNotesSlides(zip: JSZip): Promise<NotesSlideResult[]>;
export declare function linkNotesToMasters(notesSlides: NotesSlideResult[], masters: NotesMasterResult[]): void;

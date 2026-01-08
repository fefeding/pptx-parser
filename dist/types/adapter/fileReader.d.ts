import { EnvType } from "../types/index";
export declare const detectEnv: () => EnvType;
export declare const readPptxFile: (source: string | ArrayBuffer | Uint8Array, env?: EnvType) => Promise<Uint8Array>;

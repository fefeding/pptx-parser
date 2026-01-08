import { EnvType } from "../types/index";
export declare const detectEnv: () => EnvType;
export declare const readPptxFile: (source: string | Buffer, env?: EnvType) => Promise<Buffer>;

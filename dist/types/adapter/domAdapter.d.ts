import { EnvType } from "../types/index";
export declare class DomAdapter {
    private env;
    constructor(env?: EnvType);
    createElement(tag: string, attrs?: Record<string, string>, content?: string): HTMLElement | string;
    mount(el: HTMLElement | string, targetId: string): void | string;
}

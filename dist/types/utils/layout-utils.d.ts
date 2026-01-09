export declare const ALIGNMENT_CLASSES: {
    readonly v_up: "v-up";
    readonly v_mid: "v-mid";
    readonly v_down: "v-down";
    readonly h_left: "h-left";
    readonly h_mid: "h-mid";
    readonly h_right: "h-right";
    readonly up_center: "up-center";
    readonly up_left: "up-left";
    readonly up_right: "up-right";
    readonly center_center: "center-center";
    readonly center_left: "center-left";
    readonly center_right: "center-right";
    readonly down_center: "down-center";
    readonly down_left: "down-left";
    readonly down_right: "down-right";
};
export declare function getVerticalAlignClass(vAlign?: 'top' | 'middle' | 'bottom'): string;
export declare function getHorizontalAlignClass(hAlign?: 'left' | 'center' | 'right'): string;
export declare function getAlignmentClass(hAlign?: 'left' | 'center' | 'right', vAlign?: 'top' | 'middle' | 'bottom'): string;
export declare function getPlaceholderLayoutStyle(placeholder: {
    rect: {
        x: number;
        y: number;
        width: number;
        height: number;
    };
    hAlign?: 'left' | 'center' | 'right';
    vAlign?: 'top' | 'middle' | 'bottom';
}): {
    style: string;
    className: string;
};
export declare function getSlideContainerStyle(width: number, height: number, background?: {
    type: 'color' | 'image' | 'none';
    value?: string;
    relId?: string;
}, scale?: number): string;
export declare function findPlaceholder(placeholders: Array<{
    type?: string | 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';
    idx?: number;
} & Record<string, any>>, type: string, idx?: number): (Record<string, any> & {
    type?: string | 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';
    idx?: number;
}) | undefined;
export declare function mergePlaceholderStyles(baseStyle: Record<string, any>, inheritedStyle: Record<string, any>): Record<string, any>;

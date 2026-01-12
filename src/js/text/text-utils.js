/**
 * text-utils.js
 * 文本处理工具模块
 * 负责文本相关的工具函数，如 dingbat 字符映射等
 */

(function () {
    var TextUtils = {};

    // Dingbat unicode mapping for special font characters
    var dingbat_unicode = [
        // Wingdings 2
        {f: "Wingdings 2", code: 0x001, unicode: 0x2713},
        {f: "Wingdings 2", code: 0x002, unicode: 0x2714},
        {f: "Wingdings 2", code: 0x003, unicode: 0x2715},
        {f: "Wingdings 2", code: 0x004, unicode: 0x2716},
        {f: "Wingdings 2", code: 0x005, unicode: 0x2717},
        {f: "Wingdings 2", code: 0x006, unicode: 0x2718},
        {f: "Wingdings 2", code: 0x007, unicode: 0x2719},
        {f: "Wingdings 2", code: 0x008, unicode: 0x271A},
        {f: "Wingdings 2", code: 0x009, unicode: 0x271B},
        {f: "Wingdings 2", code: 0x00A, unicode: 0x271C},
        {f: "Wingdings 2", code: 0x00B, unicode: 0x271D},
        {f: "Wingdings 2", code: 0x00C, unicode: 0x271E},
        {f: "Wingdings 2", code: 0x00D, unicode: 0x271F},
        {f: "Wingdings 2", code: 0x00E, unicode: 0x2720},
        {f: "Wingdings 2", code: 0x00F, unicode: 0x2721},
        {f: "Wingdings 2", code: 0x010, unicode: 0x2722},
        {f: "Wingdings 2", code: 0x011, unicode: 0x2723},
        {f: "Wingdings 2", code: 0x012, unicode: 0x2724},
        {f: "Wingdings 2", code: 0x013, unicode: 0x2725},
        {f: "Wingdings 2", code: 0x014, unicode: 0x2726},
        {f: "Wingdings 2", code: 0x015, unicode: 0x2727},
        {f: "Wingdings 2", code: 0x016, unicode: 0x2728},
        {f: "Wingdings 2", code: 0x017, unicode: 0x2729},
        {f: "Wingdings 2", code: 0x018, unicode: 0x272A},
        {f: "Wingdings 2", code: 0x019, unicode: 0x272B},
        {f: "Wingdings 2", code: 0x01A, unicode: 0x272C},
        {f: "Wingdings 2", code: 0x01B, unicode: 0x272D},
        {f: "Wingdings 2", code: 0x01C, unicode: 0x272E},
        {f: "Wingdings 2", code: 0x01D, unicode: 0x272F},
        {f: "Wingdings 2", code: 0x01E, unicode: 0x2730},
        {f: "Wingdings 2", code: 0x01F, unicode: 0x2731},
        {f: "Wingdings 2", code: 0x020, unicode: 0x2732},
        {f: "Wingdings 2", code: 0x021, unicode: 0x2733},
        {f: "Wingdings 2", code: 0x022, unicode: 0x2734},
        {f: "Wingdings 2", code: 0x023, unicode: 0x2735},
        {f: "Wingdings 2", code: 0x024, unicode: 0x2736},
        {f: "Wingdings 2", code: 0x025, unicode: 0x2737},
        {f: "Wingdings 2", code: 0x026, unicode: 0x2738},
        {f: "Wingdings 2", code: 0x027, unicode: 0x2739},
        {f: "Wingdings 2", code: 0x028, unicode: 0x273A},
        {f: "Wingdings 2", code: 0x029, unicode: 0x273B},
        {f: "Wingdings 2", code: 0x02A, unicode: 0x273C},
        {f: "Wingdings 2", code: 0x02B, unicode: 0x273D},
        {f: "Wingdings 2", code: 0x02C, unicode: 0x273E},
        {f: "Wingdings 2", code: 0x02D, unicode: 0x273F},
        {f: "Wingdings 2", code: 0x02E, unicode: 0x2740},
        {f: "Wingdings 2", code: 0x02F, unicode: 0x2741},
        {f: "Wingdings 2", code: 0x030, unicode: 0x2742},
        {f: "Wingdings 2", code: 0x031, unicode: 0x2743},
        {f: "Wingdings 2", code: 0x032, unicode: 0x2744},
        {f: "Wingdings 2", code: 0x033, unicode: 0x2745},
        {f: "Wingdings 2", code: 0x034, unicode: 0x2746},
        {f: "Wingdings 2", code: 0x035, unicode: 0x2747},
        {f: "Wingdings 2", code: 0x036, unicode: 0x2748},
        {f: "Wingdings 2", code: 0x037, unicode: 0x2749},
        {f: "Wingdings 2", code: 0x038, unicode: 0x274A},
        {f: "Wingdings 2", code: 0x039, unicode: 0x274B},
        {f: "Wingdings 2", code: 0x03A, unicode: 0x274C},
        {f: "Wingdings 2", code: 0x03B, unicode: 0x274D},
        {f: "Wingdings 2", code: 0x03C, unicode: 0x274E},
        {f: "Wingdings 2", code: 0x03D, unicode: 0x274F},
        {f: "Wingdings 2", code: 0x03E, unicode: 0x2750},
        {f: "Wingdings 2", code: 0x03F, unicode: 0x2751},
        {f: "Wingdings 2", code: 0x040, unicode: 0x2752},
        {f: "Wingdings 2", code: 0x041, unicode: 0x2753},
        {f: "Wingdings 2", code: 0x042, unicode: 0x2754},
        {f: "Wingdings 2", code: 0x043, unicode: 0x2755},
        {f: "Wingdings 2", code: 0x044, unicode: 0x2756},
        {f: "Wingdings 2", code: 0x045, unicode: 0x2757},
        {f: "Wingdings 2", code: 0x046, unicode: 0x2758},
        {f: "Wingdings 2", code: 0x047, unicode: 0x2759},
        {f: "Wingdings 2", code: 0x048, unicode: 0x275A},
        {f: "Wingdings 2", code: 0x049, unicode: 0x275B},
        {f: "Wingdings 2", code: 0x04A, unicode: 0x275C},
        {f: "Wingdings 2", code: 0x04B, unicode: 0x275D},
        {f: "Wingdings 2", code: 0x04C, unicode: 0x275E},
        {f: "Wingdings 2", code: 0x04D, unicode: 0x275F},
        {f: "Wingdings 2", code: 0x04E, unicode: 0x2760},
        {f: "Wingdings 2", code: 0x04F, unicode: 0x2761},
        {f: "Wingdings 2", code: 0x050, unicode: 0x2762},
        {f: "Wingdings 2", code: 0x051, unicode: 0x2763},
        {f: "Wingdings 2", code: 0x052, unicode: 0x2764},
        {f: "Wingdings 2", code: 0x053, unicode: 0x2765},
        {f: "Wingdings 2", code: 0x054, unicode: 0x2766},
        {f: "Wingdings 2", code: 0x055, unicode: 0x2767},
        {f: "Wingdings 2", code: 0x056, unicode: 0x2768},
        {f: "Wingdings 2", code: 0x057, unicode: 0x2769},
        {f: "Wingdings 2", code: 0x058, unicode: 0x276A},
        {f: "Wingdings 2", code: 0x059, unicode: 0x276B},
        {f: "Wingdings 2", code: 0x05A, unicode: 0x276C},
        {f: "Wingdings 2", code: 0x05B, unicode: 0x276D},
        {f: "Wingdings 2", code: 0x05C, unicode: 0x276E},
        {f: "Wingdings 2", code: 0x05D, unicode: 0x276F},
        {f: "Wingdings 2", code: 0x05E, unicode: 0x2770},
        {f: "Wingdings 2", code: 0x05F, unicode: 0x2771},
        {f: "Wingdings 2", code: 0x060, unicode: 0x2772},
        {f: "Wingdings 2", code: 0x061, unicode: 0x2773},
        {f: "Wingdings 2", code: 0x062, unicode: 0x2774},
        {f: "Wingdings 2", code: 0x063, unicode: 0x2775},
        {f: "Wingdings 2", code: 0x064, unicode: 0x2794},
        {f: "Wingdings 2", code: 0x065, unicode: 0x2798},
        {f: "Wingdings 2", code: 0x066, unicode: 0x2799},
        {f: "Wingdings 2", code: 0x067, unicode: 0x279A},
        {f: "Wingdings 2", code: 0x068, unicode: 0x279B},
        {f: "Wingdings 2", code: 0x069, unicode: 0x279C},
        {f: "Wingdings 2", code: 0x06A, unicode: 0x279D},
        {f: "Wingdings 2", code: 0x06B, unicode: 0x279E},
        {f: "Wingdings 2", code: 0x06C, unicode: 0x279F},
        {f: "Wingdings 2", code: 0x06D, unicode: 0x27A0},
        {f: "Wingdings 2", code: 0x06E, unicode: 0x27A1},
        {f: "Wingdings 2", code: 0x06F, unicode: 0x27A2},
        {f: "Wingdings 2", code: 0x070, unicode: 0x27A3},
        {f: "Wingdings 2", code: 0x071, unicode: 0x27A4},
        {f: "Wingdings 2", code: 0x072, unicode: 0x27A5},
        {f: "Wingdings 2", code: 0x073, unicode: 0x27A6},
        {f: "Wingdings 2", code: 0x074, unicode: 0x27A7},
        {f: "Wingdings 2", code: 0x075, unicode: 0x27A8},
        {f: "Wingdings 2", code: 0x076, unicode: 0x27A9},
        {f: "Wingdings 2", code: 0x077, unicode: 0x27AA},
        {f: "Wingdings 2", code: 0x078, unicode: 0x27AB},
        {f: "Wingdings 2", code: 0x079, unicode: 0x27AC},
        {f: "Wingdings 2", code: 0x07A, unicode: 0x27AD},
        {f: "Wingdings 2", code: 0x07B, unicode: 0x27AE},
        {f: "Wingdings 2", code: 0x07C, unicode: 0x27AF},
        {f: "Wingdings 2", code: 0x07D, unicode: 0x27B0},
        {f: "Wingdings 2", code: 0x07E, unicode: 0x27B1},
        {f: "Wingdings 2", code: 0x07F, unicode: 0x27B2},
        {f: "Wingdings 2", code: 0x080, unicode: 0x27B3},
        {f: "Wingdings 2", code: 0x081, unicode: 0x27B4},
        {f: "Wingdings 2", code: 0x082, unicode: 0x27B5},
        {f: "Wingdings 2", code: 0x083, unicode: 0x27B6},
        {f: "Wingdings 2", code: 0x084, unicode: 0x27B7},
        {f: "Wingdings 2", code: 0x085, unicode: 0x27B8},
        {f: "Wingdings 2", code: 0x086, unicode: 0x27B9},
        {f: "Wingdings 2", code: 0x087, unicode: 0x27BA},
        {f: "Wingdings 2", code: 0x088, unicode: 0x27BB},
        {f: "Wingdings 2", code: 0x089, unicode: 0x27BC},
        {f: "Wingdings 2", code: 0x08A, unicode: 0x27BD},
        {f: "Wingdings 2", code: 0x08B, unicode: 0x27BE},
        {f: "Wingdings 2", code: 0x08C, unicode: 0x27F3},
        {f: "Wingdings 2", code: 0x08D, unicode: 0x27F4},
        {f: "Wingdings 2", code: 0x08E, unicode: 0x27F5},
        {f: "Wingdings 2", code: 0x08F, unicode: 0x27F6},
        {f: "Wingdings 2", code: 0x090, unicode: 0x27F7},
        {f: "Wingdings 2", code: 0x091, unicode: 0x27F8},
        {f: "Wingdings 2", code: 0x092, unicode: 0x27F9},
        {f: "Wingdings 2", code: 0x093, unicode: 0x27FA},
        {f: "Wingdings 2", code: 0x094, unicode: 0x27FB},
        {f: "Wingdings 2", code: 0x095, unicode: 0x27FC},
        {f: "Wingdings 2", code: 0x096, unicode: 0x27FD},
        {f: "Wingdings 2", code: 0x097, unicode: 0x27FE},
        {f: "Wingdings 2", code: 0x097, unicode: 0x27FF},
        // Wingdings 3
        {f: "Wingdings 3", code: 0x021, unicode: 0x2761},
        {f: "Wingdings 3", code: 0x022, unicode: 0x2762},
        {f: "Wingdings 3", code: 0x023, unicode: 0x2763},
        {f: "Wingdings 3", code: 0x024, unicode: 0x2764},
        {f: "Wingdings 3", code: 0x025, unicode: 0x2765},
        {f: "Wingdings 3", code: 0x026, unicode: 0x2766},
        {f: "Wingdings 3", code: 0x027, unicode: 0x2767},
        {f: "Wingdings 3", code: 0x028, unicode: 0x267F},
        {f: "Wingdings 3", code: 0x029, unicode: 0x2680},
        {f: "Wingdings 3", code: 0x02A, unicode: 0x2681},
        {f: "Wingdings 3", code: 0x02B, unicode: 0x2682},
        {f: "Wingdings 3", code: 0x02C, unicode: 0x2683},
        {f: "Wingdings 3", code: 0x02D, unicode: 0x2684},
        {f: "Wingdings 3", code: 0x02E, unicode: 0x2685},
        {f: "Wingdings 3", code: 0x02F, unicode: 0x2686},
        {f: "Wingdings 3", code: 0x030, unicode: 0x2687},
        {f: "Wingdings 3", code: 0x031, unicode: 0x2688},
        {f: "Wingdings 3", code: 0x032, unicode: 0x2689},
        {f: "Wingdings 3", code: 0x033, unicode: 0x268A},
        {f: "Wingdings 3", code: 0x034, unicode: 0x268B},
        {f: "Wingdings 3", code: 0x035, unicode: 0x268C},
        {f: "Wingdings 3", code: 0x036, unicode: 0x268D},
        {f: "Wingdings 3", code: 0x037, unicode: 0x268E},
        {f: "Wingdings 3", code: 0x038, unicode: 0x268F},
        {f: "Wingdings 3", code: 0x039, unicode: 0x2690},
        {f: "Wingdings 3", code: 0x03A, unicode: 0x2691},
        {f: "Wingdings 3", code: 0x03B, unicode: 0x2692},
        {f: "Wingdings 3", code: 0x03C, unicode: 0x2693},
        {f: "Wingdings 3", code: 0x03D, unicode: 0x2694},
        {f: "Wingdings 3", code: 0x03E, unicode: 0x2695},
        {f: "Wingdings 3", code: 0x03F, unicode: 0x2696},
        {f: "Wingdings 3", code: 0x040, unicode: 0x2697},
        {f: "Wingdings 3", code: 0x041, unicode: 0x2698},
        {f: "Wingdings 3", code: 0x042, unicode: 0x2699},
        {f: "Wingdings 3", code: 0x043, unicode: 0x269A},
        {f: "Wingdings 3", code: 0x044, unicode: 0x269B},
        {f: "Wingdings 3", code: 0x045, unicode: 0x269C},
        {f: "Wingdings 3", code: 0x046, unicode: 0x269D},
        {f: "Wingdings 3", code: 0x047, unicode: 0x269E},
        {f: "Wingdings 3", code: 0x048, unicode: 0x269F},
        {f: "Wingdings 3", code: 0x049, unicode: 0x26A0},
        {f: "Wingdings 3", code: 0x04A, unicode: 0x26A1},
        {f: "Wingdings 3", code: 0x04B, unicode: 0x26A2},
        {f: "Wingdings 3", code: 0x04C, unicode: 0x26A3},
        {f: "Wingdings 3", code: 0x04D, unicode: 0x26A4},
        {f: "Wingdings 3", code: 0x04E, unicode: 0x26A5},
        {f: "Wingdings 3", code: 0x04F, unicode: 0x26A6},
        {f: "Wingdings 3", code: 0x050, unicode: 0x26A7},
        {f: "Wingdings 3", code: 0x051, unicode: 0x26A8},
        {f: "Wingdings 3", code: 0x052, unicode: 0x26A9},
        {f: "Wingdings 3", code: 0x053, unicode: 0x26AA},
        {f: "Wingdings 3", code: 0x054, unicode: 0x26AB},
        {f: "Wingdings 3", code: 0x055, unicode: 0x26AC},
        {f: "Wingdings 3", code: 0x056, unicode: 0x26AD},
        {f: "Wingdings 3", code: 0x057, unicode: 0x26AE},
        {f: "Wingdings 3", code: 0x058, unicode: 0x26AF},
        {f: "Wingdings 3", code: 0x059, unicode: 0x26B0},
        {f: "Wingdings 3", code: 0x05A, unicode: 0x26B1},
        {f: "Wingdings 3", code: 0x05B, unicode: 0x26B2},
        {f: "Wingdings 3", code: 0x05C, unicode: 0x26B3},
        {f: "Wingdings 3", code: 0x05D, unicode: 0x26B4},
        {f: "Wingdings 3", code: 0x05E, unicode: 0x26B5},
        {f: "Wingdings 3", code: 0x05F, unicode: 0x26B6},
        {f: "Wingdings 3", code: 0x060, unicode: 0x26B7},
        {f: "Wingdings 3", code: 0x061, unicode: 0x26B8},
        {f: "Wingdings 3", code: 0x062, unicode: 0x26B9},
        {f: "Wingdings 3", code: 0x063, unicode: 0x26BA},
        {f: "Wingdings 3", code: 0x064, unicode: 0x26BB},
        {f: "Wingdings 3", code: 0x065, unicode: 0x26BC},
        {f: "Wingdings 3", code: 0x066, unicode: 0x26BD},
        {f: "Wingdings 3", code: 0x067, unicode: 0x26BE},
        {f: "Wingdings 3", code: 0x068, unicode: 0x26BF},
        {f: "Wingdings 3", code: 0x069, unicode: 0x26C0},
        {f: "Wingdings 3", code: 0x06A, unicode: 0x26C1},
        {f: "Wingdings 3", code: 0x06B, unicode: 0x26C2},
        {f: "Wingdings 3", code: 0x06C, unicode: 0x26C3},
        {f: "Wingdings 3", code: 0x06D, unicode: 0x26C4},
        {f: "Wingdings 3", code: 0x06E, unicode: 0x26C5},
        {f: "Wingdings 3", code: 0x06F, unicode: 0x26C6},
        {f: "Wingdings 3", code: 0x070, unicode: 0x26C7},
        {f: "Wingdings 3", code: 0x071, unicode: 0x26C8},
        {f: "Wingdings 3", code: 0x072, unicode: 0x26C9},
        {f: "Wingdings 3", code: 0x073, unicode: 0x26CA},
        {f: "Wingdings 3", code: 0x074, unicode: 0x26CB},
        {f: "Wingdings 3", code: 0x075, unicode: 0x26CC},
        {f: "Wingdings 3", code: 0x076, unicode: 0x26CD},
        {f: "Wingdings 3", code: 0x077, unicode: 0x26CE},
        {f: "Wingdings 3", code: 0x078, unicode: 0x26CF},
        {f: "Wingdings 3", code: 0x079, unicode: 0x26D0},
        {f: "Wingdings 3", code: 0x07A, unicode: 0x26D1},
        {f: "Wingdings 3", code: 0x07B, unicode: 0x26D2},
        {f: "Wingdings 3", code: 0x07C, unicode: 0x26D3},
        {f: "Wingdings 3", code: 0x07D, unicode: 0x26D4},
        {f: "Wingdings 3", code: 0x07E, unicode: 0x26D5},
        {f: "Wingdings 3", code: 0x07F, unicode: 0x26D6},
        {f: "Wingdings 3", code: 0x080, unicode: 0x26D7},
        {f: "Wingdings 3", code: 0x081, unicode: 0x26D8},
        {f: "Wingdings 3", code: 0x082, unicode: 0x26D9},
        {f: "Wingdings 3", code: 0x083, unicode: 0x26DA},
        {f: "Wingdings 3", code: 0x084, unicode: 0x26DB},
        {f: "Wingdings 3", code: 0x085, unicode: 0x26DC},
        {f: "Wingdings 3", code: 0x086, unicode: 0x26DD},
        {f: "Wingdings 3", code: 0x087, unicode: 0x26DE},
        {f: "Wingdings 3", code: 0x088, unicode: 0x26DF},
        {f: "Wingdings 3", code: 0x089, unicode: 0x26E0},
        {f: "Wingdings 3", code: 0x08A, unicode: 0x26E1},
        {f: "Wingdings 3", code: 0x08B, unicode: 0x26E2},
        {f: "Wingdings 3", code: 0x08C, unicode: 0x26E3},
        {f: "Wingdings 3", code: 0x08D, unicode: 0x26E4},
        {f: "Wingdings 3", code: 0x08E, unicode: 0x26E5},
        {f: "Wingdings 3", code: 0x08F, unicode: 0x26E6},
        {f: "Wingdings 3", code: 0x090, unicode: 0x26E7},
        {f: "Wingdings 3", code: 0x091, unicode: 0x26E8},
        {f: "Wingdings 3", code: 0x092, unicode: 0x26E9},
        {f: "Wingdings 3", code: 0x093, unicode: 0x26EA},
        {f: "Wingdings 3", code: 0x094, unicode: 0x26EB},
        {f: "Wingdings 3", code: 0x095, unicode: 0x26EC},
        {f: "Wingdings 3", code: 0x096, unicode: 0x26ED},
        {f: "Wingdings 3", code: 0x097, unicode: 0x26EE},
        {f: "Wingdings 3", code: 0x098, unicode: 0x26EF},
        {f: "Wingdings 3", code: 0x099, unicode: 0x26F0},
        {f: "Wingdings 3", code: 0x09A, unicode: 0x26F1},
        {f: "Wingdings 3", code: 0x09B, unicode: 0x26F2},
        {f: "Wingdings 3", code: 0x09C, unicode: 0x26F3},
        {f: "Wingdings 3", code: 0x09D, unicode: 0x26F4},
        {f: "Wingdings 3", code: 0x09E, unicode: 0x26F5},
        {f: "Wingdings 3", code: 0x09F, unicode: 0x26F6},
        {f: "Wingdings 3", code: 0x0A0, unicode: 0x26F7},
        {f: "Wingdings 3", code: 0x0A1, unicode: 0x26F8},
        {f: "Wingdings 3", code: 0x0A2, unicode: 0x26F9},
        {f: "Wingdings 3", code: 0x0A3, unicode: 0x26FA},
        {f: "Wingdings 3", code: 0x0A4, unicode: 0x26FB},
        {f: "Wingdings 3", code: 0x0A5, unicode: 0x26FC},
        {f: "Wingdings 3", code: 0x0A6, unicode: 0x26FD},
        {f: "Wingdings 3", code: 0x0A7, unicode: 0x26FE},
        {f: "Wingdings 3", code: 0x0A8, unicode: 0x26FF}
    ];

    /**
     * 将 Dingbat 字符映射为 Unicode 字符
     * @param {string} typefaceNode - 字体类型（如 "Wingdings 2"）
     * @param {string} buChar - Dingbat 字符
     * @returns {string|null} Unicode 字符或 null
     */
    TextUtils.getDingbatToUnicode = function(typefaceNode, buChar) {
        if (dingbat_unicode) {
            var dingbat_code = buChar.codePointAt(0) & 0xFFF;
            var char_unicode = null;
            var len = dingbat_unicode.length;
            var i = 0;
            while (len--) {
                var item = dingbat_unicode[i];
                if (item.f == typefaceNode && item.code == dingbat_code) {
                    char_unicode = item.unicode;
                    break;
                }
                i++;
            }
            return char_unicode;
        }
    };

    /**
     * Get HTML bullet character
     * Convert special bullet characters to their HTML unicode equivalents
     * @param {String} typefaceNode - The font typeface (e.g., "Wingdings", "Wingdings 2")
     * @param {String} buChar - The bullet character
     * @returns {String} HTML entity for the bullet character
     */
    TextUtils.getHtmlBullet = function(typefaceNode, buChar) {
        //http://www.alanwood.net/demos/wingdings.html
        //not work for IE11
        switch (buChar) {
            case "§":
                return "&#9632;"; // U+25A0 | Black square
            case "q":
                return "&#10065;"; // U+2751 | Lower right shadowed white square
            case "v":
                return "&#10070;"; // U+2756 | Black diamond minus white X
            case "Ø":
                return "&#11162;"; // U+2B9A | Three-D top-lighted rightwards equilateral arrowhead
            case "ü":
                return "&#10004;"; // U+2714 | Heavy check mark
            default:
                if (typefaceNode == "Wingdings" || typefaceNode == "Wingdings 2" || typefaceNode == "Wingdings 3") {
                    var wingCharCode = this.getDingbatToUnicode(typefaceNode, buChar);
                    if (wingCharCode !== null) {
                        return "&#" + wingCharCode + ";";
                    }
                }
                return "&#" + (buChar.charCodeAt(0)) + ";";
        }
    };

    // Export to global scope
    window.TextUtils = TextUtils;

})();

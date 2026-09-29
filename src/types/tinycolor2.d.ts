/**
 * tinycolor2 未随包提供类型声明，也未安装 @types/tinycolor2。
 * 迁移期以 any 声明其默认导出，待后续按需补充精确类型。
 */
declare module 'tinycolor2' {
    const tinycolor: any;
    export default tinycolor;
}

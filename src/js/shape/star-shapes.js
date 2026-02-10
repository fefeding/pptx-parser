import { PPTXXmlUtils } from '../utils/xml.js';

const SLIDE_FACTOR = 0.0001;

/**
 * Render star shapes (star4, star5, star6, star7, star8, star10, star12, star16, star24, star32)
 * @param {string} shapType - Shape type
 * @param {number} w - Width
 * @param {number} h - Height
 * @param {boolean} imgFillFlg - Image fill flag
 * @param {boolean} grndFillFlg - Gradient fill flag
 * @param {string} fillColor - Fill color
 * @param {object} border - Border object with color, width, strokeDasharray
 * @param {string} shpId - Shape ID
 * @param {object} shapeArcAlt - Shape arc alt
 * @param {object} node - XML node for shape adjustments
 * @returns {string} SVG string
 */
export function renderStar(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt, node) {
    let result = '';
    const hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
    const fill = !imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")";

    switch (shapType) {
        case "star4": {
            const adj = getAdjValue(node, "adj", 19098);
            const cnstVal1 = 50000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, cnstVal1);
            const iwd2 = wd2 * a / cnstVal1;
            const ihd2 = hd2 * a / cnstVal1;
            const sdx = iwd2 * Math.cos(0.7853981634);
            const sdy = ihd2 * Math.sin(0.7853981634);
            const sx1 = hc - sdx;
            const sx2 = hc + sdx;
            const sy1 = vc - sdy;
            const sy2 = vc + sdy;

            const d = "M0" + "," + vc +
                " L" + sx1 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx2 + "," + sy1 +
                " L" + w + "," + vc +
                " L" + sx2 + "," + sy2 +
                " L" + hc + "," + h +
                " L" + sx1 + "," + sy2 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star5": {
            const adj = getAdjValue(node, "adj", 19098);
            const hf = getAdjValue(node, "hf", 105146);
            const vf = getAdjValue(node, "vf", 110557);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const cnstVal1 = 100000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const swd2 = wd2 * hf / cnstVal1;
            const shd2 = hd2 * vf / cnstVal1;
            const svc = vc * vf / cnstVal1;
            const dx1 = swd2 * Math.cos(0.31415926536);
            const dx2 = swd2 * Math.cos(5.3407075111);
            const dy1 = shd2 * Math.sin(0.31415926536);
            const dy2 = shd2 * Math.sin(5.3407075111);
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc + dx2;
            const x4 = hc + dx1;
            const y1 = svc - dy1;
            const y2 = svc - dy2;
            const iwd2 = swd2 * a / maxAdj;
            const ihd2 = shd2 * a / maxAdj;
            const sdx1 = iwd2 * Math.cos(5.9690260418);
            const sdx2 = iwd2 * Math.cos(0.94247779608);
            const sdy1 = ihd2 * Math.sin(0.94247779608);
            const sdy2 = ihd2 * Math.sin(5.9690260418);
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc + sdx2;
            const sx4 = hc + sdx1;
            const sy1 = svc - sdy1;
            const sy2 = svc - sdy2;
            const sy3 = svc + ihd2;

            const d = "M" + x1 + "," + y1 +
                " L" + sx2 + "," + sy1 +
                " L" + hc + "," + 0 +
                " L" + sx3 + "," + sy1 +
                " L" + x4 + "," + y1 +
                " L" + sx4 + "," + sy2 +
                " L" + x3 + "," + y2 +
                " L" + hc + "," + sy3 +
                " L" + x2 + "," + y2 +
                " L" + sx1 + "," + sy2 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star6": {
            const adj = getAdjValue(node, "adj", 28868);
            const hf = getAdjValue(node, "hf", 115470);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const cnstVal1 = 100000 * SLIDE_FACTOR;
            const hd4 = h / 4;
            const a = clamp(adj, 0, maxAdj);
            const swd2 = wd2 * hf / cnstVal1;
            const dx1 = swd2 * Math.cos(0.5235987756);
            const x1 = hc - dx1;
            const x2 = hc + dx1;
            const y2 = vc + hd4;
            const iwd2 = swd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx2 = iwd2 / 2;
            const sx1 = hc - iwd2;
            const sx2 = hc - sdx2;
            const sx3 = hc + sdx2;
            const sx4 = hc + iwd2;
            const sdy1 = ihd2 * Math.sin(1.0471975512);
            const sy1 = vc - sdy1;
            const sy2 = vc + sdy1;

            const d = "M" + x1 + "," + hd4 +
                " L" + sx2 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx3 + "," + sy1 +
                " L" + x2 + "," + hd4 +
                " L" + sx4 + "," + vc +
                " L" + x2 + "," + y2 +
                " L" + sx3 + "," + sy2 +
                " L" + hc + "," + h +
                " L" + sx2 + "," + sy2 +
                " L" + x1 + "," + y2 +
                " L" + sx1 + "," + vc +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star7": {
            const adj = getAdjValue(node, "adj", 34601);
            const hf = getAdjValue(node, "hf", 102572);
            const vf = getAdjValue(node, "vf", 105210);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const cnstVal1 = 100000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const swd2 = wd2 * hf / cnstVal1;
            const shd2 = hd2 * vf / cnstVal1;
            const svc = vc * vf / cnstVal1;
            const dx1 = swd2 * 97493 / 100000;
            const dx2 = swd2 * 78183 / 100000;
            const dx3 = swd2 * 43388 / 100000;
            const dy1 = shd2 * 62349 / 100000;
            const dy2 = shd2 * 22252 / 100000;
            const dy3 = shd2 * 90097 / 100000;
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc - dx3;
            const x4 = hc + dx3;
            const x5 = hc + dx2;
            const x6 = hc + dx1;
            const y1 = svc - dy1;
            const y2 = svc + dy2;
            const y3 = svc + dy3;
            const iwd2 = swd2 * a / maxAdj;
            const ihd2 = shd2 * a / maxAdj;
            const sdx1 = iwd2 * 97493 / 100000;
            const sdx2 = iwd2 * 78183 / 100000;
            const sdx3 = iwd2 * 43388 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc + sdx3;
            const sx5 = hc + sdx2;
            const sx6 = hc + sdx1;
            const sdy1 = ihd2 * 90097 / 100000;
            const sdy2 = ihd2 * 22252 / 100000;
            const sdy3 = ihd2 * 62349 / 100000;
            const sy1 = svc - sdy1;
            const sy2 = svc - sdy2;
            const sy3 = svc + sdy3;
            const sy4 = svc + ihd2;

            const d = "M" + x1 + "," + y2 +
                " L" + sx1 + "," + sy2 +
                " L" + x2 + "," + y1 +
                " L" + sx3 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx4 + "," + sy1 +
                " L" + x5 + "," + y1 +
                " L" + sx6 + "," + sy2 +
                " L" + x6 + "," + y2 +
                " L" + sx5 + "," + sy3 +
                " L" + x4 + "," + y3 +
                " L" + hc + "," + sy4 +
                " L" + x3 + "," + y3 +
                " L" + sx2 + "," + sy3 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star8": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const cnstVal1 = 100000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * Math.cos(0.7853981634);
            const x1 = hc - dx1;
            const x2 = hc + dx1;
            const dy1 = hd2 * Math.sin(0.7853981634);
            const y1 = vc - dy1;
            const y2 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 92388 / 100000;
            const sdx2 = iwd2 * 38268 / 100000;
            const sdy1 = ihd2 * 92388 / 100000;
            const sdy2 = ihd2 * 38268 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc + sdx2;
            const sx4 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc + sdy2;
            const sy4 = vc + sdy1;

            const d = "M0" + "," + vc +
                " L" + sx1 + "," + sy2 +
                " L" + x1 + "," + y1 +
                " L" + sx2 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx3 + "," + sy1 +
                " L" + x2 + "," + y1 +
                " L" + sx4 + "," + sy2 +
                " L" + w + "," + vc +
                " L" + sx4 + "," + sy3 +
                " L" + x2 + "," + y2 +
                " L" + sx3 + "," + sy4 +
                " L" + hc + "," + h +
                " L" + sx2 + "," + sy4 +
                " L" + x1 + "," + y2 +
                " L" + sx1 + "," + sy3 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star10": {
            const adj = getAdjValue(node, "adj", 42533);
            const hf = getAdjValue(node, "hf", 105146);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const cnstVal1 = 100000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const swd2 = wd2 * hf / cnstVal1;
            const dx1 = swd2 * 95106 / 100000;
            const dx2 = swd2 * 58779 / 100000;
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc + dx2;
            const x4 = hc + dx1;
            const dy1 = hd2 * 80902 / 100000;
            const dy2 = hd2 * 30902 / 100000;
            const y1 = vc - dy1;
            const y2 = vc - dy2;
            const y3 = vc + dy2;
            const y4 = vc + dy1;
            const iwd2 = swd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 80902 / 100000;
            const sdx2 = iwd2 * 30902 / 100000;
            const sdy1 = ihd2 * 95106 / 100000;
            const sdy2 = ihd2 * 58779 / 100000;
            const sx1 = hc - iwd2;
            const sx2 = hc - sdx1;
            const sx3 = hc - sdx2;
            const sx4 = hc + sdx2;
            const sx5 = hc + sdx1;
            const sx6 = hc + iwd2;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc + sdy2;
            const sy4 = vc + sdy1;

            const d = "M" + x1 + "," + y2 +
                " L" + sx2 + "," + sy2 +
                " L" + x2 + "," + y1 +
                " L" + sx3 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx4 + "," + sy1 +
                " L" + x3 + "," + y1 +
                " L" + sx5 + "," + sy2 +
                " L" + x4 + "," + y2 +
                " L" + sx6 + "," + vc +
                " L" + x4 + "," + y3 +
                " L" + sx5 + "," + sy3 +
                " L" + x3 + "," + y4 +
                " L" + sx4 + "," + sy4 +
                " L" + hc + "," + h +
                " L" + sx3 + "," + sy4 +
                " L" + x2 + "," + y4 +
                " L" + sx2 + "," + sy3 +
                " L" + x1 + "," + y3 +
                " L" + sx1 + "," + vc +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star12": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const hd4 = h / 4;
            const wd4 = w / 4;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * Math.cos(0.5235987756);
            const dy1 = hd2 * Math.sin(1.0471975512);
            const x1 = hc - dx1;
            const x3 = w * 3 / 4;
            const x4 = hc + dx1;
            const y1 = vc - dy1;
            const y3 = h * 3 / 4;
            const y4 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * Math.cos(0.2617993878);
            const sdx2 = iwd2 * Math.cos(0.7853981634);
            const sdx3 = iwd2 * Math.cos(1.308996939);
            const sdy1 = ihd2 * Math.sin(1.308996939);
            const sdy2 = ihd2 * Math.sin(0.7853981634);
            const sdy3 = ihd2 * Math.sin(0.2617993878);
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc + sdx3;
            const sx5 = hc + sdx2;
            const sx6 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc - sdy3;
            const sy4 = vc + sdy3;
            const sy5 = vc + sdy2;
            const sy6 = vc + sdy1;

            const d = "M0" + "," + vc +
                " L" + sx1 + "," + sy3 +
                " L" + x1 + "," + hd4 +
                " L" + sx2 + "," + sy2 +
                " L" + wd4 + "," + y1 +
                " L" + sx3 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx4 + "," + sy1 +
                " L" + x3 + "," + y1 +
                " L" + sx5 + "," + sy2 +
                " L" + x4 + "," + hd4 +
                " L" + sx6 + "," + sy3 +
                " L" + w + "," + vc +
                " L" + sx6 + "," + sy4 +
                " L" + x4 + "," + y3 +
                " L" + sx5 + "," + sy5 +
                " L" + x3 + "," + y4 +
                " L" + sx4 + "," + sy6 +
                " L" + hc + "," + h +
                " L" + sx3 + "," + sy6 +
                " L" + wd4 + "," + y4 +
                " L" + sx2 + "," + sy5 +
                " L" + x1 + "," + y3 +
                " L" + sx1 + "," + sy4 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star16": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * 92388 / 100000;
            const dx2 = wd2 * 70711 / 100000;
            const dx3 = wd2 * 38268 / 100000;
            const dy1 = hd2 * 92388 / 100000;
            const dy2 = hd2 * 70711 / 100000;
            const dy3 = hd2 * 38268 / 100000;
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc - dx3;
            const x4 = hc + dx3;
            const x5 = hc + dx2;
            const x6 = hc + dx1;
            const y1 = vc - dy1;
            const y2 = vc - dy2;
            const y3 = vc - dy3;
            const y4 = vc + dy3;
            const y5 = vc + dy2;
            const y6 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 98079 / 100000;
            const sdx2 = iwd2 * 83147 / 100000;
            const sdx3 = iwd2 * 55557 / 100000;
            const sdx4 = iwd2 * 19509 / 100000;
            const sdy1 = ihd2 * 98079 / 100000;
            const sdy2 = ihd2 * 83147 / 100000;
            const sdy3 = ihd2 * 55557 / 100000;
            const sdy4 = ihd2 * 19509 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc - sdx4;
            const sx5 = hc + sdx4;
            const sx6 = hc + sdx3;
            const sx7 = hc + sdx2;
            const sx8 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc - sdy3;
            const sy4 = vc - sdy4;
            const sy5 = vc + sdy4;
            const sy6 = vc + sdy3;
            const sy7 = vc + sdy2;
            const sy8 = vc + sdy1;

            const d = "M0" + "," + vc +
                " L" + sx1 + "," + sy4 +
                " L" + x1 + "," + y3 +
                " L" + sx2 + "," + sy3 +
                " L" + x2 + "," + y2 +
                " L" + sx3 + "," + sy2 +
                " L" + x3 + "," + y1 +
                " L" + sx4 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx5 + "," + sy1 +
                " L" + x4 + "," + y1 +
                " L" + sx6 + "," + sy2 +
                " L" + x5 + "," + y2 +
                " L" + sx7 + "," + sy3 +
                " L" + x6 + "," + y3 +
                " L" + sx8 + "," + sy4 +
                " L" + w + "," + vc +
                " L" + sx8 + "," + sy5 +
                " L" + x6 + "," + y4 +
                " L" + sx7 + "," + sy6 +
                " L" + x5 + "," + y5 +
                " L" + sx6 + "," + sy7 +
                " L" + x4 + "," + y6 +
                " L" + sx5 + "," + sy8 +
                " L" + hc + "," + h +
                " L" + sx4 + "," + sy8 +
                " L" + x3 + "," + y6 +
                " L" + sx3 + "," + sy7 +
                " L" + x2 + "," + y5 +
                " L" + sx2 + "," + sy6 +
                " L" + x1 + "," + y4 +
                " L" + sx1 + "," + sy5 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star24": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const hd4 = h / 4;
            const wd4 = w / 4;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * Math.cos(0.2617993878);
            const dx2 = wd2 * Math.cos(0.5235987756);
            const dx3 = wd2 * Math.cos(0.7853981634);
            const dx4 = wd4;
            const dx5 = wd2 * Math.cos(1.308996939);
            const dy1 = hd2 * Math.sin(1.308996939);
            const dy2 = hd2 * Math.sin(1.0471975512);
            const dy3 = hd2 * Math.sin(0.7853981634);
            const dy4 = hd4;
            const dy5 = hd2 * Math.sin(0.2617993878);
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc - dx3;
            const x4 = hc - dx4;
            const x5 = hc - dx5;
            const x6 = hc + dx5;
            const x7 = hc + dx4;
            const x8 = hc + dx3;
            const x9 = hc + dx2;
            const x10 = hc + dx1;
            const y1 = vc - dy1;
            const y2 = vc - dy2;
            const y3 = vc - dy3;
            const y4 = vc - dy4;
            const y5 = vc - dy5;
            const y6 = vc + dy5;
            const y7 = vc + dy4;
            const y8 = vc + dy3;
            const y9 = vc + dy2;
            const y10 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 99144 / 100000;
            const sdx2 = iwd2 * 92388 / 100000;
            const sdx3 = iwd2 * 79335 / 100000;
            const sdx4 = iwd2 * 60876 / 100000;
            const sdx5 = iwd2 * 38268 / 100000;
            const sdx6 = iwd2 * 13053 / 100000;
            const sdy1 = ihd2 * 99144 / 100000;
            const sdy2 = ihd2 * 92388 / 100000;
            const sdy3 = ihd2 * 79335 / 100000;
            const sdy4 = ihd2 * 60876 / 100000;
            const sdy5 = ihd2 * 38268 / 100000;
            const sdy6 = ihd2 * 13053 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc - sdx4;
            const sx5 = hc - sdx5;
            const sx6 = hc - sdx6;
            const sx7 = hc + sdx6;
            const sx8 = hc + sdx5;
            const sx9 = hc + sdx4;
            const sx10 = hc + sdx3;
            const sx11 = hc + sdx2;
            const sx12 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc - sdy3;
            const sy4 = vc - sdy4;
            const sy5 = vc - sdy5;
            const sy6 = vc - sdy6;
            const sy7 = vc + sdy6;
            const sy8 = vc + sdy5;
            const sy9 = vc + sdy4;
            const sy10 = vc + sdy3;
            const sy11 = vc + sdy2;
            const sy12 = vc + sdy1;

            const d = "M0" + "," + vc +
                " L" + sx1 + "," + sy6 +
                " L" + x1 + "," + y5 +
                " L" + sx2 + "," + sy5 +
                " L" + x2 + "," + y4 +
                " L" + sx3 + "," + sy4 +
                " L" + x3 + "," + y3 +
                " L" + sx4 + "," + sy3 +
                " L" + x4 + "," + y2 +
                " L" + sx5 + "," + sy2 +
                " L" + x5 + "," + y1 +
                " L" + sx6 + "," + sy1 +
                " L" + hc + "," + 0 +
                " L" + sx7 + "," + sy1 +
                " L" + x6 + "," + y1 +
                " L" + sx8 + "," + sy2 +
                " L" + x7 + "," + y2 +
                " L" + sx9 + "," + sy3 +
                " L" + x8 + "," + y3 +
                " L" + sx10 + "," + sy4 +
                " L" + x9 + "," + y4 +
                " L" + sx11 + "," + sy5 +
                " L" + x10 + "," + y5 +
                " L" + sx12 + "," + sy6 +
                " L" + w + "," + vc +
                " L" + sx12 + "," + sy7 +
                " L" + x10 + "," + y6 +
                " L" + sx11 + "," + sy8 +
                " L" + x9 + "," + y7 +
                " L" + sx10 + "," + sy9 +
                " L" + x8 + "," + y8 +
                " L" + sx9 + "," + sy10 +
                " L" + x7 + "," + y9 +
                " L" + sx8 + "," + sy11 +
                " L" + x6 + "," + y10 +
                " L" + sx7 + "," + sy12 +
                " L" + hc + "," + h +
                " L" + sx6 + "," + sy12 +
                " L" + x5 + "," + y10 +
                " L" + sx5 + "," + sy11 +
                " L" + x4 + "," + y9 +
                " L" + sx4 + "," + sy10 +
                " L" + x3 + "," + y8 +
                " L" + sx3 + "," + sy9 +
                " L" + x2 + "," + y7 +
                " L" + sx2 + "," + sy8 +
                " L" + x1 + "," + y6 +
                " L" + sx1 + "," + sy7 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star32": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const hd4 = h / 4;
            const wd4 = w / 4;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * 98079 / 100000;
            const dx2 = wd2 * 92388 / 100000;
            const dx3 = wd2 * 83147 / 100000;
            const dx4 = wd2 * Math.cos(0.7853981634);
            const dx5 = wd2 * 55557 / 100000;
            const dx6 = wd2 * 38268 / 100000;
            const dx7 = wd2 * 19509 / 100000;
            const dy1 = hd2 * 98079 / 100000;
            const dy2 = hd2 * 92388 / 100000;
            const dy3 = hd2 * 83147 / 100000;
            const dy4 = hd2 * Math.sin(0.7853981634);
            const dy5 = hd2 * 55557 / 100000;
            const dy6 = hd2 * 38268 / 100000;
            const dy7 = hd2 * 19509 / 100000;
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc - dx3;
            const x4 = hc - dx4;
            const x5 = hc - dx5;
            const x6 = hc - dx6;
            const x7 = hc - dx7;
            const x8 = hc + dx7;
            const x9 = hc + dx6;
            const x10 = hc + dx5;
            const x11 = hc + dx4;
            const x12 = hc + dx3;
            const x13 = hc + dx2;
            const x14 = hc + dx1;
            const y1 = vc - dy1;
            const y2 = vc - dy2;
            const y3 = vc - dy3;
            const y4 = vc - dy4;
            const y5 = vc - dy5;
            const y6 = vc - dy6;
            const y7 = vc - dy7;
            const y8 = vc + dy7;
            const y9 = vc + dy6;
            const y10 = vc + dy5;
            const y11 = vc + dy4;
            const y12 = vc + dy3;
            const y13 = vc + dy2;
            const y14 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 99518 / 100000;
            const sdx2 = iwd2 * 95694 / 100000;
            const sdx3 = iwd2 * 88192 / 100000;
            const sdx4 = iwd2 * 77301 / 100000;
            const sdx5 = iwd2 * 63439 / 100000;
            const sdx6 = iwd2 * 47140 / 100000;
            const sdx7 = iwd2 * 29028 / 100000;
            const sdx8 = iwd2 * 9802 / 100000;
            const sdy1 = ihd2 * 99518 / 100000;
            const sdy2 = ihd2 * 95694 / 100000;
            const sdy3 = ihd2 * 88192 / 100000;
            const sdy4 = ihd2 * 77301 / 100000;
            const sdy5 = ihd2 * 63439 / 100000;
            const sdy6 = ihd2 * 47140 / 100000;
            const sdy7 = ihd2 * 29028 / 100000;
            const sdy8 = ihd2 * 9802 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc - sdx4;
            const sx5 = hc - sdx5;
            const sx6 = hc - sdx6;
            const sx7 = hc - sdx7;
            const sx8 = hc - sdx8;
            const sx9 = hc + sdx8;
            const sx10 = hc + sdx7;
            const sx11 = hc + sdx6;
            const sx12 = hc + sdx5;
            const sx13 = hc + sdx4;
            const sx14 = hc + sdx3;
            const sx15 = hc + sdx2;
            const sx16 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc - sdy3;
            const sy4 = vc - sdy4;
            const sy5 = vc - sdy5;
            const sy6 = vc - sdy6;
            const sy7 = vc - sdy7;
            const sy8 = vc - sdy8;
            const sy9 = vc + sdy8;
            const sy10 = vc + sdy7;
            const sy11 = vc + sdy6;
            const sy12 = vc + sdy5;
            const sy13 = vc + sdy4;
            const sy14 = vc + sdy3;
            const sy15 = vc + sdy2;
            const sy16 = vc + sdy1;

            const d = "M0" + "," + vc +
                " L" + sx1 + "," + sy8 +
                " L" + x1 + "," + y7 +
                " L" + sx2 + "," + sy7 +
                " L" + x2 + "," + y6 +
                " L" + sx3 + "," + sy6 +
                " L" + x3 + "," + y5 +
                " L" + sx4 + "," + sy5 +
                " L" + x4 + "," + y4 +
                " L" + sx5 + "," + sy4 +
                " L" + x5 + "," + y3 +
                " L" + sx6 + "," + sy3 +
                " L" + x6 + "," + y2 +
                " L" + sx7 + "," + sy2 +
                " L" + x7 + "," + y1 +
                " L" + sx8 + "," + sy1 +
                " L" + hc + "," + 0 +
                " L" + sx9 + "," + sy1 +
                " L" + x8 + "," + y1 +
                " L" + sx10 + "," + sy2 +
                " L" + x9 + "," + y2 +
                " L" + sx11 + "," + sy3 +
                " L" + x10 + "," + y3 +
                " L" + sx12 + "," + sy4 +
                " L" + x11 + "," + y4 +
                " L" + sx13 + "," + sy5 +
                " L" + x12 + "," + y5 +
                " L" + sx14 + "," + sy6 +
                " L" + x13 + "," + y6 +
                " L" + sx15 + "," + sy7 +
                " L" + x14 + "," + y7 +
                " L" + sx16 + "," + sy8 +
                " L" + w + "," + vc +
                " L" + sx16 + "," + sy9 +
                " L" + x14 + "," + y8 +
                " L" + sx15 + "," + sy10 +
                " L" + x13 + "," + y9 +
                " L" + sx14 + "," + sy11 +
                " L" + x12 + "," + y10 +
                " L" + sx13 + "," + sy12 +
                " L" + x11 + "," + y11 +
                " L" + sx12 + "," + sy13 +
                " L" + x10 + "," + y12 +
                " L" + sx11 + "," + sy14 +
                " L" + x9 + "," + y13 +
                " L" + sx10 + "," + sy15 +
                " L" + x8 + "," + y14 +
                " L" + sx9 + "," + sy16 +
                " L" + hc + "," + h +
                " L" + sx8 + "," + sy16 +
                " L" + x7 + "," + y14 +
                " L" + sx7 + "," + sy15 +
                " L" + x6 + "," + y13 +
                " L" + sx6 + "," + sy14 +
                " L" + x5 + "," + y12 +
                " L" + sx5 + "," + sy13 +
                " L" + x4 + "," + y11 +
                " L" + sx4 + "," + sy12 +
                " L" + x3 + "," + y10 +
                " L" + sx3 + "," + sy11 +
                " L" + x2 + "," + y9 +
                " L" + sx2 + "," + sy10 +
                " L" + x1 + "," + y8 +
                " L" + sx1 + "," + sy9 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
    }

    return result;
}

/**
 * Check if shape type is a star shape
 * @param {string} shapType - Shape type
 * @returns {boolean}
 */
export function isStar(shapType) {
    return ["star4", "star5", "star6", "star7", "star8", "star10", "star12", "star16", "star24", "star32"].includes(shapType);
}

/**
 * Get adjustment value from node
 * @param {object} node - XML node
 * @param {string} name - Adjustment name
 * @param {number} defaultValue - Default value
 * @returns {number}
 */
function getAdjValue(node, name, defaultValue) {
    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    if (shapAdjst !== undefined) {
        if (Array.isArray(shapAdjst)) {
            for (let key of Object.keys(shapAdjst)) {
                if (shapAdjst[key] && shapAdjst[key]["attrs"] && shapAdjst[key]["attrs"]["name"] === name) {
                    return parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * SLIDE_FACTOR;
                }
            }
        } else if (shapAdjst["attrs"] && shapAdjst["attrs"]["name"] === name) {
            return parseInt(shapAdjst["attrs"]["fmla"].substr(4)) * SLIDE_FACTOR;
        }
    }
    return defaultValue * SLIDE_FACTOR;
}

/**
 * Clamp value between min and max
 * @param {number} value - Value to clamp
 * @param {number} min - Minimum value
 * @param {number} max - Maximum value
 * @returns {number}
 */
function clamp(value, min, max) {
    return value < min ? min : value > max ? max : value;
}

/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";
(function (window, undefined) {
	window['AscCommon'] = window['AscCommon'] || {};

	const cd16 = 1.0 / 6.0;
	const cd13 = 1.0 / 3.0;
	const cd23 = 2.0 / 3.0;
	const max_hls = 255.0;

	// ---------------- sRGB gamma ----------------

	function srgbToLinear(c) {
		if (c <= 0.04045) {
			return c / 12.92;
		}
		return Math.pow((c + 0.055) / 1.055, 2.4);
	}

	function linearToSrgb(v) {
		if (v <= 0.0031308) {
			return 12.92 * v;
		}
		return 1.055 * Math.pow(v, 1 / 2.4) - 0.055;
	}

	function srgbByteToLinear(b) {
		return srgbToLinear(b / 255);
	}

	function linearToSrgbByte(v) {
		return Math.round(linearToSrgb(v) * 255);
	}

	// ---------------- WCAG ----------------

	function relativeLuminance(rgb) {
		return 0.2126 * srgbByteToLinear(rgb.R)
			+ 0.7152 * srgbByteToLinear(rgb.G)
			+ 0.0722 * srgbByteToLinear(rgb.B);
	}

	function wcagContrastRatio(lum1, lum2) {
		const lighter = Math.max(lum1, lum2);
		const darker = Math.min(lum1, lum2);
		return (lighter + 0.05) / (darker + 0.05);
	}

	function wcagContrastRatioRgb(rgb1, rgb2) {
		return wcagContrastRatio(relativeLuminance(rgb1), relativeLuminance(rgb2));
	}

	// ---------------- RGB <-> HSL (Office scale: H, S, L all 0..255) ----------------

	function rgbToHsl(rgb) {
		const R = rgb.R, G = rgb.G, B = rgb.B;
		let iMin = (R < G ? R : G);
		iMin = iMin < B ? iMin : B;
		let iMax = (R > G ? R : G);
		iMax = iMax > B ? iMax : B;
		const iDelta = iMax - iMin;
		const dMax = (iMax + iMin) / 255.0;
		let dDelta = iDelta / 255.0;
		let H = 0;
		let S = 0;
		let L = dMax / 2.0;

		if (iDelta != 0) {
			if (L < 0.5) S = dDelta / dMax;
			else S = dDelta / (2.0 - dMax);

			dDelta = dDelta * 1530.0;
			const dR = (iMax - R) / dDelta;
			const dG = (iMax - G) / dDelta;
			const dB = (iMax - B) / dDelta;

			if (R == iMax) H = dB - dG;
			else if (G == iMax) H = cd13 + dR - dB;
			else if (B == iMax) H = cd23 + dG - dR;

			if (H < 0.0) H += 1.0;
			if (H > 1.0) H -= 1.0;
		}

		H = H * max_hls;
		if (H < 0) H = 0;
		if (H > 255) H = 255;

		S = S * max_hls;
		if (S < 0) S = 0;
		if (S > 255) S = 255;

		L = L * max_hls;
		if (L < 0) L = 0;
		if (L > 255) L = 255;

		return { H: H, S: S, L: L };
	}

	function hue2Rgb(v1, v2, vH) {
		if (vH < 0.0) vH += 1.0;
		if (vH > 1.0) vH -= 1.0;
		if (vH < cd16) return v1 + (v2 - v1) * 6.0 * vH;
		if (vH < 0.5) return v2;
		if (vH < cd23) return v1 + (v2 - v1) * (cd23 - vH) * 6.0;
		return v1;
	}

	// Match AscFormat.ClampColor (defined in ColorArray.js): round + clamp to 0..255.
	// Inlined here so colorUtils.js has no dependency on AscFormat at load time.
	function clampColor(c) {
		const t = (c + 0.5) >> 0;
		return (t < 0) ? 0 : (t > 255 ? 255 : t);
	}

	function hslToRgb(hsl, bRoundValues) {
		if (hsl.S == 0) {
			const clampL = bRoundValues ? clampColor(hsl.L) : hsl.L;
			return { R: clampL, G: clampL, B: clampL };
		}
		const H = hsl.H / max_hls;
		const S = hsl.S / max_hls;
		const L = hsl.L / max_hls;
		const v2 = L < 0.5 ? L * (1.0 + S) : L + S - S * L;
		const v1 = 2.0 * L - v2;

		const R = (255 * hue2Rgb(v1, v2, H + cd13));
		const G = (255 * hue2Rgb(v1, v2, H));
		const B = (255 * hue2Rgb(v1, v2, H - cd13));

		if (bRoundValues) {
			return { R: clampColor(R), G: clampColor(G), B: clampColor(B) };
		}
		return { R: R, G: G, B: B };
	}

	// ---------------- RGB -> LAB ----------------
	// Returns {L, a, b} as floats: L in 0..100, a/b roughly -128..127.
	// asc_CColor.RGB2LAB wrapper rounds and rescales to byte form.

	function rgbToLab(rgb) {
		const eps = 216.0 / 24389.0;
		const k = 24389.0 / 27.0;

		const Xr = 0.95047; // reference white D65
		const Yr = 1.0;
		const Zr = 1.08883;

		let r = rgb.R / 255;
		let g = rgb.G / 255;
		let b = rgb.B / 255;

		if (r <= 0.04045) r = r / 12.92; else r = Math.pow((r + 0.055) / 1.055, 2.4);
		if (g <= 0.04045) g = g / 12.92; else g = Math.pow((g + 0.055) / 1.055, 2.4);
		if (b <= 0.04045) b = b / 12.92; else b = Math.pow((b + 0.055) / 1.055, 2.4);

		const X = 0.4124564 * r + 0.3575761 * g + 0.1804375 * b;
		const Y = 0.2126729 * r + 0.7151522 * g + 0.0721750 * b;
		const Z = 0.0193339 * r + 0.1191920 * g + 0.9503041 * b;

		const xr = X / Xr;
		const yr = Y / Yr;
		const zr = Z / Zr;

		const fx = xr > eps ? Math.pow(xr, 1 / 3.) : ((k * xr + 16.) / 116.);
		const fy = yr > eps ? Math.pow(yr, 1 / 3.) : ((k * yr + 16.) / 116.);
		const fz = zr > eps ? Math.pow(zr, 1 / 3.) : ((k * zr + 16.) / 116);

		const Ls = (116 * fy) - 16;
		const as = 500 * (fx - fy);
		const bs = 200 * (fy - fz);

		return { L: Ls, a: as, b: bs };
	}

	// Delta E 2000 (CIEDE2000). Caller controls the L scale: passing canonical L
	// (0..100) gives canonical Delta E; passing byte-scale L (0..255) gives the
	// scaled distance used by asc_CColor.getColorDiff.
	function deltaE2000(lab1, lab2) {
		const d2r = window['AscCommon'].deg2rad;

		const L1 = lab1.L, a1 = lab1.a, b1 = lab1.b;
		const L2 = lab2.L, a2 = lab2.a, b2 = lab2.b;

		const k_L = 1.0, k_C = 1.0, k_H = 1.0;
		const deg360InRad = d2r(360.0);
		const deg180InRad = d2r(180.0);
		const pow25To7 = 6103515625.0;

		function fAE(a, b) { return Math.abs(a - b) < 1e-15; }

		const C1 = Math.sqrt((a1 * a1) + (b1 * b1));
		const C2 = Math.sqrt((a2 * a2) + (b2 * b2));
		const barC = (C1 + C2) / 2.0;
		const G = 0.5 * (1 - Math.sqrt(Math.pow(barC, 7) / (Math.pow(barC, 7) + pow25To7)));
		const a1Prime = (1.0 + G) * a1;
		const a2Prime = (1.0 + G) * a2;
		const CPrime1 = Math.sqrt((a1Prime * a1Prime) + (b1 * b1));
		const CPrime2 = Math.sqrt((a2Prime * a2Prime) + (b2 * b2));

		let hPrime1;
		if (fAE(b1, 0.0) && fAE(a1Prime, 0.0)) hPrime1 = 0.0;
		else { hPrime1 = Math.atan2(b1, a1Prime); if (hPrime1 < 0) hPrime1 += deg360InRad; }
		let hPrime2;
		if (fAE(b2, 0.0) && fAE(a2Prime, 0.0)) hPrime2 = 0.0;
		else { hPrime2 = Math.atan2(b2, a2Prime); if (hPrime2 < 0) hPrime2 += deg360InRad; }

		const deltaLPrime = L2 - L1;
		const deltaCPrime = CPrime2 - CPrime1;
		let deltahPrime;
		const CPrimeProduct = CPrime1 * CPrime2;
		if (fAE(CPrimeProduct, 0.0)) deltahPrime = 0;
		else {
			deltahPrime = hPrime2 - hPrime1;
			if (deltahPrime < -deg180InRad) deltahPrime += deg360InRad;
			else if (deltahPrime > deg180InRad) deltahPrime -= deg360InRad;
		}
		const deltaHPrime = 2.0 * Math.sqrt(CPrimeProduct) * Math.sin(deltahPrime / 2.0);

		const barLPrime = (L1 + L2) / 2.0;
		const barCPrime = (CPrime1 + CPrime2) / 2.0;
		const hPrimeSum = hPrime1 + hPrime2;
		let barhPrime;
		if (fAE(CPrime1 * CPrime2, 0.0)) {
			barhPrime = hPrimeSum;
		} else {
			if (Math.abs(hPrime1 - hPrime2) <= deg180InRad) barhPrime = hPrimeSum / 2.0;
			else {
				if (hPrimeSum < deg360InRad) barhPrime = (hPrimeSum + deg360InRad) / 2.0;
				else barhPrime = (hPrimeSum - deg360InRad) / 2.0;
			}
		}
		const T = 1.0 - (0.17 * Math.cos(barhPrime - d2r(30.0))) + (0.24 * Math.cos(2.0 * barhPrime)) + (0.32 * Math.cos((3.0 * barhPrime) + d2r(6.0))) - (0.20 * Math.cos((4.0 * barhPrime) - d2r(63.0)));
		const deltaTheta = d2r(30.0) * Math.exp(-Math.pow((barhPrime - d2r(275.0)) / d2r(25.0), 2.0));
		const R_C = 2.0 * Math.sqrt(Math.pow(barCPrime, 7.0) / (Math.pow(barCPrime, 7.0) + pow25To7));
		const S_L = 1 + ((0.015 * Math.pow(barLPrime - 50.0, 2.0)) / Math.sqrt(20 + Math.pow(barLPrime - 50.0, 2.0)));
		const S_C = 1 + (0.045 * barCPrime);
		const S_H = 1 + (0.015 * barCPrime * T);
		const R_T = (-Math.sin(2.0 * deltaTheta)) * R_C;

		return Math.sqrt(
			Math.pow(deltaLPrime / (k_L * S_L), 2.0)
			+ Math.pow(deltaCPrime / (k_C * S_C), 2.0)
			+ Math.pow(deltaHPrime / (k_H * S_H), 2.0)
			+ (R_T * (deltaCPrime / (k_C * S_C)) * (deltaHPrime / (k_H * S_H)))
		);
	}

	// ---------------- Hex ----------------

	function byteToHex(b) {
		return window['AscCommon'].ByteToHex(b);
	}

	// ---------------- Exports ----------------

	const CU = {
		srgbToLinear: srgbToLinear,
		linearToSrgb: linearToSrgb,
		srgbByteToLinear: srgbByteToLinear,
		linearToSrgbByte: linearToSrgbByte,
		relativeLuminance: relativeLuminance,
		wcagContrastRatio: wcagContrastRatio,
		wcagContrastRatioRgb: wcagContrastRatioRgb,
		rgbToHsl: rgbToHsl,
		hslToRgb: hslToRgb,
		rgbToLab: rgbToLab,
		deltaE2000: deltaE2000,
		byteToHex: byteToHex
	};

	window['AscCommon'].ColorUtils = CU;
	window['AscCommon'].CU = CU;
})(window);

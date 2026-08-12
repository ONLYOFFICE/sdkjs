/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function (window, undefined) {

	const SVG_EMU_PER_MM = 36000;
	const SVG_TEXT_LINE_RATIO = 1.2;
	const SVG_TEXT_DESCENT_RATIO = 0.22;
	let svgCssRules = [];

	function parseSvgCss(cssText, rules) {
		if (!cssText) {
			return;
		}
		const text = cssText.replace(/\/\*[\s\S]*?\*\//g, "");
		const re = /([^{}]+)\{([^}]*)\}/g;
		let m;
		while ((m = re.exec(text)) !== null) {
			const decls = {};
			const parts = m[2].split(";");
			for (let i = 0; i < parts.length; ++i) {
				const idx = parts[i].indexOf(":");
				if (idx > 0) {
					decls[parts[i].substring(0, idx).trim()] = parts[i].substring(idx + 1).trim();
				}
			}
			const sels = m[1].split(",");
			for (let i = 0; i < sels.length; ++i) {
				const sel = sels[i].trim();
				if (sel) {
					rules.push({sel: sel, decls: decls});
				}
			}
		}
	}

	function svgCssSelectorMatches(node, sel) {
		const tag = node.tagName ? node.tagName.toLowerCase() : "";
		const id = node.getAttribute("id");
		const clsList = (node.getAttribute("class") || "").split(/\s+/);
		const tagMatch = sel.match(/^[a-zA-Z*][\w-]*/);
		if (tagMatch && tagMatch[0] !== "*" && tagMatch[0].toLowerCase() !== tag) {
			return false;
		}
		const classes = sel.match(/\.[\w-]+/g) || [];
		for (let i = 0; i < classes.length; ++i) {
			if (clsList.indexOf(classes[i].substring(1)) === -1) {
				return false;
			}
		}
		const ids = sel.match(/#[\w-]+/g) || [];
		for (let i = 0; i < ids.length; ++i) {
			if (id !== ids[i].substring(1)) {
				return false;
			}
		}
		return true;
	}

	function svgCssProps(node) {
		const out = {};
		for (let i = 0; i < svgCssRules.length; ++i) {
			if (svgCssSelectorMatches(node, svgCssRules[i].sel)) {
				const d = svgCssRules[i].decls;
				for (const k in d) {
					if (d.hasOwnProperty(k)) {
						out[k] = d[k];
					}
				}
			}
		}
		return out;
	}

	const svgNamedColors = {
		"black": [0, 0, 0], "white": [255, 255, 255], "red": [255, 0, 0],
		"green": [0, 128, 0], "blue": [0, 0, 255], "yellow": [255, 255, 0],
		"gray": [128, 128, 128], "grey": [128, 128, 128], "silver": [192, 192, 192],
		"maroon": [128, 0, 0], "olive": [128, 128, 0], "lime": [0, 255, 0],
		"aqua": [0, 255, 255], "cyan": [0, 255, 255], "teal": [0, 128, 128],
		"navy": [0, 0, 128], "fuchsia": [255, 0, 255], "magenta": [255, 0, 255],
		"purple": [128, 0, 128], "orange": [255, 165, 0]
	};

	function svgArcToCubics(startX, startY, endX, endY, rx, ry, rotation, largeArc, sweep) {
		rx = Math.abs(rx);
		ry = Math.abs(ry);
		if (rx === 0 || ry === 0 || (startX === endX && startY === endY)) {
			return [{type: "lineto", x: endX, y: endY}];
		}
		const phi = (rotation * Math.PI) / 180;
		const cosPhi = Math.cos(phi);
		const sinPhi = Math.sin(phi);
		const dx = (startX - endX) / 2;
		const dy = (startY - endY) / 2;
		const x1p = cosPhi * dx + sinPhi * dy;
		const y1p = -sinPhi * dx + cosPhi * dy;

		let rx2 = rx * rx, ry2 = ry * ry;
		const lambda = (x1p * x1p) / rx2 + (y1p * y1p) / ry2;
		if (lambda > 1) {
			const s = Math.sqrt(lambda);
			rx *= s;
			ry *= s;
			rx2 = rx * rx;
			ry2 = ry * ry;
		}

		let factor = (rx2 * ry2 - rx2 * y1p * y1p - ry2 * x1p * x1p) / (rx2 * y1p * y1p + ry2 * x1p * x1p);
		factor = Math.sqrt(Math.max(0, factor));
		if (largeArc === sweep) {
			factor = -factor;
		}
		const cxp = (factor * rx * y1p) / ry;
		const cyp = (-factor * ry * x1p) / rx;
		const cx = cosPhi * cxp - sinPhi * cyp + (startX + endX) / 2;
		const cy = sinPhi * cxp + cosPhi * cyp + (startY + endY) / 2;

		const ux = (x1p - cxp) / rx;
		const uy = (y1p - cyp) / ry;
		const vx = (-x1p - cxp) / rx;
		const vy = (-y1p - cyp) / ry;
		let theta1 = Math.atan2(uy, ux);
		let theta2 = Math.atan2(vy, vx);
		if (sweep === 0 && theta2 > theta1) {
			theta2 -= 2 * Math.PI;
		} else if (sweep !== 0 && theta2 < theta1) {
			theta2 += 2 * Math.PI;
		}

		const segments = Math.max(1, Math.ceil(Math.abs(theta2 - theta1) / (Math.PI / 2)));
		const delta = (theta2 - theta1) / segments;
		const alpha = (Math.sin(delta) * (Math.sqrt(4 + 3 * Math.pow(Math.tan(delta / 2), 2)) - 1)) / 3;

		const commands = [];
		let theta = theta1;
		for (let i = 0; i < segments; ++i) {
			const next = theta + delta;
			const cosT1 = Math.cos(theta), sinT1 = Math.sin(theta);
			const cosT2 = Math.cos(next), sinT2 = Math.sin(next);

			const p1x = cx + rx * cosT1 * cosPhi - ry * sinT1 * sinPhi;
			const p1y = cy + rx * cosT1 * sinPhi + ry * sinT1 * cosPhi;
			const p2x = cx + rx * cosT2 * cosPhi - ry * sinT2 * sinPhi;
			const p2y = cy + rx * cosT2 * sinPhi + ry * sinT2 * cosPhi;

			const d1x = -rx * sinT1 * cosPhi - ry * cosT1 * sinPhi;
			const d1y = -rx * sinT1 * sinPhi + ry * cosT1 * cosPhi;
			const d2x = -rx * sinT2 * cosPhi - ry * cosT2 * sinPhi;
			const d2y = -rx * sinT2 * sinPhi + ry * cosT2 * cosPhi;

			commands.push({
				type: "cubicBezier",
				x1: p1x + alpha * d1x, y1: p1y + alpha * d1y,
				x2: p2x - alpha * d2x, y2: p2y - alpha * d2y,
				x: p2x, y: p2y
			});
			theta = next;
		}
		return commands;
	}

	function parseSvgPathData(d) {
		const commands = [];
		if (typeof d !== "string" || d.length === 0) {
			return commands;
		}
		const regex = /([MLHVCSQTAZ])([^MLHVCSQTAZ]*)/gi;
		let curX = 0, curY = 0, startSubX = 0, startSubY = 0;
		let prevType = "", prevC2X = 0, prevC2Y = 0, prevQX = 0, prevQY = 0;
		let match;
		while ((match = regex.exec(d)) !== null) {
			const letter = match[1];
			const upper = letter.toUpperCase();
			const absolute = (letter === upper);
			const nums = [];
			const numRe = /[+-]?(?:\d*\.\d+|\d+\.?)(?:[eE][+-]?\d+)?/g;
			let nm;
			while ((nm = numRe.exec(match[2])) !== null) {
				nums.push(parseFloat(nm[0]));
			}

			if (upper === "M") {
				for (let i = 0; i + 1 < nums.length; i += 2) {
					let x = nums[i], y = nums[i + 1];
					if (!absolute) { x += curX; y += curY; }
					if (i === 0) {
						commands.push({type: "moveto", x: x, y: y});
						startSubX = x; startSubY = y;
					} else {
						commands.push({type: "lineto", x: x, y: y});
					}
					curX = x; curY = y;
				}
				prevType = "M";
			} else if (upper === "L") {
				for (let i = 0; i + 1 < nums.length; i += 2) {
					let x = nums[i], y = nums[i + 1];
					if (!absolute) { x += curX; y += curY; }
					commands.push({type: "lineto", x: x, y: y});
					curX = x; curY = y;
				}
				prevType = "L";
			} else if (upper === "H") {
				for (let i = 0; i < nums.length; ++i) {
					let x = nums[i];
					if (!absolute) { x += curX; }
					commands.push({type: "lineto", x: x, y: curY});
					curX = x;
				}
				prevType = "H";
			} else if (upper === "V") {
				for (let i = 0; i < nums.length; ++i) {
					let y = nums[i];
					if (!absolute) { y += curY; }
					commands.push({type: "lineto", x: curX, y: y});
					curY = y;
				}
				prevType = "V";
			} else if (upper === "C") {
				for (let i = 0; i + 5 < nums.length; i += 6) {
					let x1 = nums[i], y1 = nums[i + 1], x2 = nums[i + 2], y2 = nums[i + 3], x = nums[i + 4], y = nums[i + 5];
					if (!absolute) { x1 += curX; y1 += curY; x2 += curX; y2 += curY; x += curX; y += curY; }
					commands.push({type: "cubicBezier", x1: x1, y1: y1, x2: x2, y2: y2, x: x, y: y});
					prevC2X = x2; prevC2Y = y2;
					curX = x; curY = y;
				}
				prevType = "C";
			} else if (upper === "S") {
				for (let i = 0; i + 3 < nums.length; i += 4) {
					let x2 = nums[i], y2 = nums[i + 1], x = nums[i + 2], y = nums[i + 3];
					if (!absolute) { x2 += curX; y2 += curY; x += curX; y += curY; }
					let x1 = curX, y1 = curY;
					if (prevType === "C" || prevType === "S") {
						x1 = 2 * curX - prevC2X;
						y1 = 2 * curY - prevC2Y;
					}
					commands.push({type: "cubicBezier", x1: x1, y1: y1, x2: x2, y2: y2, x: x, y: y});
					prevC2X = x2; prevC2Y = y2;
					curX = x; curY = y;
					prevType = "S";
				}
				prevType = "S";
			} else if (upper === "Q") {
				for (let i = 0; i + 3 < nums.length; i += 4) {
					let x1 = nums[i], y1 = nums[i + 1], x = nums[i + 2], y = nums[i + 3];
					if (!absolute) { x1 += curX; y1 += curY; x += curX; y += curY; }
					commands.push({type: "quadraticBezier", x1: x1, y1: y1, x2: x, y2: y});
					prevQX = x1; prevQY = y1;
					curX = x; curY = y;
				}
				prevType = "Q";
			} else if (upper === "T") {
				for (let i = 0; i + 1 < nums.length; i += 2) {
					let x = nums[i], y = nums[i + 1];
					if (!absolute) { x += curX; y += curY; }
					let x1 = curX, y1 = curY;
					if (prevType === "Q" || prevType === "T") {
						x1 = 2 * curX - prevQX;
						y1 = 2 * curY - prevQY;
					}
					commands.push({type: "quadraticBezier", x1: x1, y1: y1, x2: x, y2: y});
					prevQX = x1; prevQY = y1;
					curX = x; curY = y;
					prevType = "T";
				}
				prevType = "T";
			} else if (upper === "A") {
				for (let i = 0; i + 6 < nums.length; i += 7) {
					const rx = nums[i], ry = nums[i + 1], rot = nums[i + 2];
					const largeArc = nums[i + 3] ? 1 : 0, sweep = nums[i + 4] ? 1 : 0;
					let x = nums[i + 5], y = nums[i + 6];
					if (!absolute) { x += curX; y += curY; }
					const arcCmds = svgArcToCubics(curX, curY, x, y, rx, ry, rot, largeArc, sweep);
					for (let j = 0; j < arcCmds.length; ++j) {
						commands.push(arcCmds[j]);
					}
					curX = x; curY = y;
				}
				prevType = "A";
			} else if (upper === "Z") {
				commands.push({type: "closepath"});
				curX = startSubX; curY = startSubY;
				prevType = "Z";
			}
		}
		return commands;
	}

	function svgRectToPath(el) {
		const x = parseFloat(el.getAttribute("x")) || 0;
		const y = parseFloat(el.getAttribute("y")) || 0;
		const w = parseFloat(el.getAttribute("width")) || 0;
		const h = parseFloat(el.getAttribute("height")) || 0;
		let rx = parseFloat(el.getAttribute("rx"));
		let ry = parseFloat(el.getAttribute("ry"));
		if (isNaN(rx) && isNaN(ry)) {
			return "M " + x + " " + y + " L " + (x + w) + " " + y + " L " + (x + w) + " " + (y + h) + " L " + x + " " + (y + h) + " Z";
		}
		if (isNaN(rx)) { rx = ry; }
		if (isNaN(ry)) { ry = rx; }
		rx = Math.min(rx, w / 2);
		ry = Math.min(ry, h / 2);
		return "M " + (x + rx) + " " + y +
			" L " + (x + w - rx) + " " + y +
			" A " + rx + " " + ry + " 0 0 1 " + (x + w) + " " + (y + ry) +
			" L " + (x + w) + " " + (y + h - ry) +
			" A " + rx + " " + ry + " 0 0 1 " + (x + w - rx) + " " + (y + h) +
			" L " + (x + rx) + " " + (y + h) +
			" A " + rx + " " + ry + " 0 0 1 " + x + " " + (y + h - ry) +
			" L " + x + " " + (y + ry) +
			" A " + rx + " " + ry + " 0 0 1 " + (x + rx) + " " + y + " Z";
	}

	function svgCircleToPath(el) {
		const cx = parseFloat(el.getAttribute("cx")) || 0;
		const cy = parseFloat(el.getAttribute("cy")) || 0;
		const r = parseFloat(el.getAttribute("r")) || 0;
		return "M " + (cx - r) + " " + cy + " a " + r + " " + r + " 0 1 0 " + (r * 2) + " 0 a " + r + " " + r + " 0 1 0 " + (-r * 2) + " 0 Z";
	}

	function svgEllipseToPath(el) {
		const cx = parseFloat(el.getAttribute("cx")) || 0;
		const cy = parseFloat(el.getAttribute("cy")) || 0;
		const rx = parseFloat(el.getAttribute("rx")) || 0;
		const ry = parseFloat(el.getAttribute("ry")) || 0;
		return "M " + (cx - rx) + " " + cy + " a " + rx + " " + ry + " 0 1 0 " + (rx * 2) + " 0 a " + rx + " " + ry + " 0 1 0 " + (-rx * 2) + " 0 Z";
	}

	function svgPointsToPath(el, close) {
		const raw = (el.getAttribute("points") || "").trim();
		if (!raw) {
			return "";
		}
		const nums = raw.split(/[\s,]+/);
		let d = "";
		for (let i = 0; i + 1 < nums.length; i += 2) {
			d += (i === 0 ? "M " : "L ") + nums[i] + " " + nums[i + 1] + " ";
		}
		return close ? d + "Z" : d;
	}

	function parseSvgColor(value, currentColor) {
		if (!value) {
			return null;
		}
		value = ("" + value).trim().toLowerCase();
		if (value === "none" || value === "transparent") {
			return false;
		}
		if (value === "currentcolor") {
			return currentColor ? parseSvgColor(currentColor, null) : [0, 0, 0];
		}
		if (value.charAt(0) === "#") {
			let hex = value.substring(1);
			if (hex.length === 3) {
				hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
			}
			if (hex.length >= 6) {
				const r = parseInt(hex.substring(0, 2), 16);
				const g = parseInt(hex.substring(2, 4), 16);
				const b = parseInt(hex.substring(4, 6), 16);
				if (!isNaN(r) && !isNaN(g) && !isNaN(b)) {
					return [r, g, b];
				}
			}
			return null;
		}
		if (value.indexOf("rgb") === 0) {
			const m = value.match(/[\d.]+/g);
			if (m && m.length >= 3) {
				return [Math.round(parseFloat(m[0])), Math.round(parseFloat(m[1])), Math.round(parseFloat(m[2]))];
			}
			return null;
		}
		if (svgNamedColors[value]) {
			return svgNamedColors[value];
		}
		return null;
	}

	function parseSvgTransform(str) {
		let m = [1, 0, 0, 1, 0, 0];
		if (!str) {
			return m;
		}
		const re = /(matrix|translate|scale|rotate|skewX|skewY)\s*\(([^)]*)\)/g;
		let t;
		while ((t = re.exec(str)) !== null) {
			const name = t[1];
			const a = t[2].split(/[\s,]+/).filter(function (s) { return s !== ""; }).map(parseFloat);
			let cm;
			if (name === "matrix" && a.length >= 6) {
				cm = [a[0], a[1], a[2], a[3], a[4], a[5]];
			} else if (name === "translate") {
				cm = [1, 0, 0, 1, a[0] || 0, a.length > 1 ? a[1] : 0];
			} else if (name === "scale") {
				cm = [a[0] || 1, 0, 0, (a.length > 1 ? a[1] : a[0]) || 1, 0, 0];
			} else if (name === "rotate") {
				const ang = (a[0] || 0) * Math.PI / 180, cos = Math.cos(ang), sin = Math.sin(ang);
				if (a.length >= 3) {
					const cx = a[1], cy = a[2];
					cm = [cos, sin, -sin, cos, cx - cos * cx + sin * cy, cy - sin * cx - cos * cy];
				} else {
					cm = [cos, sin, -sin, cos, 0, 0];
				}
			} else if (name === "skewX") {
				cm = [1, 0, Math.tan((a[0] || 0) * Math.PI / 180), 1, 0, 0];
			} else if (name === "skewY") {
				cm = [1, Math.tan((a[0] || 0) * Math.PI / 180), 0, 1, 0, 0];
			} else {
				cm = [1, 0, 0, 1, 0, 0];
			}
			m = matMulSvg(m, cm);
		}
		return m;
	}

	function matMulSvg(m1, m2) {
		return [
			m1[0] * m2[0] + m1[2] * m2[1],
			m1[1] * m2[0] + m1[3] * m2[1],
			m1[0] * m2[2] + m1[2] * m2[3],
			m1[1] * m2[2] + m1[3] * m2[3],
			m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
			m1[1] * m2[4] + m1[3] * m2[5] + m1[5]
		];
	}

	function matScaleSvg(m) {
		const sx = Math.sqrt(m[0] * m[0] + m[1] * m[1]);
		const sy = Math.sqrt(m[2] * m[2] + m[3] * m[3]);
		return ((sx + sy) / 2) || 1;
	}

	function transformSvgCommands(commands, m) {
		function tx(x, y) { return m[0] * x + m[2] * y + m[4]; }
		function ty(x, y) { return m[1] * x + m[3] * y + m[5]; }
		const out = [];
		for (let i = 0; i < commands.length; ++i) {
			const c = commands[i];
			if (c.type === "moveto" || c.type === "lineto") {
				out.push({type: c.type, x: tx(c.x, c.y), y: ty(c.x, c.y)});
			} else if (c.type === "cubicBezier") {
				out.push({type: c.type, x1: tx(c.x1, c.y1), y1: ty(c.x1, c.y1), x2: tx(c.x2, c.y2), y2: ty(c.x2, c.y2), x: tx(c.x, c.y), y: ty(c.x, c.y)});
			} else if (c.type === "quadraticBezier") {
				out.push({type: c.type, x1: tx(c.x1, c.y1), y1: ty(c.x1, c.y1), x2: tx(c.x2, c.y2), y2: ty(c.x2, c.y2)});
			} else {
				out.push(c);
			}
		}
		return out;
	}

	function parseSvgInlineStyle(node) {
		const obj = {};
		const inline = node.getAttribute("style");
		if (inline) {
			const parts = inline.split(";");
			for (let i = 0; i < parts.length; ++i) {
				const idx = parts[i].indexOf(":");
				if (idx > 0) {
					obj[parts[i].substring(0, idx).trim()] = parts[i].substring(idx + 1).trim();
				}
			}
		}
		return obj;
	}

	function inheritSvgStyle(node, parent) {
		const inlineObj = parseSvgInlineStyle(node);
		const cssObj = svgCssProps(node);
		function prop(name) {
			if (inlineObj[name] !== undefined) {
				return inlineObj[name];
			}
			if (cssObj[name] !== undefined) {
				return cssObj[name];
			}
			const a = node.getAttribute(name);
			return (a !== null && a !== undefined && a !== "") ? a : undefined;
		}
		const s = {
			fill: parent.fill, stroke: parent.stroke, strokeWidth: parent.strokeWidth,
			color: parent.color, opacity: parent.opacity, fillOpacity: parent.fillOpacity,
			strokeOpacity: parent.strokeOpacity, fillRule: parent.fillRule, strokeDash: parent.strokeDash,
			strokeCap: parent.strokeCap, clip: parent.clip, display: true
		};
		let v;
		if ((v = prop("color")) !== undefined) { s.color = v; }
		if ((v = prop("fill")) !== undefined) { s.fill = v; }
		if ((v = prop("stroke")) !== undefined) { s.stroke = v; }
		if ((v = prop("stroke-width")) !== undefined) { s.strokeWidth = parseFloat(v); }
		if ((v = prop("stroke-dasharray")) !== undefined) { s.strokeDash = v; }
		if ((v = prop("stroke-linecap")) !== undefined) { s.strokeCap = v; }
		if ((v = prop("fill-opacity")) !== undefined) { s.fillOpacity = parseFloat(v); }
		if ((v = prop("stroke-opacity")) !== undefined) { s.strokeOpacity = parseFloat(v); }
		if ((v = prop("fill-rule")) !== undefined) { s.fillRule = v; }
		if ((v = prop("opacity")) !== undefined) { s.opacity = parent.opacity * parseFloat(v); }
		if (prop("display") === "none") { s.display = false; }
		v = prop("visibility");
		if (v === "hidden" || v === "collapse") { s.display = false; }
		return s;
	}

	function svgElementChildren(node) {
		const out = [];
		const ch = node.childNodes;
		if (ch) {
			for (let i = 0; i < ch.length; ++i) {
				if (ch[i].nodeType === 1) {
					out.push(ch[i]);
				}
			}
		}
		return out;
	}

	function svgGradientStops(gradNode, currentColor) {
		const stops = gradNode.getElementsByTagName("stop");
		const out = [];
		for (let i = 0; i < stops.length; ++i) {
			const inlineObj = parseSvgInlineStyle(stops[i]);
			const sc = (inlineObj["stop-color"] !== undefined) ? inlineObj["stop-color"] : stops[i].getAttribute("stop-color");
			const col = parseSvgColor(sc, currentColor) || [0, 0, 0];
			if (col.length !== 3) {
				continue;
			}
			let off = inlineObj["offset"] !== undefined ? inlineObj["offset"] : stops[i].getAttribute("offset");
			let pos = parseFloat(off);
			if (isNaN(pos)) {
				pos = 0;
			} else if (("" + off).indexOf("%") !== -1) {
				pos = pos / 100;
			}
			let so = (inlineObj["stop-opacity"] !== undefined) ? inlineObj["stop-opacity"] : stops[i].getAttribute("stop-opacity");
			let alpha = parseFloat(so);
			alpha = isNaN(alpha) ? 255 : Math.round(Math.max(0, Math.min(1, alpha)) * 255);
			out.push({pos: Math.max(0, Math.min(1, pos)), rgb: col, alpha: alpha});
		}
		return out;
	}

	function svgResolveGradient(value, idMap, currentColor) {
		if (value === undefined || value === null) {
			return null;
		}
		const v = ("" + value).trim();
		if (v.indexOf("url(") !== 0) {
			return null;
		}
		const m = v.match(/#([^)\s'"]+)/);
		if (!m || !idMap[m[1]]) {
			return null;
		}
		const ref = idMap[m[1]];
		const tag = ref.tagName ? ref.tagName.toLowerCase() : "";
		if (tag.indexOf("gradient") === -1) {
			return null;
		}
		const stops = svgGradientStops(ref, currentColor);
		if (stops.length === 0) {
			return null;
		}
		let angle = 0;
		const radial = (tag === "radialgradient");
		if (tag === "lineargradient") {
			const x1 = parseFloat(ref.getAttribute("x1")) || 0;
			const y1 = parseFloat(ref.getAttribute("y1")) || 0;
			const x2 = ref.getAttribute("x2") !== null ? parseFloat(ref.getAttribute("x2")) : 1;
			const y2 = parseFloat(ref.getAttribute("y2")) || 0;
			let deg = Math.atan2(y2 - y1, x2 - x1) * 180 / Math.PI;
			if (deg < 0) {
				deg += 360;
			}
			angle = Math.round(deg * 60000);
		}
		return {stops: stops, angle: angle, radial: radial};
	}

	function svgSolidFill(rgb, alpha255) {
		const a = (alpha255 === undefined || alpha255 === null) ? 255 : alpha255;
		const fill = AscFormat.CreateSolidFillRGB(rgb[0], rgb[1], rgb[2]);
		if (a < 255) {
			fill.transparent = a;
			fill.addColorMod(new AscFormat.CColorMod("alpha", Math.max(0, Math.round(a / 255 * 100000))));
		}
		return fill;
	}

	function svgPatternFill(rgb) {
		const ftype = AscCommon.global_hatch_offsets ? AscCommon.global_hatch_offsets["dotGrid"] : undefined;
		if (ftype === undefined) {
			return svgSolidFill(rgb, 255);
		}
		const uni = new AscFormat.CUniFill();
		uni.fill = new AscFormat.CPattFill();
		uni.fill.ftype = ftype;
		uni.fill.fgClr = AscFormat.CreateUniColorRGB(rgb[0], rgb[1], rgb[2]);
		uni.fill.bgClr = AscFormat.CreateUniColorRGB(255, 255, 255);
		return uni;
	}

	function svgResolvePattern(value, idMap, currentColor) {
		if (value === undefined || value === null) {
			return null;
		}
		const v = ("" + value).trim();
		if (v.indexOf("url(") !== 0) {
			return null;
		}
		const m = v.match(/#([^)\s'"]+)/);
		if (!m || !idMap[m[1]]) {
			return null;
		}
		const ref = idMap[m[1]];
		if (!ref.tagName || ref.tagName.toLowerCase() !== "pattern") {
			return null;
		}
		let col = null;
		const child = ref.querySelector("[fill]");
		if (child) {
			col = parseSvgColor(child.getAttribute("fill"), currentColor);
		}
		return {rgb: (col && col.length === 3) ? col : [128, 128, 128]};
	}

	function svgResolveMask(value, idMap, currentColor) {
		if (!value) {
			return null;
		}
		const v = ("" + value).trim();
		if (v.indexOf("url(") !== 0) {
			return null;
		}
		const m = v.match(/#([^)\s'"]+)/);
		if (!m || !idMap[m[1]]) {
			return null;
		}
		const maskNode = idMap[m[1]];
		if (!maskNode.tagName || maskNode.tagName.toLowerCase() !== "mask") {
			return null;
		}
		const el = maskNode.querySelector("[fill]");
		if (!el) {
			return null;
		}
		const grad = svgResolveGradient(el.getAttribute("fill"), idMap, currentColor);
		if (!grad) {
			return null;
		}
		const stops = [];
		for (let i = 0; i < grad.stops.length; ++i) {
			const s = grad.stops[i];
			const lum = (0.299 * s.rgb[0] + 0.587 * s.rgb[1] + 0.114 * s.rgb[2]) / 255;
			const sa = (s.alpha === undefined || s.alpha === null) ? 1 : s.alpha / 255;
			stops.push({pos: s.pos, alpha: Math.round(lum * sa * 255)});
		}
		return {stops: stops, angle: grad.angle, radial: grad.radial};
	}

	function buildSvgGradFill(grad) {
		const uni = new AscFormat.CUniFill();
		const gf = new AscFormat.CGradFill();
		uni.setFill(gf);
		gf.rotateWithShape = true;
		if (grad.radial) {
			gf.path = new AscFormat.GradPath();
		} else {
			const lin = new AscFormat.GradLin();
			gf.setLin(lin);
			lin.setAngle(grad.angle || 0);
			lin.setScale(true);
		}
		const stops = grad.stops;
		for (let i = 0; i < stops.length; ++i) {
			const gs = new AscFormat.CGs();
			gf.addColor(gs);
			gs.setPos(Math.round(stops[i].pos * 100000));
			const col = AscFormat.CreateUniColorRGB(stops[i].rgb[0], stops[i].rgb[1], stops[i].rgb[2]);
			const a = (stops[i].alpha === undefined || stops[i].alpha === null) ? 255 : stops[i].alpha;
			if (a < 255) {
				col.addColorMod(new AscFormat.CColorMod("alpha", Math.round(a / 255 * 100000)));
				if (col.RGBA) {
					col.RGBA.A = a;
				}
			}
			gs.setColor(col);
		}
		return uni;
	}

	function resolveSvgPaint(value, idMap, currentColor) {
		if (value === undefined || value === null) {
			return null;
		}
		const v = ("" + value).trim();
		if (v.indexOf("url(") === 0) {
			const m = v.match(/#([^)\s'"]+)/);
			if (m && idMap[m[1]]) {
				const ref = idMap[m[1]];
				const tag = ref.tagName ? ref.tagName.toLowerCase() : "";
				if (tag.indexOf("gradient") !== -1) {
					const stops = svgGradientStops(ref, currentColor);
					if (stops.length > 0) {
						let r = 0, g = 0, b = 0;
						for (let i = 0; i < stops.length; ++i) { r += stops[i].rgb[0]; g += stops[i].rgb[1]; b += stops[i].rgb[2]; }
						return [Math.round(r / stops.length), Math.round(g / stops.length), Math.round(b / stops.length)];
					}
					return null;
				}
				const child = ref.querySelector("[fill]");
				if (child) {
					return parseSvgColor(child.getAttribute("fill"), currentColor);
				}
			}
			return null;
		}
		return parseSvgColor(v, currentColor);
	}

	function svgClampAlpha(a) {
		if (isNaN(a)) { return 255; }
		return Math.round(Math.max(0, Math.min(1, a)) * 255);
	}

	function svgDashPreset(dash, cap, strokeW) {
		if (!dash || ("" + dash).trim() === "none") {
			return null;
		}
		const nums = ("" + dash).match(/[\d.]+/g);
		if (!nums || nums.length === 0) {
			return null;
		}
		const a = nums.map(parseFloat);
		let maxGap = 0;
		for (let i = 1; i < a.length; i += 2) {
			maxGap = Math.max(maxGap, a[i]);
		}
		if ((cap === "square" || cap === "round") && maxGap <= (strokeW || 0)) {
			return null;
		}
		if (a.length >= 4) {
			return "dashDot";
		}
		const dashLen = a[0];
		if (dashLen <= 2) {
			return "sysDot";
		}
		if (dashLen >= 8) {
			return "lgDash";
		}
		return "dash";
	}

	function pushSvgShapeElement(node, d, matrix, style, idMap, elements) {
		if (!d) {
			return;
		}
		let commands = parseSvgPathData(d);
		if (commands.length === 0) {
			return;
		}
		commands = transformSvgCommands(commands, matrix);
		const fillGrad = svgResolveGradient(style.fill, idMap, style.color);
		const fillPattern = fillGrad ? null : svgResolvePattern(style.fill, idMap, style.color);
		const fillCol = (fillGrad || fillPattern) ? null : resolveSvgPaint(style.fill, idMap, style.color);
		const strokeCol = (style.stroke === null || style.stroke === undefined) ? null : resolveSvgPaint(style.stroke, idMap, style.color);
		const el = {commands: commands};
		if (fillGrad) {
			el.fillGrad = fillGrad;
		} else if (fillPattern) {
			el.fillPattern = fillPattern;
		} else if (fillCol && fillCol.length === 3) {
			el.fill = fillCol;
			el.fillAlpha = svgClampAlpha(style.opacity * style.fillOpacity);
		}
		if (strokeCol && strokeCol.length === 3) {
			el.stroke = strokeCol;
			el.strokeAlpha = svgClampAlpha(style.opacity * style.strokeOpacity);
			el.strokeWidth = (style.strokeWidth || 1) * matScaleSvg(matrix);
			el.strokeDash = svgDashPreset(style.strokeDash, ("" + (style.strokeCap || "")).toLowerCase(), style.strokeWidth || 1);
		}
		if (style.clip) {
			el.clip = style.clip;
		}
		if (style.mask && style.mask.stops.length > 0 && el.fill && !el.fillGrad) {
			const mstops = style.mask.stops;
			const gstops = [];
			for (let i = 0; i < mstops.length; ++i) {
				gstops.push({pos: mstops[i].pos, rgb: el.fill, alpha: mstops[i].alpha});
			}
			el.fillGrad = {stops: gstops, angle: style.mask.angle, radial: style.mask.radial};
			el.fill = null;
			el.fillAlpha = null;
		}
		if (el.fill || el.fillGrad || el.fillPattern || el.stroke) {
			elements.push(el);
		}
	}

	const svgSkipTags = {
		defs: 1, symbol: 1, clippath: 1, mask: 1, pattern: 1, marker: 1, lineargradient: 1,
		radialgradient: 1, stop: 1, style: 1, title: 1, desc: 1, metadata: 1, filter: 1,
		tspan: 1, textpath: 1
	};

	function resolveSvgFontFamily(value) {
		return AscFonts.g_fontApplication.GetFontNameDictionary("" + (value || ""), true);
	}

	function svgTextPropGetter(node) {
		const inlineObj = parseSvgInlineStyle(node);
		const cssObj = svgCssProps(node);
		return function (name) {
			if (inlineObj[name] !== undefined) {
				return inlineObj[name];
			}
			if (cssObj[name] !== undefined) {
				return cssObj[name];
			}
			const attr = node.getAttribute(name);
			return (attr !== null && attr !== undefined && attr !== "") ? attr : undefined;
		};
	}

	function svgTextLength(raw, fontSize) {
		if (raw === null || raw === undefined || ("" + raw).trim() === "") {
			return undefined;
		}
		const text = ("" + raw).trim();
		const value = parseFloat(text);
		if (isNaN(value)) {
			return undefined;
		}
		if (text.indexOf("em") !== -1) {
			return value * fontSize;
		}
		if (text.indexOf("ex") !== -1) {
			return value * fontSize / 2;
		}
		return value;
	}

	function inheritSvgTextFormat(node, parent) {
		const prop = svgTextPropGetter(node);
		const format = {
			fontSize: parent.fontSize, fontFamily: parent.fontFamily, weight: parent.weight,
			fontStyle: parent.fontStyle, fill: parent.fill, color: parent.color, anchor: parent.anchor
		};
		const fontSize = svgTextLength(prop("font-size"), parent.fontSize);
		if (fontSize > 0) {
			format.fontSize = fontSize;
		}
		let v;
		if ((v = prop("font-family")) !== undefined) { format.fontFamily = v; }
		if ((v = prop("font-weight")) !== undefined) { format.weight = ("" + v).toLowerCase(); }
		if ((v = prop("font-style")) !== undefined) { format.fontStyle = ("" + v).toLowerCase(); }
		if ((v = prop("fill")) !== undefined) { format.fill = v; }
		if ((v = prop("color")) !== undefined) { format.color = v; }
		if ((v = prop("text-anchor")) !== undefined) { format.anchor = ("" + v).toLowerCase(); }
		return format;
	}

	function buildSvgTextRun(text, format, matrix) {
		const fillCol = resolveSvgPaint(format.fill, {}, format.color);
		return {
			text: text,
			fontSize: format.fontSize * matScaleSvg(matrix),
			fontFamily: resolveSvgFontFamily(format.fontFamily),
			bold: (format.weight === "bold" || format.weight === "bolder" || parseInt(format.weight, 10) >= 600),
			italic: (format.fontStyle === "italic" || format.fontStyle === "oblique"),
			color: (fillCol && fillCol.length === 3) ? fillCol : [0, 0, 0]
		};
	}

	function isSameSvgTextRun(first, second) {
		return first.fontSize === second.fontSize && first.fontFamily === second.fontFamily &&
			first.bold === second.bold && first.italic === second.italic &&
			first.color[0] === second.color[0] && first.color[1] === second.color[1] && first.color[2] === second.color[2];
	}

	function addSvgTextRun(state, format, rawText) {
		const text = ("" + rawText).replace(/\s+/g, " ");
		if (text.length === 0) {
			return;
		}
		const matrix = state.matrix;
		if (!state.current) {
			state.current = {
				isText: true,
				runs: [],
				x: matrix[0] * state.x + matrix[2] * state.y + matrix[4],
				y: matrix[1] * state.x + matrix[3] * state.y + matrix[5],
				anchor: format.anchor,
				rot: Math.atan2(matrix[1], matrix[0])
			};
		}
		const runs = state.current.runs;
		const run = buildSvgTextRun(text, format, matrix);
		const last = runs[runs.length - 1];
		if (last && isSameSvgTextRun(last, run)) {
			last.text += run.text;
			return;
		}
		runs.push(run);
	}

	function flushSvgTextElement(state) {
		const element = state.current;
		state.current = null;
		if (!element) {
			return;
		}
		const runs = element.runs;
		runs[0].text = runs[0].text.replace(/^\s+/, "");
		runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, "");
		const kept = [];
		let fontSize = 0;
		for (let i = 0; i < runs.length; ++i) {
			if (runs[i].text.length > 0) {
				kept.push(runs[i]);
				fontSize = Math.max(fontSize, runs[i].fontSize);
			}
		}
		if (kept.length === 0) {
			return;
		}
		element.runs = kept;
		element.fontSize = fontSize;
		state.elements.push(element);
	}

	function moveSvgTextPos(state, node, fontSize) {
		const x = svgTextLength(node.getAttribute("x"), fontSize);
		const y = svgTextLength(node.getAttribute("y"), fontSize);
		const dy = svgTextLength(node.getAttribute("dy"), fontSize);
		const newX = (x !== undefined) ? x : state.x;
		const newY = (y !== undefined) ? y : ((dy !== undefined) ? state.y + dy : state.y);
		if (newX === state.x && newY === state.y) {
			return;
		}
		flushSvgTextElement(state);
		state.x = newX;
		state.y = newY;
	}

	function walkSvgTextNode(node, format, state) {
		const kids = node.childNodes;
		if (!kids) {
			return;
		}
		for (let i = 0; i < kids.length; ++i) {
			const kid = kids[i];
			if (kid.nodeType === 3) {
				addSvgTextRun(state, format, kid.nodeValue);
				continue;
			}
			if (kid.nodeType !== 1) {
				continue;
			}
			const tag = kid.tagName ? kid.tagName.toLowerCase() : "";
			if (tag !== "tspan" && tag !== "textpath" && tag !== "a") {
				continue;
			}
			const childFormat = inheritSvgTextFormat(kid, format);
			moveSvgTextPos(state, kid, childFormat.fontSize);
			walkSvgTextNode(kid, childFormat, state);
		}
	}

	function pushSvgTextElements(node, matrix, style, elements) {
		const format = inheritSvgTextFormat(node, {
			fontSize: 16, fontFamily: undefined, weight: "", fontStyle: "",
			fill: style.fill, color: style.color, anchor: "start"
		});
		const state = {
			matrix: matrix,
			elements: elements,
			current: null,
			x: svgTextLength(node.getAttribute("x"), format.fontSize) || 0,
			y: svgTextLength(node.getAttribute("y"), format.fontSize) || 0
		};
		walkSvgTextNode(node, format, state);
		flushSvgTextElement(state);
	}

	function svgResolveClip(value, idMap, matrix) {
		if (!value) {
			return null;
		}
		const v = ("" + value).trim();
		if (v.indexOf("url(") !== 0) {
			return null;
		}
		const m = v.match(/#([^)\s'"]+)/);
		if (!m || !idMap[m[1]]) {
			return null;
		}
		const cpNode = idMap[m[1]];
		if (!cpNode.tagName || cpNode.tagName.toLowerCase() !== "clippath") {
			return null;
		}
		const kids = svgElementChildren(cpNode);
		let commands = [];
		for (let i = 0; i < kids.length; ++i) {
			const kid = kids[i];
			const ktag = kid.tagName ? kid.tagName.toLowerCase() : "";
			let d = null;
			if (ktag === "path") { d = kid.getAttribute("d"); }
			else if (ktag === "rect") { d = svgRectToPath(kid); }
			else if (ktag === "circle") { d = svgCircleToPath(kid); }
			else if (ktag === "ellipse") { d = svgEllipseToPath(kid); }
			else if (ktag === "polygon") { d = svgPointsToPath(kid, true); }
			else if (ktag === "polyline") { d = svgPointsToPath(kid, false); }
			if (d) {
				const km = matMulSvg(matrix, parseSvgTransform(kid.getAttribute("transform")));
				commands = commands.concat(transformSvgCommands(parseSvgPathData(d), km));
			}
		}
		return commands.length ? commands : null;
	}

	function walkSvgNode(node, parentMatrix, parentStyle, idMap, elements, depth) {
		if (depth > 24) {
			return;
		}
		const tag = node.tagName ? node.tagName.toLowerCase() : "";
		if (svgSkipTags[tag]) {
			return;
		}
		const matrix = matMulSvg(parentMatrix, parseSvgTransform(node.getAttribute("transform")));
		const style = inheritSvgStyle(node, parentStyle);
		if (style.display === false) {
			return;
		}
		const clipRef = node.getAttribute("clip-path");
		if (clipRef) {
			const cc = svgResolveClip(clipRef, idMap, matrix);
			if (cc) {
				style.clip = cc;
			}
		}
		const maskRef = node.getAttribute("mask");
		style.mask = maskRef ? svgResolveMask(maskRef, idMap, style.color) : null;
		if (tag === "svg" && depth > 0) {
			walkSvgNestedViewport(node, matrix, style, idMap, elements, depth);
			return;
		}
		if (tag === "g" || tag === "a" || tag === "switch") {
			const kids = svgElementChildren(node);
			const groupChildren = [];
			for (let i = 0; i < kids.length; ++i) {
				walkSvgNode(kids[i], matrix, style, idMap, groupChildren, depth + 1);
			}
			if (groupChildren.length === 1) {
				elements.push(groupChildren[0]);
			} else if (groupChildren.length > 1) {
				elements.push({isGroupNode: true, children: groupChildren});
			}
			return;
		}
		if (tag === "svg") {
			const kids = svgElementChildren(node);
			for (let i = 0; i < kids.length; ++i) {
				walkSvgNode(kids[i], matrix, style, idMap, elements, depth + 1);
			}
			return;
		}
		if (tag === "use") {
			walkSvgUse(node, matrix, style, idMap, elements, depth);
			return;
		}
		if (tag === "text") {
			pushSvgTextElements(node, matrix, style, elements);
			return;
		}
		let d = null;
		if (tag === "path") {
			d = node.getAttribute("d");
		} else if (tag === "rect") {
			d = svgRectToPath(node);
		} else if (tag === "circle") {
			d = svgCircleToPath(node);
		} else if (tag === "ellipse") {
			d = svgEllipseToPath(node);
		} else if (tag === "line") {
			d = "M " + (node.getAttribute("x1") || 0) + " " + (node.getAttribute("y1") || 0) + " L " + (node.getAttribute("x2") || 0) + " " + (node.getAttribute("y2") || 0);
		} else if (tag === "polyline") {
			d = svgPointsToPath(node, false);
		} else if (tag === "polygon") {
			d = svgPointsToPath(node, true);
		}
		if (d) {
			pushSvgShapeElement(node, d, matrix, style, idMap, elements);
			if (tag === "path" || tag === "line" || tag === "polyline" || tag === "polygon") {
				svgRenderMarkers(node, d, matrix, style, idMap, elements, depth);
			}
		}
	}

	function svgRenderMarkers(node, d, matrix, style, idMap, container, depth) {
		const mEnd = node.getAttribute("marker-end");
		const mStart = node.getAttribute("marker-start");
		if (!mEnd && !mStart) {
			return;
		}
		const cmds = parseSvgPathData(d);
		const pts = [];
		let cx = 0, cy = 0;
		for (let i = 0; i < cmds.length; ++i) {
			const c = cmds[i];
			if (c.type === "moveto" || c.type === "lineto") { cx = c.x; cy = c.y; pts.push([cx, cy]); }
			else if (c.type === "cubicBezier") { pts.push([c.x2, c.y2]); cx = c.x; cy = c.y; pts.push([cx, cy]); }
			else if (c.type === "quadraticBezier") { pts.push([c.x1, c.y1]); cx = c.x2; cy = c.y2; pts.push([cx, cy]); }
		}
		if (pts.length < 1) {
			return;
		}
		const strokeW = style.strokeWidth || 1;
		function place(ref, p, dirAngle) {
			const idm = ("" + ref).match(/#([^)\s'"]+)/);
			if (!idm || !idMap[idm[1]]) {
				return;
			}
			const mk = idMap[idm[1]];
			if (!mk.tagName || mk.tagName.toLowerCase() !== "marker") {
				return;
			}
			const mw = parseFloat(mk.getAttribute("markerWidth"));
			const mh = parseFloat(mk.getAttribute("markerHeight"));
			const refX = parseFloat(mk.getAttribute("refX")) || 0;
			const refY = parseFloat(mk.getAttribute("refY")) || 0;
			const units = ("" + (mk.getAttribute("markerUnits") || "strokewidth")).toLowerCase();
			const unitScale = (units === "userspaceonuse") ? 1 : strokeW;
			let vsx = 1, vsy = 1;
			const vb = mk.getAttribute("viewBox");
			if (vb) {
				const p4 = vb.split(/[\s,]+/).map(parseFloat);
				if (p4.length >= 4 && p4[2] > 0 && p4[3] > 0) {
					vsx = (mw > 0 ? mw : p4[2]) / p4[2];
					vsy = (mh > 0 ? mh : p4[3]) / p4[3];
				}
			}
			const orient = ("" + (mk.getAttribute("orient") || "")).toLowerCase();
			let ang = 0;
			if (orient === "auto" || orient === "auto-start-reverse" || orient === "") {
				ang = dirAngle;
			} else {
				ang = (parseFloat(orient) || 0) * Math.PI / 180;
			}
			const cos = Math.cos(ang), sin = Math.sin(ang);
			let mm = matMulSvg(matrix, [1, 0, 0, 1, p[0], p[1]]);
			mm = matMulSvg(mm, [cos, sin, -sin, cos, 0, 0]);
			mm = matMulSvg(mm, [unitScale * vsx, 0, 0, unitScale * vsy, 0, 0]);
			mm = matMulSvg(mm, [1, 0, 0, 1, -refX, -refY]);
			const kids = svgElementChildren(mk);
			for (let i = 0; i < kids.length; ++i) {
				walkSvgNode(kids[i], mm, inheritSvgStyle(mk, style), idMap, container, depth + 1);
			}
		}
		if (mEnd && pts.length >= 1) {
			const e = pts[pts.length - 1];
			const pPrev = pts.length >= 2 ? pts[pts.length - 2] : e;
			place(mEnd, e, Math.atan2(e[1] - pPrev[1], e[0] - pPrev[0]));
		}
		if (mStart && pts.length >= 1) {
			const s0 = pts[0];
			const pNext = pts.length >= 2 ? pts[1] : s0;
			place(mStart, s0, Math.atan2(pNext[1] - s0[1], pNext[0] - s0[0]));
		}
	}

	function walkSvgNestedViewport(node, matrix, style, idMap, elements, depth) {
		const x = parseFloat(node.getAttribute("x")) || 0;
		const y = parseFloat(node.getAttribute("y")) || 0;
		let vm = matMulSvg(matrix, [1, 0, 0, 1, x, y]);
		const w = parseFloat(node.getAttribute("width"));
		const h = parseFloat(node.getAttribute("height"));
		const vb = node.getAttribute("viewBox");
		if (vb) {
			const p = vb.split(/[\s,]+/).map(parseFloat);
			if (p.length >= 4 && p[2] > 0 && p[3] > 0) {
				if (w > 0 && h > 0) {
					const par = ("" + (node.getAttribute("preserveAspectRatio") || "")).trim().toLowerCase();
					let sx = w / p[2], sy = h / p[3], tx = 0, ty = 0;
					if (par.indexOf("none") !== 0) {
						const s = par.indexOf("slice") !== -1 ? Math.max(sx, sy) : Math.min(sx, sy);
						sx = s; sy = s;
						const alignX = par.indexOf("xmax") !== -1 ? 1 : (par.indexOf("xmin") !== -1 ? 0 : 0.5);
						const alignY = par.indexOf("ymax") !== -1 ? 1 : (par.indexOf("ymin") !== -1 ? 0 : 0.5);
						tx = (w - p[2] * s) * alignX;
						ty = (h - p[3] * s) * alignY;
					}
					vm = matMulSvg(vm, [sx, 0, 0, sy, tx - p[0] * sx, ty - p[1] * sy]);
				} else {
					vm = matMulSvg(vm, [1, 0, 0, 1, -p[0], -p[1]]);
				}
			}
		}
		const kids = svgElementChildren(node);
		for (let i = 0; i < kids.length; ++i) {
			walkSvgNode(kids[i], vm, style, idMap, elements, depth + 1);
		}
	}

	function walkSvgUse(node, matrix, style, idMap, elements, depth) {
		let href = (node.getAttribute("href") || node.getAttribute("xlink:href")) || null;
		if (!href && node.getAttributeNS) {
			href = node.getAttributeNS("http://www.w3.org/1999/xlink", "href");
		}
		if (!href) {
			return;
		}
		const idm = href.match(/#(.+)$/);
		if (!idm || !idMap[idm[1]]) {
			return;
		}
		const ref = idMap[idm[1]];
		const ux = parseFloat(node.getAttribute("x")) || 0;
		const uy = parseFloat(node.getAttribute("y")) || 0;
		let useMatrix = matMulSvg(matrix, [1, 0, 0, 1, ux, uy]);
		const refTag = ref.tagName ? ref.tagName.toLowerCase() : "";
		if (refTag === "symbol" || refTag === "svg") {
			const vb = ref.getAttribute("viewBox");
			const uw = parseFloat(node.getAttribute("width"));
			const uh = parseFloat(node.getAttribute("height"));
			if (vb) {
				const p = vb.split(/[\s,]+/).map(parseFloat);
				if (p.length >= 4 && p[2] > 0 && p[3] > 0) {
					if (uw > 0 && uh > 0) {
						useMatrix = matMulSvg(useMatrix, [uw / p[2], 0, 0, uh / p[3], -p[0] * (uw / p[2]), -p[1] * (uh / p[3])]);
					} else {
						useMatrix = matMulSvg(useMatrix, [1, 0, 0, 1, -p[0], -p[1]]);
					}
				}
			}
			const symStyle = inheritSvgStyle(ref, style);
			const kids = svgElementChildren(ref);
			for (let i = 0; i < kids.length; ++i) {
				walkSvgNode(kids[i], useMatrix, symStyle, idMap, elements, depth + 1);
			}
		} else {
			walkSvgNode(ref, useMatrix, style, idMap, elements, depth + 1);
		}
	}

	function parseSvgString(svgString) {
		if (typeof DOMParser === "undefined") {
			return null;
		}
		const doc = new DOMParser().parseFromString(svgString, "image/svg+xml");
		const svg = doc.querySelector("svg");
		if (!svg) {
			return null;
		}
		let vbMinX = 0, vbMinY = 0;
		let width = parseFloat(svg.getAttribute("width"));
		let height = parseFloat(svg.getAttribute("height"));
		const viewBox = svg.getAttribute("viewBox");
		if (viewBox) {
			const vb = viewBox.split(/[\s,]+/).map(parseFloat);
			if (vb.length >= 4 && vb[2] > 0 && vb[3] > 0) {
				vbMinX = vb[0]; vbMinY = vb[1]; width = vb[2]; height = vb[3];
			}
		}
		if (!(width > 0) || !(height > 0)) {
			return null;
		}

		const idMap = {};
		const withId = svg.querySelectorAll("[id]");
		for (let i = 0; i < withId.length; ++i) {
			const id = withId[i].getAttribute("id");
			if (id && !idMap[id]) {
				idMap[id] = withId[i];
			}
		}

		svgCssRules = [];
		const styleNodes = svg.getElementsByTagName("style");
		for (let i = 0; i < styleNodes.length; ++i) {
			parseSvgCss(styleNodes[i].textContent, svgCssRules);
		}

		const elements = [];
		const rootStyle = {
			fill: "#000000", stroke: null, strokeWidth: 1, color: svg.getAttribute("color") || "#000000",
			opacity: 1, fillOpacity: 1, strokeOpacity: 1, fillRule: "nonzero", display: true
		};
		walkSvgNode(svg, [1, 0, 0, 1, -vbMinX, -vbMinY], rootStyle, idMap, elements, 0);
		if (elements.length === 0) {
			return null;
		}
		return {width: width, height: height, elements: elements};
	}

	function getSvgCommandsBounds(commands) {
		let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
		let curX = 0, curY = 0, has = false;

		function point(x, y) {
			if (x < minX) minX = x;
			if (y < minY) minY = y;
			if (x > maxX) maxX = x;
			if (y > maxY) maxY = y;
			has = true;
		}

		function addBezierAxis(v0, v1, v2, v3, isX) {
			const a = -v0 + 3 * v1 - 3 * v2 + v3;
			const b = 2 * (v0 - 2 * v1 + v2);
			const c = -v0 + v1;
			const roots = [];
			if (Math.abs(a) < 1e-9) {
				if (Math.abs(b) > 1e-9) {
					roots.push(-c / b);
				}
			} else {
				const disc = b * b - 4 * a * c;
				if (disc >= 0) {
					const sq = Math.sqrt(disc);
					roots.push((-b + sq) / (2 * a));
					roots.push((-b - sq) / (2 * a));
				}
			}
			for (let i = 0; i < roots.length; ++i) {
				const t = roots[i];
				if (t > 0 && t < 1) {
					const mt = 1 - t;
					const val = mt * mt * mt * v0 + 3 * mt * mt * t * v1 + 3 * mt * t * t * v2 + t * t * t * v3;
					if (isX) {
						if (val < minX) minX = val;
						if (val > maxX) maxX = val;
					} else {
						if (val < minY) minY = val;
						if (val > maxY) maxY = val;
					}
				}
			}
		}

		for (let i = 0; i < commands.length; ++i) {
			const cmd = commands[i];
			if (cmd.type === "moveto" || cmd.type === "lineto") {
				point(cmd.x, cmd.y);
				curX = cmd.x; curY = cmd.y;
			} else if (cmd.type === "cubicBezier") {
				point(curX, curY);
				point(cmd.x, cmd.y);
				addBezierAxis(curX, cmd.x1, cmd.x2, cmd.x, true);
				addBezierAxis(curY, cmd.y1, cmd.y2, cmd.y, false);
				curX = cmd.x; curY = cmd.y;
			} else if (cmd.type === "quadraticBezier") {
				point(curX, curY);
				point(cmd.x2, cmd.y2);
				addBezierAxis(curX, (curX + 2 * cmd.x1) / 3, (cmd.x2 + 2 * cmd.x1) / 3, cmd.x2, true);
				addBezierAxis(curY, (curY + 2 * cmd.y1) / 3, (cmd.y2 + 2 * cmd.y1) / 3, cmd.y2, false);
				curX = cmd.x2; curY = cmd.y2;
			}
		}
		if (!has) {
			return null;
		}
		return {minX: minX, minY: minY, w: maxX - minX, h: maxY - minY};
	}

	function buildSvgGeometry(commands, bounds, scaleX, scaleY, hasStroke) {
		const pathW = Math.max(1, Math.round(bounds.w * scaleX));
		const pathH = Math.max(1, Math.round(bounds.h * scaleY));
		const geometry = new AscFormat.Geometry();
		const path = new AscFormat.Path();
		path.setStroke(hasStroke === true);
		path.setFill("norm");
		path.setExtrusionOk(false);
		path.setPathW(pathW);
		path.setPathH(pathH);

		function sx(v) { return Math.round((v - bounds.minX) * scaleX); }
		function sy(v) { return Math.round((v - bounds.minY) * scaleY); }

		for (let i = 0; i < commands.length; ++i) {
			const cmd = commands[i];
			if (cmd.type === "moveto") {
				path.moveTo(sx(cmd.x), sy(cmd.y));
			} else if (cmd.type === "lineto") {
				path.lnTo(sx(cmd.x), sy(cmd.y));
			} else if (cmd.type === "cubicBezier") {
				path.cubicBezTo(sx(cmd.x1), sy(cmd.y1), sx(cmd.x2), sy(cmd.y2), sx(cmd.x), sy(cmd.y));
			} else if (cmd.type === "quadraticBezier") {
				path.quadBezTo(sx(cmd.x1), sy(cmd.y1), sx(cmd.x2), sy(cmd.y2));
			} else if (cmd.type === "closepath") {
				path.close();
			}
		}
		geometry.AddPath(path);
		return geometry;
	}

	function createSvgShape(isPdf, isWord, worksheet) {
		const shape = isPdf ? new AscPDF.CPdfShape() : new AscFormat.CShape();
		shape.setWordShape(isWord === true);
		shape.setBDeleted(false);
		if (worksheet) {
			shape.setWorksheet(worksheet);
		}
		return shape;
	}

	function createSvgTextRunPr(run, mmPerSvgY) {
		let ptSize = Math.round(run.fontSize * mmPerSvgY * 72 / 25.4 * 2) / 2;
		if (ptSize < 1) {
			ptSize = 1;
		} else if (ptSize > 300) {
			ptSize = 300;
		}
		const textPr = new CTextPr();
		textPr.FontSize = ptSize;
		textPr.FontSizeCS = ptSize;
		textPr.RFonts.Ascii = {Name: run.fontFamily, Index: -1};
		textPr.RFonts.HAnsi = {Name: run.fontFamily, Index: -1};
		textPr.RFonts.CS = {Name: run.fontFamily, Index: -1};
		textPr.RFonts.EastAsia = {Name: run.fontFamily, Index: -1};
		textPr.Bold = !!run.bold;
		textPr.Italic = !!run.italic;
		textPr.Unifill = AscFormat.CreateSolidFillRGB(run.color[0], run.color[1], run.color[2]);
		return textPr;
	}

	function createSvgTextShape(el, opts, mmPerSvgX, mmPerSvgY) {
		const shape = createSvgShape(opts.isPdf, opts.isWord, opts.worksheet);
		if (opts.isWord) {
			shape.createTextBoxContent();
		} else {
			shape.createTextBodyOnEdit();
		}

		const runs = el.runs;
		const ascent = el.fontSize * (SVG_TEXT_LINE_RATIO - SVG_TEXT_DESCENT_RATIO);
		let estWidthSvg = 0;
		for (let i = 0; i < runs.length; ++i) {
			estWidthSvg += runs[i].text.length * runs[i].fontSize * 0.6;
		}
		estWidthSvg = Math.max(1, estWidthSvg);
		let leftSvg = el.x;
		if (el.anchor === "middle") {
			leftSvg = el.x - estWidthSvg / 2;
		} else if (el.anchor === "end") {
			leftSvg = el.x - estWidthSvg;
		}

		const spPr = new AscFormat.CSpPr();
		const xfrm = new AscFormat.CXfrm();
		xfrm.setOffX(opts.x + leftSvg * mmPerSvgX);
		xfrm.setOffY(opts.y + (el.y - ascent) * mmPerSvgY);
		xfrm.setExtX(estWidthSvg * mmPerSvgX);
		xfrm.setExtY(el.fontSize * SVG_TEXT_LINE_RATIO * mmPerSvgY);
		if (el.rot) {
			xfrm.setRot(el.rot);
		}
		spPr.setXfrm(xfrm);
		xfrm.setParent(spPr);
		spPr.setFill(AscFormat.CreateNoFillUniFill());
		spPr.setLn(AscFormat.CreateNoFillLine());
		spPr.setGeometry(AscFormat.CreateGeometry("rect"));
		shape.setSpPr(spPr);
		spPr.setParent(shape);

		const content = shape.getDocContent();
		AscFormat.AddToContentFromString(content, runs[0].text);
		content.SetApplyToAll(true);
		content.AddToParagraph(new ParaTextPr(createSvgTextRunPr(runs[0], mmPerSvgY)));
		content.SetParagraphAlign(el.anchor === "middle" ? AscCommon.align_Center : (el.anchor === "end" ? AscCommon.align_Right : AscCommon.align_Left));
		content.SetParagraphSpacing({Line: 1, LineRule: Asc.linerule_Auto, Before: 0, After: 0});
		content.SetApplyToAll(false);
		const paragraph = content.Content[0];
		for (let i = 1; i < runs.length; ++i) {
			const paraRun = new ParaRun(paragraph, false);
			paraRun.SetPr(createSvgTextRunPr(runs[i], mmPerSvgY));
			paraRun.AddText(runs[i].text);
			paragraph.AddToContent(paragraph.Content.length - 1, paraRun);
		}

		const bodyPr = shape.getBodyPr().createDuplicate();
		bodyPr.rot = 0;
		bodyPr.wrap = AscFormat.nTWTNone;
		bodyPr.setDefaultInsets();
		bodyPr.lIns = 0;
		bodyPr.tIns = 0;
		bodyPr.rIns = 0;
		bodyPr.bIns = 0;
		bodyPr.anchor = 0;
		bodyPr.textFit = new AscFormat.CTextFit();
		bodyPr.textFit.type = AscFormat.text_fit_Auto;
		if (opts.isWord) {
			shape.setBodyPr(bodyPr);
		} else {
			shape.txBody.setBodyPr(bodyPr);
		}

		return shape;
	}

	function svgCommandsToCompoundPath(commands) {
		const cp = new AscCommon["PathBoolean"]["CompoundPath"]();
		let curX = 0, curY = 0;
		for (let i = 0; i < commands.length; ++i) {
			const c = commands[i];
			if (c.type === "moveto") {
				cp["moveTo"](c.x, c.y); curX = c.x; curY = c.y;
			} else if (c.type === "lineto") {
				cp["lineTo"](c.x, c.y); curX = c.x; curY = c.y;
			} else if (c.type === "cubicBezier") {
				cp["cubicCurveTo"](c.x1, c.y1, c.x2, c.y2, c.x, c.y); curX = c.x; curY = c.y;
			} else if (c.type === "quadraticBezier") {
				const c1x = curX + 2 / 3 * (c.x1 - curX), c1y = curY + 2 / 3 * (c.y1 - curY);
				const c2x = c.x2 + 2 / 3 * (c.x1 - c.x2), c2y = c.y2 + 2 / 3 * (c.y1 - c.y2);
				cp["cubicCurveTo"](c1x, c1y, c2x, c2y, c.x2, c.y2); curX = c.x2; curY = c.y2;
			} else if (c.type === "closepath") {
				cp["closePath"]();
			}
		}
		return cp;
	}

	function svgCompoundPathToCommands(result) {
		const out = [];
		const children = (result instanceof AscCommon["PathBoolean"]["CompoundPath"]) ? result["getChildren"]() : [result];
		for (let k = 0; k < children.length; ++k) {
			const path = children[k];
			const segs = path["getSegments"]();
			if (!segs.length) {
				continue;
			}
			for (let i = 0; i < segs.length; ++i) {
				const seg = segs[i];
				const pt = seg["getPoint"]();
				if (i === 0) {
					out.push({type: "moveto", x: pt["getX"](), y: pt["getY"]()});
					continue;
				}
				const prev = seg["getPrevious"]();
				const pp = prev["getPoint"](), pho = prev["getHandleOut"](), hi = seg["getHandleIn"]();
				out.push({type: "cubicBezier", x1: pp["getX"]() + pho["getX"](), y1: pp["getY"]() + pho["getY"](), x2: pt["getX"]() + hi["getX"](), y2: pt["getY"]() + hi["getY"](), x: pt["getX"](), y: pt["getY"]()});
			}
			if (path["isClosed"]()) {
				const last = segs[segs.length - 1], first = segs[0];
				const lp = last["getPoint"](), lho = last["getHandleOut"](), fhi = first["getHandleIn"](), fp = first["getPoint"]();
				out.push({type: "cubicBezier", x1: lp["getX"]() + lho["getX"](), y1: lp["getY"]() + lho["getY"](), x2: fp["getX"]() + fhi["getX"](), y2: fp["getY"]() + fhi["getY"](), x: fp["getX"](), y: fp["getY"]()});
				out.push({type: "closepath"});
			}
		}
		return out;
	}

	function svgIntersectCommands(commands, clipCommands) {
		if (!AscCommon["PathBoolean"]) {
			return commands;
		}
		try {
			const a = svgCommandsToCompoundPath(commands);
			const b = svgCommandsToCompoundPath(clipCommands);
			const res = a["intersect"](b);
			if (!res) {
				return [];
			}
			const out = svgCompoundPathToCommands(res);
			return out;
		} catch (e) {
			return commands;
		}
	}

	function svgFlattenSubpaths(commands) {
		const subs = [];
		let cur = null, cx = 0, cy = 0, sx = 0, sy = 0;
		function sampleCubic(p0x, p0y, c1x, c1y, c2x, c2y, p3x, p3y) {
			const N = 8;
			for (let i = 1; i <= N; ++i) {
				const t = i / N, mt = 1 - t;
				cur.push([mt * mt * mt * p0x + 3 * mt * mt * t * c1x + 3 * mt * t * t * c2x + t * t * t * p3x,
					mt * mt * mt * p0y + 3 * mt * mt * t * c1y + 3 * mt * t * t * c2y + t * t * t * p3y]);
			}
		}
		for (let i = 0; i < commands.length; ++i) {
			const c = commands[i];
			if (c.type === "moveto") {
				cur = [[c.x, c.y]];
				subs.push(cur);
				cx = c.x; cy = c.y; sx = c.x; sy = c.y;
			} else if (c.type === "lineto") {
				if (cur) { cur.push([c.x, c.y]); }
				cx = c.x; cy = c.y;
			} else if (c.type === "cubicBezier") {
				if (cur) { sampleCubic(cx, cy, c.x1, c.y1, c.x2, c.y2, c.x, c.y); }
				cx = c.x; cy = c.y;
			} else if (c.type === "quadraticBezier") {
				if (cur) { sampleCubic(cx, cy, (cx + 2 * c.x1) / 3, (cy + 2 * c.y1) / 3, (c.x2 + 2 * c.x1) / 3, (c.y2 + 2 * c.y1) / 3, c.x2, c.y2); }
				cx = c.x2; cy = c.y2;
			} else if (c.type === "closepath") {
				if (cur) { cur.push([sx, sy]); }
				cx = sx; cy = sy;
			}
		}
		return subs;
	}

	function svgStrokeToOutlineCommands(commands, w) {
		const hw = (w > 0 ? w : 1) / 2;
		const subs = svgFlattenSubpaths(commands);
		const out = [];
		for (let s = 0; s < subs.length; ++s) {
			const pts = subs[s];
			for (let i = 0; i + 1 < pts.length; ++i) {
				const a = pts[i], b = pts[i + 1];
				let dx = b[0] - a[0], dy = b[1] - a[1];
				const len = Math.sqrt(dx * dx + dy * dy);
				if (len < 1e-6) {
					continue;
				}
				dx /= len; dy /= len;
				const px = -dy * hw, py = dx * hw;
				out.push({type: "moveto", x: a[0] + px, y: a[1] + py});
				out.push({type: "lineto", x: b[0] + px, y: b[1] + py});
				out.push({type: "lineto", x: b[0] - px, y: b[1] - py});
				out.push({type: "lineto", x: a[0] - px, y: a[1] - py});
				out.push({type: "closepath"});
			}
		}
		return out;
	}

	function buildSvgGroupShape(shapes) {
		let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
		for (let i = 0; i < shapes.length; ++i) {
			const xfrm = shapes[i].spPr.xfrm;
			minX = Math.min(minX, xfrm.offX);
			minY = Math.min(minY, xfrm.offY);
			maxX = Math.max(maxX, xfrm.offX + xfrm.extX);
			maxY = Math.max(maxY, xfrm.offY + xfrm.extY);
		}
		const extX = maxX - minX;
		const extY = maxY - minY;
		const group = new AscFormat.CGroupShape();
		group.setBDeleted(false);
		group.setSpPr(new AscFormat.CSpPr());
		group.spPr.setParent(group);
		group.spPr.setXfrm(new AscFormat.CXfrm());
		const gx = group.spPr.xfrm;
		gx.setParent(group.spPr);
		gx.setOffX(minX);
		gx.setOffY(minY);
		gx.setExtX(extX);
		gx.setExtY(extY);
		gx.setChOffX(0);
		gx.setChOffY(0);
		gx.setChExtX(extX);
		gx.setChExtY(extY);
		for (let i = 0; i < shapes.length; ++i) {
			const sp = shapes[i];
			sp.spPr.xfrm.setOffX(sp.spPr.xfrm.offX - minX);
			sp.spPr.xfrm.setOffY(sp.spPr.xfrm.offY - minY);
			sp.setGroup(group);
			group.addToSpTree(group.spTree.length, sp);
		}
		return group;
	}

	function buildSvgLeaf(element, opts, mmPerSvgX, mmPerSvgY, scaleX, scaleY) {
		if (element.isText) {
			const textShape = createSvgTextShape(element, opts, mmPerSvgX, mmPerSvgY);
			return textShape ? [textShape] : null;
		}
		let isClosed = false;
		for (let i = 0; i < element.commands.length; ++i) {
			if (element.commands[i].type === "closepath") { isClosed = true; break; }
		}
		const hasFill = !!(element.fill || element.fillGrad || element.fillPattern);

		let fillCommands = null;
		let strokeOutline = null;
		let drawStroke = !!element.stroke;
		if (element.clip) {
			drawStroke = false;
			if (hasFill && isClosed) {
				const clipped = svgIntersectCommands(element.commands, element.clip);
				if (clipped && clipped.length > 0) {
					fillCommands = clipped;
				}
			}
			if (element.stroke) {
				const outline = svgStrokeToOutlineCommands(element.commands, element.strokeWidth || 1);
				const clipped = svgIntersectCommands(outline, element.clip);
				if (clipped && clipped.length > 0) {
					strokeOutline = clipped;
				}
			}
		} else {
			fillCommands = element.commands;
		}

		function emit(commands, strokeFill, withStroke) {
			const bounds = getSvgCommandsBounds(commands);
			if (!bounds) {
				return null;
			}
			const hasArea = bounds.w > 0 && bounds.h > 0;
			if (!hasArea && !withStroke) {
				return null;
			}
			const shape = createSvgShape(opts.isPdf, opts.isWord, opts.worksheet);
			const spPr = new AscFormat.CSpPr();
			const xfrm = new AscFormat.CXfrm();
			xfrm.setOffX(opts.x + bounds.minX * mmPerSvgX);
			xfrm.setOffY(opts.y + bounds.minY * mmPerSvgY);
			xfrm.setExtX(Math.max(bounds.w * mmPerSvgX, 0.01));
			xfrm.setExtY(Math.max(bounds.h * mmPerSvgY, 0.01));
			spPr.setXfrm(xfrm);
			xfrm.setParent(spPr);

			if (strokeFill) {
				spPr.setFill(svgSolidFill(strokeFill.rgb, strokeFill.alpha));
			} else if (element.fillGrad) {
				spPr.setFill(buildSvgGradFill(element.fillGrad));
			} else if (element.fillPattern) {
				spPr.setFill(svgPatternFill(element.fillPattern.rgb));
			} else if (element.fill) {
				spPr.setFill(svgSolidFill(element.fill, element.fillAlpha));
			} else {
				spPr.setFill(AscFormat.CreateNoFillUniFill());
			}
			if (withStroke) {
				const ln = new AscFormat.CLn();
				ln.setFill(svgSolidFill(element.stroke, element.strokeAlpha));
				const wEmu = Math.round(element.strokeWidth * (scaleX + scaleY) / 2);
				ln.setW(wEmu > 0 ? wEmu : 6350);
				if (element.strokeDash) {
					ln.setPrstDash(ln.GetDashCode(element.strokeDash));
				}
				spPr.setLn(ln);
			} else {
				spPr.setLn(AscFormat.CreateNoFillLine());
			}

			spPr.setGeometry(buildSvgGeometry(commands, bounds, scaleX, scaleY, withStroke));
			shape.setSpPr(spPr);
			spPr.setParent(shape);
			return shape;
		}

		const out = [];
		if (fillCommands) {
			const fillShape = emit(fillCommands, null, drawStroke);
			if (fillShape) {
				out.push(fillShape);
			}
		}
		if (strokeOutline) {
			const strokeShape = emit(strokeOutline, {rgb: element.stroke, alpha: element.strokeAlpha}, false);
			if (strokeShape) {
				out.push(strokeShape);
			}
		}
		return out.length > 0 ? out : null;
	}

	function buildSvgTree(nodes, out, opts, mmPerSvgX, mmPerSvgY, scaleX, scaleY) {
		for (let i = 0; i < nodes.length; ++i) {
			const node = nodes[i];
			if (node.isGroupNode) {
				if (opts.isPdf) {
					buildSvgTree(node.children, out, opts, mmPerSvgX, mmPerSvgY, scaleX, scaleY);
				} else {
					const childDrawings = [];
					buildSvgTree(node.children, childDrawings, opts, mmPerSvgX, mmPerSvgY, scaleX, scaleY);
					if (childDrawings.length === 1) {
						out.push(childDrawings[0]);
					} else if (childDrawings.length > 1) {
						out.push(buildSvgGroupShape(childDrawings));
					}
				}
			} else {
				const drawings = buildSvgLeaf(node, opts, mmPerSvgX, mmPerSvgY, scaleX, scaleY);
				if (drawings) {
					for (let j = 0; j < drawings.length; ++j) {
						out.push(drawings[j]);
					}
				}
			}
		}
	}

	function createShapesFromSvg(svgString, opts) {
		try {
			const parsed = parseSvgString(svgString);
			if (!parsed) {
				return [];
			}
			const mmPerSvgX = opts.extX / parsed.width;
			const mmPerSvgY = opts.extY / parsed.height;
			const scaleX = mmPerSvgX * SVG_EMU_PER_MM;
			const scaleY = mmPerSvgY * SVG_EMU_PER_MM;
			const top = [];
			buildSvgTree(parsed.elements, top, opts, mmPerSvgX, mmPerSvgY, scaleX, scaleY);
			for (let i = 0; i < top.length; ++i) {
				top[i].checkDrawingUniNvPr();
			}
			return top;
		} catch (e) {
			return [];
		}
	}

	function isSvgImage(rasterImageId) {
		if (typeof rasterImageId !== "string" || rasterImageId.length === 0) {
			return false;
		}
		if (rasterImageId.indexOf("data:image/svg") === 0) {
			return true;
		}
		if (AscCommon.GetFileExtension(rasterImageId) === "svg") {
			return true;
		}
		const local = AscCommon.g_oDocumentUrls.getImageLocal(rasterImageId);
		return !!(local && AscCommon.GetFileExtension(local) === "svg");
	}

	function decodeSvgDataUri(uri) {
		const comma = uri.indexOf(",");
		if (comma === -1) {
			return null;
		}
		const meta = uri.substring(0, comma);
		const data = uri.substring(comma + 1);
		try {
			if (meta.indexOf("base64") !== -1) {
				return decodeURIComponent(escape(window.atob(data)));
			}
			return decodeURIComponent(data);
		} catch (e) {
			return null;
		}
	}

	function loadSvgText(rasterImageId, callback) {
		if (rasterImageId.indexOf("data:") === 0) {
			callback(decodeSvgDataUri(rasterImageId));
			return;
		}
		try {
			const xhr = new XMLHttpRequest();
			xhr.open("GET", AscCommon.getFullImageSrc2(rasterImageId), true);
			xhr.responseType = "arraybuffer";
			xhr.onload = function () {
				if ((xhr.status === 200 || xhr.status === 0) && xhr.response) {
					const bytes = new Uint8Array(xhr.response);
					callback(AscCommon.UTF8ArrayToString(bytes, 0, bytes.length));
				} else {
					callback(null);
				}
			};
			xhr.onerror = function () {
				callback(null);
			};
			xhr.send(null);
		} catch (e) {
			callback(null);
		}
	}

	window["AscFormat"] = window["AscFormat"] || {};
	window["AscFormat"].CreateShapesFromSvg = createShapesFromSvg;
	window["AscFormat"].IsSvgImage = isSvgImage;
	window["AscFormat"].LoadSvgText = loadSvgText;
})(window);

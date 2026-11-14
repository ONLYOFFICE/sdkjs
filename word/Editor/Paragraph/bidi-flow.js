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

(function(window)
{
	// UAX #9 kinds. Bits so weak/neutral rules can test groups with a mask.
	const K = {
		L   : 1 << 0,
		R   : 1 << 1,
		AL  : 1 << 2,
		EN  : 1 << 3,
		ES  : 1 << 4,
		ET  : 1 << 5,
		AN  : 1 << 6,
		CS  : 1 << 7,
		NSM : 1 << 8,
		BN  : 1 << 9,
		B   : 1 << 10,
		S   : 1 << 11,
		WS  : 1 << 12,
		ON  : 1 << 13,
		LRE : 1 << 14,
		LRO : 1 << 15,
		RLE : 1 << 16,
		RLO : 1 << 17,
		PDF : 1 << 18,
		LRI : 1 << 19,
		RLI : 1 << 20,
		FSI : 1 << 21,
		PDI : 1 << 22
	};
	
	const STRONG          = K.L | K.R | K.AL;
	const BN_LIKE         = K.BN | K.LRE | K.LRO | K.RLE | K.RLO | K.PDF;
	const ISOLATE_INIT    = K.LRI | K.RLI | K.FSI;
	const NI              = K.B | K.S | K.WS | K.ON | K.FSI | K.LRI | K.RLI | K.PDI;
	const TRAILING        = K.S | K.WS | K.B | ISOLATE_INIT | K.PDI | BN_LIKE;
	const FORMATTING      = K.RLE | K.LRE | K.RLO | K.LRO | ISOLATE_INIT | K.PDI | K.PDF | K.B;
	const R_FOR_NEUTRALS  = K.R | K.EN | K.AN;
	const STRONG_FOR_N    = R_FOR_NEUTRALS | K.L;
	const MAX_DEPTH       = 125;
	
	// Opening paired brackets from Unicode Bidi_Paired_Bracket (class "o").
	// U+003C/U+003E and guillemets are mirrored glyphs here, but they are not bidi brackets.
	const OPENING_BRACKET = {};
	const OPENING_LIST = [
		0x0028, 0x005B, 0x007B, 0x0F3A, 0x0F3C, 0x169B, 0x2045, 0x207D, 0x208D, 0x2308, 0x230A, 0x2329,
		0x2768, 0x276A, 0x276C, 0x276E, 0x2770, 0x2772, 0x2774, 0x27C5, 0x27E6, 0x27E8, 0x27EA, 0x27EC,
		0x27EE, 0x2983, 0x2985, 0x2987, 0x2989, 0x298B, 0x298D, 0x298F, 0x2991, 0x2993, 0x2995, 0x2997,
		0x29D8, 0x29DA, 0x29FC, 0x2E22, 0x2E24, 0x2E26, 0x2E28, 0x2E55, 0x2E57, 0x2E59, 0x2E5B, 0x3008,
		0x300A, 0x300C, 0x300E, 0x3010, 0x3014, 0x3016, 0x3018, 0x301A, 0xFE59, 0xFE5B, 0xFE5D, 0xFF08,
		0xFF3B, 0xFF5B, 0xFF5F, 0xFF62
	];
	for (let i = 0; i < OPENING_LIST.length; i++)
		OPENING_BRACKET[OPENING_LIST[i]] = true;
	
	function kindFromType(type)
	{
		const T = AscBidi.TYPE;
		if (type === T.L) return K.L;
		if (type === T.R) return K.R;
		if (type === T.AL) return K.AL;
		if (type === T.EN) return K.EN;
		if (type === T.ES) return K.ES;
		if (type === T.ET) return K.ET;
		if (type === T.AN) return K.AN;
		if (type === T.CS) return K.CS;
		if (type === T.NSM) return K.NSM;
		if (type === T.BN) return K.BN;
		if (type === T.B) return K.B;
		if (type === T.S) return K.S;
		if (type === T.WS) return K.WS;
		if (type === T.ON) return K.ON;
		if (type === T.LRE) return K.LRE;
		if (type === T.LRO) return K.LRO;
		if (type === T.RLE) return K.RLE;
		if (type === T.RLO) return K.RLO;
		if (type === T.PDF) return K.PDF;
		if (type === T.LRI) return K.LRI;
		if (type === T.RLI) return K.RLI;
		if (type === T.FSI) return K.FSI;
		if (type === T.PDI) return K.PDI;
		return K.ON;
	}
	
	function codePointOf(element)
	{
		if (typeof element === "string" && element.length)
			return element.codePointAt(0);
		if (typeof element === "number")
			return element;
		if (Array.isArray(element))
		{
			// Numbering draws [codePoint, fontCoeff]. Paragraph runs are [para_*, runElement, ...].
			if (typeof element[0] === "number" && typeof element[1] === "number")
				return element[0];
			if (element[1] && typeof element[1].Value === "number")
				return element[1].Value;
		}
		if (element && typeof element === "object")
		{
			if (typeof element.char === "number")
				return element.char;
			if (typeof element.char === "string" && element.char.length)
				return element.char.codePointAt(0);
			if (typeof element.cp === "number")
				return element.cp;
			if (typeof element.Value === "number")
				return element.Value;
		}
		return -1;
	}
	
	function nextEven(level)
	{
		return level + ((level & 1) ? 1 : 2);
	}
	function nextOdd(level)
	{
		return level + ((level & 1) ? 2 : 1);
	}
	function pairedBracket(cp)
	{
		if (cp < 0 || !AscBidi.getPairedBracket)
			return -1;
		let pair = AscBidi.getPairedBracket(cp);
		return (typeof pair === "number" && pair > 0) ? pair : -1;
	}
	function canonicalBracket(cp)
	{
		// BD16 canonical equivalents of the angle brackets.
		if (cp === 0x2329) return 0x3008;
		if (cp === 0x232A) return 0x3009;
		return cp;
	}
	function isUnicodeBracket(cp)
	{
		// Less-than, greater-than and guillemets are mirrored by the renderer, not by N0.
		if (cp === 0x003C || cp === 0x003E || cp === 0x00AB || cp === 0x00BB)
			return false;
		return pairedBracket(cp) !== -1;
	}
	function isOpeningBracket(cp)
	{
		return isUnicodeBracket(cp) && !!OPENING_BRACKET[cp];
	}
	function isClosingBracket(cp)
	{
		return isUnicodeBracket(cp) && !OPENING_BRACKET[cp];
	}
	function bracketsMatch(openingCp, closingCp)
	{
		if (openingCp === pairedBracket(closingCp))
			return true;
		if (openingCp === pairedBracket(canonicalBracket(closingCp)))
			return true;
		return pairedBracket(canonicalBracket(openingCp)) === closingCp;
	}
	
	/**
	 * Resolve UAX #9 embedding levels for one paragraph.
	 * European numbers that follow a strong L (rule W7) stay in that LTR run, so
	 * an RTL paragraph "Abc def!23" is displayed as "Abc def!23" rather than "23!Abc def".
	 * @param {number[]} originalKinds
	 * @param {number[]} codePoints
	 * @param {number} paragraphLevel 0 = LTR, 1 = RTL
	 * @return {number[]}
	 */
	function resolveLevels(originalKinds, codePoints, paragraphLevel)
	{
		const n = originalKinds.length;
		const kinds = originalKinds.slice();
		const levels = new Array(n);
		const isolationPairs = {};
		
		function autoEmbed(start, isFSI)
		{
			let depth = 0;
			for (let i = start; i < n; i++)
			{
				const kind = originalKinds[i];
				if (!depth && (kind & (K.R | K.AL)))
					return 1;
				if (!depth && (kind & (K.B | K.L) || (isFSI && kind === K.PDI)))
					return 0;
				if (kind & ISOLATE_INIT)
					depth++;
				else if (kind === K.PDI && depth)
					depth--;
			}
			return 0;
		}
		
		const status = [{level: paragraphLevel, override: 0, isolate: false, init: -1}];
		let overflowIsolate = 0;
		let overflowEmbedding = 0;
		let validIsolate = 0;
		
		for (let i = 0; i < n; i++)
		{
			let kind = kinds[i];
			let top = status[status.length - 1];
			
			if (kind & FORMATTING)
			{
				if (kind & (K.RLE | K.LRE))
				{
					levels[i] = top.level;
					const level = (kind === K.RLE ? nextOdd : nextEven)(top.level);
					if (level <= MAX_DEPTH && !overflowIsolate && !overflowEmbedding)
						status.push({level: level, override: 0, isolate: false, init: -1});
					else if (!overflowIsolate)
						overflowEmbedding++;
				}
				else if (kind & (K.RLO | K.LRO))
				{
					levels[i] = top.level;
					const level = (kind === K.RLO ? nextOdd : nextEven)(top.level);
					if (level <= MAX_DEPTH && !overflowIsolate && !overflowEmbedding)
						status.push({level: level, override: (kind & K.RLO) ? K.R : K.L, isolate: false, init: -1});
					else if (!overflowIsolate)
						overflowEmbedding++;
				}
				else if (kind & ISOLATE_INIT)
				{
					if (kind & K.FSI)
						kind = autoEmbed(i + 1, true) === 1 ? K.RLI : K.LRI;
					levels[i] = top.level;
					if (top.override)
						kinds[i] = top.override;
					const level = (kind === K.RLI ? nextOdd : nextEven)(top.level);
					if (level <= MAX_DEPTH && !overflowIsolate && !overflowEmbedding)
					{
						validIsolate++;
						status.push({level: level, override: 0, isolate: true, init: i});
					}
					else
						overflowIsolate++;
				}
				else if (kind & K.PDI)
				{
					if (overflowIsolate > 0)
						overflowIsolate--;
					else if (validIsolate > 0)
					{
						overflowEmbedding = 0;
						while (status.length > 1 && !status[status.length - 1].isolate)
							status.pop();
						const init = status[status.length - 1].init;
						if (init >= 0)
						{
							isolationPairs[init] = i;
							isolationPairs[i] = init;
						}
						status.pop();
						validIsolate--;
					}
					top = status[status.length - 1];
					levels[i] = top.level;
					if (top.override)
						kinds[i] = top.override;
				}
				else if (kind & K.PDF)
				{
					if (!overflowIsolate)
					{
						if (overflowEmbedding > 0)
							overflowEmbedding--;
						else if (!top.isolate && status.length > 1)
						{
							status.pop();
							top = status[status.length - 1];
						}
					}
					levels[i] = top.level;
				}
				else if (kind & K.B)
				{
					levels[i] = paragraphLevel;
				}
			}
			else
			{
				levels[i] = top.level;
				if (top.override && kind !== K.BN)
					kinds[i] = top.override;
			}
		}
		
		// X10. Isolating run sequences. BN-like characters stay in place (X9 / 5.2) but are skipped.
		const runs = [];
		let run = null;
		for (let i = 0; i < n; i++)
		{
			const kind = kinds[i];
			if (kind & BN_LIKE)
				continue;
			const isInit = !!(kind & ISOLATE_INIT);
			const isPDI = kind === K.PDI;
			if (run && levels[i] === run.level)
			{
				run.end = i;
				run.endsWithInit = isInit;
			}
			else
			{
				run = {start: i, end: i, level: levels[i], startsWithPDI: isPDI, endsWithInit: isInit};
				runs.push(run);
			}
		}
		
		const sequences = [];
		for (let r = 0; r < runs.length; r++)
		{
			run = runs[r];
			if (run.startsWithPDI && isolationPairs[run.start] != null)
				continue;
			
			const parts = [run];
			let current = run;
			while (current && current.endsWithInit && isolationPairs[current.end] != null)
			{
				const pdi = isolationPairs[current.end];
				let next = null;
				for (let i = r + 1; i < runs.length; i++)
				{
					if (runs[i].start === pdi)
					{
						next = runs[i];
						break;
					}
				}
				if (!next)
					break;
				parts.push(next);
				current = next;
			}
			
			const seq = [];
			for (let p = 0; p < parts.length; p++)
			{
				for (let i = parts[p].start; i <= parts[p].end; i++)
					seq.push(i);
			}
			if (!seq.length)
				continue;
			
			const first = seq[0];
			let prevLevel = paragraphLevel;
			for (let i = first - 1; i >= 0; i--)
			{
				if (!(kinds[i] & BN_LIKE))
				{
					prevLevel = levels[i];
					break;
				}
			}
			const last = seq[seq.length - 1];
			let nextLevel = paragraphLevel;
			if (!(kinds[last] & ISOLATE_INIT))
			{
				for (let i = last + 1; i < n; i++)
				{
					if (!(kinds[i] & BN_LIKE))
					{
						nextLevel = levels[i];
						break;
					}
				}
			}
			sequences.push({
				seq      : seq,
				sos      : (Math.max(prevLevel, levels[first]) & 1) ? K.R : K.L,
				eos      : (Math.max(nextLevel, levels[last]) & 1) ? K.R : K.L,
				embedDir : (levels[first] & 1) ? K.R : K.L
			});
		}
		
		for (let s = 0; s < sequences.length; s++)
		{
			const seq = sequences[s].seq;
			const sos = sequences[s].sos;
			const eos = sequences[s].eos;
			const embedDir = sequences[s].embedDir;
			const len = seq.length;
			
			// W1. NSM takes the previous non-BN type (or sos).
			for (let si = 0; si < len; si++)
			{
				const i = seq[si];
				if (kinds[i] !== K.NSM)
					continue;
				let prev = sos;
				for (let sj = si - 1; sj >= 0; sj--)
				{
					if (!(kinds[seq[sj]] & BN_LIKE))
					{
						prev = kinds[seq[sj]];
						break;
					}
				}
				kinds[i] = (prev & (ISOLATE_INIT | K.PDI)) ? K.ON : prev;
			}
			
			// W2. EN becomes AN when the previous strong type is AL.
			for (let si = 0; si < len; si++)
			{
				const i = seq[si];
				if (kinds[i] !== K.EN)
					continue;
				let prev = sos;
				for (let sj = si - 1; sj >= 0; sj--)
				{
					if (kinds[seq[sj]] & STRONG)
					{
						prev = kinds[seq[sj]];
						break;
					}
				}
				if (prev === K.AL)
					kinds[i] = K.AN;
			}
			
			// W3. AL becomes R.
			for (let si = 0; si < len; si++)
			{
				if (kinds[seq[si]] === K.AL)
					kinds[seq[si]] = K.R;
			}
			
			// W4. ES between EN, or CS between the same number type, joins that number.
			for (let si = 1; si < len - 1; si++)
			{
				const i = seq[si];
				if (!(kinds[i] & (K.ES | K.CS)))
					continue;
				let prev = 0;
				let next = 0;
				for (let sj = si - 1; sj >= 0; sj--)
				{
					prev = kinds[seq[sj]];
					if (!(prev & BN_LIKE))
						break;
				}
				for (let sj = si + 1; sj < len; sj++)
				{
					next = kinds[seq[sj]];
					if (!(next & BN_LIKE))
						break;
				}
				if (prev === next && (kinds[i] === K.ES ? prev === K.EN : (prev & (K.EN | K.AN))))
					kinds[i] = prev;
			}
			
			// W5. ET next to EN becomes EN.
			for (let si = 0; si < len; si++)
			{
				if (kinds[seq[si]] !== K.EN)
					continue;
				for (let sj = si - 1; sj >= 0; sj--)
				{
					const k = kinds[seq[sj]];
					if (k === K.ET || (k & BN_LIKE))
						kinds[seq[sj]] = K.EN;
					else
						break;
				}
				let sj = si + 1;
				while (sj < len)
				{
					const k = kinds[seq[sj]];
					if (k === K.ET || (k & BN_LIKE) || k === K.EN)
					{
						if (k !== K.EN)
							kinds[seq[sj]] = K.EN;
						sj++;
					}
					else
						break;
				}
				si = sj - 1;
			}
			
			// W6. Remaining separators and terminators become ON, and so do adjacent BN.
			for (let si = 0; si < len; si++)
			{
				const i = seq[si];
				if (!(kinds[i] & (K.ET | K.ES | K.CS)))
					continue;
				kinds[i] = K.ON;
				for (let sj = si - 1; sj >= 0 && (kinds[seq[sj]] & BN_LIKE); sj--)
					kinds[seq[sj]] = K.ON;
				for (let sj = si + 1; sj < len && (kinds[seq[sj]] & BN_LIKE); sj++)
					kinds[seq[sj]] = K.ON;
			}
			
			// W7. EN becomes L when the previous strong type is L.
			// This is what keeps "23" on "def" in an RTL paragraph of "Abc def!23".
			let prevStrong = sos;
			for (let si = 0; si < len; si++)
			{
				const i = seq[si];
				if (kinds[i] === K.EN)
				{
					if (prevStrong === K.L)
						kinds[i] = K.L;
				}
				else if (kinds[i] & STRONG)
				{
					prevStrong = kinds[i];
				}
			}
			
			// N0. Bracket pairs.
			const pairs = [];
			const stack = [];
			for (let si = 0; si < len; si++)
			{
				if (!(kinds[seq[si]] & NI))
					continue;
				const cp = codePoints[seq[si]];
				if (isOpeningBracket(cp))
				{
					if (stack.length < 63)
						stack.push({cp: cp, si: si});
					else
						break;
				}
				else if (isClosingBracket(cp))
				{
					for (let sp = stack.length - 1; sp >= 0; sp--)
					{
						if (bracketsMatch(stack[sp].cp, cp))
						{
							pairs.push([stack[sp].si, si]);
							stack.length = sp;
							break;
						}
					}
				}
			}
			pairs.sort(function(a, b) { return a[0] - b[0]; });
			
			for (let p = 0; p < pairs.length; p++)
			{
				const openSi = pairs[p][0];
				const closeSi = pairs[p][1];
				let found = false;
				let dir = 0;
				for (let si = openSi + 1; si < closeSi; si++)
				{
					if (kinds[seq[si]] & STRONG_FOR_N)
					{
						found = true;
						const lr = (kinds[seq[si]] & R_FOR_NEUTRALS) ? K.R : K.L;
						if (lr === embedDir)
						{
							dir = lr;
							break;
						}
					}
				}
				if (found && !dir)
				{
					dir = sos;
					for (let si = openSi - 1; si >= 0; si--)
					{
						if (kinds[seq[si]] & STRONG_FOR_N)
						{
							const lr = (kinds[seq[si]] & R_FOR_NEUTRALS) ? K.R : K.L;
							dir = (lr !== embedDir) ? lr : embedDir;
							break;
						}
					}
				}
				if (!dir)
					continue;
				
				kinds[seq[openSi]] = dir;
				kinds[seq[closeSi]] = dir;
				if (dir !== embedDir)
				{
					applyNsmAfterBracket(seq, kinds, originalKinds, openSi);
					applyNsmAfterBracket(seq, kinds, originalKinds, closeSi);
				}
			}
			
			// N1 / N2. Neutrals take the surrounding direction, or the embedding direction.
			for (let si = 0; si < len; si++)
			{
				if (!(kinds[seq[si]] & NI))
					continue;
				let start = si;
				let prev = sos;
				for (let sj = si - 1; sj >= 0; sj--)
				{
					if (kinds[seq[sj]] & BN_LIKE)
						start = sj;
					else
					{
						prev = (kinds[seq[sj]] & R_FOR_NEUTRALS) ? K.R : K.L;
						break;
					}
				}
				let end = si;
				let next = eos;
				for (let sj = si + 1; sj < len; sj++)
				{
					if (kinds[seq[sj]] & (NI | BN_LIKE))
						end = sj;
					else
					{
						next = (kinds[seq[sj]] & R_FOR_NEUTRALS) ? K.R : K.L;
						break;
					}
				}
				const resolved = (prev === next) ? prev : embedDir;
				for (let sj = start; sj <= end; sj++)
					kinds[seq[sj]] = resolved;
				si = end;
			}
		}
		
		// I1 / I2, then BN level inheritance and L1.
		for (let i = 0; i < n; i++)
		{
			const level = levels[i];
			const kind = kinds[i];
			if (level & 1)
			{
				if (kind & (K.L | K.EN | K.AN))
					levels[i] = level + 1;
			}
			else if (kind & K.R)
			{
				levels[i] = level + 1;
			}
			else if (kind & (K.AN | K.EN))
			{
				levels[i] = level + 2;
			}
			
			if (kind & BN_LIKE)
				levels[i] = i === 0 ? paragraphLevel : levels[i - 1];
			
			if (i === n - 1 || (originalKinds[i] & (K.S | K.B)))
			{
				for (let j = i; j >= 0 && (originalKinds[j] & TRAILING); j--)
					levels[j] = paragraphLevel;
			}
		}
		
		return levels;
	}
	
	function applyNsmAfterBracket(seq, kinds, originalKinds, bracketSi)
	{
		for (let si = bracketSi + 1; si < seq.length; si++)
		{
			if (!(kinds[seq[si]] & BN_LIKE))
			{
				if (originalKinds[seq[si]] === K.NSM)
					kinds[seq[si]] = kinds[seq[bracketSi]];
				break;
			}
		}
	}
	
	function visualOrder(levels, paragraphLevel)
	{
		const n = levels.length;
		const order = new Array(n);
		for (let i = 0; i < n; i++)
			order[i] = i;
		
		let maxLevel = paragraphLevel;
		let minOdd = Infinity;
		for (let i = 0; i < n; i++)
		{
			if (levels[i] > maxLevel)
				maxLevel = levels[i];
			if (levels[i] < minOdd)
				minOdd = levels[i] | 1;
		}
		if (minOdd === Infinity)
			return order;
		
		for (let lvl = maxLevel; lvl >= minOdd; lvl--)
		{
			for (let i = 0; i < n; i++)
			{
				if (levels[i] < lvl)
					continue;
				let j = i;
				while (j + 1 < n && levels[j + 1] >= lvl)
					j++;
				for (let a = i, b = j; a < b; a++, b--)
				{
					const tmp = order[a];
					order[a] = order[b];
					order[b] = tmp;
				}
				i = j;
			}
		}
		return order;
	}
	
	/**
	 * Class for handling bidirectional flow of text or other content.
	 * Characters are collected for the whole line and reordered with UAX #9 when the line ends.
	 * @param handler - handler for elements in the flow
	 * @constructor
	 */
	function BidiFlow(handler)
	{
		this.handler   = handler;
		this.buffer    = [];
		this.direction = null;
		this._items    = [];
		
		this.neutralBuffer = [];
		this.bufferLTR     = [];
	}
	/**
	 * @param isRtlDirection - main flow direction
	 */
	BidiFlow.prototype.begin = function(isRtlDirection)
	{
		this.direction = isRtlDirection ? AscBidi.DIRECTION.R : AscBidi.DIRECTION.L;
		this.reset();
	};
	BidiFlow.prototype.add = function(element, bidiType)
	{
		// Tabs, breaks and paragraph marks follow the paragraph direction.
		if (bidiType === AscBidi.TYPE.PM)
			bidiType = AscBidi.DIRECTION.R === this.direction ? AscBidi.TYPE.R : AscBidi.TYPE.L;
		else if (null == bidiType)
			bidiType = AscBidi.TYPE.ON;
		
		// A paragraph separator finishes the current sequence before itself.
		if (bidiType === AscBidi.TYPE.B)
			this.end();
		
		this._items.push({
			element : element,
			kind    : kindFromType(bidiType),
			cp      : codePointOf(element)
		});
	};
	BidiFlow.prototype.end = function()
	{
		if (!this._items.length)
			return;
		
		const items = this._items;
		this._items = [];
		
		const paragraphLevel = AscBidi.DIRECTION.R === this.direction ? 1 : 0;
		const kinds = new Array(items.length);
		const codePoints = new Array(items.length);
		for (let i = 0; i < items.length; i++)
		{
			kinds[i] = items[i].kind;
			codePoints[i] = items[i].cp;
		}
		
		const levels = resolveLevels(kinds, codePoints, paragraphLevel);
		const order = visualOrder(levels, paragraphLevel);
		for (let i = 0; i < order.length; i++)
		{
			const index = order[i];
			const rtl = (levels[index] & 1) === 1;
			this.handler.handleBidiFlow(items[index].element, rtl ? AscBidi.DIRECTION.R : AscBidi.DIRECTION.L);
		}
	};
	BidiFlow.prototype.reset = function()
	{
		this._items.length = 0;
		this.buffer.length = 0;
		this.neutralBuffer.length = 0;
		this.bufferLTR.length = 0;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.BidiFlow = BidiFlow;
	
})(window);

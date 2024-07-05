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

(function (window) {

	let type = false;
	let Paragraph = null;
	let isAllowAutoCorrect = true;

	let functionNames = [
		'cos', 'acos', 'acosh', 'sin', 'tan', 'asin', 'asinh', 'sec',
		'acsc', 'atan', 'atanh', 'acsch', 'arcsinh', 'cot', 'acot', 'def',
		'arg', 'deg', 'det', 'dim', 'erf', 'acoth', 'csc', 'arcsin',
		'gcd', 'inf', 'asec', 'ker', 'asech', 'arccos', 'hom', 'lg',
		'arctan', 'sup', 'arcsec', 'arccot', 'arccsc', 'sinh', 'cosh',
		'tanh', 'coth', 'sech', 'csch', 'srcsinh', 'arctanh', 'arcsech', 'arccosh',
		'arccoth', 'arccsch', 'Pr', 'lin', 'exp', "sgn",
	];
	const limitFunctions = [
		"lim", "min", "max", "log", "ln"
	];
	const UnicodeSpecialScript = {
		"⁰": "0",
		"¹": "1",
		"²": "2",
		"³": "3",
		"⁴": "4",
		"⁵": "5",
		"⁶": "6",
		"⁷": "7",
		"⁸": "8",
		"⁹": "9",
		"ⁱ": "i",
		"ⁿ": "n",
		"⁺": "+",
		"⁻": "-",
		"⁼": "=",
		"⁽": "(",
		"⁾": ")",

		"₀": "0",
		"₁": "1",
		"₂": "2",
		"₃": "3",
		"₄": "4",
		"₅": "5",
		"₆": "6",
		"₇": "7",
		"₈": "8",
		"₉": "9",
		"₊": "+",
		"₋": "-",
		"₌": "=",
		"₍": "(",
		"₎": ")",
	}
	const GetTypeFont = {
		// Standart Word functions with higher proirity for linear format
		"\\mathcal": 7,
		"\\mathsf": 3,
		"\\mathrm": -1,
		"\\mathit": 1,
		"\\mathfrak": 9,
		"\\mathbfcal": 8,
		"\\mathbf": 0,
		"\\mathbb": 12,

		// other LaTeX functions
		"\\sf": 3,
		"\\script": 7,
		"\\scr": 7,
		"\\rm": -1,
		"\\oldstyle": 7,
		"\\mathtt": 11,
		"\\mathsfit": 5,
		"\\mathsfbfit": 6,
		"\\mathsfbf": 4,
		"\\mathbfit": 2,
		"\\mathbffrak": 10,
		"\\it": 1,
		"\\fraktur": 9,
		"\\frak": 9,
		"\\double": 12,
	}

	function SetIsAllowAutoCorrect (isAllow)
	{
		isAllowAutoCorrect = isAllow;
	}

	// List of MathFont
	const GetMathFontChar = {
		'A': { 0: '𝐀', 1: '𝐴', 2: '𝑨', 3: '𝖠', 4: '𝗔', 5: '𝘈', 6: '𝘼', 7: '𝒜', 8: '𝓐', 9: '𝔄', 10: '𝕬', 11: '𝙰', 12: '𝔸'},
		'B': { 0: '𝐁', 1: '𝐵', 2: '𝑩', 3: '𝖡', 4: '𝗕', 5: '𝘉', 6: '𝘽', 7: 'ℬ', 8: '𝓑', 9: '𝔅', 10: '𝕭', 11: '𝙱', 12: '𝔹'},
		'C': { 0: '𝐂', 1: '𝐶', 2: '𝑪', 3: '𝖢', 4: '𝗖', 5: '𝘊', 6: '𝘾', 7: '𝒞', 8: '𝓒', 9: 'ℭ', 10: '𝕮', 11: '𝙲', 12: 'ℂ'},
		'D': { 0: '𝐃', 1: '𝐷', 2: '𝑫', 3: '𝖣', 4: '𝗗', 5: '𝘋', 6: '𝘿', 7: '𝒟', 8: '𝓓', 9: '𝔇', 10: '𝕯', 11: '𝙳', 12: '𝔻'},
		'E': { 0: '𝐄', 1: '𝐸', 2: '𝑬', 3: '𝖤', 4: '𝗘', 5: '𝘌', 6: '𝙀', 7: 'ℰ', 8: '𝓔', 9: '𝔈', 10: '𝕰', 11: '𝙴', 12: '𝔼'},
		'F': { 0: '𝐅', 1: '𝐹', 2: '𝑭', 3: '𝖥', 4: '𝗙', 5: '𝘍', 6: '𝙁', 7: 'ℱ', 8: '𝓕', 9: '𝔉', 10: '𝕱', 11: '𝙵', 12: '𝔽'},
		'G': { 0: '𝐆', 1: '𝐺', 2: '𝑮', 3: '𝖦', 4: '𝗚', 5: '𝘎', 6: '𝙂', 7: '𝒢', 8: '𝓖', 9: '𝔊', 10: '𝕲', 11: '𝙶', 12: '𝔾'},
		'H': { 0: '𝐇', 1: '𝐻', 2: '𝑯', 3: '𝖧', 4: '𝗛', 5: '𝘏', 6: '𝙃', 7: 'ℋ', 8: '𝓗', 9: 'ℌ', 10: '𝕳', 11: '𝙷', 12: 'ℍ'},
		'I': { 0: '𝐈', 1: '𝐼', 2: '𝑰', 3: '𝖨', 4: '𝗜', 5: '𝘐', 6: '𝙄', 7: 'ℐ', 8: '𝓘', 9: 'ℑ', 10: '𝕴', 11: '𝙸', 12: '𝕀'},
		'J': { 0: '𝐉', 1: '𝐽', 2: '𝑱', 3: '𝖩', 4: '𝗝', 5: '𝘑', 6: '𝙅', 7: '𝒥', 8: '𝓙', 9: '𝔍', 10: '𝕵', 11: '𝙹', 12: '𝕁'},
		'K': { 0: '𝐊', 1: '𝐾', 2: '𝑲', 3: '𝖪', 4: '𝗞', 5: '𝘒', 6: '𝙆', 7: '𝒦', 8: '𝓚', 9: '𝔎', 10: '𝕶', 11: '𝙺', 12: '𝕂'},
		'L': { 0: '𝐋', 1: '𝐿', 2: '𝑳', 3: '𝖫', 4: '𝗟', 5: '𝘓', 6: '𝙇', 7: 'ℒ', 8: '𝓛', 9: '𝔏', 10: '𝕷', 11: '𝙻', 12: '𝕃'},
		'M': { 0: '𝐌', 1: '𝑀', 2: '𝑴', 3: '𝖬', 4: '𝗠', 5: '𝘔', 6: '𝙈', 7: 'ℳ', 8: '𝓜', 9: '𝔐', 10: '𝕸', 11: '𝙼', 12: '𝕄'},
		'N': { 0: '𝐍', 1: '𝑁', 2: '𝑵', 3: '𝖭', 4: '𝗡', 5: '𝘕', 6: '𝙉', 7: '𝒩', 8: '𝓝', 9: '𝔑', 10: '𝕹', 11: '𝙽', 12: 'ℕ'},
		'O': { 0: '𝐎', 1: '𝑂', 2: '𝑶', 3: '𝖮', 4: '𝗢', 5: '𝘖', 6: '𝙊', 7: '𝒪', 8: '𝓞', 9: '𝔒', 10: '𝕺', 11: '𝙾', 12: '𝕆'},
		'P': { 0: '𝐏', 1: '𝑃', 2: '𝑷', 3: '𝖯', 4: '𝗣', 5: '𝘗', 6: '𝙋', 7: '𝒫', 8: '𝓟', 9: '𝔓', 10: '𝕻', 11: '𝙿', 12: 'ℙ'},
		'Q': { 0: '𝐐', 1: '𝑄', 2: '𝑸', 3: '𝖰', 4: '𝗤', 5: '𝘘', 6: '𝙌', 7: '𝒬', 8: '𝓠', 9: '𝔔', 10: '𝕼', 11: '𝚀', 12: 'ℚ'},
		'R': { 0: '𝐑', 1: '𝑅', 2: '𝑹', 3: '𝖱', 4: '𝗥', 5: '𝘙', 6: '𝙍', 7: 'ℛ', 8: '𝓡', 9: 'ℜ', 10: '𝕽', 11: '𝚁', 12: 'ℝ'},
		'S': { 0: '𝐒', 1: '𝑆', 2: '𝑺', 3: '𝖲', 4: '𝗦', 5: '𝘚', 6: '𝙎', 7: '𝒮', 8: '𝓢', 9: '𝔖', 10: '𝕾', 11: '𝚂', 12: '𝕊'},
		'T': { 0: '𝐓', 1: '𝑇', 2: '𝑻', 3: '𝖳', 4: '𝗧', 5: '𝘛', 6: '𝙏', 7: '𝒯', 8: '𝓣', 9: '𝔗', 10: '𝕿', 11: '𝚃', 12: '𝕋'},
		'U': { 0: '𝐔', 1: '𝑈', 2: '𝑼', 3: '𝖴', 4: '𝗨', 5: '𝘜', 6: '𝙐', 7: '𝒰', 8: '𝓤', 9: '𝔘', 10: '𝖀', 11: '𝚄', 12: '𝕌'},
		'V': { 0: '𝐕', 1: '𝑉', 2: '𝑽', 3: '𝖵', 4: '𝗩', 5: '𝘝', 6: '𝙑', 7: '𝒱', 8: '𝓥', 9: '𝔙', 10: '𝖁', 11: '𝚅', 12: '𝕍'},
		'W': { 0: '𝐖', 1: '𝑊', 2: '𝑾', 3: '𝖶', 4: '𝗪', 5: '𝘞', 6: '𝙒', 7: '𝒲', 8: '𝓦', 9: '𝔚', 10: '𝖂', 11: '𝚆', 12: '𝕎'},
		'X': { 0: '𝐗', 1: '𝑋', 2: '𝑿', 3: '𝖷', 4: '𝗫', 5: '𝘟', 6: '𝙓', 7: '𝒳', 8: '𝓧', 9: '𝔛', 10: '𝖃', 11: '𝚇', 12: '𝕏'},
		'Y': { 0: '𝐘', 1: '𝑌', 2: '𝒀', 3: '𝖸', 4: '𝗬', 5: '𝘠', 6: '𝙔', 7: '𝒴', 8: '𝓨', 9: '𝔜', 10: '𝖄', 11: '𝚈', 12: '𝕐'},
		'Z': { 0: '𝐙', 1: '𝑍', 2: '𝒁', 3: '𝖹', 4: '𝗭', 5: '𝘡', 6: '𝙕', 7: '𝒵', 8: '𝓩', 9: 'ℨ', 10: '𝖅', 11: '𝚉', 12: 'ℤ'},
		'a': { 0: '𝐚', 1: '𝑎', 2: '𝒂', 3: '𝖺', 4: '𝗮', 5: '𝘢', 6: '𝙖', 7: '𝒶', 8: '𝓪', 9: '𝔞', 10: '𝖆', 11: '𝚊', 12: '𝕒'},
		'b': { 0: '𝐛', 1: '𝑏', 2: '𝒃', 3: '𝖻', 4: '𝗯', 5: '𝘣', 6: '𝙗', 7: '𝒷', 8: '𝓫', 9: '𝔟', 10: '𝖇', 11: '𝚋', 12: '𝕓'},
		'c': { 0: '𝐜', 1: '𝑐', 2: '𝒄', 3: '𝖼', 4: '𝗰', 5: '𝘤', 6: '𝙘', 7: '𝒸', 8: '𝓬', 9: '𝔠', 10: '𝖈', 11: '𝚌', 12: '𝕔'},
		'd': { 0: '𝐝', 1: '𝑑', 2: '𝒅', 3: '𝖽', 4: '𝗱', 5: '𝘥', 6: '𝙙', 7: '𝒹', 8: '𝓭', 9: '𝔡', 10: '𝖉', 11: '𝚍', 12: '𝕕'},
		'e': { 0: '𝐞', 1: '𝑒', 2: '𝒆', 3: '𝖾', 4: '𝗲', 5: '𝘦', 6: '𝙚', 7: 'ℯ', 8: '𝓮', 9: '𝔢', 10: '𝖊', 11: '𝚎', 12: '𝕖'},
		'f': { 0: '𝐟', 1: '𝑓', 2: '𝒇', 3: '𝖿', 4: '𝗳', 5: '𝘧', 6: '𝙛', 7: '𝒻', 8: '𝓯', 9: '𝔣', 10: '𝖋', 11: '𝚏', 12: '𝕗'},
		'g': { 0: '𝐠', 1: '𝑔', 2: '𝒈', 3: '𝗀', 4: '𝗴', 5: '𝘨', 6: '𝙜', 7: 'ℊ', 8: '𝓰', 9: '𝔤', 10: '𝖌', 11: '𝚐', 12: '𝕘'},
		'h': { 0: '𝐡', 1: 'ℎ', 2: '𝒉', 3: '𝗁', 4: '𝗵', 5: '𝘩', 6: '𝙝', 7: '𝒽', 8: '𝓱', 9: '𝔥', 10: '𝖍', 11: '𝚑', 12: '𝕙'},
		'i': { 0: '𝐢', 1: '𝑖', 2: '𝒊', 3: '𝗂', 4: '𝗶', 5: '𝘪', 6: '𝙞', 7: '𝒾', 8: '𝓲', 9: '𝔦', 10: '𝖎', 11: '𝚒', 12: '𝕚'},
		'j': { 0: '𝐣', 1: '𝑗', 2: '𝒋', 3: '𝗃', 4: '𝗷', 5: '𝘫', 6: '𝙟', 7: '𝒿', 8: '𝓳', 9: '𝔧', 10: '𝖏', 11: '𝚓', 12: '𝕛'},
		'k': { 0: '𝐤', 1: '𝑘', 2: '𝒌', 3: '𝗄', 4: '𝗸', 5: '𝘬', 6: '𝙠', 7: '𝓀', 8: '𝓴', 9: '𝔨', 10: '𝖐', 11: '𝚔', 12: '𝕜'},
		'l': { 0: '𝐥', 1: '𝑙', 2: '𝒍', 3: '𝗅', 4: '𝗹', 5: '𝘭', 6: '𝙡', 7: '𝓁', 8: '𝓵', 9: '𝔩', 10: '𝖑', 11: '𝚕', 12: '𝕝'},
		'm': { 0: '𝐦', 1: '𝑚', 2: '𝒎', 3: '𝗆', 4: '𝗺', 5: '𝘮', 6: '𝙢', 7: '𝓂', 8: '𝓶', 9: '𝔪', 10: '𝖒', 11: '𝚖', 12: '𝕞'},
		'n': { 0: '𝐧', 1: '𝑛', 2: '𝒏', 3: '𝗇', 4: '𝗻', 5: '𝘯', 6: '𝙣', 7: '𝓃', 8: '𝓷', 9: '𝔫', 10: '𝖓', 11: '𝚗', 12: '𝕟'},
		'o': { 0: '𝐨', 1: '𝑜', 2: '𝒐', 3: '𝗈', 4: '𝗼', 5: '𝘰', 6: '𝙤', 7: 'ℴ', 8: '𝓸', 9: '𝔬', 10: '𝖔', 11: '𝚘', 12: '𝕠'},
		'p': {0: '𝐩',1: '𝑝',2: '𝒑',3: '𝗉',4: '𝗽',5: '𝘱',6: '𝙥',7: '𝓅',8: '𝓹',9: '𝔭',10: '𝖕',11: '𝚙',12: '𝕡'},
		'q': { 0: '𝐪', 1: '𝑞', 2: '𝒒', 3: '𝗊', 4: '𝗾', 5: '𝘲', 6: '𝙦', 7: '𝓆', 8: '𝓺', 9: '𝔮', 10: '𝖖', 11: '𝚚', 12: '𝕢'},
		'r': { 0: '𝐫', 1: '𝑟', 2: '𝒓', 3: '𝗋', 4: '𝗿', 5: '𝘳', 6: '𝙧', 7: '𝓇', 8: '𝓻', 9: '𝔯', 10: '𝖗', 11: '𝚛', 12: '𝕣'},
		's': { 0: '𝐬', 1: '𝑠', 2: '𝒔', 3: '𝗌', 4: '𝘀', 5: '𝘴', 6: '𝙨', 7: '𝓈', 8: '𝓼', 9: '𝔰', 10: '𝖘', 11: '𝚜', 12: '𝕤'},
		't': { 0: '𝐭', 1: '𝑡', 2: '𝒕', 3: '𝗍', 4: '𝘁', 5: '𝘵', 6: '𝙩', 7: '𝓉', 8: '𝓽', 9: '𝔱', 10: '𝖙', 11: '𝚝', 12: '𝕥'},
		'u': { 0: '𝐮', 1: '𝑢', 2: '𝒖', 3: '𝗎', 4: '𝘂', 5: '𝘶', 6: '𝙪', 7: '𝓊', 8: '𝓾', 9: '𝔲', 10: '𝖚', 11: '𝚞', 12: '𝕦'},
		'v': { 0: '𝐯', 1: '𝑣', 2: '𝒗', 3: '𝗏', 4: '𝘃', 5: '𝘷', 6: '𝙫', 7: '𝓋', 8: '𝓿', 9: '𝔳', 10: '𝖛', 11: '𝚟', 12: '𝕧'},
		'w': { 0: '𝐰', 1: '𝑤', 2: '𝒘', 3: '𝗐', 4: '𝘄', 5: '𝘸', 6: '𝙬', 7: '𝓌', 8: '𝔀', 9: '𝔴', 10: '𝖜', 11: '𝚠', 12: '𝕨'},
		'x': { 0: '𝐱', 1: '𝑥', 2: '𝒙', 3: '𝗑', 4: '𝘅', 5: '𝘹', 6: '𝙭', 7: '𝓍', 8: '𝔁', 9: '𝔵', 10: '𝖝', 11: '𝚡', 12: '𝕩'},
		'y': { 0: '𝐲', 1: '𝑦', 2: '𝒚', 3: '𝗒', 4: '𝘆', 5: '𝘺', 6: '𝙮', 7: '𝓎', 8: '𝔂', 9: '𝔶', 10: '𝖞', 11: '𝚢', 12: '𝕪'},
		'z': { 0: '𝐳', 1: '𝑧', 2: '𝒛', 3: '𝗓', 4: '𝘇', 5: '𝘻', 6: '𝙯', 7: '𝓏', 8: '𝔃', 9: '𝔷', 10: '𝖟', 11: '𝚣', 12: '𝕫'},
		// 'ı': {mathit: '𝚤'},
		// 'ȷ': {mathit: '𝚥'},
		'Α': {0: '𝚨', 1: '𝛢', 2: '𝜜', 4: '𝝖', 6: '𝞐'},
		'Β': {0: '𝚩', 1: '𝛣', 2: '𝜝', 4: '𝝗', 6: '𝞑'},
		'Γ': {0: '𝚪', 1: '𝛤', 2: '𝜞', 4: '𝝘', 6: '𝞒'},
		'Δ': {0: '𝚫', 1: '𝛥', 2: '𝜟', 4: '𝝙', 6: '𝞓'},
		'Ε': {0: '𝚬', 1: '𝛦', 2: '𝜠', 4: '𝝚', 6: '𝞔'},
		'Ζ': {0: '𝚭', 1: '𝛧', 2: '𝜡', 4: '𝝛', 6: '𝞕'},
		'Η': {0: '𝚮', 1: '𝛨', 2: '𝜢', 4: '𝝜', 6: '𝞖'},
		'Θ': {0: '𝚯', 1: '𝛩', 2: '𝜣', 4: '𝝝', 6: '𝞗'},
		'Ι': {0: '𝚰', 1: '𝛪', 2: '𝜤', 4: '𝝞', 6: '𝞘'},
		'Κ': {0: '𝚱', 1: '𝛫', 2: '𝜥', 4: '𝝟', 6: '𝞙'},
		'Λ': {0: '𝚲', 1: '𝛬', 2: '𝜦', 4: '𝝠', 6: '𝞚'},
		'Μ': {0: '𝚳', 1: '𝛭', 2: '𝜧', 4: '𝝡', 6: '𝞛'},
		'Ν': {0: '𝚴', 1: '𝛮', 2: '𝜨', 4: '𝝢', 6: '𝞜'},
		'Ξ': {0: '𝚵', 1: '𝛯', 2: '𝜩', 4: '𝝣', 6: '𝞝'},
		'Ο': {0: '𝚶', 1: '𝛰', 2: '𝜪', 4: '𝝤', 6: '𝞞'},
		'Π': {0: '𝚷', 1: '𝛱', 2: '𝜫', 4: '𝝥', 6: '𝞟'},
		'Ρ': {0: '𝚸', 1: '𝛲', 2: '𝜬', 4: '𝝦', 6: '𝞠'},
		'ϴ': {0: '𝚹', 1: '𝛳', 2: '𝜭', 4: '𝝧', 6: '𝞡'},
		'Σ': {0: '𝚺', 1: '𝛴', 2: '𝜮', 4: '𝝨', 6: '𝞢'},
		'Τ': {0: '𝚻', 1: '𝛵', 2: '𝜯', 4: '𝝩', 6: '𝞣'},
		'Υ': {0: '𝚼', 1: '𝛶', 2: '𝜰', 4: '𝝪', 6: '𝞤'},
		'Φ': {0: '𝚽', 1: '𝛷', 2: '𝜱', 4: '𝝫', 6: '𝞥'},
		'Χ': {0: '𝚾', 1: '𝛸', 2: '𝜲', 4: '𝝬', 6: '𝞦'},
		'Ψ': {0: '𝚿', 1: '𝛹', 2: '𝜳', 4: '𝝭', 6: '𝞧'},
		'Ω': {0: '𝛀', 1: '𝛺', 2: '𝜴', 4: '𝝮', 6: '𝞨'},
		'∇': {0: '𝛁', 1: '𝛻', 2: '𝜵', 4: '𝝯', 6: '𝞩'},
		'α': {0: '𝛂', 1: '𝛼', 2: '𝜶', 4: '𝝰', 6: '𝞪'},
		'β': {0: '𝛃', 1: '𝛽', 2: '𝜷', 4: '𝝱', 6: '𝞫'},
		'γ': {0: '𝛄', 1: '𝛾', 2: '𝜸', 4: '𝝲', 6: '𝞬'},
		'δ': {0: '𝛅', 1: '𝛿', 2: '𝜹', 4: '𝝳', 6: '𝞭'},
		'ε': {0: '𝛆', 1: '𝜀', 2: '𝜺', 4: '𝝴', 6: '𝞮'},
		'ζ': {0: '𝛇', 1: '𝜁', 2: '𝜻', 4: '𝝵', 6: '𝞯'},
		'η': {0: '𝛈', 1: '𝜂', 2: '𝜼', 4: '𝝶', 6: '𝞰'},
		'θ': {0: '𝛉', 1: '𝜃', 2: '𝜽', 4: '𝝷', 6: '𝞱'},
		'ι': {0: '𝛊', 1: '𝜄', 2: '𝜾', 4: '𝝸', 6: '𝞲'},
		'κ': {0: '𝛋', 1: '𝜅', 2: '𝜿', 4: '𝝹', 6: '𝞳'},
		'λ': {0: '𝛌', 1: '𝜆', 2: '𝝀', 4: '𝝺', 6: '𝞴'},
		'μ': {0: '𝛍', 1: '𝜇', 2: '𝝁', 4: '𝝻', 6: '𝞵'},
		'ν': {0: '𝛎', 1: '𝜈', 2: '𝝂', 4: '𝝼', 6: '𝞶'},
		'ξ': {0: '𝛏', 1: '𝜉', 2: '𝝃', 4: '𝝽', 6: '𝞷'},
		'ο': {0: '𝛐', 1: '𝜊', 2: '𝝄', 4: '𝝾', 6: '𝞸'},
		'π': {0: '𝛑', 1: '𝜋', 2: '𝝅', 4: '𝝿', 6: '𝞹'},
		'ρ': {0: '𝛒', 1: '𝜌', 2: '𝝆', 4: '𝞀', 6: '𝞺'},
		'ς': {0: '𝛓', 1: '𝜍', 2: '𝝇', 4: '𝞁', 6: '𝞻'},
		'σ': {0: '𝛔', 1: '𝜎', 2: '𝝈', 4: '𝞂', 6: '𝞼'},
		'τ': {0: '𝛕', 1: '𝜏', 2: '𝝉', 4: '𝞃', 6: '𝞽'},
		'υ': {0: '𝛖', 1: '𝜐', 2: '𝝊', 4: '𝞄', 6: '𝞾'},
		'φ': {0: '𝛗', 1: '𝜑', 2: '𝝋', 4: '𝞅', 6: '𝞿'},
		'χ': {0: '𝛘', 1: '𝜒', 2: '𝝌', 4: '𝞆', 6: '𝟀'},
		'ψ': {0: '𝛙', 1: '𝜓', 2: '𝝍', 4: '𝞇', 6: '𝟁'},
		'ω': {0: '𝛚', 1: '𝜔', 2: '𝝎', 4: '𝞈', 6: '𝟂'},
		'∂': {0: '𝛛', 1: '𝜕', 2: '𝝏', 4: '𝞉', 6: '𝟃'},
		'ϵ': {0: '𝛜', 1: '𝜖', 2: '𝝐', 4: '𝞊', 6: '𝟄'},
		'ϑ': {0: '𝛝', 1: '𝜗', 2: '𝝑', 4: '𝞋', 6: '𝟅'},
		'ϰ': {0: '𝛞', 1: '𝜘', 2: '𝝒', 4: '𝞌', 6: '𝟆'},
		'ϕ': {0: '𝛟', 1: '𝜙', 2: '𝝓', 4: '𝞍', 6: '𝟇'},
		'ϱ': {0: '𝛠', 1: '𝜚', 2: '𝝔', 4: '𝞎', 6: '𝟈'},
		'ϖ': {0: '𝛡', 1: '𝜛', 2: '𝝕', 4: '𝞏', 6: '𝟉'},
		'Ϝ': {0: '𝟊'},
		'ϝ': {0: '𝟋'},
		'0': {0: '𝟎', 12: '𝟘', 3: '𝟢', 4: '𝟬', 11: '𝟶'},
		'1': {0: '𝟏', 12: '𝟙', 3: '𝟣', 4: '𝟭', 11: '𝟷'},
		'2': {0: '𝟐', 12: '𝟚', 3: '𝟤', 4: '𝟮', 11: '𝟸'},
		'3': {0: '𝟑', 12: '𝟛', 3: '𝟥', 4: '𝟯', 11: '𝟹'},
		'4': {0: '𝟒', 12: '𝟜', 3: '𝟦', 4: '𝟰', 11: '𝟺'},
		'5': {0: '𝟓', 12: '𝟝', 3: '𝟧', 4: '𝟱', 11: '𝟻'},
		'6': {0: '𝟔', 12: '𝟞', 3: '𝟨', 4: '𝟲', 11: '𝟼'},
		'7': {0: '𝟕', 12: '𝟟', 3: '𝟩', 4: '𝟳', 11: '𝟽'},
		'8': {0: '𝟖', 12: '𝟠', 3: '𝟪', 4: '𝟴', 11: '𝟾'},
		'9': {0: '𝟗', 12: '𝟡', 3: '𝟫', 4: '𝟵', 11: '𝟿'},
	};

	function LexerLiterals()
	{
		this.Unicode = {};
		this.LaTeX = {};
		this.isUseLaTeXBrackets = false;

		this.Init();
	}
	LexerLiterals.prototype.Init = function (isUseUnicodeInLaTeX)
	{
		let names = Object.keys(this.LaTeX);

		if (names.length < 1)
			return false;

		for (let i = 0; i < names.length; i++)
		{
			let name = names[i];
			let data = this.LaTeX[name];

			if (typeof(data) === "string")
			{
				this.SetUnicodeFromLaTeX(data, name);
				if (isUseUnicodeInLaTeX)
					this.private_AddToLaTeX(data, data);
			}
		}

		return true;
	};
	LexerLiterals.prototype.IsLaTeXInclude = function (name)
	{
		if (!this.LaTeX)
			return false;

		return this.LaTeX[name] !== undefined;
	};
	LexerLiterals.prototype.IsUnicodeInclude = function (name)
	{
		if (!this.Unicode)
			return false;
		return this.Unicode[name] !== undefined;
	};
	LexerLiterals.prototype.AddToLaTeX = function (name, data)
	{
		if (!this.IsLaTeXInclude(name))
			this.private_AddToLaTeX(name, data);
	};
	LexerLiterals.prototype.AddToUnicode = function (name, data)
	{
		if (!this.IsUnicodeInclude(name))
			this.private_AddToUnicode(name, data);
	};
	LexerLiterals.prototype.private_AddToLaTeX = function (name, data)
	{
		this.LaTeX[name] = data;
		this.SetUnicodeFromLaTeX(data, name);
	};
	LexerLiterals.prototype.private_AddToUnicode = function (name, data)
	{
		this.Unicode[name] = data;
		this.SetLaTeXFromUnicode(data, name);
	};
	LexerLiterals.prototype.private_GetLaTeXWord = function (arrStr)
	{
		if (!arrStr || !arrStr[0])
			return;

		let strFunc = ""
		//"\\";

		// remove regexp

		if (this.isUseLaTeXBrackets)
		{
			let isStartBracket = false;
			let isEndBracket = false;
			let isSlashes = false;
			for (let index = 0; arrStr[index] && /[a-zA-Z{}\\]/.test(arrStr[index]); index++)
			{
				if (arrStr[index] === "{")
				{
					if (!isStartBracket)
						isStartBracket = true;
					else
						return strFunc;
				}
				else if (arrStr[index] === "}")
				{
					if (!isEndBracket && isStartBracket)
						isEndBracket = true;
					else
						return strFunc;
				}
				else if (arrStr[index] === "\\")
				{
					if (!isSlashes)
						isSlashes = true;
					else
						return strFunc;
				}

				strFunc += arrStr[index];

				if (this.LaTeX && this.LaTeX[strFunc])
					return strFunc;
			}
		}
		else
		{
			let isSlashes = false;


			for (let index = 0; arrStr[index] && /[a-zA-Z\\]/.test(arrStr[index]); index++)
			{
				if (arrStr[index] === "\\")
				{
					if (!isSlashes)
						isSlashes = true;
					else
						return strFunc;
				}

				strFunc += arrStr[index]

				if (this.LaTeX && this.LaTeX[strFunc])
					return strFunc;
			}
		};

		return strFunc;
	};
	LexerLiterals.prototype.SetUnicodeFromLaTeX= function (name, data)
	{
		this.Unicode[name] = data;
	};
	LexerLiterals.prototype.SetLaTeXFromUnicode = function (name, data)
	{
		this.LaTeX[name] = data;
	};
	LexerLiterals.prototype.GetToken = function (type, str)
	{
		if (this.GetByOneRule)
			return this.GetByOneRule(str);

		if (!type)
			return this.GetUnicodeToken(str);
		else
			return this.GetLaTeXToken(str);
	};
	LexerLiterals.prototype.GetUnicodeToken = function (str)
	{
		if (this.IsUnicodeInclude(str[0]))
			return str[0];
	};
	LexerLiterals.prototype.GetLaTeXToken = function (str)
	{
		let word = this.private_GetLaTeXWord(str);

		if (typeof word === "string" && this.IsLaTeXInclude(word))
			return word;

		else if (this.IsLaTeXInclude(str[0]))
			return str[0];
	};
	LexerLiterals.prototype.GetLaTeXWordFromSymbol = function (str)
	{
		let arr = Object.entries(this.LaTeX);

		for (let i = 0; i < arr.length; i++)
		{
			let curArr = arr[i];

			if (curArr[1] === str)
				return curArr[0];
		}
	}
	// Search in Unicode group of tokens
	LexerLiterals.prototype.SearchU = function (str)
	{
		return this.IsUnicodeInclude(str);
	};
	// Search in LaTeX group of tokens
	LexerLiterals.prototype.SearchL = function (str)
	{
		return this.IsLaTeXInclude(str);
	};

	function TokenChars()
	{
		this.id = 0;
	}
	TokenChars.prototype = Object.create(LexerLiterals.prototype);
	TokenChars.prototype.constructor = TokenChars;
	TokenChars.prototype.GetByOneRule = function(arrStr)
	{
		if (arrStr[0])
			return arrStr[0];
	};
	TokenChars.prototype.SearchU = function (str)
	{
		if (str[0])
			return true;
	}

	function TokenNumbers()
	{
		this.id = 1;
	}
	TokenNumbers.prototype = Object.create(LexerLiterals.prototype);
	TokenNumbers.prototype.constructor = TokenNumbers;
	TokenNumbers.prototype.GetByOneRule = function (arrStr)
	{
		if (['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'].includes(arrStr[0]))
			return arrStr[0];
	};

	function TokenOperators()
	{
		this.id = 2;
		this.LaTeX = {
			"\\angle" : "∠",
			"\\approx" : "≈",
			"\\ast" : "∗",
			"\\asymp" : "≍",
			"\\because" : "∵",
			"\\bot" : "⊥",
			"\\bowtie" : "⋈",
			"\\bullet" : "∙",
			"\\cap" : "∩",
			"\\cdot" : "⋅",
			"\\cdots" : "⋯",
			"\\circ" : "∘",
			"\\clubsuit" : "♣",
			"\\cong" : "≅",
			"\\cup" : "∪",
			"\\ddots" : "⋱",
			"\\diamond" : "⋄",
			"\\diamondsuit" : "♢",
			"\\div" : "÷",
			"\\doteq" : "≐",
			"\\dots" : "…",
			"\\Downarrow" : "⇓",
			"\\downarrow" : "↓",
			"\\equiv" : "≡",
			"\\exists" : "∃",
			"\\forall" : "∀",
			"\\ge" : "≥",
			"\\geq" : "≥",
			"\\gg" : "≫",
			"\\heartsuit" : "♡",
			"\\in" : "∈",
			"\\ldots" : "…",
			"\\le" : "≤",
			"\\leq" : "≤",
			"\\ll" : "≪",
			"\\Longleftarrow" : "⟸",
			"\\longleftarrow" : "⟵",
			"\\Longleftrightarrow" : "⟺",
			"\\longleftrightarrow" : "⟷",
			"\\Longrightarrow" : "⟹",
			"\\longrightarrow" : "⟶",
			"\\ne" : "≠",
			"\\nearrow" : "↗",
			"\\neg" : "¬",
			"\\neq" : "≠",
			"\\ni" : "∋",
			"\\nwarrow" : "↖",
			"\\odot" : "⊙",
			"\\ominus" : "⊖",
			"\\oplus" : "⊕",
			//"\\oslash" : "⊘", // todo
			"\\otimes" : "⊗",
			"\\parallel" : "∥",
			"\\prcue" : "≼",
			"\\prec" : "≺",
			"\\preceq" : "⪯",
			"\\preccurlyeq" : "≼",
			"\\propto" : "∝",
			"\\ratio" : "∶",
			"\\rddots" : "⋰",
			"\\searrow" : "↙",
			"\\setminus" : "∖",
			"\\sim" : "∼",
			"\\simeq" : "≃",
			"\\sqcap" : "⊓",
			"\\sqcup" : "⊔",
			"\\sqsubseteq" : "⊑",
			"\\sqsuperseteq" : "⊒",
			"\\star" : "⋆",
			"\\subset" : "⊂",
			"\\subseteq" : "⊆",
			"\\succ" : "≻",
			"\\succeq" : "≽",
			"\\superset" : "⊃",
			"\\superseteq" : "⊇",
			"\\swarrow" : "↘",
			"\\therefore" : "∴",
			"\\times" : "×",
			"\\top" : "⊤",
			"\\Uparrow" : "⇑",
			"\\uparrow" : "↑",
			"\\Updownarrow" : "⇕",
			"\\updownarrow" : "↕",
			"\\uplus" : "⊎",
			"\\vdots" : "⋮",
			"\\vee" : "∨",
			"\\wedge" : "∧",
			"\\wr" : "≀",
			"\\boxdot" : "⊡",
			"\\boxminus" : "⊟",
			"\\boxplus" : "⊞",
			"\\defeq" : "≝",
			"\\degc" : "℃",
			"\\degf" : "℉",
			"\\Deltaeq": "≜",
			"\\frown": "⌑",
			"\\mp" : "∓",
			"\\notcontain" : "∌",
			"\\notelement" : "∉",
			"\\notin" : "∉",
		};
		this.Unicode = {
			"⁣" : 1,
			"⁤" : 1,
			"⨯" : 1,
			"⨝" : 1,
			"⟕" : 1,
			"⟖" : 1,
			"⟗" : 1,
			"⋉" : 1,
			"⋊" : 1,
			"▷" : 1,
			"+" : 1,
			"-" : 1,
			"*" : 1,
			"=" : 1,
			"≶" : 1,
			"≷" : 1,
		};
		this.Init(true);
	}
	TokenOperators.prototype = Object.create(LexerLiterals.prototype);
	TokenOperators.prototype.constructor = TokenOperators;

	function TokenOperand()
	{
		this.id = 3;
		this.LaTeX = {
			"\\aleph" : "ℵ",
			"\\alpha" : "α",
			"\\Alpha" : "Α",
			"\\beta" : "β",
			"\\beth" : "ℶ",
			"\\bet" : "ℶ",
			"\\chi" : "χ",
			"\\daleth" : "ℸ",
			"\\Dd" : "ⅅ",
			"\\dd" : "ⅆ",
			"\\degree" : "°",
			"\\Delta" : "Δ",
			"\\delta" : "δ",
			"\\ee" : "ⅇ",
			"\\ell" : "ℓ",
			"\\emptyset" : "∅",
			"\\epsilon" : "ϵ",
			"\\eta" : "η",
			"\\Gamma" : "Γ",
			"\\G" : "Γ",
			"\\gamma" : "γ",
			"\\gimel" : "ℷ",
			"\\hbar" : "ℏ",
			"\\ii" : "ⅈ",
			"\\Im" : "ℑ",
			"\\imath" : "ı",
			"\\inc" : "∆",
			"\\infty" : "∞",
			"\\iota" : "ι",
			"\\jj" : "ⅉ",
			"\\jmath" : "ȷ",
			"\\kappa" : "κ",
			"\\Lambda" : "Λ",
			"\\lambda" : "λ",
			"\\mu" : "μ",
			"\\nabla" : "∇",
			"\\nu" : "ν",
			"\\Omega" : "Ω",
			"\\omega" : "ω",
			"\\partial" : "∂",
			"\\Phi" : "Φ",
			"\\phi" : "π",
			"\\Psi" : "Ψ",
			"\\psi" : "ψ",
			"\\Re" : "ℜ",
			"\\rho" : "ρ",
			"\\Sigma" : "Σ",
			"\\sigma" : "σ",
			"\\tau" : "τ",
			"\\Theta" : "Θ",
			"\\theta" : "θ",
			"\\Upsilon" : "Υ",
			"\\upsilon" : "υ",
			"\\varepsilon" : "ε",
			"\\varphi" : "φ",
			"\\varpi" : "ϖ",
			"\\varrho" : "ϱ",
			"\\varsigma" : "ς",
			"\\vartheta" : "ϑ",
			"\\wp" : "℘",
			"\\Xi" : "Ξ",
			"\\xi" : "ξ",
			"\\zeta" : "ζ",
			"\\Beta"		:	"Β",
			"\\Epsilon"		:	"Ε",
			"\\Zeta"		:	"Ζ",
			"\\Eta"			: 	"Η",
			"\\Iota"		:	"Ι",
			"\\Kappa"		:	"Κ",
			"\\Mu"			:	"Μ",
			"\\Nu"			:	"Ν",
			"\\O"			: 	"Ο",
			"\\o"			:	"ο",
			"\\pi"			:	"π",
			"\\Pi"			:	"Π",
			"\\Rho"			:	"Ρ",
			"\\Tau"			:	"Τ",
			"\\Chi"			:	"Χ",
		};
		this.Unicode = {};
		this.Init();
	}
	TokenOperand.prototype = Object.create(LexerLiterals.prototype);
	TokenOperand.prototype.constructor = TokenOperand;

	function TokenOpenBrackets()
	{
		this.id = 4;
		this.Unicode = {
			"(" : 1,
		};
		this.LaTeX = {
			"\\begin" : "〖",
			"\\langle" : "⟨",
			"\\lbrace" : "{",
			"\\lbrack" : "[",
			"\\lceil" : "⌈",
			"\\lfloor" : "⌊",
			"\\lbbrack" : "⟦",
			"\\lmoust" : "⎰",
			"\\bra" : "⟨",
			"{" : "{",
			"(" : "(",
			"⟨" : "⟨",
			"[" : "[",
			"⌈" : "⌈",
			"⌊" : "⌊",
			"⟦" : "⟦",
			"⎰" : "⎰",
			"\\left" : "\\left",
		};
		this.Init();
	}
	TokenOpenBrackets.prototype = Object.create(LexerLiterals.prototype);
	TokenOpenBrackets.prototype.constructor = TokenOpenBrackets;
	TokenOpenBrackets.prototype.IsSimple = function (str)
	{
		return str === "(" ||
			str === "[" ||
			str === "{" ||
			str === "|"
	}

	function TokenSpecialBrackets()
	{
		this.id = 544;
		this.Unicode = {};
		this.LaTeX = {
			"\\open" : "├",
			"\\close" : "┤",
		};

		this.Init();
	}
	TokenSpecialBrackets.prototype = Object.create(LexerLiterals.prototype);
	TokenSpecialBrackets.prototype.constructor = TokenSpecialBrackets;

	function TokenCloseBrackets()
	{
		this.id = 5;
		this.Unicode = {
			")" : 1,
			"⟫" : 1,
			"⟧" : 1,
		};
		this.LaTeX = {
			"\\end" : "〗",
			"\\rangle" : "⟩",
			"\\ket" : "⟩",
			"\\rbrace" : "}",
			"\\rbrack" : "]",
			"\\rceil" : "⌉",
			"\\rfloor" : "⌋",
			"\\Rbrack" : "⟧",
			"\\right" : "\\right",
			"}" : "}",
			")" : ")",
			"⟩" : "⟩",
			"]" : "]",
			"⌉" : "⌉",
			"⌋" : "⌋",
			"⟧" : "⟧",
		};
		this.Init();
	}
	TokenCloseBrackets.prototype = Object.create(LexerLiterals.prototype);
	TokenCloseBrackets.prototype.constructor = TokenCloseBrackets;
	TokenCloseBrackets.prototype.IsSimple = function (str)
	{
		return str === ")" ||
			str === "]" ||
			str === "}" ||
			str === "|"
	}

	function TokenOpenCloseBrackets()
	{
		this.id = 6;
		this.Unicode = {};
		this.LaTeX = {
			"\\norm" : "‖",
			"\\Vert" : "‖",
			"\\vert" : "|",
			"‖"	:	"‖",
			"|"	:	"|",
		};
		this.Init();
	}
	TokenOpenCloseBrackets.prototype = Object.create(LexerLiterals.prototype);
	TokenOpenCloseBrackets.prototype.constructor = TokenOpenCloseBrackets;

	function TokenPhantom()
	{
		this.id = 7;
		this.LaTeX = {
			"\\asmash" : "⬆",
			"\\dsmash" : "⬇",
			"\\hphantom" : "⬄",
			"\\hsmash" : "⬌",
			"\\phantom" : "⟡",
			"\\smash" : "⬍",
			"\\vphantom" : "⇳",
		};
		this.Unicode = {};
		this.Init();
	}
	TokenPhantom.prototype = Object.create(LexerLiterals.prototype);
	TokenPhantom.prototype.constructor = TokenPhantom;

	function TokenHorizontalStretch()
	{
		this.id = 8;
		this.LaTeX = {
			"\\dashv" : "⊣",
			"\\gets" : "←",
			"\\hookleftarrow" : "↩",
			"\\hookrightarrow" : "↪",
			"\\Leftarrow" : "⇐",
			"\\leftarrow" : "←",
			"\\leftharpoondown" : "↽",
			"\\leftharpoonup" : "↼",
			"\\Leftrightarrow" : "⇔",
			"\\leftrightarrow" : "↔",
			"\\mapsto" : "↦",
			"\\models" : "⊨",
			"\\Rightarrow" : "⇒",
			"\\rightarrow" : "→",
			"\\rightharpoondown" : "⇁",
			"\\rightharpoonup" : "⇀",
			"\\to" : "→",
			"\\vdash" : "⊢",
		};
		this.Unicode = {};
		this.Init();
	}
	TokenHorizontalStretch.prototype = Object.create(LexerLiterals.prototype);
	TokenHorizontalStretch.prototype.constructor = TokenHorizontalStretch;

	function TokenDivide()
	{
		this.id = 11;
		this.LaTeX = {
			"\\atop" : "¦",
			"\\ndiv" : "⊘",
			"\\over" : "/",
			"\\sdiv" : "⁄",
			"\\ldiv" : "∕",
			"\\ldivide" : "∕",
		};
		this.Unicode = {};
		this.Init();
	}
	TokenDivide.prototype = Object.create(LexerLiterals.prototype);
	TokenDivide.prototype.constructor = TokenDivide;

	function TokenEqArray()
	{
		this.id = 12;
		this.LaTeX = {
			"\\substack" : "\\substack",
		};
		this.Unicode = {};
		this.Init();
	}
	TokenEqArray.prototype = Object.create(LexerLiterals.prototype);
	TokenEqArray.prototype.constructor = TokenEqArray;

	function TokenMarker()
	{
		this.id = 13;
		this.LaTeX = {
			"\\eqno" : "#",
		};
		this.Unicode = {};
		this.Init();
	}
	TokenMarker.prototype = Object.create(LexerLiterals.prototype);
	TokenMarker.prototype.constructor = TokenMarker;

	function TokenSubSup()
	{
		this.id = 14;
		this.LaTeX = {
			"\\above" : "┴",
			"\\below" : "┬",
			"\\pppprime" : "⁗",
			"\\ppprime" : "‴",
			"\\pprime" : "″",
			"\\prime" : "′",
			"_" : "_",
			"^" : "^",
		};
		this.Unicode = {
			"_" : 1,
			"^" : 1,
		};
		this.Init();
	}
	TokenSubSup.prototype = Object.create(LexerLiterals.prototype);
	TokenSubSup.prototype.constructor = TokenSubSup;

	function TokenNary()
	{
		this.id = 15;
		this.Unicode = {
		};
		this.LaTeX = {
			"\\amalg" : "∐",
			"\\aoint": "∳",
			"\\bigcap" : "⋂",
			"\\bigcup" : "⋃",
			"\\bigodot" : "⨀",
			"\\bigoplus" : "⨁",
			"\\bigotimes" : "⨂",
			"\\bigsqcup" : "⨆",
			"\\biguplus" : "⨄",
			"\\bigvee" : "⋁",
			"\\bigwedge" : "⋀",
			"\\coint" : "∲",
			"\\iiiint" : "⨌",
			"\\iiint" : "∭",
			"\\iint" : "∬",
			"\\int" : "∫",
			"\\oiiint" : "∰",
			"\\oiint" : "∯",
			"\\oint" : "∮",
			"\\prod" : "∏",
			"\\sum" : "∑",
		};
		this.Init();
	}
	TokenNary.prototype = Object.create(LexerLiterals.prototype);
	TokenNary.prototype.constructor = TokenNary;

	function TokenRadical()
	{
		this.id = 16;
		this.Unicode = {};
		this.LaTeX = {
			"\\cbrt" : "∛",
			"\\qdrt" : "∜",
			"\\sqrt" : "√",

			"\\root" : "⒭",
		};
		this.Init();
	}
	TokenRadical.prototype = Object.create(LexerLiterals.prototype);
	TokenRadical.prototype.constructor = TokenRadical;

	function TokenRrect()
	{
		this.id = 17;
		this.Unicode = {};
		this.LaTeX = {
			"\\rrect" : "▢",
		};
		this.Init();
	}
	TokenRrect.prototype = Object.create(LexerLiterals.prototype);
	TokenRrect.prototype.constructor = TokenRrect;

	function TokenDelimiter()
	{
		this.id = 18;
		this.Unicode = {};
		this.LaTeX = {
			"\\mid" : "∣",
			"\\vbar" : "│",

		};
		this.Init();
	}
	TokenDelimiter.prototype = Object.create(LexerLiterals.prototype);
	TokenDelimiter.prototype.constructor = TokenDelimiter;

	function TokenAccent()
	{
		this.id = 19;
		this.name = "AccentLiterals";
		this.LaTeX = {
			"\\hat": "̂",
			"\\widehat": "̂",
			"\\check": "̌",
			"\\tilde": "̃",
			"\\widetilde": "～",
			"\\acute": "́",
			"\\grave": "̀",
			"\\dot": "̇",
			"\\ddddot" : "⃜",
			"\\ddot": "̈",
			"\\dddot": "⃛",
			"\\breve": "̆",
			"\\bar": "̅",
			"\\Bar": "̿",
			"\\vec": "⃗",
			"\\lhvec" : "⃐",
			"\\hvec" : "⃑",
			"\\tvec" : "⃡",
			"\\lvec" : "⃖",
		};
		this.Unicode = {};

		this.Init();
	}
	TokenAccent.prototype = Object.create(LexerLiterals.prototype);
	TokenAccent.prototype.IsUnicodeToken = function (str)
	{
		if (!str || !str[0])
			return;

		let strFirstSymbol = str[0];

		let code = strFirstSymbol.charCodeAt(0);
		const isFirstBlocks = function (code) {
			return code >= 768 && code <= 879
		}
		const isSecondBlocks = function (code) {
			return code >= 8400 && code <= 8447
		}

		if (isFirstBlocks(code) || isSecondBlocks(code))
			return strFirstSymbol;
	};

	function TokenBox()
	{
		this.id = 20;
		this.Unicode = {};
		this.LaTeX = {
			"\\box" : "□"
		};
		this.Init();
	}
	TokenBox.prototype = Object.create(LexerLiterals.prototype);
	TokenBox.prototype.constructor = TokenBox;

	function TokenMatrix()
	{
		this.id = 21;
		this.data = ["⒩", "■"];
		this.Unicode = {};
		this.LaTeX = {
			"\\matrix" : "■",
			"\\eqarray" : "█",

			"\\begin{cases}" : 1,
			"\\begin{pmatrix}" : 1,
			"\\begin{matrix}" : 1,
			"\\begin{bmatrix}" : 1,
			"\\begin{Bmatrix}" : 1,
			"\\begin{vmatrix}" : 1,
			"\\begin{Vmatrix}" : 1,
			"\\begin{array}" : 1,
			"\\begin{equation}" : 1,

			"\\end{cases}" : 2,
			"\\end{pmatrix}" : 2,
			"\\end{matrix}" : 2,
			"\\end{bmatrix}" : 2,
			"\\end{Bmatrix}" : 2,
			"\\end{vmatrix}" : 2,
			"\\end{Vmatrix}" : 2,
			"\\end{array}" : 2,
			"\\end{equation}" : 2,
		};
		this.Init();
	}
	TokenMatrix.prototype = Object.create(LexerLiterals.prototype);
	TokenMatrix.prototype.constructor = TokenMatrix;
	TokenMatrix.prototype.isUseLaTeXBrackets = true;

	function TokenRect()
	{
		this.id = 22;
		this.Unicode = {};
		this.LaTeX = {
			"\\rect" : "▭",
		};
		this.Init();
	}
	TokenRect.prototype = Object.create(LexerLiterals.prototype);
	TokenRect.prototype.constructor = TokenRect;

	function TokenSpace()
	{
		this.id = 23;
		this.Unicode = {
			"  " 	: 	1,			// 2/18em space  very thin math space
			"  "	:	1,			// 7/18em space  very very thick math space
			" "			:	1,
			"\t"		:	1,
			"\n"		:	1,
			" "		:	1,
			"‌"		:	1,
		};
		this.LaTeX = {
			"\\nbsp"	:	" ",		// space width && no-break space
			"\\numsp"	:	" ",		// digit width
			"\\emsp"	:	" ",		// 18/18 em
			"\\ensp"	:	" ",		// 9/18 em
			"\\vthicksp":	" ",	// 6/18 em verythickmathspace
			"\\thicksp"	:	" ",	// 5/18 em thickmathspace
			"\\medsp"	:	" ",		// 4/18 em mediummathspace
			"\\thinsp"	:	" ",		// 3/18 em thinmathspace
			"\\hairsp"	:	" ",		// 3/18 em veryverythinmathspace
			"\\zwsp"	: 	"​",
			"\\zwnj"	: 	"‌",
		};
		this.Init();
	}
	TokenSpace.prototype = Object.create(LexerLiterals.prototype);
	TokenSpace.prototype.constructor = TokenSpace;

	function TokenLaTeXWords()
	{
		this.id = 24;
		this.isClassEqalData = true;
	}
	TokenLaTeXWords.prototype = Object.create(LexerLiterals.prototype);
	TokenLaTeXWords.prototype.constructor = TokenLaTeXWords;
	TokenLaTeXWords.prototype.SearchForLaTeXToken = function (arrStr)
	{
		return this.private_GetLaTeXWord(arrStr);
	};

	function TokenFunctionLiteral()
	{
		this.id = 25;
	}
	TokenFunctionLiteral.prototype = Object.create(LexerLiterals.prototype);
	TokenFunctionLiteral.prototype.constructor = TokenFunctionLiteral;
	TokenFunctionLiteral.prototype.IsLaTeXInclude = function (str)
	{
		if (MathAutoCorrectionFuncNames.includes(str.slice(1)) || limitFunctions.includes(str.slice(1)) || (str.length > 1 && str[0] === "\\"))
			return str;
	};
	TokenFunctionLiteral.prototype.IsLaTeXIncludeNormal = function (str)
	{
		if (MathAutoCorrectionFuncNames.includes(str.slice(1)) || (str.length > 1 && str[0] === "\\"))
			return str;
	}
	TokenFunctionLiteral.prototype.IsLaTeXIncludeLimit = function (str)
	{
		if (limitFunctions.includes(str.slice(1)))
			return str;
	}
	TokenFunctionLiteral.prototype.IsUnicodeInclude = function(arrStr)
	{
		if (!arrStr)
			return;

		if (!Array.isArray(arrStr))
			arrStr = [arrStr];

		let strLast = "";

		function IsFunc(arrStr)
		{
			let str = arrStr.join("");
			if (oMathAutoCorrection.arrFunctionsNames[str] === null)
				strLast = str;
		}

		for (let i = 1; i <= arrStr.length; i++)
		{
			IsFunc(arrStr.slice(0, i), i);
		}

		if (strLast)
			return strLast;
	};
	TokenFunctionLiteral.prototype.GetUnicodeToken = function (arrStr)
	{
		return this.IsUnicodeInclude(arrStr);
	};

	function TokenSpecialLiteral()
	{
		this.id = 26;
		this.Unicode = {
			// "┬" : 1,
			// "┴" : 1,
			"&" : 1,
			"@" : 1,
		};
		this.LaTeX = {
			// "\\cases" : "Ⓒ",
			// "\\j" : "Jay",
		}

	}
	TokenSpecialLiteral.prototype = Object.create(LexerLiterals.prototype);
	TokenSpecialLiteral.prototype.constructor = TokenSpecialLiteral;

	function TokenOther()
	{
		this.id = 27;
		this.Unicode = {};
		this.LaTeX = {};
		this.Init();
	}
	TokenOther.prototype = Object.create(LexerLiterals.prototype);
	TokenOther.prototype.constructor = TokenOther;
	TokenOther.prototype.GetUnicodeToken = function(arrStr)
	{
		let intCode = GetFixedCharCodeAt(arrStr[0]);
		if (intCode >= 0x1D400 && intCode <= 0x1D7FF)
			return arrStr[0];
	};

	function TokenHorizontalBrackets()
	{
		this.id = 28;
		this.LaTeX = {
			"\\overparen": "⏜",
			"\\underparen": "⏝",
			"\\overbrace": "⏞",
			"\\overline": "¯",
			"\\underbar" : "▁",
			"\\underbrace": "⏟",
			"\\overshell": "⏠",
			"\\undershell": "⏡",
			"\\overbracket": "⎴",
			"\\underbracket": "⎵",
		};
		this.Unicode = {};
		this.Init();
	}
	TokenHorizontalBrackets.prototype = Object.create(LexerLiterals.prototype);
	TokenHorizontalBrackets.prototype.constructor = TokenHorizontalBrackets;
	TokenHorizontalBrackets.prototype.GetPos = function (str)
	{
		switch (str) {
			case "⏜": return VJUST_TOP;
			case "⏝": return VJUST_BOT;
			case "⏞": return VJUST_TOP;
			case "⏟": return VJUST_BOT;
			case "⏠": return VJUST_TOP;
			case "⏡": return VJUST_BOT;
			case "⎴": return VJUST_BOT;
			case "⎵": return VJUST_TOP;
			case "▁": return VJUST_BOT;
			case "¯": return VJUST_TOP;
		}
	}

	function TokenInvisibleOperators()
	{
		this.id = 29;
		this.Unicode = {};
		this.LaTeX = {
			"\\funcapply" : "⁡",  // Invisible function application
			"\\itimes" : "⁢",
		};
		this.Init();
	}
	TokenInvisibleOperators.prototype = Object.create(LexerLiterals.prototype);
	TokenInvisibleOperators.prototype.constructor = TokenInvisibleOperators;

	function TokenAlphanumeric()
	{
		this.id = 30;
		this.Unicode = {};
		this.LaTeX = {};
		this.Init();
	}
	TokenAlphanumeric.prototype = Object.create(LexerLiterals.prototype);
	TokenAlphanumeric.prototype.constructor = TokenAlphanumeric;
	TokenAlphanumeric.prototype.GetUnicodeToken = function(arrStr)
	{
		let intCode = GetFixedCharCodeAt(arrStr[0]);
		if (intCode >= 0x1D400 && intCode <= 0x1D7FF || intCode >= 0x2102 && intCode <= 0x2134)
			return arrStr[0];
	};

	function TokenFont()
	{
		this.id = 31;
		this.Unicode = {};
		this.LaTeX = {};
		this.Init();
	}
	TokenFont.prototype = Object.create(LexerLiterals.prototype);
	TokenFont.prototype.constructor = TokenFont;
	TokenFont.prototype.GetTypes = function ()
	{
		return {
			"\\sf": 3,
			"\\script": 7,
			"\\scr": 7,
			"\\rm": -1,
			"\\oldstyle": 7,
			"\\mathtt": 11,
			"\\mathsfit": 5,
			"\\mathsfbfit": 6,
			"\\mathsfbf": 4,
			"\\mathsf": 3,
			"\\mathrm": -1,
			"\\mathit": 1,
			"\\mathfrak": 9,
			"\\mathcal": 7,
			"\\mathbfit": 2,
			"\\mathbffrak": 10,
			"\\mathbfcal": 8,
			"\\mathbf": 0,
			"\\mathbb": 12,
			"\\it": 1,
			"\\fraktur": 9,
			"\\frak": 9,
			"\\double": 12,
		}
	};

	function TokenOf()
	{
		this.id = 32;
		this.Unicode = {};
		this.LaTeX = {
			"\\of" : "▒",
			"\\naryand" : "▒",
		};
		this.Init();
	}
	TokenOf.prototype = Object.create(LexerLiterals.prototype)
	TokenOf.prototype.constructor = TokenOf;

	function TokenArrayMatrix()
	{
		this.id = 33;
		this.Unicode = {};
		this.LaTeX = {
			"\\\\" : "\\\\",
		};
		this.Init();
	}
	TokenArrayMatrix.prototype = Object.create(LexerLiterals.prototype)
	TokenArrayMatrix.prototype.constructor = TokenArrayMatrix;
	TokenArrayMatrix.prototype.GetLaTeXToken = function (str)
	{
		if (str[0] === "\\" && str[1] === "\\")
			return "\\\\"
	}

	//---------------------------------------Initialize data for Tokenizer----------------------------------------------

	// List of tokens types for parsers processing
	const MathLiterals = {
		subSup:			new TokenSubSup(),
		func:			new TokenFunctionLiteral(),
		LaTeX:			new TokenLaTeXWords(),
		alphanumeric:	new TokenAlphanumeric(),
		other:			new TokenOther(),
		radical: 		new TokenRadical(),
		nary: 			new TokenNary(),
		invisible:		new TokenInvisibleOperators(),
		divide: 		new TokenDivide(),
		hbrack:			new TokenHorizontalBrackets(),
		specialBrac:	new TokenSpecialBrackets(),
		lrBrackets: 	new TokenOpenCloseBrackets(),
		rBrackets: 		new TokenCloseBrackets(),
		lBrackets: 		new TokenOpenBrackets(),
		operand:		new TokenOperand(),
		operator:		new TokenOperators(),
		space: 			new TokenSpace(),
		accent: 		new TokenAccent(),
		number:			new TokenNumbers(),
		special:		new TokenSpecialLiteral(),
		box: 			new TokenBox(),
		rect:			new TokenRect(),
		matrix: 		new TokenMatrix(),
		font:			new TokenFont(),
		of:				new TokenOf(),
		delimiter:		new TokenDelimiter(),
		char:			new TokenChars(),
		horizontal: 	new TokenHorizontalStretch(),
		arrayMatrix:	new TokenArrayMatrix(),
		eqArray:		new TokenEqArray(),
	};

	// The array defines the sequence in which the tokens are checked by the lexer
	// the higher an element is, the lower its priority
	const arrTokensCheckerList = [
		MathLiterals.char,
		MathLiterals.func,
		MathLiterals.LaTeX,
		MathLiterals.delimiter,
		MathLiterals.special,
		MathLiterals.of,
		MathLiterals.number,
		MathLiterals.accent,
		MathLiterals.space,
		MathLiterals.operator,
		MathLiterals.rect,
		MathLiterals.operand,
		MathLiterals.lBrackets,
		MathLiterals.rBrackets,
		MathLiterals.lrBrackets,
		MathLiterals.specialBrac,
		MathLiterals.hbrack,
		MathLiterals.divide,
		MathLiterals.invisible,
		MathLiterals.horizontal,
		MathLiterals.matrix,
		MathLiterals.nary,
		MathLiterals.radical,
		MathLiterals.other,
		MathLiterals.alphanumeric,
		MathLiterals.subSup,
		MathLiterals.arrayMatrix,
		MathLiterals.eqArray,
	];

	//-------------------------------------Generating AutoCorrection Rules----------------------------------------------

	// Special autocorrection elements (doesn't start with //)
	const SpecialAutoCorrection = {
		"!!" : "‼",
		"...": "…",
		"::" : "∷",
		":=" : "≔",

		"~=" : "≅",
		"+-" : "±",
		"-+" : "∓",
		"<<" : "≪",
		"<=" : "≤",
		"->" : "→",
		">=" : "≥",
		">>" : "≫",
		"/<" : "≮",
		"/=" : "≠",
	};
	const MathAutoCorrectionLong = {
		"\\binomial" : "(a+b)^n=∑_(k=0)^n ▒(n¦k)a^k b^(n-k)",
		"\\integral": "1/2π ∫_0^2π ▒ⅆθ/(a+b sin θ)=1/√(a^2-b^2)",
		"\\identitymatrix": "(■(1&0&0@0&1&0@0&0&1))",
		"\\break": "⤶",
		"\\limit" : "lim_(n→∞)⁡〖(1+1/n)^n〗=e",
	}

	// Generate autocorrection rules
	function MathAutoCorrectionList() // oMathAutoCorrection
	{
		this.arrRuleList = [];
		this.oGeneralRules = {};
		this.oSpecialList = {};
		this.arrFunctionsNames = {
			'arcsin'	: null,
			'asin' 		: null,
			'sin'		: null,
			'arcsinh'	: null,
			'asinh'		: null,
			'sinh'		: null,
			'arcsec'	: null,
			'sec'		: null,
			'asec' 		: null,
			'arcsech'	: null,
			'asech'		: null,
			'sech'		: null,
			'arccos'	: null,
			'acos'		: null,
			'cos'		: null,
			'arccosh'	: null,
			'acosh'		: null,
			'cosh'		: null,
			'arccsc'	: null,
			'acsc'		: null,
			'csc'		: null,
			'arccsch'	: null,
			'acsch'		: null,
			'csch'		: null,
			'arctan'	: null,
			'atan'		: null,
			'tan' 		: null,
			'arctanh'	: null,
			'atanh'		: null,
			'tanh'		: null,
			'arccot'	: null,
			'acot' 		: null,
			'cot'		: null,
			'arccoth'	: null,
			'acoth'		: null,
			'coth'		: null,
			'arg' 		: null,
			'det'		: null,
			'exp'		: null,
			'inf'		: null,
			'lim'		: null,
			'min'		: null,
			'def'		: null,
			'dim'		: null,
			'gcd'		: null,
			'log'		: null,
			'Pr'		: null,
			'deg'		: null,
			'erf'		: null,
 			'lg'		: null,
			'ln'		: null,
			'max'		: null,
			'sup'		: null,
			"ker"		: null,
			'hom'		: null,
			'sgn'		: null,
		};
		this.GenerateTokens();

		return this;
	}
	MathAutoCorrectionList.prototype.AddObjectToGeneral = function (oObj)
	{
		this.oGeneralRules = Object.assign(this.oGeneralRules, oObj);
	};
	MathAutoCorrectionList.prototype.AddData = function (name, data)
	{
		this.oGeneralRules[name] = data;
	};
	MathAutoCorrectionList.prototype.GenerateTokens = function ()
	{
		this.GenerateAutoCorrectionList();
		this.GenerateTokensByFont();
		this.GenerateSpecialRules();
		this.GenerateCustomRules();

		const CheckSort = function (a,b)
		{
			if (a[0] < b[0])
			{
				return -1;
			}
			else if (a[0] > b[0])
			{
				return 1;
			}

			return 0;
		};
		const IsSpecialRule = function(rule)
		{
			return rule[0][0] !== "\\";
		}

		for (let i = 0; i < this.arrRuleList.length; i++)
		{
			let arrCurrentRule = this.arrRuleList[i];
			if (IsSpecialRule(arrCurrentRule))
			{
				let strName = arrCurrentRule[0];
				let oData 	= arrCurrentRule[1];
				this.oSpecialList[strName] = oData;
			}
		}

		console.log(this.oSpecialList);

		this.arrRuleList.sort(CheckSort);
	};
	MathAutoCorrectionList.prototype.GenerateTokensByFont = function ()
	{
		let arr_Literals = [
			'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z',
			'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
			'0', '1', '2', '3', '4', '5', '6', '7', '8', '9'
		];

		let oFontTypes = MathLiterals.font.GetTypes();
		let name = Object.keys(oFontTypes);

		for (let j = 0; j < name.length; j++)
		{
			let indexOfFont = oFontTypes[name[j]];
			for (let i = 0; i < arr_Literals.length; i++)
			{
				let Current = GetMathFontChar[arr_Literals[i]];
				if (Current[indexOfFont])
				{
					let strName = name[j] + arr_Literals[i];
					let intData = Current[indexOfFont].codePointsArray([]);
					let arrCorrectionRule = [strName, intData];
					this.AddData(strName, intData);
					this.arrRuleList.push(arrCorrectionRule);
				}
			}
		}
	};
	MathAutoCorrectionList.prototype.GenerateSpecialRules = function ()
	{
		let name = Object.keys(MathAutoCorrectionLong);

		for (let j = 0; j < name.length; j++)
		{
			let strName = name[j];
			let strData = AscCommon.convertUTF16toUnicode(MathAutoCorrectionLong[strName]);
			let arrAutoCorrectionRule = [strName, strData];

			this.arrRuleList.push(arrAutoCorrectionRule);
		}
	};
	MathAutoCorrectionList.prototype.GenerateAutoCorrectionList = function ()
	{
		let names = Object.keys(MathLiterals);

		for (let i = 0; i < names.length; i++)
		{
			let name = names[i];
			let oData = MathLiterals[name].LaTeX;

			if (oData)
			{
				let innerNames = Object.keys(oData);
				for (let i = 0; i < innerNames.length; i++)
				{
					let strName = innerNames[i];
					let strData = AscCommon.convertUTF16toUnicode(oData[strName]);

					if (strData)
					{
						let arrRule = [strName, strData]
						this.AddData(strName, strData);
						this.arrRuleList.push(arrRule);
					}
				}
			}
		}
	};
	MathAutoCorrectionList.prototype.GenerateCustomRules = function ()
	{
		let name = Object.keys(SpecialAutoCorrection);

		for (let j = 0; j < name.length; j++)
		{
			let strName = name[j];
			let strData = AscCommon.convertUTF16toUnicode(SpecialAutoCorrection[strName]);
			let arrAutoCorrectionRule = [strName, strData];

			this.AddData(strName, strData);
			this.arrRuleList.push(arrAutoCorrectionRule);
		}
	};

	// Array defining which words cannot be corrected during LaTeX processing
	const arrDoNotConvertWordsForLaTeX = [
		"\\left", "\\right",
		"\\array",
		"\\begin", "\\end",
		"\\matrix",
		"\\below", "\\above",
		"\\box", "\\fbox", "\\rect",

		"\\sum", "\\prod", "\\amalg", "\\coprod", "\\bigwedge",
		"\\bigvee", "\\bigcup", "\\bigcap", "\\bigsqcup", "\\biguplus",
		"\\bigodot", "\\bigoplus", "\\bigotimes",
		"\\int", "\\iint", "\\iiint", "\\iiiint", "\\oint", "\\oiint",
		"\\oiiint", "\\coint", "\\aouint",
	];
	let oMathAutoCorrection = new MathAutoCorrectionList();
	// Default list of autocorrection elements
	let AutoCorrectionList = oMathAutoCorrection.arrRuleList;
	// Array with function names for autocorrection
	const MathAutoCorrectionFuncNames = [
		'arcsin', 'asin', 'sin', 'arcsinh', 'asinh', 'sinh',
		'arcsec', 'sec', 'asec', 'arcsech', 'asech','sech',
		'arccos', 'acos', 'cos', 'arccosh','acosh', 'cosh',
		'arccsc', 'acsc', 'csc', 'arccsch', 'acsch', 'csch',
		'arctan', 'atan', 'tan', 'arctanh', 'atanh', 'tanh',
		'arccot', 'acot', 'cot', 'arccoth', 'acoth', 'coth',
		'arg', 'det', 'exp', 'inf', 'lim', 'min',
		'def', 'dim', 'gcd', 'ker', 'log', 'Pr',
		'deg', 'erf', 'hom', 'lg', 'ln', 'max', 'sup', "ker",
		'hom', 'sgn',
	];
	// List of structures types that generate parsers
	const MathStructures = {
		char:	0,
		space:	1,
		number: 2,
		other:	3,
		frac: 	5,
		bracket_block: 6,
		minus: 	7,
		plain: 	8,
		bar:	9,
		nary:	10,
		box:	11,
		rect:	12,
		radical:13,
		func: 	14,
		pre_script: 15,
		sub_sub: 16,
		func_lim: 18,
		limit: 19,
		diacritic_base: 20,
		matrix: 21,
		accent: 22,
		group_character: 23,
		horizontal: 24,
		array: 25,
	};

	const oNamesOfLiterals = {
		fractionLiteral: 			[0, "FractionLiteral"],
		spaceLiteral: 				[1, "SpaceLiteral", MathLiterals.space],
		charLiteral: 				[2, "CharLiteral"],
		operatorLiteral: 			[5, "OperatorLiteral"],
		binomLiteral: 				[6, "BinomLiteral"],
		bracketBlockLiteral: 		[7, "BracketBlock"],
		functionLiteral: 			[8, "FunctionLiteral"],
		subSupLiteral: 				[9, "SubSupLiteral"],
		sqrtLiteral: 				[10, "SqrtLiteral"],
		numberLiteral: 				[11, "NumberLiteral"],
		mathOperatorLiteral: 		[12, "MathOperatorLiteral"],
		rectLiteral: 				[13, "RectLiteral"],
		boxLiteral: 				[14, "BoxLiteral"],
		borderBoxLiteral:			[58, "BorderBoxLiteral"],
		preScriptLiteral: 			[15, "PreScriptLiteral"],
		mathFontLiteral: 			[16, "MathFontLiteral"],
		overLiteral: 				[17, "OverLiteral"],
		diacriticLiteral: 			[18, "DiacriticLiteral"],
		diacriticBaseLiteral: 		[19, "DiacriticBaseLiteral"],
		otherLiteral: 				[20, "OtherLiteral"],
		anMathLiteral: 				[21, "AnMathLiteral"],
		opBuildupLiteral: 			[22, "opBuildUpLiteral"],
		opOpenBracket: 				[23, "opOpenLiteral"],
		opCloseBracket: 			[24, "opCLoseLiteral"],
		opOpenCloseBracket: 		[25, "opCloseLiteral"],
		hBracketLiteral: 			[28, "hBracketLiteral"],
		opNaryLiteral: 				[29, "opNaryLiteral"],
		asciiLiteral: 				[30, "asciiLiteral"],
		opArrayLiteral: 			[31, "opArrayLiteral"],
		opDecimal: 					[32, "opDecimal"],

		specialScriptNumberLiteral: [33, "specialScriptLiteral"],
		specialScriptCharLiteral: 	[34, "specialScriptLiteral"],
		specialScriptBracketLiteral: [35, "specialScriptBracketLiteral"],
		specialScriptOperatorLiteral: [36, "specialScriptBracketLiteral"],

		specialIndexNumberLiteral: 	[37, "specialScriptLiteral"],
		specialIndexCharLiteral: 	[38, "specialScriptLiteral"],
		specialIndexBracketLiteral: [39, "specialScriptBracketLiteral"],
		specialIndexOperatorLiteral: [40, "specialScriptBracketLiteral"],

		textPlainLiteral: 				[41, "textPlainLiteral"],
		nthrtLiteral: 				[42, "nthrtLiteral"],
		fourthrtLiteral: 			[43, "fourthrtLiteral"],
		cubertLiteral: 				[44, "cubertLiteral"],
		overBarLiteral: 			[45, "overBarLiteral"],

		factorialLiteral: 			[46, "factorialLiteral"],
		rowLiteral: 				[47, "rowLiteral"],
		rowsLiteral: 				[48, "rowsLiteral"],

		minusLiteral: 				[49, "minusLiteral"],
		LaTeXLiteral: 				[50, "LaTeXLiteral"],

		functionWithLimitLiteral: 	[51, "functionWithLimitLiteral"],
		functionNameLiteral: 		[52, "functionNameLiteral"],
		matrixLiteral: 				[53, "matrixLiteral"],
		arrayLiteral: 				[53, "arrayLiteral"],

		skewedFractionLiteral: 		[54, "skewedFractionLiteral"],
		EqArrayliteral: 			[55, "EqArrayliteral"],

		groupLiteral:				[56, "GroupLiteral"],
		belowAboveLiteral:			[57, "BelowAboveLiteral"],

	};

	function AddFunctionAutoCorrection(str)
	{
		if (undefined === str || null === str)
			return;

		MathAutoCorrectionFuncNames.push(str);
	}
	function SearchFunctionName(str)
	{
		if (undefined === str || null === str)
			return false;

		return MathAutoCorrectionFuncNames.includes(str);
	}

	//---------------------------------------------Tokenizer section----------------------------------------------------
	function Tokenizer(isLaTeX)
	{
		this._string = [];
		this._styles = {};
		this._cursor = 0;

		this.state = [];
		this.isLaTeX = isLaTeX;
	}
	Tokenizer.prototype.Init = function (string)
	{
		if (string instanceof MathTextAndStyles)
		{
			let arrContent = [];
			let oAdditionalData = {};
			let nLength = 0;

			function Proceed (oMathTextAndStyles, context)
			{
				let arr = oMathTextAndStyles.arr;

				for (let i = 0; i < arr.length; i++)
				{
					let oCurrentElement 	= arr[i];
					if (oCurrentElement instanceof MathTextAndStyles)
					{
						Proceed(oCurrentElement, context);
					}
					else if (oCurrentElement instanceof MathText)
					{
						let oLocalAdditionalData	= oCurrentElement.GetAdditionalData();
						let strCurrent 			= oCurrentElement.GetText();
						let tempLength 			= context.GetStringLength(strCurrent);
						let arrCurrent 			= context.GetSymbols(strCurrent);
						arrContent 				= arrContent.concat(arrCurrent);

						for (let j = nLength; j < tempLength + nLength; j++)
						{
							oAdditionalData[j] = oLocalAdditionalData;
						}
						nLength += tempLength;
					}
					else
					{
						let arrCurrent 			= context.GetSymbols(oCurrentElement);
						nLength 				+= arrCurrent.length;
						arrContent 				= arrContent.concat(arrCurrent);
						oAdditionalData[nLength] = undefined;
					}
				}
			}

			Proceed(string, this)

			this._string = arrContent;
			this._styles = oAdditionalData;
		}
		else
		{
			let arrContent 		= this.GetSymbols(string);
			this._string 		= arrContent;
		}
	};
	Tokenizer.prototype.GetSymbols = function (str)
	{
		let output = [];
		for (let oIter = str.getUnicodeIterator(); oIter.check(); oIter.next())
		{
			output.push(String.fromCodePoint(oIter.value()));
		}
		return output;
	};
	Tokenizer.prototype.GetStringLength = function (str)
	{
		let intLen = 0;
		for (let oIter = str.getUnicodeIterator(); oIter.check(); oIter.next())
		{
			intLen++;
		}
		return intLen;
	};
	Tokenizer.prototype.IsHasMoreTokens = function ()
	{
		return this._cursor < this._string.length;
	};
	Tokenizer.prototype.GetTextOfToken = function (intIndex, isLaTeX)
	{
		let arrToken = arrTokensCheckerList[intIndex];

		if (typeof arrToken[0] !== "function")
		{
			if (isLaTeX && arrToken[1] !== undefined)
			{
				return arrToken[0];
			}
			else if (!isLaTeX && arrToken[1] !== undefined)
			{
				return arrToken[1];
			}
		}
	};
	Tokenizer.prototype.GetNextToken = function ()
	{
		if (!this.IsHasMoreTokens())
			return {
				class: undefined,
				data: undefined,
			};

		let autoCorrectRule,
			tokenValue,
			tokenClass,
			string = this._string.slice(this._cursor);

		for (let i = arrTokensCheckerList.length - 1; i >= 0; i--)
		{
			autoCorrectRule = arrTokensCheckerList[i];
			tokenValue = this.MatchToken(autoCorrectRule, string);

			if (tokenValue === null)
				continue;

			else if (!Array.isArray(autoCorrectRule))
			{
				tokenClass = (autoCorrectRule.isClassEqalData)
					? tokenValue
					: autoCorrectRule.id;
			}
			else if (autoCorrectRule.length === 1)
			{
				tokenClass = MathLiterals.char.id;
			}
			else if (autoCorrectRule.length === 2)
			{
				tokenClass = (autoCorrectRule[1] === true)
					? autoCorrectRule[0]
					: autoCorrectRule[1];
			}

			let oStyle = this.GetStyle(this._cursor);

			return {
				class: tokenClass,
				data: tokenValue,
				style: oStyle,
			}
		}
	};
	Tokenizer.prototype.GetStyle = function (nCursorPos)
	{
		return this._styles[nCursorPos - 1];
	};
	Tokenizer.prototype.ProcessString = function (str, char)
	{
		let intLenOfRule = 0;

		while (intLenOfRule <= char.length - 1)
		{
			if (char[intLenOfRule] === str[intLenOfRule])
				intLenOfRule++;
			else
				return;
		}
		return char;
	};
	Tokenizer.prototype.MatchToken = function (fMathCheck, arrStr)
	{
		if (undefined === fMathCheck)
			return null;

		let oMatched = fMathCheck.GetToken(this.isLaTeX, arrStr);

		if (oMatched === null || oMatched === undefined)
			return null;

		this._cursor += this.GetStringLength(oMatched);

		if (fMathCheck.IsNeedReturnCorrected_Unicode === true && !this.isLaTeX)
			oMatched = fMathCheck.LaTeX[oMatched];

		return oMatched;
	};
	Tokenizer.prototype.SaveState = function (oLookahead)
	{
		let strClass = oLookahead.class;
		let data = oLookahead.data;

		this.state.push({
			_string: this._string,
			_cursor: this._cursor,
			oLookahead: { class: strClass, data: data},
		})
	};
	Tokenizer.prototype.RestoreState = function ()
	{
		if (this.state.length > 0) {
			let oState = this.state.shift();
			this._cursor = oState._cursor;
			this._string = oState._string;
			return oState.oLookahead;
		}
	};
	Tokenizer.prototype.IsTextContent = function(intClass, intTokenClass)
	{
		return (intClass !== intTokenClass) ||
			(
				intTokenClass !== 0
			&& intTokenClass !== 1
			&& intTokenClass !== 3
			)
	};
	Tokenizer.prototype.IsContentOfOneType = function()
	{
		let intTokenClass = null;
		while (this.IsHasMoreTokens())
		{
			let intClass = this.GetNextToken().class;

			if (intTokenClass === null)
				intTokenClass = intClass;
			else if (intClass === undefined)
				return true;
			else if (this.IsTextContent(intClass, intTokenClass))
				return false;
		}
		return true;
	};

	//-----------------------Functions for convert tokens array in inner math format------------------------------------
	function GetPrForFunction(oIndex)
	{
		let isHide = true;
		if (oIndex)
			isHide = false;

		return {
			degHide: isHide,
		}
	}
	// Convert tokens to math objects
	function ConvertTokens(oTokens, oContext)
	{
		Paragraph = oContext.Paragraph;

		if (typeof oTokens === "object")
		{
			if (oTokens.type === "LaTeXEquation" || oTokens.type === "UnicodeEquation")
			{
				type = oTokens.type === "LaTeXEquation" ? 1 : 0;
				oTokens = oTokens.body;
			}

			if (Array.isArray(oTokens))
			{
				for (let i = 0; i < oTokens.length; i++)
				{
					if (Array.isArray(oTokens[i]))
					{
						let oToken = oTokens[i];

						for (let j = 0; j < oTokens[i].length; j++)
						{
							SelectObject(oToken[j], oContext);
						}
					}
					else
					{
						SelectObject(oTokens[i], oContext);
					}
				}
			}
			else
			{
				SelectObject(oTokens, oContext)
			}
		}
		else
		{
			oContext.Add_Text(oTokens);
		}
	}
	// Find token in all types for convert
	function SelectObject (oTokens, oContext)
	{
		let num = 1; // needs for debugging
		if (oTokens)
		{
			if (oTokens instanceof MathText)
			{
				oContext.Add_Text(oTokens.text, undefined, undefined, oTokens);
				return;
			}
			switch (oTokens.type)
			{
				case undefined:
					for (let i = 0; i < oTokens.length; i++) {
						ConvertTokens(
							oTokens[i],
							oContext,
						);
					}
					break;
				case MathStructures.other:
					let intCharCode = oTokens.value.codePointAt()
					oContext.Add_Symbol(intCharCode);
					break;
				case oNamesOfLiterals.functionNameLiteral[num]:
				case oNamesOfLiterals.specialScriptNumberLiteral[num]:
				case oNamesOfLiterals.specialScriptCharLiteral[num]:
				case oNamesOfLiterals.specialScriptBracketLiteral[num]:
				case oNamesOfLiterals.specialScriptOperatorLiteral[num]:
				case oNamesOfLiterals.specialIndexNumberLiteral[num]:
				case oNamesOfLiterals.specialIndexCharLiteral[num]:
				case oNamesOfLiterals.specialIndexBracketLiteral[num]:
				case oNamesOfLiterals.specialIndexOperatorLiteral[num]:
				case oNamesOfLiterals.opDecimal[num]:
				case MathStructures.char:
				case MathStructures.space:
				case oNamesOfLiterals.mathOperatorLiteral[num]:
				case MathStructures.number:
					if (oTokens.decimal)
					{
						ConvertTokens(
							oTokens.left,
							oContext,
						);
						oContext.Add_Text(oTokens.decimal)
						ConvertTokens(
							oTokens.right,
							oContext,
						);
					}
					else
					{
						if (Array.isArray(oTokens.style))
						{
							for (let nTokenStyle = 0; nTokenStyle < oTokens.style.length; nTokenStyle++) {
								oContext.Add_Text(oTokens.value[nTokenStyle], undefined, undefined, oTokens.style[nTokenStyle]);
							}
						}
						else
						{
							oContext.Add_Text(oTokens.value, undefined, undefined, oTokens.style);
						}
					}
					break;
				case oNamesOfLiterals.textPlainLiteral[num]:
					oContext.Add_Text(oTokens.value, Paragraph, STY_PLAIN);
					break
				case MathStructures.nary:
					let lPr = {
						ctrPrp: oTokens.style.style,
						chr: oTokens.value.charCodeAt(0),
						subHide: true,
						supHide: true,
					}

					let oNary = oContext.Add_NAry(lPr, null, null, null);

					if (oTokens.third) {
						ConvertTokens(
							oTokens.third,
							oNary.getBase(),
						);
					}
					let oBase = oNary.getBase();
					oBase.setCtrPrp(oTokens.thirdStyle.style);
					break;
				case MathStructures.pre_script:
					let oPreSubSup = oContext.Add_Script(
						oTokens.up && oTokens.down,
						{ctrPrp: new CTextPr(), type: DEGREE_PreSubSup},
						null,
						null,
						null
					);
					ConvertTokens(
						oTokens.value,
						oPreSubSup.getBase()
					);
					UnicodeArgument(
						oTokens.up,
						MathStructures.bracket_block,
						oPreSubSup.getUpperIterator()
					)
					UnicodeArgument(
						oTokens.down,
						MathStructures.bracket_block,
						oPreSubSup.getLowerIterator()
					)
					break;
				case MathStructures.accent:
					let oAccent = oContext.Add_Accent(
						oTokens.value.style.style,
						GetFixedCharCodeAt(oTokens.value.data),
						null
					);
					UnicodeArgument(
						oTokens.base,
						MathStructures.bracket_block,
						oAccent.getBase()
					)
					break;
				case oNamesOfLiterals.skewedFractionLiteral[num]:
				case MathStructures.frac:
					if (oTokens.fracType === LITTLE_FRACTION)
					{
						let oBox = new CBox({ctrPrp: new CTextPr()});
						oContext.Add_Element(oBox);
						let BoxMathContent = oBox.getBase();
						BoxMathContent.SetArgSize(-1);
						let oFraction = BoxMathContent.Add_Fraction({ctrPrp: oTokens.style.style, type: BAR_FRACTION}, null, null);
						UnicodeArgument(
							oTokens.up,
							MathStructures.bracket_block,
							oFraction.getNumeratorMathContent()
						);
						UnicodeArgument(
							oTokens.down,
							MathStructures.bracket_block,
							oFraction.getDenominatorMathContent()
						);
					}
					else
					{
						let oFraction = oContext.Add_Fraction(
							{ctrPrp: oTokens.style.style, type: oTokens.fracType},
							null,
							null
						);
						oFraction.SetReviewTypeWithInfo(oTokens.style.reviewData.reviewType, oTokens.style.reviewData.reviewInfo);
						oFraction.ReviewInfo.Update();
						UnicodeArgument(
							oTokens.up,
							MathStructures.bracket_block,
							oFraction.getNumeratorMathContent()
						);
						UnicodeArgument(
							oTokens.down,
							MathStructures.bracket_block,
							oFraction.getDenominatorMathContent()
						);
					}
					break;
				case MathStructures.sub_sub:
					if (oTokens.value && oTokens.value.type === MathStructures.func)
					{
						let oFunc = oContext.Add_Function({ctrPrp: oTokens.value.style.style}, null, null);
						let oFuncName = oFunc.getFName();

						let Pr = (oTokens.up && oTokens.down)
							? {}
							: (oTokens.up)
								? {type: DEGREE_SUPERSCRIPT}
								: {type: DEGREE_SUBSCRIPT}

						let SubSup = oFuncName.Add_Script(
							oTokens.up && oTokens.down,
							Pr,
							null,
							null,
							null
						);

						oTokens.value.type = 0;
						UnicodeArgument(
							oTokens.value,
							MathStructures.bracket_block,
							SubSup.getBase()
						);

						if (oTokens.up) {
							UnicodeArgument(
								oTokens.up,
								MathStructures.bracket_block,
								SubSup.getUpperIterator()
							)
						}
						if (oTokens.down) {
							UnicodeArgument(
								oTokens.down,
								MathStructures.bracket_block,
								SubSup.getLowerIterator()
							)
						}
						if (oTokens.third) {
							let oFuncArgument = oFunc.getArgument();
							UnicodeArgument(
								oTokens.third,
								MathStructures.bracket_block,
								oFuncArgument
							)
						}

						// Set styles
						let oUpper = SubSup.getUpperIterator();
						if (oUpper && oTokens.style.supStyle)
							oUpper.CtrPrp.Merge(oTokens.style.supStyle.style);

						let oLower = SubSup.getLowerIterator();
						if (oLower && oTokens.style.subStyle)
							oLower.CtrPrp.Merge(oTokens.style.subStyle.style);
					}
					else if (oTokens.value && oTokens.value.type === MathStructures.func_lim)
					{
						let oFuncWithLimit = oContext.Add_FunctionWithTypeLimit(
							{},
							null,
							null,
							null,
							oTokens.up ? LIMIT_UP : LIMIT_LOW
						);
						oFuncWithLimit
							.getFName()
							.Content[0]
							.getFName()
							.Add_Text(oTokens.value.value, Paragraph, STY_PLAIN);

						let oLimitIterator = oFuncWithLimit
							.getFName()
							.Content[0]
							.getIterator();

						if (oTokens.up || oTokens.down)
						{
							UnicodeArgument(
								oTokens.up === undefined ? oTokens.down : oTokens.up,
								MathStructures.bracket_block,
								oLimitIterator
							)
						}

						UnicodeArgument(
							oTokens.third,
							MathStructures.bracket_block,
							oFuncWithLimit.getArgument()
						)
					}
					else if (oTokens.value && oTokens.value.type === MathStructures.nary)
					{
						let Pr = {
							ctrPrp: oTokens.value.style.style,
							chr: oTokens.value.value.charCodeAt(0),
							subHide: oTokens.down === undefined,
							supHide: oTokens.up === undefined,
						}

						let oNary = oContext.Add_NAry(Pr, null, null, null);

						ConvertTokens(
							oTokens.third,
							oNary.getBase(),
						);

						let oBase = oNary.getBase();

						if (oTokens.style.ofStyle)
							oBase.setCtrPrp(oTokens.style.ofStyle.style);

						UnicodeArgument(
							oTokens.up,
							MathStructures.bracket_block,
							oNary.getSupMathContent()
						);

						let oUp = oNary.getSupMathContent();

						if (oTokens.style.supStyle)
							oUp.setCtrPrp(oTokens.style.supStyle.style);

						UnicodeArgument(
							oTokens.down,
							MathStructures.bracket_block,
							oNary.getSubMathContent()
						)

						let oDown = oNary.getSubMathContent();
						if (oTokens.style.subStyle)
							oDown.setCtrPrp(oTokens.style.subStyle.style);

					}
					else
					{
						let isSubSup = ((Array.isArray(oTokens.up) && oTokens.up.length > 0) || (!Array.isArray(oTokens.up) && oTokens.up !== undefined)) &&
							((Array.isArray(oTokens.down) && oTokens.down.length > 0) || (!Array.isArray(oTokens.down) && oTokens.down !== undefined))

						let oCurrentStyle = oTokens.style.subStyle ? oTokens.style.subStyle : oTokens.style.supStyle;
						let Pr = {};
						Pr.ctrPrp = oCurrentStyle.style;

						if (!isSubSup)
						{
							if (oTokens.up)
								Pr.type = DEGREE_SUPERSCRIPT;
							else if (oTokens.down)
								Pr.type = DEGREE_SUBSCRIPT;
						}

						let SubSup = oContext.Add_Script(
							isSubSup,
							Pr,
							null,
							null,
							null
						);

						SubSup.SetReviewTypeWithInfo(oCurrentStyle.reviewData.reviewType, oCurrentStyle.reviewData.reviewInfo);

						ConvertTokens(
							oTokens.value,
							SubSup.getBase()
						);

						UnicodeArgument(
							oTokens.up,
							MathStructures.bracket_block,
							SubSup.getUpperIterator()
						);

						UnicodeArgument(
							oTokens.down,
							MathStructures.bracket_block,
							SubSup.getLowerIterator()
						);

						// Set styles
						let oUpper = SubSup.getUpperIterator();
						if (oUpper && oTokens.style.subStyle)
						{
							oUpper.CtrPrp.Merge(oTokens.style.subStyle.style);
							if (oTokens.style.subStyle.reviewData.reviewInfo)
								oUpper.SetReviewTypeWithInfo(oTokens.style.subStyle.reviewData.reviewType, oTokens.style.subStyle.reviewData.reviewInfo);
						}

						let oLower = SubSup.getLowerIterator();
						if (oLower && oTokens.style.supStyle)
						{
							oLower.CtrPrp.Merge(oTokens.style.supStyle.style);
							if (oTokens.style.supStyle.reviewData.reviewInfo)
								oLower.SetReviewTypeWithInfo(oTokens.style.supStyle.reviewData.reviewType, oTokens.style.supStyle.reviewData.reviewInfo);
						}
					}
					break;
				case MathStructures.func_lim:
					let MathFunc = new CMathFunc({});
					oContext.Add_Element(MathFunc);

					let FuncName = MathFunc.getFName();

					let Limit = new CLimit({ctrPrp : oTokens.style.style, type : oTokens.down !== undefined ? LIMIT_LOW : LIMIT_UP});
					FuncName.Add_Element(Limit);

					let LimitName = Limit.getFName();

					UnicodeArgument(
						oTokens.value,
						undefined,
						LimitName
					);

					if (oTokens.up || oTokens.down) {
						UnicodeArgument(
							oTokens.up === undefined ? oTokens.down : oTokens.up,
							MathStructures.bracket_block,
							Limit.getIterator()
						)
					}

					if (oTokens.third)
					{
						ConvertTokens(
							oTokens.third,
							MathFunc.getArgument(),
						)
					}

					break;
				case MathStructures.group_character:
					let intBracketPos = !isNaN(oTokens.isBelow)
						? oTokens.isBelow
						: MathLiterals.hbrack.GetPos(oTokens.hBrack.data);

					if (oTokens.hBrack.data === "¯" || oTokens.hBrack.data === "▁")
					{
						let oBar = (oTokens.hBrack.data === "¯")
							? oContext.Add_Bar({ctrPrp : oTokens.style.style, pos : LOCATION_TOP}, null)
							: oContext.Add_Bar({ctrPrp : oTokens.style.style, pos : LOCATION_BOT}, null);

						UnicodeArgument(
							oTokens.value,
							MathStructures.bracket_block,
							oBar.getBase()
						);
					}
					else
					{
						let Pr = (intBracketPos === VJUST_TOP)
							? {ctrPrp : oTokens.hBrack.style.style, pos : VJUST_TOP, vertJc : VJUST_BOT, chr: oTokens.hBrack.data.charCodeAt(0)}
							: {
								ctrPrp : typeof oTokens.hBrack.data === "object" ? oTokens.hBrack.data.style.style : oTokens.hBrack.style.style,
								vertJc : VJUST_TOP,
								chr : typeof oTokens.hBrack.data === "object" ?  oTokens.hBrack.data.data.charCodeAt(0) :  oTokens.hBrack.data.charCodeAt(0)
							};

						let Group = new CGroupCharacter(Pr);
						oContext.Add_Element(Group);

						UnicodeArgument(
							oTokens.value,
							MathStructures.bracket_block,
							Group.getBase(),
						);

						let oBase = Group.getBase();
						oBase.setCtrPrp(oTokens.style.style);
					}

					break;
				case MathStructures.bracket_block:
					let arr = [null]
					let oPr =
						{
							ctrPrp : oTokens.style.startStyle.style,
							column : oTokens.value.length > 0 ? oTokens.value.length : 1,
							begChr : GetBracketCode(oTokens.left),
							endChr : GetBracketCode(oTokens.right),
						};


					let oBracket = oContext.Add_Delimiter(
						oPr,
						oTokens.value.length > 0 ? oTokens.value.length : 1,
						arr
					);

					if (oTokens.value.length >= 0)
					{
						for (let intCount = 0; intCount < oTokens.value.length; intCount++)
						{
							ConvertTokens(
								oTokens.value[intCount],
								oBracket.getElementMathContent(intCount)
							);

							//Last content
							if (intCount === oTokens.value.length - 1 && oTokens.style.endStyle)
							{
								let oCon = oBracket.getElementMathContent(intCount);
								oCon.setCtrPrp(oTokens.style.endStyle.style);
							}

							if (oTokens.style.middle && oTokens.style.middle[intCount - 1])
							{
								let oContent = oBracket.getElementMathContent(intCount - 1);
								oContent.setCtrPrp(oTokens.style.middle[intCount - 1].style);
							}
						}
					}
					else
					{
						ConvertTokens(
							oTokens.value,
							oBracket.getElementMathContent(0)
						);
					}

					break;
				case MathStructures.radical:
					let Pr = GetPrForFunction(oTokens.index);
					Pr.ctrPrp = oTokens.style.style;
					let oRadical = oContext.Add_Radical(
						Pr,
						null,
						null
					);

					oRadical.SetReviewTypeWithInfo(oTokens.style.reviewData.reviewType, oTokens.style.reviewData.reviewInfo);
					UnicodeArgument(
						oTokens.value,
						MathStructures.bracket_block,
						oRadical.getBase()
					)
					ConvertTokens(
						oTokens.index,
						oRadical.getDegree()
					);
					break;
				case MathStructures.func:
					let oFunc = oContext.Add_Function({ctrPrp: oTokens.style.style}, null, null);

					ConvertTokens(
						oTokens.value,
						oFunc.getFName(),
					);

					let oName = oFunc.getFName();
					oName.CtrPrp.SetItalic(false);

					ConvertTokens(
						oTokens.third,
						oFunc.getArgument(),
					)
					break;
				case oNamesOfLiterals.mathFontLiteral[num]:
					ConvertTokens(
						oTokens.value,
						oContext,
					);
					break;
				case MathStructures.matrix:
					let strStartBracket, strEndBracket;

					if (oTokens.strMatrixType)
					{
						if (oTokens.strMatrixType.length === 2)
						{
							strStartBracket = oTokens.strMatrixType[0].charCodeAt(0)
							strEndBracket = oTokens.strMatrixType[1].charCodeAt(0)
						}
						else
						{
							strEndBracket = strStartBracket = oTokens.strMatrixType[0].charCodeAt(0)
						}
					}

					let rows = oTokens.value.length;
					let cols;

					if (oTokens.value[0] && oTokens.value[0].length)
					{
						cols = oTokens.value[0].length;
					}

					if (strEndBracket && strStartBracket)
					{
						let Delimiter = oContext.Add_DelimiterEx(
							new CTextPr(),
							1,
							[null],
							strStartBracket,
							strEndBracket
						);
						oContext = Delimiter.getElementMathContent(0);
					}

					let oMatrix = oContext.Add_Matrix(
						oTokens.style.head.style,
						rows,
						cols,
						false,
						[]
					);

					for (let intRow = 0; intRow < rows; intRow++)
					{
						for (let intCol = 0; intCol < cols; intCol++)
						{
							let oContent = oMatrix.getContentElement(intRow, intCol);
							ConvertTokens(
								oTokens.value[intRow][intCol],
								oContent,
							);

							let oPr = oTokens.style.cols[intRow]
							if (oPr && intCol === cols - 1)
							{
								oContent.setCtrPrp(oPr.style);
								continue;
							}

							let rPr = oTokens.style.rows[intRow];
							if (rPr)
							{
								let cPr = rPr[intCol];
								if (cPr)
									oContent.setCtrPrp(cPr.style);
							}
						}
					}
					break;
				case MathStructures.array:
					let intCountOfRows = oTokens.value.length;

					let arrayPr = {
						ctrPrp: oTokens.style.style,
						row: intCountOfRows
					};

					let oEqArray = oContext.Add_EqArray(arrayPr, null, null);

					for (let i = 0; i < oTokens.value.length; i++)
					{
						let oMathContent = oEqArray.getElementMathContent(i);
						ConvertTokens(
							oTokens.value[i],
							oMathContent,
						);
					}
					break;
				case MathStructures.box:
					let oBox = oContext.Add_Box({ctrPrp: oTokens.style.style, opEmu : 1}, null);
					if (oTokens.argSize)
					{
						let BoxMathContent = oBox.getBase();
						BoxMathContent.SetArgSize(oTokens.argSize);
					}
					UnicodeArgument(
						oTokens.value,
						MathStructures.bracket_block,
						oBox.getBase(),
					)
					break;
				case MathStructures.rect:
					let oBorderBox = oContext.Add_BorderBox({ctrPrp: oTokens.style.style}, null);
					UnicodeArgument(
						oTokens.value,
						MathStructures.bracket_block,
						oBorderBox.getBase(),
					)
					break;
				// case MathStructures.bar:
				// 	let intLocation = oTokens.overUnder === "▁" ? LOCATION_BOT : LOCATION_TOP;
				// 	let oBar = oContext.Add_Bar({ctrPrp: new CTextPr(), pos: intLocation}, null);
				// 	UnicodeArgument(
				// 		oTokens.value,
				// 		MathStructures.bracket_block,
				// 		oBar.getBase(),
				// 	);
				// 	break;
				case MathStructures.limit:
					let oLimit = oContext.Add_Limit({ctrPrp: oTokens.style.style, type: oTokens.isBelow});
					UnicodeArgument(
						oTokens.base,
						MathStructures.bracket_block,
						oLimit.getFName(),
					);
					UnicodeArgument(
						oTokens.value,
						MathStructures.bracket_block,
						oLimit.getIterator(),
					);

					break;
			}
		}
	}
	// Trow content and may skip bracket block
	function UnicodeArgument (oInput, oComparison, oContext)
	{
		if (oInput && type === 0 && oInput.type === oComparison && oInput.left === "(" && oInput.right === ")" && oInput.counter === 1)
		{
			ConvertTokens(
				oInput.value,
				oContext
			)
		}
		else if (oInput)
		{
			ConvertTokens(
				oInput,
				oContext
			)
		}
	}
//--------------------------------------Helper functions for lexer and converter------------------------------------
	function GetBracketCode(code)
	{
		const oBrackets = {
			".": -1,
			"\\{": "{".charCodeAt(0),
			"\\}": "}".charCodeAt(0),
			"\\|": "‖".charCodeAt(0),
			"|": 124,
			"〖": -1,
			"〗": -1,
			"⟨" : 10216,
			"⟩": 10217,
			"├": -1,
			"┤": -1,

		}
		if (code) {
			if (MathLiterals.rBrackets.LaTeX[code])
			{
				return MathLiterals.rBrackets.LaTeX[code].charCodeAt(0);
			}
			else if (MathLiterals.lrBrackets.LaTeX[code])
			{
				return MathLiterals.lrBrackets.LaTeX[code].charCodeAt(0);
			}
			else if (MathLiterals.lBrackets.LaTeX[code])
			{
				return MathLiterals.lBrackets.LaTeX[code].charCodeAt(0);
			}

			let strBracket = oBrackets[code];
			if (strBracket) {
				return strBracket
			}
			return code.charCodeAt(0)
		}
	}
	function GetFixedCharCodeAt(str)
	{
		let code = str.charCodeAt(0);
		let hi, low;

		if (0xd800 <= code && code <= 0xdbff) {
			hi = code;
			low = str.charCodeAt(1);
			if (isNaN(low)) {
				return null;
			}
			return (hi - 0xd800) * 0x400 + (low - 0xdc00) + 0x10000;
		}
		if (0xdc00 <= code && code <= 0xdfff) {
			return false;
		}
		return code;
	}

	function GetLaTeXFromValue(value)
	{
		if (!isGetLaTeX || value === "{" || value === "}")
			return undefined;

		let arrValue = Object.keys(AutoCorrection).filter(function(key) {
			return AutoCorrection[key] === value;
		});

		for (let i = 0; i < arrValue.length; i++)
		{
			let currentValue = arrValue[i];
			if (currentValue[0] === "\\")
			{
				return currentValue;
			}
		}
		return undefined;
	}

	let AutoCorrection = {

		'/<' : '≮',
		'/=' : '≠',
		'/>' : '≯',

		'/\\approx' : "≉",
		'/\\asymp'	: '≭',
		'/\\cong'	: '≇',
		'/\\equiv'	: '≢',
		'/\\exists'	: '∄',
		'/\\ge'		: '≱',
		'/\\gtrless': '≹',
		'/\\in'		: '∉',
		'/\\le'		: '≰',
		'/\\lessgtr': '≸',
		'/\\ni'		: '∌',
		'/\\prec'	: '⊀',
		'/\\preceq' : '⋠',
		'/\\sim'	: '≁',
		'/\\simeq'	: '≄',
		'/\\sqsubseteq' : '⋢',
		'/\\sqsuperseteq': '⋣',
		'/\\sqsupseteq' : '⋣',
		'/\\subset': '⊄',
		'/\\subseteq': '⊈',
		'/\\succ': '⊁',
		'/\\succeq': '⋡',
		'/\\supset': '⊅',
		'/\\superset': '⊅',
		'/\\superseteq': '⊉',
		'/\\supseteq': '⊉',

		"\\above": "┴",
		"\\acute": "́",
		"\\aleph": "ℵ",
		"\\alpha": "α",
		"\\Alpha": "Α",
		"\\amalg": "∐", //?
		"\\angle": "∠",
		"\\aoint": "∳",
		"\\approx": "≈",
		"\\asmash": "⬆",
		"\\ast": "∗",
		"\\asymp": "≍",
		"\\atop": "¦",
		"\\array": "■",

		"\\Bar": "̿",
		"\\bar": "̅",
		"\\backslash": "\\",
		"\\backprime": "‵",
		"\\because": "∵",
		"\\begin": "〖",
		"\\below": "┬",
		"\\bet": "ℶ",
		"\\beta": "β",
		"\\Beta": "Β",
		"\\beth": "ℶ",
		"\\bigcap": "⋂",
		"\\bigcup": "⋃",
		"\\bigodot": "⨀",
		"\\bigoplus": "⨁",
		"\\bigotimes": "⨂",
		"\\bigsqcup": "⨆",
		"\\biguplus": "⨄",
		"\\bigvee": "⋁",
		"\\bigwedge": "⋀",
		"\\binomial": "(a+b)^n=∑_(k=0)^n ▒(n¦k)a^k b^(n-k)",
		"\\bot": "⊥",
		"\\bowtie": "⋈",
		"\\box": "□",
		"\\boxdot": "⊡",
		"\\boxminus": "⊟",
		"\\boxplus": "⊞",
		"\\bra": "⟨",
		"\\break": "⤶",
		"\\breve": "̆",
		"\\bullet": "∙",

		"\\cap": "∩",
		"\\cases": "Ⓒ", //["\\cases", "█", true], TODO CHECK
		"\\cbrt": "∛",
		"\\cdot": "⋅",
		"\\cdots": "⋯",
		"\\check": "̌",
		"\\chi": "χ",
		"\\Chi": "Χ",
		"\\circ": "∘",
		"\\close": "┤",
		"\\clubsuit": "♣",
		"\\coint": "∲",
		"\\cong": "≅",
		"\\contain": "∋",
		"\\coprod": "∐",
		"\\cup": "∪",

		"\\dalet": "ℸ",
		"\\daleth": "ℸ",
		"\\dashv": "⊣",
		"\\dd": "ⅆ",
		"\\Dd": "ⅅ",
		"\\ddddot": "⃜",
		"\\dddot": "⃛",
		"\\ddot": "̈",
		"\\ddots": "⋱",
		"\\defeq": "≝",
		"\\degc": "℃",
		"\\degf": "℉",
		"\\degree": "°",
		"\\delta": "δ",
		"\\Delta": "Δ",
		"\\Deltaeq": "≜",
		"\\diamond": "⋄",
		"\\diamondsuit": "♢",
		"\\div": "÷",
		"\\dot": "̇",
		"\\doteq": "≐",
		"\\dots": "…",
		"\\doublea": "𝕒",
		"\\doubleA": "𝔸",
		"\\doubleb": "𝕓",
		"\\doubleB": "𝔹",
		"\\doublec": "𝕔",
		"\\doubleC": "ℂ",
		"\\doubled": "𝕕",
		"\\doubleD": "𝔻",
		"\\doublee": "𝕖",
		"\\doubleE": "𝔼",
		"\\doublef": "𝕗",
		"\\doubleF": "𝔽",
		"\\doubleg": "𝕘",
		"\\doubleG": "𝔾",
		"\\doubleh": "𝕙",
		"\\doubleH": "ℍ",
		"\\doublei": "𝕚",
		"\\doubleI": "𝕀",
		"\\doublej": "𝕛",
		"\\doubleJ": "𝕁",
		"\\doublek": "𝕜",
		"\\doubleK": "𝕂",
		"\\doublel": "𝕝",
		"\\doubleL": "𝕃",
		"\\doublem": "𝕞",
		"\\doubleM": "𝕄",
		"\\doublen": "𝕟",
		"\\doubleN": "ℕ",
		"\\doubleo": "𝕠",
		"\\doubleO": "𝕆",
		"\\doublep": "𝕡",
		"\\doubleP": "ℙ",
		"\\doubleq": "𝕢",
		"\\doubleQ": "ℚ",
		"\\doubler": "𝕣",
		"\\doubleR": "ℝ",
		"\\doubles": "𝕤",
		"\\doubleS": "𝕊",
		"\\doublet": "𝕥",
		"\\doubleT": "𝕋",
		"\\doubleu": "𝕦",
		"\\doubleU": "𝕌",
		"\\doublev": "𝕧",
		"\\doubleV": "𝕍",
		"\\doublew": "𝕨",
		"\\doubleW": "𝕎",
		"\\doublex": "𝕩",
		"\\doubleX": "𝕏",
		"\\doubley": "𝕪",
		"\\doubleY": "𝕐",
		"\\doublez": "𝕫",
		"\\doubleZ": "ℤ",
		"\\downarrow": "↓",
		"\\Downarrow": "⇓",
		"\\dsmash": "⬇",

		"\\ee": "ⅇ",
		"\\ell": "ℓ",
		"\\emptyset": "∅",
		"\\emsp": " ",
		"\\end": "〗",
		"\\ensp": " ",
		"\\epsilon": "ϵ",
		"\\Epsilon": "Ε",
		"\\eqarray": "█",
		"\\equiv": "≡",
		"\\eta": "η",
		"\\Eta": "Η",
		"\\exists": "∃",

		"\\forall": "∀",
		"\\fraktura": "𝔞",
		"\\frakturA": "𝔄",
		"\\frakturb": "𝔟",
		"\\frakturB": "𝔅",
		"\\frakturc": "𝔠",
		"\\frakturC": "ℭ",
		"\\frakturd": "𝔡",
		"\\frakturD": "𝔇",
		"\\frakture": "𝔢",
		"\\frakturE": "𝔈",
		"\\frakturf": "𝔣",
		"\\frakturF": "𝔉",
		"\\frakturg": "𝔤",
		"\\frakturG": "𝔊",
		"\\frakturh": "𝔥",
		"\\frakturH": "ℌ",
		"\\frakturi": "𝔦",
		"\\frakturI": "ℑ",
		"\\frakturj": "𝔧",
		"\\frakturJ": "𝔍",
		"\\frakturk": "𝔨",
		"\\frakturK": "𝔎",
		"\\frakturl": "𝔩",
		"\\frakturL": "𝔏",
		"\\frakturm": "𝔪",
		"\\frakturM": "𝔐",
		"\\frakturn": "𝔫",
		"\\frakturN": "𝔑",
		"\\frakturo": "𝔬",
		"\\frakturO": "𝔒",
		"\\frakturp": "𝔭",
		"\\frakturP": "𝔓",
		"\\frakturq": "𝔮",
		"\\frakturQ": "𝔔",
		"\\frakturr": "𝔯",
		"\\frakturR": "ℜ",
		"\\frakturs": "𝔰",
		"\\frakturS": "𝔖",
		"\\frakturt": "𝔱",
		"\\frakturT": "𝔗",
		"\\frakturu": "𝔲",
		"\\frakturU": "𝔘",
		"\\frakturv": "𝔳",
		"\\frakturV": "𝔙",
		"\\frakturw": "𝔴",
		"\\frakturW": "𝔚",
		"\\frakturx": "𝔵",
		"\\frakturX": "𝔛",
		"\\fraktury": "𝔶",
		"\\frakturY": "𝔜",
		"\\frakturz": "𝔷",
		"\\frakturZ": "ℨ",
		"\\frown": "⌑",
		"\\funcapply": "⁡",

		"\\G": "Γ",
		"\\gamma": "γ",
		"\\Gamma": "Γ",
		"\\ge": "≥",
		"\\geq": "≥",
		"\\gets": "←",
		"\\gg": "≫",
		"\\gimel": "ℷ",
		"\\grave": "̀",

		"\\hairsp": " ",
		"\\hat": "̂",
		"\\hbar": "ℏ",
		"\\heartsuit": "♡",
		"\\hookleftarrow": "↩",
		"\\hookrightarrow": "↪",
		"\\hphantom": "⬄",
		"\\hsmash": "⬌",
		"\\hvec": "⃑",

		"\\identitymatrix": "(■(1&0&0@0&1&0@0&0&1))",
		"\\ii": "ⅈ",
		"\\iiiint": "⨌",
		"\\iiint": "∭",
		"\\iint": "∬",
		"\\Im": "ℑ",
		"\\imath": "ı",
		"\\in": "∈",
		"\\inc": "∆",
		"\\infty": "∞",
		"\\int": "∫",
		"\\integral": "1/2π ∫_0^2π ▒ⅆθ/(a+b sin θ)=1/√(a^2-b^2)",
		"\\iota": "ι",
		"\\Iota": "Ι",
		"\\itimes": "⁢",
		
		"\\j": "Jay",
		"\\jj": "ⅉ",
		"\\jmath": "ȷ",
		"\\kappa": "κ",
		"\\Kappa": "Κ",
		"\\ket": "⟩",
		"\\lambda": "λ",
		"\\Lambda": "Λ",
		"\\langle": "〈",
		"\\lbbrack": "⟦",
		"\\lbrace": "\{",
		"\\lbrack": "[",
		"\\lceil": "⌈",
		"\\ldiv": "∕",
		"\\ldivide": "∕",
		"\\ldots": "…",
		"\\le": "≤",
		"\\left": "├",
		"\\leftarrow": "←",
		"\\Leftarrow": "⇐",
		"\\leftharpoondown": "↽",
		"\\leftharpoonup": "↼",
		"\\Leftrightarrow": "⇔",
		"\\leftrightarrow": "↔",

		"\\leq": "≤",
		"\\lfloor": "⌊",
		"\\lhvec": "⃐",
		"\\limit": "lim_(n→∞)⁡〖(1+1/n)^n〗=e",
		"\\ll": "≪",
		"\\lmoust": "⎰",
		"\\Longleftarrow": "⟸",
		"\\Longleftrightarrow": "⟺",
		"\\Longrightarrow": "⟹",
		"\\lrhar": "⇋",
		"\\lvec": "⃖",

		"\\mapsto": "↦",
		"\\matrix": "■",
		"\\medsp": " ",
		"\\mid": "∣",
		"\\middle": "ⓜ",
		"\\models": "⊨",
		"\\mp": "∓",
		"\\mu": "μ",
		"\\Mu": "Μ",

		"\\nabla": "∇",
		"\\naryand": "▒",
		"\\nbsp": " ",
		"\\ndiv": "⊘",
		"\\neg": "¬",
		"\\norm": "‖",
		"\\notcontain": "∌",
		"\\notelement": "∉",
		"\\nu": "ν",
		"\\Nu": "Ν",
		"\\nwarrow": "↖",

		"\\nLeftarrow" : "⇍",
		"\\nLeftrightarrow" : "⇎",
		"\\nRightarrow" : "⇏",
		"\\nVDash" : "⊯",
		"\\nVdash" : "⊮",
		"\\nVleftarrow" : "⇺",
		"\\nVleftarrowtail" : "⬺",
		"\\nVleftrightarrow" : "⇼",
		"\\nVrightarrow" : "⇻",
		"\\nVrightarrowtail" : "⤕",
		"\\nVtwoheadleftarrow" : "⬵",
		"\\nVtwoheadleftarrowtail" : "⬽",
		"\\nVtwoheadrightarrow" : "⤁",
		"\\nVtwoheadrightarrowtail" : "⤘",
		"\\napprox" : "≉",
		"\\nasymp" : "≭",
		"\\ncong" : "≇",
		"\\ne" : "≠",
		"\\nearrow" : "↗",
		"\\neq" : "≠",
		"\\nequiv" : "≢",
		"\\neswarrow" : "⤢",
		"\\ngeq" : "≱",
		"\\ngtr" : "≯",
		"\\ngtrless" : "≹",
		"\\ngtrsim" : "≵",
		"\\nhpar" : "⫲",
		"\\ni" : "∋",
		"\\niobar" : "⋾",
		"\\nis" : "⋼",
		"\\nisd" : "⋺",
		"\\nleftarrow" : "↚",
		"\\nleftrightarrow" : "↮",
		"\\nleq" : "≰",
		"\\nless" : "≮",
		"\\nlessgtr" : "≸",
		"\\nlesssim" : "≴",
		"\\nmid" : "∤",
		"\\nni" : "∌",
		"\\notasymp" : "≭",
		"\\notin" : "∉",
		"\\notslash" : "⌿",
		"\\nparallel" : "∦",
		"\\nprec" : "⊀",
		"\\npreccurlyeq" : "⋠",
		"\\npreceq" : "⋠",
		"\\nrightarrow" : "↛",
		"\\nsim" : "≁",
		"\\nsime" : "≄",
		"\\nsimeq" : "≄",
		"\\nsqsubseteq" : "⋢",
		"\\nsqsupseteq" : "⋣",
		"\\nsubset" : "⊄",
		"\\nsubseteq" : "⊈",
		"\\nsucc" : "⊁",
		"\\nsucccurlyeq" : "⋡",
		"\\nsucceq" : "⋡",
		"\\nsupset" : "⊅",
		"\\nsupseteq" : "⊉",
		"\\ntriangleleft" : "⋪",
		"\\ntrianglelefteq" : "⋬",
		"\\ntriangleright" : "⋫",
		"\\ntrianglerighteq" : "⋭",
		"\\nvDash" : "⊭",
		"\\nvLeftarrow" : "⤂",
		"\\nvLeftrightarrow" : "⤄",
		"\\nvRightarrow" : "⤃",
		"\\nvdash" : "⊬",
		"\\nvleftarrow" : "⇷",
		"\\nvleftarrowtail" : "⬹",
		"\\nvleftrightarrow" : "⇹",
		"\\nvrightarrow" : "⇸",
		"\\nvrightarrowtail" : "⤔",
		"\\nvtwoheadleftarrow" : "⬴",
		"\\nvtwoheadleftarrowtail" : "⬼",
		"\\nvtwoheadrightarrow" : "⤀",
		"\\nvtwoheadrightarrowtail" : "⤗",

		"\\o": "ο",
		"\\O": "Ο",
		"\\odot": "⊙",
		"\\of": "▒",
		"\\oiiint": "∰",
		"\\oiint": "∯",
		"\\oint": "∮",
		"\\omega": "ω",
		"\\Omega": "Ω",
		"\\ominus": "⊖",
		"\\open": "├",
		"\\oplus": "⊕",
		"\\otimes": "⊗",
		"\\overbar": "¯",
		"\\overbrace": "⏞",
		"\\overbracket": "⎴",
		"\\overline": "¯",
		"\\overparen": "⏜",
		"\\overshell": "⏠",

		"\\parallel": "∥",
		"\\partial": "∂",
		"\\perp": "⊥",
		"\\phantom": "⟡",
		"\\phi": "ϕ",
		"\\Phi": "Φ",
		"\\pi": "π",
		"\\Pi": "Π",
		"\\pm": "±",
		"\\pmatrix": "⒨",
		"\\pppprime": "⁗",
		"\\ppprime": "‴",
		"\\pprime": "″",
		"\\prec": "≺",
		"\\preceq": "≼",
		"\\prime": "′",
		"\\prod": "∏",
		"\\propto": "∝",
		"\\psi": "ψ",
		"\\Psi": "Ψ",

		"\\qdrt": "∜",
		"\\quad": " ",
		"\\quadratic": "x=(-b±√(b^2-4ac))/2a",

		"\\rangle": "〉",
		"\\Rangle": "⟫",
		"\\ratio": "∶",
		"\\rbrace": "}",
		"\\rbrack": "]",
		"\\Rbrack": "⟧",
		"\\rceil": "⌉",
		"\\rddots": "⋰",
		"\\Re": "ℜ",
		"\\rect": "▭",
		"\\rfloor": "⌋",
		"\\rho": "ρ",
		"\\Rho": "Ρ",
		"\\rhvec": "⃑",
		"\\right": "┤",
		"\\rightarrow": "→",
		"\\Rightarrow": "⇒",
		"\\rightharpoondown": "⇁",
		"\\rightharpoonup": "⇀",
		"\\rmoust": "⎱",
		"\\root": "⒭",

		"\\scripta": "𝒶",
		"\\scriptA": "𝒜",
		"\\scriptb": "𝒷",
		"\\scriptB": "ℬ",
		"\\scriptc": "𝒸",
		"\\scriptC": "𝒞",
		"\\scriptd": "𝒹",
		"\\scriptD": "𝒟",
		"\\scripte": "ℯ",
		"\\scriptE": "ℰ",
		"\\scriptf": "𝒻",
		"\\scriptF": "ℱ",
		"\\scriptg": "ℊ",
		"\\scriptG": "𝒢",
		"\\scripth": "𝒽",
		"\\scriptH": "ℋ",
		"\\scripti": "𝒾",
		"\\scriptI": "ℐ",
		"\\scriptj": "𝒥",
		"\\scriptk": "𝓀",
		"\\scriptK": "𝒦",
		"\\scriptl": "ℓ",
		"\\scriptL": "ℒ",
		"\\scriptm": "𝓂",
		"\\scriptM": "ℳ",
		"\\scriptn": "𝓃",
		"\\scriptN": "𝒩",
		"\\scripto": "ℴ",
		"\\scriptO": "𝒪",
		"\\scriptp": "𝓅",
		"\\scriptP": "𝒫",
		"\\scriptq": "𝓆",
		"\\scriptQ": "𝒬",
		"\\scriptr": "𝓇",
		"\\scriptR": "ℛ",
		"\\scripts": "𝓈",
		"\\scriptS": "𝒮",
		"\\scriptt": "𝓉",
		"\\scriptT": "𝒯",
		"\\scriptu": "𝓊",
		"\\scriptU": "𝒰",
		"\\scriptv": "𝓋",
		"\\scriptV": "𝒱",
		"\\scriptw": "𝓌",
		"\\scriptW": "𝒲",
		"\\scriptx": "𝓍",
		"\\scriptX": "𝒳",
		"\\scripty": "𝓎",
		"\\scriptY": "𝒴",
		"\\scriptz": "𝓏",
		"\\scriptZ": "𝒵",
		"\\sdiv": "⁄",
		"\\sdivide": "⁄",
		"\\searrow": "↘",
		"\\setminus": "∖",
		"\\sigma": "σ",
		"\\Sigma": "Σ",
		"\\sim": "∼",
		"\\simeq": "≃",
		"\\smash": "⬍",
		"\\smile": "⌣",
		"\\spadesuit": "♠",
		"\\sqcap": "⊓",
		"\\sqcup": "⊔",
		"\\sqrt": "√",
		"\\sqsubseteq": "⊑",
		"\\sqsuperseteq": "⊒",
		"\\star": "⋆",
		"\\subset": "⊂",
		"\\subseteq": "⊆",
		"\\succ": "≻",
		"\\succeq": "≽",
		"\\sum": "∑",
		"\\superset": "⊃",
		"\\superseteq": "⊇",
		"\\swarrow": "↙",

		"\\tau": "τ",
		"\\Tau": "Τ",
		"\\therefore": "∴",
		"\\theta": "θ",
		"\\Theta": "Θ",
		"\\thicksp": " ",
		"\\thinsp": " ",
		"\\tilde": "̃",
		"\\times": "×",
		"\\to": "→",
		"\\top": "⊤",
		"\\tvec": "⃡",

		"\\ubar": "̲",
		"\\Ubar": "̳",
		"\\underbar": "▁",
		"\\underbrace": "⏟",
		"\\underbracket": "⎵",
		"\\underline": "▁",
		"\\underparen": "⏝",
		"\\uparrow": "↑",
		"\\Uparrow": "⇑",
		"\\updownarrow": "↕",
		"\\Updownarrow": "⇕",
		"\\uplus": "⊎",
		"\\upsilon": "υ",
		"\\Upsilon": "Υ",
		
		"\\varepsilon": "ε",
		"\\varphi": "φ",
		"\\varpi": "ϖ",
		"\\varrho": "ϱ",
		"\\varsigma": "ς",
		"\\vartheta": "ϑ",
		"\\vbar": "│",
		"\\vdots": "⋮",
		"\\vec": "⃗",
		"\\vee": "∨",
		"\\vert": "|",
		"\\Vert": "‖",
		"\\Vmatrix": "⒩",
		"\\vphantom": "⇳",
		"\\vthicksp": " ",

		"\\wedge": "∧",
		"\\wp": "℘",
		"\\wr": "≀",
		
		"\\xi": "ξ",
		"\\Xi": "Ξ",

		"\\zeta": "ζ",
		"\\Zeta": "Ζ",
		"\\zwnj": "‌",
		"\\zwsp": "​",
	};

	function UpdateAutoCorrection()
	{
		let arrG_AutoCorrectionList = window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathSymbols;
		AutoCorrection = {};
		for (let i = 0; i < arrG_AutoCorrectionList.length; i++)
		{
			let arrCurrentElement = arrG_AutoCorrectionList[i];
			let data = AscCommon.convertUnicodeToUTF16(Array.isArray(arrCurrentElement[1]) ? arrCurrentElement[1] : [arrCurrentElement[1]]);
			let name = arrCurrentElement[0];
			AutoCorrection[name] = data;
		}
	}

	function UpdateFuncCorrection()
	{
		functionNames = window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathFuncs;
	}

	const SymbolsToLaTeX = {
		"∞" : "\\infty",
		"→" : "\\to",
		"…" : "\\ldots",

		"ϵ" : "\\epsilon",
		"∃" : "\\exists",
		"∀" : "\\forall",
		"≠" : "\\neq",
		"≤" : "\\le",
		"≥" : "\\geq",
		"≮" : "\\nless",
		"≰" : "\\nleq",
		"≯" : "\\ngt",
		"≱" : "\\ngeq",
		"≡" : "\\equiv",
		"∼" : "\\sim",
		"≃" : "\\simeq",
		"≈" : "\\approx",
		"≅" : "\\cong",
		"≢" : "\\nequiv",
		//"≄" : "\\nsimeq",
		"≉" : "\\napprox",
		"≇" : "\\ncong",
		"≪" : "\\ll",
		"≫" : "\\gg",
		"∈" : "\\in",
		"∋" : "\\ni",
		"∉" : "\\notin",
		"⊂" : "\\subset",
		"⊃" : "\\supset",
		"⊆" : "\\subseteq",
		"⊇" : "\\supseteq",
		"≺" : "\\prcue",
		"≻" : "\\succ",
		"≼" : "\\preccurlyeq",
		"≽" : "\\succcurlyeq",
		"⊏" : "\\sqsubset",
		"⊐" : "\\sqsupset",
		"⊑" : "\\sqsubseteq",
		"⊒" : "\\sqsupseteq",
		"∥" : "\\parallel",
		"⊥" : "\\bot",
		"⊢" : "\\vdash",
		"⊣" : "\\dashv",
		"⋈" : "\\bowtie",
		"≍" : "\\asymp",
		"∔" : "\\dotplus",
		"∸" : "\\dotminus",
		"∖" : "\\setminus",
		"⋒" : "\\Cap",
		"⋓" : "\\Cup",
		"⊟" : "\\boxminus",
		"⊠" : "\\boxtimes",
		"⊡" : "\\boxdot",
		"⊞" : "\\boxplus",
		"⋇" : "\\divideontimes",
		"⋉" : "\\ltimes",
		"⋊" : "\\rtimes",
		"⋋" : "\\leftthreetimes",
		"⋌" : "\\rightthreetimes",
		"⋏" : "\\curlywedge",
		"⋎" : "\\curlyvee",
		"⊝" : "\\odash",
		"⊺" : "\\intercal",
		"⊕" : "\\oplus",
		"⊖" : "\\ominus",
		"⊗" : "\\otimes",
		"⊘" : "\\oslash",
		"⊙" : "\\odot",
		"⊛" : "\\oast",
		"⊚" : "\\ocirc",
		"†" : "\\dag",
		"‡" : "\\ddag",
		"⋆" : "\\star",
		"⋄" : "\\diamond",
		"≀" : "\\wr",
		"△" : "\\triangle",
		"⋀" : "\\bigwedge",
		"⋁" : "\\bigvee",
		"⨀" : "\\bigodot",
		"⨂" : "\\bigotimes",
		"⨁" : "\\bigoplus",
		"⨅" : "\\bigsqcap",
		"⨆" : "\\bigsqcup",
		"⨄" : "\\biguplus",
		"⨃" : "\\bigudot",
		"∴" : "\\therefore",
		"∵" : "\\because",
		"⋘" : "\\lll",
		"⋙" : "\\ggg",
		"≦" : "\\leqq",
		"≧" : "\\geqq",
		"≲" : "\\lesssim",
		"≳" : "\\gtrsim",
		"⋖" : "\\lessdot",
		"⋗" : "\\gtrdot",
		"≶" : "\\lessgtr",
		"⋚" : "\\lesseqgtr",
		"≷" : "\\gtrless",
		"⋛" : "\\gtreqless",
		"≑" : "\\Doteq",
		"≒" : "\\fallingdotseq",
		"≓" : "\\risingdotseq",
		"∽" : "\\backsim",
		"≊" : "\\approxeq",
		"⋍" : "\\backsimeq",
		"⋞" : "\\curlyeqprec",
		"⋟" : "\\curlyeqsucc",
		"≾" : "\\precsim",
		"≿" : "\\succsim",
		"⋜" : "\\eqless",
		"⋝" : "\\eqgtr",
		"⊲" : "\\vartriangleleft",
		"⊳" : "\\vartriangleright",
		"⊴" : "\\trianglelefteq",
		"⊵" : "\\trianglerighteq",
		"⊨" : "\\models",
		"⋐" : "\\Subset",
		"⋑" : "\\Supset",
		"⊩" : "\\Vdash",
		"⊪" : "\\Vvdash",
		"≖" : "\\eqcirc",
		"≗" : "\\circeq",
		"≜" : "\\Deltaeq",
		"≏" : "\\bumpeq",
		"≎" : "\\Bumpeq",
		"∝" : "\\propto",
		"≬" : "\\between",
		"⋔" : "\\pitchfork",
		"≐" : "\\doteq",

		"⇍" : "\\nLeftarrow",
		"⇎" : "\\nLeftrightarrow",
		"⇏" : "\\nRightarrow",
		"⊯" : "\\nVDash",
		"⊮" : "\\nVdash",
		"⇺" : "\\nVleftarrow",
		"⬺" : "\\nVleftarrowtail",
		"⇼" : "\\nVleftrightarrow",
		"⇻" : "\\nVrightarrow",
		"⤕" : "\\nVrightarrowtail",
		"⬵" : "\\nVtwoheadleftarrow",
		"⬽" : "\\nVtwoheadleftarrowtail",
		"⤁" : "\\nVtwoheadrightarrow",
		"⤘" : "\\nVtwoheadrightarrowtail",
		//"≉" : "\\napprox",
		"≭" : "\\nasymp",
		// "≇" : "\\ncong",
		// "≠" : "\\ne",
		"↗" : "\\nearrow",
		// "≠" : "\\neq",
		// "≢" : "\\nequiv",
		"⤢" : "\\neswarrow",
		// "≱" : "\\ngeq",
		// "≯" : "\\ngtr",
		"≹" : "\\ngtrless",
		"≵" : "\\ngtrsim",
		"⫲" : "\\nhpar",
		// "∋" : "\\ni",
		"⋾" : "\\niobar",
		"⋼" : "\\nis",
		"⋺" : "\\nisd",
		"↚" : "\\nleftarrow",
		"↮" : "\\nleftrightarrow",
		// "≰" : "\\nleq",
		// "≮" : "\\nless",
		"≸" : "\\nlessgtr",
		"≴" : "\\nlesssim",
		"∤" : "\\nmid",
		"∌" : "\\nni",
		// "≭" : "\\notasymp",
		// "∉" : "\\notin",
		"⌿" : "\\notslash",
		"∦" : "\\nparallel",
		"⊀" : "\\nprec",
		// "⋠" : "\\npreccurlyeq",
		// "⋠" : "\\npreceq",
		"↛" : "\\nrightarrow",
		"≁" : "\\nsim",
		// "≄" : "\\nsime",
		// "≄" : "\\nsimeq",
		"⋢" : "\\nsqsubseteq",
		"⋣" : "\\nsqsupseteq",
		"⊄" : "\\nsubset",
		"⊈" : "\\nsubseteq",
		"⊁" : "\\nsucc",
		// "⋡" : "\\nsucccurlyeq",
		// "⋡" : "\\nsucceq",
		"⊅" : "\\nsupset",
		"⊉" : "\\nsupseteq",
		"⋪" : "\\ntriangleleft",
		"⋬" : "\\ntrianglelefteq",
		"⋫": "\\ntriangleright",
		"⋭" : "\\ntrianglerighteq",
		"⊭" : "\\nvDash",
		"⤂" : "\\nvLeftarrow",
		"⤄" : "\\nvLeftrightarrow",
		"⤃" : "\\nvRightarrow",
		"⊬" : "\\nvdash",
		"⇷" : "\\nvleftarrow",
		"⬹" : "\\nvleftarrowtail",
		"⇹" : "\\nvleftrightarrow",
		"⇸" : "\\nvrightarrow",
		"⤔" : "\\nvrightarrowtail",
		"⬴" : "\\nvtwoheadleftarrow",
		"⬼" : "\\nvtwoheadleftarrowtail",
		"⤀" : "\\nvtwoheadrightarrow",
		"⤗" : "\\nvtwoheadrightarrowtail",

		"ⅆ"        :"\\dd"			,
		"ⅅ" 		:"\\Dd"			,
		"ⅇ" 		:"\\ee"			,
		"ℓ" 		:"\\ell"		,
		"ℏ" 		:"\\hbar"		,
		"ⅈ" 		:"\\ii"			,
		"ℑ" 		:"\\Im"			,
		"ı" 		:"\\imath"		,
		"Jay" 		:"\\j"			,
		"ⅉ" 		:"\\jj"			,
		"ȷ" 		:"\\jmath"		,
		"∂" 		:"\\partial"	,
		"ℜ" 		:"\\Re"			,
		"℘" 		:"\\wp"			,
		"ℵ" 		:"\\aleph"		,
		"ℶ" 		:"\\bet"		,
		"ℷ" 		:"\\gimel"		,
		"ℸ" 		:"\\dalet"		,

		"Α" 		:"\\Alpha"		,
		"α" 		:"\\alpha"		,
		"Β" 		:"\\Beta"		,
		"β" 		:"\\beta"		,
		"γ" 		:"\\gamma"		,
		"Γ" 		:"\\Gamma"		,
		"Δ" 		:"\\Delta"		,
		"δ" 		:"\\delta"		,
		"Ε" 		:"\\Epsilon"	,
		"ε" 		:"\\varepsilon"	,
		"ζ" 		:"\\zeta"		,
		"Ζ" 		:"\\Zeta"		,
		"η" 		:"\\eta"		,
		"Η" 		:"\\Eta"		,
		"θ" 		:"\\theta"		,
		"Θ" 		:"\\Theta"		,
		"ϑ" 		:"\\vartheta"	,
		"ι" 		:"\\iota"		,
		"Ι" 		:"\\Iota"		,
		"κ" 		:"\\kappa"		,
		"Κ" 		:"\\Kappa"		,
		"λ" 		:"\\lambda"		,
		"Λ" 		:"\\Lambda"		,
		"μ" 		:"\\mu"			,
		"Μ" 		:"\\Mu"			,
		"ν" 		:"\\nu"			,
		"Ν" 		:"\\Nu"			,
		"ξ" 		:"\\xi"			,
		"Ξ" 		:"\\Xi"			,
		"Ο" 		:"\\O"			,
		"ο" 		:"\\o"			,
		"π" 		:"\\pi"			,
		"Π" 		:"\\Pi"			,
		"ϖ" 		:"\\varpi"		,
		"ρ" 		:"\\rho"		,
		"Ρ" 		:"\\Rho"		,
		"ϱ" 		:"\\varrho"		,
		"σ" 		:"\\sigma"		,
		"Σ" 		:"\\Sigma"		,
		"ς" 		:"\\varsigma"	,
		"τ" 		:"\\tau"		,
		"Τ" 		:"\\Tau"		,
		"υ" 		:"\\upsilon"	,
		"Υ" 		:"\\Upsilon"	,
		"ϕ" 		:"\\phi"		,
		"Φ" 		:"\\Phi"		,
		"φ" 		:"\\varphi"		,
		"χ" 		:"\\chi"		,
		"Χ" 		:"\\Chi"		,
		"ψ" 		:"\\psi"		,
		"Ψ" 		:"\\Psi"		,
		"ω" 		:"\\omega"		,
		"Ω" 		:"\\Omega"		,

		"┴" : "\\above",
		"́" : "\\acute",
		"∐"  : "\\amalg",
		"∠" : "\\angle",
		"∳" : "\\aoint",
		"⬆" : "\\asmash",
		"∗" : "\\ast",
		"¦" : "\\atop",
		"■" : "\\array",

		"̿" : "\\Bar",
		"̅" : "\\bar",
		"‵" : "\\backprime",
		"〖" : "\\begin",
		"┬" : "\\below",
		"⋂" : "\\bigcap",
		"⋃" : "\\bigcup",
		"□" : "\\box",
		"⟨" : "\\bra",
		"⤶" : "\\break",
		"̆" : "\\breve",
		"∙" : "\\bullet",

		"∩" : "\\cap",
		"Ⓒ" : "\\cases",
		"∛" : "\\cbrt",
		"⋅" : "\\cdot",
		"⋯" : "\\cdots",
		"̌" : "\\check",
		"∘" : "\\circ",
		"┤" : "\\close",
		"♣" : "\\clubsuit",
		"∲" : "\\coint",
		"∪" : "\\cup",

		"⃜" : "\\ddddot",
		"⃛" : "\\dddot",
		"̈" : "\\ddot",
		"⋱" : "\\ddots",
		"≝" : "\\defeq",
		"℃" : "\\degc",
		"℉" : "\\degf",
		"°" : "\\degree",
		"♢" : "\\diamondsuit",
		"÷" : "\\div",
		"̇" : "\\dot",
		"𝕒" : "\\doublea",
		"𝔸" : "\\doubleA",
		"𝕓" : "\\doubleb",
		"𝔹" : "\\doubleB",
		"𝕔" : "\\doublec",
		"ℂ" : "\\doubleC",
		"𝕕" : "\\doubled",
		"𝔻" : "\\doubleD",
		"𝕖" : "\\doublee",
		"𝔼" : "\\doubleE",
		"𝕗" : "\\doublef",
		"𝔽" : "\\doubleF",
		"𝕘" : "\\doubleg",
		"𝔾" : "\\doubleG",
		"𝕙" : "\\doubleh",
		"ℍ" : "\\doubleH",
		"𝕚" : "\\doublei",
		"𝕀" : "\\doubleI",
		"𝕛" : "\\doublej",
		"𝕁" : "\\doubleJ",
		"𝕜" : "\\doublek",
		"𝕂" : "\\doubleK",
		"𝕝" : "\\doublel",
		"𝕃" : "\\doubleL",
		"𝕞" : "\\doublem",
		"𝕄" : "\\doubleM",
		"𝕟" : "\\doublen",
		"ℕ" : "\\doubleN",
		"𝕠" : "\\doubleo",
		"𝕆" : "\\doubleO",
		"𝕡" : "\\doublep",
		"ℙ" : "\\doubleP",
		"𝕢" : "\\doubleq",
		"ℚ" : "\\doubleQ",
		"𝕣" : "\\doubler",
		"ℝ" : "\\doubleR",
		"𝕤" : "\\doubles",
		"𝕊" : "\\doubleS",
		"𝕥" : "\\doublet",
		"𝕋" : "\\doubleT",
		"𝕦" : "\\doubleu",
		"𝕌" : "\\doubleU",
		"𝕧" : "\\doublev",
		"𝕍" : "\\doubleV",
		"𝕨" : "\\doublew",
		"𝕎" : "\\doubleW",
		"𝕩" : "\\doublex",
		"𝕏" : "\\doubleX",
		"𝕪" : "\\doubley",
		"𝕐" : "\\doubleY",
		"𝕫" : "\\doublez",
		"ℤ" : "\\doubleZ",
		"↓" : "\\downarrow",
		"⇓" : "\\Downarrow",
		"⬇" : "\\dsmash",

		"∅" : "\\emptyset",
		" " : "\\emsp",
		"〗" : "\\end",
		" " : "\\ensp",
		"█" : "\\eqarray",

		"𝔞" : "\\fraktura",
		"𝔄" : "\\frakturA",
		"𝔟" : "\\frakturb",
		"𝔅" : "\\frakturB",
		"𝔠" : "\\frakturc",
		"ℭ" : "\\frakturC",
		"𝔡" : "\\frakturd",
		"𝔇" : "\\frakturD",
		"𝔢" : "\\frakture",
		"𝔈" : "\\frakturE",
		"𝔣" : "\\frakturf",
		"𝔉" : "\\frakturF",
		"𝔤" : "\\frakturg",
		"𝔊" : "\\frakturG",
		"𝔥" : "\\frakturh",
		"ℌ" : "\\frakturH",
		"𝔦" : "\\frakturi",
		"𝔧" : "\\frakturj",
		"𝔍" : "\\frakturJ",
		"𝔨" : "\\frakturk",
		"𝔎" : "\\frakturK",
		"𝔩" : "\\frakturl",
		"𝔏" : "\\frakturL",
		"𝔪" : "\\frakturm",
		"𝔐" : "\\frakturM",
		"𝔫" : "\\frakturn",
		"𝔑" : "\\frakturN",
		"𝔬" : "\\frakturo",
		"𝔒" : "\\frakturO",
		"𝔭" : "\\frakturp",
		"𝔓" : "\\frakturP",
		"𝔮" : "\\frakturq",
		"𝔔" : "\\frakturQ",
		"𝔯" : "\\frakturr",
		"𝔰" : "\\frakturs",
		"𝔖" : "\\frakturS",
		"𝔱" : "\\frakturt",
		"𝔗" : "\\frakturT",
		"𝔲" : "\\frakturu",
		"𝔘" : "\\frakturU",
		"𝔳" : "\\frakturv",
		"𝔙" : "\\frakturV",
		"𝔴" : "\\frakturw",
		"𝔚" : "\\frakturW",
		"𝔵" : "\\frakturx",
		"𝔛" : "\\frakturX",
		"𝔶" : "\\fraktury",
		"𝔜" : "\\frakturY",
		"𝔷" : "\\frakturz",
		"ℨ" : "\\frakturZ",
		"⌑" : "\\frown",
		"⁡" : "\\funcapply",

		"←" : "\\gets",
		"̀" : "\\grave",

		" " : "\\hairsp",
		"̂" : "\\hat",
		"♡" : "\\heartsuit",
		"↩" : "\\hookleftarrow",
		"↪" : "\\hookrightarrow",
		"⬄" : "\\hphantom",
		"⬌" : "\\hsmash",
		"⃑" : "\\hvec",

		"⨌" : "\\iiiint",
		"∭" : "\\iiint",
		"∬" : "\\iint",

		"∆" : "\\inc",
		"∫" : "\\int",
		"⁢" : "\\itimes",


		"⟩" : "\\ket",
		"〈" : "\\langle",
		"⟦" : "\\lbbrack",
		"[" : "\\lbrack",
		"⌈" : "\\lceil",
		"├" : "\\left",
		"⇐" : "\\Leftarrow",
		"↽" : "\\leftharpoondown",
		"↼" : "\\leftharpoonup",
		"⇔" : "\\Leftrightarrow",
		"↔" : "\\leftrightarrow",

		"⌊" : "\\lfloor",
		"⃐" : "\\lhvec",
		"⎰" : "\\lmoust",
		"⟸" : "\\Longleftarrow",
		"⟺" : "\\Longleftrightarrow",
		"⟹" : "\\Longrightarrow",
		"⇋" : "\\lrhar",
		"⃖" : "\\lvec",

		"↦" : "\\mapsto",
		" " : "\\medsp",
		"∣" : "\\mid",
		"ⓜ" : "\\middle",
		"∓" : "\\mp",
		"∇" : "\\nabla",
		"▒" : "\\naryand",
		" " : "\\nbsp",
		"¬" : "\\neg",
		"‖" : "\\norm",
		"↖" : "\\nwarrow",

		"∰" : "\\oiiint",
		"∯" : "\\oiint",
		"∮" : "\\oint",
		"⏞" : "\\overbrace",
		"⎴" : "\\overbracket",
		"¯" : "\\overline",
		"⏜" : "\\overparen",
		"⏠" : "\\overshell",

		"⟡" : "\\phantom",
		"±" : "\\pm",
		"⒨" : "\\pmatrix",
		"⁗" : "\\pppprime",
		"‴" : "\\ppprime",
		"″" : "\\pprime",
		"′" : "\\prime",
		"∏" : "\\prod",

		"\\qdrt": "∜",
		"\\quad": " ",

		"〉" : "\\rangle",
		"⟫" : "\\Rangle",
		"∶" : "\\ratio",
		"]" : "\\rbrack",
		"⟧" : "\\Rbrack",
		"⌉" : "\\rceil",
		"⋰" : "\\rddots",
		"▭" : "\\rect",
		"⌋" : "\\rfloor",
		"⇒" : "\\Rightarrow",
		"⇁" : "\\rightharpoondown",
		"⇀" : "\\rightharpoonup",
		"⎱" : "\\rmoust",
		"⒭" : "\\root",

		"𝒶" : "\\scripta",
		"𝒜" : "\\scriptA",
		"𝒷" : "\\scriptb",
		"ℬ" : "\\scriptB",
		"𝒸" : "\\scriptc",
		"𝒞" : "\\scriptC",
		"𝒹" : "\\scriptd",
		"𝒟" : "\\scriptD",
		"ℯ" : "\\scripte",
		"ℰ" : "\\scriptE",
		"𝒻" : "\\scriptf",
		"ℱ" : "\\scriptF",
		"ℊ" : "\\scriptg",
		"𝒢" : "\\scriptG",
		"𝒽" : "\\scripth",
		"ℋ" : "\\scriptH",
		"𝒾" : "\\scripti",
		"ℐ" : "\\scriptI",
		"𝒥" : "\\scriptj",
		"𝓀" : "\\scriptk",
		"𝒦" : "\\scriptK",
		"ℒ" : "\\scriptL",
		"𝓂" : "\\scriptm",
		"ℳ" : "\\scriptM",
		"𝓃" : "\\scriptn",
		"𝒩" : "\\scriptN",
		"ℴ" : "\\scripto",
		"𝒪" : "\\scriptO",
		"𝓅" : "\\scriptp",
		"𝒫" : "\\scriptP",
		"𝓆" : "\\scriptq",
		"𝒬" : "\\scriptQ",
		"𝓇" : "\\scriptr",
		"ℛ" : "\\scriptR",
		"𝓈" : "\\scripts",
		"𝒮" : "\\scriptS",
		"𝓉" : "\\scriptt",
		"𝒯" : "\\scriptT",
		"𝓊" : "\\scriptu",
		"𝒰" : "\\scriptU",
		"𝓋" : "\\scriptv",
		"𝒱" : "\\scriptV",
		"𝓌" : "\\scriptw",
		"𝒲" : "\\scriptW",
		"𝓍" : "\\scriptx",
		"𝒳" : "\\scriptX",
		"𝓎" : "\\scripty",
		"𝒴" : "\\scriptY",
		"𝓏" : "\\scriptz",
		"𝒵" : "\\scriptZ",

		"↘" : "\\searrow",
		"⬍" : "\\smash",
		"⌣" : "\\smile",
		"♠" : "\\spadesuit",
		"⊓" : "\\sqcap",
		"⊔" : "\\sqcup",
		"√" : "\\sqrt",
		"∑" : "\\sum",
		"↙" : "\\swarrow",

		" " : "\\thicksp",
		" " : "\\thinsp",
		"̃" : "\\tilde",
		"×" : "\\times",
		"⊤" : "\\top",
		"⃡" : "\\tvec",

		"̲" : "\\ubar",
		"̳" : "\\Ubar",
		"⏟" : "\\underbrace",
		"⎵" : "\\underbracket",
		"▁" : "\\underline",
		"⏝" : "\\underparen",
		"↑" : "\\uparrow",
		"⇑" : "\\Uparrow",
		"↕" : "\\updownarrow",
		"⇕" : "\\Updownarrow",
		"⊎" : "\\uplus",

		"│" : "\\vbar",
		"⋮" : "\\vdots",
		"⃗" : "\\vec",
		"∨" : "\\vee",
		"|" : "\\vert",
		"⒩" : "\\Vmatrix",
		"⇳" : "\\vphantom",
		" " : "\\vthicksp",

		"∧" : "\\wedge",

		"‌" : "\\zwnj",
		"​" : "\\zwsp",
	};

	function CMathContentIterator(oCMathContent)
	{
		if (oCMathContent instanceof CMathContent)
		{
			this._content 	= oCMathContent.Content;
			this._paraRun 	= null;
			this._nParaRun	= 0;
			this._index 	= oCMathContent.Content.length - 1; // индекс текущего элемента
			this.counter 	= 0; 								// количество отданных элементов
			this.currentEl 	= null;
		}
	}
	CMathContentIterator.prototype.Count = function ()
	{
		this.counter++;
	};
	CMathContentIterator.prototype.Next = function()
	{
		if (!this.IsHasContent())
			return false;

		if (this._nParaRun >= 0 && this._paraRun)
		{
			return this.GetValue();
		}
		else
		{
			let oCurrentContent = this._content[this._index];

			if (!oCurrentContent instanceof ParaRun)
			{
				// прерываем обработку здесь точно не слово для автокоррекции
				return false;
			}
			else
			{
				this._index--;
				this._paraRun 	= oCurrentContent;
				this._nParaRun 	= oCurrentContent.GetElementsCount() - 1;
				return this.GetValue();
			}
		}
	};
	CMathContentIterator.prototype.IsHasContent = function ()
	{
		return this._index >= 0 || this._nParaRun >= 0;
	};
	CMathContentIterator.prototype.GetValue = function()
	{
		if (this._nParaRun >= 0)
		{
			this.Count();
			this._nParaRun--;
			let oMathText = this._paraRun.GetElement(this._nParaRun + 1);

			// если не текст просто прерываем обработку, здесь точно не слово для автокоррекции
			if (!(oMathText instanceof CMathText))
				return false;

			this.currentEl = oMathText;
			return oMathText.GetCodePoint();
		}
		return false;
	}
	CMathContentIterator.prototype.GetCurrentEl = function ()
	{
		return this.currentEl;
	}
	function CorrectWordOnCursor(oCMathContent, IsLaTeX, isSkipFirstLetter)
	{
		let isConvert 		= false;
		let isSkipFirst 	= isSkipFirstLetter === true;
		let strLast = oCMathContent.GetLastTextElement();
		let isLastOperator 	= oCMathContent.IsLastElement(AscMath.MathLiterals.operator) || strLast === "(" || strLast === ")";
		let oContent= new CMathContentIterator(oCMathContent);
		let oLastOperator;

		if (strLast === " ")
			isSkipFirst = true;

		let str = "";

		while (oContent.IsHasContent())
		{
			let nElement = oContent.Next();

			if (nElement === false)
				break;

			let strElement = String.fromCharCode(nElement);

			if (oContent.counter === 1 && isSkipFirst)
			{
				if (isLastOperator)
				{
					oLastOperator = strElement;
				}
				continue;
			}
			
			let isContinue =
				(nElement >= 97 && nElement <= 122)
				|| (nElement >= 65 && nElement <= 90)
				|| (nElement >= 48 && nElement <= 57)
				|| nElement === 92
				|| nElement === 47; // a-zA-z && 0-9

			if (!isContinue)
				return false;

			str = strElement + str;

			if (nElement === 92 || nElement === 47)
				break;
		}

		let oCurrentEl = oContent.GetCurrentEl();
		let nSlash = oContent.Next();
		if (nSlash === 47)
		{
			str = "/" + str;
			oCurrentEl = oContent.GetCurrentEl();
		}

		let strCorrection = ConvertWord(str, IsLaTeX);
		if (strCorrection)
		{
			let oRun = RemoveCountFormMathContent(oCMathContent, str.length + 1, isLastOperator);
			//let oRun = RemoveCountFormMathContent(oCMathContent, oContent.counter - 1, isLastOperator);
			let nPos = isLastOperator ? oRun.Content.length - 1 : oRun.Content.length;

			if (MathLiterals.accent.SearchU(strCorrection))
				strCorrection = String.fromCharCode(160) + strCorrection; //add nbsp before accent, like word

			for (let i = 0; i < strCorrection.length; i++)
			{
				let oAddMath = new MathTextAdditionalData(oCurrentEl.Parent);
				oCMathContent.Add_Text(strCorrection[i], undefined, undefined, oAddMath);
			}
			isConvert = true;
		}

		oCMathContent.MoveCursorToEndPos();
		return isConvert;
	}
	function RemoveCountFormMathContent (oContent, nCount, isSkipFirst)
	{
		let oCurrentContent;
		for (let i = oContent.Content.length - 1; i >= 0; i--)
		{
			let isSkippedFirst = false;
			oCurrentContent = oContent.Content[i];
			for (let j = oCurrentContent.Content.length - 1; j >= 0; j--)
			{
				if (isSkipFirst === true)
				{
					isSkipFirst = false;
					isSkippedFirst = true;
					continue;
				}
				oCurrentContent.RemoveFromContent(j, 1, true);
				nCount--;

				if (nCount === 0)
					return oCurrentContent;
			}
		}
		return oCurrentContent;
	}

	function CorrectSpecialWordOnCursor(oCMathContent, IsLaTeX)
	{
		let oContent= new CMathContentIterator(oCMathContent);

		if (oContent.IsHasContent())
		{
			let nSecond = oContent.Next();
			let nFirst = oContent.Next();
			if (nSecond && nFirst)
			{
				let strSecondLetter = String.fromCharCode(nSecond);
				let strFirstLetter = String.fromCharCode(nFirst);

				if (strFirstLetter !== "\\" && strSecondLetter !== "\\" && CorrectSpecial(oCMathContent, strFirstLetter, strSecondLetter))
				{
					oContent._paraRun.MoveCursorToEndPos();
					return true;
				}
			}
		}
	}
	function ConvertWord(str, IsLaTeX)
	{
		if (!IsNotConvertedLaTeXWords(str) || !IsLaTeX)
		{
			return AutoCorrection[str];
		}
	}

	function IsNotConvertedLaTeXWords(str)
	{
		return arrDoNotConvertWordsForLaTeX.includes(str);
	}

	function CorrectAllWords (oCMathContent, isLaTeX)
	{
		let isConvert = false;
	
		if (oCMathContent.Type === 49)
		{
			for (let nCount = 0; nCount < oCMathContent.Content.length; nCount++)
			{
				if (oCMathContent.Content[nCount].value === 92)
				{
					let str = oCMathContent.Content[nCount].GetTextOfElement();
					let intStart = nCount;
					let intEnd = 0;

					for (let i = nCount + 1; i < oCMathContent.Content.length; i++) {

						let oContent = oCMathContent.Content[i];
						let intCode = oContent.value;
						
						if (intCode >= 97 && intCode <= 122 || intCode >= 65 && intCode <= 90) {
							intEnd = i;
							str += oContent.GetTextOfElement();
						}
						else
						{
							break;
						}

						nCount++;
					}

					if (intEnd > intStart) {

						let strCorrection = ConvertWord(str, isLaTeX);
						if (strCorrection) {
							nCount -= (intEnd - intStart);
							oCMathContent.RemoveFromContent(intStart, intEnd - intStart + 1, true);
							oCMathContent.AddText(strCorrection, intStart);
							isConvert = true;
						}
					}
				}
			}
		}
		else
		{
			for (let nCount = 0; nCount < oCMathContent.Content.length; nCount++) {
				isConvert = CorrectAllWords(oCMathContent.Content[nCount], isLaTeX) || isConvert;
			}
		}
	
		return isConvert;
	}
	function CorrectAllSpecialWords(oCMathContent, isLaTeX)
	{
		let isConvert = false;

		if (oCMathContent.Type === 49)
		{
			for (let nCount = oCMathContent.Content.length - 1; nCount >= 1; nCount--)
			{
				let str = oCMathContent.Content[nCount].GetTextOfElement();
				let strPrev = oCMathContent.Content[nCount - 1].GetTextOfElement();
				if (CorrectSpecial(oCMathContent, strPrev, str))
					nCount--;
			}
		}
		else
		{
			for (let nCount = 0; nCount < oCMathContent.Content.length; nCount++) {
				isConvert = CorrectAllSpecialWords(oCMathContent.Content[nCount], isLaTeX) || isConvert;
			}
		}

		return isConvert;
	}
	function CorrectSpecial(oCMathContent, strPrev, strNext)
	{
		for (let i = 0; i < g_DefaultAutoCorrectMathSymbolsList.length; i++)
		{
			let current = g_DefaultAutoCorrectMathSymbolsList[i];
			let strToken = strPrev + strNext;
			if (current[0] === strToken)
			{
				let data = current[1],
					str = "";

				if (Array.isArray(data))
				{
					for (let count = 0; i < data.length; i++)
					{
						data[count] = String.fromCharCode(data[count]);
					}
					str = data.join("");
				}
				else {
					str = String.fromCharCode(data);
				}

				if (str)
				{
					let nCounter = 0;
					for (let i = oCMathContent.Content.length - 1; i >= 0 && nCounter !== strToken.length; i--)
					{
						let oCurrentElement = oCMathContent.Content[i];
						if (!oCurrentElement || !oCurrentElement.Content)
							return false;

						let oCurrentElementCounter = oCurrentElement.Content.length;

						if (oCurrentElementCounter > strToken.length)
						{
							oCurrentElement.RemoveFromContent(oCurrentElementCounter - strToken.length, strToken.length);
						}
						else
						{
							nCounter += oCurrentElementCounter;
							oCMathContent.RemoveFromContent(i, 1);
						}
					}
					oCMathContent.Add_TextOnPos(oCMathContent.Content.length, str);
					return true;
				}
			}
		}
	}
	function IsStartAutoCorrection(nInputType, intCode)
	{
		if (nInputType === 0) // Unicode
		{
			return !(
				(intCode >= 97 && intCode <= 122) || //a-zA-Z
				(intCode >= 65 && intCode <= 90) || //a-zA-Z
				(intCode >= 48 && intCode <= 57) || // 0-9
				intCode === 92 ||			// "\\"
				intCode === 95 ||			// _
				intCode === 94 ||			// ^
				MathLiterals.lBrackets.SearchU(String.fromCodePoint(intCode)) ||
				MathLiterals.rBrackets.SearchU(String.fromCodePoint(intCode)) ||
				intCode === 40 ||			// (
				intCode === 41 ||			// )
				intCode === 47 ||			// /
				intCode === 46 ||			// .
				intCode === 44 ||				// ,
				intCode > 65533
			)

		}
		else if (nInputType === 1) //LaTeX
		{
			return !(
				(intCode >= 97 && intCode <= 122) || //a-zA-Z
				(intCode >= 65 && intCode <= 90) || // a-zA-Z
				(intCode >= 48 && intCode <= 57) || // 0-9
				intCode === 92||					// "\\"
				intCode === 123 ||					// {
				intCode === 125 ||					// }
				MathLiterals.lBrackets.SearchU(String.fromCodePoint(intCode)) ||
				MathLiterals.rBrackets.SearchU(String.fromCodePoint(intCode)) ||
				intCode === 95 ||					// _
				intCode === 94 ||					// ^
				intCode === 91 ||					// [
				intCode === 93 ||					// ]
				intCode === 46 ||					// .
				intCode === 44						// ,
			)
		}
	}
	function GetConvertContent(nInputType, strConversionData, oContext)
	{
		oContext.CurPos++;
		nInputType === Asc.c_oAscMathInputType.Unicode
			? AscMath.CUnicodeConverter(strConversionData, oContext)
			: AscMath.ConvertLaTeXToTokensList(strConversionData, oContext);
	}

	let isGetLaTeX = true;

	function SetIsLaTeXGetParaRun(isConvert)
	{
		isGetLaTeX = isConvert;
	}

	function GetIsLaTeXGetParaRun()
	{
		return isGetLaTeX;
	}

	function GetFractionType(strToken)
	{
		switch (strToken)
		{
			case "/"		:	return BAR_FRACTION
			case "⁄"		:	return SKEWED_FRACTION
			case "⊘"		:	return BAR_FRACTION
			case "∕"		:	return LINEAR_FRACTION
			case "¦"		:	return NO_BAR_FRACTION
			case "⒞"		:	return NO_BAR_FRACTION

			case "\\binom"	:	return NO_BAR_FRACTION
			case "\\sfrac"	:	return SKEWED_FRACTION
			case "\\frac"	:	return BAR_FRACTION
			case "\\cfrac"	:	return BAR_FRACTION
		}
	}

	function ContentWithStylesIterator(arr)
	{
		let oArr = [];
		for (let i = 0; i < arr.length; i++)
		{
			let CurrentElement = arr[i];

			if (Array.isArray(CurrentElement))
			{
				let strTemp = ContentWithStylesIterator(CurrentElement);
				oArr = oArr.concat(strTemp);
			}
			else if (CurrentElement instanceof MathText)
			{
				oArr.push(CurrentElement)
			}
			else
			{
				oArr.push(CurrentElement)
			}
		}
		return oArr;
	}
	function ContentWithStylesToText(arr)
	{
		let arrInput = ContentWithStylesIterator(arr);
		let str = "";

		for (let i = 0; i < arrInput.length; i++)
		{
			let oCurrentElement = arrInput[i];
			str += oCurrentElement.GetText();
		}

		return str;
	}
	function ConvertMathTextToText(arr)
	{
		if (arr.length === 0)
			return "";

		if (arr instanceof MathText)
		{
			return arr.GetText();
		}
		else if (arr.Type)
		{
			return arr.GetTextOfElement(false, true)
		}

		let strContent = "";

		for (let nCount = 0; nCount < arr.length; nCount++)
		{
			let CurrentElement = arr[nCount];

			if (undefined === CurrentElement)
				continue;

			if (Array.isArray(CurrentElement))
			{
				let strTemp = ConvertMathTextToText(CurrentElement);
				if (strTemp)
				{
					strContent += strTemp;
				}
			}
			else if (CurrentElement instanceof MathText)
			{
				strContent += CurrentElement.GetText();
			}
			else
			{
				strContent += CurrentElement;
			}
		}

		return strContent;
	}
	function GetOnlyText(oContent, nInputType)
	{
		let one = oContent.GetTextOfElement(nInputType);
		return ConvertMathTextToText(one);
	}

	function PosInMathText(nPos, nLength)
	{
		this.pos = nPos; 		// EndPos
		this.length = nLength;	// Length of content
	}

	/**
	 * @param {boolean|undefined|MathTextAndStyles} [isLaTeX]
	 * @constructor
	 */
	function MathTextAndStyles (isLaTeX)
	{
		if (isLaTeX instanceof MathTextAndStyles)
			return isLaTeX;

		if (isLaTeX === undefined)
			isLaTeX = false;

		this.LaTeX				= isLaTeX;
		this.arr				= [];
		this.reviewInfo			= undefined;

		this.nPos				= 0;
		this.IsWrap				= false;
		this.IsBracket			= false;
		this.IsNumbers			= false;
		this.Positions			= [];
		this.IsUnicodeBracket	= false;

		this.globalStyle		= undefined;
		this.innerStyle		= undefined; // стиль последнего элемента
		this.IsGetStyleFromFirst = true;

		this.IsSetLastInner		= false;
	}
	MathTextAndStyles.prototype.IsEmpty = function ()
	{
		return this.arr.length === 0;
	}

	MathTextAndStyles.prototype.SetGlobalStyle = function (oContent, isCtrPr)
	{
		this.globalStyle = new MathTextAdditionalData(oContent, isCtrPr)
	}
	MathTextAndStyles.prototype.GetGlobalStyle = function ()
	{
		let oStyle = this.globalStyle;
		this.globalStyle = undefined;
		return oStyle;
	}
	MathTextAndStyles.prototype.ResetGlobalStyle = function ()
	{
		this.globalStyle = undefined;
	}

	MathTextAndStyles.prototype.GetFirstStyle = function ()
	{
		if (this.arr.length > 0)
		{
			let oLastItem = this.arr[0];

			if (oLastItem instanceof MathTextAndStyles)
				return oLastItem.GetFirstStyle(this.IsSetLastInner);
			else
			{
				this.IsSetLastInner = false;
				return oLastItem.GetAdditionalData().Copy();
			}
		}
	}
	MathTextAndStyles.prototype.SetGetInnerStyle = function (isGet)
	{
		this.IsSetLastInner = isGet;
	}
	MathTextAndStyles.prototype.GetStyleFromFirst = function (oContent)
	{
		if (!this.IsGetStyleFromFirst)
		{
			this.IsGetStyleFromFirst = true;
			return oContent;
		}

		for (let i = 0; i < this.arr.length; i++)
		{
			if (this.arr[i] instanceof MathTextAndStyles && !this.arr[i].IsEmpty())
				return this.arr[i].GetStyleFromFirst();
			else if (this.arr[i] instanceof MathTextAndStyles)
				continue;
			else
				return this.arr[i].additionalMathData;
		}
	}
	MathTextAndStyles.prototype.SetNotGetStyleFromFirst = function ()
	{
		this.IsGetStyleFromFirst = false;
	}

	MathTextAndStyles.prototype.SetIsNumbers = function (isNumbers)
	{
		this.IsNumbers = isNumbers;
	}
	MathTextAndStyles.prototype.SetIsWrap = function (isWrap)
	{
		this.IsWrap = isWrap;
	};
	MathTextAndStyles.prototype.CreateInnerCopy = function()
	{
		return new MathTextAndStyles(this.LaTeX);
	};
	MathTextAndStyles.prototype.AddContainer = function()
	{
		let oMathTextAndStyles = this.CreateInnerCopy();
		this.arr.push(oMathTextAndStyles);
		return oMathTextAndStyles;
	};
	MathTextAndStyles.prototype.DelEmptyContainer = function()
	{
		if (this.arr.length > 0)
		{
			let lastContainer = this.arr[this.arr.length - 1];
			if (lastContainer instanceof MathTextAndStyles && lastContainer.IsEmpty())
			{
				this.arr.splice(this.arr.length - 1, 1);
			}
		}
	};
	MathTextAndStyles.prototype.IsLaTeX = function()
	{
		return this.LaTeX;
	};
	MathTextAndStyles.prototype.GetLengthOfContentByPos =  function(oPos)
	{
		let oContentElement = this.GetExact(oPos);
		if (oContentElement instanceof MathTextAndStyles)
		{
			return oContentElement.arr.length > 1;
		}
	};
	MathTextAndStyles.prototype.GetLastPos = function ()
	{
		if (this.Positions.length > 0)
			return this.Positions[this.Positions.length - 1];
		else
			return false;
	}
	/**
	 *
	 * @param oContent
	 * @param isNew {boolean} - Нужно ли отделять текущий контент в отдельный MathTextAndStyles
	 * @param [Wrap] {number|Array} Unicode: 0 - not wrap; 1 - special wrap; LaTeX: 1 - wrap
	 * @param isSelectedText
	 * @return {PosInMathText|false}
	 * @constructor
	 */
	MathTextAndStyles.prototype.Add = function(oContent, isNew, Wrap, isSelectedText)
	{
		if (!(oContent instanceof MathTextAndStyles) && oContent.Content.length === 0)
			return this.GetLastPos();

		let nPosCopy = this.nPos;

		if (oContent instanceof MathTextAndStyles)
		{
			this.arr.push(oContent);
			return this.AddPosition(this.nPos - nPosCopy);
		}

		if (isNew)
		{
			let oMath = this.AddContainer();
			oContent.GetTextOfElement(oMath, isSelectedText);
			this.DelEmptyContainer();

			if (oMath.IsEmpty())
				return this.GetLastPos();

			this.Increase();
			let oPos = this.AddPosition(this.nPos - nPosCopy);

			if (this.LaTeX === false)
			{
				if (Array.isArray(Wrap))
					this.WrapExactElement(oPos, Wrap[0], Wrap[1], oContent);
				else if (Wrap === 0 || oContent instanceof ParaRun)
					return oPos;
				else if (Wrap === 1 && oContent.haveMixedContent() && !oMath.IsUnicodeBracket)
					this.WrapExactElement(oPos, "〖", "〗", oContent);
				else if (!(oContent.Parent instanceof CDelimiter) && oContent.haveMixedContent && oContent.haveMixedContent() && !oMath.IsUnicodeBracket)
					this.WrapExactElement(oPos, "(", ")", oContent);
			}
			else
			{
				if (Array.isArray(Wrap))
					this.WrapExactElement(oPos, Wrap[0], Wrap[1], this.GetFirstStyle());
				else if (Wrap === 0 || oContent instanceof ParaRun)
					return oPos;
				else if (!(oContent instanceof CDelimiter) && oContent.haveMixedContent && oContent.haveMixedContent() || Wrap === 1)
					this.WrapExactElement(oPos, "{", "}", this.GetStyleFromFirst(oContent));
			}

			this.ClearReviewInfo();
			return oPos;
		}
		else
		{
			if (this.IsSetLastInner)
				this.SetGetInnerStyle(true);

			oContent.GetTextOfElement(this, isSelectedText);

			this.ClearReviewInfo();

			if (this.nPos === nPosCopy)
				return this.Get_Position();

			return this.AddPosition(this.nPos - nPosCopy);
		}
	};
	MathTextAndStyles.prototype.AddText = function(oContent, isNew)
	{
		let nPosCopy = this.nPos;

		if (this.arr.length > 0 && (this.arr[this.arr.length - 1] instanceof MathTextAndStyles ||
			!this.arr[this.arr.length - 1].IsAdditionalDataEqual(oContent.additionalMathData)
		)) {
			isNew = true;
		}

		if (this.arr[this.arr.length - 1] && this.arr[this.arr.length - 1] instanceof MathText && !isNew) {
			this.arr[this.arr.length - 1].text += oContent.text;
		}
		else
		{
			this.arr.push(oContent);
			this.Increase();
		}

		return this.AddPosition(this.nPos - nPosCopy);
	};
	MathTextAndStyles.prototype.Get_Position = function()
	{
		for (let nCounter = 0; nCounter < this.Positions.length; nCounter++)
		{
			let oCurrentPos = this.Positions[nCounter];
			if (oCurrentPos.pos === this.nPos)
			{
				return oCurrentPos;
			}
		}
	};
	MathTextAndStyles.prototype.AddPosition = function(nLength, nPos)
	{
		let oPos;

		if (nPos !== undefined)
		{
			oPos = new PosInMathText(nPos, nLength);
			for (let i = 0; i < this.Positions.length; i++)
			{
				let oCurrentPos = this.Positions[i];
				if (oCurrentPos.pos > nPos)
				{
					this.Positions.splice(i, 0, oPos);
					return oPos;
				}
			}
		}

		oPos = this.Get_Position();

		if (!oPos)
		{
			oPos = new PosInMathText(this.nPos, nLength);
			this.Positions.push(oPos);
		}

		return oPos;
	};
	MathTextAndStyles.prototype.ClearReviewInfo = function ()
	{
		this.reviewInfo = undefined;
	}
	MathTextAndStyles.prototype.GetExact = function(oPos, isText)
	{
		let oCurrent = this.arr[oPos.pos - 1];

		if (isText && oCurrent instanceof MathText)
			return oCurrent.GetText();
		else
			return oCurrent;
	};
	MathTextAndStyles.prototype.GetArrPos = function(oPos, isBefore)
	{
		let arrAfterPos = [];

		for (let nCounter = 0; nCounter < this.Positions.length; nCounter++)
		{
			let oCurrentPos = this.Positions[nCounter];
			let nCurrentPos = oCurrentPos.pos;

			if (isBefore)
			{
				if (nCurrentPos > oPos.pos - 1)
					arrAfterPos.push(oCurrentPos);
			}
			else
			{
				if (nCurrentPos > oPos.pos )
					arrAfterPos.push(oCurrentPos);
			}
		}
		return arrAfterPos;
	};
	MathTextAndStyles.prototype.ChangePositions = function(arrPositions, oFunc)
	{
		for (let nCounter = 0; nCounter < arrPositions.length; nCounter++)
		{
			let arrCurrentPos = arrPositions[nCounter];
			oFunc(arrCurrentPos);
		}
	};
	MathTextAndStyles.prototype.AddAfter = function(oPos, oContent, isNotCopyStyle)
	{
		let arrPositions = this.GetArrPos(oPos);
		let oCurrentContainer = this.GetExact(oPos);
		let oCurrent = !isNotCopyStyle && oCurrentContainer instanceof MathText ? oCurrentContainer.GetAdditionalData() : undefined;
		let oNew = oContent instanceof MathTextAndStyles || oContent instanceof MathText ? oContent :new MathText(oContent, oCurrent);

		this.arr.splice(oPos.pos, 0, oNew);

		this.nPos++;
		this.ChangePositions(arrPositions, function(oPos) {oPos.pos++});
		return  this.AddPosition(oContent.length, oPos.pos + 1);
	};
	MathTextAndStyles.prototype.AddBefore = function(oPos, oContent, isNotCopyStyle)
	{
		if (!oContent)
			return;

		let arrPositions = this.GetArrPos(oPos, true);
		let oCurrentContainer = this.GetExact(oPos);
		let oCurrent = !isNotCopyStyle && oCurrentContainer instanceof MathText ?  oCurrentContainer.GetAdditionalData() : undefined;
		let oNew = oContent instanceof MathTextAndStyles || oContent instanceof MathText ? oContent : new MathText(oContent, oCurrent);
		let nPos = oPos.pos - oPos.length;

		this.arr.splice(nPos, 0, oNew);

		this.nPos++;
		this.ChangePositions(arrPositions, function(oPos) {oPos.pos++});
		return this.AddPosition(oContent.length, nPos);
	};
	MathTextAndStyles.prototype.Wrap = function(strStart, strEnd)
	{
		this.AddBefore(this.Positions[0], strStart);
		this.AddAfter(this.Positions[this.Positions.length - 1], strEnd);
	};
	MathTextAndStyles.prototype.IsNeedWrap = function()
	{
		return this.arr.length > 1;
	};
	MathTextAndStyles.prototype.GetLength = function ()
	{
		return this.arr.length;
	}
	MathTextAndStyles.prototype.WrapExactElement = function(oPos, strOne, strTwo, oContent)
	{
		let oToken;

		if (oPos instanceof MathTextAndStyles)
			oToken = oPos;
		else
			oToken = this.GetExact(oPos);

		if (strOne && strTwo)
		{
			if (this.IsLaTeX())
			{
				if (strOne instanceof MathText && strTwo instanceof MathText)
					oToken.Wrap(strOne, strTwo);
				else
					oToken.Wrap(new MathText(strOne, oContent), new MathText(strTwo, oContent));
			}
			else
			{
				oToken.Wrap(
					new MathText(strOne, this.globalStyle ? this.globalStyle : oContent),
					new MathText(strTwo, this.globalStyle ? this.globalStyle : oContent)
				);
			}

			return;
		}

		if (!this.IsLaTeX())
		{
			oToken.Wrap(new MathText("(", this.globalStyle ? this.globalStyle : oContent), new MathText(")", this.globalStyle ? this.globalStyle : oContent));
		}
		else
		{
			oToken.Wrap(new MathText("{", this.globalStyle ? this.globalStyle : oContent), new MathText("}", this.globalStyle ? this.globalStyle : oContent));
		}
	};
	MathTextAndStyles.prototype.Increase = function()
	{
		this.nPos++;
	};
	MathTextAndStyles.prototype.GetText = function()
	{
		let strOutput = "";

		for (let nCounter = 0; nCounter < this.arr.length; nCounter++)
		{
			let oCurrentElement = this.arr[nCounter];
			if (oCurrentElement instanceof MathText)
			{
				strOutput += oCurrentElement.GetText();
			}
			else if (oCurrentElement instanceof MathTextAndStyles)
			{
				strOutput += oCurrentElement.GetText();
			}
			else
			{
				strOutput += oCurrentElement;
			}
		}

		return strOutput;
	};
	MathTextAndStyles.prototype.IsHasText = function ()
	{
		return this.GetText().trim().length > 0;
	}
	MathTextAndStyles.prototype.ChangeContent = function (str)
	{
		this.Positions = [];
		this.arr = [];
		this.nPos = 0;
		if (str instanceof AscMath.MathText)
		{
			return this.AddText(str);
		}
		else
		{
			return this.AddText(new AscMath.MathText(str));
		}
	};
	MathTextAndStyles.prototype.Flat = function (oCMathContent)
	{
		let newArr = [];
		let Flat = function (arr)
		{
			for (let i = 0; i < arr.length; i++)
			{
				let oCurrentElement 	= arr[i];
				if (oCurrentElement instanceof MathTextAndStyles)
				{
					Flat(oCurrentElement.arr);
				}
				else if (oCurrentElement instanceof MathText)
				{
					newArr.push(oCurrentElement);
				}
				else
				{
					newArr.push(new MathText(oCurrentElement));
				}
			}
		}
		Flat(this.arr);

		oCMathContent.Remove_Content(0,oCMathContent.Content.length);
		this.AddTextToMathContent(newArr, oCMathContent);
		oCMathContent.Correct_Content(true);
	}
	MathTextAndStyles.prototype.AddTextToMathContent = function (newArr, oCMathContent)
	{
		for (let i = 0; i < newArr.length; i++)
		{
			oCMathContent.Add_Text(newArr[i].text, undefined, undefined, newArr[i].additionalMathData);
		}
	}
	MathTextAndStyles.prototype.DelLastSpace = function ()
	{
		let GetLast = function (arr)
		{
			if (arr.length === 0)
				return;

			let oLast = arr[arr.length - 1];

			if (oLast instanceof MathTextAndStyles)
			{
				return GetLast(oLast.arr);
			}
			else
			{
				if (oLast.text[oLast.text.length - 1] === " ")
					oLast.text = oLast.text.slice(0, -1);
			}

		}
		GetLast(this.arr);

		return this
	}

	function MathTextAdditionalData (oContent, isCtrPr)
	{
		this.style = undefined;
		this.reviewData = {
			reviewType : undefined,
			reviewInfo : undefined
		}

		if (oContent)
			this.SetAdditionalDataFromContent(oContent, isCtrPr);
	}
	MathTextAdditionalData.prototype.Copy = function()
	{
		let oNewMath = new MathTextAdditionalData();

		oNewMath.SetAdditionalStyleData(this.style);
		oNewMath.SetAdditionalReviewData(this.reviewData);

		return oNewMath;
	};
	MathTextAdditionalData.prototype.GetAdditionalStyleData = function()
	{
		return this.style;
	};
	MathTextAdditionalData.prototype.SetAdditionalStyleData = function (oStyle)
	{
		this.style = oStyle;
	}
	MathTextAdditionalData.prototype.IsAdditionalStyleData = function()
	{
		return this.style !== undefined;
	};
	MathTextAdditionalData.prototype.GetAdditionalReviewData = function()
	{
		return this.reviewData;
	};
	MathTextAdditionalData.prototype.SetAdditionalReviewData = function (oReviewData)
	{
		this.reviewData = oReviewData;
	}
	MathTextAdditionalData.prototype.GetAdditionalReviewType = function()
	{
		return this.reviewData.reviewType;
	};
	MathTextAdditionalData.prototype.SetAdditionalReviewType = function (nReviewType)
	{
		this.reviewData.reviewType = nReviewType;
	};
	MathTextAdditionalData.prototype.IsAdditionalReviewType = function ()
	{
		return this.reviewData.reviewType !== undefined;
	};
	MathTextAdditionalData.prototype.GetAdditionalReviewInfo = function()
	{
		return this.reviewData.reviewInfo;
	};
	MathTextAdditionalData.prototype.SetAdditionalReviewInfo = function (oReviewInfo)
	{
		this.reviewData.reviewInfo = oReviewInfo;
	};
	MathTextAdditionalData.prototype.IsStyleEqual = function (oStyle)
	{
		if (this.style === undefined)
			return false;

		if (oStyle instanceof MathTextAdditionalData)
			return this.style.IsEqual(oStyle.GetAdditionalStyleData());
		else
			return this.style.IsEqual(oStyle);
	}
	MathTextAdditionalData.prototype.SetAdditionalDataFromContent = function (oContent, isCtrPrp)
	{
		let oPr;

		if (oContent instanceof ParaRun)
		{
			oPr = oContent.Pr.Copy();
		}
		else if (oContent instanceof CMathContent && !oContent.CtrPrp.IsEmpty())
		{
			oPr = oContent.GetCtrPrp();
		}
		else if (oContent instanceof CMathContent && oContent.CtrPrp.IsEmpty() && oContent.Content.length > 0)
		{
			let oItem = oContent.Content[0];
			if (oItem instanceof ParaRun)
				oPr = oItem.Pr.Copy();
		}
		else if (oContent instanceof CMathMatrix)
		{
			oPr = oContent.TextPrControlLetter.Copy();
		}
		else
		{
			oPr = oContent.CompiledCtrPrp;
		}

		this.SetAdditionalStyleData(oPr);
		this.SetAdditionalReviewType(oContent.ReviewType);
		this.SetAdditionalReviewInfo(oContent.ReviewInfo);
	}

	/**
	 * @param {string|undefined} str text of math literal
	 * @param [oContent] {MathTextAdditionalData|}
	 * @param [isCtrPr] {boolean}
	 *
	 */
	function MathText(str, oContent, isCtrPr)
	{
		this.text				= str;
		this.additionalMathData	= new MathTextAdditionalData;
		this.isWrap				= false;

		if (oContent instanceof MathTextAdditionalData)
			this.additionalMathData = oContent;
		else
			this.SetAdditionalDataFromContent(oContent, isCtrPr);
	}
	MathText.prototype.GetAdditionalData = function ()
	{
		return this.additionalMathData;
	}
	MathText.prototype.GetText = function ()
	{
		return this.text;
	};
	MathText.prototype.SetText = function (str)
	{
		this.text = str;
	}
	MathText.prototype.Check = function (func)
	{
		return func(this.text);
	};
	MathText.prototype.Wrap = function (strFirst, strSecond)
	{
		this.text = strFirst + this.text + strSecond;
		this.isWrap = true;
	};
	MathText.prototype.IsWrap = function()
	{
		return this.isWrap;
	};
	MathText.prototype.IsAdditionalDataEqual = function (oStyle)
	{
		return this.additionalMathData.IsStyleEqual(oStyle);
	}
	MathText.prototype.SetAdditionalDataFromContent = function (oContent, isCtrPrp)
	{
		this.additionalMathData.SetAdditionalDataFromContent(oContent, isCtrPrp);
	}

	//--------------------------------------------------------export----------------------------------------------------
	window["AscMath"] = window["AscMath"] || {};
	window["AscMath"].oNamesOfLiterals 				= oNamesOfLiterals;
	window["AscMath"].ConvertTokens 				= ConvertTokens;
	window["AscMath"].Tokenizer 					= Tokenizer;
	window["AscMath"].UnicodeSpecialScript 			= UnicodeSpecialScript;
	window["AscMath"].LimitFunctions 				= limitFunctions;
	window["AscMath"].functionNames 				= functionNames;
	window["AscMath"].GetMathFontChar 				= GetMathFontChar;
	window["AscMath"].AutoCorrection 				= AutoCorrection;
	window["AscMath"].CorrectWordOnCursor 			= CorrectWordOnCursor;
	window["AscMath"].CorrectAllWords 				= CorrectAllWords;
	window["AscMath"].CorrectAllSpecialWords 		= CorrectAllSpecialWords;
	window["AscMath"].CorrectSpecialWordOnCursor 	= CorrectSpecialWordOnCursor;
	window["AscMath"].IsStartAutoCorrection 		= IsStartAutoCorrection;
	window["AscMath"].GetConvertContent 			= GetConvertContent;
	window["AscMath"].MathLiterals 					= MathLiterals;
	window["AscMath"].SymbolsToLaTeX 				= SymbolsToLaTeX;
	window["AscMath"].UpdateAutoCorrection 			= UpdateAutoCorrection;
	window["AscMath"].GetLaTeXFromValue 			= GetLaTeXFromValue;
	window["AscMath"].SetIsLaTeXGetParaRun 			= SetIsLaTeXGetParaRun;
	window["AscMath"].GetIsLaTeXGetParaRun 			= GetIsLaTeXGetParaRun;
	window["AscMath"].GetFractionType 				= GetFractionType;
	window["AscMath"].UpdateFuncCorrection 			= UpdateFuncCorrection;
	window["AscMath"].MathStructures				= MathStructures;
	window["AscMath"].MathText						= MathText;
	window["AscMath"].MathTextAdditionalData		= MathTextAdditionalData;
	window["AscMath"].ConvertMathTextToText			= ConvertMathTextToText;
	window["AscMath"].GetOnlyText					= GetOnlyText;
	window["AscMath"].ContentWithStylesIterator		= ContentWithStylesIterator;
	window["AscMath"].MathTextAndStyles				= MathTextAndStyles;
})(window);

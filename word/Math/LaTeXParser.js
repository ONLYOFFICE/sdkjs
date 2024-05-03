/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
	const num = 1;//needs for debug, default value: 0
	const MathLiterals = AscMath.MathLiterals;

	const oLiteralNames = window.AscMath.oNamesOfLiterals;
	const ConvertTokens = window.AscMath.ConvertTokens;
	const Tokenizer = window.AscMath.Tokenizer;
	const LimitFunctions = window.AscMath.LimitFunctions;
	const FunctionNames = window.AscMath.functionNames;
	const GetTypeFont = window.AscMath.GetTypeFont;
	const GetMathFontChar = window.AscMath.GetMathFontChar;

	function CLaTeXParser() {
		this.oTokenizer = new Tokenizer(true);
		this.intMathFontType = -1;
		this.isReceiveOneTokenAtTime = false;
		this.isNowMatrix = false;
		this.EscapeSymbol = "";
	}
	CLaTeXParser.prototype.IsNotEscapeSymbol = function ()
	{
		return this.oLookahead.data !== this.EscapeSymbol;
	};
	CLaTeXParser.prototype.ReadTokensWhileEnd = function (arrTypeOfLiteral)
	{
		let arrLiterals = [];
		let isOne = this.isReceiveOneTokenAtTime;

		if (isOne)
		{
			let strValue = this.EatToken(arrTypeOfLiteral[0]).data;
			let oLiteral = {
				type: arrTypeOfLiteral[num],
				value: this.intMathFontType === -1
					? strValue
					: GetMathFontChar[strValue][this.intMathFontType],
			};
			arrLiterals.push(oLiteral);
		}
		else
		{
			let strLiteral = "";
			while (this.oLookahead.class === arrTypeOfLiteral[0])
			{
				let strConvert = AscMath.AutoCorrection[this.oLookahead.data];
				if (strConvert)
				{
					this.EatToken(this.oLookahead.class);
					strLiteral += strConvert;
					if (this.oLookahead.data === " ")
					{
						this.EatToken(this.oLookahead.class);
					}
				}
				else
				{
					let oCurrentChar = this.EatToken(arrTypeOfLiteral[0]).data;

					if (this.intMathFontType === -1
						|| !GetMathFontChar[oCurrentChar]
						|| !GetMathFontChar[oCurrentChar][this.intMathFontType])
					{
						strLiteral += oCurrentChar;
					}
					else
					{
						strLiteral += GetMathFontChar[oCurrentChar][this.intMathFontType];
					}
				}
			}
			arrLiterals.push({
				type: arrTypeOfLiteral[num],
				value: strLiteral,
			})
		}
		return this.GetContentOfLiteral(arrLiterals);
	};
	CLaTeXParser.prototype.SaveState = function (oLookahead)
	{
		this.oTokenizer.SaveState(oLookahead);
	};
	CLaTeXParser.prototype.RestoreState = function ()
	{
		this.oLookahead = this.oTokenizer.RestoreState();
	};
	CLaTeXParser.prototype.Parse = function (string)
	{
		this.oTokenizer.Init(string);
		this.oLookahead = this.oTokenizer.GetNextToken();
		return this.GetASTTree();
	};
	CLaTeXParser.prototype.GetASTTree = function ()
	{
		let arrExp = [];
		while (this.oLookahead.data)
		{
			if (this.IsElementLiteral())
			{
				arrExp.push(this.GetExpressionLiteral())
			}
			else
			{
				let strValue = this.oTokenizer.GetTextOfToken(this.oLookahead.index, true);
				if (undefined === strValue)
				{
					strValue = this.EatToken(this.oLookahead.class).data
				}
				else
				{
					this.EatToken(this.oLookahead.class);
				}

				if ("\\bmod" === strValue) // todo в новой версии конвертора добавить отдельный модуль для такого типа токенов
				{
					strValue = " mod "; // в обратную сторону (линейную) такие токены вряд ли получится конвертнуть,
										// а ворде такого токена просто нет
										// todo продумать как будет происходить преобразование в линейную форму
				}

				arrExp.push({
					type: oLiteralNames.charLiteral[num],
					value: strValue
				})
			}
		}
		return {
			type: "LaTeXEquation",
			body: arrExp,
		};
	};
	CLaTeXParser.prototype.GetCharLiteral = function ()
	{
		return this.ReadTokensWhileEnd(oLiteralNames.charLiteral)
	};
	CLaTeXParser.prototype.GetOtherLiteral = function ()
	{
		return this.ReadTokensWhileEnd(oLiteralNames.otherLiteral)
	};
	CLaTeXParser.prototype.GetSpaceLiteral = function ()
	{
		//todo LaTex skip all normal spaces
		return this.ReadTokensWhileEnd(oLiteralNames.spaceLiteral);
	};
	CLaTeXParser.prototype.GetNumberLiteral = function ()
	{
		return this.ReadTokensWhileEnd(oLiteralNames.numberLiteral);
	};
	CLaTeXParser.prototype.GetOperatorLiteral = function ()
	{
		const strToken = this.EatToken(oLiteralNames.operatorLiteral[0]);
		return {
			type: oLiteralNames.operatorLiteral[num],
			value: strToken.data,
		};
	};
	CLaTeXParser.prototype.IsAccentLiteral = function ()
	{
		return this.oLookahead.class === MathLiterals.accent.id;
	};
	CLaTeXParser.prototype.GetAccentLiteral = function (oBase)
	{
		let strAccent,
			oResultAccent;

		//this.oLookahead.data = MathLiterals.accent.toSymbols[this.oLookahead.data];

		if (this.oLookahead.data === "'" || this.oLookahead.data === "''")
		{
			return this.GetSubSupLiteral(oBase);
			// strAccent = this.EatToken(this.oLookahead.class).data;
			// oResultAccent = {
			// 	type: oLiteralNames.subSupLiteral[num],
			// 	value: oBase,
			// 	up: {
			// 		type: oLiteralNames.charLiteral[num],
			// 		value: strAccent,
			// 	}
			// };
		}
		else
		{
			strAccent = this.oLookahead.data;
			this.EatToken(this.oLookahead.class);

			if (MathLiterals.accent.toSymbols[strAccent])
				strAccent = MathLiterals.accent.toSymbols[strAccent];

			if (this.IsGetBelowAboveLiteral())
			{
				return this.GetBelowAboveLiteral({
					type: oLiteralNames.charLiteral[num],
					value: strAccent,
				})
			}

			oBase = this.GetArguments(1);
			oBase = this.GetContentOfLiteral(oBase);

			// \bar{\bar{}}
			if (oBase && oBase.type === MathLiterals.accent.id && oBase.value === "̅" && strAccent === "̅")
			{
				oResultAccent = {
					type: MathLiterals.accent.id,
					base: oBase.base,
					value: "̿",
				}
			}
			else
			{
				oResultAccent = {
					type: MathLiterals.accent.id,
					base: oBase,
					value: strAccent,
				};
			}
		}

		return oResultAccent;
	};
	CLaTeXParser.prototype.IsFractionLiteral = function ()
	{
		return (
			this.oLookahead.class === "\\frac"
			|| this.oLookahead.data === "\\binom"
			|| this.oLookahead.class === "\\cfrac"
			|| this.oLookahead.class === "\\sfrac"
		);
	};
	CLaTeXParser.prototype.GetFractionLiteral = function ()
	{
		let intTypeFraction = AscMath.GetFractionType(this.oLookahead.data);
		this.EatToken(this.oLookahead.class);
		const oResult = this.GetArguments(2);

		return {
			type: oLiteralNames.fractionLiteral[num],
			up: oResult[0],
			down: oResult[1],
			fracType: intTypeFraction,
		};
	};
	CLaTeXParser.prototype.IsExpBracket = function ()
	{
		return (
			this.oLookahead.class === oLiteralNames.opOpenBracket[0] ||
			this.oLookahead.class === oLiteralNames.opOpenCloseBracket[0] ||
			this.oLookahead.data === "\\left"
		);
	};
	CLaTeXParser.prototype.IsBracketLiteral = function ()
	{
		return this.oLookahead.class === oLiteralNames.opOpenBracket[0] ||
			this.oLookahead.class === oLiteralNames.opOpenCloseBracket[0] ||
			this.oLookahead.class === oLiteralNames.opCloseBracket[0];
	}
	CLaTeXParser.prototype.GetBracketLiteral = function ()
	{
		let arrBracketContent,
			strLeftSymbol,
			strRightSymbol;

		if (this.oLookahead.data === "\\left")
		{
			this.EatToken(this.oLookahead.class);

			if (this.IsBracketLiteral() ||  this.oLookahead.data === ".")
				strLeftSymbol = this.EatToken(this.oLookahead.class).data;

			arrBracketContent = this.GetContentOfBracket();

			if (this.oLookahead.data === "\\right")
			{
				this.EatToken(this.oLookahead.class);
				if (this.IsBracketLiteral() || this.oLookahead.data === ".") {
					strRightSymbol = this.EatToken(this.oLookahead.class).data;
				}
			}
		}
		else if (this.oLookahead.class === oLiteralNames.opOpenBracket[0] || this.oLookahead.class === oLiteralNames.opOpenCloseBracket[0])
		{
			strLeftSymbol = this.EatToken(this.oLookahead.class).data;
			
			if (this.oLookahead.data === "_" || this.oLookahead.data === "^")
			{
				return this.GetPreScriptLiteral()
			}

			if (strLeftSymbol === "|" || strLeftSymbol === "‖")
			{
				this.SaveState(this.oLookahead);
				arrBracketContent = this.GetContentOfBracket(strLeftSymbol)
			}
			else
			{
				arrBracketContent = this.GetContentOfBracket()
			}

			// if (this.oLookahead.class === undefined) {
			// 	this.RestoreState();
			// 	return {
			// 		type: oLiteralNames.charLiteral[num],
			// 		value: strLeftSymbol,
			// 	}
			// }

			if (this.oLookahead.class === oLiteralNames.opCloseBracket[0] || this.oLookahead.class === oLiteralNames.opOpenCloseBracket[0])
			{
				strRightSymbol = this.EatToken(this.oLookahead.class).data;
			}
		}
		return {
			type: oLiteralNames.bracketBlockLiteral[num],
			left: strLeftSymbol,
			right: strRightSymbol,
			value: arrBracketContent,
		};
	};
	CLaTeXParser.prototype.GetContentOfBracket = function (strLeftSymbol)
	{
		let arrContent = [];
		let intCountOfBracketBlock = 1;

		while (this.IsElementLiteral() || this.oLookahead.data === "∣" || this.oLookahead.data === "\\mid"|| this.oLookahead.data === "ⓜ")
		{
			if (this.IsElementLiteral())
			{
				let oToken = [this.GetExpressionLiteral(strLeftSymbol)];
				if ((oToken && !Array.isArray(oToken)) || Array.isArray(oToken) && oToken.length > 0)
				{
					arrContent.push(oToken)
				}
			}
			else
			{
				this.EatToken(this.oLookahead.class);
				intCountOfBracketBlock++;
			}
		}

		while (arrContent.length < intCountOfBracketBlock)
		{
			arrContent.push([]);
		}

		return arrContent;
	};
	CLaTeXParser.prototype.IsElementLiteral = function ()
	{
		return this.oLookahead.class !== null && this.IsNotEscapeSymbol() && (
			this.IsFractionLiteral() ||
			this.oLookahead.class === oLiteralNames.numberLiteral[0] ||
			this.oLookahead.class === oLiteralNames.charLiteral[0] ||
			this.oLookahead.class === oLiteralNames.otherLiteral[0] ||
			this.oLookahead.class === oLiteralNames.spaceLiteral[0] ||
			this.IsSqrtLiteral() ||
			this.IsExpBracket() ||
			this.IsFuncLiteral() ||
			this.oLookahead.class === "\\middle" ||
			this.IsAccentLiteral() ||
			this.IsPreScript() ||
			this.IsChangeMathFont() ||
			this.oLookahead.class === "{" ||
			this.oLookahead.class === oLiteralNames.operatorLiteral[0] ||
			this.IsReactLiteral() ||
			this.IsBoxLiteral() ||
			this.oLookahead.class === oLiteralNames.opDecimal[0] ||
			this.IsMatrixLiteral() ||
			this.IsHBracket() ||
			this.oLookahead.data === "\\below" ||
			this.oLookahead.data === "\\above" ||
			this.IsOverUnderBarLiteral() ||
			this.IsTextLiteral() ||
			this.IsSpecialSymbol()
		);
	};
	CLaTeXParser.prototype.IsSpecialSymbol = function ()
	{
		return this.oLookahead.data === "/" ||
			this.oLookahead.data === "&" ||
			this.oLookahead.data === "@" ||
			this.oLookahead.data === "." ||
			this.oLookahead.data === ","
	}
	CLaTeXParser.prototype.GetSpecialSymbol = function ()
	{
		return {
			type: oLiteralNames.charLiteral[num],
			value: this.EatToken(this.oLookahead.class).data
		}
	}
	CLaTeXParser.prototype.GetElementLiteral = function ()
	{
		if  (this.IsSpecialSymbol())
		{
			return this.GetSpecialSymbol()
		}
		else if (this.IsFractionLiteral())
		{
			return this.GetFractionLiteral();
		}
		else if (this.oLookahead.class === oLiteralNames.numberLiteral[0])
		{
			return this.GetNumberLiteral();
		}
		else if (this.oLookahead.class === oLiteralNames.charLiteral[0])
		{
			return this.GetCharLiteral();
		}
		else if (this.oLookahead.class === oLiteralNames.otherLiteral[0])
		{
			return this.GetOtherLiteral();
		}
		else if (this.oLookahead.class === oLiteralNames.opDecimal[0])
		{
			let strDecimalLiteral = this.EatToken(this.oLookahead.class).data;
			return {
				type: oLiteralNames.opDecimal[num],
				value: strDecimalLiteral
			}
		}
		else if (this.oLookahead.class === oLiteralNames.spaceLiteral[0])
		{
			return this.GetSpaceLiteral();
		}
		else if (this.IsSqrtLiteral())
		{
			return this.GetSqrtLiteral();
		}
		else if (this.IsExpBracket())
		{
			return this.GetBracketLiteral();
		}
		else if (this.IsFuncLiteral())
		{
			return this.GetFuncLiteral();
		}
		else if (this.oLookahead.class === "\\middle")
		{
			this.EatToken("\\middle");
			return {
				type: "MiddleLiteral",
				value: this.EatToken(this.oLookahead.class).data,
			};
		}
		else if (this.IsAccentLiteral())
		{
			return this.GetAccentLiteral();
		}
		else if (this.IsPreScript())
		{
			return this.GetPreScriptLiteral();
		}
		else if (this.IsChangeMathFont())
		{
			return this.GetMathFontLiteral();
		}
		// else if (this.IsSymbolLiteral()) {
		// 	return this.GetSymbolLiteral()
		// }
		else if (this.oLookahead.data === "{")
		{
			return this.GetArguments(1)[0];
		}
		else if (this.oLookahead.class === oLiteralNames.operatorLiteral[0])
		{
			return this.GetOperatorLiteral()
		}
		else if (this.IsReactLiteral())
		{
			return this.GetRectLiteral()
		}
		else if (this.IsBoxLiteral())
		{
			return this.GetBoxLiteral()
		}
		else if (this.IsMatrixLiteral())
		{
			return this.GetMatrixLiteral();
		}
		else if (this.IsOverUnderBarLiteral())
		{
			return this.GetUnderOverBarLiteral();
		}
		else if (this.IsHBracket())
		{
			return this.GetHBracketLiteral()
		}
		else if (this.IsTextLiteral())
		{
			return this.GetTextLiteral();
		}
		else if (this.oLookahead.data === "/")
		{
			this.EatToken(this.oLookahead.class);
			return {
				type: oLiteralNames.charLiteral[num],
				value: "/",
			}
		}
	};
	CLaTeXParser.prototype.IsGetBelowAboveLiteral = function()
	{
		return this.oLookahead.data === "\\below" || this.oLookahead.data === "\\above";
	}
	CLaTeXParser.prototype.GetBelowAboveLiteral = function(base)
	{
		let isBelow = true;
		if (this.oLookahead.data === "\\above")
			isBelow = false;

		let strBaseContent = AscMath.AutoCorrection[base.value];
		if (strBaseContent)
		{
			base.value = strBaseContent;
		}

		this.EatToken(this.oLookahead.class);
		let oContent = this.GetArguments(1);

		if(base && (base.type === oLiteralNames.functionLiteral[num] || base.type === oLiteralNames.functionWithLimitLiteral[num]))
		{
			let third = this.GetArguments(1);
			return {
				type: oLiteralNames.functionWithLimitLiteral[num],
				value: base.value,
				up: !isBelow ? oContent : undefined,
				down: isBelow ? oContent : undefined,
				third: third,
			}
		}

		return {
			type: oLiteralNames.belowAboveLiteral[num],
			base: base,
			value: oContent,
			isBelow: isBelow,
		};
	}
	CLaTeXParser.prototype.IsTextLiteral = function ()
	{
		return this.oLookahead.data === "\\text"
	}
	CLaTeXParser.prototype.GetTextLiteral = function ()
	{
		this.EatToken(this.oLookahead.class);
		let oContent = this.GetTextArgument();

		return {
			type: oLiteralNames.textPlainLiteral[num],
			value: oContent,
		}
	}
	CLaTeXParser.prototype.GetTextArgument = function ()
	{
		let strText = "";

		this.EatToken(this.oLookahead.class); // {

		while (this.oLookahead.data !== "}" && this.oLookahead.data !== undefined)
		{
			strText += this.EatToken(this.oLookahead.class).data;
		}

		this.EatToken(this.oLookahead.class); // }

		return strText;
	}
	CLaTeXParser.prototype.IsFuncLiteral = function ()
	{
		return this.oLookahead.class === oLiteralNames.functionLiteral[0] || this.oLookahead.class === oLiteralNames.opNaryLiteral[0]
	};
	CLaTeXParser.prototype.GetFuncLiteral = function ()
	{
		let oOutput;
		let oFuncContent = this.EatToken(this.oLookahead.class);
		if (this.oLookahead.class === "\\limits") {
			this.EatToken("\\limits");
		}

		if (this.oLookahead.data === " ")
		{
			this.EatToken(this.oLookahead.class);
		}

		let oThirdContent = !this.IsSubSup() && !this.IsGetBelowAboveLiteral()
			? this.GetArguments(1)
			: undefined;

		let name = oFuncContent.data[0] === "\\" ? oFuncContent.data.slice(1) : oFuncContent.data;
		if (LimitFunctions.includes(name)) {
			oOutput = {
				type: oLiteralNames.functionWithLimitLiteral[num],
				value: name,
			}
		}
		else if (oFuncContent.class === oLiteralNames.opNaryLiteral[0]) {
			let str = MathLiterals.nary.toSymbols[oFuncContent.data];
			oOutput = {
				type: oLiteralNames.opNaryLiteral[num],
				value: str,
			}
		}
		else {
			if (FunctionNames.includes(name)) {
				oFuncContent.data = name;
			}
			oOutput = {
				type: oLiteralNames.functionLiteral[num],
				value: oFuncContent.data,
			};
		}

		if (oThirdContent) {
			oOutput.third = oThirdContent;
		}
		return oOutput;
	};
	CLaTeXParser.prototype.IsReactLiteral = function ()
	{
		return this.oLookahead.class === oLiteralNames.rectLiteral[0]
	};
	CLaTeXParser.prototype.GetRectLiteral = function ()
	{
		this.EatToken(this.oLookahead.class);
		let oContent = this.GetArguments(1);
		return {
			type: oLiteralNames.rectLiteral[num],
			value: oContent,
		}
	};
	CLaTeXParser.prototype.IsOverUnderBarLiteral = function ()
	{
		return this.oLookahead.data === "▁" || this.oLookahead.data === "¯"
	};
	CLaTeXParser.prototype.GetUnderOverBarLiteral = function ()
	{
		let strUnderOverLine = this.EatToken(this.oLookahead.class).data;
		let oOperand = this.GetArguments(1);
		return {
			type: oLiteralNames.overBarLiteral[num],
			overUnder: strUnderOverLine,
			value: oOperand,
		};
	};
	CLaTeXParser.prototype.IsBoxLiteral = function ()
	{
		return this.oLookahead.class === oLiteralNames.boxLiteral[0];
	}
	CLaTeXParser.prototype.GetBoxLiteral = function ()
	{
		this.EatToken(this.oLookahead.class);
		let oContent = this.GetArguments(1);
		return {
			type: oLiteralNames.boxLiteral[num],
			value: oContent,
		}
	};
	CLaTeXParser.prototype.GetBorderBoxLiteral = function ()
	{
		return this.oLookahead.class === oLiteralNames.borderBoxLiteral[0];
	}
	CLaTeXParser.prototype.IsGetBorderBoxLiteral = function ()
	{
		this.EatToken(this.oLookahead.class);
		let oContent = this.GetArguments(1);
		return {
			type: oLiteralNames.borderBoxLiteral[num],
			value: oContent,
		}
	}
	CLaTeXParser.prototype.IsHBracket = function ()
	{
		return this.oLookahead.class === oLiteralNames.hBracketLiteral[0];
	};
	CLaTeXParser.prototype.GetHBracketLiteral = function ()
	{
		let oDown, oUp;
		let hBrack = this.oLookahead.data;
		this.EatToken(this.oLookahead.class);

		switch (hBrack)
		{
			case "\\overbar": hBrack = "¯"; break;
			case "\\overbrace": hBrack = "⏞"; break;
			case "\\overbracket": hBrack = "⎴"; break;
			case "\\overline": hBrack = "¯"; break;
			case "\\underline": hBrack = "▁"; break;
			case "\\overparen": hBrack = "⏜"; break;
			case "\\overshell": hBrack = "⏠"; break;
			case "\\underparen": hBrack = "⏝"; break;
			case "\\underbrace": hBrack = "⏟"; break;
			case "\\undershell": hBrack = "⏡"; break;
			case "\\underbracket": hBrack = "⎵"; break;
		}

		let oContent = this.GetArguments(1);
		if (this.oLookahead.data === "_" || this.oLookahead.data === "^")
		{
			if (this.oLookahead.class === "_")
			{
				this.EatToken(this.oLookahead.class);
				oDown = this.GetArguments(1);
			}
			else
			{
				this.EatToken(this.oLookahead.class);
				oUp = this.GetArguments(1);
			}
		}
		return {
			type: oLiteralNames.hBracketLiteral[num],
			value: oContent,
			hBrack: hBrack,
			down: oDown,
			up: oUp,
		}
	};
	CLaTeXParser.prototype.GetWrapperElementLiteral = function ()
	{
		if (!this.IsSubSup() && this.oLookahead.class !== "\\over")
		{
			let oWrapperContent = this.GetElementLiteral();

			if (this.IsSubSup() || this.oLookahead.class === "\\limits")
			{
				return this.GetSubSupLiteral(oWrapperContent);
			}
			else if (this.IsGetBelowAboveLiteral())
			{
				return this.GetBelowAboveLiteral(oWrapperContent)
			}

			// else if (this.oLookahead.class === "\\over") {
			// 	//TODO
			// }

			return oWrapperContent;
		}
	};
	CLaTeXParser.prototype.GetWrapperElement2 = function ()
	{
		let oWrapperContent = this.GetElementLiteral();

		if (this.oLookahead.class === MathLiterals.accent.id)
		{
			return this.GetAccentLiteral(oWrapperContent);
		}
		else if (this.IsGetBelowAboveLiteral())
		{
			return this.GetBelowAboveLiteral(oWrapperContent)
		}

		return oWrapperContent;
	};
	CLaTeXParser.prototype.IsSubSup = function ()
	{
		return (this.oLookahead.class === "^" || this.oLookahead.class === "_");
	};
	CLaTeXParser.prototype.GetSubSupLiteral = function (oBaseContent, isSingle)
	{
		let isLimits, oDownContent, oUpContent, oThirdContent;

		if (undefined === oBaseContent)
		{
			oBaseContent = this.GetElementLiteral();
		}
		if (this.oLookahead.class === "\\limits")
		{
			this.EatToken("\\limits");
			isLimits = true;
		}

		if (oBaseContent.type === oLiteralNames.bracketBlockLiteral[num] && oBaseContent.left === "{" && oBaseContent.right === "}")
		{
			oBaseContent = oBaseContent.value;
		}

		if (this.oLookahead.data === "'" || this.oLookahead.data === "''")
		{
			oUpContent = {
				type: oLiteralNames.charLiteral[num],
				value: this.EatToken(this.oLookahead.class).data
			}
		}

		if (this.oLookahead.class === "_")
		{
			oDownContent = this.GetPartOfSupSup();
			if (this.oLookahead.class === "^" && isSingle !== true)
			{
				oUpContent = this.GetPartOfSupSup();
			}
			else if (oDownContent && oDownContent.down === undefined && oDownContent.base)
			{
				oDownContent = oDownContent.base;
			}
		}
		else if (this.oLookahead.class === "^")
		{
			oUpContent = this.GetPartOfSupSup();
			if (this.oLookahead.class === "_" && isSingle !== true)
			{
				oDownContent = this.GetPartOfSupSup();
			}
			else if (oUpContent && oUpContent.up === undefined && oUpContent.base && oUpContent.type !== "BelowAboveLiteral")
			{
				oUpContent = oUpContent.base;
			}
		}

		if (
			oBaseContent &&
			(oBaseContent.type === oLiteralNames.functionLiteral[num] ||
				oBaseContent.type === oLiteralNames.opNaryLiteral[num] ||
				oBaseContent.type === oLiteralNames.functionWithLimitLiteral[num])
		) {
			oThirdContent = this.GetArguments(1);
		}

		return {
			type: oLiteralNames.subSupLiteral[num],
			value: oBaseContent,
			up: oUpContent,
			down: oDownContent,
			third: oThirdContent,
			isLimits: isLimits
		};
	};
	CLaTeXParser.prototype.GetOneChar = function ()
	{
		let char = this.oLookahead.data;
		this.EatToken(this.oLookahead.class);
		return char;
	}
	CLaTeXParser.prototype.GetPartOfSupSup = function ()
	{
		let oElement;
		let strSymbol = this.oLookahead.class;
		this.EatToken(strSymbol);

		if (this.oLookahead.data === "'" || this.oLookahead.data === "''")
		{
			oElement = {
				type: oLiteralNames.charLiteral[num],
				value: this.EatToken(this.oLookahead.class).data
			}
		}
		else {
			oElement = (this.oLookahead.data === "{")
				? this.GetArguments(1)
				: this.GetOneChar();
		}

		if (this.oLookahead.class === strSymbol) {
			oElement = this.GetSubSupLiteral(oElement, true);
		}
		return oElement;
	};
	CLaTeXParser.prototype.IsPreScript = function ()
	{
		return (this.oLookahead.class === "^" || this.oLookahead.class === "_");
	};
	CLaTeXParser.prototype.GetPreScriptLiteral = function ()
	{
		let oUpContent;
		let oDownContent;
		let oBaseContent;
		let oOutput;

		if (this.oLookahead.class === "_") {
			oDownContent = this.GetPartOfSupSup();
			if (this.oLookahead.class === "^") {
				oUpContent = this.GetPartOfSupSup();
			}
		}
		else if (this.oLookahead.class === "^") {
			oUpContent = this.GetPartOfSupSup();
			if (this.oLookahead.class === "_") {
				oDownContent = this.GetPartOfSupSup();
			}
		}

		if (this.oLookahead.data === "}") {
			this.EatToken(this.oLookahead.class)
		}

		oBaseContent = this.GetElementLiteral();

		oOutput = {
			type: oLiteralNames.preScriptLiteral[num],
		};
		if (oUpContent) {
			oOutput.up = oUpContent;
		}
		if (oDownContent) {
			oOutput.down = oDownContent;
		}
		if (oBaseContent) {
			oOutput.value = oBaseContent;
		}
		return oOutput;
	};
	CLaTeXParser.prototype.IsSqrtLiteral = function ()
	{
		return this.oLookahead.class === oLiteralNames.sqrtLiteral[0];
	};
	CLaTeXParser.prototype.GetSqrtLiteral = function ()
	{
		let oBaseContent, oIndexContent, oOutput;
		this.EatToken(oLiteralNames.sqrtLiteral[0]);
		if (this.oLookahead.data === "[") {
			this.EatToken(this.oLookahead.class);
			oIndexContent = this.GetExpressionLiteral("]");
			if (this.oLookahead.data === "]") {
				this.EatToken(this.oLookahead.class);
			}
		}
		oBaseContent = this.GetArguments(1);
		oOutput = {
			type: oLiteralNames.sqrtLiteral[num],
			value: oBaseContent,
		};
		if (oIndexContent) {
			oOutput.index = oIndexContent;
		}
		return oOutput;
	};
	CLaTeXParser.prototype.IsChangeMathFont = function ()
	{
		return this.oLookahead.class === oLiteralNames.mathFontLiteral[0]
	};
	CLaTeXParser.prototype.GetMathFontLiteral = function ()
	{
		let intPrevType = this.intMathFontType;
		this.intMathFontType = GetTypeFont[this.oLookahead.data];
		this.EatToken(this.oLookahead.class);

		if (this.oLookahead.data !== "{") {
			this.isReceiveOneTokenAtTime = true;
		}

		let oOutput = {
			type: oLiteralNames.mathFontLiteral[num],
			value: this.GetArguments(1)
		};
		this.isReceiveOneTokenAtTime = false;
		this.intMathFontType = intPrevType;
		return oOutput;
	};
	CLaTeXParser.prototype.IsMatrixLiteral = function ()
	{
		return (
			this.oLookahead.class === oLiteralNames.matrixLiteral[0] ||
			this.oLookahead.data === "█" ||
			this.oLookahead.data === "■" ||
			this.oLookahead.data === "\\substack"
		)
	};
	CLaTeXParser.prototype.IsAlignBlockForArray = function ()
	{
		if (this.oLookahead.data !== "{")
			return false;

		this.SaveState(this.oLookahead);

		let oAlignBlock = this.GetArguments(1);

		const IsAlignContent = function (str)
		{
			let arr = [];
			for (let i = 0; i < str.length; i++)
			{
				if (str[i] === "l" || str[i] === "c" || str[i] === "r")
					arr.push(true);
			}

			if (arr.length === str.length)
				return true;

			return false;
		}

		if (oAlignBlock.type === oLiteralNames.charLiteral[num])
		{
			let strAlignBlock = oAlignBlock.value.trim();

			if (IsAlignContent(strAlignBlock))
				return strAlignBlock;
		}
		else
		{
			this.RestoreState();
		}
	}
	CLaTeXParser.prototype.GetMatrixLiteral = function ()
	{
		let strMatrixType;

		switch (this.oLookahead.data) {
			case "\\begin{cases}":
				strMatrixType = "{";
				break
			case "\\begin{pmatrix}":
			case "\\pmatrix":
			case "⒨":
				strMatrixType = "()";
				break;
			case "\\begin{bmatrix}":
			case "\\bmatrix":
				strMatrixType = "[]";
				break;
			case "\\begin{Bmatrix}":
			case "\\Bmatrix":
				strMatrixType = "{}";
				break;
			case "\\begin{vmatrix}":
			case "\\vmatrix":
				strMatrixType = "|";
				break;
			case "\\begin{Vmatrix}":
			case "⒩":
			case "\\Vmatrix":
				strMatrixType = "‖";
				break;
			case "\\begin{array}":
			case "\\begin{equation}":
			case "\\substack":
			case "■":
			//case "█":
			default:
				strMatrixType = "";
		}

		this.isNowMatrix = true;

		let name = this.EatToken(this.oLookahead.class).data;

		if (name === "\\substack")
		{
			this.EatToken(this.oLookahead.class);
		}

		while (this.oLookahead.data === "[")
		{
			this.GetArguments(1);
		}

		//TODO align
		let align = this.IsAlignBlockForArray();

		while (this.oLookahead.data === "[")
		{
			this.GetArguments(1);
		}

		let arrMatrixContent = [];

		let nCounter = 0;
		while (this.oLookahead.data !== "}" && this.oLookahead.class !== "endOfMatrix")
		{
			let oContent = this.GetRayOfMatrixLiteral();
			if (oContent === undefined)
			{
				oContent = [];
				nCounter++;
			}
			arrMatrixContent.push(oContent);
		}

		while(arrMatrixContent.length < 1)
		{
			arrMatrixContent.push([]);
		}

		let intMaxLengthOfMatrixRow = -Infinity;
		let intIndexOfMaxMatrixRow = -1;

		for (let i = 0; i < arrMatrixContent.length; i++)
		{
			let arrContent = arrMatrixContent[i];
			if (arrContent === undefined)
			{
				arrMatrixContent[i] = arrContent = [];
			}
			intMaxLengthOfMatrixRow = arrContent.length;
			intIndexOfMaxMatrixRow = i;
		}

		for (let i = 0; i < arrMatrixContent.length; i++)
		{

			if (i !== intIndexOfMaxMatrixRow) {

				let arrMatrix = arrMatrixContent[i];

				for (let j = arrMatrix.length; j < intMaxLengthOfMatrixRow; j++)
				{
					arrMatrix.push({});
				}
			}
		}

		if (this.oLookahead.data === "}" || this.oLookahead.class === "endOfMatrix")
		{
			this.EatToken(this.oLookahead.class)
		}

		this.isNowMatrix = false;

		if (strMatrixType.length > 0)
		{
			return {
				type: oLiteralNames.bracketBlockLiteral[num],
				value: {
					type: oLiteralNames.matrixLiteral[num],
					value: arrMatrixContent,
				},
				left: strMatrixType.length === 1 ? strMatrixType : strMatrixType[0],
				right: strMatrixType.length === 1 ? strMatrixType : strMatrixType[1],
			}
		}

		return {
			type: oLiteralNames.matrixLiteral[num],
			value: arrMatrixContent
		}
	};
	CLaTeXParser.prototype.GetRayOfMatrixLiteral = function ()
	{
		let arrRayContent;

		while (this.oLookahead.data !== "\\\\" && this.oLookahead.data !== "}" && this.oLookahead.class !== "endOfMatrix") {
			arrRayContent = this.GetElementOfMatrix();
		}

		if (this.oLookahead.data === "\\\\")
		{
			this.EatToken(this.oLookahead.class);
		}

		return arrRayContent;
	};
	CLaTeXParser.prototype.GetElementOfMatrix = function ()
	{
		let arrRow = [];
		let intLength = 0;
		let intCount = 0;
		let isAlredyGetContent = false;

		while (this.IsElementLiteral() || this.oLookahead.class === "&") {
			let intCopyOfLength = intLength;

			if (this.oLookahead.class !== "&")
			{
				arrRow.push(this.GetExpressionLiteral("&"));
				intLength++;
				isAlredyGetContent = true;
			}
			else
			{
				this.EatToken("&");

				if (isAlredyGetContent === false)
				{
					arrRow.push({});
					intCount++;
					intLength++;
				}
				else if (intCopyOfLength === intLength)
				{
					intCount++;
				}
			}
		}

		if (intLength !== intCount + 1)
		{
			for (let j = intLength; j <= intCount; j++)
			{
				arrRow.push({});
			}
		}

		return arrRow;
	};
	CLaTeXParser.prototype.IsExpressionLiteral = function(strBreak)
	{
		const arrEndOfExpression = ["}", "\\endgroup", "\\end", "┤"];

		//todo refactor
		return 	this.IsElementLiteral() &&
				this.oLookahead.data !== strBreak &&
				this.oLookahead.data !== this.EscapeSymbol &&
				!arrEndOfExpression.includes(this.oLookahead.data)
	};
	CLaTeXParser.prototype.GetExpressionLiteral = function (strBreakSymbol)
	{
		this.EscapeSymbol = strBreakSymbol;
		const arrExpList = [];

		while (this.IsExpressionLiteral(strBreakSymbol))
		{
			if (this.IsPreScript())
				arrExpList.push(this.GetPreScriptLiteral());
			else
				arrExpList.push(this.GetWrapperElementLiteral());
		}

		this.EscapeSymbol = undefined;
		return this.GetContentOfLiteral(arrExpList)
	};
	CLaTeXParser.prototype.EatToken = function (tokenType)
	{
		if (tokenType || this.oLookahead.class === tokenType) {
			const oToken = this.oLookahead;
			if (oToken === null) {
				console.log('Unexpected end of input, expected: ' + tokenType);
			}
			if (oToken.class !== tokenType) {
				console.log('Unexpected token: ' + oToken.class + ', expected: ' + tokenType);
			}
			this.oLookahead = this.oTokenizer.GetNextToken();
			return oToken;
		}
	};
	CLaTeXParser.prototype.SkipFreeSpace = function ()
	{
		while (this.oLookahead.class === oLiteralNames.spaceLiteral[0]) {
			this.oLookahead = this.oTokenizer.GetNextToken();
		}
	};
	CLaTeXParser.prototype.GetArguments = function (intCountOfArguments)
	{
		let oArgument = [];
		while (intCountOfArguments > 0) {
			if (this.oLookahead.data === "{") {
				this.EatToken(this.oLookahead.class);
				oArgument.push(this.GetExpressionLiteral());
				this.EatToken(this.oLookahead.class);
			}
			else {
				oArgument.push(this.GetWrapperElementLiteral());
			}
			intCountOfArguments--;
		}
		if (oArgument.length === 1 && Array.isArray(oArgument)) {
			return oArgument[0];
		}
		return oArgument;
	};
	CLaTeXParser.prototype.GetContentOfLiteral = function (oContent)
	{
		if (Array.isArray(oContent))
		{
			if (oContent.length === 1)
				return oContent[0];

			return oContent;
		}

		return oContent;
	};
	function ConvertLaTeXToTokensList(str, oContext, isGetOnlyTokens)
	{
		if (undefined === str || null === str)
			return;

		const oConverter = new CLaTeXParser(true);
		const oTokens = oConverter.Parse(str);

		if (!isGetOnlyTokens)
			ConvertTokens(oTokens, oContext);
		else
			return oTokens;

		return true;
	};

	//---------------------------------------export----------------------------------------------------
	window["AscMath"] = window["AscMath"] || {};
	window["AscMath"].ConvertLaTeXToTokensList = ConvertLaTeXToTokensList;
})(window);

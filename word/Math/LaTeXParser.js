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
	const Literals			= AscMath.MathLiterals;
	const Struc				= AscMath.MathStructures;
	const ConvertTokens		= AscMath.ConvertTokens;
	const Tokenizer			= AscMath.Tokenizer;
	const GetTypeFont		= AscMath.GetTypeFont;
	const GetMathFontChar	= AscMath.GetMathFontChar;

	function CLaTeXParser()
	{
		this.oTokenizer					= new Tokenizer(true);
		this.intMathFontType			= -1;
		this.isReceiveOneTokenAtTime	= false;
		this.isNowMatrix				= false;
		this.EscapeSymbol				= "";
	}
	CLaTeXParser.prototype.IsNotEscapeSymbol = function ()
	{
		return this.oLookahead.data !== this.EscapeSymbol;
	};
	CLaTeXParser.prototype.ReadTokensWhileEnd = function (arrTypeOfLiteral, type)
	{
		let arrLiterals = [];
		let strLiteral = "";
		let styles = [];

		let isOne = this.isReceiveOneTokenAtTime;
		let isSymbol;

		if (isOne && this.oLookahead.class === arrTypeOfLiteral.id)
		{
			let oItem		= this.oLookahead;
			let strValue	= this.EatToken(arrTypeOfLiteral.id).data;
			let oStyle		= oItem.style;
			type			= Struc.char;

			strLiteral = this.intMathFontType === -1
					? strValue
					: GetMathFontChar[strValue][this.intMathFontType];

			styles.push(oStyle);
		}
		else
		{
			while (this.oLookahead.class === arrTypeOfLiteral.id
				&& this.EscapeSymbol !== this.oLookahead.data
				&& (styles.length === 0 || styles[styles.length - 1].IsStyleEqual(this.oLookahead.style)))
			{
				let oCurrentItem	= this.oLookahead;
				let strCurrent		= oCurrentItem.data;
				let oStyle			= oCurrentItem.style;
				this.EatToken(arrTypeOfLiteral.id);

				if (GetMathFontChar[strCurrent] && GetMathFontChar[strCurrent][this.intMathFontType])
				{
					strLiteral += GetMathFontChar[strCurrent][this.intMathFontType];
					isSymbol = true;
				}
				else
				{
					strLiteral += strCurrent;
				}

				styles.push(oStyle);
			}
		}

		if (styles.length === 1)
			styles = styles[0];

		arrLiterals.push({type: isSymbol ? Struc.other : type, value: strLiteral, style: styles});

		if (arrLiterals.length === 1)
			return arrLiterals[0];

		return arrLiterals
	};
	CLaTeXParser.prototype.SaveState = function (oLookahead)
	{
		this.oTokenizer.SaveState(oLookahead);
	};
	CLaTeXParser.prototype.RestoreState = function ()
	{
		let oRestore =  this.oTokenizer.RestoreState();
		if (oRestore)
			this.oLookahead = oRestore;
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
				let oData = this.oLookahead;
				this.EatToken(this.oLookahead.class);
				let strValue = oData.data;
				let oStyle = oData.style;

				if ("\\bmod" === strValue) // todo в новой версии конвертора добавить отдельный модуль для такого типа токенов
				{
					strValue = " mod "; // в обратную сторону (линейную) такие токены вряд ли получится конвертнуть,
										// а ворде такого токена просто нет
										// todo продумать как будет происходить преобразование в линейную форму
				}

				arrExp.push({
					type: Struc.char,
					value: strValue,
					style: oStyle,
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
		return this.ReadTokensWhileEnd(Literals.char, Struc.char);
	};
	CLaTeXParser.prototype.GetOperandLiteral = function ()
	{
		let oPr = this.oLookahead.style;
		const strToken = this.EatToken(Literals.operand.id);
		return {
			type: Struc.char,
			value: strToken.data,
			style: oPr
		};
	}
	CLaTeXParser.prototype.GetOtherLiteral = function ()
	{
		return this.ReadTokensWhileEnd(Literals.other, Struc.other);
	};
	CLaTeXParser.prototype.GetSpaceLiteral = function ()
	{
		// LaTex skip all normal spaces
		while (this.oLookahead.data === " ")
		{
			this.EatToken(this.oLookahead.class);
		}

		let oStyle = this.oLookahead.style;

		if (this.oLookahead.data === "\\ "
			// || this.oLookahead.data === "\\quad"
			// || this.oLookahead.data === "\\:"
			// || this.oLookahead.data === "\\;"
		)
		{
			let strSpace = "";

			if (this.oLookahead.data === "\\ ")
				strSpace = " ";
			// else if (this.oLookahead.data === "\\quad")
			// 	strSpace = " ";
			// else if (this.oLookahead.data === "\\:")
			// 	strSpace = " ";
			// else if (this.oLookahead.data === "\\;")
			// 	strSpace = " ";

			this.EatToken(this.oLookahead.class)

			return {
				type: Struc.char,
				value: strSpace,
				style: oStyle
			}
		}

		//this.ReadTokensWhileEnd(Literals.space, Struc.space);
	};
	CLaTeXParser.prototype.GetNumberLiteral = function ()
	{
		return this.ReadTokensWhileEnd(Literals.number, Struc.number);
	};
	CLaTeXParser.prototype.GetOperatorLiteral = function ()
	{
		let oPr = this.oLookahead.style;
		const strToken = this.EatToken(Literals.operator.id);
		return {
			type: Struc.char,
			value: strToken.data,
			style: oPr
		};
	};
	CLaTeXParser.prototype.IsAccentLiteral = function ()
	{
		return this.oLookahead.class === Literals.accent.id;
	};
	CLaTeXParser.prototype.GetAccentLiteral = function (oBase)
	{
		let strAccent,
			oResultAccent,
			oPr;

		if (this.oLookahead.data === "'" || this.oLookahead.data === "''")
		{
			return this.GetSubSupLiteral(oBase);
		}
		else
		{
			oPr = this.oLookahead.style;
			strAccent = this.oLookahead;
			this.EatToken(this.oLookahead.class);

			strAccent.value = AscMath.MathLiterals.accent.LaTeX[strAccent.data];

			oBase = this.GetArguments(1);
			oBase = this.GetContentOfLiteral(oBase);

			oResultAccent = {
				type: Struc.accent,
				base: oBase,
				value: strAccent,
				style: oPr,
			};
		}

		return oResultAccent;
	};
	CLaTeXParser.prototype.IsFractionLiteral = function ()
	{
		return this.oLookahead.data === "\\frac"
			|| this.oLookahead.data === "\\binom"
			|| this.oLookahead.data === "\\cfrac"
			|| this.oLookahead.data === "\\sfrac";
	};
	CLaTeXParser.prototype.GetFractionType = function (str)
	{
		switch (str)
		{
			case "\\frac"	:	return BAR_FRACTION;    break;
			case "\\over"	:	return BAR_FRACTION;    break;
			case "\\binom"	:	return NO_BAR_FRACTION; break;
			case "\\atop"	:	return NO_BAR_FRACTION; break;
			default         :   return SKEWED_FRACTION; break;
		}
	}
	CLaTeXParser.prototype.GetFractionLiteral = function ()
	{
		let strName		= this.oLookahead.data;
		let oFracStyle	= this.oLookahead.style;
		let type		= this.GetFractionType(this.oLookahead.data);

		this.EatToken(this.oLookahead.class);

		let oResult		= this.GetArguments(2);

		let oFrac = {
			type: Struc.frac,
			up: oResult[0] || {},
			down: oResult[1] || {},
			fracType: type,
			style: oFracStyle,
		};

		if (strName === '\\binom')
		{
			return {
				type: Struc.bracket_block,
				left: "(",
				right: ")",
				value: [oFrac],
				style: {
					startStyle : oFracStyle,
					endStyle : oFracStyle,
					middle: oFracStyle,
				},
			};
		}

		return oFrac
	};
	CLaTeXParser.prototype.IsOverLikeLiteral = function ()
	{
		return this.oLookahead.data === "\\over"
			|| this.oLookahead.data === "\\atop"
	}
	CLaTeXParser.prototype.GetOverLikeLiteral = function (oNum)
	{
		if (!oNum)
			oNum		= this.GetArguments(1);

		let oFracStyle	= this.oLookahead.style;
		let type		= this.GetFractionType(this.oLookahead.data);
		this.EatToken(this.oLookahead.class);
		let oDen		= this.GetArguments(1);

		return {
			type: Struc.frac,
			up: oNum || {},
			down: oDen || {},
			fracType: type,
			style: oFracStyle,
		};
	}
	CLaTeXParser.prototype.IsExpBracket = function ()
	{
		return this.oLookahead.class === Literals.lrBrackets.id
			|| this.oLookahead.class === Literals.lBrackets.id
			|| this.oLookahead.data === "\\left";
	};
	CLaTeXParser.prototype.GetBracketLiteral = function ()
	{
		let arrBracketContent,
			strLeftSymbol,
			strRightSymbol,
			startStyle,
			middle_styles = [],
			endStyle;

		let isRightAndLeft = false;
		this.SaveState(this.oLookahead);

		if (this.oLookahead.data === "\\left")
		{
			if (this.oLookahead.data === "\\left")
				isRightAndLeft = true;

			this.EatToken(this.oLookahead.class);

			if (this.oLookahead.data !== "\\left"
				&& this.oLookahead.class === Literals.lBrackets.id
					|| this.oLookahead.data === "."
					|| this.oLookahead.class === Literals.lrBrackets.id
					|| this.oLookahead.class === Literals.rBrackets.id)
			{
				startStyle = this.oLookahead.style;
				strLeftSymbol = this.EatToken(this.oLookahead.class).data;
			}
			else
			{
				this.RestoreState();
				let oData = this.oLookahead;
				this.EatToken(this.oLookahead.class);
				return {
					type: Struc.char,
					value: oData.data,
					style: oData.style,
				}
			}

			arrBracketContent = this.GetContentOfBracket("\\right", middle_styles);

			if (this.oLookahead.data === "\\right")
			{
				if (this.oLookahead.data === "\\right" && isRightAndLeft)
					isRightAndLeft = true;
				else
					isRightAndLeft = false;

				this.EatToken(this.oLookahead.class);
				if (this.oLookahead.data !== "\\right"
					&& this.oLookahead.class === Literals.rBrackets.id
						|| this.oLookahead.data === "."
						|| this.oLookahead.class === Literals.lrBrackets.id
						|| this.oLookahead.class === Literals.lBrackets.id)
				{
					endStyle = this.oLookahead.style;
					strRightSymbol = this.EatToken(this.oLookahead.class).data;
				}
			}
		}
		else if (this.oLookahead.class === Literals.lBrackets.id || this.oLookahead.class === Literals.lrBrackets.id)
		{
			startStyle		= this.oLookahead.style;
			strLeftSymbol	= this.EatToken(this.oLookahead.class).data;
			
			if (this.oLookahead.data === "_" || this.oLookahead.data === "^")
				return this.GetPreScriptLiteral();

			if (this.oLookahead.class === undefined) {
				return {
					type: Struc.char,
					value: strLeftSymbol,
					style: startStyle,
				}
			}

			if (strLeftSymbol === "|" || strLeftSymbol === "‖" || strLeftSymbol === "\\|")
			{
				arrBracketContent = this.GetContentOfBracket(strLeftSymbol, middle_styles);
			}
			else
			{
				arrBracketContent = this.GetContentOfBracket(undefined, middle_styles);
			}

			if (this.oLookahead.class === Literals.rBrackets.id || this.oLookahead.class === Literals.lrBrackets.id)
			{
				endStyle = this.oLookahead.style;
				strRightSymbol = this.EatToken(this.oLookahead.class).data;
			}
		}

		if (strLeftSymbol === "{" && strRightSymbol === "}" && !isRightAndLeft)
			return arrBracketContent;

		return {
			type: Struc.bracket_block,
			left: strLeftSymbol,
			right: strRightSymbol,
			value: arrBracketContent,
			style: {
				startStyle : startStyle,
				endStyle : endStyle,
				middle: middle_styles,
			},
		};
	};
	CLaTeXParser.prototype.GetContentOfBracket = function (strLeftSymbol, arrMiddleStyles)
	{
		let arrContent = [];
		let intCountOfBracketBlock = 1;

		while (this.IsElementLiteral()
			|| this.oLookahead.data === "∣"
			|| this.oLookahead.data === "\\mid"
			|| this.oLookahead.data === "ⓜ")
		{
			if (strLeftSymbol && this.oLookahead.data === strLeftSymbol)
				break;

			if (this.oLookahead.data === "\\right")
				break;

			if (this.IsElementLiteral())
			{
				// normal space always skip in LaTeX brackets
				this.SkipFreeSpace();

				let oToken = [this.GetExpressionLiteral([strLeftSymbol])];
				if ((oToken && !Array.isArray(oToken)) || Array.isArray(oToken) && oToken.length > 0)
				{
					arrContent.push(oToken)
				}

				// normal space always skip in LaTeX brackets
				this.SkipFreeSpace();
			}
			else
			{
				arrMiddleStyles.push(this.oLookahead.style)
				this.EatToken(this.oLookahead.class);
				intCountOfBracketBlock++;
				// normal space always skip in LaTeX brackets
				this.SkipFreeSpace();
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
			this.oLookahead.class === Literals.number.id ||
			this.oLookahead.class === Literals.char.id ||
			this.oLookahead.class === Literals.space.id ||
			this.IsSqrtLiteral() ||
			this.IsExpBracket() ||
			this.IsFuncLiteral() ||
			this.oLookahead.class === "\\middle" ||
			this.IsAccentLiteral() ||
			this.IsPreScript() ||
			this.IsMathFontLiteral() ||
			this.oLookahead.class === "{" ||
			this.oLookahead.class === Literals.operator.id ||
			this.IsReactLiteral() ||
			this.IsBoxLiteral() ||
		//	this.oLookahead.class === oLiteralNames.opDecimal[0] ||
			this.IsArrayLiteral() ||
			this.IsMatrixLiteral() ||
			this.IsHBracket() ||
			this.oLookahead.data === "\\below" ||
			this.oLookahead.data === "\\above" ||
			this.IsOverUnderBarLiteral() ||
			this.IsTextLiteral() ||
			this.IsSpecialSymbol() ||
			this.oLookahead.class === Literals.other.id ||
			this.oLookahead.class === Literals.operand.id ||
			this.IsHorizontalArrowLiteral() ||
			this.oLookahead.class === Literals.punct.id
		);
	};
	CLaTeXParser.prototype.IsSpecialSymbol = function ()
	{
		return this.oLookahead.data === "/" ||
			this.oLookahead.data === "&" ||
			this.oLookahead.data === "@";
	}
	CLaTeXParser.prototype.GetSpecialSymbol = function ()
	{
		let oData = this.oLookahead;
		this.EatToken(this.oLookahead.class)
		return {
			type: Struc.char,
			value: oData.data,
			style: oData.style,
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
		else if (this.oLookahead.class === Literals.number.id)
		{
			return this.GetNumberLiteral();
		}
		else if (this.oLookahead.class === Literals.operand.id)
		{
			return this.GetOperandLiteral();
		}
		else if (this.oLookahead.class === Literals.char.id)
		{
			return this.GetCharLiteral();
		}
		else if (this.oLookahead.class === Literals.punct.id)
		{
			return this.ReadTokensWhileEnd(Literals.punct, Struc.char);
		}
		else if (this.IsHorizontalArrowLiteral())
		{
			return this.GetHorizontalArrowLiteral();
		}
		else if (this.oLookahead.class === Literals.char.id)
		{
			return this.GetOtherLiteral();
		}
		// else if (this.oLookahead.class === oLiteralNames.opDecimal[0])
		// {
		// 	let strDecimalLiteral = this.EatToken(this.oLookahead.class).data;
		// 	return {
		// 		type: oLiteralNames.opDecimal[num],
		// 		value: strDecimalLiteral
		// 	}
		// }
		else if (this.oLookahead.class === Literals.space.id)
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
		else if (this.IsMathFontLiteral())
		{
			return this.GetMathFontLiteral();
		}
		else if (this.oLookahead.data === "{")
		{
			return this.GetArguments(1)[0];
		}
		else if (this.oLookahead.class ===Literals.operator.id)
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
		else if (this.IsArrayLiteral())
		{
			return this.GetArrayLiteral();
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
				type: Struc.char,
				value: "/",
			}
		}
	};
	CLaTeXParser.prototype.IsHorizontalArrowLiteral = function()
	{
		return this.oLookahead.class === Literals.horizontal.id;
	}
	CLaTeXParser.prototype.GetHorizontalArrowLiteral = function ()
	{
		let oStyle = this.oLookahead.style;
		let oData = this.EatToken(this.oLookahead.class);
		oData.data = Literals.horizontal.LaTeX[oData.data];

		if (this.IsGetBelowAboveLiteral())
		{
			let isBelow = this.oLookahead.data === "\\below" ? VJUST_TOP :  VJUST_BOT;
			this.EatToken(this.oLookahead.class);
			let oContent = this.GetArguments(1);

			return {
				type: Struc.group_character,
				value: oContent,
				hBrack: oData,
				down: undefined,
				up: undefined,
				style: oStyle,
				isBelow: isBelow,
			}
		}

		return {
			type: Struc.char,
			value: oData.data,
			style: oStyle,
		}
	}
	CLaTeXParser.prototype.IsGetBelowAboveLiteral = function()
	{
		return this.oLookahead.data === "\\below" || this.oLookahead.data === "\\above";
	}
	CLaTeXParser.prototype.GetBelowAboveLiteral = function(base)
	{
		let oStyle = this.oLookahead.style;
		let isBelow = false;
		if (this.oLookahead.data === "\\above")
			isBelow = true;

		this.EatToken(this.oLookahead.class);
		let oContent = this.GetArguments(1);

		if(base && base.type === Struc.func)
		{
			this.SkipFreeSpace();
			let third = this.GetArguments(1);
			return {
				type: Struc.func_lim,
				value: {
					type: Struc.char,
					value: base.value,
					style: base.style
				},
				up: isBelow ? oContent : undefined,
				down: !isBelow ? oContent : undefined,
				third: third,
				style: base.style,
			}
		}

		let strNewBaseValue = AscMath.ConvertWord(base.value, true);
		if (strNewBaseValue)
			base.value = strNewBaseValue;

		return {
			type: Struc.limit,
			base: base,
			value: oContent,
			isBelow: (isBelow === false) ? VJUST_TOP :  VJUST_BOT,
			style: oStyle,
		}
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
			type: Struc.char,
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
		return this.oLookahead.class === Literals.func.id || this.oLookahead.class === Literals.nary.id
	};
	CLaTeXParser.prototype.GetFuncLiteral = function ()
	{
		let oFuncContent			= this.EatToken(this.oLookahead.class);
		let oPr						= oFuncContent.style;
		let oThirdStyle;
		let SetStyleAndGetArgument	= function (oThis)
		{
			oThirdStyle = oThis.oLookahead.style;
			return oThis.GetArguments(1);
		}

		if (this.oLookahead.class === "\\limits")
			this.EatToken("\\limits");

		if (this.oLookahead.data === " ")
			this.EatToken(this.oLookahead.class);

		let oThirdContent = !this.IsSubSup() && !this.IsGetBelowAboveLiteral()
			? SetStyleAndGetArgument(this)
			: undefined;

		let name = oFuncContent.data;

		if (oFuncContent.class === Literals.nary.id)
		{
			return {
				type: Struc.nary,
				value: Literals.nary.LaTeX[oFuncContent.data],
				style: oPr,
				third: oThirdContent,
				thirdStyle: oThirdStyle,
			}
		}
		else if (!oThirdContent)
		{
			return {
				type: Struc.func,
				value: oFuncContent.data.slice(1),
				style: oPr
			}
		}
		else if (AscMath.MathLiterals.func.IsLaTeXIncludeNormal(name))
		{
			return {
				type: Struc.func,
				value: {
					type: Struc.char,
					value: oFuncContent.data.slice(1),
					style: oPr
				},
				style: oPr,
				third: oThirdContent
			}
		}
		else if (AscMath.MathLiterals.func.IsLaTeXIncludeLimit(name))
		{
			return {
				type: Struc.func_lim,
				value: {
					type: Struc.char,
					value: oFuncContent.data.slice(1),
					style: oPr
				},
				style: oPr,
				third: oThirdContent
			}
		}
	};
	CLaTeXParser.prototype.IsReactLiteral = function ()
	{
		return this.oLookahead.class === Literals.rect.id;
	};
	CLaTeXParser.prototype.GetRectLiteral = function ()
	{
		let oCtrPr = this.oLookahead.style;
		this.EatToken(this.oLookahead.class);
		let oContent = this.GetArguments(1);
		return {
			type: Struc.rect,
			value: oContent,
			style: oCtrPr,
		}
	};
	CLaTeXParser.prototype.IsOverUnderBarLiteral = function ()
	{
		return this.oLookahead.data === "\\underline" || this.oLookahead.data === "\\overline"
	};
	CLaTeXParser.prototype.GetUnderOverBarLiteral = function ()
	{
		let oStyle				= this.oLookahead.style;
		let strUnderOverLine 	= this.EatToken(this.oLookahead.class);
		let oOperand			= this.GetArguments(1);

		strUnderOverLine.data = (strUnderOverLine.data === "\\underline")
			? "▁"
			: "¯";

		return {
			type: Struc.bar,
			bar: strUnderOverLine,
			value: oOperand,
			style: oStyle,
		};
	};
	CLaTeXParser.prototype.IsBoxLiteral = function ()
	{
		return this.oLookahead.class === Literals.box.id;
	}
	CLaTeXParser.prototype.GetBoxLiteral = function ()
	{
		this.EatToken(this.oLookahead.class);
		let oContent = this.GetArguments(1);
		return {
			type: Struc.box,
			value: oContent,
		}
	};
	CLaTeXParser.prototype.GetBorderBoxLiteral = function ()
	{
		return this.oLookahead.class === Literals.rect.id;
	}
	CLaTeXParser.prototype.IsGetBorderBoxLiteral = function ()
	{
		this.EatToken(this.oLookahead.class);
		let oContent = this.GetArguments(1);
		return {
			type: Struc.rect,
			value: oContent,
		}
	}
	CLaTeXParser.prototype.IsHBracket = function ()
	{
		return this.oLookahead.class === Literals.hbrack.id;
	};
	CLaTeXParser.prototype.GetHBracketLiteral = function ()
	{
		let oHBracket = this.oLookahead,
			oPr = this.oLookahead.style,
			oDown,
			oUp;
		oHBracket.data = Literals.hbrack.LaTeX[oHBracket.data];

		let isBelow = Literals.hbrack.GetPos(oHBracket.data);

		this.EatToken(this.oLookahead.class);

		let oContent = this.GetArguments(1);
		this.SkipFreeSpace();

		if (this.IsSubSup())
		{
			if (this.oLookahead.data === "_")
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
			type: Struc.group_character,
			value: oContent,
			hBrack: oHBracket,
			down: oDown,
			up: oUp,
			style: oPr,
			isBelow: isBelow,
		}
	};
	CLaTeXParser.prototype.GetWrapperElementLiteral = function ()
	{
		if (!this.IsSubSup() && this.oLookahead.class !== "\\over")
		{
			let oWrapperContent = this.GetElementLiteral();

			if (this.IsSubSup() || this.oLookahead.class === "\\limits")
			{
				while (this.IsSubSup())
				{
					oWrapperContent = this.GetSubSupLiteral(oWrapperContent);
				}
				return oWrapperContent;
			}
			else if (this.oLookahead.class === Literals.accent.id)
			{
				return this.GetAccentLiteral(oWrapperContent);
			}
			else if (this.IsGetBelowAboveLiteral())
			{
				return this.GetBelowAboveLiteral(oWrapperContent)
			}
			else if (this.IsOverLikeLiteral())
			{
				return this.GetOverLikeLiteral(oWrapperContent);
			}

			return oWrapperContent;
		}
	};
	CLaTeXParser.prototype.GetWrapperElement2 = function ()
	{
		let oWrapperContent = this.GetElementLiteral();

		if (this.oLookahead.class === Literals.accent.id)
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
		return this.oLookahead.data === "^" || this.oLookahead.data === "_";
	};
	CLaTeXParser.prototype.GetSubSupLiteral = function (oBaseContent, isSingle)
	{
		let isLimits,
			oDownContent,
			oUpContent,
			oThirdContent,
			oSubStyle,
			oSupStyle;

		if (undefined === oBaseContent)
		{
			oBaseContent = this.GetElementLiteral();
		}
		if (this.oLookahead.class === "\\limits")
		{
			this.EatToken("\\limits");
			isLimits = true;
		}

		if (oBaseContent && oBaseContent.type === Struc.bracket_block && oBaseContent.left === "{" && oBaseContent.right === "}")
		{
			oBaseContent = oBaseContent.value;
		}

		if (this.oLookahead.data === "'" || this.oLookahead.data === "''")
		{
			oUpContent = {
				type: Struc.char,
				value: this.EatToken(this.oLookahead.class).data
			}
		}

		if (this.oLookahead.data === "_")
		{
			oSubStyle = this.oLookahead.style;
			oDownContent = this.GetPartOfSupSup();
			if (this.oLookahead.data === "^" && isSingle !== true)
			{
				oSupStyle = this.oLookahead.style;
				oUpContent = this.GetPartOfSupSup();
			}
			else if (oDownContent && oDownContent.down === undefined && oDownContent.base)
			{
				oDownContent = oDownContent.base;
			}
		}
		else if (this.oLookahead.data === "^")
		{
			oSupStyle = this.oLookahead.style;
			oUpContent = this.GetPartOfSupSup();
			if (this.oLookahead.data === "_" && isSingle !== true)
			{
				oSubStyle = this.oLookahead.style;
				oDownContent = this.GetPartOfSupSup();
			}
			else if (oUpContent && oUpContent.up === undefined && oUpContent.base && oUpContent.type !== "BelowAboveLiteral")
			{
				oUpContent = oUpContent.base;
			}
		}

		if (oBaseContent && (oBaseContent.type === Struc.func || oBaseContent.type == Struc.func_lim || oBaseContent.type === Struc.nary))
		{
			this.SkipOneSpace();
			oThirdContent = this.GetArguments(1);
		}

		return {
			type: Struc.sub_sub,
			value: oBaseContent,
			up: oUpContent,
			down: oDownContent,
			third: oThirdContent,
			isLimits: isLimits,
			style: {supStyle: oSupStyle, subStyle: oSubStyle},
		};
	};
	CLaTeXParser.prototype.GetPartOfSupSup = function ()
	{
		let oElement;
		let strSymbol = this.oLookahead.class;
		this.EatToken(strSymbol);

		if (this.oLookahead.data === "'" || this.oLookahead.data === "''")
		{
			oElement = {
				type: Struc.char,
				value: this.EatToken(this.oLookahead.class).data
			}
		}
		else
		{
			if (this.oLookahead.data === "{")
			{
				oElement = this.GetArguments(1);

			}
			else
			{
				oElement = this.GetWrapperElement2();
			}
		}

		if (this.oLookahead.data === strSymbol) {
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

		if (this.oLookahead.data === "_") {
			oDownContent = this.GetPartOfSupSup();
			if (this.oLookahead.data === "^") {
				oUpContent = this.GetPartOfSupSup();
			}
		}
		else if (this.oLookahead.data === "^") {
			oUpContent = this.GetPartOfSupSup();
			if (this.oLookahead.data === "_") {
				oDownContent = this.GetPartOfSupSup();
			}
		}

		if (this.oLookahead.data === "}") {
			this.EatToken(this.oLookahead.class)
		}

		this.SkipFreeSpace();
		oBaseContent = this.GetElementLiteral();

		oOutput = { type: Struc.pre_script };

		if (oUpContent)
			oOutput.up = oUpContent;

		if (oDownContent)
			oOutput.down = oDownContent;

		if (oBaseContent)
			oOutput.value = oBaseContent;

		return oOutput;
	};
	CLaTeXParser.prototype.IsSqrtLiteral = function ()
	{
		return this.oLookahead.class === Literals.radical.id;
	};
	CLaTeXParser.prototype.GetSqrtLiteral = function ()
	{
		let oBaseContent,
			oIndexContent,
			oOutput,
			oStyle = this.oLookahead.style;

		this.EatToken(Literals.radical.id);
		if (this.oLookahead.data === "[")
		{
			this.EatToken(this.oLookahead.class);
			oIndexContent = this.GetExpressionLiteral(["]"]);
			if (this.oLookahead.data === "]")
			{
				this.EatToken(this.oLookahead.class);
			}
		}
		oBaseContent = this.GetArguments(1);
		oOutput = {
			type: Struc.radical,
			value: oBaseContent,
			style: oStyle,
		};

		if (oIndexContent)
			oOutput.index = oIndexContent;
		return oOutput;
	};
	CLaTeXParser.prototype.IsMathFontLiteral = function ()
	{
		return this.oLookahead.class === Literals.font.id;
	}
	CLaTeXParser.prototype.GetMathFontLiteral = function ()
	{
		let intPrevType = this.intMathFontType;
		this.intMathFontType = GetTypeFont[this.oLookahead.data];

		this.EatToken(this.oLookahead.class);
		let oOutput = this.GetArguments(1);

		this.intMathFontType = intPrevType;
		return oOutput;
	};
	CLaTeXParser.prototype.IsMatrixLiteral = function ()
	{
		return (this.oLookahead.class === Literals.matrix.id
			|| this.oLookahead.data === "\\substack")
			&& !this.IsEndMatrixLiteral()
	};
	CLaTeXParser.prototype.IsEndMatrixLiteral = function ()
	{
		return this.oLookahead.class === Literals.matrix.id
			&& Literals.matrix.LaTeX[this.oLookahead.data] === 2
			|| this.oLookahead.data === "}";
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

		if (oAlignBlock.type === Struc.char)
		{
			let strAlignBlock = oAlignBlock.value.trim();

			if (IsAlignContent(strAlignBlock))
				return strAlignBlock;
		}
		else
		{
			this.RestoreState();
		}
	};
	CLaTeXParser.prototype.IsArrayLiteral = function ()
	{
		return this.oLookahead.class === Literals.array.id && Literals.array.LaTeX[this.oLookahead.data] === 1
	}
	CLaTeXParser.prototype.IsEndArrayLiteral = function ()
	{
		return this.oLookahead.class === Literals.array.id && Literals.array.LaTeX[this.oLookahead.data] === 2
	}
	CLaTeXParser.prototype.GetArrayLiteral = function ()
	{
		let name	= this.EatToken(this.oLookahead.class).data,
			data;

		if (name !== "\\array{")
		{
			// set align in NamesOfLiterals
			let align	= this.IsAlignBlockForArray();
			data		= this.GetContentOfMatrixOrArray(function (oThis) {return oThis.IsEndArrayLiteral()});
		}
		else
		{
			data		= this.GetContentOfMatrixOrArray(function (oThis) {return oThis.oLookahead.data === "}"});
		}

		return {
			type:	Struc.array,
			value:	data.value,
			style:	data.style,
		}
	}

	CLaTeXParser.prototype.GetMatrixLiteral = function ()
	{
		let strMatrixType;

		switch (this.oLookahead.data)
		{
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
			case "\\begin{matrix}":
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

		this.SkipFreeSpace();

		let data = this.GetContentOfMatrixOrArray(function (oThis) {return oThis.IsEndMatrixLiteral()});

		this.isNowMatrix = false;

		return {
			type:	Struc.matrix,
			value:	data.value,
			style:	data.style,
			strMatrixType: strMatrixType,
		}
	};
	CLaTeXParser.prototype.GetContentOfMatrixOrArray = function (oFunc)
	{
		let arrMatrixContent	= [];
		let styles				= {};
		let nRow				= 0;

		styles.head				= this.oLookahead.style;
		styles.cols				= {};
		styles.rows				= {};

		while (this.oLookahead.class !== undefined && !oFunc(this))
		{
			let oContent = this.GetRayLiteral(styles.cols, styles.rows, nRow, oFunc)

			if (oContent)
				arrMatrixContent.push(oContent);
			else if (oContent === false)
				arrMatrixContent.push({}, {});
			else if (oContent === undefined)
				arrMatrixContent.push({});
			else if (this.IsEndMatrixLiteral())
				arrMatrixContent.push({}, {});

			nRow++;
		}

		let intMaxLengthOfMatrixRow		= -Infinity;
		let intIndexOfMaxMatrixRow		= -1;

		for (let i = 0; i < arrMatrixContent.length; i++)
		{
			let arrContent				= arrMatrixContent[i];
			intMaxLengthOfMatrixRow		= arrContent.length;
			intIndexOfMaxMatrixRow		= i;
		}

		for (let i = 0; i < arrMatrixContent.length; i++)
		{
			if (i !== intIndexOfMaxMatrixRow)
			{
				let arrMatrix = arrMatrixContent[i];
				for (let j = arrMatrix.length; j < intMaxLengthOfMatrixRow; j++)
				{
					arrMatrix.push({});
				}
			}
		}

		if (oFunc(this))
		{
			this.EatToken(this.oLookahead.class)
		}

		return {
			value: arrMatrixContent,
			style: styles,
		}
	};
	CLaTeXParser.prototype.GetRayLiteral = function (cols, rows, nRow, oFunc)
	{
		let arrRayContent;

		while (this.oLookahead.class !== Literals.arrayMatrix.id && !oFunc(this))
		{
			rows[nRow]		= {}
			arrRayContent	= this.GetElementOfMatrix(rows[nRow]);
			nRow++;
		}

		if (this.oLookahead.class === Literals.arrayMatrix.id)
		{
			cols[nRow]		= this.oLookahead.style;
			this.EatToken(this.oLookahead.class)

			if (!arrRayContent && this.IsEndMatrixLiteral())
				return false;
		}

		this.SkipFreeSpace();
		return arrRayContent
	};
	CLaTeXParser.prototype.GetElementOfMatrix = function (oStyle)
	{
		let arrRow				= [];
		let intLength			= 0;
		let intCount			= 0;
		let isAlreadyGetContent	= false;

		while (this.IsElementLiteral() || this.oLookahead.data === "&" )
		{
			let intCopyOfLength = intLength;

			if (this.oLookahead.data === "\\\\" || this.IsEndMatrixLiteral())
				break;

			if (this.oLookahead.data !== "&")
			{
				arrRow.push(this.GetExpressionLiteral(["&", "\\\\"]));
				intLength++;
				isAlreadyGetContent = true;
				this.SkipFreeSpace();
			}
			else
			{
				oStyle[intCount] = this.oLookahead.style;
				this.EatToken(this.oLookahead.class);

				if (isAlreadyGetContent === false)
				{
					arrRow.push({});
					intCount++;
					intLength++;
				}
				else if (intCopyOfLength === intLength)
				{
					intCount++;
				}

				this.SkipFreeSpace();
			}

		}

		if (intLength !== intCount + 1)
		{
			for (let j = intLength; j <= intCount; j++)
			{
				arrRow.push({});
			}
		}
		else if (arrRow.length === 1 && arrRow[0].length === 0)
		{
			arrRow.push([]);
		}

		return arrRow;
	};
	CLaTeXParser.prototype.IsExpressionLiteral = function(arrBreak)
	{
		const arrEndOfExpression = ["}", "\\endgroup", "\\end", "┤"];
		let isEndOfExp = !arrEndOfExpression.includes(this.oLookahead.data);

		let isEnd = false;

		if (Array.isArray(arrBreak))
			isEnd = !arrBreak.includes(this.oLookahead.data);
		else
			isEnd = this.oLookahead.data !== arrBreak;

		let isEscape = this.oLookahead.data !== this.EscapeSymbol

		return isEnd
			&& isEscape
			&& isEndOfExp
			&& this.IsElementLiteral();
	};
	CLaTeXParser.prototype.GetExpressionLiteral = function (arrBreakSymbol)
	{
		this.EscapeSymbol = arrBreakSymbol;
		const arrExpList = [];

		while (this.IsExpressionLiteral(arrBreakSymbol))
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
		if (tokenType !== undefined && this.oLookahead.class === tokenType) {
			const oToken = this.oLookahead;
			// if (oToken === null) {
			// 	console.log('Unexpected end of input, expected: ' + tokenType);
			// }
			// if (oToken.class !== tokenType) {
			// 	console.log('Unexpected token: ' + oToken.class + ', expected: ' + tokenType);
			// }
			this.oLookahead = this.oTokenizer.GetNextToken();
			return oToken;
		}
	};
	CLaTeXParser.prototype.SkipFreeSpace = function ()
	{
		while (this.oLookahead.data === " ") {
			this.oLookahead = this.oTokenizer.GetNextToken();
		}
	};
	CLaTeXParser.prototype.SkipOneSpace = function ()
	{
		if (this.oLookahead.data === " ") {
			this.oLookahead = this.oTokenizer.GetNextToken();
		}
	}
	CLaTeXParser.prototype.GetArguments = function (intCountOfArguments)
	{
		let oArgument = [];
		while (intCountOfArguments > 0) {
			this.SkipFreeSpace();
			if (this.oLookahead.data === "{") {
				this.SkipFreeSpace();
				this.EatToken(this.oLookahead.class);
				oArgument.push(this.GetExpressionLiteral());
				this.EatToken(this.oLookahead.class);
			}
			else {
				this.SkipFreeSpace();
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

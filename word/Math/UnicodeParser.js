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

(function (window)
{
	const Literals				= AscMath.MathLiterals;
	const Struc					= AscMath.MathStructures;
	const oLiteralNames			= AscMath.oNamesOfLiterals;
	const UnicodeSpecialScript	= AscMath.UnicodeSpecialScript;
	const ConvertTokens			= AscMath.ConvertTokens;
	const Tokenizer				= AscMath.Tokenizer;
	const FunctionNames			= AscMath.functionNames;
	const LimitNames			= AscMath.LimitFunctions;
	const LRBracketType			= { open: 1, close: 0};

	function CUnicodeParser() {
		this.oTokenizer			= new Tokenizer(false);
		this.isOneSubSup		= false;
		this.isTextLiteral		= false;
		this.arrSavedTokens		= [];
		this.isSaveTokens		= false;
		this.isSpaceExit		= false;

		this.prevLookahead		= null;
		this.bracketStack		= [];
		this.barStack			= [];
		this.nBarCount			= 0;
	}
	CUnicodeParser.prototype.GetSpaceExitFunction = function (oFunc, oArg)
	{
		this.isSpaceExit = true;
		let oContent = oFunc.call(this, oArg);
		this.isSpaceExit = false;

		return oContent;
	}
	CUnicodeParser.prototype.IsSpaceExit = function ()
	{
		// if (this.oLookahead.class !== Literals.space.id || !this.isSpaceExit)
		// 	return true;

		let isSpace = this.isSpaceExit
			&& (this.oLookahead.class === Literals.space.id || this.oLookahead.class === Literals.invisible.id);

		return !isSpace;
	}
	CUnicodeParser.prototype.Parse = function (string)
	{
		this.oTokenizer.Init(string);
		this.prevLookahead	= null;
		this.oLookahead		= this.oTokenizer.GetNextToken();
		return this.Program();
	};
	CUnicodeParser.prototype.Program = function ()
	{
		const arrExp = [];
		while (this.oLookahead.data)
		{
			if (this.IsExpLiteral())
				arrExp.push(this.GetExpLiteral());
			else
				this.WriteDataAsCharLiteral(arrExp);
		}

		return { type: "UnicodeEquation", body: arrExp};
	};
	CUnicodeParser.prototype.WriteDataAsCharLiteral = function(arrExp)
	{
		let newData = {
			type: Literals.char.id,
			value: this.oLookahead.data,
			style: this.oLookahead.style
		};

		arrExp.push(newData);
		this.EatToken(this.oLookahead.class);
	};
	CUnicodeParser.prototype.GetSpaceLiteral = function ()
	{
		const oSpaceLiteral = this.EatToken(Literals.space.id);
		return {
			type: Struc.space,
			value: oSpaceLiteral.data,
			style: oSpaceLiteral.style,
		};
	};
	CUnicodeParser.prototype.IsSpaceLiteral = function ()
	{
		return this.oLookahead.class === Literals.space.id;
	};
	CUnicodeParser.prototype.GetOpCloseLiteral = function ()
	{
		let oCloseLiteral;

		if (this.oLookahead.data === "┤") //return this.EatToken(this.oLookahead.class).data
			this.EatToken(this.oLookahead.class);

		if (this.oLookahead.class === Literals.lrBrackets.id) {
			oCloseLiteral = this.EatToken(Literals.lrBrackets.id);
			return oCloseLiteral.data;
		}

		oCloseLiteral = this.EatToken(Literals.rBrackets.id);
		return oCloseLiteral.data;
	};
	// CUnicodeParser.prototype.GetOpCloserLiteral = function ()
	// {
	// 	switch (this.oLookahead.class) {
	// 		case "\\close":
	// 			return {
	// 				type: oLiteralNames.opCloseBracket[num],
	// 				value: this.EatToken("\\close").data,
	// 			};
	// 		case "┤":
	// 			return {
	// 				type: oLiteralNames.opCloseBracket[num],
	// 				value: this.EatToken("┤").data,
	// 			};
	// 		case Literals.rBrackets.id:
	// 			return this.GetOpCloseLiteral();
	// 		case Literals.lrBrackets.id:
	// 			return this.GetOpCloseLiteral();
	// 	}
	// };
	CUnicodeParser.prototype.IsOpNaryLiteral = function ()
	{
		return this.oLookahead.class === Literals.nary.id;
	};
	CUnicodeParser.prototype.GetOpNaryLiteral = function ()
	{
		let oContent;
		let oNaryPr = this.oLookahead.style;
		let oOfStyle;
		let strNaryLiteral = this.EatToken(Literals.nary.id).data;

		if (this.oLookahead.class === Literals.of.id)
		{
			oOfStyle = this.oLookahead.style;
			this.EatToken( Literals.of.id);

			if (this.oLookahead.data === "〖" )
				oContent = this.GetExpBracketLiteral();
			else
				oContent = this.GetExpLiteral();

			if (oContent.type === Struc.bracket_block && oContent.left === "(" && oContent.right === ")")
			{
				oContent = oContent.value;
			}
		}
		// else if (this.oLookahead.class !== Literals.subSup.id)
		// {
		// 	oContent = this.GetElementLiteral()
		// }

		return {
			type: Struc.nary,
			value: strNaryLiteral,
			third: oContent,
			thirdStyle: oOfStyle,
			style: oNaryPr,
		}

	};
	CUnicodeParser.prototype.GetOpOpenLiteral = function ()
	{
		let oOpLiteral;
		if (this.oLookahead.class === Literals.lrBrackets.id) {
			oOpLiteral = this.EatToken(Literals.lrBrackets.id);
			return oOpLiteral.data;
		}
		oOpLiteral = this.EatToken(Literals.lBrackets.id);
		return oOpLiteral.data;
	};
	CUnicodeParser.prototype.IsOpOpenLiteral = function ()
	{
		return this.oLookahead.class === Literals.lrBrackets.id ||
				this.oLookahead.class === Literals.lBrackets.id;
	};
	CUnicodeParser.prototype.IsOpOpenerLiteral = function ()
	{
		return this.oLookahead.class === Literals.lBrackets.id;
	};
	CUnicodeParser.prototype.GetDigitsLiteral = function ()
	{
		return this.ReadTokensWhileEnd(Literals.char, Struc.char, false)
		// const arrNASCIIList = [this.GetASCIILiteral()];
		// while (this.oLookahead.class === "nASCII") {
		// 	arrNASCIIList.push(this.GetASCIILiteral());
		// }
		// return this.GetContentOfLiteral(arrNASCIIList);
	};
	CUnicodeParser.prototype.IsDigitsLiteral = function ()
	{
		return this.oLookahead.class === Literals.number.id;
	};
	CUnicodeParser.prototype.GetNumberLiteral = function ()
	{
		return this.GetDigitsLiteral();
	};
	CUnicodeParser.prototype.IsNumberLiteral = function ()
	{
		return this.IsDigitsLiteral();
	};
	CUnicodeParser.prototype.EatCloseOrOpenBracket = function ()
	{
		let strOpenLiteral,
			strCloseLiteral,
			oExp,
			startStyle,
			endStyle;

		if (this.oLookahead.data === "├") {
			this.strBreakSymbol.push("|", "‖");
			this.EatToken("├");

			startStyle = this.oLookahead.style;

			if (this.oLookahead.class === Literals.lrBrackets.id || this.oLookahead.class === Literals.lBrackets.id || this.oLookahead.class === Literals.rBrackets.id)
				strOpenLiteral = this.EatBracket().data;
			else
				strOpenLiteral = ".";

			let arrContent = this.GetContentOfBracket();
			oExp = arrContent[0];
			let counter = arrContent[1];

			if (this.oLookahead.data === "┤")
			{
				endStyle = this.oLookahead.style;
				this.EatToken("┤");
			}

			if (this.oLookahead.class === Literals.lrBrackets.id || this.oLookahead.class === Literals.rBrackets.id || this.oLookahead.class === Literals.lBrackets.id)
			{
				endStyle = this.oLookahead.style;
				strCloseLiteral = this.EatBracket().data;
			}
			else
				strCloseLiteral = ".";

			return {
				type: Struc.bracket_block,
				left: strOpenLiteral,
				right: strCloseLiteral,
				value: oExp,
				counter: counter,
				style: {
					startStyle : startStyle,
					endStyle : endStyle,
					middle: [],
				}
			};
		}
	};
	CUnicodeParser.prototype.EatBracket = function ()
	{
		return this.EatToken(this.oLookahead.class).data;
	};
	CUnicodeParser.prototype.GetWordLiteral = function ()
	{
		const arrWordList = [this.GetASCIILiteral()];
		while (this.oLookahead.class === oLiteralNames.asciiLiteral[0]) {
			arrWordList.push(this.GetASCIILiteral());
		}
		return {
			type: Struc.char,
			value: this.GetContentOfLiteral(arrWordList),
		};
	};
	CUnicodeParser.prototype.IsWordLiteral = function ()
	{
		return this.oLookahead.class === Literals.char.id// oLiteralNames.asciiLiteral[0];
	};
	CUnicodeParser.prototype.GetSoOperandLiteral = function (isSubSup)
	{
		if (this.IsOperandLiteral()) {
			return this.GetOperandLiteral(isSubSup);
		}
		else if (this.IsDoubleIteratorDegree())
		{
			let data = this.oLookahead.data;
			this.EatToken(this.oLookahead.class);
			return {
				type: Struc.char,
				value: data,
			}
		}

		switch (this.oLookahead.data) {
			case "-":
				let minus = this.EatToken(Literals.operator.id);
				if (this.IsOperandLiteral()) {
					const operand = this.GetOperandLiteral();
					return {
						type: Struc.minus,
						value: operand,
					};
				}

				return {
					type: Struc.char,
					value: minus.data,
				}
			case "-∞":
				const token = this.EatToken(Literals.operator.id);
				return token.data;
			case "∞":
				const tokens = this.EatToken(Literals.operator.id);
				return tokens.data;
		}

		if (this.oLookahead.class === Literals.operator.id) {
			let one = this.GetOperandLiteral(isSubSup);
			return one;
		}

	};
	CUnicodeParser.prototype.IsSoOperandLiteral = function ()
	{
		return this.IsOperandLiteral() ||
			this.oLookahead.data === "-" ||
			this.oLookahead.data === "-∞" ||
			this.oLookahead.data === "∞" ||
			this.IsDoubleIteratorDegree();
	};
	CUnicodeParser.prototype.IsTextLiteral = function ()
	{
		// only \" for text
		return this.oLookahead.data === "\"" && !this.isTextLiteral
	};
	CUnicodeParser.prototype.GetTextLiteral = function ()
	{
		let arrStyles = [];
		let strSymbol = this.EatToken(this.oLookahead.class);
		let strExp = "";

		while (this.oLookahead.data !== "\"" && this.oLookahead.class !== undefined)
		{
			strExp += this.oLookahead.data;
			arrStyles.push(this.oLookahead.style);
			this.EatToken(this.oLookahead.class)
		}

		if (strExp === "" && this.oLookahead.data !== "\"")
		{
			return {
				type: Struc.char,
				value: strSymbol.data,
				style: strSymbol.style,
			}
		}

		if (this.oLookahead.data === "\"")
			this.EatToken(this.oLookahead.class);

		return {
			type: Struc.char,
			value: strExp,
			style: arrStyles,
		}
	}
	CUnicodeParser.prototype.IsBoxLiteral = function ()
	{
		return this.oLookahead.data === "□";
	};
	CUnicodeParser.prototype.GetBoxLiteral = function ()
	{
		this.SaveTokensWhileReturn();
		if (this.oLookahead.data === "□")
		{
			let ctrlPr = this.oLookahead.style;
			this.EatToken(this.oLookahead.class);
			if (this.IsOperandLiteral())
			{
				const oToken = this.GetOperandLiteral();
				return {
					type: Struc.box,
					value: oToken,
					style: ctrlPr,
				};
			}
			else
			{
				return {
					type: Struc.box,
					value: {},
					style: ctrlPr,
				};
			}
		}
		return this.WriteSavedTokens();
	};
	CUnicodeParser.prototype.isRectLiteral = function ()
	{
		return this.oLookahead.class === Literals.rect.id;
	};
	CUnicodeParser.prototype.GetRectLiteral = function ()
	{
		this.SaveTokensWhileReturn();

		if (this.oLookahead.data === "▭")
		{
			let oCtrPr = this.oLookahead.style;
			this.EatToken(this.oLookahead.class);
			if (this.IsOperandLiteral())
			{
				const oToken = this.GetOperandLiteral();
				return {
					type: Struc.rect,
					value: oToken,
					style: oCtrPr,
				};
			}
			else
			{
				return {
					type: Struc.rect,
					value: {},
					style: oCtrPr,
				};
			}
		}

		return this.WriteSavedTokens();
	};
	//Custom horizontal (1+2)\underbrace2
	CUnicodeParser.prototype.GetSpecialHBracket = function (oBase)
	{
		let strHBracket = this.oLookahead;
		this.EatToken(Literals.hbrack.id);

		let oPos = Literals.hbrack.GetPos(strHBracket),
			oOperand = this.GetOperandLiteral("custom"),
			oUp,
			oDown;

		if (oPos === VJUST_BOT)
			oDown = oOperand;
		else
			oUp = oOperand;

		return {
			type: Struc.group_character,
			hBrack: strHBracket,
			value: oBase,
			up: oUp,
			down: oDown,
			isBelow: oPos,
		};
	};
	CUnicodeParser.prototype.GetHBracketLiteral = function (oBase)
	{
		let strHBracket		= this.oLookahead,
			nPos			= Literals.hbrack.GetPos(strHBracket.data),
			oPr				= this.oLookahead.style,
			oUp,
			oDown;

		this.EatToken(this.oLookahead.class);

		if (this.IsOperandLiteral())
		{
			if (!oBase)
				oBase = this.GetOperandLiteral("custom");

			if (this.IsScriptStandardContentLiteral())
			{
				if (this.oLookahead.data === "_")
				{
					this.EatToken(this.oLookahead.class);
					oDown = this.GetSoOperandLiteral();
				}
				else if (this.oLookahead.data === "^")
				{
					this.EatToken(this.oLookahead.class);
					oUp = this.GetSoOperandLiteral();
				}
			}
			// else if (this.IsScriptBelowOrAboveContent())
			// {
			// 	if (this.oLookahead.data === "┬")
			// 	{
			// 		this.EatToken();
			// 		oDown = this.GetElementLiteral();
			// 	}
			// 	else if (this.oLookahead.data === "┴")
			// 	{
			// 		this.EatToken();
			// 		oUp = this.GetElementLiteral();
			// 	}
			// }
			else if (this.IsOperandLiteral())
				oDown = this.GetOperandLiteral("custom");
		}

		return {
			type: Struc.group_character,
			hBrack: strHBracket,
			value: oBase,
			up: oUp,
			down: oDown,
			style: oPr,
			isBelow : nPos,
		};
	};
	CUnicodeParser.prototype.IsHBracketLiteral = function ()
	{
		return this.oLookahead.class === Literals.hbrack.id;
	};
	CUnicodeParser.prototype.IsRootLiteral = function ()
	{
		return this.oLookahead.data === "⒭";
	};
	CUnicodeParser.prototype.GetRootLiteral = function ()
	{
		let oRadicalStyle = this.oLookahead.style;
		this.EatToken(this.oLookahead.class);
		let oIndex = this.GetExpLiteral();
		let oBase;
		if (this.oLookahead.data === Literals.of.id) {
			this.EatToken(this.oLookahead.class);
			oBase = this.GetExpLiteral();
		}
		return {
			type: Struc.radical,
			index: oIndex,
			value: oBase,
			style: oRadicalStyle,
		}
	};
	CUnicodeParser.prototype.GetCubertLiteral = function ()
	{
		let oRadicalStyle = this.oLookahead.style;
		this.EatToken(this.oLookahead.class);

		return this.GetContentOfAnyTypeRadical({
			type: Struc.char,
			value: "3",
			style: oRadicalStyle,
		}, oRadicalStyle);
	};
	CUnicodeParser.prototype.IsCubertLiteral = function ()
	{
		return this.oLookahead.data === "∛" && this.oLookahead.class !== Literals.operator.id;
	};
	CUnicodeParser.prototype.GetFourthrtLiteral = function ()
	{
		let oRadicalStyle = this.oLookahead.style;
		this.EatToken(Literals.radical.id);

		return this.GetContentOfAnyTypeRadical({
			type: Struc.char,
			value: "4",
			style: oRadicalStyle,
		}, oRadicalStyle);
	};
	CUnicodeParser.prototype.IsFourthrtLiteral = function ()
	{
		return this.oLookahead.data === "∜" && this.oLookahead.class !== Literals.operator.id;
	};
	CUnicodeParser.prototype.GetNthrtLiteral = function ()
	{
		let oPr = this.oLookahead.style;
		this.EatToken(this.oLookahead.class);
		return this.GetContentOfAnyTypeRadical(undefined, oPr);
	};
	CUnicodeParser.prototype.IsNthrtLiteral = function ()
	{
		return this.oLookahead.data === "√" && this.oLookahead.class !== Literals.operator.id || this.oLookahead.data === "√(";
	};
	CUnicodeParser.prototype.GetContentOfAnyTypeRadical = function(index, oRadicalStyle)
	{
		let oIndex, oContent;

		if (this.IsOpOpenLiteral()) {
			this.GetOpOpenLiteral();

			if (this.IsOperandLiteral())
			{
				oIndex = this.GetExpLiteral(undefined, true);

				if (this.oLookahead.data === "&")
				{
					this.EatToken(this.oLookahead.class);

					if (this.IsOperandLiteral())
						oContent = this.GetExpLiteral();
				}
				else
				{
					oContent = oIndex;
					oIndex = undefined;
				}
			}

			if (this.oLookahead.class === Literals.rBrackets.id)
				this.EatToken(Literals.rBrackets.id);
		}
		else if (this.IsOperandLiteral())
		{
			oContent = this.GetOperandLiteral();
		}

		return {
			type: Struc.radical,
			index: index ? index : oIndex,
			value: oContent,
			style: oRadicalStyle
		};
	}
	CUnicodeParser.prototype.IsFunctionLiteral = function ()
	{
		return (
			this.IsRootLiteral() ||
			this.IsCubertLiteral() ||
			this.IsFourthrtLiteral() ||
			this.IsNthrtLiteral() ||
			this.IsBoxLiteral() ||
			this.isRectLiteral() ||
			this.IsHBracketLiteral()    ||
			this.IsStretchArrow()       ||
			this.IsGetNameOfFunction() ||
			this.oLookahead.class === "▁" ||
			this.oLookahead.class === "¯"
		);
	};
	CUnicodeParser.prototype.GetFunctionLiteral = function ()
	{
		let oFunctionContent;

		if (this.IsRootLiteral())
			oFunctionContent = this.GetRootLiteral();
		else if (this.IsCubertLiteral())
			oFunctionContent = this.GetCubertLiteral();
		else if (this.IsFourthrtLiteral())
			oFunctionContent = this.GetFourthrtLiteral();
		else if (this.IsNthrtLiteral())
			oFunctionContent = this.GetNthrtLiteral();
		else if (this.IsBoxLiteral())
			oFunctionContent = this.GetBoxLiteral();
		else if (this.isRectLiteral())
			oFunctionContent = this.GetRectLiteral();
		else if (this.IsBarLiteral())
			oFunctionContent = this.GetBarLiteral();
		else if (this.IsHBracketLiteral())
			oFunctionContent = this.GetHBracketLiteral();
		else if (this.IsStretchArrow())
			oFunctionContent = this.GetStretchArrow();
		else if (this.IsGetNameOfFunction())
			oFunctionContent = this.GetNameOfFunction();

		return oFunctionContent;
	};
	CUnicodeParser.prototype.IsBarLiteral = function ()
	{
		return this.oLookahead.data === "▁" || this.oLookahead.data === "¯";
	}
	CUnicodeParser.prototype.GetBarLiteral = function ()
	{
		let oStyle				= this.oLookahead.style;
		let strUnderOverLine	= this.EatToken(this.oLookahead.class);
		strUnderOverLine.class	= Struc.char;
		let oOperand			= this.GetOperandLiteral("custom");

		return {
			type: Struc.bar,
			bar: strUnderOverLine,
			value: oOperand,
			style: oStyle,
		};
	}
	CUnicodeParser.prototype.IsStretchArrow = function ()
	{
		return this.oLookahead.class === Literals.horizontal.id;
	}
	CUnicodeParser.prototype.GetStretchArrow = function ()
	{
		// todo
		let oPr = this.oLookahead.style;
		let data = this.EatToken(this.oLookahead.class);
		return {
			type: Struc.char, //Struc.horizontal,
			value: data.data,
			style: oPr,
		}
	}
	CUnicodeParser.prototype.IsFuncApplySymbol = function ()
	{
		return	this.oLookahead.data &&
				this.oLookahead.data.length === 1 &&
				this.oLookahead.data.charCodeAt(0) === 8289; //funcapply symbol ⁡
	}
	CUnicodeParser.prototype.IsGetNameOfFunction = function ()
	{
		return this.oLookahead.class === Literals.func.id;
	};
	CUnicodeParser.prototype.GetNameOfFunction = function ()
	{
		let oName = this.EatToken(this.oLookahead.class);
		let oFunctionName = {
			type: Struc.char,
			value: oName.data,
			style: oName.style,
		}

		if (this.IsApplicationFunction())
			return this.GetFunctionApplication(oFunctionName);

		return oFunctionName;
	};
	CUnicodeParser.prototype.IsBracketLiteral = function ()
	{
		return Literals.rBrackets.SearchU(this.oLookahead.data)
		|| Literals.lBrackets.SearchU(this.oLookahead.data)
		|| Literals.lrBrackets.SearchU(this.oLookahead.data)
	}
	CUnicodeParser.prototype.IsOpCloserLiteral = function()
	{
		return this.oLookahead.class === Literals.rBrackets.id || this.oLookahead.class === Literals.lrBrackets.id || this.oLookahead.data === "┤" || this.oLookahead.data === "\\right"
	}
	CUnicodeParser.prototype.IsExpBracketLiteral = function ()
	{
		return (this.oLookahead.class === Literals.lBrackets.id
			|| this.oLookahead.class === Literals.lrBrackets.id
			|| this.oLookahead.data === "├"
			|| this.oLookahead.data === "\\left"
			|| this.oLookahead.data === "\\open") && !this.oLookahead.isClose;
	};
	CUnicodeParser.prototype.GetTypeOfLastLRInStack = function ()
	{
		if (this.bracketStack.length > 0)
			return this.bracketStack[this.bracketStack.length - 1];

		return false;
	}
	CUnicodeParser.prototype.GetExpBracketLiteral = function ()
	{
		let strOpen,
			strClose,
			startStyle,
			endStyle,
			oExp;

		let oNextLookahead = this.oTokenizer.GetNextToken(true);

		if (this.oLookahead.data === "|" || this.oLookahead.data === "‖")
		{
			let oLastType = this.GetTypeOfLastLRInStack();

			if (this.nBarCount % 2 === 0)
			{
				if ((oLastType === false
					|| this.prevLookahead.class === Literals.operator.id
					|| this.prevLookahead.class === Literals.lBrackets.id
					|| this.prevLookahead.class === Literals.special.id
					|| this.prevLookahead.class === Literals.char.id
					|| this.prevLookahead.data === " ")
					&& oNextLookahead.class !== Literals.operator.id
					&& oNextLookahead.class !== undefined
					)
				{
					this.bracketStack.push(LRBracketType.open);
					this.barStack.push(this.nBarCount);
					this.nBarCount++;
				}
				else
				{
					this.bracketStack.pop();
					this.oLookahead.isClose = true;
					this.nBarCount++;
					return [];
				}
			}
			else
			{
				if (oNextLookahead.class === undefined
					|| oNextLookahead.class === Literals.rBrackets.id
					|| oNextLookahead.class === Literals.operator.id
					|| oNextLookahead.class === Literals.char.id
					)
				{
					this.bracketStack.pop();
					this.oLookahead.isClose = true;
					this.nBarCount++;
					return [];
				}
				else
				{
					this.bracketStack.push(LRBracketType.open);
					this.barStack.push(this.nBarCount);
					this.nBarCount++;
				}
			}
		}
		else if (this.oLookahead.class === Literals.lBrackets.id)
		{
			this.bracketStack.push(LRBracketType.open);
			this.barStack.push(this.nBarCount);
			this.nBarCount = 0;
		}
		else if (this.oLookahead.class === Literals.rBrackets.id)
		{
			let oLastType = this.GetTypeOfLastLRInStack();


			if (oLastType === LRBracketType.open)
			{
				this.bracketStack.pop();
				this.oLookahead.isClose = true;

				if (this.barStack.length > 0)
					this.nBarCount = this.barStack.pop();

				return [];
			}
		}

		if (this.oLookahead.data === "├" || this.oLookahead.data === "\\left" || this.oLookahead.data === "\\open")
		{
			startStyle	= this.oLookahead.style;
			strOpen		= this.EatToken(this.oLookahead.class).data;

			if (this.IsBracketLiteral())
			{
				strOpen		= this.oLookahead.data;
				this.EatToken(this.oLookahead.class);
			}
			else if (strOpen !== "\\left" && strOpen !== "\\open")
			{
				strOpen		= ".";
			}
		}
		else if (this.IsBracketLiteral())
		{
			startStyle		= this.oLookahead.style;
			strOpen			= this.GetOpOpenLiteral();
		}

		if (this.IsPreScriptLiteral() && strOpen === "(")
		{
			return this.GetPreScriptLiteral(strOpen, true);
		}

		let isTempSpaceExit = undefined;
		if (this.isSpaceExit === true)
		{
			isTempSpaceExit = this.isSpaceExit;
			this.isSpaceExit = false;
		}

		let arrContent		= this.GetContentOfBracket();

		oExp				= arrContent[0];
		let counter			= arrContent[1];
		let middle_styles	= arrContent[2];

		if (isTempSpaceExit !== undefined)
			this.isSpaceExit = isTempSpaceExit;

		if (oExp.length === 0 && !this.IsOpCloserLiteral())
		{
			return {
				type: Struc.char,
				value: strOpen,
				style: startStyle
			}
		}

		if (this.oLookahead.data === "┤" || this.oLookahead.data === "\\right" || this.oLookahead.data === "\\close")
		{
			this.EatToken(this.oLookahead.class);

			if (this.IsBracketLiteral())
			{
				strClose	= this.oLookahead.data;
				endStyle	= this.oLookahead.style;

				this.EatToken(this.oLookahead.class);
			}
			else
			{
				strClose	= ".";
			}
		}
		else if (this.IsBracketLiteral())
		{
			endStyle	= this.oLookahead.style;
			strClose	= this.GetOpCloseLiteral();
		}

		if (!strClose)
		{
			return [{
				type: Struc.char,
				value: strOpen,
				style: startStyle,
			}, oExp]
		}

		if (strOpen === "〖" && strClose === "〗" && counter === 1)
			return oExp;

		return {
			type: Struc.bracket_block,
			value: oExp,
			left: strOpen,
			right: strClose,
			counter: counter,
			style: {
				startStyle : startStyle,
				endStyle : endStyle,
				middle: middle_styles,
			},
		};
	};
	CUnicodeParser.prototype.GetContentOfBracket = function ()
	{
		let arrContent = [];
		let intCountOfBracketBlock = 1;
		let styles = [];

		while (this.IsExpLiteral() || this.oLookahead.class === Literals.delimiter.id || this.oLookahead.data === "ⓜ" && this.oLookahead.class !== Literals.lrBrackets.id)
		{
			if (this.IsExpLiteral())
			{
				let oToken = this.GetExpLiteral();
				if (oToken && !Array.isArray(oToken) || (Array.isArray(oToken) && oToken.length > 0))
					arrContent.push(oToken)
			}
			else
			{
				if (arrContent.length === 0)
				{
					arrContent.push({});
				}

				styles.push(this.oLookahead.style);
				this.EatToken(this.oLookahead.class);

				if (!this.IsExpLiteral())
				{
					arrContent.push({});
				}

				intCountOfBracketBlock++;
			}
		}

		return [arrContent, intCountOfBracketBlock, styles];
	}
	CUnicodeParser.prototype.GetPreScriptLiteral = function (strOpen, isInsideBracket)
	{
		let oFirstSoOperand,
			oSecondSoOperand,
			oBase,
			isApostrope = false,
			oSubStyle,
			oSupStyle;

		let strTypeOfPreScript = this.oLookahead.data;

		if (strTypeOfPreScript === "_")
		{
			oSubStyle = this.oLookahead.style;
		}
		else
		{
			oSupStyle = this.oLookahead.style;
		}

		this.SaveTokensWhileReturn();
		if (this.oLookahead.data === "_" || this.oLookahead.data === "^")
			this.EatToken(this.oLookahead.class);

		if (this.IsDoubleIteratorDegree())
		{
			oSecondSoOperand = this.oLookahead.data;
			this.EatToken(this.oLookahead.class);
			isApostrope = true;
		}
		else if (strTypeOfPreScript === "_")
		{
			oFirstSoOperand = this.GetSpaceExitFunction(this.GetSoOperandLiteral, "preScript");
		}
		else
		{
			oSecondSoOperand = this.GetSpaceExitFunction(this.GetSoOperandLiteral, "preScript");
		}

		if (!oFirstSoOperand && !oSecondSoOperand)
		{
			return {
				type: Struc.spaces,
				value: {},
				down: strTypeOfPreScript === "_" ? {} : undefined,
				up: strTypeOfPreScript === "^" ? {} : undefined,
				style: {supStyle: oSupStyle, subStyle: oSubStyle},
			}
		}

		if (this.oLookahead.data !== strTypeOfPreScript && this.IsPreScriptLiteral())
		{
			this.EatToken(this.oLookahead.class);
			if (strTypeOfPreScript === "_")
				oSecondSoOperand = this.GetSpaceExitFunction(this.GetSoOperandLiteral, "preScript");
			else
				oFirstSoOperand = this.GetSpaceExitFunction(this.GetSoOperandLiteral, "preScript");
		}

		let isBracket = false;

		if (this.oLookahead.class === Literals.lrBrackets.id)
		{
			this.EatToken(Literals.lrBrackets.id);
			isBracket = true;
		}
		else if (this.oLookahead.class === Literals.rBrackets.id)
		{
			this.EatToken(Literals.rBrackets.id);
			isBracket = true;
		}

		if (this.oLookahead.class === Literals.space.id || isBracket || isApostrope)
		{
			if (!isBracket && !isApostrope)
				this.EatToken(this.oLookahead.class);

			oBase = this.IsElementLiteral()
				? this.GetSpaceExitFunction(this.GetElementLiteral)
				: {};

			if (this.oLookahead.class === Literals.space.id)
				this.EatToken(this.oLookahead.class);

			return {
				type: Struc.pre_script,
				value: oBase,
				down: oFirstSoOperand,
				up: oSecondSoOperand,
				style: {supStyle: oSupStyle, subStyle: oSubStyle},
			}
		}

		if (!isInsideBracket)
		{
			return {
				type: Struc.pre_script,
				value: {
					type: Struc.char,
					value: ''
				},
				down: oFirstSoOperand,
				up: oSecondSoOperand,
				style: {supStyle: oSupStyle, subStyle: oSubStyle},
			}
		}

		return this.WriteSavedTokens();
	};
	CUnicodeParser.prototype.IsPreScriptLiteral = function ()
	{
		return (this.oLookahead.data === "_" || this.oLookahead.data === "^" || this.IsDoubleIteratorDegree())
	};
	CUnicodeParser.prototype.GetScriptBaseLiteral = function ()
	{
		if (this.IsWordLiteral()) {
			let token = this.GetWordLiteral();
			if (this.oLookahead.class === Literals.number.id) {
				token.nASCII = this.GetASCIILiteral();
			}
			return token;
		}
		// else if (this.oLookahead.class === oLiteralNames.anMathLiteral[0]) {
		// 	return this.GetAnMathLiteral();
		// }
		else if (this.IsNumberLiteral()) {
			return this.GetNumberLiteral();
		}
		else if (this.isOtherLiteral()) {
			return this.GetOtherLiteral();
		}
		else if (this.IsExpBracketLiteral()) {
			return this.GetExpBracketLiteral();
		}
		// else if (this.oLookahead.class === oLiteralNames.opBuildupLiteral[0]) {
		// 	return this.GetOpNaryLiteral();
		// }
		else if (this.IsAnOtherLiteral()) {
			return this.GetAnOtherLiteral();
		}
		else if (this.oLookahead.class === Literals.char.id) {
			return this.GetCharLiteral();
		}
		else if (this.IsCubertLiteral()) {
			return this.GetCubertLiteral();
		}
		else if (this.IsFourthrtLiteral()) {
			return this.GetFourthrtLiteral();
		}
		else if (this.IsNthrtLiteral()) {
			return this.GetNthrtLiteral();
		}
	};
	// CUnicodeParser.prototype.IsScriptBaseLiteral = function () {
	// 	return (
	// 		this.IsWordLiteral() ||
	// 		this.IsNumberLiteral() ||
	// 		this.isOtherLiteral() ||
	// 		this.IsExpBracketLiteral() ||
	// 		this.oLookahead.class === "anOther" ||
	// 		this.oLookahead.class === Literals.nary.id ||
	// 		this.oLookahead.class === "┬" ||
	// 		this.oLookahead.class === "┴" ||
	// 		this.oLookahead.class === Literals.char.id ||
	// 		this.oLookahead.class === oLiteralNames.anMathLiteral[0]
	// 	);
	// };
	// CUnicodeParser.prototype.GetScriptSpecialContent = function (base)
	// {
	// 	let oFirstSoOperand = [],
	// 		oSecondSoOperand = [];
	//
	// 	const ProceedScript = function (context) {
	// 		while (context.IsScriptSpecialContent()) {
	// 			if (context.oLookahead.class === oLiteralNames.specialScriptNumberLiteral[0]) {
	// 				let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialScriptNumberLiteral, true);
	// 				oFirstSoOperand.push(oSpecial);
	// 			}
	// 			if (context.oLookahead.class === oLiteralNames.specialScriptCharLiteral[0]) {
	// 				let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialScriptCharLiteral, true);
	// 				oFirstSoOperand.push(oSpecial);
	// 			}
	// 			if (context.oLookahead.class === oLiteralNames.specialScriptBracketLiteral[0]) {
	// 				let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialScriptBracketLiteral, true)
	// 				oFirstSoOperand.push(oSpecial);
	// 			}
	// 			if (context.oLookahead.class === oLiteralNames.specialScriptOperatorLiteral[0]) {
	// 				let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialScriptOperatorLiteral, true)
	// 				oFirstSoOperand.push(oSpecial);
	// 			}
	// 		}
	// 	};
	// 	const ProceedIndex = function (context) {
	// 		while (context.IsIndexSpecialContent()) {
	// 			if (context.oLookahead.class === oLiteralNames.specialIndexNumberLiteral[0]) {
	// 				let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialIndexNumberLiteral, true);
	// 				oSecondSoOperand.push(oSpecial);
	// 			}
	// 			if (context.oLookahead.class === oLiteralNames.specialIndexCharLiteral[0]) {
	// 				let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialIndexCharLiteral, true);
	// 				oSecondSoOperand.push(oSpecial);
	// 			}
	// 			if (context.oLookahead.class === oLiteralNames.specialIndexBracketLiteral[0]) {
	// 				let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialIndexBracketLiteral, true)
	// 				oSecondSoOperand.push(oSpecial);
	// 			}
	// 			if (context.oLookahead.class === oLiteralNames.specialIndexOperatorLiteral[0]) {
	// 				let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialIndexOperatorLiteral, true)
	// 				oSecondSoOperand.push(oSpecial);
	// 			}
	// 		}
	// 	};
	//
	// 	if (this.IsScriptSpecialContent()) {
	// 		ProceedScript(this);
	// 		if (this.IsIndexSpecialContent()) {
	// 			ProceedIndex(this);
	// 		}
	// 	}
	// 	else if (this.IsIndexSpecialContent()) {
	// 		ProceedIndex(this);
	// 		if (this.IsScriptSpecialContent()) {
	// 			ProceedScript(this);
	// 		}
	// 	}
	//
	// 	return {
	// 		type: Struc.sub_sub,
	// 		value: base,
	// 		down: oSecondSoOperand,
	// 		up: oFirstSoOperand,
	// 	};
	// }
	// CUnicodeParser.prototype.IsSpecialContent = function ()
	// {
	// 	return (
	// 		this.oLookahead.class === oLiteralNames.specialScriptNumberLiteral[0] ||
	// 		this.oLookahead.class === oLiteralNames.specialScriptCharLiteral[0] ||
	// 		this.oLookahead.class === oLiteralNames.specialScriptBracketLiteral[0] ||
	// 		this.oLookahead.class === oLiteralNames.specialScriptOperatorLiteral[0] ||
	//
	// 		this.oLookahead.class === oLiteralNames.specialIndexNumberLiteral[0] ||
	// 		this.oLookahead.class === oLiteralNames.specialIndexCharLiteral[0] ||
	// 		this.oLookahead.class === oLiteralNames.specialIndexBracketLiteral[0] ||
	// 		this.oLookahead.class === oLiteralNames.specialIndexOperatorLiteral[0]
	//
	// 	);
	// };
	// CUnicodeParser.prototype.IsScriptSpecialContent = function ()
	// {
	// 	return (
	// 		this.oLookahead.class === oLiteralNames.specialScriptNumberLiteral[0] ||
	// 		this.oLookahead.class === oLiteralNames.specialScriptCharLiteral[0] ||
	// 		this.oLookahead.class === oLiteralNames.specialScriptBracketLiteral[0] ||
	// 		this.oLookahead.class === oLiteralNames.specialScriptOperatorLiteral[0]
	// 	);
	// };
	// CUnicodeParser.prototype.IsIndexSpecialContent = function ()
	// {
	// 	return (
	// 		this.oLookahead.class === oLiteralNames.specialIndexNumberLiteral[0] ||
	// 		this.oLookahead.class === oLiteralNames.specialIndexCharLiteral[0] ||
	// 		this.oLookahead.class === oLiteralNames.specialIndexBracketLiteral[0] ||
	// 		this.oLookahead.class === oLiteralNames.specialIndexOperatorLiteral[0]
	// 	);
	// };
	CUnicodeParser.prototype.GetTextOfContent = function (oInput)
	{
		let str = '';

		if (Array.isArray(oInput))
		{
			for (let i = 0; i < oInput.length; i++)
			{
				let oCurrent = oInput[i];
				if (oCurrent.type === Struc.char)
				{
					str += oCurrent.value;
				}
			}
		}
		else
			return oInput.value;

		return str;
	}
	CUnicodeParser.prototype.IsExpSubSupLiteral = function ()
	{
		return (
			this.IsScriptStandardContentLiteral() ||
			this.IsScriptBelowOrAboveContent()
			//this.IsSpecialContent()
		);
	};
	CUnicodeParser.prototype.GetExpSubSupLiteral = function (oBase)
	{
		let oThirdSoOperand, 
			oContent,
			oOfStyle;

		if (undefined === oBase) {
			oBase = this.GetScriptBaseLiteral();
		}
		else
		{
			let strBase = this.GetTextOfContent(oBase);
			if (FunctionNames.includes(strBase)) {
				oBase.type = Struc.func;
			}
			else if (LimitNames.includes(strBase)) {
				oBase.type = Struc.func_lim;
			}
		}

		// if (this.isPreScriptLiteral()) {
		// 	return this.preScriptLiteral();
		// }

		if (this.IsScriptStandardContentLiteral()) {
			oContent = this.GetScriptStandardContentLiteral(oBase);
		}
		else if (this.IsScriptBelowOrAboveContent()) {
			oContent = this.GetScriptBelowOrAboveContent(oBase);
		}
		// else if (this.IsSpecialContent()) {
		// 	oContent = this.GetScriptSpecialContent(oBase);
		// }

		if (oBase && oBase.value === " ")
		{
			return this.WriteSavedTokens();
		}

		if (this.oLookahead.class === Literals.of.id)
		{
			if (
				//oBase.type === oLiteralNames.opBuildupLiteral[num] ||
				oBase.type === Struc.nary ||
				oBase.type === Struc.func ||
				oBase.type === Struc.func_lim ||
				oBase.type === Struc.radical //todo
			)
			{
				oOfStyle = this.oLookahead.style;
				this.EatToken("▒");

				if (oBase.type === Struc.nary)
				{
					if (this.oLookahead.data ===  "〖" )
						oThirdSoOperand = this.GetExpBracketLiteral();
					else
						oThirdSoOperand = this.GetExpLiteral();
				}
				else
					oThirdSoOperand = this.GetElementLiteral();

					return {
					type: Struc.sub_sub,
					value: oBase,
					down: oContent.down,
					up: oContent.up,
					third: oThirdSoOperand,
					style: {
						subStyle: oContent.style.subStyle,
						supStyle: oContent.style.supStyle,
						ofStyle: oOfStyle,
					},
				};
			}
			else
			{
				this.EatToken(this.oLookahead.class);
				return oContent;
			}
		}
		else if (
			oBase.type ===  Struc.nary ||
			oBase.type ===  Struc.func ||
			oBase.type ===  Struc.func_lim
		)
		{
			if (this.oLookahead.class)
			{
				if (this.oLookahead.data.charCodeAt(0) === 8289 || this.oLookahead.data === " ")
					this.EatToken(this.oLookahead.class);

				oThirdSoOperand = this.GetOperandLiteral();

				return {
					type: Struc.sub_sub,
					value: oBase,
					down: oContent.down,
					up: oContent.up,
					third: oThirdSoOperand,
					style: {
						subStyle: oContent.style.subStyle,
						supStyle: oContent.style.supStyle,
						oOfStyle: oOfStyle,
					}
				};
			}
		}

		return oContent;
	};
	CUnicodeParser.prototype.EatOneSpace = function()
	{
		if (this.oLookahead.class === Literals.space.id && this.oLookahead.data === " ")
			this.EatToken(this.oLookahead.class);
	}
	CUnicodeParser.prototype.GetScriptStandardContentLiteral = function (oBase)
	{
		let oFirstElement,
			oSecondElement,
			oSubStyle,
			oSupStyle;

		if (this.oLookahead.data === "_")
		{
			oSubStyle = this.oLookahead.style;
			this.EatToken(this.oLookahead.class);

			if (this.IsSoOperandLiteral())
			{
				oFirstElement = (oBase && oBase.type === Struc.nary)
					? this.GetSoOperandLiteral("custom")
					: this.GetSpaceExitFunction(this.GetSoOperandLiteral, "_");
			}
			else if (this.IsExpLiteral())
			{
				oFirstElement = this.GetExpLiteral();
			}

			// if (this.oLookahead.class === Literals.space.id)
			// 	this.EatToken(this.oLookahead.class);

			if (undefined === oFirstElement)
				oFirstElement = {};

			// Get second element
			if (this.oLookahead.data === "^" && !this.isOneSubSup)
			{
				oSupStyle = this.oLookahead.style;
				this.EatToken(this.oLookahead.class);

				if (this.IsSoOperandLiteral())
				{
					oSecondElement = this.GetSpaceExitFunction(this.GetSoOperandLiteral, "^");
				}
				else if (this.IsExpLiteral())
				{
					oSecondElement = this.GetExpLiteral();
				}

				if (this.oLookahead.class === Literals.space.id)
					this.EatToken(this.oLookahead.class);

				return {
					type: Struc.sub_sub,
					value: oBase,
					down: oFirstElement,
					up: oSecondElement,
					style: {supStyle: oSupStyle, subStyle: oSubStyle},
				};
			}

			return {
				type: Struc.sub_sub,
				value: oBase,
				down: oFirstElement,
				style: {supStyle: oSupStyle, subStyle: oSubStyle},
			};
		}
		else if (this.oLookahead.data === "^")
		{
			oSupStyle = this.oLookahead.style;
			this.EatToken(this.oLookahead.class);

			if (this.IsSoOperandLiteral())
			{
				oSecondElement = (oBase && oBase.type === Struc.nary)
					? this.GetSoOperandLiteral("custom")
					: this.GetSpaceExitFunction(this.GetSoOperandLiteral, "^");
			}
			else if (this.IsExpLiteral())
			{
				oSecondElement = this.GetExpLiteral();
			}
			// if (this.oLookahead.class === Literals.space.id)
			// 	this.EatToken(this.oLookahead.class);

			if (oSecondElement && (oSecondElement.value === "′" || oSecondElement.value === "′′" || oSecondElement === "‵"))
			{
				oSecondElement = oSecondElement.value;
			}

			if (this.oLookahead.data === "_")
			{
				oSubStyle = this.oLookahead.style;
				this.EatToken(this.oLookahead.class);

				if (this.IsSoOperandLiteral()) {
					oFirstElement = this.GetSpaceExitFunction(this.GetSoOperandLiteral, "_");
				}
				else if (this.IsExpLiteral())
				{
					oFirstElement = this.GetExpLiteral();
				}

				if (this.oLookahead.class === Literals.space.id)
					this.EatToken(this.oLookahead.class);

				return {
					type: Struc.sub_sub,
					value: oBase,
					down: oFirstElement,
					up: oSecondElement,
					style: {supStyle: oSupStyle, subStyle: oSubStyle},
				};
			}

			return {
				type: Struc.sub_sub,
				value: oBase,
				up: oSecondElement,
				style: {supStyle: oSupStyle, subStyle: oSubStyle},
			};
		}
	};
	CUnicodeParser.prototype.IsScriptStandardContentLiteral = function ()
	{
		return this.oLookahead.data === "_" || this.oLookahead.data === "^";
	};
	CUnicodeParser.prototype.GetScriptBelowOrAboveContent = function (base)
	{
		let oBelowAbove,
			type,
			strType = this.oLookahead.data,
			oStyle = this.oLookahead.style;

		if (strType === "┬" || strType === "\\below")
			type = LIMIT_LOW;
		else if (strType === "┴" || strType === "\\above")
			type = LIMIT_UP;

		this.EatToken(this.oLookahead.class);

		oBelowAbove = this.GetSpaceExitFunction(this.GetOperandLiteral)

		if (Literals.horizontal.SearchU(base.value))
		{
			return {
				type: Struc.horizontal,
				hBrack: base,
				value: oBelowAbove,
				VJUSTType: type,
				style: oStyle,
			}
		}
		else if (this.IsApplicationFunction())
		{
			return this.GetFunctionApplication({
				type: Struc.limit,
				base: base,
				value: oBelowAbove,
				isBelow: type,
				style: oStyle,
			})
		}

		return {
			type: Struc.limit,
			base: base,
			value: oBelowAbove,
			isBelow: type,
			style: oStyle,
		};

		// if(base && base.type === oLiteralNames.functionLiteral[num])
		// {
		// 	if (this.oLookahead.data && this.oLookahead.data.charCodeAt(0) === 8289) //funcapply symbol ⁡)
		// 	{
		// 		this.EatToken(this.oLookahead.class);
		// 	}
		// 	let third = this.GetOperandLiteral();
		// 	//todo
		// 	return {
		// 		type: oLiteralNames.functionWithLimitLiteral[num],
		// 		value: base.value,
		// 		up: strType === "┴" ? oBelowAbove : undefined,
		// 		down: strType !== "┴" ? oBelowAbove : undefined,
		// 		third: third,
		// 	}
		// }
		//
		// return {
		// 	type: Struc.limit,
		// 	base: base,
		// 	value: oBelowAbove,
		// 	isBelow: type,
		// 	style: oStyle,
		// };
	};
	CUnicodeParser.prototype.IsScriptBelowOrAboveContent = function ()
	{
		return this.oLookahead.data === "┬"
			|| this.oLookahead.data === "┴"
			|| this.oLookahead.data === "\\below"
			|| this.oLookahead.data === "\\above"
	};
	CUnicodeParser.prototype.GetFractionLiteral = function (oNumerator)
	{
		let oOperand;

		if (undefined === oNumerator) {
			oNumerator = this.GetOperandLiteral();
		}

		if (this.oLookahead.class === Literals.divide.id && this.oLookahead.style.metaData.isEscapedSlash !== true)
		{
			let oFracStyle				= this.oLookahead.style,
				strOpOver				= this.EatToken(Literals.divide.id).data,
				oAddData				= this.oLookahead.style,
				oMathMetaData			= oAddData ? oAddData.GetMathMetaData() : undefined,
				isLinearMetaData		= oMathMetaData ? oMathMetaData.getIsLinearFraction() : undefined,
				isBinomialWithBrackets	= strOpOver === "⒞",
				intTypeFraction			= AscMath.GetFractionType(strOpOver);

			if (isLinearMetaData)
			{
				intTypeFraction = LINEAR_FRACTION;
			}

			if (this.IsOperandLiteral())
				oOperand = this.GetSpaceExitFunction(this.GetFractionLiteral);

			this.EatOneSpace();

			if (isBinomialWithBrackets)
			{
				return {
					type: Struc.bracket_block,
					value: {
						type: Struc.frac,
						up: oNumerator || {},
						down: oOperand || {},
						fracType: intTypeFraction,
						style: oFracStyle,
					},
					left: "(",
					right: ")",
				}
			}

			return {
				type: Struc.frac,
				up: oNumerator || {},
				down: oOperand || {},
				fracType: this.GetFractionType(strOpOver),
				style: oFracStyle,
			};
		}
		else
		{
			if (this.oLookahead.class !== undefined && this.oLookahead.style.metaData.isEscapedSlash)
			{
				let oEscSlash = this.EatToken();
				return [
					oNumerator,
					{
						type: Struc.char,
						value: oEscSlash.data,
						style: oEscSlash.style
					},
				];
			}
			return oNumerator;
		}
	};
	CUnicodeParser.prototype.GetFractionType = function (str)
	{
		switch (str)
		{
			case "/"	:	return BAR_FRACTION;
			case "⁄"	:	return SKEWED_FRACTION;
			case "⊘"	:	return LITTLE_FRACTION;
			case "¦"	:	return NO_BAR_FRACTION;
			case "∕"	:	return LINEAR_FRACTION;
		}
	}
	CUnicodeParser.prototype.IsFractionLiteral = function ()
	{
		return this.IsOperandLiteral();
	};
	CUnicodeParser.prototype.GetOtherLiteral = function ()
	{
		return this.GetCharLiteral();
	};
	CUnicodeParser.prototype.isOtherLiteral = function ()
	{
		return this.oLookahead.class === Literals.char.id;
	};
	CUnicodeParser.prototype.GetOperatorLiteral = function ()
	{
		let oStyle = this.oLookahead.style;
		let oOperator = this.EatToken(Literals.operator.id);

		return {
			type: Literals.operator.id,
			value: oOperator.data,
			style: oStyle,
		};
	};
	CUnicodeParser.prototype.GetASCIILiteral = function ()
	{
		return this.ReadTokensWhileEnd(Literals.number, Struc.number)
	};
	CUnicodeParser.prototype.GetAlphanumericLiteral = function ()
	{
		let oStyle = this.oLookahead.style;
		let oAlphanumeric = this.EatToken(Literals.alphanumeric.id);

		return {
			type: Struc.char,
			value: oAlphanumeric.data,
			style: oStyle,
		};

		return this.ReadTokensWhileEnd(Literals.alphanumeric, Struc.char);
	}
	CUnicodeParser.prototype.GetCharLiteral = function ()
	{
		return this.ReadTokensWhileEnd(Literals.char, Struc.char)
	};
	CUnicodeParser.prototype.GetAnMathLiteral = function ()
	{
		const oAnMathLiteral = this.EatToken(oLiteralNames.anMathLiteral[0]);
		return {
			type: Struc.char,
			value: oAnMathLiteral.data,
			style: oAnMathLiteral.style,
		};
	};
	CUnicodeParser.prototype.IsAnMathLiteral = function ()
	{
		return false
	};
	CUnicodeParser.prototype.GetAnOtherLiteral = function ()
	{
		if (this.oLookahead.class === Literals.other.id)
		{
			return this.ReadTokensWhileEnd(Literals.other, Struc.other)
		}
		else if (this.oLookahead.class === Literals.char.id) {
			return this.GetCharLiteral();
		}
		else if (this.oLookahead.class === Literals.alphanumeric.id)
		{
			return this.GetAlphanumericLiteral();
		}
		else if (this.oLookahead.class === Literals.number.id) {
			return this.ReadTokensWhileEnd(Literals.number, Struc.char, false)
		}
		else if (this.oLookahead.class === Literals.operand.id) {
			return this.ReadTokensWhileEnd(Literals.operand, Struc.char, false)
		}
		else if (this.oLookahead.data === "." || this.oLookahead.data === ",") {
			return {
				type: Literals.char.id,
				value: this.EatToken(this.oLookahead.class).data,
			}
		}
	};
	CUnicodeParser.prototype.IsAnOtherLiteral = function ()
	{
		return (
			this.oLookahead.class === Literals.other.id ||
			this.oLookahead.class === Literals.char.id ||
			this.oLookahead.class === Literals.number.id ||
			this.oLookahead.class === Literals.operand.id ||
			this.oLookahead.class === Literals.alphanumeric.id ||
			this.oLookahead.data === "." || this.oLookahead.data === ","
		) && this.IsSpaceExit();
	};
	CUnicodeParser.prototype.GetAnLiteral = function ()
	{
		if (this.IsAnOtherLiteral()) {
			return this.GetAnOtherLiteral();
		}
		return this.GetAnMathLiteral();
	};
	CUnicodeParser.prototype.IsAnLiteral = function ()
	{
		return this.IsAnOtherLiteral() || this.IsAnMathLiteral();
	};
	CUnicodeParser.prototype.GetDiacriticBaseLiteral = function ()
	{
		let arrData = [];

		while(this.IsAnLiteral())
		{
			arrData.push(this.GetAnLiteral());
		}

		return arrData
	};
	CUnicodeParser.prototype.IsDiacriticBaseLiteral = function ()
	{
		return (
			this.IsAnLiteral() ||
			this.oLookahead.class === Literals.number.id ||
			this.oLookahead.class === "("
		);
	};
	CUnicodeParser.prototype.GetDiacriticsLiteral = function ()
	{
		const arrDiacriticList = [];

		arrDiacriticList.push({
			type:Struc.char,
			value: this.oLookahead.data,
			style : this.oLookahead.style,
		});
		this.EatToken(this.oLookahead.class);

		return this.GetContentOfLiteral(arrDiacriticList);
	};
	CUnicodeParser.prototype.IsDiacriticsLiteral = function ()
	{
		return this.oLookahead.class === Literals.accent.id;
	};
	CUnicodeParser.prototype.GetAtomLiteral = function ()
	{
		return this.GetDiacriticBaseLiteral();
	};
	CUnicodeParser.prototype.IsAtomLiteral = function ()
	{
		return !this.IsTextLiteral()
			&& (this.IsAnLiteral() || this.IsDiacriticBaseLiteral())
	};
	CUnicodeParser.prototype.GetAtomsLiteral = function ()
	{
		const arrAtomsList = [];
		while (this.IsAtomLiteral())
		{
			let atom = this.GetAtomLiteral();

			if (this.IsPrimeLiteral())
				arrAtomsList.push(this.GetPrimeLiteral(atom))
			else
				arrAtomsList.push(atom);

		}
		return this.GetContentOfLiteral(arrAtomsList);
	};
	CUnicodeParser.prototype.IsAtomsLiteral = function ()
	{
		return this.IsAtomLiteral();
	};
	CUnicodeParser.prototype.GetTextOperandLiteral = function()
	{
		return this.ReadTokensWhileEnd(Literals.operand, Struc.char, false)
	};
	CUnicodeParser.prototype.GetEntityLiteral = function ()
	{
		if (this.IsTextLiteral()) {
			return this.GetTextLiteral()
		}
		else if (this.IsAtomsLiteral()) {
			return this.GetAtomsLiteral();
		}
		else if (this.IsExpBracketLiteral()) {
			return this.GetExpBracketLiteral();
		}
		else if (this.IsNumberLiteral()) {
			return this.GetNumberLiteral();
		}
		else if (this.IsOpNaryLiteral()) {
			return this.GetOpNaryLiteral();
		}
		else if (this.IsHBracketLiteral()) {
			return this.GetHBracketLiteral();
		}
		// else if (this.oLookahead.class === Literals.operand.id)
		// {
		// 	return this.GetTextOperandLiteral()
		// }

	};
	CUnicodeParser.prototype.IsEntityLiteral = function ()
	{
		return (
			this.IsAtomsLiteral() ||
			this.IsExpBracketLiteral() ||
			this.IsNumberLiteral() ||
			this.IsOpNaryLiteral() ||
			this.IsTextLiteral() ||
			this.IsHBracketLiteral()
			//this.oLookahead.class === Literals.operand.id
		);
	};
	CUnicodeParser.prototype.IsPrimeLiteral = function ()
	{
		return this.oLookahead.data === "⁗" ||
			this.oLookahead.data === "‴" ||
			this.oLookahead.data === "″" ||
			this.oLookahead.data === "′";
	};
	CUnicodeParser.prototype.GetPrimeLiteral = function (oBase)
	{
		let oSupStyle = this.oLookahead.style;
		let strPrime = this.EatToken(this.oLookahead.class).data;

		return {
			type: Struc.sub_sub,
			up: {
				type: Struc.char,
				value: strPrime,
			},
			value: oBase,
			style: {supStyle: oSupStyle, subStyle: undefined},
		}
	};
	CUnicodeParser.prototype.GetFactorLiteral = function ()
	{
		if (this.IsDiacriticsLiteral())
		{
			const oDiacritic = this.GetDiacriticsLiteral();
			return {
				type: Struc.accent,
				value: oDiacritic,
			};
		}
		else if (this.IsEntityLiteral() && !this.IsFunctionLiteral())
		{
			let oEntity = this.GetEntityLiteral();
			if (this.IsPrimeLiteral())
				return this.GetPrimeLiteral(oEntity);

			if (this.oLookahead.data === " ") //nbsp
				this.EatToken(this.oLookahead.class);

			while (this.IsDiacriticsLiteral())
			{
				const oDiacritic = this.GetDiacriticsLiteral();
				if (oDiacritic === "''" || oDiacritic === "'")
				{
					oEntity = {
						type: Struc.sub_sub,
						value: oEntity,
						up: oDiacritic,
					}
				}

				//nbsp processing for accents
				if (oEntity[oEntity.length - 1] && oEntity[oEntity.length - 1].value === String.fromCharCode(160)) //nbsp
					oEntity.length--;
				else if (oEntity.value === String.fromCharCode(160))
					oEntity = {};

				return {
					type: Struc.accent,
					base: oEntity,
					value: oDiacritic,
				};

				if (this.oLookahead.class === Literals.space.id)
					this.EatToken(this.oLookahead.class);
			}

			if (this.IsHBracketLiteral())
			{
				return this.GetSpecialHBracket(oEntity);
			}
			return oEntity;
		}
		else if (this.IsFunctionLiteral()) {
			return this.GetFunctionLiteral();
		}
		else if (this.IsExpSubSupLiteral()) {
			return this.GetExpSubSupLiteral();
		}
		if (this.IsMatrixLiteral()) {
			return this.GetMatrixLiteral();
		}
	};
	CUnicodeParser.prototype.IsFactorLiteral = function ()
	{
		return this.IsEntityLiteral()
			|| this.IsFunctionLiteral()
			|| this.IsDiacriticsLiteral()
			|| this.IsMatrixLiteral()
	};
	CUnicodeParser.prototype.IsSpecial = function (isNoSubSup)
	{
		return this.oLookahead.data === isNoSubSup || (
			!isNoSubSup && this.IsScriptStandardContentLiteral() ||
			!isNoSubSup && this.IsScriptBelowOrAboveContent() ||
		//	!isNoSubSup && this.IsSpecialContent() ||
			!isNoSubSup && this.IsDoubleIteratorDegree()
		)
	}
	CUnicodeParser.prototype.GetOperandLiteral = function (isNoSubSup)
	{
		const arrFactorList = [];

		if (undefined === isNoSubSup)
			isNoSubSup = false;

		let isBreak = false;

		while (this.IsFactorLiteral() && !this.IsExpSubSupLiteral() && !isBreak)
		{
			if (this.IsFactorLiteral() && !this.IsExpSubSupLiteral())
			{
				arrFactorList.push(this.GetFactorLiteral());

				if (!this.IsSpaceExit())
					return this.GetContentOfLiteral(arrFactorList);

				if (arrFactorList[arrFactorList.length - 1])
					isBreak = arrFactorList[arrFactorList.length - 1].type === Struc.bracket_block;
			}

			if (this.IsSpecial(isNoSubSup)
				|| this.IsHBracketLiteral()
				|| this.IsApplicationFunction()
				&& arrFactorList[arrFactorList.length - 1])
			{
				let oContent = arrFactorList[arrFactorList.length - 1];

				while (this.IsSpecial(isNoSubSup)
					|| this.IsHBracketLiteral()
					|| this.IsApplicationFunction()
					&& oContent
					&& oContent.value !== " ")
				{
					//if next token "_" or "^" proceed as index/degree
					if (this.oLookahead.data === isNoSubSup || !isNoSubSup && this.IsScriptStandardContentLiteral()) {
						oContent = this.GetExpSubSupLiteral(oContent);
					}
					//if next token "┬" or "┴" proceed as below/above
					else if (this.oLookahead.data === isNoSubSup || !isNoSubSup && this.IsScriptBelowOrAboveContent()) {
						oContent = this.GetScriptBelowOrAboveContent(oContent);
					}
					else if (this.IsHBracketLiteral())
					{
						oContent = this.GetHBracketLiteral(oContent);
					}
					//if next token like ⁶⁷⁸⁹ or ₂₃₄ proceed as special degree/index
					// else if (this.oLookahead.data === isNoSubSup || !isNoSubSup && this.IsSpecialContent()) {
					// 	oContent = this.GetScriptSpecialContent(oContent);
					// }
					else if (this.oLookahead.data === isNoSubSup || !isNoSubSup && this.IsDoubleIteratorDegree() && oContent.value !== " ")
					{
						oContent = this.GetDoubleIteratorDegree(oContent);
					}
					else if (this.IsApplicationFunction())
					{
						oContent = this.GetFunctionApplication(oContent);
					}
				}
				arrFactorList[arrFactorList.length - 1] = oContent;

				if (this.oLookahead.class === Literals.space.id)
				{
					this.EatToken(this.oLookahead.class);
					arrFactorList[arrFactorList.length - 1] = oContent;
					return this.GetContentOfLiteral(arrFactorList);
				}
			}
			else if (!isNoSubSup && this.oLookahead.class === Literals.space.id && !this.isSpaceExit) {
				let oTemp = this.GetContentOfLiteral(arrFactorList);
				oTemp = this.GetContentOfLiteral(oTemp);
				if (Array.isArray(oTemp) && oTemp.length > 0 && oTemp[oTemp.length - 1].type !== Literals.char.id)
					this.EatToken(this.oLookahead.class);
				else if (!Array.isArray(oTemp) && oTemp && oTemp.type !== Literals.char.id)
					this.EatToken(this.oLookahead.class);
			}
		}

		return this.GetContentOfLiteral(arrFactorList);
	};
	CUnicodeParser.prototype.IsApplicationFunction = function()
	{
		return this.oLookahead.class === Literals.invisible.id;
	}
	CUnicodeParser.prototype.GetFunctionApplication = function(oBase)
	{
		let oPr = this.oLookahead.style;
		this.EatToken(this.oLookahead.class);

		if (this.IsExpSubSupLiteral())
		{
			let oContent = this.GetExpSubSupLiteral(this.GetContentOfLiteral(oBase));

			if (oBase.type === Struc.func_lim)
			{
				return {
					type: Struc.func_lim,
					value: { type: Struc.char, value: oContent.value.value},
					up: oContent.up,
					down: oContent.down,
					isBelow: oContent.down ? LIMIT_LOW : LIMIT_UP,
					style: oContent.style,
					third: oContent.third,
				};
			}
			else
			{
				return oContent;
			}
		}

		let oValue = this.GetSpaceExitFunction(this.GetOperandLiteral);

		return {
			type: Struc.func,
			value: oBase,
			third: oValue,
			style: oPr,
		}
	}
	CUnicodeParser.prototype.IsDoubleIteratorDegree = function ()
	{
		return this.oLookahead.data === "′"
		|| this.oLookahead.data === "'"
		|| this.oLookahead.data === "″"
		|| this.oLookahead.data === "‴"
		|| this.oLookahead.data === "⁗"
	}
	CUnicodeParser.prototype.GetDoubleIteratorDegree = function (oBase)
	{
		let strIterator = this.oLookahead.data;
		this.EatToken(this.oLookahead.class);

		if (oBase && oBase.type === Literals.space.id)
		{
			oBase = {};
		}

		return {
			type: Struc.sub_sub,
			value: oBase,
			down: undefined,
			up: {
				type: Struc.char,
				value: strIterator,
			},
			//third: oThirdSoOperand,
		};
	}
	CUnicodeParser.prototype.IsOperandLiteral = function ()
	{
		return this.IsFactorLiteral();
	};
	CUnicodeParser.prototype.IsRowLiteral = function ()
	{
		return this.IsExpLiteral() || this.oLookahead.data === "&";
	};
	CUnicodeParser.prototype.GetRowLiteral = function (oStyle)
	{
		let arrRow				= [];
		let intLength			= 0;
		let intCount			= 0;
		let isAlreadyGetContent = false;

		while (this.IsExpLiteral() || this.oLookahead.data === "&")
		{
			let intCopyOfLength = intLength;
			
			if (this.oLookahead.data !== "&")
			{
				arrRow.push(this.GetExpLiteral(undefined, true));
				intLength++;
				isAlreadyGetContent = true;
			}
			else
			{
				oStyle[intCount] = this.oLookahead.style;
				this.EatToken("&");
				
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
	CUnicodeParser.prototype.GetRowsLiteral = function (cols, rows)
	{
		let arrRows			= [];
		let nRow			= 0;
		let isHasContent	= false;

		while (this.IsRowLiteral() || this.oLookahead.data === "@")
		{
			if (this.oLookahead.data === "@")
			{
				cols[nRow] = this.oLookahead.style;
				this.EatToken("@");

				if (!rows[nRow])
				{
					arrRows.push([]);

					if (!this.IsRowLiteral() && this.oLookahead.class === Literals.rBrackets.id)
						arrRows.push([]);
				}

				nRow++;
				isHasContent = true;
			}
			else
			{
				rows[nRow] = {}
				arrRows.push(this.GetRowLiteral(rows[nRow]));
			}
		}

		if (arrRows.length === 0)
		{
			for (let i = 0; i <= nRow; i++)
			{
				arrRows.push([]);
			}
		}
		else if (arrRows.length === 1 && arrRows[0].length === 0)
		{
			arrRows.push([]);
		}

		return arrRows
	};
	CUnicodeParser.prototype.GetMatrixLiteral = function ()
	{
		this.SaveTokensWhileReturn();

		let oStyles = {
			head: this.oLookahead.style,
			cols: {},
			rows: {},
		};

		let isArray = this.oLookahead.data === "█";

		let strType = this.EatToken(this.oLookahead.class).data;

		if (this.oLookahead.data !== "(")
		{
			return {
				type:	Literals.char.id,
				value:	strType,
				style:	oStyles.head,
			}
		}
		else
		{
			this.EatToken(this.oLookahead.class);
		}

		let arrMatrixContent			= this.GetRowsLiteral(oStyles.cols, oStyles.rows);
		let intMaxLengthOfMatrixRow		= -Infinity;
		let intIndexOfMaxMatrixRow		= -1;

		for (let i = 0; i < arrMatrixContent.length; i++)
		{
			let arrContent			= arrMatrixContent[i];
			intMaxLengthOfMatrixRow	= arrContent.length;
			intIndexOfMaxMatrixRow	= i;
		}

		for (let i = 0; i < arrMatrixContent.length; i++)
		{
			if (i !== intIndexOfMaxMatrixRow)
			{
				let oRow = arrMatrixContent[i];
				for (let j = oRow.length; j < intMaxLengthOfMatrixRow; j++)
				{
					oRow.push({});
				}
			}
		}

		if (this.oLookahead.data === ")")
		{
			this.EatToken(Literals.rBrackets.id);
			return {
				type:	isArray ? Struc.array : Struc.matrix,
				value:	arrMatrixContent,
				style:	oStyles,
			};
		}

		// get error by parsing matrix or array - write as text
		return this.WriteSavedTokens();
	};
	CUnicodeParser.prototype.IsMatrixLiteral = function ()
	{
		return this.oLookahead.class === Literals.matrix.id
	};
	CUnicodeParser.prototype.IsElementLiteral = function ()
	{
		return this.IsFractionLiteral() || this.IsOperandLiteral();
	};
	CUnicodeParser.prototype.GetElementLiteral = function ()
	{
		let oOperandLiteral = this.GetOperandLiteral();
		
		while (this.oLookahead.class === Literals.divide.id)
			oOperandLiteral = this.GetFractionLiteral(oOperandLiteral);

		return oOperandLiteral;
	};
	CUnicodeParser.prototype.IsExpLiteral = function ()
	{
		return this.IsElementLiteral() ||
			this.oLookahead.class === Literals.operator.id ||
			this.oLookahead.class === Literals.divide.id ||
			this.oLookahead.class === Literals.space.id ||
			this.IsPreScriptLiteral();
	};
	CUnicodeParser.prototype.GetExpLiteral = function (arrCorrectSymbols, isMatrix)
	{
		if (!arrCorrectSymbols)
			arrCorrectSymbols = [];

		const oExpLiteral = [];

		while (this.IsExpLiteral() || (!isMatrix && (this.oLookahead.data === "@" || this.oLookahead.data === "&")) || (arrCorrectSymbols.includes(this.oLookahead.data)))
		{
			if (this.oLookahead.class === Literals.divide.id)
			{
				oExpLiteral.push(this.GetFractionLiteral({}));
			}
			else if (this.oLookahead.class === Literals.space.id)
			{
				oExpLiteral.push(this.GetSpaceLiteral());
			}
			else if (this.IsElementLiteral())
			{
				let oElement = this.GetElementLiteral();
				
				if (oElement !== null)
					oExpLiteral.push(oElement);

				if (oElement && oElement.length > 0 && oElement[oElement.length - 1].type !== Literals.char.id)
					this.EatOneSpace();
			}
			else if (arrCorrectSymbols.includes(this.oLookahead.data))
			{
				oExpLiteral.push({
					type: Struc.char,
					value: this.EatToken(this.oLookahead.class).data
				})
			}
			else if (this.oLookahead.class === Literals.operator.id && !this.IsDoubleIteratorDegree())
			{
				oExpLiteral.push(this.GetOperatorLiteral())
			}
			else if (this.IsPreScriptLiteral())
			{
				let oPreScriptLiteral = this.GetPreScriptLiteral();
				if (oPreScriptLiteral.type && oPreScriptLiteral.type === Struc.pre_script && oExpLiteral.length > 0 && oExpLiteral[0].value === " ")
				{
					oExpLiteral.length--;
				}
				oExpLiteral.push(oPreScriptLiteral);
			}
			else if (this.IsDoubleIteratorDegree())
			{
				oExpLiteral.push(this.GetDoubleIteratorDegree());
			}
			else if (this.oLookahead.data === "@" || this.oLookahead.data === "&")
			{
				let oPr = this.oLookahead.style;
				oExpLiteral.push({
					type: Struc.char,
					value: this.EatToken(this.oLookahead.class).data,
					style: oPr,
				})
			}
		}

		if (oExpLiteral.length === 1) {
			return oExpLiteral[0];
		}

		return oExpLiteral;
	};
	/**
	 * Метод позволяет обрабатывать токены одного типа, пока они не прервутся другим типом токенов
	 *
	 * @param arrTypeOfLiteral {LiteralType}
	 * @param isSpecial {boolean}
	 * @return {array} Обработанные токены
	 * @constructor
	 */
	CUnicodeParser.prototype.ReadTokensWhileEnd = function (arrTypeOfLiteral, type, isSpecial)
	{
		let arrLiterals = [];
		let strLiteral = "";
		let styles = [];

		while (this.oLookahead.class === arrTypeOfLiteral.id
		&& (styles.length === 0 || styles[styles.length - 1].IsStyleEqual(this.oLookahead.style)))
		{
			styles.push(this.oLookahead.style);

			if (isSpecial)
				strLiteral += UnicodeSpecialScript[this.EatToken(arrTypeOfLiteral.id).data];
			else
				strLiteral += this.EatToken(arrTypeOfLiteral.id).data;
		}

		arrLiterals.push({type: type, value: strLiteral, style: styles});

		if (arrLiterals.length === 1)
			return arrLiterals[0];

		return arrLiterals
	};
	CUnicodeParser.prototype.EatToken = function (tokenType)
	{
		const token = this.oLookahead;

		if (this.isSaveTokens)
			this.arrSavedTokens.push(this.oLookahead);

		this.prevLookahead	= this.oLookahead;
		this.oLookahead		= this.oTokenizer.GetNextToken();
		return token;
	};
	CUnicodeParser.prototype.GetContentOfLiteral = function (oContent)
	{
		if (Array.isArray(oContent))
		{
			if (oContent.length === 1)
				return oContent[0];

			return oContent;
		}
		return oContent;
	};
	CUnicodeParser.prototype.SkipSpace = function ()
	{
		while (this.oLookahead.class === Literals.space.id)
		{
			this.EatToken(Literals.space.id);
		}
	};
	CUnicodeParser.prototype.SaveTokensWhileReturn = function ()
	{
		//todo recursive
		if (!this.isSaveTokens)
		{
			this.isSaveTokens = true;
			this.arrSavedTokens = [];
		}
	};
	CUnicodeParser.prototype.WriteSavedTokens = function ()
	{
		let intSavedTokensLength = this.arrSavedTokens.length;
		let arrOutput = [];

		for (let i = 0; i < intSavedTokensLength; i++)
		{
			let oCurrentEl = this.arrSavedTokens[i];
			arrOutput.push({
				type: Struc.char,
				value: oCurrentEl.data,
				style: oCurrentEl.style,
			})
		}
		this.isSaveTokens = false;
		return arrOutput;
	};

	function CUnicodeConverter(str, oContext, isGetOnlyTokens)
	{
		if (undefined === str || null === str)
			return;

		const oParser = new CUnicodeParser();
		const oTokens = oParser.Parse(str);

		if (!isGetOnlyTokens)
			ConvertTokens(oTokens, oContext);
		else
			return oTokens;
	}

	//--------------------------------------------------------export----------------------------------------------------
	window["AscMath"] = window["AscMath"] || {};
	window["AscMath"].CUnicodeConverter = CUnicodeConverter;

})(window);

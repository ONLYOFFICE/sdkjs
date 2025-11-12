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

(function(window)
{
	/**
	 * Базовый класс элемента поиска в патерне
	 * @constructor
	 */
	function CSearchTextItemBase()
	{
	}
	CSearchTextItemBase.prototype.GetType = function()
	{
		return this.Type;
	};
	CSearchTextItemBase.prototype.GetValue = function()
	{
		return this.Value;
	};
	/**
	 * Проверяем, подходит ли заданный символ под текущий
	 * @param {CSearchTextItemBase} oItem
	 * @returns {boolean}
	 */
	CSearchTextItemBase.prototype.IsMatch = function(oItem)
	{
		return false;
	};
	/**
	 * Конвертируем данный элемент в элемент рана для вставки в документ
	 * @param {boolean} isMathRun
	 * @returns {?AscWord.CRunElementBase}
	 */
	CSearchTextItemBase.prototype.ToRunElement = function(isMathRun)
	{
		return null
	};
	CSearchTextItemBase.prototype.IsChar = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsNewLine = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsTab = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsParaEnd = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsAnySymbol = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsAnyDigit = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsAnyLetter = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsColumnBreak = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsEndnoteMark = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsField = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsPageBreak = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsFootnoteMark = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsGraphicObject = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsNonBreakingHyphen = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsNonBreakingSpace = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsAnySpace = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsEmDash = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsEnDash = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsSectionCharacter = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsParagraphCharacter = function()
	{
		return false;
	};
	CSearchTextItemBase.prototype.IsAnyDash = function()
	{
		return false;
	};


	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextItemChar(nCharCode)
	{
		this.Value = nCharCode;
	}
	CSearchTextItemChar.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextItemChar.prototype.IsMatch = function(oItem)
	{
		return ((oItem.IsChar() && this.GetValue() === oItem.GetValue())
			|| (0x2D === this.Value && oItem.IsNonBreakingHyphen()));
	};
	CSearchTextItemChar.prototype.ToRunElement = function(isMathRun)
	{
		if (isMathRun)
		{
			var oMathText = new CMathText();
			oMathText.add(this.Value);
			return oMathText;
		}
		else
		{
			if (AscCommon.IsSpace(this.Value))
				return new AscWord.CRunSpace(this.Value);
			else
				return new AscWord.CRunText(this.Value);
		}
	};
	CSearchTextItemChar.prototype.IsChar = function()
	{
		return true;
	};
	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialLineBreak()
	{
	}
	CSearchTextSpecialLineBreak.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialLineBreak.prototype.IsMatch = function(oItem)
	{
		return oItem.IsNewLine();
	};
	CSearchTextSpecialLineBreak.prototype.ToRunElement = function(isMathRun)
	{
		if (isMathRun)
			return null;

		return new AscWord.CRunBreak(AscWord.break_Line);
	};
	CSearchTextSpecialLineBreak.prototype.IsNewLine = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialTab()
	{
	}
	CSearchTextSpecialTab.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialTab.prototype.IsMatch = function(oItem)
	{
		return oItem.IsTab();
	};
	CSearchTextSpecialTab.prototype.ToRunElement = function(isMathRun)
	{
		if (isMathRun)
			return null;

		return new AscWord.CRunTab();
	};
	CSearchTextSpecialTab.prototype.IsTab = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialParaEnd()
	{
	}
	CSearchTextSpecialParaEnd.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialParaEnd.prototype.IsMatch = function(oItem)
	{
		return oItem.IsParaEnd();
	};
	CSearchTextSpecialParaEnd.prototype.IsParaEnd = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialAnySymbol()
	{
	}
	CSearchTextSpecialAnySymbol.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialAnySymbol.prototype.IsMatch = function(oItem)
	{
		return (oItem.IsChar()
			|| oItem.IsAnySymbol()
			|| oItem.IsAnyDigit()
			|| oItem.IsAnyLetter()
			|| oItem.IsNonBreakingHyphen()
			|| oItem.IsNonBreakingSpace()
			|| oItem.IsNewLine()
			|| oItem.IsAnySpace()
			|| oItem.IsEmDash()
			|| oItem.IsEnDash());
	};
	CSearchTextSpecialAnySymbol.prototype.IsAnySymbol = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialAnyDigit()
	{
	}
	CSearchTextSpecialAnyDigit.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialAnyDigit.prototype.IsMatch = function(oItem)
	{
		return ((oItem.IsChar() && AscCommon.IsDigit(oItem.GetValue()))
			|| oItem.IsAnySymbol()
			|| oItem.IsAnyDigit());
	};
	CSearchTextSpecialAnyDigit.prototype.IsAnyDigit = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialAnyLetter()
	{
	}
	CSearchTextSpecialAnyLetter.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialAnyLetter.prototype.IsMatch = function(oItem)
	{
		return ((oItem.IsChar() && AscCommon.IsLetter(oItem.GetValue()))
			|| oItem.IsAnyLetter()
			|| oItem.IsAnySymbol());
	};
	CSearchTextSpecialAnyLetter.prototype.IsAnyLetter = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialColumnBreak()
	{
	}
	CSearchTextSpecialColumnBreak.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialColumnBreak.prototype.IsMatch = function(oItem)
	{
		return oItem.IsColumnBreak();
	};
	CSearchTextSpecialColumnBreak.prototype.ToRunElement = function(isMathRun)
	{
		if (isMathRun)
			return null;

		return new AscWord.CRunBreak(AscWord.break_Column);
	};
	CSearchTextSpecialColumnBreak.prototype.IsColumnBreak = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialEndnoteMark()
	{
	}
	CSearchTextSpecialEndnoteMark.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialEndnoteMark.prototype.IsMatch = function(oItem)
	{
		return oItem.IsEndnoteMark();
	};
	CSearchTextSpecialEndnoteMark.prototype.IsEndnoteMark = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialField()
	{
	}
	CSearchTextSpecialField.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialField.prototype.IsMatch = function(oItem)
	{
		return oItem.IsField();
	};
	CSearchTextSpecialField.prototype.IsField = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialFootnoteMark()
	{
	}
	CSearchTextSpecialFootnoteMark.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialFootnoteMark.prototype.IsMatch = function(oItem)
	{
		return oItem.IsFootnoteMark();
	};
	CSearchTextSpecialFootnoteMark.prototype.IsFootnoteMark = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialGraphicObject()
	{
		this.DrawingType = 1;
	}
	CSearchTextSpecialGraphicObject.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialGraphicObject.prototype.IsMatch = function(oItem)
	{
		return oItem.IsGraphicObject();
	};
	CSearchTextSpecialGraphicObject.prototype.IsGraphicObject = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialPageBreak()
	{
	}
	CSearchTextSpecialPageBreak.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialPageBreak.prototype.IsMatch = function(oItem)
	{
		return oItem.IsPageBreak();
	};
	CSearchTextSpecialPageBreak.prototype.ToRunElement = function(isMathRun)
	{
		if (isMathRun)
			return null;

		return new AscWord.CRunBreak(AscWord.break_Page);
	};
	CSearchTextSpecialPageBreak.prototype.IsPageBreak = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialNonBreakingHyphen()
	{
	}
	CSearchTextSpecialNonBreakingHyphen.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialNonBreakingHyphen.prototype.IsMatch = function(oItem)
	{
		return (oItem.IsNonBreakingHyphen()
			|| oItem.IsAnySymbol());
	};
	CSearchTextSpecialNonBreakingHyphen.prototype.ToRunElement = function(isMathRun)
	{
		if (isMathRun)
		{
			var oMathText = new CMathText();
			oMathText.add(0x2D);
			return oMathText;
		}
		else
		{
			return AscWord.CreateNonBreakingHyphen();
		}
	};
	CSearchTextSpecialNonBreakingHyphen.prototype.IsNonBreakingHyphen = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialNonBreakingSpace()
	{
		this.Value = 160;
	}
	CSearchTextSpecialNonBreakingSpace.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialNonBreakingSpace.prototype.IsMatch = function(oItem)
	{
		return ((oItem.IsChar() && 0xA0 === oItem.GetValue())
			|| oItem.IsNonBreakingSpace()
			|| oItem.IsAnySymbol());
	};
	CSearchTextSpecialNonBreakingSpace.prototype.ToRunElement = function(isMathRun)
	{
		if (isMathRun)
			return null;

		return new AscWord.CRunText(0x00A0);
	};
	CSearchTextSpecialNonBreakingSpace.prototype.IsNonBreakingSpace = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialAnySpace()
	{
	}
	CSearchTextSpecialAnySpace.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialAnySpace.prototype.IsMatch = function(oItem)
	{
		return ((oItem.IsChar() && (AscCommon.IsSpace(oItem.GetValue()) || 0xA0 === oItem.GetValue()))
			|| oItem.IsAnySpace()
			|| oItem.IsNonBreakingSpace()
			|| oItem.IsAnySymbol());
	};
	CSearchTextSpecialAnySpace.prototype.IsAnySpace = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialEmDash()
	{
		this.Value = 0x2014;
	}
	CSearchTextSpecialEmDash.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialEmDash.prototype.IsMatch = function(oItem)
	{
		return ((oItem.IsChar() && 0x2014 === oItem.GetValue())
			|| oItem.IsEmDash()
			|| oItem.IsAnySymbol());
	};
	CSearchTextSpecialEmDash.prototype.ToRunElement = function(isMathRun)
	{
		if (isMathRun)
		{
			var oMathText = new CMathText();
			oMathText.add(0x2014);
			return oMathText;
		}
		else
		{
			return new AscWord.CRunText(0x2014);
		}
	};
	CSearchTextSpecialEmDash.prototype.IsEmDash = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialEnDash()
	{
		this.Value = 0x2013;
	}
	CSearchTextSpecialEnDash.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialEnDash.prototype.IsMatch = function(oItem)
	{
		return ((oItem.IsChar() && 0x2013 === oItem.GetValue())
			|| oItem.IsEnDash()
			|| oItem.IsAnySymbol());
	};
	CSearchTextSpecialEnDash.prototype.ToRunElement = function(isMathRun)
	{
		if (isMathRun)
		{
			var oMathText = new CMathText();
			oMathText.add(0x2013);
			return oMathText;
		}
		else
		{
			return new AscWord.CRunText(0x2013);
		}
	};
	CSearchTextSpecialEnDash.prototype.IsEnDash = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialSectionCharacter()
	{
		this.Value = 167;
	}
	CSearchTextSpecialSectionCharacter.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialSectionCharacter.prototype.IsMatch = function(oItem)
	{
		return oItem.IsSectionCharacter();
	};
	CSearchTextSpecialSectionCharacter.prototype.IsSectionCharacter = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialParagraphCharacter()
	{
		this.Value = 182;
	}
	CSearchTextSpecialParagraphCharacter.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialParagraphCharacter.prototype.IsMatch = function(oItem)
	{
		return oItem.IsParagraphCharacter();
	};
	CSearchTextSpecialParagraphCharacter.prototype.IsParagraphCharacter = function()
	{
		return true;
	};

	/**
	 * @constructor
	 * @extends CSearchTextItemBase
	 */
	function CSearchTextSpecialAnyDash()
	{
	}
	CSearchTextSpecialAnyDash.prototype = Object.create(CSearchTextItemBase.prototype);
	CSearchTextSpecialAnyDash.prototype.IsMatch = function(oItem)
	{
		return ((oItem.IsChar() && (0x2D === oItem.GetValue() || (0x2010 <= oItem.GetValue() && oItem.GetValue() <= 0x2015)))
			|| oItem.IsAnyDash()
			|| oItem.IsEmDash()
			|| oItem.IsEnDash()
			|| oItem.IsAnySymbol());
	};
	CSearchTextSpecialAnyDash.prototype.IsAnyDash = function()
	{
		return true;
	};


	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].CSearchTextItemChar                  = CSearchTextItemChar;
	window['AscCommonWord'].CSearchTextSpecialLineBreak          = CSearchTextSpecialLineBreak;
	window['AscCommonWord'].CSearchTextSpecialTab                = CSearchTextSpecialTab;
	window['AscCommonWord'].CSearchTextSpecialParaEnd            = CSearchTextSpecialParaEnd;
	window['AscCommonWord'].CSearchTextSpecialAnySymbol          = CSearchTextSpecialAnySymbol;
	window['AscCommonWord'].CSearchTextSpecialAnyDigit           = CSearchTextSpecialAnyDigit;
	window['AscCommonWord'].CSearchTextSpecialAnyLetter          = CSearchTextSpecialAnyLetter;
	window['AscCommonWord'].CSearchTextSpecialColumnBreak        = CSearchTextSpecialColumnBreak;
	window['AscCommonWord'].CSearchTextSpecialEndnoteMark        = CSearchTextSpecialEndnoteMark;
	window['AscCommonWord'].CSearchTextSpecialField              = CSearchTextSpecialField;
	window['AscCommonWord'].CSearchTextSpecialFootnoteMark       = CSearchTextSpecialFootnoteMark;
	window['AscCommonWord'].CSearchTextSpecialGraphicObject      = CSearchTextSpecialGraphicObject;
	window['AscCommonWord'].CSearchTextSpecialPageBreak          = CSearchTextSpecialPageBreak;
	window['AscCommonWord'].CSearchTextSpecialNonBreakingHyphen  = CSearchTextSpecialNonBreakingHyphen;
	window['AscCommonWord'].CSearchTextSpecialNonBreakingSpace   = CSearchTextSpecialNonBreakingSpace;
	window['AscCommonWord'].CSearchTextSpecialAnySpace           = CSearchTextSpecialAnySpace;
	window['AscCommonWord'].CSearchTextSpecialEmDash             = CSearchTextSpecialEmDash;
	window['AscCommonWord'].CSearchTextSpecialEnDash             = CSearchTextSpecialEnDash;
	window['AscCommonWord'].CSearchTextSpecialSectionCharacter   = CSearchTextSpecialSectionCharacter;
	window['AscCommonWord'].CSearchTextSpecialParagraphCharacter = CSearchTextSpecialParagraphCharacter;
	window['AscCommonWord'].CSearchTextSpecialAnyDash            = CSearchTextSpecialAnyDash;

})(window);

/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
/**
 * User: Ilja.Kirillov
 * Date: 20.10.2017
 * Time: 15:46
 */

var fieldtype_UNKNOWN    = 0x0000;
var fieldtype_MERGEFIELD = 0x0001;
var fieldtype_PAGENUM    = 0x0002;
var fieldtype_PAGECOUNT  = 0x0003;
var fieldtype_FORMTEXT   = 0x0004;
var fieldtype_TOC        = 0x0005;
var fieldtype_PAGEREF    = 0x0006;
var fieldtype_PAGE       = fieldtype_PAGENUM;
var fieldtype_NUMPAGES   = fieldtype_PAGECOUNT;

var fieldtype_ASK        = 0x0007;
var fieldtype_REF        = 0x0008;
var fieldtype_HYPERLINK  = 0x0009;
var fieldtype_TIME       = 0x000A;
var fieldtype_DATE       = 0x000B;
var fieldtype_FORMULA    = 0x0010;
var fieldtype_SEQ        = 0x0011;
var fieldtype_STYLEREF   = 0x0012;
var fieldtype_NOTEREF    = 0x0013;
var fieldtype_ADDIN      = 0x0014;


//--------------------------------------------------------export----------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};

window['AscCommonWord'].fieldtype_UNKNOWN    = fieldtype_UNKNOWN;
window['AscCommonWord'].fieldtype_MERGEFIELD = fieldtype_MERGEFIELD;
window['AscCommonWord'].fieldtype_PAGENUM    = fieldtype_PAGENUM;
window['AscCommonWord'].fieldtype_PAGECOUNT  = fieldtype_PAGECOUNT;
window['AscCommonWord'].fieldtype_FORMTEXT   = fieldtype_FORMTEXT;
window['AscCommonWord'].fieldtype_TOC        = fieldtype_TOC;
window['AscCommonWord'].fieldtype_PAGEREF    = fieldtype_PAGEREF;
window['AscCommonWord'].fieldtype_PAGE       = fieldtype_PAGE;
window['AscCommonWord'].fieldtype_NUMPAGES   = fieldtype_NUMPAGES;
window['AscCommonWord'].fieldtype_ASK        = fieldtype_ASK;
window['AscCommonWord'].fieldtype_REF        = fieldtype_REF;
window['AscCommonWord'].fieldtype_HYPERLINK  = fieldtype_HYPERLINK;
window['AscCommonWord'].fieldtype_TIME       = fieldtype_TIME;
window['AscCommonWord'].fieldtype_DATE       = fieldtype_DATE;
window['AscCommonWord'].fieldtype_FORMULA    = fieldtype_FORMULA;
window['AscCommonWord'].fieldtype_SEQ        = fieldtype_SEQ;
window['AscCommonWord'].fieldtype_STYLEREF   = fieldtype_STYLEREF;
window['AscCommonWord'].fieldtype_NOTEREF    = fieldtype_NOTEREF;
window['AscCommonWord'].fieldtype_ADDIN      = fieldtype_ADDIN;


/**
 * Базовый класс для инструкции сложного поля.
 * @constructor
 */
function CFieldInstructionBase()
{
	this.ComplexField    = null;
	this.InstructionLine = "";
}
CFieldInstructionBase.prototype.Type = fieldtype_UNKNOWN;
CFieldInstructionBase.prototype.GetType = function()
{
	return this.Type;
};
CFieldInstructionBase.prototype.SetComplexField = function(oComplexField)
{
	this.ComplexField = oComplexField;
};
CFieldInstructionBase.prototype.GetComplexField = function()
{
	return this.ComplexField;
};
CFieldInstructionBase.prototype.ToString = function()
{
	return "";
};
CFieldInstructionBase.prototype.SetPr = function()
{
};
CFieldInstructionBase.prototype.SetInstructionLine = function(sLine)
{
	this.InstructionLine = sLine;
};
CFieldInstructionBase.prototype.CheckInstructionLine = function(sLine)
{
	return (this.InstructionLine === sLine);
};
//--------------------------------------------------------export----------------------------------------------------
window['AscWord'] = window['AscWord'] || {};
window['AscWord'].CFieldInstructionBase = CFieldInstructionBase;

/**
* FORMULA field
* @constructor
*/
function CFieldInstructionFORMULA()
{
	CFieldInstructionBase.call(this);
	this.ParseQueue = null;
	this.Error = null;
	this.ErrStr = null;
	this.ResultStr = null;
	this.Format = null;
	this.ParentContent = null;
}

CFieldInstructionFORMULA.prototype = Object.create(CFieldInstructionBase.prototype);
CFieldInstructionFORMULA.prototype.constructor = CFieldInstructionFORMULA;
CFieldInstructionFORMULA.prototype.Type = fieldtype_FORMULA;
CFieldInstructionFORMULA.prototype.SetFormat = function(oFormat)
{
    this.Format = oFormat;
};
CFieldInstructionFORMULA.prototype.SetParseQueue = function(oParseQueue)
{
    this.ParseQueue = oParseQueue;
};
CFieldInstructionFORMULA.prototype.SetError = function(oError)
{
    this.Error = oError;
};

CFieldInstructionFORMULA.prototype.SetFormula = function(sFormula)
{
    this.Formula = sFormula;
};

CFieldInstructionFORMULA.prototype.GetErrorStr = function (oErr) {
	var ret = "!";
	if(oErr)
	{
		if(typeof oErr.Type === 'string')
		{
			ret += AscCommon.translateManager.getValue(oErr.Type);
		}
		if(typeof oErr.Data === 'string')
		{
			ret += (", " + oErr.Data);
		}
	}
	return ret;
};

CFieldInstructionFORMULA.prototype.Calculate = function(oLogicDocument)
{
	this.ErrStr = null;
	this.ResultStr = null;
	this.private_Calculate(oLogicDocument);
	if(this.Error)
	{
		this.ErrStr = this.GetErrorStr(this.Error);
		return;
	}
	if(this.ParseQueue)
	{
		var oCalcError = this.ParseQueue.calculate(oLogicDocument);
		if(oCalcError)
		{
			this.ErrStr = this.GetErrorStr(oCalcError);
			return;
		}
		if(typeof this.ParseQueue.resultS === "string")
		{
			this.ResultStr = this.ParseQueue.resultS;
		}
		else
		{
			this.ResultStr = '';
		}
	}
	else
	{
		this.ResultStr = '';
	}
};
CFieldInstructionFORMULA.prototype.private_Calculate = function (oLogicDocument)
{
	var sListSeparator  = ",";
	var sDigitSeparator = ".";

	if (oLogicDocument && oLogicDocument.GetDecimalSymbol && oLogicDocument.GetListSeparator && oLogicDocument.GetDecimalSymbol() !== oLogicDocument.GetListSeparator())
	{
		sListSeparator  = oLogicDocument.GetListSeparator();
		sDigitSeparator = oLogicDocument.GetDecimalSymbol();
	}

	var oParser = new AscCommonWord.CFormulaParser(sListSeparator, sDigitSeparator);
	oParser.parse(this.Formula, this.ParentContent);

	this.SetParseQueue(oParser.parseQueue);
	if (oParser.parseQueue)
	{
		oParser.parseQueue.format = this.Format;
	}
	this.SetError(oParser.error);
};

CFieldInstructionFORMULA.prototype.SetComplexField = function(oComplexField){
	CFieldInstructionBase.prototype.SetComplexField.call(this, oComplexField);
	this.ParentContent = null;
	var oBeginChar = oComplexField.BeginChar;
	if(oBeginChar)
	{
		var oRun = oBeginChar.Run;
		if(oRun)
		{
			var oParagraph = oRun.Paragraph;
			if(oParagraph)
			{
				this.ParentContent = oParagraph.Parent;
			}
		}
	}
};

/**
 * PAGE field
 * @constructor
 */
function CFieldInstructionPAGE()
{
	CFieldInstructionBase.call(this);
}

CFieldInstructionPAGE.prototype = Object.create(CFieldInstructionBase.prototype);
CFieldInstructionPAGE.prototype.constructor = CFieldInstructionPAGE;
CFieldInstructionPAGE.prototype.Type = fieldtype_PAGE;

/**
 * PAGEREF field
 * @constructor
 */
function CFieldInstructionPAGEREF(sBookmarkName, isHyperlink, isPositionRelative)
{
	CFieldInstructionBase.call(this);

	this.BookmarkName = sBookmarkName ? sBookmarkName : "";
	this.Hyperlink    = isHyperlink ? true : false;
	this.PosRelative  = isPositionRelative ? true : false;
}

CFieldInstructionPAGEREF.prototype = Object.create(CFieldInstructionBase.prototype);
CFieldInstructionPAGEREF.prototype.constructor = CFieldInstructionPAGEREF;
CFieldInstructionPAGEREF.prototype.Type = fieldtype_PAGEREF;
CFieldInstructionPAGEREF.prototype.SetHyperlink = function(isHyperlink)
{
	this.Hyperlink   = isHyperlink ? true : false;
};
CFieldInstructionPAGEREF.prototype.SetPositionRelative = function(isPosRel)
{
	this.PosRelative = isPosRel ? true : false;
};
CFieldInstructionPAGEREF.prototype.IsHyperlink = function()
{
	return this.Hyperlink;
};
CFieldInstructionPAGEREF.prototype.IsPositionRelative = function()
{
	return this.PosRelative;
};
CFieldInstructionPAGEREF.prototype.GetBookmarkName = function()
{
	return this.BookmarkName;
};

/**
 * TOC field
 * @constructor
 */
function CFieldInstructionTOC()
{
	CFieldInstructionBase.call(this);

	this.PreserveTabs     = false;
	this.RemoveBreaks     = true;
	this.Hyperlinks       = false;
	this.Separator        = "";
	this.HeadingS         = -1;
	this.HeadingE         = -1;
	this.Styles           = [];
	this.SkipPageRef      = false;
	this.SkipPageRefStart = -1;
	this.SkipPageRefEnd   = -1;
	this.ForceTabLeader   = undefined;
	this.Caption          = undefined;
	this.CaptionOnlyText  = undefined;
}

CFieldInstructionTOC.prototype = Object.create(CFieldInstructionBase.prototype);
CFieldInstructionTOC.prototype.constructor = CFieldInstructionTOC;
CFieldInstructionTOC.prototype.Type = fieldtype_TOC;
CFieldInstructionTOC.prototype.IsPreserveTabs = function()
{
	return this.PreserveTabs;
};
CFieldInstructionTOC.prototype.SetPreserveTabs = function(isPreserve)
{
	this.PreserveTabs = isPreserve;
};
CFieldInstructionTOC.prototype.IsRemoveBreaks = function()
{
	return this.RemoveBreaks;
};
CFieldInstructionTOC.prototype.SetRemoveBreaks = function(isRemove)
{
	this.RemoveBreaks = isRemove;
};
CFieldInstructionTOC.prototype.IsHyperlinks = function()
{
	return this.Hyperlinks;
};
CFieldInstructionTOC.prototype.SetHyperlinks = function(isHyperlinks)
{
	this.Hyperlinks = isHyperlinks;
};
CFieldInstructionTOC.prototype.SetSeparator = function(sSeparator)
{
	this.Separator = sSeparator;
};
CFieldInstructionTOC.prototype.GetSeparator = function()
{
	return this.Separator;
};
CFieldInstructionTOC.prototype.SetHeadingRange = function(nStart, nEnd)
{
	this.HeadingS = nStart;
	this.HeadingE = nEnd;
};
CFieldInstructionTOC.prototype.GetHeadingRangeStart = function()
{
	return this.HeadingS;
};
CFieldInstructionTOC.prototype.GetHeadingRangeEnd = function()
{
	return this.HeadingE;
};
CFieldInstructionTOC.prototype.SetStylesArrayRaw = function(sString)
{
	var oLogicDocument = editor.WordControl.m_oLogicDocument;
	var sListSeparator = oLogicDocument.GetListSeparator();

	var arrValues = sString.split(sListSeparator);
	var arrStyles = [];

	for (var nIndex = 0, nCount = arrValues.length; nIndex < nCount; ++nIndex)
	{
		var sName = arrValues[nIndex];
		var nLvl  = nIndex + 1 >= nCount ? 1 : parseInt(arrValues[nIndex + 1]);
		if (isNaN(nLvl))
			nLvl = undefined;
		else
			nIndex++;

		arrStyles.push({
			Name : sName,
			Lvl  : nLvl
		});
	}

	this.SetStylesArray(arrStyles);
};
CFieldInstructionTOC.prototype.SetStylesArray = function(arrStyles)
{
	this.Styles = arrStyles;
};
CFieldInstructionTOC.prototype.GetStylesArray = function()
{
	return this.Styles;
};
CFieldInstructionTOC.prototype.SetPageRefSkippedLvls = function(isSkip, nSkipStart, nSkipEnd)
{
	this.SkipPageRef = isSkip;

	if (true === isSkip
		&& null !== nSkipStart
		&& undefined !== nSkipStart
		&& null !== nSkipEnd
		&& undefined !== nSkipEnd)
	{
		this.SkipPageRefStart = nSkipStart;
		this.SkipPageRefEnd   = nSkipEnd;
	}
	else
	{
		this.SkipPageRefStart = -1;
		this.SkipPageRefEnd   = -1;
	}
};
CFieldInstructionTOC.prototype.IsSkipPageRefLvl = function(nLvl)
{
	if (undefined === nLvl)
		return this.SkipPageRef;

	if (false === this.SkipPageRef)
		return false;

	if (-1 === this.SkipPageRefStart || -1 === this.SkipPageRefEnd)
		return true;

	return  (nLvl >= this.SkipPageRefStart - 1 && nLvl <= this.SkipPageRefEnd - 1);
};
CFieldInstructionTOC.prototype.SetCaption = function (sCaption)
{
	this.Caption = sCaption;
};
CFieldInstructionTOC.prototype.GetCaption = function()
{
	return this.Caption;
};
CFieldInstructionTOC.prototype.SetCaptionOnlyText = function (sVal)
{
	this.CaptionOnlyText = sVal;
};
CFieldInstructionTOC.prototype.GetCaptionOnlyText = function()
{
	return this.CaptionOnlyText;
};
CFieldInstructionTOC.prototype.SetPr = function(oPr)
{
	if (!(oPr instanceof Asc.CTableOfContentsPr))
		return;

	this.SetStylesArray(oPr.get_Styles());
	this.SetHeadingRange(oPr.get_OutlineStart(), oPr.get_OutlineEnd());
	this.SetHyperlinks(oPr.get_Hyperlink());

	if (oPr.PageNumbers)
		this.SetPageRefSkippedLvls(false);
	else
		this.SetPageRefSkippedLvls(true);

	if (oPr.RightTab)
		this.SetSeparator("");
	else
		this.SetSeparator(" ");

	this.ForceTabLeader = oPr.TabLeader;
	var sCaption = oPr.get_CaptionForInstruction();
	if(sCaption !== undefined)
	{
		if(sCaption || this.Styles.length > 0)
		{
			if(oPr.IsIncludeLabelAndNumber)
			{
				this.SetCaption(sCaption);
				this.SetCaptionOnlyText(undefined);
			}
			else
			{
				this.SetCaptionOnlyText(sCaption);
				this.SetCaption(undefined);
			}
			this.SetHeadingRange(-1, -1);
		}
	}
};
CFieldInstructionTOC.prototype.GetForceTabLeader = function()
{
	var nTabLeader = this.ForceTabLeader;
	this.ForceTabLeader = undefined;
	return nTabLeader;
};
CFieldInstructionTOC.prototype.ToString = function()
{
	var oLogicDocument = editor.WordControl.m_oLogicDocument;
	var sListSeparator = oLogicDocument.GetListSeparator();

	var sInstr = "TOC ";

	if (this.HeadingS >= 1
		&& this.HeadingS <= 9
		&& this.HeadingE >= this.HeadingS
		&& this.HeadingE <= 9)
		sInstr +=  "\\o " + "\"" + this.HeadingS + "-" + this.HeadingE + "\" ";

	if (this.SkipPageRef)
	{
		sInstr += "\\n ";

		if (this.SkipPageRefStart >= 1
			&& this.SkipPageRefStart <= 9
			&& this.SkipPageRefEnd >= this.SkipPageRefStart
			&& this.SkipPageRefEnd <= 9)
			sInstr +=  "\"" + this.SkipPageRefStart + "-" + this.SkipPageRefEnd + "\" ";
	}

	if (this.Hyperlinks)
		sInstr += "\\h ";

	if (!this.RemoveBreaks)
		sInstr += "\\x ";

	if (this.PreserveTabs)
		sInstr += "\\w ";

	if (this.Separator)
		sInstr += "\\p \"" + this.Separator + "\"";

	if (this.Styles.length > 0)
	{
		sInstr += "\\t \"";

		for (var nIndex = 0, nCount = this.Styles.length; nIndex < nCount; ++nIndex)
		{
			if (0 === nIndex)
				sInstr += this.Styles[nIndex].Name;
			else
				sInstr += sListSeparator + this.Styles[nIndex].Name;

			if (undefined !== this.Styles[nIndex].Lvl && null !== this.Styles[nIndex].Lvl)
				sInstr += sListSeparator + this.Styles[nIndex].Lvl;
		}

		sInstr += "\" ";
	}
	if(this.Caption !== undefined)
	{
		sInstr += "\\c ";
		if(typeof this.Caption === "string" && this.Caption.length > 0)
		{
			sInstr += "\"" + this.Caption + "\"";
		}
	}
	if(this.CaptionOnlyText !== undefined)
	{
		sInstr += "\\a ";
		if(typeof this.CaptionOnlyText === "string" && this.CaptionOnlyText.length > 0)
		{
			sInstr += "\"" + this.CaptionOnlyText + "\"";
		}
	}

	return sInstr;
};
CFieldInstructionTOC.prototype.IsTableOfFigures = function ()
{
	if(this.Caption !== undefined || this.CaptionOnlyText !== undefined)
	{
		return true;
	}
	return false;
};
CFieldInstructionTOC.prototype.IsTableOfContents = function ()
{
	return !this.IsTableOfFigures();
};
/**
 * ASK field
 * @constructor
 */
function CFieldInstructionASK()
{
	CFieldInstructionBase.call(this);

	this.BookmarkName = "";
	this.PromptText   = "";
}
CFieldInstructionASK.prototype = Object.create(CFieldInstructionBase.prototype);
CFieldInstructionASK.prototype.constructor = CFieldInstructionASK;
CFieldInstructionASK.prototype.Type = fieldtype_ASK;
CFieldInstructionASK.prototype.SetBookmarkName = function(sBookmarkName)
{
	this.BookmarkName = sBookmarkName;
};
CFieldInstructionASK.prototype.GetBookmarkName = function()
{
	return this.BookmarkName;
};
CFieldInstructionASK.prototype.SetPromptText = function(sText)
{
	this.PromptText = sText;
};
CFieldInstructionASK.prototype.GetPromptText = function()
{
	if (!this.PromptText)
		return this.BookmarkName;

	return this.PromptText;
};

/**
 * REF field
 * @constructor
 */
function CFieldInstructionREF()
{
	CFieldInstructionBase.call(this);

	this.GeneralSwitches = [];
	this.BookmarkName = "";
	this.Hyperlink = false; // \h - is hyperlink
	this.bIsNumberNoContext = false; // \n - paragraph number (no context)
	this.bIsNumberFullContext = false; // \w - paragraph number (full context)
	this.bIsNumber = false; // \r - paragraph number in realtive context
	this.bIsPosition = false; // \p - above/below
	this.Delimiter = null;
}
CFieldInstructionREF.prototype = Object.create(CFieldInstructionBase.prototype);
CFieldInstructionREF.prototype.constructor = CFieldInstructionREF;
CFieldInstructionREF.prototype.Type = fieldtype_REF;

CFieldInstructionREF.prototype.SetGeneralSwitches = function (aSwitches)
{
	this.GeneralSwitches = aSwitches;
};
CFieldInstructionREF.prototype.SetBookmarkName = function(sBookmarkName)
{
	this.BookmarkName = sBookmarkName;
};
CFieldInstructionREF.prototype.GetBookmarkName = function()
{
	return this.BookmarkName;
};
CFieldInstructionREF.prototype.SetHyperlink = function(bIsHyperlink)
{
	this.Hyperlink = bIsHyperlink;
};
CFieldInstructionREF.prototype.GetHyperlink = function()
{
	return this.Hyperlink;
};
CFieldInstructionREF.prototype.SetIsNumberNoContext = function(bVal)
{
	this.bIsNumberNoContext = bVal;
};
CFieldInstructionREF.prototype.IsNumberNoContext = function()
{
	return this.bIsNumberNoContext;
};
CFieldInstructionREF.prototype.SetIsNumberFullContext = function(bVal)
{
	this.bIsNumberFullContext = bVal;
};
CFieldInstructionREF.prototype.IsNumberFullContext = function()
{
	return this.bIsNumberFullContext;
};
CFieldInstructionREF.prototype.HaveNumberFlag = function()
{
	return this.IsNumber() || this.IsNumberFullContext() || this.IsNumberNoContext();
};
CFieldInstructionREF.prototype.SetIsNumber = function(bVal)
{
	this.bIsNumber = bVal;
};
CFieldInstructionREF.prototype.IsNumber = function()
{
	return this.bIsNumber;
};
CFieldInstructionREF.prototype.SetIsPosition = function(bVal)
{
	this.bIsPosition = bVal;
};
CFieldInstructionREF.prototype.IsPosition = function()
{
	return this.bIsPosition;
};
CFieldInstructionREF.prototype.SetDelimiter = function(bVal)
{
	this.Delimiter = bVal;
};
CFieldInstructionREF.prototype.GetDelimiter = function()
{
	return this.Delimiter;
};
CFieldInstructionREF.prototype.ToString = function()
{
	var sInstruction = " REF ";
	sInstruction += this.BookmarkName;
	for(var nSwitch = 0; i < this.GeneralSwitches.length; ++nSwitch)
	{
		sInstruction +=  " \\* " + this.GeneralSwitches[nSwitch];
	}
	if(this.GetHyperlink())
	{
		sInstruction += " \\h";
	}
	if(this.IsNumberNoContext())
	{
		sInstruction += " \\n";
	}
	if(this.IsNumberFullContext())
	{
		sInstruction += " \\w";
	}
	if(this.IsNumber())
	{
		sInstruction += " \\r"
	}
	if(this.IsPosition())
	{
		sInstruction += " \\p";
	}
	if(typeof this.Delimiter === "string" && this.Delimiter.length > 0)
	{
		sInstruction += " \\d " + this.Delimiter;
	}
	return sInstruction;
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для совместимости с обычным ParaHyperlink
//----------------------------------------------------------------------------------------------------------------------
CFieldInstructionREF.prototype.GetAnchor = function()
{
	var sBookmarkName = this.GetBookmarkName();
	var sAnchor = sBookmarkName;
	if(this.ComplexField)
	{
		var oLogicDoc = this.ComplexField.LogicDocument;
		if(oLogicDoc)
		{
			var oBookmarksManager = oLogicDoc.GetBookmarksManager();
			if(oBookmarksManager)
			{
				var oBookmark = oBookmarksManager.GetBookmarkByName(sBookmarkName);
				if(!oBookmark)
				{
					sAnchor = "_top";
				}
			}
		}
	}
	return sAnchor;
};
CFieldInstructionREF.prototype.GetValue = function()
{
	return "";
};
CFieldInstructionREF.prototype.SetVisited = function(isVisited)
{
};
/**
 * Проверяем является ли данная ссылка ссылкой в начало документа
 * @returns {boolean}
 */
CFieldInstructionREF.prototype.IsTopOfDocument = function()
{
	return (this.GetAnchor() === "_top");
};
CFieldInstructionREF.prototype.SetToolTip = function(sToolTip)
{
};
CFieldInstructionREF.prototype.GetToolTip = function()
{
	var sTooltip = this.BookmarkName;
	if(!sTooltip || '_' === sTooltip.charAt(0))
	{
		sTooltip = AscCommon.translateManager.getValue("Current Document");
	}
	return sTooltip;
};

/**
 * NUMPAGES field
 * @constructor
 */
function CFieldInstructionNUMPAGES()
{
	CFieldInstructionBase.call(this);
}
CFieldInstructionNUMPAGES.prototype = Object.create(CFieldInstructionBase.prototype);
CFieldInstructionNUMPAGES.prototype.constructor = CFieldInstructionNUMPAGES;
CFieldInstructionNUMPAGES.prototype.Type = fieldtype_NUMPAGES;

/**
 * HYPERLINK field
 * @constructor
 */
function CFieldInstructionHYPERLINK()
{
	CFieldInstructionBase.call(this);

	this.ToolTip      = "";
	this.Link         = "";
	this.BookmarkName = "";
}
CFieldInstructionHYPERLINK.prototype = Object.create(CFieldInstructionBase.prototype);
CFieldInstructionHYPERLINK.prototype.constructor = CFieldInstructionHYPERLINK;
CFieldInstructionHYPERLINK.prototype.Type = fieldtype_HYPERLINK;
CFieldInstructionHYPERLINK.prototype.SetToolTip = function(sToolTip)
{
	this.ToolTip = sToolTip;
};
CFieldInstructionHYPERLINK.prototype.GetToolTip = function()
{
	if ("" === this.ToolTip)
	{
		if (this.Link)
			return (this.BookmarkName ? this.Link + "#" + this.BookmarkName : this.Link)
		else if (this.BookmarkName)
			return AscCommon.translateManager.getValue("Current Document");

		return "";
	}

	return this.ToolTip;
};
CFieldInstructionHYPERLINK.prototype.SetLink = function(sLink)
{
	this.Link = sLink;
};
CFieldInstructionHYPERLINK.prototype.GetLink = function()
{
	return this.Link;
};
CFieldInstructionHYPERLINK.prototype.SetBookmarkName = function(sBookmarkName)
{
	this.BookmarkName = sBookmarkName;
};
CFieldInstructionHYPERLINK.prototype.GetBookmarkName = function()
{
	return this.BookmarkName;
};
CFieldInstructionHYPERLINK.prototype.ToString = function()
{
	var sInstr = "HYPERLINK ";
	if (this.Link)
		sInstr +=  "\"" + this.Link + "\"";

	if (this.ToolTip)
		sInstr += "\\o \"" + this.ToolTip + "\"";

	if (this.BookmarkName)
		sInstr += "\\l " + this.BookmarkName;

	return sInstr;
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для совместимости с обычным ParaHyperlink
//----------------------------------------------------------------------------------------------------------------------
CFieldInstructionHYPERLINK.prototype.GetAnchor = function()
{
	return this.GetBookmarkName();
};
CFieldInstructionHYPERLINK.prototype.GetValue = function()
{
	return this.GetLink();
};
CFieldInstructionHYPERLINK.prototype.SetVisited = function(isVisited)
{
};
/**
 * Проверяем является ли данная ссылка ссылкой в начало документа
 * @returns {boolean}
 */
CFieldInstructionHYPERLINK.prototype.IsTopOfDocument = function()
{
	return (this.GetBookmarkName() === "_top");
};

/**
 * TIME field
 * @constructor
 */
function CFieldInstructionTIME()
{
	CFieldInstructionBase.call(this);

	this.Format = "";
}
CFieldInstructionTIME.prototype = Object.create(CFieldInstructionBase.prototype);
CFieldInstructionTIME.prototype.constructor = CFieldInstructionTIME;
CFieldInstructionTIME.prototype.Type = fieldtype_TIME;
CFieldInstructionTIME.prototype.ToString = function()
{
	return ("TIME \\@ \"" + this.sFormat + "\"");
};
CFieldInstructionTIME.prototype.SetFormat = function(sFormat)
{
	this.Format = sFormat;
};
CFieldInstructionTIME.prototype.GetFormat = function()
{
	return this.Format;
};

/**
 * TIME field
 * @constructor
 */
function CFieldInstructionDATE()
{
	CFieldInstructionBase.call(this);

	this.Format = "";
}
CFieldInstructionDATE.prototype = Object.create(CFieldInstructionBase.prototype);
CFieldInstructionDATE.prototype.constructor = CFieldInstructionDATE;
CFieldInstructionDATE.prototype.Type = fieldtype_DATE;
CFieldInstructionDATE.prototype.ToString = function()
{
	return ("TIME \\@ \"" + this.sFormat + "\"");
};
CFieldInstructionDATE.prototype.SetFormat = function(sFormat)
{
	this.Format = sFormat;
};
CFieldInstructionDATE.prototype.GetFormat = function()
{
	return this.Format;
};
/**
 * SEQ field
 * @constructor
 */
function CFieldInstructionSEQ()
{

	CFieldInstructionBase.call(this);
	this.Id = null;
	this.C = false;
	this.H = false;
	this.N = false;
	this.R = null;
	this.S = null;
	this.NumFormat = Asc.c_oAscNumberingFormat.Decimal;
	this.GeneralSwitches = [];

	this.ParentContent = null;
}
CFieldInstructionSEQ.prototype = Object.create(CFieldInstructionBase.prototype);
CFieldInstructionSEQ.prototype.constructor = CFieldInstructionSEQ;
CFieldInstructionSEQ.prototype.Type = fieldtype_SEQ;
CFieldInstructionSEQ.prototype.ToString = function ()
{
	var sInstruction = " SEQ ";
	if(this.Id)
	{
		sInstruction += this.Id;
	}
	for(var i = 0; i < this.GeneralSwitches.length; ++i)
	{
		sInstruction +=  " \\* " + this.GeneralSwitches[i];
	}
	if(this.C)
	{
		sInstruction += " \\c"
	}
	if(this.H)
	{
		sInstruction += " \\h";
	}
	if(this.R)
	{
		sInstruction += " \\r " + this.R;
	}
	if(this.S)
	{
		sInstruction += " \\s " + this.S;
	}
	return sInstruction;
};
CFieldInstructionSEQ.prototype.SetComplexField = function (oComplexField)
{
	CFieldInstructionBase.prototype.SetComplexField.call(this, oComplexField);
	this.ParentContent = null;
	var oBeginChar = oComplexField.BeginChar;
	if(oBeginChar)
	{
		var oRun = oBeginChar.Run;
		if(oRun)
		{
			var oParagraph = oRun.Paragraph;
			if(oParagraph)
			{
				this.ParentContent = oParagraph.Parent;
			}
		}
	}
};

CFieldInstructionSEQ.prototype.GetRestartNum = function ()
{
	if(typeof this.R === "string" && this.R.length > 0)
	{
		var aTest = /[0-9]+/.exec(this.R);
		var nResult;
		if(Array.isArray(aTest) && aTest.length > 0)
		{
			nResult = parseInt(aTest[0]);
			if(!isNaN(nResult))
			{
				return nResult;
			}
		}
	}
	return null;
};

CFieldInstructionSEQ.prototype.GetText = function ()
{
	if(!this.ParentContent)
	{
		return "";
	}
	var oTopDocument = this.ParentContent.Is_TopDocument(true);
	var aFields, oField, i, nIndex, nLvl, nCounter;
	if(!oTopDocument)
	{
		return "";
	}
	if(oTopDocument.IsHdrFtr(false) || oTopDocument.IsFootnote(false))
	{
		return AscCommon.translateManager.getValue("Error! Main Document Only.");
	}

	if(this.H)
	{
		if(this.GeneralSwitches.length === 0)
		{
			return "";
		}
	}
	nIndex = this.GetRestartNum();
	if(nIndex === null)
	{
		aFields = [];
		oTopDocument.GetAllSeqFieldsByType(this.Id, aFields);
		nIndex = -1;
		if(this.S)
		{
			nLvl = parseInt(this.S);
			if(!isNaN(nLvl))
			{
				--nLvl;
				for(i = aFields.length - 1; i > -1; --i)
				{
					oField = aFields[i];
					if(AscCommon.isRealObject(oField) && this.ComplexField === oField)
					{
						break;
					}
				}
				if(i > -1)
				{
					nCounter = i;
					for(i = i - 1; i > -1; --i)
					{
						oField = aFields[i];
						if(AscFormat.isRealNumber(oField) && oField >= nLvl)
						{
							aFields = aFields.splice(i + 1, nCounter - i);
							break;
						}
					}
				}
			}
		}
		nCounter = 1;
		for(i = 0; i < aFields.length; ++i)
		{
			oField = aFields[i];
			if(AscCommon.isRealObject(oField))
			{
				if(this.ComplexField === oField)
				{
					nIndex = nCounter;
					break;
				}
				if(!(oField.Instruction && oField.Instruction.C))
				{
					++nCounter;
				}
			}
		}
	}
	if(nIndex > -1)
	{
		return AscCommon.IntToNumberFormat(nIndex, this.NumFormat);
	}
	return AscCommon.translateManager.getValue("Error! Main Document Only.");
};
CFieldInstructionSEQ.prototype.SetId = function (sVal)
{
	this.Id = sVal;
};
CFieldInstructionSEQ.prototype.SetC = function (sVal)
{
	this.C = sVal;
};
CFieldInstructionSEQ.prototype.SetH = function (sVal)
{
	this.H = sVal;
};
CFieldInstructionSEQ.prototype.SetN = function (sVal)
{
	this.N = sVal;
};
CFieldInstructionSEQ.prototype.SetR = function (sVal)
{
	this.R = sVal;
};
CFieldInstructionSEQ.prototype.SetS = function (sVal)
{
	this.S = sVal;
};
CFieldInstructionSEQ.prototype.SetGeneralSwitches = function (aSwitches)
{
	this.GeneralSwitches = aSwitches;
	for(var i = 0; i < aSwitches.length; ++i)
	{
		this.NumFormat = GeneralToNumFormat(aSwitches[i]);
	}
};

function GeneralToNumFormat(sFormat)
{
	if(typeof sFormat === 'string')
	{
		if(sFormat.toLowerCase() === 'arabic')
		{
			return Asc.c_oAscNumberingFormat.Decimal;
		}
		else if(sFormat.toLowerCase() === 'alphabetic')
		{
			if(sFormat[0] === 'A')
			{
				return Asc.c_oAscNumberingFormat.UpperLetter;
			}
			else
			{
				return Asc.c_oAscNumberingFormat.LowerLetter;
			}
		}
		else if(sFormat.toLowerCase() === 'roman')
		{
			if(sFormat[0] === 'r')
			{
				return Asc.c_oAscNumberingFormat.LowerRoman;
			}
			else
			{
				return Asc.c_oAscNumberingFormat.UpperRoman;
			}
		}
	}
	return Asc.c_oAscNumberingFormat.Decimal;
}

/**
 * STYLEREF field
 * @constructor
 */

function CFieldInstructionSTYLEREF()
{

	CFieldInstructionBase.call(this);
	this.StyleName = null;
	this.L = null;
	this.N = null;
	this.P = null;
	this.R = null;
	this.T = null;
	this.W = null;
	this.S = null;
	this.GeneralSwitches = [];
	this.ParentContent = null;
	this.ParentParagraph = null;
}
CFieldInstructionSTYLEREF.prototype = Object.create(CFieldInstructionBase.prototype);
CFieldInstructionSTYLEREF.prototype.constructor = CFieldInstructionSTYLEREF;
CFieldInstructionSTYLEREF.prototype.Type = fieldtype_STYLEREF;
CFieldInstructionSTYLEREF.prototype.SetL = function(v){this.L = v;};
CFieldInstructionSTYLEREF.prototype.SetN = function(v){this.N = v;};
CFieldInstructionSTYLEREF.prototype.SetP = function(v){this.P = v;};
CFieldInstructionSTYLEREF.prototype.SetR = function(v){this.R = v;};
CFieldInstructionSTYLEREF.prototype.SetT = function(v){this.T = v;};
CFieldInstructionSTYLEREF.prototype.SetW = function(v){this.W = v;};
CFieldInstructionSTYLEREF.prototype.SetS = function(v){this.S = v;};
CFieldInstructionSTYLEREF.prototype.SetGeneralSwitches = function(v){this.GeneralSwitches = v;};
CFieldInstructionSTYLEREF.prototype.GetText = function()
{
	var sDefaultMessage = "Error! No text of specified style in document.";
	if(this.ParentContent)
	{
		var oHdrFtr = this.ParentContent.IsHdrFtr(true);
		if (oHdrFtr)
		{
			//TODO
		}
		else
		{
			var oFootNote = this.ParentContent.IsFootnote(true);
			if(oFootNote)
			{
				//TODO
			}
			else
			{
				if(this.ParentParagraph)
				{
					var oParagraph = null;
					var sRet = "";
					var bAbove = true;
					var oStyles = this.ParentContent.Styles;
					var sId = oStyles.GetStyleIdByName(this.StyleName);
					var nStartIndex, oTmpContent;
					var oShape, oMainGroup, oDrawing, oCell, oRow, oTable, oBLSdt;
					var oParentParagraph, oParentContent, nParentIdx;
					if(sId)
					{
						oParentParagraph = this.ParentParagraph;
						oParentContent = this.ParentContent;
						nParentIdx = this.ParentParagraph.GetIndex();
						oShape = this.ParentContent.Is_DrawingShape(true);
						if(oShape)
						{
							if(oShape.group)
							{
								oMainGroup = oShape.getMainGroup();
								oDrawing = oMainGroup.parent;
							}
							else
							{
								oDrawing = oShape.parent;
							}
							if(!oDrawing)
							{
								return AscCommon.translateManager.getValue(sDefaultMessage);
							}
							oParentParagraph = oDrawing.GetParagraph();
							oParentContent = oParentParagraph.GetParent();
							nParentIdx = oParentParagraph.GetIndex();
						}
						if(oParentParagraph.GetParagraphStyle() === sId)
						{
							oParagraph = oParentParagraph;
						}
						oTmpContent = oParentContent;
						nStartIndex = nParentIdx;
						while(oTmpContent && !oParagraph)
						{
							oParagraph = oTmpContent.FindParaWithStyle(sId, true, nStartIndex);
							if(oParagraph)
							{
								break;
							}
							oCell = oTmpContent.IsTableCellContent(true);
							if(oCell)
							{
								oRow = oCell.GetRow();
								oTable = oRow.GetTable();
								if(!oRow || !oTable)
								{
									return AscCommon.translateManager.getValue(sDefaultMessage);
								}
								oParagraph = oRow.FindParaWithStyle(sId, true, oCell.GetIndex() - 1);
								if(!oParagraph)
								{
									oParagraph = oTable.FindParaWithStyle(sId, true, oRow.GetIndex() - 1);
								}
								oTmpContent = oTable.Parent;
								nStartIndex = oTable.GetIndex() - 1;
							}
							else if(oTmpContent.IsBlockLevelSdtContent())
							{
								oBLSdt = oTmpContent.GetParent();
								oTmpContent = oBLSdt.Parent;
								nStartIndex = oBLSdt.GetIndex() - 1;
							}
							else
							{
								break;
							}
						}
						if(!oParagraph)
						{
							oTmpContent = oParentContent;
							nStartIndex = nParentIdx + 1;
							while(oTmpContent && !oParagraph)
							{
								oParagraph = oTmpContent.FindParaWithStyle(sId, false, nStartIndex);
								if(oParagraph)
								{
									break;
								}
								oCell = oTmpContent.IsTableCellContent(true);
								if(oCell)
								{
									oRow = oCell.GetRow();
									oTable = oRow.GetTable();
									if(!oRow || !oTable)
									{
										return AscCommon.translateManager.getValue(sDefaultMessage);
									}
									oParagraph = oRow.FindParaWithStyle(sId, false, oCell.GetIndex() + 1);
									if(!oParagraph)
									{
										oParagraph = oTable.FindParaWithStyle(sId, false, oRow.GetIndex() + 1);
									}
									oTmpContent = oTable.Parent;
									nStartIndex = oTable.GetIndex() + 1;
								}
								else if(oTmpContent.IsBlockLevelSdtContent())
								{
									oBLSdt = oTmpContent.GetParent();
									oTmpContent = oBLSdt.Parent;
									nStartIndex = oBLSdt.GetIndex() + 1;
								}
								else
								{
									break;
								}
							}
						}
						if(oParagraph)
						{
							if(this.N || this.R || this.W || this.S)
							{
								if(oParagraph.IsNumberedNumbering())
								{
									sRet += oParagraph.GetNumberingText(true);
								}
								else
								{
									sRet += "0";
								}
							}
							else
							{
								oParagraph.ApplyToAll = true;
								sRet = oParagraph.GetSelectedText(true, {});
								oParagraph.ApplyToAll = false;
							}
							if(this.P)
							{
								sRet += (" " + AscCommon.translateManager.getValue(bAbove ? "above" : "below"));
							}
							return sRet;
						}
					}
					return AscCommon.translateManager.getValue(sDefaultMessage);
				}
			}
		}
	}
    return AscCommon.translateManager.getValue(sDefaultMessage);
};
CFieldInstructionSTYLEREF.prototype.SetStyleName = function(v)
{
	this.StyleName = v;
};
CFieldInstructionSTYLEREF.prototype.ToString = function()
{
	var sRet = " STYLEREF ";

	if(this.S)
	{
		sRet += " \\s"
	}
	if(this.StyleName)
	{
		sRet += this.StyleName;
	}
	if(this.L)
	{
		sRet += " \\l";
	}
	if(this.N)
	{
		sRet += " \\n"
	}
	if(this.P)
	{
		sRet += " \\p"
	}
	if(this.R)
	{
		sRet += " \\r"
	}
	if(this.T)
	{
		sRet += " \\t"
	}
	if(this.W)
	{
		sRet += " \\w"
	}
	return sRet;
};
CFieldInstructionSTYLEREF.prototype.SetComplexField = function (oComplexField)
{
	CFieldInstructionBase.prototype.SetComplexField.call(this, oComplexField);
	this.ParentContent = null;
	var oBeginChar = oComplexField.BeginChar;
	if(oBeginChar)
	{
		var oRun = oBeginChar.Run;
		if(oRun)
		{
			var oParagraph = oRun.Paragraph;
			if(oParagraph)
			{
				this.ParentParagraph = oParagraph;
				this.ParentContent = oParagraph.Parent;
			}
		}
	}
};


/**
 * NOTEREF field
 * @constructor
 */
function CFieldInstructionNOTEREF()
{
	CFieldInstructionBase.call(this);

	this.GeneralSwitches = [];
	this.BookmarkName = "";
	this.Hyperlink = false; // \h - is hyperlink
	this.bIsPosition = false; // \p - above/below
	this.bFormatting = false; // \p - format as notes or endnotes
}
CFieldInstructionNOTEREF.prototype = Object.create(CFieldInstructionBase.prototype);
CFieldInstructionNOTEREF.prototype.constructor = CFieldInstructionNOTEREF;
CFieldInstructionNOTEREF.prototype.Type = fieldtype_NOTEREF;


CFieldInstructionNOTEREF.prototype.SetGeneralSwitches = function (aSwitches)
{
	this.GeneralSwitches = aSwitches;
};
CFieldInstructionNOTEREF.prototype.SetBookmarkName = function(sBookmarkName)
{
	this.BookmarkName = sBookmarkName;
};
CFieldInstructionNOTEREF.prototype.GetBookmarkName = function()
{
	return this.BookmarkName;
};
CFieldInstructionNOTEREF.prototype.SetHyperlink = function(bIsHyperlink)
{
	this.Hyperlink = bIsHyperlink;
};
CFieldInstructionNOTEREF.prototype.GetHyperlink = function()
{
	return this.Hyperlink;
};
CFieldInstructionNOTEREF.prototype.SetIsPosition = function(bVal)
{
	this.bIsPosition = bVal;
};
CFieldInstructionNOTEREF.prototype.IsPosition = function()
{
	return this.bIsPosition;
};
CFieldInstructionNOTEREF.prototype.SetIsFormatting = function(bVal)
{
	this.bFormatting = bVal;
};
CFieldInstructionNOTEREF.prototype.IsFormatting = function()
{
	return this.bFormatting;
};
CFieldInstructionNOTEREF.prototype.ToString = function()
{
	var sInstruction = " NOTEREF ";
	sInstruction += this.BookmarkName;
	for(var nSwitch = 0; i < this.GeneralSwitches.length; ++nSwitch)
	{
		sInstruction +=  " \\* " + this.GeneralSwitches[nSwitch];
	}
	if(this.GetHyperlink())
	{
		sInstruction += " \\h";
	}
	if(this.IsPosition())
	{
		sInstruction += " \\p";
	}
	if(this.IsFormatting())
	{
		sInstruction += " \\f";
	}
	return sInstruction;
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для совместимости с обычным ParaHyperlink
//----------------------------------------------------------------------------------------------------------------------
CFieldInstructionNOTEREF.prototype.IsTopOfDocument = function()
{
	return (this.GetAnchor() === "_top");
};
CFieldInstructionNOTEREF.prototype.GetAnchor = function()
{
	var sBookmarkName = this.GetBookmarkName();
	var sAnchor = sBookmarkName;
	if(this.ComplexField)
	{
		var oLogicDoc = this.ComplexField.LogicDocument;
		if(oLogicDoc)
		{
			var oBookmarksManager = oLogicDoc.GetBookmarksManager();
			if(oBookmarksManager)
			{
				var oBookmark = oBookmarksManager.GetBookmarkByName(sBookmarkName);
				if(!oBookmark)
				{
					sAnchor = "_top";
				}
			}
		}
	}
	return sAnchor;
};
CFieldInstructionNOTEREF.prototype.GetValue = function()
{
	return "";
};
CFieldInstructionNOTEREF.prototype.SetVisited = function(isVisited)
{
};

CFieldInstructionREF.prototype.SetToolTip = function(sToolTip)
{
};
CFieldInstructionNOTEREF.prototype.GetToolTip = function()
{
	return AscCommon.translateManager.getValue("Current Document");
};



/**
 * Класс для разбора строки с инструкцией
 * @constructor
 */
function CFieldInstructionParser()
{
	this.Line   = "";
	this.Pos    = 0;
	this.Buffer = "";
	this.Result = null;

	this.SavedStates = [];
}
CFieldInstructionParser.prototype.GetInstructionClass = function(sLine)
{
	this.Line   = sLine;
	this.Pos    = 0;
	this.Buffer = "";
	this.Result = null;

	this.private_Parse();

	return this.Result;
};
CFieldInstructionParser.prototype.InitParaFieldArguments = function(Type, sLine, oParaFiles)
{
	this.Line   = sLine;
	this.Pos    = 0;
	this.Buffer = "";
	this.Result = null;

	oParaFiles.FieldType = Type;
	oParaFiles.Arguments = this.private_ReadArguments().slice(1);
	//todo Switches
	oParaFiles.Switches = this.Pos < this.Line.length ? [this.Line.substring(this.Pos)] : [];
};
CFieldInstructionParser.prototype.private_Parse = function()
{
	if (!this.private_ReadNext())
		return this.private_ReadREF("");
	
	var sBuffer = this.Buffer.toUpperCase();
	if("PAGE" === sBuffer)
	{
		this.private_ReadPAGE();
	}
	else if("PAGEREF" === sBuffer)
	{
		this.private_ReadPAGEREF();
	}
	else if("TOC" === sBuffer)
	{
		this.private_ReadTOC();
	}
	else if("ASC" === sBuffer)
	{
		this.private_ReadASK();
	}
	else if("REF" === sBuffer)
	{
		this.private_ReadREF();
	}
	else if("NOTEREF" === sBuffer)
	{
		this.private_ReadNOTEREF();
	}
	else if("NUMPAGES" === sBuffer)
	{
		this.private_ReadNUMPAGES();
	}
	else if("HYPERLINK" === sBuffer)
	{
		this.private_ReadHYPERLINK();
	}
	else if("SEQ" === sBuffer)
	{
		this.private_ParseSEQ();
	}
	else if("STYLEREF" === sBuffer)
	{
		this.private_ParseSTYLEREF();
	}
	else if ("TIME" === sBuffer)
	{
		this.private_ReadTIME();
	}
	else if ("DATE" === sBuffer)
	{
		this.private_ReadDATE();
	}
	else if ("ADDIN" === sBuffer)
	{
		this.private_ReadADDIN();
	}
	else if(sBuffer.indexOf("=") === 0)
	{
		this.private_ReadFORMULA();
	}
	else
	{
		this.private_ReadREF();
	}
};
CFieldInstructionParser.prototype.private_ReadNext = function()
{
	var nLen  = this.Line.length,
		bWord = false;

	this.Buffer = "";

	while (this.Pos < nLen)
	{
		var nCharCode = this.Line.charCodeAt(this.Pos);
		if (32 === nCharCode || 9 === nCharCode)
		{
			if (bWord)
				return true;
		}
		else if (34 === nCharCode && (0 === this.Pos || 92 !== this.Line.charCodeAt(this.Pos - 1)))
		{
			// Кавычки
			this.Pos++;
			while (this.Pos < nLen)
			{
				nCharCode = this.Line.charCodeAt(this.Pos);
				if (34 === nCharCode && 92 !== this.Line.charCodeAt(this.Pos - 1))
				{
					this.Pos++;
					break;
				}

				bWord = true;

				if (34 === nCharCode && 92 === this.Line.charCodeAt(this.Pos - 1) && this.Buffer.length > 0)
					this.Buffer = this.Buffer.substring(0, this.Buffer.length - 1);

				this.Buffer += this.Line.charAt(this.Pos);

				this.Pos++;
			}

			return bWord;
		}
		else
		{
			this.Buffer += this.Line.charAt(this.Pos);
			bWord = true;
		}

		this.Pos++;
	}

	if (bWord)
		return true;

	return false;
};
CFieldInstructionParser.prototype.private_ReadTillEnd = function()
{
	this.Buffer = this.Line.substr(this.Pos).trim();
	return !!this.Buffer;
};
CFieldInstructionParser.prototype.private_ReadArguments = function()
{
	var arrArguments = [];

	var sArgument = this.private_ReadArgument();
	while (null !== sArgument)
	{
		arrArguments.push(sArgument);
		sArgument = this.private_ReadArgument();
	}

	return arrArguments;
};
CFieldInstructionParser.prototype.private_ReadArgument = function()
{
	this.private_SaveState();

	if (!this.private_ReadNext())
		return null;

	if (this.private_IsSwitch())
	{
		this.private_RestoreState();
		return null;
	}

	this.private_RemoveLastState();
	return this.Buffer;
};
CFieldInstructionParser.prototype.private_IsSwitch = function()
{
	return this.Buffer.charAt(0) === '\\';
};
CFieldInstructionParser.prototype.private_GetSwitchLetter = function()
{
	return this.Buffer.charAt(1);
};
CFieldInstructionParser.prototype.private_SaveState = function()
{
	this.SavedStates.push(this.Pos);
};
CFieldInstructionParser.prototype.private_RestoreState = function()
{
	if (this.SavedStates.length > 0)
		this.Pos = this.SavedStates[this.SavedStates.length - 1];

	this.private_RemoveLastState();
};
CFieldInstructionParser.prototype.private_RemoveLastState = function()
{
	if (this.SavedStates.length > 0)
		this.SavedStates.splice(this.SavedStates.length - 1, 1);
};
CFieldInstructionParser.prototype.private_ReadGeneralFormatSwitch = function()
{
	if (!this.private_IsSwitch() || this.Buffer.charAt(1) !== '*')
		return;

	if (!this.private_ReadNext() || this.private_IsSwitch())
		return;

	// TODO: Тут надо прочитать поле

	//console.log("General switch: " + this.Buffer);
};
CFieldInstructionParser.prototype.private_ReadPAGE = function()
{
	this.Result = new CFieldInstructionPAGE();

	// Zero or more general-formatting-switches

	while (this.private_ReadNext())
	{
		if (this.private_IsSwitch())
			this.private_ReadGeneralFormatSwitch();
	}
};
CFieldInstructionParser.prototype.private_ReadFORMULA = function()
{
	this.Result = new CFieldInstructionFORMULA();
    var sFormula = this.Buffer.slice(1, this.Buffer.length);
    var sFormat = null;
    var bFormat = false;
    var bNumFormat = false;
    while(this.private_ReadNext())
	{
    	if(this.private_IsSwitch())
    	{
			bFormat = true;
			if ('#' === this.Buffer.charAt(1))
			{
				bNumFormat = true;
			}
		}
		else
		{
			if(bFormat)
			{
				if(bNumFormat)
				{
                    sFormat = this.Buffer;
				}
			}
			else
			{
                sFormula += this.Buffer;
			}
            bFormat = false;
            bNumFormat = false;
		}
	}
	sFormula = sFormula.toUpperCase();
	var oFormat;
	if(null !== sFormat)
	{
        oFormat = AscCommon.oNumFormatCache.get(sFormat, AscCommon.NumFormatType.WordFieldNumeric);
        this.Result.SetFormat(oFormat);
	}
	this.Result.SetFormula(sFormula);

};
CFieldInstructionParser.prototype.private_ReadPAGEREF = function()
{
	var sBookmarkName = null;
	var isHyperlink = false, isPageRel = false;

	var isSwitch = false, isBookmark = false;

	while (this.private_ReadNext())
	{
		if (this.private_IsSwitch())
		{
			isSwitch = true;

			if ('p' === this.Buffer.charAt(1))
				isPageRel = true;
			else if ('h' === this.Buffer.charAt(1))
				isHyperlink = true;
		}
		else if (!isSwitch && !isBookmark)
		{
			sBookmarkName = this.Buffer;
			isBookmark    = true;
		}
	}

	this.Result = new CFieldInstructionPAGEREF(sBookmarkName, isHyperlink, isPageRel);
};
CFieldInstructionParser.prototype.private_ReadTOC = function()
{
	// TODO: \a, \b, \c, \d, \f, \l, \s, \z, \u

	this.Result = new CFieldInstructionTOC();
	var arrArguments;

	var isOutline = false, isStyles = false;
	while (this.private_ReadNext())
	{
		if (this.private_IsSwitch())
		{
			var sType = this.private_GetSwitchLetter();
			if ('w' === sType)
			{
				this.Result.SetPreserveTabs(true);
			}
			else if ('x' === sType)
			{
				this.Result.SetRemoveBreaks(false);
			}
			else if ('h' === sType)
			{
				this.Result.SetHyperlinks(true);
			}
			else if ('p' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
					this.Result.SetSeparator(arrArguments[0]);
			}
			else if ('o' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
				{
					var arrRange = this.private_ParseIntegerRange(arrArguments[0]);
					if (null !== arrRange)
						this.Result.SetHeadingRange(arrRange[0], arrRange[1]);
				}
				else
				{
					this.Result.SetHeadingRange(1, 9);
				}

				isOutline = true;
			}
			else if ('t' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
				{
					this.Result.SetStylesArrayRaw(arrArguments[0]);
					isStyles = true;
				}
			}
			else if ('n' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
				{
					var arrRange = this.private_ParseIntegerRange(arrArguments[0]);
					if (null !== arrRange)
						this.Result.SetPageRefSkippedLvls(true, arrRange[0], arrRange[1]);
					else
						this.Result.SetPageRefSkippedLvls(true, -1, -1);
				}
				else
				{
					this.Result.SetPageRefSkippedLvls(true, -1, -1);
				}
			}
			else if('c' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if(arrArguments.length > 0)
				{
					var sCaption = arrArguments[0];
					if(typeof sCaption === "string" && sCaption.length > 0)
					{
						this.Result.SetCaption(sCaption);
					}
					else
					{
						this.Result.SetCaption(null);
					}
				}
				else
				{
					this.Result.SetCaption(null);
				}
			}
			else if('a' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if(arrArguments.length > 0)
				{
					var sCaptionOnlyText = arrArguments[0];
					if(typeof sCaptionOnlyText === "string" && sCaptionOnlyText.length > 0)
					{
						this.Result.SetCaptionOnlyText(sCaptionOnlyText);
					}
					else
					{
						this.Result.SetCaptionOnlyText(null);
					}
				}
				else
				{
					this.Result.SetCaptionOnlyText(null);
				}
			}
		}
	}

	if (!isOutline && !isStyles)
		this.Result.SetHeadingRange(1, 9);
};
CFieldInstructionParser.prototype.private_ReadASK = function()
{
	this.Result = new CFieldInstructionASK();

	var arrArguments = this.private_ReadArguments();

	if (arrArguments.length >= 2)
		this.Result.SetPromptText(arrArguments[1]);

	if (arrArguments.length >= 1)
		this.Result.SetBookmarkName(arrArguments[0]);

	// TODO: Switches
};
CFieldInstructionParser.prototype.private_ReadREF = function(sBookmarkName)
{
	this.Result = new CFieldInstructionREF();

	if (undefined !== sBookmarkName)
	{
		this.Result.SetBookmarkName(sBookmarkName);
	}
	else
	{
		var arrArguments = this.private_ReadArguments();
		if (arrArguments.length > 0)
		{
			this.Result.SetBookmarkName(arrArguments[0]);
		}
	}
	while (this.private_ReadNext())
	{
		if (this.private_IsSwitch())
		{
			var sType = this.private_GetSwitchLetter();
			if ('*' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
					this.Result.SetGeneralSwitches(arrArguments);
			}
			else if("d" === sType)
			{
				arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
				{
					if(typeof arrArguments[0] === "string" && arrArguments[0].length > 0)
					{
						this.Result.SetDelimiter(arrArguments[0]);
					}
				}
			}
			else if("h" === sType)
			{
				this.Result.SetHyperlink(true);
			}
			else if("n" === sType)
			{
				this.Result.SetIsNumberNoContext(true);
			}
			else if("w" === sType)
			{
				this.Result.SetIsNumberFullContext(true);
			}
			else if("r" === sType)
			{
				this.Result.SetIsNumber(true);
			}
			else if("p" === sType)
			{
				this.Result.SetIsPosition(true);
			}
		}
	}
};
CFieldInstructionParser.prototype.private_ReadNOTEREF = function()
{
	this.Result = new CFieldInstructionNOTEREF();

	var arrArguments = this.private_ReadArguments();
	if (arrArguments.length > 0)
	{
		this.Result.SetBookmarkName(arrArguments[0]);
	}
	while (this.private_ReadNext())
	{
		if (this.private_IsSwitch())
		{
			var sType = this.private_GetSwitchLetter();
			if ('*' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
					this.Result.SetGeneralSwitches(arrArguments);
			}
			else if("h" === sType) {
				this.Result.SetHyperlink(true);
			}
			else if("f" === sType) {
				this.Result.SetIsFormatting(true);
			}
			else if("p" === sType) {
				this.Result.SetIsPosition(true);
			}
		}
	}
};
CFieldInstructionParser.prototype.private_ReadNUMPAGES = function()
{
	this.Result = new CFieldInstructionNUMPAGES();

	// TODO: Switches
};
CFieldInstructionParser.prototype.private_ReadHYPERLINK = function()
{
	this.Result = new CFieldInstructionHYPERLINK();
	var arrArguments = this.private_ReadArguments();
	if (arrArguments.length > 0)
		this.Result.SetLink(arrArguments[0]);

	while (this.private_ReadNext())
	{
		if (this.private_IsSwitch())
		{
			var sType = this.private_GetSwitchLetter();
			if ('o' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
					this.Result.SetToolTip(arrArguments[0]);
			}
			else if ('l' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
					this.Result.SetBookmarkName(arrArguments[0]);
			}

			// TODO: Остальные флаги \m \n \t для нас бесполезны
		}
	}
};
CFieldInstructionParser.prototype.private_ParseIntegerRange = function(sValue)
{
	// value1-value2

	var nSepPos = sValue.indexOf("-");
	if (-1 === nSepPos)
		return null;

	var nValue1 = parseInt(sValue.substr(0, nSepPos));
	var nValue2 = parseInt(sValue.substr(nSepPos + 1));

	if (isNaN(nValue1) || isNaN(nValue2))
		return null;

	return [nValue1, nValue2];
};
CFieldInstructionParser.prototype.private_ParseSEQ = function()
{
	this.Result = new CFieldInstructionSEQ();
	var arrArguments = this.private_ReadArguments();
	if (arrArguments.length > 0)
		this.Result.SetId(arrArguments[0]);

	while (this.private_ReadNext())
	{
		if (this.private_IsSwitch())
		{
			var sType = this.private_GetSwitchLetter();
			if ('*' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
					this.Result.SetGeneralSwitches(arrArguments);
			}
			else if ('c' === sType)
			{
				this.Result.SetC(true);
			}
			else if ('h' === sType)
			{
				this.Result.SetH(true);
			}
			else if ('n' === sType)
			{
				this.Result.SetN(true);
			}
			else if ('r' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
				{
					this.Result.SetR(arrArguments[0]);
				}
			}
			else if('s' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
				{
					this.Result.SetS(arrArguments[0]);
				}
			}
		}
	}
};
CFieldInstructionParser.prototype.private_ParseSTYLEREF = function()
{
	this.Result = new CFieldInstructionSTYLEREF();
	var arrArguments = this.private_ReadArguments();
	if (arrArguments.length > 0)
		this.Result.SetStyleName(arrArguments[0]);

	while (this.private_ReadNext())
	{
		if (this.private_IsSwitch())
		{
			var sType = this.private_GetSwitchLetter();
			if ('*' === sType)
			{
				arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
					this.Result.SetGeneralSwitches(arrArguments);
			}
			else if ('l' === sType)
			{
				this.Result.SetL(true);
			}
			else if ('n' === sType)
			{
				this.Result.SetN(true);
			}
			else if ('p' === sType)
			{
				this.Result.SetP(true);
			}
			else if ('r' === sType)
			{
				this.Result.SetR(true);
			}
			else if('t' === sType)
			{
				this.Result.SetT(true)
			}
			else if('w' === sType)
			{
				this.Result.SetW(true)
			}
			else if ('s' === sType)
			{
				this.Result.SetS(true);
			}
		}
	}
};
CFieldInstructionParser.prototype.private_ReadTIME = function()
{
	this.Result = new CFieldInstructionTIME();

	while (this.private_ReadNext())
	{
		if (this.private_IsSwitch())
		{
			if ('@' === this.Buffer.charAt(1))
			{
				var arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
					this.Result.SetFormat(arrArguments[0]);
			}
		}
	}
};
CFieldInstructionParser.prototype.private_ReadDATE = function()
{
	this.Result = new CFieldInstructionDATE();

	while (this.private_ReadNext())
	{
		if (this.private_IsSwitch())
		{
			if ('@' === this.Buffer.charAt(1))
			{
				var arrArguments = this.private_ReadArguments();
				if (arrArguments.length > 0)
					this.Result.SetFormat(arrArguments[0]);
			}
		}
	}
};
CFieldInstructionParser.prototype.private_ReadADDIN = function()
{
	this.Result = new AscWord.CFieldInstructionADDIN();
	
	if (this.private_ReadTillEnd())
		this.Result.SetValue(this.Buffer);
};

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

AscDFH.changesFactory[AscDFH.historyitem_Paragraph_AddItem]                   = CChangesParagraphAddItem;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_RemoveItem]                = CChangesParagraphRemoveItem;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Numbering]                 = CChangesParagraphNumbering;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Align]                     = CChangesParagraphAlign;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Ind_First]                 = CChangesParagraphIndFirst;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Ind_Right]                 = CChangesParagraphIndRight;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Ind_Left]                  = CChangesParagraphIndLeft;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_ContextualSpacing]         = CChangesParagraphContextualSpacing;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_KeepLines]                 = CChangesParagraphKeepLines;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_KeepNext]                  = CChangesParagraphKeepNext;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_PageBreakBefore]           = CChangesParagraphPageBreakBefore;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Spacing_Line]              = CChangesParagraphSpacingLine;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Spacing_LineRule]          = CChangesParagraphSpacingLineRule;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Spacing_Before]            = CChangesParagraphSpacingBefore;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Spacing_After]             = CChangesParagraphSpacingAfter;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Spacing_AfterAutoSpacing]  = CChangesParagraphSpacingAfterAutoSpacing;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Spacing_BeforeAutoSpacing] = CChangesParagraphSpacingBeforeAutoSpacing;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Shd_Value]                 = CChangesParagraphShdValue;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Shd_Color]                 = CChangesParagraphShdColor;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Shd_Unifill]               = CChangesParagraphShdUnifill;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Shd]                       = CChangesParagraphShd;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_WidowControl]              = CChangesParagraphWidowControl;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Tabs]                      = CChangesParagraphTabs;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_PStyle]                    = CChangesParagraphPStyle;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Borders_Between]           = CChangesParagraphBordersBetween;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Borders_Bottom]            = CChangesParagraphBordersBottom;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Borders_Left]              = CChangesParagraphBordersLeft;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Borders_Right]             = CChangesParagraphBordersRight;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Borders_Top]               = CChangesParagraphBordersTop;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Pr]                        = CChangesParagraphPr;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_PresentationPr_Bullet]     = CChangesParagraphPresentationPrBullet;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_PresentationPr_Level]      = CChangesParagraphPresentationPrLevel;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_FramePr]                   = CChangesParagraphFramePr;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_SectionPr]                 = CChangesParagraphSectPr;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_PrChange]                  = CChangesParagraphPrChange;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_PrReviewInfo]              = CChangesParagraphPrReviewInfo;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_OutlineLvl]                = CChangesParagraphOutlineLvl;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_DefaultTabSize]            = CChangesParagraphDefaultTabSize;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_SuppressLineNumbers]       = CChangesParagraphSuppressLineNumbers;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Shd_Fill]                  = CChangesParagraphShdFill;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Shd_ThemeFill]             = CChangesParagraphShdThemeFill;
AscDFH.changesFactory[AscDFH.historyitem_Paragraph_Bidi]                      = CChangesParagraphBidi;

function private_ParagraphChangesOnLoadPr(oColor)
{
	this.Redo();

	if (oColor)
		this.Class.private_AddCollPrChange(oColor);
}

function private_ParagraphChangesOnSetValue(oParagraph)
{
	oParagraph.RecalcInfo.Set_Type_0(pararecalc_0_All);
	oParagraph.OnContentChange();
}

//----------------------------------------------------------------------------------------------------------------------
// Карта зависимости изменений
//----------------------------------------------------------------------------------------------------------------------
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_AddItem]                   = [
	AscDFH.historyitem_Paragraph_AddItem,
	AscDFH.historyitem_Paragraph_RemoveItem
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_RemoveItem]                = [
	AscDFH.historyitem_Paragraph_AddItem,
	AscDFH.historyitem_Paragraph_RemoveItem
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Numbering]                 = [
	AscDFH.historyitem_Paragraph_Numbering,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Align]                     = [
	AscDFH.historyitem_Paragraph_Align,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_DefaultTabSize]            = [
	AscDFH.historyitem_Paragraph_DefaultTabSize,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Ind_First]                 = [
	AscDFH.historyitem_Paragraph_Ind_First,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Ind_Right]                 = [
	AscDFH.historyitem_Paragraph_Ind_Right,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Ind_Left]                  = [
	AscDFH.historyitem_Paragraph_Ind_Left,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_ContextualSpacing]         = [
	AscDFH.historyitem_Paragraph_ContextualSpacing,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_KeepLines]                 = [
	AscDFH.historyitem_Paragraph_KeepLines,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_KeepNext]                  = [
	AscDFH.historyitem_Paragraph_KeepNext,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_PageBreakBefore]           = [
	AscDFH.historyitem_Paragraph_PageBreakBefore,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Spacing_Line]              = [
	AscDFH.historyitem_Paragraph_Spacing_Line,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Spacing_LineRule]          = [
	AscDFH.historyitem_Paragraph_Spacing_LineRule,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Spacing_Before]            = [
	AscDFH.historyitem_Paragraph_Spacing_Before,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Spacing_After]             = [
	AscDFH.historyitem_Paragraph_Spacing_After,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Spacing_AfterAutoSpacing]  = [
	AscDFH.historyitem_Paragraph_Spacing_AfterAutoSpacing,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Spacing_BeforeAutoSpacing] = [
	AscDFH.historyitem_Paragraph_Spacing_BeforeAutoSpacing,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Shd_Value]                 = [
	AscDFH.historyitem_Paragraph_Shd_Value,
	AscDFH.historyitem_Paragraph_Shd,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Shd_Color]                 = [
	AscDFH.historyitem_Paragraph_Shd_Color,
	AscDFH.historyitem_Paragraph_Shd,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Shd_Unifill]               = [
	AscDFH.historyitem_Paragraph_Shd_Unifill,
	AscDFH.historyitem_Paragraph_Shd,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Shd]                       = [
	AscDFH.historyitem_Paragraph_Shd_Value,
	AscDFH.historyitem_Paragraph_Shd_Color,
	AscDFH.historyitem_Paragraph_Shd_Unifill,
	AscDFH.historyitem_Paragraph_Shd_Fill,
	AscDFH.historyitem_Paragraph_Shd_ThemeFill,
	AscDFH.historyitem_Paragraph_Shd,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_WidowControl]              = [
	AscDFH.historyitem_Paragraph_WidowControl,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Tabs]                      = [
	AscDFH.historyitem_Paragraph_Tabs,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_PStyle]                    = [
	AscDFH.historyitem_Paragraph_PStyle,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Borders_Between]           = [
	AscDFH.historyitem_Paragraph_Borders_Between,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Borders_Bottom]            = [
	AscDFH.historyitem_Paragraph_Borders_Bottom,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Borders_Left]              = [
	AscDFH.historyitem_Paragraph_Borders_Left,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Borders_Right]             = [
	AscDFH.historyitem_Paragraph_Borders_Right,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Borders_Top]               = [
	AscDFH.historyitem_Paragraph_Borders_Top,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Pr]                        = [
	AscDFH.historyitem_Paragraph_Pr,
	AscDFH.historyitem_Paragraph_Numbering,
	AscDFH.historyitem_Paragraph_Align,
	AscDFH.historyitem_Paragraph_DefaultTabSize,
	AscDFH.historyitem_Paragraph_Ind_First,
	AscDFH.historyitem_Paragraph_Ind_Right,
	AscDFH.historyitem_Paragraph_Ind_Left,
	AscDFH.historyitem_Paragraph_ContextualSpacing,
	AscDFH.historyitem_Paragraph_KeepLines,
	AscDFH.historyitem_Paragraph_KeepNext,
	AscDFH.historyitem_Paragraph_PageBreakBefore,
	AscDFH.historyitem_Paragraph_Spacing_Line,
	AscDFH.historyitem_Paragraph_Spacing_LineRule,
	AscDFH.historyitem_Paragraph_Spacing_Before,
	AscDFH.historyitem_Paragraph_Spacing_After,
	AscDFH.historyitem_Paragraph_Spacing_AfterAutoSpacing,
	AscDFH.historyitem_Paragraph_Spacing_BeforeAutoSpacing,
	AscDFH.historyitem_Paragraph_Shd_Value,
	AscDFH.historyitem_Paragraph_Shd_Color,
	AscDFH.historyitem_Paragraph_Shd_Unifill,
	AscDFH.historyitem_Paragraph_Shd_Fill,
	AscDFH.historyitem_Paragraph_Shd_ThemeFill,
	AscDFH.historyitem_Paragraph_Shd,
	AscDFH.historyitem_Paragraph_WidowControl,
	AscDFH.historyitem_Paragraph_Tabs,
	AscDFH.historyitem_Paragraph_PStyle,
	AscDFH.historyitem_Paragraph_Borders_Between,
	AscDFH.historyitem_Paragraph_Borders_Bottom,
	AscDFH.historyitem_Paragraph_Borders_Left,
	AscDFH.historyitem_Paragraph_Borders_Right,
	AscDFH.historyitem_Paragraph_Borders_Top,
	AscDFH.historyitem_Paragraph_PresentationPr_Bullet,
	AscDFH.historyitem_Paragraph_PresentationPr_Level,
	AscDFH.historyitem_Paragraph_FramePr,
	AscDFH.historyitem_Paragraph_PrChange,
	AscDFH.historyitem_Paragraph_PrReviewInfo,
	AscDFH.historyitem_Paragraph_OutlineLvl,
	AscDFH.historyitem_Paragraph_SuppressLineNumbers,
	AscDFH.historyitem_Paragraph_Bidi
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_PresentationPr_Bullet]     = [
	AscDFH.historyitem_Paragraph_PresentationPr_Bullet,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_PresentationPr_Level]      = [
	AscDFH.historyitem_Paragraph_PresentationPr_Level,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_FramePr]                   = [
	AscDFH.historyitem_Paragraph_FramePr,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_SectionPr]                 = [
	AscDFH.historyitem_Paragraph_SectionPr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_PrChange]                  = [
	AscDFH.historyitem_Paragraph_Pr,
	AscDFH.historyitem_Paragraph_PrChange,
	AscDFH.historyitem_Paragraph_PrReviewInfo
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_PrReviewInfo]              = [
	AscDFH.historyitem_Paragraph_Pr,
	AscDFH.historyitem_Paragraph_PrChange,
	AscDFH.historyitem_Paragraph_PrReviewInfo
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_OutlineLvl]                = [
	AscDFH.historyitem_Paragraph_OutlineLvl
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_SuppressLineNumbers]       = [
	AscDFH.historyitem_Paragraph_SuppressLineNumbers,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Shd_Fill]                  = [
	AscDFH.historyitem_Paragraph_Shd_Fill,
	AscDFH.historyitem_Paragraph_Shd,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Shd_ThemeFill]               = [
	AscDFH.historyitem_Paragraph_Shd_ThemeFill,
	AscDFH.historyitem_Paragraph_Shd,
	AscDFH.historyitem_Paragraph_Pr
];
AscDFH.changesRelationMap[AscDFH.historyitem_Paragraph_Bidi] = [
	AscDFH.historyitem_Paragraph_Bidi,
	AscDFH.historyitem_Paragraph_Pr
];

// Общая функция Merge для изменений, которые зависят только от себя и AscDFH.historyitem_Paragraph_Pr
function private_ParagraphChangesOnMergePr(oChange)
{
	if (oChange.Class !== this.Class)
		return true;

	if (oChange.Type === this.Type || oChange.Type === AscDFH.historyitem_Paragraph_Pr)
		return false;

	return true;
}
// Общая функция Merge для изменений, которые зависят от себя, AscDFH.historyitem_Paragraph_Shd, AscDFH.historyitem_Paragraph_Pr
function private_ParagraphChangesOnMergeShdPr(oChange)
{
	if (oChange.Class !== this.Class)
		return true;

	if (oChange.Type === this.Type || oChange.Type === AscDFH.historyitem_Paragraph_Pr || oChange.Type === AscDFH.historyitem_Paragraph_Shd)
		return false;

	return true;
}
//----------------------------------------------------------------------------------------------------------------------

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesParagraphAddItem(Class, Pos, Items)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, true);
}
CChangesParagraphAddItem.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesParagraphAddItem.prototype.constructor = CChangesParagraphAddItem;
CChangesParagraphAddItem.prototype.Type = AscDFH.historyitem_Paragraph_AddItem;
CChangesParagraphAddItem.prototype.Undo = function()
{
	var oParagraph = this.Class;
	oParagraph.Content.splice(this.Pos, this.Items.length);
	oParagraph.updateTrackRevisions();
	oParagraph.private_CheckUpdateBookmarks(this.Items);
	oParagraph.private_UpdateSelectionPosOnRemove(this.Pos, this.Items.length);
	private_ParagraphChangesOnSetValue(this.Class);
};
CChangesParagraphAddItem.prototype.Redo = function()
{
	var oParagraph  = this.Class;
	var Array_start = oParagraph.Content.slice(0, this.Pos);
	var Array_end   = oParagraph.Content.slice(this.Pos);

	oParagraph.Content = Array_start.concat(this.Items, Array_end);
	oParagraph.updateTrackRevisions();
	oParagraph.private_CheckUpdateBookmarks(this.Items);
	oParagraph.private_UpdateSelectionPosOnAdd(this.Pos, this.Items.length);
	private_ParagraphChangesOnSetValue(this.Class);

	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		var oItem = this.Items[nIndex];
		oItem.Parent = this.Class;

		if (oItem.SetParent)
			oItem.SetParent(this.Class);

		if (oItem.SetParagraph)
			oItem.SetParagraph(this.Class);

		if (oItem.Recalc_RunsCompiledPr)
			oItem.Recalc_RunsCompiledPr();
	}
};
CChangesParagraphAddItem.prototype.private_WriteItem = function(Writer, Item)
{
	Writer.WriteString2(Item.Get_Id());
};
CChangesParagraphAddItem.prototype.private_ReadItem = function(Reader)
{
	return AscCommon.g_oTableId.Get_ById(Reader.GetString2());
};
CChangesParagraphAddItem.prototype.Load = function(Color)
{
	var oParagraph = this.Class;
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		var Pos     = oParagraph.m_oContentChanges.Check(AscCommon.contentchanges_Add, this.PosArray[nIndex]);
		var Element = this.Items[nIndex];

		if (null != Element)
		{
			if (para_Comment === Element.Type)
			{
				var oComment = AscCommon.g_oTableId.Get_ById(Element.CommentId);
				if (oComment instanceof AscCommon.CComment)
					oComment.UpdatePosition();
			}

			if (Element.SetParagraph)
				Element.SetParagraph(oParagraph);

			if (Element.SetParent)
				Element.SetParent(oParagraph);

			oParagraph.Content.splice(Pos, 0, Element);
			oParagraph.private_UpdateSelectionPosOnAdd(Pos, 1);
			AscCommon.CollaborativeEditing.Update_DocumentPositionsOnAdd(oParagraph, Pos);

			if (Element.Recalc_RunsCompiledPr)
				Element.Recalc_RunsCompiledPr();
		}
	}

	oParagraph.private_ResetSelection();
	oParagraph.updateTrackRevisions();
	oParagraph.private_CheckUpdateBookmarks(this.Items);
	oParagraph.UpdateDocumentOutline();

	private_ParagraphChangesOnSetValue(this.Class);
};
CChangesParagraphAddItem.prototype.IsRelated = function(oChanges)
{
	if (this.Class === oChanges.Class && (AscDFH.historyitem_Paragraph_AddItem === oChanges.Type || AscDFH.historyitem_Paragraph_RemoveItem === oChanges.Type))
		return true;

	return false;
};
CChangesParagraphAddItem.prototype.CreateReverseChange = function()
{
	return this.private_CreateReverseChange(CChangesParagraphRemoveItem);
};
CChangesParagraphAddItem.prototype.IsParagraphSimpleChanges = function()
{
	// Простыми измененями считаем добавление комментариев и добавление ранов с простым текстом
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		var oItem = this.Items[nIndex];
		if ((para_Run !== oItem.Type || !oItem.IsContentSuitableForParagraphSimpleChanges())
			&& para_Comment !== oItem.Type
			&& para_Bookmark !== oItem.Type)
		{
			return false;
		}
	}

	return true;
};
CChangesParagraphAddItem.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesParagraphRemoveItem(Class, Pos, Items)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, false);
}
CChangesParagraphRemoveItem.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesParagraphRemoveItem.prototype.constructor = CChangesParagraphRemoveItem;
CChangesParagraphRemoveItem.prototype.Type = AscDFH.historyitem_Paragraph_RemoveItem;
CChangesParagraphRemoveItem.prototype.Undo = function()
{
	var oParagraph  = this.Class;
	var Array_start = oParagraph.Content.slice(0, this.Pos);
	var Array_end   = oParagraph.Content.slice(this.Pos);

	oParagraph.Content = Array_start.concat(this.Items, Array_end);
	oParagraph.updateTrackRevisions();
	oParagraph.private_CheckUpdateBookmarks(this.Items);
	oParagraph.private_UpdateSelectionPosOnAdd(this.Pos, this.Items.length);
	private_ParagraphChangesOnSetValue(this.Class);

	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		var oItem = this.Items[nIndex];

		oItem.Parent = this.Class;

		if (oItem.SetParent)
			oItem.SetParent(this.Class);

		if (oItem.SetParagraph)
			oItem.SetParagraph(this.Class);

		if (oItem.Recalc_RunsCompiledPr)
			oItem.Recalc_RunsCompiledPr();
	}
};
CChangesParagraphRemoveItem.prototype.Redo = function()
{
	var oParagraph  = this.Class;
	oParagraph.Content.splice(this.Pos, this.Items.length);
	oParagraph.updateTrackRevisions();
	oParagraph.private_CheckUpdateBookmarks(this.Items);
	oParagraph.private_UpdateSelectionPosOnRemove(this.Pos, this.Items.length);
	private_ParagraphChangesOnSetValue(this.Class);
};
CChangesParagraphRemoveItem.prototype.private_WriteItem = function(Writer, Item)
{
	Writer.WriteString2(Item.Get_Id());
};
CChangesParagraphRemoveItem.prototype.private_ReadItem = function(Reader)
{
	return AscCommon.g_oTableId.Get_ById(Reader.GetString2());
};
CChangesParagraphRemoveItem.prototype.Load = function(Color)
{
	var oParagraph = this.Class;
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		var ChangesPos = oParagraph.m_oContentChanges.Check(AscCommon.contentchanges_Remove, this.PosArray[nIndex]);

		if (false === ChangesPos)
			continue;

		oParagraph.Content.splice(ChangesPos, 1);
		oParagraph.private_UpdateSelectionPosOnRemove(ChangesPos, 1);
		AscCommon.CollaborativeEditing.Update_DocumentPositionsOnRemove(oParagraph, ChangesPos, 1);
	}
	oParagraph.private_ResetSelection();
	oParagraph.updateTrackRevisions();
	oParagraph.private_CheckUpdateBookmarks(this.Items);
	oParagraph.UpdateDocumentOutline();
	private_ParagraphChangesOnSetValue(this.Class);
};
CChangesParagraphRemoveItem.prototype.IsRelated = function(oChanges)
{
	if (this.Class === oChanges.Class && (AscDFH.historyitem_Paragraph_AddItem === oChanges.Type || AscDFH.historyitem_Paragraph_RemoveItem === oChanges.Type))
		return true;

	return false;
};
CChangesParagraphRemoveItem.prototype.CreateReverseChange = function()
{
	return this.private_CreateReverseChange(CChangesParagraphAddItem);
};
CChangesParagraphRemoveItem.prototype.IsParagraphSimpleChanges = function()
{
	// Простыми измененями считаем добавление комментариев и добавление ранов с простым текстом
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		var oItem = this.Items[nIndex];
		if ((para_Run !== oItem.Type || !oItem.IsContentSuitableForParagraphSimpleChanges())
			&& para_Comment !== oItem.Type
			&& para_Bookmark !== oItem.Type)
		{
			return false;
		}
	}

	return true;
};
CChangesParagraphRemoveItem.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphNumbering(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphNumbering.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphNumbering.prototype.constructor = CChangesParagraphNumbering;
CChangesParagraphNumbering.prototype.Type = AscDFH.historyitem_Paragraph_Numbering;
CChangesParagraphNumbering.prototype.private_CreateObject = function()
{
	return new AscWord.NumPr();
};
CChangesParagraphNumbering.prototype.private_SetValue = function(newNumPr)
{
	var oParagraph = this.Class;
	
	let oldNumPr   = oParagraph.Pr.NumPr;
	oParagraph.Pr.NumPr = newNumPr;
	
	oParagraph.private_RefreshNumbering(oldNumPr);
	oParagraph.private_RefreshNumbering(newNumPr);
	
	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphNumbering.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphNumbering.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphNumbering.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesParagraphAlign(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphAlign.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesParagraphAlign.prototype.constructor = CChangesParagraphAlign;
CChangesParagraphAlign.prototype.Type = AscDFH.historyitem_Paragraph_Align;
CChangesParagraphAlign.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.Jc = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphAlign.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphAlign.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphAlign.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseDoubleProperty}
 */
function CChangesParagraphIndFirst(Class, Old, New, Color)
{
	AscDFH.CChangesBaseDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphIndFirst.prototype = Object.create(AscDFH.CChangesBaseDoubleProperty.prototype);
CChangesParagraphIndFirst.prototype.constructor = CChangesParagraphIndFirst;
CChangesParagraphIndFirst.prototype.Type = AscDFH.historyitem_Paragraph_Ind_First;
CChangesParagraphIndFirst.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Ind)
		oParagraph.Pr.Ind = new CParaInd();

	oParagraph.Pr.Ind.FirstLine = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphIndFirst.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphIndFirst.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphIndFirst.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseDoubleProperty}
 */
function CChangesParagraphDefaultTabSize(Class, Old, New, Color)
{
	AscDFH.CChangesBaseDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphDefaultTabSize.prototype = Object.create(AscDFH.CChangesBaseDoubleProperty.prototype);
CChangesParagraphDefaultTabSize.prototype.constructor = CChangesParagraphDefaultTabSize;
CChangesParagraphDefaultTabSize.prototype.Type = AscDFH.historyitem_Paragraph_DefaultTabSize;
CChangesParagraphDefaultTabSize.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;


	oParagraph.Pr.DefaultTab = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphDefaultTabSize.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphDefaultTabSize.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphDefaultTabSize.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseDoubleProperty}
 */
function CChangesParagraphIndLeft(Class, Old, New, Color)
{
	AscDFH.CChangesBaseDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphIndLeft.prototype = Object.create(AscDFH.CChangesBaseDoubleProperty.prototype);
CChangesParagraphIndLeft.prototype.constructor = CChangesParagraphIndLeft;
CChangesParagraphIndLeft.prototype.Type = AscDFH.historyitem_Paragraph_Ind_Left;
CChangesParagraphIndLeft.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Ind)
		oParagraph.Pr.Ind = new CParaInd();

	oParagraph.Pr.Ind.Left = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphIndLeft.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphIndLeft.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphIndLeft.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseDoubleProperty}
 */
function CChangesParagraphIndRight(Class, Old, New, Color)
{
	AscDFH.CChangesBaseDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphIndRight.prototype = Object.create(AscDFH.CChangesBaseDoubleProperty.prototype);
CChangesParagraphIndRight.prototype.constructor = CChangesParagraphIndRight;
CChangesParagraphIndRight.prototype.Type = AscDFH.historyitem_Paragraph_Ind_Right;
CChangesParagraphIndRight.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Ind)
		oParagraph.Pr.Ind = new CParaInd();

	oParagraph.Pr.Ind.Right = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphIndRight.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphIndRight.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphIndRight.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesParagraphContextualSpacing(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphContextualSpacing.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesParagraphContextualSpacing.prototype.constructor = CChangesParagraphContextualSpacing;
CChangesParagraphContextualSpacing.prototype.Type = AscDFH.historyitem_Paragraph_ContextualSpacing;
CChangesParagraphContextualSpacing.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.ContextualSpacing = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphContextualSpacing.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphContextualSpacing.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphContextualSpacing.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesParagraphKeepLines(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphKeepLines.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesParagraphKeepLines.prototype.constructor = CChangesParagraphKeepLines;
CChangesParagraphKeepLines.prototype.Type = AscDFH.historyitem_Paragraph_KeepLines;
CChangesParagraphKeepLines.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.KeepLines = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphKeepLines.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphKeepLines.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphKeepLines.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesParagraphKeepNext(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphKeepNext.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesParagraphKeepNext.prototype.constructor = CChangesParagraphKeepNext;
CChangesParagraphKeepNext.prototype.Type = AscDFH.historyitem_Paragraph_KeepNext;
CChangesParagraphKeepNext.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.KeepNext = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphKeepNext.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphKeepNext.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphKeepNext.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesParagraphPageBreakBefore(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphPageBreakBefore.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesParagraphPageBreakBefore.prototype.constructor = CChangesParagraphPageBreakBefore;
CChangesParagraphPageBreakBefore.prototype.Type = AscDFH.historyitem_Paragraph_PageBreakBefore;
CChangesParagraphPageBreakBefore.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.PageBreakBefore = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphPageBreakBefore.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphPageBreakBefore.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphPageBreakBefore.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseDoubleProperty}
 */
function CChangesParagraphSpacingLine(Class, Old, New, Color)
{
	AscDFH.CChangesBaseDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphSpacingLine.prototype = Object.create(AscDFH.CChangesBaseDoubleProperty.prototype);
CChangesParagraphSpacingLine.prototype.constructor = CChangesParagraphSpacingLine;
CChangesParagraphSpacingLine.prototype.Type = AscDFH.historyitem_Paragraph_Spacing_Line;
CChangesParagraphSpacingLine.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Spacing)
		oParagraph.Pr.Spacing = new CParaSpacing();

	oParagraph.Pr.Spacing.Line = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphSpacingLine.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphSpacingLine.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphSpacingLine.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesParagraphSpacingLineRule(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphSpacingLineRule.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesParagraphSpacingLineRule.prototype.constructor = CChangesParagraphSpacingLineRule;
CChangesParagraphSpacingLineRule.prototype.Type = AscDFH.historyitem_Paragraph_Spacing_LineRule;
CChangesParagraphSpacingLineRule.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Spacing)
		oParagraph.Pr.Spacing = new CParaSpacing();

	oParagraph.Pr.Spacing.LineRule = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphSpacingLineRule.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphSpacingLineRule.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphSpacingLineRule.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseDoubleProperty}
 */
function CChangesParagraphSpacingBefore(Class, Old, New, Color)
{
	AscDFH.CChangesBaseDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphSpacingBefore.prototype = Object.create(AscDFH.CChangesBaseDoubleProperty.prototype);
CChangesParagraphSpacingBefore.prototype.constructor = CChangesParagraphSpacingBefore;
CChangesParagraphSpacingBefore.prototype.Type = AscDFH.historyitem_Paragraph_Spacing_Before;
CChangesParagraphSpacingBefore.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Spacing)
		oParagraph.Pr.Spacing = new CParaSpacing();

	oParagraph.Pr.Spacing.Before = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphSpacingBefore.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphSpacingBefore.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphSpacingBefore.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseDoubleProperty}
 */
function CChangesParagraphSpacingAfter(Class, Old, New, Color)
{
	AscDFH.CChangesBaseDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphSpacingAfter.prototype = Object.create(AscDFH.CChangesBaseDoubleProperty.prototype);
CChangesParagraphSpacingAfter.prototype.constructor = CChangesParagraphSpacingAfter;
CChangesParagraphSpacingAfter.prototype.Type = AscDFH.historyitem_Paragraph_Spacing_After;
CChangesParagraphSpacingAfter.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Spacing)
		oParagraph.Pr.Spacing = new CParaSpacing();

	oParagraph.Pr.Spacing.After = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphSpacingAfter.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphSpacingAfter.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphSpacingAfter.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesParagraphSpacingAfterAutoSpacing(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphSpacingAfterAutoSpacing.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesParagraphSpacingAfterAutoSpacing.prototype.constructor = CChangesParagraphSpacingAfterAutoSpacing;
CChangesParagraphSpacingAfterAutoSpacing.prototype.Type = AscDFH.historyitem_Paragraph_Spacing_AfterAutoSpacing;
CChangesParagraphSpacingAfterAutoSpacing.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Spacing)
		oParagraph.Pr.Spacing = new CParaSpacing();

	oParagraph.Pr.Spacing.AfterAutoSpacing = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphSpacingAfterAutoSpacing.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphSpacingAfterAutoSpacing.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphSpacingAfterAutoSpacing.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesParagraphSpacingBeforeAutoSpacing(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphSpacingBeforeAutoSpacing.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesParagraphSpacingBeforeAutoSpacing.prototype.constructor = CChangesParagraphSpacingBeforeAutoSpacing;
CChangesParagraphSpacingBeforeAutoSpacing.prototype.Type = AscDFH.historyitem_Paragraph_Spacing_BeforeAutoSpacing;
CChangesParagraphSpacingBeforeAutoSpacing.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Spacing)
		oParagraph.Pr.Spacing = new CParaSpacing();

	oParagraph.Pr.Spacing.BeforeAutoSpacing = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphSpacingBeforeAutoSpacing.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphSpacingBeforeAutoSpacing.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphSpacingBeforeAutoSpacing.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseByteProperty}
 */
function CChangesParagraphShdValue(Class, Old, New, Color)
{
	AscDFH.CChangesBaseByteProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphShdValue.prototype = Object.create(AscDFH.CChangesBaseByteProperty.prototype);
CChangesParagraphShdValue.prototype.constructor = CChangesParagraphShdValue;
CChangesParagraphShdValue.prototype.Type = AscDFH.historyitem_Paragraph_Shd_Value;
CChangesParagraphShdValue.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Shd)
		oParagraph.Pr.Shd = new CDocumentShd();

	oParagraph.Pr.Shd.Value = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphShdValue.prototype.Merge = private_ParagraphChangesOnMergeShdPr;
CChangesParagraphShdValue.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphShdValue.prototype.IsNeedRecalculate = function()
{
	return false;
};
CChangesParagraphShdValue.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphShdColor(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphShdColor.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphShdColor.prototype.constructor = CChangesParagraphShdColor;
CChangesParagraphShdColor.prototype.Type = AscDFH.historyitem_Paragraph_Shd_Color;
CChangesParagraphShdColor.prototype.private_CreateObject = function()
{
	return new CDocumentColor(0, 0, 0);
};
CChangesParagraphShdColor.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Shd)
		oParagraph.Pr.Shd = new CDocumentShd();

	oParagraph.Pr.Shd.Color = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphShdColor.prototype.Merge = private_ParagraphChangesOnMergeShdPr;
CChangesParagraphShdColor.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphShdColor.prototype.IsNeedRecalculate = function()
{
	return false;
};
CChangesParagraphShdColor.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphShdUnifill(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphShdUnifill.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphShdUnifill.prototype.constructor = CChangesParagraphShdUnifill;
CChangesParagraphShdUnifill.prototype.Type = AscDFH.historyitem_Paragraph_Shd_Unifill;
CChangesParagraphShdUnifill.prototype.private_CreateObject = function()
{
	return new AscFormat.CUniFill();
};
CChangesParagraphShdUnifill.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Shd)
		oParagraph.Pr.Shd = new CDocumentShd();

	oParagraph.Pr.Shd.Unifill = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphShdUnifill.prototype.Merge = private_ParagraphChangesOnMergeShdPr;
CChangesParagraphShdUnifill.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphShdUnifill.prototype.IsNeedRecalculate = function()
{
	return false;
};
CChangesParagraphShdUnifill.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphShd(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphShd.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphShd.prototype.constructor = CChangesParagraphShd;
CChangesParagraphShd.prototype.Type = AscDFH.historyitem_Paragraph_Shd;
CChangesParagraphShd.prototype.private_CreateObject = function()
{
	return new CDocumentShd();
};
CChangesParagraphShd.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.Shd = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphShd.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type || oChange.Type === AscDFH.historyitem_Paragraph_Pr)
		return false;

	if (AscDFH.historyitem_Paragraph_Shd_Value === oChange.Type)
	{
		if (!this.New)
			this.New = new CDocumentShd();

		this.New.Value = oChange.New;
	}
	else if (AscDFH.historyitem_Paragraph_Shd_Color === oChange.Type)
	{
		if (!this.New)
			this.New = new CDocumentShd();

		this.New.Color = oChange.New;
	}
	else if (AscDFH.historyitem_Paragraph_Shd_Unifill === oChange.Type)
	{
		if (!this.New)
			this.New = new CDocumentShd();

		this.New.Unifill = oChange.New;
	}

	return true;
};
CChangesParagraphShd.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphShd.prototype.IsNeedRecalculate = function()
{
	return false;
};
CChangesParagraphShd.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesParagraphWidowControl(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphWidowControl.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesParagraphWidowControl.prototype.constructor = CChangesParagraphWidowControl;
CChangesParagraphWidowControl.prototype.Type = AscDFH.historyitem_Paragraph_WidowControl;
CChangesParagraphWidowControl.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.WidowControl = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphWidowControl.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphWidowControl.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphWidowControl.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphTabs(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphTabs.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphTabs.prototype.constructor = CChangesParagraphTabs;
CChangesParagraphTabs.prototype.Type = AscDFH.historyitem_Paragraph_Tabs;
CChangesParagraphTabs.prototype.private_CreateObject = function()
{
	return new CParaTabs();
};
CChangesParagraphTabs.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.Tabs = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphTabs.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphTabs.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphTabs.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesParagraphPStyle(Class, Old, New, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphPStyle.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesParagraphPStyle.prototype.constructor = CChangesParagraphPStyle;
CChangesParagraphPStyle.prototype.Type = AscDFH.historyitem_Paragraph_PStyle;
CChangesParagraphPStyle.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.PStyle = Value;

	oParagraph.RecalcCompiledPr(true);
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
	oParagraph.Recalc_RunsCompiledPr();
	oParagraph.UpdateDocumentOutline();
	oParagraph.private_RefreshNumbering();
	private_ParagraphChangesOnSetValue(this.Class);
};
CChangesParagraphPStyle.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphPStyle.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphPStyle.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphBordersBetween(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphBordersBetween.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphBordersBetween.prototype.constructor = CChangesParagraphBordersBetween;
CChangesParagraphBordersBetween.prototype.Type = AscDFH.historyitem_Paragraph_Borders_Between;
CChangesParagraphBordersBetween.prototype.private_CreateObject = function()
{
	return new CDocumentBorder();
};
CChangesParagraphBordersBetween.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.Brd.Between = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphBordersBetween.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphBordersBetween.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphBordersBetween.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphBordersBottom(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphBordersBottom.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphBordersBottom.prototype.constructor = CChangesParagraphBordersBottom;
CChangesParagraphBordersBottom.prototype.Type = AscDFH.historyitem_Paragraph_Borders_Bottom;
CChangesParagraphBordersBottom.prototype.private_CreateObject = function()
{
	return new CDocumentBorder();
};
CChangesParagraphBordersBottom.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.Brd.Bottom = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphBordersBottom.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphBordersBottom.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphBordersBottom.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphBordersLeft(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphBordersLeft.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphBordersLeft.prototype.constructor = CChangesParagraphBordersLeft;
CChangesParagraphBordersLeft.prototype.Type = AscDFH.historyitem_Paragraph_Borders_Left;
CChangesParagraphBordersLeft.prototype.private_CreateObject = function()
{
	return new CDocumentBorder();
};
CChangesParagraphBordersLeft.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.Brd.Left = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphBordersLeft.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphBordersLeft.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphBordersLeft.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphBordersRight(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphBordersRight.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphBordersRight.prototype.constructor = CChangesParagraphBordersRight;
CChangesParagraphBordersRight.prototype.Type = AscDFH.historyitem_Paragraph_Borders_Right;
CChangesParagraphBordersRight.prototype.private_CreateObject = function()
{
	return new CDocumentBorder();
};
CChangesParagraphBordersRight.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.Brd.Right = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphBordersRight.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphBordersRight.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphBordersRight.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphBordersTop(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphBordersTop.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphBordersTop.prototype.constructor = CChangesParagraphBordersTop;
CChangesParagraphBordersTop.prototype.Type = AscDFH.historyitem_Paragraph_Borders_Top;
CChangesParagraphBordersTop.prototype.private_CreateObject = function()
{
	return new CDocumentBorder();
};
CChangesParagraphBordersTop.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.Brd.Top = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphBordersTop.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphBordersTop.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphBordersTop.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphPr(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphPr.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphPr.prototype.constructor = CChangesParagraphPr;
CChangesParagraphPr.prototype.Type = AscDFH.historyitem_Paragraph_Pr;
CChangesParagraphPr.prototype.private_CreateObject = function()
{
	return new CParaPr();
};
CChangesParagraphPr.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	var oNumPr     = oParagraph.Pr.NumPr;
	oParagraph.Pr  = Value;

	oParagraph.RecalcCompiledPr(true);
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
	oParagraph.UpdateDocumentOutline();
	private_ParagraphChangesOnSetValue(this.Class);

	if (!oNumPr || !oParagraph.Pr.NumPr || oNumPr.NumId !== oParagraph.Pr.NumPr.NumId || oNumPr.Lvl !== oParagraph.Pr.NumPr.Lvl)
	{
		oParagraph.private_RefreshNumbering(oNumPr);
		oParagraph.private_RefreshNumbering(oParagraph.Pr.NumPr);
	}
};
CChangesParagraphPr.prototype.private_IsCreateEmptyObject = function()
{
	return true;
};
CChangesParagraphPr.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return;

	if (AscDFH.historyitem_Paragraph_Pr === oChange.Type)
		return false;

	if (!this.New)
		this.New = new CParaPr();

	switch (oChange.Type)
	{
		case AscDFH.historyitem_Paragraph_Numbering:
		{
			this.New.NumPr = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Align:
		{
			this.New.Jc = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_DefaultTabSize:
		{
			this.New.DefaultTab = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Ind_First:
		{
			if (!this.New.Ind)
				this.New.Ind = new CParaInd();

			this.New.Ind.FirstLine = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Ind_Right:
		{
			if (!this.New.Ind)
				this.New.Ind = new CParaInd();

			this.New.Ind.Right = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Ind_Left:
		{
			if (!this.New.Ind)
				this.New.Ind = new CParaInd();

			this.New.Ind.Left = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_ContextualSpacing:
		{
			this.New.ContextualSpacing = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_KeepLines:
		{
			this.New.KeepLines = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_KeepNext:
		{
			this.New.KeepNext = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_PageBreakBefore:
		{
			this.New.PageBreakBefore = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Spacing_Line:
		{
			if (!this.New.Spacing)
				this.New.Spacing = new CParaSpacing();

			this.New.Spacing.Line = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Spacing_LineRule:
		{
			if (!this.New.Spacing)
				this.New.Spacing = new CParaSpacing();

			this.New.Spacing.LineRule = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Spacing_Before:
		{
			if (!this.New.Spacing)
				this.New.Spacing = new CParaSpacing();

			this.New.Spacing.Before = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Spacing_After:
		{
			if (!this.New.Spacing)
				this.New.Spacing = new CParaSpacing();

			this.New.Spacing.After = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Spacing_AfterAutoSpacing:
		{
			if (!this.New.Spacing)
				this.New.Spacing = new CParaSpacing();

			this.New.Spacing.AfterAutoSpacing = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Spacing_BeforeAutoSpacing:
		{
			if (!this.New.Spacing)
				this.New.Spacing = new CParaSpacing();

			this.New.Spacing.BeforeAutoSpacing = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Shd_Value:
		{
			if (!this.New.Shd)
				this.New.Shd = new CDocumentShd();

			this.New.Shd.Value = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Shd_Color:
		{
			if (!this.New.Shd)
				this.New.Shd = new CDocumentShd();

			this.New.Shd.Color = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Shd_Unifill:
		{
			if (!this.New.Shd)
				this.New.Shd = new CDocumentShd();

			this.New.Shd.Unifill = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Shd:
		{
			this.New.Shd = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_WidowControl:
		{
			this.New.WidowControl = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Tabs:
		{
			this.New.Tabs = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_PStyle:
		{
			this.New.PStyle = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Borders_Between:
		{
			this.New.Brd.Between = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Borders_Bottom:
		{
			this.New.Brd.Bottom = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Borders_Left:
		{
			this.New.Brd.Left = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Borders_Right:
		{
			this.New.Brd.Right = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_Borders_Top:
		{
			this.New.Brd.Top = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_PresentationPr_Bullet:
		{
			this.New.Bullet = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_PresentationPr_Level:
		{
			this.New.Lvl = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_FramePr:
		{
			this.New.FramePr = oChange.New;
			break;
		}
		case AscDFH.historyitem_Paragraph_PrChange:
		{
			this.New.PrChange   = oChange.New.PrChange;
			this.New.ReviewInfo = oChange.New.ReviewInfo;
			break;
		}
		case AscDFH.historyitem_Paragraph_PrReviewInfo:
		{
			this.New.ReviewInfo = oChange.New;
			break;
		}
	}

	return true;
};
CChangesParagraphPr.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphPr.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphPresentationPrBullet(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphPresentationPrBullet.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphPresentationPrBullet.prototype.constructor = CChangesParagraphPresentationPrBullet;
CChangesParagraphPresentationPrBullet.prototype.Type = AscDFH.historyitem_Paragraph_PresentationPr_Bullet;
CChangesParagraphPresentationPrBullet.prototype.private_CreateObject = function()
{
	return new AscFormat.CBullet();
};
CChangesParagraphPresentationPrBullet.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphPresentationPrBullet.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.Bullet = Value;
	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.Recalc_RunsCompiledPr();
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphPresentationPrBullet.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesParagraphPresentationPrLevel(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphPresentationPrLevel.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesParagraphPresentationPrLevel.prototype.constructor = CChangesParagraphPresentationPrLevel;
CChangesParagraphPresentationPrLevel.prototype.Type = AscDFH.historyitem_Paragraph_PresentationPr_Level;
CChangesParagraphPresentationPrLevel.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphPresentationPrLevel.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.Lvl = Value;
	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.Recalc_RunsCompiledPr();
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphPresentationPrLevel.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphFramePr(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphFramePr.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphFramePr.prototype.constructor = CChangesParagraphFramePr;
CChangesParagraphFramePr.prototype.Type = AscDFH.historyitem_Paragraph_FramePr;
CChangesParagraphFramePr.prototype.private_CreateObject = function()
{
	return new CFramePr();
};
CChangesParagraphFramePr.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.FramePr = Value;
	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphFramePr.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphFramePr.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesParagraphSectPr(Class, Old, New)
{
	AscDFH.CChangesBase.call(this, Class);

	this.Old = Old;
	this.New = New;
}
CChangesParagraphSectPr.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesParagraphSectPr.prototype.constructor = CChangesParagraphSectPr;
CChangesParagraphSectPr.prototype.Type = AscDFH.historyitem_Paragraph_SectionPr;
CChangesParagraphSectPr.prototype.Undo = function()
{
	var oParagraph = this.Class;
	var oOldSectPr = oParagraph.SectPr;
	oParagraph.SectPr = this.Old;
	
	let logicDocument = oParagraph.GetLogicDocument();
	if (logicDocument)
		logicDocument.UpdateSectionInfo(oOldSectPr, this.Old, false);
};
CChangesParagraphSectPr.prototype.Redo = function()
{
	var oParagraph = this.Class;
	var oOldSectPr = oParagraph.SectPr;
	oParagraph.SectPr = this.New;
	
	let logicDocument = oParagraph.GetLogicDocument();
	if (logicDocument)
		logicDocument.UpdateSectionInfo(oOldSectPr, this.New, false);
};
CChangesParagraphSectPr.prototype.WriteToBinary = function(Writer)
{
	// Long  : Flag
	// 1-bit : IsUndefined New
	// 2-bit : IsUndefined Old
	// String : Id of New
	// String : Id of Old

	var nFlags = 0;

	if (undefined === this.New)
		nFlags |= 1;

	if (undefined === this.Old)
		nFlags |= 2;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New)
		Writer.WriteString2(this.New.Get_Id());

	if (undefined !== this.Old)
		Writer.WriteString2(this.Old.Get_Id());
};
CChangesParagraphSectPr.prototype.ReadFromBinary = function(Reader)
{
	// Long  : Flag
	// 1-bit : IsUndefined New
	// 2-bit : IsUndefined Old
	// String : Id of New
	// String : Id of Old

	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.New = undefined;
	else
		this.New = AscCommon.g_oTableId.Get_ById(Reader.GetString2());

	if (nFlags & 2)
		this.Old = undefined;
	else
		this.Old = AscCommon.g_oTableId.Get_ById(Reader.GetString2());
};
CChangesParagraphSectPr.prototype.CreateReverseChange = function()
{
	return new CChangesParagraphSectPr(this.Class, this.New, this.Old);
};
CChangesParagraphSectPr.prototype.Merge = function(oChange)
{
	if (oChange.Class === this.Class && oChange.Type === this.Type)
		return false;

	return true;
};
CChangesParagraphSectPr.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesParagraphPrChange(Class, Old, New)
{
	AscDFH.CChangesBase.call(this, Class);

	this.Old = Old;
	this.New = New;
}
CChangesParagraphPrChange.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesParagraphPrChange.prototype.constructor = CChangesParagraphPrChange;
CChangesParagraphPrChange.prototype.Type = AscDFH.historyitem_Paragraph_PrChange;
CChangesParagraphPrChange.prototype.Undo = function()
{
	var oParagraph = this.Class;
	oParagraph.Pr.PrChange   = this.Old.PrChange;
	oParagraph.Pr.ReviewInfo = this.Old.ReviewInfo;
	oParagraph.updateTrackRevisions();
	private_ParagraphChangesOnSetValue(this.Class);
};
CChangesParagraphPrChange.prototype.Redo = function()
{
	var oParagraph = this.Class;
	oParagraph.Pr.PrChange   = this.New.PrChange;
	oParagraph.Pr.ReviewInfo = this.New.ReviewInfo;
	oParagraph.updateTrackRevisions();
	private_ParagraphChangesOnSetValue(this.Class);
};
CChangesParagraphPrChange.prototype.WriteToBinary = function(Writer)
{
	// Long : Flags
	// 1-bit : is New.PrChange undefined ?
	// 2-bit : is New.ReviewInfo undefined ?
	// 3-bit : is Old.PrChange undefined ?
	// 4-bit : is Old.ReviewInfo undefined ?
	// Variable(CParaPr)     : New.PrChange   (1bit = 0)
	// Variable(AscWord.ReviewInfo) : New.ReviewInfo (2bit = 0)
	// Variable(CParaPr)     : Old.PrChange   (3bit = 0)
	// Variable(AscWord.ReviewInfo) : Old.ReviewInfo (4bit = 0)
	var nFlags = 0;
	if (undefined === this.New.PrChange)
		nFlags |= 1;

	if (undefined === this.New.ReviewInfo)
		nFlags |= 2;

	if (undefined === this.Old.PrChange)
		nFlags |= 4;

	if (undefined === this.Old.ReviewInfo)
		nFlags |= 8;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New.PrChange)
		this.New.PrChange.Write_ToBinary(Writer);

	if (undefined !== this.New.ReviewInfo)
		this.New.ReviewInfo.Write_ToBinary(Writer);

	if (undefined !== this.Old.PrChange)
		this.Old.PrChange.Write_ToBinary(Writer);

	if (undefined !== this.Old.ReviewInfo)
		this.Old.ReviewInfo.Write_ToBinary(Writer);
};
CChangesParagraphPrChange.prototype.ReadFromBinary = function(Reader)
{
	// Long : Flags
	// 1-bit : is New.PrChange undefined ?
	// 2-bit : is New.ReviewInfo undefined ?
	// 3-bit : is Old.PrChange undefined ?
	// 4-bit : is Old.ReviewInfo undefined ?
	// Variable(CParaPr)     : New.PrChange   (1bit = 0)
	// Variable(AscWord.ReviewInfo) : New.ReviewInfo (2bit = 0)
	// Variable(CParaPr)     : Old.PrChange   (3bit = 0)
	// Variable(AscWord.ReviewInfo) : Old.ReviewInfo (4bit = 0)
	var nFlags = Reader.GetLong();

	this.New = {
		PrChange   : undefined,
		ReviewInfo : undefined
	};

	this.Old = {
		PrChange   : undefined,
		ReviewInfo : undefined
	};

	if (nFlags & 1)
	{
		this.New.PrChange = undefined;
	}
	else
	{
		this.New.PrChange = new CParaPr();
		this.New.PrChange.Read_FromBinary(Reader);
	}

	if (nFlags & 2)
	{
		this.New.ReviewInfo = undefined;
	}
	else
	{
		this.New.ReviewInfo = new AscWord.ReviewInfo();
		this.New.ReviewInfo.Read_FromBinary(Reader);
	}

	if (nFlags & 4)
	{
		this.Old.PrChange = undefined;
	}
	else
	{
		this.Old.PrChange = new CParaPr();
		this.Old.PrChange.Read_FromBinary(Reader);
	}

	if (nFlags & 8)
	{
		this.Old.ReviewInfo = undefined;
	}
	else
	{
		this.Old.ReviewInfo = new AscWord.ReviewInfo();
		this.Old.ReviewInfo.Read_FromBinary(Reader);
	}
};
CChangesParagraphPrChange.prototype.CreateReverseChange = function()
{
	return new CChangesParagraphPrChange(this.Class, this.New, this.Old);
};
CChangesParagraphPrChange.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (oChange.Type === this.Type || AscDFH.historyitem_Paragraph_Pr === oChange.Type)
		return false;

	if (AscDFH.historyitem_Paragraph_PrReviewInfo === oChange.Type)
		this.New.ReviewInfo = oChange.New;

	return true;
};
CChangesParagraphPrChange.prototype.IsChangedNumbering = function()
{
	var oNewNumPr = this.New.PrChange ? this.New.PrChange.NumPr : null;
	var oOldNumPr = this.Old.PrChange ? this.Old.PrChange.NumPr : null;

	if ((!oNewNumPr && oOldNumPr)
		|| (oNewNumPr && !oOldNumPr)
		|| (oNewNumPr && oOldNumPr && (oNewNumPr.NumId !== oOldNumPr.NumId || oNewNumPr.Lvl !== oOldNumPr.Lvl)))
	{
		return true;
	}

	return false;
};
CChangesParagraphPrChange.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphPrReviewInfo(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphPrReviewInfo.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphPrReviewInfo.prototype.constructor = CChangesParagraphPrReviewInfo;
CChangesParagraphPrReviewInfo.prototype.Type = AscDFH.historyitem_Paragraph_PrReviewInfo;
CChangesParagraphPrReviewInfo.prototype.private_CreateObject = function()
{
	return new AscWord.ReviewInfo();
};
CChangesParagraphPrReviewInfo.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.ReviewInfo = Value;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphPrReviewInfo.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (oChange.Type === this.Type || AscDFH.historyitem_Paragraph_Pr === oChange.Type || AscDFH.historyitem_Paragraph_PrChange === oChange.Type)
		return false;

	return true;
};
CChangesParagraphPrReviewInfo.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesParagraphOutlineLvl(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphOutlineLvl.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesParagraphOutlineLvl.prototype.constructor = CChangesParagraphOutlineLvl;
CChangesParagraphOutlineLvl.prototype.Type = AscDFH.historyitem_Paragraph_OutlineLvl;
CChangesParagraphOutlineLvl.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.OutlineLvl = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphOutlineLvl.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphOutlineLvl.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphOutlineLvl.prototype.IsNeedRecalculate = function()
{
	return false;
};
CChangesParagraphOutlineLvl.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesParagraphSuppressLineNumbers(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphSuppressLineNumbers.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesParagraphSuppressLineNumbers.prototype.constructor = CChangesParagraphSuppressLineNumbers;
CChangesParagraphSuppressLineNumbers.prototype.Type = AscDFH.historyitem_Paragraph_SuppressLineNumbers;
CChangesParagraphSuppressLineNumbers.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;
	oParagraph.Pr.SuppressLineNumbers = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);

	// TODO: Запросить пересчет номеров строк
};
CChangesParagraphSuppressLineNumbers.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphSuppressLineNumbers.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphSuppressLineNumbers.prototype.IsNeedRecalculate = function()
{
	return false;
};
CChangesParagraphSuppressLineNumbers.prototype.IsNeedRecalculateLineNumbers = function()
{
	return true;
};
CChangesParagraphSuppressLineNumbers.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphShdFill(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphShdFill.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphShdFill.prototype.constructor = CChangesParagraphShdFill;
CChangesParagraphShdFill.prototype.Type = AscDFH.historyitem_Paragraph_Shd_Fill;
CChangesParagraphShdFill.prototype.private_CreateObject = function()
{
	return new CDocumentColor(0, 0, 0);
};
CChangesParagraphShdFill.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Shd)
		oParagraph.Pr.Shd = new CDocumentShd();

	oParagraph.Pr.Shd.Fill = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphShdFill.prototype.Merge = private_ParagraphChangesOnMergeShdPr;
CChangesParagraphShdFill.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphShdFill.prototype.IsNeedRecalculate = function()
{
	return false;
};
CChangesParagraphShdFill.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesParagraphShdThemeFill(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphShdThemeFill.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesParagraphShdThemeFill.prototype.constructor = CChangesParagraphShdThemeFill;
CChangesParagraphShdThemeFill.prototype.Type = AscDFH.historyitem_Paragraph_Shd_ThemeFill;
CChangesParagraphShdThemeFill.prototype.private_CreateObject = function()
{
	return new AscFormat.CUniFill();
};
CChangesParagraphShdThemeFill.prototype.private_SetValue = function(Value)
{
	var oParagraph = this.Class;

	if (undefined === oParagraph.Pr.Shd)
		oParagraph.Pr.Shd = new CDocumentShd();

	oParagraph.Pr.Shd.ThemeFill = Value;

	oParagraph.CompiledPr.NeedRecalc = true;
	oParagraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphShdThemeFill.prototype.Merge = private_ParagraphChangesOnMergeShdPr;
CChangesParagraphShdThemeFill.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphShdThemeFill.prototype.IsNeedRecalculate = function()
{
	return false;
};
CChangesParagraphShdThemeFill.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesParagraphBidi(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesParagraphBidi.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesParagraphBidi.prototype.constructor = CChangesParagraphBidi;
CChangesParagraphBidi.prototype.Type = AscDFH.historyitem_Paragraph_Bidi;
CChangesParagraphBidi.prototype.private_SetValue = function(value)
{
	let paragraph = this.Class;
	paragraph.Pr.Bidi = value;
	paragraph.CompiledPr.NeedRecalc = true;
	paragraph.private_UpdateTrackRevisionOnChangeParaPr(false);
};
CChangesParagraphBidi.prototype.Merge = private_ParagraphChangesOnMergePr;
CChangesParagraphBidi.prototype.Load = private_ParagraphChangesOnLoadPr;
CChangesParagraphBidi.prototype.IsNeedRecalculate = function()
{
	return true;
};
CChangesParagraphBidi.prototype.CheckLock = private_ParagraphContentChangesCheckLock;

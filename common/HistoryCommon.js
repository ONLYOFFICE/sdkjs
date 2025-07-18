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

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
	function(window, undefined)
{
	function GetHistoryPointStringDescription(nDescription)
	{
		var sString = "Unknown";
		switch (nDescription)
		{
			case AscDFH.historydescription_Cut                                         :
				sString = "Cut";
				break;
			case AscDFH.historydescription_PasteButtonIE                               :
				sString = "PasteButtonIE";
				break;
			case AscDFH.historydescription_PasteButtonNotIE                            :
				sString = "PasteButtonNotIE";
				break;
			case AscDFH.historydescription_ChartDrawingObjects                         :
				sString = "ChartDrawingObjects";
				break;
			case AscDFH.historydescription_CommonControllerCheckChartText              :
				sString = "CommonControllerCheckChartText";
				break;
			case AscDFH.historydescription_CommonControllerUnGroup                     :
				sString = "CommonControllerUnGroup";
				break;
			case AscDFH.historydescription_CommonControllerCheckSelected               :
				sString = "CommonControllerCheckSelected";
				break;
			case AscDFH.historydescription_CommonControllerSetGraphicObject            :
				sString = "CommonControllerSetGraphicObject";
				break;
			case AscDFH.historydescription_CommonStatesAddNewShape                     :
				sString = "CommonStatesAddNewShape";
				break;
			case AscDFH.historydescription_CommonStatesRotate                          :
				sString = "CommonStatesRotate";
				break;
			case AscDFH.historydescription_PasteNative                                 :
				sString = "PasteNative";
				break;
			case AscDFH.historydescription_Document_GroupUnGroup                       :
				sString = "Document_GroupUnGroup";
				break;
			case AscDFH.historydescription_Document_SetDefaultLanguage                 :
				sString = "Document_SetDefaultLanguage";
				break;
			case AscDFH.historydescription_Document_ChangeColorScheme                  :
				sString = "Document_ChangeColorScheme";
				break;
			case AscDFH.historydescription_Document_AddChart                           :
				sString = "Document_AddChart";
				break;
			case AscDFH.historydescription_Document_EditChart                          :
				sString = "Document_EditChart";
				break;
			case AscDFH.historydescription_Document_DragText                           :
				sString = "Document_DragText";
				break;
			case AscDFH.historydescription_Document_DocumentContentExtendToPos         :
				sString = "Document_DocumentContentExtendToPos";
				break;
			case AscDFH.historydescription_Document_AddHeader                          :
				sString = "Document_AddHeader";
				break;
			case AscDFH.historydescription_Document_AddFooter                          :
				sString = "Document_AddFooter";
				break;
			case AscDFH.historydescription_Document_ParagraphExtendToPos               :
				sString = "Document_ParagraphExtendToPos";
				break;
			case AscDFH.historydescription_Document_ParagraphChangeFrame               :
				sString = "Document_ParagraphChangeFrame";
				break;
			case AscDFH.historydescription_Document_ReplaceAll                         :
				sString = "Document_ReplaceAll";
				break;
			case AscDFH.historydescription_Document_ReplaceSingle                      :
				sString = "Document_ReplaceSingle";
				break;
			case AscDFH.historydescription_Document_TableAddNewRowByTab                :
				sString = "Document_TableAddNewRowByTab";
				break;
			case AscDFH.historydescription_Document_AddNewShape                        :
				sString = "Document_AddNewShape";
				break;
			case AscDFH.historydescription_Document_EditWrapPolygon                    :
				sString = "Document_EditWrapPolygon";
				break;
			case AscDFH.historydescription_Document_MoveInlineObject                   :
				sString = "Document_MoveInlineObject";
				break;
			case AscDFH.historydescription_Document_CopyAndMoveInlineObject            :
				sString = "Document_CopyAndMoveInlineObject";
				break;
			case AscDFH.historydescription_Document_RotateInlineDrawing                :
				sString = "Document_RotateInlineDrawing";
				break;
			case AscDFH.historydescription_Document_RotateFlowDrawingCtrl              :
				sString = "Document_RotateFlowDrawingCtrl";
				break;
			case AscDFH.historydescription_Document_RotateFlowDrawingNoCtrl            :
				sString = "Document_RotateFlowDrawingNoCtrl";
				break;
			case AscDFH.historydescription_Document_MoveInGroup                        :
				sString = "Document_MoveInGroup";
				break;
			case AscDFH.historydescription_Document_ChangeWrapContour                  :
				sString = "Document_ChangeWrapContour";
				break;
			case AscDFH.historydescription_Document_ChangeWrapContourAddPoint          :
				sString = "Document_ChangeWrapContourAddPoint";
				break;
			case AscDFH.historydescription_Document_GrObjectsBringToFront              :
				sString = "Document_GrObjectsBringToFront";
				break;
			case AscDFH.historydescription_Document_GrObjectsBringForwardGroup         :
				sString = "Document_GrObjectsBringForwardGroup";
				break;
			case AscDFH.historydescription_Document_GrObjectsBringForward              :
				sString = "Document_GrObjectsBringForward";
				break;
			case AscDFH.historydescription_Document_GrObjectsSendToBackGroup           :
				sString = "Document_GrObjectsSendToBackGroup";
				break;
			case AscDFH.historydescription_Document_GrObjectsSendToBack                :
				sString = "Document_GrObjectsSendToBack";
				break;
			case AscDFH.historydescription_Document_GrObjectsBringBackwardGroup        :
				sString = "Document_GrObjectsBringBackwardGroup";
				break;
			case AscDFH.historydescription_Document_GrObjectsBringBackward             :
				sString = "Document_GrObjectsBringBackward";
				break;
			case AscDFH.historydescription_Document_GrObjectsChangeWrapPolygon         :
				sString = "Document_GrObjectsChangeWrapPolygon";
				break;
			case AscDFH.historydescription_Document_MathAutoCorrect                    :
				sString = "Document_MathAutoCorrect";
				break;
			case AscDFH.historydescription_Document_SetFramePrWithFontFamily           :
				sString = "Document_SetFramePrWithFontFamily";
				break;
			case AscDFH.historydescription_Document_SetFramePr                         :
				sString = "Document_SetFramePr";
				break;
			case AscDFH.historydescription_Document_SetFramePrWithFontFamilyLong       :
				sString = "Document_SetFramePrWithFontFamilyLong";
				break;
			case AscDFH.historydescription_Document_SetTextFontName                    :
				sString = "Document_SetTextFontName";
				break;
			case AscDFH.historydescription_Document_SetTextFontSize                    :
				sString = "Document_SetTextFontSize";
				break;
			case AscDFH.historydescription_Document_SetTextBold                        :
				sString = "Document_SetTextBold";
				break;
			case AscDFH.historydescription_Document_SetTextItalic                      :
				sString = "Document_SetTextItalic";
				break;
			case AscDFH.historydescription_Document_SetTextUnderline                   :
				sString = "Document_SetTextUnderline";
				break;
			case AscDFH.historydescription_Document_SetTextStrikeout                   :
				sString = "Document_SetTextStrikeout";
				break;
			case AscDFH.historydescription_Document_SetTextDStrikeout                  :
				sString = "Document_SetTextDStrikeout";
				break;
			case AscDFH.historydescription_Document_SetTextSpacing                     :
				sString = "Document_SetTextSpacing";
				break;
			case AscDFH.historydescription_Document_SetTextCaps                        :
				sString = "Document_SetTextCaps";
				break;
			case AscDFH.historydescription_Document_SetTextSmallCaps                   :
				sString = "Document_SetTextSmallCaps";
				break;
			case AscDFH.historydescription_Document_SetTextPosition                    :
				sString = "Document_SetTextPosition";
				break;
			case AscDFH.historydescription_Document_SetTextLang                        :
				sString = "Document_SetTextLang";
				break;
			case AscDFH.historydescription_Document_SetParagraphLineSpacing            :
				sString = "Document_SetParagraphLineSpacing";
				break;
			case AscDFH.historydescription_Document_SetParagraphLineSpacingBeforeAfter :
				sString = "Document_SetParagraphLineSpacingBeforeAfter";
				break;
			case AscDFH.historydescription_Document_IncFontSize                        :
				sString = "Document_IncFontSize";
				break;
			case AscDFH.historydescription_Document_DecFontSize                        :
				sString = "Document_DecFontSize";
				break;
			case AscDFH.historydescription_Document_SetParagraphBorders                :
				sString = "Document_SetParagraphBorders";
				break;
			case AscDFH.historydescription_Document_SetParagraphPr                     :
				sString = "Document_SetParagraphPr";
				break;
			case AscDFH.historydescription_Document_SetParagraphAlign                  :
				sString = "Document_SetParagraphAlign";
				break;
			case AscDFH.historydescription_Document_SetTextVertAlign                   :
				sString = "Document_SetTextVertAlign";
				break;
			case AscDFH.historydescription_Document_SetParagraphNumbering              :
				sString = "Document_SetParagraphNumbering";
				break;
			case AscDFH.historydescription_Document_SetParagraphStyle                  :
				sString = "Document_SetParagraphStyle";
				break;
			case AscDFH.historydescription_Document_SetParagraphPageBreakBefore        :
				sString = "Document_SetParagraphPageBreakBefore";
				break;
			case AscDFH.historydescription_Document_SetParagraphWidowControl           :
				sString = "Document_SetParagraphWidowControl";
				break;
			case AscDFH.historydescription_Document_SetParagraphKeepLines              :
				sString = "Document_SetParagraphKeepLines";
				break;
			case AscDFH.historydescription_Document_SetParagraphKeepNext               :
				sString = "Document_SetParagraphKeepNext";
				break;
			case AscDFH.historydescription_Document_SetParagraphContextualSpacing      :
				sString = "Document_SetParagraphContextualSpacing";
				break;
			case AscDFH.historydescription_Document_SetTextHighlightNone               :
				sString = "Document_SetTextHighlightNone";
				break;
			case AscDFH.historydescription_Document_SetTextHighlightColor              :
				sString = "Document_SetTextHighlightColor";
				break;
			case AscDFH.historydescription_Document_SetTextColor                       :
				sString = "Document_SetTextColor";
				break;
			case AscDFH.historydescription_Document_SetParagraphShd                    :
				sString = "Document_SetParagraphShd";
				break;
			case AscDFH.historydescription_Document_SetParagraphIndent                 :
				sString = "Document_SetParagraphIndent";
				break;
			case AscDFH.historydescription_Document_IncParagraphIndent                 :
				sString = "Document_IncParagraphIndent";
				break;
			case AscDFH.historydescription_Document_DecParagraphIndent                 :
				sString = "Document_DecParagraphIndent";
				break;
			case AscDFH.historydescription_Document_SetParagraphIndentRight            :
				sString = "Document_SetParagraphIndentRight";
				break;
			case AscDFH.historydescription_Document_SetParagraphIndentFirstLine        :
				sString = "Document_SetParagraphIndentFirstLine";
				break;
			case AscDFH.historydescription_Document_SetPageOrientation                 :
				sString = "Document_SetPageOrientation";
				break;
			case AscDFH.historydescription_Document_SetPageSize                        :
				sString = "Document_SetPageSize";
				break;
			case AscDFH.historydescription_Document_AddPageBreak                       :
				sString = "Document_AddPageBreak";
				break;
			case AscDFH.historydescription_Document_AddPageNumToHdrFtr                 :
				sString = "Document_AddPageNumToHdrFtr";
				break;
			case AscDFH.historydescription_Document_AddPageNumToCurrentPos             :
				sString = "Document_AddPageNumToCurrentPos";
				break;
			case AscDFH.historydescription_Document_SetHdrFtrDistance                  :
				sString = "Document_SetHdrFtrDistance";
				break;
			case AscDFH.historydescription_Document_SetHdrFtrFirstPage                 :
				sString = "Document_SetHdrFtrFirstPage";
				break;
			case AscDFH.historydescription_Document_SetHdrFtrEvenAndOdd                :
				sString = "Document_SetHdrFtrEvenAndOdd";
				break;
			case AscDFH.historydescription_Document_SetHdrFtrLink                      :
				sString = "Document_SetHdrFtrLink";
				break;
			case AscDFH.historydescription_Document_AddTable                           :
				sString = "Document_AddTable";
				break;
			case AscDFH.historydescription_Document_TableAddRowAbove                   :
				sString = "Document_TableAddRowAbove";
				break;
			case AscDFH.historydescription_Document_TableAddRowBelow                   :
				sString = "Document_TableAddRowBelow";
				break;
			case AscDFH.historydescription_Document_TableAddColumnLeft                 :
				sString = "Document_TableAddColumnLeft";
				break;
			case AscDFH.historydescription_Document_TableAddColumnRight                :
				sString = "Document_TableAddColumnRight";
				break;
			case AscDFH.historydescription_Document_TableRemoveRow                     :
				sString = "Document_TableRemoveRow";
				break;
			case AscDFH.historydescription_Document_TableRemoveColumn                  :
				sString = "Document_TableRemoveColumn";
				break;
			case AscDFH.historydescription_Document_RemoveTable                        :
				sString = "Document_RemoveTable";
				break;
			case AscDFH.historydescription_Document_MergeTableCells                    :
				sString = "Document_MergeTableCells";
				break;
			case AscDFH.historydescription_Document_SplitTableCells                    :
				sString = "Document_SplitTableCells";
				break;
			case AscDFH.historydescription_Document_ApplyTablePr                       :
				sString = "Document_ApplyTablePr";
				break;
			case AscDFH.historydescription_Document_AddImageUrl                        :
				sString = "Document_AddImageUrl";
				break;
			case AscDFH.historydescription_Document_AddImageUrlLong                    :
				sString = "Document_AddImageUrlLong";
				break;
			case AscDFH.historydescription_Document_AddImageToPage                     :
				sString = "Document_AddImageToPage";
				break;
			case AscDFH.historydescription_Document_ApplyImagePrWithUrl                :
				sString = "Document_ApplyImagePrWithUrl";
				break;
			case AscDFH.historydescription_Document_ApplyImagePrWithUrlLong            :
				sString = "Document_ApplyImagePrWithUrlLong";
				break;
			case AscDFH.historydescription_Document_ApplyImagePrWithFillUrl            :
				sString = "Document_ApplyImagePrWithFillUrl";
				break;
			case AscDFH.historydescription_Document_ApplyImagePrWithFillUrlLong        :
				sString = "Document_ApplyImagePrWithFillUrlLong";
				break;
			case AscDFH.historydescription_Document_ApplyImagePr                       :
				sString = "Document_ApplyImagePr";
				break;
			case AscDFH.historydescription_Document_AddHyperlink                       :
				sString = "Document_AddHyperlink";
				break;
			case AscDFH.historydescription_Document_ChangeHyperlink                    :
				sString = "Document_ChangeHyperlink";
				break;
			case AscDFH.historydescription_Document_RemoveHyperlink                    :
				sString = "Document_RemoveHyperlink";
				break;
			case AscDFH.historydescription_Document_ReplaceMisspelledWord              :
				sString = "Document_ReplaceMisspelledWord";
				break;
			case AscDFH.historydescription_Document_AddComment                         :
				sString = "Document_AddComment";
				break;
			case AscDFH.historydescription_Document_RemoveComment                      :
				sString = "Document_RemoveComment";
				break;
			case AscDFH.historydescription_Document_ChangeComment                      :
				sString = "Document_ChangeComment";
				break;
			case AscDFH.historydescription_Document_SetTextFontNameLong                :
				sString = "Document_SetTextFontNameLong";
				break;
			case AscDFH.historydescription_Document_AddImage                           :
				sString = "Document_AddImage";
				break;
			case AscDFH.historydescription_Document_ClearFormatting                    :
				sString = "Document_ClearFormatting";
				break;
			case AscDFH.historydescription_Document_AddSectionBreak                    :
				sString = "Document_AddSectionBreak";
				break;
			case AscDFH.historydescription_Document_AddMath                            :
				sString = "Document_AddMath";
				break;
			case AscDFH.historydescription_Document_SetParagraphTabs                   :
				sString = "Document_SetParagraphTabs";
				break;
			case AscDFH.historydescription_Document_SetParagraphIndentFromRulers       :
				sString = "Document_SetParagraphIndentFromRulers";
				break;
			case AscDFH.historydescription_Document_SetDocumentMargin_Hor              :
				sString = "Document_SetDocumentMargin_Hor";
				break;
			case AscDFH.historydescription_Document_SetTableMarkup_Hor                 :
				sString = "Document_SetTableMarkup_Hor";
				break;
			case AscDFH.historydescription_Document_SetDocumentMargin_Ver              :
				sString = "Document_SetDocumentMargin_Ver";
				break;
			case AscDFH.historydescription_Document_SetHdrFtrBounds                    :
				sString = "Document_SetHdrFtrBounds";
				break;
			case AscDFH.historydescription_Document_SetTableMarkup_Ver                 :
				sString = "Document_SetTableMarkup_Ver";
				break;
			case AscDFH.historydescription_Document_DocumentExtendToPos                :
				sString = "Document_DocumentExtendToPos";
				break;
			case AscDFH.historydescription_Document_AddDropCap                         :
				sString = "Document_AddDropCap";
				break;
			case AscDFH.historydescription_Document_RemoveDropCap                      :
				sString = "Document_RemoveDropCap";
				break;
			case AscDFH.historydescription_Document_SetTextHighlight                   :
				sString = "Document_SetTextHighlight";
				break;
			case AscDFH.historydescription_Document_BackSpaceButton                    :
				sString = "Document_BackSpaceButton";
				break;
			case AscDFH.historydescription_Document_MoveParagraphByTab                 :
				sString = "Document_MoveParagraphByTab";
				break;
			case AscDFH.historydescription_Document_AddTab                             :
				sString = "Document_AddTab";
				break;
			case AscDFH.historydescription_Document_EnterButton                        :
				sString = "Document_EnterButton";
				break;
			case AscDFH.historydescription_Document_SpaceButton                        :
				sString = "Document_SpaceButton";
				break;
			case AscDFH.historydescription_Document_ShiftInsert                        :
				sString = "Document_ShiftInsert";
				break;
			case AscDFH.historydescription_Document_ShiftInsertSafari                  :
				sString = "Document_ShiftInsertSafari";
				break;
			case AscDFH.historydescription_Document_DeleteButton                       :
				sString = "Document_DeleteButton";
				break;
			case AscDFH.historydescription_Document_ShiftDeleteButton                  :
				sString = "Document_ShiftDeleteButton";
				break;
			case AscDFH.historydescription_Document_Shortcut_SetStyleHeading1:
				sString = "Document_Shortcut_SetStyleHeading1";
				break;
			case AscDFH.historydescription_Document_Shortcut_SetStyleHeading2:
				sString = "Document_Shortcut_SetStyleHeading2";
				break;
			case AscDFH.historydescription_Document_Shortcut_SetStyleHeading3:
				sString = "Document_Shortcut_SetStyleHeading3";
				break;
			case AscDFH.historydescription_Document_SetTextStrikeoutHotKey             :
				sString = "Document_SetTextStrikeoutHotKey";
				break;
			case AscDFH.historydescription_Document_SetTextBoldHotKey                  :
				sString = "Document_SetTextBoldHotKey";
				break;
			case AscDFH.historydescription_Document_SetParagraphAlignHotKey            :
				sString = "Document_SetParagraphAlignHotKey";
				break;
			case AscDFH.historydescription_Document_AddEuroLetter                      :
				sString = "Document_AddEuroLetter";
				break;
			case AscDFH.historydescription_Document_SetTextItalicHotKey                :
				sString = "Document_SetTextItalicHotKey";
				break;
			case AscDFH.historydescription_Document_SetParagraphAlignHotKey2           :
				sString = "Document_SetParagraphAlignHotKey2";
				break;
			case AscDFH.historydescription_Document_SetParagraphNumberingHotKey        :
				sString = "Document_SetParagraphNumberingHotKey";
				break;
			case AscDFH.historydescription_Document_SetParagraphAlignHotKey3           :
				sString = "Document_SetParagraphAlignHotKey3";
				break;
			case AscDFH.historydescription_Document_AddPageNumHotKey                   :
				sString = "Document_AddPageNumHotKey";
				break;
			case AscDFH.historydescription_Document_SetParagraphAlignHotKey4           :
				sString = "Document_SetParagraphAlignHotKey4";
				break;
			case AscDFH.historydescription_Document_SetTextUnderlineHotKey             :
				sString = "Document_SetTextUnderlineHotKey";
				break;
			case AscDFH.historydescription_Document_FormatPasteHotKey                  :
				sString = "Document_FormatPasteHotKey";
				break;
			case AscDFH.historydescription_Document_PasteHotKey                        :
				sString = "Document_PasteHotKey";
				break;
			case AscDFH.historydescription_Document_PasteSafariHotKey                  :
				sString = "Document_PasteSafariHotKey";
				break;
			case AscDFH.historydescription_Document_CutHotKey                          :
				sString = "Document_CutHotKey";
				break;
			case AscDFH.historydescription_Document_SetTextVertAlignHotKey             :
				sString = "Document_SetTextVertAlignHotKey";
				break;
			case AscDFH.historydescription_Document_AddMathHotKey                      :
				sString = "Document_AddMathHotKey";
				break;
			case AscDFH.historydescription_Document_SetTextVertAlignHotKey2            :
				sString = "Document_SetTextVertAlignHotKey2";
				break;
			case AscDFH.historydescription_Document_MinusButton                        :
				sString = "Document_MinusButton";
				break;
			case AscDFH.historydescription_Document_SetTextVertAlignHotKey3            :
				sString = "Document_SetTextVertAlignHotKey3";
				break;
			case AscDFH.historydescription_Document_AddLetter                          :
				sString = "Document_AddLetter";
				break;
			case AscDFH.historydescription_Document_MoveTableBorder                    :
				sString = "Document_MoveTableBorder";
				break;
			case AscDFH.historydescription_Document_FormatPasteHotKey2                 :
				sString = "Document_FormatPasteHotKey2";
				break;
			case AscDFH.historydescription_Document_SetTextHighlight2                  :
				sString = "Document_SetTextHighlight2";
				break;
			case AscDFH.historydescription_Document_AddTextFromTextBox                 :
				sString = "Document_AddTextFromTextBox";
				break;
			case AscDFH.historydescription_Document_AddMailMergeField                  :
				sString = "Document_AddMailMergeField";
				break;
			case AscDFH.historydescription_Document_MoveInlineTable                    :
				sString = "Document_MoveInlineTable";
				break;
			case AscDFH.historydescription_Document_MoveFlowTable                      :
				sString = "Document_MoveFlowTable";
				break;
			case AscDFH.historydescription_Document_RestoreFieldTemplateText           :
				sString = "Document_RestoreFieldTemplateText";
				break;

			case AscDFH.historydescription_Spreadsheet_SetCellFontName                 :
				sString = "Spreadsheet_SetCellFontName";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellFontSize                 :
				sString = "Spreadsheet_SetCellFontSize";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellBold                     :
				sString = "Spreadsheet_SetCellBold";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellItalic                   :
				sString = "Spreadsheet_SetCellItalic";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellUnderline                :
				sString = "Spreadsheet_SetCellUnderline";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellStrikeout                :
				sString = "Spreadsheet_SetCellStrikeout";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellSubscript                :
				sString = "Spreadsheet_SetCellSubscript";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellSuperscript              :
				sString = "Spreadsheet_SetCellSuperscript";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellAlign                    :
				sString = "Spreadsheet_SetCellAlign";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellVertAlign                :
				sString = "Spreadsheet_SetCellVertAlign";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellTextColor                :
				sString = "Spreadsheet_SetCellTextColor";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellBackgroundColor          :
				sString = "Spreadsheet_SetCellBackgroundColor";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellIncreaseFontSize         :
				sString = "Spreadsheet_SetCellIncreaseFontSize";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellDecreaseFontSize         :
				sString = "Spreadsheet_SetCellDecreaseFontSize";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellHyperlinkAdd             :
				sString = "Spreadsheet_SetCellHyperlinkAdd";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellHyperlinkModify          :
				sString = "Spreadsheet_SetCellHyperlinkModify";
				break;
			case AscDFH.historydescription_Spreadsheet_SetCellHyperlinkRemove          :
				sString = "Spreadsheet_SetCellHyperlinkRemove";
				break;
			case AscDFH.historydescription_Spreadsheet_EditChart                       :
				sString = "Spreadsheet_EditChart";
				break;
			case AscDFH.historydescription_Spreadsheet_Remove                          :
				sString = "Spreadsheet_Remove";
				break;
			case AscDFH.historydescription_Spreadsheet_AddTab                          :
				sString = "Spreadsheet_AddTab";
				break;
			case AscDFH.historydescription_Spreadsheet_AddNewParagraph                 :
				sString = "Spreadsheet_AddNewParagraph";
				break;
			case AscDFH.historydescription_Spreadsheet_AddSpace                        :
				sString = "Spreadsheet_AddSpace";
				break;
			case AscDFH.historydescription_Spreadsheet_AddItem                         :
				sString = "Spreadsheet_AddItem";
				break;
			case AscDFH.historydescription_Spreadsheet_PutPrLineSpacing                :
				sString = "Spreadsheet_PutPrLineSpacing";
				break;
			case AscDFH.historydescription_Spreadsheet_SetParagraphSpacing             :
				sString = "Spreadsheet_SetParagraphSpacing";
				break;
			case AscDFH.historydescription_Spreadsheet_SetGraphicObjectsProps          :
				sString = "Spreadsheet_SetGraphicObjectsProps";
				break;
			case AscDFH.historydescription_Spreadsheet_ParaApply                       :
				sString = "Spreadsheet_ParaApply";
				break;
			case AscDFH.historydescription_Spreadsheet_GraphicObjectLayer              :
				sString = "Spreadsheet_GraphicObjectLayer";
				break;
			case AscDFH.historydescription_Spreadsheet_ParagraphAdd                    :
				sString = "Spreadsheet_ParagraphAdd";
				break;
			case AscDFH.historydescription_Spreadsheet_CreateGroup                     :
				sString = "Spreadsheet_CreateGroup";
				break;
			case AscDFH.historydescription_CommonDrawings_ChangeAdj                    :
				sString = "CommonDrawings_ChangeAdj";
				break;
			case AscDFH.historydescription_CommonDrawings_EndTrack                     :
				sString = "CommonDrawings_EndTrack";
				break;
			case AscDFH.historydescription_CommonDrawings_CopyCtrl                     :
				sString = "CommonDrawings_CopyCtrl";
				break;
			case AscDFH.historydescription_Presentation_ParaApply                      :
				sString = "Presentation_ParaApply";
				break;
			case AscDFH.historydescription_Presentation_ParaFormatPaste                :
				sString = "Presentation_ParaFormatPaste";
				break;
			case AscDFH.historydescription_Presentation_AddNewParagraph                :
				sString = "Presentation_AddNewParagraph";
				break;
			case AscDFH.historydescription_Presentation_CreateGroup                    :
				sString = "Presentation_CreateGroup";
				break;
			case AscDFH.historydescription_Presentation_UnGroup                        :
				sString = "Presentation_UnGroup";
				break;
			case AscDFH.historydescription_Presentation_AddChart                       :
				sString = "Presentation_AddChart";
				break;
			case AscDFH.historydescription_Presentation_EditChart                      :
				sString = "Presentation_EditChart";
				break;
			case AscDFH.historydescription_Presentation_ParagraphAdd                   :
				sString = "Presentation_ParagraphAdd";
				break;
			case AscDFH.historydescription_Presentation_ParagraphClearFormatting       :
				sString = "Presentation_ParagraphClearFormatting";
				break;
			case AscDFH.historydescription_Presentation_SetParagraphAlign              :
				sString = "Presentation_SetParagraphAlign";
				break;
			case AscDFH.historydescription_Presentation_SetParagraphSpacing            :
				sString = "Presentation_SetParagraphSpacing";
				break;
			case AscDFH.historydescription_Presentation_SetParagraphTabs               :
				sString = "Presentation_SetParagraphTabs";
				break;
			case AscDFH.historydescription_Presentation_SetParagraphIndent             :
				sString = "Presentation_SetParagraphIndent";
				break;
			case AscDFH.historydescription_Presentation_SetParagraphNumbering          :
				sString = "Presentation_SetParagraphNumbering";
				break;
			case AscDFH.historydescription_Presentation_ParagraphIncDecFontSize        :
				sString = "Presentation_ParagraphIncDecFontSize";
				break;
			case AscDFH.historydescription_Presentation_ParagraphIncDecIndent          :
				sString = "Presentation_ParagraphIncDecIndent";
				break;
			case AscDFH.historydescription_Presentation_SetImageProps                  :
				sString = "Presentation_SetImageProps";
				break;
			case AscDFH.historydescription_Presentation_SetShapeProps                  :
				sString = "Presentation_SetShapeProps";
				break;
			case AscDFH.historydescription_Presentation_ChartApply                     :
				sString = "Presentation_ChartApply";
				break;
			case AscDFH.historydescription_Presentation_ChangeShapeType                :
				sString = "Presentation_ChangeShapeType";
				break;
			case AscDFH.historydescription_Presentation_SetVerticalAlign               :
				sString = "Presentation_SetVerticalAlign";
				break;
			case AscDFH.historydescription_Presentation_HyperlinkAdd                   :
				sString = "Presentation_HyperlinkAdd";
				break;
			case AscDFH.historydescription_Presentation_HyperlinkModify                :
				sString = "Presentation_HyperlinkModify";
				break;
			case AscDFH.historydescription_Presentation_HyperlinkRemove                :
				sString = "Presentation_HyperlinkRemove";
				break;
			case AscDFH.historydescription_Presentation_DistHor                        :
				sString = "Presentation_DistHor";
				break;
			case AscDFH.historydescription_Presentation_DistVer                        :
				sString = "Presentation_DistVer";
				break;
			case AscDFH.historydescription_Presentation_BringToFront                   :
				sString = "Presentation_BringToFront";
				break;
			case AscDFH.historydescription_Presentation_BringForward                   :
				sString = "Presentation_BringForward";
				break;
			case AscDFH.historydescription_Presentation_SendToBack                     :
				sString = "Presentation_SendToBack";
				break;
			case AscDFH.historydescription_Presentation_BringBackward                  :
				sString = "Presentation_BringBackward";
				break;
			case AscDFH.historydescription_Presentation_ApplyTransition                :
				sString = "Presentation_ApplyTransition";
				break;
			case AscDFH.historydescription_Presentation_MoveSlidesToEnd                :
				sString = "Presentation_MoveSlidesToEnd";
				break;
			case AscDFH.historydescription_Presentation_MoveSlidesNextPos              :
				sString = "Presentation_MoveSlidesNextPos";
				break;
			case AscDFH.historydescription_Presentation_MoveSlidesPrevPos              :
				sString = "Presentation_MoveSlidesPrevPos";
				break;
			case AscDFH.historydescription_Presentation_MoveSlidesToStart              :
				sString = "Presentation_MoveSlidesToStart";
				break;
			case AscDFH.historydescription_Presentation_MoveComments                   :
				sString = "Presentation_MoveComments";
				break;
			case AscDFH.historydescription_Presentation_TableBorder                    :
				sString = "Presentation_TableBorder";
				break;
			case AscDFH.historydescription_Presentation_AddFlowImage                   :
				sString = "Presentation_AddFlowImage";
				break;
			case AscDFH.historydescription_Presentation_AddFlowTable                   :
				sString = "Presentation_AddFlowTable";
				break;
			case AscDFH.historydescription_Presentation_ChangeBackground               :
				sString = "Presentation_ChangeBackground";
				break;
			case AscDFH.historydescription_Presentation_AddNextSlide                   :
				sString = "Presentation_AddNextSlide";
				break;
			case AscDFH.historydescription_Presentation_ShiftSlides                    :
				sString = "Presentation_ShiftSlides";
				break;
			case AscDFH.historydescription_Presentation_DeleteSlides                   :
				sString = "Presentation_DeleteSlides";
				break;
			case AscDFH.historydescription_Presentation_ChangeLayout                   :
				sString = "Presentation_ChangeLayout";
				break;
			case AscDFH.historydescription_Presentation_ChangeSlideSize                :
				sString = "Presentation_ChangeSlideSize";
				break;
			case AscDFH.historydescription_Presentation_ChangeColorScheme              :
				sString = "Presentation_ChangeColorScheme";
				break;
			case AscDFH.historydescription_Presentation_AddComment                     :
				sString = "Presentation_AddComment";
				break;
			case AscDFH.historydescription_Presentation_ChangeComment                  :
				sString = "Presentation_ChangeComment";
				break;
			case AscDFH.historydescription_Presentation_PutTextPrFontName              :
				sString = "Presentation_PutTextPrFontName";
				break;
			case AscDFH.historydescription_Presentation_PutTextPrFontSize              :
				sString = "Presentation_PutTextPrFontSize";
				break;
			case AscDFH.historydescription_Presentation_PutTextPrBold                  :
				sString = "Presentation_PutTextPrBold";
				break;
			case AscDFH.historydescription_Presentation_PutTextPrItalic                :
				sString = "Presentation_PutTextPrItalic";
				break;
			case AscDFH.historydescription_Presentation_PutTextPrUnderline             :
				sString = "Presentation_PutTextPrUnderline";
				break;
			case AscDFH.historydescription_Presentation_PutTextPrStrikeout             :
				sString = "Presentation_PutTextPrStrikeout";
				break;
			case AscDFH.historydescription_Presentation_PutTextPrLineSpacing           :
				sString = "Presentation_PutTextPrLineSpacing";
				break;
			case AscDFH.historydescription_Presentation_PutTextPrSpacingBeforeAfter    :
				sString = "Presentation_PutTextPrSpacingBeforeAfter";
				break;
			case AscDFH.historydescription_Presentation_PutTextPrIncreaseFontSize      :
				sString = "Presentation_PutTextPrIncreaseFontSize";
				break;
			case AscDFH.historydescription_Presentation_PutTextPrDecreaseFontSize      :
				sString = "Presentation_PutTextPrDecreaseFontSize";
				break;
			case AscDFH.historydescription_Presentation_PutTextPrAlign                 :
				sString = "Presentation_PutTextPrAlign";
				break;
			case AscDFH.historydescription_Presentation_PutTextPrBaseline              :
				sString = "Presentation_PutTextPrBaseline";
				break;
			case AscDFH.historydescription_Presentation_PutTextPrListType              :
				sString = "Presentation_PutTextPrListType";
				break;
			case AscDFH.historydescription_Presentation_PutTextColor                   :
				sString = "Presentation_PutTextColor";
				break;
			case AscDFH.historydescription_Presentation_PutTextColor2                  :
				sString = "Presentation_PutTextColor2";
				break;
			case AscDFH.historydescription_Presentation_PutPrIndent                    :
				sString = "Presentation_PutPrIndent";
				break;
			case AscDFH.historydescription_Presentation_PutPrIndentRight               :
				sString = "Presentation_PutPrIndentRight";
				break;
			case AscDFH.historydescription_Presentation_PutPrFirstLineIndent           :
				sString = "Presentation_PutPrFirstLineIndent";
				break;
			case AscDFH.historydescription_Presentation_AddPageBreak                   :
				sString = "Presentation_AddPageBreak";
				break;
			case AscDFH.historydescription_Presentation_AddRowAbove                    :
				sString = "Presentation_AddRowAbove";
				break;
			case AscDFH.historydescription_Presentation_AddRowBelow                    :
				sString = "Presentation_AddRowBelow";
				break;
			case AscDFH.historydescription_Presentation_AddColLeft                     :
				sString = "Presentation_AddColLeft";
				break;
			case AscDFH.historydescription_Presentation_AddColRight                    :
				sString = "Presentation_AddColRight";
				break;
			case AscDFH.historydescription_Presentation_RemoveRow                      :
				sString = "Presentation_RemoveRow";
				break;
			case AscDFH.historydescription_Presentation_RemoveCol                      :
				sString = "Presentation_RemoveCol";
				break;
			case AscDFH.historydescription_Presentation_RemoveTable                    :
				sString = "Presentation_RemoveTable";
				break;
			case AscDFH.historydescription_Presentation_MergeCells                     :
				sString = "Presentation_MergeCells";
				break;
			case AscDFH.historydescription_Presentation_SplitCells                     :
				sString = "Presentation_SplitCells";
				break;
			case AscDFH.historydescription_Presentation_TblApply                       :
				sString = "Presentation_TblApply";
				break;
			case AscDFH.historydescription_Presentation_RemoveComment                  :
				sString = "Presentation_RemoveComment";
				break;
			case AscDFH.historydescription_Presentation_EndFontLoad                    :
				sString = "Presentation_EndFontLoad";
				break;
			case AscDFH.historydescription_Presentation_ChangeTheme                    :
				sString = "Presentation_ChangeTheme";
				break;
			case AscDFH.historydescription_Presentation_TableMoveFromRulers            :
				sString = "Presentation_TableMoveFromRulers";
				break;
			case AscDFH.historydescription_Presentation_TableMoveFromRulersInline      :
				sString = "Presentation_TableMoveFromRulersInline";
				break;
			case AscDFH.historydescription_Presentation_PasteOnThumbnails              :
				sString = "Presentation_PasteOnThumbnails";
				break;
			case AscDFH.historydescription_Presentation_PasteOnThumbnailsSafari        :
				sString = "Presentation_PasteOnThumbnailsSafari";
				break;
			case AscDFH.historydescription_Document_ConvertOldEquation                 :
				sString = "Document_ConvertOldEquation";
				break;
			case AscDFH.historydescription_Document_AddNewStyle                        :
				sString = "Document_AddNewStyle";
				break;
			case AscDFH.historydescription_Document_RemoveStyle                        :
				sString = "Document_RemoveStyle";
				break;
			case AscDFH.historydescription_Document_AddTextArt                         :
				sString = "Document_AddTextArt";
				break;
			case AscDFH.historydescription_Document_RemoveAllCustomStyles              :
				sString = "Document_RemoveAllCustomStyles";
				break;
			case AscDFH.historydescription_Document_AcceptAllRevisionChanges           :
				sString = "Document_AcceptAllRevisionChanges";
				break;
			case AscDFH.historydescription_Document_RejectAllRevisionChanges           :
				sString = "Document_RejectAllRevisionChanges";
				break;
			case AscDFH.historydescription_Document_AcceptRevisionChange               :
				sString = "Document_AcceptRevisionChange";
				break;
			case AscDFH.historydescription_Document_RejectRevisionChange               :
				sString = "Document_RejectRevisionChange";
				break;
			case AscDFH.historydescription_Document_AcceptRevisionChangesBySelection   :
				sString = "Document_AcceptRevisionChangesBySelection";
				break;
			case AscDFH.historydescription_Document_RejectRevisionChangesBySelection   :
				sString = "Document_RejectRevisionChangesBySelection";
				break;
			case AscDFH.historydescription_Document_AddLetterUnion                     :
				sString = "Document_AddLetterUnion";
				break;
			case AscDFH.historydescription_Document_SetColumnsFromRuler                :
				sString = "Document_SetColumnsFromRuler";
				break;
			case AscDFH.historydescription_Document_SetColumnsProps                    :
				sString = "Document_SetColumnsProps";
				break;
			case AscDFH.historydescription_Document_AddColumnBreak                     :
				sString = "Document_AddColumnBreak";
				break;
			case AscDFH.historydescription_Document_AddTabToMath                       :
				sString = "Document_AddTabToMath";
				break;
			case AscDFH.historydescription_Document_ApplyPrToMath:
				sString = "Document_ApplyPrToMath";
				break;
			case AscDFH.historydescription_Document_SetMathProps:
				sString = "Document_SetMathProps";
				break;
			case AscDFH.historydescription_Document_SetSectionProps:
				sString = "Document_SetColumnsProps";
				break;
			case AscDFH.historydescription_Document_ApiBuilder:
				sString = "Document_ApiBuilder";
				break;
			case AscDFH.historydescription_Document_AddOleObject:
				sString = "Document_AddOleObject";
				break;
			case AscDFH.historydescription_Document_EditOleObject:
				sString = "Document_EditOleObject";
				break;
			case AscDFH.historydescription_Document_CompositeInput:
				sString = "Document_CompositeInput";
				break;
			case AscDFH.historydescription_Document_CompositeInputReplace:
				sString = "Document_CompositeInputReplace";
				break;
			case AscDFH.historydescription_Document_AddPageCount:
				sString = "Document_AddPageCount";
				break;
			case AscDFH.historydescription_Document_AddFootnote:
				sString = "Document_AddFootnote";
				break;
			case AscDFH.historydescription_Document_SetFootnotePr:
				sString = "Document_SetFootnotePr";
				break;
			case AscDFH.historydescription_Document_RemoveAllFootnotes:
				sString = "Document_RemoveAllFootnotes";
				break;
			case AscDFH.historydescription_Document_InsertDocumentsByUrls:
				sString = "Document_InsertDocumentsByUrls";
				break;
			case AscDFH.historydescription_Document_AddBlockLevelContentControl:
				sString = "Document_AddBlockLevelContentControl";
				break;
			case AscDFH.historydescription_Document_AddInlineLevelContentControl:
				sString = "Document_AddInlineLevelContentControl";
				break;
			case AscDFH.historydescription_Document_RemoveContentControl:
				sString = "Document_RemoveContentControl";
				break;
			case AscDFH.historydescription_Document_RemoveContentControlWrapper:
				sString = "Document_RemoveContentControlWrapper";
				break;
			case AscDFH.historydescription_Document_ChangeContentControlProperties:
				sString = "Document_ChangeContentControlProperties";
				break;
			case AscDFH.historydescription_DocumentMacros_Data:
				sString = "DocumentMacros_Data";
				break;
			case AscDFH.historydescription_Document_AddBookmark:
				sString = "Document_AddBookmark";
				break;
			case AscDFH.historydescription_Document_AddTableOfContents:
				sString = "Document_AddTableOfContents";
				break;
			case AscDFH.historydescription_Document_ChangeOutlineLevel:
				sString = "Document_ChangeOutlineLevel";
				break;
			case AscDFH.historydescription_Document_AddElementToOutline:
				sString = "Document_AddElementToOutline";
				break;
			case AscDFH.historydescription_Document_ResizeTable:
				sString = "Document_ResizeTable";
				break;
			case AscDFH.historydescription_Document_RemoveComplexField:
				sString = "Document_RemoveComplexField";
				break;
			case AscDFH.historydescription_Document_SetComplexFieldPr:
				sString = "Document_SetComplexFieldPr";
				break;
			case AscDFH.historydescription_Document_UpdateTableOfContents:
				sString = "Document_UpdateTableOfContents";
				break;
			case AscDFH.historydescription_Document_SectionStartPage:
				sString = "Document_SectionStartPage";
				break;
			case AscDFH.historydescription_Document_DistributeTableCells:
				sString = "Document_DistributeTableCells";
				break;
			case AscDFH.historydescription_Document_RemoveBookmark:
				sString = "Document_RemoveBookmark";
				break;
			case AscDFH.historydescription_Document_ContinueNumbering:
				sString = "Document_ContinueNumbering";
				break;
			case AscDFH.historydescription_Document_RestartNumbering:
				sString = "Document_RestartNumbering";
				break;
			case AscDFH.historydescription_Document_AutomaticListAsType:
				sString = "Document_AutomaticListAsType";
				break;
			case AscDFH.historydescription_Document_CreateNum:
				sString = "Document_CreateNum";
				break;
			case AscDFH.historydescription_Document_ChangeNumLvl:
				sString = "Document_ChangeNumLvl";
				break;
			case AscDFH.historydescription_Document_AutoCorrectSmartQuotes:
				sString = "Document_AutoCorrectSmartQuotes";
				break;
			case AscDFH.historydescription_Document_AutoCorrectHyphensWithDash:
				sString = "Document_AutoCorrectHyphensWithDash";
				break;
			case AscDFH.historydescription_Document_SetGlobalSdtHighlightColor:
				sString = "Document_SetGlobalSdtHighlightColor";
				break;
			case AscDFH.historydescription_Document_SetGlobalSdtShowHighlight:
				sString = "Document_SetGlobalSdtShowHighlight";
				break;
			case AscDFH.historydescription_Document_UpdateFields:
				sString = "Document_UpdateFields";
				break;
			case AscDFH.historydescription_Document_AddBlankPage:
				sString = "Document_AddBlankPage";
				break;
			case AscDFH.historydescription_Document_AddTableFormula:
				sString = "Document_AddTableFormula";
				break;
			case AscDFH.historydescription_Document_ChangeTableFormula:
				sString = "Document_ChangeTableFormula";
				break;
			case AscDFH.historydescription_Document_SetParagraphOutlineLvl:
				sString = "Document_SetParagraphOutlineLvl";
				break;
			case AscDFH.historydescription_Document_RemoveTableCells:
				sString = "Document_RemoveTableCells";
				break;
			case AscDFH.historydescription_Document_AddContentControlCheckBox:
				sString = "Document_AddContentControlCheckBox";
				break;
			case AscDFH.historydescription_Document_SetContentControlCheckBoxPr:
				sString = "Document_SetContentControlCheckBoxPr";
				break;
			case AscDFH.historydescription_Document_AddContentControlPicture:
				sString = "Document_AddContentControlPicture";
				break;
			case AscDFH.historydescription_Document_SetContentControlPictureUrl:
				sString = "Document_SetContentControlPictureUrl";
				break;
			case AscDFH.historydescription_Document_RemoveAllComments:
				sString = "Document_RemoveAllComments";
				break;
			case AscDFH.historydescription_Document_AddContentControlList:
				sString = "Document_AddContentControlList";
				break;
			case AscDFH.historydescription_Document_SetContentControlListPr:
				sString = "Document_SetContentControlListPr";
				break;
			case AscDFH.historydescription_Document_SelectContentControlListItem:
				sString = "Document_SelectContentControlListItem";
				break;
			case AscDFH.historydescription_Document_AddContentControlDatePicker:
				sString = "Document_AddContentControlDatePicker";
				break;
			case AscDFH.historydescription_Document_SetContentControlDatePickerPr:
				sString = "Document_SetContentControlDatePickerPr";
				break;
			case AscDFH.historydescription_Document_AddTextWithProperties:
				sString = "Document_AddTextWithProperties";
				break;
			case AscDFH.historydescription_Document_AddCaption:
				sString = "Document_AddCaption";
				break;
			case AscDFH.historydescription_Document_CompareDocuments:
				sString = "Document_CompareDocuments";
				break;
			case AscDFH.historydescription_Document_DrawNewTable:
				sString = "Document_DrawNewTable";
				break;
			case AscDFH.historydescription_Document_DrawTable:
				sString = "Document_DrawTable";
				break;
			case AscDFH.historydescription_Document_AddDateTimeField:
				sString = "Document_AddDateTimeField";
				break;
			case AscDFH.historydescription_Document_SetContentControlTextPlaceholder:
				sString = "Document_SetContentControlTextPlaceholder";
				break;
			case AscDFH.historydescription_Document_AddEndnote:
				sString = "Document_AddEndnote";
				break;
			case AscDFH.historydescription_Document_AddContentControlTextForm:
				sString = "Document_AddContentControlTextForm";
				break;
			case AscDFH.historydescription_Document_SetEndnotePr:
				sString = "Document_SetEndnotePr";
				break;
			case AscDFH.historydescription_Document_ConvertFootnoteType:
				sString = "Document_ConvertFootnoteType";
				break;
			case AscDFH.historydescription_Document_AutoCorrectCommon:
				sString = "Document_AutoCorrectCommon";
				break;
			case AscDFH.historydescription_Document_Shortcut_ClearFormatting:
				sString = "Document_Shortcut_ClearFormatting";
				break;
			case AscDFH.historydescription_Document_Shortcut_AddNonBreakingSpace:
				sString = "Document_Shortcut_AddNonBreakingSpace";
				break;
			case AscDFH.historydescription_Document_SetParagraphSuppressLineNumbers:
				sString = "Document_SetParagraphSuppressLineNumbers";
				break;
			case AscDFH.historydescription_Document_SetLineNumbersProps:
				sString = "Document_SetLineNumbersProps";
				break;
			case AscDFH.historydescription_Document_AddCrossRef:
				sString = "Document_AddCrossRef";
				break;
			case AscDFH.historydescription_Document_ClearAllSpecialForms:
				sString = "Document_ClearAllSpecialForms";
				break;
			case AscDFH.historydescription_Document_ChangeTextCase:
				sString = "Document_ChangeTextCase";
				break;
			case AscDFH.historydescription_Document_SetNumberingLvl:
				sString = "Document_SetNumberingLvl";
				break;
			case AscDFH.historydescription_Document_SetTrackRevisions:
				sString = "Document_SetTrackRevisions";
				break;
			case AscDFH.historydescription_Document_SetContentControlText:
				sString = "Document_SetContentControlText";
				break;
			case AscDFH.historydescription_Document_ClearContentControl:
				sString = "Document_ClearContentControl";
				break;
			case AscDFH.historydescription_Document_AutoCorrectFirstLetterOfSentence:
				sString = "Document_AutoCorrectFirstLetterOfSentence";
				break;
			case AscDFH.historydescription_Document_ConvertTextToTable:
				sString = "Document_ConvertTextToTable";
				break;
			case AscDFH.historydescription_Document_ConvertTableToText:
				sString = "Document_ConvertTableToText";
				break;
			case AscDFH.historydescription_Document_ResolveAllComments:
				sString = "Document_ResolveAllComments";
				break;
			case AscDFH.historydescription_Document_ConvertFormFixedType:
				sString = "Document_ConvertFormFixedType";
				break;
			case AscDFH.historydescription_Document_ReplaceCurrentWord:
				sString = "Document_ReplaceCurrentWord";
				break;
			case AscDFH.historydescription_Document_Docxf_To_Docx:
				sString = "Document_Docxf_To_Docx";
				break;
			case AscDFH.historydescription_Document_ConvertMathView:
				sString = "Document_ConvertMathView";
				break;
			case AscDFH.historydescription_Document_ConvertMathDisplayMode:
				sString = "Document_ConvertMathDisplayMode";
				break;
			case AscDFH.historydescription_Document_RemoveHdrFtr:
				sString = "Document_RemoveHdrFtr";
				break;
			case AscDFH.historydescription_Document_AddParagraphToTOC:
				sString = "Document_AddParagraphToTOC";
				break;
			case AscDFH.historydescription_Document_FillFormsByTags:
				sString = "Document_FillFormsByTags";
				break;
			case AscDFH.historydescription_Document_FillFormInPlugin:
				sString = "Document_FillFormInPlugin";
				break;
			case AscDFH.historydescription_Document_AddComplexForm:
				sString = "Document_AddComplexForm";
				break;
			case AscDFH.historydescription_Document_CorrectFormTextByFormat:
				sString = "Document_CorrectFormTextByFormat";
				break;
			case AscDFH.historydescription_Document_CorrectEnterText:
				sString = "Document_CorrectEnterText";
				break;
			case AscDFH.historydescription_Document_DocumentProtection:
				sString = "Document_DocumentProtection";
				break;
			case AscDFH.historydescription_Collaborative_Undo:
				sString = "Collaborative_Undo";
				break;
			case AscDFH.historydescription_Document_AddPlaceholderImages:
				sString = "Document_AddPlaceholderImages";
				break;
			case AscDFH.historydescription_OForm_AddRole:
				sString = "OForm_AddRole";
				break;
			case AscDFH.historydescription_OForm_RemoveRole:
				sString = "OForm_RemoveRole";
				break;
			case AscDFH.historydescription_OForm_EditRole:
				sString = "OForm_EditRole";
				break;
			case AscDFH.historydescription_OForm_ChangeRoleOrder:
				sString = "OForm_ChangeRoleOrder";
				break;
			case AscDFH.historydescription_Document_FillContentControlPlaceholderOnBlur:
				sString = "Document_FillContentControlPlaceholderOnBlur";
				break;
			case AscDFH.historydescription_Document_SetAutoHyphenation:
				sString = "Document_SetAutoHyphenation";
				break;
			case AscDFH.historydescription_Document_SetConsecutiveHyphenLimit:
				sString = "Document_SetConsecutiveHyphenLimit";
				break;
			case AscDFH.historydescription_Document_SetHyphenateCaps:
				sString = "Document_SetHyphenateCaps";
				break;
			case AscDFH.historydescription_Document_RemoveMathShortcut:
				sString = "Document_RemoveMathShortcut";
				break;
			case AscDFH.historydescription_Document_SetAllFormsData:
				sString = "Document_SetAllFormsData";
				break;
			case AscDFH.historydescription_Document_ComplexField_MergeFormat:
				sString = "Document_ComplexField_MergeFormat";
				break;
			case AscDFH.historydescription_Presentation_ResetSlideBackground:
				sString = "Presentation_ResetSlideBackground";
				break;
			case AscDFH.historydescription_Presentation_ApplyBackgroundToAll:
				sString = "Presentation_ApplyBackgroundToAll";
				break;
			case AscDFH.historydescription_Presentation_ShowMasterShapes:
				sString = "Presentation_ShowMasterShapes";
				break;
			case AscDFH.historydescription_BuilderScript:
				sString = "BuilderScript";
				break;
			case AscDFH.historydescription_Document_AddRemoveBeforeAfterParagraph:
				sString = "Document_AddRemoveBeforeAfterParagraph";
				break;
			case AscDFH.historydescription_Document_SectionPageNumFormat:
				sString = "Document_SectionPageNumFormat";
				break;
			case AscDFH.historydescription_Document_SetPageColor:
				sString = "Document_SetPageColor";
				break;
			case AscDFH.historydescription_Document_InsertTextFromFile:
				sString = "Document_InsertTextFromFile";
				break;
			case AscDFH.historydescription_Document_AddComplexField:
				sString = "Document_AddComplexField";
				break;
			case AscDFH.historydescription_Document_EditComplexFieldInstruction:
				sString = "Document_EditComplexFieldInstruction";
				break;
			case AscDFH.historydescription_Collaborative_DeletedTextRecovery:
				sString = "Collaborative_DeletedTextRecovery";
				break;
			case AscDFH.historydescription_Presentation_MergeSelectedShapes:
				sString = "Presentation_MergeSelectedShapes";
				break;
		}
		return sString;
	}
	function GetHistoryClassTypeByChangeType(nChangeType)
	{
		return (nChangeType >> 16) & 0x0000FFFF;
	}
	function GetChange(base64Binary)
	{
		let change = new AscCommon.CCollaborativeChanges();
		change.Set_Data(base64Binary);
		return change.ToHistoryChange();
	}

	//------------------------------------------------------------export--------------------------------------------------
	window['AscDFH']                                  = window['AscDFH'] || {};
	window['AscDFH'].GetHistoryPointStringDescription = GetHistoryPointStringDescription;
	window['AscDFH'].GetHistoryClassTypeByChangeType  = GetHistoryClassTypeByChangeType;
	window['AscDFH'].GetChange                        = GetChange;
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//
	// Типы произошедших изменений
	//
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	window['AscDFH'].historyitem_recalctype_Inline    = 0; // Изменения произошли в обычном тексте (с верхним классом CDocument)
	window['AscDFH'].historyitem_recalctype_Flow      = 1; // Изменения произошли в "плавающем" объекте
	window['AscDFH'].historyitem_recalctype_HdrFtr    = 2; // Изменения произошли в колонтитуле
	window['AscDFH'].historyitem_recalctype_Drawing   = 3; // Изменения произошли в drawing'е
	window['AscDFH'].historyitem_recalctype_NotesEnd  = 4; // Изменение произошли в сносках, которые идут в конце документа
	window['AscDFH'].historyitem_recalctype_FromStart = 0xFFFF; // Изменения требуют полного пересчета документа с самого начала
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//
	// Типы классов, в которых происходили изменения (типы нужны для совместного редактирования)
	//
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	window['AscDFH'].historyitem_type_Unknown            = 0 << 16;
	window['AscDFH'].historyitem_type_TableId            = 1 << 16;
	window['AscDFH'].historyitem_type_Document           = 2 << 16;
	window['AscDFH'].historyitem_type_Paragraph          = 3 << 16;
	window['AscDFH'].historyitem_type_TextPr             = 4 << 16;
	window['AscDFH'].historyitem_type_Drawing            = 5 << 16;
	window['AscDFH'].historyitem_type_DrawingObjects     = 6 << 16; // obsolete
	window['AscDFH'].historyitem_type_FlowObjects        = 7 << 16; // obsolete
	window['AscDFH'].historyitem_type_FlowImage          = 8 << 16; // obsolete
	window['AscDFH'].historyitem_type_Table              = 9 << 16;
	window['AscDFH'].historyitem_type_TableRow           = 10 << 16;
	window['AscDFH'].historyitem_type_TableCell          = 11 << 16;
	window['AscDFH'].historyitem_type_DocumentContent    = 12 << 16;
	window['AscDFH'].historyitem_type_FlowTable          = 13 << 16; // obsolete
	window['AscDFH'].historyitem_type_HdrFtrController   = 14 << 16; // obsolete
	window['AscDFH'].historyitem_type_HdrFtr             = 15 << 16;
	window['AscDFH'].historyitem_type_AbstractNum        = 16 << 16;
	window['AscDFH'].historyitem_type_Comment            = 17 << 16;
	window['AscDFH'].historyitem_type_Comments           = 18 << 16;
	window['AscDFH'].historyitem_type_Image              = 19 << 16;
	window['AscDFH'].historyitem_type_GrObjects          = 20 << 16;
	window['AscDFH'].historyitem_type_Hyperlink          = 21 << 16;
	window['AscDFH'].historyitem_type_Style              = 23 << 16;
	window['AscDFH'].historyitem_type_Styles             = 24 << 16;
	window['AscDFH'].historyitem_type_ChartTitle         = 25 << 16;
	window['AscDFH'].historyitem_type_Math               = 26 << 16;
	window['AscDFH'].historyitem_type_CommentMark        = 27 << 16;
	window['AscDFH'].historyitem_type_ParaRun            = 28 << 16;
	window['AscDFH'].historyitem_type_MathContent        = 29 << 16;
	window['AscDFH'].historyitem_type_Section            = 30 << 16;
	window['AscDFH'].historyitem_type_acc                = 31 << 16;
	window['AscDFH'].historyitem_type_bar                = 32 << 16;
	window['AscDFH'].historyitem_type_borderBox          = 33 << 16;
	window['AscDFH'].historyitem_type_box                = 34 << 16;
	window['AscDFH'].historyitem_type_delimiter          = 35 << 16;
	window['AscDFH'].historyitem_type_eqArr              = 36 << 16;
	window['AscDFH'].historyitem_type_frac               = 37 << 16;
	window['AscDFH'].historyitem_type_mathFunc           = 38 << 16;
	window['AscDFH'].historyitem_type_groupChr           = 39 << 16;
	window['AscDFH'].historyitem_type_lim                = 40 << 16;
	window['AscDFH'].historyitem_type_matrix             = 41 << 16;
	window['AscDFH'].historyitem_type_nary               = 42 << 16;
	window['AscDFH'].historyitem_type_integral           = 43 << 16;
	window['AscDFH'].historyitem_type_double_integral    = 44 << 16;
	window['AscDFH'].historyitem_type_triple_integral    = 45 << 16;
	window['AscDFH'].historyitem_type_contour_integral   = 46 << 16;
	window['AscDFH'].historyitem_type_surface_integral   = 47 << 16;
	window['AscDFH'].historyitem_type_volume_integral    = 48 << 16;
	window['AscDFH'].historyitem_type_phant              = 49 << 16;
	window['AscDFH'].historyitem_type_rad                = 50 << 16;
	window['AscDFH'].historyitem_type_deg_subsup         = 51 << 16;
	window['AscDFH'].historyitem_type_iterators          = 52 << 16;
	window['AscDFH'].historyitem_type_deg                = 53 << 16;
	window['AscDFH'].historyitem_type_ParaComment        = 54 << 16;
	window['AscDFH'].historyitem_type_Field              = 55 << 16;
	window['AscDFH'].historyitem_type_Footnotes          = 56 << 16;
	window['AscDFH'].historyitem_type_FootEndNote        = 57 << 16;
	window['AscDFH'].historyitem_type_Presentation       = 58 << 16;
	window['AscDFH'].historyitem_type_BlockLevelSdt      = 59 << 16;
	window['AscDFH'].historyitem_type_SdtPr              = 60 << 16;
	window['AscDFH'].historyitem_type_InlineLevelSdt     = 61 << 16;
	window['AscDFH'].historyitem_type_ParaBookmark       = 62 << 16;
	window['AscDFH'].historyitem_type_Num                = 63 << 16;
	window['AscDFH'].historyitem_type_PresentationField  = 64 << 16;
	window['AscDFH'].historyitem_type_ParaRevisionMove   = 65 << 16;
	window['AscDFH'].historyitem_type_RunRevisionMove    = 66 << 16;
	window['AscDFH'].historyitem_type_GlossaryDocument   = 67 << 16;
	window['AscDFH'].historyitem_type_DocPart            = 68 << 16;
	window['AscDFH'].historyitem_type_Endnotes           = 69 << 16;
	window['AscDFH'].historyitem_type_ParagraphPermStart = 70 << 16;
	window['AscDFH'].historyitem_type_ParagraphPermEnd   = 71 << 16;
	window['AscDFH'].historyitem_type_CustomXmlManager   = 72 << 16;
	window['AscDFH'].historyitem_type_CustomXml          = 73 << 16;

	window['AscDFH'].historyitem_type_CommonShape            = 1000 << 16; // Этот класс добавлен для элементов, у которых нет конкретного класса

	window['AscDFH'].historyitem_type_ColorMod               = 1001 << 16;
	window['AscDFH'].historyitem_type_ColorModifiers         = 1002 << 16;
	window['AscDFH'].historyitem_type_SysColor               = 1003 << 16;
	window['AscDFH'].historyitem_type_PrstColor              = 1004 << 16;
	window['AscDFH'].historyitem_type_RGBColor               = 1005 << 16;
	window['AscDFH'].historyitem_type_SchemeColor            = 1006 << 16;
	window['AscDFH'].historyitem_type_UniColor               = 1007 << 16;
	window['AscDFH'].historyitem_type_SrcRect                = 1008 << 16;
	window['AscDFH'].historyitem_type_BlipFill               = 1009 << 16;
	window['AscDFH'].historyitem_type_SolidFill              = 1010 << 16;
	window['AscDFH'].historyitem_type_Gs                     = 1011 << 16;
	window['AscDFH'].historyitem_type_GradLin                = 1012 << 16;
	window['AscDFH'].historyitem_type_GradPath               = 1013 << 16;
	window['AscDFH'].historyitem_type_GradFill               = 1014 << 16;
	window['AscDFH'].historyitem_type_PathFill               = 1015 << 16;
	window['AscDFH'].historyitem_type_NoFill                 = 1016 << 16;
	window['AscDFH'].historyitem_type_UniFill                = 1017 << 16;
	window['AscDFH'].historyitem_type_EndArrow               = 1018 << 16;
	window['AscDFH'].historyitem_type_LineJoin               = 1019 << 16;
	window['AscDFH'].historyitem_type_Ln                     = 1020 << 16;
	window['AscDFH'].historyitem_type_DefaultShapeDefinition = 1021 << 16;
	window['AscDFH'].historyitem_type_CNvPr                  = 1022 << 16;
	window['AscDFH'].historyitem_type_NvPr                   = 1023 << 16;
	window['AscDFH'].historyitem_type_Ph                     = 1024 << 16;
	window['AscDFH'].historyitem_type_UniNvPr                = 1025 << 16;
	window['AscDFH'].historyitem_type_StyleRef               = 1026 << 16;
	window['AscDFH'].historyitem_type_FontRef                = 1027 << 16;
	window['AscDFH'].historyitem_type_Chart                  = 1028 << 16;
	window['AscDFH'].historyitem_type_ChartSpace             = 1029 << 16;
	window['AscDFH'].historyitem_type_Legend                 = 1030 << 16;
	window['AscDFH'].historyitem_type_Layout                 = 1031 << 16;
	window['AscDFH'].historyitem_type_LegendEntry            = 1032 << 16;
	window['AscDFH'].historyitem_type_PivotFmt               = 1033 << 16;
	window['AscDFH'].historyitem_type_DLbl                   = 1034 << 16;
	window['AscDFH'].historyitem_type_Marker                 = 1035 << 16;
	window['AscDFH'].historyitem_type_PlotArea               = 1036 << 16;
	window['AscDFH'].historyitem_type_Axis                   = 1037 << 16;
	window['AscDFH'].historyitem_type_NumFmt                 = 1038 << 16;
	window['AscDFH'].historyitem_type_Scaling                = 1039 << 16;
	window['AscDFH'].historyitem_type_DTable                 = 1040 << 16;
	window['AscDFH'].historyitem_type_LineChart              = 1041 << 16;
	window['AscDFH'].historyitem_type_DLbls                  = 1042 << 16;
	window['AscDFH'].historyitem_type_UpDownBars             = 1043 << 16;
	window['AscDFH'].historyitem_type_BarChart               = 1044 << 16;
	window['AscDFH'].historyitem_type_BubbleChart            = 1045 << 16;
	window['AscDFH'].historyitem_type_DoughnutChart          = 1046 << 16;
	window['AscDFH'].historyitem_type_OfPieChart             = 1047 << 16;
	window['AscDFH'].historyitem_type_PieChart               = 1048 << 16;
	window['AscDFH'].historyitem_type_RadarChart             = 1049 << 16;
	window['AscDFH'].historyitem_type_ScatterChart           = 1050 << 16;
	window['AscDFH'].historyitem_type_StockChart             = 1051 << 16;
	window['AscDFH'].historyitem_type_SurfaceChart           = 1052 << 16;
	window['AscDFH'].historyitem_type_BandFmt                = 1053 << 16;
	window['AscDFH'].historyitem_type_AreaChart              = 1054 << 16;
	window['AscDFH'].historyitem_type_ScatterSer             = 1055 << 16;
	window['AscDFH'].historyitem_type_DPt                    = 1056 << 16;
	window['AscDFH'].historyitem_type_ErrBars                = 1057 << 16;
	window['AscDFH'].historyitem_type_MinusPlus              = 1058 << 16;
	window['AscDFH'].historyitem_type_NumLit                 = 1059 << 16;
	window['AscDFH'].historyitem_type_NumericPoint           = 1060 << 16;
	window['AscDFH'].historyitem_type_NumRef                 = 1061 << 16;
	window['AscDFH'].historyitem_type_TrendLine              = 1062 << 16;
	window['AscDFH'].historyitem_type_Tx                     = 1063 << 16;
	window['AscDFH'].historyitem_type_StrRef                 = 1064 << 16;
	window['AscDFH'].historyitem_type_StrCache               = 1065 << 16;
	window['AscDFH'].historyitem_type_StrPoint               = 1066 << 16;
	window['AscDFH'].historyitem_type_XVal                   = 1067 << 16;
	window['AscDFH'].historyitem_type_MultiLvlStrRef         = 1068 << 16;
	window['AscDFH'].historyitem_type_MultiLvlStrCache       = 1069 << 16;
	window['AscDFH'].historyitem_type_StringLiteral          = 1070 << 16;
	window['AscDFH'].historyitem_type_YVal                   = 1071 << 16;
	window['AscDFH'].historyitem_type_AreaSeries             = 1072 << 16;
	window['AscDFH'].historyitem_type_Cat                    = 1073 << 16;
	window['AscDFH'].historyitem_type_PictureOptions         = 1074 << 16;
	window['AscDFH'].historyitem_type_RadarSeries            = 1075 << 16;
	window['AscDFH'].historyitem_type_BarSeries              = 1076 << 16;
	window['AscDFH'].historyitem_type_LineSeries             = 1077 << 16;
	window['AscDFH'].historyitem_type_PieSeries              = 1078 << 16;
	window['AscDFH'].historyitem_type_SurfaceSeries          = 1079 << 16;
	window['AscDFH'].historyitem_type_BubbleSeries           = 1080 << 16;
	window['AscDFH'].historyitem_type_ExternalData           = 1081 << 16;
	window['AscDFH'].historyitem_type_PivotSource            = 1082 << 16;
	window['AscDFH'].historyitem_type_Protection             = 1083 << 16;
	window['AscDFH'].historyitem_type_ChartWall              = 1084 << 16;
	window['AscDFH'].historyitem_type_View3d                 = 1085 << 16;
	window['AscDFH'].historyitem_type_ChartText              = 1086 << 16;
	window['AscDFH'].historyitem_type_ShapeStyle             = 1087 << 16;
	window['AscDFH'].historyitem_type_Xfrm                   = 1088 << 16;
	window['AscDFH'].historyitem_type_SpPr                   = 1089 << 16;
	window['AscDFH'].historyitem_type_ClrScheme              = 1090 << 16;
	window['AscDFH'].historyitem_type_ClrMap                 = 1091 << 16;
	window['AscDFH'].historyitem_type_ExtraClrScheme         = 1092 << 16;
	window['AscDFH'].historyitem_type_FontCollection         = 1093 << 16;
	window['AscDFH'].historyitem_type_FontScheme             = 1094 << 16;
	window['AscDFH'].historyitem_type_FormatScheme           = 1095 << 16;
	window['AscDFH'].historyitem_type_ThemeElements          = 1096 << 16;
	window['AscDFH'].historyitem_type_HF                     = 1097 << 16;
	window['AscDFH'].historyitem_type_BgPr                   = 1098 << 16;
	window['AscDFH'].historyitem_type_Bg                     = 1099 << 16;
	window['AscDFH'].historyitem_type_PrintSettings          = 1100 << 16;
	window['AscDFH'].historyitem_type_HeaderFooterChart      = 1101 << 16;
	window['AscDFH'].historyitem_type_PageMarginsChart       = 1102 << 16;
	window['AscDFH'].historyitem_type_PageSetup              = 1103 << 16;
	window['AscDFH'].historyitem_type_Shape                  = 1104 << 16;
	window['AscDFH'].historyitem_type_DispUnits              = 1105 << 16;
	window['AscDFH'].historyitem_type_GroupShape             = 1106 << 16;
	window['AscDFH'].historyitem_type_ImageShape             = 1107 << 16;
	window['AscDFH'].historyitem_type_Geometry               = 1108 << 16;
	window['AscDFH'].historyitem_type_Path                   = 1109 << 16;
	window['AscDFH'].historyitem_type_TextBody               = 1110 << 16;
	window['AscDFH'].historyitem_type_CatAx                  = 1111 << 16;
	window['AscDFH'].historyitem_type_ValAx                  = 1112 << 16;
	window['AscDFH'].historyitem_type_WrapPolygon            = 1113 << 16;
	window['AscDFH'].historyitem_type_DateAx                 = 1114 << 16;
	window['AscDFH'].historyitem_type_SerAx                  = 1115 << 16;
	window['AscDFH'].historyitem_type_Title                  = 1116 << 16;
	window['AscDFH'].historyitem_type_Slide                  = 1117 << 16;
	window['AscDFH'].historyitem_type_SlideLayout            = 1118 << 16;
	window['AscDFH'].historyitem_type_SlideMaster            = 1119 << 16;
	window['AscDFH'].historyitem_type_SlideComments          = 1120 << 16;
	window['AscDFH'].historyitem_type_PropLocker             = 1121 << 16;
	window['AscDFH'].historyitem_type_Theme                  = 1122 << 16;
	window['AscDFH'].historyitem_type_GraphicFrame           = 1123 << 16;
	window['AscDFH'].historyitem_type_GrpFill                = 1124 << 16;
	window['AscDFH'].historyitem_type_OleObject              = 1125 << 16;
	window['AscDFH'].historyitem_type_DrawingContent         = 1126 << 16;
	window['AscDFH'].historyitem_type_Sparkline              = 1127 << 16;
	window['AscDFH'].historyitem_type_NotesMaster            = 1128 << 16;
	window['AscDFH'].historyitem_type_Notes                  = 1129 << 16;
	window['AscDFH'].historyitem_type_Cnx                    = 1130 << 16;
	window['AscDFH'].historyitem_type_PresentationSection    = 1131 << 16;
	window['AscDFH'].historyitem_type_PivotTableDefinition   = 1132 << 16;
	window['AscDFH'].historyitem_type_LockedCanvas           = 1133 << 16;
	window['AscDFH'].historyitem_type_RelSizeAnchor          = 1134 << 16;
	window['AscDFH'].historyitem_type_AbsSizeAnchor          = 1135 << 16;
	window['AscDFH'].historyitem_type_Core                   = 1136 << 16;
	window['AscDFH'].historyitem_type_SlicerView             = 1137 << 16;
	window['AscDFH'].historyitem_type_PivotWorksheetSource   = 1138 << 16;
	window['AscDFH'].historyitem_type_NamedSheetView         = 1139 << 16;
	window['AscDFH'].historyitem_type_DataValidation         = 1140 << 16;
	window['AscDFH'].historyitem_type_EmptyObject            = 1141 << 16;
	window['AscDFH'].historyitem_type_Timing                 = 1142 << 16;
	window['AscDFH'].historyitem_type_CommonTimingList       = 1143 << 16;
	window['AscDFH'].historyitem_type_ObjectTarget           = 1144 << 16;
	window['AscDFH'].historyitem_type_BldBase                = 1145 << 16;
	window['AscDFH'].historyitem_type_BldDgm                 = 1146 << 16;
	window['AscDFH'].historyitem_type_BldGraphic             = 1147 << 16;
	window['AscDFH'].historyitem_type_BldOleChart            = 1148 << 16;
	window['AscDFH'].historyitem_type_BldP                   = 1149 << 16;
	window['AscDFH'].historyitem_type_BldSub                 = 1150 << 16;
	window['AscDFH'].historyitem_type_DirTransition          = 1151 << 16;
	window['AscDFH'].historyitem_type_OptBlackTransition     = 1152 << 16;
	window['AscDFH'].historyitem_type_GraphicEl              = 1153 << 16;
	window['AscDFH'].historyitem_type_IndexRg                = 1154 << 16;
	window['AscDFH'].historyitem_type_Tmpl                   = 1155 << 16;
	window['AscDFH'].historyitem_type_Anim                   = 1156 << 16;
	window['AscDFH'].historyitem_type_CBhvr                  = 1157 << 16;
	window['AscDFH'].historyitem_type_CTn                    = 1158 << 16;
	window['AscDFH'].historyitem_type_Cond                   = 1159 << 16;
	window['AscDFH'].historyitem_type_Rtn                    = 1160 << 16;
	window['AscDFH'].historyitem_type_TgtEl                  = 1161 << 16;
	window['AscDFH'].historyitem_type_SndTgt                 = 1162 << 16;
	window['AscDFH'].historyitem_type_SpTgt                  = 1163 << 16;
	window['AscDFH'].historyitem_type_IterateData            = 1164 << 16;
	window['AscDFH'].historyitem_type_Tm                     = 1165 << 16;
	window['AscDFH'].historyitem_type_Tav                    = 1166 << 16;
	window['AscDFH'].historyitem_type_AnimVariant            = 1167 << 16;
	window['AscDFH'].historyitem_type_AnimClr                = 1168 << 16;
	window['AscDFH'].historyitem_type_AnimEffect             = 1169 << 16;
	window['AscDFH'].historyitem_type_AnimMotion             = 1170 << 16;
	window['AscDFH'].historyitem_type_AnimRot                = 1171 << 16;
	window['AscDFH'].historyitem_type_AnimScale              = 1172 << 16;
	window['AscDFH'].historyitem_type_Audio                  = 1173 << 16;
	window['AscDFH'].historyitem_type_CMediaNode             = 1174 << 16;
	window['AscDFH'].historyitem_type_Cmd                    = 1175 << 16;
	window['AscDFH'].historyitem_type_TimeNodeContainer      = 1176 << 16;
	window['AscDFH'].historyitem_type_Seq                    = 1177 << 16;
	window['AscDFH'].historyitem_type_Set                    = 1178 << 16;
	window['AscDFH'].historyitem_type_Video                  = 1179 << 16;
	window['AscDFH'].historyitem_type_OleChartEl             = 1180 << 16;
	window['AscDFH'].historyitem_type_TlPoint                = 1181 << 16;
	window['AscDFH'].historyitem_type_SndAc                  = 1182 << 16;
	window['AscDFH'].historyitem_type_StSnd                  = 1183 << 16;
	window['AscDFH'].historyitem_type_TxEl                   = 1184 << 16;
	window['AscDFH'].historyitem_type_Wheel                  = 1185 << 16;
	window['AscDFH'].historyitem_type_AttrName               = 1186 << 16;
	window['AscDFH'].historyitem_type_AttrNameLst            = 1187 << 16;
	window['AscDFH'].historyitem_type_BldLst                 = 1188 << 16;
	window['AscDFH'].historyitem_type_CondLst                = 1189 << 16;
	window['AscDFH'].historyitem_type_ChildTnLst             = 1190 << 16;
	window['AscDFH'].historyitem_type_Par                    = 1191 << 16;
	window['AscDFH'].historyitem_type_Excl                   = 1192 << 16;
	window['AscDFH'].historyitem_type_TmplLst                = 1193 << 16;
	window['AscDFH'].historyitem_type_TnLst                  = 1194 << 16;
	window['AscDFH'].historyitem_type_TavLst                 = 1195 << 16;
	window['AscDFH'].historyitem_type_SldSz                  = 1196 << 16;
	window['AscDFH'].historyitem_type_ChartStyle             = 1197 << 16;
	window['AscDFH'].historyitem_type_ChartStyleEntry        = 1198 << 16;
	window['AscDFH'].historyitem_type_MarkerLayout           = 1199 << 16;
	window['AscDFH'].historyitem_type_TimelineSlicerView     = 1200 << 16;
	window['AscDFH'].historyitem_type_ImageBlipFillPart      = 1201 << 16;
	window['AscDFH'].historyitem_type_ImageBlipStart         = 1202 << 16;
	window['AscDFH'].historyitem_type_ImageBlipEnd           = 1203 << 16;


	window['AscDFH'].historyitem_type_Address                          = 1201 << 16;
	window['AscDFH'].historyitem_type_AxisUnits                        = 1202 << 16;
	window['AscDFH'].historyitem_type_AxisUnitsLabel                   = 1200 << 16;
	window['AscDFH'].historyitem_type_Binning                          = 1203 << 16;
	window['AscDFH'].historyitem_type_CategoryAxisScaling              = 1204 << 16;
	window['AscDFH'].historyitem_type_ChartData                        = 1205 << 16;
	window['AscDFH'].historyitem_type_Clear                            = 1206 << 16;
	window['AscDFH'].historyitem_type_Copyrights                       = 1207 << 16;
	window['AscDFH'].historyitem_type_Data                             = 1208 << 16;
	window['AscDFH'].historyitem_type_DataLabel                        = 1209 << 16;
	window['AscDFH'].historyitem_type_DataLabelHidden                  = 1200 << 16;
	window['AscDFH'].historyitem_type_DataLabels                       = 1210 << 16;
	window['AscDFH'].historyitem_type_DataLabelVisibilities            = 1211 << 16;
	window['AscDFH'].historyitem_type_DataPoint                        = 1212 << 16;
	window['AscDFH'].historyitem_type_FormatOverride                   = 1213 << 16;
	window['AscDFH'].historyitem_type_FormatOverrides                  = 1214 << 16;
	window['AscDFH'].historyitem_type_Formula                          = 1215 << 16;
	window['AscDFH'].historyitem_type_GeoCache                         = 1216 << 16;
	window['AscDFH'].historyitem_type_GeoChildEntities                 = 1217 << 16;
	window['AscDFH'].historyitem_type_GeoChildEntitiesQuery            = 1218 << 16;
	window['AscDFH'].historyitem_type_GeoChildEntitiesQueryResult      = 1219 << 16;
	window['AscDFH'].historyitem_type_GeoChildEntitiesQueryResults     = 1220 << 16;
	window['AscDFH'].historyitem_type_GeoChildTypes                    = 1221 << 16;
	window['AscDFH'].historyitem_type_GeoData                          = 1222 << 16;
	window['AscDFH'].historyitem_type_GeoDataEntityQuery               = 1223 << 16;
	window['AscDFH'].historyitem_type_GeoDataEntityQueryResult         = 1224 << 16;
	window['AscDFH'].historyitem_type_GeoDataEntityQueryResults        = 1225 << 16;
	window['AscDFH'].historyitem_type_GeoDataPointQuery                = 1226 << 16;
	window['AscDFH'].historyitem_type_GeoDataPointToEntityQuery        = 1227 << 16;
	window['AscDFH'].historyitem_type_GeoDataPointToEntityQueryResult  = 1228 << 16;
	window['AscDFH'].historyitem_type_GeoDataPointToEntityQueryResults = 1229 << 16;
	window['AscDFH'].historyitem_type_Geography                        = 1230 << 16;
	window['AscDFH'].historyitem_type_GeoHierarchyEntity               = 1231 << 16;
	window['AscDFH'].historyitem_type_GeoLocation                      = 1232 << 16;
	window['AscDFH'].historyitem_type_GeoLocationQuery                 = 1233 << 16;
	window['AscDFH'].historyitem_type_GeoLocationQueryResult           = 1234 << 16;
	window['AscDFH'].historyitem_type_GeoLocationQueryResults          = 1235 << 16;
	window['AscDFH'].historyitem_type_GeoLocations                     = 1236 << 16;
	window['AscDFH'].historyitem_type_GeoPolygon                       = 1237 << 16;
	window['AscDFH'].historyitem_type_GeoPolygons                      = 1238 << 16;
	window['AscDFH'].historyitem_type_Gridlines                        = 1239 << 16;
	window['AscDFH'].historyitem_type_Dimension                        = 1240 << 16;
	window['AscDFH'].historyitem_type_NumericDimension                 = 1241 << 16;
	window['AscDFH'].historyitem_type_PercentageColorPosition          = 1242 << 16;
	window['AscDFH'].historyitem_type_PlotAreaRegion                   = 1243 << 16;
	window['AscDFH'].historyitem_type_PlotSurface                      = 1244 << 16;
	window['AscDFH'].historyitem_type_Series                           = 1245 << 16;
	window['AscDFH'].historyitem_type_SeriesElementVisibilities        = 1246 << 16;
	window['AscDFH'].historyitem_type_SeriesLayoutProperties           = 1247 << 16;
	window['AscDFH'].historyitem_type_Statistics                       = 1248 << 16;
	window['AscDFH'].historyitem_type_StringDimension                  = 1249 << 16;
	window['AscDFH'].historyitem_type_Subtotals                        = 1250 << 16;
	window['AscDFH'].historyitem_type_TextData                         = 1251 << 16;
	window['AscDFH'].historyitem_type_TickMarks                        = 1252 << 16;
	window['AscDFH'].historyitem_type_ValueAxisScaling                 = 1253 << 16;
	window['AscDFH'].historyitem_type_ValueColorEndPosition            = 1254 << 16;
	window['AscDFH'].historyitem_type_ValueColorMiddlePosition         = 1255 << 16;
	window['AscDFH'].historyitem_type_ValueColorPositions              = 1256 << 16;
	window['AscDFH'].historyitem_type_ValueColors                      = 1257 << 16;


	window['AscDFH'].historyitem_type_DocumentMacros         = 2000 << 16;
	window['AscDFH'].historyitem_type_PrSet                  = 2001 << 16;
	window['AscDFH'].historyitem_type_CCommonDataList        = 2002 << 16;
	window['AscDFH'].historyitem_type_Point                  = 2003 << 16;
	window['AscDFH'].historyitem_type_PtLst                  = 2004 << 16;
	window['AscDFH'].historyitem_type_DataModel              = 2005 << 16;
	window['AscDFH'].historyitem_type_CxnLst                 = 2006 << 16;
	window['AscDFH'].historyitem_type_BgFormat               = 2008 << 16;
	window['AscDFH'].historyitem_type_Whole                  = 2009 << 16;
	window['AscDFH'].historyitem_type_Cxn                    = 2010 << 16;
	window['AscDFH'].historyitem_type_LayoutDef              = 2012 << 16;
	window['AscDFH'].historyitem_type_CatLst                 = 2013 << 16;
	window['AscDFH'].historyitem_type_SCat                   = 2014 << 16;
	window['AscDFH'].historyitem_type_LayoutNode             = 2017 << 16;
	window['AscDFH'].historyitem_type_Alg                    = 2018 << 16;
	window['AscDFH'].historyitem_type_Param                  = 2019 << 16;
	window['AscDFH'].historyitem_type_Choose                 = 2020 << 16;
	window['AscDFH'].historyitem_type_IteratorAttributes     = 2022 << 16;
	window['AscDFH'].historyitem_type_Else                   = 2023 << 16;
	window['AscDFH'].historyitem_type_AxisType               = 2024 << 16;
	window['AscDFH'].historyitem_type_If                     = 2025 << 16;
	window['AscDFH'].historyitem_type_ElementType            = 2026 << 16;
	window['AscDFH'].historyitem_type_ConstrLst              = 2027 << 16;
	window['AscDFH'].historyitem_type_Constr                 = 2028 << 16;
	window['AscDFH'].historyitem_type_PresOf                 = 2029 << 16;
	window['AscDFH'].historyitem_type_RuleLst                = 2030 << 16;
	window['AscDFH'].historyitem_type_Rule                   = 2031 << 16;
	window['AscDFH'].historyitem_type_SShape                 = 2032 << 16;
	window['AscDFH'].historyitem_type_AdjLst                 = 2033 << 16;
	window['AscDFH'].historyitem_type_Adj                    = 2034 << 16;
	window['AscDFH'].historyitem_type_VarLst                 = 2035 << 16;
	window['AscDFH'].historyitem_type_DiagramTitle           = 2042 << 16;
	window['AscDFH'].historyitem_type_ColorsDef              = 2047 << 16;
	window['AscDFH'].historyitem_type_ColorDefStyleLbl       = 2048 << 16;
	window['AscDFH'].historyitem_type_ClrLst                 = 2049 << 16;
	window['AscDFH'].historyitem_type_StyleDef               = 2059 << 16;
	window['AscDFH'].historyitem_type_Scene3d                = 2060 << 16;
	window['AscDFH'].historyitem_type_StyleDefStyleLbl       = 2061 << 16;
	window['AscDFH'].historyitem_type_Backdrop               = 2063 << 16;
	window['AscDFH'].historyitem_type_BackdropNorm           = 2064 << 16;
	window['AscDFH'].historyitem_type_BackdropUp             = 2065 << 16;
	window['AscDFH'].historyitem_type_Camera                 = 2066 << 16;
	window['AscDFH'].historyitem_type_Rot                    = 2067 << 16;
	window['AscDFH'].historyitem_type_LightRig               = 2068 << 16;
	window['AscDFH'].historyitem_type_Sp3d                   = 2069 << 16;
	window['AscDFH'].historyitem_type_Bevel                  = 2070 << 16;
	window['AscDFH'].historyitem_type_BackdropAnchor         = 2077 << 16;
	window['AscDFH'].historyitem_type_SampData               = 2079 << 16;
	window['AscDFH'].historyitem_type_ForEach                = 2080 << 16;
	window['AscDFH'].historyitem_type_ParameterVal           = 2084 << 16;
	window['AscDFH'].historyitem_type_Coordinate             = 2085 << 16;
	window['AscDFH'].historyitem_type_SmartArt               = 2088 << 16;
	window['AscDFH'].historyitem_type_SmartArtDrawing        = 2091 << 16;
	window['AscDFH'].historyitem_type_DiagramData            = 2092 << 16;
	window['AscDFH'].historyitem_type_CCommonDataListNoId    = 2093 << 16;

	window['AscDFH'].historyitem_type_VMLArc                 = 2099 << 16;
	window['AscDFH'].historyitem_type_VMLCurve               = 2100 << 16;
	window['AscDFH'].historyitem_type_VMLGroup               = 2101 << 16;
	window['AscDFH'].historyitem_type_VMLImage               = 2102 << 16;
	window['AscDFH'].historyitem_type_VMLLine                = 2103 << 16;
	window['AscDFH'].historyitem_type_VMLOval                = 2104 << 16;
	window['AscDFH'].historyitem_type_VMLPolyLine            = 2105 << 16;
	window['AscDFH'].historyitem_type_VMLRect                = 2106 << 16;
	window['AscDFH'].historyitem_type_VMLRoundRect           = 2107 << 16;
	window['AscDFH'].historyitem_type_VMLShape               = 2108 << 16;
	window['AscDFH'].historyitem_type_VMLShapeType           = 2109 << 16;
	window['AscDFH'].historyitem_type_VMLClientData          = 2110 << 16;

	window['AscDFH'].historyitem_type_OleSizeSelection       = 2111 << 16;
	window['AscDFH'].historyitem_type_ViewPr                 = 2112 << 16;
	window['AscDFH'].historyitem_type_CommonViewPr           = 2113 << 16;
	window['AscDFH'].historyitem_type_CSldViewPr             = 2114 << 16;
	window['AscDFH'].historyitem_type_CViewPr                = 2115 << 16;
	window['AscDFH'].historyitem_type_ViewPrScale            = 2116 << 16;
	window['AscDFH'].historyitem_type_ViewPrGuide            = 2117 << 16;

	window['AscDFH'].historyitem_type_OForm_UserMaster       = 2200 << 16;
	window['AscDFH'].historyitem_type_OForm_User             = 2201 << 16;
	window['AscDFH'].historyitem_type_OForm_FieldMaster      = 2202 << 16;
	window['AscDFH'].historyitem_type_OForm_Document         = 2203 << 16;
	window['AscDFH'].historyitem_type_OForm_FieldGroup       = 2204 << 16;
	
	window['AscDFH'].historyitem_type_PDF_Document			= 2210 << 16;
	window['AscDFH'].historyitem_type_Pdf_Form				= 2211 << 16;
	window['AscDFH'].historyitem_type_Pdf_Comment			= 2212 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot				= 2213 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_Ink			= 2214 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_Line		= 2215 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_FreeText	= 2216 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_Text		= 2217 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_Circle		= 2218 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_Square		= 2219 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_Polygon		= 2220 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_Polyline	= 2221 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_Highlight	= 2222 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_Underline	= 2223 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_Strikeout	= 2224 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_Squiggly	= 2225 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_Caret		= 2226 << 16;
	window['AscDFH'].historyitem_type_Pdf_Pushbutton		= 2227 << 16;
	window['AscDFH'].historyitem_type_Pdf_Drawing			= 2228 << 16;
	window['AscDFH'].historyitem_type_Pdf_Page				= 2229 << 16;
	window['AscDFH'].historyitem_type_Pdf_Annot_Stamp		= 2230 << 16;
	window['AscDFH'].historyitem_type_Pdf_PropLocker		= 2231 << 16;
	window['AscDFH'].historyitem_type_Pdf_Checkbox_Field	= 2232 << 16;
	window['AscDFH'].historyitem_type_Pdf_Combobox_Field	= 2233 << 16;
	window['AscDFH'].historyitem_type_Pdf_Listbox_Field		= 2234 << 16;
	window['AscDFH'].historyitem_type_Pdf_Button_Field		= 2235 << 16;
	window['AscDFH'].historyitem_type_Pdf_Radiobutton_Field	= 2236 << 16;
	window['AscDFH'].historyitem_type_Pdf_Signature_Field	= 2237 << 16;
	window['AscDFH'].historyitem_type_Pdf_Text_Field		= 2238 << 16;
	
	window['AscDFH'].historyitem_type_CustomProperties      = 2301 << 16;

	window['AscDFH'].historyitem_type_CEffectProperties      = 2302 << 16;

	

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//
	// Типы изменений, разбитые по классам
	//
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	window['AscDFH'].historyitem_Unknown_Unknown = window['AscDFH'].historyitem_type_Unknown | 0;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CTableId
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_TableId_Add         = window['AscDFH'].historyitem_type_TableId | 1;
	window['AscDFH'].historyitem_TableId_Description = window['AscDFH'].historyitem_type_TableId | 0xFFFF;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CDocument
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Document_AddItem                         = window['AscDFH'].historyitem_type_Document | 1;
	window['AscDFH'].historyitem_Document_RemoveItem                      = window['AscDFH'].historyitem_type_Document | 2;
	window['AscDFH'].historyitem_Document_DefaultTab                      = window['AscDFH'].historyitem_type_Document | 3;
	window['AscDFH'].historyitem_Document_EvenAndOddHeaders               = window['AscDFH'].historyitem_type_Document | 4;
	window['AscDFH'].historyitem_Document_DefaultLanguage                 = window['AscDFH'].historyitem_type_Document | 5;
	window['AscDFH'].historyitem_Document_MathSettings                    = window['AscDFH'].historyitem_type_Document | 6;
	window['AscDFH'].historyitem_Document_SdtGlobalSettings               = window['AscDFH'].historyitem_type_Document | 7;
	window['AscDFH'].historyitem_Document_Settings_GutterAtTop            = window['AscDFH'].historyitem_type_Document | 8;
	window['AscDFH'].historyitem_Document_Settings_MirrorMargins          = window['AscDFH'].historyitem_type_Document | 9;
	window['AscDFH'].historyitem_Document_SpecialFormsGlobalSettings      = window['AscDFH'].historyitem_type_Document | 10;
	window['AscDFH'].historyitem_Document_Settings_TrackRevisions         = window['AscDFH'].historyitem_type_Document | 11;
	window['AscDFH'].historyitem_Document_Settings_AutoHyphenation        = window['AscDFH'].historyitem_type_Document | 12;
	window['AscDFH'].historyitem_Document_Settings_ConsecutiveHyphenLimit = window['AscDFH'].historyitem_type_Document | 13;
	window['AscDFH'].historyitem_Document_Settings_DoNotHyphenateCaps     = window['AscDFH'].historyitem_type_Document | 14;
	window['AscDFH'].historyitem_Document_Settings_HyphenationZone        = window['AscDFH'].historyitem_type_Document | 15;
	window['AscDFH'].historyitem_Document_PageColor                       = window['AscDFH'].historyitem_type_Document | 16;
	
	window['AscDFH'].historyitem_Document_DisconnectEveryone              = window['AscDFH'].historyitem_type_Document | 10000;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе Paragraph
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Paragraph_AddItem                   = window['AscDFH'].historyitem_type_Paragraph | 1;
	window['AscDFH'].historyitem_Paragraph_RemoveItem                = window['AscDFH'].historyitem_type_Paragraph | 2;
	window['AscDFH'].historyitem_Paragraph_Numbering                 = window['AscDFH'].historyitem_type_Paragraph | 3;
	window['AscDFH'].historyitem_Paragraph_Align                     = window['AscDFH'].historyitem_type_Paragraph | 4;
	window['AscDFH'].historyitem_Paragraph_Ind_First                 = window['AscDFH'].historyitem_type_Paragraph | 5;
	window['AscDFH'].historyitem_Paragraph_Ind_Right                 = window['AscDFH'].historyitem_type_Paragraph | 6;
	window['AscDFH'].historyitem_Paragraph_Ind_Left                  = window['AscDFH'].historyitem_type_Paragraph | 7;
	window['AscDFH'].historyitem_Paragraph_ContextualSpacing         = window['AscDFH'].historyitem_type_Paragraph | 8;
	window['AscDFH'].historyitem_Paragraph_KeepLines                 = window['AscDFH'].historyitem_type_Paragraph | 9;
	window['AscDFH'].historyitem_Paragraph_KeepNext                  = window['AscDFH'].historyitem_type_Paragraph | 10;
	window['AscDFH'].historyitem_Paragraph_PageBreakBefore           = window['AscDFH'].historyitem_type_Paragraph | 11;
	window['AscDFH'].historyitem_Paragraph_Spacing_Line              = window['AscDFH'].historyitem_type_Paragraph | 12;
	window['AscDFH'].historyitem_Paragraph_Spacing_LineRule          = window['AscDFH'].historyitem_type_Paragraph | 13;
	window['AscDFH'].historyitem_Paragraph_Spacing_Before            = window['AscDFH'].historyitem_type_Paragraph | 14;
	window['AscDFH'].historyitem_Paragraph_Spacing_After             = window['AscDFH'].historyitem_type_Paragraph | 15;
	window['AscDFH'].historyitem_Paragraph_Spacing_AfterAutoSpacing  = window['AscDFH'].historyitem_type_Paragraph | 16;
	window['AscDFH'].historyitem_Paragraph_Spacing_BeforeAutoSpacing = window['AscDFH'].historyitem_type_Paragraph | 17;
	window['AscDFH'].historyitem_Paragraph_Shd_Value                 = window['AscDFH'].historyitem_type_Paragraph | 18;
	window['AscDFH'].historyitem_Paragraph_Shd_Color                 = window['AscDFH'].historyitem_type_Paragraph | 19;
	window['AscDFH'].historyitem_Paragraph_Shd_Unifill               = window['AscDFH'].historyitem_type_Paragraph | 20;
	window['AscDFH'].historyitem_Paragraph_Shd                       = window['AscDFH'].historyitem_type_Paragraph | 21;
	window['AscDFH'].historyitem_Paragraph_WidowControl              = window['AscDFH'].historyitem_type_Paragraph | 22;
	window['AscDFH'].historyitem_Paragraph_Tabs                      = window['AscDFH'].historyitem_type_Paragraph | 23;
	window['AscDFH'].historyitem_Paragraph_PStyle                    = window['AscDFH'].historyitem_type_Paragraph | 24;
	window['AscDFH'].historyitem_Paragraph_Borders_Between           = window['AscDFH'].historyitem_type_Paragraph | 25;
	window['AscDFH'].historyitem_Paragraph_Borders_Bottom            = window['AscDFH'].historyitem_type_Paragraph | 26;
	window['AscDFH'].historyitem_Paragraph_Borders_Left              = window['AscDFH'].historyitem_type_Paragraph | 27;
	window['AscDFH'].historyitem_Paragraph_Borders_Right             = window['AscDFH'].historyitem_type_Paragraph | 28;
	window['AscDFH'].historyitem_Paragraph_Borders_Top               = window['AscDFH'].historyitem_type_Paragraph | 29;
	window['AscDFH'].historyitem_Paragraph_Pr                        = window['AscDFH'].historyitem_type_Paragraph | 30;
	window['AscDFH'].historyitem_Paragraph_PresentationPr_Bullet     = window['AscDFH'].historyitem_type_Paragraph | 31;
	window['AscDFH'].historyitem_Paragraph_PresentationPr_Level      = window['AscDFH'].historyitem_type_Paragraph | 32;
	window['AscDFH'].historyitem_Paragraph_FramePr                   = window['AscDFH'].historyitem_type_Paragraph | 33;
	window['AscDFH'].historyitem_Paragraph_SectionPr                 = window['AscDFH'].historyitem_type_Paragraph | 34;
	window['AscDFH'].historyitem_Paragraph_PrChange                  = window['AscDFH'].historyitem_type_Paragraph | 35;
	window['AscDFH'].historyitem_Paragraph_PrReviewInfo              = window['AscDFH'].historyitem_type_Paragraph | 36;
	window['AscDFH'].historyitem_Paragraph_OutlineLvl                = window['AscDFH'].historyitem_type_Paragraph | 37;
	window['AscDFH'].historyitem_Paragraph_DefaultTabSize            = window['AscDFH'].historyitem_type_Paragraph | 38;
	window['AscDFH'].historyitem_Paragraph_SuppressLineNumbers       = window['AscDFH'].historyitem_type_Paragraph | 39;
	window['AscDFH'].historyitem_Paragraph_Shd_Fill                  = window['AscDFH'].historyitem_type_Paragraph | 40;
	window['AscDFH'].historyitem_Paragraph_Shd_ThemeFill             = window['AscDFH'].historyitem_type_Paragraph | 41;
	window['AscDFH'].historyitem_Paragraph_Bidi                      = window['AscDFH'].historyitem_type_Paragraph | 42;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе ParaTextPr
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_TextPr_Bold                  = window['AscDFH'].historyitem_type_TextPr | 1;
	window['AscDFH'].historyitem_TextPr_Italic                = window['AscDFH'].historyitem_type_TextPr | 2;
	window['AscDFH'].historyitem_TextPr_Strikeout             = window['AscDFH'].historyitem_type_TextPr | 3;
	window['AscDFH'].historyitem_TextPr_Underline             = window['AscDFH'].historyitem_type_TextPr | 4;
	window['AscDFH'].historyitem_TextPr_FontSize              = window['AscDFH'].historyitem_type_TextPr | 5;
	window['AscDFH'].historyitem_TextPr_Color                 = window['AscDFH'].historyitem_type_TextPr | 6;
	window['AscDFH'].historyitem_TextPr_VertAlign             = window['AscDFH'].historyitem_type_TextPr | 7;
	window['AscDFH'].historyitem_TextPr_HighLight             = window['AscDFH'].historyitem_type_TextPr | 8;
	window['AscDFH'].historyitem_TextPr_RStyle                = window['AscDFH'].historyitem_type_TextPr | 9;
	window['AscDFH'].historyitem_TextPr_Spacing               = window['AscDFH'].historyitem_type_TextPr | 10;
	window['AscDFH'].historyitem_TextPr_DStrikeout            = window['AscDFH'].historyitem_type_TextPr | 11;
	window['AscDFH'].historyitem_TextPr_Caps                  = window['AscDFH'].historyitem_type_TextPr | 12;
	window['AscDFH'].historyitem_TextPr_SmallCaps             = window['AscDFH'].historyitem_type_TextPr | 13;
	window['AscDFH'].historyitem_TextPr_Position              = window['AscDFH'].historyitem_type_TextPr | 14;
	window['AscDFH'].historyitem_TextPr_Value                 = window['AscDFH'].historyitem_type_TextPr | 15;
	window['AscDFH'].historyitem_TextPr_RFonts                = window['AscDFH'].historyitem_type_TextPr | 16;
	window['AscDFH'].historyitem_TextPr_RFonts_Ascii          = window['AscDFH'].historyitem_type_TextPr | 17;
	window['AscDFH'].historyitem_TextPr_RFonts_HAnsi          = window['AscDFH'].historyitem_type_TextPr | 18;
	window['AscDFH'].historyitem_TextPr_RFonts_CS             = window['AscDFH'].historyitem_type_TextPr | 19;
	window['AscDFH'].historyitem_TextPr_RFonts_EastAsia       = window['AscDFH'].historyitem_type_TextPr | 20;
	window['AscDFH'].historyitem_TextPr_RFonts_Hint           = window['AscDFH'].historyitem_type_TextPr | 21;
	window['AscDFH'].historyitem_TextPr_Lang                  = window['AscDFH'].historyitem_type_TextPr | 22;
	window['AscDFH'].historyitem_TextPr_Lang_Bidi             = window['AscDFH'].historyitem_type_TextPr | 23;
	window['AscDFH'].historyitem_TextPr_Lang_EastAsia         = window['AscDFH'].historyitem_type_TextPr | 24;
	window['AscDFH'].historyitem_TextPr_Lang_Val              = window['AscDFH'].historyitem_type_TextPr | 25;
	window['AscDFH'].historyitem_TextPr_Unifill               = window['AscDFH'].historyitem_type_TextPr | 26;
	window['AscDFH'].historyitem_TextPr_FontSizeCS            = window['AscDFH'].historyitem_type_TextPr | 27;
	window['AscDFH'].historyitem_TextPr_Outline               = window['AscDFH'].historyitem_type_TextPr | 28;
	window['AscDFH'].historyitem_TextPr_Fill                  = window['AscDFH'].historyitem_type_TextPr | 29;
	window['AscDFH'].historyitem_TextPr_HighlightColor        = window['AscDFH'].historyitem_type_TextPr | 30;
	window['AscDFH'].historyitem_TextPr_RFonts_Ascii_Theme    = window['AscDFH'].historyitem_type_TextPr | 31;
	window['AscDFH'].historyitem_TextPr_RFonts_HAnsi_Theme    = window['AscDFH'].historyitem_type_TextPr | 32;
	window['AscDFH'].historyitem_TextPr_RFonts_CS_Theme       = window['AscDFH'].historyitem_type_TextPr | 33;
	window['AscDFH'].historyitem_TextPr_RFonts_EastAsia_Theme = window['AscDFH'].historyitem_type_TextPr | 34;
	window['AscDFH'].historyitem_TextPr_BoldCS                = window['AscDFH'].historyitem_type_TextPr | 35;
	window['AscDFH'].historyitem_TextPr_ItalicCS              = window['AscDFH'].historyitem_type_TextPr | 36;
	window['AscDFH'].historyitem_TextPr_Ligatures             = window['AscDFH'].historyitem_type_TextPr | 37;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе ParaDrawing
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Drawing_DrawingType       = window['AscDFH'].historyitem_type_Drawing | 1;
	window['AscDFH'].historyitem_Drawing_WrappingType      = window['AscDFH'].historyitem_type_Drawing | 2;
	window['AscDFH'].historyitem_Drawing_Distance          = window['AscDFH'].historyitem_type_Drawing | 3;
	window['AscDFH'].historyitem_Drawing_AllowOverlap      = window['AscDFH'].historyitem_type_Drawing | 4;
	window['AscDFH'].historyitem_Drawing_PositionH         = window['AscDFH'].historyitem_type_Drawing | 5;
	window['AscDFH'].historyitem_Drawing_PositionV         = window['AscDFH'].historyitem_type_Drawing | 6;
	window['AscDFH'].historyitem_Drawing_BehindDoc         = window['AscDFH'].historyitem_type_Drawing | 7;
	window['AscDFH'].historyitem_Drawing_SetGraphicObject  = window['AscDFH'].historyitem_type_Drawing | 8;
	window['AscDFH'].historyitem_Drawing_SetSimplePos      = window['AscDFH'].historyitem_type_Drawing | 9;
	window['AscDFH'].historyitem_Drawing_SetExtent         = window['AscDFH'].historyitem_type_Drawing | 10;
	window['AscDFH'].historyitem_Drawing_SetWrapPolygon    = window['AscDFH'].historyitem_type_Drawing | 11;
	window['AscDFH'].historyitem_Drawing_SetLocked         = window['AscDFH'].historyitem_type_Drawing | 12;
	window['AscDFH'].historyitem_Drawing_SetRelativeHeight = window['AscDFH'].historyitem_type_Drawing | 13;
	window['AscDFH'].historyitem_Drawing_SetEffectExtent   = window['AscDFH'].historyitem_type_Drawing | 14;
	window['AscDFH'].historyitem_Drawing_SetParent         = window['AscDFH'].historyitem_type_Drawing | 15;
	window['AscDFH'].historyitem_Drawing_SetParaMath       = window['AscDFH'].historyitem_type_Drawing | 16;
	window['AscDFH'].historyitem_Drawing_LayoutInCell      = window['AscDFH'].historyitem_type_Drawing | 17;
	window['AscDFH'].historyitem_Drawing_SetSizeRelH       = window['AscDFH'].historyitem_type_Drawing | 18;
	window['AscDFH'].historyitem_Drawing_SetSizeRelV       = window['AscDFH'].historyitem_type_Drawing | 19;
	window['AscDFH'].historyitem_Drawing_Form              = window['AscDFH'].historyitem_type_Drawing | 20;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CTable
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Table_TableW                = window['AscDFH'].historyitem_type_Table | 1;
	window['AscDFH'].historyitem_Table_TableCellMar          = window['AscDFH'].historyitem_type_Table | 2;
	window['AscDFH'].historyitem_Table_TableAlign            = window['AscDFH'].historyitem_type_Table | 3;
	window['AscDFH'].historyitem_Table_TableInd              = window['AscDFH'].historyitem_type_Table | 4;
	window['AscDFH'].historyitem_Table_TableBorder_Left      = window['AscDFH'].historyitem_type_Table | 5;
	window['AscDFH'].historyitem_Table_TableBorder_Top       = window['AscDFH'].historyitem_type_Table | 6;
	window['AscDFH'].historyitem_Table_TableBorder_Right     = window['AscDFH'].historyitem_type_Table | 7;
	window['AscDFH'].historyitem_Table_TableBorder_Bottom    = window['AscDFH'].historyitem_type_Table | 8;
	window['AscDFH'].historyitem_Table_TableBorder_InsideH   = window['AscDFH'].historyitem_type_Table | 9;
	window['AscDFH'].historyitem_Table_TableBorder_InsideV   = window['AscDFH'].historyitem_type_Table | 10;
	window['AscDFH'].historyitem_Table_TableShd              = window['AscDFH'].historyitem_type_Table | 11;
	window['AscDFH'].historyitem_Table_Inline                = window['AscDFH'].historyitem_type_Table | 12;
	window['AscDFH'].historyitem_Table_AddRow                = window['AscDFH'].historyitem_type_Table | 13;
	window['AscDFH'].historyitem_Table_RemoveRow             = window['AscDFH'].historyitem_type_Table | 14;
	window['AscDFH'].historyitem_Table_TableGrid             = window['AscDFH'].historyitem_type_Table | 15;
	window['AscDFH'].historyitem_Table_TableLook             = window['AscDFH'].historyitem_type_Table | 16;
	window['AscDFH'].historyitem_Table_TableStyleRowBandSize = window['AscDFH'].historyitem_type_Table | 17;
	window['AscDFH'].historyitem_Table_TableStyleColBandSize = window['AscDFH'].historyitem_type_Table | 18;
	window['AscDFH'].historyitem_Table_TableStyle            = window['AscDFH'].historyitem_type_Table | 19;
	window['AscDFH'].historyitem_Table_AllowOverlap          = window['AscDFH'].historyitem_type_Table | 20;
	window['AscDFH'].historyitem_Table_PositionH             = window['AscDFH'].historyitem_type_Table | 21;
	window['AscDFH'].historyitem_Table_PositionV             = window['AscDFH'].historyitem_type_Table | 22;
	window['AscDFH'].historyitem_Table_Distance              = window['AscDFH'].historyitem_type_Table | 23;
	window['AscDFH'].historyitem_Table_Pr                    = window['AscDFH'].historyitem_type_Table | 24;
	window['AscDFH'].historyitem_Table_TableLayout           = window['AscDFH'].historyitem_type_Table | 25;
	window['AscDFH'].historyitem_Table_TableDescription      = window['AscDFH'].historyitem_type_Table | 26;
	window['AscDFH'].historyitem_Table_TableCaption          = window['AscDFH'].historyitem_type_Table | 27;
	window['AscDFH'].historyitem_Table_TableGridChange       = window['AscDFH'].historyitem_type_Table | 28;
	window['AscDFH'].historyitem_Table_PrChange              = window['AscDFH'].historyitem_type_Table | 29;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CTableRow
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_TableRow_Before      = window['AscDFH'].historyitem_type_TableRow | 1;
	window['AscDFH'].historyitem_TableRow_After       = window['AscDFH'].historyitem_type_TableRow | 2;
	window['AscDFH'].historyitem_TableRow_CellSpacing = window['AscDFH'].historyitem_type_TableRow | 3;
	window['AscDFH'].historyitem_TableRow_Height      = window['AscDFH'].historyitem_type_TableRow | 4;
	window['AscDFH'].historyitem_TableRow_AddCell     = window['AscDFH'].historyitem_type_TableRow | 5;
	window['AscDFH'].historyitem_TableRow_RemoveCell  = window['AscDFH'].historyitem_type_TableRow | 6;
	window['AscDFH'].historyitem_TableRow_TableHeader = window['AscDFH'].historyitem_type_TableRow | 7;
	window['AscDFH'].historyitem_TableRow_Pr          = window['AscDFH'].historyitem_type_TableRow | 8;
	window['AscDFH'].historyitem_TableRow_PrChange    = window['AscDFH'].historyitem_type_TableRow | 9;
	window['AscDFH'].historyitem_TableRow_ReviewType  = window['AscDFH'].historyitem_type_TableRow | 10;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CTableCell
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_TableCell_GridSpan      = window['AscDFH'].historyitem_type_TableCell | 1;
	window['AscDFH'].historyitem_TableCell_Margins       = window['AscDFH'].historyitem_type_TableCell | 2;
	window['AscDFH'].historyitem_TableCell_Shd           = window['AscDFH'].historyitem_type_TableCell | 3;
	window['AscDFH'].historyitem_TableCell_VMerge        = window['AscDFH'].historyitem_type_TableCell | 4;
	window['AscDFH'].historyitem_TableCell_Border_Left   = window['AscDFH'].historyitem_type_TableCell | 5;
	window['AscDFH'].historyitem_TableCell_Border_Right  = window['AscDFH'].historyitem_type_TableCell | 6;
	window['AscDFH'].historyitem_TableCell_Border_Top    = window['AscDFH'].historyitem_type_TableCell | 7;
	window['AscDFH'].historyitem_TableCell_Border_Bottom = window['AscDFH'].historyitem_type_TableCell | 8;
	window['AscDFH'].historyitem_TableCell_VAlign        = window['AscDFH'].historyitem_type_TableCell | 9;
	window['AscDFH'].historyitem_TableCell_W             = window['AscDFH'].historyitem_type_TableCell | 10;
	window['AscDFH'].historyitem_TableCell_Pr            = window['AscDFH'].historyitem_type_TableCell | 11;
	window['AscDFH'].historyitem_TableCell_TextDirection = window['AscDFH'].historyitem_type_TableCell | 12;
	window['AscDFH'].historyitem_TableCell_NoWrap        = window['AscDFH'].historyitem_type_TableCell | 13;
	window['AscDFH'].historyitem_TableCell_HMerge        = window['AscDFH'].historyitem_type_TableCell | 14;
	window['AscDFH'].historyitem_TableCell_PrChange      = window['AscDFH'].historyitem_type_TableCell | 15;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CDocumentContent
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_DocumentContent_AddItem    = window['AscDFH'].historyitem_type_DocumentContent | 1;
	window['AscDFH'].historyitem_DocumentContent_RemoveItem = window['AscDFH'].historyitem_type_DocumentContent | 2;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CAbstractNum
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_AbstractNum_LvlChange    = window['AscDFH'].historyitem_type_AbstractNum | 1;
	window['AscDFH'].historyitem_AbstractNum_TextPrChange = window['AscDFH'].historyitem_type_AbstractNum | 2;
	window['AscDFH'].historyitem_AbstractNum_ParaPrChange = window['AscDFH'].historyitem_type_AbstractNum | 3;
	window['AscDFH'].historyitem_AbstractNum_StyleLink    = window['AscDFH'].historyitem_type_AbstractNum | 4;
	window['AscDFH'].historyitem_AbstractNum_NumStyleLink = window['AscDFH'].historyitem_type_AbstractNum | 5;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CNum
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Num_LvlOverrideChange = window['AscDFH'].historyitem_type_Num | 1;
	window['AscDFH'].historyitem_Num_AbstractNum       = window['AscDFH'].historyitem_type_Num | 2;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CPresentationField
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_PresentationField_Guid      = window['AscDFH'].historyitem_type_PresentationField | 1;
	window['AscDFH'].historyitem_PresentationField_FieldType = window['AscDFH'].historyitem_type_PresentationField | 2;
	window['AscDFH'].historyitem_PresentationField_PPr       = window['AscDFH'].historyitem_type_PresentationField | 3;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе СComment
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Comment_Change     = window['AscDFH'].historyitem_type_Comment | 1;
	window['AscDFH'].historyitem_Comment_TypeInfo   = window['AscDFH'].historyitem_type_Comment | 2;
	window['AscDFH'].historyitem_Comment_Position   = window['AscDFH'].historyitem_type_Comment | 3;
	window['AscDFH'].historyitem_Comment_RangeStart = window['AscDFH'].historyitem_type_Comment | 4;
	window['AscDFH'].historyitem_Comment_RangeEnd   = window['AscDFH'].historyitem_type_Comment | 5;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе AscCommon.CComments
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Comments_Add    = window['AscDFH'].historyitem_type_Comments | 1;
	window['AscDFH'].historyitem_Comments_Remove = window['AscDFH'].historyitem_type_Comments | 2;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе ParaHyperlink
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Hyperlink_Value      = window['AscDFH'].historyitem_type_Hyperlink | 1;
	window['AscDFH'].historyitem_Hyperlink_ToolTip    = window['AscDFH'].historyitem_type_Hyperlink | 2;
	window['AscDFH'].historyitem_Hyperlink_AddItem    = window['AscDFH'].historyitem_type_Hyperlink | 3;
	window['AscDFH'].historyitem_Hyperlink_RemoveItem = window['AscDFH'].historyitem_type_Hyperlink | 4;
	window['AscDFH'].historyitem_Hyperlink_Anchor     = window['AscDFH'].historyitem_type_Hyperlink | 5;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CStyle
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Style_TextPr          = window['AscDFH'].historyitem_type_Style | 1;
	window['AscDFH'].historyitem_Style_ParaPr          = window['AscDFH'].historyitem_type_Style | 2;
	window['AscDFH'].historyitem_Style_TablePr         = window['AscDFH'].historyitem_type_Style | 3;
	window['AscDFH'].historyitem_Style_TableRowPr      = window['AscDFH'].historyitem_type_Style | 4;
	window['AscDFH'].historyitem_Style_TableCellPr     = window['AscDFH'].historyitem_type_Style | 5;
	window['AscDFH'].historyitem_Style_TableBand1Horz  = window['AscDFH'].historyitem_type_Style | 6;
	window['AscDFH'].historyitem_Style_TableBand1Vert  = window['AscDFH'].historyitem_type_Style | 7;
	window['AscDFH'].historyitem_Style_TableBand2Horz  = window['AscDFH'].historyitem_type_Style | 8;
	window['AscDFH'].historyitem_Style_TableBand2Vert  = window['AscDFH'].historyitem_type_Style | 9;
	window['AscDFH'].historyitem_Style_TableFirstCol   = window['AscDFH'].historyitem_type_Style | 10;
	window['AscDFH'].historyitem_Style_TableFirstRow   = window['AscDFH'].historyitem_type_Style | 11;
	window['AscDFH'].historyitem_Style_TableLastCol    = window['AscDFH'].historyitem_type_Style | 12;
	window['AscDFH'].historyitem_Style_TableLastRow    = window['AscDFH'].historyitem_type_Style | 13;
	window['AscDFH'].historyitem_Style_TableTLCell     = window['AscDFH'].historyitem_type_Style | 14;
	window['AscDFH'].historyitem_Style_TableTRCell     = window['AscDFH'].historyitem_type_Style | 15;
	window['AscDFH'].historyitem_Style_TableBLCell     = window['AscDFH'].historyitem_type_Style | 16;
	window['AscDFH'].historyitem_Style_TableBRCell     = window['AscDFH'].historyitem_type_Style | 17;
	window['AscDFH'].historyitem_Style_TableWholeTable = window['AscDFH'].historyitem_type_Style | 18;
	window['AscDFH'].historyitem_Style_Name            = window['AscDFH'].historyitem_type_Style | 101;
	window['AscDFH'].historyitem_Style_BasedOn         = window['AscDFH'].historyitem_type_Style | 102;
	window['AscDFH'].historyitem_Style_Next            = window['AscDFH'].historyitem_type_Style | 103;
	window['AscDFH'].historyitem_Style_Type            = window['AscDFH'].historyitem_type_Style | 104;
	window['AscDFH'].historyitem_Style_QFormat         = window['AscDFH'].historyitem_type_Style | 105;
	window['AscDFH'].historyitem_Style_UiPriority      = window['AscDFH'].historyitem_type_Style | 106;
	window['AscDFH'].historyitem_Style_Hidden          = window['AscDFH'].historyitem_type_Style | 107;
	window['AscDFH'].historyitem_Style_SemiHidden      = window['AscDFH'].historyitem_type_Style | 108;
	window['AscDFH'].historyitem_Style_UnhideWhenUsed  = window['AscDFH'].historyitem_type_Style | 109;
	window['AscDFH'].historyitem_Style_Link            = window['AscDFH'].historyitem_type_Style | 110;
	window['AscDFH'].historyitem_Style_Custom          = window['AscDFH'].historyitem_type_Style | 111;
	window['AscDFH'].historyitem_Style_StyleId         = window['AscDFH'].historyitem_type_Style | 112;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CStyles
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Styles_Add                 = window['AscDFH'].historyitem_type_Styles | 1;
	window['AscDFH'].historyitem_Styles_Remove              = window['AscDFH'].historyitem_type_Styles | 2;
	window['AscDFH'].historyitem_Styles_ChangeDefaultTextPr = window['AscDFH'].historyitem_type_Styles | 3;
	window['AscDFH'].historyitem_Styles_ChangeDefaultParaPr = window['AscDFH'].historyitem_type_Styles | 4;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе ParaMath
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_MathContent_AddItem      = window['AscDFH'].historyitem_type_Math | 101;
	window['AscDFH'].historyitem_MathContent_RemoveItem   = window['AscDFH'].historyitem_type_Math | 102;
	window['AscDFH'].historyitem_MathContent_ArgSize      = window['AscDFH'].historyitem_type_Math | 103;
	window['AscDFH'].historyitem_MathPara_Jc              = window['AscDFH'].historyitem_type_Math | 201;
	window['AscDFH'].historyitem_MathBase_AddItems        = window['AscDFH'].historyitem_type_Math | 301;
	window['AscDFH'].historyitem_MathBase_RemoveItems     = window['AscDFH'].historyitem_type_Math | 302;
	window['AscDFH'].historyitem_MathBase_FontSize        = window['AscDFH'].historyitem_type_Math | 303;
	window['AscDFH'].historyitem_MathBase_Shd             = window['AscDFH'].historyitem_type_Math | 304;
	window['AscDFH'].historyitem_MathBase_Color           = window['AscDFH'].historyitem_type_Math | 305;
	window['AscDFH'].historyitem_MathBase_Unifill         = window['AscDFH'].historyitem_type_Math | 306;
	window['AscDFH'].historyitem_MathBase_Underline       = window['AscDFH'].historyitem_type_Math | 307;
	window['AscDFH'].historyitem_MathBase_Strikeout       = window['AscDFH'].historyitem_type_Math | 308;
	window['AscDFH'].historyitem_MathBase_DoubleStrikeout = window['AscDFH'].historyitem_type_Math | 309;
	window['AscDFH'].historyitem_MathBase_Italic          = window['AscDFH'].historyitem_type_Math | 310;
	window['AscDFH'].historyitem_MathBase_Bold            = window['AscDFH'].historyitem_type_Math | 311;
	window['AscDFH'].historyitem_MathBase_RFontsAscii     = window['AscDFH'].historyitem_type_Math | 312;
	window['AscDFH'].historyitem_MathBase_RFontsHAnsi     = window['AscDFH'].historyitem_type_Math | 313;
	window['AscDFH'].historyitem_MathBase_RFontsCS        = window['AscDFH'].historyitem_type_Math | 314;
	window['AscDFH'].historyitem_MathBase_RFontsEastAsia  = window['AscDFH'].historyitem_type_Math | 315;
	window['AscDFH'].historyitem_MathBase_RFontsHint      = window['AscDFH'].historyitem_type_Math | 316;
	window['AscDFH'].historyitem_MathBase_HighLight       = window['AscDFH'].historyitem_type_Math | 317;
	window['AscDFH'].historyitem_MathBase_ReviewInfo      = window['AscDFH'].historyitem_type_Math | 318;
	window['AscDFH'].historyitem_MathBase_TextFill        = window['AscDFH'].historyitem_type_Math | 319;
	window['AscDFH'].historyitem_MathBase_TextOutline     = window['AscDFH'].historyitem_type_Math | 320;
	window['AscDFH'].historyitem_MathBase_HighlightColor  = window['AscDFH'].historyitem_type_Math | 321;
	window['AscDFH'].historyitem_MathBox_AlnAt            = window['AscDFH'].historyitem_type_Math | 401;
	window['AscDFH'].historyitem_MathBox_ForcedBreak      = window['AscDFH'].historyitem_type_Math | 402;
	window['AscDFH'].historyitem_MathFraction_Type        = window['AscDFH'].historyitem_type_Math | 501;
	window['AscDFH'].historyitem_MathRadical_HideDegree   = window['AscDFH'].historyitem_type_Math | 601;
	window['AscDFH'].historyitem_MathNary_LimLoc          = window['AscDFH'].historyitem_type_Math | 701;
	window['AscDFH'].historyitem_MathNary_UpperLimit      = window['AscDFH'].historyitem_type_Math | 702;
	window['AscDFH'].historyitem_MathNary_LowerLimit      = window['AscDFH'].historyitem_type_Math | 703;
	window['AscDFH'].historyitem_MathDelimiter_BegOper    = window['AscDFH'].historyitem_type_Math | 801;
	window['AscDFH'].historyitem_MathDelimiter_EndOper    = window['AscDFH'].historyitem_type_Math | 802;
	window['AscDFH'].historyitem_MathDelimiter_Grow       = window['AscDFH'].historyitem_type_Math | 803;
	window['AscDFH'].historyitem_MathDelimiter_Shape      = window['AscDFH'].historyitem_type_Math | 804;
	window['AscDFH'].historyitem_MathDelimiter_SetColumn  = window['AscDFH'].historyitem_type_Math | 805;
	window['AscDFH'].historyitem_MathGroupChar_Pr         = window['AscDFH'].historyitem_type_Math | 901;
	window['AscDFH'].historyitem_MathLimit_Type           = window['AscDFH'].historyitem_type_Math | 1001;
	window['AscDFH'].historyitem_MathBorderBox_Top        = window['AscDFH'].historyitem_type_Math | 1101;
	window['AscDFH'].historyitem_MathBorderBox_Bot        = window['AscDFH'].historyitem_type_Math | 1102;
	window['AscDFH'].historyitem_MathBorderBox_Left       = window['AscDFH'].historyitem_type_Math | 1103;
	window['AscDFH'].historyitem_MathBorderBox_Right      = window['AscDFH'].historyitem_type_Math | 1104;
	window['AscDFH'].historyitem_MathBorderBox_Hor        = window['AscDFH'].historyitem_type_Math | 1105;
	window['AscDFH'].historyitem_MathBorderBox_Ver        = window['AscDFH'].historyitem_type_Math | 1106;
	window['AscDFH'].historyitem_MathBorderBox_TopLTR     = window['AscDFH'].historyitem_type_Math | 1107;
	window['AscDFH'].historyitem_MathBorderBox_TopRTL     = window['AscDFH'].historyitem_type_Math | 1108;
	window['AscDFH'].historyitem_MathBar_LinePos          = window['AscDFH'].historyitem_type_Math | 1201;
	window['AscDFH'].historyitem_MathMatrix_AddRow        = window['AscDFH'].historyitem_type_Math | 1301;
	window['AscDFH'].historyitem_MathMatrix_RemoveRow     = window['AscDFH'].historyitem_type_Math | 1302;
	window['AscDFH'].historyitem_MathMatrix_AddColumn     = window['AscDFH'].historyitem_type_Math | 1303;
	window['AscDFH'].historyitem_MathMatrix_RemoveColumn  = window['AscDFH'].historyitem_type_Math | 1304;
	window['AscDFH'].historyitem_MathMatrix_BaseJc        = window['AscDFH'].historyitem_type_Math | 1305;
	window['AscDFH'].historyitem_MathMatrix_ColumnJc      = window['AscDFH'].historyitem_type_Math | 1306;
	window['AscDFH'].historyitem_MathMatrix_Interval      = window['AscDFH'].historyitem_type_Math | 1307;
	window['AscDFH'].historyitem_MathMatrix_Plh           = window['AscDFH'].historyitem_type_Math | 1308;
	window['AscDFH'].historyitem_MathDegree_SubSupType    = window['AscDFH'].historyitem_type_Math | 1401;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе ParaRun
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_ParaRun_AddItem               = window['AscDFH'].historyitem_type_ParaRun | 1;
	window['AscDFH'].historyitem_ParaRun_RemoveItem            = window['AscDFH'].historyitem_type_ParaRun | 2;
	window['AscDFH'].historyitem_ParaRun_Bold                  = window['AscDFH'].historyitem_type_ParaRun | 3;
	window['AscDFH'].historyitem_ParaRun_Italic                = window['AscDFH'].historyitem_type_ParaRun | 4;
	window['AscDFH'].historyitem_ParaRun_Strikeout             = window['AscDFH'].historyitem_type_ParaRun | 5;
	window['AscDFH'].historyitem_ParaRun_Underline             = window['AscDFH'].historyitem_type_ParaRun | 6;
	window['AscDFH'].historyitem_ParaRun_FontFamily            = window['AscDFH'].historyitem_type_ParaRun | 7; // obsolete
	window['AscDFH'].historyitem_ParaRun_FontSize              = window['AscDFH'].historyitem_type_ParaRun | 8;
	window['AscDFH'].historyitem_ParaRun_Color                 = window['AscDFH'].historyitem_type_ParaRun | 9;
	window['AscDFH'].historyitem_ParaRun_VertAlign             = window['AscDFH'].historyitem_type_ParaRun | 10;
	window['AscDFH'].historyitem_ParaRun_HighLight             = window['AscDFH'].historyitem_type_ParaRun | 11;
	window['AscDFH'].historyitem_ParaRun_RStyle                = window['AscDFH'].historyitem_type_ParaRun | 12;
	window['AscDFH'].historyitem_ParaRun_Spacing               = window['AscDFH'].historyitem_type_ParaRun | 13;
	window['AscDFH'].historyitem_ParaRun_DStrikeout            = window['AscDFH'].historyitem_type_ParaRun | 14;
	window['AscDFH'].historyitem_ParaRun_Caps                  = window['AscDFH'].historyitem_type_ParaRun | 15;
	window['AscDFH'].historyitem_ParaRun_SmallCaps             = window['AscDFH'].historyitem_type_ParaRun | 16;
	window['AscDFH'].historyitem_ParaRun_Position              = window['AscDFH'].historyitem_type_ParaRun | 17;
	window['AscDFH'].historyitem_ParaRun_Value                 = window['AscDFH'].historyitem_type_ParaRun | 18; // obsolete
	window['AscDFH'].historyitem_ParaRun_RFonts                = window['AscDFH'].historyitem_type_ParaRun | 19;
	window['AscDFH'].historyitem_ParaRun_Lang                  = window['AscDFH'].historyitem_type_ParaRun | 20;
	window['AscDFH'].historyitem_ParaRun_RFonts_Ascii          = window['AscDFH'].historyitem_type_ParaRun | 21;
	window['AscDFH'].historyitem_ParaRun_RFonts_HAnsi          = window['AscDFH'].historyitem_type_ParaRun | 22;
	window['AscDFH'].historyitem_ParaRun_RFonts_CS             = window['AscDFH'].historyitem_type_ParaRun | 23;
	window['AscDFH'].historyitem_ParaRun_RFonts_EastAsia       = window['AscDFH'].historyitem_type_ParaRun | 24;
	window['AscDFH'].historyitem_ParaRun_RFonts_Hint           = window['AscDFH'].historyitem_type_ParaRun | 25;
	window['AscDFH'].historyitem_ParaRun_Lang_Bidi             = window['AscDFH'].historyitem_type_ParaRun | 26;
	window['AscDFH'].historyitem_ParaRun_Lang_EastAsia         = window['AscDFH'].historyitem_type_ParaRun | 27;
	window['AscDFH'].historyitem_ParaRun_Lang_Val              = window['AscDFH'].historyitem_type_ParaRun | 28;
	window['AscDFH'].historyitem_ParaRun_TextPr                = window['AscDFH'].historyitem_type_ParaRun | 29;
	window['AscDFH'].historyitem_ParaRun_Unifill               = window['AscDFH'].historyitem_type_ParaRun | 30;
	window['AscDFH'].historyitem_ParaRun_Shd                   = window['AscDFH'].historyitem_type_ParaRun | 31;
	window['AscDFH'].historyitem_ParaRun_MathStyle             = window['AscDFH'].historyitem_type_ParaRun | 32;
	window['AscDFH'].historyitem_ParaRun_MathPrp               = window['AscDFH'].historyitem_type_ParaRun | 33;
	window['AscDFH'].historyitem_ParaRun_ReviewType            = window['AscDFH'].historyitem_type_ParaRun | 34;
	window['AscDFH'].historyitem_ParaRun_PrChange              = window['AscDFH'].historyitem_type_ParaRun | 35;
	window['AscDFH'].historyitem_ParaRun_TextFill              = window['AscDFH'].historyitem_type_ParaRun | 36;
	window['AscDFH'].historyitem_ParaRun_TextOutline           = window['AscDFH'].historyitem_type_ParaRun | 37;
	window['AscDFH'].historyitem_ParaRun_PrReviewInfo          = window['AscDFH'].historyitem_type_ParaRun | 38;
	window['AscDFH'].historyitem_ParaRun_ContentReviewInfo     = window['AscDFH'].historyitem_type_ParaRun | 39;
	window['AscDFH'].historyitem_ParaRun_OnStartSplit          = window['AscDFH'].historyitem_type_ParaRun | 40;
	window['AscDFH'].historyitem_ParaRun_OnEndSplit            = window['AscDFH'].historyitem_type_ParaRun | 41;
	window['AscDFH'].historyitem_ParaRun_MathAlnAt             = window['AscDFH'].historyitem_type_ParaRun | 42;
	window['AscDFH'].historyitem_ParaRun_MathForcedBreak       = window['AscDFH'].historyitem_type_ParaRun | 43;
	window['AscDFH'].historyitem_ParaRun_HighlightColor        = window['AscDFH'].historyitem_type_ParaRun | 44;
	window['AscDFH'].historyitem_ParaRun_RFonts_Ascii_Theme    = window['AscDFH'].historyitem_type_ParaRun | 45;
	window['AscDFH'].historyitem_ParaRun_RFonts_HAnsi_Theme    = window['AscDFH'].historyitem_type_ParaRun | 46;
	window['AscDFH'].historyitem_ParaRun_RFonts_CS_Theme       = window['AscDFH'].historyitem_type_ParaRun | 47;
	window['AscDFH'].historyitem_ParaRun_RFonts_EastAsia_Theme = window['AscDFH'].historyitem_type_ParaRun | 48;
	window['AscDFH'].historyitem_ParaRun_BoldCS                = window['AscDFH'].historyitem_type_ParaRun | 49;
	window['AscDFH'].historyitem_ParaRun_ItalicCS              = window['AscDFH'].historyitem_type_ParaRun | 50;
	window['AscDFH'].historyitem_ParaRun_FontSizeCS            = window['AscDFH'].historyitem_type_ParaRun | 51;
	window['AscDFH'].historyitem_ParaRun_Ligatures             = window['AscDFH'].historyitem_type_ParaRun | 52;
	window['AscDFH'].historyitem_ParaRun_CS                    = window['AscDFH'].historyitem_type_ParaRun | 53;
	window['AscDFH'].historyitem_ParaRun_RTL                   = window['AscDFH'].historyitem_type_ParaRun | 54;
	window['AscDFH'].historyitem_ParaRun_MathMetaData          = window['AscDFH'].historyitem_type_ParaRun | 55;

	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CSectionPr
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Section_PageSize_Orient       = window['AscDFH'].historyitem_type_Section | 1;
	window['AscDFH'].historyitem_Section_PageSize_Size         = window['AscDFH'].historyitem_type_Section | 2;
	window['AscDFH'].historyitem_Section_PageMargins           = window['AscDFH'].historyitem_type_Section | 3;
	window['AscDFH'].historyitem_Section_Type                  = window['AscDFH'].historyitem_type_Section | 4;
	window['AscDFH'].historyitem_Section_Borders_Left          = window['AscDFH'].historyitem_type_Section | 5;
	window['AscDFH'].historyitem_Section_Borders_Top           = window['AscDFH'].historyitem_type_Section | 6;
	window['AscDFH'].historyitem_Section_Borders_Right         = window['AscDFH'].historyitem_type_Section | 7;
	window['AscDFH'].historyitem_Section_Borders_Bottom        = window['AscDFH'].historyitem_type_Section | 8;
	window['AscDFH'].historyitem_Section_Borders_Display       = window['AscDFH'].historyitem_type_Section | 9;
	window['AscDFH'].historyitem_Section_Borders_OffsetFrom    = window['AscDFH'].historyitem_type_Section | 10;
	window['AscDFH'].historyitem_Section_Borders_ZOrder        = window['AscDFH'].historyitem_type_Section | 11;
	window['AscDFH'].historyitem_Section_Header_First          = window['AscDFH'].historyitem_type_Section | 12;
	window['AscDFH'].historyitem_Section_Header_Even           = window['AscDFH'].historyitem_type_Section | 13;
	window['AscDFH'].historyitem_Section_Header_Default        = window['AscDFH'].historyitem_type_Section | 14;
	window['AscDFH'].historyitem_Section_Footer_First          = window['AscDFH'].historyitem_type_Section | 15;
	window['AscDFH'].historyitem_Section_Footer_Even           = window['AscDFH'].historyitem_type_Section | 16;
	window['AscDFH'].historyitem_Section_Footer_Default        = window['AscDFH'].historyitem_type_Section | 17;
	window['AscDFH'].historyitem_Section_TitlePage             = window['AscDFH'].historyitem_type_Section | 18;
	window['AscDFH'].historyitem_Section_PageMargins_Header    = window['AscDFH'].historyitem_type_Section | 19;
	window['AscDFH'].historyitem_Section_PageMargins_Footer    = window['AscDFH'].historyitem_type_Section | 20;
	window['AscDFH'].historyitem_Section_PageNumType_Start     = window['AscDFH'].historyitem_type_Section | 21;
	window['AscDFH'].historyitem_Section_Columns_EqualWidth    = window['AscDFH'].historyitem_type_Section | 22;
	window['AscDFH'].historyitem_Section_Columns_Space         = window['AscDFH'].historyitem_type_Section | 23;
	window['AscDFH'].historyitem_Section_Columns_Num           = window['AscDFH'].historyitem_type_Section | 24;
	window['AscDFH'].historyitem_Section_Columns_Sep           = window['AscDFH'].historyitem_type_Section | 25;
	window['AscDFH'].historyitem_Section_Columns_Col           = window['AscDFH'].historyitem_type_Section | 26;
	window['AscDFH'].historyitem_Section_Columns_SetCols       = window['AscDFH'].historyitem_type_Section | 27;
	window['AscDFH'].historyitem_Section_Footnote_Pos          = window['AscDFH'].historyitem_type_Section | 28;
	window['AscDFH'].historyitem_Section_Footnote_NumStart     = window['AscDFH'].historyitem_type_Section | 29;
	window['AscDFH'].historyitem_Section_Footnote_NumRestart   = window['AscDFH'].historyitem_type_Section | 30;
	window['AscDFH'].historyitem_Section_Footnote_NumFormat    = window['AscDFH'].historyitem_type_Section | 31;
	window['AscDFH'].historyitem_Section_PageMargins_Gutter    = window['AscDFH'].historyitem_type_Section | 32;
	window['AscDFH'].historyitem_Section_Gutter_RTL            = window['AscDFH'].historyitem_type_Section | 33;
	window['AscDFH'].historyitem_Section_Endnote_Pos           = window['AscDFH'].historyitem_type_Section | 34;
	window['AscDFH'].historyitem_Section_Endnote_NumStart      = window['AscDFH'].historyitem_type_Section | 35;
	window['AscDFH'].historyitem_Section_Endnote_NumRestart    = window['AscDFH'].historyitem_type_Section | 36;
	window['AscDFH'].historyitem_Section_Endnote_NumFormat     = window['AscDFH'].historyitem_type_Section | 37;
	window['AscDFH'].historyitem_Section_LnNumType             = window['AscDFH'].historyitem_type_Section | 38;
	window['AscDFH'].historyitem_Section_PageNumType_Format    = window['AscDFH'].historyitem_type_Section | 39;
	window['AscDFH'].historyitem_Section_PageNumType_ChapStyle = window['AscDFH'].historyitem_type_Section | 40;
	window['AscDFH'].historyitem_Section_PageNumType_ChapSep   = window['AscDFH'].historyitem_type_Section | 41;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе AscCommon.ParaComment
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_ParaComment_CommentId = window['AscDFH'].historyitem_type_ParaComment | 1;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе ParaField
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Field_AddItem              = window['AscDFH'].historyitem_type_Field | 1;
	window['AscDFH'].historyitem_Field_RemoveItem           = window['AscDFH'].historyitem_type_Field | 2;
	window['AscDFH'].historyitem_Field_FormFieldName        = window['AscDFH'].historyitem_type_Field | 3;
	window['AscDFH'].historyitem_Field_FormFieldDefaultText = window['AscDFH'].historyitem_type_Field | 4;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе Footnotes
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Footnotes_AddFootnote              = window['AscDFH'].historyitem_type_Footnotes | 1;
	window['AscDFH'].historyitem_Footnotes_SetSeparator             = window['AscDFH'].historyitem_type_Footnotes | 2;
	window['AscDFH'].historyitem_Footnotes_SetContinuationSeparator = window['AscDFH'].historyitem_type_Footnotes | 3;
	window['AscDFH'].historyitem_Footnotes_SetContinuationNotice    = window['AscDFH'].historyitem_type_Footnotes | 4;
	window['AscDFH'].historyitem_Footnotes_RemoveFootnote           = window['AscDFH'].historyitem_type_Footnotes | 5;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CGlossaryDocument
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_GlossaryDocument_AddDocPart = window['AscDFH'].historyitem_type_GlossaryDocument | 1;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CDocPart
	//------------------------------------------------------------------------------------------------------------------
	window["AscDFH"].historyitem_DocPart_Name        = window["AscDFH"].historyitem_type_DocPart | 1;
	window["AscDFH"].historyitem_DocPart_Style       = window["AscDFH"].historyitem_type_DocPart | 2;
	window["AscDFH"].historyitem_DocPart_Types       = window["AscDFH"].historyitem_type_DocPart | 3;
	window["AscDFH"].historyitem_DocPart_Description = window["AscDFH"].historyitem_type_DocPart | 4;
	window["AscDFH"].historyitem_DocPart_GUID        = window["AscDFH"].historyitem_type_DocPart | 5;
	window["AscDFH"].historyitem_DocPart_Category    = window["AscDFH"].historyitem_type_DocPart | 6;
	window["AscDFH"].historyitem_DocPart_Behavior    = window["AscDFH"].historyitem_type_DocPart | 7;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CSdtPr
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_SdtPr_Alias            = window['AscDFH'].historyitem_type_SdtPr | 1;
	window['AscDFH'].historyitem_SdtPr_Id               = window['AscDFH'].historyitem_type_SdtPr | 2;
	window['AscDFH'].historyitem_SdtPr_Tag              = window['AscDFH'].historyitem_type_SdtPr | 3;
	window['AscDFH'].historyitem_SdtPr_Label            = window['AscDFH'].historyitem_type_SdtPr | 4;
	window['AscDFH'].historyitem_SdtPr_Lock             = window['AscDFH'].historyitem_type_SdtPr | 5;
	window['AscDFH'].historyitem_SdtPr_DocPartObj       = window['AscDFH'].historyitem_type_SdtPr | 6;
	window['AscDFH'].historyitem_SdtPr_Appearance       = window['AscDFH'].historyitem_type_SdtPr | 7;
	window['AscDFH'].historyitem_SdtPr_Color            = window['AscDFH'].historyitem_type_SdtPr | 8;
	window['AscDFH'].historyitem_SdtPr_CheckBox         = window['AscDFH'].historyitem_type_SdtPr | 9;
	window['AscDFH'].historyitem_SdtPr_CheckBox_Checked = window['AscDFH'].historyitem_type_SdtPr | 10;
	window['AscDFH'].historyitem_SdtPr_Picture          = window['AscDFH'].historyitem_type_SdtPr | 11;
	window['AscDFH'].historyitem_SdtPr_ComboBox         = window['AscDFH'].historyitem_type_SdtPr | 12;
	window['AscDFH'].historyitem_SdtPr_DropDownList     = window['AscDFH'].historyitem_type_SdtPr | 13;
	window['AscDFH'].historyitem_SdtPr_DatePicker       = window['AscDFH'].historyitem_type_SdtPr | 14;
	window['AscDFH'].historyitem_SdtPr_TextPr           = window['AscDFH'].historyitem_type_SdtPr | 15;
	window['AscDFH'].historyitem_SdtPr_Placeholder      = window['AscDFH'].historyitem_type_SdtPr | 16;
	window['AscDFH'].historyitem_SdtPr_ShowingPlcHdr    = window['AscDFH'].historyitem_type_SdtPr | 17;
	window['AscDFH'].historyitem_SdtPr_Equation         = window['AscDFH'].historyitem_type_SdtPr | 18;
	window['AscDFH'].historyitem_SdtPr_Text             = window['AscDFH'].historyitem_type_SdtPr | 19;
	window['AscDFH'].historyitem_SdtPr_Temporary        = window['AscDFH'].historyitem_type_SdtPr | 20;
	window['AscDFH'].historyitem_SdtPr_TextForm         = window['AscDFH'].historyitem_type_SdtPr | 21;
	window['AscDFH'].historyitem_SdtPr_FormPr           = window['AscDFH'].historyitem_type_SdtPr | 22;
	window['AscDFH'].historyitem_SdtPr_PictureFormPr    = window['AscDFH'].historyitem_type_SdtPr | 23;
	window['AscDFH'].historyitem_SdtPr_ComplexFormPr    = window['AscDFH'].historyitem_type_SdtPr | 24;
	window['AscDFH'].historyitem_SdtPr_OForm            = window['AscDFH'].historyitem_type_SdtPr | 25;
	window['AscDFH'].historyitem_SdtPr_DataBinding      = window['AscDFH'].historyitem_type_SdtPr | 26;
	window['AscDFH'].historyitem_SdtPr_ShdColor         = window['AscDFH'].historyitem_type_SdtPr | 27;
	window['AscDFH'].historyitem_SdtPr_BorderColor      = window['AscDFH'].historyitem_type_SdtPr | 28;
	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CSdtPr
	//------------------------------------------------------------------------------------------------------------------
	window["AscDFH"].historyitem_Endnotes_AddEndnote    = window["AscDFH"].historyitem_type_Endnotes | 1;
	window["AscDFH"].historyitem_Endnotes_RemoveEndnote = window["AscDFH"].historyitem_type_Endnotes | 2;
	//------------------------------------------------------------------------------------------------------------------
	// Графические классы общего назначение (без привязки к конкретному классу)
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_AutoShapes_SetDrawingBaseCoors      = window['AscDFH'].historyitem_type_CommonShape | 101;
	window['AscDFH'].historyitem_AutoShapes_SetWorksheet             = window['AscDFH'].historyitem_type_CommonShape | 102;
	window['AscDFH'].historyitem_AutoShapes_AddToDrawingObjects      = window['AscDFH'].historyitem_type_CommonShape | 103;
	window['AscDFH'].historyitem_AutoShapes_RemoveFromDrawingObjects = window['AscDFH'].historyitem_type_CommonShape | 104;
	window['AscDFH'].historyitem_AutoShapes_SetBFromSerialize        = window['AscDFH'].historyitem_type_CommonShape | 105;
	window['AscDFH'].historyitem_AutoShapes_SetLocks                 = window['AscDFH'].historyitem_type_CommonShape | 106;
	window['AscDFH'].historyitem_AutoShapes_SetDrawingBaseType       = window['AscDFH'].historyitem_type_CommonShape | 107;
	window['AscDFH'].historyitem_AutoShapes_SetDrawingBaseExt        = window['AscDFH'].historyitem_type_CommonShape | 108;
	window['AscDFH'].historyitem_AutoShapes_SetDrawingBasePos        = window['AscDFH'].historyitem_type_CommonShape | 109;
	window['AscDFH'].historyitem_AutoShapes_SetDrawingBaseEditAs     = window['AscDFH'].historyitem_type_CommonShape | 110;

	window['AscDFH'].historyitem_ChartFormatSetChart = window['AscDFH'].historyitem_type_CommonShape | 201;

	window['AscDFH'].historyitem_CommonChart_RemoveSeries         = window['AscDFH'].historyitem_type_CommonShape | 301;
	window['AscDFH'].historyitem_CommonSeries_RemoveDPt           = window['AscDFH'].historyitem_type_CommonShape | 302;
	window['AscDFH'].historyitem_CommonLit_RemoveDPt              = window['AscDFH'].historyitem_type_CommonShape | 303;
	window['AscDFH'].historyitem_CommonChartFormat_SetParent      = window['AscDFH'].historyitem_type_CommonShape | 304;
	window['AscDFH'].historyitem_CommonSeries_SetIdx              = window['AscDFH'].historyitem_type_CommonShape | 305;
	window['AscDFH'].historyitem_CommonSeries_SetOrder            = window['AscDFH'].historyitem_type_CommonShape | 306;
	window['AscDFH'].historyitem_CommonSeries_SetTx               = window['AscDFH'].historyitem_type_CommonShape | 307;
	window['AscDFH'].historyitem_CommonSeries_SetSpPr             = window['AscDFH'].historyitem_type_CommonShape | 308;
	window['AscDFH'].historyitem_CommonChart_AddSeries            = window['AscDFH'].historyitem_type_CommonShape | 309;
	window['AscDFH'].historyitem_CommonChart_SetDlbls             = window['AscDFH'].historyitem_type_CommonShape | 310;
	window['AscDFH'].historyitem_CommonChart_AddAxId              = window['AscDFH'].historyitem_type_CommonShape | 311;
	window['AscDFH'].historyitem_CommonChart_SetVaryColors        = window['AscDFH'].historyitem_type_CommonShape | 312;
	window['AscDFH'].historyitem_CommonChart_AddFilteredSeries    = window['AscDFH'].historyitem_type_CommonShape | 313;
	window['AscDFH'].historyitem_CommonChart_RemoveFilteredSeries = window['AscDFH'].historyitem_type_CommonShape | 314;
	window['AscDFH'].historyitem_CommonChart_DataLabelsRange      = window['AscDFH'].historyitem_type_CommonShape | 315;
	window['AscDFH'].historyitem_CommonChart_AddErrBars           = window['AscDFH'].historyitem_type_CommonShape | 316;

	window['AscDFH'].historyitem_Common_AddWatermark = window['AscDFH'].historyitem_type_CommonShape | 401;
	//------------------------------------------------------------------------------------------------------------------
	// Графические классы
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_Presentation_AddSlide                    = window['AscDFH'].historyitem_type_Presentation | 1;
	window['AscDFH'].historyitem_Presentation_RemoveSlide                 = window['AscDFH'].historyitem_type_Presentation | 2;
	window['AscDFH'].historyitem_Presentation_SlideSize                   = window['AscDFH'].historyitem_type_Presentation | 3;
	window['AscDFH'].historyitem_Presentation_AddSlideMaster              = window['AscDFH'].historyitem_type_Presentation | 4;
	window['AscDFH'].historyitem_Presentation_ChangeTheme                 = window['AscDFH'].historyitem_type_Presentation | 5;
	window['AscDFH'].historyitem_Presentation_ChangeColorScheme           = window['AscDFH'].historyitem_type_Presentation | 6;
	window['AscDFH'].historyitem_Presentation_SetShowPr                   = window['AscDFH'].historyitem_type_Presentation | 7;
	window['AscDFH'].historyitem_Presentation_SetDefaultTextStyle         = window['AscDFH'].historyitem_type_Presentation | 8;
	window['AscDFH'].historyitem_Presentation_AddSection                  = window['AscDFH'].historyitem_type_Presentation | 9;
	window['AscDFH'].historyitem_Presentation_RemoveSection               = window['AscDFH'].historyitem_type_Presentation | 10;
	window["AscDFH"].historyitem_Presentation_SetFirstSlideNum            = window["AscDFH"].historyitem_type_Presentation | 11;
	window["AscDFH"].historyitem_Presentation_SetShowSpecialPlsOnTitleSld = window["AscDFH"].historyitem_type_Presentation | 12;
	window['AscDFH'].historyitem_Presentation_RemoveSlideMaster           = window['AscDFH'].historyitem_type_Presentation | 13;
	window['AscDFH'].historyitem_Presentation_ViewPr                      = window['AscDFH'].historyitem_type_Presentation | 14;
	window['AscDFH'].historyitem_Presentation_NotesSz                     = window['AscDFH'].historyitem_type_Presentation | 15;

	window['AscDFH'].historyitem_ColorMod_SetName = window['AscDFH'].historyitem_type_ColorMod | 1;
	window['AscDFH'].historyitem_ColorMod_SetVal  = window['AscDFH'].historyitem_type_ColorMod | 2;

	window['AscDFH'].historyitem_ColorModifiers_AddColorMod    = window['AscDFH'].historyitem_type_ColorModifiers | 1;
	window['AscDFH'].historyitem_ColorModifiers_RemoveColorMod = window['AscDFH'].historyitem_type_ColorModifiers | 2;

	window['AscDFH'].historyitem_SysColor_SetId = window['AscDFH'].historyitem_type_SysColor | 1;
	window['AscDFH'].historyitem_SysColor_SetR  = window['AscDFH'].historyitem_type_SysColor | 2;
	window['AscDFH'].historyitem_SysColor_SetG  = window['AscDFH'].historyitem_type_SysColor | 3;
	window['AscDFH'].historyitem_SysColor_SetB  = window['AscDFH'].historyitem_type_SysColor | 4;

	window['AscDFH'].historyitem_PrstColor_SetId = window['AscDFH'].historyitem_type_PrstColor | 1;

	window['AscDFH'].historyitem_RGBColor_SetColor = window['AscDFH'].historyitem_type_RGBColor | 1;

	window['AscDFH'].historyitem_SchemeColor_SetId = window['AscDFH'].historyitem_type_SchemeColor | 1;

	window['AscDFH'].historyitem_UniColor_SetColor = window['AscDFH'].historyitem_type_UniColor | 1;
	window['AscDFH'].historyitem_UniColor_SetMods  = window['AscDFH'].historyitem_type_UniColor | 2;

	window['AscDFH'].historyitem_SrcRect_SetLTRB = window['AscDFH'].historyitem_type_SrcRect | 1;

	window['AscDFH'].historyitem_BlipFill_SetRasterImageId  = window['AscDFH'].historyitem_type_BlipFill | 1;
	window['AscDFH'].historyitem_BlipFill_SetVectorImageBin = window['AscDFH'].historyitem_type_BlipFill | 2;
	window['AscDFH'].historyitem_BlipFill_SetSrcRect        = window['AscDFH'].historyitem_type_BlipFill | 3;
	window['AscDFH'].historyitem_BlipFill_SetStretch        = window['AscDFH'].historyitem_type_BlipFill | 4;
	window['AscDFH'].historyitem_BlipFill_SetTile           = window['AscDFH'].historyitem_type_BlipFill | 5;
	window['AscDFH'].historyitem_BlipFill_SetRotWithShape   = window['AscDFH'].historyitem_type_BlipFill | 6;

	window['AscDFH'].historyitem_SolidFill_SetColor = window['AscDFH'].historyitem_type_SolidFill | 1;

	window['AscDFH'].historyitem_Gs_SetColor = window['AscDFH'].historyitem_type_Gs | 1;
	window['AscDFH'].historyitem_Gs_SetPos   = window['AscDFH'].historyitem_type_Gs | 2;

	window['AscDFH'].historyitem_GradLin_SetAngle = window['AscDFH'].historyitem_type_GradLin | 1;
	window['AscDFH'].historyitem_GradLin_SetScale = window['AscDFH'].historyitem_type_GradLin | 2;

	window['AscDFH'].historyitem_GradPath_SetPath = window['AscDFH'].historyitem_type_GradPath | 1;
	window['AscDFH'].historyitem_GradPath_SetRect = window['AscDFH'].historyitem_type_GradPath | 2;

	window['AscDFH'].historyitem_GradFill_AddColor = window['AscDFH'].historyitem_type_GradFill | 1;
	window['AscDFH'].historyitem_GradFill_SetLin   = window['AscDFH'].historyitem_type_GradFill | 2;
	window['AscDFH'].historyitem_GradFill_SetPath  = window['AscDFH'].historyitem_type_GradFill | 3;

	window['AscDFH'].historyitem_PathFill_SetFType = window['AscDFH'].historyitem_type_PathFill | 1;
	window['AscDFH'].historyitem_PathFill_SetFgClr = window['AscDFH'].historyitem_type_PathFill | 2;
	window['AscDFH'].historyitem_PathFill_SetBgClr = window['AscDFH'].historyitem_type_PathFill | 3;

	window['AscDFH'].historyitem_UniFill_SetFill        = window['AscDFH'].historyitem_type_UniFill | 1;
	window['AscDFH'].historyitem_UniFill_SetTransparent = window['AscDFH'].historyitem_type_UniFill | 2;

	window['AscDFH'].historyitem_EndArrow_SetType = window['AscDFH'].historyitem_type_EndArrow | 1;
	window['AscDFH'].historyitem_EndArrow_SetLen  = window['AscDFH'].historyitem_type_EndArrow | 2;
	window['AscDFH'].historyitem_EndArrow_SetW    = window['AscDFH'].historyitem_type_EndArrow | 3;

	window['AscDFH'].historyitem_LineJoin_SetType  = window['AscDFH'].historyitem_type_LineJoin | 1;
	window['AscDFH'].historyitem_LineJoin_SetLimit = window['AscDFH'].historyitem_type_LineJoin | 2;

	window['AscDFH'].historyitem_Ln_SetFill     = window['AscDFH'].historyitem_type_Ln | 1;
	window['AscDFH'].historyitem_Ln_SetPrstDash = window['AscDFH'].historyitem_type_Ln | 2;
	window['AscDFH'].historyitem_Ln_SetJoin     = window['AscDFH'].historyitem_type_Ln | 3;
	window['AscDFH'].historyitem_Ln_SetHeadEnd  = window['AscDFH'].historyitem_type_Ln | 4;
	window['AscDFH'].historyitem_Ln_SetTailEnd  = window['AscDFH'].historyitem_type_Ln | 5;
	window['AscDFH'].historyitem_Ln_SetAlgn     = window['AscDFH'].historyitem_type_Ln | 6;
	window['AscDFH'].historyitem_Ln_SetCap      = window['AscDFH'].historyitem_type_Ln | 7;
	window['AscDFH'].historyitem_Ln_SetCmpd     = window['AscDFH'].historyitem_type_Ln | 8;
	window['AscDFH'].historyitem_Ln_SetW        = window['AscDFH'].historyitem_type_Ln | 9;

	window['AscDFH'].historyitem_DefaultShapeDefinition_SetSpPr     = window['AscDFH'].historyitem_type_DefaultShapeDefinition | 1;
	window['AscDFH'].historyitem_DefaultShapeDefinition_SetBodyPr   = window['AscDFH'].historyitem_type_DefaultShapeDefinition | 2;
	window['AscDFH'].historyitem_DefaultShapeDefinition_SetLstStyle = window['AscDFH'].historyitem_type_DefaultShapeDefinition | 3;
	window['AscDFH'].historyitem_DefaultShapeDefinition_SetStyle    = window['AscDFH'].historyitem_type_DefaultShapeDefinition | 4;

	window['AscDFH'].historyitem_CNvPr_SetId         = window['AscDFH'].historyitem_type_CNvPr | 1;
	window['AscDFH'].historyitem_CNvPr_SetName       = window['AscDFH'].historyitem_type_CNvPr | 2;
	window['AscDFH'].historyitem_CNvPr_SetIsHidden   = window['AscDFH'].historyitem_type_CNvPr | 3;
	window['AscDFH'].historyitem_CNvPr_SetDescr      = window['AscDFH'].historyitem_type_CNvPr | 4;
	window['AscDFH'].historyitem_CNvPr_SetTitle      = window['AscDFH'].historyitem_type_CNvPr | 5;
	window['AscDFH'].historyitem_CNvPr_SetHlinkClick = window['AscDFH'].historyitem_type_CNvPr | 6;
	window['AscDFH'].historyitem_CNvPr_SetHlinkHover = window['AscDFH'].historyitem_type_CNvPr | 7;

	window['AscDFH'].historyitem_NvPr_SetIsPhoto   = window['AscDFH'].historyitem_type_NvPr | 1;
	window['AscDFH'].historyitem_NvPr_SetUserDrawn = window['AscDFH'].historyitem_type_NvPr | 2;
	window['AscDFH'].historyitem_NvPr_SetPh        = window['AscDFH'].historyitem_type_NvPr | 3;
	window['AscDFH'].historyitem_NvPr_SetUniMedia  = window['AscDFH'].historyitem_type_NvPr | 4;
	window['AscDFH'].historyitem_NvPr_AddExt       = window['AscDFH'].historyitem_type_NvPr | 5;

	window['AscDFH'].historyitem_Ph_SetHasCustomPrompt = window['AscDFH'].historyitem_type_Ph | 1;
	window['AscDFH'].historyitem_Ph_SetIdx             = window['AscDFH'].historyitem_type_Ph | 2;
	window['AscDFH'].historyitem_Ph_SetOrient          = window['AscDFH'].historyitem_type_Ph | 3;
	window['AscDFH'].historyitem_Ph_SetSz              = window['AscDFH'].historyitem_type_Ph | 4;
	window['AscDFH'].historyitem_Ph_SetType            = window['AscDFH'].historyitem_type_Ph | 5;

	window['AscDFH'].historyitem_UniNvPr_SetCNvPr    = window['AscDFH'].historyitem_type_UniNvPr | 1;
	window['AscDFH'].historyitem_UniNvPr_SetUniPr    = window['AscDFH'].historyitem_type_UniNvPr | 2;
	window['AscDFH'].historyitem_UniNvPr_SetNvPr     = window['AscDFH'].historyitem_type_UniNvPr | 3;
	window['AscDFH'].historyitem_UniNvPr_SetUniSpPr  = window['AscDFH'].historyitem_type_UniNvPr | 4;

	window['AscDFH'].historyitem_StyleRef_SetIdx   = window['AscDFH'].historyitem_type_StyleRef | 1;
	window['AscDFH'].historyitem_StyleRef_SetColor = window['AscDFH'].historyitem_type_StyleRef | 2;

	window['AscDFH'].historyitem_FontRef_SetIdx   = window['AscDFH'].historyitem_type_FontRef | 1;
	window['AscDFH'].historyitem_FontRef_SetColor = window['AscDFH'].historyitem_type_FontRef | 2;

	window['AscDFH'].historyitem_Chart_SetAutoTitleDeleted = window['AscDFH'].historyitem_type_Chart | 1;
	window['AscDFH'].historyitem_Chart_SetBackWall         = window['AscDFH'].historyitem_type_Chart | 2;
	window['AscDFH'].historyitem_Chart_SetDispBlanksAs     = window['AscDFH'].historyitem_type_Chart | 3;
	window['AscDFH'].historyitem_Chart_SetFloor            = window['AscDFH'].historyitem_type_Chart | 4;
	window['AscDFH'].historyitem_Chart_SetLegend           = window['AscDFH'].historyitem_type_Chart | 5;
	window['AscDFH'].historyitem_Chart_AddPivotFmt         = window['AscDFH'].historyitem_type_Chart | 6;
	window['AscDFH'].historyitem_Chart_SetPlotArea         = window['AscDFH'].historyitem_type_Chart | 7;
	window['AscDFH'].historyitem_Chart_SetPlotVisOnly      = window['AscDFH'].historyitem_type_Chart | 8;
	window['AscDFH'].historyitem_Chart_SetShowDLblsOverMax = window['AscDFH'].historyitem_type_Chart | 9;
	window['AscDFH'].historyitem_Chart_SetSideWall         = window['AscDFH'].historyitem_type_Chart | 10;
	window['AscDFH'].historyitem_Chart_SetTitle            = window['AscDFH'].historyitem_type_Chart | 11;
	window['AscDFH'].historyitem_Chart_SetView3D           = window['AscDFH'].historyitem_type_Chart | 12;

	window['AscDFH'].historyitem_ChartSpace_SetChart          = window['AscDFH'].historyitem_type_ChartSpace | 1;
	window['AscDFH'].historyitem_ChartSpace_SetClrMapOvr      = window['AscDFH'].historyitem_type_ChartSpace | 2;
	window['AscDFH'].historyitem_ChartSpace_SetDate1904       = window['AscDFH'].historyitem_type_ChartSpace | 3;
	window['AscDFH'].historyitem_ChartSpace_SetExternalData   = window['AscDFH'].historyitem_type_ChartSpace | 4;
	window['AscDFH'].historyitem_ChartSpace_SetLang           = window['AscDFH'].historyitem_type_ChartSpace | 5;
	window['AscDFH'].historyitem_ChartSpace_SetPivotSource    = window['AscDFH'].historyitem_type_ChartSpace | 6;
	window['AscDFH'].historyitem_ChartSpace_SetPrintSettings  = window['AscDFH'].historyitem_type_ChartSpace | 7;
	window['AscDFH'].historyitem_ChartSpace_SetProtection     = window['AscDFH'].historyitem_type_ChartSpace | 8;
	window['AscDFH'].historyitem_ChartSpace_SetRoundedCorners = window['AscDFH'].historyitem_type_ChartSpace | 9;
	window['AscDFH'].historyitem_ChartSpace_SetSpPr           = window['AscDFH'].historyitem_type_ChartSpace | 10;
	window['AscDFH'].historyitem_ChartSpace_SetStyle          = window['AscDFH'].historyitem_type_ChartSpace | 11;
	window['AscDFH'].historyitem_ChartSpace_SetTxPr           = window['AscDFH'].historyitem_type_ChartSpace | 12;
	window['AscDFH'].historyitem_ChartSpace_SetUserShapes     = window['AscDFH'].historyitem_type_ChartSpace | 13;
	window['AscDFH'].historyitem_ChartSpace_SetThemeOverride  = window['AscDFH'].historyitem_type_ChartSpace | 14;
	window['AscDFH'].historyitem_ChartSpace_SetGroup          = window['AscDFH'].historyitem_type_ChartSpace | 15;
	window['AscDFH'].historyitem_ChartSpace_SetParent         = window['AscDFH'].historyitem_type_ChartSpace | 16;
	window['AscDFH'].historyitem_ChartSpace_SetNvGrFrProps    = window['AscDFH'].historyitem_type_ChartSpace | 17;
	window['AscDFH'].historyitem_ChartSpace_AddUserShape      = window['AscDFH'].historyitem_type_ChartSpace | 18;
	window['AscDFH'].historyitem_ChartSpace_RemoveUserShape   = window['AscDFH'].historyitem_type_ChartSpace | 19;
	window['AscDFH'].historyitem_ChartSpace_ChartStyle        = window['AscDFH'].historyitem_type_ChartSpace | 20;
	window['AscDFH'].historyitem_ChartSpace_ChartColors       = window['AscDFH'].historyitem_type_ChartSpace | 21;
	window['AscDFH'].historyitem_ChartSpace_SetChartData      = window['AscDFH'].historyitem_type_ChartSpace | 22;

	window['AscDFH'].historyitem_Legend_SetLayout      = window['AscDFH'].historyitem_type_Legend | 1;
	window['AscDFH'].historyitem_Legend_AddLegendEntry = window['AscDFH'].historyitem_type_Legend | 2;
	window['AscDFH'].historyitem_Legend_SetLegendPos   = window['AscDFH'].historyitem_type_Legend | 3;
	window['AscDFH'].historyitem_Legend_SetOverlay     = window['AscDFH'].historyitem_type_Legend | 4;
	window['AscDFH'].historyitem_Legend_SetSpPr        = window['AscDFH'].historyitem_type_Legend | 5;
	window['AscDFH'].historyitem_Legend_SetTxPr        = window['AscDFH'].historyitem_type_Legend | 6;
	window['AscDFH'].historyitem_Legend_SetAlign       = window['AscDFH'].historyitem_type_Legend | 7;

	window['AscDFH'].historyitem_Layout_SetH            = window['AscDFH'].historyitem_type_Layout | 1;
	window['AscDFH'].historyitem_Layout_SetHMode        = window['AscDFH'].historyitem_type_Layout | 2;
	window['AscDFH'].historyitem_Layout_SetLayoutTarget = window['AscDFH'].historyitem_type_Layout | 3;
	window['AscDFH'].historyitem_Layout_SetW            = window['AscDFH'].historyitem_type_Layout | 4;
	window['AscDFH'].historyitem_Layout_SetWMode        = window['AscDFH'].historyitem_type_Layout | 5;
	window['AscDFH'].historyitem_Layout_SetX            = window['AscDFH'].historyitem_type_Layout | 6;
	window['AscDFH'].historyitem_Layout_SetXMode        = window['AscDFH'].historyitem_type_Layout | 7;
	window['AscDFH'].historyitem_Layout_SetY            = window['AscDFH'].historyitem_type_Layout | 8;
	window['AscDFH'].historyitem_Layout_SetYMode        = window['AscDFH'].historyitem_type_Layout | 9;
	window['AscDFH'].historyitem_Layout_SetParent       = window['AscDFH'].historyitem_type_Layout | 10;

	window['AscDFH'].historyitem_LegendEntry_SetDelete = window['AscDFH'].historyitem_type_LegendEntry | 1;
	window['AscDFH'].historyitem_LegendEntry_SetIdx    = window['AscDFH'].historyitem_type_LegendEntry | 2;
	window['AscDFH'].historyitem_LegendEntry_SetTxPr   = window['AscDFH'].historyitem_type_LegendEntry | 3;

	window['AscDFH'].historyitem_PivotFmt_SetDLbl   = window['AscDFH'].historyitem_type_PivotFmt | 1;
	window['AscDFH'].historyitem_PivotFmt_SetIdx    = window['AscDFH'].historyitem_type_PivotFmt | 2;
	window['AscDFH'].historyitem_PivotFmt_SetMarker = window['AscDFH'].historyitem_type_PivotFmt | 3;
	window['AscDFH'].historyitem_PivotFmt_SetSpPr   = window['AscDFH'].historyitem_type_PivotFmt | 4;
	window['AscDFH'].historyitem_PivotFmt_SetTxPr   = window['AscDFH'].historyitem_type_PivotFmt | 5;

	window['AscDFH'].historyitem_DLbl_SetDelete         = window['AscDFH'].historyitem_type_DLbl | 1;
	window['AscDFH'].historyitem_DLbl_SetDLblPos        = window['AscDFH'].historyitem_type_DLbl | 2;
	window['AscDFH'].historyitem_DLbl_SetIdx            = window['AscDFH'].historyitem_type_DLbl | 3;
	window['AscDFH'].historyitem_DLbl_SetLayout         = window['AscDFH'].historyitem_type_DLbl | 4;
	window['AscDFH'].historyitem_DLbl_SetNumFmt         = window['AscDFH'].historyitem_type_DLbl | 5;
	window['AscDFH'].historyitem_DLbl_SetSeparator      = window['AscDFH'].historyitem_type_DLbl | 6;
	window['AscDFH'].historyitem_DLbl_SetShowBubbleSize = window['AscDFH'].historyitem_type_DLbl | 7;
	window['AscDFH'].historyitem_DLbl_SetShowCatName    = window['AscDFH'].historyitem_type_DLbl | 8;
	window['AscDFH'].historyitem_DLbl_SetShowLegendKey  = window['AscDFH'].historyitem_type_DLbl | 9;
	window['AscDFH'].historyitem_DLbl_SetShowPercent    = window['AscDFH'].historyitem_type_DLbl | 10;
	window['AscDFH'].historyitem_DLbl_SetShowSerName    = window['AscDFH'].historyitem_type_DLbl | 11;
	window['AscDFH'].historyitem_DLbl_SetShowVal        = window['AscDFH'].historyitem_type_DLbl | 12;
	window['AscDFH'].historyitem_DLbl_SetSpPr           = window['AscDFH'].historyitem_type_DLbl | 13;
	window['AscDFH'].historyitem_DLbl_SetTx             = window['AscDFH'].historyitem_type_DLbl | 14;
	window['AscDFH'].historyitem_DLbl_SetTxPr           = window['AscDFH'].historyitem_type_DLbl | 15;
	window['AscDFH'].historyitem_DLbl_SetParent         = window['AscDFH'].historyitem_type_DLbl | 16;
	window['AscDFH'].historyitem_DLbl_SetShowDLblsRange = window['AscDFH'].historyitem_type_DLbl | 17;

	window['AscDFH'].historyitem_Marker_SetSize   = window['AscDFH'].historyitem_type_Marker | 1;
	window['AscDFH'].historyitem_Marker_SetSpPr   = window['AscDFH'].historyitem_type_Marker | 2;
	window['AscDFH'].historyitem_Marker_SetSymbol = window['AscDFH'].historyitem_type_Marker | 3;

	window['AscDFH'].historyitem_PlotArea_AddChart          = window['AscDFH'].historyitem_type_PlotArea | 1;
	window['AscDFH'].historyitem_PlotArea_SetCatAx          = window['AscDFH'].historyitem_type_PlotArea | 2;
	window['AscDFH'].historyitem_PlotArea_SetDateAx         = window['AscDFH'].historyitem_type_PlotArea | 3;
	window['AscDFH'].historyitem_PlotArea_SetDTable         = window['AscDFH'].historyitem_type_PlotArea | 4;
	window['AscDFH'].historyitem_PlotArea_SetLayout         = window['AscDFH'].historyitem_type_PlotArea | 5;
	window['AscDFH'].historyitem_PlotArea_SetSerAx          = window['AscDFH'].historyitem_type_PlotArea | 6;
	window['AscDFH'].historyitem_PlotArea_SetSpPr           = window['AscDFH'].historyitem_type_PlotArea | 7;
	window['AscDFH'].historyitem_PlotArea_SetValAx          = window['AscDFH'].historyitem_type_PlotArea | 8;
	window['AscDFH'].historyitem_PlotArea_AddAxis           = window['AscDFH'].historyitem_type_PlotArea | 9;
	window['AscDFH'].historyitem_PlotArea_RemoveChart       = window['AscDFH'].historyitem_type_PlotArea | 10;
	window['AscDFH'].historyitem_PlotArea_RemoveAxis        = window['AscDFH'].historyitem_type_PlotArea | 11;
	window['AscDFH'].historyitem_PlotArea_SetPlotAreaRegion = window['AscDFH'].historyitem_type_PlotArea | 12;

	window['AscDFH'].historyitem_NumFmt_SetFormatCode   = window['AscDFH'].historyitem_type_NumFmt | 1;
	window['AscDFH'].historyitem_NumFmt_SetSourceLinked = window['AscDFH'].historyitem_type_NumFmt | 2;

	window['AscDFH'].historyitem_Scaling_SetLogBase     = window['AscDFH'].historyitem_type_Scaling | 1;
	window['AscDFH'].historyitem_Scaling_SetMax         = window['AscDFH'].historyitem_type_Scaling | 2;
	window['AscDFH'].historyitem_Scaling_SetMin         = window['AscDFH'].historyitem_type_Scaling | 3;
	window['AscDFH'].historyitem_Scaling_SetOrientation = window['AscDFH'].historyitem_type_Scaling | 4;
	window['AscDFH'].historyitem_Scaling_SetParent      = window['AscDFH'].historyitem_type_Scaling | 5;

	window['AscDFH'].historyitem_DTable_SetShowHorzBorder = window['AscDFH'].historyitem_type_DTable | 1;
	window['AscDFH'].historyitem_DTable_SetShowKeys       = window['AscDFH'].historyitem_type_DTable | 2;
	window['AscDFH'].historyitem_DTable_SetShowOutline    = window['AscDFH'].historyitem_type_DTable | 3;
	window['AscDFH'].historyitem_DTable_SetShowVertBorder = window['AscDFH'].historyitem_type_DTable | 4;
	window['AscDFH'].historyitem_DTable_SetSpPr           = window['AscDFH'].historyitem_type_DTable | 5;
	window['AscDFH'].historyitem_DTable_SetTxPr           = window['AscDFH'].historyitem_type_DTable | 6;

	window['AscDFH'].historyitem_LineChart_AddAxId       = window['AscDFH'].historyitem_type_LineChart | 1;
	window['AscDFH'].historyitem_LineChart_SetDLbls      = window['AscDFH'].historyitem_type_LineChart | 2;
	window['AscDFH'].historyitem_LineChart_SetDropLines  = window['AscDFH'].historyitem_type_LineChart | 3;
	window['AscDFH'].historyitem_LineChart_SetGrouping   = window['AscDFH'].historyitem_type_LineChart | 4;
	window['AscDFH'].historyitem_LineChart_SetHiLowLines = window['AscDFH'].historyitem_type_LineChart | 5;
	window['AscDFH'].historyitem_LineChart_SetMarker     = window['AscDFH'].historyitem_type_LineChart | 6;
	window['AscDFH'].historyitem_LineChart_SetSmooth     = window['AscDFH'].historyitem_type_LineChart | 7;
	window['AscDFH'].historyitem_LineChart_SetUpDownBars = window['AscDFH'].historyitem_type_LineChart | 8;
	window['AscDFH'].historyitem_LineChart_SetVaryColors = window['AscDFH'].historyitem_type_LineChart | 9;

	window['AscDFH'].historyitem_DLbls_SetDelete          = window['AscDFH'].historyitem_type_DLbls | 1;
	window['AscDFH'].historyitem_DLbls_SetDLbl            = window['AscDFH'].historyitem_type_DLbls | 2;
	window['AscDFH'].historyitem_DLbls_SetDLblPos         = window['AscDFH'].historyitem_type_DLbls | 3;
	window['AscDFH'].historyitem_DLbls_SetLeaderLines     = window['AscDFH'].historyitem_type_DLbls | 4;
	window['AscDFH'].historyitem_DLbls_SetNumFmt          = window['AscDFH'].historyitem_type_DLbls | 5;
	window['AscDFH'].historyitem_DLbls_SetSeparator       = window['AscDFH'].historyitem_type_DLbls | 6;
	window['AscDFH'].historyitem_DLbls_SetShowBubbleSize  = window['AscDFH'].historyitem_type_DLbls | 7;
	window['AscDFH'].historyitem_DLbls_SetShowCatName     = window['AscDFH'].historyitem_type_DLbls | 8;
	window['AscDFH'].historyitem_DLbls_SetShowLeaderLines = window['AscDFH'].historyitem_type_DLbls | 9;
	window['AscDFH'].historyitem_DLbls_SetShowLegendKey   = window['AscDFH'].historyitem_type_DLbls | 10;
	window['AscDFH'].historyitem_DLbls_SetShowPercent     = window['AscDFH'].historyitem_type_DLbls | 11;
	window['AscDFH'].historyitem_DLbls_SetShowSerName     = window['AscDFH'].historyitem_type_DLbls | 12;
	window['AscDFH'].historyitem_DLbls_SetShowVal         = window['AscDFH'].historyitem_type_DLbls | 13;
	window['AscDFH'].historyitem_DLbls_SetSpPr            = window['AscDFH'].historyitem_type_DLbls | 14;
	window['AscDFH'].historyitem_DLbls_SetTxPr            = window['AscDFH'].historyitem_type_DLbls | 15;

	window['AscDFH'].historyitem_UpDownBars_SetDownBars = window['AscDFH'].historyitem_type_UpDownBars | 1;
	window['AscDFH'].historyitem_UpDownBars_SetGapWidth = window['AscDFH'].historyitem_type_UpDownBars | 2;
	window['AscDFH'].historyitem_UpDownBars_SetUpBars   = window['AscDFH'].historyitem_type_UpDownBars | 3;

	window['AscDFH'].historyitem_BarChart_AddAxId       = window['AscDFH'].historyitem_type_BarChart | 1;
	window['AscDFH'].historyitem_BarChart_SetBarDir     = window['AscDFH'].historyitem_type_BarChart | 2;
	window['AscDFH'].historyitem_BarChart_SetDLbls      = window['AscDFH'].historyitem_type_BarChart | 3;
	window['AscDFH'].historyitem_BarChart_SetGapWidth   = window['AscDFH'].historyitem_type_BarChart | 4;
	window['AscDFH'].historyitem_BarChart_SetGrouping   = window['AscDFH'].historyitem_type_BarChart | 5;
	window['AscDFH'].historyitem_BarChart_SetOverlap    = window['AscDFH'].historyitem_type_BarChart | 6;
	window['AscDFH'].historyitem_BarChart_SetSerLines   = window['AscDFH'].historyitem_type_BarChart | 7;
	window['AscDFH'].historyitem_BarChart_SetVaryColors = window['AscDFH'].historyitem_type_BarChart | 8;
	window['AscDFH'].historyitem_BarChart_Set3D         = window['AscDFH'].historyitem_type_BarChart | 9;
	window['AscDFH'].historyitem_BarChart_SetGapDepth   = window['AscDFH'].historyitem_type_BarChart | 10;
	window['AscDFH'].historyitem_BarChart_SetShape      = window['AscDFH'].historyitem_type_BarChart | 11;

	window['AscDFH'].historyitem_BubbleChart_AddAxId           = window['AscDFH'].historyitem_type_BubbleChart | 1;
	window['AscDFH'].historyitem_BubbleChart_SetBubble3D       = window['AscDFH'].historyitem_type_BubbleChart | 2;
	window['AscDFH'].historyitem_BubbleChart_SetBubbleScale    = window['AscDFH'].historyitem_type_BubbleChart | 3;
	window['AscDFH'].historyitem_BubbleChart_SetDLbls          = window['AscDFH'].historyitem_type_BubbleChart | 4;
	window['AscDFH'].historyitem_BubbleChart_SetShowNegBubbles = window['AscDFH'].historyitem_type_BubbleChart | 5;
	window['AscDFH'].historyitem_BubbleChart_SetSizeRepresents = window['AscDFH'].historyitem_type_BubbleChart | 6;
	window['AscDFH'].historyitem_BubbleChart_SetVaryColors     = window['AscDFH'].historyitem_type_BubbleChart | 7;

	window['AscDFH'].historyitem_DoughnutChart_SetDLbls         = window['AscDFH'].historyitem_type_DoughnutChart | 1;
	window['AscDFH'].historyitem_DoughnutChart_SetFirstSliceAng = window['AscDFH'].historyitem_type_DoughnutChart | 2;
	window['AscDFH'].historyitem_DoughnutChart_SetHoleSize      = window['AscDFH'].historyitem_type_DoughnutChart | 3;
	window['AscDFH'].historyitem_DoughnutChart_SetVaryColor     = window['AscDFH'].historyitem_type_DoughnutChart | 4;

	window['AscDFH'].historyitem_OfPieChart_AddCustSplit     = window['AscDFH'].historyitem_type_OfPieChart | 1;
	window['AscDFH'].historyitem_OfPieChart_SetDLbls         = window['AscDFH'].historyitem_type_OfPieChart | 2;
	window['AscDFH'].historyitem_OfPieChart_SetGapWidth      = window['AscDFH'].historyitem_type_OfPieChart | 3;
	window['AscDFH'].historyitem_OfPieChart_SetOfPieType     = window['AscDFH'].historyitem_type_OfPieChart | 4;
	window['AscDFH'].historyitem_OfPieChart_SetSecondPieSize = window['AscDFH'].historyitem_type_OfPieChart | 5;
	window['AscDFH'].historyitem_OfPieChart_SetSerLines      = window['AscDFH'].historyitem_type_OfPieChart | 6;
	window['AscDFH'].historyitem_OfPieChart_SetSplitPos      = window['AscDFH'].historyitem_type_OfPieChart | 7;
	window['AscDFH'].historyitem_OfPieChart_SetSplitType     = window['AscDFH'].historyitem_type_OfPieChart | 8;
	window['AscDFH'].historyitem_OfPieChart_SetVaryColors    = window['AscDFH'].historyitem_type_OfPieChart | 9;

	window['AscDFH'].historyitem_PieChart_SetDLbls         = window['AscDFH'].historyitem_type_PieChart | 1;
	window['AscDFH'].historyitem_PieChart_SetFirstSliceAng = window['AscDFH'].historyitem_type_PieChart | 2;
	window['AscDFH'].historyitem_PieChart_SetVaryColors    = window['AscDFH'].historyitem_type_PieChart | 3;
	window['AscDFH'].historyitem_PieChart_3D               = window['AscDFH'].historyitem_type_PieChart | 4;

	window['AscDFH'].historyitem_RadarChart_AddAxId       = window['AscDFH'].historyitem_type_RadarChart | 1;
	window['AscDFH'].historyitem_RadarChart_SetDLbls      = window['AscDFH'].historyitem_type_RadarChart | 2;
	window['AscDFH'].historyitem_RadarChart_SetRadarStyle = window['AscDFH'].historyitem_type_RadarChart | 3;
	window['AscDFH'].historyitem_RadarChart_SetVaryColors = window['AscDFH'].historyitem_type_RadarChart | 4;

	window['AscDFH'].historyitem_ScatterChart_AddAxId         = window['AscDFH'].historyitem_type_ScatterChart | 1;
	window['AscDFH'].historyitem_ScatterChart_SetDLbls        = window['AscDFH'].historyitem_type_ScatterChart | 2;
	window['AscDFH'].historyitem_ScatterChart_SetScatterStyle = window['AscDFH'].historyitem_type_ScatterChart | 3;
	window['AscDFH'].historyitem_ScatterChart_SetVaryColors   = window['AscDFH'].historyitem_type_ScatterChart | 4;

	window['AscDFH'].historyitem_StockChart_AddAxId       = window['AscDFH'].historyitem_type_StockChart | 1;
	window['AscDFH'].historyitem_StockChart_SetDLbls      = window['AscDFH'].historyitem_type_StockChart | 2;
	window['AscDFH'].historyitem_StockChart_SetDropLines  = window['AscDFH'].historyitem_type_StockChart | 3;
	window['AscDFH'].historyitem_StockChart_SetHiLowLines = window['AscDFH'].historyitem_type_StockChart | 4;
	window['AscDFH'].historyitem_StockChart_SetUpDownBars = window['AscDFH'].historyitem_type_StockChart | 5;

	window['AscDFH'].historyitem_SurfaceChart_AddAxId      = window['AscDFH'].historyitem_type_SurfaceChart | 1;
	window['AscDFH'].historyitem_SurfaceChart_AddBandFmt   = window['AscDFH'].historyitem_type_SurfaceChart | 2;
	window['AscDFH'].historyitem_SurfaceChart_SetWireframe = window['AscDFH'].historyitem_type_SurfaceChart | 3;

	window['AscDFH'].historyitem_BandFmt_SetIdx  = window['AscDFH'].historyitem_type_BandFmt | 1;
	window['AscDFH'].historyitem_BandFmt_SetSpPr = window['AscDFH'].historyitem_type_BandFmt | 2;

	window['AscDFH'].historyitem_AreaChart_AddAxId       = window['AscDFH'].historyitem_type_AreaChart | 1;
	window['AscDFH'].historyitem_AreaChart_SetDLbls      = window['AscDFH'].historyitem_type_AreaChart | 2;
	window['AscDFH'].historyitem_AreaChart_SetDropLines  = window['AscDFH'].historyitem_type_AreaChart | 3;
	window['AscDFH'].historyitem_AreaChart_SetGrouping   = window['AscDFH'].historyitem_type_AreaChart | 4;
	window['AscDFH'].historyitem_AreaChart_SetVaryColors = window['AscDFH'].historyitem_type_AreaChart | 5;

	window['AscDFH'].historyitem_ScatterSer_SetDLbls     = window['AscDFH'].historyitem_type_ScatterSer | 1;
	window['AscDFH'].historyitem_ScatterSer_SetDPt       = window['AscDFH'].historyitem_type_ScatterSer | 2;
	window['AscDFH'].historyitem_ScatterSer_SetErrBars   = window['AscDFH'].historyitem_type_ScatterSer | 3;
	window['AscDFH'].historyitem_ScatterSer_SetIdx       = window['AscDFH'].historyitem_type_ScatterSer | 4;
	window['AscDFH'].historyitem_ScatterSer_SetMarker    = window['AscDFH'].historyitem_type_ScatterSer | 5;
	window['AscDFH'].historyitem_ScatterSer_SetOrder     = window['AscDFH'].historyitem_type_ScatterSer | 6;
	window['AscDFH'].historyitem_ScatterSer_SetSmooth    = window['AscDFH'].historyitem_type_ScatterSer | 7;
	window['AscDFH'].historyitem_ScatterSer_SetSpPr      = window['AscDFH'].historyitem_type_ScatterSer | 8;
	window['AscDFH'].historyitem_ScatterSer_SetTrendline = window['AscDFH'].historyitem_type_ScatterSer | 9;
	window['AscDFH'].historyitem_ScatterSer_SetXVal      = window['AscDFH'].historyitem_type_ScatterSer | 10;
	window['AscDFH'].historyitem_ScatterSer_SetYVal      = window['AscDFH'].historyitem_type_ScatterSer | 11;

	window['AscDFH'].historyitem_DPt_SetBubble3D         = window['AscDFH'].historyitem_type_DPt | 1;
	window['AscDFH'].historyitem_DPt_SetExplosion        = window['AscDFH'].historyitem_type_DPt | 2;
	window['AscDFH'].historyitem_DPt_SetIdx              = window['AscDFH'].historyitem_type_DPt | 3;
	window['AscDFH'].historyitem_DPt_SetInvertIfNegative = window['AscDFH'].historyitem_type_DPt | 4;
	window['AscDFH'].historyitem_DPt_SetMarker           = window['AscDFH'].historyitem_type_DPt | 5;
	window['AscDFH'].historyitem_DPt_SetPictureOptions   = window['AscDFH'].historyitem_type_DPt | 6;
	window['AscDFH'].historyitem_DPt_SetSpPr             = window['AscDFH'].historyitem_type_DPt | 7;

	window['AscDFH'].historyitem_ErrBars_SetErrBarType = window['AscDFH'].historyitem_type_ErrBars | 1;
	window['AscDFH'].historyitem_ErrBars_SetErrDir     = window['AscDFH'].historyitem_type_ErrBars | 2;
	window['AscDFH'].historyitem_ErrBars_SetErrValType = window['AscDFH'].historyitem_type_ErrBars | 3;
	window['AscDFH'].historyitem_ErrBars_SetMinus      = window['AscDFH'].historyitem_type_ErrBars | 4;
	window['AscDFH'].historyitem_ErrBars_SetNoEndCap   = window['AscDFH'].historyitem_type_ErrBars | 5;
	window['AscDFH'].historyitem_ErrBars_SetPlus       = window['AscDFH'].historyitem_type_ErrBars | 6;
	window['AscDFH'].historyitem_ErrBars_SetSpPr       = window['AscDFH'].historyitem_type_ErrBars | 7;
	window['AscDFH'].historyitem_ErrBars_SetVal        = window['AscDFH'].historyitem_type_ErrBars | 8;

	window['AscDFH'].historyitem_MinusPlus_SetNumLit = window['AscDFH'].historyitem_type_MinusPlus | 1;
	window['AscDFH'].historyitem_MinusPlus_SetNumRef = window['AscDFH'].historyitem_type_MinusPlus | 2;

	window['AscDFH'].historyitem_NumLit_SetFormatCode = window['AscDFH'].historyitem_type_NumLit | 1;
	window['AscDFH'].historyitem_NumLit_AddPt         = window['AscDFH'].historyitem_type_NumLit | 2;
	window['AscDFH'].historyitem_NumLit_SetPtCount    = window['AscDFH'].historyitem_type_NumLit | 3;
	window['AscDFH'].historyitem_NumLit_SetName       = window['AscDFH'].historyitem_type_NumLit | 4;

	window['AscDFH'].historyitem_NumericPoint_SetFormatCode = window['AscDFH'].historyitem_type_NumericPoint | 1;
	window['AscDFH'].historyitem_NumericPoint_SetIdx        = window['AscDFH'].historyitem_type_NumericPoint | 2;
	window['AscDFH'].historyitem_NumericPoint_SetVal        = window['AscDFH'].historyitem_type_NumericPoint | 3;

	window['AscDFH'].historyitem_NumRef_SetF        = window['AscDFH'].historyitem_type_NumRef | 1;
	window['AscDFH'].historyitem_NumRef_SetNumCache = window['AscDFH'].historyitem_type_NumRef | 2;

	window['AscDFH'].historyitem_Trendline_SetBackward      = window['AscDFH'].historyitem_type_TrendLine | 1;
	window['AscDFH'].historyitem_Trendline_SetDispEq        = window['AscDFH'].historyitem_type_TrendLine | 2;
	window['AscDFH'].historyitem_Trendline_SetDispRSqr      = window['AscDFH'].historyitem_type_TrendLine | 3;
	window['AscDFH'].historyitem_Trendline_SetForward       = window['AscDFH'].historyitem_type_TrendLine | 4;
	window['AscDFH'].historyitem_Trendline_SetIntercept     = window['AscDFH'].historyitem_type_TrendLine | 5;
	window['AscDFH'].historyitem_Trendline_SetName          = window['AscDFH'].historyitem_type_TrendLine | 6;
	window['AscDFH'].historyitem_Trendline_SetOrder         = window['AscDFH'].historyitem_type_TrendLine | 7;
	window['AscDFH'].historyitem_Trendline_SetPeriod        = window['AscDFH'].historyitem_type_TrendLine | 8;
	window['AscDFH'].historyitem_Trendline_SetSpPr          = window['AscDFH'].historyitem_type_TrendLine | 9;
	window['AscDFH'].historyitem_Trendline_SetTrendlineLbl  = window['AscDFH'].historyitem_type_TrendLine | 10;
	window['AscDFH'].historyitem_Trendline_SetTrendlineType = window['AscDFH'].historyitem_type_TrendLine | 11;

	window['AscDFH'].historyitem_Tx_SetStrRef = window['AscDFH'].historyitem_type_Tx | 1;
	window['AscDFH'].historyitem_Tx_SetVal    = window['AscDFH'].historyitem_type_Tx | 2;

	window['AscDFH'].historyitem_StrRef_SetF        = window['AscDFH'].historyitem_type_StrRef | 1;
	window['AscDFH'].historyitem_StrRef_SetStrCache = window['AscDFH'].historyitem_type_StrRef | 2;

	window['AscDFH'].historyitem_StrCache_AddPt      = window['AscDFH'].historyitem_type_StrCache | 1;
	window['AscDFH'].historyitem_StrCache_SetPtCount = window['AscDFH'].historyitem_type_StrCache | 2;
	window['AscDFH'].historyitem_StrCache_SetName    = window['AscDFH'].historyitem_type_StrCache | 3;

	window['AscDFH'].historyitem_StrPoint_SetIdx = window['AscDFH'].historyitem_type_StrPoint | 1;
	window['AscDFH'].historyitem_StrPoint_SetVal = window['AscDFH'].historyitem_type_StrPoint | 2;


	window['AscDFH'].historyitem_MultiLvlStrRef_SetF                = window['AscDFH'].historyitem_type_MultiLvlStrRef | 1;
	window['AscDFH'].historyitem_MultiLvlStrRef_SetMultiLvlStrCache = window['AscDFH'].historyitem_type_MultiLvlStrRef | 2;

	window['AscDFH'].historyitem_MultiLvlStrCache_SetLvl     = window['AscDFH'].historyitem_type_MultiLvlStrCache | 1;
	window['AscDFH'].historyitem_MultiLvlStrCache_SetPtCount = window['AscDFH'].historyitem_type_MultiLvlStrCache | 2;

	window['AscDFH'].historyitem_StringLiteral_SetPt      = window['AscDFH'].historyitem_type_StringLiteral | 1;
	window['AscDFH'].historyitem_StringLiteral_SetPtCount = window['AscDFH'].historyitem_type_StringLiteral | 2;

	window['AscDFH'].historyitem_YVal_SetNumLit = window['AscDFH'].historyitem_type_YVal | 1;
	window['AscDFH'].historyitem_YVal_SetNumRef = window['AscDFH'].historyitem_type_YVal | 2;

	window['AscDFH'].historyitem_AreaSeries_SetCat            = window['AscDFH'].historyitem_type_AreaSeries | 1;
	window['AscDFH'].historyitem_AreaSeries_SetDLbls          = window['AscDFH'].historyitem_type_AreaSeries | 2;
	window['AscDFH'].historyitem_AreaSeries_SetDPt            = window['AscDFH'].historyitem_type_AreaSeries | 3;
	window['AscDFH'].historyitem_AreaSeries_SetErrBars        = window['AscDFH'].historyitem_type_AreaSeries | 4;
	window['AscDFH'].historyitem_AreaSeries_SetIdx            = window['AscDFH'].historyitem_type_AreaSeries | 5;
	window['AscDFH'].historyitem_AreaSeries_SetOrder          = window['AscDFH'].historyitem_type_AreaSeries | 6;
	window['AscDFH'].historyitem_AreaSeries_SetPictureOptions = window['AscDFH'].historyitem_type_AreaSeries | 7;
	window['AscDFH'].historyitem_AreaSeries_SetSpPr           = window['AscDFH'].historyitem_type_AreaSeries | 8;
	window['AscDFH'].historyitem_AreaSeries_SetTrendline      = window['AscDFH'].historyitem_type_AreaSeries | 9;
	window['AscDFH'].historyitem_AreaSeries_SetVal            = window['AscDFH'].historyitem_type_AreaSeries | 10;

	window['AscDFH'].historyitem_Cat_SetMultiLvlStrRef = window['AscDFH'].historyitem_type_Cat | 1;
	window['AscDFH'].historyitem_Cat_SetNumLit         = window['AscDFH'].historyitem_type_Cat | 2;
	window['AscDFH'].historyitem_Cat_SetNumRef         = window['AscDFH'].historyitem_type_Cat | 3;
	window['AscDFH'].historyitem_Cat_SetStrLit         = window['AscDFH'].historyitem_type_Cat | 4;
	window['AscDFH'].historyitem_Cat_SetStrRef         = window['AscDFH'].historyitem_type_Cat | 5;

	window['AscDFH'].historyitem_PictureOptions_SetApplyToEnd       = window['AscDFH'].historyitem_type_PictureOptions | 1;
	window['AscDFH'].historyitem_PictureOptions_SetApplyToFront     = window['AscDFH'].historyitem_type_PictureOptions | 2;
	window['AscDFH'].historyitem_PictureOptions_SetApplyToSides     = window['AscDFH'].historyitem_type_PictureOptions | 3;
	window['AscDFH'].historyitem_PictureOptions_SetPictureFormat    = window['AscDFH'].historyitem_type_PictureOptions | 4;
	window['AscDFH'].historyitem_PictureOptions_SetPictureStackUnit = window['AscDFH'].historyitem_type_PictureOptions | 5;

	window['AscDFH'].historyitem_RadarSeries_SetCat    = window['AscDFH'].historyitem_type_RadarSeries | 1;
	window['AscDFH'].historyitem_RadarSeries_SetDLbls  = window['AscDFH'].historyitem_type_RadarSeries | 2;
	window['AscDFH'].historyitem_RadarSeries_SetDPt    = window['AscDFH'].historyitem_type_RadarSeries | 3;
	window['AscDFH'].historyitem_RadarSeries_SetIdx    = window['AscDFH'].historyitem_type_RadarSeries | 4;
	window['AscDFH'].historyitem_RadarSeries_SetMarker = window['AscDFH'].historyitem_type_RadarSeries | 5;
	window['AscDFH'].historyitem_RadarSeries_SetOrder  = window['AscDFH'].historyitem_type_RadarSeries | 6;
	window['AscDFH'].historyitem_RadarSeries_SetSpPr   = window['AscDFH'].historyitem_type_RadarSeries | 7;
	window['AscDFH'].historyitem_RadarSeries_SetVal    = window['AscDFH'].historyitem_type_RadarSeries | 8;

	window['AscDFH'].historyitem_BarSeries_SetCat              = window['AscDFH'].historyitem_type_BarSeries | 1;
	window['AscDFH'].historyitem_BarSeries_SetDLbls            = window['AscDFH'].historyitem_type_BarSeries | 2;
	window['AscDFH'].historyitem_BarSeries_SetDPt              = window['AscDFH'].historyitem_type_BarSeries | 3;
	window['AscDFH'].historyitem_BarSeries_SetErrBars          = window['AscDFH'].historyitem_type_BarSeries | 4;
	window['AscDFH'].historyitem_BarSeries_SetIdx              = window['AscDFH'].historyitem_type_BarSeries | 5;
	window['AscDFH'].historyitem_BarSeries_SetInvertIfNegative = window['AscDFH'].historyitem_type_BarSeries | 6;
	window['AscDFH'].historyitem_BarSeries_SetOrder            = window['AscDFH'].historyitem_type_BarSeries | 7;
	window['AscDFH'].historyitem_BarSeries_SetPictureOptions   = window['AscDFH'].historyitem_type_BarSeries | 8;
	window['AscDFH'].historyitem_BarSeries_SetShape            = window['AscDFH'].historyitem_type_BarSeries | 9;
	window['AscDFH'].historyitem_BarSeries_SetSpPr             = window['AscDFH'].historyitem_type_BarSeries | 10;
	window['AscDFH'].historyitem_BarSeries_SetTrendline        = window['AscDFH'].historyitem_type_BarSeries | 11;
	window['AscDFH'].historyitem_BarSeries_SetVal              = window['AscDFH'].historyitem_type_BarSeries | 12;

	window['AscDFH'].historyitem_LineSeries_SetCat       = window['AscDFH'].historyitem_type_LineSeries | 1;
	window['AscDFH'].historyitem_LineSeries_SetDLbls     = window['AscDFH'].historyitem_type_LineSeries | 2;
	window['AscDFH'].historyitem_LineSeries_SetDPt       = window['AscDFH'].historyitem_type_LineSeries | 3;
	window['AscDFH'].historyitem_LineSeries_SetErrBars   = window['AscDFH'].historyitem_type_LineSeries | 4;
	window['AscDFH'].historyitem_LineSeries_SetIdx       = window['AscDFH'].historyitem_type_LineSeries | 5;
	window['AscDFH'].historyitem_LineSeries_SetMarker    = window['AscDFH'].historyitem_type_LineSeries | 6;
	window['AscDFH'].historyitem_LineSeries_SetOrder     = window['AscDFH'].historyitem_type_LineSeries | 7;
	window['AscDFH'].historyitem_LineSeries_SetSmooth    = window['AscDFH'].historyitem_type_LineSeries | 8;
	window['AscDFH'].historyitem_LineSeries_SetSpPr      = window['AscDFH'].historyitem_type_LineSeries | 9;
	window['AscDFH'].historyitem_LineSeries_SetTrendline = window['AscDFH'].historyitem_type_LineSeries | 10;
	window['AscDFH'].historyitem_LineSeries_SetVal       = window['AscDFH'].historyitem_type_LineSeries | 11;

	window['AscDFH'].historyitem_PieSeries_SetCat       = window['AscDFH'].historyitem_type_PieSeries | 1;
	window['AscDFH'].historyitem_PieSeries_SetDLbls     = window['AscDFH'].historyitem_type_PieSeries | 2;
	window['AscDFH'].historyitem_PieSeries_SetDPt       = window['AscDFH'].historyitem_type_PieSeries | 3;
	window['AscDFH'].historyitem_PieSeries_SetExplosion = window['AscDFH'].historyitem_type_PieSeries | 4;
	window['AscDFH'].historyitem_PieSeries_SetIdx       = window['AscDFH'].historyitem_type_PieSeries | 5;
	window['AscDFH'].historyitem_PieSeries_SetOrder     = window['AscDFH'].historyitem_type_PieSeries | 6;
	window['AscDFH'].historyitem_PieSeries_SetSpPr      = window['AscDFH'].historyitem_type_PieSeries | 7;
	window['AscDFH'].historyitem_PieSeries_SetVal       = window['AscDFH'].historyitem_type_PieSeries | 8;

	window['AscDFH'].historyitem_SurfaceSeries_SetCat   = window['AscDFH'].historyitem_type_SurfaceSeries | 1;
	window['AscDFH'].historyitem_SurfaceSeries_SetIdx   = window['AscDFH'].historyitem_type_SurfaceSeries | 2;
	window['AscDFH'].historyitem_SurfaceSeries_SetOrder = window['AscDFH'].historyitem_type_SurfaceSeries | 3;
	window['AscDFH'].historyitem_SurfaceSeries_SetSpPr  = window['AscDFH'].historyitem_type_SurfaceSeries | 4;
	window['AscDFH'].historyitem_SurfaceSeries_SetVal   = window['AscDFH'].historyitem_type_SurfaceSeries | 5;

	window['AscDFH'].historyitem_BubbleSeries_SetBubble3D         = window['AscDFH'].historyitem_type_BubbleSeries | 1;
	window['AscDFH'].historyitem_BubbleSeries_SetBubbleSize       = window['AscDFH'].historyitem_type_BubbleSeries | 2;
	window['AscDFH'].historyitem_BubbleSeries_SetDLbls            = window['AscDFH'].historyitem_type_BubbleSeries | 3;
	window['AscDFH'].historyitem_BubbleSeries_SetDPt              = window['AscDFH'].historyitem_type_BubbleSeries | 4;
	window['AscDFH'].historyitem_BubbleSeries_SetErrBars          = window['AscDFH'].historyitem_type_BubbleSeries | 5;
	window['AscDFH'].historyitem_BubbleSeries_SetIdx              = window['AscDFH'].historyitem_type_BubbleSeries | 6;
	window['AscDFH'].historyitem_BubbleSeries_SetInvertIfNegative = window['AscDFH'].historyitem_type_BubbleSeries | 7;
	window['AscDFH'].historyitem_BubbleSeries_SetOrder            = window['AscDFH'].historyitem_type_BubbleSeries | 8;
	window['AscDFH'].historyitem_BubbleSeries_SetSpPr             = window['AscDFH'].historyitem_type_BubbleSeries | 9;
	window['AscDFH'].historyitem_BubbleSeries_SetTrendline        = window['AscDFH'].historyitem_type_BubbleSeries | 10;
	window['AscDFH'].historyitem_BubbleSeries_SetXVal             = window['AscDFH'].historyitem_type_BubbleSeries | 11;
	window['AscDFH'].historyitem_BubbleSeries_SetYVal             = window['AscDFH'].historyitem_type_BubbleSeries | 12;

	window['AscDFH'].historyitem_ExternalData_SetAutoUpdate = window['AscDFH'].historyitem_type_ExternalData | 1;
	window['AscDFH'].historyitem_ExternalData_SetId         = window['AscDFH'].historyitem_type_ExternalData | 2;

	window['AscDFH'].historyitem_PivotSource_SetFmtId = window['AscDFH'].historyitem_type_PivotSource | 1;
	window['AscDFH'].historyitem_PivotSource_SetName  = window['AscDFH'].historyitem_type_PivotSource | 2;

	window['AscDFH'].historyitem_Protection_SetChartObject   = window['AscDFH'].historyitem_type_Protection | 1;
	window['AscDFH'].historyitem_Protection_SetData          = window['AscDFH'].historyitem_type_Protection | 2;
	window['AscDFH'].historyitem_Protection_SetFormatting    = window['AscDFH'].historyitem_type_Protection | 3;
	window['AscDFH'].historyitem_Protection_SetSelection     = window['AscDFH'].historyitem_type_Protection | 4;
	window['AscDFH'].historyitem_Protection_SetUserInterface = window['AscDFH'].historyitem_type_Protection | 5;

	window['AscDFH'].historyitem_ChartWall_SetPictureOptions = window['AscDFH'].historyitem_type_ChartWall | 1;
	window['AscDFH'].historyitem_ChartWall_SetSpPr           = window['AscDFH'].historyitem_type_ChartWall | 2;
	window['AscDFH'].historyitem_ChartWall_SetThickness      = window['AscDFH'].historyitem_type_ChartWall | 3;

	window['AscDFH'].historyitem_View3d_SetDepthPercent = window['AscDFH'].historyitem_type_View3d | 1;
	window['AscDFH'].historyitem_View3d_SetHPercent     = window['AscDFH'].historyitem_type_View3d | 2;
	window['AscDFH'].historyitem_View3d_SetPerspective  = window['AscDFH'].historyitem_type_View3d | 3;
	window['AscDFH'].historyitem_View3d_SetRAngAx       = window['AscDFH'].historyitem_type_View3d | 4;
	window['AscDFH'].historyitem_View3d_SetRotX         = window['AscDFH'].historyitem_type_View3d | 5;
	window['AscDFH'].historyitem_View3d_SetRotY         = window['AscDFH'].historyitem_type_View3d | 6;

	window['AscDFH'].historyitem_ChartText_SetRich   = window['AscDFH'].historyitem_type_ChartText | 1;
	window['AscDFH'].historyitem_ChartText_SetStrRef = window['AscDFH'].historyitem_type_ChartText | 2;
	window['AscDFH'].historyitem_ChartText_SetTxData = window['AscDFH'].historyitem_type_ChartText | 3;

	window['AscDFH'].historyitem_ShapeStyle_SetLnRef     = window['AscDFH'].historyitem_type_ShapeStyle | 1;
	window['AscDFH'].historyitem_ShapeStyle_SetFillRef   = window['AscDFH'].historyitem_type_ShapeStyle | 2;
	window['AscDFH'].historyitem_ShapeStyle_SetFontRef   = window['AscDFH'].historyitem_type_ShapeStyle | 3;
	window['AscDFH'].historyitem_ShapeStyle_SetEffectRef = window['AscDFH'].historyitem_type_ShapeStyle | 4;

	window['AscDFH'].historyitem_Xfrm_SetOffX   = window['AscDFH'].historyitem_type_Xfrm | 1;
	window['AscDFH'].historyitem_Xfrm_SetOffY   = window['AscDFH'].historyitem_type_Xfrm | 2;
	window['AscDFH'].historyitem_Xfrm_SetExtX   = window['AscDFH'].historyitem_type_Xfrm | 3;
	window['AscDFH'].historyitem_Xfrm_SetExtY   = window['AscDFH'].historyitem_type_Xfrm | 4;
	window['AscDFH'].historyitem_Xfrm_SetChOffX = window['AscDFH'].historyitem_type_Xfrm | 5;
	window['AscDFH'].historyitem_Xfrm_SetChOffY = window['AscDFH'].historyitem_type_Xfrm | 6;
	window['AscDFH'].historyitem_Xfrm_SetChExtX = window['AscDFH'].historyitem_type_Xfrm | 7;
	window['AscDFH'].historyitem_Xfrm_SetChExtY = window['AscDFH'].historyitem_type_Xfrm | 8;
	window['AscDFH'].historyitem_Xfrm_SetFlipH  = window['AscDFH'].historyitem_type_Xfrm | 9;
	window['AscDFH'].historyitem_Xfrm_SetFlipV  = window['AscDFH'].historyitem_type_Xfrm | 10;
	window['AscDFH'].historyitem_Xfrm_SetRot    = window['AscDFH'].historyitem_type_Xfrm | 11;
	window['AscDFH'].historyitem_Xfrm_SetParent = window['AscDFH'].historyitem_type_Xfrm | 12;

	window['AscDFH'].historyitem_SpPr_SetBwMode   = window['AscDFH'].historyitem_type_SpPr | 1;
	window['AscDFH'].historyitem_SpPr_SetXfrm     = window['AscDFH'].historyitem_type_SpPr | 2;
	window['AscDFH'].historyitem_SpPr_SetGeometry = window['AscDFH'].historyitem_type_SpPr | 3;
	window['AscDFH'].historyitem_SpPr_SetFill     = window['AscDFH'].historyitem_type_SpPr | 4;
	window['AscDFH'].historyitem_SpPr_SetLn       = window['AscDFH'].historyitem_type_SpPr | 5;
	window['AscDFH'].historyitem_SpPr_SetParent   = window['AscDFH'].historyitem_type_SpPr | 6;
	window['AscDFH'].historyitem_SpPr_SetEffectPr = window['AscDFH'].historyitem_type_SpPr | 7;

	window['AscDFH'].historyitem_ClrScheme_AddClr  = window['AscDFH'].historyitem_type_ClrScheme | 1;
	window['AscDFH'].historyitem_ClrScheme_SetName = window['AscDFH'].historyitem_type_ClrScheme | 2;

	window['AscDFH'].historyitem_ClrMap_SetClr = window['AscDFH'].historyitem_type_ClrMap | 1;

	window['AscDFH'].historyitem_ExtraClrScheme_SetClrScheme = window['AscDFH'].historyitem_type_ExtraClrScheme | 1;
	window['AscDFH'].historyitem_ExtraClrScheme_SetClrMap    = window['AscDFH'].historyitem_type_ExtraClrScheme | 2;

	window['AscDFH'].historyitem_FontCollection_SetFontScheme = window['AscDFH'].historyitem_type_FontCollection | 1;
	window['AscDFH'].historyitem_FontCollection_SetLatin      = window['AscDFH'].historyitem_type_FontCollection | 2;
	window['AscDFH'].historyitem_FontCollection_SetEA         = window['AscDFH'].historyitem_type_FontCollection | 3;
	window['AscDFH'].historyitem_FontCollection_SetCS         = window['AscDFH'].historyitem_type_FontCollection | 4;

	window['AscDFH'].historyitem_FontScheme_SetName      = window['AscDFH'].historyitem_type_FontScheme | 1;
	window['AscDFH'].historyitem_FontScheme_SetMajorFont = window['AscDFH'].historyitem_type_FontScheme | 2;
	window['AscDFH'].historyitem_FontScheme_SetMinorFont = window['AscDFH'].historyitem_type_FontScheme | 3;

	window['AscDFH'].historyitem_FormatScheme_SetName             = window['AscDFH'].historyitem_type_FormatScheme | 1;
	window['AscDFH'].historyitem_FormatScheme_AddFillToStyleLst   = window['AscDFH'].historyitem_type_FormatScheme | 2;
	window['AscDFH'].historyitem_FormatScheme_AddLnToStyleLst     = window['AscDFH'].historyitem_type_FormatScheme | 3;
	window['AscDFH'].historyitem_FormatScheme_AddEffectToStyleLst = window['AscDFH'].historyitem_type_FormatScheme | 4;
	window['AscDFH'].historyitem_FormatScheme_AddBgFillToStyleLst = window['AscDFH'].historyitem_type_FormatScheme | 5;

	window['AscDFH'].historyitem_ThemeElements_SetClrScheme  = window['AscDFH'].historyitem_type_ThemeElements | 1;
	window['AscDFH'].historyitem_ThemeElements_SetFontScheme = window['AscDFH'].historyitem_type_ThemeElements | 2;
	window['AscDFH'].historyitem_ThemeElements_SetFmtScheme  = window['AscDFH'].historyitem_type_ThemeElements | 3;

	window['AscDFH'].historyitem_HF_SetDt     = window['AscDFH'].historyitem_type_HF | 1;
	window['AscDFH'].historyitem_HF_SetFtr    = window['AscDFH'].historyitem_type_HF | 2;
	window['AscDFH'].historyitem_HF_SetHdr    = window['AscDFH'].historyitem_type_HF | 3;
	window['AscDFH'].historyitem_HF_SetSldNum = window['AscDFH'].historyitem_type_HF | 4;

	window['AscDFH'].historyitem_BgPr_SetFill         = window['AscDFH'].historyitem_type_BgPr | 1;
	window['AscDFH'].historyitem_BgPr_SetShadeToTitle = window['AscDFH'].historyitem_type_BgPr | 2;

	window['AscDFH'].historyitem_BgSetBwMode = window['AscDFH'].historyitem_type_Bg | 1;
	window['AscDFH'].historyitem_BgSetBgPr   = window['AscDFH'].historyitem_type_Bg | 2;
	window['AscDFH'].historyitem_BgSetBgRef  = window['AscDFH'].historyitem_type_Bg | 3;

	window['AscDFH'].historyitem_PrintSettingsSetHeaderFooter = window['AscDFH'].historyitem_type_PrintSettings | 1;
	window['AscDFH'].historyitem_PrintSettingsSetPageMargins  = window['AscDFH'].historyitem_type_PrintSettings | 2;
	window['AscDFH'].historyitem_PrintSettingsSetPageSetup    = window['AscDFH'].historyitem_type_PrintSettings | 3;

	window['AscDFH'].historyitem_HeaderFooterChartSetAlignWithMargins = window['AscDFH'].historyitem_type_HeaderFooterChart | 1;
	window['AscDFH'].historyitem_HeaderFooterChartSetDifferentFirst   = window['AscDFH'].historyitem_type_HeaderFooterChart | 2;
	window['AscDFH'].historyitem_HeaderFooterChartSetDifferentOddEven = window['AscDFH'].historyitem_type_HeaderFooterChart | 3;
	window['AscDFH'].historyitem_HeaderFooterChartSetEvenFooter       = window['AscDFH'].historyitem_type_HeaderFooterChart | 4;
	window['AscDFH'].historyitem_HeaderFooterChartSetEvenHeader       = window['AscDFH'].historyitem_type_HeaderFooterChart | 5;
	window['AscDFH'].historyitem_HeaderFooterChartSetFirstFooter      = window['AscDFH'].historyitem_type_HeaderFooterChart | 6;
	window['AscDFH'].historyitem_HeaderFooterChartSetFirstHeader      = window['AscDFH'].historyitem_type_HeaderFooterChart | 7;
	window['AscDFH'].historyitem_HeaderFooterChartSetOddFooter        = window['AscDFH'].historyitem_type_HeaderFooterChart | 8;
	window['AscDFH'].historyitem_HeaderFooterChartSetOddHeader        = window['AscDFH'].historyitem_type_HeaderFooterChart | 9;

	window['AscDFH'].historyitem_PageMarginsSetB      = window['AscDFH'].historyitem_type_PageMarginsChart | 1;
	window['AscDFH'].historyitem_PageMarginsSetFooter = window['AscDFH'].historyitem_type_PageMarginsChart | 2;
	window['AscDFH'].historyitem_PageMarginsSetHeader = window['AscDFH'].historyitem_type_PageMarginsChart | 3;
	window['AscDFH'].historyitem_PageMarginsSetL      = window['AscDFH'].historyitem_type_PageMarginsChart | 4;
	window['AscDFH'].historyitem_PageMarginsSetR      = window['AscDFH'].historyitem_type_PageMarginsChart | 5;
	window['AscDFH'].historyitem_PageMarginsSetT      = window['AscDFH'].historyitem_type_PageMarginsChart | 6;

	window['AscDFH'].historyitem_PageSetupSetBlackAndWhite    = window['AscDFH'].historyitem_type_PageSetup | 1;
	window['AscDFH'].historyitem_PageSetupSetCopies           = window['AscDFH'].historyitem_type_PageSetup | 2;
	window['AscDFH'].historyitem_PageSetupSetDraft            = window['AscDFH'].historyitem_type_PageSetup | 3;
	window['AscDFH'].historyitem_PageSetupSetFirstPageNumber  = window['AscDFH'].historyitem_type_PageSetup | 4;
	window['AscDFH'].historyitem_PageSetupSetHorizontalDpi    = window['AscDFH'].historyitem_type_PageSetup | 5;
	window['AscDFH'].historyitem_PageSetupSetOrientation      = window['AscDFH'].historyitem_type_PageSetup | 6;
	window['AscDFH'].historyitem_PageSetupSetPaperHeight      = window['AscDFH'].historyitem_type_PageSetup | 7;
	window['AscDFH'].historyitem_PageSetupSetPaperSize        = window['AscDFH'].historyitem_type_PageSetup | 8;
	window['AscDFH'].historyitem_PageSetupSetPaperWidth       = window['AscDFH'].historyitem_type_PageSetup | 9;
	window['AscDFH'].historyitem_PageSetupSetUseFirstPageNumb = window['AscDFH'].historyitem_type_PageSetup | 10;
	window['AscDFH'].historyitem_PageSetupSetVerticalDpi      = window['AscDFH'].historyitem_type_PageSetup | 11;

	window['AscDFH'].historyitem_ShapeSetBDeleted               = window['AscDFH'].historyitem_type_Shape | 1;
	window['AscDFH'].historyitem_ShapeSetNvSpPr                 = window['AscDFH'].historyitem_type_Shape | 2;
	window['AscDFH'].historyitem_ShapeSetSpPr                   = window['AscDFH'].historyitem_type_Shape | 3;
	window['AscDFH'].historyitem_ShapeSetStyle                  = window['AscDFH'].historyitem_type_Shape | 4;
	window['AscDFH'].historyitem_ShapeSetTxBody                 = window['AscDFH'].historyitem_type_Shape | 5;
	window['AscDFH'].historyitem_ShapeSetTextBoxContent         = window['AscDFH'].historyitem_type_Shape | 6;
	window['AscDFH'].historyitem_ShapeSetParent                 = window['AscDFH'].historyitem_type_Shape | 7;
	window['AscDFH'].historyitem_ShapeSetGroup                  = window['AscDFH'].historyitem_type_Shape | 8;
	window['AscDFH'].historyitem_ShapeSetBodyPr                 = window['AscDFH'].historyitem_type_Shape | 9;
	window['AscDFH'].historyitem_ShapeSetWordShape              = window['AscDFH'].historyitem_type_Shape | 10;
	window['AscDFH'].historyitem_ShapeSetSignature              = window['AscDFH'].historyitem_type_Shape | 11;
	window['AscDFH'].historyitem_ShapeSetMacro                  = window['AscDFH'].historyitem_type_Shape | 12;
	window['AscDFH'].historyitem_ShapeSetTextLink               = window['AscDFH'].historyitem_type_Shape | 13;
	window['AscDFH'].historyitem_ShapeSetModelId                = window['AscDFH'].historyitem_type_Shape | 14;
	window['AscDFH'].historyitem_ShapeSetTxXfrm                 = window['AscDFH'].historyitem_type_Shape | 15;
	window['AscDFH'].historyitem_ShapeSetFLocksText             = window['AscDFH'].historyitem_type_Shape | 17;
	window['AscDFH'].historyitem_ShapeSetClientData             = window['AscDFH'].historyitem_type_Shape | 18;
	window['AscDFH'].historyitem_ShapeSetUseBgFill              = window['AscDFH'].historyitem_type_Shape | 19;

	window['AscDFH'].historyitem_OleSizeSelectionSetRange = window['AscDFH'].historyitem_type_OleSizeSelection | 1;

	window['AscDFH'].historyitem_DispUnitsSetBuiltInUnit  = window['AscDFH'].historyitem_type_DispUnits | 1;
	window['AscDFH'].historyitem_DispUnitsSetCustUnit     = window['AscDFH'].historyitem_type_DispUnits | 2;
	window['AscDFH'].historyitem_DispUnitsSetDispUnitsLbl = window['AscDFH'].historyitem_type_DispUnits | 3;
	window['AscDFH'].historyitem_DispUnitsSetParent       = window['AscDFH'].historyitem_type_DispUnits | 4;

	window['AscDFH'].historyitem_GroupShapeSetNvGrpSpPr     = window['AscDFH'].historyitem_type_GroupShape | 1;
	window['AscDFH'].historyitem_GroupShapeSetSpPr          = window['AscDFH'].historyitem_type_GroupShape | 2;
	window['AscDFH'].historyitem_GroupShapeAddToSpTree      = window['AscDFH'].historyitem_type_GroupShape | 3;
	window['AscDFH'].historyitem_GroupShapeSetParent        = window['AscDFH'].historyitem_type_GroupShape | 4;
	window['AscDFH'].historyitem_GroupShapeSetGroup         = window['AscDFH'].historyitem_type_GroupShape | 5;
	window['AscDFH'].historyitem_GroupShapeRemoveFromSpTree = window['AscDFH'].historyitem_type_GroupShape | 6;

	window['AscDFH'].historyitem_ImageShapeSetNvPicPr            = window['AscDFH'].historyitem_type_ImageShape | 1;
	window['AscDFH'].historyitem_ImageShapeSetSpPr               = window['AscDFH'].historyitem_type_ImageShape | 2;
	window['AscDFH'].historyitem_ImageShapeSetBlipFill           = window['AscDFH'].historyitem_type_ImageShape | 3;
	window['AscDFH'].historyitem_ImageShapeSetParent             = window['AscDFH'].historyitem_type_ImageShape | 4;
	window['AscDFH'].historyitem_ImageShapeSetGroup              = window['AscDFH'].historyitem_type_ImageShape | 5;
	window['AscDFH'].historyitem_ImageShapeSetStyle              = window['AscDFH'].historyitem_type_ImageShape | 6;
	window['AscDFH'].historyitem_ImageShapeSetData               = window['AscDFH'].historyitem_type_ImageShape | 7;
	window['AscDFH'].historyitem_ImageShapeSetApplicationId      = window['AscDFH'].historyitem_type_ImageShape | 8;
	window['AscDFH'].historyitem_ImageShapeSetPixSizes           = window['AscDFH'].historyitem_type_ImageShape | 9;
	window['AscDFH'].historyitem_ImageShapeSetObjectFile	     = window['AscDFH'].historyitem_type_ImageShape | 10;
	window['AscDFH'].historyitem_ImageShapeSetOleType   	     = window['AscDFH'].historyitem_type_ImageShape | 11;
	window['AscDFH'].historyitem_ImageShapeSetStartBinaryData    = window['AscDFH'].historyitem_type_ImageShape | 12;
	window['AscDFH'].historyitem_ImageShapeSetPartBinaryData     = window['AscDFH'].historyitem_type_ImageShape | 13;
	window['AscDFH'].historyitem_ImageShapeSetEndBinaryData      = window['AscDFH'].historyitem_type_ImageShape | 14;
	window['AscDFH'].historyitem_ImageShapeSetMathObject         = window['AscDFH'].historyitem_type_ImageShape | 15;
	window['AscDFH'].historyitem_ImageShapeSetDataLink           = window['AscDFH'].historyitem_type_ImageShape | 16;
	window['AscDFH'].historyitem_ImageShapeSetDrawAspect         = window['AscDFH'].historyitem_type_ImageShape | 17;
	window['AscDFH'].historyitem_ImageShapeLoadImagesfromContent = window['AscDFH'].historyitem_type_ImageShape | 18;

	window['AscDFH'].historyitem_GeometrySetParent      = window['AscDFH'].historyitem_type_Geometry | 1;
	window['AscDFH'].historyitem_GeometryAddAdj         = window['AscDFH'].historyitem_type_Geometry | 2;
	window['AscDFH'].historyitem_GeometryAddGuide       = window['AscDFH'].historyitem_type_Geometry | 3;
	window['AscDFH'].historyitem_GeometryAddCnx         = window['AscDFH'].historyitem_type_Geometry | 4;
	window['AscDFH'].historyitem_GeometryAddHandleXY    = window['AscDFH'].historyitem_type_Geometry | 5;
	window['AscDFH'].historyitem_GeometryAddHandlePolar = window['AscDFH'].historyitem_type_Geometry | 6;
	window['AscDFH'].historyitem_GeometryAddPath        = window['AscDFH'].historyitem_type_Geometry | 7;
	window['AscDFH'].historyitem_GeometryAddRect        = window['AscDFH'].historyitem_type_Geometry | 8;
	window['AscDFH'].historyitem_GeometrySetPreset      = window['AscDFH'].historyitem_type_Geometry | 9;

	window['AscDFH'].historyitem_PathSetStroke      = window['AscDFH'].historyitem_type_Path | 1;
	window['AscDFH'].historyitem_PathSetExtrusionOk = window['AscDFH'].historyitem_type_Path | 2;
	window['AscDFH'].historyitem_PathSetFill        = window['AscDFH'].historyitem_type_Path | 3;
	window['AscDFH'].historyitem_PathSetPathH       = window['AscDFH'].historyitem_type_Path | 4;
	window['AscDFH'].historyitem_PathSetPathW       = window['AscDFH'].historyitem_type_Path | 5;
	window['AscDFH'].historyitem_PathAddPathCommand = window['AscDFH'].historyitem_type_Path | 6;

	window['AscDFH'].historyitem_TextBodySetBodyPr   = window['AscDFH'].historyitem_type_TextBody | 1;
	window['AscDFH'].historyitem_TextBodySetLstStyle = window['AscDFH'].historyitem_type_TextBody | 2;
	window['AscDFH'].historyitem_TextBodySetContent  = window['AscDFH'].historyitem_type_TextBody | 3;
	window['AscDFH'].historyitem_TextBodySetParent   = window['AscDFH'].historyitem_type_TextBody | 4;

	window['AscDFH'].historyitem_CatAxSetAuto           = window['AscDFH'].historyitem_type_CatAx | 1;
	window['AscDFH'].historyitem_CatAxSetAxId           = window['AscDFH'].historyitem_type_CatAx | 2;
	window['AscDFH'].historyitem_CatAxSetAxPos          = window['AscDFH'].historyitem_type_CatAx | 3;
	window['AscDFH'].historyitem_CatAxSetCrossAx        = window['AscDFH'].historyitem_type_CatAx | 4;
	window['AscDFH'].historyitem_CatAxSetCrosses        = window['AscDFH'].historyitem_type_CatAx | 5;
	window['AscDFH'].historyitem_CatAxSetCrossesAt      = window['AscDFH'].historyitem_type_CatAx | 6;
	window['AscDFH'].historyitem_CatAxSetDelete         = window['AscDFH'].historyitem_type_CatAx | 7;
	window['AscDFH'].historyitem_CatAxSetLblAlgn        = window['AscDFH'].historyitem_type_CatAx | 8;
	window['AscDFH'].historyitem_CatAxSetLblOffset      = window['AscDFH'].historyitem_type_CatAx | 9;
	window['AscDFH'].historyitem_CatAxSetMajorGridlines = window['AscDFH'].historyitem_type_CatAx | 10;
	window['AscDFH'].historyitem_CatAxSetMajorTickMark  = window['AscDFH'].historyitem_type_CatAx | 11;
	window['AscDFH'].historyitem_CatAxSetMinorGridlines = window['AscDFH'].historyitem_type_CatAx | 12;
	window['AscDFH'].historyitem_CatAxSetMinorTickMark  = window['AscDFH'].historyitem_type_CatAx | 13;
	window['AscDFH'].historyitem_CatAxSetNoMultiLvlLbl  = window['AscDFH'].historyitem_type_CatAx | 14;
	window['AscDFH'].historyitem_CatAxSetNumFmt         = window['AscDFH'].historyitem_type_CatAx | 15;
	window['AscDFH'].historyitem_CatAxSetScaling        = window['AscDFH'].historyitem_type_CatAx | 16;
	window['AscDFH'].historyitem_CatAxSetSpPr           = window['AscDFH'].historyitem_type_CatAx | 17;
	window['AscDFH'].historyitem_CatAxSetTickLblPos     = window['AscDFH'].historyitem_type_CatAx | 18;
	window['AscDFH'].historyitem_CatAxSetTickLblSkip    = window['AscDFH'].historyitem_type_CatAx | 19;
	window['AscDFH'].historyitem_CatAxSetTickMarkSkip   = window['AscDFH'].historyitem_type_CatAx | 20;
	window['AscDFH'].historyitem_CatAxSetTitle          = window['AscDFH'].historyitem_type_CatAx | 21;
	window['AscDFH'].historyitem_CatAxSetTxPr           = window['AscDFH'].historyitem_type_CatAx | 22;

	window['AscDFH'].historyitem_ValAxSetAxId           = window['AscDFH'].historyitem_type_ValAx | 1;
	window['AscDFH'].historyitem_ValAxSetAxPos          = window['AscDFH'].historyitem_type_ValAx | 2;
	window['AscDFH'].historyitem_ValAxSetCrossAx        = window['AscDFH'].historyitem_type_ValAx | 3;
	window['AscDFH'].historyitem_ValAxSetCrossBetween   = window['AscDFH'].historyitem_type_ValAx | 4;
	window['AscDFH'].historyitem_ValAxSetCrosses        = window['AscDFH'].historyitem_type_ValAx | 5;
	window['AscDFH'].historyitem_ValAxSetCrossesAt      = window['AscDFH'].historyitem_type_ValAx | 6;
	window['AscDFH'].historyitem_ValAxSetDelete         = window['AscDFH'].historyitem_type_ValAx | 7;
	window['AscDFH'].historyitem_ValAxSetDispUnits      = window['AscDFH'].historyitem_type_ValAx | 8;
	window['AscDFH'].historyitem_ValAxSetMajorGridlines = window['AscDFH'].historyitem_type_ValAx | 9;
	window['AscDFH'].historyitem_ValAxSetMajorTickMark  = window['AscDFH'].historyitem_type_ValAx | 10;
	window['AscDFH'].historyitem_ValAxSetMajorUnit      = window['AscDFH'].historyitem_type_ValAx | 11;
	window['AscDFH'].historyitem_ValAxSetMinorGridlines = window['AscDFH'].historyitem_type_ValAx | 12;
	window['AscDFH'].historyitem_ValAxSetMinorTickMark  = window['AscDFH'].historyitem_type_ValAx | 13;
	window['AscDFH'].historyitem_ValAxSetMinorUnit      = window['AscDFH'].historyitem_type_ValAx | 14;
	window['AscDFH'].historyitem_ValAxSetNumFmt         = window['AscDFH'].historyitem_type_ValAx | 15;
	window['AscDFH'].historyitem_ValAxSetScaling        = window['AscDFH'].historyitem_type_ValAx | 16;
	window['AscDFH'].historyitem_ValAxSetSpPr           = window['AscDFH'].historyitem_type_ValAx | 17;
	window['AscDFH'].historyitem_ValAxSetTickLblPos     = window['AscDFH'].historyitem_type_ValAx | 18;
	window['AscDFH'].historyitem_ValAxSetTitle          = window['AscDFH'].historyitem_type_ValAx | 19;
	window['AscDFH'].historyitem_ValAxSetTxPr           = window['AscDFH'].historyitem_type_ValAx | 20;

	window['AscDFH'].historyitem_WrapPolygonSetEdited    = window['AscDFH'].historyitem_type_WrapPolygon | 1;
	window['AscDFH'].historyitem_WrapPolygonSetRelPoints = window['AscDFH'].historyitem_type_WrapPolygon | 2;
	window['AscDFH'].historyitem_WrapPolygonSetWrapSide  = window['AscDFH'].historyitem_type_WrapPolygon | 3;

	window['AscDFH'].historyitem_DateAxAuto           = window['AscDFH'].historyitem_type_DateAx | 1;
	window['AscDFH'].historyitem_DateAxAxId           = window['AscDFH'].historyitem_type_DateAx | 2;
	window['AscDFH'].historyitem_DateAxAxPos          = window['AscDFH'].historyitem_type_DateAx | 3;
	window['AscDFH'].historyitem_DateAxBaseTimeUnit   = window['AscDFH'].historyitem_type_DateAx | 4;
	window['AscDFH'].historyitem_DateAxCrossAx        = window['AscDFH'].historyitem_type_DateAx | 5;
	window['AscDFH'].historyitem_DateAxCrosses        = window['AscDFH'].historyitem_type_DateAx | 6;
	window['AscDFH'].historyitem_DateAxCrossesAt      = window['AscDFH'].historyitem_type_DateAx | 7;
	window['AscDFH'].historyitem_DateAxDelete         = window['AscDFH'].historyitem_type_DateAx | 8;
	window['AscDFH'].historyitem_DateAxLblOffset      = window['AscDFH'].historyitem_type_DateAx | 9;
	window['AscDFH'].historyitem_DateAxMajorGridlines = window['AscDFH'].historyitem_type_DateAx | 10;
	window['AscDFH'].historyitem_DateAxMajorTickMark  = window['AscDFH'].historyitem_type_DateAx | 11;
	window['AscDFH'].historyitem_DateAxMajorTimeUnit  = window['AscDFH'].historyitem_type_DateAx | 12;
	window['AscDFH'].historyitem_DateAxMajorUnit      = window['AscDFH'].historyitem_type_DateAx | 13;
	window['AscDFH'].historyitem_DateAxMinorGridlines = window['AscDFH'].historyitem_type_DateAx | 14;
	window['AscDFH'].historyitem_DateAxMinorTickMark  = window['AscDFH'].historyitem_type_DateAx | 15;
	window['AscDFH'].historyitem_DateAxMinorTimeUnit  = window['AscDFH'].historyitem_type_DateAx | 16;
	window['AscDFH'].historyitem_DateAxMinorUnit      = window['AscDFH'].historyitem_type_DateAx | 17;
	window['AscDFH'].historyitem_DateAxNumFmt         = window['AscDFH'].historyitem_type_DateAx | 18;
	window['AscDFH'].historyitem_DateAxScaling        = window['AscDFH'].historyitem_type_DateAx | 19;
	window['AscDFH'].historyitem_DateAxSpPr           = window['AscDFH'].historyitem_type_DateAx | 20;
	window['AscDFH'].historyitem_DateAxTickLblPos     = window['AscDFH'].historyitem_type_DateAx | 21;
	window['AscDFH'].historyitem_DateAxTitle          = window['AscDFH'].historyitem_type_DateAx | 22;
	window['AscDFH'].historyitem_DateAxTxPr           = window['AscDFH'].historyitem_type_DateAx | 23;

	window['AscDFH'].historyitem_SerAxSetAxId           = window['AscDFH'].historyitem_type_SerAx | 1;
	window['AscDFH'].historyitem_SerAxSetAxPos          = window['AscDFH'].historyitem_type_SerAx | 2;
	window['AscDFH'].historyitem_SerAxSetCrossAx        = window['AscDFH'].historyitem_type_SerAx | 3;
	window['AscDFH'].historyitem_SerAxSetCrosses        = window['AscDFH'].historyitem_type_SerAx | 4;
	window['AscDFH'].historyitem_SerAxSetCrossesAt      = window['AscDFH'].historyitem_type_SerAx | 5;
	window['AscDFH'].historyitem_SerAxSetDelete         = window['AscDFH'].historyitem_type_SerAx | 6;
	window['AscDFH'].historyitem_SerAxSetMajorGridlines = window['AscDFH'].historyitem_type_SerAx | 7;
	window['AscDFH'].historyitem_SerAxSetMajorTickMark  = window['AscDFH'].historyitem_type_SerAx | 8;
	window['AscDFH'].historyitem_SerAxSetMinorGridlines = window['AscDFH'].historyitem_type_SerAx | 9;
	window['AscDFH'].historyitem_SerAxSetMinorTickMark  = window['AscDFH'].historyitem_type_SerAx | 10;
	window['AscDFH'].historyitem_SerAxSetNumFmt         = window['AscDFH'].historyitem_type_SerAx | 11;
	window['AscDFH'].historyitem_SerAxSetScaling        = window['AscDFH'].historyitem_type_SerAx | 12;
	window['AscDFH'].historyitem_SerAxSetSpPr           = window['AscDFH'].historyitem_type_SerAx | 13;
	window['AscDFH'].historyitem_SerAxSetTickLblPos     = window['AscDFH'].historyitem_type_SerAx | 14;
	window['AscDFH'].historyitem_SerAxSetTickLblSkip    = window['AscDFH'].historyitem_type_SerAx | 15;
	window['AscDFH'].historyitem_SerAxSetTickMarkSkip   = window['AscDFH'].historyitem_type_SerAx | 16;
	window['AscDFH'].historyitem_SerAxSetTitle          = window['AscDFH'].historyitem_type_SerAx | 17;
	window['AscDFH'].historyitem_SerAxSetTxPr           = window['AscDFH'].historyitem_type_SerAx | 18;

	window['AscDFH'].historyitem_Title_SetLayout  = window['AscDFH'].historyitem_type_Title | 1;
	window['AscDFH'].historyitem_Title_SetOverlay = window['AscDFH'].historyitem_type_Title | 2;
	window['AscDFH'].historyitem_Title_SetSpPr    = window['AscDFH'].historyitem_type_Title | 3;
	window['AscDFH'].historyitem_Title_SetTx      = window['AscDFH'].historyitem_type_Title | 4;
	window['AscDFH'].historyitem_Title_SetTxPr    = window['AscDFH'].historyitem_type_Title | 5;
	window['AscDFH'].historyitem_Title_SetAlign   = window['AscDFH'].historyitem_type_Title | 6;
	window['AscDFH'].historyitem_Title_SetPos     = window['AscDFH'].historyitem_type_Title | 7;

	window['AscDFH'].historyitem_SlideSetComments       = window['AscDFH'].historyitem_type_Slide | 1;
	window['AscDFH'].historyitem_SlideSetShow           = window['AscDFH'].historyitem_type_Slide | 2;
	window['AscDFH'].historyitem_SlideSetShowPhAnim     = window['AscDFH'].historyitem_type_Slide | 3;
	window['AscDFH'].historyitem_SlideSetShowMasterSp   = window['AscDFH'].historyitem_type_Slide | 4;
	window['AscDFH'].historyitem_SlideSetLayout         = window['AscDFH'].historyitem_type_Slide | 5;
	window['AscDFH'].historyitem_SlideSetNum            = window['AscDFH'].historyitem_type_Slide | 6;
	window['AscDFH'].historyitem_SlideSetTransition     = window['AscDFH'].historyitem_type_Slide | 7;
	window['AscDFH'].historyitem_SlideSetSize           = window['AscDFH'].historyitem_type_Slide | 8;
	window['AscDFH'].historyitem_SlideSetBg             = window['AscDFH'].historyitem_type_Slide | 9;
	window['AscDFH'].historyitem_SlideSetLocks          = window['AscDFH'].historyitem_type_Slide | 10;
	window['AscDFH'].historyitem_SlideRemoveFromSpTree  = window['AscDFH'].historyitem_type_Slide | 11;
	window['AscDFH'].historyitem_SlideAddToSpTree       = window['AscDFH'].historyitem_type_Slide | 12;
	window['AscDFH'].historyitem_SlideSetCSldName       = window['AscDFH'].historyitem_type_Slide | 13;
	window['AscDFH'].historyitem_SlideSetClrMapOverride = window['AscDFH'].historyitem_type_Slide | 14;
	window['AscDFH'].historyitem_SlideSetNotes          = window['AscDFH'].historyitem_type_Slide | 15;
	window['AscDFH'].historyitem_SlideSetTiming         = window['AscDFH'].historyitem_type_Slide | 16;

	window['AscDFH'].historyitem_SlideLayoutSetMaster         = window['AscDFH'].historyitem_type_SlideLayout | 1;
	window['AscDFH'].historyitem_SlideLayoutSetMatchingName   = window['AscDFH'].historyitem_type_SlideLayout | 2;
	window['AscDFH'].historyitem_SlideLayoutSetType           = window['AscDFH'].historyitem_type_SlideLayout | 3;
	window['AscDFH'].historyitem_SlideLayoutSetBg             = window['AscDFH'].historyitem_type_SlideLayout | 4;
	window['AscDFH'].historyitem_SlideLayoutSetCSldName       = window['AscDFH'].historyitem_type_SlideLayout | 5;
	window['AscDFH'].historyitem_SlideLayoutSetShow           = window['AscDFH'].historyitem_type_SlideLayout | 6;
	window['AscDFH'].historyitem_SlideLayoutSetShowPhAnim     = window['AscDFH'].historyitem_type_SlideLayout | 7;
	window['AscDFH'].historyitem_SlideLayoutSetShowMasterSp   = window['AscDFH'].historyitem_type_SlideLayout | 8;
	window['AscDFH'].historyitem_SlideLayoutSetClrMapOverride = window['AscDFH'].historyitem_type_SlideLayout | 9;
	window['AscDFH'].historyitem_SlideLayoutAddToSpTree       = window['AscDFH'].historyitem_type_SlideLayout | 10;
	window['AscDFH'].historyitem_SlideLayoutSetSize           = window['AscDFH'].historyitem_type_SlideLayout | 11;
	window['AscDFH'].historyitem_SlideLayoutSetHF             = window['AscDFH'].historyitem_type_SlideLayout | 12;
	window['AscDFH'].historyitem_SlideLayoutSetTiming         = window['AscDFH'].historyitem_type_SlideLayout | 13;
	window['AscDFH'].historyitem_SlideLayoutSetTransition     = window['AscDFH'].historyitem_type_SlideLayout | 14;
	window['AscDFH'].historyitem_SlideLayoutRemoveFromSpTree  = window['AscDFH'].historyitem_type_SlideLayout | 15;
	window['AscDFH'].historyitem_SlideLayoutSetPreserve       = window['AscDFH'].historyitem_type_SlideLayout | 16;

	window['AscDFH'].historyitem_SlideMasterAddToSpTree       = window['AscDFH'].historyitem_type_SlideMaster | 1;
	window['AscDFH'].historyitem_SlideMasterSetTheme          = window['AscDFH'].historyitem_type_SlideMaster | 2;
	window['AscDFH'].historyitem_SlideMasterSetBg             = window['AscDFH'].historyitem_type_SlideMaster | 3;
	window['AscDFH'].historyitem_SlideMasterSetTxStyles       = window['AscDFH'].historyitem_type_SlideMaster | 4;
	window['AscDFH'].historyitem_SlideMasterSetCSldName       = window['AscDFH'].historyitem_type_SlideMaster | 5;
	window['AscDFH'].historyitem_SlideMasterSetClrMapOverride = window['AscDFH'].historyitem_type_SlideMaster | 6;
	window['AscDFH'].historyitem_SlideMasterAddLayout         = window['AscDFH'].historyitem_type_SlideMaster | 7;
	window['AscDFH'].historyitem_SlideMasterSetThemeIndex     = window['AscDFH'].historyitem_type_SlideMaster | 8;
	window['AscDFH'].historyitem_SlideMasterSetSize           = window['AscDFH'].historyitem_type_SlideMaster | 9;
	window['AscDFH'].historyitem_SlideMasterSetHF             = window['AscDFH'].historyitem_type_SlideMaster | 10;
	window['AscDFH'].historyitem_SlideMasterSetTiming         = window['AscDFH'].historyitem_type_SlideMaster | 11;
	window['AscDFH'].historyitem_SlideMasterSetTransition     = window['AscDFH'].historyitem_type_SlideMaster | 12;
	window['AscDFH'].historyitem_SlideMasterRemoveLayout      = window['AscDFH'].historyitem_type_SlideMaster | 13;
	window['AscDFH'].historyitem_SlideMasterRemoveFromSpTree  = window['AscDFH'].historyitem_type_SlideMaster | 14;
	window['AscDFH'].historyitem_SlideMasterSetPreserve       = window['AscDFH'].historyitem_type_SlideMaster | 15;

	window['AscDFH'].historyitem_SlideCommentsAddComment    = window['AscDFH'].historyitem_type_SlideComments | 1;
	window['AscDFH'].historyitem_SlideCommentsRemoveComment = window['AscDFH'].historyitem_type_SlideComments | 2;

	window['AscDFH'].historyitem_PropLockerSetId = window['AscDFH'].historyitem_type_PropLocker | 1;

	window['AscDFH'].historyitem_ThemeSetColorScheme        = window['AscDFH'].historyitem_type_Theme | 1;
	window['AscDFH'].historyitem_ThemeSetFontScheme         = window['AscDFH'].historyitem_type_Theme | 2;
	window['AscDFH'].historyitem_ThemeSetFmtScheme          = window['AscDFH'].historyitem_type_Theme | 3;
	window['AscDFH'].historyitem_ThemeSetName               = window['AscDFH'].historyitem_type_Theme | 4;
	window['AscDFH'].historyitem_ThemeSetIsThemeOverride    = window['AscDFH'].historyitem_type_Theme | 5;
	window['AscDFH'].historyitem_ThemeSetSpDef              = window['AscDFH'].historyitem_type_Theme | 6;
	window['AscDFH'].historyitem_ThemeSetLnDef              = window['AscDFH'].historyitem_type_Theme | 7;
	window['AscDFH'].historyitem_ThemeSetTxDef              = window['AscDFH'].historyitem_type_Theme | 8;
	window['AscDFH'].historyitem_ThemeAddExtraClrScheme     = window['AscDFH'].historyitem_type_Theme | 9;
	window['AscDFH'].historyitem_ThemeRemoveExtraClrScheme  = window['AscDFH'].historyitem_type_Theme | 10;

	window['AscDFH'].historyitem_GraphicFrameSetSpPr          = window['AscDFH'].historyitem_type_GraphicFrame | 1;
	window['AscDFH'].historyitem_GraphicFrameSetGraphicObject = window['AscDFH'].historyitem_type_GraphicFrame | 2;
	window['AscDFH'].historyitem_GraphicFrameSetSetNvSpPr     = window['AscDFH'].historyitem_type_GraphicFrame | 3;
	window['AscDFH'].historyitem_GraphicFrameSetSetParent     = window['AscDFH'].historyitem_type_GraphicFrame | 4;
	window['AscDFH'].historyitem_GraphicFrameSetSetGroup      = window['AscDFH'].historyitem_type_GraphicFrame | 5;


    window['AscDFH'].historyitem_Sparkline_Type = window['AscDFH'].historyitem_type_Sparkline | 1;
    window['AscDFH'].historyitem_Sparkline_LineWeight = window['AscDFH'].historyitem_type_Sparkline | 2;
    window['AscDFH'].historyitem_Sparkline_DisplayEmptyCellsAs = window['AscDFH'].historyitem_type_Sparkline | 3;
    window['AscDFH'].historyitem_Sparkline_Markers = window['AscDFH'].historyitem_type_Sparkline | 4;
    window['AscDFH'].historyitem_Sparkline_High = window['AscDFH'].historyitem_type_Sparkline | 5;
    window['AscDFH'].historyitem_Sparkline_Low = window['AscDFH'].historyitem_type_Sparkline | 6;
    window['AscDFH'].historyitem_Sparkline_First = window['AscDFH'].historyitem_type_Sparkline | 7;
    window['AscDFH'].historyitem_Sparkline_Last = window['AscDFH'].historyitem_type_Sparkline | 8;
    window['AscDFH'].historyitem_Sparkline_Negative = window['AscDFH'].historyitem_type_Sparkline | 9;
    window['AscDFH'].historyitem_Sparkline_DisplayXAxis = window['AscDFH'].historyitem_type_Sparkline | 10;
    window['AscDFH'].historyitem_Sparkline_DisplayHidden = window['AscDFH'].historyitem_type_Sparkline | 11;
    window['AscDFH'].historyitem_Sparkline_MinAxisType = window['AscDFH'].historyitem_type_Sparkline | 12;
    window['AscDFH'].historyitem_Sparkline_MaxAxisType = window['AscDFH'].historyitem_type_Sparkline | 13;
    window['AscDFH'].historyitem_Sparkline_RightToLeft = window['AscDFH'].historyitem_type_Sparkline | 14;
    window['AscDFH'].historyitem_Sparkline_ManualMax = window['AscDFH'].historyitem_type_Sparkline | 15;
    window['AscDFH'].historyitem_Sparkline_ManualMin = window['AscDFH'].historyitem_type_Sparkline | 16;
    window['AscDFH'].historyitem_Sparkline_DateAxis = window['AscDFH'].historyitem_type_Sparkline | 17;
    window['AscDFH'].historyitem_Sparkline_ColorSeries = window['AscDFH'].historyitem_type_Sparkline | 18;
    window['AscDFH'].historyitem_Sparkline_ColorNegative = window['AscDFH'].historyitem_type_Sparkline | 19;
    window['AscDFH'].historyitem_Sparkline_ColorAxis = window['AscDFH'].historyitem_type_Sparkline | 20;
    window['AscDFH'].historyitem_Sparkline_ColorMarkers = window['AscDFH'].historyitem_type_Sparkline | 21;
    window['AscDFH'].historyitem_Sparkline_ColorFirst = window['AscDFH'].historyitem_type_Sparkline | 22;
    window['AscDFH'].historyitem_Sparkline_colorLast = window['AscDFH'].historyitem_type_Sparkline | 23;
    window['AscDFH'].historyitem_Sparkline_ColorHigh = window['AscDFH'].historyitem_type_Sparkline | 24;
    window['AscDFH'].historyitem_Sparkline_ColorLow = window['AscDFH'].historyitem_type_Sparkline | 25;
    window['AscDFH'].historyitem_Sparkline_F = window['AscDFH'].historyitem_type_Sparkline | 26;
    window['AscDFH'].historyitem_Sparkline_ChangeData = window['AscDFH'].historyitem_type_Sparkline | 27;
    window['AscDFH'].historyitem_Sparkline_RemoveData = window['AscDFH'].historyitem_type_Sparkline | 28;
    window['AscDFH'].historyitem_Sparkline_RemoveSparkline = window['AscDFH'].historyitem_type_Sparkline | 29;
    window['AscDFH'].historyitem_Sparkline_Worksheet = window['AscDFH'].historyitem_type_Sparkline | 30;


    window['AscDFH'].historyitem_NotesMasterSetHF          = window['AscDFH'].historyitem_type_NotesMaster | 1;
    window['AscDFH'].historyitem_NotesMasterSetNotesStyle  = window['AscDFH'].historyitem_type_NotesMaster | 2;
    window['AscDFH'].historyitem_NotesMasterSetNotesTheme  = window['AscDFH'].historyitem_type_NotesMaster | 3;
    window['AscDFH'].historyitem_NotesMasterAddToSpTree    = window['AscDFH'].historyitem_type_NotesMaster | 4;
    window['AscDFH'].historyitem_NotesMasterRemoveFromTree = window['AscDFH'].historyitem_type_NotesMaster | 5;
    window['AscDFH'].historyitem_NotesMasterSetBg          = window['AscDFH'].historyitem_type_NotesMaster | 6;
    window['AscDFH'].historyitem_NotesMasterAddToNotesLst  = window['AscDFH'].historyitem_type_NotesMaster | 7;
    window['AscDFH'].historyitem_NotesMasterSetName        = window['AscDFH'].historyitem_type_NotesMaster | 8;
    window['AscDFH'].historyitem_NotesMasterSetClrMap      = window['AscDFH'].historyitem_type_NotesMaster | 9;


    window['AscDFH'].historyitem_NotesSetClrMap           = window['AscDFH'].historyitem_type_Notes | 1;
    window['AscDFH'].historyitem_NotesSetShowMasterPhAnim = window['AscDFH'].historyitem_type_Notes | 2;
    window['AscDFH'].historyitem_NotesSetShowMasterSp     = window['AscDFH'].historyitem_type_Notes | 3;
    window['AscDFH'].historyitem_NotesAddToSpTree         = window['AscDFH'].historyitem_type_Notes | 4;
    window['AscDFH'].historyitem_NotesRemoveFromTree      = window['AscDFH'].historyitem_type_Notes | 5;
    window['AscDFH'].historyitem_NotesSetBg               = window['AscDFH'].historyitem_type_Notes | 6;
    window['AscDFH'].historyitem_NotesSetName             = window['AscDFH'].historyitem_type_Notes | 7;
    window['AscDFH'].historyitem_NotesSetSlide            = window['AscDFH'].historyitem_type_Notes | 8;
    window['AscDFH'].historyitem_NotesSetNotesMaster      = window['AscDFH'].historyitem_type_Notes | 9;

    window['AscDFH'].historyitem_PresentationSectionSetName       = window['AscDFH'].historyitem_type_PresentationSection | 1;
    window['AscDFH'].historyitem_PresentationSectionSetGuid       = window['AscDFH'].historyitem_type_PresentationSection | 2;
    window['AscDFH'].historyitem_PresentationSectionSetStartIndex = window['AscDFH'].historyitem_type_PresentationSection | 3;

	window['AscDFH'].historyitem_RelSizeAnchorFromX  = window['AscDFH'].historyitem_type_RelSizeAnchor | 1;
	window['AscDFH'].historyitem_RelSizeAnchorFromY  = window['AscDFH'].historyitem_type_RelSizeAnchor | 2;
	window['AscDFH'].historyitem_RelSizeAnchorToX    = window['AscDFH'].historyitem_type_RelSizeAnchor | 3;
	window['AscDFH'].historyitem_RelSizeAnchorToY    = window['AscDFH'].historyitem_type_RelSizeAnchor | 4;
	window['AscDFH'].historyitem_RelSizeAnchorObject = window['AscDFH'].historyitem_type_RelSizeAnchor | 5;
	window['AscDFH'].historyitem_RelSizeAnchorParent = window['AscDFH'].historyitem_type_RelSizeAnchor | 6;

	window['AscDFH'].historyitem_AbsSizeAnchorFromX   = window['AscDFH'].historyitem_type_AbsSizeAnchor | 1;
	window['AscDFH'].historyitem_AbsSizeAnchorFromY   = window['AscDFH'].historyitem_type_AbsSizeAnchor | 2;
	window['AscDFH'].historyitem_AbsSizeAnchorExtX    = window['AscDFH'].historyitem_type_AbsSizeAnchor | 3;
	window['AscDFH'].historyitem_AbsSizeAnchorExtY    = window['AscDFH'].historyitem_type_AbsSizeAnchor | 4;
	window['AscDFH'].historyitem_AbsSizeAnchorObject  = window['AscDFH'].historyitem_type_AbsSizeAnchor | 5;
	window['AscDFH'].historyitem_AbsSizeAnchorParent  = window['AscDFH'].historyitem_type_AbsSizeAnchor | 6;

	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений класса CDocumentMacros
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_DocumentMacros_Data = window['AscDFH'].historyitem_type_DocumentMacros | 1;

	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений класса CCore
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_CoreProperties = window['AscDFH'].historyitem_type_Core | 1;

	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений класса CSlicerView
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_SlicerViewName = window['AscDFH'].historyitem_type_SlicerView | 1;

	AscDFH.historyitem_TimingBldLst = AscDFH.historyitem_type_Timing | 1;
	AscDFH.historyitem_TimingTnLst  = AscDFH.historyitem_type_Timing | 2;

	AscDFH.historyitem_CommonTimingListAdd    = AscDFH.historyitem_type_CommonTimingList | 1;
	AscDFH.historyitem_CommonTimingListRemove = AscDFH.historyitem_type_CommonTimingList | 2;

	AscDFH.historyitem_ObjectTargetSpid = AscDFH.historyitem_type_ObjectTarget | 1;

	AscDFH.historyitem_BldBaseGrpId    = AscDFH.historyitem_type_BldBase | 1;
	AscDFH.historyitem_BldBaseUIExpand = AscDFH.historyitem_type_BldBase | 2;

	AscDFH.historyitem_BldDgmBld = AscDFH.historyitem_type_BldDgm | 1;

	AscDFH.historyitem_BldGraphicBldAsOne = AscDFH.historyitem_type_BldGraphic | 1;
	AscDFH.historyitem_BldGraphicBldSub   = AscDFH.historyitem_type_BldGraphic | 2;

	AscDFH.historyitem_BldOleChartAnimBg = AscDFH.historyitem_type_BldOleChart | 1;

	AscDFH.historyitem_BldPTmplLst          = AscDFH.historyitem_type_BldP | 1;
	AscDFH.historyitem_BldPAdvAuto          = AscDFH.historyitem_type_BldP | 2;
	AscDFH.historyitem_BldPAutoUpdateAnimBg = AscDFH.historyitem_type_BldP | 3;
	AscDFH.historyitem_BldPBldLvl           = AscDFH.historyitem_type_BldP | 4;
	AscDFH.historyitem_BldPBuild            = AscDFH.historyitem_type_BldP | 5;
	AscDFH.historyitem_BldPRev              = AscDFH.historyitem_type_BldP | 6;

	AscDFH.historyitem_BldSubChart = AscDFH.historyitem_type_BldSub    | 1;
	AscDFH.historyitem_BldSubAnimBg = AscDFH.historyitem_type_BldSub   | 2;
	AscDFH.historyitem_BldSubBldChart = AscDFH.historyitem_type_BldSub | 3;
	AscDFH.historyitem_BldSubBldDgm = AscDFH.historyitem_type_BldSub   | 4;
	AscDFH.historyitem_BldSubRev = AscDFH.historyitem_type_BldSub      | 5;

	AscDFH.historyitem_DirTransitionDir = AscDFH.historyitem_type_DirTransition | 1;

	AscDFH.historyitem_OptBlackTransitionThruBlk = AscDFH.historyitem_type_OptBlackTransition | 1;


	AscDFH.historyitem_GraphicElDgmId          = AscDFH.historyitem_type_GraphicEl | 1;
	AscDFH.historyitem_GraphicElDgmBuildStep   = AscDFH.historyitem_type_GraphicEl | 2;
	AscDFH.historyitem_GraphicElChartBuildStep = AscDFH.historyitem_type_GraphicEl | 3;
	AscDFH.historyitem_GraphicElSeriesIdx      = AscDFH.historyitem_type_GraphicEl | 4;
	AscDFH.historyitem_GraphicElCategoryIdx    = AscDFH.historyitem_type_GraphicEl | 5;

	AscDFH.historyitem_IndexRgSt  = AscDFH.historyitem_type_IndexRg | 1;
	AscDFH.historyitem_IndexRgEnd = AscDFH.historyitem_type_IndexRg | 2;

	AscDFH.historyitem_TmplLvl   = AscDFH.historyitem_type_Tmpl | 1;
	AscDFH.historyitem_TmplTnLst = AscDFH.historyitem_type_Tmpl | 2;

	AscDFH.historyitem_AnimCBhvr     = AscDFH.historyitem_type_Anim | 1;
	AscDFH.historyitem_AnimTavLst    = AscDFH.historyitem_type_Anim | 2;
	AscDFH.historyitem_AnimBy        = AscDFH.historyitem_type_Anim | 3;
	AscDFH.historyitem_AnimCalcmode  = AscDFH.historyitem_type_Anim | 4;
	AscDFH.historyitem_AnimFrom      = AscDFH.historyitem_type_Anim | 5;
	AscDFH.historyitem_AnimTo        = AscDFH.historyitem_type_Anim | 6;
	AscDFH.historyitem_AnimValueType = AscDFH.historyitem_type_Anim | 7;

	AscDFH.historyitem_CBhvrAttrNameLst = AscDFH.historyitem_type_CBhvr | 1;
	AscDFH.historyitem_CBhvrCTn         = AscDFH.historyitem_type_CBhvr | 2;
	AscDFH.historyitem_CBhvrTgtEl       = AscDFH.historyitem_type_CBhvr | 3;
	AscDFH.historyitem_CBhvrAccumulate  = AscDFH.historyitem_type_CBhvr | 4;
	AscDFH.historyitem_CBhvrAdditive    = AscDFH.historyitem_type_CBhvr | 5;
	AscDFH.historyitem_CBhvrBy          = AscDFH.historyitem_type_CBhvr | 6;
	AscDFH.historyitem_CBhvrFrom        = AscDFH.historyitem_type_CBhvr | 7;
	AscDFH.historyitem_CBhvrOverride    = AscDFH.historyitem_type_CBhvr | 8;
	AscDFH.historyitem_CBhvrRctx        = AscDFH.historyitem_type_CBhvr | 9;
	AscDFH.historyitem_CBhvrTo          = AscDFH.historyitem_type_CBhvr | 10;
	AscDFH.historyitem_CBhvrXfrmType    = AscDFH.historyitem_type_CBhvr | 11;

	AscDFH.historyitem_CTnChildTnLst    = AscDFH.historyitem_type_CTn | 1;
	AscDFH.historyitem_CTnEndCondLst    = AscDFH.historyitem_type_CTn | 2;
	AscDFH.historyitem_CTnEndSync       = AscDFH.historyitem_type_CTn | 3;
	AscDFH.historyitem_CTnIterate       = AscDFH.historyitem_type_CTn | 4;
	AscDFH.historyitem_CTnStCondLst     = AscDFH.historyitem_type_CTn | 5;
	AscDFH.historyitem_CTnSubTnLst      = AscDFH.historyitem_type_CTn | 6;
	AscDFH.historyitem_CTnAccel         = AscDFH.historyitem_type_CTn | 7;
	AscDFH.historyitem_CTnAfterEffect   = AscDFH.historyitem_type_CTn | 8;
	AscDFH.historyitem_CTnAutoRev       = AscDFH.historyitem_type_CTn | 9;
	AscDFH.historyitem_CTnBldLvl        = AscDFH.historyitem_type_CTn | 10;
	AscDFH.historyitem_CTnDecel         = AscDFH.historyitem_type_CTn | 11;
	AscDFH.historyitem_CTnDisplay       = AscDFH.historyitem_type_CTn | 12;
	AscDFH.historyitem_CTnDur           = AscDFH.historyitem_type_CTn | 13;
	AscDFH.historyitem_CTnEvtFilter     = AscDFH.historyitem_type_CTn | 14;
	AscDFH.historyitem_CTnFill          = AscDFH.historyitem_type_CTn | 15;
	AscDFH.historyitem_CTnGrpId         = AscDFH.historyitem_type_CTn | 16;
	AscDFH.historyitem_CTnId            = AscDFH.historyitem_type_CTn | 17;
	AscDFH.historyitem_CTnMasterRel     = AscDFH.historyitem_type_CTn | 18;
	AscDFH.historyitem_CTnNodePh        = AscDFH.historyitem_type_CTn | 19;
	AscDFH.historyitem_CTnNodeType      = AscDFH.historyitem_type_CTn | 20;
	AscDFH.historyitem_CTnPresetClass   = AscDFH.historyitem_type_CTn | 21;
	AscDFH.historyitem_CTnPresetID      = AscDFH.historyitem_type_CTn | 22;
	AscDFH.historyitem_CTnPresetSubtype = AscDFH.historyitem_type_CTn | 23;
	AscDFH.historyitem_CTnRepeatCount   = AscDFH.historyitem_type_CTn | 24;
	AscDFH.historyitem_CTnRepeatDur     = AscDFH.historyitem_type_CTn | 25;
	AscDFH.historyitem_CTnRestart       = AscDFH.historyitem_type_CTn | 26;
	AscDFH.historyitem_CTnSpd           = AscDFH.historyitem_type_CTn | 27;
	AscDFH.historyitem_CTnSyncBehavior  = AscDFH.historyitem_type_CTn | 28;
	AscDFH.historyitem_CTnTmFilter      = AscDFH.historyitem_type_CTn | 29;

	AscDFH.historyitem_CondRtn   = AscDFH.historyitem_type_Cond | 1;
	AscDFH.historyitem_CondTgtEl = AscDFH.historyitem_type_Cond | 2;
	AscDFH.historyitem_CondTn    = AscDFH.historyitem_type_Cond | 3;
	AscDFH.historyitem_CondDelay = AscDFH.historyitem_type_Cond | 4;
	AscDFH.historyitem_CondEvt   = AscDFH.historyitem_type_Cond | 5;

	AscDFH.historyitem_RtnVal = AscDFH.historyitem_type_Rtn | 1;

	AscDFH.historyitem_TgtElInkTgt = AscDFH.historyitem_type_TgtEl | 1;
	AscDFH.historyitem_TgtElSldTgt = AscDFH.historyitem_type_TgtEl | 2;
	AscDFH.historyitem_TgtElSndTgt = AscDFH.historyitem_type_TgtEl | 3;
	AscDFH.historyitem_TgtElSpTgt  = AscDFH.historyitem_type_TgtEl | 4;

	AscDFH.historyitem_SndTgtEmbed   = AscDFH.historyitem_type_SndTgt | 1;
	AscDFH.historyitem_SndTgtName    = AscDFH.historyitem_type_SndTgt | 2;
	AscDFH.historyitem_SndTgtBuiltIn = AscDFH.historyitem_type_SndTgt | 3;

	AscDFH.historyitem_SpTgtBg         = AscDFH.historyitem_type_SpTgt | 1;
	AscDFH.historyitem_SpTgtGraphicEl  = AscDFH.historyitem_type_SpTgt | 2;
	AscDFH.historyitem_SpTgtOleChartEl = AscDFH.historyitem_type_SpTgt | 3;
	AscDFH.historyitem_SpTgtSubSpId    = AscDFH.historyitem_type_SpTgt | 4;
	AscDFH.historyitem_SpTgtTxEl       = AscDFH.historyitem_type_SpTgt | 5;

	AscDFH.historyitem_IterateDataTmAbs     = AscDFH.historyitem_type_IterateData | 1;
	AscDFH.historyitem_IterateDataTmPct     = AscDFH.historyitem_type_IterateData | 2;
	AscDFH.historyitem_IterateDataBackwards = AscDFH.historyitem_type_IterateData | 3;
	AscDFH.historyitem_IterateDataType      = AscDFH.historyitem_type_IterateData | 4;

	AscDFH.historyitem_TmVal = AscDFH.historyitem_type_Tm | 1;

	AscDFH.historyitem_TavVal  = AscDFH.historyitem_type_Tav | 1;
	AscDFH.historyitem_TavFmla = AscDFH.historyitem_type_Tav | 2;
	AscDFH.historyitem_TavTm   = AscDFH.historyitem_type_Tav | 3;

	AscDFH.historyitem_AnimVariantBoolVal = AscDFH.historyitem_type_AnimVariant | 1;
	AscDFH.historyitem_AnimVariantClrVal  = AscDFH.historyitem_type_AnimVariant | 2;
	AscDFH.historyitem_AnimVariantFltVal  = AscDFH.historyitem_type_AnimVariant | 3;
	AscDFH.historyitem_AnimVariantIntVal  = AscDFH.historyitem_type_AnimVariant | 4;
	AscDFH.historyitem_AnimVariantStrVal  = AscDFH.historyitem_type_AnimVariant | 5;

	AscDFH.historyitem_AnimClrByRGB  = AscDFH.historyitem_type_AnimClr | 1;
	AscDFH.historyitem_AnimClrByHSL  = AscDFH.historyitem_type_AnimClr | 2;
	AscDFH.historyitem_AnimClrCBhvr  = AscDFH.historyitem_type_AnimClr | 3;
	AscDFH.historyitem_AnimClrFrom   = AscDFH.historyitem_type_AnimClr | 4;
	AscDFH.historyitem_AnimClrTo     = AscDFH.historyitem_type_AnimClr | 5;
	AscDFH.historyitem_AnimClrClrSpc = AscDFH.historyitem_type_AnimClr | 6;
	AscDFH.historyitem_AnimClrDir    = AscDFH.historyitem_type_AnimClr | 7;

	AscDFH.historyitem_AnimEffectCBhvr      = AscDFH.historyitem_type_AnimEffect | 1;
	AscDFH.historyitem_AnimEffectProgress   = AscDFH.historyitem_type_AnimEffect | 2;
	AscDFH.historyitem_AnimEffectFilter     = AscDFH.historyitem_type_AnimEffect | 3;
	AscDFH.historyitem_AnimEffectPrLst      = AscDFH.historyitem_type_AnimEffect | 4;
	AscDFH.historyitem_AnimEffectTransition = AscDFH.historyitem_type_AnimEffect | 5;

	AscDFH.historyitem_AnimMotionBy           = AscDFH.historyitem_type_AnimMotion | 1;
	AscDFH.historyitem_AnimMotionCBhvr        = AscDFH.historyitem_type_AnimMotion | 2;
	AscDFH.historyitem_AnimMotionFrom         = AscDFH.historyitem_type_AnimMotion | 3;
	AscDFH.historyitem_AnimMotionRCtr         = AscDFH.historyitem_type_AnimMotion | 4;
	AscDFH.historyitem_AnimMotionTo           = AscDFH.historyitem_type_AnimMotion | 5;
	AscDFH.historyitem_AnimMotionOrigin       = AscDFH.historyitem_type_AnimMotion | 6;
	AscDFH.historyitem_AnimMotionPath         = AscDFH.historyitem_type_AnimMotion | 7;
	AscDFH.historyitem_AnimMotionPathEditMode = AscDFH.historyitem_type_AnimMotion | 8;
	AscDFH.historyitem_AnimMotionPtsTypes     = AscDFH.historyitem_type_AnimMotion | 9;
	AscDFH.historyitem_AnimMotionRAng         = AscDFH.historyitem_type_AnimMotion | 10;

	AscDFH.historyitem_AnimRotCBhvr = AscDFH.historyitem_type_AnimRot | 1;
	AscDFH.historyitem_AnimRotBy    = AscDFH.historyitem_type_AnimRot | 2;
	AscDFH.historyitem_AnimRotFrom  = AscDFH.historyitem_type_AnimRot | 3;
	AscDFH.historyitem_AnimRotTo    = AscDFH.historyitem_type_AnimRot | 4;

	AscDFH.historyitem_AnimScaleCBhvr        = AscDFH.historyitem_type_AnimScale | 1;
	AscDFH.historyitem_AnimScaleBy           = AscDFH.historyitem_type_AnimScale | 2;
	AscDFH.historyitem_AnimScaleFrom         = AscDFH.historyitem_type_AnimScale | 3;
	AscDFH.historyitem_AnimScaleTo           = AscDFH.historyitem_type_AnimScale | 4;
	AscDFH.historyitem_AnimScaleZoomContents = AscDFH.historyitem_type_AnimScale | 5;

	AscDFH.historyitem_AudioCMediaNode  = AscDFH.historyitem_type_Audio | 1;
	AscDFH.historyitem_AudioIsNarration = AscDFH.historyitem_type_Audio | 2;

	AscDFH.historyitem_CMediaNodeCTn             = AscDFH.historyitem_type_CMediaNode | 1;
	AscDFH.historyitem_CMediaNodeTgtEl           = AscDFH.historyitem_type_CMediaNode | 2;
	AscDFH.historyitem_CMediaNodeMute            = AscDFH.historyitem_type_CMediaNode | 3;
	AscDFH.historyitem_CMediaNodeNumSld          = AscDFH.historyitem_type_CMediaNode | 4;
	AscDFH.historyitem_CMediaNodeShowWhenStopped = AscDFH.historyitem_type_CMediaNode | 5;
	AscDFH.historyitem_CMediaNodeVol             = AscDFH.historyitem_type_CMediaNode | 6;

	AscDFH.historyitem_CmdCBhvr = AscDFH.historyitem_type_Cmd | 1;
	AscDFH.historyitem_CmdCmd   = AscDFH.historyitem_type_Cmd | 2;
	AscDFH.historyitem_CmdType  = AscDFH.historyitem_type_Cmd | 3;


	AscDFH.historyitem_OleChartElLvl  = AscDFH.historyitem_type_OleChartEl | 1;
	AscDFH.historyitem_OleChartElType = AscDFH.historyitem_type_OleChartEl | 2;


	AscDFH.historyitem_TimeNodeContainerCTn = AscDFH.historyitem_type_TimeNodeContainer | 1;

	AscDFH.historyitem_SeqNextCondLst = AscDFH.historyitem_type_Seq | 1;
	AscDFH.historyitem_SeqPrevCondLst = AscDFH.historyitem_type_Seq | 2;
	AscDFH.historyitem_SeqConcurrent  = AscDFH.historyitem_type_Seq | 3;
	AscDFH.historyitem_SeqNextAc      = AscDFH.historyitem_type_Seq | 4;
	AscDFH.historyitem_SeqPrevAc      = AscDFH.historyitem_type_Seq | 5;

	AscDFH.historyitem_SetCBhvr = AscDFH.historyitem_type_Set | 1;
	AscDFH.historyitem_SetTo    = AscDFH.historyitem_type_Set | 2;

	AscDFH.historyitem_VideoCMediaNode = AscDFH.historyitem_type_Video | 1;
	AscDFH.historyitem_VideoFullScrn   = AscDFH.historyitem_type_Video | 2;

	AscDFH.historyitem_OleChartElLvl  = AscDFH.historyitem_type_OleChartEl | 1;
	AscDFH.historyitem_OleChartElType = AscDFH.historyitem_type_OleChartEl | 2;

	AscDFH.historyitem_TlPointX = AscDFH.historyitem_type_TlPoint | 1;
	AscDFH.historyitem_TlPointY = AscDFH.historyitem_type_TlPoint | 2;

	AscDFH.historyitem_SndAcEndSnd = AscDFH.historyitem_type_SndAc | 1;
	AscDFH.historyitem_SndAcStSnd  = AscDFH.historyitem_type_SndAc | 2;

	AscDFH.historyitem_StSndSnd  = AscDFH.historyitem_type_StSnd | 1;
	AscDFH.historyitem_StSndLoop = AscDFH.historyitem_type_StSnd | 2;

	AscDFH.historyitem_TxElCharRg = AscDFH.historyitem_type_TxEl | 1;
	AscDFH.historyitem_TxElPRg    = AscDFH.historyitem_type_TxEl | 2;

	AscDFH.historyitem_WheelSpokes = AscDFH.historyitem_type_Wheel | 1;

	AscDFH.historyitem_AttrNameText = AscDFH.historyitem_type_AttrName | 1;

	AscDFH.historyitem_SldSzCX   = AscDFH.historyitem_type_SldSz | 1;
	AscDFH.historyitem_SldSzCY   = AscDFH.historyitem_type_SldSz | 2;
	AscDFH.historyitem_SldSzType = AscDFH.historyitem_type_SldSz | 3;

	AscDFH.historyitem_PrSetCoherent3DOff        = AscDFH.historyitem_type_PrSet | 1;
	AscDFH.historyitem_PrSetCsCatId              = AscDFH.historyitem_type_PrSet | 2;
	AscDFH.historyitem_PrSetCsTypeId             = AscDFH.historyitem_type_PrSet | 3;
	AscDFH.historyitem_PrSetCustAng              = AscDFH.historyitem_type_PrSet | 4;
	AscDFH.historyitem_PrSetCustFlipHor          = AscDFH.historyitem_type_PrSet | 5;
	AscDFH.historyitem_PrSetCustFlipVert         = AscDFH.historyitem_type_PrSet | 6;
	AscDFH.historyitem_PrSetCustLinFactNeighborX = AscDFH.historyitem_type_PrSet | 7;
	AscDFH.historyitem_PrSetCustLinFactNeighborY = AscDFH.historyitem_type_PrSet | 8;
	AscDFH.historyitem_PrSetCustLinFactX         = AscDFH.historyitem_type_PrSet | 9;
	AscDFH.historyitem_PrSetCustLinFactY         = AscDFH.historyitem_type_PrSet | 10;
	AscDFH.historyitem_PrSetCustRadScaleInc      = AscDFH.historyitem_type_PrSet | 11;
	AscDFH.historyitem_PrSetCustRadScaleRad      = AscDFH.historyitem_type_PrSet | 12;
	AscDFH.historyitem_PrSetCustScaleX           = AscDFH.historyitem_type_PrSet | 13;
	AscDFH.historyitem_PrSetCustScaleY           = AscDFH.historyitem_type_PrSet | 14;
	AscDFH.historyitem_PrSetCustSzX              = AscDFH.historyitem_type_PrSet | 15;
	AscDFH.historyitem_PrSetCustSzY              = AscDFH.historyitem_type_PrSet | 16;
	AscDFH.historyitem_PrSetCustT                = AscDFH.historyitem_type_PrSet | 17;
	AscDFH.historyitem_PrSetLoCatId              = AscDFH.historyitem_type_PrSet | 18;
	AscDFH.historyitem_PrSetLoTypeId             = AscDFH.historyitem_type_PrSet | 19;
	AscDFH.historyitem_PrSetPhldr                = AscDFH.historyitem_type_PrSet | 20;
	AscDFH.historyitem_PrSetPhldrT               = AscDFH.historyitem_type_PrSet | 21;
	AscDFH.historyitem_PrSetPresAssocID          = AscDFH.historyitem_type_PrSet | 22;
	AscDFH.historyitem_PrSetPresName             = AscDFH.historyitem_type_PrSet | 23;
	AscDFH.historyitem_PrSetPresStyleCnt         = AscDFH.historyitem_type_PrSet | 24;
	AscDFH.historyitem_PrSetPresStyleIdx         = AscDFH.historyitem_type_PrSet | 25;
	AscDFH.historyitem_PrSetPresStyleLbl         = AscDFH.historyitem_type_PrSet | 26;
	AscDFH.historyitem_PrSetQsCatId              = AscDFH.historyitem_type_PrSet | 27;
	AscDFH.historyitem_PrSetQsTypeId             = AscDFH.historyitem_type_PrSet | 28;
	AscDFH.historyitem_PrSetStyle                = AscDFH.historyitem_type_PrSet | 29;
	AscDFH.historyitem_PrSetPresLayoutVars       = AscDFH.historyitem_type_PrSet | 30;

	AscDFH.historyitem_PointCxnId   = AscDFH.historyitem_type_Point | 1;
	AscDFH.historyitem_PointModelId = AscDFH.historyitem_type_Point | 2;
	AscDFH.historyitem_PointType    = AscDFH.historyitem_type_Point | 3;
	AscDFH.historyitem_PointPrSet   = AscDFH.historyitem_type_Point | 5;
	AscDFH.historyitem_PointSpPr    = AscDFH.historyitem_type_Point | 6;
	AscDFH.historyitem_PointT       = AscDFH.historyitem_type_Point | 7;

	AscDFH.historyitem_CCommonDataListAdd    = AscDFH.historyitem_type_CCommonDataList | 1;
	AscDFH.historyitem_CCommonDataListRemove = AscDFH.historyitem_type_CCommonDataList | 2;

	AscDFH.historyitem_BgFormatFill  = AscDFH.historyitem_type_BgFormat | 1;
	AscDFH.historyitem_BgFormatEffect = AscDFH.historyitem_type_BgFormat | 2;


	AscDFH.historyitem_WholeEffect    = AscDFH.historyitem_type_Whole | 1;
	AscDFH.historyitem_WholeLn        = AscDFH.historyitem_type_Whole | 2;

	AscDFH.historyitem_CxnDestId     = AscDFH.historyitem_type_Cxn | 1;
	AscDFH.historyitem_CxnDestOrd    = AscDFH.historyitem_type_Cxn | 2;
	AscDFH.historyitem_CxnModelId    = AscDFH.historyitem_type_Cxn | 3;
	AscDFH.historyitem_CxnParTransId = AscDFH.historyitem_type_Cxn | 4;
	AscDFH.historyitem_CxnPresId     = AscDFH.historyitem_type_Cxn | 5;
	AscDFH.historyitem_CxnSibTransId = AscDFH.historyitem_type_Cxn | 6;
	AscDFH.historyitem_CxnSrcId      = AscDFH.historyitem_type_Cxn | 7;
	AscDFH.historyitem_CxnSrcOrd     = AscDFH.historyitem_type_Cxn | 8;
	AscDFH.historyitem_CxnType       = AscDFH.historyitem_type_Cxn | 9;

	AscDFH.historyitem_DataModelBg     = AscDFH.historyitem_type_DataModel | 1;
	AscDFH.historyitem_DataModelCxnLst = AscDFH.historyitem_type_DataModel | 2;
	AscDFH.historyitem_DataModelPtLst  = AscDFH.historyitem_type_DataModel | 4;
	AscDFH.historyitem_DataModelWhole  = AscDFH.historyitem_type_DataModel | 5;

	AscDFH.historyitem_DiagramDataDataModel = AscDFH.historyitem_type_DiagramData | 1;

	AscDFH.historyitem_Scene3dBackdrop = AscDFH.historyitem_type_Scene3d | 1;
	AscDFH.historyitem_Scene3dCamera = AscDFH.historyitem_type_Scene3d | 2;
	AscDFH.historyitem_Scene3dLightRig = AscDFH.historyitem_type_Scene3d | 4;

	AscDFH.historyitem_Scene3dBackdrop = AscDFH.historyitem_type_Scene3d | 1;
	AscDFH.historyitem_Scene3dCamera = AscDFH.historyitem_type_Scene3d | 2;
	AscDFH.historyitem_Scene3dLightRig = AscDFH.historyitem_type_Scene3d | 4;

	AscDFH.historyitem_BackdropAnchor = AscDFH.historyitem_type_Backdrop | 1;
	AscDFH.historyitem_BackdropNorm = AscDFH.historyitem_type_Backdrop | 3;
	AscDFH.historyitem_BackdropUp = AscDFH.historyitem_type_Backdrop | 4;

	AscDFH.historyitem_BackdropNormDx = AscDFH.historyitem_type_BackdropNorm | 1;
	AscDFH.historyitem_BackdropNormDy = AscDFH.historyitem_type_BackdropNorm | 2;
	AscDFH.historyitem_BackdropNormDz = AscDFH.historyitem_type_BackdropNorm | 3;

	AscDFH.historyitem_BackdropUpDx = AscDFH.historyitem_type_BackdropUp | 1;
	AscDFH.historyitem_BackdropUpDy = AscDFH.historyitem_type_BackdropUp | 2;
	AscDFH.historyitem_BackdropUpDz = AscDFH.historyitem_type_BackdropUp | 3;

	AscDFH.historyitem_CameraFov = AscDFH.historyitem_type_Camera | 1;
	AscDFH.historyitem_CameraPrst = AscDFH.historyitem_type_Camera | 2;
	AscDFH.historyitem_CameraZoom = AscDFH.historyitem_type_Camera | 3;
	AscDFH.historyitem_CameraRot = AscDFH.historyitem_type_Camera | 4;

	AscDFH.historyitem_RotLat = AscDFH.historyitem_type_Rot | 1;
	AscDFH.historyitem_RotLon = AscDFH.historyitem_type_Rot | 2;
	AscDFH.historyitem_RotRev = AscDFH.historyitem_type_Rot | 3;

	AscDFH.historyitem_LightRigDir = AscDFH.historyitem_type_LightRig | 1;
	AscDFH.historyitem_LightRigRig = AscDFH.historyitem_type_LightRig | 2;
	AscDFH.historyitem_LightRigRot = AscDFH.historyitem_type_LightRig | 3;

	AscDFH.historyitem_Sp3dContourW = AscDFH.historyitem_type_Sp3d | 1;
	AscDFH.historyitem_Sp3dExtrusionH = AscDFH.historyitem_type_Sp3d | 2;
	AscDFH.historyitem_Sp3dPrstMaterial = AscDFH.historyitem_type_Sp3d | 3;
	AscDFH.historyitem_Sp3dZ = AscDFH.historyitem_type_Sp3d | 4;
	AscDFH.historyitem_Sp3dBevelB = AscDFH.historyitem_type_Sp3d | 5;
	AscDFH.historyitem_Sp3dBevelT = AscDFH.historyitem_type_Sp3d | 6;
	AscDFH.historyitem_Sp3dContourClr = AscDFH.historyitem_type_Sp3d | 7;
	AscDFH.historyitem_Sp3dExtrusionClr = AscDFH.historyitem_type_Sp3d | 9;

	AscDFH.historyitem_BevelH = AscDFH.historyitem_type_Bevel | 1;
	AscDFH.historyitem_BevelPrst = AscDFH.historyitem_type_Bevel | 2;
	AscDFH.historyitem_BevelW = AscDFH.historyitem_type_Bevel | 3;

	AscDFH.historyitem_TxPrFlatTx = AscDFH.historyitem_type_TxPr | 1;
	AscDFH.historyitem_TxPrSp3d = AscDFH.historyitem_type_TxPr | 2;

	AscDFH.historyitem_FlatTxZ = AscDFH.historyitem_type_FlatTx | 1;

	AscDFH.historyitem_BackdropAnchorX = AscDFH.historyitem_type_BackdropAnchor | 1;
	AscDFH.historyitem_BackdropAnchorY = AscDFH.historyitem_type_BackdropAnchor | 2;
	AscDFH.historyitem_BackdropAnchorZ = AscDFH.historyitem_type_BackdropAnchor | 3;

	AscDFH.historyitem_CoordinateCoordinateUnqualified = AscDFH.historyitem_type_Coordinate | 1;
	AscDFH.historyitem_CoordinateUniversalMeasure      = AscDFH.historyitem_type_Coordinate | 2;

	AscDFH.historyitem_ChartStyleAxisTitle          = AscDFH.historyitem_type_ChartStyle | 1;
	AscDFH.historyitem_ChartStyleCategoryAxis       = AscDFH.historyitem_type_ChartStyle | 2;
	AscDFH.historyitem_ChartStyleChartArea          = AscDFH.historyitem_type_ChartStyle | 3;
	AscDFH.historyitem_ChartStyleDataLabel          = AscDFH.historyitem_type_ChartStyle | 4;
	AscDFH.historyitem_ChartStyleDataLabelCallout   = AscDFH.historyitem_type_ChartStyle | 5;
	AscDFH.historyitem_ChartStyleDataPoint          = AscDFH.historyitem_type_ChartStyle | 5;
	AscDFH.historyitem_ChartStyleDataPoint3D        = AscDFH.historyitem_type_ChartStyle | 6;
	AscDFH.historyitem_ChartStyleDataPointLine      = AscDFH.historyitem_type_ChartStyle | 7;
	AscDFH.historyitem_ChartStyleDataPointMarker    = AscDFH.historyitem_type_ChartStyle | 8;
	AscDFH.historyitem_ChartStyleDataPointWireframe = AscDFH.historyitem_type_ChartStyle | 9;
	AscDFH.historyitem_ChartStyleDataTable          = AscDFH.historyitem_type_ChartStyle | 10;
	AscDFH.historyitem_ChartStyleDownBar            = AscDFH.historyitem_type_ChartStyle | 11;
	AscDFH.historyitem_ChartStyleDropLine           = AscDFH.historyitem_type_ChartStyle | 12;
	AscDFH.historyitem_ChartStyleErrorBar           = AscDFH.historyitem_type_ChartStyle | 13;
	AscDFH.historyitem_ChartStyleFloor              = AscDFH.historyitem_type_ChartStyle | 14;
	AscDFH.historyitem_ChartStyleGridlineMajor      = AscDFH.historyitem_type_ChartStyle | 15;
	AscDFH.historyitem_ChartStyleGridlineMinor      = AscDFH.historyitem_type_ChartStyle | 16;
	AscDFH.historyitem_ChartStyleHiLoLine           = AscDFH.historyitem_type_ChartStyle | 17;
	AscDFH.historyitem_ChartStyleLeaderLine         = AscDFH.historyitem_type_ChartStyle | 18;
	AscDFH.historyitem_ChartStyleLegend             = AscDFH.historyitem_type_ChartStyle | 19;
	AscDFH.historyitem_ChartStylePlotArea           = AscDFH.historyitem_type_ChartStyle | 20;
	AscDFH.historyitem_ChartStylePlotArea3D         = AscDFH.historyitem_type_ChartStyle | 21;
	AscDFH.historyitem_ChartStyleSeriesAxis         = AscDFH.historyitem_type_ChartStyle | 22;
	AscDFH.historyitem_ChartStyleSeriesLine         = AscDFH.historyitem_type_ChartStyle | 23;
	AscDFH.historyitem_ChartStyleTitle              = AscDFH.historyitem_type_ChartStyle | 24;
	AscDFH.historyitem_ChartStyleTrendline          = AscDFH.historyitem_type_ChartStyle | 25;
	AscDFH.historyitem_ChartStyleTrendlineLabel     = AscDFH.historyitem_type_ChartStyle | 26;
	AscDFH.historyitem_ChartStyleUpBar              = AscDFH.historyitem_type_ChartStyle | 27;
	AscDFH.historyitem_ChartStyleValueAxis          = AscDFH.historyitem_type_ChartStyle | 28;
	AscDFH.historyitem_ChartStyleWall               = AscDFH.historyitem_type_ChartStyle | 29;
	AscDFH.historyitem_ChartStyleMarkerLayout       = AscDFH.historyitem_type_ChartStyle | 30;
	AscDFH.historyitem_ChartStyleMarkerId           = AscDFH.historyitem_type_ChartStyle | 31;

	AscDFH.historyitem_ChartStyleEntryType           = AscDFH.historyitem_type_ChartStyleEntry | 1;
	AscDFH.historyitem_ChartStyleEntryLineWidthScale = AscDFH.historyitem_type_ChartStyleEntry | 2;
	AscDFH.historyitem_ChartStyleEntryLnRef          = AscDFH.historyitem_type_ChartStyleEntry | 3;
	AscDFH.historyitem_ChartStyleEntryFillRef        = AscDFH.historyitem_type_ChartStyleEntry | 4;
	AscDFH.historyitem_ChartStyleEntryEffectRef      = AscDFH.historyitem_type_ChartStyleEntry | 5;
	AscDFH.historyitem_ChartStyleEntryFontRef        = AscDFH.historyitem_type_ChartStyleEntry | 6;
	AscDFH.historyitem_ChartStyleEntryDefRPr         = AscDFH.historyitem_type_ChartStyleEntry | 7;
	AscDFH.historyitem_ChartStyleEntryBodyPr         = AscDFH.historyitem_type_ChartStyleEntry | 8;
	AscDFH.historyitem_ChartStyleEntrySpPr           = AscDFH.historyitem_type_ChartStyleEntry | 9;

	AscDFH.historyitem_MarkerLayoutSymbol = AscDFH.historyitem_type_MarkerLayout | 1;
	AscDFH.historyitem_MarkerLayoutSize   = AscDFH.historyitem_type_MarkerLayout | 2;

	AscDFH.historyitem_TimelineSlicerViewName   = AscDFH.historyitem_type_TimelineSlicerView | 1;

	AscDFH.historyitem_SmartArtColorsDef = AscDFH.historyitem_type_SmartArt | 1;
	AscDFH.historyitem_SmartArtDrawing   = AscDFH.historyitem_type_SmartArt | 2;
	AscDFH.historyitem_SmartArtLayoutDef = AscDFH.historyitem_type_SmartArt | 3;
	AscDFH.historyitem_SmartArtDataModel = AscDFH.historyitem_type_SmartArt | 4;
	AscDFH.historyitem_SmartArtStyleDef  = AscDFH.historyitem_type_SmartArt | 5;
	AscDFH.historyitem_SmartArtParent    = AscDFH.historyitem_type_SmartArt | 6;
	AscDFH.historyitem_SmartArtType      = AscDFH.historyitem_type_SmartArt | 7;

	AscDFH.historyitem_ViewPrGridSpacing        = AscDFH.historyitem_type_ViewPr | 1;
	AscDFH.historyitem_ViewPrSlideViewerPr      = AscDFH.historyitem_type_ViewPr | 2;
	AscDFH.historyitem_ViewPrLastView           = AscDFH.historyitem_type_ViewPr | 3;
	AscDFH.historyitem_ViewPrShowComments       = AscDFH.historyitem_type_ViewPr | 4;

	AscDFH.historyitem_CommonViewPrCSldViewPr   = AscDFH.historyitem_type_CommonViewPr | 1;

	AscDFH.historyitem_CSldViewPrCViewPr        = AscDFH.historyitem_type_CSldViewPr | 1;
	AscDFH.historyitem_CSldViewPrGuideLst       = AscDFH.historyitem_type_CSldViewPr | 2;
	AscDFH.historyitem_CSldViewPrShowGuides     = AscDFH.historyitem_type_CSldViewPr | 3;
	AscDFH.historyitem_CSldViewPrSnapToGrid     = AscDFH.historyitem_type_CSldViewPr | 4;
	AscDFH.historyitem_CSldViewPrSnapToObjects  = AscDFH.historyitem_type_CSldViewPr | 5;

	AscDFH.historyitem_CViewPrOrigin            = AscDFH.historyitem_type_CViewPr | 1;
	AscDFH.historyitem_CViewPrScale             = AscDFH.historyitem_type_CViewPr | 2;

	AscDFH.historyitem_ViewPrScaleSx             = AscDFH.historyitem_type_ViewPrScale | 1;
	AscDFH.historyitem_ViewPrScaleSy             = AscDFH.historyitem_type_ViewPrScale | 2;

	AscDFH.historyitem_ViewPrGuideOrient         = AscDFH.historyitem_type_ViewPrGuide | 1;
	AscDFH.historyitem_ViewPrGuidePos            = AscDFH.historyitem_type_ViewPrGuide | 2;


	AscDFH.historyitem_Address_SetAddress1       = AscDFH.historyitem_type_Address | 1;
	AscDFH.historyitem_Address_SetCountryRegion  = AscDFH.historyitem_type_Address | 2;
	AscDFH.historyitem_Address_SetAdminDistrict1 = AscDFH.historyitem_type_Address | 3;
	AscDFH.historyitem_Address_SetAdminDistrict2 = AscDFH.historyitem_type_Address | 4;
	AscDFH.historyitem_Address_SetPostalCode     = AscDFH.historyitem_type_Address | 5;
	AscDFH.historyitem_Address_SetLocality       = AscDFH.historyitem_type_Address | 6;
	AscDFH.historyitem_Address_SetISOCountryCode = AscDFH.historyitem_type_Address | 7;

	AscDFH.historyitem_Axis_SetUnits             = AscDFH.historyitem_type_Axis | 1;
	AscDFH.historyitem_Axis_SetTickLabels        = AscDFH.historyitem_type_Axis | 2;
	AscDFH.historyitem_Axis_SetHidden            = AscDFH.historyitem_type_Axis | 3;

	AscDFH.historyitem_AxisUnits_SetUnitsLabel = AscDFH.historyitem_type_AxisUnits | 1;
	AscDFH.historyitem_AxisUnits_SetUnit       = AscDFH.historyitem_type_AxisUnits | 2;

	AscDFH.historyitem_AxisUnitsLabel_SetTx   = AscDFH.historyitem_type_AxisUnitsLabel | 1;
	AscDFH.historyitem_AxisUnitsLabel_SetSpPr = AscDFH.historyitem_type_AxisUnitsLabel | 2;
	AscDFH.historyitem_AxisUnitsLabel_SetTxPr = AscDFH.historyitem_type_AxisUnitsLabel | 3;

	AscDFH.historyitem_Binning_SetBinSize        = AscDFH.historyitem_type_Binning | 1;
	AscDFH.historyitem_Binning_SetBinCount       = AscDFH.historyitem_type_Binning | 2;
	AscDFH.historyitem_Binning_SetIntervalClosed = AscDFH.historyitem_type_Binning | 3;
	AscDFH.historyitem_Binning_SetUnderflow      = AscDFH.historyitem_type_Binning | 4;
	AscDFH.historyitem_Binning_SetOverflow       = AscDFH.historyitem_type_Binning | 5;

	AscDFH.historyitem_CategoryAxisScaling_SetGapWidth = AscDFH.historyitem_type_CategoryAxisScaling | 1;

	AscDFH.historyitem_ChartData_SetExternalData = AscDFH.historyitem_type_ChartData | 1;
	AscDFH.historyitem_ChartData_AddData         = AscDFH.historyitem_type_ChartData | 2;
	AscDFH.historyitem_ChartData_RemoveData      = AscDFH.historyitem_type_ChartData | 3;

	AscDFH.historyitem_Clear_SetGeoLocationQueryResults          = AscDFH.historyitem_type_Clear | 1;
	AscDFH.historyitem_Clear_SetGeoDataEntityQueryResults        = AscDFH.historyitem_type_Clear | 2;
	AscDFH.historyitem_Clear_SetGeoDataPointToEntityQueryResults = AscDFH.historyitem_type_Clear | 3;
	AscDFH.historyitem_Clear_SetGeoChildEntitiesQueryResults     = AscDFH.historyitem_type_Clear | 4;

	AscDFH.historyitem_Copyrights_SetCopyright = AscDFH.historyitem_type_Copyrights | 1;

	AscDFH.historyitem_Data_AddDimension    = AscDFH.historyitem_type_Data | 1;
	AscDFH.historyitem_Data_RemoveDimension = AscDFH.historyitem_type_Data | 2;
	AscDFH.historyitem_Data_SetId           = AscDFH.historyitem_type_Data | 3;

	AscDFH.historyitem_DataLabel_SetNumFmt     = AscDFH.historyitem_type_DataLabel | 1;
	AscDFH.historyitem_DataLabel_SetSpPr       = AscDFH.historyitem_type_DataLabel | 2;
	AscDFH.historyitem_DataLabel_SetTxPr       = AscDFH.historyitem_type_DataLabel | 3;
	AscDFH.historyitem_DataLabel_SetVisibility = AscDFH.historyitem_type_DataLabel | 4;
	AscDFH.historyitem_DataLabel_SetSeparator  = AscDFH.historyitem_type_DataLabel | 5;
	AscDFH.historyitem_DataLabel_SetIdx        = AscDFH.historyitem_type_DataLabel | 6;
	AscDFH.historyitem_DataLabel_SetPos        = AscDFH.historyitem_type_DataLabel | 7;

	AscDFH.historyitem_DataLabelHidden_SetIdx = AscDFH.historyitem_type_DataLabelHidden | 1;

	AscDFH.historyitem_DataLabels_SetNumFmt             = AscDFH.historyitem_type_DataLabels | 1;
	AscDFH.historyitem_DataLabels_SetSpPr               = AscDFH.historyitem_type_DataLabels | 2;
	AscDFH.historyitem_DataLabels_SetTxPr               = AscDFH.historyitem_type_DataLabels | 3;
	AscDFH.historyitem_DataLabels_SetVisibility         = AscDFH.historyitem_type_DataLabels | 4;
	AscDFH.historyitem_DataLabels_SetSeparator          = AscDFH.historyitem_type_DataLabels | 5;
	AscDFH.historyitem_DataLabels_SetDataLabel          = AscDFH.historyitem_type_DataLabels | 6;
	AscDFH.historyitem_DataLabels_AddDataLabel          = AscDFH.historyitem_type_DataLabels | 7;
	AscDFH.historyitem_DataLabels_RemoveDataLabel       = AscDFH.historyitem_type_DataLabels | 8;
	AscDFH.historyitem_DataLabels_AddDataLabelHidden    = AscDFH.historyitem_type_DataLabels | 9;
	AscDFH.historyitem_DataLabels_RemoveDataLabelHidden = AscDFH.historyitem_type_DataLabels | 10;
	AscDFH.historyitem_DataLabels_SetPos                = AscDFH.historyitem_type_DataLabels | 11;


	AscDFH.historyitem_DataLabelVisibilities_SetSeriesName   = AscDFH.historyitem_type_DataLabelVisibilities | 1;
	AscDFH.historyitem_DataLabelVisibilities_SetCategoryName = AscDFH.historyitem_type_DataLabelVisibilities | 2;
	AscDFH.historyitem_DataLabelVisibilities_SetValue        = AscDFH.historyitem_type_DataLabelVisibilities | 3;

	AscDFH.historyitem_DataPoint_SetSpPr = AscDFH.historyitem_type_DataPoint | 1;
	AscDFH.historyitem_DataPoint_SetIdx  = AscDFH.historyitem_type_DataPoint | 2;

	AscDFH.historyitem_FormatOverride_SetSpPr    = AscDFH.historyitem_type_FormatOverride | 1;
	AscDFH.historyitem_FormatOverride_SetIdx     = AscDFH.historyitem_type_FormatOverride | 2;
	AscDFH.historyitem_FormatOverrides_SetFmtOvr = AscDFH.historyitem_type_FormatOverride | 3;

	AscDFH.historyitem_Formula_SetDir     = AscDFH.historyitem_type_Formula | 1;
	AscDFH.historyitem_Formula_SetContent = AscDFH.historyitem_type_Formula | 2;

	AscDFH.historyitem_GeoCache_SetBinary   = AscDFH.historyitem_type_GeoCache | 1;
	AscDFH.historyitem_GeoCache_SetClear    = AscDFH.historyitem_type_GeoCache | 2;
	AscDFH.historyitem_GeoCache_SetProvider = AscDFH.historyitem_type_GeoCache | 3;

	AscDFH.historyitem_GeoChildEntities_SetGeoHierarchyEntity = AscDFH.historyitem_type_GeoChildEntities | 1;

	AscDFH.historyitem_GeoChildEntitiesQuery_SetGeoChildTypes = AscDFH.historyitem_type_GeoChildEntitiesQuery | 1;
	AscDFH.historyitem_GeoChildEntitiesQuery_SetEntityId      = AscDFH.historyitem_type_GeoChildEntitiesQuery | 2;

	AscDFH.historyitem_GeoChildEntitiesQueryResult_SetGeoChildEntitiesQuery = AscDFH.historyitem_type_GeoChildEntitiesQueryResult | 1;
	AscDFH.historyitem_GeoChildEntitiesQueryResult_SetGeoChildEntities      = AscDFH.historyitem_type_GeoChildEntitiesQueryResult | 2;

	AscDFH.historyitem_GeoChildEntitiesQueryResults_SetGeoChildEntitiesQueryResult = AscDFH.historyitem_type_GeoChildEntitiesQueryResults | 1;

	AscDFH.historyitem_GeoChildTypes_SetEntityType = AscDFH.historyitem_type_GeoChildTypes | 1;

	AscDFH.historyitem_GeoData_SetGeoPolygons = AscDFH.historyitem_type_GeoData | 1;
	AscDFH.historyitem_GeoData_SetCopyrights  = AscDFH.historyitem_type_GeoData | 2;
	AscDFH.historyitem_GeoData_SetEntityName  = AscDFH.historyitem_type_GeoData | 3;
	AscDFH.historyitem_GeoData_SetEntityId    = AscDFH.historyitem_type_GeoData | 4;
	AscDFH.historyitem_GeoData_SetEast        = AscDFH.historyitem_type_GeoData | 5;
	AscDFH.historyitem_GeoData_SetWest        = AscDFH.historyitem_type_GeoData | 6;
	AscDFH.historyitem_GeoData_SetNorth       = AscDFH.historyitem_type_GeoData | 7;
	AscDFH.historyitem_GeoData_SetSouth       = AscDFH.historyitem_type_GeoData | 8;

	AscDFH.historyitem_GeoDataEntityQuery_SetEntityType = AscDFH.historyitem_type_GeoDataEntityQuery | 1;
	AscDFH.historyitem_GeoDataEntityQuery_SetEntityId   = AscDFH.historyitem_type_GeoDataEntityQuery |2;

	AscDFH.historyitem_GeoDataEntityQueryResult_SetGeoDataEntityQuery = AscDFH.historyitem_type_GeoDataEntityQueryResult | 1;
	AscDFH.historyitem_GeoDataEntityQueryResult_SetGeoData            = AscDFH.historyitem_type_GeoDataEntityQueryResult | 2;

	AscDFH.historyitem_GeoDataEntityQueryResults_SetGeoDataEntityQueryResult = AscDFH.historyitem_type_GeoDataEntityQueryResults | 1;

	AscDFH.historyitem_GeoDataPointQuery_SetEntityType = AscDFH.historyitem_type_GeoDataPointQuery | 1;
	AscDFH.historyitem_GeoDataPointQuery_SetLatitude   = AscDFH.historyitem_type_GeoDataPointQuery | 2;
	AscDFH.historyitem_GeoDataPointQuery_SetLongitude  = AscDFH.historyitem_type_GeoDataPointQuery | 3;

	AscDFH.historyitem_GeoDataPointToEntityQuery_SetEntityType = AscDFH.historyitem_type_GeoDataPointToEntityQuery | 1;
	AscDFH.historyitem_GeoDataPointToEntityQuery_SetEntityId   = AscDFH.historyitem_type_GeoDataPointToEntityQuery | 2;

	AscDFH.historyitem_GeoDataPointToEntityQueryResult_SetGeoDataPointQuery         = AscDFH.historyitem_type_GeoDataPointToEntityQueryResult | 1;
	AscDFH.historyitem_GeoDataPointToEntityQueryResult_SetGeoDataPointToEntityQuery = AscDFH.historyitem_type_GeoDataPointToEntityQueryResult | 2;

    AscDFH.historyitem_GeoDataPointToEntityQueryResults_SetGeoDataPointToEntityQueryResult = AscDFH.historyitem_type_GeoDataPointToEntityQueryResults | 1;

	AscDFH.historyitem_Geography_SetGeoCache         = AscDFH.historyitem_type_Geography | 1;
	AscDFH.historyitem_Geography_SetProjectionType   = AscDFH.historyitem_type_Geography | 2;
	AscDFH.historyitem_Geography_SetViewedRegionType = AscDFH.historyitem_type_Geography | 3;
	AscDFH.historyitem_Geography_SetCultureLanguage  = AscDFH.historyitem_type_Geography | 4;
	AscDFH.historyitem_Geography_SetCultureRegion    = AscDFH.historyitem_type_Geography | 5;
	AscDFH.historyitem_Geography_SetAttribution      = AscDFH.historyitem_type_Geography | 6;

	AscDFH.historyitem_GeoHierarchyEntity_SetEntityName = AscDFH.historyitem_type_GeoHierarchyEntity | 1;
	AscDFH.historyitem_GeoHierarchyEntity_SetEntityId   = AscDFH.historyitem_type_GeoHierarchyEntity | 2;
	AscDFH.historyitem_GeoHierarchyEntity_SetEntityType = AscDFH.historyitem_type_GeoHierarchyEntity | 3;

	AscDFH.historyitem_GeoLocation_SetAddress    = AscDFH.historyitem_type_GeoLocation | 1;
	AscDFH.historyitem_GeoLocation_SetLatitude   = AscDFH.historyitem_type_GeoLocation | 2;
	AscDFH.historyitem_GeoLocation_SetLongitude  = AscDFH.historyitem_type_GeoLocation | 3;
	AscDFH.historyitem_GeoLocation_SetEntityName = AscDFH.historyitem_type_GeoLocation | 4;
	AscDFH.historyitem_GeoLocation_SetEntityType = AscDFH.historyitem_type_GeoLocation | 5;

	AscDFH.historyitem_GeoLocationQuery_SetCountryRegion  = AscDFH.historyitem_type_GeoLocationQuery | 1;
	AscDFH.historyitem_GeoLocationQuery_SetAdminDistrict1 = AscDFH.historyitem_type_GeoLocationQuery | 2;
	AscDFH.historyitem_GeoLocationQuery_SetAdminDistrict2 = AscDFH.historyitem_type_GeoLocationQuery | 3;
	AscDFH.historyitem_GeoLocationQuery_SetPostalCode     = AscDFH.historyitem_type_GeoLocationQuery | 4;
	AscDFH.historyitem_GeoLocationQuery_SetEntityType     = AscDFH.historyitem_type_GeoLocationQuery | 5;

	AscDFH.historyitem_GeoLocationQueryResult_SetGeoLocationQuery = AscDFH.historyitem_type_GeoLocationQueryResult | 1;
	AscDFH.historyitem_GeoLocationQueryResult_SetGeoLocations     = AscDFH.historyitem_type_GeoLocationQueryResult | 2;

	AscDFH.historyitem_GeoLocationQueryResults_SetGeoLocationQueryResult = AscDFH.historyitem_type_GeoLocationQueryResults | 1;

	AscDFH.historyitem_GeoLocations_SetGeoLocation = AscDFH.historyitem_type_GeoLocations | 1;

	AscDFH.historyitem_GeoPolygon_SetPolygonId = AscDFH.historyitem_type_GeoPolygon | 1;
	AscDFH.historyitem_GeoPolygon_SetNumPoints = AscDFH.historyitem_type_GeoPolygon | 2;
	AscDFH.historyitem_GeoPolygon_SetPcaRings  = AscDFH.historyitem_type_GeoPolygon | 3;

	AscDFH.historyitem_GeoPolygons_SetGeoPolygon = AscDFH.historyitem_type_GeoPolygons | 1;

	AscDFH.historyitem_Gridlines_SetSpPr = AscDFH.historyitem_type_Gridlines | 1;
	AscDFH.historyitem_Gridlines_SetName = AscDFH.historyitem_type_Gridlines | 2;

	AscDFH.historyitem_Dimension_SetF    = AscDFH.historyitem_type_Dimension | 1;
	AscDFH.historyitem_Dimension_SetNf   = AscDFH.historyitem_type_Dimension | 2;
	AscDFH.historyitem_Dimension_SetType = AscDFH.historyitem_type_Dimension | 3;

	AscDFH.historyitem_NumericDimension_AddLevelData    = AscDFH.historyitem_type_NumericDimension | 1;
	AscDFH.historyitem_NumericDimension_RemoveLevelData = AscDFH.historyitem_type_NumericDimension | 2;

	AscDFH.historyitem_PercentageColorPosition_SetVal = AscDFH.historyitem_type_PercentageColorPosition | 1;

	AscDFH.historyitem_PlotAreaRegion_SetPlotSurface = AscDFH.historyitem_type_PlotAreaRegion | 1;
	AscDFH.historyitem_PlotAreaRegion_AddSeries      = AscDFH.historyitem_type_PlotAreaRegion | 2;
	AscDFH.historyitem_PlotAreaRegion_RemoveSeries   = AscDFH.historyitem_type_PlotAreaRegion | 3;

	AscDFH.historyitem_PlotSurface_SetSpPr = AscDFH.historyitem_type_PlotSurface | 1;

	AscDFH.historyitem_Series_AddDataPt     = AscDFH.historyitem_type_Series | 1;
	AscDFH.historyitem_Series_RemoveDataPt  = AscDFH.historyitem_type_Series | 2;
	AscDFH.historyitem_Series_SetDataLabels = AscDFH.historyitem_type_Series | 3;
	AscDFH.historyitem_Series_SetDataId     = AscDFH.historyitem_type_Series | 4;
	AscDFH.historyitem_Series_SetLayoutPr   = AscDFH.historyitem_type_Series | 5;
	AscDFH.historyitem_Series_AddAxisId     = AscDFH.historyitem_type_Series | 6;
	AscDFH.historyitem_Series_RemoveAxisId  = AscDFH.historyitem_type_Series | 7;
	AscDFH.historyitem_Series_SetLayoutId   = AscDFH.historyitem_type_Series | 8;
	AscDFH.historyitem_Series_SetHidden     = AscDFH.historyitem_type_Series | 9;
	AscDFH.historyitem_Series_SetOwnerIdx   = AscDFH.historyitem_type_Series | 10;
	AscDFH.historyitem_Series_SetUniqueId   = AscDFH.historyitem_type_Series | 11;
	AscDFH.historyitem_Series_SetFormatIdx  = AscDFH.historyitem_type_Series | 12;

	AscDFH.historyitem_SeriesElementVisibilities_SetConnectorLines = AscDFH.historyitem_type_SeriesElementVisibilities | 1;
	AscDFH.historyitem_SeriesElementVisibilities_SetMeanLine       = AscDFH.historyitem_type_SeriesElementVisibilities | 2;
	AscDFH.historyitem_SeriesElementVisibilities_SetMeanMarker     = AscDFH.historyitem_type_SeriesElementVisibilities | 3;
	AscDFH.historyitem_SeriesElementVisibilities_SetNonoutliers    = AscDFH.historyitem_type_SeriesElementVisibilities | 4;
	AscDFH.historyitem_SeriesElementVisibilities_SetOutliers       = AscDFH.historyitem_type_SeriesElementVisibilities | 5;

	AscDFH.historyitem_SeriesLayoutProperties_SetParentLabelLayout = AscDFH.historyitem_type_SeriesLayoutProperties | 1;
	AscDFH.historyitem_SeriesLayoutProperties_SetRegionLabelLayout = AscDFH.historyitem_type_SeriesLayoutProperties | 2;
	AscDFH.historyitem_SeriesLayoutProperties_SetVisibility        = AscDFH.historyitem_type_SeriesLayoutProperties | 3;
	AscDFH.historyitem_SeriesLayoutProperties_SetAggregation       = AscDFH.historyitem_type_SeriesLayoutProperties | 4;
	AscDFH.historyitem_SeriesLayoutProperties_SetBinning           = AscDFH.historyitem_type_SeriesLayoutProperties | 5;
	AscDFH.historyitem_SeriesLayoutProperties_SetGeography         = AscDFH.historyitem_type_SeriesLayoutProperties | 6;
	AscDFH.historyitem_SeriesLayoutProperties_SetStatistics        = AscDFH.historyitem_type_SeriesLayoutProperties | 7;
	AscDFH.historyitem_SeriesLayoutProperties_SetSubtotals         = AscDFH.historyitem_type_SeriesLayoutProperties | 8;

	AscDFH.historyitem_Statistics_SetQuartileMethod    = AscDFH.historyitem_type_Statistics | 1;

	AscDFH.historyitem_StringDimension_AddLevelData    = AscDFH.historyitem_type_StringDimension | 1;
	AscDFH.historyitem_StringDimension_RemoveLevelData = AscDFH.historyitem_type_StringDimension | 2;

	AscDFH.historyitem_Subtotals_AddIdx    = AscDFH.historyitem_type_Subtotals | 1;
	AscDFH.historyitem_Subtotals_RemoveIdx = AscDFH.historyitem_type_Subtotals | 2;

	AscDFH.historyitem_TextData_SetF = AscDFH.historyitem_type_TextData | 1;
	AscDFH.historyitem_TextData_SetV = AscDFH.historyitem_type_TextData | 2;

	AscDFH.historyitem_TickMarks_SetType = AscDFH.historyitem_type_TickMarks | 1;
	AscDFH.historyitem_TickMarks_SetName = AscDFH.historyitem_type_TickMarks | 2;

	AscDFH.historyitem_ValueAxisScaling_SetMax         = AscDFH.historyitem_type_ValueAxisScaling | 1;
	AscDFH.historyitem_ValueAxisScaling_SetMin         = AscDFH.historyitem_type_ValueAxisScaling | 2;
	AscDFH.historyitem_ValueAxisScaling_SetMajorUnit   = AscDFH.historyitem_type_ValueAxisScaling | 3;
	AscDFH.historyitem_ValueAxisScaling_SetMinorUnit   = AscDFH.historyitem_type_ValueAxisScaling | 4;

	AscDFH.historyitem_ValueColorEndPosition_SetExtremeValue = AscDFH.historyitem_type_ValueColorEndPosition | 1;
	AscDFH.historyitem_ValueColorEndPosition_SetNumber       = AscDFH.historyitem_type_ValueColorEndPosition | 2;
	AscDFH.historyitem_ValueColorEndPosition_SetPercent      = AscDFH.historyitem_type_ValueColorEndPosition | 3;

	AscDFH.historyitem_ValueColorMiddlePosition_SetNumber  = AscDFH.historyitem_type_ValueColorMiddlePosition | 1;
	AscDFH.historyitem_ValueColorMiddlePosition_SetPercent = AscDFH.historyitem_type_ValueColorMiddlePosition | 2;

	AscDFH.historyitem_ValueColorPositions_SetMin   = AscDFH.historyitem_type_ValueColorPositions | 1;
	AscDFH.historyitem_ValueColorPositions_SetMid   = AscDFH.historyitem_type_ValueColorPositions | 2;
	AscDFH.historyitem_ValueColorPositions_SetMax   = AscDFH.historyitem_type_ValueColorPositions | 3;
	AscDFH.historyitem_ValueColorPositions_SetCount = AscDFH.historyitem_type_ValueColorPositions | 4;

	AscDFH.historyitem_ValueColors_SetMinColor = AscDFH.historyitem_type_ValueColors | 1;
	AscDFH.historyitem_ValueColors_SetMidColor = AscDFH.historyitem_type_ValueColors | 2;
	AscDFH.historyitem_ValueColors_SetMaxColor = AscDFH.historyitem_type_ValueColors | 3;
	

	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в PDF Forms 
	//------------------------------------------------------------------------------------------------------------------

	// common
	AscDFH.historyitem_Pdf_Form_Value			= AscDFH.historyitem_type_Pdf_Form | 1;
	AscDFH.historyitem_Pdf_Form_Add_Kid			= AscDFH.historyitem_type_Pdf_Form | 2;
	AscDFH.historyitem_Pdf_Form_Change_Display	= AscDFH.historyitem_type_Pdf_Form | 4;
	AscDFH.historyitem_Pdf_Form_Changed			= AscDFH.historyitem_type_Pdf_Form | 5;
	AscDFH.historyitem_Pdf_Form_Parent_Value	= AscDFH.historyitem_type_Pdf_Form | 6;
	AscDFH.historyitem_Pdf_Form_Format_Value	= AscDFH.historyitem_type_Pdf_Form | 7;
	AscDFH.historyitem_Pdf_Form_Border_Color	= AscDFH.historyitem_type_Pdf_Form | 8;
	AscDFH.historyitem_Pdf_Form_BG_Color		= AscDFH.historyitem_type_Pdf_Form | 9;
	AscDFH.historyitem_Pdf_Form_Border_Style	= AscDFH.historyitem_type_Pdf_Form | 10;
	AscDFH.historyitem_Pdf_Form_Required		= AscDFH.historyitem_type_Pdf_Form | 11;
	AscDFH.historyitem_Pdf_Form_Text_Color		= AscDFH.historyitem_type_Pdf_Form | 12;
	AscDFH.historyitem_Pdf_Form_Text_Font		= AscDFH.historyitem_type_Pdf_Form | 13;
	AscDFH.historyitem_Pdf_Form_Text_Size		= AscDFH.historyitem_type_Pdf_Form | 14;
	AscDFH.historyitem_Pdf_Form_Default_Value	= AscDFH.historyitem_type_Pdf_Form | 15;
	AscDFH.historyitem_Pdf_Form_Rect			= AscDFH.historyitem_type_Pdf_Form | 16;
	AscDFH.historyitem_Pdf_Form_Actions			= AscDFH.historyitem_type_Pdf_Form | 17;
	AscDFH.historyitem_Pdf_Form_Partial_Name	= AscDFH.historyitem_type_Pdf_Form | 18;
	AscDFH.historyitem_Pdf_Form_Meta			= AscDFH.historyitem_type_Pdf_Form | 19;
	AscDFH.historyitem_Pdf_Form_Read_Only		= AscDFH.historyitem_type_Pdf_Form | 20;
	AscDFH.historyitem_Pdf_Form_No_Export		= AscDFH.historyitem_type_Pdf_Form | 21;
	AscDFH.historyitem_Pdf_Form_Border_Width	= AscDFH.historyitem_type_Pdf_Form | 22;
	AscDFH.historyitem_Pdf_Form_Locked			= AscDFH.historyitem_type_Pdf_Form | 23;
	AscDFH.historyitem_Pdf_Form_Rotate			= AscDFH.historyitem_type_Pdf_Form | 24;
	
	// text
	AscDFH.historyitem_Pdf_Text_Form_Multiline			= AscDFH.historyitem_type_Pdf_Text_Field | 1;
	AscDFH.historyitem_Pdf_Text_Form_Align				= AscDFH.historyitem_type_Pdf_Text_Field | 2;
	AscDFH.historyitem_Pdf_Text_Form_Char_Limit			= AscDFH.historyitem_type_Pdf_Text_Field | 3;
	AscDFH.historyitem_Pdf_Text_Form_Comb				= AscDFH.historyitem_type_Pdf_Text_Field | 4;
	AscDFH.historyitem_Pdf_Text_Form_DoNot_Scroll		= AscDFH.historyitem_type_Pdf_Text_Field | 5;
	AscDFH.historyitem_Pdf_Text_Form_Password			= AscDFH.historyitem_type_Pdf_Text_Field | 6;
	AscDFH.historyitem_Pdf_Text_Form_File_Select		= AscDFH.historyitem_type_Pdf_Text_Field | 7;
	AscDFH.historyitem_Pdf_Text_Form_DoNot_Spell_Check	= AscDFH.historyitem_type_Pdf_Text_Field | 8;
	
	// combobox
	AscDFH.historyitem_Pdf_Combobox_Form_Editable = AscDFH.historyitem_type_Pdf_Combobox_Field | 1;

	// list
	AscDFH.historyitem_Pdf_List_Form_Cur_Idxs				= AscDFH.historyitem_type_Pdf_Listbox_Field | 1;
	AscDFH.historyitem_Pdf_List_Form_Parent_Cur_Idxs		= AscDFH.historyitem_type_Pdf_Listbox_Field | 2;
	AscDFH.historyitem_Pdf_List_Form_Top_Idx				= AscDFH.historyitem_type_Pdf_Listbox_Field | 3;
	AscDFH.historyitem_Pdf_List_Form_Option					= AscDFH.historyitem_type_Pdf_Listbox_Field | 4;
	AscDFH.historyitem_Pdf_List_Form_Content_Option			= AscDFH.historyitem_type_Pdf_Listbox_Field | 5;
	AscDFH.historyitem_Pdf_List_Form_Commit_On_Sel_Change	= AscDFH.historyitem_type_Pdf_Listbox_Field | 6;
	AscDFH.historyitem_Pdf_List_Form_Multiple_Selection		= AscDFH.historyitem_type_Pdf_Listbox_Field | 7;

	// button
	AscDFH.historyitem_Pdf_Pushbutton_Image				= AscDFH.historyitem_type_Pdf_Pushbutton | 1;
	AscDFH.historyitem_Pdf_Pushbutton_Layout			= AscDFH.historyitem_type_Pdf_Pushbutton | 2;
	AscDFH.historyitem_Pdf_Pushbutton_Icon_Pos			= AscDFH.historyitem_type_Pdf_Pushbutton | 3;
	AscDFH.historyitem_Pdf_Pushbutton_Highlight_Type	= AscDFH.historyitem_type_Pdf_Pushbutton | 4;
	AscDFH.historyitem_Pdf_Pushbutton_Scale_When_Type	= AscDFH.historyitem_type_Pdf_Pushbutton | 5;
	AscDFH.historyitem_Pdf_Pushbutton_Scale_How_Type	= AscDFH.historyitem_type_Pdf_Pushbutton | 6;
	AscDFH.historyitem_Pdf_Pushbutton_Fit_Bounds		= AscDFH.historyitem_type_Pdf_Pushbutton | 7;
	AscDFH.historyitem_Pdf_Pushbutton_Caption			= AscDFH.historyitem_type_Pdf_Pushbutton | 8;
	
	// checkbox/radio
	AscDFH.historyitem_Pdf_Checkbox_No_Toggle_To_Off	= AscDFH.historyitem_type_Pdf_Checkbox_Field | 1;
	AscDFH.historyitem_Pdf_Checkbox_Style				= AscDFH.historyitem_type_Pdf_Checkbox_Field | 2;
	AscDFH.historyitem_Pdf_Checkbox_Export_Value		= AscDFH.historyitem_type_Pdf_Checkbox_Field | 3;
	AscDFH.historyitem_Pdf_Checkbox_Options				= AscDFH.historyitem_type_Pdf_Checkbox_Field | 4;
	
	// radio
	AscDFH.historyitem_Pdf_Radiobutton_Is_Unison		= AscDFH.historyitem_type_Pdf_Radiobutton_Field | 1;
	

	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в PDF Annots 
	//------------------------------------------------------------------------------------------------------------------

	// Common
	AscDFH.historyitem_Pdf_Annot_Rect				= AscDFH.historyitem_type_Pdf_Annot | 1;
	AscDFH.historyitem_Pdf_Annot_Pos				= AscDFH.historyitem_type_Pdf_Annot | 2;
	AscDFH.historyitem_Pdf_Annot_Contents			= AscDFH.historyitem_type_Pdf_Annot | 3;
	AscDFH.historyitem_Pdf_Annot_Page				= AscDFH.historyitem_type_Pdf_Annot | 4;
	AscDFH.historyitem_Pdf_Annot_RD					= AscDFH.historyitem_type_Pdf_Annot | 5;
	AscDFH.historyitem_Pdf_Annot_Vertices			= AscDFH.historyitem_type_Pdf_Annot | 6;
	AscDFH.historyitem_Pdf_Annot_Creation_Date		= AscDFH.historyitem_type_Pdf_Annot | 7;
	AscDFH.historyitem_Pdf_Annot_Mod_Date			= AscDFH.historyitem_type_Pdf_Annot | 8;
	AscDFH.historyitem_Pdf_Annot_Author				= AscDFH.historyitem_type_Pdf_Annot | 9;
	AscDFH.historyitem_Pdf_Annot_Display			= AscDFH.historyitem_type_Pdf_Annot | 10;
	AscDFH.historyitem_Pdf_Annot_Name				= AscDFH.historyitem_type_Pdf_Annot | 11;
	AscDFH.historyitem_Pdf_Annot_File_Idx			= AscDFH.historyitem_type_Pdf_Annot | 12;
	AscDFH.historyitem_Pdf_Annot_Stroke				= AscDFH.historyitem_type_Pdf_Annot | 13;
	AscDFH.historyitem_Pdf_Annot_StrokeWidth		= AscDFH.historyitem_type_Pdf_Annot | 14;
	AscDFH.historyitem_Pdf_Annot_Fill				= AscDFH.historyitem_type_Pdf_Annot | 15;
	AscDFH.historyitem_Pdf_Annot_Opacity			= AscDFH.historyitem_type_Pdf_Annot | 16;
	AscDFH.historyitem_Pdf_Annot_Quads				= AscDFH.historyitem_type_Pdf_Annot | 17;
	AscDFH.historyitem_Pdf_Annot_Intent				= AscDFH.historyitem_type_Pdf_Annot | 18;
	AscDFH.historyitem_Pdf_Annot_Rotate				= AscDFH.historyitem_type_Pdf_Annot | 19;
	AscDFH.historyitem_Pdf_Annot_User_Id			= AscDFH.historyitem_type_Pdf_Annot | 20;
	AscDFH.historyitem_Pdf_Annot_Changed			= AscDFH.historyitem_type_Pdf_Annot | 21;
	AscDFH.historyitem_Pdf_Annot_Changed_View		= AscDFH.historyitem_type_Pdf_Annot | 22;

	// Comment
	AscDFH.historyitem_Pdf_Comment_Data			= AscDFH.historyitem_type_Pdf_Comment | 1;

	// Ink
	AscDFH.historyitem_Pdf_Ink_Points			= AscDFH.historyitem_type_Pdf_Annot_Ink | 1;
	AscDFH.historyitem_Pdf_Ink_FlipV			= AscDFH.historyitem_type_Pdf_Annot_Ink | 2;
	AscDFH.historyitem_Pdf_Ink_FlipH			= AscDFH.historyitem_type_Pdf_Annot_Ink | 3;

	// FreeText
	AscDFH.historyitem_type_Pdf_Annot_FreeText_CL			= AscDFH.historyitem_type_Pdf_Annot_FreeText | 1;
	AscDFH.historyitem_type_Pdf_Annot_FreeText_RC			= AscDFH.historyitem_type_Pdf_Annot_FreeText | 2;
	AscDFH.historyitem_type_Pdf_Annot_FreeText_Align		= AscDFH.historyitem_type_Pdf_Annot_FreeText | 3;
	AscDFH.historyitem_type_Pdf_Annot_FreeText_Rotate		= AscDFH.historyitem_type_Pdf_Annot_FreeText | 4;

	// annot line
	AscDFH.historyitem_Pdf_Line_Points			= AscDFH.historyitem_type_Pdf_Annot_Line | 1;
	
	// annot stamp
	AscDFH.historyitem_Pdf_Stamp_Type			 = AscDFH.historyitem_type_Pdf_Annot_Stamp | 1;
	AscDFH.historyitem_Pdf_Stamp_InRect			 = AscDFH.historyitem_type_Pdf_Annot_Stamp | 2;
	AscDFH.historyitem_Pdf_Stamp_Rect			 = AscDFH.historyitem_type_Pdf_Annot_Stamp | 3;
	AscDFH.historyitem_Pdf_Stamp_RenderStructure = AscDFH.historyitem_type_Pdf_Annot_Stamp | 4;

	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в PDF drawing prototype
	//------------------------------------------------------------------------------------------------------------------

	AscDFH.historyitem_type_Pdf_Drawing_Page		= AscDFH.historyitem_type_Pdf_Drawing | 1;

	//------------------------------------------------------------------------------------------------------------------
	// Типы изменений в классе CPDFDoc
	//------------------------------------------------------------------------------------------------------------------
	window['AscDFH'].historyitem_PDF_Document_AnnotsContent		= window['AscDFH'].historyitem_type_PDF_Document | 1;
	window['AscDFH'].historyitem_PDF_Document_DrawingsContent	= window['AscDFH'].historyitem_type_PDF_Document | 2;
	window['AscDFH'].historyitem_PDF_Document_FieldsContent		= window['AscDFH'].historyitem_type_PDF_Document | 3;
	window['AscDFH'].historyitem_PDF_Document_PagesContent		= window['AscDFH'].historyitem_type_PDF_Document | 4;
	window['AscDFH'].historyitem_PDF_Document_RotatePage		= window['AscDFH'].historyitem_type_PDF_Document | 5;
	window['AscDFH'].historyitem_PDF_Document_RecognizePage		= window['AscDFH'].historyitem_type_PDF_Document | 6;
	window['AscDFH'].historyitem_PDF_Document_SetDocument		= window['AscDFH'].historyitem_type_PDF_Document | 7;
	window['AscDFH'].historyitem_PDF_Document_PageLocks			= window['AscDFH'].historyitem_type_PDF_Document | 8;
	window['AscDFH'].historyitem_PDF_PropLocker_ObjectId		= window['AscDFH'].historyitem_type_PDF_Document | 9;
	window['AscDFH'].historyitem_Pdf_Document_Start_Merge_Pages	= window['AscDFH'].historyitem_type_PDF_Document | 10;
	window['AscDFH'].historyitem_Pdf_Document_Part_Merge_Pages	= window['AscDFH'].historyitem_type_PDF_Document | 11;
	window['AscDFH'].historyitem_Pdf_Document_End_Merge_Pages	= window['AscDFH'].historyitem_type_PDF_Document | 12;
	window['AscDFH'].historyitem_Pdf_Document_Calc_Order		= window['AscDFH'].historyitem_type_PDF_Document | 13;
	window['AscDFH'].historyitem_PDF_Document_Locks				= window['AscDFH'].historyitem_type_PDF_Document | 14;


	AscDFH.historyitem_CustomPropertiesAddProperty = AscDFH.historyitem_type_CustomProperties | 0;
	AscDFH.historyitem_CustomPropertiesRemoveProperty = AscDFH.historyitem_type_CustomProperties | 1;

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//
	// Типы действий, который может совершить пользователь
	//
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	window['AscDFH'].historydescription_Unknown                                     = 0x0000;
	window['AscDFH'].historydescription_Cut                                         = 0x0001;
	window['AscDFH'].historydescription_PasteButtonIE                               = 0x0002;
	window['AscDFH'].historydescription_PasteButtonNotIE                            = 0x0003;
	window['AscDFH'].historydescription_ChartDrawingObjects                         = 0x0004;
	window['AscDFH'].historydescription_CommonControllerCheckChartText              = 0x0005;
	window['AscDFH'].historydescription_CommonControllerUnGroup                     = 0x0006;
	window['AscDFH'].historydescription_CommonControllerCheckSelected               = 0x0007;
	window['AscDFH'].historydescription_CommonControllerSetGraphicObject            = 0x0008;
	window['AscDFH'].historydescription_CommonStatesAddNewShape                     = 0x0009;
	window['AscDFH'].historydescription_CommonStatesRotate                          = 0x000a;
	window['AscDFH'].historydescription_PasteNative                                 = 0x000b;
	window['AscDFH'].historydescription_Document_GroupUnGroup                       = 0x000c;
	window['AscDFH'].historydescription_Document_SetDefaultLanguage                 = 0x000d;
	window['AscDFH'].historydescription_Document_ChangeColorScheme                  = 0x000e;
	window['AscDFH'].historydescription_Document_AddChart                           = 0x000f;
	window['AscDFH'].historydescription_Document_EditChart                          = 0x0010;
	window['AscDFH'].historydescription_Document_DragText                           = 0x0011;
	window['AscDFH'].historydescription_Document_DocumentContentExtendToPos         = 0x0012;
	window['AscDFH'].historydescription_Document_AddHeader                          = 0x0013;
	window['AscDFH'].historydescription_Document_AddFooter                          = 0x0014;
	window['AscDFH'].historydescription_Document_ParagraphExtendToPos               = 0x0015;
	window['AscDFH'].historydescription_Document_ParagraphChangeFrame               = 0x0016;
	window['AscDFH'].historydescription_Document_ReplaceAll                         = 0x0017;
	window['AscDFH'].historydescription_Document_ReplaceSingle                      = 0x0018;
	window['AscDFH'].historydescription_Document_TableAddNewRowByTab                = 0x0019;
	window['AscDFH'].historydescription_Document_AddNewShape                        = 0x001a;
	window['AscDFH'].historydescription_Document_EditWrapPolygon                    = 0x001b;
	window['AscDFH'].historydescription_Document_MoveInlineObject                   = 0x001c;
	window['AscDFH'].historydescription_Document_CopyAndMoveInlineObject            = 0x001d;
	window['AscDFH'].historydescription_Document_RotateInlineDrawing                = 0x001e;
	window['AscDFH'].historydescription_Document_RotateFlowDrawingCtrl              = 0x001f;
	window['AscDFH'].historydescription_Document_RotateFlowDrawingNoCtrl            = 0x0020;
	window['AscDFH'].historydescription_Document_MoveInGroup                        = 0x0021;
	window['AscDFH'].historydescription_Document_ChangeWrapContour                  = 0x0022;
	window['AscDFH'].historydescription_Document_ChangeWrapContourAddPoint          = 0x0023;
	window['AscDFH'].historydescription_Document_GrObjectsBringToFront              = 0x0024;
	window['AscDFH'].historydescription_Document_GrObjectsBringForwardGroup         = 0x0025;
	window['AscDFH'].historydescription_Document_GrObjectsBringForward              = 0x0026;
	window['AscDFH'].historydescription_Document_GrObjectsSendToBackGroup           = 0x0027;
	window['AscDFH'].historydescription_Document_GrObjectsSendToBack                = 0x0028;
	window['AscDFH'].historydescription_Document_GrObjectsBringBackwardGroup        = 0x0029;
	window['AscDFH'].historydescription_Document_GrObjectsBringBackward             = 0x002a;
	window['AscDFH'].historydescription_Document_GrObjectsChangeWrapPolygon         = 0x002b;
	window['AscDFH'].historydescription_Document_MathAutoCorrect                    = 0x002c;
	window['AscDFH'].historydescription_Document_SetFramePrWithFontFamily           = 0x002d;
	window['AscDFH'].historydescription_Document_SetFramePr                         = 0x002e;
	window['AscDFH'].historydescription_Document_SetFramePrWithFontFamilyLong       = 0x002f;
	window['AscDFH'].historydescription_Document_SetTextFontName                    = 0x0030;
	window['AscDFH'].historydescription_Document_SetTextFontSize                    = 0x0031;
	window['AscDFH'].historydescription_Document_SetTextBold                        = 0x0032;
	window['AscDFH'].historydescription_Document_SetTextItalic                      = 0x0033;
	window['AscDFH'].historydescription_Document_SetTextUnderline                   = 0x0034;
	window['AscDFH'].historydescription_Document_SetTextStrikeout                   = 0x0035;
	window['AscDFH'].historydescription_Document_SetTextDStrikeout                  = 0x0036;
	window['AscDFH'].historydescription_Document_SetTextSpacing                     = 0x0037;
	window['AscDFH'].historydescription_Document_SetTextCaps                        = 0x0038;
	window['AscDFH'].historydescription_Document_SetTextSmallCaps                   = 0x0039;
	window['AscDFH'].historydescription_Document_SetTextPosition                    = 0x003a;
	window['AscDFH'].historydescription_Document_SetTextLang                        = 0x003b;
	window['AscDFH'].historydescription_Document_SetParagraphLineSpacing            = 0x003c;
	window['AscDFH'].historydescription_Document_SetParagraphLineSpacingBeforeAfter = 0x003d;
	window['AscDFH'].historydescription_Document_IncFontSize                        = 0x003e;
	window['AscDFH'].historydescription_Document_DecFontSize                        = 0x003f;
	window['AscDFH'].historydescription_Document_SetParagraphBorders                = 0x0040;
	window['AscDFH'].historydescription_Document_SetParagraphPr                     = 0x0041;
	window['AscDFH'].historydescription_Document_SetParagraphAlign                  = 0x0042;
	window['AscDFH'].historydescription_Document_SetTextVertAlign                   = 0x0043;
	window['AscDFH'].historydescription_Document_SetParagraphNumbering              = 0x0044;
	window['AscDFH'].historydescription_Document_SetParagraphStyle                  = 0x0045;
	window['AscDFH'].historydescription_Document_SetParagraphPageBreakBefore        = 0x0046;
	window['AscDFH'].historydescription_Document_SetParagraphWidowControl           = 0x0047;
	window['AscDFH'].historydescription_Document_SetParagraphKeepLines              = 0x0048;
	window['AscDFH'].historydescription_Document_SetParagraphKeepNext               = 0x0049;
	window['AscDFH'].historydescription_Document_SetParagraphContextualSpacing      = 0x004a;
	window['AscDFH'].historydescription_Document_SetTextHighlightNone               = 0x004b;
	window['AscDFH'].historydescription_Document_SetTextHighlightColor              = 0x004c;
	window['AscDFH'].historydescription_Document_SetTextColor                       = 0x004d;
	window['AscDFH'].historydescription_Document_SetParagraphShd                    = 0x004e;
	window['AscDFH'].historydescription_Document_SetParagraphIndent                 = 0x004f;
	window['AscDFH'].historydescription_Document_IncParagraphIndent                 = 0x0050;
	window['AscDFH'].historydescription_Document_DecParagraphIndent                 = 0x0051;
	window['AscDFH'].historydescription_Document_SetParagraphIndentRight            = 0x0052;
	window['AscDFH'].historydescription_Document_SetParagraphIndentFirstLine        = 0x0053;
	window['AscDFH'].historydescription_Document_SetPageOrientation                 = 0x0054;
	window['AscDFH'].historydescription_Document_SetPageSize                        = 0x0055;
	window['AscDFH'].historydescription_Document_AddPageBreak                       = 0x0056;
	window['AscDFH'].historydescription_Document_AddPageNumToHdrFtr                 = 0x0057;
	window['AscDFH'].historydescription_Document_AddPageNumToCurrentPos             = 0x0058;
	window['AscDFH'].historydescription_Document_SetHdrFtrDistance                  = 0x0059;
	window['AscDFH'].historydescription_Document_SetHdrFtrFirstPage                 = 0x005a;
	window['AscDFH'].historydescription_Document_SetHdrFtrEvenAndOdd                = 0x005b;
	window['AscDFH'].historydescription_Document_SetHdrFtrLink                      = 0x005c;
	window['AscDFH'].historydescription_Document_AddTable                           = 0x005d;
	window['AscDFH'].historydescription_Document_TableAddRowAbove                   = 0x005e;
	window['AscDFH'].historydescription_Document_TableAddRowBelow                   = 0x005f;
	window['AscDFH'].historydescription_Document_TableAddColumnLeft                 = 0x0060;
	window['AscDFH'].historydescription_Document_TableAddColumnRight                = 0x0061;
	window['AscDFH'].historydescription_Document_TableRemoveRow                     = 0x0062;
	window['AscDFH'].historydescription_Document_TableRemoveColumn                  = 0x0063;
	window['AscDFH'].historydescription_Document_RemoveTable                        = 0x0064;
	window['AscDFH'].historydescription_Document_MergeTableCells                    = 0x0065;
	window['AscDFH'].historydescription_Document_SplitTableCells                    = 0x0066;
	window['AscDFH'].historydescription_Document_ApplyTablePr                       = 0x0067;
	window['AscDFH'].historydescription_Document_AddImageUrl                        = 0x0068;
	window['AscDFH'].historydescription_Document_AddImageUrlLong                    = 0x0069;
	window['AscDFH'].historydescription_Document_AddImageToPage                     = 0x006a;
	window['AscDFH'].historydescription_Document_ApplyImagePrWithUrl                = 0x006b;
	window['AscDFH'].historydescription_Document_ApplyImagePrWithUrlLong            = 0x006c;
	window['AscDFH'].historydescription_Document_ApplyImagePrWithFillUrl            = 0x006d;
	window['AscDFH'].historydescription_Document_ApplyImagePrWithFillUrlLong        = 0x006e;
	window['AscDFH'].historydescription_Document_ApplyImagePr                       = 0x006f;
	window['AscDFH'].historydescription_Document_AddHyperlink                       = 0x0070;
	window['AscDFH'].historydescription_Document_ChangeHyperlink                    = 0x0071;
	window['AscDFH'].historydescription_Document_RemoveHyperlink                    = 0x0072;
	window['AscDFH'].historydescription_Document_ReplaceMisspelledWord              = 0x0073;
	window['AscDFH'].historydescription_Document_AddComment                         = 0x0074;
	window['AscDFH'].historydescription_Document_RemoveComment                      = 0x0075;
	window['AscDFH'].historydescription_Document_ChangeComment                      = 0x0076;
	window['AscDFH'].historydescription_Document_SetTextFontNameLong                = 0x0077;
	window['AscDFH'].historydescription_Document_AddImage                           = 0x0078;
	window['AscDFH'].historydescription_Document_ClearFormatting                    = 0x0079;
	window['AscDFH'].historydescription_Document_AddSectionBreak                    = 0x007a;
	window['AscDFH'].historydescription_Document_AddMath                            = 0x007b;
	window['AscDFH'].historydescription_Document_SetParagraphTabs                   = 0x007c;
	window['AscDFH'].historydescription_Document_SetParagraphIndentFromRulers       = 0x007d;
	window['AscDFH'].historydescription_Document_SetDocumentMargin_Hor              = 0x007e;
	window['AscDFH'].historydescription_Document_SetTableMarkup_Hor                 = 0x007f;
	window['AscDFH'].historydescription_Document_SetDocumentMargin_Ver              = 0x0080;
	window['AscDFH'].historydescription_Document_SetHdrFtrBounds                    = 0x0081;
	window['AscDFH'].historydescription_Document_SetTableMarkup_Ver                 = 0x0082;
	window['AscDFH'].historydescription_Document_DocumentExtendToPos                = 0x0083;
	window['AscDFH'].historydescription_Document_AddDropCap                         = 0x0084;
	window['AscDFH'].historydescription_Document_RemoveDropCap                      = 0x0085;
	window['AscDFH'].historydescription_Document_SetTextHighlight                   = 0x0086;
	window['AscDFH'].historydescription_Document_BackSpaceButton                    = 0x0087;
	window['AscDFH'].historydescription_Document_MoveParagraphByTab                 = 0x0088;
	window['AscDFH'].historydescription_Document_AddTab                             = 0x0089;
	window['AscDFH'].historydescription_Document_EnterButton                        = 0x008a;
	window['AscDFH'].historydescription_Document_SpaceButton                        = 0x008b;
	window['AscDFH'].historydescription_Document_ShiftInsert                        = 0x008c;
	window['AscDFH'].historydescription_Document_ShiftInsertSafari                  = 0x008d;
	window['AscDFH'].historydescription_Document_DeleteButton                       = 0x008e;
	window['AscDFH'].historydescription_Document_ShiftDeleteButton                  = 0x008f;
	window['AscDFH'].historydescription_Document_Shortcut_SetStyleHeading1          = 0x0090;
	window['AscDFH'].historydescription_Document_Shortcut_SetStyleHeading2          = 0x0091;
	window['AscDFH'].historydescription_Document_Shortcut_SetStyleHeading3          = 0x0092;
	window['AscDFH'].historydescription_Document_SetTextStrikeoutHotKey             = 0x0093;
	window['AscDFH'].historydescription_Document_SetTextBoldHotKey                  = 0x0094;
	window['AscDFH'].historydescription_Document_SetParagraphAlignHotKey            = 0x0095;
	window['AscDFH'].historydescription_Document_AddEuroLetter                      = 0x0096;
	window['AscDFH'].historydescription_Document_SetTextItalicHotKey                = 0x0097;
	window['AscDFH'].historydescription_Document_SetParagraphAlignHotKey2           = 0x0098;
	window['AscDFH'].historydescription_Document_SetParagraphNumberingHotKey        = 0x0099;
	window['AscDFH'].historydescription_Document_SetParagraphAlignHotKey3           = 0x009a;
	window['AscDFH'].historydescription_Document_AddPageNumHotKey                   = 0x009b;
	window['AscDFH'].historydescription_Document_SetParagraphAlignHotKey4           = 0x009c;
	window['AscDFH'].historydescription_Document_SetTextUnderlineHotKey             = 0x009d;
	window['AscDFH'].historydescription_Document_FormatPasteHotKey                  = 0x009e;
	window['AscDFH'].historydescription_Document_PasteHotKey                        = 0x009f;
	window['AscDFH'].historydescription_Document_PasteSafariHotKey                  = 0x00a0;
	window['AscDFH'].historydescription_Document_CutHotKey                          = 0x00a1;
	window['AscDFH'].historydescription_Document_SetTextVertAlignHotKey             = 0x00a2;
	window['AscDFH'].historydescription_Document_AddMathHotKey                      = 0x00a3;
	window['AscDFH'].historydescription_Document_SetTextVertAlignHotKey2            = 0x00a4;
	window['AscDFH'].historydescription_Document_MinusButton                        = 0x00a5;
	window['AscDFH'].historydescription_Document_SetTextVertAlignHotKey3            = 0x00a6;
	window['AscDFH'].historydescription_Document_AddLetter                          = 0x00a7;
	window['AscDFH'].historydescription_Document_MoveTableBorder                    = 0x00a8;
	window['AscDFH'].historydescription_Document_FormatPasteHotKey2                 = 0x00a9;
	window['AscDFH'].historydescription_Document_SetTextHighlight2                  = 0x00aa;
	window['AscDFH'].historydescription_Document_AddTextFromTextBox                 = 0x00ab;
	window['AscDFH'].historydescription_Document_AddMailMergeField                  = 0x00ac;
	window['AscDFH'].historydescription_Document_MoveInlineTable                    = 0x00ad;
	window['AscDFH'].historydescription_Document_MoveFlowTable                      = 0x00ae;
	window['AscDFH'].historydescription_Document_RestoreFieldTemplateText           = 0x00af;
	window['AscDFH'].historydescription_Spreadsheet_SetCellFontName                 = 0x00b0;
	window['AscDFH'].historydescription_Spreadsheet_SetCellFontSize                 = 0x00b1;
	window['AscDFH'].historydescription_Spreadsheet_SetCellBold                     = 0x00b2;
	window['AscDFH'].historydescription_Spreadsheet_SetCellItalic                   = 0x00b3;
	window['AscDFH'].historydescription_Spreadsheet_SetCellUnderline                = 0x00b4;
	window['AscDFH'].historydescription_Spreadsheet_SetCellStrikeout                = 0x00b5;
	window['AscDFH'].historydescription_Spreadsheet_SetCellSubscript                = 0x00b6;
	window['AscDFH'].historydescription_Spreadsheet_SetCellSuperscript              = 0x00b7;
	window['AscDFH'].historydescription_Spreadsheet_SetCellAlign                    = 0x00b8;
	window['AscDFH'].historydescription_Spreadsheet_SetCellVertAlign                = 0x00b9;
	window['AscDFH'].historydescription_Spreadsheet_SetCellTextColor                = 0x00ba;
	window['AscDFH'].historydescription_Spreadsheet_SetCellBackgroundColor          = 0x00bb;
	window['AscDFH'].historydescription_Spreadsheet_SetCellIncreaseFontSize         = 0x00bc;
	window['AscDFH'].historydescription_Spreadsheet_SetCellDecreaseFontSize         = 0x00bd;
	window['AscDFH'].historydescription_Spreadsheet_SetCellHyperlinkAdd             = 0x00be;
	window['AscDFH'].historydescription_Spreadsheet_SetCellHyperlinkModify          = 0x00bf;
	window['AscDFH'].historydescription_Spreadsheet_SetCellHyperlinkRemove          = 0x00c0;
	window['AscDFH'].historydescription_Spreadsheet_EditChart                       = 0x00c1;
	window['AscDFH'].historydescription_Spreadsheet_Remove                          = 0x00c2;
	window['AscDFH'].historydescription_Spreadsheet_AddTab                          = 0x00c3;
	window['AscDFH'].historydescription_Spreadsheet_AddNewParagraph                 = 0x00c4;
	window['AscDFH'].historydescription_Spreadsheet_AddSpace                        = 0x00c5;
	window['AscDFH'].historydescription_Spreadsheet_AddItem                         = 0x00c6;
	window['AscDFH'].historydescription_Spreadsheet_PutPrLineSpacing                = 0x00c7;
	window['AscDFH'].historydescription_Spreadsheet_SetParagraphSpacing             = 0x00c8;
	window['AscDFH'].historydescription_Spreadsheet_SetGraphicObjectsProps          = 0x00c9;
	window['AscDFH'].historydescription_Spreadsheet_ParaApply                       = 0x00ca;
	window['AscDFH'].historydescription_Spreadsheet_GraphicObjectLayer              = 0x00cb;
	window['AscDFH'].historydescription_Spreadsheet_ParagraphAdd                    = 0x00cc;
	window['AscDFH'].historydescription_Spreadsheet_CreateGroup                     = 0x00cd;
	window['AscDFH'].historydescription_CommonDrawings_ChangeAdj                    = 0x00ce;
	window['AscDFH'].historydescription_CommonDrawings_EndTrack                     = 0x00cf;
	window['AscDFH'].historydescription_CommonDrawings_CopyCtrl                     = 0x00d0;
	window['AscDFH'].historydescription_Presentation_ParaApply                      = 0x00d1;
	window['AscDFH'].historydescription_Presentation_ParaFormatPaste                = 0x00d2;
	window['AscDFH'].historydescription_Presentation_AddNewParagraph                = 0x00d3;
	window['AscDFH'].historydescription_Presentation_CreateGroup                    = 0x00d4;
	window['AscDFH'].historydescription_Presentation_UnGroup                        = 0x00d5;
	window['AscDFH'].historydescription_Presentation_AddChart                       = 0x00d6;
	window['AscDFH'].historydescription_Presentation_EditChart                      = 0x00d7;
	window['AscDFH'].historydescription_Presentation_ParagraphAdd                   = 0x00d8;
	window['AscDFH'].historydescription_Presentation_ParagraphClearFormatting       = 0x00d9;
	window['AscDFH'].historydescription_Presentation_SetParagraphAlign              = 0x00da;
	window['AscDFH'].historydescription_Presentation_SetParagraphSpacing            = 0x00db;
	window['AscDFH'].historydescription_Presentation_SetParagraphTabs               = 0x00dc;
	window['AscDFH'].historydescription_Presentation_SetParagraphIndent             = 0x00dd;
	window['AscDFH'].historydescription_Presentation_SetParagraphNumbering          = 0x00de;
	window['AscDFH'].historydescription_Presentation_ParagraphIncDecFontSize        = 0x00df;
	window['AscDFH'].historydescription_Presentation_ParagraphIncDecIndent          = 0x00e0;
	window['AscDFH'].historydescription_Presentation_SetImageProps                  = 0x00e1;
	window['AscDFH'].historydescription_Presentation_SetShapeProps                  = 0x00e2;
	window['AscDFH'].historydescription_Presentation_ChartApply                     = 0x00e3;
	window['AscDFH'].historydescription_Presentation_ChangeShapeType                = 0x00e4;
	window['AscDFH'].historydescription_Presentation_SetVerticalAlign               = 0x00e5;
	window['AscDFH'].historydescription_Presentation_HyperlinkAdd                   = 0x00e6;
	window['AscDFH'].historydescription_Presentation_HyperlinkModify                = 0x00e7;
	window['AscDFH'].historydescription_Presentation_HyperlinkRemove                = 0x00e8;
	window['AscDFH'].historydescription_Presentation_DistHor                        = 0x00e9;
	window['AscDFH'].historydescription_Presentation_DistVer                        = 0x00ea;
	window['AscDFH'].historydescription_Presentation_BringToFront                   = 0x00eb;
	window['AscDFH'].historydescription_Presentation_BringForward                   = 0x00ec;
	window['AscDFH'].historydescription_Presentation_SendToBack                     = 0x00ed;
	window['AscDFH'].historydescription_Presentation_BringBackward                  = 0x00ef;
	window['AscDFH'].historydescription_Presentation_ApplyTransition                = 0x00f0;
	window['AscDFH'].historydescription_Presentation_MoveSlidesToEnd                = 0x00f1;
	window['AscDFH'].historydescription_Presentation_MoveSlidesNextPos              = 0x00f2;
	window['AscDFH'].historydescription_Presentation_MoveSlidesPrevPos              = 0x00f3;
	window['AscDFH'].historydescription_Presentation_MoveSlidesToStart              = 0x00f4;
	window['AscDFH'].historydescription_Presentation_MoveComments                   = 0x00f5;
	window['AscDFH'].historydescription_Presentation_TableBorder                    = 0x00f6;
	window['AscDFH'].historydescription_Presentation_AddFlowImage                   = 0x00f7;
	window['AscDFH'].historydescription_Presentation_AddFlowTable                   = 0x00f8;
	window['AscDFH'].historydescription_Presentation_ChangeBackground               = 0x00f9;
	window['AscDFH'].historydescription_Presentation_AddNextSlide                   = 0x00fa;
	window['AscDFH'].historydescription_Presentation_ShiftSlides                    = 0x00fb;
	window['AscDFH'].historydescription_Presentation_DeleteSlides                   = 0x00fc;
	window['AscDFH'].historydescription_Presentation_ChangeLayout                   = 0x00fd;
	window['AscDFH'].historydescription_Presentation_ChangeSlideSize                = 0x00fe;
	window['AscDFH'].historydescription_Presentation_ChangeColorScheme              = 0x00ff;
	window['AscDFH'].historydescription_Presentation_AddComment                     = 0x0100;
	window['AscDFH'].historydescription_Presentation_ChangeComment                  = 0x0101;
	window['AscDFH'].historydescription_Presentation_PutTextPrFontName              = 0x0102;
	window['AscDFH'].historydescription_Presentation_PutTextPrFontSize              = 0x0103;
	window['AscDFH'].historydescription_Presentation_PutTextPrBold                  = 0x0104;
	window['AscDFH'].historydescription_Presentation_PutTextPrItalic                = 0x0105;
	window['AscDFH'].historydescription_Presentation_PutTextPrUnderline             = 0x0106;
	window['AscDFH'].historydescription_Presentation_PutTextPrStrikeout             = 0x0107;
	window['AscDFH'].historydescription_Presentation_PutTextPrLineSpacing           = 0x0108;
	window['AscDFH'].historydescription_Presentation_PutTextPrSpacingBeforeAfter    = 0x0109;
	window['AscDFH'].historydescription_Presentation_PutTextPrIncreaseFontSize      = 0x010a;
	window['AscDFH'].historydescription_Presentation_PutTextPrDecreaseFontSize      = 0x010b;
	window['AscDFH'].historydescription_Presentation_PutTextPrAlign                 = 0x010c;
	window['AscDFH'].historydescription_Presentation_PutTextPrBaseline              = 0x010d;
	window['AscDFH'].historydescription_Presentation_PutTextPrListType              = 0x010e;
	window['AscDFH'].historydescription_Presentation_PutTextColor                   = 0x010f;
	window['AscDFH'].historydescription_Presentation_PutTextColor2                  = 0x010f;
	window['AscDFH'].historydescription_Presentation_PutPrIndent                    = 0x010f;
	window['AscDFH'].historydescription_Presentation_PutPrIndentRight               = 0x010f;
	window['AscDFH'].historydescription_Presentation_PutPrFirstLineIndent           = 0x010f;
	window['AscDFH'].historydescription_Presentation_AddPageBreak                   = 0x010f;
	window['AscDFH'].historydescription_Presentation_AddRowAbove                    = 0x0110;
	window['AscDFH'].historydescription_Presentation_AddRowBelow                    = 0x0111;
	window['AscDFH'].historydescription_Presentation_AddColLeft                     = 0x0112;
	window['AscDFH'].historydescription_Presentation_AddColRight                    = 0x0113;
	window['AscDFH'].historydescription_Presentation_RemoveRow                      = 0x0114;
	window['AscDFH'].historydescription_Presentation_RemoveCol                      = 0x0115;
	window['AscDFH'].historydescription_Presentation_RemoveTable                    = 0x0116;
	window['AscDFH'].historydescription_Presentation_MergeCells                     = 0x0117;
	window['AscDFH'].historydescription_Presentation_SplitCells                     = 0x0118;
	window['AscDFH'].historydescription_Presentation_TblApply                       = 0x0119;
	window['AscDFH'].historydescription_Presentation_RemoveComment                  = 0x011a;
	window['AscDFH'].historydescription_Presentation_EndFontLoad                    = 0x011b;
	window['AscDFH'].historydescription_Presentation_ChangeTheme                    = 0x011c;
	window['AscDFH'].historydescription_Presentation_TableMoveFromRulers            = 0x011d;
	window['AscDFH'].historydescription_Presentation_TableMoveFromRulersInline      = 0x011e;
	window['AscDFH'].historydescription_Presentation_PasteOnThumbnails              = 0x011f;
	window['AscDFH'].historydescription_Presentation_PasteOnThumbnailsSafari        = 0x0120;
	window['AscDFH'].historydescription_Document_ConvertOldEquation                 = 0x0121;
	window['AscDFH'].historydescription_Presentation_SetVert                        = 0x0122;
	window['AscDFH'].historydescription_Document_AddNewStyle                        = 0x0123;
	window['AscDFH'].historydescription_Document_RemoveStyle                        = 0x0124;
	window['AscDFH'].historydescription_Document_AddTextArt                         = 0x0125;
	window['AscDFH'].historydescription_Document_RemoveAllCustomStyles              = 0x0126;
	window['AscDFH'].historydescription_Document_AcceptAllRevisionChanges           = 0x0127;
	window['AscDFH'].historydescription_Document_RejectAllRevisionChanges           = 0x0128;
	window['AscDFH'].historydescription_Document_AcceptRevisionChange               = 0x0129;
	window['AscDFH'].historydescription_Document_RejectRevisionChange               = 0x012a;
	window['AscDFH'].historydescription_Document_AcceptRevisionChangesBySelection   = 0x012b;
	window['AscDFH'].historydescription_Document_RejectRevisionChangesBySelection   = 0x012c;
	window['AscDFH'].historydescription_Document_AddLetterUnion                     = 0x012d;
	window['AscDFH'].historydescription_Presentation_ApplyTransitionToAll           = 0x012e;
	window['AscDFH'].historydescription_Document_SetColumnsFromRuler                = 0x012f;
	window['AscDFH'].historydescription_Document_SetColumnsProps                    = 0x0130;
	window['AscDFH'].historydescription_Document_AddColumnBreak                     = 0x0131;
	window['AscDFH'].historydescription_Document_SetSectionProps                    = 0x0132;
	window['AscDFH'].historydescription_Document_AddTabToMath                       = 0x0133;
	window['AscDFH'].historydescription_Document_SetMathProps                       = 0x0134;
	window['AscDFH'].historydescription_Document_ApplyPrToMath                      = 0x0135;
	window['AscDFH'].historydescription_Document_ApiBuilder                         = 0x0136;
	window['AscDFH'].historydescription_Document_AddOleObject                       = 0x0137;
	window['AscDFH'].historydescription_Document_EditOleObject                      = 0x0138;
	window['AscDFH'].historydescription_Document_CompositeInput                     = 0x0139;
	window['AscDFH'].historydescription_Document_CompositeInputReplace              = 0x013a;
	window['AscDFH'].historydescription_Document_AddPageCount                       = 0x013b;
	window['AscDFH'].historydescription_Document_AddFootnote                        = 0x013c;
	window['AscDFH'].historydescription_Document_SetFootnotePr                      = 0x013d;
	window['AscDFH'].historydescription_Document_RemoveAllFootnotes                 = 0x013e;
	window['AscDFH'].historydescription_Document_InsertDocumentsByUrls              = 0x013f;
	window['AscDFH'].historydescription_Document_InsertSignatureLine                = 0x0140;
	window['AscDFH'].historydescription_Document_AddBlockLevelContentControl        = 0x0141;
	window['AscDFH'].historydescription_Document_AddInlineLevelContentControl       = 0x0142;
	window['AscDFH'].historydescription_Document_RemoveContentControl               = 0x0143;
	window['AscDFH'].historydescription_Document_RemoveContentControlWrapper        = 0x0144;
	window['AscDFH'].historydescription_Document_ChangeContentControlProperties     = 0x0145;
    window['AscDFH'].historydescription_Presentation_HideSlides                     = 0x0146;
	window['AscDFH'].historydescription_DocumentMacros_Data                         = 0x0147;
    window['AscDFH'].historydescription_Document_AddBookmark                        = 0x0148;
	window['AscDFH'].historydescription_Document_AddTableOfContents                 = 0x0149;
	window['AscDFH'].historydescription_Document_ChangeOutlineLevel                 = 0x014a;
	window['AscDFH'].historydescription_Document_AddElementToOutline                = 0x014b;
	window['AscDFH'].historydescription_Document_ResizeTable                        = 0x014c;
	window['AscDFH'].historydescription_Document_RemoveComplexField                 = 0x014d;
	window['AscDFH'].historydescription_Document_SetComplexFieldPr                  = 0x014e;
	window['AscDFH'].historydescription_Document_UpdateTableOfContents              = 0x014f;
	window['AscDFH'].historydescription_Document_SectionStartPage                   = 0x0150;
	window['AscDFH'].historydescription_Document_DistributeTableCells               = 0x0151;
	window['AscDFH'].historydescription_Document_RemoveBookmark                     = 0x0152;
	window['AscDFH'].historydescription_Document_ContinueNumbering                  = 0x0153;
	window['AscDFH'].historydescription_Document_RestartNumbering                   = 0x0154;
	window['AscDFH'].historydescription_Document_AutomaticListAsType                = 0x0155;
	window['AscDFH'].historydescription_Document_CreateNum                          = 0x0156;
	window['AscDFH'].historydescription_Document_ChangeNumLvl                       = 0x0157;
	window['AscDFH'].historydescription_Document_AutoCorrectSmartQuotes             = 0x0158;
	window['AscDFH'].historydescription_Document_AutoCorrectHyphensWithDash         = 0x0159;
	window['AscDFH'].historydescription_Document_SetGlobalSdtHighlightColor         = 0x015a;
	window['AscDFH'].historydescription_Document_SetGlobalSdtShowHighlight          = 0x015b;
	window['AscDFH'].historydescription_Document_UpdateFields                       = 0x015c;
	window['AscDFH'].historydescription_Document_AddBlankPage                       = 0x015d;
	window['AscDFH'].historydescription_Document_AddTableFormula                    = 0x015e;
	window['AscDFH'].historydescription_Document_ChangeTableFormula                 = 0x015e;
	window['AscDFH'].historydescription_SetCoreproperties                           = 0x015f;
	window['AscDFH'].historydescription_Document_AddWatermark                       = 0x0160;
	window['AscDFH'].historydescription_Presentation_SetHF                          = 0x0161;
	window['AscDFH'].historydescription_Presentation_AddSlideNumber                 = 0x0162;
	window['AscDFH'].historydescription_Document_SetParagraphOutlineLvl             = 0x0163;
	window['AscDFH'].historydescription_Document_RemoveTableCells                   = 0x0164;
	window['AscDFH'].historydescription_Document_AddContentControlCheckBox          = 0x0165;
	window['AscDFH'].historydescription_Document_SetContentControlCheckBoxPr        = 0x0166;
	window['AscDFH'].historydescription_Document_AddContentControlPicture           = 0x0167;
	window['AscDFH'].historydescription_Document_SetContentControlPictureUrl        = 0x0168;
	window['AscDFH'].historydescription_Document_RemoveAllComments                  = 0x0169;
	window['AscDFH'].historydescription_Document_AddContentControlList              = 0x016a;
	window['AscDFH'].historydescription_Document_SetContentControlListPr            = 0x016b;
	window['AscDFH'].historydescription_Document_SelectContentControlListItem       = 0x016c;
	window['AscDFH'].historydescription_Document_AddContentControlDatePicker        = 0x016d;
	window['AscDFH'].historydescription_Document_SetContentControlDatePickerPr      = 0x016e;
	window['AscDFH'].historydescription_Presentation_AddToLayout                    = 0x016f;
	window['AscDFH'].historydescription_Presentation_FitImagesToSlide               = 0x0170;
	window['AscDFH'].historydescription_Document_AddTextWithProperties              = 0x0171;
	window['AscDFH'].historydescription_Document_AddCaption                         = 0x0172;
	window['AscDFH'].historydescription_Document_CompareDocuments                   = 0x0173;
	window['AscDFH'].historydescription_Document_DrawNewTable                       = 0x0174;
	window['AscDFH'].historydescription_Document_DrawTable                          = 0x0175;
	window['AscDFH'].historydescription_Document_AddDateTimeField                   = 0x0176;
	window['AscDFH'].historydescription_Document_SetContentControlTextPlaceholder   = 0x0177;
	window['AscDFH'].historydescription_Document_AddEndnote                         = 0x0178;
	window['AscDFH'].historydescription_Document_AddContentControlTextForm          = 0x0179;
	window['AscDFH'].historydescription_Document_SetEndnotePr                       = 0x017a;
	window['AscDFH'].historydescription_Document_ConvertFootnoteType                = 0x017b;
	window['AscDFH'].historydescription_Document_AutoCorrectCommon                  = 0x017c;
	window['AscDFH'].historydescription_Document_Shortcut_ClearFormatting           = 0x017d;
	window['AscDFH'].historydescription_Document_Shortcut_AddNonBreakingSpace       = 0x017e;
	window['AscDFH'].historydescription_Document_SetParagraphSuppressLineNumbers    = 0x017f;
	window['AscDFH'].historydescription_Document_SetLineNumbersProps                = 0x0180;
	window['AscDFH'].historydescription_Document_AddCrossRef                        = 0x0181;
	window['AscDFH'].historydescription_Document_ClearAllSpecialForms               = 0x0182;
	window['AscDFH'].historydescription_Document_ChangeTextCase                     = 0x0183;
	window['AscDFH'].historydescription_Document_SetNumberingLvl                    = 0x0184;
	window['AscDFH'].historydescription_Document_SetTrackRevisions                  = 0x0185;
	window['AscDFH'].historydescription_Document_SetContentControlText              = 0x0186;
	window['AscDFH'].historydescription_Document_ClearContentControl                = 0x0187;
	window['AscDFH'].historydescription_Document_AutoCorrectFirstLetterOfSentence   = 0x0188;
	window['AscDFH'].historydescription_Document_ConvertTextToTable                 = 0x0189;
	window['AscDFH'].historydescription_Document_ConvertTableToText                 = 0x018a;
	window['AscDFH'].historydescription_Document_ResolveAllComments                 = 0x018b;
	window['AscDFH'].historydescription_Document_ConvertFormFixedType               = 0x018c;
	window['AscDFH'].historydescription_Document_ReplaceCurrentWord                 = 0x018d;
	window['AscDFH'].historydescription_Document_ChangeGeometryEdit                 = 0x018e;
	window['AscDFH'].historydescription_Document_Docxf_To_Docx                      = 0x018f;
	window['AscDFH'].historydescription_Document_ConvertMathDisplayMode             = 0x0190;
	window['AscDFH'].historydescription_Document_RemoveHdrFtr                       = 0x0191;
	window['AscDFH'].historydescription_Document_AddParagraphToTOC                  = 0x0192;
	window['AscDFH'].historydescription_Document_FillFormsByTags                    = 0x0193;
	window['AscDFH'].historydescription_Document_FillFormInPlugin                   = 0x0194;
	window['AscDFH'].historydescription_Document_AddSmartArt                        = 0x0195;
	window['AscDFH'].historydescription_Document_AddComplexForm                     = 0x0196;
	window['AscDFH'].historydescription_Document_CorrectFormTextByFormat            = 0x0197;
	window['AscDFH'].historydescription_Document_CorrectEnterText                   = 0x0198;
	window['AscDFH'].historydescription_Document_ConvertMathView                    = 0x0199;
	window['AscDFH'].historydescription_Document_DocumentProtection                 = 0x019a;
	window['AscDFH'].historydescription_Collaborative_Undo                          = 0x019b;
	window['AscDFH'].historydescription_Document_AddPlaceholderImages               = 0x019c;
	window['AscDFH'].historydescription_OForm_AddRole                               = 0x019d;
	window['AscDFH'].historydescription_OForm_RemoveRole                            = 0x019e;
	window['AscDFH'].historydescription_OForm_EditRole                              = 0x019f;
	window['AscDFH'].historydescription_OForm_ChangeRoleOrder                       = 0x01a0;
	window['AscDFH'].historydescription_Document_AddAddinField                      = 0x01a1;
	window['AscDFH'].historydescription_Document_UpdateAddinFields                  = 0x01a2;
	window['AscDFH'].historydescription_Document_RemoveComplexFieldWrapper          = 0x01a3;
	window['AscDFH'].historydescription_Document_MergeDocuments                     = 0x01a4;
	window['AscDFH'].historydescription_Document_FillContentControlPlaceholderOnBlur= 0x01a5;
	window['AscDFH'].historydescription_Document_SetAutoHyphenation                 = 0x01a6;
	window['AscDFH'].historydescription_Document_SetConsecutiveHyphenLimit          = 0x01a7;
	window['AscDFH'].historydescription_Document_SetHyphenateCaps                   = 0x01a8;
	window['AscDFH'].historydescription_Document_RemoveMathShortcut                 = 0x01a9;
	window['AscDFH'].historydescription_Document_SetAllFormsData                    = 0x01aa;
	window['AscDFH'].historydescription_Document_ComplexField_MergeFormat           = 0x01ab;
	window['AscDFH'].historydescription_Presentation_ResetSlideBackground           = 0x01ac;
	window['AscDFH'].historydescription_Presentation_ApplyBackgroundToAll           = 0x01ad;
	window['AscDFH'].historydescription_Presentation_ShowMasterShapes               = 0x01ae;
	window['AscDFH'].historydescription_BuilderScript                               = 0x01af;
	window['AscDFH'].historydescription_Document_AddRemoveBeforeAfterParagraph      = 0x01b0;
	window['AscDFH'].historydescription_Document_SectionPageNumFormat               = 0x01b1;
	window['AscDFH'].historydescription_Document_SetPageColor                       = 0x01b2;
	window['AscDFH'].historydescription_Document_InsertTextFromFile                 = 0x01b3;
	window['AscDFH'].historydescription_Document_AddComplexField                    = 0x01b4;
	window['AscDFH'].historydescription_Document_EditComplexFieldInstruction        = 0x01b5;
	window['AscDFH'].historydescription_Collaborative_DeletedTextRecovery           = 0x01b6;
	window['AscDFH'].historydescription_Document_AutoCorrectMath                    = 0x01b7;
	window['AscDFH'].historydescription_CustomProperties_Add                        = 0x01b8;
	window['AscDFH'].historydescription_CustomProperties_Remove                     = 0x01b9;
	window['AscDFH'].historydescription_CustomProperties_Modify                     = 0x01c0;
	window['AscDFH'].historydescription_Presentation_MergeSelectedShapes            = 0x01c1;
	window['AscDFH'].historydescription_Presentation_SaveAnnotations                = 0x01c2;
	window['AscDFH'].historydescription_Document_SetParagraphBidi                   = 0x01c3;
	window['AscDFH'].historydescription_RemoveAllInks                               = 0x01c4;
	window['AscDFH'].historydescription_DisconnectEveryone                          = 0x01c5;
	window['AscDFH'].historydescription_OForm_RoleFilled                            = 0x01c6;
	window['AscDFH'].historydescription_OForm_CompletePreparation                   = 0x01c7;
	window['AscDFH'].historydescription_Presentation_SetPreserveSlideMaster         = 0x01c8;
	window['AscDFH'].historydescription_Document_AddMathML                          = 0x01c9;
	window['AscDFH'].historydescription_OForm_CancelFilling                         = 0x01ca;

	// pdf
	window['AscDFH'].historydescription_Pdf_AddAnnot			= 0x29a;
	window['AscDFH'].historydescription_Pdf_FreeTextGeom		= 0x29b;
	window['AscDFH'].historydescription_Pdf_AddPage				= 0x29c;
	window['AscDFH'].historydescription_Pdf_RemovePage			= 0x29d;
	window['AscDFH'].historydescription_Pdf_AddHighlightAnnot	= 0x29e;
	window['AscDFH'].historydescription_Pdf_EraseInk			= 0x29f;
	window['AscDFH'].historydescription_Pdf_RemoveComment		= 0x2a0;
	window['AscDFH'].historydescription_Pdf_EditPage			= 0x2a1;
	window['AscDFH'].historydescription_Pdf_ContextMenuRemove	= 0x2a2;
	window['AscDFH'].historydescription_Pdf_RotatePage			= 0x2a3;
	window['AscDFH'].historydescription_Pdf_UpdateAnnotRC		= 0x2a4;
	window['AscDFH'].historydescription_Pdf_ClickCheckbox		= 0x2a5;
	window['AscDFH'].historydescription_Pdf_FieldCommit			= 0x2a6;
	window['AscDFH'].historydescription_Pdf_FieldImportImage	= 0x2a6;
	window['AscDFH'].historydescription_Pdf_FieldSelectOption	= 0x2a7;
	window['AscDFH'].historydescription_Pdf_ExecActions			= 0x2a8;
	window['AscDFH'].historydescription_Pdf_FreeTextFitTextBox	= 0x2a9;
	window['AscDFH'].historydescription_Pdf_AddComment			= 0x2b0;
	window['AscDFH'].historydescription_Pdf_ChangeStrokeColor	= 0x2b1;
	window['AscDFH'].historydescription_Pdf_ChangeFillColor		= 0x2b2;
	window['AscDFH'].historydescription_Pdf_ChangeOpacity		= 0x2b3;
	window['AscDFH'].historydescription_Pdf_MovePage			= 0x2b4;
	window['AscDFH'].historydescription_Pdf_AddField			= 0x2b5;
	window['AscDFH'].historydescription_Pdf_ChangeField			= 0x2b6;

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//
	// Фабрика изменений (заполняется там же, где и определяются классы изменений)
	//
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	window['AscDFH'].changesFactory = {};
	window['AscDFH'].changesFactory[window['AscDFH'].historyitem_Unknown_Unknown] = CChangesBase;
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//
	// Карта зависимости изменений. Изменения зависят только от изменений для того же класса, но вот типы могут быть
	// разными. В основном изменения зависят только от изменений такого же типа.
	//
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	window['AscDFH'].changesRelationMap = {};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//
	// Базовые классы для изменений
	//
	// Разница между классами Property и Value в том, что Property могут быть undefined, а Value всегда значение
	// заданного типа.
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	/**
	 * Базовый класс для всех изменений совместного редактирования.
	 * @constructor
	 */
	function CChangesBase(Class)
	{
		this.Class = Class;

		this.Reverted = false;
	}
	CChangesBase.prototype.Type = window['AscDFH'].historyitem_Unknown_Unknown;
	CChangesBase.prototype.GetType = function()
	{
		return this.Type;
	};
	CChangesBase.prototype.Undo = function()
	{
		if (this.Class && this.Class.Undo)
			this.Class.Undo(this);
	};
	CChangesBase.prototype.Redo = function()
	{
		if (this.Class && this.Class.Redo)
			this.Class.Redo(this);
	};
	CChangesBase.prototype.WriteToBinary = function(Writer)
	{
		if (this.Class && this.Class.Save_Changes)
			this.Class.Save_Changes(this, Writer);
	};
	CChangesBase.prototype.ReadFromBinary = function(Reader)
	{

	};
	CChangesBase.prototype.Load = function()
	{
		// В большинстве случаев загрузка чужого изменения работает как простое Redo
		this.Redo();
	};
	CChangesBase.prototype.GetClass = function()
	{
		return this.Class;
	};
	CChangesBase.prototype.RefreshRecalcData = function()
	{
		if (this.Class && this.Class.Refresh_RecalcData)
			this.Class.Refresh_RecalcData(this);
	};
	CChangesBase.prototype.IsContentChange = function()
	{
		return false;
	};
	CChangesBase.prototype.IsSpreadsheetChange = function()
	{
		return false;
	};
	CChangesBase.prototype.CreateReverseChange = function()
	{
		return null;
	};
	CChangesBase.prototype.Merge = function(oChange)
	{
		return true;
	};
	CChangesBase.prototype.IsPosExtChange = function(oChange)
	{
		return false;
	};
	CChangesBase.prototype.IsReverted = function()
	{
		return this.Reverted;
	};
	CChangesBase.prototype.SetReverted = function(isReverted)
	{
		this.Reverted = isReverted;
	};
	CChangesBase.prototype.IsParagraphSimpleChanges = function()
	{
		return false;
	};
	CChangesBase.prototype.IsNeedRecalculate = function()
	{
		return true;
	};
	CChangesBase.prototype.IsNeedRecalculateLineNumbers = function()
	{
		return false;
	};
	CChangesBase.prototype.IsDescriptionChange = function()
	{
		return false;
	};
	CChangesBase.prototype.CheckLock = function(lockData)
	{
	};
	CChangesBase.prototype.CheckNeedRecalculate = function()
	{
		if (!this.IsNeedRecalculate())
			return;
		
		let obj = this.GetClass();
		if (obj && obj.SetIsRecalculated)
			obj.SetIsRecalculated(false);
	};
	window['AscDFH'].CChangesBase = CChangesBase;
	/**
	 * Базовый класс для изменений, которые меняют содержимое родительского класса.*
	 * @constructor
	 * @extends {AscDFH.CChangesBase}
	 */
	function CChangesBaseContentChange(Class, Pos, Items, isAdd)
	{
		CChangesBase.call(this, Class);

		this.Pos      = Pos;
		this.Items    = Items;
		this.UseArray = false;
		this.PosArray = [];
		this.Add      = isAdd;

		this.Reverted = false;
	}

	CChangesBaseContentChange.prototype = Object.create(CChangesBase.prototype);
	CChangesBaseContentChange.prototype.constructor = CChangesBaseContentChange;
	CChangesBaseContentChange.prototype.IsContentChange = function()
	{
		return true;
	};
	CChangesBaseContentChange.prototype.GetContentChangesClass = function()
	{
		if (this.Class && this.Class.m_oContentChanges)
			return this.Class.m_oContentChanges;
		
		return null;
	};
	CChangesBaseContentChange.prototype.IsAdd = function()
	{
		return this.Add;
	};
	CChangesBaseContentChange.prototype.Copy = function()
	{
		var oChanges = new this.constructor(this.Class, this.Pos, this.Items, this.Add);

		oChanges.UseArray = this.UseArray;

		for (var nIndex = 0, nCount = this.PosArray.length; nIndex < nCount; ++nIndex)
			oChanges.PosArray[nIndex] = this.PosArray[nIndex];

		return oChanges;
	};
	CChangesBaseContentChange.prototype.GetItemsCount = function()
	{
		return this.Items.length;
	};
	CChangesBaseContentChange.prototype.GetItem = function(nIndex)
	{
		return this.Items[nIndex];
	};
	CChangesBaseContentChange.prototype.GetPos = function(nIndex)
	{
		if (nIndex === undefined)
			nIndex = 0;

		return this.UseArray ? this.PosArray[nIndex] : this.Pos + nIndex;
	};
	CChangesBaseContentChange.prototype.WriteToBinary = function(Writer)
	{
		// Long : Количество элементов
		// Array of
		// {
		//    Long     : позиции элементов
		//    Variable : Item
		// }
		// Long : Поле Color

		var bArray = this.UseArray;
		var nCount = this.Items.length;

		var nStartPos = Writer.GetCurPosition();
		Writer.Skip(4);
		var nRealCount = nCount;

		for (var nIndex = 0; nIndex < nCount; ++nIndex)
		{
			if (true === bArray)
			{
				if (false === this.PosArray[nIndex])
				{
					nRealCount--;
				}
				else
				{
					Writer.WriteLong(this.PosArray[nIndex]);
					this.private_WriteItem(Writer, this.Items[nIndex]);
				}
			}
			else
			{
				Writer.WriteLong(this.Pos);
				this.private_WriteItem(Writer, this.Items[nIndex]);
			}
		}

		var nEndPos = Writer.GetCurPosition();
		Writer.Seek(nStartPos);
		Writer.WriteLong(nRealCount);
		Writer.Seek(nEndPos);

		var nColor = 0;
		if (undefined !== this.Color)
		{
			nColor |= 1;
			if (true === this.Color)
				nColor |= 2;
		}
		Writer.WriteLong(nColor);
	};
	CChangesBaseContentChange.prototype.ReadFromBinary = function(Reader)
	{
		// Long : Количество элементов
		// Array of
		// {
		//    Long     : позиции элементов
		//    Variable : Item
		// }
		// Long : поле Color

		this.UseArray = true;
		this.Items    = [];
		this.PosArray = [];

		var nCount = Reader.GetLong();
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
		{
			this.PosArray[nIndex] = Reader.GetLong();
			this.Items[nIndex]    = this.private_ReadItem(Reader);
		}

		var nColor = Reader.GetLong();
		if (nColor & 1)
			this.Color = (nColor & 2) ? true : false;
	};
	CChangesBaseContentChange.prototype.private_WriteItem = function(Writer, Item)
	{
	};
	CChangesBaseContentChange.prototype.private_ReadItem = function(Reader)
	{
		return null;
	};
	CChangesBaseContentChange.prototype.ConvertToSimpleActions = function()
	{
		let arrActions = [];

		if (this.UseArray)
		{
			for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
			{
				arrActions.push({
					Item : this.Items[nIndex],
					Pos  : this.PosArray[nIndex],
					Add  : this.Add
				});
			}
		}
		else
		{
			let Pos = this.Pos;
			if (this.Add)
			{
				for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
				{
					arrActions.push({
						Item : this.Items[nIndex],
						Pos  : Pos + nIndex,
						Add  : true
					});
				}
			}
			else
			{
				for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
				{
					arrActions.push({
						Item : this.Items[nIndex],
						Pos  : Pos,
						Add  : false
					});
				}
			}
		}

		return arrActions;
	};
	CChangesBaseContentChange.prototype.ConvertToSimpleChanges = function()
	{
		let arrSimpleActions = this.ConvertToSimpleActions();
		let arrChanges       = [];

		for (let nIndex = 0, nCount = arrSimpleActions.length; nIndex < nCount; ++nIndex)
		{
			let oAction = arrSimpleActions[nIndex];
			let oChange = new this.constructor(this.Class, oAction.Pos, [oAction.Item], oAction.Add);
			arrChanges.push(oChange);
		}

		return arrChanges;
	};
	CChangesBaseContentChange.prototype.ConvertFromSimpleActions = function(arrActions)
	{
		this.UseArray = true;
		this.Pos      = 0;
		this.Items    = [];
		this.PosArray = [];

		for (var nIndex = 0, nCount = arrActions.length; nIndex < nCount; ++nIndex)
		{
			this.PosArray[nIndex] = arrActions[nIndex].Pos;
			this.Items[nIndex]    = arrActions[nIndex].Item;
		}
	};
	CChangesBaseContentChange.prototype.IsRelated = function(oChanges)
	{
		if (this.Class !== oChanges.GetClass() || this.Type !== oChanges.Type)
			return false;

		return true;
	};
	CChangesBaseContentChange.prototype.private_CreateReverseChange = function(fConstructor)
	{
		var oChange = new fConstructor();

		oChange.Class    = this.Class;
		oChange.Pos      = this.Pos;
		oChange.Items    = this.Items;
		oChange.Add      = !this.Add;
		oChange.UseArray = this.UseArray;
		oChange.PosArray = [];

		for (var nIndex = 0, nCount = this.PosArray.length; nIndex < nCount; ++nIndex)
			oChange.PosArray[nIndex] = this.PosArray[nIndex];

		return oChange;
	};
	CChangesBaseContentChange.prototype.Merge = function(oChange)
	{
		// TODO: Сюда надо бы перенести работу с ContentChanges
		return true;
	};
	CChangesBaseContentChange.prototype.GetMinPos = function()
	{
		var nPos = null;
		if (this.UseArray)
		{
			for (var nIndex = 0, nCount = this.PosArray.length; nIndex < nCount; ++nIndex)
			{
				if (null === nPos || nPos > this.PosArray[nIndex])
					nPos = this.PosArray[nIndex];
			}

			if (null === nPos)
				nPos = 0;
		}
		else
		{
			nPos = this.Pos;
		}

		return nPos;
	};
	window['AscDFH'].CChangesBaseContentChange = CChangesBaseContentChange;
	/**
	 * Базовый класс для изменения свойств.
	 * @constructor
	 * @extends {AscDFH.CChangesBase}
	 */
	function CChangesBaseProperty(Class, Old, New, Color)
	{
		CChangesBase.call(this, Class);

		this.Color = true === Color ? true : false;
		this.Old   = Old;
		this.New   = New;
	}

	CChangesBaseProperty.prototype = Object.create(CChangesBase.prototype);
	CChangesBaseProperty.prototype.constructor = CChangesBaseProperty;
	CChangesBaseProperty.prototype.Undo = function()
	{
		this.private_SetValue(this.Old);
	};
	CChangesBaseProperty.prototype.Redo = function()
	{
		this.private_SetValue(this.New);
	};
	CChangesBaseProperty.prototype.private_SetValue = function(Value)
	{
		// Эту функцию нужно переопределить в дочернем классе
	};
	CChangesBaseProperty.prototype.CreateReverseChange = function()
	{
		return new this.constructor(this.Class, this.New, this.Old, this.Color);
	};
	CChangesBaseProperty.prototype.Merge = function(oChange)
	{
		if (oChange.Class === this.Class && oChange.Type === this.Type)
		{
			this.New = oChange.New;
			return false;
		}

		return true;
	};

	window['AscDFH'].CChangesBaseProperty = CChangesBaseProperty;
	/**
	 * Базовый класс для изменения булевых свойств.
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesBaseBoolProperty(Class, Old, New, Color)
	{
		CChangesBaseProperty.call(this, Class, Old, New, Color);
	}

	CChangesBaseBoolProperty.prototype = Object.create(CChangesBaseProperty.prototype);
	CChangesBaseBoolProperty.prototype.constructor = CChangesBaseBoolProperty;
	CChangesBaseBoolProperty.prototype.WriteToBinary = function(Writer)
	{
		// Long  : Flag
		// 1-bit : Подсвечивать ли данные изменения
		// 2-bit : IsUndefined New
		// 3-bit : New value
		// 4-bit : IsUndefined Old
		// 5-bit : Old value

		var nFlags = 0;

		if (false !== this.Color)
			nFlags |= 1;

		if (undefined === this.New)
			nFlags |= 2;
		else if (true === this.New)
			nFlags |= 4;

		if (undefined === this.Old)
			nFlags |= 8;
		else if (true === this.Old)
			nFlags |= 16;

		Writer.WriteLong(nFlags);
	};
	CChangesBaseBoolProperty.prototype.ReadFromBinary = function(Reader)
	{
		// Long  : Flag
		// 1-bit : Подсвечивать ли данные изменения
		// 2-bit : IsUndefined New
		// 3-bit : New value
		// 4-bit : IsUndefined Old
		// 5-bit : Old value

		var nFlags = Reader.GetLong();

		if (nFlags & 1)
			this.Color = true;
		else
			this.Color = false;

		if (nFlags & 2)
			this.New = undefined;
		else if (nFlags & 4)
			this.New = true;
		else
			this.New = false;

		if (nFlags & 8)
			this.Old = undefined;
		else if (nFlags & 16)
			this.Old = true;
		else
			this.Old = false;
	};
	window['AscDFH'].CChangesBaseBoolProperty = CChangesBaseBoolProperty;
	/**
	 * Базовый класс для изменения числовых (double) свойств.
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesBaseDoubleProperty(Class, Old, New, Color)
	{
		CChangesBaseProperty.call(this, Class, Old, New, Color);
	}

	CChangesBaseDoubleProperty.prototype = Object.create(CChangesBaseProperty.prototype);
	CChangesBaseDoubleProperty.prototype.constructor = CChangesBaseDoubleProperty;
	CChangesBaseDoubleProperty.prototype.WriteToBinary = function(Writer)
	{
		// Long  : Flag
		// 1-bit : Подсвечивать ли данные изменения
		// 2-bit : IsUndefined New
		// 3-bit : IsUndefined Old
		// double : New
		// double : Old

		var nFlags = 0;

		if (false !== this.Color)
			nFlags |= 1;

		if (undefined === this.New)
			nFlags |= 2;

		if (undefined === this.Old)
			nFlags |= 4;

		Writer.WriteLong(nFlags);

		if (undefined !== this.New)
			Writer.WriteDouble(this.New);

		if (undefined !== this.Old)
			Writer.WriteDouble(this.Old);
	};
	CChangesBaseDoubleProperty.prototype.ReadFromBinary = function(Reader)
	{
		// Long  : Flag
		// 1-bit : Подсвечивать ли данные изменения
		// 2-bit : IsUndefined New
		// 3-bit : IsUndefined Old
		// double : New
		// double : Old


		var nFlags = Reader.GetLong();

		if (nFlags & 1)
			this.Color = true;
		else
			this.Color = false;

		if (nFlags & 2)
			this.New = undefined;
		else
			this.New = Reader.GetDouble();

		if (nFlags & 4)
			this.Old = undefined;
		else
			this.Old = Reader.GetDouble();
	};
	window['AscDFH'].CChangesBaseDoubleProperty = CChangesBaseDoubleProperty;
	/**
	 * Базовый класс для изменения объектных свойств, т.е. если свойство задано объектом.
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesBaseObjectProperty(Class, Old, New, Color)
	{
		CChangesBaseProperty.call(this, Class, Old, New, Color);
	}

	CChangesBaseObjectProperty.prototype = Object.create(CChangesBaseProperty.prototype);
	CChangesBaseObjectProperty.prototype.constructor = CChangesBaseObjectProperty;
	CChangesBaseObjectProperty.prototype.WriteToBinary = function(Writer)
	{
		// Long  : Flag
		// 1-bit : Подсвечивать ли данные изменения
		// 2-bit : IsUndefined New
		// 3-bit : IsUndefined Old
		// Variable : New
		// Variable : Old

		var nFlags = 0;

		if (false !== this.Color)
			nFlags |= 1;

		if (undefined === this.New)
			nFlags |= 2;

		if (undefined === this.Old)
			nFlags |= 4;

		Writer.WriteLong(nFlags);

		if (undefined !== this.New && this.New.Write_ToBinary)
			this.New.Write_ToBinary(Writer);

		if (undefined !== this.Old && this.Old.Write_ToBinary)
			this.Old.Write_ToBinary(Writer);
		
		this.WriteAdditional(Writer);
	};
	CChangesBaseObjectProperty.prototype.ReadFromBinary = function(Reader)
	{
		// Long  : Flag
		// 1-bit : Подсвечивать ли данные изменения
		// 2-bit : IsUndefined New
		// 3-bit : IsUndefined Old
		// Variable : New
		// Variable : Old

		var nFlags = Reader.GetLong();

		if (nFlags & 1)
			this.Color = true;
		else
			this.Color = false;

		if (nFlags & 2)
		{
			if (true === this.private_IsCreateEmptyObject())
				this.New = this.private_CreateObject();
			else
				this.New = undefined;
		}
		else
		{
			this.New = this.private_CreateObject();
			if (this.New && this.New.Read_FromBinary)
				this.New.Read_FromBinary(Reader);
		}

		if (nFlags & 4)
		{
			if (true === this.private_IsCreateEmptyObject())
				this.Old = this.private_CreateObject();
			else
				this.Old = undefined;
		}
		else
		{
			this.Old = this.private_CreateObject();
			if (this.Old && this.Old.Read_FromBinary)
				this.Old.Read_FromBinary(Reader);
		}
		
		this.ReadAdditional(Reader);
	};
	CChangesBaseObjectProperty.prototype.WriteAdditional = function(writer)
	{
	};
	CChangesBaseObjectProperty.prototype.ReadAdditional = function(reader)
	{
	};
	CChangesBaseObjectProperty.prototype.private_CreateObject = function()
	{
		// Эту функцию нужно переопределить в дочернем классе
		return null;
	};
	CChangesBaseObjectProperty.prototype.private_IsCreateEmptyObject = function()
	{
		return false;
	};
	window['AscDFH'].CChangesBaseObjectProperty = CChangesBaseObjectProperty;
	/**
	 * Базовый класс для изменения числовых (long) свойств.
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesBaseLongProperty(Class, Old, New, Color)
	{
		CChangesBaseProperty.call(this, Class, Old, New, Color);
	}

	CChangesBaseLongProperty.prototype = Object.create(CChangesBaseProperty.prototype);
	CChangesBaseLongProperty.prototype.constructor = CChangesBaseLongProperty;
	CChangesBaseLongProperty.prototype.WriteToBinary = function(Writer)
	{
		// Long  : Flag
		// 1-bit : Подсвечивать ли данные изменения
		// 2-bit : IsUndefined New
		// 3-bit : IsUndefined Old
		// long : New
		// long : Old

		var nFlags = 0;

		if (false !== this.Color)
			nFlags |= 1;

		if (undefined === this.New)
			nFlags |= 2;

		if (undefined === this.Old)
			nFlags |= 4;

		Writer.WriteLong(nFlags);

		if (undefined !== this.New)
			Writer.WriteLong(this.New);

		if (undefined !== this.Old)
			Writer.WriteLong(this.Old);
	};
	CChangesBaseLongProperty.prototype.ReadFromBinary = function(Reader)
	{
		// Long  : Flag
		// 1-bit : Подсвечивать ли данные изменения
		// 2-bit : IsUndefined New
		// 3-bit : IsUndefined Old
		// long : New
		// long : Old


		var nFlags = Reader.GetLong();

		if (nFlags & 1)
			this.Color = true;
		else
			this.Color = false;

		if (nFlags & 2)
			this.New = undefined;
		else
			this.New = Reader.GetLong();

		if (nFlags & 4)
			this.Old = undefined;
		else
			this.Old = Reader.GetLong();
	};
	window['AscDFH'].CChangesBaseLongProperty = CChangesBaseLongProperty;
	/**
	 * Базовый класс для изменения строковых свойств.
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesBaseStringProperty(Class, Old, New, Color)
	{
		CChangesBaseProperty.call(this, Class, Old, New, Color);
	}

	CChangesBaseStringProperty.prototype = Object.create(CChangesBaseProperty.prototype);
	CChangesBaseStringProperty.prototype.constructor = CChangesBaseStringProperty;
	CChangesBaseStringProperty.prototype.WriteToBinary = function(Writer)
	{
		// Long  : Flag
		// 1-bit : Подсвечивать ли данные изменения
		// 2-bit : IsUndefined New
		// 3-bit : IsUndefined Old
		// string : New
		// string : Old

		var nFlags = 0;

		if (false !== this.Color)
			nFlags |= 1;

		if (undefined === this.New)
			nFlags |= 2;

		if (undefined === this.Old)
			nFlags |= 4;

		Writer.WriteLong(nFlags);

		if (undefined !== this.New)
			Writer.WriteString2(this.New);

		if (undefined !== this.Old)
			Writer.WriteString2(this.Old);
	};
	CChangesBaseStringProperty.prototype.ReadFromBinary = function(Reader)
	{
		// Long  : Flag
		// 1-bit : Подсвечивать ли данные изменения
		// 2-bit : IsUndefined New
		// 3-bit : IsUndefined Old
		// string : New
		// string : Old


		var nFlags = Reader.GetLong();

		if (nFlags & 1)
			this.Color = true;
		else
			this.Color = false;

		if (nFlags & 2)
			this.New = undefined;
		else
			this.New = Reader.GetString2();

		if (nFlags & 4)
			this.Old = undefined;
		else
			this.Old = Reader.GetString2();
	};
	window['AscDFH'].CChangesBaseStringProperty = CChangesBaseStringProperty;
	/**
	 * Базовый класс для изменения числовых (byte) свойств.
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesBaseByteProperty(Class, Old, New, Color)
	{
		CChangesBaseProperty.call(this, Class, Old, New, Color);
	}

	CChangesBaseByteProperty.prototype = Object.create(CChangesBaseProperty.prototype);
	CChangesBaseByteProperty.prototype.constructor = CChangesBaseByteProperty;
	CChangesBaseByteProperty.prototype.WriteToBinary = function(Writer)
	{
		// Long  : Flag
		// 1-bit : Подсвечивать ли данные изменения
		// 2-bit : IsUndefined New
		// 3-bit : IsUndefined Old
		// byte  : New
		// byte  : Old

		var nFlags = 0;

		if (false !== this.Color)
			nFlags |= 1;

		if (undefined === this.New)
			nFlags |= 2;

		if (undefined === this.Old)
			nFlags |= 4;

		Writer.WriteLong(nFlags);

		if (undefined !== this.New)
			Writer.WriteByte(this.New);

		if (undefined !== this.Old)
			Writer.WriteByte(this.Old);
	};
	CChangesBaseByteProperty.prototype.ReadFromBinary = function(Reader)
	{
		// Long  : Flag
		// 1-bit : Подсвечивать ли данные изменения
		// 2-bit : IsUndefined New
		// 3-bit : IsUndefined Old
		// byte  : New
		// byte  : Old


		var nFlags = Reader.GetLong();

		if (nFlags & 1)
			this.Color = true;
		else
			this.Color = false;

		if (nFlags & 2)
			this.New = undefined;
		else
			this.New = Reader.GetByte();

		if (nFlags & 4)
			this.Old = undefined;
		else
			this.Old = Reader.GetByte();
	};
	window['AscDFH'].CChangesBaseByteProperty = CChangesBaseByteProperty;
	/**
	 * Базовый класс для изменения числовых(long) значений.
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesBaseLongValue(Class, Old, New, Color)
	{
		CChangesBaseProperty.call(this, Class, Old, New, Color);
	}

	CChangesBaseLongValue.prototype = Object.create(CChangesBaseProperty.prototype);
	CChangesBaseLongValue.prototype.constructor = CChangesBaseLongValue;
	CChangesBaseLongValue.prototype.WriteToBinary = function(Writer)
	{
		// Long  : New
		// Long  : Old

		Writer.WriteLong(this.New);
		Writer.WriteLong(this.Old);
	};
	CChangesBaseLongValue.prototype.ReadFromBinary = function(Reader)
	{
		// Long  : New
		// Long  : Old

		this.New = Reader.GetLong();
		this.Old = Reader.GetLong();
	};
	window['AscDFH'].CChangesBaseLongValue = CChangesBaseLongValue;
	/**
	 * Базовый класс для изменения булевых значений.
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesBaseBoolValue(Class, Old, New, Color)
	{
		CChangesBaseProperty.call(this, Class, Old, New, Color);
	}

	CChangesBaseBoolValue.prototype = Object.create(CChangesBaseProperty.prototype);
	CChangesBaseBoolValue.prototype.constructor = CChangesBaseBoolValue;
	CChangesBaseBoolValue.prototype.WriteToBinary = function(Writer)
	{
		// Bool  : New
		// Bool  : Old

		Writer.WriteBool(this.New);
		Writer.WriteBool(this.Old);
	};
	CChangesBaseBoolValue.prototype.ReadFromBinary = function(Reader)
	{
		// Bool  : New
		// Bool  : Old

		this.New = Reader.GetBool();
		this.Old = Reader.GetBool();
	};
	window['AscDFH'].CChangesBaseBoolValue = CChangesBaseBoolValue;
	/**
	 * Базовый класс для изменения объектных значений.
	 * @constructor
	 * @extends {AscDFH.CChangesBaseObjectProperty}
	 */
	function CChangesBaseObjectValue(Class, Old, New, Color)
	{
		CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
	}

	CChangesBaseObjectValue.prototype = Object.create(CChangesBaseObjectProperty.prototype);
	CChangesBaseObjectValue.prototype.constructor = CChangesBaseObjectValue;
	CChangesBaseObjectValue.prototype.private_IsCreateEmptyObject = function()
	{
		return true;
	};
	window['AscDFH'].CChangesBaseObjectValue = CChangesBaseObjectValue;
	/**
	 * Базовый класс для изменения строковых значений.
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesBaseStringValue(Class, Old, New, Color)
	{
		CChangesBaseProperty.call(this, Class, Old, New, Color);
	}

	CChangesBaseStringValue.prototype = Object.create(CChangesBaseProperty.prototype);
	CChangesBaseStringValue.prototype.constructor = CChangesBaseStringValue;
	CChangesBaseStringValue.prototype.WriteToBinary = function(Writer)
	{
		// String : New
		// String : Old

		Writer.WriteString2(this.New);
		Writer.WriteString2(this.Old);
	};
	CChangesBaseStringValue.prototype.ReadFromBinary = function(Reader)
	{
		// String : New
		// String : Old

		this.New = Reader.GetString2();
		this.Old = Reader.GetString2();
	};
	window['AscDFH'].CChangesBaseStringValue = CChangesBaseStringValue;
	/**
	 * Базовый класс для изменения числовых(byte) значений.
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesBaseByteValue(Class, Old, New, Color)
	{
		CChangesBaseProperty.call(this, Class, Old, New, Color);
	}

	CChangesBaseByteValue.prototype = Object.create(CChangesBaseProperty.prototype);
	CChangesBaseByteValue.prototype.constructor = CChangesBaseByteValue;
	CChangesBaseByteValue.prototype.WriteToBinary = function(Writer)
	{
		// Byte  : New
		// Byte  : Old

		Writer.WriteByte(this.New);
		Writer.WriteByte(this.Old);
	};
	CChangesBaseByteValue.prototype.ReadFromBinary = function(Reader)
	{
		// Byte  : New
		// Byte  : Old

		this.New = Reader.GetByte();
		this.Old = Reader.GetByte();
	};
	window['AscDFH'].CChangesBaseByteValue = CChangesBaseByteValue;
	/**
	 * Базовый класс для изменения числовых(double) значений.
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesBaseDoubleValue(Class, Old, New, Color)
	{
		CChangesBaseProperty.call(this, Class, Old, New, Color);
	}

	CChangesBaseDoubleValue.prototype = Object.create(CChangesBaseProperty.prototype);
	CChangesBaseDoubleValue.prototype.constructor = CChangesBaseDoubleValue;
	CChangesBaseDoubleValue.prototype.WriteToBinary = function(Writer)
	{
		// Double : New
		// Double : Old

		Writer.WriteDouble(this.New);
		Writer.WriteDouble(this.Old);
	};
	CChangesBaseDoubleValue.prototype.ReadFromBinary = function(Reader)
	{
		// Double : New
		// Double : Old

		this.New = Reader.GetDouble();
		this.Old = Reader.GetDouble();
	};
	window['AscDFH'].CChangesBaseDoubleValue = CChangesBaseDoubleValue;
	//------------------------------------------------------------------------------------------------------------------
	function DoNotRecalculate()
	{
		return false;
	}
	window['AscDFH'].InheritBaseChange = function(changeClass, baseChange, type)
	{
		window['AscDFH'].changesFactory[type] = changeClass;
		
		changeClass.prototype             = Object.create(baseChange.prototype);
		changeClass.prototype.constructor = changeClass;
		changeClass.prototype.Type        = type;
	};
	window['AscDFH'].InheritPropertyChange = function(changeClass, baseChange, type, setFunction, needRecalculate)
	{
		window['AscDFH'].changesFactory[type]   = changeClass;

		changeClass.prototype                   = Object.create(baseChange.prototype);
		changeClass.prototype.constructor       = changeClass;
		changeClass.prototype.Type              = type;
		changeClass.prototype.private_SetValue  = setFunction;

		if (undefined !== needRecalculate && !needRecalculate)
			changeClass.prototype.IsNeedRecalculate = DoNotRecalculate;
	};

})(window);

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

function CGraphicObjectsPdf(document, drawingDocument, api)
{
    AscCommonWord.CGraphicObjects.call(this, document, drawingDocument, api);
}

CGraphicObjectsPdf.prototype.constructor = CGraphicObjectsPdf;
CGraphicObjectsPdf.prototype = Object.create(AscCommonWord.CGraphicObjects.prototype);

CGraphicObjectsPdf.prototype.updateSelectionState = function(bNoCheck) {
    let text_object, drawingDocument = this.drawingDocument;
    if (this.selection.textSelection) {
        text_object = this.selection.textSelection;
    } else if (this.selection.groupSelection) {
        if (this.selection.groupSelection.selection.textSelection) {
            text_object = this.selection.groupSelection.selection.textSelection;
        } else if (this.selection.groupSelection.selection.chartSelection && this.selection.groupSelection.selection.chartSelection.selection.textSelection) {
            text_object = this.selection.groupSelection.selection.chartSelection.selection.textSelection;
        }
    } else if (this.selection.chartSelection && this.selection.chartSelection.selection.textSelection) {
        text_object = this.selection.chartSelection.selection.textSelection;
    }
    if (isRealObject(text_object)) {
        text_object.updateSelectionState(drawingDocument);
    } else if (bNoCheck !== true) {
        drawingDocument.UpdateTargetTransform(null);
        if (this.document.GetActiveObject() == null)
            drawingDocument.TargetEnd();
    }
    let oMathTrackHandler = this.document.MathTrackHandler;
    if (oMathTrackHandler) {
        this.setEquationTrack(oMathTrackHandler, this.canEdit());
    }
};

CGraphicObjectsPdf.prototype.setTableProps = function(props) {
    let by_type = this.getSelectedObjectsByTypes();
    if(by_type.tables.length === 1) {
        let sCaption = props.TableCaption;
        let sDescription = props.TableDescription;
        let sName = props.TableName;
        let dRowHeight = props.RowHeight;
        let oTable = by_type.tables[0];
        oTable.setTitle(sCaption);
        oTable.setDescription(sDescription);
        oTable.setName(sName);
        props.TableCaption = undefined;
        props.TableDescription = undefined;
        let bIgnoreHeight = false;
        if(AscFormat.isRealNumber(props.RowHeight)) {
            if(AscFormat.fApproxEqual(props.RowHeight, 0.0)) {
                props.RowHeight = 1.0;
            }
            bIgnoreHeight = false;
        }
        let target_text_object = AscFormat.getTargetTextObject(this);
        if (target_text_object === oTable) {
            oTable.graphicObject.Set_Props(props);
        }
        else {
            oTable.graphicObject.SelectAll();
            oTable.graphicObject.Set_Props(props);
            oTable.graphicObject.RemoveSelection();
        }
        props.TableCaption = sCaption;
        props.TableDescription = sDescription;
        props.RowHeight = dRowHeight;
        if(!oTable.setFrameTransform(props)) {
            this.Check_GraphicFrameRowHeight(oTable, bIgnoreHeight);
        }
    }
};
CGraphicObjectsPdf.prototype.Check_GraphicFrameRowHeight = function (grFrame, bIgnoreHeight) {
	grFrame.recalculate();
	let oTable = grFrame.graphicObject;
	oTable.private_SetTableLayoutFixedAndUpdateCellsWidth(-1);
	let content = oTable.Content, i, j;
	for (i = 0; i < content.length; ++i) {
		let row = content[i];
		if (!bIgnoreHeight && row.Pr && row.Pr.Height && row.Pr.Height.HRule === Asc.linerule_AtLeast
			&& AscFormat.isRealNumber(row.Pr.Height.Value) && row.Pr.Height.Value > 0) {
			continue;
		}
		row.Set_Height(row.Height, Asc.linerule_AtLeast);
	}
};
CGraphicObjectsPdf.prototype.setEquationTrack = AscFormat.DrawingObjectsController.prototype.setEquationTrack;

CGraphicObjectsPdf.prototype.cursorMoveLeft = function(AddToSelect/*Shift*/, Word/*Ctrl*/) {
    var target_text_object = AscFormat.getTargetTextObject(this);
    var oStartContent, oStartPara;
    if (target_text_object) {

        if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
            oStartContent = this.getTargetDocContent(false, false);
            if (oStartContent) {
                oStartPara = oStartContent.GetCurrentParagraph();
            }
            target_text_object.graphicObject.MoveCursorLeft(AddToSelect, Word);
            // this.checkRedrawOnChangeCursorPosition(oStartContent, oStartPara);
        } else {
            var content = this.getTargetDocContent(undefined, true);
            if (content) {
                oStartContent = content;
                if (oStartContent) {
                    oStartPara = oStartContent.GetCurrentParagraph();
                }
                content.MoveCursorLeft(AddToSelect, Word);
                // this.checkRedrawOnChangeCursorPosition(oStartContent, oStartPara);
            }
        }
        this.updateSelectionState();
    } else {
        if (this.selectedObjects.length === 0)
            return;

        this.moveSelectedObjectsByDir([-1, null], Word);
    }
};
CGraphicObjectsPdf.prototype.cursorMoveRight = function(AddToSelect, Word, bFromPaste) {
    var target_text_object = AscFormat.getTargetTextObject(this);
    var oStartContent, oStartPara;
    if (target_text_object) {
        if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
            oStartContent = this.getTargetDocContent(false, false);
            if (oStartContent) {
                oStartPara = oStartContent.GetCurrentParagraph();
            }
            target_text_object.graphicObject.MoveCursorRight(AddToSelect, Word, bFromPaste);
            // this.checkRedrawOnChangeCursorPosition(oStartContent, oStartPara);
        } else {
            var content = this.getTargetDocContent(undefined, true);
            if (content) {
                oStartContent = content;
                if (oStartContent) {
                    oStartPara = oStartContent.GetCurrentParagraph();
                }
                content.MoveCursorRight(AddToSelect, Word, bFromPaste);
                // this.checkRedrawOnChangeCursorPosition(oStartContent, oStartPara);
            }
        }
        this.updateSelectionState();
    } else {
        if (this.selectedObjects.length === 0)
            return;

        this.moveSelectedObjectsByDir([1, null], Word);
    }
};


CGraphicObjectsPdf.prototype.cursorMoveUp = function(AddToSelect, Word) {
    var target_text_object = AscFormat.getTargetTextObject(this);
    var oStartContent, oStartPara;
    if (target_text_object) {
        if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
            oStartContent = this.getTargetDocContent(false, false);
            if (oStartContent) {
                oStartPara = oStartContent.GetCurrentParagraph();
            }
            target_text_object.graphicObject.MoveCursorUp(AddToSelect);
            // this.checkRedrawOnChangeCursorPosition(oStartContent, oStartPara);
        } else {
            var content = this.getTargetDocContent(undefined, true);
            if (content) {
                oStartContent = content;
                if (oStartContent) {
                    oStartPara = oStartContent.GetCurrentParagraph();
                }
                content.MoveCursorUp(AddToSelect);
                // this.checkRedrawOnChangeCursorPosition(oStartContent, oStartPara);
            }
        }
        this.updateSelectionState();
    } else {
        if (this.selectedObjects.length === 0)
            return;
        this.moveSelectedObjectsByDir([null, -1], Word);
    }
};

CGraphicObjectsPdf.prototype.cursorMoveDown = function(AddToSelect, Word) {
    var target_text_object = AscFormat.getTargetTextObject(this);
    var oStartContent, oStartPara;
    if (target_text_object) {
        if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
            oStartContent = this.getTargetDocContent(false, false);
            if (oStartContent) {
                oStartPara = oStartContent.GetCurrentParagraph();
            }
            target_text_object.graphicObject.MoveCursorDown(AddToSelect);
            // this.checkRedrawOnChangeCursorPosition(oStartContent, oStartPara);
        } else {
            var content = this.getTargetDocContent(undefined, true);
            if (content) {
                oStartContent = content;
                if (oStartContent) {
                    oStartPara = oStartContent.GetCurrentParagraph();
                }
                content.MoveCursorDown(AddToSelect);
                // this.checkRedrawOnChangeCursorPosition(oStartContent, oStartPara);
            }
        }
        this.updateSelectionState();
    } else {
        if (this.selectedObjects.length === 0)
            return;
        this.moveSelectedObjectsByDir([null, 1], Word);
    }
};

CGraphicObjectsPdf.prototype.getDrawingProps = function () {
    return this.getDrawingPropsFromArray(this.getSelectedArray());
};

CGraphicObjectsPdf.prototype.addTextWithPr = function(sText, oSettings) {
    if (this.checkSelectedObjectsProtectionText()) {
        return;
    }
    this.checkSelectedObjectsAndCallback(function () {

        if (!oSettings)
            oSettings = new AscCommon.CAddTextSettings();

        let oTargetDocContent = this.getTargetDocContent(true, false);
        if (oTargetDocContent) {
            oTargetDocContent.Remove(-1, true, true, true, undefined);
            let oCurrentTextPr = oTargetDocContent.GetDirectTextPr();
            let oParagraph = oTargetDocContent.GetCurrentParagraph();
            if (oParagraph && oParagraph.GetParent()) {
                let oTempPara = new AscWord.Paragraph(oParagraph.GetParent());
                let oRun = new ParaRun(oTempPara, false);
                oRun.AddText(sText);
                oTempPara.AddToContent(0, oRun);

                oRun.SetPr(oCurrentTextPr.Copy());

                let oTextPr = oSettings.GetTextPr();
                if (oTextPr)
                    oRun.ApplyPr(oTextPr);

                let oAnchorPos = oParagraph.GetCurrentAnchorPosition();

                let oSelectedContent = new AscCommonWord.CSelectedContent();
                let oSelectedElement = new AscCommonWord.CSelectedElement();

                oSelectedElement.Element = oTempPara;
                oSelectedElement.SelectedAll = false;
                oSelectedContent.Add(oSelectedElement);
                oSelectedContent.EndCollect(oTargetDocContent);
                oSelectedContent.ForceInlineInsert();
                oSelectedContent.PlaceCursorInLastInsertedRun(!oSettings.IsMoveCursorOutside());
                oSelectedContent.Insert(oAnchorPos);

                let oTargetTextObject = AscFormat.getTargetTextObject(this);
                if (oTargetTextObject) {
                    oTargetTextObject.SetNeedRecalc(true);
                    
                    if (oTargetTextObject.checkExtentsByDocContent)
                        oTargetTextObject.checkExtentsByDocContent();
                }
            }

        }
    }, [], false, AscDFH.historydescription_Document_AddTextWithProperties);
};

window["AscPDF"].CGraphicObjectsPdf = CGraphicObjectsPdf;



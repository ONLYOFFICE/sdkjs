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

CGraphicObjectsPdf.prototype.OnMouseDown = function(e, x, y, pageIndex)
{
    //console.log("down " + this.curState.id);
    this.checkInkState();
    this.curState.onMouseDown(e, x, y, pageIndex);
    if(this.arrTrackObjects.length === 0)
    {
        this.document.GetApi().sendEvent("asc_onSelectionEnd");
    }

    if(e.ClickCount < 2)
    {
        this.updateOverlay();
        this.updateSelectionState();
    }
};

CGraphicObjectsPdf.prototype.updateSelectionState = function(bNoCheck)
{
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
        drawingDocument.TargetEnd();
        drawingDocument.SelectEnabled(false);
        drawingDocument.SelectShow();
    }
    let oMathTrackHandler = null;
    if (Asc.editor.wbModel && Asc.editor.wbModel.mathTrackHandler) {
        oMathTrackHandler = Asc.editor.wbModel.mathTrackHandler;
    } else {
        if (this.drawingObjects.cSld) {
            oMathTrackHandler = editor.WordControl.m_oLogicDocument.MathTrackHandler;
        }
    }
    if (oMathTrackHandler) {
        this.setEquationTrack(oMathTrackHandler, this.canEdit());
    }
};
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
},


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
},

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

CGraphicObjectsPdf.prototype.createTextAddState = function(object, x, y, e) {
	return new AscFormat.TextAddState2(this, object, x, y, e.Button)
};

window["AscPDF"].CGraphicObjectsPdf = CGraphicObjectsPdf;



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
    }
    else if (this.selection.groupSelection) {
        let oDrawing = this.selection.groupSelection;
        if (oDrawing.selection.textSelection) {
            text_object = oDrawing.selection.textSelection;
        }
        else if (oDrawing.selection.chartSelection && oDrawing.selection.chartSelection.selection.textSelection) {
            text_object = oDrawing.selection.chartSelection.selection.textSelection;
        }
        else if (oDrawing.IsAnnot() && oDrawing.IsFreeText() && oDrawing.IsInTextBox()) {
            text_object = oDrawing.spTree[0];
        }
    }
    else if (this.selection.chartSelection && this.selection.chartSelection.selection.textSelection) {
        text_object = this.selection.chartSelection.selection.textSelection;
    }

    if (isRealObject(text_object)) {
        text_object.updateSelectionState(drawingDocument);
    }
    else if (bNoCheck !== true) {
        drawingDocument.UpdateTargetTransform(null);
        if (this.document.GetActiveObject() == null)
            drawingDocument.TargetEnd();
    }

    let oMathTrackHandler   = this.document.MathTrackHandler;
    let oAnnotTextPrTrackHandler = this.document.AnnotTextPrTrackHandler;

    this.setEquationTrack(oMathTrackHandler, this.canEdit());
    this.setAnnotTextPrTrack(oAnnotTextPrTrackHandler, this.canEdit());
};

CGraphicObjectsPdf.prototype.setAnnotTextPrTrack = function(oAnnotTextPrTrackHandler, IsShowAnnotTrack) {
    let oDoc            = this.document;
    let oAnnot          = oDoc.mouseDownAnnot;
    let bSelection      = false;
    let bEmptySelection = true;

    let bShowTrack = oAnnot && oAnnot.IsFreeText() && oAnnot.IsInTextBox();

    oAnnotTextPrTrackHandler.SetTrackObject(IsShowAnnotTrack && bShowTrack ? oAnnot : null, 0, false === bSelection || true === bEmptySelection);
};

CGraphicObjectsPdf.prototype.canIncreaseParagraphLevel = function(bIncrease)
{
    let oDocContent = this.getTargetDocContent();
    if(oDocContent)
    {
        let oTextObject = AscFormat.getTargetTextObject(this);
        if(oTextObject && oTextObject.getObjectType() === AscDFH.historyitem_type_Shape)
        {
            if(oTextObject.isPlaceholder())
            {
                let nPhType = oTextObject.getPlaceholderType();
                if(nPhType === AscFormat.phType_title || nPhType === AscFormat.phType_ctrTitle)
                {
                    return false;
                }
            }
            return oDocContent.Can_IncreaseParagraphLevel(bIncrease);
        }
    }
    return false;
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

CGraphicObjectsPdf.prototype.fitImagesToPage = function() {
    this.checkSelectedObjectsAndCallback(function () {
        let oDoc    = this.document;
        let oApi    = this.getEditorApi();
        if (!oApi) {
            return;
        }
        
        let aSelectedObjects    = this.selection.groupSelection ? this.selection.groupSelection.selectedObjects : this.selectedObjects;
        let dWidth              = oDoc.GetPageWidthMM();
        let dHeight             = oDoc.GetPageHeightMM();

        for (let i = 0; i < aSelectedObjects.length; ++i) {
            let oDrawing = aSelectedObjects[i];
            if (oDrawing.getObjectType() === AscDFH.historyitem_type_ImageShape || oDrawing.getObjectType() === AscDFH.historyitem_type_OleObject) {
                let sImageId = oDrawing.getImageUrl();
                if(typeof sImageId === "string") {
                    sImageId = AscCommon.getFullImageSrc2(sImageId);
                    let _image = oApi.ImageLoader.map_image_index[sImageId];
                    if (_image && _image.Image) {
                        let __w     = Math.max((_image.Image.width * AscCommon.g_dKoef_pix_to_mm), 1);
                        let __h     = Math.max((_image.Image.height * AscCommon.g_dKoef_pix_to_mm), 1);
                        let fKoeff  = 1.0/Math.max(__w/dWidth, __h/dHeight);
                        let _w      = Math.max(5, __w*fKoeff);
                        let _h      = Math.max(5, __h*fKoeff);

                        AscFormat.CheckSpPrXfrm(oDrawing, true);
                        oDrawing.spPr.xfrm.setOffX((dWidth - _w)/ 2.0);
                        oDrawing.spPr.xfrm.setOffY((dHeight - _h)/ 2.0);
                        oDrawing.spPr.xfrm.setExtX(_w);
                        oDrawing.spPr.xfrm.setExtY(_h);
                    }
                }
            }
        }
    }, [], false, AscDFH.historydescription_Presentation_FitImagesToSlide)
};

CGraphicObjectsPdf.prototype.cursorMoveLeft = function(AddToSelect/*Shift*/, Word/*Ctrl*/) {
    let oViewer = Asc.editor.getDocumentRenderer();

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
        oViewer.onUpdateOverlay();
    } else {
        if (this.selectedObjects.length === 0)
            return;

        this.moveSelectedObjectsByDir([-1, null], Word);
    }
};
CGraphicObjectsPdf.prototype.cursorMoveRight = function(AddToSelect, Word, bFromPaste) {
    let oViewer = Asc.editor.getDocumentRenderer();

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
        oViewer.onUpdateOverlay();
    } else {
        if (this.selectedObjects.length === 0)
            return;

        this.moveSelectedObjectsByDir([1, null], Word);
    }
};


CGraphicObjectsPdf.prototype.cursorMoveUp = function(AddToSelect, Word) {
    let oViewer = Asc.editor.getDocumentRenderer();

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
        oViewer.onUpdateOverlay();
    } else {
        if (this.selectedObjects.length === 0)
            return;
        this.moveSelectedObjectsByDir([null, -1], Word);
    }
};

CGraphicObjectsPdf.prototype.cursorMoveDown = function(AddToSelect, Word) {
    let oViewer = Asc.editor.getDocumentRenderer();

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
        oViewer.onUpdateOverlay();
    } else {
        if (this.selectedObjects.length === 0)
            return;
        this.moveSelectedObjectsByDir([null, 1], Word);
    }
};

CGraphicObjectsPdf.prototype.getDrawingProps = function () {
    return this.getDrawingPropsFromArray(this.getSelectedArray());
};

CGraphicObjectsPdf.prototype.addTextWithPr = function (sText, oSettings) {
    if (this.checkSelectedObjectsProtectionText()) {
        return;
    }
    let oController = this;
    this.checkSelectedObjectsAndCallback(function () {

        if (!oSettings)
            oSettings = new AscCommon.CAddTextSettings();

        let fContentFunction = function() {
            let oTargetDocContent = this;
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

                let oTargetTextObject = AscFormat.getTargetTextObject(oController);
                if (oTargetTextObject) {
                    if (oTargetTextObject.group && oTargetTextObject.group.IsAnnot()) {
                        oTargetTextObject.group.SetInTextBox(true);
                        oTargetTextObject.group.SetNeedRecalc(true);
                    }
                    else {
                        oTargetTextObject.SetInTextBox(true);
                        oTargetTextObject.SetNeedRecalc(true);
                        
                        if (oTargetTextObject.checkExtentsByDocContent)
                            oTargetTextObject.checkExtentsByDocContent();
                    }
                }
            }
        };
        let fTableFunction = function () {
            let oContent = this.CurCell.Content;
            fContentFunction.call(oContent);
        };
        this.applyTextFunction(fContentFunction, fTableFunction, []);
    }, [], false, AscDFH.historydescription_Document_AddTextWithProperties);
};

CGraphicObjectsPdf.prototype.getParagraphParaPr = function () {
    let target_text_object = AscFormat.getTargetTextObject(this);
    if (target_text_object) {
        if (target_text_object.group && target_text_object.group.IsAnnot())
            return null;
        
        if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
            return target_text_object.graphicObject.GetCalculatedParaPr();
        } else {
            let content = this.getTargetDocContent();
            if (content) {
                return content.GetCalculatedParaPr();
            }
        }
    } else {
        let result, cur_pr, selected_objects, i;
        let getPropsFromArr = function (arr) {
            let cur_pr, result_pr, content;
            for (let i = 0; i < arr.length; ++i) {
                cur_pr = null;
                if (arr[i].getObjectType() === AscDFH.historyitem_type_GroupShape) {
                    cur_pr = getPropsFromArr(arr[i].arrGraphicObjects);
                } else {
                    if (arr[i].getDocContent && arr[i].getObjectType() !== AscDFH.historyitem_type_ChartSpace) {
                        content = arr[i].getDocContent();
                        if (content) {
                            content.SetApplyToAll(true);
                            cur_pr = content.GetCalculatedParaPr();
                            content.SetApplyToAll(false);
                        }
                    }
                }

                if (cur_pr) {
                    if (!result_pr)
                        result_pr = cur_pr;
                    else
                        result_pr.Compare(cur_pr);
                }
            }
            return result_pr;
        };

        if (this.selection.groupSelection && !this.selection.groupSelection.IsAnnot) {
            result = getPropsFromArr(this.selection.groupSelection.selectedObjects);
        } else {
            result = getPropsFromArr(this.selectedObjects);
        }
        return result;
    }
};

CGraphicObjectsPdf.prototype.getParagraphTextPr = function () {
    let target_text_object = AscFormat.getTargetTextObject(this);
    if (target_text_object) {
        if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
            return target_text_object.graphicObject.GetCalculatedTextPr();
        } else {
            let content = this.getTargetDocContent();
            if (content) {
                return content.GetCalculatedTextPr();
            }
        }
    } else {
        let result, cur_pr, selected_objects, i;
        let getPropsFromArr = function (arr) {
            let cur_pr, result_pr, content;
            for (let i = 0; i < arr.length; ++i) {
                cur_pr = null;
                if (arr[i].getObjectType() === AscDFH.historyitem_type_GroupShape) {
                    cur_pr = getPropsFromArr(arr[i].arrGraphicObjects);
                } else if (arr[i].getObjectType() === AscDFH.historyitem_type_ChartSpace) {
                    cur_pr = arr[i].getParagraphTextPr();
                } else {
                    if (arr[i].getDocContent) {
                        content = arr[i].getDocContent();
                        if (content) {
                            content.SetApplyToAll(true);
                            cur_pr = content.GetCalculatedTextPr();
                            content.SetApplyToAll(false);
                        }
                    }
                }

                if (cur_pr) {
                    if (!result_pr)
                        result_pr = cur_pr;
                    else
                        result_pr.Compare(cur_pr);
                }
            }
            return result_pr;
        };

        if (this.selection.groupSelection && !this.selection.groupSelection.IsAnnot) {
            result = getPropsFromArr(this.selection.groupSelection.selectedObjects);
        } else if (this.selectedObjects
            && 1 === this.selectedObjects.length
            && this.selectedObjects[0].getObjectType() === AscDFH.historyitem_type_ImageShape
            && this.selectedObjects[0].parent
            && this.selectedObjects[0].parent.Parent
            && this.selectedObjects[0].parent.Parent.GetCalculatedTextPr) {
            let oParaDrawing = this.selectedObjects[0].parent;
            let oParagraph = oParaDrawing.Parent;
            oParagraph.MoveCursorToDrawing(oParaDrawing.Get_Id(), true);
            result = oParagraph.GetCalculatedTextPr();
        } else {
            result = getPropsFromArr(this.selectedObjects);
        }
        return result;
    }
};

CGraphicObjectsPdf.prototype.getDrawingObjects = function(nPage) {
    let oViewer     = Asc.editor.getDocumentRenderer();
    let oPageInfo   = oViewer.pagesInfo.pages[nPage];
    
    return oPageInfo ? oPageInfo.drawings : [];
};

CGraphicObjectsPdf.prototype.canEditText = function () {
    let content = this.getTargetDocContent();
    if (content) {
        let oShape = content.GetParent().parent;
        return oShape.IsInTextBox();
    }
    return false;
};

CGraphicObjectsPdf.prototype.alignCenter = function(bSelected) {
    let selected_objects = this.getSelectedArray(), i, boundsObject, centerPos;

    if (selected_objects.length > 0) {
        if (bSelected && selected_objects.length > 1) {
            boundsObject = AscFormat.getAbsoluteRectBoundsArr(selected_objects);
            centerPos = boundsObject.minX + (boundsObject.maxX - boundsObject.minX) / 2;
        }
        else {
            centerPos = this.document.GetPageWidthMM(selected_objects[0].GetPage()) / 2;
        }

        this.checkSelectedObjectsForMove();
        this.swapTrackObjects();
        
        let move_state, oTrack, oDrawing, oBounds;
        if (!this.selection.groupSelection) {
            move_state = new AscFormat.MoveState(this, this.selectedObjects[0], 0, 0);
        } else {
            move_state = new AscFormat.MoveInGroupState(this, this.selection.groupSelection.selectedObjects[0], this.selection.groupSelection, 0, 0);
        }
        
        for (i = 0; i < this.arrTrackObjects.length; ++i) {
            oTrack = this.arrTrackObjects[i];
            oDrawing = oTrack.originalObject;
            oBounds = AscFormat.getAbsoluteRectBoundsObject(oDrawing);
            oTrack.track(centerPos - (oBounds.maxX - oBounds.minX) / 2 - oBounds.minX, 0, oDrawing.selectStartPage);
        }
        
        move_state.bSamePos = false;
        move_state.onMouseUp({}, 0, 0, 0);
    }
};

CGraphicObjectsPdf.prototype.alignRight = function(bSelected) {
    let selected_objects = this.getSelectedArray(), i, boundsObject, rightPos;

    if (selected_objects.length > 0) {
        if (bSelected && selected_objects.length > 1) {
            boundsObject = AscFormat.getAbsoluteRectBoundsArr(selected_objects);
            rightPos = boundsObject.maxX;
        }
        else {
            rightPos = this.document.GetPageWidthMM(selected_objects[0].GetPage());
        }

        this.checkSelectedObjectsForMove();
        this.swapTrackObjects();

        let move_state, oTrack, oDrawing, oBounds;
        if (!this.selection.groupSelection) {
            move_state = new AscFormat.MoveState(this, this.selectedObjects[0], 0, 0);
        }
        else {
            move_state = new AscFormat.MoveInGroupState(this, this.selection.groupSelection.selectedObjects[0], this.selection.groupSelection, 0, 0);
        }

        for (i = 0; i < this.arrTrackObjects.length; ++i) {
            oTrack = this.arrTrackObjects[i];
            oDrawing = oTrack.originalObject;
            oBounds = AscFormat.getAbsoluteRectBoundsObject(oDrawing);
            oTrack.track(rightPos - oBounds.maxX, 0, oDrawing.selectStartPage);
        }
        
        move_state.bSamePos = false;
        move_state.onMouseUp({}, 0, 0, 0);
    }
};

CGraphicObjectsPdf.prototype.alignMiddle = function(bSelected) {
    let selected_objects = this.getSelectedArray(), i, boundsObject, middlePos;

    if (selected_objects.length > 0) {
        if (bSelected && selected_objects.length > 1) {
            boundsObject = AscFormat.getAbsoluteRectBoundsArr(selected_objects);
            middlePos = boundsObject.minY + (boundsObject.maxY - boundsObject.minY) / 2;
        }
        else {
            middlePos = this.document.GetPageHeightMM(selected_objects[0].GetPage()) / 2;
        }
        
        this.checkSelectedObjectsForMove();
        this.swapTrackObjects();
        
        let move_state, oTrack, oDrawing, oBounds;
        if (!this.selection.groupSelection) {
            move_state = new AscFormat.MoveState(this, this.selectedObjects[0], 0, 0);
        }
        else {
            move_state = new AscFormat.MoveInGroupState(this, this.selection.groupSelection.selectedObjects[0], this.selection.groupSelection, 0, 0);
        }
        
        for (i = 0; i < this.arrTrackObjects.length; ++i) {
            oTrack = this.arrTrackObjects[i];
            oDrawing = oTrack.originalObject;
            oBounds = AscFormat.getAbsoluteRectBoundsObject(oDrawing);
            oTrack.track(0, middlePos - (oBounds.maxY - oBounds.minY) / 2 - oBounds.minY, oDrawing.selectStartPage);
        }
        
        move_state.bSamePos = false;
        move_state.onMouseUp({}, 0, 0, 0);
    }
};

CGraphicObjectsPdf.prototype.alignBottom = function(bSelected) {
    var selected_objects = this.getSelectedArray(), i, boundsObject, bottomPos;

    if (selected_objects.length > 0) {
        if (bSelected && selected_objects.length > 1) {
            boundsObject = AscFormat.getAbsoluteRectBoundsArr(selected_objects);
            bottomPos = boundsObject.maxY;
        }
        else {
            bottomPos = this.document.GetPageHeightMM(selected_objects[0].GetPage());
        }
        
        this.checkSelectedObjectsForMove();
        this.swapTrackObjects();
        
        var move_state, oTrack, oDrawing, oBounds;
        if (!this.selection.groupSelection) {
            move_state = new AscFormat.MoveState(this, this.selectedObjects[0], 0, 0);
        }
        else {
            move_state = new AscFormat.MoveInGroupState(this, this.selection.groupSelection.selectedObjects[0], this.selection.groupSelection, 0, 0);
        }
        
        for (i = 0; i < this.arrTrackObjects.length; ++i) {
            oTrack = this.arrTrackObjects[i];
            oDrawing = oTrack.originalObject;
            oBounds = AscFormat.getAbsoluteRectBoundsObject(oDrawing);
            oTrack.track(0, bottomPos - oBounds.maxY, oDrawing.selectStartPage);
        }
        
        move_state.bSamePos = false;
        move_state.onMouseUp({}, 0, 0, 0);
    }
};
CGraphicObjectsPdf.prototype.distributeHor = function(bSelected) {
    let selected_objects = this.getSelectedArray(), i, boundsObject, pos1, pos2, gap, sortObjects, lastPos;
    let oTrack, oDrawing, oBounds, oSortObject;

    if (selected_objects.length > 0) {
        let nPageW = this.document.GetPageWidthMM(selected_objects[0].GetPage());
        console.log(`PageW: ${nPageW}`);
        boundsObject = AscFormat.getAbsoluteRectBoundsArr(selected_objects);
        this.checkSelectedObjectsForMove();
        this.swapTrackObjects();

        sortObjects = [];
        for (i = 0; i < selected_objects.length; ++i) {
            oTrack = this.arrTrackObjects[i];
            oDrawing = oTrack.originalObject;
            oBounds = AscFormat.getAbsoluteRectBoundsObject(oDrawing);
            sortObjects.push({trackObject: this.arrTrackObjects[i], bounds: oBounds});
        }
        sortObjects.sort(function (obj1, obj2) {
            return (obj1.bounds.maxX + obj1.bounds.minX) / 2 - (obj2.bounds.maxX + obj2.bounds.minX) / 2
        });
        
        if (bSelected && selected_objects.length > 2) {
            pos1 = sortObjects[0].bounds.minX;
            pos2 = sortObjects[sortObjects.length - 1].bounds.maxX;
            gap = (pos2 - pos1 - boundsObject.summWidth) / (sortObjects.length - 1);
        }
        else {
            if (boundsObject.summWidth < nPageW) {
                gap = (nPageW - boundsObject.summWidth) / (sortObjects.length + 1);
                pos1 = gap;
                pos2 = nPageW - gap;
            }
            else {
                pos1 = 0;
                pos2 = nPageW;
                gap = (nPageW - boundsObject.summWidth) / (sortObjects.length - 1);
            }
        }
        
        let move_state;
        if (!this.selection.groupSelection) {
            move_state = new AscFormat.MoveState(this, this.selectedObjects[0], 0, 0);
        }
        else {
            move_state = new AscFormat.MoveInGroupState(this, this.selection.groupSelection.selectedObjects[0], this.selection.groupSelection, 0, 0);
        }

        lastPos = pos1;
        for (i = 0; i < sortObjects.length; ++i) {
            oSortObject = sortObjects[i];
            oTrack = oSortObject.trackObject;
            oDrawing = oTrack.originalObject;
            oBounds = oSortObject.bounds;
            oTrack.track(lastPos - oBounds.minX, 0, oDrawing.selectStartPage);
            lastPos += (gap + (oBounds.maxX - oBounds.minX));
        }
        
        move_state.bSamePos = false;
        move_state.onMouseUp({}, 0, 0, 0);
    }
};
CGraphicObjectsPdf.prototype.distributeVer = function(bSelected) {
    let selected_objects = this.getSelectedArray(), i, boundsObject, pos1, pos2, gap, sortObjects, lastPos;
    let oTrack, oDrawing, oBounds, oSortObject;
    
    if (selected_objects.length > 0) {
        let nPageH = this.document.GetPageHeightMM(selected_objects[0].GetPage());
        console.log(`PageH: ${nPageH}`);

        boundsObject = AscFormat.getAbsoluteRectBoundsArr(selected_objects);
        this.checkSelectedObjectsForMove();
        this.swapTrackObjects();
        
        sortObjects = [];
        for (i = 0; i < selected_objects.length; ++i) {
            oTrack = this.arrTrackObjects[i];
            oDrawing = oTrack.originalObject;
            oBounds = AscFormat.getAbsoluteRectBoundsObject(oDrawing);
            sortObjects.push({trackObject: this.arrTrackObjects[i], bounds: oBounds});
        }
        sortObjects.sort(function (obj1, obj2) {
            return (obj1.bounds.maxY + obj1.bounds.minY) / 2 - (obj2.bounds.maxY + obj2.bounds.minY) / 2
        });
        
        if (bSelected && selected_objects.length > 2) {
            pos1 = sortObjects[0].bounds.minY;
            pos2 = sortObjects[sortObjects.length - 1].bounds.maxY;
            gap = (pos2 - pos1 - boundsObject.summHeight) / (sortObjects.length - 1);
        }
        else {
            if (boundsObject.summHeight < nPageH) {
                gap = (nPageH - boundsObject.summHeight) / (sortObjects.length + 1);
                pos1 = gap;
                pos2 = nPageH - gap;
            }
            else {
                pos1 = 0;
                pos2 = nPageH;
                gap = (nPageH - boundsObject.summHeight) / (sortObjects.length - 1);
            }
        }
        
        let move_state;
        if (!this.selection.groupSelection) {
            move_state = new AscFormat.MoveState(this, this.selectedObjects[0], 0, 0);
        }
        else {
            move_state = new AscFormat.MoveInGroupState(this, this.selection.groupSelection.selectedObjects[0], this.selection.groupSelection, 0, 0);
        }
        
        lastPos = pos1;
        for (i = 0; i < sortObjects.length; ++i) {
            oSortObject = sortObjects[i];
            oTrack = oSortObject.trackObject;
            oDrawing = oTrack.originalObject;
            oBounds = oSortObject.bounds;
            oTrack.track(0, lastPos - oBounds.minY, oDrawing.selectStartPage);
            lastPos += (gap + (oBounds.maxY - oBounds.minY));
        }
        
        move_state.bSamePos = false;
        move_state.onMouseUp({}, 0, 0, 0);
    }
};
CGraphicObjectsPdf.prototype.checkChartTextSelection = function (bNoRedraw) {
    if (this.bNoCheckChartTextSelection === true)
        return false;

    let chart_selection, bRet = false;
    let nPageNum1, nPageNum2;

    if (this.selection.chartSelection) {
        chart_selection = this.selection.chartSelection;
    } else if (this.selection.groupSelection && this.selection.groupSelection.selection.chartSelection) {
        chart_selection = this.selection.groupSelection.selection.chartSelection;
    }

    if (chart_selection && (chart_selection.selection.textSelection || chart_selection.selection.title)) {
        let oTitle = chart_selection.selection.textSelection;

        if (!oTitle) {
            oTitle = chart_selection.selection.title;
            nPageNum2 = this.drawingObjects.num;
        }

        let content = oTitle.getDocContent(), bDeleteTitle = false;
        if (content) {
            if (content.Is_Empty()) {
                if (chart_selection.selection.title && chart_selection.selection.title.parent) {
                    AscCommon.History.Create_NewPoint(AscDFH.historydescription_CommonControllerCheckChartText);
                    chart_selection.selection.title.parent.setTitle(null);
                    bDeleteTitle = true;
                }
            }
        }

        if (chart_selection.recalcInfo.bRecalculatedTitle || bDeleteTitle) {
            chart_selection.recalcInfo.recalcTitle = null;
            chart_selection.handleUpdateInternalChart(false);

            chart_selection.recalculate();
            nPageNum1 = chart_selection.selectStartPage;

            chart_selection.recalcInfo.bRecalculatedTitle = false;
        }
    }
    let oTargetTextObject = AscFormat.getTargetTextObject(this);
    let nSelectStartPage = 0, bNoNeedRecalc = false;

    if (oTargetTextObject) {
        nSelectStartPage = oTargetTextObject.selectStartPage;
    }

    if (oTargetTextObject) {
        let bRedraw = false;
        let warpGeometry = oTargetTextObject.recalcInfo && oTargetTextObject.recalcInfo.warpGeometry;
        
        if (warpGeometry && warpGeometry.preset !== "textNoShape" || oTargetTextObject.worksheet) {
            if (oTargetTextObject.recalcInfo.bRecalculatedTitle) {
                oTargetTextObject.recalcInfo.recalcTitle = null;
                oTargetTextObject.recalcInfo.bRecalculatedTitle = false;
                AscFormat.ExecuteNoHistory(function () {
                    if (oTargetTextObject.bWordShape) {
                        if (!bNoNeedRecalc) {
                            oTargetTextObject.recalcInfo.oContentMetrics = oTargetTextObject.recalculateTxBoxContent();
                            oTargetTextObject.recalcInfo.recalculateTxBoxContent = false;
                            oTargetTextObject.recalcInfo.AllDrawings = [];
                            let oContent = oTargetTextObject.getDocContent();
                            if (oContent) {
                                oContent.GetAllDrawingObjects(oTargetTextObject.recalcInfo.AllDrawings);
                            }
                        }
                    } else {
                        oTargetTextObject.recalcInfo.oContentMetrics = oTargetTextObject.recalculateContent();
                        oTargetTextObject.recalcInfo.recalculateContent = false;
                    }
                }, this, []);

            }
            oTargetTextObject.AddToRedraw();
            bRedraw = true;
        }

        let oDocContent = this.getTargetDocContent();
        if (oDocContent) {
            let oParagraph = oDocContent.GetElement(0);
            let oForm;
            if (oParagraph && oParagraph.IsParagraph() && oParagraph.IsInFixedForm() && (oForm = oParagraph.GetInnerForm())) {
                oDocContent.ResetShiftView();
                oTargetTextObject.AddToRedraw();
                bRedraw = true;
            }
        }

        if (bRedraw) {
            nPageNum2 = nSelectStartPage;
        }
    }

    if (AscFormat.isRealNumber(nPageNum1)) {
        bRet = true;
    }

    if (AscFormat.isRealNumber(nPageNum2) && nPageNum2 !== nPageNum1) {
        bRet = true;
    }
    return bRet;
};

CGraphicObjectsPdf.prototype.setParagraphNumbering = function(Bullet) {
    this.applyDocContentFunction(AscWord.CDocumentContent.prototype.Set_ParagraphPresentationNumbering, [Bullet], AscWord.CTable.prototype.Set_ParagraphPresentationNumbering);
};

CGraphicObjectsPdf.prototype.selectAll = function () {
    var i;
    var target_text_object = AscFormat.getTargetTextObject(this);
    if (target_text_object) {
        if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
            target_text_object.graphicObject.SelectAll();
        } else {
            var content = this.getTargetDocContent();
            if (content) {
                content.SelectAll();
            }
        }
    }
    else {
        if (this.selection.groupSelection) {
            if (!this.selection.groupSelection.selection.chartSelection) {
                this.selection.groupSelection.resetSelection(this);
                for (i = this.selection.groupSelection.arrGraphicObjects.length - 1; i > -1; --i) {
                    this.selection.groupSelection.selectObject(this.selection.groupSelection.arrGraphicObjects[i], 0);
                }
            }
        } else if (!this.selection.chartSelection) {
            this.resetSelection();
            var drawings = this.getDrawingObjects();
            for (i = drawings.length - 1; i > -1; --i) {
                this.selectObject(drawings[i], 0);
            }
        }
    }
};

CGraphicObjectsPdf.prototype.getSelectedArray = function () {
    if (this.selection.groupSelection) {
        let oGroup = this.selection.groupSelection;
        if (oGroup.IsAnnot && oGroup.IsAnnot() && oGroup.IsFreeText() && oGroup.selection.textSelection) {
            return oGroup;
        }

        return this.selection.groupSelection.selectedObjects;
    }
    return this.selectedObjects;
};

CGraphicObjectsPdf.prototype.loadDocumentStateAfterLoadChanges = function() {};

CGraphicObjectsPdf.prototype.setEquationTrack       = AscFormat.DrawingObjectsController.prototype.setEquationTrack;
CGraphicObjectsPdf.prototype.getParagraphTextPr     = AscFormat.DrawingObjectsController.prototype.getParagraphTextPr;
CGraphicObjectsPdf.prototype.alignLeft              = AscFormat.DrawingObjectsController.prototype.alignLeft;
CGraphicObjectsPdf.prototype.alignTop               = AscFormat.DrawingObjectsController.prototype.alignTop;
CGraphicObjectsPdf.prototype.convertMathView        = AscFormat.DrawingObjectsController.prototype.convertMathView;
CGraphicObjectsPdf.prototype.setMathProps           = AscFormat.DrawingObjectsController.prototype.setMathProps;
CGraphicObjectsPdf.prototype.paraApplyCallback      = AscFormat.DrawingObjectsController.prototype.paraApplyCallback;
CGraphicObjectsPdf.prototype.setParagraphIndent     = AscFormat.DrawingObjectsController.prototype.setParagraphIndent;
CGraphicObjectsPdf.prototype.setParagraphAlign      = AscFormat.DrawingObjectsController.prototype.setParagraphAlign;
CGraphicObjectsPdf.prototype.setParagraphSpacing    = AscFormat.DrawingObjectsController.prototype.setParagraphSpacing;
CGraphicObjectsPdf.prototype.setParagraphTabs       = AscFormat.DrawingObjectsController.prototype.setParagraphTabs;
CGraphicObjectsPdf.prototype.setDefaultTabSize      = AscFormat.DrawingObjectsController.prototype.setDefaultTabSize;
CGraphicObjectsPdf.prototype.changeTextCase         = AscFormat.DrawingObjectsController.prototype.changeTextCase;

CGraphicObjectsPdf.prototype.startRecalculate = function() {};

window["AscPDF"].CGraphicObjectsPdf = CGraphicObjectsPdf;



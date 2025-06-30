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

    function CGraphicObjects(document, drawingDocument, api)
    {
        AscCommonWord.CGraphicObjects.call(this, document, drawingDocument, api);
    }

    CGraphicObjects.prototype.constructor = CGraphicObjects;
    CGraphicObjects.prototype = Object.create(AscCommonWord.CGraphicObjects.prototype);

	CGraphicObjects.prototype.createShape = function() {
		return new AscPDF.CPdfShape();
	};
    CGraphicObjects.prototype.updateSelectionState = function(bNoCheck) {
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
            let oActiveObj  = this.document.GetActiveObject();
            let oDocContent = oActiveObj ? oActiveObj.GetDocContent() : null;
            if (oDocContent == null)
                drawingDocument.TargetEnd();
        }

        let oMathTrackHandler   = this.document.MathTrackHandler;
        let oAnnotTextPrTrackHandler = this.document.AnnotTextPrTrackHandler;

        this.setEquationTrack(oMathTrackHandler, this.canEdit());
        this.setAnnotTextPrTrack(oAnnotTextPrTrackHandler, this.api.isRestrictionView());
    };

    CGraphicObjects.prototype.setAnnotTextPrTrack = function(oAnnotTextPrTrackHandler, IsShowAnnotTrack) {
        let oDoc            = this.document;
        let oAnnot          = oDoc.mouseDownAnnot;
        let bSelection      = false;
        let bEmptySelection = true;

        let bShowTrack = oAnnot && oAnnot.IsFreeText() && oAnnot.IsInTextBox();

        oAnnotTextPrTrackHandler.SetTrackObject(IsShowAnnotTrack && bShowTrack ? oAnnot : null, 0, false === bSelection || true === bEmptySelection);
    };

    CGraphicObjects.prototype.paragraphIncDecIndent = function(bIncrease) {
        this.applyDocContentFunction(AscWord.CDocumentContent.prototype.Increase_ParagraphLevel, [bIncrease], AscWord.CTable.prototype.Increase_ParagraphLevel);
    };
    CGraphicObjects.prototype.canIncreaseParagraphLevel = function(bIncrease)
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

    CGraphicObjects.prototype.getTargetTextObject = function() {
        return AscFormat.getTargetTextObject(this);
    };

    CGraphicObjects.prototype.setTableProps = function(props) {
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
            oTable.SetNeedRecalc(true);
            props.TableCaption = sCaption;
            props.TableDescription = sDescription;
            props.RowHeight = dRowHeight;
            if(!oTable.setFrameTransform(props)) {
                this.Check_GraphicFrameRowHeight(oTable, bIgnoreHeight);
            }
        }
    };
    CGraphicObjects.prototype.Check_GraphicFrameRowHeight = function (grFrame, bIgnoreHeight) {
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

    CGraphicObjects.prototype.fitImagesToPage = function() {
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
                            let __w     = Math.max((_image.Image.width * g_dKoef_pix_to_mm), 1);
                            let __h     = Math.max((_image.Image.height * g_dKoef_pix_to_mm), 1);
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

    CGraphicObjects.prototype.cursorMoveLeft = function(AddToSelect/*Shift*/, Word/*Ctrl*/) {
        let oViewer = Asc.editor.getDocumentRenderer();
        let oDoc = Asc.editor.getPDFDoc()

        var target_text_object = AscFormat.getTargetTextObject(this);
        var oStartContent, oStartPara;
        if (target_text_object) {

            if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
                oStartContent = this.getTargetDocContent(false, false);
                if (oStartContent) {
                    oStartPara = oStartContent.GetCurrentParagraph();
                }
                target_text_object.graphicObject.MoveCursorLeft(AddToSelect, Word);
            } else {
                var content = this.getTargetDocContent(undefined, true);
                if (content) {
                    oStartContent = content;
                    if (oStartContent) {
                        oStartPara = oStartContent.GetCurrentParagraph();
                    }
                    content.MoveCursorLeft(AddToSelect, Word);
                }
            }
            oViewer.onUpdateOverlay();
            oDoc.UpdateInterface();
        } else {
            if (this.selectedObjects.length === 0)
                return;

            this.moveSelectedObjectsByDir([-1, null], Word);
        }
    };
    CGraphicObjects.prototype.cursorMoveRight = function(AddToSelect, Word, bFromPaste) {
        let oViewer = Asc.editor.getDocumentRenderer();
        let oDoc = Asc.editor.getPDFDoc();
        
        var target_text_object = AscFormat.getTargetTextObject(this);
        var oStartContent, oStartPara;
        if (target_text_object) {
            if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
                oStartContent = this.getTargetDocContent(false, false);
                if (oStartContent) {
                    oStartPara = oStartContent.GetCurrentParagraph();
                }
                target_text_object.graphicObject.MoveCursorRight(AddToSelect, Word, bFromPaste);
            } else {
                var content = this.getTargetDocContent(undefined, true);
                if (content) {
                    oStartContent = content;
                    if (oStartContent) {
                        oStartPara = oStartContent.GetCurrentParagraph();
                    }
                    content.MoveCursorRight(AddToSelect, Word, bFromPaste);
                }
            }

            oViewer.onUpdateOverlay();
            oDoc.UpdateInterface();
        } else {
            if (this.selectedObjects.length === 0)
                return;

            this.moveSelectedObjectsByDir([1, null], Word);
        }
    };


    CGraphicObjects.prototype.cursorMoveUp = function(AddToSelect, Word) {
        let oViewer = Asc.editor.getDocumentRenderer();
        let oDoc = Asc.editor.getPDFDoc();

        var target_text_object = AscFormat.getTargetTextObject(this);
        var oStartContent, oStartPara;
        if (target_text_object) {
            if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
                oStartContent = this.getTargetDocContent(false, false);
                if (oStartContent) {
                    oStartPara = oStartContent.GetCurrentParagraph();
                }
                target_text_object.graphicObject.MoveCursorUp(AddToSelect);
            } else {
                var content = this.getTargetDocContent(undefined, true);
                if (content) {
                    oStartContent = content;
                    if (oStartContent) {
                        oStartPara = oStartContent.GetCurrentParagraph();
                    }
                    content.MoveCursorUp(AddToSelect);
                }
            }

            oViewer.onUpdateOverlay();
            oDoc.UpdateInterface();
        } else {
            if (this.selectedObjects.length === 0)
                return;
            this.moveSelectedObjectsByDir([null, -1], Word);
        }
    };

    CGraphicObjects.prototype.cursorMoveDown = function(AddToSelect, Word) {
        let oViewer = Asc.editor.getDocumentRenderer();
        let oDoc = Asc.editor.getPDFDoc();

        var target_text_object = AscFormat.getTargetTextObject(this);
        var oStartContent, oStartPara;
        if (target_text_object) {
            if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
                oStartContent = this.getTargetDocContent(false, false);
                if (oStartContent) {
                    oStartPara = oStartContent.GetCurrentParagraph();
                }
                target_text_object.graphicObject.MoveCursorDown(AddToSelect);
            } else {
                var content = this.getTargetDocContent(undefined, true);
                if (content) {
                    oStartContent = content;
                    if (oStartContent) {
                        oStartPara = oStartContent.GetCurrentParagraph();
                    }
                    content.MoveCursorDown(AddToSelect);
                }
            }
            
            oViewer.onUpdateOverlay();
            oDoc.UpdateInterface();
        } else {
            if (this.selectedObjects.length === 0)
                return;
            this.moveSelectedObjectsByDir([null, 1], Word);
        }
    };

    CGraphicObjects.prototype.getDrawingProps = function () {
        return this.getDrawingPropsFromArray(this.getSelectedArray());
    };

    CGraphicObjects.prototype.addTextWithPr = function (sText, oSettings) {
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

    CGraphicObjects.prototype.getParagraphParaPr = function () {
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
                    if (arr[i].IsAnnot && arr[i].IsAnnot()
                        || (arr[i].group && arr[i].group.IsAnnot && arr[i].group.IsAnnot())
                        || arr[i].IsDrawing() && arr[i].IsShape() && arr[i].GetEditField()) {
                        return null;
                    }

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

            if (this.selection.groupSelection) {
                result = getPropsFromArr(this.selection.groupSelection.selectedObjects);
            } else {
                result = getPropsFromArr(this.selectedObjects);
            }
            return result;
        }
    };

    CGraphicObjects.prototype.getParagraphTextPr = function () {
        let target_text_object = AscFormat.getTargetTextObject(this);
        if (target_text_object) {
            if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
                return target_text_object.graphicObject.GetCalculatedTextPr();
            }
            else {
                let content = this.getTargetDocContent();
                if (content) {
                    return content.GetCalculatedTextPr();
                }
            }
        }
        else {
            let result;
            let getPropsFromArr = function (arr) {
                let cur_pr, result_pr, content;
                for (let i = 0; i < arr.length; ++i) {
                    cur_pr = null;

                    if (arr[i].getObjectType() === AscDFH.historyitem_type_GroupShape) {
                        cur_pr = getPropsFromArr(arr[i].arrGraphicObjects);
                    }
                    else if (arr[i].getObjectType() === AscDFH.historyitem_type_ChartSpace) {
                        cur_pr = arr[i].getParagraphTextPr();
                    }
                    else {
                        if (arr[i].getDocContent) {
                            if (arr[i].IsDrawing() && arr[i].IsEditFieldShape()) {
                                let oField = arr[i].GetEditField();
                                let oContent = oField.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Format) ? oField.contentFormat : oField.content;
                                let oCalcedTextPr = oContent ? oContent.GetCalculatedTextPr() : null;

                                cur_pr = new AscWord.CTextPr();

                                // font family
                                cur_pr.FontFamily = {
                                    Name: oField.GetTextFontActual() || oField.GetTextFont(),
                                    Index: -1
                                };

                                // font size
                                cur_pr.SetFontSize(oCalcedTextPr ? oCalcedTextPr.GetFontSize() : oField.GetTextSize());
                                
                                // font color
                                let aColor = oField.GetTextColor();
                                let oRGB = oField.GetRGBColor(aColor);
                                let oUnifill = AscFormat.CreateUnfilFromRGB(oRGB.r, oRGB.g, oRGB.b);
                                cur_pr.Unifill = oUnifill;
                                cur_pr.Unifill.fill.color.RGBA = cur_pr.Unifill.fill.color.color.RGBA;
                            }
                            else {
                                content = arr[i].getDocContent();
                            }

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

            if (this.selection.groupSelection) {
                result = getPropsFromArr(this.selection.groupSelection.selectedObjects);
            }
            else if (this.selectedObjects
                && 1 === this.selectedObjects.length
                && this.selectedObjects[0].getObjectType() === AscDFH.historyitem_type_ImageShape
                && this.selectedObjects[0].parent
                && this.selectedObjects[0].parent.Parent
                && this.selectedObjects[0].parent.Parent.GetCalculatedTextPr) {
                let oParaDrawing = this.selectedObjects[0].parent;
                let oParagraph = oParaDrawing.Parent;
                oParagraph.MoveCursorToDrawing(oParaDrawing.Get_Id(), true);
                result = oParagraph.GetCalculatedTextPr();
            }
            else {
                result = getPropsFromArr(this.selectedObjects);
            }

            return result;
        }
    };

    CGraphicObjects.prototype.getDrawingObjects = function(nPage) {
        let oViewer     = Asc.editor.getDocumentRenderer();
        let oPageInfo   = oViewer.pagesInfo.pages[nPage];
        
        return oPageInfo ? oPageInfo.drawings : [];
    };

    CGraphicObjects.prototype.getDrawingArray = function() {
        return this.document.drawings;
    };

    CGraphicObjects.prototype.canEditText = function () {
        let content = this.getTargetDocContent();
        if (content) {
            let oShape = content.GetParent().parent;
            return oShape.IsInTextBox();
        }
        return false;
    };

    CGraphicObjects.prototype.alignCenter = function(bSelected) {
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

    CGraphicObjects.prototype.alignRight = function(bSelected) {
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

    CGraphicObjects.prototype.alignMiddle = function(bSelected) {
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

    CGraphicObjects.prototype.alignBottom = function(bSelected) {
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
    CGraphicObjects.prototype.distributeHor = function(bSelected) {
        let selected_objects = this.getSelectedArray(), i, boundsObject, pos1, pos2, gap, sortObjects, lastPos;
        let oTrack, oDrawing, oBounds, oSortObject;

        if (selected_objects.length > 0) {
            let nPageW = this.document.GetPageWidthMM(selected_objects[0].GetPage());
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
    CGraphicObjects.prototype.distributeVer = function(bSelected) {
        let selected_objects = this.getSelectedArray(), i, boundsObject, pos1, pos2, gap, sortObjects, lastPos;
        let oTrack, oDrawing, oBounds, oSortObject;
        
        if (selected_objects.length > 0) {
            let nPageH = this.document.GetPageHeightMM(selected_objects[0].GetPage());

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
    CGraphicObjects.prototype.checkChartTextSelection = function (bNoRedraw) {
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

    CGraphicObjects.prototype.setParagraphNumbering = function(Bullet) {
        this.applyDocContentFunction(AscWord.CDocumentContent.prototype.Set_ParagraphPresentationNumbering, [Bullet], AscWord.CTable.prototype.Set_ParagraphPresentationNumbering);
    };

    CGraphicObjects.prototype.selectAll = function () {
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
                let nCurPage = this.document.GetCurPage();
                this.resetSelection();
                var drawings = this.getDrawingObjects(nCurPage);
                for (i = drawings.length - 1; i > -1; --i) {
                    this.selectObject(drawings[i], nCurPage);
                }
            }
        }
    };

    CGraphicObjects.prototype.getSelectedArray = function () {
        if (this.selection.groupSelection) {
            let oGroup = this.selection.groupSelection;
            if (oGroup.IsAnnot && oGroup.IsAnnot() && oGroup.IsFreeText() && oGroup.selection.textSelection) {
                return [oGroup];
            }

            return this.selection.groupSelection.selectedObjects;
        }
        return this.selectedObjects;
    };

    CGraphicObjects.prototype.updateCursorType = function(pageIndex, x, y, e, bTextFlag)
    {
        var ret;
        this.handleEventMode = HANDLE_EVENT_MODE_CURSOR;
        ret = this.curState.onMouseDown(global_mouseEvent, x, y, pageIndex, bTextFlag);
        this.handleEventMode = HANDLE_EVENT_MODE_HANDLE;
        if(ret)
        {
            if(ret.cursorType !== "text")
            {
                let oApi = Asc.editor || editor;
                let oDoc = this.document;
                let isDrawHandles = oApi ? oApi.isShowShapeAdjustments() : true;

                let oObject = AscCommon.g_oTableId.Get_ById(ret.objectId) || oDoc.GetShapeBasedAnnotById(ret.objectId) || this.selectedObjects.find(function(obj) { return obj.GetId() == ret.objectId});
                let isViewerObj = this.document.IsViewerObject(oObject);

                if (!isDrawHandles && isViewerObj) {
                    isDrawHandles = true;
                }

                if(isDrawHandles === false)
                {
                    this.drawingDocument.SetCursorType("default");
                    return true;
                }
				let sCursorType = ret.cursorType;
				let oAPI = this.getEditorApi();
				if(oAPI.isFormatPainterOn())
				{
					if(sCursorType !== "text")
					{
						let oData = oAPI.getFormatPainterData();
						if(oData.isDrawingData())
						{
							sCursorType = AscCommon.Cursors.ShapeCopy;
						}
					}
				}
                this.drawingDocument.SetCursorType(sCursorType);
            }
            return true;
        }
        return false;
    };
    CGraphicObjects.prototype.handleTextHit = function (object, e, x, y, group, pageIndex, bWord) {
        var content, invert_transform_text, tx, ty, check_hyperlink;
    
        if (this.handleEventMode === HANDLE_EVENT_MODE_HANDLE) {
            if (object.IsDrawing() && object.IsEditFieldShape()) {
                return this.handleMoveHit(object, e, x, y, group, false, pageIndex, bWord);
            }
    
            // Обработка выбора объекта (для одиночного объекта или группы)
            if (!group) {
                if (this.selection.textSelection !== object) {
                    this.resetSelection(true);
                    this.selectObject(object, pageIndex);
                    this.selection.textSelection = object;
                } else if (this.checkTargetSelection(object, x, y, object.invertTransformText)) {
                    return true;
                }
            } else {
                if (this.selection.groupSelection !== group || group.selection.textSelection !== object) {
                    this.resetSelection(true);
                    group.selectObject(object, pageIndex);
                    this.selectObject(group, pageIndex);
                    this.selection.groupSelection = group;
                    group.selection.textSelection = object;
                } else if (this.checkTargetSelection(object, x, y, object.invertTransformText)) {
                    return true;
                }
            }
    
            // Вызываем начало выделения без проверки isSlideShow
            object.selectionSetStart(e, x, y, pageIndex);
            this.changeCurrentState(new AscFormat.TextAddState(this, object, x, y, e.Button));
            return true;
    
        } else {
            var ret = { objectId: object.Get_Id(), cursorType: "text" };
            content = object.getDocContent();
            invert_transform_text = object.invertTransformText;
            if (content && invert_transform_text) {
                tx = invert_transform_text.TransformPointX(x, y);
                ty = invert_transform_text.TransformPointY(x, y);
    
                // Отрисовка контролов для Document Editor
                if (this.document.IsDocumentEditor() && object instanceof AscFormat.CShape && object.isForm()) {
                    var oForm = object.getInnerForm();
                    if (oForm)
                        oForm.DrawContentControlsTrack(AscCommon.ContentControlTrack.Hover, tx, ty, 0, false);
                }
    
                var nPageIndex = pageIndex;
                if (this.drawingObjects.cSld && !this.noNeedUpdateCursorType && AscFormat.isRealNumber(this.drawingObjects.num)) {
                    nPageIndex = this.drawingObjects.num;
                }
                content.UpdateCursorType(tx, ty, 0);
                ret.updated = true;
            } else if (this.drawingObjects) {
                check_hyperlink = AscFormat.fCheckObjectHyperlink(object, x, y);
                if (isRealObject(check_hyperlink)) {
                    ret.hyperlink = check_hyperlink;
                }
            }
            return ret;
        }
    };
    CGraphicObjects.prototype.checkSelectedObjectsAndCallback = function(callback, args, bNoSendProps, nHistoryPointType, aAdditionaObjects, bNoCheckLock) {
    	let oDoc = this.document;
        let check_type = AscCommon.changestype_Drawing_Props;

        if(bNoCheckLock || oDoc.Document_Is_SelectionLocked(check_type, null, false, false) === false) {
            let nPointType = AscFormat.isRealNumber(nHistoryPointType) ? nHistoryPointType : AscDFH.historydescription_CommonControllerCheckSelected;
			oDoc.StartAction(nPointType);
            callback.apply(this, args);
			oDoc.Recalculate();
			oDoc.FinalizeAction(true);
        }
    };
    CGraphicObjects.prototype.selectObject = function (object, pageIndex) {
        let oDoc = this.document;
        object.select(this, pageIndex);
        if (this.selectedObjects.length == 1 && !oDoc.GetActiveObject()) {
            oDoc.SetMouseDownObject(object);
        }

        if (AscFormat.MoveAnimationDrawObject) {
            if (object instanceof AscFormat.MoveAnimationDrawObject) {
                for (let i = this.selectedObjects.length - 1; i > -1; --i) {
                    if (!this.selectedObjects[i].isMoveAnimObject()) {
                        object.selected = false;
                        this.selectedObjects.splice(i, 1);
                        return;
                    }
                }
            } else {
                for (let i = this.selectedObjects.length - 1; i > -1; --i) {
                    if (this.selectedObjects[i].isMoveAnimObject()) {
                        object.selected = false;
                        this.selectedObjects.splice(i, 1);
                        return;
                    }
                }
            }
        }
        this.lastSelectedObject = null;
        this.checkShowMediaControlOnSelect();
    }
    CGraphicObjects.prototype.drawSelect = function (pageIndex) {
        let drawingDocument = this.drawingDocument;

        if (undefined !== drawingDocument.BeginDrawTracking)
            drawingDocument.BeginDrawTracking();


        const oTrackDrawer  = drawingDocument.AutoShapesTrack;
        oTrackDrawer.SetCurrentPage(pageIndex, true);

        const oApi = this.getEditorApi();
        let isDrawHandles = oApi ? oApi.isShowShapeAdjustments() : true;
        const nSelectedCount = this.selectedObjects;
        const oFirstSelected = this.selectedObjects[0];

        if (nSelectedCount === 1
            && oFirstSelected.isForm()
            && oFirstSelected.getInnerForm()
            && oFirstSelected.getInnerForm().IsFormLocked())
            isDrawHandles = false;

        var i;
        const oTx           = this.selection.textSelection;
        const oCrop         = this.selection.cropSelection;
        const oGm           = this.selection.geometrySelection;
        const oGrp          = this.selection.groupSelection;
        const oChart        = this.selection.chartSelection;
        const oWrp          = this.selection.wrapPolygonSelection;

        if (oCrop) {
            if (this.arrTrackObjects.length === 0) {
                if (oCrop.selectStartPage === pageIndex) {
                    const cropObject = oCrop.getCropObject();
                    if (cropObject) {
                        let oldGlobalAlpha;
                        if (oTrackDrawer.Graphics) {
                            oldGlobalAlpha = oTrackDrawer.Graphics.globalAlpha;
                            oTrackDrawer.Graphics.put_GlobalAlpha(false, 1.0);
                        }
                        cropObject.draw(oTrackDrawer);
                        oTrackDrawer.CorrectOverlayBounds();

                        oCrop.draw(oTrackDrawer);
                        oTrackDrawer.CorrectOverlayBounds();

                        if (oTrackDrawer.Graphics) {
                            oTrackDrawer.Graphics.put_GlobalAlpha(true, oldGlobalAlpha);
                        }

                        oTrackDrawer.CheckCanvasTransform();

                        drawingDocument.DrawTrack(
                            AscFormat.TYPE_TRACK.SHAPE,
                            cropObject.getTransformMatrix(),
                            0,
                            0,
                            cropObject.extX,
                            cropObject.extY,
                            false,
                            false,
                            undefined,
                            isDrawHandles && cropObject.canEdit()
                        );
                        drawingDocument.DrawTrack(
                            AscFormat.TYPE_TRACK.CROP,
                            oCrop.getTransformMatrix(),
                            0,
                            0,
                            oCrop.extX,
                            oCrop.extY,
                            false,
                            false,
                            undefined,
                            isDrawHandles && oCrop.canEdit()
                        );
                    }
                }
            }
        } else if (oGm) {
            oGm.drawSelect(pageIndex, drawingDocument);
        } else if (oTx) {
            if (oTx.selectStartPage === pageIndex) {
                if (!oTx.isForm()) {
                    drawingDocument.DrawTrack(
                        AscFormat.TYPE_TRACK.TEXT,
                        oTx.getTransformMatrix(),
                        0,
                        0,
                        oTx.extX,
                        oTx.extY,
                        AscFormat.CheckObjectLine(oTx),
                        oTx.canRotate(),
                        undefined,
                        isDrawHandles && oTx.canEdit()
                    );
                    oTx.drawAdjustments(drawingDocument);
                }
            }
        } else if (oGrp) {
            if (oGrp.selectStartPage === pageIndex) {
                isDrawHandles = !oGrp.IsAnnot();
                let bDrawGroupTrack = !oGrp.IsAnnot() || oGrp.IsAnnot() && !oGrp.IsInTextBox(); 

                bDrawGroupTrack && drawingDocument.DrawTrack(
                    AscFormat.TYPE_TRACK.GROUP_PASSIVE,
                    oGrp.getTransformMatrix(),
                    0,
                    0,
                    oGrp.extX,
                    oGrp.extY,
                    false,
                    oGrp.canRotate(),
                    undefined,
                    isDrawHandles && oGrp.canEdit()
                );
                const oGrpTx = oGrp.selection.textSelection;
                const oGrpChart = oGrp.selection.chartSelection;
                const aGrpSelected = oGrp.selectedObjects;
                if (oGrpTx && !oGrp.IsAnnot) {
                    drawingDocument.DrawTrack(
                        AscFormat.TYPE_TRACK.TEXT,
                        oGrpTx.transform,
                        0,
                        0,
                        oGrpTx.extX,
                        oGrpTx.extY,
                        AscFormat.CheckObjectLine(oGrpTx),
                        oGrpTx.canRotate(),
                        undefined,
                        isDrawHandles && this.selection.groupSelection.canEdit()
                    );
                } else if (oGrpChart) {
                    oGrpChart.drawSelect(drawingDocument, pageIndex);
                } else {
                    for (i = 0; i < aGrpSelected.length; ++i) {
                        let oDrawing = aGrpSelected[i];
                        if (oDrawing instanceof AscFormat.CConnectionShape) {
                            continue;
                        }

                        drawingDocument.DrawTrack(
                            AscFormat.TYPE_TRACK.SHAPE,
                            oDrawing.transform,
                            0,
                            0,
                            oDrawing.extX,
                            oDrawing.extY,
                            AscFormat.CheckObjectLine(oDrawing),
                            false,
                            undefined,
                            (isDrawHandles || this.document.IsViewerObject(oDrawing)) && oGrp.canEdit());
                    }
                }
                if (aGrpSelected.length === 1) {
                    aGrpSelected[0].drawAdjustments(drawingDocument);
                }
            }
        } else if (oChart) {
            oChart.drawSelect(drawingDocument, pageIndex);
        } else if (oWrp) {
            if (oWrp.selectStartPage === pageIndex) {
                if(oTrackDrawer.DrawEditWrapPointsPolygon) {
                    oTrackDrawer.DrawEditWrapPointsPolygon(oWrp.parent.wrappingPolygon.calculatedPoints, new AscCommon.CMatrix());
                }
            }
        } else {
            for (i = 0; i < this.selectedObjects.length; ++i) {
                let oDrawing = this.selectedObjects[i];
                if (oDrawing.selectStartPage === pageIndex) {

                    if (oDrawing.IsForm()) {
                        oDrawing = oDrawing.editShape;
                    }

                    let nType = oDrawing.IsAnnot() && oDrawing.IsStamp() ? AscFormat.TYPE_TRACK.ANNOT_STAMP : AscFormat.TYPE_TRACK.SHAPE;

                    if (oDrawing.IsAnnot() && (oDrawing.IsTextMarkup() || oDrawing.IsComment())) {
                        oDrawing.DrawSelected(drawingDocument.Overlay);
                    }
                    else {
                        if (oDrawing.IsAnnot() && oDrawing.IsFreeText()) {
                            drawingDocument.DrawTrack(
                                nType,
                                oDrawing.getTransformMatrix(),
                                0,
                                0,
                                oDrawing.extX,
                                oDrawing.extY,
                                AscFormat.CheckObjectLine(oDrawing),
                                oDrawing.canRotate(),
                                undefined,
                                true
                            );
                        }
                        else {
                            drawingDocument.DrawTrack(
                                nType,
                                oDrawing.getTransformMatrix(),
                                0,
                                0,
                                oDrawing.extX,
                                oDrawing.extY,
                                AscFormat.CheckObjectLine(oDrawing),
                                oDrawing.canRotate(),
                                undefined,
                                (isDrawHandles || this.document.IsViewerObject(oDrawing)) && oDrawing.canEdit()
                            );
                        }
                    }
                }
            }
            if (this.selectedObjects.length === 1 && this.selectedObjects[0].drawAdjustments && this.selectedObjects[0].selectStartPage === pageIndex) {
                this.selectedObjects[0].drawAdjustments(drawingDocument);
            }
        }
        if (this.document) {
            if (!oGm) {
                if (nSelectedCount === 1 && oFirstSelected.parent && !oFirstSelected.parent.Is_Inline()) {
                    let anchor_pos;
                    let oFirstTrack = this.arrTrackObjects[0];
                    let page_index;
                    if (this.arrTrackObjects.length === 1 &&
                        !(oFirstTrack instanceof TrackPointWrapPointWrapPolygon || oFirstTrack instanceof TrackNewPointWrapPolygon)) {
                        page_index = AscFormat.isRealNumber(oFirstTrack.pageIndex) ? this.arrTrackObjects[0].pageIndex : (AscFormat.isRealNumber(oFirstTrack.selectStartPage) ? oFirstTrack.selectStartPage : 0);
                        if (page_index === pageIndex) {
                            var bounds = oFirstTrack.getBounds();
                            var nearest_pos = this.document.Get_NearestPos(page_index, bounds.min_x, bounds.min_y, true, this.selectedObjects[0].parent);
                            nearest_pos.Page = page_index;
                            oTrackDrawer.drawFlowAnchor(nearest_pos.X, nearest_pos.Y);
                        }
                    } else {
                        page_index = oFirstSelected.selectStartPage;
                        if (page_index === pageIndex) {
                            var paragraph = oFirstSelected.parent.Get_ParentParagraph();
                            anchor_pos = paragraph.Get_AnchorPos(oFirstSelected.parent);
                            if (anchor_pos) {
                                oTrackDrawer.drawFlowAnchor(anchor_pos.X, anchor_pos.Y);
                            }
                        }
                    }
                }
            }
        }
        if (this.selectionRect) {
            drawingDocument.DrawTrackSelectShapes(this.selectionRect.x, this.selectionRect.y, this.selectionRect.w, this.selectionRect.h);
        }


        if (this.connector) {
            this.connector.drawConnectors(oTrackDrawer);
            this.connector = null;
        }


        if (undefined !== drawingDocument.EndDrawTracking)
            drawingDocument.EndDrawTracking();


    };
    CGraphicObjects.prototype.deselectObject = function (object) {
        let oDoc = this.document;
        for (let i = 0; i < this.selectedObjects.length; ++i) {
            if (this.selectedObjects[i] === object) {
                object.selected = false;
                this.selectedObjects.splice(i, 1);
                if(this.selectedObjects.length === 0) {
                    this.lastSelectedObject = object;
                }
                this.checkShowMediaControlOnSelect();
                if (!object.IsFreeText || !object.IsFreeText()) {
                    let oAcitveObj = oDoc.GetActiveObject();
                    if (oAcitveObj) {
                        let oPrev = this.selectedObjects[0];

                        // edit field shape
                        if (object.IsDrawing() && object.IsEditFieldShape() && oAcitveObj == object.GetEditField()) {
                            oDoc.SetMouseDownObject(oPrev && oPrev.GetEditField() || null, false);
                            oDoc.activeForm = oPrev && oPrev.GetEditField() || null;
                        }
                        // other objects
                        else if (oAcitveObj == object) {
                            oDoc.SetMouseDownObject(oPrev || null, false);
                        }
                    }
                }
                return;
            }
        }
    };
    CGraphicObjects.prototype.hyperlinkCanAdd = function (bCheckInHyperlink) {
        var content = this.getTargetDocContent();
        if (content) {
            return content.CanAddHyperlink(bCheckInHyperlink);
        }
        return false;
    };

    CGraphicObjects.prototype.setParagraphIndent = function(Indent) {
        if(AscCommon.isRealObject(Indent) && AscFormat.isRealNumber(Indent.Left) && Indent.Left < 0) {
            Indent.Left = 0;
        }
        this.applyDocContentFunction(AscWord.CDocumentContent.prototype.SetParagraphIndent, [Indent], AscWord.CTable.prototype.SetParagraphIndent);
    };
    CGraphicObjects.prototype.endTrackNewShape = function() {
        this.curState.bStart = this.curState.bStart !== false;
        let aTracks = this.arrTrackObjects;
        let bNewShape = false;
        let bRet = false;
        if (aTracks.length > 0) {
            let nT;
            for (nT = 0; nT < aTracks.length; ++nT) {
                let oTrack = aTracks[nT];
                if (!oTrack.getShape) {
                    break;
                }
            }
            if (nT === aTracks.length) {
                bNewShape = true;
            }
            if (bNewShape) {
                bRet = AscFormat.StartAddNewShape.prototype.onMouseUp.call(this.curState, {
                    ClickCount: 1,
                    X: 0,
                    Y: 0
                }, 0, 0, 0);
            }
            else {
                this.curState.onMouseUp({
                    ClickCount: 1,
                    X: 0,
                    Y: 0
                }, 0, 0, 0);
                bRet = true;
            }
        }
        else {
            bRet = AscFormat.StartAddNewShape.prototype.onMouseUp.call(this.curState, {
                ClickCount: 1,
                X: 0,
                Y: 0
            }, 0, 0, 0);
        }
    
        const oApi = this.getEditorApi();
        if (oApi.isInkDrawerOn()) {
            oApi.stopInkDrawer();
        }

        return bRet;
    };

    CGraphicObjects.prototype.saveDocumentState = function(){};
	
	CGraphicObjects.prototype.getAllRasterImagesOnPage = function(pageIndex) {
		if (!this.api)
			return [];
		
		let viewer = this.api.getDocumentRenderer();
		if (!viewer || !viewer.pagesInfo.pages[pageIndex])
			return [];
		
		let page = viewer.pagesInfo.pages[pageIndex];
		
		let result = []
		
		page.fields.forEach(function(field){
			field.getAllRasterImages(result);
		});
		page.drawings.forEach(function(drawing){
			drawing.getAllRasterImages(result);
		});
		
		return result;
	};
    CGraphicObjects.prototype.applyTextFunction = function(docContentFunction, tableFunction, args) {
        let bApplyTextPr = false;
        if(docContentFunction === CDocumentContent.prototype.AddToParagraph && args[0].Type === para_TextPr) {
            bApplyTextPr = true;
        }
        if (this.selection.textSelection) {
            if(!this.selection.textSelection.checkEditTextStyle(docContentFunction, args, true)) {
                this.selection.textSelection.checkDrawingPartWithHistory();
                this.selection.textSelection.applyTextFunction(docContentFunction, tableFunction, args);
            }
        } else if (this.selection.groupSelection) {
            var oOldDoc = this.selection.groupSelection.document;
            this.selection.groupSelection.document = this.document;
            this.selection.groupSelection.applyTextFunction(docContentFunction, tableFunction, args);
            this.selection.groupSelection.document = oOldDoc;
        } else if (this.selection.chartSelection) {
            this.selection.chartSelection.applyTextFunction(docContentFunction, tableFunction, args);
            if (this.document) {
                this.document.Recalculate();
            }
        } else {
            if (bApplyTextPr || docContentFunction === CDocumentContent.prototype.PasteFormatting) {
                this.applyDocContentFunction(docContentFunction, args, tableFunction, docContentFunction);
            } else if (this.selectedObjects.length === 1 && ((this.selectedObjects[0].getObjectType() === AscDFH.historyitem_type_Shape && this.selectedObjects[0].canEditText()) || this.selectedObjects[0].getObjectType() === AscDFH.historyitem_type_GraphicFrame)) {
                this.selection.textSelection = this.selectedObjects[0];
                if (this.selectedObjects[0].getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
                    this.selectedObjects[0].graphicObject.MoveCursorToStartPos(false);
                    this.selectedObjects[0].applyTextFunction(docContentFunction, tableFunction, args);
                } else {
                    this.selectedObjects[0].checkDrawingPartWithHistory && this.selectedObjects[0].checkDrawingPartWithHistory();
                    var oDocContent = this.selectedObjects[0].getDocContent();
                    if (oDocContent) {
                        oDocContent.MoveCursorToEndPos(false);
                    }
                    this.selectedObjects[0].applyTextFunction(docContentFunction, tableFunction, args);
                    this.selection.textSelection.select(this, this.selection.textSelection.selectStartPage);
                }
            } else if (this.parent && this.parent.GoToText) {
                this.parent.GoToText();
                this.resetSelection();
                if (this.document && (docpostype_DrawingObjects !== this.document.GetDocPosType() || isRealObject(AscFormat.getTargetTextObject(this.document.DrawingObjects))) && CDocumentContent.prototype.AddNewParagraph === docContentFunction) {
                    this.document.AddNewParagraph(args[0]);
                }
            } else if (this.selectedObjects.length > 0 && this.selectedObjects[0].parent && this.selectedObjects[0].parent.GoToText) {
                this.selectedObjects[0].parent.GoToText();
                this.resetSelection();
                if (this.document && (docpostype_DrawingObjects !== this.document.GetDocPosType() || isRealObject(AscFormat.getTargetTextObject(this))) && CDocumentContent.prototype.AddNewParagraph === docContentFunction) {
                    this.document.AddNewParagraph(args[0]);
                }
            }
        }
    };
    CGraphicObjects.prototype.applyDocContentFunction = function (f, args, tableFunction, fOrigDocContentMethod) {
        let oThis = this;
        let isIncreaseDecreaseFunction = f === CDocumentContent.prototype.IncreaseDecreaseFontSize;

        function applyToArrayDrawings(arr) {
            var ret = false, ret2;
            for (var i = 0; i < arr.length; ++i) {
                let drawing = arr[i];
                const newShape = arr[i].checkDrawingPartWithHistory && arr[i].checkDrawingPartWithHistory();
                if (newShape) {
                    drawing = newShape;
                }
                if (drawing.getObjectType() === AscDFH.historyitem_type_GroupShape || drawing.getObjectType() === AscDFH.historyitem_type_SmartArt) {
                    ret2 = applyToArrayDrawings(drawing.arrGraphicObjects);
                    if (ret2) {
                        ret = true;
                    }
                } else if (drawing.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
                    drawing.graphicObject.SetApplyToAll(true);
                    tableFunction.apply(drawing.graphicObject, args);
                    drawing.graphicObject.SetApplyToAll(false);
                    ret = true;
                } else if (drawing.getObjectType() === AscDFH.historyitem_type_ChartSpace) {
                    if (args[0].Type === para_TextPr) {
                        var oChartSpace = drawing;

                        var fCallback = function (oElement) {
                            AscFormat.CheckObjectTextPr(oElement, args[0].Value, oThis.getDrawingDocument());
                        };
                        oChartSpace.applyLabelsFunction(fCallback, args[0].Value);
                    }
                    if (f === CDocumentContent.prototype.IncreaseDecreaseFontSize) {
                        drawing.paragraphIncDecFontSize(args[0]);
                    }
                } else if (drawing.getDocContent) {
                    if(!drawing.checkEditTextStyle(fOrigDocContentMethod || f, args, false)) {
                        if (drawing.IsDrawing() && drawing.IsEditFieldShape()) {
                            let oField = drawing.GetEditField();
                            if (f == CDocumentContent.prototype.AddToParagraph) {
                                let oTextPr = args[0].Value;

                                if (oTextPr.FontFamily && oTextPr.FontFamily.Name) {
                                    oField.SetTextFont(oTextPr.FontFamily.Name);
                                    oField.SetTextFontActual(oTextPr.FontFamily.Name);
                                }
                                if (oTextPr.GetFontSize()) {
                                    oField.SetTextSize(oTextPr.GetFontSize());
                                }
                                if (oTextPr.Unifill) {
                                    let oRGBA = oTextPr.Unifill.fill.color.color.RGBA;
                                    oField.SetTextColor([oRGBA.R / 255, oRGBA.G / 255, oRGBA.B / 255]);
                                }
                            }
                            else if (f == CDocumentContent.prototype.SetParagraphAlign){
                                oField.SetAlign(AscPDF.getPdfTypeAlignByInternal(args[0]));
                            }
                            else if (f == CDocumentContent.prototype.IncreaseDecreaseFontSize) {
                                let oContent = oField.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Format) ? oField.contentFormat : oField.content;
                                let oCalcedTextPr = oContent.GetCalculatedTextPr();
                                
                                let nCurFontSize = oCalcedTextPr.GetFontSize();
                                let oTextPr = new AscWord.CTextPr();
                                oTextPr.SetFontSize(nCurFontSize);
                                oField.SetTextSize(oTextPr.GetIncDecFontSize(args[0]));
                            }
                        }
                        else {
                            let content = drawing.getDocContent();
                            if (content) {
                                content.SetApplyToAll(true);
                                f.apply(content, args);
                                content.SetApplyToAll(false);
                                ret = true;
                                if (isIncreaseDecreaseFunction && drawing.isObjectInSmartArt()) {
                                    drawing.setCustT(true);
                                }
                            } else {
                                if (drawing.getObjectType() === AscDFH.historyitem_type_Shape) {
                                    if (drawing.canEditText()) {
                                        if (drawing.bWordShape) {
                                            drawing.createTextBoxContent();
                                        } else {
                                            drawing.createTextBody();
                                        }
                                        content = drawing.getDocContent();
                                        if (content) {
                                            content.SetApplyToAll(true);
                                            f.apply(content, args);
                                            content.SetApplyToAll(false);
                                            ret = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                if (drawing.checkExtentsByDocContent) {
                    drawing.checkExtentsByDocContent();
                }
            }
            return ret;
        }

        function applyToChartSelection(chart) {
            var content;
            if (chart.selection.textSelection) {
                chart.selection.textSelection.checkDocContent();
                content = chart.selection.textSelection.getDocContent();
                if (content) {
                    f.apply(content, args);
                }
            } else if (chart.selection.title) {
                content = chart.selection.title.getDocContent();
                if (content) {
                    content.SetApplyToAll(true);
                    f.apply(content, args);
                    content.SetApplyToAll(false);
                }
            }
        }

        if (this.selection.textSelection) {
            if (this.selection.textSelection.getObjectType() !== AscDFH.historyitem_type_GraphicFrame) {
                if(!this.selection.textSelection.checkEditTextStyle(fOrigDocContentMethod || f, args, true)) {
                    f.apply(this.selection.textSelection.getDocContent(), args);
                }
                this.selection.textSelection.checkExtentsByDocContent();
            } else {
                tableFunction.apply(this.selection.textSelection.graphicObject, args);
            }
        } else if (this.selection.groupSelection) {
            if (this.selection.groupSelection.selection.textSelection) {
                if (this.selection.groupSelection.selection.textSelection.getObjectType() !== AscDFH.historyitem_type_GraphicFrame) {
                    var frame = this.selection.groupSelection.selection.textSelection;
                    frame.checkDrawingPartWithHistory();
                    frame = this.selection.groupSelection.selection.textSelection;
                    f.apply(frame.getDocContent(), args);
                    if (frame.isObjectInSmartArt() && isIncreaseDecreaseFunction) {
                        frame.setCustT(true);
                    }
                    frame.checkExtentsByDocContent();
                } else {
                    tableFunction.apply(this.selection.groupSelection.selection.textSelection.graphicObject, args);
                }
            } else if (this.selection.groupSelection.selection.chartSelection) {
                if (isIncreaseDecreaseFunction) {
                    this.selection.groupSelection.selection.chartSelection.paragraphIncDecFontSize(args[0]);
                } else {
                    applyToChartSelection(this.selection.groupSelection.selection.chartSelection);
                }

            } else
                applyToArrayDrawings(this.selection.groupSelection.selectedObjects);
        } else if (this.selection.chartSelection) {
            if (isIncreaseDecreaseFunction) {
                this.selection.chartSelection.paragraphIncDecFontSize(args[0]);
            } else {
                applyToChartSelection(this.selection.chartSelection);
            }
        } else {
            applyToArrayDrawings(this.selectedObjects);
        }
        if (this.document) {
            this.document.Recalculate();
        }
    }

    // import
    CGraphicObjects.prototype.setEquationTrack          = AscFormat.DrawingObjectsController.prototype.setEquationTrack;
    CGraphicObjects.prototype.alignLeft                 = AscFormat.DrawingObjectsController.prototype.alignLeft;
    CGraphicObjects.prototype.alignTop                  = AscFormat.DrawingObjectsController.prototype.alignTop;
    CGraphicObjects.prototype.convertMathView           = AscFormat.DrawingObjectsController.prototype.convertMathView;
    CGraphicObjects.prototype.setMathProps              = AscFormat.DrawingObjectsController.prototype.setMathProps;
    CGraphicObjects.prototype.paraApplyCallback         = AscFormat.DrawingObjectsController.prototype.paraApplyCallback;
    CGraphicObjects.prototype.setParagraphAlign         = AscFormat.DrawingObjectsController.prototype.setParagraphAlign;
    CGraphicObjects.prototype.setParagraphSpacing       = AscFormat.DrawingObjectsController.prototype.setParagraphSpacing;
    CGraphicObjects.prototype.setParagraphTabs          = AscFormat.DrawingObjectsController.prototype.setParagraphTabs;
    CGraphicObjects.prototype.setDefaultTabSize         = AscFormat.DrawingObjectsController.prototype.setDefaultTabSize;
    CGraphicObjects.prototype.changeTextCase            = AscFormat.DrawingObjectsController.prototype.changeTextCase;
    CGraphicObjects.prototype.handleDblClickEmptyShape  = AscFormat.DrawingObjectsController.prototype.handleDblClickEmptyShape;
    CGraphicObjects.prototype.getDrawingsPasteShift     = AscFormat.DrawingObjectsController.prototype.getDrawingsPasteShift;
    CGraphicObjects.prototype.removeCallback            = AscFormat.DrawingObjectsController.prototype.removeCallback;
    CGraphicObjects.prototype.getAllSingularDrawings    = AscFormat.DrawingObjectsController.prototype.getAllSingularDrawings;
    CGraphicObjects.prototype.setParagraphBidi          = AscFormat.DrawingObjectsController.prototype.setParagraphBidi;
    CGraphicObjects.prototype.loadDocumentStateAfterLoadChanges = AscFormat.DrawingObjectsController.prototype.loadDocumentStateAfterLoadChanges;

    CGraphicObjects.prototype.startRecalculate = function() {};

    // export 
    window["AscPDF"].CGraphicObjects = CGraphicObjects;

})(window);

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

(function(){

    /**
	 * Class representing a pdf text shape.
	 * @constructor
    */
    function CPdfDrawingPrototype() {
        this._isFromScan = false;   // флаг, что был прочитан из скана страницы 

        this._doc           = undefined;
        this._needRecalc    = true;
        this._redactIds     = [];

        this.redactsArrayChanges = new AscCommon.CContentChanges(); // related redacts list
    }
    
    CPdfDrawingPrototype.prototype.constructor = CPdfDrawingPrototype;
    
    CPdfDrawingPrototype.prototype.IsAnnot = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.IsForm = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.IsShape = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.IsImage = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.IsChart = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.IsDrawing = function() {
        return true;
    };
    CPdfDrawingPrototype.prototype.IsPdfObject = function() {
        return true;
    };
    CPdfDrawingPrototype.prototype.IsEditFieldShape = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.OnContentChange = function() {
        return this.SetNeedRecalc(true);
    };
    CPdfDrawingPrototype.prototype.IsPdfDrawing = function() {
        return true;
    };
    CPdfDrawingPrototype.prototype.IsSmartArt = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.IsGraphicFrame = function() {
        return false;
    };
	CPdfDrawingPrototype.prototype.IsUseInDocument = function() {
		let oDoc = this.GetDocument();
		if (!oDoc) {
			return false;
		}

        let oGroup = this.getMainGroup();
		if (oGroup && oGroup != this && oGroup.IsUseInDocument)
			return oGroup.IsUseInDocument();
		
        let oPage = this.GetParentPage();
        if (oPage && oPage.drawings.includes(this) && oPage.GetIndex() !== -1) {
            return true;
        }

        if (this.IsShape() && oDoc.IsEditFieldsMode()) {
            let oEditFiled = this.GetEditField();
            if (oEditFiled) {
                return oEditFiled.IsUseInDocument();
            }
        }
        
        return false;
    };
    CPdfDrawingPrototype.prototype.OnBlur = function() {
        AscCommon.History.ForbidUnionPoint();
    };
    CPdfDrawingPrototype.prototype.SetParentPage = function(oParent) {
        this.parent = oParent;
    };
    CPdfDrawingPrototype.prototype.setParent2 = function() {};
    CPdfDrawingPrototype.prototype.GetParentPage = function() {
        return this.parent;
    };
    CPdfDrawingPrototype.prototype.GetTopParentObj = function() {
        let oParent = this;
        while (!oParent.AddToRedraw) {
            if (oParent.GetParent) {
                oParent = oParent.GetParent();
            }
            else if (oParent.parent) {
                oParent = oParent.parent;
            }
            else if (oParent.Parent) {
                oParent = oParent.Parent;
            }
            else if (oParent.GetTable) {
                oParent = oParent.GetTable();
            }
            else {
                return null;
            }
        }

        return oParent;
    }
    
    CPdfDrawingPrototype.prototype.GetSelectionQuads = function() {
        let oDoc        = this.GetDocument();
        let oViewer     = oDoc.Viewer;
        let oContent    = this.GetDocContent();
        let aInfo       = [];
        let nPage       = this.GetPage();

        if (!oContent || !oContent.IsSelectionUse()) {
            return aInfo;
        }

        let oTextTr = oContent.Get_ParentTextTransform();

        let nStart = oContent.Selection.StartPos;
        let nEnd   = oContent.Selection.EndPos;
        if (nStart > nEnd) {
            let temp = nStart;
            nStart = nEnd;
            nEnd = temp;
        }

        let oInfo = {
            page: nPage,
            quads: []
        }

        for (let i = nStart; i <= nEnd; i++) {
            let oPara = oContent.GetElement(i);

            let nStartInPara = oPara.Selection.StartPos;
            let nEndInPara   = oPara.Selection.EndPos;
            if (nStartInPara > nEndInPara) {
                let temp = nStartInPara;
                nStartInPara = nEndInPara;
                nEndInPara = temp;
            }

            let nStartLine = oPara.Pages[0].StartLine;
			let nEndLine   = oPara.Pages[0].EndLine;

            if (nStartInPara > oPara.Lines[nEndLine].Get_EndPos() || nEndInPara < oPara.Lines[nStartLine].Get_StartPos()) {
				return aInfo;
            }
			else {
				nStartInPara = Math.max(nStartInPara, oPara.Lines[nStartLine].Get_StartPos());
				nEndInPara   = Math.min(nEndInPara, ( nEndLine !== oPara.Lines.length - 1 ? oPara.Lines[nEndLine].Get_EndPos() : oPara.Content.length - 1 ));
			}

            let oDrawSelectionState = new AscWord.ParagraphDrawSelectionState(oPara);
			oDrawSelectionState.resetPage(0);

            for (let iLine = nStartLine; iLine <= nEndLine; ++iLine) {
                let oLine = oPara.Lines[iLine];
                oDrawSelectionState.resetLine(iLine);
                
                for (let iRange = 0, rangeCount = oLine.Ranges.length; iRange < rangeCount; ++iRange) {
                    let oRange = oLine.Ranges[iRange];
                    
                    let nRangeStart = oRange.StartPos;
                    let nRangeEnd   = oRange.EndPos;
                    
                    // Если пересечение пустое с селектом, тогда пропускаем данный отрезок
                    if (nStartInPara > nRangeEnd || nEndInPara < nRangeStart)
                        continue;
                    
                    oDrawSelectionState.beginRange(iRange);
                    
                    for (let pos = nRangeStart; pos <= nRangeEnd; ++pos) {
                        oPara.Content[pos].drawSelectionInRange(iLine, iRange, oDrawSelectionState);
                    }
                    
                    oDrawSelectionState.endRange();
                    
                    let aAnchored = oDrawSelectionState.getAnchoredObjects();
                    for (let index = 0; index < aAnchored.length; ++index) {
                        aAnchored[index].Draw_Selection();
                    }
                }
            }

            let aSelectionRanges = oDrawSelectionState.getSelectionRanges();
            for (let iSel = 0; iSel < aSelectionRanges.length; ++iSel) {
                let x = aSelectionRanges[iSel].x;
                let w = aSelectionRanges[iSel].w;
                let y = aSelectionRanges[iSel].y;
                let h = aSelectionRanges[iSel].h;
                
                if (oPara.CalculatedFrame) {
                    let Frame_X_min = oPara.CalculatedFrame.L2;
                    let Frame_Y_min = oPara.CalculatedFrame.T2;
                    let Frame_X_max = oPara.CalculatedFrame.L2 + oPara.CalculatedFrame.W2;
                    let Frame_Y_max = oPara.CalculatedFrame.T2 + oPara.CalculatedFrame.H2;
                    
                    x = Math.min(Math.max(Frame_X_min, x), Frame_X_max);
                    y = Math.min(Math.max(Frame_Y_min, y), Frame_Y_max);
                    w = Math.min(w, Frame_X_max - x);
                    h = Math.min(h, Frame_Y_max - y);
                }
                
                let isTextMatrixUse = ((null != oTextTr) && !global_MatrixTransformer.IsIdentity(oTextTr));
                if (isTextMatrixUse) {
                    let oPt1 = oTextTr.TransformPoint(x, y);            // левый верхний
                    let oPt2 = oTextTr.TransformPoint(x + w, y);        // правый верхний
                    let oPt3 = oTextTr.TransformPoint(x + w, y + h);    // правый нижний
                    let oPt4 = oTextTr.TransformPoint(x, y + h);        // левый нижний

                    let nKoeff = oViewer.getDrawingPageScale(nPage) * g_dKoef_pix_to_mm;

                    oInfo.quads.push([oPt1.x / nKoeff, oPt1.y / nKoeff, oPt2.x / nKoeff, oPt2.y / nKoeff, oPt4.x / nKoeff, oPt4.y / nKoeff, oPt3.x / nKoeff, oPt3.y / nKoeff]);
                }
            }
        }

        if (oInfo.quads.length != 0) {
            aInfo.push(oInfo);
        }

        return aInfo;
    };
    CPdfDrawingPrototype.prototype.GetRect = function() {
        let bounds = this.bounds;
        
        return [
            (bounds.l) * g_dKoef_mm_to_pt,
            (bounds.t) * g_dKoef_mm_to_pt,
            (bounds.r) * g_dKoef_mm_to_pt,
            (bounds.b) * g_dKoef_mm_to_pt,
        ];
    };
    CPdfDrawingPrototype.prototype.AddRedactId = function(id) {
        AscCommon.History.Add(new CChangesPDFDrawingRedacts(this, this._redactIds.length, [id], true));

        this._redactIds.push(id);
    };
    CPdfDrawingPrototype.prototype.GetRedactIds = function() {
        return this._redactIds;
    };
    CPdfDrawingPrototype.prototype.RemoveRedactId = function(nPos) {
        let ids = this._redactIds.splice(nPos, 1);
        AscCommon.History.Add(new CChangesPDFDrawingRedacts(this, nPos, ids, false));

        return ids[0];
    };
    CPdfDrawingPrototype.prototype.ClearRedacts = function() {
        let nCount = this._redactIds.length;
        for (let i = 0; i < nCount; i++) {
            this.RemoveRedactId(0);
        }
    };
    
    CPdfDrawingPrototype.prototype.SetFromScan = function(bFromScan) {
        this._isFromScan = bFromScan;
    };
    CPdfDrawingPrototype.prototype.IsFromScan = function() {
        return this._isFromScan;
    };
    CPdfDrawingPrototype.prototype.SetDocument = function(oDoc) {
        if (this._doc == oDoc) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFObjectSetDocument(this, this._doc, oDoc));
        this._doc = oDoc;
    };
    CPdfDrawingPrototype.prototype.OnContentChange = function() {
        let oGroup = this.getMainGroup();
        if (oGroup) {
            oGroup.SetNeedRecalc && oGroup.SetNeedRecalc(true);
        }
        else {
            this.SetNeedRecalc(true);
        }
    };
    CPdfDrawingPrototype.prototype.OnTextPrChange = function() {
        let oGroup = this.getMainGroup();
        if (oGroup) {
            oGroup.SetNeedRecalc && oGroup.SetNeedRecalc(true);
        }
        else {
            this.SetNeedRecalc(true);
        }
    };
    CPdfDrawingPrototype.prototype.GetDocument = function() {
        let oGroup = this.getMainGroup();
        if (oGroup && oGroup != this)
            return oGroup.getLogicDocument();

        return this._doc;
    };
    CPdfDrawingPrototype.prototype.SetPage = function(nPage) {
        if (this.GetPage() == nPage) {
            return;
        }
        
        let oDoc        = this.GetDocument();
        let oNewPage    = oDoc.GetPageInfo(nPage);

        if (oNewPage) {
            oDoc.RemoveDrawing(this.GetId(), true);
            oDoc.AddDrawing(this, nPage);
            this.selectStartPage = nPage;
        }
    };
    CPdfDrawingPrototype.prototype.GetPage = function() {
        let oGroup = this.getMainGroup();
        if (oGroup && oGroup != this) {
            return oGroup.GetPage();
        }
        
        let oParentPage = this.GetParentPage();
        if (!oParentPage || !(oParentPage instanceof AscPDF.CPageInfo)) {
            return -1;
        }

        return oParentPage.GetIndex();
    };
    
    CPdfDrawingPrototype.prototype.AddToRedraw = function() {
        let oViewer = Asc.editor.getDocumentRenderer();
        let nPage   = this.GetPage();
        
        if (false == this.IsUseInDocument()) {
            return;
        }
        
        let oGroup = this.getMainGroup();
        if (oGroup && oGroup != this && oGroup.IsAnnot()) {
            oGroup.AddToRedraw();
            return;
        }

        function setRedrawPageOnRepaint() {
            if (oViewer.pagesInfo.pages[nPage]) {
                oViewer.pagesInfo.pages[nPage].needRedrawDrawings = true;
                oViewer.thumbnails && oViewer.thumbnails._repaintPage(nPage);
            }
        }

        oViewer.paint(setRedrawPageOnRepaint);
    };
    
    CPdfDrawingPrototype.prototype.GetRot = function() {
        return this.rot;
    };
    CPdfDrawingPrototype.prototype.Recalculate = function() {
    };
    CPdfDrawingPrototype.prototype.IsNeedRecalc = function() {
       return this._needRecalc;
    };
    CPdfDrawingPrototype.prototype.SetNeedRecalc = function(bRecalc, bSkipAddToRedraw) {
        if (bRecalc == false) {
            this._needRecalc = false;
        }
        else {
            let oGroup = this.getMainGroup();
            if (oGroup && oGroup != this) {
                if (!oGroup.IsPdfObject || !oGroup.IsPdfObject()) {
                    return;
                }

                oGroup.SetNeedRecalc(bRecalc, bSkipAddToRedraw);
                return;
            }

            let oDoc = Asc.editor.getPDFDoc();
            if (false == this.IsEditFieldShape()) {
                oDoc.SetNeedUpdateSearch(true);
            }

            oDoc.SetNeedUpdateTarget(true);
            this._needRecalc = true;
            if (bSkipAddToRedraw != true)
                this.AddToRedraw();

            if (oGroup && oGroup != this) {
                oGroup.SetNeedRecalc(true);
            }
        }
    };
    CPdfDrawingPrototype.prototype.addToRecalculate = function() {
        this.SetNeedRecalc(true);
    };
    CPdfDrawingPrototype.prototype.GetAllFonts = function(fontMap) {
        fontMap = fontMap || {};
        return fontMap;
    };
    CPdfDrawingPrototype.prototype.CheckTextOnOpen = function() {};
    CPdfDrawingPrototype.prototype.Draw = function(oGraphicsWord) {
        function normRect(r) {
			const x1 = Math.min(r[0], r[2]);
			const y1 = Math.min(r[1], r[3]);
			const x2 = Math.max(r[0], r[2]);
			const y2 = Math.max(r[1], r[3]);
			return [x1, y1, x2, y2];
		}

        this.Recalculate();
        if (this.IsEditFieldShape()) {
            this.draw(oGraphicsWord);
            return;
        }

        let _t = this;
        let oDoc = Asc.editor.getPDFDoc();
        let nPage = this.GetPage();
        let aRedactIds = this.GetRedactIds();
        
        let aRedactAnnots = [];
        oDoc.annots.forEach(function(annot) {
            if (!annot.IsRedact() || !annot.GetRedactId() || annot.GetPage() != _t.GetPage()) {
                return;
            }

            let sRedactId = annot.GetRedactId();
            if (aRedactIds.includes(sRedactId)) {
                aRedactAnnots.push(annot);
            }
        });

        const aRectsList = [];
        aRedactAnnots.forEach(function(annot) {
            const aQuadsParts = annot.GetQuads();

            aQuadsParts.forEach(function(quads) {
                const r = normRect([quads[0], quads[1], quads[6], quads[7]]);
                aRectsList.push(r);
            });
        });

        let unredactedPolygon = null;
        let zoom = AscCommon.AscBrowser.convertToRetinaValue(oGraphicsWord.m_lWidthPix) / (oDoc.GetPageWidthMM(nPage) * g_dKoef_mm_to_pix);
        let mm2px = AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_mm_to_pix * zoom;

        if (aRectsList.length) {
            let nPageW = oDoc.GetPageWidthMM(nPage) * mm2px;
            let nPageH = oDoc.GetPageHeightMM(nPage) * mm2px;

            unredactedPolygon = {
                inverted : false,
                regions : [
                    [
                        [0, 0],
                        [nPageW, 0],
                        [nPageW, nPageH],
                        [0, nPageH]
                    ]
                ]
            };
        }

        for (let i = 0; i < aRectsList.length; i++) {
            const clip = aRectsList[i];
            let x = clip[0] * g_dKoef_pt_to_mm * mm2px;
            let y = clip[1] * g_dKoef_pt_to_mm * mm2px;
            let w = (clip[2] - clip[0]) * g_dKoef_pt_to_mm * mm2px;
            let h = (clip[3] - clip[1]) * g_dKoef_pt_to_mm * mm2px;

            let redactRect = {
                inverted : false,
                regions : [
                    [
                        [x, y],
                        [x + w, y],
                        [x + w, y + h],
                        [x, y + h]
                    ]
                ]
            };

            unredactedPolygon = AscGeometry.PolyBool.difference(unredactedPolygon, redactRect);
        }

        if (unredactedPolygon) {
            let ctx = oGraphicsWord.m_oContext;
            ctx.save();
            ctx.beginPath();

            for (let i = 0, countPolygons = unredactedPolygon.regions.length; i < countPolygons; i++)
            {
                let region = unredactedPolygon.regions[i];
                let countPoints = region.length;

                if (2 > countPoints)
                    continue;

                ctx.moveTo(region[0][0], region[0][1]);

                for (let j = 1, countPoints = region.length; j < countPoints; j++)
                {
                    ctx.lineTo(region[j][0], region[j][1]);
                }

                ctx.closePath();
            }

            ctx.clip("evenodd");
            ctx.beginPath();
            ctx.save();
        }

        this.draw(oGraphicsWord);

        if (unredactedPolygon) {
            oGraphicsWord.restore();
            oGraphicsWord.restore();
        }
    };
    CPdfDrawingPrototype.prototype.onMouseDown = function(x, y, e) {};
    CPdfDrawingPrototype.prototype.onMouseUp = function(x, y, e) {};
    CPdfDrawingPrototype.prototype.GetDocContent = function() {
        return null;
    };
    CPdfDrawingPrototype.prototype.IsInTextBox = function() {
        let oDoc = editor.getPDFDoc();
        let oController = oDoc.GetController();

        if (oDoc.GetActiveObject() == this && oController.getTargetTextObject()) {
            return !!this.GetDocContent();
        }

        return false;
    };
	CPdfDrawingPrototype.prototype.Remove = function(direction, isWord) {
		let doc = this.GetDocument();
		let content = this.GetDocContent();
		
		if (!doc || !content)
			return;
		
		content.Remove(direction, true, false, false, isWord);
		this.SetNeedRecalc(true);
		content.RecalculateCurPos();
	};
	CPdfDrawingPrototype.prototype.EnterText = function(codePoints) {
		let doc = this.GetDocument();
        let controller = doc.GetController();
		let content = this.GetDocContent();
		if (!doc || !content)
			return false;
		
        for (let nIdx = 0; nIdx < codePoints.length; ++nIdx) {
            let nCode = codePoints[nIdx];
            let oItem = AscCommon.IsSpace(nCode) ? new AscWord.CRunSpace(nCode) : new AscWord.CRunText(nCode);
            controller.paragraphAdd(oItem, false);
        }

		return true;
	};
	CPdfDrawingPrototype.prototype.canBeginCompositeInput = function() {
		return true;
	};
	CPdfDrawingPrototype.prototype.beforeCompositeInput = function() {
		let docContent = this.GetDocContent();
		if (docContent.IsSelectionUse()) {
			docContent.Remove(1, true, false, true);
			docContent.RemoveSelection();
		}
	};
    
    //////////////////////////////////////////////////////////////////////////////
    ///// Overrides
    /////////////////////////////////////////////////////////////////////////////
    
    CPdfDrawingPrototype.prototype.GetAbsolutePage = function(nCurPage) {
        return this.GetPage() != undefined ? this.GetPage() : nCurPage;
    };
    CPdfDrawingPrototype.prototype.getLogicDocument = function() {
        return this.GetDocument();
    };
    CPdfDrawingPrototype.prototype.IsThisElementCurrent = function() {
        return this.selected;
    };
    CPdfDrawingPrototype.prototype.getDrawingDocument = function() {
        return Asc.editor.getPDFDoc().GetDrawingDocument();
    };
    CPdfDrawingPrototype.prototype.handleUpdateRot = function() {
        this.recalcTransformText && this.recalcTransformText();
        this.SetNeedRecalc(true);
    };

    /////////////////////////////
    /// saving
    ////////////////////////////

    CPdfDrawingPrototype.prototype.WriteToBinary = function(memory) {
        if (Asc.editor.getShapeSerializeType() === "xml") {
            this.toXml(memory, '');
        } else {
            // Write base64 binaryData
            let writer = new AscCommon.CBinaryFileWriter();
            writer.WriteSpTreeElem(this);
            memory.WriteXmlString(writer.GetBase64Memory());
        }
    };
    CPdfDrawingPrototype.prototype.WriteRedactIds = function(oWriter) {
        let aRedactIds = this.GetRedactIds();

        oWriter.StartRecord(0xFF);
        oWriter.WriteULong(aRedactIds.length);
        aRedactIds.forEach(function(id) {
            oWriter.WriteString2(id);
        });
        oWriter.EndRecord();
    };

    window["AscPDF"].CPdfDrawingPrototype = CPdfDrawingPrototype;
})();


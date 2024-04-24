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
    function CPdfGraphicFrame()
    {
        AscFormat.CGraphicFrame.call(this);
    }
    
    CPdfGraphicFrame.prototype.constructor = CPdfGraphicFrame;
    CPdfGraphicFrame.prototype = Object.create(AscFormat.CGraphicFrame.prototype);
    Object.assign(CPdfGraphicFrame.prototype, AscPDF.PdfDrawingPrototype.prototype);

    CPdfGraphicFrame.prototype.IsGraphicFrame = function() {
        return true;
    };

    CPdfGraphicFrame.prototype.GetDocContent = function() {
        return this.getDocContent();
    };
    CPdfGraphicFrame.prototype.EnterText = function(aChars) {
        let oDoc        = this.GetDocument();
        let oContent    = this.GetDocContent();
        let oParagraph  = oContent.GetCurrentParagraph();

        oDoc.CreateNewHistoryPoint({objects: [this]});

        // удаляем текст в селекте
        if (oContent.IsSelectionUse())
            oContent.Remove(-1);

        for (let index = 0; index < aChars.length; ++index) {
            let oRun = AscPDF.codePointToRunElement(aChars[index]);
            if (oRun)
                oParagraph.AddToParagraph(oRun, true);
        }

        this.SetNeedRecalc(true);
        return true;
    };
    /**
     * Removes char in current position by direction.
     * @memberof CTextField
     * @typeofeditors ["PDF"]
     */
    CPdfGraphicFrame.prototype.Remove = function(nDirection, isCtrlKey) {
        let oDoc = this.GetDocument();
        oDoc.CreateNewHistoryPoint({objects: [this]});

        let oContent = this.GetDocContent();
        oContent.Remove(nDirection, true, false, false, isCtrlKey);
        this.SetNeedRecalc(true);

        if (AscCommon.History.Is_LastPointEmpty()) {
            AscCommon.History.Remove_LastPoint();
        }
        else {
            this.SetNeedRecalc(true);
        }
    };
    
    CPdfGraphicFrame.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        this.recalcInfo.recalculateTransform = true;
        this.recalcInfo.recalculateSizes = true;
        
        this.recalculate();
        this.updateTransformMatrix();
        this.SetNeedRecalc(false);
    };
    CPdfGraphicFrame.prototype.SetInTextBox = function(bIn) {
        this.isInTextBox = bIn;
    };
    CPdfGraphicFrame.prototype.SetNeedRecalc = function(bRecalc, bSkipAddToRedraw) {
        if (bRecalc == false) {
            this._needRecalc = false;
        }
        else {
            this.GetDocument().SetNeedUpdateTarget(true);
            this._needRecalc = true;
            this.recalcInfo.recalculateTable = true;
           
            if (bSkipAddToRedraw != true)
                this.AddToRedraw();
        }
    };
    CPdfGraphicFrame.prototype.onMouseDown = function(x, y, e) {
        let oDoc                = this.GetDocument();
        let oDrawingObjects     = oDoc.Viewer.DrawingObjects;
        let oDrDoc              = oDoc.GetDrawingDocument();
        this.selectStartPage    = this.GetPage();

        let oPos    = oDrDoc.ConvertCoordsFromCursor2(x, y);
        let X       = oPos.X;
        let Y       = oPos.Y;

        if (this.hitInBoundingRect(X, Y) || this.hitToHandles(X, Y) != -1 || this.hitInPath(X, Y)) {
            this.SetInTextBox(false);
        }
        else {
            this.SetInTextBox(true);
        }

        oDrawingObjects.OnMouseDown(e, X, Y, this.selectStartPage);
    };
    
    CPdfGraphicFrame.prototype.onMouseUp = function(x, y, e) {
        let oViewer = Asc.editor.getDocumentRenderer();
        let oDoc    = this.GetDocument();
        let oDrDoc  = oDoc.GetDrawingDocument();

        this.selectStartPage    = this.GetPage();
        let oContent            = this.GetDocContent();

        this.graphicObject.Selection.Start = false;

        if (oDrDoc.m_sLockedCursorType.indexOf("resize") != -1) {
            this.SetNeedRecalc(true);
        }

        if (!oContent) {
            return;
        }

        if (global_mouseEvent.ClickCount == 2) {
            oContent.SelectAll();
            if (oContent.IsSelectionEmpty() == false)
                oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd();
            else
                oContent.RemoveSelection();
        }
                
        if (oContent.IsSelectionEmpty())
            oContent.RemoveSelection();
    };
    
    //////////////////////////////////////////////////////////////////////////////
    ///// Overrides
    /////////////////////////////////////////////////////////////////////////////
    
    CPdfGraphicFrame.prototype.Get_Styles = function (level) {
		if (AscFormat.isRealNumber(level)) {
			if (!this.compiledStyles[level]) {
				AscFormat.CShape.prototype.recalculateTextStyles.call(this, level);
			}
			return this.compiledStyles[level];
		}
        else {
			return Asc.editor.getPDFDoc().globalTableStyles;
		}
	};
    CPdfGraphicFrame.prototype.selectionSetStart = function (e, x, y) {
		if (AscCommon.g_mouse_button_right === e.Button) {
			this.rightButtonFlag = true;
			return;
		}
		if (isRealObject(this.graphicObject)) {
			var tx, ty;
			tx = this.invertTransform.TransformPointX(x, y);
			ty = this.invertTransform.TransformPointY(x, y);
			if (AscCommon.g_mouse_event_type_down === e.Type) {
				if (this.graphicObject.IsTableBorder(tx, ty, 0)) {
					if (editor.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props) !== false) {
						return;
					} else {
					}
				}
			}

			if (!(/*content.IsTextSelectionUse() && */e.ShiftKey)) {
				if (editor.WordControl.m_oLogicDocument.CurPosition) {
					editor.WordControl.m_oLogicDocument.CurPosition.X = tx;
					editor.WordControl.m_oLogicDocument.CurPosition.Y = ty;
				}
				this.graphicObject.Selection_SetStart(tx, ty, this.GetPage(), e);
			} else {
				if (!this.graphicObject.IsSelectionUse()) {
					this.graphicObject.StartSelectionFromCurPos();
				}
				this.graphicObject.Selection_SetEnd(tx, ty, this.GetPage(), e);
			}
			this.graphicObject.RecalculateCurPos();
		}
	};
    CPdfGraphicFrame.prototype.updateSelectionState = function () {
        let oDoc    = this.GetDocument();
        let oDrDoc  = oDoc.GetDrawingDocument();

		if (isRealObject(this.graphicObject)) {
			let graphicObject = this.graphicObject;
			if (true === graphicObject.IsSelectionUse() && !graphicObject.IsSelectionEmpty()) {
				oDrDoc.UpdateTargetTransform(this.transform);
				oDrDoc.TargetEnd();
				oDrDoc.Overlay && graphicObject.DrawSelectionOnPage(0);
			} else {
				graphicObject.RecalculateCurPos();
				oDrDoc.UpdateTargetTransform(this.transform);
				oDrDoc.TargetShow();
			}
		} else {
			oDrDoc.UpdateTargetTransform(null);
			oDrDoc.TargetEnd();
		}
	};
    CPdfGraphicFrame.prototype.Set_CurrentElement = function () {
        let oDoc        = this.GetDocument();
        let oController = oDoc.GetController();

		oController.resetSelection(true);
        if (this.group) {
            var main_group = this.group.getMainGroup();
            oController.selectObject(main_group, 0);
            main_group.selectObject(this, 0);
            main_group.selection.textSelection = this;
        }
        else {
            oController.selectObject(this, 0);
            oController.selection.textSelection = this;
        }
	};
    CPdfGraphicFrame.prototype.Get_PageFields = function (nPage) {
        return this.Get_PageLimits(nPage);
    };

    window["AscPDF"].CPdfGraphicFrame = CPdfGraphicFrame;
})();


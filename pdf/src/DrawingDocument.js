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
    function CDrawingDocument()
    {
        AscCommon.CDrawingDocument.call(this);
        var oThis = this;

        this.DrawGuiCanvasTextProps = function(props) {
            var bIsChange = false;
            if (null == this.GuiLastTextProps)
            {
                bIsChange = true;
        
                this.GuiLastTextProps = new Asc.asc_CParagraphProperty();
        
                this.GuiLastTextProps.Subscript   = props.Subscript;
                this.GuiLastTextProps.Superscript = props.Superscript;
                this.GuiLastTextProps.SmallCaps   = props.SmallCaps;
                this.GuiLastTextProps.AllCaps     = props.AllCaps;
                this.GuiLastTextProps.Strikeout   = props.Strikeout;
                this.GuiLastTextProps.DStrikeout  = props.DStrikeout;
        
                this.GuiLastTextProps.TextSpacing = props.TextSpacing;
                this.GuiLastTextProps.Position    = props.Position;
            }
            else
            {
                if (this.GuiLastTextProps.Subscript != props.Subscript)
                {
                    this.GuiLastTextProps.Subscript = props.Subscript;
                    bIsChange                       = true;
                }
                if (this.GuiLastTextProps.Superscript != props.Superscript)
                {
                    this.GuiLastTextProps.Superscript = props.Superscript;
                    bIsChange                         = true;
                }
                if (this.GuiLastTextProps.SmallCaps != props.SmallCaps)
                {
                    this.GuiLastTextProps.SmallCaps = props.SmallCaps;
                    bIsChange                       = true;
                }
                if (this.GuiLastTextProps.AllCaps != props.AllCaps)
                {
                    this.GuiLastTextProps.AllCaps = props.AllCaps;
                    bIsChange                     = true;
                }
                if (this.GuiLastTextProps.Strikeout != props.Strikeout)
                {
                    this.GuiLastTextProps.Strikeout = props.Strikeout;
                    bIsChange                       = true;
                }
                if (this.GuiLastTextProps.DStrikeout != props.DStrikeout)
                {
                    this.GuiLastTextProps.DStrikeout = props.DStrikeout;
                    bIsChange                        = true;
                }
                if (this.GuiLastTextProps.TextSpacing != props.TextSpacing)
                {
                    this.GuiLastTextProps.TextSpacing = props.TextSpacing;
                    bIsChange                         = true;
                }
                if (this.GuiLastTextProps.Position != props.Position)
                {
                    this.GuiLastTextProps.Position = props.Position;
                    bIsChange                      = true;
                }
            }
        
            if (undefined !== this.GuiLastTextProps.Position && isNaN(this.GuiLastTextProps.Position))
                this.GuiLastTextProps.Position = undefined;
            if (undefined !== this.GuiLastTextProps.TextSpacing && isNaN(this.GuiLastTextProps.TextSpacing))
                this.GuiLastTextProps.TextSpacing = undefined;
        
            if (!bIsChange)
                return;
        
        
            AscFormat.ExecuteNoHistory(function(){
        
                var _oldTurn      = editor.isViewMode;
                editor.isViewMode = true;
        
                var docContent = new CDocumentContent(this.m_oWordControl.m_oLogicDocument, this.m_oWordControl.m_oDrawingDocument, 0, 0, 1000, 1000, false, false, true);
                var par        = docContent.Content[0];
        
                par.MoveCursorToStartPos();
        
                par.Set_Pr(new CParaPr());
                var _textPr        = new CTextPr();
                _textPr.FontFamily = {Name : "Arial", Index : -1};
                _textPr.FontSize = (AscCommon.AscBrowser.convertToRetinaValue(11 << 1, true) >> 0) * 0.5;
                _textPr.RFonts.SetAll("Arial");
        
                _textPr.Strikeout = this.GuiLastTextProps.Strikeout;
        
                if (true === this.GuiLastTextProps.Subscript)
                    _textPr.VertAlign = AscCommon.vertalign_SubScript;
                else if (true === this.GuiLastTextProps.Superscript)
                    _textPr.VertAlign = AscCommon.vertalign_SuperScript;
                else
                    _textPr.VertAlign = AscCommon.vertalign_Baseline;
        
                _textPr.DStrikeout = this.GuiLastTextProps.DStrikeout;
                _textPr.Caps       = this.GuiLastTextProps.AllCaps;
                _textPr.SmallCaps  = this.GuiLastTextProps.SmallCaps;
        
                _textPr.Spacing  = this.GuiLastTextProps.TextSpacing;
                _textPr.Position = this.GuiLastTextProps.Position;
        
                var parRun = new ParaRun(par);
                parRun.Set_Pr(_textPr);
                parRun.AddText("Hello World");
                par.AddToContent(0, parRun);
        
                docContent.Recalculate_Page(0, true);
        
                var baseLineOffset = par.Lines[0].Y;
                var _bounds        = par.Get_PageBounds(0);
        
                var ctx  = this.GuiCanvasTextProps.getContext('2d');
                var _wPx = this.GuiCanvasTextProps.width;
                var _hPx = this.GuiCanvasTextProps.height;
        
                var _wMm = _wPx * g_dKoef_pix_to_mm;
                var _hMm = _hPx * g_dKoef_pix_to_mm;
        
                ctx.fillStyle = "#FFFFFF";
                ctx.fillRect(0, 0, _wPx, _hPx);
        
                var _pxBoundsW = par.Lines[0].Ranges[0].W * g_dKoef_mm_to_pix;//(_bounds.Right - _bounds.Left) * g_dKoef_mm_to_pix;
                var _pxBoundsH = (_bounds.Bottom - _bounds.Top) * g_dKoef_mm_to_pix;
        
                if (this.GuiLastTextProps.Position !== undefined && this.GuiLastTextProps.Position != null && this.GuiLastTextProps.Position != 0)
                {
                    // TODO: нужна высота без учета Position
                    // _pxBoundsH -= (this.GuiLastTextProps.Position * g_dKoef_mm_to_pix);
                }
        
                if (_pxBoundsH < _hPx && _pxBoundsW < _wPx)
                {
                    // рисуем линию
                    var _lineY = (((_hPx + _pxBoundsH) / 2) >> 0) + 0.5;
                    var _lineW = (((_wPx - _pxBoundsW) / 4) >> 0);
        
                    ctx.strokeStyle = "#000000";
                    ctx.lineWidth   = 1;
        
                    ctx.beginPath();
                    ctx.moveTo(0, _lineY);
                    ctx.lineTo(_lineW, _lineY);
        
                    ctx.moveTo(_wPx - _lineW, _lineY);
                    ctx.lineTo(_wPx, _lineY);
        
                    ctx.stroke();
                    ctx.beginPath();
                }
        
                var _yOffset = (((_hPx + _pxBoundsH) / 2) - baseLineOffset * g_dKoef_mm_to_pix) >> 0;
                var _xOffset = ((_wPx - _pxBoundsW) / 2) >> 0;
        
                var graphics = new AscCommon.CGraphics();
                graphics.init(ctx, _wPx, _hPx, _wMm, _hMm);
                graphics.m_oFontManager = AscCommon.g_fontManager;
        
                graphics.m_oCoordTransform.tx = _xOffset;
                graphics.m_oCoordTransform.ty = _yOffset;
        
                graphics.transform(1, 0, 0, 1, 0, 0);
        
                var old_marks                            = this.m_oWordControl.m_oApi.ShowParaMarks;
                this.m_oWordControl.m_oApi.ShowParaMarks = false;
                par.Draw(0, graphics);
                this.m_oWordControl.m_oApi.ShowParaMarks = old_marks;
                editor.isViewMode = _oldTurn;
            }, this, []);
        };
        this.DrawImageTextureFillTextArt = function(url) {
            if (this.GuiCanvasFillTextureTextArt == null)
            {
                this.InitGuiCanvasTextArt(this.GuiCanvasFillTextureParentIdTextArt);
            }

            if (this.GuiCanvasFillTextureTextArt == null || this.GuiCanvasFillTextureCtxTextArt == null || url == this.LastDrawingUrlTextArt)
                return;

            this.LastDrawingUrlTextArt = url;
            var _width                 = this.GuiCanvasFillTextureTextArt.width;
            var _height                = this.GuiCanvasFillTextureTextArt.height;

            this.GuiCanvasFillTextureCtxTextArt.clearRect(0, 0, _width, _height);

            if (null == this.LastDrawingUrlTextArt)
                return;

            var _img = this.m_oWordControl.m_oApi.ImageLoader.map_image_index[AscCommon.getFullImageSrc2(this.LastDrawingUrlTextArt)];
            if (_img != undefined && _img.Image != null && _img.Status != AscFonts.ImageLoadStatus.Loading)
            {
                var _x = 0;
                var _y = 0;
                var _w = Math.max(_img.Image.width, 1);
                var _h = Math.max(_img.Image.height, 1);

                var dAspect1 = _width / _height;
                var dAspect2 = _w / _h;

                _w = _width;
                _h = _height;
                if (dAspect1 >= dAspect2)
                {
                    _w = dAspect2 * _height;
                    _x = (_width - _w) / 2;
                }
                else
                {
                    _h = _w / dAspect2;
                    _y = (_height - _h) / 2;
                }

                this.GuiCanvasFillTextureCtxTextArt.drawImage(_img.Image, _x, _y, _w, _h);
            }
            else
            {
                this.GuiCanvasFillTextureCtxTextArt.lineWidth = 1;

                this.GuiCanvasFillTextureCtxTextArt.beginPath();
                this.GuiCanvasFillTextureCtxTextArt.moveTo(0, 0);
                this.GuiCanvasFillTextureCtxTextArt.lineTo(_width, _height);
                this.GuiCanvasFillTextureCtxTextArt.moveTo(_width, 0);
                this.GuiCanvasFillTextureCtxTextArt.lineTo(0, _height);
                this.GuiCanvasFillTextureCtxTextArt.strokeStyle = "#FF0000";
                this.GuiCanvasFillTextureCtxTextArt.stroke();

                this.GuiCanvasFillTextureCtxTextArt.beginPath();
                this.GuiCanvasFillTextureCtxTextArt.moveTo(0, 0);
                this.GuiCanvasFillTextureCtxTextArt.lineTo(_width, 0);
                this.GuiCanvasFillTextureCtxTextArt.lineTo(_width, _height);
                this.GuiCanvasFillTextureCtxTextArt.lineTo(0, _height);
                this.GuiCanvasFillTextureCtxTextArt.closePath();

                this.GuiCanvasFillTextureCtxTextArt.strokeStyle = "#000000";
                this.GuiCanvasFillTextureCtxTextArt.stroke();
                this.GuiCanvasFillTextureCtxTextArt.beginPath();
            }
        };
        this.InitGuiCanvasTextArt = function(div_id) {
            if (null != this.GuiCanvasFillTextureTextArt)
            {
                var _div_elem = document.getElementById(this.GuiCanvasFillTextureParentIdTextArt);
                if (!_div_elem)
                    _div_elem.removeChild(this.GuiCanvasFillTextureTextArt);

                this.GuiCanvasFillTextureTextArt    = null;
                this.GuiCanvasFillTextureCtxTextArt = null;
            }

            this.GuiCanvasFillTextureParentIdTextArt = div_id;
            var _div_elem                            = document.getElementById(this.GuiCanvasFillTextureParentIdTextArt);
            if (!_div_elem)
                return;

            this.GuiCanvasFillTextureTextArt        = document.createElement('canvas');
            this.GuiCanvasFillTextureTextArt.width  = parseInt(_div_elem.style.width);
            this.GuiCanvasFillTextureTextArt.height = parseInt(_div_elem.style.height);

            this.LastDrawingUrlTextArt          = "";
            this.GuiCanvasFillTextureCtxTextArt = this.GuiCanvasFillTextureTextArt.getContext('2d');

            _div_elem.appendChild(this.GuiCanvasFillTextureTextArt);
        };
        this.InitGuiCanvasShape = function (div_id) {
            if (null != this.GuiCanvasFillTexture)
            {
                var _div_elem = document.getElementById(this.GuiCanvasFillTextureParentId);
                if (_div_elem)
                    _div_elem.removeChild(this.GuiCanvasFillTexture);

                this.GuiCanvasFillTexture = null;
                this.GuiCanvasFillTextureCtx = null;
            }

            this.GuiCanvasFillTextureParentId = div_id;
            var _div_elem = document.getElementById(this.GuiCanvasFillTextureParentId);
            if (!_div_elem)
                return;

            this.GuiCanvasFillTexture = document.createElement('canvas');
            this.GuiCanvasFillTexture.width = parseInt(_div_elem.style.width);
            this.GuiCanvasFillTexture.height = parseInt(_div_elem.style.height);

            this.LastDrawingUrl = "";
            this.GuiCanvasFillTextureCtx = this.GuiCanvasFillTexture.getContext('2d');

            _div_elem.appendChild(this.GuiCanvasFillTexture);
        };
        this.DrawTarget = function() {
            let oDoc            = Asc.editor.getPDFDoc();
            let oActiveObj      = oDoc.GetActiveObject();
            let isViewObject    = oActiveObj && (oActiveObj.IsAnnot() || oActiveObj.IsForm());

            if (oThis.NeedTarget) {
                let isActive = true;
                let api = oThis.m_oWordControl.m_oApi;

                if (!oThis.m_oWordControl.IsFocus)
                    isActive = false;
                else if (oThis.m_oWordControl.m_oApi.isBlurEditor)
                    isActive = false;
                else if (api.isViewMode || (api.isRestrictionView() && !isViewObject))
                    isActive = false;

                if (isActive)
                    oThis.showTarget(!oThis.isShowTarget());
                else
                    oThis.showTarget(true);
            }
        };
        this.isHideTarget = function() {
            let oDoc            = Asc.editor.getPDFDoc();
            let oActiveObj      = oDoc.GetActiveObject();
            let isViewObject    = oActiveObj && (oActiveObj.IsAnnot() || oActiveObj.IsForm());
            let api             = this.m_oWordControl.m_oApi;

            if (api.isViewMode || (api.isRestrictionView() && !api.isRestrictionForms() && !isViewObject))
                return this.isHideTargetBeforeFirstClick;
            return false;
        };
        this.TargetStart = function (bShowHide) {
            if (this.m_lTimerTargetId != -1)
                clearInterval(this.m_lTimerTargetId);
            this.m_lTimerTargetId = setInterval(oThis.DrawTarget, 500);

            this.showTarget(bShowHide);
        };
        this.Collaborative_UpdateTarget = function (_id, _shortId, _x, _y, _size, _page, _transform, is_from_paint) {
            this.AutoShapesTrack.SetCurrentPage(_page, true);
            let TextMatrix = this.AutoShapesTrack.transformPageMatrix(_transform);

            if (is_from_paint !== true) {
                this.CollaborativeTargetsUpdateTasks.push([_id, _shortId, _x, _y, _size, _page, _transform]);
                return;
            }

            for (let i = 0; i < this.CollaborativeTargets.length; i++) {
                if (_id == this.CollaborativeTargets[i].Id) {
                    this.CollaborativeTargets[i].CheckPosition(_x, _y, _size, _page, TextMatrix);
                    return;
                }
            }

            let _target = new CDrawingCollaborativeTarget(this);
            _target.Id = _id;
            _target.ShortId = _shortId;
            _target.CheckPosition(_x, _y, _size, _page, TextMatrix);
            this.CollaborativeTargets[this.CollaborativeTargets.length] = _target;
        };
        this.OnRecalculatePage = function() {};
        this.OnEndRecalculate = function() {};
        this.ConvertCoordsToAnotherPage = function (x, y, pageCoord, pageNeed) {
            return AscPDF.ConvertCoordsToAnotherPage(x, y, pageCoord, pageNeed);
        }

        this.CheckRasterImageOnScreen = function(src)
        {
            let redrawPages = [];
            let viewer = this.m_oDocumentRenderer;
            let thumbnails = viewer.thumbnails;

            let thumbStartVisiblePage = thumbnails && thumbnails.getStartVisiblePage();
            let thumbEndVisiblePage = thumbnails && thumbnails.getEndVisiblePage();

            let startVisiblePage = thumbStartVisiblePage != undefined ? thumbStartVisiblePage : viewer.startVisiblePage;
            let endVisiblePage = thumbEndVisiblePage != undefined ? thumbEndVisiblePage : viewer.endVisiblePage;

            for (let i = startVisiblePage; i <= endVisiblePage; i++)
            {
                let imgs = viewer.DrawingObjects.getAllRasterImagesOnPage(i);
                for (let j = 0, len = imgs.length; j < len; j++)
                {
                    if (AscCommon.getFullImageSrc2(imgs[j]) === src)
                    {
                        redrawPages.push(i);
                        break;
                    }
                }
            }

            if (redrawPages.length > 0)
            {
                viewer.onUpdatePages(redrawPages);
                viewer.onRepaintForms(redrawPages);
            }
        };
    }

    CDrawingDocument.prototype.constructor = CDrawingDocument;

    window["AscPDF"].CDrawingDocument = CDrawingDocument;

})(window);

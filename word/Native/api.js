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

var sdkCheck = true;
var spellCheck = true;
var _api = null;
var _internalStorage = {
    changesReview : []
};

// SERIALIZE
function asc_menu_WriteHeaderFooterPr(_hdrftrPr, _stream)
{
    if (_hdrftrPr.Type !== undefined && _hdrftrPr.Type !== null)
    {
        _stream["WriteByte"](0);
        _stream["WriteLong"](_hdrftrPr.Type);
    }
    if (_hdrftrPr.Position !== undefined && _hdrftrPr.Position !== null)
    {
        _stream["WriteByte"](1);
        _stream["WriteDouble2"](_hdrftrPr.Position);
    }

    if (_hdrftrPr.DifferentFirst !== undefined && _hdrftrPr.DifferentFirst !== null)
    {
        _stream["WriteByte"](2);
        _stream["WriteBool"](_hdrftrPr.DifferentFirst);
    }
    if (_hdrftrPr.DifferentEvenOdd !== undefined && _hdrftrPr.DifferentEvenOdd !== null)
    {
        _stream["WriteByte"](3);
        _stream["WriteBool"](_hdrftrPr.DifferentEvenOdd);
    }
    if (_hdrftrPr.LinkToPrevious !== undefined && _hdrftrPr.LinkToPrevious !== null)
    {
        _stream["WriteByte"](4);
        _stream["WriteBool"](_hdrftrPr.LinkToPrevious);
    }
    if (_hdrftrPr.Locked !== undefined && _hdrftrPr.Locked !== null)
    {
        _stream["WriteByte"](5);
        _stream["WriteBool"](_hdrftrPr.DifferentFirst);
    }

    _stream["WriteByte"](255);
}

function Deserialize_Table_Markup(_params, _cols, _margins, _rows)
{
    var _markup = new CTableMarkup(null);
    _markup.Internal.RowIndex   = _params[0];
    _markup.Internal.CellIndex  = _params[1];
    _markup.Internal.PageNum    = _params[2];
    _markup.X                   = _params[3];
    _markup.CurCol              = _params[4];
    _markup.CurRow              = _params[5];
    // 6 - DragPos
    _markup.TransformX          = _params[7];
    _markup.TransformY          = _params[8];

    _markup.Cols    = _cols;

    var _len = _margins.length;
    for (var i = 0; i < _len; i += 2)
    {
        _markup.Margins.push({ Left : _margins[i], Right : _margins[i + 1] });
    }

    _len = _rows.length;
    for (var i = 0; i < _len; i += 2)
    {
        _markup.Rows.push({ Y : _rows[i], H : _rows[i + 1] });
    }

    return _markup;
}

// editor
Asc['asc_docs_api'].prototype["NativeAfterLoad"] = function()
{
    this.WordControl.m_oDrawingDocument.AfterLoad();
    this.WordControl.m_oLogicDocument.SetUseTextShd(false);
};
Asc['asc_docs_api'].prototype["GetNativePageMeta"] = function(pageIndex)
{
    this.WordControl.m_oDrawingDocument.LogicDocument = _api.WordControl.m_oDrawingDocument.m_oLogicDocument;
    this.WordControl.m_oDrawingDocument.RenderPage(pageIndex);
};

Asc['asc_docs_api'].prototype["Native_Editor_Initialize_Settings"] = function(_params)
{
    window["NativeSupportTimeouts"] = true;

    if (!_params)
        return;

    var _current = { pos : 0 };
    var _continue = true;
    while (_continue)
    {
        var _attr = _params[_current.pos++];
        switch (_attr)
        {
            case 0:
            {
                AscCommon.GlobalSkin.STYLE_THUMBNAIL_WIDTH = _params[_current.pos++];
                break;
            }
            case 1:
            {
                AscCommon.GlobalSkin.STYLE_THUMBNAIL_HEIGHT = _params[_current.pos++];
                break;
            }
            case 2:
            {
                TABLE_STYLE_WIDTH_PIX = _params[_current.pos++];
                break;
            }
            case 3:
            {
                TABLE_STYLE_HEIGHT_PIX = _params[_current.pos++];
                break;
            }
            case 4:
            {
                this.chartPreviewManager.CHART_PREVIEW_WIDTH_PIX = _params[_current.pos++];
                break;
            }
            case 5:
            {
                this.chartPreviewManager.CHART_PREVIEW_HEIGHT_PIX = _params[_current.pos++];
                break;
            }
            case 6:
            {
                var _val = _params[_current.pos++];
                if (_val === true)
                {
                    this.ShowParaMarks = false;
                    AscCommon.CollaborativeEditing.Set_GlobalLock(true);

                    this.isViewMode = true;
                    this.WordControl.m_oDrawingDocument.IsViewMode = true;
                }
                break;
            }
            case 100:
            {
                this.WordControl.m_oDrawingDocument.IsRetina = _params[_current.pos++];
                break;
            }
            case 101:
            {
                this.WordControl.m_oDrawingDocument.IsMobile = _params[_current.pos++];
                window.AscAlwaysSaveAspectOnResizeTrack = true;
                break;
            }
            case 255:
            default:
            {
                _continue = false;
                break;
            }
        }
    }

    AscCommon.AscBrowser.isRetina = this.WordControl.m_oDrawingDocument.IsRetina;
};

// HTML page interface
Asc['asc_docs_api'].prototype["Call_OnUpdateOverlay"] = function(param)
{
    this.WordControl.m_oDrawingDocument.OnUpdateOverlay();
};

Asc['asc_docs_api'].prototype["Call_OnMouseDown"] = function(e)
{
    return this.WordControl.m_oDrawingDocument.OnMouseDown(e);
};
Asc['asc_docs_api'].prototype["Call_OnMouseUp"] = function(e)
{
    return this.WordControl.m_oDrawingDocument.OnMouseUp(e);
};
Asc['asc_docs_api'].prototype["Call_OnMouseMove"] = function(e)
{
    return this.WordControl.m_oDrawingDocument.OnMouseMove(e);
};
Asc['asc_docs_api'].prototype["Call_OnCheckMouseDown"] = function(e)
{
    return this.WordControl.m_oDrawingDocument.OnCheckMouseDown(e);
};

Asc['asc_docs_api'].prototype["Call_OnKeyDown"] = function(e)
{
    this.WordControl.m_oDrawingDocument.OnKeyDown(e);
};
Asc['asc_docs_api'].prototype["Call_OnKeyPress"] = function(e)
{
    this.WordControl.m_oDrawingDocument.OnKeyPress(e);
};
Asc['asc_docs_api'].prototype["Call_OnKeyUp"] = function(e)
{
    this.WordControl.m_oDrawingDocument.OnKeyUp(e);
};
Asc['asc_docs_api'].prototype["Call_OnKeyboardEvent"] = function(e)
{
    this.WordControl.m_oDrawingDocument.OnKeyboardEvent(e);
};

Asc['asc_docs_api'].prototype["Call_CalculateResume"] = function()
{
    Document_Recalculate_Page();
};

Asc['asc_docs_api'].prototype["Call_TurnOffRecalculate"] = function()
{
    this.WordControl.m_oLogicDocument.TurnOff_Recalculate();
};
Asc['asc_docs_api'].prototype["Call_TurnOnRecalculate"] = function()
{
    this.WordControl.m_oLogicDocument.TurnOn_Recalculate();
    this.WordControl.m_oLogicDocument.Recalculate();
};

Asc['asc_docs_api'].prototype["Call_CheckTargetUpdate"] = function()
{
    this.WordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
    this.WordControl.m_oLogicDocument.CheckTargetUpdate();
    this.WordControl.m_oDrawingDocument.CheckTargetShow();
    this.WordControl.m_oDrawingDocument.UpdateTargetFromPaint = false;

    this.WordControl.m_oDrawingDocument.Collaborative_TargetsUpdate(false);
};

Asc['asc_docs_api'].prototype["Call_Common"] = function(type, param)
{
    switch (type)
    {
        case 1:
        {
            this.WordControl.m_oLogicDocument.MoveCursorLeft();
            break;
        }
        case 67:
        {
            this.startGetDocInfo();
            break;
        }
        case 68:
        {
            this.stopGetDocInfo();
            break;
        }
        default:
            break;
    }
};

Asc['asc_docs_api'].prototype["Call_HR_Tabs"] = function(arrT, arrP)
{
    var _arr = new AscCommonWord.CParaTabs();
    var _c = arrT.length;
    for (var i = 0; i < _c; i++)
    {
        if (arrT[i] == 1)
            _arr.Add( new CParaTab( tab_Left, arrP[i] ) );
        if (arrT[i] == 2)
            _arr.Add( new CParaTab( tab_Right, arrP[i] ) );
        if (arrT[i] == 3)
            _arr.Add( new CParaTab( tab_Center, arrP[i] ) );
    }

    var _logic = this.WordControl.m_oLogicDocument;
    if ( false === _logic.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Properties) )
    {
        _logic.StartAction();
        _logic.SetParagraphTabs(_arr);
        _logic.FinalizeAction();
    }
};
Asc['asc_docs_api'].prototype["Call_HR_Pr"] = function(_indent_left, _indent_right, _indent_first)
{
    var _logic = this.WordControl.m_oLogicDocument;
    if ( false === _logic.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Properties) )
    {
        _logic.StartAction();
        _logic.SetParagraphIndent( { Left : _indent_left, Right : _indent_right, FirstLine: _indent_first } );
		_logic.FinalizeAction();
    }
};
Asc['asc_docs_api'].prototype["Call_HR_Margins"] = function(_margin_left, _margin_right)
{
    var _logic = this.WordControl.m_oLogicDocument;
    if ( false === _logic.Document_Is_SelectionLocked(AscCommon.changestype_Document_SectPr) )
    {
        _logic.StartAction();
        _logic.SetDocumentMargin( { Left : _margin_left, Right : _margin_right });
		_logic.FinalizeAction();
    }
};
Asc['asc_docs_api'].prototype["Call_HR_Table"] = function(_params, _cols, _margins, _rows)
{
    var _logic = this.WordControl.m_oLogicDocument;
    if ( false === _logic.Document_Is_SelectionLocked(AscCommon.changestype_Table_Properties) )
    {
        _logic.StartAction();

        var _table_murkup = Deserialize_Table_Markup(_params, _cols, _margins, _rows);
        _table_murkup.Table = this.WordControl.m_oDrawingDocument.Table;

        _table_murkup.CorrectTo();
        _table_murkup.Table.Update_TableMarkupFromRuler(_table_murkup, true, _params[6]);
        _table_murkup.CorrectFrom();

        _logic.FinalizeAction();
    }
};

Asc['asc_docs_api'].prototype["Call_VR_Margins"] = function(_top, _bottom)
{
    var _logic = this.WordControl.m_oLogicDocument;
    if ( false === _logic.Document_Is_SelectionLocked(AscCommon.changestype_Document_SectPr) )
    {
        _logic.StartAction();
        _logic.SetDocumentMargin( { Top : _top, Bottom : _bottom });
        _logic.FinalizeAction();
    }
};
Asc['asc_docs_api'].prototype["Call_VR_Header"] = function(_header_top, _header_bottom)
{
    var _logic = this.WordControl.m_oLogicDocument;
    if ( false === _logic.Document_Is_SelectionLocked(AscCommon.changestype_HdrFtr) )
    {
        _logic.StartAction();
        _logic.Document_SetHdrFtrBounds(_header_top, _header_bottom);
        _logic.FinalizeAction();
    }
};
Asc['asc_docs_api'].prototype["Call_VR_Table"] = function(_params, _cols, _margins, _rows)
{
    var _logic = this.WordControl.m_oLogicDocument;
    if ( false === _logic.Document_Is_SelectionLocked(AscCommon.changestype_Table_Properties) )
    {
        _logic.StartAction();

        var _table_murkup = Deserialize_Table_Markup(_params, _cols, _margins, _rows);
        _table_murkup.Table = this.WordControl.m_oDrawingDocument.Table;

        _table_murkup.CorrectTo();
        _table_murkup.Table.Update_TableMarkupFromRuler(_table_murkup, false, _params[6]);
        _table_murkup.CorrectFrom();

        _logic.FinalizeAction();
    }
};

Asc['asc_docs_api'].prototype["Call_Menu_Event"] = function(type, _params)
{
    if (this.WordControl.m_oDrawingDocument.m_bIsMouseLockDocument)
    {
        // не делаем ничего. Как в веб версии отрубаем клавиатуру
        return undefined;
    }

    var _return = undefined;
    var _current = { pos : 0 };
    var _continue = true;
    switch (type)
    {
        case 1: // ASC_MENU_EVENT_TYPE_TEXTPR
        {
            var _textPr = new AscCommonWord.CTextPr();
            while (_continue)
            {
                var _attr = _params[_current.pos++];
                switch (_attr)
                {
                    case 0:
                    {
                        _textPr.Bold = _params[_current.pos++];
                        break;
                    }
                    case 1:
                    {
                        _textPr.Italic = _params[_current.pos++];
                        break;
                    }
                    case 2:
                    {
                        _textPr.Underline = _params[_current.pos++];
                        break;
                    }
                    case 3:
                    {
                        _textPr.Strikeout = _params[_current.pos++];
                        break;
                    }
                    case 4:
                    {
                        _textPr.FontFamily = asc_menu_ReadFontFamily(_params, _current);
                        break;
                    }
                    case 5:
                    {
                        _textPr.FontSize = _params[_current.pos++];
                        break;
                    }
                    case 6:
                    {
                        var Unifill = new AscFormat.CUniFill();
                        Unifill.fill = new AscFormat.CSolidFill();
                        var color = AscCommon.asc_menu_ReadColor(_params, _current);
                        Unifill.fill.color = AscFormat.CorrectUniColor(color, Unifill.fill.color, 1);
                        _textPr.Unifill = Unifill;
                        break;
                    }
                    case 7:
                    {
                        _textPr.VertAlign = _params[_current.pos++];
                        break;
                    }
                    case 8:
                    {
                        var color = AscCommon.asc_menu_ReadColor(_params, _current);
                        if (color.a < 1) {
                            _textPr.HighLight = AscCommonWord.highlight_None;
                        } else {
                            _textPr.HighLight = { r: color.r, g: color.g, b: color.b };
                        }
                        break;
                    }
                    case 9:
                    {
                        _textPr.DStrikeout = _params[_current.pos++];
                        break;
                    }
                    case 10:
                    {
                        _textPr.Caps = _params[_current.pos++];
                        break;
                    }
                    case 11:
                    {
                        _textPr.SmallCaps = _params[_current.pos++];
                        break;
                    }
                    case 12:
                    {
                        _textPr.HighLight = AscCommonWord.highlight_None;
                        break;
                    }
                    case 13:
                    {
                        _textPr.Spacing = _params[_current.pos++];
                        break;
                    }
                    case 255:
                    default:
                    {
                        _continue = false;
                        break;
                    }
                }
            }

            this.WordControl.m_oLogicDocument.StartAction();
            this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr(_textPr));
            this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
            break;
        }
        case 2: // ASC_MENU_EVENT_TYPE_PARAPR
        {
            var _textPr = undefined;

            this.WordControl.m_oLogicDocument.StartAction();

            while (_continue)
            {
                var _attr = _params[_current.pos++];
                switch (_attr)
                {
                    case 0:
                    {
                        this.WordControl.m_oLogicDocument.SetParagraphContextualSpacing( _params[_current.pos++] );
                        break;
                    }
                    case 1:
                    {
                        var _ind = asc_menu_ReadParaInd(_params, _current);
                        this.WordControl.m_oLogicDocument.SetParagraphIndent( _ind );
                        break;
                    }
                    case 2:
                    {
                        this.WordControl.m_oLogicDocument.SetParagraphKeepLines( _params[_current.pos++] );
                        break;
                    }
                    case 3:
                    {
                        this.WordControl.m_oLogicDocument.SetParagraphKeepNext( _params[_current.pos++] );
                        break;
                    }
                    case 4:
                    {
                        this.WordControl.m_oLogicDocument.SetParagraphWidowControl( _params[_current.pos++] );
                        break;
                    }
                    case 5:
                    {
                        this.WordControl.m_oLogicDocument.SetParagraphPageBreakBefore( _params[_current.pos++] );
                        break;
                    }
                    case 6:
                    {
                        var _spacing = asc_menu_ReadParaSpacing(_params, _current);
                        this.WordControl.m_oLogicDocument.SetParagraphSpacing( _spacing );
                        break;
                    }
                    case 7:
                    {
                        // TODO:
                        var _brds = asc_menu_ReadParaBorders(_params, _current);

                        if (_brds.Left && _brds.Left.Color)
                        {
                            _brds.Left.Unifill = AscFormat.CreateUnifillFromAscColor(_brds.Left.Color);
                        }
                        if (_brds.Top && _brds.Top.Color)
                        {
                            _brds.Top.Unifill = AscFormat.CreateUnifillFromAscColor(_brds.Top.Color);
                        }
                        if (_brds.Right && _brds.Right.Color)
                        {
                            _brds.Right.Unifill = AscFormat.CreateUnifillFromAscColor(_brds.Right.Color);
                        }
                        if (_brds.Bottom && _brds.Bottom.Color)
                        {
                            _brds.Bottom.Unifill = AscFormat.CreateUnifillFromAscColor(_brds.Bottom.Color);
                        }

                        this.WordControl.m_oLogicDocument.SetParagraphBorders( _brds );
                        break;
                    }
                    case 8:
                    {
                        var _shd = asc_menu_ReadParaShd(_params, _current);
                        this.WordControl.m_oLogicDocument.SetParagraphShd( _shd );
                        break;
                    }
                    case 9:
                    case 10:
                    case 11:
                    {
                        // nothing
                        _current.pos++;
                        break;
                    }
                    case 12:
                    {
                        this.WordControl.m_oLogicDocument.Set_DocumentDefaultTab( _params[_current.pos++] );
                        break;
                    }
                    case 13:
                    {
                        var _tabs = asc_menu_ReadParaTabs(_params, _current);
                        // TODO:
                        this.WordControl.m_oLogicDocument.SetParagraphTabs( _tabs.Tabs );
                        break;
                    }
                    case 14:
                    {
                        var _framePr = asc_menu_ReadParaFrame(_params, _current);
                        this.WordControl.m_oLogicDocument.SetParagraphFramePr( _framePr );
                        break;
                    }
                    case 15:
                    {
                        if (_textPr === undefined)
                            _textPr = new AscCommonWord.CTextPr();
                        if (true == _params[_current.pos++])
                            _textPr.VertAlign = AscCommon.vertalign_SubScript;
                        else
                            _textPr.VertAlign = AscCommon.vertalign_Baseline;
                        break;
                    }
                    case 16:
                    {
                        if (_textPr === undefined)
                            _textPr = new AscCommonWord.CTextPr();
                        if (true == _params[_current.pos++])
                            _textPr.VertAlign = AscCommon.vertalign_SuperScript;
                        else
                            _textPr.VertAlign = AscCommon.vertalign_Baseline;
                        break;
                    }
                    case 17:
                    {
                        if (_textPr === undefined)
                            _textPr = new AscCommonWord.CTextPr();
                        _textPr.SmallCaps = _params[_current.pos++];
                        _textPr.Caps   = false;
                        break;
                    }
                    case 18:
                    {
                        if (_textPr === undefined)
                            _textPr = new AscCommonWord.CTextPr();
                        _textPr.Caps = _params[_current.pos++];
                        if (true == _textPr.Caps)
                            _textPr.SmallCaps = false;
                        break;
                    }
                    case 19:
                    {
                        if (_textPr === undefined)
                            _textPr = new AscCommonWord.CTextPr();
                        _textPr.Strikeout  = _params[_current.pos++];
                        _textPr.DStrikeout = false;
                        break;
                    }
                    case 20:
                    {
                        if (_textPr === undefined)
                            _textPr = new AscCommonWord.CTextPr();
                        _textPr.DStrikeout  = _params[_current.pos++];
                        if (true == _textPr.DStrikeout)
                            _textPr.Strikeout = false;
                        break;
                    }
                    case 21:
                    {
                        if (_textPr === undefined)
                            _textPr = new AscCommonWord.CTextPr();
                        _textPr.TextSpacing = _params[_current.pos++];
                        break;
                    }
                    case 22:
                    {
                        if (_textPr === undefined)
                            _textPr = new AscCommonWord.CTextPr();
                        _textPr.Position = _params[_current.pos++];
                        break;
                    }
                    case 23:
                    {
                        var _listType = asc_menu_ReadParaListType(_params, _current);
						this.put_ListType(_listType.asc_getListType(), _listType.asc_getListSubType());
                        break;
                    }
                    case 24:
                    {
                        this.WordControl.m_oLogicDocument.SetParagraphStyle( _params[_current.pos++] );
                        break;
                    }
                    case 25:
                    {
                        this.WordControl.m_oLogicDocument.SetParagraphAlign( _params[_current.pos++] );
                        break;
                    }
                    case 255:
                    default:
                    {
                        _continue = false;
                        break;
                    }
                }
            }

            if (undefined !== _textPr)
                this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr(_textPr));

			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
            break;
        }
        case 22003: //ASC_MENU_EVENT_TYPE_ON_EDIT_TEXT
        {
            var oController = this.WordControl.m_oLogicDocument.DrawingObjects;
            if(oController)
            {
                oController.startEditTextCurrentShape();
            }
            break;
        }
        case 3: // ASC_MENU_EVENT_TYPE_UNDO
        {
            this.WordControl.m_oLogicDocument.Document_Undo();
            break;
        }
        case 4: // ASC_MENU_EVENT_TYPE_REDO
        {
            this.WordControl.m_oLogicDocument.Document_Redo();
            break;
        }
        case 7: // ASC_MENU_EVENT_TYPE_HEADERFOOTER
        {
            var bIsApply = (this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_HdrFtr) === false) ? true : false;

            if (bIsApply)
                this.WordControl.m_oLogicDocument.StartAction();

            while (_continue)
            {
                var _attr = _params[_current.pos++];
                switch (_attr)
                {
                    case 0:
                    {
                        _current.pos++;
                        break;
                    }
                    case 1:
                    {
                        if (bIsApply)
                            this.WordControl.m_oLogicDocument.Document_SetHdrFtrDistance(_params[_current.pos]);

                        _current.pos++;
                        break;
                    }
                    case 2:
                    {
                        if (bIsApply)
                            this.WordControl.m_oLogicDocument.Document_SetHdrFtrFirstPage(_params[_current.pos]);

                        _current.pos++;
                        break;
                    }
                    case 3:
                    {
                        if (bIsApply)
                            this.WordControl.m_oLogicDocument.Document_SetHdrFtrEvenAndOddHeaders(_params[_current.pos]);

                        _current.pos++;
                        break;
                    }
                    case 4:
                    {
                        if (bIsApply)
                            this.WordControl.m_oLogicDocument.Document_SetHdrFtrLink(_params[_current.pos]);

                        _current.pos++;
                        break;
                    }
                    case 5:
                    {
                        _current.pos++;
                        break;
                    }
                    case 255:
                    default:
                    {
                        _continue = false;
                        break;
                    }
                }
            }

            if (bIsApply)
            	this.WordControl.m_oLogicDocument.FinalizeAction();

            break;
        }
        case 13: // ASC_MENU_EVENT_TYPE_INCREASEPARAINDENT
        {
            this.IncreaseIndent();
            break;
        }
        case 14: // ASC_MENU_EVENT_TYPE_DECREASEPARAINDENT
        {
            this.DecreaseIndent();
            break;
        }
        case 54: // ASC_MENU_EVENT_TYPE_INSERT_PAGEBREAK
        {
            this.put_AddPageBreak();
            break;
        }
        case 55: // ASC_MENU_EVENT_TYPE_INSERT_LINEBREAK
        {
            this.put_AddLineBreak();
            break;
        }
        case 56: // ASC_MENU_EVENT_TYPE_INSERT_PAGENUMBER
        {
            if (_params[0] < 0) {
                this.put_PageNum(-1);
            } else {
                this.put_PageNum((_params[0] >> 16) & 0xFFFF, _params[0] & 0xFFFF);
            }
            break;
        }
        case 57: // ASC_MENU_EVENT_TYPE_INSERT_SECTIONBREAK
        {
            this.add_SectionBreak(_params[0]);
            break;
        }
        case 10: // ASC_MENU_EVENT_TYPE_TABLE
        {
            var _tablePr = new Asc.CTableProp();
            while (_continue)
            {
                var _attr = _params[_current.pos++];
                switch (_attr)
                {
                    case 0:
                    {
                        _tablePr.CanBeFlow = _params[_current.pos++];
                        break;
                    }
                    case 1:
                    {
                        _tablePr.CellSelect = _params[_current.pos++];
                        break;
                    }
                    case 2:
                    {
                        _tablePr.TableWidth = _params[_current.pos++];
                        break;
                    }
                    case 3:
                    {
                        _tablePr.TableSpacing = _params[_current.pos++];
                        break;
                    }
                    case 4:
                    {
                        _tablePr.TableDefaultMargins = AscCommon.asc_menu_ReadPaddings(_params, _current);
                        break;
                    }
                    case 5:
                    {
                        _tablePr.CellMargins = asc_menu_ReadCellMargins(_params, _current);
                        break;
                    }
                    case 6:
                    {
                        _tablePr.TableAlignment = _params[_current.pos++];
                        break;
                    }
                    case 7:
                    {
                        _tablePr.TableIndent = _params[_current.pos++];
                        break;
                    }
                    case 8:
                    {
                        _tablePr.TableWrappingStyle = _params[_current.pos++];
                        break;
                    }
                    case 9:
                    {
                        _tablePr.TablePaddings = AscCommon.asc_menu_ReadPaddings(_params, _current);
                        break;
                    }
                    case 10:
                    {
                        _tablePr.TableBorders = asc_menu_ReadCellBorders(_params, _current);
                        break;
                    }
                    case 11:
                    {
                        _tablePr.CellBorders = asc_menu_ReadCellBorders(_params, _current);
                        break;
                    }
                    case 12:
                    {
                        _tablePr.TableBackground = asc_menu_ReadCellBackground(_params, _current);
                        break;
                    }
                    case 13:
                    {
                        _tablePr.CellsBackground = asc_menu_ReadCellBackground(_params, _current);
                        break;
                    }
                    case 14:
                    {
                        _tablePr.Position = asc_menu_ReadPosition(_params, _current);
                        break;
                    }
                    case 15:
                    {
                        _tablePr.PositionH = asc_menu_ReadImagePosition(_params, _current);
                        break;
                    }
                    case 16:
                    {
                        _tablePr.PositionV = asc_menu_ReadImagePosition(_params, _current);
                        break;
                    }
                    case 17:
                    {
                        _tablePr.Internal_Position = asc_menu_ReadTableAnchorPosition(_params, _current);
                        break;
                    }
                    case 18:
                    {
                        _tablePr.ForSelectedCells = _params[_current.pos++];
                        break;
                    }
                    case 19:
                    {
                        _tablePr.TableStyle = _params[_current.pos++];
                        break;
                    }
                    case 20:
                    {
                        _tablePr.TableLook = asc_menu_ReadTableLook(_params, _current);
                        break;
                    }
                    case 21:
                    {
                        _tablePr.RowsInHeader = _params[_current.pos++];
                        break;
                    }
                    case 22:
                    {
                        _tablePr.CellsVAlign = _params[_current.pos++];
                        break;
                    }
                    case 23:
                    {
                        _tablePr.AllowOverlap = _params[_current.pos++];
                        break;
                    }
                    case 24:
                    {
                        _tablePr.TableLayout = _params[_current.pos++];
                        break;
                    }
                    case 25:
                    {
                        _tablePr.Locked = _params[_current.pos++];
                        break;
                    }
                    case 255:
                    default:
                    {
                        _continue = false;
                        break;
                    }
                }
            }

            this.tblApply(_tablePr);
            break;
        }
        case 9 : // ASC_MENU_EVENT_TYPE_IMAGE
        {
            var _imagePr = new Asc.asc_CImgProperty();
            while (_continue)
            {
                var _attr = _params[_current.pos++];
                switch (_attr)
                {
                    case 0:
                    {
                        _imagePr.CanBeFlow = _params[_current.pos++];
                        break;
                    }
                    case 1:
                    {
                        _imagePr.Width = _params[_current.pos++];
                        break;
                    }
                    case 2:
                    {
                        _imagePr.Height = _params[_current.pos++];
                        break;
                    }
                    case 3:
                    {
                        _imagePr.WrappingStyle = _params[_current.pos++];
                        break;
                    }
                    case 4:
                    {
                        _imagePr.Paddings = AscCommon.asc_menu_ReadPaddings(_params, _current);
                        break;
                    }
                    case 5:
                    {
                        _imagePr.Position = asc_menu_ReadPosition(_params, _current);
                        break;
                    }
                    case 6:
                    {
                        _imagePr.AllowOverlap = _params[_current.pos++];
                        break;
                    }
                    case 7:
                    {
                        _imagePr.PositionH = asc_menu_ReadImagePosition(_params, _current);
                        break;
                    }
                    case 8:
                    {
                        _imagePr.PositionV = asc_menu_ReadImagePosition(_params, _current);
                        break;
                    }
                    case 9:
                    {
                        _imagePr.Internal_Position = _params[_current.pos++];
                        break;
                    }
                    case 10:
                    {
                        _imagePr.ImageUrl = _params[_current.pos++];
                        break;
                    }
                    case 11:
                    {
                        _imagePr.Locked = _params[_current.pos++];
                        break;
                    }
                    case 12:
                    {
                        _imagePr.ChartProperties = asc_menu_ReadChartPr(_params, _current);
                        break;
                    }
                    case 13:
                    {
                        _imagePr.ShapeProperties = asc_menu_ReadShapePr(_params, _current);
                        break;
                    }
                    case 14:
                    {
                        _imagePr.put_ChangeLevel(parseInt(_params[_current.pos++]));
                        break;
                    }
                    case 15:
                    {
                        _imagePr.Group = _params[_current.pos++];
                        break;
                    }
                    case 16:
                    {
                        _imagePr.fromGroup = _params[_current.pos++];
                        break;
                    }
                    case 17:
                    {
                        _imagePr.severalCharts = _params[_current.pos++];
                        break;
                    }
                    case 18:
                    {
                        _imagePr.severalChartTypes = _params[_current.pos++];
                        break;
                    }
                    case 19:
                    {
                        _imagePr.severalChartStyles = _params[_current.pos++];
                        break;
                    }
                    case 20:
                    {
                        _imagePr.verticalTextAlign = _params[_current.pos++];
                        break;
                    }
                    case 21:
                    {
                        var bIsNeed = _params[_current.pos++];

                        if (bIsNeed) {
                            var properties = this.WordControl.m_oLogicDocument.DrawingObjects.Get_Props();
                            if (properties) {
                                for (var i = 0; i < properties.length; i++) {
                                    if (undefined !== properties[i].ImageUrl && null != properties[i].ImageUrl) {                                
                                        var section_select = this.WordControl.m_oLogicDocument.Get_PageSizesByDrawingObjects();
                                        var page_width = AscCommon.Page_Width;
                                        var page_height = AscCommon.Page_Height;
                                        var page_x_left_margin = AscCommon.X_Left_Margin;
                                        var page_y_top_margin = AscCommon.Y_Top_Margin;
                                        var page_x_right_margin = AscCommon.X_Right_Margin;
                                        var page_y_bottom_margin = AscCommon.Y_Bottom_Margin;

                                        if (section_select) {
                                            if (section_select.W) {
                                                page_width = section_select.W;
                                            }

                                            if (section_select.H) {
                                                page_height = section_select.H;
                                            }
                                        }
                 
                                        var boundingWidth  = Math.max(1, page_width  - (page_x_left_margin + page_x_right_margin));
                                        var boundingHeight = Math.max(1, page_height - (page_y_top_margin  + page_y_bottom_margin));

                                        var size = this.WordControl.m_oDrawingDocument.Native["DD_GetOriginalImageSize"](properties[i].ImageUrl);

                                        var w = (undefined !== size[0]) ? Math.max(size[0] * AscCommon.g_dKoef_pix_to_mm, 1) : 1;
                                        var h = (undefined !== size[1]) ? Math.max(size[1] * AscCommon.g_dKoef_pix_to_mm, 1) : 1;
                                        
                                        var mW = boundingWidth  / w;
                                        var mH = boundingHeight / h;

                                        if (mH < mW) {
                                            boundingWidth  = boundingHeight / h * w;
                                        } else if (mW < mH) {
                                            boundingHeight = boundingWidth  / w * h;
                                        }

                                        //var __w = Math.max(1, page_width  - (page_x_left_margin + page_x_right_margin));
                                        //var __h = Math.max(1, page_height - (page_y_top_margin  + page_y_bottom_margin));

                                        //w = Math.max(5, boundingWidth);
                                        //h = Math.max(5, boundingHeight);

                                        _imagePr.Width  = Math.max(5, boundingWidth);
                                        _imagePr.Height = Math.max(5, boundingHeight);
                                        _imagePr.ImageUrl = undefined;

                                        break;
                                    }
                                }
                            }
                        }

                        break;
                    }
                    case 255:
                    default:
                    {
                        _continue = false;
                        break;
                    }
                }
            }

            this.ImgApply(_imagePr);
            break;
        }
        case 15: // ASC_MENU_EVENT_TYPE_TABLEMERGECELLS
        {
            this.MergeCells();
            break;
        }
        case 16: // ASC_MENU_EVENT_TYPE_TABLESPLITCELLS
        {
            this.SplitCell(_params[0], _params[1]);
            break;
        }
        case 51: // ASC_MENU_EVENT_TYPE_INSERT_TABLE
        {
            var _rows = 2;
            var _cols = 2;
            var _style = null;

            while (_continue)
            {
                var _attr = _params[_current.pos++];
                switch (_attr)
                {
                    case 0:
                    {
                        _rows = _params[_current.pos++];
                        break;
                    }
                    case 1:
                    {
                        _cols = _params[_current.pos++];
                        break;
                    }
                    case 2:
                    {
                        _style = _params[_current.pos++];
                        break;
                    }
                    case 255:
                    default:
                    {
                        _continue = false;
                        break;
                    }
                }
            }
            _style = _style + "";
            this.put_Table(_rows, _cols, _style);
            break;
        }
        case 50: // ASC_MENU_EVENT_TYPE_INSERT_IMAGE
        {
            var _src = _params[_current.pos++];
            var _w = _params[_current.pos++];
            var _h = _params[_current.pos++];
            var _pageNum = _params[_current.pos++];
            var _additionalParams = _params[_current.pos++];
            var _posX = _params[_current.pos++];
            var _posY = _params[_current.pos++];
            var _wrapType = _params[_current.pos++];

            this.AddImageUrlNative(_src, _w, _h, _pageNum);
            break;
        }
        case 401: // ASC_MENU_EVENT_TYPE_INSERT_SCREEN_IMAGE
        {
            var _src = _params[_current.pos++];
            var _w = _params[_current.pos++];
            var _h = _params[_current.pos++];
            var _pageNum = _params[_current.pos++];
            var _additionalParams = _params[_current.pos++];
            var _posX = _params[_current.pos++];
            var _posY = _params[_current.pos++];
            var _wrapType = _params[_current.pos++];

            this.AddImageUrlAtPosNative(_src, _w, _h, _pageNum, _posX, _posY, _wrapType);
            break;
        }
        case 53: // ASC_MENU_EVENT_TYPE_INSERT_SHAPE
        {
            var _shapeProp = asc_menu_ReadShapePr(_params, _current);
            var _pageNum = _shapeProp.InsertPageNum;
            this.WordControl.m_oLogicDocument.DrawingObjects.addShapeOnPage(_shapeProp.type, _shapeProp.InsertPageNum);
            break;
        }
        case 52: // ASC_MENU_EVENT_TYPE_INSERT_HYPERLINK
        {
            var _props = asc_menu_ReadHyperPr(_params, _current);
            if ( false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content) )
            {
                this.WordControl.m_oLogicDocument.StartAction();
                this.WordControl.m_oLogicDocument.AddHyperlink( _props );
                this.WordControl.m_oLogicDocument.FinalizeAction();
            }
            break;
        }
        case 8: // ASC_MENU_EVENT_TYPE_HYPERLINK
        {
            var _props = asc_menu_ReadHyperPr(_params, _current);
            var oHyperProps = null;
            for(var i = 0; i < this.SelectedObjectsStack.length; ++i)
            {
                if(this.SelectedObjectsStack[i].Type === Asc.c_oAscTypeSelectElement.Hyperlink)
                {
                    oHyperProps = this.SelectedObjectsStack[i].Value;
                    _props.Class = oHyperProps.Class;
                    break;
                }
            }
            if ( oHyperProps && false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content) )
            {
                this.WordControl.m_oLogicDocument.StartAction();
                this.WordControl.m_oLogicDocument.ModifyHyperlink( _props );
                this.WordControl.m_oLogicDocument.FinalizeAction();
            }
            break;
        }
        case 59: // ASC_MENU_EVENT_TYPE_REMOVE_HYPERLINK
        {

            var oHyperProps = null;
            for(var i = 0; i < this.SelectedObjectsStack.length; ++i)
            {
                if(this.SelectedObjectsStack[i].Type === Asc.c_oAscTypeSelectElement.Hyperlink)
                {
                    oHyperProps = this.SelectedObjectsStack[i].Value;
                    break;
                }
            }
            if (oHyperProps && false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content) )
            {
                this.WordControl.m_oLogicDocument.StartAction();
                this.WordControl.m_oLogicDocument.RemoveHyperlink(oHyperProps);
                this.WordControl.m_oLogicDocument.FinalizeAction();
            }
            break;
        }
        case 58: // ASC_MENU_EVENT_TYPE_CAN_ADD_HYPERLINK
        {
            var bCanAdd = this.WordControl.m_oLogicDocument.CanAddHyperlink(true);

            var _stream = global_memory_stream_menu;
            _stream["ClearNoAttack"]();
            if ( true === bCanAdd )
            {
                var _text = this.WordControl.m_oLogicDocument.GetSelectedText(true);
                if (null == _text)
                    _stream["WriteByte"](1);
                else
                {
                    _stream["WriteByte"](2);
                    _stream["WriteString2"](_text);
                }
            }
            else
            {
                _stream["WriteByte"](0);
            }
            _return = _stream;
            break;
        }
        case 62: // ASC_MENU_EVENT_TYPE_SEARCH_FINDTEXT
        {
            var searchSettings = new AscCommon.CSearchSettings();
            searchSettings.put_Text(_params[0]);
            searchSettings.put_MatchCase(_params[2]);

            var _ret = _api.asc_findText(searchSettings, _params[1]);
            var _stream = global_memory_stream_menu;
            _stream["ClearNoAttack"]();
            _stream["WriteLong"](_ret);
            _return = _stream;
            break;
        }
        case 63: // ASC_MENU_EVENT_TYPE_SEARCH_REPLACETEXT
        {
            var searchSettings = new AscCommon.CSearchSettings();
            searchSettings.put_Text(_params[0]);
            searchSettings.put_MatchCase(_params[3]);

            var _ret = _api.asc_replaceText(searchSettings, _params[1], _params[2]);
            var _stream = global_memory_stream_menu;
            _stream["ClearNoAttack"]();
            _stream["WriteBool"](_ret);
            _return = _stream;
            break;
        }
        case 64: // ASC_MENU_EVENT_TYPE_SEARCH_SELECTRESULTS
        {
            this.asc_selectSearchingResults(_params[0]);
            break;
        }
        case 65: // ASC_MENU_EVENT_TYPE_SEARCH_ISSELECTRESULTS
        {
            var _ret = this.asc_isSelectSearchingResults();
            var _stream = global_memory_stream_menu;
            _stream["ClearNoAttack"]();
            _stream["WriteBool"](_ret);
            _return = _stream;
            break;
        }
        case 67: // ASC_MENU_EVENT_TYPE_STATISTIC_START
        {
            _api.startGetDocInfo();
            break;
        }
        case 68: // ASC_MENU_EVENT_TYPE_STATISTIC_STOP
        {
            _api.stopGetDocInfo();
            break;
        }
        case 17:
        {
            var _sect_width = undefined;
            var _sect_height = undefined;
            var _sect_orient = undefined;

            while (_continue)
            {
                var _attr = _params[_current.pos++];
                switch (_attr)
                {
                    case 0:
                    {
                        _sect_width = _params[_current.pos++];
                        break;
                    }
                    case 1:
                    {
                        _sect_height = _params[_current.pos++];
                        break;
                    }
                    case 2:
                    {
                        // margin_left
                        _current.pos++;
                        break;
                    }
                    case 3:
                    {
                        // margin_top
                        _current.pos++;
                        break;
                    }
                    case 4:
                    {
                        // margin_right
                        _current.pos++;
                        break;
                    }
                    case 5:
                    {
                        // margin_bottom
                        _current.pos++;
                        break;
                    }
                    case 6:
                    {
                        _sect_orient = _params[_current.pos++];
                        break;
                    }
                    case 255:
                    default:
                    {
                        _continue = false;
                        break;
                    }
                }
            }

            if (undefined !== _sect_width && undefined !== _sect_height)
                this.change_DocSize(_sect_width, _sect_height);

            if (undefined !== _sect_orient)
                this.change_PageOrient(_sect_orient);

            break;
        }
        case 200: // ASC_MENU_EVENT_TYPE_DOCUMENT_BASE64
        {
            var _stream = global_memory_stream_menu;
            _stream["ClearNoAttack"]();
            _stream["WriteStringA"](this["asc_nativeGetFileData"]());
            _return = _stream;
            break;
        }
        case 202: // ASC_MENU_EVENT_TYPE_DOCUMENT_PDFBASE64
        case 203: // ASC_MENU_EVENT_TYPE_DOCUMENT_PDFBASE64_PRINT
        {
            var _stream = global_memory_stream_menu;
            _stream["ClearNoAttack"]();
            _stream["WriteStringA"](this.WordControl.m_oDrawingDocument.ToRenderer(203 === type));
            _return = _stream;
            break;
        }
        case 110: // ASC_MENU_EVENT_TYPE_CONTEXTMENU_COPY
        {
            _return = this.Call_Menu_Context_Copy();
            break;
        }
        case 111 : // ASC_MENU_EVENT_TYPE_CONTEXTMENU_CUT
        {
            _return = this.Call_Menu_Context_Cut();
            break;
        }
        case 112: // ASC_MENU_EVENT_TYPE_CONTEXTMENU_PASTE
        {
            this.Call_Menu_Context_Paste(_params[0], _params[1]);
            break;
        }
        case 113: // ASC_MENU_EVENT_TYPE_CONTEXTMENU_DELETE
        {
            this.Call_Menu_Context_Delete();
            break;
        }
        case 114: // ASC_MENU_EVENT_TYPE_CONTEXTMENU_SELECT
        {
            this.Call_Menu_Context_Select();
            break;
        }
        case 115: // ASC_MENU_EVENT_TYPE_CONTEXTMENU_SELECTALL
        {
            this.Call_Menu_Context_SelectAll();
            break;
        }
        case 201: // ASC_MENU_EVENT_TYPE_DOCUMENT_CHARTSTYLES
        {
            var _typeChart = _params[0];
            this.chartPreviewManager.getChartPreviews(_typeChart);
            _return = global_memory_stream_menu;
            break;
        }
        case 71: // ASC_MENU_EVENT_TYPE_TABLE_INSERTDELETE_ROWCOLUMN
        {
            if (typeof _params[0] === 'string') {
                var json = JSON.parse(_params[0]);
                if (json) {
                    var isInsert = json["insert"] || false;
                    var isDelete = json["delete"] || false;
                    var type = json["type"] || "table";

                    if (isInsert) {
                        if (type === "row") {
                            json["above"] ? _api.addRowAbove() : _api.addRowBelow();
                        } else if (type == "column") {
                            json["left"] ? _api.addColumnLeft() : _api.addColumnRight();
                        }
                    } else if (isDelete) {
                        if (type === "row") {
                            _api.remRow();
                        } else if (type === "column") {
                            _api.remColumn();
                        } else {
                            _api.remTable();
                        }
                    }
                }
            } else {
                var _type = 0;
                var _is_add = true;
                var _is_above = true;
                while (_continue) {
                    var _attr = _params[_current.pos++];
                    switch (_attr) {
                        case 0:
                            {
                                _type = _params[_current.pos++];
                                break;
                            }
                        case 1:
                            {
                                _is_add = _params[_current.pos++];
                                break;
                            }
                        case 2:
                            {
                                _is_above = _params[_current.pos++];
                                break;
                            }
                        case 255:
                        default:
                            {
                                _continue = false;
                                break;
                            }
                    }
                }

                if (1 == _type) {
                    if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Table_Properties)) {
                        this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_TableAddColumnLeft);
                        if (_is_add)
                            this.WordControl.m_oLogicDocument.AddTableColumn(!_is_above);
                        else
                            this.WordControl.m_oLogicDocument.RemoveTableColumn();

                        this.WordControl.m_oLogicDocument.FinalizeAction();
                    }
                }
                else if (2 == _type) {
                    if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Table_Properties)) {
                        this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_TableAddColumnLeft);
                        if (_is_add)
                            this.WordControl.m_oLogicDocument.AddTableRow(!_is_above);
                        else
                            this.WordControl.m_oLogicDocument.RemoveTableRow();

                        this.WordControl.m_oLogicDocument.FinalizeAction();
                    }
                }

            }
            break;
        }

        case 440:   // ASC_MENU_EVENT_TYPE_ADD_CHART_DATA
        {
            if (undefined !== _params) {
                var chartData = _params[0];
                if (chartData && chartData.length > 0) {
                    var json = JSON.parse(chartData);
                    if (json) {
                        _api.asc_addChartDrawingObject(json);
                    }
                }
            }
            break;
        } 
            
        case 450:   // ASC_MENU_EVENT_TYPE_GET_CHART_DATA
        {
            var index = null;
            if (undefined !== _params) {
                index = parseInt(_params);
            }

            var chart = _api.asc_getChartObject(index);
            
            var _stream = global_memory_stream_menu;
            _stream["ClearNoAttack"]();
            _stream["WriteStringA"](JSON.stringify(new Asc.asc_CChartBinary(chart)));
            _return = _stream;
            break;
        }
        
        case 460:   // ASC_MENU_EVENT_TYPE_SET_CHART_DATA
        {
            if (undefined !== _params) {
                var chartData = _params[0];
                if (chartData && chartData.length > 0) {
                    var json = JSON.parse(chartData);
                    if (json) {
                        _api.asc_editChartDrawingObject(json);
                    }
                }
            }
            break;
        }
            
        case 2415: // ASC_MENU_EVENT_TYPE_CHANGE_COLOR_SCHEME
        {
            if (undefined !== _params) {
                var indexScheme = parseInt(_params);
                _api.asc_ChangeColorSchemeByIdx(indexScheme);
            }
            break;
        }

        case 2416: // ASC_MENU_EVENT_TYPE_GET_COLOR_SCHEME
        {
            var index = _api.asc_GetCurrentColorSchemeIndex();
            var stream = global_memory_stream_menu;
            stream["ClearNoAttack"]();
            stream["WriteLong"](index);
            _return = stream;
            break;
        }

        case 10000: // ASC_SOCKET_EVENT_TYPE_OPEN
        {
            _api.CoAuthoringApi._CoAuthoringApi.socketio.onMessage("connect");
            break;
        }

        case 10010: // ASC_SOCKET_EVENT_TYPE_ON_CLOSE
        {
            // NOT USED
            break;
        }

        case 10020: // ASC_SOCKET_EVENT_TYPE_MESSAGE
        {
            _api.CoAuthoringApi._CoAuthoringApi.socketio.onMessage("message", _params ? JSON.parse(_params) : {});
            break;
        }

        case 11010: // ASC_SOCKET_EVENT_TYPE_ON_DISCONNECT
        {
            _api.CoAuthoringApi._CoAuthoringApi.socketio.onMessage("disconnect", _params || "");
            break;
        }

        case 11020: // ASC_SOCKET_EVENT_TYPE_TRY_RECONNECT
        {
            // NOT USED
            break;
        }

        case 21000: // ASC_COAUTH_EVENT_TYPE_INSERT_URL_IMAGE
        {
            var urls = JSON.parse(_params[0]);
            AscCommon.g_oDocumentUrls.addUrls(urls);
            var firstUrl;
            for (var i in urls) {
                if (urls.hasOwnProperty(i)) {
                    firstUrl = urls[i];
                    break;
                }
            }

            var _src = firstUrl;
            var _w = _params[1];
            var _h = _params[2];
            var _pageNum = _params[3];

            this.AddImageUrlActionNative(_src, _w, _h, _pageNum);
            break;
        }

        case 21001: // ASC_COAUTH_EVENT_TYPE_LOAD_URL_IMAGE
        {
          _api.WordControl.m_oDrawingDocument.ClearCachePages();
          _api.WordControl.m_oDrawingDocument.FirePaint();

          break;
        }

        case 21002: // ASC_COAUTH_EVENT_TYPE_REPLACE_URL_IMAGE
        {
            var urls = JSON.parse(_params[0]);
            AscCommon.g_oDocumentUrls.addUrls(urls);
            var firstUrl;
            for (var i in urls) {
                if (urls.hasOwnProperty(i)) {
                    firstUrl = urls[i];
                    break;
                }
            }

            var _src = firstUrl;

            var imageProp = new asc_CImgProperty();
            imageProp.ImageUrl = _src;
            this.ImgApply(imageProp);

            break;
        }

        case 21003: // ASC_COAUTH_EVENT_TYPE_INSERT_SCREEN_URL_IMAGE
        {
            var urls = JSON.parse(_params[_current.pos++]);
            AscCommon.g_oDocumentUrls.addUrls(urls);
            var firstUrl;
            for (var i in urls) {
                if (urls.hasOwnProperty(i)) {
                    firstUrl = urls[i];
                    break;
                }
            }

            var _src = firstUrl;
            var _w = _params[_current.pos++];
            var _h = _params[_current.pos++];
            var _pageNum = _params[_current.pos++];
            var _additionalParams = _params[_current.pos++];
            var _posX = _params[_current.pos++];
            var _posY = _params[_current.pos++];
            var _wrapType = _params[_current.pos++];

            this.AddImageUrlAtPosNative(_src, _w, _h, _pageNum, _posX, _posY, _wrapType);
            break;
        }

        case 22001: // ASC_MENU_EVENT_TYPE_SET_PASSWORD
        {
          _api.asc_setDocumentPassword(_params[0]);
          break;
        }

        case 22000: // ASC_MENU_EVENT_TYPE_ADVANCED_OPTIONS
        {
            var obj = JSON.parse(_params);
            var type = parseInt(obj["type"]);
            var encoding = parseInt(obj["encoding"]);

            _api.advancedOptionsAction = AscCommon.c_oAscAdvancedOptionsAction.Open;
            _api.documentFormat = "txt";
           
            _api.asc_setAdvancedOptions(type, new Asc.asc_CTextOptions(encoding));
            
            break;
        } 

        case 22004: // ASC_EVENT_TYPE_SPELLCHECK_MESSAGE
        {
            var spellData = JSON.parse(_params[0]);
            if (_api.SpellCheckApi && spellData)
                _api.SpellCheckApi.onSpellCheck(spellData);
            break;
        }

        case 22005: // ASC_EVENT_TYPE_SPELLCHECK_TURN_ON
        {
            var status = parseInt(_params[0]);
            if (status !== undefined) {
                this.asc_setSpellCheck(status == 0 ? false : true);
            }
            break;
        }

        case 22006: // ASC_EVENT_TYPE_DO_NONPRINTING_DISPLAY
        {
            var json = JSON.parse(_params[0]);
            if (json) {
                var display = json["display"] || false;
                if (_api.put_ShowParaMarks) {
                    _api.put_ShowParaMarks(display);
                }
                if (_api.put_ShowTableEmptyLine) {
                    _api.put_ShowTableEmptyLine(display);
                }
                _api.WordControl.m_oDrawingDocument.ClearCachePages();
                _api.WordControl.m_oDrawingDocument.FirePaint();
            }
            break;
        }

        case 23101: // ASC_MENU_EVENT_TYPE_DO_SELECT_COMMENT
        {
            var json = JSON.parse(_params[0]);
            if (json && json["id"]) {
                var id = parseInt(json["id"]);
                if (_api.asc_selectComment && id) {
                    _api.asc_selectComment(id);
                }
            }
            break;
        }

        case 23102: // ASC_MENU_EVENT_TYPE_DO_SHOW_COMMENT
        {
            var json = JSON.parse(_params[0]);
            if (json && json["id"]) {
                if (_api.asc_showComment) {
                    _api.asc_showComment(json["id"], json["isNew"] === true);
                }
            }
            break;
        }

        case 23103: // ASC_MENU_EVENT_TYPE_DO_SELECT_COMMENTS
        {
            var json = JSON.parse(_params[0]);
            if (json) {
                if (_api.asc_showComments) {
                    _api.asc_showComments(json["resolved"] === true);
                }
            }
            break;
        }

        case 23104: // ASC_MENU_EVENT_TYPE_DO_DESELECT_COMMENTS
        {
            if (_api.asc_hideComments) {
                _api.asc_hideComments();
            }
            break;
        }

        case 23105: // ASC_MENU_EVENT_TYPE_DO_ADD_COMMENT
        {
            var json = JSON.parse(_params[0]);
            if (json) {
                var buildCommentData = function () {
                    if (typeof Asc.asc_CCommentDataWord !== 'undefined') {
                        return new Asc.asc_CCommentDataWord(null);
                    }
                    return new Asc.asc_CCommentData(null);
                };

                var comment = buildCommentData();
                var now = new Date();
                var timeZoneOffsetInMs = (new Date()).getTimezoneOffset() * 60000;
                var currentUserId = _internalStorage.externalUserInfo.asc_getId();
                var currentUserName = _internalStorage.externalUserInfo.asc_getFullName();

                if (comment) {
                    comment.asc_putText(json["text"]);
                    comment.asc_putTime((now.getTime() - timeZoneOffsetInMs).toString());
                    comment.asc_putOnlyOfficeTime(now.getTime().toString());
                    comment.asc_putUserId(currentUserId);
                    comment.asc_putUserName(currentUserName);
                    comment.asc_putSolved(false);

                    if (comment.asc_putDocumentFlag) {
                        comment.asc_putDocumentFlag(json["unattached"]);
                    }

                    _api.asc_addComment(comment);
                }
            }
            break;
        }

        case 23106: // ASC_MENU_EVENT_TYPE_DO_REMOVE_COMMENT
        {
            var json = JSON.parse(_params[0]);
            if (json && json["id"]) { // id - String
                if (_api.asc_removeComment) {
                    _api.asc_removeComment(json["id"]);
                }
            }
            break;
        }

        case 23107: // ASC_MENU_EVENT_TYPE_DO_REMOVE_ALL_COMMENTS
        {
            var json = JSON.parse(_params[0]),
                type = json["type"],
                canEditComments = json["canEditComments"];
            if (json && type) {
                if (_api.asc_RemoveAllComments) {
                    _api.asc_RemoveAllComments(type=='my' || !(canEditComments === true), type=='current'); // 1 param = true if remove only my comments, 2 param - remove current comments
                }
            }
            break;
        }

        case 23108: // ASC_MENU_EVENT_TYPE_DO_CHANGE_COMMENT
        {
            var json = JSON.parse(_params[0]),
                commentId = json["id"],
                comment = json["comment"],
                updateAuthor = json["updateAuthor"] || false;

            if (json && commentId) {
                var timeZoneOffsetInMs = (new Date()).getTimezoneOffset() * 60000;
                var currentUserId = _internalStorage.externalUserInfo.asc_getId();
                var currentUserName = _internalStorage.externalUserInfo.asc_getFullName();
                var buildCommentData = function () {
                    if (typeof Asc.asc_CCommentDataWord !== 'undefined') {
                        return new Asc.asc_CCommentDataWord(null);
                    }
                    return new Asc.asc_CCommentData(null);
                };
                var ooDateToString = function (date) {
                    if (Object.prototype.toString.call(date) === '[object Date]')
                        return (date.getTime()).toString();
                    return "";
                };
                var utcDateToString = function (date) {
                    if (Object.prototype.toString.call(date) === '[object Date]')
                        return (date.getTime() - timeZoneOffsetInMs).toString();
                    return "";
                };
                var ascComment = buildCommentData();

                if (ascComment && comment && _api.asc_changeComment) {
                    var sTime = new Date(parseInt(comment["date"]));
                    ascComment.asc_putText(comment["text"]);
                    ascComment.asc_putQuoteText(comment["quoteText"]);
                    ascComment.asc_putTime(utcDateToString(sTime));
                    ascComment.asc_putOnlyOfficeTime(ooDateToString(sTime));
                    ascComment.asc_putUserId(updateAuthor ? currentUserId : comment["userId"]);
                    ascComment.asc_putUserName(updateAuthor ? currentUserName : comment["userName"]);
                    ascComment.asc_putSolved(comment["solved"]);
                    ascComment.asc_putGuid(comment["id"]);

                    if (ascComment.asc_putDocumentFlag !== undefined) {
                        ascComment.asc_putDocumentFlag(comment["unattached"]);
                    }

                    var replies = comment["replies"];

                    if (replies && replies.length) {
                        replies.forEach(function (reply) {
                            var addReply = buildCommentData();   //  new asc_CCommentData(null);
                            if (addReply) {
                                var sTime = new Date(parseInt(reply["date"]));
                                addReply.asc_putText(reply["text"]);
                                addReply.asc_putTime(utcDateToString(sTime));
                                addReply.asc_putOnlyOfficeTime(ooDateToString(sTime));
                                addReply.asc_putUserId(reply["userId"]);
                                addReply.asc_putUserName(reply["userName"]);

                                ascComment.asc_addReply(addReply);
                            }
                        });
                    }

                    _api.asc_changeComment(commentId, ascComment);
                }
            }
            break;
        }

        case 23109: // ASC_MENU_EVENT_TYPE_DO_CAN_ADD_QUOTED_COMMENT
        {
            var _stream = global_memory_stream_menu;
            _stream["ClearNoAttack"]();
            _stream["WriteString2"](JSON.stringify({
                result: this.can_AddQuotedComment() !== false
            }));
            _return = _stream;
            break;
        }

        case 24101: // ASC_MENU_EVENT_TYPE_DO_SET_TRACK_REVISIONS
        {
            var json = JSON.parse(_params[0]);
            if (json) {
                if (_api.asc_SetTrackRevisions) {
                    _api.asc_SetTrackRevisions(json["state"] === true);
                }
            }
            break;
        }
        case 24102: // ASC_MENU_EVENT_TYPE_DO_BEGIN_VIEWMODE_IN_REVIEW
        {
            var json = JSON.parse(_params[0]);
            if (json) {
                if (_api.asc_BeginViewModeInReview) {
                    _api.asc_BeginViewModeInReview(json["review"] === true);
                }
            }
            break;
        }
        case 24103: // ASC_MENU_EVENT_TYPE_DO_END_VIEWMODE_IN_REVIEW
        {
            if (_api.asc_EndViewModeInReview) {
                _api.asc_EndViewModeInReview();
            }
            break;
        }
        case 24104: // ASC_MENU_EVENT_TYPE_DO_ACCEPT_ALL_CHANGES
        {
            if (_api.asc_AcceptAllChanges) {
                _api.asc_AcceptAllChanges();
            }
            break;
        }
        case 24105: // ASC_MENU_EVENT_TYPE_DO_REJECT_ALL_CHANGES
        {
            if (_api.asc_RejectAllChanges) {
                _api.asc_RejectAllChanges();
            }
            break;
        }
        case 24106: // ASC_MENU_EVENT_TYPE_DO_GET_PREV_REVISIONS_CHANGE
        {
            if (_api.asc_GetPrevRevisionsChange) {
                _api.asc_GetPrevRevisionsChange();
            }
            break;
        }
        case 24107: // ASC_MENU_EVENT_TYPE_DO_GET_NEXT_REVISIONSCHANGE
        {
            if (_api.asc_GetNextRevisionsChange) {
                _api.asc_GetNextRevisionsChange();
            }
            break;
        }
        case 24108: // ASC_MENU_EVENT_TYPE_DO_ACCEPT_CHANGES
        {
            var api = _api;
            if (api.asc_AcceptChanges) {
                api.asc_AcceptChanges(_internalStorage.changesReview[0]);

                if (api.asc_GetNextRevisionsChange) {
                    setTimeout(function () {
                        api.asc_GetNextRevisionsChange();
                    }, 10);
                }
            }
            break;
        }
        case 24109: // ASC_MENU_EVENT_TYPE_DO_REJECT_CHANGES
        {
            var api = _api;
            if (api.asc_RejectChanges) {
                api.asc_RejectChanges(_internalStorage.changesReview[0]);
                if (api.asc_GetNextRevisionsChange) {
                    setTimeout(function () {
                        api.asc_GetNextRevisionsChange();
                    }, 10);
                }
            }
            break;
        }
        case 24110: // ASC_MENU_EVENT_TYPE_DO_FOLLOW_REVISION_MOVE
        {
            if (_api.asc_FollowRevisionMove) {
                _api.asc_FollowRevisionMove(_internalStorage.changesReview[0]);
            }
            break;
        }

        case 25001: // ASC_MENU_EVENT_TYPE_DO_API_FUNCTION_CALL
        {
            var json = JSON.parse(_params[0]),
                func = json["func"],
                params = json["params"] || [],
                returnable = json["returnable"] || false; // need return result

            if (json && func) {
                if (_api[func]) {
                    if (returnable) {
                        var _stream = global_memory_stream_menu;
                        _stream["ClearNoAttack"]();
                        var result = _api[func].apply(_api, params);
                        _stream["WriteString2"](JSON.stringify({
                            result: result
                        }));
                        _return = _stream;
                    } else {
                        _api[func].apply(_api, params);
                    }
                }
            }
            break;
        }

        case 26003: // ASC_MENU_EVENT_TYPE_DO_SET_CONTENTCONTROL_PICTURE
        {
            var _src = _params[_current.pos++];
            var _w = _params[_current.pos++];
            var _h = _params[_current.pos++];
            var _pageNum = _params[_current.pos++];
            var _additionalParams = _params[_current.pos++];

            var json = JSON.parse(_additionalParams);
            if (json) {
                var internalId = json["internalId"] || "";
                _api.SetContentControlPictureUrlNative(_src, internalId)
            }
            break;
        }

        case 26004: // ASC_MENU_EVENT_TYPE_DO_SET_CONTENTCONTROL_PICTURE_URL
        {
            var urls = JSON.parse(_params[0]);
            AscCommon.g_oDocumentUrls.addUrls(urls);
            var firstUrl;
            for (var i in urls) {
                if (urls.hasOwnProperty(i)) {
                    firstUrl = urls[i];
                    break;
                }
            }

            var _src = firstUrl;
            var _w = _params[1];
            var _h = _params[2];
            var _pageNum = _params[3];
            var _additionalParams = _params[4];

            var json = JSON.parse(_additionalParams);
            if (json) {
                var internalId = json["internalId"] || "";
                _api.SetContentControlPictureUrlNative(_src, internalId);
            }
            break;
        }

        case 27001: //ASC_MENU_EVENT_TYPE_SET_FOOTNOTE_PROP
        {
            var json = JSON.parse(_params),
                pos = json["Pos"],
                numFormat = json["NumFormat"],
                numStart = json["NumStart"],
                numRestart = json["NumRestart"],
                isAll = json["IsAll"] || true,
                isEndnote = json["IsEndnote"] || false;

            var props = new Asc.CAscFootnotePr();
            props.put_Pos(pos);
            props.put_NumFormat(numFormat);
            props.put_NumStart(numStart);
            props.put_NumRestart(numRestart);

            if (isEndnote) {
                _api.asc_SetEndnoteProps(props, isAll);
            } else {
                _api.asc_SetFootnoteProps(props, isAll);
            }
            break;
        }
        case 2500: // ASC_MENU_EVENT_TYPE_CHANGE_MOBILE_MODE
        {
            _api.ChangeReaderMode();
            break;
        }

        default:
            break;
    }
    return _return;
};

// STYLES
Asc['asc_docs_api'].prototype.GenerateNativeStyles = function()
{
    var StylesPainter = new CStylesPainter();
    StylesPainter.GenerateStyles(this, this.LoadedObjectDS);
};

Asc['asc_docs_api'].prototype.asc_setDocumentPassword = function(password)
{
    var v = {
      "id": this.documentId,
      "userid": this.documentUserId,
      "format": this.documentFormat,
      "c": "reopen",
      "title": this.documentTitle,
      "password": password
    };

    AscCommon.sendCommand(this, null, v);
};

Asc['asc_docs_api'].prototype.Update_ParaInd = function( Ind )
{
    this.WordControl.m_oDrawingDocument.Update_ParaInd(Ind);
};

Asc['asc_docs_api'].prototype.Internal_Update_Ind_Left = function(Left)
{
};

Asc['asc_docs_api'].prototype.Internal_Update_Ind_Right = function(Right)
{
};

Asc['asc_docs_api'].prototype.UpdateTextPr = function(TextPr)
{
    if (!TextPr)
        return;

    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();

    if (TextPr.Bold !== undefined)
    {
        _stream["WriteByte"](0);
        _stream["WriteBool"](TextPr.Bold);
    }
    if (TextPr.Italic !== undefined)
    {
        _stream["WriteByte"](1);
        _stream["WriteBool"](TextPr.Italic);
    }
    if (TextPr.Underline !== undefined)
    {
        _stream["WriteByte"](2);
        _stream["WriteBool"](TextPr.Underline);
    }
    if (TextPr.Strikeout !== undefined)
    {
        _stream["WriteByte"](3);
        _stream["WriteBool"](TextPr.Strikeout);
    }

    asc_menu_WriteFontFamily(4, TextPr.FontFamily, _stream);

    if (TextPr.FontSize !== undefined)
    {
        _stream["WriteByte"](5);
        _stream["WriteDouble2"](TextPr.FontSize);
    }

    if(TextPr.Unifill && TextPr.Unifill.fill && TextPr.Unifill.fill.type === Asc.c_oAscFill.FILL_TYPE_SOLID && TextPr.Unifill.fill.color)
    {
        var _color = AscCommon.CreateAscColor(TextPr.Unifill.fill.color);
        asc_menu_WriteColor(6, AscCommon.CreateAscColorCustom(_color.r, _color.g, _color.b, false), _stream);
    }
    else if (TextPr.Color !== undefined)
    {
        asc_menu_WriteColor(6, AscCommon.CreateAscColorCustom(TextPr.Color.r, TextPr.Color.g, TextPr.Color.b, TextPr.Color.Auto), _stream);
    }

    if (TextPr.VertAlign !== undefined)
    {
        _stream["WriteByte"](7);
        _stream["WriteLong"](TextPr.VertAlign);
    }

    if (TextPr.HighLight !== undefined)
    {
        if (TextPr.HighLight === AscCommonWord.highlight_None)
        {
            _stream["WriteByte"](12);
        }
        else
        {
            asc_menu_WriteColor(8, AscCommon.CreateAscColorCustom(TextPr.HighLight.r, TextPr.HighLight.g, TextPr.HighLight.b), _stream);
        }
    }

    if (TextPr.DStrikeout !== undefined)
    {
        _stream["WriteByte"](9);
        _stream["WriteBool"](TextPr.DStrikeout);
    }
    if (TextPr.Caps !== undefined)
    {
        _stream["WriteByte"](10);
        _stream["WriteBool"](TextPr.Caps);
    }
    if (TextPr.SmallCaps !== undefined)
    {
        _stream["WriteByte"](11);
        _stream["WriteBool"](TextPr.SmallCaps);
    }
    if (TextPr.Spacing !== undefined)
    {
        _stream["WriteByte"](13);
        _stream["WriteDouble2"](TextPr.Spacing);
    }

    _stream["WriteByte"](255);

    this.Send_Menu_Event(1);
};

Asc['asc_docs_api'].prototype.UpdateParagraphProp = function(ParaPr)
{
    // TODO: как только разъединят настройки параграфа и текста переделать тут
    var TextPr = this.WordControl.m_oLogicDocument.GetCalculatedTextPr();
    ParaPr.Subscript   = ( TextPr.VertAlign === AscCommon.vertalign_SubScript   ? true : false );
    ParaPr.Superscript = ( TextPr.VertAlign === AscCommon.vertalign_SuperScript ? true : false );
    ParaPr.Strikeout   = TextPr.Strikeout;
    ParaPr.DStrikeout  = TextPr.DStrikeout;
    ParaPr.AllCaps     = TextPr.Caps;
    ParaPr.SmallCaps   = TextPr.SmallCaps;
    ParaPr.TextSpacing = TextPr.Spacing;
    ParaPr.Position    = TextPr.Position;
    //-----------------------------------------------------------------------------

    if ( -1 === ParaPr.PStyle )
        ParaPr.StyleName = "";
    else if ( undefined === ParaPr.PStyle )
        ParaPr.StyleName = this.WordControl.m_oLogicDocument.Styles.Style[this.WordControl.m_oLogicDocument.Styles.Get_Default_Paragraph()].Name;
    else
        ParaPr.StyleName = this.WordControl.m_oLogicDocument.Styles.Style[ParaPr.PStyle].Name;

    var NumType    = -1;
    var NumSubType = -1;
    if ( !(null == ParaPr.NumPr || 0 === ParaPr.NumPr.NumId || "0" === ParaPr.NumPr.NumId) )
    {
        var oNum = this.WordControl.m_oLogicDocument.GetNumbering().GetNum(ParaPr.NumPr.NumId);

        if (oNum && oNum.GetLvl(ParaPr.NumPr.Lvl))
        {
            var Lvl = oNum.GetLvl(ParaPr.NumPr.Lvl);
            var NumFormat = Lvl.GetFormat();
            var NumText   = Lvl.GetLvlText();

            if ( Asc.c_oAscNumberingFormat.Bullet === NumFormat )
            {
                NumType    = 0;
                NumSubType = 0;

                var TextLen = NumText.length;
                if ( 1 === TextLen && numbering_lvltext_Text === NumText[0].Type )
                {
                    var NumVal = NumText[0].Value.charCodeAt(0);

                    if ( 0x00B7 === NumVal )
                        NumSubType = 1;
                    else if ( 0x006F === NumVal )
                        NumSubType = 2;
                    else if ( 0x00A7 === NumVal )
                        NumSubType = 3;
                    else if ( 0x0076 === NumVal )
                        NumSubType = 4;
                    else if ( 0x00D8 === NumVal )
                        NumSubType = 5;
                    else if ( 0x00FC === NumVal )
                        NumSubType = 6;
                    else if ( 0x00A8 === NumVal )
                        NumSubType = 7;
                }
            }
            else
            {
                NumType    = 1;
                NumSubType = 0;

                var TextLen = NumText.length;
                if ( 2 === TextLen && numbering_lvltext_Num === NumText[0].Type && numbering_lvltext_Text === NumText[1].Type )
                {
                    var NumVal2 = NumText[1].Value;

                    if ( Asc.c_oAscNumberingFormat.Decimal === NumFormat )
                    {
                        if ( "." === NumVal2 )
                            NumSubType = 1;
                        else if ( ")" === NumVal2 )
                            NumSubType = 2;
                    }
                    else if ( Asc.c_oAscNumberingFormat.UpperRoman === NumFormat )
                    {
                        if ( "." === NumVal2 )
                            NumSubType = 3;
                    }
                    else if ( Asc.c_oAscNumberingFormat.UpperLetter === NumFormat )
                    {
                        if ( "." === NumVal2 )
                            NumSubType = 4;
                    }
                    else if ( Asc.c_oAscNumberingFormat.LowerLetter === NumFormat )
                    {
                        if ( ")" === NumVal2 )
                            NumSubType = 5;
                        else if ( "." === NumVal2 )
                            NumSubType = 6;
                    }
                    else if ( Asc.c_oAscNumberingFormat.LowerRoman === NumFormat )
                    {
                        if ( "." === NumVal2 )
                            NumSubType = 7;
                    }
                }
            }
        }
    }

    ParaPr.ListType = { Type : NumType, SubType : NumSubType };

    if ( undefined !== ParaPr.FramePr && undefined !== ParaPr.FramePr.Wrap )
    {
        if ( wrap_NotBeside === ParaPr.FramePr.Wrap )
            ParaPr.FramePr.Wrap = false;
        else if ( wrap_Around === ParaPr.FramePr.Wrap )
            ParaPr.FramePr.Wrap = true;
        else
            ParaPr.FramePr.Wrap = undefined;
    }

    var _len = this.SelectedObjectsStack.length;
    if (_len > 0)
    {
        if (this.SelectedObjectsStack[_len - 1].Type == Asc.c_oAscTypeSelectElement.Paragraph)
        {
            this.SelectedObjectsStack[_len - 1].Value = ParaPr;
            return;
        }
    }

    this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new AscCommon.asc_CSelectedObject ( Asc.c_oAscTypeSelectElement.Paragraph, ParaPr );
};

Asc['asc_docs_api'].prototype.put_PageNum = function(where,align)
{
    if ( where >= 0 )
    {
        if ( false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, { Type : AscCommon.changestype_2_HdrFtr }) )
        {
            this.WordControl.m_oLogicDocument.StartAction();
            this.WordControl.m_oLogicDocument.Document_AddPageNum( where, align );
			this.WordControl.m_oLogicDocument.FinalizeAction();
        }
    }
    else
    {
        if ( false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content) )
        {
            this.WordControl.m_oLogicDocument.StartAction();
            this.WordControl.m_oLogicDocument.Document_AddPageNum( where, align );
			this.WordControl.m_oLogicDocument.FinalizeAction();
        }
    }
};

Asc['asc_docs_api'].prototype.put_AddPageBreak = function()
{
    if ( false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content) )
    {
        var Document = this.WordControl.m_oLogicDocument;

        if ( null === Document.IsCursorInHyperlink(false) )
        {
            Document.StartAction();
            Document.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Page));
			Document.FinalizeAction();
        }
    }
};

Asc['asc_docs_api'].prototype.add_SectionBreak = function(_Type)
{
    if ( false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content) )
    {
        this.WordControl.m_oLogicDocument.StartAction();
        this.WordControl.m_oLogicDocument.Add_SectionBreak(_Type);
		this.WordControl.m_oLogicDocument.FinalizeAction();
    }
};

Asc['asc_docs_api'].prototype.put_AddLineBreak = function()
{
    if ( false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content) )
    {
        var Document = this.WordControl.m_oLogicDocument;

        if ( null === Document.IsCursorInHyperlink(false) )
        {
            Document.StartAction();
            Document.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Line));
			Document.FinalizeAction();
        }
    }
};

Asc['asc_docs_api'].prototype.ImgApply = function(obj)
{

    var ImagePr = obj;

    // Если у нас меняется с Float->Inline мы также должны залочить соответствующий параграф
    var AdditionalData = null;
    var LogicDocument = this.WordControl.m_oLogicDocument;
    if(obj && obj.ChartProperties && obj.ChartProperties.type === Asc.c_oAscChartTypeSettings.stock)
    {
        var selectedObjectsByType = LogicDocument.DrawingObjects.getSelectedObjectsByTypes();
        if(selectedObjectsByType.charts[0])
        {
            var chartSpace = selectedObjectsByType.charts[0];
            if(chartSpace && chartSpace.chart && chartSpace.chart.plotArea && chartSpace.chart.plotArea.charts[0] && chartSpace.chart.plotArea.charts[0].getObjectType() !== AscDFH.historyitem_type_StockChart)
            {
                if(chartSpace.chart.plotArea.charts[0].series.length !== 4)
                {
                    this.sendEvent("asc_onError", Asc.c_oAscError.ID.StockChartError,Asc.c_oAscError.Level.NoCritical);
                    this.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
                    return;
                }
            }
        }
    }

    /* change z-index */
    if (AscFormat.isRealNumber(ImagePr.ChangeLevel))
    {
        switch (ImagePr.ChangeLevel)
        {
            case 0:
            {
                this.WordControl.m_oLogicDocument.DrawingObjects.bringToFront();
                break;
            }
            case 1:
            {
                this.WordControl.m_oLogicDocument.DrawingObjects.bringForward();
                break;
            }
            case 2:
            {
                this.WordControl.m_oLogicDocument.DrawingObjects.sendToBack();
                break;
            }
            case 3:
            {
                this.WordControl.m_oLogicDocument.DrawingObjects.bringBackward();
            }
        }
        return;
    }

    if ( false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props, AdditionalData) )
    {

        if (ImagePr.ShapeProperties)
            ImagePr.ImageUrl = "";

        if(ImagePr.ImageUrl != undefined && ImagePr.ImageUrl != null && ImagePr.ImageUrl != "")
        {
            this.WordControl.m_oLogicDocument.StartAction();
            this.WordControl.m_oLogicDocument.SetImageProps( ImagePr );
			this.WordControl.m_oLogicDocument.FinalizeAction();
        }
        else if (ImagePr.ShapeProperties && ImagePr.ShapeProperties.fill && ImagePr.ShapeProperties.fill.fill &&
                 ImagePr.ShapeProperties.fill.fill.url !== undefined && ImagePr.ShapeProperties.fill.fill.url != null && ImagePr.ShapeProperties.fill.fill.url != "")
        {
            this.WordControl.m_oLogicDocument.StartAction();
            this.WordControl.m_oLogicDocument.SetImageProps( ImagePr );
			this.WordControl.m_oLogicDocument.FinalizeAction();
        }
        else
        {
            ImagePr.ImageUrl = null;


            if(!this.noCreatePoint || this.exucuteHistory)
            {
                if( !this.noCreatePoint && !this.exucuteHistory && this.exucuteHistoryEnd)
                {
                    this.WordControl.m_oLogicDocument.SetImageProps( ImagePr );
                    this.exucuteHistoryEnd = false;
                }
                else
                {
                    this.WordControl.m_oLogicDocument.StartAction();
                    this.WordControl.m_oLogicDocument.SetImageProps( ImagePr );
					this.WordControl.m_oLogicDocument.FinalizeAction();
                }
                if(this.exucuteHistory)
                {
                    this.exucuteHistory = false;
                }
                if(this.exucuteHistoryEnd)
                {
                    this.exucuteHistoryEnd = false;
                }
            }
            else
            {
                AscFormat.ExecuteNoHistory(function(){
                                           this.WordControl.m_oLogicDocument.SetImageProps( ImagePr );
                                           }, this, []);
            }

            this.WordControl.m_oLogicDocument.SetImageProps( ImagePr );
        }
    }
};

Asc['asc_docs_api'].prototype.MergeCells = function()
{
    if ( false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Table_Properties) )
    {
        this.WordControl.m_oLogicDocument.StartAction();
        this.WordControl.m_oLogicDocument.MergeTableCells();
		this.WordControl.m_oLogicDocument.FinalizeAction();
    }
}
Asc['asc_docs_api'].prototype.SplitCell = function(Cols, Rows)
{
    if ( false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Table_Properties) )
    {
        this.WordControl.m_oLogicDocument.StartAction();
        this.WordControl.m_oLogicDocument.SplitTableCells(Cols, Rows);
		this.WordControl.m_oLogicDocument.FinalizeAction();
    }
}

Asc['asc_docs_api'].prototype.StartAddShape = function(sPreset, is_apply)
{
    this.isStartAddShape = true;
    this.addShapePreset = sPreset;
    if (is_apply)
    {
        this.WordControl.m_oDrawingDocument.LockCursorType("crosshair");
    }
    else
    {
        editor.sync_EndAddShape();
        editor.sync_StartAddShapeCallback(false);
    }
};

Asc['asc_docs_api'].prototype.AddImageUrlNative = function(url, _w, _h, _pageNum)
{
    var _section_select = this.WordControl.m_oLogicDocument.Get_PageSizesByDrawingObjects();
    var _page_width             = AscCommon.Page_Width;
    var _page_height            = AscCommon.Page_Height;
    var _page_x_left_margin     = AscCommon.X_Left_Margin;
    var _page_y_top_margin      = AscCommon.Y_Top_Margin;
    var _page_x_right_margin    = AscCommon.X_Right_Margin;
    var _page_y_bottom_margin   = AscCommon.Y_Bottom_Margin;

    if (_section_select)
    {
        if (_section_select.W)
            _page_width = _section_select.W;

        if (_section_select.H)
            _page_height = _section_select.H;
    }

    var __w = Math.max(1, (_page_width - (_page_x_left_margin + _page_x_right_margin)) / 2);
    var __h = Math.max(1, (_page_height - (_page_y_top_margin + _page_y_bottom_margin)) / 2);

    var wI = (undefined !== _w) ? Math.max(_w * AscCommon.g_dKoef_pix_to_mm, 1) : 1;
    var hI = (undefined !== _h) ? Math.max(_h * AscCommon.g_dKoef_pix_to_mm, 1) : 1;

    if (wI < 5)
        wI = 5;
    if (hI < 5)
        hI = 5;

    if (wI > __w || hI > __h)
    {
        var _koef = Math.min(__w / wI, __h / hI);
        wI *= _koef;
        hI *= _koef;
    }

    this.WordControl.m_oLogicDocument.StartAction();
    this.WordControl.m_oLogicDocument.AddInlineImage(wI, hI, url);
	this.WordControl.m_oLogicDocument.Recalculate();
	this.WordControl.m_oLogicDocument.FinalizeAction();
};
Asc['asc_docs_api'].prototype.AddImageUrlAtPosNative = function(url, _w, _h, _pageNum, _posX, _posY, _wrapType)
{
    _api.AddImageToPage(url, _pageNum, _posX, _posY, _w, _h, _wrapType);
};
Asc['asc_docs_api'].prototype.SetContentControlPictureUrlNative = function(sUrl, sId)
{
	if (this.WordControl && this.WordControl.m_oDrawingDocument)
	{
		this.WordControl.m_oDrawingDocument.UnlockCursorType();
	}

	var oLogicDocument = this.private_GetLogicDocument();
	if (!oLogicDocument || AscCommon.isNullOrEmptyString(sUrl))
		return;

	var oCC = oLogicDocument.GetContentControl(sId);
	oCC.SkipSpecialContentControlLock(true);
	if (!oCC || !oCC.IsPicture() || !oCC.SelectPicture() || !oCC.CanBeEdited())
	{
		oCC.SkipSpecialContentControlLock(false);
		return;
	}

	if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_Image_Properties, undefined, false, oLogicDocument.IsFormFieldEditing()))
	{
		oCC.SkipSpecialContentControlLock(false);
		oLogicDocument.StartAction(AscDFH.historydescription_Document_ApplyImagePrWithUrl);
		oLogicDocument.SetImageProps({ImageUrl : sUrl});
		oCC.SetShowingPlcHdr(false);
		oLogicDocument.UpdateTracks();
		oLogicDocument.FinalizeAction();
	}
	else
	{
		oCC.SkipSpecialContentControlLock(false);
	}
}
Asc['asc_docs_api'].prototype.AddImageUrlActionNative = function(src, _w, _h, _pageNum)
{
  var section_select = this.WordControl.m_oLogicDocument.Get_PageSizesByDrawingObjects();
  var page_width = AscCommon.Page_Width;
  var page_height = AscCommon.Page_Height;
  var page_x_left_margin = AscCommon.X_Left_Margin;
  var page_y_top_margin = AscCommon.Y_Top_Margin;
  var page_x_right_margin = AscCommon.X_Right_Margin;
  var page_y_bottom_margin = AscCommon.Y_Bottom_Margin;

  var boundingWidth  = Math.max(1, page_width  - (page_x_left_margin + page_x_right_margin));
  var boundingHeight = Math.max(1, page_height - (page_y_top_margin  + page_y_bottom_margin));

  var w = Math.max(_w * AscCommon.g_dKoef_pix_to_mm, 1);
  var h = Math.max(_h * AscCommon.g_dKoef_pix_to_mm, 1);
                                        
  var mW = boundingWidth  / w;
  var mH = boundingHeight / h;

  if (mH < mW) {
    boundingWidth  = boundingHeight / h * w;
  } else if (mW < mH) {
    boundingHeight = boundingWidth  / w * h;
  }

  _w = Math.max(5, boundingWidth);
  _h = Math.max(5, boundingHeight);  

  if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
  {
    var imageLocal = AscCommon.g_oDocumentUrls.getImageLocal(src);
    if (imageLocal)
    {
      src = imageLocal;
    }
    this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddImageUrlLong);
    //if (undefined === imgProp || undefined === imgProp.WrappingStyle || 0 == imgProp.WrappingStyle)
      this.WordControl.m_oLogicDocument.AddInlineImage(_w, _h, src);
    //else
    //  this.WordControl.m_oLogicDocument.AddInlineImage(_w, _h, src, null, true);
    this.WordControl.m_oLogicDocument.FinalizeAction();
  }
};

Asc['asc_docs_api'].prototype.Send_Menu_Event = function(type)
{
    window["native"]["OnCallMenuEvent"](type, global_memory_stream_menu);
};

Asc['asc_docs_api'].prototype.sync_EndCatchSelectedElements = function(isExternalTrigger)
{
    if (this.WordControl && this.WordControl.m_oDrawingDocument)
        this.WordControl.m_oDrawingDocument.EndTableStylesCheck();

    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();

    var _count = this.SelectedObjectsStack.length;
    var _naturalCount = 0;

    for (var i = 0; i < _count; i++)
    {
        switch (this.SelectedObjectsStack[i].Type)
        {
            case Asc.c_oAscTypeSelectElement.Paragraph:
            case Asc.c_oAscTypeSelectElement.Header:
            case Asc.c_oAscTypeSelectElement.Table:
            case Asc.c_oAscTypeSelectElement.Image:
            case Asc.c_oAscTypeSelectElement.Hyperlink:
            case Asc.c_oAscTypeSelectElement.Math:
            {
                ++_naturalCount;
                break;
            }
            default:
                break;
        }
    }

    _stream["WriteLong"](_naturalCount);

    for (var i = 0; i < _count; i++)
    {
        switch (this.SelectedObjectsStack[i].Type)
        {
                //console.log(JSON.stringify(this.SelectedObjectsStack[i].Value));
            case Asc.c_oAscTypeSelectElement.Paragraph:
            {
                _stream["WriteLong"](Asc.c_oAscTypeSelectElement.Paragraph);
                asc_menu_WriteParagraphPr(this.SelectedObjectsStack[i].Value, _stream);
                break;
            }
            case Asc.c_oAscTypeSelectElement.Header:
            {
                _stream["WriteLong"](Asc.c_oAscTypeSelectElement.Header);
                asc_menu_WriteHeaderFooterPr(this.SelectedObjectsStack[i].Value, _stream);
                break;
            }
            case Asc.c_oAscTypeSelectElement.Table:
            {
                _stream["WriteLong"](Asc.c_oAscTypeSelectElement.Table);
                asc_menu_WriteTablePr(this.SelectedObjectsStack[i].Value, _stream);
                break;
            }
            case Asc.c_oAscTypeSelectElement.Image:
            {
                _stream["WriteLong"](Asc.c_oAscTypeSelectElement.Image);
                asc_menu_WriteImagePr(this.SelectedObjectsStack[i].Value, _stream);
                break;
            }
            case Asc.c_oAscTypeSelectElement.Hyperlink:
            {
                _stream["WriteLong"](Asc.c_oAscTypeSelectElement.Hyperlink);
                asc_menu_WriteHyperPr(this.SelectedObjectsStack[i].Value, _stream);
                break;
            }
            case Asc.c_oAscTypeSelectElement.Math:
            {
                _stream["WriteLong"](Asc.c_oAscTypeSelectElement.Math);
                asc_menu_WriteMath(this.SelectedObjectsStack[i].Value, _stream);
                break;
            }
            case Asc.c_oAscTypeSelectElement.SpellCheck:
            default:
            {
                // none
                break;
            }
        }
    }

    this.Send_Menu_Event(6);
    this.sendEvent("asc_onFocusObject", this.SelectedObjectsStack, !isExternalTrigger);
};

Asc['asc_docs_api'].prototype.startGetDocInfo = function()
{
    /*
     Возвращаем объект следующего вида:
     {
     PageCount: 12,
     WordsCount: 2321,
     ParagraphCount: 45,
     SymbolsCount: 232345,
     SymbolsWSCount: 34356
     }
     */
    this.sync_GetDocInfoStartCallback();

    if (null != this.WordControl.m_oLogicDocument)
        this.WordControl.m_oLogicDocument.Statistics_Start();
};
Asc['asc_docs_api'].prototype.stopGetDocInfo = function()
{
    this.sync_GetDocInfoStopCallback();

    if (null != this.WordControl.m_oLogicDocument)
        this.WordControl.m_oLogicDocument.Statistics_Stop();
};
Asc['asc_docs_api'].prototype.sync_DocInfoCallback = function(obj)
{
    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();
    _stream["WriteLong"](obj.PageCount);
    _stream["WriteLong"](obj.WordsCount);
    _stream["WriteLong"](obj.ParagraphCount);
    _stream["WriteLong"](obj.SymbolsCount);
    _stream["WriteLong"](obj.SymbolsWSCount);

    window["native"]["OnCallMenuEvent"](70, _stream); // ASC_MENU_EVENT_TYPE_STATISTIC_INFO
};
Asc['asc_docs_api'].prototype.sync_GetDocInfoStartCallback = function()
{
    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();

    window["native"]["OnCallMenuEvent"](67, _stream); // ASC_MENU_EVENT_TYPE_STATISTIC_START
};
Asc['asc_docs_api'].prototype.sync_GetDocInfoStopCallback = function()
{
    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();

    window["native"]["OnCallMenuEvent"](68, _stream); // ASC_MENU_EVENT_TYPE_STATISTIC_STOP
};
Asc['asc_docs_api'].prototype.sync_GetDocInfoEndCallback = function()
{
    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();

    window["native"]["OnCallMenuEvent"](69, _stream); // ASC_MENU_EVENT_TYPE_STATISTIC_END
};

Asc['asc_docs_api'].prototype.sync_CanUndoCallback = function(bCanUndo)
{
    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();
    _stream["WriteBool"](bCanUndo);
    window["native"]["OnCallMenuEvent"](60, _stream); // ASC_MENU_EVENT_TYPE_CAN_UNDO
};
Asc['asc_docs_api'].prototype.sync_CanRedoCallback = function(bCanRedo)
{
    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();
    _stream["WriteBool"](bCanRedo);
    window["native"]["OnCallMenuEvent"](61, _stream); // ASC_MENU_EVENT_TYPE_CAN_REDO
};

Asc['asc_docs_api'].prototype.SetDocumentModified = function(bValue)
{
    this.isDocumentModify = bValue;

    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();
    _stream["WriteBool"](this.isDocumentModify);
    window["native"]["OnCallMenuEvent"](66, _stream); // ASC_MENU_EVENT_TYPE_DOCUMETN_MODIFITY
};

// find -------------------------------------------------------------------------------------------------
Asc['asc_docs_api'].prototype._selectSearchingResults = function(bShow)
{
    this.WordControl.m_oLogicDocument.HighlightSearchResults(bShow);
};

Asc['asc_docs_api'].prototype.asc_isSelectSearchingResults = function()
{
    return this.WordControl.m_oLogicDocument.IsHighlightSearchResults();
};
// endfind ----------------------------------------------------------------------------------------------

// sectionPr --------------------------------------------------------------------------------------------
Asc['asc_docs_api'].prototype.change_PageOrient = function(isPortrait)
{
    if ( false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Document_SectPr) )
    {
        this.WordControl.m_oLogicDocument.StartAction();
        if (isPortrait)
        {
            this.WordControl.m_oLogicDocument.Set_DocumentOrientation(Asc.c_oAscPageOrientation.PagePortrait);
            this.DocumentOrientation = isPortrait;
        }
        else
        {
            this.WordControl.m_oLogicDocument.Set_DocumentOrientation(Asc.c_oAscPageOrientation.PageLandscape);
            this.DocumentOrientation = isPortrait;
        }
		this.WordControl.m_oLogicDocument.FinalizeAction();
        this.sync_PageOrientCallback(editor.get_DocumentOrientation());
    }
};
Asc['asc_docs_api'].prototype.change_DocSize = function(width,height)
{
    if ( false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Document_SectPr) )
    {
        this.WordControl.m_oLogicDocument.StartAction();
        if (this.DocumentOrientation)
            this.WordControl.m_oLogicDocument.Set_DocumentPageSize(width, height);
        else
            this.WordControl.m_oLogicDocument.Set_DocumentPageSize(height, width);

		this.WordControl.m_oLogicDocument.FinalizeAction();
    }
};
Asc['asc_docs_api'].prototype.sync_PageOrientCallback = function(isPortrait)
{
    this.DocumentOrientation = isPortrait;
    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();
    _stream["WriteByte"](6);
    _stream["WriteBool"](this.DocumentOrientation);
    _stream["WriteByte"](255);
    this.Send_Menu_Event(17, _stream); // ASC_MENU_EVENT_TYPE_SECTION
};
Asc['asc_docs_api'].prototype.sync_DocSizeCallback = function(width,height)
{
    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();
    _stream["WriteByte"](0);
    _stream["WriteDouble2"](width);
    _stream["WriteByte"](1);
    _stream["WriteDouble2"](height);
    _stream["WriteByte"](255);
    this.Send_Menu_Event(17, _stream); // ASC_MENU_EVENT_TYPE_SECTION
};
// endsectionPr -----------------------------------------------------------------------------------------

// STYLES
function CStylesPainter()
{
    this.STYLE_THUMBNAIL_WIDTH  = AscCommon.GlobalSkin.STYLE_THUMBNAIL_WIDTH;
    this.STYLE_THUMBNAIL_HEIGHT = AscCommon.GlobalSkin.STYLE_THUMBNAIL_HEIGHT;

    this.CurrentTranslate = null;
    this.IsRetinaEnabled = false;

    this.defaultStyles = [];
    this.docStyles = [];
    this.mergedStyles = [];
}
CStylesPainter.prototype =
{
    GenerateStyles: function(_api, ds)
    {
        if (_api.WordControl.bIsRetinaSupport)
        {
            this.STYLE_THUMBNAIL_WIDTH  <<= 1;
            this.STYLE_THUMBNAIL_HEIGHT <<= 1;
            this.IsRetinaEnabled = true;
        }

        this.CurrentTranslate = _api.CurrentTranslate;

        var _stream = global_memory_stream_menu;
        var _graphics = new CDrawingStream();

        _api.WordControl.m_oDrawingDocument.Native["DD_PrepareNativeDraw"]();

        this.GenerateDefaultStyles(_api, ds, _graphics);
        this.GenerateDocumentStyles(_api, _graphics);

        // стили сформированы. осталось просто сформировать единый список
        var _count_default = this.defaultStyles.length;
        var _count_doc = 0;
        if (null != this.docStyles)
            _count_doc = this.docStyles.length;

        var aPriorityStyles = [];
        var fAddToPriorityStyles = function(style){
            var index = style.Style.uiPriority;
            if(null == index)
                index = 0;
            var aSubArray = aPriorityStyles[index];
            if(null == aSubArray)
            {
                aSubArray = [];
                aPriorityStyles[index] = aSubArray;
            }
            aSubArray.push(style);
        };
        var _map_document = {};

        for (var i = 0; i < _count_doc; i++)
        {
            var style = this.docStyles[i];
            _map_document[style.Name] = 1;
            fAddToPriorityStyles(style);
        }

        for (var i = 0; i < _count_default; i++)
        {
            var style = this.defaultStyles[i];
            if(null == _map_document[style.Name])
                fAddToPriorityStyles(style);
        }

        this.mergedStyles = [];
        for(var index in aPriorityStyles)
        {
            var aSubArray = aPriorityStyles[index];
            aSubArray.sort(function(a, b){
                           if(a.Name < b.Name)
                           return -1;
                           else if(a.Name > b.Name)
                           return 1;
                           else
                           return 0;
                           });
            for(var i = 0, length = aSubArray.length; i < length; ++i)
            {
                this.mergedStyles.push(aSubArray[i]);
            }
        }

        var _count = this.mergedStyles.length;
        for (var i = 0; i < _count; i++)
        {
            this.drawStyle(_graphics, this.mergedStyles[i].Style, _api);
        }

        _stream["ClearNoAttack"]();

        _stream["WriteByte"](1);

        _api.WordControl.m_oDrawingDocument.Native["DD_EndNativeDraw"](_stream);
    },
    GenerateDefaultStyles: function(_api, ds, _graphics)
    {
        var styles = ds;

        for (var i in styles)
        {
            var style = styles[i];
            if (true == style.qFormat)
            {
                this.defaultStyles.push({ Name: style.Name, Style: style });
            }
        }
    },

    GenerateDocumentStyles: function(_api, _graphics)
    {
        if (_api.WordControl.m_oLogicDocument == null)
            return;

        var __Styles = _api.WordControl.m_oLogicDocument.Get_Styles();
        var styles = __Styles.Style;

        if (styles == null)
            return;

        for (var i in styles)
        {
            var style = styles[i];
            if (true == style.qFormat)
            {
                // как только меняется сериалайзер - меняется и код здесь. Да, не очень удобно,
                // зато быстро делается
                var formalStyle = i.toLowerCase().replace(/\s/g, "");
                var res = formalStyle.match(/^heading([1-9][0-9]*)$/);
                var index = (res) ? res[1] - 1 : -1;

                var _dr_style = __Styles.Get_Pr(i, styletype_Paragraph);
                _dr_style.Name = style.Name;
                _dr_style.Id = i;

                var _name = _dr_style.Name;
                
                // алгоритм смены имени
                if (style.Default)
                {
                    switch (style.Default)
                    {
                        case 1:
                            break;
                        case 2:
                            _name = "No List";
                            break;
                        case 3:
                            _name = "Normal";
                            break;
                        case 4:
                            _name = "Normal Table";
                            break;
                    }
                }
                else if (index != -1)
                {
                    _name = "Heading ".concat(index + 1);
                }

                this.docStyles.push({ Name: _name, Style: _dr_style });
            }
        }
    },

    drawStyle: function(graphics, style, _api)
    {
        var _w_px = this.STYLE_THUMBNAIL_WIDTH;
        var _h_px = this.STYLE_THUMBNAIL_HEIGHT;
        var dKoefToMM = AscCommon.g_dKoef_pix_to_mm;

        if (AscCommon.AscBrowser.isRetina)
            dKoefToMM /= 2;

        _api.WordControl.m_oDrawingDocument.Native["DD_StartNativeDraw"](_w_px, _h_px, _w_px * dKoefToMM, _h_px * dKoefToMM);

        AscCommon.g_oTableId.m_bTurnOff = true;
        AscCommon.History.TurnOff();

        var oldDefTabStop = AscCommonWord.Default_Tab_Stop;
        AscCommonWord.Default_Tab_Stop = 1;

        var hdr = new CHeaderFooter(_api.WordControl.m_oLogicDocument.HdrFtr, _api.WordControl.m_oLogicDocument, _api.WordControl.m_oDrawingDocument, AscCommon.hdrftr_Header);
        var _dc = hdr.Content;//new CDocumentContent(editor.WordControl.m_oLogicDocument, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, false, true, false);

        var par = new AscCommonWord.Paragraph(_api.WordControl.m_oDrawingDocument, _dc, 0, 0, 0, 0, false);
        var run = new AscCommonWord.ParaRun(par, false);
        run.AddText(AscCommon.translateManager.getValue(style.Name));

        _dc.Internal_Content_Add(0, par, false);
        par.Add_ToContent(0, run);
        par.Style_Add(style.Id, false);
        par.Set_Align(AscCommon.align_Left);
        par.Set_Tabs(new AscCommonWord.CParaTabs());

        var _brdL = style.ParaPr.Brd.Left;
        if ( undefined !== _brdL && null !== _brdL )
        {
            var brdL = new CDocumentBorder();
            brdL.Set_FromObject(_brdL);
            brdL.Space = 0;
            par.Set_Border(brdL, AscDFH.historyitem_Paragraph_Borders_Left);
        }

        var _brdT = style.ParaPr.Brd.Top;
        if ( undefined !== _brdT && null !== _brdT )
        {
            var brd = new CDocumentBorder();
            brd.Set_FromObject(_brdT);
            brd.Space = 0;
            par.Set_Border(brd, AscDFH.historyitem_Paragraph_Borders_Top);
        }

        var _brdB = style.ParaPr.Brd.Bottom;
        if ( undefined !== _brdB && null !== _brdB )
        {
            var brd = new CDocumentBorder();
            brd.Set_FromObject(_brdB);
            brd.Space = 0;
            par.Set_Border(brd, AscDFH.historyitem_Paragraph_Borders_Bottom);
        }

        var _brdR = style.ParaPr.Brd.Right;
        if ( undefined !== _brdR && null !== _brdR )
        {
            var brd = new CDocumentBorder();
            brd.Set_FromObject(_brdR);
            brd.Space = 0;
            par.Set_Border(brd, AscDFH.historyitem_Paragraph_Borders_Right);
        }

        var _ind = new CParaInd();
        _ind.FirstLine = 0;
        _ind.Left = 0;
        _ind.Right = 0;
        par.Set_Ind(_ind, false);

        var _sp = new CParaSpacing();
        _sp.Line              = 1;
        _sp.LineRule          = Asc.linerule_Auto;
        _sp.Before            = 0;
        _sp.BeforeAutoSpacing = false;
        _sp.After             = 0;
        _sp.AfterAutoSpacing  = false;
        par.Set_Spacing(_sp, false);

        _dc.Reset(0, 0, 10000, 10000);
        _dc.Recalculate_Page(0, true);

        _dc.Reset(0, 0, par.Lines[0].Ranges[0].W + 0.001, 10000);
        _dc.Recalculate_Page(0, true);

        var y = 0;
        var b = dKoefToMM * _h_px;
        var w = dKoefToMM * _w_px;
        var off = 10 * dKoefToMM;
        var off2 = 5 * dKoefToMM;
        var off3 = 1 * dKoefToMM;

        graphics.transform(1,0,0,1,0,0);
        graphics.save();
        graphics._s();
        graphics._m(off2, y + off3);
        graphics._l(w - off, y + off3);
        graphics._l(w - off, b - off3);
        graphics._l(off2, b - off3);
        graphics._z();
        graphics.clip();

        var baseline = par.Lines[0].Y;
        par.Shift(0, off + 0.5, y + 0.75 * (b - y) - baseline);
        par.Draw(0, graphics);

        graphics.restore();

        AscCommonWord.Default_Tab_Stop = oldDefTabStop;

        AscCommon.g_oTableId.m_bTurnOff = false;
        AscCommon.History.TurnOn();

        var _stream = global_memory_stream_menu;

        _stream["ClearNoAttack"]();

        _stream["WriteByte"](0);
        _stream["WriteString2"](style.Name);

        _api.WordControl.m_oDrawingDocument.Native["DD_EndNativeDraw"](_stream);
        graphics.ClearParams();
    }
};

// -------------------------------------------------

/***************************** COPY|PASTE *******************************/

Asc['asc_docs_api'].prototype.Call_Menu_Context_Copy = function()
{
    var dataBuffer = {};

    var clipboard = {};
    clipboard.pushData = function(type, data) {

        if (AscCommon.c_oAscClipboardDataFormat.Text === type) {

            dataBuffer.text = data;

        } else if (AscCommon.c_oAscClipboardDataFormat.Internal === type) {

            if (null != data.drawingUrls && data.drawingUrls.length > 0) {
                dataBuffer.drawingUrls = data.drawingUrls[0];
            }

            dataBuffer.sBase64 = data.sBase64;
        }
    }

    this.asc_CheckCopy(clipboard, AscCommon.c_oAscClipboardDataFormat.Internal|AscCommon.c_oAscClipboardDataFormat.Text);

    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();

    if (dataBuffer.text) {
        _stream["WriteByte"](0);
        _stream["WriteString2"](dataBuffer.text);
    }

    if (dataBuffer.drawingUrls) {
        _stream["WriteByte"](1);
        _stream["WriteStringA"](dataBuffer.drawingUrls);
    }

    if (dataBuffer.sBase64) {
        _stream["WriteByte"](2);
        _stream["WriteStringA"](dataBuffer.sBase64);
    }

    _stream["WriteByte"](255);

    return _stream;
};
Asc['asc_docs_api'].prototype.Call_Menu_Context_Cut = function()
{
    var dataBuffer = {};

    var clipboard = {};
    clipboard.pushData = function(type, data) {

        if (AscCommon.c_oAscClipboardDataFormat.Text === type) {

            dataBuffer.text = data;

        } else if (AscCommon.c_oAscClipboardDataFormat.Internal === type) {

            if (null != data.drawingUrls && data.drawingUrls.length > 0) {
                dataBuffer.drawingUrls = data.drawingUrls[0];
            }

            dataBuffer.sBase64 = data.sBase64;
        }
    }

    this.asc_CheckCopy(clipboard, AscCommon.c_oAscClipboardDataFormat.Internal|AscCommon.c_oAscClipboardDataFormat.Text);

    this.asc_SelectionCut();

    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();

    if (dataBuffer.text) {
        _stream["WriteByte"](0);
        _stream["WriteString2"](dataBuffer.text);
    }

    if (dataBuffer.drawingUrls) {
        _stream["WriteByte"](1);
        _stream["WriteStringA"](dataBuffer.drawingUrls);
    }

    if (dataBuffer.sBase64) {
        _stream["WriteByte"](2);
        _stream["WriteStringA"](dataBuffer.sBase64);
    }

    _stream["WriteByte"](255);

    return _stream;
};
Asc['asc_docs_api'].prototype.Call_Menu_Context_Paste = function(type, param)
{
    if (0 == type)
    {
        this.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, param);
    }
    else if (1 == type)
    {
        this.AddImageUrlNative(param, 200, 200);
    }
    else if (2 == type)
    {
        this.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Internal, param);
    }
};
Asc['asc_docs_api'].prototype.Call_Menu_Context_Delete = function()
{
    if ( false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Delete) )
    {
        this.WordControl.m_oLogicDocument.StartAction();
        this.WordControl.m_oLogicDocument.Remove( 1, true );
		this.WordControl.m_oLogicDocument.FinalizeAction();
    }
};
Asc['asc_docs_api'].prototype.Call_Menu_Context_Select = function()
{
    this.WordControl.m_oLogicDocument.MoveCursorLeft(false, true);
    this.WordControl.m_oLogicDocument.MoveCursorRight(true, true);
    this.WordControl.m_oLogicDocument.Document_UpdateSelectionState();
};
Asc['asc_docs_api'].prototype.Call_Menu_Context_SelectAll = function()
{
    this.WordControl.m_oLogicDocument.SelectAll();
};
/************************************************************************/

// chat styles
AscCommon.ChartPreviewManager.prototype.clearPreviews = function()
{
    window["native"]["ClearCacheChartStyles"]();
};

AscCommon.ChartPreviewManager.prototype.createChartPreview = function(_graphics, type, styleIndex)
{
    return AscFormat.ExecuteNoHistory(function(){
      var chart_space = this.checkChartForPreview(type, AscCommon.g_oChartStyles[type][styleIndex]);
      // sizes for imageView
      window["native"]["DD_StartNativeDraw"](85 * 2, 85 * 2, 75, 75);

      chart_space.draw(_graphics);
      _graphics.ClearParams();

      var _stream = global_memory_stream_menu;
      _stream["ClearNoAttack"]();
      _stream["WriteByte"](4);
      _stream["WriteLong"](type);
      _stream["WriteLong"](styleIndex);
      window["native"]["DD_EndNativeDraw"](_stream);

      }, this, []);
};

AscCommon.ChartPreviewManager.prototype.getChartPreviews = function(chartType)
{
    if (AscFormat.isRealNumber(chartType))
    {
        var bIsCached = window["native"]["IsCachedChartStyles"](chartType);
        if (!bIsCached)
        {
            window["native"]["DD_PrepareNativeDraw"]();

            var _graphics = new CDrawingStream();

            if(AscCommon.g_oChartStyles[chartType]){
                var nStylesCount = AscCommon.g_oChartStyles[chartType].length;
                for(var i = 0; i < nStylesCount; ++i)
                    this.createChartPreview(_graphics, chartType, i);
            }

            var _stream = global_memory_stream_menu;
            _stream["ClearNoAttack"]();
            _stream["WriteByte"](5);
            _api.WordControl.m_oDrawingDocument.Native["DD_EndNativeDraw"](_stream);
        }
    }
};

function initTrackRevisions() {
    this.trackRevisionsTimerId = setInterval(function () {
        _api.WordControl.m_oLogicDocument.ContinueTrackRevisions();
    }, 40);
}

function initSpellCheckApi() {
    
    _api.SpellCheckApi = new AscCommon.CSpellCheckApi();
    _api.isSpellCheckEnable = true;

    _api.SpellCheckApi.spellCheck = function (spellData) {
        window["native"]["SpellCheck"](JSON.stringify(spellData));
    };
    
    _api.SpellCheckApi.disconnect = function () {};

    _api.sendEvent('asc_onSpellCheckInit', [
        "1026",
        "1027",
        "1029",
        "1030",
        "1031",
        "1032",
        "1033",
        "1036",
        "1038",
        "1040",
        "1042",
        "1043",
        "1044",
        "1045",
        "1046",
        "1048",
        "1049",
        "1050",
        "1051",
        "1053",
        "1055",
        "1057",
        "1058",
        "1060",
        "1062",
        "1063",
        "1066",
        "1068",
        "1069",
        "1087",
        "1104",
        "1110",
        "1134",
        "2051",
        "2055",
        "2057",
        "2068",
        "2070",
        "3079",
        "3081",
        "3082",
        "4105",
        "7177",
        "9242",
        "10266"
    ]);

    _api.SpellCheckApi.onInit = function (e) {
        _api.sendEvent('asc_onSpellCheckInit', e);
    };

    _api.SpellCheckApi.onSpellCheck = function (e) {
        _api.SpellCheck_CallBack(e);
    };

    _api.SpellCheckApi.init(_api.documentId);

    _api.asc_setSpellCheck(spellCheck);
}

function NativeOpenFile3(_params, documentInfo)
{
    window.NATIVE_DOCUMENT_TYPE = window["native"]["GetEditorType"]();
    if (window.NATIVE_DOCUMENT_TYPE == "presentation" || window.NATIVE_DOCUMENT_TYPE == "document")
    {
        sdkCheck = documentInfo["sdkCheck"];
        spellCheck = documentInfo["spellCheck"];

        var translations = documentInfo["translations"];
        if (undefined != translations && null != translations && translations.length > 0) {
            translations = JSON.parse(translations)
        } else {
            translations = "";
        }

        window["_api"] = window["API"] = _api = new window["Asc"]["asc_docs_api"](translations);
        window["_editor"] = window.editor;
        
        AscCommon.g_clipboardBase.Init(_api);

        if (undefined !== _api["Native_Editor_Initialize_Settings"])
        {
            _api["Native_Editor_Initialize_Settings"](_params);
        }

        window.documentInfo = documentInfo;

        var userInfo = new Asc.asc_CUserInfo();
        userInfo.asc_putId(window.documentInfo["docUserId"]);
        userInfo.asc_putFullName(window.documentInfo["docUserName"]);
        userInfo.asc_putFirstName(window.documentInfo["docUserFirstName"]);
        userInfo.asc_putLastName(window.documentInfo["docUserLastName"]);

        var docInfo = new Asc.asc_CDocInfo();
        docInfo.put_Id(window.documentInfo["docKey"]);
        docInfo.put_Url(window.documentInfo["docURL"]);
        docInfo.put_Format("docx");
        docInfo.put_UserInfo(userInfo);
        docInfo.put_Token(window.documentInfo["token"]);

        _internalStorage.externalUserInfo = userInfo;
        _internalStorage.externalDocInfo = docInfo;

        var permissions = window.documentInfo["permissions"];
        if (undefined != permissions && null != permissions && permissions.length > 0) {
            docInfo.put_Permissions(JSON.parse(permissions));
        }

        // Settings

        _api.asc_setDocInfo(docInfo);

        _api.asc_registerCallback("asc_onAdvancedOptions", function(type, options) {
                                  var stream = global_memory_stream_menu;
                                  if (options === undefined) {
                                    options = {};
                                  }
                                  options["optionId"] = type;
                                  stream["ClearNoAttack"]();
                                  stream["WriteString2"](JSON.stringify(options));
                                  window["native"]["OnCallMenuEvent"](22000, stream); // ASC_MENU_EVENT_TYPE_ADVANCED_OPTIONS
                                  });

        _api.asc_registerCallback("asc_onSendThemeColorSchemes", function(schemes) {
                                  var stream = global_memory_stream_menu;
                                  stream["ClearNoAttack"]();
                                  asc_WriteColorSchemes(schemes, stream);
                                  window["native"]["OnCallMenuEvent"](2404, stream); // ASC_SPREADSHEETS_EVENT_TYPE_COLOR_SCHEMES
                                  });
        _api.asc_registerCallback("asc_onSendThemeColors", onApiSendThemeColors);

        _api.asc_registerCallback("asc_onFocusObject", onFocusObject);
        _api.asc_registerCallback('asc_onStartAction', onApiLongActionBegin);
        _api.asc_registerCallback('asc_onEndAction', onApiLongActionEnd);
        _api.asc_registerCallback('asc_onError', onApiError);
        
        // Comments

        _api.asc_registerCallback("asc_onAddComment", onApiAddComment);
        _api.asc_registerCallback("asc_onAddComments", onApiAddComments);
        _api.asc_registerCallback("asc_onRemoveComment", onApiRemoveComment);
        _api.asc_registerCallback("asc_onChangeComments", onApiChangeComments);
        _api.asc_registerCallback("asc_onRemoveComments", onApiRemoveComments);
        _api.asc_registerCallback("asc_onChangeCommentData", onApiChangeCommentData);
        _api.asc_registerCallback("asc_onLockComment", onApiLockComment);
        _api.asc_registerCallback("asc_onUnLockComment", onApiUnLockComment);
        _api.asc_registerCallback("asc_onShowComment", onApiShowComment);
        _api.asc_registerCallback("asc_onHideComment", onApiHideComment);
        _api.asc_registerCallback("asc_onUpdateCommentPosition", onApiUpdateCommentPosition);

        // Revisions Change

        _api.asc_registerCallback("asc_onDocumentPlaceChanged", onDocumentPlaceChanged);
        _api.asc_registerCallback("asc_onShowRevisionsChange", onApiShowRevisionsChange);

        // Fill forms
        
        _api.asc_registerCallback('asc_onShowContentControlsActions', onShowContentControlsActions);
        _api.asc_registerCallback('asc_onHideContentControlsActions', onHideContentControlsActions);

        // Co-authoring

        if (window.documentInfo["iscoauthoring"]) {

            _api.isSpellCheckEnable = false;

            _api.asc_setAutoSaveGap(1);
            _api._coAuthoringInit();
            _api.asc_SetFastCollaborative(true);
            _api.SetCollaborativeMarksShowType(Asc.c_oAscCollaborativeMarksShowType.None);
           
            window["native"]["onTokenJWT"](_api.CoAuthoringApi.get_jwt());

            _api.asc_registerCallback("asc_onAuthParticipantsChanged", onApiAuthParticipantsChanged);
            _api.asc_registerCallback("asc_onParticipantsChanged", onApiParticipantsChanged);

            _api.asc_registerCallback("asc_onGetEditorPermissions", function(state) {

                                      var rData = {
                                      "c"             : "open",
                                      "id"            : window.documentInfo["docKey"],
                                      "userid"        : window.documentInfo["docUserId"],
                                      "format"        : "docx",
                                      "vkey"          : undefined,
                                      "url"           : window.documentInfo["docURL"],
                                      "title"         : this.documentTitle,
                                      "nobase64"      : true
                                      };

                                      _api.CoAuthoringApi.auth(window.documentInfo["viewmode"], rData);
                                      });

            _api.asc_registerCallback("asc_onDocumentUpdateVersion", function(callback) {
                                      var me = this;
                                      me.needToUpdateVersion = true;
                                      if (callback) callback.call(me);
                                      });


        } else {
            var doc_bin = window["native"]["GetFileString"]("native_open_file");
            _api["asc_nativeOpenFile"](doc_bin);

           	if (window.documentInfo["viewmode"]) {
            	_api.ShowParaMarks = false;
            	AscCommon.CollaborativeEditing.Set_GlobalLock(true);
            	_api.isViewMode = true;
            	_api.WordControl.m_oDrawingDocument.IsViewMode = true;
          	}

            if (null != _api.WordControl.m_oLogicDocument)
            {
                _api.WordControl.m_oDrawingDocument.CheckGuiControlColors();
                _api.sendColorThemes(_api.WordControl.m_oLogicDocument.theme);
            }

            _api["NativeAfterLoad"]();

            initSpellCheckApi();
            initTrackRevisions();
        }
    }
    _api.isDocumentLoadComplete = true;
}

// Revisions Change

function onApiShowRevisionsChange(data) {
    _internalStorage.changesReview = data
    var revisionChanges = [];

    if (data && data.length > 0) {
        var recalcFromMM = function (value) {
            return parseFloat((value / 10.).toFixed(4)); // value in mm need to convert to cm
        };
        var metricName = "|cm|";
        var c_paragraphLinerule = {
            LINERULE_LEAST: 0,
            LINERULE_AUTO: 1,
            LINERULE_EXACT: 2
        };

        data.forEach(function (item) {
            var commonChanges = [], changes = [],
                value = item.get_Value(),
                movetype = item.get_MoveType();
            switch (item.get_Type()) {
                case Asc.c_oAscRevisionsChangeType.TextAdd:
                    commonChanges.push((movetype == Asc.c_oAscRevisionsMove.NoMove) ? "|Inserted|" : "|Moved|");
                    if (typeof value == 'object') {
                        value.forEach(function (obj) {
                            if (typeof obj === 'string') {
                                changes.push(obj);
                            } else {
                                switch (obj) {
                                    case 0:
                                        changes.push("<|Image|>");
                                        break;
                                    case 1:
                                        changes.push("<|Shape|>");
                                        break;
                                    case 2:
                                        changes.push("<|Chart|>");
                                        break;
                                    case 3:
                                        changes.push("<|Equation|>");
                                        break;
                                }
                            }
                        });
                    } else if (typeof value === 'string') {
                        changes.push(value);
                    }
                    break;
                case Asc.c_oAscRevisionsChangeType.TextRem:
                    commonChanges.push((movetype == Asc.c_oAscRevisionsMove.NoMove) ? "|Deleted|" : (item.is_MovedDown() ? "|Moved Down|" : "|Moved Up"));
                    if (typeof value == 'object') {
                        value.forEach(function (obj) {
                            if (typeof obj === 'string') {
                                changes.push(obj);
                            } else {
                                switch (obj) {
                                    case 0:
                                        changes.push("<|Image|>");
                                        break;
                                    case 1:
                                        changes.push("<|Shape|>");
                                        break;
                                    case 2:
                                        changes.push("<|Chart|>");
                                        break;
                                    case 3:
                                        changes.push("<|Equation|>");
                                        break;
                                }
                            }
                        });
                    } else if (typeof value === 'string') {
                        changes.push(value);
                    }
                    break;
                case Asc.c_oAscRevisionsChangeType.ParaAdd:
                    commonChanges.push("|Paragraph Inserted|");
                    break;
                case Asc.c_oAscRevisionsChangeType.ParaRem:
                    commonChanges.push("|Paragraph Deleted|");
                    break;
                case Asc.c_oAscRevisionsChangeType.TextPr:
                    commonChanges.push("|Formatted|");
                    if (value.Get_Bold() !== undefined)
                        changes.push((value.Get_Bold() ? '' : ("|Not|" + " ")) + "|Bold|");
                    if (value.Get_Italic() !== undefined)
                        changes.push((value.Get_Italic() ? '' : ("|Not|" + " ")) + "|Italic|");
                    if (value.Get_Underline() !== undefined)
                        changes.push((value.Get_Underline() ? '' : ("|Not|" + " ")) + "|Underline|");
                    if (value.Get_Strikeout() !== undefined)
                        changes.push((value.Get_Strikeout() ? '' : ("|Not|" + " ")) + "|Strikeout|");
                    if (value.Get_DStrikeout() !== undefined)
                        changes.push((value.Get_DStrikeout() ? '' : ("|Not|" + " ")) + "|Double strikeout|");
                    if (value.Get_Caps() !== undefined)
                        changes.push((value.Get_Caps() ? '' : ("|Not|" + " ")) + "|All caps|");
                    if (value.Get_SmallCaps() !== undefined)
                        changes.push((value.Get_SmallCaps() ? '' : ("|Not|" + " ")) + "|Small caps|");
                    if (value.GetVertAlign() !== undefined)
                        changes.push(((value.GetVertAlign() === AscCommon.vertalign_SuperScript) ? "|Superscript|" : ((value.GetVertAlign() === AscCommon.vertalign_SubScript) ? "|Subscript|" : "|Baseline|")));
                    if (value.Get_Color() !== undefined)
                        changes.push("|Font color|");
                    if (value.Get_Highlight() !== undefined)
                        changes.push("|Highlight color|");
                    if (value.Get_Shd() !== undefined)
                        changes.push("|Background color|");
                    if (value.Get_FontFamily() !== undefined)
                        changes.push(value.Get_FontFamily());
                    if (value.Get_FontSize() !== undefined)
                        changes.push(value.Get_FontSize());
                    if (value.Get_Spacing() !== undefined)
                        changes.push("|Spacing|" + " " + recalcFromMM(value.Get_Spacing()).toFixed(2) + " " + metricName);
                    if (value.Get_Position() !== undefined)
                        changes.push("|Position|" + " " + recalcFromMM(value.Get_Position()).toFixed(2) + " " + metricName);
                    if (value.Get_Lang() !== undefined)
                        changes.push("[lang]" + value.Get_Lang());
                    break;
                case Asc.c_oAscRevisionsChangeType.ParaPr:
                    commonChanges.push("|Paragraph Formatted|");
                    if (value.Get_ContextualSpacing())
                        changes.push(value.Get_ContextualSpacing() ? "|Don't add interval between paragraphs of the same style|" : "|Add interval between paragraphs of the same style|");
                    if (value.Get_IndLeft() !== undefined)
                        changes.push("|Indent left|" + " " + recalcFromMM(value.Get_IndLeft()).toFixed(2) + " " + metricName);
                    if (value.Get_IndRight() !== undefined)
                        changes.push("|Indent right|" + " " + recalcFromMM(value.Get_IndRight()).toFixed(2) + " " + metricName);
                    if (value.Get_IndFirstLine() !== undefined)
                        changes.push("|First line|" + " " + recalcFromMM(value.Get_IndFirstLine()).toFixed(2) + " " + metricName);
                    if (value.Get_Jc() !== undefined) {
                        switch (value.Get_Jc()) {
                            case 0:
                                changes.push("|Align right|");
                                break;
                            case 1:
                                changes.push("|Align left|");
                                break;
                            case 2:
                                changes.push("|Align center|");
                                break;
                            case 3:
                                changes.push("|Align justify|");
                                break;
                        }
                    }
                    if (value.Get_KeepLines() !== undefined)
                        changes.push(value.Get_KeepLines() ? "|Keep lines together|" : "|Don't keep lines together|");
                    if (value.Get_KeepNext())
                        changes.push(value.Get_KeepNext() ? "|Keep with next|" : "|Don't keep with next|");
                    if (value.Get_PageBreakBefore())
                        changes.push(value.Get_PageBreakBefore() ? "|Page break before|" : "|No page break before|");
                    if (value.Get_SpacingLineRule() !== undefined && value.Get_SpacingLine() !== undefined) {
                        changes.push("|Line Spacing|: ");
                        changes.push((value.Get_SpacingLineRule() == c_paragraphLinerule.LINERULE_LEAST) ? "|at least|" : ((value.Get_SpacingLineRule() == c_paragraphLinerule.LINERULE_AUTO) ? "|multiple|" : "|exactly|"));
                        changes.push((value.Get_SpacingLineRule() == c_paragraphLinerule.LINERULE_AUTO) ? value.Get_SpacingLine() : recalcFromMM(value.Get_SpacingLine()).toFixed(2) + " " + metricName);
                    }
                    if (value.Get_SpacingBeforeAutoSpacing())
                        changes.push("|Spacing before| |auto|");
                    else if (value.Get_SpacingBefore() !== undefined)
                        changes.push("|Spacing before|" + " " + recalcFromMM(value.Get_SpacingBefore()).toFixed(2) + ' ' + metricName);
                    if (value.Get_SpacingAfterAutoSpacing())
                        changes.push("|Spacing after| |auto|");
                    else if (value.Get_SpacingAfter() !== undefined)
                        changes.push("|Spacing after| " + recalcFromMM(value.Get_SpacingAfter()).toFixed(2) + " " + metricName);
                    if (value.Get_WidowControl())
                        changes.push(value.Get_WidowControl() ? "|Widow control|" : "|No widow control|");
                    if (value.Get_Tabs() !== undefined)
                        changes.push("|Change tabs|");
                    if (value.Get_NumPr() !== undefined)
                        changes.push("|Change numbering|");
                    if (value.GetPStyle() !== undefined) {
                        var style = me.api.asc_GetStyleNameById(value.GetPStyle());
                        if (!_.isEmpty(style)) {
                            changes.push(style)
                        }
                    }
                    break;
                case Asc.c_oAscRevisionsChangeType.TablePr:
                case Asc.c_oAscRevisionsChangeType.TableRowPr:
                    commonChanges.push("|Table Settings Changed|");
                    break;
                case Asc.c_oAscRevisionsChangeType.RowsAdd:
                    commonChanges.push("|Table Rows Added|");
                    break;
                case Asc.c_oAscRevisionsChangeType.RowsRem:
                    commonChanges.push("|Table Rows Deleted|");
                    break;
            }

            var userName = '';

            if (item.get_UserName() != '') {
                userName = item.get_UserName()
            } else {
                if (_internalStorage.externalUserInfo && _internalStorage.externalUserInfo.asc_getFullName) {
                    userName = _internalStorage.externalUserInfo.asc_getFullName()
                }
            }

            var revisionChange = {
                userName: userName,
                userId: item.get_UserId(),
                lock: (item.get_LockUserId()!==null),
                date: (item.get_DateTime() == '' ? new Date().getMilliseconds() : item.get_DateTime()),
                goto: (item.get_MoveType() == Asc.c_oAscRevisionsMove.MoveTo || item.get_MoveType() == Asc.c_oAscRevisionsMove.MoveFrom),
                commonChanges: commonChanges,
                changes: changes
            };

            Object.keys(revisionChange).forEach(function (key) {
                if (typeof revisionChange[key] === 'undefined') {
                    delete revisionChange[key];
                }
            });

            revisionChanges.push(revisionChange);
        });
    }
    postDataAsJSONString(revisionChanges, 24001); // ASC_MENU_EVENT_TYPE_SHOW_REVISIONS_CHANGE
};

// Fill forms

function readSDKContentControl(props, selectedObjects) {
    var type = props.get_SpecificType(),
        internalId = props.get_InternalId(),
        specProps;

    var result = {
        get_InternalId: internalId,
        get_PlaceholderText: props.get_PlaceholderText(),
        get_Lock: props.get_Lock(),
        get_SpecificType: type
    };

    // for list controls
    if (type == Asc.c_oAscContentControlSpecificType.ComboBox || type == Asc.c_oAscContentControlSpecificType.DropDownList) {
        specProps = (type == Asc.c_oAscContentControlSpecificType.ComboBox) ? props.get_ComboBoxPr() : props.get_DropDownListPr();
        if (specProps) {
            var count = specProps.get_ItemsCount();
            var arr = [];
            for (var i = 0; i < count; i++) {
                (specProps.get_ItemValue(i) !== '') && arr.push({
                    value: specProps.get_ItemValue(i),
                    name: specProps.get_ItemDisplayText(i)
                });
            }
            result["values"] = arr;
        }
    } else if (type == Asc.c_oAscContentControlSpecificType.CheckBox) {
        specProps = props.get_CheckBoxPr();
    } else if (type == Asc.c_oAscContentControlSpecificType.Picture) {
        for (i = 0; i < selectedObjects.length; i++) {
            var eltype = selectedObjects[i].get_ObjectType();

            if (eltype === Asc.c_oAscTypeSelectElement.Image) {
                var value = selectedObjects[i].get_ObjectValue();
                if (value.get_ChartProperties() == null && value.get_ShapeProperties() == null) {
                    result["get_ImageUrl"] = value.get_ImageUrl();
                    break;
                }
            }
        }
    } else if (type == Asc.c_oAscContentControlSpecificType.DateTime) {
        specProps = props.get_DateTimePr();
        result["get_FullDate"] = specProps ? specProps.get_FullDate() : null;
    }

    // form settings
    var formPr = props.get_FormPr();
    if (formPr) {
        var data = [];
        if (type == Asc.c_oAscContentControlSpecificType.CheckBox) {
            data = _api.asc_GetCheckBoxFormKeys();
            result["asc_GetCheckBoxFormKeys"] = data;
        } else if (type == Asc.c_oAscContentControlSpecificType.Picture) {
            data = _api.asc_GetPictureFormKeys();
            result["asc_GetPictureFormKeys"] = data;
        } else {
            data = _api.asc_GetTextFormKeys();
            result["asc_GetTextFormKeys"] = data;
        }

        var arr = [];
        data.forEach(function (item) {
            arr.push({ displayValue: item, value: item });
        });

        result["FormKeys"] = arr;

        var val = formPr.get_Key();
        result["get_Key"] = val;

        if (val) {
            val = _api.asc_GetFormsCountByKey(val);
            result["asc_GetFormsCountByKey"] = val;
        }
        
        val = formPr.get_HelpText();
        result["get_HelpText"] = val;

        if (type == Asc.c_oAscContentControlSpecificType.CheckBox && specProps) {
            val = specProps.get_GroupKey();
            result["get_GroupKey"] = val;
            
            var ischeckbox = (typeof val !== 'string');
            result["isCheckBox"] = ischeckbox;

            val = specProps.get_Checked();
            result["get_Checked"] = val;

            val = _api.asc_IsContentControlCheckBoxChecked(internalId);
            result["asc_IsContentControlCheckBoxChecked"] = val;

            
            if (!ischeckbox) {
                data = _api.asc_GetRadioButtonGroupKeys();
                result["asc_GetRadioButtonGroupKeys"] = data;

                var arr = [];
                data.forEach(function(item) {
                    arr.push({ displayValue: item,  value: item });
                });

                result["GroupKeys"] = arr;
            }
        }

        var formTextPr = props.get_TextFormPr();
        if (formTextPr) {
            val = formTextPr.get_Comb();
            result["get_Comb"] = val;

            val = formTextPr.get_Width();
            result["get_Width"] = val;
            
            val = _api.asc_GetTextFormAutoWidth();
            result["asc_GetTextFormAutoWidth"] = val;
            
            val = formTextPr.get_MaxCharacters();
            result["get_MaxCharacters"] = val;
        }
    }

    return result;
}

function onShowContentControlsActions(obj, x, y) {
    var type = obj.type;
    var data = {
        "x": x,
        "y": y,
        "type": type
    };
    var contentControllJSON = {};

    switch (type) {
        case Asc.c_oAscContentControlSpecificType.DateTime:
            contentControllJSON = contentControllDateTimeToJSON(obj);
            break;
        case Asc.c_oAscContentControlSpecificType.Picture:
            if (obj.pr && obj.pr.get_Lock) {
                var lock = obj.pr.get_Lock();
                if (lock == Asc.c_oAscSdtLockType.SdtContentLocked || lock == Asc.c_oAscSdtLockType.ContentLocked)
                    return;
            }
            break;
        case Asc.c_oAscContentControlSpecificType.DropDownList:
        case Asc.c_oAscContentControlSpecificType.ComboBox:
            contentControllJSON = contentControllListAToJSON(obj);
            break;
    }

    // merge
    for(var key in contentControllJSON) {
        data[key] = contentControllJSON[key];
    }

    postDataAsJSONString(data, 26001); // ASC_MENU_EVENT_TYPE_SHOW_CONTENT_CONTROLS_ACTIONS
}

function onHideContentControlsActions() {
    postDataAsJSONString(null, 26002); // ASC_MENU_EVENT_TYPE_HIDE_CONTENT_CONTROLS_ACTIONS
}

function contentControllDateTimeToJSON(obj) {
    var props = obj.pr,
        specProps = props.get_DateTimePr();

    return {
        date: specProps ? specProps.get_FullDate() : null
    }
}

function contentControllListAToJSON(obj) {
    var type = obj.type,
        props = obj.pr,
        specProps = (type == Asc.c_oAscContentControlSpecificType.ComboBox) ? props.get_ComboBoxPr() : props.get_DropDownListPr(),
        isForm = !!props.get_FormPr(),
        internalId = props.get_InternalId()
        items = [];

    if (specProps) {
        if (isForm) { // for dropdown and combobox form control always add placeholder item
            var text = props.get_PlaceholderText();
            items.push({
                caption: text,
                value: ''
            });
        }
        var count = specProps.get_ItemsCount();
        for (var i = 0; i < count; i++) {
            (specProps.get_ItemValue(i) !== '' || !isForm) && items.push({
                caption: specProps.get_ItemDisplayText(i),
                value: specProps.get_ItemValue(i)
            });
        }
        if (!isForm && menu.items.length < 1) {
            items.push({
                caption: '',
                value: '-1'
            });
        }
    }

    return {
        internalId: internalId,
        isForm: isForm,
        items: items
    }
}

// Common

function onFocusObject(SelectedObjects, localTrigger) {
    var settings = [];
    var control_props = _api.asc_IsContentControl() ? _api.asc_GetContentControlProperties() : null;

    for (var i = 0; i < SelectedObjects.length; i++) {
        var eltype = SelectedObjects[i].get_ObjectType();
        var value = SelectedObjects[i].get_ObjectValue();
        
        switch (eltype)
        {
            case Asc.c_oAscTypeSelectElement.Paragraph:
            case Asc.c_oAscTypeSelectElement.Header:
            case Asc.c_oAscTypeSelectElement.Table:
            case Asc.c_oAscTypeSelectElement.Image:
            case Asc.c_oAscTypeSelectElement.Hyperlink:
            case Asc.c_oAscTypeSelectElement.Math:
            {
                settings.push({
                    type: eltype,
                    localTrigger: (typeof localTrigger === 'boolean') ? localTrigger : true,
                    rawValue: JSON.prune(value, 5)
                });
                break;
            }
            case Asc.c_oAscTypeSelectElement.SpellCheck:
            default:
            {
                break;
            }
        }
    }

    // Form object
    if (control_props) {
        var spectype = control_props.get_SpecificType();
        settings.push({
            type: Asc.c_oAscTypeSelectElement.ContentControl,
            spectype: spectype,
            localTrigger: (typeof localTrigger === 'boolean') ? localTrigger : true,
            rawValue: JSON.prune(control_props, 4),
            value: readSDKContentControl(control_props, SelectedObjects)
        });
    }

    postDataAsJSONString(settings, 26101); // ASC_MENU_EVENT_TYPE_FOCUS_OBJECT
}

// Others

var DocumentPageSize = new function()
{
    this.oSizes    = [{name : "US Letter", w_mm : 215.9, h_mm : 279.4, w_tw : 12240, h_tw : 15840},
                      {name : "US Legal", w_mm : 215.9, h_mm : 355.6, w_tw : 12240, h_tw : 20160},
                      {name : "A4", w_mm : 210, h_mm : 297, w_tw : 11907, h_tw : 16839},
                      {name : "A5", w_mm : 148.1, h_mm : 209.9, w_tw : 8391, h_tw : 11907},
                      {name : "B5", w_mm : 176, h_mm : 250.1, w_tw : 9979, h_tw : 14175},
                      {name : "Envelope #10", w_mm : 104.8, h_mm : 241.3, w_tw : 5940, h_tw : 13680},
                      {name : "Envelope DL", w_mm : 110.1, h_mm : 220.1, w_tw : 6237, h_tw : 12474},
                      {name : "Tabloid", w_mm : 279.4, h_mm : 431.7, w_tw : 15842, h_tw : 24477},
                      {name : "A3", w_mm : 297, h_mm : 420.1, w_tw : 16840, h_tw : 23820},
                      {name : "Tabloid Oversize", w_mm : 304.8, h_mm : 457.1, w_tw : 17282, h_tw : 25918},
                      {name : "ROC 16K", w_mm : 196.8, h_mm : 273, w_tw : 11164, h_tw : 15485},
                      {name : "Envelope Coukei 3", w_mm : 119.9, h_mm : 234.9, w_tw : 6798, h_tw : 13319},
                      {name : "Super B/A3", w_mm : 330.2, h_mm : 482.5, w_tw : 18722, h_tw : 27358}
                      ];
    this.sizeEpsMM = 0.5;
    this.getSize   = function(widthMm, heightMm)
    {
        for (var index in this.oSizes)
        {
            var item = this.oSizes[index];
            if (Math.abs(widthMm - item.w_mm) < this.sizeEpsMM && Math.abs(heightMm - item.h_mm) < this.sizeEpsMM)
                return item;
        }
        return {w_mm : widthMm, h_mm : heightMm};
    };
};

Asc["asc_docs_api"].prototype["asc_nativeOpenFile2"] = function(base64File, version)
{
    this.SpellCheckUrl = '';

    this.WordControl.m_bIsRuler = false;
    this.WordControl.Init();

    this.InitEditor();
    this.DocumentType   = 2;
    this.LoadedObjectDS = this.WordControl.m_oLogicDocument.CopyStyle();

    AscCommon.g_oIdCounter.Set_Load(true);

    var openParams        = {checkFileSize : /*this.isMobileVersion*/false, charCount : 0, parCount : 0};
    var oBinaryFileReader = new AscCommonWord.BinaryFileReader(this.WordControl.m_oLogicDocument, openParams);

    if (undefined === version)
    {
        if (oBinaryFileReader.Read(base64File))
        {
            AscCommon.g_oIdCounter.Set_Load(false);
            this.LoadedObject = 1;

            this.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Open);

            // проверяем какие шрифты нужны
            //this.WordControl.m_oDrawingDocument.CheckFontNeeds();
            AscCommon.pptx_content_loader.CheckImagesNeeds(this.WordControl.m_oLogicDocument);

            //this.FontLoader.LoadDocumentFonts(this.WordControl.m_oLogicDocument.Fonts, false);
        }
        else
            this.sendEvent("asc_onError", Asc.c_oAscError.ID.MobileUnexpectedCharCount, Asc.c_oAscError.Level.Critical);
    }
    else
    {
        AscCommon.CurFileVersion = version;
        if (oBinaryFileReader.Read(base64File))
        {
            AscCommon.g_oIdCounter.Set_Load(false);
            this.LoadedObject = 1;

            this.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Open);
        }
        else
            this.sendEvent("asc_onError", Asc.c_oAscError.ID.MobileUnexpectedCharCount, Asc.c_oAscError.Level.Critical);
    }

    /*
     if (window["NATIVE_EDITOR_ENJINE"] === true && undefined != window["native"])
     {
     AscCommon.CDocsCoApi.prototype.askSaveChanges = function(callback)
     {
     callback({"saveLock" : false});
     };
     AscCommon.CDocsCoApi.prototype.saveChanges    = function(arrayChanges, deleteIndex, excelAdditionalInfo)
     {
     if (window["native"]["SaveChanges"])
     window["native"]["SaveChanges"](arrayChanges.join("\",\""), deleteIndex, arrayChanges.length);
     };
     }
     */

    if (undefined != window["Native"])
        return;

    //callback
    this.DocumentOrientation = (null == editor.WordControl.m_oLogicDocument) ? true : !editor.WordControl.m_oLogicDocument.Orientation;
    var sizeMM;
    if (this.DocumentOrientation)
        sizeMM = DocumentPageSize.getSize(AscCommon.Page_Width, AscCommon.Page_Height);
    else
        sizeMM = DocumentPageSize.getSize(AscCommon.Page_Height, AscCommon.Page_Width);
    this.sync_DocSizeCallback(sizeMM.w_mm, sizeMM.h_mm);
    this.sync_PageOrientCallback(editor.get_DocumentOrientation());

    if (this.GenerateNativeStyles !== undefined)
    {
        this.GenerateNativeStyles();

        if (this.WordControl.m_oDrawingDocument.CheckTableStylesOne !== undefined)
            this.WordControl.m_oDrawingDocument.CheckTableStylesOne();
    }
};

Asc['asc_docs_api'].prototype.openDocument = function(file)
{
    _api.asc_nativeOpenFile2(file.data);

    if (window.documentInfo["viewmode"]) {
        _api.ShowParaMarks = false;
        _api.isViewMode = true;
        _api.WordControl.m_oDrawingDocument.IsViewMode = true;
    }    
    
    if (!sdkCheck) {
        
        console.log("OPEN FILE ONLINE READ MODE");
       
        _api["NativeAfterLoad"]();
     
        this.ImageLoader.bIsLoadDocumentFirst = true;
        
        if (null != _api.WordControl.m_oLogicDocument)
        {
            _api.WordControl.m_oDrawingDocument.CheckGuiControlColors();
            _api.sendColorThemes(_api.WordControl.m_oLogicDocument.theme);
        }
  
        window["native"]["onEndLoadingFile"]();
        
        return;
    }

    var version;
    if (file.changes && this.VersionHistory)
    {
        this.VersionHistory.changes = file.changes;
        this.VersionHistory.applyChanges(this);
    }
       
    _api["NativeAfterLoad"]();

    //console.log("ImageMap : " + JSON.stringify(this.WordControl.m_oLogicDocument));

    this.ImageLoader.bIsLoadDocumentFirst = true;
    this.ImageLoader.LoadDocumentImages(this.WordControl.m_oLogicDocument.ImageMap);

    this.WordControl.m_oLogicDocument.Continue_FastCollaborativeEditing();

    //this.asyncFontsDocumentEndLoaded();

    if (null != _api.WordControl.m_oLogicDocument)
    {
        _api.WordControl.m_oDrawingDocument.CheckGuiControlColors();
        _api.sendColorThemes(_api.WordControl.m_oLogicDocument.theme);
    }

    window["native"]["onTokenJWT"](_api.CoAuthoringApi.get_jwt());
    window["native"]["onEndLoadingFile"]();

    this.WordControl.m_oDrawingDocument.Collaborative_TargetsUpdate(true);

    initSpellCheckApi();
    initTrackRevisions();

    var t = this;
    setInterval(function () {
        t._autoSave();
    }, 40);
};

Asc["asc_docs_api"].prototype["asc_nativeGetFileData"] = function()
{
    var oBinaryFileWriter = new AscCommonWord.BinaryFileWriter(this.WordControl.m_oLogicDocument);
    var memory = oBinaryFileWriter.memory;

    oBinaryFileWriter.Write(true);

    window["native"]["GetFileData"](memory.ImData.data, memory.GetCurPosition());

    return true;
};

Asc['asc_docs_api'].prototype.asc_setSpellCheck = function(isOn)
{
    if (this.WordControl && this.WordControl.m_oLogicDocument)
    {
        var oLogicDoc = this.WordControl.m_oLogicDocument;
        if(isOn)
        {
            this.spellCheckTimerId = setInterval(function(){oLogicDoc.ContinueSpellCheck();}, 500);
        }
        else
        {
            if(this.spellCheckTimerId)
            {
               clearInterval(this.spellCheckTimerId);
            }
        }
        editor.WordControl.m_oLogicDocument.Spelling.Use = isOn;
        editor.WordControl.m_oDrawingDocument.ClearCachePages();
        editor.WordControl.m_oDrawingDocument.FirePaint();
    }
};

// The helper function, called from the native application,
// returns information about the document as a JSON string.
Asc["asc_docs_api"].prototype["asc_nativeGetCoreProps"] = function() {
    var props = (_api) ? _api.asc_getCoreProps() : null,
        value;

    if (props) {
        var coreProps = {};
        coreProps["asc_getModified"] = props.asc_getModified();

        value = props.asc_getLastModifiedBy();
        if (value)
        coreProps["asc_getLastModifiedBy"] = AscCommon.UserInfoParser.getParsedName(value);

        coreProps["asc_getTitle"] = props.asc_getTitle();
        coreProps["asc_getSubject"] = props.asc_getSubject();
        coreProps["asc_getDescription"] = props.asc_getDescription();
        coreProps["asc_getCreated"] = props.asc_getCreated();

        var authors = [];
        value = props.asc_getCreator();//"123\"\"\"\<\>,456";
        value && value.split(/\s*[,;]\s*/).forEach(function (item) {
            authors.push(item);
        });

        coreProps["asc_getCreator"] = authors;

        return coreProps;
    }

    return {};
}

// The helper function, wrap of asc_SetContentControlDatePickerDate
Asc["asc_docs_api"].prototype["asc_nativeSetContentControlDatePickerDate"] = function(textDate, sId) {
    var oLogicDocument = this.WordControl.m_oLogicDocument;
    if (!oLogicDocument)
        return;

    var oContentControl = oLogicDocument.GetContentControl(sId);
    if (!oContentControl || !oContentControl.IsDatePicker() || !oContentControl.CanBeEdited())
        return;

    var oPr = oContentControl.GetContentControlPr().get_DateTimePr();
    oPr.put_FullDate(new  Date(textDate));

    _api.asc_SetContentControlDatePickerPr(oPr, sId, true);
}

Asc["asc_docs_api"].prototype["asc_nativeAddText"] = function(text, wrapWithSpaces) {
    var settings = new AscCommon.CAddTextSettings();

    if (wrapWithSpaces) {
        settings.SetWrapWithSpaces(true);
    }
    
    _api.asc_AddText(text, settings);
}

Asc["asc_docs_api"].prototype["asc_nativeGetDocumentProtection"] = function() {
    var props = (_api) ? _api.asc_getDocumentProtection() : null;
    if (props) {
        return {
            "asc_getEditType": props.asc_getEditType(),
            "asc_getIsPassword": props.asc_getIsPassword()
        }
    }
    return {};
}

window["AscCommon"].getFullImageSrc2 = function(src) {
    var start = src.slice(0, 6);
    if (0 === start.indexOf("theme") && editor.ThemeLoader) {
      return editor.ThemeLoader.ThemesUrlAbs + src;
    }
    if (0 !== start.indexOf("http:") && 0 !== start.indexOf("data:") && 0 !== start.indexOf("https:") && 0 !== start.indexOf("file:") && 0 !== start.indexOf("ftp:")) {
      var srcFull = AscCommon.g_oDocumentUrls.getImageUrl(src);
      var srcFull2 = srcFull;
      if (src.endsWith(".svg")) {
        var sName = src.slice(0, src.length - 3);
        src = sName + "wmf";
        srcFull = AscCommon.g_oDocumentUrls.getImageUrl(src);
        if (!srcFull) {
          src = sName + "emf";
          srcFull = AscCommon.g_oDocumentUrls.getImageUrl(src);
        }
      }
      if (srcFull) {
        window["native"]["loadUrlImage"](srcFull, src);
        return srcFull2;
      }
    }
    return src;
  };

window["AscCommon"].sendImgUrls = function(api, images, callback)
{
	var _data = [];
	callback(_data);
};

window["native"]["offline_of"] = function(_params, documentInfo) {NativeOpenFile3(_params, documentInfo);};

window["GetNativePageMeta"] = function(pageIndex)
{
    return window["API"].GetNativePageMeta(pageIndex);
}

window.native.Call_CalculateResume = function ()
{
    return window["API"].Call_CalculateResume();
};

window.native.Call_TurnOffRecalculate = function ()
{
    return window["API"].Call_TurnOffRecalculate();
};
window.native.Call_TurnOnRecalculate = function ()
{
    return window["API"].Call_TurnOnRecalculate();
};

window.native.Call_CheckTargetUpdate = function ()
{
    return window["API"].Call_CheckTargetUpdate();
};
window.native.Call_Common = function (type, param)
{
    return window["API"].Call_Common(type, param);
};

window.native.Call_HR_Tabs = function (arrT, arrP)
{
    return window["API"].Call_HR_Tabs(arrT, arrP);
};
window.native.Call_HR_Pr = function (_indent_left, _indent_right, _indent_first)
{
    return window["API"].Call_HR_Pr(_indent_left, _indent_right, _indent_first);
};
window.native.Call_HR_Margins = function (_margin_left, _margin_right)
{
    return window["API"].Call_HR_Margins(_margin_left, _margin_right);
};
window.native.Call_HR_Table = function (_params, _cols, _margins, _rows)
{
    return window["API"].Call_HR_Table(_params, _cols, _margins, _rows);
};

window.native.Call_VR_Margins = function (_top, _bottom)
{
    return window["API"].Call_VR_Margins(_top, _bottom);
};
window.native.Call_VR_Header = function (_header_top, _header_bottom)
{
    return window["API"].Call_VR_Header(_header_top, _header_bottom);
};
window.native.Call_VR_Table = function (_params, _cols, _margins, _rows)
{
    return window["API"].Call_VR_Table(_params, _cols, _margins, _rows);
};

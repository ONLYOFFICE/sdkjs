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

// Import
var CColor                   = AscCommon.CColor;
var g_oTextMeasurer          = AscCommon.g_oTextMeasurer;
var global_keyboardEvent     = AscCommon.global_keyboardEvent;
var global_mouseEvent        = AscCommon.global_mouseEvent;
var History                  = AscCommon.History;
var global_MatrixTransformer = AscCommon.global_MatrixTransformer;
var g_dKoef_pix_to_mm        = AscCommon.g_dKoef_pix_to_mm;
var g_dKoef_mm_to_pix        = AscCommon.g_dKoef_mm_to_pix;

var FontStyle         = AscFonts.FontStyle;
var g_fontApplication = AscFonts.g_fontApplication;
var ImageLoadStatus   = AscFonts.ImageLoadStatus;

var FOCUS_OBJECT_THUMBNAILS = 0;
var FOCUS_OBJECT_MAIN       = 1;
var FOCUS_OBJECT_NOTES      = 2;
var FOCUS_OBJECT_ANIM_PANE  = 3;

AscCommon.FOCUS_OBJECT_THUMBNAILS = FOCUS_OBJECT_THUMBNAILS;
AscCommon.FOCUS_OBJECT_MAIN       = FOCUS_OBJECT_MAIN;
AscCommon.FOCUS_OBJECT_NOTES      = FOCUS_OBJECT_NOTES;
AscCommon.FOCUS_OBJECT_ANIM_PANE  = FOCUS_OBJECT_ANIM_PANE;

var COMMENT_WIDTH  = 18;
var COMMENT_HEIGHT = 16;

const LAYOUT_SCALE = 0.87;

// text measurer wrapper
AscCommon.CTextMeasurer.prototype.GetAscender  = function()
{
	var UnitsPerEm = this.m_oManager.m_lUnits_Per_Em;
	//var Ascender   = this.m_oManager.m_lAscender;
	var Ascender   = ( 0 !== this.m_oManager.m_lLineHeight ) ? 1.2 * UnitsPerEm * this.m_oManager.m_lAscender / this.m_oManager.m_lLineHeight : this.m_oManager.m_lAscender;

	return Ascender * this.m_oLastFont.SetUpSize / UnitsPerEm * g_dKoef_pt_to_mm;
};
AscCommon.CTextMeasurer.prototype.GetDescender = function()
{
	var UnitsPerEm = this.m_oManager.m_lUnits_Per_Em;
	//var Descender  = this.m_oManager.m_lDescender;
	var Descender  = ( 0 !== this.m_oManager.m_lLineHeight ) ? 1.2 * UnitsPerEm * this.m_oManager.m_lDescender / this.m_oManager.m_lLineHeight : this.m_oManager.m_lDescender;

	return Descender * this.m_oLastFont.SetUpSize / UnitsPerEm * g_dKoef_pt_to_mm;
};
AscCommon.CTextMeasurer.prototype.GetHeight    = function()
{
	var UnitsPerEm = this.m_oManager.m_lUnits_Per_Em;
	//var Height     = this.m_oManager.m_lLineHeight;
	var Height     = ( 0 !== this.m_oManager.m_lLineHeight ) ? 1.2 * UnitsPerEm : this.m_oManager.m_lLineHeight;

	return Height * this.m_oLastFont.SetUpSize / UnitsPerEm * g_dKoef_pt_to_mm;
};

function CTableOutlineDr()
{
	this.TableOutline = null;
	this.Counter = 0;
	this.bIsNoTable = true;
	this.bIsTracked = false;

	this.CurPos = null;
	this.TrackTablePos = 0; // 0 - left_top, 1 - right_top, 2 - right_bottom, 3 - left_bottom
	this.TrackOffsetX = 0;
	this.TrackOffsetY = 0;

	this.InlinePos = null;

	this.IsChangeSmall = true;
	this.ChangeSmallPoint = null;

	this.TableMatrix = null;
	this.CurrentPageIndex = null;

	this.checkMouseDown = function (pos, word_control)
	{
		if (null == this.TableOutline)
			return false;

		var _table_track = this.TableOutline;
		var _d = 13 * g_dKoef_pix_to_mm * 100 / word_control.m_nZoomValue;

		this.IsChangeSmall = true;
		this.ChangeSmallPoint = pos;

		this.CurPos = {X: this.ChangeSmallPoint.X, Y: this.ChangeSmallPoint.Y, Page: this.ChangeSmallPoint.Page};
		this.TrackOffsetX = 0;
		this.TrackOffsetY = 0;

		if (!this.TableMatrix || global_MatrixTransformer.IsIdentity(this.TableMatrix))
		{
			if (word_control.MobileTouchManager)
			{
				var _move_point = word_control.MobileTouchManager.TableMovePoint;

				if (_move_point == null || pos.Page != _table_track.PageNum)
					return false;

				var _pos1 = word_control.m_oDrawingDocument.ConvertCoordsToCursorWR(pos.X, pos.Y, pos.Page);
				var _pos2 = word_control.m_oDrawingDocument.ConvertCoordsToCursorWR(_move_point.X, _move_point.Y, pos.Page);

				var _eps = word_control.MobileTouchManager.TrackTargetEps;

				var _offset1 = word_control.MobileTouchManager.TableRulersRectOffset;
				var _offset2 = _offset1 + word_control.MobileTouchManager.TableRulersRectSize;
				if ((_pos1.X >= (_pos2.X - _offset2 - _eps)) && (_pos1.X <= (_pos2.X - _offset1 + _eps)) &&
					(_pos1.Y >= (_pos2.Y - _offset2 - _eps)) && (_pos1.Y <= (_pos2.Y - _offset1 + _eps)))
				{
					this.TrackTablePos = 0;
					return true;
				}

				return false;
			}

			switch (this.TrackTablePos)
			{
				case 1:
				{
					var _x = _table_track.X + _table_track.W;
					var _b = _table_track.Y;
					var _y = _b - _d;
					var _r = _x + _d;

					if ((pos.X > _x) && (pos.X < _r) && (pos.Y > _y) && (pos.Y < _b))
					{
						this.TrackOffsetX = pos.X - _x;
						this.TrackOffsetY = pos.Y - _b;

						this.CurPos.X -= this.TrackOffsetX;
						this.CurPos.Y -= this.TrackOffsetY;
						return true;
					}
					return false;
				}
				case 2:
				{
					var _x = _table_track.X + _table_track.W;
					var _y = _table_track.Y + _table_track.H;
					var _r = _x + _d;
					var _b = _y + _d;

					if ((pos.X > _x) && (pos.X < _r) && (pos.Y > _y) && (pos.Y < _b))
					{
						this.TrackOffsetX = pos.X - _x;
						this.TrackOffsetY = pos.Y - _y;
						return true;
					}
					return false;
				}
				case 3:
				{
					var _r = _table_track.X;
					var _x = _r - _d;
					var _y = _table_track.Y + _table_track.H;
					var _b = _y + _d;

					if ((pos.X > _x) && (pos.X < _r) && (pos.Y > _y) && (pos.Y < _b))
					{
						this.TrackOffsetX = pos.X - _r;
						this.TrackOffsetY = pos.Y - _y;

						this.CurPos.X -= this.TrackOffsetX;
						this.CurPos.Y -= this.TrackOffsetY;
						return true;
					}
					return false;
				}
				case 0:
				default:
				{
					var _r = _table_track.X;
					var _b = _table_track.Y;
					var _x = _r - _d;
					var _y = _b - _d;

					if ((pos.X > _x) && (pos.X < _r) && (pos.Y > _y) && (pos.Y < _b))
					{
						this.TrackOffsetX = pos.X - _r;
						this.TrackOffsetY = pos.Y - _b;

						this.CurPos.X -= this.TrackOffsetX;
						this.CurPos.Y -= this.TrackOffsetY;
						return true;
					}
					return false;
				}
			}
		}
		else
		{
			if (word_control.MobileTouchManager)
			{
				var _invert = global_MatrixTransformer.Invert(this.TableMatrix);
				var _posx = _invert.TransformPointX(pos.X, pos.Y);
				var _posy = _invert.TransformPointY(pos.X, pos.Y);

				var _move_point = word_control.MobileTouchManager.TableMovePoint;

				if (_move_point == null || pos.Page != _table_track.PageNum)
					return false;

				var _koef = g_dKoef_pix_to_mm * 100 / word_control.m_nZoomValue;
				var _eps = word_control.MobileTouchManager.TrackTargetEps * _koef;

				var _offset1 = word_control.MobileTouchManager.TableRulersRectOffset * _koef;
				var _offset2 = _offset1 + word_control.MobileTouchManager.TableRulersRectSize * _koef;
				if ((_posx >= (_move_point.X - _offset2 - _eps)) && (_posx <= (_move_point.X - _offset1 + _eps)) &&
					(_posy >= (_move_point.Y - _offset2 - _eps)) && (_posy <= (_move_point.Y - _offset1 + _eps)))
				{
					this.TrackTablePos = 0;
					return true;
				}

				return false;
			}

			var _invert = global_MatrixTransformer.Invert(this.TableMatrix);
			var _posx = _invert.TransformPointX(pos.X, pos.Y);
			var _posy = _invert.TransformPointY(pos.X, pos.Y);
			switch (this.TrackTablePos)
			{
				case 1:
				{
					var _x = _table_track.X + _table_track.W;
					var _b = _table_track.Y;
					var _y = _b - _d;
					var _r = _x + _d;

					if ((_posx > _x) && (_posx < _r) && (_posy > _y) && (_posy < _b))
					{
						this.TrackOffsetX = _posx - _x;
						this.TrackOffsetY = _posy - _b;

						this.CurPos.X -= this.TrackOffsetX;
						this.CurPos.Y -= this.TrackOffsetY;
						return true;
					}
					return false;
				}
				case 2:
				{
					var _x = _table_track.X + _table_track.W;
					var _y = _table_track.Y + _table_track.H;
					var _r = _x + _d;
					var _b = _y + _d;

					if ((_posx > _x) && (_posx < _r) && (_posy > _y) && (_posy < _b))
					{
						this.TrackOffsetX = _posx - _x;
						this.TrackOffsetY = _posy - _y;

						this.CurPos.X -= this.TrackOffsetX;
						this.CurPos.Y -= this.TrackOffsetY;
						return true;
					}
					return false;
				}
				case 3:
				{
					var _r = _table_track.X;
					var _x = _r - _d;
					var _y = _table_track.Y + _table_track.H;
					var _b = _y + _d;

					if ((_posx > _x) && (_posx < _r) && (_posy > _y) && (_posy < _b))
					{
						this.TrackOffsetX = _posx - _r;
						this.TrackOffsetY = _posy - _y;

						this.CurPos.X -= this.TrackOffsetX;
						this.CurPos.Y -= this.TrackOffsetY;
						return true;
					}
					return false;
				}
				case 0:
				default:
				{
					var _r = _table_track.X;
					var _b = _table_track.Y;
					var _x = _r - _d;
					var _y = _b - _d;

					if ((_posx > _x) && (_posx < _r) && (_posy > _y) && (_posy < _b))
					{
						this.TrackOffsetX = _posx - _r;
						this.TrackOffsetY = _posy - _b;

						this.CurPos.X -= this.TrackOffsetX;
						this.CurPos.Y -= this.TrackOffsetY;
						return true;
					}
					return false;
				}
			}
		}

		return false;
	};

	this.checkMouseUp = function (X, Y, word_control)
	{
		this.bIsTracked = false;

		if (null == this.TableOutline || (true === this.IsChangeSmall) || word_control.m_oApi.isViewMode)
			return false;

		var _d = 13 * g_dKoef_pix_to_mm * 100 / word_control.m_nZoomValue;

		var _outline = this.TableOutline;
		var _table = _outline.Table;

		_table.MoveCursorToStartPos();
		_table.Document_SetThisElementCurrent(true);

		if (!_table.Is_Inline())
		{
			var pos;
			switch (this.TrackTablePos)
			{
				case 1:
				{
					var _w_pix = this.TableOutline.W * g_dKoef_mm_to_pix * word_control.m_nZoomValue / 100;
					pos = word_control.m_oDrawingDocument.ConvertCoordsFromCursor2(X - _w_pix, Y);
					break;
				}
				case 2:
				{
					var _w_pix = this.TableOutline.W * g_dKoef_mm_to_pix * word_control.m_nZoomValue / 100;
					var _h_pix = this.TableOutline.H * g_dKoef_mm_to_pix * word_control.m_nZoomValue / 100;
					pos = word_control.m_oDrawingDocument.ConvertCoordsFromCursor2(X - _w_pix, Y - _h_pix);
					break;
				}
				case 3:
				{
					var _h_pix = this.TableOutline.H * g_dKoef_mm_to_pix * word_control.m_nZoomValue / 100;
					pos = word_control.m_oDrawingDocument.ConvertCoordsFromCursor2(X, Y - _h_pix);
					break;
				}
				case 0:
				default:
				{
					pos = word_control.m_oDrawingDocument.ConvertCoordsFromCursor2(X, Y);
					break;
				}
			}

			var NearestPos = word_control.m_oLogicDocument.Get_NearestPos(pos.Page, pos.X - this.TrackOffsetX, pos.Y - this.TrackOffsetY);
			_table.Move(pos.X - this.TrackOffsetX, pos.Y - this.TrackOffsetY, pos.Page, NearestPos);
			_outline.X = pos.X - this.TrackOffsetX;
			_outline.Y = pos.Y - this.TrackOffsetY;
			_outline.PageNum = pos.Page;
		}
		else
		{
			if (null != this.InlinePos)
			{
				// inline move
				_table.Move(this.InlinePos.X, this.InlinePos.Y, this.InlinePos.Page, this.InlinePos);
			}
		}
	};

	this.checkMouseMove = function (X, Y, word_control)
	{
		if (null == this.TableOutline)
			return false;

		if (true === this.IsChangeSmall)
		{
			var _pos = word_control.m_oDrawingDocument.ConvertCoordsFromCursor2(X, Y);
			var _dist = 15 * g_dKoef_pix_to_mm * 100 / word_control.m_nZoomValue;
			if ((Math.abs(_pos.X - this.ChangeSmallPoint.X) < _dist) && (Math.abs(_pos.Y - this.ChangeSmallPoint.Y) < _dist) && (_pos.Page == this.ChangeSmallPoint.Page))
			{
				this.CurPos = {
					X: this.ChangeSmallPoint.X,
					Y: this.ChangeSmallPoint.Y,
					Page: this.ChangeSmallPoint.Page
				};

				switch (this.TrackTablePos)
				{
					case 1:
					{
						this.CurPos.X -= this.TableOutline.W;
						break;
					}
					case 2:
					{
						this.CurPos.X -= this.TableOutline.W;
						this.CurPos.Y -= this.TableOutline.H;
						break;
					}
					case 3:
					{
						this.CurPos.Y -= this.TableOutline.H;
						break;
					}
					case 0:
					default:
					{
						break;
					}
				}

				this.CurPos.X -= this.TrackOffsetX;
				this.CurPos.Y -= this.TrackOffsetY;
				return;
			}
			this.IsChangeSmall = false;

			this.TableOutline.Table.RemoveSelection();
			this.TableOutline.Table.MoveCursorToStartPos();
			editor.WordControl.m_oLogicDocument.Document_UpdateSelectionState();
		}

		var _d = 13 * g_dKoef_pix_to_mm * 100 / word_control.m_nZoomValue;
		switch (this.TrackTablePos)
		{
			case 1:
			{
				var _w_pix = this.TableOutline.W * g_dKoef_mm_to_pix * word_control.m_nZoomValue / 100;
				this.CurPos = word_control.m_oDrawingDocument.ConvertCoordsFromCursor2(X - _w_pix, Y);
				break;
			}
			case 2:
			{
				var _w_pix = this.TableOutline.W * g_dKoef_mm_to_pix * word_control.m_nZoomValue / 100;
				var _h_pix = this.TableOutline.H * g_dKoef_mm_to_pix * word_control.m_nZoomValue / 100;
				this.CurPos = word_control.m_oDrawingDocument.ConvertCoordsFromCursor2(X - _w_pix, Y - _h_pix);
				break;
			}
			case 3:
			{
				var _h_pix = this.TableOutline.H * g_dKoef_mm_to_pix * word_control.m_nZoomValue / 100;
				this.CurPos = word_control.m_oDrawingDocument.ConvertCoordsFromCursor2(X, Y - _h_pix);
				break;
			}
			case 0:
			default:
			{
				this.CurPos = word_control.m_oDrawingDocument.ConvertCoordsFromCursor2(X, Y);
				break;
			}
		}

		this.CurPos.X -= this.TrackOffsetX;
		this.CurPos.Y -= this.TrackOffsetY;
	};

	this.CheckStartTrack = function (word_control, transform)
	{
		this.TableMatrix = null;
		if (transform)
			this.TableMatrix = transform.CreateDublicate();

		if (!this.TableMatrix || global_MatrixTransformer.IsIdentity(this.TableMatrix))
		{
			var pos = word_control.m_oDrawingDocument.ConvertCoordsToCursor(this.TableOutline.X, this.TableOutline.Y, this.TableOutline.PageNum, true);

			var _x0 = word_control.m_oEditor.AbsolutePosition.L;
			var _y0 = word_control.m_oEditor.AbsolutePosition.T;

			if (pos.X < _x0 && pos.Y < _y0)
			{
				this.TrackTablePos = 2;
			}
			else if (pos.X < _x0)
			{
				this.TrackTablePos = 1;
			}
			else if (pos.Y < _y0)
			{
				this.TrackTablePos = 3;
			}
			else
			{
				this.TrackTablePos = 0;
			}
		}
		else
		{
			var _x = this.TableOutline.X;
			var _y = this.TableOutline.Y;
			var _r = _x + this.TableOutline.W;
			var _b = _y + this.TableOutline.H;

			var x0 = transform.TransformPointX(_x, _y);
			var y0 = transform.TransformPointY(_x, _y);

			var x1 = transform.TransformPointX(_r, _y);
			var y1 = transform.TransformPointY(_r, _y);

			var x2 = transform.TransformPointX(_r, _b);
			var y2 = transform.TransformPointY(_r, _b);

			var x3 = transform.TransformPointX(_x, _b);
			var y3 = transform.TransformPointY(_x, _b);

			var _x0 = word_control.m_oEditor.AbsolutePosition.L * g_dKoef_mm_to_pix;
			var _y0 = word_control.m_oEditor.AbsolutePosition.T * g_dKoef_mm_to_pix;
			var _x1 = word_control.m_oEditor.AbsolutePosition.R * g_dKoef_mm_to_pix;
			var _y1 = word_control.m_oEditor.AbsolutePosition.B * g_dKoef_mm_to_pix;

			var pos0 = word_control.m_oDrawingDocument.ConvertCoordsToCursor(x0, y0, this.TableOutline.PageNum, true);
			if (pos0.X > _x0 && pos0.X < _x1 && pos0.Y > _y0 && pos0.Y < _y1)
			{
				this.TrackTablePos = 0;
				return;
			}

			pos0 = word_control.m_oDrawingDocument.ConvertCoordsToCursor(x1, y1, this.TableOutline.PageNum, true);
			if (pos0.X > _x0 && pos0.X < _x1 && pos0.Y > _y0 && pos0.Y < _y1)
			{
				this.TrackTablePos = 1;
				return;
			}

			pos0 = word_control.m_oDrawingDocument.ConvertCoordsToCursor(x3, y3, this.TableOutline.PageNum, true);
			if (pos0.X > _x0 && pos0.X < _x1 && pos0.Y > _y0 && pos0.Y < _y1)
			{
				this.TrackTablePos = 3;
				return;
			}

			pos0 = word_control.m_oDrawingDocument.ConvertCoordsToCursor(x2, y2, this.TableOutline.PageNum, true);
			if (pos0.X > _x0 && pos0.X < _x1 && pos0.Y > _y0 && pos0.Y < _y1)
			{
				this.TrackTablePos = 2;
				return;
			}

			this.TrackTablePos = 0;
		}
	};
}

function CDrawingPage()
{
	this.left   = 0;
	this.top    = 0;
	this.right  = 0;
	this.bottom = 0;

	this.cachedImage = null;
}


function CDrawingCollaborativeTarget(DrawingDocument)
{
	AscCommon.CDrawingCollaborativeTargetBase.call(this);
	this.Page = -1;
	this.DrawingDocument = DrawingDocument;
}
CDrawingCollaborativeTarget.prototype = Object.create(AscCommon.CDrawingCollaborativeTargetBase.prototype);

CDrawingCollaborativeTarget.prototype.IsFocusOnNotes = function() {
	return this.DrawingDocument.m_oWordControl.m_oLogicDocument.IsFocusOnNotes();
};
CDrawingCollaborativeTarget.prototype.GetZoom = function()
{
	if(this.IsFocusOnNotes())
	{
		return 1.0;
	}
	return this.DrawingDocument.m_oWordControl.m_nZoomValue / 100;
};
CDrawingCollaborativeTarget.prototype.ConvertCoords = function(x, y)
{
	if (this.IsFocusOnNotes())
	{
		var oNotesApi = this.DrawingDocument.m_oWordControl.m_oNotesApi;
		var _offsetX = oNotesApi.OffsetX;
		return { X : (AscCommon.AscBrowser.convertToRetinaValue(_offsetX) + x * g_dKoef_mm_to_pix), Y : (y * g_dKoef_mm_to_pix - oNotesApi.Scroll) };
	}
	return this.DrawingDocument.ConvertCoordsToCursor(x, y);
};
CDrawingCollaborativeTarget.prototype.GetMobileTouchManager = function()
{
	return this.DrawingDocument.m_oWordControl.MobileTouchManager;
};
CDrawingCollaborativeTarget.prototype.GetParentElement = function()
{
	if(this.IsFocusOnNotes())
	{
		return this.DrawingDocument.m_oWordControl.m_oNotesContainer.HtmlElement;
	}
	return this.DrawingDocument.m_oWordControl.m_oMainView.HtmlElement;
};
CDrawingCollaborativeTarget.prototype.CheckPosition = function(_x, _y, _size, _page, _transform)
{
	this.Transform = _transform;
	this.Size = _size;
	this.X = _x;
	this.Y = _y;
	this.Page = _page;
	this.Update();
};
CDrawingCollaborativeTarget.prototype.CheckStyleDisplay = function()
{
	if (this.DrawingDocument.m_oWordControl.m_oApi.isReporterMode)
	{
		this.HtmlElement.style.display = "none";
		return;
	}
	if (this.HtmlElement.style.display != "block")
	{
		this.HtmlElement.style.display = "block";
	}
	if (this.HtmlElement.style.zIndex != 7)
	{
		this.HtmlElement.style.zIndex = 7;
	}
};
CDrawingCollaborativeTarget.prototype.CheckNeedDraw = function()
{
	if (this.Page != this.DrawingDocument.SlideCurrent)
	{
		this.HtmlElement.style.display = "none";
		return false;
	}
	return true;
};

function CDrawingDocument()
{
	AscCommon.g_oHtmlCursor.register(AscCommon.Cursors.MarkerFormat, "14 8", "pointer");
	AscCommon.g_oHtmlCursor.register(AscCommon.Cursors.TableEraser, "8 19", "pointer");

	this.IsLockObjectsEnable = false;

	// presentation
	this.m_oWordControl   = null;
	this.m_oLogicDocument = null;


	this.IsEmptyPresentation = false;

	this.SlideCurrent     = -1;
	this.SlideCurrectRect = new CDrawingPage();


	this.MasterCurrent = -1;
	this.LayoutCurrent = -1;


	this.isTabButtonShow = true;

	// target
	this.m_lTimerTargetId      = -1;
	this.m_dTargetX            = -1;
	this.m_dTargetY            = -1;
	this.m_dTargetSize         = 1;
	this.m_dTargetAscent	   = 0;
	this.TargetHtmlElement     = null;
	this.TargetHtmlElementLeft = 0;
	this.TargetHtmlElementTop  = 0;
	this.TargetHtmlElementOnSlide = true;

	this.UpdateTargetFromPaint = false;
	this.NeedTarget            = true;
	this.TextMatrix            = null;
	this.TargetShowFlag        = false;
	this.TargetShowNeedFlag    = false;

	this.m_lTimerUpdateTargetID = -1;
	this.m_tempX                = 0;
	this.m_tempY                = 0;
	this.m_tempPageIndex        = 0;

	this.TargetCursorColor = {R : 0, G : 0, B : 0};

	this.CollaborativeTargets            = [];
	this.CollaborativeTargetsUpdateTasks = [];

	this.m_sLockedCursorType = "";

	// select
	this.m_bIsSelection = false;
	this.SelectionMatrix = null;

	// search
	this.m_bIsSearching = false;
	this.CurrentSearchNavi = null;

	// tracks
	this.MathTrack = new AscCommon.CMathTrack();

	this.TableOutlineDr      = new CTableOutlineDr();

	this.HorVerAnchors = [];

	this.InlineTextTrackEnabled = false;
	this.InlineTextTrack = null;
	this.InlineTextTrackPage = -1;
	this.InlineTextInNotes = false;

	this.AutoShapesTrack = null;

	// draw
	this.m_oCacheManager = new CCacheManager(true);

	// hit
	this.CanvasHit        = document.createElement('canvas');
	this.CanvasHit.width  = 10;
	this.CanvasHit.height = 10;
	this.CanvasHitContext = this.CanvasHit.getContext('2d');

	// interface ui
	this.GuiControlColorsMap  = null;
	this.IsSendStandartColors = false;

	this.GuiCanvasFillTextureParentId = "";
	this.GuiCanvasFillTexture         = null;
	this.GuiCanvasFillTextureCtx      = null;
	this.LastDrawingUrl               = "";

	this.GuiCanvasFillTextureParentIdSlide = "";
	this.GuiCanvasFillTextureSlide         = null;
	this.GuiCanvasFillTextureCtxSlide      = null;
	this.LastDrawingUrlSlide               = "";

	this.GuiCanvasFillTextureParentIdTextArt = "";
	this.GuiCanvasFillTextureTextArt         = null;
	this.GuiCanvasFillTextureCtxTextArt      = null;
	this.LastDrawingUrlTextArt               = "";

	this.TableStylesLastTheme = null;
	this.TableStylesLastColorScheme = null;
	this.TableStylesLastColorMap = null;
	this.TableStylesLastLook = null;

	// transitions
	this.TransitionSlide = new CTransitionAnimation(null);

	//all images on start
	this.LoadingImages = [];

	// notes
	this.isDrawingNotes = false;

	// metafile
	this.m_lCurrentRendererPage = -1;
	this.m_oDocRenderer         = null;
	this.m_bOldShowMarks        = false;

    // placeholders
    this.placeholders = new AscCommon.DrawingPlaceholders(this);

	var oThis = this;

	// methods ---

	// target


	this.GetSlidesCount = function()
	{
		return Asc.editor.getCountSlides();
	};
	this.MoveTargetInInputContext = function()
	{
		if (AscCommon.g_inputContext)
			AscCommon.g_inputContext.move(this.TargetHtmlElementLeft, this.TargetHtmlElementTop);
	};
	this.GetTargetColor = function(isFocusOnSlide)
	{
		if (false !== isFocusOnSlide)
			return this.TargetCursorColor;

		// check dark theme
		if (AscCommon.GlobalSkin.Type !== "dark" || (this.TargetCursorColor.R > 10 || this.TargetCursorColor.R > 10 || this.TargetCursorColor.R > 10))
			return this.TargetCursorColor;

		return { R : 255 - this.TargetCursorColor.R, G : 255 - this.TargetCursorColor.G, B : 255 - this.TargetCursorColor.B };
	};
	this.SetTargetColor = function(r, g, b)
	{
		this.TargetCursorColor.R = r;
		this.TargetCursorColor.G = g;
		this.TargetCursorColor.B = b;
	};

	this.SetCursorType     = function(sType, Data)
	{
		let elem = this.m_oWordControl.m_oMainContent.HtmlElement;
		if (this.m_oWordControl.DemonstrationManager.Mode)
			elem = this.m_oWordControl.DemonstrationManager.Canvas;

		if ("" === this.m_sLockedCursorType)
		{
			elem.style.cursor = AscCommon.g_oHtmlCursor.value(sType);
		}
		else
			elem.style.cursor = AscCommon.g_oHtmlCursor.value(this.m_sLockedCursorType);

		if ("undefined" === typeof(Data) || null === Data)
			Data = new AscCommon.CMouseMoveData();

		editor.sync_MouseMoveCallback(Data);
	};
	this.LockCursorType    = function(sType)
	{
		this.m_sLockedCursorType                                    = sType;
		this.m_oWordControl.m_oMainContent.HtmlElement.style.cursor = AscCommon.g_oHtmlCursor.value(this.m_sLockedCursorType);
	};
	this.LockCursorTypeCur = function()
	{
		this.m_sLockedCursorType = this.m_oWordControl.m_oMainContent.HtmlElement.style.cursor;
	};
	this.UnlockCursorType  = function()
	{
		this.m_sLockedCursorType = "";
	};

	this.TargetStart          = function()
	{
		if (this.m_lTimerTargetId != -1)
			clearInterval(this.m_lTimerTargetId);
		this.m_lTimerTargetId = setInterval(oThis.DrawTarget, 500);
	};
	this.TargetEnd            = function()
	{
		if (!this.TargetShowFlag)
			return;

		this.TargetShowFlag     = false;
		this.TargetShowNeedFlag = false;
		clearInterval(this.m_lTimerTargetId);
		this.m_lTimerTargetId                = -1;
		this.TargetHtmlElement.style.display = "none";
	};
	this.UpdateTargetNoAttack = function()
	{
		if (null == this.m_oWordControl)
			return;

		this.CheckTargetDraw(this.m_dTargetX, this.m_dTargetY, !this.m_oLogicDocument.IsFocusOnNotes());
	};

	this.CheckTargetDraw = function(x, y, isFocusOnSlide)
	{
		var isReporter = this.m_oWordControl.m_oApi.isReporterMode;
		if (this.TargetHtmlElementOnSlide != isFocusOnSlide)
		{
			if (this.TargetHtmlElementOnSlide)
			{
				this.m_oWordControl.m_oMainView.HtmlElement.removeChild(this.TargetHtmlElement);
				this.m_oWordControl.m_oNotesContainer.HtmlElement.appendChild(this.TargetHtmlElement);
				this.TargetHtmlElement.style.zIndex = isReporter ? 0 : 9;

				AscCommon.g_inputContext.TargetOffsetY = (this.m_oWordControl.m_oNotesContainer.AbsolutePosition.T * AscCommon.g_dKoef_mm_to_pix) >> 0;
			}
			else
			{
				this.m_oWordControl.m_oNotesContainer.HtmlElement.removeChild(this.TargetHtmlElement);
				this.m_oWordControl.m_oMainView.HtmlElement.appendChild(this.TargetHtmlElement);
				this.TargetHtmlElement.style.zIndex = isReporter ? 0 : 9;

				AscCommon.g_inputContext.TargetOffsetY = 0;
			}

			this.TargetHtmlElementOnSlide = isFocusOnSlide;
		}
		else if (!this.TargetHtmlElementOnSlide)
		{
			AscCommon.g_inputContext.TargetOffsetY = (this.m_oWordControl.m_oNotesContainer.AbsolutePosition.T * AscCommon.g_dKoef_mm_to_pix) >> 0;
		}

		var targetZoom = isFocusOnSlide ? this.m_oWordControl.m_nZoomValue : 100;

		var oldW = this.TargetHtmlElement.width;
		var oldH = this.TargetHtmlElement.height;

		var newW = 2;
		var newH = (this.m_dTargetSize * targetZoom * g_dKoef_mm_to_pix / 100) >> 0;

		this.TargetHtmlElement.style.transformOrigin = "top left";

		if (oldW !== newW || oldH !== newH)
		{
			var pixNewW = ((newW * AscCommon.AscBrowser.retinaPixelRatio) >> 0) / AscCommon.AscBrowser.retinaPixelRatio;

			this.TargetHtmlElement.style.width = pixNewW + "px";
			this.TargetHtmlElement.style.height = newH + "px";
			this.TargetHtmlElement.oldColor = null;
		}

		var oldColor = this.TargetHtmlElement.oldColor;
		var newColor = this.GetTargetColor(isFocusOnSlide);
		if (!oldColor ||
			oldColor.R !== newColor.R ||
			oldColor.G !== newColor.G ||
			oldColor.B !== newColor.B)
		{
			this.TargetHtmlElement.style.backgroundColor = "rgb(" + newColor.R + "," + newColor.G + "," + newColor.B + ")";
			this.TargetHtmlElement.oldColor = { R : newColor.R, G : newColor.G, B : newColor.B };
		}

		if (null == this.TextMatrix || global_MatrixTransformer.IsIdentity2(this.TextMatrix))
		{
			if (null != this.TextMatrix)
			{
				x += this.TextMatrix.tx;
				y += this.TextMatrix.ty;
			}

			var pos = this.ConvertCoordsToCursor(x, y);

			if (!isFocusOnSlide)
			{
				pos.X = x * g_dKoef_mm_to_pix + AscCommon.AscBrowser.convertToRetinaValue(this.m_oWordControl.m_oNotesApi.OffsetX);
				pos.Y = y * g_dKoef_mm_to_pix - this.m_oWordControl.m_oNotesApi.Scroll;
			}

			this.TargetHtmlElementLeft = pos.X >> 0;
			this.TargetHtmlElementTop = pos.Y >> 0;

			this.TargetHtmlElement.style["transform"] = "";
			this.TargetHtmlElement.style["msTransform"] = "";
			this.TargetHtmlElement.style["mozTransform"] = "";
			this.TargetHtmlElement.style["webkitTransform"] = "";

			if ((!this.m_oWordControl.MobileTouchManager && !AscCommon.AscBrowser.isSafariMacOs) || !AscCommon.AscBrowser.isWebkit)
			{
				this.TargetHtmlElement.style.left = this.TargetHtmlElementLeft + "px";
				this.TargetHtmlElement.style.top = this.TargetHtmlElementTop + "px";
			}
			else
			{
				this.TargetHtmlElement.style.left = "0px";
				this.TargetHtmlElement.style.top = "0px";
				this.TargetHtmlElement.style["webkitTransform"] = "matrix(1, 0, 0, 1, " + oThis.TargetHtmlElementLeft + "," + oThis.TargetHtmlElementTop + ")";
			}
		}
		else
		{
			var x1 = this.TextMatrix.TransformPointX(x, y);
			var y1 = this.TextMatrix.TransformPointY(x, y);

			var pos1 = this.ConvertCoordsToCursor(x1, y1);
			pos1.X -= (newW / 2);

			this.TargetHtmlElementLeft = pos1.X >> 0;
			this.TargetHtmlElementTop = pos1.Y >> 0;

			var transform = "matrix(" + this.TextMatrix.sx + ", " + this.TextMatrix.shy + ", " + this.TextMatrix.shx + ", " +
				this.TextMatrix.sy + ", " + pos1.X + ", " + pos1.Y + ")";

			this.TargetHtmlElement.style.left = "0px";
			this.TargetHtmlElement.style.top = "0px";

			this.TargetHtmlElement.style["transform"] = transform;
			this.TargetHtmlElement.style["msTransform"] = transform;
			this.TargetHtmlElement.style["mozTransform"] = transform;
			this.TargetHtmlElement.style["webkitTransform"] = transform;
		}

		this.MoveTargetInInputContext();
	};

	this.UpdateTarget = function(x, y, pageIndex)
	{
		if (this.m_oWordControl)
			this.m_oWordControl.m_oApi.checkLastWork();

		this.m_oWordControl.m_oLogicDocument.Set_TargetPos(x, y, pageIndex);
		if (pageIndex != this.SlideCurrent && !this.m_oWordControl.DemonstrationManager.Mode)
		{
			// сначала выставим страницу
			this.m_oWordControl.GoToPage(pageIndex);
		}

		if (this.UpdateTargetFromPaint === false)
		{
			this.UpdateTargetCheck = true;
			return;
		}

		var isTargetOnNotes = this.m_oLogicDocument.IsFocusOnNotes();
		var targetZoom = isTargetOnNotes ? 100 : this.m_oWordControl.m_nZoomValue;

		/// detect need scrolling
		var targetSize 	= Number(this.m_dTargetSize * targetZoom * g_dKoef_mm_to_pix / 100);
		var pos        	= null;
		var _x        	= x;
		var _y        	= y;
		if (this.TextMatrix)
		{
			_x = this.TextMatrix.TransformPointX(x, y);
			_y = this.TextMatrix.TransformPointY(x, y);
		}

		this.m_dTargetX = x;
		this.m_dTargetY = y;

		if (!isTargetOnNotes)
		{
			// focus ton slide
			pos = this.ConvertCoordsToCursor(_x, _y);

			var _ww = this.m_oWordControl.m_oEditor.HtmlElement.width;
			var _hh = this.m_oWordControl.m_oEditor.HtmlElement.height;

			_ww /= AscCommon.AscBrowser.retinaPixelRatio;
			_hh /= AscCommon.AscBrowser.retinaPixelRatio;

			var boxX = 0;
			var boxY = 0;
			var boxR = _ww - 2;
			var boxB = _hh - targetSize;

			var nValueScrollHor = 0;
			if (pos.X < boxX)
			{
				nValueScrollHor = (this.m_oWordControl.m_dScrollX + pos.X - boxX) >> 0;
			}
			if (pos.X > boxR)
			{
				nValueScrollHor = (this.m_oWordControl.m_dScrollX + pos.X - boxR) >> 0;
			}

			var nValueScrollVer = 0;
			if (pos.Y < boxY)
			{
				nValueScrollVer = (this.m_oWordControl.m_dScrollY + pos.Y - boxY) >> 0;
			}
			if (pos.Y > boxB)
			{
				nValueScrollVer = (this.m_oWordControl.m_dScrollY + pos.Y - boxB) >> 0;
			}

			/// check scroll
			var isNeedScroll = false;
			if (0 != nValueScrollHor && this.m_oWordControl.m_oScrollHorApi)
			{
				isNeedScroll                                  = true;
				this.m_oWordControl.m_bIsUpdateTargetNoAttack = true;

				if (nValueScrollHor > this.m_oWordControl.m_dScrollX_max)
					nValueScrollHor = this.m_oWordControl.m_dScrollX_max;
				if (0 > nValueScrollHor)
					nValueScrollHor = 0;

				if (this.m_oWordControl.m_oTimerScrollSelect == -1)
					this.m_oWordControl.m_oScrollHorApi.scrollToX(nValueScrollHor, false);
			}
			if (0 != nValueScrollVer)
			{
				isNeedScroll                                  = true;
				this.m_oWordControl.m_bIsUpdateTargetNoAttack = true;

				if (nValueScrollVer > this.m_oWordControl.SlideScrollMAX)
					nValueScrollVer = this.m_oWordControl.SlideScrollMAX - 1;
				if (this.m_oWordControl.SlideScrollMIN > nValueScrollVer)
					nValueScrollVer = this.m_oWordControl.SlideScrollMIN;

				if (this.m_oWordControl.m_oTimerScrollSelect == -1)
					this.m_oWordControl.m_oScrollVerApi.scrollToY(nValueScrollVer, false);
			}

			if (true == isNeedScroll)
			{
				this.m_oWordControl.m_bIsUpdateTargetNoAttack = true;
				this.m_oWordControl.OnScroll();
				return;
			}
		}
		else if (this.m_oWordControl.m_oNotesApi)
		{
			var yPos = _y * g_dKoef_mm_to_pix - this.m_oWordControl.m_oNotesApi.Scroll;
			var _hh = this.m_oWordControl.m_oNotes.HtmlElement.height;
			_hh /= AscCommon.AscBrowser.retinaPixelRatio;

			var boxY = 0;
			var targetSizeAscent = (this.m_dTargetAscent * g_dKoef_mm_to_pix) >> 0;

			var boxB = _hh - (targetSize - targetSizeAscent);
			if (boxB < 0)
				boxB = _hh;

			yPos += targetSizeAscent;

			var nValueScrollVer = 0;
			if (yPos < boxY)
			{
				nValueScrollVer = (this.m_oWordControl.m_oNotesApi.Scroll + yPos - boxY) >> 0;
			}
			if (yPos > boxB)
			{
				nValueScrollVer = (this.m_oWordControl.m_oNotesApi.Scroll + yPos - boxB) >> 0;
			}

			/// check scroll
			if (0 != nValueScrollVer)
			{
				this.m_oWordControl.m_bIsUpdateTargetNoAttack = true;
				this.m_oWordControl.m_oScrollNotes_.scrollToY(nValueScrollVer, false);
				this.m_oWordControl.OnScroll();
				return;
			}
		}

		this.CheckTargetDraw(x, y, !isTargetOnNotes);
	};

	this.SetTargetSize   = function(size, ascent)
	{
		this.m_dTargetSize = size;
		this.m_dTargetAscent = (undefined === ascent) ? 0 : ascent;
	};
	this.DrawTarget      = function()
	{
		let isNeedTarget = (0 != oThis.GetSlidesCount()) && oThis.NeedTarget && !oThis.TransitionSlide.IsPlaying();
		let isShow = false;
		if (isNeedTarget)
		{
			let api = oThis.m_oWordControl.m_oApi;
			if (api.isBlurEditor)
				isShow = true;
			else if (api.isViewMode || api.isRestrictionView())
				isShow = true;
			else if ("block" !== oThis.TargetHtmlElement.style.display)
				isShow = true;
		}

		oThis.TargetHtmlElement.style.display = isShow ? "block" : "none";
	};
	this.isDrawTargetGlass = function()
	{
		let isActive = true;
		let api = oThis.m_oWordControl.m_oApi;

		if (api.isBlurEditor)
			isActive = false;
		else if (api.isViewMode || api.isRestrictionView())
			isActive = false;
		if (-1 === this.m_lTimerTargetId)
			isActive = false;

		return isActive;
	};

	this.TargetShow      = function()
	{
		this.TargetShowNeedFlag = true;
	};
	this.CheckTargetShow = function()
	{
		if (this.TargetShowFlag && this.TargetShowNeedFlag)
		{
			this.TargetHtmlElement.style.display = "block";
			this.TargetShowNeedFlag              = false;
			return;
		}

		if (!this.TargetShowNeedFlag)
			return;

		this.TargetShowNeedFlag = false;

		if (-1 == this.m_lTimerTargetId)
			this.TargetStart();

		if (oThis.NeedTarget)
			this.TargetHtmlElement.style.display = "block";

		this.TargetShowFlag = true;
	};

	this.UpdateTargetTransform = function(matrix)
	{
		this.TextMatrix = matrix;
	};

	this.MultiplyTargetTransform = function (matrix)
	{
		if (!this.TextMatrix)
			this.TextMatrix = matrix;
		else if (matrix)
		{
			this.TextMatrix.Multiply(matrix, AscCommon.MATRIX_ORDER_PREPEND);
		}
	};

	// select
	this.SelectEnabled = function(bIsEnabled)
	{
		this.m_bIsSelection = bIsEnabled;
		if (false === this.m_bIsSelection)
		{
			this.SelectClear();
			this.m_oWordControl.OnUpdateOverlay();
			this.m_oWordControl.m_oOverlayApi.m_oContext.globalAlpha = 1.0;
		}
	};

	this.SelectClear   = function()
	{
		this.m_oWordControl.OnUpdateOverlay();
	};

	this.AddPageSelection = function(pageIndex, x, y, width, height)
	{
		if (null == this.SelectionMatrix)
			this.SelectionMatrix = this.TextMatrix;

		if (pageIndex < 0 || pageIndex != this.SlideCurrent || Math.abs(width) < 0.001 || Math.abs(height) < 0.001)
			return;

		var dPR = AscCommon.AscBrowser.retinaPixelRatio;
		var xDst = this.SlideCurrectRect.left * dPR;
		var yDst = this.SlideCurrectRect.top * dPR;
		var wDst = (this.SlideCurrectRect.right - this.SlideCurrectRect.left) * dPR;
		var hDst = (this.SlideCurrectRect.bottom - this.SlideCurrectRect.top) * dPR;
		var indent = 0.5 * Math.round(dPR);

		var dKoefX = wDst / this.m_oLogicDocument.GetWidthMM();
		var dKoefY = hDst / this.m_oLogicDocument.GetHeightMM();

		var overlay = this.m_oWordControl.m_oOverlayApi;
		if (this.m_oWordControl.IsNotesSupported() && this.m_oWordControl.m_oNotesApi && this.m_oLogicDocument.IsFocusOnNotes())
		{
			overlay = this.m_oWordControl.m_oNotesApi.m_oOverlayApi;
			xDst = this.m_oWordControl.m_oNotesApi.OffsetX;
			yDst = -dPR * this.m_oWordControl.m_oNotesApi.Scroll;
			dKoefX = g_dKoef_mm_to_pix * dPR;
			dKoefY = g_dKoef_mm_to_pix * dPR;
		}

		if (null == this.TextMatrix || global_MatrixTransformer.IsIdentity(this.TextMatrix))
		{
			var _x = ((xDst + dKoefX * x + indent) >> 0) - indent;
			var _y = ((yDst + dKoefY * y + indent) >> 0) - indent;

			var _r = ((xDst + dKoefX * (x + width) + indent) >> 0) - indent;
			var _b = ((yDst + dKoefY * (y + height) + indent) >> 0) - indent;

			if (_x < overlay.min_x)
				overlay.min_x = _x;
			if (_r > overlay.max_x)
				overlay.max_x = _r;

			if (_y < overlay.min_y)
				overlay.min_y = _y;
			if (_b > overlay.max_y)
				overlay.max_y = _b;

			overlay.m_oContext.rect(_x, _y, _r - _x + 1, _b - _y + 1);
		}
		else
		{
			var _x1 = this.TextMatrix.TransformPointX(x, y);
			var _y1 = this.TextMatrix.TransformPointY(x, y);

			var _x2 = this.TextMatrix.TransformPointX(x + width, y);
			var _y2 = this.TextMatrix.TransformPointY(x + width, y);

			var _x3 = this.TextMatrix.TransformPointX(x + width, y + height);
			var _y3 = this.TextMatrix.TransformPointY(x + width, y + height);

			var _x4 = this.TextMatrix.TransformPointX(x, y + height);
			var _y4 = this.TextMatrix.TransformPointY(x, y + height);

			var x1 = xDst + dKoefX * _x1;
			var y1 = yDst + dKoefY * _y1;

			var x2 = xDst + dKoefX * _x2;
			var y2 = yDst + dKoefY * _y2;

			var x3 = xDst + dKoefX * _x3;
			var y3 = yDst + dKoefY * _y3;

			var x4 = xDst + dKoefX * _x4;
			var y4 = yDst + dKoefY * _y4;

			if (global_MatrixTransformer.IsIdentity2(this.TextMatrix))
			{
				x1 = (x1 >> 0) + indent;
				y1 = (y1 >> 0) + indent;

				x2 = (x2 >> 0) + indent;
				y2 = (y2 >> 0) + indent;

				x3 = (x3 >> 0) + indent;
				y3 = (y3 >> 0) + indent;

				x4 = (x4 >> 0) + indent;
				y4 = (y4 >> 0) + indent;
			}

			overlay.CheckPoint(x1, y1);
			overlay.CheckPoint(x2, y2);
			overlay.CheckPoint(x3, y3);
			overlay.CheckPoint(x4, y4);

			var ctx = overlay.m_oContext;
			ctx.moveTo(x1, y1);
			ctx.lineTo(x2, y2);
			ctx.lineTo(x3, y3);
			ctx.lineTo(x4, y4);
			ctx.closePath();
		}
	};

	this.SelectShow       = function()
	{
		this.m_oWordControl.OnUpdateOverlay();
	};

	this.OnSelectEnd = function()
	{
	};

	// search
	this.SearchClear   = function()
	{
		let SlidesCount = this.GetSlidesCount();
		for (var i = 0; i < SlidesCount; i++)
		{
			this.Slide.searchingArray.splice(0, this.Slide.searchingArray.length);
		}
		this.m_oWordControl.m_oOverlayApi.Clear();
		this.m_bIsSearching    = false;
		this.CurrentSearchNavi = null;
	};

	this.AddPageSearch = function(findText, rects)
	{
		var _len = rects.length;
		if (_len == 0)
			return;

		if (this.m_oWordControl.m_oOverlay.HtmlElement.style.display == "none")
		{
			this.m_oWordControl.ShowOverlay();
			this.m_oWordControl.m_oOverlayApi.m_oContext.globalAlpha = 0.2;
		}

		var navigator = {Page : rects[0].PageNum, Place : rects.slice(0, _len)};

		var _find = {text : findText, navigator : navigator};
		this.m_oWordControl.m_oApi.sync_SearchFoundCallback(_find);

		var is_update = false;

		var _pages = this.m_oLogicDocument.Slides;
		for (var i = 0; i < _len; i++)
		{
			var r                                                                     = rects[i];
			_pages[r.PageNum].searchingArray[_pages[r.PageNum].searchingArray.length] = r;

			if (r.PageNum >= this.m_lDrawingFirst && r.PageNum <= this.m_lDrawingEnd)
				is_update = true;
		}

		if (is_update)
			this.m_oWordControl.OnUpdateOverlay();

	};

	this.StartSearch      = function()
	{
		this.SearchClear();
		if (this.m_bIsSelection)
			this.m_oWordControl.OnUpdateOverlay();
		this.m_bIsSearching    = true;
		this.CurrentSearchNavi = null;
	};

	this.EndSearch        = function(bIsChange)
	{
		if (bIsChange)
		{
			this.SearchClear();
			this.m_bIsSearching = false;
			this.m_oWordControl.OnUpdateOverlay();
		}
		else
		{
			this.m_bIsSearching = true;
			this.m_oWordControl.OnUpdateOverlay();
		}
		this.m_oWordControl.m_oApi.sync_SearchEndCallback();
	};

	this.DrawSearch = function(overlay)
	{
		var xDst = this.SlideCurrectRect.left;
		var yDst = this.SlideCurrectRect.top;
		var wDst = this.SlideCurrectRect.right - this.SlideCurrectRect.left;
		var hDst = this.SlideCurrectRect.bottom - this.SlideCurrectRect.top;

		var dKoefX = wDst / this.m_oLogicDocument.GetWidthMM();
		var dKoefY = hDst / this.m_oLogicDocument.GetHeightMM();

		var ctx            = overlay.m_oContext;
		var searchingArray = this.m_oLogicDocument.Slides[this.SlideCurrent].searchingArray;
		for (var i = 0; i < searchingArray.length; i++)
		{
			var place = searchingArray[i];

			if (undefined === place.Ex)
			{
				var _x = ((xDst + dKoefX * place.X) >> 0) - 0.5;
				var _y = ((yDst + dKoefY * place.Y) >> 0) - 0.5;

				var _w = ((dKoefX * place.W) >> 0) + 1;
				var _h = ((dKoefY * place.H) >> 0) + 1;

				if (_x < overlay.min_x)
					overlay.min_x = _x;
				if ((_x + _w) > overlay.max_x)
					overlay.max_x = _x + _w;

				if (_y < overlay.min_y)
					overlay.min_y = _y;
				if ((_y + _h) > overlay.max_y)
					overlay.max_y = _y + _h;

				ctx.rect(_x, _y, _w, _h);
			}
			else
			{
				var _x1 = (xDst + dKoefX * place.X) >> 0;
				var _y1 = (yDst + dKoefY * place.Y) >> 0;

				var x2  = place.X + place.W * place.Ex;
				var y2  = place.Y + place.W * place.Ey;
				var _x2 = (xDst + dKoefX * x2) >> 0;
				var _y2 = (yDst + dKoefY * y2) >> 0;

				var x3  = x2 - place.H * place.Ey;
				var y3  = y2 + place.H * place.Ex;
				var _x3 = (xDst + dKoefX * x3) >> 0;
				var _y3 = (yDst + dKoefY * y3) >> 0;

				var x4  = place.X - place.H * place.Ey;
				var y4  = place.Y + place.H * place.Ex;
				var _x4 = (xDst + dKoefX * x4) >> 0;
				var _y4 = (yDst + dKoefY * y4) >> 0;

				overlay.CheckPoint(_x1, _y1);
				overlay.CheckPoint(_x2, _y2);
				overlay.CheckPoint(_x3, _y3);
				overlay.CheckPoint(_x4, _y4);

				ctx.moveTo(_x1, _y1);
				ctx.lineTo(_x2, _y2);
				ctx.lineTo(_x3, _y3);
				ctx.lineTo(_x4, _y4);
				ctx.lineTo(_x1, _y1);
			}
		}
	};

	this.DrawSearchCur = function(overlay, place)
	{
		var xDst = this.SlideCurrectRect.left;
		var yDst = this.SlideCurrectRect.top;
		var wDst = this.SlideCurrectRect.right - this.SlideCurrectRect.left;
		var hDst = this.SlideCurrectRect.bottom - this.SlideCurrectRect.top;

		var dKoefX = wDst / this.m_oLogicDocument.GetWidthMM();
		var dKoefY = hDst / this.m_oLogicDocument.GetHeightMM();

		var ctx = overlay.m_oContext;

		if (undefined === place.Ex)
		{
			var _x = ((xDst + dKoefX * place.X) >> 0) - 0.5;
			var _y = ((yDst + dKoefY * place.Y) >> 0) - 0.5;

			var _w = ((dKoefX * place.W) >> 0) + 1;
			var _h = ((dKoefY * place.H) >> 0) + 1;

			if (_x < overlay.min_x)
				overlay.min_x = _x;
			if ((_x + _w) > overlay.max_x)
				overlay.max_x = _x + _w;

			if (_y < overlay.min_y)
				overlay.min_y = _y;
			if ((_y + _h) > overlay.max_y)
				overlay.max_y = _y + _h;

			ctx.rect(_x, _y, _w, _h);
		}
		else
		{
			var _x1 = (xDst + dKoefX * place.X) >> 0;
			var _y1 = (yDst + dKoefY * place.Y) >> 0;

			var x2  = place.X + place.W * place.Ex;
			var y2  = place.Y + place.W * place.Ey;
			var _x2 = (xDst + dKoefX * x2) >> 0;
			var _y2 = (yDst + dKoefY * y2) >> 0;

			var x3  = x2 - place.H * place.Ey;
			var y3  = y2 + place.H * place.Ex;
			var _x3 = (xDst + dKoefX * x3) >> 0;
			var _y3 = (yDst + dKoefY * y3) >> 0;

			var x4  = place.X - place.H * place.Ey;
			var y4  = place.Y + place.H * place.Ex;
			var _x4 = (xDst + dKoefX * x4) >> 0;
			var _y4 = (yDst + dKoefY * y4) >> 0;

			overlay.CheckPoint(_x1, _y1);
			overlay.CheckPoint(_x2, _y2);
			overlay.CheckPoint(_x3, _y3);
			overlay.CheckPoint(_x4, _y4);

			ctx.moveTo(_x1, _y1);
			ctx.lineTo(_x2, _y2);
			ctx.lineTo(_x3, _y3);
			ctx.lineTo(_x4, _y4);
			ctx.lineTo(_x1, _y1);
		}
	};

	// recalculate
	this.OnStartRecalculate = function(pageCount)
	{
		if (this.m_oWordControl)
			this.m_oWordControl.m_oApi.checkLastWork();
	};


	this.OnRecalculateSlide = function(index)
	{

		let thpages = this.m_oWordControl.Thumbnails.m_arrPages;
		if(index < 0 || index >= thpages.length) return;

		if (this.m_oWordControl && this.m_oWordControl.MobileTouchManager)
		{
			this.m_oWordControl.MobileTouchManager.ClearContextMenu();
		}

		if (this.TransitionSlide && this.TransitionSlide.IsPlaying())
			this.TransitionSlide.End(true);

		editor.sendEvent("asc_onDocumentChanged");

		if (true === this.m_bIsSearching)
		{
			this.SearchClear();
			this.m_oWordControl.OnUpdateOverlay();
			this.SendChangeDocumentToApi(true);
		}

		if (thpages.length > index)
		{
			thpages[index].IsRecalc = true;
		}

		if (index === this.SlideCurrent)
		{
			this.m_oWordControl.Thumbnails.LockMainObjType = true;

			// так как серега посылает по сто раз - делаем такую заглушку ---------------------

			this.m_oWordControl.SlideDrawer.CheckSlide(this.SlideCurrent);
			this.m_oWordControl.CalculateDocumentSize(false);
			// --------------------------------------------------------------------------------

			this.m_oWordControl.OnScroll();
			this.m_oWordControl.Thumbnails.LockMainObjType = false;
		}
	};

	this.OnEndRecalculate = function()
	{
		if (this.m_oWordControl)
			this.m_oWordControl.m_oApi.checkLastWork();

		this.m_oWordControl.Thumbnails.LockMainObjType = true;
		this.SlidesCount                               = this.m_oLogicDocument.GetSlidesCount();
		this.m_oWordControl.CalculateDocumentSize();
		this.m_oWordControl.m_oApi.sync_countPagesCallback(this.SlidesCount);
		this.m_oWordControl.Thumbnails.LockMainObjType = false;
	};

	// draw
	this.LockSlide = function(slideNum)
	{
		var _th_manager = this.m_oWordControl.Thumbnails;

		if (slideNum >= 0 && slideNum < _th_manager.m_arrPages.length)
			_th_manager.m_arrPages[slideNum].IsLocked = true;

		_th_manager.OnUpdateOverlay();
	};

	this.UnLockSlide = function(slideNum)
	{
		var _th_manager = this.m_oWordControl.Thumbnails;

		if (slideNum >= 0 && slideNum < _th_manager.m_arrPages.length)
			_th_manager.m_arrPages[slideNum].IsLocked = false;

		_th_manager.OnUpdateOverlay();
	};

	this.UpdateThumbnailsAttack = function()
	{
		this.m_oWordControl.Thumbnails.RecalculateAll();
	};

	this.ChangePageAttack = function(pageIndex)
	{
		if (pageIndex != this.SlideCurrent)
			return;

		this.StopRenderingPage(pageIndex);
		this.m_oWordControl.OnScroll();
	};

	this.StopRenderingPage = function(pageIndex)
	{
		return;
	};

	this.ClearCachePages = function()
	{
		if (this.m_oWordControl.m_oApi.bInit_word_control && 0 <= this.SlideCurrent)
			this.m_oWordControl.SlideDrawer.CheckSlide(this.SlideCurrent);
	};

	this.FirePaint = function()
	{
		//this.m_oWordControl.OnScroll();

		this.m_oWordControl.Thumbnails.LockMainObjType = true;

		// так как серега посылает по сто раз - делаем такую заглушку ---------------------
		this.m_oWordControl.SlideDrawer.CheckSlide(this.SlideCurrent);
		this.m_oWordControl.CalculateDocumentSize(false);
		// --------------------------------------------------------------------------------
		this.m_oWordControl.OnScroll();
		this.m_oWordControl.Thumbnails.LockMainObjType = false;
	};

	// interface ui
	this.InitGuiCanvasTextProps = function(div_id)
	{
		var _div_elem = document.getElementById(div_id);
		if (null != this.GuiCanvasTextProps)
		{
			var elem = _div_elem.getElementsByTagName('canvas');
			if (elem.length == 0)
			{
				_div_elem.appendChild(this.GuiCanvasTextProps);
			}
			else
			{
				var _width = parseInt(_div_elem.offsetWidth);
				var _height = parseInt(_div_elem.offsetHeight);
				if (0 == _width)
					_width = 300;
				if (0 == _height)
					_height = 80;

				this.GuiCanvasTextProps.style.width = _width + "px";
				this.GuiCanvasTextProps.style.height = _height + "px";
			}

			var old_width = this.GuiCanvasTextProps.width;
			var old_height = this.GuiCanvasTextProps.height;
			AscCommon.calculateCanvasSize(this.GuiCanvasTextProps);

			if (old_width !== this.GuiCanvasTextProps.width || old_height !== this.GuiCanvasTextProps.height)
				this.GuiLastTextProps = null;
		}
		else
		{
			this.GuiCanvasTextProps = document.createElement('canvas');
			this.GuiCanvasTextProps.style.position = "absolute";
			this.GuiCanvasTextProps.style.left = "0px";
			this.GuiCanvasTextProps.style.top = "0px";

			var _width = parseInt(_div_elem.offsetWidth);
			var _height = parseInt(_div_elem.offsetHeight);
			if (0 == _width)
				_width = 300;
			if (0 == _height)
				_height = 80;

			this.GuiCanvasTextProps.style.width = _width + "px";
			this.GuiCanvasTextProps.style.height = _height + "px";

			AscCommon.calculateCanvasSize(this.GuiCanvasTextProps);

			_div_elem.appendChild(this.GuiCanvasTextProps);
		}
	};

	this.DrawGuiCanvasTextProps = function(props)
	{
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

	this.CheckGuiControlColors = function(bIsAttack)
	{
		let oPresentation = this.m_oWordControl.m_oLogicDocument;
		let oTheme = oPresentation.Get_Theme();
		let oClrMap = oPresentation.Get_ColorMap();
		if(!oTheme || !oClrMap) return;

		let arr_colors = new Array(10);

		let rgba               = {R : 0, G : 0, B : 0, A : 255};
		// bg1,tx1,bg2,tx2,accent1 - accent6
		let array_colors_types = [6, 15, 7, 16, 0, 1, 2, 3, 4, 5];
		let _count             = array_colors_types.length;
		let color   = new AscFormat.CUniColor();
		color.color = new AscFormat.CSchemeColor();
		for (let i = 0; i < _count; ++i)
		{
			color.color.id = array_colors_types[i];
			color.Calculate(oTheme, null, null, null, rgba, oClrMap);
			let _rgba     = color.RGBA;
			arr_colors[i] = new Asc.asc_CColor(_rgba.R, _rgba.G, _rgba.B);
			arr_colors[i].setColorSchemeId(color.color.id);
		}

		// теперь проверим
		let bIsSend = false;
		if (this.GuiControlColorsMap != null)
		{
			for (let i = 0; i < _count; ++i)
			{
				let _color1 = this.GuiControlColorsMap[i];
				let _color2 = arr_colors[i];

				if ((_color1.r != _color2.r) || (_color1.g != _color2.g) || (_color1.b != _color2.b))
				{
					bIsSend = true;
					break;
				}
			}
		}
		else
		{
			this.GuiControlColorsMap = new Array(_count);
			bIsSend                  = true;
		}

		if (bIsSend || (bIsAttack === true))
		{
			for (let i = 0; i < _count; ++i)
			{
				this.GuiControlColorsMap[i] = arr_colors[i];
			}

			this.SendControlColors(bIsAttack);
		}
	};

	this.SendControlColors = function(bIsAttack)
	{
		let standart_colors = null;
		if (!this.IsSendStandartColors || (bIsAttack === true))
		{
			let standartColors = AscCommon.g_oStandartColors;
			let _c_s           = standartColors.length;
			standart_colors    = new Array(_c_s);

			for (var i = 0; i < _c_s; ++i)
			{
				standart_colors[i] = new Asc.asc_CColor(standartColors[i].R, standartColors[i].G, standartColors[i].B);
			}

			this.IsSendStandartColors = true;
		}

		let _count = this.GuiControlColorsMap.length;

		let _ret_array = new Array(_count * 6);
		let _cur_index = 0;

		let array_colors_types = [6, 15, 7, 16, 0, 1, 2, 3, 4, 5];
		for (let i = 0; i < _count; ++i)
		{
			let _color_src = this.GuiControlColorsMap[i];

			_ret_array[_cur_index] = new Asc.asc_CColor(_color_src.r, _color_src.g, _color_src.b);
			_ret_array[_cur_index].setColorSchemeId(array_colors_types[i]);
			_cur_index++;

			// теперь с модификаторами
			let _count_mods = 5;
			for (let j = 0; j < _count_mods; ++j)
			{
				let dst_mods  = new AscFormat.CColorModifiers();
				dst_mods.Mods = AscCommon.GetDefaultMods(_color_src.r, _color_src.g, _color_src.b, j + 1, 0);

				let _rgba = {R : _color_src.r, G : _color_src.g, B : _color_src.b, A : 255};
				dst_mods.Apply(_rgba);

				let oColor = new Asc.asc_CColor(_rgba.R, _rgba.G, _rgba.B);
				oColor.put_effectValue(dst_mods.getEffectValue());
				oColor.setColorSchemeId(array_colors_types[i]);
				_ret_array[_cur_index] = oColor;
				_cur_index++;
			}
		}

		this.m_oWordControl.m_oApi.sync_SendThemeColors(_ret_array, standart_colors);
	};

	this.DrawGuiImage = function (oTexture, oCtx, sUrl, sLastDrawingUrl)
	{

		let _width          = oTexture.width;
		let _height         = oTexture.height;

		oCtx.clearRect(0, 0, _width, _height);

		if (!sLastDrawingUrl)
			return;

		let _img = Asc.editor.ImageLoader.map_image_index[AscCommon.getFullImageSrc2(sUrl)];
		if (_img != undefined && _img.Image != null && _img.Status != ImageLoadStatus.Loading)
		{
			let _x = 0;
			let _y = 0;
			let _w = Math.max(_img.Image.width, 1);
			let _h = Math.max(_img.Image.height, 1);

			let dAspect1 = _width / _height;
			let dAspect2 = _w / _h;

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

			oCtx.drawImage(_img.Image, _x, _y, _w, _h);
		}
		else
		{
			if(_img.Status === ImageLoadStatus.Loading)
			{

			}
			else
			{
				oCtx.lineWidth = 1;

				oCtx.beginPath();
				oCtx.moveTo(0, 0);
				oCtx.lineTo(_width, _height);
				oCtx.moveTo(_width, 0);
				oCtx.lineTo(0, _height);
				oCtx.strokeStyle = "#FF0000";
				oCtx.stroke();

				oCtx.beginPath();
				oCtx.moveTo(0, 0);
				oCtx.lineTo(_width, 0);
				oCtx.lineTo(_width, _height);
				oCtx.lineTo(0, _height);
				oCtx.closePath();

				oCtx.strokeStyle = "#000000";
				oCtx.stroke();
				oCtx.beginPath();
			}
		}
	};

	this.DrawImageTextureFillShape = function(url)
	{
		if (this.GuiCanvasFillTexture == null)
		{
			this.InitGuiCanvasShape(this.GuiCanvasFillTextureParentId);
		}

		if (this.GuiCanvasFillTexture == null || this.GuiCanvasFillTextureCtx == null || url == this.LastDrawingUrl)
			return;

		this.LastDrawingUrl = url;
		this.DrawGuiImage(this.GuiCanvasFillTexture, this.GuiCanvasFillTextureCtx, url, this.LastDrawingUrl);
		if(url)
		{
			let _img = Asc.editor.ImageLoader.map_image_index[AscCommon.getFullImageSrc2(url)];
			if (!(_img && _img.Image && _img.Status !== ImageLoadStatus.Loading))
				this.LastDrawingUrl = "";
		}
	};

	this.DrawImageTextureFillSlide = function(url)
	{
		if (this.GuiCanvasFillTextureSlide == null)
		{
			this.InitGuiCanvasSlide(this.GuiCanvasFillTextureParentIdSlide);
		}

		if (this.GuiCanvasFillTextureSlide == null || this.GuiCanvasFillTextureCtxSlide == null || url == this.LastDrawingUrlSlide)
			return;

		this.LastDrawingUrlSlide = url;
		this.DrawGuiImage(this.GuiCanvasFillTextureSlide, this.GuiCanvasFillTextureCtxSlide, url, this.LastDrawingUrlSlide);
		if(url)
		{
			let _img = Asc.editor.ImageLoader.map_image_index[AscCommon.getFullImageSrc2(url)];
			if (!(_img && _img.Image && _img.Status !== ImageLoadStatus.Loading))
				this.LastDrawingUrlSlide = "";
		}
	};

	this.DrawImageTextureFillTextArt = function(url)
	{
		if (this.GuiCanvasFillTextureTextArt == null)
		{
			this.InitGuiCanvasTextArt(this.GuiCanvasFillTextureParentIdTextArt);
		}

		if (this.GuiCanvasFillTextureTextArt == null || this.GuiCanvasFillTextureCtxTextArt == null || url == this.LastDrawingUrlTextArt)
			return;

		this.LastDrawingUrlTextArt = url;

		this.DrawGuiImage(this.GuiCanvasFillTextureTextArt, this.GuiCanvasFillTextureCtxTextArt, url, this.LastDrawingUrlTextArt);
		if(url)
		{
			let _img = Asc.editor.ImageLoader.map_image_index[AscCommon.getFullImageSrc2(url)];
			if (!(_img && _img.Image && _img.Status !== ImageLoadStatus.Loading))
				this.LastDrawingUrlTextArt = "";
		}
	};

	this.InitGuiCanvasShape = function(div_id)
	{
		if (null != this.GuiCanvasFillTexture)
		{
			var _div_elem = document.getElementById(this.GuiCanvasFillTextureParentId);
			if (!_div_elem)
				_div_elem.removeChild(this.GuiCanvasFillTexture);

			this.GuiCanvasFillTexture    = null;
			this.GuiCanvasFillTextureCtx = null;
		}

		this.GuiCanvasFillTextureParentId = div_id;
		var _div_elem                     = document.getElementById(this.GuiCanvasFillTextureParentId);
		if (!_div_elem)
			return;

		this.GuiCanvasFillTexture        = document.createElement('canvas');
		this.GuiCanvasFillTexture.width  = parseInt(_div_elem.style.width);
		this.GuiCanvasFillTexture.height = parseInt(_div_elem.style.height);

		this.LastDrawingUrl          = "";
		this.GuiCanvasFillTextureCtx = this.GuiCanvasFillTexture.getContext('2d');

		_div_elem.appendChild(this.GuiCanvasFillTexture);
	};

	this.InitGuiCanvasSlide = function(div_id)
	{
		if (null != this.GuiCanvasFillTextureSlide)
		{
			var _div_elem = document.getElementById(this.GuiCanvasFillTextureParentIdSlide);
			if (!_div_elem)
				_div_elem.removeChild(this.GuiCanvasFillTextureSlide);

			this.GuiCanvasFillTextureSlide    = null;
			this.GuiCanvasFillTextureCtxSlide = null;
		}

		this.GuiCanvasFillTextureParentIdSlide = div_id;
		var _div_elem                          = document.getElementById(this.GuiCanvasFillTextureParentIdSlide);
		if (!_div_elem)
			return;

		this.GuiCanvasFillTextureSlide        = document.createElement('canvas');
		this.GuiCanvasFillTextureSlide.width  = parseInt(_div_elem.style.width);
		this.GuiCanvasFillTextureSlide.height = parseInt(_div_elem.style.height);

		this.LastDrawingUrlSlide          = "";
		this.GuiCanvasFillTextureCtxSlide = this.GuiCanvasFillTextureSlide.getContext('2d');

		_div_elem.appendChild(this.GuiCanvasFillTextureSlide);
	};

	this.InitGuiCanvasTextArt = function(div_id)
	{
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

	this.CheckTableStyles = function(tableLook)
	{
		// сначала проверим, подписан ли кто на этот евент
		// а то во вьюере не стоит ничего посылать
		if (!this.m_oWordControl.m_oApi.asc_checkNeedCallback("asc_onInitTableTemplates"))
			return;
		var isChanged = this.m_oWordControl.m_oLogicDocument.CheckNeedUpdateTableStyles(tableLook);
		if(!isChanged)
		{
			return;
		}
		this.m_oWordControl.m_oApi.sync_InitEditorTableStyles();
	};

	this.CheckTableStylesDefault = function ()
	{
		let tableLook = new AscCommon.CTableLook(true, true, false, false, true, false);
		return this.CheckTableStyles(tableLook);
	};

	this.GetTableStylesPreviews = function(bUseDefault)
	{
		return (new AscCommon.CTableStylesPreviewGenerator(this.m_oWordControl.m_oLogicDocument)).GetAllPreviews(bUseDefault);
	};

	this.GetTableLook = function(isDefault)
	{
		let oTableLook;

		if (isDefault)
		{
			oTableLook = new AscCommon.CTableLook();
			oTableLook.SetDefault();
		}
		else
		{
			oTableLook = this.TableStylesLastLook;
		}

		return oTableLook;
	};

	// ruler
	this.Set_RulerState_Table = function(markup, transform)
	{
		this.TabButtonEnable(true);

		var hor_ruler = this.m_oWordControl.m_oHorRuler;
		var ver_ruler = this.m_oWordControl.m_oVerRuler;

		hor_ruler.CurrentObjectType = RULER_OBJECT_TYPE_TABLE;
		hor_ruler.m_oTableMarkup    = markup.CreateDublicate();
		hor_ruler.m_oColumnMarkup	= null;
		hor_ruler.CalculateMargins();

		ver_ruler.CurrentObjectType = RULER_OBJECT_TYPE_TABLE;
		ver_ruler.m_oTableMarkup    = markup.CreateDublicate();

		if (transform)
		{
			hor_ruler.m_oTableMarkup.TransformX = transform.tx;
			hor_ruler.m_oTableMarkup.TransformY = transform.ty;

			ver_ruler.m_oTableMarkup.TransformX = transform.tx;
			ver_ruler.m_oTableMarkup.TransformY = transform.ty;

			hor_ruler.m_oTableMarkup.CorrectFrom();
			ver_ruler.m_oTableMarkup.CorrectFrom();
		}

		hor_ruler.CalculateMargins();
		this.TableOutlineDr.CurrentPageIndex = this.m_lCurrentPage;

		if (0 <= this.SlideCurrent && this.SlideCurrent < this.GetSlidesCount())
		{
			this.m_oWordControl.CreateBackgroundHorRuler();
			this.m_oWordControl.CreateBackgroundVerRuler();
		}

		if (!this.m_oWordControl.IsEnabledRulerMarkers)
		{
			this.m_oWordControl.EnableRulerMarkers();
		}
		else
		{
			this.m_oWordControl.UpdateHorRuler();
		}
		this.m_oWordControl.UpdateVerRuler();

		if (this.m_oWordControl.MobileTouchManager)
		{
			this.m_oWordControl.MobileTouchManager.TableStartTrack_Check = true;
			markup.Table.StartTrackTable();
			this.m_oWordControl.MobileTouchManager.TableStartTrack_Check = false;
		}
	};

	this.Set_RulerState_Paragraph = function(obj, margins)
	{
		this.TabButtonEnable((margins !== undefined) ? true : false);

		var hor_ruler = this.m_oWordControl.m_oHorRuler;
		var ver_ruler = this.m_oWordControl.m_oVerRuler;

		//if (hor_ruler.CurrentObjectType == RULER_OBJECT_TYPE_PARAGRAPH && ver_ruler.CurrentObjectType == RULER_OBJECT_TYPE_PARAGRAPH)
		//    return;

		hor_ruler.CurrentObjectType = RULER_OBJECT_TYPE_PARAGRAPH;
		hor_ruler.m_oTableMarkup    = null;
		hor_ruler.m_oColumnMarkup	= null;

		ver_ruler.CurrentObjectType = RULER_OBJECT_TYPE_PARAGRAPH;
		ver_ruler.m_oTableMarkup    = null;

		// вообще надо посмотреть... может и был параграф до этого.
		// тогда вэкграунд перерисовывать не нужно. Только надо знать, на той же странице это было или нет
		if (-1 != this.SlideCurrent)
		{
			this.m_oWordControl.CreateBackgroundHorRuler(margins);
			this.m_oWordControl.CreateBackgroundVerRuler(margins);
		}

		if (!this.m_oWordControl.IsEnabledRulerMarkers && margins !== undefined)
			this.m_oWordControl.EnableRulerMarkers();
		else if (this.m_oWordControl.IsEnabledRulerMarkers && margins === undefined)
			this.m_oWordControl.DisableRulerMarkers();
		else
		{
			this.m_oWordControl.UpdateHorRuler();
			this.m_oWordControl.UpdateVerRuler();
		}
	};

	this.Set_RulerState_HdrFtr = function(bHeader, Y0, Y1)
	{
		this.TabButtonEnable(true);

		var hor_ruler = this.m_oWordControl.m_oHorRuler;
		var ver_ruler = this.m_oWordControl.m_oVerRuler;

		hor_ruler.CurrentObjectType = RULER_OBJECT_TYPE_PARAGRAPH;
		hor_ruler.m_oTableMarkup    = null;
		hor_ruler.m_oColumnMarkup	= null;

		ver_ruler.CurrentObjectType = (true === bHeader) ? RULER_OBJECT_TYPE_HEADER : RULER_OBJECT_TYPE_FOOTER;
		ver_ruler.header_top        = Y0;
		ver_ruler.header_bottom     = Y1;
		ver_ruler.m_oTableMarkup    = null;

		// вообще надо посмотреть... может и бал параграф до этого.
		// тогда вэкграунд перерисовывать не нужно. Только надо знать, на той же странице это было или нет
		if (-1 != this.m_lCurrentPage)
		{
			this.m_oWordControl.CreateBackgroundHorRuler();
			this.m_oWordControl.CreateBackgroundVerRuler();
		}

		this.m_oWordControl.UpdateHorRuler();
		this.m_oWordControl.UpdateVerRuler();
	};

	this.Set_RulerState_Columns = function (markup)
	{
		this.TabButtonEnable(true);

		var hor_ruler = this.m_oWordControl.m_oHorRuler;
		var ver_ruler = this.m_oWordControl.m_oVerRuler;

		hor_ruler.CurrentObjectType = RULER_OBJECT_TYPE_COLUMNS;
		hor_ruler.m_oTableMarkup = null;
		hor_ruler.m_oColumnMarkup = markup.CreateDuplicate();

		ver_ruler.CurrentObjectType = RULER_OBJECT_TYPE_PARAGRAPH;
		ver_ruler.m_oTableMarkup = null;

		this.TableOutlineDr.TableMatrix = null;
		this.TableOutlineDr.CurrentPageIndex = this.m_lCurrentPage;

		hor_ruler.CalculateMargins();

		if (0 <= this.m_lCurrentPage && this.m_lCurrentPage < this.m_lPagesCount)
		{
			this.m_oWordControl.CreateBackgroundHorRuler();
			this.m_oWordControl.CreateBackgroundVerRuler();
		}

		this.m_oWordControl.UpdateHorRuler();
		this.m_oWordControl.UpdateVerRuler();
	};

	this.Update_ParaTab = function(Default_Tab, ParaTabs)
	{
		var hor_ruler = this.m_oWordControl.m_oHorRuler;

		var __tabs = ParaTabs.Tabs;
		if (undefined === __tabs)
			__tabs = ParaTabs;

		var _len = __tabs.length;
		if ((Default_Tab == hor_ruler.m_dDefaultTab) && (hor_ruler.m_arrTabs.length == _len) && (_len == 0))
		{
			// потом можно и проверить сами табы
			return;
		}

		hor_ruler.m_dDefaultTab = Default_Tab;
		hor_ruler.m_arrTabs     = [];
		var _ar                 = hor_ruler.m_arrTabs;

		for (var i = 0; i < _len; i++)
		{
			if (__tabs[i].Value == tab_Left || __tabs[i].Value == tab_Center || __tabs[i].Value == tab_Right)
				_ar[i] = new CTab(__tabs[i].Pos, __tabs[i].Value);
			else
			{
				// не должно такого быть. но приходит
				_ar[i] = new CTab(__tabs[i].Pos, tab_Left);
			}
		}

		hor_ruler.CorrectTabs();
		this.m_oWordControl.UpdateHorRuler();
	};

	this.UpdateTableRuler = function(isCols, index, position)
	{
		var dKoef_mm_to_pix = g_dKoef_mm_to_pix * this.m_oWordControl.m_nZoomValue / 100;
		if (false === isCols)
		{
			var markup = this.m_oWordControl.m_oVerRuler.m_oTableMarkup;
			if (markup == null)
				return;

			position += markup.TransformY;
			if (0 == index)
			{
				var delta        = position - markup.Rows[0].Y;
				markup.Rows[0].Y = position;
				markup.Rows[0].H -= delta;
			}
			else
			{
				var delta = (markup.Rows[index - 1].Y + markup.Rows[index - 1].H) - position;

				markup.Rows[index - 1].H -= delta;

				if (index != markup.Rows.length)
				{
					markup.Rows[index].Y -= delta;
					markup.Rows[index].H += delta;
				}
			}

			if ("none" == this.m_oWordControl.m_oOverlay.HtmlElement.style.display)
				this.m_oWordControl.ShowOverlay();

			this.m_oWordControl.UpdateVerRulerBack();
			this.m_oWordControl.m_oOverlayApi.HorLine(this.SlideCurrectRect.top + position * dKoef_mm_to_pix);
		}
		else
		{
			var markup = this.m_oWordControl.m_oHorRuler.m_oTableMarkup;
			if (markup == null)
				return;

			position += markup.TransformX;

			if (0 == index)
			{
				markup.X = position;
			}
			else
			{
				var _start = markup.X;
				for (var i = 0; i < (index - 1); i++)
				{
					_start += markup.Cols[i];
				}

				var _old               = markup.Cols[index - 1];
				markup.Cols[index - 1] = position - _start;

				if (index != markup.Cols.length)
				{
					markup.Cols[index] += (_old - markup.Cols[index - 1]);
				}
			}

			if ("none" == this.m_oWordControl.m_oOverlay.HtmlElement.style.display)
				this.m_oWordControl.ShowOverlay();

			this.m_oWordControl.UpdateHorRulerBack();
			this.m_oWordControl.m_oOverlayApi.VertLine(this.SlideCurrectRect.left + position * dKoef_mm_to_pix);
		}
	};

	this.TabButtonEnable = function(isEnabled)
	{
		if (this.isTabButtonShow == isEnabled)
			return;

		this.isTabButtonShow = isEnabled;
		if (this.m_oWordControl && this.m_oWordControl.m_oLeftRuler_buttonsTabs)
		{
			this.m_oWordControl.m_oLeftRuler_buttonsTabs.HtmlElement.style.display = this.isTabButtonShow ? "block" : "none";
		}
	};

	this.CorrectRulerPosition = function(pos)
	{
		if (global_keyboardEvent.AltKey)
			return pos;

		return ((pos / 2.5 + 0.5) >> 0) * 2.5;
	};

	// coords
	this.ConvertCoordsToAnotherPage = function(x, y)
	{
		return {X : x, Y : y};
	};

	this.ConvertCoordsFromCursor2 = function(x, y)
	{
		var wc = this.m_oWordControl;
		var _x = x - wc.X - (wc.m_oMainParent.AbsolutePosition.L + wc.m_oMainView.AbsolutePosition.L) * g_dKoef_mm_to_pix;
		var _y = y - wc.Y - (wc.m_oMainParent.AbsolutePosition.T + wc.m_oMainView.AbsolutePosition.T) * g_dKoef_mm_to_pix;

		var dKoef = (100 * g_dKoef_pix_to_mm / wc.m_nZoomValue);

		var Pos = {X : 0, Y : 0, Page : this.SlideCurrent};

		if (this.SlideCurrent != -1)
		{
			var rect = this.SlideCurrectRect;

			var x_mm = (_x - rect.left) * dKoef;
			var y_mm = (_y - rect.top) * dKoef;

			Pos.X = x_mm;
			Pos.Y = y_mm;
		}

		return Pos;
	};

	this.ConvertCoordsToCursorWR = function(x, y, pageIndex, transform, isMainAttack)
	{
		var wc = this.m_oWordControl;

		var _x = x;
		var _y = y;
		if (transform)
		{
			_x = transform.TransformPointX(x, y);
			_y = transform.TransformPointY(x, y);
		}

		if (isMainAttack || !wc.m_oLogicDocument.IsFocusOnNotes())
		{
			if (wc.DemonstrationManager.Mode)
			{
				return wc.DemonstrationManager.convertCoordsToCursorWR(_x, _y);
			}

			var dKoef = (this.m_oWordControl.m_nZoomValue * g_dKoef_mm_to_pix / 100);
			var x_pix = (this.SlideCurrectRect.left + _x * dKoef + (wc.m_oMainParent.AbsolutePosition.L + wc.m_oMainView.AbsolutePosition.L) * g_dKoef_mm_to_pix) >> 0;
			var y_pix = (this.SlideCurrectRect.top  + _y * dKoef + (wc.m_oMainParent.AbsolutePosition.T + wc.m_oMainView.AbsolutePosition.T) * g_dKoef_mm_to_pix) >> 0;

			return {X: x_pix, Y: y_pix, Error: false};
		}
		else
		{
			var x_pix = (_x * g_dKoef_mm_to_pix + 10 + (wc.m_oMainParent.AbsolutePosition.L + wc.m_oBottomPanesContainer.AbsolutePosition.L + wc.m_oNotesContainer.AbsolutePosition.L) * g_dKoef_mm_to_pix) >> 0;
			var y_pix = (_y * g_dKoef_mm_to_pix      + (wc.m_oMainParent.AbsolutePosition.T + wc.m_oBottomPanesContainer.AbsolutePosition.T + wc.m_oNotesContainer.AbsolutePosition.T) * g_dKoef_mm_to_pix) >> 0;

			return {X: x_pix, Y: y_pix, Error: false};
		}
	};

	this.ConvertAnimPaneCoordsToCursor = function(x, y)
	{
		const wc = this.m_oWordControl;
		const x_pix = (x * g_dKoef_mm_to_pix + (wc.m_oMainParent.AbsolutePosition.L + wc.m_oBottomPanesContainer.AbsolutePosition.L + wc.m_oAnimationPaneContainer.AbsolutePosition.L) * g_dKoef_mm_to_pix) >> 0;
		const y_pix = (y * g_dKoef_mm_to_pix + (wc.m_oMainParent.AbsolutePosition.T + wc.m_oBottomPanesContainer.AbsolutePosition.T + wc.m_oAnimationPaneContainer.AbsolutePosition.T) * g_dKoef_mm_to_pix) >> 0;
		return {X: x_pix, Y: y_pix, Error: false};
	};

	this.ConvertCoordsToCursor3 = function (x, y, isGlobal)
	{
		var wc = this.m_oWordControl;
		var dKoef = (wc.m_nZoomValue * g_dKoef_mm_to_pix / 100);

		var _x = 0;
		var _y = 0;
		if (isGlobal)
		{
			_x = wc.X;
			_y = wc.Y;

			_x += (wc.m_oMainParent.AbsolutePosition.L + wc.m_oMainView.AbsolutePosition.L) * g_dKoef_mm_to_pix;
			_y += (wc.m_oMainParent.AbsolutePosition.T + wc.m_oMainView.AbsolutePosition.T) * g_dKoef_mm_to_pix;
		}

		var x_pix = (this.SlideCurrectRect.left + x * dKoef + _x) >> 0;
		var y_pix = (this.SlideCurrectRect.top  + y * dKoef + _y) >> 0;

		return {X: x_pix, Y: y_pix, Error: false};
	};

	this.ConvertCoordsToCursorWR_Comment = function(x, y)
	{
		var wc = this.m_oWordControl;
		var dKoef = (wc.m_nZoomValue * g_dKoef_mm_to_pix / 100);

		var x_pix = (this.SlideCurrectRect.left + x * dKoef + wc.m_oMainView.AbsolutePosition.L * g_dKoef_mm_to_pix) >> 0;
		var y_pix = (this.SlideCurrectRect.top  + y * dKoef + wc.m_oMainView.AbsolutePosition.T * g_dKoef_mm_to_pix) >> 0;

		x_pix += COMMENT_WIDTH;
		y_pix += ((COMMENT_HEIGHT / 3) >> 0);

		return {X : x_pix, Y : y_pix, Error : false};
	};

	this.ConvertCoordsToCursor = function(x, y)
	{
		var dKoef = (this.m_oWordControl.m_nZoomValue * g_dKoef_mm_to_pix / 100);

		var x_pix = (this.SlideCurrectRect.left + x * dKoef) >> 0;
		var y_pix = (this.SlideCurrectRect.top  + y * dKoef) >> 0;

		return {X : x_pix, Y : y_pix, Error : false};
	};

	this.GetDotsPerMM     = function(value)
	{
		if (!this.isDrawingNotes)
			return value * this.m_oWordControl.m_nZoomValue * g_dKoef_mm_to_pix / 100;
		return value * g_dKoef_mm_to_pix;
	};

	this.GetMMPerDot        = function(value)
	{
		return value / this.GetDotsPerMM(1);
	};

	this.GetVisibleMMHeight = function()
	{
		var pixHeigth = this.m_oWordControl.m_oEditor.HtmlElement.height;
		pixHeigth /= AscCommon.AscBrowser.retinaPixelRatio;
		var pixBetweenPages = 20 * (this.m_lDrawingEnd - this.m_lDrawingFirst);

		return (pixHeigth - pixBetweenPages) * g_dKoef_pix_to_mm * 100 / this.m_oWordControl.m_nZoomValue;
	};

	// tracks
	this.StartTrackTable = function (obj, transform)
	{
		if (this.m_oWordControl.MobileTouchManager)
		{
			if (!this.m_oWordControl.MobileTouchManager.TableStartTrack_Check)
				return;
		}

		this.TableOutlineDr.TableOutline = obj;
		this.TableOutlineDr.Counter = 0;
		this.TableOutlineDr.bIsNoTable = false;
		this.TableOutlineDr.CheckStartTrack(this.m_oWordControl, transform);

		if (this.m_oWordControl.MobileTouchManager)
			this.m_oWordControl.OnUpdateOverlay();
	};

	this.EndTrackTable = function (pointer, bIsAttack)
	{
		if (this.TableOutlineDr.TableOutline != null)
		{
			if (pointer == this.TableOutlineDr.TableOutline.Table || bIsAttack)
			{
				this.TableOutlineDr.TableOutline = null;
				this.TableOutlineDr.Counter = 0;
			}
		}
	};

	this.StartTrackAutoshape = function()
	{
		this.m_oWordControl.ShowOverlay();
	};

	this.EndTrackAutoShape   = function()
	{
		this.m_oWordControl.OnUpdateOverlay();
	};

	this.DrawMathTrack = function (overlay)
	{
		if (!this.MathTrack.IsActive())
			return;
		if(!this.TextMatrix)
			return;

		overlay.Show();
		let nIndex, nCount;
		let oPath;
		let dKoefX, dKoefY;
		let PathLng = this.MathTrack.GetPolygonsCount();

		let xDst, yDst, wDst, hDst;
		let bIsMain = !this.m_oLogicDocument.IsFocusOnNotes();
		if(bIsMain)
		{
			xDst = this.SlideCurrectRect.left;
			yDst = this.SlideCurrectRect.top;
			wDst = this.SlideCurrectRect.right - this.SlideCurrectRect.left;
			hDst = this.SlideCurrectRect.bottom - this.SlideCurrectRect.top;

			dKoefX = wDst / this.m_oLogicDocument.GetWidthMM();
			dKoefY = hDst / this.m_oLogicDocument.GetHeightMM();
		}
		else
		{
			xDst = this.m_oWordControl.m_oNotesApi.OffsetX;
			yDst = - this.m_oWordControl.m_oNotesApi.Scroll;
			dKoefX = g_dKoef_mm_to_pix;
			dKoefY = g_dKoef_mm_to_pix;
		}
		let oTextMatrix = this.TextMatrix;
		for (nIndex = 0; nIndex < PathLng; nIndex++)
		{
			oPath = this.MathTrack.GetPolygon(nIndex);
			this.MathTrack.Draw(overlay, oPath, 0, 0, "#939393", dKoefX, dKoefY, xDst, yDst, oTextMatrix);
			this.MathTrack.Draw(overlay, oPath, 1, 1, "#FFFFFF", dKoefX, dKoefY, xDst, yDst, oTextMatrix);
		}
		for (nIndex = 0, nCount = this.MathTrack.GetSelectPathsCount(); nIndex < nCount; nIndex++)
		{
			oPath =  this.MathTrack.GetSelectPath(nIndex);
			this.MathTrack.DrawSelectPolygon(overlay, oPath, dKoefX, dKoefY, xDst, yDst, oTextMatrix);
		}
	};

	this.Update_MathTrack = function (IsActive, IsContentActive, oMath)
	{
		let PixelError;
		if(this.m_oLogicDocument.IsFocusOnNotes())
		{
			PixelError = 3 / g_dKoef_mm_to_pix;
		}
		else
		{
			PixelError = this.GetMMPerDot(3);
		}
		this.MathTrack.Update(IsActive, IsContentActive, oMath, PixelError);
	};

	this.DrawVerAnchor = function(pageNum, xPos, bIsFromDrawings)
	{
		if (undefined === bIsFromDrawings)
		{
			if (this.m_oWordControl.m_oApi.ShowSmartGuides)
			{
				this.HorVerAnchors.push({Type : 0, Pos : xPos});
			}
			return;
		}

		var _pos = this.ConvertCoordsToCursor(xPos, 0);
		if (_pos.Error === false)
		{
			this.m_oWordControl.m_oOverlayApi.DashLineColor = "#FF0000";
			this.m_oWordControl.m_oOverlayApi.VertLine2(_pos.X);
			this.m_oWordControl.m_oOverlayApi.DashLineColor = "#000000";
		}
	};

	this.DrawHorAnchor = function(pageNum, yPos, bIsFromDrawings)
	{
		if (undefined === bIsFromDrawings)
		{
			if (this.m_oWordControl.m_oApi.ShowSmartGuides)
			{
				this.HorVerAnchors.push({Type : 1, Pos : yPos});
			}
			return;
		}

		var _pos = this.ConvertCoordsToCursor(0, yPos);
		if (_pos.Error === false)
		{
			this.m_oWordControl.m_oOverlayApi.DashLineColor = "#FF0000";
			this.m_oWordControl.m_oOverlayApi.HorLine2(_pos.Y);
			this.m_oWordControl.m_oOverlayApi.DashLineColor = "#000000";
		}
	};

	this.DrawHorVerAnchor = function()
	{
		for (var i = 0; i < this.HorVerAnchors.length; i++)
		{
			var _anchor = this.HorVerAnchors[i];
			if (_anchor.Type == 0)
				this.DrawVerAnchor(0, _anchor.Pos, true);
			else
				this.DrawHorAnchor(0, _anchor.Pos, true);
		}
		this.HorVerAnchors.splice(0, this.HorVerAnchors.length);
	};

	this.StartTrackText = function ()
	{
		this.InlineTextTrackEnabled = true;
		this.InlineTextTrack = null;
		this.InlineTextTrackPage = -1;
		this.InlineTextInNotes = false;
	};

	this.EndTrackText = function (isOnlyMoveTarget)
	{
		this.InlineTextTrackEnabled = false;

		if (true !== isOnlyMoveTarget)
			this.m_oWordControl.m_oLogicDocument.OnEndTextDrag(this.InlineTextTrack, AscCommon.global_keyboardEvent.CtrlKey);
		else if (this.InlineTextTrack)
		{
			var Paragraph = this.InlineTextTrack.Paragraph;
			Paragraph.Cursor_MoveToNearPos(this.InlineTextTrack);
			Paragraph.Document_SetThisElementCurrent(false);

			this.m_oWordControl.m_oLogicDocument.Document_UpdateSelectionState();
			this.m_oWordControl.m_oLogicDocument.Document_UpdateInterfaceState();
			this.m_oWordControl.m_oLogicDocument.Document_UpdateRulersState();
		}

		this.InlineTextTrack = null;
		this.InlineTextTrackPage = -1;
		this.InlineTextInNotes = false;
	};

	this.IsTrackText = function ()
	{
		return this.InlineTextTrackEnabled;
	};

	this.CancelTrackText = function ()
	{
		this.InlineTextTrackEnabled = false;
		this.InlineTextTrack = null;
		this.InlineTextTrackPage = -1;
		this.InlineTextInNotes = false;
	};

	this.DrawTrack = function(type, matrix, left, top, width, height, isLine, canRotate, isNoMove, isDrawHandles)
	{
		this.AutoShapesTrack.DrawTrack(type, matrix, left, top, width, height, isLine, canRotate, isNoMove, isDrawHandles);
	};

	this.DrawTrackSelectShapes = function(x, y, w, h)
	{
		this.AutoShapesTrack.DrawTrackSelectShapes(x, y, w, h);
	};

	this.DrawAdjustment = function(matrix, x, y, bTextWarp)
	{
		this.AutoShapesTrack.DrawAdjustment(matrix, x, y, bTextWarp);
	};

	// metafile
	this.RenderDocument = function(Renderer)
	{

		let SlidesCount = this.GetSlidesCount();
		for (let i = 0; i < SlidesCount; i++)
		{
			Renderer.BeginPage(this.m_oLogicDocument.GetWidthMM(), this.m_oLogicDocument.GetHeightMM());
			this.m_oLogicDocument.DrawPage(i, Renderer);
			Renderer.EndPage();
		}
	};

	this.ToRenderer = function()
	{
		var Renderer                             = new AscCommon.CDocumentRenderer();
		Renderer.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);
		Renderer.IsNoDrawingEmptyPlaceholder     = true;
		Renderer.VectorMemoryForPrint            = new AscCommon.CMemory();
		var old_marks                            = this.m_oWordControl.m_oApi.ShowParaMarks;
		this.m_oWordControl.m_oApi.ShowParaMarks = false;
		this.RenderDocument(Renderer);
		this.m_oWordControl.m_oApi.ShowParaMarks = old_marks;
		var ret                                  = Renderer.Memory.GetBase64Memory();

		// DEBUG
		//console.log(ret);

		return ret;
	};

	this.ToRenderer2    = function()
	{
		var Renderer = new AscCommon.CDocumentRenderer();
		Renderer.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);

		var old_marks                            = this.m_oWordControl.m_oApi.ShowParaMarks;
		this.m_oWordControl.m_oApi.ShowParaMarks = false;

		var ret = "";
		let SlidesCount = this.GetSlidesCount();
		for (var i = 0; i < SlidesCount; i++)
		{
			Renderer.BeginPage(this.m_oLogicDocument.GetWidthMM(), this.m_oLogicDocument.GetHeightMM());
			this.m_oLogicDocument.DrawPage(i, Renderer);
			Renderer.EndPage();

			ret += Renderer.Memory.GetBase64Memory();
			Renderer.Memory.Seek(0);
		}

		this.m_oWordControl.m_oApi.ShowParaMarks = old_marks;
		return ret;
	};

	this.ToRendererPart = function(noBase64, isSelection)
	{
		var watermark = this.m_oWordControl.m_oApi.watermarkDraw;
		let oPresentation = this.m_oWordControl.m_oLogicDocument;

		var pagescount = oPresentation.isVisioDocument ? oPresentation.GetSlidesCount() : oPresentation.Slides.length;

		if (-1 == this.m_lCurrentRendererPage)
		{
			if (watermark)
				watermark.StartRenderer();

			this.m_oDocRenderer                             = new AscCommon.CDocumentRenderer();
			this.m_oDocRenderer.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);
			this.m_oDocRenderer.VectorMemoryForPrint        = new AscCommon.CMemory();
			this.m_lCurrentRendererPage                     = 0;
			this.m_bOldShowMarks                            = this.m_oWordControl.m_oApi.ShowParaMarks;
			this.m_oWordControl.m_oApi.ShowParaMarks        = false;
			this.m_oDocRenderer.IsNoDrawingEmptyPlaceholder = true;
		}

		var start = this.m_lCurrentRendererPage;
		var end   = pagescount - 1;

		var renderer = this.m_oDocRenderer;
		renderer.Memory.Seek(0);
		renderer.VectorMemoryForPrint.ClearNoAttack();
		renderer.DocInfo(this.m_oWordControl.m_oApi.asc_getCoreProps());

		for (var i = start; i <= end; i++)
		{
			if ((true === isSelection && !this.m_oLogicDocument.IsMasterMode()) && !this.m_oWordControl.Thumbnails.isSelectedPage(i))
				continue;

			if (oPresentation.isVisioDocument)
			{
				//todo override
				renderer.BeginPage(this.m_oLogicDocument.GetWidthMM(), this.m_oLogicDocument.GetHeightMM());
				this.m_oLogicDocument.DrawPage(i, renderer);
				renderer.EndPage();
			}
			else
			{
				let oSlide = this.m_oLogicDocument.Slides[i];

				if (!oSlide.isVisible())
					continue;

				renderer.BeginPage(this.m_oLogicDocument.GetWidthMM(), this.m_oLogicDocument.GetHeightMM());
				oSlide.draw(renderer);
				renderer.EndPage();
			}

			if (watermark)
				watermark.DrawOnRenderer(renderer, this.m_oLogicDocument.GetWidthMM(), this.m_oLogicDocument.GetHeightMM());
		}

		if (end == -1)
		{
			renderer.BeginPage(this.m_oLogicDocument.GetWidthMM(), this.m_oLogicDocument.GetHeightMM());
			renderer.EndPage()
		}

		this.m_lCurrentRendererPage = end + 1;

		if (this.m_lCurrentRendererPage >= pagescount)
		{
			if (watermark)
				watermark.EndRenderer();

			this.m_lCurrentRendererPage              = -1;
			this.m_oDocRenderer                      = null;
			this.m_oWordControl.m_oApi.ShowParaMarks = this.m_bOldShowMarks;
		}

		if (noBase64) {
			return renderer.Memory.GetData();
		} else {
			return renderer.Memory.GetBase64Memory();
		}
	};

	// common
	this.Start_CollaborationEditing = function()
	{
		this.IsLockObjectsEnable = true;
		this.m_oWordControl.OnRePaintAttack();
	};

	this.IsMobileVersion = function()
	{
		if (this.m_oWordControl.MobileTouchManager)
			return true;
		return false;
	};

	this.SendChangeDocumentToApi = function(bIsAttack)
	{
		if (bIsAttack || !this.m_bIsSendApiDocChanged)
		{
			this.m_bIsSendApiDocChanged = true;
			this.m_oWordControl.m_oApi.SetDocumentModified(true);

			this.m_oWordControl.m_oApi.sendEvent("asc_onDocumentChanged");
		}
	};

	this.CloseFile = function ()
	{
		this.IsEmptyPresentation = true;
		this.SlideCurrent     = -1;
		this.ClearCachePages();
		this.FirePaint();
	};

	this.IsCursorInTableCur = function(x, y, page, checkArea)
	{
		var _table = this.TableOutlineDr.TableOutline;
		if (_table == null)
			return false;

		if (page != _table.PageNum)
			return false;

		var _dist = this.TableOutlineDr.mover_size * g_dKoef_pix_to_mm;
		_dist *= (100 / this.m_oWordControl.m_nZoomValue);

		var _x = _table.X;
		var _y = _table.Y;
		var _r = _x + _table.W;
		var _b = _y + _table.H;

		if ((x > (_x - _dist)) && (x < _r) && (y > (_y - _dist)) && (y < _b))
		{
			if ((x < _x && y < _y) || (checkArea && (x < _x || y < _y)))
				return true;
		}
		return false;
	};

	this.SetCurrentPage = function(PageIndex)
	{
		if (PageIndex >= this.GetSlidesCount())
			return;
		if (this.SlideCurrent == PageIndex)
			return;

		this.SlideCurrent = PageIndex;
		this.m_oWordControl.SetCurrentPage();
	};

	// вот оооочень важные функции. она выкидывает из кэша неиспользуемые шрифты
	this.CheckFontCache = function()
	{
		var map_used = this.m_oWordControl.m_oLogicDocument.Document_CreateFontMap();

		var _measure_map = g_oTextMeasurer.m_oManager.m_oFontsCache.Fonts;
		var _drawing_map = AscCommon.g_fontManager.m_oFontsCache.Fonts;

		var map_keys = {};
		var api      = this.m_oWordControl.m_oApi;
		for (var i in map_used)
		{
			var key       = AscFonts.GenerateMapId(api, g_fontApplication.GetFontInfoName(map_used[i].Name), map_used[i].Style, map_used[i].Size);
			map_keys[key] = true;
		}

		// а теперь просто пробегаем по кэшам и удаляем ненужное
		for (var i in _measure_map)
		{
			if (map_keys[i] == undefined)
			{
				//_measure_map[i] = undefined;
				delete _measure_map[i];
			}
		}
		for (var i in _drawing_map)
		{
			if (map_keys[i] == undefined)
			{
				//_drawing_map[i] = undefined;
				if (null != _drawing_map[i])
					_drawing_map[i].Destroy();
				delete _drawing_map[i];
			}
		}
	};

	this.CheckFontNeeds = function()
	{
		var map_keys = this.m_oWordControl.m_oLogicDocument.Document_Get_AllFontNames();
		var dstfonts = [];
		for (var i in map_keys)
		{
			dstfonts[dstfonts.length] = new AscFonts.CFont(i);
		}
        AscFonts.FontPickerByCharacter.extendFonts(dstfonts);
		this.m_oWordControl.m_oLogicDocument.Fonts = dstfonts;
		return;
	};

	// comments
	this.GetCommentWidth = function(type)
	{
		var _index = 0;
		if ((type & 0x02) == 0x02)
			_index = 2;
		if ((type & 0x01) == 0x01)
			_index += 1;

		return AscCommon.g_comment_image_offsets[_index][2] * g_dKoef_pix_to_mm * 100 / this.m_oWordControl.m_nZoomValue;
	};

	this.GetCommentHeight = function(type)
	{
		var _index = 0;
		if ((type & 0x02) == 0x02)
			_index = 2;
		if ((type & 0x01) == 0x01)
			_index += 1;

		return AscCommon.g_comment_image_offsets[_index][3] * g_dKoef_pix_to_mm * 100 / this.m_oWordControl.m_nZoomValue;
	};

	// collaborative targets
	this.Collaborative_UpdateTarget      = function(_id, _shortId, _x, _y, _size, _page, _transform, is_from_paint)
	{
		if (is_from_paint !== true)
		{
			this.CollaborativeTargetsUpdateTasks.push([_id, _shortId, _x, _y, _size, _page, _transform]);
			return;
		}

		for (var i = 0; i < this.CollaborativeTargets.length; i++)
		{
			if (_id == this.CollaborativeTargets[i].Id)
			{
				this.CollaborativeTargets[i].CheckPosition(_x, _y, _size, _page, _transform);
				return;
			}
		}
		var _target     = new CDrawingCollaborativeTarget(this);
		_target.Id      = _id;
		_target.ShortId = _shortId;
		_target.CheckPosition(_x, _y, _size, _page, _transform);
		this.CollaborativeTargets[this.CollaborativeTargets.length] = _target;
	};
	this.Collaborative_RemoveTarget      = function(_id)
	{
		var i = 0;
		for (i = 0; i < this.CollaborativeTargets.length; i++)
		{
			if (_id == this.CollaborativeTargets[i].Id)
			{
				this.CollaborativeTargets[i].Remove(this);
				this.CollaborativeTargets.splice(i, 1);
				i--;
			}
		}

		for (i = 0; i < this.CollaborativeTargetsUpdateTasks.length; i++)
		{
			var _tmp = this.CollaborativeTargetsUpdateTasks[i];
			if (_tmp[0] == _id)
			{
				this.CollaborativeTargetsUpdateTasks.splice(i, 1);
				i--;
			}
		}
	};
	this.Collaborative_TargetsUpdate     = function(bIsChangePosition)
	{
		var _len_tasks = this.CollaborativeTargetsUpdateTasks.length;
		var i          = 0;
		for (i = 0; i < _len_tasks; i++)
		{
			var _tmp = this.CollaborativeTargetsUpdateTasks[i];
			this.Collaborative_UpdateTarget(_tmp[0], _tmp[1], _tmp[2], _tmp[3], _tmp[4], _tmp[5], _tmp[6], true);
		}
		if (_len_tasks != 0)
			this.CollaborativeTargetsUpdateTasks.splice(0, _len_tasks);

		if (bIsChangePosition)
		{
			for (i = 0; i < this.CollaborativeTargets.length; i++)
			{
				this.CollaborativeTargets[i].Update(this);
			}
		}
	};
	this.Collaborative_GetTargetPosition = function(UserId)
	{
		for (var i = 0; i < this.CollaborativeTargets.length; i++)
		{
			if (UserId == this.CollaborativeTargets[i].Id)
				return {X : this.CollaborativeTargets[i].HtmlElementX, Y : this.CollaborativeTargets[i].HtmlElementY};
		}

		return null;
	};

	// notes
	this.Notes_GetWidth = function()
	{
		if (!this.m_oWordControl.IsNotesSupported())
			return 0;

		if (!this.m_oWordControl.m_oNotesApi)
			return 0;

		return this.m_oWordControl.m_oNotesApi.GetNotesWidth();
	};

	this.Notes_OnRecalculate = function(slideNum, width, height)
	{
		if (!this.m_oWordControl.IsNotesSupported())
			return;

		if (!this.m_oWordControl.m_oNotesApi)
			return;

		this.m_oWordControl.m_oNotesApi.OnRecalculateNote(slideNum, width, height);
	};

	// mouse events
	this.checkMouseDown_Drawing = function (pos)
	{
        var oWordControl = this.m_oWordControl;
		var bIsReturn = false;

		if (!this.isButtonsDisabled() && this.placeholders.onPointerDown(pos, this.SlideCurrectRect, this.m_oLogicDocument.GetWidthMM(), this.m_oLogicDocument.GetHeightMM()))
		{
			bIsReturn = true;
			this.m_oWordControl.onMouseUpMainSimple();
		}

        if (bIsReturn)
		{
            oWordControl.OnUpdateOverlay();
            oWordControl.EndUpdateOverlay();
		}
		return bIsReturn;
	};

	this.checkMouseMove_Drawing = function (pos)
	{
		var oWordControl = this.m_oWordControl;
        var bIsReturn = false;

		if (this.InlineTextTrackEnabled)
		{
			if (-1 != oWordControl.m_oTimerScrollSelect)
			{
				clearInterval(oWordControl.m_oTimerScrollSelect);
				oWordControl.m_oTimerScrollSelect = -1;
			}

			this.InlineTextTrack = oWordControl.m_oLogicDocument.Get_NearestPos(pos.Page, pos.X, pos.Y, pos.isNotes);
			this.InlineTextTrackPage = pos.Page;
			this.InlineTextInNotes = pos.isNotes ? true : false;

            bIsReturn = true;
		}
		else
		{
			if(!AscCommon.global_mouseEvent.IsLocked)
			{
				if (!this.isButtonsDisabled() && this.placeholders.onPointerMove(pos, this.SlideCurrectRect, this.m_oLogicDocument.GetWidthMM(), this.m_oLogicDocument.GetHeightMM()))
				{
					oWordControl.OnUpdateOverlay();
					oWordControl.EndUpdateOverlay();
					bIsReturn = true;
				}
			}
		}

        if (bIsReturn)
        {
            oWordControl.OnUpdateOverlay();
            oWordControl.EndUpdateOverlay();
        }
        return bIsReturn;
	};

	this.isButtonsDisabled = function()
	{
		return Asc.editor.isEyedropperStarted() || Asc.editor.isDrawInkMode();
	};

	this.checkMouseUp_Drawing = function (pos)
	{
		var oWordControl = this.m_oWordControl;
        var bIsReturn = false;

		if (this.InlineTextTrackEnabled)
		{
			this.InlineTextTrack = oWordControl.m_oLogicDocument.Get_NearestPos(pos.Page, pos.X, pos.Y, pos.isNotes);
			this.InlineTextTrackPage = pos.Page;
			this.InlineTextInNotes = pos.isNotes ? true : false;
			this.EndTrackText();

            bIsReturn = true;
		}
        else if (!this.isButtonsDisabled() && this.placeholders.onPointerUp(pos, this.SlideCurrectRect, this.m_oLogicDocument.GetWidthMM(), this.m_oLogicDocument.GetHeightMM()))
            bIsReturn = true;

        if (bIsReturn)
        {
            oWordControl.OnUpdateOverlay();
            oWordControl.EndUpdateOverlay();
        }
        return bIsReturn;
	};

	this.GetCurSld = function()
	{
		return this.m_oWordControl.Thumbnails.GetCurSld();
	};


	this.OnStartImagesLoading = function(aImages)
	{
		this.LoadingImages = [].concat(aImages);
	};
	this.CheckRasterImageOnScreen = function (src)
	{
		const oPresentation = oThis.m_oWordControl.m_oLogicDocument;
		let oCurSlide = oPresentation.GetCurrentSlide();
		if(!oCurSlide)
			return;
		let bRedraw = false;
		let sCheckImage = AscCommon.getFullImageSrc2(src);
		let nCurIdx = oPresentation.GetSlideIndex(oCurSlide);
		if(oCurSlide.checkImageDraw(sCheckImage))
		{
			this.OnRecalculateSlide(nCurIdx);
			bRedraw = true;
		}
		//check visible thumbnails
		let oWordControl = oThis.m_oWordControl;
		let oThumbnails = oWordControl.Thumbnails;
		let nStart = oThumbnails.m_lDrawingFirst;
		let nEnd = oThumbnails.m_lDrawingEnd;
		for(let nIdx = nStart; nIdx <= nEnd; ++nIdx)
		{
			if(nIdx !== nCurIdx)
			{
				let oSlide = oPresentation.GetSlide(nIdx);
				if(oSlide && oSlide.checkImageDraw(sCheckImage))
				{
					this.OnRecalculateSlide(nIdx);
					bRedraw = true;
				}
			}
		}
		if(bRedraw)
		{
			this.OnEndRecalculate();
		}

		for(let nImg = 0; nImg < this.LoadingImages.length; ++nImg)
		{
			if(this.LoadingImages[nImg] === sCheckImage)
			{
				this.LoadingImages.splice(nImg, 1);
				break;
			}
		}

		if(this.LoadingImages.length === 0)
		{
			oWordControl.CheckLayouts(true);
			oPresentation.SendThemesThumbnails();
			this.DrawImageTextureFillShape(this.LastDrawingUrl);
			this.DrawImageTextureFillSlide(this.LastDrawingUrlSlide);
			this.DrawImageTextureFillTextArt(this.LastDrawingUrlTextArt);
		}
	};

}

function CThPage()
{
	this.PageIndex = -1;

	this.left   = 0;
	this.top    = 0;
	this.right  = 0;
	this.bottom = 0;

	this.IsRecalc    = false;
	this.cachedImage = null;

	this.IsSelected = false;
	this.IsFocused  = false;
	this.IsLocked   = false;
}

CThPage.prototype.Draw = function(context, xDst, yDst, wDst, hDst)
{
	if (wDst <= 0 || hDst <= 0)
		return;

	if (null != this.cachedImage)
	{
		// потом посмотреть на кусочную отрисовку
		context.drawImage(this.cachedImage.image, xDst, yDst, wDst, hDst);
	}
	else
	{
		context.fillStyle = "#FFFFFF";
		context.fillRect(xDst, yDst, wDst, hDst);
	}
};
CThPage.prototype.Hit = function(x, y)
{
	return y >= this.top && y <= this.bottom;
};
CThPage.prototype.SetRecalc = function(bValue)
{
	this.IsRecalc = bValue;
};
CThPage.prototype.GetRecalc = function(bValue)
{
	return this.IsRecalc;
};
CThPage.prototype.SetLocked = function(bValue)
{
	this.IsLocked = bValue;
};
CThPage.prototype.SetSelected = function(bValue)
{
	this.IsSelected = bValue;
};
CThPage.prototype.SetFocused = function(bValue)
{
	this.IsFocused = bValue;
};
CThPage.prototype.IsPageSelected = function()
{
	return this.IsSelected;
};
CThPage.prototype.GetPosition = function()
{
	return this.Position;
};

function DrawBackground(graphics, unifill, w, h)
{
	// первым делом рисуем белый рект!
	if (true)
	{
		// ну какой-то бэкграунд должен быть
		graphics.SetIntegerGrid(false);

		var _l = 0;
		var _t = 0;
		var _r = (0 + w);
		var _b = (0 + h);

		graphics._s();
		graphics._m(_l, _t);
		graphics._l(_r, _t);
		graphics._l(_r, _b);
		graphics._l(_l, _b);
		graphics._z();

		graphics.b_color1(255, 255, 255, 255);
		graphics.df();
		graphics._e();
	}

	if (unifill == null || unifill.fill == null)
		return;

	graphics.SetIntegerGrid(false);

	var _shape = {};

	_shape.brush           = unifill;
	_shape.pen             = null;
	_shape.TransformMatrix = new AscCommon.CMatrix();
	_shape.extX            = w;
	_shape.extY            = h;
	_shape.check_bounds    = function(checker)
	{
		checker._s();
		checker._m(0, 0);
		checker._l(this.extX, 0);
		checker._l(this.extX, this.extY);
		checker._l(0, this.extY);
		checker._z();
		checker._e();
	};

	var shape_drawer = new AscCommon.CShapeDrawer();
	shape_drawer.fromShape2(_shape, graphics, null);
	shape_drawer.draw(null);
}

function CMouseDownTrack(Thumbnails)
{
	this.Thumbnails = Thumbnails;
	this.SetDefault();
}
CMouseDownTrack.prototype.Check = function()
{
	if(this.Started)
	{
		let aThPages = this.Thumbnails.m_arrPages;
		if(!aThPages[this.Page])
		{
			this.Reset();
		}
	}
};
CMouseDownTrack.prototype.SetDefault = function()
{
	this.Started = false;
	this.StartedSimple = true;
	this.Page = -1;
	this.X = -1;
	this.Y = -1;
	this.Position = -1;
};
CMouseDownTrack.prototype.Reset = function()
{
	this.SetDefault();
};
CMouseDownTrack.prototype.Start = function(Page, X, Y)
{
	this.Started = true;
	this.StartedSimple = true;
	this.Page = Page;
	this.X = X;
	this.Y = Y;
};
CMouseDownTrack.prototype.IsStarted = function()
{
	this.Check();
	return this.Started;
};
CMouseDownTrack.prototype.IsSimple = function()
{
	this.Check();
	return this.StartedSimple;
};
CMouseDownTrack.prototype.IsDragged = function()
{
	return this.IsStarted() && !this.IsSimple() && this.GetPosition() !== -1;
};
CMouseDownTrack.prototype.ResetSimple = function(Position)
{
	this.StartedSimple = false;
	this.Position = Position;
};
CMouseDownTrack.prototype.SetPosition = function(Position)
{
	this.Position = Position;
};
CMouseDownTrack.prototype.GetPosition = function()
{
	this.Check();
	return this.Position;
};
CMouseDownTrack.prototype.GetPage = function()
{
	this.Check();
	return this.Page;
};
CMouseDownTrack.prototype.GetX = function()
{
	this.Check();
	return this.X;
};
CMouseDownTrack.prototype.GetY = function()
{
	this.Check();
	return this.Y;
};

CMouseDownTrack.prototype.IsMoved = function(X, Y)
{
	this.Check();
	if (Math.abs(this.X - X) > 10 || Math.abs(this.Y - Y) > 10)
		return true;
	return false;
};
CMouseDownTrack.prototype.IsSamePos = function()
{
	this.Check();
	return this.Position === this.Page || this.Position === (this.Page + 1);
};




function CThPosBase(oThumbnails)
{
	this.thumbnails = oThumbnails;
}
CThPosBase.prototype.Increment = function() {};
CThPosBase.prototype.Decrement = function() {};
CThPosBase.prototype.IsLess = function(oPos)
{
	return false;
};
CThPosBase.prototype.IsGreater = function(oPos)
{
	return false;
};
CThPosBase.prototype.IsLessOrEqual = function(oPos)
{
	return this.IsLess(oPos) || this.IsEqual(oPos);
};
CThPosBase.prototype.IsGreaterOrEqual = function(oPos)
{
	return this.IsGreater(oPos) || this.IsEqual(oPos);
};
CThPosBase.prototype.IsEqual = function(oPos)
{
	return false;
};
CThPosBase.prototype.Copy = function()
{
	return new CThPosBase(this.thumbnails);
};
CThPosBase.prototype.GetPage = function()
{
	return this.thumbnails.GetPage(this);
};
CThPosBase.prototype.GetSld = function()
{
	return null;
};
CThPosBase.prototype.GetPresentation = function()
{
	return Asc.editor.WordControl.m_oLogicDocument;
};

function CSlideThPos(oThumbnails)
{
	CThPosBase.call(this, oThumbnails);
	this.Idx = -1;
}
AscFormat.InitClassWithoutType(CSlideThPos, CThPosBase);
CSlideThPos.prototype.Check = function()
{
	let aThs = this.thumbnails.slides;
	let nMaxIdx = aThs.length - 1;
	let nMinIdx = Math.min(nMaxIdx, 0);
	this.Idx = Math.max(nMinIdx, Math.min(this.Idx, nMaxIdx));
};
CSlideThPos.prototype.Increment = function()
{
	++this.Idx;
	this.Check();
};
CSlideThPos.prototype.Decrement = function()
{
	--this.Idx;
	this.Check();
};
CSlideThPos.prototype.IsLess = function(oPos)
{
	return this.Idx < oPos.Idx;
};
CSlideThPos.prototype.IsGreater = function(oPos)
{
	return this.Idx > oPos.Idx;
};
CSlideThPos.prototype.IsEqual = function(oPos)
{
	return this.Idx === oPos.Idx;
};
CSlideThPos.prototype.Copy = function()
{
	let oPos = new CSlideThPos(this.thumbnails);
	oPos.Idx = this.Idx;
	return oPos;
};
CSlideThPos.prototype.GetSld = function()
{
	return this.GetPresentation().Slides[this.Idx];
};


function CMasterThPos(oThumbnails)
{
	CThPosBase.call(this, oThumbnails);
	this.MasterIdx = -1;
	this.LayoutIdx = -1;
}
AscFormat.InitClassWithoutType(CMasterThPos, CThPosBase);
CMasterThPos.prototype.Check = function()
{
	let nOldMasterIdx = this.MasterIdx;
	let aThs = this.thumbnails.masters;
	let nMaxMasterIdx = aThs.length - 1;
	let nMinMasterIdx = Math.min(nMaxMasterIdx, 0);
	this.MasterIdx = Math.max(nMinMasterIdx, Math.min(this.MasterIdx, nMaxMasterIdx));
	if(this.MasterIdx !== nOldMasterIdx)
	{
		this.LayoutIdx = -1;
		return;
	}
	if(this.LayoutIdx === -1)
	{
		return;
	}
	let oMasterTh = aThs[this.MasterIdx];
	let aLayoutsTh = oMasterTh.layoutPages;
	let nMaxLayoutIdx = aLayoutsTh.length - 1;
	let nMinLayoutIdx = Math.min(nMaxLayoutIdx, 0);
	this.LayoutIdx = Math.max(nMinLayoutIdx, Math.min(this.LayoutIdx, nMaxLayoutIdx));
};
CMasterThPos.prototype.Increment = function()
{
	let oMasterThs = this.thumbnails.GetMasterThumbnails(this.MasterIdx);
	let nLayoutsCount = oMasterThs.GetLayoutsCount();
	if(this.LayoutIdx + 1 < nLayoutsCount)
	{
		++this.LayoutIdx;
	}
	else
	{
		if(this.MasterIdx + 1 < this.thumbnails.GetMastersCount())
		{
			++this.MasterIdx;
			this.LayoutIdx = - 1;
		}
	}
};
CMasterThPos.prototype.Decrement = function()
{
	if(this.LayoutIdx > -1)
	{
		--this.LayoutIdx;
		return;
	}
	this.MasterIdx = Math.max(-1, this.MasterIdx - 1);
	let oMasterThs = this.thumbnails.GetMasterThumbnails(this.MasterIdx);
	if(oMasterThs)
	{
		this.LayoutIdx = oMasterThs.GetLayoutsCount() - 1;
	}
};
CMasterThPos.prototype.IsLess = function(oPos)
{
	if(this.MasterIdx > oPos.MasterIdx) return false;
	if(this.MasterIdx < oPos.MasterIdx) return true;
	return this.LayoutIdx < oPos.LayoutIdx;
};
CMasterThPos.prototype.IsGreater = function(oPos)
{
	return !this.IsLess(oPos) && !this.IsEqual(oPos);
};
CMasterThPos.prototype.IsEqual = function(oPos)
{
	return (oPos.MasterIdx === this.MasterIdx && oPos.LayoutIdx === this.LayoutIdx);
};
CMasterThPos.prototype.Copy = function()
{
	let oPos = new CMasterThPos(this.thumbnails);
	oPos.MasterIdx = this.MasterIdx;
	oPos.LayoutIdx = this.LayoutIdx;
	return oPos;
};
CMasterThPos.prototype.GetSld = function()
{
	let oPresentation = this.GetPresentation();
	let oMaster = oPresentation.slideMasters[this.MasterIdx];
	if(!oMaster) return null;
	if(this.LayoutIdx === -1) {
		return oMaster;
	}
	return oMaster.sldLayoutLst[this.LayoutIdx];
};

function CThumbnailsBase() {
	this.curPos = this.GetIterator();
	this.firstVisible = this.GetIterator();
	this.lastVisible = this.GetIterator();
}
CThumbnailsBase.prototype.GetIterator = function()
{
	throw new Error();
};
CThumbnailsBase.prototype.SetSlideRecalc = function (idx) {
	let oTh = this.GetPage(idx);
	if(oTh)
	{
		oTh.SetRecalc(true);
	}
};
CThumbnailsBase.prototype.LockSlide = function(idx) {
	let oTh = this.GetPage(idx);
	if(oTh)
	{
		oTh.SetLocked(true);
	}
};
CThumbnailsBase.prototype.UnLockSlide = function(idx) {
	let oTh = this.GetPage(idx);
	if(oTh)
	{
		oTh.SetLocked(false);
	}
};
CThumbnailsBase.prototype.GetPage = function (oIdx) {
	return null;
};
CThumbnailsBase.prototype.SelectAll = function()
{
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage) {
		oPage.SetSelected(true);
		return false;
	});
};
CThumbnailsBase.prototype.FindThUnderCursor = function(X, Y)
{
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage) {
		if(oPage.Hit(X, Y))
		{
			return oPage;
		}
	});
};
CThumbnailsBase.prototype.ResetSelection = function()
{
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage) {
		oPage.SetSelected(false);
		return false;
	});
};
CThumbnailsBase.prototype.SetAllFocused = function(bValue)
{
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage) {
		oPage.SetFocused(bValue);
		return false;
	});
};

CThumbnailsBase.prototype.SetAllRecalc = function (bValue) {
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage) {
		oPage.SetRecalc(bValue);
		return false;
	});
};
CThumbnailsBase.prototype.SelectRange = function(oStartPos, oEndPos)
{
	this.SetSelectionRange(oStartPos, oEndPos, true);
};

CThumbnailsBase.prototype.SetSelectionRange = function(oStartPos, oEndPos, bValue)
{
	this.IteratePages(oStartPos, oEndPos, function(oPage) {
		oPage.SetSelected(bValue);
		return false;
	});
};
CThumbnailsBase.prototype.GetStartPos = function()
{
	throw new Error();
};
CThumbnailsBase.prototype.GetEndPos = function()
{
	throw new Error();
};
CThumbnailsBase.prototype.IteratePages = function(oStartPos, oEndPos, fAction)
{
	if(oStartPos.IsLessOrEqual(oEndPos))
	{
		for(let oPos = oStartPos.Copy(); oPos.IsLessOrEqual(oEndPos); oPos.Increment())
		{
			let oPage = oPos.GetPage();
			if(!oPage)
				return null;
			if(oPage)
			{
				let oResult = fAction(oPos.GetPage());
				if(oResult)
				{
					return oResult;
				}
			}
		}
	}
	return null;
};
CThumbnailsBase.prototype.GetPagesCount = function()
{
	let nCount = 0;
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage) {
		nCount++;
		return false;
	});
	return nCount;
};
CThumbnailsBase.prototype.ConvertCoords2 = function(X, Y)
{

	let pages_count = this.GetPagesCount();
	if (0 == pages_count)
		return -1;

	let oStartPos = this.GetStartPos();
	let oEndPos = this.GetEndPos();
	let _min = Math.abs(Y - this.GetStartPos().GetPage().top);
	let _MinPositionPage = 0;

	this.IteratePages(oStartPos, oEndPos, function(oPage) {

		let _min1 = Math.abs(Y - oPage.top);
		let _min2 = Math.abs(Y - oPage.bottom);

		if (_min1 < _min)
		{
			_min = _min1;
			_MinPositionPage = oPage.GetPosition();
		}
		if (_min2 < _min)
		{
			_min = _min2;
			_MinPositionPage = oPage.GetPosition().Increment();
		}
	});

	return _MinPositionPage;
};
CThumbnailsBase.prototype.GetPagePosition = function(oPageIdx)
{
	let oPage = this.GetPage(oPageIdx);
	if(!oPage) return null;
	let _ret = {
		X: AscCommon.AscBrowser.convertToRetinaValue(oPage.left),
		Y: AscCommon.AscBrowser.convertToRetinaValue(oPage.top),
		W: AscCommon.AscBrowser.convertToRetinaValue(oPage.right - oPage.left),
		H: AscCommon.AscBrowser.convertToRetinaValue(oPage.bottom - oPage.top)
	};
	return _ret;
};
CThumbnailsBase.prototype.GetCurSld = function()
{
	return this.curPos.GetSld();
};
CThumbnailsBase.prototype.IndexToPos = function(nIndex)
{
	return null;
};
CThumbnailsBase.prototype.PosToIndex = function(oPos)
{
	return null;
};

function CSlidesThumbnails() {
	CThumbnailsBase.call(this);
	this.slides = [];
}
AscFormat.InitClassWithoutType(CSlidesThumbnails, CThumbnailsBase);
CSlidesThumbnails.prototype.GetIterator = function()
{
	return new CSlideThPos(this);
};
CSlidesThumbnails.prototype.clear = function ()
{
	this.slides.length = 0;
	this.cusSlide = -1;
	this.firstVisible = -1;
	this.lastVisible = -1;
};
CSlidesThumbnails.prototype.GetStartPos = function()
{
	let oPos = this.GetIterator();
	oPos.Idx = Math.min(0, this.slides.length - 1);
	return oPos;
};
CSlidesThumbnails.prototype.GetEndPos = function()
{
	let oPos = this.GetIterator();
	oPos.Idx = this.slides.length - 1;
	return oPos;
};
CSlidesThumbnails.prototype.IndexToPos = function(nIndex)
{
	let oPos = this.GetIterator();
	oPos.Idx = nIndex;
	return oPos;
};
CSlidesThumbnails.prototype.PosToIndex = function(oPos)
{
	if(!oPos) return -1;
	return oPos.Idx;
};


function CMastersThumbnails()
{
	CThumbnailsBase.call(this);
	this.masters = [];
}
AscFormat.InitClassWithoutType(CMastersThumbnails, CThumbnailsBase);
CMastersThumbnails.prototype.clear = function ()
{
	for(let nIdx = 0; nIdx < this.masters.length; ++nIdx)
	{
		this.masters[nIdx].preDelete();
	}
	this.masters.length = 0;
};
CMastersThumbnails.prototype.GetPage = function(oPos)
{
	if(!Asc.Format.isRealObject(oPos)) return null;
	let oMasterTh = this.masters[oPos.MasterIdx];
	if(!oMasterTh) return null;
	return oMasterTh.GetPage(oPos);
};
CMastersThumbnails.prototype.GetMastersCount = function ()
{
	return this.masters.length;
};
CMastersThumbnails.prototype.GetMasterThumbnails = function (nIdx)
{
	return this.masters[nIdx] || null;
};
CMastersThumbnails.prototype.GetIterator = function()
{
	return new CMasterThPos(this);
};
CMastersThumbnails.prototype.GetStartPos = function()
{
	let oPos = this.GetIterator();
	oPos.MasterIdx = Math.min(0, this.masters.length - 1);
};
CMastersThumbnails.prototype.GetEndPos = function()
{
	let oPos = this.GetIterator();
	oPos.MasterIdx = this.masters.length - 1;
	let oMasterThs = this.GetMasterThumbnails(oPos.MasterIdx);
	if(oMasterThs)
	{
		oPos.LayoutIdx = oMasterThs.GetLayoutsCount() - 1;
	}
	return oPos;
};
CMastersThumbnails.prototype.IndexToPos = function(nIndex)
{
	let nIdx = 0;
	let oPosition = null;
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage, oPos) {
		if(nIdx === nIndex) {
			oPosition = oPos;
			return true;
		}
		++nIdx;
	});
	return oPosition;
};
CSlidesThumbnails.prototype.PosToIndex = function(oPos)
{
	let nIdx = -1;
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage, oPosition) {
		++nIdx;
		if(oPos.IsEqual(oPosition)) {
			return true;
		}
		return false;
	});
	return nIdx;
};

function CMasterThumbnails(oMaster)
{
	this.master = oMaster;

	this.masterPage = null;//new CThPage()
	this.layoutPages = [];//new CThPage()[]

	this.id = this.master.Id;

	let oThis = this;
	Asc.editor.attachEvent("slideMasterUpdate", function (oMaster) {
		if(oMaster === this.master) {
			oThis.update();
		}
	}, this.id);
	this.update();
}
CMasterThumbnails.prototype.GetPage = function(oPos)
{
	if(!oPos) return null;
	return this.layoutPages[oPos.LayoutIdx] || this.masterPage;
};
CMasterThumbnails.prototype.GetLayoutsCount = function()
{
	return this.layoutPages.length;
};
CMasterThumbnails.prototype.preDelete = function() {
	Asc.editor.detachEvent("slideMasterUpdate", this.id);
};
CMasterThumbnails.prototype.update = function() {
	if(!this.masterPage) {
		this.masterPage = new CThPage();
	}
	let aLayouts = this.master.sldLayoutLst;
	this.layoutPages.length = 0;
	for(let nIdx = 0; nIdx < aLayouts.length; ++nIdx) {
		this.layoutPages.push(new CThPage());
	}
};

function CThumbnailsManager()
{
	this.isInit = false;
	this.lastPixelRatio = 0;
	this.m_oFontManager = new AscFonts.CFontManager();

	this.m_bIsScrollVisible = true;

	// cached measure
	this.DigitWidths = [];

	// skin
	this.backgroundColor = "#B0B0B0";
	this.overColor = "#D3D3D3";
	this.selectColor = "#FFD86B";
	this.selectoverColor = "#FFE065";

	this.const_offset_x = 0;
	this.const_offset_y = 0;
	this.const_offset_r = 4;
	this.const_offset_b = 0;
	this.const_border_w = 4;

	// size & position
	this.SlideWidth = 297;
	this.SlideHeight = 210;

	this.m_dScrollY = 0;
	this.m_dScrollY_max = 0;

	this.m_bIsVisible = false;


	this.m_bIsUpdate = false;

	//regular mode
	this.thumbnails = new CSlidesThumbnails();

	this.m_nCurrentPage = -1;
	this.m_arrPages = [];
	this.m_lDrawingFirst = -1;
	this.m_lDrawingEnd = -1;



	this.bIsEmptyDrawed = false;

	this.m_oCacheManager = new CCacheManager(true);

	this.FocusObjType = FOCUS_OBJECT_MAIN;
	this.LockMainObjType = false;

	this.SelectPageEnabled = true;

	this.MouseDownTrack = new CMouseDownTrack(this);

	this.MouseTrackCommonImage = null;

	this.MouseThumbnailsAnimateScrollTopTimer = -1;
	this.MouseThumbnailsAnimateScrollBottomTimer = -1;

	this.ScrollerHeight = 0;

	this.m_oWordControl = null;
	var oThis = this;


	this.IsMasterMode = function() {
		return Asc.editor.isMasterMode();
	};
	// init
	this.Init = function()
	{
		this.isInit = true;
		this.m_oFontManager.Initialize(true);
		this.m_oFontManager.SetHintsProps(true, true);

		this.InitCheckOnResize();

		this.MouseTrackCommonImage = document.createElement("canvas");

		var _im_w = 9;
		var _im_h = 9;

		this.MouseTrackCommonImage.width = _im_w;
		this.MouseTrackCommonImage.height = _im_h;

		var _ctx = this.MouseTrackCommonImage.getContext('2d');
		var _data = _ctx.createImageData(_im_w, _im_h);
		var _pixels = _data.data;

		var _ind = 0;
		for (var j = 0; j < _im_h; ++j)
		{
			var _off1 = (j > (_im_w >> 1)) ? (_im_w - j - 1) : j;
			var _off2 = _im_w - _off1 - 1;

			for (var r = 0; r < _im_w; ++r)
			{
				if (r <= _off1 || r >= _off2)
				{
					_pixels[_ind++] = 183;
					_pixels[_ind++] = 183;
					_pixels[_ind++] = 183;
					_pixels[_ind++] = 255;
				} else
				{
					_pixels[_ind + 3] = 0;
					_ind += 4;
				}
			}
		}

		_ctx.putImageData(_data, 0, 0);
	};

	this.InitCheckOnResize = function()
	{
		if (!this.isInit)
			return;

		if (Math.abs(this.lastPixelRatio - AscCommon.AscBrowser.retinaPixelRatio) < 0.01)
			return;

		this.lastPixelRatio = AscCommon.AscBrowser.retinaPixelRatio;
		this.SetFont({
			FontFamily: {Name: "Arial", Index: -1},
			Italic: false,
			Bold: false,
			FontSize: Math.round(10 * AscCommon.AscBrowser.retinaPixelRatio)
		});

		// измеряем все цифры
		for (var i = 0; i < 10; i++)
		{
			var _meas = this.m_oFontManager.MeasureChar(("" + i).charCodeAt(0));
			if (_meas)
				this.DigitWidths[i] = _meas.fAdvanceX * 25.4 / 96;
			else
				this.DigitWidths[i] = 10;
		}

		this.const_offset_y = Math.round(this.lastPixelRatio * 17);
		this.const_offset_b = this.const_offset_y;
		this.const_offset_r = Math.round(this.lastPixelRatio * 10);
		this.const_border_w = Math.round(this.lastPixelRatio * 7);
	};

	this.initEvents = function()
	{
		var control = this.m_oWordControl.m_oThumbnails.HtmlElement;

		AscCommon.addMouseEvent(control, "down", this.onMouseDown);
		AscCommon.addMouseEvent(control, "move", this.onMouseMove);
		AscCommon.addMouseEvent(control, "up", this.onMouseUp);

		control.onmouseout = this.onMouseLeave;

		control.onmousewheel = this.onMouseWhell;
		if (control.addEventListener)
		{
			control.addEventListener("DOMMouseScroll", this.onMouseWhell, false);
		}
	};

	this.GetPageByPos = function(oPos)
	{
		return this.thumbnails.GetPage(oPos);
	};

	this.GetSlidesCount = function()
	{
		return Asc.editor.getCountSlides();
	};

	this.SetFont = function(font)
	{
		font.FontFamily.Name = g_fontApplication.GetFontFileWeb(font.FontFamily.Name, 0).m_wsFontName;

		if (-1 == font.FontFamily.Index)
			font.FontFamily.Index = AscFonts.g_map_font_index[font.FontFamily.Name];

		if (font.FontFamily.Index == undefined || font.FontFamily.Index == -1)
			return;

		var bItalic = true === font.Italic;
		var bBold   = true === font.Bold;

		var oFontStyle = FontStyle.FontStyleRegular;
		if (!bItalic && bBold)
			oFontStyle = FontStyle.FontStyleBold;
		else if (bItalic && !bBold)
			oFontStyle = FontStyle.FontStyleItalic;
		else if (bItalic && bBold)
			oFontStyle = FontStyle.FontStyleBoldItalic;

		g_fontApplication.LoadFont(font.FontFamily.Name, AscCommon.g_font_loader, this.m_oFontManager, font.FontSize, oFontStyle, 96, 96);
	};

	this.SetSlideRecalc = function (nIdx)
	{
		this.thumbnails.SetSlideRecalc(nIdx);
	};

	this.LockSlide = function(nIdx)
	{

		this.thumbnails.LockSlide(nIdx);
		this.OnUpdateOverlay();
	};

	this.UnLockSlide = function(nIdx)
	{
		this.thumbnails.UnLockSlide(nIdx);
		this.OnUpdateOverlay();
	};
	this.SelectAll = function()
	{
		this.thumbnails.SelectAll();
		this.OnUpdateOverlay();
	};

	this.GetFirstSelectedType = function()
	{
		return this.m_oWordControl.m_oLogicDocument.GetFirstSelectedType();
	};
	this.GetSlideType = function(nIdx)
	{
		return this.m_oWordControl.m_oLogicDocument.GetSlideType(nIdx);
	};
	this.IsMixedSelection = function()
	{
		return this.m_oWordControl.m_oLogicDocument.IsMixedSelection();
	};

	// events
	this.onMouseDown = function(e)
	{
		let mobileTouchManager = oThis.m_oWordControl ? oThis.m_oWordControl.MobileTouchManagerThumbnails : null;
		if (mobileTouchManager && mobileTouchManager.checkTouchEvent(e))
		{
			mobileTouchManager.startTouchingInProcess();
			let res = mobileTouchManager.mainOnTouchStart(e);
			mobileTouchManager.stopTouchingInProcess();
			return res;
		}

		if (oThis.m_oWordControl)
		{
			oThis.m_oWordControl.m_oApi.checkInterfaceElementBlur();
			oThis.m_oWordControl.m_oApi.checkLastWork();

			// после fullscreen возможно изменение X, Y после вызова Resize.
			oThis.m_oWordControl.checkBodyOffset();
		}
		
		AscCommon.stopEvent(e);

		if (AscCommon.g_inputContext && AscCommon.g_inputContext.externalChangeFocus())
			return;

		var control = oThis.m_oWordControl.m_oThumbnails.HtmlElement;
		if (global_mouseEvent.IsLocked == true && global_mouseEvent.Sender != control)
		{
			// кто-то зажал мышку. кто-то другой
			return false;
		}

		AscCommon.check_MouseDownEvent(e);
		global_mouseEvent.LockMouse();
		const oPresentation = oThis.m_oWordControl.m_oLogicDocument;
		let nStartHistoryIndex = oPresentation.History.Index;
		function checkSelectionEnd()
		{
			if(nStartHistoryIndex === oPresentation.History.Index)
			{
				Asc.editor.sendEvent("asc_onSelectionEnd");
			}
		}

		oThis.m_oWordControl.m_oApi.sync_EndAddShape();
		if (global_mouseEvent.Sender != control)
		{
			// такого быть не должно
			return false;
		}

		if (global_mouseEvent.Button == undefined)
			global_mouseEvent.Button = 0;

		oThis.SetFocusElement(FOCUS_OBJECT_THUMBNAILS);

		var pos = oThis.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);
		if (pos.Page == -1)
		{
			if (global_mouseEvent.Button == 2)
			{
				oThis.showContextMenu(false);
			}
			checkSelectionEnd();
			return false;
		}

		if (global_keyboardEvent.CtrlKey && !oThis.m_oWordControl.m_oApi.isReporterMode)
		{
			if (oThis.m_arrPages[pos.Page].IsSelected === true)
			{
				oThis.m_arrPages[pos.Page].IsSelected = false;
				var arr = oThis.GetSelectedArray();
				if (0 == arr.length)
				{
					oThis.m_arrPages[pos.Page].IsSelected = true;
					oThis.ShowPage(pos.Page);
				} else
				{
					oThis.OnUpdateOverlay();

					oThis.SelectPageEnabled = false;
					oThis.m_oWordControl.GoToPage(arr[0]);
					oThis.SelectPageEnabled = true;

					oThis.ShowPage(arr[0]);
				}
			} else
			{
				if(oThis.GetFirstSelectedType() === oThis.GetSlideType(pos.Page))
				{
					oThis.m_arrPages[pos.Page].IsSelected = true;
					oThis.OnUpdateOverlay();

					oThis.SelectPageEnabled = false;
					oThis.m_oWordControl.GoToPage(pos.Page);
					oThis.SelectPageEnabled = true;
					oThis.ShowPage(pos.Page);
				}
			}
		} else if (global_keyboardEvent.ShiftKey && !oThis.m_oWordControl.m_oApi.isReporterMode)
		{

			var pages_count = oThis.m_arrPages.length;
			for (var i = 0; i < pages_count; i++)
			{
				oThis.m_arrPages[i].IsSelected = false;
			}

			var _max = pos.Page;
			var _min = oThis.m_oWordControl.m_oDrawingDocument.SlideCurrent;
			if (_min > _max)
			{
				var _temp = _max;
				_max = _min;
				_min = _temp;
			}

			let nSlideType = oThis.GetSlideType(_min);
			for (var i = _min; i <= _max; i++)
			{
				if(nSlideType === oThis.GetSlideType(i))
				{
					oThis.m_arrPages[i].IsSelected = true;
				}
			}

			oThis.OnUpdateOverlay();
			oThis.ShowPage(pos.Page);
			oPresentation.Document_UpdateInterfaceState();
		} else if (0 == global_mouseEvent.Button || 2 == global_mouseEvent.Button)
		{
			
			let isMouseDownOnAnimPreview = false;
			if (0 == global_mouseEvent.Button)
			{
				if (oThis.m_arrPages[pos.Page].animateLabelRect) // click on the current star, preview animation button, slide transition
				{
					let animateLabelRect = oThis.m_arrPages[pos.Page].animateLabelRect;
					if (pos.X >= animateLabelRect.minX && pos.X <= animateLabelRect.maxX && pos.Y >= animateLabelRect.minY && pos.Y <= animateLabelRect.maxY) 
					isMouseDownOnAnimPreview = true
				} 
				
				if (!isMouseDownOnAnimPreview) // приготавливаемся к треку
				{
					oThis.MouseDownTrack.Start(pos.Page, global_mouseEvent.X, global_mouseEvent.Y);
				}
			}

			if (oThis.m_arrPages[pos.Page].IsSelected)
			{
				let isStartedAnimPreview = (oThis.m_oWordControl.m_oLogicDocument.IsStartedPreview() || (oThis.m_oWordControl.m_oDrawingDocument.TransitionSlide && oThis.m_oWordControl.m_oDrawingDocument.TransitionSlide.IsPlaying()));

				oThis.SelectPageEnabled = false;
				oThis.m_oWordControl.GoToPage(pos.Page);
				oThis.SelectPageEnabled = true;
				
				if (!isStartedAnimPreview)
				{
					if (isMouseDownOnAnimPreview) 
					{
						oThis.m_oWordControl.m_oApi.SlideTransitionPlay(function(){ oThis.m_oWordControl.m_oApi.asc_StartAnimationPreview() });
					}
				}

				if (oThis.m_oWordControl.m_oNotesApi.IsEmptyDraw)
				{
					oThis.m_oWordControl.m_oNotesApi.IsEmptyDraw = false;
					oThis.m_oWordControl.m_oNotesApi.IsRepaint = true;
				}

				if (global_mouseEvent.Button == 2 && !global_keyboardEvent.CtrlKey)
				{
					oThis.showContextMenu(false);
				}
				checkSelectionEnd();
				return false;
			}

			var pages_count = oThis.m_arrPages.length;
			for (var i = 0; i < pages_count; i++)
			{
				oThis.m_arrPages[i].IsSelected = false;
			}

			oThis.m_arrPages[pos.Page].IsSelected = true;

			oThis.OnUpdateOverlay();

			if (global_mouseEvent.Button == 0 && oThis.m_arrPages[pos.Page].animateLabelRect) // click on the current star, preview animation button, slide transition
			{
				let animateLabelRect = oThis.m_arrPages[pos.Page].animateLabelRect;
				let isMouseDownOnAnimPreview = false;
				if (pos.X >= animateLabelRect.minX && pos.X <= animateLabelRect.maxX && pos.Y >= animateLabelRect.minY && pos.Y <= animateLabelRect.maxY);
					isMouseDownOnAnimPreview = true;
			}

			oThis.SelectPageEnabled = false;
			oThis.m_oWordControl.GoToPage(pos.Page);
			oThis.SelectPageEnabled = true;

			if (isMouseDownOnAnimPreview) 
			{
				oThis.m_oWordControl.m_oApi.SlideTransitionPlay(function(){ oThis.m_oWordControl.m_oApi.asc_StartAnimationPreview() });
			}
			
			oThis.ShowPage(pos.Page);
		}

		if (global_mouseEvent.Button == 2 && !global_keyboardEvent.CtrlKey)
		{
			oThis.showContextMenu(false);
		}
		checkSelectionEnd();
		return false;
	};

	this.onMouseMove = function(e)
	{
		let mobileTouchManager = oThis.m_oWordControl ? oThis.m_oWordControl.MobileTouchManagerThumbnails : null;
		if (mobileTouchManager && mobileTouchManager.checkTouchEvent(e))
		{
			mobileTouchManager.startTouchingInProcess();
			let res = mobileTouchManager.mainOnTouchMove(e);
			mobileTouchManager.stopTouchingInProcess();
			return res;
		}

		if (oThis.m_oWordControl)
			oThis.m_oWordControl.m_oApi.checkLastWork();

		var control = oThis.m_oWordControl.m_oThumbnails.HtmlElement;
		if (global_mouseEvent.IsLocked == true && global_mouseEvent.Sender != control)
		{
			// кто-то зажал мышку. кто-то другой
			return;
		}
		AscCommon.check_MouseMoveEvent(e);
		if (global_mouseEvent.Sender != control)
		{
			// такого быть не должно
			return;
		}

		if (oThis.MouseDownTrack.IsStarted())
		{
			// это трек для перекидывания слайдов
			if (oThis.MouseDownTrack.IsSimple() && !oThis.m_oWordControl.m_oApi.isViewMode)
			{
				if (Math.abs(oThis.MouseDownTrack.GetX() - global_mouseEvent.X) > 10 || Math.abs(oThis.MouseDownTrack.GetY() - global_mouseEvent.Y) > 10)
					oThis.MouseDownTrack.ResetSimple(oThis.ConvertCoords2(global_mouseEvent.X, global_mouseEvent.Y));
			}
			else
			{
				if (!oThis.MouseDownTrack.IsSimple())
				{
					// нужно определить активная позиция между слайдами
					oThis.MouseDownTrack.SetPosition(oThis.ConvertCoords2(global_mouseEvent.X, global_mouseEvent.Y));
				}
			}


			oThis.OnUpdateOverlay();

			// теперь нужно посмотреть, нужно ли проскроллить
			if (oThis.m_bIsScrollVisible)
			{
				var _Y = global_mouseEvent.Y - oThis.m_oWordControl.Y;
				var _abs_pos = oThis.m_oWordControl.m_oThumbnails.AbsolutePosition;
				var _YMax = (_abs_pos.B - _abs_pos.T) * g_dKoef_mm_to_pix;

				var _check_type = -1;
				if (/*oThis.MouseDownTrack.GetPosition() != -1 && _Y >= 0 && */_Y < 30)
				{
					_check_type = 0;
				} else if (/*oThis.MouseDownTrack.GetPosition() != -1 &&*/_Y >= (_YMax - 30)/* && _Y < _YMax*/)
				{
					_check_type = 1;
				}

				oThis.CheckNeedAnimateScrolls(_check_type);
			}

			if (!oThis.MouseDownTrack.IsSimple())
			{
				var cursor_dragged = "default";
				if (AscCommon.AscBrowser.isWebkit)
					cursor_dragged = "-webkit-grabbing";
				else if (AscCommon.AscBrowser.isMozilla)
					cursor_dragged = "-moz-grabbing";

				oThis.m_oWordControl.m_oThumbnails.HtmlElement.style.cursor = cursor_dragged;
			}

			return;
		}

		var pos = oThis.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);

		var _is_old_focused = false;

		var pages_count = oThis.m_arrPages.length;
		for (var i = 0; i < pages_count; i++)
		{
			if (oThis.m_arrPages[i].IsFocused)
			{
				_is_old_focused = true;
				oThis.m_arrPages[i].IsFocused = false;
			}
		}

		var cursor_moved = "default";

		if (pos.Page != -1)
		{
			oThis.m_arrPages[pos.Page].IsFocused = true;
			oThis.OnUpdateOverlay();

			cursor_moved = "pointer";
		} else if (_is_old_focused)
		{
			oThis.OnUpdateOverlay();
		}

		oThis.m_oWordControl.m_oThumbnails.HtmlElement.style.cursor = cursor_moved;
	};

	this.onMouseUp = function(e, bIsWindow)
	{
		let mobileTouchManager = oThis.m_oWordControl ? oThis.m_oWordControl.MobileTouchManagerThumbnails : null;
		if (mobileTouchManager && mobileTouchManager.checkTouchEvent(e))
		{
			mobileTouchManager.startTouchingInProcess();
			let res = mobileTouchManager.mainOnTouchEnd(e);
			mobileTouchManager.stopTouchingInProcess();
			return res;
		}

		if (oThis.m_oWordControl)
			oThis.m_oWordControl.m_oApi.checkLastWork();

		var _oldSender = global_mouseEvent.Sender;
		AscCommon.check_MouseUpEvent(e);
		global_mouseEvent.UnLockMouse();

		var control = oThis.m_oWordControl.m_oThumbnails.HtmlElement;
		if (global_mouseEvent.Sender != control)
		{
			if (_oldSender != control || true !== bIsWindow)
			{
				// такого быть не должно
				return;
			}
		}

		oThis.CheckNeedAnimateScrolls(-1);

		if (!oThis.MouseDownTrack.IsStarted())
			return;

		// теперь смотрим, просто ли это селект, или же это трек
		if (oThis.MouseDownTrack.IsSimple())
		{
			if (oThis.MouseDownTrack.IsMoved(global_mouseEvent.X, global_mouseEvent.Y))
				oThis.MouseDownTrack.ResetSimple(oThis.ConvertCoords2(global_mouseEvent.X, global_mouseEvent.Y));
		}

		if (oThis.MouseDownTrack.IsSimple())
		{
			// это просто селект
			var pages_count = oThis.m_arrPages.length;
			for (var i = 0; i < pages_count; i++)
			{
				oThis.m_arrPages[i].IsSelected = false;
			}

			oThis.m_arrPages[oThis.MouseDownTrack.GetPage()].IsSelected = true;

			oThis.OnUpdateOverlay();

			// послали уже на mouseDown
			//oThis.SelectPageEnabled = false;
			//oThis.m_oWordControl.GoToPage(oThis.MouseDownTrack.GetPage());
			//oThis.SelectPageEnabled = true;
		} else
		{
			// это трек
			oThis.MouseDownTrack.SetPosition(oThis.ConvertCoords2(global_mouseEvent.X, global_mouseEvent.Y));

			if (-1 !== oThis.MouseDownTrack.GetPosition() && (!oThis.MouseDownTrack.IsSamePos() || AscCommon.global_mouseEvent.CtrlKey))
			{
				// вызвать функцию апи для смены слайдов местами
				var _array = oThis.GetSelectedArray();
				oThis.m_oWordControl.m_oLogicDocument.shiftSlides(oThis.MouseDownTrack.GetPosition(), _array);
				oThis.ClearCacheAttack();
			}

			oThis.OnUpdateOverlay();
		}
		oThis.MouseDownTrack.Reset();

		oThis.onMouseMove(e);
	};

	this.onMouseLeave = function(e)
	{
		var pages_count = oThis.m_arrPages.length;
		for (var i = 0; i < pages_count; i++)
		{
			oThis.m_arrPages[i].IsFocused = false;
		}
		oThis.OnUpdateOverlay();
	};

	this.onMouseWhell = function(e)
	{
		if (false === this.m_bIsScrollVisible || !oThis.m_oWordControl.m_oScrollThumbApi)
			return;

		if (global_keyboardEvent.CtrlKey)
			return;

		if (undefined !== window["AscDesktopEditor"])
		{
			if (false === window["AscDesktopEditor"]["CheckNeedWheel"]())
				return;
		}

		var delta = GetWheelDeltaY(e);

		oThis.m_oWordControl.m_oScrollThumbApi.scrollBy(0, delta, false);

		if (e.preventDefault)
			e.preventDefault();
		else
			e.returnValue = false;

		AscCommon.stopEvent(e);
		return false;
	};

	this.onKeyDown = function(e)
	{
		var control = oThis.m_oWordControl.m_oThumbnails.HtmlElement;
		if (global_mouseEvent.IsLocked == true && global_mouseEvent.Sender == control)
		{
			e.preventDefault();
			return false;
		}
		AscCommon.check_KeyboardEvent(e);
		let oApi = this.m_oWordControl.m_oApi;
		let oPresentation = this.m_oWordControl.m_oLogicDocument;
		let oDrawingDocument = this.m_oWordControl.m_oDrawingDocument;
		let oEvent = global_keyboardEvent;
		let nShortCutAction = oApi.getShortcut(oEvent);
		let bReturnValue = false, bPreventDefault = true;
		let sSelectedIdx;
		let nStartHistoryIndex = oPresentation.History.Index;
		switch (nShortCutAction)
		{
			case Asc.c_oAscPresentationShortcutType.EditSelectAll:
			{
				this.SelectAll();
				bReturnValue = false;
				bPreventDefault = true;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.EditUndo:
			case Asc.c_oAscPresentationShortcutType.EditRedo:
			{
				bReturnValue = true;
				bPreventDefault = false;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.Cut:
			case Asc.c_oAscPresentationShortcutType.Copy:
			case Asc.c_oAscPresentationShortcutType.Paste:
			{
				bReturnValue = undefined;
				bPreventDefault = false;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.Duplicate:
			{
				if (oPresentation.CanEdit())
				{
					oApi.DublicateSlide();
				}
				bReturnValue = false;
				bPreventDefault = true;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.Print:
			{
				oApi.onPrint();
				bReturnValue = false;
				bPreventDefault = true;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.Save:
			{
				if (!oPresentation.IsViewMode())
				{
					oApi.asc_Save();
				}
				bReturnValue = false;
				bPreventDefault = true;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.ShowContextMenu:
			{
				oThis.showContextMenu(true);
				bReturnValue = false;
				bPreventDefault = false;
				break;
			}
			default:
			{
				break;
			}
		}
		if (!nShortCutAction)
		{
			switch (oEvent.KeyCode)
			{
				case 13: // enter
				{
					if (oPresentation.CanEdit())
					{
						sSelectedIdx = this.GetSelectedArray();
						bReturnValue = false;
						bPreventDefault = true;
						if (sSelectedIdx.length > 0)
						{
							this.m_oWordControl.GoToPage(sSelectedIdx[sSelectedIdx.length - 1]);
							oPresentation.addNextSlide();
							bPreventDefault = false;
						}
					}
					break;
				}
				case 46: // delete
				case 8: // backspace
				{
					if (oPresentation.CanEdit())
					{
						sSelectedIdx = this.GetSelectedArray();
						if (!oApi.IsSupportEmptyPresentation)
						{
							if (sSelectedIdx.length === oDrawingDocument.GetSlidesCount())
							{
								sSelectedIdx.splice(0, 1);
							}
						}
						if (sSelectedIdx.length !== 0)
						{
							oPresentation.deleteSlides(sSelectedIdx);
						}
						if (0 === oPresentation.GetSlidesCount())
						{
							this.m_bIsUpdate = true;
						}
						bReturnValue = false;
						bPreventDefault = true;
					}
					break;
				}
				case 34: // PgDown
				case 40: // bottom arrow
				{
					if (oEvent.CtrlKey && oEvent.ShiftKey)
					{
						oPresentation.moveSelectedSlidesToEnd();
					} else if (oEvent.CtrlKey)
					{
						oPresentation.moveSlidesNextPos();
					} else if (oEvent.ShiftKey)
					{
						this.CorrectShiftSelect(false, false);
					} else
					{
						if (oDrawingDocument.SlideCurrent < oDrawingDocument.GetSlidesCount() - 1)
						{
							this.m_oWordControl.GoToPage(oDrawingDocument.SlideCurrent + 1);
						}
					}
					break;
				}
				case 36: // home
				{
					if (!oEvent.ShiftKey)
					{
						if (oDrawingDocument.SlideCurrent > 0)
						{
							this.m_oWordControl.GoToPage(0);
						}
					} else
					{
						if (oDrawingDocument.SlideCurrent > 0)
						{
							this.CorrectShiftSelect(true, true);
						}
					}
					break;
				}
				case 35: // end
				{
					var slidesCount = oDrawingDocument.GetSlidesCount();
					if (!oEvent.ShiftKey)
					{
						if (oDrawingDocument.SlideCurrent !== (slidesCount - 1))
						{
							this.m_oWordControl.GoToPage(slidesCount - 1);
						}
					} else
					{
						if (oDrawingDocument.SlideCurrent !== (slidesCount - 1))
						{
							this.CorrectShiftSelect(false, true);
						}
					}
					break;
				}
				case 33: // PgUp
				case 38: // UpArrow
				{
					if (oEvent.CtrlKey && oEvent.ShiftKey)
					{
						oPresentation.moveSelectedSlidesToStart();
					} else if (oEvent.CtrlKey)
					{
						oPresentation.moveSlidesPrevPos();
					} else if (oEvent.ShiftKey)
					{
						this.CorrectShiftSelect(true, false);
					} else
					{
						if (oDrawingDocument.SlideCurrent > 0)
						{
							this.m_oWordControl.GoToPage(oDrawingDocument.SlideCurrent - 1);
						}
					}
					break;
				}
				case 77:
				{
					if (oEvent.CtrlKey)
					{
						if (oPresentation.CanEdit())
						{
							sSelectedIdx = this.GetSelectedArray();
							if (sSelectedIdx.length > 0)
							{
								this.m_oWordControl.GoToPage(sSelectedIdx[sSelectedIdx.length - 1]);
								oPresentation.addNextSlide();
								bReturnValue = false;
							} else if (this.GetSlidesCount() === 0)
							{
								oPresentation.addNextSlide();
								this.m_oWordControl.GoToPage(0);
							}
						}
					}
					bReturnValue = false;
					bPreventDefault = true;
					break;
				}
				case 122:
				case 123:
				{
					return;
				}
				default:
					break;
			}
		}
		if (bPreventDefault)
		{
			e.preventDefault();
		}
		if(nStartHistoryIndex === oPresentation.History.Index)
		{
			oApi.sendEvent("asc_onSelectionEnd");
		}
		return bReturnValue;
	};

	this.SetFocusElement = function(type)
	{
		if(this.FocusObjType === type)
			return;
		switch (type)
		{
			case FOCUS_OBJECT_MAIN:
			{
				this.FocusObjType = FOCUS_OBJECT_MAIN;
				break;
			}
			case FOCUS_OBJECT_THUMBNAILS:
			{
				if (this.LockMainObjType)
					return;

				this.FocusObjType = FOCUS_OBJECT_THUMBNAILS;
				if (this.m_oWordControl.m_oLogicDocument)
				{
					this.m_oWordControl.m_oLogicDocument.resetStateCurSlide(true);
				}
				break;
			}
			case FOCUS_OBJECT_ANIM_PANE:
			{
				this.FocusObjType = FOCUS_OBJECT_ANIM_PANE;
				break;
			}
			case FOCUS_OBJECT_NOTES:
			{
				break;
			}
			default:
				break;
		}
	};

	// draw
	this.ShowPage = function(pageNum)
	{
		var y1 = this.m_arrPages[pageNum].top - this.const_border_w;
		var y2 = this.m_arrPages[pageNum].bottom + this.const_border_w;

		if (!this.m_oWordControl.m_oScrollThumbApi)
			return;

		if (y1 < 0)
		{
			var _sizeH = y2 - y1;
			this.m_oWordControl.m_oScrollThumbApi.scrollToY(pageNum * _sizeH + (pageNum + 1) * this.const_border_w);
		} else if (y2 > this.m_oWordControl.m_oThumbnails.HtmlElement.height)
		{
			this.m_oWordControl.m_oScrollThumbApi.scrollByY(y2 - this.m_oWordControl.m_oThumbnails.HtmlElement.height);
		}
	};

	this.ClearCacheAttack = function()
	{
		var pages_count = this.m_arrPages.length;

		for (var i = 0; i < pages_count; i++)
		{
			this.m_arrPages[i].IsRecalc = true;
		}

		this.m_bIsUpdate = true;
	};

	this.FocusRectDraw = function(ctx, x, y, r, b)
	{
		ctx.rect(x - this.const_border_w, y, r - x + this.const_border_w, b - y);
	};
	this.FocusRectFlat = function(_color, ctx, x, y, r, b, isFocus)
	{
		ctx.beginPath();
		ctx.strokeStyle = _color;
		var lineW = Math.round((isFocus ? 2 : 3) * AscCommon.AscBrowser.retinaPixelRatio);
		var dist = Math.round(2 * AscCommon.AscBrowser.retinaPixelRatio);
		ctx.lineWidth = lineW;
		var extend = dist + (lineW / 2);

		ctx.rect(x - extend, y - extend, r - x + (2 * extend), b - y + (2 * extend));
		ctx.stroke();

		ctx.beginPath();
	};

	this.DrawAnimLabel = function(oGraphics, nX, nY, oColor)
	{
		let fCX = function(nVal)
		{
			return AscCommon.AscBrowser.convertToRetinaValue(nVal, true) + nX;
		};
		let fCY = function(nVal)
		{
			return AscCommon.AscBrowser.convertToRetinaValue(nVal, true) + nY;
		};
		oGraphics.b_color1(oColor.R, oColor.G, oColor.B, 255);
		oGraphics.SaveGrState();
		oGraphics.SetIntegerGrid(true);
		let oCtx = oGraphics.m_oContext;
		oCtx.beginPath();
		oCtx.moveTo(fCX(10.5), fCY(4));
		oCtx.lineTo(fCX(12), fCY(8));
		oCtx.lineTo(fCX(16), fCY(8));
		oCtx.lineTo(fCX(12.5), fCY(10.5));
		oCtx.lineTo(fCX(14), fCY(15));
		oCtx.lineTo(fCX(10.5), fCY(12.5));
		oCtx.lineTo(fCX(7), fCY(15));
		oCtx.lineTo(fCX(8.5), fCY(10.5));
		oCtx.lineTo(fCX(5), fCY(8));
		oCtx.lineTo(fCX(9), fCY(8));
		oCtx.lineTo(fCX(10.5), fCY(4));
		oCtx.closePath();
		oCtx.fill();
		
		oCtx.beginPath();
		oCtx.moveTo(fCX(6), fCY(5))
		oCtx.lineTo(fCX(9), fCY(5));
		oCtx.lineTo(fCX(9), fCY(4))
		oCtx.lineTo(fCX(6), fCY(4));
		oCtx.lineTo(fCX(6), fCY(5))
		oCtx.closePath();
		oCtx.fill();

		oCtx.beginPath();
		oCtx.moveTo(fCX(4), fCY(7));
		oCtx.lineTo(fCX(8), fCY(7));
		oCtx.lineTo(fCX(8), fCY(6));
		oCtx.lineTo(fCX(4), fCY(6));
		oCtx.lineTo(fCX(4), fCY(7));
		oCtx.closePath();
		oCtx.fill();
		oCtx.beginPath();
		oGraphics.RestoreGrState();
		return {maxX: fCX(16), minX: fCX(4), maxY: fCY(15), minY: fCY(4)}
	};

	this.OnPaint = function()
	{
		if (!this.m_bIsVisible)
			return;

		var word_control = this.m_oWordControl;
		var canvas = word_control.m_oThumbnails.HtmlElement;

		if (null == canvas)
			return;

		this.m_oWordControl.m_oApi.clearEyedropperImgData();

		var context = canvas.getContext("2d");
		var _width = canvas.width;
		var _height = canvas.height;

		context.clearRect(0, 0, _width, _height);

		var _digit_distance = this.const_offset_x * g_dKoef_pix_to_mm;

		var _logicDocument = word_control.m_oLogicDocument;
		for (var i = 0; i < this.GetSlidesCount(); i++)
		{
			var page = this.m_arrPages[i];

			if (i < this.m_lDrawingFirst || i > this.m_lDrawingEnd)
			{
				this.m_oCacheManager.UnLock(page.cachedImage);
				page.cachedImage = null;
				continue;
			}

			// создаем отрисовщик
			var g = new AscCommon.CGraphics();
			g.init(context, _width, _height, _width * g_dKoef_pix_to_mm, _height * g_dKoef_pix_to_mm);
			g.m_oFontManager = this.m_oFontManager;
			g.transform(1, 0, 0, 1, 0, 0);

			var font = {
				FontFamily: {Name: "Arial", Index: -1},
				Italic: false,
				Bold: false,
				FontSize: Math.round(10 * AscCommon.AscBrowser.retinaPixelRatio)
			};
			g.SetFont(font);



			page.Draw(context, page.left, page.top, page.right - page.left, page.bottom - page.top);

			// меряем надпись номера слайда
			let nSlideNumber = _logicDocument.GetSlideNumber(i);
			if(nSlideNumber !== null)
			{

				var DrawNumSlide = nSlideNumber;
				var num_slide_text_width = 0;
				while (DrawNumSlide !== 0)
				{
					var _last_dig = DrawNumSlide % 10;
					num_slide_text_width += this.DigitWidths[_last_dig];
					DrawNumSlide = (DrawNumSlide / 10) >> 0;
				}
				var text_color = null;
				let sTextColor = AscCommon.GlobalSkin.ThumbnailsPageNumberText;
				if (!page.IsLocked)
					sTextColor = AscCommon.GlobalSkin.ThumbnailsPageNumberText;
				else
					sTextColor = AscCommon.GlobalSkin.ThumbnailsLockColor;

				text_color = AscCommon.RgbaHexToRGBA(sTextColor);
				g.b_color1(text_color.R, text_color.G, text_color.B, 255);

				let dX = (_digit_distance - num_slide_text_width) / 2;
				let dY = page.top * g_dKoef_pix_to_mm + 3 * AscCommon.AscBrowser.retinaPixelRatio;
				let _bounds = g.t("" + nSlideNumber, dX, dY, true);
				let oSlide = _logicDocument.GetSlide(i);
				if (oSlide && !oSlide.isVisible())
				{
					context.lineWidth =  AscCommon.AscBrowser.convertToRetinaValue(1, true);
					context.strokeStyle = sTextColor;
					context.beginPath();
					context.moveTo(_bounds.x - 3, _bounds.y);
					context.lineTo(_bounds.r + 3, _bounds.b);
					context.stroke();
					context.beginPath();
					context.fillStyle = AscCommon.GlobalSkin.BackgroundColorThumbnails;
					context.globalAlpha = 0.5;
					context.fillRect(page.left, page.top, page.right - page.left, page.bottom - page.top);
					context.globalAlpha = 1;
				}
				page.animateLabelRect = null;
				if (_logicDocument.isSlideAnimated(i))
				{
					let nX = (_bounds.x + _bounds.r) / 2 - AscCommon.AscBrowser.convertToRetinaValue(9.5, true);
					let nY = _bounds.b + 3;
					let nIconH = AscCommon.AscBrowser.convertToRetinaValue(15, true);
					if(nY + nIconH < page.bottom)
					{
						let oColor = text_color;
						let resCords = this.DrawAnimLabel(g, nX, nY, oColor);
						page.animateLabelRect = resCords
					}
				}
			}
		}

		this.OnUpdateOverlay();
	};

	this.onCheckUpdate = function()
	{
		if (!this.m_bIsVisible || 0 == this.DigitWidths.length)
			return;

		if (this.m_oWordControl.m_oApi.isSaveFonts_Images)
			return;

		if (this.m_lDrawingFirst == -1 || this.m_lDrawingEnd == -1)
		{
			if (this.m_oWordControl.m_oDrawingDocument.IsEmptyPresentation)
			{
				if (!this.bIsEmptyDrawed)
				{
					this.bIsEmptyDrawed = true;
					this.OnPaint();
				}
				return;
			}

			this.bIsEmptyDrawed = false;
			return;
		}

		this.bIsEmptyDrawed = false;

		if (!this.m_bIsUpdate)
		{
			// определяем, нужно ли пересчитать и закэшировать табнейл (хотя бы один)
			for (var i = this.m_lDrawingFirst; i <= this.m_lDrawingEnd; i++)
			{
				var page = this.m_arrPages[i];
				if (null == page.cachedImage || page.IsRecalc)
				{
					this.m_bIsUpdate = true;
					break;
				}
				if ((page.cachedImage.image.width != (page.right - page.left)) || (page.cachedImage.image.height != (page.bottom - page.top)))
				{
					this.m_bIsUpdate = true;
					break;
				}
			}
		}

		if (!this.m_bIsUpdate)
			return;

		for (var i = this.m_lDrawingFirst; i <= this.m_lDrawingEnd; i++)
		{
			var page = this.m_arrPages[i];
			var w = page.right - page.left;
			var h = page.bottom - page.top;

			if (null != page.cachedImage)
			{
				if ((page.cachedImage.image.width != w) || (page.cachedImage.image.height != h) || page.IsRecalc)
				{
					this.m_oCacheManager.UnLock(page.cachedImage);
					page.cachedImage = null;
				}
			}

			if (null == page.cachedImage)
			{
				page.cachedImage = this.m_oCacheManager.Lock(w, h);

				var g = new AscCommon.CGraphics();
				g.IsNoDrawingEmptyPlaceholder = true;
				g.IsThumbnail = true;
				if (this.m_oWordControl.m_oLogicDocument.isVisioDocument)
				{
					//todo override CThumbnailsManager
					let SlideWidth = this.m_oWordControl.m_oLogicDocument.GetWidthMM(i);
					let SlideHeight = this.m_oWordControl.m_oLogicDocument.GetHeightMM(i);
					g.init(page.cachedImage.image.ctx, w, h, SlideWidth, SlideHeight);
				}
				else
				{
					g.init(page.cachedImage.image.ctx, w, h, this.SlideWidth, this.SlideHeight);
				}
				g.m_oFontManager = this.m_oFontManager;

				g.transform(1, 0, 0, 1, 0, 0);

				var bIsShowPars = this.m_oWordControl.m_oApi.ShowParaMarks;
				this.m_oWordControl.m_oApi.ShowParaMarks = false;
				this.m_oWordControl.m_oLogicDocument.DrawPage(i, g);
				this.m_oWordControl.m_oApi.ShowParaMarks = bIsShowPars;

				page.IsRecalc = false;

				this.m_bIsUpdate = true;
				break;
			}
		}

		this.OnPaint();
		this.m_bIsUpdate = false;
	};

	// overlay
	this.OnUpdateOverlay = function()
	{
		if (!this.m_bIsVisible)
			return;

		var canvas = this.m_oWordControl.m_oThumbnailsBack.HtmlElement;
		if (null == canvas)
			return;

		if (this.m_oWordControl)
			this.m_oWordControl.m_oApi.checkLastWork();

		var context = canvas.getContext("2d");
		var _width = canvas.width;
		var _height = canvas.height;

		context.fillStyle = GlobalSkin.BackgroundColorThumbnails;
		context.fillRect(0, 0, _width, _height);

		var _style_select = GlobalSkin.ThumbnailsPageOutlineActive;
		var _style_focus = GlobalSkin.ThumbnailsPageOutlineHover;
		var _style_select_focus = GlobalSkin.ThumbnailsPageOutlineActive;

		// selected pages
		context.fillStyle = _style_select;
		var _border = this.const_border_w;
		for (var i = 0; i < this.GetSlidesCount(); i++)
		{
			var page = this.m_arrPages[i];

			if (page.IsLocked)
			{
				var _lock_focus = AscCommon.GlobalSkin.ThumbnailsLockColor;
				var _lock_color = AscCommon.GlobalSkin.ThumbnailsLockColor;

				if (page.IsFocused)
				{
					this.FocusRectFlat(_lock_focus, context, page.left, page.top, page.right, page.bottom);
				} else
				{
					this.FocusRectFlat(_lock_color, context, page.left, page.top, page.right, page.bottom);
				}

				continue;
			}

			if (page.IsSelected && page.IsFocused)
			{
				this.FocusRectFlat(_style_select_focus, context, page.left, page.top, page.right, page.bottom);
			} else if (page.IsSelected)
			{
				this.FocusRectFlat(_style_select, context, page.left, page.top, page.right, page.bottom);
			} else if (page.IsFocused)
			{
				this.FocusRectFlat(_style_focus, context, page.left, page.top, page.right, page.bottom, true);
			}
		}

		if (this.MouseDownTrack.IsDragged())
		{
			// теперь нужно просто нарисовать линию
			context.strokeStyle = "#DEDEDE";
			let y = (0.5 * this.const_offset_y) >> 0;
			let nPosition = this.MouseDownTrack.GetPosition();
			let oPage = this.m_arrPages[nPosition - 1];
			if (oPage)
				y = (oPage.bottom + 1.5 * this.const_border_w) >> 0;

			var _left_pos = 0;
			var _right_pos = _width;
			if (this.m_arrPages.length > 0)
			{
				_left_pos = this.m_arrPages[0].left + 4;
				_right_pos = this.m_arrPages[0].right - 4;
			}

			context.lineWidth = 3;
			context.beginPath();
			context.moveTo(_left_pos, y + 0.5);
			context.lineTo(_right_pos, y + 0.5);
			context.stroke();
			context.beginPath();

			if (null != this.MouseTrackCommonImage)
			{
				context.drawImage(this.MouseTrackCommonImage, 0, 0,
					(this.MouseTrackCommonImage.width + 1) >> 1, this.MouseTrackCommonImage.height,
					_left_pos - this.MouseTrackCommonImage.width,
					y - (this.MouseTrackCommonImage.height >> 1),
					(this.MouseTrackCommonImage.width + 1) >> 1, this.MouseTrackCommonImage.height);

				context.drawImage(this.MouseTrackCommonImage, this.MouseTrackCommonImage.width >> 1, 0,
					(this.MouseTrackCommonImage.width + 1) >> 1, this.MouseTrackCommonImage.height,
					_right_pos + (this.MouseTrackCommonImage.width >> 1),
					y - (this.MouseTrackCommonImage.height >> 1),
					(this.MouseTrackCommonImage.width + 1) >> 1, this.MouseTrackCommonImage.height);
			}
		}
	};

	// select
	this.SelectSlides = function(aSelectedIdx, bReset)
	{
		if (aSelectedIdx.length > 0)
		{
			if(bReset === undefined || bReset)
			{
				for (let i = 0; i < this.m_arrPages.length; i++)
				{
					this.m_arrPages[i].IsSelected = false;
				}
			}
			let nCount = aSelectedIdx.length;
			let nSlideType = this.GetSlideType(aSelectedIdx[0]);
			for (let nIdx = 0; nIdx < nCount; nIdx++)
			{
				if(this.GetSlideType(aSelectedIdx[nIdx]) === nSlideType)
				{
					let oPage = this.m_arrPages[aSelectedIdx[nIdx]];
					if (oPage)
					{
						oPage.IsSelected = true;
					}
				}
			}
			this.OnUpdateOverlay();
		}
	};

	this.GetSelectedSlidesRange = function()
	{
		var _min = this.m_oWordControl.m_oDrawingDocument.SlideCurrent;
		var _max = _min;
		var pages_count = this.m_arrPages.length;
		for (var i = 0; i < pages_count; i++)
		{
			if (this.m_arrPages[i].IsSelected)
			{
				if (i < _min)
					_min = i;
				if (i > _max)
					_max = i;
			}
		}
		return {Min: _min, Max: _max};
	};

	this.GetSelectedArray = function()
	{
		var _array = [];
		var pages_count = this.m_arrPages.length;
		for (var i = 0; i < pages_count; i++)
		{
			if (this.m_arrPages[i].IsSelected)
			{
				_array[_array.length] = i;
			}
		}
		return _array;
	};

	this.CorrectShiftSelect = function(isTop, isEnd)
	{
		var drDoc = this.m_oWordControl.m_oDrawingDocument;
		var slidesCount = drDoc.GetSlidesCount();
		var min_max = this.GetSelectedSlidesRange();

		var _page = this.m_oWordControl.m_oDrawingDocument.SlideCurrent;
		if (isEnd)
		{
			_page = isTop ? 0 : slidesCount - 1;
		} else if (isTop)
		{
			if (min_max.Min != _page)
			{
				_page = min_max.Min - 1;
				if (_page < 0)
					_page = 0;
			} else
			{
				_page = min_max.Max - 1;
				if (_page < 0)
					_page = 0;
			}
		} else
		{
			if (min_max.Min != _page)
			{
				_page = min_max.Min + 1;
				if (_page >= slidesCount)
					_page = slidesCount - 1;
			} else
			{
				_page = min_max.Max + 1;
				if (_page >= slidesCount)
					_page = slidesCount - 1;
			}
		}

		var _max = _page;
		var _min = this.m_oWordControl.m_oDrawingDocument.SlideCurrent;
		if (_min > _max)
		{
			var _temp = _max;
			_max = _min;
			_min = _temp;
		}

		for (var i = 0; i < _min; i++)
		{
			this.m_arrPages[i].IsSelected = false;
		}
		let nSlideType = this.GetSlideType(_min);
		for (var i = _min; i <= _max; i++)
		{
			if(this.GetSlideType(i) === nSlideType)
			{
				this.m_arrPages[i].IsSelected = true;
			}
		}
		for (var i = _max + 1; i < slidesCount; i++)
		{
			this.m_arrPages[i].IsSelected = false;
		}


		this.m_oWordControl.m_oLogicDocument.Document_UpdateInterfaceState();
		this.OnUpdateOverlay();
		this.ShowPage(_page);
	};

	this.SelectAll = function()
	{
		for (var i = 0; i < this.m_arrPages.length; i++)
		{
			this.m_arrPages[i].IsSelected = true;
		}
		this.UpdateInterface();
		this.OnUpdateOverlay();
	};

	this.IsSlideHidden = function(aSelected)
	{
		var oPresentation = oThis.m_oWordControl.m_oLogicDocument;
		for (var i = 0; i < aSelected.length; ++i)
		{
			if (oPresentation.IsVisibleSlide(aSelected[i]))
				return false;
		}
		return true;
	};

	this.isSelectedPage = function(pageNum)
	{
		if (this.m_arrPages[pageNum] && this.m_arrPages[pageNum].IsSelected)
			return true;
		return false;
	};


	this.SelectPage = function(pageNum)
	{
		if (!this.SelectPageEnabled)
			return;

		var pages_count = this.m_arrPages.length;

		if (pageNum >= 0 && pageNum < pages_count)
		{
			var bIsUpdate = false;
			for (var i = 0; i < pages_count; i++)
			{
				if (this.m_arrPages[i].IsSelected === true && i != pageNum)
					bIsUpdate = true;
				this.m_arrPages[i].IsSelected = false;
			}

			if (this.m_arrPages[pageNum].IsSelected === false)
				bIsUpdate = true;
			this.m_arrPages[pageNum].IsSelected = true;

			this.m_bIsUpdate = bIsUpdate;

			if (bIsUpdate && this.m_oWordControl.m_oScrollThumbApi != null)
			{
				var y1 = this.m_arrPages[pageNum].top - this.const_border_w;
				var y2 = this.m_arrPages[pageNum].bottom + this.const_border_w;

				if (y1 < 0)
				{
					var _sizeH = y2 - y1;
					this.m_oWordControl.m_oScrollThumbApi.scrollToY(pageNum * _sizeH + (pageNum + 1) * this.const_border_w);
				} else if (y2 > this.m_oWordControl.m_oThumbnails.HtmlElement.height)
				{
					this.m_oWordControl.m_oScrollThumbApi.scrollByY(y2 - this.m_oWordControl.m_oThumbnails.HtmlElement.height);
				}
			}
		}
	};

	// position
	this.ConvertCoords = function(x, y)
	{
		var Pos = {X: x, Y: y};
		Pos.X -= this.m_oWordControl.X;
		Pos.Y -= this.m_oWordControl.Y;

		Pos.X = AscCommon.AscBrowser.convertToRetinaValue(Pos.X, true);
		Pos.Y = AscCommon.AscBrowser.convertToRetinaValue(Pos.Y, true);

		Pos
		Pos.Page = -1;
		let pages_count = this.m_arrPages.length;
		for (let i = 0; i < pages_count; i++)
		{
			if(this.m_arrPages[i].Hit(Pos.X, Pos.Y))
			{
				Pos.Page = i;
				break;
			}
		}
		return Pos;
	};

	this.GetContextMenuPos = function()
	{
		let oPos = {};
		oPos.X_abs = global_mouseEvent.X - ((oThis.m_oWordControl.m_oThumbnails.AbsolutePosition.L * g_dKoef_mm_to_pix) >> 0) - oThis.m_oWordControl.X;
		oPos.Y_abs = global_mouseEvent.Y - ((oThis.m_oWordControl.m_oThumbnails.AbsolutePosition.T * g_dKoef_mm_to_pix) >> 0) - oThis.m_oWordControl.Y;
		return oPos;
	};
	this.ConvertCoords2 = function(x, y)
	{
		var Pos = {X: x, Y: y};
		Pos.X -= this.m_oWordControl.X;
		Pos.Y -= this.m_oWordControl.Y;

		Pos.X = AscCommon.AscBrowser.convertToRetinaValue(Pos.X, true);
		Pos.Y = AscCommon.AscBrowser.convertToRetinaValue(Pos.Y, true);

		var _abs_pos = this.m_oWordControl.m_oThumbnails.AbsolutePosition;
		var _controlW = (_abs_pos.R - _abs_pos.L) * g_dKoef_mm_to_pix * AscCommon.AscBrowser.retinaPixelRatio;
		var _controlH = (_abs_pos.B - _abs_pos.T) * g_dKoef_mm_to_pix * AscCommon.AscBrowser.retinaPixelRatio;

		if (Pos.X < 0 || Pos.X > _controlW || Pos.Y < 0 || Pos.Y > _controlH)
			return -1;

		var pages_count = this.m_arrPages.length;
		if (0 == pages_count)
			return -1;

		var _min = Math.abs(Pos.Y - this.m_arrPages[0].top);
		var _MinPositionPage = 0;

		for (var i = 0; i < pages_count; i++)
		{
			var _min1 = Math.abs(Pos.Y - this.m_arrPages[i].top);
			var _min2 = Math.abs(Pos.Y - this.m_arrPages[i].bottom);

			if (_min1 < _min)
			{
				_min = _min1;
				_MinPositionPage = i;
			}
			if (_min2 < _min)
			{
				_min = _min2;
				_MinPositionPage = i + 1;
			}
		}

		return _MinPositionPage;
	};

	this.CalculatePlaces = function()
	{
		if (!this.m_bIsVisible)
			return;

		var word_control = this.m_oWordControl;

		var dKoefToPix = AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_mm_to_pix;

		if (word_control && word_control.MobileTouchManagerThumbnails)
			word_control.MobileTouchManagerThumbnails.ClearContextMenu();

		var canvas = word_control.m_oThumbnails.HtmlElement;
		if (null == canvas)
			return;

		var _width = canvas.width;
		var _height = canvas.height;

		var bIsFoundFirst = false;
		var bIsFoundEnd = false;

		var lCurrentTopInDoc = (this.m_dScrollY) >> 0;

		var __w = word_control.m_oThumbnails.AbsolutePosition.R - word_control.m_oThumbnails.AbsolutePosition.L;
		var __h = word_control.m_oThumbnails.AbsolutePosition.B - word_control.m_oThumbnails.AbsolutePosition.T;
		var nWidthSlide = (__w * dKoefToPix) >> 0;

		nWidthSlide -= (this.const_offset_x + this.const_offset_r);
		var nHeightSlide = (nWidthSlide * this.SlideHeight / this.SlideWidth) >> 0;

		var lStart = this.const_offset_y;
		let SlidesCount = this.GetSlidesCount();
		for (var i = 0; i < SlidesCount; i++)
		{
			if (this.m_oWordControl.m_oLogicDocument.isVisioDocument)
			{
				//todo override CThumbnailsManager
				let SlideWidth = this.m_oWordControl.m_oLogicDocument.GetWidthMM(i);
				let SlideHeight = this.m_oWordControl.m_oLogicDocument.GetHeightMM(i);
				nHeightSlide = (nWidthSlide * SlideHeight / SlideWidth) >> 0;
			}

			if (i >= this.m_arrPages.length)
			{
				this.m_arrPages[i] = new CThPage();
				if (0 == i)
					this.m_arrPages[0].IsSelected = true;
			}

			if (false === bIsFoundFirst)
			{
				if ((lStart + nHeightSlide) > lCurrentTopInDoc)
				{
					this.m_lDrawingFirst = i;
					bIsFoundFirst = true;
				}
			}

			var drawRect = this.m_arrPages[i];
			let nHeight = nHeightSlide;
			drawRect.left = this.const_offset_x;
			drawRect.top = lStart - lCurrentTopInDoc;
			drawRect.right = drawRect.left + nWidthSlide;
			drawRect.bottom = drawRect.top + nHeight;
			drawRect.pageIndex = i;
			let oSlide = this.m_oWordControl.m_oLogicDocument.GetSlide(i);
			if(oSlide.getObjectType() === AscDFH.historyitem_type_SlideLayout)
			{
				let nWidth = (drawRect.right - drawRect.left) * LAYOUT_SCALE;
				drawRect.left = (drawRect.right - nWidth) + 0.5 >> 0;
				nHeight = nHeightSlide * LAYOUT_SCALE + 0.5 >> 0;
				drawRect.bottom = drawRect.top + nHeight;
			}

			if (false === bIsFoundEnd)
			{
				if (drawRect.top > _height)
				{
					this.m_lDrawingEnd = i - 1;
					bIsFoundEnd = true;
				}
			}

			lStart += (nHeight + 3 * this.const_border_w);
		}

		if (this.m_arrPages.length > SlidesCount)
			this.m_arrPages.splice(SlidesCount, this.m_arrPages.length - SlidesCount);

		if (false === bIsFoundEnd)
		{
			this.m_lDrawingEnd = SlidesCount - 1;
		}
	};

	this.GetThumbnailPagePosition = function(pageIndex)
	{
		if (pageIndex < 0 || pageIndex >= this.m_arrPages.length)
			return null;

		var drawRect = this.m_arrPages[pageIndex];
		var _ret = {
			X: AscCommon.AscBrowser.convertToRetinaValue(drawRect.left),
			Y: AscCommon.AscBrowser.convertToRetinaValue(drawRect.top),
			W: AscCommon.AscBrowser.convertToRetinaValue(drawRect.right - drawRect.left),
			H: AscCommon.AscBrowser.convertToRetinaValue(drawRect.bottom - drawRect.top)
		};
		return _ret;
	};

	this.getSpecialPasteButtonCoords = function(sSlideId)
	{
		if(!sSlideId) return null;
		let nSlideIdx = null;
		let oPresentation = this.m_oWordControl.m_oLogicDocument;
		let aSlides = oPresentation.GetAllSlides();
		for (let nSld = 0; nSld < aSlides.length; ++nSld)
		{
			if (aSlides[nSld].Get_Id() === sSlideId)
			{
				nSlideIdx = nSld;
				break;
			}
		}
		if(nSlideIdx === null)
		{
			return null;
		}
		let oRect = this.GetThumbnailPagePosition(nSlideIdx);
		if (!oRect)
		{
			return null;
		}
		let nX, nY;
		let oThContainer = editor.WordControl.m_oThumbnailsContainer;
		if (oThContainer.HtmlElement.style.display === "none")
		{
			nX = 0;
		} else
		{
			nX = oRect.X + oRect.W - AscCommon.specialPasteElemWidth;
		}
		nY = oRect.Y + oRect.H;
		return {X: nX, Y: nY};
	};

	// scroll
	this.OnScrollTrackTop = function()
	{
		oThis.m_oWordControl.m_oScrollThumbApi.scrollByY(-45);
	};
	this.OnScrollTrackBottom = function()
	{
		oThis.m_oWordControl.m_oScrollThumbApi.scrollByY(45);
	};

	this.CheckNeedAnimateScrolls = function(type)
	{
		if (type == -1)
		{
			// нужно застопить все
			if (this.MouseThumbnailsAnimateScrollTopTimer != -1)
			{
				clearInterval(this.MouseThumbnailsAnimateScrollTopTimer);
				this.MouseThumbnailsAnimateScrollTopTimer = -1;
			}

			if (this.MouseThumbnailsAnimateScrollBottomTimer != -1)
			{
				clearInterval(this.MouseThumbnailsAnimateScrollBottomTimer);
				this.MouseThumbnailsAnimateScrollBottomTimer = -1;
			}
		}

		if (type == 0)
		{
			if (this.MouseThumbnailsAnimateScrollBottomTimer != -1)
			{
				clearInterval(this.MouseThumbnailsAnimateScrollBottomTimer);
				this.MouseThumbnailsAnimateScrollBottomTimer = -1;
			}

			if (-1 == this.MouseThumbnailsAnimateScrollTopTimer)
			{
				this.MouseThumbnailsAnimateScrollTopTimer = setInterval(this.OnScrollTrackTop, 50);
			}
			return;
		}
		if (type == 1)
		{
			if (this.MouseThumbnailsAnimateScrollTopTimer != -1)
			{
				clearInterval(this.MouseThumbnailsAnimateScrollTopTimer);
				this.MouseThumbnailsAnimateScrollTopTimer = -1;
			}

			if (-1 == this.MouseThumbnailsAnimateScrollBottomTimer)
			{
				this.MouseThumbnailsAnimateScrollBottomTimer = setInterval(this.OnScrollTrackBottom, 50);
			}
			return;
		}
	};

	this.verticalScroll = function(sender, scrollPositionY, maxY, isAtTop, isAtBottom)
	{
		if (false === this.m_oWordControl.m_oApi.bInit_word_control || false === this.m_bIsScrollVisible)
			return;

		this.m_dScrollY = Math.max(0, Math.min(scrollPositionY, maxY));
		this.m_dScrollY_max = maxY;

		this.CalculatePlaces();
		AscCommon.g_specialPasteHelper.SpecialPasteButton_Update_Position();
		this.m_bIsUpdate = true;

		if (!this.m_oWordControl.m_oApi.isMobileVersion)
			this.SetFocusElement(FOCUS_OBJECT_THUMBNAILS);
	};

	// calculate
	this.RecalculateAll = function()
	{
		this.SlideWidth = this.m_oWordControl.m_oLogicDocument.GetWidthMM();
		this.SlideHeight = this.m_oWordControl.m_oLogicDocument.GetHeightMM();
		this.CheckSizes();

		this.ClearCacheAttack();
	};

	// size
	this.CheckSizes = function()
	{
		this.InitCheckOnResize();
		var word_control = this.m_oWordControl;
		let oPresentation = word_control.m_oLogicDocument;
		var dKoefToPix = AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_mm_to_pix;

		var __w = word_control.m_oThumbnailsContainer.AbsolutePosition.R - word_control.m_oThumbnailsContainer.AbsolutePosition.L;
		var __h = word_control.m_oThumbnailsContainer.AbsolutePosition.B - word_control.m_oThumbnailsContainer.AbsolutePosition.T;
		var nWidthSlide = (__w * dKoefToPix) >> 0;

		if (__w < 1 || __h < 0)
		{
			this.m_bIsVisible = false;
			return;
		}
		this.m_bIsVisible = true;

		nWidthSlide -= this.const_offset_r;

		var _tmpDig = 0;
		if (this.DigitWidths.length > 5)
			_tmpDig = this.DigitWidths[5];


		let SlidesCount = this.GetSlidesCount();
		this.const_offset_x = (_tmpDig * dKoefToPix * (("") + (SlidesCount + oPresentation.getFirstSlideNumber())).length) >> 0;
		if (this.const_offset_x < 25)
			this.const_offset_x = 25;

		// focus/select rects
		this.const_offset_x += Math.round(9 * AscCommon.AscBrowser.retinaPixelRatio);

		nWidthSlide -= this.const_offset_x;

		var nHeightSlide = (nWidthSlide * this.SlideHeight / this.SlideWidth) >> 0;

		let nSumThHeight = 0;
		if (this.m_oWordControl.m_oLogicDocument.isVisioDocument)
		{
			//todo override CThumbnailsManager
			for (let nIdx = 0; nIdx < this.GetSlidesCount(); ++nIdx)
			{
				let SlideWidth = this.m_oWordControl.m_oLogicDocument.GetWidthMM(nIdx);
				let SlideHeight = this.m_oWordControl.m_oLogicDocument.GetHeightMM(nIdx);
				nSumThHeight += (nWidthSlide * SlideHeight / SlideWidth) >> 0;
			}
		}
		else
		{
			if(!this.IsMasterMode())
			{
				nSumThHeight = nHeightSlide * SlidesCount;
			}
			else
			{
				nSumThHeight = 0;
				for(let nIdx = 0; nIdx < oPresentation.slideMasters.length; ++nIdx)
				{
					nSumThHeight += nHeightSlide;
					let oMaster = oPresentation.slideMasters[nIdx];
					nSumThHeight += LAYOUT_SCALE * oMaster.sldLayoutLst.length * nHeightSlide;
				}
				nSumThHeight = nSumThHeight + 0.5 >> 0;
			}
		}
		var nHeightPix = this.const_offset_y + this.const_offset_y + nSumThHeight;
		if (SlidesCount > 0)
			nHeightPix += (SlidesCount - 1) * 3 * this.const_border_w;

		var dPosition = 0;
		if (this.m_dScrollY_max != 0)
		{
			dPosition = this.m_dScrollY / this.m_dScrollY_max;
		}

		var heightThumbs = (__h * dKoefToPix) >> 0;
		if (nHeightPix < heightThumbs)
		{
			// все убралось. скролл не нужен
			if (this.m_bIsScrollVisible && GlobalSkin.ThumbnailScrollWidthNullIfNoScrolling)
			{
				word_control.m_oThumbnails.Bounds.R = 0;
				word_control.m_oThumbnailsBack.Bounds.R = 0;
				word_control.m_oThumbnails_scroll.Bounds.AbsW = 0;

				word_control.m_oThumbnailsContainer.Resize(__w, __h, word_control);
			} else
			{
				word_control.m_oThumbnails_scroll.HtmlElement.style.display = "none";
			}
			this.m_bIsScrollVisible = false;
			this.m_dScrollY = 0;
		} else
		{
			// нужен скролл
			if (!this.m_bIsScrollVisible)
			{
				if (GlobalSkin.ThumbnailScrollWidthNullIfNoScrolling)
				{
					word_control.m_oThumbnailsBack.Bounds.R = word_control.ScrollWidthPx * dKoefToPix;
					word_control.m_oThumbnails.Bounds.R = word_control.ScrollWidthPx * dKoefToPix;

					var _width_mm_scroll = (true) ? 10 : word_control.ScrollWidthPx;
					word_control.m_oThumbnails_scroll.Bounds.AbsW = _width_mm_scroll * dKoefToPix;
				} else
				{
					if (!word_control.m_oApi.isMobileVersion)
						word_control.m_oThumbnails_scroll.HtmlElement.style.display = "block";
				}

				word_control.m_oThumbnailsContainer.Resize(__w, __h, word_control);
			}
			this.m_bIsScrollVisible = true;

			__w = word_control.m_oThumbnails.AbsolutePosition.R - word_control.m_oThumbnails.AbsolutePosition.L;
			__h = word_control.m_oThumbnails.AbsolutePosition.B - word_control.m_oThumbnails.AbsolutePosition.T;
			nWidthSlide = (__w * dKoefToPix) >> 0;

			nWidthSlide -= (this.const_offset_x + this.const_offset_r);

			var nHeightSlide = (nWidthSlide * this.SlideHeight / this.SlideWidth) >> 0;

			let nSumThHeight = 0;
			if (this.m_oWordControl.m_oLogicDocument.isVisioDocument)
			{
				//todo override CThumbnailsManager
				for (let nIdx = 0; nIdx < this.GetSlidesCount(); ++nIdx)
				{
					let SlideWidth = this.m_oWordControl.m_oLogicDocument.GetWidthMM(nIdx);
					let SlideHeight = this.m_oWordControl.m_oLogicDocument.GetHeightMM(nIdx);
					nSumThHeight += (nWidthSlide * SlideHeight / SlideWidth) >> 0;
				}
			}
			else
			{
				if(!this.IsMasterMode())
				{
					nSumThHeight = nHeightSlide * SlidesCount;
				}
				else
				{
					nSumThHeight = 0;
					for(let nIdx = 0; nIdx < oPresentation.slideMasters.length; ++nIdx)
					{
						nSumThHeight += nHeightSlide;
						let oMaster = oPresentation.slideMasters[nIdx];
						nSumThHeight += LAYOUT_SCALE * oMaster.sldLayoutLst.length * nHeightSlide;
					}
					nSumThHeight = nSumThHeight + 0.5 >> 0;
				}
			}
			var nHeightPix = this.const_offset_y + this.const_offset_y + nSumThHeight;
			if (SlidesCount > 0)
				nHeightPix += (SlidesCount - 1) * 3 * this.const_border_w;

			// теперь нужно выставить размеры
			var settings = new AscCommon.ScrollSettings();
			settings.showArrows = false;
			settings.screenW = word_control.m_oThumbnails.HtmlElement.width;
			settings.screenH = word_control.m_oThumbnails.HtmlElement.height;
			settings.cornerRadius = 1;
			settings.slimScroll = true;

			settings.scrollBackgroundColor = GlobalSkin.BackgroundColorThumbnails;
			settings.scrollBackgroundColorHover = GlobalSkin.BackgroundColorThumbnails;
			settings.scrollBackgroundColorActive = GlobalSkin.BackgroundColorThumbnails;

			settings.scrollerColor = GlobalSkin.ScrollerColor;
			settings.scrollerHoverColor = GlobalSkin.ScrollerHoverColor;
			settings.scrollerActiveColor = GlobalSkin.ScrollerActiveColor;

			settings.arrowColor = GlobalSkin.ScrollArrowColor;
			settings.arrowHoverColor = GlobalSkin.ScrollArrowHoverColor;
			settings.arrowActiveColor = GlobalSkin.ScrollArrowActiveColor;

			settings.strokeStyleNone = GlobalSkin.ScrollOutlineColor;
			settings.strokeStyleOver = GlobalSkin.ScrollOutlineHoverColor;
			settings.strokeStyleActive = GlobalSkin.ScrollOutlineActiveColor;

			settings.targetColor = GlobalSkin.ScrollerTargetColor;
			settings.targetHoverColor = GlobalSkin.ScrollerTargetHoverColor;
			settings.targetActiveColor = GlobalSkin.ScrollerTargetActiveColor;

			settings.contentH = nHeightPix;

			if (word_control.m_oScrollThumb_)
			{
				word_control.m_oScrollThumb_.Repos(settings);
				word_control.m_oScrollThumb_.isHorizontalScroll = false;
			} else
			{
				word_control.m_oScrollThumb_ = new AscCommon.ScrollObject("id_vertical_scroll_thmbnl", settings);
				word_control.m_oScrollThumb_.bind("scrollvertical", function(evt)
				{
					oThis.verticalScroll(this, evt.scrollD, evt.maxScrollY);
				});
				word_control.m_oScrollThumbApi = word_control.m_oScrollThumb_;
				word_control.m_oScrollThumb_.isHorizontalScroll = false;
			}
		}

		if (this.m_bIsScrollVisible)
		{
			var lPosition = (dPosition * word_control.m_oScrollThumbApi.getMaxScrolledY()) >> 0;
			word_control.m_oScrollThumbApi.scrollToY(lPosition);
		}

		this.ScrollerHeight = nHeightPix;
		if (word_control.MobileTouchManagerThumbnails)
		{
			word_control.MobileTouchManagerThumbnails.Resize();
		}

		this.CalculatePlaces();
		AscCommon.g_specialPasteHelper.SpecialPasteButton_Update_Position();
		this.m_bIsUpdate = true;
	};

	// interface
	this.UpdateInterface = function()
	{
		this.m_oWordControl.m_oLogicDocument.Document_UpdateInterfaceState();
	};

	this.showContextMenu = function(bPosBySelect)
	{
		let sSelectedIdx = this.GetSelectedArray();
		let oMenuPos;
		let bIsSlideSelect = false;
		if(bPosBySelect)
		{
			oMenuPos = this.GetThumbnailPagePosition(Math.min.apply(Math, sSelectedIdx));
			bIsSlideSelect = sSelectedIdx.length > 0;
		}
		else
		{
			let oEditorCtrl = oThis.m_oWordControl;
			let oThCtrlPos = oEditorCtrl.m_oThumbnails.AbsolutePosition;
			oMenuPos = {
				X: global_mouseEvent.X - ((oThCtrlPos.L * g_dKoef_mm_to_pix) >> 0) - oEditorCtrl.X,
				Y: global_mouseEvent.Y - ((oThCtrlPos.T * g_dKoef_mm_to_pix) >> 0) - oEditorCtrl.Y
			};
			
			let pos = this.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);
			bIsSlideSelect = pos.Page !== -1;
		}
		if (oMenuPos)
		{
			let oFirstSlide = this.m_oWordControl.m_oLogicDocument.GetSlide(sSelectedIdx[0]);
			let nType = Asc.c_oAscContextMenuTypes.Thumbnails;
			if(oFirstSlide)
			{
				if(oFirstSlide.getObjectType() === AscDFH.historyitem_type_SlideLayout)
				{
					nType = Asc.c_oAscContextMenuTypes.Layout;
				}
				else if(oFirstSlide.getObjectType() === AscDFH.historyitem_type_SlideMaster)
				{
					nType = Asc.c_oAscContextMenuTypes.Master;
				}
			}
			let oData =
				{
					Type: nType,
					X_abs: oMenuPos.X,
					Y_abs: oMenuPos.Y,
					IsSlideSelect: bIsSlideSelect,
					IsSlideHidden: this.IsSlideHidden(sSelectedIdx)
				};
			editor.sync_ContextMenuCallback(new AscCommonSlide.CContextMenuData(oData));
		}
	};

	this.GetCurSld = function()
	{
		return this.thumbnails.GetCurSld();
	};
}

function CSlideDrawer()
{
	this.m_oWordControl             = null;
	this.CONST_MAX_SLIDE_CACHE_SIZE = 104857600; // 100 megabytes
	if (AscCommon.AscBrowser.isAppleDevices)
		this.CONST_MAX_SLIDE_CACHE_SIZE = 16777216;

	this.CONST_BORDER               = 10; // in px

	this.IsCached        = false;
	this.CachedCanvas    = null;
	this.CachedCanvasCtx = null;

	this.BoundsChecker = new AscFormat.CSlideBoundsChecker();
	this.BoundsChecker2 = new AscFormat.CSlideBoundsChecker();

	this.CacheSlidePixW = 1;
	this.CacheSlidePixH = 1;

	this.bIsEmptyPresentation = false;
	this.IsRecalculateSlide   = false;

	this.EmptyPresenattionTextHeight = 60;

	// TODO: максимальная ширина всех линий и запас под локи
	this.SlideEps = 20;

	this.CheckRecalculateSlide = function()
	{
		if (this.IsRecalculateSlide)
		{
			this.IsRecalculateSlide = false;
			this.m_oWordControl.m_oDrawingDocument.FirePaint();
		}
	}

	this.CheckSlideSize = function(zoom, slideNum)
	{
		if (-1 == slideNum)
			this.bIsEmptyPresentation = true;

		var dKoef = zoom * g_dKoef_mm_to_pix / 100;
		dKoef *= AscCommon.AscBrowser.retinaPixelRatio;

		var w_mm = this.m_oWordControl.m_oLogicDocument.GetWidthMM();
		var h_mm = this.m_oWordControl.m_oLogicDocument.GetHeightMM();
		var w_px = (w_mm * dKoef) >> 0;
		var h_px = (h_mm * dKoef) >> 0;

		this.BoundsChecker2.init(w_px, h_px, w_mm, h_mm);
		this.BoundsChecker2.transform(1, 0, 0, 1, 0, 0);

		if (this.bIsEmptyPresentation)
		{
			this.BoundsChecker2._s();
			this.BoundsChecker2._m(0, 0);
			this.BoundsChecker2._l(w_mm, 0);
			this.BoundsChecker2._l(w_mm, h_mm);
			this.BoundsChecker2._l(0, h_mm);
			this.BoundsChecker2._z();

			return;
		}

		this.m_oWordControl.m_oLogicDocument.DrawPage(slideNum, this.BoundsChecker2);
	}

	this.CheckSlide = function(slideNum)
	{
		if (this.m_oWordControl.m_oApi.isReporterMode)
			return;

		if (this.m_oWordControl.m_oApi.isSaveFonts_Images)
		{
			this.IsRecalculateSlide = true;
			return;
		}
		this.IsRecalculateSlide = false;

		this.bIsEmptyPresentation = false;
		if (-1 == slideNum)
			this.bIsEmptyPresentation = true;

		var dKoef = this.m_oWordControl.m_nZoomValue * g_dKoef_mm_to_pix / 100;
		dKoef *= AscCommon.AscBrowser.retinaPixelRatio;

		var w_mm = this.m_oWordControl.m_oLogicDocument.GetWidthMM();
		var h_mm = this.m_oWordControl.m_oLogicDocument.GetHeightMM();
		var w_px = (w_mm * dKoef) >> 0;
		var h_px = (h_mm * dKoef) >> 0;

		this.BoundsChecker.init(w_px, h_px, w_mm, h_mm);
		this.BoundsChecker.transform(1, 0, 0, 1, 0, 0);

		if (this.bIsEmptyPresentation)
		{
			this.BoundsChecker._s();
			this.BoundsChecker._m(0, 0);
			this.BoundsChecker._l(w_mm, 0);
			this.BoundsChecker._l(w_mm, h_mm);
			this.BoundsChecker._l(0, h_mm);
			this.BoundsChecker._z();

			return;
		}

		this.m_oWordControl.m_oLogicDocument.DrawPage(slideNum, this.BoundsChecker);

		var bIsResize = this.m_oWordControl.CheckCalculateDocumentSize(this.BoundsChecker.Bounds);

		if (true)
		{
			// поидее если был ресайз только
			dKoef = this.m_oWordControl.m_nZoomValue * g_dKoef_mm_to_pix / 100;
			dKoef *= AscCommon.AscBrowser.retinaPixelRatio;

			w_mm = this.m_oWordControl.m_oLogicDocument.GetWidthMM();
			h_mm = this.m_oWordControl.m_oLogicDocument.GetHeightMM();
			w_px = (w_mm * dKoef) >> 0;
			h_px = (h_mm * dKoef) >> 0;
		}

		// теперь смотрим, используем ли кэш для скролла
		var _need_pix_width  = this.BoundsChecker.Bounds.max_x - this.BoundsChecker.Bounds.min_x + 1 + 2 * this.SlideEps;
		var _need_pix_height = this.BoundsChecker.Bounds.max_y - this.BoundsChecker.Bounds.min_y + 1 + 2 * this.SlideEps;

		if (this.m_oWordControl.NoneRepaintPages)
			return;

		this.CacheSlidePixW = _need_pix_width;
		this.CacheSlidePixH = _need_pix_height;

		this.IsCached = false;
		if (4 * _need_pix_width * _need_pix_height < this.CONST_MAX_SLIDE_CACHE_SIZE)
			this.IsCached = true;

		// See bug 68871
		if (this.m_oWordControl.m_oApi.isMobileVersion)
			this.IsCached = false;

		if (this.IsCached)
		{
			// кэш используется. теперь нужно решить, нужно ли создать картинку, или управимся и старой
			var _need_reinit_image = false;
			if (null == this.CachedCanvas)
				_need_reinit_image = true;
			else
			{
				if (this.CachedCanvas.width < _need_pix_width || this.CachedCanvas.height < _need_pix_height)
					_need_reinit_image = true;
			}

			if (_need_reinit_image)
			{
				// все равно перевыделяем память - сделаем небольшой задел, чтобы обезопасить
				// себя от чувака, который будет быстро выдвигать/задвигать элемент на "чуть-чуть"

				this.CachedCanvas        = null;
				this.CachedCanvas        = document.createElement('canvas');
				this.CachedCanvas.width  = _need_pix_width + 100;
				this.CachedCanvas.height = _need_pix_height + 100;

				this.CachedCanvasCtx = this.CachedCanvas.getContext('2d');
			}
			else
			{
				// здесь просто нужно очистить место под новую отрисовку
				this.CachedCanvasCtx.setTransform(1, 0, 0, 1, 0, 0);
				this.CachedCanvasCtx.clearRect(0, 0, _need_pix_width, _need_pix_height);
			}

			// и сразу отрисуем его на кешированной картинке
			var g = new AscCommon.CGraphics();
			g.init(this.CachedCanvasCtx, w_px, h_px, w_mm, h_mm);
			g.m_oFontManager = AscCommon.g_fontManager;

			g.m_oCoordTransform.tx = -this.BoundsChecker.Bounds.min_x + this.SlideEps;
			g.m_oCoordTransform.ty = -this.BoundsChecker.Bounds.min_y + this.SlideEps;
			g.transform(1, 0, 0, 1, 0, 0);

			if (this.m_oWordControl.m_oApi.isViewMode)
				g.IsNoDrawingEmptyPlaceholderText = true;

			this.m_oWordControl.m_oLogicDocument.DrawPage(slideNum, g);
		}
		else
		{
			if (null != this.CachedCanvas)
			{
				this.CachedCanvas    = null;
				this.CachedCanvasCtx = null;
			}
		}
	}

	this.DrawSlide = function(outputCtx, scrollX, scrollX_max, scrollY, scrollY_max, slideNum)
	{
		if (this.m_oWordControl.m_oApi.isReporterMode)
			return;

		var _rect   = this.m_oWordControl.m_oDrawingDocument.SlideCurrectRect;
		var _bounds = this.BoundsChecker.Bounds;

		var _x = _rect.left + _bounds.min_x;
		var _y = _rect.top + _bounds.min_y;

		_x = ((_rect.left * AscCommon.AscBrowser.retinaPixelRatio) >> 0) + _bounds.min_x;
		_y = ((_rect.top * AscCommon.AscBrowser.retinaPixelRatio) >> 0) + _bounds.min_y;

		if (this.bIsEmptyPresentation)
		{
			var w_px = _bounds.max_x - _bounds.min_x + 1;
			var h_px = _bounds.max_y - _bounds.min_y + 1;

			outputCtx.lineWidth   = 1;
			outputCtx.strokeStyle = "#000000";

			outputCtx.beginPath();
			this.m_oWordControl.m_oDrawingDocument.AutoShapesTrack.AddRectDashClever(outputCtx, _x >> 0, _y >> 0, (_x + w_px) >> 0, (_y + h_px) >> 0, 2, 2, true);
			outputCtx.beginPath();

			outputCtx.fillStyle = "#3C3C3C";

			var fontCtx = ((this.m_oWordControl.m_nZoomValue * 60 / 100) >> 0) + "px Arial";
			outputCtx.font = fontCtx;

			var text = AscCommon.translateManager.getValue("Click to add first slide");
			var textWidth = outputCtx.measureText(text).width;

			var xPos = ((_x >> 0) + ((w_px - textWidth) >> 1));
			var yPos = ((_y >> 0) + ((h_px - this.EmptyPresenattionTextHeight) >> 1));
			outputCtx.fillText(text, xPos, yPos);

			return;
		}

		if (this.IsCached)
		{
			if (!this.m_oWordControl.NoneRepaintPages)
			{
				var w_px = (_bounds.max_x - _bounds.min_x + 1 + 2 * this.SlideEps) >> 0;
				var h_px = (_bounds.max_y - _bounds.min_y + 1 + 2 * this.SlideEps) >> 0;

				if (w_px > this.CachedCanvas.width)
					w_px = this.CachedCanvas.width;
				if (h_px > this.CachedCanvas.height)
					h_px = this.CachedCanvas.height;

				outputCtx.drawImage(this.CachedCanvas, 0, 0, w_px, h_px, (_x >> 0) - this.SlideEps, (_y >> 0) - this.SlideEps, w_px, h_px);
			}
			else
			{
				var w_px = (_bounds.max_x - _bounds.min_x + 1 + 2 * this.SlideEps) >> 0;
				var h_px = (_bounds.max_y - _bounds.min_y + 1 + 2 * this.SlideEps) >> 0;

				var w_px_src = this.CacheSlidePixW;
				var h_px_src = this.CacheSlidePixH;
				if (w_px_src > this.CachedCanvas.width)
					w_px_src = this.CachedCanvas.width;
				if (h_px_src > this.CachedCanvas.height)
					h_px_src = this.CachedCanvas.height;

				outputCtx.drawImage(this.CachedCanvas, 0, 0, w_px_src, h_px_src, (_x >> 0) - this.SlideEps, (_y >> 0) - this.SlideEps, w_px, h_px);
			}
		}
		else
		{
			var dKoef = this.m_oWordControl.m_nZoomValue * g_dKoef_mm_to_pix / 100;
			dKoef *= AscCommon.AscBrowser.retinaPixelRatio;

			var w_mm = this.m_oWordControl.m_oLogicDocument.GetWidthMM();
			var h_mm = this.m_oWordControl.m_oLogicDocument.GetHeightMM();
			var w_px = (w_mm * dKoef) >> 0;
			var h_px = (h_mm * dKoef) >> 0;

			var g = new AscCommon.CGraphics();
			g.init(outputCtx, w_px, h_px, w_mm, h_mm);
			g.m_oFontManager = AscCommon.g_fontManager;

			g.m_oCoordTransform.tx = _x - _bounds.min_x;
			g.m_oCoordTransform.ty = _y - _bounds.min_y;
			g.transform(1, 0, 0, 1, 0, 0);

			if (this.m_oWordControl.m_oApi.isViewMode)
				g.IsNoDrawingEmptyPlaceholderText = true;

			this.m_oWordControl.m_oLogicDocument.DrawPage(slideNum, g);
		}

		if (this.m_oWordControl.m_oApi.watermarkDraw &&
			!(this.m_oWordControl.DemonstrationManager && this.m_oWordControl.DemonstrationManager.Mode) &&
			!this.m_oWordControl.m_oDrawingDocument.TransitionSlide.IsPlaying())
		{
			this.m_oWordControl.m_oApi.watermarkDraw.Draw(outputCtx,
				AscCommon.AscBrowser.convertToRetinaValue(_rect.left, true),
				AscCommon.AscBrowser.convertToRetinaValue(_rect.top, true),
				AscCommon.AscBrowser.convertToRetinaValue(_rect.right - _rect.left, true),
				AscCommon.AscBrowser.convertToRetinaValue(_rect.bottom - _rect.top, true));
		}
	}
}

function CreateScrollSettings(height, element, OffsetY)
{
	var settings = new AscCommon.ScrollSettings();
	settings.screenW = element.width;
	settings.screenH = element.height;
	settings.vscrollStep = 45;
	settings.hscrollStep = 45;
	settings.contentW = 1;
	settings.contentH = 2 * OffsetY + ((height * g_dKoef_mm_to_pix) >> 0);
	settings.scrollerMinHeight = 5;

	settings.screenW = AscCommon.AscBrowser.convertToRetinaValue(settings.screenW);
	settings.screenH = AscCommon.AscBrowser.convertToRetinaValue(settings.screenH);

	settings.scrollBackgroundColor = GlobalSkin.ScrollBackgroundColor;
	settings.scrollBackgroundColorHover = GlobalSkin.ScrollBackgroundColor;
	settings.scrollBackgroundColorActive = GlobalSkin.ScrollBackgroundColor;

	settings.scrollerColor = GlobalSkin.ScrollerColor;
	settings.scrollerHoverColor = GlobalSkin.ScrollerHoverColor;
	settings.scrollerActiveColor = GlobalSkin.ScrollerActiveColor;

	settings.arrowColor = GlobalSkin.ScrollArrowColor;
	settings.arrowHoverColor = GlobalSkin.ScrollArrowHoverColor;
	settings.arrowActiveColor = GlobalSkin.ScrollArrowActiveColor;

	settings.strokeStyleNone = GlobalSkin.ScrollOutlineColor;
	settings.strokeStyleOver = GlobalSkin.ScrollOutlineHoverColor;
	settings.strokeStyleActive = GlobalSkin.ScrollOutlineActiveColor;

	settings.targetColor = GlobalSkin.ScrollerTargetColor;
	settings.targetHoverColor = GlobalSkin.ScrollerTargetHoverColor;
	settings.targetActiveColor = GlobalSkin.ScrollerTargetActiveColor;

	return settings;
}

function CNotesDrawer(page)
{
	this.Width = 0;
	this.Height = 0;

	this.HtmlPage = page;
	this.TargetHtmlElement = document.getElementById('id_notes_target_cursor');

	this.IsRepaint = false;
	this.Slide = -1;

	this.m_oOverlayApi = new AscCommon.COverlay();
	this.m_oOverlayApi.m_oControl = this.HtmlPage.m_oNotesOverlay;
	this.m_oOverlayApi.m_oHtmlPage = this.HtmlPage;
	this.m_oOverlayApi.Clear();

	this.m_oOverlayApi.getNotesOffsets = function()
	{
		return { X : this.m_oHtmlPage.m_oNotesApi.OffsetX, Y : AscCommon.AscBrowser.convertToRetinaValue(-this.m_oHtmlPage.m_oNotesApi.Scroll, true) };
	};

	this.OffsetX = 10;
	this.OffsetY = 10;

	this.Scroll = 0;
	this.ScrollMax = 0;

	this.fontManager = new AscFonts.CFontManager();
	this.fontManager.Initialize(true);
	this.fontManager.SetHintsProps(true, true);

	this.m_oTimerScrollSelect = -1;

	this.IsEmptyDrawCheck = false;
	this.IsEmptyDraw = false;

	var oThis = this;

	this.Init = function()
	{
		var _elem = this.HtmlPage.m_oNotes;
		var _elemOverlay = this.HtmlPage.m_oNotesOverlay;

		_elem.HtmlElement.onmousedown = this.onMouseDown;
		_elem.HtmlElement.onmousemove = this.onMouseMove;
		_elem.HtmlElement.onmouseup = this.onMouseUp;

		_elemOverlay.HtmlElement.onmousedown = this.onMouseDown;
		_elemOverlay.HtmlElement.onmousemove = this.onMouseMove;
		_elemOverlay.HtmlElement.onmouseup = this.onMouseUp;

		this.HtmlPage.m_oNotesContainer.HtmlElement.onmousewheel = this.onMouseWhell;
		if (this.HtmlPage.m_oNotesContainer.HtmlElement.addEventListener)
		{
			this.HtmlPage.m_oNotesContainer.HtmlElement.addEventListener("DOMMouseScroll", this.onMouseWhell, false);
		}
	};

	// paint
	this.CheckPaint = function()
	{
		if (this.IsRepaint)
			this.OnPaint();
	};

	this.OnPaint = function()
	{
		var element = this.HtmlPage.m_oNotes.HtmlElement;
		var ctx = element.getContext('2d');
		ctx.clearRect(0, 0, element.width, element.height);

		if (-1 == this.Slide || this.IsEmptyDraw)
		{
			this.IsRepaint = false;
			return;
		}

		var dKoef = g_dKoef_mm_to_pix;
		dKoef *= AscCommon.AscBrowser.retinaPixelRatio;

		var w_mm = this.Width;
		var h_mm = this.Height;
		var w_px = (w_mm * dKoef) >> 0;
		var h_px = (h_mm * dKoef) >> 0;

		var g = new AscCommon.CGraphics();
		g.init(ctx, w_px, h_px, w_mm, h_mm);
		g.m_oFontManager = this.fontManager;

		if (AscCommon.GlobalSkin.Type === "dark")
		{
			g.setDarkMode(1);
		}

		g.SaveGrState();

		g.m_oCoordTransform.tx = this.OffsetX;
		g.m_oCoordTransform.ty = AscCommon.AscBrowser.convertToRetinaValue(-this.Scroll, true);
		g.transform(1, 0, 0, 1, 0, 0);

		//g.IsNoDrawingEmptyPlaceholder = true;
		//g.IsNoDrawingEmptyPlaceholderText = true;

        if (this.HtmlPage.m_oApi.isViewMode)
            g.IsNoDrawingEmptyPlaceholderText = true;

		this.HtmlPage.m_oDrawingDocument.isDrawingNotes = true;
		this.HtmlPage.m_oLogicDocument.Notes_Draw(this.Slide, g);
		this.HtmlPage.m_oDrawingDocument.isDrawingNotes = false;
		this.IsRepaint = false;

		g.RestoreGrState();
	};

	// recalculate
	this.OnRecalculateNote = function(slideNum, width, height)
	{
		var isChangedSlide = (this.Slide != slideNum) ? true : false;
		this.Slide = slideNum;
		this.Width = width;
		this.Height = height;
		this.IsRepaint = true;

		if (window["NATIVE_EDITOR_ENJINE"])
			return;

		this.IsEmptyDraw = false;
		if (this.HtmlPage.m_oApi.isReporterMode && this.IsEmptyDrawCheck && !this.HtmlPage.m_oLogicDocument.IsVisibleSlide(this.Slide))
		{
			height = 0;
			this.IsEmptyDraw = true;
		}

		var settings = CreateScrollSettings(height, this.HtmlPage.m_oNotes.HtmlElement, this.OffsetY);

		this.ScrollMax = Math.max(0, settings.contentH - settings.screenH);
		if (this.Scroll > this.ScrollMax)
			this.Scroll = this.ScrollMax;

		document.getElementById('panel_right_scroll_notes').style.height = settings.contentH + "px";

		if (this.HtmlPage.m_oScrollNotes_)
		{
			this.HtmlPage.m_oScrollNotes_.Repos(settings, undefined, true);

			if (isChangedSlide)
				this.HtmlPage.m_oScrollNotes_.scrollToY(0);
		}
		else
		{
			this.HtmlPage.m_oScrollNotes_ = new AscCommon.ScrollObject("id_vertical_scroll_notes", settings);

			this.HtmlPage.m_oScrollNotes_.onLockMouse = function (evt)
			{
				AscCommon.check_MouseDownEvent(evt, true);
				global_mouseEvent.LockMouse();
			};
			this.HtmlPage.m_oScrollNotes_.offLockMouse = function (evt)
			{
				AscCommon.check_MouseUpEvent(evt);
			};

			this.HtmlPage.m_oScrollNotes_.bind("scrollvertical", function (evt)
			{
				oThis.Scroll = (oThis.ScrollMax * evt.scrollD / Math.max(evt.maxScrollY, 1)) >> 0;
				oThis.HtmlPage.m_bIsUpdateTargetNoAttack = true;
				oThis.IsRepaint = true;

				oThis.HtmlPage.StartUpdateOverlay();
				oThis.HtmlPage.OnUpdateOverlay();
				oThis.HtmlPage.EndUpdateOverlay();
			});
		}
	};

	// common
	this.GetCurrentSlideNumber = function()
	{
		return oThis.GetDrawingDocument().SlideCurrent;
	};

	this.GetPresentation = function()
	{
		return oThis.HtmlPage.m_oLogicDocument;
	};
	this.GetDrawingDocument = function()
	{
		return oThis.HtmlPage.m_oDrawingDocument;
	};

	this.GetRelPos = function ()
	{
		var _x = global_mouseEvent.X - oThis.HtmlPage.X - ((oThis.HtmlPage.m_oMainParent.AbsolutePosition.L * g_dKoef_mm_to_pix + 0.5) >> 0);
		var _y = global_mouseEvent.Y - oThis.HtmlPage.Y - ((oThis.HtmlPage.m_oNotesContainer.Parent.AbsolutePosition.T * g_dKoef_mm_to_pix + 0.5) >> 0);
		return {X: _x, Y: _y};
	};

	this.GetPosition = function ()
	{
		var oPos = oThis.GetRelPos();
		var _x = oPos.X;
		var _y = oPos.Y;

		_x -= oThis.OffsetX;
		_y += oThis.Scroll;
		_x *= g_dKoef_pix_to_mm;
		_y *= g_dKoef_pix_to_mm;
		return { Page : oThis.GetCurrentSlideNumber(), X : _x, Y : _y, isNotes : true };
	};

	this.GetNotesWidth = function()
	{
		var _pix_width = this.HtmlPage.m_oNotes.HtmlElement.width - AscCommon.AscBrowser.convertToRetinaValue(30, true);
		if (_pix_width < 10)
			_pix_width = 10;
		_pix_width = AscCommon.AscBrowser.convertToRetinaValue(_pix_width);
		if (window["NATIVE_EDITOR_ENJINE"] && _pix_width < 100) _pix_width = 100;
		return _pix_width / g_dKoef_mm_to_pix;
	};

	// events
	this.onMouseDown = function (e)
	{
		if (-1 == oThis.HtmlPage.m_oDrawingDocument.SlideCurrent)
			return;

		// после fullscreen возможно изменение X, Y после вызова Resize.
		oThis.HtmlPage.checkBodyOffset();

		AscCommon.check_MouseDownEvent(e, true);
		global_mouseEvent.LockMouse();

		oThis.HtmlPage.Thumbnails.SetFocusElement(FOCUS_OBJECT_MAIN);

		if (-1 == oThis.m_oTimerScrollSelect)
		{
			oThis.m_oTimerScrollSelect = setInterval(oThis.onSelectWheel, 20);
		}

		var pos = oThis.GetPosition();

		var ret = oThis.HtmlPage.m_oDrawingDocument.checkMouseDown_Drawing(pos);
		if (ret === true)
		{
			AscCommon.stopEvent(e);
			return;
		}

		oThis.HtmlPage.StartUpdateOverlay();
		oThis.HtmlPage.m_oLogicDocument.Notes_OnMouseDown(global_mouseEvent, pos.X, pos.Y);
		oThis.HtmlPage.EndUpdateOverlay();
	};
	this.onMouseMove = function (e, is_overlay_attack)
	{
		if (-1 == oThis.HtmlPage.m_oDrawingDocument.SlideCurrent)
			return;

		if (e)
			AscCommon.check_MouseMoveEvent(e);


		var pos = oThis.GetPosition();
		var _x = pos.X;
		var _y = pos.Y;

		if (oThis.HtmlPage.m_oDrawingDocument.InlineTextTrackEnabled)
		{
			if (_y < 0)
				return;
		}
		oThis.HtmlPage.StartUpdateOverlay();
		if ((-1 != oThis.m_oTimerScrollSelect) || (is_overlay_attack === true))
			oThis.HtmlPage.OnUpdateOverlay();

		var is_drawing = oThis.HtmlPage.m_oDrawingDocument.checkMouseMove_Drawing(pos);
		if (is_drawing === true)
			return;

		oThis.HtmlPage.m_oLogicDocument.Notes_OnMouseMove(global_mouseEvent, _x, _y);
		oThis.HtmlPage.EndUpdateOverlay();
	};
	this.onMouseUp = function (e)
	{
		if (-1 == oThis.HtmlPage.m_oDrawingDocument.SlideCurrent)
			return;

		AscCommon.check_MouseUpEvent(e);

		var pos = oThis.GetPosition();
		var _x = pos.X;
		var _y = pos.Y;

		if (-1 != oThis.m_oTimerScrollSelect)
		{
			clearInterval(oThis.m_oTimerScrollSelect);
			oThis.m_oTimerScrollSelect = -1;
		}

		if (oThis.HtmlPage.m_oDrawingDocument.InlineTextTrackEnabled)
		{
			if (_y < 0)
				return;
		}

		oThis.HtmlPage.StartUpdateOverlay();


		var is_drawing = oThis.HtmlPage.m_oDrawingDocument.checkMouseUp_Drawing(pos);
		if (is_drawing === true)
			return;

		oThis.HtmlPage.m_oLogicDocument.Notes_OnMouseUp(global_mouseEvent, _x, _y);
		oThis.HtmlPage.EndUpdateOverlay();

		oThis.HtmlPage.m_bIsMouseLock = false;
	};

	this.onMouseWhell = function(e)
	{
		if (false === oThis.HtmlPage.m_oApi.bInit_word_control)
			return;

		var _ctrl = false;
		if (e.metaKey !== undefined)
			_ctrl = e.ctrlKey || e.metaKey;
		else
			_ctrl = e.ctrlKey;

		if (true === _ctrl)
		{
			AscCommon.stopEvent(e);
			return false;
		}

		var deltaY = GetWheelDeltaY(e);

		if (0 != deltaY)
			oThis.HtmlPage.m_oScrollNotes_.scrollBy(0, deltaY, false);

		// здесь - имитируем моус мув ---------------------------
		oThis.onMouseMove(CreateBrowserEventObject(), true);
		// ------------------------------------------------------

		AscCommon.stopEvent(e);
		return false;
	};

	this.onSelectWheel = function()
	{
		if (false === oThis.HtmlPage.m_oApi.bInit_word_control)
			return;

		var _y = oThis.GetRelPos().Y;

		var positionMinY = 0;
		var positionMaxY = oThis.HtmlPage.m_oNotes.AbsolutePosition.B * g_dKoef_mm_to_pix;

		var scrollYVal = 0;
		if (_y < positionMinY)
		{
			var delta = 30;
			if (20 > (positionMinY - _y))
				delta = 10;

			scrollYVal = -delta;
		}
		else if (_y > positionMaxY)
		{
			var delta = 30;
			if (20 > (_y - positionMaxY))
				delta = 10;

			scrollYVal = delta;
		}

		if (0 != scrollYVal)
		{
			oThis.HtmlPage.m_oScrollNotes_.scrollByY(scrollYVal, false);
			oThis.onMouseMove();
		}
	};

	this.OnResize = function ()
	{
		if (this.HtmlPage.m_oLogicDocument)
		{
			var oldEmpty = this.IsEmptyDraw;
            if (!this.HtmlPage.m_oLogicDocument.Notes_OnResize())
			{
				this.OnRecalculateNote(this.Slide, this.Width, this.Height);
				this.HtmlPage.m_oLogicDocument.RecalculateCurPos();
			}
			this.IsEmptyDraw = oldEmpty;
		}
	};
}

function CAnimPaneDrawTask()
{
	this.bDraw = false;
	this.Rect = null;
}
CAnimPaneDrawTask.prototype.Check = function(oRect)
{
	this.bDraw = true;
	if(this.Rect)
	{
		if(!oRect)
		{
			this.Rect = null;
		}
		else
		{
			this.Rect.checkByOther(oRect);
		}
	}
};
CAnimPaneDrawTask.prototype.NeedRedraw = function()
{
	return this.bDraw;
};
CAnimPaneDrawTask.prototype.Clear = function()
{
	this.bDraw = false;
	this.Rect = null;
};
CAnimPaneDrawTask.prototype.GetRect = function()
{
	return this.Rect;
};

function GetWheelDeltaY(e)
{
	var delta  = 0;
	var deltaX = 0;
	var deltaY = 0;

	if (undefined != e.wheelDelta && e.wheelDelta != 0)
	{
		//delta = (e.wheelDelta > 0) ? -45 : 45;
		delta = -45 * e.wheelDelta / 120;
	}
	else if (undefined != e.detail && e.detail != 0)
	{
		//delta = (e.detail > 0) ? 45 : -45;
		delta = 45 * e.detail / 3;
	}

	// New school multidimensional scroll (touchpads) deltas
	deltaY = delta;
	deltaY >>= 0;
	return deltaY;
}

function CreateBrowserEventObject()
{
	var _e   = {};
	_e.pageX = global_mouseEvent.X;
	_e.pageY = global_mouseEvent.Y;

	_e.clientX = global_mouseEvent.X;
	_e.clientY = global_mouseEvent.Y;

	_e.altKey   = global_mouseEvent.AltKey;
	_e.shiftKey = global_mouseEvent.ShiftKey;
	_e.ctrlKey  = global_mouseEvent.CtrlKey;
	_e.metaKey  = global_mouseEvent.CtrlKey;

	_e.srcElement = global_mouseEvent.Sender;
	return _e;
}

function CPaneDrawerBase(page, htmlElement, parentDrawer, pageControl)
{
	this.parentDrawer = parentDrawer ? parentDrawer : null;
	if(this.parentDrawer)
	{
		this.fontManager = this.parentDrawer.fontManager;
	}
	else
	{
		this.fontManager = new AscFonts.CFontManager();
		this.fontManager.Initialize(true);
		this.fontManager.SetHintsProps(true, true);
	}

	this.HtmlPage = page;
	this.HtmlElement = htmlElement;
	this.PageParentControl = pageControl;

	this.DrawTask = new CAnimPaneDrawTask();

	this.Control = null;


	this.Scroll = 0;
	this.ScrollMax = 0;

	var oThis = this;

	oThis.CreateControl = function()
	{
	};
	oThis.CheckControl = function()
	{
		if(!oThis.HtmlPage.m_oApi.isDocumentLoadComplete)
		{
			return;
		}
		if(!oThis.Control)
		{
			if(oThis.HtmlPage.m_oLogicDocument)
			{
				oThis.CreateControl();
			}
		}
	};
	oThis.GetControl = function()
	{
		oThis.CheckControl();
		return oThis.Control;
	};
	oThis.GetHtmlElement = function()
	{
		return oThis.HtmlElement;
	};
	oThis.CheckSubscribeMouseWheel = function()
	{};
	oThis.Init = function ()
	{

		var oHtmlElem = oThis.GetHtmlElement();
		AscCommon.addMouseEvent(oHtmlElem, "up", oThis.onMouseUp);
		oThis.CheckSubscribeMouseWheel();
	};
	oThis.GetCurrentSlideNumber = function ()
	{
		return oThis.GetDrawingDocument().SlideCurrent;
	};
	oThis.GetPresentation = function ()
	{
		return oThis.HtmlPage.m_oLogicDocument;
	};
	oThis.GetDrawingDocument = function ()
	{
		return oThis.HtmlPage.m_oDrawingDocument;
	};
	oThis.OnPaint = function ()
	{
		var oControl = oThis.GetControl();
		if(!oControl)
		{
			return;
		}
		var element = oThis.GetHtmlElement();
		var ctx = element.getContext('2d');

		var dKoef = g_dKoef_mm_to_pix;
		dKoef *= AscCommon.AscBrowser.retinaPixelRatio;

		var w_mm = oThis.GetWidth();
		var h_mm = oThis.GetHeight();
		var w_px = (w_mm * dKoef) >> 0;
		var h_px = (h_mm * dKoef) >> 0;

		var g = new AscCommon.CGraphics();
		g.init(ctx, w_px, h_px, w_mm, h_mm);
		g.m_oCoordTransform.ty = AscCommon.AscBrowser.convertToRetinaValue(-oThis.Scroll, true);
		g.m_oFontManager = oThis.fontManager;

		var oUpdateRect = oThis.DrawTask.GetRect();
		var oClipRect = null;
		if(oUpdateRect)
		{
			var oT = g.m_oCoordTransform;
			var l = (oT.TransformPointX(oUpdateRect.l, oUpdateRect.t) >> 0) - 1;
			var t = (oT.TransformPointY(oUpdateRect.l, oUpdateRect.t) >> 0) - 1;
			var r = (oT.TransformPointX(oUpdateRect.r, oUpdateRect.b) >> 0) + 1;
			var b = (oT.TransformPointY(oUpdateRect.r, oUpdateRect.b) >> 0) + 1;
			oClipRect = new AscFormat.CGraphicBounds(l, t, r, b);
			ctx.clearRect(oClipRect.l, oClipRect.t, oClipRect.w, oClipRect.h);
			g.updateRect = oUpdateRect;
		}
		else
		{
			ctx.clearRect(0, 0, element.width, element.height);
		}
		ctx.fillStyle = AscCommon.GlobalSkin.AnimPaneBackground;
		ctx.fillRect(0, 0, element.width, element.height);
		if(oClipRect)
		{
			ctx.save();
			ctx.beginPath();
			ctx.rect(oClipRect.x, oClipRect.y, oClipRect.w, oClipRect.h);
			ctx.clip();
			ctx.save();
		}
		g.SaveGrState();

		if (AscCommon.GlobalSkin.Type === "dark")
		{
			g.setDarkMode(1);
		}
		oControl.draw(g);
		g.RestoreGrState();
		if(oClipRect)
		{
			ctx.restore();
			ctx.restore();
		}
	};
	oThis.CheckPaint = function ()
	{
		if(oThis.DrawTask.NeedRedraw())
		{
			oThis.OnPaint();
			oThis.DrawTask.Clear();
		}
	};
	oThis.IsPresentationEmpty = function ()
	{
		return -1 === oThis.GetCurrentSlideNumber();
	};
	oThis.GetPosition = function ()
	{
		let nLeftPos = oThis.HtmlPage.m_oMainParent.AbsolutePosition.L;
		let nTopPos = oThis.HtmlPage.m_oBottomPanesContainer.AbsolutePosition.T;
		nTopPos += oThis.HtmlPage.m_oAnimationPaneContainer.AbsolutePosition.T;
		if(oThis.PageParentControl)
		{
			nLeftPos += oThis.PageParentControl.AbsolutePosition.L;
			nTopPos += oThis.PageParentControl.AbsolutePosition.T;
		}

		let _x = global_mouseEvent.X - oThis.HtmlPage.X - ((nLeftPos * g_dKoef_mm_to_pix + 0.5) >> 0);
		let _y = global_mouseEvent.Y - oThis.HtmlPage.Y - ((nTopPos * g_dKoef_mm_to_pix + 0.5) >> 0);
		_y += oThis.Scroll;
		_x *= g_dKoef_pix_to_mm;
		_y *= g_dKoef_pix_to_mm;
		return { Page : oThis.GetCurrentSlideNumber(), X : _x, Y : _y, isNotes : false };
	};
	oThis.onMouseDown = function (e)
	{
		if(oThis.IsPresentationEmpty())
			return;
		var oControl = oThis.GetControl();
		if(!oControl)
		{
			return;
		}

		AscCommon.check_MouseDownEvent(e, true);
		global_mouseEvent.LockMouse();
		oThis.HtmlPage.Thumbnails.SetFocusElement(FOCUS_OBJECT_MAIN);
		var pos = oThis.GetPosition(e);
		var _x = pos.X;
		var _y = pos.Y;
		var ret = oThis.GetDrawingDocument().checkMouseDown_Drawing(pos);
		if (ret === true)
		{
			AscCommon.stopEvent(e);
			return;
		}
		oControl.onMouseDown(global_mouseEvent, _x, _y);
		//oThis.GetPresentation().AnimPane_OnMouseDown(global_mouseEvent, _x, _y);

		return oControl.hit(_x, _y);
	};
	oThis.onMouseMove = function (e)
	{
		if(oThis.IsPresentationEmpty())
			return;
		var oControl = oThis.GetControl();
		if(!oControl)
		{
			return;
		}

		if (e)
			AscCommon.check_MouseMoveEvent(e);

		var pos = oThis.GetPosition(e);
		var _x = pos.X;
		var _y = pos.Y;

		if (oThis.GetDrawingDocument().InlineTextTrackEnabled)
		{
			if (_y < 0)
				return;
		}

		var is_drawing = oThis.GetDrawingDocument().checkMouseMove_Drawing(pos);
		if (is_drawing === true)
			return;

		oControl.onMouseMove(global_mouseEvent, _x, _y);
		//oThis.GetPresentation().AnimPane_OnMouseMove(global_mouseEvent, _x, _y);
	};
	oThis.onMouseUp = function (e)
	{
		if(oThis.IsPresentationEmpty())
			return;
		var oControl = oThis.GetControl();
		if(!oControl)
		{
			return;
		}

		AscCommon.check_MouseUpEvent(e);


		var pos = oThis.GetPosition(e);
		var _x = pos.X;
		var _y = pos.Y;


		if (oThis.GetDrawingDocument().InlineTextTrackEnabled)
		{
			if (_y < 0)
				return;
		}

		var is_drawing = oThis.GetDrawingDocument().checkMouseUp_Drawing(pos);
		if (is_drawing === true)
			return;


		oControl.onMouseUp(global_mouseEvent, _x, _y);
		//oThis.GetPresentation().AnimPane_OnMouseUp(global_mouseEvent, _x, _y);

		oThis.HtmlPage.m_bIsMouseLock = false;
	};
	oThis.onMouseWhell = function(e)
	{
		// if (false === oThis.HtmlPage.m_oApi.bInit_word_control)
		// 	return;
		// var oControl = oThis.GetControl();
		// if(!oControl)
		// {
		// 	return;
		// }
		//
		// var deltaY = GetWheelDeltaY(e);
		//
		// if (0 != deltaY)
		// {
		// 	var pos = oThis.GetPosition(e);
		// 	var _x = pos.X;
		// 	var _y = pos.Y;
		// 	oControl.onMouseWheel(global_mouseEvent, deltaY, _x, _y);
		// }
		//
		//
		// // здесь - имитируем моус мув ---------------------------
		// var _e   = CreateBrowserEventObject();
		//
		// oThis.onMouseMove(_e, true);
		// // ------------------------------------------------------
		//
		// AscCommon.stopEvent(e);
		return false;
	};
	oThis.onSelectWheel = function()
	{
	};
	oThis.CheckScroll = function()
	{};
	oThis.OnResize = function ()
	{
		var oControl = oThis.GetControl();
		if(!oControl)
		{
			return;
		}
		oControl.onResize();
		oThis.CheckScroll();
	};
	oThis.OnAnimPaneChanged = function (oRect)
	{
		oThis.DrawTask.Check(oRect);
	};

	oThis.GetWidth = function()
	{
		return AscCommon.AscBrowser.convertToRetinaValue(oThis.GetHtmlElement().width) / g_dKoef_mm_to_pix;
	};
	oThis.GetHeight = function()
	{
		return AscCommon.AscBrowser.convertToRetinaValue(oThis.GetHtmlElement().height) / g_dKoef_mm_to_pix;
	};
}

function CAnimPaneHeaderDrawer(page, htmlElement, parentDrawer)
{
	CPaneDrawerBase.call(this, page, htmlElement, parentDrawer, page.m_oAnimPaneHeaderContainer);
	var oThis = this;
	oThis.CreateControl = function()
	{
		oThis.Control = new AscCommon.CAnimPaneHeader(this);
		oThis.Control.onResize();
	};

	oThis.UpdateButtonsState = function () {
		Asc.editor.asc_canStartAnimationPreview() ?
			oThis.Control.playButton.enable() : oThis.Control.playButton.disable()
		Asc.editor.asc_canMoveAnimationEarlier() ?
			oThis.Control.moveUpButton.enable() : oThis.Control.moveUpButton.disable()
		Asc.editor.asc_canMoveAnimationLater() ?
			oThis.Control.moveDownButton.enable() : oThis.Control.moveDownButton.disable()
	}
	oThis.UpdateState = function () {
		oThis.Control.checkLayout();
		oThis.UpdateButtonsState();
	}
}

function CAnimPaneListDrawer(page, htmlElement, parentDrawer)
{
	CPaneDrawerBase.call(this, page, htmlElement, parentDrawer, page.m_oAnimPaneListContainer);
	this.SlideNum = -1;
	var oThis = this;
	oThis.CreateControl = function()
	{
		oThis.Control = new AscCommon.CSeqListContainer(this);
		oThis.Control.onResize();
	};

	oThis.CheckSubscribeMouseWheel = function()
	{
		if(oThis.parentDrawer)
		{
			var oHtmlElem = oThis.parentDrawer.GetHtmlElement();
			oHtmlElem.onmousewheel = oThis.onMouseWhell;
			if (oHtmlElem.addEventListener)
			{
				oHtmlElem.addEventListener("DOMMouseScroll", oThis.onMouseWhell, false);
			}
		}
	};

	oThis.onMouseWhell = function(e)
	{
		if (false === oThis.HtmlPage.m_oApi.bInit_word_control)
			return;

		var _ctrl = false;
		if (e.metaKey !== undefined)
			_ctrl = e.ctrlKey || e.metaKey;
		else
			_ctrl = e.ctrlKey;

		if (true === _ctrl)
		{
			AscCommon.stopEvent(e);
			return false;
		}

		var deltaY = GetWheelDeltaY(e);

		if (0 != deltaY)
		{
			if(oThis.HtmlPage.m_oScrollAnim_)
			{
				oThis.HtmlPage.m_oScrollAnim_.scrollBy(0, deltaY, false);
			}
		}

		// здесь - имитируем моус мув ---------------------------
		oThis.onMouseMove(CreateBrowserEventObject(), true);
		// ------------------------------------------------------

		AscCommon.stopEvent(e);
		return false;
	};

	oThis.OnUpdateList = function()
	{
		var settings = CreateScrollSettings(oThis.Control.getHeight(), oThis.HtmlElement, 0);
		oThis.ScrollMax = Math.max(0, settings.contentH - settings.screenH);
		if (oThis.Scroll > oThis.ScrollMax)
			oThis.Scroll = oThis.ScrollMax;

		document.getElementById('panel_right_scroll_notes').style.height = settings.contentH + "px";
		var nCurSlideNum = oThis.Control.getSlideNum();
		var bChangedSlide = false;
		if(oThis.SlideNum !== nCurSlideNum)
		{
			bChangedSlide = true;
			oThis.SlideNum = nCurSlideNum;
		}
		if (oThis.HtmlPage.m_oScrollAnim_)
		{
			oThis.HtmlPage.m_oScrollAnim_.Repos(settings, undefined, true);
			if(bChangedSlide)
			{
				oThis.HtmlPage.m_oScrollAnim_.scrollToY(0);
			}
		}
		else
		{
			oThis.HtmlPage.m_oScrollAnim_ = new AscCommon.ScrollObject("id_anim_list_scroll", settings);

			oThis.HtmlPage.m_oScrollAnim_.onLockMouse = function (evt)
			{
				AscCommon.check_MouseDownEvent(evt, true);
				global_mouseEvent.LockMouse();
			};
			oThis.HtmlPage.m_oScrollAnim_.offLockMouse = function (evt)
			{
				AscCommon.check_MouseUpEvent(evt);
			};

			oThis.HtmlPage.m_oScrollAnim_.bind("scrollvertical", function (evt)
			{
				oThis.Scroll = (oThis.ScrollMax * evt.scrollD / Math.max(evt.maxScrollY, 1)) >> 0;
				oThis.DrawTask.Check(nCurSlideNum, null);
			});
		}
	};

	oThis.CheckScroll = function()
	{
		oThis.OnUpdateList();
	};
}

function CAnimPaneTimelineDrawer(page, htmlElement, parentDrawer)
{
	CPaneDrawerBase.call(this, page, htmlElement, parentDrawer, page.m_oAnimPaneTimelineContainer);
	const oThis = this;

	// this.oCurSlide = editor.WordControl.m_oLogicDocument.GetCurrentSlide();
	this.oCurSlide = editor.WordControl.m_oLogicDocument ? editor.WordControl.m_oLogicDocument.GetCurrentSlide() : {};

	oThis.CreateControl = function()
	{
		oThis.Control = new AscCommon.CTimelineContainer(this);
		oThis.Control.onResize();
	};
	oThis.UpdateState = function () {
		const timing = oThis.Control.getTiming()
		timing && timing.hasEffects() ?
			oThis.Control.show() : oThis.Control.hide()

		const actualCurrentSlide = Asc.editor.WordControl.m_oLogicDocument.GetCurrentSlide();
		if (this.oCurSlide !== actualCurrentSlide) {
			this.oCurSlide = actualCurrentSlide;
			oThis.Control.timeline.setStartTime(0);
			oThis.Control.timeline.setScrollOffset(0);
		}

	}
}

function CAnimationPaneDrawer(page, htmlElement)
{
	CPaneDrawerBase.call(this, page, htmlElement, null, null);
	this.header = new CAnimPaneHeaderDrawer(page, page.m_oAnimPaneHeader.HtmlElement, this);
	this.list = new CAnimPaneListDrawer(page, page.m_oAnimPaneList.HtmlElement, this);
	this.timeline = new CAnimPaneTimelineDrawer(page, page.m_oAnimPaneTimeline.HtmlElement, this);

	var oThis = this;
	oThis.Init = function()
	{
		oThis.header.Init();
		oThis.list.Init();
		oThis.timeline.Init();

		var oHtmlElem = oThis.GetHtmlElement();

		AscCommon.addMouseEvent(oHtmlElem, "down", oThis.onMouseDown);
		AscCommon.addMouseEvent(oHtmlElem, "move", oThis.onMouseMove);
		AscCommon.addMouseEvent(oHtmlElem, "up", oThis.onMouseUp);

		Asc.editor.asc_registerCallback('asc_onFocusObject', function () {
			// Here we need to check if all animEffects havent been changed
			// If they were - recalculate corresponding elements
			// If they were not - redraw animItems based on "selected" state of effects

			const seqListContainer = oThis.list.Control;
			if (!seqListContainer) { return; }

			// Compare number of sequences (main and interactive ones)
			const timing = seqListContainer.getTiming();
			const newSeqList = timing ? timing.getRootSequences() : [];
			const oldSeqList = seqListContainer.seqList
				? seqListContainer.seqList.children.map(function (animSequence) {
					return animSequence.getSeq();
				})
				: [];

			if (oldSeqList.length !== newSeqList.length) {
				recalculateSeqListContainer();
				return;
			}

			oldSeqList.some(function (_, nSeq) {
				// Compare sequences by Id
				const oldSeq = oldSeqList[nSeq];
				const newSeq = newSeqList[nSeq];

				if (oldSeq.Id !== newSeq.Id) {
					recalculateSeqListContainer();
					return true;
				}

				// Compare number of groups in current sequence
				const oldSeqGroups = seqListContainer.seqList.children[nSeq].animGroups.map(function (animGroup) {
					return animGroup.effects;
				});
				const newSeqGroupsAsObject = AscFormat.groupBy(
					newSeq.getAllEffects(),
					function (effect) { return effect.getIndexInSequence(); }
				);
				const newSeqGroups = Object.keys(newSeqGroupsAsObject).map(function (groupIndex) {
					return newSeqGroupsAsObject[groupIndex];
				})

				if (oldSeqGroups.length !== newSeqGroups.length) {
					recalculateSeqListContainer();
					return true;
				}

				for (let nGroup = 0; nGroup < oldSeqGroups.length; ++nGroup) {
					// Compare number of effects in current group
					const oldSeqGroup = oldSeqGroups[nGroup];
					const newSeqGroup = newSeqGroups[nGroup];

					if (oldSeqGroup.length !== newSeqGroup.length) {
						recalculateSeqListContainer();
						return true;
					}

					for (let nEffect = 0; nEffect < oldSeqGroup.length; ++nEffect) {
						// Compare effects in currect group by Id
						const oldEffect = oldSeqGroup[nEffect];
						const newEffect = newSeqGroup[nEffect];

						if (oldEffect.Id !== newEffect.Id) {
							recalculateSeqListContainer();
							return true;
						}
					}
				}

				seqListContainer.seqList.forEachAnimItem(function (animItem) {
					animItem.onUpdate();
				})

				return false;
			});

			function recalculateSeqListContainer() {
				oThis.list.Control.seqList.recalculateChildren();
				oThis.list.Control.seqList.recalculateChildrenLayout();
				oThis.list.Control.recalculateChildrenLayout();
				oThis.list.Control.onUpdate();
				oThis.list.CheckScroll();
			}
		});

		Asc.editor.asc_registerCallback('asc_onFocusObject', function () {
			/*
				Update effectLabel if needed
				(when shape name has been changed)
			*/

			if (!oThis.list.Control) return;

			let changedLabelsCount = 0;
			oThis.list.Control.seqList.forEachAnimItem(function (animItem) {
				if (animItem.effectLabel.string !== animItem.getEffectLabelText()) {
					animItem.effectLabel.string = animItem.getEffectLabelText();
					changedLabelsCount++;
				}
			});

			if (changedLabelsCount > 0) {
				oThis.list.Control.recalculateChildrenLayout();
			}
		});
	};
	oThis.onMouseDown = function (e)
	{
		oThis.HtmlPage.Thumbnails.SetFocusElement(FOCUS_OBJECT_ANIM_PANE);

		// Order matters
		if (oThis.header.onMouseDown(e)) { return true }
		if (oThis.timeline.onMouseDown(e)) { return true }
		if (oThis.list.onMouseDown(e)) { return true }
		return false;
	};
	oThis.onMouseMove = function (e)
	{
		oThis.sentMouseMoveData = null; // can be overwritten in onMouseMove's below

		oThis.header.onMouseMove(e);
		oThis.list.onMouseMove(e);
		oThis.timeline.onMouseMove(e);

		if (oThis.sentMouseMoveData === null) {
			const data = new AscCommon.CMouseMoveData();

			Asc.editor.sync_MouseMoveStartCallback();
			Asc.editor.sync_MouseMoveCallback(data);
			Asc.editor.sync_MouseMoveEndCallback();
		}
	};
	oThis.onMouseUp = function (e)
	{
	};
	oThis.onMouseWhell = function(e)
	{
		return oThis.list.onMouseWhell(e);
	};
	oThis.onSelectWheel = function()
	{
	};
	oThis.OnResize = function ()
	{
		oThis.header.OnResize();
		oThis.list.OnResize();
		oThis.timeline.OnResize();
	};
	oThis.OnPaint = function ()
	{
	};
	oThis.CheckPaint = function ()
	{
		oThis.header.CheckPaint();
		oThis.list.CheckPaint();
		oThis.timeline.CheckPaint();
	};

	oThis.UpdateState = function () {
		if (!oThis.header.Control) { return }

		oThis.header.UpdateState();
		// TODO: oThis.list.UpdateState();
		oThis.timeline.UpdateState();
	};
	oThis.SetCursorType = function(sType, Data) {
		let elem = this.HtmlElement;
		// if (Asc.editor.WordControl.DemonstrationManager.Mode)
		// 	elem = Asc.editor.WordControl.DemonstrationManager.Canvas;

		if (Asc.editor.WordControl.m_oDrawingDocument.m_sLockedCursorType === '') {
			elem.style.cursor = AscCommon.g_oHtmlCursor.value(sType);
		} else
			elem.style.cursor = AscCommon.g_oHtmlCursor.value(this.m_sLockedCursorType);

		if ("undefined" === typeof (Data) || null === Data)
			Data = new AscCommon.CMouseMoveData();

		Asc.editor.sync_MouseMoveStartCallback();
		Asc.editor.sync_MouseMoveCallback(Data);
		Asc.editor.sync_MouseMoveEndCallback();
	};
}

CAnimationPaneDrawer.prototype = Object.create(CPaneDrawerBase);

//--------------------------------------------------------export----------------------------------------------------
window['AscCommon']                  = window['AscCommon'] || {};
window['AscCommon'].CDrawingDocument = CDrawingDocument;

window['AscCommon'].g_comment_image = new Image();
window['AscCommon'].g_comment_image;
window['AscCommon'].g_comment_image.asc_complete = false;
window['AscCommon'].g_comment_image.onload       = function()
{
    window['AscCommon'].g_comment_image.asc_complete = true;
};
window['AscCommon'].g_comment_image.src = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAR0AAAAlCAMAAABI6s09AAAABGdBTUEAALGPC/xhBQAAAMBQTFRF2Ypk03pP0nZJ1oNb68Ku03hL1X9U7MWy1X9V9d7T3JNv5bCW0nZJ9uXc////8dPE+vDq5rSb5a6U4aOF5K6T2Idg////1HxR6r6o03lN0nZJ6Lef+u7p8NDA0ndK////78++0nZJ8dPE////0nZJ1HtP2Ihh////03hM////////3JVy0nZJ0nZJ9+Xc0nZJ////3JVy3JVx////////AAAA3Zh38tbJ3ZZ08dXH4qaJ5a6U9N3R+/Tw////0nZJQeZHfgAAADZ0Uk5TZJaoUk+v3lzdEi9hDlIbKVN8eY54xTTnc/NKegRg9EJm8WQz7+3EFPYoSKborlkmCqeoJ00ATKvZ0wAAA3tJREFUaN7dmOdi2jAUhelukjaQkkGCGIHYbDAYNYOh93+r2mZIsq6WwVHD+RfukWx93OtjUsCxHoaE0fAB7yWvHFCcjB6JRI+jCc7kzMVfSEzD5zWj5yFdPuSXDLm9C7/nVOMyX5SvnDwRhZ4mWZz5+Dd0yJoToevTK7jNLxHzByryRYZcqemzK0fkbbWWaPVGRqxTqVH6tJ8/XbBfGPPVjVtlT/Tr9t/ToZ+l6bR2l2hxdITJQfLil6/syjqRwonwkDrrVKqePu15fy5XWfTr9s/eO+I0EvlYnRFuz7VCRHF1ZSdHavfOEIaEUHBZE/0XJbjTmuWfyf7Ze0ckqjgWeh86AVaoKPrlrVb6ztGx7h2RKLesRa8UUcUiHei0MJ2KePMVgY4+rQJj/7fzy0YZ6h2AzuacTYCOee8cRKcq0qmm78YgrZCNH/1w2zvHnSyTHOT9mjQsUjreK7vbq0d38fhVnqp3PFXvePnSMclB3q9Jw4DS+XNHFvHuq0X82d013SWqMGIrwjSia6B3dgPJrczhuWNC3Io7onQ6jfk0wrNazOJLNzp0l7iS2IWK0Duoo+gdbmUOz52j08GUTqQwwrOYhkAShjEesSKfRuVA5jRZJsTTO1fgMK8AdHA4+AvCiSsAHMU0KgfyP6JThelUITo4rIaS9yiwIp/GTXGW3NsUKEInUdGpAE+cd56s+EjS10xJRT6N8oHMQOdqzOjKFR17yadxgwcufsTnTjY80mlUFD/kcyeTOhmKXfWbW5d1KtW1nKyu5WR1D6WTRb76rd9nnUr5lnR8Szq+Czq1+/j6L0t698sXel/3tbRTJtZp8KT/5dWUz51Kmo5Xc0Gn3bxJRmaPZ8kMy02zLTrBseKcJnRabZ4Ol4VCGnp+q+2CTpD802m2x7Pc/k7ZqB8ATiqJ02CyEO/XTVa8vws6OLjtM3g4OP3bAHSKcHinCR3er6PTbwfYCZ1EvS2eBE5P69zB6R2agzZp6I7OFo8eDoNH7jTPQZs0dEgnOvRUfWQLp3kO2qShSzo4jA89nYdHcJrnoE0aOqUTHXpgBEfvNM9B1j9goQxEv1s60aHN4Oid5jnI+gcQHOp3TAeH4TGd5jm470gKB9jfNR1nOZjCA8I5NToWOcjhgeGcHB2LHGTwSOCcHh2LHNz7ZXBOkI5FDmr9J0jHIgd1/n8LiumvxDAoYwAAAABJRU5ErkJggg==";
window['AscCommon'].g_comment_image_offsets = [
	[5, 0, 16, 15],
	[31, 0, 16, 15],
	[57, 0, 19, 18],
	[86, 0, 19, 18],
	[115, 0, 32, 30],
	[157, 0, 32, 30],
	[199, 0, 38, 36],
	[247, 0, 38, 36]
];


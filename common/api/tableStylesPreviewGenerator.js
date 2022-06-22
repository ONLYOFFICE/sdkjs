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

"use strict";

(function (window)
{
	let g_oCanvas = null;
	let g_oTable  = null;

	function GetCanvas()
	{
		if (g_oCanvas == null)
		{
			g_oCanvas = document.createElement('canvas');

			g_oCanvas.width = (TABLE_STYLE_WIDTH_PIX * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
			g_oCanvas.height = (TABLE_STYLE_HEIGHT_PIX * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
		}

		return g_oCanvas;
	}
	function GetTable(oLogicDocument)
	{
		if (g_oTable == null)
		{
			g_oTable = oLogicDocument.GetTableForPreview();
		}
		oLogicDocument.CheckTableForPreview(g_oTable);
		return g_oTable;
	}

	/**
	 * @param oLogicDocument
	 * @constructor
	 * @extends AscCommon.CActionOnTimerBase
	 */
	function CTableStylesPreviewGenerator(oLogicDocument)
	{
		AscCommon.CActionOnTimerBase.call(this);

		this.FirstActionOnTimer = true;

		this.Api             = oLogicDocument.GetApi();
		this.LogicDocument   = oLogicDocument;
		this.DrawingDocument = oLogicDocument.GetDrawingDocument();
		this.TableStyles     = [];
		this.TableLook       = null;
		this.Index           = -1;
		this.Buffer          = [];
	}
	CTableStylesPreviewGenerator.prototype = Object.create(AscCommon.CActionOnTimerBase.prototype);
	CTableStylesPreviewGenerator.prototype.constructor = CTableStylesPreviewGenerator;
	CTableStylesPreviewGenerator.prototype.OnBegin = function(isDefaultTableLook)
	{
		this.TableStyles = this.LogicDocument.GetAllTableStyles();
		this.Index       = -1;
		this.TableLook   = this.DrawingDocument.GetTableLook(isDefaultTableLook);

		this.Api.sendEvent("asc_onBeginTableStylesPreview", this.TableStyles.length);
	};
	CTableStylesPreviewGenerator.prototype.OnEnd = function()
	{
		this.Api.sendEvent("asc_onEndTableStylesPreview");
	};
	CTableStylesPreviewGenerator.prototype.IsContinue = function()
	{
		return (this.Index < this.TableStyles.length);
	};
	CTableStylesPreviewGenerator.prototype.DoAction = function()
	{
		let oPreview = this.GetPreview(this.TableStyles[this.Index]);
		if (oPreview)
			this.Buffer.push(oPreview);

		this.Index++;
	};
	CTableStylesPreviewGenerator.prototype.OnEndTimer = function()
	{
		this.Api.sendEvent("asc_onAddTableStylesPreview", this.Buffer);
		this.Buffer = [];
	};
	CTableStylesPreviewGenerator.prototype.GetAllPreviews = function(isDefaultTableLook)
	{
		let oTableLookOld = this.TableLook;
		this.TableLook    = this.DrawingDocument.GetTableLook(isDefaultTableLook);

		let arrStyles = this.LogicDocument.GetAllTableStyles();

		let arrPreviews = [];
		for (let nIndex = 0 , nCount = arrStyles.length; nIndex < nCount; ++nIndex)
		{
			let oPreview = this.GetPreview(arrStyles[nIndex]);
			if (oPreview)
				arrPreviews.push(oPreview);
		}

		this.TableLook = oTableLookOld;

		return arrPreviews;
	};
	CTableStylesPreviewGenerator.prototype.GetAllPreviewsNative = function(isDefaultTableLook, oGraphics, oStream, oNative, dW, dH, nW, nH)
	{
		let oTableLookOld = this.TableLook;
		this.TableLook    = this.DrawingDocument.GetTableLook(isDefaultTableLook);
		let arrStyles = this.LogicDocument.GetAllTableStyles();
		oNative["DD_PrepareNativeDraw"]();
		for (let nIndex = 0 , nCount = arrStyles.length; nIndex < nCount; ++nIndex)
		{
			let oStyle = arrStyles[nIndex];
			let oTable = this.GetTable(oStyle);
			oNative["DD_StartNativeDraw"](nW, nH, dW, dH);
			this.DrawTable(oGraphics, oTable);
			oStream["ClearNoAttack"]();
			oStream["WriteByte"](2);
			oStream["WriteString2"]("" + oStyle.GetId());
			oNative["DD_EndNativeDraw"](oStream);
			oGraphics.ClearParams();
		}
		oStream["ClearNoAttack"]();
		oStream["WriteByte"](3);
		oNative["DD_EndNativeDraw"](oStream);
		this.TableLook = oTableLookOld;
	};
	CTableStylesPreviewGenerator.prototype.GetPreview = function(oStyle)
	{
		if (!oStyle)
			return null;

		var _pageW = 297;
		var _pageH = 210;
		var _canvas = GetCanvas();
		var ctx = _canvas.getContext('2d');
		ctx.fillStyle = "#FFFFFF";
		ctx.fillRect(0, 0, _canvas.width, _canvas.height);
		var graphics = new AscCommon.CGraphics();
		graphics.init(ctx, _canvas.width, _canvas.height, _pageW, _pageH);
		graphics.m_oFontManager = AscCommon.g_fontManager;
		graphics.transform(1, 0, 0, 1, 0, 0);

		let oTable = this.GetTable(oStyle);
		this.DrawTable(graphics, oTable);

		var _styleD = new AscCommon.CStyleImage();
		_styleD.type = AscCommon.c_oAscStyleImage.Default;
		_styleD.image = _canvas.toDataURL("image/png");
		_styleD.name = oStyle.GetId();
		_styleD.displayName = oStyle.GetName();

		return _styleD;
	};
	CTableStylesPreviewGenerator.prototype.GetTable = function(oStyle)
	{
		let oTable     = GetTable(this.LogicDocument);
		let oTableLook = this.TableLook;
		AscCommon.ExecuteNoHistory(function(){
			oTable.Set_Props({
				TableStyle : oStyle.GetId(),
				TableLook  : oTableLook,
				CellSelect : false
			});
			oTable.Recalc_CompiledPr2();
		}, this.LogicDocument);

		return oTable;
	};
	CTableStylesPreviewGenerator.prototype.DrawTable = function(oGraphics, oTable)
	{
		oTable.Recalculate_Page(0);
		let _old_mode = editor.isViewMode;
		editor.isViewMode = true;
		editor.isShowTableEmptyLineAttack = this.LogicDocument.IsDocumentEditor();
		oTable.Draw(0, oGraphics, false);
		editor.isShowTableEmptyLineAttack = false;
		editor.isViewMode = _old_mode;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].CTableStylesPreviewGenerator = CTableStylesPreviewGenerator;

})(window);

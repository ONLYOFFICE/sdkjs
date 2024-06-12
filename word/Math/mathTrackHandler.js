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

(function(window)
{
	/**
	 * Класс контролирует события работы трека формулы. Вызывать у этого класса события обновления можно
	 * сколько угодно раз, а этот класс уже отрисовщику и в интерфейс посылает события, только когда реально что-то
	 * изменилось
	 *
	 * @constructor
	 */
	function CMathTrackHandler(drawingDocument, eventHandler)
	{
		this.DrawingDocument = drawingDocument;
		this.EventHandler    = eventHandler;

		this.Math        = null;
		this.PageNum     = -1;
		this.ForceUpdate = true;
	}

	CMathTrackHandler.prototype.SetTrackObject = function(math, pageNum, isActive)
	{
		// TODO: Сейчас посылаем сообщение в отрисовщик трека по старому всегда

		if (math)
			this.DrawingDocument.Update_MathTrack(true, isActive, math);
		else
			this.DrawingDocument.Update_MathTrack(false);

		if (math !== this.Math
			|| (math && (this.PageNum !== pageNum || this.ForceUpdate)))
		{
			this.Math        = math;
			this.ForceUpdate = false;
			this.PageNum     = pageNum;

			let bounds = null;
			if (this.Math)
				bounds = this.GetBounds();

			if (bounds)
			{
				this.OnShow(bounds);
			}
			else
			{
				this.Math    = null;
				this.PageNum = -1;

				this.OnHide();
			}
		}
	};
	CMathTrackHandler.prototype.Update = function()
	{
		this.ForceUpdate = true;
	};
	CMathTrackHandler.prototype.OnChangePosition = function()
	{
		let bounds = this.GetBounds();
		if (!bounds)
			return;

		this.OnShow(bounds);
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CMathTrackHandler.prototype.GetBounds = function()
	{
		let math = this.Math;

		if (!math)
			return null;

		let paragraph  = math.GetParagraph();
		let mathBounds = math.GetBounds();

		let oTextTransform = paragraph.Get_ParentTextTransform();
		if (!mathBounds || !mathBounds.length || !paragraph)
			return null;

		let firstBounds = null;
		for (let index = 0, count = mathBounds.length; index < count; ++index)
		{
			for (let innerIndex = 0, innerCount = mathBounds[index].length; innerIndex < innerCount; ++innerIndex)
			{
				let bounds = mathBounds[index][innerIndex];

				if (bounds.W < 0.001 || bounds.H < 0.001)
					continue;
				
				if (!firstBounds)
					firstBounds = bounds;

				if (this.PageNum === bounds.Page)
				{
					firstBounds = bounds;
					break;
				}
			}
		}
		
		if (!firstBounds)
		{
			if (!math.IsEmpty() && mathBounds.length > 0 && mathBounds[0].length > 0)
			{
				let logicDocument = paragraph.GetLogicDocument();
				let shift         = logicDocument ? logicDocument.GetDrawingDocument().GetMMPerDot(5) : 0.1;

				let tmpBounds = mathBounds[0][0];
				firstBounds = {
					Page : tmpBounds.Page,
					X    : tmpBounds.X,
					Y    : tmpBounds.Y,
					W    : Math.max(tmpBounds.W, shift),
					H    : Math.max(tmpBounds.H, shift)
				};
			}
			else
			{
				return null;
			}
		}
		
		let pageNum = firstBounds.Page;
		let x0 = firstBounds.X;
		let y0 = firstBounds.Y;
		let x1 = firstBounds.X + firstBounds.W;
		let y1 = firstBounds.Y + firstBounds.H;

		for (let index = 0, count = mathBounds.length; index < count; ++index)
		{
			for (let innerIndex = 0, innerCount = mathBounds[index].length; innerIndex < innerCount; ++innerIndex)
			{
				let bounds = mathBounds[index][innerIndex];
				if (bounds.Page !== pageNum
					|| bounds.W < 0.001
					|| bounds.H < 0.001)
					continue;

				if (x0 > bounds.X)
					x0 = bounds.X;

				if (x1 < bounds.X + bounds.W)
					x1 = bounds.X + bounds.W;

				if (y0 > bounds.Y)
					y0 = bounds.Y;

				if (y1 < bounds.Y + bounds.H)
					y1 = bounds.Y + bounds.H;
			}
		}
		if(oTextTransform)
		{
			let aX = [];
			let aY = [];
			aX.push(oTextTransform.TransformPointX(x0, y0));
			aX.push(oTextTransform.TransformPointX(x0, y1));
			aX.push(oTextTransform.TransformPointX(x1, y0));
			aX.push(oTextTransform.TransformPointX(x1, y1));
			aY.push(oTextTransform.TransformPointY(x0, y0));
			aY.push(oTextTransform.TransformPointY(x0, y1));
			aY.push(oTextTransform.TransformPointY(x1, y0));
			aY.push(oTextTransform.TransformPointY(x1, y1));
			x0 = Math.min.apply(Math, aX);
			y0 = Math.min.apply(Math, aY);
			x1 = Math.max.apply(Math, aX);
			y1 = Math.max.apply(Math, aY);
		}

		let pos0 = {};
		let pos1 = {};

		if (Asc.editor.isPdfEditor())
		{
			let oDoc	= Asc.editor.getPDFDoc();
			let oFile	= Asc.editor.getDocumentRenderer().file;

			let oTr		= oDoc.pagesTransform[pageNum].normal.CreateDublicate();
			let inchC	= (25.4 / oFile.pages[pageNum].Dpi);
			AscCommon.global_MatrixTransformer.ScaleAppend(oTr, inchC, inchC);
			oTr.Invert();

			let oPt1 = oTr.TransformPoint(x0, y0);
			let oPt2 = oTr.TransformPoint(x1, y1);

			pos0 = {
				X: Math.min(oPt1.x, oPt2.x),
				Y: Math.min(oPt1.y, oPt2.y)
			};
			pos1 = {
				X: Math.max(oPt1.x, oPt2.x),
				Y: Math.max(oPt1.y, oPt2.y)
			};
		}
		else
		{
			pos0 = this.DrawingDocument.ConvertCoordsToCursorWR(x0, y0, pageNum);
			pos1 = this.DrawingDocument.ConvertCoordsToCursorWR(x1, y1, pageNum);
		}

		return [pos0.X, pos0.Y, pos1.X, pos1.Y];
	};
	CMathTrackHandler.prototype.OnHide = function()
	{
		this.EventHandler.sendEvent("asc_onHideMathTrack");
	};
	CMathTrackHandler.prototype.OnShow = function(bounds)
	{
		this.EventHandler.sendEvent("asc_onShowMathTrack", bounds);
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CMathTrackHandler = CMathTrackHandler;

})(window);

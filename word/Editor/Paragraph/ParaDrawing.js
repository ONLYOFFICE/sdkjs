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

var drawing_Inline = 0x01;
var drawing_Anchor = 0x02;

var WRAPPING_TYPE_NONE           = 0x00;
var WRAPPING_TYPE_SQUARE         = 0x01;
var WRAPPING_TYPE_THROUGH        = 0x02;
var WRAPPING_TYPE_TIGHT          = 0x03;
var WRAPPING_TYPE_TOP_AND_BOTTOM = 0x04;

var WRAP_HIT_TYPE_POINT   = 0x00;
var WRAP_HIT_TYPE_SECTION = 0x01;
var c_oAscAlignH         = Asc.c_oAscAlignH;
var c_oAscAlignV         = Asc.c_oAscAlignV;

/**
 * Оберточный класс для автофигур и картинок. Именно он непосредственно лежит в ране.
 * @constructor
 * @extends {AscWord.CRunElementBase}
 */
function ParaDrawing(W, H, GraphicObj, DrawingDocument, DocumentContent, Parent)
{
	AscWord.CRunElementBase.call(this);

	this.Id          = AscCommon.g_oIdCounter.Get_NewId();
	this.DrawingType = drawing_Inline;
	this.GraphicObj  = GraphicObj;

	this.Form = false; // Флаг, означающий, является ли данная автофигура контейнером для специальной формой

	this.X      = 0;
	this.Y      = 0;
	this.Width  = 0;
	this.Height = 0;
	this.OrigX = 0;
	this.OrigY = 0;
	this.ShiftX = 0;
	this.ShiftY = 0;

	this.PageNum = 0;
	this.LineNum = 0;
	this.YOffset = 0;

	this.DocumentContent = DocumentContent;
	this.DrawingDocument = DrawingDocument;
	this.Parent          = Parent;
	this.LogicDocument   = DrawingDocument ? DrawingDocument.m_oLogicDocument : null;
	
	if (!this.LogicDocument && Asc.editor)
		this.LogicDocument = Asc.editor.getLogicDocument();
	
	// Расстояние до окружающего текста
	this.Distance = {
		T : 0,
		B : 0,
		L : 0,
		R : 0
	};

	// Расположение в таблице
	this.LayoutInCell = true;

	// Z-order
	this.RelativeHeight = undefined;

	//
	this.SimplePos = {
		Use : false,
		X   : 0,
		Y   : 0
	};

	// Ширина и высота
	this.Extent = {
		W : W,
		H : H
	};

	this.EffectExtent = {
		L : 0,
		T : 0,
		R : 0,
		B : 0
	};

	this.docPr = new AscFormat.CNvPr();

	this.SizeRelH = undefined;
	this.SizeRelV = undefined;
	//{RelativeFrom      : c_oAscRelativeFromH.Column, Percent: ST_PositivePercentage}

	this.AllowOverlap = true;

	//привязка к параграфу
	this.Locked = false;

	//скрытые drawing'и
	this.Hidden = null;

	// Позиция по горизонтали
	this.PositionH = {
		RelativeFrom : c_oAscRelativeFromH.Column, // Относительно чего вычисляем координаты
		Align        : false,                      // true : В поле Value лежит тип прилегания, false - в поле Value
												   // лежит точное значени
		Value        : 0,                          //
		Percent      : false                       // Значение Valuе задано в процентах
	};

	// Позиция по горизонтали
	this.PositionV = {
		RelativeFrom : c_oAscRelativeFromV.Paragraph, // Относительно чего вычисляем координаты
		Align        : false,                         // true : В поле Value лежит тип прилегания, false - в поле Value
													  // лежит точное значени
		Value        : 0,                             //
		Percent      : false                          // Значение Valuе задано в процентах
	};

	// Данный поля используются для перемещения Flow-объекта
	this.PositionH_Old = undefined;
	this.PositionV_Old = undefined;

	this.Internal_Position = new CAnchorPosition();

	// Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
	//--------------------------------------------------------
	this.wrappingType = WRAPPING_TYPE_THROUGH;
	this.useWrap      = true;

	if (typeof CWrapPolygon !== "undefined")
		this.wrappingPolygon = new CWrapPolygon(this);

	this.document        = this.LogicDocument;
	this.drawingDocument = DrawingDocument;
	this.graphicObjects  = this.LogicDocument ? this.LogicDocument.getDrawingObjects() : null;
	
	this.selected        = false;

	this.behindDoc    = false;
	this.bNoNeedToAdd = false;

	this.pageIndex = -1;
	this.Lock      = new AscCommon.CLock();

	this.ParaMath = null;

	this.SkipOnRecalculate = false;
	//for clip
	this.LineTop = null;
	this.LineBottom = null;
	//------------------------------------------------------------
	AscCommon.g_oTableId.Add(this, this.Id);

	if (this.graphicObjects)
	{
		this.Set_RelativeHeight(this.graphicObjects.getZIndex());
		if (History.CanRegisterClasses())
			this.graphicObjects.addGraphicObject(this);
	}
	this.bCellHF = false;
}
ParaDrawing.prototype = Object.create(AscWord.CRunElementBase.prototype);
ParaDrawing.prototype.constructor = ParaDrawing;

ParaDrawing.prototype.Type = para_Drawing;
ParaDrawing.prototype.SetParent = function(oParent)
{
	if (!oParent)
		return;

	if (oParent instanceof ParaRun)
		this.Parent = oParent.GetParagraph();
	else if (oParent instanceof Paragraph)
		this.Parent = oParent;
};
ParaDrawing.prototype.GetParent = function()
{
	return this.Parent;
};
ParaDrawing.prototype.Get_Type = function()
{
	return this.Type;
};
ParaDrawing.prototype.GetWidth = function()
{
	return this.Width * this.GetScaleCoefficient();
};
ParaDrawing.prototype.GetInlineWidth = function()
{
	if (!this.IsInline())
		return 0;

	return this.Width;
};
ParaDrawing.prototype.Get_Height = function()
{
	return this.Height * this.GetScaleCoefficient();
};
ParaDrawing.prototype.GetHeight = function()
{
	return this.Get_Height();
};
ParaDrawing.prototype.getHeight = function()
{
	return this.Get_Height();
};
ParaDrawing.prototype.GetWidthVisible = function()
{
	return this.WidthVisible * this.GetScaleCoefficient();
};
ParaDrawing.prototype.SetWidthVisible = function(WidthVisible)
{
	this.WidthVisible = WidthVisible;
};
ParaDrawing.prototype.GetSelectedContent = function(SelectedContent)
{
	if (this.GraphicObj && this.GraphicObj.GetSelectedContent)
	{
		this.GraphicObj.GetSelectedContent(SelectedContent);
	}
};
ParaDrawing.prototype.GetSearchElementId = function(bNext, bCurrent)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.GetSearchElementId === "function")
		return this.GraphicObj.GetSearchElementId(bNext, bCurrent);

	return null;
};
ParaDrawing.prototype.FindNextFillingForm = function(isNext, isCurrent)
{
	if (isCurrent && this.IsForm())
		return null;
	
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.FindNextFillingForm === "function")
		return this.GraphicObj.FindNextFillingForm(isNext, isCurrent);

	return null;
};
ParaDrawing.prototype.CheckCorrect = function()
{
	if (!this.GraphicObj)
	{
		return false;
	}
	if (this.GraphicObj && this.GraphicObj.checkCorrect)
	{
		return this.GraphicObj.checkCorrect();
	}
	return true;
};
ParaDrawing.prototype.GetAllDrawingObjects = function(arrDrawingObjects)
{
	if (!arrDrawingObjects)
		arrDrawingObjects = [];

	if (this.GraphicObj.GetAllDrawingObjects)
		this.GraphicObj.GetAllDrawingObjects(arrDrawingObjects);

	return arrDrawingObjects;
};
ParaDrawing.prototype.GetAllOleObjects = function(sPluginId, arrObjects)
{
	if (!Array.isArray(arrObjects))
	{
		arrObjects = [];
	}

	if (this.GraphicObj.GetAllOleObjects)
		this.GraphicObj.GetAllOleObjects(sPluginId, arrObjects);

	return arrObjects;
};
ParaDrawing.prototype.canRotate = function()
{
	return AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.canRotate == "function" && this.GraphicObj.canRotate();
};
ParaDrawing.prototype.canResize = function()
{
	return AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.canResize == "function" && this.GraphicObj.canResize();
};
ParaDrawing.prototype.GetParagraph = function()
{
	return this.Get_ParentParagraph();
};
ParaDrawing.prototype.GetRun = function()
{
	return this.Get_Run();
};
ParaDrawing.prototype.GetDocumentContent = function()
{
	// TODO: Check why do we need to skip BlockLevelSdt here. If it's not necessary then merge this method with GetParentDocumentContent
	const oParagraph = this.GetParagraph();
	let oDocumentContent = (oParagraph ? oParagraph.GetParent() : null);
	if (oDocumentContent && oDocumentContent.IsBlockLevelSdtContent())
	{
		oDocumentContent = oDocumentContent.Parent.Parent;
	}
	return oDocumentContent;
};
ParaDrawing.prototype.GetParentDocumentContent = function()
{
	let para = this.GetParagraph();
	return para ? para.GetParent() : null;
};
ParaDrawing.prototype.Get_Run = function()
{
	var oParagraph = this.Get_ParentParagraph();
	if (oParagraph)
		return oParagraph.Get_DrawingObjectRun(this.Id);

	return null;
};
ParaDrawing.prototype.Get_Props = function(OtherProps)
{
	var Props    = {};
	Props.Width  = this.GraphicObj.extX;
	Props.Height = this.GraphicObj.extY;
	Props.Form   = this.IsForm();
	if (drawing_Inline === this.DrawingType)
		Props.WrappingStyle = c_oAscWrapStyle2.Inline;
	else if (WRAPPING_TYPE_NONE === this.wrappingType)
		Props.WrappingStyle = ( this.behindDoc === true ? c_oAscWrapStyle2.Behind : c_oAscWrapStyle2.InFront );
	else
	{
		switch (this.wrappingType)
		{
			case WRAPPING_TYPE_SQUARE         :
				Props.WrappingStyle = c_oAscWrapStyle2.Square;
				break;
			case WRAPPING_TYPE_TIGHT          :
				Props.WrappingStyle = c_oAscWrapStyle2.Tight;
				break;
			case WRAPPING_TYPE_THROUGH        :
				Props.WrappingStyle = c_oAscWrapStyle2.Through;
				break;
			case WRAPPING_TYPE_TOP_AND_BOTTOM :
				Props.WrappingStyle = c_oAscWrapStyle2.TopAndBottom;
				break;
			default                           :
				Props.WrappingStyle = c_oAscWrapStyle2.Inline;
				break;
		}
	}

	if (drawing_Inline === this.DrawingType)
	{
		Props.Paddings =
		{
			Left   : AscFormat.DISTANCE_TO_TEXT_LEFTRIGHT,
			Right  : AscFormat.DISTANCE_TO_TEXT_LEFTRIGHT,
			Top    : 0,
			Bottom : 0
		};
	}
	else
	{
		var oDistance  = this.Get_Distance();
		Props.Paddings =
		{
			Left   : oDistance.L,
			Right  : oDistance.R,
			Top    : oDistance.T,
			Bottom : oDistance.B
		};
	}

	Props.AllowOverlap = this.AllowOverlap;

	Props.Position =
	{
		X : this.X,
		Y : this.Y
	};

	Props.PositionH =
	{
		RelativeFrom : this.PositionH.RelativeFrom,
		UseAlign     : this.PositionH.Align,
		Align        : ( true === this.PositionH.Align ? this.PositionH.Value : undefined ),
		Value        : ( true === this.PositionH.Align ? 0 : this.PositionH.Value ),
		Percent      : this.PositionH.Percent
	};

	Props.PositionV =
	{
		RelativeFrom : this.PositionV.RelativeFrom,
		UseAlign     : this.PositionV.Align,
		Align        : ( true === this.PositionV.Align ? this.PositionV.Value : undefined ),
		Value        : ( true === this.PositionV.Align ? 0 : this.PositionV.Value ),
		Percent      : this.PositionV.Percent
	};


	if (this.SizeRelH && this.SizeRelH.Percent > 0)
	{
		Props.SizeRelH =
		{
			RelativeFrom : AscFormat.ConvertRelSizeHToRelPosition(this.SizeRelH.RelativeFrom),
			Value        : (this.SizeRelH.Percent * 100) >> 0
		};
	}

	if (this.SizeRelV && this.SizeRelV.Percent > 0)
	{
		Props.SizeRelV =
		{
			RelativeFrom : AscFormat.ConvertRelSizeVToRelPosition(this.SizeRelV.RelativeFrom),
			Value        : (this.SizeRelV.Percent * 100) >> 0
		};
	}

	Props.Internal_Position = this.Internal_Position;

	Props.Locked        = this.Lock.Is_Locked();
	var ParentParagraph = this.Get_ParentParagraph();
	if (ParentParagraph && undefined !== ParentParagraph.Parent)
	{
		var DocContent = ParentParagraph.Parent;
		if (true === DocContent.Is_DrawingShape() || (DocContent.GetTopDocumentContent() instanceof CFootEndnote))
			Props.CanBeFlow = false;
	}

    Props.title = this.docPr.title !== null ? this.docPr.title : undefined;
    Props.description = this.docPr.descr !== null ? this.docPr.descr : undefined;

	if (null != OtherProps && undefined != OtherProps)
	{
		// Соединяем
		if (undefined === OtherProps.Width || 0.001 > Math.abs(Props.Width - OtherProps.Width))
			Props.Width = undefined;

		if (undefined === OtherProps.Height || 0.001 > Math.abs(Props.Height - OtherProps.Height))
			Props.Height = undefined;

		if (undefined === OtherProps.WrappingStyle || Props.WrappingStyle != OtherProps.WrappingStyle)
			Props.WrappingStyle = undefined;

		if (undefined === OtherProps.ImageUrl || Props.ImageUrl != OtherProps.ImageUrl)
			Props.ImageUrl = undefined;

		if (undefined === OtherProps.Paddings.Left || 0.001 > Math.abs(Props.Paddings.Left - OtherProps.Paddings.Left))
			Props.Paddings.Left = undefined;

		if (undefined === OtherProps.Paddings.Right || 0.001 > Math.abs(Props.Paddings.Right - OtherProps.Paddings.Right))
			Props.Paddings.Right = undefined;

		if (undefined === OtherProps.Paddings.Top || 0.001 > Math.abs(Props.Paddings.Top - OtherProps.Paddings.Top))
			Props.Paddings.Top = undefined;

		if (undefined === OtherProps.Paddings.Bottom || 0.001 > Math.abs(Props.Paddings.Bottom - OtherProps.Paddings.Bottom))
			Props.Paddings.Bottom = undefined;

		if (undefined === OtherProps.AllowOverlap || Props.AllowOverlap != OtherProps.AllowOverlap)
			Props.AllowOverlap = undefined;

		if (undefined === OtherProps.Position.X || 0.001 > Math.abs(Props.Position.X - OtherProps.Position.X))
			Props.Position.X = undefined;

		if (undefined === OtherProps.Position.Y || 0.001 > Math.abs(Props.Position.Y - OtherProps.Position.Y))
			Props.Position.Y = undefined;

		if (undefined === OtherProps.PositionH.RelativeFrom || Props.PositionH.RelativeFrom != OtherProps.PositionH.RelativeFrom)
			Props.PositionH.RelativeFrom = undefined;

		if (undefined === OtherProps.PositionH.UseAlign || Props.PositionH.UseAlign != OtherProps.PositionH.UseAlign)
			Props.PositionH.UseAlign = undefined;

		if (Props.PositionH.RelativeFrom === OtherProps.PositionH.RelativeFrom && Props.PositionH.UseAlign === OtherProps.PositionH.UseAlign)
		{
			if (true != Props.PositionH.UseAlign && 0.001 > Math.abs(Props.PositionH.Value - OtherProps.PositionH.Value))
				Props.PositionH.Value = undefined;

			if (true === Props.PositionH.UseAlign && Props.PositionH.Align != OtherProps.PositionH.Align)
				Props.PositionH.Align = undefined;
		}

		if (undefined === OtherProps.PositionV.RelativeFrom || Props.PositionV.RelativeFrom != OtherProps.PositionV.RelativeFrom)
			Props.PositionV.RelativeFrom = undefined;

		if (undefined === OtherProps.PositionV.UseAlign || Props.PositionV.UseAlign != OtherProps.PositionV.UseAlign)
			Props.PositionV.UseAlign = undefined;

		if (Props.PositionV.RelativeFrom === OtherProps.PositionV.RelativeFrom && Props.PositionV.UseAlign === OtherProps.PositionV.UseAlign)
		{
			if (true != Props.PositionV.UseAlign && 0.001 > Math.abs(Props.PositionV.Value - OtherProps.PositionV.Value))
				Props.PositionV.Value = undefined;

			if (true === Props.PositionV.UseAlign && Props.PositionV.Align != OtherProps.PositionV.Align)
				Props.PositionV.Align = undefined;
		}


		if (false === OtherProps.Locked)
			Props.Locked = false;

		if (false === OtherProps.CanBeFlow || false === Props.CanBeFlow)
			Props.CanBeFlow = false;
		else
			Props.CanBeFlow = true;

		if(undefined === OtherProps.title || Props.title !== OtherProps.title){
			Props.title = undefined;
		}

		if(undefined === OtherProps.description || Props.description !== OtherProps.description){
			Props.description = undefined;
		}
		
		if (OtherProps.Form || Props.Form)
			Props.Form = true;
	}

	return Props;
};
ParaDrawing.prototype.IsUseInDocument = function()
{
	if (this.Parent)
	{
		var Run = this.Parent.Get_DrawingObjectRun(this.Id);
		if (Run)
		{
			return Run.IsUseInDocument(this.GetId());
		}
	}
	return false;
};
ParaDrawing.prototype.CheckGroupSizes = function()
{
	if (this.GraphicObj && this.GraphicObj.CheckGroupSizes)
	{
		this.GraphicObj.CheckGroupSizes();
	}
};
ParaDrawing.prototype.Set_DrawingType = function(DrawingType)
{
	History.Add(new CChangesParaDrawingDrawingType(this, this.DrawingType, DrawingType));
	this.DrawingType = DrawingType;
};
ParaDrawing.prototype.Set_WrappingType = function(WrapType)
{
	History.Add(new CChangesParaDrawingWrappingType(this, this.wrappingType, WrapType));
	this.wrappingType = WrapType;
};
ParaDrawing.prototype.Set_Distance = function(L, T, R, B)
{
	var oDistance = this.Get_Distance();
	if (!AscFormat.isRealNumber(L))
	{
		L = oDistance.L;
	}
	if (!AscFormat.isRealNumber(T))
	{
		T = oDistance.T;
	}
	if (!AscFormat.isRealNumber(R))
	{
		R = oDistance.R;
	}
	if (!AscFormat.isRealNumber(B))
	{
		B = oDistance.B;
	}

	History.Add(new CChangesParaDrawingDistance(this,
		{
			Left   : this.Distance.L,
			Top    : this.Distance.T,
			Right  : this.Distance.R,
			Bottom : this.Distance.B
		},
		{
			Left   : L,
			Top    : T,
			Right  : R,
			Bottom : B
		}));

	this.Distance.L = L;
	this.Distance.R = R;
	this.Distance.T = T;
	this.Distance.B = B;
};
ParaDrawing.prototype.Set_AllowOverlap = function(AllowOverlap)
{
	History.Add(new CChangesParaDrawingAllowOverlap(this, this.AllowOverlap, AllowOverlap));
	this.AllowOverlap = AllowOverlap;
};
ParaDrawing.prototype.Set_PositionH = function(RelativeFrom, Align, Value, Percent)
{
    var _Value, _Percent;
    if(AscFormat.isRealNumber(Value) && AscFormat.fApproxEqual(Value, 0.0) && true === Percent)
    {
        _Value        = 0;
        _Percent      = false;
    }
    else
    {
        _Value        = Value;
        _Percent      = Percent;
    }
	History.Add(new CChangesParaDrawingPositionH(this,
		{
			RelativeFrom : this.PositionH.RelativeFrom,
			Align        : this.PositionH.Align,
			Value        : this.PositionH.Value,
			Percent      : this.PositionH.Percent
		},
		{
			RelativeFrom : RelativeFrom,
			Align        : Align,
			Value        : _Value,
			Percent      : _Percent
		}));
	this.PositionH.RelativeFrom = RelativeFrom;
	this.PositionH.Align        = Align;
    this.PositionH.Value        = _Value;
    this.PositionH.Percent      = _Percent;
};
ParaDrawing.prototype.Set_PositionV = function(RelativeFrom, Align, Value, Percent)
{
    var _Value, _Percent;
    if(AscFormat.isRealNumber(Value) && AscFormat.fApproxEqual(Value, 0.0) && true === Percent)
    {
        _Value        = 0;
        _Percent      = false;
    }
    else
    {
        _Value        = Value;
        _Percent      = Percent;
    }
	History.Add(new CChangesParaDrawingPositionV(this,
		{
			RelativeFrom : this.PositionV.RelativeFrom,
			Align        : this.PositionV.Align,
			Value        : this.PositionV.Value,
			Percent      : this.PositionV.Percent
		},
		{
			RelativeFrom : RelativeFrom,
			Align        : Align,
			Value        : _Value,
			Percent      : _Percent
		}));

	this.PositionV.RelativeFrom = RelativeFrom;
	this.PositionV.Align        = Align;
    this.PositionV.Value        = _Value;
    this.PositionV.Percent      = _Percent;
};
ParaDrawing.prototype.GetPositionH = function()
{
	return {
		RelativeFrom : this.PositionH.RelativeFrom,
		Align        : this.PositionH.Align,
		Value        : this.PositionH.Value,
		Percent      : this.PositionH.Percent
	};
};
ParaDrawing.prototype.GetPositionV = function()
{
	return {
		RelativeFrom : this.PositionV.RelativeFrom,
		Align        : this.PositionV.Align,
		Value        : this.PositionV.Value,
		Percent      : this.PositionV.Percent
	};
};
ParaDrawing.prototype.Set_BehindDoc = function(BehindDoc)
{
	History.Add(new CChangesParaDrawingBehindDoc(this, this.behindDoc, BehindDoc));
	this.behindDoc = BehindDoc;
};
ParaDrawing.prototype.Set_GraphicObject = function(graphicObject)
{
	var oldId = AscCommon.isRealObject(this.GraphicObj) ? this.GraphicObj.Get_Id() : null;
	var newId = AscCommon.isRealObject(graphicObject) ? graphicObject.Get_Id() : null;

	History.Add(new CChangesParaDrawingGraphicObject(this, oldId, newId));

	if (graphicObject)
		graphicObject.handleUpdateExtents();

	this.GraphicObj = graphicObject;
};
ParaDrawing.prototype.setSimplePos = function(use, x, y)
{
	History.Add(new CChangesParaDrawingSimplePos(this,
		{
			Use : this.SimplePos.Use,
			X   : this.SimplePos.X,
			Y   : this.SimplePos.Y
		},
		{
			Use : use,
			X   : x,
			Y   : y
		}));

	this.SimplePos.Use = use;
	this.SimplePos.X   = x;
	this.SimplePos.Y   = y;
};
ParaDrawing.prototype.setExtent = function(extX, extY)
{
	History.Add(new CChangesParaDrawingExtent(this,
		{
			W : this.Extent.W,
			H : this.Extent.H
		},
		{
			W : extX,
			H : extY
		}));

	this.Extent.W = extX;
	this.Extent.H = extY;
};
ParaDrawing.prototype.addWrapPolygon = function(wrapPolygon)
{
	History.Add(new CChangesParaDrawingWrapPolygon(this, this.wrappingPolygon, wrapPolygon));
	this.wrappingPolygon = wrapPolygon;
};
ParaDrawing.prototype.Set_Locked = function(bLocked)
{
	History.Add(new CChangesParaDrawingLocked(this, this.Locked, bLocked));
	this.Locked = bLocked;
};
ParaDrawing.prototype.Set_RelativeHeight = function(nRelativeHeight)
{
	History.Add(new CChangesParaDrawingRelativeHeight(this, this.RelativeHeight, nRelativeHeight));
	this.Set_RelativeHeight2(nRelativeHeight);
};
ParaDrawing.prototype.Set_RelativeHeight2 = function(nRelativeHeight)
{
	this.RelativeHeight = nRelativeHeight;
	if (this.graphicObjects && AscFormat.isRealNumber(nRelativeHeight) && nRelativeHeight > this.graphicObjects.maximalGraphicObjectZIndex)
	{
		this.graphicObjects.maximalGraphicObjectZIndex = nRelativeHeight;
	}
};
ParaDrawing.prototype.setEffectExtent = function(L, T, R, B)
{
	var oEE = this.EffectExtent;
	History.Add(new CChangesParaDrawingEffectExtent(this,
		{
			L : oEE.L,
			T : oEE.T,
			R : oEE.R,
			B : oEE.B
		},
		{
			L : L,
			T : T,
			R : R,
			B : B
		}));

	this.EffectExtent.L = L;
	this.EffectExtent.T = T;
	this.EffectExtent.R = R;
	this.EffectExtent.B = B;
};
ParaDrawing.prototype.Set_Parent = function(oParent)
{
	History.Add(new CChangesParaDrawingParent(this, this.Parent, oParent));
	this.Parent = oParent;
};
ParaDrawing.prototype.IsWatermark = function()
{
	if(!this.GraphicObj)
	{
		return false;
	}
	if(this.GraphicObj.getObjectType() !== AscDFH.historyitem_type_Shape && this.GraphicObj.getObjectType() !== AscDFH.historyitem_type_ImageShape)
	{
		return false;
	}
	if(this.Is_Inline())
	{
		return false;
	}
	var oParagraph = this.GetParagraph();
	if(!(oParagraph instanceof Paragraph))
	{
		return false;
	}
	var oContent = oParagraph.Parent;
	if(!oContent || oContent.Is_DrawingShape(false))
	{
		return false;
	}
	var oHdrFtr = oContent.IsHdrFtr(true);
	if(!oHdrFtr)
	{
		return false;
	}

	var oRun = this.Get_Run();
	if (!oRun)
		return false;

	var arrDocPos = oRun.GetDocumentPositionFromObject();
	for (var nIndex = 0, nCount = arrDocPos.length; nIndex < nCount; ++nIndex)
	{
		var oClass = arrDocPos[nIndex].Class;
		var oSdt   = null;
		if (oClass instanceof CDocumentContent && oClass.Parent instanceof CBlockLevelSdt)
			oSdt = oClass.Parent;
		else if (oClass instanceof CInlineLevelSdt)
			oSdt = oClass;

		if (oSdt)
		{
			var oPr = oSdt.Pr;
			if (AscCommon.isRealObject(oPr) && AscCommon.isRealObject(oPr.DocPartObj) && oPr.DocPartObj.Gallery === "Watermarks")
				return true;
		}
	}
	if(this.docPr) 
	{
		if(this.docPr.name && 
			this.docPr.name.indexOf("PowerPlusWaterMarkObject") > -1) {
				return true;
		}
	}
	return false;
};
ParaDrawing.prototype.Set_ParaMath = function(ParaMath)
{
	History.Add(new CChangesParaDrawingParaMath(this, this.ParaMath, ParaMath));
	this.ParaMath = ParaMath;
};
ParaDrawing.prototype.Set_LayoutInCell = function(LayoutInCell)
{
	if (this.LayoutInCell === LayoutInCell)
		return;

	History.Add(new CChangesParaDrawingLayoutInCell(this, this.LayoutInCell, LayoutInCell));
	this.LayoutInCell = LayoutInCell;
};
ParaDrawing.prototype.SetSizeRelH = function(oSize)
{
	History.Add(new CChangesParaDrawingSizeRelH(this, this.SizeRelH, oSize));
	this.SizeRelH = oSize;
};
ParaDrawing.prototype.SetSizeRelV  = function(oSize)
{
	History.Add(new CChangesParaDrawingSizeRelV(this, this.SizeRelV, oSize));
	this.SizeRelV = oSize;
};
ParaDrawing.prototype.getExtX = function()
{
	return this.getXfrmExtX() * this.GetScaleCoefficient();
};
ParaDrawing.prototype.getExtY = function()
{
	return this.getXfrmExtY() * this.GetScaleCoefficient();
};
ParaDrawing.prototype.getXfrmExtX = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && AscCommon.isRealObject(this.GraphicObj.spPr) && AscCommon.isRealObject(this.GraphicObj.spPr.xfrm) && AscFormat.isRealNumber(this.GraphicObj.spPr.xfrm.extX))
		return this.GraphicObj.spPr.xfrm.extX;
	if (AscFormat.isRealNumber(this.Extent.W))
		return this.Extent.W;
	return 0;
};
ParaDrawing.prototype.getXfrmExtY = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && AscCommon.isRealObject(this.GraphicObj.spPr) && AscCommon.isRealObject(this.GraphicObj.spPr.xfrm) && AscFormat.isRealNumber(this.GraphicObj.spPr.xfrm.extY))
		return this.GraphicObj.spPr.xfrm.extY;
	if (AscFormat.isRealNumber(this.Extent.H))
		return this.Extent.H;
	return 0;
};
ParaDrawing.prototype.getXfrmRot = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && AscCommon.isRealObject(this.GraphicObj.spPr) && AscCommon.isRealObject(this.GraphicObj.spPr.xfrm) && AscFormat.isRealNumber(this.GraphicObj.spPr.xfrm.rot))
		return this.GraphicObj.spPr.xfrm.rot;
	return 0;
};
ParaDrawing.prototype.Get_Bounds = function()
{
	var InsL, InsT, InsR, InsB;
	InsL = 0.0;
	InsT = 0.0;
	InsR = 0.0;
	InsB = 0.0;
	if(!this.Is_Inline())
	{
		var oDistance = this.Get_Distance();
		if(oDistance)
		{
			InsL = oDistance.L;
			InsT = oDistance.T;
			InsR = oDistance.R;
			InsB = oDistance.B;
		}
	}
	var ExtX = this.getXfrmExtX();
	var ExtY = this.getXfrmExtY();
	var Rot = this.getXfrmRot();
	var X, Y, W, H;
	if(AscFormat.checkNormalRotate(Rot))
	{
		X = this.X;
		Y = this.Y;
		W = ExtX;
		H = ExtY;
	}
	else
	{
		X = this.X + ExtX / 2.0 - ExtY / 2.0;
		Y = this.Y + ExtY / 2.0 - ExtX / 2.0;
		W = ExtY;
		H = ExtX;
	}
	return {Left : X - this.EffectExtent.L - InsL, Top : Y - this.EffectExtent.T - InsT, Bottom : Y + H + this.EffectExtent.B +  InsB, Right : X + W + this.EffectExtent.R + InsR};

};
ParaDrawing.prototype.Search = function(SearchEngine, Type)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.Search === "function")
	{
		this.GraphicObj.Search(SearchEngine, Type)
	}
};
ParaDrawing.prototype.Set_Props = function(Props)
{
	var bCheckWrapPolygon = false;

	var isPictureCC = false;

	var oRun = this.GetRun();
	if (oRun)
	{
		var arrContentControls = oRun.GetParentContentControls();
		for (var nIndex = 0, nCount = arrContentControls.length; nIndex < nCount; ++nIndex)
		{
			if (arrContentControls[nIndex].IsPicture())
			{
				isPictureCC = true;
				break;
			}
		}
	}

	if (undefined != Props.WrappingStyle && !isPictureCC)
	{
		if (drawing_Inline === this.DrawingType && c_oAscWrapStyle2.Inline != Props.WrappingStyle && undefined === Props.Paddings)
		{
			this.Set_Distance(3.2, 0, 3.2, 0);
		}

		this.Set_DrawingType(c_oAscWrapStyle2.Inline === Props.WrappingStyle ? drawing_Inline : drawing_Anchor);
		if (c_oAscWrapStyle2.Inline === Props.WrappingStyle)
		{
			if (AscCommon.isRealObject(this.GraphicObj.bounds) && AscFormat.isRealNumber(this.GraphicObj.bounds.w) && AscFormat.isRealNumber(this.GraphicObj.bounds.h))
			{
				this.CheckWH();
			}
		}
		if (c_oAscWrapStyle2.Behind === Props.WrappingStyle || c_oAscWrapStyle2.InFront === Props.WrappingStyle)
		{
			this.Set_WrappingType(WRAPPING_TYPE_NONE);
			this.Set_BehindDoc(c_oAscWrapStyle2.Behind === Props.WrappingStyle ? true : false);
		}
		else
		{
			switch (Props.WrappingStyle)
			{
				case c_oAscWrapStyle2.Square      :
					this.Set_WrappingType(WRAPPING_TYPE_SQUARE);
					break;
				case c_oAscWrapStyle2.Tight       :
				{
					bCheckWrapPolygon = true;
					this.Set_WrappingType(WRAPPING_TYPE_TIGHT);
					break;
				}
				case c_oAscWrapStyle2.Through     :
				{
					this.Set_WrappingType(WRAPPING_TYPE_THROUGH);
					bCheckWrapPolygon = true;
					break;
				}
				case c_oAscWrapStyle2.TopAndBottom:
					this.Set_WrappingType(WRAPPING_TYPE_TOP_AND_BOTTOM);
					break;
				default                           :
					this.Set_WrappingType(WRAPPING_TYPE_SQUARE);
					break;
			}

			this.Set_BehindDoc(false);
		}
	}

	if (undefined != Props.Paddings)
		this.Set_Distance(Props.Paddings.Left, Props.Paddings.Top, Props.Paddings.Right, Props.Paddings.Bottom);

	if (undefined != Props.AllowOverlap)
		this.Set_AllowOverlap(Props.AllowOverlap);

	if (undefined != Props.PositionH)
	{
		this.Set_PositionH(Props.PositionH.RelativeFrom, Props.PositionH.UseAlign, ( true === Props.PositionH.UseAlign ? Props.PositionH.Align : Props.PositionH.Value ), Props.PositionH.Percent);
	}
	if (undefined != Props.PositionV)
	{
		this.Set_PositionV(Props.PositionV.RelativeFrom, Props.PositionV.UseAlign, ( true === Props.PositionV.UseAlign ? Props.PositionV.Align : Props.PositionV.Value ), Props.PositionV.Percent);
	}
	if (undefined != Props.SizeRelH)
	{
		this.SetSizeRelH({
			RelativeFrom : AscFormat.ConvertRelPositionHToRelSize(Props.SizeRelH.RelativeFrom),
			Percent      : Props.SizeRelH.Value / 100.0
		});
	}

	if (undefined != Props.SizeRelV)
	{
		this.SetSizeRelV({
			RelativeFrom : AscFormat.ConvertRelPositionVToRelSize(Props.SizeRelV.RelativeFrom),
			Percent      : Props.SizeRelV.Value / 100.0
		});
	}

	if (this.SizeRelH && !this.SizeRelV)
	{
		this.SetSizeRelV({RelativeFrom : AscCommon.c_oAscSizeRelFromV.sizerelfromvPage, Percent : 0});
	}

	if (this.SizeRelV && !this.SizeRelH)
	{
		this.SetSizeRelH({RelativeFrom : AscCommon.c_oAscSizeRelFromH.sizerelfromhPage, Percent : 0});
	}

	if (bCheckWrapPolygon)
	{
		this.Check_WrapPolygon();
	}

	if(undefined != Props.description){
		this.docPr.setDescr(Props.description);
	}
	if(undefined != Props.title){
		this.docPr.setTitle(Props.title);
	}
};
ParaDrawing.prototype.CheckFitToColumn = function()
{
	var oLogicDoc = this.document;
	if(oLogicDoc)
	{
		var oColumnSize = oLogicDoc.GetColumnSize();
		if(oColumnSize)
		{
			if(this.Extent.W > oColumnSize.W || this.Extent.H > oColumnSize.H)
			{
				if(oColumnSize.W > 0 && oColumnSize.H > 0)
				{
					if(this.GraphicObj && !this.GraphicObj.isGroupObject())
					{
						var oXfrm = this.GraphicObj.spPr && this.GraphicObj.spPr.xfrm;
						if(oXfrm)
						{
							var dScaleW = oColumnSize.W/oXfrm.extX;
							var dScaleH = oColumnSize.H/oXfrm.extY;
							var dScale = Math.min(dScaleW, dScaleH);
							oXfrm.setExtX(oXfrm.extX * dScale);
							oXfrm.setExtY(oXfrm.extY * dScale);
							this.CheckWH();
						}
					}
				}
			}
		}
	}
};
ParaDrawing.prototype.CheckWH = function()
{
	if (!this.GraphicObj)
		return;
	var oldExtW = this.Extent.W;
	var oldExtH = this.Extent.H;
	if (this.GraphicObj.spPr && this.GraphicObj.spPr.xfrm
		&& AscFormat.isRealNumber(this.GraphicObj.spPr.xfrm.extX)
		&& AscFormat.isRealNumber(this.GraphicObj.spPr.xfrm.extY))
	{
		this.Extent.W = this.GraphicObj.spPr.xfrm.extX;
		this.Extent.H = this.GraphicObj.spPr.xfrm.extY;
	}
	if(this.GraphicObj.getObjectType() === AscDFH.historyitem_type_Shape ||
		this.GraphicObj.getObjectType() === AscDFH.historyitem_type_SmartArt)
	{
		this.GraphicObj.handleUpdateExtents();
	}
	this.GraphicObj.recalculate();
	this.Extent.W = oldExtW;
	this.Extent.H = oldExtH;
	var extX, extY, rot;
	if (this.GraphicObj.spPr && this.GraphicObj.spPr.xfrm )
	{
		if(AscFormat.isRealNumber(this.GraphicObj.spPr.xfrm.extX) && AscFormat.isRealNumber(this.GraphicObj.spPr.xfrm.extY))
		{
            extX = this.GraphicObj.spPr.xfrm.extX;
            extY = this.GraphicObj.spPr.xfrm.extY;
		}
		else
		{
			extX = 5;
			extY = 5;
		}
		if(AscFormat.isRealNumber(this.GraphicObj.spPr.xfrm.rot))
		{
			rot = this.GraphicObj.spPr.xfrm.rot;
		}
		else
		{
			rot = 0;
		}
	}
	else
	{
		extX = 5;
		extY = 5;
		rot = 0;
	}

	let dKoef = this.GetScaleCoefficient();
	this.setExtent(extX / dKoef, extY / dKoef);


	var EEL = 0.0, EET = 0.0, EER = 0.0, EEB = 0.0;
	var addEEL = 0.0, addEET = 0.0, addEER = 0.0, addEEB = 0.0;

	if (this.GraphicObj.getObjectType() === AscDFH.historyitem_type_SmartArt) {
		addEER = 57150 * AscCommonWord.g_dKoef_emu_to_mm;
		addEEL = 38100 * AscCommonWord.g_dKoef_emu_to_mm;
	}
	//if(this.Is_Inline())
	{
		var xc          = this.GraphicObj.localTransform.TransformPointX(this.GraphicObj.extX / 2.0, this.GraphicObj.extY / 2.0);
		var yc          = this.GraphicObj.localTransform.TransformPointY(this.GraphicObj.extX / 2.0, this.GraphicObj.extY / 2.0);
		var oBounds     = this.GraphicObj.bounds;
		var LineCorrect = 0;
		if (this.GraphicObj.pen && this.GraphicObj.pen.Fill && this.GraphicObj.pen.Fill.fill)
		{
			LineCorrect = (this.GraphicObj.pen.w == null) ? 12700 : parseInt(this.GraphicObj.pen.w);
			LineCorrect /= 72000.0;
		}


		var l = oBounds.x;
		var r = l + oBounds.w;
		var t = oBounds.y;
		var b = t + oBounds.h;

		var startX, startY;
		if(!AscFormat.checkNormalRotate(rot)){
			var temp = extX;
			extX = extY;
			extY = temp;
		}


		startX = xc - extX/2.0;
		startY = yc - extY/2.0;

		if(l > startX){
			l = startX;
		}
		if(r < startX + extX){
			r = startX + extX;
		}
		if(t > startY){
			t = startY;
		}
		if(b < startY + extY){
			b = startY + extY;
		}

		EEL = (xc - extX / 2) - l + LineCorrect + addEEL;
		EET = (yc - extY / 2) - t + LineCorrect + addEET;
		EER = r + LineCorrect - (xc + extX / 2) + addEER;
		EEB = b + LineCorrect - (yc + extY / 2) + addEEB;
	}
	this.setEffectExtent(EEL, EET, EER, EEB);
	this.Check_WrapPolygon();
};
ParaDrawing.prototype.Check_WrapPolygon = function()
{
	if ((this.wrappingType === WRAPPING_TYPE_TIGHT || this.wrappingType === WRAPPING_TYPE_THROUGH) && this.wrappingPolygon && !this.wrappingPolygon.edited)
	{
		this.GraphicObj.recalculate();
		this.wrappingPolygon.setArrRelPoints(this.wrappingPolygon.calculate(this.GraphicObj));
	}
};
ParaDrawing.prototype.Draw = function( X, Y, pGraphics, PDSE)
{
	var nPageIndex = null;
	if(AscCommon.isRealObject(PDSE))
	{
		nPageIndex = PDSE.Page;
	}
	if (pGraphics.isTextDrawer())
	{
		pGraphics.m_aDrawings.push(new AscFormat.ParaDrawingStruct(undefined, this));
		return;
	}
	if (this.Is_Inline())
	{
		pGraphics.shapePageIndex = nPageIndex;
		this.draw(pGraphics, PDSE);
		pGraphics.shapePageIndex = null;
	}

	pGraphics.End_Command();
};

ParaDrawing.prototype.Measure = function()
{
	if (!this.GraphicObj)
	{
		this.Width  = 0;
		this.Height = 0;
		return;
	}
	if (AscFormat.isRealNumber(this.Extent.W) && AscFormat.isRealNumber(this.Extent.H) && (!this.GraphicObj.checkAutofit || !this.GraphicObj.checkAutofit()) && !this.SizeRelH && !this.SizeRelV)
	{
		var oEffectExtent = this.EffectExtent;

		var W, H;
		if(AscFormat.isRealNumber(this.GraphicObj.rot)){
            if(AscFormat.checkNormalRotate(this.GraphicObj.rot))
            {
                W = this.Extent.W;
                H = this.Extent.H;
            }
            else
			{
                W = this.Extent.H;
                H = this.Extent.W;
			}
		}
		else{
			W = this.Extent.W;
			H = this.Extent.H;
		}
		this.Width        = W + AscFormat.getValOrDefault(oEffectExtent.L, 0) + AscFormat.getValOrDefault(oEffectExtent.R, 0);
		this.Height       = H + AscFormat.getValOrDefault(oEffectExtent.T, 0) + AscFormat.getValOrDefault(oEffectExtent.B, 0);
		this.WidthVisible = this.Width;
	}
	else
	{
		this.GraphicObj.recalculate();
		if (this.GraphicObj.recalculateText)
		{
			this.GraphicObj.recalculateText();
		}
		if (this.PositionH.UseAlign || this.Is_Inline())
		{
			this.Width = this.GraphicObj.bounds.w;
		}
		else
		{
			this.Width = this.GraphicObj.extX;
		}
		this.WidthVisible = this.Width;
		if (this.PositionV.UseAlign || this.Is_Inline())
		{
			this.Height = this.GraphicObj.bounds.h;
		}
		else
		{
			this.Height = this.GraphicObj.extY;
		}
	}
};
ParaDrawing.prototype.GetScaleCoefficient = function ()
{
	let paragraph = this.GetParagraph();
	return paragraph ? paragraph.getLayoutScaleCoefficient() : 1;
};
ParaDrawing.prototype.createPlaceholderControl = function (arrObjects)
{
	this.GraphicObj && this.GraphicObj.createPlaceholderControl(arrObjects);
};
ParaDrawing.prototype.IsNeedSaveRecalculateObject = function()
{
	return true;
};
ParaDrawing.prototype.SaveRecalculateObject = function(Copy)
{
	var DrawingObj = {};

	DrawingObj.Type         = this.Type;
	DrawingObj.DrawingType  = this.DrawingType;
	DrawingObj.WrappingType = this.wrappingType;

	if (drawing_Anchor === this.Get_DrawingType() && true === this.Use_TextWrap())
	{
		var oDistance      = this.Get_Distance();
		DrawingObj.FlowPos =
		{
			X : this.X - oDistance.L,
			Y : this.Y - oDistance.T,
			W : this.Width + oDistance.R,
			H : this.Height + oDistance.B
		}
	}
	DrawingObj.PageNum         = this.PageNum;
	DrawingObj.X               = this.X;
	DrawingObj.Y               = this.Y;
	DrawingObj.spRecaclcObject = this.GraphicObj.getRecalcObject();

	return DrawingObj;
};
ParaDrawing.prototype.LoadRecalculateObject = function(RecalcObj)
{
	this.updatePosition3(RecalcObj.PageNum, RecalcObj.X, RecalcObj.Y);
	this.GraphicObj.setRecalcObject(RecalcObj.spRecaclcObject);
};
ParaDrawing.prototype.Reassign_ImageUrls = function(mapUrls)
{
	if (this.GraphicObj)
	{
		this.GraphicObj.Reassign_ImageUrls(mapUrls);
	}
};
ParaDrawing.prototype.PrepareRecalculateObject = function()
{
};
ParaDrawing.prototype.CanAddNumbering = function()
{
	return this.IsInline();
};
ParaDrawing.prototype.Copy = function(oPr)
{
	let c = new ParaDrawing(this.Extent.W, this.Extent.H, null, this.DrawingDocument, null, null);
	c.Set_DrawingType(this.DrawingType);
	if (AscCommon.isRealObject(this.GraphicObj))
	{
		var oCopyPr = new AscFormat.CCopyObjectProperties();
		oCopyPr.contentCopyPr = oPr;
		c.Set_GraphicObject(this.GraphicObj.copy(oCopyPr));
		c.GraphicObj.setParent(c);
	}

	var d = this.Distance;
	c.Set_PositionH(this.PositionH.RelativeFrom, this.PositionH.Align, this.PositionH.Value, this.PositionH.Percent);
	c.Set_PositionV(this.PositionV.RelativeFrom, this.PositionV.Align, this.PositionV.Value, this.PositionV.Percent);
	c.Set_Distance(d.L, d.T, d.R, d.B);
	c.Set_AllowOverlap(this.AllowOverlap);
	c.Set_WrappingType(this.wrappingType);
	if (this.wrappingPolygon)
	{
		c.wrappingPolygon.fromOther(this.wrappingPolygon);
	}
	c.Set_BehindDoc(this.behindDoc);
	c.Set_RelativeHeight(this.RelativeHeight);
	if (this.SizeRelH)
	{
		c.SetSizeRelH({RelativeFrom : this.SizeRelH.RelativeFrom, Percent : this.SizeRelH.Percent});
	}
	if (this.SizeRelV)
	{
		c.SetSizeRelV({RelativeFrom : this.SizeRelV.RelativeFrom, Percent : this.SizeRelV.Percent});
	}
	if (AscFormat.isRealNumber(this.Extent.W) && AscFormat.isRealNumber(this.Extent.H))
	{
		c.setExtent(this.Extent.W, this.Extent.H);
	}
	var EE = this.EffectExtent;
	if (EE.L > 0 || EE.T > 0 || EE.R > 0 || EE.B > 0)
	{
		c.setEffectExtent(EE.L, EE.T, EE.R, EE.B);
	}
	c.docPr.setFromOther(this.docPr);
	if (this.ParaMath)
		c.Set_ParaMath(this.ParaMath.Copy());

	c.SetForm(this.Form);
	return c;
};
ParaDrawing.prototype.IsEqual = function(oElement)
{
	return false;
};
ParaDrawing.prototype.Get_Id = function()
{
	return this.Id;
};
ParaDrawing.prototype.GetId = function()
{
	return this.Id;
};
ParaDrawing.prototype.setParagraphTabs = function(tabs)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.setParagraphTabs === "function")
		this.GraphicObj.setParagraphTabs(tabs);
};
ParaDrawing.prototype.IsMovingTableBorder = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.IsMovingTableBorder === "function")
		return this.GraphicObj.IsMovingTableBorder();
	return false;
};
ParaDrawing.prototype.SetVerticalClip = function(Top, Bottom)
{
	this.LineTop = Top;
	this.LineBottom = Bottom;
};
ParaDrawing.prototype.Update_Position = function(Paragraph, ParaLayout, PageLimits, PageLimitsOrigin, LineNum)
{
	if (undefined != this.PositionH_Old)
	{
		this.PositionH.RelativeFrom = this.PositionH_Old.RelativeFrom2;
		this.PositionH.Align        = this.PositionH_Old.Align2;
		this.PositionH.Value        = this.PositionH_Old.Value2;
		this.PositionH.Percent      = this.PositionH_Old.Percent2;
	}

	if (undefined != this.PositionV_Old)
	{
		this.PositionV.RelativeFrom = this.PositionV_Old.RelativeFrom2;
		this.PositionV.Align        = this.PositionV_Old.Align2;
		this.PositionV.Value        = this.PositionV_Old.Value2;
		this.PositionV.Percent      = this.PositionV_Old.Percent2;
	}

	this.Parent            = Paragraph;
	const oDocumentContent = this.GetDocumentContent();
	this.DocumentContent   = oDocumentContent;
	let PageNum            = ParaLayout.PageNum;

	let floatObjectsOnPage = this.graphicObjects ? this.graphicObjects.getAllFloatObjectsOnPage(PageNum, this.Parent.Parent) : [];
	var bInline            = this.Is_Inline();
	this.Internal_Position.SetScaleFactor(this.GetScaleCoefficient());
	this.Internal_Position.Set(this.GraphicObj.extX, this.GraphicObj.extY, this.getXfrmRot(), this.EffectExtent, this.YOffset, ParaLayout, PageLimits);
	this.Internal_Position.Calculate_X(bInline, this.PositionH.RelativeFrom, this.PositionH.Align, this.PositionH.Value, this.PositionH.Percent);
	this.Internal_Position.Calculate_Y(bInline, this.PositionV.RelativeFrom, this.PositionV.Align, this.PositionV.Value, this.PositionV.Percent);


	let bCorrect = false;
	if(oDocumentContent && oDocumentContent.IsTableCellContent && oDocumentContent.IsTableCellContent(false))
	{
		bCorrect = true;
	}
	if(this.PositionH.RelativeFrom !== c_oAscRelativeFromH.Page || this.PositionV.RelativeFrom !== c_oAscRelativeFromV.Page)
	{
		bCorrect = true;
	}
	this.Internal_Position.Correct_Values(bInline, PageLimits, this.AllowOverlap, this.Use_TextWrap(), floatObjectsOnPage, bCorrect);
	this.GraphicObj.bounds.l = this.GraphicObj.bounds.x + this.Internal_Position.CalcX;
	this.GraphicObj.bounds.r =  this.GraphicObj.bounds.x  + this.GraphicObj.bounds.w + this.Internal_Position.CalcX;
	this.GraphicObj.bounds.t = this.GraphicObj.bounds.y + this.Internal_Position.CalcY;
	this.GraphicObj.bounds.b = this.GraphicObj.bounds.y + this.GraphicObj.bounds.h + this.Internal_Position.CalcY;

	let OldPageNum = this.PageNum;
	this.PageNum   = PageNum;
	this.LineNum   = LineNum;
	this.X         = this.Internal_Position.CalcX;
	this.Y         = this.Internal_Position.CalcY;

	if (undefined != this.PositionH_Old)
	{
		// Восстанови старые значения, чтобы в историю изменений все нормально записалось
		this.PositionH.RelativeFrom = this.PositionH_Old.RelativeFrom;
		this.PositionH.Align        = this.PositionH_Old.Align;
		this.PositionH.Value        = this.PositionH_Old.Value;
		this.PositionH.Percent      = this.PositionH_Old.Percent;

		// Рассчитаем сдвиг с учетом старой привязки
		var Value = this.Internal_Position.Calculate_X_Value(this.PositionH_Old.RelativeFrom);
		this.Set_PositionH(this.PositionH_Old.RelativeFrom, false, Value, false);
		// На всякий случай пересчитаем заново координату
		this.X = this.Internal_Position.Calculate_X(bInline, this.PositionH.RelativeFrom, this.PositionH.Align, this.PositionH.Value, this.PositionH.Percent);
	}

	if (undefined != this.PositionV_Old)
	{
		// Восстанови старые значения, чтобы в историю изменений все нормально записалось
		this.PositionV.RelativeFrom = this.PositionV_Old.RelativeFrom;
		this.PositionV.Align        = this.PositionV_Old.Align;
		this.PositionV.Value        = this.PositionV_Old.Value;
		this.PositionV.Percent      = this.PositionV_Old.Percent;

		// Рассчитаем сдвиг с учетом старой привязки
		var Value = this.Internal_Position.Calculate_Y_Value(this.PositionV_Old.RelativeFrom);
		this.Set_PositionV(this.PositionV_Old.RelativeFrom, false, Value, false);
		// На всякий случай пересчитаем заново координату
		this.Y = this.Internal_Position.Calculate_Y(bInline, this.PositionV.RelativeFrom, this.PositionV.Align, this.PositionV.Value, this.PositionV.Percent);
	}
	this.OrigX = this.X;
	this.OrigY = this.Y;
	this.ShiftX = 0;
	this.ShiftY = 0;
	this.updatePosition3(this.PageNum, this.X, this.Y, OldPageNum);
	this.useWrap = this.Use_TextWrap();
};
ParaDrawing.prototype.GetClipRect = function ()
{
	if (this.Is_Inline() || this.Use_TextWrap())
	{
		const oCell = this.isTableCellChild(true);
		if (oCell)
		{
			var arrPages = oCell.GetCurPageByAbsolutePage(this.PageNum);
			for (var nIndex = 0, nCount = arrPages.length; nIndex < nCount; ++nIndex)
			{
				var oPageBounds = oCell.GetPageBounds(arrPages[nIndex]);
				if(this.GraphicObj.bounds.isIntersect(oPageBounds.Left, oPageBounds.Top, oPageBounds.Right, oPageBounds.Bottom))
				{
					return new AscFormat.CGraphicBounds(oPageBounds.Left, oPageBounds.Top, oPageBounds.Right, oPageBounds.Bottom);
				}
			}
		}
	}
	return null;
};
ParaDrawing.prototype.Update_PositionYHeaderFooter = function(TopMarginY, BottomMarginY)
{
	var oParagraph = this.GetParagraph();
	if (!oParagraph)
		return;

	if (oParagraph.IsTableCellContent() && this.IsLayoutInCell())
		return;

	this.Internal_Position.Update_PositionYHeaderFooter(TopMarginY, BottomMarginY);
	this.Internal_Position.Calculate_Y(this.Is_Inline(), this.PositionV.RelativeFrom, this.PositionV.Align, this.PositionV.Value, this.PositionV.Percent);
	this.OrigY = this.Internal_Position.CalcY;
	this.Y = this.OrigY + this.ShiftY;
	this.updatePosition3(this.PageNum, this.X, this.Y, this.PageNum);
};
ParaDrawing.prototype.Reset_SavedPosition = function()
{
	this.PositionV_Old = undefined;
	this.PositionH_Old = undefined;
};
ParaDrawing.prototype.setParagraphBorders = function(val)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.setParagraphBorders === "function")
		this.GraphicObj.setParagraphBorders(val);
};
ParaDrawing.prototype.deselect = function()
{
	this.selected = false;
	if (this.GraphicObj && this.GraphicObj.deselect)
		this.GraphicObj.deselect();
};

ParaDrawing.prototype.updatePosition3 = function(pageIndex, x, y, oldPageNum)
{
	let _x = x;
	let _y = y;

	if(this.graphicObjects)
	{
		this.graphicObjects.removeById(pageIndex, this.Get_Id());
		if (AscFormat.isRealNumber(oldPageNum))
		{
			this.graphicObjects.removeById(oldPageNum, this.Get_Id());
		}
	}
	var bChangePageIndex = this.pageIndex !== pageIndex;
	this.setPageIndex(pageIndex);
	let bIsHfdFtr = this.isHdrFtrChild(false);
	let oDocContent = this.GetDocumentContent();
	if (typeof this.GraphicObj.setStartPage === "function")
	{
		this.GraphicObj.setStartPage(pageIndex, bIsHfdFtr, bIsHfdFtr || bChangePageIndex);
	}
	if(this.graphicObjects)
	{
		if (!(bIsHfdFtr && oDocContent && oDocContent.Get_StartPage_Absolute() !== pageIndex))
		{
			// TODO: ситуацию в колонтитуле с привязкой к полю нужно отдельно обрабатывать через дополнительные пересчеты
			if (!(bIsHfdFtr && this.GetPositionV().RelativeFrom !== Asc.c_oAscRelativeFromV.Margin))
			{
				this.graphicObjects.addObjectOnPage(pageIndex, this.GraphicObj);
			}
			this.bNoNeedToAdd = false;
		}
		else
		{
			this.bNoNeedToAdd = true;
		}
	}

	if(this.bCellHF)
	{
		let oDocContent = this.DocumentContent;
		let oParagraph = this.GetParagraph();
		if(oDocContent && oParagraph)
		{
			let dExtX = this.getXfrmExtX();
			if(dExtX > oDocContent.XLimit)
			{
				let nAlign = oParagraph.GetParagraphAlign();
				switch (nAlign)
				{
					case AscCommon.align_Center:
					{
						_x = oDocContent.X + oDocContent.XLimit / 2 - dExtX / 2;
						break;
					}
					case AscCommon.align_Right:
					{
						_x = oDocContent.X + oDocContent.XLimit - dExtX;
						break;
					}
					case AscCommon.align_Left:
					{
						_x = oDocContent.X;
						break;
					}
				}
			}
		}
	}

	if (this.GraphicObj.bNeedUpdatePosition || !(AscFormat.isRealNumber(this.GraphicObj.posX) && AscFormat.isRealNumber(this.GraphicObj.posY)) || !(Math.abs(this.GraphicObj.posX - _x) < MOVE_DELTA && Math.abs(this.GraphicObj.posY - _y) < MOVE_DELTA))
		this.GraphicObj.updatePosition(_x, _y);
	if (this.GraphicObj.bNeedUpdatePosition || !(AscFormat.isRealNumber(this.wrappingPolygon.posX) && AscFormat.isRealNumber(this.wrappingPolygon.posY)) || !(Math.abs(this.wrappingPolygon.posX - _x) < MOVE_DELTA && Math.abs(this.wrappingPolygon.posY - _y) < MOVE_DELTA))
		this.wrappingPolygon.updatePosition(_x, _y);

	this.calculateSnapArrays();
	
	if (this.GraphicObj.checkFormShiftView)
		this.GraphicObj.checkFormShiftView();
};
ParaDrawing.prototype.calculateAfterChangeTheme = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.calculateAfterChangeTheme === "function")
	{
		this.GraphicObj.calculateAfterChangeTheme();
	}
};
ParaDrawing.prototype.selectionIsEmpty = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.selectionIsEmpty === "function")
		return this.GraphicObj.selectionIsEmpty();
	return false;
};
ParaDrawing.prototype.recalculateDocContent = function()
{
};
ParaDrawing.prototype.Shift = function(nShiftX, nShiftY, nPageAbs)
{
	if (undefined !== nPageAbs)
		this.PageNum = nPageAbs;

	this.ShiftX += nShiftX;
	this.ShiftY += nShiftY;

	this.X = this.OrigX + this.ShiftX;
	this.Y = this.OrigY + this.ShiftY;

	this.updatePosition3(this.PageNum, this.X, this.Y);
};
ParaDrawing.prototype.IsMoveWithTextHorizontally = function()
{
	var oHorRelative = this.GetPositionH().RelativeFrom;
	return (Asc.c_oAscRelativeFromH.Column === oHorRelative || Asc.c_oAscRelativeFromH.Character === oHorRelative);
};
ParaDrawing.prototype.IsMoveWithTextVertically = function()
{
	var oVertRelative = this.GetPositionV().RelativeFrom;
	return (Asc.c_oAscRelativeFromV.Paragraph === oVertRelative || Asc.c_oAscRelativeFromV.Line === oVertRelative);
};
ParaDrawing.prototype.IsLayoutInCell = function()
{
	// Начиная с 15-ой версии Word автофигуры всегда считает расположенными внутри ячейки, и данный флаг считается как true
	if (this.LogicDocument && this.LogicDocument.GetCompatibilityMode() >= AscCommon.document_compatibility_mode_Word15)
		return true;

	return this.LayoutInCell;
};
ParaDrawing.prototype.Get_Distance = function()
{
	var oDist = this.Distance;
	return new AscFormat.CDistance(AscFormat.getValOrDefault(oDist.L, AscFormat.DISTANCE_TO_TEXT_LEFTRIGHT), AscFormat.getValOrDefault(oDist.T, 0), AscFormat.getValOrDefault(oDist.R, AscFormat.DISTANCE_TO_TEXT_LEFTRIGHT), AscFormat.getValOrDefault(oDist.B, 0));
};
ParaDrawing.prototype.Set_XYForAdd = function(X, Y, NearPos, PageNum)
{
	if (null !== NearPos)
	{
		var Layout = NearPos.Paragraph.Get_Layout(NearPos.ContentPos, this);
		this.private_SetXYByLayout(X, Y, PageNum, Layout, true, true);

		var nRecalcIndex   = null;
		var oLogicDocument = this.document;
		if (oLogicDocument)
		{
			nRecalcIndex = oLogicDocument.Get_History().GetRecalculateIndex();
			this.SetSkipOnRecalculate(true);
            oLogicDocument.TurnOff_InterfaceEvents();
            oLogicDocument.Recalculate();
            oLogicDocument.TurnOn_InterfaceEvents(false);
			this.SetSkipOnRecalculate(false);
		}

		if (null !== nRecalcIndex)
			oLogicDocument.Get_History().SetRecalculateIndex(nRecalcIndex);

		Layout = NearPos.Paragraph.Get_Layout(NearPos.ContentPos, this);
		this.private_SetXYByLayout(X, Y, PageNum, Layout, true, true);
	}
};
ParaDrawing.prototype.SetSkipOnRecalculate = function(isSkip)
{
	this.SkipOnRecalculate = isSkip;
};
ParaDrawing.prototype.IsSkipOnRecalculate = function()
{
	return this.SkipOnRecalculate;
};
ParaDrawing.prototype.Set_XY = function(X, Y, Paragraph, PageNum, bResetAlign)
{
	if (Paragraph)
	{
		var PageNumOld = this.PageNum;
		var ContentPos = Paragraph.Get_DrawingObjectContentPos(this.Get_Id());
		if (null === ContentPos)
			return;

		var Layout = Paragraph.Get_Layout(ContentPos, this);
		this.private_SetXYByLayout(X, Y, PageNum, Layout, (bResetAlign || true !== this.PositionH.Align ? true : false), (bResetAlign || true !== this.PositionV.Align ? true : false));

		var nRecalcIndex   = null;
		var oLogicDocument = this.document;
		if (oLogicDocument)
		{
			nRecalcIndex = oLogicDocument.Get_History().GetRecalculateIndex();
			this.SetSkipOnRecalculate(true);
			oLogicDocument.Recalculate();
			this.SetSkipOnRecalculate(false);
		}

		if (null !== nRecalcIndex)
			oLogicDocument.Get_History().SetRecalculateIndex(nRecalcIndex);

		if (!this.LogicDocument
			|| null === this.LogicDocument.FullRecalc.Id
			|| (PageNum < this.LogicDocument.FullRecalc.PageIndex
			&& PageNumOld < this.LogicDocument.FullRecalc.PageIndex))
			Layout = Paragraph.Get_Layout(ContentPos, this);

		this.private_SetXYByLayout(X, Y, PageNum, Layout, (bResetAlign || true !== this.PositionH.Align ? true : false), (bResetAlign || true !== this.PositionV.Align ? true : false));
	}
};
ParaDrawing.prototype.private_SetXYByLayout = function(X, Y, PageNum, Layout, bChangeX, bChangeY)
{
	if(!Layout)
	{
		return;
	}
	this.PageNum = PageNum;

	this.Internal_Position.Set(this.GraphicObj.extX, this.GraphicObj.extY, this.getXfrmRot(), this.EffectExtent, this.YOffset, Layout.ParagraphLayout, Layout.PageLimitsOrigin);
	this.Internal_Position.Calculate_X(false, c_oAscRelativeFromH.Page, false, X - Layout.PageLimitsOrigin.X, false);
	this.Internal_Position.Calculate_Y(false, c_oAscRelativeFromV.Page, false, Y - Layout.PageLimitsOrigin.Y, false);
	let bCorrect = false;
	if(this.isTableCellChild(false))
	{
		bCorrect = true;
	}
	if(this.PositionH.RelativeFrom !== c_oAscRelativeFromH.Page || this.PositionV.RelativeFrom !== c_oAscRelativeFromV.Page)
	{
		bCorrect = true;
	}
	this.Internal_Position.Correct_Values(false, Layout.PageLimits, this.AllowOverlap, this.Use_TextWrap(), [], bCorrect);

	if (true === bChangeX)
	{
		this.X = this.Internal_Position.CalcX;

		// Рассчитаем сдвиг с учетом старой привязки
		var ValueX = this.Internal_Position.Calculate_X_Value(this.PositionH.RelativeFrom);
		this.Set_PositionH(this.PositionH.RelativeFrom, false, ValueX, false);

		// На всякий случай пересчитаем заново координату
		this.X = this.Internal_Position.Calculate_X(false, this.PositionH.RelativeFrom, this.PositionH.Align, this.PositionH.Value, this.PositionH.Percent);
	}

	if (true === bChangeY)
	{
		this.Y = this.Internal_Position.CalcY;

		// Рассчитаем сдвиг с учетом старой привязки
		var ValueY = this.Internal_Position.Calculate_Y_Value(this.PositionV.RelativeFrom);
		this.Set_PositionV(this.PositionV.RelativeFrom, false, ValueY, false);

		// На всякий случай пересчитаем заново координату
		this.Y = this.Internal_Position.Calculate_Y(false, this.PositionV.RelativeFrom, this.PositionV.Align, this.PositionV.Value, this.PositionV.Percent);
	}
};
ParaDrawing.prototype.Get_DrawingType = function()
{
	return this.DrawingType;
};
ParaDrawing.prototype.Is_Inline = function()
{
	if (Asc.editor.isPdfEditor()) {
		return drawing_Inline === this.DrawingType;
	}

	if(this.Parent &&
		this.Parent.Get_ParentTextTransform &&
		this.Parent.Get_ParentTextTransform())
	{
		return true;
	}
	
	
	let para = this.GetParagraph();
	let logicDocument = para ? para.GetLogicDocument() : null;
	if (logicDocument
		&& logicDocument.IsDocumentEditor()
		&& this.IsForm()
		&& logicDocument.GetDocumentLayout().IsReadMode())
	{
		return true;
	}

	return drawing_Inline === this.DrawingType;
};
ParaDrawing.prototype.IsInline = function()
{
	return this.Is_Inline();
};
ParaDrawing.prototype.MakeInline = function()
{
	if (drawing_Inline === this.Get_DrawingType())
		return;
	
	this.Set_DrawingType(drawing_Inline);
	
	if (AscCommon.isRealObject(this.GraphicObj.bounds)
		&& AscFormat.isRealNumber(this.GraphicObj.bounds.w)
		&& AscFormat.isRealNumber(this.GraphicObj.bounds.h))
	{
		this.CheckWH();
	}
};
ParaDrawing.prototype.IsForm = function()
{
	return this.Form;
};
ParaDrawing.prototype.SetForm = function(isForm)
{
	History.Add(new CChangesParaDrawingForm(this, this.DrawingType, isForm));
	this.Form = isForm;
}
ParaDrawing.prototype.IsInForm = function()
{
	let run = this.GetRun();
	if (!run)
		return false;
	
	let parentCCs = run.GetParentContentControls();
	for (let i = 0; i < parentCCs.length; ++i)
	{
		if (parentCCs[i].IsForm())
			return true;
	}
	
	return false;
};
ParaDrawing.prototype.GetInnerForm = function()
{
	return this.GraphicObj ? this.GraphicObj.getInnerForm() : null;
};
ParaDrawing.prototype.Use_TextWrap = function()
{
	if (this.IsInline())
		return false;
	
	// TODO: Проверить, возможно данную проверку можно заменить на Paragraph.IsInline()
	// Если автофигура привязана к параграфу с рамкой, обтекание не делается
	if (!this.Parent
		|| !this.Parent.GetFramePr
		|| (this.Parent.GetFramePr() && !this.Parent.GetFramePr().IsInline()))
		return false;

	// здесь должна быть проверка, нужно ли использовать обтекание относительно данного объекта,
	// или он просто лежит над или под текстом.
	return ( drawing_Anchor === this.DrawingType && !(this.wrappingType === WRAPPING_TYPE_NONE) );
};
ParaDrawing.prototype.IsUseTextWrap = function()
{
	return this.Use_TextWrap();
};
ParaDrawing.prototype.Draw_Selection = function()
{
	var Padding = this.DrawingDocument.GetMMPerDot(6);
	var extX, extY;
	if(this.GraphicObj)
	{
		extX = this.GraphicObj.extX;
		extY = this.GraphicObj.extY;
	}
	else
	{
		extX = this.getXfrmExtX();
		extY = this.getXfrmExtY();
	}
	var rot = this.getXfrmRot();
	if(AscFormat.checkNormalRotate(rot))
	{
		this.DrawingDocument.AddPageSelection(this.PageNum, this.X - this.EffectExtent.L - Padding, this.Y - this.EffectExtent.T - Padding, this.EffectExtent.L + extX + this.EffectExtent.R + 2 * Padding, this.EffectExtent.T + extY + this.EffectExtent.B + 2 * Padding);
	}
	else
	{

		this.DrawingDocument.AddPageSelection(this.PageNum, this.X + extX / 2.0 - extY / 2.0 - this.EffectExtent.L - Padding, this.Y + extY / 2.0 - extX / 2.0 - this.EffectExtent.T - Padding, this.EffectExtent.L + extY + this.EffectExtent.R + 2 * Padding, this.EffectExtent.T + extX + this.EffectExtent.B + 2 * Padding);
	}

};
ParaDrawing.prototype.CanInsertToPos = function(oAnchorPos)
{
	// Автофигуры не вставляем в другие автофигуры, сноски и концевые сноски
	if (!oAnchorPos || !oAnchorPos.Paragraph || !oAnchorPos.Paragraph.Parent)
		return false;

	return !((this.IsShape() || this.IsGroup()) && (true === oAnchorPos.Paragraph.Parent.Is_DrawingShape() || true === oAnchorPos.Paragraph.Parent.IsFootnote()));
};
ParaDrawing.prototype.OnEnd_MoveInline = function(NearPos)
{
	if (!this.Parent)
		return;

	NearPos.Paragraph.Check_NearestPos(NearPos);

	var oRun       = this.GetRun();
	var oPictureCC = false;
	if (oRun)
	{
		var arrContentControls = oRun.GetParentContentControls();
		for (var nIndex = arrContentControls.length - 1; nIndex >= 0; --nIndex)
		{
			if (arrContentControls[nIndex].IsPicture())
			{
				oPictureCC = arrContentControls[nIndex];
				break;
			}
		}
	}

	var oDstRun = null;
	var arrClasses = NearPos.Paragraph.GetClassesByPos(NearPos.ContentPos);
	for (var nIndex = arrClasses.length - 1; nIndex >= 0; --nIndex)
	{
		if (arrClasses[nIndex] instanceof ParaRun)
		{
			oDstRun = arrClasses[nIndex];
			break;
		}
	}

	// Ничего никуда не переносим в такой ситуации
	if (!oDstRun || (oPictureCC && oDstRun === oRun) || oDstRun.GetParentForm())
	{
		NearPos.Paragraph.Clear_NearestPosArray();
		return;
	}

	// При переносе всегда создаем копию, чтобы в совместном редактировании не было проблем
	var NewParaDrawing = this.Copy();
	var RunPr = this.Remove_FromDocument(false);

	NewParaDrawing.AddToDocument(NearPos, RunPr, undefined, oPictureCC);
	NewParaDrawing.SelectAsDrawing();
};
ParaDrawing.prototype.Get_ParentTextTransform = function()
{
	if (this.Parent)
	{
		return this.Parent.Get_ParentTextTransform();
	}
	return null;
};
ParaDrawing.prototype.GoToText = function(bBefore, bUpdateStates)
{
	var Paragraph = this.Get_ParentParagraph();
	if (Paragraph)
	{
		Paragraph.MoveCursorToDrawing(this.Id, bBefore);
		Paragraph.Document_SetThisElementCurrent(undefined === bUpdateStates ? true : bUpdateStates);
	}
};
ParaDrawing.prototype.GetDocumentPositionFromObject = function()
{
	var oParagraph = this.GetParagraph();
	if (!oParagraph)
		return [];

	var oLogicDocument = oParagraph.GetLogicDocument();
	if (!oLogicDocument)
		return [];

	var oParaContentPos = oParagraph.GetPosByDrawing(this.GetId());
	if (!oParaContentPos)
		return [];

	var nInRunPos = oParaContentPos.Get(oParaContentPos.GetDepth());

	var oRunPos = oParaContentPos.Copy();
	oRunPos.DecreaseDepth(1);

	var oRun = oParagraph.GetClassByPos(oRunPos);

	if (!(oRun instanceof ParaRun))
		return [];

	var oRunDocPos = oRun.GetDocumentPositionFromObject();
	if (!oRunDocPos || !oRunDocPos.length)
		return [];

	oRunDocPos.push({Class : oRun, Position : nInRunPos});

	return oRunDocPos;
};
ParaDrawing.prototype.Remove_FromDocument = function(bRecalculate)
{
	var oResult = null;
	if(!this.Parent)
	{
		return oResult;
	}

	var oRun = this.Parent.Get_DrawingObjectRun(this.Id);
	if (oRun)
	{

		var oGrObject = this.GraphicObj;
		if(oGrObject && oGrObject.getObjectType() === AscDFH.historyitem_type_Shape)
		{
			if(oGrObject.signatureLine)
			{
				oGrObject.setSignature(oGrObject.signatureLine);
			}
		}
		var oPictureCC         = null;
		var arrContentControls = oRun.GetParentContentControls();
		for (var nIndex = arrContentControls.length - 1; nIndex >= 0; --nIndex)
		{
			if (arrContentControls[nIndex].IsPicture())
			{
				oPictureCC = arrContentControls[nIndex];
				break;
			}
		}

		if (oPictureCC)
			oPictureCC.RemoveContentControlWrapper();

		oRun.Remove_DrawingObject(this.Id);
		if (oRun.IsInHyperlink())
			oResult = new CTextPr();
		else
			oResult = oRun.GetTextPr();

		if(oGrObject && oGrObject.getObjectType() === AscDFH.historyitem_type_Shape)
		{
			if(oGrObject.signatureLine)
			{
				editor.sendEvent("asc_onRemoveSignature", oGrObject.signatureLine.id);
				oGrObject.setSignature(oGrObject.signatureLine);
			}
		}
	}

	if (false != bRecalculate)
		editor.WordControl.m_oLogicDocument.Recalculate();

	return oResult;
};
ParaDrawing.prototype.Get_ParentParagraph = function()
{
	if (this.Parent instanceof Paragraph)
		return this.Parent;

	if (this.Parent instanceof ParaRun)
		return this.Parent.Paragraph;

	if (this.Parent && this.Parent.GetParagraph)
		return this.Parent.GetParagraph();

	return null;
};
ParaDrawing.prototype.SelectAsText = function()
{
	var oParagraph = this.GetParagraph();
	var oRun       = this.GetRun();
	if (!oParagraph || !oRun)
		return;

	var oDocument = oParagraph.GetLogicDocument();
	if (!oDocument)
		return;

	oDocument.RemoveSelection();

	oRun.Make_ThisElementCurrent(false);
	oRun.SetCursorPosition(oRun.GetElementPosition(this));

	var oStartPos = oDocument.GetContentPosition(false);
	oRun.SetCursorPosition(oRun.GetElementPosition(this) + 1);
	var oEndPos = oDocument.GetContentPosition(false);

	oDocument.RemoveSelection();
	oDocument.SetSelectionByContentPositions(oStartPos, oEndPos);
};
ParaDrawing.prototype.SelectAsDrawing = function()
{
	let oParagraph = this.GetParagraph();
	let oRun       = this.GetRun();
	if (!oParagraph || !oRun)
		return;

	let oDocument = oParagraph.GetLogicDocument();
	if (!oDocument)
		return;

	oDocument.RemoveSelection();
	oDocument.Select_DrawingObject(this.GetId());
};
ParaDrawing.prototype.AddToDocument = function(oAnchorPos, oRunPr, oRun, oPictureCC)
{
	let oAnchorParagraph = oAnchorPos.Paragraph;
	oAnchorParagraph.Check_NearestPos(oAnchorPos);

	let oInsertParagraph = new AscWord.Paragraph();
	var oDrawingRun = new AscCommonWord.ParaRun();
	oDrawingRun.AddToContent(0, this);

	if (oRunPr)
		oDrawingRun.SetPr(oRunPr.Copy());

	if (oRun)
		oDrawingRun.SetReviewTypeWithInfo(oRun.GetReviewType(), oRun.GetReviewInfo());

	if (oPictureCC)
	{
		let oSdt = new CInlineLevelSdt();
		oSdt.SetPicturePr(true);
		oSdt.ReplacePlaceHolderWithContent();
		oSdt.AddToContent(0, oDrawingRun);
		oInsertParagraph.AddToContent(0, oSdt);
		oPictureCC.private_CopyPrTo(oSdt);
	}
	else
	{
		oInsertParagraph.AddToContent(0, oDrawingRun);
	}

	let oSelectedContent = new AscCommonWord.CSelectedContent();
	oSelectedContent.Add(new AscCommonWord.CSelectedElement(oInsertParagraph, false));
	oSelectedContent.EndCollect();
	oSelectedContent.SetMoveDrawing(true);
	oSelectedContent.Insert(oAnchorPos, true);
	this.SelectAsDrawing();
	oAnchorParagraph.Clear_NearestPosArray();
};
ParaDrawing.prototype.AddToParagraph = function(oParagraph)
{
	let oRun = new AscCommonWord.ParaRun();
	oRun.AddToContent(0, this);
	oParagraph.AddToContent(0, oRun);
};
ParaDrawing.prototype.UpdateCursorType = function(X, Y, PageIndex)
{
	this.DrawingDocument.SetCursorType("move", new AscCommon.CMouseMoveData());

	if (null != this.Parent)
	{
		var Lock = this.Parent.Lock;
		if (true === Lock.Is_Locked())
		{
			var PNum = Math.max(0, Math.min(PageIndex - this.Parent.PageNum, this.Parent.Pages.length - 1));
			var _X   = this.Parent.Pages[PNum].X;
			var _Y   = this.Parent.Pages[PNum].Y;

			var MMData              = new AscCommon.CMouseMoveData();
			var Coords              = this.DrawingDocument.ConvertCoordsToCursorWR(_X, _Y, this.Parent.Get_StartPage_Absolute() + ( PageIndex - this.Parent.PageNum ));
			MMData.X_abs            = Coords.X - 5;
			MMData.Y_abs            = Coords.Y;
			MMData.Type             = Asc.c_oAscMouseMoveDataTypes.LockedObject;
			MMData.UserId           = Lock.Get_UserId();
			MMData.HaveChanges      = Lock.Have_Changes();
			MMData.LockedObjectType = c_oAscMouseMoveLockedObjectType.Common;

			editor.sync_MouseMoveCallback(MMData);
		}
	}
};
ParaDrawing.prototype.Get_AnchorPos = function()
{
	if(!this.Parent)
	{
		return null;
	}
	return this.Parent.Get_AnchorPos(this);
};
ParaDrawing.prototype.CheckRecalcAutoFit = function(oSectPr)
{
	if (this.GraphicObj && this.GraphicObj.CheckNeedRecalcAutoFit)
	{
		if (this.GraphicObj.CheckNeedRecalcAutoFit(oSectPr))
		{
			if (this.GraphicObj)
			{
				this.GraphicObj.recalcWrapPolygon && this.GraphicObj.recalcWrapPolygon();
			}
			this.Measure();
		}
	}
};
//----------------------------------------------------------------------------------------------------------------------
// Undo/Redo функции
//----------------------------------------------------------------------------------------------------------------------
ParaDrawing.prototype.Get_ParentObject_or_DocumentPos = function()
{
	if (this.Parent != null)
		return this.Parent.Get_ParentObject_or_DocumentPos();
};
ParaDrawing.prototype.Refresh_RecalcData = function(Data)
{
	var oRun = this.GetRun();
	if (oRun)
	{
		if (AscCommon.isRealObject(Data))
		{
			switch (Data.Type)
			{
				case AscDFH.historyitem_Drawing_Distance:
				{
					if (this.GraphicObj)
					{
						this.GraphicObj.recalcWrapPolygon && this.GraphicObj.recalcWrapPolygon();
						this.GraphicObj.addToRecalculate();
					}
					break;
				}

				case AscDFH.historyitem_Drawing_SetExtent:
				{
					oRun.RecalcInfo.Measure = true;
					break;
				}

				case AscDFH.historyitem_Drawing_SetSizeRelH:
				case AscDFH.historyitem_Drawing_SetSizeRelV:
				case AscDFH.historyitem_Drawing_SetGraphicObject:
				{
					if (this.GraphicObj)
					{
						this.GraphicObj.handleUpdateExtents && this.GraphicObj.handleUpdateExtents();
						this.GraphicObj.addToRecalculate();
					}
					oRun.RecalcInfo.Measure = true;
					break;
				}
				case AscDFH.historyitem_Drawing_WrappingType:
				{
					if (this.GraphicObj)
					{
						this.GraphicObj.recalcWrapPolygon && this.GraphicObj.recalcWrapPolygon();
						this.GraphicObj.addToRecalculate()
					}
					break;
				}
			}
		}
		return oRun.Refresh_RecalcData2();
	}
};
ParaDrawing.prototype.Refresh_RecalcData2 = function(Data)
{
	var oRun = this.GetRun();
	if (oRun)
		return oRun.Refresh_RecalcData2();
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для совместного редактирования
//----------------------------------------------------------------------------------------------------------------------
ParaDrawing.prototype.hyperlinkCheck = function(bCheck)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.hyperlinkCheck === "function")
		return this.GraphicObj.hyperlinkCheck(bCheck);
	return null;
};
ParaDrawing.prototype.hyperlinkCanAdd = function(bCheckInHyperlink)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.hyperlinkCanAdd === "function")
		return this.GraphicObj.hyperlinkCanAdd(bCheckInHyperlink);
	return false;
};
ParaDrawing.prototype.hyperlinkRemove = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.hyperlinkCanAdd === "function")
		return this.GraphicObj.hyperlinkRemove();
	return false;
};
ParaDrawing.prototype.hyperlinkModify = function( HyperProps )
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.hyperlinkModify === "function")
		return this.GraphicObj.hyperlinkModify(HyperProps);
};
ParaDrawing.prototype.hyperlinkAdd = function( HyperProps )
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.hyperlinkAdd === "function")
		return this.GraphicObj.hyperlinkAdd(HyperProps);
};
ParaDrawing.prototype.documentStatistics = function(stat)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.documentStatistics === "function")
		this.GraphicObj.documentStatistics(stat);
};
ParaDrawing.prototype.documentCreateFontCharMap = function(fontMap)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.documentCreateFontCharMap === "function")
		this.GraphicObj.documentCreateFontCharMap(fontMap);
};
ParaDrawing.prototype.documentCreateFontMap = function(fontMap)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.documentCreateFontMap === "function")
		this.GraphicObj.documentCreateFontMap(fontMap);
};
ParaDrawing.prototype.tableCheckSplit = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.tableCheckSplit === "function")
		return this.GraphicObj.tableCheckSplit();
	return false;
};
ParaDrawing.prototype.tableCheckMerge = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.tableCheckMerge === "function")
		return this.GraphicObj.tableCheckMerge();
	return false;
};
ParaDrawing.prototype.tableSelect = function( Type )
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.tableSelect === "function")
		return this.GraphicObj.tableSelect(Type);
};
ParaDrawing.prototype.tableRemoveTable = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.tableRemoveTable === "function")
		return this.GraphicObj.tableRemoveTable();
};
ParaDrawing.prototype.tableSplitCell = function(Cols, Rows)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.tableSplitCell === "function")
		return this.GraphicObj.tableSplitCell(Cols, Rows);
};
ParaDrawing.prototype.tableMergeCells = function(Cols, Rows)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.tableMergeCells === "function")
		return this.GraphicObj.tableMergeCells(Cols, Rows);
};
ParaDrawing.prototype.tableRemoveCol = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.tableRemoveCol === "function")
		return this.GraphicObj.tableRemoveCol();
};
ParaDrawing.prototype.tableAddCol = function(bBefore, nCount)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.tableAddCol === "function")
		return this.GraphicObj.tableAddCol(bBefore, nCount);
};
ParaDrawing.prototype.tableRemoveRow = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.tableRemoveRow === "function")
		return this.GraphicObj.tableRemoveRow();
};
ParaDrawing.prototype.tableAddRow = function(bBefore, nCount)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.tableAddRow === "function")
		return this.GraphicObj.tableAddRow(bBefore, nCount);
};
ParaDrawing.prototype.getCurrentParagraph = function(bIgnoreSelection, arrSelectedParagraphs)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.getCurrentParagraph === "function")
		return this.GraphicObj.getCurrentParagraph(bIgnoreSelection, arrSelectedParagraphs);

	if (this.Parent instanceof Paragraph)
		return this.Parent;
};
ParaDrawing.prototype.GetSelectedText = function(bClearText, oPr)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.GetSelectedText === "function")
		return this.GraphicObj.GetSelectedText(bClearText, oPr);
	return "";
};
ParaDrawing.prototype.getCurPosXY = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.getCurPosXY === "function")
		return this.GraphicObj.getCurPosXY();
	return {X : 0, Y : 0};
};
ParaDrawing.prototype.setParagraphKeepLines = function(Value)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.setParagraphKeepLines === "function")
		return this.GraphicObj.setParagraphKeepLines(Value);
};
ParaDrawing.prototype.setParagraphKeepNext = function(Value)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.setParagraphKeepNext === "function")
		return this.GraphicObj.setParagraphKeepNext(Value);
};
ParaDrawing.prototype.setParagraphWidowControl = function(Value)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.setParagraphWidowControl === "function")
		return this.GraphicObj.setParagraphWidowControl(Value);
};
ParaDrawing.prototype.setParagraphPageBreakBefore = function(Value)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.setParagraphPageBreakBefore === "function")
		return this.GraphicObj.setParagraphPageBreakBefore(Value);
};
ParaDrawing.prototype.isTextSelectionUse = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.isTextSelectionUse === "function")
		return this.GraphicObj.isTextSelectionUse();
	return false;
};
ParaDrawing.prototype.pasteFormatting = function(oData)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.pasteFormatting === "function")
		return this.GraphicObj.pasteFormatting(oData);
};
ParaDrawing.prototype.getNearestPos = function(x, y, pageIndex)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.getNearestPos === "function")
		return this.GraphicObj.getNearestPos(x, y, pageIndex);
	return null;
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для записи/чтения в поток
//----------------------------------------------------------------------------------------------------------------------
ParaDrawing.prototype.Write_ToBinary = function(Writer)
{
	// Long   : Type
	// String : Id

	Writer.WriteLong(this.Type);
	Writer.WriteString2(this.Id);
};
ParaDrawing.prototype.Write_ToBinary2 = function(Writer)
{
	Writer.WriteLong(AscDFH.historyitem_type_Drawing);
	Writer.WriteString2(this.Id);
	AscFormat.writeDouble(Writer, this.Extent.W);
	AscFormat.writeDouble(Writer, this.Extent.H);
	AscFormat.writeObject(Writer, this.GraphicObj);
	AscFormat.writeObject(Writer, this.DocumentContent);
	AscFormat.writeObject(Writer, this.Parent);
	AscFormat.writeObject(Writer, this.wrappingPolygon);
	AscFormat.writeLong(Writer, this.RelativeHeight);
	AscFormat.writeObject(Writer, this.docPr);
};
ParaDrawing.prototype.Read_FromBinary2 = function(Reader)
{
	this.Id              = Reader.GetString2();
	this.DrawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;
	this.LogicDocument   = this.DrawingDocument ? this.DrawingDocument.m_oLogicDocument : null;

	this.Extent.W        = AscFormat.readDouble(Reader);
	this.Extent.H        = AscFormat.readDouble(Reader);
	this.GraphicObj      = AscFormat.readObject(Reader);
	this.DocumentContent = AscFormat.readObject(Reader);
	this.Parent          = AscFormat.readObject(Reader);
	this.wrappingPolygon = AscFormat.readObject(Reader);
	this.RelativeHeight  = AscFormat.readLong(Reader);
	this.docPr = AscFormat.readObject(Reader);
	if (this.wrappingPolygon)
	{
		this.wrappingPolygon.wordGraphicObject = this;
	}

	this.drawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;
	this.document        = editor.WordControl.m_oLogicDocument;
	this.graphicObjects  = editor.WordControl.m_oLogicDocument.DrawingObjects;
	this.graphicObjects.addGraphicObject(this);
	g_oTableId.Add(this, this.Id);
};
ParaDrawing.prototype.draw = function(graphics, PDSE)
{
	let iO = AscCommon.isRealObject;
	let iN = AscFormat.isRealNumber;
	if (this.GraphicObj)
	{
		graphics.SaveGrState();
		let bInline = this.Is_Inline();
		let oBounds = this.GraphicObj.bounds;
		let paragraph = this.GetParagraph();
		let paraPr = paragraph.GetCompiledParaPr();
		let spacing = paraPr.Spacing;
		if(bInline && spacing.LineRule !== Asc.linerule_Exact && PDSE && iN(this.LineTop) && iN(this.LineBottom) && iO(oBounds))
		{
			let x, y, w, h;
			let oEffectExtent = this.EffectExtent;
			x = PDSE.X;
			y = this.LineTop;
			w = oBounds.r - oBounds.l + AscFormat.getValOrDefault(oEffectExtent.R, 0) + AscFormat.getValOrDefault(oEffectExtent.L, 0);
			h = this.LineBottom - this.LineTop;
			graphics.AddClipRect(x, y, w, h);
		}
		this.GraphicObj.draw(graphics);
		graphics.RestoreGrState();
	}
};
ParaDrawing.prototype.drawAdjustments = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.drawAdjustments === "function")
	{
		this.GraphicObj.drawAdjustments();
	}
};
ParaDrawing.prototype.getTransformMatrix = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.getTransformMatrix === "function")
	{
		return this.GraphicObj.getTransformMatrix();
	}
	return null;
};
ParaDrawing.prototype.getExtensions = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.getExtensions === "function")
	{
		return this.GraphicObj.getExtensions();
	}
	return null;
};
ParaDrawing.prototype.isGroup = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.isGroup === "function")
		return this.GraphicObj.isGroup();
	return false;
};
ParaDrawing.prototype.isShapeChild = function(bRetShape)
{
	const oDocContent = this.GetDocumentContent();
	if (!this.Is_Inline() || !oDocContent)
		return bRetShape ? null : false;

	let oCurDocContent = oDocContent;
	let oCell;
	while (oCell = oCurDocContent.IsTableCellContent(true))
	{
		oCurDocContent = oCell.Row.Table.Parent;
	}
	if(oCurDocContent && oCurDocContent.Is_DrawingShape)
	{
		return oCurDocContent.Is_DrawingShape(bRetShape);
	}
	return bRetShape ? null : false;
};
ParaDrawing.prototype.isHdrFtrChild = function(bRetHdrFtr)
{
	const oContent = this.GetDocumentContent();
	if(!oContent)
	{
		return bRetHdrFtr ? null : false;
	}
	return oContent.IsHdrFtr(bRetHdrFtr);
};
ParaDrawing.prototype.isFootnoteChild = function(bRetFootnote)
{
	const oContent = this.GetDocumentContent();
	if(!oContent)
	{
		return bRetFootnote ? null : false;
	}
	return oContent.IsFootnote(bRetFootnote);
};
ParaDrawing.prototype.isTableCellChild = function(bRetCell)
{
	const oContent = this.GetDocumentContent();
	if(!oContent)
	{
		return bRetCell ? null : false;
	}
	return oContent.IsTableCellContent(bRetCell);
};
ParaDrawing.prototype.isInTopDocument = function(bRetDoc)
{
	const oContent = this.GetDocumentContent();
	if(!oContent)
	{
		return bRetDoc ? null : false;
	}
	return oContent.Is_TopDocument(bRetDoc);
};
ParaDrawing.prototype.checkShapeChildAndGetTopParagraph = function(paragraph)
{
	var parent_paragraph   = !paragraph ? this.Get_ParentParagraph() : paragraph;
	var parent_doc_content = parent_paragraph.Parent;
	if (parent_doc_content.Parent instanceof AscFormat.CShape)
	{
		if (!parent_doc_content.Parent.group)
		{
			return parent_doc_content.Parent.parent.Get_ParentParagraph();
		}
		else
		{
			var top_group = parent_doc_content.Parent.group;
			while (top_group.group)
				top_group = top_group.group;
			return top_group.parent.Get_ParentParagraph();
		}
	}
	else if (parent_doc_content.IsTableCellContent())
	{
		var top_doc_content = parent_doc_content;
		var oCell;
		while (oCell = top_doc_content.IsTableCellContent(true))
		{
			top_doc_content = oCell.Row.Table.Parent;
		}
		if (top_doc_content.Parent instanceof AscFormat.CShape)
		{
			if (!top_doc_content.Parent.group)
			{
				return top_doc_content.Parent.parent.Get_ParentParagraph();
			}
			else
			{
				var top_group = top_doc_content.Parent.group;
				while (top_group.group)
					top_group = top_group.group;
				return top_group.parent.Get_ParentParagraph();
			}
		}
		else
		{
			return parent_paragraph;
		}

	}
	return parent_paragraph;
};
ParaDrawing.prototype.hit = function(x, y)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.hit === "function")
	{
		return this.GraphicObj.hit(x, y);
	}
	return false;
};
ParaDrawing.prototype.hitToTextRect = function(x, y)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.hitToTextRect === "function")
	{
		return this.GraphicObj.hitToTextRect(x, y);
	}
	return false;
};
ParaDrawing.prototype.cursorGetPos = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.cursorGetPos === "function")
	{
		return this.GraphicObj.cursorGetPos();
	}
	return {X : 0, Y : 0};
};
ParaDrawing.prototype.getResizeCoefficients = function(handleNum, x, y)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.getResizeCoefficients === "function")
	{
		return this.GraphicObj.getResizeCoefficients(handleNum, x, y);
	}
	return {kd1 : 1, kd2 : 1};
};
ParaDrawing.prototype.getParagraphParaPr = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.getParagraphParaPr === "function")
	{
		return this.GraphicObj.getParagraphParaPr();
	}
	return null;
};
ParaDrawing.prototype.getParagraphTextPr = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.getParagraphTextPr === "function")
	{
		return this.GraphicObj.getParagraphTextPr();
	}
	return null;
};
ParaDrawing.prototype.getAngle = function(x, y)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.getAngle === "function")
		return this.GraphicObj.getAngle(x, y);
	return 0;
};
ParaDrawing.prototype.calculateSnapArrays = function()
{
	this.GraphicObj.snapArrayX.length = 0;
	this.GraphicObj.snapArrayY.length = 0;
	if (this.GraphicObj)
		this.GraphicObj.recalculateSnapArrays();

};
ParaDrawing.prototype.recalculateCurPos = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.recalculateCurPos === "function")
	{
		this.GraphicObj.recalculateCurPos();
	}
};
ParaDrawing.prototype.setPageIndex = function(newPageIndex)
{
	this.pageIndex = newPageIndex;
	this.PageNum   = newPageIndex;


	if (this.IsShape())
	{
		var arrDrawings = this.GetAllDrawingObjects();
		for (var nIndex = 0, nCount = arrDrawings.length; nIndex < nCount; ++nIndex)
		{
			var oParaDrawing = arrDrawings[nIndex];
			oParaDrawing.setPageIndex(newPageIndex);
		}
	}
};
ParaDrawing.prototype.GetPageNum = function()
{
	return this.PageNum;
};
ParaDrawing.prototype.GetAllParagraphs = function(Props, ParaArray)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.GetAllParagraphs === "function")
		this.GraphicObj.GetAllParagraphs(Props, ParaArray);
};
ParaDrawing.prototype.GetAllTables = function(oProps, arrTables)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.GetAllTables === "function")
		this.GraphicObj.GetAllTables(oProps, arrTables);
};
ParaDrawing.prototype.GetAllDocContents = function(aDocContents)
{
	var _ret = Array.isArray(aDocContents) ? aDocContents : [];
	if(this.GraphicObj)
	{
		this.GraphicObj.getAllDocContents(_ret);
	}
	return _ret;
};
ParaDrawing.prototype.getTableProps = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.getTableProps === "function")
		return this.GraphicObj.getTableProps();
	return null;
};
ParaDrawing.prototype.canGroup = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.canGroup === "function")
		return this.GraphicObj.canGroup();
	return false;
};
ParaDrawing.prototype.canUnGroup = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.canGroup === "function")
		return this.GraphicObj.canUnGroup();
	return false;
};
ParaDrawing.prototype.select = function(pageIndex)
{
	this.selected = true;
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.select === "function")
		this.GraphicObj.select(pageIndex);

};
ParaDrawing.prototype.isSelected = function()
{
	return this.GraphicObj ? this.GraphicObj.selected : false;
};
ParaDrawing.prototype.paragraphClearFormatting = function(isClearParaPr, isClearTextPr)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.paragraphAdd === "function")
		this.GraphicObj.paragraphClearFormatting(isClearParaPr, isClearTextPr);
};
ParaDrawing.prototype.paragraphAdd = function(paraItem, bRecalculate)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.paragraphAdd === "function")
		this.GraphicObj.paragraphAdd(paraItem, bRecalculate);
};
ParaDrawing.prototype.setParagraphShd = function(Shd)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.setParagraphShd === "function")
		this.GraphicObj.setParagraphShd(Shd);
};
ParaDrawing.prototype.getArrayWrapPolygons = function()
{
	if ((AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.getArrayWrapPolygons === "function"))
		return this.GraphicObj.getArrayWrapPolygons();

	return [];
};
ParaDrawing.prototype.getArrayWrapIntervals = function(x0, y0, x1, y1, Y0Sp, Y1Sp, LeftField, RightField, arr_intervals, bMathWrap)
{
	if (this.wrappingType === WRAPPING_TYPE_THROUGH || this.wrappingType === WRAPPING_TYPE_TIGHT)
	{
		y0 = Y0Sp;
		y1 = Y1Sp;
	}

	this.wrappingPolygon.wordGraphicObject = this;
	return this.wrappingPolygon.getArrayWrapIntervals(x0, y0, x1, y1, LeftField, RightField, arr_intervals, bMathWrap);
};
ParaDrawing.prototype.setAllParagraphNumbering = function(numInfo)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.setAllParagraphNumbering === "function")
		this.GraphicObj.setAllParagraphNumbering(numInfo);
};
ParaDrawing.prototype.addNewParagraph = function(bRecalculate)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.addNewParagraph === "function")
		this.GraphicObj.addNewParagraph(bRecalculate);
};
ParaDrawing.prototype.addInlineTable = function(nCols, nRows, nMode)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.addInlineTable === "function")
		return this.GraphicObj.addInlineTable(nCols, nRows, nMode);

	return null;
};
ParaDrawing.prototype.applyTextPr = function(paraItem, bRecalculate)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.applyTextPr === "function")
		this.GraphicObj.applyTextPr(paraItem, bRecalculate);
};
ParaDrawing.prototype.allIncreaseDecFontSize = function(bIncrease)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.allIncreaseDecFontSize === "function")
		this.GraphicObj.allIncreaseDecFontSize(bIncrease);
};
ParaDrawing.prototype.setParagraphNumbering = function(NumInfo)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.setParagraphNumbering === "function")
		this.GraphicObj.setParagraphNumbering(NumInfo);
};
ParaDrawing.prototype.allIncreaseDecIndent = function(bIncrease)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.allIncreaseDecIndent === "function")
		this.GraphicObj.allIncreaseDecIndent(bIncrease);
};
ParaDrawing.prototype.allSetParagraphAlign = function(align)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.allSetParagraphAlign === "function")
		this.GraphicObj.allSetParagraphAlign(align);
};
ParaDrawing.prototype.paragraphIncreaseDecFontSize = function(bIncrease)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.paragraphIncreaseDecFontSize === "function")
		this.GraphicObj.paragraphIncreaseDecFontSize(bIncrease);
};
ParaDrawing.prototype.paragraphIncreaseDecIndent = function(bIncrease)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.paragraphIncreaseDecIndent === "function")
		this.GraphicObj.paragraphIncreaseDecIndent(bIncrease);
};
ParaDrawing.prototype.setParagraphAlign = function(align)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.setParagraphAlign === "function")
		this.GraphicObj.setParagraphAlign(align);
};
ParaDrawing.prototype.setParagraphSpacing = function(Spacing)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.setParagraphSpacing === "function")
		this.GraphicObj.setParagraphSpacing(Spacing);
};
ParaDrawing.prototype.updatePosition = function(x, y)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.updatePosition === "function")
	{
		this.GraphicObj.updatePosition(x, y);
	}
};
ParaDrawing.prototype.updatePosition2 = function(x, y)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.updatePosition === "function")
	{
		this.GraphicObj.updatePosition2(x, y);
	}
};
ParaDrawing.prototype.addInlineImage = function(W, H, Img, GraphicObject, bFlow)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.addInlineImage === "function")
		this.GraphicObj.addInlineImage(W, H, Img, GraphicObject, bFlow);
};
ParaDrawing.prototype.addSignatureLine = function(oSignatureDrawing)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.addSignatureLine === "function")
		this.GraphicObj.addSignatureLine(oSignatureDrawing);
};
ParaDrawing.prototype.canAddComment = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.canAddComment === "function")
		return this.GraphicObj.canAddComment();
	return false;
};
ParaDrawing.prototype.addComment = function(commentData)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.addComment === "function")
		return this.GraphicObj.addComment(commentData);
};
ParaDrawing.prototype.selectionSetStart = function(x, y, event, pageIndex)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.selectionSetStart === "function")
		this.GraphicObj.selectionSetStart(x, y, event, pageIndex);
};
ParaDrawing.prototype.selectionSetEnd = function(x, y, event, pageIndex)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.selectionSetEnd === "function")
		this.GraphicObj.selectionSetEnd(x, y, event, pageIndex);
};
ParaDrawing.prototype.selectionRemove = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.selectionRemove === "function")
		this.GraphicObj.selectionRemove();
};
ParaDrawing.prototype.updateSelectionState = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.updateSelectionState === "function")
		this.GraphicObj.updateSelectionState();
};
ParaDrawing.prototype.cursorMoveLeft = function(AddToSelect, Word)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.cursorMoveLeft === "function")
		this.GraphicObj.cursorMoveLeft(AddToSelect, Word);
};
ParaDrawing.prototype.cursorMoveRight = function(AddToSelect, Word)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.cursorMoveRight === "function")
		this.GraphicObj.cursorMoveRight(AddToSelect, Word);
};
ParaDrawing.prototype.cursorMoveUp = function(AddToSelect)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.cursorMoveUp === "function")
		this.GraphicObj.cursorMoveUp(AddToSelect);
};
ParaDrawing.prototype.cursorMoveDown = function(AddToSelect)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.cursorMoveDown === "function")
		this.GraphicObj.cursorMoveDown(AddToSelect);
};
ParaDrawing.prototype.cursorMoveEndOfLine = function(AddToSelect)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.cursorMoveEndOfLine === "function")
		this.GraphicObj.cursorMoveEndOfLine(AddToSelect);
};
ParaDrawing.prototype.cursorMoveStartOfLine = function(AddToSelect)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.cursorMoveStartOfLine === "function")
		this.GraphicObj.cursorMoveStartOfLine(AddToSelect);
};
ParaDrawing.prototype.remove = function(Count, isRemoveWholeElement, bRemoveOnlySelection, bOnTextAdd)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.remove === "function")
		this.GraphicObj.remove(Count, isRemoveWholeElement, bRemoveOnlySelection, bOnTextAdd);
};
ParaDrawing.prototype.hitToWrapPolygonPoint = function(x, y)
{
	if (this.wrappingPolygon && this.wrappingPolygon.arrPoints.length > 0)
	{
		var radius      = this.drawingDocument.GetMMPerDot(AscCommon.TRACK_CIRCLE_RADIUS);
		var arr_point   = this.wrappingPolygon.calculatedPoints;
		var point_count = arr_point.length;
		var dx, dy;

		var previous_point;
		for (var i = 0; i < arr_point.length; ++i)
		{
			var cur_point = arr_point[i];
			dx            = x - cur_point.x;
			dy            = y - cur_point.y;
			if (Math.sqrt(dx * dx + dy * dy) < radius)
				return {hit : true, hitType : WRAP_HIT_TYPE_POINT, pointNum : i};
		}

		cur_point      = arr_point[0];
		previous_point = arr_point[arr_point.length - 1];
		var vx, vy;
		vx             = cur_point.x - previous_point.x;
		vy             = cur_point.y - previous_point.y;
		if (Math.abs(vx) > 0 || Math.abs(vy) > 0)
		{
			if (HitInLine(this.drawingDocument.CanvasHitContext, x, y, previous_point.x, previous_point.y, cur_point.x, cur_point.y))
				return {hit : true, hitType : WRAP_HIT_TYPE_SECTION, pointNum1 : arr_point.length - 1, pointNum2 : 0};
		}

		for (var point_index = 1; point_index < point_count; ++point_index)
		{
			cur_point      = arr_point[point_index];
			previous_point = arr_point[point_index - 1];

			vx = cur_point.x - previous_point.x;
			vy = cur_point.y - previous_point.y;

			if (Math.abs(vx) > 0 || Math.abs(vy) > 0)
			{
				if (HitInLine(this.drawingDocument.CanvasHitContext, x, y, previous_point.x, previous_point.y, cur_point.x, cur_point.y))
					return {hit   : true,
						hitType   : WRAP_HIT_TYPE_SECTION,
						pointNum1 : point_index - 1,
						pointNum2 : point_index
					};
			}
		}
	}
	return {hit : false};
};
ParaDrawing.prototype.documentGetAllFontNames = function(AllFonts)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.documentGetAllFontNames === "function")
		this.GraphicObj.documentGetAllFontNames(AllFonts);
};
ParaDrawing.prototype.isCurrentElementParagraph = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.isCurrentElementParagraph === "function")
		return this.GraphicObj.isCurrentElementParagraph();
	return false;
};
ParaDrawing.prototype.isCurrentElementTable = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.isCurrentElementTable === "function")
		return this.GraphicObj.isCurrentElementTable();
	return false;
};
ParaDrawing.prototype.canChangeWrapPolygon = function()
{
	if (this.Is_Inline())
		return false;
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.canChangeWrapPolygon === "function")
		return this.GraphicObj.canChangeWrapPolygon();
	return false;
};
ParaDrawing.prototype.init = function()
{
};
ParaDrawing.prototype.calculateAfterOpen = function()
{
};
ParaDrawing.prototype.getBounds = function()
{

	return this.GraphicObj.bounds;
};
ParaDrawing.prototype.getWrapContour = function()
{
	if (AscCommon.isRealObject(this.wrappingPolygon))
	{
		var kw         = 1 / 36000;
		var kh         = 1 / 36000;
		var rel_points = this.wrappingPolygon.relativeArrPoints;
		var ret        = [];
		for (var i = 0; i < rel_points.length; ++i)
		{
			ret[i] = {x : rel_points[i].x * kw, y : rel_points[i].y * kh};
		}
		return ret;
	}
	return [];
};
ParaDrawing.prototype.getDrawingArrayType = function()
{
	if (this.Is_Inline())
		return DRAWING_ARRAY_TYPE_INLINE;
	if (this.behindDoc === true){
		if(this.wrappingType === WRAPPING_TYPE_NONE || (this.document && this.document.GetCompatibilityMode && this.document.GetCompatibilityMode() < AscCommon.document_compatibility_mode_Word15)){
			return DRAWING_ARRAY_TYPE_BEHIND;
		}
	}
	return DRAWING_ARRAY_TYPE_BEFORE;
};
ParaDrawing.prototype.GetWatermarkProps = function()
{
	return this.GraphicObj.getWatermarkProps();
};
ParaDrawing.prototype.documentSearch = function(String, search_Common)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof this.GraphicObj.documentSearch === "function")
		this.GraphicObj.documentSearch(String, search_Common)
};
ParaDrawing.prototype.setParagraphContextualSpacing = function(Value)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.setParagraphContextualSpacing === "function")
		this.GraphicObj.setParagraphContextualSpacing(Value);
};
ParaDrawing.prototype.setParagraphStyle = function(style)
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.setParagraphStyle === "function")
		this.GraphicObj.setParagraphStyle(style);
};

ParaDrawing.prototype.CopyComments = function()
{
	if(!this.GraphicObj)
	{
		return;
	}
	this.GraphicObj.copyComments(this.LogicDocument);
};

ParaDrawing.prototype.copy = function()
{
	var c = new ParaDrawing(this.Extent.W, this.Extent.H, null, editor.WordControl.m_oLogicDocument.DrawingDocument, null, null);
	c.Set_DrawingType(this.DrawingType);
	if (AscCommon.isRealObject(this.GraphicObj))
	{
		var g = this.GraphicObj.copy(undefined);
		c.Set_GraphicObject(g);
		g.setParent(c);
	}
	var d = this.Distance;
	c.Set_PositionH(this.PositionH.RelativeFrom, this.PositionH.Align, this.PositionH.Value, this.PositionH.Percent);
	c.Set_PositionV(this.PositionV.RelativeFrom, this.PositionV.Align, this.PositionV.Value, this.PositionV.Percent);
	c.Set_Distance(d.L, d.T, d.R, d.B);
	c.Set_AllowOverlap(this.AllowOverlap);
	c.Set_WrappingType(this.wrappingType);
	c.Set_BehindDoc(this.behindDoc);
	c.SetForm(this.Form);
	var EE = this.EffectExtent;
	c.setEffectExtent(EE.L, EE.T, EE.R, EE.B);
	return c;
};
ParaDrawing.prototype.SetDrawingPrFromDrawing = function(oAnotherDrawing)
{
	this.Set_DrawingType(oAnotherDrawing.DrawingType);
	var d = oAnotherDrawing.Distance;
	this.Set_PositionH(oAnotherDrawing.PositionH.RelativeFrom, oAnotherDrawing.PositionH.Align, oAnotherDrawing.PositionH.Value, oAnotherDrawing.PositionH.Percent);
	this.Set_PositionV(oAnotherDrawing.PositionV.RelativeFrom, oAnotherDrawing.PositionV.Align, oAnotherDrawing.PositionV.Value, oAnotherDrawing.PositionV.Percent);
	this.Set_Distance(d.L, d.T, d.R, d.B);
	this.Set_AllowOverlap(oAnotherDrawing.AllowOverlap);
	this.Set_WrappingType(oAnotherDrawing.wrappingType);
	this.Set_BehindDoc(oAnotherDrawing.behindDoc);
	this.docPr.setFromOther(oAnotherDrawing.docPr);
};
ParaDrawing.prototype.OnContentReDraw = function()
{
	if (this.Parent && this.Parent.Parent)
		this.Parent.Parent.OnContentReDraw(this.PageNum, this.PageNum);
};
ParaDrawing.prototype.getBase64Img = function()
{
	if (AscCommon.isRealObject(this.GraphicObj) && typeof  this.GraphicObj.getBase64Img === "function")
		return this.GraphicObj.getBase64Img();
	return null;
};
ParaDrawing.prototype.isPointInObject = function(x, y, pageIndex)
{
	if (this.pageIndex === pageIndex)
	{
		if (AscCommon.isRealObject(this.GraphicObj))
		{
			var hit         = (typeof  this.GraphicObj.hit === "function") ? this.GraphicObj.hit(x, y) : false;
			var hit_to_text = (typeof  this.GraphicObj.hitToTextRect === "function") ? this.GraphicObj.hitToTextRect(x, y) : false;
			return hit || hit_to_text;
		}
	}
	return false;
};
ParaDrawing.prototype.RestartSpellCheck = function()
{
	this.GraphicObj && this.GraphicObj.RestartSpellCheck && this.GraphicObj.RestartSpellCheck();
};
/**
 * Проверяем является ли данная автофигура формулой в старом формате
 * @returns {boolean}
 */
ParaDrawing.prototype.IsMathEquation = function()
{
	if (undefined !== this.ParaMath && null !== this.ParaMath)
		return true;

	return false;
};
/**
 * Конвертируем данную автофигуру, представляющую формулу в старом формате, в новый формат
 * @param {boolean} isUpdatePos - обновлять или нет позицию документа
 */
ParaDrawing.prototype.ConvertToMath = function(isUpdatePos)
{
	if (!this.IsMathEquation())
		return;

	var oParagraph = this.GetParagraph();
	if (!oParagraph)
		return;

	var oLogicDocument = oParagraph.GetLogicDocument();
	if (!oLogicDocument)
		return;

	var oParaContentPos = oParagraph.GetPosByDrawing(this.GetId());
	if (!oParaContentPos)
		return;

	var nDepth         = oParaContentPos.GetDepth();
	var nTopElementPos = oParaContentPos.Get(0);
	var nBotElementPos = oParaContentPos.Get(nDepth);
	var oTopElement    = oParagraph.Content[nTopElementPos];

	// Уменьшаем глубину на 1, чтобы получить позицию родительского класса
	var oRunPos = oParaContentPos.Copy();
	oRunPos.DecreaseDepth(1);

	var oRun = oParagraph.Get_ElementByPos(oRunPos);

	if (!oTopElement || !oTopElement.Content || !(oRun instanceof ParaRun))
		return;

	// Коректируем формулу после конвертации
	this.ParaMath.Correct_AfterConvertFromEquation();
	this.ParaMath.ProcessingOldEquationConvert();

	// Сначала удаляем Drawing из рана
	oRun.RemoveFromContent(nBotElementPos, 1);

	// TODO: Тут возможно лучше взять настройки предыдущего элемента, но пока просто удалим самое неприятное
	// свойство.
	if (true === oRun.IsEmpty())
		oRun.Set_Position(undefined);

	// Теперь разделяем параграф по заданной позиции и добавляем туда новую формулу.
	var oRightElement = oTopElement.Split(oParaContentPos, 1);
	oParagraph.AddToContent(nTopElementPos + 1, oRightElement);
	oParagraph.AddToContent(nTopElementPos + 1, this.ParaMath);
	oParagraph.Correct_Content(nTopElementPos, nTopElementPos + 2);

	if (isUpdatePos)
	{
		oRightElement.MoveCursorToStartPos();
		oParagraph.CurPos.ContentPos = nTopElementPos + 2;
		oParagraph.Document_SetThisElementCurrent(false);
	}
};
ParaDrawing.prototype.GetRevisionsChangeElement = function(SearchEngine)
{
	if (this.GraphicObj && this.GraphicObj.GetRevisionsChangeElement)
		this.GraphicObj.GetRevisionsChangeElement(SearchEngine);
};
ParaDrawing.prototype.Get_ObjectType = function()
{
	if (this.GraphicObj)
		return this.GraphicObj.getObjectType();

	return AscDFH.historyitem_type_Drawing;
};
ParaDrawing.prototype.getObjectType = ParaDrawing.prototype.Get_ObjectType;

ParaDrawing.prototype.GetAllContentControls = function(arrContentControls)
{
	if(this.GraphicObj)
	{
		this.GraphicObj.GetAllContentControls(arrContentControls);
	}
};
ParaDrawing.prototype.UpdateBookmarks = function(oManager)
{
	var arrDocContents = this.GetAllDocContents();
	for (var nIndex = 0, nCount = arrDocContents.length; nIndex < nCount; ++nIndex)
	{
		arrDocContents[nIndex].UpdateBookmarks(oManager);
	}
};
ParaDrawing.prototype.PreDelete = function()
{
	if(this.bNotPreDelete === true) {
		//TODO: remove
		return;
	}
	var arrDocContents = this.GetAllDocContents();
	for (var nIndex = 0, nCount = arrDocContents.length; nIndex < nCount; ++nIndex)
	{
		arrDocContents[nIndex].PreDelete();
		arrDocContents[nIndex].ClearContent(true);
	}
	var oGrObject = this.GraphicObj;
	if(oGrObject && oGrObject.signatureLine)
	{
		var oOldSignature = oGrObject.signatureLine;
		oGrObject.setSignature(null);
		editor && editor.sendEvent("asc_onRemoveSignature", oOldSignature);
		oGrObject.setSignature(oOldSignature);
	}
};
ParaDrawing.prototype.CheckSignatureLineOnAdd = function()
{
	var oGrObject = this.GraphicObj;
	if(oGrObject && oGrObject.signatureLine)
	{
		editor && editor.sendEvent("asc_onAddSignature", oGrObject.signatureLine.id);
	}
};
ParaDrawing.prototype.CheckContentControlEditingLock = function()
{
	const oDocContent = this.GetDocumentContent();
	if(oDocContent && oDocContent.CheckContentControlEditingLock)
	{
		oDocContent.CheckContentControlEditingLock();
	}
};
ParaDrawing.prototype.Document_Is_SelectionLocked = function(CheckType)
{
	if(CheckType === AscCommon.changestype_Drawing_Props)
	{
		this.Lock.Check(this.Get_Id());
	}
};

ParaDrawing.prototype.CheckDeletingLock = function()
{
	let logicDocument = this.LogicDocument;
	let form          = this.GetInnerForm();
	
	if (logicDocument && logicDocument.IsDocumentEditor())
	{
		// Если в форму зайти нельзя, то проверяем можно ли удалять её внутреннюю часть
		if (form && this.IsForm() && !form.CanPlaceCursorInside())
		{
			if (Asc.c_oAscSdtLockType.SdtLocked === form.GetContentControlLock() || Asc.c_oAscSdtLockType.SdtContentLocked === form.GetContentControlLock())
				AscCommon.CollaborativeEditing.Add_CheckLock(true);
			
			return;
		}
		
		if (!this.LogicDocument.CanEdit() || this.LogicDocument.IsFillingFormMode())
		{
			// В формах мы не удаляем ничего, а чистим форму, если нужно
			if (this.IsForm() && form && !form.IsPlaceHolder())
				return;
			
			return AscCommon.CollaborativeEditing.Add_CheckLock(true);
		}
	}
	
	if (form)
		form.SkipSpecialContentControlLock(true);
	
	var arrDocContents = this.GetAllDocContents();
	for (var nIndex = 0, nCount = arrDocContents.length; nIndex < nCount; ++nIndex)
	{
		arrDocContents[nIndex].SetApplyToAll(true);
		arrDocContents[nIndex].Document_Is_SelectionLocked(AscCommon.changestype_Remove);
		arrDocContents[nIndex].SetApplyToAll(false);
	}
	
	if (form)
		form.SkipSpecialContentControlLock(false);
};
ParaDrawing.prototype.GetAllFields = function(isUseSelection, arrFields)
{
	if(this.GraphicObj)
	{
		return this.GraphicObj.GetAllFields(isUseSelection, arrFields);
	}
	return arrFields ? arrFields : [];
};

ParaDrawing.prototype.GetAllSeqFieldsByType = function(sType, aFields)
{
	if(this.GraphicObj)
	{
		return this.GraphicObj.GetAllSeqFieldsByType(sType, aFields);
	}
};
/**
 * Является ли данная автофигура картинкой
 * @returns {boolean}
 */
ParaDrawing.prototype.IsPicture = function()
{
	return (this.GraphicObj.getObjectType() === AscDFH.historyitem_type_ImageShape);
};
/**
 * Получаем объект картинки, если автофигура является картинкой
 * @returns {?CImageShape}
 */
ParaDrawing.prototype.GetPicture = function()
{
	return this.GraphicObj.getObjectType() === AscDFH.historyitem_type_ImageShape ? this.GraphicObj : null;
};
ParaDrawing.prototype.GetPictureUrl = function()
{
	let oPicture = this.GetPicture();
	return (oPicture ? oPicture.getImageUrl() : null);
};
/**
 * Является ли объект фигурой
 * @returns {boolean}
 */
ParaDrawing.prototype.IsShape = function()
{
	return (this.GraphicObj.getObjectType() === AscDFH.historyitem_type_Shape);
};
/**
 * Является ли объект группой
 * @returns {boolean}
 */
ParaDrawing.prototype.IsGroup = function()
{
	return (this.GraphicObj.isGroupObject());
};
ParaDrawing.prototype.IsComparable = function(oDrawing)
{
	if(!oDrawing)
	{
		return false;
	}
	if(!this.docPr.hasSameNameAndId(oDrawing.docPr))
	{
		return false;
	}
	if(!this.GraphicObj || !oDrawing.GraphicObj)
	{
		return false;
	}
	if(this.GraphicObj.getObjectType() !== oDrawing.GraphicObj.getObjectType())
	{
		return false;
	}
	return this.GraphicObj.isComparable(oDrawing.GraphicObj);
};
ParaDrawing.prototype.ToSearchElement = function(oProps)
{
	if (this.IsInline())
		return new AscCommonWord.CSearchTextSpecialGraphicObject();

	return null;
};
ParaDrawing.prototype.IsDrawing = function()
{
	return true;
};
ParaDrawing.prototype.CheckRunContent = function(fCheck)
{
	if(this.GraphicObj)
	{
		this.GraphicObj.checkRunContent(fCheck);
	}
};
/**
 * Класс, описывающий текущее положение параграфа при рассчете позиции автофигуры.
 * @constructor
 */
function CParagraphLayout(X, Y, PageNum, LastItemW, ColumnStartX, ColumnEndX, Left_Margin, Right_Margin, Page_W, Top_Margin, Bottom_Margin, Page_H, MarginH, MarginV, LineTop, ParagraphTop)
{
	this.X             = X;
	this.Y             = Y;
	this.PageNum       = PageNum;
	this.LastItemW     = LastItemW;
	this.ColumnStartX  = ColumnStartX;
	this.ColumnEndX    = ColumnEndX;
	this.Left_Margin   = Left_Margin;
	this.Right_Margin  = Right_Margin;
	this.Page_W        = Page_W;
	this.Top_Margin    = Top_Margin;
	this.Bottom_Margin = Bottom_Margin;
	this.Page_H        = Page_H;
	this.Margin_H      = MarginH;
	this.Margin_V      = MarginV;
	this.LineTop       = LineTop;
	this.ParagraphTop  = ParagraphTop;
}
/**
 * Класс, описывающий позицию автофигуры на странице
 * @constructor
 */
function CAnchorPosition()
{
	// Данный коэффициент говорит нам насколько изменены все относительные позиции
	// Используется в режиме чтения, когда реальный размер страницы, на которой мы рисуем, меньшее, чем тот,
	// что задан в настройках
	this.ScaleFactor   = 1;

	// Рассчитанные координаты
	this.CalcX         = 0;
	this.CalcY         = 0;

	// Данные для Inline-объектов
	this.YOffset       = 0;

	// Данные для Flow-объектов
	this.W             = 0;
	this.H             = 0;
	this.X             = 0;
	this.Y             = 0;
	this.PageNum       = 0;
	this.LastItemW     = 0;
	this.ColumnStartX  = 0;
	this.ColumnEndX    = 0;
	this.Left_Margin   = 0;
	this.Right_Margin  = 0;
	this.Page_W        = 0;
	this.Top_Margin    = 0;
	this.Bottom_Margin = 0;
	this.Page_H        = 0;
	this.Margin_H      = 0;
	this.Margin_V      = 0;
	this.LineTop       = 0;
	this.ParagraphTop  = 0;
	this.Page_X        = 0;
	this.Page_Y        = 0;
}
CAnchorPosition.prototype.SetScaleFactor = function(nScale)
{
	this.ScaleFactor = nScale;
};
CAnchorPosition.prototype.Set = function(W, H, Rot, EffectExtent, YOffset, ParaLayout, PageLimits)
{
	this.W = W;
	this.H = H;
	this.Rot = Rot;
	this.EffectExtentL = EffectExtent.L;
	this.EffectExtentT = EffectExtent.T;
	this.EffectExtentR = EffectExtent.R;
	this.EffectExtentB = EffectExtent.B;

	this.YOffset = YOffset;

	this.X             = ParaLayout.X;
	this.Y             = ParaLayout.Y;
	this.PageNum       = ParaLayout.PageNum;
	this.LastItemW     = ParaLayout.LastItemW;
	this.ColumnStartX  = ParaLayout.ColumnStartX;
	this.ColumnEndX    = ParaLayout.ColumnEndX;
	this.Left_Margin   = ParaLayout.Left_Margin;
	this.Right_Margin  = ParaLayout.Right_Margin;
	this.Page_W        = PageLimits.XLimit - PageLimits.X;// ParaLayout.Page_W;
	this.Top_Margin    = ParaLayout.Top_Margin;
	this.Bottom_Margin = ParaLayout.Bottom_Margin;
	this.Page_H        = PageLimits.YLimit - PageLimits.Y;// ParaLayout.Page_H;
	this.Margin_H      = ParaLayout.Margin_H;
	this.Margin_V      = ParaLayout.Margin_V;
	this.LineTop       = ParaLayout.LineTop;
	this.ParagraphTop  = ParaLayout.ParagraphTop;
	this.Page_X        = PageLimits.X;
	this.Page_Y        = PageLimits.Y;

};
CAnchorPosition.prototype.Calculate_X = function(bInline, RelativeFrom, bAlign, Value, bPercent)
{
	var _W;
	if(AscFormat.checkNormalRotate(this.Rot))
	{
		_W = this.W;
	}
	else
	{
		_W = this.H;
	}
	var Width = _W + this.EffectExtentL + this.EffectExtentR;
	var Shift = this.EffectExtentL + _W / 2.0 - this.W / 2.0;

	if (true === bInline)
	{
		this.CalcX = this.X + Shift;
	}
	else
	{
		var _RelativeFrom = RelativeFrom;
		if (_RelativeFrom === c_oAscRelativeFromH.InsideMargin)
		{
			if (0 === this.PageNum % 2)
				_RelativeFrom = c_oAscRelativeFromH.LeftMargin;
			else
				_RelativeFrom = c_oAscRelativeFromH.RightMargin;
		}
		else if (_RelativeFrom === c_oAscRelativeFromH.OutsideMargin)
		{
			if (0 === this.PageNum % 2)
				_RelativeFrom = c_oAscRelativeFromH.RightMargin;
			else
				_RelativeFrom = c_oAscRelativeFromH.LeftMargin;
		}

		// Вычисляем координату по X
		switch (_RelativeFrom)
		{
			case c_oAscRelativeFromH.Character:
			{
				// Почему то Word при позиционировании относительно символа использует не
				// текущуюю позицию, а позицию предыдущего элемента (именно для этого мы
				// храним параметр LastItemW).

				var _X = this.X - this.LastItemW;

				if (true === bAlign)
				{
					switch (Value)
					{
						case c_oAscAlignH.Center:
						{
							this.CalcX = _X - this.W / 2;
							break;
						}

						case c_oAscAlignH.Inside:
						case c_oAscAlignH.Outside:
						case c_oAscAlignH.Left:
						{
							this.CalcX = _X + Shift;
							break;
						}
						case c_oAscAlignH.Right:
						{
							this.CalcX = _X - this.EffectExtentR - _W / 2.0 - this.W / 2.0;
							break;
						}
					}
				}
				else
				{
					this.CalcX = _X + Value;
				}
				break;
			}

			case c_oAscRelativeFromH.Column:
			{
				if (true === bAlign)
				{
					switch (Value)
					{
						case c_oAscAlignH.Center:
						{
							this.CalcX = (this.ColumnEndX + this.ColumnStartX - Width) / 2.0 + this.EffectExtentL + _W / 2.0 - this.W /2.0;
							break;
						}

						case c_oAscAlignH.Inside:
						case c_oAscAlignH.Outside:
						case c_oAscAlignH.Left:
						{

							this.CalcX = this.ColumnStartX + Shift;
							break;
						}

						case c_oAscAlignH.Right:
						{
							this.CalcX = this.ColumnEndX - this.EffectExtentR - _W / 2.0 - this.W / 2.0;
							break;
						}
					}
				}
				else
				{
					this.CalcX = this.ColumnStartX + Value * this.ScaleFactor;
				}

				break;
			}

			case c_oAscRelativeFromH.LeftMargin:
			{
				if (true === bAlign)
				{
					switch (Value)
					{
						case c_oAscAlignH.Center:
						{
							this.CalcX = (this.Left_Margin - Width) / 2 + Shift;
							break;
						}

						case c_oAscAlignH.Inside:
						case c_oAscAlignH.Outside:
						case c_oAscAlignH.Left:
						{
							this.CalcX = Shift;
							break;
						}

						case c_oAscAlignH.Right:
						{
							this.CalcX = this.Left_Margin - (_W / 2.0 + this.EffectExtentR) - this.W / 2.0;
							break;
						}
					}
				}
				else if (true === bPercent)
				{
					this.CalcX = this.Page_X + this.Left_Margin * Value / 100 + Shift;
				}
				else
				{
					this.CalcX = Value;
				}

				break;
			}

			case c_oAscRelativeFromH.Margin:
			{
				var X_s = this.Page_X + this.Left_Margin;
				var X_e = this.Page_X + this.Page_W - this.Right_Margin;

				if (true === bAlign)
				{
					switch (Value)
					{
						case c_oAscAlignH.Center:
						{
							this.CalcX = (X_e + X_s - Width) / 2 + Shift;
							break;
						}

						case c_oAscAlignH.Inside:
						case c_oAscAlignH.Outside:
						case c_oAscAlignH.Left:
						{
							this.CalcX = X_s + Shift;
							break;
						}

						case c_oAscAlignH.Right:
						{
							this.CalcX = X_e - (_W / 2.0 + this.EffectExtentR) - this.W / 2.0;
							break;
						}
					}
				}
				else if (true === bPercent)
				{
					this.CalcX = X_s + (X_e - X_s) * Value / 100 + Shift;
				}
				else
				{
					this.CalcX = this.Margin_H + Value * this.ScaleFactor;
				}

				break;
			}

			case c_oAscRelativeFromH.Page:
			{
				if (true === bAlign)
				{
					switch (Value)
					{
						case c_oAscAlignH.Center:
						{
							this.CalcX = (this.Page_W - Width) / 2 + Shift;
							break;
						}

						case c_oAscAlignH.Inside:
						case c_oAscAlignH.Outside:
						case c_oAscAlignH.Left:
						{
							this.CalcX = Shift;
							break;
						}

						case c_oAscAlignH.Right:
						{
							this.CalcX = this.Page_W - Width + Shift;
							break;
						}
					}
				}
				else if (true === bPercent)
				{
					this.CalcX = this.Page_X + this.Page_W * Value / 100 + Shift;
				}
				else
				{
					this.CalcX = Value * this.ScaleFactor + this.Page_X;
				}

				break;
			}

			case c_oAscRelativeFromH.RightMargin:
			{
				var X_s = this.Page_X + this.Page_W - this.Right_Margin;
				var X_e = this.Page_X + this.Page_W;

				if (true === bAlign)
				{
					switch (Value)
					{
						case c_oAscAlignH.Center:
						{
							this.CalcX = (X_e + X_s - Width) / 2 + Shift;
							break;
						}

						case c_oAscAlignH.Inside:
						case c_oAscAlignH.Outside:
						case c_oAscAlignH.Left:
						{
							this.CalcX = X_s + Shift;
							break;
						}

						case c_oAscAlignH.Right:
						{
							this.CalcX = X_e - Width + Shift;
							break;
						}
					}
				}
				else if (true === bPercent)
				{
					this.CalcX = X_s + (X_e - X_s) * Value / 100 + Shift;
				}
				else
				{
					this.CalcX = X_s + Value;
				}

				break;
			}
		}
	}

	return this.CalcX;
};
CAnchorPosition.prototype.Calculate_Y = function(bInline, RelativeFrom, bAlign, Value, bPercent)
{
	var _H;
	if(AscFormat.checkNormalRotate(this.Rot))
	{
		_H = this.H;
	}
	else
	{
		_H = this.W;
	}
	var Height = this.EffectExtentB + _H + this.EffectExtentT;
	var Shift = this.EffectExtentT + _H / 2.0 - this.H / 2.0;
	if (true === bInline)
	{
		this.CalcY = this.Y - this.YOffset - Height + Shift;
	}
	else
	{
		// Вычисляем координату по Y
		switch (RelativeFrom)
		{
			case c_oAscRelativeFromV.BottomMargin:
			case c_oAscRelativeFromV.InsideMargin:
			case c_oAscRelativeFromV.OutsideMargin:
			{
				var _Y = this.Page_H - this.Bottom_Margin;

				if (true === bAlign)
				{
					switch (Value)
					{
						case c_oAscAlignV.Bottom:
						case c_oAscAlignV.Outside:
						{
							this.CalcY = this.Page_H - Height + Shift;
							break;
						}
						case c_oAscAlignV.Center:
						{
							this.CalcY = (_Y + this.Page_H - Height) / 2 + Shift;
							break;
						}

						case c_oAscAlignV.Inside:
						case c_oAscAlignV.Top:
						{
							this.CalcY = _Y + Shift;
							break;
						}
					}
				}
				else if (true === bPercent)
				{
					if (Math.abs(this.Page_Y) > 0.001)
						this.CalcY = this.Margin_V + Shift;
					else
						this.CalcY = _Y + this.Bottom_Margin * Value / 100 + Shift;
				}
				else
				{
					this.CalcY = _Y + Value;
				}

				break;
			}

			case c_oAscRelativeFromV.Line:
			{
				var _Y = this.LineTop;

				if (true === bAlign)
				{
					switch (Value)
					{
						case c_oAscAlignV.Bottom:
						case c_oAscAlignV.Outside:
						{
							this.CalcY = _Y - this.EffectExtentB - Height + Shift;
							break;
						}
						case c_oAscAlignV.Center:
						{
							this.CalcY = _Y - Height / 2 + Shift;
							break;
						}

						case c_oAscAlignV.Inside:
						case c_oAscAlignV.Top:
						{
							this.CalcY = _Y + Shift;
							break;
						}
					}
				}
				else
					this.CalcY = _Y + Value;

				break;
			}

			case c_oAscRelativeFromV.Margin:
			{
				var Y_s = this.Top_Margin + this.Page_Y;
				var Y_e = this.Page_H - this.Bottom_Margin;

				if (true === bAlign)
				{
					switch (Value)
					{
						case c_oAscAlignV.Bottom:
						case c_oAscAlignV.Outside:
						{
							this.CalcY = Y_e - Height + Shift;
							break;
						}
						case c_oAscAlignV.Center:
						{
							this.CalcY = (Y_s + Y_e - Height) / 2 + Shift;
							break;
						}

						case c_oAscAlignV.Inside:
						case c_oAscAlignV.Top:
						{
							this.CalcY = Y_s + Shift;
							break;
						}
					}
				}
				else if (true === bPercent)
				{
					if (Math.abs(this.Page_Y) > 0.001)
						this.CalcY = this.Margin_V + Shift;
					else
						this.CalcY = Y_s + (Y_e - Y_s) * Value / 100 + Shift;
				}
				else
				{
					this.CalcY = this.Margin_V + Value * this.ScaleFactor;
				}

				break;
			}

			case c_oAscRelativeFromV.Page:
			{
				if (true === bAlign)
				{
					switch (Value)
					{
						case c_oAscAlignV.Bottom:
						case c_oAscAlignV.Outside:
						{
							this.CalcY = this.Page_H - Height + Shift;
							break;
						}
						case c_oAscAlignV.Center:
						{
							this.CalcY = (this.Page_H - Height) / 2 + Shift;
							break;
						}

						case c_oAscAlignV.Inside:
						case c_oAscAlignV.Top:
						{
							this.CalcY = Shift;
							break;
						}
					}
				}
				else if (true === bPercent)
				{
					if (Math.abs(this.Page_Y) > 0.001)
						this.CalcY = this.Margin_V + Shift;
					else
						this.CalcY = this.Page_H * Value / 100 + Shift;
				}
				else
				{
					this.CalcY = Value * this.ScaleFactor + this.Page_Y;
				}

				break;
			}

			case c_oAscRelativeFromV.Paragraph:
			{
				// Почему то Word не дает возможности использовать прилегание
				// относительно абзаца, только абсолютные позиции
				var _Y = this.ParagraphTop;

				if (true === bAlign)
					this.CalcY = _Y + Shift;
				else
					this.CalcY = _Y + Value;

				break;
			}

			case c_oAscRelativeFromV.TopMargin:
			{
				var Y_s = 0;
				var Y_e = this.Top_Margin;

				if (true === bAlign)
				{
					switch (Value)
					{
						case c_oAscAlignV.Bottom:
						case c_oAscAlignV.Outside:
						{
							this.CalcY = Y_e - Height + Shift;
							break;
						}
						case c_oAscAlignV.Center:
						{
							this.CalcY = (Y_s + Y_e - Height) / 2 + Shift;
							break;
						}

						case c_oAscAlignV.Inside:
						case c_oAscAlignV.Top:
						{
							this.CalcY = Y_s + Shift;
							break;
						}
					}
				}
				else if (true === bPercent)
				{
					if (Math.abs(this.Page_Y) > 0.001)
						this.CalcY = this.Margin_V + Shift;
					else
						this.CalcY = this.Top_Margin * Value / 100 + Shift;
				}
				else
				{
					this.CalcY = Y_s + Value * this.ScaleFactor;
				}

				break;
			}
		}
	}

	return this.CalcY;
};
CAnchorPosition.prototype.Update_PositionYHeaderFooter = function(TopMarginY, BottomMarginY)
{
	var TopY    = Math.max(this.Page_Y, Math.min(TopMarginY, this.Page_H));
	var BottomY = Math.max(this.Page_Y, Math.min(BottomMarginY, this.Page_H));

	this.Margin_V      = TopY;
	this.Top_Margin    = TopY;
	this.Bottom_Margin = this.Page_H - BottomY;
};
CAnchorPosition.prototype.Correct_Values = function(bInline, PageLimits, AllowOverlap, UseTextWrap, OtherFlowObjects, bCorrect)
{
	if (true != bInline)
	{
		var X_min = PageLimits.X;
		var Y_min = PageLimits.Y;
		var X_max = PageLimits.XLimit;
		var Y_max = PageLimits.YLimit;

		var W = this.W;
		var H = this.H;

		var CurX = this.CalcX;
		var CurY = this.CalcY;

		var bBreak = false;
		while (true != bBreak)
		{
			bBreak = true;
			for (var Index = 0; Index < OtherFlowObjects.length; Index++)
			{
				var Drawing = OtherFlowObjects[Index];
				if (( false === AllowOverlap || false === Drawing.AllowOverlap ) && true === Drawing.Use_TextWrap() && true === UseTextWrap && ( CurX <= Drawing.X + Drawing.W && CurX + W >= Drawing.X && CurY <= Drawing.Y + Drawing.H && CurY + H >= Drawing.Y ))
				{
					// Если убирается справа, размещаем справа от картинки
					if (Drawing.X + Drawing.W < X_max - W - 0.001)
						CurX = Drawing.X + Drawing.W + 0.001;
					else
					{
						CurX = this.CalcX;
						CurY = Drawing.Y + Drawing.H + 0.001;
					}

					bBreak = false;
				}
			}
		}

		// Автофигуры с обтеканием за/перед текстом могут лежать где угодно
		if (true === UseTextWrap && true === bCorrect)
		{
			// Скорректируем рассчитанную позицию, так чтобы объект не выходил за заданные пределы
			var _W, _H;
			if(AscFormat.checkNormalRotate(this.Rot))
			{
				_W = this.W;
				_H = this.H;
			}
			else
			{
				_W = this.H;
				_H = this.W;
			}
			var Right = CurX + this.W / 2.0 + _W / 2.0 + this.EffectExtentR;
			if (Right > X_max)
			{
				CurX -= (Right - X_max);
			}

			var Left =  CurX + this.W / 2.0 - _W / 2.0 - this.EffectExtentR;
			if (Left < X_min)
			{
				CurX += (X_min - Left);
			}

			// Скорректируем рассчитанную позицию, так чтобы объект не выходил за заданные пределы
			var Bottom = CurY + this.H / 2.0 + _H / 2.0 + this.EffectExtentB;
			if (Bottom > Y_max)
			{
				CurY -= (Bottom - Y_max);
			}

			var Top = CurY + this.H / 2.0 - _H / 2.0 - this.EffectExtentT;
			if (Top < Y_min)
			{
				CurY += (Y_min - Top);
			}
		}

		this.CalcX = CurX;
		this.CalcY = CurY;
	}
};
CAnchorPosition.prototype.Calculate_X_Value = function(RelativeFrom)
{
	var Value = 0;
	switch (RelativeFrom)
	{
		case c_oAscRelativeFromH.Character:
		{
			// Почему то Word при позиционировании относительно символа использует не
			// текущуюю позицию, а позицию предыдущего элемента (именно для этого мы
			// храним параметр LastItemW).
			Value = this.CalcX - this.X + this.LastItemW;

			break;
		}

		case c_oAscRelativeFromH.Column:
		{
			Value = this.CalcX - this.ColumnStartX;

			break;
		}

		case c_oAscRelativeFromH.InsideMargin:
		case c_oAscRelativeFromH.LeftMargin:
		case c_oAscRelativeFromH.OutsideMargin:
		{
			Value = this.CalcX;

			break;
		}

		case c_oAscRelativeFromH.Margin:
		{
			Value = this.CalcX - this.Margin_H;

			break;
		}

		case c_oAscRelativeFromH.Page:
		{
			Value = this.CalcX - this.Page_X;

			break;
		}

		case c_oAscRelativeFromH.RightMargin:
		{
			Value = this.CalcX - this.Page_W + this.Right_Margin;

			break;
		}
	}

	return Value;
};
CAnchorPosition.prototype.Calculate_Y_Value = function(RelativeFrom)
{
	var Value = 0;

	switch (RelativeFrom)
	{
		case c_oAscRelativeFromV.BottomMargin:
		case c_oAscRelativeFromV.InsideMargin:
		case c_oAscRelativeFromV.OutsideMargin:
		{
			Value = this.CalcY - this.Page_H + this.Bottom_Margin;

			break;
		}

		case c_oAscRelativeFromV.Line:
		{
			Value = this.CalcY - this.LineTop;

			break;
		}

		case c_oAscRelativeFromV.Margin:
		{
			Value = this.CalcY - this.Margin_V;

			break;
		}

		case c_oAscRelativeFromV.Page:
		{
			Value = this.CalcY - this.Page_Y;

			break;
		}

		case c_oAscRelativeFromV.Paragraph:
		{
			Value = this.CalcY - this.ParagraphTop;

			break;
		}

		case c_oAscRelativeFromV.TopMargin:
		{
			Value = this.CalcY;

			break;
		}
	}

	return Value;
};

//--------------------------------------------------------export----------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].ParaDrawing = ParaDrawing;

window['AscWord'] = window['AscWord'] || {};
window['AscWord'].ParaDrawing = ParaDrawing;


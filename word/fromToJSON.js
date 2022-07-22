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
(function(window, undefined){

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

	function private_PtToMM(pt)
	{
		return 25.4 / 72.0 * pt;
	}
	function private_Twips2MM(twips)
	{
		return 25.4 / 72.0 / 20 * twips;
	}
	function private_GetDrawingDocument()
	{
		return editor.WordControl.m_oLogicDocument.DrawingDocument;
	}
	function private_EMU2MM(EMU)
	{
		return EMU / 36000.0;
	}
	function private_MM2EMU(MM)
	{
		return MM * 36000.0;
	}
	function private_GetLogicDocument()
	{
		return editor.WordControl.m_oLogicDocument;
	}
	function private_GetStyles()
	{
		var oLogicDocument = private_GetLogicDocument();

		return oLogicDocument instanceof AscCommonWord.CDocument ? oLogicDocument.Get_Styles() : oLogicDocument.globalTableStyles;
	}
	function private_MM2Twips(mm)
	{
		return mm / (25.4 / 72.0 / 20);
	}
	function private_Twips2Px(twips)
	{
		return twips * (4 / 3 / 20);
	}
	function private_Px2Twips(px)
	{
		return px / (4 / 3 / 20);
	}
	/**
	 * Check relative pos by document position
	 * @typeofeditors ["CDE"]
	 * @param {Array} firstPos - first doc pos of element
	 * @param {Array} secondPos - second doc pos of element
	 * @return {1 || 0 || - 1}
	 * If returns 1  -> first element placed before second
	 * If returns 0  -> first element placed like second
	 * If returns -1 -> first element placed after second
	 */
	function private_checkRelativePos(firstPos, secondPos)
	{
		for (var nPos = 0, nLen = Math.min(firstPos.length, secondPos.length); nPos < nLen; ++nPos)
		{
			if (!secondPos[nPos] || !firstPos[nPos] || firstPos[nPos].Class !== secondPos[nPos].Class)
				return 1;

			if (firstPos[nPos].Position < secondPos[nPos].Position)
				return 1;
			else if (firstPos[nPos].Position > secondPos[nPos].Position)
				return -1;
		}

		return 0;
	}
	function private_MM2Pt(mm)
	{
		return mm / (25.4 / 72.0);
	}

	function GetRectAlgnStrType(nAlgnType)
	{
		var sAlgnType = undefined;
		switch (nAlgnType) 
		{
			case AscCommon.c_oAscRectAlignType.b:
				sAlgnType = "b";
				break;
			case AscCommon.c_oAscRectAlignType.bl:
				sAlgnType = "bl";
				break;
			case AscCommon.c_oAscRectAlignType.br:
				sAlgnType = "br";
				break;
			case AscCommon.c_oAscRectAlignType.ctr:
				sAlgnType = "ctr";
				break;
			case AscCommon.c_oAscRectAlignType.l:
				sAlgnType = "l";
				break;
			case AscCommon.c_oAscRectAlignType.r:
				sAlgnType = "r";
				break;
			case AscCommon.c_oAscRectAlignType.t:
				sAlgnType = "t";
				break;
			case AscCommon.c_oAscRectAlignType.tl:
				sAlgnType = "tl";
				break;
			case AscCommon.c_oAscRectAlignType.tr:
				sAlgnType = "tr";
				break;
		}

		return sAlgnType;
	}
	function GetRectAlgnNumType(sAlgnType)
	{
		var nAlgnType = undefined;
		switch (sAlgnType) 
		{
			case "b":
				nAlgnType = AscCommon.c_oAscRectAlignType.b;
				break;
			case "bl":
				nAlgnType = AscCommon.c_oAscRectAlignType.bl;
				break;
			case "br":
				nAlgnType = AscCommon.c_oAscRectAlignType.br;
				break;
			case "ctr":
				nAlgnType = AscCommon.c_oAscRectAlignType.ctr;
				break;
			case "l":
				nAlgnType = AscCommon.c_oAscRectAlignType.l;
				break;
			case "r":
				nAlgnType = AscCommon.c_oAscRectAlignType.r;
				break;
			case "t":
				nAlgnType = AscCommon.c_oAscRectAlignType.t;
				break;
			case "tl":
				nAlgnType = AscCommon.c_oAscRectAlignType.tl;
				break;
			case "tr":
				nAlgnType = AscCommon.c_oAscRectAlignType.tr;
				break;
		}

		return nAlgnType;
	}

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// End of private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    function WriterToJSON()
	{
		this.layoutsMap     = {};
		this.mastersMap     = {};
		this.notesMasterMap = {};
		this.themesMap      = {};
	}

    WriterToJSON.prototype.SerHlink = function(oHlink)
	{
		if (!oHlink)
			return oHlink;
		
		return {
			action:         oHlink.action,
			endSnd:         oHlink.endSnd,
			highlightClick: oHlink.highlightClick,
			history:        oHlink.history,
			id:             oHlink.id,
			invalidUrl:     oHlink.invalidUrl,
			tgtFrame:       oHlink.tgtFrame,
			tooltip:        oHlink.tooltip
		}
	};
	WriterToJSON.prototype.SerExternalData = function(oExtData)
	{
		if (!oExtData)
			return oExtData;
		
		return {
			autoUpdate: oExtData.autoUpdate,
			id:         oExtData.id
		}
	};
	WriterToJSON.prototype.SerProtection = function(oProtection)
	{
		if (!oProtection)
			return oProtection;
		
		return {
			chartObject:   oProtection.chartObject,
			data:          oProtection.data,
			formatting:    oProtection.formatting,
			selection:     oProtection.selection,
			userInterface: oProtection.userInterface
		}
	};
	WriterToJSON.prototype.SerPivotSource = function(oPivotSource)
	{
		if (!oPivotSource)
			return oPivotSource;
		
		return {
			fmtId: oPivotSource.fmtId,
			name:  oPivotSource.name
		}
	};
	WriterToJSON.prototype.SerSimplePos = function(oSimplePos)
	{
		if (!oSimplePos)
			return oSimplePos;
		
		return {
			x:   private_MM2EMU(oSimplePos.X),
			y:   private_MM2EMU(oSimplePos.Y),
			use: oSimplePos.Use
		}
	};
	WriterToJSON.prototype.SerPositionV = function(oPosV)
	{
		if (!oPosV)
			return oPosV;
		
		// anchorV
		var sRelFromV = undefined;
		switch (oPosV.RelativeFrom)
		{
			case Asc.c_oAscRelativeFromV.BottomMargin:
				sRelFromV = "bottomMargin";
				break;
			case Asc.c_oAscRelativeFromV.InsideMargin:
				sRelFromV = "insideMargin";
				break;
			case Asc.c_oAscRelativeFromV.Line:
				sRelFromV = "line";
				break;
			case Asc.c_oAscRelativeFromV.Margin:
				sRelFromV = "margin";
				break;
			case Asc.c_oAscRelativeFromV.OutsideMargin:
				sRelFromV = "outsideMargin";
				break;
			case Asc.c_oAscRelativeFromV.Page:
				sRelFromV = "page";
				break;
			case Asc.c_oAscRelativeFromV.Paragraph:
				sRelFromV = "paragraph";
				break;
			case Asc.c_oAscRelativeFromV.TopMargin:
				sRelFromV = "topMargin";
				break;
		}

		// alignV
		var sVerAlign = undefined;
		if (oPosV.Align)
		{
			switch (oPosV.Value)
			{
				case Asc.c_oAscYAlign.Bottom:
					sVerAlign = "bottom";
					break;
				case Asc.c_oAscYAlign.Center:
					sVerAlign = "center";
					break;
				case Asc.c_oAscYAlign.Inline:
					sVerAlign = "inline";
					break;
				case Asc.c_oAscYAlign.Inside:
					sVerAlign = "inside";
					break;
				case Asc.c_oAscYAlign.Outside:
					sVerAlign = "outside";
					break;
				case Asc.c_oAscYAlign.Top:
					sVerAlign = "top";
					break;
			}
		}

		// offset
		var posOffset = oPosV.Align ? undefined : (oPosV.Percent ? oPosV.Value : private_MM2EMU(oPosV.Value));

		return {
			align:        sVerAlign,
			relativeFrom: sRelFromV,
			posOffset:    posOffset,
			percent:      oPosV.Percent
		}
	};
	WriterToJSON.prototype.SerPositionH = function(oPosH)
	{
		if (!oPosH)
			return oPosH;
		
		// anchorH
		var sRelFromH = undefined;
		switch (oPosH.RelativeFrom)
		{
			case Asc.c_oAscRelativeFromH.Character:
				sRelFromH = "character";
				break;
			case Asc.c_oAscRelativeFromH.Column:
				sRelFromH = "column";
				break;
			case Asc.c_oAscRelativeFromH.InsideMargin:
				sRelFromH = "insideMargin";
				break;
			case Asc.c_oAscRelativeFromH.LeftMargin:
				sRelFromH = "leftMargin";
				break;
			case Asc.c_oAscRelativeFromH.Margin:
				sRelFromH = "margin";
				break;
			case Asc.c_oAscRelativeFromH.OutsideMargin:
				sRelFromH = "outsideMargin";
				break;
			case Asc.c_oAscRelativeFromH.Page:
				sRelFromH = "page";
				break;
			case Asc.c_oAscRelativeFromH.RightMargin:
				sRelFromH = "rightMargin";
				break;
		}

		// alignH
		var sHorAlign = undefined;
		if (oPosH.Align)
		{
			switch (oPosH.Value)
			{
				case Asc.c_oAscXAlign.Center:
					sHorAlign = "center";
					break;
				case Asc.c_oAscXAlign.Inside:
					sHorAlign = "inside";
					break;
				case Asc.c_oAscXAlign.Left:
					sHorAlign = "left";
					break;
				case Asc.c_oAscXAlign.Outside:
					sHorAlign = "outside";
					break;
				case Asc.c_oAscXAlign.Right:
					sHorAlign = "right";
					break;
			}
		}

		// offset
		var posOffSet =  oPosH.Align ? undefined : (oPosH.Percent ? oPosH.Value : private_MM2EMU(oPosH.Value));

		return {
			align:        sHorAlign,
			relativeFrom: sRelFromH,
			posOffset:    posOffSet,
			percent:      oPosH.Percent
		}
	};
	WriterToJSON.prototype.SerExtent = function(oExtent)
	{
		if (!oExtent)
			return oExtent;

		return {
			cy: private_MM2EMU(oExtent.H),
			cx: private_MM2EMU(oExtent.W)
		}
	};
	WriterToJSON.prototype.SerEffectExtent = function(oEffExt)
	{
		if (!oEffExt)
			return oEffExt;

		return {
			b: oEffExt.B,
			l: oEffExt.L,
			r: oEffExt.R,
			t: oEffExt.T
		}
	};
	WriterToJSON.prototype.GetWrapStrType = function(nType)
	{
		switch (nType)
		{
			case WRAPPING_TYPE_NONE:
				return "none";
			case WRAPPING_TYPE_SQUARE:
				return "square";
			case WRAPPING_TYPE_THROUGH:
				return "through";
			case WRAPPING_TYPE_TIGHT:
				return "tight";
			case WRAPPING_TYPE_TOP_AND_BOTTOM:
				return "top_and_bottom";
			default:
				return "none";
		}
	};
	WriterToJSON.prototype.SerGraphicObject = function(oGraphicObj, aComplexFieldsToSave)
	{
		if (!oGraphicObj)
			return null;

		var oResultObj = null;

		if (oGraphicObj instanceof AscFormat.CChartSpace)
			oResultObj = this.SerChartSpace(oGraphicObj, aComplexFieldsToSave);
		else if (oGraphicObj instanceof AscFormat.COleObject)
			oResultObj = this.SerOleObject(oGraphicObj, aComplexFieldsToSave);
		else if (oGraphicObj instanceof AscFormat.CImageShape)
			oResultObj = this.SerImage(oGraphicObj, aComplexFieldsToSave);
		else if (oGraphicObj instanceof AscFormat.CShape)
		{
			if (oGraphicObj.constructor.name === "CConnectionShape")
			{
				var oShapeObj = this.SerShape(oGraphicObj, aComplexFieldsToSave);
				oShapeObj.type = "connectShape";

				oResultObj = oShapeObj;
			}
			else
				oResultObj = this.SerShape(oGraphicObj, aComplexFieldsToSave);
		}
		else if (oGraphicObj instanceof AscFormat.CGraphicFrame)
			oResultObj = this.SerGraphicFrame(oGraphicObj);
		else if (oGraphicObj instanceof AscFormat.SmartArt)
			oResultObj = this.SerSmartArt(oGraphicObj);
		else if (oGraphicObj instanceof AscFormat.CLockedCanvas)
			oResultObj = this.SerLockedCanvas(oGraphicObj);
		else if (oGraphicObj instanceof AscFormat.CGroupShape && oGraphicObj.constructor.name === "CGroupShape")
			oResultObj = this.SerGroupShape(oGraphicObj);
		
		oResultObj.id = oGraphicObj.Id;

		return oResultObj;
	};
	WriterToJSON.prototype.SerOleObject = function(oOleObject)
	{
		var oResult  = this.SerImage(oOleObject);

		oResult.appId = oOleObject.m_sApplicationId;
		oResult.data  = oOleObject.m_sData;

		oResult.dxaOrig = private_Px2Twips(oOleObject.m_nPixWidth);
		oResult.dyaOrig = private_Px2Twips(oOleObject.m_nPixHeight);

		oResult.objFile    = oOleObject.m_sObjectFile;
		oResult.oleType    = To_XML_OleObj_Type(oOleObject.m_nOleType);
		oResult.binaryData = oOleObject.m_aBinaryData;
		oResult.mathObj    = this.SerParaMath(oOleObject.m_oMathObject);

		oResult.type = "oleObject";
        return oResult;
	};
	WriterToJSON.prototype.SerLockedCanvas = function(oLockedCanvas)
	{
		var oResult  = this.SerDrawing(oLockedCanvas);
        oResult.type = "lockedCanvas";

        return oResult;
	};
	WriterToJSON.prototype.SerGroupShape = function(oGroupShape)
    {
        var oResult  = this.SerDrawing(oGroupShape);
        oResult.type = "groupShape";

        return oResult;
    };
	WriterToJSON.prototype.SerSmartArt = function(oSmartArt)
	{
		return {
			colorsDef: this.SerColorsDef(oSmartArt.colorsDef),
			dataModel: this.SerData(oSmartArt.dataModel, "diagramData"),
			layoutDef: this.SerLayoutDef(oSmartArt.layoutDef),
			styleDef:  this.SerStyleDef(oSmartArt.styleDef),
			drawing:   this.SerDrawing(oSmartArt.drawing),
			nvGrpSpPr: this.SerUniNvPr(oSmartArt.nvGrpSpPr),
			spPr:      this.SerSpPr(oSmartArt.spPr),
			type:      "smartArt",
			artType:   oSmartArt.type
		}
	};
	WriterToJSON.prototype.SerStyleDef = function(oStyleDef)
	{
		if (!oStyleDef)
			return oStyleDef;

		var aStyleLbl = [];
		for (var nStyle = 0; nStyle < oStyleDef.styleLbl.length; nStyle++)
			aStyleLbl.push(this.SerDefStyleLbl(oStyleDef.styleLbl[nStyle]));

		return {
			catLst:     this.SerCatLst(oStyleDef.catLst),
			desc:       this.SerDesc(oStyleDef.desc, "desc"),
			scene3d:    this.SerScene3d(oStyleDef.scene3d),
			styleLbl:   aStyleLbl,
			minVer:     oStyleDef.minVer,
			title:      this.SerDesc(oStyleDef.title, "diagramTitle"),
			uniqueId:   oStyleDef.uniqueId
		}
	};
	WriterToJSON.prototype.SerDefStyleLbl = function(oDefStyleLbl)
	{
		if (!oDefStyleLbl)
			return oDefStyleLbl;

		return {
			scene3d: this.SerScene3d(oDefStyleLbl.scene3d),
			sp3d:    this.SerSp3D(oDefStyleLbl.sp3d),
			style:   this.SerSpStyle(oDefStyleLbl.style),
			txPr:    this.SerTxPr(oDefStyleLbl.txPr),
			name:    oDefStyleLbl.name
		}
	};
	WriterToJSON.prototype.SerSp3D = function(oSp3D)
	{
		if (!oSp3D)
			return oSp3D;

		return {
			bevelB:       this.SerBevel(oSp3D.bevelB),
			bevelT:       this.SerBevel(oSp3D.bevelT),
			contourClr:   this.SerContourClr(oSp3D.contourClr),
			extrusionClr: this.SerExtrusionClr(oSp3D.extrusionClr),
			contourW:     oSp3D.contourW,
			extrusionH:   oSp3D.extrusionH,
			prstMaterial: To_XML_ST_PresetMaterialType(oSp3D.prstMaterial),
			z:            oSp3D.z
		}
	};
	WriterToJSON.prototype.SerExtrusionClr = function(oExtrusionClr)
	{
		if (!oExtrusionClr)
			return oExtrusionClr;

		return {
			color: this.SerColor(oExtrusionClr.color)
		}
	};
	WriterToJSON.prototype.SerContourClr = function(oContourClr)
	{
		if (!oContourClr)
			return oContourClr;

		return {
			color: this.SerColor(oContourClr.color)
		}
	};
	WriterToJSON.prototype.SerBevel = function(oBevel)
	{
		if (!oBevel)
			return oBevel;

		return {
			h:    oBevel.h,
			prst: To_XML_ST_BevelPresetType(oBevel.prst),
			w:    oBevel.w
		}
	};
	WriterToJSON.prototype.SerScene3d = function(oScene3d)
	{
		if (!oScene3d)
			return oScene3d;

		return {
			backdrop: this.SerBackdrop(oScene3d.backdrop),
			camera:   this.SerCamera(oScene3d.camera),
			lightRig: this.SerLightRig(oScene3d.lightRig)
		}
	};
	WriterToJSON.prototype.SerLightRig = function(oLightRig)
	{
		if (!oLightRig)
			return oLightRig;

		return {
			rot: this.SerCameraRot(oLightRig.rot),
			dir: To_XML_ST_LightRigDirection(oLightRig.dir),
			rig: To_XML_ST_LightRigType(oLightRig.rig)
		}
	};
	WriterToJSON.prototype.SerCamera = function(oCamera)
	{
		if (!oCamera)
			return oCamera;

		return {
			rot:  this.SerCameraRot(oCamera.rot),
			fov:  oCamera.fov,
			prst: To_XML_ST_PresetCameraType(oCamera.prst),
			zoom: oCamera.zoom
		}
	};
	WriterToJSON.prototype.SerCameraRot = function(oCameraRot)
	{
		if (!oCameraRot)
			return oCameraRot;

		return {
			lat: oCameraRot.lat,
			lon: oCameraRot.lon,
			rev: oCameraRot.rev
		}
	};
	WriterToJSON.prototype.SerBackdrop = function(oBackdrop)
	{
		if (!oBackdrop)
			return oBackdrop;

		return {
			anchor: { // function BackdropAnchor() 
				x: oBackdrop.anchor.x,
				y: oBackdrop.anchor.y,
				z: oBackdrop.anchor.z
			},
			norm: { // function BackdropNorm() 
				dx: oBackdrop.norm.dx,
				dy: oBackdrop.norm.dy,
				dz: oBackdrop.norm.dz
			},
			up: { // function BackdropUp() 
				dx: oBackdrop.up.dx,
				dy: oBackdrop.up.dy,
				dz: oBackdrop.up.dz
			}
		}
	};
	WriterToJSON.prototype.SerLayoutDef = function(oLayoutDef)
	{
		if (!oLayoutDef)
			return oLayoutDef;

		return {
			catLst:     this.SerCatLst(oLayoutDef.catLst),
			clrData:    this.SerData(oLayoutDef.clrData, "clrData"),
			defStyle:   oLayoutDef.defStyle,
			desc:       this.SerDesc(oLayoutDef.desc, "desc"),
			layoutNode: this.SerLayoutNode(oLayoutDef.layoutNode),
			minVer:     oLayoutDef.minVer,
			sampData:   this.SerData(oLayoutDef.sampData, "sampData"),
			styleData:  this.SerData(oLayoutDef.styleData, "styleData"),
			title:      this.SerDesc(oLayoutDef.title, "diagramTitle"),
			uniqueId:   oLayoutDef.uniqueId
		}
	};
	WriterToJSON.prototype.SerLayoutNode = function(oLayoutNode)
	{
		if (!oLayoutNode)
			return oLayoutNode;

		var aItems = [];
		for (var nItem = 0; nItem < oLayoutNode.list.length; nItem++)
			aItems.push(this.SerNodeItem(oLayoutNode.list[nItem]));

		return {
			chOrder:  To_XML_ST_ChildOrderType(oLayoutNode.chOrder),
			moveWith: oLayoutNode.moveWith,
			name:     oLayoutNode.name,
			styleLbl: oLayoutNode.styleLbl,
			list:     aItems,
			objType:  "layoutNode"
		}
	};
	WriterToJSON.prototype.SerNodeItem = function(oNodeItem)
	{
		if (oNodeItem instanceof AscFormat.Alg)
			return this.SerAlg(oNodeItem);
		if (oNodeItem instanceof AscFormat.Choose)
			return this.SerChoose(oNodeItem);
		if (oNodeItem instanceof AscFormat.ConstrLst)
			return this.SerConstrLst(oNodeItem);
		if (oNodeItem instanceof AscFormat.ForEach)
			return this.SerForEach(oNodeItem);
		if (oNodeItem instanceof AscFormat.LayoutNode)
			return this.SerLayoutNode(oNodeItem);
		if (oNodeItem instanceof AscFormat.PresOf)
			return this.SerPresOf(oNodeItem);
		if (oNodeItem instanceof AscFormat.RuleLst)
			return this.SerRuleLst(oNodeItem);
		if (oNodeItem instanceof AscFormat.SShape)
			return this.SerSShape(oNodeItem);
		if (oNodeItem instanceof AscFormat.VarLst)
			return this.SerVarLst(oNodeItem);
	};
	WriterToJSON.prototype.SerSShape = function(oSShape)
	{
		return {
			blip:      oSShape.blip,
			blipPhldr: oSShape.blipPhldr,
			hideGeom:  oSShape.hideGeom,
			lkTxEntry: oSShape.lkTxEntry,
			rot:       oSShape.rot,
			type:      oSShape.type,
			zOrderOff: oSShape.zOrderOff,
			adjLst:    this.SerAdjLst(oSShape.adjLst),
			objType:   "sshape"
		}
	};
	WriterToJSON.prototype.SerAdjLst = function(oSerAdjLst)
	{
		var aSerAdjLst = [];
		for (var nRule = 0; nRule < oSerAdjLst.list.length; nRule++)
			aSerAdjLst.push(this.SerAdj(oSerAdjLst.list[nRule]));

		return {
			list: aSerAdjLst
		};
	};
	WriterToJSON.prototype.SerAdj = function(oSerAdj)
	{
		return {
			idx: oSerAdj.idx,
			val: oSerAdj.val
		};
	};
	WriterToJSON.prototype.SerRuleLst = function(oRuleLst)
	{
		var aRuleLst = [];
		for (var nRule = 0; nRule < oRuleLst.list.length; nRule++)
			aRuleLst.push(this.SerRule(oRuleLst.list[nRule]));

		return {
			list:    aRuleLst,
			objType: "ruleLst"
		};
	};
	WriterToJSON.prototype.SerRule = function(oRule)
	{
		return {
			fact:    isNaN(oRule.fact) ? "NaN" : oRule.fact,
			for:     To_XML_ST_ConstraintRelationship(oRule.for),
			forName: oRule.forName,
			max:     isNaN(oRule.max) ? "NaN" : oRule.max,
			ptType:  this.SerBaseFormatObj(oRule.ptType, "element"),
			type:    To_XML_ST_ConstraintType(oRule.type),
			val:     oRule.val
		}
	};
	WriterToJSON.prototype.SerPresOf = function(oPresOf)
	{
		var oResult        = this.SerIteratorAttributes(oPresOf);
		oResult.objType = "presOf";

		return oResult;
	};
	WriterToJSON.prototype.SerForEach = function(oForEach)
	{
		var aNodeItems = [];
		for (var nItem = 0; nItem < oForEach.list.length; nItem++)
			aNodeItems.push(this.SerNodeItem(oForEach.list[nItem]));

		var oResult = this.SerIteratorAttributes(oForEach);

		oResult.list    = aNodeItems;
		oResult.name    = oForEach.name;
		oResult.ref     = oForEach.ref;
		oResult.objType = "forEach";

		return oResult;
	};
	WriterToJSON.prototype.SerChoose = function(oChoose)
	{
		var aIf = [];
		for (var nIf = 0; nIf < oChoose.if.length; nIf++)
			aIf.push(this.SerIf(oChoose.if[nIf]));

		return {
			name:    oChoose.name,
			if:      aIf,
			else:    this.SerElse(oChoose.else),
			objType: "choose"
		}
	};
	WriterToJSON.prototype.SerIf = function(oIf)
	{
		var aNodeItems = [];
		for (var nItem = 0; nItem < oIf.list.length; nItem++)
			aNodeItems.push(this.SerNodeItem(oIf.list[nItem]));

		var oResult = this.SerIteratorAttributes(oIf);

		oResult.list = aNodeItems;
		oResult.arg  = oIf.arg;
		oResult.func = To_XML_ST_FunctionType(oIf.func);
		oResult.name = oIf.name;
		oResult.op   = To_XML_ST_FunctionOperator(oIf.op);
		oResult.val  = oIf.val;

		return oResult;
	};
	WriterToJSON.prototype.SerElse = function(oElse)
	{
		var aNodeItems = [];
		for (var nItem = 0; nItem < oElse.list.length; nItem++)
			aNodeItems.push(this.SerNodeItem(oElse.list[nItem]));

		return {
			list: aNodeItems,
			name: oElse.name
		}
	};
	WriterToJSON.prototype.SerIteratorAttributes = function(oIteratorAttributes)
	{
		var aAxis = [];
		for (var nAxie = 0; nAxie < oIteratorAttributes.axis.length; nAxie++)
			aAxis.push(this.SerBaseFormatObj(oIteratorAttributes.axis[nAxie], "axie"));
		
		var aCnt = [];
		for (var nCnt = 0; nCnt < oIteratorAttributes.cnt.length; nCnt++)
			aCnt.push(oIteratorAttributes.cnt[nCnt]);

		var aHideLastTrans = [];
		for (var nItem = 0; nItem < oIteratorAttributes.hideLastTrans.length; nItem++)
			aHideLastTrans.push(oIteratorAttributes.hideLastTrans[nItem]);

		var aPtTypes = [];
		for (var nPtType = 0; nPtType < oIteratorAttributes.ptType.length; nPtType++)
			aPtTypes.push(this.SerBaseFormatObj(oIteratorAttributes.ptType[nPtType], "element"));

		var aSt = [];
		for (var nSt = 0; nSt < oIteratorAttributes.st.length; nSt++)
			aSt.push(oIteratorAttributes.st[nSt]);
		
		var aStep = [];
		for (var nStep = 0; nStep < oIteratorAttributes.step.length; nStep++)
			aStep.push(oIteratorAttributes.step[nStep]);
			
		return {
			axis:          aAxis,
			cnt:           aCnt,
			hideLastTrans: aHideLastTrans,
			ptType:        aPtTypes,
			st:            aSt,
			step:          aStep
		}
	};
	WriterToJSON.prototype.SerConstrLst = function(oConstrLst)
	{
		var aContsLst = [];
		for (var nItem = 0; nItem < oConstrLst.list.length; nItem++)
			aContsLst.push(this.SerConstr(oConstrLst.list[nItem]));

		return {
			list:    aContsLst,
			objType: "constrLst"
		}
	};
	WriterToJSON.prototype.SerConstr = function(oConstr)
	{
		return {
			fact:       oConstr.fact,
			for:        To_XML_ST_ConstraintRelationship(oConstr.for),
			forName:    oConstr.forName,
			op:         To_XML_ST_BoolOperator(oConstr.op),
			ptType:     this.SerBaseFormatObj(oConstr.ptType, "element"),
			refFor:     To_XML_ST_ConstraintRelationship(oConstr.refFor),
			refForName: oConstr.refForName,
			refPtType:  this.SerBaseFormatObj(oConstr.refPtType, "element"),
			refType:    To_XML_ST_ConstraintType(oConstr.refType),
			type:       To_XML_ST_ConstraintType(oConstr.type),
			val:        oConstr.val,
			objType:    "constrLst"
		}
	};
	WriterToJSON.prototype.SerAlg = function(oAlg)
	{
		var aParams = [];
		for (var nParam = 0; nParam < oAlg.param.length; nParam++)
			aParams.push(this.SerAlgParam(oAlg.param[nParam]));

		return {
			param:   aParams,
			rev:     oAlg.rev,
			type:    To_XML_ST_AlgorithmType(oAlg.type),
			objType: "alg"
		}
	};
	WriterToJSON.prototype.SerAlgParam = function(oAlgParam)
	{
		return {
			type: To_XML_ST_ParameterId(oAlgParam.type),
			val:  oAlgParam.val
		}
	};
	WriterToJSON.prototype.SerData = function(oData, sType)
	{
		if (!oData)
			return oData;

		var oResult = {
			dataModel: this.SerDataModel(oData.dataModel),
			type:      sType
		}
		
		switch (sType)
		{
			case "diagramData":
				break;
			case "styleData":
			case "clrData":
			case "sampData":
				oResult.useDef = oData.useDef;
				break;
		} 

		return oResult;
	};
	WriterToJSON.prototype.SerDataModel = function(oDataModel)
	{
		if (!oDataModel)
			return oDataModel;

		return {
			bg:     this.SerBgFormat(oDataModel.bg),
			cxnLst: this.SerCxnLst(oDataModel.cxnLst),
			ptLst:  this.SerPtLst(oDataModel.ptLst),
			whole:  this.SerWhole(oDataModel.whole)
		}
	};
	WriterToJSON.prototype.SerWhole = function(oWhole)
	{
		if (!oWhole)
			return oWhole;

		return {
			effectLst: oWhole.effect ? this.SerEffectLst(oWhole.effect.EffectLst) : oWhole.effect,
			effectDag: oWhole.effect ? this.SerEffectDag(oWhole.effect.EffectDag) : oWhole.effect,
			ln: this.SerLn(oWhole.ln)
		}
	};
	WriterToJSON.prototype.SerPtLst = function(oPtLst)
	{
		if (!oPtLst)
			return oPtLst;

		var aPtLst = [];
		for (var nItem = 0; nItem < oPtLst.list.length; nItem++)
			aPtLst.push(this.SerPt(oPtLst.list[nItem]));
			
		return {
			list: aPtLst
		}
	};
	WriterToJSON.prototype.SerPt = function(oPt)
	{
		return {
			prSet:   this.SerPtPrSet(oPt.prSet),
			spPr:    this.SerSpPr(oPt.spPr),
			t:       this.SerTxPr(oPt.t),
			cxnId:   oPt.cxnId,
			modelId: oPt.modelId,
			type:    To_XML_ST_PtType(oPt.type)
		}
	};
	WriterToJSON.prototype.SerPtPrSet = function(oPtPrSet)
	{
		if (!oPtPrSet)
			return oPtPrSet;

		return {
			coherent3DOff:        oPtPrSet.coherent3DOff,
			csCatId:              oPtPrSet.csCatId,
			csTypeId:             oPtPrSet.csTypeId,
			custAng:              oPtPrSet.custAng,
			custFlipHor:          oPtPrSet.custFlipHor,
			custFlipVert:         oPtPrSet.custFlipVert,
			custLinFactNeighborX: oPtPrSet.custLinFactNeighborX,
			custLinFactNeighborY: oPtPrSet.custLinFactNeighborY,
			custLinFactX:         oPtPrSet.custLinFactX,
			custLinFactY:         oPtPrSet.custLinFactY,
			custRadScaleInc:      oPtPrSet.custRadScaleInc,
			custRadScaleRad:      oPtPrSet.custRadScaleRad,
			custScaleX:           oPtPrSet.custScaleX,
			custScaleY:           oPtPrSet.custScaleY,
			custSzX:              oPtPrSet.custSzX,
			custSzY:              oPtPrSet.custSzY,
			custT:                oPtPrSet.custT,
			loCatId:              oPtPrSet.loCatId,
			loTypeId:             oPtPrSet.loTypeId,
			phldr:                oPtPrSet.phldr,
			phldrT:               oPtPrSet.phldrT,
			presAssocID:          oPtPrSet.presAssocID,
			presLayoutVars:       this.SerNodeItem(oPtPrSet.presLayoutVars),
			presName:             oPtPrSet.presName,
			presStyleCnt:         oPtPrSet.presStyleCnt,
			presStyleIdx:         oPtPrSet.presStyleIdx,
			presStyleLbl:         oPtPrSet.presStyleLbl,
			qsCatId:              oPtPrSet.qsCatId,
			qsTypeId:             oPtPrSet.qsTypeId,
			style:                this.SerData(oPtPrSet.style, "styleData")
		}
	};
	WriterToJSON.prototype.SerVarLst = function(oPresLayoutVars)
	{
		if (!oPresLayoutVars)
			return oPresLayoutVars;

		return {
			animLvl:       this.SerBaseFormatObj(oPresLayoutVars.animLvl, "animLvl"),
			animOne:       this.SerBaseFormatObj(oPresLayoutVars.animOne, "animOne"),
			bulletEnabled: this.SerBaseFormatObj(oPresLayoutVars.bulletEnabled, "bulletEnabled"),
			chMax:         this.SerBaseFormatObj(oPresLayoutVars.chMax, "chMax"),
			chPref:        this.SerBaseFormatObj(oPresLayoutVars.chPref, "chPref"),
			dir:           this.SerBaseFormatObj(oPresLayoutVars.dir, "dir"),
			hierBranch:    this.SerBaseFormatObj(oPresLayoutVars.hierBranch, "hierBranch"),
			orgChart:      this.SerBaseFormatObj(oPresLayoutVars.orgChart, "orgChart"),
			resizeHandles: this.SerBaseFormatObj(oPresLayoutVars.resizeHandles, "resizeHandles"),
			objType:       "varLst"
		}
	};
	WriterToJSON.prototype.SerBaseFormatObj = function(oBaseFormatObj, sType)
	{
		if (!oBaseFormatObj)
			return oBaseFormatObj;

		var oResult = {
			type: sType
		};
		switch (sType)
		{
			case "animLvl":
				oResult.val = To_XML_ST_AnimLvlStr(oBaseFormatObj.val);
				break;
			case "animOne":
				oResult.val = To_XML_ST_AnimOneStr(oBaseFormatObj.val);
				break;
			case "bulletEnabled":
			case "chMax":
			case "chPref":
			case "orgChart":
				oResult.val = oBaseFormatObj.val;
				break;
			case "dir":
				oResult.val = To_XML_ST_Direction(oBaseFormatObj.val)
				break;
			case "hierBranch":
				oResult.val = To_XML_ST_HierBranchStyle(oBaseFormatObj.val);
				break;
			case "resizeHandles":
				oResult.val = To_XML_ST_ResizeHandlesStr(oBaseFormatObj.val);
				break;
			case "element":
				oResult.val = To_XML_ST_ElementType(oBaseFormatObj.val);
				break;
			case "axie":
				oResult.val = To_XML_ST_AxisType(oBaseFormatObj.val);
				break;
		} 

		return oResult;
	};
	WriterToJSON.prototype.SerCxnLst = function(oCxnLst)
	{
		if (!oCxnLst)
			return oCxnLst;

		var aCxnLst = [];
		for (var nItem = 0; nItem < oCxnLst.list.length; nItem++)
			aCxnLst.push(this.SerCxn(oCxnLst.list[nItem]));
			
		return {
			list: aCxnLst
		}
	};
	WriterToJSON.prototype.SerCxn = function(oCxn)
	{
		return {
			destId:     oCxn.destId,
			destOrd:    oCxn.destOrd,
			modelId:    oCxn.modelId,
			parTransId: oCxn.parTransId,
			presId:     oCxn.presId,
			sibTransId: oCxn.sibTransId,
			srcId:      oCxn.srcId,
			srcOrd:     oCxn.srcOrd,
			type:       oCxn.type
		}
	};
	WriterToJSON.prototype.SerBgFormat = function(oBgFormat)
	{
		if (!oBgFormat)
			return oBgFormat;

		return {
			effectLst: oBgFormat.effect ? this.SerEffectLst(oBgFormat.effect.EffectLst) : oBgFormat.effect,
			effectDag: oBgFormat.effect ? this.SerEffectDag(oBgFormat.effect.EffectDag) : oBgFormat.effect,
			fill: this.SerFill(oBgFormat.fill)
		}
	};
	WriterToJSON.prototype.SerColorsDef = function(oColorsDef)
	{
		if (!oColorsDef)
			return oColorsDef;

		var aStyleLbl = [];
		for (var nStyle = 0; nStyle < oColorsDef.styleLbl.length; nStyle++)
			aStyleLbl.push(this.SerColorDefStyleLbl(oColorsDef.styleLbl[nStyle]));

		return {
			catLst:   this.SerCatLst(oColorsDef.catLst),
			desc:     this.SerDesc(oColorsDef.desc, "desc"),
			styleLbl: aStyleLbl,
			title:    this.SerDesc(oColorsDef.title, "diagramTitle"),
			minVer:   oColorsDef.minVer,
			uniqueId: oColorsDef.uniqueId
		}
	};
	WriterToJSON.prototype.SerColorDefStyleLbl = function(oStyleLbl)
	{
		if (!oStyleLbl)
			return oStyleLbl;

		return {
			effectClrLst:   this.SerClrLst(oStyleLbl.effectClrLst, "effectClrLst"),
			fillClrLst:     this.SerClrLst(oStyleLbl.fillClrLst, "fillClrLst"),
			linClrLst:      this.SerClrLst(oStyleLbl.linClrLst, "linClrLst"),
			txEffectClrLst: this.SerClrLst(oStyleLbl.txEffectClrLst, "txEffectClrLst"),
			txFillClrLst:   this.SerClrLst(oStyleLbl.txFillClrLst, "txFillClrLst"),
			txLinClrLst:    this.SerClrLst(oStyleLbl.txLinClrLst, "txLinClrLst"),
			name:           oStyleLbl.name
		}
	};
	WriterToJSON.prototype.SerClrLst = function(oClrLst, sType)
	{
		if (!oClrLst)
			return oClrLst;

		var aColors = [];
		for (var nColor = 0; nColor < oClrLst.list.length; nColor++)
			aColors.push(this.SerColor(oClrLst.list[nColor]));

		return {
			list:   aColors,
			hueDir: To_XML_ST_HueDir(oClrLst.hueDir),
			meth:   To_XML_ST_ClrAppMethod(oClrLst.meth),
			type:   sType
		}
	};
	WriterToJSON.prototype.SerDesc = function(oDesc, sType)
	{
		if (!oDesc)
			return oDesc;

		return {
			lang: oDesc.lang,
			val:  oDesc.val,
			type: sType
		}
	};
	WriterToJSON.prototype.SerCatLst = function(oCatLst)
	{
		if (!oCatLst)
			return oCatLst;

		var aCatLst = [];
		for (var nCat = 0; nCat < oCatLst.list.length; nCat++)
			aCatLst.push(this.SerSCat(oCatLst.list[nCat]));

		return {
			list: aCatLst
		}
	};
	WriterToJSON.prototype.SerSCat = function(oSCat)
	{
		return {
			pri:  oSCat.pri,
			type: oSCat.type
		}
	};
	WriterToJSON.prototype.SerDrawing = function(oDrawing)
	{
		var aSpTree = [];
		for (var nDrawing = 0; nDrawing < oDrawing.spTree.length; nDrawing++)
			aSpTree.push(this.SerGraphicObject(oDrawing.spTree[nDrawing]));

		return {
			spTree: aSpTree,
			spPr:   this.SerSpPr(oDrawing.spPr),
			type:   "drawing"
		}
	};
	WriterToJSON.prototype.SerGraphicFrame = function(oGraphicFrame)
	{
		return {
			graphic:          this.SerTable(oGraphicFrame.graphicObject),
			nvGraphicFramePr: this.SerUniNvPr(oGraphicFrame.nvGraphicFramePr),
			spPr:             this.SerSpPr(oGraphicFrame.spPr),
			type:             "graphicFrame"
		}
	};
	WriterToJSON.prototype.SerCNvPr = function(oCNvPr)
	{
		if (!oCNvPr)
			return null;
		
		return {
			hlinkClick: this.SerHlink(oCNvPr.hlinkClick),
			hlinkHover: this.SerHlink(oCNvPr.hlinkHover),
			descr:      oCNvPr.descr,
			hidden:     oCNvPr.isHidden,
			id:         oCNvPr.id,
			name:       oCNvPr.name,
			title:      oCNvPr.title
		}
	};
	WriterToJSON.prototype.SerNvPr = function(oNvPr)
	{
		if (!oNvPr)
			return oNvPr;

		return {
			isPhoto:   oNvPr.isPhoto,
			ph:        this.SerPlaceholder(oNvPr.ph),
			unimedia:  this.SerUniMedia(oNvPr.unimedia),
			userDrawn: oNvPr.userDrawn
		}
	};
	WriterToJSON.prototype.SerPlaceholder = function(oPh)
	{
		if (!oPh)
			return oPh;

		// orient
		var sOrient = typeof(oPh.orient) === "number" ? (oPh.orient === 1 ? "vert" : "horz") : null;
		
		// size
		var sPhSz = null;
		switch (oPh.sz)
		{
			case 0:
				sPhSz = "full";
				break;
			case 1:
				sPhSz = "half";
				break;
			case 2:
				sPhSz = "quarter";
				break;
		}

		return {
			hasCustomPrompt: oPh.hasCustomPrompt,
			idx:             oPh.idx,
			orient:          sOrient,
			sz:              sPhSz,
			type:            this.GetStrPhType(oPh.type)
		}
	};
	WriterToJSON.prototype.SerUniMedia = function(oUniMedia)
	{
		if (!oUniMedia)
			return oUniMedia;

		return {
			type:  oUniMedia.type,
			media: oUniMedia.media
		}
	};
	WriterToJSON.prototype.GetStrPhType = function(nType)
	{
		switch (nType)
		{
			case AscFormat.phType_body:
				return "body";
			case AscFormat.phType_chart:
				return "sPhType";
			case AscFormat.phType_clipArt:
				return "clipArt";
			case AscFormat.phType_ctrTitle:
				return "ctrTitle";
			case AscFormat.phType_dgm:
				return "dgm";
			case AscFormat.phType_dt:
				return "dt";
			case AscFormat.phType_ftr:
				return "ftr";
			case AscFormat.phType_hdr:
				return "hdr";
			case AscFormat.phType_media:
				return "media";
			case AscFormat.phType_obj:
				return "obj";
			case AscFormat.phType_pic:
				return "pic";
			case AscFormat.phType_sldImg:
				return "sldImg";
			case AscFormat.phType_sldNum:
				return "sldNum";
			case AscFormat.phType_subTitle:
				return "subTitle";
			case AscFormat.phType_tbl:
				return "tbl";
			case AscFormat.phType_title:
				return "title";
		}
	};
	WriterToJSON.prototype.SerNumLit = function(oNumLit)
	{
		if (!oNumLit)
			return null;

		var arrResultPts = [];
		for (var nItem = 0; nItem < oNumLit.pts.length; nItem++)
		{
			arrResultPts.push({
				v:          oNumLit.pts[nItem].val,
				formatCode: oNumLit.pts[nItem].formatCode,
				idx:        oNumLit.pts[nItem].idx
			});
		}

		return {
			formatCode: oNumLit.formatCode,
			pt:         arrResultPts,
			ptCount:    oNumLit.ptCount
		}
	};
	WriterToJSON.prototype.SerStrLit = function(oStrLit)
	{
		if (!oStrLit)
			return oStrLit;

		var arrResultPts = [];
		for (var nItem = 0; nItem < oStrLit.pts.length; nItem++)
		{
			arrResultPts.push({
				v:   oStrLit.pts[nItem].val,
				idx: oStrLit.pts[nItem].idx,
			});
		}

		return {
			pt:         arrResultPts,
			ptCount:    oStrLit.ptCount
		}
	};
	WriterToJSON.prototype.SerErrBars = function(oErrBars)
	{
		if (!oErrBars)
			return oErrBars;

		var sErrBarType = undefined;
		switch(oErrBars.errBarType)
		{
			case AscFormat.st_errbartypeBOTH:
				sErrBarType = "both";
				break;
			case AscFormat.st_errbartypeMINUS:
				sErrBarType = "minus";
				break;
			case AscFormat.st_errbartypePLUS:
				sErrBarType = "plus";
				break;
		}

		var sErrDir = oErrBars.errDir === AscFormat.st_errdirX ? "x" : "y";

		var sErrValType = undefined;
		switch(oErrBars.errValType)
		{
			case AscFormat.st_errvaltypeCUST:
				sErrValType = "cust";
				break;
			case AscFormat.st_errvaltypeFIXEDVAL:
				sErrValType = "fixedVal";
				break;
			case AscFormat.st_errvaltypePERCENTAGE:
				sErrValType = "percentage";
				break;
			case AscFormat.st_errvaltypeSTDDEV:
				sErrValType = "stdDev";
				break;
			case AscFormat.st_errvaltypeSTDERR:
				sErrValType = "stdErr";
				break;
		}

		return {
			errBarType: sErrBarType,
			errDir:     sErrDir,
			errValType: sErrValType,
			minus:      this.SerMinusPlus(oErrBars.minus),
			noEndCap:   oErrBars.noEndCap,
			plus:       this.SerMinusPlus(oErrBars.plus),
			spPr:       this.SerSpPr(oErrBars.spPr),
			val:        oErrBars.val
		}
	};
	WriterToJSON.prototype.SerMinusPlus = function(oMinusPlus)
	{
		if (!oMinusPlus)
			return oMinusPlus;

		return {
			numLit: this.SerNumLit(oMinusPlus.numLit),
			numRef: this.SerNumRef(oMinusPlus.numRef)
		} 
	};
	WriterToJSON.prototype.SerDataPoints = function(arrDpts)
	{
		var arrResultDataPoints = [];

		for (var nItem = 0; nItem < arrDpts.length; nItem++)
		{
			arrResultDataPoints.push({
				bubble3D:         arrDpts[nItem].bubble3D,
				explosion:        arrDpts[nItem].explosion,
				idx:              arrDpts[nItem].idx,
				invertIfNegative: arrDpts[nItem].invertIfNegative,
				marker:           this.SerMarker(arrDpts[nItem].marker),
				pictureOptions:   this.SerPicOptions(arrDpts[nItem].pictureOptions),
				spPr:             this.SerSpPr(arrDpts[nItem].spPr)
			});
		}

		return arrResultDataPoints;
	};
	WriterToJSON.prototype.SerMultiLvlStrRef = function(oRef)
	{
		if (!oRef)
			return oRef;
		
		return {
			f:                oRef.f,
			multiLvlStrCache: this.SerMultiLvlStrCache(oRef.multiLvlStrCache)
		}
	};
	WriterToJSON.prototype.SerMultiLvlStrCache = function(oCache)
	{
		if (!oCache)
			return oCache;

		var arrLvl = [];

		for (var nLvl = 0; nLvl < oCache.lvl.length; nLvl++)
			arrLvl.push(this.SerStrLit(oCache.lvl[nLvl]));
		
		return {
			lvl:     arrLvl,
			ptCount: oCache.ptCount
		}
	};
	WriterToJSON.prototype.SerTrendline = function(oTrendLine)
	{
		if (!oTrendLine)
			return oTrendLine;

		var sTrendlineType = undefined;
		switch(oTrendLine.trendlineType)
		{
			case AscFormat.st_trendlinetypeEXP:
				sTrendlineType = "exp";
				break;
			case AscFormat.st_trendlinetypeLINEAR:
				sTrendlineType = "linear";
				break;
			case AscFormat.st_trendlinetypeLOG:
				sTrendlineType = "log";
				break;
			case AscFormat.st_trendlinetypeMOVINGAVG:
				sTrendlineType = "movingAvg";
				break;
			case AscFormat.st_trendlinetypePOLY:
				sTrendlineType = "poly";
				break;
			case AscFormat.st_trendlinetypePOWER:
				sTrendlineType = "power";
				break;
		}
		
		return {
			backward:      oTrendLine.backward,
			dispEq:        oTrendLine.dispEq,
			dispRSqr:      oTrendLine.dispRSqr,
			forward:       oTrendLine.forward,
			intercept:     oTrendLine.intercept,
			name:          oTrendLine.name,
			order:         oTrendLine.order,
			period:        oTrendLine.period,
			spPr:          this.SerSpPr(oTrendLine.spPr),
			trendlineLbl:  this.SerDlbl(oTrendLine.trendlineLbl),
			trendlineType: sTrendlineType
		}
	};
	WriterToJSON.prototype.SerPicOptions = function(oPicOptions)
	{
		if (!oPicOptions)
			return oPicOptions;

		return {
			applyToEnd:       oPicOptions.applyToEnd,
			applyToFront:     oPicOptions.applyToFront,
			applyToSides:     oPicOptions.applyToSides,
			pictureFormat:    oPicOptions.pictureFormat,
			pictureStackUnit: oPicOptions.pictureStackUnit
		};
	};
	WriterToJSON.prototype.SerBarSeries = function(arrSeries)
	{
		var arrResultSeries = [];

		for (var nItem = 0; nItem < arrSeries.length; nItem++)
		{
			var sShapeType = undefined;
			switch(arrSeries[nItem].shape)
			{
				case AscFormat.BAR_SHAPE_CONE:
					sShapeType = "cone";
					break;
				case AscFormat.BAR_SHAPE_CONETOMAX:
					sShapeType = "coneToMax";
					break;
				case AscFormat.BAR_SHAPE_BOX:
					sShapeType = "box";
					break;
				case AscFormat.BAR_SHAPE_CYLINDER:
					sShapeType = "cylinder";
					break;
				case AscFormat.BAR_SHAPE_PYRAMID:
					sShapeType = "pyramid";
					break;
				case AscFormat.BAR_SHAPE_PYRAMIDTOMAX:
					sShapeType = "pyramidToMax";
					break;
			}
			arrResultSeries.push({
				cat:              this.SerCat(arrSeries[nItem].cat),
				dLbls:            this.SerDLbls(arrSeries[nItem].dLbls),
				dPt:              this.SerDataPoints(arrSeries[nItem].dPt),
				errBars:          this.SerErrBars(arrSeries[nItem].errBars),
				idx:              arrSeries[nItem].idx,
				invertIfNegative: arrSeries[nItem].invertIfNegative,
				order:            arrSeries[nItem].order,
				pictureOptions:   this.SerPicOptions(arrSeries[nItem].pictureOptions),
				shape:            sShapeType,
				spPr:             this.SerSpPr(arrSeries[nItem].spPr),
				trendline:        this.SerTrendline(arrSeries[nItem].trendline),
				tx:               this.SerSerTx(arrSeries[nItem].tx),
				val:              this.SerYVAL(arrSeries[nItem].val)
			});
		}

		return arrResultSeries;
	};
	WriterToJSON.prototype.SerYVAL = function(oVal)
	{
		return {
			numLit: this.SerNumLit(oVal.numLit),
			numRef: this.SerNumRef(oVal.numRef)
		}
	};
	WriterToJSON.prototype.SerCat = function(oCat)
	{
		if (!oCat)
			return oCat;

		return {
			multiLvlStrRef: this.SerMultiLvlStrRef(oCat.multiLvlStrRef),
			numLit:         this.SerNumLit(oCat.numLit),
			numRef:         this.SerNumRef(oCat.numRef),
			strLit:         this.SerStrLit(oCat.strLit),
			strRef:         this.SerStrRef(oCat.strRef)
		}
	};
	WriterToJSON.prototype.SerStrRef = function(oStrRef)
	{
		if (!oStrRef)
			return oStrRef;

		var arrResultPts = [];
		for (var nItem = 0; nItem < oStrRef.strCache.pts.length; nItem++)
		{
			arrResultPts.push({
				v:   oStrRef.strCache.pts[nItem].val,
				idx: oStrRef.strCache.pts[nItem].idx
			});
		}

		return {
			f: oStrRef.f,
			strCache: {
				ptCount: oStrRef.strCache.ptCount,
				pt:      arrResultPts
			}
		}
	};
	WriterToJSON.prototype.SerNumRef = function(oNumRef)
	{
		if (!oNumRef)
			return oNumRef;

		return {
			f: oNumRef.f,
			numCache: this.SerNumLit(oNumRef.numCache)
		}
	};
	WriterToJSON.prototype.SerGeometry = function(oGeometry)
	{
		if (!oGeometry)
			return oGeometry;
		
		// AhPolar
		var arrAhPolarResult = [];
		for (var nItem = 0; nItem < oGeometry.ahPolarLstInfo.length; nItem++)
		{
			arrAhPolarResult.push({
				pos: {
					x: oGeometry.ahPolarLstInfo[nItem].posX,
					y: oGeometry.ahPolarLstInfo[nItem].posY
				},
				gdRefAng: oGeometry.ahPolarLstInfo[nItem].gdRefAng,
				gdRefR: oGeometry.ahPolarLstInfo[nItem].gdRefR,
				maxAng: oGeometry.ahPolarLstInfo[nItem].maxAng,
				maxR: oGeometry.ahPolarLstInfo[nItem].maxR,
				minAng: oGeometry.ahPolarLstInfo[nItem].minAng,
				minR: oGeometry.ahPolarLstInfo[nItem].minR
			});
		}

		// AhXY
		var arrAhXYResult    = [];
		for (nItem = 0; nItem < oGeometry.ahXYLstInfo.length; nItem++)
		{
			arrAhXYResult.push({
				pos: {
					x: oGeometry.ahXYLstInfo[nItem].posX,
					y: oGeometry.ahXYLstInfo[nItem].posY
				},
				gdRefX: oGeometry.ahXYLstInfo[nItem].gdRefX,
				gdRefY: oGeometry.ahXYLstInfo[nItem].gdRefY,
				maxX: oGeometry.ahXYLstInfo[nItem].maxX,
				maxY: oGeometry.ahXYLstInfo[nItem].maxY,
				minX: oGeometry.ahXYLstInfo[nItem].minX,
				minY: oGeometry.ahXYLstInfo[nItem].minY
			});
		}

		// Av
		var oAvResult      = {};
		for (var key in oGeometry.avLst)
			oAvResult[key] = oGeometry.avLst[key]

		// adj
		var oAdjLst = {};
		for (key in oAvResult)
		{
			if (oAvResult[key])
				oAdjLst[key] = oGeometry.gdLst[key];
		}

		// Cnx
		var arrCxnResult     = [];
		for (nItem = 0; nItem < oGeometry.cnxLstInfo.length; nItem++)
		{
			arrCxnResult.push({
				pos: {
					x: oGeometry.cnxLstInfo[nItem].x,
					y: oGeometry.cnxLstInfo[nItem].y
				},
				ang: oGeometry.cnxLstInfo[nItem].ang
			});
		}

		// Gd
		var arrGdResult      = [];
		for (nItem = 0; nItem < oGeometry.gdLstInfo.length; nItem++)
		{
			arrGdResult.push({
				fmla: this.GetFormulaStrType(oGeometry.gdLstInfo[nItem].formula),
				x:    oGeometry.gdLstInfo[nItem].x,
				y:    oGeometry.gdLstInfo[nItem].y,
				z:    oGeometry.gdLstInfo[nItem].z,
				name: oGeometry.gdLstInfo[nItem].name
			});
		}

		// Commands
		var arrPathResult    = [];
		for (nItem = 0; nItem < oGeometry.pathLst.length; nItem++)
			arrPathResult.push(this.SerGeomPath(oGeometry.pathLst[nItem]));

		return {
			ahLst: {
				ahPolar: arrAhPolarResult,
				ahXY:    arrAhXYResult
			},
			avLst:   oAvResult,
			adjLst:  oAdjLst,
			cnxLst:  arrCxnResult,
			gdLst:   arrGdResult,
			pathLst: arrPathResult,
			rect:    oGeometry.rectS,
			preset:  oGeometry.preset
		}
	};
	WriterToJSON.prototype.SerGeomPath = function(oPath)
	{
		if (!oPath)
			return oPath;

		var arrCommands = [];
		for (var nCommand = 0; nCommand < oPath.ArrPathCommandInfo.length; nCommand++)
		{
			var arrObjKeys = Object.keys(oPath.ArrPathCommandInfo[nCommand]);
			var oCommand   = {};
			switch (oPath.ArrPathCommandInfo[nCommand].id)
			{
				case 0:
					oCommand.id =  "moveTo";
					break;
				case 1:
					oCommand.id =  "lnTo";
					break;
				case 2:
					oCommand.id =  "arcTo";
					break;
				case 3:
					oCommand.id =  "cubicBezTo";
					break;
				case 4:
					oCommand.id =  "quadBezTo";
					break;
				case 5:
					oCommand.id =  "close";
					break;
			}
			if (oCommand.id !== "arcTo" && oCommand.id !== "close")
				oCommand.pt = {};
			for (var nKey = 0; nKey < arrObjKeys.length; nKey++)
			{
				if (arrObjKeys[nKey] === "id")
					continue;
				if (!oCommand.pt)
					oCommand[arrObjKeys[nKey]] = oPath.ArrPathCommandInfo[nCommand][arrObjKeys[nKey]];
				else 
					oCommand.pt[arrObjKeys[nKey].toLowerCase()] = oPath.ArrPathCommandInfo[nCommand][arrObjKeys[nKey]];
			}

			arrCommands.push(oCommand);
		}

		return {
			commands:    arrCommands,
			extrusionOk: oPath.extrusionOk,
			fill:        oPath.fill,
			h:           oPath.pathH,
			stroke:      oPath.stroke,
			w:           oPath.pathW
		}
	};
	WriterToJSON.prototype.GetFormulaStrType = function(nFormulaType)
	{
		var sFormulaType = undefined;
		switch(nFormulaType)
		{
			case AscFormat.FORMULA_TYPE_MULT_DIV:
				sFormulaType = "*/";
				break;
			case AscFormat.FORMULA_TYPE_PLUS_MINUS:
				sFormulaType = "+-";
				break;
			case AscFormat.FORMULA_TYPE_PLUS_DIV:
				sFormulaType = "+/";
				break;
			case AscFormat.FORMULA_TYPE_IF_ELSE:
				sFormulaType = "?:";
				break;
			case AscFormat.FORMULA_TYPE_ABS:
				sFormulaType = "abs";
				break;
			case AscFormat.FORMULA_TYPE_AT2:
				sFormulaType = "at2";
				break;
			case AscFormat.FORMULA_TYPE_CAT2:
				sFormulaType = "cat2";
				break;
			case AscFormat.FORMULA_TYPE_COS:
				sFormulaType = "cos";
				break;
			case AscFormat.FORMULA_TYPE_MAX:
				sFormulaType = "max";
				break;
			case AscFormat.FORMULA_TYPE_MOD:
				sFormulaType = "mod";
				break;
			case AscFormat.FORMULA_TYPE_PIN:
				sFormulaType = "pin";
				break;
			case AscFormat.FORMULA_TYPE_SAT2:
				sFormulaType = "sat2";
				break;
			case AscFormat.FORMULA_TYPE_SIN:
				sFormulaType = "sin";
				break;
			case AscFormat.FORMULA_TYPE_SQRT:
				sFormulaType = "sqrt";
				break;
			case AscFormat.FORMULA_TYPE_TAN:
				sFormulaType = "tan";
				break;
			case AscFormat.FORMULA_TYPE_VALUE:
				sFormulaType = "val";
				break;
			case AscFormat.FORMULA_TYPE_MIN:
				sFormulaType = "min";
				break;
		}

		return sFormulaType;
	};
	WriterToJSON.prototype.SerSpPr = function(oSpPr)
	{
		if (!oSpPr)
			return oSpPr;

		var oEffectLst = oSpPr.effectProps ? this.SerEffectLst(oSpPr.effectProps.EffectLst) : oSpPr.effectProps;
		var oEffectDag = oSpPr.effectProps ? this.SerEffectDag(oSpPr.effectProps.EffectDag) : oSpPr.effectProps;
		
		return {
			custGeom:  this.SerGeometry(oSpPr.geometry),
			effectDag: oEffectDag,
			effectLst: oEffectLst,
			ln: this.SerLn(oSpPr.ln),

			fill:   this.SerFill(oSpPr.Fill),
			xfrm:   this.SerXfrm(oSpPr.xfrm),
			bwMode: oSpPr.bwMode
		}
	};
	WriterToJSON.prototype.SerXfrm = function(oXfrm)
	{
		if (!oXfrm)
			return oXfrm;

		return {
			ext: {
				cx: oXfrm.extX ? private_MM2EMU(oXfrm.extX) : oXfrm.extX,
				cy: oXfrm.extY ? private_MM2EMU(oXfrm.extY) : oXfrm.extY
			},
			off: {
				x: oXfrm.offX ? private_MM2EMU(oXfrm.offX) : oXfrm.offX,
				y: oXfrm.offY ? private_MM2EMU(oXfrm.offY) : oXfrm.offY
			},
			flipH: oXfrm.flipH,
			flipV: oXfrm.flipV,
			rot: oXfrm.rot,

			chOffX: oXfrm.chOffX ? private_MM2EMU(oXfrm.chOffX) : oXfrm.chOffX,
    		chOffY: oXfrm.chOffY ? private_MM2EMU(oXfrm.chOffY) : oXfrm.chOffY,
    		chExtX: oXfrm.chExtX ? private_MM2EMU(oXfrm.chExtX) : oXfrm.chExtX,
   			chExtY: oXfrm.chExtY ? private_MM2EMU(oXfrm.chExtY) : oXfrm.chExtY
		}
	};
	WriterToJSON.prototype.SerLn = function(oLn)
	{
		if (!oLn)
			return oLn;
		
		var sAlgnType = oLn.algn != undefined ? (oLn.algn === 0 ? "ctr" : "in") : oLn.algn;
		var sCapType  = undefined;
		switch (oLn.cap)
		{
			case 0:
				sCapType = "flat";
				break;
			case 1:
				sCapType = "rnd";
				break;
			case 2:
				sCapType = "sq";
				break;
		}
		var sCmpdType = undefined;
		switch(oLn.cmpd)
		{
			case 0:
				sCmpdType = "dbl";
				break;
			case 1:
				sCmpdType = "sng";
				break;
			case 2:
				sCmpdType = "thickThin";
				break;
			case 3:
				sCmpdType = "thinThick";
				break;
			case 4:
				sCmpdType = "tri";
				break;
		}
		
		return {
			fill:     this.SerFill(oLn.Fill),
			lineJoin: this.SerLineJoin(oLn.Join),
			
			algn:     sAlgnType,
			cap:      sCapType,
			cmpd:     sCmpdType,
			w:        oLn.w,

			headEnd:  this.SerEndArrow(oLn.headEnd),

			prstDash: this.GetPenDashStrType(oLn.prstDash),

			tailEnd:  this.SerEndArrow(oLn.tailEnd),
			type:     "stroke"
		}
	};
	WriterToJSON.prototype.SerLineJoin = function(oLine)
	{
		if (!oLine)
			return oLine;

		var sType = undefined;
		switch (oLine.type)
		{
			case AscFormat.LineJoinType.Round:
				sType = "round";
				break;
			case  AscFormat.LineJoinType.Bevel:
				sType = "bevel";
				break;
			case  AscFormat.LineJoinType.Miter:
				sType = "miter";
				break;
			case  AscFormat.LineJoinType.Empty:
				sType = "empty";
				break;
		}

		return {
			type: sType,
			lim:  oLine.limit
		}
	};
	WriterToJSON.prototype.SerEndArrow = function(oArrow)
	{
		if (!oArrow)
			return oArrow;
		
		var sType = null;
		switch(oArrow.type)
		{
			case AscFormat.LineEndType.None:
				sType = "none";
				break;
			case AscFormat.LineEndType.Arrow:
				sType = "arrow";
				break;
			case AscFormat.LineEndType.Diamond:
				sType = "diamond";
				break;
			case AscFormat.LineEndType.Oval:
				sType = "oval";
				break;
			case AscFormat.LineEndType.Stealth:
				sType = "stealth";
				break;
			case AscFormat.LineEndType.Triangle:
				sType = "triangle";
				break;
		}

		var sLineEndSize = null;
		switch(oArrow.len)
		{
			case AscFormat.LineEndSize.Large:
				sLineEndSize = "lg";
				break;
			case AscFormat.LineEndSize.Mid:
				sLineEndSize = "med";
				break;
			case AscFormat.LineEndSize.Small:
				sLineEndSize = "sm";
				break;
		}

		var sLineEndWidth = null;
		switch(oArrow.w)
		{
			case AscFormat.LineEndSize.Large:
				sLineEndWidth = "lg";
				break;
			case AscFormat.LineEndSize.Mid:
				sLineEndWidth = "med";
				break;
			case AscFormat.LineEndSize.Small:
				sLineEndWidth = "sm";
				break;
		}
		
		return {
			len:  sLineEndSize,
			type: sType,
			w:    sLineEndWidth
		}
	};
	WriterToJSON.prototype.SerEffectLst = function(oEffectLst)
	{
		if (!oEffectLst)
			return oEffectLst;
		
		return {
			blur:        this.SerEffect(oEffectLst.blur),
			fillOverlay: this.SerEffect(oEffectLst.fillOverlay),
			glow:        this.SerEffect(oEffectLst.glow),
			innerShdw:   this.SerEffect(oEffectLst.innerShdw),
			outerShdw:   this.SerEffect(oEffectLst.outerShdw),
			prstShdw:    this.SerEffect(oEffectLst.prstShdw),
			reflection:  this.SerEffect(oEffectLst.reflection),
			softEdge:    this.SerEffect(oEffectLst.softEdge)
		}
	};
	WriterToJSON.prototype.SerEffectDag = function(oEffectDag)
	{
		if (!oEffectDag)
			return oEffectDag;
			
		var arrEffectLst = this.SerEffects(oEffectDag.effectList);
		
		return {
			effectList: arrEffectLst,
			name:       oEffectDag.name,
			type:       oEffectDag.type
		}
	};
	WriterToJSON.prototype.SerEffects = function(arrEffects)
	{
		var arrEffectLst = [];
		for (var nEffect = 0; nEffect < arrEffects.length; nEffect++)
			arrEffectLst.push(this.SerEffect(arrEffects[nEffect]));

		return arrEffectLst;
	};
	WriterToJSON.prototype.SerEffect = function(oEffect)
	{
		var oElm = null;
		var sBlendType;

		if (oEffect instanceof AscFormat.CAlphaBiLevel)
		{
			oElm = {
				thresh: oEffect.tresh,
				type:   "alphaBiLvl"
			}
		}
		else if (oEffect instanceof AscFormat.CAlphaCeiling)
		{
			oElm = {
				type: "alphaCeiling"
			}
		}
		else if (oEffect instanceof AscFormat.CAlphaFloor)
		{
			oElm = {
				type: "alphaFloor"
			}
		}
		else if (oEffect instanceof AscFormat.CAlphaInv)
		{
			oElm = {
				color: this.SerColor(oEffect.unicolor),
				type:  "alphaInv"
			}
		}
		else if (oEffect instanceof AscFormat.CAlphaMod)
		{
			oElm = {
				cont: this.SerEffectDag(oEffect.cont),
				type: "alphaMod"
			}
		}
		else if (oEffect instanceof AscFormat.CAlphaModFix)
		{
			oElm = {
				amt:  oEffect.amt,
				type: "alphaModFix"
			}
		}
		else if (oEffect instanceof AscFormat.CAlphaOutset)
		{
			oElm = {
				rad:  oEffect.rad,
				type: "alphaOutset"
			}
		}
		else if (oEffect instanceof AscFormat.CAlphaRepl)
		{
			oElm = {
				a:     oEffect.a,
				type: "alphaRepl"
			}
		}
		else if (oEffect instanceof AscFormat.CBiLevel)
		{
			oElm = {
				thresh: oEffect.thresh,
				type:   "biLvl"
			}
		}
		else if (oEffect instanceof AscFormat.CBlend)
		{
			sBlendType = To_XML_ST_BlendMode(oEffect.blend);

			oElm = {
				cont:  this.SerEffectDag(oEffect.cont),
				blend: sBlendType,
				type:  "blend"
			}
		}
		else if (oEffect instanceof AscFormat.CBlur)
		{
			oElm = {
				grow: oEffect.grow,
				rad:  oEffect.rad,
				type: "blur"
			}
		}
		else if (oEffect instanceof AscFormat.CClrChange)
		{
			oElm = {
				clrFrom: this.SerColor(oEffect.clrFrom),
				clrTo:   this.SerColor(oEffect.clrTo),
				useA:    oEffect.useA,
				type:    "clrChange"
			}
		}
		else if (oEffect instanceof AscFormat.CClrRepl)
		{
			oElm = {
				color: this.SerColor(oEffect.color),
				type:  "clrRepl"
			}
		}
		else if (oEffect instanceof AscFormat.CEffectContainer)
		{
			oElm = {
				cont: this.SerEffectDag(oEffect),
				type: "effCont"
			}
		}
		else if (oEffect instanceof AscFormat.CDuotone)
		{
			var arrColors = [];
			for (var nColor = 0; nColor < oEffect.colors.length; nColor++)
				arrColors.push(this.SerColor(oEffect.colors[nColor]));

			oElm = {
				colors: arrColors,
				type:   "duotone"
			}
		}
		else if (oEffect instanceof AscFormat.CEffectElement)
		{
			oElm = {
				ref:  oEffect.ref,
				type: "effect"
			}
		}
		else if (oEffect instanceof AscFormat.CFillEffect)
		{
			oElm = {
				fill: this.SerFill(oEffect.fill),
				type: "fill"
			}
		}
		else if (oEffect instanceof AscFormat.CFillOverlay)
		{
			sBlendType = To_XML_ST_BlendMode(oEffect.blend);

			oElm = {
				fill:  this.SerFill(oEffect.fill),
				blend: sBlendType,
				type:  "fillOvrl"
			}
		}
		else if (oEffect instanceof AscFormat.CGlow)
		{
			oElm = {
				color: this.SerColor(oEffect.color),
				rad:   oEffect.rad,
				type:  "glow"
			}
		}
		else if (oEffect instanceof AscFormat.CGrayscl)
		{
			oElm = {
				type: "gray"
			}
		}
		else if (oEffect instanceof AscFormat.CHslEffect)
		{
			oElm = {
				hue:  oEffect.h,
				lum:  oEffect.l,
				sat:  oEffect.s,
				type: "hsl"
			}
		}
		else if (oEffect instanceof AscFormat.CInnerShdw)
		{
			oElm = {
				color:   this.SerColor(oEffect.color),
				blurRad: oEffect.blurRad,
				dir:     oEffect.dir,
				dist:    oEffect.dist,
				type:    "innerShdw"
			}
		}
		else if (oEffect instanceof AscFormat.CLumEffect)
		{
			oElm = {
				bright:   oEffect.bright,
				contrast: oEffect.contrast,
				type:     "lum"
			}
		}
		else if (oEffect instanceof AscFormat.COuterShdw)
		{
			oElm = {
				color:        this.SerColor(oEffect.color),
				algn:         GetRectAlgnStrType(oEffect.algn),
				blurRad:      oEffect.blurRad,
				dir:          oEffect.dir,
				dist:         oEffect.dist,
				kx:           oEffect.kx,
				ky:           oEffect.ky,
				rotWithShape: oEffect.rotWithShape,
				sx:           oEffect.sx,
				sy:           oEffect.sy,
				type:         "outerShdw"
			}
		}
		else if (oEffect instanceof AscFormat.CPrstShdw)
		{
			var sPrstType = undefined;
			switch (oEffect.prst)
			{
				case c_oAscPresetShadowVal.shdw1:
					sPrstType = "shdw1";
					break;
				case c_oAscPresetShadowVal.shdw2:
					sPrstType = "shdw2";
					break;
				case c_oAscPresetShadowVal.shdw3:
					sPrstType = "shdw3";
					break;
				case c_oAscPresetShadowVal.shdw4:
					sPrstType = "shdw4";
					break;
				case c_oAscPresetShadowVal.shdw5:
					sPrstType = "shdw5";
					break;
				case c_oAscPresetShadowVal.shdw6:
					sPrstType = "shdw6";
					break;
				case c_oAscPresetShadowVal.shdw7:
					sPrstType = "shdw7";
					break;
				case c_oAscPresetShadowVal.shdw8:
					sPrstType = "shdw8";
					break;
				case c_oAscPresetShadowVal.shdw9:
					sPrstType = "shdw9";
					break;
				case c_oAscPresetShadowVal.shdw10:
					sPrstType = "shdw10";
					break;
				case c_oAscPresetShadowVal.shdw11:
					sPrstType = "shdw11";
					break;
				case c_oAscPresetShadowVal.shdw12:
					sPrstType = "shdw12";
					break;
				case c_oAscPresetShadowVal.shdw13:
					sPrstType = "shdw13";
					break;
				case c_oAscPresetShadowVal.shdw14:
					sPrstType = "shdw14";
					break;
				case c_oAscPresetShadowVal.shdw15:
					sPrstType = "shdw15";
					break;
				case c_oAscPresetShadowVal.shdw16:
					sPrstType = "shdw16";
					break;
				case c_oAscPresetShadowVal.shdw17:
					sPrstType = "shdw17";
					break;
				case c_oAscPresetShadowVal.shdw18:
					sPrstType = "shdw18";
					break;
				case c_oAscPresetShadowVal.shdw19:
					sPrstType = "shdw19";
					break;
				case c_oAscPresetShadowVal.shdw20:
					sPrstType = "shdw20";
					break;
			}

			oElm = {
				color: this.SerColor(oEffect.color),
				dir:   oEffect.dir,
				dist:  oEffect.dis,
				prst:  sPrstType,
				type:  "prstShdw"
			}
		}
		else if (oEffect instanceof AscFormat.CReflection)
		{
			oElm = {
				algn:         GetRectAlgnStrType(oEffect.algn),
				blurRad:      oEffect.blurRad,
				dir:          oEffect.dir,
				dist:         oEffect.dist,
				endA:         oEffect.endA,
				endPos:       oEffect.endPos,
				fadeDir:      oEffect.fadeDir,
				kx:           oEffect.kx,
				ky:           oEffect.ky,
				rotWithShape: oEffect.rotWithShape,
				stA:          oEffect.stA,
				stPos:        oEffect.stPos,
				sx:           oEffect.sx,
				sy:           oEffect.sy,
				type: "reflection"
			}
		}
		else if (oEffect instanceof AscFormat.CRelOff)
		{
			oElm = {
				tx:   oEffect.tx,
				ty:   oEffect.ty,
				type: "relOff"
			}
		}
		else if (oEffect instanceof AscFormat.CSoftEdge)
		{
			oElm = {
				rad:  oEffect.rad,
				type: "softEdge"
			}
		}
		else if (oEffect instanceof AscFormat.CTintEffect)
		{
			oElm = {
				amt: oEffect.amt,
				hue: oEffect.hue,
				type: "tint"
			}
		}
		else if (oEffect instanceof AscFormat.CXfrmEffect)
		{
			oElm = {
				kx: oEffect.kx,
				ky: oEffect.ky,
				sx: oEffect.sx,
				sy: oEffect.sy,
				tx: oEffect.tx,
				ty: oEffect.ty,
				type: "xfrm"
			}
		}

		return oElm;
	};
	WriterToJSON.prototype.SerFill = function(oFill)
	{
		if (!oFill)
			return oFill;
		
		var oFillObj = null;
		if (oFill.fill)
		{
			switch (oFill.fill.type)
			{
				case Asc.c_oAscFill.FILL_TYPE_NONE:
					oFillObj = {
						type: "none"
					}
					break;
				case Asc.c_oAscFill.FILL_TYPE_SOLID:
					oFillObj = {
						color: this.SerColor(oFill.fill.color),
						type: "solid"
					}
					break;
				case Asc.c_oAscFill.FILL_TYPE_BLIP:
					oFillObj = this.SerBlipFill(oFill.fill);
					break;
				case Asc.c_oAscFill.FILL_TYPE_NOFILL:
					oFillObj = {
						type: "noFill"
					}
					break;
				case Asc.c_oAscFill.FILL_TYPE_GRAD:
					oFillObj = this.SerGradFill(oFill.fill);
					break;
				case Asc.c_oAscFill.FILL_TYPE_PATT:
					oFillObj = this.SerPattFill(oFill.fill);
					break;
				case Asc.c_oAscFill.FILL_TYPE_GRP:
					oFillObj = {
						type: "grp"
					};
					break;
			}
		}
		
		var sFillType = undefined;
		switch (oFill.type)
		{
			case Asc.c_oAscFill.FILL_TYPE_NONE:
				sFillType = "none";
				break;
			case Asc.c_oAscFill.FILL_TYPE_BLIP:
				sFillType = "blip";
				break;
			case Asc.c_oAscFill.FILL_TYPE_NOFILL:
				sFillType = "noFill";
				break;
			case Asc.c_oAscFill.FILL_TYPE_SOLID:
				sFillType = "solid";
				break;
			case Asc.c_oAscFill.FILL_TYPE_GRAD:
				sFillType = "grad";
				break;
			case Asc.c_oAscFill.FILL_TYPE_PATT:
				sFillType = "patt";
				break;
			case Asc.c_oAscFill.FILL_TYPE_GRP:
				sFillType = "grp";
				break;
		}
		
		return {
			fill:        oFillObj,
			transparent: oFill.transparent,
			type:        "fill",
			fillType:    sFillType
		}
	};
	WriterToJSON.prototype.SerShd = function(oShd)
	{
		if (!oShd)
			return oShd;

		return {
			val: oShd.Value,
			color: oShd.Color ? {
				auto: oShd.Color.Auto,
				r:    oShd.Color.r,
				g:    oShd.Color.g,
				b:    oShd.Color.b
			} : oShd.Color,
			fill: oShd.Fill ? {
				auto: oShd.Fill.Auto,
				r:    oShd.Fill.r,
				g:    oShd.Fill.g,
				b:    oShd.Fill.b
			} : oShd.Fill,
			fillRef: oShd.FillRef ? {
				idx:   oShd.FillRef.idx,
				color: this.SerColor(oShd.FillRef.Color)
			} : oShd.FillRef,
			themeColor: this.SerFill(oShd.Unifill), // Unifill
			themeFill:  this.SerFill(oShd.themeFill)
		}
	};
	WriterToJSON.prototype.SerColor = function(oColor)
	{
		if (!oColor)
			return oColor;

		var oRGBA = oColor.RGBA ? {
			red:   oColor.RGBA.R,
			green: oColor.RGBA.G,
			blue:  oColor.RGBA.B,
			alpha: oColor.RGBA.A 
		} : oColor.RGBA;

		var oColorObj  = null;
		var oColorType = undefined;
		if (oColor.color)
		{
			switch (oColor.color.type)
			{
				case Asc.c_oAscColor.COLOR_TYPE_NONE:
					oColorType = "none";
					oColorObj  = {
						type: oColorType
					}
					break;
				case Asc.c_oAscColor.COLOR_TYPE_SRGB:
				{
					oColorType = "srgb";
					oColorObj  = {
						rgba: {
							red:   oColor.color.RGBA.R,
							green:  oColor.color.RGBA.G,
							blue:  oColor.color.RGBA.B,
							alpha: oColor.color.RGBA.A 
						},
						type: oColorType
					}
					break;
				}
				case Asc.c_oAscColor.COLOR_TYPE_PRST:
				{
					oColorType = "prst";
					oColorObj  = {
						rgba: {
							red:   oColor.color.RGBA.R,
							green:  oColor.color.RGBA.G,
							blue:  oColor.color.RGBA.B,
							alpha: oColor.color.RGBA.A 
						},
						id:   oColor.color.id,
						type: oColorType
					}
					break;
				}
				case Asc.c_oAscColor.COLOR_TYPE_SCHEME:
				{
					oColorType = "scheme";
					oColorObj  = {
						rgba: {
							red:   oColor.color.RGBA.R,
							green:  oColor.color.RGBA.G,
							blue:  oColor.color.RGBA.B,
							alpha: oColor.color.RGBA.A 
						},
						id:   oColor.color.id,
						type: oColorType
					}
					break;
				}
				case Asc.c_oAscColor.COLOR_TYPE_SYS:
				{
					oColorType = "sys";
					oColorObj  = {
						rgba: {
							red:   oColor.color.RGBA.R,
							green: oColor.color.RGBA.G,
							blue:  oColor.color.RGBA.B,
							alpha: oColor.color.RGBA.A 
						},
						id:   oColor.color.id,
						type: oColorType
					}
					break;
				}
				case Asc.c_oAscColor.COLOR_TYPE_STYLE:
				{
					oColorType = "style";
					oColorObj  = {
						auto: oColor.color.bAuto,
						val:  oColor.color.val,
						type: oColorType
					}
					break;
				}
			}
		}
		
		return {
			rgba:  oRGBA,
			color: oColorObj,
			mods:  this.SerColorModifiers(oColor.Mods),
			type:  "uniColor"
		}
	};
	WriterToJSON.prototype.SerColorModifiers = function(oColorModifiers)
	{
		if (!oColorModifiers)
			return oColorModifiers;

		var arrColorMods = [];
		for (var nMod = 0; nMod < oColorModifiers.Mods.length; nMod++)
		{
			arrColorMods.push({
				name: oColorModifiers.Mods[nMod].name,
				val:  oColorModifiers.Mods[nMod].val
			});
		}

		return arrColorMods;
	};
	WriterToJSON.prototype.SerGradFill = function(oGradFill)
	{
		if (!oGradFill)
			return oGradFill;
		
		var arrGsLst = [];
		for (var nGs = 0; nGs < oGradFill.colors.length; nGs++)
			arrGsLst.push(this.SerGradStop(oGradFill.colors[nGs]));

		var sPathShadeType = undefined;
		if (oGradFill.path)
		{
			switch(oGradFill.path.path)
			{
				case 0:
					sPathShadeType = "circle";
					break;
				case 1:
					sPathShadeType = "rect";
					break;
				case 2:
					sPathShadeType = "shape";
					break;
			}
		}

		return {
			gsLst: arrGsLst,

			lin:   oGradFill.lin ? {
				ang:    oGradFill.lin.angle,
				scaled: oGradFill.lin.scale
			} : oGradFill.lin,

			path: oGradFill.path ? {
				path: sPathShadeType,
				fillToRect: oGradFill.path.rect ? {
					b: oGradFill.path.rect.b,
					l: oGradFill.path.rect.l,
					r: oGradFill.path.rect.r,
					t: oGradFill.path.rect.t
				} : oGradFill.path.rect
			} : oGradFill.path,

			rotWithShape: oGradFill.rotateWithShape,
			type: "gradFill"
		}
	};
	WriterToJSON.prototype.SerGradStop = function(oGradStop)
	{
		if (!oGradStop)
			return oGradStop;

		return {
			color: this.SerColor(oGradStop.color),
			pos:   oGradStop.pos,
			type:  "gradStop"
		}
	};
	WriterToJSON.prototype.SerPattFill = function(oPattFill)
	{
		if (!oPattFill)
			return oPattFill;

		return {
			bgClr: oPattFill.bgClr ? this.SerColor(oPattFill.bgClr) : oPattFill.bgClr,
			fgClr: oPattFill.fgClr ? this.SerColor(oPattFill.fgClr) : oPattFill.fgClr,
			prst:  this.GetPresetStrType(oPattFill.ftype),
			type: "pattFill"
		}
	};
	WriterToJSON.prototype.GetPresetStrType = function(nType)
	{
		switch (nType)
		{
			case AscCommon.global_hatch_offsets.cross:
				return "cross";
			case AscCommon.global_hatch_offsets.dashDnDiag:
				return "dashDnDiag";
			case AscCommon.global_hatch_offsets.dashHorz:
				return "dashHorz";
			case AscCommon.global_hatch_offsets.dashUpDiag:
				return "dashUpDiag";
			case AscCommon.global_hatch_offsets.dashVert:
				return "dashVert";
			case AscCommon.global_hatch_offsets.diagBrick:
				return "diagBrick";
			case AscCommon.global_hatch_offsets.diagCross:
				return "diagCross";
			case AscCommon.global_hatch_offsets.divot:
				return "divot";
			case AscCommon.global_hatch_offsets.dkDnDiag:
				return "dkDnDiag";
			case AscCommon.global_hatch_offsets.dkHorz:
				return "dkHorz";
			case AscCommon.global_hatch_offsets.dkUpDiag:
				return "dkUpDiag";
			case AscCommon.global_hatch_offsets.dkVert:
				return "dkVert";
			case AscCommon.global_hatch_offsets.dnDiag:
				return "dnDiag";
			case AscCommon.global_hatch_offsets.dotDmnd:
				return "dotDmnd";
			case AscCommon.global_hatch_offsets.dotGrid:
				return "dotGrid";
			case AscCommon.global_hatch_offsets.horz:
				return "horz";
			case AscCommon.global_hatch_offsets.horzBrick:
				return "horzBrick";
			case AscCommon.global_hatch_offsets.lgCheck:
				return "lgCheck";
			case AscCommon.global_hatch_offsets.lgConfetti:
				return "lgConfetti";
			case AscCommon.global_hatch_offsets.lgGrid:
				return "lgGrid";
			case AscCommon.global_hatch_offsets.ltDnDiag:
				return "ltDnDiag";
			case AscCommon.global_hatch_offsets.ltHorz:
				return "ltHorz";
			case AscCommon.global_hatch_offsets.ltUpDiag:
				return "ltUpDiag";
			case AscCommon.global_hatch_offsets.ltVert:
				return "ltVert";
			case AscCommon.global_hatch_offsets.narHorz:
				return "narHorz";
			case AscCommon.global_hatch_offsets.narVert:
				return "narVert";
			case AscCommon.global_hatch_offsets.openDmnd:
				return "openDmnd";
			case AscCommon.global_hatch_offsets.pct10:
				return "pct10";
			case AscCommon.global_hatch_offsets.pct20:
				return "pct20";
			case AscCommon.global_hatch_offsets.pct25:
				return "pct25";
			case AscCommon.global_hatch_offsets.pct30:
				return "pct30";
			case AscCommon.global_hatch_offsets.pct40:
				return "pct40";
			case AscCommon.global_hatch_offsets.pct5:
				return "pct5";
			case AscCommon.global_hatch_offsets.pct50:
				return "pct50";
			case AscCommon.global_hatch_offsets.pct60:
				return "pct60";
			case AscCommon.global_hatch_offsets.pct70:
				return "pct70";
			case AscCommon.global_hatch_offsets.pct75:
				return "pct75";
			case AscCommon.global_hatch_offsets.pct80:
				return "pct80";
			case AscCommon.global_hatch_offsets.pct90:
				return "pct90";
			case AscCommon.global_hatch_offsets.plaid:
				return "plaid";
			case AscCommon.global_hatch_offsets.shingle:
				return "shingle";
			case AscCommon.global_hatch_offsets.smCheck:
				return "smCheck";
			case AscCommon.global_hatch_offsets.smConfetti:
				return "smConfetti";
			case AscCommon.global_hatch_offsets.smGrid:
				return "smGrid";
			case AscCommon.global_hatch_offsets.solidDmnd:
				return "solidDmnd";
			case AscCommon.global_hatch_offsets.sphere:
				return "sphere";
			case AscCommon.global_hatch_offsets.trellis:
				return "trellis";
			case AscCommon.global_hatch_offsets.upDiag:
				return "upDiag";
			case AscCommon.global_hatch_offsets.vert:
				return "vert";
			case AscCommon.global_hatch_offsets.wave:
				return "wave";
			case AscCommon.global_hatch_offsets.wdDnDiag:
				return "wdDnDiag";
			case AscCommon.global_hatch_offsets.wdUpDiag:
				return "wdUpDiag";
			case AscCommon.global_hatch_offsets.weave:
				return "weave";
			case AscCommon.global_hatch_offsets.zigZag:
				return "zigZag";
		}
	};
	WriterToJSON.prototype.SerTxPr = function(oTxPr)
	{
		if (!oTxPr)
			return oTxPr;

		return {
			bodyPr:   this.SerBodyPr(oTxPr.bodyPr),
			lstStyle: this.SerLstStyle(oTxPr.lstStyle),
			content:  this.SerDrawingDocContent(oTxPr.content, undefined, undefined, undefined, true)
		}
	};
	WriterToJSON.prototype.SerBodyPr = function(oBodyPr)
	{
		if (!oBodyPr)
			return oBodyPr;

		var sAnchorType = null;
		switch(oBodyPr.anchor)
		{
			case AscFormat.VERTICAL_ANCHOR_TYPE_BOTTOM:
				sAnchorType = "b";
				break;
			case AscFormat.VERTICAL_ANCHOR_TYPE_CENTER:
				sAnchorType = "ctr";
				break;
			case AscFormat.VERTICAL_ANCHOR_TYPE_DISTRIBUTED:
				sAnchorType = "dist";
				break;
			case AscFormat.VERTICAL_ANCHOR_TYPE_JUSTIFIED:
				sAnchorType = "just";
				break;
			case AscFormat.VERTICAL_ANCHOR_TYPE_TOP:
				sAnchorType = "t";
				break;
		}

		var sHorzOverflow = null;
		switch(oBodyPr.horzOverflow)
		{
			case AscFormat.nOTClip:
				sHorzOverflow = "clip";
				break;
			case AscFormat.nOTOwerflow:
				sHorzOverflow = "overflow";
				break;
		}

		var sVertOverflow = null;
		switch( oBodyPr.vertOverflow)
		{
			case AscFormat.nOTClip:
				sVertOverflow = "clip";
				break;
			case AscFormat.nOTEllipsis:
				sVertOverflow = "ellipsis";
				break;
			case AscFormat.nOTOwerflow:
				sVertOverflow = "overflow";
				break;
		}

		var sVertType = null;
		switch(oBodyPr.vert)
		{
			case AscFormat.nVertTTeaVert:
				sVertType = "eaVert";
				break;
			case AscFormat.nVertTThorz:
				sVertType = "horz";
				break;
			case AscFormat.nVertTTmongolianVert:
				sVertType = "mongolianVert";
				break;
			case AscFormat.nVertTTvert:
				sVertType = "vert";
				break;
			case AscFormat.nVertTTvert270:
				sVertType = "vert270";
				break;
			case AscFormat.nVertTTwordArtVert:
				sVertType = "wordArtVert";
				break;
			case AscFormat.nVertTTwordArtVertRtl:
				sVertType = "wordArtVertRtl";
				break;
		}

		var sWrapType = null;
		switch (oBodyPr.wrap)
		{
			case AscFormat.nTWTNone:
				sWrapType = "none";
				break;
			case AscFormat.nTWTSquare:
				sWrapType = "square";
				break;
		}

		return {
			flatTx:           oBodyPr.flatTx != null ? private_MM2EMU(oBodyPr.flatTx) : oBodyPr.flatTx,
			normAutofit:      this.SerTextFit(oBodyPr.textFit),
			prstTxWarp:       this.SerGeometry(oBodyPr.prstTxWarp),
			anchor:           sAnchorType,
			anchorCtr:        oBodyPr.anchorCtr,
			bIns:             typeof(oBodyPr.bIns) === "number" ? private_MM2EMU(oBodyPr.bIns) : oBodyPr.bIns,
			compatLnSpc:      oBodyPr.compatLnSpc,
			forceAA:          oBodyPr.forceAA,
			fromWordArt:      oBodyPr.fromWordArt,
			horzOverflow:     sHorzOverflow,
			lIns:             typeof(oBodyPr.lIns) === "number" ? private_MM2EMU(oBodyPr.lIns) : oBodyPr.lIns,
			numCol:           oBodyPr.numCol,
			rIns:             typeof(oBodyPr.rIns) === "number" ? private_MM2EMU(oBodyPr.rIns) : oBodyPr.rIns,
			rot:              oBodyPr.rot,
			rtlCol:           oBodyPr.rtlCol,
			spcCol:           typeof(oBodyPr.spcCol) === "number" ? private_MM2EMU(oBodyPr.spcCol) : oBodyPr.spcCol,
			spcFirstLastPara: oBodyPr.spcFirstLastPara,
			tIns:             typeof(oBodyPr.tIns) === "number" ? private_MM2EMU(oBodyPr.tIns) : oBodyPr.tIns,
			upright:          oBodyPr.upright,
			vert:             sVertType,
			vertOverflow:     sVertOverflow,
			wrap:             sWrapType
		}
	};
	WriterToJSON.prototype.SerTextFit = function(oTextFit)
	{
		if (!oTextFit)
			return oTextFit;

		return {
			type:           oTextFit.type,
			fontScale:      oTextFit.fontScale,
			lnSpcReduction: oTextFit.lnSpcReduction
		}
	};
	WriterToJSON.prototype.SerLstStyle = function(oListStyle)
	{
		if (!oListStyle)
			return oListStyle;
		
		var arrResult = [];

		for (var nLvl = 0; nLvl < oListStyle.levels.length; nLvl++)
		{
			arrResult.push(this.SerParaPr(oListStyle.levels[nLvl]));
		}

		return arrResult;
	};
	WriterToJSON.prototype.SerStyle = function(oStyle)
	{
		if (!oStyle)
			return oStyle;

		var oStyleWithFullPr  = oStyle.Copy();
		oStyleWithFullPr.Link = oStyle.Link;
		oStyleWithFullPr.Id   = oStyle.Id;

		var oStyles = private_GetStyles();
		function GetFullStylePr(styleObj)
		{
			var nBasedOnId    = styleObj.GetBasedOn();
			var oStyleBasedOn = null;
			if (nBasedOnId)
			{
				oStyleBasedOn = oStyles.Get(nBasedOnId);
				oStyleBasedOn = oStyleBasedOn ? oStyleBasedOn.Copy() : null;
			}

			var nBasedOnBasedOnId    = oStyleBasedOn ? oStyleBasedOn.GetBasedOn() : null;
			var oStyleBasedOnBasedOn = nBasedOnBasedOnId ? oStyles.Get(nBasedOnBasedOnId) : null;
			
			if (!oStyleBasedOn)
				return;

			if (oStyleBasedOnBasedOn)
				GetFullStylePr(oStyleBasedOn);

			oStyleBasedOn.TextPr.Merge(styleObj.TextPr);
			oStyleBasedOn.ParaPr.Merge(styleObj.ParaPr);
			oStyleBasedOn.TablePr.Merge(styleObj.TablePr);
			oStyleBasedOn.TableRowPr.Merge(styleObj.TableRowPr);
			oStyleBasedOn.TableCellPr.Merge(styleObj.TableCellPr);

			styleObj.TextPr      = oStyleBasedOn.TextPr;
			styleObj.ParaPr      = oStyleBasedOn.ParaPr;
			styleObj.TablePr     = oStyleBasedOn.TablePr;
			styleObj.TableRowPr  = oStyleBasedOn.TableRowPr;
			styleObj.TableCellPr = oStyleBasedOn.TableCellPr;

			if (styleObj.TableBand1Horz && oStyleBasedOn.TableBand1Horz)
			{
				oStyleBasedOn.TableBand1Horz.Merge(styleObj.TableBand1Horz);
				oStyleBasedOn.TableBand1Vert.Merge(styleObj.TableBand1Vert);
				oStyleBasedOn.TableBand2Horz.Merge(styleObj.TableBand2Horz);
				oStyleBasedOn.TableBand2Vert.Merge(styleObj.TableBand2Vert);
				oStyleBasedOn.TableFirstCol.Merge(styleObj.TableFirstCol);
				oStyleBasedOn.TableFirstRow.Merge(styleObj.TableFirstRow);
				oStyleBasedOn.TableLastCol.Merge(styleObj.TableLastCol);
				oStyleBasedOn.TableLastRow.Merge(styleObj.TableLastRow);
				oStyleBasedOn.TableTLCell.Merge(styleObj.TableTLCell);
				oStyleBasedOn.TableTRCell.Merge(styleObj.TableTRCell);
				oStyleBasedOn.TableBLCell.Merge(styleObj.TableBLCell);
				oStyleBasedOn.TableBRCell.Merge(styleObj.TableBRCell);
				oStyleBasedOn.TableWholeTable.Merge(styleObj.TableWholeTable);

				styleObj.TableBand1Horz  = oStyleBasedOn.TableBand1Horz;
				styleObj.TableBand1Vert  = oStyleBasedOn.TableBand1Vert;
				styleObj.TableBand2Horz  = oStyleBasedOn.TableBand2Horz;
				styleObj.TableBand2Vert  = oStyleBasedOn.TableBand2Vert;
				styleObj.TableFirstCol   = oStyleBasedOn.TableFirstCol;
				styleObj.TableFirstRow   = oStyleBasedOn.TableFirstRow;
				styleObj.TableLastCol    = oStyleBasedOn.TableLastCol;
				styleObj.TableLastRow    = oStyleBasedOn.TableLastRow;
				styleObj.TableTLCell     = oStyleBasedOn.TableTLCell;
				styleObj.TableTRCell     = oStyleBasedOn.TableTRCell;
				styleObj.TableBLCell     = oStyleBasedOn.TableBLCell;
				styleObj.TableBRCell     = oStyleBasedOn.TableBRCell;
				styleObj.TableWholeTable = oStyleBasedOn.TableWholeTable;
			}
		}

		GetFullStylePr(oStyleWithFullPr);

		var sStyleType;
		switch (oStyleWithFullPr.Type)
		{
			case styletype_Paragraph:
				sStyleType = "paragraphStyle";
				break;
			case styletype_Numbering:
				sStyleType = "numberingStyle";
				break;
			case styletype_Table:
				sStyleType = "tableStyle";
				break;
			case styletype_Character:
				sStyleType = "characterStyle";
				break;
			default:
				sStyleType = "paragraphStyle";
				break;
		}

		// Style Conditional Table Formatting Properties
		var oTblStylesPr = oStyleWithFullPr.TableBand1Horz ? {
			band1Horz:  this.SerTableStylePr(oStyleWithFullPr.TableBand1Horz, "bandedRow"),
			band1Vert:  this.SerTableStylePr(oStyleWithFullPr.TableBand1Vert, "bandedColumn"),
			band2Horz:  this.SerTableStylePr(oStyleWithFullPr.TableBand2Horz, "bandedRowEven"),
			band2Vert:  this.SerTableStylePr(oStyleWithFullPr.TableBand2Vert, "bandedColumnEven"),
			firstCol:   this.SerTableStylePr(oStyleWithFullPr.TableFirstCol, "firstColumn"),
			firstRow:   this.SerTableStylePr(oStyleWithFullPr.TableFirstRow, "firstRow"),
			lastCol:    this.SerTableStylePr(oStyleWithFullPr.TableLastCol, "lastColumn"),
			lastRow:    this.SerTableStylePr(oStyleWithFullPr.TableLastRow, "lastRow"),
			neCell:     this.SerTableStylePr(oStyleWithFullPr.TableTRCell, "topRightCell"),
			nwCell:     this.SerTableStylePr(oStyleWithFullPr.TableTLCell, "topLeftCell"),
			seCell:     this.SerTableStylePr(oStyleWithFullPr.TableBRCell, "bottomRightCell"),
			swCell:     this.SerTableStylePr(oStyleWithFullPr.TableBLCell, "bottomLeftCell"),
			wholeTable: this.SerTableStylePr(oStyleWithFullPr.TableWholeTable, "wholeTable")
		} : oStyleWithFullPr.TableBand1Horz;

		return {
			basedOn:        oStyleWithFullPr.BasedOn,
			hidden:         oStyleWithFullPr.hidden,
			link:           oStyleWithFullPr.Link,
			name:           oStyleWithFullPr.Name,
			next:           oStyleWithFullPr.Next,
			pPr:            this.SerParaPr(oStyleWithFullPr.ParaPr),
			qFormat:        oStyleWithFullPr.qFormat,
			rPr:            this.SerTextPr(oStyleWithFullPr.TextPr),
			semiHidden:     oStyleWithFullPr.semiHidden,
			tblPr:          this.SerTablePr(oStyleWithFullPr.TablePr),
			tblStylePr:     oTblStylesPr,
			tcPr:           this.SerTableCellPr(oStyleWithFullPr.TableCellPr),
			trPr:           this.SerTableRowPr(oStyleWithFullPr.TableRowPr),
			uiPriority:     oStyleWithFullPr.uiPriority,
			unhideWhenUsed: oStyleWithFullPr.unhideWhenUsed,
			customStyle:    oStyleWithFullPr.Custom,
			styleId:        oStyleWithFullPr.Id,
			styleType:      sStyleType,
			type:           "style"
		}
	};
	WriterToJSON.prototype.SerTableStylePr = function(oPr, sStyleType)
	{
		if (!oPr)
			return oPr;

		return {
			pPr:       this.SerParaPr(oPr.ParaPr),
			rPr:       this.SerTextPr(oPr.TextPr),
			tblPr:     this.SerTablePr(oPr.TablePr, null),
			tcPr:      this.SerTableCellPr(oPr.TableCellPr),
			trPr:      this.SerTableRowPr(oPr.TableRowPr),
			styleType: sStyleType,
			type:      "tableStyle"
		}
	};
	WriterToJSON.prototype.SerTableMeasurement = function(oMeasurement)
	{
		if (!oMeasurement)
			return oMeasurement;

		switch (oMeasurement.Type)
		{
			case tblwidth_Auto:
				return {
					type: "auto",
					w:    0
				};
			case tblwidth_Mm:
				return {
					type: "dxa",
					w:    private_MM2Twips(oMeasurement.W)
				};
			case tblwidth_Nil:
				return {
					type: "nil",
					w:    0
				};
			case tblwidth_Pct:
				return {
					type: "pct",
					w:    oMeasurement.W
				};
		}

		return undefined;
	};
	WriterToJSON.prototype.SerTablePr = function(oPr, oTable)
	{
		if (!oPr)
			return oPr;

		var sJc          = undefined;
		var sLayoutType  = oPr.TableLayout == undefined ? oPr.TableLayout : (oPr.TableLayout === tbllayout_Fixed ? "fixed" : "autofit");
		var sOverlapType = oTable ? (oTable.AllowOverlap ? "overlap" : "never") : "never";
		var isInline     = oTable ? oTable.Inline : false;
		switch (oPr.Jc)
		{
			case AscCommon.align_Left:
				sJc = "start";
				break;
			case AscCommon.align_Center:
				sJc = "center";
				break;
			case AscCommon.align_Right:
				sJc = "end";
				break;
		}

		// table look
		var oTableLook = oTable ? (oTable.TableLook ? {
			firstColumn: oTable.TableLook.m_bFirst_Col,
			firstRow:    oTable.TableLook.m_bFirst_Row,
			lastColumn:  oTable.TableLook.m_bLast_Col,
			lastRow:     oTable.TableLook.m_bLast_Row,
			noHBand:     !oTable.TableLook.m_bBand_Hor,
			noVBand:     !oTable.TableLook.m_bBand_Ver
		} : undefined) : undefined;
		
		// anchorH
		var sHorAnchor = undefined;
		if (oTable && oTable.PositionH)
		{
			switch (oTable.PositionH.RelativeFrom)
			{
				case Asc.c_oAscHAnchor.Margin:
					sHorAnchor = "margin";
					break;
				case Asc.c_oAscHAnchor.Text:
					sHorAnchor = "text";
					break;
				case Asc.c_oAscHAnchor.Page:
					sHorAnchor = "page";
					break;
			}
		}

		// anchorV
		var sVerAnchor = undefined;
		if (oTable && oTable.PositionV)
		{
			switch (oTable.PositionV.RelativeFrom)
			{
				case Asc.c_oAscVAnchor.Margin:
					sVerAnchor = "margin";
					break;
				case Asc.c_oAscVAnchor.Text:
					sVerAnchor = "text";
					break;
				case Asc.c_oAscVAnchor.Page:
					sVerAnchor = "page";
					break;
			}
		}

		// alignH
		var sHorAlign = undefined;
		if (oTable && oTable.PositionH && oTable.PositionH.Align)
		{
			switch (oTable.PositionH.Value)
			{
				case Asc.c_oAscXAlign.Center:
					sHorAlign = "center";
					break;
				case Asc.c_oAscXAlign.Inside:
					sHorAlign = "inside";
					break;
				case Asc.c_oAscXAlign.Left:
					sHorAlign = "left";
					break;
				case Asc.c_oAscXAlign.Outside:
					sHorAlign = "outside";
					break;
				case Asc.c_oAscXAlign.Right:
					sHorAlign = "right";
					break;
			}
		}
		
		// alignV
		var sVerAlign = undefined;
		if (oTable && oTable.PositionV && oTable.PositionV.Align)
		{
			switch (oTable.PositionV.Value)
			{
				case Asc.c_oAscYAlign.Bottom:
					sVerAlign = "bottom";
					break;
				case Asc.c_oAscYAlign.Center:
					sVerAlign = "center";
					break;
				case Asc.c_oAscYAlign.Inline:
					sVerAlign = "inline";
					break;
				case Asc.c_oAscYAlign.Inside:
					sVerAlign = "inside";
					break;
				case Asc.c_oAscYAlign.Outside:
					sVerAlign = "outside";
					break;
				case Asc.c_oAscYAlign.Top:
					sVerAlign = "top";
					break;
			}
		}

		// tablePosPr
		var oTblPosPr = oTable ? (oTable.PositionH && oTable.PositionV ? {
			horzAnchor:     sHorAnchor,
			vertAnchor:     sVerAnchor,
			tblpXSpec:      sHorAlign,
			tblpYSpec:      sVerAlign,
			tblpX:          private_MM2Twips(oTable.PositionH.Value),
			tblpY:          private_MM2Twips(oTable.PositionV.Value),
			bottomFromText: private_MM2Twips(oTable.Distance.B),
			leftFromText:   private_MM2Twips(oTable.Distance.L),
			rightFromText:  private_MM2Twips(oTable.Distance.R),
			topFromText:    private_MM2Twips(oTable.Distance.T)

		} : undefined) : undefined;

		var oStyles = private_GetStyles();
		// table style
		var oTableStyle = oTable ? oStyles.Get(oTable.TableStyle) : undefined;

		return {
			jc:         sJc,
			shd:        this.SerShd(oPr.Shd),

			tblBorders: {
				bottom:  this.SerBorder(oPr.TableBorders.Bottom),
				end:     this.SerBorder(oPr.TableBorders.Right),
				insideH: this.SerBorder(oPr.TableBorders.InsideH),
				insideV: this.SerBorder(oPr.TableBorders.InsideV),
				start:   this.SerBorder(oPr.TableBorders.Left),
				top:     this.SerBorder(oPr.TableBorders.Top)
			},
			tblCaption: oPr.TableCaption,

			tblCellMar: oPr.TableCellMar ? {
				bottom: this.SerTableMeasurement(oPr.TableCellMar.Bottom),
				left:   this.SerTableMeasurement(oPr.TableCellMar.Left),
				right:  this.SerTableMeasurement(oPr.TableCellMar.Right),
				top:    this.SerTableMeasurement(oPr.TableCellMar.Top)
			} : oPr.TableCellMar,

			tblCellSpacing:      oPr.TableCellMar.TableCellSpacing ? private_MM2Twips(oPr.TableCellMar.TableCellSpacing) : oPr.TableCellMar.TableCellSpacing,
			tblDescription:      oPr.TableDescription,
			tblInd:              oPr.TableInd ? private_MM2Twips(oPr.TableInd) : oPr.TableInd,
			tblLayout:           sLayoutType,
			tblLook:             oTableLook,
			tblOverlap:          sOverlapType,
			tblpPr:              oTblPosPr,
			tblPrChange:         this.SerTablePr(oPr.PrChange),
			tblStyle:            oTableStyle ? this.SerStyle(oTableStyle) : oTableStyle,
			tblStyleColBandSize: oPr.TableStyleColBandSize,
			tblStyleRowBandSize: oPr.TableStyleRowBandSize,
			tblW:                this.SerTableMeasurement(oPr.TableW),
			inline:              isInline,

			reviewInfo: this.SerReviewInfo(oPr.ReviewInfo),
			type:       "tablePr"
		}
	};
	WriterToJSON.prototype.SerTable = function(oTable, aComplexFieldsToSave, oMapCommentsInfo)
	{
		if (!oTable)
			return oTable;

		var oLogicDocument = private_GetLogicDocument();
		var oTableObj = {
			bPresentation: oTable.bPresentation,
			tblGrid: [],
			tblPr:   this.SerTablePr(oTable.Pr, oTable),
			content: [],
			changes: [],
			//tableMarkup: this.SerTableMarkup(oTable.Markup),
			type:    "table"
		}

		for (var nGrid = 0; nGrid < oTable.TableGrid.length; nGrid++)
			oTableObj.tblGrid.push({
				w:    private_MM2Twips(oTable.TableGrid[nGrid]),
				type: "gridCol"
			});

		for (var nRow = 0; nRow < oTable.Content.length; nRow++)
			oTableObj.content.push(this.SerTableRow(oTable.Content[nRow], aComplexFieldsToSave, oMapCommentsInfo));

		// Revisions
		var aChanges = oLogicDocument.TrackRevisionsManager ? oLogicDocument.TrackRevisionsManager.GetElementChanges(oTable.Id) : [];

		for (var nChange = 0; nChange < aChanges.length; nChange++)
		{
			oTableObj.changes.push(this.SerRevisionChange(aChanges[nChange]));
		}

		return oTableObj;	
	};
	WriterToJSON.prototype.SerTableMarkup = function(oMarkup)
	{
		if (!oMarkup)
			return oMarkup;

		return {
			cols:   oMarkup.Cols,
			curCol: oMarkup.CurCol,
			curRow: oMarkup.CurRow,
			internal: {
				cellIndex: oMarkup.Internal.CellIndex,
				pageNum:   oMarkup.Internal.PageNum,
				rowIndex:  oMarkup.Internal.RowIndex
			},
			margins:    oMarkup.Margins,
			rows:       oMarkup.Rows,
			transformX: oMarkup.TransformX,
			transformY: oMarkup.TransformY,
			x:          oMarkup.X
		}
	};
	WriterToJSON.prototype.SerTableCellPr = function(oPr)
	{
		if (!oPr)
			return oPr;

		var sHMerge = oPr.HMerge ? (oPr.HMerge === 2 ? "continue" : "restart") : oPr.HMerge;
		var sVMerge = oPr.VMerge ? (oPr.VMerge === 2 ? "continue" : "restart") : oPr.VMerge;
		var sVAlign = undefined;

		// alignV
		if (oPr.VAlign)
		{
			switch (oPr.VAlign)
			{
				case vertalignjc_Top:
					sVAlign = "top";
					break;
				case vertalignjc_Center:
					sVAlign = "center";
					break;
				case vertalignjc_Bottom:
					sVAlign = "bottom";
					break;
			}
		}

		// text direction
		var sTextDir = undefined;
		switch (oPr.TextDirection)
		{
			case textdirection_LRTB:
				sTextDir = "lrtb";
				break;
			case textdirection_TBRL:
				sTextDir = "tbrl";
				break;
			case textdirection_BTLR:
				sTextDir = "btlr";
				break;
			case textdirection_LRTBV:
				sTextDir = "lrtbV";
				break;
			case textdirection_TBRLV:
				sTextDir = "tbrlV";
				break;
			case textdirection_TBLRV:
				sTextDir = "tblrV";
				break;
		}

		return {
			gridSpan: oPr.GridSpan,
			hMerge:   sHMerge,
			noWrap:   oPr.NoWrap,
			shd:      this.SerShd(oPr.Shd),

			tcBorders: {
				bottom:  this.SerBorder(oPr.TableCellBorders.Bottom),
				end:     this.SerBorder(oPr.TableCellBorders.Right),
				start:   this.SerBorder(oPr.TableCellBorders.Left),
				top:     this.SerBorder(oPr.TableCellBorders.Top)
			},

			tcMar: oPr.TableCellMar ? {
				bottom: this.SerTableMeasurement(oPr.TableCellMar.Bottom),
				left:   this.SerTableMeasurement(oPr.TableCellMar.Left),
				right:  this.SerTableMeasurement(oPr.TableCellMar.Right),
				top:    this.SerTableMeasurement(oPr.TableCellMar.Top),
			} : oPr.TableCellMar,

			tcPrChange:    this.SerTableCellPr(oPr.PrChange),
			tcW:           this.SerTableMeasurement(oPr.TableCellW),
			textDirection: sTextDir,
			vAlign:        sVAlign,
			vMerge:        sVMerge,
			type:          "tableCellPr"
		}
	};
	WriterToJSON.prototype.SerTableCell = function(oCell, aComplexFieldsToSave, oMapCommentsInfo)
	{
		if (!oCell)
			return oCell;

		return {
			content: this.SerDocContent(oCell.Content, aComplexFieldsToSave, oMapCommentsInfo),
			tcPr:    this.SerTableCellPr(oCell.Pr),
			id:      oCell.Id,
			type:    "tblCell"
		}
	};
	WriterToJSON.prototype.SerTableRow = function(oRow, aComplexFieldsToSave, oMapCommentsInfo)
	{
		if (!oRow)
			return oRow;

		var sReviewType = undefined;
		switch (oRow.ReviewType)
		{
			case reviewtype_Common:
				sReviewType = "common";
				break;
			case reviewtype_Remove:
				sReviewType = "remove";
				break;
			case reviewtype_Add:
				sReviewType = "add";
				break;
		}

		var oRowObj = {
			content: [],
			reviewInfo: this.SerReviewInfo(oRow.ReviewInfo),
			reviewType: sReviewType,
			trPr:    this.SerTableRowPr(oRow.Pr),
			type:    "tblRow"
		}
		
		for (var nCell = 0; nCell < oRow.Content.length; nCell++)
		{
			oRowObj.content.push(this.SerTableCell(oRow.Content[nCell], aComplexFieldsToSave, oMapCommentsInfo));
		}

		return oRowObj;
	};
	WriterToJSON.prototype.SerTableRowPr = function(oPr)
	{
		if (!oPr)
			return oPr;

		// rowJc
		var sRowJc;
		switch (oPr.Jc)
		{
			case align_Left:
				sRowJc = "start";
				break;
			case align_Center:
				sRowJc = "center";
				break;
			case align_Right:
				sRowJc = "end";
				break;
			default:
				sRowJc = undefined;
				break;
		}

		// rowHeight
		var oRowHeight = undefined;
		if (oPr.Height)
		{
			switch (oPr.Height.HRule)
			{
				case linerule_AtLeast:
					oRowHeight = {
						val:   private_MM2Twips(oPr.Height.Value),
						hRule: "atLeast"
					};
					break;
				case linerule_Auto:
					oRowHeight = {
						val:   oPr.Height.Value,
						hRule: "auto"
					};
					break;
				case linerule_Exact:
					oRowHeight = {
						val:   private_MM2Twips(oPr.Height.Value),
						hRule: "exact"
					};
					break;
			}
		}

		return {
			cantSplit:      oPr.CantSplit,
			gridAfter:      oPr.GridAfter,
			gridBefore:     oPr.GridBefore,
			jc:             sRowJc,
			tblCellSpacing: oPr.TableCellSpacing ? private_MM2Twips(oPr.TableCellSpacing) : oPr.TableCellSpacing,
			tblHeader:      oPr.TableHeader,
			trHeight:       oRowHeight,
			trPrChange:     this.SerTableRowPr(oPr.PrChange),
			wAfter:         this.SerTableMeasurement(oPr.WAfter),
			wBefore:        this.SerTableMeasurement(oPr.WBefore),
			type:           "tableRowPr"
		}
	};
	WriterToJSON.prototype.SerBorder = function(oBorder)
	{
		if (!oBorder)
			return oBorder;

		var sBorderType = "none";
		if (oBorder.Value === border_Single)
			sBorderType = "single";

		return {
			color: oBorder.Color ? {
				auto: oBorder.Color.Auto,
				r:    oBorder.Color.r,
				g:    oBorder.Color.g,
				b:    oBorder.Color.b
			} : oBorder.Color,

			lineRef: oBorder.LineRef ? {
				idx:   oBorder.LineRef.idx,
				color: this.SerColor(oBorder.LineRef.Color)
			} : oBorder.LineRef,

			sz:         private_MM2Twips(oBorder.Size),
			space:      private_MM2Twips(oBorder.Space),
			themeColor: this.SerFill(oBorder.Unifill),
			value:      sBorderType
		}
	};
	WriterToJSON.prototype.SerDocContent = function(oDocContent, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo, bAllCompFields)
	{
		return {
			bPresentation: oDocContent.bPresentation,
			content:       this.SerContent(oDocContent.Content, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo, bAllCompFields),
			type:          "docContent"
		}
	};
	WriterToJSON.prototype.SerDrawingDocContent = function(oDocContent, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo, bAllCompFields)
	{
		return {
			content: this.SerContent(oDocContent.Content, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo, bAllCompFields),
			type:    "drawingDocContent"
		}
	};
	WriterToJSON.prototype.SerContent = function(aContent, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo, bAllCompFields)
	{
		var aResult = [];

		if (!aComplexFieldsToSave)
			aComplexFieldsToSave = this.GetComplexFieldsToSave(aContent, undefined, undefined, bAllCompFields);

		if (!oMapCommentsInfo)
			oMapCommentsInfo = this.GetMapCommentsInfo(aContent, oMapCommentsInfo);

		if (!oMapBookmarksInfo)
			oMapBookmarksInfo = this.GetMapBookmarksInfo(aContent, oMapBookmarksInfo);
		
		var TempElm = null;
		for (var nElm = 0; nElm < aContent.length; nElm++)
		{
			TempElm = aContent[nElm];

			if (TempElm instanceof AscCommonWord.Paragraph)
				aResult.push(this.SerParagraph(TempElm, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo));
			else if (TempElm instanceof AscCommonWord.CTable)
				aResult.push(this.SerTable(TempElm, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo));
			else if (TempElm instanceof AscCommonWord.CBlockLevelSdt)
				aResult.push(this.SerBlockLvlSdt(TempElm, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo));
		}

		return aResult;
	};
	WriterToJSON.prototype.SerParaPr = function(oParaPr)
	{
		if (!oParaPr)
			return oParaPr;

		var oStyles = private_GetStyles();
		// paragraph style
		var oParaStyle   = oParaPr.PStyle ? oStyles.Get(oParaPr.PStyle) : undefined;
		
		// horizontal align
		var sJc = undefined;
		switch (oParaPr.Jc)
		{
			case AscCommon.align_Right:
				sJc = "end";
				break;
			case AscCommon.align_Left:
				sJc = "start";
				break;
			case AscCommon.align_Center:
				sJc = "center";
				break;
			case AscCommon.align_Justify:
				sJc = "both";
				break;
			case AscCommon.align_Distributed:
				sJc = "distribute";
				break;
		}

		return {
			contextualSpacing: oParaPr.ContextualSpacing,
			framePr:           this.SerFramePr(oParaPr.FramePr),
			ind:               this.SerParaInd(oParaPr.Ind),
			jc:                sJc,
			keepLines:         oParaPr.KeepLines,
			keepNext:          oParaPr.KeepNext,
			numPr:             this.SerNumPr(oParaPr.NumPr),
			outlineLvl:        oParaPr.OutlineLvl,
			pageBreakBefore:   oParaPr.PageBreakBefore,
			defaultRunPr:      this.SerTextPr(oParaPr.DefaultRunPr),
			lvl:               oParaPr.Lvl,
			defTabSz:          oParaPr.DefaultTab != undefined ? private_MM2EMU(oParaPr.DefaultTab) : oParaPr.DefaultTab,

			pBdr: oParaPr.Brd ? 
			{
				between: this.SerBorder(oParaPr.Brd.Between),
				bottom:  this.SerBorder(oParaPr.Brd.Bottom),
				left:    this.SerBorder(oParaPr.Brd.Left),
				right:   this.SerBorder(oParaPr.Brd.Right),
				top:     this.SerBorder(oParaPr.Brd.Top)
			} : oParaPr.Brd,

			pPrChange:    this.SerParaPr(oParaPr.PrChange),
			pStyle:       this.SerStyle(oParaStyle),
			shd:          this.SerShd(oParaPr.Shd),
			spacing:      this.SerParaSpacing(oParaPr.Spacing),
			tabs:         this.SerTabs(oParaPr.Tabs),
			widowControl: oParaPr.WidowControl,
			bullet:       this.SerBullet(oParaPr.Bullet),
			type:         "paraPr"
		}
	};
	WriterToJSON.prototype.SerFramePr = function(oFramePr)
	{
		if (!oFramePr)
			return oFramePr;
		
		var sDropCapType = undefined;
		switch (oFramePr.DropCap)
		{
			case Asc.c_oAscDropCap.None:
				sDropCapType = "none";
				break;
			case Asc.c_oAscDropCap.Drop:
				sDropCapType = "drop";
				break;
			case Asc.c_oAscDropCap.Margin:
				sDropCapType = "margin";
				break;
		}

		var sHAnchor = undefined;
		switch (oFramePr.HAnchor)
		{
			case Asc.c_oAscHAnchor.Margin:
				sHAnchor = "margin";
				break;
			case Asc.c_oAscHAnchor.Page:
				sHAnchor = "page";
				break;
			case Asc.c_oAscHAnchor.Text:
				sHAnchor = "text";
				break;
		}

		var sVAnchor = undefined;
		switch (oFramePr.VAnchor)
		{
			case Asc.c_oAscHAnchor.Margin:
				sVAnchor = "margin";
				break;
			case Asc.c_oAscHAnchor.Page:
				sVAnchor = "page";
				break;
			case Asc.c_oAscHAnchor.Text:
				sVAnchor = "text";
				break;
		}

		var sLineRule = undefined;
		switch (oFramePr.HRule)
		{
			case Asc.linerule_AtLeast:
				sLineRule = "atLeast";
				break;
			case Asc.linerule_Auto:
				sLineRule = "auto";
				break;
			case Asc.linerule_Exact:
				sLineRule = "exact";
				break;
		}

		var sWrapType = undefined;
		switch (oFramePr.Wrap)
		{
			case AscCommonWord.wrap_Around:
				sWrapType = "around";
				break;
			case AscCommonWord.wrap_Auto:
				sWrapType = "auto";
				break;
			case AscCommonWord.wrap_None:
				sWrapType = "none";
				break;
			case AscCommonWord.wrap_NotBeside:
				sWrapType = "notBeside";
				break;
			case AscCommonWord.wrap_Through:
				sWrapType = "through";
				break;
			case AscCommonWord.wrap_Tight:
				sWrapType = "tight";
				break;
		}

		var sXAlign = undefined;
		switch (oFramePr.XAlign)
		{
			case Asc.c_oAscXAlign.Center:
				sXAlign = "center";
				break;
			case Asc.c_oAscXAlign.Inside:
				sXAlign = "inside";
				break;
			case Asc.c_oAscXAlign.Left:
				sXAlign = "left";
				break;
			case Asc.c_oAscXAlign.Outside:
				sXAlign = "outside";
				break;
			case Asc.c_oAscXAlign.Right:
				sXAlign = "right";
				break;
		}

		var sYAlign = undefined;
		switch (oFramePr.YAlign)
		{
			case Asc.c_oAscYAlign.Bottom:
				sYAlign = "bottom";
				break;
			case Asc.c_oAscYAlign.Center:
				sYAlign = "center";
				break;
			case Asc.c_oAscYAlign.Inline:
				sYAlign = "inline";
				break;
			case Asc.c_oAscYAlign.Inside:
				sYAlign = "inside";
				break;
			case Asc.c_oAscYAlign.Outside:
				sYAlign = "outside";
				break;
			case Asc.c_oAscYAlign.Top:
				sYAlign = "top";
				break;
		}

		return {
			dropCap: sDropCapType,
			h:       typeof(oFramePr.H) === "number" ? private_MM2Twips(oFramePr.H) : oFramePr.H,
			hAnchor: sHAnchor,
			hRule:   sLineRule,
			hSpace:  typeof(oFramePr.HSpace) === "number" ? private_MM2Twips(oFramePr.HSpace) : oFramePr.HSpace,
			lines:   oFramePr.Lines,
			vAnchor: sVAnchor,
			vSpace:  typeof(oFramePr.VSpace) === "number" ? private_MM2Twips(oFramePr.VSpace) : oFramePr.VSpace,
			w:       typeof(oFramePr.W) === "number" ? private_MM2Twips(oFramePr.W) : oFramePr.W,
			wrap:    sWrapType,
			x:       typeof(oFramePr.X) === "number" ? private_MM2Twips(oFramePr.X) : oFramePr.X,
			xAlign:  sXAlign,
			y:       typeof(oFramePr.Y) === "number" ? private_MM2Twips(oFramePr.Y) : oFramePr.Y,
			yAlign:  sYAlign
		}
	};
	WriterToJSON.prototype.SerParaInd = function(oParaInd)
	{
		if (!oParaInd)
			return oParaInd;
		
		return {
			left:      oParaInd.Left != undefined ? private_MM2Twips(oParaInd.Left)      : oParaInd.Left,
			right:     oParaInd.Right != undefined ? private_MM2Twips(oParaInd.Right)     : oParaInd.Right,
			firstLine: oParaInd.FirstLine != undefined ? private_MM2Twips(oParaInd.FirstLine) : oParaInd.FirstLine
		}
	};
	WriterToJSON.prototype.SerNumPr = function(oNumPr)
	{
		if (!oNumPr)
			return oNumPr;

		return {
			ilvl:  oNumPr.Lvl,
			numId: oNumPr.NumId
		}
	};
	WriterToJSON.prototype.SerParaSpacing = function(oParaSpacing)
	{
		if (!oParaSpacing)
			return oParaSpacing;
		
		// spacing
		var oSpacing = {
			before:            oParaSpacing.Before ? private_MM2Twips(oParaSpacing.Before) : oParaSpacing.Before,
			beforePct:         oParaSpacing.BeforePct,
			beforeAutoSpacing: oParaSpacing.BeforeAutoSpacing !== undefined ? (oParaSpacing.BeforeAutoSpacing === true ? "on" : "off") : oParaSpacing.BeforeAutoSpacing,
			after:             oParaSpacing.After ? private_MM2Twips(oParaSpacing.After) : oParaSpacing.After,
			afterPct:          oParaSpacing.AfterPct,
			afterAutoSpacing:  oParaSpacing.AfterAutoSpacing !== undefined ? (oParaSpacing.AfterAutoSpacing === true ? "on" : "off") : oParaSpacing.AfterAutoSpacing
		};

		switch (oParaSpacing.LineRule)
		{
			case linerule_AtLeast:
				oSpacing.lineRule = "atLeast";
				oSpacing.line     = private_MM2Twips(oParaSpacing.Line)
				break;
			case linerule_Auto:
				oSpacing.lineRule = "auto";
				oSpacing.line     = private_MM2Twips(oParaSpacing.Line)
				break;
			case linerule_Exact:
				oSpacing.lineRule = "exact";
				oSpacing.line     = private_MM2Twips(oParaSpacing.Line)
				break;
			default:
				oSpacing.lineRule = undefined;
				oSpacing.line     = undefined;
				break;
		}

		return oSpacing;
	};
	WriterToJSON.prototype.SerBlockLvlSdt = function(oSdt, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo)
	{
		if (!aComplexFieldsToSave)
			aComplexFieldsToSave = this.GetComplexFieldsToSave(oSdt.Content.Content);

		if (!oMapCommentsInfo)
			oMapCommentsInfo = this.GetMapCommentsInfo(oSdt.Content.Content, oMapCommentsInfo);

		if (!oMapBookmarksInfo)
			oMapBookmarksInfo = this.GetMapBookmarksInfo(oSdt.Content.Content, oMapBookmarksInfo);

		return {
			sdtPr: this.SerSdtPr(oSdt.Pr),
			sdtContent: this.SerDocContent(oSdt.Content, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo),
			type: "blockLvlSdt"
		}
	};
	WriterToJSON.prototype.SerInlineLvlSdt = function(oSdt, aComplexFieldsToSave)
	{
		var oInlineSdt = 
		{
			sdtPr: this.SerSdtPr(oSdt.Pr),
			content: [],
			type: "inlineLvlSdt"
		}

		var TempElm = null;
		for (var nElm = 0; nElm < oSdt.Content.length; nElm++)
		{
			TempElm = oSdt.Content[nElm];

			if (TempElm instanceof AscCommonWord.ParaRun)
				oInlineSdt.content.push(this.SerParaRun(TempElm, aComplexFieldsToSave));
			else if (TempElm instanceof AscCommonWord.ParaHyperlink)
				oInlineSdt.content.push(this.SerHyperlink(TempElm, aComplexFieldsToSave));
			else if (TempElm instanceof AscCommonWord.CInlineLevelSdt)
				oInlineSdt.content.push(this.SerInlineLvlSdt(TempElm, aComplexFieldsToSave));
		}

		return oInlineSdt;
	};
	WriterToJSON.prototype.SerParagraph = function(oPara, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo)
	{
		var oParaObject = {
			bFromDocument: oPara.bFromDocument,
			pPr:           this.SerParaPr(oPara.Pr),
			rPr:           this.SerTextPr(oPara.TextPr.Value),
			content:       [],
			changes:       [],
			type:          "paragraph"
		};
		
		if (!oMapCommentsInfo)
			oMapCommentsInfo = this.GetMapCommentsInfo([oPara], oMapCommentsInfo);

		if (!oMapBookmarksInfo && AscCommonWord.CParagraphBookmark)
			oMapBookmarksInfo = this.GetMapBookmarksInfo([oPara], oMapBookmarksInfo);

		if (!aComplexFieldsToSave)
			aComplexFieldsToSave = this.GetComplexFieldsToSave([oPara]);

		oParaObject.pPr.sectPr = this.SerSectionPr(oPara.SectPr);

		var oTempElm    = null;
		var oTempResult = null;

		// numbering
		var oNumPr           = null;
		var oGlobalNumbering = null;
		var oParaParent      = oPara.GetParent();
		var oLogicDocument   = private_GetLogicDocument();

		if (!(oParaParent instanceof AscFormat.CDrawingDocContent) && oLogicDocument instanceof AscCommonWord.CDocument)
		{
			oNumPr           = oPara.GetNumPr();
			oGlobalNumbering = oLogicDocument.GetNumbering();
		}
		
		var oNum = null;
		if (oNumPr && oGlobalNumbering)
			oNum = oGlobalNumbering.GetNum(oNumPr.NumId);

		if (oNum)
			oParaObject.numbering = this.SerNumbering(oNum);

		for (var nElm = 0; nElm < oPara.Content.length; nElm++)
		{
			oTempElm = oPara.Content[nElm];
			
			if (oTempElm instanceof AscCommonWord.ParaRun && !oTempElm.FieldType && !oTempElm.Guid)
			{
				var oRunObject = this.SerParaRun(oTempElm, aComplexFieldsToSave);
				//if (oRunObject.content.length !== 0)
				oParaObject.content.push(oRunObject);
			}
			else if (oTempElm instanceof AscCommonWord.ParaMath)
				oParaObject.content.push(this.SerParaMath(oTempElm));
			else if (oTempElm instanceof AscCommonWord.ParaHyperlink)
				oParaObject.content.push(this.SerHyperlink(oTempElm));
			else if (oTempElm instanceof AscCommon.ParaComment)
			{
				oTempResult = this.SerParaComment(oTempElm, oMapCommentsInfo);
				if (oTempResult)
					oParaObject.content.push(oTempResult);
			}
			else if (AscCommonWord.CParagraphBookmark && oTempElm instanceof AscCommonWord.CParagraphBookmark)
			{
				oTempResult = this.SerParaBookmark(oTempElm, oMapBookmarksInfo);
				if (oTempResult)
					oParaObject.content.push(oTempResult);
			}
			else if (oTempElm instanceof AscCommonWord.CInlineLevelSdt)
			{
				var oSdt = this.SerInlineLvlSdt(oTempElm, aComplexFieldsToSave);
				if (oSdt)
					oParaObject.content.push(oSdt);
			}
			else if (oTempElm instanceof AscCommon.CParaRevisionMove)
			{
				oParaObject.content.push(this.SerRevisionMove(oTempElm));
			}
			else if (oTempElm instanceof AscCommonWord.CPresentationField && oTempElm.FieldType && oTempElm.Guid)
			{
				oParaObject.content.push(this.SerPresField(oTempElm));
			}
		}

		// Revisions
		var aChanges = oLogicDocument.TrackRevisionsManager ? oLogicDocument.TrackRevisionsManager.GetElementChanges(oPara.Id) : [];

		for (var nChange = 0; nChange < aChanges.length; nChange++)
		{
			oParaObject.changes.push(this.SerRevisionChange(aChanges[nChange]));
		}

		return oParaObject;
	};
	WriterToJSON.prototype.SerPresField = function(oPresField)
	{
		var oPresFieldObj = this.SerParaRun(oPresField);

		oPresFieldObj.type = "presField";
		oPresFieldObj.fldType = oPresField.FieldType;

		return oPresFieldObj;
	};
	WriterToJSON.prototype.SerFootEndnote = function(oFootEndnote)
	{
		var aComplexFieldsToSave = this.GetComplexFieldsToSave(oFootEndnote.Content);
		var oMapCommentsInfo     = this.GetMapCommentsInfo(oFootEndnote.Content, []);
		var oMapBookmarksInfo    = this.GetMapBookmarksInfo(oFootEndnote.Content, []);

		var oFootEndnoteObj = {
			content:          [],
			customMarkFollow: oFootEndnote.customMarkFollow,
			hint:             oFootEndnote.Hint,
			id:               oFootEndnote.Id,
			number:           oFootEndnote.Number,
			columnsCount:     oFootEndnote.ColumnsCount,
			sectPr:           this.SerSectionPr(oFootEndnote.SectPr),
			type:             oFootEndnote.parent instanceof CFootnotesController ? "footnote" : "endnote"
		};

		var TempElm = null;
		for (var nElm = 0; nElm < oFootEndnote.Content.length; nElm++)
		{
			TempElm = oFootEndnote.Content[nElm];

			if (TempElm instanceof AscCommonWord.Paragraph)
				oFootEndnoteObj.content.push(this.SerParagraph(TempElm, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo));
			else if (TempElm instanceof AscCommonWord.CTable)
				oFootEndnoteObj.content.push(this.SerTable(TempElm, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo));
			else if (TempElm instanceof AscCommonWord.CBlockLevelSdt)
				oFootEndnoteObj.content.push(this.SerBlockLvlSdt(TempElm, aComplexFieldsToSave, oMapCommentsInfo, oMapBookmarksInfo));
		}

		return oFootEndnoteObj;
	};
	WriterToJSON.prototype.SerRevisionMove = function(oMove)
	{
		if (!oMove)
			return oMove;

		return {
			start:      oMove.Start,
			id:         oMove.Id,
			from:       oMove.From,
			name:       oMove.Name,
			reviewInfo: this.SerReviewInfo(oMove.ReviewInfo),
			type:       "revisionMove"
		}
	};
	WriterToJSON.prototype.SerRevisionChange = function(oChange)
	{
		if (!oChange)
			return oChange;

		var sChangeType = undefined;
		switch (oChange.Type)
		{
			case Asc.c_oAscRevisionsChangeType.Unknown:
				sChangeType = "unknown";
				break;
			case Asc.c_oAscRevisionsChangeType.TextAdd:
				sChangeType = "textAdd";
				break;
			case Asc.c_oAscRevisionsChangeType.TextRem:
				sChangeType = "textRem";
				break;
			case Asc.c_oAscRevisionsChangeType.ParaAdd:
				sChangeType = "paraAdd";
				break;
			case Asc.c_oAscRevisionsChangeType.ParaRem:
				sChangeType = "paraRem";
				break;
			case Asc.c_oAscRevisionsChangeType.TextPr:
				sChangeType = "textPr";
				break;
			case Asc.c_oAscRevisionsChangeType.ParaPr:
				sChangeType = "paraPr";
				break;
			case Asc.c_oAscRevisionsChangeType.TablePr:
				sChangeType = "tablePr";
				break;
			case Asc.c_oAscRevisionsChangeType.RowsAdd:
				sChangeType = "rowsAdd";
				break;
			case Asc.c_oAscRevisionsChangeType.RowsRem:
				sChangeType = "rowsRem";
				break;
			case Asc.c_oAscRevisionsChangeType.MoveMark:
				sChangeType = "moveMark";
				break;
			case Asc.c_oAscRevisionsChangeType.MoveMarkRemove:
				sChangeType = "moveMarkRemove";
				break;
		}

		var sMoveType = undefined;
		switch (oChange.MoveType)
		{
			case Asc.c_oAscRevisionsMove.NoMove:
				sMoveType = "noMove";
				break;
			case Asc.c_oAscRevisionsMove.MoveTo:
				sMoveType = "moveTo";
				break;
			case Asc.c_oAscRevisionsMove.MoveFrom:
				sMoveType = "moveFrom";
				break;
		}

		var changeValue;

		if (oChange.Value instanceof AscCommonWord.CTextPr)
			changeValue = this.SerTextPr(oChange.Value);
		else if (oChange.Value instanceof CParaRevisionMove)
			changeValue = "";
		else
			changeValue = oChange.Value;

		var startPos;
		if (oChange.StartPos instanceof CParagraphContentPos)
		{
			startPos = {
				data:         oChange.StartPos.Data,
				death:        oChange.StartPos.Depth,
				bPlaceholder: oChange.StartPos.bPlaceholder
			}
		}
		else
			startPos = oChange.StartPos;

		var endPos;
		if (oChange.EndPos instanceof CParagraphContentPos)
		{
			endPos = {
				data:         oChange.EndPos.Data,
				death:        oChange.EndPos.Depth,
				bPlaceholder: oChange.EndPos.bPlaceholder
			}
		}
		else
			endPos = oChange.EndPos;

		return {
			date:      oChange.DateTime,
			start:     startPos,
			end:       endPos,
			userId:    oChange.UserId,
			autor:     oChange.UserName,
			value:     changeValue,
			moveType:  sMoveType,
			type:      sChangeType
		}
	};
	WriterToJSON.prototype.SerDocument = function(oDocument)
	{
		var oDocObject = {
			content: [],
			hdrFtr:  [],
			sectPr:  this.SerSectionPr(oDocument.SectPr),
			changes: [],
			type:    "document"
		};

		var aComplexFieldsToSave = this.GetComplexFieldsToSave(oDocument.Content);

		var TempElm = null;
		for (var nElm = 0; nElm < oDocument.Content.length; nElm++)
		{
			TempElm = oDocument.Content[nElm];

			if (TempElm instanceof AscCommonWord.Paragraph)
				oDocObject.content.push(this.SerParagraph(TempElm, aComplexFieldsToSave));
			else if (TempElm instanceof AscCommonWord.CTable)
				oDocObject.content.push(this.SerTable(TempElm, aComplexFieldsToSave));
			else if (TempElm instanceof AscCommonWord.CBlockLevelSdt)
				oDocObject.content.push(this.SerBlockLvlSdt(TempElm, aComplexFieldsToSave));
		}

		// header and footer
		// var oHdrFtr = oDocument.GetHdrFtr();
		// for (var nPage = 0; nPage < oHdrFtr.Pages.length; nPage++)
		// 	oDocObject.hdrFtr.push(this.(oHdrFtr.Pages[nPage]));

		return oDocObject;
	};
	WriterToJSON.prototype.SerTheme = function(oTheme)
	{
		var aExtraClrSchemeLst = [];
		for (var nElm = 0; nElm < oTheme.extraClrSchemeLst.length; nElm++)
			aExtraClrSchemeLst.push(this.SerExtraClrScheme(oTheme.extraClrSchemeLst[nElm]));

		var oThemeObj = {
			custClrLst: this.SerColorMapOvr(oTheme.clrMap), // ??? maybe not supported
			name:       oTheme.name,
			objectDefaults: {
				lnDef: this.SerDefSpDefinition(oTheme.lnDef), // AscFormat.DefaultShapeDefinition
				spDef: this.SerDefSpDefinition(oTheme.spDef),
				txDef: this.SerDefSpDefinition(oTheme.txDef)
			},
			themeElements: {
				clrScheme:  this.SerClrScheme(oTheme.themeElements.clrScheme),
				fmtScheme:  this.SerFmtScheme(oTheme.themeElements.fmtScheme),
				fontScheme: this.SerFontScheme(oTheme.themeElements.fontScheme)
			},

			extraClrSchemeLst: aExtraClrSchemeLst, // AscFormat.ExtraClrScheme:
			isThemeOverride:   oTheme.isThemeOverride,
			id:                oTheme.Id
		}

		// мамим, чтобы не записывать несколько раз
		this.themesMap[oTheme.Id] = oThemeObj;

		return oThemeObj;
	};
	WriterToJSON.prototype.SerClrScheme = function(oClrScheme)
	{
		if (!oClrScheme)
			return oClrScheme;

		return {
			name:     oClrScheme.name,
			dk1:      this.SerColor(oClrScheme.colors[0]),
			lt1:      this.SerColor(oClrScheme.colors[1]),
			dk2:      this.SerColor(oClrScheme.colors[2]),
			lt2:      this.SerColor(oClrScheme.colors[3]),
			accent1:  this.SerColor(oClrScheme.colors[4]),
			accent2:  this.SerColor(oClrScheme.colors[5]),
			accent3:  this.SerColor(oClrScheme.colors[8]),
			accent4:  this.SerColor(oClrScheme.colors[9]),
			accent5:  this.SerColor(oClrScheme.colors[10]),
			accent6:  this.SerColor(oClrScheme.colors[11]),
			hlink:    this.SerColor(oClrScheme.colors[12]),
			folHlink: this.SerColor(oClrScheme.colors[13]),
			type:     "clrScheme"
		}
	};
	WriterToJSON.prototype.SerFmtScheme = function(oFmtScheme)
	{
		if (!oFmtScheme)
			return oFmtScheme;

		var aBgFillStyleLst = [];
		for (var nFill = 0; nFill < oFmtScheme.bgFillStyleLst.length; nFill++)
			aBgFillStyleLst.push(this.SerFill(oFmtScheme.bgFillStyleLst[nFill]));

		var aEffectStyleLst = []; // пока не поддерживаем

		var aFillStyleLst = [];
		for (nFill = 0; nFill < oFmtScheme.fillStyleLst.length; nFill++)
			aFillStyleLst.push(this.SerFill(oFmtScheme.fillStyleLst[nFill]));

		var aLnStyleLst = [];
		for (var nLn = 0; nLn < oFmtScheme.lnStyleLst.length; nLn++)
			aLnStyleLst.push(this.SerLn(oFmtScheme.lnStyleLst[nLn]));

		return {
			name:           oFmtScheme.name,
			bgFillStyleLst: aBgFillStyleLst,
			fillStyleLst:   aFillStyleLst,
			lnStyleLst:     aLnStyleLst,
			type:           "fmtScheme"
		}
	};
	WriterToJSON.prototype.SerFontScheme = function(oFontScheme)
	{
		if (!oFontScheme)
			return oFontScheme;

		return {
			name:      oFontScheme.name,
			majorFont: this.SerFontCollection(oFontScheme.majorFont),
			minorFont: this.SerFontCollection(oFontScheme.minorFont),
			type:      "fontScheme"
		}
	};
	WriterToJSON.prototype.SerFontCollection = function(oFontCollection)
	{
		if (!oFontCollection)
			return oFontCollection;

		return {
			cs:    oFontCollection.cs,
			ea:    oFontCollection.ea,
			latin: oFontCollection.latin
		}
	};
	WriterToJSON.prototype.SerExtraClrScheme = function(oExtraClrScheme)
	{
		if (!oExtraClrScheme)
			return oExtraClrScheme;

		return {
			clrMap:    this.SerColorMapOvr(oExtraClrScheme.clrMap),
			clrScheme: this.SerClrScheme(oExtraClrScheme.clrScheme)
		}
	};
	WriterToJSON.prototype.SerDefSpDefinition = function(oDefinition)
	{
		if (!oDefinition)
			return oDefinition;

		return {
			bodyPr:   this.SerBodyPr(oDefinition.bodyPr),
			lstStyle: this.SerLstStyle(oDefinition.lstStyle),
			spPr:     this.SerSpPr(oDefinition.spPr),
			style:    this.SerSpStyle(oDefinition.style)
		}
	};
	ReaderFromJSON.prototype.ThemeFromJSON = function(oParsedTheme)
	{
		var oTheme = new AscFormat.CTheme();
		for (var nElm = 0; nElm < oParsedTheme.extraClrSchemeLst.length; nElm++)
			oTheme.addExtraClrSceme(this.ExtraClrSchemeFromJSON(oParsedTheme.extraClrSchemeLst[nElm]));

		oTheme.setName(oParsedTheme.name);
		oParsedTheme.objectDefaults.lnDef && oTheme.setLnDef(this.DefSpDefinitionFromJSON(oParsedTheme.objectDefaults.lnDef));
		oParsedTheme.objectDefaults.spDef && oTheme.setSpDef(this.DefSpDefinitionFromJSON(oParsedTheme.objectDefaults.spDef));
		oParsedTheme.objectDefaults.txDef && oTheme.setTxDef(this.DefSpDefinitionFromJSON(oParsedTheme.objectDefaults.txDef));
		oParsedTheme.themeElements.clrScheme && oTheme.setColorScheme(this.ClrSchemeFromJSON(oParsedTheme.themeElements.clrScheme));
		oParsedTheme.themeElements.fmtScheme && oTheme.setFormatScheme(this.FmtSchemeFromJSON(oParsedTheme.themeElements.fmtScheme));
		oParsedTheme.themeElements.fontScheme && oTheme.setFontScheme(this.FontSchemeFromJSON(oParsedTheme.themeElements.fontScheme));
		oTheme.setIsThemeOverride(oParsedTheme.isThemeOverride);

		this.themesMap[oParsedTheme.Id] = oTheme;

		return oTheme;
	};
	ReaderFromJSON.prototype.ThemeFromJSON = function(oParsedTheme)
	{
		var oTheme = new AscFormat.CTheme();
		for (var nElm = 0; nElm < oParsedTheme.extraClrSchemeLst.length; nElm++)
			oTheme.addExtraClrSceme(this.ExtraClrSchemeFromJSON(oParsedTheme.extraClrSchemeLst[nElm]));

		oTheme.setName(oParsedTheme.name);
		oParsedTheme.objectDefaults.lnDef && oTheme.setLnDef(this.DefSpDefinitionFromJSON(oParsedTheme.objectDefaults.lnDef));
		oParsedTheme.objectDefaults.spDef && oTheme.setSpDef(this.DefSpDefinitionFromJSON(oParsedTheme.objectDefaults.spDef));
		oParsedTheme.objectDefaults.txDef && oTheme.setTxDef(this.DefSpDefinitionFromJSON(oParsedTheme.objectDefaults.txDef));
		oParsedTheme.themeElements.clrScheme && oTheme.setColorScheme(this.ClrSchemeFromJSON(oParsedTheme.themeElements.clrScheme));
		oParsedTheme.themeElements.fmtScheme && oTheme.setFormatScheme(this.FmtSchemeFromJSON(oParsedTheme.themeElements.fmtScheme));
		oParsedTheme.themeElements.fontScheme && oTheme.setFontScheme(this.FontSchemeFromJSON(oParsedTheme.themeElements.fontScheme));
		oTheme.setIsThemeOverride(oParsedTheme.isThemeOverride);

		this.themesMap[oParsedTheme.Id] = oTheme;

		return oTheme;
	};
	ReaderFromJSON.prototype.ExtraClrSchemeFromJSON = function(oParsedExtrClrScheme)
	{
		var oExtraClrScheme = new AscFormat.ExtraClrScheme();
		oParsedExtrClrScheme.clrMap && oExtraClrScheme.setClrMap(this.ColorMapOvrFromJSON(oParsedExtrClrScheme.clrMap));
		oParsedExtrClrScheme.clrScheme && oExtraClrScheme.setClrScheme(this.ClrSchemeFromJSON(oParsedExtrClrScheme.clrScheme));

		return oExtraClrScheme;
	};
	ReaderFromJSON.prototype.ClrSchemeFromJSON = function(oParsedClrScheme)
	{
		var oClrScheme = new AscFormat.ClrScheme();
		oClrScheme.setName(oParsedClrScheme.name);
		oParsedClrScheme.dk1 && oClrScheme.addColor(0, this.ColorFromJSON(oParsedClrScheme.dk1));
		oParsedClrScheme.lt1 && oClrScheme.addColor(1, this.ColorFromJSON(oParsedClrScheme.lt1));
		oParsedClrScheme.dk2 && oClrScheme.addColor(2, this.ColorFromJSON(oParsedClrScheme.dk2));
		oParsedClrScheme.lt2 && oClrScheme.addColor(3, this.ColorFromJSON(oParsedClrScheme.lt2));
		oParsedClrScheme.accent1 && oClrScheme.addColor(4, this.ColorFromJSON(oParsedClrScheme.accent1));
		oParsedClrScheme.accent2 && oClrScheme.addColor(5, this.ColorFromJSON(oParsedClrScheme.accent2));
		oParsedClrScheme.accent3 && oClrScheme.addColor(8, this.ColorFromJSON(oParsedClrScheme.accent3));
		oParsedClrScheme.accent4 && oClrScheme.addColor(9, this.ColorFromJSON(oParsedClrScheme.accent4));
		oParsedClrScheme.accent5 && oClrScheme.addColor(10, this.ColorFromJSON(oParsedClrScheme.accent5));
		oParsedClrScheme.accent6 && oClrScheme.addColor(11, this.ColorFromJSON(oParsedClrScheme.accent6));
		oParsedClrScheme.hlink && oClrScheme.addColor(12, this.ColorFromJSON(oParsedClrScheme.hlink));
		oParsedClrScheme.folHlink && oClrScheme.addColor(13, this.ColorFromJSON(oParsedClrScheme.folHlink));

		return oClrScheme;
	};
	ReaderFromJSON.prototype.FmtSchemeFromJSON = function(oParsedFmtScheme)
	{
		var oFmtScheme = new AscFormat.FmtScheme();

		for (var nBgFill = 0; nBgFill < oParsedFmtScheme.bgFillStyleLst.length; nBgFill++)
			oFmtScheme.addBgFillToStyleLst(this.FillFromJSON(oParsedFmtScheme.bgFillStyleLst[nBgFill]));

		for (var nFill = 0; nFill < oParsedFmtScheme.fillStyleLst.length; nFill++)
			oFmtScheme.addFillToStyleLst(this.FillFromJSON(oParsedFmtScheme.fillStyleLst[nFill]));
		
		for (nFill = 0; nFill < oParsedFmtScheme.lnStyleLst.length; nFill++)
			oFmtScheme.addLnToStyleLst(this.LnFromJSON(oParsedFmtScheme.lnStyleLst[nFill]));

		oParsedFmtScheme.name && oFmtScheme.setName(oParsedFmtScheme.name);

		return oFmtScheme;
	};
	ReaderFromJSON.prototype.FontSchemeFromJSON = function(oParsedFntScheme)
	{
		var oFontScheme = new AscFormat.FontScheme();
		this.FontCollectionFromJSON(oParsedFntScheme.majorFont, "major", oFontScheme);
		this.FontCollectionFromJSON(oParsedFntScheme.minorFont, "minor", oFontScheme);

		oParsedFntScheme.name && oFontScheme.setName(oParsedFntScheme.name);

		return oFontScheme;
	};
	ReaderFromJSON.prototype.FontCollectionFromJSON = function(oParsedFntColl, sType, oParentFntScheme)
	{
		if (sType === "major")
		{
			oParentFntScheme.majorFont.setLatin(oParsedFntColl.latin);
			oParentFntScheme.majorFont.setEA(oParsedFntColl.ea);
			oParentFntScheme.majorFont.setCS(oParsedFntColl.cs);
		}
		if (sType === "minor")
		{
			oParentFntScheme.minorFont.setLatin(oParsedFntColl.latin);
			oParentFntScheme.minorFont.setEA(oParsedFntColl.ea);
			oParentFntScheme.minorFont.setCS(oParsedFntColl.cs);
		}
	};
	ReaderFromJSON.prototype.DefSpDefinitionFromJSON = function(oParsedDefSpDef)
	{
		var oDefSpDefinition = new AscFormat.DefaultShapeDefinition();
		oParsedDefSpDef.bodyPr   && oDefSpDefinition.setBodyPr(this.BodyPrFromJSON(oParsedDefSpDef.bodyPr));
		oParsedDefSpDef.lstStyle && oDefSpDefinition.setLstStyle(this.LstStyleFromJSON(oParsedDefSpDef.lstStyle));
		oParsedDefSpDef.spPr     && oDefSpDefinition.setSpPr(this.SpPrFromJSON(oParsedDefSpDef.spPr));
		oParsedDefSpDef.style    && oDefSpDefinition.setStyle(this.SpStyleFromJSON(oParsedDefSpDef.style));

		return oDefSpDefinition;
	};
	WriterToJSON.prototype.SerHeader = function(oHdr)
	{
		if (!oHdr)
			return oHdr;

		return {
			content: this.SerDocContent(oHdr.Content, undefined, undefined, undefined, true),
			type: "hdr"
		}
	};
	WriterToJSON.prototype.SerFooter = function(oFtr)
	{
		if (!oFtr)
			return oFtr;

		return {
			content: this.SerDocContent(oFtr.Content, undefined, undefined, undefined, true),
			type: "ftr"
		}
	};
	WriterToJSON.prototype.SerSectionPr = function(oSectionPr)
	{
		if (!oSectionPr)
			return oSectionPr;

		var sSectionType = undefined;
		switch(oSectionPr.Type)
		{
			case Asc.c_oAscSectionBreakType.NextPage:
				sSectionType = "nextPage";
				break;
			case Asc.c_oAscSectionBreakType.OddPage:
				sSectionType = "oddPage";
				break;
			case Asc.c_oAscSectionBreakType.EvenPage:
				sSectionType = "evenPage";
				break;
			case Asc.c_oAscSectionBreakType.Continuous:
				sSectionType = "continuous";
				break;
			case Asc.c_oAscSectionBreakType.Column:
				sSectionType = "nextColumn";
				break;
		}
		var oFooterReference = {
			first:   oSectionPr.FooterFirst   ? this.SerFooter(oSectionPr.FooterFirst)   : oSectionPr.FooterFirst,
			default: oSectionPr.FooterDefault ? this.SerFooter(oSectionPr.FooterDefault) : oSectionPr.FooterDefault,
			even:    oSectionPr.FooterEven    ? this.SerFooter(oSectionPr.FooterEven)    : oSectionPr.FooterEven,
		};
		var oHeaderReference = {
			first:   oSectionPr.HeaderFirst   ? this.SerHeader(oSectionPr.HeaderFirst)   : oSectionPr.HeaderFirst,
			default: oSectionPr.HeaderDefault ? this.SerHeader(oSectionPr.HeaderDefault) : oSectionPr.HeaderDefault,
			even:    oSectionPr.HeaderEven    ? this.SerHeader(oSectionPr.HeaderEven)    : oSectionPr.HeaderEven,
		}

		return {
			cols:            this.SerSectionColumns(oSectionPr.Columns),
			endnotePr:       this.SerEndNotePr(oSectionPr.EndnotePr),
			footerReference: oFooterReference,
			footnotePr:      this.SerFootnotePr(oSectionPr.FootnotePr),
			headerReference: oHeaderReference,
			lnNumType:       this.SerLnNumType(oSectionPr.LnNumType),
			pgBorders:       this.SerPageBorders(oSectionPr.Borders),
			pgMar:           this.SerPageMargins(oSectionPr.PageMargins),
			//pgNumType:       oSectionPr.PageNumType.Start,
			pgSz:            this.SerPageSize(oSectionPr.PageSize),
			rtlGutter:       oSectionPr.GutterRTL,
			titlePg:         oSectionPr.TitlePage,
			type:            sSectionType
		}
	};
	WriterToJSON.prototype.SerSectionColumns = function(oSectColumns)
	{
		if (!oSectColumns)
			return oSectColumns;

		var aCols = [];
		for (var nCol = 0; nCol < oSectColumns.Cols.length; nCol++)
			aCols.push(this.SerSectionCol(oSectColumns.Cols[nCol]));

		return {
			col:        aCols,
			equalWidth: oSectColumns.EqualWidth,
			num:        oSectColumns.Num,
			sep:        oSectColumns.Sep,
			space:      private_MM2Twips(oSectColumns.Space)
		}
	};
	WriterToJSON.prototype.SerSectionCol = function(oSectionCol)
	{
		if (!oSectionCol)
			return oSectionCol;

		return {
			space: private_MM2Twips(oSectionCol.Space),
			w:     private_MM2Twips(oSectionCol.W)
		}
	};
	WriterToJSON.prototype.SerEndNotePr = function(oEndNotePr)
	{
		if (!oEndNotePr)
			return oEndNotePr;

		var sNumRestart = undefined;
		switch(oEndNotePr.NumRestart)
		{
			case section_footnote_RestartContinuous:
				sNumRestart = "continuous";
				break;
			case section_footnote_RestartEachPage:
				sNumRestart = "eachPage";
				break;
			case section_footnote_RestartEachSect:
				sNumRestart = "eachSect";
				break;
		}

		var sEndPos = undefined;
		switch(oEndNotePr.Pos)
		{
			case Asc.c_oAscEndnotePos.DocEnd:
				sEndPos = "docEnd";
				break;
			case Asc.c_oAscEndnotePos.SectEnd:
				sEndPos = "sectEnd";
				break;
		}
		return {
			numFmt:     To_XML_c_oAscNumberingFormat(oEndNotePr.NumFormat),
			numRestart: sNumRestart,
			numStart:   oEndNotePr.NumStart,
			pos:        sEndPos
		}
	};
	WriterToJSON.prototype.SerFootnotePr = function(oFootnotePr)
	{
		if (!oFootnotePr)
			return oFootnotePr;

		var sNumRestart = undefined;
		switch(oFootnotePr.NumRestart)
		{
			case section_footnote_RestartContinuous:
				sNumRestart = "continuous";
				break;
			case section_footnote_RestartEachPage:
				sNumRestart = "eachPage";
				break;
			case section_footnote_RestartEachSect:
				sNumRestart = "eachSect";
				break;
		}

		var sEndPos = undefined;
		switch(oFootnotePr.Pos)
		{
			case Asc.c_oAscFootnotePos.BeneathText:
				sEndPos = "beneathText";
				break;
			case Asc.c_oAscFootnotePos.DocEnd:
				sEndPos = "docEnd";
				break;
			case Asc.c_oAscFootnotePos.PageBottom:
				sEndPos = "pgBottom";
				break;
			case Asc.c_oAscFootnotePos.SectEnd:
				sEndPos = "sectEnd";
				break;
		}
		return {
			numFmt:     To_XML_c_oAscNumberingFormat(oFootnotePr.NumFormat),
			numRestart: sNumRestart,
			numStart:   oFootnotePr.NumStart,
			pos:        sEndPos
		}
	};
	WriterToJSON.prototype.SerLnNumType = function(oLnNumType)
	{
		if (!oLnNumType)
			return oLnNumType;

		var sRestartType = undefined;
		switch(oLnNumType.Restart)
		{
			case Asc.c_oAscLineNumberRestartType.Continuous:
				sRestartType = "continuous";
				break;
			case Asc.c_oAscLineNumberRestartType.NewPage:
				sRestartType = "newPage";
				break;
			case Asc.c_oAscLineNumberRestartType.NewSection:
				sRestartType = "newSection";
				break;
		}
		return {
			countBy:  oLnNumType.CountBy,
			distance: private_MM2Twips(oLnNumType.Distance),
			start:    oLnNumType.Start,
			restart:  sRestartType
		}
	};
	WriterToJSON.prototype.SerPageBorders = function(oPageBorders)
	{
		if (!oPageBorders)
			return oPageBorders;

		var sDisplayType = undefined;
		switch(oPageBorders.Display)
		{
			case section_borders_DisplayAllPages:
				sDisplayType = "allPages";
				break;
			case section_borders_DisplayFirstPage:
				sDisplayType = "firstPage";
				break;
			case section_borders_DisplayNotFirstPage:
				sDisplayType = "notFirstPage";
				break;
		}
		return {
			bottom:     this.SerBorder(oPageBorders.Bottom),
			left:       this.SerBorder(oPageBorders.Left),
			right:      this.SerBorder(oPageBorders.Right),
			top:        this.SerBorder(oPageBorders.Top),
			display:    sDisplayType,
			offsetFrom: private_MM2Twips(oPageBorders.OffsetFrom),
			zOrder:     oPageBorders.ZOrder === section_borders_ZOrderFront ? "front" : "back"
		}
	};
	WriterToJSON.prototype.SerPageMargins = function(oPageMargins)
	{
		if (!oPageMargins)
			return oPageMargins;

		return {
			bottom: private_MM2Twips(oPageMargins.Bottom),
			footer: private_MM2Twips(oPageMargins.Footer),
			gutter: private_MM2Twips(oPageMargins.Gutter),
			header: private_MM2Twips(oPageMargins.Header),
			left: private_MM2Twips(oPageMargins.Left),
			right: private_MM2Twips(oPageMargins.Right),
			top: private_MM2Twips(oPageMargins.Top)
		}
	};
	WriterToJSON.prototype.SerPageSize = function(oPageSize)
	{
		if (!oPageSize)
			return oPageSize;

		var sOrientType = undefined;
		switch(oPageSize.Orient)
		{
			case Asc.c_oAscPageOrientation.PagePortrait:
				sOrientType = "portrait";
				break;
			case Asc.c_oAscPageOrientation.PageLandscape:
				sOrientType = "landscape";
				break;
		}
		return {
			h:      private_MM2Twips(oPageSize.H),
			orinet: sOrientType,
			w:      private_MM2Twips(oPageSize.W)
		}
	};
	WriterToJSON.prototype.SerHyperlink = function(oHyperlink, aComplexFieldsToSave)
	{
		var oLinkObject = {
			anchor:  oHyperlink.Anchor  ? oHyperlink.Anchor  : undefined,
			tooltip: oHyperlink.ToolTip ? oHyperlink.ToolTip : undefined,
			value:   oHyperlink.Value   ? oHyperlink.Value   : undefined,
			content: [],
			type:    "hyperlink"
		};

		if (!aComplexFieldsToSave)
			aComplexFieldsToSave = this.GetAllParaComplexFields(oHyperlink);
		
		var oTempElm = null;
		for (var nElm = 0; nElm < oHyperlink.Content.length; nElm++)
		{
			oTempElm = oHyperlink.Content[nElm];
			
			if (oTempElm instanceof AscCommonWord.ParaRun)
				oLinkObject.content.push(this.SerParaRun(oTempElm, aComplexFieldsToSave));
			else if (oTempElm instanceof AscCommonWord.CInlineLevelSdt)
				oLinkObject.content.push(this.SerInlineLvlSdt(oTempElm, aComplexFieldsToSave));
					
		}

		return oLinkObject;
	};
	WriterToJSON.prototype.SerParaRun = function(oRun, aComplexFieldsToSave)
	{
		var sReviewType = undefined;
		switch (oRun.ReviewType)
		{
			case reviewtype_Common:
				sReviewType = "common";
				break;
			case reviewtype_Remove:
				sReviewType = "remove";
				break;
			case reviewtype_Add:
				sReviewType = "add";
				break;
		}

		var oRunObject = {
			rPr:        this.SerTextPr(oRun.Pr),
			content:    [],
			footnotes:  [],
			endnotes:   [],
			reviewInfo: this.SerReviewInfo(oRun.ReviewInfo),
			reviewType: sReviewType,
			type:       "run"
		}
		
		if (oRun.IsMathRun())
			oRunObject.type = "mathRun";

		if (!aComplexFieldsToSave)
		{
			aComplexFieldsToSave = this.GetComplexFieldToSave(oRun);
		}
			
		if (oRun.Type === para_Math_Run)
			oRunObject.mathPr = oRun.MathPrp;

		function SerPageNum(oPageNum)
		{
			if (!oPageNum)
				return oPageNum;

			return {
				type: "pgNum"
			}
		}
		function SerPageCount(oPageCount)
		{
			if (!oPageCount)
				return [];

			return ToComplexField(oPageCount);
		}
		function SerCompFieldContent(oContent)
		{
			var aResult = [];
			for (var CurPos = 0; CurPos < oContent.length; CurPos++)
			{
				var oItem = oContent[CurPos];

				if (oItem.Type !== para_FieldChar)
					continue;

				var oCompField  = oItem.GetComplexField();
				var sInstuction = oCompField.InstructionLine;
				 
				if (oItem.IsBegin())
				{
					aResult.push({
						type:        "fldChar",
						fldCharType: "begin"
					});
					aResult.push({
						type:  "instrText",
						instr: sInstuction
					});
				}
				else if (oItem.IsSeparate())
				{
					aResult.push({
						type:        "fldChar",
						fldCharType: "separate"
					});
				}
				else if (oItem.IsEnd())
				{
					aResult.push({
						type:        "fldChar",
						fldCharType: "end"
					});
				}
			}
			
			return aResult;
		}
		function ToComplexField(oElement)
		{
			var arrComplexFieldRuns = [];
			var oFldCharBegin = {
				type:        "fldChar",
				fldCharType: "begin"
			}
			var oFldCharSep   = {
				type:        "fldChar",
				fldCharType: "separate"
			}
			var oFldCharEnd   = {
				type:        "fldChar",
				fldCharType: "end"
			}
			var oInstrText    = {
				type: "instrText"
			}

			var sResultOfField = "";
			arrComplexFieldRuns.push(oFldCharBegin);
			switch (oElement.Type)
			{
				case para_PageCount:
					oInstrText.instr = "PAGE";
					sResultOfField      = oElement.String;
			}
			arrComplexFieldRuns.push(oInstrText);
			arrComplexFieldRuns.push(oFldCharSep);
			arrComplexFieldRuns.push(sResultOfField);
			arrComplexFieldRuns.push(oFldCharEnd);

			return arrComplexFieldRuns;
		}
		function SerParaNewLine(oParaNewLine)
		{
			if (!oParaNewLine)
				return oParaNewLine;
				
			var sBreakType = "";
			if (oParaNewLine.IsLineBreak())
				sBreakType = "textWrapping";
			else if (oParaNewLine.IsPageBreak())
				sBreakType = "page";
			else if (oParaNewLine.IsColumnBreak())
				sBreakType = "column";

			return {
				type: "break",
				breakType: sBreakType
			}
		}
		var ContentLen        = oRun.Content.length;
		var sTempRunText      = '';
		var allowAddCompField = false;

		for (var CurPos = 0; CurPos < ContentLen; CurPos++)
		{
			var Item     = oRun.Content[CurPos];
			var ItemType = Item.Type;

			switch (ItemType)
			{
				case para_PageNum:
					sTempRunText !== "" && oRunObject.content.push(sTempRunText);
					oRunObject.content.push(SerPageNum(Item));
					sTempRunText = '';
					break;
				case para_PageCount:
					sTempRunText !== "" && oRunObject.content.push(sTempRunText);
					oRunObject.content = oRunObject.content.concat(SerPageCount(Item));
					sTempRunText = '';
					break;
				case para_Drawing:
				{
					sTempRunText !== "" && oRunObject.content.push(sTempRunText);
					oRunObject.content.push(this.SerParaDrawing(Item));
					sTempRunText = '';
					break;
				}
				case para_End:
				{
					oRunObject.type = "endRun";
					break;
				}
				case para_Text:
				{
					sTempRunText += String.fromCharCode(Item.Value);
					break;
				}
				case para_Math_BreakOperator:
				case para_Math_Placeholder:
				case para_Math_Text:
					sTempRunText !== "" && oRunObject.content.push(sTempRunText);
					oRunObject.content.push(this.SerMathText(Item));
					sTempRunText = '';
					break;
				case para_NewLine:
					sTempRunText !== "" && oRunObject.content.push(sTempRunText);
					oRunObject.content.push(SerParaNewLine(Item));
					sTempRunText = '';
					break;
				case para_Space:
				{
					sTempRunText += " ";
					break;
				}
				case para_Tab:
					sTempRunText !== "" && oRunObject.content.push(sTempRunText);
					oRunObject.content.push({
						type: "tab"
					});
					sTempRunText = '';
					break;
				case para_FieldChar:
					var oComplexField    = Item.GetComplexField();
					var oCurStartPos     = oComplexField.GetStartDocumentPosition();
					var oCurEndPos       = oComplexField.GetEndDocumentPosition();
					var oTempStartPos    = null;
					var oTempEndPos      = null;
					for (var nCompField = 0; nCompField < aComplexFieldsToSave.length; nCompField++)
					{
						oTempStartPos = aComplexFieldsToSave[nCompField].GetStartDocumentPosition();
						oTempEndPos   = aComplexFieldsToSave[nCompField].GetEndDocumentPosition();

						if (private_checkRelativePos(oCurStartPos, oTempStartPos) === 0 && private_checkRelativePos(oCurEndPos, oTempEndPos) === 0)
						{
							allowAddCompField = true;
							oRunObject.content = oRunObject.content.concat(SerCompFieldContent(oRun.Content));
							break;
						}
					}
					break;
				case para_RevisionMove:
					oRunObject.content.push(this.SerRevisionMove(Item));
					break;
				case para_FootnoteReference:
					oRunObject.content.push(this.SerParaFootEndNoteRef(Item));
					oRunObject.footnotes.push(this.SerFootEndnote(Item.Footnote));
					break;
				case para_FootnoteRef:
					oRunObject.content.push(this.SerParaFootEndNoteRef(Item));
					break;
				case para_EndnoteReference:
					oRunObject.content.push(this.SerParaFootEndNoteRef(Item));
					oRunObject.endnotes.push(this.SerFootEndnote(Item.Footnote));
					break;
				case para_EndnoteRef:
					oRunObject.content.push(this.SerParaFootEndNoteRef(Item));
					break;
			}

			if (allowAddCompField)
				break;
		}
		if (ContentLen !== 0)
			sTempRunText !== "" && oRunObject.content.push(sTempRunText);

		return oRunObject;
	};
	WriterToJSON.prototype.SerMathText = function(oMathText)
	{
		return {
			value: oMathText.value,
			type: "mathTxt"
		}
	};
	WriterToJSON.prototype.SerParaFootEndNoteRef = function(oFootnoteRef)
	{
		var sType = undefined;
		switch (oFootnoteRef.Type)
		{
			// Ссылка на сноску
			case para_FootnoteReference:
				sType = "footnoteRef";
				break;
			// Номер сноски (должен быть только внутри сноски)
			case para_FootnoteRef:
				sType = "footnoteNum";
				break;
			// Ссылка на сноску
			case para_EndnoteReference:
				sType = "endnoteRef";
				break;
			// Номер сноски (должен быть только внутри сноски)
			case para_EndnoteRef:
				sType = "endnoteNum"
		}

		return {
			customMark: oFootnoteRef.CustomMark,
			footnote:   oFootnoteRef.Footnote.Id,
			numFormat:  oFootnoteRef.NumFormat,
			number:     oFootnoteRef.Number,
			type:       sType
		}
	};
	WriterToJSON.prototype.SerReviewInfo = function(oReviewInfo)
	{
		if (!oReviewInfo)
			return oReviewInfo;

		// move type
		var sMoveType = undefined;
		switch (oReviewInfo.MoveType)
		{
			case Asc.c_oAscRevisionsMove.NoMove:
				sMoveType = "noMove";
				break;
			case Asc.c_oAscRevisionsMove.MoveTo:
				sMoveType = "moveTo";
				break;
			case Asc.c_oAscRevisionsMove.MoveFrom:
				sMoveType = "moveFrom";
				break;
		}

		// prev type
		var sPrevType = -1;
		switch (oReviewInfo.PrevType)
		{
			case Asc.c_oAscRevisionsMove.NoMove:
				sPrevType = "noMove";
				break;
			case Asc.c_oAscRevisionsMove.MoveTo:
				sPrevType = "moveTo";
				break;
			case Asc.c_oAscRevisionsMove.MoveFrom:
				sPrevType = "moveFrom";
				break;
		}

		return {
			userId:   oReviewInfo.UserId,
			autor:    oReviewInfo.UserName,
			date:     oReviewInfo.DateTime,
			moveType: sMoveType,
			prevType: sPrevType,
			prevInfo: this.SerReviewInfo(oReviewInfo.PrevInfo)
		}
	};
	WriterToJSON.prototype.SerParaDrawing = function(oDrawing, aComplexFieldsToSave)
	{
		return {
			docPr:          this.SerCNvPr(oDrawing.docPr),
			effectExtent:   this.SerEffectExtent(oDrawing.EffectExtent),
			extent:         this.SerExtent(oDrawing.Extent),
			graphic:        this.SerGraphicObject(oDrawing.GraphicObj, aComplexFieldsToSave),
			positionH:      this.SerPositionH(oDrawing.PositionH),
			positionV:      this.SerPositionV(oDrawing.PositionV),
			simplePos:      this.SerSimplePos(oDrawing.SimplePos),
			distB:          oDrawing.Distance ? private_MM2EMU(oDrawing.Distance.B) : oDrawing.Distance,
			distL:          oDrawing.Distance ? private_MM2EMU(oDrawing.Distance.L) : oDrawing.Distance,
			distR:          oDrawing.Distance ? private_MM2EMU(oDrawing.Distance.R) : oDrawing.Distance,
			distT:          oDrawing.Distance ? private_MM2EMU(oDrawing.Distance.T) : oDrawing.Distance,
			allowOverlap:   oDrawing.AllowOverlap,
			behindDoc:      oDrawing.behindDoc,
			hidden:         oDrawing.Hidden,
			layoutInCell:   oDrawing.LayoutInCell,
			locked:         oDrawing.Locked,
			relativeHeight: oDrawing.RelativeHeight,
			wrapType:       this.GetWrapStrType(oDrawing.wrappingType),
			drawingType:    oDrawing.DrawingType === drawing_Inline ? "inline" : "anchor",
			type:           "paraDrawing"
		};
	};
	WriterToJSON.prototype.GetAllParaComplexFields = function(oPara)
	{
		var arrComplexFields = [];
		var arrTemp          = [];
		var nEndPos          = oPara.Content.length;

		for (var nIndex = 0; nIndex < nEndPos; nIndex++)
		{
			if (oPara.Content[nIndex].GetCurrentComplexFields)
			{
				oPara.Content[nIndex].GetCurrentComplexFields(arrTemp);
				arrComplexFields = arrComplexFields.concat(arrTemp);
				arrTemp          = [];
			}
		}

		return arrComplexFields;
	};
	/**
	 * Get all complex fields to save from content (Takes into account the positions of the beginning and end of fields)
	 * @param  {Array} arrContent     - array of document content 
	 * @param  {Array} minStartDocPos - the minimum allowable position not exceeding which we will save the field
	 * @param  {Array} maxStartDocPos - the maximum allowable position not exceeding which we will save the field
	 * @param  {bool} bAll            - get all complex fields in content
	 * @return {Array}                - returns array with complex fields to save
	 */
	WriterToJSON.prototype.GetComplexFieldsToSave = function(arrContent, minStartDocPos, maxStartDocPos, bAll)
	{
		var oMinStartPos          = minStartDocPos ? minStartDocPos : (arrContent.length !== 0 ? arrContent[0].GetDocumentPositionFromObject() : null);
		var oMaxStartPos          = maxStartDocPos ? maxStartDocPos : (arrContent.length !== 0 ? arrContent[arrContent.length - 1].GetDocumentPositionFromObject() : null);
		if (oMinStartPos[0].Position === -1 || oMaxStartPos[0].Position === -1)
			bAll = true;

		var arrCompexFieldsToSave = [];
		var arrElmContent;

		for (var nElm = 0; nElm < arrContent.length; nElm++)
		{
			var oElm                  = arrContent[nElm];
			var oFieldStartPos        = null;
			var oFieldEndPos          = null;
			var arrTemp               = [];

			if (oElm instanceof AscCommonWord.Paragraph)
			{
				arrTemp = this.GetAllParaComplexFields(oElm);
				if (!bAll)
				{
					for (var nField = 0; nField < arrTemp.length; nField++)
					{
						oFieldStartPos = arrTemp[nField].GetStartDocumentPosition();
						oFieldEndPos   = arrTemp[nField].GetEndDocumentPosition();

						if (private_checkRelativePos(oFieldStartPos, oMinStartPos) === 1 || private_checkRelativePos(oFieldEndPos, oMaxStartPos) === -1)
						{
							arrTemp.splice(nField, 1);
							nField--;
						}
					}
				}
				
				arrCompexFieldsToSave = arrCompexFieldsToSave.concat(arrTemp);
			}
			if (oElm instanceof AscCommonWord.CTable)
			{
				arrElmContent = null;
				for (var nRow = 0; nRow < oElm.Content.length; nRow++)
				{
					for (var nCell = 0; nCell < oElm.Content[nRow].Content.length; nCell++)
					{
						arrElmContent         = oElm.Content[nRow].Content[nCell].Content.Content;
						arrCompexFieldsToSave = arrCompexFieldsToSave.concat(this.GetComplexFieldsToSave(arrElmContent, minStartDocPos, maxStartDocPos));
					}
				}
			}
			if (oElm instanceof AscCommonWord.CBlockLevelSdt)
			{
				arrElmContent     = oElm.Content.Content;
				arrCompexFieldsToSave = arrCompexFieldsToSave.concat(this.GetComplexFieldsToSave(arrElmContent, minStartDocPos, maxStartDocPos));
			}
		}

		return arrCompexFieldsToSave;
	};
	/**
	 * Get all comments to save from content (Takes into account the positions of the beginning and end of fields)
	 * @param {Array} arrContent - array of document content 
	 * @param {object} oMap
	 * @return {object} - returns map with comments to save
	 */
	WriterToJSON.prototype.GetMapCommentsInfo = function(arrContent, oMap)
	{
		if (!oMap)
			oMap = {};

		var arrElmContent;

		for (var nElm = 0; nElm < arrContent.length; nElm++)
		{
			var oElm = arrContent[nElm];

			if (oElm instanceof AscCommonWord.Paragraph)
			{
				var aParaComments = oElm.GetAllComments();
				for (var nComment = 0; nComment < aParaComments.length; nComment++)
				{
					if (!oMap[aParaComments[nComment].Comment.CommentId])
						oMap[aParaComments[nComment].Comment.CommentId] = {};

					if (aParaComments[nComment].Comment.Start)
						oMap[aParaComments[nComment].Comment.CommentId].Start = true;
					else
						oMap[aParaComments[nComment].Comment.CommentId].End = true;
				}
			}
			if (oElm instanceof AscCommonWord.CTable)
			{
				arrElmContent = null;
				for (var nRow = 0; nRow < oElm.Content.length; nRow++)
				{
					for (var nCell = 0; nCell < oElm.Content[nRow].Content.length; nCell++)
					{
						arrElmContent         = oElm.Content[nRow].Content[nCell].Content.Content;
						this.GetMapCommentsInfo(arrElmContent, oMap);
					}
				}
			}
			if (oElm instanceof AscCommonWord.CBlockLevelSdt)
			{
				arrElmContent = oElm.Content.Content;
				this.GetMapCommentsInfo(arrElmContent, oMap);
			}
		}

		return oMap;
	};
	/**
	 * Get all bookmarks to save from content (Takes into account the positions of the beginning and end of fields)
	 * @param {Array} arrContent - array of document content 
	 * @param {object} oMap
	 * @return {object} - returns map with boomarks to save
	 */
	WriterToJSON.prototype.GetMapBookmarksInfo = function(arrContent, oMap)
	{
		if (!oMap)
			oMap = {};

		var arrElmContent;

		if (!AscCommonWord.CParagraphBookmark)
			return oMap;

		for (var nElm = 0; nElm < arrContent.length; nElm++)
		{
			var oElm = arrContent[nElm];

			if (oElm instanceof AscCommonWord.Paragraph)
			{
				for (var nItem = 0; nItem < oElm.Content.length; nItem++)
				{
					if (!(oElm.Content[nItem] instanceof AscCommonWord.CParagraphBookmark))
						continue;

					if (!oMap[oElm.Content[nItem].BookmarkId])
						oMap[oElm.Content[nItem].BookmarkId] = {};

					if (oElm.Content[nItem].Start)
						oMap[oElm.Content[nItem].BookmarkId].Start = true;
					else
						oMap[oElm.Content[nItem].BookmarkId].End = true;
				}
			}
			if (oElm instanceof AscCommonWord.CTable)
			{
				arrElmContent = null;
				for (var nRow = 0; nRow < oElm.Content.length; nRow++)
				{
					for (var nCell = 0; nCell < oElm.Content[nRow].Content.length; nCell++)
					{
						arrElmContent = oElm.Content[nRow].Content[nCell].Content.Content;
						this.GetMapBookmarksInfo(arrElmContent, oMap);
					}
				}
			}
			if (oElm instanceof AscCommonWord.CBlockLevelSdt)
			{
				arrElmContent = oElm.Content.Content;
				this.GetMapBookmarksInfo(arrElmContent, oMap);
			}
		}

		return oMap;
	};
	WriterToJSON.prototype.GetComplexFieldToSave = function(oRun)
	{
		var aFieldsToSave = []; 
		var arrRunContent = oRun.Content;
		for (var CurPos = 0; CurPos < arrRunContent.length; CurPos++)
		{
			var oItem = arrRunContent[CurPos];

			if (oItem.Type !== para_FieldChar)
				continue;

			var oCompField     = oItem.GetComplexField();
			var oFieldStartPos = oCompField.GetStartDocumentPosition();
			var oFieldEndPos   = oCompField.GetEndDocumentPosition();
			
			var RunStartPos = oRun.GetDocumentPositionFromObject();
			var RunEndPos   = oRun.GetDocumentPositionFromObject();

			RunStartPos.push({Class: oRun, Position: 0});
			RunEndPos.push({Class: oRun, Position: oRun.Content.length});

			if (private_checkRelativePos(oFieldStartPos, RunStartPos) === 1 || private_checkRelativePos(oFieldEndPos, RunEndPos) === -1)
				break;
			else
			{
				aFieldsToSave.push(oCompField);
				break;
			}
		}

		return aFieldsToSave;
	};
	WriterToJSON.prototype.SerNumbering = function(oNum)
	{
		if (!oNum)
			return oNum;

		var arrNumLvls     = [];
		var abtrNumb       = oNum.Numbering.AbstractNum[oNum.AbstractNumId];
		var arrLvlOverride = [];

		// fill arrNumLvls
		for (var nLvl = 0; nLvl < abtrNumb.Lvl.length; nLvl++)
			arrNumLvls.push(this.SerNumLvl(abtrNumb.Lvl[nLvl], nLvl));

		// fill arrLvlOverride
		for (var nOvrd = 0; nOvrd < oNum.LvlOverride.length; nOvrd++)
		{
			arrLvlOverride.push({
				lvl:           this.SerNumLvl(oNum.LvlOverride[nOvrd], oNum.LvlOverride[nOvrd].Lvl),
				startOverride: oNum.LvlOverride[nOvrd].StartOverride,
				ilvl:          oNum.LvlOverride[nOvrd].Lvl
			});
		}
		
		return {
			abstractNum: {
				lvl:           arrNumLvls,
				numStyleLink:  abtrNumb.NumStyleLink,
				styleLink:     abtrNumb.StyleLink,
				abstractNumId: abtrNumb.Id
			},
			num: {
				abstractNumId: oNum.AbstractNumId,
				lvlOverride:   arrLvlOverride,
				numId:         oNum.Id
			},
			type: "numbering"
		}
	};
	WriterToJSON.prototype.SerNumLvl = function(oLvl, nLvl)
	{
		// align
		var sJc = undefined;
		switch (oLvl.Jc)
		{
			case AscCommon.align_Right:
				sJc = "end";
				break;
			case AscCommon.align_Left:
				sJc = "start";
				break;
			case AscCommon.align_Center:
				sJc = "center";
				break;
			case AscCommon.align_Justify:
				sJc = "both";
				break;
			case AscCommon.align_Distributed:
				sJc = "distribute";
				break;
		}

		var oStyles = private_GetStyles();
		// style
		var oStyle   = oLvl.PStyle ? oStyles.Get(oLvl.PStyle) : undefined;

		// suff
		var sSuffType = undefined;
		switch (oLvl.Suff)
		{
			case Asc.c_oAscNumberingSuff.Tab:
				sSuffType = "tab";
				break;
			case Asc.c_oAscNumberingSuff.Space:
				sSuffType = "space";
				break;
			case Asc.c_oAscNumberingSuff.None:
				sSuffType = "nothing";
				break;
		}

		// format type
		var sFormatType = To_XML_c_oAscNumberingFormat(oLvl.Format);

		// lvl text
		var oLvlText = {
			val: oLvl.LvlText.length === 2 ? oLvl.LvlText[1].Value : oLvl.LvlText[0].Value
		}
		if (oLvl.LvlText.length)
			oLvlText.numValue = oLvl.LvlText[0].Value;

		return {
			isLgl: oLvl.IsLgl,

			legacy: oLvl.Legacy ? {
				legacy:       oLvl.Legacy.Legacy,
				legacyIndent: oLvl.Legacy.Indent,
				legacySpace:  oLvl.Legacy.Space
			} : oLvl.Legacy,

			lvlJc:   sJc,
			lvlText: oLvlText,
			numFmt:  {
				val: sFormatType
			},
			pPr:     this.SerParaPr(oLvl.ParaPr),
			pStyle:  this.SerStyle(oStyle),
			rPr:     this.SerTextPr(oLvl.TextPr),
			restart: oLvl.Restart,
			start:   oLvl.Start,
			suff:    sSuffType,
			ilvl:    nLvl
		}
	};
	WriterToJSON.prototype.SerTabs = function(oTabs)
	{
		if (!oTabs)
			return oTabs;
			
		var oTabsObj = {
			tabs: []
		};

		for (var nTab = 0; nTab < oTabs.Tabs.length; nTab++)
		{
			oTabsObj.tabs.push({
				val:    oTabs.Tabs[nTab].Value,
				pos:    private_MM2Twips(oTabs.Tabs[nTab].Pos),
				leader: oTabs.Tabs[nTab].Leader
			});
		}

		return oTabsObj;
	};
	WriterToJSON.prototype.SerSerTx = function(oTx)
	{
		if (!oTx)
			return oTx;

		return {
			strRef: this.SerStrRef(oTx.strRef),
			v:      oTx.val
		}
	};
	WriterToJSON.prototype.SerChartTx = function(oTx)
	{
		if (!oTx)
			return oTx;

		return {
			strRef: this.SerStrRef(oTx.strRef),
			rich:   this.SerTxPr(oTx.rich)
		}
	};
	WriterToJSON.prototype.SerScaling = function(oScaling)
	{
		if (!oScaling)
			return oScaling;
		
		var sOrientType = oScaling.orientation === AscFormat.ORIENTATION_MAX_MIN ? "maxMin" : "minMax"; 
		return {
			logBase:     oScaling.logBase,
			max:         oScaling.max,
			min:         oScaling.min,
			orientation: sOrientType
		}
	};
	WriterToJSON.prototype.SerNumFmt = function(oNumFmt)
	{
		if (!oNumFmt)
			return oNumFmt;
		
		return {
			formatCode:   oNumFmt.formatCode,
			sourceLinked: oNumFmt.sourceLinked
		}
	};
	WriterToJSON.prototype.SerDispUnits = function(oDispUnits)
	{
		if (!oDispUnits)
			return oDispUnits;
		
		var sBuiltInUnit = undefined;
		switch(oDispUnits.builtInUnit)
		{
			case Asc.c_oAscValAxUnits.none:
				sBuiltInUnit = "none";
				break;
			case Asc.c_oAscValAxUnits.BILLIONS:
				sBuiltInUnit = "billions";
				break;
			case Asc.c_oAscValAxUnits.HUNDRED_MILLIONS:
				sBuiltInUnit = "hundredMillions";
				break;
			case Asc.c_oAscValAxUnits.HUNDREDS:
				sBuiltInUnit = "hundreds";
				break;
			case Asc.c_oAscValAxUnits.HUNDRED_THOUSANDS:
				sBuiltInUnit = "hundredThousands";
				break;
			case Asc.c_oAscValAxUnits.MILLIONS:
				sBuiltInUnit = "millions";
				break;
			case Asc.c_oAscValAxUnits.TEN_MILLIONS:
				sBuiltInUnit = "tenMillions";
				break;
			case Asc.c_oAscValAxUnits.TEN_THOUSANDS:
				sBuiltInUnit = "tenThousands";
				break;
			case Asc.c_oAscValAxUnits.TRILLIONS:
				sBuiltInUnit = "trillions";
				break;
			case Asc.c_oAscValAxUnits.CUSTOM:
				sBuiltInUnit = "custom";
				break;
			case Asc.c_oAscValAxUnits.THOUSANDS:
				sBuiltInUnit = "thousands";
				break;
		}
		return {
			builtInUnit:  sBuiltInUnit,
			custUnit:     oDispUnits.custUnit,
			dispUnitsLbl: this.SerDlbl(oDispUnits.dispUnitsLbl)
		}
	};
	WriterToJSON.prototype.SerValAx = function(oValAx)
	{
		if (!oValAx)
			return oValAx;
		
		var sCrossBetweenType = oValAx.crossBetween === AscFormat.CROSS_BETWEEN_BETWEEN ? "between" : "micCat";
		
		return {
			axId:           oValAx.axId,
			axPos:          this.GetAxPosStrType(oValAx.axPos),
			crossAx:        oValAx.crossAx.axId,
			crossBetween:   sCrossBetweenType,
			crosses:        this.GetCrossesStrType(oValAx.crosses),
			crossesAt:      oValAx.crossesAt,
			delete:         oValAx.bDelete,
			dispUnits:      this.SerDispUnits(oValAx.dispUnits),
			extLst:         oValAx.extLst, /// ???
			majorGridlines: this.SerSpPr(oValAx.majorGridlines),
			majorTickMark:  this.GetTickMarkStrType(oValAx.majorTickMark),
			majorUnit:      oValAx.majorUnit,
			minorGridlines: this.SerSpPr(oValAx.minorGridlines),
			minorTickMark:  this.GetTickMarkStrType(oValAx.minorTickMark),
			minorUnit:      oValAx.minorUnit,
			numFmt:         this.SerNumFmt(oValAx.numFmt),
			scaling:        this.SerScaling(oValAx.scaling),
			spPr:           this.SerSpPr(oValAx.spPr),
			tickLblPos:     this.GetTickLabelStrPos(oValAx.tickLblPos),
			title:          this.SerTitle(oValAx.title),
			txPr:           this.SerTxPr(oValAx.txPr),
			type:           "valAx"
		}
	};
	WriterToJSON.prototype.GetAxPosStrType = function(nType)
	{
		var sAxPos = undefined;
		switch (nType)
		{
			case AscFormat.AX_POS_B:
				sAxPos = "b";
				break;
			case AscFormat.AX_POS_L:
				sAxPos = "l";
				break;
			case AscFormat.AX_POS_R:
				sAxPos = "r";
				break;
			case AscFormat.AX_POS_T:
				sAxPos = "t";
				break;
		}

		return sAxPos;
	};
	WriterToJSON.prototype.GetTickLabelStrPos = function(nType)
	{
		var sTickLblPos = undefined;
		switch (nType)
		{
			case Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_HIGH:
				sTickLblPos = "high";
				break;
			case Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_LOW:
				sTickLblPos = "low";
				break;
			case Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO:
				sTickLblPos = "nextTo";
				break;
			case Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE:
				sTickLblPos = "none";
				break;
		}

		return sTickLblPos;
	};
    WriterToJSON.prototype.SerAxis = function(oAxis)
	{
		if (oAxis instanceof AscFormat.CCatAx)
			return this.SerCatAx(oAxis);
		if (oAxis instanceof AscFormat.CValAx)
			return this.SerValAx(oAxis);
		if (oAxis instanceof AscFormat.CDateAx)
			return this.SerDateAx(oAxis);
		if (oAxis instanceof AscFormat.CSerAx)
			return this.SerSerAx(oAxis);
	};
	WriterToJSON.prototype.SerPlotArea = function(oPlotArea)
	{
		if (!oPlotArea)
			return oPlotArea;

		var arrAxId = [];

		for (var nAx = 0; nAx < oPlotArea.axId.length; nAx++)
			arrAxId.push(this.SerAxis(oPlotArea.axId[nAx]))

		return {
			axId:     arrAxId,
			charts:   this.SerCharts(oPlotArea.charts),
			catAx:    oPlotArea.catAx ? oPlotArea.catAx.axId : oPlotArea.catAx,
			dateAx:   oPlotArea.dateAx ? oPlotArea.dateAx.axId : oPlotArea.dateAx,
			valAx:    oPlotArea.valAx ? oPlotArea.valAx.axId : oPlotArea.valAx,
			serAx:    oPlotArea.serAx ? oPlotArea.serAx.axId : oPlotArea.serAx,
			dTable:   this.SerDataTable(oPlotArea.dTable),
			layout:   this.SerLayout(oPlotArea.layout),
			spPr:     this.SerSpPr(oPlotArea.spPr)
		}
	};
	WriterToJSON.prototype.SerDataTable = function(oData)
	{
		if (!oData)
			return oData;

		return {
			showHorzBorder: oData.showHorzBorder,
			showKeys:       oData.showKeys,
			showOutline:    oData.showOutline,
			showVertBorder: oData.showVertBorder,
			spPr:           this.SerSpPr(oData.spPr),
			txPr:           this.SerTxPr(oData.txPr)
		};
	};
	WriterToJSON.prototype.SerSerAx = function(oSerAx)
	{
		if (!oSerAx)
			return oSerAx;

		return {
			axId:           oSerAx.axId,
			axPos:          this.GetAxPosStrType(oSerAx.axPos),
			crossAx:        oSerAx.crossAx.axId,
			crosses:        this.GetCrossesStrType(oSerAx.crosses),
			crossesAt:      oSerAx.crossesAt,
			delete:         oSerAx.bDelete,
			extLst:         oSerAx.extLst, /// ?
			majorGridlines: this.SerSpPr(oSerAx.majorGridlines),
			majorTickMark:  this.GetTickMarkStrType(oSerAx.majorTickMark),
			minorGridlines: this.SerSpPr(oSerAx.minorGridlines),
			minorTickMark:  this.GetTickMarkStrType(oSerAx.minorTickMark),
			numFmt:         this.SerNumFmt(oSerAx.numFmt),
			scaling:        this.SerScaling(oSerAx.scaling),
			spPr:           this.SerSpPr(oSerAx.spPr),
			tickLblPos:     this.GetTickLabelStrPos(oSerAx.tickLblPos),
			tickLblSkip:    oSerAx.tickLblSkip,
			tickMarkSkip:   oSerAx.tickMarkSkip,
			title:          this.SerTitle(oSerAx.title),
			txPr:           this.SerTxPr(oSerAx.txPr),
			type:           "serAx"
		}
	};
	WriterToJSON.prototype.SerCatAx = function(oCatAx)
	{
		if (!oCatAx)
			return oCatAx;

		var sLblAlgn = undefined;
		switch(oCatAx.lblAlgn)
		{
			case AscFormat.LBL_ALG_CTR:
				sLblAlgn = "ctr";
				break;
			case AscFormat.LBL_ALG_L:
				sLblAlgn = "l";
				break;
			case AscFormat.LBL_ALG_R:
				sLblAlgn = "r";
				break;
		}

		return {
			auto:           oCatAx.auto,
			axId:           oCatAx.axId,
			axPos:          this.GetAxPosStrType(oCatAx.axPos),
			crossAx:        oCatAx.crossAx.axId,
			crosses:        this.GetCrossesStrType(oCatAx.crosses),
			crossesAt:      oCatAx.crossesAt,
			delete:         oCatAx.bDelete,
			extLst:         oCatAx.extLst, /// ?
			lblAlgn:        sLblAlgn,
			lblOffset:      oCatAx.lblOffset,
			majorGridlines: this.SerSpPr(oCatAx.majorGridlines),
			majorTickMark:  this.GetTickMarkStrType(oCatAx.majorTickMark),
			minorGridlines: this.SerSpPr(oCatAx.minorGridlines),
			minorTickMark:  this.GetTickMarkStrType(oCatAx.minorTickMark),
			noMultiLvlLbl:  oCatAx.noMultiLvlLbl,
			numFmt:         this.SerNumFmt(oCatAx.numFmt),
			scaling:        this.SerScaling(oCatAx.scaling),
			spPr:           this.SerSpPr(oCatAx.spPr),
			tickLblPos:     this.GetTickLabelStrPos(oCatAx.tickLblPos),
			tickLblSkip:    oCatAx.tickLblSkip,
			tickMarkSkip:   oCatAx.tickMarkSkip,
			title:          this.SerTitle(oCatAx.title),
			txPr:           this.SerTxPr(oCatAx.txPr),
			type:           "catAx"
		}
	};
	WriterToJSON.prototype.SerDateAx = function(oDateAx)
	{
		if (!oDateAx)
			return oDateAx;

		return {
			auto:           oDateAx.auto,
			axId:           oDateAx.axId,
			axPos:          this.GetAxPosStrType(oDateAx.axPos),
			baseTimeUnit:   this.GetTimeUnitStrType(oDateAx.baseTimeUnit),
			crossAx:        oDateAx.crossAx ? oDateAx.crossAx.axId : oDateAx.crossAx,
			crosses:        this.GetCrossesStrType(oDateAx.crosses),
			crossesAt:      oDateAx.crossesAt,
			delete:         oDateAx.bDelete,
			extLst:         oDateAx.extLst,
			lblOffset:      oDateAx.lblOffset,
			majorGridlines: this.SerSpPr(oDateAx.majorGridlines),
			majorTickMark:  this.GetTickMarkStrType(oDateAx.majorTickMark),
			majorTimeUnit:  this.GetTimeUnitStrType(oDateAx.majorTimeUnit),
			majorUnit:      oDateAx.majorUnit,
			minorGridlines: this.SerSpPr(oDateAx.minorGridlines),
			minorTickMark:  this.GetTickMarkStrType(oDateAx.minorTickMark),
			minorTimeUnit:  this.GetTimeUnitStrType(oDateAx.minorTimeUnit),
			minorUnit:      oDateAx.minorUnit,
			numFmt:         this.SerNumFmt(oDateAx.numFmt),
			scaling:        this.SerScaling(oDateAx.scaling),
			spPr:           this.SerSpPr(oDateAx.spPr),
			tickLblPos:     this.GetTickLabelStrPos(oDateAx.tickLblPos),
			title:          this.SerTitle(oDateAx.title),
			txPr:           this.SerTxPr(oDateAx.txPr),
			type:           "dateAx"
		}
	};
	WriterToJSON.prototype.GetCrossesStrType = function(nType)
	{
		var sType = undefined;

		switch(nType)
		{
			case AscFormat.CROSSES_AUTO_ZERO:
				sType = "autoZero";
				break;
			case AscFormat.CROSSES_MAX:
				sType = "max";
				break;
			case AscFormat.CROSSES_MIN:
				sType = "min";
				break;
		}

		return sType;
	};
	WriterToJSON.prototype.GetTickMarkStrType = function(nType)
	{
		var sType = undefined;
		switch(nType)
		{
			case Asc.c_oAscTickMark.TICK_MARK_CROSS:
				sType = "cross";
				break;
			case Asc.c_oAscTickMark.TICK_MARK_IN:
				sType = "in";
				break;
			case Asc.c_oAscTickMark.TICK_MARK_NONE:
				sType = "none";
				break;
			case Asc.c_oAscTickMark.TICK_MARK_OUT:
				sType = "out";
				break;
		}

		return sType;
	};
	WriterToJSON.prototype.GetTimeUnitStrType = function(nType)
	{
		var sTimeUnit = undefined;
		switch(nType)
		{
			case AscFormat.TIME_UNIT_DAYS:
				sTimeUnit = "days";
				break;
			case AscFormat.TIME_UNIT_MONTHS:
				sTimeUnit = "months";
				break;
			case AscFormat.TIME_UNIT_YEARS:
				sTimeUnit = "years";
				break;	
		}
		
		return sTimeUnit;
	};
	WriterToJSON.prototype.SerLayout = function(oLayout)
	{
		if (!oLayout)
			return oLayout;
		
		return {
			h:            oLayout.h,
			hMode:        oLayout.hMode != undefined ? (oLayout.hMode === AscFormat.LAYOUT_MODE_EDGE ? "edge" : "factor") : oLayout.hMode,
			layoutTarget: oLayout.layoutTarget != undefined ? (oLayout.layoutTarget === AscFormat.LAYOUT_TARGET_INNER ? "inner" : "outer") : oLayout.layoutTarget,
			w:            oLayout.w,
			wMode:        oLayout.wMode != undefined ? (oLayout.wMode === AscFormat.LAYOUT_MODE_EDGE ? "edge" : "factor") : oLayout.wMode,
			x:            oLayout.x,
			xMode:        oLayout.xMode != undefined ? (oLayout.xMode === AscFormat.LAYOUT_MODE_EDGE ? "edge" : "factor") : oLayout.xMode,
			y:            oLayout.y,
			yMode:        oLayout.yMode != undefined ? (oLayout.yMode === AscFormat.LAYOUT_MODE_EDGE ? "edge" : "factor") : oLayout.yMode
		}
	};
	WriterToJSON.prototype.SerDlbl = function(oDlbl)
	{
		if (!oDlbl)
			return oDlbl;
		
		// TickLblPos
		var sDLblPos = undefined;
		switch (oDlbl.dLblPos)
		{
			case Asc.c_oAscChartDataLabelsPos.b:
				sDLblPos = "b";
				break;
			case Asc.c_oAscChartDataLabelsPos.bestFit:
				sDLblPos = "bestFit";
				break;
			case Asc.c_oAscChartDataLabelsPos.ctr:
				sDLblPos = "ctr";
				break;
			case Asc.c_oAscChartDataLabelsPos.inBase:
				sDLblPos = "inBase";
				break;
			case Asc.c_oAscChartDataLabelsPos.inEnd:
				sDLblPos = "inEnd";
				break;
			case Asc.c_oAscChartDataLabelsPos.l:
				sDLblPos = "l";
				break;
			case Asc.c_oAscChartDataLabelsPos.outEnd:
				sDLblPos = "outEnd";
				break;
			case Asc.c_oAscChartDataLabelsPos.r:
				sDLblPos = "r";
				break;
			case Asc.c_oAscChartDataLabelsPos.t:
				sDLblPos = "t";
				break;
		}

		return {
			delete:         oDlbl.bDelete,
			dLblPos:        sDLblPos,
			idx:            oDlbl.idx,
			layout:         this.SerLayout(oDlbl.layout),
			numFmt:         this.SerNumFmt(oDlbl.numFmt),
			separator:      oDlbl.separator,
			showBubbleSize: oDlbl.showBubbleSize,
			showCatName:    oDlbl.showCatName,
			showLegendKey:  oDlbl.showLegendKey,
			showPercent:    oDlbl.showPercent,
			showSerName:    oDlbl.showSerName,
			showVal:        oDlbl.showVal,
			spPr:           this.SerSpPr(oDlbl.spPr),
			txPr:           this.SerTxPr(oDlbl.txPr),
			tx:             this.SerChartTx(oDlbl.tx)
		}
	};
	WriterToJSON.prototype.SerDLbls = function(dLbls)
	{
		if (!dLbls)
			return dLbls;
		
		// TickLblPos
		var sDLblPos = undefined;
		switch (dLbls.dLblPos)
		{
			case Asc.c_oAscChartDataLabelsPos.b:
				sDLblPos = "b";
				break;
			case Asc.c_oAscChartDataLabelsPos.bestFit:
				sDLblPos = "bestFit";
				break;
			case Asc.c_oAscChartDataLabelsPos.ctr:
				sDLblPos = "ctr";
				break;
			case Asc.c_oAscChartDataLabelsPos.inBase:
				sDLblPos = "inBase";
				break;
			case Asc.c_oAscChartDataLabelsPos.inEnd:
				sDLblPos = "inEnd";
				break;
			case Asc.c_oAscChartDataLabelsPos.l:
				sDLblPos = "l";
				break;
			case Asc.c_oAscChartDataLabelsPos.outEnd:
				sDLblPos = "outEnd";
				break;
			case Asc.c_oAscChartDataLabelsPos.r:
				sDLblPos = "r";
				break;
			case Asc.c_oAscChartDataLabelsPos.t:
				sDLblPos = "t";
				break;
		}

		//dlbl
		var arrDlbl = [];
		for (var nDlbl = 0; nDlbl < dLbls.dLbl.length; nDlbl++)
			arrDlbl.push(this.SerDlbl(dLbls.dLbl[nDlbl]));

		return {
			delete:          dLbls.bDelete,
			dLbl:            arrDlbl,
			dLblPos:         sDLblPos,
			leaderLines:     this.SerSpPr(dLbls.leaderLines),
			numFmt:          this.SerNumFmt(dLbls.numFmt),
			separator:       dLbls.separator,
			showBubbleSize:  dLbls.showBubbleSize,
			showCatName:     dLbls.showCatName,
			showLeaderLines: dLbls.showLeaderLines,
			showLegendKey:   dLbls.showLegendKey,
			showPercent:     dLbls.showPercent,
			showSerName:     dLbls.showSerName,
			showVal:         dLbls.showVal,
			spPr:            this.SerSpPr(dLbls.spPr),
			txPr:            this.SerTxPr(dLbls.txPr)
		}
	};
	WriterToJSON.prototype.SerCharts = function(arrBarCharts)
	{
		var arrResult = [];

		for (var nChart = 0; nChart < arrBarCharts.length; nChart++)
		{
			if (arrBarCharts[nChart] instanceof AscFormat.CBarChart)
			{
				arrResult.push(this.SerBarChart(arrBarCharts[nChart]));
				continue;
			}
			if (arrBarCharts[nChart] instanceof AscFormat.CLineChart)
			{
				arrResult.push(this.SerLineChart(arrBarCharts[nChart]));
				continue;
			}
			if (arrBarCharts[nChart] instanceof AscFormat.CPieChart)
			{
				arrResult.push(this.SerPieChart(arrBarCharts[nChart]));
				continue;
			}
			if (arrBarCharts[nChart] instanceof AscFormat.CDoughnutChart)
			{
				arrResult.push(this.SerDoughnutChart(arrBarCharts[nChart]));
				continue;
			}
			if (arrBarCharts[nChart] instanceof AscFormat.CAreaChart)
			{
				arrResult.push(this.SerAreaChart(arrBarCharts[nChart]));
				continue;
			}
			if (arrBarCharts[nChart] instanceof AscFormat.CStockChart)
			{
				arrResult.push(this.SerStockChart(arrBarCharts[nChart]));
				continue;
			}
			if (arrBarCharts[nChart] instanceof AscFormat.CScatterChart)
			{
				arrResult.push(this.SerScatterChart(arrBarCharts[nChart]));
				continue;
			}
			if (arrBarCharts[nChart] instanceof AscFormat.CBubbleChart)
			{
				arrResult.push(this.SerBubbleChart(arrBarCharts[nChart]));
				continue;
			}
			if (arrBarCharts[nChart] instanceof AscFormat.CSurfaceChart)
			{
				arrResult.push(this.SerSurfaceChart(arrBarCharts[nChart]));
				continue;
			}
			if (arrBarCharts[nChart] instanceof AscFormat.CRadarChart)
			{
				arrResult.push(this.SerRadarChart(arrBarCharts[nChart]));
			}
		}

		return arrResult;
	};
	WriterToJSON.prototype.SerDoughnutChart = function(oDoughnutChart)
	{
		return {
			dLbls:         this.SerDLbls(oDoughnutChart.dLbls),
			firstSliceAng: oDoughnutChart.firstSliceAng,
			holeSize:      oDoughnutChart.holeSize,
			ser:           this.SerPieSeries(oDoughnutChart.series),
			varyColors:    oDoughnutChart.varyColors,
			type:          "doughnutChart"
		}
	};
	WriterToJSON.prototype.SerRadarChart = function(oRadarChart)
	{
		var sRadarStyle = undefined;
		switch(oRadarChart.radarStyle)
		{
			case AscFormat.RADAR_STYLE_STANDARD:
				sRadarStyle = "standard";
				break;
			case AscFormat.RADAR_STYLE_MARKER:
				sRadarStyle = "marker";
				break;
			case AscFormat.RADAR_STYLE_FILLED:
				sRadarStyle = "filled";
				break;
		}
		
		var arrAxId = [];
		for (var nAxis = 0; nAxis < oRadarChart.axId.length; nAxis++)
			arrAxId.push(oRadarChart.axId[nAxis].axId);

		return {
			axId:         arrAxId,
			dLbls:        this.SerDLbls(oRadarChart.dLbls),
			radarStyle:   sRadarStyle,
			ser:          this.SerRadarSeries(oRadarChart.series),
			varyColors:   oRadarChart.varyColors,
			type:         "radarChart"
		}
	};
	WriterToJSON.prototype.SerRadarSeries = function(arrRadarSeries)
	{
		var arrResultSeries = [];

		for (var nItem = 0; nItem < arrRadarSeries.length; nItem++)
		{
			arrResultSeries.push({
				cat:              this.SerCat(arrRadarSeries[nItem].cat),
				dLbls:            this.SerDLbls(arrRadarSeries[nItem].dLbls),
				dPt:              this.SerDataPoints(arrRadarSeries[nItem].dPt),
				idx:              arrRadarSeries[nItem].idx,
				marker:           this.SerMarker(arrRadarSeries[nItem].marker),
				order:            arrRadarSeries[nItem].order,
				spPr:             this.SerSpPr(arrRadarSeries[nItem].spPr),
				tx:               this.SerSerTx(arrRadarSeries[nItem].tx),
				val:              this.SerYVAL(arrRadarSeries[nItem].val)
			});
		}

		return arrResultSeries;
	};
	WriterToJSON.prototype.SerSurfaceChart = function(oSurfaceChart)
	{
		var arrAxId = [];
		for (var nAxis = 0; nAxis < oSurfaceChart.axId.length; nAxis++)
			arrAxId.push(oSurfaceChart.axId[nAxis].axId);

		return {
			axId:           arrAxId,
			bandFmts:       this.SerBandFmts(oSurfaceChart.bandFmts),
			ser:            this.SerSurfaceSeries(oSurfaceChart.series),
			wireframe:      oSurfaceChart.wireframe,
			type:           "surfaceChart"
		}
	};
	WriterToJSON.prototype.SerSurfaceSeries = function(arrSurfaceSeries)
	{
		var arrResultSeries = [];

		for (var nItem = 0; nItem < arrSurfaceSeries.length; nItem++)
		{
			arrResultSeries.push({
				cat:              this.SerCat(arrSurfaceSeries[nItem].cat),
				idx:              arrSurfaceSeries[nItem].idx,
				order:            arrSurfaceSeries[nItem].order,
				spPr:             this.SerSpPr(arrSurfaceSeries[nItem].spPr),
				tx:               this.SerSerTx(arrSurfaceSeries[nItem].tx),
				val:              this.SerYVAL(arrSurfaceSeries[nItem].val)
			});
		}

		return arrResultSeries;
	};
	WriterToJSON.prototype.SerBandFmts = function(arrBandFmts)
	{
		var arrResult = [];

		for (var nBand = 0; nBand < arrBandFmts.length; nBand++)
		{
			arrResult.push({
				idx:  arrBandFmts[nBand].idx,
				spPr: this.SerSpPr(arrBandFmts[nBand.spPr])
			});
		}

		return arrResult;		
	};
	WriterToJSON.prototype.SerBubbleChart = function(oBubbleChart)
	{
		var arrAxId = [];
		for (var nAxis = 0; nAxis < oBubbleChart.axId.length; nAxis++)
			arrAxId.push(oBubbleChart.axId[nAxis].axId);

		var sSizeRepresents = oBubbleChart.sizeRepresents === AscFormat.SIZE_REPRESENTS_AREA ? "area" : "w";
		
		return {
			axId:           arrAxId,
			bubble3D:       oBubbleChart.bubble3D,
			bubbleScale:    oBubbleChart.bubbleScale,
			dLbls:          this.SerDLbls(oBubbleChart.dLbls),
			ser:            this.SerBubbleSeries(oBubbleChart.series),
			showNegBubbles: oBubbleChart.showNegBubbles,
			sizeRepresents: sSizeRepresents,
			varyColors:     oBubbleChart.varyColors,
			type:           "bubbleChart"
		}
	};
	WriterToJSON.prototype.SerBubbleSeries = function(arrBubbleSeries)
	{
		var arrResultSeries = [];

		for (var nItem = 0; nItem < arrBubbleSeries.length; nItem++)
		{
			arrResultSeries.push({
				bubble3D:         arrBubbleSeries.bubble3D,
				bubbleSize:       this.SerYVAL(arrBubbleSeries[nItem].bubbleSize),
				dLbls:            this.SerDLbls(arrBubbleSeries[nItem].dLbls),
				dPt:              this.SerDataPoints(arrBubbleSeries[nItem].dPt),
				errBars:          this.SerErrBars(arrBubbleSeries[nItem].errBars),
				idx:              arrBubbleSeries[nItem].idx,
				invertIfNegative: arrBubbleSeries[nItem].invertIfNegative,
				order:            arrBubbleSeries[nItem].order,
				spPr:             this.SerSpPr(arrBubbleSeries[nItem].spPr),
				trendline:        this.SerTrendline(arrBubbleSeries[nItem].trendline),
				tx:               this.SerSerTx(arrBubbleSeries[nItem].tx),
				xVal:             this.SerCat(arrBubbleSeries[nItem].xVal),
				yVal:             this.SerYVAL(arrBubbleSeries[nItem].yVal)
			});
		}

		return arrResultSeries;
	};
	WriterToJSON.prototype.SerScatterChart = function(oScatterChart)
	{
		var sScatterStyle = undefined;
		switch(oScatterChart.scatterStyle)
		{
			case AscFormat.SCATTER_STYLE_LINE:
				sScatterStyle = "line";
				break;
			case AscFormat.SCATTER_STYLE_LINE_MARKER:
				sScatterStyle = "lineMarker";
				break;
			case AscFormat.SCATTER_STYLE_MARKER:
				sScatterStyle = "marker";
				break;
			case AscFormat.SCATTER_STYLE_NONE:
				sScatterStyle = "none";
				break;
			case AscFormat.SCATTER_STYLE_SMOOTH:
				sScatterStyle = "smooth";
				break;
			case AscFormat.SCATTER_STYLE_SMOOTH_MARKER:
				sScatterStyle = "smoothMarker";
				break;
		}
		
		var arrAxId = [];
		for (var nAxis = 0; nAxis < oScatterChart.axId.length; nAxis++)
			arrAxId.push(oScatterChart.axId[nAxis].axId);

		return {
			axId:         arrAxId,
			dLbls:        this.SerDLbls(oScatterChart.dLbls),
			scatterStyle: sScatterStyle,
			ser:          this.SerScatterSeries(oScatterChart.series),
			type:         "scatterChart"
		}
	};
	WriterToJSON.prototype.SerScatterSeries = function(arrScatterSeries)
	{
		var arrResultSeries = [];

		for (var nItem = 0; nItem < arrScatterSeries.length; nItem++)
		{
			arrResultSeries.push({
				dLbls:            this.SerDLbls(arrScatterSeries[nItem].dLbls),
				dPt:              this.SerDataPoints(arrScatterSeries[nItem].dPt),
				errBars:          this.SerErrBars(arrScatterSeries[nItem].errBars),
				idx:              arrScatterSeries[nItem].idx,
				marker:           this.SerMarker(arrScatterSeries[nItem].marker),
				order:            arrScatterSeries[nItem].order,
				smooth:           arrScatterSeries[nItem].smooth,
				spPr:             this.SerSpPr(arrScatterSeries[nItem].spPr),
				trendline:        this.SerTrendline(arrScatterSeries[nItem].trendline),
				tx:               this.SerSerTx(arrScatterSeries[nItem].tx),
				xVal:             this.SerCat(arrScatterSeries[nItem].xVal),
				yVal:             this.SerYVAL(arrScatterSeries[nItem].yVal)
			});
		}

		return arrResultSeries;
	};
	WriterToJSON.prototype.SerStockChart = function(oStockChart)
	{
		var arrAxId = [];
		for (var nAxis = 0; nAxis < oStockChart.axId.length; nAxis++)
			arrAxId.push(oStockChart.axId[nAxis].axId);

		return {
			axId:       arrAxId,
			dLbls:      this.SerDLbls(oStockChart.dLbls),
			dropLines:  this.SerSpPr(oStockChart.dropLines),
			hiLowLines: this.SerSpPr(oStockChart.hiLowLines),
			ser:        this.SerLineSeries(oStockChart.series),
			upDownBars: this.SerUpDownBars(oStockChart.upDownBars),
			type:       "stockChart"
		}
	};
	WriterToJSON.prototype.SerStockSeries = function(arrStockSeries)
	{
		var arrResultSeries = [];

		for (var nItem = 0; nItem < arrStockSeries.length; nItem++)
		{
			arrResultSeries.push({
				cat:              this.SerCat(arrStockSeries[nItem].cat),
				dLbls:            this.SerDLbls(arrStockSeries[nItem].dLbls),
				dPt:              this.SerDataPoints(arrStockSeries[nItem].dPt),
				errBars:          this.SerErrBars(arrStockSeries[nItem].errBars),
				idx:              arrStockSeries[nItem].idx,
				marker:           this.SerMarker(arrStockSeries[nItem].marker),
				order:            arrStockSeries[nItem].order,
				smooth:           arrStockSeries[nItem].smooth,
				spPr:             this.SerSpPr(arrStockSeries[nItem].spPr),
				trendline:        this.SerTrendline(arrStockSeries[nItem].trendline),
				tx:               this.SerSerTx(arrStockSeries[nItem].tx),
				val:              this.SerYVAL(arrStockSeries[nItem].val)
			});
		}

		return arrResultSeries;
	};
	WriterToJSON.prototype.SerAreaChart = function(oAreaChart)
	{
		var sGroupingType = undefined;
		switch (oAreaChart.grouping)
		{
			case AscFormat.GROUPING_PERCENT_STACKED:
				sGroupingType = "percentStacked";
				break;
			case AscFormat.GROUPING_STACKED:
				sGroupingType = "stacked";
				break;
			case AscFormat.GROUPING_STANDARD:
				sGroupingType = "standard";
				break;
		
		}
		
		var arrAxId = [];
		for (var nAxis = 0; nAxis < oAreaChart.axId.length; nAxis++)
			arrAxId.push(oAreaChart.axId[nAxis].axId);

		return {
			axId:       arrAxId,
			dLbls:      this.SerDLbls(oAreaChart.dLbls),
			dropLines:  this.SerSpPr(oAreaChart.dropLines),
			grouping:   sGroupingType,
			ser:        this.SerAreaSeries(oAreaChart.series),
			varyColors: oAreaChart.varyColors,
			type:       "areaChart"
		}
	};
	WriterToJSON.prototype.SerAreaSeries = function(arrAreaSeries)
	{
		var arrResultSeries = [];

		for (var nItem = 0; nItem < arrAreaSeries.length; nItem++)
		{
			arrResultSeries.push({
				cat:              this.SerCat(arrAreaSeries[nItem].cat),
				dLbls:            this.SerDLbls(arrAreaSeries[nItem].dLbls),
				dPt:              this.SerDataPoints(arrAreaSeries[nItem].dPt),
				errBars:          this.SerErrBars(arrAreaSeries[nItem].errBars),
				idx:              arrAreaSeries[nItem].idx,
				order:            arrAreaSeries[nItem].order,
				pictureOptions:   this.SerPicOptions(arrAreaSeries[nItem].pictureOptions),
				spPr:             this.SerSpPr(arrAreaSeries[nItem].spPr),
				trendline:        this.SerTrendline(arrAreaSeries[nItem].trendline),
				tx:               this.SerSerTx(arrAreaSeries[nItem].tx),
				val:              this.SerYVAL(arrAreaSeries[nItem].val)
			});
		}

		return arrResultSeries;
	};
	WriterToJSON.prototype.SerPieChart = function(oPieChart)
	{
		return {
			b3D:           oPieChart.b3D,
			firstSliceAng: oPieChart.firstSliceAng,
			dLbls:         this.SerDLbls(oPieChart.dLbls),
			ser:           this.SerPieSeries(oPieChart.series),
			varyColors:    oPieChart.varyColors,
			type:          "pieChart"
		}
	};
	WriterToJSON.prototype.SerPieSeries = function(arrPieSeries)
	{
		var arrResultSeries = [];

		for (var nItem = 0; nItem < arrPieSeries.length; nItem++)
		{
			arrResultSeries.push({
				cat:              this.SerCat(arrPieSeries[nItem].cat),
				dLbls:            this.SerDLbls(arrPieSeries[nItem].dLbls),
				dPt:              this.SerDataPoints(arrPieSeries[nItem].dPt),
				explosion:        arrPieSeries[nItem].explosion,
				idx:              arrPieSeries[nItem].idx,
				order:            arrPieSeries[nItem].order,
				spPr:             this.SerSpPr(arrPieSeries[nItem].spPr),
				tx:               this.SerSerTx(arrPieSeries[nItem].tx),
				val:              this.SerYVAL(arrPieSeries[nItem].val)
			});
		}

		return arrResultSeries;
	};
	WriterToJSON.prototype.SerBarChart = function(oBarChart)
	{
		var sBarDirType = oBarChart.barDir === AscFormat.BAR_DIR_BAR ? "bar" : "col";
		var sGroupingType = undefined;
		switch (oBarChart.grouping)
		{
			case AscFormat.BAR_GROUPING_CLUSTERED:
				sGroupingType = "clustered";
				break;
			case AscFormat.BAR_GROUPING_PERCENT_STACKED:
				sGroupingType = "percentStacked";
				break;
			case AscFormat.BAR_GROUPING_STACKED:
				sGroupingType = "stacked";
				break;
			case AscFormat.BAR_GROUPING_STANDARD:
				sGroupingType = "standard";
				break;
		
		}

		var arrAxId = [];
		for (var nAxis = 0; nAxis < oBarChart.axId.length; nAxis++)
			arrAxId.push(oBarChart.axId[nAxis].axId);

		return {
			b3D:        oBarChart.b3D,
			axId:       arrAxId,
			barDir:     sBarDirType,
			dLbls:      this.SerDLbls(oBarChart.dLbls),
			gapWidth:   oBarChart.gapWidth,
			grouping:   sGroupingType,
			overlap:    oBarChart.overlap,
			serLines:   this.SerSpPr(oBarChart.serLines),
			ser:        this.SerBarSeries(oBarChart.series),
			varyColors: oBarChart.varyColors,
			type:       "barChart"
		}
	};
	WriterToJSON.prototype.SerLineChart = function(oLineChart)
	{
		var sGroupingType = undefined;
		switch (oLineChart.grouping)
		{
			case AscFormat.GROUPING_PERCENT_STACKED:
				sGroupingType = "percentStacked";
				break;
			case AscFormat.GROUPING_STACKED:
				sGroupingType = "stacked";
				break;
			case AscFormat.GROUPING_STANDARD:
				sGroupingType = "standard";
				break;
		
		}
		
		var arrAxId = [];
		for (var nAxis = 0; nAxis < oLineChart.axId.length; nAxis++)
			arrAxId.push(oLineChart.axId[nAxis].axId);

		return {
			b3D:        oLineChart.b3D,
			axId:       arrAxId,
			dLbls:      this.SerDLbls(oLineChart.dLbls),
			dropLines:  this.SerSpPr(oLineChart.dropLines),
			grouping:   sGroupingType,
			hiLowLines: this.SerSpPr(oLineChart.hiLowLines),
			marker:     oLineChart.marker,
			ser:        this.SerLineSeries(oLineChart.series),
			smooth:     oLineChart.smooth,
			upDownBars: this.SerUpDownBars(oLineChart.upDownBars),
			varyColors: oLineChart.varyColors,
			type:       "lineChart"
		}
	};
	WriterToJSON.prototype.SerUpDownBars = function(oUpDownBars)
	{
		if (!oUpDownBars)
			return oUpDownBars;

		return {
			downBars: this.SerSpPr(oUpDownBars.downBars),
			gapWidth: oUpDownBars.gapWidth,
			upBars:   this.SerSpPr(oUpDownBars.upBars)
		}
	};
	WriterToJSON.prototype.SerLineSeries = function(arrLineSeries)
	{
		var arrResultSeries = [];

		for (var nItem = 0; nItem < arrLineSeries.length; nItem++)
		{
			arrResultSeries.push({
				cat:              this.SerCat(arrLineSeries[nItem].cat),
				dLbls:            this.SerDLbls(arrLineSeries[nItem].dLbls),
				dPt:              this.SerDataPoints(arrLineSeries[nItem].dPt),
				errBars:          this.SerErrBars(arrLineSeries[nItem].errBars),
				idx:              arrLineSeries[nItem].idx,
				marker:           this.SerMarker(arrLineSeries[nItem].marker),
				order:            arrLineSeries[nItem].order,
				smooth:           arrLineSeries[nItem].smooth,
				spPr:             this.SerSpPr(arrLineSeries[nItem].spPr),
				trendline:        this.SerTrendline(arrLineSeries[nItem].trendline),
				tx:               this.SerSerTx(arrLineSeries[nItem].tx),
				val:              this.SerYVAL(arrLineSeries[nItem].val)
			});
		}

		return arrResultSeries;
	};
	WriterToJSON.prototype.SerTitle = function(oTitle)
	{
		if (!oTitle)
			return oTitle;
		
		return {
			layout:  this.SerLayout(oTitle.layout),
			overlay: oTitle.overlay,
			spPr:    this.SerSpPr(oTitle.spPr),
			tx:      this.SerChartTx(oTitle.tx),
			txPr:    this.SerTxPr(oTitle.txPr)
		}
	};
	WriterToJSON.prototype.SerPrintSettings = function(oPrintSettings)
	{
		if (!oPrintSettings)
			return oPrintSettings;

		return {
			headerFooter: this.SerHeaderFooterChart(oPrintSettings.headerFooter),
			pageMargins:  this.SerPageMarginsChart(oPrintSettings.pageMargins),
			pageSetup:    this.SerPageSetup(oPrintSettings.pageSetup)
		}
	};
	WriterToJSON.prototype.SerPageMarginsChart = function(oPageMarginsChart)
	{
		if (!oPageMarginsChart)
			return oPageMarginsChart;

		return {
			b:      oPageMarginsChart.b,
			footer: oPageMarginsChart.footer,
			header: oPageMarginsChart.header,
			l:      oPageMarginsChart.l,
			r:      oPageMarginsChart.r,
			t:      oPageMarginsChart.t,
		}
	};
	WriterToJSON.prototype.SerPageSetup = function(oPageSetup)
	{
		if (!oPageSetup)
			return oPageSetup;
		
		var sOrientType = undefined;
		switch(oPageSetup.orientation)
		{
			case AscFormat.PAGE_SETUP_ORIENTATION_DEFAULT:
				sOrientType = "default";
				break;
			case AscFormat.PAGE_SETUP_ORIENTATION_PORTRAIT:
				sOrientType = "portrait";
				break;
			case AscFormat.PAGE_SETUP_ORIENTATION_LANDSCAPE:
				sOrientType = "landscape";
				break;
		}

		return {
			blackAndWhite:    oPageSetup.blackAndWhite,
			copies:           oPageSetup.copies,
			draft:            oPageSetup.draft,
			firstPageNumber:  oPageSetup.firstPageNumber,
			horizontalDpi:    oPageSetup.horizontalDpi,
			orientation:      sOrientType,
			paperHeight:      oPageSetup.paperHeight,
			paperSize:        oPageSetup.paperSize,
			paperWidth:       oPageSetup.paperWidth,
			useFirstPageNumb: oPageSetup.useFirstPageNumb,
			verticalDpi:      oPageSetup.verticalDpi
		}
	};
	WriterToJSON.prototype.SerHeaderFooterChart = function(oHeaderFooter)
	{
		if (!oHeaderFooter)
			return oHeaderFooter;

		return	{
			evenFooter:       oHeaderFooter.evenFooter,
			evenHeader:       oHeaderFooter.evenHeader,
			firstFooter:      oHeaderFooter.firstFooter,
			firstHeader:      oHeaderFooter.firstHeader,
			oddFooter:        oHeaderFooter.oddFooter,
			oddHeader:        oHeaderFooter.oddHeader,
			alignWithMargins: oHeaderFooter.alignWithMargins,
			differentFirst:   oHeaderFooter.differentFirst,
			differentOddEven: oHeaderFooter.differentOddEven
		}
	};
	WriterToJSON.prototype.SerWall = function(oWall)
	{
		if (!oWall)
			return oWall;
		
		return {
			pictureOptions: this.SerPicOptions(oWall.pictureOptions),
			spPr:           this.SerSpPr(oWall.spPr),
			thickness:      oWall.thickness
		}
	};
	WriterToJSON.prototype.SerMarker = function(oMarker)
	{
		if (!oMarker)
			return oMarker;
		
		var sSymbolType = undefined;
		switch(oMarker.symbol)
		{
			case AscFormat.SYMBOL_CIRCLE:
				sSymbolType = "circle";
				break;
			case AscFormat.SYMBOL_DASH:
				sSymbolType = "dash";
				break;
			case AscFormat.SYMBOL_DIAMOND:
				sSymbolType = "diamond";
				break;
			case AscFormat.SYMBOL_DOT:
				sSymbolType = "dot";
				break;
			case AscFormat.SYMBOL_NONE:
				sSymbolType = "none";
				break;
			case AscFormat.SYMBOL_PICTURE:
				sSymbolType = "picture";
				break;
			case AscFormat.SYMBOL_PLUS:
				sSymbolType = "plus";
				break;
			case AscFormat.SYMBOL_SQUARE:
				sSymbolType = "square";
				break;
			case AscFormat.SYMBOL_STAR:
				sSymbolType = "star";
				break;
			case AscFormat.SYMBOL_TRIANGLE:
				sSymbolType = "triangle";
				break;
			case AscFormat.SYMBOL_X:
				sSymbolType = "x";
				break;
		}
		
		return {
			size:   oMarker.size,
			spPr:   this.SerSpPr(oMarker.spPr),
			symbol: sSymbolType
		}
	};
	WriterToJSON.prototype.SerPivotFmt = function(oFmt)
	{
		if (!oFmt)
			return oFmt;
		
		return {
			dLbl:   this.SerDlbl(oFmt.dLbl),
			idx:    oFmt.idx,
			marker: this.SerMarker(oFmt.marker),
			spPr:   this.SerSpPr(oFmt.spPr),
			txPr:   this.SerTxPr(oFmt.txPr)
		}
	};
	WriterToJSON.prototype.SerPivotFmts = function(arrFmts)
	{
		var arrResult = [];

		for (var nItem = 0; nItem < arrFmts.length; nItem++)
		{
			arrResult.push(this.SerPivotFmt(arrFmts[nItem]));
		}

		return arrResult;
	};
	WriterToJSON.prototype.SerView3D = function(oView3D)
	{
		if (!oView3D)
			return oView3D;
		
		return {
			depthPercent: oView3D.depthPercent,
			hPercent:     oView3D.hPercent,
			perspective:  oView3D.perspective,
			rAngAx:       oView3D.rAngAx,
			rotX:         oView3D.rotX,
			rotY:         oView3D.rotY
		}
	};
	WriterToJSON.prototype.SerLegendEntry = function(oLegendEntry)
	{
		if (!oLegendEntry)
			return oLegendEntry;
		
		return {
			delete: oLegendEntry.bDelete,
			idx:    oLegendEntry.idx,
			txPr:   this.SerTxPr(oLegendEntry.txPr)
		}
	};
	WriterToJSON.prototype.SerLegendEntries = function(arrEntries)
	{
		var arrResults = [];

		for (var nItem = 0; nItem < arrEntries.length; nItem++)
		{
			arrResults.push(this.SerLegendEntry(arrEntries[nItem]));	
		}

		return arrResults;
	};
	WriterToJSON.prototype.SerBlipFill = function(oBlipFill, oParent)
	{
		if (!oBlipFill)
			return oBlipFill;

		var rasterImageId = oBlipFill.getBase64RasterImageId(true);
		
		return {
			blip: this.SerEffects(oBlipFill.Effects),

			srcRect: oBlipFill.srcRect ? {
				b: oBlipFill.srcRect.b,
				l: oBlipFill.srcRect.l,
				r: oBlipFill.srcRect.r,
				t: oBlipFill.srcRect.t
			} : oBlipFill.srcRect,

			tile: oBlipFill.tile ? {
				algn: GetRectAlgnStrType(oBlipFill.tile.algn),
				flip: oBlipFill.tile.flip,
				sx:   oBlipFill.tile.sx,
				sy:   oBlipFill.tile.sy,
				tx:   oBlipFill.tile.tx,
				ty:   oBlipFill.tile.ty
			} : oBlipFill.tile,

			stretch:       oBlipFill.stretch,
			rotWithShape:  oBlipFill.rotWithShape,
			rasterImageId: rasterImageId,
			type: "blipFill"
		};
	};
	WriterToJSON.prototype.SerUniNvPr = function(oUniNvPr)
	{
		if (!oUniNvPr)
			return oUniNvPr;

		return {
			cNvPr:     this.SerCNvPr(oUniNvPr.cNvPr),
			nvPr:      this.SerNvPr(oUniNvPr.nvPr),
			nvUniSpPr: this.SerNvUniSpPr(oUniNvPr.nvUniSpPr)
		};
	};
	WriterToJSON.prototype.SerNvUniSpPr = function(oNvUniSpPr)
	{
		if (!oNvUniSpPr)
			return oNvUniSpPr;

		return {
			locks     : oNvUniSpPr.locks,
			stCnxIdx  : oNvUniSpPr.stCnxIdx,
			stCnxId   : oNvUniSpPr.stCnxId,
			endCnxIdx : oNvUniSpPr.endCnxIdx,
			endCnxId  : oNvUniSpPr.endCnxId
		}
	};
	WriterToJSON.prototype.SerBullet = function(oBullet)
	{
		if (!oBullet)
			return oBullet;

		function SerBulletColor(oBulletColor)
		{
			if (!oBulletColor)
				return oBulletColor;

			var sBulletColorType = "none";
			switch (oBulletColor.type)
			{
				case AscFormat.BULLET_TYPE_COLOR_NONE:
					sBulletColorType = "none";
					break;
				case AscFormat.BULLET_TYPE_COLOR_CLRTX:
					sBulletColorType = "clrtx";
					break;
				case AscFormat.BULLET_TYPE_COLOR_CLR:
					sBulletColorType = "clr";
					break;	
			}

			return {
				type: sBulletColorType,
				color: this.SerColor(oBulletColor.UniColor)
			}
		}
		
		function SerBulletSize(oBulletSize)
		{
			if (!oBulletSize)
				return oBulletSize;

			var sBulleSizeType = "none";
			switch (oBulletSize.type)
			{
				case AscFormat.BULLET_TYPE_SIZE_NONE:
					sBulleSizeType = "none";
					break;
				case AscFormat.BULLET_TYPE_SIZE_TX:
					sBulleSizeType = "tx";
					break;
				case AscFormat.BULLET_TYPE_SIZE_PCT:
					sBulleSizeType = "pct";
					break;	
				case AscFormat.BULLET_TYPE_SIZE_PTS:
					sBulleSizeType = "pts";
					break;	
			}

			return {
				type: sBulleSizeType,
				val:  oBulletSize.val
			}
		}
		
		function SerBulletTypeFace(oBulletTypeface)
		{
			if (!oBulletTypeface)
				return oBulletTypeface;

			var sBulleTypefaceType = "none";
			switch (oBulletTypeface.type)
			{
				case AscFormat.BULLET_TYPE_TYPEFACE_NONE:
					sBulleTypefaceType = "none";
					break;
				case AscFormat.BULLET_TYPE_TYPEFACE_TX:
					sBulleTypefaceType = "tx";
					break;
				case AscFormat.BULLET_TYPE_TYPEFACE_BUFONT:
					sBulleTypefaceType = "bufont";
					break;	
			}

			return {
				type:     sBulleTypefaceType,
				typeface: oBulletTypeface.typeface
			}
		}
		
		function SerBulletType(oBulletType)
		{
			if (!oBulletType)
				return oBulletType;

			var sBulleType = "none";
			var sAutoNumType = undefined;
			switch (oBulletType.AutoNumType)
			{
				case AscFormat.numbering_presentationnumfrmt_AlphaLcParenBoth:
					sAutoNumType = "alphaLcParenBoth";
					break;
				case AscFormat.numbering_presentationnumfrmt_AlphaLcParenR:
					sAutoNumType = "alphaLcParenR";
					break;
				case AscFormat.numbering_presentationnumfrmt_AlphaLcPeriod:
					sAutoNumType = "alphaLcPeriod";
					break;
				case AscFormat.numbering_presentationnumfrmt_AlphaUcParenBoth:
					sAutoNumType = "alphaUcParenBoth";
					break;
				case AscFormat.numbering_presentationnumfrmt_AlphaUcParenR:
					sAutoNumType = "alphaUcParenR";
					break;
				case AscFormat.numbering_presentationnumfrmt_AlphaUcPeriod:
					sAutoNumType = "alphaUcPeriod";
					break;
				case AscFormat.numbering_presentationnumfrmt_Arabic1Minus:
					sAutoNumType = "arabic1Minus";
					break;
				case AscFormat.numbering_presentationnumfrmt_Arabic2Minus:
					sAutoNumType = "arabic2Minus";
					break;
				case AscFormat.numbering_presentationnumfrmt_ArabicDbPeriod:
					sAutoNumType = "arabicDbPeriod";
					break;
				case AscFormat.numbering_presentationnumfrmt_ArabicDbPlain:
					sAutoNumType = "arabicDbPlain";
					break;
				case AscFormat.numbering_presentationnumfrmt_ArabicParenBoth:
					sAutoNumType = "arabicParenBoth";
					break;
				case AscFormat.numbering_presentationnumfrmt_ArabicParenR:
					sAutoNumType = "arabicParenR";
					break;
				case AscFormat.numbering_presentationnumfrmt_ArabicPeriod:
					sAutoNumType = "arabicPeriod";
					break;
				case AscFormat.numbering_presentationnumfrmt_ArabicPlain:
					sAutoNumType = "arabicPlain";
					break;
				case AscFormat.numbering_presentationnumfrmt_CircleNumDbPlain:
					sAutoNumType = "circleNumDbPlain";
					break;
				case AscFormat.numbering_presentationnumfrmt_CircleNumWdBlackPlain:
					sAutoNumType = "circleNumWdBlackPlain";
					break;
				case AscFormat.numbering_presentationnumfrmt_CircleNumWdWhitePlain:
					sAutoNumType = "circleNumWdWhitePlain";
					break;
				case AscFormat.numbering_presentationnumfrmt_Ea1ChsPeriod:
					sAutoNumType = "ea1ChsPeriod";
					break;
				case AscFormat.numbering_presentationnumfrmt_Ea1ChsPlain:
					sAutoNumType = "ea1ChsPlain";
					break;
				case AscFormat.numbering_presentationnumfrmt_Ea1ChtPeriod:
					sAutoNumType = "ea1ChtPeriod";
					break;
				case AscFormat.numbering_presentationnumfrmt_Ea1ChtPlain:
					sAutoNumType = "ea1ChtPlain";
					break;
				case AscFormat.numbering_presentationnumfrmt_Ea1JpnChsDbPeriod:
					sAutoNumType = "ea1JpnChsDbPeriod";
					break;
				case AscFormat.numbering_presentationnumfrmt_Ea1JpnKorPeriod:
					sAutoNumType = "ea1JpnKorPeriod";
					break;
				case AscFormat.numbering_presentationnumfrmt_Ea1JpnKorPlain:
					sAutoNumType = "ea1JpnKorPlain";
					break;
				case AscFormat.numbering_presentationnumfrmt_Hebrew2Minus:
					sAutoNumType = "hebrew2Minus";
					break;
				case AscFormat.numbering_presentationnumfrmt_HindiAlpha1Period:
					sAutoNumType = "hindiAlpha1Period";
					break;
				case AscFormat.numbering_presentationnumfrmt_HindiAlphaPeriod:
					sAutoNumType = "hindiAlphaPeriod";
					break;
				case AscFormat.numbering_presentationnumfrmt_HindiNumParenR:
					sAutoNumType = "hindiNumParenR";
					break;
				case AscFormat.numbering_presentationnumfrmt_HindiNumPeriod:
					sAutoNumType = "hindiNumPeriod";
					break;
				case AscFormat.numbering_presentationnumfrmt_RomanLcParenBoth:
					sAutoNumType = "romanLcParenBoth";
					break;
				case AscFormat.numbering_presentationnumfrmt_RomanLcParenR:
					sAutoNumType = "romanLcParenR";
					break;
				case AscFormat.numbering_presentationnumfrmt_RomanLcPeriod:
					sAutoNumType = "romanLcPeriod";
					break;
				case AscFormat.numbering_presentationnumfrmt_RomanUcParenBoth:
					sAutoNumType = "romanUcParenBoth";
					break;
				case AscFormat.numbering_presentationnumfrmt_RomanUcParenR:
					sAutoNumType = "romanUcParenR";
					break;
				case AscFormat.numbering_presentationnumfrmt_RomanUcPeriod:
					sAutoNumType = "romanUcPeriod";
					break;
				case AscFormat.numbering_presentationnumfrmt_ThaiAlphaParenBoth:
					sAutoNumType = "thaiAlphaParenBoth";
					break;
				case AscFormat.numbering_presentationnumfrmt_ThaiAlphaParenR:
					sAutoNumType = "thaiAlphaParenR";
					break;
				case AscFormat.numbering_presentationnumfrmt_ThaiAlphaPeriod:
					sAutoNumType = "thaiAlphaPeriod";
					break;
				case AscFormat.numbering_presentationnumfrmt_ThaiNumParenBoth:
					sAutoNumType = "thaiNumParenBoth";
					break;
				case AscFormat.numbering_presentationnumfrmt_ThaiNumParenR:
					sAutoNumType = "thaiNumParenR";
					break;
				case AscFormat.numbering_presentationnumfrmt_ThaiNumPeriod:
					sAutoNumType = "thaiNumPeriod";
					break;
			}

			switch (oBulletType.type)
			{
				case AscFormat.BULLET_TYPE_BULLET_NONE:
					sBulleType = "none";
					break;
				case AscFormat.BULLET_TYPE_BULLET_CHAR:
					sBulleType = "char";
					break;
				case AscFormat.BULLET_TYPE_BULLET_AUTONUM:
					sBulleType = "autonum";
					break;	
				case AscFormat.BULLET_TYPE_BULLET_BLIP:
					sBulleType = "blip";
					break;
			}

			return {
				type:        sBulleType,
				char:        oBulletType.Char,
				autoNumType: sAutoNumType,
				startAt:     oBulletType.startAt
			}
		}
		
		return {
			bulletColor:    SerBulletColor.call(this, oBullet.bulletColor),
			bulletSize:     SerBulletSize.call(this, oBullet.bulletSize),
			bulletTypeface: SerBulletTypeFace.call(this, oBullet.bulletTypeface),
			bulletType:     SerBulletType.call(this, oBullet.bulletType)
		}
	};
	WriterToJSON.prototype.SerSpStyle = function(oStyle)
	{
		if (!oStyle)
			return oStyle;
		
		return {
			lnRef:     this.SerStyleRef(oStyle.lnRef),
			fillRef:   this.SerStyleRef(oStyle.fillRef),
			effectRef: this.SerStyleRef(oStyle.effectRef),
			fontRef:   this.SerFontRef(oStyle.fontRef)
		}
	};
	WriterToJSON.prototype.SerStyleRef = function(oStyleRef)
	{
		if (!oStyleRef)
			return oStyleRef;

		return {
			idx:   oStyleRef.idx,
			color: this.SerColor(oStyleRef.Color)
		}
	};
	WriterToJSON.prototype.SerFontRef = function(oFontRef)
	{
		if (!oFontRef)
			return oFontRef;
			
		return {
			idx:   oFontRef.idx,
			color: this.SerColor(oFontRef.Color)
		}
	};
	WriterToJSON.prototype.SerChartSpace = function(oChartSpace)
	{
		if (!oChartSpace)
			return oChartSpace;

		var aUserShapes = [];
		for (var nUserShape = 0; nUserShape < oChartSpace.userShapes.length; nUserShape++)
			aUserShapes.push(this.SerUserShape(oChartSpace.userShapes[nUserShape]));

		return {
			extX:           private_MM2EMU(oChartSpace.extX),
			extY:           private_MM2EMU(oChartSpace.extY),
			chart:          this.SerChart(oChartSpace.chart),
			chartColors:    this.SerChartColors(oChartSpace.chartColors),
			chartStyle:     this.SerChartStyle(oChartSpace.chartStyle),
			clrMapOvr:      this.SerColorMapOvr(oChartSpace.clrMapOvr),
			date1904:       oChartSpace.date1904,
			externalData:   this.SerExternalData(oChartSpace.externalData),
			lang:           oChartSpace.lang,
			pivotSource:    this.SerPivotSource(oChartSpace.pivotSource),
			printSettings:  this.SerPrintSettings(oChartSpace.printSettings),
			protection:     this.SerProtection(oChartSpace.protection),
			roundedCorners: oChartSpace.roundedCorners,
			spPr:           this.SerSpPr(oChartSpace.spPr),
			style:          oChartSpace.style,
			txPr:           this.SerTxPr(oChartSpace.txPr),
			userShapes:     aUserShapes,
			type:           "chartSpace"
		};
	};
	WriterToJSON.prototype.SerColorMapOvr = function(oClrMap)
	{
		if (!oClrMap)
			return oClrMap;

		var aClpMap = [];
		
		for (var nClrScheme in oClrMap.color_map)
		{
			var sClrSchemeType = null;
			switch (oClrMap.color_map[nClrScheme])
			{
				case c_oAscColorSchemeIndex.Accent1:
					sClrSchemeType = "accent1";
					break;
				case c_oAscColorSchemeIndex.Accent2:
					sClrSchemeType = "accent2";
					break;
				case c_oAscColorSchemeIndex.Accent3:
					sClrSchemeType = "accent3";
					break;
				case c_oAscColorSchemeIndex.Accent4:
					sClrSchemeType = "accent4";
					break;
				case c_oAscColorSchemeIndex.Accent5:
					sClrSchemeType = "accent5";
					break;
				case c_oAscColorSchemeIndex.Accent6:
					sClrSchemeType = "accent6";
					break;
				case c_oAscColorSchemeIndex.Bg1:
					sClrSchemeType = "bg1";
					break;
				case c_oAscColorSchemeIndex.Bg2:
					sClrSchemeType = "bg2";
					break;
				case c_oAscColorSchemeIndex.Dk1:
					sClrSchemeType = "dk1";
					break;
				case c_oAscColorSchemeIndex.Dk2:
					sClrSchemeType = "dk2";
					break;
				case c_oAscColorSchemeIndex.FolHlink:
					sClrSchemeType = "folHlink";
					break;
				case c_oAscColorSchemeIndex.Hlink:
					sClrSchemeType = "hlink";
					break;
				case c_oAscColorSchemeIndex.Lt1:
					sClrSchemeType = "lt1";
					break;
				case c_oAscColorSchemeIndex.Lt2:
					sClrSchemeType = "lt2";
					break;
				case c_oAscColorSchemeIndex.PhClr:
					sClrSchemeType = "phClr";
					break;
				case c_oAscColorSchemeIndex.Tx1:
					sClrSchemeType = "tx1";
					break;
				case c_oAscColorSchemeIndex.Tx2:
					sClrSchemeType = "tx2";
					break;
			}
			aClpMap[nClrScheme] = sClrSchemeType;
		}

		return aClpMap;
	};
	WriterToJSON.prototype.SerUserShape = function(oUserShape)
	{
		if (!oUserShape)
			return oUserShape;

		var sType = undefined;
		if (oUserShape instanceof AscFormat.CRelSizeAnchor)
			sType = "relSizeAnchor";
		else if (oUserShape instanceof AscFormat.CAbsSizeAnchor)
			sType = "absSizeAnchor";

		return {
			fromX:  oUserShape.fromX != undefined ? private_MM2Twips(oUserShape.fromX) : oUserShape.fromX,
			fromY:  oUserShape.fromY != undefined ? private_MM2Twips(oUserShape.fromY) : oUserShape.fromY,
			object: this.SerGraphicObject(oUserShape.object),
			toX:    oUserShape.toX,
			toY:    oUserShape.toY,
			type:   sType
		}
	};
	WriterToJSON.prototype.SerChartColors = function(oChartColors)
	{
		if (!oChartColors)
			return oChartColors;

		var aItems = [];
		for (var nItem = 0; nItem < oChartColors.items.length; nItem++)
		{
			if (oChartColors.items[nItem] instanceof AscFormat.CUniColor)
				aItems.push(this.SerColor(oChartColors.items[nItem]));
			else if (oChartColors.items[nItem] instanceof AscFormat.CColorModifiers)
				aItems.push(this.SerColorModifiers(oChartColors.items[nItem]));
		}

		return {
			id:    oChartColors.id,
			items: aItems,
			meth:  oChartColors.meth
		}
	};
	WriterToJSON.prototype.SerChartStyle = function(oChartStyle)
	{
		if (!oChartStyle)
			return oChartStyle;

		return {
			id:                 oChartStyle.id,
			axisTitle:          this.SerStyleEntry(oChartStyle.axisTitle),
			categoryAxis:       this.SerStyleEntry(oChartStyle.categoryAxis),
			chartArea:          this.SerStyleEntry(oChartStyle.chartArea),
			dataLabel:          this.SerStyleEntry(oChartStyle.dataLabel),
			dataLabelCallout:   this.SerStyleEntry(oChartStyle.dataLabelCallout),
			dataPoint:          this.SerStyleEntry(oChartStyle.dataPoint),
			dataPoint3D:        this.SerStyleEntry(oChartStyle.dataPoint3D),
			dataPointLine:      this.SerStyleEntry(oChartStyle.dataPointLine),
			dataPointMarker:    this.SerStyleEntry(oChartStyle.dataPointMarker),
			dataPointWireframe: this.SerStyleEntry(oChartStyle.dataPointWireframe),
			dataTable:          this.SerStyleEntry(oChartStyle.dataTable),
			downBar:            this.SerStyleEntry(oChartStyle.downBar),
			dropLine:           this.SerStyleEntry(oChartStyle.dropLine),
			errorBar:           this.SerStyleEntry(oChartStyle.errorBar),
			floor:              this.SerStyleEntry(oChartStyle.floor),
			gridlineMajor:      this.SerStyleEntry(oChartStyle.gridlineMajor),
			gridlineMinor:      this.SerStyleEntry(oChartStyle.gridlineMinor),
			hiLoLine:           this.SerStyleEntry(oChartStyle.hiLoLine),
			leaderLine:         this.SerStyleEntry(oChartStyle.leaderLine),
			legend:             this.SerStyleEntry(oChartStyle.legend),
			plotArea:           this.SerStyleEntry(oChartStyle.plotArea),
			plotArea3D:         this.SerStyleEntry(oChartStyle.plotArea3D),
			seriesAxis:         this.SerStyleEntry(oChartStyle.seriesAxis),
			seriesLine:         this.SerStyleEntry(oChartStyle.seriesLine),
			title:              this.SerStyleEntry(oChartStyle.title),
			trendline:          this.SerStyleEntry(oChartStyle.trendline),
			trendlineLabel:     this.SerStyleEntry(oChartStyle.trendlineLabel),
			upBar:              this.SerStyleEntry(oChartStyle.upBar),
			valueAxis:          this.SerStyleEntry(oChartStyle.valueAxis),
			wall:               this.SerStyleEntry(oChartStyle.wall),
			markerLayout:       this.SerMarkerLayout(oChartStyle.markerLayout)
		}
	};
	WriterToJSON.prototype.SerStyleEntry = function(oStyleEntry)
	{
		if (!oStyleEntry)
			return oStyleEntry;

		return {
			type:           oStyleEntry.type,
			lineWidthScale: oStyleEntry.lineWidthScale,
			lnRef:          this.SerStyleRef(oStyleEntry.lnRef),
			fillRef:        this.SerStyleRef(oStyleEntry.fillRef),
			effectRef:      this.SerStyleRef(oStyleEntry.effectRef),
			fontRef:        this.SerFontRef(oStyleEntry.fontRef),
			defRPr:         this.SerTextPr(oStyleEntry.defRPr),
			bodyPr:         this.SerBodyPr(oStyleEntry.bodyPr),
			spPr:           this.SerSpPr(oStyleEntry.spPr),
		}
	};
	WriterToJSON.prototype.SerMarkerLayout = function(oMarkerLayout)
	{
		if (!oMarkerLayout)
			return oMarkerLayout;

		return {
			size:   oMarkerLayout.size,
			symbol: oMarkerLayout.symbol
		}
	};
	WriterToJSON.prototype.SerChart = function(oChart)
	{
		if (!oChart)
			return oChart;

		var sDispBlanksAs = undefined;
		switch(oChart.dispBlanksAs)
		{
			case AscFormat.DISP_BLANKS_AS_SPAN:
				sDispBlanksAs = "span";
				break;
			case AscFormat.DISP_BLANKS_AS_GAP:
				sDispBlanksAs = "gap";
				break;
			case AscFormat.DISP_BLANKS_AS_ZERO:
				sDispBlanksAs = "zero";
				break;
		}

		return {
			autoTitleDeleted: oChart.autoTitleDeleted,
			backwall:         this.SerWall(oChart.backWall),
			dispBlanksAs:     sDispBlanksAs,
			floor:            this.SerWall(oChart.floor),
			legend:           this.SerLegend(oChart.legend),
			pivotFmts:        this.SerPivotFmts(oChart.pivotFmts),
			plotArea:         this.SerPlotArea(oChart.plotArea),
			plotVisOnly:      oChart.plotVisOnly,
			showDLblsOverMax: oChart.showDLblsOverMax,
			sideWall:         this.SerWall(oChart.sideWall),
			title:            this.SerTitle(oChart.title),
			view3D:           this.SerView3D(oChart.view3D)
		}
	};
	WriterToJSON.prototype.SerLegend = function(oLegend)
	{
		if (!oLegend)
			return oLegend;
		
		var sLegendPos = undefined;
		switch (oLegend.legendPos)
		{
			case Asc.c_oAscChartLegendShowSettings.bottom:
				sLegendPos = "b";
				break;
			case Asc.c_oAscChartLegendShowSettings.left:
				sLegendPos = "l";
				break;
			case Asc.c_oAscChartLegendShowSettings.right:
				sLegendPos = "r";
				break;
			case Asc.c_oAscChartLegendShowSettings.top:
				sLegendPos = "t";
				break;
			case Asc.c_oAscChartLegendShowSettings.topRight:
				sLegendPos = "tr";
				break;
		}

		return	{
			layout:      this.SerLayout(oLegend.layout),
			legendEntry: this.SerLegendEntries(oLegend.legendEntryes),
			legendPos:   sLegendPos,
			overlay:     oLegend.overlay,
			spPr:        this.SerSpPr(oLegend.spPr),
			txPr:        this.SerTxPr(oLegend.txPr)
		}
	};
	WriterToJSON.prototype.SerImage = function(oImgObject)
	{
		if (!oImgObject)
			return oImgObject;
		
		return {
			extX:     private_MM2EMU(oImgObject.extX),
			extY:     private_MM2EMU(oImgObject.extY),
			blipFill: this.SerBlipFill(oImgObject.blipFill),
			nvPicPr:  this.SerUniNvPr(oImgObject.nvPicPr),
			spPr:     this.SerSpPr(oImgObject.spPr),
			type:     "image"
		}
	};
	WriterToJSON.prototype.SerShape = function(oShape)
	{
		if (!oShape)
			return oShape;

		var oSerContent = null;
		if (oShape.textBoxContent)
			oSerContent = this.SerDocContent(oShape.textBoxContent);
		else if (oShape.txBody)
			oSerContent = this.SerTxPr(oShape.txBody);
		
		return {
			bWordShape:  oShape.bWordShape,
			extX:        private_MM2EMU(oShape.extX),
			extY:        private_MM2EMU(oShape.extY),
			nvSpPr:      this.SerUniNvPr(oShape.nvSpPr),
			spPr:        this.SerSpPr(oShape.spPr),
			style:       this.SerSpStyle(oShape.style),
			bodyPr:      this.SerBodyPr(oShape.bodyPr),
			content:     oSerContent,
			modelId:     oShape.modelId,
			txXfrm:      this.SerXfrm(oShape.txXfrm),
			type:        "shape",
		}
	};
	WriterToJSON.prototype.SerTextPr = function(oTextPr)
	{
		if (!oTextPr)
			return oTextPr;
		
		var oStyles = private_GetStyles();
		// run style
		var oRunStyle = oTextPr.RStyle ? oStyles.Get(oTextPr.RStyle) : undefined;
		
		var sVAlign = undefined;

		// alignV
		if (oTextPr.VertAlign)
		{
			switch (oTextPr.VertAlign)
			{
				case 0:
					sVAlign = "baseline";
					break;
				case 1:
					sVAlign = "superscript";
					break;
				case 2:
					sVAlign = "subscript";
					break;
			}
		}
		return {
			b:         oTextPr.Bold,
			bCs:       oTextPr.BoldCs,
			caps:      oTextPr.Caps,
			color:     oTextPr.Color ? {
				auto: oTextPr.Color.Auto,
				r:    oTextPr.Color.r,
				g:    oTextPr.Color.g,
				b:    oTextPr.Color.b
			} : oTextPr.Color,
			cs:        oTextPr.CS,
			dstrike:   oTextPr.DStrikeout,
			highlight: oTextPr.HighLight ? (oTextPr.HighLight !== -1 ? {
				auto: oTextPr.HighLight.Auto,
				r:    oTextPr.HighLight.r,
				g:    oTextPr.HighLight.g,
				b:    oTextPr.HighLight.b
			} : "none") : oTextPr.HighLight,
			i:         oTextPr.Italic,
			iCs:       oTextPr.ItalicCS,
			lang:      oTextPr.Lang ? {
				bidi:     oTextPr.Lang.Bidi,
				eastAsia: oTextPr.Lang.EastAsia,
				val:      oTextPr.Lang.Val
			} : oTextPr.Lang,
			outline:   this.SerLn(oTextPr.TextOutline),
			position:  oTextPr.Position ? 2.0 * private_MM2Pt(oTextPr.Position) : oTextPr.Position, /// нужно ли умножать на 2?
			rFonts:    oTextPr.RFonts ? {
				ascii:         oTextPr.RFonts.Ascii,
				asciiTheme:    oTextPr.RFonts.AsciiTheme,
				cs:            oTextPr.RFonts.CS,
				cstheme:       oTextPr.RFonts.CSTheme,
				eastAsia:      oTextPr.RFonts.EastAsia,
				eastAsiaTheme: oTextPr.RFonts.EastAsiaTheme,
				hAnsi:         oTextPr.RFonts.HAnsi,
				hAnsiTheme:    oTextPr.RFonts.HAnsiTheme,
				hint:          oTextPr.RFonts.Hint
			} : oTextPr.RFonts,
			fontFamily: oTextPr.FontFamily ? {
				name: oTextPr.FontFamily.Name,
				idx:  oTextPr.FontFamily.Index
			} : oTextPr.FontFamily,
			rPrChange: this.SerTextPr(oTextPr.PrChange),
			rStyle:    this.SerStyle(oRunStyle),
			rtl:       oTextPr.RTL,
			shd:       this.SerShd(oTextPr.Shd),
			smallCaps: oTextPr.SmallCaps,
			spacing:   oTextPr.Spacing ? private_MM2Twips(oTextPr.Spacing) : oTextPr.Spacing,
			strike:    oTextPr.Strikeout,
			sz:        oTextPr.FontSize ? 2.0 * oTextPr.FontSize : oTextPr.FontSize,
			szCs:      oTextPr.FontSizeCS ? 2.0 * oTextPr.FontSizeCS : oTextPr.FontSizeCS,
			u:         oTextPr.Underline,
			vanish:    oTextPr.Vanish,
			vertAlign: sVAlign,
			fontRef:   this.SerFontRef(oTextPr.FontRef),
			uniFill:   this.SerFill(oTextPr.Unifill),
			textFill:  this.SerFill(oTextPr.TextFill),
			reviewInfo: this.SerReviewInfo(oTextPr.ReviewInfo),
			type:      "textPr"
		}
	};
	WriterToJSON.prototype.SerSdtPr = function(oSdtPr)
	{
		if (!oSdtPr)
			return oSdtPr;

		var arrListItemComboBox = [];
		if (oSdtPr.ComboBox)
		{
			for (var nItem = 0; nItem < oSdtPr.ComboBox.ListItems; nItem++)
			{
				arrListItemComboBox.push({
					displayText: oSdtPr.ComboBox.ListItems[nItem].DisplayText,
					value: oSdtPr.ComboBox.ListItems[nItem].Value,
				});
			}
		}

		var arrListItemDropDown = [];
		if (oSdtPr.DropDown)
		{
			for (var nItem = 0; nItem < oSdtPr.DropDown.ListItems; nItem++)
			{
				arrListItemDropDown.push({
					displayText: oSdtPr.DropDown.ListItems[nItem].DisplayText,
					value: oSdtPr.DropDown.ListItems[nItem].Value,
				});
			}
		}

		var sLockType = undefined;
		switch (oSdtPr.Lock)
		{
			case Asc.c_oAscSdtLockType.ContentLocked:
				sLockType = "contentLocked";
				break;
			case Asc.c_oAscSdtLockType.SdtContentLocked:
				sLockType = "sdtContentLocked";
				break;
			case Asc.c_oAscSdtLockType.SdtLocked:
				sLockType = "sdtLocked";
				break;
			case Asc.c_oAscSdtLockType.Unlocked:
				sLockType = "unlocked";
				break;
		}
		return {
			alias: oSdtPr.Alias,

			comboBox: oSdtPr.ComboBox ? {
				listItem:  arrListItemComboBox,
				lastValue: oSdtPr.ComboBox.LastValue
			} : oSdtPr.ComboBox,

			date: oSdtPr.Date ? {
				calendar:   oSdtPr.Date.Calendar,
				dateFormat: oSdtPr.Date.DateFormat,
				lid:        oSdtPr.Date.LangId,
				fullDate:   oSdtPr.Date.FullDate
			} : oSdtPr.Date,

			docPartObj: oSdtPr.DocPartObj ? {
				docPartCategory: oSdtPr.DocPartObj.Category,
				docPartGallery:  oSdtPr.DocPartObj.Gallery,
				docPartUnique:   oSdtPr.DocPartObj.Unique
			} : oSdtPr.DocPartObj,

			dropDownList: oSdtPr.DropDown ? {
				listItem:  arrListItemDropDown,
				lastValue: oSdtPr.DropDown.LastValue
			} : oSdtPr.DropDown,

			equation : oSdtPr.Equation,
			id:        oSdtPr.Id,
			label:     oSdtPr.Label,
			lock:      sLockType,
			picture:   oSdtPr.Picture,

			placeholder: oSdtPr.Placeholder ? {
				docPart: oSdtPr.Placeholder
			} : oSdtPr.Placeholder,

			rPr:           this.SerTextPr(oSdtPr.TextPr),
			showingPlcHdr: oSdtPr.ShowingPlcHdr,
			tag:           oSdtPr.Tag,
			temporary:     oSdtPr.Temporary,
			text:          {
				multiLine: oSdtPr.Text
			}
		}
	};
	WriterToJSON.prototype.SerParaMath = function(oParaMath)
	{
		if (!oParaMath)
			return oParaMath;

		var sJc = undefined;
		switch (oParaMath.Jc)
		{
			case JC_CENTER:
				sJc = "center";
				break;
			case JC_CENTERGROUP:
				sJc = "centerGroup";
				break;
			case JC_LEFT:
				sJc = "left";
				break;
			case JC_RIGHT:
				sJc = "right";
				break;
		}

		function SerMathContent(oMathContent)
		{
			if (!oMathContent)
				return oMathContent;

			var oTempElm   = null;
			var arrContent = [];

			if (oMathContent.constructor.name === "CDenominator" || oMathContent.constructor.name === "CNumerator")
				arrContent.push(SerFracArg.call(this, oMathContent));
			else 
			{
				for (var nElm = 0; nElm < oMathContent.Content.length; nElm++)
				{
					oTempElm = oMathContent.Content[nElm];
					if (oTempElm instanceof AscCommonWord.ParaRun)
						arrContent.push(this.SerParaRun(oTempElm));
					else if (oTempElm instanceof AscCommonWord.CFraction)
						arrContent.push(SerFraction.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CDegree)
						arrContent.push(SerDegree.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CDegreeSubSup)
						arrContent.push(SerSupSubDegree.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CRadical)
						arrContent.push(SerRadical.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CNary)
						arrContent.push(SerNary.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CEqArray)
						arrContent.push(SerEqArray.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CDelimiter)
						arrContent.push(SerDelimiter.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CMathFunc)
						arrContent.push(SerMathFunc.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CAccent)
						arrContent.push(SerAccent.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CGroupCharacter)
						arrContent.push(SerGroupCharacter.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CBorderBox)
						arrContent.push(SerBorderBox.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CBar)
						arrContent.push(SerBar.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CLimit)
						arrContent.push(SerLimit.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CBox)
						arrContent.push(SerBox.call(this, oTempElm));
					else if (oTempElm instanceof AscCommonWord.CMathMatrix)
						arrContent.push(SerMathMatrix.call(this, oTempElm));
				}
			}

			return arrContent;
		}

		function SerFraction(oFraction)
		{
			if (!oFraction)
				return oFraction;
			
			var sFracType;
			switch (oFraction.Pr.type)
			{
				case BAR_FRACTION:
					sFracType = "bar";
					break;
				case SKEWED_FRACTION:
					sFracType = "skw";
					break;
				case LINEAR_FRACTION:
					sFracType = "lin";
					break;
				default:
					sFracType = "noBar";
					break;
			}

			return {
				fPr: {
					ctrlPr: this.SerTextPr(oFraction.CtrPrp),
					type:   sFracType
				},
				num:  SerFracArg.call(this, oFraction.Numerator),
				den:  SerFracArg.call(this, oFraction.Denominator),     
				type: "fraction"
			}
		}

		function SerFracArg(oFractionArg)
		{
			if (!oFractionArg)
				return oFractionArg;

			var sArgType = "";
			if (oFractionArg.constructor.name === "CDenominator")
				sArgType = "den";
			else if (oFractionArg.constructor.name === "CNumerator")
				sArgType = "num";

			return {
				argPr: {
					argSize: oFractionArg.ArgSize.value
				},  
				ctrlPr:  this.SerTextPr(oFractionArg.CtrPrp),
				content: SerMathContent.call(this, oFractionArg.elements[0][0]),
				type:    sArgType
			}
		}
		function SerDegree(oDegree)
		{
			if (!oDegree)
				return oDegree;

			var oDegreeObj = {
				e: SerMathContent.call(this, oDegree.baseContent),
			}
			switch (oDegree.Pr.type)
			{
				case DEGREE_SUPERSCRIPT:
					oDegreeObj.type   = "superScript";
					oDegreeObj.sSupPr = {
						ctrlPr: this.SerTextPr(oDegree.CtrPrp)
					}
					oDegreeObj.sup    = SerMathContent.call(this, oDegree.iterContent);
					break;
				case DEGREE_SUBSCRIPT:
					oDegreeObj.type   = "supScript";
					oDegreeObj.sSubPr = {
						ctrlPr: this.SerTextPr(oDegree.CtrPrp)
					}
					oDegreeObj.sub    = SerMathContent.call(this, oDegree.iterContent);
					break;
			}

			return oDegreeObj;
		}

		function SerSupSubDegree(oDegreeSubSup)
		{
			if (!oDegreeSubSup)
				return oDegreeSubSup;

			var oDegreeObj = {
				e:   SerMathContent.call(this, oDegreeSubSup.baseContent),
				sup: SerMathContent.call(this, oDegreeSubSup.iters.iterUp),
				sub: SerMathContent.call(this, oDegreeSubSup.iters.iterDn)
			}

			switch (oDegreeSubSup.Pr.type)
			{
				case DEGREE_SubSup:
					oDegreeObj.type      = "subSupScript";
					oDegreeObj.sSubSupPr = {
						ctrlPr: this.SerTextPr(oDegreeSubSup.CtrPrp)
					}
					break;
				case DEGREE_PreSubSup:
					oDegreeObj.type      = "preSubSupScript";
					oDegreeObj.sPrePr    = {
						ctrlPr: this.SerTextPr(oDegreeSubSup.CtrPrp)
					}
					break;
			}


			return oDegreeObj;
		}

		function SerRadical(oRadical)
		{
			if (!oRadical)
				return oRadical;

			var sRadType = oRadical.Pr.type === SQUARE_RADICAL ? "radSquare" : "radDegree";

			return {
				radPr: {
					ctrlPr:  this.SerTextPr(oRadical.CtrPrp),
					degHide: oRadical.Pr.degHide
				},
				e:     SerMathContent.call(this, oRadical.RealBase),
				deg:   SerMathContent.call(this, oRadical.Iterator),
				type:  sRadType
			}
		}

		function SerNary(oNary)
		{
			if (!oNary)
				return oNary;

			var sLimLoc = "";
			switch (oNary.Pr.limLoc)
			{
				case NARY_UndOvr:
					sLimLoc = "undOvr";
					break;
				case NARY_SubSup:
					sLimLoc = "subSup";
					break;
			}
			return {
				e:      SerMathContent.call(this, oNary.Arg),
				sup:    SerMathContent.call(this, oNary.UpperIterator),
				sub:    SerMathContent.call(this, oNary.LowerIterator),
				naryPr: {
					chr:     oNary.Pr.chr,
					ctrlPr:  this.SerTextPr(oNary.CtrPrp),
					grow:    oNary.Pr.grow,
					limLoc:  sLimLoc,
					subHide: oNary.Pr.subHide,
					supHide: oNary.Pr.supHide
				},
				type:   "nary"
			}
		}

		function SerEqArray(oEqArray)
		{
			if (!oEqArray)
				return oEqArray;

			
			var sJcType = undefined;
			switch (oEqArray.Pr.baseJc)
			{
				case BASEJC_CENTER:
					sJcType = "center";
					break;
				case BASEJC_TOP:
					sJcType = "top";
					break;
				case BASEJC_BOTTOM:
					sJcType = "bottom";
					break;
				case BASEJC_INLINE:
					sJcType = "inline";
					break;
				case BASEJC_INSIDE:
					sJcType = "inside";
					break;
				case BASEJC_OUTSIDE:
					sJcType = "outside";
					break;
			}

			var oEqArrayObj = {
				eqArrPr: {
					baseJc:  sJcType,
					ctrlPr:  this.SerTextPr(oEqArray.CtrPrp),
					maxDist: oEqArray.Pr.maxDist,
					objDist: oEqArray.Pr.objDist,
					rSp:     oEqArray.Pr.rSp,
					rSpRule: oEqArray.Pr.rSpRule,
					row:     oEqArray.Pr.row
				},
				e:  [],
				type: "eqArray"
			}

			for (var nArg = 0; nArg < oEqArray.elements.length; nArg++)
				oEqArrayObj.e.push(SerMathContent.call(this, oEqArray.elements[nArg][0]));

			return oEqArrayObj;
		}

		function SerDelimiter(oDelimiter)
		{
			if (!oDelimiter)
				return oDelimiter;

			var sShpType = oDelimiter.Pr.shp === DELIMITER_SHAPE_CENTERED ? "centered" : "match";

			var oDelimiterObj = {
				dPr:  {
					begChr: oDelimiter.Pr.begChr,
					ctrlPr: this.SerTextPr(oDelimiter.CtrPrp),
					endChr: oDelimiter.Pr.endChr,
					grow:   oDelimiter.Pr.grow,
					sepChr: oDelimiter.Pr.sepChr,
					shp:    sShpType
				},
				e:    [],
				type: "delimiter"
			}

			for (var nArg = 0; nArg < oDelimiter.elements[0].length; nArg++)
				oDelimiterObj.e.push(SerMathContent.call(this, oDelimiter.elements[0][nArg]));

			return oDelimiterObj;
		}

		function SerMathFunc(oMathFunc)
		{
			if (!oMathFunc)
				return oMathFunc;

			return {
				fName:  SerMathContent.call(this, oMathFunc.elements[0][0]),
				e:      SerMathContent.call(this, oMathFunc.elements[0][1]),
				funcPr: {
					ctrlPr: this.SerTextPr(oMathFunc.CtrPrp)
				},
				type: "mathFunc"
			}
		}

		function SerAccent(oAccent)
		{
			if (!oAccent)
				return oAccent;

			return {
				accPr: {
					ctrlPr: this.SerTextPr(oAccent.CtrPrp),
					chr:    oAccent.Pr.chr
				},
				e:    SerMathContent.call(this, oAccent.elements[0][0]),
				type: "accent"
			}
		}

		function SerGroupCharacter(oGrpChar)
		{
			if (!oGrpChar)
				return oGrpChar;

			var sPos = undefined;
			switch (oGrpChar.Pr.pos)
			{
				case LOCATION_BOT:
					sPos = "bot";
					break;
				case LOCATION_TOP:
					sPos = "top";
					break;
			}

			var sVertJc = undefined;
			switch (oGrpChar.Pr.vertJc)
			{
				case VJUST_TOP:
					sVertJc = "top";
					break;
				case VJUST_BOT:
					sVertJc = "bot";
					break;
			}

			return {
				groupChrPr: {
					chr:    oGrpChar.Pr.chr,
					ctrlPr: this.SerTextPr(oGrpChar.CtrPrp),
					pos:    sPos,
					vertJc: sVertJc
				},
				e:    SerMathContent.call(this, oGrpChar.getBase()),
				type: "groupChr"
			}
		}

		function SerBorderBox(oBox)
		{
			if (!oBox)
				return oBox;

			return {
				borderBoxPr: {
					ctrlPr:     this.SerTextPr(oBox.CtrPrp),
					hideBot:    oBox.Pr.hideBot,
					hideLeft:   oBox.Pr.hideLeft,
					hideRight:  oBox.Pr.hideRight,
					hideTop:    oBox.Pr.hideTop,
					strikeBLTR: oBox.Pr.strikeBLTR,
					strikeH:    oBox.Pr.strikeH,
					strikeTLBR: oBox.Pr.strikeTLBR,
					strikeV:    oBox.Pr.strikeV
				},
				e:    SerMathContent.call(this, oBox.elements[0][0]),
				type: "borderBox"
			}
		}

		function SerBox(oBox)
		{
			if (!oBox)
				return oBox;

			return {
				boxPr: {
					ctrlPr:  this.SerTextPr(oBox.CtrPrp),
					aln:     oBox.Pr.aln,
					brk:     oBox.Pr.brk ? {
						alnAt: oBox.Pr.brk.alnAt
					} : oBox.Pr.brk,
					diff:    oBox.Pr.diff,
					noBreak: oBox.Pr.noBreak,
					opEmu:   oBox.Pr.opEmu
				},
				e:      SerMathContent.call(this, oBox.getBase()),
				type:   "box"
			}
		}

		function SerBar(oBar)
		{
			if (!oBar)
				return oBar;

			var sPos = undefined;
			switch (oBar.Pr.pos)
			{
				case LOCATION_BOT:
					sPos = "bot";
					break;
				case LOCATION_TOP:
					sPos = "top";
					break;
			}

			return {
				barPr: {
					ctrlPr: this.SerTextPr(oBar.CtrPrp),
					pos:    sPos
				},
				e:     SerMathContent.call(this, oBar.elements[0][0]),
				type:  "bar"
			}
		}

		function SerLimit(oLimit)
		{
			if (!oLimit)
				return oLimit;

			var oLimObj = {
				e:     SerMathContent.call(this, oLimit.getFName()),
				limit: SerMathContent.call(this, oLimit.getIterator())
			}
			
			if (oLimit.Pr.type === LIMIT_UP)
			{
				oLimObj.type     = "limUpp";
				oLimObj.limUppPr = {
					ctrlPr: this.SerTextPr(oLimit.CtrPrp)
				};
			}
			else
			{
				oLimObj.type     = "limLow";
				oLimObj.limLowPr = {
					ctrlPr: this.SerTextPr(oLimit.CtrPrp)
				};
			}
				
			return oLimObj;
		}

		function SerMathMatrix(oMatrix)
		{
			if (!oMatrix)
				return oMatrix;

			var sJcType = undefined;
			switch (oMatrix.Pr.baseJc)
			{
				case BASEJC_CENTER:
					sJcType = "center";
					break;
				case BASEJC_TOP:
					sJcType = "top";
					break;
				case BASEJC_BOTTOM:
					sJcType = "bottom";
					break;
				case BASEJC_INLINE:
					sJcType = "inline";
					break;
				case BASEJC_INSIDE:
					sJcType = "inside";
					break;
				case BASEJC_OUTSIDE:
					sJcType = "outside";
					break;
			}

			var arrMatrixRow = [];

			for (var nRow = 0; nRow < oMatrix.elements.length; nRow++)
			{
				var arrCells = [];
				for (var nCell = 0; nCell < oMatrix.elements[nRow].length; nCell++)
					arrCells.push(SerMathContent.call(this, oMatrix.elements[nRow][nCell]));
				
				arrMatrixRow.push(arrCells);
			}

			var arrMatrixColsPr = [];
			for (var nPr = 0; nPr < oMatrix.Pr.mcs.length; nPr++)
			{
				var sColPrJcType = undefined;
				switch (oMatrix.Pr.mcs[nPr].mcJc)
				{
					case MCJC_CENTER:
						sColPrJcType = "center";
						break;
					case MCJC_LEFT:
						sColPrJcType = "left";
						break;
					case MCJC_RIGHT:
						sColPrJcType = "right";
						break;
					case MCJC_INSIDE:
						sColPrJcType = "inside";
						break;
					case MCJC_OUTSIDE:
						sColPrJcType = "outside";
						break;
				}

				arrMatrixColsPr.push({
					count: oMatrix.Pr.mcs[nPr].count,
					mcJc:  sColPrJcType
				});
			}
			return {
				mPr: {
					baseJc:  sJcType,
					cGp:     oMatrix.Pr.cGp,
					cGpRule: oMatrix.Pr.cGpRule,
					cSp:     oMatrix.Pr.cSp,
					ctrlPr:  this.SerTextPr(oMatrix.CtrPrp),
					mcs:     arrMatrixColsPr,
					plcHide: oMatrix.Pr.plcHide,
					rSp:     oMatrix.Pr.rSp,
					rSpRule: oMatrix.Pr.rSpRule
				},
				mr:   arrMatrixRow,
				type: "matrix"
			}
		}

		return {
			oMathParaPr: {
				jc: sJc
			},
			content: SerMathContent.call(this, oParaMath.Root),
			type: "paraMath"
		}
	};
	WriterToJSON.prototype.SerParaComment = function(oComment, oMapCommentsInfo)
	{
		if (!oComment || !oComment.Paragraph)
			return null;

		if (!oMapCommentsInfo[oComment.CommentId].Start || !oMapCommentsInfo[oComment.CommentId].End)
			return null;

		var oDocument      = private_GetLogicDocument();
		var oCommentData   = oDocument.Comments.GetById(oComment.CommentId).GetData();  
		var isStartComment = oComment.Start;
		var oCommentObj    = {
			id:    oComment.CommentId,
			autor: oCommentData.Get_Name(),
			text:  oCommentData.Get_Text()
		};

		if (isStartComment)
			oCommentObj.type = "commentRangeStart";
		else
			oCommentObj.type = "commentRangeEnd";

		return oCommentObj;
	};
	WriterToJSON.prototype.SerParaBookmark = function(oBookmark, oMapBookmarksInfo)
	{
		if (!oBookmark || !oBookmark.Paragraph)
			return null;

		if (!oMapBookmarksInfo[oBookmark.BookmarkId].Start || !oMapBookmarksInfo[oBookmark.BookmarkId].End)
			return null;

		var isStartBookmark   = oBookmark.Start;
		var oTempBookmark     = null;
		var oBookmarkObj      = {
			id:   oBookmark.BookmarkId,
			name: oBookmark.BookmarkName
		};
		if (isStartBookmark)
			oBookmarkObj.type = "bookmarkStart";
		else
			oBookmarkObj.type = "bookmarkEnd";

		for (var nElm = 0; nElm < oBookmark.Paragraph.Content.length; nElm++)
		{
			if (oBookmark.Paragraph.Content[nElm] instanceof AscCommonWord.CParagraphBookmark)
			{
				oTempBookmark = oBookmark.Paragraph.Content[nElm];
				if (oTempBookmark.CommentId === oBookmark.CommentId && oTempBookmark.Start !== isStartBookmark)
				{
					return oBookmarkObj;
				}
			}
		}

		return null;
	};
	WriterToJSON.prototype.GetPenDashStrType = function(nType)
	{
		switch (nType)
		{
			case Asc.c_oDashType.dash:
				return "dash";
			case Asc.c_oDashType.dashDot:
				return "dashDot";
			case Asc.c_oDashType.dot:
				return "dot";
			case Asc.c_oDashType.lgDash:
				return "lgDash";
			case Asc.c_oDashType.lgDashDot:
				return "lgDashDot";
			case Asc.c_oDashType.lgDashDotDot:
				return "lgDashDotDot";
			case Asc.c_oDashType.solid:
				return "solid";
			case Asc.c_oDashType.sysDash:
				return "sysDash";
			case Asc.c_oDashType.sysDashDot:
				return "sysDashDot";
			case Asc.c_oDashType.sysDashDotDot:
				return "sysDashDotDot";
			case Asc.c_oDashType.sysDot:
				return "sysDot";
			default:
				return nType;
		}
	};

	function ReaderFromJSON()
	{
		this.MoveMap = {};
		this.FootEndNoteMap = {};
		this.layoutsMap     = {};
		this.mastersMap     = {};
		this.notesMasterMap = {};
		this.themesMap      = {};
		this.drawingsMap    = {};
	}

	ReaderFromJSON.prototype.ParaRunFromJSON = function(oParsedRun, oParentPara, notCompletedFields)
	{
		var aContent         = oParsedRun.content;
		var oPr              = oParsedRun.rPr;
		var oCurComplexField = null;
		var oDocument        = private_GetLogicDocument();

		if (!notCompletedFields)
			notCompletedFields = [];

		var oRun;
		switch (oParsedRun.type)
		{
			case "mathRun":
				oRun = new ParaRun(oParentPara, true);
				break;
			case "presField":
				oRun = new AscCommonWord.CPresentationField(oParentPara);
				break;
			default:
				oRun = new ParaRun(oParentPara, false);
				break;
		}

		if (oParsedRun.type === "presField")
		{
			oRun.SetGuid(AscCommon.CreateGUID());
			oRun.SetFieldType(oParsedRun.fldType);
		}

		// Footnotes
		for (var nElm = 0; nElm < oParsedRun.footnotes.length; nElm++)
			this.ParaFootEndNoteFromJSON(oParsedRun.footnotes[nElm]);
		// Endnotes
		for (nElm = 0; nElm < oParsedRun.endnotes.length; nElm++)
			this.ParaFootEndNoteFromJSON(oParsedRun.endnotes[nElm]);

		// review info
		var nReviewType = undefined;
		switch (oParsedRun.reviewType)
		{
			case "common":
				nReviewType = reviewtype_Common;
				break;
			case "remove":
				nReviewType = reviewtype_Remove;
				break;
			case "add":
				nReviewType = reviewtype_Add;
				break;
		}
		var oReviewInfo = oParsedRun.reviewInfo ? this.ReviewInfoFromJSON(oParsedRun.reviewInfo) : oRun.ReviewInfo;
		oRun.SetReviewTypeWithInfo(nReviewType, oReviewInfo);

		oRun.Apply_Pr(this.TextPrFromJSON(oPr));
		oParsedRun.type === "mathRun" && oRun.Math_Apply_Style(oParsedRun.mathPr.sty);

		if (oParsedRun.type === "endRun")
		{
			oRun.Add_ToContent( 0, new AscWord.CRunParagraphMark() );
			return oRun;
		}

		for (nElm = 0; nElm < aContent.length; nElm++)
		{
			// записываем текстовый контент в ран(либо обычный ран либо mathRun)
			if (typeof aContent[nElm] === "string" || aContent[nElm].type === "mathTxt")
			{
				if (oParsedRun.type === "mathRun")
				{
					// if (0x0026 == aContent[nElm].charCodeAt(0))
					// 	var oText = new CMathAmp();
					// else
					// {
					// 	var oText = new CMathText(false);
					// 	if (aContent[nElm][nChar] === String.fromCharCode(StartTextElement))
					// 	{
					// 		oText.SetPlaceholder();
					// 		oRun.Add_ToContent(0, oText, false);
					// 	}
					// 	else
					// 	{
					// 		oText.add(aContent[nElm][nChar]);
					// 		oRun.Add(oText, true);
					// 	}
					// }
					var oText;
					if (0x0026 == aContent[nElm].value)
						oText = new CMathAmp();
					else
					{
						oText = new CMathText(false);
						if (aContent[nElm].value === StartTextElement)
						{
							oText.SetPlaceholder();
							oRun.Add_ToContent(0, oText, false);
						}
						else
						{
							oText.add(aContent[nElm].value);
							oRun.Add(oText);
						}
					}
				}
				else
					oRun.AddText(aContent[nElm]);

				continue;
			}

			switch (aContent[nElm].type)
			{
				case "break":
					switch(aContent[nElm].breakType)
					{
						case "textWrapping":
							oRun.AddToContent(-1, new AscWord.CRunBreak(AscWord.break_Line));
							break;
						case "page":
							oRun.AddToContent(-1, new AscWord.CRunBreak(AscWord.break_Page));
							break;
						case "column":
							oRun.AddToContent(-1, new AscWord.CRunBreak(AscWord.break_Column));
							break;
					}
					break;
				case "pgNum":
					oRun.AddToContent(-1, new AscWord.CRunPageNum());
					break;
				case "tab":
					oRun.AddToContent(-1, new AscWord.CRunTab());
					break;
				case "fldChar":
					switch (aContent[nElm].fldCharType)
					{
						case "begin":
							var oBeginChar    = new ParaFieldChar(fldchartype_Begin, private_GetLogicDocument());
							oRun.AddToContent(-1, oBeginChar);
							oBeginChar.SetRun(oRun);
							oCurComplexField = oBeginChar.GetComplexField();
							oCurComplexField.SetBeginChar(oBeginChar);
							notCompletedFields.push(oCurComplexField);
    						break;
						case "separate":
							var oSeparateChar = new ParaFieldChar(fldchartype_Separate, private_GetLogicDocument());
							oRun.AddToContent(-1, oSeparateChar);
							oSeparateChar.SetRun(oRun);
							oCurComplexField.SetSeparateChar(oSeparateChar);
							break;
						case "end":
							var oEndChar = new ParaFieldChar(fldchartype_End, private_GetLogicDocument());
							oRun.AddToContent(-1, oEndChar);
							oEndChar.SetRun(oRun);
							notCompletedFields[notCompletedFields.length - 1].SetEndChar(oEndChar);
							notCompletedFields.splice(notCompletedFields.length - 1, 1);
							break;
					}
					break;
				case "instrText":
					oRun.AddInstrText(aContent[nElm].instr);
					oCurComplexField.SetInstructionLine(aContent[nElm].instr);
					break;
				case "paraDrawing":
					var oDrawing = this.ParaDrawingFromJSON(aContent[nElm]);
					oRun.Add_ToContent(-1, oDrawing);
					oDrawing.Set_Parent(oRun.Paragraph);
					break;
				case "revisionMove":
					// создаём новый moveId(moveName), мапим к нему значения moveId(moveName) из JSON
					// таким образом задаем соответствия
					var sMoveName = undefined;
					if (!this.MoveMap[aContent[nElm].name])
					{
						sMoveName = oDocument.TrackRevisionsManager.GetNewMoveId();
						this.MoveMap[aContent[nElm].name] = sMoveName;
					}

					var oRevisionMove = new AscWord.CRunRevisionMove(aContent[nElm].start, aContent[nElm].from, this.MoveMap[aContent[nElm].name], aContent[nElm].reviewInfo);
					oRun.Add_ToContent(-1, oRevisionMove);
					break;
				case "footnoteRef":
				case "footnoteNum":
				case "endnoteRef":
				case "endnoteNum":
					oRun.Add_ToContent(-1, this.ParaFootEndNoteRefFromJSON(aContent[nElm]));
					break;
			}
		}

		return oRun;
	};
	ReaderFromJSON.prototype.ParaFootEndNoteRefFromJSON = function(oParsedFootEndnoteRef)
	{
		var oFootEndnoteRef = null;
		switch (oParsedFootEndnoteRef.type)
		{
			case "footnoteRef":
				oFootEndnoteRef = new AscWord.CRunFootnoteReference(this.FootEndNoteMap[oParsedFootEndnoteRef.footnote]);
				break;
			case "footnoteNum":
				oFootEndnoteRef = new AscWord.CRunFootnoteRef(this.FootEndNoteMap[oParsedFootEndnoteRef.footnote]);
				break;
			case "endnoteRef":
				oFootEndnoteRef = new AscWord.CRunEndnoteReference(this.FootEndNoteMap[oParsedFootEndnoteRef.footnote]);
				break;
			case "endnoteNum":
				oFootEndnoteRef = new AscWord.CRunEndnoteRef(this.FootEndNoteMap[oParsedFootEndnoteRef.footnote]);
				break;
		}
		
		oFootEndnoteRef.CustomMark = oParsedFootEndnoteRef.customMark;
		oFootEndnoteRef.NumFormat  = oParsedFootEndnoteRef.numFormat;
		//oFootEndnoteRef.Number   = oParsedFootEndnoteRef.number;

		return oFootEndnoteRef;
	};
	ReaderFromJSON.prototype.ParaFootEndNoteFromJSON = function(oParsedFootEndnote)
	{
		var oDocument    = private_GetLogicDocument();
		var aContent     = oParsedFootEndnote.content;
		var oFootEndnote = oParsedFootEndnote === "footnote" ? oDocument.Footnotes.CreateFootnote() : oDocument.Endnotes.CreateEndnote();
		
		// мапим сносну, для привязки к ссылке внутри рана
		this.FootEndNoteMap[oParsedFootEndnote.id] = oFootEndnote;

		var notCompletedFields = [];
		var oMapCommentsInfo   = [];
		var oMapBookmarksInfo  = [];

		var oPrevNumIdInfo = {};

		for (var nElm = 0; nElm < aContent.length; nElm++)
		{
			switch (aContent[nElm].type)
			{
				case "paragraph":
					oFootEndnote.AddToContent(oFootEndnote.Content.length, this.ParagraphFromJSON(aContent[nElm], oFootEndnote, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo, oPrevNumIdInfo), false);
					break;
				case "table":
					oFootEndnote.AddToContent(oFootEndnote.Content.length, this.TableFromJSON(aContent[nElm], oFootEndnote, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo), false);
					break;
				case "blockLvlSdt":
					oFootEndnote.AddToContent(oFootEndnote.Content.length, this.BlockLvlSdtFromJSON(aContent[nElm], oFootEndnote, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo), false);
					break;
			}
		}

		if (oFootEndnote.Content.length > 1)
			// удаляем параграф, который добавляется при создании CDocumentContent
			oFootEndnote.RemoveFromContent(0, 1);

		oFootEndnote.CustomMarkFollow = oParsedFootEndnote.customMarkFollow;
		oFootEndnote.Hint             = oParsedFootEndnote.hint;
		//oFootEndnote.number           = oParsedFootEndnote.number;
		oFootEndnote.ColumnsCount     = oParsedFootEndnote.columnsCount;
		oFootEndnote.SectPr           = oParsedFootEndnote.sectPr ? this.SectPrFromJSON(oParsedFootEndnote.sectPr) : oFootEndnote.SectPr;

		return oFootEndnote;
	};
	ReaderFromJSON.prototype.ReviewInfoFromJSON = function(oParsedReviewInfo)
	{
		var oReviewInfo = new CReviewInfo();

		// move type
		var nMoveType = undefined;
		switch (oParsedReviewInfo.moveType)
		{
			case "noMove":
				nMoveType = Asc.c_oAscRevisionsMove.NoMove;
				break;
			case "moveTo":
				nMoveType = Asc.c_oAscRevisionsMove.MoveTo;
				break;
			case "moveFrom":
				nMoveType = Asc.c_oAscRevisionsMove.MoveFrom;
				break;
		}

		// move type
		var nPrevType = undefined;
		switch (oParsedReviewInfo.prevType)
		{
			case "noMove":
				nPrevType = Asc.c_oAscRevisionsMove.NoMove;
				break;
			case "moveTo":
				nPrevType = Asc.c_oAscRevisionsMove.MoveTo;
				break;
			case "moveFrom":
				nPrevType = Asc.c_oAscRevisionsMove.MoveFrom;
				break;
		}

		oReviewInfo.UserId   = oParsedReviewInfo.userId;
		oReviewInfo.UserName = oParsedReviewInfo.autor;
		oReviewInfo.DateTime = oParsedReviewInfo.date;
		oReviewInfo.MoveType = nMoveType;
		oReviewInfo.PrevType = nPrevType;
		oReviewInfo.PrevInfo = oParsedReviewInfo.prevInfo ? this.ReviewInfoFromJSON(oParsedReviewInfo.prevInfo) : oReviewInfo.PrevInfo;

		return oReviewInfo;
	};
	ReaderFromJSON.prototype.TextPrFromJSON = function(oPr)
	{
		var oTextPr  = new AscCommonWord.CTextPr();

		// alignV
		var nVAlign  = undefined;
		if (oPr.vertAlign)
		{
			switch (oPr.vertAlign)
			{
				case "baseline":
					nVAlign = 0;
					break;
				case "superscript":
					nVAlign = 1;
					break;
				case "subscript":
					nVAlign = 2;
					break;
			}
		}

		// style 
		var oStyle    = oPr.rStyle ? this.StyleFromJSON(oPr.rStyle) : oPr.rStyle;
		var oStyles = private_GetStyles();
		if (oStyle)
		{
			var nExistingStyle = oStyles.GetStyleIdByName(oStyle.Name);
			// если такого стиля нет - добавляем новый
			if (nExistingStyle === null)
			{
				oStyle.Set_BasedOn(null);
				oStyles.Add(oStyle);
			}
			else
			{
				var oExistingStyle = oStyles.Get(nExistingStyle);
				// если стили идентичны, стиль не добавляем
				if (!oStyle.IsEqual(oExistingStyle))
				{
					oStyle.Set_BasedOn(null);
					oStyles.Add(oStyle);
				}
				else
					oStyle = oExistingStyle;
			}
		}

		oTextPr.Bold                  = oPr.b;
		oTextPr.BoldCS                = oPr.bCs;
		oTextPr.Caps                  = oPr.caps;
		oTextPr.Color                 = oPr.color ? new AscCommonWord.CDocumentColor(oPr.color.r, oPr.color.g, oPr.color.b, oPr.color.auto) : oPr.color;
		oTextPr.CS                    = oPr.cs;
		oTextPr.DStrikeout            = oPr.dstrike;
		oTextPr.HighLight             = oPr.highlight === "none" ? -1 : oPr.highlight != undefined ? new AscCommonWord.CDocumentColor(oPr.highlight.r, oPr.highlight.g, oPr.highlight.b, oPr.highlight.auto) : oPr.highlight;
		oTextPr.Italic                = oPr.i;
		oTextPr.ItalicCS              = oPr.iCs;

		oTextPr.Lang.Bidi             = oPr.lang.bidi;
		oTextPr.Lang.EastAsia         = oPr.lang.eastAsia;
		oTextPr.Lang.Val              = oPr.lang.val;

		oTextPr.TextOutline           = oPr.outline ? this.LnFromJSON(oPr.outline) : oTextPr.TextOutline;
		oTextPr.Position              = oPr.position ? private_PtToMM(oPr.position / 2.0) : oPr.position;

		oTextPr.RFonts.Ascii          = oPr.rFonts.ascii;
		oTextPr.RFonts.AsciiTheme     = oPr.rFonts.asciiTheme;
		oTextPr.RFonts.CS             = oPr.rFonts.cs;
		oTextPr.RFonts.CSTheme        = oPr.rFonts.cstheme;
		oTextPr.RFonts.EastAsia       = oPr.rFonts.eastAsia;
		oTextPr.RFonts.EastAsiaTheme  = oPr.rFonts.eastAsiaTheme;
		oTextPr.RFonts.HAnsi          = oPr.rFonts.hAnsi;
		oTextPr.RFonts.HAnsiTheme     = oPr.rFonts.hAnsiTheme;
		oTextPr.RFonts.Hint           = oPr.rFonts.hint;

		oTextPr.FontFamily            = oPr.fontFamily ? {Index: oPr.fontFamily.idx, Name: oPr.fontFamily.name} : oTextPr.FontFamily;
		oTextPr.PrChange              = oPr.rPrChange ? this.TextPrFromJSON(oPr.rPrChange) : oPr.rPrChange;
		oTextPr.RStyle                = oStyle ? oStyle.Id : oTextPr.RStyle;
		oTextPr.RTL                   = oPr.rtl;
		oTextPr.Shd                   = oPr.shd ? this.ShadeFromJSON(oPr.shd) : oPr.shd;
		oTextPr.SmallCaps             = oPr.smallCaps;
		oTextPr.Spacing               = oPr.spacing ? private_Twips2MM(oPr.spacing) : oPr.spacing;
		oTextPr.Strikeout             = oPr.strike;
		oTextPr.FontSize              = oPr.sz ? oPr.sz / 2.0 : oPr.sz;
		oTextPr.FontSizeCS            = oPr.szCs ? oPr.szCs / 2.0 : oPr.szCs;
		oTextPr.Underline             = oPr.u;
		oTextPr.Vanish                = oPr.vanish;
		oTextPr.VertAlign             = nVAlign;
		oTextPr.FontRef               = oPr.fontRef ? this.FontRefFromJSON(oPr.fontRef) : oTextPr.FontRef;
		oTextPr.Unifill               = oPr.uniFill ? this.FillFromJSON(oPr.uniFill) : oPr.uniFill;
		oTextPr.TextFill              = oPr.textFill ? this.FillFromJSON(oPr.textFill) : oPr.TextFill;

		oTextPr.ReviewInfo = oPr.reviewInfo ? this.ReviewInfoFromJSON(oPr.reviewInfo) : oTextPr.ReviewInfo;
		return oTextPr;
	};
	ReaderFromJSON.prototype.ShadeFromJSON = function(oShd) /// To do
	{
		var oShade  = new AscCommonWord.CDocumentShd();
		oShade.Fill = oShd.fill ? new CDocumentColor() : oShd.fill;

		oShade.Value = oShd.val;

		oShade.Color.r    = oShd.color.r;
		oShade.Color.g    = oShd.color.g;
		oShade.Color.b    = oShd.color.b;
		oShade.Color.Auto = oShd.color.auto;

		if (oShade.Fill)
		{
			oShade.Fill.r    = oShd.fill.r;
			oShade.Fill.g    = oShd.fill.g;
			oShade.Fill.b    = oShd.fill.b;
			oShade.Fill.Auto = oShd.fill.auto;	
		}

		oShade.FillRef   = oShd.fillRef    ? this.StyleRefFromJSON(oShd.fillRef) : oShd.fillRef;
		oShade.Unifill   = oShd.themeColor ? this.FillFromJSON(oShd.themeColor)  : oShd.themeColor;
		oShade.themeFill = oShd.themeFill  ? this.FillFromJSON(oShd.themeFill)   : oShd.themeFill;

		return oShade;
	};
	ReaderFromJSON.prototype.ParagraphFromJSON = function(oParsedPara, oParent, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo, oPrevNumIdInfo)
	{
		if (!notCompletedFields)
			notCompletedFields = [];
		if (!oMapCommentsInfo)
			oMapCommentsInfo = {};
		if (!oMapBookmarksInfo)
			oMapBookmarksInfo = {};

		var aContent  = oParsedPara.content;
		var oPr       = oParsedPara.pPr;
		var oDocument = private_GetLogicDocument();
		var oParaPr   = this.ParaPrFromJSON(oPr, oParsedPara.numbering, oPrevNumIdInfo);
		var oPara     = new AscCommonWord.Paragraph(private_GetDrawingDocument(), oParent || oDocument, !oParsedPara.bFromDocument);

		// символ конца параграфа
		oPara.TextPr.Apply_TextPr(this.TextPrFromJSON(oParsedPara.rPr));

		oPara.SetParagraphPr(oParaPr);
		
		// section prop.
		oPr.sectPr && oPara.Set_SectionPr(this.SectPrFromJSON(oPr.sectPr));

		// массив меток переноса
		var arrTrackMoves = [];

		var Comment, CommentStart, CommentEnd, oBookmark, sBookmarkId,
			sBookmarkName, sMoveName, oReviewInfo, oRevisionMove;

		for (var nElm = 0; nElm < aContent.length; nElm++)
		{
			switch (aContent[nElm].type)
			{
				case "run":
				case "mathRun":
				case "presField":
					oPara.AddToContent(oPara.Content.length - 1, this.ParaRunFromJSON(aContent[nElm], oPara, notCompletedFields));
					break;
				case "endRun":
					oPara.RemoveFromContent(oPara.Content.length - 1, 1);
					var oEndRun = this.ParaRunFromJSON(aContent[nElm]);
					oPara.AddToContent(oPara.Content.length, oEndRun);
					oPara.TextPr.Apply_TextPr(oEndRun.Pr);
					break;
				case "paraMath":
					oPara.AddToContent(oPara.Content.length - 1, this.ParaMathFromJSON(aContent[nElm], oPara, notCompletedFields));
					break;
				case "hyperlink":
					oPara.AddToContent(oPara.Content.length - 1, this.HyperlinkFromJSON(aContent[nElm], oPara, notCompletedFields));
					break;
				case "inlineLvlSdt":
					oPara.AddToContent(oPara.Content.length - 1, this.InlineLvlSdtFromJSON(aContent[nElm], oPara, notCompletedFields));
					break;
				case "commentRangeStart":
					var CommentData = new AscCommon.CCommentData();
					CommentData.SetText(aContent[nElm].text);
					CommentData.SetUserName(aContent[nElm].autor);
					Comment = new AscCommon.CComment(oDocument.Comments, CommentData);
					oDocument.Comments.Add(Comment);
					CommentStart = new AscCommon.ParaComment(true, Comment.Get_Id());
					oPara.Add_ToContent(oPara.Content.length - 1, CommentStart);
					oMapCommentsInfo[aContent[nElm].id] = CommentStart.CommentId;
					editor.sync_AddComment(Comment.Get_Id(), CommentData);
					break;
				case "commentRangeEnd":
					Comment = oDocument.Comments.Get_ById(oMapCommentsInfo[aContent[nElm].id]);
					CommentEnd = new AscCommon.ParaComment(false, Comment.Get_Id());
					oPara.Add_ToContent(oPara.Content.length - 1, CommentEnd);
					break;
				case "bookmarkStart":
					sBookmarkName = aContent[nElm].name;
					if (oDocument.BookmarksManager.GetBookmarkByName(sBookmarkName))
							break;

					sBookmarkId = oDocument.BookmarksManager.GetNewBookmarkId();
					oBookmark   = new CParagraphBookmark(true, sBookmarkId, sBookmarkName);
					oPara.Add_ToContent(oPara.Content.length - 1, oBookmark);
					oDocument.BookmarksManager.NeedUpdate = true;

					oMapBookmarksInfo[aContent[nElm].id] = oBookmark;
					break;
				case "bookmarkEnd":
					oBookmark = oMapBookmarksInfo[aContent[nElm].id];
					if (!oBookmark)
						break;
					sBookmarkName = oBookmark.BookmarkName;
					oPara.Add_ToContent(oPara.Content.length - 1, new CParagraphBookmark(false, oBookmark.BookmarkId, sBookmarkName));
					oDocument.BookmarksManager.NeedUpdate = true;
					break;
				case "revisionMove":
					// создаём новый moveId(moveName), мапим к нему значения moveId(moveName) из JSON
					// таким образом задаем соответствия
					sMoveName = undefined;
					if (!this.MoveMap[aContent[nElm].name])
					{
						sMoveName = oDocument.TrackRevisionsManager.GetNewMoveId();
						this.MoveMap[aContent[nElm].name] = sMoveName;
					}

					oReviewInfo   = aContent[nElm].reviewInfo ? this.ReviewInfoFromJSON(aContent[nElm].reviewInfo) : aContent[nElm].reviewInfo;
					oRevisionMove = new CParaRevisionMove(aContent[nElm].start, aContent[nElm].from, this.MoveMap[aContent[nElm].name], oReviewInfo);
					arrTrackMoves.push(oRevisionMove);
					oPara.Add_ToContent(oPara.Content.length - 1, oRevisionMove);
					break;
			}
		}

		for (var nChange = 0; nChange < oParsedPara.changes.length; nChange++)
		{
			if (oParsedPara.changes[nChange].type === "moveMark")
				this.RevisionFromJSON(oParsedPara.changes[nChange], oPara, arrTrackMoves.shift());
			else
				this.RevisionFromJSON(oParsedPara.changes[nChange], oPara);
		}

		return oPara;
	};
	ReaderFromJSON.prototype.ParaPrFromJSON = function(oParsedParaPr, oParsedNumbernig, oPrevNumIdInfo)
	{
		var oParaPr = new AscCommonWord.CParaPr();
		
		oParaPr.ContextualSpacing = oParsedParaPr.contextualSpacing;
		
		// align
		var nJc = undefined;
		switch (oParsedParaPr.jc)
		{
			case "end":
				nJc = AscCommon.align_Right;
				break;
			case "start":
				nJc = AscCommon.align_Left;
				break;
			case "center":
				nJc = AscCommon.align_Center;
				break;
			case "both":
				nJc = AscCommon.align_Justify;
				break;
			case "distribute":
				nJc = AscCommon.align_Distributed;
				break;
		}

		// style 
		var oStyle    = oParsedParaPr.pStyle ? this.StyleFromJSON(oParsedParaPr.pStyle) : oParsedParaPr.pStyle;
		var oStyles   = private_GetStyles();
		if (oStyle)
		{
			var nExistingStyle = oStyles.GetStyleIdByName(oStyle.Name);
			// если такого стиля нет - добавляем новый
			if (nExistingStyle === null)
			{
				oStyle.Set_BasedOn(null);
				oStyles.Add(oStyle);
			}
			else
			{
				var oExistingStyle = oStyles.Get(nExistingStyle);
				// если стили идентичны, стиль не добавляем
				if (!oStyle.IsEqual(oExistingStyle))
				{
					oStyle.Set_BasedOn(null);
					oStyles.Add(oStyle);
				}
				else
					oStyle = oExistingStyle;
			}
		}

		oParaPr.Lvl             = oParsedParaPr.lvl != undefined ? oParsedParaPr.lvl : oParaPr.Lvl;
		oParaPr.FramePr         = oParsedParaPr.framePr != undefined ? this.FramePrFromJSON(oParsedParaPr.framePr) : oParaPr.FramePr;
		oParaPr.Spacing         = oParsedParaPr.spacing != undefined ? this.ParaSpacingFromJSON(oParsedParaPr.spacing) : oParaPr.Spacing;
		oParaPr.Ind.Left        = oParsedParaPr.ind.left != undefined ? private_Twips2MM(oParsedParaPr.ind.left) : oParaPr.Ind.Left;
		oParaPr.Ind.Right       = oParsedParaPr.ind.right != undefined ? private_Twips2MM(oParsedParaPr.ind.right) : oParaPr.Ind.Right;
		oParaPr.Ind.FirstLine   = oParsedParaPr.ind.firstLine != undefined ? private_Twips2MM(oParsedParaPr.ind.firstLine) : oParaPr.Ind.FirstLine;
		oParaPr.Jc              = nJc;
		oParaPr.DefaultTab        = oParsedParaPr.defTabSz != undefined ? private_EMU2MM(oParsedParaPr.defTabSz) : oParsedParaPr.defTabSz;
		oParaPr.KeepLines       = oParsedParaPr.keepLines;
		oParaPr.KeepNext        = oParsedParaPr.keepNext;
		oParaPr.NumPr           = oParsedParaPr.numPr        ? this.NumPrFromJSON(oParsedParaPr.numPr, oParsedNumbernig, oPrevNumIdInfo) : oParaPr.NumPr;
		oParaPr.OutlineLvl      = oParsedParaPr.outlineLvl;
		oParaPr.DefaultRunPr    = oParsedParaPr.defaultRunPr ? this.TextPrFromJSON(oParsedParaPr.defaultRunPr)     : oParaPr.DefaultRunPr;
		oParaPr.Brd.Between     = oParsedParaPr.pBdr.between ? this.DocBorderFromJSON(oParsedParaPr.pBdr.between)  : oParaPr.Brd.Between;
		oParaPr.Brd.Bottom      = oParsedParaPr.pBdr.bottom  ? this.DocBorderFromJSON(oParsedParaPr.pBdr.bottom)   : oParaPr.Brd.Bottom;
		oParaPr.Brd.Left        = oParsedParaPr.pBdr.left    ? this.DocBorderFromJSON(oParsedParaPr.pBdr.left)     : oParaPr.Brd.Left;
		oParaPr.Brd.Right       = oParsedParaPr.pBdr.right   ? this.DocBorderFromJSON(oParsedParaPr.pBdr.right)    : oParaPr.Brd.Right;
		oParaPr.Brd.Top         = oParsedParaPr.pBdr.top     ? this.DocBorderFromJSON(oParsedParaPr.pBdr.top)      : oParaPr.Brd.Top;
		oParaPr.PStyle          = oStyle ? oStyle.Id : oParaPr.PStyle;
		oParaPr.PageBreakBefore = oParsedParaPr.pageBreakBefore;
		oParaPr.Shd             = oParsedParaPr.shd  ? this.ShadeFromJSON(oParsedParaPr.shd) : oParaPr.Shd;
		oParaPr.Tabs            = oParsedParaPr.tabs ? this.TabsFromJSON(oParsedParaPr.tabs) : oParaPr.Tabs;
		oParaPr.WidowControl    = oParsedParaPr.widowControl;
		oParaPr.Bullet          = oParsedParaPr.bullet ? this.BulletFromJSON(oParsedParaPr.bullet) : oParaPr.Bullet;

		return oParaPr;
	};
	ReaderFromJSON.prototype.BulletFromJSON = function(oParsedBullet)
	{
		var oBullet = new AscFormat.CBullet();

		function BulletColorFromJSON(oParsedBulletClr)
		{
			var oBulletColor = new AscFormat.CBulletColor();

			var nBulletColorType = AscFormat.BULLET_TYPE_COLOR_NONE;
			switch (oParsedBulletClr.type)
			{
				case "none":
					nBulletColorType = AscFormat.BULLET_TYPE_COLOR_NONE;
					break;
				case "clrtx":
					nBulletColorType = AscFormat.BULLET_TYPE_COLOR_CLRTX;
					break;
				case "clr":
					nBulletColorType = AscFormat.BULLET_TYPE_COLOR_CLR;
					break;	
			}

			oBulletColor.type = nBulletColorType;
			oBulletColor.UniColor = oParsedBulletClr.color ? this.ColorFromJSON(oParsedBulletClr.color) : oBulletColor.UniColor;

			return oBulletColor;
		}

		function BulletSizeFromJSON(oParsedBulletSz)
		{
			var oBulletSize = new AscFormat.CBulletSize();

			var nBulleSizeType = AscFormat.BULLET_TYPE_SIZE_NONE;
			switch (oParsedBulletSz.type)
			{
				case "none":
					nBulleSizeType = AscFormat.BULLET_TYPE_SIZE_NONE;
					break;
				case "tx":
					nBulleSizeType = AscFormat.BULLET_TYPE_SIZE_TX;
					break;
				case "pct":
					nBulleSizeType = AscFormat.BULLET_TYPE_SIZE_PCT;
					break;	
				case "pts":
					nBulleSizeType = AscFormat.BULLET_TYPE_SIZE_PTS;
					break;	
			}

			oBulletSize.type = nBulleSizeType;
			oBulletSize.val  = oParsedBulletSz.val;

			return oBulletSize;
		}

		function BulletTypeFaceFromJSON(oParsedBulletTpFace)
		{
			var oBulletTypeface = new AscFormat.CBulletTypeface();

			var nBulleTypefaceType = "none";
			switch (oParsedBulletTpFace.type)
			{
				case "none":
					nBulleTypefaceType = AscFormat.BULLET_TYPE_TYPEFACE_NONE;
					break;
				case "tx":
					nBulleTypefaceType = AscFormat.BULLET_TYPE_TYPEFACE_TX;
					break;
				case "bufont":
					nBulleTypefaceType = AscFormat.BULLET_TYPE_TYPEFACE_BUFONT;
					break;	
			}

			oBulletTypeface.type     = nBulleTypefaceType;
			oBulletTypeface.typeface = oParsedBulletTpFace.typeface;

			return oBulletTypeface;
		}

		function BulletTypeFromJSON(oParsedBulletType)
		{
			var oBulletType = new AscFormat.CBulletType();

			var nBulleType = AscFormat.BULLET_TYPE_BULLET_NONE;
			var nAutoNumType = undefined;
			switch (oParsedBulletType.AutoNumType)
			{
				case "alphaLcParenBoth":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_AlphaLcParenBoth;
					break;
				case "alphaLcParenR":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_AlphaLcParenR;
					break;
				case "alphaLcPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_AlphaLcPeriod;
					break;
				case "alphaUcParenBoth":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_AlphaUcParenBoth;
					break;
				case "alphaUcParenR":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_AlphaUcParenR;
					break;
				case "alphaUcPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_AlphaUcPeriod;
					break;
				case "arabic1Minus":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_Arabic1Minus;
					break;
				case "arabic2Minus":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_Arabic2Minus;
					break;
				case "arabicDbPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_ArabicDbPeriod;
					break;
				case "arabicDbPlain":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_ArabicDbPlain;
					break;
				case "arabicParenBoth":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_ArabicParenBoth;
					break;
				case "arabicParenR":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_ArabicParenR;
					break;
				case "arabicPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_ArabicPeriod;
					break;
				case "arabicPlain":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_ArabicPlain;
					break;
				case "circleNumDbPlain":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_CircleNumDbPlain;
					break;
				case "circleNumWdBlackPlain":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_CircleNumWdBlackPlain;
					break;
				case "circleNumWdWhitePlain":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_CircleNumWdWhitePlain;
					break;
				case "ea1ChsPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_Ea1ChsPeriod;
					break;
				case "ea1ChsPlain":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_Ea1ChsPlain;
					break;
				case "ea1ChtPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_Ea1ChtPeriod;
					break;
				case "ea1ChtPlain":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_Ea1ChtPlain;
					break;
				case "ea1JpnChsDbPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_Ea1JpnChsDbPeriod;
					break;
				case "ea1JpnKorPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_Ea1JpnKorPeriod;
					break;
				case "ea1JpnKorPlain":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_Ea1JpnKorPlain;
					break;
				case "hebrew2Minus":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_Hebrew2Minus;
					break;
				case "hindiAlpha1Period":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_HindiAlpha1Period;
					break;
				case "hindiAlphaPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_HindiAlphaPeriod;
					break;
				case "hindiNumParenR":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_HindiNumParenR;
					break;
				case "hindiNumPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_HindiNumPeriod;
					break;
				case "romanLcParenBoth":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_RomanLcParenBoth;
					break;
				case "romanLcParenR":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_RomanLcParenR;
					break;
				case "romanLcPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_RomanLcPeriod;
					break;
				case "romanUcParenBoth":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_RomanUcParenBoth;
					break;
				case "romanUcParenR":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_RomanUcParenR;
					break;
				case "romanUcPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_RomanUcPeriod;
					break;
				case "thaiAlphaParenBoth":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_ThaiAlphaParenBoth;
					break;
				case "thaiAlphaParenR":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_ThaiAlphaParenR;
					break;
				case "thaiAlphaPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_ThaiAlphaPeriod;
					break;
				case "thaiNumParenBoth":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_ThaiNumParenBoth;
					break;
				case "thaiNumParenR":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_ThaiNumParenR;
					break;
				case "thaiNumPeriod":
					nAutoNumType = AscFormat.numbering_presentationnumfrmt_ThaiNumPeriod;
					break;
			}

			switch (oParsedBulletType.type)
			{
				case "none":
					nBulleType = AscFormat.BULLET_TYPE_BULLET_NONE;
					break;
				case "char":
					nBulleType = AscFormat.BULLET_TYPE_BULLET_CHAR;
					break;
				case "autonum":
					nBulleType = AscFormat.BULLET_TYPE_BULLET_AUTONUM;
					break;	
				case "blip":
					nBulleType = AscFormat.BULLET_TYPE_BULLET_BLIP;
					break;
			}

			oBulletType.type        = nBulleType;
			oBulletType.Char        = oParsedBulletType.char;
			oBulletType.AutoNumType = nAutoNumType;
			oBulletType.startAt     = oParsedBulletType.startAt;

			return oBulletType;
		}

		oBullet.bulletColor    = oParsedBullet.bulletColor ? BulletColorFromJSON.call(this, oParsedBullet.bulletColor) : oBullet.bulletColor;
		oBullet.bulletSize     = oParsedBullet.bulletSize ? BulletSizeFromJSON.call(this, oParsedBullet.bulletSize) : oBullet.bulletSize;
		oBullet.bulletTypeface = oParsedBullet.bulletTypeface ? BulletTypeFaceFromJSON.call(this, oParsedBullet.bulletTypeface) : oBullet.bulletTypeface;
		oBullet.bulletType     = oParsedBullet.bulletType ? BulletTypeFromJSON.call(this, oParsedBullet.bulletType) : oBullet.bulletType;
	
		return oBullet;
	};
	ReaderFromJSON.prototype.ParaMathFromJSON = function(oParsedParaMath)
	{
		var oParaMath = new AscCommonWord.ParaMath();
		var nCurPos   = 0;

		var nJc = undefined;
		switch (oParsedParaMath.oMathParaPr.jc)
		{
			case "center":
				nJc = JC_CENTER;
				break;
			case "centerGroup":
				nJc = JC_CENTERGROUP;
				break;
			case "left":
				nJc = JC_LEFT;
				break;
			case "right":
				nJc = JC_RIGHT;
				break;
		}

		nJc !== undefined && oParaMath.Set_Align(nJc);

		function MathContentFromJSON(oParsedMathContent)
		{
			var aContent = [];

			for (var nElm = 0; nElm < oParsedMathContent.length; nElm++)
			{
				switch (oParsedMathContent[nElm].type)
				{
					case "mathRun":
					case "run":
						aContent.push(this.ParaRunFromJSON(oParsedMathContent[nElm]));
						break;
					case "fraction":
						aContent.push(FractionFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "superScript":
					case "supScript":
						aContent.push(DegreeFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "subSupScript":
					case "preSubSupScript":
						aContent.push(SupSubDegreeFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "radSquare":
					case "radDegree":
						aContent.push(RadicalFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "nary":
						aContent.push(NaryFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "eqArray":
						aContent.push(EqArrayFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "delimiter":
						aContent.push(DelimiterFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "mathFunc":
						aContent.push(MathFuncFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "accent":
						aContent.push(AccentFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "groupChr":
						aContent.push(GroupCharacterFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "borderBox":
						aContent.push(BorderBoxFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "box":
						aContent.push(BoxFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "bar":
						aContent.push(BarFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "limUpp":
					case "limLow":
						aContent.push(LimitFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					case "matrix":
						aContent.push(MathMatrixFromJSON.call(this, oParsedMathContent[nElm]));
						break;
					
				}
			}

			return aContent;
		}
		function FractionFromJSON(oParsedFraction)
		{
			// fraction type
			var nFracType;
			switch (oParsedFraction.fPr.type)
			{
				case "bar":
					nFracType = BAR_FRACTION;
					break;
				case "skw":
					nFracType = SKEWED_FRACTION;
					break;
				case "lin":
					nFracType = LINEAR_FRACTION;
					break;
				default:
					nFracType = BAR_FRACTION;
					break;
			}

			var oFracPr   = this.TextPrFromJSON(oParsedFraction.fPr.ctrlPr);
			var oFraction = new AscCommonWord.CFraction({ctrPrp: oFracPr, type: nFracType});

			var oDenMathContent = oFraction.getDenominatorMathContent();
			var aDenMathContent = MathContentFromJSON.call(this, oParsedFraction.den.content);
			var nCurPos = 0;
			for (var nElm = 0; nElm < aDenMathContent.length; nElm++)
			{
				oDenMathContent.Internal_Content_Add(nCurPos, aDenMathContent[nElm], false);
				nCurPos++;
			}

			var oNumMathContent = oFraction.getNumeratorMathContent();
			var aNumMathContent = MathContentFromJSON.call(this, oParsedFraction.num.content);
			nCurPos = 0;
			for (nElm = 0; nElm < aNumMathContent.length; nElm++)
			{
				oNumMathContent.Internal_Content_Add(nCurPos, aNumMathContent[nElm], false);
				nCurPos++;
			}

			oFraction.Denominator.setCtrPrp(this.TextPrFromJSON(oParsedFraction.den.ctrlPr));
			oParsedFraction.den.argPr.argSize != undefined && oFraction.Denominator.ArgSize.SetValue(oParsedFraction.den.argPr.argSize);

			oFraction.Numerator.setCtrPrp(this.TextPrFromJSON(oParsedFraction.num.ctrlPr));
			oParsedFraction.num.argPr.argSize != undefined && oFraction.Numerator.ArgSize.SetValue(oParsedFraction.num.argPr.argSize);
		
			return oFraction;
		}
		function DegreeFromJSON(oParsedDegree)
		{
			var nType = oParsedDegree.type === "superScript" ? DEGREE_SUPERSCRIPT : DEGREE_SUBSCRIPT;

			var oParsedCtrlPr = nType === DEGREE_SUPERSCRIPT ? oParsedDegree.sSupPr.ctrlPr : oParsedDegree.sSubPr.ctrlPr;

			var oCtrlPr = this.TextPrFromJSON(oParsedCtrlPr);
			var oDegree = new AscCommonWord.CDegree({ctrPrp: oCtrlPr, type: nType});

			var oBaseContent = oDegree.getBase();
			var aBaseContent = MathContentFromJSON.call(this, oParsedDegree.e);
			var nCurPos      = 0;
			for (var nElm = 0; nElm < aBaseContent.length; nElm++)
			{
				oBaseContent.Internal_Content_Add(nCurPos, aBaseContent[nElm]);
				nCurPos++;
			}

			var oIterContent = oDegree.getUpperIterator();
			var aIterContent = MathContentFromJSON.call(this, oParsedDegree.sup || oParsedDegree.sub);
			nCurPos          = 0;
			for (nElm = 0; nElm < aIterContent.length; nElm++)
			{
				oIterContent.Internal_Content_Add(nCurPos, aIterContent[nElm]);
				nCurPos++;
			}

			return oDegree;
		}
		function SupSubDegreeFromJSON(oParsedDegree)
		{
			var nType = oParsedDegree.type === "subSupScript" ? DEGREE_SubSup : DEGREE_PreSubSup;

			var oParsedCtrlPr = nType === DEGREE_SubSup ? oParsedDegree.sSubSupPr.ctrlPr : oParsedDegree.sPrePr.ctrlPr;
			var oCtrlPr = this.TextPrFromJSON(oParsedCtrlPr);
			var oSupSubDegree = new AscCommonWord.CDegreeSubSup({ctrPrp: oCtrlPr, type: nType});

			var oBaseContent = oSupSubDegree.getBase();
			var aBaseContent = MathContentFromJSON.call(this, oParsedDegree.e);
			var nCurPos      = 0;
			for (var nElm = 0; nElm < aBaseContent.length; nElm++)
			{
				oBaseContent.Internal_Content_Add(nCurPos, aBaseContent[nElm]);
				nCurPos++;
			}

			var oUpContent = oSupSubDegree.getUpperIterator();
			var aUpContent = MathContentFromJSON.call(this, oParsedDegree.sup);
			nCurPos          = 0;
			for (nElm = 0; nElm < aUpContent.length; nElm++)
			{
				oUpContent.Internal_Content_Add(nCurPos, aUpContent[nElm]);
				nCurPos++;
			}

			var oDnContent = oSupSubDegree.getLowerIterator();
			var aDnContent = MathContentFromJSON.call(this, oParsedDegree.sub);
			nCurPos          = 0;
			for (nElm = 0; nElm < aDnContent.length; nElm++)
			{
				oDnContent.Internal_Content_Add(nCurPos, aDnContent[nElm]);
				nCurPos++;
			}

			return oSupSubDegree;
		}
		function RadicalFromJSON(oParsedRad)
		{
			var nType   = oParsedRad.type === "radSquare" ? SQUARE_RADICAL : DEGREE_RADICAL;
			var oCtrlPr = this.TextPrFromJSON(oParsedRad.radPr.ctrlPr);

			var oRadical = new AscCommonWord.CRadical({ctrPrp: oCtrlPr, degHide: oParsedRad.radPr.degHide, type: nType});

			var oBaseContent = oRadical.getBase();
			var aBaseContent = MathContentFromJSON.call(this, oParsedRad.e);
			var nCurPos      = 0;
			for (var nElm = 0; nElm < aBaseContent.length; nElm++)
			{
				oBaseContent.Internal_Content_Add(nCurPos, aBaseContent[nElm]);
				nCurPos++;
			}

			var oDegreeContent = oRadical.getDegree();
			var aDegreeContent = MathContentFromJSON.call(this, oParsedRad.deg);
			nCurPos            = 0;
			for (nElm = 0; nElm < aDegreeContent.length; nElm++)
			{
				oDegreeContent.Internal_Content_Add(nCurPos, aDegreeContent[nElm]);
				nCurPos++;
			}

			return oRadical;
		}
		function NaryFromJSON(oParsedNary)
		{
			var oCtrlPr = this.TextPrFromJSON(oParsedNary.naryPr.ctrlPr);
			var oPr = {
				chr: oParsedNary.naryPr.chr,
				ctrPrp: oCtrlPr,
				limLoc: oParsedNary.naryPr.limLoc === "undOvr" ? NARY_UndOvr : NARY_SubSup,
				subHide: oParsedNary.naryPr.subHide,
				supHide: oParsedNary.naryPr.supHide
			};

			var oNary = new AscCommonWord.CNary(oPr);

			var oBaseContent = oNary.getBase();
			var aBaseContent = MathContentFromJSON.call(this, oParsedNary.e);
			var nCurPos      = 0;
			for (var nElm = 0; nElm < aBaseContent.length; nElm++)
			{
				oBaseContent.Internal_Content_Add(nCurPos, aBaseContent[nElm]);
				nCurPos++;
			}

			var oSubContent = oNary.getSubMathContent();
			var aSubContent = MathContentFromJSON.call(this, oParsedNary.sub);
			nCurPos            = 0;
			for (nElm = 0; nElm < aSubContent.length; nElm++)
			{
				oSubContent.Internal_Content_Add(nCurPos, aSubContent[nElm]);
				nCurPos++;
			}

			var oSupContent = oNary.getSupMathContent();
			var aSupContent = MathContentFromJSON.call(this, oParsedNary.sup);
			nCurPos            = 0;
			for (nElm = 0; nElm < aSupContent.length; nElm++)
			{
				oSupContent.Internal_Content_Add(nCurPos, aSupContent[nElm]);
				nCurPos++;
			}

			return oNary;
		}
		function EqArrayFromJSON(oParsedEqArray)
		{
			var oCtrlPr = this.TextPrFromJSON(oParsedEqArray.eqArrPr.ctrlPr);

			var nJcType = undefined;
			switch (oParsedEqArray.eqArrPr.baseJc)
			{
				case "center":
					nJcType = BASEJC_CENTER;
					break;
				case "top":
					nJcType = BASEJC_TOP;
					break;
				case "bottom":
					nJcType = BASEJC_BOTTOM;
					break;
				case "inline":
					nJcType = BASEJC_INLINE;
					break;
				case "inside":
					nJcType = BASEJC_INSIDE;
					break;
				case "outside":
					nJcType = BASEJC_OUTSIDE;
					break;
			}

			var oPr = {
				ctrPrp:  oCtrlPr,
				baseJc:  nJcType,
				maxDist: oParsedEqArray.eqArrPr.maxDist,
				objDist: oParsedEqArray.eqArrPr.objDist,
				rSp:     oParsedEqArray.eqArrPr.rSp,
				rSpRule: oParsedEqArray.eqArrPr.rSpRule,
				row:     oParsedEqArray.eqArrPr.row,
			};

			var oEqArray = new AscCommonWord.CEqArray(oPr);

			for (var Index = 0; Index < oPr.row; Index++)
			{
				var oMathContent = oEqArray.getElementMathContent(Index);
				var aMathContent = MathContentFromJSON.call(this, oParsedEqArray.e[Index]);
				var nCurPos      = 0;
				for (var nElm = 0; nElm < aMathContent.length; nElm++)
				{
					oMathContent.Internal_Content_Add(nCurPos, aMathContent[nElm]);
					nCurPos++;
				}
			}

			return oEqArray;
		}
		function DelimiterFromJSON(oParsedDelimiter)
		{
			var oCtrlPr  = this.TextPrFromJSON(oParsedDelimiter.dPr.ctrlPr);
			var nShpType = oParsedDelimiter.dPr.shp === "centered" ? DELIMITER_SHAPE_CENTERED : DELIMITER_SHAPE_MATCH;

			var oPr = {
				ctrPrp:  oCtrlPr,
				begChr:  oParsedDelimiter.dPr.begChr,
				endChr:  oParsedDelimiter.dPr.endChr,
				grow:    oParsedDelimiter.dPr.grow,
				sepChr:  oParsedDelimiter.dPr.sepChr,
				shp:     nShpType,
				column:  oParsedDelimiter.e.length
			};

			var oDelimiter = new AscCommonWord.CDelimiter(oPr);

			for (var Index = 0; Index < oPr.column; Index++)
			{
				var oMathContent = oDelimiter.getElementMathContent(Index);
				var aMathContent = MathContentFromJSON.call(this, oParsedDelimiter.e[Index]);
				var nCurPos      = 0;
				for (var nElm = 0; nElm < aMathContent.length; nElm++)
				{
					oMathContent.Internal_Content_Add(nCurPos, aMathContent[nElm]);
					nCurPos++;
				}
			}

			return oDelimiter;
		}
		function MathFuncFromJSON(oParsedMathFunc)
		{
			var oCtrlPr  = this.TextPrFromJSON(oParsedMathFunc.funcPr.ctrlPr);

			var oPr = {
				ctrPrp:  oCtrlPr,
			};

			var oMathFunc = new AscCommonWord.CMathFunc(oPr);

			var oMathNameContent = oMathFunc.getFName();
			var aMathNameContent = MathContentFromJSON.call(this, oParsedMathFunc.fName);
			var nCurPos      = 0;
			for (var nElm = 0; nElm < aMathNameContent.length; nElm++)
			{
				oMathNameContent.Internal_Content_Add(nCurPos, aMathNameContent[nElm]);
				nCurPos++;
			}

			var oMathArgContent = oMathFunc.getArgument();
			var aMathArgContent = MathContentFromJSON.call(this, oParsedMathFunc.e);
			nCurPos      = 0;
			for (nElm = 0; nElm < aMathArgContent.length; nElm++)
			{
				oMathArgContent.Internal_Content_Add(nCurPos, aMathArgContent[nElm]);
				nCurPos++;
			}

			return oMathFunc;
		}
		function AccentFromJSON(oParsedAccent)
		{
			var oCtrlPr  = this.TextPrFromJSON(oParsedAccent.accPr.ctrlPr);

			var oPr = {
				ctrPrp: oCtrlPr,
				chr:    oParsedAccent.accPr.chr
			};

			var oAccent = new CAccent(oPr);

			var oMathContent = oAccent.getBase();
			var aMathContent = MathContentFromJSON.call(this, oParsedAccent.e);
			var nCurPos      = 0;
			for (var nElm = 0; nElm < aMathContent.length; nElm++)
			{
				oMathContent.Internal_Content_Add(nCurPos, aMathContent[nElm]);
				nCurPos++;
			}

			return oAccent;
		}
		function GroupCharacterFromJSON(oParsedGrpChr)
		{
			var nPos     = oParsedGrpChr.groupChrPr.pos === "bot" ? LOCATION_BOT : LOCATION_TOP;
			var nVertJc  = oParsedGrpChr.groupChrPr.vertJc === "top" ? VJUST_TOP : VJUST_BOT;
			var oCtrlPr  = this.TextPrFromJSON(oParsedGrpChr.groupChrPr.ctrlPr);

			var oPr = {
				ctrPrp: oCtrlPr,
				chr:    oParsedGrpChr.groupChrPr.chr,
				pos:    nPos,
				vertJc: nVertJc
			};

			var oGroupCharacter = new CGroupCharacter(oPr);

			var oMathContent = oGroupCharacter.getBase();
			var aMathContent = MathContentFromJSON.call(this, oParsedGrpChr.e);
			var nCurPos      = 0;
			for (var nElm = 0; nElm < aMathContent.length; nElm++)
			{
				oMathContent.Internal_Content_Add(nCurPos, aMathContent[nElm]);
				nCurPos++;
			}

			return oGroupCharacter;
		}
		function BorderBoxFromJSON(oParsedBorderBox)
		{
			var oCtrlPr  = this.TextPrFromJSON(oParsedBorderBox.borderBoxPr.ctrlPr);
			var oPr = {
				ctrPrp:     oCtrlPr,
				hideBot:    oParsedBorderBox.borderBoxPr.hideBot,
				hideLeft:   oParsedBorderBox.borderBoxPr.hideLeft,
				hideRight:  oParsedBorderBox.borderBoxPr.hideRight,
				hideTop:    oParsedBorderBox.borderBoxPr.hideTop,
				strikeBLTR: oParsedBorderBox.borderBoxPr.strikeBLTR,
				strikeH:    oParsedBorderBox.borderBoxPr.strikeH,
				strikeTLBR: oParsedBorderBox.borderBoxPr.strikeTLBR,
				strikeV:    oParsedBorderBox.borderBoxPr.strikeV
			};

			var oBorderBox = new CBorderBox(oPr);

			var oMathContent = oBorderBox.getBase();
			var aMathContent = MathContentFromJSON.call(this, oParsedBorderBox.e);
			var nCurPos      = 0;
			for (var nElm = 0; nElm < aMathContent.length; nElm++)
			{
				oMathContent.Internal_Content_Add(nCurPos, aMathContent[nElm]);
				nCurPos++;
			}

			return oBorderBox;
		}
		function BoxFromJSON(oParsedBox)
		{
			var oCtrlPr = this.TextPrFromJSON(oParsedBox.boxPr.ctrlPr);

			var oPr = {
				ctrPrp:  oCtrlPr,
				brk:     oParsedBox.boxPr.brk,
				aln:     oParsedBox.boxPr.aln,
				diff:    oParsedBox.boxPr.diff,
				noBreak: oParsedBox.boxPr.noBreak,
				opEmu:   oParsedBox.boxPr.opEmu
			};

			var oBox = new CBox(oPr);

			var oMathContent = oBox.getBase();
			var aMathContent = MathContentFromJSON.call(this, oParsedBox.e);
			var nCurPos      = 0;
			for (var nElm = 0; nElm < aMathContent.length; nElm++)
			{
				oMathContent.Internal_Content_Add(nCurPos, aMathContent[nElm]);
				nCurPos++;
			}

			return oBox;
		}
		function BarFromJSON(oParsedBar)
		{
			var oCtrlPr = this.TextPrFromJSON(oParsedBar.barPr.ctrlPr);

			var nPos = undefined;
			switch (oParsedBar.barPr.pos)
			{
				case "bot":
					nPos = LOCATION_BOT;
					break;
				case "top":
					nPos = LOCATION_TOP;
					break;
			}

			var oPr = {
				ctrPrp: oCtrlPr,
				pos:   nPos
			};

			var oBar = new CBar(oPr);

			var oMathContent = oBar.getBase();
			var aMathContent = MathContentFromJSON.call(this, oParsedBar.e);
			var nCurPos      = 0;
			for (var nElm = 0; nElm < aMathContent.length; nElm++)
			{
				oMathContent.Internal_Content_Add(nCurPos, aMathContent[nElm]);
				nCurPos++;
			}

			return oBar;
		}
		function LimitFromJSON(oParsedLimit)
		{
			var nLimitType = oParsedLimit.type === "limLow" ? LIMIT_LOW : LIMIT_UP;
			var oCtrlPr    = nLimitType === LIMIT_LOW ? this.TextPrFromJSON(oParsedLimit.limLowPr.ctrlPr) : this.TextPrFromJSON(oParsedLimit.limUppPr.ctrlPr);

			var oPr = {
				ctrPrp: oCtrlPr,
				type:   nLimitType
			};

			var oLimit = new CLimit(oPr);

			var oMathFNameContent = oLimit.getFName();
			var aMathFNameContent = MathContentFromJSON.call(this, oParsedLimit.e);
			var nCurPos           = 0;
			for (var nElm = 0; nElm < aMathFNameContent.length; nElm++)
			{
				oMathFNameContent.Internal_Content_Add(nCurPos, aMathFNameContent[nElm]);
				nCurPos++;
			}

			var oMathIterContent = oLimit.getIterator();
			var aMathIterContent = MathContentFromJSON.call(this, oParsedLimit.limit);
			nCurPos              = 0;
			for (nElm = 0; nElm < aMathIterContent.length; nElm++)
			{
				oMathIterContent.Internal_Content_Add(nCurPos, aMathIterContent[nElm]);
				nCurPos++;
			}

			return oLimit;
		}
		function MathMatrixFromJSON(oParsedMatrix)
		{
			var oCtrlPr    = this.TextPrFromJSON(oParsedMatrix.mPr.ctrlPr);

			var nJcType = undefined;
			switch (oParsedMatrix.mPr.baseJc)
			{
				case "center":
					nJcType = BASEJC_CENTER;
					break;
				case "top":
					nJcType = BASEJC_TOP;
					break;
				case "bottom":
					nJcType = BASEJC_BOTTOM;
					break;
				case "inline":
					nJcType = BASEJC_INLINE;
					break;
				case "inside":
					nJcType = BASEJC_INSIDE;
					break;
				case "outside":
					nJcType = BASEJC_OUTSIDE;
					break;
			}

			var oPr = {
				ctrPrp:  oCtrlPr,
				baseJc:  nJcType,
				cGp:     oParsedMatrix.mPr.cGp,
				cGpRule: oParsedMatrix.mPr.cGpRule,
				cSp:     oParsedMatrix.mPr.cSp,
				plcHide: oParsedMatrix.mPr.plcHide,
				rSp:     oParsedMatrix.mPr.rSp,
				rSpRule: oParsedMatrix.mPr.rSpRule,
				mcs:     oParsedMatrix.mPr.mcs,
				row:     oParsedMatrix.mr.length
			};

			var oMatrix    = new CMathMatrix(oPr);
			var nRowsCount = oMatrix.getRowsCount();
    		var nColsCount = oMatrix.getColsCount();

			for (var nRow = 0; nRow < nRowsCount; nRow++)
			{
				for (var nCell = 0; nCell < nColsCount; nCell++)
				{
					var oMathContent = oMatrix.getContentElement(nRow, nCell);
					var aMathContent = MathContentFromJSON.call(this, oParsedMatrix.mr[nRow][nCell]);
					var nCurPos           = 0;
					for (var nElm = 0; nElm < aMathContent.length; nElm++)
					{
						oMathContent.Internal_Content_Add(nCurPos, aMathContent[nElm]);
						nCurPos++;
					}
				}
			}

			return oMatrix;
		}

		var aMathContent = MathContentFromJSON.call(this, oParsedParaMath.content);
		for (var nElm = 0; nElm < aMathContent.length; nElm++)
		{
			oParaMath.Root.Internal_Content_Add(nCurPos, aMathContent[nElm]);
			nCurPos++;
		}
		
		return oParaMath;
	};
	ReaderFromJSON.prototype.RevisionFromJSON = function(oParsedChange, oElement, Value)
	{
		var oChange   = new CRevisionsChange();
		var oDocument = private_GetLogicDocument();

		// change type
		var nChangeType = undefined;
		switch (oParsedChange.type)
		{
			case "unknown":
				nChangeType = Asc.c_oAscRevisionsChangeType.Unknown;
				break;
			case "textAdd":
				nChangeType = Asc.c_oAscRevisionsChangeType.TextAdd;
				break;
			case "textRem":
				nChangeType = Asc.c_oAscRevisionsChangeType.TextRem;
				break;
			case "paraAdd":
				nChangeType = Asc.c_oAscRevisionsChangeType.ParaAdd;
				break;
			case "paraRem":
				nChangeType = Asc.c_oAscRevisionsChangeType.ParaRem;
				break;
			case "textPr":
				nChangeType = Asc.c_oAscRevisionsChangeType.TextPr;
				break;
			case "paraPr":
				nChangeType = Asc.c_oAscRevisionsChangeType.ParaPr;
				break;
			case "tablePr":
				nChangeType = Asc.c_oAscRevisionsChangeType.TablePr;
				break;
			case "rowsAdd":
				nChangeType = Asc.c_oAscRevisionsChangeType.RowsAdd;
				break;
			case "rowsRem":
				nChangeType = Asc.c_oAscRevisionsChangeType.RowsRem;
				break;
			case "moveMark":
				nChangeType = Asc.c_oAscRevisionsChangeType.MoveMark;
				break;
			case "moveMarkRemove":
				nChangeType = Asc.c_oAscRevisionsChangeType.MoveMarkRemove;
				break;
		}

		// move type
		var nMoveType = undefined;
		switch (oParsedChange.moveType)
		{
			case "noMove":
				nMoveType = Asc.c_oAscRevisionsMove.NoMove;
				break;
			case "moveTo":
				nMoveType = Asc.c_oAscRevisionsMove.MoveTo;
				break;
			case "moveFrom":
				nMoveType = Asc.c_oAscRevisionsMove.MoveFrom;
				break;
		}

		// value
		var changeValue = Value;
		if (!changeValue)
		{
			if (oParsedChange.value.type && oParsedChange.value.type === "textPr")
				changeValue = this.TextPrFromJSON(oParsedChange.value);
			else
				changeValue = oParsedChange.value;
		}
		
		// start pos
		var startPos;
		if (oParsedChange.start.data)
		{
			startPos              = new CParagraphContentPos()
			startPos.Data         = oParsedChange.start.data;
			startPos.Depth        = oParsedChange.start.depth;
			startPos.bPlaceholder = oParsedChange.start.bPlaceholder;
		}
		else
			startPos = oParsedChange.start;
		

		// end pos
		var endPos;
		if (oParsedChange.end.data)
		{
			endPos              = new CParagraphContentPos();
			endPos.Data         = oParsedChange.end.data;
			endPos.Depth        = oParsedChange.end.depth;
			endPos.bPlaceholder = oParsedChange.end.bPlaceholder;
		}
		else
			endPos = oParsedChange.end;
		
		// color
		// oChange.UserColor.a = oParsedChange.userColor.a;
		// oChange.UserColor.g = oParsedChange.userColor.g;
		// oChange.UserColor.b = oParsedChange.userColor.b;
		// oChange.UserColor.r = oParsedChange.userColor.r;

		oChange.SetType(nChangeType);
		oChange.SetElement(oElement);
		oChange.put_Value(changeValue);
		oChange.put_StartPos(startPos);
		oChange.put_EndPos(endPos);
		oChange.SetMoveType(nMoveType);
		oChange.SetUserId(oParsedChange.userId);
		oChange.SetUserName(oParsedChange.autor);
		oChange.SetDateTime(oParsedChange.date);

		oDocument.TrackRevisionsManager.AddChange(oElement.GetId(), oChange);
	};
	ReaderFromJSON.prototype.SectPrFromJSON = function(oParsedSectPr)
	{
		var oSectPr = new AscCommonWord.CSectionPr(private_GetLogicDocument());

		var nSectionType = undefined;
		switch(oParsedSectPr.type)
		{
			case "nextPage":
				nSectionType = Asc.c_oAscSectionBreakType.NextPage;
				break;
			case "oddPage":
				nSectionType = Asc.c_oAscSectionBreakType.OddPage;
				break;
			case "evenPage":
				nSectionType = Asc.c_oAscSectionBreakType.EvenPage;
				break;
			case "continuous":
				nSectionType = Asc.c_oAscSectionBreakType.Continuous;
				break;
			case "nextColumn":
				nSectionType = Asc.c_oAscSectionBreakType.Column;
				break;
		}

		oSectPr.FooterFirst   = oParsedSectPr.footerReference.first   ? this.FooterFromJSON(oParsedSectPr.footerReference.first)   : oSectPr.FooterFirst;
		oSectPr.FooterDefault = oParsedSectPr.footerReference.default ? this.FooterFromJSON(oParsedSectPr.footerReference.default) : oSectPr.FooterDefault;
		oSectPr.FooterEven    = oParsedSectPr.footerReference.even    ? this.FooterFromJSON(oParsedSectPr.footerReference.even)    : oSectPr.FooterEven;

		oSectPr.HeaderFirst   = oParsedSectPr.headerReference.first   ? this.HeaderFromJSON(oParsedSectPr.headerReference.first)   : oSectPr.HeaderFirst;
		oSectPr.HeaderDefault = oParsedSectPr.headerReference.default ? this.HeaderFromJSON(oParsedSectPr.headerReference.default) : oSectPr.HeaderDefault;
		oSectPr.HeaderEven    = oParsedSectPr.headerReference.even    ? this.HeaderFromJSON(oParsedSectPr.headerReference.even)    : oSectPr.HeaderEven;

		oSectPr.Columns       = this.SectionColumnsFromJSON(oParsedSectPr.cols, oSectPr);
		oSectPr.EndnotePr     = this.EndnotePrFromJSON(oParsedSectPr.endnotePr);
		oSectPr.FootnotePr    = this.FootnotePrFromJSON(oParsedSectPr.footnotePr);
		oSectPr.LnNumType     = oParsedSectPr.lnNumType ? this.LnNumTypeFromJSON(oParsedSectPr.lnNumType) : oSectPr.LnNumType;
		oSectPr.Borders       = this.PageBordersFromJSON(oParsedSectPr.pgBorders);
		oSectPr.PageMargins   = this.PageMarginsFromJSON(oParsedSectPr.pgMar);
		oSectPr.PageSize      = this.PageSizeFromJSON(oParsedSectPr.pgSz);
		oSectPr.GutterRTL     = oParsedSectPr.rtlGutter;
		oSectPr.TitlePage     = oParsedSectPr.titlePg;
		oSectPr.Type          = nSectionType;

		return oSectPr;
	};
	ReaderFromJSON.prototype.PageMarginsFromJSON = function(oParsedMargins)
	{
		var oPageMargins = new CSectionPageMargins();

		oPageMargins.Bottom = private_Twips2MM(oParsedMargins.bottom);
		oPageMargins.Footer = private_Twips2MM(oParsedMargins.footer);
		oPageMargins.Gutter = private_Twips2MM(oParsedMargins.gutter);
		oPageMargins.Header = private_Twips2MM(oParsedMargins.header);
		oPageMargins.Left   = private_Twips2MM(oParsedMargins.left);
		oPageMargins.Right  = private_Twips2MM(oParsedMargins.right);
		oPageMargins.Top    = private_Twips2MM(oParsedMargins.top);

		return oPageMargins;
	};
	ReaderFromJSON.prototype.PageSizeFromJSON = function(oParsedPgSize)
	{
		var oPageSize = new CSectionPageSize();

		var nOrientType = undefined;
		switch(oParsedPgSize.orinet)
		{
			case "portrait":
				nOrientType = Asc.c_oAscPageOrientation.PagePortrait;
				break;
			case "landscape":
				nOrientType = Asc.c_oAscPageOrientation.PageLandscape;
				break;
		}

		oPageSize.H = private_Twips2MM(oParsedPgSize.h);
		oPageSize.W = private_Twips2MM(oParsedPgSize.w);

		return oPageSize;
	};
	ReaderFromJSON.prototype.PageBordersFromJSON = function(oParsedPgBorders)
	{
		var oPageBorders = new CSectionBorders();

		var nDisplayType = undefined;
		switch(oParsedPgBorders.display)
		{
			case "allPages":
				nDisplayType = section_borders_DisplayAllPages;
				break;
			case "firstPage":
				nDisplayType = section_borders_DisplayFirstPage;
				break;
			case "notFirstPage":
				nDisplayType = section_borders_DisplayNotFirstPage;
				break;
		}

		oPageBorders.Bottom = this.DocBorderFromJSON(oParsedPgBorders.bottom);
		oPageBorders.Left   = this.DocBorderFromJSON(oParsedPgBorders.left);
		oPageBorders.Right  = this.DocBorderFromJSON(oParsedPgBorders.right);
		oPageBorders.Top    = this.DocBorderFromJSON(oParsedPgBorders.top);

		oPageBorders.Display    = nDisplayType;
		oPageBorders.OffsetFrom = private_Twips2MM(oParsedPgBorders.offsetFrom);
		oPageBorders.ZOrder     = oParsedPgBorders.zOrder === "front" ? section_borders_ZOrderFront : section_borders_ZOrderBack;

		return oPageBorders;
	};
	ReaderFromJSON.prototype.LnNumTypeFromJSON = function(oParsedLnNumType)
	{
		var nRestartType = undefined;
		switch(oParsedLnNumType.restart)
		{
			case "continuous":
				nRestartType = Asc.c_oAscLineNumberRestartType.Continuous;
				break;
			case "newPage":
				nRestartType = Asc.c_oAscLineNumberRestartType.NewPage;
				break;
			case "newSection":
				nRestartType = Asc.c_oAscLineNumberRestartType.NewSection;
				break;
		}

		return new Asc.CSectionLnNumType(oParsedLnNumType.countBy, private_Twips2MM(oParsedLnNumType.distance), oParsedLnNumType.start, nRestartType);
	};
	ReaderFromJSON.prototype.EndnotePrFromJSON = function(oParsedPr)
	{
		var oEndnotePr = new CFootnotePr();

		var nNumRestart = undefined;
		switch(oParsedPr.numRestart)
		{
			case "continuous":
				nNumRestart = section_footnote_RestartContinuous;
				break;
			case "eachPage":
				nNumRestart = section_footnote_RestartEachPage;
				break;
			case "eachSect":
				nNumRestart = section_footnote_RestartEachSect;
				break;
		}

		var nEndPos = undefined;
		switch(oParsedPr.pos)
		{
			case "docEnd":
				nEndPos = Asc.c_oAscEndnotePos.DocEnd;
				break;
			case "sectEnd":
				nEndPos = Asc.c_oAscEndnotePos.SectEnd;
				break;
		}

		oEndnotePr.NumFormat  = From_XML_c_oAscNumberingFormat(oParsedPr.numFmt);
		oEndnotePr.NumRestart = nNumRestart;
		oEndnotePr.NumStart   = oParsedPr.numStart;
		oEndnotePr.Pos        = nEndPos;

		return oEndnotePr;
	};
	ReaderFromJSON.prototype.FootnotePrFromJSON = function(oParsedPr)
	{
		var oFootnotePr = new CFootnotePr();

		var nNumRestart = undefined;
		switch(oParsedPr.numRestart)
		{
			case "continuous":
				nNumRestart = section_footnote_RestartContinuous;
				break;
			case "eachPage":
				nNumRestart = section_footnote_RestartEachPage;
				break;
			case "eachSect":
				nNumRestart = section_footnote_RestartEachSect;
				break;
		}

		var nEndPos = undefined;
		switch(oParsedPr.pos)
		{
			case "beneathText":
				nEndPos = Asc.c_oAscFootnotePos.BeneathText;
				break;
			case "docEnd":
				nEndPos = Asc.c_oAscFootnotePos.DocEnd;
				break;
			case "pgBottom":
				nEndPos = Asc.c_oAscFootnotePos.PageBottom;
				break;
			case "sectEnd":
				nEndPos = Asc.c_oAscFootnotePos.SectEnd;
				break;
		}

		oFootnotePr.NumFormat  = From_XML_c_oAscNumberingFormat(oParsedPr.numFmt);
		oFootnotePr.NumRestart = nNumRestart;
		oFootnotePr.NumStart   = oParsedPr.numStart;
		oFootnotePr.Pos        = nEndPos;

		return oFootnotePr;
	};
	ReaderFromJSON.prototype.SectionColumnsFromJSON = function(oParsedSectCols, oParentSectPr)
	{
		var oSectionCols = new CSectionColumns(oParentSectPr);

		for (var nCol = 0; nCol < oParsedSectCols.col.length; nCol++)
			oSectionCols.Cols.push(this.SectionColFromJSON(oParsedSectCols.col[nCol]));

		oSectionCols.EqualWidth = oParsedSectCols.equalWidth;
		oSectionCols.Num        = oParsedSectCols.num;
		oSectionCols.Sep        = oParsedSectCols.sep;
		oSectionCols.Space      = private_Twips2MM(oParsedSectCols.space);

		return oSectionCols;
	};
	ReaderFromJSON.prototype.SectionColFromJSON = function(oParsedSectCol)
	{
		var oSectionColumn = new CSectionColumn();

		oSectionColumn.W     = private_Twips2MM(oParsedSectCol.w);
		oSectionColumn.Space = private_Twips2MM(oParsedSectCol.space);

		return oSectionColumn;
	};
	ReaderFromJSON.prototype.HeaderFromJSON = function(oParsedHdr)
	{
		var oDocument         = private_GetLogicDocument();
		var oDrawingDocuemnt  = private_GetDrawingDocument();
		var oHdrFtrController = oDocument.GetHdrFtr();

		var oHeader = new AscCommonWord.CHeaderFooter(oHdrFtrController, oDocument, oDrawingDocuemnt, AscCommon.hdrftr_Header);
		oHeader.Content = this.DocContentFromJSON(oParsedHdr.content, oHeader);

		return oHeader;
	};
	ReaderFromJSON.prototype.FooterFromJSON = function(oParsedHdr)
	{
		var oDocument         = private_GetLogicDocument();
		var oDrawingDocuemnt  = private_GetDrawingDocument();
		var oHdrFtrController = oDocument.GetHdrFtr();

		var oFooter = new AscCommonWord.CHeaderFooter(oHdrFtrController, oDocument, oDrawingDocuemnt, AscCommon.hdrftr_Footer);
		oFooter.Content = this.DocContentFromJSON(oParsedHdr.content, oFooter);

		return oFooter;
	};
	ReaderFromJSON.prototype.FramePrFromJSON = function(oParsedFramePr)
	{
		var oFramePr = new CFramePr();

		var nDropCapType = undefined;
		switch (oParsedFramePr.dropCap)
		{
			case "none":
				nDropCapType = Asc.c_oAscDropCap.None;
				break;
			case "drop":
				nDropCapType = Asc.c_oAscDropCap.Drop;
				break;
			case "margin":
				nDropCapType = Asc.c_oAscDropCap.Margin;
				break;
		}

		var nHAnchor = undefined;
		switch (oParsedFramePr.hAnchor)
		{
			case "margin":
				nHAnchor = Asc.c_oAscHAnchor.Margin;
				break;
			case "page":
				nHAnchor = Asc.c_oAscHAnchor.Page;
				break;
			case "text":
				nHAnchor = Asc.c_oAscHAnchor.Text;
				break;
		}

		var nVAnchor = undefined;
		switch (oParsedFramePr.vAnchor)
		{
			case "margin":
				nVAnchor = Asc.c_oAscHAnchor.Margin;
				break;
			case "page":
				nVAnchor = Asc.c_oAscHAnchor.Page;
				break;
			case "text":
				nVAnchor = Asc.c_oAscHAnchor.Text;
				break;
		}

		var nLineRule = undefined;
		switch (oParsedFramePr.hRule)
		{
			case "atLeast":
				nLineRule = Asc.linerule_AtLeast;
				break;
			case "auto":
				nLineRule = Asc.linerule_Auto;
				break;
			case "exact":
				nLineRule = Asc.linerule_Exact;
				break;
		}

		var nWrapType = undefined;
		switch (oParsedFramePr.wrap)
		{
			case "around":
				nWrapType = AscCommonWord.wrap_Around;
				break;
			case "auto":
				nWrapType = AscCommonWord.wrap_Auto;
				break;
			case "none":
				nWrapType = AscCommonWord.wrap_None;
				break;
			case "notBeside":
				nWrapType = AscCommonWord.wrap_NotBeside;
				break;
			case "through":
				nWrapType = AscCommonWord.wrap_Through;
				break;
			case "tight":
				nWrapType = AscCommonWord.wrap_Tight;
				break;
		}

		var nXAlign = undefined;
		switch (oParsedFramePr.xAlign)
		{
			case "center":
				nXAlign = Asc.c_oAscXAlign.Center;
				break;
			case "inside":
				nXAlign = Asc.c_oAscXAlign.Inside;
				break;
			case "left":
				nXAlign = Asc.c_oAscXAlign.Left;
				break;
			case "outside":
				nXAlign = Asc.c_oAscXAlign.Outside;
				break;
			case "right":
				nXAlign = Asc.c_oAscXAlign.Right;
				break;
		}

		var nYAlign = undefined;
		switch (oParsedFramePr.yAlign)
		{
			case Asc.c_oAscYAlign.Bottom:
				nYAlign = "bottom";
				break;
			case Asc.c_oAscYAlign.Center:
				nYAlign = "center";
				break;
			case Asc.c_oAscYAlign.Inline:
				nYAlign = "inline";
				break;
			case Asc.c_oAscYAlign.Inside:
				nYAlign = "inside";
				break;
			case Asc.c_oAscYAlign.Outside:
				nYAlign = "outside";
				break;
			case Asc.c_oAscYAlign.Top:
				nYAlign = "top";
				break;
		}

		oFramePr.DropCap = nDropCapType;
		oFramePr.H       = oParsedFramePr.h !== undefined ? private_Twips2MM(oParsedFramePr.h) : oFramePr.H;
		oFramePr.HAnchor = nHAnchor;
		oFramePr.HRule   = nLineRule;
		oFramePr.HSpace  = oParsedFramePr.hSpace !== undefined ? private_Twips2MM(oParsedFramePr.hSpace) : oFramePr.HSpace;
		oFramePr.Lines   = oParsedFramePr.lines;
		oFramePr.VAnchor = nVAnchor;
		oFramePr.VSpace  = oParsedFramePr.vSpace !== undefined ? private_Twips2MM(oParsedFramePr.vSpace) : oFramePr.VSpace;
		oFramePr.W       = oParsedFramePr.w !== undefined ? private_Twips2MM(oParsedFramePr.w) : oFramePr.W;
		oFramePr.Wrap    = nWrapType;
		oFramePr.X       = oParsedFramePr.x !== undefined ? private_Twips2MM(oParsedFramePr.x) : oFramePr.X;
		oFramePr.XAlign  = nXAlign;
		oFramePr.Y       = oParsedFramePr.y !== undefined ? private_Twips2MM(oParsedFramePr.y) : oFramePr.Y;
		oFramePr.YAlign  = nYAlign;

		return oFramePr;		
	};
	ReaderFromJSON.prototype.NumPrFromJSON = function(oParsedNumPr, oParsedNumbering, oPrevNumIdInfo)
	{
		if (oPrevNumIdInfo && oPrevNumIdInfo.sPrevCreatedNumId)
		{
			if (oParsedNumPr.numId === oPrevNumIdInfo.nNumId && oPrevNumIdInfo.sPrevCreatedNumId)
			{
				return new CNumPr(oPrevNumIdInfo.sPrevCreatedNumId, oParsedNumPr.ilvl);
			}
		}
		
		// numbering
		var sNumId  = oParsedNumbering ? this.NumberingFromJSON(oParsedNumbering) : oParsedNumbering; // тут создаем AbstrackNum и CNum 
		var nNumLvl = oParsedNumPr.ilvl;

		if (oPrevNumIdInfo)
		{
			oPrevNumIdInfo.sPrevCreatedNumId = sNumId;
			oPrevNumIdInfo.nNumId            = oParsedNumPr.numId;
		}

		var oNumPr = new CNumPr(sNumId, nNumLvl);
		oNumPr.NumId = oPrevNumIdInfo ? oNumPr.NumId : undefined;

		return oNumPr;
	};
	ReaderFromJSON.prototype.ParaSpacingFromJSON = function(oParsedSpacing)
	{
		var oSpacing = new CParaSpacing();

		oSpacing.Before            = typeof(oParsedSpacing.before) === "number" ? private_Twips2MM(oParsedSpacing.before) : oSpacing.Before;
		oSpacing.BeforePct         = oParsedSpacing.beforePct != undefined ? oParsedSpacing.beforePct : oSpacing.BeforePct;
		oSpacing.BeforeAutoSpacing = oParsedSpacing.beforeAutoSpacing != undefined ? (oParsedSpacing.beforeAutoSpacing === "on" ? true : false) : oSpacing.BeforeAutoSpacing;
		oSpacing.After             = typeof(oParsedSpacing.after) === "number" ? private_Twips2MM(oParsedSpacing.after) : oSpacing.After;
		oSpacing.AfterPct          = oParsedSpacing.afterPct != undefined ? oParsedSpacing.afterPct : oSpacing.AfterPct;
		oSpacing.AfterAutoSpacing  = oParsedSpacing.afterAutoSpacing != undefined ? (oParsedSpacing.afterAutoSpacing === "on" ? true : false) : oSpacing.AfterAutoSpacing;

		switch (oParsedSpacing.lineRule)
		{
			case "atLeast":
				oSpacing.LineRule = linerule_AtLeast;
				oSpacing.Line     = typeof(oParsedSpacing.line) === "number" ? private_Twips2MM(oParsedSpacing.line) : oSpacing.Line;
				break;
			case "auto":
				oSpacing.LineRule = linerule_Auto;
				oSpacing.Line     = typeof(oParsedSpacing.line) === "number" ? private_Twips2MM(oParsedSpacing.line) : oSpacing.Line;
				break;
			case "exact":
				oSpacing.LineRule = linerule_Exact;
				oSpacing.Line     = typeof(oParsedSpacing.line) === "number" ? private_Twips2MM(oParsedSpacing.line) : oSpacing.Line;
				break;
			default:
				oSpacing.LineRule = undefined;
				oSpacing.Line     = undefined;
				break;
		}

		return oSpacing;
	};
	ReaderFromJSON.prototype.NumberingFromJSON = function(oParsedNumbering)
	{
		var oDocument    = private_GetLogicDocument();
		var oAbstractNum = this.AbstractNumFromJSON(oParsedNumbering.abstractNum);
		var oNum         = new CNum(private_GetLogicDocument().Numbering, oAbstractNum.GetId());
		oNum             = this.CNumFromJSON(oNum, oParsedNumbering.num.lvlOverride);
		var sNumId       = oNum.GetId();

		oDocument.Numbering.Num[sNumId] = oNum;

		return sNumId;
	};
	ReaderFromJSON.prototype.AbstractNumFromJSON = function(oParsedAbstrNum)
	{
		var oDocument    = private_GetLogicDocument();
		var oAbstractNum = new CAbstractNum();

		oAbstractNum.Lvl = [];
		for (var nLvl = 0; nLvl < oParsedAbstrNum.lvl.length; nLvl++)
			oAbstractNum.Lvl.push(this.NumLvlFromJSON(oParsedAbstrNum.lvl[nLvl]));

		oAbstractNum.NumStyleLink = oParsedAbstrNum.numStyleLink;
		oAbstractNum.styleLink    = oParsedAbstrNum.StyleLink;

		//if (oDocument.Numbering.AbstractNum[oParsedAbstrNum.abstractNumId] && oAbstractNum.isEqual(oDocument.Numbering.AbstractNum[oParsedAbstrNum.abstractNumId]))
		//	return oDocument.Numbering.AbstractNum[oParsedAbstrNum.abstractNumId];

		oDocument.Numbering.AbstractNum[oAbstractNum.GetId()] = oAbstractNum;
		return oAbstractNum;
	};
	ReaderFromJSON.prototype.CNumFromJSON = function(oNum, arrLvlOverride)
	{
		for (var nLvl = 0; nLvl < arrLvlOverride; nLvl++)
			oNum.LvlOverride.push(new CLvlOverride(this.NumLvlFromJSON(arrLvlOverride[nLvl].lvl), arrLvlOverride[nLvl].ilvl, arrLvlOverride[nLvl].startOverride));
	
		return oNum;
	};
	ReaderFromJSON.prototype.NumLvlFromJSON = function(oParsedNumLvl)
	{
		var oNumLvl = new CNumberingLvl();

		// align
		var nJc = undefined;
		switch (oParsedNumLvl.lvlJc)
		{
			case "end":
				nJc = AscCommon.align_Right;
				break;
			case "start":
				nJc = AscCommon.align_Left;
				break;
			case "center":
				nJc = AscCommon.align_Center;
				break;
			case "both":
				nJc = AscCommon.align_Justify;
				break;
			case "distribute":
				nJc = AscCommon.align_Distributed;
				break;
		}

		// style 
		var oStyle    = oParsedNumLvl.pStyle ? this.StyleFromJSON(oParsedNumLvl.pStyle) : oParsedNumLvl.pStyle;
		var oStyles   = private_GetStyles();
		if (oStyle)
		{
			var nExistingStyle = oStyles.GetStyleIdByName(oStyle.Name);
			// если такого стиля нет - добавляем новый
			if (nExistingStyle === null)
			{
				oStyle.Set_BasedOn(null);
				oStyles.Add(oStyle);
			}
			else
			{
				var oExistingStyle = oStyles.Get(nExistingStyle);
				// если стили идентичны, стиль не добавляем
				if (!oStyle.IsEqual(oExistingStyle))
				{
					oStyle.Set_BasedOn(null);
					oStyles.Add(oStyle);
				}
				else
					oStyle = oExistingStyle;
			}
		}

		// suff
		var nSuffType = undefined;
		switch (oParsedNumLvl.suff)
		{
			case "tab":
				nSuffType = Asc.c_oAscNumberingSuff.Tab;
				break;
			case "space":
				nSuffType = Asc.c_oAscNumberingSuff.Space;
				break;
			case "nothing":
				nSuffType = Asc.c_oAscNumberingSuff.None;
				break;
		}

		// format type
		var nFormatType = From_XML_c_oAscNumberingFormat(oParsedNumLvl.numFmt.val);

		oNumLvl.IsLgl   = oParsedNumLvl.isLgl;
		oNumLvl.Legacy  = oParsedNumLvl.legacy ? new CNumberingLvlLegacy(oParsedNumLvl.legacy.legacy, oParsedNumLvl.legacy.legacyIndent, oParsedNumLvl.legacy.legacySpace) : oNumLvl.Legacy;
		oNumLvl.Jc      = nJc;
		oNumLvl.Format  = nFormatType;
		oNumLvl.ParaPr  = this.ParaPrFromJSON(oParsedNumLvl.pPr);
		oNumLvl.TextPr  = this.TextPrFromJSON(oParsedNumLvl.rPr);
		oNumLvl.Suff    = nSuffType;
		oNumLvl.Restart = oParsedNumLvl.restart;
		oNumLvl.Start   = oParsedNumLvl.start;
		oNumLvl.PStyle  = oStyle ? oStyle.Id : oNumLvl.PStyle;
		if (oParsedNumLvl.lvlText.numValue !== undefined)
		{
			oNumLvl.LvlText[0] = new CNumberingLvlTextNum(oParsedNumLvl.lvlText.numValue);
			oNumLvl.LvlText[1] = new CNumberingLvlTextString(oParsedNumLvl.lvlText.val);
		}
		else
			oNumLvl.LvlText[0] = new CNumberingLvlTextString(oParsedNumLvl.lvlText.val);

		return oNumLvl;
	};
	ReaderFromJSON.prototype.HyperlinkFromJSON = function(oParsedLink, oParentPara, notCompletedFields)
	{
		var aContent = oParsedLink.content;   
		var oHyper   = new AscCommonWord.ParaHyperlink();

		// Заполняем гиперссылку полями
		if (undefined !== oParsedLink.anchor && null !== oParsedLink.anchor)
		{
			oHyper.SetAnchor(oParsedLink.anchor);
			oHyper.SetValue("")
		}
		else if (undefined != oParsedLink.value)
		{
			oHyper.SetValue(oParsedLink.value);
			oHyper.SetAnchor("");
		}
		if (undefined != oParsedLink.toolTip)
			oHyper.SetToolTip(oParsedLink.toolTip);


		for (var nElm = 0; nElm < aContent.length; nElm++)
		{
			switch(aContent[nElm].type)
			{
				case "run":
				case "mathRun":
					oHyper.AddToContent(oHyper.Content.length, this.ParaRunFromJSON(aContent[nElm], oParentPara, notCompletedFields));
			}
		}

		return oHyper;
	};
	ReaderFromJSON.prototype.InlineLvlSdtFromJSON = function(oParsedSdt, oParentPara, notCompletedFields)
	{
		if (!notCompletedFields)
			notCompletedFields = [];

		var oContentControl = new AscCommonWord.CInlineLevelSdt();
		var aContent        = oParsedSdt.content;

		oContentControl.Pr = this.SdtPrFromJSON(oParsedSdt.sdtPr);

		for (var nElm = 0; nElm < aContent.length; nElm++)
		{
			switch (aContent[nElm].type)
			{
				case "run":
				case "mathRun":
					oContentControl.AddToContent(oContentControl.Content.length, this.ParaRunFromJSON(aContent[nElm], oParentPara, notCompletedFields));
					break;
				case "hyperlink":
					oContentControl.AddToContent(oContentControl.Content.length, this.HyperlinkFromJSON(aContent[nElm], notCompletedFields));
					break;
			}
		}

		return oContentControl;
	};
	ReaderFromJSON.prototype.BlockLvlSdtFromJSON = function(oParsedSdt, oParent, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo)
	{
		if (!notCompletedFields)
			notCompletedFields = [];
		if (!oMapCommentsInfo)
			oMapCommentsInfo = {};
		if (!oMapBookmarksInfo)
			oMapBookmarksInfo = {};

		var oContentControl = new AscCommonWord.CBlockLevelSdt(private_GetLogicDocument(), oParent || private_GetLogicDocument());
		
		oContentControl.SetPr(this.SdtPrFromJSON(oParsedSdt.sdtPr));
		
		oContentControl.Content.Copy2(this.DocContentFromJSON(oParsedSdt.sdtContent, oContentControl, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo));

		return oContentControl;
	};
	ReaderFromJSON.prototype.TableCellFromJSON = function(oParsedCell, oParentRow, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo)
	{
		var oCell = new CTableCell(oParentRow);

		oParsedCell.tcPr && oCell.Set_Pr(this.TableCellPrFromJSON(oParsedCell.tcPr));

		oCell.Content.Copy2(this.DocContentFromJSON(oParsedCell.content, oCell, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo));
	
		return oCell;
	};
	ReaderFromJSON.prototype.TableCellPrFromJSON = function(oParsedPr)
	{
		var oTableCellPr = new CTableCellPr();

		oTableCellPr.GridSpan   = oParsedPr.gridSpan;
		oTableCellPr.HMerge     = oParsedPr.hMerge ? (oParsedPr.hMerge === "continue" ? 2 : 1) : oParsedPr.hMerge;
		oTableCellPr.VMerge     = oParsedPr.vMerge ? (oParsedPr.vMerge === "continue" ? 2 : 1) : oParsedPr.vMerge;
		oTableCellPr.NoWrap     = oParsedPr.noWrap;
		oTableCellPr.Shd        = oParsedPr.shd ? this.ShadeFromJSON(oParsedPr.shd) : oParsedPr.shd;

		oTableCellPr.TableCellBorders.Bottom = oParsedPr.tcBorders.bottom ? this.DocBorderFromJSON(oParsedPr.tcBorders.bottom) : oParsedPr.tcBorders.bottom;
		oTableCellPr.TableCellBorders.Right  = oParsedPr.tcBorders.end    ? this.DocBorderFromJSON(oParsedPr.tcBorders.end)    : oParsedPr.tcBorders.end;
		oTableCellPr.TableCellBorders.Left   = oParsedPr.tcBorders.start  ? this.DocBorderFromJSON(oParsedPr.tcBorders.start)  : oParsedPr.tcBorders.start;
		oTableCellPr.TableCellBorders.Top    = oParsedPr.tcBorders.top    ? this.DocBorderFromJSON(oParsedPr.tcBorders.top)    : oParsedPr.tcBorders.top;
		
		if (oParsedPr.tcMar)
		{
			oTableCellPr.TableCellMar        = {};
			oTableCellPr.TableCellMar.Bottom = oParsedPr.tcMar.bottom ? this.TableMeasurementFromJSON(oParsedPr.tcMar.bottom) : oParsedPr.tcMar.bottom;
			oTableCellPr.TableCellMar.Right  = oParsedPr.tcMar.right  ? this.TableMeasurementFromJSON(oParsedPr.tcMar.right)  : oParsedPr.tcMar.right;
			oTableCellPr.TableCellMar.Left   = oParsedPr.tcMar.left   ? this.TableMeasurementFromJSON(oParsedPr.tcMar.left)   : oParsedPr.tcMar.left;
			oTableCellPr.TableCellMar.Top    = oParsedPr.tcMar.top    ? this.TableMeasurementFromJSON(oParsedPr.tcMar.top)    : oParsedPr.tcMar.top;
		}

		oTableCellPr.PrChange   = oParsedPr.tcPrChange ? this.TableCellPrFromJSON(oParsedPr.tcPrChange) : oParsedPr.tcPrChange;
		oTableCellPr.TableCellW = oParsedPr.tcW        ? this.TableMeasurementFromJSON(oParsedPr.tcW)   : oParsedPr.tcW;

		// text direction
		var nTextDir = undefined;
		switch (oParsedPr.textDirection)
		{
			case "lrtb":
				nTextDir = textdirection_LRTB;
				break;
			case "tbrl":
				nTextDir = textdirection_TBRL;
				break;
			case "btlr":
				nTextDir = textdirection_BTLR;
				break;
			case "lrtbV":
				nTextDir = textdirection_LRTBV;
				break;
			case "tbrlV":
				nTextDir = textdirection_TBRLV;
				break;
			case "tblrV":
				nTextDir = textdirection_TBLRV;
				break;
		}

		oTableCellPr.TextDirection = nTextDir;

		// alignV
		var nVAlign = undefined;
		switch (oParsedPr.vAlign)
		{
			case "top":
				nVAlign = vertalignjc_Top;
				break;
			case "center":
				nVAlign = vertalignjc_Center;
				break;
			case "bottom":
				nVAlign = vertalignjc_Bottom;
				break;
		}

		oTableCellPr.VAlign = nVAlign;

		return oTableCellPr;
	};
	ReaderFromJSON.prototype.TableMeasurementFromJSON = function(oParsedObj)
	{
		var nType = tblwidth_Auto;
		var nW    = 0;

		switch (oParsedObj.type)
		{
			case "auto":
				nType = tblwidth_Auto;
				nW    = 0;
				break;
			case "dxa":
				nType = tblwidth_Mm;
				nW    = private_Twips2MM(oParsedObj.w)
				break;
			case "nil":
				nType = tblwidth_Nil;
				nW    = 0;
				break;
			case "pct":
				nType = tblwidth_Pct;
				nW    = oParsedObj.w;
				break;
		}

		return new CTableMeasurement(nType, nW);
	};
	ReaderFromJSON.prototype.TableRowFromJSON = function(oParsedRow, oParentTable, nIndex, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo)
	{
		var oRow     = new CTableRow(oParentTable);
		var aContent = oParsedRow.content;

		if (nIndex >= 0)
			oRow.Index = nIndex;

		// review info
		var nReviewType = undefined;
		switch (oParsedRow.reviewType)
		{
			case "common":
				nReviewType = reviewtype_Common;
				break;
			case "remove":
				nReviewType = reviewtype_Remove;
				break;
			case "add":
				nReviewType = reviewtype_Add;
				break;
		}
		var oReviewInfo = oParsedRow.reviewInfo ? this.ReviewInfoFromJSON(oParsedRow.reviewInfo) : oRow.ReviewInfo;
		oRow.SetReviewTypeWithInfo(nReviewType, oReviewInfo);

		oRow.Set_Pr(this.TableRowPrFromJSON(oParsedRow.trPr));

		for (var nCell = 0; nCell < aContent.length; nCell++)
			oRow.Add_Cell(nCell, oRow, this.TableCellFromJSON(aContent[nCell], oRow, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo), false);

		return oRow;
	};
	ReaderFromJSON.prototype.TableRowPrFromJSON = function(oParsedPr)
	{
		var oRowPr = new CTableRowPr();

		oRowPr.CantSplit  = oParsedPr.cantSplit;
		oRowPr.GridAfter  = oParsedPr.gridAfter;
		oRowPr.GridBefore = oParsedPr.gridBefore;

		// spacing
		oRowPr.TableCellSpacing = typeof(oParsedPr.tblCellSpacing) === "number" ? private_Twips2MM(oParsedPr.tblCellSpacing) : oParsedPr.tblCellSpacing;
		
		// rowJc
		var nRowJc = undefined;
		switch (oParsedPr.jc)
		{
			case "start":
				nRowJc = align_Left;
				break;
			case "center":
				nRowJc = align_Center;
				break;
			case "end":
				nRowJc = align_Right;
				break;
		}
		oRowPr.Jc = nRowJc;

		// rowHeight
		var oRowHeight = undefined;
		if (oParsedPr.trHeight)
		{
			switch (oParsedPr.trHeight.hRule)
			{
				case "atLeast":
					oRowHeight = new CTableRowHeight(private_Twips2MM(oParsedPr.trHeight.val), linerule_AtLeast)
					break;
				case "auto":
					oRowHeight = new CTableRowHeight(0, linerule_Auto);
					break;
				case "exact":
					oRowHeight = new CTableRowHeight(private_Twips2MM(oParsedPr.trHeight.val), linerule_Exact)
					break;
			}
		}

		oRowPr.Height      = oRowHeight;
		oRowPr.PrChange    = oParsedPr.trPrChange ? this.TableRowPrFromJSON(oParsedPr.trPrChange) : oParsedPr.trPrChange;
		oRowPr.WAfter      = oParsedPr.wAfter     ? this.TableMeasurementFromJSON(oParsedPr.wAfter)  : oParsedPr.wAfter;
		oRowPr.WBefore     = oParsedPr.wBefore    ? this.TableMeasurementFromJSON(oParsedPr.wBefore) : oParsedPr.wBefore;
		oRowPr.TableHeader = oParsedPr.tblHeader;
		
		return oRowPr;
	};
	ReaderFromJSON.prototype.TableFromJSON = function(oParsedTable, oParent, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo)
	{
		var aTableGrid = [];
		for (var nGrid = 0; nGrid < oParsedTable.tblGrid.length; nGrid++)
			aTableGrid.push(private_Twips2MM(oParsedTable.tblGrid[nGrid].w));

		var oTable     = new CTable(private_GetDrawingDocument(), oParent || private_GetLogicDocument(), true, 0, 0, aTableGrid, oParsedTable.bPresentation);
		var aContent   = oParsedTable.content; 
	
		// table prop.
		oTable.SetPr(this.TablePrFromJSON(oTable, oParsedTable.tblPr));

		// fill table content
		for (var nRow = 0; nRow < aContent.length; nRow++)
			oTable.private_AddRow(nRow, 0, false, this.TableRowFromJSON(aContent[nRow], oTable, nRow, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo));

		// выставляем текущую ячейку
		oTable.CurCell = oTable.Content[0].Get_Cell(0);

		//oTable.Update_TableMarkupFromRuler(this.TableMarkupFromJSON(oParsedTable.tableMarkup, oTable), true, 0);
		//oTable.Markup = this.TableMarkupFromJSON(oParsedTable.tableMarkup, oTable);

		for (var nChange = 0; nChange < oParsedTable.changes.length; nChange++)
			this.RevisionFromJSON(oParsedTable.changes[nChange], oTable);

		return oTable;
	};
	ReaderFromJSON.prototype.TableMarkupFromJSON = function(oParsedMarkup, oParentTable)
	{
		var oMarkup = new AscCommon.CTableMarkup(oParentTable);

		oMarkup.Cols   = oParsedMarkup.cols;
		oMarkup.CurCol = oParsedMarkup.curCol;
		oMarkup.CurRow = oParsedMarkup.curRow;

		oMarkup.Internal = {
			CellIndex: oParsedMarkup.internal.cellIndex,
			PageNum:   oParsedMarkup.internal.pageNum,
			RowIndex:  oParsedMarkup.internal.rowIndex
		}

		oMarkup.Margins    = oParsedMarkup.margins;
		oMarkup.Rows       = oParsedMarkup.rows;
		oMarkup.TransformX = oParsedMarkup.transformX;
		oMarkup.TransformY = oParsedMarkup.transformY;
		oMarkup.X          = oParsedMarkup.x;

		return oMarkup;
	};
	ReaderFromJSON.prototype.TablePrFromJSON = function(oParentTable, oParsedPr)
	{
		var oTablePr = new CTablePr();

		oTablePr.TableStyleColBandSize = oParsedPr.tblStyleColBandSize;
		oTablePr.TableStyleRowBandSize = oParsedPr.tblStyleRowBandSize;

		// jc
		var nJc          = undefined;
		switch (oParsedPr.jc)
		{
			case "start":
				nJc = AscCommon.align_Left;
				break;
			case "center":
				nJc = AscCommon.align_Center;
				break;
			case "end":
				nJc = AscCommon.align_Right;
				break;
		}
		oTablePr.Jc  = nJc;

		oTablePr.Shd = oParsedPr.shd ? this.ShadeFromJSON(oParsedPr.shd) : oParsedPr.shd;

		// borders
		oTablePr.TableBorders.Bottom  = oParsedPr.tblBorders.bottom  ? this.DocBorderFromJSON(oParsedPr.tblBorders.bottom)  : oParsedPr.tblBorders.bottom;
		oTablePr.TableBorders.Right   = oParsedPr.tblBorders.end     ? this.DocBorderFromJSON(oParsedPr.tblBorders.end)     : oParsedPr.tblBorders.end;
		oTablePr.TableBorders.InsideH = oParsedPr.tblBorders.insideH ? this.DocBorderFromJSON(oParsedPr.tblBorders.insideH) : oParsedPr.tblBorders.insideH;
		oTablePr.TableBorders.InsideV = oParsedPr.tblBorders.insideV ? this.DocBorderFromJSON(oParsedPr.tblBorders.insideV) : oParsedPr.tblBorders.insideV;
		oTablePr.TableBorders.Left    = oParsedPr.tblBorders.start   ? this.DocBorderFromJSON(oParsedPr.tblBorders.start)   : oParsedPr.tblBorders.start;
		oTablePr.TableBorders.Top     = oParsedPr.tblBorders.top     ? this.DocBorderFromJSON(oParsedPr.tblBorders.top)     : oParsedPr.tblBorders.top;

		// margins
		if (oParsedPr.tblCellMar)
		{
			oTablePr.TableCellMar.Bottom = oParsedPr.tblCellMar.bottom ? this.TableMeasurementFromJSON(oParsedPr.tblCellMar.bottom) : oParsedPr.tblCellMar.bottom;
			oTablePr.TableCellMar.Left   = oParsedPr.tblCellMar.left   ? this.TableMeasurementFromJSON(oParsedPr.tblCellMar.left)   : oParsedPr.tblCellMar.left;
			oTablePr.TableCellMar.Right  = oParsedPr.tblCellMar.right  ? this.TableMeasurementFromJSON(oParsedPr.tblCellMar.right)  : oParsedPr.tblCellMar.right;
			oTablePr.TableCellMar.Top    = oParsedPr.tblCellMar.top    ? this.TableMeasurementFromJSON(oParsedPr.tblCellMar.top)    : oParsedPr.tblCellMar.top;
		}

		oTablePr.TableCellSpacing = oParsedPr.tblCellSpacing ? private_Twips2MM(oParsedPr.tblCellSpacing) : oParsedPr.tblCellSpacing;
		oTablePr.TableInd         = oParsedPr.tblInd         ? private_Twips2MM(oParsedPr.tblInd)         : oParsedPr.tblInd;
		oTablePr.TableW           = oParsedPr.tblW           ? this.TableMeasurementFromJSON(oParsedPr.tblW)   : oParsedPr.tblW;

		oTablePr.TableLayout      = oParsedPr.tblLayout == undefined ? oParsedPr.tblLayout : (oParsedPr.tblLayout === "fixed" ? tbllayout_Fixed : tbllayout_AutoFit);
		oTablePr.TableDescription = oParsedPr.tblDescription;
		oTablePr.TableCaption     = oParsedPr.tblCaption;

		var oPrChane    = oParsedPr.tblPrChange ? this.TablePrFromJSON(oParentTable, oParsedPr.tblPrChange) : oParsedPr.tblPrChange;
		var oReviewInfo = oParsedPr.reviewInfo ? this.ReviewInfoFromJSON(oParsedPr.reviewInfo) : oTablePr.ReviewInfo;
		oPrChane && oReviewInfo && oTablePr.SetPrChange(oPrChane, oReviewInfo);

		// if oParentTalbe is exist
		if (oParentTable)
		{
			oParsedPr.tblOverlap === "overlap" ? oParentTable.Set_AllowOverlap(true) : oParentTable.Set_AllowOverlap(false);
			oParentTable.Inline !== undefined && oParentTable.Set_Inline(oParsedPr.inline);
			
			if (oParsedPr.tblLook)
			{
				var oNewLook = new CTableLook(oParsedPr.tblLook.firstColumn, oParsedPr.tblLook.firstRow, oParsedPr.tblLook.lastColumn, oParsedPr.tblLook.lastRow, !oParsedPr.tblLook.noHBand, !oParsedPr.tblLook.noVBand);
				oParentTable.Set_TableLook(oNewLook);
			}

			// position prop
			if (oParsedPr.tblpPr)
			{
				// hAnchor
				var nHorAnchor = undefined;
				switch (oParsedPr.tblpPr.horzAnchor)
				{
					case "margin":
						nHorAnchor = Asc.c_oAscHAnchor.Margin;
						break;
					case "text":
						nHorAnchor = Asc.c_oAscHAnchor.Text;
						break;
					case "page":
						nHorAnchor = Asc.c_oAscHAnchor.Page;
						break;
				}

				// vAnchor
				var nVerAnchor = undefined;
				switch (oParsedPr.tblpPr.vertAnchor)
				{
					case "margin":
						nVerAnchor = Asc.c_oAscHAnchor.Margin;
						break;
					case "text":
						nVerAnchor = Asc.c_oAscHAnchor.Text;
						break;
					case "page":
						nVerAnchor = Asc.c_oAscHAnchor.Page;
						break;
				}

				// alignH
				var nHorAlign = undefined;
				switch (oParsedPr.tblpPr.tblpXSpec)
				{
					case "center":
						nHorAlign = Asc.c_oAscXAlign.Center;
						break;
					case "inside":
						nHorAlign = Asc.c_oAscXAlign.Inside;
						break;
					case "left":
						nHorAlign = Asc.c_oAscXAlign.Left;
						break;
					case "outside":
						nHorAlign = Asc.c_oAscXAlign.Outside;
						break;
					case "right":
						nHorAlign = Asc.c_oAscXAlign.Right;
						break;
				}
				
				// alignV
				var nVerAlign = undefined;
				switch (oParsedPr.tblpPr.tblpYSpec)
				{
					case "bottom":
						nVerAlign = Asc.c_oAscYAlign.Bottom;
						break;
					case "center":
						nVerAlign = Asc.c_oAscYAlign.Center;
						break;
					case "inline":
						nVerAlign = Asc.c_oAscYAlign.Inline;
						break;
					case "inside":
						nVerAlign = Asc.c_oAscYAlign.Inside;
						break;
					case "outside":
						nVerAlign = Asc.c_oAscYAlign.Outside;
						break;
					case "top":
						nVerAlign = Asc.c_oAscYAlign.Top;
						break;
				}

				oParentTable.Set_PositionH(nHorAnchor, nHorAlign != undefined, nHorAlign == undefined ? private_Twips2MM(oParsedPr.tblpPr.tblpX) : nHorAlign);
				oParentTable.Set_PositionV(nVerAnchor, nVerAlign != undefined, nVerAlign == undefined ? private_Twips2MM(oParsedPr.tblpPr.tblpY) : nVerAlign);

				oParentTable.Set_Distance(private_Twips2MM(oParsedPr.tblpPr.bottomFromText), private_Twips2MM(oParsedPr.tblpPr.leftFromText), private_Twips2MM(oParsedPr.tblpPr.rightFromText), private_Twips2MM(oParsedPr.tblpPr.topFromText));
			}

			// style
			if (oParsedPr.tblStyle)
			{
				// style 
				var oStyle    = this.StyleFromJSON(oParsedPr.tblStyle);
				var oStyles   = private_GetStyles();
				if (oStyle)
				{
					var nExistingStyle = oStyles.GetStyleIdByName(oStyle.Name);
					// если такого стиля нет - добавляем новый
					if (nExistingStyle === null)
					{
						oStyle.Set_BasedOn(null);
						oStyles.Add(oStyle);
					}
					else
					{
						var oExistingStyle = oStyles.Get(nExistingStyle);
						// если стили идентичны, стиль не добавляем
						if (!oStyle.IsEqual(oExistingStyle))
						{
							oStyle.Set_BasedOn(null);
							oStyles.Add(oStyle);
						}
						else
							oStyle = oExistingStyle;
					}
				}

				oParentTable.Set_TableStyle(oStyle.Id);
			}
		}
	
		return oTablePr;
	};
	ReaderFromJSON.prototype.DocContentFromJSON = function(oParsedDocContent, oParent, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo)
	{
		var oDocContent = new AscCommonWord.CDocumentContent(oParent || private_GetLogicDocument(), private_GetDrawingDocument(), 0, 0, 0, 0, true, false, oParsedDocContent.bPresentation);
		var aContent    = oParsedDocContent.content;
	
		if (!notCompletedFields)
			notCompletedFields = [];
		if (!oMapCommentsInfo)
			oMapCommentsInfo = {};
		if (!oMapBookmarksInfo)
			oMapBookmarksInfo = {};

		var oPrevNumIdInfo = {};

		for (var nElm = 0; nElm < aContent.length; nElm++)
		{
			switch (aContent[nElm].type)
			{
				case "paragraph":
					oDocContent.AddToContent(oDocContent.Content.length, this.ParagraphFromJSON(aContent[nElm], oDocContent, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo, oPrevNumIdInfo), false);
					break;
				case "table":
					oDocContent.AddToContent(oDocContent.Content.length, this.TableFromJSON(aContent[nElm], oDocContent, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo), false);
					break;
				case "blockLvlSdt":
					oDocContent.AddToContent(oDocContent.Content.length, this.BlockLvlSdtFromJSON(aContent[nElm], oDocContent, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo), false);
					break;
			}
		}

		if (oDocContent.Content.length > 1)
			// удаляем параграф, который добавляется при создании CDocumentContent
			oDocContent.RemoveFromContent(0, 1);

		return oDocContent;
	};
	ReaderFromJSON.prototype.ContentFromJSON = function(oParsedDocContent, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo)
	{
		var aContent = oParsedDocContent;
		var aResult = [];
		var oDocument = private_GetLogicDocument()

		if (!notCompletedFields)
			notCompletedFields = [];
		if (!oMapCommentsInfo)
			oMapCommentsInfo = {};
		if (!oMapBookmarksInfo)
			oMapBookmarksInfo = {};

		var oPrevNumIdInfo = {};

		for (var nElm = 0; nElm < aContent.length; nElm++)
		{
			switch (aContent[nElm].type)
			{
				case "paragraph":
					aResult.push(this.ParagraphFromJSON(aContent[nElm], oDocument, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo, oPrevNumIdInfo));
					break;
				case "table":
					aResult.push(this.TableFromJSON(aContent[nElm], oDocument, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo));
					break;
				case "blockLvlSdt":
					aResult.push(this.BlockLvlSdtFromJSON(aContent[nElm], oDocument, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo));
					break;
			}
		}

		return aResult;
	};
	ReaderFromJSON.prototype.DrawingDocContentFromJSON = function(oParsedDocContent, oParent, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo)
	{
		var oDocContent = new AscFormat.CDrawingDocContent(oParent ? oParent : private_GetLogicDocument(), private_GetDrawingDocument(), 0, 0, 0, 0, true, false, false);
		var aContent    = oParsedDocContent.content;
	
		if (!notCompletedFields)
			notCompletedFields = [];
		if (!oMapCommentsInfo)
			oMapCommentsInfo = {};
		if (!oMapBookmarksInfo)
			oMapBookmarksInfo = {};
			
		var oPrevNumIdInfo = {};

		for (var nElm = 0; nElm < aContent.length; nElm++)
		{
			switch (aContent[nElm].type)
			{
				case "paragraph":
					oDocContent.AddToContent(oDocContent.Content.length, this.ParagraphFromJSON(aContent[nElm], oDocContent, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo, oPrevNumIdInfo), false);
					break;
				case "table":
					oDocContent.AddToContent(oDocContent.Content.length, this.TableFromJSON(aContent[nElm], oDocContent, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo), false);
					break;
				case "blockLvlSdt":
					oDocContent.AddToContent(oDocContent.Content.length, this.BlockLvlSdtFromJSON(aContent[nElm], oDocContent, notCompletedFields, oMapCommentsInfo, oMapBookmarksInfo), false);
					break;
			}
		}

		if (oDocContent.Content.length > 1)
			// удаляем параграф, который добавляется при создании CDocumentContent
			oDocContent.RemoveFromContent(0, 1);

		return oDocContent;
	};
	ReaderFromJSON.prototype.SdtPrFromJSON = function(oParsedSdtPr)
	{
		var oSdtPr = new AscCommonWord.CSdtPr();
		var oTempListItem;

		oSdtPr.Alias = oParsedSdtPr.alias;

		// comboboxPr
		if (oParsedSdtPr.comboBox)
		{
			var oComboboxPr   = new AscWord.CSdtComboBoxPr();
			oTempListItem = null;

			oComboboxPr.LastValue = oParsedSdtPr.comboBox.lastValue;
			for (var nItem = 0; nItem < oParsedSdtPr.comboBox.listItem.length; nItem++)
			{
				oTempListItem             = new AscWord.CSdtListItem();
				oTempListItem.DisplayText = oParsedSdtPr.comboBox.listItem[nItem].displayText;
				oTempListItem.Value       = oParsedSdtPr.comboBox.listItem[nItem].value;
				
				oComboboxPr.ListItems.push(oTempListItem);
			}

			oSdtPr.ComboBox = oComboboxPr;
		}

		// date
		if (oParsedSdtPr.date)
		{
			var oDate = new AscWord.CSdtDatePickerPr();

			oDate.FullDate   = oParsedSdtPr.date.fullDate;
			oDate.LangId     = oParsedSdtPr.date.lid;
			oDate.DateFormat = oParsedSdtPr.date.dateFormat;
			oDate.Calendar   = oParsedSdtPr.date.calendar;

			oSdtPr.Date = oDate;
		}

		// docPartObj
		oSdtPr.DocPartObj.Gallery  = oParsedSdtPr.docPartObj.docPartGallery;
		oSdtPr.DocPartObj.Category = oParsedSdtPr.docPartObj.docPartCategory;
		oSdtPr.DocPartObj.Unique   = oParsedSdtPr.docPartObj.docPartUnique;

		// dropdown
		if (oParsedSdtPr.comboBox)
		{
			var oDropDownPr   = new AscWord.CSdtComboBoxPr();
			oTempListItem = null;

			oDropDownPr.LastValue = oParsedSdtPr.dropDownList.lastValue;
			for (var nItem = 0; nItem < oParsedSdtPr.dropDownList.listItem.length; nItem++)
			{
				oTempListItem             = new AscWord.CSdtListItem();
				oTempListItem.DisplayText = oParsedSdtPr.dropDownList.listItem[nItem].displayText;
				oTempListItem.Value       = oParsedSdtPr.dropDownList.listItem[nItem].value;
				
				oDropDownPr.ListItems.push(oTempListItem);
			}

			oSdtPr.DropDown = oDropDownPr;
		}

		// lock type
		var nLockType = undefined;
		switch (oSdtPr.Lock)
		{
			case "contentLocked":
				nLockType = Asc.c_oAscSdtLockType.ContentLocked;
				break;
			case "sdtContentLocked":
				nLockType = Asc.c_oAscSdtLockType.SdtContentLocked;
				break;
			case "sdtLocked":
				nLockType = Asc.c_oAscSdtLockType.SdtLocked;
				break;
			case "unlocked":
				nLockType = Asc.c_oAscSdtLockType.Unlocked;
				break;
		}
		oSdtPr.Equation      = oParsedSdtPr.equation;
		oSdtPr.Id            = oParsedSdtPr.id;
		oSdtPr.Label         = oParsedSdtPr.label;
		oSdtPr.Lock          = nLockType;
		oSdtPr.Picture       = oParsedSdtPr.picture;
		oSdtPr.Placeholder   = oParsedSdtPr.placeholder ? oParsedSdtPr.placeholder.docPart : oSdtPr.Placeholder;
		oSdtPr.TextPr        = this.TextPrFromJSON(oParsedSdtPr.rPr);
		oSdtPr.ShowingPlcHdr = oParsedSdtPr.showingPlcHdr;
		oSdtPr.Tag           = oParsedSdtPr.tag;
		oSdtPr.Temporary     = oParsedSdtPr.temporary;
		oSdtPr.Text          = oParsedSdtPr.text.multiLine;

		return oSdtPr;
	};
	ReaderFromJSON.prototype.TabsFromJSON = function(oTabs)
	{
		var oTabsObj = new CParaTabs();

		for (var nTab = 0; nTab < oTabs.tabs.length; nTab++)
			oTabsObj.Tabs.push(new CParaTab(oTabs.tabs[nTab].val, private_Twips2MM(oTabs.tabs[nTab].pos), oTabs.tabs[nTab].leader));

		return oTabsObj;
	};
	ReaderFromJSON.prototype.DocBorderFromJSON = function(oBorder)
	{
		var oBorderObj = new CDocumentBorder();

		oBorderObj.Color.r    = oBorder.color.r;
		oBorderObj.Color.g    = oBorder.color.g;
		oBorderObj.Color.b    = oBorder.color.b;
		oBorderObj.Color.Auto = oBorder.color.auto;

		oBorderObj.Size  = private_Twips2MM(oBorder.sz);
		oBorderObj.Space = private_Twips2MM(oBorder.space);
		oBorderObj.Value = oBorder.value === "none" ? border_None : border_Single;

		oBorderObj.LineRef    = oBorder.lineRef    ? this.StyleRefFromJSON(oBorder.lineRef) : oBorder.lineRef;
		oBorderObj.Unifill    = oBorder.themeColor ? this.FillFromJSON(oBorder.themeColor)  : oBorder.themeColor;

		return oBorderObj;
	};
	ReaderFromJSON.prototype.StyleRefFromJSON = function(oParsedRef)
	{
		var oStyleRefObj = new AscFormat.StyleRef();
		
		oStyleRefObj.idx   = oParsedRef.idx;
		oStyleRefObj.Color = oParsedRef.color ? this.ColorFromJSON(oParsedRef.color) : oStyleRefObj.Color;
		
		return oStyleRefObj;
	};
	ReaderFromJSON.prototype.FontRefFromJSON = function(oParsedRef)
	{
		var oFontRefObj = new AscFormat.FontRef();
		
		oFontRefObj.idx   = oParsedRef.idx;
		oFontRefObj.Color = oParsedRef.color ? this.ColorFromJSON(oParsedRef.color) : oFontRefObj.Color;
		
		return oFontRefObj;
	};
	ReaderFromJSON.prototype.FillFromJSON = function(oParsedFill)
	{
		var oFillObj = new AscFormat.CUniFill();
		if (oParsedFill.type)
		{
			switch (oParsedFill.fill.type)
			{
				case "none":
					oFillObj.fill = null;
					break;
				case "solid":
					oFillObj.fill       = new AscFormat.CSolidFill();
					oFillObj.fill.color = this.ColorFromJSON(oParsedFill.fill.color);
					break;
				case "blipFill":
					oFillObj.fill = this.BlipFillFromJSON(oParsedFill.fill);
					break;
				case "noFill":
					oFillObj.fill = new AscFormat.CNoFill();
					break;
				case "gradFill":
					oFillObj.fill = this.GradFillFromJSON(oParsedFill.fill);
					break;
				case "pattFill":
					oFillObj.fill = this.PattFillFromJSON(oParsedFill.fill);
					break;
				case "grp":
					oFillObj.fill = new AscFormat.CGrpFill();
					break;
			}
		}

		var nFillType = -1;
		switch (oParsedFill.fillType)
		{
			case "none":
				nFillType = Asc.c_oAscFill.FILL_TYPE_NONE;
				break;
			case "blip":
				nFillType = Asc.c_oAscFill.FILL_TYPE_BLIP;
				break;
			case "noFill":
				nFillType = Asc.c_oAscFill.FILL_TYPE_NOFILL;
				break;
			case "solid":
				nFillType = Asc.c_oAscFill.FILL_TYPE_SOLID;
				break;
			case "grad":
				nFillType = Asc.c_oAscFill.FILL_TYPE_GRAD;
				break;
			case "patt":
				nFillType = Asc.c_oAscFill.FILL_TYPE_PATT;
				break;
			case "grp":
				nFillType = Asc.c_oAscFill.FILL_TYPE_GRP;
				break;
		}

		oFillObj.transparent = oParsedFill.transparent;
		if (nFillType !== -1)
			oFillObj.type = nFillType;

		return oFillObj;
	};
	ReaderFromJSON.prototype.GradFillFromJSON = function(oParsedGradFill)
	{
		var oGradFill = new AscFormat.CGradFill();

		for (var nGs = 0; nGs < oParsedGradFill.gsLst.length; nGs++)
			oGradFill.colors.push(this.GradStopFromJSON(oParsedGradFill.gsLst[nGs]));

		if (oParsedGradFill.lin)
		{
			oGradFill.lin  = new AscFormat.GradLin();
			oGradFill.lin.angle = oParsedGradFill.lin.ang;
			oGradFill.lin.scale = oParsedGradFill.lin.scaled;
		}
		if (oParsedGradFill.path)
		{
			oGradFill.path = new AscFormat.GradPath();
			var nPathShadeType = undefined;
			if (oGradFill.path)
			{
				switch(oParsedGradFill.path.path)
				{
					case "circle":
						nPathShadeType = 0;
						break;
					case "rect":
						nPathShadeType = 1;
						break;
					case "shape":
						nPathShadeType = 2;
						break;
				}
			}
			oGradFill.path.path = nPathShadeType;

			if (oGradFill.path.fillToRect)
			{
				oGradFill.rect      = new AscFormat.CSrcRect();

				oGradFill.rect.b = oParsedGradFill.path.fillToRect.b;
				oGradFill.rect.l = oParsedGradFill.path.fillToRect.l;
				oGradFill.rect.r = oParsedGradFill.path.fillToRect.r;
				oGradFill.rect.t = oParsedGradFill.path.fillToRect.t;
			}
		}

		oGradFill.rotWithShape = oParsedGradFill.rotWithShape;
		
		return oGradFill;
	};
	ReaderFromJSON.prototype.GradStopFromJSON = function(oParsedGradStop)
	{
		var oGs   = new AscFormat.CGs();
		oGs.color = this.ColorFromJSON(oParsedGradStop.color);
		oGs.pos   = oParsedGradStop.pos;

		return oGs;
	};
	ReaderFromJSON.prototype.PattFillFromJSON = function(oParsedPattFill)
	{
		var oPattFill = new AscFormat.CPattFill();

		oPattFill.bgClr  = oParsedPattFill.bgClr ? this.ColorFromJSON(oParsedPattFill.bgClr) : oPattFill.bgClr;
		oPattFill.fgClr  = oParsedPattFill.fgClr ? this.ColorFromJSON(oParsedPattFill.fgClr) : oPattFill.fgClr;
		oPattFill.ftype  = this.GetPresetNumType(oParsedPattFill.prst);

		return oPattFill;
	};
	ReaderFromJSON.prototype.ColorFromJSON = function(oParsedColor)
	{
		var oRGBA = oParsedColor.rgba ? {
			R:   oParsedColor.rgba.red,
			G:   oParsedColor.rgba.green,
			B:   oParsedColor.rgba.blue,
			A:   oParsedColor.rgba.alpha 
		} : oParsedColor.rgba;

		var oColorObj  = new AscFormat.CUniColor();
		oColorObj.RGBA = oRGBA;

		if (oParsedColor.color && oParsedColor.color.type)
		{
			switch (oParsedColor.color.type)
			{
				case "none":
					oColorObj.color = null;
					break;
				case "srgb":
				{
					oColorObj.color   = new AscFormat.CRGBColor();
					oColorObj.color.RGBA.R = oParsedColor.color.rgba.red;
					oColorObj.color.RGBA.G = oParsedColor.color.rgba.green;
					oColorObj.color.RGBA.B = oParsedColor.color.rgba.blue;
					oColorObj.color.RGBA.A = oParsedColor.color.rgba.alpha;


					break;
				}
				case "prst":
				{
					oColorObj.color    = new AscFormat.CPrstColor();
					oColorObj.color.RGBA.R  = oParsedColor.color.rgba.red;
					oColorObj.color.RGBA.G  = oParsedColor.color.rgba.green;
					oColorObj.color.RGBA.B  = oParsedColor.color.rgba.blue;
					oColorObj.color.RGBA.A  = oParsedColor.color.rgba.alpha;
					oColorObj.color.id = oParsedColor.color.id;
					break;
				}
				case "scheme":
				{
					oColorObj.color    = new AscFormat.CSchemeColor();
					oColorObj.color.RGBA.R  = oParsedColor.color.rgba.red;
					oColorObj.color.RGBA.G  = oParsedColor.color.rgba.green;
					oColorObj.color.RGBA.B  = oParsedColor.color.rgba.blue;
					oColorObj.color.RGBA.A  = oParsedColor.color.rgba.alpha;
					oColorObj.color.id = oParsedColor.color.id;
					break;
				}
				case "sys":
				{
					oColorObj.color    = new AscFormat.CSysColor();
					oColorObj.color.RGBA.R  = oParsedColor.color.rgba.red;
					oColorObj.color.RGBA.G  = oParsedColor.color.rgba.green;
					oColorObj.color.RGBA.B  = oParsedColor.color.rgba.blue;
					oColorObj.color.RGBA.A  = oParsedColor.color.rgba.alpha;
					oColorObj.color.id = oParsedColor.color.id;
					break;
				}
				case "style":
				{
					oColorObj.color       = new AscFormat.CStyleColor();
					oColorObj.color.bAuto = oParsedColor.color.auto;
					oColorObj.color.val   = oParsedColor.color.val;
					break;
				}
			}
		}

		oColorObj.Mods = oParsedColor.mods ? this.ColorModifiersFromJSON(oParsedColor.mods) : oColorObj.Mods;

		return oColorObj;
	};
	ReaderFromJSON.prototype.ColorModifiersFromJSON = function(oParsedModifiers)
	{
		var oMods = new AscFormat.CColorModifiers();

		for (var nMod = 0; nMod < oParsedModifiers.length; nMod++)
		{
			var oMod  = new AscFormat.CColorMod();
			oMod.name = oParsedModifiers[nMod].name;
			oMod.val  = oParsedModifiers[nMod].val;

			oMods.Mods.push(oMod);
		}

		return oMods;
	};
	ReaderFromJSON.prototype.BlipFillFromJSON = function(oParsedFill)
	{
		var oBlipFill = new AscFormat.CBlipFill();

		if (oParsedFill.srcRect)
		{
			oBlipFill.srcRect   = new AscFormat.CSrcRect();
			oBlipFill.srcRect.b = oParsedFill.srcRect.b;
			oBlipFill.srcRect.l = oParsedFill.srcRect.l;
			oBlipFill.srcRect.r = oParsedFill.srcRect.r;
			oBlipFill.srcRect.t = oParsedFill.srcRect.t;
		}
		oBlipFill.stretch = oParsedFill.stretch;

		if (oParsedFill.tile)
		{
			oBlipFill.tile      = new AscFormat.CBlipFillTile();
			oBlipFill.tile.tx   = oParsedFill.tx;
			oBlipFill.tile.ty   = oParsedFill.ty;
			oBlipFill.tile.sx   = oParsedFill.sx;
			oBlipFill.tile.sy   = oParsedFill.sy;
			oBlipFill.tile.flip = oParsedFill.flip;
			oBlipFill.tile.algn = GetRectAlgnNumType(oParsedFill.algn);
		}

		for (var nEffect = 0; nEffect < oParsedFill.blip.length; nEffect++)
			oBlipFill.Effects.push(this.EffectFromJSON(oParsedFill.blip[nEffect]));

		oParsedFill.rotWithShape != undefined && oBlipFill.setRotWithShape(oParsedFill.rotWithShape);
		oParsedFill.rasterImageId != undefined &&  oBlipFill.setRasterImageId(oParsedFill.rasterImageId);

		return oBlipFill;
	};
	ReaderFromJSON.prototype.EffectFromJSON = function(oParsedEff)
	{
		var oEffect = null;
		var nBlendType;

		switch (oParsedEff.type)
		{
			case "alphaBiLvl":
				oEffect              = new AscFormat.CAlphaBiLevel();
				oEffect.tresh        = oParsedEff.thresh;
				return oEffect;
			case "alphaCeiling":
				oEffect              = new AscFormat.CAlphaCeiling();
				return oEffect;
			case "alphaFloor":
				oEffect              = new AscFormat.CAlphaFloor();
				return oEffect;
			case "alphaInv":
				oEffect              = new AscFormat.CAlphaInv();
				oEffect.unicolor     = this.ColorFromJSON(oParsedEff.color);
				return oEffect;
			case "alphaMod":
				oEffect              = new AscFormat.CAlphaMod();
				oEffect.cont         = this.EffectContainerFromJSON(oParsedEff.cont);
				return oEffect;
			case "alphaModFix":
				oEffect              = new AscFormat.CAlphaModFix();
				oEffect.amt          = oParsedEff.amt;
				return oEffect;
			case "alphaOutset":
				oEffect              = new AscFormat.CAlphaOutset();
				oEffect.rad          = oParsedEff.rad;
				return oEffect;
			case "alphaRepl":
				oEffect              = new AscFormat.CAlphaRepl();
				oEffect.a            = oParsedEff.a;
				return oEffect;
			case "biLvl":
				oEffect              = new AscFormat.CBiLevel();
				oEffect.thresh       = oParsedEff.thresh;
				return oEffect;
			case "blend":
				nBlendType       = From_XML_ST_BlendMode(oParsedEff.blend);
				
				oEffect              = new AscFormat.CBlend();
				oEffect.cont         = this.EffectContainerFromJSON(oParsedEff.cont);
				oEffect.blend        = nBlendType;
				return oEffect;
			case "blur":
				oEffect              = new AscFormat.CBlur();
				oEffect.grow         = oParsedEff.grow;
				oEffect.rad          = oParsedEff.rad;
				return oEffect;
			case "clrChange":
				oEffect              = new AscFormat.CClrChange();
				oEffect.clrFrom      = this.ColorFromJSON(oParsedEff.clrFrom);
				oEffect.clrTo        = this.ColorFromJSON(oParsedEff.clrTo);
				oEffect.useA         = oParsedEff.useA;
				return oEffect;
			case "clrRepl":
				oEffect              = new AscFormat.CClrRepl();
				oEffect.color        = this.ColorFromJSON(oParsedEff.color);
				return oEffect;
			case "effCont":
				oEffect              = new AscFormat.CEffectContainer();
				oEffect.cont         = this.EffectContainerFromJSON(oParsedEff.cont);
				return oEffect;
			case "duotone":
				oEffect              = new AscFormat.CDuotone();
				for (var nColor = 0; nColor < oParsedEff.colors.length; nColor++)
					oEffect.colors.push(this.ColorFromJSON(oParsedEff.colors[nColor]));
				return oEffect;
			case "effect":
				oEffect              = new AscFormat.CEffectElement();
				oEffect.ref          = oParsedEff.ref;
				return oEffect;
			case "fill":
				oEffect              = new AscFormat.CFillEffect();
				oEffect.fill         = this.FillFromJSON(oParsedEff.fill);
				return oEffect;
			case "fillOvrl":
				nBlendType       = From_XML_ST_BlendMode(oParsedEff.blend);
				
				oEffect              = new AscFormat.CFillOverlay();
				oEffect.fill         = this.FillFromJSON(oParsedEff.fill);
				oEffect.blend        = nBlendType;
				return oEffect;
			case "glow":
				oEffect              = new AscFormat.CGlow();
				oEffect.color        = this.ColorFromJSON(oParsedEff.color);
				oEffect.rad          = oParsedEff.rad;
				return oEffect;
			case "gray":
				oEffect              = new AscFormat.CGrayscl();
				return oEffect;
			case "hsl":
				oEffect              = new AscFormat.CHslEffect();
				oEffect.h            = oParsedEff.hue;
				oEffect.l            = oParsedEff.lum;
				oEffect.s            = oParsedEff.sat;
				return oEffect;
			case "innerShdw":
				oEffect              = new AscFormat.CInnerShdw();
				oEffect.color        = this.ColorFromJSON(oParsedEff.color);
				oEffect.blurRad      = oParsedEff.blurRad;
				oEffect.dir          = oParsedEff.dir;
				oEffect.dist         = oParsedEff.dist;
				return oEffect;
			case "lum":
				oEffect              = new AscFormat.CLumEffect();
				oEffect.bright       = oParsedEff.bright;
				oEffect.contrast     = oParsedEff.contrast;
				return oEffect;
			case "outerShdw":
				oEffect              = new AscFormat.COuterShdw();
				oEffect.color        = this.ColorFromJSON(oParsedEff.color);
				oEffect.algn         = GetRectAlgnNumType(oParsedEff.algn);
				oEffect.blurRad      = oParsedEff.blurRad;
				oEffect.dir          = oParsedEff.dir;
				oEffect.dist         = oParsedEff.dist;
				oEffect.kx           = oParsedEff.kx;
				oEffect.ky           = oParsedEff.ky;
				oEffect.rotWithShape = oParsedEff.rotWithShape;
				oEffect.sx           = oParsedEff.sx;
				oEffect.sy           = oParsedEff.sy;
				return oEffect;
			case "prstShdw":
				var nPrstType = undefined;
				switch (oParsedEff.prst)
				{
					case "shdw1":
						nPrstType = c_oAscPresetShadowVal.shdw1;
						break;
					case "shdw2":
						nPrstType = c_oAscPresetShadowVal.shdw2;
						break;
					case "shdw3":
						nPrstType = c_oAscPresetShadowVal.shdw3;
						break;
					case "shdw4":
						nPrstType = c_oAscPresetShadowVal.shdw4;
						break;
					case "shdw5":
						nPrstType = c_oAscPresetShadowVal.shdw5;
						break;
					case "shdw6":
						nPrstType = c_oAscPresetShadowVal.shdw6;
						break;
					case "shdw7":
						nPrstType = c_oAscPresetShadowVal.shdw7;
						break;
					case "shdw8":
						nPrstType = c_oAscPresetShadowVal.shdw8;
						break;
					case "shdw9":
						nPrstType = c_oAscPresetShadowVal.shdw9;
						break;
					case "shdw10":
						nPrstType = c_oAscPresetShadowVal.shdw10;
						break;
					case "shdw11":
						nPrstType = c_oAscPresetShadowVal.shdw11;
						break;
					case "shdw12":
						nPrstType = c_oAscPresetShadowVal.shdw12;
						break;
					case "shdw13":
						nPrstType = c_oAscPresetShadowVal.shdw13;
						break;
					case "shdw14":
						nPrstType = c_oAscPresetShadowVal.shdw14;
						break;
					case "shdw15":
						nPrstType = c_oAscPresetShadowVal.shdw15;
						break;
					case "shdw16":
						nPrstType = c_oAscPresetShadowVal.shdw16;
						break;
					case "shdw17":
						nPrstType = c_oAscPresetShadowVal.shdw17;
						break;
					case "shdw18":
						nPrstType = c_oAscPresetShadowVal.shdw18;
						break;
					case "shdw19":
						nPrstType = c_oAscPresetShadowVal.shdw19;
						break;
					case "shdw20":
						nPrstType = c_oAscPresetShadowVal.shdw20;
						break;
				}

				oEffect              = new AscFormat.CPrstShdw();
				oEffect.color        = this.ColorFromJSON(oParsedEff.color);
				oEffect.dir          = oParsedEff.dir;
				oEffect.dis          = oParsedEff.dist;
				oEffect.prst         = nPrstType;
				return oEffect;
			case "reflection":
				oEffect              = new AscFormat.CReflection();
				oEffect.algn         = GetRectAlgnNumType(oParsedEff.algn);
				oEffect.blurRad      = oParsedEff.blurRad;
				oEffect.dir          = oParsedEff.dir;
				oEffect.dist         = oParsedEff.dist;
				oEffect.endA         = oParsedEff.endA;
				oEffect.endPos       = oParsedEff.endPos;
				oEffect.fadeDir      = oParsedEff.fadeDir;
				oEffect.kx           = oParsedEff.kx;
				oEffect.ky           = oParsedEff.ky;
				oEffect.rotWithShape = oParsedEff.rotWithShape;
				oEffect.stA          = oParsedEff.stA;
				oEffect.stPos        = oParsedEff.stPos;
				oEffect.sx           = oParsedEff.sx;
				oEffect.sy           = oParsedEff.sy;
				return oEffect;
			case "relOff":
				oEffect              = new AscFormat.CRelOff();
				oEffect.tx           = oParsedEff.tx;
				oEffect.ty           = oParsedEff.ty;
				return oEffect;
			case "softEdge":
				oEffect              = new AscFormat.CSoftEdge();
				oEffect.rad          = oParsedEff.rad;
				return oEffect;
			case "tint":
				oEffect              = new AscFormat.CTintEffect();
				oEffect.amt          = oParsedEff.amt;
				oEffect.hue          = oParsedEff.hue;
				return oEffect;
			case "xfrm":
				oEffect              = new AscFormat.CXfrmEffect();
				oEffect.kx           = oParsedEff.kx;
				oEffect.ky           = oParsedEff.ky;
				oEffect.sx           = oParsedEff.sx;
				oEffect.sy           = oParsedEff.sy;
				oEffect.tx           = oParsedEff.tx;
				oEffect.ty           = oParsedEff.ty;
				return oEffect;
		}
	};
	ReaderFromJSON.prototype.EffectContainerFromJSON = function(oParsedCont)
	{
		var oEffectContainer = new AscFormat.CEffectContainer();

		oEffectContainer.type = oParsedCont.type;
		oEffectContainer.name = oParsedCont.name;

		for (var nEffect = 0; nEffect < oParsedCont.effectList.length; nEffect++)
			oEffectContainer.push(this.EffectFromJSON(oParsedCont.effectList[nEffect]));

		return oEffectContainer;
	};
	ReaderFromJSON.prototype.EffectLstFromJSON = function(oParsedLst)
	{
		var oEffectLst = new AscFormat.CEffectLst();

		oEffectLst.blur        = oParsedLst.blur        ? this.EffectFromJSON(oParsedLst.blur)        : oParsedLst.blur;
		oEffectLst.fillOverlay = oParsedLst.fillOverlay ? this.EffectFromJSON(oParsedLst.fillOverlay) : oParsedLst.fillOverlay;
		oEffectLst.innerShdw   = oParsedLst.innerShdw   ? this.EffectFromJSON(oParsedLst.innerShdw)   : oParsedLst.innerShdw;
		oEffectLst.outerShdw   = oParsedLst.outerShdw   ? this.EffectFromJSON(oParsedLst.outerShdw)   : oParsedLst.outerShdw;
		oEffectLst.prstShdw    = oParsedLst.prstShdw    ? this.EffectFromJSON(oParsedLst.prstShdw)    : oParsedLst.prstShdw;
		oEffectLst.reflection  = oParsedLst.reflection  ? this.EffectFromJSON(oParsedLst.reflection)  : oParsedLst.reflection;
		oEffectLst.softEdge    = oParsedLst.softEdge    ? this.EffectFromJSON(oParsedLst.softEdge)    : oParsedLst.softEdge;

		return oEffectLst;
	};
	ReaderFromJSON.prototype.StyleFromJSON = function(oParsedStyle)
	{
		var sStyleName       = oParsedStyle.name;
		var nBasedOnId       = oParsedStyle.basedOn;
		var nNextId          = oParsedStyle.next;
		var nStyleType       = styletype_Paragraph;
		var bNoCreateTablePr = !oParsedStyle.tblStylePr;

		switch (oParsedStyle.styleType)
		{
			case "paragraphStyle":
				nStyleType = styletype_Paragraph;
				break;
			case "numberingStyle":
				nStyleType = styletype_Numbering;
				break;
			case "tableStyle":
				nStyleType = styletype_Table;
				break;
			case "characterStyle":
				nStyleType = styletype_Character;
				break;
		}
		var oStyle = new CStyle(sStyleName, nBasedOnId, nNextId, nStyleType, bNoCreateTablePr);

		oParsedStyle.link != undefined && oStyle.SetLink(oParsedStyle.link);
		oParsedStyle.customStyle != undefined && oStyle.SetCustom(oParsedStyle.customStyle);
		oParsedStyle.qFormat != undefined && oStyle.SetQFormat(oParsedStyle.qFormat);
		oParsedStyle.uiPriority != undefined && oStyle.SetUiPriority(oParsedStyle.uiPriority);
		oParsedStyle.hidden != undefined && oStyle.SetHidden(oParsedStyle.hidden);
		oParsedStyle.semiHidden != undefined && oStyle.SetSemiHidden(oParsedStyle.semiHidden);
		oParsedStyle.unhideWhenUsed != undefined && oStyle.SetUnhideWhenUsed(oParsedStyle.unhideWhenUsed);
		oParsedStyle.rPr != undefined && oStyle.SetTextPr(this.TextPrFromJSON(oParsedStyle.rPr));
		oParsedStyle.pPr != undefined && oStyle.SetParaPr(this.ParaPrFromJSON(oParsedStyle.pPr));
		oParsedStyle.tblPr != undefined && oStyle.Set_TablePr(this.TablePrFromJSON(null, oParsedStyle.tblPr));
		oParsedStyle.trPr != undefined && oStyle.Set_TableRowPr(this.TableRowPrFromJSON(oParsedStyle.trPr));
		oParsedStyle.tcPr != undefined && oStyle.Set_TableCellPr(this.TableCellPrFromJSON(oParsedStyle.tcPr));

		if (!bNoCreateTablePr)
		{
			oParsedStyle.tblStylePr.band1Horz && oStyle.Set_TableBand1Horz(this.TableStylePrFromJSON(oParsedStyle.tblStylePr.band1Horz));
			oParsedStyle.tblStylePr.band1Vert && oStyle.Set_TableBand1Vert(this.TableStylePrFromJSON(oParsedStyle.tblStylePr.band1Vert));
			oParsedStyle.tblStylePr.band2Horz && oStyle.Set_TableBand2Horz(this.TableStylePrFromJSON(oParsedStyle.tblStylePr.band2Horz));
			oParsedStyle.tblStylePr.band2Vert && oStyle.Set_TableBand2Vert(this.TableStylePrFromJSON(oParsedStyle.tblStylePr.band2Vert));
			oParsedStyle.tblStylePr.firstCol && oStyle.Set_TableFirstCol(this.TableStylePrFromJSON(oParsedStyle.tblStylePr.firstCol));
			oParsedStyle.tblStylePr.firstRow && oStyle.Set_TableFirstRow(this.TableStylePrFromJSON(oParsedStyle.tblStylePr.firstRow));
			oParsedStyle.tblStylePr.lastCol && oStyle.Set_TableLastCol(this.TableStylePrFromJSON(oParsedStyle.tblStylePr.lastCol));
			oParsedStyle.tblStylePr.lastRow && oStyle.Set_TableLastRow(this.TableStylePrFromJSON(oParsedStyle.tblStylePr.lastRow));
			oParsedStyle.tblStylePr.neCell && oStyle.Set_TableTLCell(this.TableStylePrFromJSON(oParsedStyle.tblStylePr.neCell));
			oParsedStyle.tblStylePr.nwCell && oStyle.Set_TableTRCell(this.TableStylePrFromJSON(oParsedStyle.tblStylePr.nwCell));
			oParsedStyle.tblStylePr.seCell && oStyle.Set_TableBLCell(this.TableStylePrFromJSON(oParsedStyle.tblStylePr.seCell));
			oParsedStyle.tblStylePr.swCell && oStyle.Set_TableBRCell(this.TableStylePrFromJSON(oParsedStyle.tblStylePr.swCell));
			oParsedStyle.tblStylePr.wholeTable && oStyle.Set_TableWholeTable(this.TableStylePrFromJSON(oParsedStyle.tblStylePr.wholeTable));
		}

		return oStyle;
	};
	ReaderFromJSON.prototype.TableStylePrFromJSON = function(oParsedPr)
	{
		var oTableStylePr = new CTableStylePr();

		oTableStylePr.TextPr      = oParsedPr.rPr   ? this.TextPrFromJSON(oParsedPr.rPr)          : oParsedPr.rPr;
		oTableStylePr.ParaPr      = oParsedPr.pPr   ? this.ParaPrFromJSON(oParsedPr.pPr)          : oParsedPr.pPr;
		oTableStylePr.TablePr     = oParsedPr.tblPr ? this.TablePrFromJSON(null, oParsedPr.tblPr) : oParsedPr.tblPr;
		oTableStylePr.TableRowPr  = oParsedPr.trPr  ? this.TableRowPrFromJSON(oParsedPr.trPr)     : oParsedPr.trPr;
		oTableStylePr.TableCellPr = oParsedPr.tcPr  ? this.TableCellPrFromJSON(oParsedPr.tcPr)    : oParsedPr.tcPr;

		return oTableStylePr;
	};
	ReaderFromJSON.prototype.ParaDrawingFromJSON = function(oParsedDrawing)
	{
		var oDrawing = new ParaDrawing(undefined, undefined, null, private_GetDrawingDocument(), private_GetLogicDocument(), null);
		
		// doc prop
		oDrawing.docPr.setFromOther(this.CNvPrFromJSON(oParsedDrawing.docPr));

		// effect extent
		oDrawing.setEffectExtent(oParsedDrawing.effectExtent.l, oParsedDrawing.effectExtent.t, oParsedDrawing.effectExtent.r, oParsedDrawing.effectExtent.b);

		// extent
		oDrawing.setExtent(private_EMU2MM(oParsedDrawing.extent.cx), private_EMU2MM(oParsedDrawing.extent.cy));

		// posH posY
		var oPosH = this.PositionHFromJSON(oParsedDrawing.positionH);
		var oPosV = this.PositionVFromJSON(oParsedDrawing.positionV);
		oDrawing.Set_PositionH(oPosH.RelativeFrom, oPosH.Align, oPosH.Value, oPosH.Percent);
		oDrawing.Set_PositionV(oPosV.RelativeFrom, oPosV.Align, oPosV.Value, oPosV.Percent);

		// simple pos
		oDrawing.setSimplePos(oParsedDrawing.simplePos.use, private_EMU2MM(oParsedDrawing.simplePos.x), private_EMU2MM(oParsedDrawing.simplePos.x));

		// distance
		oDrawing.Set_Distance(private_EMU2MM(oParsedDrawing.distL), private_EMU2MM(oParsedDrawing.distT), private_EMU2MM(oParsedDrawing.distR), private_EMU2MM(oParsedDrawing.distB));

		// overlap
		oDrawing.Set_AllowOverlap(oParsedDrawing.allowOverlap);

		// behind doc
		oDrawing.Set_BehindDoc(oParsedDrawing.behindDoc);

		oDrawing.Hidden         = oParsedDrawing.hidden;
		oDrawing.Set_LayoutInCell(oParsedDrawing.layoutInCell);
		oDrawing.Set_Locked(oParsedDrawing.locked);
		oDrawing.Set_RelativeHeight(oParsedDrawing.relativeHeight);
		oDrawing.Set_WrappingType(this.GetWrapNumType(oParsedDrawing.wrapType));
		oDrawing.Set_DrawingType(oParsedDrawing.drawingType === "inline" ? drawing_Inline : drawing_Anchor);

		var oGraphicObj = this.GraphicObjFromJSON(oParsedDrawing.graphic, oDrawing);

		if (!oDrawing.GraphicObj)
			oDrawing.Set_GraphicObject(oGraphicObj);
//		oDrawing.CheckWH();

		return oDrawing;
	};
	ReaderFromJSON.prototype.GraphicObjFromJSON = function(oParsedObj, oParent)
	{
		var oGraphicObj = null;
		switch (oParsedObj.type)
		{
			case "image":
				oGraphicObj = this.ImageFromJSON(oParsedObj, oParent);
				break;
			case "shape":
			case "connectShape":
				oGraphicObj = this.ShapeFromJSON(oParsedObj, oParent);
				break;
			case "chartSpace":
				oGraphicObj = this.ChartSpaceFromJSON(oParsedObj, oParent);
				break;
			case "graphicFrame":
				oGraphicObj = this.GraphicFrameFromJSON(oParsedObj);
				break;
			case "smartArt":
				oGraphicObj = this.SmartArtFromJSON(oParsedObj, oParent);
				break;
			case "groupShape":
				oGraphicObj = this.GroupShapeFromJSON(oParsedObj, oParent);
				break;
			case "lockedCanvas":
				oGraphicObj = this.LockedCanvasFromJSON(oParsedObj, oParent);
				break;
			case "oleObject":
				oGraphicObj = this.OleObjectFromJSOM(oParsedObj, oParent);
				break;
		}

		this.drawingsMap[oParsedObj.id] = oGraphicObj;
		
		return oGraphicObj;
	};
	ReaderFromJSON.prototype.LockedCanvasFromJSON = function(oParsedLocCanvas, oParentDrawing)
	{
		var oLockedCanvas = new AscFormat.CLockedCanvas();

		if (oParentDrawing)
		{
			oLockedCanvas.setParent(oParentDrawing);
			oParentDrawing.Set_GraphicObject(oLockedCanvas);
		}

		for (var nDrawing = 0; nDrawing < oParsedLocCanvas.spTree.length; nDrawing++)
			oLockedCanvas.addToSpTree(oLockedCanvas.spTree.length, this.GraphicObjFromJSON(oParsedLocCanvas.spTree[nDrawing]));

		oParsedLocCanvas.spPr && oLockedCanvas.setSpPr(this.SpPrFromJSON(oParsedLocCanvas.spPr, oLockedCanvas));

		oLockedCanvas.setBDeleted(false);
		return oLockedCanvas;
	};
	ReaderFromJSON.prototype.GroupShapeFromJSON = function(oParsedGrpShp, oParentDrawing)
	{
		var oGroupShape = new AscFormat.CGroupShape();
		var oTempGraphObj = null;
		
		if (oParentDrawing)
		{
			oGroupShape.setParent(oParentDrawing);
			oParentDrawing.Set_GraphicObject(oGroupShape);
		}

		for (var nDrawing = 0; nDrawing < oParsedGrpShp.spTree.length; nDrawing++)
		{
			oTempGraphObj = this.GraphicObjFromJSON(oParsedGrpShp.spTree[nDrawing]);
			oGroupShape.addToSpTree(oGroupShape.spTree.length, oTempGraphObj);
			oTempGraphObj.setGroup(oGroupShape);
		}
			
		oParsedGrpShp.spPr && oGroupShape.setSpPr(this.SpPrFromJSON(oParsedGrpShp.spPr, oGroupShape));

		oGroupShape.setBDeleted(false);
		return oGroupShape;
	};
	ReaderFromJSON.prototype.SmartArtFromJSON = function(oParsedArt, oParentDrawing)
	{
		var oSmartArt = new AscFormat.SmartArt();

		oParsedArt.artType != undefined && oSmartArt.setType(oParsedArt.artType);
		oParsedArt.colorsDef && oSmartArt.setColorsDef(this.ColorsDefFromJSON(oParsedArt.colorsDef));
		oParsedArt.dataModel && oSmartArt.setDataModel(this.DataFromJSON(oParsedArt.dataModel));
		oParsedArt.layoutDef && oSmartArt.setLayoutDef(this.LayoutDefFromJSON(oParsedArt.layoutDef));
		oParsedArt.styleDef && oSmartArt.setStyleDef(this.StyleDefFromJSON(oParsedArt.styleDef));

		oParsedArt.drawing && oSmartArt.setDrawing(this.DrawingFromJSON(oParsedArt.drawing));
		if(oSmartArt.drawing)
		{
			oSmartArt.drawing.setGroup(oSmartArt);
			oSmartArt.addToSpTree(0, oSmartArt.drawing);
		}
		oParsedArt.nvGrpSpPr && oSmartArt.setNvGrpSpPr(this.UniNvPrFromJSON(oParsedArt.nvGrpSpPr));
		oParsedArt.spPr && oSmartArt.setSpPr(this.SpPrFromJSON(oParsedArt.spPr, oSmartArt));
		
		if (oParentDrawing)
		{
			oSmartArt.setParent(oParentDrawing);
			oParentDrawing.Set_GraphicObject(oSmartArt);
		}

		oSmartArt.setBDeleted(false);
		//oSmartArt.recalculate();
		return oSmartArt;
	};
	ReaderFromJSON.prototype.DrawingFromJSON = function(oParsedDrawing)
	{
		var oDrawing = new AscFormat.Drawing();
		var oTempGraphObj = null; 

		for (var nDrawing = 0; nDrawing < oParsedDrawing.spTree.length; nDrawing++)
		{
			oTempGraphObj = this.GraphicObjFromJSON(oParsedDrawing.spTree[nDrawing]);
			oDrawing.addToSpTree(oDrawing.spTree.length, oTempGraphObj);
			oTempGraphObj.setGroup(oDrawing);
		}

		oParsedDrawing.spPr && oDrawing.setSpPr(this.SpPrFromJSON(oParsedDrawing.spPr, oDrawing));

		oDrawing.setBDeleted(false);
		//oDrawing.recalculate();
		return oDrawing;
	};
	ReaderFromJSON.prototype.StyleDefFromJSON = function(oParsedStyleDef)
	{
		var oStyleDef = new AscFormat.StyleDef();

		for (var nStyle = 0; nStyle < oParsedStyleDef.styleLbl.length; nStyle++)
			oStyleDef.addToLstStyleLbl(oStyleDef.styleLbl.length, this.DefStyleLblFromJSON(oParsedStyleDef.styleLbl[nStyle]));

		oParsedStyleDef.catLst && oStyleDef.setCatLst(this.CatLstFromJSON(oParsedStyleDef.catLst));
		oParsedStyleDef.desc && oStyleDef.setDesc(this.DescFromJSON(oParsedStyleDef.desc));
		oParsedStyleDef.scene3d && oStyleDef.setScene3d(this.Scene3dFromJSON(oParsedStyleDef.scene3d));
		oParsedStyleDef.minVer != undefined && oStyleDef.setMinVer(oParsedStyleDef.minVer);
		oParsedStyleDef.title && oStyleDef.setTitle(this.DescFromJSON(oParsedStyleDef.title));
		oParsedStyleDef.uniqueId != undefined && oStyleDef.setUniqueId(oParsedStyleDef.uniqueId);

		return oStyleDef;
	};
	ReaderFromJSON.prototype.DefStyleLblFromJSON = function(oParsedDefStyleLbl)
	{
		var oDefStyleLbl = new AscFormat.StyleDefStyleLbl();

		oParsedDefStyleLbl.scene3d && oDefStyleLbl.setScene3d(this.Scene3dFromJSON(oParsedDefStyleLbl.scene3d));
		oParsedDefStyleLbl.sp3d && oDefStyleLbl.setSp3d(this.Sp3DFromJSON(oParsedDefStyleLbl.sp3d));
		oParsedDefStyleLbl.style && oDefStyleLbl.setStyle(this.SpStyleFromJSON(oParsedDefStyleLbl.style));
		oParsedDefStyleLbl.txPr && oDefStyleLbl.setTxPr(this.TxPrFromJSON(oParsedDefStyleLbl.txPr));
		oParsedDefStyleLbl.name != undefined && oDefStyleLbl.setName(oParsedDefStyleLbl.name);

		return oDefStyleLbl;
	};
	ReaderFromJSON.prototype.Sp3DFromJSON = function(oParsedSp3D)
	{
		var oSp3D = new AscFormat.Sp3d();

		oParsedSp3D.bevelB && oSp3D.setBevelB(this.BevelFromJSON(oParsedSp3D.bevelB));
		oParsedSp3D.bevelT && oSp3D.setBevelT(this.BevelFromJSON(oParsedSp3D.bevelT));
		oParsedSp3D.contourClr && oSp3D.setContourClr(this.ContourClrFromJSON(oParsedSp3D.contourClr));
		oParsedSp3D.extrusionClr && oSp3D.setExtrusionClr(this.ExtrusionClrFromJSON(oParsedSp3D.extrusionClr));
		oParsedSp3D.contourW != undefined && oSp3D.setContourW(oParsedSp3D.contourW);
		oParsedSp3D.extrusionH != undefined && oSp3D.setExtrusionH(oParsedSp3D.extrusionH);
		oParsedSp3D.prstMaterial != undefined && oSp3D.setPrstMaterial(From_XML_ST_PresetMaterialType(oParsedSp3D.prstMaterial));
		oParsedSp3D.z != undefined && oSp3D.setZ(oParsedSp3D.z);

		return oSp3D;
	};
	ReaderFromJSON.prototype.ExtrusionClrFromJSON = function(oParsedExtrusionClr)
	{
		var oExtrusionClr = new AscFormat.ExtrusionClr();

		oParsedExtrusionClr.color && oExtrusionClr.setColor(this.ColorFromJSON(oParsedExtrusionClr.color));

		return oExtrusionClr;
	};
	ReaderFromJSON.prototype.ContourClrFromJSON = function(oParsedContourClr)
	{
		var oContourClr = new AscFormat.ContourClr();

		oParsedContourClr.color && oContourClr.setColor(this.ColorFromJSON(oParsedContourClr.color));

		return oContourClr;
	};
	ReaderFromJSON.prototype.BevelFromJSON = function(oParsedBevel)
	{
		var oBevel = new AscFormat.Bevel();

		oParsedBevel.h != undefined && oBevel.setH(oParsedBevel.h);
		oParsedBevel.w != undefined && oBevel.setW(oParsedBevel.w);
		oParsedBevel.prst != undefined && oBevel.setPrst(From_XML_ST_BevelPresetType(oParsedBevel.prst));

		return oBevel;
	};
	ReaderFromJSON.prototype.Scene3dFromJSON = function(oParsedScene3d)
	{
		var oScene3d = new AscFormat.Scene3d();

		oParsedScene3d.backdrop && oScene3d.setBackdrop(this.BackdropFromJSON(oParsedScene3d.backdrop));
		oParsedScene3d.camera && oScene3d.setCamera(this.CameraFromJSON(oParsedScene3d.camera));
		oParsedScene3d.lightRig && oScene3d.setLightRig(this.LightRigFromJSON(oParsedScene3d.lightRig));

		return oScene3d;
	};
	ReaderFromJSON.prototype.LightRigFromJSON = function(oParsedLightRig)
	{
		var oLightRig = new AscFormat.LightRig();

		oParsedLightRig.rot && oLightRig.setRot(this.CameraRotFromJSON(oParsedLightRig.rot));
		oParsedLightRig.dir != undefined && oLightRig.setDir(From_XML_ST_LightRigDirection(oParsedLightRig.dir));
		oParsedLightRig.rig != undefined && oLightRig.setRig(From_XML_ST_LightRigType(oParsedLightRig.rig));

		return oLightRig;
	};
	ReaderFromJSON.prototype.CameraFromJSON = function(oParsedCamera)
	{
		var oCamera = new AscFormat.Camera();

		oParsedCamera.rot && oCamera.setRot(this.CameraRotFromJSON(oParsedCamera.rot));
		oParsedCamera.fov != undefined && oCamera.setFov(oParsedCamera.fov);
		oParsedCamera.prst != undefined && oCamera.setPrst(From_XML_ST_PresetCameraType(oParsedCamera.prst));
		oParsedCamera.zoom != undefined && oCamera.setZoom(oParsedCamera.zoom);
		
		return oCamera;
	};
	ReaderFromJSON.prototype.CameraRotFromJSON = function(oParsedCameraRot)
	{
		var oCameraRot = new AscFormat.Rot();

		oParsedCameraRot.lat != undefined && oCameraRot.setLat(oParsedCameraRot.lat);
		oParsedCameraRot.lon != undefined && oCameraRot.setLon(oParsedCameraRot.lon);
		oParsedCameraRot.rev != undefined && oCameraRot.setRev(oParsedCameraRot.rev);

		return oCameraRot;
	};
	ReaderFromJSON.prototype.BackdropFromJSON = function(oParsedBackdrop)
	{
		var oBackdrop = new AscFormat.Backdrop();

		oBackdrop.setAnchor(this.BackdropAnchorFromJSON(oParsedBackdrop.anchor));
		oBackdrop.setNorm(this.BackdropNormFromJSON(oParsedBackdrop.norm));
		oBackdrop.setUp(this.BackdropUpFromJSON(oParsedBackdrop.up));

		return oBackdrop;
	};
	ReaderFromJSON.prototype.BackdropUpFromJSON = function(oParsedBackdropUp)
	{
		var oBackdropUp = new AscFormat.BackdropUp();

		oParsedBackdropUp.dx != undefined && oBackdropUp.setDx(oParsedBackdropUp.dx);
		oParsedBackdropUp.dy != undefined && oBackdropUp.setDy(oParsedBackdropUp.dy);
		oParsedBackdropUp.dz != undefined && oBackdropUp.setDz(oParsedBackdropUp.dz);

		return oBackdropUp;
	};
	ReaderFromJSON.prototype.BackdropNormFromJSON = function(oParsedBackdropNorm)
	{
		var oBackdropNorm = new AscFormat.BackdropNorm();

		oParsedBackdropNorm.dx != undefined && oBackdropNorm.setDx(oParsedBackdropNorm.dx);
		oParsedBackdropNorm.dy != undefined && oBackdropNorm.setDy(oParsedBackdropNorm.dy);
		oParsedBackdropNorm.dz != undefined && oBackdropNorm.setDz(oParsedBackdropNorm.dz);

		return oBackdropNorm;
	};
	ReaderFromJSON.prototype.BackdropAnchorFromJSON = function(oParsedBackdropAnchor)
	{
		var oBackdropAnchor = new AscFormat.BackdropAnchor();

		oParsedBackdropAnchor.x != undefined && oBackdropAnchor.setX(oParsedBackdropAnchor.x);
		oParsedBackdropAnchor.y != undefined && oBackdropAnchor.setY(oParsedBackdropAnchor.y);
		oParsedBackdropAnchor.z != undefined && oBackdropAnchor.setZ(oParsedBackdropAnchor.z);

		return oBackdropAnchor;
	};
	ReaderFromJSON.prototype.LayoutDefFromJSON = function(oParsedLayoutDef)
	{
		var oLayoutDef = new AscFormat.LayoutDef();

		oParsedLayoutDef.catLst && oLayoutDef.setCatLst(this.CatLstFromJSON(oParsedLayoutDef.catLst));
		oParsedLayoutDef.clrData && oLayoutDef.setClrData(this.DataFromJSON(oParsedLayoutDef.clrData));
		oParsedLayoutDef.defStyle && oLayoutDef.setDefStyle(oParsedLayoutDef.defStyle);
		oParsedLayoutDef.desc && oLayoutDef.setDesc(this.DescFromJSON(oParsedLayoutDef.desc));
		oParsedLayoutDef.layoutNode && oLayoutDef.setLayoutNode(this.LayoutNodeFromJSON(oParsedLayoutDef.layoutNode));
		oParsedLayoutDef.minVer && oLayoutDef.setMinVer(oParsedLayoutDef.minVer);
		oParsedLayoutDef.sampData && oLayoutDef.setSampData(this.DataFromJSON(oParsedLayoutDef.sampData));
		oParsedLayoutDef.styleData && oLayoutDef.setStyleData(this.DataFromJSON(oParsedLayoutDef.styleData));
		oParsedLayoutDef.title && oLayoutDef.setTitle(this.DescFromJSON(oParsedLayoutDef.title));
		oParsedLayoutDef.uniqueId && oLayoutDef.setUniqueId(oParsedLayoutDef.uniqueId);

		return oLayoutDef;
	};
	ReaderFromJSON.prototype.DataFromJSON = function(oParsedData)
	{
		var oData = null;
		switch (oParsedData.type)
		{
			case "diagramData":
				oData = new AscFormat.DiagramData();
				break;
			case "styleData":
				oData = new AscFormat.StyleData();
				oParsedData.useDef != undefined && oData.setUseDef(oParsedData.useDef);
				break;
			case "clrData":
				oData = new AscFormat.ClrData();
				oParsedData.useDef != undefined && oData.setUseDef(oParsedData.useDef);
				break;
			case "sampData":
				oData = new AscFormat.SampData();
				oParsedData.useDef != undefined && oData.setUseDef(oParsedData.useDef);
				break;
		}

		if (oData)
			oParsedData.dataModel && oData.setDataModel(this.DataModelFromJSON(oParsedData.dataModel));

		return oData;
	};
	ReaderFromJSON.prototype.DataModelFromJSON = function(oParsedDataModel)
	{
		var oDataModel = new AscFormat.DataModel();

		oParsedDataModel.bg && oDataModel.setBg(this.BgFormatFromJSON(oParsedDataModel.bg));
		oParsedDataModel.cxnLst && oDataModel.setCxnLst(this.CxnLstFromJSON(oParsedDataModel.cxnLst));
		oParsedDataModel.ptLst && oDataModel.setPtLst(this.PtLstFromJSON(oParsedDataModel.ptLst));
		oParsedDataModel.whole && oDataModel.setWhole(this.WholeFromJSON(oParsedDataModel.whole));

		return oDataModel;
	};
	ReaderFromJSON.prototype.WholeFromJSON = function(oParsedWhole)
	{
		var oWhole = new AscFormat.Whole();

		(oParsedWhole.effectDag || oParsedWhole.effectLst) && oWhole.setEffect(this.EffectPropsFromJSON(oParsedWhole.effectDag, oParsedWhole.effectLst));
		oParsedWhole.whole && oWhole.setLn(this.LnFromJSON(oParsedWhole.whole));

		return oWhole;
	};
	ReaderFromJSON.prototype.PtLstFromJSON = function(oParsedPtLst)
	{
		var oPtLst = new AscFormat.PtLst();

		for (var nItem = 0; nItem < oParsedPtLst.list.length; nItem++)
			oPtLst.addToLst(oPtLst.list.length, this.PtFromJSON(oParsedPtLst.list[nItem]));

		return oPtLst;
	};
	ReaderFromJSON.prototype.PtFromJSON = function(oParsedPt)
	{
		var oPt = new AscFormat.Point();

		oParsedPt.prSet && oPt.setPrSet(this.PtPrSetFromJSON(oParsedPt.prSet));
		oParsedPt.spPr && oPt.setSpPr(this.SpPrFromJSON(oParsedPt.spPr));
		oParsedPt.t && oPt.setT(this.TxPrFromJSON(oParsedPt.t));
		oParsedPt.cxnId != undefined && oPt.setCxnId(oParsedPt.cxnId);
		oParsedPt.modelId != undefined && oPt.setModelId(oParsedPt.modelId);
		oParsedPt.type != undefined && oPt.setType(From_XML_ST_PtType(oParsedPt.type));

		return oPt;
	};
	ReaderFromJSON.prototype.PtPrSetFromJSON = function(oParsedPtPrSet)
	{
		var oPtPrSet = new AscFormat.PrSet();

		oParsedPtPrSet.coherent3DOff != undefined && oPtPrSet.setCoherent3DOff(oParsedPtPrSet.coherent3DOff);
		oParsedPtPrSet.csCatId != undefined && oPtPrSet.setCsCatId(oParsedPtPrSet.csCatId);
		oParsedPtPrSet.csTypeId != undefined && oPtPrSet.setCsTypeId(oParsedPtPrSet.csTypeId);
		oParsedPtPrSet.custAng != undefined && oPtPrSet.setCustAng(oParsedPtPrSet.custAng);
		oParsedPtPrSet.custFlipHor != undefined && oPtPrSet.setCustFlipHor(oParsedPtPrSet.custFlipHor);
		oParsedPtPrSet.custFlipVert != undefined && oPtPrSet.setCustFlipVert(oParsedPtPrSet.custFlipVert);
		oParsedPtPrSet.custLinFactNeighborX != undefined && oPtPrSet.setCustLinFactNeighborX(oParsedPtPrSet.custLinFactNeighborX);
		oParsedPtPrSet.custLinFactNeighborY != undefined && oPtPrSet.setCustLinFactNeighborY(oParsedPtPrSet.custLinFactNeighborY);
		oParsedPtPrSet.custLinFactX != undefined && oPtPrSet.setCustLinFactX(oParsedPtPrSet.custLinFactX);
		oParsedPtPrSet.custLinFactY != undefined && oPtPrSet.setCustLinFactY(oParsedPtPrSet.custLinFactY);
		oParsedPtPrSet.custRadScaleInc != undefined && oPtPrSet.setCustRadScaleInc(oParsedPtPrSet.custRadScaleInc);
		oParsedPtPrSet.custRadScaleRad != undefined && oPtPrSet.setCustRadScaleRad(oParsedPtPrSet.custRadScaleRad);
		oParsedPtPrSet.custScaleX != undefined && oPtPrSet.setCustScaleX(oParsedPtPrSet.custScaleX);
		oParsedPtPrSet.custScaleY != undefined && oPtPrSet.setCustScaleY(oParsedPtPrSet.custScaleY);
		oParsedPtPrSet.custSzX != undefined && oPtPrSet.setCustSzX(oParsedPtPrSet.custSzX);
		oParsedPtPrSet.custSzY != undefined && oPtPrSet.setCustSzY(oParsedPtPrSet.custSzY);
		oParsedPtPrSet.custT != undefined && oPtPrSet.setCustT(oParsedPtPrSet.custT);
		oParsedPtPrSet.loCatId != undefined && oPtPrSet.setLoCatId(oParsedPtPrSet.loCatId);
		oParsedPtPrSet.loTypeId != undefined && oPtPrSet.setLoTypeId(oParsedPtPrSet.loTypeId);
		oParsedPtPrSet.phldr != undefined && oPtPrSet.setPhldr(oParsedPtPrSet.phldr);
		oParsedPtPrSet.phldrT != undefined && oPtPrSet.setPhldrT(oParsedPtPrSet.phldrT);
		oParsedPtPrSet.presAssocID != undefined && oPtPrSet.setPresAssocID(oParsedPtPrSet.presAssocID);
		oParsedPtPrSet.presLayoutVars != undefined && oPtPrSet.setPresLayoutVars(this.NodeItemFromJSON(oParsedPtPrSet.presLayoutVars));
		oParsedPtPrSet.presName != undefined && oPtPrSet.setPresName(oParsedPtPrSet.presName);
		oParsedPtPrSet.presStyleCnt != undefined && oPtPrSet.setPresStyleCnt(oParsedPtPrSet.presStyleCnt);
		oParsedPtPrSet.presStyleIdx != undefined && oPtPrSet.setPresStyleIdx(oParsedPtPrSet.presStyleIdx);
		oParsedPtPrSet.presStyleLbl != undefined && oPtPrSet.setPresStyleLbl(oParsedPtPrSet.presStyleLbl);
		oParsedPtPrSet.qsCatId != undefined && oPtPrSet.setQsCatId(oParsedPtPrSet.qsCatId);
		oParsedPtPrSet.qsTypeId != undefined && oPtPrSet.setQsTypeId(oParsedPtPrSet.qsTypeId);
		oParsedPtPrSet.style != undefined && oPtPrSet.setStyle(this.DataFromJSON(oParsedPtPrSet.style));

		return oPtPrSet;
	};
	ReaderFromJSON.prototype.NodeItemFromJSON = function(oParsedNodeItem)
	{
		switch (oParsedNodeItem.objType)
		{
			case "alg":
				return this.AlgFromJSON(oParsedNodeItem);
			case "choose":
				return this.ChooseFromJSON(oParsedNodeItem);
			case "constrLst":
				return this.ConstrLstFromJSON(oParsedNodeItem);
			case "forEach":
				return this.ForEachFromJSON(oParsedNodeItem);
			case "layoutNode":
				return this.LayoutNodeFromJSON(oParsedNodeItem);
			case "presOf":
				return this.PresOfFromJSON(oParsedNodeItem);
			case "ruleLst":
				return this.RuleLstFromJSON(oParsedNodeItem);
			case "sshape":
				return this.SShapeFromJSON(oParsedNodeItem);
			case "varLst":
				return this.VarLstFromJSON(oParsedNodeItem);
			default:
				return null;
		}
	};
	ReaderFromJSON.prototype.SShapeFromJSON = function(oParsedSShape)
	{
		var oSShape = new AscFormat.SShape();

		oParsedSShape.blip != undefined && oSShape.setBlip(oParsedSShape.blip);
		oParsedSShape.blipPhldr != undefined && oSShape.setBlipPhldr(oParsedSShape.blipPhldr);
		oParsedSShape.hideGeom != undefined && oSShape.setHideGeom(oParsedSShape.hideGeom);
		oParsedSShape.lkTxEntry != undefined && oSShape.setLkTxEntry(oParsedSShape.lkTxEntry);
		oParsedSShape.rot != undefined && oSShape.setRot(oParsedSShape.rot);
		oParsedSShape.type != undefined && oSShape.setType(oParsedSShape.type);
		oParsedSShape.zOrderOff != undefined && oSShape.setZOrderOff(oParsedSShape.zOrderOff);
		oParsedSShape.adjLst != undefined && oSShape.setAdjLst(this.AdjLstFromJSON(oParsedSShape.adjLst));
				
		return oSShape;
	};
	ReaderFromJSON.prototype.AdjLstFromJSON = function(oParsedAdjLst)
	{
		var oAdjLst = new AscFormat.AdjLst();

		for (var nItem = 0; nItem < oParsedAdjLst.list.length; nItem++)
			oAdjLst.addToLst(oAdjLst.list.length, this.AdjFromJSON(oParsedAdjLst.list[nItem]));
		
		return oAdjLst;
	};
	ReaderFromJSON.prototype.AdjFromJSON = function(oParsedAdj)
	{
		var oAdj = new AscFormat.Adj();

		oParsedAdj.idx != undefined && oAdj.setIdx(oParsedAdj.idx);
		oParsedAdj.val != undefined && oAdj.setVal(oParsedAdj.val);

		return oAdj;
	};
	ReaderFromJSON.prototype.RuleLstFromJSON = function(oParsedRuleLst)
	{
		var oRuleLst = new AscFormat.RuleLst();

		for (var nItem = 0; nItem < oParsedRuleLst.list.length; nItem++)
			oRuleLst.addToLst(oRuleLst.list.length, this.RuleFromJSON(oParsedRuleLst.list[nItem]));

		return oRuleLst;
	};
	ReaderFromJSON.prototype.RuleFromJSON = function(oParsedRule)
	{
		var oRule = new AscFormat.Rule();

		var fact = oParsedRule.fact === "Nan" ? NaN : oParsedRule.fact;
		var max = oParsedRule.max === "Nan" ? NaN : oParsedRule.max;

		fact != undefined && oRule.setFact(fact);
		oParsedRule.for != undefined && oRule.setFor(From_XML_ST_ConstraintRelationship(oParsedRule.for));
		oParsedRule.forName != undefined && oRule.setForName(oParsedRule.forName);
		max != undefined && oRule.setMax(max);
		oParsedRule.ptType != undefined && oRule.setPtType(this.BaseFormatObjFromJSON(oParsedRule.ptType));
		oParsedRule.type != undefined && oRule.setType(From_XML_ST_ConstraintType(oParsedRule.type));
		oParsedRule.val != undefined && oRule.setVal(oParsedRule.val);

		return oRule;
	};
	ReaderFromJSON.prototype.PresOfFromJSON = function(oParsedPresOf)
	{
		var oPresOf = new AscFormat.PresOf();

		this.IteratorAttributesFromJSON(oParsedPresOf, oPresOf);

		return oPresOf;
	};
	ReaderFromJSON.prototype.LayoutNodeFromJSON = function(oParsedLayoutNode)
	{
		var oLayoutNode = new AscFormat.LayoutNode();

		for (var nItem = 0; nItem < oParsedLayoutNode.list.length; nItem++)
			oLayoutNode.addToLst(oLayoutNode.list.length, this.NodeItemFromJSON(oParsedLayoutNode.list[nItem]));

		oParsedLayoutNode.chOrder != undefined && oLayoutNode.setChOrder(From_XML_ST_ChildOrderType(oParsedLayoutNode.chOrder));
		oParsedLayoutNode.moveWith != undefined && oLayoutNode.setMoveWith(oParsedLayoutNode.moveWith);
		oParsedLayoutNode.name != undefined && oLayoutNode.setName(oParsedLayoutNode.name);
		oParsedLayoutNode.styleLbl != undefined && oLayoutNode.setStyleLbl(oParsedLayoutNode.styleLbl);
		
		return oLayoutNode;
	};
	ReaderFromJSON.prototype.ForEachFromJSON = function(oParsedForEach)
	{
		var oForEach = new AscFormat.ForEach();

		for (var nItem = 0; nItem < oParsedForEach.list.length; nItem++)
			oForEach.addToLstList(oForEach.list.length, this.NodeItemFromJSON(oParsedForEach.list[nItem]));

		oParsedForEach.name != undefined && oForEach.setName(oParsedForEach.name);
		oParsedForEach.ref != undefined && oForEach.setRef(oParsedForEach.ref);

		this.IteratorAttributesFromJSON(oParsedForEach, oForEach);

		return oForEach;
	};
	ReaderFromJSON.prototype.ConstrLstFromJSON = function(oParsedConstrLst)
	{
		var oConstrLst = new AscFormat.ConstrLst();

		for (var nItem = 0; nItem < oParsedConstrLst.list.length; nItem++)
			oConstrLst.addToLst(oConstrLst.list.length, this.ConstrFromJSON(oParsedConstrLst.list[nItem]));
			
		return oConstrLst;
	};
	ReaderFromJSON.prototype.ConstrFromJSON = function(oParsedConstr)
	{
		var oConstr = new AscFormat.Constr();

		oParsedConstr.fact != undefined && oConstr.setFact(oParsedConstr.fact);
		oParsedConstr.for != undefined && oConstr.setFor(From_XML_ST_ConstraintRelationship(oParsedConstr.for));
		oParsedConstr.forName != undefined && oConstr.setForName(oParsedConstr.forName);
		oParsedConstr.op != undefined && oConstr.setOp(From_XML_ST_BoolOperator(oParsedConstr.op));
		oParsedConstr.ptType != undefined && oConstr.setPtType(this.BaseFormatObjFromJSON(oParsedConstr.ptType));
		oParsedConstr.refFor != undefined && oConstr.setRefFor(From_XML_ST_ConstraintRelationship(oParsedConstr.refFor));
		oParsedConstr.refForName != undefined && oConstr.setRefForName(oParsedConstr.refForName);
		oParsedConstr.refPtType != undefined && oConstr.setRefPtType(this.BaseFormatObjFromJSON(oParsedConstr.refPtType));
		oParsedConstr.refType != undefined && oConstr.setRefType(From_XML_ST_ConstraintType(oParsedConstr.refType));
		oParsedConstr.type != undefined && oConstr.setType(From_XML_ST_ConstraintType(oParsedConstr.type));
		oParsedConstr.val != undefined && oConstr.setVal(oParsedConstr.val);

		return oConstr;
	};
	ReaderFromJSON.prototype.ChooseFromJSON = function(oParsedChoose)
	{
		var oChoose = new AscFormat.Choose();

		for (var nIf = 0; nIf < oParsedChoose.if.length; nIf++)
			oChoose.addToLstIf(oChoose.if.length, this.IfFromJSON(oParsedChoose.if[nIf]));

		oParsedChoose.else && oChoose.setElse(this.ElseFromJSON(oParsedChoose.else));
		oParsedChoose.name != undefined && oChoose.setName(oParsedChoose.name);
		
		return oChoose;
	};
	ReaderFromJSON.prototype.ElseFromJSON = function(oParsedElse)
	{
		var oElse = new AscFormat.Else();

		for (var nItem = 0; nItem < oParsedElse.list.length; nItem++)
			oElse.addToLst(oElse.list.length, this.NodeItemFromJSON(oParsedElse.list[nItem]));

		oParsedElse.name != undefined && oElse.setName(oParsedElse.name);

		return oElse;
	};
	ReaderFromJSON.prototype.IfFromJSON = function(oParsedIf)
	{
		var oIf = new AscFormat.If();

		for (var nItem = 0; nItem < oParsedIf.list.length; nItem++)
			oIf.addToLstList(oIf.list.length, this.NodeItemFromJSON(oParsedIf.list[nItem]));

		oParsedIf.arg != undefined && oIf.setArg(oParsedIf.arg);
		oParsedIf.func != undefined && oIf.setFunc(From_XML_ST_FunctionType(oParsedIf.func));
		oParsedIf.name != undefined && oIf.setName(oParsedIf.name);
		oParsedIf.op != undefined && oIf.setOp(From_XML_ST_FunctionOperator(oParsedIf.op));
		oParsedIf.val != undefined && oIf.setVal(oParsedIf.val);

		this.IteratorAttributesFromJSON(oParsedIf, oIf);

		return oIf;
	};
	ReaderFromJSON.prototype.IteratorAttributesFromJSON = function(oParsedIterAttr, oParent)
	{
		for (var nAxie = 0; nAxie < oParsedIterAttr.axis.length; nAxie++)
			oParent.addToLstAxis(oParent.axis.length, this.BaseFormatObjFromJSON(oParsedIterAttr.axis[nAxie]));
		
		for (var nCnt = 0; nCnt < oParsedIterAttr.cnt.length; nCnt++)
			oParent.addToLstCnt(oParent.cnt.length, oParsedIterAttr.cnt[nCnt]);
		
		for (var nItem = 0; nItem < oParsedIterAttr.hideLastTrans.length; nItem++)
			oParent.addToLstHideLastTrans(oParent.hideLastTrans.length, oParsedIterAttr.hideLastTrans[nItem]);

		for (var nPtType = 0; nPtType < oParsedIterAttr.ptType.length; nPtType++)
			oParent.addToLstPtType(oParent.ptType.length, this.BaseFormatObjFromJSON(oParsedIterAttr.ptType[nPtType]));

		for (var nSt = 0; nSt < oParsedIterAttr.st.length; nSt++)
			oParent.addToLstSt(oParent.st.length, oParsedIterAttr.st[nSt]);

		for (var nStep = 0; nStep < oParsedIterAttr.step.length; nStep++)
			oParent.addToLstStep(oParent.step.length, oParsedIterAttr.step[nStep]);
	};
	ReaderFromJSON.prototype.AlgFromJSON = function(oParsedAlg)
	{
		var oAlg = new AscFormat.Alg();

		for (var nParam = 0; nParam < oParsedAlg.param.length; nParam++)
			oAlg.addToLstParam(oAlg.param.length, this.AlgParamFromJSON(oParsedAlg.param[nParam]));

		oParsedAlg.rev != undefined && oAlg.setRev(oParsedAlg.rev);
		oParsedAlg.type != undefined && oAlg.setType(From_XML_ST_AlgorithmType(oParsedAlg.type));

		return oAlg;
	};
	ReaderFromJSON.prototype.AlgParamFromJSON = function(oParsedAlgParam)
	{
		var oAlgParam = new AscFormat.Param();

		oParsedAlgParam.type != undefined && oAlgParam.setType(From_XML_ST_ParameterId(oParsedAlgParam.type));
		oParsedAlgParam.val != undefined && oAlgParam.setVal(oParsedAlgParam.val);

		return oAlgParam;
	};
	ReaderFromJSON.prototype.VarLstFromJSON = function(oParsedVarLst)
	{
		var oVarLst = new AscFormat.VarLst();

		oParsedVarLst.animLvl && oVarLst.setAnimLvl(this.BaseFormatObjFromJSON(oParsedVarLst.animLvl));
		oParsedVarLst.animOne && oVarLst.setAnimOne(this.BaseFormatObjFromJSON(oParsedVarLst.animOne));
		oParsedVarLst.bulletEnabled && oVarLst.setBulletEnabled(this.BaseFormatObjFromJSON(oParsedVarLst.bulletEnabled));
		oParsedVarLst.chMax && oVarLst.setChMax(this.BaseFormatObjFromJSON(oParsedVarLst.chMax));
		oParsedVarLst.chPref && oVarLst.setChPref(this.BaseFormatObjFromJSON(oParsedVarLst.chPref));
		oParsedVarLst.dir && oVarLst.setDir(this.BaseFormatObjFromJSON(oParsedVarLst.dir));
		oParsedVarLst.hierBranch && oVarLst.setHierBranch(this.BaseFormatObjFromJSON(oParsedVarLst.hierBranch));
		oParsedVarLst.orgChart && oVarLst.setOrgChart(this.BaseFormatObjFromJSON(oParsedVarLst.orgChart));
		oParsedVarLst.resizeHandles && oVarLst.setResizeHandles(this.BaseFormatObjFromJSON(oParsedVarLst.resizeHandles));

		return oVarLst;
	};
	ReaderFromJSON.prototype.BaseFormatObjFromJSON = function(oParsedBaseFormatObj)
	{
		var oBaseFormatObj = null;

		switch (oParsedBaseFormatObj.type)
		{
			case "animLvl":
				oBaseFormatObj = new AscFormat.AnimLvl();
				oParsedBaseFormatObj.val != undefined && oBaseFormatObj.setVal(From_XML_ST_AnimLvlStr(oParsedBaseFormatObj.val));
				break;
			case "animOne":
				oBaseFormatObj = new AscFormat.AnimOne();
				oParsedBaseFormatObj.val != undefined && oBaseFormatObj.setVal(From_XML_ST_AnimOneStr(oParsedBaseFormatObj.val));
				break;
			case "bulletEnabled":
				oBaseFormatObj = new AscFormat.BulletEnabled();
				oParsedBaseFormatObj.val != undefined && oBaseFormatObj.setVal(oParsedBaseFormatObj.val);
				break;
			case "chMax":
				oBaseFormatObj = new AscFormat.ChMax();
				oParsedBaseFormatObj.val != undefined && oBaseFormatObj.setVal(oParsedBaseFormatObj.val);
				break;
			case "chPref":
				oBaseFormatObj = new AscFormat.ChPref();
				oParsedBaseFormatObj.val != undefined && oBaseFormatObj.setVal(oParsedBaseFormatObj.val);
				break;
			case "orgChart":
				oBaseFormatObj = new AscFormat.OrgChart();
				oParsedBaseFormatObj.val != undefined && oBaseFormatObj.setVal(oParsedBaseFormatObj.val);
				break;
			case "dir":
				oBaseFormatObj = new AscFormat.DiagramDirection();
				oParsedBaseFormatObj.val != undefined && oBaseFormatObj.setVal(From_XML_ST_Direction(oParsedBaseFormatObj.val));
				break;
			case "hierBranch":
				oBaseFormatObj = new AscFormat.HierBranch();
				oParsedBaseFormatObj.val != undefined && oBaseFormatObj.setVal(From_XML_ST_HierBranchStyle(oParsedBaseFormatObj.val));
				break;
			case "resizeHandles":
				oBaseFormatObj = new AscFormat.ResizeHandles();
				oParsedBaseFormatObj.val != undefined && oBaseFormatObj.setVal(From_XML_ST_ResizeHandlesStr(oParsedBaseFormatObj.val));
				break;
			case "element":
				oBaseFormatObj = new AscFormat.ElementType();
				oParsedBaseFormatObj.val != undefined && oBaseFormatObj.setVal(From_XML_ST_ElementType(oParsedBaseFormatObj.val));
				break;
			case "axie":
				oBaseFormatObj = new AscFormat.AxisType();
				oParsedBaseFormatObj.val != undefined && oBaseFormatObj.setVal(From_XML_ST_AxisType(oParsedBaseFormatObj.val));
				break;
		}

		return oBaseFormatObj;
	};
	ReaderFromJSON.prototype.CxnLstFromJSON = function(oParsedCxnLst)
	{
		var oCxnLst = new AscFormat.CxnLst();

		for (var nItem = 0; nItem < oParsedCxnLst.list.length; nItem++)
			oCxnLst.addToLst(oCxnLst.list.length, this.CxnFromJSON(oParsedCxnLst.list[nItem]));

		return oCxnLst;
	};
	ReaderFromJSON.prototype.CxnFromJSON = function(oParsedCxn)
	{
		var oCxn = new AscFormat.Cxn();

		oParsedCxn.destId != undefined && oCxn.setDestId(oParsedCxn.destId);
		oParsedCxn.destOrd != undefined && oCxn.setDestOrd(oParsedCxn.destOrd);
		oParsedCxn.modelId != undefined && oCxn.setModelId(oParsedCxn.modelId);
		oParsedCxn.parTransId != undefined && oCxn.setParTransId(oParsedCxn.parTransId);
		oParsedCxn.presId != undefined && oCxn.setPresId(oParsedCxn.presId);
		oParsedCxn.sibTransId != undefined && oCxn.setSibTransId(oParsedCxn.sibTransId);
		oParsedCxn.srcId != undefined && oCxn.setSrcId(oParsedCxn.srcId);
		oParsedCxn.srcOrd != undefined && oCxn.setSrcOrd(oParsedCxn.srcOrd);
		oParsedCxn.type != undefined && oCxn.setType(oParsedCxn.type);

		return oCxn;
	};
	ReaderFromJSON.prototype.BgFormatFromJSON = function(oParsedBgFormat)
	{
		var oBgFormat = new AscFormat.BgFormat();

		(oParsedBgFormat.effectDag || oParsedBgFormat.effectLst) && oBgFormat.setEffect(this.EffectPropsFromJSON(oParsedBgFormat.effectDag, oParsedBgFormat.effectLst));
		oParsedBgFormat.fill && oBgFormat.setFill(this.FillFromJSON(oParsedBgFormat.fill));

		return oBgFormat;
	};
	ReaderFromJSON.prototype.ColorsDefFromJSON = function(oParsedColorsDef)
	{
		var oColorsDef = new AscFormat.ColorsDef();

		oParsedColorsDef.catLst && oColorsDef.setCatLst(this.CatLstFromJSON(oParsedColorsDef.catLst));
		oParsedColorsDef.desc && oColorsDef.setDesc(this.DescFromJSON(oParsedColorsDef.desc));

		for (var nStyle = 0; nStyle < oParsedColorsDef.styleLbl.length; nStyle++)
			oColorsDef.addToLstStyleLbl(oColorsDef.styleLbl.length, this.ColorDefStyleLblFromJSON(oParsedColorsDef.styleLbl[nStyle]));

		oParsedColorsDef.title && oColorsDef.setTitle(this.DescFromJSON(oParsedColorsDef.title));

		oParsedColorsDef.minVer != undefined && oColorsDef.setMinVer(oParsedColorsDef.minVer);
		oParsedColorsDef.uniqueId != undefined && oColorsDef.setUniqueId(oParsedColorsDef.uniqueId);

		return oColorsDef;
	};
	ReaderFromJSON.prototype.ColorDefStyleLblFromJSON = function(oParsedLbl)
	{
		var oLbl = new AscFormat.ColorDefStyleLbl();

		oParsedLbl.effectClrLst && oLbl.setEffectClrLst(this.ClrLstFromJSON(oParsedLbl.effectClrLst));
		oParsedLbl.fillClrLst && oLbl.setFillClrLst(this.ClrLstFromJSON(oParsedLbl.fillClrLst));
		oParsedLbl.linClrLst && oLbl.setLinClrLst(this.ClrLstFromJSON(oParsedLbl.linClrLst));
		oParsedLbl.txEffectClrLst && oLbl.setTxEffectClrLst(this.ClrLstFromJSON(oParsedLbl.txEffectClrLst));
		oParsedLbl.txFillClrLst && oLbl.setTxFillClrLst(this.ClrLstFromJSON(oParsedLbl.txFillClrLst));
		oParsedLbl.txLinClrLst && oLbl.setTxLinClrLst(this.ClrLstFromJSON(oParsedLbl.txLinClrLst));
		oParsedLbl.name != undefined && oLbl.setName(oParsedLbl.name);

		return oLbl;
	};
	ReaderFromJSON.prototype.ClrLstFromJSON = function(oParsedClrLst)
	{
		var oClrLst = null;

		switch (oParsedClrLst.type)
		{
			case "effectClrLst":
				oClrLst = new AscFormat.EffectClrLst();
				break;
			case "fillClrLst":
				oClrLst = new AscFormat.FillClrLst();
				break;
			case "linClrLst":
				oClrLst = new AscFormat.LinClrLst();
				break;
			case "txEffectClrLst":
				oClrLst = new AscFormat.TxEffectClrLst();
				break;
			case "txFillClrLst":
				oClrLst = new AscFormat.TxFillClrLst();
				break;
			case "txLinClrLst":
				oClrLst = new AscFormat.TxLinClrLst();
				break;
		}

		if (oClrLst)
		{
			oParsedClrLst.hueDir != undefined && oClrLst.setHueDir(From_XML_ST_HueDir(oParsedClrLst.hueDir));
			oParsedClrLst.meth != undefined && oClrLst.setMeth(From_XML_ST_ClrAppMethod(oParsedClrLst.meth));

			for (var nColor = 0; nColor < oParsedClrLst.list.length; nColor++)
				oClrLst.addToLst(oClrLst.list.length, this.ColorFromJSON(oParsedClrLst.list[nColor]));
		}

		return oClrLst;
	};
	ReaderFromJSON.prototype.DescFromJSON = function(oParsedDesc)
	{
		var oDesc = null;
		if (oParsedDesc.type === "desc")
			oDesc = new AscFormat.Desc();
		else if (oParsedDesc.type === "diagramTitle")
			oDesc = new AscFormat.DiagramTitle();

		if (oDesc)
		{
			oParsedDesc.val  != undefined && oDesc.setVal(oParsedDesc.val);
			oParsedDesc.lang != undefined && oDesc.setLang(oParsedDesc.lang);
		}

		return oDesc;
	};
	ReaderFromJSON.prototype.CatLstFromJSON = function(oParsedCatLst)
	{
		var oCatLst = new AscFormat.CatLst();
	
		for (var nCat = 0; nCat < oParsedCatLst.list.length; nCat++)
			oCatLst.addToLst(oCatLst.list.length, this.SCatFromJSON(oParsedCatLst.list[nCat]));

		return oCatLst;
	};
	ReaderFromJSON.prototype.SCatFromJSON = function(oParsedSCat)
	{
		var oSCat = new AscFormat.SCat();

		oParsedSCat.pri != undefined && oSCat.setPri(oParsedSCat.pri);
		oParsedSCat.type != undefined && oSCat.setType(oParsedSCat.type);

		return oSCat;
	};
	ReaderFromJSON.prototype.ChartSpaceFromJSON = function(oParsedChart, oParentDrawing)
	{
		var oChartSpace = new AscFormat.CChartSpace();

		if(!oChartSpace)
			return null;

		oChartSpace.setBDeleted(false);
		oChartSpace.setChart(this.ChartFromJSON(oParsedChart.chart, oChartSpace));
		oChartSpace.setSpPr(this.SpPrFromJSON(oParsedChart.spPr, oChartSpace));
		oParsedChart.chartColors   && oChartSpace.setChartColors(this.ChartColorsFromJSON(oParsedChart.chartColors));
		oParsedChart.chartStyle    && oChartSpace.setChartStyle(this.ChartStyleFromJSON(oParsedChart.chartStyle));
		oParsedChart.clrMapOvr     && oChartSpace.setClrMapOvr(this.ColorMapOvrFromJSON());
		oParsedChart.pivotSource   && oChartSpace.setPivotSource(this.PivotSourceFromJSON(oParsedChart.pivotSource, oChartSpace));
		oParsedChart.printSettings && oChartSpace.setPrintSettings(this.PrintSettingsFromJSON(oParsedChart.printSettings, oChartSpace));
		oParsedChart.protection    && oChartSpace.setProtection(this.ProtectionFromJSON(oParsedChart.protection));
		oParsedChart.txPr          && oChartSpace.setTxPr(this.TxPrFromJSON(oParsedChart.txPr, oChartSpace));
		oChartSpace.setDate1904(oParsedChart.date1904);
		oChartSpace.setLang(oParsedChart.lang);
		oChartSpace.setRoundedCorners(oParsedChart.roundedCorners);
		oChartSpace.setStyle(oParsedChart.style);

		for (var nUserShape = 0; nUserShape < oParsedChart.userShapes.length; nUserShape++)
			oChartSpace.addUserShape(undefined, this.UserShapeFromJSON(oParsedChart.userShapes[nUserShape]));

		if (oParentDrawing)
		{
			oChartSpace.setParent(oParentDrawing);
			oParentDrawing.Set_GraphicObject(oChartSpace);
			oChartSpace.spPr.xfrm && oParentDrawing.setExtent(oChartSpace.spPr.xfrm.extX, oChartSpace.spPr.xfrm.extY);
		}

		oChartSpace.bDeleted = false;
		oChartSpace.updateLinks();
		oChartSpace.recalculate();
		return oChartSpace;
	};
	ReaderFromJSON.prototype.GraphicFrameFromJSON = function(oParsedGraphFrame)
	{
		var oGraphicFrame = new AscFormat.CGraphicFrame();

		oParsedGraphFrame.nvGraphicFramePr && oGraphicFrame.setNvSpPr(this.UniNvPrFromJSON(oParsedGraphFrame.nvGraphicFramePr));
		oParsedGraphFrame.spPr && oGraphicFrame.setSpPr(this.SpPrFromJSON(oParsedGraphFrame.spPr));
		oParsedGraphFrame.graphic && oGraphicFrame.setGraphicObject(this.TableFromJSON(oParsedGraphFrame.graphic, oGraphicFrame));

		oGraphicFrame.setBDeleted(false);
		
		return oGraphicFrame;
	};
	ReaderFromJSON.prototype.ColorMapOvrFromJSON = function(oParsedClrMap)
	{
		var oClrMap = new AscFormat.ClrMap();

		for (var nClrScheme = 0; nClrScheme < oParsedClrMap.length; nClrScheme++)
		{
			var nClrSchemeType = null;
			switch (oParsedClrMap[nClrScheme])
			{
				case "accent1":
					nClrSchemeType = c_oAscColorSchemeIndex.Accent1;
					break;
				case "accent2":
					nClrSchemeType = c_oAscColorSchemeIndex.Accent2;
					break;
				case "accent3":
					nClrSchemeType = c_oAscColorSchemeIndex.Accent3;
					break;
				case "accent4":
					nClrSchemeType = c_oAscColorSchemeIndex.Accent4;
					break;
				case "accent5":
					nClrSchemeType = c_oAscColorSchemeIndex.Accent5;
					break;
				case "accent6":
					nClrSchemeType = c_oAscColorSchemeIndex.Accent6;
					break;
				case "bg1":
					nClrSchemeType = c_oAscColorSchemeIndex.Bg1;
					break;
				case "bg2":
					nClrSchemeType = c_oAscColorSchemeIndex.Bg2;
					break;
				case "dk1":
					nClrSchemeType = c_oAscColorSchemeIndex.Dk1;
					break;
				case "dk2":
					nClrSchemeType = c_oAscColorSchemeIndex.Dk2;
					break;
				case "folHlink":
					nClrSchemeType = c_oAscColorSchemeIndex.FolHlink;
					break;
				case "hlink":
					nClrSchemeType = c_oAscColorSchemeIndex.Hlink;
					break;
				case "lt1":
					nClrSchemeType = c_oAscColorSchemeIndex.Lt1;
					break;
				case "lt2":
					nClrSchemeType = c_oAscColorSchemeIndex.Lt2;
					break;
				case "phClr":
					nClrSchemeType = c_oAscColorSchemeIndex.PhClr;
					break;
				case "tx1":
					nClrSchemeType = c_oAscColorSchemeIndex.Tx1;
					break;
				case "tx2":
					nClrSchemeType = c_oAscColorSchemeIndex.Tx2;
					break;
			}
			nClrSchemeType != null && oClrMap.setClr(nClrScheme, nClrSchemeType);
		}

		return oClrMap;
	};
	ReaderFromJSON.prototype.UserShapeFromJSON = function(oParsedUserShape)
	{
		var oUserShape  = oParsedUserShape.type === "relSizeAnchor" ? new AscFormat.CRelSizeAnchor() : new AscFormat.CAbsSizeAnchor();
		var oGraphicObj = this.GraphicObjFromJSON(oParsedUserShape.object);

		oUserShape.setFromTo(private_Twips2MM(oParsedUserShape.fromX), private_Twips2MM(oParsedUserShape.fromY), oParsedUserShape.toX, oParsedUserShape.toY);
		oUserShape.setObject(oGraphicObj);

		return oUserShape;
	};
	ReaderFromJSON.prototype.ChartColorsFromJSON = function(oParsedChartColors)
	{
		var oChartColors = new AscFormat.CChartColors();

		for (var nItem = 0; nItem < oParsedChartColors.items.length; nItem++)
		{
			if (oParsedChartColors.items[nItem].type && oParsedChartColors.items[nItem].type === "uniColor")
				oChartColors.addItem(this.ColorFromJSON(oParsedChartColors.items[nItem]));
			else
				oChartColors.addItem(this.ColorModifiersFromJSON(oParsedChartColors.items[nItem]));
		}

		oChartColors.setId(oParsedChartColors.id);
		oChartColors.setMeth(oParsedChartColors.meth);

		return oChartColors;
	};
	ReaderFromJSON.prototype.ChartStyleFromJSON = function(oParsedChartStyle)
	{
		var oChartStyle = new AscFormat.CChartStyle();

		oChartStyle.setId(oParsedChartStyle.id);
        oParsedChartStyle.axisTitle && oChartStyle.setAxisTitle(this.StyleEntryFromJSON(oParsedChartStyle.axisTitle));
        oParsedChartStyle.categoryAxis && oChartStyle.setCategoryAxis(this.StyleEntryFromJSON(oParsedChartStyle.categoryAxis));
        oParsedChartStyle.chartArea && oChartStyle.setChartArea(this.StyleEntryFromJSON(oParsedChartStyle.chartArea));
        oParsedChartStyle.dataLabel && oChartStyle.setDataLabel(this.StyleEntryFromJSON(oParsedChartStyle.dataLabel));
        oParsedChartStyle.dataLabelCallout && oChartStyle.setDataLabelCallout(this.StyleEntryFromJSON(oParsedChartStyle.dataLabelCallout));
        oParsedChartStyle.dataPoint && oChartStyle.setDataPoint(this.StyleEntryFromJSON(oParsedChartStyle.dataPoint));
        oParsedChartStyle.dataPoint3D && oChartStyle.setDataPoint3D(this.StyleEntryFromJSON(oParsedChartStyle.dataPoint3D));
        oParsedChartStyle.dataPointLine && oChartStyle.setDataPointLine(this.StyleEntryFromJSON(oParsedChartStyle.dataPointLine));
        oParsedChartStyle.dataPointMarker && oChartStyle.setDataPointMarker(this.StyleEntryFromJSON(oParsedChartStyle.dataPointMarker));
        oParsedChartStyle.dataPointWireframe && oChartStyle.setDataPointWireframe(this.StyleEntryFromJSON(oParsedChartStyle.dataPointWireframe));
        oParsedChartStyle.dataTable && oChartStyle.setDataTable(this.StyleEntryFromJSON(oParsedChartStyle.dataTable));
        oParsedChartStyle.downBar && oChartStyle.setDownBar(this.StyleEntryFromJSON(oParsedChartStyle.downBar));
        oParsedChartStyle.dropLine && oChartStyle.setDropLine(this.StyleEntryFromJSON(oParsedChartStyle.dropLine));
        oParsedChartStyle.errorBar && oChartStyle.setErrorBar(this.StyleEntryFromJSON(oParsedChartStyle.errorBar));
        oParsedChartStyle.floor && oChartStyle.setFloor(this.StyleEntryFromJSON(oParsedChartStyle.floor));
        oParsedChartStyle.gridlineMajor && oChartStyle.setGridlineMajor(this.StyleEntryFromJSON(oParsedChartStyle.gridlineMajor));
        oParsedChartStyle.gridlineMinor && oChartStyle.setGridlineMinor(this.StyleEntryFromJSON(oParsedChartStyle.gridlineMinor));
        oParsedChartStyle.hiLoLine && oChartStyle.setHiLoLine(this.StyleEntryFromJSON(oParsedChartStyle.hiLoLine));
        oParsedChartStyle.leaderLine && oChartStyle.setLeaderLine(this.StyleEntryFromJSON(oParsedChartStyle.leaderLine));
        oParsedChartStyle.legend && oChartStyle.setLegend(this.StyleEntryFromJSON(oParsedChartStyle.legend));
        oParsedChartStyle.plotArea && oChartStyle.setPlotArea(this.StyleEntryFromJSON(oParsedChartStyle.plotArea));
        oParsedChartStyle.plotArea3D && oChartStyle.setPlotArea3D(this.StyleEntryFromJSON(oParsedChartStyle.plotArea3D));
        oParsedChartStyle.seriesAxis && oChartStyle.setSeriesAxis(this.StyleEntryFromJSON(oParsedChartStyle.seriesAxis));
        oParsedChartStyle.seriesLine && oChartStyle.setSeriesLine(this.StyleEntryFromJSON(oParsedChartStyle.seriesLine));
        oParsedChartStyle.title && oChartStyle.setTitle(this.StyleEntryFromJSON(oParsedChartStyle.title));
        oParsedChartStyle.trendline && oChartStyle.setTrendline(this.StyleEntryFromJSON(oParsedChartStyle.trendline));
        oParsedChartStyle.trendlineLabel && oChartStyle.setTrendlineLabel(this.StyleEntryFromJSON(oParsedChartStyle.trendlineLabel));
        oParsedChartStyle.upBar && oChartStyle.setUpBar(this.StyleEntryFromJSON(oParsedChartStyle.upBar));
        oParsedChartStyle.valueAxis && oChartStyle.setValueAxis(this.StyleEntryFromJSON(oParsedChartStyle.valueAxis));
        oParsedChartStyle.wall && oChartStyle.setWall(this.StyleEntryFromJSON(oParsedChartStyle.wall));
        oParsedChartStyle.markerLayout && oChartStyle.setMarkerLayout(this.MarkerLayoutFromJSON(oParsedChartStyle.markerLayout));
	
		return oChartStyle;
	};
	ReaderFromJSON.prototype.StyleEntryFromJSON = function(oParsedStyleEntry)
	{
		var oStyleEntry = new AscFormat.CStyleEntry();

		oStyleEntry.setType(oParsedStyleEntry.type);
        oStyleEntry.setLineWidthScale(oParsedStyleEntry.lineWidthScale);
        oParsedStyleEntry.lnRef && oStyleEntry.setLnRef(this.StyleRefFromJSON(oParsedStyleEntry.lnRef));
        oParsedStyleEntry.fillRef && oStyleEntry.setFillRef(this.StyleRefFromJSON(oParsedStyleEntry.fillRef));
        oParsedStyleEntry.effectRef && oStyleEntry.setEffectRef(this.StyleRefFromJSON(oParsedStyleEntry.effectRef));
        oParsedStyleEntry.fontRef && oStyleEntry.setFontRef(this.FontRefFromJSON(oParsedStyleEntry.fontRef));
        oParsedStyleEntry.defRPr && oStyleEntry.setDefRPr(this.TextPrFromJSON(oParsedStyleEntry.defRPr));
        oParsedStyleEntry.bodyPr && oStyleEntry.setBodyPr(this.BodyPrFromJSON(oParsedStyleEntry.bodyPr));
        oParsedStyleEntry.spPr && oStyleEntry.setSpPr(this.SpPrFromJSON(oParsedStyleEntry.spPr, oStyleEntry));

		return oStyleEntry;
	};
	ReaderFromJSON.prototype.MarkerLayoutFromJSON = function(oParsedMarkerLayout)
	{
		var oMarkerLayout = new AscFormat.CMarkerLayout();

		oMarkerLayout.setSize(oParsedMarkerLayout.size);
		oMarkerLayout.setSymbol(oParsedMarkerLayout.symbol);

		return oMarkerLayout;
	};
	ReaderFromJSON.prototype.PivotSourceFromJSON = function(oParsedPivotSource, oParentChart)
	{
		var oPivotSource = new AscFormat.CPivotSource();

		oPivotSource.setParent(oParentChart);

		oPivotSource.setFmtId(oParsedPivotSource.fmtId);
		oPivotSource.setName(oParsedPivotSource.name);

		return oPivotSource;
	};
	ReaderFromJSON.prototype.PrintSettingsFromJSON = function(oParsedPrintSettings, oParentChart)
	{
		var oPrintSettings = new AscFormat.CPrintSettings();

		oPrintSettings.setParent(oParentChart);

		oParsedPrintSettings.headerFooter && oPrintSettings.setHeaderFooter(this.HeaderFooterChartFromJSON(oParsedPrintSettings.headerFooter));
		oParsedPrintSettings.pageMargins && oPrintSettings.setPageMargins(this.PageMarginsChartFromJSON(oParsedPrintSettings.pageMargins));
		oParsedPrintSettings.pageSetup && oPrintSettings.setPageSetup(this.PageSetupFromJSON(oParsedPrintSettings.pageSetup));

		return oPrintSettings;
	};
	ReaderFromJSON.prototype.ProtectionFromJSON = function(oParsedProtection)
	{
		var oProtection = new AscFormat.CProtection();

		oProtection.setChartObject(oParsedProtection.chartObject);
		oProtection.setData(oParsedProtection.data);
		oProtection.setFormatting(oParsedProtection.formatting);
		oProtection.setSelection(oParsedProtection.selection);
		oProtection.setUserInterface(oParsedProtection.userInterface);

		return oProtection;
	};
	ReaderFromJSON.prototype.HeaderFooterChartFromJSON = function(oParsedHdrFtrCart)
	{
		var oHeaderFooter = new AscFormat.CHeaderFooterChart();

		oHeaderFooter.setEvenFooter(oParsedHdrFtrCart.evenFooter);
		oHeaderFooter.setEvenFooter(oParsedHdrFtrCart.evenHeader);
		oHeaderFooter.setFirstFooter(oParsedHdrFtrCart.firstFooter);
		oHeaderFooter.setFirstHeader(oParsedHdrFtrCart.firstHeader);
		oHeaderFooter.setOddFooter(oParsedHdrFtrCart.oddFooter);
		oHeaderFooter.setOddHeader(oParsedHdrFtrCart.oddHeader);
		oHeaderFooter.setAlignWithMargins(oParsedHdrFtrCart.alignWithMargins);
		oHeaderFooter.setDifferentFirst(oParsedHdrFtrCart.differentFirst);
		oHeaderFooter.setDifferentOddEven(oParsedHdrFtrCart.differentOddEven);

		return oHeaderFooter;
	};
	ReaderFromJSON.prototype.PageMarginsChartFromJSON = function(oParsedPgMargins)
	{
		var oPageMargins = new AscFormat.CPageMarginsChart();
		
		oPageMargins.setB(oParsedPgMargins.b);
		oPageMargins.setFooter(oParsedPgMargins.footer);
		oPageMargins.setHeader(oParsedPgMargins.header);
		oPageMargins.setL(oParsedPgMargins.l);
		oPageMargins.setR(oParsedPgMargins.r);
		oPageMargins.setT(oParsedPgMargins.t);

		return oPageMargins;
	};
	ReaderFromJSON.prototype.PageSetupFromJSON = function(oParsedPageSetup)
	{
		var oPageSetup = new AscFormat.CPageSetup();

		var nOrientType = undefined;
		switch(oParsedPageSetup.orientation)
		{
			case "default":
				nOrientType = AscFormat.PAGE_SETUP_ORIENTATION_DEFAULT;
				break;
			case "portrait":
				nOrientType = AscFormat.PAGE_SETUP_ORIENTATION_PORTRAIT;
				break;
			case "landscape":
				nOrientType = AscFormat.PAGE_SETUP_ORIENTATION_LANDSCAPE;
				break;
		}

		oPageSetup.setBlackAndWhite(oParsedPageSetup.blackAndWhite);
		oPageSetup.setCopies(oParsedPageSetup.copies);
		oPageSetup.setDraft(oParsedPageSetup.draft);
		oPageSetup.setFirstPageNumber(oParsedPageSetup.firstPageNumber);
		oPageSetup.setHorizontalDpi(oParsedPageSetup.horizontalDpi);
		oPageSetup.setOrientation(nOrientType);
		oPageSetup.setPaperHeight(oParsedPageSetup.paperHeight);
		oPageSetup.setPaperSize(oParsedPageSetup.paperSize);
		oPageSetup.setPaperWidth(oParsedPageSetup.paperWidth);
		oPageSetup.setUseFirstPageNumb(oParsedPageSetup.useFirstPageNumb);
		oPageSetup.setVerticalDpi(oParsedPageSetup.verticalDpi);

		return oPageSetup;
	};
	ReaderFromJSON.prototype.ShapeFromJSON = function(oParsedShape, oParentDrawing)
	{
		var oShape = oParsedShape.type === "shape" ? new AscFormat.CShape() : new AscFormat.CConnectionShape();

		if (oParentDrawing)
		{
			oShape.setParent(oParentDrawing);
			oParentDrawing.Set_GraphicObject(oShape);
		}
		
		oParsedShape.bWordShape && oShape.setWordShape(oParsedShape.bWordShape);
		oParsedShape.nvSpPr && oShape.setNvSpPr(this.UniNvPrFromJSON(oParsedShape.nvSpPr));

		oParsedShape.spPr && oShape.setSpPr(this.SpPrFromJSON(oParsedShape.spPr, oShape));

		oShape.setStyle(this.SpStyleFromJSON(oParsedShape.style));
		oParsedShape.bodyPr && oShape.setBodyPr(this.BodyPrFromJSON(oParsedShape.bodyPr));

		if (oParsedShape.content)
		{
			if (oParsedShape.content.type === "docContent")
				oShape.setTextBoxContent(this.DocContentFromJSON(oParsedShape.content, oShape));
			else
				oShape.setTxBody(this.TxPrFromJSON(oParsedShape.content, oShape));
		}

		oParsedShape.modelId != undefined && oShape.setModelId(oParsedShape.modelId);
		oParsedShape.txXfrm != undefined && oShape.setTxXfrm(this.XfrmFromJSON(oParsedShape.txXfrm));

		oShape.setBDeleted(false);
		oShape.recalculate();

		return oShape;
	};
	ReaderFromJSON.prototype.OleObjectFromJSOM = function(oParsedOleObj, oParentDrawing)
	{
		var oOleObject = new AscFormat.COleObject();
		oOleObject.setBDeleted(false);
		
		if (oParentDrawing)
		{
			oOleObject.setParent(oParentDrawing);
			oParentDrawing.Set_GraphicObject(oOleObject);
		}

		oParsedOleObj.appId      != undefined && oOleObject.setApplicationId(oParsedOleObj.appId);
		oParsedOleObj.data       != undefined && oOleObject.setData(oParsedOleObj.data);
		oParsedOleObj.objFile    != undefined && oOleObject.setObjectFile(oParsedOleObj.objFile);
		oParsedOleObj.oleType    != undefined && oOleObject.setOleType(From_XML_OleObj_Type(oParsedOleObj.oleType));
		oParsedOleObj.binaryData != undefined && oOleObject.setBinaryData(oParsedOleObj.binaryData);
		oParsedOleObj.mathObj    != undefined && oOleObject.setMathObject(this.ParaMathFromJSON(oParsedOleObj.mathObj));
		
		if (oParsedOleObj.dxaOrig > 0 && oParsedOleObj.dyaOrig > 0) {
			oOleObject.setPixSizes(private_Twips2Px(oParsedOleObj.dxaOrig), private_Twips2Px(oParsedOleObj.dyaOrig));
		}

		oOleObject.setBlipFill(this.BlipFillFromJSON(oParsedOleObj.blipFill));
		oOleObject.setNvPicPr(this.UniNvPrFromJSON(oParsedOleObj.nvPicPr));
		oOleObject.setSpPr(this.SpPrFromJSON(oParsedOleObj.spPr, oOleObject));

		return oOleObject;
	};
	ReaderFromJSON.prototype.ImageFromJSON = function(oParsedImage, oParentDrawing)
	{
		var oImage = new AscFormat.CImageShape();

		if (oParentDrawing)
		{
			oImage.setParent(oParentDrawing);
			oParentDrawing.Set_GraphicObject(oImage);
		}

		oImage.setBlipFill(this.BlipFillFromJSON(oParsedImage.blipFill));
		oImage.setNvPicPr(this.UniNvPrFromJSON(oParsedImage.nvPicPr));
		oImage.setSpPr(this.SpPrFromJSON(oParsedImage.spPr, oImage));

		oImage.setNoChangeAspect(true);
        oImage.setBDeleted(false);

		return oImage;
	};
	ReaderFromJSON.prototype.ChartFromJSON = function(oParsedChart, oParentChartSpace)
	{
		var oChart = new AscFormat.CChart();

		var nDispBlanksAs = undefined;
		switch(oParsedChart.dispBlanksAs)
		{
			case "span":
				nDispBlanksAs = AscFormat.DISP_BLANKS_AS_SPAN;
				break;
			case "gap":
				nDispBlanksAs = AscFormat.DISP_BLANKS_AS_GAP;
				break;
			case "zero":
				nDispBlanksAs = AscFormat.DISP_BLANKS_AS_ZERO;
		}

		oParentChartSpace && oChart.setParent(oParentChartSpace);

		oChart.setPlotArea(this.PlotAreaFromJSON(oParsedChart.plotArea));
		oChart.setAutoTitleDeleted(oParsedChart.autoTitleDeleted);
		oParsedChart.backWall && oChart.setBackWall(this.WallFromJSON(oParsedChart.backWall, oChart));
		oChart.setDispBlanksAs(nDispBlanksAs);
		oParsedChart.floor && oChart.setFloor(this.WallFromJSON(oParsedChart.floor, oChart));
		oParsedChart.legend && oChart.setLegend(this.LegendFromJSON(oParsedChart.legend, oChart));
		this.PivotFmtsFromJSON(oParsedChart.pivotFmts, oChart);
		oChart.setPlotVisOnly(oParsedChart.plotVisOnly);
		oChart.setShowDLblsOverMax(oParsedChart.showDLblsOverMax);
		oParsedChart.sideWall && oChart.setSideWall(this.WallFromJSON(oParsedChart.sideWall, oChart));
		oParsedChart.title && oChart.setTitle(this.TitleFromJSON(oParsedChart.title, oChart));
		oParsedChart.view3D && oChart.setView3D(this.View3DFromJSON(oParsedChart.view3D));

		return oChart;
	};
	ReaderFromJSON.prototype.View3DFromJSON = function(oParsedView3D)
	{
		var oView3D = new AscFormat.CView3d();

		oView3D.setDepthPercent(oParsedView3D.depthPercent);
		oView3D.setHPercent(oParsedView3D.hPercent);
		oView3D.setPerspective(oParsedView3D.perspective);
		oView3D.setRAngAx(oParsedView3D.rAngAx);
		oView3D.setRotX(oParsedView3D.rotX);
		oView3D.setRotY(oParsedView3D.rotY);

		return oView3D;
	};
	ReaderFromJSON.prototype.PlotAreaFromJSON = function(oParsedArea)
	{
		var oPlotArea = new AscFormat.CPlotArea();

		var oAxisMap = {};
		
		for (var nAxis = 0; nAxis < oParsedArea.axId.length; nAxis++)
		{
			var oAxis = this.AxisFromJSON(oParsedArea.axId[nAxis]);
			oPlotArea.addAxis(oAxis);
			oAxisMap[oParsedArea.axId[nAxis].axId] = oAxis;
		}

		oParsedArea.dTable && oPlotArea.setDTable(this.DataTableFromJSON(oParsedArea.dTable));
		oParsedArea.layout && oPlotArea.setLayout(this.LayoutFromJSON(oParsedArea.layout, oPlotArea));
		oParsedArea.spPr && oPlotArea.setSpPr(this.SpPrFromJSON(oParsedArea.spPr, oPlotArea));

		for (nAxis = 0; nAxis < oPlotArea.axId.length; nAxis++)
			oPlotArea.axId[nAxis].setCrossAx(oAxisMap[oPlotArea.axId[nAxis].crossAx]);

		this.ChartsFromJSON(oParsedArea.charts, oPlotArea, oAxisMap);

		return oPlotArea;
	};
	ReaderFromJSON.prototype.AxisFromJSON = function(oParsedAxis)
	{
		switch (oParsedAxis.type)
		{
			case "catAx":
				return this.CatAxFromJSON(oParsedAxis);
			case "valAx":
				return this.ValAxFromJSON(oParsedAxis);
			case "dateAx":
				return this.DateAxFromJSON(oParsedAxis);
			case "serAx":
				return this.SerAxFromJSON(oParsedAxis);
		}

		return null;
	};
	ReaderFromJSON.prototype.DataTableFromJSON = function(oParsedDataTable)
	{
		var oDTable = new AscFormat.CDTable();

		oDTable.setShowHorzBorder(oParsedDataTable.showHorzBorder);
		oDTable.setShowKeys(oParsedDataTable.showKeys);
		oDTable.setShowOutline(oParsedDataTable.showOutline);
		oDTable.setShowVertBorder(oParsedDataTable.showVertBorder);
		oParsedDataTable.spPr && oDTable.setSpPr(this.SpPrFromJSON(oParsedDataTable.spPr, oDTable));
		oParsedDataTable.txPr && oDTable.setTxPr(this.TxPrFromJSON(oParsedDataTable.txPr, oDTable));
	
		return oDTable;
	};
	ReaderFromJSON.prototype.SerAxFromJSON = function(oParsedSerAx, oParentPlotArea)
	{
		var oSerAx = new AscFormat.CSerAx();

		oSerAx.setParent(oParentPlotArea);

		oSerAx.setAxId(++AscFormat.Ax_Counter.GLOBAL_AX_ID_COUNTER);
		oSerAx.setAxPos(this.GetAxPosNumType(oParsedSerAx.axPos));
		oSerAx.crossAx        = oParsedSerAx.crossAx ? oParsedSerAx.crossAx : oSerAx.crossAx;
		oSerAx.setCrosses(this.GetCrossesNumType(oParsedSerAx.crosses));
		oSerAx.setCrossesAt(oParsedSerAx.crossesAt);
		oSerAx.setDelete(oParsedSerAx.delete);
		oSerAx.extLst = oParsedSerAx.extLst; /// ?
		oParsedSerAx.majorGridlines && oSerAx.setMajorGridlines(this.SpPrFromJSON(oParsedSerAx.majorGridlines, oSerAx));
		oSerAx.setMajorTickMark(this.GetTickMarkNumType(oParsedSerAx.majorTickMark));
		oParsedSerAx.minorGridlines && oSerAx.setMinorGridlines(this.SpPrFromJSON(oParsedSerAx.minorGridlines, oSerAx));
		oSerAx.setMinorTickMark(this.GetTickMarkNumType(oParsedSerAx.minorTickMark));
		oParsedSerAx.numFmt && oSerAx.setNumFmt(this.NumFmtFromJSON(oParsedSerAx.numFmt));
		oParsedSerAx.scaling && oSerAx.setScaling(this.ScalingFromJSON(oParsedSerAx.scaling, oSerAx));
		oParsedSerAx.spPr && oSerAx.setSpPr(this.SpPrFromJSON(oParsedSerAx.spPr, oSerAx));
		oSerAx.setTickLblSkip(this.GetTickLabelNumPos(oParsedSerAx.tickLblPos));
		oSerAx.setTickLblSkip(oParsedSerAx.tickLblSkip);
		oSerAx.setTickMarkSkip(oParsedSerAx.tickMarkSkip);
		oParsedSerAx.title && oSerAx.setTitle(this.TitleFromJSON(oParsedSerAx.title));
		oParsedSerAx.txPr && oSerAx.setTxPr(this.TxPrFromJSON(oParsedSerAx.txPr, oSerAx));
		
		return oSerAx;
	};
	ReaderFromJSON.prototype.DateAxFromJSON = function(oParsedDateAx, oParentPlotArea)
	{
		var oDateAx = new AscFormat.CDateAx();

		oDateAx.setParent(oParentPlotArea);

		oDateAx.setAuto(oParsedDateAx.auto);
		oDateAx.setAxId(++AscFormat.Ax_Counter.GLOBAL_AX_ID_COUNTER);
		oDateAx.setAxPos(this.GetAxPosNumType(oParsedDateAx.axPos));
		oDateAx.setBaseTimeUnit(this.GetTimeUnitNumType(oParsedDateAx.baseTimeUnit));
		oDateAx.crossAx = oParsedDateAx.crossAx ? oParsedDateAx.crossAx : oDateAx.crossAx;
		oDateAx.setCrosses(this.GetCrossesNumType(oParsedDateAx.crosses));
		oDateAx.setCrossesAt(oParsedDateAx.crossesAt);
		oDateAx.setDelete(oParsedDateAx.delete);
		oDateAx.extLst = oParsedDateAx.extLst; /// ???
		oDateAx.setLblOffset(oParsedDateAx.lblOffset);
		oParsedDateAx.majorGridlines && oDateAx.setMajorGridlines(this.SpPrFromJSON(oParsedDateAx.majorGridlines, oDateAx));
		oDateAx.setMajorTickMark(this.GetTickMarkNumType(oParsedDateAx.majorTickMark));
		oDateAx.setMajorTimeUnit(this.GetTimeUnitNumType(oParsedDateAx.majorTimeUnit));
		oDateAx.setMajorUnit(oParsedDateAx.majorUnit);
		oParsedDateAx.minorGridlines && oDateAx.setMinorGridlines(this.SpPrFromJSON(oParsedDateAx.minorGridlines, oDateAx));
		oDateAx.setMinorTickMark(this.GetTickMarkNumType(oParsedDateAx.minorTickMark));
		oDateAx.setMinorTimeUnit(this.GetTimeUnitNumType(oParsedDateAx.minorTimeUnit));
		oDateAx.setMinorUnit(oParsedDateAx.minorUnit);
		oParsedDateAx.numFmt && oDateAx.setNumFmt(this.NumFmtFromJSON(oParsedDateAx.numFmt));
		oParsedDateAx.scaling && oDateAx.setScaling(this.ScalingFromJSON(oParsedDateAx.scaling, oDateAx));
		oParsedDateAx.spPr && oParsedDateAx.setSpPr(this.SpPrFromJSON(oParsedDateAx.spPr, oDateAx));
		oDateAx.setTickLblPos(this.GetTickLabelNumPos(oParsedDateAx.tickLblPos));
		oParsedDateAx.title && oDateAx.setTitle(this.TitleFromJSON(oParsedDateAx.title));
		oParsedDateAx.txPr && oDateAx.setTxPr(this.TxPrFromJSON(oParsedDateAx.txPr, oDateAx));

		return oDateAx;
	};
	ReaderFromJSON.prototype.BarChartFromJSON = function(oParsedBarChart, oAxisMap)
	{
		var oBarChart = new AscFormat.CBarChart();

		var nBarDirType = oParsedBarChart.barDir === "bar" ? AscFormat.BAR_DIR_BAR : AscFormat.BAR_DIR_COL;

		var nGroupingType = undefined;
		switch (oParsedBarChart.grouping)
		{
			case "clustered":
				nGroupingType = AscFormat.BAR_GROUPING_CLUSTERED;
				break;
			case "percentStacked":
				nGroupingType = AscFormat.BAR_GROUPING_PERCENT_STACKED;
				break;
			case "stacked":
				nGroupingType = AscFormat.BAR_GROUPING_STACKED;
				break;
			case "standard":
				nGroupingType = AscFormat.BAR_GROUPING_STANDARD;
				break;
		}

		for (var nAxis = 0; nAxis < oParsedBarChart.axId.length; nAxis++)
			oBarChart.addAxId(oAxisMap[oParsedBarChart.axId[nAxis]]);
			
		oBarChart.setBarDir(nBarDirType);
		oBarChart.setDLbls(this.DLblsFromJSON(oParsedBarChart.dLbls, oBarChart));
		oBarChart.setGapWidth(oParsedBarChart.gapWidth);
		oBarChart.setGrouping(nGroupingType);
		oBarChart.setOverlap(oParsedBarChart.overlap);
		oParsedBarChart.serLines && oBarChart.setSerLines(this.SpPrFromJSON(oParsedBarChart.serLines, oBarChart));
		this.BarSeriesFromJSON(oParsedBarChart.ser, oBarChart);
		oBarChart.setVaryColors(oParsedBarChart.varyColors);
		
		return oBarChart;
	};
	ReaderFromJSON.prototype.BarSeriesFromJSON = function(arrParsedBarSeries, oParentChart)
	{
		for (var nBarSeries = 0; nBarSeries < arrParsedBarSeries.length; nBarSeries++)
		{
			var oItem       = arrParsedBarSeries[nBarSeries];
			var oBarSeries  = new AscFormat.CBarSeries();

			var nShapeType = undefined;
			switch(arrParsedBarSeries[nBarSeries].shape)
			{
				case "cone":
					nShapeType = AscFormat.BAR_SHAPE_CONE;
					break;
				case "coneToMax":
					nShapeType = AscFormat.BAR_SHAPE_CONETOMAX;
					break;
				case "box":
					nShapeType = AscFormat.BAR_SHAPE_BOX;
					break;
				case "cylinder":
					nShapeType = AscFormat.BAR_SHAPE_CYLINDER;
					break;
				case "pyramid":
					nShapeType = AscFormat.BAR_SHAPE_PYRAMID;
					break;
				case "pyramidToMax":
					nShapeType = AscFormat.BAR_SHAPE_PYRAMIDTOMAX;
					break;
			}

			oItem.cat && oBarSeries.setCat(this.CatFromJSON(oItem.cat, oBarSeries));
			oItem.dLbls && oBarSeries.setDLbls(this.DLblsFromJSON(oItem.dLbls, oBarSeries));
			this.DataPointsFromJSON(oItem.dPt, oBarSeries);
			oItem.errBars && oBarSeries.setErrBars(this.ErrBarsFromJSON(oItem.errBars));
			oBarSeries.setIdx(oItem.idx);
			oBarSeries.setInvertIfNegative(oItem.invertIfNegative);
			oBarSeries.setOrder(oItem.order);
			oItem.pictureOptions && oBarSeries.setPictureOptions(this.PicOptionsFromJSON(oItem.pictureOptions));
			oBarSeries.setShape(nShapeType);
			oItem.spPr && oBarSeries.setSpPr(this.SpPrFromJSON(oItem.spPr, oBarSeries));
			oItem.trendline && oBarSeries.setTrendline(this.TrendlineFromJSON(oItem.trendline));
			oItem.tx && oBarSeries.setTx(this.TxFromJSON(oItem.tx, oBarSeries));
			oItem.val && oBarSeries.setVal(this.YVALFromJSON(oItem.val, oBarSeries));
			
			oParentChart.addSer(oBarSeries);
		}
	};
	ReaderFromJSON.prototype.LineChartFromJSON = function(oParsedLineChart, oAxisMap)
	{
		var oLineChart = new AscFormat.CLineChart();

		var nGroupingType = undefined;
		switch (oParsedLineChart.grouping)
		{
			case "percentStacked":
				nGroupingType = AscFormat.GROUPING_PERCENT_STACKED;
				break;
			case "stacked":
				nGroupingType = AscFormat.GROUPING_STACKED;
				break;
			case "standard":
				nGroupingType = AscFormat.GROUPING_STANDARD;
				break;
		}

		for (var nAxis = 0; nAxis < oParsedLineChart.axId.length; nAxis++)
			oLineChart.addAxId(oAxisMap[oParsedLineChart.axId[nAxis]]);

		oParsedLineChart.dLbls && oLineChart.setDLbls(this.DLblsFromJSON(oParsedLineChart.dLbls));
		oParsedLineChart.dropLines && oLineChart.setDropLines(this.SpPrFromJSON(oParsedLineChart.dropLines, oLineChart));
		oLineChart.setGrouping(nGroupingType);
		oParsedLineChart.hiLowLines && oLineChart.setHiLowLines(this.SpPrFromJSON(oParsedLineChart.hiLowLines, oLineChart));
		oLineChart.setMarker(oParsedLineChart.marker);
		this.LineSeriesFromJSON(oParsedLineChart.ser, oLineChart);
		oLineChart.setSmooth(oParsedLineChart.smooth);
		oParsedLineChart.upDownBars && oLineChart.setUpDownBars(this.UpDownBarsFromJSON(oParsedLineChart.upDownBars));
		oLineChart.setVaryColors(oParsedLineChart.varyColors);

		return oLineChart;
	};
	ReaderFromJSON.prototype.LineSeriesFromJSON = function(arrParsedLineSeries, oParentChart)
	{
		for (var nLineSeries = 0; nLineSeries < arrParsedLineSeries.length; nLineSeries++)
		{
			var oItem       = arrParsedLineSeries[nLineSeries];
			var oLineSeries = new AscFormat.CLineSeries();

			oLineSeries.setParent(oParentChart);

			oItem.cat && oLineSeries.setCat(this.CatFromJSON(oItem.cat, oLineSeries));
			oItem.dLbls && oLineSeries.setDLbls(this.DLblsFromJSON(oItem.dLbls, oLineSeries));
			this.DataPointsFromJSON(oItem.dPt, oLineSeries);
			oItem.errBars && oLineSeries.setErrBars(this.ErrBarsFromJSON(oItem.errBars));
			oLineSeries.setIdx(oItem.idx);
			oItem.marker && oLineSeries.setMarker(this.MarkerFromJSON(oItem.marker, oLineSeries));
			oLineSeries.setOrder(oItem.order);
			oLineSeries.setSmooth(oItem.smooth);
			oItem.spPr && oLineSeries.setSpPr(this.SpPrFromJSON(oItem.spPr, oLineSeries));
			oItem.trendline && oLineSeries.setTrendline(this.TrendlineFromJSON(oItem.trendline));
			oItem.tx && oLineSeries.setTx(this.TxFromJSON(oItem.tx, oLineSeries));
			oItem.val && oLineSeries.setVal(this.YVALFromJSON(oItem.val, oLineSeries));

			oParentChart.addSer(oLineSeries);
		}
	};
	ReaderFromJSON.prototype.PieChartFromJSON = function(oParsedPieChart)
	{
		var oPieChart = new AscFormat.CPieChart();

		oPieChart.set3D(oParsedPieChart.b3D);
		oPieChart.setFirstSliceAng(oParsedPieChart.firstSliceAng);
		oParsedPieChart.dLbls && oPieChart.setDLbls(this.DLblsFromJSON(oParsedPieChart.dLbls, oPieChart));
		this.PieSeriesFromJSON(oParsedPieChart.ser, oPieChart);
		oPieChart.setVaryColors(oParsedPieChart.varyColors);

		return oPieChart;
	};
	ReaderFromJSON.prototype.DoughnutChartFromJSON = function(oParsedDoughnutChart)
	{
		var oDoughnutChart = new AscFormat.CDoughnutChart();

		oDoughnutChart.setFirstSliceAng(oParsedDoughnutChart.firstSliceAng);
		oDoughnutChart.setHoleSize(oParsedDoughnutChart.holeSize);
		oParsedDoughnutChart.dLbls && oDoughnutChart.setDLbls(this.DLblsFromJSON(oParsedDoughnutChart.dLbls, oDoughnutChart));
		this.PieSeriesFromJSON(oParsedDoughnutChart.ser);
		oDoughnutChart.setVaryColors(oParsedDoughnutChart.varyColors);

		return oDoughnutChart;
	};
	ReaderFromJSON.prototype.PieSeriesFromJSON = function(arrParsedPieSeries, oParentChart)
	{
		for (var nPieSeries = 0; nPieSeries < arrParsedPieSeries.length; nPieSeries++)
		{
			var oItem      = arrParsedPieSeries[nPieSeries];
			var oPieSeries = new AscFormat.CPieSeries();

			oPieSeries.setParent(oParentChart);

			oItem.cat && oPieSeries.setCat(this.CatFromJSON(oItem.cat, oPieSeries));
			oItem.dLbls && oPieSeries.setDLbls(this.DLblsFromJSON(oItem.dLbls, oPieSeries));
			this.DataPointsFromJSON(oItem.dPt, oPieSeries);
			oPieSeries.setExplosion(oItem.explosion);
			oPieSeries.setIdx(oItem.idx);
			oPieSeries.setIdx(oItem.order);
			oItem.spPr && oPieSeries.setSpPr(this.SpPrFromJSON(oItem.spPr, oPieSeries));
			oItem.tx && oPieSeries.setTx(this.TxFromJSON(oItem.tx, oPieSeries));
			oItem.val && oPieSeries.setVal(this.YVALFromJSON(oItem.val, oPieSeries));

			oParentChart.addSer(oPieSeries);
		}
	};
	ReaderFromJSON.prototype.AreaChartFromJSON = function(oParsedAreaChart, oAxisMap)
	{
		var oAreaChart = new AscFormat.CAreaChart();

		var nGroupingType = undefined;
		switch (oParsedAreaChart.grouping)
		{
			case "percentStacked":
				nGroupingType = AscFormat.GROUPING_PERCENT_STACKED;
				break;
			case "stacked":
				nGroupingType = AscFormat.GROUPING_STACKED;
				break;
			case "standard":
				nGroupingType = AscFormat.GROUPING_STANDARD;
				break;
		
		}

		for (var nAxis = 0; nAxis < oParsedAreaChart.axId.length; nAxis++)
			oAreaChart.addAxId(oAxisMap[oParsedAreaChart.axId[nAxis]]);

		oParsedAreaChart.dLbls && oAreaChart.setDLbls(this.DLblsFromJSON(oParsedAreaChart.dLbls));
		oParsedAreaChart.dropLines && oAreaChart.setDropLines(this.SpPrFromJSON(oParsedAreaChart.dropLines, oAreaChart));
		oAreaChart.setGrouping(nGroupingType);
		this.AreaSeriesFromJSON(oParsedAreaChart.ser, oAreaChart);
		oAreaChart.setVaryColors(oParsedAreaChart.varyColors);
		
		return oAreaChart;
	};
	ReaderFromJSON.prototype.AreaSeriesFromJSON = function(arrParsedAreaSeries, oParentChart)
	{
		for (var nAreaSeries = 0; nAreaSeries < arrParsedAreaSeries.length; nAreaSeries++)
		{
			var oItem       = arrParsedAreaSeries[nAreaSeries];
			var oAreaSeries = new AscFormat.CAreaSeries();

			oAreaSeries.setParent(oParentChart);

			oItem.cat && oAreaSeries.setCat(this.CatFromJSON(oItem.cat, oAreaSeries));
			oItem.dLbls && oAreaSeries.setDLbls(this.DLblsFromJSON(oItem.dLbls));
			this.DataPointsFromJSON(oItem.dPt, oAreaSeries);
			oItem.errBars && oAreaSeries.setErrBars(this.ErrBarsFromJSON(oItem.errBars));
			oAreaSeries.setIdx(oItem.idx);
			oAreaSeries.setOrder(oItem.order);
			oItem.pictureOptions && oAreaSeries.setPictureOptions(this.PicOptionsFromJSON(oItem.pictureOptions));
			oItem.spPr && oAreaSeries.setSpPr(this.SpPrFromJSON(oItem.spPr, oAreaSeries));
			oItem.trendline && oAreaSeries.setTrendline(this.TrendlineFromJSON(oItem.trendline));
			oItem.tx && oAreaSeries.setTx(this.TxFromJSON(oItem.tx, oAreaSeries));
			oItem.val && oAreaSeries.setVal(this.YVALFromJSON(oItem.val, oAreaSeries));
			
			oParentChart.addSer(oAreaSeries);
		}
	};
	ReaderFromJSON.prototype.StockChartFromJSON = function(oParsedStockChart, oAxisMap)
	{
		var oStockChart = new AscFormat.CStockChart();

		for (var nAxis = 0; nAxis < oParsedStockChart.axId.length; nAxis++)
			oStockChart.addAxId(oAxisMap[oParsedStockChart.axId[nAxis]]);

		oParsedStockChart.dLbls && oStockChart.setDLbls(this.DLblsFromJSON(oParsedStockChart.dLbls));
		oParsedStockChart.dropLines && oStockChart.setDropLines(this.SpPrFromJSON(oParsedStockChart.dropLines, oStockChart));
		oParsedStockChart.hiLowLines && oStockChart.setHiLowLines(this.SpPrFromJSON(oParsedStockChart.hiLowLines, oStockChart));
		this.LineSeriesFromJSON(oParsedStockChart.ser, oStockChart);
		oParsedStockChart.upDownBars && oStockChart.setUpDownBars(this.UpDownBarsFromJSON(oParsedStockChart.upDownBars));

		return oStockChart;
	};
	ReaderFromJSON.prototype.StockSeriesFromJSON = function(arrParsedStockSeries, oParentChart)
	{
		var arrStockSeriesResult = [];
		// for (var nStockSeries = 0; nStockSeries < arrParsedStockSeries.length; nStockSeries++)
		// {
		// 	var oItem       = arrParsedStockSeries[nStockSeries];
		// 	var oStockSeries = new AscFormat.CStockSeries();
		//
		// 	oStockSeries.setParent(oParentChart);
		//
		// 	oStockSeries.cat            = oItem.cat ? this.CatFromJSON(oItem.cat, oStockSeries) : oStockSeries.cat;
		// 	oItem.dLbls && oStockSeries.setDLbls(this.DLblsFromJSON(oItem.dLbls));
		// 	this.DataPointsFromJSON(oItem.dPt, oStockSeries);
		// 	oStockSeries.errBars        = oItem.errBars ? this.ErrBarsFromJSON(oItem.errBars) : oStockSeries.errBars;
		// 	oStockSeries.idx            = oItem.idx;
		// 	oStockSeries.order          = oItem.order;
		// 	oStockSeries.pictureOptions = oItem.pictureOptions ? this.PicOptionsFromJSON(oItem.pictureOptions) : oStockSeries.pictureOptions;
		// 	oItem.spPr && oStockSeries.setSpPr(this.SpPrFromJSON(oItem.spPr, oStockSeries));
		// 	oStockSeries.trendline      = oItem.trendline ? this.TrendlineFromJSON(oItem.trendline) : oStockSeries.trendline;
		// 	oStockSeries.tx             = oItem.tx ? this.TxFromJSON(oItem.tx, oStockSeries) : oStockSeries.tx;
		// 	oStockSeries.val            = oItem.val ? this.YVALFromJSON(oItem.val, oStockSeries) : oStockSeries.val;
		//
		// 	arrStockSeriesResult.push(oStockSeries);
		// }

		return arrStockSeriesResult;
	};
	ReaderFromJSON.prototype.ScatterChartFromJSON = function(oParsedScatterChart, oAxisMap)
	{
		var oScatterChart = new AscFormat.CScatterChart();

		var nScatterStyle = undefined;
		switch(oParsedScatterChart.scatterStyle)
		{
			case "line":
				nScatterStyle = AscFormat.SCATTER_STYLE_LINE;
				break;
			case "lineMarker":
				nScatterStyle = AscFormat.SCATTER_STYLE_LINE_MARKER;
				break;
			case "marker":
				nScatterStyle = AscFormat.SCATTER_STYLE_MARKER;
				break;
			case "none":
				nScatterStyle = AscFormat.SCATTER_STYLE_NONE;
				break;
			case "smooth":
				nScatterStyle = AscFormat.SCATTER_STYLE_SMOOTH;
				break;
			case "smoothMarker":
				nScatterStyle = AscFormat.SCATTER_STYLE_SMOOTH_MARKER;
				break;
		}

		for (var nAxis = 0; nAxis < oParsedScatterChart.axId.length; nAxis++)
			oScatterChart.addAxId(oAxisMap[oParsedScatterChart.axId[nAxis]]);

		oParsedScatterChart.dLbls && oScatterChart.setDLbls(this.DLblsFromJSON(oParsedScatterChart.dLbls, oScatterChart));
		oScatterChart.setScatterStyle(nScatterStyle);
		this.ScatterSeriesFromJSON(oParsedScatterChart.ser, oScatterChart);

		return oScatterChart;
	};
	ReaderFromJSON.prototype.ScatterSeriesFromJSON = function(arrParsedScatterSeries, oParentChart)
	{
		for (var nScatterSeries = 0; nScatterSeries < arrParsedScatterSeries.length; nScatterSeries++)
		{
			var oItem          = arrParsedScatterSeries[nScatterSeries];
			var oScatterSeries = new AscFormat.CScatterSeries();

			oScatterSeries.setParent(oParentChart);
			
			oItem.dLbls && oScatterSeries.setDLbls(this.DLblsFromJSON(oItem.dLbls));
			this.DataPointsFromJSON(oItem.dPt, oScatterSeries);
			oItem.errBars && oScatterSeries.setErrBars(this.ErrBarsFromJSON(oItem.errBars));
			oScatterSeries.setIdx(oItem.idx);
			oItem.marker && oScatterSeries.setMarker(this.MarkerFromJSON(oItem.marker, oScatterSeries));
			oScatterSeries.setOrder(oItem.order);
			oScatterSeries.setSmooth(oItem.smooth);
			oItem.spPr && oScatterSeries.setSpPr(this.SpPrFromJSON(oItem.spPr, oScatterSeries));
			oItem.trendline && oScatterSeries.setTrendline(this.TrendlineFromJSON(oItem.trendline));
			oItem.tx && oScatterSeries.setTx(this.TxFromJSON(oItem.tx, oScatterSeries));
			oItem.yVal && oScatterSeries.setYVal(this.YVALFromJSON(oItem.yVal, oScatterSeries));
			oItem.xVal && oScatterSeries.setXVal(this.CatFromJSON(oItem.xVal, oScatterSeries));

			oParentChart.addSer(oScatterSeries);
		}
	};
	ReaderFromJSON.prototype.SurfaceChartFromJSON = function(oParsedSurfaceChart, oAxisMap)
	{
		var oSurfaceChart = new AscFormat.CSurfaceChart();

		for (var nAxis = 0; nAxis < oParsedSurfaceChart.axId.length; nAxis++)
			oSurfaceChart.addAxId(oAxisMap[oParsedSurfaceChart.axId[nAxis]]);

		this.BandFmtsFromJSON(oParsedSurfaceChart.bandFmts);
		this.SurfaceSeriesFromJSON(oParsedSurfaceChart.series, oSurfaceChart);
		oSurfaceChart.setWireframe(oParsedSurfaceChart.wireframe);

		return oSurfaceChart;
	};
	ReaderFromJSON.prototype.SurfaceSeriesFromJSON = function(arrSurfaceSeries, oParentChart)
	{
		for (var nSurfaceSeries = 0; nSurfaceSeries < arrSurfaceSeries.length; nSurfaceSeries++)
		{
			var oItem          = arrSurfaceSeries[nSurfaceSeries];
			var oSurfaceSeries = new AscFormat.CSurfaceSeries();

			oSurfaceSeries.setParent(oParentChart);

			oItem.cat && oSurfaceSeries.setCat(this.CatFromJSON(oItem.cat, oSurfaceSeries));
			oSurfaceSeries.setIdx(oItem.idx);
			oSurfaceSeries.setOrder(oItem.order);
			oItem.spPr && oSurfaceSeries.setSpPr(this.SpPrFromJSON(oItem.spPr, oSurfaceSeries));
			oItem.tx && oSurfaceSeries.setTx(this.TxFromJSON(oItem.tx, oSurfaceSeries));
			oItem.val && oSurfaceSeries.setVal(this.YVALFromJSON(oItem.val, oSurfaceSeries));
			
			oParentChart.addSer(oSurfaceSeries);
		}
	};
	ReaderFromJSON.prototype.BandFmtsFromJSON = function(arrParsedBandFmts, oParentChart)
	{
		for (var nBand = 0; nBand < arrParsedBandFmts.length; nBand++)
		{
			var oBandFmt     = new AscFormat.CBandFmt();

			oBandFmt.setIdx(arrParsedBandFmts[nBand].idx);
			arrParsedBandFmts[nBand].spPr && oBandFmt.setSpPr(this.SpPrFromJSON(arrParsedBandFmts[nBand].spPr, oBandFmt));
			
			oParentChart.addBandFmt(oBandFmt);
		}
	};
	ReaderFromJSON.prototype.BubbleChartFromJSON = function(oParsedBubbleChart, oAxisMap)
	{
		var oBubbleChart = new AscFormat.CBubbleChart();

		var nSizeRepresents = oParsedBubbleChart.sizeRepresents === "area" ? AscFormat.SIZE_REPRESENTS_AREA : AscFormat.SIZE_REPRESENTS_W;

		for (var nAxis = 0; nAxis < oParsedBubbleChart.axId.length; nAxis++)
			oBubbleChart.addAxId(oAxisMap[oParsedBubbleChart.axId[nAxis]]);

		oBubbleChart.setBubble3D(oParsedBubbleChart.bubble3D);
		oBubbleChart.setBubbleScale(oParsedBubbleChart.bubbleScale);
		oParsedBubbleChart.dLbls && oBubbleChart.setDLbls(this.DLblsFromJSON(oParsedBubbleChart.dLbls, oBubbleChart));
		this.BubbleSeriesFromJSON(oParsedBubbleChart.series, oBubbleChart);
		oBubbleChart.setShowNegBubbles(oParsedBubbleChart.showNegBubbles);
		oBubbleChart.setSizeRepresents(nSizeRepresents);
		oBubbleChart.setVaryColors(oParsedBubbleChart.varyColors);
	
		return oBubbleChart;
	};
	ReaderFromJSON.prototype.BubbleSeriesFromJSON = function(arrParsedBubbleSeries, oParentChart)
	{
		for (var nBubbleSeries = 0; nBubbleSeries < arrParsedBubbleSeries.length; nBubbleSeries++)
		{
			var oItem         = arrParsedBubbleSeries[nBubbleSeries];
			var oBubbleSeries = new AscFormat.CBubbleSeries();

			oBubbleSeries.setParent(oParentChart);

			oBubbleSeries.setBubble3D(oItem.bubble3D);
			oBubbleSeries.setBubbleSize(this.YVALFromJSON(oItem.bubbleSize, oBubbleSeries));
			oItem.dLbls && oBubbleSeries.setDLbls(this.DLblsFromJSON(oItem.dLbls, oBubbleSeries));
			this.DataPointsFromJSON(oItem.dPt, oBubbleSeries);
			oItem.errBars && oBubbleSeries.setErrBars(this.ErrBarsFromJSON(oItem.errBars));
			oBubbleSeries.setIdx(oItem.idx);
			oBubbleSeries.setInvertIfNegative(oItem.invertIfNegative);
			oBubbleSeries.setOrder(oItem.order);
			oItem.spPr && oBubbleSeries.setSpPr(this.SpPrFromJSON(oItem.spPr, oBubbleSeries));
			oItem.trendline && oBubbleSeries.setTrendline(this.TrendlineFromJSON(oItem.trendline));
			oItem.tx && oBubbleSeries.setTx(this.TxFromJSON(oItem.tx, oBubbleSeries));
			oItem.yVal && oBubbleSeries.setYVal(this.YVALFromJSON(oItem.yVal, oBubbleSeries));
			oItem.xVal && oBubbleSeries.setXVal(this.CatFromJSON(oItem.xVal, oBubbleSeries));
			
			oParentChart.addSer(oBubbleSeries);
		}
	};
	ReaderFromJSON.prototype.RadarChartFromJSON = function(oParsedRadarChart, oAxisMap)
	{
		var oRadarChart = new AscFormat.CRadarChart();

		var nRadarStyle = undefined;
		switch(oRadarChart.radarStyle)
		{
			case "standard":
				nRadarStyle = AscFormat.RADAR_STYLE_STANDARD;
				break;
			case "marker":
				nRadarStyle = AscFormat.RADAR_STYLE_MARKER;
				break;
			case "filled":
				nRadarStyle = AscFormat.RADAR_STYLE_FILLED;
				break;
		}

		for (var nAxis = 0; nAxis < oParsedRadarChart.axId.length; nAxis++)
			oRadarChart.addAxId(oAxisMap[oParsedRadarChart.axId[nAxis]]);

		oParsedRadarChart.dLbls && oRadarChart.setDLbls(this.DLblsFromJSON(oParsedRadarChart.dLbls, oRadarChart));
		oRadarChart.setRadarStyle(nRadarStyle);
		this.RadarSeriesFromJSOM(oParsedRadarChart.ser, oRadarChart);
		oRadarChart.setVaryColors(oParsedRadarChart.varyColors);

		return oRadarChart;
	};
	ReaderFromJSON.prototype.RadarSeriesFromJSOM = function(arrParsedRadarSeries, oParentChart)
	{
		for (var nRadarSeries = 0; nRadarSeries < arrParsedRadarSeries.length; nRadarSeries++)
		{
			var oItem        = arrParsedRadarSeries[nRadarSeries];
			var oRadarSeries = new AscFormat.CRadarSeries();

			oItem.cat && oRadarSeries.setCat(this.CatFromJSON(oItem.cat, oRadarSeries));
			oItem.dLbls && oRadarSeries.setDLbls(this.DLblsFromJSON(oItem.dLbls, oRadarSeries));
			this.DataPointsFromJSON(oItem.dPt, oRadarSeries);
			oRadarSeries.setIdx(oItem.idx);
			oItem.marker && oRadarSeries.setMarker(this.MarkerFromJSON(oItem.marker, oRadarSeries));
			oRadarSeries.setOrder(oItem.order);
			oItem.spPr && oRadarSeries.setSpPr(this.SpPrFromJSON(oItem.spPr, oRadarSeries));
			oItem.tx && oRadarSeries.setTx(this.TxFromJSON(oItem.tx, oRadarSeries));
			oRadarSeries.val && oRadarSeries.setVal(this.YVALFromJSON(oItem.val, oRadarSeries));

			oParentChart.addSer(oRadarSeries);
		}
	};
	ReaderFromJSON.prototype.UpDownBarsFromJSON = function(oParsedUpDownBars)
	{
		var oUpDownBars = new AscFormat.CUpDownBars();

		oParsedUpDownBars.downBars && oUpDownBars.setDownBars(this.SpPrFromJSON(oParsedUpDownBars.downBars, oUpDownBars));
		oUpDownBars.setGapWidth(oParsedUpDownBars.gapWidth);
		oParsedUpDownBars.upBars && oUpDownBars.setUpBars(this.SpPrFromJSON(oParsedUpDownBars.upBars, oUpDownBars));

		return oUpDownBars;
	};
	ReaderFromJSON.prototype.TxFromJSON = function(oParsedTx, oParent)
	{
		var oTx = new AscFormat.CTx();

		oTx.setParent(oParent);

		oTx.setVal(oParsedTx.v);
		oParsedTx.strRef && oTx.setStrRef(this.StrRefFromJSON(oParsedTx.strRef));
		
		return oTx;
	};
	ReaderFromJSON.prototype.YVALFromJSON = function(oParsedYVal, oParent)
	{
		var oYVal = new AscFormat.CYVal();

		oYVal.setParent(oParent);

		oParsedYVal.numLit && oYVal.setNumLit(this.NumLitFromJSON(oParsedYVal.numLit, oYVal));
		oParsedYVal.numRef && oYVal.setNumRef(this.NumRefFromJSON(oParsedYVal.numRef, oYVal));
		
		return oYVal;
	};
	ReaderFromJSON.prototype.TrendlineFromJSON = function(oParsedTrendLine)
	{
		var oTrendLine = new AscFormat.CTrendLine();

		var nTrendlineType = undefined;
		switch(oTrendLine.trendlineType)
		{
			case "exp":
				nTrendlineType = AscFormat.st_trendlinetypeEXP;
				break;
			case "linear":
				nTrendlineType = AscFormat.st_trendlinetypeLINEAR;
				break;
			case "log":
				nTrendlineType = AscFormat.st_trendlinetypeLOG;
				break;
			case "movingAvg":
				nTrendlineType = AscFormat.st_trendlinetypeMOVINGAVG;
				break;
			case "poly":
				nTrendlineType = AscFormat.st_trendlinetypePOLY;
				break;
			case "power":
				nTrendlineType = AscFormat.st_trendlinetypePOWER;
				break;
		}

		oTrendLine.setBackward(oParsedTrendLine.backward);
		oTrendLine.setDispEq(oParsedTrendLine.dispEq);
		oTrendLine.setDispRSqr(oParsedTrendLine.dispRSqr);
		oTrendLine.setForward(oParsedTrendLine.forward);
		oTrendLine.setIntercept(oParsedTrendLine.intercept);
		oTrendLine.setName(oParsedTrendLine.name);
		oTrendLine.setOrder(oParsedTrendLine.order);
		oTrendLine.setPeriod(oParsedTrendLine.period);
		oParsedTrendLine.spPr && oTrendLine.setSpPr(this.SpPrFromJSON(oParsedTrendLine.spPr, oTrendLine));
		oParsedTrendLine.trendlineLbl && oTrendLine.setTrendlineLbl(this.DlblFromJSON(oParsedTrendLine.trendlineLbl));
		oTrendLine.setTrendlineType(nTrendlineType);

		return oTrendLine;
	};
	ReaderFromJSON.prototype.ErrBarsFromJSON = function(oParsedErrBars)
	{
		var oErrBars = new AscFormat.CErrBars();

		var nErrBarType = undefined;
		switch(oParsedErrBars.errBarType)
		{
			case "both":
				nErrBarType = AscFormat.st_errbartypeBOTH;
				break;
			case "minus":
				nErrBarType = AscFormat.st_errbartypeMINUS;
				break;
			case "plus":
				nErrBarType = AscFormat.st_errbartypePLUS;
				break;
		}

		var nErrDir = oParsedErrBars.errDir === "x" ? AscFormat.st_errdirX : AscFormat.st_errdirY;

		var nErrValType = undefined;
		switch(oParsedErrBars.errValType)
		{
			case "cust":
				nErrValType = AscFormat.st_errvaltypeCUST;
				break;
			case "fixedVal":
				nErrValType = AscFormat.st_errvaltypeFIXEDVAL;
				break;
			case "percentage":
				nErrValType = AscFormat.st_errvaltypePERCENTAGE;
				break;
			case "stdDev":
				nErrValType = AscFormat.st_errvaltypeSTDDEV;
				break;
			case "stdErr":
				nErrValType = AscFormat.st_errvaltypeSTDERR;
				break;
		}

		oErrBars.setErrBarType(nErrBarType);
		oErrBars.setErrDir(nErrDir);
		oErrBars.setErrValType(nErrValType);
		oParsedErrBars.minus && oErrBars.setMinus(this.MinusPlusFromJSON(oParsedErrBars.minus));
		oParsedErrBars.plus && oErrBars.setPlus(this.MinusPlusFromJSON(oParsedErrBars.plus));
		oErrBars.setNoEndCap(oParsedErrBars.noEndCap);
		oParsedErrBars.spPr && oErrBars.setSpPr(this.SpPrFromJSON(oParsedErrBars.spPr, oErrBars));
		oErrBars.setVal(oParsedErrBars.val);

		return oErrBars;
	};
	ReaderFromJSON.prototype.MinusPlusFromJSON = function(oParsedMinusPlus)
	{
		var oMinusPlus = new AscFormat.CMinusPlus();

		oParsedMinusPlus.numLit && oMinusPlus.setNumLit(this.NumLitFromJSON(oParsedMinusPlus.numLit, oMinusPlus));
		oParsedMinusPlus.numRef && oMinusPlus.setNumRef(this.NumRefFromJSON(oParsedMinusPlus.numRef, oMinusPlus));

		return oMinusPlus;
	};
	ReaderFromJSON.prototype.DataPointsFromJSON = function(oParsedDataPoints, oParent)
	{
		for (var nItem = 0; nItem < oParsedDataPoints.length; nItem++)
		{
			var oDataPoint = new AscFormat.CDPt();

			oDataPoint.setBubble3D(oParsedDataPoints[nItem].bubble3D);
			oDataPoint.setExplosion(oParsedDataPoints[nItem].explosion);
			oDataPoint.setIdx(oParsedDataPoints[nItem].idx);
			oDataPoint.setInvertIfNegative(oParsedDataPoints[nItem].invertIfNegative);
			oParsedDataPoints[nItem].marker && oDataPoint.setMarker(this.MarkerFromJSON(oParsedDataPoints[nItem].marker, oDataPoint));
			oParsedDataPoints[nItem].pictureOptions && oDataPoint.setPictureOptions(this.PicOptionsFromJSON(oParsedDataPoints[nItem].pictureOptions));
			oParsedDataPoints[nItem].spPr && oDataPoint.setSpPr(this.SpPrFromJSON(oParsedDataPoints[nItem].spPr, oDataPoint));

			oParent.addDPt(oDataPoint);
		}
	};
	ReaderFromJSON.prototype.CatFromJSON = function(oParsedCat, oParent)
	{
		var oCat = new AscFormat.CCat();

		oCat.setParent(oParent);

		oParsedCat.multiLvlStrRef && oCat.setMultiLvlStrRef(this.MultiLvlStrRefFromJSON(oParsedCat.multiLvlStrRef));
		oParsedCat.numLit && oCat.setNumLit(this.NumLitFromJSON(oParsedCat.numLit, oCat));
		oParsedCat.numRef && oCat.setNumRef(this.NumRefFromJSON(oParsedCat.numRef, oCat));
		oParsedCat.strLit && oCat.setStrLit(this.StrLitFromJSON(oParsedCat.strLit));
		oParsedCat.strRef && oCat.setStrRef(this.StrRefFromJSON(oParsedCat.strRef));

		return oCat;
	};
	ReaderFromJSON.prototype.NumLitFromJSON = function(oParsedNumLit)
	{
		var oNumLit = new AscFormat.CNumLit();

		for (var nPt = 0; nPt < oParsedNumLit.pt.length; nPt++)
		{
			var oPt = new AscFormat.CNumericPoint();

			oPt.setFormatCode(oParsedNumLit.pt[nPt].formatCode);
			oPt.setIdx(oParsedNumLit.pt[nPt].idx);
			oPt.setVal(oParsedNumLit.pt[nPt].v);
			
			oNumLit.addPt(oPt);
		}

		oNumLit.setPtCount(oParsedNumLit.ptCount);
		oNumLit.setFormatCode(oParsedNumLit.formatCode);

		return oNumLit;
	};
	ReaderFromJSON.prototype.NumRefFromJSON = function(oParsedNumRef)
	{
		var oNumRef = new AscFormat.CNumRef();

		oNumRef.setF(oParsedNumRef.f);
		oParsedNumRef.numCache && oNumRef.setNumCache(this.NumLitFromJSON(oParsedNumRef.numCache));

		return oNumRef;
	};
	ReaderFromJSON.prototype.MultiLvlStrRefFromJSON = function(oParsedMultiLvl)
	{
		var oMultiLvlStrRef = new AscFormat.CMultiLvlStrRef();
		oParsedMultiLvl.multiLvlStrCache && oMultiLvlStrRef.setMultiLvlStrCache(new AscFormat.CMultiLvlStrCache());

		if (oMultiLvlStrRef.multiLvlStrCache)
		{
			for (var nLvl = 0; nLvl < oParsedMultiLvl.multiLvlStrCache.lvl.length; nLvl++)
				oMultiLvlStrRef.multiLvlStrCache.addLvl(this.StrLitFromJSON(oParsedMultiLvl.multiLvlStrCache[nLvl]));
		
			oMultiLvlStrRef.multiLvlStrCache.setPtCount(oParsedMultiLvl.multiLvlStrCache.ptCount);
		}

		oMultiLvlStrRef.setF(oParsedMultiLvl.f);

		return oMultiLvlStrRef;
	};
	ReaderFromJSON.prototype.StrLitFromJSON = function(oParsedStrLit)
	{
		var oStrLit = new AscFormat.CStrCache();

		for (var nPt = 0; nPt < oParsedStrLit.pt.length; nPt++)
		{
			var oPt = new AscFormat.CStringPoint();

			oPt.setIdx(oParsedStrLit.pt[nPt].idx);
			oPt.setVal(oParsedStrLit.pt[nPt].v);

			oStrLit.addPt(oPt);
		}
		oStrLit.setPtCount(oParsedStrLit.ptCount);
		
		return oStrLit;
	};
	ReaderFromJSON.prototype.DLblsFromJSON = function(oParsedDLbls, oParent)
	{
		var oDlbls = new AscFormat.CDLbls();

		oDlbls.setParent(oParent);

		// TickLblPos
		var nDLblPos = undefined;
		switch (oParsedDLbls.dLblPos)
		{
			case "b":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.b;
				break;
			case "bestFit":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.bestFit;
				break;
			case "ctr":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.ctr;
				break;
			case "inBase":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.inBase;
				break;
			case "inEnd":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.inEnd;
				break;
			case "l":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.l;
				break;
			case "outEnd":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.outEnd;
				break;
			case "r":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.r;
				break;
			case "t":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.t;
				break;
		}

		oDlbls.setDelete(oParsedDLbls.delete);
		oDlbls.setDLblPos(nDLblPos);
		for (var nDlbl = 0; nDlbl < oParsedDLbls.dLbl.length; nDlbl++)
			oDlbls.addDLbl(this.DlblFromJSON(oParsedDLbls.dLbl[nDlbl]));

		oParsedDLbls.leaderLines && oDlbls.setLeaderLines(this.SpPrFromJSON(oParsedDLbls.leaderLines, oDlbls));
		oParsedDLbls.numFmt && oDlbls.setNumFmt(this.NumFmtFromJSON(oParsedDLbls.numFmt));
		oDlbls.setSeparator(oParsedDLbls.separator);
		oDlbls.setShowBubbleSize(oParsedDLbls.showBubbleSize);
		oDlbls.setShowCatName(oParsedDLbls.showCatName);
		oDlbls.setShowLeaderLines(oParsedDLbls.showLeaderLines);
		oDlbls.setShowLegendKey(oParsedDLbls.showLegendKey);
		oDlbls.setShowPercent(oParsedDLbls.showPercent);
		oDlbls.setShowSerName(oParsedDLbls.showSerName);
		oDlbls.setShowVal(oParsedDLbls.showVal);
		oParsedDLbls.spPr && oDlbls.setSpPr(this.SpPrFromJSON(oParsedDLbls.spPr, oDlbls));
		oParsedDLbls.txPr && oDlbls.setTxPr(this.TxPrFromJSON(oParsedDLbls.txPr, oDlbls));

		return oDlbls;
	};
	
	ReaderFromJSON.prototype.CatAxFromJSON = function(oParsedCatAx, oParentPlotArea)
	{
		var oCatAx = new AscFormat.CCatAx();

		oCatAx.setParent(oParentPlotArea);

		var nLblAlgn = undefined;
		switch(oParsedCatAx.lblAlgn)
		{
			case "ctr":
				nLblAlgn = AscFormat.LBL_ALG_CTR;
				break;
			case "l":
				nLblAlgn = AscFormat.LBL_ALG_L;
				break;
			case "r":
				nLblAlgn = AscFormat.LBL_ALG_R;
				break;
		}

		oCatAx.setAuto(oParsedCatAx.auto);
		oCatAx.setAxId(++AscFormat.Ax_Counter.GLOBAL_AX_ID_COUNTER);
		oCatAx.setAxPos(this.GetAxPosNumType(oParsedCatAx.axPos));
		oCatAx.crossAx        = oParsedCatAx.crossAx ? oParsedCatAx.crossAx : oCatAx.crossAx;
		oCatAx.setCrosses(this.GetCrossesNumType(oParsedCatAx.crosses));
		oCatAx.setCrossesAt(oParsedCatAx.crossesAt);
		oCatAx.setDelete(oParsedCatAx.delete);
		oCatAx.extLst = oParsedCatAx.extLst; /// ???
		oCatAx.setLblAlgn(nLblAlgn);
		oCatAx.setLblOffset(oParsedCatAx.lblOffset);
		oParsedCatAx.majorGridlines && oCatAx.setMajorGridlines(this.SpPrFromJSON(oParsedCatAx.majorGridlines, oCatAx));
		oCatAx.setMajorTickMark(this.GetTickMarkNumType(oParsedCatAx.majorTickMark));
		oParsedCatAx.minorGridlines && oCatAx.setMinorGridlines(this.SpPrFromJSON(oParsedCatAx.minorGridlines, oCatAx));
		oCatAx.setMinorTickMark(this.GetTickMarkNumType(oParsedCatAx.minorTickMark));
		oCatAx.setNoMultiLvlLbl(oParsedCatAx.noMultiLvlLbl);
		oParsedCatAx.numFmt && oCatAx.setNumFmt(this.NumFmtFromJSON(oParsedCatAx.numFmt));
		oParsedCatAx.scaling && oCatAx.setScaling(this.ScalingFromJSON(oParsedCatAx.scaling, oCatAx));
		oParsedCatAx.spPr && oCatAx.setSpPr(this.SpPrFromJSON(oParsedCatAx.spPr, oCatAx));
		oCatAx.setTickLblPos(this.GetTickLabelNumPos(oParsedCatAx.tickLblPos));
		oCatAx.setTickLblSkip(oParsedCatAx.tickLblSkip);
		oCatAx.setTickMarkSkip(oParsedCatAx.tickMarkSkip);
		oParsedCatAx.title && oCatAx.setTitle(this.TitleFromJSON(oParsedCatAx.title));
		oParsedCatAx.txPr && oCatAx.setTxPr(this.TxPrFromJSON(oParsedCatAx.txPr, oCatAx));

		return oCatAx;
	};
	ReaderFromJSON.prototype.ValAxFromJSON = function(oParsedValAx, oParentPlotArea)
	{
		var oValAx = new AscFormat.CValAx();

		oValAx.setParent(oParentPlotArea);

		var sCrossBetweenType = oParsedValAx.crossBetween === "between" ? AscFormat.CROSS_BETWEEN_BETWEEN : AscFormat.CROSS_BETWEEN_MID_CAT;
		
		oValAx.setAxId(++AscFormat.Ax_Counter.GLOBAL_AX_ID_COUNTER);
		oValAx.setAxPos(this.GetAxPosNumType(oParsedValAx.axPos));
		oValAx.crossAx        = oParsedValAx.crossAx ? oParsedValAx.crossAx : oValAx.crossAx;
		oValAx.setCrossBetween(sCrossBetweenType);
		oValAx.setCrosses(this.GetCrossesNumType(oParsedValAx.crosses));
		oValAx.setCrossesAt(oParsedValAx.crossesAt);
		oValAx.setDelete(oParsedValAx.delete);
		oParsedValAx.dispUnits && oValAx.setDispUnits(this.DispUnitsFromJSON(oParsedValAx.dispUnits));
		oValAx.extLst = oParsedValAx.extLst; /// ???
		oParsedValAx.majorGridlines && oValAx.setMajorGridlines(this.SpPrFromJSON(oParsedValAx.majorGridlines, oValAx));
		oValAx.setMajorTickMark(this.GetTickMarkNumType(oParsedValAx.majorTickMark));
		oValAx.setMajorUnit(oParsedValAx.majorUnit);
		oParsedValAx.minorGridlines && oValAx.setMinorGridlines(this.SpPrFromJSON(oParsedValAx.minorGridlines));
		oValAx.setMinorTickMark(this.GetTickMarkNumType(oParsedValAx.minorTickMark));
		oValAx.setMinorUnit(oParsedValAx.minorUnit);
		oParsedValAx.numFmt && oValAx.setNumFmt(this.NumFmtFromJSON(oParsedValAx.numFmt));
		oParsedValAx.scaling && oValAx.setScaling(this.ScalingFromJSON(oParsedValAx.scaling, oValAx));
		oParsedValAx.spPr  && oValAx.setSpPr(this.SpPrFromJSON(oParsedValAx.spPr, oValAx));
		oValAx.setTickLblPos(this.GetTickLabelNumPos(oParsedValAx.tickLblPos));
		oParsedValAx.title && oValAx.setTitle(this.TitleFromJSON(oParsedValAx.title));
		oParsedValAx.txPr  && oValAx.setTxPr(this.TxPrFromJSON(oParsedValAx.txPr, oValAx));

		return oValAx;
	};
	ReaderFromJSON.prototype.TitleFromJSON = function(oParsedTitle, oParent)
	{
		var oTitle = new AscFormat.CTitle();

		oTitle.setParent(oParent);

		oTitle.setOverlay(oParsedTitle.overlay);
		oParsedTitle.layout && oTitle.setLayout(this.LayoutFromJSON(oParsedTitle.layout, oTitle));
		oParsedTitle.spPr   && oTitle.setSpPr(this.SpPrFromJSON(oParsedTitle.spPr, oTitle));
		oParsedTitle.tx     && oTitle.setTx(this.ChartTxFromJSON(oParsedTitle.tx));
		oParsedTitle.txPr   && oTitle.setTxPr(this.TxPrFromJSON(oParsedTitle.txPr, oTitle));
		
		return oTitle;
	};
	ReaderFromJSON.prototype.ScalingFromJSON = function(oParsedScaling, oParent)
	{
		var oScaling = new AscFormat.CScaling();

		oScaling.setParent(oParent);

		var nOrientType = oParsedScaling.orientation === "maxMin" ? AscFormat.ORIENTATION_MAX_MIN : AscFormat.ORIENTATION_MIN_MAX;

		oScaling.setLogBase(oParsedScaling.logBase);
		oScaling.setMax(oParsedScaling.max);
		oScaling.setMin(oParsedScaling.min);
		oScaling.setOrientation(nOrientType);

		return oScaling;
	};
	ReaderFromJSON.prototype.DispUnitsFromJSON = function(oParsedDispUnits)
	{
		var oDispUnits = new AscFormat.CDispUnits();

		var nBuiltInUnit = undefined;
		switch(oDispUnits.builtInUnit)
		{
			case "none":
				nBuiltInUnit = Asc.c_oAscValAxUnits.none;
				break;
			case "billions":
				nBuiltInUnit = Asc.c_oAscValAxUnits.BILLIONS;
				break;
			case "hundredMillions":
				nBuiltInUnit = Asc.c_oAscValAxUnits.HUNDRED_MILLIONS;
				break;
			case "hundreds":
				nBuiltInUnit = Asc.c_oAscValAxUnits.HUNDREDS;
				break;
			case "hundredThousands":
				nBuiltInUnit = Asc.c_oAscValAxUnits.HUNDRED_THOUSANDS;
				break;
			case "millions":
				nBuiltInUnit = Asc.c_oAscValAxUnits.MILLIONS;
				break;
			case "tenMillions":
				nBuiltInUnit = Asc.c_oAscValAxUnits.TEN_MILLIONS;
				break;
			case "tenThousands":
				nBuiltInUnit = Asc.c_oAscValAxUnits.TEN_THOUSANDS;
				break;
			case "trillions":
				nBuiltInUnit = Asc.c_oAscValAxUnits.TRILLIONS;
				break;
			case "custom":
				nBuiltInUnit = Asc.c_oAscValAxUnits.CUSTOM;
				break;
			case "thousands":
				nBuiltInUnit = Asc.c_oAscValAxUnits.THOUSANDS;
				break;
		}

		oDispUnits.setBuiltInUnit(nBuiltInUnit);
		oDispUnits.setCustUnit(oParsedDispUnits.custUnit);
		oParsedDispUnits.dispUnitsLbl && oDispUnits.setDispUnitsLbl(this.DlblFromJSON(oParsedDispUnits.dispUnitsLbl));

		return oDispUnits;
	};
	ReaderFromJSON.prototype.ChartsFromJSON = function(arrParsedCharts, oParentPlotArea, oAxisMap)
	{
		for (var nChart = 0; nChart < arrParsedCharts.length; nChart++)
		{
			switch (arrParsedCharts[nChart].type)
			{
				case "barChart":
					oParentPlotArea.addChart(this.BarChartFromJSON(arrParsedCharts[nChart], oAxisMap));
					break;
				case "lineChart":
					oParentPlotArea.addChart(this.LineChartFromJSON(arrParsedCharts[nChart], oAxisMap));
					break;
				case "pieChart":
					oParentPlotArea.addChart(this.PieChartFromJSON(arrParsedCharts[nChart]));
					break;
				case "doughnutChart":
					oParentPlotArea.addChart(this.DoughnutChartFromJSON(arrParsedCharts[nChart]));
					break;
				case "areaChart":
					oParentPlotArea.addChart(this.AreaChartFromJSON(arrParsedCharts[nChart], oAxisMap));
					break;
				case "stockChart":
					oParentPlotArea.addChart(this.StockChartFromJSON(arrParsedCharts[nChart], oAxisMap));
					break;
				case "scatterChart":
					oParentPlotArea.addChart(this.ScatterChartFromJSON(arrParsedCharts[nChart], oAxisMap));
					break;
				case "bubbleChart":
					oParentPlotArea.addChart(this.BubbleChartFromJSON(arrParsedCharts[nChart], oAxisMap));
					break;
				case "surfaceChart":
					oParentPlotArea.addChart(this.SurfaceChartFromJSON(arrParsedCharts[nChart], oAxisMap));
					break;
				case "radarChart":
					oParentPlotArea.addChart(this.RadarChartFromJSON(arrParsedCharts[nChart], oAxisMap));
					break;
			}
		}
	};
	ReaderFromJSON.prototype.PivotFmtFromJSON = function(oParsedPivotFmt)
	{
		var oPivotFmt = new AscFormat.CPivotFmt();

		oPivotFmt.setIdx(oParsedPivotFmt.idx);
		oParsedPivotFmt.dLbl   && oPivotFmt.setLbl(this.DlblFromJSON(oParsedPivotFmt.dLbl));
		oParsedPivotFmt.marker && oPivotFmt.setMarker(this.MarkerFromJSON(oParsedPivotFmt.marker, oPivotFmt));
		oParsedPivotFmt.spPr   && oPivotFmt.setSpPr(this.SpPrFromJSON(oParsedPivotFmt.spPr, oPivotFmt));
		oParsedPivotFmt.txPr   && oPivotFmt.setTxPr(this.TxPrFromJSON(oParsedPivotFmt.txPr, oPivotFmt));
		
		return oPivotFmt;
	};
	ReaderFromJSON.prototype.PivotFmtsFromJSON = function(oParsedPivotFmts, oParent)
	{
		for (var nItem = 0; nItem < oParsedPivotFmts.length; nItem++)
			oParent.setPivotFmts(this.PivotFmtFromJSON(oParsedPivotFmts[nItem]));
	};
	ReaderFromJSON.prototype.MarkerFromJSON = function(oParsedMarker, oParent)
	{
		var oMarker = new AscFormat.CMarker();

		oMarker.setParent(oParent);

		var nSymbolType = undefined;
		switch(oParsedMarker.symbol)
		{
			case "circle":
				nSymbolType = AscFormat.SYMBOL_CIRCLE;
				break;
			case "dash":
				nSymbolType = AscFormat.SYMBOL_DASH;
				break;
			case "diamond":
				nSymbolType = AscFormat.SYMBOL_DIAMOND;
				break;
			case "dot":
				nSymbolType = AscFormat.SYMBOL_DOT;
				break;
			case "none":
				nSymbolType = AscFormat.SYMBOL_NONE;
				break;
			case "picture":
				nSymbolType = AscFormat.SYMBOL_PICTURE;
				break;
			case "plus":
				nSymbolType = AscFormat.SYMBOL_PLUS;
				break;
			case "square":
				nSymbolType = AscFormat.SYMBOL_SQUARE;
				break;
			case "star":
				nSymbolType = AscFormat.SYMBOL_STAR;
				break;
			case "triangle":
				nSymbolType = AscFormat.SYMBOL_TRIANGLE;
				break;
			case "x":
				nSymbolType = AscFormat.SYMBOL_X;
				break;
		}

		oMarker.setSize(oParsedMarker.size);
		oParsedMarker.spPr && oMarker.setSpPr(this.SpPrFromJSON(oParsedMarker.spPr, oMarker));
		oMarker.setSymbol(nSymbolType);

		return oMarker;
	};
	ReaderFromJSON.prototype.DlblFromJSON = function(oParsedDlbl)
	{
		var oDlbl = new AscFormat.CDLbl();

		// TickLblPos
		var nDLblPos = undefined;
		switch (oParsedDlbl.dLblPos)
		{
			case "b":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.b;
				break;
			case "bestFit":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.bestFit;
				break;
			case "ctr":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.ctr;
				break;
			case "inBase":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.inBase;
				break;
			case "inEnd":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.inEnd;
				break;
			case "l":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.l;
				break;
			case "outEnd":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.outEnd;
				break;
			case "r":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.r;
				break;
			case "t":
				nDLblPos = Asc.c_oAscChartDataLabelsPos.t;
				break;
		}

		oDlbl.setDelete(oParsedDlbl.delete);
		oDlbl.setDLblPos(nDLblPos);
		oDlbl.setIdx(oParsedDlbl.idx);
		oParsedDlbl.layout && oDlbl.setLayout(this.LayoutFromJSON(oParsedDlbl.layout, oDlbl));
		oParsedDlbl.numFmt && oDlbl.setNumFmt(this.NumFmtFromJSON(oParsedDlbl.numFmt));
		oDlbl.setSeparator(oParsedDlbl.separator);
		oDlbl.setShowBubbleSize(oParsedDlbl.showBubbleSize);
		oDlbl.setShowCatName(oParsedDlbl.showCatName);
		oDlbl.setShowLegendKey(oParsedDlbl.showLegendKey);
		oDlbl.setShowPercent(oParsedDlbl.showPercent);
		oDlbl.setShowSerName(oParsedDlbl.showSerName);
		oDlbl.setShowVal(oParsedDlbl.showVal);
		oParsedDlbl.spPr && oDlbl.setSpPr(this.SpPrFromJSON(oParsedDlbl.spPr, oDlbl));
		oParsedDlbl.txPr && oDlbl.setTxPr(this.TxPrFromJSON(oParsedDlbl.txPr, oDlbl));
		oParsedDlbl.tx && oDlbl.setTx(this.ChartTxFromJSON(oParsedDlbl.tx));
		return oDlbl;
	};
	ReaderFromJSON.prototype.NumFmtFromJSON = function(oParsedNumFmt)
	{
		var oNumFmt = new AscFormat.CNumFmt();

		oNumFmt.setFormatCode(oParsedNumFmt.formatCode);
		oNumFmt.setSourceLinked(oParsedNumFmt.sourceLinked);

		return oNumFmt;
	};
	ReaderFromJSON.prototype.ChartTxFromJSON = function(oParsedChartTx)
	{
		var oChartTx = new AscFormat.CChartText();

		oParsedChartTx.strRef && oChartTx.setStrRef(this.StrRefFromJSON(oParsedChartTx.strRef));
		oParsedChartTx.rich && oChartTx.setRich(this.TxPrFromJSON(oParsedChartTx.rich));

		return oChartTx;
	};
	ReaderFromJSON.prototype.StrRefFromJSON = function(oParsedStrRef)
	{
		var oStrRef = new AscFormat.CStrRef();
		
		oStrRef.setF(oParsedStrRef.f);
		oStrRef.setStrCache(this.StrLitFromJSON(oParsedStrRef.strCache));

		return oStrRef;
	};
	ReaderFromJSON.prototype.LegendFromJSON = function(oParsedLegend, oParentChart)
	{
		var oLegend = new AscFormat.CLegend();

		oLegend.setParent(oParentChart);

		var nLegendPos = undefined;
		switch (oParsedLegend.legendPos)
		{
			case "b":
				nLegendPos = Asc.c_oAscChartLegendShowSettings.bottom;
				break;
			case "l":
				nLegendPos = Asc.c_oAscChartLegendShowSettings.left;
				break;
			case "r":
				nLegendPos = Asc.c_oAscChartLegendShowSettings.right;
				break;
			case "t":
				nLegendPos = Asc.c_oAscChartLegendShowSettings.top;
				break;
			case "tr":
				nLegendPos = Asc.c_oAscChartLegendShowSettings.topRight;
				break;
		}

		oParsedLegend.layout && oLegend.setLayout(this.LayoutFromJSON(oParsedLegend.layout));
		this.LegendEntriesFromJSON(oParsedLegend.legendEntry, oLegend);
		oLegend.setLegendPos(nLegendPos);
		oLegend.setOverlay(oParsedLegend.overlay);
		oParsedLegend.spPr && oLegend.setSpPr(this.SpPrFromJSON(oParsedLegend.spPr, oLegend));
		oParsedLegend.txPr && oLegend.setTxPr(this.TxPrFromJSON(oParsedLegend.txPr, oLegend));

		return oLegend;
	};
	ReaderFromJSON.prototype.TxPrFromJSON = function(oParsedTxPr, oParent)
	{
		var oTxPr = new AscFormat.CTextBody();

		oTxPr.setParent(oParent);

		oParsedTxPr.bodyPr && oTxPr.setBodyPr(this.BodyPrFromJSON(oParsedTxPr.bodyPr));
		oParsedTxPr.lstStyle && oTxPr.setLstStyle(this.LstStyleFromJSON(oParsedTxPr.lstStyle));
		oParsedTxPr.content && oTxPr.setContent(this.DrawingDocContentFromJSON(oParsedTxPr.content, oTxPr, undefined, undefined));
		
		return oTxPr;
	};
	ReaderFromJSON.prototype.LstStyleFromJSON = function(oParsedStyleLvls)
	{
		var oTxtLstStyle = new AscFormat.TextListStyle();

		for (var nLvl = 0; nLvl < oParsedStyleLvls.length; nLvl++)
		{
			if (oParsedStyleLvls[nLvl])
				oTxtLstStyle.levels[nLvl] = this.ParaPrFromJSON(oParsedStyleLvls[nLvl]);
		}

		return oTxtLstStyle;
	};
	ReaderFromJSON.prototype.LayoutFromJSON = function(oParsedLayout, oParent)
	{
		var oLayout = new AscFormat.CLayout();

		oLayout.setParent(oParent);

		oLayout.setH(oParsedLayout.h);
		oParsedLayout.hMode != undefined && oLayout.setHMode(oParsedLayout.hMode === "edge" ? AscFormat.LAYOUT_MODE_EDGE : AscFormat.LAYOUT_MODE_FACTOR);
		oParsedLayout.layoutTarget != undefined && oLayout.setLayoutTarget(oParsedLayout.layoutTarget === "inner" ? AscFormat.LAYOUT_TARGET_INNER : AscFormat.LAYOUT_TARGET_OUTER);
		oLayout.setW(oParsedLayout.w);
		oParsedLayout.wMode != undefined && oLayout.setWMode(oParsedLayout.wMode === "edge" ? AscFormat.LAYOUT_MODE_EDGE : AscFormat.LAYOUT_MODE_FACTOR);
		oLayout.setX(oParsedLayout.x);
		oParsedLayout.xMode != undefined && oLayout.setXMode(oParsedLayout.xMode === "edge" ? AscFormat.LAYOUT_MODE_EDGE : AscFormat.LAYOUT_MODE_FACTOR);
		oLayout.setY(oParsedLayout.y);
		oParsedLayout.yMode != undefined && oLayout.setYMode(oParsedLayout.yMode === "edge" ? AscFormat.LAYOUT_MODE_EDGE : AscFormat.LAYOUT_MODE_FACTOR);

		return oLayout;
	};
	ReaderFromJSON.prototype.LegendEntryFromJSON = function(oParsedEntry)
	{
		var oLegendEntry = new AscFormat.CLegendEntry();
		
		oLegendEntry.setDelete(oParsedEntry.delete);
		oLegendEntry.setIdx(oParsedEntry.idx);
		oParsedEntry.txPr && oLegendEntry.setTxPr(this.TxPrFromJSON(oParsedEntry.txPr, oLegendEntry));

		return oLegendEntry;
	};
	ReaderFromJSON.prototype.LegendEntriesFromJSON = function(oParsedEntries, oParent)
	{
		for (var nItem = 0; nItem < oParsedEntries.length; nItem++)
			oParent.addLegendEntry(this.LegendEntryFromJSON(oParsedEntries[nItem]));
	};
	ReaderFromJSON.prototype.WallFromJSON = function(oParsedWall, oParent)
	{
		var oWall = new AscFormat.CChartWall();

		oWall.setParent(oParent);

		oParsedWall.pictureOptions && oWall.setPictureOptions(this.PicOptionsFromJSON(oParsedWall.pictureOptions));
		oParsedWall.spPr && oWall.setSpPr(this.SpPrFromJSON(oParsedWall.spPr, oWall));
		oWall.setThickness(oParsedWall.thickness);

		return oWall;
	};
	ReaderFromJSON.prototype.PicOptionsFromJSON = function(oParsedPicOpt)
	{
		var oPicOptions = new AscFormat.CPictureOptions();

		oPicOptions.setApplyToEnd(oParsedPicOpt.applyToEnd);
		oPicOptions.setApplyToFront(oParsedPicOpt.applyToFront);
		oPicOptions.setApplyToSides(oParsedPicOpt.applyToSides);
		oPicOptions.setPictureFormat(oParsedPicOpt.pictureFormat);
		oPicOptions.setPictureStackUnit(oParsedPicOpt.pictureStackUnit);

		return oPicOptions;
	};
	ReaderFromJSON.prototype.SpStyleFromJSON = function(oParsedSpStyle)
	{
		if (!oParsedSpStyle)
			return oParsedSpStyle;

		var oStyle = new AscFormat.CShapeStyle();

		oParsedSpStyle.lnRef && oStyle.setLnRef(this.StyleRefFromJSON(oParsedSpStyle.lnRef));
		oParsedSpStyle.fillRef && oStyle.setFillRef(this.StyleRefFromJSON(oParsedSpStyle.fillRef));
		oParsedSpStyle.effectRef && oStyle.setEffectRef(this.StyleRefFromJSON(oParsedSpStyle.effectRef));
		oParsedSpStyle.fontRef && oStyle.setFontRef(this.FontRefFromJSON(oParsedSpStyle.fontRef));
		
		return oStyle;
	};
	ReaderFromJSON.prototype.BodyPrFromJSON = function(oParsedBodyPr)
	{
		var oBodyPr = new AscFormat.CBodyPr();

		var nAnchorType = null;
		switch(oParsedBodyPr.anchor)
		{
			case "b":
				nAnchorType = AscFormat.VERTICAL_ANCHOR_TYPE_BOTTOM;
				break;
			case "ctr":
				nAnchorType = AscFormat.VERTICAL_ANCHOR_TYPE_CENTER;
				break;
			case "dist":
				nAnchorType = AscFormat.VERTICAL_ANCHOR_TYPE_DISTRIBUTED;
				break;
			case "just":
				nAnchorType = AscFormat.VERTICAL_ANCHOR_TYPE_JUSTIFIED;
				break;
			case "t":
				nAnchorType = AscFormat.VERTICAL_ANCHOR_TYPE_TOP;
				break;
		}

		var nHorzOverflow = null;
		switch(oParsedBodyPr.horzOverflow)
		{
			case "clip":
				nHorzOverflow = AscFormat.nOTClip;
				break;
			case "overflow":
				nHorzOverflow = AscFormat.nOTOwerflow;
				break;
		}

		var nVertOverflow = null;
		switch(oParsedBodyPr.vertOverflow)
		{
			case "clip":
				nVertOverflow = AscFormat.nOTClip;
				break;
			case "ellipsis":
				nVertOverflow = AscFormat.nOTEllipsis;
				break;
			case "overflow":
				nVertOverflow = AscFormat.nOTOwerflow;
				break;
		}

		var nVertType = null;
		switch(oParsedBodyPr.vert)
		{
			case "eaVert":
				nVertType = AscFormat.nVertTTeaVert;
				break;
			case "horz":
				nVertType = AscFormat.nVertTThorz;
				break;
			case "mongolianVert":
				nVertType = AscFormat.nVertTTmongolianVert;
				break;
			case "vert":
				nVertType = AscFormat.nVertTTvert;
				break;
			case "vert270":
				nVertType = AscFormat.nVertTTvert270;
				break;
			case "wordArtVert":
				nVertType = AscFormat.nVertTTwordArtVert;
				break;
			case "wordArtVertRtl":
				nVertType = AscFormat.nVertTTwordArtVertRtl;
				break;
		}

		var nWrapType = null;
		switch (oParsedBodyPr.wrap)
		{
			case "none":
				nWrapType = AscFormat.nTWTNone;
				break;
			case "square":
				nWrapType = AscFormat.nTWTSquare;
				break;
		}

		oBodyPr.flatTx           = oParsedBodyPr.flatTx != undefined ? private_EMU2MM(oParsedBodyPr.flatTx) : oBodyPr.flatTx;
		oBodyPr.textFit          = oParsedBodyPr.normAutofit ? this.TextFitFromJSON(oParsedBodyPr.normAutofit) : oBodyPr.textFit;
		oBodyPr.prstTxWarp       = oParsedBodyPr.prstTxWarp ? this.GeometryFromJSON(oParsedBodyPr.prstTxWarp) : oBodyPr.prstTxWarp;
		oBodyPr.anchor           = nAnchorType;
		oBodyPr.anchorCtr        = oParsedBodyPr.anchorCtr;
		oBodyPr.bIns             = oParsedBodyPr.bIns != undefined ? private_EMU2MM(oParsedBodyPr.bIns) : oParsedBodyPr.bIns;
		oBodyPr.compatLnSpc      = oParsedBodyPr.compatLnSpc;
		oBodyPr.forceAA          = oParsedBodyPr.forceAA;
		oBodyPr.fromWordArt      = oParsedBodyPr.fromWordArt;
		oBodyPr.horzOverflow     = nHorzOverflow;
		oBodyPr.lIns             = oParsedBodyPr.lIns != undefined ? private_EMU2MM(oParsedBodyPr.lIns) : oParsedBodyPr.lIns;
		oBodyPr.numCol           = oParsedBodyPr.numCol;
		oBodyPr.rIns             = oParsedBodyPr.rIns != undefined ? private_EMU2MM(oParsedBodyPr.rIns) : oParsedBodyPr.rIns;
		oBodyPr.rot              = oParsedBodyPr.rot;
		oBodyPr.rtlCol           = oParsedBodyPr.rtlCol;
		oBodyPr.spcCol           = oParsedBodyPr.spcCol != undefined ? private_EMU2MM(oParsedBodyPr.spcCol) : oParsedBodyPr.spcCol;
		oBodyPr.spcFirstLastPara = oParsedBodyPr.spcFirstLastPara;
		oBodyPr.tIns             = oParsedBodyPr.tIns != undefined ? private_EMU2MM(oParsedBodyPr.tIns) : oParsedBodyPr.tIns;
		oBodyPr.upright          = oParsedBodyPr.upright;
		oBodyPr.vert             = nVertType;
		oBodyPr.vertOverflow     = nVertOverflow;
		oBodyPr.wrap             = nWrapType;

		return oBodyPr;
	};
	ReaderFromJSON.prototype.TextFitFromJSON = function(oParsedTextFit)
	{
		var oTextFit = new AscFormat.CTextFit();

		oTextFit.type           = oParsedTextFit.type;
		oTextFit.fontScale      = oParsedTextFit.fontScale;
		oTextFit.lnSpcReduction = oParsedTextFit.lnSpcReduction;

		return oTextFit;
	};
	ReaderFromJSON.prototype.SpPrFromJSON = function(oParsedPr, oParent)
	{
		var oSpPr = new AscFormat.CSpPr();

		oSpPr.setParent(oParent);

		oParsedPr.fill && oSpPr.setFill(this.FillFromJSON(oParsedPr.fill));
		(oParsedPr.effectDag || oParsedPr.effectLst) && oSpPr.setEffectPr(this.EffectPropsFromJSON(oParsedPr.effectDag, oParsedPr.effectLst));
		oSpPr.setBwMode(oParsedPr.bwMode);
		oParsedPr.custGeom && oSpPr.setGeometry(this.GeometryFromJSON(oParsedPr.custGeom));
		oParsedPr.ln && oSpPr.setLn(this.LnFromJSON(oParsedPr.ln));
		oParsedPr.xfrm && oSpPr.setXfrm(this.XfrmFromJSON(oParsedPr.xfrm, oSpPr));

		return oSpPr;
	};
	ReaderFromJSON.prototype.LnFromJSON = function(oParsedLn)
	{
		var oLn = new AscFormat.CLn();

		oParsedLn.fill && oLn.setFill(this.FillFromJSON(oParsedLn.fill));
		oParsedLn.lineJoin && oLn.setJoin(this.LineJoinFromJSON(oParsedLn.lineJoin));
		oParsedLn.headEnd && oLn.setHeadEnd(this.EndArrowFromJSON(oParsedLn.headEnd));
		oParsedLn.tailEnd && oLn.setTailEnd(this.EndArrowFromJSON(oParsedLn.tailEnd));
		oLn.setPrstDash(this.GetPenDashNumType(oParsedLn.prstDash));

		var nAlgnType = oParsedLn.algn != undefined ? (oParsedLn.algn === "ctr" ? 0 : 1) : oLn.algn;
		var nCapType  = undefined;
		switch (oParsedLn.cap)
		{
			case "flat":
				nCapType = 0;
				break;
			case "rnd":
				nCapType = 1;
				break;
			case "sq":
				nCapType = 2;
				break;
		}
		var nCmpdType = undefined;
		switch( oParsedLn.cmpd)
		{
			case "dbl":
				nCmpdType = 0;
				break;
			case "sng":
				nCmpdType = 1;
				break;
			case "thickThin":
				nCmpdType = 2;
				break;
			case "thinThick":
				nCmpdType = 3;
				break;
			case "tri":
				nCmpdType = 4;
				break;
		}

		oLn.setAlgn(nAlgnType);
		oLn.setCap(nCapType);
		oLn.setCmpd(nCmpdType);
		oLn.setW(oParsedLn.w);

		return oLn;
	};
	ReaderFromJSON.prototype.GetTimeUnitNumType = function(sType)
	{
		var nTimeUnit = undefined;
		switch(sType)
		{
			case "days":
				nTimeUnit = AscFormat.TIME_UNIT_DAYS;
				break;
			case "months":
				nTimeUnit = AscFormat.TIME_UNIT_MONTHS;
				break;
			case "years":
				nTimeUnit = AscFormat.TIME_UNIT_YEARS;
				break;	
		}

		return nTimeUnit;
	};
	ReaderFromJSON.prototype.GetTickMarkNumType = function(sType)
	{
		var nType = undefined;
		switch(sType)
		{
			case "cross":
				nType = Asc.c_oAscTickMark.TICK_MARK_CROSS;
				break;
			case "in":
				nType = Asc.c_oAscTickMark.TICK_MARK_IN;
				break;
			case "none":
				nType = Asc.c_oAscTickMark.TICK_MARK_NONE;
				break;
			case "out":
				nType = Asc.c_oAscTickMark.TICK_MARK_OUT;
				break;
		}

		return nType;
	};
	ReaderFromJSON.prototype.GetCrossesNumType = function(sType)
	{
		var nType = undefined;

		switch(sType)
		{
			case "autoZero":
				nType = AscFormat.CROSSES_AUTO_ZERO;
				break;
			case "max":
				nType = AscFormat.CROSSES_MAX;
				break;
			case "min":
				nType = AscFormat.CROSSES_MIN;
				break;
		}

		return nType;
	};
	ReaderFromJSON.prototype.GetTickLabelNumPos = function(sType)
	{
		var nTickLblPos = undefined;
		switch (sType)
		{
			case "high":
				nTickLblPos = Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_HIGH;
				break;
			case "low":
				nTickLblPos = Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_LOW;
				break;
			case "nextTo":
				nTickLblPos = Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO;
				break;
			case "none":
				nTickLblPos = Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE;
				break;
		}

		return nTickLblPos;
	};
	ReaderFromJSON.prototype.GetAxPosNumType = function(sType)
	{
		var nAxPos = undefined;
		switch (sType)
		{
			case "b":
				nAxPos = AscFormat.AX_POS_B;
				break;
			case "l":
				nAxPos = AscFormat.AX_POS_L;
				break;
			case "r":
				nAxPos = AscFormat.AX_POS_R;
				break;
			case "t":
				nAxPos = AscFormat.AX_POS_B;
				break;
		}

		return nAxPos;
	};
	ReaderFromJSON.prototype.GetPresetNumType = function(sType)
	{
		switch (sType)
		{
			case "cross":
				return AscCommon.global_hatch_offsets.cross;
			case "dashDnDiag":
				return AscCommon.global_hatch_offsets.dashDnDiag;
			case "dashHorz":
				return AscCommon.global_hatch_offsets.dashHorz;
			case "dashUpDiag":
				return AscCommon.global_hatch_offsets.dashUpDiag;
			case "dashVert":
				return AscCommon.global_hatch_offsets.dashVert;
			case "diagBrick":
				return AscCommon.global_hatch_offsets.diagBrick;
			case "diagCross":
				return AscCommon.global_hatch_offsets.diagCross;
			case "divot":
				return AscCommon.global_hatch_offsets.divot;
			case "dkDnDiag":
				return AscCommon.global_hatch_offsets.dkDnDiag;
			case "dkHorz":
				return AscCommon.global_hatch_offsets.dkHorz;
			case "dkUpDiag":
				return AscCommon.global_hatch_offsets.dkUpDiag;
			case "dkVert":
				return AscCommon.global_hatch_offsets.dkVert;
			case "dnDiag":
				return AscCommon.global_hatch_offsets.dnDiag;
			case "dotDmnd":
				return AscCommon.global_hatch_offsets.dotDmnd;
			case "dotGrid":
				return AscCommon.global_hatch_offsets.dotGrid;
			case "horz":
				return AscCommon.global_hatch_offsets.horz;
			case "horzBrick":
				return AscCommon.global_hatch_offsets.horzBrick;
			case "lgCheck":
				return AscCommon.global_hatch_offsets.lgCheck;
			case "lgConfetti":
				return AscCommon.global_hatch_offsets.lgConfetti;
			case "lgGrid":
				return AscCommon.global_hatch_offsets.lgGrid;
			case "ltDnDiag":
				return AscCommon.global_hatch_offsets.ltDnDiag;
			case "ltHorz":
				return AscCommon.global_hatch_offsets.ltHorz;
			case "ltUpDiag":
				return AscCommon.global_hatch_offsets.ltUpDiag;
			case "ltVert":
				return AscCommon.global_hatch_offsets.ltVert;
			case "narHorz":
				return AscCommon.global_hatch_offsets.narHorz;
			case "narVert":
				return AscCommon.global_hatch_offsets.narVert;
			case "openDmnd":
				return AscCommon.global_hatch_offsets.openDmnd;
			case "pct10":
				return AscCommon.global_hatch_offsets.pct10;
			case "pct20":
				return AscCommon.global_hatch_offsets.pct20;
			case "pct25":
				return AscCommon.global_hatch_offsets.pct25;
			case "pct30":
				return AscCommon.global_hatch_offsets.pct30;
			case "pct40":
				return AscCommon.global_hatch_offsets.pct40;
			case "pct5":
				return AscCommon.global_hatch_offsets.pct5;
			case "pct50":
				return AscCommon.global_hatch_offsets.pct50;
			case "pct60":
				return AscCommon.global_hatch_offsets.pct60;
			case "pct70":
				return AscCommon.global_hatch_offsets.pct70;
			case "pct75":
				return AscCommon.global_hatch_offsets.pct75;
			case "pct80":
				return AscCommon.global_hatch_offsets.pct80;
			case "pct90":
				return AscCommon.global_hatch_offsets.pct90;
			case "plaid":
				return AscCommon.global_hatch_offsets.plaid;
			case "shingle":
				return AscCommon.global_hatch_offsets.shingle;
			case "smCheck":
				return AscCommon.global_hatch_offsets.smCheck;
			case "smConfetti":
				return AscCommon.global_hatch_offsets.smConfetti;
			case "smGrid":
				return AscCommon.global_hatch_offsets.smGrid;
			case "solidDmnd":
				return AscCommon.global_hatch_offsets.solidDmnd;
			case "sphere":
				return AscCommon.global_hatch_offsets.sphere;
			case "trellis":
				return AscCommon.global_hatch_offsets.trellis;
			case "upDiag":
				return AscCommon.global_hatch_offsets.upDiag;
			case "vert":
				return AscCommon.global_hatch_offsets.vert;
			case "wave":
				return AscCommon.global_hatch_offsets.wave;
			case "wdDnDiag":
				return AscCommon.global_hatch_offsets.wdDnDiag;
			case "wdUpDiag":
				return AscCommon.global_hatch_offsets.wdUpDiag;
			case "weave":
				return AscCommon.global_hatch_offsets.weave;
			case "zigZag":
				return AscCommon.global_hatch_offsets.zigZag;
		}
	};
	ReaderFromJSON.prototype.GetFormulaNumType = function(sFormulaType)
	{
		var nFormulaType = undefined;
		switch(sFormulaType)
		{
			case "*/":
				nFormulaType = AscFormat.FORMULA_TYPE_MULT_DIV;
				break;
			case "+-":
				nFormulaType = AscFormat.FORMULA_TYPE_PLUS_MINUS;
				break;
			case "+/":
				nFormulaType = AscFormat.FORMULA_TYPE_PLUS_DIV;
				break;
			case "?:":
				nFormulaType = AscFormat.FORMULA_TYPE_IF_ELSE;
				break;
			case "abs":
				nFormulaType = AscFormat.FORMULA_TYPE_ABS;
				break;
			case "at2":
				nFormulaType = AscFormat.FORMULA_TYPE_AT2;
				break;
			case "cat2":
				nFormulaType = AscFormat.FORMULA_TYPE_CAT2;
				break;
			case "cos":
				nFormulaType = AscFormat.FORMULA_TYPE_COS;
				break;
			case "max":
				nFormulaType = AscFormat.FORMULA_TYPE_MAX;
				break;
			case "mod":
				nFormulaType = AscFormat.FORMULA_TYPE_MOD;
				break;
			case "pin":
				nFormulaType = AscFormat.FORMULA_TYPE_PIN;
				break;
			case "sat2":
				nFormulaType = AscFormat.FORMULA_TYPE_SAT2;
				break;
			case "sin":
				nFormulaType = AscFormat.FORMULA_TYPE_SIN;
				break;
			case "sqrt":
				nFormulaType = AscFormat.FORMULA_TYPE_SQRT;
				break;
			case "tan":
				nFormulaType = AscFormat.FORMULA_TYPE_TAN;
				break;
			case "val":
				nFormulaType = AscFormat.FORMULA_TYPE_VALUE;
				break;
			case "min":
				nFormulaType = AscFormat.FORMULA_TYPE_MIN;
				break;
		}

		return nFormulaType;
	};
	ReaderFromJSON.prototype.GetNumPhType = function(sType)
	{
		switch (sType)
		{
			case "body":
				return AscFormat.phType_body;
			case "sPhType":
				return AscFormat.phType_chart;
			case "clipArt":
				return AscFormat.phType_clipArt;
			case "ctrTitle":
				return AscFormat.phType_ctrTitle;
			case "dgm":
				return AscFormat.phType_dgm;
			case "dt":
				return AscFormat.phType_dt;
			case "ftr":
				return AscFormat.phType_ftr;
			case "hdr":
				return AscFormat.phType_hdr;
			case "media":
				return AscFormat.phType_media;
			case "obj":
				return AscFormat.phType_obj;
			case "pic":
				return AscFormat.phType_pic;
			case "sldImg":
				return AscFormat.phType_sldImg;
			case "sldNum":
				return AscFormat.phType_sldNum;
			case "subTitle":
				return AscFormat.phType_subTitle;
			case "tbl":
				return AscFormat.phType_tbl;
			case "title":
				return AscFormat.phType_title;
		}
	};
	ReaderFromJSON.prototype.GetWrapNumType = function(sType)
	{
		switch (sType)
		{
			case "none":
				return WRAPPING_TYPE_NONE;
			case "square":
				return WRAPPING_TYPE_SQUARE;
			case "through":
				return WRAPPING_TYPE_THROUGH;
			case "tight":
				return WRAPPING_TYPE_TIGHT;
			case "top_and_bottom":
				return WRAPPING_TYPE_TOP_AND_BOTTOM;
			default:
				return WRAPPING_TYPE_NONE;
		}
	};
	ReaderFromJSON.prototype.GetPenDashNumType = function(sType)
	{
		switch (sType)
		{
			case "dash":
				return Asc.c_oDashType.dash;
			case "dashDot":
				return Asc.c_oDashType.dashDot;
			case "dot":
				return Asc.c_oDashType.dot;
			case "lgDash":
				return Asc.c_oDashType.lgDash;
			case "lgDashDot":
				return Asc.c_oDashType.lgDashDot;
			case "lgDashDotDot":
				return Asc.c_oDashType.lgDashDotDot;
			case "solid":
				return Asc.c_oDashType.solid;
			case "sysDash":
				return Asc.c_oDashType.sysDash;
			case "sysDashDot":
				return Asc.c_oDashType.sysDashDot;
			case "sysDashDotDot":
				return Asc.c_oDashType.sysDashDotDot;
			case "sysDot":
				return Asc.c_oDashType.sysDot;
			default:
				return sType;
		}
	};
	ReaderFromJSON.prototype.EndArrowFromJSON = function(oParsedEndArrow)
	{
		var nType = null;
		switch(oParsedEndArrow.type)
		{
			case "none":
				nType = AscFormat.LineEndType.None;
				break;
			case "arrow":
				nType = AscFormat.LineEndType.Arrow;
				break;
			case "diamond":
				nType = AscFormat.LineEndType.Diamond;
				break;
			case "oval":
				nType = AscFormat.LineEndType.Oval;
				break;
			case "stealth":
				nType = AscFormat.LineEndType.Stealth;
				break;
			case "triangle":
				nType = AscFormat.LineEndType.Triangle;
				break;
		}

		var nLineEndSize = null;
		switch(oParsedEndArrow.len)
		{
			case "lg":
				nLineEndSize = AscFormat.LineEndSize.Large;
				break;
			case "med":
				nLineEndSize = AscFormat.LineEndSize.Mid;
				break;
			case "sm":
				nLineEndSize = AscFormat.LineEndSize.Small;
				break;
		}

		var nLineEndWidth = null;
		switch(oParsedEndArrow.w)
		{
			case "lg":
				nLineEndWidth = AscFormat.LineEndSize.Large;
				break;
			case "med":
				nLineEndWidth = AscFormat.LineEndSize.Mid;
				break;
			case "sm":
				nLineEndWidth = AscFormat.LineEndSize.Small;
				break;
		}

		var oEndArrow = new AscFormat.EndArrow();
		oEndArrow.setType(nType);
		oEndArrow.setLen(nLineEndSize);
		oEndArrow.setW(nLineEndWidth);

		return oEndArrow;
	};
	ReaderFromJSON.prototype.LineJoinFromJSON = function(oParsedLineJoin)
	{
		var oLineJoin = new AscFormat.LineJoin();

		var nType = undefined;
		switch (oParsedLineJoin.type)
		{
			case "round":
				nType = AscFormat.LineJoinType.Round;
				break;
			case "bevel":
				nType = AscFormat.LineJoinType.Bevel;
				break;
			case "miter":
				nType = AscFormat.LineJoinType.Miter;
				break;
			case "empty":
				nType = AscFormat.LineJoinType.Empty;
				break;
		}

		oLineJoin.setType(nType);
		oLineJoin.setLimit(oParsedLineJoin.lim);

		return oLineJoin;
	};
	ReaderFromJSON.prototype.XfrmFromJSON = function(oParsedXfrm, oParent)
	{
		var oXfrm = new AscFormat.CXfrm();

		oParent && oXfrm.setParent(oParent);

		oParsedXfrm.ext.cx != undefined && oXfrm.setExtX(private_EMU2MM(oParsedXfrm.ext.cx));
		oParsedXfrm.ext.cy != undefined && oXfrm.setExtY(private_EMU2MM(oParsedXfrm.ext.cy));
		oParsedXfrm.off.x != undefined && oXfrm.setOffX(private_EMU2MM(oParsedXfrm.off.x));
		oParsedXfrm.off.y != undefined && oXfrm.setOffY(private_EMU2MM(oParsedXfrm.off.y));
		oParsedXfrm.flipH != undefined && oXfrm.setFlipH(oParsedXfrm.flipH);
		oParsedXfrm.flipV != undefined && oXfrm.setFlipV(oParsedXfrm.flipV);
		oParsedXfrm.rot != undefined && oXfrm.setRot(oParsedXfrm.rot);

		oParsedXfrm.chOffX != undefined && oXfrm.setChOffX(private_EMU2MM(oParsedXfrm.chOffX));
		oParsedXfrm.chOffY != undefined && oXfrm.setChOffY(private_EMU2MM(oParsedXfrm.chOffY));
		oParsedXfrm.chExtX != undefined && oXfrm.setChExtX(private_EMU2MM(oParsedXfrm.chExtX));
		oParsedXfrm.chExtY != undefined && oXfrm.setChExtY(private_EMU2MM(oParsedXfrm.chExtY));

		return oXfrm;
	};
	ReaderFromJSON.prototype.EffectPropsFromJSON = function(oParsedEffectDag, oParsedEffectList)
	{
		var oEffectProps = new AscFormat.CEffectProperties();

		oEffectProps.EffectDag = oParsedEffectDag  ? this.EffectContainerFromJSON(oParsedEffectDag) : oEffectProps.EffectDag;
		oEffectProps.EffectLst = oParsedEffectList ? this.EffectLstFromJSON(oParsedEffectList) : oEffectProps.EffectLst;

		return oEffectProps;
	};
	ReaderFromJSON.prototype.GeometryFromJSON = function(oParsedGeom)
	{
		var oGeom = new AscFormat.Geometry();

		var oItem;
		// AhPolar
		for (var nItem = 0; nItem < oParsedGeom.ahLst.ahPolar.length; nItem++)
		{
			oItem = oParsedGeom.ahLst.ahPolar[nItem];
			oGeom.AddHandlePolar(oItem.gdRefAng, oItem.minAng, oItem.maxAng, oItem.gdRefR, oItem.minR, oItem.maxR, oItem.pos.x, oItem.pos.y)
		}
		
		// AhXY
		for (nItem = 0; nItem < oParsedGeom.ahLst.ahXY.length; nItem++)
		{
			oItem = oParsedGeom.ahLst.ahXY[nItem];
			oGeom.AddHandleXY(oItem.gdRefX, oItem.minX, oItem.maxX, oItem.gdRefY, oItem.minY, oItem.maxY, oItem.pos.x, oItem.pos.y);
		}

		// Av
		for (var key in oParsedGeom.avLst)
			oGeom.avLst[key] = oParsedGeom.avLst[key];

		// adj
		for (key in oParsedGeom.adjLst)
			oGeom.AddAdj(key, undefined, oParsedGeom.adjLst[key]);

		// Cnx
		for (nItem = 0; nItem < oParsedGeom.cnxLst.length; nItem++)
		{
			oItem = oParsedGeom.cnxLst[nItem];
			oGeom.AddCnx(oItem.ang, oItem.pos.x, oItem.pos.y);
		}

		// gdLst
		for (var nGd = 0; nGd < oParsedGeom.gdLst.length; nGd++)
		{
			oItem = oParsedGeom.gdLst[nGd];
			oGeom.AddGuide(oItem.name, this.GetFormulaNumType(oItem.fmla), oItem.x, oItem.y, oItem.z);
		}

		// pathLst
		for (var nPath = 0; nPath < oParsedGeom.pathLst.length; nPath++)
			oGeom.AddPath(this.GeomPathFromJSON(oParsedGeom.pathLst[nPath]));

		oParsedGeom.rect && oGeom.AddRect(oParsedGeom.rect.l, oParsedGeom.rect.t, oParsedGeom.rect.r, oParsedGeom.rect.b);
		if (oParsedGeom.preset)
			oGeom.setPreset(oParsedGeom.preset);

		return oGeom;
	};
	ReaderFromJSON.prototype.GeomPathFromJSON = function(oParsedPath)
	{
		var oPath = new AscFormat.Path();

		for (var nCommand = 0; nCommand < oParsedPath.commands.length; nCommand++)
		{
			var arrKeys  = Object.keys(oParsedPath.commands[nCommand]);
			var oCommand = {};
			switch (oParsedPath.commands[nCommand].id)
			{
				case "moveTo":
					oCommand.id =  0;
					break;
				case "lnTo":
					oCommand.id =  1;
					break;
				case "arcTo":
					oCommand.id =  2;
					break;
				case "cubicBezTo":
					oCommand.id =  3;
					break;
				case "quadBezTo":
					oCommand.id =  4;
					break;
				case "close":
					oCommand.id =  5;
					break;
			}
			if (oParsedPath.commands[nCommand].id !== "arcTo" && oParsedPath.commands[nCommand].id !== "close")
			{
				var arrPtKeys  = Object.keys(oParsedPath.commands[nCommand].pt);
				for (var ptKey in arrPtKeys)
					oCommand[arrPtKeys[ptKey].toUpperCase()] = oParsedPath.commands[nCommand].pt[arrPtKeys[ptKey]];
			}
			for (var key in arrKeys)
			{
				if (arrKeys[key] === "id")
					continue;
				if (arrKeys[key] !== "pt")
					oCommand[arrKeys[key]] = oParsedPath.commands[nCommand][arrKeys[key]];
			}
				

			oPath.addPathCommand(oCommand);
		}

		oParsedPath.extrusionOk != undefined && oPath.setExtrusionOk(oParsedPath.extrusionOk);
		oParsedPath.fill        != undefined && oPath.setFill(oParsedPath.fill);
		oParsedPath.h           != undefined && oPath.setPathH(oParsedPath.h);
		oParsedPath.stroke      != undefined && oPath.setStroke(oParsedPath.stroke);
		oParsedPath.w           != undefined && oPath.setPathW(oParsedPath.w);

		return oPath;
	};
	ReaderFromJSON.prototype.UniNvPrFromJSON = function(oParsedPr)
	{
		var oUniNvPr = new AscFormat.UniNvPr();
		
		oUniNvPr.setCNvPr(this.CNvPrFromJSON(oParsedPr.cNvPr));
		oUniNvPr.setNvPr(this.NvPrFromJSON(oParsedPr.nvPr));
		oUniNvPr.setUniSpPr(this.NvUniSpPrFromJSON(oParsedPr.nvUniSpPr, oUniNvPr));

		return oUniNvPr;
	};
	ReaderFromJSON.prototype.CNvPrFromJSON = function(oParsedPr)
	{
		var oCNvPr = new AscFormat.CNvPr();

		oCNvPr.setName(oParsedPr.name);
		oCNvPr.setIsHidden(oParsedPr.hidden);
		oCNvPr.setDescr(oParsedPr.descr);
		oCNvPr.setTitle(oParsedPr.title);
		oParsedPr.hlinkClick && oCNvPr.setHlinkClick(this.HLinkFromJSON(oParsedPr.hlinkClick));
		oParsedPr.hlinkHover && oCNvPr.setHlinkHover(this.HLinkFromJSON(oParsedPr.hlinkHover));
		oCNvPr.setId(oParsedPr.id);

		return oCNvPr;
	};
	ReaderFromJSON.prototype.NvPrFromJSON = function(oParsedPr)
	{
		var oNvPr = new AscFormat.NvPr();

		oNvPr.setIsPhoto(oParsedPr.isPhoto);
		oParsedPr.ph && oNvPr.setPh(this.PlaceholderFromJSON(oParsedPr.ph));
		oParsedPr.unimedia && oNvPr.setUniMedia(this.UniMediaFromJSON(oParsedPr.unimedia));
		oNvPr.setUserDrawn(oParsedPr.userDrawn);

		return oNvPr;
	};
	ReaderFromJSON.prototype.PlaceholderFromJSON = function(oParsedPh)
	{
		var oPlaceholder = new AscFormat.Ph();

		// orient
		var nOrient = typeof(oParsedPh.orient) === "string" ? (oParsedPh.orient === "vert" ? 1 : 0) : null;
		
		// size
		var nPhSz   = null;
		switch (oParsedPh.sz)
		{
			case "full":
				nPhSz = 0;
				break;
			case "half":
				nPhSz = 1;
				break;
			case "quarter":
				nPhSz = 2;
				break;
		}
		
		oParsedPh.hasCustomPrompt != undefined && oPlaceholder.setHasCustomPrompt(oParsedPh.hasCustomPrompt);
		oParsedPh.idx != undefined && oPlaceholder.setIdx(oParsedPh.idx);
		nOrient != undefined && oPlaceholder.setOrient(nOrient);
		nPhSz != undefined && oPlaceholder.setSz(nPhSz);
		oParsedPh.type != undefined && oPlaceholder.setType(this.GetNumPhType(oParsedPh.type));
		
		return oPlaceholder;
	};
	ReaderFromJSON.prototype.UniMediaFromJSON = function(oParsedUniMedia)
	{
		var oUniMedia = new AscFormat.UniMedia();

		oUniMedia.type  = oParsedUniMedia.type;
		oUniMedia.media = oParsedUniMedia.media;

		return oUniMedia;
	};
	ReaderFromJSON.prototype.NvUniSpPrFromJSON = function(oParsedPr)
	{
		var oNvUniSpPr = new AscFormat.CNvUniSpPr();

		oNvUniSpPr.locks     = oParsedPr.locks;
		oNvUniSpPr.stCnxIdx  = oParsedPr.stCnxIdx;
		oNvUniSpPr.stCnxId   = oParsedPr.stCnxId;
		oNvUniSpPr.endCnxIdx = oParsedPr.endCnxIdx;
		oNvUniSpPr.endCnxId  = oParsedPr.endCnxId;

		return oNvUniSpPr;
	};
	ReaderFromJSON.prototype.HLinkFromJSON = function(oParsedHLink)
	{
		var oHLink = new AscFormat.CT_Hyperlink();

		oHLink.id             = oParsedHLink.id;
		oHLink.invalidUrl     = oParsedHLink.invalidUrl;
		oHLink.action         = oParsedHLink.action;
		oHLink.tgtFrame       = oParsedHLink.tgtFrame;
		oHLink.tooltip        = oParsedHLink.tooltip;
		oHLink.history        = oParsedHLink.history;
		oHLink.highlightClick = oParsedHLink.highlightClick;
		oHLink.endSnd         = oParsedHLink.endSnd;

		return oHLink;
	};
	ReaderFromJSON.prototype.PositionHFromJSON = function(oParsedPosH)
	{
		// anchorH
		var nRelFromH = undefined;
		switch (oParsedPosH.relativeFrom)
		{
			case "character":
				nRelFromH = Asc.c_oAscRelativeFromH.Character;
				break;
			case "column":
				nRelFromH = Asc.c_oAscRelativeFromH.Column;
				break;
			case "insideMargin":
				nRelFromH = Asc.c_oAscRelativeFromH.InsideMargin;
				break;
			case "leftMargin":
				nRelFromH = Asc.c_oAscRelativeFromH.LeftMargin;
				break;
			case "margin":
				nRelFromH = Asc.c_oAscRelativeFromH.Margin;
				break;
			case "outsideMargin":
				nRelFromH = Asc.c_oAscRelativeFromH.OutsideMargin;
				break;
			case "page":
				nRelFromH = Asc.c_oAscRelativeFromH.Page;
				break;
			case "rightMargin":
				nRelFromH = Asc.c_oAscRelativeFromH.RightMargin;
				break;
		}

		// alignH
		var nHorAlign = undefined;
		if (oParsedPosH.align)
		{
			switch (oParsedPosH.align)
			{
				case "center":
					nHorAlign = Asc.c_oAscXAlign.Center;
					break;
				case "inside":
					nHorAlign = Asc.c_oAscXAlign.Inside;
					break;
				case "left":
					nHorAlign = Asc.c_oAscXAlign.Left;
					break;
				case "outside":
					nHorAlign = Asc.c_oAscXAlign.Outside;
					break;
				case "right":
					nHorAlign = Asc.c_oAscXAlign.Right;
					break;
			}
		}

		var nValue = oParsedPosH.align ? nHorAlign : (oParsedPosH.percent ? oParsedPosH.posOffset : private_EMU2MM(oParsedPosH.posOffset));

		return {
			Align:        !!oParsedPosH.align,
			Percent:      !!oParsedPosH.percent,
			RelativeFrom: nRelFromH,
			Value:        nValue
		}
	};
	ReaderFromJSON.prototype.PositionVFromJSON = function(oParsedPosV)
	{
		// anchorV
		var nRelFromV = undefined;
		switch (oParsedPosV.relativeFrom)
		{
			case "bottomMargin":
				nRelFromV = Asc.c_oAscRelativeFromV.BottomMargin;
				break;
			case "insideMargin":
				nRelFromV = Asc.c_oAscRelativeFromV.InsideMargin;
				break;
			case "line":
				nRelFromV = Asc.c_oAscRelativeFromV.Line;
				break;
			case "margin":
				nRelFromV = Asc.c_oAscRelativeFromV.Margin;
				break;
			case "outsideMargin":
				nRelFromV = Asc.c_oAscRelativeFromV.OutsideMargin;
				break;
			case "page":
				nRelFromV = Asc.c_oAscRelativeFromV.Page;
				break;
			case "paragraph":
				nRelFromV = Asc.c_oAscRelativeFromV.Paragraph;
				break;
			case "topMargin":
				nRelFromV = Asc.c_oAscRelativeFromV.TopMargin;
				break;
		}

		// alignV
		var nVerAlign = undefined;
		if (oParsedPosV.Align)
		{
			switch (oParsedPosV.Value)
			{
				case "bottom":
					nVerAlign = c_oAscYAlign.Bottom;
					break;
				case "center":
					nVerAlign = c_oAscYAlign.Center;
					break;
				case "inline":
					nVerAlign = c_oAscYAlign.Inline;
					break;
				case "inside":
					nVerAlign = c_oAscYAlign.Inside;
					break;
				case "outside":
					nVerAlign = c_oAscYAlign.Outside;
					break;
				case "top":
					nVerAlign = c_oAscYAlign.Top;
					break;
			}
		}

		var nValue = oParsedPosV.align ? nVerAlign : (oParsedPosV.percent ? oParsedPosV.posOffset : private_EMU2MM(oParsedPosV.posOffset));

		return {
			Align:        oParsedPosV.align ? true : false,
			Percent:      oParsedPosV.percent ? true : false,
			RelativeFrom: nRelFromV,
			Value:        nValue
		}
	};
	function To_XML_ST_HueDir(nType)
	{
		var sDir = undefined;
		switch (nType)
		{
			case AscCommon.ST_HueDir.Ccw:
				sDir = "ccw";
				break;
			case AscCommon.ST_HueDir.Cw:
				sDir = "cw";
				break;
		}

		return sDir;
	}
	function From_XML_ST_HueDir(sType)
	{
		var nDir = undefined;
		switch (sType)
		{
			case "ccw":
				nDir = AscCommon.ST_HueDir.Ccw;
				break;
			case "cw":
				nDir = AscCommon.ST_HueDir.Cw;
				break;
		}

		return nDir;
	}
	function To_XML_ST_ClrAppMethod(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommon.ST_ClrAppMethod.cycle:
				sType = "cycle";
				break;
			case AscCommon.ST_ClrAppMethod.repeat:
				sType = "repeat";
				break;
			case AscCommon.ST_ClrAppMethod.span:
				sType = "span";
				break;
		}

		return sType;
	}
	function From_XML_ST_ClrAppMethod(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "cycle":
				nType = AscCommon.ST_ClrAppMethod.cycle;
				break;
			case "repeat":
				nType = AscCommon.ST_ClrAppMethod.repeat;
				break;
			case "span":
				nType = AscCommon.ST_ClrAppMethod.span;
				break;
		}

		return nType;
	}
	function To_XML_ST_AnimLvlStr(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_AnimLvlStr.ctr:
				sVal = "ctr";
				break;
			case AscCommon.ST_AnimLvlStr.lvl:
				sVal = "lvl";
				break;
			case AscCommon.ST_AnimLvlStr.none:
				sVal = "none";
				break;
		}

		return sVal;
	}
	function From_XML_ST_AnimLvlStr(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "ctr":
				nVal = AscCommon.ST_AnimLvlStr.ctr;
				break;
			case "lvl":
				nVal = AscCommon.ST_AnimLvlStr.lvl;
				break;
			case "none":
				nVal = AscCommon.ST_AnimLvlStr.none;
				break;
		}

		return nVal;
	}
	function To_XML_ST_AnimOneStr(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_AnimOneStr.branch:
				sVal = "branch";
				break;
			case AscCommon.ST_AnimOneStr.none:
				sVal = "none";
				break;
			case AscCommon.ST_AnimOneStr.one:
				sVal = "one";
				break;
		}

		return sVal;
	}
	function From_XML_ST_AnimOneStr(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "branch":
				nVal = AscCommon.ST_AnimOneStr.branch;
				break;
			case "none":
				nVal = AscCommon.ST_AnimOneStr.none;
				break;
			case "one":
				nVal = AscCommon.ST_AnimOneStr.one;
				break;
		}

		return nVal;
	}
	function To_XML_ST_Direction(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_Direction.norm:
				sVal = "norm";
				break;
			case AscCommon.ST_Direction.rev:
				sVal = "rev";
				break;
		}

		return sVal;
	}
	function From_XML_ST_Direction(sVal)
	{
		var nVal;
		switch (sVal)
		{
			case "norm":
				nVal = AscCommon.ST_Direction.norm;
				break;
			case "rev":
				nVal = AscCommon.ST_Direction.rev;
				break;
		}

		return nVal;
	}
	function To_XML_ST_HierBranchStyle(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_HierBranchStyle.hang:
				sVal = "hang";
				break;
			case AscCommon.ST_HierBranchStyle.init:
				sVal = "init";
				break;
			case AscCommon.ST_HierBranchStyle.l:
				sVal = "l";
				break;
			case AscCommon.ST_HierBranchStyle.r:
				sVal = "r";
				break;
			case AscCommon.ST_HierBranchStyle.std:
				sVal = "std";
				break;
		}

		return sVal;
	}
	function From_XML_ST_HierBranchStyle(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "hang":
				nVal = AscCommon.ST_HierBranchStyle.hang;
				break;
			case "init":
				nVal = AscCommon.ST_HierBranchStyle.init;
				break;
			case "l":
				nVal = AscCommon.ST_HierBranchStyle.l;
				break;
			case "r":
				nVal = AscCommon.ST_HierBranchStyle.r;
				break;
			case "std":
				nVal = AscCommon.ST_HierBranchStyle.std;
				break;
		}

		return nVal;
	}
	function To_XML_ST_ResizeHandlesStr(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_ResizeHandlesStr.exact:
				sVal = "exact";
				break;
			case AscCommon.ST_ResizeHandlesStr.rel:
				sVal = "rel";
				break;
		}

		return sVal;
	}
	function From_XML_ST_ResizeHandlesStr(sVal)
	{
		var nVal;
		switch (sVal)
		{
			case "exact":
				nVal = AscCommon.ST_ResizeHandlesStr.exact;
				break;
			case "rel":
				nVal = AscCommon.ST_ResizeHandlesStr.rel;
				break;
		}

		return nVal;
	}
	function To_XML_ST_PtType(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_PtType.node:
				sVal = "node";
				break;
			case AscCommon.ST_PtType.asst:
				sVal = "asst";
				break;
			case AscCommon.ST_PtType.doc:
				sVal = "doc";
				break;
			case AscCommon.ST_PtType.pres:
				sVal = "pres";
				break;
			case AscCommon.ST_PtType.parTrans:
				sVal = "parTrans";
				break;
			case AscCommon.ST_PtType.sibTrans:
				sVal = "sibTrans";
				break;
		}

		return sVal;
	}
	function From_XML_ST_PtType(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "node":
				nVal = AscCommon.ST_PtType.node;
				break;
			case "asst":
				nVal = AscCommon.ST_PtType.asst;
				break;
			case "doc":
				nVal = AscCommon.ST_PtType.doc;
				break;
			case "pres":
				nVal = AscCommon.ST_PtType.pres;
				break;
			case "parTrans":
				nVal = AscCommon.ST_PtType.parTrans;
				break;
			case "sibTrans":
				nVal = AscCommon.ST_PtType.sibTrans;
				break;
		}

		return nVal;
	}
	function To_XML_ST_ChildOrderType(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_ChildOrderType.b:
				sVal = "b";
				break;
			case AscCommon.ST_ChildOrderType.t:
				sVal = "t";
				break;
		}

		return sVal;
	}
	function From_XML_ST_ChildOrderType(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "b":
				nVal = AscCommon.ST_ChildOrderType.b;
				break;
			case "t":
				nVal = AscCommon.ST_ChildOrderType.t;
				break;
		}

		return nVal;
	}
	function To_XML_ST_AlgorithmType(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_AlgorithmType.composite:
				sVal = "composite";
				break;
			case AscCommon.ST_AlgorithmType.conn:
				sVal = "conn";
				break;
			case AscCommon.ST_AlgorithmType.cycle:
				sVal = "cycle";
				break;
			case AscCommon.ST_AlgorithmType.hierChild:
				sVal = "hierChild";
				break;
			case AscCommon.ST_AlgorithmType.hierRoot:
				sVal = "hierRoot";
				break;
			case AscCommon.ST_AlgorithmType.pyra:
				sVal = "pyra";
				break;
			case AscCommon.ST_AlgorithmType.lin:
				sVal = "lin";
				break;
			case AscCommon.ST_AlgorithmType.sp:
				sVal = "sp";
				break;
			case AscCommon.ST_AlgorithmType.tx:
				sVal = "tx";
				break;
			case AscCommon.ST_AlgorithmType.snake:
				sVal = "snake";
				break;
		}

		return sVal;
	}
	function From_XML_ST_AlgorithmType(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "composite":
				nVal = AscCommon.ST_AlgorithmType.composite;
				break;
			case "conn":
				nVal = AscCommon.ST_AlgorithmType.conn;
				break;
			case "cycle":
				nVal = AscCommon.ST_AlgorithmType.cycle;
				break;
			case "hierChild":
				nVal = AscCommon.ST_AlgorithmType.hierChild;
				break;
			case "hierRoot":
				nVal = AscCommon.ST_AlgorithmType.hierRoot;
				break;
			case "pyra":
				nVal = AscCommon.ST_AlgorithmType.pyra;
				break;
			case "lin":
				nVal = AscCommon.ST_AlgorithmType.lin;
				break;
			case "sp":
				nVal = AscCommon.ST_AlgorithmType.sp;
				break;
			case "tx":
				nVal = AscCommon.ST_AlgorithmType.tx;
				break;
			case "snake":
				nVal = AscCommon.ST_AlgorithmType.snake;
				break;
		}

		return nVal;
	}
	function To_XML_ST_ConstraintRelationship(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_ConstraintRelationship.self:
				sVal = "self";
				break;
			case AscCommon.ST_ConstraintRelationship.ch:
				sVal = "ch";
				break;
			case AscCommon.ST_ConstraintRelationship.des:
				sVal = "des";
				break;
		}

		return sVal;
	}
	function From_XML_ST_ConstraintRelationship(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "self":
				nVal = AscCommon.ST_ConstraintRelationship.self;
				break;
			case "ch":
				nVal = AscCommon.ST_ConstraintRelationship.ch;
				break;
			case "des":
				nVal = AscCommon.ST_ConstraintRelationship.des;
				break;
		}

		return nVal;
	}
	function To_XML_ST_BoolOperator(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_BoolOperator.none:
				sVal = "none";
				break;
			case AscCommon.ST_BoolOperator.equ:
				sVal = "equ";
				break;
			case AscCommon.ST_BoolOperator.gte:
				sVal = "gte";
				break;
			case AscCommon.ST_BoolOperator.lte:
				sVal = "lte";
				break;
		}

		return sVal;
	}
	function From_XML_ST_BoolOperator(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "none":
				nVal = AscCommon.ST_BoolOperator.none;
				break;
			case "equ":
				nVal = AscCommon.ST_BoolOperator.equ;
				break;
			case "gte":
				nVal = AscCommon.ST_BoolOperator.gte;
				break;
			case "lte":
				nVal = AscCommon.ST_BoolOperator.lte;
				break;
		}

		return nVal;
	}
	function To_XML_ST_ElementType(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_ElementType.all:
				sVal = "all";
				break;
			case AscCommon.ST_ElementType.doc:
				sVal = "doc";
				break;
			case AscCommon.ST_ElementType.node:
				sVal = "node";
				break;
			case AscCommon.ST_ElementType.norm:
				sVal = "norm";
				break;
			case AscCommon.ST_ElementType.nonNorm:
				sVal = "nonNorm";
				break;
			case AscCommon.ST_ElementType.asst:
				sVal = "asst";
				break;
			case AscCommon.ST_ElementType.nonAsst:
				sVal = "nonAsst";
				break;
			case AscCommon.ST_ElementType.parTrans:
				sVal = "parTrans";
				break;
			case AscCommon.ST_ElementType.pres:
				sVal = "pres";
				break;
			case AscCommon.ST_ElementType.sibTrans:
				sVal = "sibTrans";
				break;
		}

		return sVal;
	}
	function From_XML_ST_ElementType(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "all":
				nVal = AscCommon.ST_ElementType.all;
				break;
			case "doc":
				nVal = AscCommon.ST_ElementType.doc;
				break;
			case "node":
				nVal = AscCommon.ST_ElementType.node;
				break;
			case "norm":
				nVal = AscCommon.ST_ElementType.norm;
				break;
			case "nonNorm":
				nVal = AscCommon.ST_ElementType.nonNorm;
				break;
			case "asst":
				nVal = AscCommon.ST_ElementType.asst;
				break;
			case "nonAsst":
				nVal = AscCommon.ST_ElementType.nonAsst;
				break;
			case "parTrans":
				nVal = AscCommon.ST_ElementType.parTrans;
				break;
			case "pres":
				nVal = AscCommon.ST_ElementType.pres;
				break;
			case "sibTrans":
				nVal = AscCommon.ST_ElementType.sibTrans;
				break;
		}

		return nVal;
	}
	function To_XML_ST_ConstraintType(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_ConstraintType.alignOff:
				sVal = "alignOff";
				break;
			case AscCommon.ST_ConstraintType.b:
				sVal = "b";
				break;
			case AscCommon.ST_ConstraintType.begMarg:
				sVal = "begMarg";
				break;
			case AscCommon.ST_ConstraintType.begPad:
				sVal = "begPad";
				break;
			case AscCommon.ST_ConstraintType.bendDist:
				sVal = "bendDist";
				break;
			case AscCommon.ST_ConstraintType.bMarg:
				sVal = "bMarg";
				break;
			case AscCommon.ST_ConstraintType.bOff:
				sVal = "bOff";
				break;
			case AscCommon.ST_ConstraintType.connDist:
				sVal = "connDist";
				break;
			case AscCommon.ST_ConstraintType.ctrX:
				sVal = "ctrX";
				break;
			case AscCommon.ST_ConstraintType.ctrXOff:
				sVal = "ctrXOff";
				break;
			case AscCommon.ST_ConstraintType.ctrY:
				sVal = "ctrY";
				break;
			case AscCommon.ST_ConstraintType.ctrYOff:
				sVal = "ctrYOff";
				break;
			case AscCommon.ST_ConstraintType.diam:
				sVal = "diam";
				break;
			case AscCommon.ST_ConstraintType.endMarg:
				sVal = "endMarg";
				break;
			case AscCommon.ST_ConstraintType.endPad:
				sVal = "endPad";
				break;
			case AscCommon.ST_ConstraintType.h:
				sVal = "h";
				break;
			case AscCommon.ST_ConstraintType.hArH:
				sVal = "hArH";
				break;
			case AscCommon.ST_ConstraintType.hOff:
				sVal = "hOff";
				break;
			case AscCommon.ST_ConstraintType.l:
				sVal = "l";
				break;
			case AscCommon.ST_ConstraintType.lMarg:
				sVal = "lMarg";
				break;
			case AscCommon.ST_ConstraintType.lOff:
				sVal = "lOff";
				break;
			case AscCommon.ST_ConstraintType.none:
				sVal = "none";
				break;
			case AscCommon.ST_ConstraintType.primFontSz:
				sVal = "primFontSz";
				break;
			case AscCommon.ST_ConstraintType.pyraAcctRatio:
				sVal = "pyraAcctRatio";
				break;
			case AscCommon.ST_ConstraintType.r:
				sVal = "r";
				break;
			case AscCommon.ST_ConstraintType.rMarg:
				sVal = "rMarg";
				break;
			case AscCommon.ST_ConstraintType.rOff:
				sVal = "rOff";
				break;
			case AscCommon.ST_ConstraintType.secFontSz:
				sVal = "secFontSz";
				break;
			case AscCommon.ST_ConstraintType.secSibSp:
				sVal = "secSibSp";
				break;
			case AscCommon.ST_ConstraintType.sibSp:
				sVal = "sibSp";
				break;
			case AscCommon.ST_ConstraintType.sp:
				sVal = "sp";
				break;
			case AscCommon.ST_ConstraintType.stemThick:
				sVal = "stemThick";
				break;
			case AscCommon.ST_ConstraintType.t:
				sVal = "t";
				break;
			case AscCommon.ST_ConstraintType.tMarg:
				sVal = "tMarg";
				break;
			case AscCommon.ST_ConstraintType.tOff:
				sVal = "tOff";
				break;
			case AscCommon.ST_ConstraintType.userA:
				sVal = "userA";
				break;
			case AscCommon.ST_ConstraintType.userB:
				sVal = "userB";
				break;
			case AscCommon.ST_ConstraintType.userC:
				sVal = "userC";
				break;
			case AscCommon.ST_ConstraintType.userD:
				sVal = "userD";
				break;
			case AscCommon.ST_ConstraintType.userE:
				sVal = "userE";
				break;
			case AscCommon.ST_ConstraintType.userF:
				sVal = "userF";
				break;
			case AscCommon.ST_ConstraintType.userG:
				sVal = "userG";
				break;
			case AscCommon.ST_ConstraintType.userH:
				sVal = "userH";
				break;
			case AscCommon.ST_ConstraintType.userI:
				sVal = "userI";
				break;
			case AscCommon.ST_ConstraintType.userJ:
				sVal = "userJ";
				break;
			case AscCommon.ST_ConstraintType.userK:
				sVal = "userK";
				break;
			case AscCommon.ST_ConstraintType.userL:
				sVal = "userL";
				break;
			case AscCommon.ST_ConstraintType.userM:
				sVal = "userM";
				break;
			case AscCommon.ST_ConstraintType.userN:
				sVal = "userN";
				break;
			case AscCommon.ST_ConstraintType.userO:
				sVal = "userO";
				break;
			case AscCommon.ST_ConstraintType.userP:
				sVal = "userP";
				break;
			case AscCommon.ST_ConstraintType.userQ:
				sVal = "userQ";
				break;
			case AscCommon.ST_ConstraintType.userR:
				sVal = "userR";
				break;
			case AscCommon.ST_ConstraintType.userS:
				sVal = "userS";
				break;
			case AscCommon.ST_ConstraintType.userT:
				sVal = "userT";
				break;
			case AscCommon.ST_ConstraintType.userU:
				sVal = "userU";
				break;
			case AscCommon.ST_ConstraintType.userV:
				sVal = "userV";
				break;
			case AscCommon.ST_ConstraintType.userW:
				sVal = "userW";
				break;
			case AscCommon.ST_ConstraintType.userX:
				sVal = "userX";
				break;
			case AscCommon.ST_ConstraintType.userY:
				sVal = "userY";
				break;
			case AscCommon.ST_ConstraintType.userZ:
				sVal = "userZ";
				break;
			case AscCommon.ST_ConstraintType.w:
				sVal = "w";
				break;
			case AscCommon.ST_ConstraintType.wArH:
				sVal = "wArH";
				break;
			case AscCommon.ST_ConstraintType.wOff:
				sVal = "wOff";
				break;
		}

		return sVal;
	}
	function From_XML_ST_ConstraintType(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "alignOff":
				nVal = AscCommon.ST_ConstraintType.alignOff;
				break;
			case "b":
				nVal = AscCommon.ST_ConstraintType.b;
				break;
			case "begMarg":
				nVal = AscCommon.ST_ConstraintType.begMarg;
				break;
			case "begPad":
				nVal = AscCommon.ST_ConstraintType.begPad;
				break;
			case "bendDist":
				nVal = AscCommon.ST_ConstraintType.bendDist;
				break;
			case "bMarg":
				nVal = AscCommon.ST_ConstraintType.bMarg;
				break;
			case "bOff":
				nVal = AscCommon.ST_ConstraintType.bOff;
				break;
			case "connDist":
				nVal = AscCommon.ST_ConstraintType.connDist;
				break;
			case "ctrX":
				nVal = AscCommon.ST_ConstraintType.ctrX;
				break;
			case "ctrXOff":
				nVal = AscCommon.ST_ConstraintType.ctrXOff;
				break;
			case "ctrY":
				nVal = AscCommon.ST_ConstraintType.ctrY;
				break;
			case "ctrYOff":
				nVal = AscCommon.ST_ConstraintType.ctrYOff;
				break;
			case "diam":
				nVal = AscCommon.ST_ConstraintType.diam;
				break;
			case "endMarg":
				nVal = AscCommon.ST_ConstraintType.endMarg;
				break;
			case "endPad":
				nVal = AscCommon.ST_ConstraintType.endPad;
				break;
			case "h":
				nVal = AscCommon.ST_ConstraintType.h;
				break;
			case "hArH":
				nVal = AscCommon.ST_ConstraintType.hArH;
				break;
			case "hOff":
				nVal = AscCommon.ST_ConstraintType.hOff;
				break;
			case "l":
				nVal = AscCommon.ST_ConstraintType.l;
				break;
			case "lMarg":
				nVal = AscCommon.ST_ConstraintType.lMarg;
				break;
			case "lOff":
				nVal = AscCommon.ST_ConstraintType.lOff;
				break;
			case "none":
				nVal = AscCommon.ST_ConstraintType.none;
				break;
			case "primFontSz":
				nVal = AscCommon.ST_ConstraintType.primFontSz;
				break;
			case "pyraAcctRatio":
				nVal = AscCommon.ST_ConstraintType.pyraAcctRatio;
				break;
			case "r":
				nVal = AscCommon.ST_ConstraintType.r;
				break;
			case "rMarg":
				nVal = AscCommon.ST_ConstraintType.rMarg;
				break;
			case "rOff":
				nVal = AscCommon.ST_ConstraintType.rOff;
				break;
			case "secFontSz":
				nVal = AscCommon.ST_ConstraintType.secFontSz;
				break;
			case "secSibSp":
				nVal = AscCommon.ST_ConstraintType.secSibSp;
				break;
			case "sibSp":
				nVal = AscCommon.ST_ConstraintType.sibSp;
				break;
			case "sp":
				nVal = AscCommon.ST_ConstraintType.sp;
				break;
			case "stemThick":
				nVal = AscCommon.ST_ConstraintType.stemThick;
				break;
			case "t":
				nVal = AscCommon.ST_ConstraintType.t;
				break;
			case "tMarg":
				nVal = AscCommon.ST_ConstraintType.tMarg;
				break;
			case "tOff":
				nVal = AscCommon.ST_ConstraintType.tOff;
				break;
			case "userA":
				nVal = AscCommon.ST_ConstraintType.userA;
				break;
			case "userB":
				nVal = AscCommon.ST_ConstraintType.userB;
				break;
			case "userC":
				nVal = AscCommon.ST_ConstraintType.userC;
				break;
			case "userD":
				nVal = AscCommon.ST_ConstraintType.userD;
				break;
			case "userE":
				nVal = AscCommon.ST_ConstraintType.userE;
				break;
			case "userF":
				nVal = AscCommon.ST_ConstraintType.userF;
				break;
			case "userG":
				nVal = AscCommon.ST_ConstraintType.userG;
				break;
			case "userH":
				nVal = AscCommon.ST_ConstraintType.userH;
				break;
			case "userI":
				nVal = AscCommon.ST_ConstraintType.userI;
				break;
			case "userJ":
				nVal = AscCommon.ST_ConstraintType.userJ;
				break;
			case "userK":
				nVal = AscCommon.ST_ConstraintType.userK;
				break;
			case "userL":
				nVal = AscCommon.ST_ConstraintType.userL;
				break;
			case "userM":
				nVal = AscCommon.ST_ConstraintType.userM;
				break;
			case "userN":
				nVal = AscCommon.ST_ConstraintType.userN;
				break;
			case "userO":
				nVal = AscCommon.ST_ConstraintType.userO;
				break;
			case "userP":
				nVal = AscCommon.ST_ConstraintType.userP;
				break;
			case "userQ":
				nVal = AscCommon.ST_ConstraintType.userQ;
				break;
			case "userR":
				nVal = AscCommon.ST_ConstraintType.userR;
				break;
			case "userS":
				nVal = AscCommon.ST_ConstraintType.userS;
				break;
			case "userT":
				nVal = AscCommon.ST_ConstraintType.userT;
				break;
			case "userU":
				nVal = AscCommon.ST_ConstraintType.userU;
				break;
			case "userV":
				nVal = AscCommon.ST_ConstraintType.userV;
				break;
			case "userW":
				nVal = AscCommon.ST_ConstraintType.userW;
				break;
			case "userX":
				nVal = AscCommon.ST_ConstraintType.userX;
				break;
			case "userY":
				nVal = AscCommon.ST_ConstraintType.userY;
				break;
			case "userZ":
				nVal = AscCommon.ST_ConstraintType.userZ;
				break;
			case "w":
				nVal = AscCommon.ST_ConstraintType.w;
				break;
			case "wArH":
				nVal = AscCommon.ST_ConstraintType.wArH;
				break;
			case "wOff":
				nVal = AscCommon.ST_ConstraintType.wOff;
				break;
		}

		return nVal;
	}
	function To_XML_ST_VariableType(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_VariableType.animLvl:
				sVal = "animLvl";
				break;
			case AscCommon.ST_VariableType.animOne:
				sVal = "animOne";
				break;
			case AscCommon.ST_VariableType.bulEnabled:
				sVal = "bulEnabled";
				break;
			case AscCommon.ST_VariableType.chMax:
				sVal = "chMax";
				break;
			case AscCommon.ST_VariableType.chPref:
				sVal = "chPref";
				break;
			case AscCommon.ST_VariableType.dir:
				sVal = "dir";
				break;
			case AscCommon.ST_VariableType.hierBranch:
				sVal = "hierBranch";
				break;
			case AscCommon.ST_VariableType.none:
				sVal = "none";
				break;
			case AscCommon.ST_VariableType.orgChart:
				sVal = "orgChart";
				break;
			case AscCommon.ST_VariableType.resizeHandles:
				sVal = "resizeHandles";
				break;
		}

		return sVal;
	}
	function From_XML_ST_VariableType(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "animLvl":
				nVal = AscCommon.ST_VariableType.animLvl;
				break;
			case "animOne":
				nVal = AscCommon.ST_VariableType.animOne;
				break;
			case "bulEnabled":
				nVal = AscCommon.ST_VariableType.bulEnabled;
				break;
			case "chMax":
				nVal = AscCommon.ST_VariableType.chMax;
				break;
			case "chPref":
				nVal = AscCommon.ST_VariableType.chPref;
				break;
			case "dir":
				nVal = AscCommon.ST_VariableType.dir;
				break;
			case "hierBranch":
				nVal = AscCommon.ST_VariableType.hierBranch;
				break;
			case "none":
				nVal = AscCommon.ST_VariableType.none;
				break;
			case "orgChart":
				nVal = AscCommon.ST_VariableType.orgChart;
				break;
			case "resizeHandles":
				nVal = AscCommon.ST_VariableType.resizeHandles;
				break;
		}

		return nVal;
	}
	function To_XML_ST_AxisType(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_AxisType.ancst:
				sVal = "ancst";
				break;
			case AscCommon.ST_AxisType.ancstOrSelf:
				sVal = "ancstOrSelf";
				break;
			case AscCommon.ST_AxisType.ch:
				sVal = "ch";
				break;
			case AscCommon.ST_AxisType.des:
				sVal = "des";
				break;
			case AscCommon.ST_AxisType.desOrSelf:
				sVal = "desOrSelf";
				break;
			case AscCommon.ST_AxisType.follow:
				sVal = "follow";
				break;
			case AscCommon.ST_AxisType.followSib:
				sVal = "followSib";
				break;
			case AscCommon.ST_AxisType.none:
				sVal = "none";
				break;
			case AscCommon.ST_AxisType.par:
				sVal = "par";
				break;
			case AscCommon.ST_AxisType.preced:
				sVal = "preced";
				break;
			case AscCommon.ST_AxisType.precedSib:
				sVal = "precedSib";
				break;
			case AscCommon.ST_AxisType.root:
				sVal = "root";
				break;
			case AscCommon.ST_AxisType.self:
				sVal = "self";
				break;
		}

		return sVal;
	}
	function From_XML_ST_AxisType(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "ancst":
				nVal = AscCommon.ST_AxisType.ancst;
				break;
			case "ancstOrSelf":
				nVal = AscCommon.ST_AxisType.ancstOrSelf;
				break;
			case "ch":
				nVal = AscCommon.ST_AxisType.ch;
				break;
			case "des":
				nVal = AscCommon.ST_AxisType.des;
				break;
			case "desOrSelf":
				nVal = AscCommon.ST_AxisType.desOrSelf;
				break;
			case "follow":
				nVal = AscCommon.ST_AxisType.follow;
				break;
			case "followSib":
				nVal = AscCommon.ST_AxisType.followSib;
				break;
			case "none":
				nVal = AscCommon.ST_AxisType.none;
				break;
			case "par":
				nVal = AscCommon.ST_AxisType.par;
				break;
			case "preced":
				nVal = AscCommon.ST_AxisType.preced;
				break;
			case "precedSib":
				nVal = AscCommon.ST_AxisType.precedSib;
				break;
			case "root":
				nVal = AscCommon.ST_AxisType.root;
				break;
			case "self":
				nVal = AscCommon.ST_AxisType.self;
				break;
		}

		return nVal;
	}
	function To_XML_ST_FunctionType(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_FunctionType.cnt:
				sVal = "cnt";
				break;
			case AscCommon.ST_FunctionType.depth:
				sVal = "depth";
				break;
			case AscCommon.ST_FunctionType.maxDepth:
				sVal = "maxDepth";
				break;
			case AscCommon.ST_FunctionType.pos:
				sVal = "pos";
				break;
			case AscCommon.ST_FunctionType.posEven:
				sVal = "posEven";
				break;
			case AscCommon.ST_FunctionType.posOdd:
				sVal = "posOdd";
				break;
			case AscCommon.ST_FunctionType.revPos:
				sVal = "revPos";
				break;
			case AscCommon.ST_FunctionType.var:
				sVal = "var";
				break;
		}

		return sVal;
	}
	function From_XML_ST_FunctionType(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "cnt":
				nVal = AscCommon.ST_FunctionType.cnt;
				break;
			case "depth":
				nVal = AscCommon.ST_FunctionType.depth;
				break;
			case "maxDepth":
				nVal = AscCommon.ST_FunctionType.maxDepth;
				break;
			case "pos":
				nVal = AscCommon.ST_FunctionType.pos;
				break;
			case "posEven":
				nVal = AscCommon.ST_FunctionType.posEven;
				break;
			case "posOdd":
				nVal = AscCommon.ST_FunctionType.posOdd;
				break;
			case "revPos":
				nVal = AscCommon.ST_FunctionType.revPos;
				break;
			case "var":
				nVal = AscCommon.ST_FunctionType.var;
				break;
		}

		return nVal;
	}
	function To_XML_ST_FunctionOperator(nVal)
	{
		var sVal = undefined;
		switch (nVal)
		{
			case AscCommon.ST_FunctionOperator.equ:
				sVal = "equ";
				break;
			case AscCommon.ST_FunctionOperator.gt:
				sVal = "gt";
				break;
			case AscCommon.ST_FunctionOperator.gte:
				sVal = "gte";
				break;
			case AscCommon.ST_FunctionOperator.lt:
				sVal = "lt";
				break;
			case AscCommon.ST_FunctionOperator.lte:
				sVal = "lte";
				break;
			case AscCommon.ST_FunctionOperator.neq:
				sVal = "neq";
				break;
		}

		return sVal;
	}
	function From_XML_ST_FunctionOperator(sVal)
	{
		var nVal = undefined;
		switch (sVal)
		{
			case "equ":
				nVal = AscCommon.ST_FunctionOperator.equ;
				break;
			case "gt":
				nVal = AscCommon.ST_FunctionOperator.gt;
				break;
			case "gte":
				nVal = AscCommon.ST_FunctionOperator.gte;
				break;
			case "lt":
				nVal = AscCommon.ST_FunctionOperator.lt;
				break;
			case "lte":
				nVal = AscCommon.ST_FunctionOperator.lte;
				break;
			case "neq":
				nVal = AscCommon.ST_FunctionOperator.neq;
				break;
		}

		return nVal;
	}
	function To_XML_ST_LayoutShapeType(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case AscCommon.ST_LayoutShapeType.conn:
				sVal = "conn";
				break;
			case AscCommon.ST_LayoutShapeType.none:
				sVal = "none";
				break;
			case AscCommon.ST_LayoutShapeType.accentBorderCallout1:
				sVal = "accentBorderCallout1";
				break;
			case AscCommon.ST_LayoutShapeType.accentBorderCallout2:
				sVal = "accentBorderCallout2";
				break;
			case AscCommon.ST_LayoutShapeType.accentBorderCallout3:
				sVal = "accentBorderCallout3";
				break;
			case AscCommon.ST_LayoutShapeType.accentCallout1:
				sVal = "accentCallout1";
				break;
			case AscCommon.ST_LayoutShapeType.accentCallout2:
				sVal = "accentCallout2";
				break;
			case AscCommon.ST_LayoutShapeType.accentCallout3:
				sVal = "accentCallout3";
				break;
			case AscCommon.ST_LayoutShapeType.actionButtonBackPrevious:
				sVal = "actionButtonBackPrevious";
				break;
			case AscCommon.ST_LayoutShapeType.actionButtonBeginning:
				sVal = "actionButtonBeginning";
				break;
			case AscCommon.ST_LayoutShapeType.actionButtonBlank:
				sVal = "actionButtonBlank";
				break;
			case AscCommon.ST_LayoutShapeType.actionButtonDocument:
				sVal = "actionButtonDocument";
				break;
			case AscCommon.ST_LayoutShapeType.actionButtonEnd:
				sVal = "actionButtonEnd";
				break;
			case AscCommon.ST_LayoutShapeType.actionButtonForwardNext:
				sVal = "actionButtonForwardNext";
				break;
			case AscCommon.ST_LayoutShapeType.actionButtonHelp:
				sVal = "actionButtonHelp";
				break;
			case AscCommon.ST_LayoutShapeType.actionButtonHome:
				sVal = "actionButtonHome";
				break;
			case AscCommon.ST_LayoutShapeType.actionButtonInformation:
				sVal = "actionButtonInformation";
				break;
			case AscCommon.ST_LayoutShapeType.actionButtonMovie:
				sVal = "actionButtonMovie";
				break;
			case AscCommon.ST_LayoutShapeType.actionButtonReturn:
				sVal = "actionButtonReturn";
				break;
			case AscCommon.ST_LayoutShapeType.actionButtonSound:
				sVal = "actionButtonSound";
				break;
			case AscCommon.ST_LayoutShapeType.arc:
				sVal = "arc";
				break;
			case AscCommon.ST_LayoutShapeType.bentArrow:
				sVal = "bentArrow";
				break;
			case AscCommon.ST_LayoutShapeType.bentConnector2:
				sVal = "bentConnector2";
				break;
			case AscCommon.ST_LayoutShapeType.bentConnector3:
				sVal = "bentConnector3";
				break;
			case AscCommon.ST_LayoutShapeType.bentConnector4:
				sVal = "bentConnector4";
				break;
			case AscCommon.ST_LayoutShapeType.bentConnector5:
				sVal = "bentConnector5";
				break;
			case AscCommon.ST_LayoutShapeType.bentUpArrow:
				sVal = "bentUpArrow";
				break;
			case AscCommon.ST_LayoutShapeType.bevel:
				sVal = "bevel";
				break;
			case AscCommon.ST_LayoutShapeType.blockArc:
				sVal = "blockArc";
				break;
			case AscCommon.ST_LayoutShapeType.borderCallout1:
				sVal = "borderCallout1";
				break;
			case AscCommon.ST_LayoutShapeType.borderCallout2:
				sVal = "borderCallout2";
				break;
			case AscCommon.ST_LayoutShapeType.borderCallout3:
				sVal = "borderCallout3";
				break;
			case AscCommon.ST_LayoutShapeType.bracePair:
				sVal = "bracePair";
				break;
			case AscCommon.ST_LayoutShapeType.bracketPair:
				sVal = "bracketPair";
				break;
			case AscCommon.ST_LayoutShapeType.callout1:
				sVal = "callout1";
				break;
			case AscCommon.ST_LayoutShapeType.callout2:
				sVal = "callout2";
				break;
			case AscCommon.ST_LayoutShapeType.callout3:
				sVal = "callout3";
				break;
			case AscCommon.ST_LayoutShapeType.can:
				sVal = "can";
				break;
			case AscCommon.ST_LayoutShapeType.chartPlus:
				sVal = "chartPlus";
				break;
			case AscCommon.ST_LayoutShapeType.chartStar:
				sVal = "chartStar";
				break;
			case AscCommon.ST_LayoutShapeType.chartX:
				sVal = "chartX";
				break;
			case AscCommon.ST_LayoutShapeType.chevron:
				sVal = "chevron";
				break;
			case AscCommon.ST_LayoutShapeType.chord:
				sVal = "chord";
				break;
			case AscCommon.ST_LayoutShapeType.circularArrow:
				sVal = "circularArrow";
				break;
			case AscCommon.ST_LayoutShapeType.cloud:
				sVal = "cloud";
				break;
			case AscCommon.ST_LayoutShapeType.cloudCallout:
				sVal = "cloudCallout";
				break;
			case AscCommon.ST_LayoutShapeType.corner:
				sVal = "corner";
				break;
			case AscCommon.ST_LayoutShapeType.cornerTabs:
				sVal = "cornerTabs";
				break;
			case AscCommon.ST_LayoutShapeType.cube:
				sVal = "cube";
				break;
			case AscCommon.ST_LayoutShapeType.curvedConnector2:
				sVal = "curvedConnector2";
				break;
			case AscCommon.ST_LayoutShapeType.curvedConnector3:
				sVal = "curvedConnector3";
				break;
			case AscCommon.ST_LayoutShapeType.curvedConnector4:
				sVal = "curvedConnector4";
				break;
			case AscCommon.ST_LayoutShapeType.curvedConnector5:
				sVal = "curvedConnector5";
				break;
			case AscCommon.ST_LayoutShapeType.curvedDownArrow:
				sVal = "curvedDownArrow";
				break;
			case AscCommon.ST_LayoutShapeType.curvedLeftArrow:
				sVal = "curvedLeftArrow";
				break;
			case AscCommon.ST_LayoutShapeType.curvedRightArrow:
				sVal = "curvedRightArrow";
				break;
			case AscCommon.ST_LayoutShapeType.curvedUpArrow:
				sVal = "curvedUpArrow";
				break;
			case AscCommon.ST_LayoutShapeType.decagon:
				sVal = "decagon";
				break;
			case AscCommon.ST_LayoutShapeType.diagStripe:
				sVal = "diagStripe";
				break;
			case AscCommon.ST_LayoutShapeType.diamond:
				sVal = "diamond";
				break;
			case AscCommon.ST_LayoutShapeType.dodecagon:
				sVal = "dodecagon";
				break;
			case AscCommon.ST_LayoutShapeType.donut:
				sVal = "donut";
				break;
			case AscCommon.ST_LayoutShapeType.doubleWave:
				sVal = "doubleWave";
				break;
			case AscCommon.ST_LayoutShapeType.downArrow:
				sVal = "downArrow";
				break;
			case AscCommon.ST_LayoutShapeType.downArrowCallout:
				sVal = "downArrowCallout";
				break;
			case AscCommon.ST_LayoutShapeType.ellipse:
				sVal = "ellipse";
				break;
			case AscCommon.ST_LayoutShapeType.ellipseRibbon:
				sVal = "ellipseRibbon";
				break;
			case AscCommon.ST_LayoutShapeType.ellipseRibbon2:
				sVal = "ellipseRibbon2";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartAlternateProcess:
				sVal = "flowChartAlternateProcess";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartCollate:
				sVal = "flowChartCollate";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartConnector:
				sVal = "flowChartConnector";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartDecision:
				sVal = "flowChartDecision";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartDelay:
				sVal = "flowChartDelay";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartDisplay:
				sVal = "flowChartDisplay";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartDocument:
				sVal = "flowChartDocument";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartExtract:
				sVal = "flowChartExtract";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartInputOutput:
				sVal = "flowChartInputOutput";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartInternalStorage:
				sVal = "flowChartInternalStorage";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartMagneticDisk:
				sVal = "flowChartMagneticDisk";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartMagneticDrum:
				sVal = "flowChartMagneticDrum";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartMagneticTape:
				sVal = "flowChartMagneticTape";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartManualInput:
				sVal = "flowChartManualInput";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartManualOperation:
				sVal = "flowChartManualOperation";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartMerge:
				sVal = "flowChartMerge";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartMultidocument:
				sVal = "flowChartMultidocument";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartOfflineStorage:
				sVal = "flowChartOfflineStorage";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartOffpageConnector:
				sVal = "flowChartOffpageConnector";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartOnlineStorage:
				sVal = "flowChartOnlineStorage";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartOr:
				sVal = "flowChartOr";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartPredefinedProcess:
				sVal = "flowChartPredefinedProcess";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartPreparation:
				sVal = "flowChartPreparation";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartProcess:
				sVal = "flowChartProcess";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartPunchedCard:
				sVal = "flowChartPunchedCard";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartPunchedTape:
				sVal = "flowChartPunchedTape";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartSort:
				sVal = "flowChartSort";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartSummingJunction:
				sVal = "flowChartSummingJunction";
				break;
			case AscCommon.ST_LayoutShapeType.flowChartTerminator:
				sVal = "flowChartTerminator";
				break;
			case AscCommon.ST_LayoutShapeType.foldedCorner:
				sVal = "foldedCorner";
				break;
			case AscCommon.ST_LayoutShapeType.frame:
				sVal = "frame";
				break;
			case AscCommon.ST_LayoutShapeType.funnel:
				sVal = "funnel";
				break;
			case AscCommon.ST_LayoutShapeType.gear6:
				sVal = "gear6";
				break;
			case AscCommon.ST_LayoutShapeType.gear9:
				sVal = "gear9";
				break;
			case AscCommon.ST_LayoutShapeType.halfFrame:
				sVal = "halfFrame";
				break;
			case AscCommon.ST_LayoutShapeType.heart:
				sVal = "heart";
				break;
			case AscCommon.ST_LayoutShapeType.heptagon:
				sVal = "heptagon";
				break;
			case AscCommon.ST_LayoutShapeType.hexagon:
				sVal = "hexagon";
				break;
			case AscCommon.ST_LayoutShapeType.homePlate:
				sVal = "homePlate";
				break;
			case AscCommon.ST_LayoutShapeType.horizontalScroll:
				sVal = "horizontalScroll";
				break;
			case AscCommon.ST_LayoutShapeType.irregularSeal1:
				sVal = "irregularSeal1";
				break;
			case AscCommon.ST_LayoutShapeType.irregularSeal2:
				sVal = "irregularSeal2";
				break;
			case AscCommon.ST_LayoutShapeType.leftArrow:
				sVal = "leftArrow";
				break;
			case AscCommon.ST_LayoutShapeType.leftArrowCallout:
				sVal = "leftArrowCallout";
				break;
			case AscCommon.ST_LayoutShapeType.leftBrace:
				sVal = "leftBrace";
				break;
			case AscCommon.ST_LayoutShapeType.leftBracket:
				sVal = "leftBracket";
				break;
			case AscCommon.ST_LayoutShapeType.leftCircularArrow:
				sVal = "leftCircularArrow";
				break;
			case AscCommon.ST_LayoutShapeType.leftRightArrow:
				sVal = "leftRightArrow";
				break;
			case AscCommon.ST_LayoutShapeType.leftRightArrowCallout:
				sVal = "leftRightArrowCallout";
				break;
			case AscCommon.ST_LayoutShapeType.leftRightCircularArrow:
				sVal = "leftRightCircularArrow";
				break;
			case AscCommon.ST_LayoutShapeType.leftRightRibbon:
				sVal = "leftRightRibbon";
				break;
			case AscCommon.ST_LayoutShapeType.leftRightUpArrow:
				sVal = "leftRightUpArrow";
				break;
			case AscCommon.ST_LayoutShapeType.leftUpArrow:
				sVal = "leftUpArrow";
				break;
			case AscCommon.ST_LayoutShapeType.lightningBolt:
				sVal = "lightningBolt";
				break;
			case AscCommon.ST_LayoutShapeType.line:
				sVal = "line";
				break;
			case AscCommon.ST_LayoutShapeType.lineInv:
				sVal = "lineInv";
				break;
			case AscCommon.ST_LayoutShapeType.mathDivide:
				sVal = "mathDivide";
				break;
			case AscCommon.ST_LayoutShapeType.mathEqual:
				sVal = "mathEqual";
				break;
			case AscCommon.ST_LayoutShapeType.mathMinus:
				sVal = "mathMinus";
				break;
			case AscCommon.ST_LayoutShapeType.mathMultiply:
				sVal = "mathMultiply";
				break;
			case AscCommon.ST_LayoutShapeType.mathNotEqual:
				sVal = "mathNotEqual";
				break;
			case AscCommon.ST_LayoutShapeType.mathPlus:
				sVal = "mathPlus";
				break;
			case AscCommon.ST_LayoutShapeType.moon:
				sVal = "moon";
				break;
			case AscCommon.ST_LayoutShapeType.nonIsoscelesTrapezoid:
				sVal = "nonIsoscelesTrapezoid";
				break;
			case AscCommon.ST_LayoutShapeType.noSmoking:
				sVal = "noSmoking";
				break;
			case AscCommon.ST_LayoutShapeType.notchedRightArrow:
				sVal = "notchedRightArrow";
				break;
			case AscCommon.ST_LayoutShapeType.octagon:
				sVal = "octagon";
				break;
			case AscCommon.ST_LayoutShapeType.parallelogram:
				sVal = "parallelogram";
				break;
			case AscCommon.ST_LayoutShapeType.pentagon:
				sVal = "pentagon";
				break;
			case AscCommon.ST_LayoutShapeType.pie:
				sVal = "pie";
				break;
			case AscCommon.ST_LayoutShapeType.pieWedge:
				sVal = "pieWedge";
				break;
			case AscCommon.ST_LayoutShapeType.plaque:
				sVal = "plaque";
				break;
			case AscCommon.ST_LayoutShapeType.plaqueTabs:
				sVal = "plaqueTabs";
				break;
			case AscCommon.ST_LayoutShapeType.plus:
				sVal = "plus";
				break;
			case AscCommon.ST_LayoutShapeType.quadArrow:
				sVal = "quadArrow";
				break;
			case AscCommon.ST_LayoutShapeType.quadArrowCallout:
				sVal = "quadArrowCallout";
				break;
			case AscCommon.ST_LayoutShapeType.rect:
				sVal = "rect";
				break;
			case AscCommon.ST_LayoutShapeType.ribbon:
				sVal = "ribbon";
				break;
			case AscCommon.ST_LayoutShapeType.ribbon2:
				sVal = "ribbon2";
				break;
			case AscCommon.ST_LayoutShapeType.rightArrow:
				sVal = "rightArrow";
				break;
			case AscCommon.ST_LayoutShapeType.rightArrowCallout:
				sVal = "rightArrowCallout";
				break;
			case AscCommon.ST_LayoutShapeType.rightBrace:
				sVal = "rightBrace";
				break;
			case AscCommon.ST_LayoutShapeType.rightBracket:
				sVal = "rightBracket";
				break;
			case AscCommon.ST_LayoutShapeType.round1Rect:
				sVal = "round1Rect";
				break;
			case AscCommon.ST_LayoutShapeType.round2DiagRect:
				sVal = "round2DiagRect";
				break;
			case AscCommon.ST_LayoutShapeType.round2SameRect:
				sVal = "round2SameRect";
				break;
			case AscCommon.ST_LayoutShapeType.roundRect:
				sVal = "roundRect";
				break;
			case AscCommon.ST_LayoutShapeType.rtTriangle:
				sVal = "rtTriangle";
				break;
			case AscCommon.ST_LayoutShapeType.smileyFace:
				sVal = "smileyFace";
				break;
			case AscCommon.ST_LayoutShapeType.snip1Rect:
				sVal = "snip1Rect";
				break;
			case AscCommon.ST_LayoutShapeType.snip2DiagRect:
				sVal = "snip2DiagRect";
				break;
			case AscCommon.ST_LayoutShapeType.snip2SameRect:
				sVal = "snip2SameRect";
				break;
			case AscCommon.ST_LayoutShapeType.snipRoundRect:
				sVal = "snipRoundRect";
				break;
			case AscCommon.ST_LayoutShapeType.squareTabs:
				sVal = "squareTabs";
				break;
			case AscCommon.ST_LayoutShapeType.star10:
				sVal = "star10";
				break;
			case AscCommon.ST_LayoutShapeType.star12:
				sVal = "star12";
				break;
			case AscCommon.ST_LayoutShapeType.star16:
				sVal = "star16";
				break;
			case AscCommon.ST_LayoutShapeType.star24:
				sVal = "star24";
				break;
			case AscCommon.ST_LayoutShapeType.star32:
				sVal = "star32";
				break;
			case AscCommon.ST_LayoutShapeType.star4:
				sVal = "star4";
				break;
			case AscCommon.ST_LayoutShapeType.star5:
				sVal = "star5";
				break;
			case AscCommon.ST_LayoutShapeType.star6:
				sVal = "star6";
				break;
			case AscCommon.ST_LayoutShapeType.star7:
				sVal = "star7";
				break;
			case AscCommon.ST_LayoutShapeType.star8:
				sVal = "star8";
				break;
			case AscCommon.ST_LayoutShapeType.straightConnector1:
				sVal = "straightConnector1";
				break;
			case AscCommon.ST_LayoutShapeType.stripedRightArrow:
				sVal = "stripedRightArrow";
				break;
			case AscCommon.ST_LayoutShapeType.sun:
				sVal = "sun";
				break;
			case AscCommon.ST_LayoutShapeType.swooshArrow:
				sVal = "swooshArrow";
				break;
			case AscCommon.ST_LayoutShapeType.teardrop:
				sVal = "teardrop";
				break;
			case AscCommon.ST_LayoutShapeType.trapezoid:
				sVal = "trapezoid";
				break;
			case AscCommon.ST_LayoutShapeType.triangle:
				sVal = "triangle";
				break;
			case AscCommon.ST_LayoutShapeType.upArrow:
				sVal = "upArrow";
				break;
			case AscCommon.ST_LayoutShapeType.upArrowCallout:
				sVal = "upArrowCallout";
				break;
			case AscCommon.ST_LayoutShapeType.upDownArrow:
				sVal = "upDownArrow";
				break;
			case AscCommon.ST_LayoutShapeType.upDownArrowCallout:
				sVal = "upDownArrowCallout";
				break;
			case AscCommon.ST_LayoutShapeType.uturnArrow:
				sVal = "uturnArrow";
				break;
			case AscCommon.ST_LayoutShapeType.verticalScroll:
				sVal = "verticalScroll";
				break;
			case AscCommon.ST_LayoutShapeType.wave:
				sVal = "wave";
				break;
			case AscCommon.ST_LayoutShapeType.wedgeEllipseCallout:
				sVal = "wedgeEllipseCallout";
				break;
			case AscCommon.ST_LayoutShapeType.wedgeRectCallout:
				sVal = "wedgeRectCallout";
				break;
			case AscCommon.ST_LayoutShapeType.wedgeRoundRectCallout:
				sVal = "wedgeRoundRectCallout";
				break;
		}

		return sVal;
	}
	function From_XML_ST_LayoutShapeType(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "conn":
				nVal = AscCommon.ST_LayoutShapeType.conn;
				break;
			case "none":
				nVal = AscCommon.ST_LayoutShapeType.none;
				break;
			case "accentBorderCallout1":
				nVal = AscCommon.ST_LayoutShapeType.accentBorderCallout1;
				break;
			case "accentBorderCallout2":
				nVal = AscCommon.ST_LayoutShapeType.accentBorderCallout2;
				break;
			case "accentBorderCallout3":
				nVal = AscCommon.ST_LayoutShapeType.accentBorderCallout3;
				break;
			case "accentCallout1":
				nVal = AscCommon.ST_LayoutShapeType.accentCallout1;
				break;
			case "accentCallout2":
				nVal = AscCommon.ST_LayoutShapeType.accentCallout2;
				break;
			case "accentCallout3":
				nVal = AscCommon.ST_LayoutShapeType.accentCallout3;
				break;
			case "actionButtonBackPrevious":
				nVal = AscCommon.ST_LayoutShapeType.actionButtonBackPrevious;
				break;
			case "actionButtonBeginning":
				nVal = AscCommon.ST_LayoutShapeType.actionButtonBeginning;
				break;
			case "actionButtonBlank":
				nVal = AscCommon.ST_LayoutShapeType.actionButtonBlank;
				break;
			case "actionButtonDocument":
				nVal = AscCommon.ST_LayoutShapeType.actionButtonDocument;
				break;
			case "actionButtonEnd":
				nVal = AscCommon.ST_LayoutShapeType.actionButtonEnd;
				break;
			case "actionButtonForwardNext":
				nVal = AscCommon.ST_LayoutShapeType.actionButtonForwardNext;
				break;
			case "actionButtonHelp":
				nVal = AscCommon.ST_LayoutShapeType.actionButtonHelp;
				break;
			case "actionButtonHome":
				nVal = AscCommon.ST_LayoutShapeType.actionButtonHome;
				break;
			case "actionButtonInformation":
				nVal = AscCommon.ST_LayoutShapeType.actionButtonInformation;
				break;
			case "actionButtonMovie":
				nVal = AscCommon.ST_LayoutShapeType.actionButtonMovie;
				break;
			case "actionButtonReturn":
				nVal = AscCommon.ST_LayoutShapeType.actionButtonReturn;
				break;
			case "actionButtonSound":
				nVal = AscCommon.ST_LayoutShapeType.actionButtonSound;
				break;
			case "arc":
				nVal = AscCommon.ST_LayoutShapeType.arc;
				break;
			case "bentArrow":
				nVal = AscCommon.ST_LayoutShapeType.bentArrow;
				break;
			case "bentConnector2":
				nVal = AscCommon.ST_LayoutShapeType.bentConnector2;
				break;
			case "bentConnector3":
				nVal = AscCommon.ST_LayoutShapeType.bentConnector3;
				break;
			case "bentConnector4":
				nVal = AscCommon.ST_LayoutShapeType.bentConnector4;
				break;
			case "bentConnector5":
				nVal = AscCommon.ST_LayoutShapeType.bentConnector5;
				break;
			case "bentUpArrow":
				nVal = AscCommon.ST_LayoutShapeType.bentUpArrow;
				break;
			case "bevel":
				nVal = AscCommon.ST_LayoutShapeType.bevel;
				break;
			case "blockArc":
				nVal = AscCommon.ST_LayoutShapeType.blockArc;
				break;
			case "borderCallout1":
				nVal = AscCommon.ST_LayoutShapeType.borderCallout1;
				break;
			case "borderCallout2":
				nVal = AscCommon.ST_LayoutShapeType.borderCallout2;
				break;
			case "borderCallout3":
				nVal = AscCommon.ST_LayoutShapeType.borderCallout3;
				break;
			case "bracePair":
				nVal = AscCommon.ST_LayoutShapeType.bracePair;
				break;
			case "bracketPair":
				nVal = AscCommon.ST_LayoutShapeType.bracketPair;
				break;
			case "callout1":
				nVal = AscCommon.ST_LayoutShapeType.callout1;
				break;
			case "callout2":
				nVal = AscCommon.ST_LayoutShapeType.callout2;
				break;
			case "callout3":
				nVal = AscCommon.ST_LayoutShapeType.callout3;
				break;
			case "can":
				nVal = AscCommon.ST_LayoutShapeType.can;
				break;
			case "chartPlus":
				nVal = AscCommon.ST_LayoutShapeType.chartPlus;
				break;
			case "chartStar":
				nVal = AscCommon.ST_LayoutShapeType.chartStar;
				break;
			case "chartX":
				nVal = AscCommon.ST_LayoutShapeType.chartX;
				break;
			case "chevron":
				nVal = AscCommon.ST_LayoutShapeType.chevron;
				break;
			case "chord":
				nVal = AscCommon.ST_LayoutShapeType.chord;
				break;
			case "circularArrow":
				nVal = AscCommon.ST_LayoutShapeType.circularArrow;
				break;
			case "cloud":
				nVal = AscCommon.ST_LayoutShapeType.cloud;
				break;
			case "cloudCallout":
				nVal = AscCommon.ST_LayoutShapeType.cloudCallout;
				break;
			case "corner":
				nVal = AscCommon.ST_LayoutShapeType.corner;
				break;
			case "cornerTabs":
				nVal = AscCommon.ST_LayoutShapeType.cornerTabs;
				break;
			case "cube":
				nVal = AscCommon.ST_LayoutShapeType.cube;
				break;
			case "curvedConnector2":
				nVal = AscCommon.ST_LayoutShapeType.curvedConnector2;
				break;
			case "curvedConnector3":
				nVal = AscCommon.ST_LayoutShapeType.curvedConnector3;
				break;
			case "curvedConnector4":
				nVal = AscCommon.ST_LayoutShapeType.curvedConnector4;
				break;
			case "curvedConnector5":
				nVal = AscCommon.ST_LayoutShapeType.curvedConnector5;
				break;
			case "curvedDownArrow":
				nVal = AscCommon.ST_LayoutShapeType.curvedDownArrow;
				break;
			case "curvedLeftArrow":
				nVal = AscCommon.ST_LayoutShapeType.curvedLeftArrow;
				break;
			case "curvedRightArrow":
				nVal = AscCommon.ST_LayoutShapeType.curvedRightArrow;
				break;
			case "curvedUpArrow":
				nVal = AscCommon.ST_LayoutShapeType.curvedUpArrow;
				break;
			case "decagon":
				nVal = AscCommon.ST_LayoutShapeType.decagon;
				break;
			case "diagStripe":
				nVal = AscCommon.ST_LayoutShapeType.diagStripe;
				break;
			case "diamond":
				nVal = AscCommon.ST_LayoutShapeType.diamond;
				break;
			case "dodecagon":
				nVal = AscCommon.ST_LayoutShapeType.dodecagon;
				break;
			case "donut":
				nVal = AscCommon.ST_LayoutShapeType.donut;
				break;
			case "doubleWave":
				nVal = AscCommon.ST_LayoutShapeType.doubleWave;
				break;
			case "downArrow":
				nVal = AscCommon.ST_LayoutShapeType.downArrow;
				break;
			case "downArrowCallout":
				nVal = AscCommon.ST_LayoutShapeType.downArrowCallout;
				break;
			case "ellipse":
				nVal = AscCommon.ST_LayoutShapeType.ellipse;
				break;
			case "ellipseRibbon":
				nVal = AscCommon.ST_LayoutShapeType.ellipseRibbon;
				break;
			case "ellipseRibbon2":
				nVal = AscCommon.ST_LayoutShapeType.ellipseRibbon2;
				break;
			case "flowChartAlternateProcess":
				nVal = AscCommon.ST_LayoutShapeType.flowChartAlternateProcess;
				break;
			case "flowChartCollate":
				nVal = AscCommon.ST_LayoutShapeType.flowChartCollate;
				break;
			case "flowChartConnector":
				nVal = AscCommon.ST_LayoutShapeType.flowChartConnector;
				break;
			case "flowChartDecision":
				nVal = AscCommon.ST_LayoutShapeType.flowChartDecision;
				break;
			case "flowChartDelay":
				nVal = AscCommon.ST_LayoutShapeType.flowChartDelay;
				break;
			case "flowChartDisplay":
				nVal = AscCommon.ST_LayoutShapeType.flowChartDisplay;
				break;
			case "flowChartDocument":
				nVal = AscCommon.ST_LayoutShapeType.flowChartDocument;
				break;
			case "flowChartExtract":
				nVal = AscCommon.ST_LayoutShapeType.flowChartExtract;
				break;
			case "flowChartInputOutput":
				nVal = AscCommon.ST_LayoutShapeType.flowChartInputOutput;
				break;
			case "flowChartInternalStorage":
				nVal = AscCommon.ST_LayoutShapeType.flowChartInternalStorage;
				break;
			case "flowChartMagneticDisk":
				nVal = AscCommon.ST_LayoutShapeType.flowChartMagneticDisk;
				break;
			case "flowChartMagneticDrum":
				nVal = AscCommon.ST_LayoutShapeType.flowChartMagneticDrum;
				break;
			case "flowChartMagneticTape":
				nVal = AscCommon.ST_LayoutShapeType.flowChartMagneticTape;
				break;
			case "flowChartManualInput":
				nVal = AscCommon.ST_LayoutShapeType.flowChartManualInput;
				break;
			case "flowChartManualOperation":
				nVal = AscCommon.ST_LayoutShapeType.flowChartManualOperation;
				break;
			case "flowChartMerge":
				nVal = AscCommon.ST_LayoutShapeType.flowChartMerge;
				break;
			case "flowChartMultidocument":
				nVal = AscCommon.ST_LayoutShapeType.flowChartMultidocument;
				break;
			case "flowChartOfflineStorage":
				nVal = AscCommon.ST_LayoutShapeType.flowChartOfflineStorage;
				break;
			case "flowChartOffpageConnector":
				nVal = AscCommon.ST_LayoutShapeType.flowChartOffpageConnector;
				break;
			case "flowChartOnlineStorage":
				nVal = AscCommon.ST_LayoutShapeType.flowChartOnlineStorage;
				break;
			case "flowChartOr":
				nVal = AscCommon.ST_LayoutShapeType.flowChartOr;
				break;
			case "flowChartPredefinedProcess":
				nVal = AscCommon.ST_LayoutShapeType.flowChartPredefinedProcess;
				break;
			case "flowChartPreparation":
				nVal = AscCommon.ST_LayoutShapeType.flowChartPreparation;
				break;
			case "flowChartProcess":
				nVal = AscCommon.ST_LayoutShapeType.flowChartProcess;
				break;
			case "flowChartPunchedCard":
				nVal = AscCommon.ST_LayoutShapeType.flowChartPunchedCard;
				break;
			case "flowChartPunchedTape":
				nVal = AscCommon.ST_LayoutShapeType.flowChartPunchedTape;
				break;
			case "flowChartSort":
				nVal = AscCommon.ST_LayoutShapeType.flowChartSort;
				break;
			case "flowChartSummingJunction":
				nVal = AscCommon.ST_LayoutShapeType.flowChartSummingJunction;
				break;
			case "flowChartTerminator":
				nVal = AscCommon.ST_LayoutShapeType.flowChartTerminator;
				break;
			case "foldedCorner":
				nVal = AscCommon.ST_LayoutShapeType.foldedCorner;
				break;
			case "frame":
				nVal = AscCommon.ST_LayoutShapeType.frame;
				break;
			case "funnel":
				nVal = AscCommon.ST_LayoutShapeType.funnel;
				break;
			case "gear6":
				nVal = AscCommon.ST_LayoutShapeType.gear6;
				break;
			case "gear9":
				nVal = AscCommon.ST_LayoutShapeType.gear9;
				break;
			case "halfFrame":
				nVal = AscCommon.ST_LayoutShapeType.halfFrame;
				break;
			case "heart":
				nVal = AscCommon.ST_LayoutShapeType.heart;
				break;
			case "heptagon":
				nVal = AscCommon.ST_LayoutShapeType.heptagon;
				break;
			case "hexagon":
				nVal = AscCommon.ST_LayoutShapeType.hexagon;
				break;
			case "homePlate":
				nVal = AscCommon.ST_LayoutShapeType.homePlate;
				break;
			case "horizontalScroll":
				nVal = AscCommon.ST_LayoutShapeType.horizontalScroll;
				break;
			case "irregularSeal1":
				nVal = AscCommon.ST_LayoutShapeType.irregularSeal1;
				break;
			case "irregularSeal2":
				nVal = AscCommon.ST_LayoutShapeType.irregularSeal2;
				break;
			case "leftArrow":
				nVal = AscCommon.ST_LayoutShapeType.leftArrow;
				break;
			case "leftArrowCallout":
				nVal = AscCommon.ST_LayoutShapeType.leftArrowCallout;
				break;
			case "leftBrace":
				nVal = AscCommon.ST_LayoutShapeType.leftBrace;
				break;
			case "leftBracket":
				nVal = AscCommon.ST_LayoutShapeType.leftBracket;
				break;
			case "leftCircularArrow":
				nVal = AscCommon.ST_LayoutShapeType.leftCircularArrow;
				break;
			case "leftRightArrow":
				nVal = AscCommon.ST_LayoutShapeType.leftRightArrow;
				break;
			case "leftRightArrowCallout":
				nVal = AscCommon.ST_LayoutShapeType.leftRightArrowCallout;
				break;
			case "leftRightCircularArrow":
				nVal = AscCommon.ST_LayoutShapeType.leftRightCircularArrow;
				break;
			case "leftRightRibbon":
				nVal = AscCommon.ST_LayoutShapeType.leftRightRibbon;
				break;
			case "leftRightUpArrow":
				nVal = AscCommon.ST_LayoutShapeType.leftRightUpArrow;
				break;
			case "leftUpArrow":
				nVal = AscCommon.ST_LayoutShapeType.leftUpArrow;
				break;
			case "lightningBolt":
				nVal = AscCommon.ST_LayoutShapeType.lightningBolt;
				break;
			case "line":
				nVal = AscCommon.ST_LayoutShapeType.line;
				break;
			case "lineInv":
				nVal = AscCommon.ST_LayoutShapeType.lineInv;
				break;
			case "mathDivide":
				nVal = AscCommon.ST_LayoutShapeType.mathDivide;
				break;
			case "mathEqual":
				nVal = AscCommon.ST_LayoutShapeType.mathEqual;
				break;
			case "mathMinus":
				nVal = AscCommon.ST_LayoutShapeType.mathMinus;
				break;
			case "mathMultiply":
				nVal = AscCommon.ST_LayoutShapeType.mathMultiply;
				break;
			case "mathNotEqual":
				nVal = AscCommon.ST_LayoutShapeType.mathNotEqual;
				break;
			case "mathPlus":
				nVal = AscCommon.ST_LayoutShapeType.mathPlus;
				break;
			case "moon":
				nVal = AscCommon.ST_LayoutShapeType.moon;
				break;
			case "nonIsoscelesTrapezoid":
				nVal = AscCommon.ST_LayoutShapeType.nonIsoscelesTrapezoid;
				break;
			case "noSmoking":
				nVal = AscCommon.ST_LayoutShapeType.noSmoking;
				break;
			case "notchedRightArrow":
				nVal = AscCommon.ST_LayoutShapeType.notchedRightArrow;
				break;
			case "octagon":
				nVal = AscCommon.ST_LayoutShapeType.octagon;
				break;
			case "parallelogram":
				nVal = AscCommon.ST_LayoutShapeType.parallelogram;
				break;
			case "pentagon":
				nVal = AscCommon.ST_LayoutShapeType.pentagon;
				break;
			case "pie":
				nVal = AscCommon.ST_LayoutShapeType.pie;
				break;
			case "pieWedge":
				nVal = AscCommon.ST_LayoutShapeType.pieWedge;
				break;
			case "plaque":
				nVal = AscCommon.ST_LayoutShapeType.plaque;
				break;
			case "plaqueTabs":
				nVal = AscCommon.ST_LayoutShapeType.plaqueTabs;
				break;
			case "plus":
				nVal = AscCommon.ST_LayoutShapeType.plus;
				break;
			case "quadArrow":
				nVal = AscCommon.ST_LayoutShapeType.quadArrow;
				break;
			case "quadArrowCallout":
				nVal = AscCommon.ST_LayoutShapeType.quadArrowCallout;
				break;
			case "rect":
				nVal = AscCommon.ST_LayoutShapeType.rect;
				break;
			case "ribbon":
				nVal = AscCommon.ST_LayoutShapeType.ribbon;
				break;
			case "ribbon2":
				nVal = AscCommon.ST_LayoutShapeType.ribbon2;
				break;
			case "rightArrow":
				nVal = AscCommon.ST_LayoutShapeType.rightArrow;
				break;
			case "rightArrowCallout":
				nVal = AscCommon.ST_LayoutShapeType.rightArrowCallout;
				break;
			case "rightBrace":
				nVal = AscCommon.ST_LayoutShapeType.rightBrace;
				break;
			case "rightBracket":
				nVal = AscCommon.ST_LayoutShapeType.rightBracket;
				break;
			case "round1Rect":
				nVal = AscCommon.ST_LayoutShapeType.round1Rect;
				break;
			case "round2DiagRect":
				nVal = AscCommon.ST_LayoutShapeType.round2DiagRect;
				break;
			case "round2SameRect":
				nVal = AscCommon.ST_LayoutShapeType.round2SameRect;
				break;
			case "roundRect":
				nVal = AscCommon.ST_LayoutShapeType.roundRect;
				break;
			case "rtTriangle":
				nVal = AscCommon.ST_LayoutShapeType.rtTriangle;
				break;
			case "smileyFace":
				nVal = AscCommon.ST_LayoutShapeType.smileyFace;
				break;
			case "snip1Rect":
				nVal = AscCommon.ST_LayoutShapeType.snip1Rect;
				break;
			case "snip2DiagRect":
				nVal = AscCommon.ST_LayoutShapeType.snip2DiagRect;
				break;
			case "snip2SameRect":
				nVal = AscCommon.ST_LayoutShapeType.snip2SameRect;
				break;
			case "snipRoundRect":
				nVal = AscCommon.ST_LayoutShapeType.snipRoundRect;
				break;
			case "squareTabs":
				nVal = AscCommon.ST_LayoutShapeType.squareTabs;
				break;
			case "star10":
				nVal = AscCommon.ST_LayoutShapeType.star10;
				break;
			case "star12":
				nVal = AscCommon.ST_LayoutShapeType.star12;
				break;
			case "star16":
				nVal = AscCommon.ST_LayoutShapeType.star16;
				break;
			case "star24":
				nVal = AscCommon.ST_LayoutShapeType.star24;
				break;
			case "star32":
				nVal = AscCommon.ST_LayoutShapeType.star32;
				break;
			case "star4":
				nVal = AscCommon.ST_LayoutShapeType.star4;
				break;
			case "star5":
				nVal = AscCommon.ST_LayoutShapeType.star5;
				break;
			case "star6":
				nVal = AscCommon.ST_LayoutShapeType.star6;
				break;
			case "star7":
				nVal = AscCommon.ST_LayoutShapeType.star7;
				break;
			case "star8":
				nVal = AscCommon.ST_LayoutShapeType.star8;
				break;
			case "straightConnector1":
				nVal = AscCommon.ST_LayoutShapeType.straightConnector1;
				break;
			case "stripedRightArrow":
				nVal = AscCommon.ST_LayoutShapeType.stripedRightArrow;
				break;
			case "sun":
				nVal = AscCommon.ST_LayoutShapeType.sun;
				break;
			case "swooshArrow":
				nVal = AscCommon.ST_LayoutShapeType.swooshArrow;
				break;
			case "teardrop":
				nVal = AscCommon.ST_LayoutShapeType.teardrop;
				break;
			case "trapezoid":
				nVal = AscCommon.ST_LayoutShapeType.trapezoid;
				break;
			case "triangle":
				nVal = AscCommon.ST_LayoutShapeType.triangle;
				break;
			case "upArrow":
				nVal = AscCommon.ST_LayoutShapeType.upArrow;
				break;
			case "upArrowCallout":
				nVal = AscCommon.ST_LayoutShapeType.upArrowCallout;
				break;
			case "upDownArrow":
				nVal = AscCommon.ST_LayoutShapeType.upDownArrow;
				break;
			case "upDownArrowCallout":
				nVal = AscCommon.ST_LayoutShapeType.upDownArrowCallout;
				break;
			case "uturnArrow":
				nVal = AscCommon.ST_LayoutShapeType.uturnArrow;
				break;
			case "verticalScroll":
				nVal = AscCommon.ST_LayoutShapeType.verticalScroll;
				break;
			case "wave":
				nVal = AscCommon.ST_LayoutShapeType.wave;
				break;
			case "wedgeEllipseCallout":
				nVal = AscCommon.ST_LayoutShapeType.wedgeEllipseCallout;
				break;
			case "wedgeRectCallout":
				nVal = AscCommon.ST_LayoutShapeType.wedgeRectCallout;
				break;
			case "wedgeRoundRectCallout":
				nVal = AscCommon.ST_LayoutShapeType.wedgeRoundRectCallout;
				break;
		}

		return nVal;
	}
	
	function To_XML_ST_ParameterId(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case AscCommon.ST_ParameterId.alignTx:
				sVal = "alignTx";
				break;
			case AscCommon.ST_ParameterId.ar:
				sVal = "ar";
				break;
			case AscCommon.ST_ParameterId.autoTxRot:
				sVal = "autoTxRot";
				break;
			case AscCommon.ST_ParameterId.begPts:
				sVal = "begPts";
				break;
			case AscCommon.ST_ParameterId.begSty:
				sVal = "begSty";
				break;
			case AscCommon.ST_ParameterId.bendPt:
				sVal = "bendPt";
				break;
			case AscCommon.ST_ParameterId.bkpt:
				sVal = "bkpt";
				break;
			case AscCommon.ST_ParameterId.bkPtFixedVal:
				sVal = "bkPtFixedVal";
				break;
			case AscCommon.ST_ParameterId.chAlign:
				sVal = "chAlign";
				break;
			case AscCommon.ST_ParameterId.chDir:
				sVal = "chDir";
				break;
			case AscCommon.ST_ParameterId.connRout:
				sVal = "connRout";
				break;
			case AscCommon.ST_ParameterId.contDir:
				sVal = "contDir";
				break;
			case AscCommon.ST_ParameterId.ctrShpMap:
				sVal = "ctrShpMap";
				break;
			case AscCommon.ST_ParameterId.dim:
				sVal = "dim";
				break;
			case AscCommon.ST_ParameterId.dstNode:
				sVal = "dstNode";
				break;
			case AscCommon.ST_ParameterId.endPts:
				sVal = "endPts";
				break;
			case AscCommon.ST_ParameterId.endSty:
				sVal = "endSty";
				break;
			case AscCommon.ST_ParameterId.fallback:
				sVal = "fallback";
				break;
			case AscCommon.ST_ParameterId.flowDir:
				sVal = "flowDir";
				break;
			case AscCommon.ST_ParameterId.grDir:
				sVal = "grDir";
				break;
			case AscCommon.ST_ParameterId.hierAlign:
				sVal = "hierAlign";
				break;
			case AscCommon.ST_ParameterId.horzAlign:
				sVal = "horzAlign";
				break;
			case AscCommon.ST_ParameterId.linDir:
				sVal = "linDir";
				break;
			case AscCommon.ST_ParameterId.lnSpAfChP:
				sVal = "lnSpAfChP";
				break;
			case AscCommon.ST_ParameterId.lnSpAfParP:
				sVal = "lnSpAfParP";
				break;
			case AscCommon.ST_ParameterId.lnSpCh:
				sVal = "lnSpCh";
				break;
			case AscCommon.ST_ParameterId.lnSpPar:
				sVal = "lnSpPar";
				break;
			case AscCommon.ST_ParameterId.nodeHorzAlign:
				sVal = "nodeHorzAlign";
				break;
			case AscCommon.ST_ParameterId.nodeVertAlign:
				sVal = "nodeVertAlign";
				break;
			case AscCommon.ST_ParameterId.off:
				sVal = "off";
				break;
			case AscCommon.ST_ParameterId.parTxLTRAlign:
				sVal = "parTxLTRAlign";
				break;
			case AscCommon.ST_ParameterId.parTxRTLAlign:
				sVal = "parTxRTLAlign";
				break;
			case AscCommon.ST_ParameterId.pyraAcctBkgdNode:
				sVal = "pyraAcctBkgdNode";
				break;
			case AscCommon.ST_ParameterId.pyraAcctPos:
				sVal = "pyraAcctPos";
				break;
			case AscCommon.ST_ParameterId.pyraAcctTxMar:
				sVal = "pyraAcctTxMar";
				break;
			case AscCommon.ST_ParameterId.pyraAcctTxNode:
				sVal = "pyraAcctTxNode";
				break;
			case AscCommon.ST_ParameterId.pyraLvlNode:
				sVal = "pyraLvlNode";
				break;
			case AscCommon.ST_ParameterId.rotPath:
				sVal = "rotPath";
				break;
			case AscCommon.ST_ParameterId.rtShortDist:
				sVal = "rtShortDist";
				break;
			case AscCommon.ST_ParameterId.secChAlign:
				sVal = "secChAlign";
				break;
			case AscCommon.ST_ParameterId.secLinDir:
				sVal = "secLinDir";
				break;
			case AscCommon.ST_ParameterId.shpTxLTRAlignCh:
				sVal = "shpTxLTRAlignCh";
				break;
			case AscCommon.ST_ParameterId.shpTxRTLAlignCh:
				sVal = "shpTxRTLAlignCh";
				break;
			case AscCommon.ST_ParameterId.spanAng:
				sVal = "spanAng";
				break;
			case AscCommon.ST_ParameterId.srcNode:
				sVal = "srcNode";
				break;
			case AscCommon.ST_ParameterId.stAng:
				sVal = "stAng";
				break;
			case AscCommon.ST_ParameterId.stBulletLvl:
				sVal = "stBulletLvl";
				break;
			case AscCommon.ST_ParameterId.stElem:
				sVal = "stElem";
				break;
			case AscCommon.ST_ParameterId.txAnchorHorz:
				sVal = "txAnchorHorz";
				break;
			case AscCommon.ST_ParameterId.txAnchorHorzCh:
				sVal = "txAnchorHorzCh";
				break;
			case AscCommon.ST_ParameterId.txAnchorVert:
				sVal = "txAnchorVert";
				break;
			case AscCommon.ST_ParameterId.txAnchorVertCh:
				sVal = "txAnchorVertCh";
				break;
			case AscCommon.ST_ParameterId.txBlDir:
				sVal = "txBlDir";
				break;
			case AscCommon.ST_ParameterId.txDir:
				sVal = "txDir";
				break;
			case AscCommon.ST_ParameterId.vertAlign:
				sVal = "vertAlign";
				break;
		}

		return sVal;
	}
	function From_XML_ST_ParameterId(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "alignTx":
				nVal = AscCommon.ST_ParameterId.alignTx;
				break;
			case "ar":
				nVal = AscCommon.ST_ParameterId.ar;
				break;
			case "autoTxRot":
				nVal = AscCommon.ST_ParameterId.autoTxRot;
				break;
			case "begPts":
				nVal = AscCommon.ST_ParameterId.begPts;
				break;
			case "begSty":
				nVal = AscCommon.ST_ParameterId.begSty;
				break;
			case "bendPt":
				nVal = AscCommon.ST_ParameterId.bendPt;
				break;
			case "bkpt":
				nVal = AscCommon.ST_ParameterId.bkpt;
				break;
			case "bkPtFixedVal":
				nVal = AscCommon.ST_ParameterId.bkPtFixedVal;
				break;
			case "chAlign":
				nVal = AscCommon.ST_ParameterId.chAlign;
				break;
			case "chDir":
				nVal = AscCommon.ST_ParameterId.chDir;
				break;
			case "connRout":
				nVal = AscCommon.ST_ParameterId.connRout;
				break;
			case "contDir":
				nVal = AscCommon.ST_ParameterId.contDir;
				break;
			case "ctrShpMap":
				nVal = AscCommon.ST_ParameterId.ctrShpMap;
				break;
			case "dim":
				nVal = AscCommon.ST_ParameterId.dim;
				break;
			case "dstNode":
				nVal = AscCommon.ST_ParameterId.dstNode;
				break;
			case "endPts":
				nVal = AscCommon.ST_ParameterId.endPts;
				break;
			case "endSty":
				nVal = AscCommon.ST_ParameterId.endSty;
				break;
			case "fallback":
				nVal = AscCommon.ST_ParameterId.fallback;
				break;
			case "flowDir":
				nVal = AscCommon.ST_ParameterId.flowDir;
				break;
			case "grDir":
				nVal = AscCommon.ST_ParameterId.grDir;
				break;
			case "hierAlign":
				nVal = AscCommon.ST_ParameterId.hierAlign;
				break;
			case "horzAlign":
				nVal = AscCommon.ST_ParameterId.horzAlign;
				break;
			case "linDir":
				nVal = AscCommon.ST_ParameterId.linDir;
				break;
			case "lnSpAfChP":
				nVal = AscCommon.ST_ParameterId.lnSpAfChP;
				break;
			case "lnSpAfParP":
				nVal = AscCommon.ST_ParameterId.lnSpAfParP;
				break;
			case "lnSpCh":
				nVal = AscCommon.ST_ParameterId.lnSpCh;
				break;
			case "lnSpPar":
				nVal = AscCommon.ST_ParameterId.lnSpPar;
				break;
			case "nodeHorzAlign":
				nVal = AscCommon.ST_ParameterId.nodeHorzAlign;
				break;
			case "nodeVertAlign":
				nVal = AscCommon.ST_ParameterId.nodeVertAlign;
				break;
			case "off":
				nVal = AscCommon.ST_ParameterId.off;
				break;
			case "parTxLTRAlign":
				nVal = AscCommon.ST_ParameterId.parTxLTRAlign;
				break;
			case "parTxRTLAlign":
				nVal = AscCommon.ST_ParameterId.parTxRTLAlign;
				break;
			case "pyraAcctBkgdNode":
				nVal = AscCommon.ST_ParameterId.pyraAcctBkgdNode;
				break;
			case "pyraAcctPos":
				nVal = AscCommon.ST_ParameterId.pyraAcctPos;
				break;
			case "pyraAcctTxMar":
				nVal = AscCommon.ST_ParameterId.pyraAcctTxMar;
				break;
			case "pyraAcctTxNode":
				nVal = AscCommon.ST_ParameterId.pyraAcctTxNode;
				break;
			case "pyraLvlNode":
				nVal = AscCommon.ST_ParameterId.pyraLvlNode;
				break;
			case "rotPath":
				nVal = AscCommon.ST_ParameterId.rotPath;
				break;
			case "rtShortDist":
				nVal = AscCommon.ST_ParameterId.rtShortDist;
				break;
			case "secChAlign":
				nVal = AscCommon.ST_ParameterId.secChAlign;
				break;
			case "secLinDir":
				nVal = AscCommon.ST_ParameterId.secLinDir;
				break;
			case "shpTxLTRAlignCh":
				nVal = AscCommon.ST_ParameterId.shpTxLTRAlignCh;
				break;
			case "shpTxRTLAlignCh":
				nVal = AscCommon.ST_ParameterId.shpTxRTLAlignCh;
				break;
			case "spanAng":
				nVal = AscCommon.ST_ParameterId.spanAng;
				break;
			case "srcNode":
				nVal = AscCommon.ST_ParameterId.srcNode;
				break;
			case "stAng":
				nVal = AscCommon.ST_ParameterId.stAng;
				break;
			case "stBulletLvl":
				nVal = AscCommon.ST_ParameterId.stBulletLvl;
				break;
			case "stElem":
				nVal = AscCommon.ST_ParameterId.stElem;
				break;
			case "txAnchorHorz":
				nVal = AscCommon.ST_ParameterId.txAnchorHorz;
				break;
			case "txAnchorHorzCh":
				nVal = AscCommon.ST_ParameterId.txAnchorHorzCh;
				break;
			case "txAnchorVert":
				nVal = AscCommon.ST_ParameterId.txAnchorVert;
				break;
			case "txAnchorVertCh":
				nVal = AscCommon.ST_ParameterId.txAnchorVertCh;
				break;
			case "txBlDir":
				nVal = AscCommon.ST_ParameterId.txBlDir;
				break;
			case "txDir":
				nVal = AscCommon.ST_ParameterId.txDir;
				break;
			case "vertAlign":
				nVal = AscCommon.ST_ParameterId.vertAlign;
				break;
		}

		return nVal;
	}
	function To_XML_ST_PresetCameraType(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case AscCommon.ST_PresetCameraType.isometricBottomDown:
				sVal = "isometricBottomDown";
				break;
			case AscCommon.ST_PresetCameraType.isometricBottomUp:
				sVal = "isometricBottomUp";
				break;
			case AscCommon.ST_PresetCameraType.isometricLeftDown:
				sVal = "isometricLeftDown";
				break;
			case AscCommon.ST_PresetCameraType.isometricLeftUp:
				sVal = "isometricLeftUp";
				break;
			case AscCommon.ST_PresetCameraType.isometricOffAxis1Left:
				sVal = "isometricOffAxis1Left";
				break;
			case AscCommon.ST_PresetCameraType.isometricOffAxis1Right:
				sVal = "isometricOffAxis1Right";
				break;
			case AscCommon.ST_PresetCameraType.isometricOffAxis1Top:
				sVal = "isometricOffAxis1Top";
				break;
			case AscCommon.ST_PresetCameraType.isometricOffAxis2Left:
				sVal = "isometricOffAxis2Left";
				break;
			case AscCommon.ST_PresetCameraType.isometricOffAxis2Right:
				sVal = "isometricOffAxis2Right";
				break;
			case AscCommon.ST_PresetCameraType.isometricOffAxis2Top:
				sVal = "isometricOffAxis2Top";
				break;
			case AscCommon.ST_PresetCameraType.isometricOffAxis3Bottom:
				sVal = "isometricOffAxis3Bottom";
				break;
			case AscCommon.ST_PresetCameraType.isometricOffAxis3Left:
				sVal = "isometricOffAxis3Left";
				break;
			case AscCommon.ST_PresetCameraType.isometricOffAxis3Right:
				sVal = "isometricOffAxis3Right";
				break;
			case AscCommon.ST_PresetCameraType.isometricOffAxis4Bottom:
				sVal = "isometricOffAxis4Bottom";
				break;
			case AscCommon.ST_PresetCameraType.isometricOffAxis4Left:
				sVal = "isometricOffAxis4Left";
				break;
			case AscCommon.ST_PresetCameraType.isometricOffAxis4Right:
				sVal = "isometricOffAxis4Right";
				break;
			case AscCommon.ST_PresetCameraType.isometricRightDown:
				sVal = "isometricRightDown";
				break;
			case AscCommon.ST_PresetCameraType.isometricRightUp:
				sVal = "isometricRightUp";
				break;
			case AscCommon.ST_PresetCameraType.isometricTopDown:
				sVal = "isometricTopDown";
				break;
			case AscCommon.ST_PresetCameraType.isometricTopUp:
				sVal = "isometricTopUp";
				break;
			case AscCommon.ST_PresetCameraType.legacyObliqueBottom:
				sVal = "legacyObliqueBottom";
				break;
			case AscCommon.ST_PresetCameraType.legacyObliqueBottomLeft:
				sVal = "legacyObliqueBottomLeft";
				break;
			case AscCommon.ST_PresetCameraType.legacyObliqueBottomRight:
				sVal = "legacyObliqueBottomRight";
				break;
			case AscCommon.ST_PresetCameraType.legacyObliqueFront:
				sVal = "legacyObliqueFront";
				break;
			case AscCommon.ST_PresetCameraType.legacyObliqueLeft:
				sVal = "legacyObliqueLeft";
				break;
			case AscCommon.ST_PresetCameraType.legacyObliqueRight:
				sVal = "legacyObliqueRight";
				break;
			case AscCommon.ST_PresetCameraType.legacyObliqueTop:
				sVal = "legacyObliqueTop";
				break;
			case AscCommon.ST_PresetCameraType.legacyObliqueTopLeft:
				sVal = "legacyObliqueTopLeft";
				break;
			case AscCommon.ST_PresetCameraType.legacyObliqueTopRight:
				sVal = "legacyObliqueTopRight";
				break;
			case AscCommon.ST_PresetCameraType.legacyPerspectiveBottom:
				sVal = "legacyPerspectiveBottom";
				break;
			case AscCommon.ST_PresetCameraType.legacyPerspectiveBottomLeft:
				sVal = "legacyPerspectiveBottomLeft";
				break;
			case AscCommon.ST_PresetCameraType.legacyPerspectiveBottomRight:
				sVal = "legacyPerspectiveBottomRight";
				break;
			case AscCommon.ST_PresetCameraType.legacyPerspectiveFront:
				sVal = "legacyPerspectiveFront";
				break;
			case AscCommon.ST_PresetCameraType.legacyPerspectiveLeft:
				sVal = "legacyPerspectiveLeft";
				break;
			case AscCommon.ST_PresetCameraType.legacyPerspectiveRight:
				sVal = "legacyPerspectiveRight";
				break;
			case AscCommon.ST_PresetCameraType.legacyPerspectiveTop:
				sVal = "legacyPerspectiveTop";
				break;
			case AscCommon.ST_PresetCameraType.legacyPerspectiveTopLeft:
				sVal = "legacyPerspectiveTopLeft";
				break;
			case AscCommon.ST_PresetCameraType.legacyPerspectiveTopRight:
				sVal = "legacyPerspectiveTopRight";
				break;
			case AscCommon.ST_PresetCameraType.obliqueBottom:
				sVal = "obliqueBottom";
				break;
			case AscCommon.ST_PresetCameraType.obliqueBottomLeft:
				sVal = "obliqueBottomLeft";
				break;
			case AscCommon.ST_PresetCameraType.obliqueBottomRight:
				sVal = "obliqueBottomRight";
				break;
			case AscCommon.ST_PresetCameraType.obliqueLeft:
				sVal = "obliqueLeft";
				break;
			case AscCommon.ST_PresetCameraType.obliqueRight:
				sVal = "obliqueRight";
				break;
			case AscCommon.ST_PresetCameraType.obliqueTop:
				sVal = "obliqueTop";
				break;
			case AscCommon.ST_PresetCameraType.obliqueTopLeft:
				sVal = "obliqueTopLeft";
				break;
			case AscCommon.ST_PresetCameraType.obliqueTopRight:
				sVal = "obliqueTopRight";
				break;
			case AscCommon.ST_PresetCameraType.orthographicFront:
				sVal = "orthographicFront";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveAbove:
				sVal = "perspectiveAbove";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveAboveLeftFacing:
				sVal = "perspectiveAboveLeftFacing";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveAboveRightFacing:
				sVal = "perspectiveAboveRightFacing";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveBelow:
				sVal = "perspectiveBelow";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveContrastingLeftFacing:
				sVal = "perspectiveContrastingLeftFacing";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveContrastingRightFacing:
				sVal = "perspectiveContrastingRightFacing";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveFront:
				sVal = "perspectiveFront";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveHeroicExtremeLeftFacing:
				sVal = "perspectiveHeroicExtremeLeftFacing";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveHeroicExtremeRightFacing:
				sVal = "perspectiveHeroicExtremeRightFacing";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveHeroicLeftFacing:
				sVal = "perspectiveHeroicLeftFacing";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveHeroicRightFacing:
				sVal = "perspectiveHeroicRightFacing";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveLeft:
				sVal = "perspectiveLeft";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveRelaxed:
				sVal = "perspectiveRelaxed";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveRelaxedModerately:
				sVal = "perspectiveRelaxedModerately";
				break;
			case AscCommon.ST_PresetCameraType.perspectiveRight:
				sVal = "perspectiveRight";
				break;
		}

		return sVal;
	}
	function From_XML_ST_PresetCameraType(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "isometricBottomDown":
				nVal = AscCommon.ST_PresetCameraType.isometricBottomDown;
				break;
			case "isometricBottomUp":
				nVal = AscCommon.ST_PresetCameraType.isometricBottomUp;
				break;
			case "isometricLeftDown":
				nVal = AscCommon.ST_PresetCameraType.isometricLeftDown;
				break;
			case "isometricLeftUp":
				nVal = AscCommon.ST_PresetCameraType.isometricLeftUp;
				break;
			case "isometricOffAxis1Left":
				nVal = AscCommon.ST_PresetCameraType.isometricOffAxis1Left;
				break;
			case "isometricOffAxis1Right":
				nVal = AscCommon.ST_PresetCameraType.isometricOffAxis1Right;
				break;
			case "isometricOffAxis1Top":
				nVal = AscCommon.ST_PresetCameraType.isometricOffAxis1Top;
				break;
			case "isometricOffAxis2Left":
				nVal = AscCommon.ST_PresetCameraType.isometricOffAxis2Left;
				break;
			case "isometricOffAxis2Right":
				nVal = AscCommon.ST_PresetCameraType.isometricOffAxis2Right;
				break;
			case "isometricOffAxis2Top":
				nVal = AscCommon.ST_PresetCameraType.isometricOffAxis2Top;
				break;
			case "isometricOffAxis3Bottom":
				nVal = AscCommon.ST_PresetCameraType.isometricOffAxis3Bottom;
				break;
			case "isometricOffAxis3Left":
				nVal = AscCommon.ST_PresetCameraType.isometricOffAxis3Left;
				break;
			case "isometricOffAxis3Right":
				nVal = AscCommon.ST_PresetCameraType.isometricOffAxis3Right;
				break;
			case "isometricOffAxis4Bottom":
				nVal = AscCommon.ST_PresetCameraType.isometricOffAxis4Bottom;
				break;
			case "isometricOffAxis4Left":
				nVal = AscCommon.ST_PresetCameraType.isometricOffAxis4Left;
				break;
			case "isometricOffAxis4Right":
				nVal = AscCommon.ST_PresetCameraType.isometricOffAxis4Right;
				break;
			case "isometricRightDown":
				nVal = AscCommon.ST_PresetCameraType.isometricRightDown;
				break;
			case "isometricRightUp":
				nVal = AscCommon.ST_PresetCameraType.isometricRightUp;
				break;
			case "isometricTopDown":
				nVal = AscCommon.ST_PresetCameraType.isometricTopDown;
				break;
			case "isometricTopUp":
				nVal = AscCommon.ST_PresetCameraType.isometricTopUp;
				break;
			case "legacyObliqueBottom":
				nVal = AscCommon.ST_PresetCameraType.legacyObliqueBottom;
				break;
			case "legacyObliqueBottomLeft":
				nVal = AscCommon.ST_PresetCameraType.legacyObliqueBottomLeft;
				break;
			case "legacyObliqueBottomRight":
				nVal = AscCommon.ST_PresetCameraType.legacyObliqueBottomRight;
				break;
			case "legacyObliqueFront":
				nVal = AscCommon.ST_PresetCameraType.legacyObliqueFront;
				break;
			case "legacyObliqueLeft":
				nVal = AscCommon.ST_PresetCameraType.legacyObliqueLeft;
				break;
			case "legacyObliqueRight":
				nVal = AscCommon.ST_PresetCameraType.legacyObliqueRight;
				break;
			case "legacyObliqueTop":
				nVal = AscCommon.ST_PresetCameraType.legacyObliqueTop;
				break;
			case "legacyObliqueTopLeft":
				nVal = AscCommon.ST_PresetCameraType.legacyObliqueTopLeft;
				break;
			case "legacyObliqueTopRight":
				nVal = AscCommon.ST_PresetCameraType.legacyObliqueTopRight;
				break;
			case "legacyPerspectiveBottom":
				nVal = AscCommon.ST_PresetCameraType.legacyPerspectiveBottom;
				break;
			case "legacyPerspectiveBottomLeft":
				nVal = AscCommon.ST_PresetCameraType.legacyPerspectiveBottomLeft;
				break;
			case "legacyPerspectiveBottomRight":
				nVal = AscCommon.ST_PresetCameraType.legacyPerspectiveBottomRight;
				break;
			case "legacyPerspectiveFront":
				nVal = AscCommon.ST_PresetCameraType.legacyPerspectiveFront;
				break;
			case "legacyPerspectiveLeft":
				nVal = AscCommon.ST_PresetCameraType.legacyPerspectiveLeft;
				break;
			case "legacyPerspectiveRight":
				nVal = AscCommon.ST_PresetCameraType.legacyPerspectiveRight;
				break;
			case "legacyPerspectiveTop":
				nVal = AscCommon.ST_PresetCameraType.legacyPerspectiveTop;
				break;
			case "legacyPerspectiveTopLeft":
				nVal = AscCommon.ST_PresetCameraType.legacyPerspectiveTopLeft;
				break;
			case "legacyPerspectiveTopRight":
				nVal = AscCommon.ST_PresetCameraType.legacyPerspectiveTopRight;
				break;
			case "obliqueBottom":
				nVal = AscCommon.ST_PresetCameraType.obliqueBottom;
				break;
			case "obliqueBottomLeft":
				nVal = AscCommon.ST_PresetCameraType.obliqueBottomLeft;
				break;
			case "obliqueBottomRight":
				nVal = AscCommon.ST_PresetCameraType.obliqueBottomRight;
				break;
			case "obliqueLeft":
				nVal = AscCommon.ST_PresetCameraType.obliqueLeft;
				break;
			case "obliqueRight":
				nVal = AscCommon.ST_PresetCameraType.obliqueRight;
				break;
			case "obliqueTop":
				nVal = AscCommon.ST_PresetCameraType.obliqueTop;
				break;
			case "obliqueTopLeft":
				nVal = AscCommon.ST_PresetCameraType.obliqueTopLeft;
				break;
			case "obliqueTopRight":
				nVal = AscCommon.ST_PresetCameraType.obliqueTopRight;
				break;
			case "orthographicFront":
				nVal = AscCommon.ST_PresetCameraType.orthographicFront;
				break;
			case "perspectiveAbove":
				nVal = AscCommon.ST_PresetCameraType.perspectiveAbove;
				break;
			case "perspectiveAboveLeftFacing":
				nVal = AscCommon.ST_PresetCameraType.perspectiveAboveLeftFacing;
				break;
			case "perspectiveAboveRightFacing":
				nVal = AscCommon.ST_PresetCameraType.perspectiveAboveRightFacing;
				break;
			case "perspectiveBelow":
				nVal = AscCommon.ST_PresetCameraType.perspectiveBelow;
				break;
			case "perspectiveContrastingLeftFacing":
				nVal = AscCommon.ST_PresetCameraType.perspectiveContrastingLeftFacing;
				break;
			case "perspectiveContrastingRightFacing":
				nVal = AscCommon.ST_PresetCameraType.perspectiveContrastingRightFacing;
				break;
			case "perspectiveFront":
				nVal = AscCommon.ST_PresetCameraType.perspectiveFront;
				break;
			case "perspectiveHeroicExtremeLeftFacing":
				nVal = AscCommon.ST_PresetCameraType.perspectiveHeroicExtremeLeftFacing;
				break;
			case "perspectiveHeroicExtremeRightFacing":
				nVal = AscCommon.ST_PresetCameraType.perspectiveHeroicExtremeRightFacing;
				break;
			case "perspectiveHeroicLeftFacing":
				nVal = AscCommon.ST_PresetCameraType.perspectiveHeroicLeftFacing;
				break;
			case "perspectiveHeroicRightFacing":
				nVal = AscCommon.ST_PresetCameraType.perspectiveHeroicRightFacing;
				break;
			case "perspectiveLeft":
				nVal = AscCommon.ST_PresetCameraType.perspectiveLeft;
				break;
			case "perspectiveRelaxed":
				nVal = AscCommon.ST_PresetCameraType.perspectiveRelaxed;
				break;
			case "perspectiveRelaxedModerately":
				nVal = AscCommon.ST_PresetCameraType.perspectiveRelaxedModerately;
				break;
			case "perspectiveRight":
				nVal = AscCommon.ST_PresetCameraType.perspectiveRight;
				break;
		}

		return nVal;
	}
	function To_XML_ST_LightRigDirection(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case AscCommon.ST_LightRigDirection.b:
				sVal = "b";
				break;
			case AscCommon.ST_LightRigDirection.bl:
				sVal = "bl";
				break;
			case AscCommon.ST_LightRigDirection.br:
				sVal = "br";
				break;
			case AscCommon.ST_LightRigDirection.l:
				sVal = "l";
				break;
			case AscCommon.ST_LightRigDirection.r:
				sVal = "r";
				break;
			case AscCommon.ST_LightRigDirection.t:
				sVal = "t";
				break;
			case AscCommon.ST_LightRigDirection.tl:
				sVal = "tl";
				break;
			case AscCommon.ST_LightRigDirection.tr:
				sVal = "tr";
				break;
		}

		return sVal;
	}
	function From_XML_ST_LightRigDirection(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "b":
				nVal = AscCommon.ST_LightRigDirection.b;
				break;
			case "bl":
				nVal = AscCommon.ST_LightRigDirection.bl;
				break;
			case "br":
				nVal = AscCommon.ST_LightRigDirection.br;
				break;
			case "l":
				nVal = AscCommon.ST_LightRigDirection.l;
				break;
			case "r":
				nVal = AscCommon.ST_LightRigDirection.r;
				break;
			case "t":
				nVal = AscCommon.ST_LightRigDirection.t;
				break;
			case "tl":
				nVal = AscCommon.ST_LightRigDirection.tl;
				break;
			case "tr":
				nVal = AscCommon.ST_LightRigDirection.tr;
				break;
		}

		return nVal;
	}
	function To_XML_ST_LightRigType(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case AscCommon.ST_LightRigType.balanced:
				sVal = "balanced";
				break;
			case AscCommon.ST_LightRigType.brightRoom:
				sVal = "brightRoom";
				break;
			case AscCommon.ST_LightRigType.chilly:
				sVal = "chilly";
				break;
			case AscCommon.ST_LightRigType.contrasting:
				sVal = "contrasting";
				break;
			case AscCommon.ST_LightRigType.flat:
				sVal = "flat";
				break;
			case AscCommon.ST_LightRigType.flood:
				sVal = "flood";
				break;
			case AscCommon.ST_LightRigType.freezing:
				sVal = "freezing";
				break;
			case AscCommon.ST_LightRigType.glow:
				sVal = "glow";
				break;
			case AscCommon.ST_LightRigType.harsh:
				sVal = "harsh";
				break;
			case AscCommon.ST_LightRigType.legacyFlat1:
				sVal = "legacyFlat1";
				break;
			case AscCommon.ST_LightRigType.legacyFlat2:
				sVal = "legacyFlat2";
				break;
			case AscCommon.ST_LightRigType.legacyFlat3:
				sVal = "legacyFlat3";
				break;
			case AscCommon.ST_LightRigType.legacyFlat4:
				sVal = "legacyFlat4";
				break;
			case AscCommon.ST_LightRigType.legacyHarsh1:
				sVal = "legacyHarsh1";
				break;
			case AscCommon.ST_LightRigType.legacyHarsh2:
				sVal = "legacyHarsh2";
				break;
			case AscCommon.ST_LightRigType.legacyHarsh3:
				sVal = "legacyHarsh3";
				break;
			case AscCommon.ST_LightRigType.legacyHarsh4:
				sVal = "legacyHarsh4";
				break;
			case AscCommon.ST_LightRigType.legacyNormal1:
				sVal = "legacyNormal1";
				break;
			case AscCommon.ST_LightRigType.legacyNormal2:
				sVal = "legacyNormal2";
				break;
			case AscCommon.ST_LightRigType.legacyNormal3:
				sVal = "legacyNormal3";
				break;
			case AscCommon.ST_LightRigType.legacyNormal4:
				sVal = "legacyNormal4";
				break;
			case AscCommon.ST_LightRigType.morning:
				sVal = "morning";
				break;
			case AscCommon.ST_LightRigType.soft:
				sVal = "soft";
				break;
			case AscCommon.ST_LightRigType.sunrise:
				sVal = "sunrise";
				break;
			case AscCommon.ST_LightRigType.sunset:
				sVal = "sunset";
				break;
			case AscCommon.ST_LightRigType.threePt:
				sVal = "threePt";
				break;
			case AscCommon.ST_LightRigType.twoPt:
				sVal = "twoPt";
				break;
		}

		return sVal;
	}
	function From_XML_ST_LightRigType(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "balanced":
				nVal = AscCommon.ST_LightRigType.balanced;
				break;
			case "brightRoom":
				nVal = AscCommon.ST_LightRigType.brightRoom;
				break;
			case "chilly":
				nVal = AscCommon.ST_LightRigType.chilly;
				break;
			case "contrasting":
				nVal = AscCommon.ST_LightRigType.contrasting;
				break;
			case "flat":
				nVal = AscCommon.ST_LightRigType.flat;
				break;
			case "flood":
				nVal = AscCommon.ST_LightRigType.flood;
				break;
			case "freezing":
				nVal = AscCommon.ST_LightRigType.freezing;
				break;
			case "glow":
				nVal = AscCommon.ST_LightRigType.glow;
				break;
			case "harsh":
				nVal = AscCommon.ST_LightRigType.harsh;
				break;
			case "legacyFlat1":
				nVal = AscCommon.ST_LightRigType.legacyFlat1;
				break;
			case "legacyFlat2":
				nVal = AscCommon.ST_LightRigType.legacyFlat2;
				break;
			case "legacyFlat3":
				nVal = AscCommon.ST_LightRigType.legacyFlat3;
				break;
			case "legacyFlat4":
				nVal = AscCommon.ST_LightRigType.legacyFlat4;
				break;
			case "legacyHarsh1":
				nVal = AscCommon.ST_LightRigType.legacyHarsh1;
				break;
			case "legacyHarsh2":
				nVal = AscCommon.ST_LightRigType.legacyHarsh2;
				break;
			case "legacyHarsh3":
				nVal = AscCommon.ST_LightRigType.legacyHarsh3;
				break;
			case "legacyHarsh4":
				nVal = AscCommon.ST_LightRigType.legacyHarsh4;
				break;
			case "legacyNormal1":
				nVal = AscCommon.ST_LightRigType.legacyNormal1;
				break;
			case "legacyNormal2":
				nVal = AscCommon.ST_LightRigType.legacyNormal2;
				break;
			case "legacyNormal3":
				nVal = AscCommon.ST_LightRigType.legacyNormal3;
				break;
			case "legacyNormal4":
				nVal = AscCommon.ST_LightRigType.legacyNormal4;
				break;
			case "morning":
				nVal = AscCommon.ST_LightRigType.morning;
				break;
			case "soft":
				nVal = AscCommon.ST_LightRigType.soft;
				break;
			case "sunrise":
				nVal = AscCommon.ST_LightRigType.sunrise;
				break;
			case "sunset":
				nVal = AscCommon.ST_LightRigType.sunset;
				break;
			case "threePt":
				nVal = AscCommon.ST_LightRigType.threePt;
				break;
			case "twoPt":
				nVal = AscCommon.ST_LightRigType.twoPt;
				break;
		}

		return nVal;
	}
	function To_XML_ST_BevelPresetType(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case AscCommon.ST_BevelPresetType.angle:
				sVal = "angle";
				break;
			case AscCommon.ST_BevelPresetType.artDeco:
				sVal = "artDeco";
				break;
			case AscCommon.ST_BevelPresetType.circle:
				sVal = "circle";
				break;
			case AscCommon.ST_BevelPresetType.convex:
				sVal = "convex";
				break;
			case AscCommon.ST_BevelPresetType.coolSlant:
				sVal = "coolSlant";
				break;
			case AscCommon.ST_BevelPresetType.cross:
				sVal = "cross";
				break;
			case AscCommon.ST_BevelPresetType.divot:
				sVal = "divot";
				break;
			case AscCommon.ST_BevelPresetType.hardEdge:
				sVal = "hardEdge";
				break;
			case AscCommon.ST_BevelPresetType.relaxedInset:
				sVal = "relaxedInset";
				break;
			case AscCommon.ST_BevelPresetType.riblet:
				sVal = "riblet";
				break;
			case AscCommon.ST_BevelPresetType.slope:
				sVal = "slope";
				break;
			case AscCommon.ST_BevelPresetType.softRound:
				sVal = "softRound";
				break;
		}

		return sVal;
	}
	function From_XML_ST_BevelPresetType(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "angle":
				nVal = AscCommon.ST_BevelPresetType.angle;
				break;
			case "artDeco":
				nVal = AscCommon.ST_BevelPresetType.artDeco;
				break;
			case "circle":
				nVal = AscCommon.ST_BevelPresetType.circle;
				break;
			case "convex":
				nVal = AscCommon.ST_BevelPresetType.convex;
				break;
			case "coolSlant":
				nVal = AscCommon.ST_BevelPresetType.coolSlant;
				break;
			case "cross":
				nVal = AscCommon.ST_BevelPresetType.cross;
				break;
			case "divot":
				nVal = AscCommon.ST_BevelPresetType.divot;
				break;
			case "hardEdge":
				nVal = AscCommon.ST_BevelPresetType.hardEdge;
				break;
			case "relaxedInset":
				nVal = AscCommon.ST_BevelPresetType.relaxedInset;
				break;
			case "riblet":
				nVal = AscCommon.ST_BevelPresetType.riblet;
				break;
			case "slope":
				nVal = AscCommon.ST_BevelPresetType.slope;
				break;
			case "softRound":
				nVal = AscCommon.ST_BevelPresetType.softRound;
				break;
		}

		return nVal;
	}
	function To_XML_ST_PresetMaterialType(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case AscCommon.ST_PresetMaterialType.clear:
				sVal = "clear";
				break;
			case AscCommon.ST_PresetMaterialType.dkEdge:
				sVal = "dkEdge";
				break;
			case AscCommon.ST_PresetMaterialType.flat:
				sVal = "flat";
				break;
			case AscCommon.ST_PresetMaterialType.legacyMatte:
				sVal = "legacyMatte";
				break;
			case AscCommon.ST_PresetMaterialType.legacyMetal:
				sVal = "legacyMetal";
				break;
			case AscCommon.ST_PresetMaterialType.legacyPlastic:
				sVal = "legacyPlastic";
				break;
			case AscCommon.ST_PresetMaterialType.legacyWireframe:
				sVal = "legacyWireframe";
				break;
			case AscCommon.ST_PresetMaterialType.matte:
				sVal = "matte";
				break;
			case AscCommon.ST_PresetMaterialType.metal:
				sVal = "metal";
				break;
			case AscCommon.ST_PresetMaterialType.plastic:
				sVal = "plastic";
				break;
			case AscCommon.ST_PresetMaterialType.powder:
				sVal = "powder";
				break;
			case AscCommon.ST_PresetMaterialType.softEdge:
				sVal = "softEdge";
				break;
			case AscCommon.ST_PresetMaterialType.softmetal:
				sVal = "softmetal";
				break;
			case AscCommon.ST_PresetMaterialType.translucentPowder:
				sVal = "translucentPowder";
				break;
			case AscCommon.ST_PresetMaterialType.warmMatte:
				sVal = "warmMatte";
				break;
		}

		return sVal;
	}
	function From_XML_ST_PresetMaterialType(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "clear":
				nVal = AscCommon.ST_PresetMaterialType.clear;
				break;
			case "dkEdge":
				nVal = AscCommon.ST_PresetMaterialType.dkEdge;
				break;
			case "flat":
				nVal = AscCommon.ST_PresetMaterialType.flat;
				break;
			case "legacyMatte":
				nVal = AscCommon.ST_PresetMaterialType.legacyMatte;
				break;
			case "legacyMetal":
				nVal = AscCommon.ST_PresetMaterialType.legacyMetal;
				break;
			case "legacyPlastic":
				nVal = AscCommon.ST_PresetMaterialType.legacyPlastic;
				break;
			case "legacyWireframe":
				nVal = AscCommon.ST_PresetMaterialType.legacyWireframe;
				break;
			case "matte":
				nVal = AscCommon.ST_PresetMaterialType.matte;
				break;
			case "metal":
				nVal = AscCommon.ST_PresetMaterialType.metal;
				break;
			case "plastic":
				nVal = AscCommon.ST_PresetMaterialType.plastic;
				break;
			case "powder":
				nVal = AscCommon.ST_PresetMaterialType.powder;
				break;
			case "softEdge":
				nVal = AscCommon.ST_PresetMaterialType.softEdge;
				break;
			case "softmetal":
				nVal = AscCommon.ST_PresetMaterialType.softmetal;
				break;
			case "translucentPowder":
				nVal = AscCommon.ST_PresetMaterialType.translucentPowder;
				break;
			case "warmMatte":
				nVal = AscCommon.ST_PresetMaterialType.warmMatte;
				break;
		}

		return nVal;
	}
	function To_XML_ST_CxnType(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case AscCommon.ST_CxnType.parOf:
				sVal = "parOf";
				break;
			case AscCommon.ST_CxnType.presOf:
				sVal = "presOf";
				break;
			case AscCommon.ST_CxnType.presParOf:
				sVal = "presParOf";
				break;
			case AscCommon.ST_CxnType.unknownRelationShip:
				sVal = "unknownRelationShip";
				break;
		}

		return sVal;
	}
	function From_XML_ST_CxnType(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "parOf":
				nVal = AscCommon.ST_CxnType.parOf;
				break;
			case "presOf":
				nVal = AscCommon.ST_CxnType.presOf;
				break;
			case "presParOf":
				nVal = AscCommon.ST_CxnType.presParOf;
				break;
			case "unknownRelationShip":
				nVal = AscCommon.ST_CxnType.unknownRelationShip;
				break;
		}

		return nVal;
	}
	function To_XML_OleObj_Type(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case AscCommon.c_oAscOleObjectTypes.document:
				sVal = "document";
				break;
			case AscCommon.c_oAscOleObjectTypes.spreadsheet:
				sVal = "spreadsheet";
				break;
			case AscCommon.c_oAscOleObjectTypes.formula:
				sVal = "formula";
				break;
		}

		return sVal;
	}
	function From_XML_OleObj_Type(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "document":
				nVal = AscCommon.c_oAscOleObjectTypes.document;
				break;
			case "spreadsheet":
				nVal = AscCommon.c_oAscOleObjectTypes.spreadsheet;
				break;
			case "formula":
				nVal = AscCommon.c_oAscOleObjectTypes.formula;
				break;
		}

		return nVal;
	}
	function To_XML_ST_BlendMode(nVal)
	{
		var sBlendType = undefined;
		switch (nVal)
		{
			case c_oAscBlendModeType.Darken:
				sBlendType = "darken";
				break;
			case c_oAscBlendModeType.Lighten:
				sBlendType = "lighten";
				break;
			case c_oAscBlendModeType.Mult:
				sBlendType = "mult";
				break;
			case c_oAscBlendModeType.Over:
				sBlendType = "over";
				break;
			case c_oAscBlendModeType.Screen:
				sBlendType = "screen";
				break;
		}

		return sBlendType;
	}
	function From_XML_ST_BlendMode(sVal)
	{
		var nBlendType = undefined;
		switch (sVal)
		{
			case "darken":
				nBlendType = c_oAscBlendModeType.Darken;
				break;
			case "lighten":
				nBlendType = c_oAscBlendModeType.Lighten;
				break;
			case "mult":
				nBlendType = c_oAscBlendModeType.Mult;
				break;
			case "over":
				nBlendType = c_oAscBlendModeType.Over;
				break;
			case "screen":
				nBlendType = c_oAscBlendModeType.Screen;
				break;
		}

		return nBlendType;
	}
	function To_XML_c_oAscNumberingFormat(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case Asc.c_oAscNumberingFormat.Aiueo:
				sVal = "aiueo";
				break;
			case Asc.c_oAscNumberingFormat.AiueoFullWidth:
				sVal = "aiueoFullWidth";
				break;
			case Asc.c_oAscNumberingFormat.ArabicAbjad:
				sVal = "arabicAbjad";
				break;
			case Asc.c_oAscNumberingFormat.ArabicAlpha:
				sVal = "arabicAlpha";
				break;
			case Asc.c_oAscNumberingFormat.BahtText:
				sVal = "bahtText";
				break;
			case Asc.c_oAscNumberingFormat.Bullet:
				sVal = "bullet";
				break;
			case Asc.c_oAscNumberingFormat.CardinalText:
				sVal = "cardinalText";
				break;
			case Asc.c_oAscNumberingFormat.Chicago:
				sVal = "chicago";
				break;
			case Asc.c_oAscNumberingFormat.ChineseCounting:
				sVal = "chineseCounting";
				break;
			case Asc.c_oAscNumberingFormat.ChineseCountingThousand:
				sVal = "chineseCountingThousand";
				break;
			case Asc.c_oAscNumberingFormat.ChineseLegalSimplified:
				sVal = "chineseLegalSimplified";
				break;
			case Asc.c_oAscNumberingFormat.Chosung:
				sVal = "chosung";
				break;
			case Asc.c_oAscNumberingFormat.Custom:
				sVal = "custom";
				break;
			case Asc.c_oAscNumberingFormat.Decimal:
				sVal = "decimal";
				break;
			case Asc.c_oAscNumberingFormat.DecimalEnclosedCircle:
				sVal = "decimalEnclosedCircle";
				break;
			case Asc.c_oAscNumberingFormat.DecimalEnclosedCircleChinese:
				sVal = "decimalEnclosedCircleChinese";
				break;
			case Asc.c_oAscNumberingFormat.DecimalEnclosedFullstop:
				sVal = "decimalEnclosedFullstop";
				break;
			case Asc.c_oAscNumberingFormat.DecimalEnclosedParen:
				sVal = "decimalEnclosedParen";
				break;
			case Asc.c_oAscNumberingFormat.DecimalFullWidth:
				sVal = "decimalFullWidth";
				break;
			case Asc.c_oAscNumberingFormat.DecimalFullWidth2:
				sVal = "decimalFullWidth2";
				break;
			case Asc.c_oAscNumberingFormat.DecimalHalfWidth:
				sVal = "decimalHalfWidth";
				break;
			case Asc.c_oAscNumberingFormat.DecimalZero:
				sVal = "decimalZero";
				break;
			case Asc.c_oAscNumberingFormat.DollarText:
				sVal = "dollarText";
				break;
			case Asc.c_oAscNumberingFormat.Ganada:
				sVal = "ganada";
				break;
			case Asc.c_oAscNumberingFormat.Hebrew1:
				sVal = "hebrew1";
				break;
			case Asc.c_oAscNumberingFormat.Hebrew2:
				sVal = "hebrew2";
				break;
			case Asc.c_oAscNumberingFormat.Hex:
				sVal = "hex";
				break;
			case Asc.c_oAscNumberingFormat.HindiConsonants:
				sVal = "hindiConsonants";
				break;
			case Asc.c_oAscNumberingFormat.HindiCounting:
				sVal = "hindiCounting";
				break;
			case Asc.c_oAscNumberingFormat.HindiNumbers:
				sVal = "hindiNumbers";
				break;
			case Asc.c_oAscNumberingFormat.HindiVowels:
				sVal = "hindiVowels";
				break;
			case Asc.c_oAscNumberingFormat.IdeographDigital:
				sVal = "ideographDigital";
				break;
			case Asc.c_oAscNumberingFormat.IdeographEnclosedCircle:
				sVal = "ideographEnclosedCircle";
				break;
			case Asc.c_oAscNumberingFormat.IdeographLegalTraditional:
				sVal = "ideographLegalTraditional";
				break;
			case Asc.c_oAscNumberingFormat.IdeographTraditional:
				sVal = "ideographTraditional";
				break;
			case Asc.c_oAscNumberingFormat.IdeographZodiac:
				sVal = "ideographZodiac";
				break;
			case Asc.c_oAscNumberingFormat.IdeographZodiacTraditional:
				sVal = "ideographZodiacTraditional";
				break;
			case Asc.c_oAscNumberingFormat.Iroha:
				sVal = "iroha";
				break;
			case Asc.c_oAscNumberingFormat.IrohaFullWidth:
				sVal = "irohaFullWidth";
				break;
			case Asc.c_oAscNumberingFormat.JapaneseCounting:
				sVal = "japaneseCounting";
				break;
			case Asc.c_oAscNumberingFormat.JapaneseDigitalTenThousand:
				sVal = "japaneseDigitalTenThousand";
				break;
			case Asc.c_oAscNumberingFormat.JapaneseLegal:
				sVal = "japaneseLegal";
				break;
			case Asc.c_oAscNumberingFormat.KoreanCounting:
				sVal = "koreanCounting";
				break;
			case Asc.c_oAscNumberingFormat.KoreanDigital:
				sVal = "koreanDigital";
				break;
			case Asc.c_oAscNumberingFormat.KoreanDigital2:
				sVal = "koreanDigital2";
				break;
			case Asc.c_oAscNumberingFormat.KoreanLegal:
				sVal = "koreanLegal";
				break;
			case Asc.c_oAscNumberingFormat.LowerLetter:
				sVal = "lowerLetter";
				break;
			case Asc.c_oAscNumberingFormat.LowerRoman:
				sVal = "lowerRoman";
				break;
			case Asc.c_oAscNumberingFormat.None:
				sVal = "none";
				break;
			case Asc.c_oAscNumberingFormat.NumberInDash:
				sVal = "numberInDash";
				break;
			case Asc.c_oAscNumberingFormat.Ordinal:
				sVal = "ordinal";
				break;
			case Asc.c_oAscNumberingFormat.OrdinalText:
				sVal = "ordinalText";
				break;
			case Asc.c_oAscNumberingFormat.RussianLower:
				sVal = "russianLower";
				break;
			case Asc.c_oAscNumberingFormat.RussianUpper:
				sVal = "russianUpper";
				break;
			case Asc.c_oAscNumberingFormat.TaiwaneseCounting:
				sVal = "taiwaneseCounting";
				break;
			case Asc.c_oAscNumberingFormat.TaiwaneseCountingThousand:
				sVal = "taiwaneseCountingThousand";
				break;
			case Asc.c_oAscNumberingFormat.TaiwaneseDigital:
				sVal = "taiwaneseDigital";
				break;
			case Asc.c_oAscNumberingFormat.ThaiCounting:
				sVal = "thaiCounting";
				break;
			case Asc.c_oAscNumberingFormat.ThaiLetters:
				sVal = "thaiLetters";
				break;
			case Asc.c_oAscNumberingFormat.ThaiNumbers:
				sVal = "thaiNumbers";
				break;
			case Asc.c_oAscNumberingFormat.UpperLetter:
				sVal = "upperLetter";
				break;
			case Asc.c_oAscNumberingFormat.UpperRoman:
				sVal = "upperRoman";
				break;
			case Asc.c_oAscNumberingFormat.VietnameseCounting:
				sVal = "vietnameseCounting";
				break;
		}

		return sVal;
	}
	function From_XML_c_oAscNumberingFormat(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "aiueo":
				nVal = Asc.c_oAscNumberingFormat.Aiueo;
				break;
			case "aiueoFullWidth":
				nVal = Asc.c_oAscNumberingFormat.AiueoFullWidth;
				break;
			case "arabicAbjad":
				nVal = Asc.c_oAscNumberingFormat.ArabicAbjad;
				break;
			case "arabicAlpha":
				nVal = Asc.c_oAscNumberingFormat.ArabicAlpha;
				break;
			case "bahtText":
				nVal = Asc.c_oAscNumberingFormat.BahtText;
				break;
			case "bullet":
				nVal = Asc.c_oAscNumberingFormat.Bullet;
				break;
			case "cardinalText":
				nVal = Asc.c_oAscNumberingFormat.CardinalText;
				break;
			case "chicago":
				nVal = Asc.c_oAscNumberingFormat.Chicago;
				break;
			case "chineseCounting":
				nVal = Asc.c_oAscNumberingFormat.ChineseCounting;
				break;
			case "chineseCountingThousand":
				nVal = Asc.c_oAscNumberingFormat.ChineseCountingThousand;
				break;
			case "chineseLegalSimplified":
				nVal = Asc.c_oAscNumberingFormat.ChineseLegalSimplified;
				break;
			case "chosung":
				nVal = Asc.c_oAscNumberingFormat.Chosung;
				break;
			case "custom":
				nVal = Asc.c_oAscNumberingFormat.Custom;
				break;
			case "decimal":
				nVal = Asc.c_oAscNumberingFormat.Decimal;
				break;
			case "decimalEnclosedCircle":
				nVal = Asc.c_oAscNumberingFormat.DecimalEnclosedCircle;
				break;
			case "decimalEnclosedCircleChinese":
				nVal = Asc.c_oAscNumberingFormat.DecimalEnclosedCircleChinese;
				break;
			case "decimalEnclosedFullstop":
				nVal = Asc.c_oAscNumberingFormat.DecimalEnclosedFullstop;
				break;
			case "decimalEnclosedParen":
				nVal = Asc.c_oAscNumberingFormat.DecimalEnclosedParen;
				break;
			case "decimalFullWidth":
				nVal = Asc.c_oAscNumberingFormat.DecimalFullWidth;
				break;
			case "decimalFullWidth2":
				nVal = Asc.c_oAscNumberingFormat.DecimalFullWidth2;
				break;
			case "decimalHalfWidth":
				nVal = Asc.c_oAscNumberingFormat.DecimalHalfWidth;
				break;
			case "decimalZero":
				nVal = Asc.c_oAscNumberingFormat.DecimalZero;
				break;
			case "dollarText":
				nVal = Asc.c_oAscNumberingFormat.DollarText;
				break;
			case "ganada":
				nVal = Asc.c_oAscNumberingFormat.Ganada;
				break;
			case "hebrew1":
				nVal = Asc.c_oAscNumberingFormat.Hebrew1;
				break;
			case "hebrew2":
				nVal = Asc.c_oAscNumberingFormat.Hebrew2;
				break;
			case "hex":
				nVal = Asc.c_oAscNumberingFormat.Hex;
				break;
			case "hindiConsonants":
				nVal = Asc.c_oAscNumberingFormat.HindiConsonants;
				break;
			case "hindiCounting":
				nVal = Asc.c_oAscNumberingFormat.HindiCounting;
				break;
			case "hindiNumbers":
				nVal = Asc.c_oAscNumberingFormat.HindiNumbers;
				break;
			case "hindiVowels":
				nVal = Asc.c_oAscNumberingFormat.HindiVowels;
				break;
			case "ideographDigital":
				nVal = Asc.c_oAscNumberingFormat.IdeographDigital;
				break;
			case "ideographEnclosedCircle":
				nVal = Asc.c_oAscNumberingFormat.IdeographEnclosedCircle;
				break;
			case "ideographLegalTraditional":
				nVal = Asc.c_oAscNumberingFormat.IdeographLegalTraditional;
				break;
			case "ideographTraditional":
				nVal = Asc.c_oAscNumberingFormat.IdeographTraditional;
				break;
			case "ideographZodiac":
				nVal = Asc.c_oAscNumberingFormat.IdeographZodiac;
				break;
			case "ideographZodiacTraditional":
				nVal = Asc.c_oAscNumberingFormat.IdeographZodiacTraditional;
				break;
			case "iroha":
				nVal = Asc.c_oAscNumberingFormat.Iroha;
				break;
			case "irohaFullWidth":
				nVal = Asc.c_oAscNumberingFormat.IrohaFullWidth;
				break;
			case "japaneseCounting":
				nVal = Asc.c_oAscNumberingFormat.JapaneseCounting;
				break;
			case "japaneseDigitalTenThousand":
				nVal = Asc.c_oAscNumberingFormat.JapaneseDigitalTenThousand;
				break;
			case "japaneseLegal":
				nVal = Asc.c_oAscNumberingFormat.JapaneseLegal;
				break;
			case "koreanCounting":
				nVal = Asc.c_oAscNumberingFormat.KoreanCounting;
				break;
			case "koreanDigital":
				nVal = Asc.c_oAscNumberingFormat.KoreanDigital;
				break;
			case "koreanDigital2":
				nVal = Asc.c_oAscNumberingFormat.KoreanDigital2;
				break;
			case "koreanLegal":
				nVal = Asc.c_oAscNumberingFormat.KoreanLegal;
				break;
			case "lowerLetter":
				nVal = Asc.c_oAscNumberingFormat.LowerLetter;
				break;
			case "lowerRoman":
				nVal = Asc.c_oAscNumberingFormat.LowerRoman;
				break;
			case "none":
				nVal = Asc.c_oAscNumberingFormat.None;
				break;
			case "numberInDash":
				nVal = Asc.c_oAscNumberingFormat.NumberInDash;
				break;
			case "ordinal":
				nVal = Asc.c_oAscNumberingFormat.Ordinal;
				break;
			case "ordinalText":
				nVal = Asc.c_oAscNumberingFormat.OrdinalText;
				break;
			case "russianLower":
				nVal = Asc.c_oAscNumberingFormat.RussianLower;
				break;
			case "russianUpper":
				nVal = Asc.c_oAscNumberingFormat.RussianUpper;
				break;
			case "taiwaneseCounting":
				nVal = Asc.c_oAscNumberingFormat.TaiwaneseCounting;
				break;
			case "taiwaneseCountingThousand":
				nVal = Asc.c_oAscNumberingFormat.TaiwaneseCountingThousand;
				break;
			case "taiwaneseDigital":
				nVal = Asc.c_oAscNumberingFormat.TaiwaneseDigital;
				break;
			case "thaiCounting":
				nVal = Asc.c_oAscNumberingFormat.ThaiCounting;
				break;
			case "thaiLetters":
				nVal = Asc.c_oAscNumberingFormat.ThaiLetters;
				break;
			case "thaiNumbers":
				nVal = Asc.c_oAscNumberingFormat.ThaiNumbers;
				break;
			case "upperLetter":
				nVal = Asc.c_oAscNumberingFormat.UpperLetter;
				break;
			case "upperRoman":
				nVal = Asc.c_oAscNumberingFormat.UpperRoman;
				break;
			case "vietnameseCounting":
				nVal = Asc.c_oAscNumberingFormat.VietnameseCounting;
				break;
		}

		return nVal;
	}
    //----------------------------------------------------------export----------------------------------------------------
    window['AscCommon']       = window['AscCommon'] || {};
    window['AscFormat']       = window['AscFormat'] || {};
	window['AscCommon'].WriterToJSON   = WriterToJSON;
	window['AscCommon'].ReaderFromJSON = ReaderFromJSON;
})(window);




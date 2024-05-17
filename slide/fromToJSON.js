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
(function(window, undefined){

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

	function private_GetLogicDocument()
	{
		if (editor)
			return editor.WordControl.m_oLogicDocument;
		
		return null;
	}
	function private_GetStyles()
	{
		var oLogicDocument = private_GetLogicDocument();

		return oLogicDocument instanceof AscCommonWord.CDocument ? oLogicDocument.Get_Styles() : oLogicDocument.globalTableStyles;
	}
	function private_GetPresentation(){
        return editor.WordControl.m_oLogicDocument;
    }

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// End of private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    var WriterToJSON   = window['AscJsonConverter'].WriterToJSON;
	var ReaderFromJSON = window['AscJsonConverter'].ReaderFromJSON;

	WriterToJSON.prototype.SerPresentation = function(oPres)
	{
		var aNotesMasters = [];
		for (var nNote = 0; nNote < oPres.notesMasters.length; nNote++)
			aNotesMasters.push(this.SerNotesMaster(oPres.notesMasters[nNote]));

		var aSlides = [];
		for (var nSlide = 0; nSlide < oPres.Slides.length; nSlide++)
			aSlides.push(this.SerSlide(oPres.Slides[nSlide], false, false, false));

		var aSldMasters = [];
		for (var nMaster = 0; nMaster < oPres.slideMasters.length; nMaster++)
			aSldMasters.push(this.SerMasterSlide(oPres.slideMasters[nMaster], true));

		var sConformanceType = oPres.pres.attrConformance === Asc.c_oAscConformanceType.Strict ? "strict" : "transitional";
		return {
			"slides":               aSlides,
			"sldSz":                this.SerSlideSize(oPres.sldSz),
			"showPr":               this.SerShowPr(oPres.showPr),
			"notesMasters":         aNotesMasters,
			"sldMasters":           aSldMasters,

			// CPres (oPresentation.pres)
			"defaultTextStyle":         this.SerLstStyle(oPres.pres.defaultTextStyle),
			"autoCompressPictures":     oPres.pres.attrAutoCompressPictures,
			"bookmarkIdSeed":           oPres.pres.attrBookmarkIdSeed,
			"compatMode":               oPres.pres.attrCompatMode,
			"conformance":              sConformanceType, 
			"embedTrueTypeFonts":       oPres.pres.attrEmbedTrueTypeFonts,
			"firstSlideNum":            oPres.pres.attrFirstSlideNum,
			"removePersonalInfoOnSave": oPres.pres.attrRemovePersonalInfoOnSave,
			"rtl":                      oPres.pres.attrRtl,
			"saveSubsetFonts":          oPres.pres.attrSaveSubsetFonts,
			"serverZoom":               oPres.pres.attrServerZoom,
			"showSpecialPlsOnTitleSld": oPres.pres.attrShowSpecialPlsOnTitleSld,
			"strictFirstAndLastChars":  oPres.pres.attrStrictFirstAndLastChars,
			"type":                     "presentation"
		}
	};
	WriterToJSON.prototype.SerSlideSize = function(oSldSz)
	{
		var sSldSzType = undefined;
		switch (oSldSz.type)
		{
			case Asc.c_oAscSlideSZType.Sz35mm:
				sSldSzType = "35mm";
				break;
			case Asc.c_oAscSlideSZType.SzA3:
				sSldSzType = "A3";
				break;
			case Asc.c_oAscSlideSZType.SzA4:
				sSldSzType = "A4";
				break;
			case Asc.c_oAscSlideSZType.SzB4ISO:
				sSldSzType = "B4ISO";
				break;
			case Asc.c_oAscSlideSZType.SzB4JIS:
				sSldSzType = "B4JIS";
				break;
			case Asc.c_oAscSlideSZType.SzB5ISO:
				sSldSzType = "B5ISO";
				break;
			case Asc.c_oAscSlideSZType.SzB5JIS:
				sSldSzType = "B5JIS";
				break;
			case Asc.c_oAscSlideSZType.SzBanner:
				sSldSzType = "banner";
				break;
			case Asc.c_oAscSlideSZType.SzCustom:
				sSldSzType = "custom";
				break;
			case Asc.c_oAscSlideSZType.SzHagakiCard:
				sSldSzType = "hagakiCard";
				break;
			case Asc.c_oAscSlideSZType.SzLedger:
				sSldSzType = "ledger";
				break;
			case Asc.c_oAscSlideSZType.SzLetter:
				sSldSzType = "letter";
				break;
			case Asc.c_oAscSlideSZType.SzOverhead:
				sSldSzType = "overhead";
				break;
			case Asc.c_oAscSlideSZType.SzScreen16x10:
				sSldSzType = "screen16x10";
				break;
			case Asc.c_oAscSlideSZType.SzScreen16x9:
				sSldSzType = "screen16x9";
				break;
			case Asc.c_oAscSlideSZType.SzScreen4x3:
				sSldSzType = "screen4x3";
				break;
			case Asc.c_oAscSlideSZType.SzWidescreen:
				sSldSzType = "wideScreen";
				break;
		}

		return {
			"cx":   oSldSz.cx,
			"cy":   oSldSz.cy,
			"type": sSldSzType
		}
	};
	WriterToJSON.prototype.SerShowPr = function(oShowPr)
	{
		if (!oShowPr)
			return undefined;

		return {
			"browse":  oShowPr.browse,
			"kiosk":   oShowPr.kiosk,
			"penClr":  this.SerColor(oShowPr.penClr),
			"present": oShowPr.present,
			"show":    oShowPr.show,
			"loop":    oShowPr.loop,
			"showAnimation": oShowPr.showAnimation,
			"showNarration": oShowPr.showNarration,
			"useTimings":    oShowPr.useTimings
		}
	};
	WriterToJSON.prototype.SerNotesMaster = function(oNoteMaster)
	{
		var oNotesMasterObj = {
			"id":         oNoteMaster.Id,
			"clrMap":     this.SerColorMapOvr(oNoteMaster.clrMap),
			"cSld":       this.SerCSld(oNoteMaster.cSld),
			"hf":         this.SerHF(oNoteMaster.hf),
			"notesStyle": this.SerLstStyle(oNoteMaster.txStyles),
			"theme":      oNoteMaster.Theme ? (this.themesMap[oNoteMaster.Theme.Id] ? oNoteMaster.Theme.Id : this.SerTheme(oNoteMaster.Theme)) : null
		}

		// мапим, чтобы не записывать несколько раз
		this.notesMasterMap[oNoteMaster.Id] = oNotesMasterObj;

		return oNotesMasterObj;
	};
	WriterToJSON.prototype.SerNotes = function(oNote)
	{
		return {
			"lock":             undefined, /// ??? вроде не нужно
			"clrMapOvr":        this.SerColorMapOvr(oNote.clrMap),
			"graphicObjects":   undefined, /// ??? вроде не нужно
			"cSld":             this.SerCSld(oNote.cSld),
			"showMasterPhAnim": oNote.showMasterPhAnim,
			"showMasterSp":     oNote.showMasterSp,
			"master":           this.notesMasterMap[oNote.Master.Id] ? oNote.Master.Id : this.SerNotesMaster(oNote.Master)
		}
	};
	WriterToJSON.prototype.SerSldComments = function(oSldComments)
	{
		var aComments = [];

		for (var nComment = 0; nComment < oSldComments.comments.length; nComment++)
			aComments.push(this.SerComment(oSldComments.comments[nComment]));

		return aComments;
	};
	WriterToJSON.prototype.SerComment = function(oComment)
	{
		return {
			"pos": {
				"x": oComment.x,
				"y": oComment.y
			},
			"text":       oComment.Data.m_sText,
			"authorName": oComment.Data.m_sUserName,
			"authorId":   oComment.Data.m_sUserId,
			"dt":         oComment.Data.m_sOOTime,
			"idx":        oComment.Parent.comments.indexOf(oComment)
		}
	};
	WriterToJSON.prototype.SerSlideLayout = function(oLayout, bWriteMaster)
	{
		var sLayoutType = undefined;
		switch (oLayout.type)
		{
			case c_oAscSlideLayoutType.Blank:
				sLayoutType = "blank";
				break;
			case c_oAscSlideLayoutType.Chart:
				sLayoutType = "chart";
				break;
			case c_oAscSlideLayoutType.ChartAndTx:
				sLayoutType = "chartAndTx";
				break;
			case c_oAscSlideLayoutType.ClipArtAndTx:
				sLayoutType = "clipArtAndTx";
				break;
			case c_oAscSlideLayoutType.ClipArtAndVertTx:
				sLayoutType = "clipArtAndVertTx";
				break;
			case c_oAscSlideLayoutType.Cust:
				sLayoutType = "cust";
				break;
			case c_oAscSlideLayoutType.Dgm:
				sLayoutType = "dgm";
				break;
			case c_oAscSlideLayoutType.FourObj:
				sLayoutType = "fourObj";
				break;
			case c_oAscSlideLayoutType.MediaAndTx:
				sLayoutType = "mediaAndTx";
				break;
			case c_oAscSlideLayoutType.Obj:
				sLayoutType = "obj";
				break;
			case c_oAscSlideLayoutType.ObjAndTwoObj:
				sLayoutType = "objAndTwoObj";
				break;
			case c_oAscSlideLayoutType.ObjAndTx:
				sLayoutType = "objAndTx";
				break;
			case c_oAscSlideLayoutType.ObjOnly:
				sLayoutType = "objOnly";
				break;
			case c_oAscSlideLayoutType.ObjOverTx:
				sLayoutType = "objOverTx";
				break;
			case c_oAscSlideLayoutType.ObjTx:
				sLayoutType = "objTx";
				break;
			case c_oAscSlideLayoutType.PicTx:
				sLayoutType = "picTx";
				break;
			case c_oAscSlideLayoutType.SecHead:
				sLayoutType = "secHead";
				break;
			case c_oAscSlideLayoutType.Tbl:
				sLayoutType = "tbl";
				break;
			case c_oAscSlideLayoutType.Title:
				sLayoutType = "title";
				break;
			case c_oAscSlideLayoutType.TitleOnly:
				sLayoutType = "titleOnly";
				break;
			case c_oAscSlideLayoutType.TwoColTx:
				sLayoutType = "twoColTx";
				break;
			case c_oAscSlideLayoutType.TwoObj:
				sLayoutType = "twoObj";
				break;
			case c_oAscSlideLayoutType.TwoObjAndObj:
				sLayoutType = "twoObjAndObj";
				break;
			case c_oAscSlideLayoutType.TwoObjAndTx:
				sLayoutType = "twoObjAndTx";
				break;
			case c_oAscSlideLayoutType.TwoObjOverTx:
				sLayoutType = "twoObjOverTx";
				break;
			case c_oAscSlideLayoutType.TwoTxTwoObj:
				sLayoutType = "twoTxTwoObj";
				break;
			case c_oAscSlideLayoutType.Tx:
				sLayoutType = "tx";
				break;
			case c_oAscSlideLayoutType.TxAndChart:
				sLayoutType = "txAndChart";
				break;
			case c_oAscSlideLayoutType.TxAndClipArt:
				sLayoutType = "txAndClipArt";
				break;
			case c_oAscSlideLayoutType.TxAndMedia:
				sLayoutType = "txAndMedia";
				break;
			case c_oAscSlideLayoutType.TxAndObj:
				sLayoutType = "txAndObj";
				break;
			case c_oAscSlideLayoutType.TxAndTwoObj:
				sLayoutType = "txAndTwoObj";
				break;
			case c_oAscSlideLayoutType.TxOverObj:
				sLayoutType = "txOverObj";
				break;
			case c_oAscSlideLayoutType.VertTitleAndTx:
				sLayoutType = "vertTitleAndTx";
				break;
			case c_oAscSlideLayoutType.VertTitleAndTxOverChart:
				sLayoutType = "vertTitleAndTxOverChart";
				break;
			case c_oAscSlideLayoutType.VertTx:
				sLayoutType = "vertTx";
				break;
		}

		return {
			"id":               oLayout.Id,
			"master":           bWriteMaster ? this.SerMasterSlide(oLayout.Master, false) : oLayout.Master.Id,
			"clrMapOvr":        this.SerColorMapOvr(oLayout.clrMap),
			"cSld":             this.SerCSld(oLayout.cSld),
			"hf":               this.SerHF(oLayout.hf),
			"timing":           this.SerTiming(oLayout.timing),
			"transition":       this.SerTransition(oLayout.transition),
			"matchingName":     oLayout.matchingName,
			"preserve":         oLayout.preserve,
			"showMasterPhAnim": oLayout.showMasterPhAnim,
			"showMasterSp":     oLayout.showMasterSp,
			"userDrawn":        oLayout.userDrawn,
			"ltType":           sLayoutType,
			"imgBase64":        oLayout.ImageBase64,
			"type":             "sldLayout"
		}
	};
	WriterToJSON.prototype.SerMasterSlide = function(oMaster, bWriteAllMasLayouts)
	{
		var aLayoutLst = [];
		if (bWriteAllMasLayouts)
		{
			for (var nLayout = 0; nLayout < oMaster.sldLayoutLst.length; nLayout++)
				aLayoutLst.push(this.SerSlideLayout(oMaster.sldLayoutLst[nLayout], false));
		}

		return {
			"id":               oMaster.Id,
			"theme":            this.SerTheme(oMaster.Theme),
			"clrMapOvr":        this.SerColorMapOvr(oMaster.clrMap),
			"cSld":             this.SerCSld(oMaster.cSld),
			"hf":               this.SerHF(oMaster.hf),
			"sldLayoutLst":     aLayoutLst,
			"timing":           this.SerTiming(oMaster.timing),
			"transition":       this.SerTransition(oMaster.transition),
			"txStyles":         this.SerTxStyles(oMaster.txStyles),
			"preserve":         oMaster.preserve,
			"imgBase64":        oMaster.ImageBase64,
			"type":             "sldMaster"
		}
	};
	WriterToJSON.prototype.SerSlide = function(oSlide, bWriteLayout, bWriteMaster, bWriteAllMasLayouts)
	{
		var oMaster = oSlide.Layout.Master.Id;
		var oLayout = oSlide.Layout.Id;

		// нет смысла тащить master за слайдом без layout
		if (bWriteLayout || (bWriteMaster && bWriteAllMasLayouts))
		{
			if (bWriteMaster)
				oMaster = this.SerMasterSlide(oSlide.Layout.Master, bWriteAllMasLayouts);

			if (!bWriteAllMasLayouts)
				oLayout = this.SerSlideLayout(oSlide.Layout, false);
		}

		return {
			"notes":            this.SerNotes(oSlide.notes),
			"master":           oMaster,
			"clrMapOvr":        this.SerColorMapOvr(oSlide.clrMap),
			"layout":           oLayout,
			"cSld":             this.SerCSld(oSlide.cSld),
			"transition":       this.SerTransition(oSlide.transition),
			"timing":           this.SerTiming(oSlide.timing),
			"comments":         this.SerSldComments(oSlide.slideComments),
			"show":             oSlide.show,
			"showMasterPhAnim": oSlide.showMasterPhAnim,
			"showMasterSp":     oSlide.showMasterSp,
			"type":             "slide"
		}
	};
	WriterToJSON.prototype.SerSlides = function(nStart, nEnd, bWriteLayout, bWriteMaster, bWriteAllMasLayouts)
	{
		var aMasters = [];
		var aLayouts = [];
		var aSlides  = [];

		var oPresentation = private_GetPresentation();
		for (var nSlide = nStart; nSlide <= nEnd; nSlide++)
		{
			aSlides.push(this.SerSlide(oPresentation.Slides[nSlide], false, false, false));

			if (bWriteLayout || (bWriteMaster && bWriteAllMasLayouts))
			{
				if (bWriteMaster)
				{
					var sMasterId = oPresentation.Slides[nSlide].Layout.Master ? oPresentation.Slides[nSlide].Layout.Master.Id : "";
					if (sMasterId && !this.mastersMap[sMasterId])
					{
						this.mastersMap[sMasterId] = this.SerMasterSlide(oPresentation.Slides[nSlide].Layout.Master, bWriteAllMasLayouts);
						aMasters.push(this.mastersMap[sMasterId]);
					}
				}

				if (!bWriteAllMasLayouts)
				{
					var sLayoutId = oPresentation.Slides[nSlide].Layout ? oPresentation.Slides[nSlide].Layout.Id : "";
					if (!this.layoutsMap[sLayoutId])
					{
						this.layoutsMap[sLayoutId] = this.SerSlideLayout(oPresentation.Slides[nSlide].Layout);
						aLayouts.push(this.layoutsMap[sLayoutId]);
					}
				}
			}
		}
			
		return {
			"masters": aMasters,
			"layouts": aLayouts,
			"slides":  aSlides,
			"type":    "slides"
		}
	};
	WriterToJSON.prototype.SerTxStyles = function(oTxStyles)
	{
		return {
			"bodyStyle":  this.SerLstStyle(oTxStyles.bodyStyle),
			"otherStyle": this.SerLstStyle(oTxStyles.otherStyle),
			"titleStyle": this.SerLstStyle(oTxStyles.titleStyle)
		}
	};
	WriterToJSON.prototype.SerTableStylesForWrite = function()
	{
		if (this.stylesForWrite == null)
			return undefined;

		var oResult = {};
		for (var key in this.stylesForWrite.styles)
			oResult[key] = this.SerTableStyle(this.stylesForWrite.styles[key]);

		return oResult;
	};
	WriterToJSON.prototype.SerTableStyle = function(oStyle)
	{
		return {
			"styleId":  oStyle.Id,
			"name":     oStyle.Name,
			"tblBg":    this.SerTableCellStyle(oStyle.TablePr, true),
			"wholeTbl": this.SerTableStylePart(oStyle.TableWholeTable),
			"band1H":   this.SerTableStylePart(oStyle.TableBand1Horz),
			"band2H":   this.SerTableStylePart(oStyle.TableBand2Horz),
			"band1V":   this.SerTableStylePart(oStyle.TableBand1Vert),
			"band2V":   this.SerTableStylePart(oStyle.TableBand2Vert),
			"lastCol":  this.SerTableStylePart(oStyle.TableLastCol),
			"firstCol": this.SerTableStylePart(oStyle.TableFirstCol),
			"lastRow":  this.SerTableStylePart(oStyle.TableLastRow),
			"seCell":   this.SerTableStylePart(oStyle.TableBRCell),
			"swCell":   this.SerTableStylePart(oStyle.TableBLCell),
			"firstRow": this.SerTableStylePart(oStyle.TableFirstRow),
			"neCell":   this.SerTableStylePart(oStyle.TableTRCell),
			"nwCell":   this.SerTableStylePart(oStyle.TableTLCell)
		}
	};
	WriterToJSON.prototype.SerTableCellStyle = function(oCellPr, isTableBorders)
	{
		if (!oCellPr)
			return undefined;

		let oBorders = isTableBorders ? oCellPr.TableBorders : oCellPr.TableCellBorders;
		let oShd = oCellPr.Shd;
		let bBorders = oBorders && (oBorders.Left || oBorders.Right || oBorders.Top ||  oBorders.Bottom ||  oBorders.InsideH || oBorders.InsideV);
		
		return {
			"tcBdr": bBorders ? {
				"left":    this.SerTableCellBorder(oBorders.Left),
				"right":   this.SerTableCellBorder(oBorders.Right),
				"top":     this.SerTableCellBorder(oBorders.Top),
				"bottom":  this.SerTableCellBorder(oBorders.Bottom),
				"insideH": this.SerTableCellBorder(oBorders.InsideH),
				"insideV": this.SerTableCellBorder(oBorders.InsideV),
			} : undefined,
			"fillRef":     oShd ? this.SerStyleRef(oShd.FillRef) : undefined,
			"fill":        oShd ? this.SerFill(oShd.Unifill) : undefined
		}
	};
	WriterToJSON.prototype.SerTableCellBorder = function(oBorder)
	{
		if (!oBorder)
			return undefined;

		let oResult = {};

		if (oBorder.LineRef)
			oResult["lnRef"] = this.SerStyleRef(oBorder.LineRef);
		else
		{
			let oLn = new AscFormat.CLn();
			oLn.fromDocumentBorder(oBorder);
			oResult["ln"] = this.SerLn(oLn);
		}

		return oResult;
	};
	WriterToJSON.prototype.SerTableStylePart = function(oStylePart)
	{
		if (!oStylePart)
			return undefined;

		return {
			"tcTxStyle": this.SerTcTxStyle(oStylePart.TextPr),
			"tcStyle":   this.SerTableCellStyle(oStylePart.TableCellPr)
		}
	};
	WriterToJSON.prototype.SerTcTxStyle = function(oTextPr)
	{
		if (!oTextPr)
			return undefined;

		let oResult;
		let oUnicolor = oTextPr.Unifill && oTextPr.Unifill.fill && oTextPr.Unifill.fill.color;
		if(oTextPr.Italic === true || oTextPr.Italic === false ||
			oTextPr.Bold === true || oTextPr.Bold === false ||
			oTextPr.FontRef || oUnicolor)
			{
				oResult = {};
				if(oTextPr.Italic === true || oTextPr.Italic === false) {
					oResult["i"] = oTextPr.Italic ? "on" : "off";
				}
				if(oTextPr.Bold === true || oTextPr.Bold === false) {
					oResult["b"] = oTextPr.Bold ? "on" : "off";
				}
	
				if(oTextPr.FontRef) {
					oResult["fontRef"] = this.SerFontRef(oTextPr.FontRef);
				}
				if(oUnicolor) {
					oResult["color"] = this.SerColor(oUnicolor);
				}
			}

		return oResult;
	};
	WriterToJSON.prototype.SerHF = function(oHf)
	{
		if (!oHf)
			return undefined;

		return {
			"dt":     oHf.dt,
			"ftr":    oHf.ftr,
			"hdr":    oHf.hdr,
			"sldNum": oHf.sldNum
		}
	};
	WriterToJSON.prototype.SerCSld = function(oCSld)
	{
		var aSpTree = [];
		for (var nElm = 0; nElm < oCSld.spTree.length; nElm++)
			aSpTree.push(this.SerGraphicObject(oCSld.spTree[nElm]));

		return {
			"bg":     this.SerBg(oCSld.Bg),
			"spTree": aSpTree,
			"name":   oCSld.name
		}
	};
	WriterToJSON.prototype.SerBg = function(oBg)
	{
		if (!oBg)
			return undefined;

		var sBwModeType = undefined;
		switch (oBg.bwMode)
		{
			case c_oAscSlideBgBwModeType.Auto:
				sBwModeType = "auto";
				break;
			case c_oAscSlideBgBwModeType.Black:
				sBwModeType = "black";
				break;
			case c_oAscSlideBgBwModeType.BlackGray:
				sBwModeType = "blackGray";
				break;
			case c_oAscSlideBgBwModeType.BlackWhite:
				sBwModeType = "blackWhite";
				break;
			case c_oAscSlideBgBwModeType.Clr:
				sBwModeType = "clr";
				break;
			case c_oAscSlideBgBwModeType.Gray:
				sBwModeType = "gray";
				break;
			case c_oAscSlideBgBwModeType.GrayWhite:
				sBwModeType = "grayWhite";
				break;
			case c_oAscSlideBgBwModeType.Hidden:
				sBwModeType = "hidden";
				break;
			case c_oAscSlideBgBwModeType.InvGray:
				sBwModeType = "invGray";
				break;
			case c_oAscSlideBgBwModeType.LtGray:
				sBwModeType = "ltGray";
				break;
			case c_oAscSlideBgBwModeType.White:
				sBwModeType = "white";
				break;
		}

		return {
			"bwMode": sBwModeType,
			"bgPr":   this.SerBgPr(oBg.bgPr),
			"bgRef":	this.SerStyleRef(oBg.bgRef)
		}
	};
	WriterToJSON.prototype.SerBgPr = function(oBgPr)
	{
		if (!oBgPr)
			return undefined;

		return {
			"fill":         this.SerFill(oBgPr.Fill),
			"shadeToTitle": oBgPr.shadeToTitle
		}
	};
	WriterToJSON.prototype.SerTransition = function(oTransition)
	{
		if (!oTransition)
			return undefined;
		
		var sTransType = undefined;
		switch (oTransition.TransitionType)
		{
			case c_oAscSlideTransitionTypes.None:
				sTransType = "none";
				break;
			case c_oAscSlideTransitionTypes.Fade:
				sTransType = "fade";
				break;
			case c_oAscSlideTransitionTypes.Push:
				sTransType = "push";
				break;
			case c_oAscSlideTransitionTypes.Wipe:
				sTransType = "wipe";
				break;
			case c_oAscSlideTransitionTypes.Split:
				sTransType = "split";
				break;
			case c_oAscSlideTransitionTypes.UnCover:
				sTransType = "unCover";
				break;
			case c_oAscSlideTransitionTypes.Cover:
				sTransType = "cover";
				break;
			case c_oAscSlideTransitionTypes.Clock:
				sTransType = "clock";
				break;
			case c_oAscSlideTransitionTypes.Zoom:
				sTransType = "zoom";
				break;
		}

		var transOption = "none";
		switch (sTransType)
		{
			case "fade":
				transOption = oTransition.TransitionOption !== c_oAscSlideTransitionParams.Fade_Smoothly;
				break;
			case "push":
			case "wipe":
			case "cover":
			case "uncover":
				switch (oTransition.TransitionOption)
				{
					case c_oAscSlideTransitionParams.Param_Left:
						transOption = "l";
						break;
					case c_oAscSlideTransitionParams.Param_Top:
						transOption = "t";
						break;
					case c_oAscSlideTransitionParams.Param_Right:
						transOption = "r";
						break;
					case c_oAscSlideTransitionParams.Param_Bottom:
						transOption = "b";
						break;
					case c_oAscSlideTransitionParams.Param_TopLeft:
						transOption = "tl";
						break;
					case c_oAscSlideTransitionParams.Param_TopRight:
						transOption = "tr";
						break;
					case c_oAscSlideTransitionParams.Param_BottomLeft:
						transOption = "bl";
						break;
					case c_oAscSlideTransitionParams.Param_BottomRight:
						transOption = "br";
						break;
				}
				break;
			case "split":
				switch (oTransition.TransitionOption)
				{
					case c_oAscSlideTransitionParams.Split_VerticalIn:
						transOption = "verIn";
						break;
					case c_oAscSlideTransitionParams.Split_VerticalOut:
						transOption = "verOut";
						break;
					case c_oAscSlideTransitionParams.Split_HorizontalIn:
						transOption = "horIn";
						break;
					case c_oAscSlideTransitionParams.Split_HorizontalOut:
						transOption = "horOut";
						break;
				}
				break;
			case "clock":
				switch (oTransition.TransitionOption)
				{
					case c_oAscSlideTransitionParams.Clock_Clockwise:
						transOption = "clockwise";
						break;
					case c_oAscSlideTransitionParams.Clock_Counterclockwise:
						transOption = "counterClockwise";
						break;
					case c_oAscSlideTransitionParams.Clock_Wedge:
						transOption = "wedge";
						break;
				}
				break;
			case "zoom":
				switch (oTransition.TransitionOption)
				{
					case c_oAscSlideTransitionParams.Zoom_In:
						transOption = "in";
						break;
					case c_oAscSlideTransitionParams.Zoom_Out:
						transOption = "out";
						break;
					case c_oAscSlideTransitionParams.Zoom_AndRotate:
						transOption = "andRotate";
						break;
				}
				break;
		}

		return {
			"option":   transOption,
			"type":     sTransType,
			"transDur": oTransition.TransitionDuration,
			"advClick": oTransition.SlideAdvanceOnMouseClick,
			"advAfter": oTransition.SlideAdvanceAfter,
			"advDur":   oTransition.SlideAdvanceDuration,
			"shwLoop":  oTransition.ShowLoop
		}
	};
	WriterToJSON.prototype.SerTiming = function(oTiming)
	{
		if (!oTiming)
			return undefined;
		
		return {
			"bldLst": this.SerBldLst(oTiming.bldLst),
			"tnLst":  this.SerTnLst(oTiming.tnLst),
			"type":   "timing"
		}
	};
	WriterToJSON.prototype.SerBldLst = function(oBldLst)
	{
		if (!oBldLst)
			return undefined;

		var aBldLst = [];
		for (var nElm = 0; nElm < oBldLst.list.length; nElm++)
		{
			if (oBldLst.list[nElm] instanceof AscFormat.CBldDgm)
				aBldLst.push(this.SerBldDgm(oBldLst.list[nElm]));
			else if (oBldLst.list[nElm] instanceof AscFormat.CBldOleChart)
				aBldLst.push(this.SerBldOleChart(oBldLst.list[nElm]));
			else if (oBldLst.list[nElm] instanceof AscFormat.CBldGraphic)
				aBldLst.push(this.SerBldGraphic(oBldLst.list[nElm]));
			else if (oBldLst.list[nElm] instanceof AscFormat.CBldP)
				aBldLst.push(this.SerBldP(oBldLst.list[nElm]));
		}

		return aBldLst;
	};
	WriterToJSON.prototype.SerBldDgm = function(oBldDgm)
	{
		if (!oBldDgm)
			return undefined;

		let nSpId = this.GetSpIdxId(oBldDgm.spid);
		return {
			"bld":      oBldDgm.bld, // ?? c_oAscSlideDgmBuildType
			"grpId":    oBldDgm.grpId,
			"spid":     nSpId != null ? nSpId : undefined,
			"uiExpand": oBldDgm.uiExpand,
			"type":     "bldDgm"
		}
	};
	WriterToJSON.prototype.SerBldOleChart = function(oBldOleChart)
	{
		if (!oBldOleChart)
			return undefined;

		let nSpId = this.GetSpIdxId(oBldOleChart.spid);

		return {
			"animBg":   oBldOleChart.animBg,
			"bld":      oBldOleChart.bld, // ?? c_oAscSlideOleChartBuildType
			"grpId":    oBldOleChart.grpId,
			"spid":     nSpId != null ? nSpId : undefined,
			"uiExpand": oBldOleChart.uiExpand,
			"type":     "bldOleChart"
		}
	};
	WriterToJSON.prototype.SerBldGraphic = function(oBldGraphic)
	{
		if (!oBldGraphic)
			return undefined;

		let nSpId = this.GetSpIdxId(oBldGraphic.spid);

		return {
			"bldAsOne": this.SerEmptyObject(oBldGraphic.bldAsOne),
			"bldSub":   this.SerBldSub(oBldGraphic.bldSub),
			"grpId":    oBldGraphic.grpId,
			"spid":     nSpId != null ? nSpId : undefined,
			"uiExpand": oBldGraphic.uiExpand,
			"type":     "bldGraphic"
		}
	};
	WriterToJSON.prototype.SerEmptyObject = function(oEmptyObject)
	{
		if (!oEmptyObject)
			return undefined;

		return {
			"type": "emptyObj"
		}
	};
	WriterToJSON.prototype.SerBldSub = function(oBldSub)
	{
		if (!oBldSub)
			return undefined;

		return {
			"chart":    oBldSub.chart,
        	"animBg":   oBldSub.animBg,
        	"bldChart": oBldSub.bldChart, // ?? c_oAscSlideAnimDgmBuildType
        	"bldDgm":   oBldSub.bldDgm, // ?? c_oAscSlideAnimChartBuildType
        	"rev":      oBldSub.rev
		}	
	};
	WriterToJSON.prototype.SerBldP = function(oBldP)
	{
		if (!oBldP)
			return undefined;

		var sBuildType = undefined;
		switch (oBldP.build)
		{
			case c_oAscSlideParaBuildType.AllAtOnce:
				sBuildType = "allAtOnce";
				break;
			case c_oAscSlideParaBuildType.Cust:
				sBuildType = "cust";
				break;
			case c_oAscSlideParaBuildType.P:
				sBuildType = "p";
				break;
			case c_oAscSlideParaBuildType.Whole:
				sBuildType = "whole";
				break;
		}
		let nSpId = this.GetSpIdxId(oBldP.spid);

		return {
			"tmplLst":         this.SerTmplLst(oBldP.tmplLst),
			"advAuto":         oBldP.advAuto,
        	"animBg":          oBldP.animBg,
        	"autoUpdateAnimB": oBldP.autoUpdateAnimB,
        	"bldLvl":          oBldP.bldLvl,
        	"build":           sBuildType,
			"grpId":           oBldP.grpId,
			"rev":             oBldP.rev,
			"spid":            nSpId != null ? nSpId : undefined,
			"uiExpand":        oBldP.uiExpand,
			"type":            "bldP"
		}	
	};
	WriterToJSON.prototype.SerTmplLst = function(oTmplLst)
	{
		if (!oTmplLst)
			return undefined;
		
		var aTmplLst = [];

		for (var nElm = 0; nElm < oTmplLst.list.length; nElm++)
			aTmplLst.push(this.SerTmpl(oTmplLst.list[nElm]));
		
		return aTmplLst;
	};
	WriterToJSON.prototype.SerTmpl = function(oTmpl)
	{
		return {
			"lvl":   oTmpl.lvl,
			"tnLst": this.SerTnLst(oTmpl.tnLst)
		}
	};
	WriterToJSON.prototype.SerTnLst = function(oTnLst)
	{
		if (!oTnLst)
			return undefined;

		var aTnLst   = [];
		var oTempElm = null;

		for (var nElm = 0; nElm < oTnLst.list.length; nElm++)
		{
			oTempElm = oTnLst.list[nElm];
			if (oTempElm instanceof AscFormat.CPar)
				aTnLst.push(this.SerPar(oTempElm));
			else if (oTempElm instanceof AscFormat.CSeq)
				aTnLst.push(this.SerSeq(oTempElm));
			else if (oTempElm instanceof AscFormat.CAudio)
				aTnLst.push(this.SerAudio(oTempElm));
			else if (oTempElm instanceof AscFormat.CVideo)
				aTnLst.push(this.SerVideo(oTempElm));
			else if (oTempElm instanceof AscFormat.CExcl)
				aTnLst.push(this.SerExcl(oTempElm));
			else if (oTempElm instanceof AscFormat.CAnim)
				aTnLst.push(this.SerAnim(oTempElm));
			else if (oTempElm instanceof AscFormat.CAnimClr)
				aTnLst.push(this.SerAnimClr(oTempElm));
			else if (oTempElm instanceof AscFormat.CAnimEffect)
				aTnLst.push(this.SerAnimEffect(oTempElm));
			else if (oTempElm instanceof AscFormat.CAnimMotion)
				aTnLst.push(this.SerAnimMotion(oTempElm));
			else if (oTempElm instanceof AscFormat.CAnimRot)
				aTnLst.push(this.SerAnimRot(oTempElm));
			else if (oTempElm instanceof AscFormat.CAnimScale)
				aTnLst.push(this.SerAnimScale(oTempElm));
			else if (oTempElm instanceof AscFormat.CCmd)
				aTnLst.push(this.SerCmd(oTempElm));
			else if (oTempElm instanceof AscFormat.CSet)
				aTnLst.push(this.SerSet(oTempElm));
		}	

		return aTnLst;
	};
	WriterToJSON.prototype.SerPar = function(oPar)
	{
		return {
			"cTn":     this.SerCTn(oPar.cTn),
			"objType": "par"
		}
	};
	WriterToJSON.prototype.SerCTn = function(oCTn)
	{
		if (!oCTn)
			return undefined;

		var sNodeFillType = undefined;
		switch (oCTn.fill)
		{
			case c_oAscSlideNodeFillType.Freeze:
				sNodeFillType = "freeze";
				break;
			case c_oAscSlideNodeFillType.Hold:
				sNodeFillType = "hold";
				break;
			case c_oAscSlideNodeFillType.Remove:
				sNodeFillType = "remove";
				break;
			case c_oAscSlideNodeFillType.Transition:
				sNodeFillType = "transition";
				break;
		}

		var sMasterRelType = undefined;
		switch (oCTn.masterRel)
		{
			case c_oAscSlideMasterRelationType.LastClick:
				sMasterRelType = "lastClick";
				break;
			case c_oAscSlideMasterRelationType.NextClick:
				sMasterRelType = "nextClick";
				break;
			case c_oAscSlideMasterRelationType.SameClick:
				sMasterRelType = "sameClick";
				break;
		}

		var sNodeType = undefined;
		switch (oCTn.nodeType)
		{
			case c_oAscSlideNodeType.AfterEffect:
				sNodeType = "afterEffect";
				break;
			case c_oAscSlideNodeType.AfterGroup:
				sNodeType = "afterGroup";
				break;
			case c_oAscSlideNodeType.ClickEffect:
				sNodeType = "clickEffect";
				break;
			case c_oAscSlideNodeType.ClickPar:
				sNodeType = "clickPar";
				break;
			case c_oAscSlideNodeType.InteractiveSeq:
				sNodeType = "interactiveSeq";
				break;
			case c_oAscSlideNodeType.MainSeq:
				sNodeType = "mainSeq";
				break;
			case c_oAscSlideNodeType.TmRoot:
				sNodeType = "tmRoot";
				break;
			case c_oAscSlideNodeType.WithEffect:
				sNodeType = "withEffect";
				break;
			case c_oAscSlideNodeType.WithGroup:
				sNodeType = "withGroup";
				break;
		}

		var sPresetClassType = undefined;
		switch (oCTn.presetClass)
		{
			case c_oAscSlidePresetClassType.Emph:
				sPresetClassType = "emph";
				break;
			case c_oAscSlidePresetClassType.Entr:
				sPresetClassType = "entr";
				break;
			case c_oAscSlidePresetClassType.Exit:
				sPresetClassType = "exit";
				break;
			case c_oAscSlidePresetClassType.Mediacall:
				sPresetClassType = "mediacall";
				break;
			case c_oAscSlidePresetClassType.Path:
				sPresetClassType = "path";
				break;
			case c_oAscSlidePresetClassType.Verb:
				sPresetClassType = "verb";
				break;
		}

		var sRestartType = undefined;
		switch (oCTn.restart)
		{
			case c_oAscSlideRestartType.Always:
				sRestartType = "always";
				break;
			case c_oAscSlideRestartType.Never:
				sRestartType = "never";
				break;
			case c_oAscSlideRestartType.WhenNotActive:
				sRestartType = "whenNotActive";
				break;
		}

		var sSyncBehaviorType = undefined;
		switch (oCTn.syncBehavior)
		{
			case c_oAscSlideSyncBehaviorType.CanSlip:
				sSyncBehaviorType = "canSlip";
				break;
			case c_oAscSlideSyncBehaviorType.Locked:
				sSyncBehaviorType = "locked";
				break;
		}

		return {
			"Id":         oCTn.Id,

			"childTnLst": this.SerTnLst(oCTn.childTnLst),
			"endCondLst": this.SerCondLst(oCTn.endCondLst),
			"endSync":    this.SerCond(oCTn.endSync),
			"iterate":    this.SerIterate(oCTn.iterate),
			"stCondLst":  this.SerCondLst(oCTn.stCondLst),
			"subTnLst":   this.SerTnLst(oCTn.subTnLst),

			"accel":         oCTn.accel,
			"afterEffect":   oCTn.afterEffect,
			"autoRev":       oCTn.autoRev,
			"bldLvl":        oCTn.bldLvl,
			"decel":         oCTn.decel,
			"display":       oCTn.display,
			"dur":           oCTn.dur,
			"evtFilter":     oCTn.evtFilter,
			"fill":          sNodeFillType,
			"grpId":         oCTn.grpId,
			"id":            oCTn.id,
			"masterRel":     sMasterRelType,
			"nodePh":        oCTn.nodePh,
			"nodeType":      sNodeType,
			"presetClass":   sPresetClassType,
			"presetID":      oCTn.presetID,
			"presetSubtype": oCTn.presetSubtype,
			"repeatCount":   oCTn.repeatCount,
			"repeatDur":     oCTn.repeatDur,
			"restart":       sRestartType,
			"spd":           oCTn.spd,
			"syncBehavior":  sSyncBehaviorType,
			"tmFilter":      oCTn.tmFilter
		}
	};
	WriterToJSON.prototype.SerSeq = function(oSeq)
	{
		var sNextAcType = undefined;
		switch (oSeq.nextAc)
		{
			case c_oAscSlideNextAcType.None:
				sNextAcType = "none";
				break;
			case c_oAscSlideNextAcType.Seek:
				sNextAcType = "seek";
				break;
		}

		var sPrevAcType = undefined;
		switch (oSeq.prevAc)
		{
			case c_oAscSlidePrevAcType.None:
				sPrevAcType = "none";
				break;
			case c_oAscSlidePrevAcType.SkipTimed:
				sPrevAcType = "skipTimed";
				break;
		}

		return {
			"cTn":         this.SerCTn(oSeq.cTn),
			"nextCondLst": this.SerCondLst(oSeq.nextCondLst),
			"prevCondLst": this.SerCondLst(oSeq.prevCondLst),
			"concurrent":  oSeq.concurrent,
			"nextAc":      sNextAcType,
			"prevAc":      sPrevAcType,
			"objType":     "seq"
		}	
	};
	WriterToJSON.prototype.SerIterate = function(oIterate)
	{
		if (!oIterate)
			return undefined;

		var sIterateType = undefined;
		switch (oIterate.type)
		{
			case c_oAscSlideIterateType.El:
				sIterateType = "el";
				break;
			case c_oAscSlideIterateType.Lt:
				sIterateType = "lt";
				break;
			case c_oAscSlideIterateType.Wd:
				sIterateType = "wd";
				break;
		}
		return {
			"tmAbs":     oIterate.tmAbs,
			"tmPct":     oIterate.tmPct,
			"backwards": oIterate.backwards,
			"type":      sIterateType
		}
	};
	WriterToJSON.prototype.SerAudio = function(oAudio)
	{
		return {
			"cMediaNode":  this.SerCMediaNode(oAudio.cMediaNode),
			"isNarration": oAudio.isNarration,
			"objType":     "audio"
		}	
	};
	WriterToJSON.prototype.SerCMediaNode = function(oCMediaNode)
	{
		if (!oCMediaNode)
			return undefined;
			
		return {
			"cTn":             this.SerCTn(oCMediaNode.cTn),
			"tgtEl":           this.SerTgtEl(oCMediaNode.tgtEl),
			"mute":            oCMediaNode.mute,
			"numSld":          oCMediaNode.numSld,
			"showWhenStopped": oCMediaNode.showWhenStopped,
			"vol":             oCMediaNode.vol
		}
	};
	WriterToJSON.prototype.SerVideo = function(oVideo)
	{
		return {
			"cMediaNode":  this.SerCMediaNode(oVideo.cMediaNode),
			"fullScrn":    oVideo.fullScrn,
			"objType":     "video"
		}	
	};
	WriterToJSON.prototype.SerExcl = function(oExcl)
	{
		return {
			"cTn":     this.SerCTn(oExcl.cTn),
			"objType": "excl"
		}
	};
	WriterToJSON.prototype.SerAnim = function(oAnim)
	{
		var sCalcmodeType = undefined;
		switch (oAnim.calcmode)
		{
			case c_oAscSlideCalcModeType.Discrete:
				sCalcmodeType = "discrete";
				break;
			case c_oAscSlideCalcModeType.Lin:
				sCalcmodeType = "lin";
				break;
			case c_oAscSlideCalcModeType.Fmla:
				sCalcmodeType = "fmla";
				break;
		}

		var sValueType = undefined;
		switch (oAnim.valueType)
		{
			case c_oAscSlideTLValueType.Num:
				sValueType = "num";
				break;
			case c_oAscSlideTLValueType.Clr:
				sValueType = "clr";
				break;
			case c_oAscSlideTLValueType.Str:
				sValueType = "str";
				break;
		}

		return {
			"cBhvr":     this.SerCBhvr(oAnim.cBhvr),
			"tavLst":    this.SerTavLst(oAnim.tavLst),
			"by":        oAnim.by,
			"calcmode":  sCalcmodeType,
			"from":      oAnim.from,
			"to":        oAnim.to,
			"valueType": sValueType,
			"objType":   "anim"
		}
	};
	WriterToJSON.prototype.SerCBhvr = function(oCBhvr)
	{
		if (!oCBhvr)
			return undefined;

		var sAccumulateType = undefined;
		switch (oCBhvr.accumulate)
		{
			case c_oAscSlideTLAccumulateType.Always:
				sAccumulateType = "always";
				break;
			case c_oAscSlideTLAccumulateType.None:
				sAccumulateType = "none";
				break;
		}

		var sAdditiveType = undefined;
		switch (oCBhvr.additive)
		{
			case c_oAscSlideTLAdditiveType.Base:
				sAdditiveType = "base";
				break;
			case c_oAscSlideTLAdditiveType.Mult:
				sAdditiveType = "mult";
				break;
			case c_oAscSlideTLAdditiveType.None:
				sAdditiveType = "none";
				break;
			case c_oAscSlideTLAdditiveType.Repl:
				sAdditiveType = "repl";
				break;
			case c_oAscSlideTLAdditiveType.Sum:
				sAdditiveType = "sum";
				break;
		}

		var sOverrideType = undefined;
		switch (oCBhvr.override)
		{
			case c_oAscSlideTLOverrideType.ChildStyle:
				sOverrideType = "childStyle";
				break;
			case c_oAscSlideTLOverrideType.Normal:
				sOverrideType = "normal";
				break;
		}

		var sXfrmTypeType = undefined;
		switch (oCBhvr.xfrmType)
		{
			case c_oAscSlideTLTransformType.Img:
				sXfrmTypeType = "img";
				break;
			case c_oAscSlideTLTransformType.Pt:
				sXfrmTypeType = "pt";
				break;
		}

		return {
			"attrNameLst": this.SerAttrNameLst(oCBhvr.attrNameLst),
			"cTn":         this.SerCTn(oCBhvr.cTn),
			"tgtEl":       this.SerTgtEl(oCBhvr.tgtEl),

			"accumulate":  sAccumulateType,
			"additive":    sAdditiveType,
			"by":          oCBhvr.by,
			"from":        oCBhvr.from,
			"override":    sOverrideType,
			"rctx":        oCBhvr.rctx,
			"to":          oCBhvr.to,
			"xfrmType":    sXfrmTypeType
		}
	};
	WriterToJSON.prototype.SerTavLst = function(oTavLst)
	{
		if (!oTavLst)
			return undefined;

		var aTavLst = [];
		for (var nTav = 0; nTav < oTavLst.list.length; nTav++)
			aTavLst.push(this.SerTav(oTavLst.list[nTav]));

		return aTavLst;
	};
	WriterToJSON.prototype.SerTav = function(oTav)
	{
		return {
			"val":  this.SerAnimVariant(oTav.val),
			"fmla": oTav.fmla,
			"tm":   oTav.tm
		};
	};
	WriterToJSON.prototype.SerAttrNameLst = function(oAttrNameLst)
	{
		if (!oAttrNameLst)
			return undefined;

		var aAttrNameLst = [];
		for (var nName = 0; nName < oAttrNameLst.list.length; nName++)
			aAttrNameLst.push(oAttrNameLst.list[nName].text);

		return aAttrNameLst;
	};
	WriterToJSON.prototype.SerAnimClr = function(oAnimClr)
	{
		var sClrSpcType = undefined;
		switch (oAnimClr.clrSpc)
		{
			case c_oAscSlideTLColorSpaceType.Rgb:
				sClrSpcType = "rgb";
				break;
			case c_oAscSlideTLColorSpaceType.Hsl:
				sClrSpcType = "hsl";
				break;
		}

		var sColorDir = undefined;
		switch (oAnimClr.dir)
		{
			case AscCommon.ST_HueDir.Ccw:
				sColorDir = "ccw";
				break;
			case AscCommon.ST_HueDir.Cw:
				sColorDir = "cw";
				break;
		}

		return {
			"by": {
				"rgb": this.SerByRGB(oAnimClr.byRGB),
				"hsl": this.SerByHSL(oAnimClr.byHSL)
			},

			"cBhvr":   this.SerCBhvr(oAnimClr.cBhvr),
			"from":    oAnimClr.from,
			"to":      oAnimClr.to,
			"clrSpc":  sClrSpcType,
			"dir":     sColorDir,
			"objType": "animClr"
		}
	};
	WriterToJSON.prototype.SerByRGB = function(oByRGB)
	{
		if (!oByRGB)
			return undefined;

		return {
			"r": oByRGB.c1,
			"g": oByRGB.c2,
			"b": oByRGB.c3
		}
	};
	WriterToJSON.prototype.SerByHSL = function(oByHSL)
	{
		if (!oByHSL)
			return undefined;

		return {
			"h": oByHSL.c1,
			"s": oByHSL.c2,
			"l": oByHSL.c3
		}
	};
	WriterToJSON.prototype.SerAnimEffect = function(oAnimEffect)
	{

		return {
			"cBhvr":      this.SerCBhvr(oAnimEffect.cBhvr),
			"progress":   this.SerAnimVariant(oAnimEffect.progress),
			"filter":     oAnimEffect.filter,
			"prLst":      oAnimEffect.prLst,
			"transition": To_XML_TLAnimateEffectTransition(oAnimEffect.transition),
			"objType":    "animEffect"
		}
	};
	WriterToJSON.prototype.SerAnimMotion = function(oAnimMotion)
	{
		var sOriginType = undefined;
		switch (oAnimMotion.origin)
		{
			case c_oAscSlideTLOriginType.Parent:
				sOriginType = "parent";
				break;
			case c_oAscSlideTLOriginType.Layout:
				sOriginType = "layout";
				break;
		}

		var sPathEditMode = undefined;
		switch (oAnimMotion.pathEditMode)
		{
			case c_oAscSlideTLPathEditMode.Fixed:
				sPathEditMode = "fixed";
				break;
			case c_oAscSlideTLPathEditMode.Relative:
				sPathEditMode = "relative";
				break;
		}

		return {
			"by":    this.SerByPoint(oAnimMotion.by),
			"cBhvr": this.SerCBhvr(oAnimMotion.cBhvr),
			"from":  this.SerByPoint(oAnimMotion.from),
			"rCtr":  this.SerByPoint(oAnimMotion.rCtr),
			"to":    this.SerByPoint(oAnimMotion.to),
			"origin":       sOriginType,
			"path":         oAnimMotion.path,
			"pathEditMode": sPathEditMode,
			"ptsTypes":     oAnimMotion.ptsTypes,
			"rAng":         oAnimMotion.rAng,
			"objType":      "animMotion"
		}
	};
	WriterToJSON.prototype.SerByPoint = function(oPoint)
	{
		if (!oPoint)
			return undefined;

		return {
			"x": oPoint.x,
			"y": oPoint.y
		}
	};
	WriterToJSON.prototype.SerAnimRot = function(oAnimRot)
	{
		return {
			"cBhvr":   this.SerCBhvr(oAnimRot.cBhvr),
			"by":      oAnimRot.by,
			"from":    oAnimRot.from,
			"to":      oAnimRot.to,
			"objType": "animRot"
		}
	};
	WriterToJSON.prototype.SerAnimScale = function(oAnimScale)
	{
		return {
			"cBhvr":        this.SerCBhvr(oAnimScale.cBhvr),
			"by":           this.SerByPoint(oAnimScale.by),
			"from":         this.SerByPoint(oAnimScale.from),
			"to":           this.SerByPoint(oAnimScale.to),
			"zoomContents": oAnimScale.zoomContents,
			"objType":      "animScale"
		}	
	};
	WriterToJSON.prototype.SerCmd = function(oCmd)
	{
		var sCommandType = undefined;
		switch (oCmd.type)
		{
			case c_oAscSlideTLCommandType.Call:
				sCommandType = "call";
				break;
			case c_oAscSlideTLCommandType.Evt:
				sCommandType = "evt";
				break;
			case c_oAscSlideTLCommandType.Verb:
				sCommandType = "verb";
				break;
		}
		return {
			"cBhvr":   this.SerCBhvr(oCmd.cBhvr),
			"cmd":     oCmd.cmd,
			"type":    sCommandType,
			"objType": "cmd"
		}
	};
	WriterToJSON.prototype.SerSet = function(oSet)
	{
		return {
			"cBhvr":   this.SerCBhvr(oSet.cBhvr),
			"to":      this.SerAnimVariant(oSet.to),
			"objType": "set"
		}
	};
	WriterToJSON.prototype.SerAnimVariant = function(oAnimVariant)
	{
		if (!oAnimVariant)
			return undefined;

		return {
			"boolVal": oAnimVariant.boolVal,
			"strVal":  oAnimVariant.strVal,
			"intVal":  oAnimVariant.intVal,
			"fltVal":  oAnimVariant.fltVal,
			"clrVal":  this.SerColor(oAnimVariant.clrVal)
		}
	};
	WriterToJSON.prototype.SerCondLst = function(oCondLst)
	{
		if (!oCondLst)
			return undefined;

		var aCondLst = [];
		for (var nCond = 0; nCond < oCondLst.list.length; nCond++)
			aCondLst.push(this.SerCond(oCondLst.list[nCond]));

		return aCondLst;
	};
	WriterToJSON.prototype.SerCond = function(oCond)
	{
		if (!oCond)
			return undefined;

		var sEventType = undefined;
		switch (oCond.evt)
		{
			case c_oAscSlideTriggerEventType.Begin:
				sEventType = "begin";
				break;
			case c_oAscSlideTriggerEventType.End:
				sEventType = "end";
				break;
			case c_oAscSlideTriggerEventType.OnBegin:
				sEventType = "onBegin";
				break;
			case c_oAscSlideTriggerEventType.OnClick:
				sEventType = "onClick";
				break;
			case c_oAscSlideTriggerEventType.OnDblClick:
				sEventType = "onDblClick";
				break;
			case c_oAscSlideTriggerEventType.OnEnd:
				sEventType = "onEnd";
				break;
			case c_oAscSlideTriggerEventType.OnMouseOut:
				sEventType = "onMouseOut";
				break;
			case c_oAscSlideTriggerEventType.OnMouseOver:
				sEventType = "onMouseOver";
				break;
			case c_oAscSlideTriggerEventType.OnNext:
				sEventType = "onNext";
				break;
			case c_oAscSlideTriggerEventType.OnPrev:
				sEventType = "onPrev";
				break;
			case c_oAscSlideTriggerEventType.OnStopAudio:
				sEventType = "onStopAudio";
				break;
		}

		return {
			"rtn":   this.SerRtn(oCond.rtn),
			"tgtEl": this.SerTgtEl(oCond.tgtEl),
			"tn":    oCond.tn,
			"delay": oCond.delay,
			"evt":   sEventType
		}
	};
	WriterToJSON.prototype.SerRtn = function(nRtn)
	{
		if (!AscFormat.isRealNumber(nRtn))
			return undefined;
		var sType = undefined;
		switch (nRtn)
		{
			case c_oAscSlideRuntimeTriggerType.All:
				sType = "all";
				break;
			case c_oAscSlideRuntimeTriggerType.First:
				sType = "first";
				break;
			case c_oAscSlideRuntimeTriggerType.Last:
				sType = "last";
				break;
		}

		return {
			"val": sType
		}
	};
	WriterToJSON.prototype.SerTgtEl = function(oTgtEl)
	{
		if (!oTgtEl)
			return undefined;

		return {
			"inkTgt": this.SerInkTgt(oTgtEl.inkTgt),
			"sldTgt": undefined, /// ???
			"sndTgt": this.SerSndTgt(oTgtEl.sndTgt),
			"spTgt":  this.SerSpTgt(oTgtEl.spTgt)
		}
	};
	WriterToJSON.prototype.SerInkTgt = function(oInkTgt)
	{
		if (!oInkTgt)
			return undefined;

		let nSpId = this.GetSpIdxId(oInkTgt.spid);

		return {
			"spid": nSpId != null ? nSpId : undefined
		}
	};
	WriterToJSON.prototype.SerSndTgt = function(oSndTgt)
	{
		if (!oSndTgt)
			return undefined;

		return {
			"embed":   oSndTgt.embed,
			"name":    oSndTgt.name,
			"builtIn": oSndTgt.builtIn
		}
	};
	WriterToJSON.prototype.SerSpTgt = function(oSpTgt)
	{
		if (!oSpTgt)
			return undefined;

		let nSpId = this.GetSpIdxId(oSpTgt.spid);
		let nSubSpId = this.GetSpIdxId(oSpTgt.subSpId);

		return {
			"spid":       nSpId != null ? nSpId : undefined,
			"bg":         oSpTgt.bg,
			"subSpId":    nSubSpId != null ? nSubSpId : undefined,
			"oleChartEl": this.SerOleChartEl(oSpTgt.oleChartEl),
			"txEl":       this.SerTxEl(oSpTgt.txEl),
			"graphicEl":  this.SerGraphicEl(oSpTgt.graphicEl)
		}
	};
	WriterToJSON.prototype.SerOleChartEl = function(oOleChartEl)
	{
		if (!oOleChartEl)
			return undefined;

		var sType = undefined;
		switch (oOleChartEl.type)
		{
			case c_oAscSlideChartSubElementType.Category:
				sType = "category";
				break;
			case c_oAscSlideChartSubElementType.GridLegend:
				sType = "gridLegend";
				break;
			case c_oAscSlideChartSubElementType.PtInCategory:
				sType = "ptInCategory";
				break;
			case c_oAscSlideChartSubElementType.PtInSeries:
				sType = "ptInSeries";
				break;
			case c_oAscSlideChartSubElementType.Series:
				sType = "series";
				break;
		}
	
		return {
			"lvl":  oOleChartEl.lvl,
			"type": sType
		}
	};
	WriterToJSON.prototype.SerTxEl = function(oTxEl)
	{
		if (!oTxEl)
			return undefined;

		return {
			"charRg": this.SerIndexRg(oTxEl.charRg),
			"pRg":    this.SerIndexRg(oTxEl.pRg)
		}
	};
	WriterToJSON.prototype.SerIndexRg = function(oIndexRg)
	{
		if (!oIndexRg)
			return undefined;
			
		return {
			"st":  oIndexRg.st,
			"end": oIndexRg.end
		}
	};
	WriterToJSON.prototype.SerGraphicEl = function(oGraphicEl)
	{
		if (!oGraphicEl)
			return undefined;
		
		let nSpId = this.GetSpIdxId(oGraphicEl.dgmId);

		return {
			"chart": {
				"bldStep":     oGraphicEl.chartBuildStep,
				"categoryIdx": oGraphicEl.categoryIdx,
				"seriesIdx":   oGraphicEl.seriesIdx
			},
			"dgm": {
				"bldStep": oGraphicEl.dgmBuildStep,
				"id":      nSpId != null ? nSpId : undefined
			}
		}
	};

	ReaderFromJSON.prototype.MasterSlideFromJSON = function(oParsedMaster, oPres)
	{
		var oMasterSlide = new AscCommonSlide.MasterSlide(undefined, null);

		var oTheme  = oParsedMaster["theme"] ? this.ThemeFromJSON(oParsedMaster["theme"]) : null;
		var oClrMap = oParsedMaster["clrMapOvr"] ? this.ColorMapOvrFromJSON(oParsedMaster["clrMapOvr"]) : null;

		oTheme && oMasterSlide.setTheme(oTheme);
		oClrMap && oMasterSlide.setClMapOverride(oClrMap);

		// cSld
		var oCSld = this.CSldFromJSON(oParsedMaster["cSld"]);
		for (var nShape = 0; nShape < oCSld.spTree.length; nShape++)
			oMasterSlide.shapeAdd(nShape, oCSld.spTree[nShape]);
		oCSld.Bg && oMasterSlide.changeBackground(oCSld.Bg);
		oCSld.name && oMasterSlide.setCSldName(oCSld.name);

		// hf
		oParsedMaster["hf"] && oMasterSlide.setHF(this.HFFromJSON(oParsedMaster["hf"]));

		// layouts
		for (var nLayout = 0; nLayout < oParsedMaster["sldLayoutLst"].length; nLayout++)
			oMasterSlide.addToSldLayoutLstToPos(oMasterSlide.sldLayoutLst.length, this.SlideLayoutFromJSON(oParsedMaster["sldLayoutLst"][nLayout]));

		// transition
		oParsedMaster["transition"] && oMasterSlide.applyTransition(this.TransitionFromJSON(oParsedMaster["transition"]));

		// timing 
		oParsedMaster["timing"] && oMasterSlide.setTiming(this.TimingFromJSON(oParsedMaster["timing"]));

		//txStyles
		oParsedMaster["txStyles"] && oMasterSlide.setTxStyles(this.TxStylesFromJSON(oParsedMaster["txStyles"]));

		oMasterSlide.preserve    = oParsedMaster["preserve"];
		oMasterSlide.ImageBase64 = oParsedMaster["imgBase64"];

		if (!oPres)
			oPres = private_GetPresentation();

		oPres.addSlideMaster(oPres.slideMasters.length, oMasterSlide);
		this.mastersMap[oParsedMaster["id"]] = oMasterSlide;

		return oMasterSlide;
	};
	ReaderFromJSON.prototype.TxStylesFromJSON = function(oParsedTxStyles)
	{
		var oTxStyles = new AscFormat.CTextStyles();

		if (oParsedTxStyles["bodyStyle"])
			oTxStyles.bodyStyle = this.LstStyleFromJSON(oParsedTxStyles["bodyStyle"]);
		if (oParsedTxStyles["otherStyle"])
			oTxStyles.otherStyle = this.LstStyleFromJSON(oParsedTxStyles["otherStyle"]);
		if (oParsedTxStyles["titleStyle"])
			oTxStyles.titleStyle = this.LstStyleFromJSON(oParsedTxStyles["titleStyle"]);

		return oTxStyles;
	};
	ReaderFromJSON.prototype.TableStylesFromJSON = function(oParsedStyles)
	{
		this.RestoredStylesMap = {};
		// восстанавливаем все стили и мапим по старым id (т.к. создаются с новыми)
		for (var key in oParsedStyles)
			this.RestoredStylesMap[key] = this.TableStyleFromJSON(oParsedStyles[key]);
	};
	ReaderFromJSON.prototype.TableStyleFromJSON = function(oParsedStyle)
	{
		let oTableStyle = new AscCommonWord.CStyle(oParsedStyle["name"], null, null, styletype_Table);

		if (oParsedStyle["band1H"])
			oTableStyle.Set_TableBand1Horz(this.TableStylePartFromJSON(oParsedStyle["band1H"]));
		if (oParsedStyle["band1V"])
			oTableStyle.Set_TableBand1Vert(this.TableStylePartFromJSON(oParsedStyle["band1V"]));
		if (oParsedStyle["band2H"])
			oTableStyle.Set_TableBand2Horz(this.TableStylePartFromJSON(oParsedStyle["band2H"]));
		if (oParsedStyle["band2V"])
			oTableStyle.Set_TableBand2Vert(this.TableStylePartFromJSON(oParsedStyle["band2V"]));
		if (oParsedStyle["firstCol"])
			oTableStyle.Set_TableFirstCol(this.TableStylePartFromJSON(oParsedStyle["firstCol"]));
		if (oParsedStyle["firstRow"])
			oTableStyle.Set_TableFirstRow(this.TableStylePartFromJSON(oParsedStyle["firstRow"]));
		if (oParsedStyle["lastCol"])
			oTableStyle.Set_TableLastCol(this.TableStylePartFromJSON(oParsedStyle["lastCol"]));
		if (oParsedStyle["lastRow"])
			oTableStyle.Set_TableLastRow(this.TableStylePartFromJSON(oParsedStyle["lastRow"]));
		if (oParsedStyle["wholeTbl"])
			oTableStyle.Set_TableWholeTable(this.TableStylePartFromJSON(oParsedStyle["wholeTbl"]));
		if (oParsedStyle["neCell"])
			oTableStyle.Set_TableTLCell(this.TableStylePartFromJSON(oParsedStyle["neCell"]));
		if (oParsedStyle["nwCell"])
			oTableStyle.Set_TableTRCell(this.TableStylePartFromJSON(oParsedStyle["nwCell"]));
		if (oParsedStyle["seCell"])
			oTableStyle.Set_TableBLCell(this.TableStylePartFromJSON(oParsedStyle["seCell"]));
		if (oParsedStyle["swCell"])
			oTableStyle.Set_TableBRCell(this.TableStylePartFromJSON(oParsedStyle["swCell"]));
		if (oParsedStyle["tblBg"])
			oTableStyle.Set_TablePr(this.TableCellStyleFromJSON(null, oParsedStyle["tblBg"], true));

		var oStyles = private_GetStyles();
		if (oTableStyle)
		{
			// если такого стиля нет - добавляем новый
			var nExistingStyleId = oStyles.GetStyleIdByName(oTableStyle.Name);
			if (nExistingStyleId === null)
			{
				oStyles.Add(oTableStyle);
			}
			else
			{
				var oExistingStyle = oStyles.Get(nExistingStyleId);
				// если стили идентичны, стиль не добавляем
				if (!oTableStyle.IsEqual(oExistingStyle))
					oStyles.Add(oTableStyle);
				else
					oTableStyle = oExistingStyle;
			}
		}

		return oTableStyle;
	};
	ReaderFromJSON.prototype.TableStylePartFromJSON = function(oParsed)
	{
		let oPart = new CTableStylePr();
		let oCellPr = oPart.TableCellPr;
		let oTextPr = oPart.TextPr;
		
		if (oParsed["tcTxStyle"])
			this.TcTxStyleFromJSON(oTextPr, oParsed["tcTxStyle"]);
		if (oParsed["tcStyle"])
			this.TableCellStyleFromJSON(oCellPr, oParsed["tcStyle"]);

		return oPart;
	};
	ReaderFromJSON.prototype.TableCellStyleFromJSON = function(oCellPr, oParsed, isTableBorders)
	{
		if (!oCellPr)
			oCellPr = new CTablePr();

		let oBorders = isTableBorders ? oCellPr.TableBorders : oCellPr.TableCellBorders;
		if (oParsed["fill"])
		{
			let oFill = this.FillFromJSON(oParsed["fill"]);
			if(!oCellPr.Shd) {
				oCellPr.Shd = new AscCommonWord.CDocumentShd();
				oCellPr.Shd.Value = c_oAscShdClear;
			}
			oCellPr.Shd.Unifill = oFill;
		}
		if (oParsed["fillRef"])
		{
			let oStyleRef = this.StyleRefFromJSON(oParsed["fillRef"]);
			if(!oCellPr.Shd) {
				oCellPr.Shd = new AscCommonWord.CDocumentShd();
				oCellPr.Shd.Value = c_oAscShdClear;
			}
			oCellPr.Shd.FillRef = oStyleRef;
		}
		if (oParsed["tcBdr"])
		{
			if (oParsed["tcBdr"]["bottom"])
				oBorders.Bottom = this.TableCellBorderFromJSON(oParsed["tcBdr"]["bottom"]);
			if (oParsed["tcBdr"]["insideH"])
				oBorders.InsideH = this.TableCellBorderFromJSON(oParsed["tcBdr"]["insideH"]);
			if (oParsed["tcBdr"]["insideV"])
				oBorders.InsideV = this.TableCellBorderFromJSON(oParsed["tcBdr"]["insideV"]);
			if (oParsed["tcBdr"]["left"])
				oBorders.Left = this.TableCellBorderFromJSON(oParsed["tcBdr"]["left"]);
			if (oParsed["tcBdr"]["right"])
				oBorders.Right = this.TableCellBorderFromJSON(oParsed["tcBdr"]["right"]);
			if (oParsed["tcBdr"]["top"])
				oBorders.Top = this.TableCellBorderFromJSON(oParsed["tcBdr"]["top"]);
		}

		return oCellPr;
	};
	ReaderFromJSON.prototype.TableCellBorderFromJSON = function(oParsed)
	{
		let oBorder = new CDocumentBorder();
		if (oParsed["ln"])
		{
			let oLn = this.LnFromJSON(oParsed["ln"]);
			oLn.fillDocumentBorder(oBorder)
		}
		else if (oParsed["lnRef"])
		{
			oBorder.LineRef = this.StyleRefFromJSON(oParsed["lnRef"]);
			oBorder.Value = AscCommonWord.border_Single;
		}

		return oBorder;
	};
	ReaderFromJSON.prototype.TcTxStyleFromJSON = function(oTcTxPr, oParsed)
	{
		if (oParsed["i"] != null)
			oTcTxPr.Italic = oParsed["i"] === "on";
		if (oParsed["b"] != null)
			oTcTxPr.Bold = oParsed["b"] === "on";
		if (oParsed["fontRef"])
			oTcTxPr.FontRef = this.FontRefFromJSON(oParsed["fontRef"]);
		if (oParsed["color"])
		{
			let oColor = this.ColorFromJSON(oParsed["color"]);
			oTcTxPr.Unifill = AscFormat.CreateUniFillByUniColor(oColor);
		}
	};
	ReaderFromJSON.prototype.SlideLayoutFromJSON = function(oParsedLayout)
	{
		var oLayout = new AscCommonSlide.SlideLayout();

		var nLayoutType = undefined;
		switch (oParsedLayout["ltType"])
		{
			case "blank":
				nLayoutType = c_oAscSlideLayoutType.Blank;
				break;
			case "chart":
				nLayoutType = c_oAscSlideLayoutType.Chart;
				break;
			case "chartAndTx":
				nLayoutType = c_oAscSlideLayoutType.ChartAndTx;
				break;
			case "clipArtAndTx":
				nLayoutType = c_oAscSlideLayoutType.ClipArtAndTx;
				break;
			case "clipArtAndVertTx":
				nLayoutType = c_oAscSlideLayoutType.ClipArtAndVertTx;
				break;
			case "cust":
				nLayoutType = c_oAscSlideLayoutType.Cust;
				break;
			case "dgm":
				nLayoutType = c_oAscSlideLayoutType.Dgm;
				break;
			case "fourObj":
				nLayoutType = c_oAscSlideLayoutType.FourObj;
				break;
			case "mediaAndTx":
				nLayoutType = c_oAscSlideLayoutType.MediaAndTx;
				break;
			case "obj":
				nLayoutType = c_oAscSlideLayoutType.Obj;
				break;
			case "objAndTwoObj":
				nLayoutType = c_oAscSlideLayoutType.ObjAndTwoObj;
				break;
			case "objAndTx":
				nLayoutType = c_oAscSlideLayoutType.ObjAndTx;
				break;
			case "objOnly":
				nLayoutType = c_oAscSlideLayoutType.ObjOnly;
				break;
			case "objOverTx":
				nLayoutType = c_oAscSlideLayoutType.ObjOverTx;
				break;
			case "objTx":
				nLayoutType = c_oAscSlideLayoutType.ObjTx;
				break;
			case "picTx":
				nLayoutType = c_oAscSlideLayoutType.PicTx;
				break;
			case "secHead":
				nLayoutType = c_oAscSlideLayoutType.SecHead;
				break;
			case "tbl":
				nLayoutType = c_oAscSlideLayoutType.Tbl;
				break;
			case "title":
				nLayoutType = c_oAscSlideLayoutType.Title;
				break;
			case "titleOnly":
				nLayoutType = c_oAscSlideLayoutType.TitleOnly;
				break;
			case "twoColTx":
				nLayoutType = c_oAscSlideLayoutType.TwoColTx;
				break;
			case "twoObj":
				nLayoutType = c_oAscSlideLayoutType.TwoObj;
				break;
			case "twoObjAndObj":
				nLayoutType = c_oAscSlideLayoutType.TwoObjAndObj;
				break;
			case "twoObjAndTx":
				nLayoutType = c_oAscSlideLayoutType.TwoObjAndTx;
				break;
			case "twoObjOverTx":
				nLayoutType = c_oAscSlideLayoutType.TwoObjOverTx;
				break;
			case "twoTxTwoObj":
				nLayoutType = c_oAscSlideLayoutType.TwoTxTwoObj;
				break;
			case "tx":
				nLayoutType = c_oAscSlideLayoutType.Tx;
				break;
			case "txAndChart":
				nLayoutType = c_oAscSlideLayoutType.TxAndChart;
				break;
			case "txAndClipArt":
				nLayoutType = c_oAscSlideLayoutType.TxAndClipArt;
				break;
			case "txAndMedia":
				nLayoutType = c_oAscSlideLayoutType.TxAndMedia;
				break;
			case "txAndObj":
				nLayoutType = c_oAscSlideLayoutType.TxAndObj;
				break;
			case "txAndTwoObj":
				nLayoutType = c_oAscSlideLayoutType.TxAndTwoObj;
				break;
			case "txOverObj":
				nLayoutType = c_oAscSlideLayoutType.TxOverObj;
				break;
			case "vertTitleAndTx":
				nLayoutType = c_oAscSlideLayoutType.VertTitleAndTx;
				break;
			case "vertTitleAndTxOverChart":
				nLayoutType = c_oAscSlideLayoutType.VertTitleAndTxOverChart;
				break;
			case "vertTx":
				nLayoutType = c_oAscSlideLayoutType.VertTx;
				break;
		}

		var oClrMap = oParsedLayout["clrMapOvr"] ? this.ColorMapOvrFromJSON(oParsedLayout["clrMapOvr"]) : null;
		oClrMap && oLayout.setClMapOverride(oClrMap);

		// cSld
		var oCSld = this.CSldFromJSON(oParsedLayout["cSld"]);
		for (var nShape = 0; nShape < oCSld.spTree.length; nShape++)
			oLayout.shapeAdd(nShape, oCSld.spTree[nShape]);
		oCSld.Bg && oLayout.changeBackground(oCSld.Bg);
		oCSld.name && oLayout.setCSldName(oCSld.name);

		// hf
		oParsedLayout["hf"] && oLayout.setHF(this.HFFromJSON(oParsedLayout["hf"]));

		// timing
		oParsedLayout["timing"] && oLayout.setTiming(this.TimingFromJSON(oParsedLayout["timing"]));
		// transition
		oParsedLayout["transition"] && oLayout.applyTransition(this.TransitionFromJSON(oParsedLayout["transition"]));

		oParsedLayout["matchingName"] && oLayout.setMatchingName(oParsedLayout["matchingName"]);
		oLayout.preserve = oParsedLayout["preserve"];	
		oParsedLayout["showMasterPhAnim"] && oLayout.setShowPhAnim(oParsedLayout["showMasterPhAnim"]);
		oParsedLayout["showMasterSp"] && oLayout.setShowMasterSp(oParsedLayout["showMasterSp"]);
		oLayout.userDrawn = oParsedLayout["userDrawn"];
		oLayout.setType(nLayoutType);
		oLayout.ImageBase64 = oParsedLayout["imgBase64"];

		// мапим, чтобы повторно не восстанавливать
		this.layoutsMap[oParsedLayout["id"]] = oLayout;

		if (this.mastersMap[oParsedLayout["master"]])
		{
			var oMaster = this.mastersMap[oParsedLayout["master"]];
			oMaster.addToSldLayoutLstToPos(oMaster.sldLayoutLst.length, oLayout);
		}

		return oLayout;
	};
	ReaderFromJSON.prototype.SlideFromJSON = function(oParsedSlide, oPres)
	{
		var oPresentation = oPres || private_GetPresentation();

		var oLayout = null;
		var oMaster = null;

		if (typeof oParsedSlide["layout"] === "object")
			oLayout = this.SlideLayoutFromJSON(oParsedSlide["layout"]);
		if (typeof oParsedSlide["master"] === "object")
			oMaster = this.MasterSlideFromJSON(oParsedSlide["master"], oPresentation);

		if (oMaster)
		{
			if (!oLayout)
				oLayout = this.layoutsMap[oParsedSlide["layout"]];
			else
				oMaster.addToSldLayoutLstToPos(oMaster.length, oLayout);
		}
			
		if (!oLayout)
		{
			if (!this.layoutsMap[oParsedSlide["layout"]])
			{
				var CurSlide = oPresentation.Slides[oPresentation.CurPage];
				if (!CurSlide)
				{
					oMaster = oPresentation.getDefaultMasterSlide();
					oLayout = oMaster.sldLayoutLst[0];
				}
				else
					oLayout = CurSlide.Layout;
			}
			else
				oLayout = this.layoutsMap[oParsedSlide["layout"]];
		}

		// установим MasterSlide для Layout, если не задан
		if (!oLayout.Master)
		{
			oMaster = oPresentation.getDefaultMasterSlide();
			oLayout.setMaster(oMaster);
			oMaster.addToSldLayoutLstToPos(oMaster.length, oLayout);
		}

		var oSlide = new AscCommonSlide.Slide(oPresentation, oLayout, 0);
		
		oParsedSlide["notes"] && oSlide.setNotes(this.NotesFromJSON(oParsedSlide["notes"], oPresentation));
		oParsedSlide["clrMapOvr"] && oSlide.setClMapOverride(this.ColorMapOvrFromJSON(oParsedSlide["clrMapOvr"]));

		// cSld
		var oCSld = this.CSldFromJSON(oParsedSlide["cSld"]);
		for (var nShape = 0; nShape < oCSld.spTree.length; nShape++)
			oSlide.shapeAdd(nShape, oCSld.spTree[nShape]);
		oCSld.Bg && oSlide.changeBackground(oCSld.Bg);
		oCSld.name && oSlide.setCSldName(oCSld.name);

		// comments
		for (var nComment = 0; nComment < oParsedSlide["comments"].length; nComment++)
			oSlide.slideComments.addComment(this.CommentFromJSON(oParsedSlide["comments"][nComment], oSlide.slideComments));

		oParsedSlide["transition"] && oSlide.applyTransition(this.TransitionFromJSON(oParsedSlide["transition"]));
		oParsedSlide["timing"] && oSlide.setTiming(this.TimingFromJSON(oParsedSlide["timing"]));
		oParsedSlide["show"] != undefined && oSlide.setShow(oParsedSlide["show"]);
		oParsedSlide["showMasterPhAnim"] != undefined && oSlide.setShowPhAnim(oParsedSlide["showMasterPhAnim"]);
		oParsedSlide["showMasterSp"] != undefined && oSlide.setShowMasterSp(oParsedSlide["showMasterSp"]);

		oSlide.setSlideSize(oPresentation.GetWidthMM(), oPresentation.GetHeightMM());

		return oSlide;
	};
	ReaderFromJSON.prototype.SlidesFromJSON = function(oParsedSlides, oPres)
	{
		var oPresentation = oPres || private_GetPresentation();
		var aSlides       = [];

		for (var nMaster = 0; nMaster < oParsedSlides["masters"].length; nMaster++)
			this.MasterSlideFromJSON(oParsedSlides["masters"][nMaster], oPresentation);
		
		for (var nLayout = 0; nLayout < oParsedSlides["layouts"].length; nLayout++)
			this.SlideLayoutFromJSON(oParsedSlides["layouts"][nLayout]);

		for (var nSlide = 0; nSlide < oParsedSlides["slides"].length; nSlide++)
			aSlides.push(this.SlideFromJSON(oParsedSlides["slides"][nSlide]));

		return aSlides;
	};
	ReaderFromJSON.prototype.SlideSizeFromJSON = function(oParsedSldSize)
	{
		var nSldSzType = undefined;
		switch (oParsedSldSize["type"])
		{
			case Asc.c_oAscSlideSZType.Sz35mm:
				nSldSzType = "35mm";
				break;
			case Asc.c_oAscSlideSZType.SzA3:
				nSldSzType = "A3";
				break;
			case Asc.c_oAscSlideSZType.SzA4:
				nSldSzType = "A4";
				break;
			case Asc.c_oAscSlideSZType.SzB4ISO:
				nSldSzType = "B4ISO";
				break;
			case Asc.c_oAscSlideSZType.SzB4JIS:
				nSldSzType = "B4JIS";
				break;
			case Asc.c_oAscSlideSZType.SzB5ISO:
				nSldSzType = "B5ISO";
				break;
			case Asc.c_oAscSlideSZType.SzB5JIS:
				nSldSzType = "B5JIS";
				break;
			case Asc.c_oAscSlideSZType.SzBanner:
				nSldSzType = "banner";
				break;
			case Asc.c_oAscSlideSZType.SzCustom:
				nSldSzType = "custom";
				break;
			case Asc.c_oAscSlideSZType.SzHagakiCard:
				nSldSzType = "hagakiCard";
				break;
			case Asc.c_oAscSlideSZType.SzLedger:
				nSldSzType = "ledger";
				break;
			case Asc.c_oAscSlideSZType.SzLetter:
				nSldSzType = "letter";
				break;
			case Asc.c_oAscSlideSZType.SzOverhead:
				nSldSzType = "overhead";
				break;
			case Asc.c_oAscSlideSZType.SzScreen16x10:
				nSldSzType = "screen16x10";
				break;
			case Asc.c_oAscSlideSZType.SzScreen16x9:
				nSldSzType = "screen16x9";
				break;
			case Asc.c_oAscSlideSZType.SzScreen4x3:
				nSldSzType = "screen4x3";
				break;
			case Asc.c_oAscSlideSZType.SzWidescreen:
				nSldSzType = "wideScreen";
				break;
		}

		var oSldSize = new AscCommonSlide.CSlideSize();

		oParsedSldSize["cx"] != undefined && oSldSize.setCX(oParsedSldSize["cx"]);
		oParsedSldSize["cy"] != undefined && oSldSize.setCY(oParsedSldSize["cy"]);
		nSldSzType != undefined && oSldSize.setType(oParsedSldSize["cx"]);

		return oSldSize;
	};
	ReaderFromJSON.prototype.ShowPrFromJSON = function(oParsedShowPr)
	{
		var oShowPr = new CShowPr();

		oShowPr.browse = oParsedShowPr["browse"] !== undefined ? oParsedShowPr["browse"] : oShowPr.browse;
		oShowPr.kiosk = oParsedShowPr["kiosk"] !== undefined ? oParsedShowPr["kiosk"] : oShowPr.kiosk;
		oShowPr.penClr = oParsedShowPr["penClr"] !== undefined ? this.ColorFromJSON(oParsedShowPr["penClr"]) : oShowPr.penClr;
		oShowPr.present = oParsedShowPr["present"] !== undefined ? oParsedShowPr["present"] : oShowPr.present;
		oShowPr.show = oParsedShowPr["show"] !== undefined ? oParsedShowPr["show"] : oShowPr.show;
		oShowPr.loop = oParsedShowPr["loop"] !== undefined ? oParsedShowPr["loop"] : oShowPr.loop;
		oShowPr.showAnimation = oParsedShowPr["showAnimation"] !== undefined ? oParsedShowPr["showAnimation"] : oShowPr.showAnimation;
		oShowPr.showNarration = oParsedShowPr["showNarration"] !== undefined ? oParsedShowPr["showNarration"] : oShowPr.showNarration;
		oShowPr.useTimings = oParsedShowPr["useTimings"] !== undefined ? oParsedShowPr["useTimings"] : oShowPr.useTimings;

		return oShowPr;
	};
	ReaderFromJSON.prototype.CommentFromJSON = function(oParsedComment, oParent)
	{
		var oAscCommentData = new Asc.asc_CCommentData({
			m_sText:     oParsedComment["text"],
			m_sUserName: oParsedComment["authorName"],
			m_sUserId:   oParsedComment["authorId"],
			m_sOOTime:   oParsedComment["dt"]
		});

		var oCommentData = new AscCommon.CCommentData();
		oCommentData.Read_FromAscCommentData(oAscCommentData);

		var oComment = new AscCommon.CComment(oParent, oCommentData);
		oComment.setPosition(oParsedComment["pos"]["x"], oParsedComment["pos"]["y"]);

		return oComment;
	};
	ReaderFromJSON.prototype.NotesFromJSON = function(oParsedNotes, oPres)
	{
		var oNotes = new AscCommonSlide.CNotes();

		oParsedNotes["clrMapOvr"] && oNotes.setClMapOverride(this.ColorMapOvrFromJSON(oParsedNotes["clrMapOvr"]));
		// cSld
		var oCSld = this.CSldFromJSON(oParsedNotes["cSld"]);
		for (var nShape = 0; nShape < oCSld.spTree.length; nShape++)
			oNotes.addToSpTreeToPos(nShape, oCSld.spTree[nShape]);
		oCSld.Bg && oNotes.changeBackground(oCSld.Bg);
		oCSld.name && oNotes.setCSldName(oCSld.name);

		oParsedNotes["showMasterPhAnim"] != undefined && oNotes.setShowMasterPhAnim(oParsedNotes["showMasterPhAnim"]);
		oParsedNotes["showMasterSp"] != undefined && oNotes.setShowMasterSp(oParsedNotes["showMasterSp"]);

		var oNotesMaster = typeof oParsedNotes["master"] === "object" ? this.NotesMasterFromJSON(oParsedNotes["master"], oPres) : oParsedNotes["master"];
		typeof oNotesMaster === "object" ? oNotes.setNotesMaster(oNotesMaster) : oNotes.setNotesMaster(this.notesMasterMap[oNotesMaster]);

		return oNotes;
	};
	ReaderFromJSON.prototype.NotesMasterFromJSON = function(oParsedNotesMaster, oPres)
	{
		var oNotesMaster = new AscCommonSlide.CNotesMaster();

		if (oParsedNotesMaster["clrMap"] != null)
			oNotesMaster.setClrMap(this.ColorMapOvrFromJSON(oParsedNotesMaster["clrMap"]));
	
		// cSld
		var oCSld = this.CSldFromJSON(oParsedNotesMaster["cSld"]);
		for (var nShape = 0; nShape < oCSld.spTree.length; nShape++)
			oNotesMaster.addToSpTreeToPos(nShape, oCSld.spTree[nShape]);
		oCSld.Bg && oNotesMaster.changeBackground(oCSld.Bg);
		oCSld.name && oNotesMaster.setCSldName(oCSld.name);

		oParsedNotesMaster["hf"] && oNotesMaster.setHF(this.HFFromJSON(oParsedNotesMaster["hf"]));
		oParsedNotesMaster["notesStyle"] && oNotesMaster.setNotesStyle(this.LstStyleFromJSON(oParsedNotesMaster["notesStyle"]));

		// oTheme здесь будет либо объект темы, либо Id по которому тема уже создана и замаплена
		var oTheme = typeof oParsedNotesMaster["theme"] === "object" ? this.ThemeFromJSON(oParsedNotesMaster["theme"]) : oParsedNotesMaster["theme"];
		typeof oParsedNotesMaster["theme"] === "object" ? oNotesMaster.setTheme(oTheme) : oNotesMaster.setTheme(this.themesMap[oTheme]);

		oPres.notesMasters[oPres.notesMasters.length] = oNotesMaster;
		this.notesMasterMap[oParsedNotesMaster["id"]] = oNotesMaster;

		return oNotesMaster;
	};
	ReaderFromJSON.prototype.TimingFromJSON = function(oParsedTimig)
	{
		var oTiming = new AscFormat.CTiming();

		oParsedTimig["bldLst"] && oTiming.setBldLst(this.BldLstFromJSON(oParsedTimig["bldLst"]));
		oParsedTimig["tnLst"] && oTiming.setTnLst(this.TnLstFromJSON(oParsedTimig["tnLst"]));

		return oTiming;
	};
	ReaderFromJSON.prototype.BldLstFromJSON = function(oParsedBldLst)
	{
		var oBldLst = new AscFormat.CBldLst();
		for (var nElm = 0; nElm < oParsedBldLst.length; nElm++)
		{
			if (oParsedBldLst[nElm]["type"] === "bldDgm")
				oBldLst.addToLst(oBldLst.list.length, this.BldDgmFromJSON(oParsedBldLst[nElm]));
			else if (oParsedBldLst[nElm]["type"] === "bldOleChart")
				oBldLst.addToLst(oBldLst.list.length, this.BldOleChartFromJSON(oParsedBldLst[nElm]));
			else if (oParsedBldLst[nElm]["type"] === "bldGraphic")
				oBldLst.addToLst(oBldLst.list.length, this.BldGraphicFromJSON(oParsedBldLst[nElm]));
			else if (oParsedBldLst[nElm]["type"] === "bldP")
				oBldLst.addToLst(oBldLst.list.length, this.BldPFromJSON(oParsedBldLst[nElm]));
		}

		return oBldLst;
	};
	ReaderFromJSON.prototype.BldDgmFromJSON = function(oParsedBldDgm)
	{
		var oBldDgm = new AscFormat.CBldDgm();

		var spId = this.drawingsMap[oParsedBldDgm["spid"]] ? this.drawingsMap[oParsedBldDgm["spid"]].Id : oParsedBldDgm["spid"];
		oParsedBldDgm["bld"] != undefined && oBldDgm.setBld(oParsedBldDgm["bld"]);
		oParsedBldDgm["grpId"] != undefined && oBldDgm.setGrpId(oParsedBldDgm["grpId"]);
		oParsedBldDgm["spid"] != undefined && oBldDgm.setSpid(spId);
		oParsedBldDgm["uiExpand"] != undefined && oBldDgm.setUiExpand(oParsedBldDgm["uiExpand"]);

		return oBldDgm;
	};
	ReaderFromJSON.prototype.BldOleChartFromJSON = function(oParsedBldOleChart)
	{
		var oBldOleChart = new AscFormat.CBldOleChart();

		var spId = this.drawingsMap[oParsedBldOleChart["spid"]] ? this.drawingsMap[oParsedBldOleChart["spid"]].Id : oParsedBldOleChart["spid"];

		oParsedBldOleChart["animBg"] != undefined && oBldOleChart.setAnimBg(oParsedBldOleChart["animBg"]);
		oParsedBldOleChart["bld"] != undefined && oBldOleChart.setBld(oParsedBldOleChart["bld"]);
		oParsedBldOleChart["grpId"] != undefined && oBldOleChart.setGrpId(oParsedBldOleChart["grpId"]);
		oParsedBldOleChart["spid"] != undefined && oBldOleChart.setSpid(spId);
		oParsedBldOleChart["uiExpand"] != undefined && oBldOleChart.setUiExpand(oParsedBldOleChart["uiExpand"]);

		return oBldOleChart;
	};
	ReaderFromJSON.prototype.BldGraphicFromJSON = function(oParsedBldGraphic)
	{
		var oBldGraphic = new AscFormat.CBldGraphic();

		var spId = this.drawingsMap[oParsedBldGraphic["spid"]] ? this.drawingsMap[oParsedBldGraphic["spid"]].Id : oParsedBldGraphic["spid"];

		oParsedBldGraphic["bldAsOne"] && oBldGraphic.setBldAsOne(new AscFormat.CEmptyObject());
		oParsedBldGraphic["bldSub"] && oBldGraphic.setBldSub(this.BldSubFromJSON(oParsedBldGraphic["bldSub"]));

		oParsedBldGraphic["grpId"] != undefined && oBldGraphic.setGrpId(oParsedBldGraphic["grpId"]);
		oParsedBldGraphic["spid"] != undefined && oBldGraphic.setSpid(spId);
		oParsedBldGraphic["uiExpand"] != undefined && oBldGraphic.setUiExpand(oParsedBldGraphic["uiExpand"]);

		return oBldGraphic;
	};
	ReaderFromJSON.prototype.BldSubFromJSON = function(oParsedBldSub)
	{
		var oBldSub = new AscFormat.CBldSub();

		oParsedBldSub["chart"] != undefined && oBldSub.setChart(oParsedBldSub["chart"]);
		oParsedBldSub["animBg"] != undefined && oBldSub.setAnimBg(oParsedBldSub["animBg"]);
		oParsedBldSub["bldChart"] != undefined && oBldSub.setBldChart(oParsedBldSub["bldChart"]);
		oParsedBldSub["bldDgm"] != undefined && oBldSub.setBldDgm(oParsedBldSub["bldDgm"]);
		oParsedBldSub["rev"] != undefined && oBldSub.setRev(oParsedBldSub["rev"]);

		return oBldSub;
	};
	ReaderFromJSON.prototype.BldPFromJSON = function(oParsedBldP)
	{
		var oBldP = new AscFormat.CBldP();
		
		var spId = this.drawingsMap[oParsedBldP["spid"]] ? this.drawingsMap[oParsedBldP["spid"]].Id : oParsedBldP["spid"];

		var nBuildType = undefined;
		switch (oParsedBldP["build"])
		{
			case "allAtOnce":
				nBuildType = c_oAscSlideParaBuildType.AllAtOnce;
				break;
			case "cust":
				nBuildType = c_oAscSlideParaBuildType.Cust;
				break;
			case "p":
				nBuildType = c_oAscSlideParaBuildType.P;
				break;
			case "whole":
				nBuildType = c_oAscSlideParaBuildType.Whole;
				break;
		}

		oParsedBldP["tmplLst"] != undefined && oBldP.setTmplLst(this.TmplLstFromJSON(oParsedBldP["tmplLst"]));
		oParsedBldP["advAuto"] != undefined && oBldP.setAdvAuto(oParsedBldP["advAuto"]);
		oParsedBldP["animBg"] != undefined && oBldP.setAnimBg(oParsedBldP["animBg"]);
		oParsedBldP["autoUpdateAnimB"] != undefined && oBldP.setAutoUpdateAnimBg(oParsedBldP["autoUpdateAnimB"]);
		oParsedBldP["bldLvl"] != undefined && oBldP.setBldLvl(oParsedBldP["bldLvl"]);
		nBuildType != undefined && oBldP.setBuild(nBuildType);
		oParsedBldP["grpId"] != undefined && oBldP.setGrpId(oParsedBldP["grpId"]);
		oParsedBldP["rev"] != undefined && oBldP.setRev(oParsedBldP["rev"]);
		oParsedBldP["spid"] != undefined && oBldP.setSpid(spId);
		oParsedBldP["uiExpand"] != undefined && oBldP.setUiExpand(oParsedBldP["uiExpand"]);

		return oBldP;
	};
	ReaderFromJSON.prototype.TmplLstFromJSON = function(oParsedTmplLst)
	{
		var oTmplLst = new AscFormat.CTmplLst();

		for (var nElm = 0; nElm < oParsedTmplLst["list"].length; nElm++)
			oTmplLst.addToLst(oTmplLst.list.length, this.TmplFromJSON(oParsedTmplLst["list"][nElm]));

		return oTmplLst;
	};
	ReaderFromJSON.prototype.TmplFromJSON = function(oParsedTmpl)
	{
		var oTmlp = new AscFormat.CTmpl();

		oParsedTmpl["lvl"] != undefined && oTmlp.setLvl(oParsedTmpl["lvl"]);
		oParsedTmpl["tnLst"] && oTmlp.setTnLst(this.TnLstFromJSON(oParsedTmpl["tnLst"]));
		
		return oTmlp;
	};
	ReaderFromJSON.prototype.TnLstFromJSON = function(oParsedTnLst, bChildLst)
	{
		var oTnLSt   = bChildLst ? new AscFormat.CChildTnLst() : new AscFormat.CTnLst();
		var oTempElm = null;
		for (var nElm = 0; nElm < oParsedTnLst.length; nElm++)
		{
			oTempElm = oParsedTnLst[nElm];
			switch (oTempElm["objType"])
			{
				case "par":
					oTnLSt.addToLst(oTnLSt.list.length, this.ParFromJSON(oTempElm));
					break;
				case "seq":
					oTnLSt.addToLst(oTnLSt.list.length, this.SeqFromJSON(oTempElm));
					break;
				case "audio":
					oTnLSt.addToLst(oTnLSt.list.length, this.AudioFromJSON(oTempElm));
					break;
				case "video":
					oTnLSt.addToLst(oTnLSt.list.length, this.VideoFromJSON(oTempElm));
					break;
				case "excl":
					oTnLSt.addToLst(oTnLSt.list.length, this.ExclFromJSON(oTempElm));
					break;
				case "anim":
					oTnLSt.addToLst(oTnLSt.list.length, this.AnimFromJSON(oTempElm));
					break;
				case "animClr":
					oTnLSt.addToLst(oTnLSt.list.length, this.AnimClrFromJSON(oTempElm));
					break;
				case "animEffect":
					oTnLSt.addToLst(oTnLSt.list.length, this.AnimEffectFromJSON(oTempElm));
					break;
				case "animMotion":
					oTnLSt.addToLst(oTnLSt.list.length, this.AnimMotionFromJSON(oTempElm));
					break;
				case "animRot":
					oTnLSt.addToLst(oTnLSt.list.length, this.AnimRotFromJSON(oTempElm));
					break;
				case "animScale":
					oTnLSt.addToLst(oTnLSt.list.length, this.AnimScaleFromJSON(oTempElm));
					break;
				case "cmd":
					oTnLSt.addToLst(oTnLSt.list.length, this.CmdFromJSON(oTempElm));
					break;
				case "set":
					oTnLSt.addToLst(oTnLSt.list.length, this.SetFromJSON(oTempElm));
					break;
			}
		}
			
		return oTnLSt;
	};
	ReaderFromJSON.prototype.ParFromJSON = function(oParsedPar)
	{
		var oPar = new AscFormat.CPar();

		oParsedPar["cTn"] && oPar.setCTn(this.CTnFromJSON(oParsedPar["cTn"]));

		return oPar;
	};
	ReaderFromJSON.prototype.CTnFromJSON = function(oParsedCTn)
	{
		var oCTn = new AscFormat.CCTn();

		oParsedCTn["childTnLst"] && oCTn.setChildTnLst(this.TnLstFromJSON(oParsedCTn["childTnLst"], true));
		oParsedCTn["endCondLst"] && oCTn.setEndCondLst(this.CondLstFromJSON(oParsedCTn["endCondLst"]));
		oParsedCTn["endSync"]    && oCTn.setEndSync(this.CondFromJSON(oParsedCTn["endSync"]));
		oParsedCTn["iterate"]    && oCTn.setIterate(this.IterateFromJSON(oParsedCTn["iterate"]));
		oParsedCTn["stCondLst"]  && oCTn.setStCondLst(this.CondLstFromJSON(oParsedCTn["stCondLst"]));
		oParsedCTn["subTnLst"]   && oCTn.setSubTnLst(this.TnLstFromJSON(oParsedCTn["subTnLst"]));

		var nNodeFillType = undefined;
		switch (oParsedCTn["fill"])
		{
			case "freeze":
				nNodeFillType = c_oAscSlideNodeFillType.Freeze;
				break;
			case "hold":
				nNodeFillType = c_oAscSlideNodeFillType.Hold;
				break;
			case "remove":
				nNodeFillType = c_oAscSlideNodeFillType.Remove;
				break;
			case "transition":
				nNodeFillType = c_oAscSlideNodeFillType.Transition;
				break;
		}

		var nMasterRelType = undefined;
		switch (oParsedCTn["masterRel"])
		{
			case "lastClick":
				nMasterRelType = c_oAscSlideMasterRelationType.LastClick;
				break;
			case "nextClick":
				nMasterRelType = c_oAscSlideMasterRelationType.NextClick;
				break;
			case "sameClick":
				nMasterRelType = c_oAscSlideMasterRelationType.SameClick;
				break;
		}

		var nNodeType = undefined;
		switch (oParsedCTn["nodeType"])
		{
			case "afterEffect":
				nNodeType = c_oAscSlideNodeType.AfterEffect;
				break;
			case "afterGroup":
				nNodeType = c_oAscSlideNodeType.AfterGroup;
				break;
			case "clickEffect":
				nNodeType = c_oAscSlideNodeType.ClickEffect;
				break;
			case "clickPar":
				nNodeType = c_oAscSlideNodeType.ClickPar;
				break;
			case "interactiveSeq":
				nNodeType = c_oAscSlideNodeType.InteractiveSeq;
				break;
			case "mainSeq":
				nNodeType = c_oAscSlideNodeType.MainSeq;
				break;
			case "tmRoot":
				nNodeType = c_oAscSlideNodeType.TmRoot;
				break;
			case "withEffect":
				nNodeType = c_oAscSlideNodeType.WithEffect;
				break;
			case "withGroup":
				nNodeType = c_oAscSlideNodeType.WithGroup;
				break;
		}

		var nPresetClassType = undefined;
		switch (oParsedCTn["presetClass"])
		{
			case "emph":
				nPresetClassType = c_oAscSlidePresetClassType.Emph;
				break;
			case "entr":
				nPresetClassType = c_oAscSlidePresetClassType.Entr;
				break;
			case "exit":
				nPresetClassType = c_oAscSlidePresetClassType.Exit;
				break;
			case "mediacall":
				nPresetClassType = c_oAscSlidePresetClassType.Mediacall;
				break;
			case "path":
				nPresetClassType = c_oAscSlidePresetClassType.Path;
				break;
			case "verb":
				nPresetClassType = c_oAscSlidePresetClassType.Verb;
				break;
		}

		var nRestartType = undefined;
		switch (oParsedCTn["restart"])
		{
			case "always":
				nRestartType = c_oAscSlideRestartType.Always;
				break;
			case "never":
				nRestartType = c_oAscSlideRestartType.Never;
				break;
			case "whenNotActive":
				nRestartType = c_oAscSlideRestartType.WhenNotActive;
				break;
		}

		var nSyncBehaviorType = undefined;
		switch (oParsedCTn["syncBehavior"])
		{
			case "canSlip":
				nSyncBehaviorType = c_oAscSlideSyncBehaviorType.CanSlip;
				break;
			case "locked":
				nSyncBehaviorType = c_oAscSlideSyncBehaviorType.Locked;
				break;
		}

		oParsedCTn["accel"] != undefined && oCTn.setAccel(oParsedCTn["accel"]);
		oParsedCTn["afterEffect"] != undefined && oCTn.setAfterEffect(oParsedCTn["afterEffect"]);
		oParsedCTn["autoRev"] != undefined && oCTn.setAutoRev(oParsedCTn["autoRev"]);
		oParsedCTn["bldLvl"] != undefined && oCTn.setBldLvl(oParsedCTn["bldLvl"]);
		oParsedCTn["decel"] != undefined && oCTn.setDecel(oParsedCTn["decel"]);
		oParsedCTn["display"] != undefined && oCTn.setDisplay(oParsedCTn["display"]);
		oParsedCTn["dur"] != undefined && oCTn.setDur(oParsedCTn["dur"]);
		oParsedCTn["evtFilter"] != undefined && oCTn.setEvtFilter(oParsedCTn["evtFilter"]);
		nNodeFillType != undefined && oCTn.setFill(nNodeFillType);
		oParsedCTn["grpId"] != undefined && oCTn.setGrpId(oParsedCTn["grpId"]);
		oParsedCTn["id"] != undefined && oCTn.setId(oParsedCTn["id"]);
		nMasterRelType != undefined && oCTn.setMasterRel(nMasterRelType);
		oParsedCTn["nodePh"] != undefined && oCTn.setNodePh(oParsedCTn["nodePh"]);
		nNodeType != undefined && oCTn.setNodeType(nNodeType);
		nPresetClassType != undefined && oCTn.setPresetClass(nPresetClassType);
		oParsedCTn["presetID"] != undefined && oCTn.setPresetID(oParsedCTn["presetID"]);
		oParsedCTn["presetSubtype"] != undefined && oCTn.setPresetSubtype(oParsedCTn["presetSubtype"]);
		oParsedCTn["repeatCount"] != undefined && oCTn.setRepeatCount(oParsedCTn["repeatCount"]);
		oParsedCTn["repeatDur"] != undefined && oCTn.setRepeatDur(oParsedCTn["repeatDur"]);
		nRestartType != undefined && oCTn.setRestart(nRestartType);
		oParsedCTn["spd"] != undefined && oCTn.setSpd(oParsedCTn["spd"]);
		nSyncBehaviorType != undefined && oCTn.setSyncBehavior(nSyncBehaviorType);
		oParsedCTn["tmFilter"] != undefined && oCTn.setTmFilter(oParsedCTn["tmFilter"]);

		return oCTn;
	};
	ReaderFromJSON.prototype.CondLstFromJSON = function(oParsedCondLst)
	{
		var oCondLst = new AscFormat.CCondLst();

		for (var nCond = 0; nCond < oParsedCondLst.length; nCond++)
			oCondLst.addToLst(oCondLst.list.length, this.CondFromJSON(oParsedCondLst[nCond]));

		return oCondLst;
	};
	ReaderFromJSON.prototype.CondFromJSON = function(oParsedCond)
	{
		var oCond = new AscFormat.CCond();

		var nEventType = undefined;
		switch (oParsedCond["evt"])
		{
			case "begin":
				nEventType = c_oAscSlideTriggerEventType.Begin;
				break;
			case "end":
				nEventType = c_oAscSlideTriggerEventType.End;
				break;
			case "onBegin":
				nEventType = c_oAscSlideTriggerEventType.OnBegin;
				break;
			case "onClick":
				nEventType = c_oAscSlideTriggerEventType.OnClick;
				break;
			case "onDblClick":
				nEventType = c_oAscSlideTriggerEventType.OnDblClick;
				break;
			case "onEnd":
				nEventType = c_oAscSlideTriggerEventType.OnEnd;
				break;
			case "onMouseOut":
				nEventType = c_oAscSlideTriggerEventType.OnMouseOut;
				break;
			case "onMouseOver":
				nEventType = c_oAscSlideTriggerEventType.OnMouseOver;
				break;
			case "onNext":
				nEventType = c_oAscSlideTriggerEventType.OnNext;
				break;
			case "onPrev":
				nEventType = c_oAscSlideTriggerEventType.OnPrev;
				break;
			case "onStopAudio":
				nEventType = c_oAscSlideTriggerEventType.OnStopAudio;
				break;
		}

		oParsedCond["rtn"] && oCond.setRtn(this.RtnFromJSON(oParsedCond["rtn"]));
		oParsedCond["tgtEl"] && oCond.setTgtEl(this.TgtElFromJSON(oParsedCond["tgtEl"]));
		oParsedCond["tn"] != undefined && oCond.setTn(oParsedCond["tn"]);
		oParsedCond["delay"] != undefined && oCond.setDelay(oParsedCond["delay"]);
		nEventType != undefined && oCond.setEvt(nEventType);

		return oCond;
	};
	ReaderFromJSON.prototype.RtnFromJSON = function(oParsedRtn)
	{

		var nType = undefined;
		switch (oParsedRtn["val"])
		{
			case "all":
				nType = c_oAscSlideRuntimeTriggerType.All;
				break;
			case "first":
				nType = c_oAscSlideRuntimeTriggerType.First;
				break;
			case "last":
				nType = c_oAscSlideRuntimeTriggerType.Last;
				break;
		}

		return nType || null;
	};
	ReaderFromJSON.prototype.TgtElFromJSON = function(oParsedTgtEl)
	{
		var oTgtEl = new AscFormat.CTgtEl();

		oParsedTgtEl["inkTgt"] && oTgtEl.setInkTgt(this.InkTgtFromJSON(oParsedTgtEl["inkTgt"]));
		//oParsedTgtEl["sldTgt"] && oTgtEl.setSldTgt(this.SldTgtFromJSON(oParsedTgtEl["sldTgt"]));
		oParsedTgtEl["sndTgt"] && oTgtEl.setSndTgt(this.SndTgtFromJSON(oParsedTgtEl["sndTgt"]));
		oParsedTgtEl["spTgt"] && oTgtEl.setSpTgt(this.SpTgtFromJSON(oParsedTgtEl["spTgt"]));
		
		return oTgtEl;
	};
	ReaderFromJSON.prototype.InkTgtFromJSON = function(oParsedInkTgt)
	{
		var oInkTgt = new AscFormat.CObjectTarget();

		//var spId = this.drawingsMap[oParsedInkTgt["spid"]] ? this.drawingsMap[oParsedInkTgt["spid"]].Id : oParsedInkTgt["spid"];

		oParsedInkTgt["spid"] != undefined && oInkTgt.setSpid(oParsedInkTgt["spid"]);
		this.AddConnectedObject(oInkTgt);

		return oInkTgt;
	};
	ReaderFromJSON.prototype.SldTgtFromJSON = function(oParsedSldTgt)
	{
		// not supported yet
	};
	ReaderFromJSON.prototype.SndTgtFromJSON = function(oParsedSndTgt)
	{
		var oSndTgt = new AscFormat.CSndTgt();

		oParsedSndTgt["embed"] != undefined && oSndTgt.setEmbed(oParsedSndTgt["embed"]);
		oParsedSndTgt["name"] != undefined && oSndTgt.setName(oParsedSndTgt["name"]);
		oParsedSndTgt["builtIn"] != undefined && oSndTgt.setBuiltIn(oParsedSndTgt["builtIn"]);

		return oSndTgt;
	};
	ReaderFromJSON.prototype.SpTgtFromJSON = function(oParsedSpTgt)
	{
		var oSpTgt = new AscFormat.CSpTgt();

		//var spId = this.drawingsMap[oParsedSpTgt["spid"]] ? this.drawingsMap[oParsedSpTgt["spid"]].Id : oParsedSpTgt["spid"];
		//var subSpId = this.drawingsMap[oParsedSpTgt["subSpId"]] ? this.drawingsMap[oParsedSpTgt["subSpId"]].Id : oParsedSpTgt["subSpId"];

		oParsedSpTgt["spid"] != undefined && oSpTgt.setSpid(oParsedSpTgt["spid"]);
		oParsedSpTgt["bg"] != undefined && oSpTgt.setBg(oParsedSpTgt["bg"]);
		oParsedSpTgt["subSpId"] != undefined && oSpTgt.setSubSpId(oParsedSpTgt["subSpId"]);
		
		oParsedSpTgt["oleChartEl"] && oSpTgt.setOleChartEl(this.OleChartElFromJSON(oParsedSpTgt["oleChartEl"]));
		oParsedSpTgt["txEl"] && oSpTgt.setTxEl(this.TxElFromJSON(oParsedSpTgt["txEl"]));
		oParsedSpTgt["graphicEl"] && oSpTgt.setGraphicEl(this.GraphicElFromJSON(oParsedSpTgt["graphicEl"]));
		this.AddConnectedObject(oSpTgt);
		return oSpTgt;
	};
	ReaderFromJSON.prototype.OleChartElFromJSON = function(oParsedOleChartEl)
	{
		var oOleChartEl = new AscFormat.COleChartEl();

		var nType = undefined;
		switch (oParsedOleChartEl["type"])
		{
			case "category":
				nType = c_oAscSlideChartSubElementType.Category;
				break;
			case "gridLegend":
				nType = c_oAscSlideChartSubElementType.GridLegend;
				break;
			case "ptInCategory":
				nType = c_oAscSlideChartSubElementType.PtInCategory;
				break;
			case "ptInSeries":
				nType = c_oAscSlideChartSubElementType.PtInSeries;
				break;
			case "series":
				nType = c_oAscSlideChartSubElementType.Series;
				break;
		}

		oParsedOleChartEl["lvl"] != undefined && oOleChartEl.setLvl(oParsedOleChartEl["lvl"]);
		nType != undefined && oOleChartEl.setType(nType);

		return oOleChartEl;
	};
	ReaderFromJSON.prototype.TxElFromJSON = function(oParsedTxEl)
	{
		var oTxEl = new AscFormat.CTxEl();

		oParsedTxEl["charRg"] && oTxEl.setCharRg(this.IndexRgFromJSON(oParsedTxEl["charRg"]));
		oParsedTxEl["pRg"] && oTxEl.setPRg(this.IndexRgFromJSON(oParsedTxEl["pRg"]));

		return oTxEl;
	};
	ReaderFromJSON.prototype.IndexRgFromJSON = function(oParsedIndexRg)
	{
		var oIndexRg = new AscFormat.CIndexRg();

		oParsedIndexRg["st"] != undefined && oIndexRg.setSt(oParsedIndexRg["st"]);
		oParsedIndexRg["end"] != undefined && oIndexRg.setEnd(oParsedIndexRg["end"]);

		return oIndexRg;
	};
	ReaderFromJSON.prototype.GraphicElFromJSON = function(oParsedGraphicEl)
	{
		var oGraphicEl = new AscFormat.CGraphicEl();

		oParsedGraphicEl["chart"]["bldStep"] != undefined && oGraphicEl.setChartBuildStep(oParsedGraphicEl["chart"]["bldStep"]);
		oParsedGraphicEl["chart"]["categoryIdx"] != undefined && oGraphicEl.setCategoryIdx(oParsedGraphicEl["chart"]["categoryIdx"]);
		oParsedGraphicEl["chart"]["seriesIdx"] != undefined && oGraphicEl.setSeriesIdx(oParsedGraphicEl["chart"]["seriesIdx"]);

		oParsedGraphicEl["dgm"]["bldStep"] != undefined && oGraphicEl.setDgmBuildStep(oParsedGraphicEl["dgm"]["bldStep"]);
		oParsedGraphicEl["dgm"]["id"] != undefined && oGraphicEl.setDgmId(oParsedGraphicEl["dgm"]["id"]);
		this.AddConnectedObject(oGraphicEl);

		return oGraphicEl;
	};
	ReaderFromJSON.prototype.IterateFromJSON = function(oParsedIterate)
	{
		var oIterate = new AscFormat.CIterateData();

		var nIterateType = undefined;
		switch (oParsedIterate["type"])
		{
			case "el":
				nIterateType = c_oAscSlideIterateType.El;
				break;
			case "lt":
				nIterateType = c_oAscSlideIterateType.Lt;
				break;
			case "wd":
				nIterateType = c_oAscSlideIterateType.Wd;
				break;
		}

		oParsedIterate["tmAbs"] != undefined && oIterate.setTmAbs(oParsedIterate["tmAbs"]);
		oParsedIterate["tmPct"] != undefined && oIterate.setTmPct(oParsedIterate["tmPct"]);
		oParsedIterate["backwards"] != undefined && oIterate.setBackwards(oParsedIterate["backwards"]);
		nIterateType != undefined && oIterate.setType(nIterateType);

		return oIterate;
	};
	ReaderFromJSON.prototype.SeqFromJSON = function(oParsedSeq)
	{
		var oSeq = new AscFormat.CSeq();
		
		var nNextAcType = undefined;
		switch (oParsedSeq["nextAc"])
		{
			case "none":
				nNextAcType = c_oAscSlideNextAcType.None;
				break;
			case "seek":
				nNextAcType = c_oAscSlideNextAcType.Seek;
				break;
		}

		var nPrevAcType = undefined;
		switch (oParsedSeq["prevAc"])
		{
			case "none":
				nPrevAcType = c_oAscSlidePrevAcType.None;
				break;
			case "skipTimed":
				nPrevAcType = c_oAscSlidePrevAcType.SkipTimed;
				break;
		}

		oParsedSeq["cTn"] && oSeq.setCTn(this.CTnFromJSON(oParsedSeq["cTn"]));
		oParsedSeq["nextCondLst"] && oSeq.setNextCondLst(this.CondLstFromJSON(oParsedSeq["nextCondLst"]));
		oParsedSeq["prevCondLst"] && oSeq.setPrevCondLst(this.CondLstFromJSON(oParsedSeq["prevCondLst"]));

		oParsedSeq["concurrent"] != undefined && oSeq.setConcurrent(oParsedSeq["concurrent"]);
		nNextAcType != undefined && oSeq.setNextAc(nNextAcType);
		nPrevAcType != undefined && oSeq.setPrevAc(nPrevAcType);

		return oSeq;
	};
	ReaderFromJSON.prototype.AudioFromJSON = function(oParsedAudio)
	{
		var oAudio = new AscFormat.CAudio();

		oParsedAudio["cMediaNode"] && oAudio.setCMediaNode(this.CMediaNodeFromJSON(oParsedAudio["cMediaNode"]));
		oParsedAudio["isNarration"] != undefined && oAudio.setIsNarration(oParsedAudio["isNarration"]);

		return oAudio;
	};
	ReaderFromJSON.prototype.CMediaNodeFromJSON = function(oParsedCMediaNode)
	{
		var oCMediaNode = new AscFormat.CCMediaNode();

		oParsedCMediaNode["cTn"] && oCMediaNode.setCTn(this.CTnFromJSON(oParsedCMediaNode["cTn"]));
		oParsedCMediaNode["tgtEl"] && oCMediaNode.setTgtEl(this.TgtElFromJSON(oParsedCMediaNode["tgtEl"]));

		oParsedCMediaNode["mute"] != undefined && oCMediaNode.setMute(oParsedCMediaNode["mute"]);
		oParsedCMediaNode["numSld"] != undefined && oCMediaNode.setNumSld(oParsedCMediaNode["numSld"]);
		oParsedCMediaNode["showWhenStopped"] != undefined && oCMediaNode.setShowWhenStopped(oParsedCMediaNode["showWhenStopped"]);
		oParsedCMediaNode["vol"] != undefined && oCMediaNode.setVol(oParsedCMediaNode["vol"]);

		return oCMediaNode;
	};
	ReaderFromJSON.prototype.VideoFromJSON = function(oParsedVideo)
	{
		var oVideo = new AscFormat.CVideo();
		
		oParsedVideo["cMediaNode"] && oVideo.setCMediaNode(this.CMediaNodeFromJSON(oParsedVideo["cMediaNode"]));
		oParsedVideo["fullScrn"] != undefined && oVideo.setFullScrn(oParsedVideo["fullScrn"]);

		return oVideo;
	};
	ReaderFromJSON.prototype.ExclFromJSON = function(oParsedExcl)
	{
		var oExcl = new AscFormat.CExcl();

		oParsedExcl["cTn"] && oExcl.setCTn(this.CTnFromJSON(oParsedExcl["cTn"]));

		return oExcl;
	};
	ReaderFromJSON.prototype.AnimFromJSON = function(oParsedAnim)
	{
		var oAnim = new AscFormat.CAnim();

		var nCalcmodeType = undefined;
		switch (oParsedAnim["calcmode"])
		{
			case "discrete":
				nCalcmodeType = c_oAscSlideCalcModeType.Discrete;
				break;
			case "lin":
				nCalcmodeType = c_oAscSlideCalcModeType.Lin;
				break;
			case "fmla":
				nCalcmodeType = c_oAscSlideCalcModeType.Fmla;
				break;
		}

		var nValueType = undefined;
		switch (oParsedAnim["valueType"])
		{
			case "num":
				nValueType = c_oAscSlideTLValueType.Num;
				break;
			case "clr":
				nValueType = c_oAscSlideTLValueType.Clr;
				break;
			case "str":
				nValueType = c_oAscSlideTLValueType.Str;
				break;
		}

		oParsedAnim["cBhvr"] && oAnim.setCBhvr(this.CBhvrFromJSON(oParsedAnim["cBhvr"]));
		oParsedAnim["tavLst"] && oAnim.setTavLst(this.TavLstFromJSON(oParsedAnim["tavLst"]));

		oParsedAnim["by"] != undefined && oAnim.setBy(oParsedAnim["by"]);
		nCalcmodeType != undefined && oAnim.setCalcmode(nCalcmodeType);
		oParsedAnim["from"] != undefined && oAnim.setFrom(oParsedAnim["from"]);
		oParsedAnim["to"] != undefined && oAnim.setTo(oParsedAnim["to"]);
		nValueType != undefined && oAnim.setValueType(nValueType);

		return oAnim;
	};
	ReaderFromJSON.prototype.CBhvrFromJSON = function(oParsedCBhvr)
	{
		var oCBhvr = new AscFormat.CCBhvr();

		var nAccumulateType = undefined;
		switch (oParsedCBhvr["accumulate"])
		{
			case "always":
				nAccumulateType = c_oAscSlideTLAccumulateType.Always;
				break;
			case "none":
				nAccumulateType = c_oAscSlideTLAccumulateType.None;
				break;
		}

		var nAdditiveType = undefined;
		switch (oParsedCBhvr["additive"])
		{
			case "base":
				nAdditiveType = c_oAscSlideTLAdditiveType.Base;
				break;
			case "mult":
				nAdditiveType = c_oAscSlideTLAdditiveType.Mult;
				break;
			case "none":
				nAdditiveType = c_oAscSlideTLAdditiveType.None;
				break;
			case "repl":
				nAdditiveType = c_oAscSlideTLAdditiveType.Repl;
				break;
			case "sum":
				nAdditiveType = c_oAscSlideTLAdditiveType.Sum;
				break;
		}

		var nOverrideType = undefined;
		switch (oParsedCBhvr["override"])
		{
			case "childStyle":
				nOverrideType = c_oAscSlideTLOverrideType.ChildStyle;
				break;
			case "normal":
				nOverrideType = c_oAscSlideTLOverrideType.Normal;
				break;
		}

		var nXfrmTypeType = undefined;
		switch (oParsedCBhvr["xfrmType"])
		{
			case "img":
				nXfrmTypeType = c_oAscSlideTLTransformType.Img;
				break;
			case "pt":
				nXfrmTypeType = c_oAscSlideTLTransformType.Pt;
				break;
		}

		oParsedCBhvr["attrNameLst"] && oCBhvr.setAttrNameLst(this.AttrNameLstFromJSON(oParsedCBhvr["attrNameLst"]));
		oParsedCBhvr["cTn"] && oCBhvr.setCTn(this.CTnFromJSON(oParsedCBhvr["cTn"]));
		oParsedCBhvr["tgtEl"] && oCBhvr.setTgtEl(this.TgtElFromJSON(oParsedCBhvr["tgtEl"]));

		nAccumulateType != undefined && oCBhvr.setAccumulate(nAccumulateType);
		nAdditiveType != undefined && oCBhvr.setAdditive(nAdditiveType);
		oParsedCBhvr["by"] != undefined && oCBhvr.setBy(oParsedCBhvr["by"]);
		oParsedCBhvr["from"] != undefined && oCBhvr.setFrom(oParsedCBhvr["from"]);
		nOverrideType != undefined && oCBhvr.setOverride(nOverrideType);
		oParsedCBhvr["rctx"] != undefined && oCBhvr.setRctx(oParsedCBhvr["rctx"]);
		oParsedCBhvr["to"] != undefined && oCBhvr.setTo(oParsedCBhvr["to"]);
		nXfrmTypeType != undefined && oCBhvr.setXfrmType(nXfrmTypeType);

		return oCBhvr;
	};
	ReaderFromJSON.prototype.AttrNameLstFromJSON = function(oParsedAttrNameLst)
	{
		var oAttrNameLst = new AscFormat.CAttrNameLst();
		var oTempName    = null;
		for (var nName = 0; nName < oParsedAttrNameLst.length; nName++)
		{
			oTempName = new AscFormat.CAttrName();
			oTempName.setText(oParsedAttrNameLst[nName]);
			oAttrNameLst.addToLst(oAttrNameLst.list.length, oTempName);
		}
			
		return oAttrNameLst;
	};
	ReaderFromJSON.prototype.TavLstFromJSON = function(oParsedTavLst)
	{
		var oTavLst = new AscFormat.CTavLst();
		for (var nTav = 0; nTav < oParsedTavLst.length; nTav++)
			oTavLst.addToLst(oTavLst.list.length, this.TavFromJSON(oParsedTavLst[nTav]));

		return oTavLst;
	};
	ReaderFromJSON.prototype.TavFromJSON = function(oParsedTav)
	{
		var oTav = new AscFormat.CTav();

		oParsedTav["val"] && oTav.setVal(this.AnimVariantFromJSON(oParsedTav["val"]));
		oParsedTav["fmla"] != undefined && oTav.setFmla(oParsedTav["fmla"]);
		oParsedTav["tm"] != undefined && oTav.setTm(oParsedTav["tm"]);

		return oTav;
	};
	ReaderFromJSON.prototype.AnimVariantFromJSON = function(oParsedAnimVariant)
	{
		var oAnimVariant = new AscFormat.CAnimVariant();

		oParsedAnimVariant["boolVal"] != undefined && oAnimVariant.setBoolVal(oParsedAnimVariant["boolVal"]);
		oParsedAnimVariant["fltVal"] != undefined && oAnimVariant.setFltVal(oParsedAnimVariant["fltVal"]);
		oParsedAnimVariant["intVal"] != undefined && oAnimVariant.setIntVal(oParsedAnimVariant["intVal"]);
		oParsedAnimVariant["strVal"] != undefined && oAnimVariant.setStrVal(oParsedAnimVariant["strVal"]);
		oParsedAnimVariant["clrVal"] && oAnimVariant.setClrVal(this.ColorFromJSON(oParsedAnimVariant["clrVal"]));

		return oAnimVariant;
	};
	ReaderFromJSON.prototype.AnimClrFromJSON = function(oParsedAnimClr)
	{
		var oAnimClr = new AscFormat.CAnimClr();

		var nClrSpcType = undefined;
		switch (oParsedAnimClr["clrSpc"])
		{
			case "rgb":
				nClrSpcType = c_oAscSlideTLColorSpaceType.Rgb;
				break;
			case "hsl":
				nClrSpcType = c_oAscSlideTLColorSpaceType.Hsl;
				break;
		}

		var nColorDir = undefined;
		switch (oParsedAnimClr["dir"])
		{
			case "ccw":
				nColorDir = AscCommon.ST_HueDir.Ccw;
				break;
			case "cw":
				nColorDir = AscCommon.ST_HueDir.Cw;
				break;
		}

		oParsedAnimClr["by"]["rgb"] && oAnimClr.setByRGB(this.ByRGBFromJSON(oParsedAnimClr["by"]["rgb"]));
		oParsedAnimClr["by"]["hsl"] && oAnimClr.setByHSL(this.ByHSLFromJSON(oParsedAnimClr["by"]["hsl"]));

		oParsedAnimClr["cBhvr"] && oAnimClr.setCBhvr(this.CBhvrFromJSON(oParsedAnimClr["cBhvr"]));
		oParsedAnimClr["from"] != undefined && oAnimClr.setFrom(oParsedAnimClr["from"]);
		oParsedAnimClr["to"] != undefined && oAnimClr.setTo(oParsedAnimClr["to"]);
		nClrSpcType != undefined && oAnimClr.setClrSpc(nClrSpcType);
		nColorDir != undefined && oAnimClr.setDir(nColorDir);

		return oAnimClr;
	};
	ReaderFromJSON.prototype.ByRGBFromJSON = function(oParsedByRGB)
	{
		var oByRGB = new AscFormat.CColorPercentage();

		oByRGB.c1 = oParsedByRGB["r"];
		oByRGB.c1 = oParsedByRGB["g"];
		oByRGB.c1 = oParsedByRGB["b"];

		return oByRGB;
	};
	ReaderFromJSON.prototype.ByHSLFromJSON = function(oParsedByHSL)
	{
		var oByHSL = new AscFormat.CColorPercentage();

		oByHSL.c1 = oParsedByHSL["h"];
		oByHSL.c1 = oParsedByHSL["s"];
		oByHSL.c1 = oParsedByHSL["l"];

		return oByHSL;
	};
	ReaderFromJSON.prototype.AnimEffectFromJSON = function(oParsedAnimEffect)
	{
		var oAnimEffect = new AscFormat.CAnimEffect();

		oParsedAnimEffect["cBhvr"] && oAnimEffect.setCBhvr(this.CBhvrFromJSON(oParsedAnimEffect["cBhvr"]));
		oParsedAnimEffect["progress"] && oAnimEffect.setProgress(this.AnimVariantFromJSON(oParsedAnimEffect["progress"]));
		oParsedAnimEffect["filter"] != undefined && oAnimEffect.setFilter(oParsedAnimEffect["filter"]);
		oParsedAnimEffect["prLst"] != undefined && oAnimEffect.setPrLst(oParsedAnimEffect["prLst"]);
		oParsedAnimEffect["transition"] && oAnimEffect.setTransition(From_XML_TLAnimateEffectTransition(oParsedAnimEffect["transition"]));

		return oAnimEffect;
	};
	ReaderFromJSON.prototype.TransitionFromJSON = function(oParsedTransition)
	{
		var oTransition = new Asc.CAscSlideTransition();

		var nTransType = undefined;
		switch (oParsedTransition["type"])
		{
			case "none":
				nTransType = c_oAscSlideTransitionTypes.None;
				break;
			case "fade":
				nTransType = c_oAscSlideTransitionTypes.Fade;
				break;
			case "push":
				nTransType = c_oAscSlideTransitionTypes.Push;
				break;
			case "wipe":
				nTransType = c_oAscSlideTransitionTypes.Wipe;
				break;
			case "split":
				nTransType = c_oAscSlideTransitionTypes.Split;
				break;
			case "unCover":
				nTransType = c_oAscSlideTransitionTypes.UnCover;
				break;
			case "cover":
				nTransType = c_oAscSlideTransitionTypes.Cover;
				break;
			case "clock":
				nTransType = c_oAscSlideTransitionTypes.Clock;
				break;
			case "zoom":
				nTransType = c_oAscSlideTransitionTypes.Zoom;
				break;
		}

		var transOption = undefined;
		switch (nTransType)
		{
			case c_oAscSlideTransitionTypes.Fade:
				transOption = oParsedTransition["option"] === false ? c_oAscSlideTransitionParams.Fade_Smoothly  : c_oAscSlideTransitionParams.Fade_Through_Black;
				break;
			case c_oAscSlideTransitionTypes.Push:
			case c_oAscSlideTransitionTypes.Wipe:
			case c_oAscSlideTransitionTypes.Cover:
			case c_oAscSlideTransitionTypes.UnCover:
				switch (oParsedTransition["option"])
				{
					case "l":
						transOption = c_oAscSlideTransitionParams.Param_Left;
						break;
					case "t":
						transOption = c_oAscSlideTransitionParams.Param_Top;
						break;
					case "r":
						transOption = c_oAscSlideTransitionParams.Param_Right;
						break;
					case "b":
						transOption = c_oAscSlideTransitionParams.Param_Bottom;
						break;
					case "tl":
						transOption = c_oAscSlideTransitionParams.Param_TopLeft;
						break;
					case "tr":
						transOption = c_oAscSlideTransitionParams.Param_TopRight;
						break;
					case "bl":
						transOption = c_oAscSlideTransitionParams.Param_BottomLeft;
						break;
					case "br":
						transOption = c_oAscSlideTransitionParams.Param_BottomRight;
						break;
				}
				break;
			case c_oAscSlideTransitionTypes.Split:
				switch (oParsedTransition["option"])
				{
					case "verIn":
						transOption = c_oAscSlideTransitionParams.Split_VerticalIn;
						break;
					case "verOut":
						transOption = c_oAscSlideTransitionParams.Split_VerticalOut;
						break;
					case "horIn":
						transOption = c_oAscSlideTransitionParams.Split_HorizontalIn;
						break;
					case "horOut":
						transOption = c_oAscSlideTransitionParams.Split_HorizontalOut;
						break;
				}
				break;
			case c_oAscSlideTransitionTypes.Clock:
				switch (oParsedTransition["option"])
				{
					case "clockwise":
						transOption = c_oAscSlideTransitionParams.Clock_Clockwise;
						break;
					case "counterClockwise":
						transOption = c_oAscSlideTransitionParams.Clock_Counterclockwise;
						break;
					case "wedge":
						transOption = c_oAscSlideTransitionParams.Clock_Wedge;
						break;
				}
				break;
			case c_oAscSlideTransitionTypes.Zoom:
				switch (oParsedTransition["option"])
				{
					case "in":
						transOption = c_oAscSlideTransitionParams.Zoom_In;
						break;
					case "out":
						transOption = c_oAscSlideTransitionParams.Zoom_Out;
						break;
					case "andRotate":
						transOption = c_oAscSlideTransitionParams.Zoom_AndRotate;
						break;
				}
				break;
		}

		oTransition.TransitionType           = nTransType;
		oTransition.TransitionOption         = transOption;
		oTransition.TransitionDuration       = oParsedTransition["transDur"];
		oTransition.SlideAdvanceOnMouseClick = oParsedTransition["advClick"];
		oTransition.SlideAdvanceAfter        = oParsedTransition["advAfter"];
		oTransition.SlideAdvanceDuration     = oParsedTransition["advDur"];
		oTransition.ShowLoop                 = oParsedTransition["shwLoop"];

		return oTransition;
	};
	ReaderFromJSON.prototype.AnimMotionFromJSON = function(oParsedAnimMotion)
	{
		var oAnimMotion = new AscFormat.CAnimMotion();

		var nOriginType = undefined;
		switch (oParsedAnimMotion["origin"])
		{
			case "parent":
				nOriginType = c_oAscSlideTLOriginType.Parent;
				break;
			case "layout":
				nOriginType = c_oAscSlideTLOriginType.Layout;
				break;
		}

		var nPathEditMode = undefined;
		switch (oParsedAnimMotion["pathEditMode"])
		{
			case "fixed":
				nPathEditMode = c_oAscSlideTLPathEditMode.Fixed;
				break;
			case "relative":
				nPathEditMode = c_oAscSlideTLPathEditMode.Relative;
				break;
		}

		oParsedAnimMotion["by"] && oAnimMotion.setBy(this.ByPointFromJSON(oParsedAnimMotion["by"]));
		oParsedAnimMotion["cBhvr"] && oAnimMotion.setCBhvr(this.CBhvrFromJSON(oParsedAnimMotion["cBhvr"]));
		oParsedAnimMotion["from"] && oAnimMotion.setFrom(this.ByPointFromJSON(oParsedAnimMotion["from"]));
		oParsedAnimMotion["rCtr"] && oAnimMotion.setRCtr(this.ByPointFromJSON(oParsedAnimMotion["rCtr"]));
		oParsedAnimMotion["to"] && oAnimMotion.setTo(this.ByPointFromJSON(oParsedAnimMotion["to"]));

		nOriginType != undefined && oAnimMotion.setOrigin(nOriginType);
		oParsedAnimMotion["path"] != undefined && oAnimMotion.setPath(oParsedAnimMotion["path"]);
		nPathEditMode != undefined && oAnimMotion.setPathEditMode(nPathEditMode);
		oParsedAnimMotion["ptsTypes"] != undefined && oAnimMotion.setPtsTypes(oParsedAnimMotion["ptsTypes"]);
		oParsedAnimMotion["rAng"] != undefined && oAnimMotion.setRAng(oParsedAnimMotion["rAng"]);

		return oAnimMotion;
	};
	ReaderFromJSON.prototype.ByPointFromJSON = function(oParsedByPoint)
	{
		var oTLPoint = new AscFormat.CTLPoint();

		oParsedByPoint["x"] != undefined && oTLPoint.setX(oParsedByPoint["x"]);
		oParsedByPoint["y"] != undefined && oTLPoint.setY(oParsedByPoint["y"]);

		return oTLPoint;
	};
	ReaderFromJSON.prototype.AnimRotFromJSON = function(oParsedAnimRot)
	{
		var oAnimRot = new AscFormat.CAnimRot();

		oParsedAnimRot["cBhvr"] && oAnimRot.setCBhvr(this.CBhvrFromJSON(oParsedAnimRot["cBhvr"]));
		oParsedAnimRot["by"] != undefined && oAnimRot.setBy(oParsedAnimRot["by"]);
		oParsedAnimRot["from"] != undefined && oAnimRot.setFrom(oParsedAnimRot["from"]);
		oParsedAnimRot["to"] != undefined && oAnimRot.setTo(oParsedAnimRot["to"]);

		return oAnimRot;
	};
	ReaderFromJSON.prototype.AnimScaleFromJSON = function(oParsedAnimScale)
	{
		var oAnimScale = new AscFormat.CAnimScale();

		oParsedAnimScale["cBhvr"] && oAnimScale.setCBhvr(this.CBhvrFromJSON(oParsedAnimScale["cBhvr"]));
		oParsedAnimScale["by"] && oAnimScale.setBy(this.ByPointFromJSON(oParsedAnimScale["by"]));
		oParsedAnimScale["from"] && oAnimScale.setFrom(this.ByPointFromJSON(oParsedAnimScale["from"]));
		oParsedAnimScale["to"] && oAnimScale.setTo(this.ByPointFromJSON(oParsedAnimScale["to"]));
		oParsedAnimScale["zoomContents"] != undefined && oAnimScale.setZoomContents(oParsedAnimScale["zoomContents"]);

		return oAnimScale;
	};
	ReaderFromJSON.prototype.CmdFromJSON = function(oParsedCmd)
	{
		var oCmd = new AscFormat.CCmd();

		var nCommandType = undefined;
		switch (oParsedCmd["type"])
		{
			case "call":
				nCommandType = c_oAscSlideTLCommandType.Call;
				break;
			case "evt":
				nCommandType = c_oAscSlideTLCommandType.Evt;
				break;
			case "verb":
				nCommandType = c_oAscSlideTLCommandType.Verb;
				break;
		}

		oParsedCmd["cBhvr"] && oCmd.setCBhvr(this.CBhvrFromJSON(oParsedCmd["cBhvr"]));
		oParsedCmd["cmd"] && oCmd.setCmd(oParsedCmd["cmd"]);
		nCommandType != undefined && oCmd.setType(nCommandType);

		return oCmd;
	};
	ReaderFromJSON.prototype.SetFromJSON = function(oParsedSet)
	{
		var oSet = new AscFormat.CSet();

		oParsedSet["cBhvr"] && oSet.setCBhvr(this.CBhvrFromJSON(oParsedSet["cBhvr"]));
		oParsedSet["to"] && oSet.setTo(this.AnimVariantFromJSON(oParsedSet["to"]));

		return oSet;
	};
	ReaderFromJSON.prototype.HFFromJSON = function(oParsedHF)
	{
		var oHF = new AscFormat.HF();

		oParsedHF["dt"] != undefined && oHF.setDt(oParsedHF["dt"]);
		oParsedHF["ftr"] != undefined && oHF.setFtr(oParsedHF["ftr"]);
		oParsedHF["hdr"] != undefined && oHF.setHdr(oParsedHF["hdr"]);
		oParsedHF["sldNum"] != undefined && oHF.setSldNum(oParsedHF["sldNum"]);

		return oHF;
	};
	ReaderFromJSON.prototype.CSldFromJSON = function(oParsedCSld)
	{
		var oCSld = new AscFormat.CSld();

		for (var nShape = 0; nShape < oParsedCSld["spTree"].length; nShape++)
			oCSld.spTree.push(this.GraphicObjFromJSON(oParsedCSld["spTree"][nShape]));
			
		oCSld.Bg   = oParsedCSld["bg"] ? this.BgFromJSON(oParsedCSld["bg"]) : oCSld.Bg;
		oCSld.name = oParsedCSld.name;

		return oCSld;
	};
	ReaderFromJSON.prototype.BgFromJSON = function(oParsedBg)
	{
		var oBg = new AscFormat.CBg();
	
		var nBwModeType = undefined;
		switch (oParsedBg["bwMode"])
		{
			case "auto":
				nBwModeType = c_oAscSlideBgBwModeType.Auto;
				break;
			case "black":
				nBwModeType = c_oAscSlideBgBwModeType.Black;
				break;
			case "blackGray":
				nBwModeType = c_oAscSlideBgBwModeType.BlackGray;
				break;
			case "blackWhite":
				nBwModeType = c_oAscSlideBgBwModeType.BlackWhite;
				break;
			case "clr":
				nBwModeType = c_oAscSlideBgBwModeType.Clr;
				break;
			case "gray":
				nBwModeType = c_oAscSlideBgBwModeType.Gray;
				break;
			case "grayWhite":
				nBwModeType = c_oAscSlideBgBwModeType.GrayWhite;
				break;
			case "hidden":
				nBwModeType = c_oAscSlideBgBwModeType.Hidden;
				break;
			case "invGray":
				nBwModeType = c_oAscSlideBgBwModeType.InvGray;
				break;
			case "ltGray":
				nBwModeType = c_oAscSlideBgBwModeType.LtGray;
				break;
			case "white":
				nBwModeType = c_oAscSlideBgBwModeType.White;
				break;
		}

		nBwModeType != undefined && oBg.setBwMode(nBwModeType);
		oParsedBg["bgPr"] && oBg.setBgPr(this.BgPrFromJSON(oParsedBg["bgPr"]));
		oParsedBg["bgRef"] && oBg.setBgRef(this.StyleRefFromJSON(oParsedBg["bgRef"]));

		return oBg;
	};
	ReaderFromJSON.prototype.BgPrFromJSON = function(oParsedBgPr)
	{
		var oBgPr = new AscFormat.CBgPr();

		oParsedBgPr["fill"] && oBgPr.setFill(this.FillFromJSON(oParsedBgPr["fill"]));
		oParsedBgPr["setShadeToTitle"] && oBgPr.setShadeToTitle(oParsedBgPr["setShadeToTitle"]);

		return oBgPr;
	};
	
	function To_XML_TLAnimateEffectTransition(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case AscCommon.ST_TLAnimateEffectTransition.in:
				sVal = "in";
				break;
			case AscCommon.ST_TLAnimateEffectTransition.out:
				sVal = "out";
				break;
			case AscCommon.ST_TLAnimateEffectTransition.none:
				sVal = "none";
				break;
		}

		return sVal;
	}
	function From_XML_TLAnimateEffectTransition(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "in":
				nVal = AscCommon.ST_TLAnimateEffectTransition.in;
				break;
			case "out":
				nVal = AscCommon.ST_TLAnimateEffectTransition.out;
				break;
			case "none":
				nVal = AscCommon.ST_TLAnimateEffectTransition.none;
				break;
		}

		return nVal;
	}
    //----------------------------------------------------------export----------------------------------------------------
    window['AscCommon']       = window['AscCommon'] || {};
    window['AscFormat']       = window['AscFormat'] || {};
	
})(window);




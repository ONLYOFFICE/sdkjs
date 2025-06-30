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
var align_Left = AscCommon.align_Left;
var align_Justify = AscCommon.align_Justify;
var vertalign_Baseline = AscCommon.vertalign_Baseline;
var changestype_Drawing_Props = AscCommon.changestype_Drawing_Props;
var g_oTableId = AscCommon.g_oTableId;
var isRealObject = AscCommon.isRealObject;
var History = AscCommon.History;

var CreateUnifillSolidFillSchemeColor = AscFormat.CreateUnifillSolidFillSchemeColor;

function DrawingCopyObject(Drawing, X, Y, ExtX, ExtY, ImageUrl) {
	this.Drawing = Drawing;
	this.X = X;
	this.Y = Y;
	this.ExtX = ExtX;
	this.ExtY = ExtY;
	this.ImageUrl = ImageUrl;
}

DrawingCopyObject.prototype.copy = function (oIdMap) {

	var _copy = this.Drawing;
	var oPr = new AscFormat.CCopyObjectProperties();
	oPr.idMap = oIdMap;
	if (this.Drawing) {
		_copy = this.Drawing.copy(oPr);
		if (AscCommon.isRealObject(oIdMap)) {
			oIdMap[this.Drawing.Id] = _copy.Id;
		}
	}
	return new DrawingCopyObject(this.Drawing ? _copy : this.Drawing, this.X, this.Y, this.ExtX, this.ExtY, this.ImageUrl);

};

const PE_SELECTED_CONTENT_EMPTY = 0;
const PE_SELECTED_CONTENT_SLIDES = 1;
const PE_SELECTED_CONTENT_DRAWINGS = 2;
const PE_SELECTED_CONTENT_DOC_CONTENT = 3;

function PresentationSelectedContent() {
	this.SlideObjects = [];
	this.Notes = [];
	this.NotesMasters = [];
	this.NotesMastersIndexes = [];
	this.NotesThemes = [];
	this.LayoutsIndexes = [];
	this.Layouts = [];
	this.MastersIndexes = [];
	this.Masters = [];
	this.ThemesIndexes = [];
	this.Themes = [];
	this.Drawings = [];
	this.DocContent = null;
	this.PresentationWidth = null;
	this.PresentationHeight = null;
	this.ThemeName = null;
}

PresentationSelectedContent.prototype.copy = function () {
	var ret = new PresentationSelectedContent(), i, oIdMap, oSlide, oNotes, oNotesMaster, oLayout, aElements,
		oSelectedElement, oElement, oParagraph;
	for (i = 0; i < this.SlideObjects.length; ++i) {
		oIdMap = {};
		oSlide = this.SlideObjects[i].createDuplicate(oIdMap);
		AscFormat.fResetConnectorsIds(oSlide.cSld.spTree, oIdMap);
		ret.SlideObjects.push(oSlide);
	}
	for (i = 0; i < this.Notes.length; ++i) {
		oIdMap = {};
		oNotes = this.Notes[i].createDuplicate(oIdMap);
		AscFormat.fResetConnectorsIds(oNotes.cSld.spTree, oIdMap);
		ret.Notes.push(oNotes);
	}
	for (i = 0; i < this.NotesMasters.length; ++i) {
		oIdMap = {};
		oNotesMaster = this.NotesMasters[i].createDuplicate(oIdMap);
		AscFormat.fResetConnectorsIds(oNotesMaster.cSld.spTree, oIdMap);
		ret.NotesMasters.push(oNotesMaster);
	}
	for (i = 0; i < this.NotesMastersIndexes.length; ++i) {
		ret.NotesMastersIndexes.push(this.NotesMastersIndexes[i]);
	}
	for (i = 0; i < this.NotesThemes.length; ++i) {
		ret.NotesThemes.push(this.NotesThemes[i].createDuplicate());
	}
	for (i = 0; i < this.LayoutsIndexes.length; ++i) {
		ret.LayoutsIndexes.push(this.LayoutsIndexes[i]);
	}
	for (i = 0; i < this.Layouts.length; ++i) {
		oIdMap = {};
		oLayout = this.Layouts[i].createDuplicate(oIdMap);
		AscFormat.fResetConnectorsIds(oLayout.cSld.spTree, oIdMap);
		ret.Layouts.push(oLayout);
	}
	for (i = 0; i < this.MastersIndexes.length; ++i) {
		ret.MastersIndexes.push(this.MastersIndexes[i]);
	}

	for (i = 0; i < this.Masters.length; ++i) {
		oIdMap = {};
		oNotesMaster = this.Masters[i].createDuplicate(oIdMap);
		AscFormat.fResetConnectorsIds(oNotesMaster.cSld.spTree, oIdMap);
		ret.Masters.push(oNotesMaster);
	}

	for (i = 0; i < this.ThemesIndexes.length; ++i) {
		ret.ThemesIndexes.push(this.ThemesIndexes[i]);
	}
	for (i = 0; i < this.Themes.length; ++i) {
		ret.Themes.push(this.Themes[i].createDuplicate());
	}


	oIdMap = {};
	var oPr = new AscFormat.CCopyObjectProperties();
	oPr.idMap = oIdMap;
	var aDrawingsCopy = [];
	for (i = 0; i < this.Drawings.length; ++i) {
		ret.Drawings.push(this.Drawings[i].copy(oPr));
		if (ret.Drawings[ret.Drawings.length - 1].Drawing) {
			aDrawingsCopy.push(ret.Drawings[ret.Drawings.length - 1].Drawing);
		}
	}
	AscFormat.fResetConnectorsIds(aDrawingsCopy, oIdMap);
	if (this.DocContent) {
		//TODO: перенести копирование в CSelectedContent;
		ret.DocContent = new AscCommonWord.CSelectedContent();
		aElements = this.DocContent.Elements;
		for (i = 0; i < aElements.length; ++i) {
			oSelectedElement = new AscCommonWord.CSelectedElement();
			oElement = aElements[i];
			oParagraph = aElements[i].Element;
			oSelectedElement.SelectedAll = oElement.SelectedAll;

			oSelectedElement.Element = oParagraph.Copy(oParagraph.Parent, oParagraph.DrawingDocument, {});
			ret.DocContent.Elements[i] = oSelectedElement;
		}
	}
	ret.PresentationWidth = this.PresentationWidth;
	ret.PresentationHeight = this.PresentationHeight;
	ret.ThemeName = this.ThemeName;
	return ret;
};

PresentationSelectedContent.prototype.getContentType = function () {
	if (this.SlideObjects.length > 0) {
		return PE_SELECTED_CONTENT_SLIDES;
	} else if (this.Drawings.length > 0) {
		return PE_SELECTED_CONTENT_DRAWINGS;
	} else if (this.DocContent) {
		return PE_SELECTED_CONTENT_DOC_CONTENT;
	}
	return PE_SELECTED_CONTENT_EMPTY;
};
PresentationSelectedContent.prototype.isSlidesContent = function () {
	return this.getContentType() === PE_SELECTED_CONTENT_SLIDES;
};
PresentationSelectedContent.prototype.isDrawingsContent = function () {
	return this.getContentType() === PE_SELECTED_CONTENT_DRAWINGS;
};
PresentationSelectedContent.prototype.isDocContent = function () {
	return this.getContentType() === PE_SELECTED_CONTENT_DOC_CONTENT;
};
PresentationSelectedContent.prototype.isValid = function() {
	return !!(
		this.DocContent ||
		(this.Drawings && this.Drawings.length) ||
		(this.SlideObjects && this.SlideObjects.length) ||
		(this.Layouts && this.Layouts.length) ||
		(this.Masters && this.Masters.length)
	);
};

function CPrSection() {
	this.name = null;
	this.startIndex = null;
	this.guid = null;
	this.Id = AscCommon.g_oIdCounter.Get_NewId();
	AscCommon.g_oTableId.Add(this, this.Id);
}

CPrSection.prototype.getObjectType = function () {
	return AscDFH.historyitem_type_PresentationSection;
};

CPrSection.prototype.Get_Id = function () {
	return this.Id;
};

CPrSection.prototype.Write_ToBinary2 = function (w) {
	w.WriteLong(this.getObjectType());
	w.WriteString2(this.Get_Id());
};
CPrSection.prototype.Read_FromBinary2 = function (r) {
	this.Id = r.GetString2();
};
CPrSection.prototype.setName = function (pr) {
	History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_PresentationSectionSetName, this.name, pr));
	this.name = pr;
};
CPrSection.prototype.setStartIndex = function (pr) {
	History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_PresentationSectionSetStartIndex, this.startIndex, pr));
	this.startIndex = pr;
};
CPrSection.prototype.setGuid = function (pr) {
	History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_PresentationSectionSetGuid, this.guid, pr));
	this.guid = pr;
};
CPrSection.prototype.Read_FromBinary2 = function (r) {
	this.Id = r.GetString2();
};

function CShowPr() {
	AscFormat.CBaseNoIdObject.call(this);
	this.browse = undefined;
	this.kiosk = undefined; //{restart: uInt}
	this.penClr = undefined;
	this.present = false;
	this.show = undefined;// {showAll: true/false, range: {start: uInt, end: uInt}, custShow: uInt};
	this.loop = undefined;
	this.showAnimation = undefined;
	this.showNarration = undefined;
	this.useTimings = undefined;
}

AscFormat.InitClass(CShowPr, AscFormat.CBaseNoIdObject, 0);

CShowPr.prototype.Write_ToBinary = function (w) {
	var nStartPos = w.GetCurPosition();
	w.Skip(4);
	var Flags = 0;
	if (AscFormat.isRealBool(this.browse)) {
		Flags |= 1;
		w.WriteBool(this.browse);
	}
	if (isRealObject(this.kiosk)) {
		Flags |= 2;
		if (AscFormat.isRealNumber(this.kiosk.restart)) {
			Flags |= 4;
			w.WriteLong(this.kiosk.restart);
		}
	}
	if (isRealObject(this.penClr)) {
		Flags |= 8;
		this.penClr.Write_ToBinary(w);
	}
	w.WriteBool(this.present);
	if (isRealObject(this.show)) {
		Flags |= 16;
		w.WriteBool(this.show.showAll);
		if (!this.show.showAll) {
			if (this.show.range) {
				Flags |= 32;
				w.WriteLong(this.show.range.start);
				w.WriteLong(this.show.range.end);
			} else if (AscFormat.isRealNumber(this.show.custShow)) {
				Flags |= 64;
				w.WriteLong(this.show.custShow);
			}
		}
	}
	if (AscFormat.isRealBool(this.loop)) {
		Flags |= 128;
		w.WriteBool(this.loop);
	}
	if (AscFormat.isRealBool(this.showAnimation)) {
		Flags |= 256;
		w.WriteBool(this.showAnimation);
	}
	if (AscFormat.isRealBool(this.showNarration)) {
		Flags |= 512;
		w.WriteBool(this.showNarration);
	}
	if (AscFormat.isRealBool(this.useTimings)) {
		Flags |= 1024;
		w.WriteBool(this.useTimings);
	}
	var nEndPos = w.GetCurPosition();
	w.Seek(nStartPos);
	w.WriteLong(Flags);
	w.Seek(nEndPos);
};

CShowPr.prototype.Read_FromBinary = function (r) {
	var Flags = r.GetLong();
	if (Flags & 1) {
		this.browse = r.GetBool();
	}
	if (Flags & 2) {
		this.kiosk = {};
		if (Flags & 4) {
			this.kiosk.restart = r.GetLong();
		}
	}
	if (Flags & 8) {
		this.penClr = new AscFormat.CUniColor();
		this.penClr.Read_FromBinary(r);
	}
	this.present = r.GetBool();
	if (Flags & 16) {
		this.show = {};
		this.show.showAll = r.GetBool();
		if (Flags & 32) {
			var start = r.GetLong();
			var end = r.GetLong();
			this.show.range = {start: start, end: end};
		} else if (Flags & 64) {
			this.show.custShow = r.GetLong();
		}
	}
	if (Flags & 128) {
		this.loop = r.GetBool();
	}
	if (Flags & 256) {
		this.showAnimation = r.GetBool();
	}
	if (Flags & 512) {
		this.showNarration = r.GetBool();
	}
	if (Flags & 1024) {
		this.useTimings = r.GetBool();
	}
};

CShowPr.prototype.Copy = function () {
	var oCopy = new CShowPr();
	oCopy.browse = this.browse;
	if (isRealObject(this.kiosk)) {
		oCopy.kiosk = {};
		if (AscFormat.isRealBool(this.kiosk.restart)) {
			oCopy.kiosk.restart = this.kiosk.restart;
		}
	}
	if (this.penClr) {
		oCopy.penClr = this.penClr.createDuplicate();
	}
	oCopy.present = this.present;
	if (isRealObject(this.show)) {
		oCopy.show = {};
		oCopy.show.showAll = this.show.showAll;
		if (isRealObject(this.show.range)) {
			oCopy.show.range = {start: this.show.range.start, end: this.show.range.end};
		} else if (AscFormat.isRealNumber(this.show.custShow)) {
			oCopy.show.custShow = this.show.custShow;
		}
	}
	oCopy.loop = this.loop;
	oCopy.showAnimation = this.showAnimation;
	oCopy.showNarration = this.showNarration;
	oCopy.useTimings = this.useTimings;
	return oCopy;
};


AscDFH.changesFactory[AscDFH.historyitem_Presentation_SetShowPr] = AscDFH.CChangesDrawingsObjectNoId;
AscDFH.changesFactory[AscDFH.historyitem_Presentation_AddSlideMaster] = AscDFH.CChangesDrawingsContent;
AscDFH.changesFactory[AscDFH.historyitem_Presentation_ChangeTheme] = AscDFH.CChangesDrawingChangeTheme;
AscDFH.changesFactory[AscDFH.historyitem_Presentation_SlideSize] = AscDFH.CChangesDrawingsObject;
AscDFH.changesFactory[AscDFH.historyitem_Presentation_NotesSz] = AscDFH.CChangesDrawingsObject;
AscDFH.changesFactory[AscDFH.historyitem_Presentation_ChangeColorScheme] = AscDFH.CChangesDrawingChangeTheme;
AscDFH.changesFactory[AscDFH.historyitem_Presentation_RemoveSlide] = AscDFH.CChangesDrawingsContentPresentation;
AscDFH.changesFactory[AscDFH.historyitem_Presentation_AddSlide] = AscDFH.CChangesDrawingsContentPresentation;
AscDFH.changesFactory[AscDFH.historyitem_Presentation_SetDefaultTextStyle] = AscDFH.CChangesDrawingsObjectNoId;
AscDFH.changesFactory[AscDFH.historyitem_Presentation_SetFirstSlideNum] = AscDFH.CChangesDrawingsLong;
AscDFH.changesFactory[AscDFH.historyitem_Presentation_SetShowSpecialPlsOnTitleSld] = AscDFH.CChangesDrawingsBool;
AscDFH.changesFactory[AscDFH.historyitem_Presentation_RemoveSlideMaster] = AscDFH.CChangesDrawingsContent;


AscDFH.changesFactory[AscDFH.historyitem_PresentationSectionSetName] = AscDFH.CChangesDrawingsString;
AscDFH.changesFactory[AscDFH.historyitem_PresentationSectionSetStartIndex] = AscDFH.CChangesDrawingsLong;
AscDFH.changesFactory[AscDFH.historyitem_PresentationSectionSetGuid] = AscDFH.CChangesDrawingsString;

AscDFH.changesFactory[AscDFH.historyitem_SldSzCX] = AscDFH.CChangesDrawingsLong;
AscDFH.changesFactory[AscDFH.historyitem_SldSzCY] = AscDFH.CChangesDrawingsLong;
AscDFH.changesFactory[AscDFH.historyitem_SldSzType] = AscDFH.CChangesDrawingsLong;

AscDFH.drawingsChangesMap[AscDFH.historyitem_SldSzCX] = function (oClass, value) {
	oClass.cx = value;
};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SldSzCY] = function (oClass, value) {
	oClass.cy = value;
};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SldSzType] = function (oClass, value) {
	oClass.type = value;
};

AscDFH.drawingsChangesMap[AscDFH.historyitem_PresentationSectionSetName] = function (oClass, value) {
	oClass.name = value;
};
AscDFH.drawingsChangesMap[AscDFH.historyitem_PresentationSectionSetStartIndex] = function (oClass, value) {
	oClass.startIndex = value;
};
AscDFH.drawingsChangesMap[AscDFH.historyitem_PresentationSectionSetGuid] = function (oClass, value) {
	oClass.guid = value;
};

AscDFH.drawingsChangesMap[AscDFH.historyitem_Presentation_SetShowPr] = function (oClass, value) {
	oClass.showPr = value;
};
AscDFH.drawingsChangesMap[AscDFH.historyitem_Presentation_SlideSize] = function (oClass, value) {
	oClass.sldSz = value;
	oClass.changeSlideSizeFunction();
};
AscDFH.drawingsChangesMap[AscDFH.historyitem_Presentation_NotesSz] = function (oClass, value) {
	oClass.notesSz = value;
};
AscDFH.drawingsChangesMap[AscDFH.historyitem_Presentation_SetDefaultTextStyle] = function (oClass, value) {
	oClass.defaultTextStyle = value;
};
AscDFH.drawingsChangesMap[AscDFH.historyitem_Presentation_SetFirstSlideNum] = function (oClass, value) {
	oClass.firstSlideNum = value;
};
AscDFH.drawingsChangesMap[AscDFH.historyitem_Presentation_SetShowSpecialPlsOnTitleSld] = function (oClass, value) {
	oClass.showSpecialPlsOnTitleSld = value;
};


AscDFH.drawingContentChanges[AscDFH.historyitem_Presentation_AddSlide] = function (oClass) {
	return oClass.Slides;
};
AscDFH.drawingContentChanges[AscDFH.historyitem_Presentation_RemoveSlide] = function (oClass) {
	return oClass.Slides;
};
AscDFH.drawingContentChanges[AscDFH.historyitem_Presentation_AddSlideMaster] = function (oClass) {
	return oClass.slideMasters;
};
AscDFH.drawingContentChanges[AscDFH.historyitem_Presentation_RemoveSlideMaster] = function (oClass) {
	return oClass.slideMasters;
};

AscDFH.drawingsConstructorsMap[AscDFH.historyitem_Presentation_SetShowPr] = CShowPr;
AscDFH.drawingsConstructorsMap[AscDFH.historyitem_Presentation_SetDefaultTextStyle] = AscFormat.TextListStyle;


function CSlideSize() {
	AscFormat.CBaseFormatObject.call(this);
	this.cx = null;
	this.cy = null;
	this.type = null;
}

AscFormat.InitClass(CSlideSize, AscFormat.CBaseFormatObject, AscDFH.historyitem_type_SldSz);
CSlideSize.prototype.DEFAULT_CX = 9144000;
CSlideSize.prototype.DEFAULT_CY = 6858000;
CSlideSize.prototype.static_CreateNotesSize = function () {
	return AscFormat.ExecuteNoHistory(function () {
		let oSize = new CSlideSize();
		oSize.setCX(this.DEFAULT_CX);
		oSize.setCY(this.DEFAULT_CY);
		return oSize;
	}, this, []);
};
CSlideSize.prototype.setCX = function (pr) {
	History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_SldSzCX, this.cx, pr));
	this.cx = pr;
};
CSlideSize.prototype.setCY = function (pr) {
	History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_SldSzCY, this.cy, pr));
	this.cy = pr;
};
CSlideSize.prototype.setType = function (pr) {
	History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_SldSzType, this.type, pr));
	this.type = pr;
};
CSlideSize.prototype.GetWidthEMU = function () {
	if (AscFormat.isRealNumber(this.cx)) {
		return this.cx;
	}
	return this.DEFAULT_CX;
};
CSlideSize.prototype.GetHeightEMU = function () {
	if (AscFormat.isRealNumber(this.cy)) {
		return this.cy;
	}
	return this.DEFAULT_CY;
};
CSlideSize.prototype.GetWidthMM = function () {
	return this.GetWidthEMU() / g_dKoef_mm_to_emu;
};
CSlideSize.prototype.GetHeightMM = function () {
	return this.GetHeightEMU() / g_dKoef_mm_to_emu;
};
CSlideSize.prototype.GetSizeType = function () {
	if (AscFormat.isRealNumber(this.type)) {
		return this.type;
	}
	return Asc.c_oAscSlideSZType.SzCustom;
};

CSlideSize.prototype.fillObject = function (oCopy, oIdMap) {
	oCopy.setCX(this.cx);
	oCopy.setCY(this.cy);
	oCopy.setType(this.type);
};

const CONFORMANCE_STRICT = 0;
const CONFORMANCE_TRANSITIONAL = 1;

function CPresentation(DrawingDocument) {
	AscFormat.CBaseFormatObject.call(this);
	this.History = History;
	this.IdCounter = AscCommon.g_oIdCounter;
	this.TableId = g_oTableId;
	this.CollaborativeEditing = (("undefined" !== typeof (CCollaborativeEditing) && AscCommon.CollaborativeEditing instanceof CCollaborativeEditing) ? AscCommon.CollaborativeEditing : null);
	this.Api = Asc.editor;
	this.TurnOffInterfaceEvents = false;
	//------------------------------------------------------------------------
	if (DrawingDocument) {
		if (this.History)
			this.History.Set_LogicDocument(this);

		if (this.CollaborativeEditing)
			this.CollaborativeEditing.SetLogicDocument(this);
	}

	//------------------------------------------------------------------------

	this.firstSlideNum = null;
	this.showSpecialPlsOnTitleSld = null;

	this.Id = AscCommon.g_oIdCounter.Get_NewId();
	//Props
	this.App = null;
	this.Core = null;
	this.CustomProperties = new AscCommon.CCustomProperties();;

	this.StartPage = 0; // Для совместимости с CDocumentContent
	this.CurPage = 0;

	this.TurnOffRecalc = false;

	this.DrawingDocument = DrawingDocument;

	this.SearchEngine = new AscCommonWord.CDocumentSearch(this);

	this.NeedUpdateTarget = false;

	this.noShowContextMenu = false;

	this.viewMode = false;
	// Класс для работы с поиском
	this.SearchInfo =
		{
			Id: null,
			StartPos: 0,
			CurPage: 0,
			String: null
		};

	// Позция каретки
	this.TargetPos =
		{
			X: 0,
			Y: 0,
			PageNum: 0
		};


	this.Lock = new AscCommon.CLock();

	this.m_oContentChanges = new AscCommon.CContentChanges(); // список изменений(добавление/удаление элементов)


	this.Slides = [];
	this.slideMasters = [];
	this.notesMasters = [];
	this.notes = [];
	this.globalTableStyles = null;
	this.TrackMoveId = null;


	this.sldSz = null;
	this.notesSz = null;
	this.viewPr = null;

	this.strideData = null;

	this.recalcMap = {};
	this.bNeedUpdateTh = false;
	this.needSelectPages = [];

	this.writecomments = [];

	this.forwardChangeThemeTimeOutId = null;
	this.backChangeThemeTimeOutId = null;
	this.startChangeThemeTimeOutId = null;
	this.TablesForInterface = null;
	this.LastTheme = null;
	this.LastColorScheme = null;
	this.LastColorMap = null;
	this.LastTableLook = null;
	this.DefaultSlideTransition = new Asc.CAscSlideTransition();
	this.DefaultSlideTransition.setDefaultParams();

	this.DefaultTableStyleId = null;
	this.TableStylesIdMap = {};
	this.bNeedUpdateChartPreview = false;
	this.LastUpdateTargetTime = 0;
	this.NeedUpdateTargetForCollaboration = false;
	this.oLastCheckContent = null;
	this.CompositeInput = null;


	this.Spelling = new AscWord.CDocumentSpellChecker();

	this.Sections = [];//array of CPrSection


	this.comments = new SlideComments(this);

	this.CheckLanguageOnTextAdd = false;

	// Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
	g_oTableId.Add(this, this.Id);
	//
	this.hdrFtrLock = new PropLocker(this.Id);
	this.themeLock = new PropLocker(this.Id);
	this.schemeLock = new PropLocker(this.Id);
	this.slideSizeLock = new PropLocker(this.Id);
	this.defaultTextStyleLock = new PropLocker(this.Id);
	this.commentsLock = new PropLocker(this.Id);
	this.viewPrLock = new PropLocker(this.Id);

	this.RecalcId = 0; // Номер пересчета
	this.CommentAuthors = {};
	this.createDefaultTableStyles();
	this.bGoToPage = false;

	this.custShowList = [];
	this.clrMru = [];
	this.prnPr = null;
	this.showPr = null;

	this.CurPosition =
		{
			X: 0, Y: 0
		};

	this.NotesWidth = -10;
	this.FocusOnNotes = false;

	this.lastMaster = null;

	this.AutoCorrectSettings = new AscCommon.CAutoCorrectSettings();

	this.MathTrackHandler = new AscWord.CMathTrackHandler(DrawingDocument, this.Api);

	this.cachedGridCanvas = null;
	this.cachedGridSpacing = null;

	this.cachedAnimationLabelText = {};
}

AscFormat.InitClass(CPresentation, AscFormat.CBaseFormatObject, AscDFH.historyitem_type_Presentation);
CPresentation.prototype.getFirstSlideImagesMap = function () {
	const oRequiredSyncImagesMap = {};
	const oFirstSlide = this.Slides[0];
	if (oFirstSlide) {
		const aImages = oFirstSlide.getAllRasterImagesOnSlide();
		for (let i = 0; i < aImages.length; i++) {
			const sImg = aImages[i];
			oRequiredSyncImagesMap[sImg] = true;
		}
	}
	return oRequiredSyncImagesMap;
};
CPresentation.prototype.sendEvent = function()
{
	if (!this.Api)
		return;
	this.Api.sendEvent.apply(this.Api, arguments);
};
CPresentation.prototype.notAllowedWithoutId = function () {
	return true;
};
CPresentation.prototype.GetApi = function () {
	return this.Api;
};
CPresentation.prototype.GetHistory = function () {
	return this.History;
};
CPresentation.prototype.IsDocumentEditor = function () {
	return false;
};
CPresentation.prototype.IsPresentationEditor = function () {
	return true;
};
CPresentation.prototype.IsSpreadSheetEditor = function () {
	return false;
};
CPresentation.prototype.IsPdfEditor = function() {
	return false;
};
CPresentation.prototype.GetWidthMM = function () {
	return this.GetWidthEMU() / g_dKoef_mm_to_emu;
};

CPresentation.prototype.GetHeightMM = function () {
	return this.GetHeightEMU() / g_dKoef_mm_to_emu;
};
CPresentation.prototype.GetNotesWidthMM = function () {
	return this.GetNotesWidthEMU() / g_dKoef_mm_to_emu;
};

CPresentation.prototype.GetNotesHeightMM = function () {
	return this.GetNotesHeightEMU() / g_dKoef_mm_to_emu;
};
CPresentation.prototype.GetWidthEMU = function () {
	if (this.sldSz) {
		return this.sldSz.GetWidthEMU();
	}
	return CSlideSize.prototype.DEFAULT_CX;
};


CPresentation.prototype.GetHeightEMU = function () {
	if (this.sldSz) {
		return this.sldSz.GetHeightEMU();
	}
	return CSlideSize.prototype.DEFAULT_CY;
};

CPresentation.prototype.GetNotesWidthEMU = function () {
	if (this.notesSz) {
		return this.notesSz.GetWidthEMU();
	}
	return CSlideSize.prototype.DEFAULT_CX;
};
CPresentation.prototype.GetNotesHeightEMU = function () {
	if (this.notesSz) {
		return this.notesSz.GetHeightEMU();
	}
	return CSlideSize.prototype.DEFAULT_CY;
};
CPresentation.prototype.GetSizeType = function () {
	if (this.sldSz) {
		return this.sldSz.GetSizeType();
	}
	return Asc.c_oAscSlideSZType.SzCustom;
};
CPresentation.prototype.setSldSz = function (pr) {
	History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_Presentation_SlideSize, this.sldSz, pr));
	this.sldSz = pr;
};
CPresentation.prototype.setNotesSz = function (pr) {
	History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_Presentation_NotesSz, this.notesSz, pr));
	this.notesSz = pr;
};
CPresentation.prototype.setViewPr = function (pr) {
	History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_Presentation_ViewPr, this.viewPr, pr));
	this.viewPr = pr;
	if (this.viewPr) {
		this.viewPr.setParent(this);
	}
};
CPresentation.prototype.getStrideData = function () {
	if (!this.strideData) {
		this.strideData = new AscCommonSlide.CStrideData(this);
	}
	return this.strideData;
};

CPresentation.prototype.changeSlideSizeFunction = function () {
	AscFormat.ExecuteNoHistory(function () {
		let i;
		const dWidth = this.GetWidthMM();
		const dHeight = this.GetHeightMM();
		let oFirstMaster = this.slideMasters[0];
		if (oFirstMaster) {
			let dOldWidth = oFirstMaster.Width;
			let dOldHeight = oFirstMaster.Height;
			let dCW = dWidth / dOldWidth;
			let dCH = dHeight / dOldHeight;
			if (!AscFormat.fApproxEqual(dCW, 1.0) || !AscFormat.fApproxEqual(dCW, 1.0)) {
				this.scaleGuides(dCW, dCH);
			}
		}
		for (i = 0; i < this.slideMasters.length; ++i) {
			this.slideMasters[i].changeSize(dWidth, dHeight);
			var master = this.slideMasters[i];
			for (var j = 0; j < master.sldLayoutLst.length; ++j) {
				master.sldLayoutLst[j].changeSize(dWidth, dHeight);
			}
		}
		for (i = 0; i < this.Slides.length; ++i) {
			this.Slides[i].changeSize(dWidth, dHeight);
		}
	}, this, []);
};

CPresentation.prototype.internalChangeSizes = function (nWidth, nHeight, nType) {
	var oSldSize = new CSlideSize();
	oSldSize.setCX(nWidth);
	oSldSize.setCY(nHeight);
	if (AscFormat.isRealNumber(nType) && nType !== Asc.c_oAscSlideSZType.SzWidescreen && nType !== Asc.c_oAscSlideSZType.SzCustom) {
		oSldSize.setType(nType);
	}
	this.setSldSz(oSldSize);
	this.changeSlideSizeFunction();
};

CPresentation.prototype.changeSlideSize = function (width, height, nType, nFirstSlideNum) {
	if (this.Document_Is_SelectionLocked(AscCommon.changestype_SlideSize) === false) {
		History.Create_NewPoint(AscDFH.historydescription_Presentation_ChangeSlideSize);
		this.internalChangeSizes(width, height, nType);
		if(AscFormat.isRealNumber(nFirstSlideNum)) {
			if(this.getFirstSlideNumber() !== nFirstSlideNum) {
				if(nFirstSlideNum === 1) {
					this.setFirstSlideNum(null);
				}
				else {
					this.setFirstSlideNum(nFirstSlideNum);
				}
			}
		}
		this.Recalculate();
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.getViewProperties = function () {
	let oPresentation = this;
	return AscFormat.ExecuteNoHistory(function () {
		if (oPresentation.viewPr) {
			return oPresentation.viewPr.createDuplicate();
		}
		return new AscFormat.CViewPr();
	}, this, []);
};

CPresentation.prototype.getGridSpacing = function () {
	if (this.viewPr) {
		return this.viewPr.getGridSpacing();
	}
	return AscFormat.CViewPr.prototype.DEFAULT_GRID_SPACING;
};
CPresentation.prototype.getGridSpacingMM = function () {
	return this.getGridSpacing() / g_dKoef_mm_to_emu;
};
CPresentation.prototype.getViewPropertiesStride = function () {
	return this.getGridSpacing();
};
CPresentation.prototype.checkViewPr = function () {
	if (!this.viewPr) {
		this.setViewPr(new AscFormat.CViewPr());
	}
	return this.viewPr;
};
CPresentation.prototype.setGridSpacing = function (nSpacing) {
	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_ViewPr, undefined, undefined, [])) {
		this.Create_NewHistoryPoint(0);
		this.checkViewPr().setGridSpacingVal(nSpacing);
		this.Recalculate();
		this.UpdateInterface();
	}
};
CPresentation.prototype.isSnapToGrid = function () {
	if (this.viewPr) {
		return this.viewPr.isSnapToGrid();
	}
	return false;
};
CPresentation.prototype.setSnapToGrid = function (bVal) {
	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_ViewPr, undefined, undefined, [])) {
		this.Create_NewHistoryPoint(0);
		this.checkViewPr().setSnapToGrid(bVal);
		this.UpdateInterface();
	}
};

CPresentation.prototype.addHorizontalGuide = function () {
	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_ViewPr, undefined, undefined, [])) {
		this.Create_NewHistoryPoint(0);
		this.checkViewPr().addHorizontalGuide();
		this.Recalculate();
		this.UpdateInterface();
	}
};
CPresentation.prototype.addVerticalGuide = function () {
	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_ViewPr, undefined, undefined, [])) {
		this.Create_NewHistoryPoint(0);
		this.checkViewPr().addVerticalGuide();
		this.Recalculate();
		this.UpdateInterface();
	}
};


CPresentation.prototype.checkEmptyGuides = function () {
	if (!this.canClearGuides()) {
		if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_ViewPr, undefined, undefined, [])) {
			this.Create_NewHistoryPoint(0);
			this.checkViewPr().addVerticalGuide();
			this.checkViewPr().addHorizontalGuide();
			this.Recalculate();
			this.UpdateInterface();
		}
	}
};


//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с составным вводом
//----------------------------------------------------------------------------------------------------------------------
/**
 * Сообщаем о начале составного ввода текста.
 * @returns {boolean} Начался или нет составной ввод.
 */

CPresentation.prototype.IsThisElementCurrent = function () {
	return false;
};

CPresentation.prototype.TurnOffCheckChartSelection = function () {
};

CPresentation.prototype.TurnOnCheckChartSelection = function () {
};

CPresentation.prototype.setFirstSlideNum = function (val) {
	History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_Presentation_SetFirstSlideNum, this.firstSlideNum, val));
	this.firstSlideNum = val;
};
CPresentation.prototype.setShowSpecialPlsOnTitleSld = function (val) {
	History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_Presentation_SetShowSpecialPlsOnTitleSld, this.showSpecialPlsOnTitleSld, val));
	this.showSpecialPlsOnTitleSld = val;
};

CPresentation.prototype.setDefaultTextStyle = function (oStyle) {
	History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Presentation_SetDefaultTextStyle, this.defaultTextStyle, oStyle));
	this.defaultTextStyle = oStyle;
};


CPresentation.prototype.addSection = function (pos, pr) {
	History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_Presentation_AddSection, pos, [pr], true));
	this.Sections.splice(pos, 0, pr);
};

CPresentation.prototype.removeSection = function (pos) {
	History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_Presentation_AddSection, pos, [], true));
	this.Sections.splice(pos, 0);
};

CPresentation.prototype.SetDefaultLanguage = function (NewLangId) {
	this.SetLanguage(NewLangId);
	this.RestartSpellCheck();
	this.Recalculate();
	this.Document_UpdateInterfaceState();
};
CPresentation.prototype.SetLanguage = function (NewLangId) {
	var oTextStyle = this.defaultTextStyle ? this.defaultTextStyle.createDuplicate() : new AscFormat.TextListStyle();
	if (!oTextStyle.levels[9]) {
		oTextStyle.levels[9] = new CParaPr();
	}
	if (!oTextStyle.levels[9].DefaultRunPr) {
		oTextStyle.levels[9].DefaultRunPr = new CTextPr();
	}
	oTextStyle.levels[9].DefaultRunPr.Lang.Val = NewLangId;
	this.setDefaultTextStyle(oTextStyle);
};

CPresentation.prototype.GetDefaultLanguage = function () {
	let oTextPr = null;
	if (this.defaultTextStyle && this.defaultTextStyle.levels[9]) {
		oTextPr = this.defaultTextStyle.levels[9].DefaultRunPr;
	}
	const lcid_EnUS = 1033;
	let nLang = lcid_EnUS;
	if(oTextPr && oTextPr.Lang && AscFormat.isRealNumber(oTextPr.Lang.Val)) {
		nLang = oTextPr.Lang.Val;
	}
	return nLang;
};

CPresentation.prototype.collectHFProps = function (oSlide) {
	if (oSlide) {
		var oParentObjects = oSlide.getParentObjects();
		var oContent, sText, oField;
		var oSlideHF = new AscCommonSlide.CAscHFProps();
		var sFieldType, oDateTimeFieldsMap;
		oSlideHF.put_Api(this.Api);
		var oDTShape = oSlide.getMatchingShape(AscFormat.phType_dt, null, false, {});

		oSlideHF.put_ShowDateTime(false);
		if (oDTShape) {
			oSlideHF.put_ShowDateTime(true);
		}
		if (!oDTShape) {
			if (oParentObjects.layout) {
				oDTShape = oParentObjects.layout.getMatchingShape(AscFormat.phType_dt, null, false, {});
			}
		}
		if (!oDTShape) {
			if (oParentObjects.master) {
				oDTShape = oParentObjects.master.getMatchingShape(AscFormat.phType_dt, null, false, {});
			}
		}
		let oDateTime;
		if (oDTShape) {
			oContent = oDTShape.getDocContent();
			if (oContent && oContent.CalculateAllFields) {
				oDateTime = new AscCommonSlide.CAscDateTime();
				sText = "";
				if(oSlideHF.get_ShowDateTime()) {
					oContent.SetApplyToAll(true);
					sText = oContent.GetSelectedText(false);
					oContent.SetApplyToAll(false);
				}
				oDateTime.put_CustomDateTime(sText);
				oContent.CalculateAllFields();
				oField = oContent.GetFieldByType2('datetime');
				if (oField) {
					oDateTimeFieldsMap = {};
					oDateTimeFieldsMap["datetime"] =
						oDateTimeFieldsMap["datetime1"] =
							oDateTimeFieldsMap["datetime2"] =
								oDateTimeFieldsMap["datetime3"] =
									oDateTimeFieldsMap["datetime4"] =
										oDateTimeFieldsMap["datetime5"] =
											oDateTimeFieldsMap["datetime6"] =
												oDateTimeFieldsMap["datetime7"] =
													oDateTimeFieldsMap["datetime8"] =
														oDateTimeFieldsMap["datetime9"] =
															oDateTimeFieldsMap["datetime10"] =
																oDateTimeFieldsMap["datetime11"] =
																	oDateTimeFieldsMap["datetime12"] =
																		oDateTimeFieldsMap["datetime13"] = true;
					if (oDateTimeFieldsMap[oField.FieldType]) {
						sFieldType = oField.FieldType;
					} else {
						if(oField.FieldType === "datetimeFigureOut") {
							sFieldType = "datetime1";
						}
						else {
							sFieldType = "datetime";
						}
					}
					oDateTime.put_DateTime(sFieldType);

					oDateTime.put_Lang(oField.Pr.Lang.Val);
				}
			}
		}
		if(!oDateTime) {
			oDateTime = new AscCommonSlide.CAscDateTime();
			oDateTime.put_CustomDateTime("");
			oDateTime.put_DateTime("datetime");
			oDateTime.put_Lang(this.GetDefaultLanguage());
		}
		oSlideHF.put_DateTime(oDateTime);


		var oSldNumShape = oSlide.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
		if (oSldNumShape) {
			oSlideHF.put_ShowSlideNum(true);
		} else {
			oSlideHF.put_ShowSlideNum(false);
		}

		var oFooterShape = oSlide.getMatchingShape(AscFormat.phType_ftr, null, false, {});
		if (oFooterShape) {
			oSlideHF.put_ShowFooter(true);
		} else {
			oSlideHF.put_ShowFooter(false);
		}
		if (!oFooterShape) {
			if (oParentObjects.layout) {
				oFooterShape = oParentObjects.layout.getMatchingShape(AscFormat.phType_ftr, null, false, {});
			}
		}
		if (!oFooterShape) {
			if (oParentObjects.master) {
				oFooterShape = oParentObjects.master.getMatchingShape(AscFormat.phType_ftr, null, false, {});
			}
		}
		if (oFooterShape) {
			oContent = oFooterShape.getDocContent();
			if (oContent) {
				oContent.SetApplyToAll(true);
				sText = oContent.GetSelectedText(false);
				oContent.SetApplyToAll(false);
				oSlideHF.put_Footer(sText);
			}
		}

		var oHeaderShape = oSlide.getMatchingShape(AscFormat.phType_hdr, null, false, {});
		if (oHeaderShape) {
			oSlideHF.put_ShowHeader(true);
		} else {
			oSlideHF.put_ShowHeader(false);
		}
		if (!oHeaderShape) {
			if (oParentObjects.layout) {
				oHeaderShape = oParentObjects.layout.getMatchingShape(AscFormat.phType_hdr, null, false, {});
			}
		}
		if (!oHeaderShape) {
			if (oParentObjects.master) {
				oHeaderShape = oParentObjects.master.getMatchingShape(AscFormat.phType_hdr, null, false, {});
			}
		}
		if (oHeaderShape) {
			oContent = oHeaderShape.getDocContent();
			if (oContent) {
				oContent.SetApplyToAll(true);
				sText = oContent.GetSelectedText(false);
				oContent.SetApplyToAll(false);
				oSlideHF.put_Header(sText);
			}
		}
		oSlideHF.put_ShowOnTitleSlide(this.showSpecialPlsOnTitleSld !== false);
		return oSlideHF;
	}
	return null;
};

CPresentation.prototype.getHFProperties = function () {
	var oProps = new AscCommonSlide.CAscHF();
	var oSlide = this.GetCurrentSlide();
	oProps.put_Slide(this.collectHFProps(oSlide));
	if (oProps.Slide) {
		oProps.Slide.slide = oSlide;
	}
	let oNotes = this.GetCurrentNotes();
	if (oNotes) {
		oProps.put_Notes(this.collectHFProps(oNotes));
		if (oProps.Notes) {
			oProps.Notes.notes = oNotes;
		}
	}
	return oProps;
};


CPresentation.prototype.setHFProperties = function (oProps, bAll) {
	if (bAll && this.Document_Is_SelectionLocked(AscCommon.changestype_HdrFtr)) {
		return;
	}

	this.StartAction(AscDFH.historydescription_Presentation_SetHF);
	let oSlideProps = oProps.get_Slide();
	let oNotesProps = oProps.get_Notes();
	let i, j, oSlide, oMaster, oParents, oLayout, oSp,
		sText, oContent, oDateTime, sDateTime, sCustomDateTime, bRemoveOnTitle, nLang, aSelectedSlides, nSlideIndex;
	let oNotes;
	let oNotesMaster;
	let nLayout;
	let bRecalculate = false;
	const nDefaultLang = this.GetDefaultLanguage();
	const oPresentation = this;
	function fAddTextToPhInSlideLikeObject(oSlideLikeObject, nPhType, sText) {
		if(typeof sText === "string") {
			let oSp = oSlideLikeObject.getMatchingShape(nPhType, null, false, {});
			let oContent = oSp && oSp.getDocContent && oSp.getDocContent();
			if (oContent) {
				let oParaPr = null;
				let oFirstParaOld = oContent.GetFirstParagraph();
				if (oFirstParaOld) {
					oParaPr = oFirstParaOld.GetDirectParaPr();
				}
				AscFormat.CheckContentTextAndAdd(oContent, sText);
				if(oParaPr) {
					let oFirstParaNew = oContent.GetFirstParagraph();
					if (oFirstParaNew) {
						oFirstParaNew.SetPr(oParaPr);
					}
				}
			}
		}
	}
	function fAddFooterToSlideLikeObject(oSlideLikeObject, sFooterText) {
		fAddTextToPhInSlideLikeObject(oSlideLikeObject, AscFormat.phType_ftr, sFooterText);
	}
	function fAddHeaderToSlideLikeObject(oSlideLikeObject, sHeaderText) {
		fAddTextToPhInSlideLikeObject(oSlideLikeObject, AscFormat.phType_hdr, sHeaderText);
	}
	function fAddDateTimeToSlideLikeObject(oSlideLikeObject, sDateTimeFieldType, sCustomDateTime, nLang) {
		let oSp = oSlideLikeObject.getMatchingShape(AscFormat.phType_dt, null, false, {});
		if (oSp) {
			let oContent = oSp.getDocContent && oSp.getDocContent();
			if (oContent) {
				if (sDateTimeFieldType) {
					oContent.ClearContent(true);
					let oParagraph = oContent.Content[0];
					let oFld = new AscCommonWord.CPresentationField(oParagraph);
					oFld.SetGuid(AscCommon.CreateGUID());
					oFld.SetFieldType(sDateTimeFieldType);
					if(AscFormat.isRealNumber(nLang)) {
						oFld.Set_Lang_Val(nLang);
					}
					else {
						oFld.Set_Lang_Val(nDefaultLang);
					}
					if (typeof sCustomDateTime === "string") {
						oFld.CanAddToContent = true;
						oFld.AddText(sCustomDateTime);
						oFld.CanAddToContent = false;
					}
					oParagraph.Internal_Content_Add(0, oFld);
				} else {
					AscFormat.CheckContentTextAndAdd(oContent, sCustomDateTime);
				}
			}
		}
	}
	function fCopyPlaceholderToLikeObject(oSlide, oTemplate, nPhType) {
		let oSp = oTemplate.getMatchingShape(nPhType, null, false, {});
		if (oSp) {
			oSp = oSp.copy(undefined);
			oSp.clearLang();
			oSlide.addToSpTreeToPos(undefined, oSp);
			oSp.setParent(oSlide);
		}
	}
	function fApplyPropsToSlide(oSlide, oSlideProps) {
		if(!oSlide) return;
		let oParents = oSlide.getParentObjects();
		let oLayout = oParents.layout;
		let bRemoveOnTitle = oLayout.type === AscFormat.nSldLtTTitle && oPresentation.showSpecialPlsOnTitleSld === false;
		let oSp, sText, oDateTime, sDateTime, sCustomDateTime;
		if (oSlideProps.get_ShowSlideNum() && !bRemoveOnTitle) {
			if (!oSlide.getMatchingShape(AscFormat.phType_sldNum, null, false, {})) {
				fCopyPlaceholderToLikeObject(oSlide, oLayout, AscFormat.phType_sldNum);
			}
		} else {
			oSp = oSlide.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
			if (oSp) {
				oSlide.removeFromSpTreeById(oSp.Get_Id());
				oSp.setBDeleted(true);
			}
		}

		if (oSlideProps.get_ShowFooter() && !bRemoveOnTitle) {
			sText = oSlideProps.get_Footer();
			oSp = oSlide.getMatchingShape(AscFormat.phType_ftr, null, false, {});
			if (!oSp) {
				fCopyPlaceholderToLikeObject(oSlide, oLayout, AscFormat.phType_ftr);
			}
			fAddFooterToSlideLikeObject(oSlide, sText);
		} else {
			oSp = oSlide.getMatchingShape(AscFormat.phType_ftr, null, false, {});
			if (oSp) {
				oSlide.removeFromSpTreeById(oSp.Get_Id());
				oSp.setBDeleted(true);
			}
		}

		if (oSlideProps.get_ShowHeader() && !bRemoveOnTitle) {
			sText = oSlideProps.get_Header();
			oSp = oSlide.getMatchingShape(AscFormat.phType_hdr, null, false, {});
			if (!oSp) {
				fCopyPlaceholderToLikeObject(oSlide, oLayout, AscFormat.phType_hdr);
			}
			fAddHeaderToSlideLikeObject(oSlide, sText);
		} else {
			oSp = oSlide.getMatchingShape(AscFormat.phType_hdr, null, false, {});
			if (oSp) {
				oSlide.removeFromSpTreeById(oSp.Get_Id());
				oSp.setBDeleted(true);
			}
		}

		if (oSlideProps.get_ShowDateTime() && !bRemoveOnTitle) {
			oDateTime = oSlideProps.get_DateTime();
			sDateTime = "";
			sCustomDateTime = "";
			if (oDateTime) {
				sDateTime = oDateTime.get_DateTime();
				sCustomDateTime = oDateTime.get_CustomDateTime();
				if (sDateTime) {
					sCustomDateTime = oDateTime.get_DateTimeExamples()[sDateTime];
				}
			}
			oSp = oSlide.getMatchingShape(AscFormat.phType_dt, null, false, {});
			if (!oSp) {
				fCopyPlaceholderToLikeObject(oSlide, oLayout, AscFormat.phType_dt);
			}
			fAddDateTimeToSlideLikeObject(oSlide, sDateTime, sCustomDateTime, oDateTime.get_Lang());
		} else {
			oSp = oSlide.getMatchingShape(AscFormat.phType_dt, null, false, {});
			if (oSp) {
				oSlide.removeFromSpTreeById(oSp.Get_Id());
				oSp.setBDeleted(true);
			}
		}
	}

	if (oSlideProps) {
		let bShowOnTitleSlide = oSlideProps.get_ShowOnTitleSlide();
		if (bShowOnTitleSlide) {
			if (this.showSpecialPlsOnTitleSld !== null) {
				this.setShowSpecialPlsOnTitleSld(null);
			}
		} else {
			if (this.showSpecialPlsOnTitleSld !== false) {
				this.setShowSpecialPlsOnTitleSld(false);
			}
		}
		if(this.IsMasterMode()) {
			if(bAll) {
				for(let nMaster = 0; nMaster < this.slideMasters.length; ++nMaster) {
					oMaster = this.slideMasters[nMaster];
					if (!oMaster.hf) {
						oMaster.setHF(new AscFormat.HF());
					}
					oMaster.hf.applySettings(oSlideProps);
					if (oSlideProps.get_ShowFooter()) {
						sText = oSlideProps.get_Footer();
						if (typeof sText === "string") {
							for (j = 0; j < oMaster.sldLayoutLst.length; ++j) {
								fAddFooterToSlideLikeObject(oMaster.sldLayoutLst[j], sText);
							}
							fAddFooterToSlideLikeObject(oMaster, sText);
						}
					}
					if (oSlideProps.get_ShowHeader()) {
						sText = oSlideProps.get_Header();
						if (typeof sText === "string") {
							for (j = 0; j < oMaster.sldLayoutLst.length; ++j) {
								fAddHeaderToSlideLikeObject(oMaster.sldLayoutLst[j], sText);
							}
							fAddHeaderToSlideLikeObject(oMaster, sText);
						}
					}
					if (oSlideProps.get_ShowDateTime()) {
						oDateTime = oSlideProps.get_DateTime();

						sDateTime = "";
						sCustomDateTime = "";
						nLang = nDefaultLang;
						if (oDateTime) {
							sDateTime = oDateTime.get_DateTime();
							sCustomDateTime = oDateTime.get_CustomDateTime();
							nLang = oDateTime.get_Lang();
							if (typeof sDateTime === "string" || typeof sCustomDateTime === "string") {
								if (sDateTime) {
									sCustomDateTime = oDateTime.get_DateTimeExamples()[sDateTime];
								}
								for (j = 0; j < oMaster.sldLayoutLst.length; ++j) {
									fAddDateTimeToSlideLikeObject(oMaster.sldLayoutLst[j], sDateTime, sCustomDateTime, nLang);
								}
								fAddDateTimeToSlideLikeObject(oMaster, sDateTime, sCustomDateTime, nLang);
							}
						}
					}
					for (nLayout = 0; nLayout < oMaster.sldLayoutLst.length; ++nLayout) {
						oLayout = oMaster.sldLayoutLst[nLayout];
						if (oLayout.hf) {
							oLayout.setHF(null);
						}
					}
				}
				for(let nSlide = 0; nSlide < this.Slides.length; ++nSlide) {
					fApplyPropsToSlide(this.Slides[nSlide], oSlideProps);
				}
			}
			else {
				let aSelected = this.GetSelectedSlideObjects();
				let oDependentSlides = {};
				for(let nIdx = 0; nIdx < aSelected.length; ++nIdx) {
					let oSlideLikeObject = aSelected[nIdx];
					if (!oSlideLikeObject.hf) {
						oSlideLikeObject.setHF(new AscFormat.HF());
					}
					oSlideLikeObject.hf.applySettings(oSlideProps);
					if (oSlideProps.get_ShowFooter()) {
						sText = oSlideProps.get_Footer();
						if (typeof sText === "string") {
							fAddFooterToSlideLikeObject(oSlideLikeObject, sText);
						}
					}
					if (oSlideProps.get_ShowHeader()) {
						sText = oSlideProps.get_Header();
						if (typeof sText === "string") {
							fAddHeaderToSlideLikeObject(oSlideLikeObject, sText);
						}
					}
					if (oSlideProps.get_ShowDateTime()) {
						oDateTime = oSlideProps.get_DateTime();
						sDateTime = "";
						sCustomDateTime = "";
						if (oDateTime) {
							sDateTime = oDateTime.get_DateTime();
							sCustomDateTime = oDateTime.get_CustomDateTime();
							if (typeof sDateTime === "string" || typeof sCustomDateTime === "string") {
								if (sDateTime) {
									sCustomDateTime = oDateTime.get_DateTimeExamples()[sDateTime];
								}
								fAddDateTimeToSlideLikeObject(oSlideLikeObject, sDateTime, sCustomDateTime, oDateTime.get_Lang());
							}
						}
					}
					for(let nSlide = 0; nSlide < this.Slides.length; ++nSlide) {
						let oSlide = this.Slides[nSlide];
						let oParentObjects = oSlide.getParentObjects();
						if(oParentObjects.layout === oSlideLikeObject || oParentObjects.master === oSlideLikeObject) {
							oDependentSlides[oSlide.Id] = oSlide;
						}
					}
				}
				for(let sId in oDependentSlides) {
					if(oDependentSlides.hasOwnProperty(sId)) {
						let oSlide = oDependentSlides[sId];
						fApplyPropsToSlide(oSlide, oSlideProps);
					}
				}
			}
		}
		else {
			if (bAll) {
				var oMastersMap = {};
				for (i = 0; i < this.Slides.length; ++i) {
					oSlide = this.Slides[i];
					oParents = oSlide.getParentObjects();
					oMaster = oParents.master;
					oLayout = oParents.layout;
					bRemoveOnTitle = oLayout.type === AscFormat.nSldLtTTitle && this.showSpecialPlsOnTitleSld === false;
					if (oMaster) {
						if (!oMaster.hf) {
							oMaster.setHF(new AscFormat.HF());
						}
						oMaster.hf.applySettings(oSlideProps);
						if (oSlideProps.get_ShowSlideNum()) {
							oSp = oSlide.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
							if (!bRemoveOnTitle) {
								if (!oSp) {
									fCopyPlaceholderToLikeObject(oSlide, oLayout, AscFormat.phType_sldNum);
								}
							} else {
								if (oSp) {
									oSlide.removeFromSpTreeById(oSp.Get_Id());
									oSp.setBDeleted(true);
								}
							}
						} else {
							oSp = oSlide.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
							if (oSp) {
								oSlide.removeFromSpTreeById(oSp.Get_Id());
								oSp.setBDeleted(true);
							}
						}

						if (oSlideProps.get_ShowFooter()) {
							sText = oSlideProps.get_Footer();
							if (!oMastersMap[oMaster.Get_Id()]) {
								if (typeof sText === "string") {
									for (j = 0; j < oMaster.sldLayoutLst.length; ++j) {
										fAddFooterToSlideLikeObject(oMaster.sldLayoutLst[j], sText);
									}
									fAddFooterToSlideLikeObject(oMaster, sText);
								}
							}
							oSp = oSlide.getMatchingShape(AscFormat.phType_ftr, null, false, {});
							if (!bRemoveOnTitle) {
								if (!oSp) {
									fCopyPlaceholderToLikeObject(oSlide, oLayout, AscFormat.phType_ftr);
								} else {
									fAddFooterToSlideLikeObject(oSlide, sText);
								}
							} else {
								if (oSp) {
									oSlide.removeFromSpTreeById(oSp.Get_Id());
									oSp.setBDeleted(true);
								}
							}
						} else {
							oSp = oSlide.getMatchingShape(AscFormat.phType_ftr, null, false, {});
							if (oSp) {
								oSlide.removeFromSpTreeById(oSp.Get_Id());
								oSp.setBDeleted(true);
							}
						}

						if (oSlideProps.get_ShowHeader()) {
							sText = oSlideProps.get_Header();
							if (!oMastersMap[oMaster.Get_Id()]) {
								if (typeof sText === "string") {
									for (j = 0; j < oMaster.sldLayoutLst.length; ++j) {
										fAddHeaderToSlideLikeObject(oMaster.sldLayoutLst[j], sText);
									}
									fAddHeaderToSlideLikeObject(oMaster, sText);
								}
							}


							oSp = oSlide.getMatchingShape(AscFormat.phType_hdr, null, false, {});
							if (!bRemoveOnTitle) {
								if (!oSp) {
									fCopyPlaceholderToLikeObject(oSlide, oLayout, AscFormat.phType_hdr);
								} else {
									fAddHeaderToSlideLikeObject(oSlide, sText);
								}
							} else {
								if (oSp) {
									oSlide.removeFromSpTreeById(oSp.Get_Id());
									oSp.setBDeleted(true);
								}
							}
						} else {
							oSp = oSlide.getMatchingShape(AscFormat.phType_hdr, null, false, {});
							if (oSp) {
								oSlide.removeFromSpTreeById(oSp.Get_Id());
								oSp.setBDeleted(true);
							}
						}

						if (oSlideProps.get_ShowDateTime()) {
							oDateTime = oSlideProps.get_DateTime();

							sDateTime = "";
							sCustomDateTime = "";
							nLang = nDefaultLang;
							if (oDateTime) {
								sDateTime = oDateTime.get_DateTime();
								sCustomDateTime = oDateTime.get_CustomDateTime();
								nLang = oDateTime.get_Lang();
								if (!oMastersMap[oMaster.Get_Id()]) {
									if (typeof sDateTime === "string" || typeof sCustomDateTime === "string") {
										if (sDateTime) {
											sCustomDateTime = oDateTime.get_DateTimeExamples()[sDateTime];
										}
										for (j = 0; j < oMaster.sldLayoutLst.length; ++j) {
											fAddDateTimeToSlideLikeObject(oMaster.sldLayoutLst[j], sDateTime, sCustomDateTime, nLang);
										}
										fAddDateTimeToSlideLikeObject(oMaster, sDateTime, sCustomDateTime, nLang);
										oSp = oMaster.getMatchingShape(AscFormat.phType_dt, null, false, {});
									}
								}
							}
							oSp = oSlide.getMatchingShape(AscFormat.phType_dt, null, false, {});
							if (!bRemoveOnTitle) {
								if (!oSp) {
									fCopyPlaceholderToLikeObject(oSlide, oLayout, AscFormat.phType_dt);
								} else {
									fAddDateTimeToSlideLikeObject(oSlide, sDateTime, sCustomDateTime, oDateTime.get_Lang());
								}
							} else {
								if (oSp) {
									oSlide.removeFromSpTreeById(oSp.Get_Id());
									oSp.setBDeleted(true);
								}
							}
						} else {
							oSp = oSlide.getMatchingShape(AscFormat.phType_dt, null, false, {});
							if (oSp) {
								oSlide.removeFromSpTreeById(oSp.Get_Id());
								oSp.setBDeleted(true);
							}
						}

						if (!oMastersMap[oMaster.Get_Id()]) {
							for (nLayout = 0; nLayout < oMaster.sldLayoutLst.length; ++nLayout) {
								oLayout = oMaster.sldLayoutLst[nLayout];
								if (oLayout.hf) {
									oLayout.setHF(null);
								}
							}
						}
						oMastersMap[oMaster.Get_Id()] = oMaster;
					}
				}
			}
			else {
				aSelectedSlides = this.GetSelectedSlides();
				for (nSlideIndex = 0; nSlideIndex < aSelectedSlides.length; ++nSlideIndex) {
					oSlide = this.GetSlide(aSelectedSlides[nSlideIndex]);
					fApplyPropsToSlide(oSlide, oSlideProps);
				}
			}
		}
		bRecalculate = true;
	}
	if (oNotesProps) {
		if (bAll) {
			const oNotesMastersMap = {};
			for (let nSlide = 0; nSlide < this.Slides.length; ++nSlide) {
				oSlide = this.Slides[nSlide];
				oNotes = oSlide.notes;
				if(!oNotes) {
					continue;
				}
				oNotesMaster = oNotes.Master;
				if(!oNotesMaster) {
					continue;
				}
				if (!oNotesMaster.hf) {
					oNotesMaster.setHF(new AscFormat.HF());
				}
				oNotesMaster.hf.applySettings(oNotesProps);
				if (oNotesProps.get_ShowSlideNum()) {
					oSp = oNotes.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
					if (!oSp) {
						oSp = oNotesMaster.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
						if (oSp) {
							oSp = oSp.copy(undefined);
							oSp.clearLang();
							oNotes.addToSpTreeToPos(undefined, oSp);
							oSp.setParent(oNotes);
						}
					}
				} else {
					oSp = oNotes.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
					if (oSp) {
						oNotes.removeFromSpTreeById(oSp.Get_Id());
						oSp.setBDeleted(true);
					}
				}

				if (oNotesProps.get_ShowFooter()) {
					sText = oNotesProps.get_Footer();
					if (!oNotesMastersMap[oNotesMaster.Get_Id()]) {
						if (typeof sText === "string") {
							oSp = oNotesMaster.getMatchingShape(AscFormat.phType_ftr, null, false, {});
							oContent = oSp && oSp.getDocContent && oSp.getDocContent();
							if (oContent) {
								AscFormat.CheckContentTextAndAdd(oContent, sText);
							}
						}
					}
					oSp = oNotes.getMatchingShape(AscFormat.phType_ftr, null, false, {});
					if (!oSp) {
						oSp = oNotesMaster.getMatchingShape(AscFormat.phType_ftr, null, false, {});
						if (oSp) {
							oSp = oSp.copy(undefined);
							oSp.clearLang();
							oNotes.addToSpTreeToPos(undefined, oSp);
							oSp.setParent(oNotes);
						}
					} else {
						oContent = oSp.getDocContent && oSp.getDocContent();
						if (oContent && typeof sText === "string") {
							AscFormat.CheckContentTextAndAdd(oContent, sText);
						}
					}
				} else {
					oSp = oNotes.getMatchingShape(AscFormat.phType_ftr, null, false, {});
					if (oSp) {
						oNotes.removeFromSpTreeById(oSp.Get_Id());
						oSp.setBDeleted(true);
					}
				}

				if (oNotesProps.get_ShowHeader()) {
					sText = oNotesProps.get_Header();
					if (!oNotesMastersMap[oNotesMaster.Get_Id()]) {
						fAddHeaderToSlideLikeObject(oNotesMaster, sText);
					}


					oSp = oNotes.getMatchingShape(AscFormat.phType_hdr, null, false, {});
					if (!oSp) {
						oSp = oNotesMaster.getMatchingShape(AscFormat.phType_hdr, null, false, {});
						if (oSp) {
							oSp = oSp.copy(undefined);
							oSp.clearLang();
							oNotes.addToSpTreeToPos(undefined, oSp);
							oSp.setParent(oNotes);
						}
					} else {
						fAddHeaderToSlideLikeObject(oNotes, sText);
					}
				} else {
					oSp = oNotes.getMatchingShape(AscFormat.phType_hdr, null, false, {});
					if (oSp) {
						oNotes.removeFromSpTreeById(oSp.Get_Id());
						oSp.setBDeleted(true);
					}
				}


				if (oNotesProps.get_ShowDateTime()) {
					oDateTime = oNotesProps.get_DateTime();

					sDateTime = "";
					sCustomDateTime = "";
					nLang = nDefaultLang;
					if (oDateTime) {
						sDateTime = oDateTime.get_DateTime();
						sCustomDateTime = oDateTime.get_CustomDateTime();
						nLang = oDateTime.get_Lang();
						if (!oNotesMastersMap[oNotesMaster.Get_Id()]) {
							if (typeof sDateTime === "string" || typeof sCustomDateTime === "string") {
								if (sDateTime) {
									sCustomDateTime = oDateTime.get_DateTimeExamples()[sDateTime];
								}
								fAddDateTimeToSlideLikeObject(oNotesMaster, sDateTime, sCustomDateTime, nLang);
							}
						}
					}
					oSp = oNotes.getMatchingShape(AscFormat.phType_dt, null, false, {});

					if (!oSp) {
						oSp = oNotesMaster.getMatchingShape(AscFormat.phType_dt, null, false, {});
						if (oSp) {
							oSp = oSp.copy(undefined);
							oSp.clearLang();
							oNotes.addToSpTreeToPos(undefined, oSp);
							oSp.setParent(oNotes);
						}
					} else {
						fAddDateTimeToSlideLikeObject(oNotes, sDateTime, sCustomDateTime, nLang);
					}
				} else {
					oSp = oNotes.getMatchingShape(AscFormat.phType_dt, null, false, {});
					if (oSp) {
						oNotes.removeFromSpTreeById(oSp.Get_Id());
						oSp.setBDeleted(true);
					}
				}
				oNotesMastersMap[oNotesMaster.Get_Id()] = oNotesMaster;
			}
		} else {
			aSelectedSlides = this.GetSelectedSlides();
			for (nSlideIndex = 0; nSlideIndex < aSelectedSlides.length; ++nSlideIndex) {
				oSlide = this.GetSlide(aSelectedSlides[nSlideIndex]);
				if(!oSlide) {
					continue;
				}
				oNotes = oSlide.notes;
				if(!oNotes) {
					continue;
				}
				oNotesMaster = oNotes.Master;
				if(!oNotesMaster) {
					continue;
				}
				if (oNotesProps.get_ShowSlideNum()) {
					if (!oNotes.getMatchingShape(AscFormat.phType_sldNum, null, false, {})) {
						oSp = oNotesMaster.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
						if (oSp) {
							oSp = oSp.copy(undefined);
							oSp.clearLang();
							oNotes.addToSpTreeToPos(undefined, oSp);
							oSp.setParent(oNotes);
						}
					}
				} else {
					oSp = oNotes.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
					if (oSp) {
						oNotes.removeFromSpTreeById(oSp.Get_Id());
						oSp.setBDeleted(true);
					}
				}

				if (oNotesProps.get_ShowFooter()) {
					sText = oNotesProps.get_Footer();
					oSp = oNotes.getMatchingShape(AscFormat.phType_ftr, null, false, {});
					if (!oSp) {
						oSp = oNotesMaster.getMatchingShape(AscFormat.phType_ftr, null, false, {});
						if (oSp) {
							oSp = oSp.copy(undefined);
							oSp.clearLang();
							oNotes.addToSpTreeToPos(undefined, oSp);
							oSp.setParent(oNotes);
						}
					}
					if (oSp) {
						oContent = oSp.getDocContent && oSp.getDocContent();
						if (oContent && typeof sText === "string") {
							AscFormat.CheckContentTextAndAdd(oContent, sText);
						}
					}
				} else {
					oSp = oNotes.getMatchingShape(AscFormat.phType_ftr, null, false, {});
					if (oSp) {
						oNotes.removeFromSpTreeById(oSp.Get_Id());
						oSp.setBDeleted(true);
					}
				}

				if (oNotesProps.get_ShowHeader()) {
					sText = oNotesProps.get_Header();
					oSp = oNotes.getMatchingShape(AscFormat.phType_hdr, null, false, {});
					if (!oSp) {
						oSp = oNotesMaster.getMatchingShape(AscFormat.phType_hdr, null, false, {});
						if (oSp) {
							oSp = oSp.copy(undefined);
							oSp.clearLang();
							oNotes.addToSpTreeToPos(undefined, oSp);
							oSp.setParent(oNotes);
						}
					}
					fAddHeaderToSlideLikeObject(oNotes, sText);
				} else {
					oSp = oNotes.getMatchingShape(AscFormat.phType_hdr, null, false, {});
					if (oSp) {
						oNotes.removeFromSpTreeById(oSp.Get_Id());
						oSp.setBDeleted(true);
					}
				}

				if (oNotesProps.get_ShowDateTime()) {
					oDateTime = oNotesProps.get_DateTime();
					sDateTime = "";
					sCustomDateTime = "";
					nLang = nDefaultLang;
					if (oDateTime) {
						sDateTime = oDateTime.get_DateTime();
						sCustomDateTime = oDateTime.get_CustomDateTime();
						if (sDateTime) {
							sCustomDateTime = oDateTime.get_DateTimeExamples()[sDateTime];
						}
						nLang = oDateTime.get_Lang();
						if (!AscFormat.isRealNumber(nLang)) {
							nLang = nDefaultLang;
						}
					}
					oSp = oNotes.getMatchingShape(AscFormat.phType_dt, null, false, {});
					if (!oSp) {
						oSp = oNotesMaster.getMatchingShape(AscFormat.phType_dt, null, false, {});
						if (oSp) {
							oSp = oSp.copy(undefined);
							oSp.clearLang();
							oNotes.addToSpTreeToPos(undefined, oSp);
							oSp.setParent(oNotes);
						}
					}
					fAddDateTimeToSlideLikeObject(oNotes, sDateTime, sCustomDateTime, nLang);
				} else {
					oSp = oNotes.getMatchingShape(AscFormat.phType_dt, null, false, {});
					if (oSp) {
						oNotes.removeFromSpTreeById(oSp.Get_Id());
						oSp.setBDeleted(true);
					}
				}
			}
		}
		bRecalculate = true;
	}
	if(bRecalculate) {
		this.Recalculate();
		this.Document_UpdateSelectionState();
		this.Document_UpdateUndoRedoState();
		this.Document_UpdateInterfaceState();
		this.Document_UpdateRulersState();
	}
	this.FinalizeAction(true);
};


CPresentation.prototype.addFieldToContent = function (fCallback) {
	var oController = this.GetCurrentController();
	if (!oController) {
		return;
	}
	var oContent = oController.getTargetDocContent(undefined, false);
	if (!oContent) {
		return;
	}
	if (false === this.Document_Is_SelectionLocked(changestype_Drawing_Props)) {
		this.StartAction(AscDFH.historydescription_Presentation_AddSlideNumber);
		oContent = oController.getTargetDocContent(true, false);
		if (oContent) {
			if (true === oContent.IsSelectionUse())
				oContent.Remove(1, true, false, true);
			var oParagraph = oContent.Content[oContent.CurPos.ContentPos];
			if (oParagraph) {
				var oFld = fCallback.call(this, oParagraph);
				if (oFld) {
					oContent.AddToParagraph(oFld, false, false);
					oController.checkCurrentTextObjectExtends();
					oContent.MoveCursorRight(false, false);
					this.Recalculate();
					this.RecalculateCurPos();
					this.Document_UpdateSelectionState();
					this.Document_UpdateInterfaceState();
				}
			}
		}
		this.FinalizeAction(true);
	}
};
CPresentation.prototype.getFirstSlideNumber = function() {
	return AscFormat.isRealNumber(this.firstSlideNum) ? this.firstSlideNum : 1;
};
CPresentation.prototype.addSlideNumber = function () {
	this.addFieldToContent(function (oParagraph) {
		var oFld = new AscCommonWord.CPresentationField(oParagraph);
		oFld.SetGuid(AscCommon.CreateGUID());
		oFld.SetFieldType("slidenum");
		var nFirstSlideNum = this.getFirstSlideNumber();
		oFld.AddText("" + (this.CurPage + nFirstSlideNum));
		return oFld;
	});
};

CPresentation.prototype.addDateTime = function (oPr) {
	this.addFieldToContent(function (oParagraph) {
		var oFld = null;
		if (oPr) {
			var sCustomDateTime = oPr.get_CustomDateTime();
			var sFieldType = oPr.get_DateTime();
			var nLang = oPr.get_Lang();
			if (!AscFormat.isRealNumber(nLang)) {
				nLang = this.GetDefaultLanguage();
			}
			if (typeof sFieldType === "string" && sFieldType.length > 0) {
				oFld = new AscCommonWord.CPresentationField(oParagraph);
				oFld.SetGuid(AscCommon.CreateGUID());
				oFld.SetFieldType(sFieldType);
				if (sFieldType) {
					sCustomDateTime = oPr.get_DateTimeExamples()[sFieldType];
				}
				oFld.Set_Lang_Val(nLang);
				if (typeof sCustomDateTime === "string" && sCustomDateTime.length > 0) {
					oFld.CanAddToContent = true;
					oFld.AddText(sCustomDateTime);
					oFld.CanAddToContent = false;
				}
			} else {
				if (typeof sCustomDateTime === "string" && sCustomDateTime.length > 0) {
					oFld = new AscCommonWord.ParaRun(oParagraph);
					oFld.AddText(sCustomDateTime);
					oFld.Set_Lang_Val(nLang);
				}
			}
		}
		return oFld;
	});
};

CPresentation.prototype.RestartSpellCheck = function () {
	this.Spelling.Reset();
	let nCount = this.GetSlidesCount()
	for (let i = 0; i < nCount; ++i) {
		this.GetSlide(i).RestartSpellCheck();
	}
};

CPresentation.prototype.GetSelectionBounds = function () {
	var oController = this.GetCurrentController();
	if (oController) {
		var oTargetDocContent = oController.getTargetDocContent();
		if (oTargetDocContent) {
			return oTargetDocContent.GetSelectionBounds();
		}
	}
	return null;
};


CPresentation.prototype.GetTextTransformMatrix = function () {
	var oController = this.GetCurrentController();
	if (oController) {
		var oTargetDocContent = oController.getTargetDocContent();
		if (oTargetDocContent) {
			return oTargetDocContent.Get_ParentTextTransform();
		}
	}
	return null;
};


CPresentation.prototype.IsViewMode = function () {
	return this.Api.getViewMode();
};


CPresentation.prototype.IsEditCommentsMode = function () {
	return this.Api.isRestrictionComments();
};

CPresentation.prototype.IsEditSignaturesMode = function () {
	return this.Api.isRestrictionSignatures();
};
CPresentation.prototype.IsViewModeInEditor = function () {
	return this.Api.isRestrictionView();
};

CPresentation.prototype.CanEdit = function () {
	if(this.IsSlideShow()) {
		return false;
	}
	return this.Api.canEdit();
};


CPresentation.prototype.StopSpellCheck = function () {
	this.Spelling.Reset();
};

CPresentation.prototype.ContinueSpellCheck = function () {
	this.Spelling.ContinueSpellCheck();
};

CPresentation.prototype.TurnOffSpellCheck = function () {
	this.Spelling.TurnOff();
};

CPresentation.prototype.TurnOnSpellCheck = function () {
	this.Spelling.TurnOn();
};
CPresentation.prototype.GetSpellCheckManager = function () {
	return this.Spelling;
};

CPresentation.prototype.Get_DrawingDocument = function () {
	return this.DrawingDocument;
};
CPresentation.prototype.GetDrawingDocument = function () {
	return this.DrawingDocument;
};



CPresentation.prototype.GetSlides = function () {
	return this.Slides;
};
CPresentation.prototype.GetAllSlides = function () {
	if(this.IsMasterMode()) {
		let aSlides = [];
		for(let nMaster = 0; nMaster < this.slideMasters.length; ++nMaster) {
			let oMaster = this.slideMasters[nMaster];
			aSlides.push(oMaster);
			let aLayouts = oMaster.sldLayoutLst;
			for(let nLt = 0; nLt < aLayouts.length; ++nLt) {
				aSlides.push(aLayouts[nLt]);
			}
		}
		return aSlides;
	}
	return this.GetSlides();
};
CPresentation.prototype.GetSlide = function (nIndex) {
	let aSlides = this.GetAllSlides();
	if (aSlides[nIndex]) {
		return aSlides[nIndex];
	}
	return null;
};
CPresentation.prototype.GetSlideType = function (nIndex) {
	let oSlide = this.GetSlide(nIndex);
	if(oSlide) return oSlide.getObjectType();
	return null;
};
CPresentation.prototype.GetFirstSelectedType = function () {
	return this.GetSlideType(this.GetSelectedSlides()[0]);
};
CPresentation.prototype.IsMixedSelection = function()
{
	let aSelectedIdx = this.GetSelectedSlides();
	return this.CheckIsMixedSelection(aSelectedIdx);
};
CPresentation.prototype.CheckIsMixedSelection = function(aSelectedIdx)
{
	if(aSelectedIdx.length === 0) return false;
	let nType = this.GetSlideType(aSelectedIdx[0]);
	for(let nIdx = 0; nIdx < aSelectedIdx.length; ++nIdx)
	{
		if(this.GetSlideType(aSelectedIdx[nIdx]) !== nType)
			return true;
	}
	return false;
};
CPresentation.prototype.GetSlidesCount = function () {
	return this.GetAllSlides().length;
};
CPresentation.prototype.GetSlideIndex = function (oSlide) {
	let aSlides = this.GetAllSlides();
	for(let nIdx = 0; nIdx < aSlides.length; ++nIdx) {
		if(aSlides[nIdx] === oSlide) {
			return nIdx;
		}
	}
	return -1;
};
CPresentation.prototype.GetThumbnailsCount = function() {

	if (this.Api.WordControl.Thumbnails) {
		return this.Api.WordControl.Thumbnails.m_arrPages.length;
	}
	return this.GetSlidesCount();
};
CPresentation.prototype.GetSlideNumber = function (nIdx) {
	if(!this.IsMasterMode()) {
		return nIdx + this.getFirstSlideNumber();
	}
	let oSlide = this.GetSlide(nIdx);
	if(oSlide.getObjectType() === AscDFH.historyitem_type_SlideMaster) {
		for(let nMaster = 0; nMaster < this.slideMasters.length; ++nMaster) {
			if(this.slideMasters[nMaster] === oSlide) {
				return nMaster + 1;
			}
		}
	}
	return null;
};
CPresentation.prototype.CorrectCurSlideIdx = function () {
	let aAllSlides = this.GetAllSlides();
	if(this.CurPage >= aAllSlides.length) {
		this.Set_CurPage(aAllSlides.length - 1);
	}
	if(this.CurPage < 0 && aAllSlides.length > 0) {
		this.Set_CurPage(0);
	}
};
CPresentation.prototype.GetCurrentSlide = function () {
	return this.GetAllSlides()[this.CurPage];
};
CPresentation.prototype.GetCurrentMaster = function () {
	let oSlide = this.GetCurrentSlide();
	if(!oSlide) return this.lastMaster;
	let oParents = oSlide.getParentObjects();
	return oParents.master;
};
CPresentation.prototype.GetUsedMasters = function () {
	const mapMasters = {};
	const arrMasters = [];
	for (let i = 0; i < this.Slides.length; i++) {
		const oSlide = this.Slides[i];
		const oMaster = oSlide.Layout && oSlide.Layout.Master;
		if (oMaster && !mapMasters[oMaster.Id]) {
			arrMasters.push(oMaster);
			mapMasters[oMaster.Id] = true;
		}
	}
	return arrMasters;
};
CPresentation.prototype.GetCurrentNotes = function () {
	let oCurSlide = this.GetCurrentSlide();
	if(!oCurSlide) return null;
	if(!this.IsMasterMode()) {
		return oCurSlide.notes;
	}
	else {
		let oSlide = this.Slides[0];
		if(!oSlide) return null;
		return oSlide.notes;
	}
};
CPresentation.prototype.GetCurrentController = function () {
	var oCurSlide = this.GetCurrentSlide();
	if (oCurSlide) {
		if (this.IsFocusOnNotes()) {
			return oCurSlide.notes && oCurSlide.notes.graphicObjects;
		} else {
			return oCurSlide.graphicObjects;
		}
	}
	return null;
};
CPresentation.prototype.GetInputLanguage = function () {
	var oController = this.GetCurrentController();
	if (oController) {
		return oController.getInputLanguage();
	}
	return lcid_enUS;

};
CPresentation.prototype.GetCurrentSlideController = function () {
	var oCurSlide = this.GetCurrentSlide();
	if (oCurSlide) {
		return oCurSlide.graphicObjects;
	}
	return null;
};
CPresentation.prototype.IsEmpty = function() {
	return this.GetSlidesCount() === 0;
};
CPresentation.prototype.Get_TargetDocContent = function () {
	var oController = this.GetCurrentController();
	if (oController) {
		return oController.getTargetDocContent(true);
	}
	return null;
};

CPresentation.prototype.Begin_CompositeInput = function () {
	let oCurSlide = this.GetCurrentSlide();
	if (!this.IsFocusOnNotes() && oCurSlide && oCurSlide.graphicObjects.selectedObjects.length === 0) {
		var oTitle = oCurSlide.getMatchingShape(AscFormat.phType_title, null);
		if (oTitle) {
			var oDocContent = oTitle.getDocContent();
			if (oDocContent.Is_Empty()) {
				oDocContent.Set_CurrentElement(0, false);
			} else {
				return;
			}
		} else {
			return;
		}
	}
	if (false === this.Document_Is_SelectionLocked(changestype_Drawing_Props, null, undefined, undefined, true)) {
		this.Create_NewHistoryPoint(AscDFH.historydescription_Document_CompositeInput);
		var oController = this.GetCurrentController();
		if (oController) {
			oController.CreateDocContent();
		}


		var oContent = this.Get_TargetDocContent();
		if (!oContent) {
			this.History.Remove_LastPoint();
			return false;
		}
		this.DrawingDocument.TargetStart();
		this.DrawingDocument.TargetShow();
		var oPara = oContent.GetCurrentParagraph();
		if (!oPara) {
			this.History.Remove_LastPoint();
			return false;
		}
		if (true === oContent.IsSelectionUse())
			oContent.Remove(1, true, false, true);
		var oRun = oPara.Get_ElementByPos(oPara.Get_ParaContentPos(false, false));
		if (!oRun || !(oRun instanceof ParaRun)) {
			this.History.Remove_LastPoint();
			return false;
		}

		this.CompositeInput = new AscWord.RunCompositeInput_Old(oRun);
		oRun.Set_CompositeInput(this.CompositeInput);

		return true;
	}

	return false;
};

CPresentation.prototype.IsFillingFormMode = function () {
	return false;
};

CPresentation.prototype.ResetWordSelection = function () {
	this.WordSelected = false;
};

CPresentation.prototype.SetWordSelection = function (isWord) {
	this.WordSelected = isWord;
};

CPresentation.prototype.IsWordSelection = function () {
	return this.WordSelected;
};


CPresentation.prototype.checkCurrentTextObjectExtends = function () {
	var oController = this.GetCurrentController();
	if (oController) {
		var oTargetTextObject = AscFormat.getTargetTextObject(oController);
		if (oTargetTextObject && oTargetTextObject.checkExtentsByDocContent) {
			oTargetTextObject.checkExtentsByDocContent(true, true);
		}
	}
};

CPresentation.prototype.addCompositeText = function (nCharCode) {
	// TODO: При таком вводе не меняется язык в зависимости от раскладки, не учитывается режим рецензирования.

	if (null === this.CompositeInput)
		return;

	var oRun = this.CompositeInput.Run;
	var nPos = this.CompositeInput.Pos + this.CompositeInput.Length;
	var oChar;
	if (para_Math_Run === oRun.Type) {
		oChar = new CMathText();
		oChar.add(nCharCode);
	} else {
		if (32 == nCharCode || 12288 == nCharCode)
			oChar = new AscWord.CRunSpace();
		else
			oChar = new AscWord.CRunText(nCharCode);
	}
	oRun.AddToContent(nPos, oChar, true);
	this.CompositeInput.Length++;
};
CPresentation.prototype.Add_CompositeText = function (nCharCode) {
	if (null === this.CompositeInput)
		return;
	this.Create_NewHistoryPoint(AscDFH.historydescription_Document_CompositeInputReplace);
	this.addCompositeText(nCharCode);
	this.checkCurrentTextObjectExtends();
	this.Recalculate();
	this.RecalculateCurPos(true, true);
	this.Document_UpdateSelectionState();
};

CPresentation.prototype.removeCompositeText = function (nCount) {
	if (null === this.CompositeInput)
		return;

	var oRun = this.CompositeInput.Run;
	var nPos = this.CompositeInput.Pos + this.CompositeInput.Length;

	var nDelCount = Math.max(0, Math.min(nCount, this.CompositeInput.Length, oRun.Content.length, nPos));
	oRun.Remove_FromContent(nPos - nDelCount, nDelCount, true);
	this.CompositeInput.Length -= nDelCount;
};

CPresentation.prototype.Remove_CompositeText = function (nCount) {
	this.removeCompositeText(nCount);
	this.checkCurrentTextObjectExtends();
	this.Recalculate();
	this.RecalculateCurPos(true, true);
	this.Document_UpdateSelectionState();
};
CPresentation.prototype.Replace_CompositeText = function (arrCharCodes) {
	if (null === this.CompositeInput)
		return;
	this.Create_NewHistoryPoint(AscDFH.historydescription_Document_CompositeInputReplace);
	this.removeCompositeText(this.CompositeInput.Length);
	for (var nIndex = 0, nCount = arrCharCodes.length; nIndex < nCount; ++nIndex) {
		this.addCompositeText(arrCharCodes[nIndex]);
	}
	this.checkCurrentTextObjectExtends();
	this.Recalculate();
	this.RecalculateCurPos(true, true);
	this.Document_UpdateSelectionState();
	if (!this.History.CheckUnionLastPoints())
		this.CompositeInput.CanUndo = false;
};
CPresentation.prototype.Set_CursorPosInCompositeText = function (nPos) {
	if (null === this.CompositeInput)
		return;

	var oRun = this.CompositeInput.Run;

	var nInRunPos = Math.max(Math.min(this.CompositeInput.Pos + nPos, this.CompositeInput.Pos + this.CompositeInput.Length, oRun.Content.length), this.CompositeInput.Pos);
	oRun.State.ContentPos = nInRunPos;
	this.RecalculateCurPos(true, true);
	this.Document_UpdateSelectionState();
};
CPresentation.prototype.Get_CursorPosInCompositeText = function () {
	if (null === this.CompositeInput)
		return 0;

	var oRun = this.CompositeInput.Run;
	var nInRunPos = oRun.State.ContentPos;
	var nPos = Math.min(this.CompositeInput.Length, Math.max(0, nInRunPos - this.CompositeInput.Pos));
	return nPos;
};
CPresentation.prototype.End_CompositeInput = function () {
	if (null === this.CompositeInput)
		return;

	var nLen = this.CompositeInput.Length;

	var oRun = this.CompositeInput.Run;
	oRun.Set_CompositeInput(null);

	if (0 === nLen && true === this.History.CanRemoveLastPoint() && true === this.CompositeInput.CanUndo) {
		this.Document_Undo();
		this.History.Clear_Redo();
	}

	this.CompositeInput = null;

	var oController = this.GetCurrentController();
	if (oController) {
		var oTargetTextObject = AscFormat.getTargetTextObject(oController);
		if (oTargetTextObject && oTargetTextObject.txWarpStructNoTransform) {
			oTargetTextObject.recalculateContent();
		}
	}
	this.Document_UpdateInterfaceState();

	this.DrawingDocument.ClearCachePages();
	this.DrawingDocument.FirePaint();
};
CPresentation.prototype.Get_MaxCursorPosInCompositeText = function () {
	if (null === this.CompositeInput)
		return 0;

	return this.CompositeInput.Length;
};


CPresentation.prototype.setShowLoop = function (value) {
	if (value === false) {
		if (!this.showPr) {

		} else {
			if (this.showPr.loop !== false) {
				var oCopyShowPr = this.showPr.Copy();
				oCopyShowPr.loop = false;
				this.setShowPr(oCopyShowPr);
			}
		}
	} else {
		if (!this.showPr) {
			var oShowPr = new CShowPr();
			oShowPr.loop = true;
			this.setShowPr(oShowPr);
		} else {
			if (!this.showPr.loop) {
				var oCopyShowPr = this.showPr.Copy();
				oCopyShowPr.loop = true;
				this.setShowPr(oCopyShowPr);
			}
		}
	}
};

CPresentation.prototype.isLoopShowMode = function () {
	if (this.showPr) {
		return this.showPr.loop === true;
	}
	return false;
};


CPresentation.prototype.setShowPr = function (oShowPr) {
	History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Presentation_SetShowPr, this.showPr, oShowPr));
	this.showPr = oShowPr;
};

CPresentation.prototype.createDefaultTableStyles = function () {
	//AscFormat.ExecuteNoHistory(function(){
	this.globalTableStyles = new CStyles(false);

	this.globalTableStyles.Id = AscCommon.g_oIdCounter.Get_NewId();
	AscCommon.g_oTableId.Add(this.globalTableStyles, this.globalTableStyles.Id);
	this.DefaultTableStyleId = AscFormat.CreatePresentationTableStyles(this.globalTableStyles, this.TableStylesIdMap);
	//}, this, []);
};

CPresentation.prototype.GetAllTableStyles = function () {
	if (!this.globalTableStyles) {
		return [];
	}
	let aStyles = [];
	for (let key in this.TableStylesIdMap) {
		if (this.TableStylesIdMap.hasOwnProperty(key)) {
			let oStyle = this.globalTableStyles.Get(key);
			if (oStyle) {
				aStyles.push(oStyle);
			}
		}
	}
	return aStyles;
};
// Проводим начальные действия, исходя из Документа
CPresentation.prototype.Init = function () {

};

CPresentation.prototype.Get_Api = function () {
	return this.Api;
};

CPresentation.prototype.GetApi = function () {
	return this.Api;
};

CPresentation.prototype.Get_CollaborativeEditing = function () {
	return this.CollaborativeEditing;
};

CPresentation.prototype.addSlideMaster = function (pos, master) {
	History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_Presentation_AddSlideMaster, pos, [master], true));
	this.slideMasters.splice(pos, 0, master);
};
CPresentation.prototype.addNotesMaster = function (pos, master) {
	//History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_Presentation_AddSlideMaster, pos, [master], true));
	this.notesMasters.splice(pos, 0, master);
};

CPresentation.prototype.removeSlideMaster = function (pos, count) {
	History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_Presentation_RemoveSlideMaster, pos, this.slideMasters.slice(pos, pos + count), false));
	this.slideMasters.splice(pos, count);
};
CPresentation.prototype.removeSlideMasterObject = function (oMaster) {
	for(let nMaster = 0; nMaster < this.slideMasters.length; ++nMaster) {
		if(this.slideMasters[nMaster] === oMaster) {
			this.removeSlideMaster(nMaster, 1);
			return;
		}
	}
};

CPresentation.prototype.Get_Id = function () {
	return this.Id;
};


CPresentation.prototype.AddNewMasterSlide = function () {
	if(!this.IsMasterMode) return;
	this.StartAction(0);
	let oMaster = AscCommonSlide.CreateDefaultMaster();
	oMaster.changeSize(this.GetWidthMM(), this.GetHeightMM());
	for(let nLt = 0 ; nLt < oMaster.sldLayoutLst.length; ++nLt) {
		oMaster.sldLayoutLst[nLt].changeSize(this.GetWidthMM(), this.GetHeightMM());
	}
	this.addSlideMaster(this.slideMasters.length, oMaster);
	this.FinalizeAction(false);
	this.DrawingDocument.m_oWordControl.GoToPage(this.GetSlideIndex(oMaster));

};
CPresentation.prototype.AddNewLayout = function () {
	if(!this.IsMasterMode) return;
	this.StartAction(0);
	let oMaster = this.GetCurrentMaster();
	if(oMaster) {
		let oLayout = oMaster.addNewLayout();
		this.Recalculate();
		this.DrawingDocument.m_oWordControl.GoToPage(this.GetSlideIndex(oLayout));
	}
	this.FinalizeAction(true);
};

CPresentation.prototype.LoadEmptyDocument = function () {
	this.DrawingDocument.TargetStart();
	this.Recalculate();

	this.Interface_Update_ParaPr();
	this.Interface_Update_TextPr();
};

CPresentation.prototype.EndPreview_MailMergeResult = function () {
};

CPresentation.prototype.CheckNeedUpdateTargetForCollaboration = function () {
	if (!this.NeedUpdateTargetForCollaboration) {
		var oController = this.GetCurrentController();
		if (oController) {
			var oTargetDocContent = oController.getTargetDocContent();
			if (oTargetDocContent !== this.oLastCheckContent) {
				this.oLastCheckContent = oTargetDocContent;
				return true;
			}
		}
		return false;
	}
	return true;
};


CPresentation.prototype.Is_OnRecalculate = function () {
	return true;
};
CPresentation.prototype.Continue_FastCollaborativeEditing = function () {
	if (true === AscCommon.CollaborativeEditing.Get_GlobalLock()) {
		if (this.Api.forceSaveUndoRequest)
			this.Api.asc_Save(true);

		return;
	}

	if (this.Api.isLongAction())
		return;
	if (true !== AscCommon.CollaborativeEditing.Is_Fast() || true === AscCommon.CollaborativeEditing.Is_SingleUser())
		return;

	var oController = this.GetCurrentController();
	if (oController) {
		if (oController.isTrackingDrawings() || this.Api.isOpenedFrameEditor) {
			return;
		}
	}

	var bHaveChanges = History.Have_Changes(true);
	if (true !== bHaveChanges && (true === AscCommon.CollaborativeEditing.Have_OtherChanges() || 0 !== AscCommon.CollaborativeEditing.getOwnLocksLength())) {
		// Принимаем чужие изменение. Своих нет, но функцию отсылки надо вызвать, чтобы снялить локи.
		AscCommon.CollaborativeEditing.Apply_Changes();
		AscCommon.CollaborativeEditing.Send_Changes();
	} else if (true === bHaveChanges || true === AscCommon.CollaborativeEditing.Have_OtherChanges()) {
		this.Api.asc_Save(true);
	}

	var CurTime = new Date().getTime();
	if (this.CheckNeedUpdateTargetForCollaboration() && (CurTime - this.LastUpdateTargetTime > 1000)) {
		this.NeedUpdateTargetForCollaboration = false;
		if (true !== bHaveChanges) {
			var CursorInfo = History.Get_DocumentPositionBinary();
			if (null !== CursorInfo) {
				this.Api.CoAuthoringApi.sendCursor(CursorInfo);
				this.LastUpdateTargetTime = CurTime;
			}
		} else {
			this.LastUpdateTargetTime = CurTime;
		}
	}
};

CPresentation.prototype.Get_DocumentPositionInfoForCollaborative = function () {
	var oController = this.GetCurrentController();
	var oRes;
	if (oController) {
		oRes = oController.getDocumentPositionForCollaborative();
		if (oRes) {
			return oRes;
		}
	}
	return {Class: this, Position: 0};
};

CPresentation.prototype.GetRecalculateMaps = function () {
	var ret = {
		layouts: {},
		masters: {}
	};

	for (var i = 0; i < this.Slides.length; ++i) {
		if (this.Slides[i].Layout) {
			ret.layouts[this.Slides[i].Layout.Id] = this.Slides[i].Layout;
			if (this.Slides[i].Layout.Master) {
				ret.masters[this.Slides[i].Layout.Master.Id] = this.Slides[i].Layout.Master;
			}
		}
	}
	return ret;
};

CPresentation.prototype.replaceMisspelledWord = function (Word, SpellCheckProperty) {
	var ParaId = SpellCheckProperty.ParaId;
	var Paragraph = g_oTableId.Get_ById(ParaId);
	Paragraph.Document_SetThisElementCurrent(true);
	var oController = this.GetCurrentController();
	if (oController) {
		oController.checkSelectedObjectsAndCallback(function () {
			Paragraph.ReplaceMisspelledWord(Word, SpellCheckProperty.Element);
		}, [], false, AscDFH.historydescription_Document_ReplaceMisspelledWord);
	}
};

CPresentation.prototype.CancelEyedropper = function() {
	if(this.Api.isEyedropperStarted()) {
		this.Api.cancelEyedropper();
		return true;
	}
	return false;
};
CPresentation.prototype.CancelInkDrawer = function() {
	if(this.Api.isInkDrawerOn()) {
		this.Api.stopInkDrawer();
		return true;
	}
	return false;
};

CPresentation.prototype.IsMasterMode = function () {
	return Asc.editor.isMasterMode();
};
CPresentation.prototype.GetSlideObjectsWithTheme = function (oTheme) {
	let result = {
		masters: [],
		layouts: [],
		slides: []
	};
	for(let nIdx = 0; nIdx < this.slideMasters.length; ++nIdx) {
		let oMaster = this.slideMasters[nIdx];
		if(oMaster.getTheme() === oTheme) {
			result.masters.push(oMaster);
			result.layouts = result.layouts.concat(oMaster.sldLayoutLst);
		}
	}
	for(let nIdx = 0; nIdx < this.Slides.length; ++nIdx) {
		let oSlide = this.Slides[nIdx];
		if(oSlide.getTheme() === oTheme) {
			result.slides.push(oSlide);
		}
	}

	return result;
};

CPresentation.prototype.Recalculate = function (RecalcData) {
	this.DrawingDocument.OnStartRecalculate(this.GetSlidesCount());
	this.StopAnimationPreview();
	++this.RecalcId;
	this.private_ClearSearchOnRecalculate();

	if (undefined === RecalcData && this.private_RecalculateFastRunRange(History.GetNonRecalculatedChanges()))
		return;

	if (this.SearchEngine.ClearOnRecalc) {
		this.SearchEngine.Clear();
		this.SearchEngine.ClearOnRecalc = false;
	}
	let _RecalcData = RecalcData ? RecalcData : History.Get_RecalcData(), key, bSync = true, i,
		bRedrawAllSlides = false, aToRedrawSlides = [], redrawSlideIndexMap = {}, slideIndex, isUpdateThemes = false;
	let bAttack = undefined;
	this.updateSlideIndexes();
	let b_check_layout = false;
	let bRedrawNotes = false;
	let oCurSlide = this.GetCurrentSlide();
	let oCurMaster = this.GetCurrentMaster();
	let aAllSlides = this.GetAllSlides();
	let oSlide;
	if (_RecalcData.Drawings.All) {
		b_check_layout = true;
		for (key in this.slideMasters) {
			if (this.slideMasters.hasOwnProperty(key)) {
				let oMaster = this.slideMasters[key];
				if (oCurMaster === oMaster) {
					bAttack = oMaster.needRecalc();
				}
				isUpdateThemes = isUpdateThemes || oMaster.needRecalc();
				if (oMaster.needRecalc()) {
					oMaster.recalculate();
					for(let nIdx = 0; nIdx < this.Slides.length; ++nIdx) {
						oSlide = this.Slides[nIdx];
						if (oSlide.Layout.Master === oMaster) {
							oSlide.checkSlideTheme();
						}
					}
				} 
				for (key in oMaster.sldLayoutLst) {
					if (oMaster.sldLayoutLst.hasOwnProperty(key)) {
						let oSlideLayout = oMaster.sldLayoutLst[key];
						if (oSlideLayout.needRecalc()) {
							if (oSlideLayout.type === AscFormat.nSldLtTTitle) {
								isUpdateThemes = true;
							}
							if (oCurMaster === oMaster) {
								bAttack = true;
							}
							for(let nIdx = 0; nIdx < this.Slides.length; ++nIdx) {
								oSlide = this.Slides[nIdx];
								if (oSlide.Layout == oSlideLayout) {
									oSlide.checkSlideTheme();
								}
							}
							oSlideLayout.ImageBase64 = "";
							oSlideLayout.recalculate();
						}
					}
				}
			}
		}
		this.bNeedUpdateChartPreview = true;
		bRedrawAllSlides = true;
		for (let nIdx = 0; nIdx < aAllSlides.length; ++nIdx) {
			let oCalcSlide = aAllSlides[nIdx];
			if (oCalcSlide.bChangeLayout) {
				oCalcSlide.checkSlideTheme();
			}
			oCalcSlide.recalcText();
			oCalcSlide.recalculate();
			oCalcSlide.recalculateNotesShape();
		}
	} else {
		let oCurNotesShape = null;
		if (oCurSlide) {
			oCurNotesShape = oCurSlide.notesShape;
		}
		let oPlaceholders = {};
		let oMasterPlaceholders = {};
		let bPlaceholders = false;
		let bMasterPlaceholders = false;
		let oMasters = {};
		for (key in _RecalcData.Drawings.Map) {
			if (_RecalcData.Drawings.Map.hasOwnProperty(key)) {
				var oDrawingObject = _RecalcData.Drawings.Map[key];
				if (AscCommon.g_oTableId.Get_ById(key) === oDrawingObject) {
					if(oDrawingObject.IsUseInDocument && !oDrawingObject.IsUseInDocument()) {
						continue;
					}
					let parent = oDrawingObject.parent;
					if(oDrawingObject instanceof AscFormat.CTheme) {
						b_check_layout = true;
						bAttack = true;
						let oThemeObjects = this.GetSlideObjectsWithTheme(oDrawingObject);
						if(!this.IsMasterMode()) {
							for(let nIdx = 0; nIdx < oThemeObjects.masters.length; ++nIdx) {
								oThemeObjects.masters[nIdx].recalculate();
							}
							for(let nIdx = 0; nIdx < oThemeObjects.layouts.length; ++nIdx) {
								oThemeObjects.layouts[nIdx].recalculate();
							}
							let aIdx = [];
							let nStartIdx = 0;
							for(let nIdx = 0; nIdx < oThemeObjects.slides.length; ++nIdx) {
								if(oThemeObjects.slides[nIdx].num === this.CurPage) {
									nStartIdx = aIdx.length;
								}
								aIdx.push(oThemeObjects.slides[nIdx].num);
							}
							AscFormat.redrawSlide(this.Slides[aIdx[nStartIdx]], this, aIdx, nStartIdx, 0, this.Slides);
						}
						else {
							for(let nIdx = 0; nIdx < oThemeObjects.slides.length; ++nIdx) {
								oThemeObjects.slides[nIdx].recalculate();
							}
							let aIdx = [];
							let nStartIdx = 0;
							for(let nIdx = 0; nIdx < oThemeObjects.masters.length; ++nIdx) {
								let nMasterIdx = this.GetSlideIndex(oThemeObjects.masters[nIdx]);
								if(nMasterIdx === this.CurPage) {
									nStartIdx = aIdx.length;
								}
								aIdx.push(nMasterIdx);
							}
							for(let nIdx = 0; nIdx < oThemeObjects.layouts.length; ++nIdx) {
								let nLayoutIdx = this.GetSlideIndex(oThemeObjects.layouts[nIdx]);
								if(nLayoutIdx === this.CurPage) {
									nStartIdx = aIdx.length;
								}
								aIdx.push(nLayoutIdx);
							}
							AscFormat.redrawSlide(this.GetSlide(aIdx[nStartIdx]), this, aIdx, nStartIdx, 0, this.GetAllSlides());
						}
					}
					else {
						oDrawingObject.recalculate();
						if (oDrawingObject instanceof AscCommonSlide.MasterSlide) {
							b_check_layout = true;
							bAttack = true;
							isUpdateThemes = true;
							bRedrawAllSlides = true;
							if(oDrawingObject.needRecalc()) {
								for (let nIdx = 0; nIdx < this.Slides.length; ++nIdx) {
									let oCalcSlide = this.Slides[nIdx];
									if (oCalcSlide.Layout.Master === oDrawingObject) {
										oCalcSlide.checkSlideTheme();
										oCalcSlide.recalculate();
										let nIdx = this.GetSlideIndex(oCalcSlide);
										if(nIdx !== -1) {
											if (redrawSlideIndexMap[nIdx] !== true) {
												redrawSlideIndexMap[nIdx] = true;
												aToRedrawSlides.push(nIdx);
											}
										}
									}
								}
								for(let nLt = 0; nLt < oDrawingObject.sldLayoutLst.length; ++nLt) {
									let oLt = oDrawingObject.sldLayoutLst[nLt];
									oLt.checkSlideTheme();
									oLt.recalculate();
									let nIdx = this.GetSlideIndex(oLt);
									if(nIdx !== -1) {
										if (redrawSlideIndexMap[nIdx] !== true) {
											redrawSlideIndexMap[nIdx] = true;
											aToRedrawSlides.push(nIdx);
										}
									}
								}
							}
							let nIdx = this.GetSlideIndex(oDrawingObject);
							if(nIdx !== -1) {
								if (redrawSlideIndexMap[nIdx] !== true) {
									redrawSlideIndexMap[nIdx] = true;
									aToRedrawSlides.push(nIdx);
								}
							}
						}
						let oLayoutObject;
						if (parent instanceof AscCommonSlide.SlideLayout) {
							oLayoutObject = parent;
						} else if (oDrawingObject instanceof AscCommonSlide.SlideLayout) {
							oLayoutObject = oDrawingObject;
						}
						if (oLayoutObject) {
							if (oLayoutObject.type === AscFormat.nSldLtTTitle) {
								isUpdateThemes = true;
							}
							oLayoutObject.ImageBase64 = "";
							b_check_layout = true;
							bAttack = true;
							for (let nIdx = 0; nIdx < this.Slides.length; ++nIdx) {
								let oCalcSlide = this.Slides[nIdx];
								if (oCalcSlide.Layout === oLayoutObject) {
									oCalcSlide.checkSlideTheme();
									oCalcSlide.recalculate();
									if (!this.IsMasterMode() && redrawSlideIndexMap[nIdx] !== true) {
										redrawSlideIndexMap[nIdx] = true;
										aToRedrawSlides.push(nIdx);
									}
								}
							}
							let nIdx = this.GetSlideIndex(oLayoutObject);
							if(nIdx !== -1) {
								if (redrawSlideIndexMap[nIdx] !== true) {
									redrawSlideIndexMap[nIdx] = true;
									aToRedrawSlides.push(nIdx);
								}
							}
						}
						if (oDrawingObject.getSlideIndex) {
							slideIndex = oDrawingObject.getSlideIndex();
							if (slideIndex !== null && slideIndex !== -1) {
								if (redrawSlideIndexMap[slideIndex] !== true) {
									redrawSlideIndexMap[slideIndex] = true;
									aToRedrawSlides.push(slideIndex);
								}
							} else {
								if (oCurNotesShape && oCurNotesShape === oDrawingObject) {
									if (oCurSlide) {
										oCurSlide.recalculateNotesShape();
									}
									bRedrawNotes = true;
								}
							}
						}
					}
					if(oDrawingObject.isDrawing ) {
						if(oDrawingObject.isPlaceholder()) {
							if(parent instanceof AscCommonSlide.SlideLayout || parent instanceof AscCommonSlide.MasterSlide) {
								oPlaceholders[oDrawingObject.Id] = oDrawingObject;
								bPlaceholders = true;
								if(parent instanceof AscCommonSlide.MasterSlide) {
									oMasterPlaceholders[oDrawingObject.Id] = oDrawingObject;
									bMasterPlaceholders = true;
									oMasters[parent.Id] = parent;
								}
							}
						}
						else {
							if(parent instanceof AscCommonSlide.MasterSlide) {
								for(let nSld = 0; nSld < aAllSlides.length; ++nSld) {
									if (redrawSlideIndexMap[nSld] !== true) {
										let oSld = aAllSlides[nSld];
										let oMS = oSld.getMaster();
										if(oMS === parent) {

											redrawSlideIndexMap[nSld] = true;
											aToRedrawSlides.push(nSld);
										}
									}
								}
							}
						}
					}
				}
			}
		}
		if(bPlaceholders) {
			for(let nSld = 0; nSld < this.Slides.length; ++nSld) {
				let bCalc = this.Slides[nSld].checkPlaceholders(oPlaceholders);
				if(bCalc) {
					if(!this.IsMasterMode()) {
						if (redrawSlideIndexMap[nSld] !== true) {
							redrawSlideIndexMap[nSld] = true;
							aToRedrawSlides.push(nSld);
						}
					}
				}
			}
		}
		if(bMasterPlaceholders) {
			for(let sKey in oMasters) {
				if(oMasters.hasOwnProperty(sKey)) {
					let oMaster = oMasters[sKey];
					let aLayouts = oMaster.sldLayoutLst;
					for(let nLt = 0; nLt < aLayouts.length; ++nLt) {
						let bCalc = aLayouts[nLt].checkPlaceholders(oMasterPlaceholders);
						if(bCalc) {
							if(this.IsMasterMode()) {
								let nIndex = this.GetSlideIndex(aLayouts[nLt]);
								if (redrawSlideIndexMap[nIndex] !== true) {
									redrawSlideIndexMap[nIndex] = true;
									aToRedrawSlides.push(nIndex);
								}
							}
						}
					}
				}
			}
		}
	}
	History.Reset_RecalcIndex();
	this.RecalculateCurPos();
	oCurSlide = this.GetCurrentSlide();
	if (bSync) {
		let bEndRecalc;
		if (bRedrawAllSlides) {
			this.bNeedUpdateTh = true;
			bEndRecalc = (this.GetSlidesCount() > 0);
			if (oCurSlide) {
				this.DrawingDocument.OnRecalculateSlide(this.CurPage);
				this.DrawingDocument.Notes_OnRecalculate(this.CurPage, oCurSlide.NotesWidth, oCurSlide.getNotesHeight());
			}

		} else {
			aToRedrawSlides.sort(AscCommon.fSortAscending);
			let nSlideMinIdx = _RecalcData.Drawings.SlideMinIdx;
			if (AscFormat.isRealNumber(nSlideMinIdx)) {
				for (i = 0; i < aToRedrawSlides.length && aToRedrawSlides[i] < nSlideMinIdx; ++i) {
					this.DrawingDocument.OnRecalculateSlide(aToRedrawSlides[i]);
				}
				for (i = nSlideMinIdx; i < this.Slides.length; ++i) {
					this.DrawingDocument.OnRecalculateSlide(i);
				}
			} else {
				for (i = 0; i < aToRedrawSlides.length; ++i) {
					this.DrawingDocument.OnRecalculateSlide(aToRedrawSlides[i]);
				}
			}
			bEndRecalc = (aToRedrawSlides.length > 0) || AscFormat.isRealNumber(nSlideMinIdx);
		}
		if (bRedrawNotes) {
			if (oCurSlide) {
				this.DrawingDocument.Notes_OnRecalculate(this.CurPage, oCurSlide.NotesWidth, oCurSlide.getNotesHeight());
			}
		}
		let nSlidesCount = this.GetSlidesCount();
		if (bEndRecalc || nSlidesCount === 0 || nSlidesCount !== this.GetThumbnailsCount()) {
			this.DrawingDocument.OnEndRecalculate();
		}
	}

	if (isUpdateThemes || this.bNeedUpdateThemes) {
		this.SendThemesThumbnails();
		this.bNeedUpdateThemes = false;
		let oMaster = this.GetCurrentMaster();
		if(oMaster)
		{
			this.Api.sendEvent("asc_onUpdateThemeIndex", oMaster.getThemeIndex());
			this.Api.sendColorThemes(oMaster.Theme);
		}
	}
	if (!oCurSlide) {
		this.DrawingDocument.m_oWordControl.GoToPage(-1);
		if (b_check_layout) {
			this.DrawingDocument.m_oWordControl.CheckLayouts(bAttack);
		}
	} else {
		if (this.bGoToPage) {
			this.DrawingDocument.m_oWordControl.GoToPage(this.CurPage);
			this.bGoToPage = false;
		} else if (b_check_layout) {
			this.DrawingDocument.m_oWordControl.CheckLayouts(bAttack);
		}
		if (this.needSelectPages.length > 0) {
			this.needSelectPages.length = 0;
		}
	}
	if (this.bNeedUpdateTh) {
		this.DrawingDocument.UpdateThumbnailsAttack();
		this.bNeedUpdateTh = false;
	}
	this.Document_UpdateSelectionState();

	if (oCurSlide) {
		if (this.DrawingDocument.placeholders)
			this.DrawingDocument.placeholders.update(oCurSlide.getPlaceholdersControls());
	}
	this.MathTrackHandler.Update();
};

CPresentation.prototype.private_RecalculateFastRunRange = function (arrChanges, nStartIndex, nEndIndex) {

	var _nStartIndex = undefined !== nStartIndex ? nStartIndex : 0;
	var _nEndIndex = undefined !== nEndIndex ? nEndIndex : arrChanges.length - 1;

	var oRun = null;
	for (var nIndex = _nStartIndex; nIndex <= _nEndIndex; ++nIndex) {
		var oChange = arrChanges[nIndex];

		if (oChange.IsDescriptionChange())
			continue;

		if (!oRun)
			oRun = oChange.GetClass();
		else if (oRun !== oChange.GetClass())
			return false;
	}

	if (!oRun || !(oRun instanceof ParaRun) || !oRun.GetParagraph())
		return false;

	var oParaPos = oRun.GetSimpleChangesRange(arrChanges, _nStartIndex, _nEndIndex);
	if (oParaPos) {

		var oParagraph = oRun.GetParagraph();
		var nRes = oParagraph.RecalculateFastRunRange(oParaPos);
		if (-1 !== nRes) {
			var oCurSlide = this.GetCurrentSlide();
			if (oCurSlide) {
				if (!this.IsFocusOnNotes()) {
					this.DrawingDocument.OnRecalculateSlide(this.CurPage);
					this.DrawingDocument.OnEndRecalculate();
				} else {
					this.DrawingDocument.Notes_OnRecalculate(this.CurPage, oCurSlide.NotesWidth, oCurSlide.getNotesHeight());
				}

			}
			History.Get_RecalcData();
			History.Reset_RecalcIndex();
			var DrawingShape = oParagraph.Parent.Is_DrawingShape(true);
			if (DrawingShape && DrawingShape.recalcInfo && DrawingShape.recalcInfo.recalcTitle) {
				DrawingShape.recalcInfo.bRecalculatedTitle = true;
				DrawingShape.recalcInfo.recalcTitle = null;
			}
			return true;
		}
	}

	return false;
};

CPresentation.prototype.updateSlideIndexes = function () {
	for (var i = 0; i < this.Slides.length; ++i) {
		this.Slides[i].changeNum(i);
	}
};

CPresentation.prototype.GenerateThumbnails = function (_drawerThemes, _drawerLayouts) {
	var _masters = this.slideMasters;
	var _len = _masters.length;
	var aLayouts, i, j;
	for (i = 0; i < _len; i++) {
		_masters[i].ImageBase64 = _drawerThemes.GetThumbnail(_masters[i]);
		aLayouts = _masters[i].sldLayoutLst;
		for (j = 0; j < aLayouts.length; ++j) {
			aLayouts[j].ImageBase64 = _drawerLayouts.GetThumbnail(aLayouts[j]);
			aLayouts[j].Width64 = _drawerLayouts.WidthPx;
			aLayouts[j].Height64 = _drawerLayouts.HeightPx;
		}
	}
};


/**
 * Получаем идентификатор текущего пересчета
 * @returns {number}
 */
CPresentation.prototype.GetRecalcId = function () {
	return this.RecalcId;
};
CPresentation.prototype.StopRecalculate = function () {
	this.clearThemeTimeouts();
//        this.DrawingDocument.OnStartRecalculate( 0 );
};
CPresentation.prototype.PauseRecalculate = function () {
	this.StopRecalculate();
};
CPresentation.prototype.ResumeRecalculate = function () {
};

CPresentation.prototype.OnContentReDraw = function (StartPage, EndPage) {
	this.ReDraw(StartPage, EndPage);
};

CPresentation.prototype.checkGridCache = function (oGraphics) {
	let oCoordTr = oGraphics.m_oCoordTransform;
	let oContext = oGraphics.m_oContext;
	if (!oContext ||
		!oCoordTr ||
		AscCommon.IsShapeToImageConverter ||
		oGraphics.IsThumbnail ||
		oGraphics.animationDrawer ||
		oGraphics.IsDemonstrationMode ||
		oGraphics.isBoundsChecker()) {
		return;
	}
	let nWidth = (oCoordTr.sx * this.GetWidthMM() + 0.5) >> 0;
	let nHeight = (oCoordTr.sy * this.GetHeightMM() + 0.5) >> 0;
	if (nWidth === 0 || nHeight === 0) {
		return;
	}
	let bUpdateCache = false;
	if (!this.cachedGridCanvas) {
		bUpdateCache = true;
	}
	if (!bUpdateCache) {
		if (this.cachedGridCanvas.width !== nWidth || this.cachedGridCanvas.height !== nHeight) {
			bUpdateCache = true
		}
	}
	if (!bUpdateCache) {
		let nGridSpacing = this.getGridSpacing();
		if (this.cachedGridSpacing !== nGridSpacing) {
			bUpdateCache = true;
		}
	}
	if (bUpdateCache) {
		if (!this.cachedGridCanvas) {
			this.cachedGridCanvas = document.createElement('canvas');
		}
		this.cachedGridCanvas.width = nWidth;
		this.cachedGridCanvas.height = nHeight;
		this.cachedGridSpacing = this.getGridSpacing();
		let oCtx = this.cachedGridCanvas.getContext('2d');
		let oCacheGraphics = new AscCommon.CGraphics();
		oCacheGraphics.init(oCtx, nWidth, nHeight, nWidth / oCoordTr.sx, nHeight / oCoordTr.sy);
		oCacheGraphics.m_oFontManager = AscCommon.g_fontManager;
		oCacheGraphics.transform(1, 0, 0, 1, 0, 0);
		this.drawGrid(oCacheGraphics);
	}
	let nX = oCoordTr.TransformPointX(0, 0) + 0.5 >> 0;
	let nY = oCoordTr.TransformPointY(0, 0) + 0.5 >> 0;
	oGraphics.SaveGrState();
	oGraphics.SetIntegerGrid(true);

	let oGrCtx = oGraphics.m_oContext;
	let sOldCompostiteOperation = oContext.globalCompositeOperation;
	oContext.globalCompositeOperation = "difference";
	oGrCtx.drawImage(this.cachedGridCanvas, nX, nY);
	oGrCtx.globalCompositeOperation = sOldCompostiteOperation;
	oGraphics.RestoreGrState();
};
CPresentation.prototype.CheckTargetUpdate = function () {
	if (this.DrawingDocument.UpdateTargetFromPaint === true) {
		if (true === this.DrawingDocument.UpdateTargetCheck)
			this.NeedUpdateTarget = this.DrawingDocument.UpdateTargetCheck;
		this.DrawingDocument.UpdateTargetCheck = false;
	}

	if (true === this.NeedUpdateTarget) {
		this.RecalculateCurPos();
		this.NeedUpdateTarget = false;
	}
};

CPresentation.prototype.RecalculateCurPos = function (bUpdateX, bUpdateY) {
	var oController = this.GetCurrentController();
	if (oController) {
		oController.recalculateCurPos(bUpdateX, bUpdateY);
	}
};

CPresentation.prototype.Set_TargetPos = function (X, Y, PageNum) {
	this.TargetPos.X = X;
	this.TargetPos.Y = Y;
	this.TargetPos.PageNum = PageNum;
};

// Вызываем перерисовку нужных страниц
CPresentation.prototype.ReDraw = function (StartPage, EndPage) {
	this.DrawingDocument.OnRecalculateSlide(StartPage);
};
CPresentation.prototype.DrawPage = function (nPageIndex, pGraphics) {
	this.Draw(nPageIndex, pGraphics);
};
CPresentation.prototype.RedrawCurSlide = function () {
	let oSlide = this.GetCurrentSlide();
	if (oSlide) {
		let oDrawingDocument = this.DrawingDocument;
		oDrawingDocument.OnRecalculateSlide(this.CurPage);
		oDrawingDocument.OnEndRecalculate();
	}
};
CPresentation.prototype.drawGrid = function (oGraphics) {
	this.getStrideData().drawGrid(oGraphics);
};
CPresentation.prototype.drawGuides = function (oGraphics) {
	if (this.viewPr) {
		this.viewPr.drawGuides(oGraphics);
	}
};
CPresentation.prototype.getHorGuidesPos = function () {
	if (this.viewPr) {
		return this.viewPr.getHorGuidesPos();
	}
	return [];

};
CPresentation.prototype.getVertGuidesPos = function () {
	if (this.viewPr) {
		return this.viewPr.getVertGuidesPos();
	}
	return [];
};
CPresentation.prototype.canClearGuides = function () {
	if (this.viewPr) {
		return this.viewPr.canClearGuides();
	}
	return false;
};
CPresentation.prototype.getGuidesCount = function () {
	if (this.viewPr) {
		return this.viewPr.getHorGuidesPos().length + this.viewPr.getVertGuidesPos().length;
	}
	return 0;
};
CPresentation.prototype.clearGuides = function () {
	if (!this.canClearGuides()) {
		return;
	}
	if (this.viewPr) {
		if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_ViewPr, undefined, undefined, [])) {
			this.Create_NewHistoryPoint(0);
			this.viewPr.clearGuides();
			this.Recalculate();
			this.UpdateInterface();
		}
	}
};
CPresentation.prototype.deleteGuide = function (sId) {
	if (this.viewPr) {
		if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_ViewPr, undefined, undefined, [])) {
			this.Create_NewHistoryPoint(0);
			this.viewPr.removeGuideById(sId);
			this.Recalculate();
			this.UpdateInterface();
		}
	}
};
CPresentation.prototype.hitInGuide = function (x, y) {
	if(!this.Api.asc_getShowGuides()) {
		return null;
	}
	if (this.viewPr) {
		return this.viewPr.hitInGuide(x, y);
	}
	return null;
};
CPresentation.prototype.scaleGuides = function (dCW, dCH) {
	if (this.viewPr) {
		this.viewPr.scaleGuides(dCW, dCH);
	}
};
CPresentation.prototype.Update_ForeignCursor = function (CursorInfo, UserId, Show, UserShortId) {
	if (!this.Api.User)
		return;

	if (UserId === this.Api.CoAuthoringApi.getUserConnectionId())
		return;

	AscFormat.drawingsUpdateForeignCursor(this.GetCurrentController(), this.DrawingDocument, CursorInfo, UserId, Show, UserShortId);
};

CPresentation.prototype.Remove_ForeignCursor = function (UserId) {
	this.DrawingDocument.Collaborative_RemoveTarget(UserId);
	AscCommon.CollaborativeEditing.Remove_ForeignCursor(UserId);
};

/**
 * Список позиций, которые мы собираемся отслеживать
 * @param arrPositions
 */
CPresentation.prototype.TrackDocumentPositions = function (arrPositions) {
	this.CollaborativeEditing.Clear_DocumentPositions();

	for (var nIndex = 0, nCount = arrPositions.length; nIndex < nCount; ++nIndex) {
		this.CollaborativeEditing.Add_DocumentPosition(arrPositions[nIndex]);
	}
};
/**
 * Обновляем отслеживаемые позиции
 * @param arrPositions
 */
CPresentation.prototype.RefreshDocumentPositions = function (arrPositions) {
	for (var nIndex = 0, nCount = arrPositions.length; nIndex < nCount; ++nIndex) {
		this.CollaborativeEditing.Update_DocumentPosition(arrPositions[nIndex]);
	}
};

CPresentation.prototype.GetTargetPosition = function () {
	var oController = this.GetCurrentController();
	var oPosition = null;
	if (oController) {
		var oTargetDocContent = oController.getTargetDocContent(false, false);
		if (oTargetDocContent) {
			var oElem = oTargetDocContent.Content[oTargetDocContent.CurPos.ContentPos];
			if (oElem) {
				var oPos = oElem.GetTargetPos();
				if (oPos) {

				}
				var x, y;
				if (oPos.Transform) {
					x = oPos.Transform.TransformPointX(oPos.X, oPos.Y + oPos.Height);
					y = oPos.Transform.TransformPointY(oPos.X, oPos.Y + oPos.Height);
				} else {
					x = oPos.X;
					y = oPos.Y + oPos.Height;
				}
				oPosition = {X: x, Y: y};
			}
		}
	}
	return oPosition;
};

// Отрисовка содержимого Документа
CPresentation.prototype.Draw = function (nPageIndex, pGraphics) {
	if (!pGraphics.isBoundsChecker()) {
		AscCommon.CollaborativeEditing.Update_ForeignCursorsPositions();
	}
	let oSlide = this.GetSlide(nPageIndex);
	if(oSlide) {
		oSlide.draw(pGraphics);
	}
};

CPresentation.prototype.AddNewParagraph = function (bRecalculate) {
	var oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.addNewParagraph, [], false, AscDFH.historydescription_Presentation_AddNewParagraph);
	this.Document_UpdateInterfaceState();
};
CPresentation.prototype.Search = function (oProps) {
	if (true === this.SearchEngine.Compare(oProps))
		return this.SearchEngine;
	this.SearchEngine.Clear();
	this.SearchEngine.Set(oProps);

	let nCount = this.GetSlidesCount();
	for (var i = 0; i < nCount; ++i) {
		this.GetSlide(i).Search(this.SearchEngine, search_Common);
	}

	this.DrawingDocument.ClearCachePages();
	this.DrawingDocument.FirePaint();
	this.SearchEngine.ClearOnRecalc = true;
	return this.SearchEngine;
};
CPresentation.prototype.ClearSearch = function () {
	let isPrevSearch = this.SearchEngine.Count > 0;
	this.SearchEngine.Clear();
	if (isPrevSearch) {
		this.Api.sync_SearchEndCallback();
	}
};
CPresentation.prototype.private_ClearSearchOnRecalculate = function () {
	if (!this.SearchEngine.ClearOnRecalc) {
		return;
	}

	this.ClearSearch();
};


CPresentation.prototype.GetSearchElementId = function (isNext) {
	if (this.Slides.length > 0) {
		var i, Id, content, start_index;
		var target_text_object;
		var oCommonController = this.GetCurrentController();
		if (oCommonController) {
			target_text_object = AscFormat.getTargetTextObject(oCommonController);
		}
		if (target_text_object) {
			if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
				Id = target_text_object.graphicObject.GetSearchElementId(isNext, true);
				if (Id !== null) {
					return Id;
				}
			} else {
				content = target_text_object.getDocContent();
				if (content) {
					Id = content.GetSearchElementId(isNext, true);
					if (Id !== null) {
						return Id;
					}
				}
			}
		}
		var sp_tree = this.GetCurrentSlide().cSld.spTree, group_shapes, group_start_index;
		let oController = this.GetCurrentController();
		var bSkipCurNotes = false;
		if (isNext) {
			if (oController.selection.groupSelection) {
				group_shapes = oController.selection.groupSelection.arrGraphicObjects;
				for (i = 0; i < group_shapes.length; ++i) {
					if (group_shapes[i].selected && group_shapes[i].getObjectType() === AscDFH.historyitem_type_Shape) {
						content = group_shapes[i].getDocContent();
						if (content) {
							Id = content.GetSearchElementId(isNext, isRealObject(target_text_object));
							if (Id !== null) {
								return Id;
							}
						}
						group_start_index = i + 1;
					}
				}
				for (i = group_start_index; i < group_shapes.length; ++i) {
					if (group_shapes[i].getObjectType() === AscDFH.historyitem_type_Shape) {
						content = group_shapes[i].getDocContent();
						if (content) {
							Id = content.GetSearchElementId(isNext, false);
							if (Id !== null) {
								return Id;
							}
						}
					}
				}

				for (i = 0; i < sp_tree.length; ++i) {
					if (sp_tree[i] === oController.selection.groupSelection) {
						start_index = i + 1;
						break;
					}
				}
				if (i === sp_tree.length) {
					start_index = sp_tree.length;
				}
			} else if (oController.selectedObjects.length === 0) {
				start_index = 0;
				if (this.IsFocusOnNotes()) {
					start_index = sp_tree.length;
					bSkipCurNotes = true;
				}
			} else {
				for (i = 0; i < sp_tree.length; ++i) {
					if (sp_tree[i].selected) {
						start_index = target_text_object ? i + 1 : i;
						break;
					}
				}
				if (i === sp_tree.length) {
					start_index = sp_tree.length;
				}
			}
			Id = this.GetCurrentSlide().GetSearchElementId(isNext, start_index);
			if (Id !== null) {
				return Id;
			}
			var oCurSlide = this.GetCurrentSlide();
			if (oCurSlide.notesShape && !bSkipCurNotes) {
				Id = oCurSlide.notesShape.GetSearchElementId(isNext, false);
				if (Id !== null) {
					return Id;
				}
			}
			let nCount = this.GetSlidesCount();
			for (i = this.CurPage + 1; i < nCount; ++i) {
				Id = this.GetSlide(i).GetSearchElementId(isNext, 0);
				if (Id !== null) {
					return Id;
				}
				if (this.GetSlide(i).notesShape) {
					Id = this.GetSlide(i).notesShape.GetSearchElementId(isNext, false);
					if (Id !== null) {
						return Id;
					}
				}
			}
			for (i = 0; i <= this.CurPage; ++i) {
				Id = this.GetSlide(i).GetSearchElementId(isNext, 0);
				if (Id !== null) {
					return Id;
				}
				if (this.GetSlide(i).notesShape) {
					Id = this.GetSlide(i).notesShape.GetSearchElementId(isNext, false);
					if (Id !== null) {
						return Id;
					}
				}
			}

		} else {
			let oController = this.GetCurrentSlideController();
			if (oController.selection.groupSelection) {
				group_shapes = oController.selection.groupSelection.arrGraphicObjects;
				for (i = group_shapes.length - 1; i > -1; --i) {
					if (group_shapes[i].selected && group_shapes[i].getObjectType() === AscDFH.historyitem_type_Shape) {
						content = group_shapes[i].getDocContent();
						if (content) {
							Id = content.GetSearchElementId(isNext, isRealObject(target_text_object));
							if (Id !== null) {
								return Id;
							}
						}
						group_start_index = i - 1;
					}
				}
				for (i = group_start_index; i > -1; --i) {
					if (group_shapes[i].getObjectType() === AscDFH.historyitem_type_Shape) {
						content = group_shapes[i].getDocContent();
						if (content) {
							Id = content.GetSearchElementId(isNext, false);
							if (Id !== null) {
								return Id;
							}
						}
					}
				}

				for (i = 0; i < sp_tree.length; ++i) {
					if (sp_tree[i] === oController.selection.groupSelection) {
						start_index = i - 1;
						break;
					}
				}
				if (i === sp_tree.length) {
					start_index = -1;
				}
			} else if (oController.selectedObjects.length === 0) {
				start_index = sp_tree.length - 1;
			} else {
				for (i = sp_tree.length - 1; i > -1; --i) {
					if (sp_tree[i].selected) {
						start_index = target_text_object ? i - 1 : i;
						break;
					}
				}
				if (i === sp_tree.length) {
					start_index = -1;
				}
			}
			Id = this.GetCurrentSlide().GetSearchElementId(isNext, start_index);
			if (Id !== null) {
				return Id;
			}
			for (i = this.CurPage - 1; i > -1; --i) {
				if (this.GetSlide(i).notesShape) {
					Id = this.GetSlide(i).notesShape.GetSearchElementId(isNext, false);
					if (Id !== null) {
						return Id;
					}
				}
				Id = this.GetSlide(i).GetSearchElementId(isNext, this.GetSlide(i).cSld.spTree.length - 1);
				if (Id !== null) {
					return Id;
				}
			}
			for (i = this.GetSlidesCount() - 1; i >= this.CurPage; --i) {
				if (this.GetSlide(i).notesShape) {
					Id = this.GetSlide(i).notesShape.GetSearchElementId(isNext, false);
					if (Id !== null) {
						return Id;
					}
				}
				Id = this.GetSlide(i).GetSearchElementId(isNext, this.GetSlide(i).cSld.spTree.length - 1);
				if (Id !== null) {
					return Id;
				}
			}
		}
	}
	return null;
};

CPresentation.prototype.SelectSearchElement = function (Id) {
	this.SearchEngine.Select(Id);

	this.Document_UpdateInterfaceState();
	this.Document_UpdateSelectionState();
	// this.Document_UpdateRulersState();
	this.Api.WordControl.OnUpdateOverlay();
};

CPresentation.prototype.ReplaceSearchElement = function (NewStr, bAll, Id, bInterfaceEvent) {
	var bResult = false;

	var oController = this.GetCurrentController();
	if (!oController) {
		return bResult;
	}
	var oContent = oController.getTargetDocContent();
	if (oContent) {
		oContent.RemoveSelection();
	}

	var CheckParagraphs = [];
	if (true === bAll) {
		var CheckParagraphsObj = {};
		for (var Id in this.SearchEngine.Elements) {
			CheckParagraphsObj[this.SearchEngine.Elements[Id].Get_Id()] = this.SearchEngine.Elements[Id];
		}

		for (var ParaId in CheckParagraphsObj) {
			CheckParagraphs.push(CheckParagraphsObj[ParaId]);
		}
	} else {
		if (undefined !== this.SearchEngine.Elements[Id])
			CheckParagraphs.push(this.SearchEngine.Elements[Id]);
	}

	var AllCount = this.SearchEngine.Count;


	AscCommon.History.Create_NewPoint(bAll ? AscDFH.historydescription_Document_ReplaceAll : AscDFH.historydescription_Document_ReplaceSingle);

	if (true === bAll) {
		this.SearchEngine.ReplaceAll(NewStr, true);
	} else {
		this.SearchEngine.Replace(NewStr, Id, false);

		// TODO: В будушем надо будет переделать, чтобы искалось заново только в том параграфе, в котором произошла замена
		//       Тут появляется проблема с вложенным поиском, если то что мы заменяем содержится в том, на что мы заменяем.
		if (true === this.IsTrackRevisions())
			this.SearchEngine.Reset();
	}

	this.SearchEngine.ClearOnRecalc = false;
	this.TurnOffInterfaceEvents = true;
	this.Recalculate();
	this.SearchEngine.ClearOnRecalc = true;
	this.RecalculateCurPos();
	this.TurnOffInterfaceEvents = false;
	bResult = true;

	if (true === bAll && false !== bInterfaceEvent)
		this.Api.sync_ReplaceAllCallback(AllCount, AllCount);


	return bResult;
};


CPresentation.prototype.findText = function (text, scanForward) {
	if (typeof (text) != "string") {
		return;
	}
	if (scanForward === undefined) {
		scanForward = true;
	}

	var slide_num;
	var search_select_data = null;
	if (scanForward) {
		let nCount = this.GetSlidesCount();
		for (slide_num = this.CurPage; slide_num < nCount; ++slide_num) {
			search_select_data = this.GetSlide(slide_num).graphicObjects.startSearchText(text, scanForward);
			if (search_select_data != null) {
				this.DrawingDocument.m_oWordControl.GoToPage(slide_num);
				this.GetSlide(slide_num).graphicObjects.setSelectionState(search_select_data);
				this.Document_UpdateSelectionState();
				return true;
			}
		}
		for (slide_num = 0; slide_num <= this.CurPage; ++slide_num) {
			search_select_data = this.GetSlide(slide_num).graphicObjects.startSearchText(text, scanForward, true);
			if (search_select_data != null) {
				this.DrawingDocument.m_oWordControl.GoToPage(slide_num);
				this.GetSlide(slide_num).graphicObjects.setSelectionState(search_select_data);
				this.Document_UpdateSelectionState();
				return true;
			}
		}
	} else {
		for (slide_num = this.CurPage; slide_num > -1; --slide_num) {
			search_select_data = this.GetSlide(slide_num).graphicObjects.startSearchText(text, scanForward);
			if (search_select_data != null) {
				this.DrawingDocument.m_oWordControl.GoToPage(slide_num);
				this.GetSlide(slide_num).graphicObjects.setSelectionState(search_select_data);
				this.Document_UpdateSelectionState();
				return true;
			}
		}
		for (slide_num = this.GetSlidesCount() - 1; slide_num >= this.CurPage; --slide_num) {
			search_select_data = this.GetSlide(slide_num).graphicObjects.startSearchText(text, scanForward, true);
			if (search_select_data != null) {
				this.DrawingDocument.m_oWordControl.GoToPage(slide_num);
				this.Slides[slide_num].graphicObjects.setSelectionState(search_select_data);
				this.Document_UpdateSelectionState();
				return true;
			}
		}
	}

	return false;
};

CPresentation.prototype.groupShapes = function () {
	var oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.createGroup, [], false, AscDFH.historydescription_Presentation_CreateGroup);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.unGroupShapes = function () {
	var oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.unGroupCallback, [], false, AscDFH.historydescription_Presentation_UnGroup);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.mergeSelectedShapes = function (operation) {
	const oController = this.GetCurrentController();
	if (oController)
		oController.checkSelectedObjectsAndCallback(AscFormat.mergeSelectedShapes, [operation], false, AscDFH.historydescription_Presentation_MergeSelectedShapes);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.SetThumbnailsFocusElement = function(nFocusType) {
	let oThumbnails = this.Api.WordControl.Thumbnails;
	if (oThumbnails) {
		oThumbnails.SetFocusElement(nFocusType);
	}
};

CPresentation.prototype.addImages = function (aImages, placeholder) {
	let oCurSlide = this.GetCurrentSlide();
	if (oCurSlide && aImages.length) {
		this.SetThumbnailsFocusElement(FOCUS_OBJECT_MAIN);
		this.FocusOnNotes = false;
		var oController = oCurSlide.graphicObjects;
		if (placeholder && undefined !== placeholder.id && aImages.length === 1 && aImages[0].Image) {
			var oPh = AscCommon.g_oTableId.Get_ById(placeholder.id);
			if (oPh) {
				History.Create_NewPoint(AscDFH.historydescription_Presentation_AddFlowImage);
				oController.resetSelection();
				if (oPh.isObjectInSmartArt && oPh.isObjectInSmartArt() && !this.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props, undefined, undefined, [oPh.group.getMainGroup()])) {
					const oMainGroup = oPh.group.getMainGroup();
					oPh.applyImagePlaceholderCallback(aImages, placeholder);
					oController.selectObject(oMainGroup, 0);
					oController.selection.groupSelection = oMainGroup;
					oMainGroup.selectObject(oPh, 0);
					oMainGroup.addToRecalculate();
				} else {
					var _w, _h;
					var _image = aImages[0];
					_w = oPh.extX;
					_h = oPh.extY;
					var __w = Math.max((_image.Image.width * AscCommon.g_dKoef_pix_to_mm), 1);
					var __h = Math.max((_image.Image.height * AscCommon.g_dKoef_pix_to_mm), 1);
					if (__w < _w && __h < _h) {
						_w = __w;
						_h = __h;
					} else {
						var fKoeff = Math.min(_w / __w, _h / __h);
						_w = Math.max(5, __w * fKoeff);
						_h = Math.max(5, __h * fKoeff);
					}
					var Image = oController.createImage(_image.src, oPh.x + oPh.extX / 2.0 - _w / 2.0, oPh.y + oPh.extY / 2.0 - _h / 2.0, _w, _h, _image.videoUrl, _image.audioUrl);
					if (AscFormat.isRealNumber(oPh.rot)) {
						if (Image.spPr && Image.spPr.xfrm) {
							Image.spPr.xfrm.setRot(oPh.rot);
						}
					}
					Image.setParent(oCurSlide);
					if (this.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props, undefined, undefined, [oPh])) {
						Image.addToDrawingObjects();
					} else {
						oCurSlide.replaceSp(oPh, Image);
					}
					oController.selectObject(Image, 0);
				}
				this.Recalculate();
				this.Document_UpdateInterfaceState();
				this.CheckEmptyPlaceholderNotes();
				return;
			} else {
				return;
			}
		}
		History.Create_NewPoint(AscDFH.historydescription_Presentation_AddFlowImage);
		oController.resetSelection();
		for (let i = 0; i < aImages.length; ++i) {
			let _image = aImages[i];
			if (_image.Image) {
				oController.addImage(_image.src, _image.Image.width, _image.Image.height, _image.videoUrl, _image.audioUrl);
			}
		}
		this.Recalculate();
		this.Document_UpdateInterfaceState();
		this.CheckEmptyPlaceholderNotes();
	}
};

CPresentation.prototype.AddOleObject = function (fWidth, fHeight, nWidthPix, nHeightPix, sLocalUrl, Data, sApplicationId, bSelect, arrImagesForAddToHistory) {
	let oSlide = this.GetCurrentSlide();
	if (oSlide) {
		var fPosX = (this.GetWidthMM() - fWidth) / 2;
		var fPosY = (this.GetHeightMM() - fHeight) / 2;
		var oController = oSlide.graphicObjects;
		var Image = oController.createOleObject(Data, sApplicationId, sLocalUrl, fPosX, fPosY, fWidth, fHeight, nWidthPix, nHeightPix, arrImagesForAddToHistory);
		Image.setParent(oSlide);
		Image.addToDrawingObjects();
		oController.resetSelection();
		if (bSelect !== false) {
			oController.selectObject(Image, 0);
		}
		this.Recalculate();
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.EditOleObject = function (oOleObject, sData, sImageUrl, fWidth, fHeight, nPixWidth, nPixHeight, arrImagesForAddToHistory) {
	oOleObject.editExternal(sData, sImageUrl, fWidth, fHeight, nPixWidth, nPixHeight, arrImagesForAddToHistory);
};

CPresentation.prototype.getImageDataFromSelection = function () {
	let oSlide = this.GetCurrentSlide();
	if (oSlide) {
		return oSlide.graphicObjects.getImageDataFromSelection();
	}
	return null;
};
CPresentation.prototype.putImageToSelection = function (sImageSrc, nWidth, nHeight, replaceMode) {
	let oSlide = this.GetCurrentSlide();
	if (oSlide) {
		oSlide.graphicObjects.putImageToSelection(sImageSrc, nWidth, nHeight, replaceMode);
	}
};


CPresentation.prototype.Get_AbsolutePage = function () {
	return 0;
};

CPresentation.prototype.Get_AbsoluteColumn = function () {
	return 0;
};


CPresentation.prototype.addChart = function (nTypeChart, isFromInterface, Placeholder) {
	var _this = this;
	var oSlide = _this.GetCurrentSlide();
	if (!oSlide) {
		return;
	}

	this.Api.inkDrawer.startSilentMode();
	History.Create_NewPoint(AscDFH.historydescription_Presentation_AddChart);
	this.Api.inkDrawer.endSilentMode();
	this.SetThumbnailsFocusElement(FOCUS_OBJECT_MAIN);
	_this.FocusOnNotes = false;

	AscFonts.IsCheckSymbols = true;
	const oChart = this.GetChartObject(nTypeChart, true);
	AscFonts.IsCheckSymbols = false;
	oChart.setParent(oSlide);

	var PosX = (this.GetWidthMM() - oChart.spPr.xfrm.extX) / 2;
	var PosY = (this.GetHeightMM() - oChart.spPr.xfrm.extY) / 2;
	if (Placeholder) {
		var oPh = AscCommon.g_oTableId.Get_ById(Placeholder.id);
		if (oPh) {
			PosX = oPh.x;
			PosY = oPh.y;
			oChart.spPr.xfrm.setExtX(oPh.extX);
			oChart.spPr.xfrm.setExtY(oPh.extY);
			if (this.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props, undefined, undefined, [oPh])) {
				oChart.addToDrawingObjects();
			} else {
				oSlide.replaceSp(oPh, oChart);
			}
		} else {
			return;
		}
	} else {
		oChart.addToDrawingObjects();
	}

	oChart.spPr.xfrm.setOffX(PosX);
	oChart.spPr.xfrm.setOffY(PosY);
	oSlide.graphicObjects.resetSelection();
	oSlide.graphicObjects.selectObject(oChart, 0);

	if (isFromInterface) {
		oChart.recalculate();
		AscFonts.FontPickerByCharacter.checkText("", this, function () {
			_this.Recalculate();
			_this.Document_UpdateInterfaceState();
			_this.CheckEmptyPlaceholderNotes();

			_this.DrawingDocument.m_oWordControl.OnUpdateOverlay();
		}, false, false, false);
	} else {
		_this.Recalculate();
		_this.Document_UpdateInterfaceState();
		_this.CheckEmptyPlaceholderNotes();

		this.DrawingDocument.m_oWordControl.OnUpdateOverlay();
	}
	return oChart;
};

CPresentation.prototype.RemoveSelection = function (bNoResetChartSelection) {
	var oController = this.GetCurrentController();
	if (oController) {
		oController.resetSelection(undefined, bNoResetChartSelection);
	}
};
CPresentation.prototype.CheckNotesShow = function () {
	if (this.Api) {
		var bIsShow = this.Api.getIsNotesShow();
		if (!bIsShow) {
			if (this.IsFocusOnNotes()) {
				this.FocusOnNotes = false;
				this.Document_UpdateInterfaceState();
				this.Document_UpdateSelectionState();
			}
		}
	}
};
CPresentation.prototype.CheckAnimPaneShow = function () {
	if (this.Api) {
		var bIsShow = this.Api.getIsAnimPaneShow();
		if (!bIsShow) {
			if (this.GetFocusObjType() === FOCUS_OBJECT_ANIM_PANE) {
				this.SetThumbnailsFocusElement(FOCUS_OBJECT_MAIN);
				this.Document_UpdateInterfaceState();
				this.Document_UpdateSelectionState();
			}
		}
	}
};

CPresentation.prototype.UpdateChart = function (binary) {
	const oSlide = this.GetCurrentSlide();
	if (oSlide)
	{
		oSlide.graphicObjects.updateChart(binary);
		this.Recalculate();
	}
};
CPresentation.prototype.FinalizeEditChart = function (binary)
{
	const oThis = this;
	const oSlide = this.GetCurrentSlide();
	if (oSlide)
	{
		oSlide.graphicObjects.loadChartData(binary['noHistory']);
		if (!binary['noHistory'])
		{
			oSlide.graphicObjects.checkSelectedObjectsAndCallback(function () {
				oThis.EditChart(binary);
			}, [], false, AscDFH.historydescription_Presentation_EditChart);
		}
	}
};
CPresentation.prototype.EditChart = function (binary) {
	const oSlide = this.GetCurrentSlide();
	oSlide.graphicObjects.editChart(binary);
	this.Document_UpdateInterfaceState();
};
CPresentation.prototype.GetChartObject = function (type) {
	return this.GetCurrentSlideController().getChartObject(type, undefined, undefined, true);
};

CPresentation.prototype.Check_GraphicFrameRowHeight = function (grFrame, bIgnoreHeight) {
	grFrame.recalculate();
	var oTable = grFrame.graphicObject;
	oTable.private_SetTableLayoutFixedAndUpdateCellsWidth(-1);
	var content = oTable.Content, i, j;
	for (i = 0; i < content.length; ++i) {
		var row = content[i];
		if (!bIgnoreHeight && row.Pr && row.Pr.Height && row.Pr.Height.HRule === Asc.linerule_AtLeast
			&& AscFormat.isRealNumber(row.Pr.Height.Value) && row.Pr.Height.Value > 0) {
			continue;
		}
		row.Set_Height(row.Height, Asc.linerule_AtLeast);
	}
};

CPresentation.prototype.Add_FlowTable = function (Cols, Rows, Placeholder, sStyleId) {

	var oSlide = this.GetCurrentSlide();
	if (!oSlide)
		return;

	var Width = undefined, Height = undefined, oPh, X = undefined, Y = undefined;
	var bLocked = false;
	if (Placeholder) {
		oPh = AscCommon.g_oTableId.Get_ById(Placeholder.id);
		if (oPh) {
			if (this.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props, undefined, undefined, [oPh])) {
				bLocked = true;
			}
			Width = oPh.extX;
			X = oPh.x;
			Y = oPh.y;
		} else {
			return;
		}
	}
	this.Api.inkDrawer.startSilentMode();
	History.Create_NewPoint(AscDFH.historydescription_Presentation_AddFlowTable);
	var graphic_frame = this.Create_TableGraphicFrame(Cols, Rows, oSlide, sStyleId || this.DefaultTableStyleId, Width, Height, X, Y);

	this.SetThumbnailsFocusElement(FOCUS_OBJECT_MAIN);
	this.FocusOnNotes = false;
	this.Check_GraphicFrameRowHeight(graphic_frame);
	if (oPh && !bLocked) {
		oSlide.replaceSp(oPh, graphic_frame);
	} else {
		oSlide.addToSpTreeToPos(oSlide.cSld.spTree.length, graphic_frame);
	}
	graphic_frame.Set_CurrentElement();
	graphic_frame.graphicObject.MoveCursorToStartPos();
	this.Recalculate();
	this.Document_UpdateInterfaceState();
	this.Api.inkDrawer.endSilentMode();
	return graphic_frame;
};


CPresentation.prototype.Create_TableGraphicFrame = function (Cols, Rows, Parent, StyleId, Width, Height, PosX, PosY, bInline) {
	var W;
	if (AscFormat.isRealNumber(Width)) {
		W = Width;
	} else {
		W = this.GetWidthMM() * 2 / 3;
	}
	var X, Y;
	if (AscFormat.isRealNumber(PosX) && AscFormat.isRealNumber(PosY)) {
		X = PosX;
		Y = PosY;
	} else {
		X = (this.GetWidthMM() - W) / 2;
		Y = this.GetHeightMM() / 5;
	}
	var Inline = false;
	if (AscFormat.isRealBool(bInline)) {
		Inline = bInline;
	}
	var Grid = [];

	for (var Index = 0; Index < Cols; Index++)
		Grid[Index] = W / Cols;

	var RowHeight;
	if (AscFormat.isRealNumber(Height)) {
		RowHeight = Height / Rows;
	}

	var graphic_frame = new AscFormat.CGraphicFrame();
	graphic_frame.setParent(Parent);
	graphic_frame.setSpPr(new AscFormat.CSpPr());
	graphic_frame.spPr.setParent(graphic_frame);
	graphic_frame.spPr.setXfrm(new AscFormat.CXfrm());
	graphic_frame.spPr.xfrm.setParent(graphic_frame.spPr);
	graphic_frame.spPr.xfrm.setOffX(X);
	graphic_frame.spPr.xfrm.setOffY(Y);
	graphic_frame.spPr.xfrm.setExtX(W);
	graphic_frame.spPr.xfrm.setExtY(7.478268771701388 * Rows);
	graphic_frame.setNvSpPr(new AscFormat.UniNvPr());

	var table = new CTable(this.DrawingDocument, graphic_frame, Inline, Rows, Cols, Grid, true);
	table.Reset(Inline ? X : 0, Inline ? Y : 0, W, 100000, 0, 0, 1, 0);
	if (!Inline) {
		table.Set_PositionH(Asc.c_oAscHAnchor.Page, false, 0);
		table.Set_PositionV(Asc.c_oAscVAnchor.Page, false, 0);
	}
	table.SetTableLayout(tbllayout_Fixed);
	if (typeof StyleId === "string") {
		table.Set_TableStyle(StyleId);
	}
	table.Set_TableLook(new AscCommon.CTableLook(false, true, false, false, true, false));
	for (var i = 0; i < table.Content.length; ++i) {
		var Row = table.Content[i];
		if (AscFormat.isRealNumber(RowHeight)) {
			Row.Set_Height(RowHeight, Asc.linerule_AtLeast);
		}
		//for(var j = 0; j < Row.Content.length; ++j)
		//{
		//    var cell = Row.Content[j];
		//    var props = new CTableCellPr();
		//    props.TableCellMar = {};
		//    props.TableCellMar.Top    = new CTableMeasurement(tblwidth_Mm, 1.27);
		//    props.TableCellMar.Left   = new CTableMeasurement(tblwidth_Mm, 2.54);
		//    props.TableCellMar.Bottom = new CTableMeasurement(tblwidth_Mm, 1.27);
		//    props.TableCellMar.Right  = new CTableMeasurement(tblwidth_Mm, 2.54);
		//    props.Merge(cell.Pr);
		//    cell.Set_Pr(props);
		//}
	}
	graphic_frame.setGraphicObject(table);
	graphic_frame.setBDeleted(false);
	return graphic_frame;
};

CPresentation.prototype.Set_MathProps = function (oMathProps) {

	var oController = this.GetCurrentController();
	if (oController) {
		oController.setMathProps(oMathProps);
	}
};


CPresentation.prototype.AddToParagraph = function (ParaItem, bRecalculate, noUpdateInterface) {
	let oSlide = this.GetCurrentSlide();
	if (oSlide) {
		var oMathShape = null;
		if (ParaItem.Type === para_Math) {
			var oController = oSlide.graphicObjects;
			if (!this.IsFocusOnNotes() && !(oController.selection.textSelection || (oController.selection.groupSelection && oController.selection.groupSelection.selection.textSelection))) {
				oController.resetSelection();
				oMathShape = oController.createTextArt(0, false, null, "");
				oMathShape.addToDrawingObjects();
				oMathShape.select(oController, this.CurPage);
				oController.selection.textSelection = oMathShape;
				oMathShape.txBody.content.MoveCursorToStartPos(false);
			}
		}
		if (this.IsFocusOnNotes()) {
			var oCurSlide = this.GetCurrentSlide();
			if (oCurSlide.notes) {
				oCurSlide.notes.graphicObjects.paragraphAdd(ParaItem, false);
				if (bRecalculate !== false) {
					this.Recalculate();
				}
				bRecalculate = false;
			}
		} else {
			oSlide.graphicObjects.paragraphAdd(ParaItem, false);
			if (bRecalculate !== false) {
				this.Recalculate();
			}
			var oTargetTextObject = AscFormat.getTargetTextObject(oSlide.graphicObjects);
			if (!oTargetTextObject || oTargetTextObject instanceof AscFormat.CGraphicFrame) {
				bRecalculate = false;
			}
			if (oMathShape) {
				oMathShape.checkExtentsByDocContent();
				oMathShape.spPr.xfrm.setOffX((this.GetWidthMM() - oMathShape.spPr.xfrm.extX) / 2);
				oMathShape.spPr.xfrm.setOffY((this.GetHeightMM() - oMathShape.spPr.xfrm.extY) / 2);
			}
		}
		if (false === bRecalculate) {
			this.Recalculate();
			oSlide.graphicObjects.recalculateCurPos();
			var oContent = oSlide.graphicObjects.getTargetDocContent(false, false);
			if (oContent) {
				var oCurrentParagraph = oContent.GetCurrentParagraph(true);
				if (oCurrentParagraph && oCurrentParagraph.GetType() === type_Paragraph) {
					oCurrentParagraph.CurPos.RealX = oCurrentParagraph.CurPos.X;
					oCurrentParagraph.CurPos.RealY = oCurrentParagraph.CurPos.Y;
				}
			}

		}
		//oSlide.graphicObjects.startRecalculate();
		//oSlide.graphicObjects.recalculateCurPos();
		if (!(noUpdateInterface === true) || (this.Api.asc_getKeyboardLanguage() !== -1)) {
			this.Document_UpdateInterfaceState();
			this.Document_UpdateRulersState();
		}
		this.NeedUpdateTargetForCollaboration = true;
	}
	// oSlide && oSlide.graphicObjects.checkSelectedObjectsAndCallback(oSlide.graphicObjects.paragraphAdd, [ParaItem, bRecalculate], false, AscDFH.historydescription_Presentation_ParagraphAdd, true);

};

CPresentation.prototype.ConvertMathView = function (isToLinear, isAll) {
	let oController = this.GetCurrentController();
	if (!oController) {
		return;
	}
	oController.convertMathView(isToLinear, isAll);
	this.UpdateSelection();
	this.UpdateInterface();
};

CPresentation.prototype.ClearParagraphFormatting = function (isClearParaPr, isClearTextPr) {
	var oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.paragraphClearFormatting, [isClearParaPr, isClearTextPr], false, AscDFH.historydescription_Presentation_ParagraphClearFormatting);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.GetSelectedBounds = function () {
	var oController = this.GetCurrentController();
	if (oController && oController.selectedObjects.length > 0) {
		return oController.getBoundsForGroup([oController.selectedObjects[0]]);
	}
	return new AscFormat.CGraphicBounds(0, 0, 0, 0);
};

CPresentation.prototype.GetFocusObjType = function () {
	if (!window["NATIVE_EDITOR_ENJINE"] && this.Api.WordControl.Thumbnails) {
		return this.Api.WordControl.Thumbnails.FocusObjType;
	} else {
		var oCurController = this.GetCurrentController();
		if (oCurController) {
			return oCurController.selectedObjects.length > 0 ? FOCUS_OBJECT_MAIN : FOCUS_OBJECT_THUMBNAILS;
		}
		return FOCUS_OBJECT_THUMBNAILS;
	}
};
CPresentation.prototype.GetSelectedSlides = function () {
	if (!window["NATIVE_EDITOR_ENJINE"] && this.Api.WordControl.Thumbnails) {
		return this.Api.WordControl.Thumbnails.GetSelectedArray();
	} else {
		//todo: check if this.CurPage is -1
		return [this.CurPage];
	}
};
CPresentation.prototype.GetSelectedSlideObjects = function () {
	let aSlideObjects = [];
	let aIdx = this.GetSelectedSlides();
	for(let nIdx = 0; nIdx < aIdx.length; ++nIdx) {
		let oSlide = this.GetSlide(aIdx[nIdx]);
		if(oSlide) {
			aSlideObjects.push(this.GetSlide(aIdx[nIdx]));
		}
	}
	return aSlideObjects;
};
CPresentation.prototype.GetAllSlideIndexes = function () {
	let aIdx = [];
	let nCount = this.GetSlidesCount();
	for(let nIdx = 0; nIdx < nCount; ++nIdx) {
		aIdx.push(nIdx);
	}
	return aIdx;
};

CPresentation.prototype.RemoveCurrentComment = function (isMine) {
	if (!this.IsFocusOnNotes()) {
		var oCurSlide = this.GetCurrentSlide();
		if (oCurSlide && oCurSlide.slideComments) {
			var oSelectedComment = oCurSlide.slideComments.getSelectedComment();
			if (oSelectedComment) {
				var aCommentData = [{comment: oSelectedComment, slide: oCurSlide}];
				if (isMine) {
					if (this.Api) {
						var oDocInfo = this.Api.DocInfo;
						if (oDocInfo) {
							var sUserId = oDocInfo.get_UserId();
							if (oSelectedComment.hasUserData(sUserId) && oSelectedComment.canBeDeleted()) {
								if (this.Document_Is_SelectionLocked(AscCommon.changestype_MoveComment, aCommentData, this.IsEditCommentsMode()) === false) {
									this.Create_NewHistoryPoint(AscDFH.historydescription_Presentation_RemoveComment);
									if (oSelectedComment.isUserComment(sUserId)) {
										oCurSlide.slideComments.removeComment(oSelectedComment.Get_Id());
										this.Api.sync_HideComment();
									} else {
										oSelectedComment.removeUserReplies(sUserId);
									}
									this.Recalculate();
									return true;
								}
							}
						}
					}
				} else {
					if (this.Document_Is_SelectionLocked(AscCommon.changestype_MoveComment, aCommentData, this.IsEditCommentsMode()) === false) {
						this.Create_NewHistoryPoint(AscDFH.historydescription_Presentation_RemoveComment);
						this.RemoveComment(oSelectedComment.Id, true);
						return true;
					}
				}
			}
		}
	}
	return false;
};

CPresentation.prototype.SendRemoveCommentEvent = function () {
	if (!this.Api) {
		return false;
	}
	if (!this.IsFocusOnNotes()) {
		var oCurSlide = this.GetCurrentSlide();
		if (oCurSlide && oCurSlide.slideComments) {
			var oSelectedComment = oCurSlide.slideComments.getSelectedComment();
			if (oSelectedComment) {
				this.Api.asc_onDeleteComment(oSelectedComment.Id, oSelectedComment.Data);
				return true;
			}
		}
	}
	return false;
};


CPresentation.prototype.RemoveMyComments = function () {
	var aAllMyComments = [];
	this.GetAllMyComments(aAllMyComments);
	if (this.Document_Is_SelectionLocked(AscCommon.changestype_MoveComment, aAllMyComments, this.IsEditCommentsMode()) === false) {
		this.Create_NewHistoryPoint(AscDFH.historydescription_Presentation_RemoveComment);
		this.comments.removeMyComments();
		for (var i = 0; i < this.Slides.length; ++i) {
			this.Slides[i].removeMyComments();
		}
		this.Recalculate();
	}
};

CPresentation.prototype.GetAllMyComments = function (aAllComments) {
	this.comments.getAllMyComments(aAllComments, null);
	for (var i = 0; i < this.Slides.length; ++i) {
		this.Slides[i].getAllMyComments(aAllComments);
	}
};

CPresentation.prototype.CheckFormAutoFit = function (oForm) {
};
CPresentation.prototype.OnChangeForm = function (oForm) {
};
CPresentation.prototype.OnChangeContentControl = function (oCC) {
};

CPresentation.prototype.RemoveAllComments = function () {
	var aAllMyComments = [];
	this.GetAllComments(aAllMyComments);
	if (this.Document_Is_SelectionLocked(AscCommon.changestype_MoveComment, aAllMyComments, this.IsEditCommentsMode()) === false) {
		this.Create_NewHistoryPoint(AscDFH.historydescription_Presentation_RemoveComment);
		this.comments.removeAllComments();
		for (var i = 0; i < this.Slides.length; ++i) {
			this.Slides[i].removeAllComments();
		}
		this.Recalculate();
	}
};
CPresentation.prototype.ResolveAllComments = function (isMine, isCurrent, arrIds) {
	var aAllMyComments = [];
	this.GetAllComments(aAllMyComments, isMine, isCurrent, arrIds);
	if (this.Document_Is_SelectionLocked(AscCommon.changestype_MoveComment, aAllMyComments, this.IsEditCommentsMode()) === false) {
		this.Create_NewHistoryPoint(AscDFH.historydescription_Presentation_RemoveComment);
		for (var nComment = 0; nComment < aAllMyComments.length; ++nComment) {
			var oComment = aAllMyComments[nComment].comment;
			if (oComment && !oComment.IsSolved() && AscCommon.UserInfoParser.canEditComment(oComment.GetUserName())) {
				if (oComment.Data) {
					var oCopyData = oComment.Data.createDuplicate(false);
					oCopyData.Set_Solved(true);
					oComment.Set_Data(oCopyData);
					this.Api.sync_ChangeCommentData(oComment.Id, oCopyData);
				}
			}
		}
		this.Recalculate();
	}
};
CPresentation.prototype.GetAllComments = function (aAllComments, isMine, isCurrent, aIds) {
	let aResult = aAllComments;
	if(!Array.isArray(aAllComments)) {
		aResult = [];
	}
	this.comments.getAllComments(aResult, isMine, isCurrent, aIds);
	for (var i = 0; i < this.Slides.length; ++i) {
		this.Slides[i].getAllComments(aResult, isMine, isCurrent, aIds);
	}
	return aResult;
};

CPresentation.prototype.Remove = function (Count, bOnlyText, bRemoveOnlySelection, bOnTextAdd, isWord) {
	if (this.GetFocusObjType() === FOCUS_OBJECT_THUMBNAILS) {
		this.deleteSlides(this.GetSelectedSlideObjects());
		return;
	}
	if ("undefined" === typeof (bRemoveOnlySelection))
		bRemoveOnlySelection = false;

	var oController = this.GetCurrentController();
	if (this.SendRemoveCommentEvent()) {
		return;
	}
	if (oController) {
		if (oController.selectedObjects.length !== 0) {
			oController.remove(Count, bOnlyText, bRemoveOnlySelection, bOnTextAdd, isWord);
			this.Document_UpdateInterfaceState();

		} else {
			var aAnims = oController.getAnimSelectionState();
			if (aAnims.length > 0) {
				var oTiming = this.GetCurTiming();
				if (oTiming) {
					if (this.IsSelectionLocked(AscCommon.changestype_Timing) === false) {
						AscCommon.History.Create_NewPoint(0);
						oTiming.removeSelectedEffects();
						this.Recalculate();
					}
				}
			}
		}
	}
};


CPresentation.prototype.MoveCursorToStartPos = function () {
	var oController = this.GetCurrentController();
	oController && oController.cursorMoveToStartPos();
	this.private_UpdateCursorXY(true, true);
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
	return true;
};

CPresentation.prototype.MoveCursorToEndPos = function () {
	var oController = this.GetCurrentController();
	oController && oController.cursorMoveToEndPos();
	this.private_UpdateCursorXY(true, true);
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
	return true;
};

CPresentation.prototype.MoveCursorLeft = function (AddToSelect, Word) {
	var oController = this.GetCurrentController();
	oController && oController.cursorMoveLeft(AddToSelect, Word);
	this.private_UpdateCursorXY(true, true);
	this.Document_UpdateInterfaceState();
	return true;
};

CPresentation.prototype.MoveCursorRight = function (AddToSelect, Word) {
	var oController = this.GetCurrentController();
	oController && oController.cursorMoveRight(AddToSelect, Word);
	this.private_UpdateCursorXY(true, true);
	this.Document_UpdateInterfaceState();
	return true;
};

CPresentation.prototype.MoveCursorUp = function (AddToSelect, CtrlKey) {
	var oController = this.GetCurrentController();
	oController && oController.cursorMoveUp(AddToSelect, CtrlKey);
	this.private_UpdateCursorXY(true, true);
	this.Document_UpdateInterfaceState();
	return true;
};

CPresentation.prototype.MoveCursorDown = function (AddToSelect, CtrlKey) {
	var oController = this.GetCurrentController();
	oController && oController.cursorMoveDown(AddToSelect, CtrlKey);
	this.private_UpdateCursorXY(true, true);
	this.Document_UpdateInterfaceState();
	return true;
};

CPresentation.prototype.MoveCursorToEndOfLine = function (AddToSelect) {
	var oController = this.GetCurrentController();
	oController && oController.cursorMoveEndOfLine(AddToSelect);
	this.private_UpdateCursorXY(true, true);
	this.Document_UpdateInterfaceState();
	return true;
};

CPresentation.prototype.MoveCursorToStartOfLine = function (AddToSelect) {
	var oController = this.GetCurrentController();
	oController && oController.cursorMoveStartOfLine(AddToSelect);
	this.private_UpdateCursorXY(true, true);
	this.Document_UpdateInterfaceState();
	return true;
};

CPresentation.prototype.MoveCursorToXY = function (X, Y, AddToSelect) {
	var oController = this.GetCurrentController();
	oController && oController.cursorMoveAt(X, Y, AddToSelect);
	this.private_UpdateCursorXY(true, true);
	this.Document_UpdateInterfaceState();
	return true;
};

CPresentation.prototype.MoveCursorToCell = function (bNext) {

};

/**
 * Проверяем будет ли добавление текста на ивенте KeyDown
 * @param e
 * @returns {Number[]} Массив юникодных значений
 */
CPresentation.prototype.GetAddedTextOnKeyDown = function (e) {
	if (e.KeyCode === 32) // Space
	{
		var oController = this.GetCurrentController();
		if (oController) {

			var oTargetDocContent = oController.getTargetDocContent();
			if (oTargetDocContent) {
				var oSelectedInfo = new CSelectedElementsInfo();
				oTargetDocContent.GetSelectedElementsInfo(oSelectedInfo);
				var oMath = oSelectedInfo.GetMath();

				if (!oMath) {
					if (e.ShiftKey && e.CtrlKey)
						return [0x00A0];
				}
			}
		}
	} else if (e.KeyCode === 69 && e.CtrlKey) // Ctrl + E + ...
	{
		if (e.AltKey) // Ctrl + Alt + E - добавляем знак евро €
			return [0x20AC];
	} else if ((e.KeyCode === 189 || e.KeyCode === 173)) // Клавиша Num-
	{
		if (e.CtrlKey && e.ShiftKey)
			return [0x2013];
	}

	return [];
};

CPresentation.prototype.ConvertEquationToMath = function (oEquation, isAll) {
	var aEquations = [];
	if (isAll) {
		var aSlides = this.Slides;
		for (var nSlide = 0; nSlide < aSlides.length; ++nSlide) {
			var oSlide = aSlides[nSlide];
			var aSpTree = oSlide.cSld.spTree;
			for (var nSp = 0; nSp < aSpTree.length; ++nSp) {
				aSpTree[nSp].collectEquations3(aEquations);
			}
		}

	} else {
		aEquations.push(oEquation);
	}
	if (aEquations.length > 0) {
		var nEquation;
		var aObjectsForCheck = [];
		for (nEquation = 0; nEquation < aEquations.length; ++nEquation) {
			aObjectsForCheck.push(aEquations[nEquation].getMainGroup() || aEquations[nEquation]);
		}
		if (!this.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props, undefined, undefined, aObjectsForCheck)) {
			AscCommon.History.Create_NewPoint(0);
			for (nEquation = 0; nEquation < aEquations.length; ++nEquation) {
				aEquations[nEquation].replaceToMath();
			}
			this.Recalculate();
		}
	}
};

CPresentation.prototype.SetParagraphBidi = function(isRtl)
{
	let oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.setParagraphBidi, [isRtl], false, AscDFH.historydescription_Document_SetParagraphBidi);
	this.Document_UpdateInterfaceState();
};
CPresentation.prototype.SetParagraphAlign = function (Align) {

	var oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.setParagraphAlign, [Align], false, AscDFH.historydescription_Presentation_SetParagraphAlign);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.SetParagraphSpacing = function (Spacing) {
	var oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.setParagraphSpacing, [Spacing], false, AscDFH.historydescription_Presentation_SetParagraphSpacing);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.SetParagraphTabs = function (Tabs) {
	var oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.setParagraphTabs, [Tabs], false, AscDFH.historydescription_Presentation_SetParagraphTabs);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.SetParagraphIndent = function (Ind) {
	var oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.setParagraphIndent, [Ind], false, AscDFH.historydescription_Presentation_SetParagraphIndent);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.SetParagraphNumbering = function (oBullet) {
	var oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.setParagraphNumbering, [oBullet], false, AscDFH.historydescription_Presentation_SetParagraphNumbering);
	this.Document_UpdateInterfaceState();   //TODO
};
CPresentation.prototype.SetParagraphHighlight = function (IsColor, r, g, b) {
	var oController = this.GetCurrentController();
	var oPresentation = this;
	if (oController) {
		var oTargetContent = oController.getTargetDocContent();
		if (!oTargetContent || oTargetContent.IsSelectionUse() && !oTargetContent.IsSelectionEmpty()) {
			oController.checkSelectedObjectsAndCallback(function () {
				if (false === IsColor) {
					oPresentation.AddToParagraph(new ParaTextPr({HighlightColor: null}));
				} else {
					oPresentation.AddToParagraph(new ParaTextPr({HighlightColor: AscFormat.CreateUniColorRGB(r, g, b)}));
				}
			}, [], false, AscDFH.historydescription_Document_SetTextHighlight);
		} else {
			if (false === IsColor) {
				oPresentation.HighlightColor = null;
			} else {
				oPresentation.HighlightColor = new AscCommonWord.CDocumentColor(r, g, b, false);
			}
		}
	}
};
CPresentation.prototype.IncreaseDecreaseFontSize = function (bIncrease) {
	var oController = this.GetCurrentController();
	var oPresentation = this;
	oController && oController.checkSelectedObjectsAndCallback(
		function () {
			oController.paragraphIncDecFontSize(bIncrease);
		}
		, [], false, AscDFH.historydescription_Presentation_ParagraphIncDecFontSize);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.IncreaseDecreaseIndent = function (bIncrease) {
	var oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.paragraphIncDecIndent, [bIncrease], false, AscDFH.historydescription_Presentation_ParagraphIncDecIndent);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.Can_IncreaseParagraphLevel = function (bIncrease) {
	var oController = this.GetCurrentController();
	return oController && oController.canIncreaseParagraphLevel(bIncrease);
};

CPresentation.prototype.SetImageProps = function (Props) {
	var oController = this.GetCurrentController();
	if (!oController) {
		return;
	}
	var aAdditionalObjects = null;
	if (AscFormat.isRealNumber(Props.Width) && AscFormat.isRealNumber(Props.Height)) {
		aAdditionalObjects = oController.getConnectorsForCheck2();
	}
	oController.checkSelectedObjectsAndCallback(oController.applyDrawingProps, [Props], false, AscDFH.historydescription_Presentation_SetImageProps, aAdditionalObjects);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.ShapeApply = function (shapeProps) {
	var oController = this.GetCurrentController();
	if (!oController) {
		return;
	}
	var aAdditionalObjects = null;
	if (AscFormat.isRealNumber(shapeProps.Width) && AscFormat.isRealNumber(shapeProps.Height)) {
		aAdditionalObjects = oController.getConnectorsForCheck2();
	}
	oController.checkSelectedObjectsAndCallback(oController.applyDrawingProps, [shapeProps], false, AscDFH.historydescription_Presentation_SetShapeProps, aAdditionalObjects);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.ChartApply = function (chartProps) {
	var oController = this.GetCurrentController();
	if (!oController) {
		return;
	}
	var aAdditionalObjects = null;
	if (AscFormat.isRealNumber(chartProps.Width) && AscFormat.isRealNumber(chartProps.Height)) {
		aAdditionalObjects = oController.getConnectorsForCheck2();
	}
	oController.checkSelectedObjectsAndCallback(oController.applyDrawingProps, [chartProps], false, AscDFH.historydescription_Presentation_ChartApply, aAdditionalObjects);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.changeShapeType = function (shapeType) {
	var oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.applyDrawingProps, [{type: shapeType}], false, AscDFH.historydescription_Presentation_ChangeShapeType);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.setVerticalAlign = function (align) {
	var oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.applyDrawingProps, [{verticalTextAlign: align}], false, AscDFH.historydescription_Presentation_SetVerticalAlign);
	this.Document_UpdateInterfaceState();
};
CPresentation.prototype.setVert = function (align) {
	var oController = this.GetCurrentController();
	oController && oController.checkSelectedObjectsAndCallback(oController.applyDrawingProps, [{vert: align}], false, AscDFH.historydescription_Presentation_SetVert);
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.Get_Styles = function () {
	var styles = AscWord.DEFAULT_STYLES;
	return {styles: styles, lastId: styles.Get_Default_Paragraph()}
};
CPresentation.prototype.GetStyles = function () {
	return this.Get_Styles();
};

CPresentation.prototype.IsTableCellContent = function (isReturnCell) {
	if (true === isReturnCell)
		return null;

	return false;
};

CPresentation.prototype.Check_AutoFit = function () {
	return false;
};


CPresentation.prototype.Get_Theme = function () {
	let oCurSlide = this.GetCurrentSlide();
	if(!oCurSlide) return this.slideMasters[0] && this.slideMasters[0].Theme;
	if(this.IsFocusOnNotes()) {
		return oCurSlide.notes.getTheme();
	}
	return oCurSlide.getTheme();
};

CPresentation.prototype.Get_ColorMap = function () {
	let oCurSlide = this.GetCurrentSlide();
	if(!oCurSlide) return AscFormat.GetDefaultColorMap();
	if(this.IsFocusOnNotes()) {
		return oCurSlide.notes.getColorMap();
	}
	return oCurSlide.getColorMap();
};

CPresentation.prototype.Get_PageFields = function () {
	return {X: 0, Y: 0, XLimit: 2000, YLimit: 2000};
};

CPresentation.prototype.Get_PageLimits = function (PageIndex) {
	return this.Get_PageFields();
};

CPresentation.prototype.CheckRange = function () {
	return [];
};

CPresentation.prototype.GetCursorRealPosition = function () {
	return {
		X: this.CurPosition.X,
		Y: this.CurPosition.Y
	};
};

CPresentation.prototype.Viewer_OnChangePosition = function () {
	var oSlide = this.GetCurrentSlide();
	if (oSlide && oSlide.slideComments && Array.isArray(oSlide.slideComments.comments)) {
		var aComments = oSlide.slideComments.comments;
		for (var i = aComments.length - 1; i > -1; --i) {
			if (aComments[i].selected) {
				var Coords = this.DrawingDocument.ConvertCoordsToCursorWR_Comment(aComments[i].x, aComments[i].y);
				this.Api.sync_UpdateCommentPosition(aComments[i].Get_Id(), Coords.X, Coords.Y);
				break;
			}
		}
	}
	AscCommon.g_specialPasteHelper.SpecialPasteButton_Update_Position();
	this.MathTrackHandler.OnChangePosition();
};

CPresentation.prototype.IsCell = function (isReturnCell) {
	if (isReturnCell)
		return null;

	return false;
};

CPresentation.prototype.GetPrevElementEndInfo = function (CurElement) {
	return null;
};
CPresentation.prototype.Get_TextBackGroundColor = function () {
	return new CDocumentColor(255, 255, 255, false);
};


CPresentation.prototype.SetTableProps = function (Props) {

	const oController = this.GetCurrentController();
	if (oController) {
		oController.setTableProps(Props);

		if (!this.IsFocusOnNotes()) {
			oController.updateConnectors(false);
		}
		this.Recalculate();

		this.Document_UpdateInterfaceState();
		this.Document_UpdateSelectionState();
	}
};

CPresentation.prototype.GetCalculatedParaPr = function () {
	var oController = this.GetCurrentController();
	if (oController) {
		var ret = oController.getParagraphParaPr();
		if (ret) {
			return ret;
		}
	}
	return new CParaPr();
};

CPresentation.prototype.GetCalculatedTextPr = function () {
	var oController = this.GetCurrentController();
	if (oController) {
		var ret = oController.getParagraphTextPr();
		if (ret) {
			var oTheme = oController.getTheme();
			if (oTheme) {
				ret.ReplaceThemeFonts(oTheme.themeElements.fontScheme);
			}
			return ret;
		}
	}
	return new CTextPr();
};

CPresentation.prototype.GetDirectTextPr = function () {
	var oController = this.GetCurrentController();
	if (oController) {
		return oController.getParagraphTextPr();
	}
	return new CTextPr();
};

CPresentation.prototype.GetDirectParaPr = function () {
	var oController = this.GetCurrentController();
	if (oController) {
		return oController.getParagraphParaPr();
	}
	return new CParaPr();
};


CPresentation.prototype.GetTableStyleIdMap = function (oMap) {
	var oSlide;
	for (let i = 0; i < this.Slides.length; ++i) {
		oSlide = this.Slides[i];
		this.CollectStyleId(oMap, oSlide.cSld.spTree);
		if (oSlide.notes) {
			this.CollectStyleId(oMap, oSlide.notes.cSld.spTree);
		}
	}
	for(let nIdx = 0; nIdx < this.slideMasters.length; ++nIdx) {
		let oMaster = this.slideMasters[nIdx];
		this.CollectStyleId(oMap, oMaster.cSld.spTree);
		for(let nLt = 0; nLt < oMaster.sldLayoutLst.length; ++nLt) {
			this.CollectStyleId(oMap, oMaster.sldLayoutLst[nLt].cSld.spTree);
		}
	}
};

CPresentation.prototype.CollectStyleId = function (oMap, aSpTree) {
	for (var i = 0; i < aSpTree.length; ++i) {
		if (aSpTree[i].getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
			if (isRealObject(aSpTree[i].graphicObject) && typeof aSpTree[i].graphicObject.TableStyle === "string" && isRealObject(g_oTableId.Get_ById(aSpTree[i].graphicObject.TableStyle))) {
				var oStyle = AscCommon.g_oTableId.Get_ById(aSpTree[i].graphicObject.TableStyle);
				if (oStyle instanceof CStyle) {
					oMap[aSpTree[i].graphicObject.TableStyle] = true;
				}
			}
		} else if (aSpTree[i].isGroupObject()) {
			this.CollectStyleId(oMap, aSpTree[i].spTree);
		}
	}
};

// Обновляем данные в интерфейсе о свойствах параграфа
CPresentation.prototype.Interface_Update_ParaPr = function () {
	var oController = this.GetCurrentController();
	if (oController) {
		var ParaPr = oController.getPropsArrays().paraPr;

		if (null != ParaPr) {
			if (undefined != ParaPr.Tabs) {
				var DefaultTab = ParaPr.DefaultTab != null ? ParaPr.DefaultTab : AscCommonWord.Default_Tab_Stop;
				this.Api.Update_ParaTab(DefaultTab, ParaPr.Tabs);
			}

			this.Api.UpdateParagraphProp(ParaPr);
		}
	}
};

// Обновляем данные в интерфейсе о свойствах текста
CPresentation.prototype.Interface_Update_TextPr = function () {
	var oController = this.GetCurrentController();
	if (oController) {
		var TextPr = oController.getPropsArrays().textPr;

		if (null != TextPr)
			this.Api.UpdateTextPr(TextPr);
	}
};

CPresentation.prototype.IsVisibleSlide = function (nIndex) {
	var oSlide = this.GetSlide(nIndex);
	if (!oSlide) {
		return false;
	}
	return oSlide.isVisible();
};


CPresentation.prototype.hideSlides = function (isHide, aSlides) {

	var aSelectedArray;
	if (Array.isArray(aSlides)) {
		aSelectedArray = aSlides;
	} else {
		aSelectedArray = this.GetSelectedSlides();
	}
	const arrSlides = [];
	for (let i = 0; i < aSelectedArray.length; i++) {
		arrSlides.push(this.Slides[aSelectedArray[i]]);
	}
	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_SlideHide, arrSlides)) {
		History.Create_NewPoint(AscDFH.historydescription_Presentation_HideSlides);
		var bShow = !isHide;
		var oSlide;
		var nIndex;
		for (var i = 0; i < aSelectedArray.length; ++i) {
			nIndex = aSelectedArray[i];
			oSlide = this.Slides[nIndex];
			if (oSlide) {
				oSlide.setShow(bShow);
				this.DrawingDocument.OnRecalculateSlide(nIndex);//need only for update index label in thumbnails; TODO: remove it
			}
		}
		this.DrawingDocument.OnEndRecalculate(false, false);
		this.Document_UpdateUndoRedoState();
	}
};


CPresentation.prototype.SelectAll = function () {

	var oController = this.GetCurrentController();
	if (oController) {
		oController.selectAll();
		this.Document_UpdateInterfaceState();
		this.Api.sendEvent("asc_onSelectionEnd");
	}
};


CPresentation.prototype.UpdateCursorType = function (X, Y, MouseEvent) {

	const oApi = this.Api;
	const isDrawHandles = oApi ? oApi.isShowShapeAdjustments() : true;
	const isMarkerFormat = oApi ? oApi.isMarkerFormat : false;

	const oController = this.GetCurrentController();
	if (oController) {

		let graphicObjectInfo;
		if(isMarkerFormat) {
			oController.noNeedUpdateCursorType = true;
			graphicObjectInfo = oController.isPointInDrawingObjects(X, Y, MouseEvent);
			oController.noNeedUpdateCursorType = false;
		}
		else {
			graphicObjectInfo = oController.isPointInDrawingObjects(X, Y, MouseEvent);
		}
		if (graphicObjectInfo) {
			if (!graphicObjectInfo.updated) {
				if (isDrawHandles !== false) {
					if(isMarkerFormat && graphicObjectInfo.cursorType === "text") {
						this.DrawingDocument.SetCursorType(AscCommon.Cursors.MarkerFormat);
					}
					else {
						this.DrawingDocument.SetCursorType(graphicObjectInfo.cursorType);
					}
				} else {
					this.DrawingDocument.SetCursorType("default");
				}
			}
		} else {
			this.DrawingDocument.SetCursorType("default");
		}
		AscCommon.CollaborativeEditing.Check_ForeignCursorsLabels(X, Y, this.CurPage);
	}
};


CPresentation.prototype.OnKeyDown = function (e) {
	var bUpdateSelection = true;
	var bRetValue = keydownresult_PreventNothing;

	this.Api.sendEvent("asc_onBeforeKeyDown", e);
	if (this.StopAnimationPreview()) {
		return keydownresult_PreventAll;
	}
	// Сбрасываем текущий элемент в поиске
	if (this.SearchEngine.Count > 0)
		this.SearchEngine.ResetCurrent();


	let nStartHistoryIndex = this.History.Index;
	var oController = this.GetCurrentController();
	var aStartAnims = [];
	if (oController) {
		aStartAnims = oController.getAnimSelectionState();
	}
	const bIsMacOs = AscCommon.AscBrowser.isMacOs;
	var nShortcutAction = this.Api.getShortcut(e);
	switch (nShortcutAction) {
		case Asc.c_oAscPresentationShortcutType.EditSelectAll: {
			this.SelectAll();
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.EditUndo: {
			if (this.CanEdit() || this.IsEditCommentsMode()) {
				this.Document_Undo();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.EditRedo: {
			if (this.CanEdit() || this.IsEditCommentsMode()) {
				this.Document_Redo();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.Duplicate: {
			if (this.CanEdit()) {
				if (oController) {
					if (oController.selectedObjects.length > 0) {
						this.Create_NewHistoryPoint(AscDFH.historydescription_Document_SetParagraphAlignHotKey);
						this.GetCurrentSlide().copySelectedObjects();
						this.Recalculate();
						this.Document_UpdateInterfaceState();
					} else {
						this.DublicateSlide();
					}
				}
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.Print: {
			this.Api.onPrint();
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.Save: {
			if (!this.IsViewMode()) {
				this.Api.asc_Save();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.ShowContextMenu: {
			if (this.GetFocusObjType() === FOCUS_OBJECT_MAIN) {
				if (oController) {
					var oPosition = oController.getContextMenuPosition(0);
					var ConvertedPos = this.DrawingDocument.ConvertCoordsToCursorWR(oPosition.X, oPosition.Y, this.CurPage);
					this.Api.sync_ContextMenuCallback(new AscCommonSlide.CContextMenuData({
						Type: Asc.c_oAscContextMenuTypes.Main,
						X_abs: ConvertedPos.X,
						Y_abs: ConvertedPos.Y
					}));
				}
			}

			bUpdateSelection = false;
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.ShowParaMarks: {
			editor.ShowParaMarks = !editor.ShowParaMarks;
			let oSlide = this.GetCurrentSlide();
			if (oSlide) {
				this.DrawingDocument.OnRecalculateSlide(this.CurPage);
				if (oSlide.notes) {
					this.DrawingDocument.Notes_OnRecalculate(this.CurPage, oSlide.NotesWidth, oSlide.getNotesHeight());
				}
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.Bold: {
			var TextPr = this.GetCalculatedTextPr();
			if (null != TextPr && this.CanEdit()) {
				if (this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
					History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
					this.AddToParagraph(new ParaTextPr({Bold: TextPr.Bold === true ? false : true}));
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscPresentationShortcutType.CopyFormat: {
			this.Document_Format_Copy();
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.CenterAlign: {
			var ParaPr = this.GetCalculatedParaPr();
			if (null != ParaPr && ParaPr.Jc !== AscCommon.align_Center) {
				this.Create_NewHistoryPoint(AscDFH.historydescription_Document_SetParagraphAlignHotKey);
				this.SetParagraphAlign(AscCommon.align_Center);
				this.Document_UpdateInterfaceState();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.EuroSign: {
			if (true === this.CollaborativeEditing.Is_Fast() || this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
				History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
				this.AddToParagraph(new AscWord.CRunText("€".charCodeAt(0)));
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.Group: {
			if (this.CanEdit()) {
				this.groupShapes();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.UnGroup: {
			if (this.CanEdit()) {
				this.unGroupShapes();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.Italic: {
			var TextPr = this.GetCalculatedTextPr();
			if (null != TextPr) {
				if (this.CanEdit() && this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
					History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
					this.AddToParagraph(new ParaTextPr({Italic: TextPr.Italic === true ? false : true}));
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscPresentationShortcutType.JustifyAlign: {
			var ParaPr = this.GetCalculatedParaPr();
			if (null != ParaPr && this.CanEdit() && ParaPr.Jc !== align_Justify) {
				this.SetParagraphAlign(align_Justify);
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.AddHyperlink: {
			if (this.CanEdit() && true === this.CanAddHyperlink(false))
				editor.sync_DialogAddHyperlink();

			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.BulletList: {
			if (this.CanEdit()) {
				const oCalcParaPr = this.GetCalculatedParaPr();
				let oListType;
				if (oCalcParaPr && oCalcParaPr.Bullet)
				{
					oListType = AscFormat.fGetListTypeFromBullet(oCalcParaPr.Bullet)
				}
				let oBullet;
				if (oListType && oListType.Type === 0 && oListType.SubType === 1)
				{
					oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: -1, SubType: -1});
				}
				else
				{
					oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: 0, SubType: 1});
				}
				this.SetParagraphNumbering(oBullet);
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.LeftAlign: {
			var ParaPr = this.GetCalculatedParaPr();
			if (null != ParaPr) {
				if (this.CanEdit() && ParaPr.Jc !== align_Left) {
					this.SetParagraphAlign(align_Left);
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscPresentationShortcutType.RightAlign: {
			var ParaPr = this.GetCalculatedParaPr();
			if (null != ParaPr) {
				if (this.CanEdit() && ParaPr.Jc !== AscCommon.align_Right) {
					this.SetParagraphAlign(AscCommon.align_Right);
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscPresentationShortcutType.Underline: {
			var TextPr = this.GetCalculatedTextPr();
			if (null != TextPr) {
				if (this.CanEdit() && this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
					History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
					this.AddToParagraph(new ParaTextPr({Underline: TextPr.Underline === true ? false : true}));
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscPresentationShortcutType.Strikethrough: {
			var TextPr = this.GetCalculatedTextPr();
			if (null != TextPr) {
				if (this.CanEdit() && this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
					History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
					this.AddToParagraph(new ParaTextPr({Strikeout: TextPr.Strikeout !== true}));
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscPresentationShortcutType.PasteFormat: {
			if (this.CanEdit()) {
				this.Document_Format_Paste();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.Superscript: {
			var TextPr = this.GetCalculatedTextPr();
			if (null != TextPr) {
				if (this.CanEdit() && this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
					History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
					this.AddToParagraph(new ParaTextPr({VertAlign: TextPr.VertAlign === AscCommon.vertalign_SuperScript ? vertalign_Baseline : AscCommon.vertalign_SuperScript}));
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscPresentationShortcutType.Subscript: {
			var TextPr = this.GetCalculatedTextPr();
			if (null != TextPr) {
				if (this.CanEdit() && this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
					History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
					this.AddToParagraph(new ParaTextPr({VertAlign: TextPr.VertAlign === AscCommon.vertalign_SubScript ? vertalign_Baseline : AscCommon.vertalign_SubScript}));
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscPresentationShortcutType.EnDash: {
			if (true === this.CollaborativeEditing.Is_Fast() || this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
				if (this.CanEdit()) {
					this.DrawingDocument.TargetStart();
					this.DrawingDocument.TargetShow();

					History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);

					var Item = new AscWord.CRunText(0x2013);
					Item.SpaceAfter = false;

					this.AddToParagraph(Item);
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscPresentationShortcutType.DecreaseFont: {
			if (this.CanEdit()) {
				editor.FontSizeOut();
				this.Document_UpdateInterfaceState();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.IncreaseFont: {
			if (this.CanEdit()) {
				editor.FontSizeIn();
				this.Document_UpdateInterfaceState();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscPresentationShortcutType.SpeechWorker: {
			AscCommon.EditorActionSpeaker.toggle();
			bRetValue = keydownresult_PreventAll;
			break;
		}
		default: {
			var oCustom = this.Api.getCustomShortcutAction(nShortcutAction);
			if (oCustom) {
				if (oController.getTargetDocContent(false, false)) {
					if (AscCommon.c_oAscCustomShortcutType.Symbol === oCustom.Type) {
						this.Api["asc_insertSymbol"](oCustom.Font, oCustom.CharCode);
					}
				}

			}
			break;
		}
	}

	if (!nShortcutAction) {
		if (e.KeyCode === 8) {// BackSpace
			if (this.CanEdit()) {
				const bIsWord = bIsMacOs ? e.AltKey : e.CtrlKey;
				this.Remove(-1, true, undefined, undefined, bIsWord);
			}
			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 9) {// Tab
			if (this.CanEdit()) {
				if (oController) {
					var graphicObjects = oController;
					var target_content = graphicObjects.getTargetDocContent(undefined, true);
					if (target_content) {
						if (target_content instanceof CTable) {
							target_content.MoveCursorToCell(e.ShiftKey ? false : true);
						} else {
							if (true === this.CollaborativeEditing.Is_Fast() || this.Api.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
								if (target_content.Selection.StartPos === target_content.Selection.EndPos &&
									target_content.Content[target_content.CurPos.ContentPos].IsCursorAtBegin() &&
									target_content.Content[target_content.CurPos.ContentPos].CompiledPr.Pr &&
									target_content.Content[target_content.CurPos.ContentPos].CompiledPr.Pr.ParaPr.Bullet &&
									target_content.Content[target_content.CurPos.ContentPos].CompiledPr.Pr.ParaPr.Bullet.isBullet() &&
									target_content.Content[target_content.CurPos.ContentPos].CompiledPr.Pr.ParaPr.Bullet.bulletType.type !== AscFormat.BULLET_TYPE_BULLET_NONE) {
									if (this.Can_IncreaseParagraphLevel(!e.ShiftKey)) {
										this.IncreaseDecreaseIndent(!e.ShiftKey);
									}
								} else {
									History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
									this.AddToParagraph(new AscWord.CRunTab());
								}


							}
						}
					} else {
						graphicObjects.selectNextObject(!e.ShiftKey ? 1 : -1);
					}
					this.Document_UpdateInterfaceState();
				}
			}
			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 13) // Enter
		{
			var Hyperlink = this.IsCursorInHyperlink(false);
			if (null != Hyperlink && false === e.ShiftKey) {
				this.Api.sync_HyperlinkClickCallback(Hyperlink.GetValue());
				Hyperlink.SetVisited(true);

				// TODO: Пока сделаем так, потом надо будет переделать
				this.DrawingDocument.ClearCachePages();
				this.DrawingDocument.FirePaint();
			} else {
				if (e.CtrlKey) {
					if (oController) {
						var bChangeSelect = false;
						if (!this.IsFocusOnNotes()) {
							var aDrawings = oController.getDrawingArray();
							for (var i = aDrawings.length - 1; i > -1; --i) {
								if (aDrawings[i].selected) {
									break;
								}
							}
							++i;
							for (; i < aDrawings.length; ++i) {
								if (aDrawings[i].getObjectType() === AscDFH.historyitem_type_Shape && aDrawings[i].isPlaceholder()) {
									var oContent = aDrawings[i].getDocContent();
									if (oContent) {
										oContent.Set_CurrentElement(0, false);
										bChangeSelect = true;
										if (!oContent.IsEmpty()) {
											oContent.SelectAll();
										}
										this.DrawingDocument.OnRecalculateSlide(this.CurPage);
										this.Document_UpdateSelectionState();
										break;
									}
								}
							}
						}
						if (this.CanEdit()) {
							if (!bChangeSelect) {
								History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
								this.addNextSlide();
							}
						}
					}
				} else {
					if (oController) {
						var oTargetDocContent = oController.getTargetDocContent();
						if (oTargetDocContent) {
							if (e.ShiftKey) {
								if (oController.selectedObjects.length !== 0) {
									if (true === this.CollaborativeEditing.Is_Fast() || this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
										History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);

										var oSelectedInfo = new CSelectedElementsInfo();
										oTargetDocContent.GetSelectedElementsInfo(oSelectedInfo);
										var oMath = oSelectedInfo.GetMath();
										if (null !== oMath && oMath.Is_InInnerContent()) {
											if (oMath.Handle_AddNewLine()) {

												let oShape = oTargetDocContent.Is_DrawingShape(true);
												if(oShape)
													oShape.checkExtentsByDocContent();
												this.Recalculate();
											}
										} else {
											this.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Line));
										}
									}
								}
							} else {
								if (oController.selectedObjects.length !== 0) {
									var aSelectedObjects = oController.selectedObjects;
									if (aSelectedObjects.length === 1 && aSelectedObjects[0].isPlaceholder && aSelectedObjects[0].isPlaceholder()
										&& aSelectedObjects[0].getPlaceholderType && (aSelectedObjects[0].getPlaceholderType() === AscFormat.phType_ctrTitle || aSelectedObjects[0].getPlaceholderType() === AscFormat.phType_title)) {
										if (true === this.CollaborativeEditing.Is_Fast() || this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
											History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
											var oSelectedInfo = new CSelectedElementsInfo();
											oTargetDocContent.GetSelectedElementsInfo(oSelectedInfo);
											var oMath = oSelectedInfo.GetMath();
											if (null !== oMath && oMath.Is_InInnerContent()) {
												if (oMath.Handle_AddNewLine()) {
													let oShape = oTargetDocContent.Is_DrawingShape(true);
													if(oShape)
														oShape.checkExtentsByDocContent();
													this.Recalculate();
												}
											} else {
												this.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Line));
											}
										}
									} else {

										var oSelectedInfo = new CSelectedElementsInfo();
										oTargetDocContent.GetSelectedElementsInfo(oSelectedInfo);
										var oMath = oSelectedInfo.GetMath();
										if (null !== oMath && oMath.Is_InInnerContent()) {
											if (true === this.CollaborativeEditing.Is_Fast() || this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
												History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
												if (oMath.Handle_AddNewLine()) {
													let oShape = oTargetDocContent.Is_DrawingShape(true);
													if(oShape)
														oShape.checkExtentsByDocContent();
													this.Recalculate();
												}
											}
										} else {
											this.AddNewParagraph();

										}
									}


								}
							}
						} else {
							var nRet = oController.handleEnter();
							if (nRet & 2) {
								if (this.GetCurrentSlide()) {
									this.DrawingDocument.OnRecalculateSlide(this.CurPage);
								}
							}
							if (nRet & 1) {
								this.Document_UpdateSelectionState();
								this.Document_UpdateInterfaceState();
								this.Document_UpdateRulersState();
							}
						}
					}
				}
			}

			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 27) // Esc
		{
			const bCancelEyedropper = this.CancelEyedropper();
			const bCancelInkDrawer = this.CancelInkDrawer();
			if (oController && !this.IsFocusOnNotes()) {
				if(!bCancelEyedropper && !bCancelInkDrawer) {
					var oDrawingObjects = oController;
					if (oDrawingObjects.isTrackingDrawings()) {
						this.Api.sync_EndAddShape();
						oDrawingObjects.endTrackNewShape();
						this.UpdateCursorType(0, 0, new AscCommon.CMouseEventHandler());
						this.UpdateInterface();
						return;
					}
					var oTargetTextObject = AscFormat.getTargetTextObject(oDrawingObjects);

					var bNeedRedraw;
					if (oTargetTextObject && oTargetTextObject.isEmptyPlaceholder()) {
						bNeedRedraw = true;
					} else {
						bNeedRedraw = false;
					}
					var bChart = oDrawingObjects.checkChartTextSelection(true);
					if (!bNeedRedraw) {
						bNeedRedraw = bChart;
					}
					if (!bNeedRedraw) {
						var oCurContent = oDrawingObjects.getTargetDocContent(false, false);
						if (oCurContent) {
							var oCurParagraph = oCurContent.GetCurrentParagraph();
							if (oCurParagraph && oCurParagraph.IsEmpty()) {
								bNeedRedraw = true;
							}
						}
					}
					if (e.ShiftKey || (!oDrawingObjects.selection.groupSelection && !oDrawingObjects.selection.textSelection && !oDrawingObjects.selection.chartSelection)) {
						oDrawingObjects.resetSelection();
					} else {
						if (oDrawingObjects.selection.groupSelection) {
							var oGroupSelection = oDrawingObjects.selection.groupSelection.selection;
							if (oGroupSelection.textSelection) {
								if (oGroupSelection.textSelection.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
									if (oGroupSelection.textSelection.graphicObject) {
										oGroupSelection.textSelection.graphicObject.RemoveSelection();
									}
								} else {
									var content = oGroupSelection.textSelection.getDocContent();
									content && content.RemoveSelection();
								}
								oGroupSelection.textSelection = null;
							} else if (oGroupSelection.chartSelection) {
								oGroupSelection.chartSelection.resetSelection(false);
								oGroupSelection.chartSelection = null;
							} else {
								oDrawingObjects.selection.groupSelection.resetSelection(oDrawingObjects);
								oDrawingObjects.selection.groupSelection = null;
							}
						} else if (oDrawingObjects.selection.textSelection) {
							if (oDrawingObjects.selection.textSelection.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
								if (oDrawingObjects.selection.textSelection.graphicObject) {
									oDrawingObjects.selection.textSelection.graphicObject.RemoveSelection();
								}
							} else {
								var content = oDrawingObjects.selection.textSelection.getDocContent();
								content && content.RemoveSelection();
							}
							oDrawingObjects.selection.textSelection = null;
						} else if (oDrawingObjects.selection.chartSelection) {
							oDrawingObjects.selection.chartSelection.resetSelection(false);
							oDrawingObjects.selection.chartSelection = null;
						}
					}
					if (bNeedRedraw) {
						this.DrawingDocument.OnRecalculateSlide(this.CurPage);
					}
					this.Document_UpdateSelectionState();
					this.Document_UpdateInterfaceState();
				}
			}
			if (true === this.DrawingDocument.IsTrackText()) {
				this.DrawingDocument.CancelTrackText();
			}
			if(bCancelEyedropper || bCancelInkDrawer) {
				this.OnMouseMove(global_mouseEvent, 0, 0, this.CurPage);
			} else if (this.Api.isFormatPainterOn()) {
				this.Api.sync_PaintFormatCallback(AscCommon.c_oAscFormatPainterState.kOff);
				this.OnMouseMove(global_mouseEvent, 0, 0, this.CurPage);
			} else if (this.Api.isMarkerFormat) {
				this.Api.sync_MarkerFormatCallback(false);
				this.OnMouseMove(global_mouseEvent, 0, 0, this.CurPage);
			} else if (this.Api.isInkDrawerOn()) {
				this.Api.stopInkDrawer();
			}
			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 32) // Space
		{
			var oController = this.GetCurrentController();
			if (e.ShiftKey && e.CtrlKey) {
				this.DrawingDocument.TargetStart();
				this.DrawingDocument.TargetShow();


				if (true === this.CollaborativeEditing.Is_Fast() || this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
					if (oController && oController.selectedObjects.length !== 0) {
						History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
						this.AddToParagraph(new AscWord.CRunText(0x00A0));
					}
				}
			} else if (e.CtrlKey) {
				this.ClearParagraphFormatting(false, true);
			} else {
				if (true === this.CollaborativeEditing.Is_Fast() || this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
					if (oController && oController.selectedObjects.length !== 0) {
						History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
						this.CheckLanguageOnTextAdd = true;
						this.AddToParagraph(new AscWord.CRunSpace());
						this.CheckLanguageOnTextAdd = false;
					}
				}
			}
		} else if (e.KeyCode === 33) // PgUp
		{
			if (e.AltKey) {
			} else {
				if (this.CurPage > 0) {
					this.DrawingDocument.m_oWordControl.GoToPage(this.CurPage - 1);
				}
			}
			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 34) // PgDn
		{
			if (e.AltKey) {
			} else {
				if (this.CurPage + 1 < this.Slides.length) {
					this.DrawingDocument.m_oWordControl.GoToPage(this.CurPage + 1);
				}
			}
			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 35) // клавиша End
		{
			if (oController.getTargetDocContent()) {
				if (e.CtrlKey) // Ctrl + End - переход в конец документа
				{
					this.MoveCursorToEndPos();
				} else // Переходим в конец строки
				{
					this.MoveCursorToEndOfLine(e.ShiftKey);
				}
			} else {
				if (!e.ShiftKey) {
					if (this.CurPage !== (this.Slides.length - 1)) {
						this.Api.WordControl.GoToPage(this.Slides.length - 1);
					}
				} else {
					if (this.Slides.length > 0) {
						this.Api.WordControl.Thumbnails && this.Api.WordControl.Thumbnails.CorrectShiftSelect(false, true);
					}
				}
			}

			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 36) // клавиша Home
		{
			if (oController.getTargetDocContent()) {
				if (e.CtrlKey) // Ctrl + Home - переход в начало документа
				{
					this.MoveCursorToStartPos();
				} else // Переходим в начало строки
				{
					this.MoveCursorToStartOfLine(e.ShiftKey);
				}
			} else {
				if (!e.ShiftKey) {
					if (this.Slides.length > 0) {
						this.Api.WordControl.GoToPage(0);
					}
				} else {
					if (this.Slides.length > 0) {
						this.Api.WordControl.Thumbnails && this.Api.WordControl.Thumbnails.CorrectShiftSelect(true, true);
					}
				}
			}
			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 37) // Left Arrow
		{
			if (this.Slides.length > 1 && !this.IsFocusOnNotes() && !e.CtrlKey && this.DrawingDocument.SlideCurrent > 0) {
				if (this.GetCurrentSlideController().selectedObjects.length === 0)
					this.DrawingDocument.m_oWordControl.GoToPage(this.DrawingDocument.SlideCurrent - 1);
			}
			const oController = this.GetCurrentController();
			if (oController)
			{
				const oTargetTextObject = AscFormat.getTargetTextObject(oController);
				if (!oTargetTextObject)
				{
					this.MoveCursorLeft(e.ShiftKey, e.CtrlKey);
					return;
				}
			}
			if (bIsMacOs && e.CtrlKey)
			{
				this.MoveCursorToStartOfLine(e.ShiftKey);
			}
			else
			{
				const bIsWord = bIsMacOs ? e.AltKey : e.CtrlKey;
				this.MoveCursorLeft(e.ShiftKey, bIsWord);
			}
			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 38) // Top Arrow
		{
			if (this.Slides.length > 1 && !this.IsFocusOnNotes() && !e.CtrlKey && this.DrawingDocument.SlideCurrent > 0) {
				if (this.GetCurrentSlideController().selectedObjects.length === 0)
					this.DrawingDocument.m_oWordControl.GoToPage(this.DrawingDocument.SlideCurrent - 1);
			}
			this.MoveCursorUp(e.ShiftKey, e.CtrlKey);
			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 39) // Right Arrow
		{
			// Чтобы при зажатой клавише курсор не пропадал
			// if ( true != e.ShiftKey )
			//     this.DrawingDocument.TargetStart();

			if (this.Slides.length > 1 && !this.IsFocusOnNotes() && !e.CtrlKey && this.DrawingDocument.SlideCurrent < (this.Slides.length - 1)) {
				if (this.GetCurrentSlideController().selectedObjects.length === 0)
					this.DrawingDocument.m_oWordControl.GoToPage(this.DrawingDocument.SlideCurrent + 1);
			}
			const oController = this.GetCurrentController();
			if (oController)
			{
				const oTargetTextObject = AscFormat.getTargetTextObject(oController);
				if (!oTargetTextObject)
				{
					this.MoveCursorRight(e.ShiftKey, e.CtrlKey);
					return;
				}
			}
			if (bIsMacOs && e.CtrlKey)
			{
				this.MoveCursorToEndOfLine(e.ShiftKey);
			}
			else
			{
				const bIsWord = bIsMacOs ? e.AltKey : e.CtrlKey;
				this.MoveCursorRight(e.ShiftKey, bIsWord);
			}
			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 40) // Bottom Arrow
		{
			// Чтобы при зажатой клавише курсор не пропадал
			//if ( true != e.ShiftKey )
			//    this.DrawingDocument.TargetStart();

			if (this.Slides.length > 1 && !this.IsFocusOnNotes() && !e.CtrlKey && this.DrawingDocument.SlideCurrent < (this.Slides.length - 1)) {
				if (this.GetCurrentSlideController().selectedObjects.length === 0)
					this.DrawingDocument.m_oWordControl.GoToPage(this.DrawingDocument.SlideCurrent + 1);
			}
			this.MoveCursorDown(e.ShiftKey, e.CtrlKey);
			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 46) // Delete
		{
			if (true != e.ShiftKey) {
				if (this.CanEdit()) {
					const bIsWord = bIsMacOs ? e.AltKey : e.CtrlKey;
					//this.Create_NewHistoryPoint();
					this.Remove(1, true, undefined, undefined, bIsWord);
				}
				bRetValue = keydownresult_PreventAll;
			}
		} else if (e.KeyCode === 77 && e.CtrlKey) // Ctrl + M + ...
		{
			if (this.CanEdit()) {
				if (oController && oController.getTargetDocContent()) {
					if (e.ShiftKey)
					{// Ctrl + Shift + M - уменьшаем левый отступ
						if (this.Can_IncreaseParagraphLevel(false))
						{
							this.Api.DecreaseIndent();
						}
					}
					else
					{ // Ctrl + M - увеличиваем левый отступ
						if (this.Can_IncreaseParagraphLevel(true))
						{
							this.Api.IncreaseIndent();
						}
					}
				} else {
					if (this.Api.WordControl.Thumbnails) {
						var _selected_thumbnails = this.GetSelectedSlides();
						if (_selected_thumbnails.length > 0) {
							var _last_selected_slide_num = _selected_thumbnails[_selected_thumbnails.length - 1];
							this.Api.WordControl.GoToPage(_last_selected_slide_num);
							this.Api.WordControl.m_oLogicDocument.addNextSlide();
						} else if (this.Slides.length === 0) {
							this.Api.WordControl.m_oLogicDocument.addNextSlide();
							this.Api.WordControl.GoToPage(0);
						}
					}
				}
			}
			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 144) // Num Lock
		{
			// Ничего не делаем
			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 145) // Scroll Lock
		{
			// Ничего не делаем
			bRetValue = keydownresult_PreventAll;
		} else if (e.KeyCode === 187) // +
		{
			if (e.IsCtrl() && e.IsShift()) {
				// Ничего не делаем
				bRetValue = keydownresult_PreventAll;
			}
		}
	}

	if (bRetValue & keydownflags_PreventKeyPress && true === bUpdateSelection)
		this.Document_UpdateSelectionState();

	if(nStartHistoryIndex === this.History.Index) {
		this.private_UpdateCursorXY(true, true);
	}
	oController = this.GetCurrentController();
	if (oController) {
		oController.checkRedrawAnimLabels(aStartAnims);
	}

	this.Api.sendEvent("asc_onKeyDown", e);
	return bRetValue;
};

CPresentation.prototype.Set_DocumentDefaultTab = function (DTab) {
	var oController = this.GetCurrentController();
	return oController && oController.setDefaultTabSize(DTab);
};

CPresentation.prototype.SetDocumentMargin = function () {

};
CPresentation.prototype.EnterText = function (value) {
	if (undefined === value
		|| null === value
		|| (Array.isArray(value) && !value.length))
		return false;
	
	let codePoints = typeof(value) === "string" ? value.codePointsArray() : value;
	
	if (!this.CanEdit())
		return false;

	let oCurSlide = this.GetCurrentSlide();
	if (!oCurSlide || !oCurSlide.graphicObjects) {
		return false;
	}
	if (this.StopAnimationPreview()) {
		return false;
	}
	if (!this.IsFocusOnNotes() && oCurSlide.graphicObjects.selectedObjects.length === 0) {
		let oTitle = oCurSlide.getMatchingShape(AscFormat.phType_title, null);
		if (oTitle) {
			let oDocContent = oTitle.getDocContent && oTitle.getDocContent();
			if (oDocContent.Is_Empty()) {
				oDocContent.Set_CurrentElement(0, false);
			} else {
				return false;
			}
		} else {
			return false;
		}
	}
	if (this.IsFocusOnNotes() && !oCurSlide.notesShape) {
		return false;
	}
	let bRetValue = false;
	let oDocContent1, oDocContent2, bUpdateInterface = false;
	let nCode;
	if (this.CollaborativeEditing.Is_Fast() || !this.Document_Is_SelectionLocked(changestype_Drawing_Props)) {
		this.Create_NewHistoryPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
		let oController = this.GetCurrentController();
		if (oController) {
			oDocContent1 = oController.getTargetDocContent();
		}
		this.CheckLanguageOnTextAdd = true;
		let oItem;
		if (Array.isArray(codePoints)) {
			for (let nIdx = 0; nIdx < codePoints.length; ++nIdx) {
				nCode = codePoints[nIdx];
				oItem = AscCommon.IsSpace(nCode) ? new AscWord.CRunSpace(nCode) : new AscWord.CRunText(nCode);
				this.AddToParagraph(oItem, false, true);
			}
		} else {
			oItem = AscCommon.IsSpace(codePoints) ? new AscWord.CRunSpace(codePoints) : new AscWord.CRunText(codePoints);
			this.AddToParagraph(oItem, false, true);
		}
		this.CheckLanguageOnTextAdd = false;
		if (oController) {
			oDocContent2 = oController.getTargetDocContent();
		}
		if (!oDocContent1 && oDocContent2) {
			bUpdateInterface = true;
			this.Document_UpdateInterfaceState();
			this.Document_UpdateRulersState();
		}
		bRetValue = true;
	}
	if (bRetValue) {
		this.Document_UpdateSelectionState();
		if (!bUpdateInterface) {
			this.Document_UpdateUndoRedoState();
			this.Document_UpdateRulersState();
		}
	}
	return bRetValue;
};

CPresentation.prototype.GetCurrentParagraph = function(bIgnoreSelection, bReturnSelectedArray, oPr) {

	let oController = this.GetCurrentController();
	if (!oController) {
		return null;
	}
	let oContent = oController.getTargetDocContent();
	if(!oContent) {
		return null;
	}
	return oContent.GetCurrentParagraph(bIgnoreSelection, bReturnSelectedArray, oPr);
};

CPresentation.prototype.CorrectEnterText = function (oldValue, newValue) {
	if (undefined === oldValue
		|| null === oldValue
		|| (Array.isArray(oldValue) && !oldValue.length))
		return this.EnterText(newValue);


	let oController = this.GetCurrentController();
	if (!oController) {
		return false;
	}
	let oDocContent = oController.getTargetDocContent(false, false);
	if (!oDocContent) {
		return false;
	}
	if (oDocContent.IsSelectionUse())
		return false;

	let newCodePoints = typeof (newValue) === "string" ? newValue.codePointsArray() : newValue;
	let oldCodePoints = typeof (oldValue) === "string" ? oldValue.codePointsArray() : oldValue;


	if (!Array.isArray(oldCodePoints))
		oldCodePoints = [oldCodePoints];

	let paragraph = oDocContent.GetCurrentParagraph();
	if (!paragraph)
		return false;

	let contentPos = paragraph.GetContentPosition(false, false);
	let run, inRunPos;
	for (let index = contentPos.length - 1; index >= 0; --index) {
		if (contentPos[index].Class instanceof AscWord.CRun) {
			run = contentPos[index].Class;
			inRunPos = contentPos[index].Position;
			break;
		}
	}

	if (!run)
		return false;

	if (!AscWord.checkAsYouTypeEnterText(run, inRunPos, oldCodePoints[oldCodePoints.length - 1]))
		return false;

	if (undefined === newCodePoints || null === newCodePoints)
		newCodePoints = [];
	else if (!Array.isArray(newCodePoints))
		newCodePoints = [newCodePoints];

	let oldText = "";
	for (let index = 0, count = oldCodePoints.length; index < count; ++index) {
		oldText += String.fromCodePoint(oldCodePoints[index]);
	}


	let state = {};
	oController.Save_DocumentStateBeforeLoadChanges(state);

	let maxShifts = oldCodePoints.length;
	let selectedText;
	while (maxShifts >= 0) {
		this.MoveCursorLeft(true, false);
		selectedText = this.GetSelectedText(true);

		if (!selectedText || selectedText === oldText)
			break;

		maxShifts--;
	}

	if (selectedText !== oldText || this.IsSelectionLocked(AscCommon.changestype_Drawing_Props, null, true)) {

		oController.resetSelection()
		oController.loadDocumentStateAfterLoadChanges(state, this.CurPage);
		return false;
	}

	this.StartAction(AscDFH.historydescription_Document_CorrectEnterText);

	this.DrawingDocument.TargetStart();
	this.DrawingDocument.TargetShow();

	oDocContent.Remove(1, true, false, true);

	for (let index = 0, count = newCodePoints.length; index < count; ++index) {
		let codePoint = newCodePoints[index];
		this.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
	}
	this.UpdateSelection();
	this.FinalizeAction();
	return true;
};
CPresentation.prototype.OnKeyPress = function (e) {

	let Code;
	if (null != e.Which)
		Code = e.Which;
	else if (e.KeyCode)
		Code = e.KeyCode;
	else
		Code = 0;//special char

	if (Code > 0x20) {
		return this.EnterText(Code);
	}

	return false;
};


CPresentation.prototype.CheckEmptyPlaceholderNotes = function () {
	let oCurSlide = this.GetCurrentSlide();
	this.DrawingDocument.CheckGuiControlColors();
	if (oCurSlide && oCurSlide.notesShape) {
		let oContent = oCurSlide.notesShape.getDocContent();
		if (oContent && oContent.Is_Empty()) {
			this.DrawingDocument.Notes_OnRecalculate(this.CurPage, oCurSlide.NotesWidth, oCurSlide.getNotesHeight());
			return true;
		}
	}
	return false;
};

CPresentation.prototype.OnMouseDown = function (e, X, Y, PageIndex) {

	this.CurPage = PageIndex;


	var _old_focus = this.FocusOnNotes;
	this.FocusOnNotes = false;
	if (PageIndex < 0)
		return;


	if (this.StopAnimationPreview()) {
		return;
	}


	// Сбрасываем текущий элемент в поиске
	if (this.SearchEngine.Count > 0)
		this.SearchEngine.ResetCurrent();

	this.CurPage = PageIndex;
	e.ctrlKey = e.CtrlKey;
	e.shiftKey = e.ShiftKey;
	let oSlide = this.GetCurrentSlide();
	let oController = oSlide && oSlide.graphicObjects;
	let oContent1, oContent2;
	var ret = null;
	if (oController) {
		oContent1 = oController.getTargetDocContent();
		var aStartAnims = oController.getAnimSelectionState();
		ret = oController.onMouseDown(e, X, Y);
		oController.checkRedrawAnimLabels(aStartAnims);
		oContent2 = oController.getTargetDocContent();
	}
	let bUpdate = true;
	if(oContent1 && oContent1 === oContent2) {
		bUpdate = false;
	}
	if(bUpdate) {
		this.private_UpdateCursorXY(true, true);
	}
	if (!ret) {
		this.Document_UpdateSelectionState();
	}
	this.Document_UpdateInterfaceState();
	if (_old_focus) {
		this.CheckEmptyPlaceholderNotes();
	}

	if (ret) {
		return keydownresult_PreventAll;
	}
	return keydownresult_PreventNothing;
};

CPresentation.prototype.OnMouseUp = function (e, X, Y, PageIndex) {
	e.ctrlKey = e.CtrlKey;
	e.shiftKey = e.ShiftKey;
	const nStartPage = this.CurPage;


	let oSlide = this.GetCurrentSlide();
	let oController = oSlide && oSlide.graphicObjects;
	if (oController) {
		const aStartAnims = oController.getAnimSelectionState();
		oController.onMouseUp(e, X, Y);
		oController.checkRedrawAnimLabels(aStartAnims);
		this.private_UpdateCursorXY(true, true);
	}
	if (nStartPage !== this.CurPage) {
		this.DrawingDocument.CheckTargetShow();
		this.Document_UpdateSelectionState();
	}
	if (e.Button === AscCommon.g_mouse_button_right && !this.noShowContextMenu) {
		const ContextData = new AscCommonSlide.CContextMenuData();
		const ConvertedPos = this.DrawingDocument.ConvertCoordsToCursorWR(X, Y, PageIndex);
		ContextData.X_abs = ConvertedPos.X;
		ContextData.Y_abs = ConvertedPos.Y;
		ContextData.IsSlideSelect = false;
		ContextData.Guide = this.hitInGuide(X, Y);
		this.Api.sync_ContextMenuCallback(ContextData);
	}

	this.noShowContextMenu = false;
	this.Document_UpdateInterfaceState();
	if (oController.isSlideShow()) {
		oController.handleEventMode = AscFormat.HANDLE_EVENT_MODE_CURSOR;
		const oResult = oController.curState.onMouseDown(e, X, Y, 0);
		oController.handleEventMode = AscFormat.HANDLE_EVENT_MODE_HANDLE;
		if (oResult) {
			return keydownresult_PreventAll;
		}
	}
	return keydownresult_PreventNothing;
};


CPresentation.prototype.IsSlideShow = function() {
	return this.Api.isSlideShow();
};
CPresentation.prototype.OnMouseMove = function (e, X, Y, PageIndex) {

	e.ctrlKey = e.CtrlKey;
	e.shiftKey = e.ShiftKey;
	this.Api.sync_MouseMoveStartCallback();
	this.CurPage = PageIndex;
	let oSlide = this.GetCurrentSlide();
	let oController = oSlide && oSlide.graphicObjects;
	if (oController) {
		oController.onMouseMove(e, X, Y);
	}
	let bOldFocus = this.FocusOnNotes;
	this.FocusOnNotes = false;
	this.UpdateCursorType(X, Y, e);
	this.FocusOnNotes = bOldFocus;
	this.Api.sync_MouseMoveEndCallback();
	if (oController.isSlideShow()) {
		oController.handleEventMode = AscFormat.HANDLE_EVENT_MODE_CURSOR;
		let oResult = oController.curState.onMouseDown(e, X, Y, 0);
		oController.handleEventMode = AscFormat.HANDLE_EVENT_MODE_HANDLE;
		if (oResult) {
			return keydownresult_PreventAll;
		}
	}
	return keydownresult_PreventNothing;
};

CPresentation.prototype.OnEndTextDrag = function (NearPos, bCopy) {
	var oController = this.GetCurrentController();
	if (!oController) {
		return;
	}
	var oContent = oController.getTargetDocContent();
	if (oContent && oContent.CheckPosInSelection(0, 0, 0, NearPos)) {
		var Paragraph = NearPos.Paragraph;
		Paragraph.Cursor_MoveToNearPos(NearPos);
		Paragraph.Document_SetThisElementCurrent(false);

		oController.onMouseUp(AscCommon.global_mouseEvent, AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
		this.private_UpdateCursorXY(true, true);
		this.Document_UpdateSelectionState();
		this.Document_UpdateInterfaceState();
		this.Document_UpdateRulersState();
	} else {

		var oParagraph = NearPos.Paragraph;
		var bUndo = true;
		History.Create_NewPoint(AscDFH.historydescription_Document_DragText);

		var oSelectedContent = this.GetSelectedContent();
		var aCheckObjects = [];
		var aSelectedObjects, oObjectFrom, bIsLocked;
		if (oSelectedContent && oSelectedContent.DocContent) {
			if (oParagraph) {
				if (oParagraph.Parent && oParagraph.Parent.Parent && oParagraph.Parent.Parent.parent) {
					var oObjectTo = oParagraph.Parent.Parent.parent;
					var initialObjectTo = oObjectTo;
					while (oObjectTo.group) {
						oObjectTo = oObjectTo.group;
					}
					oObjectFrom = AscFormat.getTargetTextObject(oController);
					if (oObjectFrom && oObjectFrom.getObjectType() === AscDFH.historyitem_type_Shape) {
						while (oObjectFrom.group) {
							oObjectFrom = oObjectFrom.group;
						}
						aCheckObjects.push(oObjectTo);
						if (!bCopy) {
							if (oObjectFrom !== oObjectTo) {
								aCheckObjects.push(oObjectFrom);
							}
						}
						aSelectedObjects = oController.selectedObjects;
						oController.selectedObjects = [];
						bIsLocked = this.Document_Is_SelectionLocked(changestype_Drawing_Props, aCheckObjects);

						oController.selectedObjects = aSelectedObjects;
						if (!bIsLocked) {

							NearPos.Paragraph.Check_NearestPos(NearPos);
							if (!bCopy) {
								oController.removeCallback(-1, undefined, undefined, undefined, undefined, true);
							}
							oController.resetSelection(false, false);
							oSelectedContent = oSelectedContent.copy();
							oSelectedContent.DocContent.Insert(NearPos, true);
							oController.onMouseUp(AscCommon.global_mouseEvent, AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
							if (initialObjectTo.isObjectInSmartArt()) {
								initialObjectTo.checkExtentsByDocContent();
							}
							this.Recalculate();
							this.Document_UpdateSelectionState();
							this.Document_UpdateUndoRedoState();
							this.Document_UpdateInterfaceState();
							this.Document_UpdateRulersState();
							bUndo = false;
						}
					}
				}
			} else if (oSelectedContent.SlideObjects.length === 0) {
				oObjectFrom = AscFormat.getTargetTextObject(oController);
				if (oObjectFrom && oObjectFrom.getObjectType() === AscDFH.historyitem_type_Shape) {
					while (oObjectFrom.group) {
						oObjectFrom = oObjectFrom.group;
					}
					aCheckObjects.push(oObjectFrom);
					aSelectedObjects = oController.selectedObjects;
					oController.selectedObjects = [];
					bIsLocked = this.Document_Is_SelectionLocked(changestype_Drawing_Props, aCheckObjects);

					oController.selectedObjects = aSelectedObjects;

					if (!bIsLocked) {
						if (!bCopy) {
							var bNoCheck = oObjectFrom.getObjectType() !== AscDFH.historyitem_type_SmartArt;
							oController.removeCallback(-1, undefined, undefined, undefined, undefined, bNoCheck);
						}
						let oCurSlide = this.GetCurrentSlide();
						oCurSlide.graphicObjects.resetSelection(undefined, false);
						oSelectedContent = oSelectedContent.copy();
						this.InsertContent(oSelectedContent);
						var oShape = oCurSlide.graphicObjects.selectedObjects[0];
						if (oShape) {
							oShape.spPr.xfrm.setOffX(NearPos.X);
							oShape.spPr.xfrm.setOffY(NearPos.Y);
						}
						this.Recalculate();
						bUndo = false;
						oController.onMouseUp(AscCommon.global_mouseEvent, AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
						this.Document_UpdateSelectionState();
						this.Document_UpdateUndoRedoState();
						this.Document_UpdateInterfaceState();
						this.Document_UpdateRulersState();
					}
				}
			}
		}
		if (bUndo) {
			History.Remove_LastPoint();
			if (oParagraph) {
				oParagraph.Clear_NearestPosArray();
			}
			oController.onMouseUp(AscCommon.global_mouseEvent, AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
		}
	}
};


/**
 * @returns {boolean}
 */
CPresentation.prototype.IsShowShapeAdjustments = function () {
	return (!!this.CanEdit());
};
/**
 * Рисовать ли трек у таблицы и давать ли возможность таскать границы
 * @returns {boolean}
 */
CPresentation.prototype.IsShowTableAdjustments = function () {
	return (!!this.CanEdit());
};
/**
 * Рисовать ли трек у таблицы и давать ли возможность таскать границы
 * @returns {boolean}
 */
CPresentation.prototype.IsShowEquationTrack = function () {
	return (!!this.CanEdit());
};
/**
 * Можем ли перетаскивать текст
 * @returns {boolean}
 */
CPresentation.prototype.CanDragAndDrop = function () {
	return (!!this.CanEdit());
};

CPresentation.prototype.IsFocusOnNotes = function () {

	if(!this.IsMasterMode() && this.FocusOnNotes) {
		let oCurSlide = this.GetCurrentSlide();
		return !!oCurSlide.notes;
	}
	return false;
};

CPresentation.prototype.IsFocusOnThumbnails = function () {
	return this.GetFocusObjType() === FOCUS_OBJECT_THUMBNAILS;
};

CPresentation.prototype.Notes_OnMouseDown = function (e, X, Y) {
	// Сбрасываем текущий элемент в поиске
	this.CancelEyedropper();
	this.CancelInkDrawer();
	if (this.SearchEngine.Count > 0)
		this.SearchEngine.ResetCurrent();
	var bFocusOnSlide = !this.FocusOnNotes;
	this.FocusOnNotes = true;
	var oCurSlide = this.GetCurrentSlide();
	var oDrawingObjects = oCurSlide.graphicObjects;
	if (oDrawingObjects.checkTrackDrawings()) {
		this.Api.sync_EndAddShape();
		this.Api.stopInkDrawer();
		oDrawingObjects.endTrackNewShape();
		this.UpdateCursorType(0, 0, new AscCommon.CMouseEventHandler());
	}
	if (oCurSlide) {
		if (bFocusOnSlide) {
			var bNeedRedraw = false;
			if (AscFormat.checkEmptyPlaceholderContent(oCurSlide.graphicObjects.getTargetDocContent(false, false))) {
				bNeedRedraw = true;
			}
			oCurSlide.graphicObjects.resetSelection(true, false);
			var aComments = oCurSlide.slideComments && oCurSlide.slideComments.comments;
			if (Array.isArray(aComments)) {
				for (var i = 0; i < aComments.length; ++i) {
					if (aComments[i].selected) {
						bNeedRedraw = true;
						aComments[i].selected = false;
						this.Api.asc_hideComments();
						break;
					}
				}
			}
			oCurSlide.graphicObjects.clearPreTrackObjects();
			oCurSlide.graphicObjects.clearTrackObjects();
			oCurSlide.graphicObjects.changeCurrentState(new AscFormat.NullState(oCurSlide.graphicObjects));
			if (bNeedRedraw) {
				this.DrawingDocument.OnRecalculateSlide(this.CurPage);
				this.DrawingDocument.OnEndRecalculate();
			}

		}
		if (oCurSlide.notes) {
			e.ctrlKey = e.CtrlKey;
			e.shiftKey = e.ShiftKey;
			var ret = oCurSlide.notes.graphicObjects.onMouseDown(e, X, Y);
			this.private_UpdateCursorXY(true, true);
			if (bFocusOnSlide) {
				this.CheckEmptyPlaceholderNotes();
			}
			if (!ret) {
				this.Document_UpdateSelectionState();
			}
			this.Document_UpdateInterfaceState();
		}
	}

};

CPresentation.prototype.Notes_OnMouseUp = function (e, X, Y) {
	if (!this.IsFocusOnNotes()) {
		return;
	}
	var oCurSlide = this.GetCurrentSlide();
	if (oCurSlide && oCurSlide.notes) {
		e.ctrlKey = e.CtrlKey;
		e.shiftKey = e.ShiftKey;
		oCurSlide.notes.graphicObjects.onMouseUp(e, X, Y);
		this.private_UpdateCursorXY(true, true);
		if (e.Button === AscCommon.g_mouse_button_right && !this.noShowContextMenu) {
			var ContextData = new AscCommonSlide.CContextMenuData();
			var ConvertedPos = this.DrawingDocument.ConvertCoordsToCursorWR(X, Y, this.CurPage);
			ContextData.X_abs = ConvertedPos.X;
			ContextData.Y_abs = ConvertedPos.Y;
			ContextData.IsSlideSelect = false;
			this.Api.sync_ContextMenuCallback(ContextData);
		}
		this.noShowContextMenu = false;
		this.Document_UpdateInterfaceState();
		this.Api.sendEvent("asc_onSelectionEnd");
	}
};

CPresentation.prototype.Notes_OnMouseMove = function (e, X, Y) {
	// if(!this.FocusOnNotes){
	//     return;
	// }
	var oCurSlide = this.GetCurrentSlide();
	if (oCurSlide) {
		if (oCurSlide.notes) {
			e.ctrlKey = e.CtrlKey;
			e.shiftKey = e.ShiftKey;
			this.Api.sync_MouseMoveStartCallback();
			oCurSlide.notes.graphicObjects.onMouseMove(e, X, Y);
			var bOldFocus = this.FocusOnNotes;
			this.FocusOnNotes = true;
			this.UpdateCursorType(X, Y, e);
			this.FocusOnNotes = bOldFocus;
			this.Api.sync_MouseMoveEndCallback();
		}
	}
};


CPresentation.prototype.Get_TableStyleForPara = function () {
	return null;
};

CPresentation.prototype.GetSelectionAnchorPos = function () {
	let oCurSlide = this.GetCurrentSlide();
	if (oCurSlide) {
		var selected_objects = oCurSlide.graphicObjects.selectedObjects;
		if (selected_objects.length > 0) {
			var last_object = selected_objects[selected_objects.length - 1];
			var Coords1 = this.Api.WordControl.m_oDrawingDocument.ConvertCoordsToCursorWR_Comment(last_object.x, last_object.y, this.CurPage);
			var Coords2 = this.Api.WordControl.m_oDrawingDocument.ConvertCoordsToCursorWR_Comment(last_object.x + last_object.extX, last_object.y, this.CurPage);
			return {X0: Coords1.X, X1: Coords2.X, Y: Coords1.Y};
		} else {

			var Pos = this.Api.WordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
			var Coords1 = this.Api.WordControl.m_oDrawingDocument.ConvertCoordsToCursorWR_Comment(0, 0, this.CurPage);
			return {X0: Coords1.X, X1: Coords1.X, Y: Coords1.Y};
		}
	}
	return {X0: 0, X1: 0, Y: 0};
};

CPresentation.prototype.Clear_ContentChanges = function () {
	this.m_oContentChanges.Clear();
};

CPresentation.prototype.Add_ContentChanges = function (Changes) {
	this.m_oContentChanges.Add(Changes);
};

CPresentation.prototype.Refresh_ContentChanges = function () {
	this.m_oContentChanges.Refresh();
};


CPresentation.prototype.GetFormattingPasteData = function () {
	let oController = this.GetCurrentController();
	if (!oController)
		return null;
	return oController.getFormatPainterData(true);
};
CPresentation.prototype.Document_Format_Copy = function () {
	this.Api.checkFormatPainterData();
};

CPresentation.prototype.Document_Format_Paste = function () {
	let oData = this.Api.getFormatPainterData();
	if(!oData) return;
	let oController = this.GetCurrentController();
	if(!oController) return;
	oController.pasteFormattingWithPoint(oData);
};

// Возвращаем выделенный текст, если в выделении не более 1 параграфа, и там нет картинок, нумерации страниц и т.д.
CPresentation.prototype.GetSelectedText = function (bClearText, oPr) {
	if (undefined === oPr)
		oPr = {};

	if (undefined === bClearText)
		bClearText = false;

	var oController = this.GetCurrentController();
	if (oController) {
		return oController.GetSelectedText(bClearText, oPr);
	}
	return "";
};

CPresentation.prototype.GetSelectedParagraphs = function () {
	var oController = this.GetCurrentController();
	if (!oController) {
		return [];
	}
	var oTargetContent = oController.getTargetDocContent();
	if (!oTargetContent) {
		return [];
	}
	var aParagraphs = [];
	oTargetContent.GetCurrentParagraph(false, aParagraphs, {});
	return aParagraphs;
};
//-----------------------------------------------------------------------------------
// Функции для работы с таблицами
//-----------------------------------------------------------------------------------

CPresentation.prototype.ApplyTableFunction = function (Function, bBefore, bAll, Cols, Rows) {
	let oController = this.GetCurrentController();
	if(!oController)
		return null;

	let result = null;

	let args;
	if (AscFormat.isRealNumber(Rows) && AscFormat.isRealNumber(Cols)) {
		args = [Cols, Rows];
	} else {
		args = [bBefore];
	}
	let oTargetText = AscFormat.getTargetTextObject(oController);
	let oTable;
	if (oTargetText && oTargetText.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
		oTable = oTargetText.graphicObject;
	}
	else {
		let oByTypes = oController.getSelectedObjectsByTypes(true);
		if (oByTypes.tables.length === 1) {
			let oGrFrame = oByTypes.tables[0];
			oTable = oGrFrame.graphicObject;
			if (Function !== AscWord.CTable.prototype.DistributeTableCells) {
				oGrFrame.Set_CurrentElement();
				if (!(bAll === true)) {
					if (bBefore) {
						oTable.MoveCursorToStartPos();
					} else {
						oTable.MoveCursorToStartPos();
					}
				} else {
					oTable.SelectAll();
				}
			}
		}
	}
	if(oTable) {
		result = Function.apply(oTable, args);
		if (oTable.Content.length === 0) {
			this.RemoveTable();
			return result;
		}
		if (!this.IsFocusOnNotes()) {
			oController.updateConnectors(false);
		}
		this.Recalculate();
		this.Document_UpdateInterfaceState();
	}
	return result;
};


CPresentation.prototype.AddTableRow = function (bBefore) {
	this.ApplyTableFunction(CTable.prototype.AddTableRow, bBefore);
};

CPresentation.prototype.AddTableColumn = function (bBefore) {
	this.ApplyTableFunction(CTable.prototype.AddTableColumn, bBefore);
};

CPresentation.prototype.RemoveTableRow = function () {
	this.ApplyTableFunction(CTable.prototype.RemoveTableRow, undefined);
};

CPresentation.prototype.RemoveTableColumn = function () {
	this.ApplyTableFunction(CTable.prototype.RemoveTableColumn, true);
};

CPresentation.prototype.DistributeTableCells = function (isHorizontally) {
	return this.ApplyTableFunction(CTable.prototype.DistributeTableCells, isHorizontally);
};

CPresentation.prototype.MergeTableCells = function () {
	this.ApplyTableFunction(CTable.prototype.MergeTableCells, false, true);
};

CPresentation.prototype.SplitTableCells = function (Cols, Rows) {
	this.ApplyTableFunction(CTable.prototype.SplitTableCells, true, true, parseInt(Cols, 10), parseInt(Rows, 10));
};

CPresentation.prototype.RemoveTable = function () {
	let oCurSlide = this.GetCurrentSlide();
	if (oCurSlide) {
		let oController = oCurSlide.graphicObjects;
		oController.deleteSelectedObjectsCallback();
		this.Recalculate();
		this.Document_UpdateInterfaceState();
		this.Document_UpdateSelectionState();
	}
};

CPresentation.prototype.SelectTable = function (Type) {
	const oController = this.GetCurrentController();
	if (oController) {
		let oByTypes = oController.getSelectedObjectsByTypes(true);
		if (oByTypes.tables.length === 1) {
			let oGrFrame = oByTypes.tables[0];
			oGrFrame.Set_CurrentElement();
			oGrFrame.graphicObject.SelectTable(Type);
			this.Document_UpdateSelectionState();
			this.Document_UpdateInterfaceState();
		}
	}
};

CPresentation.prototype.Table_CheckFunction = function (Function) {
	let oController = this.GetCurrentController();
	if (oController) {
		let oTextDrawing = AscFormat.getTargetTextObject(oController);
		if (oTextDrawing && oTextDrawing.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
			return Function.apply(oTextDrawing.graphicObject, []);
		}
	}
	return false;
};

CPresentation.prototype.CanMergeTableCells = function () {
	return this.Table_CheckFunction(CTable.prototype.CanMergeTableCells);
};

CPresentation.prototype.CanSplitTableCells = function () {
	return this.Table_CheckFunction(CTable.prototype.CanSplitTableCells);
};

CPresentation.prototype.CheckTableCoincidence = function (Table) {
	return false;
};

CPresentation.prototype.Get_PageSizesByDrawingObjects = function () {
	return {W: Page_Width, H: Page_Height};
};

CPresentation.prototype.ChangeTextCase = function (nCaseType) {
	var oController = this.GetCurrentController();
	if (!oController) {
		return;
	}
	oController.changeTextCase(nCaseType);
	this.Document_UpdateInterfaceState();
};

//-----------------------------------------------------------------------------------
// Дополнительные функции
//-----------------------------------------------------------------------------------
CPresentation.prototype.Document_CreateFontMap = function () {
	var nSlide;
	var oCheckedMap = {};
	var oFontsMap = {};
	let nCount = this.GetSlidesCount();
	for (nSlide = 0; nSlide < nCount; ++nSlide) {
		this.GetSlide(nSlide).createFontMap(oFontsMap, oCheckedMap);
	}
	return oFontsMap;
};

CPresentation.prototype.Document_CreateFontCharMap = function (FontCharMap) {
	//TODO !!!!!!!!!!
};

CPresentation.prototype.Document_Get_AllFontNames = function () {
	var AllFonts = {}, i;
	if (this.defaultTextStyle && this.defaultTextStyle.Document_Get_AllFontNames) {
		this.defaultTextStyle.Document_Get_AllFontNames(AllFonts);
	}
	for (i = 0; i < this.Slides.length; ++i) {
		this.Slides[i].getAllFonts(AllFonts)
	}
	for (i = 0; i < this.slideMasters.length; ++i) {
		this.slideMasters[i].getAllFonts(AllFonts);
	}
	if (this.globalTableStyles) {
		this.globalTableStyles.Document_Get_AllFontNames(AllFonts);
	}

	for (i = 0; i < this.notesMasters.length; ++i) {
		this.notesMasters[i].getAllFonts(AllFonts);
	}
	for (i = 0; i < this.notes.length; ++i) {
		this.notes[i].getAllFonts(AllFonts);
	}
	delete AllFonts["+mj-lt"];
	delete AllFonts["+mn-lt"];
	delete AllFonts["+mj-ea"];
	delete AllFonts["+mn-ea"];
	delete AllFonts["+mj-cs"];
	delete AllFonts["+mn-cs"];
	return AllFonts;
};


CPresentation.prototype.Get_AllImageUrls = function (aImages) {
	if (!Array.isArray(aImages)) {
		aImages = [];
	}
	let oObjectsToCheck = {};
	for (let i = 0; i < this.Slides.length; ++i) {
		let oSlide = this.Slides[i];
		oSlide.getAllRasterImages(aImages);
	}
	for (let i = 0; i < this.slideMasters.length; ++i) {
		let oMaster = this.slideMasters[i];
		oObjectsToCheck[oMaster.Id] = oMaster; 
		for (let j = 0; j < oMaster.sldLayoutLst.length; ++j) {
			let oLayout = oMaster.sldLayoutLst[j];
			oObjectsToCheck[oLayout.Id] = oLayout;
		}
		let oTheme = oMaster.Theme;
		if (oTheme) {
			oObjectsToCheck[oTheme.Id] = oTheme;
		}
	}
	for (let sKey in oObjectsToCheck) {
		if (oObjectsToCheck.hasOwnProperty(sKey)) {
			oObjectsToCheck[sKey].getAllRasterImages(aImages);
		}
	}
	return aImages;
};

CPresentation.prototype.Reassign_ImageUrls = function (images_rename) {
	let oObjectsToCheck = {};
	for (let i = 0; i < this.Slides.length; ++i) {
		let oSlide = this.Slides[i];
		oSlide.Reassign_ImageUrls(images_rename);
	}
	for (let i = 0; i < this.slideMasters.length; ++i) {
		let oMaster = this.slideMasters[i];
		oObjectsToCheck[oMaster.Id] = oMaster; 
		for (let j = 0; j < oMaster.sldLayoutLst.length; ++j) {
			let oLayout = oMaster.sldLayoutLst[j];
			oObjectsToCheck[oLayout.Id] = oLayout;
		}
		let oTheme = oMaster.Theme;
		if (oTheme) {
			oObjectsToCheck[oTheme.Id] = oTheme;
		}
	}
	for (let sKey in oObjectsToCheck) {
		if (oObjectsToCheck.hasOwnProperty(sKey)) {
			oObjectsToCheck[sKey].Reassign_ImageUrls(images_rename);
		}
	}
};


CPresentation.prototype.Get_GraphicObjectsProps = function () {
	let oController = this.GetCurrentController();
	if (oController) {
		return oController.getDrawingProps();
	}
	return null;
};

CPresentation.prototype.TurnOff_InterfaceEvents = function () {
	this.TurnOffInterfaceEvents = true;
};

CPresentation.prototype.TurnOn_InterfaceEvents = function (bUpdate) {
	this.TurnOffInterfaceEvents = false;

	if (true === bUpdate) {
		this.Document_UpdateInterfaceState();
		this.Document_UpdateSelectionState();
		this.Document_UpdateRulersState();
	}
};

CPresentation.prototype.Document_UpdateInterfaceState = function () {
	if (this.TurnOffInterfaceEvents) {
		return;
	}
	if (!this.Api) {
		return;
	}
	this.Api.sync_BeginCatchSelectedElements();
	this.Api.ClearPropObjCallback();
	let oCurSlide = this.GetCurrentSlide();
	if (oCurSlide) {
		this.Api.sync_slidePropCallback(oCurSlide);
		let oController = this.GetCurrentController();
		if (!oController) {
			this.Api.sync_EndCatchSelectedElements();
			return;
		}
		let oTargetDocContent = oController.getTargetDocContent(undefined, true);
		let oDrawingPr = oController.getDrawingProps();
		let oParaPr = oController.getParagraphParaPr();
		let oTextPr = oController.getParagraphTextPr();
		this.Api.textArtPreviewManager.clear();
		let oTheme = oController.getTheme();
		if (oTextPr) {
			oTextPr.ReplaceThemeFonts(oTheme.themeElements.fontScheme);
		}
		this.Api.sync_PrLineSpacingCallBack(oParaPr ? oParaPr.Spacing : undefined);
		if (!oTargetDocContent) {
			if (oTextPr && oParaPr) {
				this.Api.UpdateParagraphProp(oParaPr);
				this.Api.UpdateTextPr(oTextPr);
			}

		}
		let oImgPr = oDrawingPr.imageProps;
		let oSpPr = oDrawingPr.shapeProps;
		let oChartPr = oDrawingPr.chartProps;
		let oTblPr = oDrawingPr.tableProps;
		let oHyperlinkPr = oDrawingPr.hyperlinkProps;
		let bIsFocusOnSlide = !this.FocusOnNotes;
		if (bIsFocusOnSlide) {
			if (oImgPr) {
				oImgPr.Width = oImgPr.w;
				oImgPr.Height = oImgPr.h;
				oImgPr.Position = {X: oImgPr.x, Y: oImgPr.y};
				if (AscFormat.isRealBool(oImgPr.locked) && oImgPr.locked) {
					oImgPr.Locked = true;
				}
				this.Api.sync_ImgPropCallback(oImgPr);
			}
			if (oSpPr) {
				oSpPr.Position = new Asc.CPosition({X: oSpPr.x, Y: oSpPr.y});
				this.Api.sync_shapePropCallback(oSpPr);
				this.Api.sync_VerticalTextAlign(oSpPr.verticalTextAlign);
				this.Api.sync_Vert(oSpPr.vert);
			}
			if (oDrawingPr.animProps) {
				this.Api.sync_animPropCallback(oDrawingPr.animProps);
			}
			if (oChartPr && oChartPr.chartProps) {
				if (this.bNeedUpdateChartPreview) {
					this.Api.chartPreviewManager.clearPreviews();
					this.Api.sendEvent("asc_onUpdateChartStyles");
					this.bNeedUpdateChartPreview = false;
				}
				if (oSpPr) {
					oChartPr.x = oSpPr.x;
					oChartPr.y = oSpPr.y;
					if (oSpPr.Position) {
						oChartPr.Position = new Asc.CPosition(oSpPr.Position);
					}
				}
				this.Api.sync_ImgPropCallback(oChartPr);
			}
			if (oTblPr) {
				this.DrawingDocument.CheckTableStyles(oTblPr.TableLook);
				this.Api.sync_TblPropCallback(oTblPr);
				if (!oSpPr) {
					if (oTblPr.CellsVAlign === vertalignjc_Bottom) {
						this.Api.sync_VerticalTextAlign(AscFormat.VERTICAL_ANCHOR_TYPE_BOTTOM);
					} else if (oTblPr.CellsVAlign === vertalignjc_Center) {
						this.Api.sync_VerticalTextAlign(AscFormat.VERTICAL_ANCHOR_TYPE_CENTER);
					} else {
						this.Api.sync_VerticalTextAlign(AscFormat.VERTICAL_ANCHOR_TYPE_TOP);
					}
				}
			}

			if(oHyperlinkPr) {
				this.Api.sync_HyperlinkPropCallback(oHyperlinkPr);
			}
		}
		if (window['IS_NATIVE_EDITOR']) {
			if (!oTblPr) {
				this.DrawingDocument.CheckTableStylesDefault();
			}
		}
		if (oTargetDocContent) {
			oTargetDocContent.Document_UpdateInterfaceState();
		} else {
			if (oTextPr) {
				let nLang = oTextPr && oTextPr.Lang.Val ? oTextPr.Lang.Val : this.GetDefaultLanguage();
				this.Api.sendEvent("asc_onTextLanguage", nLang);
			} else {
				this.Api.sendEvent("asc_onTextLanguage", this.GetDefaultLanguage());
			}
		}
	}
	this.Api.sync_EndCatchSelectedElements();

	this.Document_UpdateUndoRedoState();
	this.Document_UpdateRulersState();
	this.Document_UpdateCanAddHyperlinkState();

	this.Api.sendEvent("asc_onPresentationSize", this.GetWidthEMU(), this.GetHeightEMU(), this.GetSizeType(), this.getFirstSlideNumber());
	this.Api.sendEvent("asc_canIncreaseIndent", this.Can_IncreaseParagraphLevel(true));
	this.Api.sendEvent("asc_canDecreaseIndent", this.Can_IncreaseParagraphLevel(false));
	this.Api.sendEvent("asc_onCanGroup", this.canGroup());
	this.Api.sendEvent("asc_onCanUnGroup", this.canUnGroup());
	this.Api.sendEvent("asc_onCanCopyCut", this.Can_CopyCut());

	if(this.IsMasterMode()) {
		let bTitle = false;
		let bFooter = false;
		let bCanDeleteLayout = false;
		let oCurSlide = this.GetCurrentSlide();
		if(oCurSlide && oCurSlide.isLayout()) {
			let oSp = oCurSlide.getMatchingShape(AscFormat.phType_title, null, false, {});
			if(!oSp) {
				oSp = oCurSlide.getMatchingShape(AscFormat.phType_ctrTitle, null, false, {});
			}
			if(oSp) {
				bTitle = true;
			}
			bFooter = true;
			let aTypes = [AscFormat.phType_ftr, AscFormat.phType_dt, AscFormat.phType_sldNum];
			for(let nIdx = 0; nIdx < aTypes.length; ++nIdx) {
				let nType = aTypes[nIdx];
				let oSp = oCurSlide.getMatchingShape(nType, null, false, {});
				if(!oSp) {
					bFooter = false;
					break;
				}
			}


			bCanDeleteLayout = true;
			let aSelectedSlides = this.GetSelectedSlides();
			for(let nIdx = 0; nIdx < aSelectedSlides.length; ++nIdx) {
				let oSlide = this.GetSlide(aSelectedSlides[nIdx]);
				if(oSlide.isLayout()) {
					if(!this.CanRemoveLayout(oSlide)) {
						bCanDeleteLayout = false;
						break;
					}
				}
				else {
					bCanDeleteLayout = false;
					break;
				}
			}
		}
		this.Api.sendEvent("asc_onLayoutTitle", bTitle);
		this.Api.sendEvent("asc_onLayoutFooter", bFooter);
		this.Api.sendEvent("asc_onCanDeleteMaster", this.CanRemoveMaster(this.lastMaster));
		this.Api.sendEvent("asc_onCanDeleteLayout", bCanDeleteLayout);
	}

	this.Api.externalChartCollector.onUpdateExternalList();

	AscCommon.g_specialPasteHelper.SpecialPasteButton_Update_Position();

	Asc.editor.WordControl.m_oAnimPaneApi && Asc.editor.WordControl.m_oAnimPaneApi.UpdateState()
};
CPresentation.prototype.CanDeleteSelectedMasters = function() {
	return this.CanRemoveMaster(this.lastMaster);
};
CPresentation.prototype.CanDeleteSelectedLayouts = function() {
	let bCanDeleteLayout = true;
	let aSelectedSlides = this.GetSelectedSlides();
	for(let nIdx = 0; nIdx < aSelectedSlides.length; ++nIdx) {
		let oSlide = this.GetSlide(aSelectedSlides[nIdx]);
		if(oSlide.isLayout()) {
			if(!this.CanRemoveLayout(oSlide)) {
				bCanDeleteLayout = false;
				break;
			}
		}
		else {
			bCanDeleteLayout = false;
			break;
		}
	}
	return bCanDeleteLayout;
};
CPresentation.prototype.changeBackground = function (bg, arrSlides, bNoCreatePoint) {
	if (bNoCreatePoint === true || this.Document_Is_SelectionLocked(AscCommon.changestype_SlideBg, arrSlides) === false) {
		if (!(bNoCreatePoint === true)) {
			History.Create_NewPoint(AscDFH.historydescription_Presentation_ChangeBackground);
		}
		for (var i = 0; i < arrSlides.length; ++i) {
			arrSlides[i].changeBackground(bg);
		}

		this.Recalculate();

		for (var i = 0; i < arrSlides.length; ++i) {
			const nIdx = this.GetSlideIndex(arrSlides[i]);
			this.DrawingDocument.OnRecalculateSlide(nIdx);
		}

		this.DrawingDocument.OnEndRecalculate(true, false);
		if (!(bNoCreatePoint === true)) {
			this.Document_UpdateInterfaceState();
		}
	}
};
CPresentation.prototype.GetTableForPreview = function () {
	return AscFormat.ExecuteNoHistory(function () {
		let _x_mar = 10;
		let _y_mar = 10;
		let _r_mar = 10;
		let _b_mar = 10;
		let _pageW = 297;
		let _pageH = 210;
		let W = (_pageW - _x_mar - _r_mar);
		let H = (_pageH - _y_mar - _b_mar);
		let oGrFrame = this.Create_TableGraphicFrame(5, 5, this.GetCurrentSlide(), this.DefaultTableStyleId, W, H, _x_mar, _y_mar, true);
		oGrFrame.setBDeleted(true);
		return oGrFrame.graphicObject;
	}, this, []);
};

CPresentation.prototype.CheckTableForPreview = function (oTable) {
	if (!oTable) {
		return;
	}
	var oGrFrame = oTable.Parent;
	if (!oGrFrame) {
		return;
	}
	var oSlide = this.GetCurrentSlide();
	if (oSlide) {
		oGrFrame.parent = oSlide;
	}
};

CPresentation.prototype.CheckNeedUpdateTableStyles = function (oTableLook) {
	if (!oTableLook) {
		return false;
	}
	let oSlide = this.GetCurrentSlide();
	if (!oSlide) {
		return false;
	}
	let oMaster = this.GetCurrentMaster();
	if (!oMaster) {
		return false;
	}
	var oColorMap = oSlide.Get_ColorMap();
	if (!oColorMap) {
		return false;
	}
	let oDrawingDocument = this.DrawingDocument;
	if (!oDrawingDocument.TableStylesLastTheme ||
		oDrawingDocument.TableStylesLastTheme !== oMaster.Theme ||
		oDrawingDocument.TableStylesLastColorScheme !== oMaster.Theme.themeElements.clrScheme ||
		!oDrawingDocument.TableStylesLastColorMap ||
		!oDrawingDocument.TableStylesLastColorMap.compare(oColorMap) ||
		!oDrawingDocument.TableStylesLastLook ||
		!oDrawingDocument.TableStylesLastLook.IsEqual(oTableLook)) {
		oDrawingDocument.TableStylesLastTheme = oMaster.Theme;
		oDrawingDocument.TableStylesLastColorScheme = oMaster.Theme.themeElements.clrScheme;
		oDrawingDocument.TableStylesLastColorMap = oColorMap;
		oDrawingDocument.TableStylesLastLook = oTableLook.Copy();
		return true;
	}
	return false;
};

// Обновляем линейки
CPresentation.prototype.Document_UpdateRulersState = function () {
	if (this.TurnOffInterfaceEvents) {
		return;
	}
	let oController = this.GetCurrentController();
	if (oController) {
		var target_content = oController.getTargetDocContent(undefined, true);
		if (target_content && target_content.Parent && target_content.Parent.getObjectType && target_content.Parent.getObjectType() === AscDFH.historyitem_type_TextBody) {
			return this.DrawingDocument.Set_RulerState_Paragraph(null, target_content.Parent.getMargins());
		} else if (target_content instanceof CTable) {
			return target_content.Document_UpdateRulersState(this.CurPage);
		}
	}
	this.DrawingDocument.Set_RulerState_Paragraph(null);
};

// Обновляем линейки
CPresentation.prototype.Document_UpdateSelectionState = function () {
	if (this.TurnOffInterfaceEvents) {
		return;
	}
	let oController = this.GetCurrentController();
	if (oController) {
		oController.updateSelectionState();
	}
};
CPresentation.prototype.Document_UpdateUndoRedoState = function () {
	if (true === this.TurnOffInterfaceEvents)
		return;

	if (true === AscCommon.CollaborativeEditing.Get_GlobalLockSelection())
		return;

	// TODO: Возможно стоит перенсти эту проверку в класс CHistory и присылать
	//       данные события при изменении значения History.Index

	// Проверяем состояние Undo/Redo

	var bCanUndo = this.History.Can_Undo();
	if (true !== bCanUndo && this.Api && this.CollaborativeEditing && true === this.CollaborativeEditing.Is_Fast() && true !== this.CollaborativeEditing.Is_SingleUser())
		bCanUndo = this.CollaborativeEditing.CanUndo();

	this.Api.sync_CanUndoCallback(bCanUndo);
	this.Api.sync_CanRedoCallback(this.History.Can_Redo());
	this.Api.CheckChangedDocument();

};

CPresentation.prototype.Document_UpdateCanAddHyperlinkState = function () {
	this.Api.sync_CanAddHyperlinkCallback(this.CanAddHyperlink(false));
};

CPresentation.prototype.Set_CurPage = function (PageNum) {
	if (-1 == PageNum) {
		this.CurPage = -1;
		this.Document_UpdateInterfaceState();
		return false;
	}

	let nSlidesCount = this.GetSlidesCount();
	let nNewCurrentPage = Math.min(nSlidesCount, Math.max(0, PageNum));
	if (nNewCurrentPage !== this.CurPage && nNewCurrentPage < nSlidesCount) {
		var oCurrentController = this.GetCurrentController();
		if (oCurrentController) {
			oCurrentController.resetSelectionState();
		}
		this.CurPage = nNewCurrentPage;
		let oCurSlide = this.GetCurrentSlide();
		this.FocusOnNotes = false;
		this.Notes_OnResize();
		this.DrawingDocument.Notes_OnRecalculate(this.CurPage, oCurSlide.NotesWidth, oCurSlide.getNotesHeight());
		this.Api.asc_hideComments();
		this.Document_UpdateInterfaceState();
		if (this.DrawingDocument.placeholders)
			this.DrawingDocument.placeholders.update(oCurSlide.getPlaceholdersControls());
		this.MathTrackHandler.Update();
		return true;
	}
	let oCurSlide = this.GetCurrentSlide();
	if (oCurSlide && oCurSlide.Layout && oCurSlide.Layout.Master) {
		this.lastMaster = oCurSlide.Layout.Master;
	}
	return false;
};

CPresentation.prototype.Get_CurPage = function () {
	return this.CurPage;
};

CPresentation.prototype.private_UpdateCursorXY = function (bUpdateX, bUpdateY) {
	let oController = this.GetCurrentController();
	if(oController) {
		let oContent = oController.getTargetDocContent();
		if(oContent) {
			if (true === oContent.Selection.Use && true !== oContent.Selection.Start)
				this.Api.sendEvent("asc_onSelectionEnd");
			else if (!oContent.Selection.Use)
				this.Api.sendEvent("asc_onCursorMove");
			this.private_CheckCursorInField();
			return;
		}
	}
	this.Api.sendEvent("asc_onSelectionEnd");
};

CPresentation.prototype.private_CheckCursorInField = function () {
	var oPresentationField = this.GetPresentationField();
	if (oPresentationField) {
		oPresentationField.SelectThisElement();
	}
};

CPresentation.prototype.GetPresentationField = function () {
	var oController = this.GetCurrentController();
	var oDocContent;
	if (oController) {
		oDocContent = oController.getTargetDocContent();
		if (oDocContent) {
			return oDocContent.GetPresentationField();
		}
	}
	return null;
};

CPresentation.prototype.resetStateCurSlide = function () {
	var oCurSlide = this.GetCurrentSlide();
	if (oCurSlide) {
		var bNeedRedraw = false;
		if (AscFormat.checkEmptyPlaceholderContent(oCurSlide.graphicObjects.getTargetDocContent(false, false))) {
			bNeedRedraw = true;
		}
		oCurSlide.graphicObjects.resetSelection(true, false);
		oCurSlide.graphicObjects.clearPreTrackObjects();
		oCurSlide.graphicObjects.clearTrackObjects();
		oCurSlide.graphicObjects.changeCurrentState(new AscFormat.NullState(oCurSlide.graphicObjects));
		if (bNeedRedraw) {
			this.DrawingDocument.OnRecalculateSlide(this.CurPage);
			this.DrawingDocument.OnEndRecalculate();
		}
		oCurSlide.graphicObjects.updateSelectionState();
	}

};

///NOTES
CPresentation.prototype.Notes_OnResize = function () {
	var oCurSlide = this.GetCurrentSlide();
	if (!oCurSlide) {
		return false;
	}
	var newNotesWidth = this.DrawingDocument.Notes_GetWidth();
	if (AscFormat.fApproxEqual(oCurSlide.NotesWidth, newNotesWidth)) {
		return false;
	}
	oCurSlide.NotesWidth = newNotesWidth;
	oCurSlide.recalculateNotesShape();
	this.DrawingDocument.Notes_OnRecalculate(this.CurPage, newNotesWidth, oCurSlide.getNotesHeight());
	return true;
};


CPresentation.prototype.Notes_GetHeight = function () {
	let oCurSlide = this.GetCurrentSlide();
	if (!oCurSlide) {
		return 0;
	}
	return oCurSlide.getNotesHeight();
};

CPresentation.prototype.Notes_Draw = function (SlideIndex, graphics) {
	if (this.Slides[SlideIndex]) {
		if (!graphics.isBoundsChecker()) {
			AscCommon.CollaborativeEditing.Update_ForeignCursorsPositions();
		}
		this.Slides[SlideIndex].drawNotes(graphics);
	}
};

//-----------------------------------------------------------------------------------
// Undo/Redo функции
//-----------------------------------------------------------------------------------
CPresentation.prototype.Create_NewHistoryPoint = function (Description) {
	this.History.Create_NewPoint(Description);
};
CPresentation.prototype.IsFastMultipleUsers = function () {
	return !!(this.Api && this.CollaborativeEditing && true === this.CollaborativeEditing.Is_Fast() && true !== this.CollaborativeEditing.Is_SingleUser());
};
CPresentation.prototype.CanAddChangesToHistory = function () {
	return this.History.CanAddChanges();
};
CPresentation.prototype.IsEditingInFastMultipleUsers = function () {
	return this.IsFastMultipleUsers() && this.CanAddChangesToHistory();
};
CPresentation.prototype.Document_Undo = function (Options) {

	if (true === AscCommon.CollaborativeEditing.Get_GlobalLock())
		return;

	if (true !== this.History.Can_Undo() && this.IsFastMultipleUsers()) {
		if (this.CollaborativeEditing.CanUndo() && true === this.Api.canSave) {
			this.CollaborativeEditing.Set_GlobalLock(true);
			this.Api.forceSaveUndoRequest = true;
		}
	} else {
		this.Api.sendEvent("asc_onBeforeUndoRedo");
		this.clearThemeTimeouts();
		var arrChanges = this.History.Undo(Options);
		this.Recalculate(this.History.Get_RecalcData(null, arrChanges));

		this.Document_UpdateSelectionState();
		this.Document_UpdateInterfaceState();
		this.Api.sendEvent("asc_onUndoRedo");
	}
};

CPresentation.prototype.Document_Redo = function () {
	if (true === AscCommon.CollaborativeEditing.Get_GlobalLock())
		return;

	this.Api.sendEvent("asc_onBeforeUndoRedo");
	this.clearThemeTimeouts();
	var arrChanges = this.History.Redo();
	this.Recalculate(this.History.Get_RecalcData(null, arrChanges));


	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
	this.Api.sendEvent("asc_onUndoRedo");
};

CPresentation.prototype.Set_FastCollaborativeEditing = function (isOn) {
	AscCommon.CollaborativeEditing.Set_Fast(isOn);
};

CPresentation.prototype.GetSelectionState = function () {
	const oSelectionState = {};
	oSelectionState.mode = Asc.editor.presentationViewMode;
	oSelectionState.CurPage = this.CurPage;
	oSelectionState.FocusOnNotes = this.FocusOnNotes;
	oSelectionState.SelectedSlides = this.GetSelectedSlides();
	let oController = this.GetCurrentController();
	if (oController) {
		oSelectionState.slideSelection = oController.getSelectionState();
	}
	oSelectionState.HistoryIndex = this.History.Index;
	return oSelectionState;
};


CPresentation.prototype.SetSelectionState = function (State) {
	if (State.CurPage > -1) {
		if(State.mode !== null && State.mode !== undefined && State.mode !== Asc.editor.presentationViewMode) {
			Asc.editor.asc_changePresentationViewMode(State.mode);
		}
		var oSlide = this.GetSlide(State.CurPage);
		if (oSlide) {
			if (State.FocusOnNotes) {
				oSlide.graphicObjects.resetSelection();
				oSlide.graphicObjects.clearPreTrackObjects();
				oSlide.graphicObjects.clearTrackObjects();
				oSlide.graphicObjects.changeCurrentState(new AscFormat.NullState(oSlide.graphicObjects));
				if (oSlide.notes) {
					this.FocusOnNotes = true;
					if (State.slideSelection) {
						oSlide.notes.graphicObjects.setSelectionState(State.slideSelection);
					}
				} else {
					this.FocusOnNotes = false;
				}
			} else {
				if (State.slideSelection) {
					oSlide.graphicObjects.setSelectionState(State.slideSelection);
				}
			}
		}
	}
	if (State.CurPage !== this.CurPage)
		this.bGoToPage = true;
	this.CurPage = State.CurPage;
};


CPresentation.prototype.Get_SelectionState2 = function () {
	var oState = this.Save_DocumentStateBeforeLoadChanges();
	var oCurController = this.GetCurrentController();
	if (oCurController) {
		var oContent = oCurController.getTargetDocContent(false, false);
		if (oContent) {
			oState.Content2 = oContent;
			oState.SelectionState2 = oContent.Get_SelectionState2();
		}
	}
	return oState;
};

CPresentation.prototype.Set_SelectionState2 = function (oDocState) {
	this.Load_DocumentStateAfterLoadChanges(oDocState);
	var oCurController = this.GetCurrentController();
	if (oCurController) {
		var oContent = oCurController.getTargetDocContent(false, false);
		if (oContent) {
			if (oContent === oDocState.Content2 && oDocState.SelectionState2) {
				oContent.Set_SelectionState2(oDocState.SelectionState2);
			}
		}
	}
};

CPresentation.prototype.Save_DocumentStateBeforeLoadChanges = function () {
	var oDocState = {};
	oDocState.Pos = [];
	oDocState.StartPos = [];
	oDocState.EndPos = [];
	oDocState.CurPage = this.CurPage;
	oDocState.FocusOnNotes = this.FocusOnNotes;
	var oController = this.GetCurrentController();
	if (oController) {
		oDocState.Slide = this.GetCurrentSlide();
		oController.Save_DocumentStateBeforeLoadChanges(oDocState);
	}

	this.CollaborativeEditing.WatchDocumentPositionsByState(oDocState);
	return oDocState;
};

CPresentation.prototype.Load_DocumentStateAfterLoadChanges = function (oState) {

	this.CollaborativeEditing.UpdateDocumentPositionsByState(oState);
	let oCurSlide = this.GetCurrentSlide();
	if(!oCurSlide) {
		this.DrawingDocument.m_oWordControl.GoToPage(this.GetSlidesCount() - 1);
		return;
	}
	if (oState.Slide) {
		var oSlide = oState.Slide;
		if (oSlide !== oCurSlide) {
			var bFind = false;
			let nCount = this.GetSlidesCount();
			for (var i = 0; i < nCount; ++i) {
				let oSlideObj = this.GetSlide(i);
				oSlideObj.setSlideNum && oSlideObj.setSlideNum(i);
				if (oSlideObj === oSlide) {
					this.CurPage = i;
					this.bGoToPage = true;
					bFind = true;
					if (oCurSlide) {

						let oDrawingObjects = oCurSlide.graphicObjects;
						oDrawingObjects.clearPreTrackObjects();
						oDrawingObjects.clearTrackObjects();
						oDrawingObjects.resetSelection(undefined, true, true);
						oDrawingObjects.changeCurrentState(new AscFormat.NullState(oDrawingObjects));
					}
				}
			}
			if (!bFind) {
				if (this.CurPage >= this.GetSlidesCount()) {
					this.CurPage = this.GetSlidesCount() - 1;
				}
				this.bGoToPage = true;
				if (oCurSlide) {
					let oDrawingObjects = oCurSlide.graphicObjects;
					oDrawingObjects.clearPreTrackObjects();
					oDrawingObjects.clearTrackObjects();
					oDrawingObjects.resetSelection(undefined, true, true);
					oDrawingObjects.changeCurrentState(new AscFormat.NullState(oDrawingObjects));
				}
				return;
			}
		}
		let oDrawingObjects = oCurSlide.graphicObjects;
		oDrawingObjects.clearPreTrackObjects();
		oDrawingObjects.clearTrackObjects();
		oDrawingObjects.resetSelection(undefined, true, true);
		oDrawingObjects.changeCurrentState(new AscFormat.NullState(oDrawingObjects));
		if (oState.FocusOnNotes) {
			if (oCurSlide.notes) {
				this.FocusOnNotes = true;
				oCurSlide.notes.graphicObjects.loadDocumentStateAfterLoadChanges(oState);
			}
		} else {
			this.FocusOnNotes = false;
			oDrawingObjects.loadDocumentStateAfterLoadChanges(oState);
		}

	} else {
		if (oState.CurPage === -1) {
			if (this.Slides.length > 0) {
				this.CurPage = 0;
				this.bGoToPage = true;
			}
		}
		if (oCurSlide) {
			var oDrawingObjects = oCurSlide.graphicObjects;
			this.FocusOnNotes = false;
			oDrawingObjects.clearPreTrackObjects();
			oDrawingObjects.clearTrackObjects();
			oDrawingObjects.resetSelection(undefined, true, true);
			oDrawingObjects.changeCurrentState(new AscFormat.NullState(oDrawingObjects));
		}
	}
};

CPresentation.prototype.GetSelectedContent = function () {
	return AscFormat.ExecuteNoHistory(function () {
		var oIdMap, curImgUrl;
		var ret = new PresentationSelectedContent(), i;
		ret.PresentationWidth = this.GetWidthMM();
		ret.PresentationHeight = this.GetHeightMM();
		let oController = this.GetCurrentSlideController();
		let oSlide = this.GetCurrentSlide();
		if (this.Slides.length > 0) {
			var FocusObjectType = this.GetFocusObjType();
			switch (FocusObjectType) {
				case FOCUS_OBJECT_MAIN: {
					var target_text_object = AscFormat.getTargetTextObject(oController);
					if (target_text_object) {
						var doc_content = oController.getTargetDocContent();
						if (target_text_object.getObjectType() === AscDFH.historyitem_type_GraphicFrame && !doc_content) {
							if (target_text_object.graphicObject) {
								var GraphicFrame = target_text_object.copy(undefined);
								var SelectedContent = new AscCommonWord.CSelectedContent();
								target_text_object.graphicObject.GetSelectedContent(SelectedContent);
								var Table = SelectedContent.Elements[0].Element;
								GraphicFrame.setGraphicObject(Table);
								Table.Set_Parent(GraphicFrame);
								curImgUrl = target_text_object.getBase64Img();
								ret.Drawings.push(new DrawingCopyObject(GraphicFrame, target_text_object.x, target_text_object.y, target_text_object.extX, target_text_object.extY, curImgUrl));
							}
						} else {
							if (doc_content) {
								var SelectedContent = new AscCommonWord.CSelectedContent();
								doc_content.GetSelectedContent(SelectedContent);
								ret.DocContent = SelectedContent;
							}
						}
					} else {

						var selector = oController.selection.groupSelection ? oController.selection.groupSelection : oController;
						if (selector.selection.chartSelection && selector.selection.chartSelection.selection.title) {
							var doc_content = selector.selection.chartSelection.selection.title.getDocContent();
							if (doc_content) {

								var SelectedContent = new AscCommonWord.CSelectedContent();
								doc_content.SetApplyToAll(true);
								doc_content.GetSelectedContent(SelectedContent);
								doc_content.SetApplyToAll(false);
								ret.DocContent = SelectedContent;
							}
						} else {
							var bRecursive = isRealObject(oController.selection.groupSelection);
							var aSpTree = bRecursive ? oController.selection.groupSelection.spTree : oSlide.cSld.spTree;
							oIdMap = {};
							collectSelectedObjects(aSpTree, ret.Drawings, bRecursive, oIdMap);
							AscFormat.fResetConnectorsIds(ret.Drawings, oIdMap);
						}
					}
					break;
				}
				case FOCUS_OBJECT_THUMBNAILS : {
					if(!this.IsMasterMode()) {
						var selected_slides = this.GetSelectedSlides();
						for (i = 0; i < selected_slides.length; ++i) {
							oIdMap = {};
							var oSlideCopy = this.Slides[selected_slides[i]].createDuplicate(oIdMap);
							ret.SlideObjects.push(oSlideCopy);
							AscFormat.fResetConnectorsIds(oSlideCopy.cSld.spTree, oIdMap);
						}
					}
				}
			}
		}
		return ret;
	}, this, []);
};

CPresentation.prototype.GetSpeechDescription = function(oBeforeSelectionState, action) {
	if(!oBeforeSelectionState) {
		return null;
	}

	const oEndSelectionState = this.GetSelectionState();

	const nFirstSlideIdx = this.getFirstSlideNumber();
	const correctSlideIndexes = function (aIndexes) {
		for(let nIdx = 0; nIdx < aIndexes.length; ++nIdx) {
			aIndexes[nIdx] += nFirstSlideIdx;
		}
	};
	const getSpeechData = function(type, obj) {
		return {type: type, obj: obj};
	};

	if(oBeforeSelectionState.CurPage !== oEndSelectionState.CurPage) {
		let aIndexes = [this.CurPage];
		correctSlideIndexes(aIndexes);
		return getSpeechData(
			AscCommon.SpeechWorkerCommands.SlidesSelected,
			{
				indexes: aIndexes
			}
		);
	}

	const aStartSelectedSlides = oBeforeSelectionState.SelectedSlides;
	const aEndSelectedSlides = oEndSelectionState.SelectedSlides;

	if(aStartSelectedSlides.length < aEndSelectedSlides.length) {
		let aIndexes = AscCommon.getArrayElementsDiff(aStartSelectedSlides, aEndSelectedSlides);
		if(aIndexes.length > 0) {
			correctSlideIndexes(aIndexes);
			return getSpeechData(
				AscCommon.SpeechWorkerCommands.SlidesSelected,
				{
					indexes: aIndexes
				}
			);
		}
	}
	else if(aStartSelectedSlides.length > aEndSelectedSlides.length) {
		let aIndexes = AscCommon.getArrayElementsDiff(aEndSelectedSlides, aStartSelectedSlides);
		if(aIndexes.length > 0) {
			correctSlideIndexes(aIndexes);
			return getSpeechData(
				AscCommon.SpeechWorkerCommands.SlidesUnselected,
				{
					indexes: aIndexes
				}
			);
		}
	}
	return AscCommon.getSpeechDescription(oBeforeSelectionState.slideSelection, oEndSelectionState.slideSelection, action);
};


CPresentation.prototype.internalResetElementsFontSize = function (aContent) {
	for (var j = 0; j < aContent.length; ++j) {
		if (aContent[j].Type === para_Run) {
			if (aContent[j].Pr && AscFormat.isRealNumber(aContent[j].Pr.FontSize)) {
				var oPr = aContent[j].Pr.Copy();
				oPr.FontSize = undefined;
				aContent[j].Set_Pr(oPr);
			}
		} else if (aContent[j].Type === para_Hyperlink) {
			this.internalResetElementsFontSize(aContent[j].Content);
		}
	}
};
/**Returns array of PresentationSelectedContent for special paste
 * @returns {Array}
 **/
CPresentation.prototype.GetSelectedContent2 = function () {
	return AscFormat.ExecuteNoHistory(function () {
		var aRet = [], oIdMap;
		var oSourceFormattingContent = new PresentationSelectedContent();
		var oEndFormattingContent = new PresentationSelectedContent();
		var oImagesSelectedContent = new PresentationSelectedContent();

		oSourceFormattingContent.PresentationWidth = this.GetWidthMM();
		oSourceFormattingContent.PresentationHeight = this.GetHeightMM();
		oEndFormattingContent.PresentationWidth = this.GetWidthMM();
		oEndFormattingContent.PresentationHeight = this.GetHeightMM();
		oImagesSelectedContent.PresentationWidth = this.GetWidthMM();
		oImagesSelectedContent.PresentationHeight = this.GetHeightMM();
		var oSelectedContent, oDocContent, oController, oTargetTextObject, oGraphicFrame, oTable, oImage, dImageWidth,
			dImageHeight, bNeedSelectAll,
			oDocContentForDraw, oParagraph, aParagraphs, dMaxWidth, oCanvas, oContext, oGraphics,
			dContentHeight, nContentIndents = 30, bOldShowParaMarks, oSelector;
		var i, j;
		if (this.Slides.length > 0) {
			var FocusObjectType = this.GetFocusObjType();
			switch (FocusObjectType) {
				case FOCUS_OBJECT_MAIN: {
					oController = this.GetCurrentController();
					oSelector = oController.selection.groupSelection ? oController.selection.groupSelection : oController;
					oTargetTextObject = AscFormat.getTargetTextObject(oController);
					bNeedSelectAll = false;
					if (!oTargetTextObject) {
						if (oSelector.selection.chartSelection && oSelector.selection.chartSelection.selection.title) {
							oDocContent = oSelector.selection.chartSelection.selection.title.getDocContent();
							if (oDocContent) {
								bNeedSelectAll = true;
							}
						}
					}
					if (oTargetTextObject) {
						if (!oDocContent) {
							oDocContent = oController.getTargetDocContent();
						}
						if (oTargetTextObject.getObjectType() === AscDFH.historyitem_type_GraphicFrame && !oDocContent) {
							if (oTargetTextObject.graphicObject) {
								oGraphicFrame = oTargetTextObject.copy(undefined);
								oSelectedContent = new AscCommonWord.CSelectedContent();
								oTargetTextObject.graphicObject.GetSelectedContent(oSelectedContent);
								oTable = oSelectedContent.Elements[0].Element;
								oGraphicFrame.setGraphicObject(oTable);
								oTable.Set_Parent(oGraphicFrame);
								oEndFormattingContent.Drawings.push(new DrawingCopyObject(oGraphicFrame, oTargetTextObject.x, oTargetTextObject.y, oTargetTextObject.extX, oTargetTextObject.extY, oTargetTextObject.getBase64Img()));
								oGraphicFrame.parent = oTargetTextObject.parent;
								oGraphicFrame.bDeleted = false;
								oGraphicFrame.recalculate();
								oSourceFormattingContent.Drawings.push(new DrawingCopyObject(oGraphicFrame.getCopyWithSourceFormatting(), oTargetTextObject.x, oTargetTextObject.y, oTargetTextObject.extX, oTargetTextObject.extY, oTargetTextObject.getBase64Img()));
								oImage = oController.createImage(oGraphicFrame.getBase64Img(), 0, 0, oGraphicFrame.extX, oGraphicFrame.extY);
								oImagesSelectedContent.Drawings.push(new DrawingCopyObject(oImage, 0, 0, oTargetTextObject.extX, oTargetTextObject.extY, oTargetTextObject.getBase64Img()));
								oGraphicFrame.parent = null;
								oGraphicFrame.bDeleted = true;
							}
						} else {
							if (oDocContent) {
								if (bNeedSelectAll) {
									oDocContent.SetApplyToAll(true);
								}
								oSelectedContent = oDocContent.GetSelectedContent();
								oEndFormattingContent.DocContent = oSelectedContent;
								for (i = 0; i < oSelectedContent.Elements.length; ++i) {
									var oElem = oSelectedContent.Elements[i].Element;
									if (oElem.GetType() === AscCommonWord.type_Paragraph) {
										if (oElem.Pr && oElem.Pr.DefaultRunPr && AscFormat.isRealNumber(oElem.Pr.DefaultRunPr.FontSize)) {
											var oPr = oElem.Pr.Copy();
											oPr.DefaultRunPr.FontSize = undefined;
											oElem.Set_Pr(oPr);
										}
										this.internalResetElementsFontSize(oElem.Content);
									}
								}
								oSelectedContent = oDocContent.GetSelectedContent();
								var aContent = [];
								for (i = 0; i < oSelectedContent.Elements.length; ++i) {
									oParagraph = oSelectedContent.Elements[i].Element;
									oParagraph.Parent = oDocContent;
									oParagraph.private_CompileParaPr();
									aContent.push(oParagraph);
								}
								AscFormat.SaveContentSourceFormatting(aContent, aContent, oDocContent.Get_Theme(), oDocContent.Get_ColorMap());
								oSourceFormattingContent.DocContent = oSelectedContent;

								var oSelectedContent2 = oDocContent.GetSelectedContent();
								aContent = [];
								for (i = 0; i < oSelectedContent2.Elements.length; ++i) {
									oParagraph = oSelectedContent2.Elements[i].Element;
									oParagraph.Parent = oDocContent;
									oParagraph.private_CompileParaPr();
									aContent.push(oParagraph);
								}
								AscFormat.SaveContentSourceFormatting(aContent, aContent, oDocContent.Get_Theme(), oDocContent.Get_ColorMap());

								if (bNeedSelectAll) {
									oDocContent.SetApplyToAll(false);
								}
								if (oSelectedContent2.Elements.length > 0) {
									oDocContentForDraw = new AscFormat.CDrawingDocContent(oDocContent.Parent, oDocContent.DrawingDocument, 0, 0, 20000, 20000);
									oSelectedContent2.ReplaceContent(oDocContentForDraw);

									var oCheckParagraph, aRuns;
									for (i = oDocContentForDraw.Content.length - 1; i > -1; --i) {
										oCheckParagraph = oDocContentForDraw.Content[i];
										if (!oCheckParagraph.IsEmpty()) {
											aRuns = oCheckParagraph.Content;
											if (aRuns.length > 1) {
												for (j = aRuns.length - 2; j > -1; --j) {
													var oRun = aRuns[j];
													if (oRun.Type === para_Run) {
														for (var k = oRun.Content.length - 1; k > -1; --k) {
															if (oRun.Content[k].Type === para_NewLine) {
																oRun.Content.splice(k, 1);
															} else {
																break;
															}
														}
														if (oRun.Content.length === 0) {
															aRuns.splice(j, 1);
														} else {
															break;
														}
													}
												}
											}
										}
										if (oCheckParagraph.IsEmpty()) {
											oDocContentForDraw.Internal_Content_Remove(i, 1, false);
										} else {

											break;
										}
									}
									for (i = 0; i < oDocContentForDraw.Content.length; ++i) {
										oCheckParagraph = oDocContentForDraw.Content[i];
										if (!oCheckParagraph.IsEmpty()) {
											aRuns = oCheckParagraph.Content;
											if (aRuns.length > 1) {
												for (j = 0; j < aRuns.length - 1; ++j) {
													var oRun = aRuns[j];
													if (oRun.Type === para_Run) {
														for (var k = 0; k < oRun.Content.length; ++k) {
															if (oRun.Content[k].Type === para_NewLine) {
																oRun.Content.splice(k, 1);
																k--;
															} else {
																break;
															}
														}
														if (oRun.Content.length === 0) {
															aRuns.splice(j, 1);
															j--;
														} else {
															break;
														}
													}
												}
											}
										}
										if (oCheckParagraph.IsEmpty()) {
											oDocContentForDraw.Internal_Content_Remove(i, 1, false);
											i--;
										} else {

											break;
										}
									}
									if (oDocContentForDraw.Content.length > 0) {

										oDocContentForDraw.Reset(0, 0, 20000, 20000);
										oDocContentForDraw.Recalculate_Page(0, true);
										aParagraphs = oDocContentForDraw.Content;
										dMaxWidth = 0;
										for (i = 0; i < aParagraphs.length; ++i) {
											oParagraph = aParagraphs[i];
											for (j = 0; j < oParagraph.Lines.length; ++j) {
												if (oParagraph.Lines[j].Ranges[0].W > dMaxWidth) {
													dMaxWidth = oParagraph.Lines[j].Ranges[0].W;
												}
											}
										}
										dMaxWidth += 1;


										oDocContentForDraw.Reset(0, 0, dMaxWidth, 20000);
										oDocContentForDraw.Recalculate_Page(0, true);
										dContentHeight = oDocContentForDraw.GetSummaryHeight();

										var oTextWarpObject = null;
										if (oDocContentForDraw.Parent && oDocContentForDraw.Parent.parent && oDocContentForDraw.Parent.parent instanceof AscFormat.CShape) {
											oTextWarpObject = oDocContentForDraw.Parent.parent.checkTextWarp(oDocContentForDraw, oDocContentForDraw.Parent.parent.getBodyPr(), dMaxWidth, dContentHeight, true, false);
										}


										oCanvas = document.createElement('canvas');
										dImageWidth = dMaxWidth;
										dImageHeight = dContentHeight;
										oCanvas.width = ((dImageWidth * AscCommon.g_dKoef_mm_to_pix) + 2 * nContentIndents + 0.5) >> 0;
										oCanvas.height = ((dImageHeight * AscCommon.g_dKoef_mm_to_pix) + 2 * nContentIndents + 0.5) >> 0;
										//if (AscCommon.AscBrowser.isRetina) {
										//    oCanvas.width <<= 1;
										//    oCanvas.height <<= 1;
										//}

										var sImageUrl;
										if (!window["NATIVE_EDITOR_ENJINE"]) {
											oContext = oCanvas.getContext('2d');
											oGraphics = new AscCommon.CGraphics();

											oGraphics.init(oContext, oCanvas.width, oCanvas.height, dImageWidth + 2.0 * nContentIndents / AscCommon.g_dKoef_mm_to_pix, dImageHeight + 2.0 * nContentIndents / AscCommon.g_dKoef_mm_to_pix);
											oGraphics.m_oFontManager = AscCommon.g_fontManager;
											oGraphics.m_oCoordTransform.tx = nContentIndents;
											oGraphics.m_oCoordTransform.ty = nContentIndents;
											oGraphics.transform(1, 0, 0, 1, 0, 0);

											if (oTextWarpObject && oTextWarpObject.oTxWarpStructNoTransform) {
												oTextWarpObject.oTxWarpStructNoTransform.draw(oGraphics, oDocContentForDraw.Parent.parent.Get_Theme(), oDocContentForDraw.Parent.parent.Get_ColorMap());
											} else {

												bOldShowParaMarks = this.Api.ShowParaMarks;
												this.Api.ShowParaMarks = false;
												oDocContentForDraw.Draw(0, oGraphics);
												this.Api.ShowParaMarks = bOldShowParaMarks;
											}
											sImageUrl = oCanvas.toDataURL("image/png");
										} else {
											sImageUrl = "";
										}

										oImage = oController.createImage(sImageUrl, 0, 0, oCanvas.width * AscCommon.g_dKoef_pix_to_mm, oCanvas.height * AscCommon.g_dKoef_pix_to_mm);
										oImagesSelectedContent.Drawings.push(new DrawingCopyObject(oImage, 0, 0, dImageWidth, dImageHeight, sImageUrl));
									}
								}
							}
						}
					} else {
						var bRecursive = isRealObject(oController.selection.groupSelection);
						var aSpTree = bRecursive ? oController.selection.groupSelection.spTree : this.GetCurrentSlide().cSld.spTree;
						oIdMap = {};

						var oTheme = oController.getTheme();
						if (oTheme) {
							oEndFormattingContent.ThemeName = oTheme.name;
							oSourceFormattingContent.ThemeName = oTheme.name;
							oImagesSelectedContent.ThemeName = oTheme.name;
						}
						collectSelectedObjects(aSpTree, oEndFormattingContent.Drawings, bRecursive, oIdMap);
						AscFormat.fResetConnectorsIds(oEndFormattingContent.Drawings, oIdMap);
						oIdMap = {};
						collectSelectedObjects(aSpTree, oSourceFormattingContent.Drawings, bRecursive, oIdMap, true);
						AscFormat.fResetConnectorsIds(oSourceFormattingContent.Drawings, oIdMap);
						let oImageData = oController.getSelectionImageData();
						if (oImageData) {
							let oBounds = oImageData.bounds;
							let sImageUrl = oImageData.src;
							oImage = oController.createImage(sImageUrl, oBounds.min_x * AscCommon.g_dKoef_pix_to_mm, oBounds.min_y * AscCommon.g_dKoef_pix_to_mm, (oImageData.width) * AscCommon.g_dKoef_pix_to_mm, (oImageData.height) * AscCommon.g_dKoef_pix_to_mm);
							oImagesSelectedContent.Drawings.push(new DrawingCopyObject(oImage, 0, 0, (oImageData.width) * AscCommon.g_dKoef_pix_to_mm, (oImageData.height) * AscCommon.g_dKoef_pix_to_mm, sImageUrl));
						}
					}
					break;
				}
				case FOCUS_OBJECT_THUMBNAILS : {

					let oMastersMap = {}, oSlide, oSlideCopy, oLayout, oMaster, oNotesCopy, oNotes;
					function fAddMaster() {
						oSourceFormattingContent.Masters.push(oMaster);
						let aLayouts = oMaster.sldLayoutLst;
						for (let nLayout = 0; nLayout < aLayouts.length; ++nLayout) {
							let oCurLayout = aLayouts[nLayout];
							oSourceFormattingContent.Layouts.push(oCurLayout);
						}
						oTheme = oMaster.Theme;
						oSourceFormattingContent.Themes.push(oTheme);
					}
					if(!this.IsMasterMode()) {
						let aSelectedSlidesIdx = this.GetSelectedSlides();
						for (let nSldIdx = 0; nSldIdx < aSelectedSlidesIdx.length; ++nSldIdx) {
							oIdMap = {};
							oSlide = this.Slides[aSelectedSlidesIdx[nSldIdx]];
							oSlideCopy = oSlide;
							oLayout = oSlide.Layout;
							oMaster = oLayout.Master;
							let sMasterId = oMaster.Get_Id();
							if (!oMastersMap[sMasterId]) {
								oMastersMap[sMasterId] = oMaster;
								fAddMaster();
							}
							oSourceFormattingContent.SlideObjects.push(oSlideCopy);
							let oEndFmtSld = oSlideCopy;
							if (oEndFmtSld.cSld && oEndFmtSld.cSld.Bg) {
								oEndFmtSld = oEndFmtSld.createDuplicate();
								oEndFmtSld.changeBackground(null);
							}
							oEndFormattingContent.SlideObjects.push(oEndFmtSld);

							if (nSldIdx === 0) {
								let sRasterImageId = oSlide.getBase64Img();
								oImage = AscFormat.DrawingObjectsController.prototype.createImage(sRasterImageId, 0, 0, this.GetWidthMM() / 2.0, this.GetHeightMM() / 2.0);
								oImagesSelectedContent.Drawings.push(new DrawingCopyObject(oImage, 0, 0, this.GetWidthMM() / 2.0, this.GetHeightMM() / 2.0, sRasterImageId));
							}
							oNotes = null;
							if (oSlide.notes) {
								oNotes = oSlide.notes;
								oNotesCopy = oNotes;//.createDuplicate();
								oSourceFormattingContent.Notes.push(oNotesCopy);
								oEndFormattingContent.Notes.push(oNotesCopy);
								for (j = 0; j < oSourceFormattingContent.NotesMasters.length; ++j) {
									if (oSourceFormattingContent.NotesMasters[j] === oNotes.Master) {
										oSourceFormattingContent.NotesMastersIndexes.push(j);
										oEndFormattingContent.NotesMastersIndexes.push(j);
										break;
									}
								}
								if (j === oSourceFormattingContent.NotesMasters.length) {
									oSourceFormattingContent.NotesMastersIndexes.push(j);
									oSourceFormattingContent.NotesMasters.push(oNotes.Master);
									oSourceFormattingContent.NotesThemes.push(oNotes.Master.Theme);
									oEndFormattingContent.NotesMastersIndexes.push(j);
									oEndFormattingContent.NotesMasters.push(oNotes.Master);
									oEndFormattingContent.NotesThemes.push(oNotes.Master.Theme);
								}
							} else {
								oSourceFormattingContent.Notes.push(null);
								oSourceFormattingContent.NotesMastersIndexes.push(-1);
								oEndFormattingContent.Notes.push(null);
								oEndFormattingContent.NotesMastersIndexes.push(-1);
							}
						}


					}
					else {
						let aSelectedSlides = this.GetSelectedSlideObjects();
						if(aSelectedSlides.length > 0) {
							let oFirstSlide = aSelectedSlides[0];
							if(oFirstSlide.isMaster()) {
								for(let nIdx = 0; nIdx < aSelectedSlides.length; ++nIdx) {
									let oSld = aSelectedSlides[nIdx];
									if(oSld.isMaster()) {
										oMaster = oSld;
										fAddMaster();
									}
								}
							}
							else {
								for(let nIdx = 0; nIdx < aSelectedSlides.length; ++nIdx) {
									let oSld = aSelectedSlides[nIdx];
									if(oSld.isLayout()) {
										oSourceFormattingContent.Layouts.push(oSld);
									}
								}
							}
						}
					}
					let aSlides = oSourceFormattingContent.SlideObjects;
					let aLayouts = oSourceFormattingContent.Layouts;
					let aMasters = oSourceFormattingContent.Masters;
					let aThemes = oSourceFormattingContent.Themes;
					for (let nSldIdx = 0; nSldIdx < aSlides.length; ++nSldIdx) {
						let oSlide = aSlides[nSldIdx];
						let oLayout = oSlide.Layout;
						for (let nIdx = 0; nIdx < aLayouts.length; ++nIdx) {
							if (aLayouts[nIdx] === oLayout) {
								oSourceFormattingContent.LayoutsIndexes[nSldIdx] = nIdx;
								break;
							}
						}
					}
					for (let nLtIdx = 0; nLtIdx < aLayouts.length; ++nLtIdx) {
						let oCurLayout = aLayouts[nLtIdx];
						for (let nIdx = 0; nIdx < aMasters.length; ++nIdx) {
							if (aMasters[nIdx] === oCurLayout.Master) {
								oSourceFormattingContent.MastersIndexes[nLtIdx] = nIdx;
								break;
							}
						}
					}
					for (let nMasterIdx = 0; nMasterIdx < aMasters.length; ++nMasterIdx) {
						oSourceFormattingContent.ThemesIndexes[nMasterIdx] = nMasterIdx;
					}
					oEndFormattingContent = oSourceFormattingContent;
				}
			}
		}
		aRet.push(oEndFormattingContent);
		aRet.push(oSourceFormattingContent);
		aRet.push(oImagesSelectedContent);
		return aRet;
	}, this, []);
};

CPresentation.prototype.CreateAndAddShapeFromSelectedContent = function (oDocContent) {
	let oCurSlide = this.GetCurrentSlide();
	if(!oCurSlide) return;
	let oParents = oCurSlide.getParentObjects();
	var track_object = new AscFormat.NewShapeTrack("textRect", 0, 0, oParents.master.Theme, oParents.master, oParents.layout, oCurSlide, this.CurPage);
	track_object.track({}, 0, 0);
	var shape = track_object.getShape(false, this.DrawingDocument, oCurSlide);
	shape.setParent(oCurSlide);
	oDocContent.ReplaceContent(shape.txBody.content);
	var body_pr = shape.getBodyPr();
	var w = shape.txBody.getMaxContentWidth(this.GetWidthMM() / 2, true) + body_pr.lIns + body_pr.rIns;
	var h = shape.txBody.content.GetSummaryHeight() + body_pr.tIns + body_pr.bIns;
	shape.spPr.xfrm.setExtX(w);
	shape.spPr.xfrm.setExtY(h);
	shape.spPr.xfrm.setOffX((this.GetWidthMM() - w) / 2);
	shape.spPr.xfrm.setOffY((this.GetHeightMM() - h) / 2);
	shape.setParent(oCurSlide);
	shape.addToDrawingObjects();
	return shape;
};

/** insert content from aContents, aContents[0] - end formatting, aContents[1] - source formatting, aContents[2] - image
 * @param {Array} aContents
 * @param {number} nIndex
 * */
CPresentation.prototype.InsertContent2 = function (aContents, nIndex) {
	//nIndex = 1;
	var oContent, oSlide, i, j, bEndFormatting = (nIndex === 0), oSourceContent, kw = 1.0, kh = 1.0;
	var nLayoutIndex, nMasterIndex, nNotesMasterIndex;
	var oLayout, oMaster, oTheme, oNotes, oNotesMaster, oNotesTheme, oCurrentMaster, bChangeSize = false;
	var bNeedGenerateThumbnails = false;
	let oRet = {};
	if (!aContents[nIndex]) {
		return oRet;
	}
	let oCurSlide = this.GetCurrentSlide();
	if(!oCurSlide) return;
	let oParents = oCurSlide.getParentObjects();
	oCurrentMaster = oParents.master;
	if (!oCurrentMaster) {
		return oRet;
	}
	oContent = aContents[nIndex].copy();
	if (oContent.PresentationWidth !== null && oContent.PresentationHeight !== null) {
		if (!AscFormat.fApproxEqual(this.GetWidthMM(), oContent.PresentationWidth) || !AscFormat.fApproxEqual(this.GetHeightMM(), oContent.PresentationHeight)) {
			bChangeSize = true;
			kw = this.GetWidthMM() / oContent.PresentationWidth;
			kh = this.GetHeightMM() / oContent.PresentationHeight;
		}
	}
	if (oContent.SlideObjects.length > 0 && !this.IsMasterMode()) {

		if (bEndFormatting) {
			oSourceContent = aContents[1];
			for (i = 0; i < oContent.SlideObjects.length; ++i) {
				oSlide = oContent.SlideObjects[i];
				if (bChangeSize) {
					oSlide.Width = oContent.PresentationWidth;
					oSlide.Height = oContent.PresentationHeight;
					oSlide.changeSize(this.GetWidthMM(), this.GetHeightMM());
				}
				nLayoutIndex = oSourceContent.LayoutsIndexes[i];
				oLayout = oSourceContent.Layouts[nLayoutIndex];
				if (oLayout) {
					oSlide.setLayout(oCurrentMaster.getMatchingLayout(oLayout.type, oLayout.matchingName, oLayout.cSld.name, true));
				} else {
					oSlide.setLayout(oCurrentMaster.sldLayoutLst[0]);
				}
				oNotes = oContent.Notes[i];
				if (!oNotes) {
					oNotes = AscCommonSlide.CreateNotes();
				}
				oSlide.setNotes(oNotes);
				oSlide.notes.setNotesMaster(this.notesMasters[0]);
				oSlide.notes.setSlide(oSlide);
			}
		} else {
			bNeedGenerateThumbnails = true;
			for (i = 0; i < oContent.Masters.length; ++i) {
				if (bChangeSize) {
					oContent.Masters[i].scale(kw, kh);
				}
				this.addSlideMaster(this.slideMasters.length, oContent.Masters[i]);
			}
			for (i = 0; i < oContent.Layouts.length; ++i) {
				oLayout = oContent.Layouts[i];
				if (bChangeSize) {
					oLayout.scale(kw, kh);
				}
				nMasterIndex = oContent.MastersIndexes[i];
				oMaster = oContent.Masters[nMasterIndex];
				if (oMaster) {
					oMaster.addLayout(oLayout);
				}

			}
			for (i = 0; i < oContent.SlideObjects.length; ++i) {
				oSlide = oContent.SlideObjects[i];
				if (bChangeSize) {
					oSlide.Width = oContent.PresentationWidth;
					oSlide.Height = oContent.PresentationHeight;
					oSlide.changeSize(this.GetWidthMM(), this.GetHeightMM());
				}
				nLayoutIndex = oContent.LayoutsIndexes[i];
				oLayout = oContent.Layouts[nLayoutIndex];
				oSlide.setLayout(oLayout);

				nLayoutIndex = oContent.LayoutsIndexes[i];
				oLayout = oContent.Layouts[nLayoutIndex];
				nMasterIndex = oContent.MastersIndexes[nLayoutIndex];
				oMaster = oContent.Masters[nMasterIndex];
				oTheme = oContent.Themes[nMasterIndex];
				oNotes = oContent.Notes[i];
				nNotesMasterIndex = oContent.NotesMastersIndexes[i];
				oNotesMaster = oContent.NotesMastersIndexes[nNotesMasterIndex];
				oNotesTheme = oContent.NotesThemes[nNotesMasterIndex];
				if (!oMaster.Theme) {
					oMaster.setTheme(oTheme);
				}
				if (!oLayout.Master) {
					oLayout.setMaster(oMaster);
				}
				if (!oNotes || !oNotesMaster || !oNotesTheme) {
					oSlide.setNotes(AscCommonSlide.CreateNotes());
					oSlide.notes.setNotesMaster(this.notesMasters[0]);
					oSlide.notes.setSlide(oSlide);
				} else {
					if (!oNotesMaster.Themes) {
						oNotesMaster.setTheme(oNotesTheme);
					}
					if (!oNotes.Master) {
						oNotes.setNotes(oNotesMaster);
					}
					if (!oSlide.notes) {
						oSlide.setNotes(oNotes);
					}
				}
			}
		}
	}
	else if(this.IsMasterMode() && (oContent.Layouts.length > 0 || oContent.Masters.length > 0)) {
		for (i = 0; i < oContent.Layouts.length; ++i) {
			oLayout = oContent.Layouts[i];
			if (bChangeSize) {
				oLayout.scale(kw, kh);
			}
			oLayout.setSlideSize(this.GetWidthMM(), this.GetHeightMM());
			nMasterIndex = oContent.MastersIndexes[i];
			oMaster = oContent.Masters[nMasterIndex];
			if (oMaster) {
				oMaster.addLayout(oLayout);
			}
		}

		let oCurSlide = this.GetCurrentSlide();
		let oCurMaster = this.GetCurrentMaster();
		if(oContent.Masters.length > 0) {
			let nPos = this.slideMasters.length;
			for(let nMaster = this.slideMasters.length; nMaster < this.slideMasters.length; ++nMaster) {
				if(this.slideMasters[nMaster] === oCurMaster) {
					nPos = nMaster + 1;
					break;
				}
			}
			for (i = 0; i < oContent.Masters.length; ++i) {
				if (bChangeSize) {
					oContent.Masters[i].scale(kw, kh);
				}
				oContent.Masters[i].setSlideSize(this.GetWidthMM(), this.GetHeightMM());
				oTheme = oContent.Themes[i];
				oContent.Masters[i].setTheme(oTheme);
				this.addSlideMaster(nPos + i, oContent.Masters[i]);
			}

			let nPage = this.GetSlideIndex(oContent.Masters[0]);
			if(nPage > -1) {
				this.CurPage = this.GetSlideIndex(oContent.Masters[0]);
				this.bGoToPage = true;
			}
		}
		else {
			if(oCurMaster) {
				let nPos = oCurMaster.sldLayoutLst.length;
				if(oCurSlide.isMaster()) {
					nPos = 0;
				}
				else {
					for(let nLt = 0; nLt < oCurMaster.sldLayoutLst.length; ++nLt) {
						if(oCurMaster.sldLayoutLst[nLt] === oCurSlide) {
							nPos = nLt + 1;
						}
					}
				}
				for (i = 0; i < oContent.Layouts.length; ++i) {
					oLayout = oContent.Layouts[i];
					oCurMaster.addToSldLayoutLstToPos(nPos + i, oLayout);
				}
				let nPage = this.GetSlideIndex(oContent.Layouts[0]);
				if(nPage > -1) {
					this.CurPage = this.GetSlideIndex(oContent.Layouts[0]);
					this.bGoToPage = true;
				}
			}
		}
	}
	if (oContent.Drawings.length > 0) {
		if (bEndFormatting) {
			oSourceContent = aContents[1];
			if (oCurSlide && !this.IsFocusOnNotes() && oSourceContent) {
				AscFormat.checkDrawingsTransformBeforePaste(oContent, oSourceContent, oCurSlide);
			}
		}
	}
	if (oContent.DocContent && oContent.DocContent.Elements.length > 0 && nIndex === 0) {
		var oTextPr, oTextPr2, oParaTextPr, nFontSize, oTextObject, oElement, aElements;
		var oController = this.GetCurrentController();
		if (oController) {
			var oTargetDocContent = oController.getTargetDocContent();
			if (oTargetDocContent) {
				var oCurParagraph = oTargetDocContent.GetCurrentParagraph();
				if (oCurParagraph) {
					oTextPr = oCurParagraph.Internal_CompiledParaPrPresentation(undefined, true).TextPr;

					var fApplyPropsToContent = function (content, textpr) {
						for (var j = 0; j < content.length; ++j) {
							if (content[j].Get_Type) {
								if (content[j].Get_Type() === para_Run) {
									if (content[j].Pr) {
										var oTextPr2 = textpr.Copy();
										oTextPr2.Merge(content[j].Pr);
										content[j].Set_Pr(oTextPr2);
									}
								} else if (content[j].Get_Type() === para_Hyperlink) {
									fApplyPropsToContent(content[j].Content, textpr);
								}
							}

						}
					};
					for (i = 0; i < oContent.DocContent.Elements.length; ++i) {
						if (oContent.DocContent.Elements[i].Element.GetType() === AscCommonWord.type_Paragraph) {
							var aContent = oContent.DocContent.Elements[i].Element.Content;
							fApplyPropsToContent(aContent, oTextPr);
						}
					}
				}
				oTextPr = oTargetDocContent.GetCalculatedTextPr();
				if (oTextPr && AscFormat.isRealNumber(oTextPr.FontSize)) {
					nFontSize = oTextPr.FontSize;
					if (!AscFormat.isRealNumber(oTextPr.FontScale) ||
						AscFormat.fApproxEqual(oTextPr.FontScale, 1.0)) {
						nFontSize = oTextPr.FontSize;
					} else {
						oTextObject = AscFormat.getTargetTextObject(oController);
						if (oTextObject && oTextObject.getObjectType() === AscDFH.historyitem_type_Shape) {
							oTextObject.bCheckAutoFitFlag = true;
							oTextObject.tmpFontScale = 100000;
							oTextObject.tmpLnSpcReduction = 0;
							oTextObject.recalculateContentWitCompiledPr();
							oTextPr = oTargetDocContent.GetCalculatedTextPr();
							if (AscFormat.isRealNumber(oTextPr.FontSize)) {
								nFontSize = oTextPr.FontSize;
							}
							oTextObject.bCheckAutoFitFlag = false;
							oTextObject.tmpFontScale = undefined;
							oTextObject.recalculateContentWitCompiledPr();
						}
					}
					oTextPr2 = new AscCommonWord.CTextPr();
					oTextPr2.FontSize = nFontSize;
					oParaTextPr = new AscCommonWord.ParaTextPr(oTextPr2);
					aElements = oContent.DocContent.Elements;
					for (i = 0; i < aElements.length; ++i) {
						oElement = aElements[i].Element;
						if (oElement.GetType() === AscCommonWord.type_Paragraph) {
							oElement.SetApplyToAll(true);
							oElement.AddToParagraph(oParaTextPr);
							oElement.SetApplyToAll(false);
						}
					}
				}
			}
		}
	}
	oRet = this.InsertContent(oContent);
	if (bNeedGenerateThumbnails) {
		this.SendThemesThumbnails();
	}
	return oRet;
};

CPresentation.prototype.InsertContent = function (Content) {
	let oResult = {};
	let selected_slides = this.GetSelectedSlides(), i;
	let oThumbnails = this.Api.WordControl.Thumbnails;
	let nNeedFocusType = null;
	if (oThumbnails) {
		nNeedFocusType = oThumbnails.FocusObjType;
	}
	let oCurSlide = this.GetCurrentSlide();
	if (Content.SlideObjects.length > 0) {
		if(!this.IsMasterMode()) {
			let las_slide_index = selected_slides.length > 0 ? selected_slides[selected_slides.length - 1] : -1;
			this.needSelectPages.length = 0;
			for (i = 0; i < Content.SlideObjects.length; ++i) {
				this.insertSlide(las_slide_index + i + 1, Content.SlideObjects[i]);
				this.needSelectPages.push(las_slide_index + i + 1);
			}
			this.CurPage = las_slide_index + 1;
			this.bGoToPage = true;
			this.bNeedUpdateTh = true;
			this.FocusOnNotes = false;
			this.CheckEmptyPlaceholderNotes();
		}
		oResult.insert = true;
		nNeedFocusType = FOCUS_OBJECT_THUMBNAILS;
	} else {
		if (!oCurSlide) {
			this.addNextSlideAction(null);
			this.CurPage = 0;
			oCurSlide = this.Slides[0];
		}
		if (oCurSlide) {
			if (Content.Drawings.length > 0) {
				let oIsSingleTable = null;
				if (Content.Drawings.length === 1 &&
					Content.Drawings[0].Drawing &&
					Content.Drawings[0].Drawing.isTable()) {
					oIsSingleTable = Content.Drawings[0].Drawing.graphicObject;
				}
				if (this.IsFocusOnNotes() && oIsSingleTable) {
					let oContent = AscFormat.ExecuteNoHistory(
						function () {
							let oTable = Content.Drawings[0].Drawing.graphicObject;
							let oResult = new AscFormat.CDrawingDocContent(this, this.DrawingDocument, 0, 0, 3000, 2000);
							for (let i = 0; i < oTable.Content.length; ++i) {
								let oRow = oTable.Content[i];
								for (let j = 0; j < oRow.Content.length; ++j) {
									let oCurDocContent = oRow.Content[j].Content;
									for (let k = 0; k < oCurDocContent.Content.length; ++k) {
										oResult.Content.push(oCurDocContent.Content[k]);
									}
								}
							}
							if (oResult.Content.length > 1) {
								oResult.Content.splice(0, 1);
							}
							return oResult;
						}, this, []
					);
					let oSelectedContent = new AscCommonWord.CSelectedContent();
					oContent.SelectAll();
					oContent.GetSelectedContent(oSelectedContent);
					let PresentSelContent = new PresentationSelectedContent();
					PresentSelContent.DocContent = oSelectedContent;
					this.InsertContent(PresentSelContent);
					this.Check_CursorMoveRight();
					nNeedFocusType = FOCUS_OBJECT_MAIN;
					return true;
				} else {
					this.FocusOnNotes = false;
					let oController = oCurSlide.graphicObjects;
					let oTextSelection = oController.selection.textSelection;
					if (oIsSingleTable && oTextSelection && oTextSelection.isTable()) {
						let oTable = oTextSelection.graphicObject;
						let oCurParagraph = oTable.GetCurrentParagraph();
						if (oCurParagraph) {
							let oParaParent = oCurParagraph.GetParent();
							if (oParaParent) {
								let oCurCell = oParaParent.IsTableCellContent(true);
								if (oCurCell) {
									oCurCell.InsertTableContent(oIsSingleTable);
									let nMaxCellsCount = 0;
									for (let nRow = 0; nRow < oTable.Content.length; nRow++) {
										let oRow = oTable.Content[nRow];
										if (nMaxCellsCount < oRow.Content.length)
											nMaxCellsCount = oRow.Content.length;
									}
									for (let nRow = 0; nRow < oTable.Content.length; nRow++) {
										let oRow = oTable.Content[nRow];
										let aCells = oRow.Content;
										let nCellsCount = 0;
										for (let nCell = 0; nCell < aCells.length; nCell++) {
											nCellsCount += aCells[nCell].GetGridSpan();
										}
										if (nCellsCount < nMaxCellsCount) {
											for (let nCell = nCellsCount; nCell < nMaxCellsCount; nCell++) {
												oRow.Add_Cell(oRow.Get_CellsCount(), oRow, null, true);
											}
										}
									}
									oResult.insert = true;
									nNeedFocusType = FOCUS_OBJECT_MAIN;
								}
							}
						}
					}
					if (!oResult.insert) {
						oController.resetSelection();
						let aPasteDrawings = [];
						for (i = 0; i < Content.Drawings.length; ++i) {
							aPasteDrawings.push(Content.Drawings[i].Drawing);
						}
						let dShift = oController.getDrawingsPasteShift(aPasteDrawings);
						if (Content.Drawings.length === 1)
						{
							const oDrawing = Content.Drawings[0];
							const oSp = oDrawing.Drawing;
							oResult.specialPasteProps = oSp.getSpecialPasteProps && oSp.getSpecialPasteProps();
						}
						for (i = 0; i < Content.Drawings.length; ++i) {
							let bInsertShape = true;
							let oCopyObject = Content.Drawings[i];
							let oSp = oCopyObject.Drawing;
							let oSlidePh, oLayoutPlaceholder;

							let nType, nIdx;
							oSp.generateSmartArtDrawingPart();
							if (oSp.isPlaceholder() && !this.IsMasterMode()) {
								let oInfo = {};
								nType = oSp.getPlaceholderType();
								nIdx = oSp.getPlaceholderIndex();
								oSlidePh = oCurSlide.getMatchingShape(nType, nIdx, false, oInfo);
								oLayoutPlaceholder = oCurSlide.Layout.getMatchingShape(nType, nIdx, false, oInfo);
								if (oSp.isEmptyPlaceholder()) {
									if (!oLayoutPlaceholder || oInfo.bBadMatch) {
										bInsertShape = false;
									} else {
										if (oSlidePh) {
											bInsertShape = false;
										}
									}
								}
							}
							if (bInsertShape) {
								if (oSp.bDeleted) {
									if (oSp.setBDeleted2) {
										oSp.setBDeleted2(false);
									} else if (oSp.setBDeleted) {
										oSp.setBDeleted(false);
									}
								}
								oSp.applySpecialPasteProps();
								oSp.setParent2(this.GetCurrentSlide());
								if (oSp.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
									this.Check_GraphicFrameRowHeight(oSp);
								}
								if (dShift > 0) {
									let oXfrm = oSp.getXfrm();
									if (oXfrm) {
										oXfrm.shift(dShift, dShift);
									}
								}
								oSp.addToDrawingObjects();
								oSp.checkExtentsByDocContent && oSp.checkExtentsByDocContent();
								if (oSp.isPlaceholder() && !this.IsMasterMode()) {
									if (oSlidePh || !oLayoutPlaceholder) {
										let oNvProps = oSp.getNvProps();
										if (oNvProps && oNvProps.ph) {
											if (oSp.txBody) {
												let oLstStyles = new AscFormat.TextListStyle(), oLstStylesTmp,
													oParentObjects;
												oParentObjects = oSp.getParentObjects();
												if (oParentObjects && oParentObjects.master && oParentObjects.master.txStyles) {

													oLstStylesTmp = oParentObjects.master.txStyles.getStyleByPhType(nType);
													if (oLstStylesTmp) {
														oLstStyles.merge(oLstStylesTmp);
													}
												}

												let aHierarhy = oSp.getHierarchy();
												let oBodyPr = new AscFormat.CBodyPr();
												for (let s = aHierarhy.length - 1; s > -1; --s) {
													if (aHierarhy[s]) {
														if (aHierarhy[s].txBody) {
															oLstStyles.merge(aHierarhy[s].txBody.lstStyle);
															oBodyPr.merge(aHierarhy[s].txBody.bodyPr);
														}
													}
												}
												oLstStyles.merge(oSp.txBody.lstStyle);
												oBodyPr.merge(oSp.txBody.bodyPr);
												oSp.txBody.setLstStyle(oLstStyles);
												oSp.txBody.setBodyPr(oBodyPr);
											}
											oNvProps.setPh(null);
										}
									}
								}
								oController.selectObject(oSp, 0);
								oResult.insert = true;
								nNeedFocusType = FOCUS_OBJECT_MAIN;
							}
						}
						if (Content.DocContent && Content.DocContent.Elements.length > 0) {
							let shape = this.CreateAndAddShapeFromSelectedContent(Content.DocContent);
							oController.selectObject(shape, 0);
							oResult.insert = true;
							nNeedFocusType = FOCUS_OBJECT_MAIN;
						}
					}
				}
			} else if (Content.DocContent) {
				Content.DocContent.EndCollect(this);
				if (Content.DocContent.Elements.length > 0) {
					let oController = this.GetCurrentController();
					let target_doc_content = oController.getTargetDocContent(true), paragraph, NearPos;
					if (target_doc_content) {
						if (target_doc_content.Selection.Use) {
							oController.removeCallback(1, undefined, undefined, undefined, undefined, undefined);
						}
						paragraph = target_doc_content.Content[target_doc_content.CurPos.ContentPos];
						if (null != paragraph && paragraph.IsParagraph()) {
							NearPos = {Paragraph: paragraph, ContentPos: paragraph.Get_ParaContentPos(false, false)};
							paragraph.Check_NearestPos(NearPos);
							Content.DocContent.Insert(NearPos);
						}
						let oTargetTextObject = AscFormat.getTargetTextObject(this.GetCurrentController());
						oTargetTextObject && oTargetTextObject.checkExtentsByDocContent && oTargetTextObject.checkExtentsByDocContent();
					} else {
						this.FocusOnNotes = false;
						let shape = this.CreateAndAddShapeFromSelectedContent(Content.DocContent);
						oController.resetSelection();
						oController.selectObject(shape, 0);
						this.CheckEmptyPlaceholderNotes();
					}
					oResult.insert = true;
					nNeedFocusType = FOCUS_OBJECT_MAIN;
				}
			}
		}
	}
	if (oResult.insert && oThumbnails) {
		if (oThumbnails.FocusObjType !== nNeedFocusType) {
			oThumbnails.SetFocusElement(nNeedFocusType);
		}
	}
	return oResult;
};


CPresentation.prototype.Get_NearestPos = function (Page, X, Y, bNotes) {
	var oCurSlide = this.GetCurrentSlide();
	if (!oCurSlide) {
		return;
	}
	var oNearestPos;
	if (bNotes) {
		if (oCurSlide.notesShape) {
			var oContent = oCurSlide.notesShape.getDocContent();
			if (oContent) {
				var tx = oCurSlide.notesShape.invertTransformText.TransformPointX(X, Y);
				var ty = oCurSlide.notesShape.invertTransformText.TransformPointY(X, Y);
				return oContent.Get_NearestPos(0, tx, ty, false);
			}
		}
	} else {
		if (oCurSlide.graphicObjects) {
			oNearestPos = oCurSlide.graphicObjects.getNearestPos3(X, Y);
			if (oNearestPos) {
				return oNearestPos;
			}
			return {
				X: X,
				Y: Y,
				Height: 0,
				PageNum: 0,
				Internal: {Line: 0, Page: 0, Range: 0},
				Transform: null,
				Paragraph: null,
				ContentPos: null,
				SearchPos: null
			};
		}
	}
	return null;
};

CPresentation.prototype.SendThemesThumbnails = function () {
	if (window['IS_NATIVE_EDITOR']) {
		this.DrawingDocument.CheckThemes();
		return;
	}

	if (!window['native']) {
		this.GenerateThumbnails(this.Api.WordControl.m_oMasterDrawer, this.Api.WordControl.m_oLayoutDrawer);
	}
	var _masters = this.slideMasters;
	var aDocumentThemes = this.Api.ThemeLoader.Themes.DocumentThemes;
	var aThemeInfo = this.Api.ThemeLoader.themes_info_document;
	aDocumentThemes.length = 0;
	aThemeInfo.length = 0;
	for (var i = 0; i < _masters.length; i++) {
		var theme_load_info = new AscCommonSlide.CThemeLoadInfo();
		theme_load_info.Master = _masters[i];
		theme_load_info.Theme = _masters[i].Theme;
		var oTheme = _masters[i].Theme;
		var _lay_cnt = _masters[i].sldLayoutLst.length;
		for (var j = 0; j < _lay_cnt; j++) {
			theme_load_info.Layouts[j] = _masters[i].sldLayoutLst[j];
		}
		var th_info = {};
		th_info.Name = typeof oTheme.name === "string" && oTheme.name.length > 0 ? oTheme.name : "Doc Theme " + (i + 1);
		th_info.Url = "";
		th_info.Thumbnail = _masters[i].ImageBase64;
		var th = new AscCommonSlide.CAscThemeInfo(th_info);
		aDocumentThemes[aDocumentThemes.length] = th;
		th.Index = -this.Api.ThemeLoader.Themes.DocumentThemes.length;
		aThemeInfo[aDocumentThemes.length - 1] = theme_load_info;
	}
	this.Api.sync_InitEditorThemes(this.Api.ThemeLoader.Themes.EditorThemes, aDocumentThemes);
	this.Api.sendEvent("asc_onUpdateThemeIndex", 0);
};

CPresentation.prototype.Check_CursorMoveRight = function () {
	var oController = this.GetCurrentController();
	if (oController) {
		if (oController.getTargetDocContent(false, false)) {
			oController.cursorMoveRight(false, false, true);
		}
	}
};


CPresentation.prototype.Get_ParentObject_or_DocumentPos = function (Index) {
	return {Type: AscDFH.historyitem_recalctype_Inline, Data: Index};
};

CPresentation.prototype.Refresh_RecalcData = function (Data) {
	var recalculateMaps, key;
	switch (Data.Type) {
		case AscDFH.historyitem_Presentation_SetFirstSlideNum: {
			for (let nSld = 0; nSld < this.Slides.length; ++nSld) {
				let oSlide = this.Slides[nSld];
				if (oSlide) {
					oSlide.refreshAllContentsFields();
				}
			}
			break;
		}
		case AscDFH.historyitem_Presentation_AddSlide:
		case AscDFH.historyitem_Presentation_RemoveSlide: {
			for (let nSld = Data.Pos; nSld < this.Slides.length; ++nSld) {
				let oSlide = this.Slides[nSld];
				if (oSlide) {
					oSlide.refreshAllContentsFields();
				}
			}
			break;
		}
		case AscDFH.historyitem_Presentation_SetDefaultTextStyle: {
			for (key = 0; key < this.Slides.length; ++key) {
				this.Slides[key].checkSlideSize();
			}
			this.RestartSpellCheck();
			break;
		}
		case AscDFH.historyitem_Presentation_SlideSize: {
			recalculateMaps = this.GetRecalculateMaps();
			for (key in recalculateMaps.masters) {
				if (recalculateMaps.masters.hasOwnProperty(key)) {
					recalculateMaps.masters[key].checkSlideSize();
				}
			}
			for (key in recalculateMaps.layouts) {
				if (recalculateMaps.layouts.hasOwnProperty(key)) {
					recalculateMaps.layouts[key].checkSlideSize();
				}
			}
			for (key = 0; key < this.Slides.length; ++key) {
				this.Slides[key].checkSlideSize();
			}
			break;
		}
		case AscDFH.historyitem_Presentation_AddSlideMaster: {
			this.bNeedUpdateThemes = true;
			break;
		}
		case AscDFH.historyitem_Presentation_RemoveSlideMaster: {
			this.bNeedUpdateThemes = true;
			break;
		}
		case AscDFH.historyitem_Presentation_ChangeTheme: {
			for (var i = 0; i < Data.aIndexes.length; ++i) {
				this.Slides[Data.aIndexes[i]] && this.Slides[Data.aIndexes[i]].checkSlideTheme();
			}
			break;
		}
		case AscDFH.historyitem_Presentation_ChangeColorScheme: {
			recalculateMaps = this.GetRecalculateMaps();
			for (key in recalculateMaps.masters) {
				if (recalculateMaps.masters.hasOwnProperty(key)) {
					recalculateMaps.masters[key].checkSlideColorScheme();
				}
			}
			for (key in recalculateMaps.layouts) {
				if (recalculateMaps.layouts.hasOwnProperty(key)) {
					recalculateMaps.layouts[key].checkSlideColorScheme();
				}
			}
			for (var i = 0; i < Data.aIndexes.length; ++i) {
				this.Slides[Data.aIndexes[i]] && this.Slides[Data.aIndexes[i]].checkSlideTheme();
			}
			break;
		}
	}
	this.Refresh_RecalcData2(Data);
};

CPresentation.prototype.Refresh_RecalcData2 = function (Data) {
	switch (Data.Type) {
		case AscDFH.historyitem_Presentation_AddSlide: {
			History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, SlideMinIdx: Data.Pos});
			break;
		}
		case AscDFH.historyitem_Presentation_RemoveSlide: {
			History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, SlideMinIdx: Data.Pos});
			break;
		}
		case AscDFH.historyitem_Presentation_SlideSize:
		case AscDFH.historyitem_Presentation_SetDefaultTextStyle: {
			History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, All: true});
			break;
		}
		case AscDFH.historyitem_Presentation_AddSlideMaster: {
			break;
		}
		case AscDFH.historyitem_Presentation_ChangeTheme: {
			History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, Theme: true, ArrInd: Data.aIndexes});
			break;
		}
		case AscDFH.historyitem_CSldViewPrGuideLst:
		case AscDFH.historyitem_ViewPrGuidePos:
		case AscDFH.historyitem_ViewPrGridSpacing:
		case AscDFH.historyitem_ViewPrSlideViewerPr: {
			History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, All: true});
			break;
		}
		case AscDFH.historyitem_ThemeSetFontScheme: {

			for (let nSlide = 0; nSlide < Data.aIndexes.length; ++nSlide) {
				let nSldIdx = Data.aIndexes[nSlide];
				let oSlide = this.Slides[nSldIdx];
				if (oSlide) {
					oSlide.checkSlideTheme();
					oSlide.addToRecalculate();
				}
			}
			break;
		}
		case AscDFH.historyitem_ThemeSetColorScheme: {
			this.bNeedUpdateThemes = true;
			History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, Theme: true, ThemeObj: Data.Theme});
			break;
		}
		case AscDFH.historyitem_ThemeSetName: {
			this.bNeedUpdateThemes = true;
			break;
		}
	}
};

//-----------------------------------------------------------------------------------
// Функции для работы с гиперссылками
//-----------------------------------------------------------------------------------
CPresentation.prototype.AddHyperlink = function (HyperProps) {
	var oController = this.GetCurrentController();
	if (oController) {
		oController.checkSelectedObjectsAndCallback(oController.hyperlinkAdd, [HyperProps], false, AscDFH.historydescription_Presentation_HyperlinkAdd);
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.ModifyHyperlink = function (HyperProps) {
	var oController = this.GetCurrentController();
	if (oController) {
		oController.checkSelectedObjectsAndCallback(oController.hyperlinkModify, [HyperProps], false, AscDFH.historydescription_Presentation_HyperlinkModify);
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.RemoveHyperlink = function () {
	var oController = this.GetCurrentController();
	if (oController) {
		oController.checkSelectedObjectsAndCallback(oController.hyperlinkRemove, [], false, AscDFH.historydescription_Presentation_HyperlinkRemove);
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.CanAddHyperlink = function (bCheckInHyperlink) {
	var oController = this.GetCurrentController();
	if (oController)
		return oController.hyperlinkCanAdd(bCheckInHyperlink);
	return false;
};

CPresentation.prototype.canGroup = function () {
	let oController = this.GetCurrentSlideController();
	if (oController)
		return oController.canGroup();
	return false
};

CPresentation.prototype.canUnGroup = function () {
	let oController = this.GetCurrentSlideController();
	if (oController)
		return oController.canUnGroup();
	return false;
};

CPresentation.prototype.getSelectedDrawingObjectsCount = function () {
	var oController = this.GetCurrentController();
	if (!oController) {
		return 0;
	}
	var aSelectedObjects = oController.selection.groupSelection ? oController.selection.groupSelection.selectedObjects : oController.selectedObjects;
	return aSelectedObjects.length;
};

CPresentation.prototype.alignLeft = function (alignType) {
	let oController = this.GetCurrentSlideController();
	if(oController) {
		oController.alignLeft(alignType === Asc.c_oAscObjectsAlignType.Selected);
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.alignRight = function (alignType) {
	let oController = this.GetCurrentSlideController();
	if(oController) {
		oController.alignRight(alignType === Asc.c_oAscObjectsAlignType.Selected);
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.alignTop = function (alignType) {
	let oController = this.GetCurrentSlideController();
	if(oController) {
		oController.alignTop(alignType === Asc.c_oAscObjectsAlignType.Selected);
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.alignBottom = function (alignType) {
	let oController = this.GetCurrentSlideController();
	if(oController) {
		oController.alignBottom(alignType === Asc.c_oAscObjectsAlignType.Selected);
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.alignCenter = function (alignType) {
	let oController = this.GetCurrentSlideController();
	if(oController) {
		oController.alignCenter(alignType === Asc.c_oAscObjectsAlignType.Selected);
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.alignMiddle = function (alignType) {
	let oController = this.GetCurrentSlideController();
	if(oController) {
		oController.alignMiddle(alignType === Asc.c_oAscObjectsAlignType.Selected);
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.distributeHor = function (alignType) {
	let oController = this.GetCurrentSlideController();
	if(oController) {
		let bSelected = (alignType === Asc.c_oAscObjectsAlignType.Selected);
		oController.checkSelectedObjectsAndCallback(oController.distributeHor, [bSelected], false, AscDFH.historydescription_Presentation_DistHor);
		this.Document_UpdateInterfaceState();
	}
};
CPresentation.prototype.distributeVer = function (alignType) {
	let oController = this.GetCurrentSlideController();
	if(oController) {
		let bSelected = (alignType === Asc.c_oAscObjectsAlignType.Selected);
		oController.checkSelectedObjectsAndCallback(oController.distributeVer, [bSelected], false, AscDFH.historydescription_Presentation_DistVer);
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.bringToFront = function () {

	let oController = this.GetCurrentSlideController();
	if(oController) {
		oController.checkSelectedObjectsAndCallback(oController.bringToFront, [], false, AscDFH.historydescription_Presentation_BringToFront);
		this.Document_UpdateInterfaceState();
	}
};
CPresentation.prototype.bringForward = function () {
	let oController = this.GetCurrentSlideController();
	if(oController) {
		oController.checkSelectedObjectsAndCallback(oController.bringForward, [], false, AscDFH.historydescription_Presentation_BringForward);
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.sendToBack = function () {
	let oController = this.GetCurrentSlideController();
	if(oController) {
		oController.checkSelectedObjectsAndCallback(oController.sendToBack, [], false, AscDFH.historydescription_Presentation_SendToBack);
		this.Document_UpdateInterfaceState();
	}
};


CPresentation.prototype.bringBackward = function () {

	let oController = this.GetCurrentSlideController();
	if(oController) {
		oController.checkSelectedObjectsAndCallback(oController.bringBackward, [], false, AscDFH.historydescription_Presentation_BringBackward);
		this.Document_UpdateInterfaceState();
	}
};

// Проверяем, находимся ли мы в гиперссылке сейчас
CPresentation.prototype.IsCursorInHyperlink = function (bCheckEnd) {
	var oController = this.GetCurrentController();
	return oController && oController.hyperlinkCheck(bCheckEnd);
};


CPresentation.prototype.RemoveBeforePaste = function () {

	var oController = this.GetCurrentController();
	if (oController) {
		var oTargetContent = oController.getTargetDocContent();
		if (oTargetContent) {
			oTargetContent.Remove(-1, true, true, true, undefined);
		}

	}
};

CPresentation.prototype.addNextSlide = function (layoutIndex) {
	if(this.IsMasterMode()) {
		return;
	}
	this.Api.inkDrawer.startSilentMode();
	History.Create_NewPoint(AscDFH.historydescription_Presentation_AddNextSlide);
	this.addNextSlideAction(layoutIndex);
	this.Recalculate();
	this.DrawingDocument.m_oWordControl.GoToPage(this.CurPage + 1);
	this.Api.inkDrawer.endSilentMode();
	this.Document_UpdateInterfaceState();
};
CPresentation.prototype.addNextSlideAction = function (layoutIndex) {
	var new_slide, layout, i, _ph_type, sp, hf, bIsSpecialPh, aLayouts, bRemoveOnTitle;
	var cur_slide = this.GetCurrentSlide();
	if (cur_slide) {
		aLayouts = this.GetAllLayouts();
		if (AscFormat.isRealNumber(layoutIndex) && aLayouts[layoutIndex]) {
			layout = aLayouts[layoutIndex];
		} else {
			if (cur_slide.Layout === aLayouts[0] && aLayouts[1]) {
				layout = aLayouts[1];
			} else {
				layout = cur_slide.Layout;
			}
		}
	}
	else {
		aLayouts = this.GetAllLayouts();
		if(AscFormat.isRealNumber(layoutIndex)) {
			layout = aLayouts[layoutIndex];
		}
		else {
			layout = aLayouts[0];
		}
	}
	hf = layout.hf || layout.Master.hf;
	new_slide = new Slide(this, layout, this.CurPage + 1);
	new_slide.setNotes(AscCommonSlide.CreateNotes());
	new_slide.notes.setNotesMaster(this.notesMasters[0]);
	new_slide.notes.setSlide(new_slide);
	bRemoveOnTitle = layout.type === AscFormat.nSldLtTTitle && this.showSpecialPlsOnTitleSld === false;
	layout.cSld.forEachSp(function(oSp) {
		if (oSp.isPlaceholder()) {
			_ph_type = oSp.getPlaceholderType();
			bIsSpecialPh = _ph_type === AscFormat.phType_dt || _ph_type === AscFormat.phType_ftr || _ph_type === AscFormat.phType_hdr || _ph_type === AscFormat.phType_sldNum;
			if (!bIsSpecialPh || hf && !bRemoveOnTitle && ((_ph_type === AscFormat.phType_dt && (hf.dt !== false)) ||
				(_ph_type === AscFormat.phType_ftr && (hf.ftr !== false)) ||
				(_ph_type === AscFormat.phType_hdr && (hf.hdr !== false)) ||
				(_ph_type === AscFormat.phType_sldNum && (hf.sldNum !== false)))) {
				sp = oSp.copy(undefined);
				sp.setParent(new_slide);
				if (sp.txBody) {
					sp.txBody.setBodyPr(new AscFormat.CBodyPr());
				}
				!bIsSpecialPh && sp.clearContent && sp.clearContent();
				new_slide.addToSpTreeToPos(new_slide.cSld.spTree.length, sp);
			}
		}
	});
	new_slide.setSlideNum(this.CurPage + 1);
	new_slide.setSlideSize(this.GetWidthMM(), this.GetHeightMM());
	this.insertSlide(this.CurPage + 1, new_slide);
	for (i = this.CurPage + 2; i < this.Slides.length; ++i) {
		this.Slides[i].setSlideNum(i);
	}
};
CPresentation.prototype.getDefaultMasterSlide = function () {
	if (this.lastMaster && this.lastMaster.sldLayoutLst.length > 0) {
		return this.lastMaster;
	}
	return this.slideMasters[0];
};
CPresentation.prototype.getLayoutsMasterSlide = function () {
	let oCurSlide = this.GetCurrentSlide();
	if(!oCurSlide ||
		!oCurSlide.Layout ||
		!oCurSlide.Layout.Master ||
		oCurSlide.Layout.Master.sldLayoutLst.length <= 0) {
		return this.getDefaultMasterSlide();
	}
	return oCurSlide.Layout.Master;
};
CPresentation.prototype.DublicateSlide = function () {
	if (this.Api.WordControl.Thumbnails) {
		var selected_slides = this.GetSelectedSlides();
		this.shiftSlides(Math.max.apply(Math, selected_slides) + 1, selected_slides, true);
	}
};
CPresentation.prototype.DublicateMaster = function () {
	if (!this.CanEdit()) {
		return;
	}
	this.StartAction(0);
	let aSelectedSlides = this.GetSelectedSlideObjects();
	let aCopyMasters = [];
	let oLastMaster;
	for(let nIdx = 0; nIdx < aSelectedSlides.length; ++nIdx) {
		let oSlide = aSelectedSlides[nIdx];
		if(oSlide.isMaster()) {
			let oMaster = oSlide.createDuplicate();
			aCopyMasters.push(oMaster);
			oLastMaster = oSlide;
		}
	}
	if(oLastMaster) {
		let nPos = this.slideMasters.length;
		for(let nIdx = this.slideMasters.length - 1; nIdx > -1; --nIdx) {
			if(this.slideMasters[nIdx] === oLastMaster) {
				nPos = nIdx + 1;
				break;
			}
		}
		for(let nIdx = 0; nIdx < aCopyMasters.length; ++nIdx) {
			this.addSlideMaster(nPos + nIdx, aCopyMasters[nIdx]);
		}
	}
	this.FinalizeAction(true);
	let oMaster = aCopyMasters[0];
	if(oMaster) {
		let nIdx = this.GetSlideIndex(oMaster);
		if(nIdx > -1) {
			this.DrawingDocument.m_oWordControl.GoToPage(nIdx);
		}
	}
};
CPresentation.prototype.DublicateLayout = function () {
	if (!this.CanEdit()) {
		return;
	}
	this.StartAction(0);
	let aSelectedSlides = this.GetSelectedSlideObjects();
	let aCopyLayouts = [];
	let oLastLayout;
	for(let nIdx = 0; nIdx < aSelectedSlides.length; ++nIdx) {
		let oSlide = aSelectedSlides[nIdx];
		if(oSlide.isLayout()) {
			let oLayout = oSlide.createDuplicate();
			aCopyLayouts.push(oLayout);
			oLastLayout = oSlide;
		}
	}
	if(oLastLayout) {
		let oMaster = oLastLayout.Master;
		let nPos = oMaster.sldLayoutLst.length;
		for(let nIdx = oMaster.sldLayoutLst.length - 1; nIdx > -1; --nIdx) {
			if(oMaster.sldLayoutLst[nIdx] === oLastLayout) {
				nPos = nIdx + 1;
				break;
			}
		}
		for(let nIdx = 0; nIdx < aCopyLayouts.length; ++nIdx) {
			oMaster.addToSldLayoutLstToPos(nPos + nIdx, aCopyLayouts[nIdx]);
		}
	}
	this.FinalizeAction(true);
	let oLayout = aCopyLayouts[0];
	if(oLayout) {
		let nIdx = this.GetSlideIndex(oLayout);
		if(nIdx > -1) {
			this.DrawingDocument.m_oWordControl.GoToPage(nIdx);
		}
	}
};

CPresentation.prototype.shiftSlides = function (pos, array, bCopy) {
	if (!this.CanEdit()) {
		return this.CurPage;
	}
	if(this.CheckIsMixedSelection(array)) {
		return this.CurPage;
	}
	if(array.length < 1) return this.CurPage;

	let oSlideLikeObject = this.GetSlide([array[0]]);
	if(oSlideLikeObject.isLayout() && pos === 0) return;
	let bCopyOnMove = (bCopy === true || AscCommon.global_mouseEvent.CtrlKey);
	History.Create_NewPoint(AscDFH.historydescription_Presentation_ShiftSlides);
	array.sort(AscCommon.fSortAscending);
	let aNewSelected = [];
	if(this.IsMasterMode()) {
		let aToInsert = [];
		if(bCopyOnMove) {
			for(let nIdx = 0; nIdx < array.length; ++nIdx) {
				let nIndexInSlides = array[nIdx];
				oSlideLikeObject = this.GetSlide(nIndexInSlides);
				aToInsert.push(oSlideLikeObject.createDuplicate({}, false));
			}
		}
		else {
			for(let nIdx = array.length - 1; nIdx > -1; --nIdx) {
				let nIndexInSlides = array[nIdx];
				oSlideLikeObject = this.GetSlide(nIndexInSlides);
				aToInsert.splice(0, 0, oSlideLikeObject);
				if(oSlideLikeObject.isMaster()) {
					this.removeSlideMasterObject(oSlideLikeObject);
				}
				else {
					oSlideLikeObject.Master.removeLayout(oSlideLikeObject);
				}
				if (nIndexInSlides < pos) {
					--pos;
				}
			}
		}
		let oPrevSlideLikeObj = this.GetSlide(pos - 1);
		let oPrevMaster;
		if(oPrevSlideLikeObj) {
			if(oPrevSlideLikeObj.isMaster()) {
				oPrevMaster = oPrevSlideLikeObj;
			}
			else {
				oPrevMaster = oPrevSlideLikeObj.Master;
			}
		}
		if(oSlideLikeObject.isMaster()) {
			let nInsertPos = null;
			if(!oPrevSlideLikeObj) {
				nInsertPos = 0;
			}
			else {
				if(oPrevMaster) {
					for(let nIdx = 0; nIdx < this.slideMasters.length; ++nIdx) {
						if(this.slideMasters[nIdx] === oPrevMaster) {
							nInsertPos = nIdx + 1;
							break;
						}
					}
				}
			}
			if(nInsertPos !== null) {
				for(let nIdx = 0; nIdx < aToInsert.length; ++nIdx) {
					this.addSlideMaster(nInsertPos + nIdx, aToInsert[nIdx]);
					aNewSelected.push(this.GetSlideIndex(aToInsert[nIdx]));
				}
			}
		}
		else {
			if(oPrevMaster) {
				let nInsertPos = null;
				let oMaster = null;
				if(oPrevSlideLikeObj.isMaster()) {
					nInsertPos = 0;
					oMaster = oPrevSlideLikeObj;
				}
				else {
					oMaster = oPrevSlideLikeObj.Master;
					for(let nIdx = 0; nIdx < oMaster.sldLayoutLst.length; ++nIdx) {
						if(oMaster.sldLayoutLst[nIdx] === oPrevSlideLikeObj) {
							nInsertPos = nIdx + 1;
							break;
						}
					}
				}
				if(oMaster !== null && nInsertPos !== null) {
					let aNewSelected = [];
					for(let nIdx = 0; nIdx < aToInsert.length; ++nIdx) {
						oMaster.addToSldLayoutLstToPos(nInsertPos + nIdx, aToInsert[nIdx]);
						aNewSelected.push(this.GetSlideIndex(aToInsert[nIdx]));
					}
				}
			}
		}
	}
	else {
		let deleted = [], i;
		if (!bCopyOnMove) {
			for (i = array.length - 1; i > -1; --i) {
				deleted.push(this.removeSlide(array[i], true));
			}

			for (i = 0; i < array.length; ++i) {
				if (array[i] < pos)
					--pos;
				else
					break;
			}
		} else {
			for (i = array.length - 1; i > -1; --i) {
				let oIdMap = {};
				let oSlideCopy = this.GetSlide([array[i]]).createDuplicate(oIdMap, false);
				AscFormat.fResetConnectorsIds(oSlideCopy.cSld.spTree, oIdMap);
				deleted.push(oSlideCopy);
			}
		}
		deleted.reverse();
		for (i = 0; i < deleted.length; ++i) {
			this.insertSlideObjectToPos(pos + i, deleted[i]);
			aNewSelected.push(pos + i);
		}
	}
	this.Recalculate();
	this.Document_UpdateUndoRedoState();
	this.DrawingDocument.OnEndRecalculate();
	if(aNewSelected.length > 0) {
		this.DrawingDocument.m_oWordControl.GoToPage(aNewSelected[0]);
		let oThumbnails = this.Api.WordControl.Thumbnails;
		if (oThumbnails) {
			oThumbnails.SelectSlides(aNewSelected);
		}
	}
};
CPresentation.prototype.deleteMaster = function() {
	if(!this.IsMasterMode()) {
		return;
	}
	let oMaster = this.GetCurrentMaster();
	if(!oMaster) return;
	let nMasterIdx = -1;
	for(let nMaster = 0; nMaster < this.slideMasters.length; ++nMaster) {
		if(this.slideMasters[nMaster] === oMaster) {
			nMasterIdx = nMaster;
			break;
		}
	}
	History.Create_NewPoint(AscDFH.historydescription_Presentation_DeleteSlides);
	let nIdx = this.GetSlideIndex(oMaster);
	this.removeSlide(nIdx);
	let nNewIdx = this.GetSlideIndex(oMaster);
	if(nNewIdx === -1) {
		let nCurIdx = 0;
		if(nMasterIdx > 0) {
			nCurIdx = nMasterIdx - 1;
		}
		let nPageIdx = this.GetSlideIndex(this.slideMasters[nCurIdx]);
		this.DrawingDocument.m_oWordControl.GoToPage(nPageIdx, undefined, undefined, true);
		this.Api.sync_HideComment();
		this.Document_UpdateUndoRedoState();
		this.Recalculate();
	}
};
CPresentation.prototype.deleteSlides = function (arrSlides) {
	const unpreserveInfo = this.getUnpreserveLayoutsAndMasters(arrSlides);
	const checkArray = [].concat(unpreserveInfo.layouts, unpreserveInfo.masters, arrSlides);
	if (arrSlides.length > 0 && (this.Document_Is_SelectionLocked(AscCommon.changestype_RemoveSlide, checkArray) === false)) {
		History.Create_NewPoint(AscDFH.historydescription_Presentation_DeleteSlides);
		let nMinSlideIndex = this.GetSlidesCount() - 1;
		for (var i = checkArray.length - 1; i > -1; --i) {
			const nSlideIndex = this.GetSlideIndex(checkArray[i]);
			if (nSlideIndex < nMinSlideIndex) {
				nMinSlideIndex = nSlideIndex;
			}
			this.removeSlideByObject(checkArray[i]);
		}
		if(!this.IsMasterMode()) {
			this.updateSlideIndexes();
		}
		this.DrawingDocument.UpdateThumbnailsAttack();
		this.DrawingDocument.m_oWordControl.GoToPage(Math.min(this.GetSlidesCount() - 1, nMinSlideIndex), undefined, undefined, true);
		this.Api.sync_HideComment();
		this.Document_UpdateUndoRedoState();
		this.Recalculate();
	}
};
CPresentation.prototype.GetAllMasters = function() {
	return [].concat(this.slideMasters);
};
CPresentation.prototype.GetAllLayouts = function() {
	let aMasters = this.GetAllMasters();
	let aLayouts = [];
	for(let nIdx = 0; nIdx < aMasters.length; ++nIdx) {
		aLayouts = aLayouts.concat(aMasters[nIdx].sldLayoutLst);
	}
	return aLayouts;
};
CPresentation.prototype.GetCurrentTheme = function() {
	let oCurSlide = this.GetCurrentSlide();
	if(oCurSlide) {
		let nType = oCurSlide.getObjectType();
		switch (nType) {
			case AscDFH.historyitem_type_Slide: {
				return oCurSlide.Layout.Master.Theme;
			}
			case AscDFH.historyitem_type_SlideLayout: {
				return oCurSlide.Master.Theme;
			}
			case AscDFH.historyitem_type_SlideMaster: {
				return oCurSlide.Theme;
			}
		}
	}
	if(this.lastMaster) {
		return this.lastMaster;
	}
	return this.slideMasters[0] || null;
};
CPresentation.prototype.changeLayout = function (_array, layout_index) {
	if(this.IsMasterMode()) return;
	let aLayouts = this.GetAllLayouts();
	let oApplyLayout = aLayouts[layout_index];
	let layout;
	if (this.Document_Is_SelectionLocked(AscCommon.changestype_Layout) === false) {
		History.Create_NewPoint(AscDFH.historydescription_Presentation_ChangeLayout);
		var oSelectionStateState = null;
		let oController = this.GetCurrentSlideController()
		if (oController) {
			oSelectionStateState = {};
			oController.Save_DocumentStateBeforeLoadChanges(oSelectionStateState);
		}
		for (var i = 0; i < _array.length; ++i) {
			let slide = this.Slides[_array[i]];
			if (!oApplyLayout) {
				layout = slide.Layout;
				slide.changeLayout(slide.Layout);
			}
			else {
				slide.changeLayout(oApplyLayout);
			}
		}
		if (oSelectionStateState) {
			oController.resetSelection();
			oController.loadDocumentStateAfterLoadChanges(oSelectionStateState, this.CurPage);
		}
		this.Recalculate();
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.clearThemeTimeouts = function () {
	if (this.startChangeThemeTimeOutId != null) {
		clearTimeout(this.startChangeThemeTimeOutId);
	}
	if (this.backChangeThemeTimeOutId != null) {
		clearTimeout(this.backChangeThemeTimeOutId);
	}
	if (this.forwardChangeThemeTimeOutId != null) {
		clearTimeout(this.forwardChangeThemeTimeOutId);
	}
};

CPresentation.prototype.changeTheme = function (themeInfo, arrInd) {
	if (this.viewMode === true) {
		return;
	}

	let arr_ind, i;
	let bReplace = false;
	if(this.IsMasterMode()) {
		let oCurSlide = this.GetCurrentSlide();
		let oCurMaster = oCurSlide.getMaster();
		if(oCurMaster === themeInfo.Master) {
			return;
		}
		for (i = 0; i < this.slideMasters.length; ++i) {
			if (this.slideMasters[i] === themeInfo.Master) {
				return;
			}
		}
		bReplace = true;
		if(oCurMaster.Theme.name === "Blank" || oCurMaster.Theme.name === "Office Theme" || oCurMaster.isPreserve()) {
			bReplace = false;
		}
		arr_ind = [];
		for(let nSlide = 0; nSlide < this.Slides.length; ++nSlide) {
			if(this.Slides[nSlide].getMaster() === oCurMaster) {
				arr_ind.push(nSlide);
			}
		}
		if(!bReplace) {
			this.addSlideMaster(this.slideMasters.length, themeInfo.Master);
		}
		else {
			for(let nMaster = 0; nMaster < this.slideMasters.length; ++nMaster) {
				if(this.slideMasters[nMaster] === oCurMaster) {
					this.removeSlideMaster(nMaster, 1);
					this.addSlideMaster(nMaster, themeInfo.Master)
					break;
				}
			}
		}
	}
	else {
		if (!Array.isArray(arrInd)) {
			let oCurMaster;
			let oCurSlide = this.GetCurrentSlide();
			arr_ind = [];
			if (oCurSlide) {
				oCurMaster = oCurSlide.Layout && oCurSlide.Layout.Master;
				for (i = 0; i < this.Slides.length; ++i) {
					let oSlide = this.Slides[i];
					let oMaster = oSlide.Layout && oSlide.Layout.Master;
					if (oMaster === oCurMaster) {
						arr_ind.push(i);
					}
				}
			}
		} else {
			arr_ind = arrInd;
		}
		for (i = 0; i < this.slideMasters.length; ++i) {
			if (this.slideMasters[i] === themeInfo.Master) {
				break;
			}
		}
		if (i === this.slideMasters.length) {
			this.addSlideMaster(this.slideMasters.length, themeInfo.Master);
		}
	}
	this.clearThemeTimeouts();

	let oCurSlide = this.GetCurrentSlide();
	let oParents = oCurSlide.getParentObjects();
	var oldMaster = oParents.master;
	var _new_master = themeInfo.Master;
	_new_master.presentation = this;
	themeInfo.Master.changeSize(this.GetWidthMM(), this.GetHeightMM());
	var oContent, oMasterSp, oMasterContent, oSp;
	if (oldMaster && oldMaster.hf) {
		themeInfo.Master.setHF(oldMaster.hf.createDuplicate());
		if (oldMaster.hf.dt !== false) {
			oMasterSp = oldMaster.getMatchingShape(AscFormat.phType_dt, null, false, {});
			if (oMasterSp) {
				oMasterContent = oMasterSp.getDocContent && oMasterSp.getDocContent();
				if (oMasterContent) {
					oSp = themeInfo.Master.getMatchingShape(AscFormat.phType_dt, null, false, {});
					if (oSp) {
						oContent = oSp.getDocContent && oSp.getDocContent();
						oContent.Copy2(oMasterContent);
					}
					for (i = 0; i < themeInfo.Master.sldLayoutLst.length; ++i) {
						oSp = themeInfo.Master.sldLayoutLst[i].getMatchingShape(AscFormat.phType_dt, null, false, {});
						if (oSp) {
							oContent = oSp.getDocContent && oSp.getDocContent();
							oContent.Copy2(oMasterContent);
						}
					}
				}
			}
		}
		if (oldMaster.hf.hdr !== false) {
			oMasterSp = oldMaster.getMatchingShape(AscFormat.phType_hdr, null, false, {});
			if (oMasterSp) {
				oMasterContent = oMasterSp.getDocContent && oMasterSp.getDocContent();
				if (oMasterContent) {
					oSp = themeInfo.Master.getMatchingShape(AscFormat.phType_hdr, null, false, {});
					if (oSp) {
						oContent = oSp.getDocContent && oSp.getDocContent();
						oContent.Copy2(oMasterContent);
					}
					for (i = 0; i < themeInfo.Master.sldLayoutLst.length; ++i) {
						oSp = themeInfo.Master.sldLayoutLst[i].getMatchingShape(AscFormat.phType_hdr, null, false, {});
						if (oSp) {
							oContent = oSp.getDocContent && oSp.getDocContent();
							oContent.Copy2(oMasterContent);
						}
					}
				}
			}
		}
		if (oldMaster.hf.ftr !== false) {
			oMasterSp = oldMaster.getMatchingShape(AscFormat.phType_ftr, null, false, {});
			if (oMasterSp) {
				oMasterContent = oMasterSp.getDocContent && oMasterSp.getDocContent();
				if (oMasterContent) {
					oSp = themeInfo.Master.getMatchingShape(AscFormat.phType_ftr, null, false, {});
					if (oSp) {
						oContent = oSp.getDocContent && oSp.getDocContent();
						oContent.Copy2(oMasterContent);
					}
					for (i = 0; i < themeInfo.Master.sldLayoutLst.length; ++i) {
						oSp = themeInfo.Master.sldLayoutLst[i].getMatchingShape(AscFormat.phType_ftr, null, false, {});
						if (oSp) {
							oContent = oSp.getDocContent && oSp.getDocContent();
							oContent.Copy2(oMasterContent);
						}
					}
				}
			}
		}
	}
	for (i = 0; i < themeInfo.Master.sldLayoutLst.length; ++i) {
		themeInfo.Master.sldLayoutLst[i].changeSize(this.GetWidthMM(), this.GetHeightMM());
	}
	var slides_array = [];
	for (i = 0; i < arr_ind.length; ++i) {
		slides_array.push(this.Slides[arr_ind[i]]);
	}
	let oReplacedMasters = {};
	let aReplacedMasters = [];
	for (i = 0; i < slides_array.length; ++i) {
		let oSlide = slides_array[i];
		let oOldMaster = oSlide.getMaster();
		if(oOldMaster) {
			if(!oReplacedMasters[oOldMaster.Id]) {
				oReplacedMasters[oOldMaster.Id] = oOldMaster;
				aReplacedMasters.push(oOldMaster);
			}
		}
		this.ChangeSlideSlideMaster(slides_array[i], _new_master);
	}

	for(let nMaster = 0; nMaster < aReplacedMasters.length; ++nMaster) {
		let oMaster = aReplacedMasters[nMaster];
		let bFound = false;
		if (oMaster.isPreserve()) {
			continue;
		}
		for(let nSlide = 0; nSlide < this.Slides.length; ++nSlide) {
			if(this.Slides[nSlide].getMaster() === oMaster) {
				bFound = true;
				break;
			}
		}
		if(!bFound) {
			this.removeSlideMasterObject(oMaster);
		}
	}

	History.Add(new AscDFH.CChangesDrawingChangeTheme(this, AscDFH.historyitem_Presentation_ChangeTheme, arr_ind));

	this.Recalculate();
	if(this.IsMasterMode()) {
		let nIdx = this.GetSlideIndex(themeInfo.Master);
		if(nIdx !== -1) {
			this.DrawingDocument.m_oWordControl.GoToPage(nIdx);
		}
	}
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.ChangeSlideSlideMaster = function(oSlide, oNewMaster) {
	if (oSlide.Layout.calculatedType == null) {
		oSlide.Layout.calculateType();
	}
	let oNewLayout = oNewMaster.getMatchingLayout(oSlide.Layout.type, oSlide.Layout.matchingName, oSlide.Layout.cSld.name, true);
	if (!isRealObject(oNewLayout)) {
		oNewLayout = oNewMaster.sldLayoutLst[0];
	}
	oSlide.setLayout(oNewLayout);
	oSlide.checkNoTransformPlaceholder();
};

CPresentation.prototype.changeColorScheme = function (colorScheme) {
	if (this.viewMode === true) {
		return;
	}

	if (!(this.Document_Is_SelectionLocked(AscCommon.changestype_Theme) === false))
		return;

	if (!(colorScheme instanceof AscFormat.ClrScheme)) {
		return;
	}
	History.Create_NewPoint(AscDFH.historydescription_Presentation_ChangeColorScheme);

	let aSlides;
	if(!this.IsMasterMode()) {
		aSlides = this.GetAllSlides();
	}
	else {
		aSlides = this.GetSelectedSlideObjects();
	}
	let oThemeMap = {};
	for(let nIdx = 0; nIdx < aSlides.length; ++nIdx) {
		let oSlide = aSlides[nIdx];
		let oTheme = oSlide.getTheme();
		if(oTheme && !oThemeMap[oTheme.Get_Id()] &&
			!oTheme.themeElements.clrScheme.isIdentical(colorScheme)) {
			oTheme.changeColorScheme(colorScheme.createDuplicate());
			oThemeMap[oTheme.Get_Id()] = oTheme;
		}
	}
	this.Recalculate();
	this.Document_UpdateInterfaceState();
};

CPresentation.prototype.removeSlide = function (pos, bNoCheck) {
	let oSlide = this.GetSlide(pos);
	return this.removeSlideByObject(oSlide, bNoCheck, pos);
};
CPresentation.prototype.removeSlideByObject = function (oSlide, bNoCheck, pos) {
	if(!oSlide) return;
	let nType = oSlide.getObjectType();
	switch (nType) {
		case AscDFH.historyitem_type_Slide: {
			pos = typeof pos === 'number' ? pos : this.Slides.indexOf(oSlide);
			History.Add(new AscDFH.CChangesDrawingsContentPresentation(this, AscDFH.historyitem_Presentation_RemoveSlide, pos, [oSlide], false));
			var aSlideComments = oSlide && oSlide.slideComments && oSlide.slideComments.comments;
			this.Api.sync_HideComment();
			if (Array.isArray(aSlideComments)) {
				for (var i = aSlideComments.length - 1; i > -1; --i) {
					var sId = aSlideComments[i].Id;
					oSlide.removeComment(sId, true);
				}
			}
			this.Slides.splice(pos, 1);
			break;
		}
		case AscDFH.historyitem_type_SlideLayout: {
			if(this.CanRemoveLayout(oSlide) || bNoCheck) {
				oSlide.Master.removeLayout(oSlide);
			}
			break;
		}
		case AscDFH.historyitem_type_SlideMaster: {
			if(this.CanRemoveMaster(oSlide)) {
				for(let nIdx = 0; nIdx < this.slideMasters.length; ++nIdx) {
					if(this.slideMasters[nIdx] === oSlide || bNoCheck) {
						this.removeSlideMaster(nIdx, 1);
						if(!bNoCheck) {
							let oNewMaster = this.slideMasters[0];
							for(let nIdx = 0; nIdx < this.Slides.length; ++nIdx) {
								let oCurSlide = this.Slides[nIdx];
								if(oCurSlide.Layout.Master === oSlide) {
									this.ChangeSlideSlideMaster(oCurSlide, oNewMaster);
								}
							}
						}
						break;
					}
				}
			}
			break;
		}
	}
	return oSlide;
}

CPresentation.prototype.CanRemoveLayout = function(oLayout) {
	for(let nIdx = 0; nIdx < this.Slides.length; ++nIdx) {
		if(this.Slides[nIdx].Layout === oLayout) return false;
	}
	return true;
};
CPresentation.prototype.CanRemoveMaster = function(oMaster) {
	return this.slideMasters.length > 1;
};

CPresentation.prototype.insertSlide = function (pos, slide) {
	History.Add(new AscDFH.CChangesDrawingsContentPresentation(this, AscDFH.historyitem_Presentation_AddSlide, pos, [slide], true));
	this.Slides.splice(pos, 0, slide);
	var aSlideComments = slide.slideComments.comments;
	for (var i = 0; i < aSlideComments.length; ++i) {
		this.Api.sync_AddComment(aSlideComments[i].Get_Id(), aSlideComments[i].Data);
	}
};
CPresentation.prototype.insertSlideObjectToPos = function (pos, slide) {
	if(this.IsMasterMode()) {
		if(slide.isMaster()) {
			let oCurSlide = this.GetSlide(pos - 1);
			let oPrevMaster = null;
			if(oCurSlide) {
				if(oCurSlide.isMaster()) {
					oPrevMaster = oCurSlide;
				}
				else {
					oPrevMaster = oCurSlide.Master;
				}
			}
			let nPos = 0;
			if(oPrevMaster) {
				for(let nIdx = 0; nIdx < this.slideMasters.length; ++nIdx) {
					if(this.slideMasters[nIdx] === oPrevMaster) {
						nPos = nIdx + 1;
						break;
					}
				}
			}
			this.addSlideMaster(nPos, slide);
		}
		else {
			let oCurSlide = this.GetSlide(pos - 1);
			let oPrevMaster = null;
			let oPrevLayout = null;
			if(oCurSlide) {
				if(oCurSlide.isMaster()) {
					oPrevMaster = oCurSlide;
				}
				else {
					oPrevLayout = oCurSlide;
					oPrevMaster = oCurSlide.Master;
				}
			}
			if(!oPrevMaster) {
				oPrevMaster = this.slideMasters[0];
			}
			if(oPrevMaster) {
				let nPos = 0;
				if(oPrevLayout) {
					for(let nIdx = 0; nIdx < oPrevMaster.sldLayoutLst.length; ++nIdx) {
						if(oPrevMaster.sldLayoutLst[nIdx] === oPrevLayout) {
							nPos = nIdx + 1;
							break;
						}
					}
				}
				oPrevMaster.addToSldLayoutLstToPos(nPos, slide);
			}
		}
	}
	else {
		this.insertSlide(pos, slide)
	}
};

CPresentation.prototype.moveSlides = function (slidesIndexes, pos) {

	var insert_pos = pos;
	var removed_slides = [];
	for (var i = slidesIndexes.length - 1; i > -1; --i) {
		removed_slides.push(this.removeSlide(slidesIndexes[i]));
	}
	removed_slides.reverse();
	for (i = 0; i < removed_slides.length; ++i) {
		this.insertSlide(insert_pos + i, removed_slides[i]);
	}
};

CPresentation.prototype.moveSelectedSlidesToEnd = function () {
	if (!this.CanEdit()) {
		return;
	}
	if(this.IsMasterMode())
		return;
	History.Create_NewPoint(AscDFH.historydescription_Presentation_MoveSlidesToEnd);
	var aSelectedIdx = this.GetSelectedSlides();
	this.moveSlides(aSelectedIdx, this.Slides.length - aSelectedIdx.length);
	this.Recalculate();
	let oThumbnails = this.Api.WordControl.Thumbnails;

	this.DrawingDocument.m_oWordControl.GoToPage(this.Slides.length - aSelectedIdx.length);
	if (oThumbnails) {
		let nCount = aSelectedIdx.length;
		let aNewSelected = [];
		for (let nIdx = 0; nIdx < nCount; nIdx++) {
			aNewSelected.push(oThumbnails.m_arrPages.length - 1 - nIdx);
		}
		oThumbnails.SelectSlides(aNewSelected);
	}
	this.Document_UpdateInterfaceState();
};
CPresentation.prototype.moveSelectedSlidesToStart = function () {
	if (!this.CanEdit()) {
		return;
	}
	if(this.IsMasterMode())
		return;
	History.Create_NewPoint(AscDFH.historydescription_Presentation_MoveSlidesToStart);
	var _selection_array = this.GetSelectedSlides();
	this.moveSlides(_selection_array, 0);
	this.Recalculate();
	let oThumbnails = this.Api.WordControl.Thumbnails;
	this.DrawingDocument.m_oWordControl.GoToPage(0);
	if (oThumbnails) {
		let aNewSelected = [];
		let nCount = _selection_array.length;
		for (let nIdx = 0; nIdx < nCount; nIdx++) {
			aNewSelected.push(nIdx);
		}
		oThumbnails.SelectSlides(aNewSelected);
	}

	this.Document_UpdateInterfaceState();
};
CPresentation.prototype.moveSlidesNextPos = function () {
	if (!this.CanEdit()) {
		return;
	}
	if(this.IsMasterMode())
		return;
	var aSelectedIdx = this.GetSelectedSlides();
	var can_move = false, first_index, i;
	for (i = aSelectedIdx.length - 1; i > -1; i--) {
		if (i === aSelectedIdx.length - 1) {
			if (aSelectedIdx[i] < this.Slides.length - 1) {
				can_move = true;
				first_index = i;
				break;
			}
		} else {
			if (Math.abs(aSelectedIdx[i] - aSelectedIdx[i + 1]) > 1) {
				can_move = true;
				first_index = i;
				break;
			}
		}
	}
	if (can_move) {
		History.Create_NewPoint(AscDFH.historydescription_Presentation_MoveSlidesNextPos);
		let aNewSelected = [];
		for (i = first_index; i > -1; --i) {
			let nOldIdx = aSelectedIdx[i];
			let nNewIdx = nOldIdx + 1;
			this.moveSlides([nOldIdx], nNewIdx);
			aNewSelected.push(nNewIdx)
		}
		this.Recalculate();

		if (aNewSelected.length > 0) {
			this.DrawingDocument.m_oWordControl.GoToPage(aNewSelected[0]);
			let oThumbnails = this.Api.WordControl.Thumbnails;
			if (oThumbnails) {
				oThumbnails.SelectSlides(aNewSelected);
			}
		}
		this.Document_UpdateInterfaceState();
	}
};
CPresentation.prototype.moveSlidesPrevPos = function () {
	if (!this.CanEdit()) {
		return;
	}
	if(this.IsMasterMode())
		return;
	var _selected_array = this.GetSelectedSlides();
	var can_move = false, first_index, i;
	for (i = 0; i < _selected_array.length; ++i) {
		if (i === 0) {
			if (_selected_array[i] > 0) {
				can_move = true;
				first_index = i;
				break;
			}
		} else {
			if (Math.abs(_selected_array[i] - _selected_array[i - 1]) > 1) {
				can_move = true;
				first_index = i;
				break;
			}
		}
	}
	if (can_move) {
		History.Create_NewPoint(AscDFH.historydescription_Presentation_MoveSlidesPrevPos);
		let aNewSelected = [];
		for (i = first_index; i < _selected_array.length; ++i) {
			let nOldIdx = _selected_array[i];
			let nNewIdx = nOldIdx - 1;
			this.moveSlides([nOldIdx], nNewIdx);
			aNewSelected.push(nNewIdx);
		}
		this.Recalculate();
		if (aNewSelected.length > 0) {
			this.DrawingDocument.m_oWordControl.GoToPage(aNewSelected[0]);
			let oThumbnails = this.Api.WordControl.Thumbnails;
			if (oThumbnails) {
				oThumbnails.SelectSlides(aNewSelected);
			}
		}
		this.Document_UpdateInterfaceState();
	}
};
//-----------------------------------------------------------------------------------
// Функции для работы с совместным редактирования
//-----------------------------------------------------------------------------------

CPresentation.prototype.IsSelectionLocked = function (nCheckType, oAdditionalData, isDontLockInFastMode, isIgnoreCanEditFlag) {
	return this.Document_Is_SelectionLocked(nCheckType, oAdditionalData, isIgnoreCanEditFlag, undefined, isDontLockInFastMode);
};
CPresentation.prototype.Document_Is_SelectionLocked = function (CheckType, AdditionalData, isIgnoreCanEditFlag, aAdditionaObjects, DontLockInFastMode) {
	if (!this.CanEdit() && true !== isIgnoreCanEditFlag)
		return true;

	if (true === AscCommon.CollaborativeEditing.Get_GlobalLock())
		return true;


	if (this.IsEmpty())
		return false;
	if (AscCommon.changestype_Document_SectPr === CheckType) {
		return true;
	}

	if (CheckType === AscCommon.changestype_None && AscCommon.isRealObject(AdditionalData) && AdditionalData.CheckType === AscCommon.changestype_Table_Properties) {
		CheckType = AscCommon.changestype_Drawing_Props;
	}


	var cur_slide = this.GetCurrentSlide();
	var slide_id;
	if (this.IsFocusOnNotes() && cur_slide.notes) {
		slide_id = cur_slide.notes.Get_Id();
	} else {
		slide_id = cur_slide.deleteLock.Get_Id();
	}

	AscCommon.CollaborativeEditing.OnStart_CheckLock();

	var oController = this.GetCurrentController();
	if (!oController) {
		return false;
	}

	if (CheckType === AscCommon.changestype_Paragraph_Content || CheckType === AscCommon.changestype_Paragraph_TextProperties) {
		var oTargetTextObject = oController.getTargetDocContent(false, true);
		if (oTargetTextObject) {
			CheckType = AscCommon.changestype_Drawing_Props;
		} else {
			return false;
		}
	}

	if (CheckType === AscCommon.changestype_Drawing_Props) {
		if (cur_slide.deleteLock.Lock.Type !== AscCommon.c_oAscLockTypes.kLockTypeMine && cur_slide.deleteLock.Lock.Type !== AscCommon.c_oAscLockTypes.kLockTypeNone)
			return true;
		var selected_objects = oController.selectedObjects;
		for (var i = 0; i < selected_objects.length; ++i) {
			var check_obj =
				{
					"type": c_oAscLockTypeElemPresentation.Object,
					"slideId": slide_id,
					"objId": selected_objects[i].Get_Id(),
					"guid": selected_objects[i].Get_Id()
				};
			selected_objects[i].Lock.Check(check_obj);
		}
		if (Array.isArray(aAdditionaObjects)) {
			for (var i = 0; i < aAdditionaObjects.length; ++i) {
				var check_obj =
					{
						"type": c_oAscLockTypeElemPresentation.Object,
						"slideId": slide_id,
						"objId": aAdditionaObjects[i].Get_Id(),
						"guid": aAdditionaObjects[i].Get_Id()
					};
				aAdditionaObjects[i].Lock.Check(check_obj);
			}
		}
	}

	if (CheckType === AscCommon.changestype_AddShape || CheckType === AscCommon.changestype_AddComment) {

		if (CheckType === AscCommon.changestype_AddComment && AdditionalData && AdditionalData.Parent === this.comments) {
			var check_obj =
				{
					"type": c_oAscLockTypeElemPresentation.Slide,
					"val": this.commentsLock.Get_Id(),
					"guid": this.commentsLock.Get_Id()
				};
			this.commentsLock.Lock.Check(check_obj);
		} else {
			if (cur_slide.deleteLock.Lock.Type !== AscCommon.c_oAscLockTypes.kLockTypeMine && cur_slide.deleteLock.Lock.Type !== AscCommon.c_oAscLockTypes.kLockTypeNone)
				return true;
			var check_obj =
				{
					"type": c_oAscLockTypeElemPresentation.Object,
					"slideId": slide_id,
					"objId": AdditionalData.Get_Id(),
					"guid": AdditionalData.Get_Id()
				};
			AdditionalData.Lock.Check(check_obj);
		}
	}
	if (CheckType === AscCommon.changestype_AddShapes) {
		if (cur_slide.deleteLock.Lock.Type !== AscCommon.c_oAscLockTypes.kLockTypeMine && cur_slide.deleteLock.Lock.Type !== AscCommon.c_oAscLockTypes.kLockTypeNone)
			return true;
		for (var i = 0; i < AdditionalData.length; ++i) {
			var check_obj =
				{
					"type": c_oAscLockTypeElemPresentation.Object,
					"slideId": slide_id,
					"objId": AdditionalData[i].Get_Id(),
					"guid": AdditionalData[i].Get_Id()
				};
			AdditionalData[i].Lock.Check(check_obj);
		}
	}

	if (CheckType === AscCommon.changestype_MoveComment) {
		if (Array.isArray(AdditionalData)) {
			for (var i = 0; i < AdditionalData.length; ++i) {
				var oCheckData = AdditionalData[i];
				if (oCheckData.slide) {
					if (oCheckData.slide.deleteLock.Lock.Type !== AscCommon.c_oAscLockTypes.kLockTypeMine && oCheckData.slide.deleteLock.Lock.Type !== AscCommon.c_oAscLockTypes.kLockTypeNone)
						return true;
					var check_obj =
						{
							"type": c_oAscLockTypeElemPresentation.Object,
							"slideId": slide_id,
							"objId": oCheckData.comment.Get_Id(),
							"guid": oCheckData.comment.Get_Id()
						};
					oCheckData.comment.Lock.Check(check_obj);
				} else {
					var check_obj =
						{
							"type": c_oAscLockTypeElemPresentation.Slide,
							"val": this.commentsLock.Get_Id(),
							"guid": this.commentsLock.Get_Id()
						};
					this.commentsLock.Lock.Check(check_obj);
				}
			}
		}
	}

	let oPres = this;
	function fCheckSlides(fGetLock, selectedSlides) {
		for (let nIdx = 0; nIdx < selectedSlides.length; ++nIdx) {
			let oSlide = selectedSlides[nIdx];
			if(oSlide) {
				let oLocker = fGetLock(oSlide);
				if(oLocker) {
					let sId = oLocker.Get_Id();
					let oCheckData = {
						"type": c_oAscLockTypeElemPresentation.Slide,
						"val": sId,
						"guid": sId
					};
					oLocker.Lock.Check(oCheckData);
				}
			}
		}
	}
	if (CheckType === AscCommon.changestype_SlideBg) {
		fCheckSlides(function (slide) {return slide.backgroundLock;}, AdditionalData);
	}
	if (CheckType === AscCommon.changestype_HdrFtr) {
		const check_obj =
			{
				"type": c_oAscLockTypeElemPresentation.Slide,
				"val": this.hdrFtrLock.Get_Id(),
				"guid": this.hdrFtrLock.Get_Id()
			};
		this.hdrFtrLock.Lock.Check(check_obj);
	}
	if (CheckType === AscCommon.changestype_SlideHide) {
		fCheckSlides(function (slide) {return slide.showLock;}, AdditionalData);
	}

	if (CheckType === AscCommon.changestype_CorePr) {
		if (this.Core) {
			this.Core.Lock.Check(
				{
					"type": c_oAscLockTypeElemPresentation.Object,
					"val": this.Core.Get_Id(),
					"guid": this.Core.Get_Id(),
					"objId": this.Core.Get_Id()
				});
		}
	}
	if (CheckType === AscCommon.changestype_CustomPr) {
		if (this.CustomProperties) {
			this.CustomProperties.Lock.Check(
				{
					"type": c_oAscLockTypeElemPresentation.Object,
					"val": this.CustomProperties.Get_Id(),
					"guid": this.CustomProperties.Get_Id(),
					"objId": this.CustomProperties.Get_Id()
				});
		}
	}

	if (CheckType === AscCommon.changestype_SlideTransition) {

		let aIdx;
		if (!AdditionalData || !AdditionalData.All) {
			aIdx = this.GetSelectedSlideObjects();
		} else {
			aIdx = this.GetAllSlides();
		}
		fCheckSlides(function (slide) {return slide.transitionLock;}, aIdx);
	}

	if (CheckType === AscCommon.changestype_Text_Props) {
		if (cur_slide.deleteLock.Lock.Type !== AscCommon.c_oAscLockTypes.kLockTypeMine && cur_slide.deleteLock.Lock.Type !== AscCommon.c_oAscLockTypes.kLockTypeNone)
			return true;
		var selected_objects = oController.selectedObjects;
		for (var i = 0; i < selected_objects.length; ++i) {
			var check_obj =
				{
					"type": c_oAscLockTypeElemPresentation.Object,
					"slideId": slide_id,
					"objId": selected_objects[i].Get_Id(),
					"guid": selected_objects[i].Get_Id()
				};
			selected_objects[i].Lock.Check(check_obj);
		}
	}

	if (CheckType === AscCommon.changestype_RemoveSlide) {
		var selected_slides = AdditionalData;
		for (var i = 0; i < selected_slides.length; ++i) {
			if (selected_slides[i].isLockedObject())
				return true;
		}
		fCheckSlides(function (slide) {return slide.deleteLock;}, selected_slides);
	}

	if (CheckType === AscCommon.changestype_Theme) {
		var check_obj =
			{
				"type": c_oAscLockTypeElemPresentation.Slide,
				"val": this.themeLock.Get_Id(),
				"guid": this.themeLock.Get_Id()
			};
		this.themeLock.Lock.Check(check_obj);
	}

	if (CheckType === AscCommon.changestype_Layout) {
		var selected_slides = this.GetSelectedSlideObjects();
		fCheckSlides(function (slide) {return slide.layoutLock;}, selected_slides);
	}
	if (CheckType === AscCommon.changestype_Timing) {
		const oSlide = this.GetCurrentSlide();
		fCheckSlides(function (slide) {return slide.timingLock;}, [oSlide]);
	}
	if (CheckType === AscCommon.changestype_ColorScheme) {
		var check_obj =
			{
				"type": c_oAscLockTypeElemPresentation.Slide,
				"val": this.schemeLock.Get_Id(),
				"guid": this.schemeLock.Get_Id()
			};
		this.schemeLock.Lock.Check(check_obj);
	}

	if (CheckType === AscCommon.changestype_SlideSize) {
		var check_obj =
			{
				"type": c_oAscLockTypeElemPresentation.Slide,
				"val": this.slideSizeLock.Get_Id(),
				"guid": this.slideSizeLock.Get_Id()
			};
		this.slideSizeLock.Lock.Check(check_obj);
	}

	if (CheckType === AscCommon.changestype_PresDefaultLang) {
		var check_obj =
			{
				"type": c_oAscLockTypeElemPresentation.Slide,
				"val": this.defaultTextStyleLock.Get_Id(),
				"guid": this.defaultTextStyleLock.Get_Id()
			};

		this.defaultTextStyleLock.Lock.Check(check_obj);
	}

	if (CheckType === AscCommon.changestype_ViewPr) {
		var check_obj =
			{
				"type": c_oAscLockTypeElemPresentation.Slide,
				"val": this.viewPrLock.Get_Id(),
				"guid": this.viewPrLock.Get_Id()
			};

		this.viewPrLock.Lock.Check(check_obj);
	}

	if (CheckType === AscCommon.changestype_None) {
		if (AdditionalData) {
			if (AscCommon.changestype_2_ElementsArray_and_Type === AdditionalData.Type) {
				if (AdditionalData.CheckType === AscCommon.changestype_Drawing_Props) {
					const arrElements = AdditionalData.Elements;
					const oMainElements = {};
					for (let i = 0; i < arrElements.length; ++i) {
						const oElement = arrElements[i].getMainGroup() || arrElements[i];
						const sElementId = oElement.Get_Id();
						if (!oMainElements[sElementId]) {
							oMainElements[sElementId] = true;
							const oCheckObj =
								{
									"type"   : c_oAscLockTypeElemPresentation.Object,
									"slideId": slide_id,
									"objId"  : sElementId,
									"guid"   : sElementId
								};
							oElement.Lock.Check(oCheckObj);
						}
					}
				}
			}
		}
	}

	var bResult = AscCommon.CollaborativeEditing.OnEnd_CheckLock(DontLockInFastMode);

	if (true === bResult) {
		this.Document_UpdateSelectionState();
		this.Document_UpdateInterfaceState();
	}

	return bResult;
};

CPresentation.prototype.Clear_CollaborativeMarks = function () {
};


CPresentation.prototype.Load_Comments = function (authors) {
	AscCommonSlide.fLoadComments(this, authors);
};

//-----------------------------------------------------------------------------------
// Функции для работы с комментариями
//-----------------------------------------------------------------------------------

CPresentation.prototype.addComment = function (comment) {
	if (AscCommon.isRealObject(this.comments)) {
		this.comments.addComment(comment);
	}
};


CPresentation.prototype.AddComment = function (CommentData, bAll) {
	let oCurSlide = this.GetCurrentSlide();
	var oSlideComments = bAll ? this.comments : (oCurSlide ? oCurSlide.slideComments : null);
	if (oSlideComments) {
		History.Create_NewPoint(AscDFH.historydescription_Presentation_AddComment);
		var Comment = new AscCommon.CComment(oSlideComments, CommentData);
		if (this.Document_Is_SelectionLocked(AscCommon.changestype_AddComment, Comment, this.IsEditCommentsMode()) === false) {
			if (!bAll) {
				var oSlide = oCurSlide;
				var aComments = oSlideComments.comments;
				Comment.selected = true;
				var selected_objects = oSlide.graphicObjects.selection.groupSelection ? oSlide.graphicObjects.selection.groupSelection.selectedObjects : oSlide.graphicObjects.selectedObjects;
				var fCommentX = 0, fCommentY = 0;
				if (selected_objects.length > 0) {
					var last_object = selected_objects[selected_objects.length - 1];
					fCommentX = last_object.x + last_object.extX;
					fCommentY = last_object.y;
				} else {
					if (oSlide) {
						fCommentX = oSlide.commentX;
						fCommentY = oSlide.commentY;
					}
				}

				var Flags = 0;
				var dd = this.Api.WordControl.m_oDrawingDocument;
				var W = dd.GetCommentWidth(Flags);
				var H = dd.GetCommentHeight(Flags);
				var fCurX = fCommentX;
				var fCurY = fCommentY;
				var nComment, oCurComment;
				var fStep = W / 2;
				while (true) {
					for (nComment = 0; nComment < aComments.length; ++nComment) {
						oCurComment = aComments[nComment];
						if (AscFormat.fApproxEqual(fCurX, oCurComment.x, 0.1) &&
							AscFormat.fApproxEqual(fCurY, oCurComment.y, 0.1)) {
							break;
						}
					}
					if (nComment === aComments.length) {
						break;
					}
					fCurX += fStep;
					fCurY += fStep;
				}
				Comment.setPosition(fCurX, fCurY);
				if (oSlide) {
					oSlide.commentX = fCommentX + W;
					oSlide.commentY = fCommentY + H;
					for (nComment = aComments.length - 1; nComment > -1; --nComment) {
						aComments[nComment].selected = false;
					}
				}
			}
			oSlideComments.addComment(Comment);
			CommentData.bDocument = bAll;
			this.Api.sync_AddComment(Comment.Get_Id(), CommentData);
			if (!bAll) {
				this.DrawingDocument.OnRecalculateSlide(this.CurPage);
				this.DrawingDocument.OnEndRecalculate();
				var Coords = this.Api.WordControl.m_oDrawingDocument.ConvertCoordsToCursorWR_Comment(Comment.x, Comment.y, this.CurPage);
				this.Api.sync_HideComment();
				this.Api.sync_ShowComment(Comment.Id, Coords.X, Coords.Y);
			}

			this.Document_UpdateInterfaceState();
			return Comment;
		} else {
			this.Document_Undo();
		}
	}
};

CPresentation.prototype.EditComment = function (Id, CommentData) {
	var comment = g_oTableId.Get_ById(Id);
	if (!comment) {
		return;
	}
	var oComments = comment.Parent;
	if (!oComments) {
		return;
	}
	var bPresComments = (oComments === this.comments);
	var nCheckType = AscCommon.changestype_MoveComment;
	if (this.Document_Is_SelectionLocked(nCheckType, [{
		comment: comment,
		slide: bPresComments ? null : oComments.slide
	}], this.IsEditCommentsMode()) === false) {
		History.Create_NewPoint(AscDFH.historydescription_Presentation_ChangeComment);
		if (!bPresComments) {
			if (AscCommon.isRealObject(oComments.slide)) {
				if (oComments.slide.num !== this.CurPage) {
					this.DrawingDocument.m_oWordControl.GoToPage(oComments.slide.num);
				}
				oComments.changeComment(Id, CommentData);
				this.Api.sync_ChangeCommentData(Id, CommentData);
				this.Recalculate()
			} else {
				return true;
			}
		} else {
			oComments.changeComment(Id, CommentData);
			this.Api.sync_ChangeCommentData(Id, CommentData);
		}
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.RemoveComment = function (Id, bSendEvent) {
	if (null === Id)
		return;

	for (var i = 0; i < this.Slides.length; ++i) {
		var comments = this.Slides[i].slideComments.comments;
		for (var j = 0; j < comments.length; ++j) {
			if (comments[j].Id === Id) {
				this.Slides[i].removeComment(Id);
				this.Recalculate();
				if (this.CurPage !== i) {
					this.DrawingDocument.m_oWordControl.GoToPage(i);
				}
				return;
			}
		}
	}
	this.comments.removeComment(Id);
	this.Api.sync_HideComment();
};

CPresentation.prototype.CanAddComment = function () {
	if (!this.CanEdit() && !this.IsEditCommentsMode())
		return false;
	if(this.IsMasterMode())
		return false;
	return true;
};

CPresentation.prototype.SelectComment = function (Id) {

};


CPresentation.prototype.GetCommentIdByGuid = function (sGuid) {
	for (var i = 0; i < this.Slides.length; ++i) {
		var comments = this.Slides[i].slideComments.comments;
		for (var j = 0; j < comments.length; ++j) {
			var oComment = comments[j];
			var oData = oComment.Data;
			if (oData) {
				if (oData.m_sGuid === sGuid) {
					return oComment.Id;
				}
				for (var t = 0; t < oData.m_aReplies.length; ++t) {
					if (oData.m_aReplies[t].m_sGuid === sGuid) {
						return oComment.Id;
					}
				}
			}
		}
	}
	return null;
};


CPresentation.prototype.ShowComment = function (Id) {

	for (var i = 0; i < this.Slides.length; ++i) {
		var comments = this.Slides[i].slideComments.comments;
		for (var j = 0; j < comments.length; ++j) {
			if (comments[j].Id === Id) {
				//this.Set_CurPage(i);
				if (this.CurPage !== i) {
					this.DrawingDocument.m_oWordControl.GoToPage(i);
				}

				var Coords = this.DrawingDocument.ConvertCoordsToCursorWR_Comment(comments[j].x, comments[j].y, i);
				this.Slides[i].showComment(Id, Coords.X, Coords.Y);
				return;
			}
		}
	}
	this.Api.sync_HideComment();
};

CPresentation.prototype.ShowComments = function () {
};

CPresentation.prototype.HideComments = function () {
	//this.GetCurrentSlide().graphicObjects.hideComment();
};
//-----------------------------------------------------------------------------------
// Функции для работы с textbox
//-----------------------------------------------------------------------------------
CPresentation.prototype.TextBox_Put = function (sText, rFonts) {
	if (true === this.CollaborativeEditing.Is_Fast() || this.Document_Is_SelectionLocked(changestype_Drawing_Props) === false) {
		History.Create_NewPoint(AscDFH.historydescription_Presentation_ParagraphAdd);
		if (!rFonts) {

			// Отключаем пересчет, включим перед последним добавлением. Поскольку,
			// у нас все добавляется в 1 параграф, так можно делать.
			this.TurnOffRecalc = true;

			for (var oIterator = sText.getUnicodeIterator(); oIterator.check(); oIterator.next()) {

				var nCharCode = oIterator.value();
				if (0x0020 === nCharCode)
					this.AddToParagraph(new AscWord.CRunSpace());
				else
					this.AddToParagraph(new AscWord.CRunText(nCharCode));

			}

			this.TurnOffRecalc = false;
			this.Recalculate();
		} else {
			var oController = this.GetCurrentController();
			if (oController) {
				oController.CreateDocContent();
				var oTargetContent = oController.getTargetDocContent(true);
				if (oTargetContent) {
					var Para = oTargetContent.GetCurrentParagraph();
					if (null === Para)
						return;

					var RunPr = Para.Get_TextPr();
					if (null === RunPr || undefined === RunPr)
						RunPr = new CTextPr();

					RunPr.RFonts = rFonts;

					var Run = new ParaRun(Para);
					Run.Set_Pr(RunPr);
					Run.AddText(sText);
					Para.Add(Run);
				}
				oController.startRecalculate();
			}
		}
	}
	this.Document_UpdateUndoRedoState();
};


CPresentation.prototype.AddShapeOnCurrentPage = function (sPreset) {

	let oDrawingObjects = this.GetCurrentSlideController();
	if (!oDrawingObjects) {
		return;
	}
	oDrawingObjects.changeCurrentState(new AscFormat.StartAddNewShape(oDrawingObjects, sPreset));
	this.OnMouseDown({}, this.GetWidthMM() / 4, this.GetHeightMM() / 4, this.CurPage);
	this.OnMouseUp({}, this.GetWidthMM() / 4, this.GetHeightMM() / 4, this.CurPage);
	this.Document_UpdateInterfaceState();
	this.Document_UpdateRulersState();
	this.Document_UpdateSelectionState();
};


CPresentation.prototype.CanEditGeometry = function () {
	if (this.IsFocusOnNotes()) {
		return false;
	}
	var oController = this.GetCurrentController();
	return oController.canEditGeometry();
};

CPresentation.prototype.StartEditGeometry = function () {
	if (this.IsFocusOnNotes()) {
		return;
	}
	var oController = this.GetCurrentController();
	return oController.startEditGeometry();
};

CPresentation.prototype.Can_CopyCut = function () {
	if (this.GetFocusObjType() === FOCUS_OBJECT_THUMBNAILS) {
		return this.GetSelectedSlides().length > 0;
	}
	var oController = this.GetCurrentController();
	if (!oController) {
		return false;
	}
	var oTargetContent = oController.getTargetDocContent(false, true);
	if (oTargetContent) {
		return oTargetContent.Can_CopyCut();
	} else {
		return oController.selectedObjects.length > 0;
	}
};

CPresentation.prototype.AddToLayout = function () {
	if (this.IsFocusOnNotes()) {
		return;
	}
	var oSlide = this.GetCurrentSlide();
	if (!oSlide) {
		return;
	}
	var oController = oSlide.graphicObjects;
	var oPresentation = this;
	oController.checkSelectedObjectsAndCallback(function () {
		var oSelectionState = oPresentation.Get_SelectionState2();
		var spTree = oSlide.cSld.spTree;
		var oLayout = oSlide.Layout;
		var k = 0;
		for (var i = spTree.length - 1; i > -1; i--) {
			var oSp = spTree[i];
			if (spTree[i].selected && !spTree[i].isPlaceholder()) {
				oSlide.removeFromSpTreeByPos(i);
				oSp.setParent(oLayout);
				oLayout.shapeAdd(oLayout.cSld.spTree.length - k, oSp);
				++k;
			}
		}
		oPresentation.Set_SelectionState2(oSelectionState);
	}, [], false, AscDFH.historydescription_Presentation_AddToLayout);
};

CPresentation.prototype.AddAnimation = function (nPresetClass, nPresetId, nPresetSubtype, oColor, bReplace, bPreview) {
	var oSlide = this.GetCurrentSlide();
	if (oSlide) {
		if (nPresetClass === AscFormat.PRESET_CLASS_PATH && nPresetId === AscFormat.MOTION_CUSTOM_PATH) {
			if (nPresetSubtype === AscFormat.MOTION_CUSTOM_PATH_CURVE) {

				oSlide.graphicObjects.changeCurrentState(new AscFormat.SplineBezierState(oSlide.graphicObjects, true, bReplace, bPreview));
			} else if (nPresetSubtype === AscFormat.MOTION_CUSTOM_PATH_SCRIBBLE) {
				oSlide.graphicObjects.changeCurrentState(new AscFormat.PolyLineAddState(oSlide.graphicObjects, true, bReplace, bPreview));
			} else {
				oSlide.graphicObjects.changeCurrentState(new AscFormat.AddPolyLine2State(oSlide.graphicObjects, true, bReplace, bPreview));
			}
			this.TurnOff_InterfaceEvents();
			return;
		}
		if (this.IsSelectionLocked(AscCommon.changestype_Timing) === false) {
			this.StartAction(0);
			var aAddedEffects = oSlide.addAnimation(nPresetClass, nPresetId, nPresetSubtype, oColor, bReplace);
			this.FinalizeAction();
			this.Document_UpdateInterfaceState();
			if (bPreview && aAddedEffects.length > 0) {
				oSlide.graphicObjects.resetSelection();
				let oTiming = this.GetCurTiming();
				oTiming.resetSelection();
				for (var nEffect = 0; nEffect < aAddedEffects.length; ++nEffect) {
					aAddedEffects[nEffect].select();
				}
				this.StartAnimationPreview();
				oTiming.checkSelectedAnimMotionShapes();
			} else {
				this.DrawingDocument.OnRecalculateSlide(this.CurPage);
			}
		}
	}
};
CPresentation.prototype.GetCurSlideObjectsNamesPairs = function () {
	var oSlide = this.GetCurrentSlide();
	if (!oSlide) {
		return []
	}
	return oSlide.cSld.getObjectsNamesPairs();
};
CPresentation.prototype.GetCurSlideObjectsNames = function () {
	var oSlide = this.GetCurrentSlide();
	if (!oSlide) {
		return []
	}
	return oSlide.cSld.getObjectsNames();
};
CPresentation.prototype.isSlidePreserved = function (nSlideIdx) {
	let oSlide = this.GetSlide(nSlideIdx);
	return oSlide instanceof AscCommonSlide.MasterSlide && oSlide.isPreserve();
};
CPresentation.prototype.isSlideAnimated = function (nSlideIdx) {
	let oSlide = this.GetSlide(nSlideIdx);
	if (!oSlide) {
		return false;
	}
	return oSlide.isAnimated();
};
CPresentation.prototype.SetAnimationProperties = function (oPr, bStartDemo) {

	var oController = this.GetCurrentController();
	if (!oController) {
		return;
	}
	var oSlide = this.GetCurrentSlide();
	if (oSlide) {


		var bStartDemo_ = bStartDemo;
		if(bStartDemo_ !== true && bStartDemo_ !== false) {
			bStartDemo_ = false;
			var oCurPr = oController.getDrawingProps().animProps;
			if (oPr && oCurPr && (oPr.asc_getSubtype() !== oCurPr.asc_getSubtype() || oCurPr.isEqualProperties(oPr))) {
				bStartDemo_ = true;
			}
		}
		if (this.IsSelectionLocked(AscCommon.changestype_Timing) === false) {
			this.StartAction(0);
			oSlide.setAnimationProperties(oPr);
			this.FinalizeAction();
			this.Document_UpdateInterfaceState();
			if (bStartDemo_) {
				this.StartAnimationPreview();
			}
		}
	}
};
CPresentation.prototype.GetCurTiming = function () {
	var oSlide = this.GetCurrentSlide();
	if (!oSlide) {
		return null;
	}
	var oTiming = oSlide.timing;
	if (!oTiming) {
		return null;
	}
	return oTiming;
};
CPresentation.prototype.CanStartAnimationPreview = function () {
	if (window['IS_NATIVE_EDITOR']) {
		return false;
	}

	var oTiming = this.GetCurTiming();
	if (!oTiming) {
		return false;
	}
	return oTiming.canStartDemo();
};
CPresentation.prototype.StartAnimationPreview = function (isAllSlideAnimations, bIncludeFollowing) {
	if (!this.CanStartAnimationPreview()) {
		return false;
	}
	var oTiming = this.GetCurTiming();
	if (!oTiming) {
		return false;
	}
	oTiming.isAllSlideAnimations = isAllSlideAnimations;
	oTiming.bIncludeFollowing = bIncludeFollowing;
	var oPlayer = oTiming.createDemoPlayer();
	if (!oPlayer) {
		return false;
	}
	var bOldLblFlag = this.Api.bIsShowAnimTab;
	this.Api.bIsShowAnimTab = false;
	this.DrawingDocument.OnRecalculateSlide(this.CurPage);
	this.Api.bIsShowAnimTab = bOldLblFlag;
	this.previewPlayer = oPlayer;
	oPlayer.start();
	this.DrawingDocument.TargetEnd();
	return true;
};
CPresentation.prototype.StopAnimationPreview = function () {
	if (this.previewPlayer) {
		if (this.previewPlayer.isStarted()) {
			this.previewPlayer.stop();
		}
		this.Api.sendEvent("asc_onAnimPreviewFinished");
		this.previewPlayer = null;
		this.UpdateSelection();
		return true;
	}
	return false;
};
CPresentation.prototype.IsStartedPreview = function () {
	if (this.previewPlayer) {
		if (this.previewPlayer.isStarted()) {
			return true;
		}
	}
	return false;
};

CPresentation.prototype.CanMoveAnimation = function (bEarlier, nPositions) {
	var oTiming = this.GetCurTiming();
	return oTiming && oTiming.canMoveAnimation(bEarlier, nPositions) || false;
};

CPresentation.prototype.MoveAnimation = function (bEarlier, nPositions) {
	History.Create_NewPoint(0);
	var oTiming = this.GetCurTiming();
	if (oTiming) {
		oTiming.moveAnimation(bEarlier, nPositions);
		this.DrawingDocument.OnRecalculateSlide(this.CurPage);
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.OnInkDrawerChangeState = function() {
	let oSlide = this.GetCurrentSlide();
	if(!oSlide) {
		return;
	}
	this.FocusOnNotes = false;
	this.SetThumbnailsFocusElement(FOCUS_OBJECT_MAIN);
	const oController = oSlide.graphicObjects;
	oController.onInkDrawerChangeState();
};

CPresentation.prototype.StartAddShape = function (preset, _is_apply, nPlaceholderType, bVertical) {
	const oCurSlide = this.GetCurrentSlide();
	if(!oCurSlide) {
		return;
	}
	let oController = oCurSlide.graphicObjects;
	if (!(_is_apply === false)) {
		this.FocusOnNotes = false;
		this.SetThumbnailsFocusElement(FOCUS_OBJECT_MAIN);
		this.Api.sync_HideComment();
		oController.startTrackNewShape(preset, nPlaceholderType, bVertical);
	} else {
		oController.clearTrackObjects();
		oController.clearPreTrackObjects();

		oController.changeCurrentState(new AscFormat.NullState(oController));
		this.DrawingDocument.m_oWordControl.OnUpdateOverlay();
		this.Api.sync_EndAddShape();
	}
};
CPresentation.prototype.EraseAllInksOnSlide = function ()
{
	if(Asc.editor.isSlideShow())
	{
		let oDemoManager = Asc.editor.getDemoManager();
		if(!oDemoManager) return;

		let oCurSlide = oDemoManager.GetCurrentSlide();
		if(!oCurSlide) return;
		let oAnnot = Asc.editor.getAnnotations();
		oAnnot.clearOnSlide(oCurSlide);
	}
};
CPresentation.prototype.SetLayoutTitle = function (bVal) {
	const oCurSlide = this.GetCurrentSlide();
	if(!oCurSlide) {
		return;
	}
	if(oCurSlide.getObjectType() !== AscDFH.historyitem_type_SlideLayout) return;
	this.StartAction(0);
	let oSp = oCurSlide.getMatchingShape(AscFormat.phType_title, null, false, {});
	if(!oSp) {
		oSp = oCurSlide.getMatchingShape(AscFormat.phType_ctrTitle, null, false, {});
	}
	if(!bVal) {
		if(oSp) {
			oCurSlide.graphicObjects.deselectObject(oSp);
			oCurSlide.removeFromSpTreeById(oSp.Get_Id());
		}
	}
	else {
		if(!oSp) {
			oSp = AscCommonSlide.CreatePlaceholder(AscFormat.phType_title);
			oCurSlide.addToSpTreeToPos(undefined, oSp);
		}
	}
	this.Recalculate()
	this.FinalizeAction(true);
};
CPresentation.prototype.SetLayoutFooter = function (bVal) {
	const oCurSlide = this.GetCurrentSlide();
	if(!oCurSlide) {
		return;
	}
	if(oCurSlide.getObjectType() !== AscDFH.historyitem_type_SlideLayout) return;
	this.StartAction(0);
	let aTypes = [AscFormat.phType_ftr, AscFormat.phType_dt, AscFormat.phType_sldNum];
	for(let nIdx = 0; nIdx < aTypes.length; ++nIdx) {
		let nType = aTypes[nIdx];
		let oSp = oCurSlide.getMatchingShape(nType, null, false, {});
		if(!bVal) {
			if(oSp) {
				oCurSlide.graphicObjects.deselectObject(oSp);
				oCurSlide.removeFromSpTreeById(oSp.Get_Id());
			}
		}
		else {
			if(!oSp) {
				oSp = oCurSlide.Master.getMatchingShape(nType, null, false, {});
				if(oSp) {
					oCurSlide.addToSpTreeToPos(undefined, oSp.copy());
				}
			}
		}
	}
	this.Recalculate()
	this.FinalizeAction(true);
};

CPresentation.prototype.canStartImageCrop = function () {
	var oCurrentController = this.GetCurrentController();
	if (!oCurrentController) {
		return false;
	}
	return oCurrentController.canStartImageCrop();
};

CPresentation.prototype.startImageCrop = function () {
	var oCurrentController = this.GetCurrentController();
	if (!oCurrentController) {
		return false;
	}
	return oCurrentController.startImageCrop();
};

CPresentation.prototype.endImageCrop = function () {
	var oCurrentController = this.GetCurrentController();
	if (!oCurrentController) {
		return false;
	}
	return oCurrentController.endImageCrop();
};


CPresentation.prototype.cropFit = function () {
	var oCurrentController = this.GetCurrentController();
	if (!oCurrentController) {
		return false;
	}
	return oCurrentController.cropFit();
};

CPresentation.prototype.cropFill = function () {
	var oCurrentController = this.GetCurrentController();
	if (!oCurrentController) {
		return false;
	}
	return oCurrentController.cropFill();
};

CPresentation.prototype.FitImagesToSlide = function () {
	var oCurrentController = this.GetCurrentController();
	if (!oCurrentController) {
		return;
	}
	oCurrentController.fitImagesToSlide();
};


/**
 * Добавляем текст в текущую позицию с заданными текстовыми настройками
 * @param sText {string}
 * @param {?AscCommon.CAddTextSettings} oSettings
 */
CPresentation.prototype.AddTextWithPr = function (sText, oSettings) {
	var oCurrentController = this.GetCurrentController();
	if (!oCurrentController) {
		return;
	}
	oCurrentController.addTextWithPr(sText, oSettings);
};

CPresentation.prototype.AddTextArt = function (nStyle) {
	let oSlide = this.GetCurrentSlide();
	if (oSlide) {
		var oDrawingObjects = oSlide.graphicObjects;
		if (oDrawingObjects.checkTrackDrawings()) {
			oDrawingObjects.endTrackNewShape();
			this.Api.sync_EndAddShape();
		}

		this.SetThumbnailsFocusElement(FOCUS_OBJECT_MAIN);

		History.Create_NewPoint(AscDFH.historydescription_Document_AddTextArt);
		var oTextArt = oDrawingObjects.createTextArt(nStyle, false);
		oTextArt.addToDrawingObjects();
		oTextArt.checkExtentsByDocContent();
		oTextArt.spPr.xfrm.setOffX((this.GetWidthMM() - oTextArt.spPr.xfrm.extX) / 2);
		oTextArt.spPr.xfrm.setOffY((this.GetHeightMM() - oTextArt.spPr.xfrm.extY) / 2);
		oDrawingObjects.resetSelection();

		if (oTextArt.bSelectedText) {
			oDrawingObjects.selectObject(oTextArt, 0);
		} else {
			var oContent = oTextArt.getDocContent();
			oContent.Content[0].Document_SetThisElementCurrent(false);
			this.SelectAll();
		}
		this.Recalculate();
		this.Document_UpdateInterfaceState();
	}
};


CPresentation.prototype.AddSignatureLine = function (oPr, Width, Height, sImgUrl) {
	let oSlide = this.GetCurrentSlide();
	if (oSlide) {
		History.Create_NewPoint(AscDFH.historydescription_Document_InsertSignatureLine);
		var fPosX = (this.GetWidthMM() - Width) / 2;
		var fPosY = (this.GetHeightMM() - Height) / 2;
		var oController = oSlide.graphicObjects;
		var Image = AscFormat.fCreateSignatureShape(oPr, false, null, Width, Height, sImgUrl);
		Image.spPr.xfrm.setOffX(fPosX);
		Image.spPr.xfrm.setOffY(fPosY);
		Image.setParent(oSlide);
		Image.addToDrawingObjects();
		oController.resetSelection();
		oController.selectObject(Image, 0);
		this.Recalculate();
		this.Document_UpdateInterfaceState();
		this.Api.sendEvent("asc_onAddSignature", Image.signatureLine.id);
	}
};

CPresentation.prototype.GetAllSignatures = function () {
	var ret = [];
	for (var i = 0; i < this.Slides.length; ++i) {
		var oController = this.Slides[i].graphicObjects;
		oController.getAllSignatures2(ret, oController.getDrawingArray());
	}
	return ret;
};

CPresentation.prototype.CallSignatureDblClickEvent = function (sGuid) {
	var ret = [], allSpr = [];
	for (var i = 0; i < this.Slides.length; ++i) {
		var oController = this.Slides[i].graphicObjects;
		allSpr = allSpr.concat(oController.getAllSignatures2(ret, oController.getDrawingArray()));
	}
	for (i = 0; i < allSpr.length; ++i) {
		if (allSpr[i].signatureLine && allSpr[i].signatureLine.id === sGuid) {
			this.Api.sendEvent("asc_onSignatureDblClick", sGuid, allSpr[i].extX, allSpr[i].extY);
		}
	}
};

CPresentation.prototype.internalCalculateData = function (aSlideComments, aWriteComments, oData) {
	aWriteComments.length = 0;

	var _comments = aSlideComments;
	var _commentsCount = _comments.length;
	for (var i = 0; i < _commentsCount; i++) {
		var _data = _comments[i].Data;
		var _commId = 0;

		var _autID = _data.m_sUserName;
		var _author = this.CommentAuthors[_autID];
		if (!_author) {
			this.CommentAuthors[_autID] = new AscCommon.CCommentAuthor();
			_author = this.CommentAuthors[_autID];
			_author.Name = _data.m_sUserName;
			_author.Calculate();

			oData._AuthorId++;
			_author.Id = oData._AuthorId;
		}

		_author.LastId++;
		_commId = _author.LastId;

		var _new_data = new AscCommon.CWriteCommentData();
		_new_data.Data = _data;
		_new_data.WriteAuthorId = _author.Id;
		_new_data.WriteCommentId = _commId;
		_new_data.WriteParentAuthorId = 0;
		_new_data.WriteParentCommentId = 0;
		_new_data.x = _comments[i].x;
		_new_data.y = _comments[i].y;

		_new_data.Calculate();
		aWriteComments.push(_new_data);

		var _comments2 = _data.m_aReplies;
		var _commentsCount2 = _comments2.length;

		for (var j = 0; j < _commentsCount2; j++) {
			var _data2 = _comments2[j];

			var _autID2 = _data2.m_sUserName;
			var _author2 = this.CommentAuthors[_autID2];
			if (!_author2) {
				this.CommentAuthors[_autID2] = new AscCommon.CCommentAuthor();
				_author2 = this.CommentAuthors[_autID2];
				_author2.Name = _data2.m_sUserName;
				_author2.Calculate();

				oData._AuthorId++;
				_author2.Id = oData._AuthorId;
			}

			_author2.LastId++;

			var _new_data2 = new AscCommon.CWriteCommentData();
			_new_data2.Data = _data2;
			_new_data2.WriteAuthorId = _author2.Id;
			_new_data2.WriteCommentId = _author2.LastId;
			_new_data2.WriteParentAuthorId = _author.Id;
			_new_data2.WriteParentCommentId = _commId;
			_new_data2.x = _new_data.x;
			_new_data2.y = _new_data.y + 136 * (j + 1); // так уж делает микрософт
			_new_data2.Calculate();
			aWriteComments.push(_new_data2);
		}
	}
};

CPresentation.prototype.CalculateComments = function () {
	this.CommentAuthors = {};
	var oData = {_AuthorId: 0};
	this.internalCalculateData(this.comments.comments, this.writecomments, oData);
	var _slidesCount = this.Slides.length;
	for (var _sldIdx = 0; _sldIdx < _slidesCount; _sldIdx++) {
		this.Slides[_sldIdx].writecomments = [];
		this.internalCalculateData(this.Slides[_sldIdx].slideComments.comments, this.Slides[_sldIdx].writecomments, oData);
	}
};

CPresentation.prototype.IsTrackRevisions = function () {
	return false;
};
CPresentation.prototype.GetLocalTrackRevisions = function () {
	return false;
};
CPresentation.prototype.SetLocalTrackRevisions = function (isTrack) {
};
CPresentation.prototype.GetGlobalTrackRevisions = function () {
	return false;
};
CPresentation.prototype.SetGlobalTrackRevisions = function (isTrack) {
};
CPresentation.prototype.IsViewModeInReview = function () {
	return false;
};
CPresentation.prototype.StartAction = function (nDescription) {
	this.Create_NewHistoryPoint(nDescription);
	this.StopAnimationPreview();
	this.Api.sendEvent("asc_onUserActionStart");
};
CPresentation.prototype.FinalizeAction = function (isCheckEmptyAction, isCheckLockedAction) {
	this.Recalculate();
	this.Api.checkChangesSize();
	if (false !== isCheckEmptyAction && AscCommon.History.Is_LastPointEmpty()) {
		AscCommon.History.RemoveLastPoint();
	} else if (isCheckLockedAction && !this.CanEdit()) {
		const arrChanges = AscCommon.History.Undo();
		this.Recalculate(this.History.Get_RecalcData(null, arrChanges));
	}
	this.Api.sendEvent("asc_onUserActionEnd");
};

CPresentation.prototype.IsSplitPageBreakAndParaMark = function () {
	return false;
};
CPresentation.prototype.IsDoNotExpandShiftReturn = function () {
	return false;
};


CPresentation.prototype.IsActionStarted = function () {
};
/**
 * Сообщаем документу, что потребуется пересчет
 */
CPresentation.prototype.Recalculate2 = function () {
	this.Recalculate();
};
/**
 * Сообщаем документу, что потребуется обновить состояние селекта
 */
CPresentation.prototype.UpdateSelection = function () {
	this.Document_UpdateSelectionState();
};
/**
 * Сообщаем документу, что потребуется обновить состояние интерфейса
 */
CPresentation.prototype.UpdateInterface = function () {
	this.Document_UpdateInterfaceState();
};
/**
 * Сообщаем документу, что потребуется обновить линейки
 */
CPresentation.prototype.UpdateRulers = function () {
	this.Document_UpdateRulersState();
};
/**
 * Сообщаем документу, что потребуется обновить состояние кнопки Unddo/Redo
 */
CPresentation.prototype.UpdateUndoRedo = function () {
	this.Document_UpdateUndoRedoState();
};


CPresentation.prototype.UpdateTracks = function () {
};


CPresentation.prototype.GetAutoCorrectSettings = function () {
	return this.AutoCorrectSettings;
};
/**
 * Устанавливаем настройку автосоздания маркированных списков
 * @param isAuto {boolean}
 */
CPresentation.prototype.SetAutomaticBulletedLists = function (isAuto) {
	this.AutoCorrectSettings.SetAutomaticBulletedLists(isAuto);
};
/**
 * Запрашиваем настройку автосоздания маркированных списков
 * @returns {boolean}
 */
CPresentation.prototype.IsAutomaticBulletedLists = function () {
	return this.AutoCorrectSettings.IsAutomaticBulletedLists();
};
/**
 * Устанавливаем настройку автосоздания нумерованных списков
 * @param isAuto {boolean}
 */
CPresentation.prototype.SetAutomaticNumberedLists = function (isAuto) {
	this.AutoCorrectSettings.SetAutomaticNumberedLists(isAuto);
};
/**
 * Запрашиваем настройку автосоздания нумерованных списков
 * @returns {boolean}
 */
CPresentation.prototype.IsAutomaticNumberedLists = function () {
	return this.AutoCorrectSettings.IsAutomaticNumberedLists();
};
/**
 * Устанавливаем параметр автозамены: заменять ли прямые кавычки "умными"
 * @param isSmartQuotes {boolean}
 */
CPresentation.prototype.SetAutoCorrectSmartQuotes = function (isSmartQuotes) {
	this.AutoCorrectSettings.SetSmartQuotes(isSmartQuotes);
};
/**
 * Запрашиваем настройку автозамены: заменять ли прямые кавычки "умными"
 * @returns {boolean}
 */
CPresentation.prototype.IsAutoCorrectSmartQuotes = function () {
	return this.AutoCorrectSettings.IsSmartQuotes();
};
/**
 * Устанавливаем параметр автозамены двух дефисов на тире
 * @param isReplace {boolean}
 */
CPresentation.prototype.SetAutoCorrectHyphensWithDash = function (isReplace) {
	this.AutoCorrectSettings.SetHyphensWithDash(isReplace);
};
/**
 * Запрашиваем настройку автозамены двух дефисов на тире
 * @returns {boolean}
 */
CPresentation.prototype.IsAutoCorrectHyphensWithDash = function () {
	return this.AutoCorrectSettings.IsHyphensWithDash();
};
/**
 * Запрашиваем настройку автозамены для французской пунктуации
 * @returns {boolean}
 */
CPresentation.prototype.IsAutoCorrectFrenchPunctuation = function () {
	return this.AutoCorrectSettings.IsFrenchPunctuation();
};
/**
 * Запрашиваем настройку автозамены двойного пробела на точку
 * @returns {boolean}
 */
CPresentation.prototype.IsAutoCorrectDoubleSpaceWithPeriod = function () {
	return this.AutoCorrectSettings.IsDoubleSpaceWithPeriod();
};
/**
 * Выставляем настройку атозамены двойного пробела на точку
 * @param {boolean} isCorrect
 */
CPresentation.prototype.SetAutoCorrectDoubleSpaceWithPeriod = function (isCorrect) {
	this.AutoCorrectSettings.SetDoubleSpaceWithPeriod(isCorrect);
};
/**
 * Выставляем настройку атозамены для первого символа предложения
 * @param {boolean} isCorrect
 */
CPresentation.prototype.SetAutoCorrectFirstLetterOfSentences = function (isCorrect) {
	this.AutoCorrectSettings.SetFirstLetterOfSentences(isCorrect);
};
/**
 * Запрашиваем настройку атозамены для первого символа предложения
 * @return {boolean}
 */
CPresentation.prototype.IsAutoCorrectFirstLetterOfSentences = function () {
	return this.AutoCorrectSettings.IsFirstLetterOfSentences();
};
/**
 * Выставляем настройку атозамены для первого символа в ячейке таблицы
 * @param {boolean} isCorrect
 */
CPresentation.prototype.SetAutoCorrectFirstLetterOfCells = function (isCorrect) {
	this.AutoCorrectSettings.SetFirstLetterOfCells(isCorrect);
}
/**
 * Запрашиваем настройку атозамены для первого символа ячейки таблицы
 * @return {boolean}
 */
CPresentation.prototype.IsAutoCorrectFirstLetterOfCells = function () {
	return this.AutoCorrectSettings.IsFirstLetterOfCells();
};
/**
 * Выставляем настройку атозамены для гиперссылок
 * @param {boolean} isCorrect
 */
CPresentation.prototype.SetAutoCorrectHyperlinks = function (isCorrect) {
	this.AutoCorrectSettings.SetHyperlinks(isCorrect);
};
/**
 * Запрашиваем настройку атозамены для гиперссылок
 * @return {boolean}
 */
CPresentation.prototype.IsAutoCorrectHyperlinks = function () {
	return this.AutoCorrectSettings.IsHyperlinks();
};
CPresentation.prototype.StopAnimation = function () {
	for (var nSlide = 0; nSlide < this.Slides.length; ++nSlide) {
		var oSlide = this.Slides[nSlide];
		var oPlayer = oSlide.animationPlayer;
		if (oPlayer) {
			oPlayer.stop();
		}
	}
};

CPresentation.prototype.createNecessaryObjectsIfNoPresent = function () {
	if (this.slideMasters.length === 0) {
		this.addSlideMaster(0, AscFormat.GenerateDefaultMasterSlide(AscFormat.GenerateDefaultTheme(this)));
	}
	if (this.slideMasters[0].sldLayoutLst.length === 0) {
		this.slideMasters[0].addLayout(AscFormat.GenerateDefaultSlideLayout(this.slideMasters[0]));
	}

	if (this.notesMasters.length === 0) {
		let oNotesMaster = AscCommonSlide.CreateNotesMaster();
		this.addNotesMaster(0, oNotesMaster);
		let oNotesTheme = this.slideMasters[0].Theme.createDuplicate();
		oNotesTheme.presentation = this;
		oNotesMaster.setTheme(oNotesTheme);
	}
	for (let nSlide = 0; nSlide < this.Slides.length; ++nSlide) {
		let oSlide = this.Slides[nSlide];
		if (!oSlide.notes) {
			oSlide.setNotes(AscCommonSlide.CreateNotes());
			oSlide.notes.setSlide(oSlide);
			oSlide.notes.setNotesMaster(this.notesMasters[0]);
		}
		if (!oSlide.notes.Master) {
			oSlide.notes.setNotesMaster(this.notesMasters[0]);
		}
	}

	if (!this.canClearGuides()) {
		this.checkViewPr().addVerticalGuide();
		this.checkViewPr().addHorizontalGuide();
	}
};
CPresentation.prototype.getDrawingObjects = function() {
	return null;
};

CPresentation.prototype.applySlideBackgroundToAll = function() {
	if(this.IsMasterMode()) return;
	if (this.Document_Is_SelectionLocked(AscCommon.changestype_SlideBg, this.Slides) === false) {
		this.StartAction(AscDFH.historydescription_Presentation_ApplyBackgroundToAll);
		const oCurSlide = this.GetCurrentSlide();
		if (!oCurSlide) {
			return;
		}
		const oBg = oCurSlide.cSld.Bg;
		const bShowMasterShapes = oCurSlide.showMasterSp;
		for (let i = 0; i < this.Slides.length; i += 1) {
			const oSlide = this.Slides[i];
			oSlide.changeBackground(null);
			oSlide.setShowMasterSp(bShowMasterShapes);
		}
		for (let i = 0; i < this.slideMasters.length; i += 1) {
			const oMaster = this.slideMasters[i];
			for (let j = 0; j < oMaster.sldLayoutLst.length; j += 1) {
				const oLayout = oMaster.sldLayoutLst[j];
				oLayout.changeBackground(null);
			}
			if (oBg) {
				oMaster.changeBackground(oBg.createFullCopy());
			}
		}

		this.FinalizeAction();
		for (let i = 0; i < this.Slides.length; i += 1) {
			this.DrawingDocument.OnRecalculateSlide(i);
		}

		this.DrawingDocument.OnEndRecalculate(true, false);
		this.Document_UpdateInterfaceState();
	}
};

CPresentation.prototype.resetSlideBackground = function(arrSlides) {
	if (this.Document_Is_SelectionLocked(AscCommon.changestype_SlideBg, arrSlides) === false) {
		this.StartAction(AscDFH.historydescription_Presentation_ResetSlideBackground);
		for(let nIdx = 0; nIdx < arrSlides.length; ++nIdx) {
			let oSlide = arrSlides[nIdx];
			if(!oSlide.isMaster()) {
				oSlide.changeBackground(null);
			}
		}
		this.FinalizeAction(true);
	}
};
CPresentation.prototype.setShowMasterSp = function(bShow, arrSlides) {
	if (this.Document_Is_SelectionLocked(AscCommon.changestype_SlideBg, arrSlides) === false) {
		this.StartAction(AscDFH.historydescription_Presentation_ShowMasterShapes);
		for (let i = 0; i < arrSlides.length; i += 1) {
			let oSlide = arrSlides[i];
			if(oSlide.setShowMasterSp) {
				oSlide.setShowMasterSp(bShow);
			}
		}
		this.FinalizeAction(true);
		for (let i = 0; i < arrSlides.length; i += 1) {
			const nIdx = this.GetSlideIndex(arrSlides[i]);
			this.DrawingDocument.OnRecalculateSlide(nIdx);
		}
		this.DrawingDocument.OnEndRecalculate(true, false);
		this.Document_UpdateInterfaceState();
	}
};
CPresentation.prototype.getLockApplyBackgroundToAll = function() {
	for (let i = 0; i < this.Slides.length; i++) {
		const oSlide = this.Slides[i];
		if (!(oSlide.backgroundLock.Lock.Type === AscCommon.c_oAscLockTypes.kLockTypeMine || oSlide.backgroundLock.Lock.Type === AscCommon.c_oAscLockTypes.kLockTypeNone)) {
			return true;
		}
	}
	return false;
};

CPresentation.prototype.AddCustomProperty = function(name, type, value) {
	if(this.Document_Is_SelectionLocked(AscCommon.changestype_CustomPr, null))
		return;
	this.StartAction(AscDFH.historydescription_CustomProperties_Add);
	this.CustomProperties.AddProperty(name, type, value);
	this.FinalizeAction(true);
};

CPresentation.prototype.ModifyCustomProperty = function(idx, name, type, value) {
	if(this.Document_Is_SelectionLocked(AscCommon.changestype_CustomPr, null))
		return;
	this.StartAction(AscDFH.historydescription_CustomProperties_Modify);
	this.CustomProperties.ModifyProperty(idx, name, type, value);
	this.FinalizeAction(true);
};

CPresentation.prototype.RemoveCustomProperty = function(idx) {
	if(this.Document_Is_SelectionLocked(AscCommon.changestype_CustomPr, null))
		return;
	this.StartAction(AscDFH.historydescription_CustomProperties_Remove);
	this.CustomProperties.RemoveProperty(idx);
	this.FinalizeAction(true);
};

CPresentation.prototype.RemoveAllInks = function () {
	const arrInks = this.GetAllInks();
	if (!arrInks.length || this.Document_Is_SelectionLocked(AscCommon.changestype_None, {Type: AscCommon.changestype_2_ElementsArray_and_Type, CheckType: AscCommon.changestype_Drawing_Props, Elements: arrInks})) {
		return;
	}

	this.StartAction(AscDFH.historydescription_RemoveAllInks);

	for (let i = 0; i < this.Slides.length; i += 1) {
		const oSlide = this.Slides[i];
		oSlide.removeAllInks();
	}
	for (let i = 0; i < this.slideMasters.length; i += 1) {
		const oSlideMaster = this.slideMasters[i];
		oSlideMaster.removeAllInks();
	}
	this.FinalizeAction();
};
CPresentation.prototype.GetAllInks = function () {
	const arrInks = [];
	for (let i = 0; i < this.Slides.length; i += 1) {
		const oSlide = this.Slides[i];
		oSlide.getAllInks(arrInks);
	}
	for (let i = 0; i < this.slideMasters.length; i += 1) {
		const oSlideMaster = this.slideMasters[i];
		oSlideMaster.getAllInks(arrInks);
	}
	return arrInks;
};
CPresentation.prototype.setPreserveSlideMaster = function (bPr) {
	if (!this.IsMasterMode()) {
		return;
	}
	const arrIndexes = this.GetSelectedSlides();
	if (!arrIndexes.length) {
		return;
	}
	arrIndexes.sort(AscCommon.fSortAscending);
	for (let i = 0; i < arrIndexes.length; i += 1) {
		const nIdx = arrIndexes[i];
		const oSlideObject = this.GetSlide(nIdx);
		if (!(oSlideObject instanceof AscCommonSlide.MasterSlide)) {
			return;
		}
	}
	const arrMasterIndexesForDelete = [];
	const arrMastersForDelete = [];
	let nSlideIndex = null;
	if (bPr) {
		nSlideIndex = arrIndexes[arrIndexes.length - 1];
	} else {
		for (let i = 0; i < arrIndexes.length; i++) {
			const nIdx = arrIndexes[i];
			const oSlideObject = this.GetSlide(nIdx);
			if (oSlideObject.IsUseInSlides()) {
				nSlideIndex = arrIndexes[i];
			} else {
				arrMasterIndexesForDelete.push(nIdx);
				arrMastersForDelete.push(oSlideObject);
			}
		}
	}

	const oThis = this;
	function callback(bDelete) {
		if (bDelete && oThis.Document_Is_SelectionLocked(AscCommon.changestype_RemoveSlide, arrMastersForDelete)) {
			return;
		}

		oThis.StartAction(AscDFH.historydescription_Presentation_SetPreserveSlideMaster);
		for (let i = 0; i < arrIndexes.length; i++) {
			const nIdx = arrIndexes[i];
			const oSlideObject = oThis.GetSlide(nIdx);
			oSlideObject.setPreserve(bPr);
		}

		if (bDelete) {
			for (let i = arrMasterIndexesForDelete.length - 1; i >= 0; i -= 1) {
				oThis.removeSlide(arrMasterIndexesForDelete[i]);
			}

			oThis.DrawingDocument.UpdateThumbnailsAttack();
			if (nSlideIndex === null) {
				nSlideIndex = Math.max(arrMasterIndexesForDelete[0] - 1, 0);
			}
			oThis.DrawingDocument.m_oWordControl.GoToPage(nSlideIndex, undefined, undefined, true);
		}

		oThis.Document_UpdateUndoRedoState();
		oThis.FinalizeAction();
	}

	if (arrMastersForDelete.length) {
		this.Api.sendEvent("asc_onRemoveUnpreserveMasters", callback);
	} else {
		callback();
	}
};
CPresentation.prototype.getUnpreserveLayoutsAndMasters = function (arrSlides) {
	const oResult = {masters: [], layouts: []};
	const mapSlides = {};
	const mapMasters = {};
	const mapLayouts = {};

	const arrUsedPresentationMasters = this.GetUsedMasters();
	const bCheckMasters = !(arrUsedPresentationMasters.length === 1 && arrUsedPresentationMasters[0] === this.slideMasters[0]);

	for (let i = 0; i < arrSlides.length; i++) {
		const oSlide = arrSlides[i];
		if (!(oSlide instanceof AscCommonSlide.Slide)) {
			continue;
		}
		mapSlides[oSlide.Id] = oSlide;
		const oSlideLayout = oSlide.Layout;
		if (oSlideLayout) {
			if (!oSlideLayout.preserve) {
				mapLayouts[oSlideLayout.Id] = oSlideLayout;
			}

			const oSlideMaster = oSlideLayout.Master;
			if (bCheckMasters && oSlideMaster && !oSlideMaster.preserve) {
				mapMasters[oSlideMaster.Id] = oSlideMaster;
			}
		}
	}
	for (let i = 0; i < this.Slides.length; i++) {
		const oSlide = this.Slides[i];
		if (!mapSlides[oSlide.Id]) {
			const oLayout = oSlide.Layout;
			if (oLayout) {
				delete mapLayouts[oLayout.Id];
				const oMaster = oLayout.Master;
				if (oMaster) {
					delete mapMasters[oMaster.Id];
				}
			}
		}
	}

	for (let i in mapMasters) {
		const oMaster = mapMasters[i];
		oResult.masters.push(oMaster);
	}

	for (let i in mapLayouts) {
		const oLayout = mapLayouts[i];
		const oMaster = oLayout.Master;
		if (oMaster && !mapMasters[oMaster.Id]) {
			oResult.layouts.push(oLayout);
		}
	}

	return oResult;
};
CPresentation.prototype.isPreserveSelectionSlides = function() {
	const arrSlides = this.GetSelectedSlideObjects();
	for (let i = 0; i < arrSlides.length; i += 1) {
		const oSlide = arrSlides[i];
		if (!oSlide.isPreserve()) {
			return false;
		}
	}

	return true;
};
function collectSelectedObjects(aSpTree, aCollectArray, bRecursive, oIdMap, bSourceFormatting) {
	var oSp;
	var oPr = new AscFormat.CCopyObjectProperties();
	oPr.idMap = oIdMap;
	oPr.bSaveSourceFormatting = bSourceFormatting;
	for (var i = 0; i < aSpTree.length; ++i) {
		oSp = aSpTree[i];
		// if(oSp.isEmptyPlaceholder())
		// {
		//     continue;
		// }
		if (oSp.selected) {
			var oCopy;
			if (oSp.isGroupObject()) {
				oCopy = oSp.copy(oPr);
				oCopy.setParent(oSp.parent);
			} else {
				if (!bSourceFormatting) {
					oCopy = oSp.copy(oPr);
					oCopy.setParent(oSp.parent);
					if (oSp.isPlaceholder && oSp.isPlaceholder()) {
						oCopy.x = oSp.x;
						oCopy.y = oSp.y;
						oCopy.extX = oSp.extX;
						oCopy.extY = oSp.extY;
						oCopy.rot = oSp.rot;
						AscFormat.CheckSpPrXfrm(oCopy, true);
					}
					oCopy.convertFromSmartArt(true);
				} else {
					oCopy = oSp.getCopyWithSourceFormatting();
					oCopy.convertFromSmartArt(true);
					oCopy.setParent(oSp.parent);
				}

			}

			aCollectArray.push(new DrawingCopyObject(oCopy, oSp.x, oSp.y, oSp.extX, oSp.extY, oSp.getBase64Img()));
			if (AscCommon.isRealObject(oIdMap)) {
				oIdMap[oSp.Id] = oCopy.Id;
			}
		}
		if (bRecursive && oSp.isGroupObject()) {
			collectSelectedObjects(oSp.spTree, aCollectArray, bRecursive, oIdMap, bSourceFormatting);
		}
	}
}

function IdList(name) {
	AscFormat.CBaseNoIdObject.call(this);
	this.name = name;
	this.list = [];
}

AscFormat.InitClass(IdList, AscFormat.CBaseNoIdObject, undefined);
let MIN_SLD_MASTER_ID = 0x80000000;
let MIN_SLD_ID = 0xFF;
let MIN_SLD_LAYOUT_ID = 0x80000000;

AscFormat.MIN_SLD_MASTER_ID = MIN_SLD_MASTER_ID;
AscFormat.MIN_SLD_ID = MIN_SLD_ID;
AscFormat.MIN_SLD_LAYOUT_ID = MIN_SLD_LAYOUT_ID;


function CPresentationProperties(oPresentation) {
	AscFormat.CBaseNoIdObject.call(this);
	this.presentation = oPresentation;
}

AscFormat.InitClass(CPresentationProperties, AscFormat.CBaseNoIdObject, 0);

function isSlideLikeObject(oObjectToCheck) {
	if(!oObjectToCheck) return false;
	if(oObjectToCheck.getObjectType) {
		let nType = oObjectToCheck.getObjectType();
		if(nType === AscDFH.historyitem_type_Slide ||
		nType === AscDFH.historyitem_type_SlideLayout ||
		nType === AscDFH.historyitem_type_SlideMaster) {
			return true;
		}
	}
	return false;
}

//------------------------------------------------------------export----------------------------------------------------
window['AscCommonSlide'] = window['AscCommonSlide'] || {};
window['AscCommonSlide'].CPresentation = CPresentation;
window['AscCommonSlide'].CPrSection = CPrSection;
window['AscCommonSlide'].CSlideSize = CSlideSize;
window['AscCommonSlide'].IdList = IdList;
window['AscCommonSlide'].CONFORMANCE_STRICT = CONFORMANCE_STRICT;
window['AscCommonSlide'].CONFORMANCE_TRANSITIONAL = CONFORMANCE_TRANSITIONAL;


window['AscFormat'] = window['AscFormat'] || {};
window['AscFormat'].CShowPr = CShowPr;
window['AscFormat'].CPresentationProperties = CPresentationProperties;
window['AscFormat'].isSlideLikeObject = isSlideLikeObject;

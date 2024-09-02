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

(function (window)
{
	const drawingDocument = {
		CanvasHit : null,
		CanvasHitContext : null,
		GoToPage: function (slideNumber)
		{
			AscTest.DrawingDocument.m_oLogicDocument.Set_CurPage(slideNumber);
			editor.WordControl.m_oDrawingDocument.SlideCurrent = slideNumber;
			this.Thumbnails && this.Thumbnails.SelectPage(slideNumber);
		},
		ConvertCoordsToCursorWR: function () {return {X:0,Y:0}},
		Notes_GetWidth: function () {return 100;},
		Notes_OnRecalculate: function () {return 100;},
		scrollToY: function () {},
		OnStartRecalculate: function () {},
		OnRecalculateSlide: function () {},
		OnEndRecalculate: function () {},
		UpdateTargetTransform : function(){},
		SetTargetColor : function(){},
		ClearCachePages : function(){},
		FirePaint : function(){},
		SetTargetSize : function(){},
		UpdateTarget : function(){},
		Set_RulerState_Paragraph : function(){},
		SelectEnabled : function(){},
		SelectShow : function(){},
		TargetStart : function(){},
		SetCursorType : function(){},
		TargetShow : function(){},
		TargetEnd : function(){},
		Update_MathTrack : function(){},
		CheckTableStyles : function(){},
		OnUpdateOverlay : function(){},
		AddPageSelection : function(){},
		IsTrackText  : function(){},
		Set_RulerState_Table  : function(){},
		SelectClear : function() {},
		GetDotsPerMM : function(value) {return 72;},
		GetMMPerDot : function(value){return value / this.GetDotsPerMM(1);},
		m_oNotesApi: {},
		clear: function () {},
		GetSlidesCount: function () {
			return editor.getCountSlides();
		},
		getGraphicController: function() {
			return AscTest.DrawingDocument.m_oLogicDocument.GetCurrentController();
		}
	};

	drawingDocument.CanvasHit = document.createElement('canvas');
	drawingDocument.CanvasHitContext = drawingDocument.CanvasHit.getContext('2d');

	window['asc_docs_api'] = AscCommon.baseEditorsApi;

	const editor = new AscCommon.baseEditorsApi({});
	editor.WordControl = drawingDocument;
	editor.WordControl.m_oDrawingDocument = drawingDocument;
	editor.WordControl.m_oDrawingDocument.m_oWordControl = drawingDocument;
	editor.WordControl.m_oApi = editor;

	editor.textArtPreviewManager = drawingDocument;

	editor.asc_hideComments = function () {};
	editor.isSlideShow = function () {return false};
	editor.sync_HideComment = function () {};
	editor.sync_CanUndoCallback = function () {};
	editor.sync_CanRedoCallback = function () {};
	editor.CheckChangedDocument = function () {};
	editor.sync_BeginCatchSelectedElements = function(){};
	editor.ClearPropObjCallback = function(){};
	editor.sync_slidePropCallback = function(){};
	editor.sync_PrLineSpacingCallBack = function(){};
	editor.sync_EndCatchSelectedElements = function(){};
	editor.sync_CanAddHyperlinkCallback = function(){};
	editor.sync_shapePropCallback = function(){};
	editor.sync_VerticalTextAlign = function(){};
	editor.sync_Vert = function(){};
	editor.sync_animPropCallback = function(){};
	editor.sync_ImgPropCallback = function(){};
	editor.sync_TblPropCallback = function(){};
	editor.sync_EndCatchSelectedElements = function(){};
	editor.Update_ParaTab = function(){};
	editor.sync_ParaStyleName = function(){};
	editor.UpdateParagraphProp = function(){};
	editor.UpdateTextPr = function(){};
	editor.sync_MathPropCallback = function(){};
	editor.sync_HyperlinkPropCallback = function(){};
	editor.DemonstrationReporterEnd = function () {};
	editor.private_GetLogicDocument = function(){return this.WordControl.m_oLogicDocument;};
	editor.asc_getKeyboardLanguage = function(){return -1;};
	editor.isMasterMode = function(){return false;};
	editor.getCountSlides = function(){
		let oPresentation = this.private_GetLogicDocument();
		if(!oPresentation) return 0;
		return oPresentation.GetSlidesCount();
	};
	editor.initCollaborativeEditing = AscCommon.SlideEditorApi.prototype.initCollaborativeEditing.bind(editor);
	//--------------------------------------------------------export----------------------------------------------------
	AscTest.DrawingDocument = editor.WordControl.m_oDrawingDocument;
	AscTest.Editor = editor;

	window.editor = editor;
	Asc.editor = editor;
	editor.initCollaborativeEditing();

})(window);

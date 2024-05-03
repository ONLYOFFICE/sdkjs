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

(function(window)
{
	const drawingDocument = {
		CanvasHit : null,
		CanvasHitContext : null,

		OnStartRecalculate : function(){},
		OnRecalculatePage : function(){},
		OnEndRecalculate : function(){},
		UpdateTargetTransform : function(){},
		SelectEnabled : function(){},
		SelectShow : function(){},
		TargetStart : function(){},
		TargetShow : function(){},
		TargetEnd : function(){},
		showTarget : function(){},
		Set_RulerState_Start : function(){},
		Set_RulerState_Paragraph : function(){},
		Set_RulerState_End : function(){},
		Update_MathTrack : function(){},
		OnDrawContentControl : function(){},
		Update_FieldTrack : function(){},
		SetTargetColor : function(){},
		SetTargetSize : function(){},
		UpdateTarget : function(){},
		ClearCachePages : function(){},
		OnRepaintPage : function(){},
		FirePaint : function(){},
		GetMMPerDot : function(value){return value / this.GetDotsPerMM(1);},
		GetDotsPerMM : function(value) {return 72;},
		EndTrackTable : function() {},
		SetCurrentPage : function(pageNum) {},
		SelectClear : function() {},
		Start_CollaborationEditing : function() {},
		End_CollaborationEditing : function() {}
	};

	drawingDocument.CanvasHit = document.createElement('canvas');
	drawingDocument.CanvasHitContext = drawingDocument.CanvasHit.getContext('2d');
	
	window['asc_docs_api'] = AscCommon.baseEditorsApi;
	
	let _callbacks = {};
	const editor = new AscCommon.baseEditorsApi({});
	editor.WordControl = drawingDocument;
	editor.WordControl.m_oDrawingDocument = drawingDocument;
	editor.WordControl.m_oDrawingDocument.m_oWordControl = drawingDocument;
	editor.WordControl.m_oApi = editor;
	editor.sync_BeginCatchRevisionsChanges = function(){};
	editor.sync_EndCatchRevisionsChanges = function(){};
	editor.sync_ChangeCommentLogicalPosition = function(){};
	editor.GetCollaborativeMarksShowType = function(){return c_oAscCollaborativeMarksShowType.All;};
	editor.asc_OnChangeContentControl = function(){};
	editor.sync_OnAllRequiredFormsFilled = function(){};
	editor.asc_OnFocusContentControl = function(){};
	editor.asc_OnBlurContentControl = function(){};
	editor.sync_CanUndoCallback = function(){};
	editor.sync_CanRedoCallback = function(){};
	editor.CheckChangedDocument = function(){};
	editor.asc_GetRevisionsChangesStack = function(){return []};
	editor.private_GetLogicDocument = function(){return this.WordControl.m_oLogicDocument;};
	editor.asc_getKeyboardLanguage = function(){return -1;};
	editor.GenerateStyles = function(){};
	editor.sync_BeginCatchSelectedElements = function(){};
	editor.sync_EndCatchSelectedElements = function(){};
	editor.ClearPropObjCallback = function(){};
	editor.Update_ParaTab = function(){};
	editor.UpdateParagraphProp = function(){};
	editor.UpdateTextPr = function(){};
	editor.sync_CanAddHyperlinkCallback = function(){};
	editor.sync_PageOrientCallback = function(){};
	editor.sync_DocSizeCallback = function(){};
	editor.sync_ColumnsPropsCallback = function(){};
	editor.sync_LineNumbersPropsCollback = function(){};
	editor.sync_SectionPropsCallback = function(){};
	editor.sendEvent = function()
	{
		var name = arguments[0];
		if (_callbacks.hasOwnProperty(name))
		{
			for (var i = 0; i < _callbacks[name].length; ++i)
			{
				_callbacks[name][i].apply(this || window, Array.prototype.slice.call(arguments, 1));
			}
			return true;
		}
		return false;
	};
	editor.asc_registerCallback = function(name, callback)
	{
		if (!_callbacks.hasOwnProperty(name))
			_callbacks[name] = [];
		_callbacks[name].push(callback);
	};
	editor.asc_unregisterCallback = function(name, callback)
	{
		if (_callbacks.hasOwnProperty(name))
		{
			for (var i = _callbacks[name].length - 1; i >= 0; --i)
			{
				if (_callbacks[name][i] === callback)
					_callbacks[name].splice(i, 1);
			}
		}
	};
	editor.getSelectionState = function()
	{
		return AscTest.GetLogicDocument().GetSelectionState();
	};
	editor.getSpeechDescription = function()
	{
		return AscTest.GetLogicDocument().getSpeechDescription(...arguments);
	};
	editor._addRemoveSpaceBeforeAfterParagraph = AscCommon.DocumentEditorApi.prototype._addRemoveSpaceBeforeAfterParagraph.bind(editor);
	editor.asc_addSpaceBeforeParagraph = AscCommon.DocumentEditorApi.prototype.asc_addSpaceBeforeParagraph.bind(editor);
	editor.asc_addSpaceAfterParagraph = AscCommon.DocumentEditorApi.prototype.asc_addSpaceAfterParagraph.bind(editor);
	editor.asc_removeSpaceBeforeParagraph = AscCommon.DocumentEditorApi.prototype.asc_removeSpaceBeforeParagraph.bind(editor);
	editor.asc_removeSpaceAfterParagraph = AscCommon.DocumentEditorApi.prototype.asc_removeSpaceAfterParagraph.bind(editor);
	editor.asc_haveSpaceBeforeParagraph = AscCommon.DocumentEditorApi.prototype.asc_haveSpaceBeforeParagraph.bind(editor);
	editor.asc_haveSpaceAfterParagraph = AscCommon.DocumentEditorApi.prototype.asc_haveSpaceAfterParagraph.bind(editor);
	
	//--------------------------------------------------------export----------------------------------------------------
	AscTest.DrawingDocument = drawingDocument;
	AscTest.Editor          = editor;

	window.editor = editor;
	Asc['editor'] = Asc.editor = editor;

})(window);

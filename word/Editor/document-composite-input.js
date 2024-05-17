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
	 * Класс для композитного ввода в документ
	 * @param document {AscWord.CDocument} можно ли разделять run во время ввода
	 * @constructor
	 */
	function DocumentCompositeInput(document)
	{
		this.document       = document;
		this.compositeInput = new AscWord.RunCompositeInput(true);
		this.inUse          = false;
		this.documentState  = null;
	}
	
	DocumentCompositeInput.prototype.isInProgress = function()
	{
		return this.inUse;
	};
	DocumentCompositeInput.prototype.begin = function()
	{
		if (this.inUse)
			return;
		
		let document = this.document;
		if (document.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, document.IsFormFieldEditing()))
			return false;
		
		this.documentState = document.GetSelectionState();
		document.StartAction(AscDFH.historydescription_Document_CompositeInput);
		
		if (document.IsDrawingSelected())
			document.GetDrawingObjects().CreateDocContent();
		
		document.RemoveBeforePaste();
		let paragraph = document.GetCurrentParagraph();
		let run = paragraph ? paragraph.Get_ElementByPos(paragraph.Get_ParaContentPos(false, false)) : null;
		
		if (run && (run instanceof AscWord.CRun))
		{
			let runParent = run.GetParent();
			if (runParent instanceof AscWord.CInlineLevelSdt && runParent.IsPlaceHolder())
			{
				runParent.ReplacePlaceHolderWithContent(false);
				run = runParent.GetElement(0);
			}
		}
		
		let result;
		if (run && (run instanceof AscWord.CRun))
		{
			this.compositeInput.begin(run);
			result = true;
		}
		else
		{
			document.CancelAction();
			result = false;
		}
		
		document.UpdateSelection();
		document.FinalizeAction(false);
		return this.inUse = result;
	};
	DocumentCompositeInput.prototype.end = function()
	{
		if (!this.inUse)
			return;
		
		this.validateInput();
		this.compositeInput.end();
		
		// UpdateInterface is necessary here since we need to fire the Api.CheckChangedDocument event
		// This event was blocked util the end of the composite input
		this.document.UpdateInterface();
		this.document.private_UpdateCursorXY(true, true);
		
		// TODO: In some cases underline of composite input isn't automatically cleared
		this.document.Redraw();
		
		this.inUse = false;
		this.documentState = null;
	};
	DocumentCompositeInput.prototype.replace = function(codePoints)
	{
		let runInput = this.compositeInput;
		this.doAction(function()
		{
			runInput.replace(codePoints);
		});
	};
	DocumentCompositeInput.prototype.remove = function(count)
	{
		let runInput = this.compositeInput;
		this.doAction(function()
		{
			runInput.remove(count);
		});
	};
	DocumentCompositeInput.prototype.add = function(codePoint)
	{
		let runInput = this.compositeInput;
		this.doAction(function()
		{
			runInput.add(codePoint);
		});
	};
	DocumentCompositeInput.prototype.setPos = function(pos)
	{
		if (!this.inUse)
			return;
		
		this.compositeInput.setPos(pos);
	};
	DocumentCompositeInput.prototype.getPos = function()
	{
		if (!this.inUse)
			return 0;
		
		return this.compositeInput.getPos();
	};
	DocumentCompositeInput.prototype.getMaxPos = function()
	{
		if (!this.inUse)
			return 0;
		
		return this.compositeInput.getLength();
	};
	DocumentCompositeInput.prototype.checkState = function()
	{
		if (!this.inUse)
			return;
		
		let run = this.compositeInput.getRun();
		if (!run.IsUseInDocument())
			AscCommon.g_inputContext.externalEndCompositeInput();
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	DocumentCompositeInput.prototype.validateInput = function()
	{
		let run = this.compositeInput.getRun();
		let parentForm = run.GetParentForm();
		if (0 !== this.compositeInput.getLength()
			&& (!parentForm || this.document.GetFormsManager().ValidateChangeOnFly(parentForm)))
			return;
		
		let changes = AscCommon.History.UndoCompositeInput(this.documentState);
		if (changes)
		{
			AscCommon.History.ClearRedo();
			this.document.UpdateAfterUndoRedo(changes);
		}
	};
	DocumentCompositeInput.prototype.doAction = function(actionFunc)
	{
		if (!this.inUse)
			return;
		
		let document = this.document;
		document.StartAction(AscDFH.historydescription_Document_CompositeInputReplace);
		
		actionFunc();
		
		document.Recalculate();
		document.UpdateSelection();
		document.UpdateUndoRedo();
		document.FinalizeAction(false);
		
		document.private_UpdateCursorXY(true, true);
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.DocumentCompositeInput = DocumentCompositeInput;
	
})(window);


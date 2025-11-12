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

	window["AscCommon"] = window.AscCommon = (window["AscCommon"] || {});

	/**
	 * 1st version:
	 *
	 * 0) Вводим текст - произносим его. Copy/Paste не произносим.
	 * --- SpeechWorker.speech(AscCommon.SpeechWorkerCommands.Text, "a");
	 * 1) Ходим курсором по тексту - произносим следующую а курсором букву. Если пробел - присылаем пустой текст.
	 * --- SpeechWorker.speech(AscCommon.SpeechWorkerCommands.Text, "a");
	 * 2) Ходим по тексту по словам - произносим следующее за курсором слово. Если конец - посылаем пустой текст.
	 * --- SpeechWorker.speech(AscCommon.SpeechWorkerCommands.Text, "hello");
	 * 3) Селект/УменьшениеСелета по клавиатуре/конец селекта мышью - произносится изменение в селекте (новый текст/тот что ушел из селекта).
	 * --- SpeechWorker.speech(AscCommon.SpeechWorkerCommands.TextSelected, { text: "текст", isBefore: false });
	 * --- SpeechWorker.speech(AscCommon.SpeechWorkerCommands.TextUnselected, { text: "текст", isBefore: false });
	 * 4) Селект автофигуры/диаграммы/картинки/...
	 * --- SpeechWorker.speech(AscCommon.SpeechWorkerCommands.DrawingSelected, { altText: "текст" });
	 * 5) Селект слайда
	 * --- SpeechWorker.speech(AscCommon.SpeechWorkerCommands.SlideSelected, { num: 1 });
	 * 6) Ходим по ячейкам в Cell
	 * --- SpeechWorker.speech(AscCommon.SpeechWorkerCommands.CellSelected, { text: "cell value", cell: "A1" });
	 * 7) Селект/УменьшениеСелета по клавиатуре/конец селекта мышью - смотрим,
	 * если +/- одна ячейка, то используем CellRangeSelectedChangeOne/CellRangeUnselectedChangeOne
	 * если нет - то CellRangeSelected/CellRangeUnselected
	 * 8) Ходим по листам - даем информацию о нем
	 * --- SpeechWorker.speech(AscCommon.SpeechWorkerCommands.SheetSelected, { ... });
	 *
	 */

	// types for SpeechWorker.speech method
	var SpeechWorkerType = {

		// text
		Text : 0,

		// { text: "text", isBefore: true  }
		TextSelected : 1,

		// { text: "text", isBefore: true  }
		TextUnselected : 2,

		// { indexes: [1, 2, ..] }
		SlidesSelected : 3,

		// { indexes: [1, 2, ..] }
		SlidesUnselected : 4,

		// { altText: "text" }
		DrawingSelected : 5,

		// { text: "text", cell: "A1" }
		CellSelected : 6,

		// { start: { text: "text", cell: "A1" }, end: { text: "text", cell: "A2" }] } }
		CellRangeSelected : 7,

		// { start: { text: "text", cell: "A1" }, end: { text: "text", cell: "A2" }] } }
		CellRangeUnselected : 8,

		// { text: "text", cell: "A1" }
		CellRangeSelectedChangeOne : 9,

		// { text: "text", cell: "A1" }
		CellRangeUnselectedChangeOne : 10,

		// { text: "text", ranges: [startCell: "A1" , endCell: "A2" }] }
		MultipleRangesSelected : 11,

		// { name: "sheet 1", cell: "A1", text: "text", cellEnd: "D5", cellsCount: 10, objectsCount: 5 }
		SheetSelected : 12

	};
	
	/**
	 * @constructor
	 */
	function CWorkerSpeech()
	{
		this.isEnabled = false;
		this.speechElement = null;
		this.isLogEnabled = false;

		this.timerSetValue = -1;
		this.value = "";
		this.valueEqualAddon = false;

		this.equalAddon = "&nbsp;";
		if (AscCommon.AscBrowser.isMacOs)
			this.equalAddon = "<br/>";

		this.setEnabled = function(isEnabled)
		{
			if (this.isEnabled === isEnabled)
				return;

			if (!AscCommon.g_inputContext)
			{
				var worker = this;
				AscCommon.inputMethodAddInitEvent(function() {
					worker.setEnabled(isEnabled);
				});
				return;
			}

			this.isEnabled = isEnabled;
			if (this.isEnabled)
			{
				this.speechElement = document.createElement("div");
				this.speechElement.innerHTML = "";
				this.speechElement.id = "area_id_screen_reader";

				let style = "position: absolute; left:0; top:-1; z-index: -2; opacity: 0;";
				this.speechElement.setAttribute("style", style);

				this.speechElement.setAttribute("role", "region");
				this.speechElement.setAttribute("aria-label", "");
				this.speechElement.setAttribute("aria-live", "assertive");
				if (!(AscCommon.AscBrowser.isMozilla && AscCommon.AscBrowser.isWindows))
					this.speechElement.setAttribute("aria-atomic", "true");
				this.speechElement.setAttribute("aria-hidden", "false");

				//AscCommon.g_inputContext.HtmlArea.setAttribute("aria-describedby", "area_id_screen_reader");
				AscCommon.g_inputContext.HtmlDiv.appendChild(this.speechElement);
			}
			else if (this.speechElement)
			{
				AscCommon.g_inputContext.HtmlArea.removeAttribute("aria-describedby");
				AscCommon.g_inputContext.HtmlDiv.removeChild(this.speechElement);
				this.speechElement = null;
			}
		};

		this._log = function(message)
		{
			if (!this.isLogEnabled)
				return;
			console.log(message);
		};

		// Variants for text setting

		// 1) Resolve the problem with equals with temporary set empty value
		this._setValueWithCheckEqual = function(value)
		{
			if (-1 !== this.timerSetValue)
			{
				clearTimeout(this.timerSetValue);
				this.timerSetValue = -1;
			}

			if (value !== this.speechElement.innerHTML)
			{
				this.speechElement.innerHTML = value;
				if (this.isLogEnabled)
					console.log("[speech]: " + value);
			}
			else
			{
				this.speechElement.innerHTML = "";

				if ("" !== value)
				{
					this.timerSetValue = setTimeout(function(){
						AscCommon.SpeechWorker.timerSetValue = -1;
						AscCommon.SpeechWorker._setValuePermanently(value);
					}, 50);
				}
			}
		};

		// 2) Set value permanently
		this._setValuePermanently = function(value)
		{
			this.speechElement.innerHTML = value;
			if (this.isLogEnabled)
				console.log("[speech]: " + value);
		};

		// 3) Resolve the problem with equals with nbsp
		this._setValuePermanentlyDiffEqual = function(value)
		{
			if (this.value !== value)
			{
				this.value = value;
				this.speechElement.innerHTML = this.value;
				this.valueEqualAddon = false;
			}
			else
			{
				this.valueEqualAddon = !this.valueEqualAddon;
				if (this.valueEqualAddon)
				{
					this.speechElement.innerHTML = this.value + this.equalAddon;
				}
				else
				{
					this.speechElement.innerHTML = this.value;
				}
			}

			if (this.isLogEnabled)
				console.log("[speech]: " + this.speechElement.innerHTML);
		};

		// 4) Simple timer
		this._setValueWithTimeout = function(value)
		{
			if (-1 !== this.timerSetValue)
			{
				clearTimeout(this.timerSetValue);
				this.timerSetValue = -1;
			}

			this.timerSetValue = setTimeout(function(){
				AscCommon.SpeechWorker._setValuePermanentlyDiffEqual(value);
				AscCommon.SpeechWorker.timerSetValue = -1;
			}, 5000);
		};

		this._setValue = this._setValuePermanentlyDiffEqual;

		this.isSpeechEnabled = function()
		{
			if (!this.isEnabled)
				return false;

			if (AscCommon.g_inputContext && AscCommon.g_inputContext.isCompositionProcess())
				return false;

			return true;
		};

		this.speech = function(type, obj)
		{
			if (!this.isSpeechEnabled())
				return;

			if (undefined === obj)
				obj = {};
			
			if (obj.cancelSelection)
				this._log("Text selection has been canceled");
		
			if (obj.moveToMainPart)
				this._log("Main document part");
			else if (obj.moveToFootnote)
				this._log("Footnote");
			else if (obj.moveToFootnote)
				this._log("Drawing");
			else if (obj.moveToHdrFtr)
				this._log("Header/Footer");
			
			if (obj.moveToStartOfDocument)
				this._log("Start of the document");
			else if (obj.moveToStartOfLine)
				this._log("Start of the line");
			else if (obj.moveToEndOfDocument)
				this._log("End of the document");
			else if (obj.moveToEndOfLine)
				this._log("End of the line");
			else if (obj.movePageUp)
				this._log("Page up");
			else if (obj.movePageDown)
				this._log("Page down");

			let translateManager = AscCommon.translateManager;
			switch (type)
			{
				case SpeechWorkerType.Text:
				case SpeechWorkerType.TextSelected:
				case SpeechWorkerType.TextUnselected:
				{
					if (obj.text === " ")
						obj.text = translateManager.getValue("space");
					break;
				}
				default:
					break;
			}

			switch (type)
			{
				case SpeechWorkerType.Text:
				{
					this._setValue(obj.text);
					break;
				}
				case SpeechWorkerType.TextSelected:
				{
					if (obj.isBefore)
						this._setValue(translateManager.getValue("select") + " " + obj.text);
					else
						this._setValue(obj.text + " " + translateManager.getValue("select"));
					break;
				}
				case SpeechWorkerType.TextUnselected:
				{
					if (obj.isBefore)
						this._setValue(translateManager.getValue("unselected") + " " + obj.text);
					else
						this._setValue(obj.text + " " + translateManager.getValue("unselected"));
					break;
				}
				case SpeechWorkerType.SlidesSelected:
				{
					if (obj.indexes.length === 1)
					{
						this._setValue(translateManager.getValue("slide") + " " + (obj.indexes[0]));
					}
					else
					{
						this._setValue(obj.indexes.length + " " + translateManager.getValue("slides added to selection"));
					}
					break;
				}
				case SpeechWorkerType.SlidesUnselected:
				{
					if (obj.indexes.length === 1)
					{
						this._setValue(translateManager.getValue("slide") + " " + (obj.indexes[0]) + " " + translateManager.getValue("unselected"));
					}
					else
					{
						this._setValue(obj.indexes.length + " " + translateManager.getValue("slides unselected"));
					}
					break;
				}
				case SpeechWorkerType.DrawingSelected:
				{
					this._setValue(translateManager.getValue("drawing select") + (obj.altText ? (" " + obj.altText) : ""));
					break;
				}
				case SpeechWorkerType.CellSelected:
				{
					this._setValue((obj.text ? obj.text : translateManager.getValue("empty cell")) + " " + obj.cell);
					break;
				}
				case SpeechWorkerType.CellRangeSelected:
				{
					let result = translateManager.getValue("selected range select") + " ";
					result += obj.start.text ? obj.start.text : translateManager.getValue("empty");
					result += (" " + obj.start.cell);
					result += obj.end.text ? " " + obj.end.text : " " + translateManager.getValue("empty");
					result += (" " + obj.end.cell);

					this._setValue(result);
					break;
				}
				case SpeechWorkerType.CellRangeUnselected:
				{
					let result = translateManager.getValue("unselected range select") + " ";
					result += obj.start.text ? obj.start.text : translateManager.getValue("empty");
					result += (" " + obj.start.cell);
					result += obj.end.text ? " " + obj.end.text : " " + translateManager.getValue("empty");
					result += (" " + obj.end.cell);

					this._setValue(result);
					break;
				}
				case SpeechWorkerType.CellRangeSelectedChangeOne:
				{
					let result = translateManager.getValue("select") + " ";
					result += obj.text ? obj.text : translateManager.getValue("empty");
					result += (" " + obj.cell);

					this._setValue(result);
					break;
				}
				case SpeechWorkerType.CellRangeUnselectedChangeOne:
				{
					let result = translateManager.getValue("unselected") + " ";
					result += obj.text ? obj.text : translateManager.getValue("empty");
					result += (" " + obj.cell);

					this._setValue(result);
					break;
				}
				case SpeechWorkerType.MultipleRangesSelected:
				{
					if (obj.ranges)
					{
						let result = translateManager.getValue("selected") + " ";
						result += obj.ranges.length + " " + translateManager.getValue("areas") + " ";

						for (let i = 0; i < obj.ranges.length; i++) {
							result += obj.ranges[i].startCell + "-" +  obj.ranges[i].endCell + " ";
						}
						result += obj.text ? obj.text : translateManager.getValue("empty");

						this._setValue(result);
					}
					break;
				}
				case SpeechWorkerType.SheetSelected:
				{
					//ms read after "objects" only selection
					//we read selection in next command
					let isEmpty = 0 === obj.cellsCount && 0 === obj.objectsCount;
					let result = "";
					if (isEmpty)
					{
						//ms not read it, read only else
						result = obj.name + " " + translateManager.getValue("empty sheet") + " "/*+ obj.cell*/;
					}
					else
					{
						result = obj.name + " " + translateManager.getValue("end of sheet") + " " + obj.cellEnd + " " +
							obj.cellsCount + " " + translateManager.getValue("cells") + " "
							obj.objectsCount + " " + translateManager.getValue("objects") /*+
							obj.text + " " + obj.cell*/;
					}
					this._setValue(result);
					break;
				}
				default:
					break;
			}
		};
	}

	window.AscCommon.SpeechWorker = new CWorkerSpeech();
	window.AscCommon.SpeechWorkerCommands = SpeechWorkerType;
	
	const SpeakerActionType = {
		unknown : 0,
		keyDown : 1,
		sheetChange : 2,
		undoRedo: 3
	};
	
	/**
	 * @constructor
	 */
	function EditorActionSpeaker()
	{
		this.speechWorker = window.AscCommon.SpeechWorker;
		this.editor = null;
		
		this.isLanched = false;
		
		this.onSelectionChange    = null;
		this.onActionStart        = null;
		this.onActionEnd          = null;
		this.onBeforeKeyDown      = null;
		this.onKeyDown            = null;
		this.onBeforeApplyChanges = null;
		this.onApplyChanges       = null;
		this.onBeforeUndoRedo     = null;
		this.onUndoRedo           = null;
		this.onCloseFile          = null;
		
		this.selectionState = null;
		this.isAction       = false;
		this.isApplyChanges = false;
		this.isKeyDown      = false;
		this.isUndoRedo     = false;

		if (AscCommon.EditorActionSpeakerInitData && AscCommon.EditorActionSpeakerInitData.isEnabled)
			this.run();
	}
	EditorActionSpeaker.prototype.toggle = function()
	{
		if (this.isLanched)
			this.stop();
		else
			this.run();
	};
	EditorActionSpeaker.prototype.run = function()
	{
		this.editor = Asc.editor;
		if (!this.editor || this.isLanched)
			return;
		
		this.isLanched = true;
		
		this.initEvents();
		this.editor.asc_registerCallback('asc_onSelectionEnd', this.onSelectionChange);
		this.editor.asc_registerCallback('asc_onCursorMove', this.onSelectionChange);
		this.editor.asc_registerCallback('asc_onUserActionStart', this.onActionStart);
		this.editor.asc_registerCallback('asc_onUserActionEnd', this.onActionEnd);
		this.editor.asc_registerCallback('asc_onBeforeKeyDown', this.onBeforeKeyDown);
		this.editor.asc_registerCallback('asc_onKeyDown', this.onKeyDown);
		this.editor.asc_registerCallback('asc_onBeforeApplyChanges', this.onBeforeApplyChanges);
		this.editor.asc_registerCallback('asc_onApplyChanges', this.onApplyChanges);
		this.editor.asc_registerCallback('asc_onBeforeUndoRedo', this.onBeforeUndoRedo);
		this.editor.asc_registerCallback('asc_onUndoRedo', this.onUndoRedo);
		this.editor.asc_registerCallback('asc_onBeforeUndoRedoInCollaboration', this.onBeforeUndoRedo);
		this.editor.asc_registerCallback('asc_onUndoRedoInCollaboration', this.onUndoRedo);
		this.editor.asc_registerCallback('asc_onCloseFile', this.onCloseFile);

		//se
		this.editor.asc_registerCallback('asc_onActiveSheetChanged', this.onActiveSheetChanged);
		
		this.selectionState = this.editor.isDocumentLoadComplete ? this.editor.getSelectionState() : null;
		this.isAction       = false;
		this.isApplyChanges = false;
		this.isKeyDown      = false;
		this.isUndoRedo     = false;
		
		this.speechWorker.setEnabled(true);
	};
	EditorActionSpeaker.prototype.stop = function()
	{
		if (!this.isLanched)
			return;
		
		this.editor.asc_unregisterCallback('asc_onSelectionEnd', this.onSelectionChange);
		this.editor.asc_unregisterCallback('asc_onCursorMove', this.onSelectionChange);
		this.editor.asc_unregisterCallback('asc_onUserActionStart', this.onActionStart);
		this.editor.asc_unregisterCallback('asc_onUserActionEnd', this.onActionEnd);
		this.editor.asc_unregisterCallback('asc_onBeforeKeyDown', this.onBeforeKeyDown);
		this.editor.asc_unregisterCallback('asc_onKeyDown', this.onKeyDown);
		this.editor.asc_unregisterCallback('asc_onBeforeApplyChanges', this.onBeforeApplyChanges);
		this.editor.asc_unregisterCallback('asc_onApplyChanges', this.onApplyChanges);
		this.editor.asc_unregisterCallback('asc_onBeforeUndoRedo', this.onBeforeUndoRedo);
		this.editor.asc_unregisterCallback('asc_onUndoRedo', this.onUndoRedo);
		this.editor.asc_unregisterCallback('asc_onBeforeUndoRedoInCollaboration', this.onBeforeUndoRedo);
		this.editor.asc_unregisterCallback('asc_onUndoRedoInCollaboration', this.onUndoRedo);
		this.editor.asc_unregisterCallback('asc_onCloseFile', this.onCloseFile);
		
		//se
		this.editor.asc_unregisterCallback('asc_onActiveSheetChanged', this.onActiveSheetChanged);
		
		this.selectionState = null;
		this.isAction       = false;
		this.isKeyDown      = false;
		this.isApplyChanges = false;
		this.isUndoRedo     = false;
		
		this.speechWorker.setEnabled(false);
		
		this.isLanched = false;
	};
	EditorActionSpeaker.prototype.initEvents = function()
	{
		let _t = this;
		
		this.onSelectionChange = function()
		{
			if (_t.isAction
				|| _t.isKeyDown
				|| _t.isApplyChanges
				|| _t.isUndoRedo)
				return;
			
			_t.handleSpeechDescription(null);
		};
		
		this.onActionStart = function()
		{
			_t.isAction = true;
		};
		
		this.onActionEnd = function()
		{
			_t.isAction = false;
			_t.updateState();
			// TODO: Если нужно, то добавить описание действия
		};
		
		this.onBeforeKeyDown = function()
		{
			_t.isKeyDown = true;
		};
		
		this.onKeyDown = function(e)
		{
			_t.isKeyDown = false;
			_t.handleSpeechDescription({type: SpeakerActionType.keyDown, event : e});
		};
		
		this.onBeforeApplyChanges = function()
		{
			_t.isApplyChanges = true;
		};
		
		this.onApplyChanges = function()
		{
			_t.isApplyChanges = false;
			_t.updateState();
			// TODO: Если дополнительно сообщить о совместке, то добавить тут
		};
		
		this.onBeforeUndoRedo = function()
		{
			_t.isUndoRedo = true;
		};
		
		this.onUndoRedo = function()
		{
			_t.isUndoRedo = false;
			_t.handleSpeechDescription({type: SpeakerActionType.undoRedo});
			_t.updateState();
			// TODO: Если дополнительно сообщить об Undo/Redo, то добавить тут
		};

		this.onActiveSheetChanged = function(index)
		{
			_t.handleSpeechDescription({type: SpeakerActionType.sheetChange, index : index});
		};
		
		this.onCloseFile = function()
		{
			_t.resetState();
		};
		
	};
	EditorActionSpeaker.prototype.handleSpeechDescription = function(action)
	{
		let state = this.editor.getSelectionState();
		if (!this.selectionState)
		{
			this.selectionState = state;
			return;
		}
		
		let speechInfo = this.editor.getSpeechDescription(this.selectionState, action);
		this.selectionState = state;
		if (!speechInfo)
			return;
		
		this.speechWorker.speech(speechInfo.type, speechInfo.obj);
	};
	EditorActionSpeaker.prototype.updateState = function()
	{
		this.selectionState = this.editor.getSelectionState();
	};
	EditorActionSpeaker.prototype.resetState = function()
	{
		this.selectionState = null;
	};
	
	window.AscCommon.EditorActionSpeaker = new EditorActionSpeaker();
	window.AscCommon.SpeakerActionType = SpeakerActionType;

})(window);

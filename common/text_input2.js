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

(function(window, undefined)
{
	window["AscInputMethod"] = window["AscInputMethod"] || {};
	///
	// такие методы нужны в апи
	// baseEditorsApi.prototype.Begin_CompositeInput = function()
	// baseEditorsApi.prototype.Replace_CompositeText = function(arrCharCodes)
	// baseEditorsApi.prototype.Set_CursorPosInCompositeText = function(nPos)
	// baseEditorsApi.prototype.Get_CursorPosInCompositeText = function()
	// baseEditorsApi.prototype.End_CompositeInput = function()
	// baseEditorsApi.prototype.Get_MaxCursorPosInCompositeText = function()

	// baseEditorsApi.prototype.onKeyDown = function(e)
	// baseEditorsApi.prototype.onKeyPress = function(e)
	// baseEditorsApi.prototype.onKeyUp = function(e)
	///

	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].inputMethodAddInitEvent = function(callback)
	{
		AscCommon.g_inputContext_create_events = AscCommon.g_inputContext_create_events || [];
		AscCommon.g_inputContext_create_events.push(callback);
	};
	window['AscCommon'].inputMethodCheckInitEvents = function()
	{
		if (!AscCommon.g_inputContext_create_events)
			return;

		for (let i = 0, len = AscCommon.g_inputContext_create_events.length; i < len; i++)
			AscCommon.g_inputContext_create_events[i]();

		delete AscCommon.g_inputContext_create_events;
	};

	var InputTextElementType = {
		TextArea           : 0,
		ContentEditableDiv : 1
	};

	function CTextInput2(api)
	{
		this.Api = api;

		this.TargetId = null; // id caret
		this.HtmlDiv  = null; // для незаметной реализации одной textarea недостаточно. parent для HtmlArea
		this.HtmlArea = null; // HtmlArea - элемент для ввода
		this.ElementType = InputTextElementType.TextArea;

		// ---------------------------------------------------------------
		// chrome element for left/top
		this.FixedPosCheckElementX = 0;
		this.FixedPosCheckElementY = 0;
		// Notes offset for slides
		this.TargetOffsetY = 0;

		this.HtmlAreaOffset = 50; // height in pix
		this.HtmlAreaWidth = 200;
		// ---------------------------------------------------------------

		// информация о текущем состоянии текста -------------------------

		// текущее значение в textarea
		this.Text = "";

		// текст до того, как пришли сообщения onCompositeStart/onCompositeUpdate
		// т.е. текст, который пришел на onInput/onTextInput, и когда мы не внутри onComposite[Begin-End]
		this.TextBeforeComposition = "";

		// в каком состоянии апи (композитный ли ввод сейчас)
		this.IsComposition = false;

		// ---------------------------------------------------------------

		// не обрабатывать keyPress после keyDown
		this.IsDisableKeyPress = false;

		this.nativeFocusElement = null;
		this.nativeFocusElementNoRemoveOnElementFocus = false;
		this.InterfaceEnableKeyEvents = true;
		this.isNoClearOnFocus = false;
		this.isGlobalDisableFocus = false;

		this.ReadOnlyCounter = 0;

		this.keyPressInput = "";
		this.isInputHelpersPresent = false;
		this.isInputHelpers = {};

		// параметры для показа/скрытия виртуальной клавиатуры.
		this.isHardCheckKeyboard = AscCommon.AscBrowser.isSailfish;

		this.virtualKeyboardClickTimeout = -1;
		this.virtualKeyboardReadOnly_ShowKeyboard = AscCommon.AscBrowser.isAndroid && AscCommon.AscBrowser.isMozilla;

		// для сброса текста при фокусе
		this.checkClearTextOnFocusTimerId = -1;

		this.isDisableKeyboard = false;

		this.moveAccurateInfo = {
			id : -1,
			x : 0,
			y : 0
		};
	}

	var CTextInputPrototype = CTextInput2.prototype;

	const TEXT_INPUT_DEBUG = false;
	CTextInputPrototype.log = function(value)
	{
		if (TEXT_INPUT_DEBUG)
			console.log(value);
	};

	// для совместимости. убрал системный ввод
	CTextInputPrototype.systemInputEnable = function()
	{
	};

	// input common
	CTextInputPrototype.isSpaceSymbol = function(e)
	{
		if (e.keyCode == 32)
			return true;

		if ((e.keyCode == 229) && ((e.code == "space") || (e.code == "Space") || (e.key == "Spacebar")))
			return true;

		return false;
	};
	CTextInputPrototype.isCompositionProcess = function()
	{
		return this.IsComposition;
	};

	// input
	CTextInputPrototype.onKeyDown = function(e)
	{
		if (this.Api.isLongAction())
		{
			AscCommon.stopEvent(e);
			return false;
		}

		// проверим - может это навигация в окне хэлпера
		if (this.isInputHelpersPresent)
		{
			switch (e.keyCode)
			{
				case 9:		// tab
				case 13:	// enter
				case 38:	// top
				case 40:	// bottom
				case 33: 	// pageup
				case 34: 	// pagedown
				case 35: 	// end
				case 36: 	// home
				case 27:	// escape
				{
					window.g_asc_plugins.onPluginEvent2("onKeyDown", { "keyCode" : e.keyCode }, this.isInputHelpers);

					AscCommon.stopEvent(e);
					return false;
				}
				default:
					break;
			}
		}

		if (null != this.nativeFocusElement)
		{
			if (this.emulateNativeKeyDown(e))
			{
				e.preventDefault();
				return false;
			}
		}

		AscCommon.check_KeyboardEvent(e);
		var arrCodes = this.Api.getAddedTextOnKeyDown(AscCommon.global_keyboardEvent);

		var isAsync = AscFonts.FontPickerByCharacter.checkTextLight(arrCodes, true);

		if (isAsync)
		{
			AscFonts.FontPickerByCharacter.loadFonts(this, function ()
			{
				this.onKeyDown(e);
				this.onKeyUp(e);

				this.setReadOnly(false);
			});

			this.setReadOnly(true);
			AscCommon.stopEvent(e);
			return false;
		}

		let isSpaceAsText = (32 === e.keyCode);
		if (isSpaceAsText)
		{
			// hotkeys
			if (AscCommon.global_keyboardEvent.AltKey ||
				AscCommon.global_keyboardEvent.CtrlKey ||
				AscCommon.global_keyboardEvent.MacCmdKey)
			{
				isSpaceAsText = false;
			}

			if (isSpaceAsText)
			{
				switch (this.Api.editorId)
				{
					case AscCommon.c_oEditorId.Spreadsheet:
					{
						if (AscCommon.global_keyboardEvent.ShiftKey)
							isSpaceAsText = false;
						break;
					}
					case AscCommon.c_oEditorId.Presentation:
					{
						if (this.Api.WordControl && this.Api.WordControl.DemonstrationManager && this.Api.WordControl.DemonstrationManager.Mode)
							isSpaceAsText = false;
						break;
					}
					default:
						break;
				}
			}
		}

		// ios копирование и вырезка через клавиатуру внешнюю - требует селекта в фокусном textarea
		// но если селектить - его видно. да и куча проблем. попробуем сэмулировать
		if (this.Api.isMobileVersion && AscCommon.AscBrowser.isAppleDevices)
		{
			if (e.metaKey)
			{
				if (e.keyCode === 67)
				{
					AscCommon.g_clipboardBase.Button_Copy();
					return;
				}
				else if (e.keyCode === 88)
				{
					AscCommon.g_clipboardBase.Button_Cut();
					return;
				}
			}
			else if (e.ctrlKey)
			{
				// safari send code 13 on ctrl + c. disable it
				if (e.keyCode === 13 && e.code === "KeyC")
					return;
			}
		}

		let ret = undefined;
		if (!isSpaceAsText)
			ret = this.Api.onKeyDown(e);

		let isSpecialClearInComposition = true;
		switch (e.keyCode)
		{
			case 8:		// backspace
				isSpecialClearInComposition = false;
			case 9:		// tab
			case 13:	// enter
			case 37:	// left
			case 38:	// top
			case 39:	// right
			case 40:	// bottom
			case 33: 	// pageup
			case 34: 	// pagedown
			case 35: 	// end
			case 36: 	// home
			case 46:	// delete
			{
				if (!this.IsComposition || isSpecialClearInComposition)
					this.clear();
			}
			default:
				break;
		}

		if (e.keyCode === 32 && AscCommon.global_keyboardEvent.CtrlKey && !AscCommon.global_keyboardEvent.ShiftKey)
		{
			if (window.g_asc_plugins)
				window.g_asc_plugins.onPluginEvent("onClick");
		}

		return ret;
	};
	CTextInputPrototype.onKeyPress = function(e)
	{
		if (this.Api.isLongAction() || !this.Api.asc_IsFocus() || this.Api.isViewMode)
		{
			AscCommon.stopEvent(e);
			return false;
		}

		// вся обработка - в onInput
	};
	CTextInputPrototype.onKeyUp = function(e)
	{
		if (this.Api.isLongAction())
		{
			AscCommon.stopEvent(e);
			return false;
		}

		AscCommon.global_keyboardEvent.Up();
		this.Api.onKeyUp(e);
	};

	CTextInputPrototype.onFocusInputText = function()
	{
		this.onFocusInputTextEnd();

		this.checkClearTextOnFocusTimerId = setTimeout(function(){
			let _t = AscCommon.g_inputContext;
			if (!_t.IsComposition)
				_t.clear(true);
		}, 500);
	};

	CTextInputPrototype.onFocusInputTextEnd = function()
	{
		if (-1 !== this.checkClearTextOnFocusTimerId)
		{
			clearTimeout(this.checkClearTextOnFocusTimerId);
			this.checkClearTextOnFocusTimerId = -1;
		}
	};

	CTextInputPrototype.onInput = function(e)
	{
		if (this.Api.isLongAction())
		{
			AscCommon.stopEvent(e);
			return false;
		}

		this.onFocusInputTextEnd();

		let type = (e.type ? ("" + e.type) : "undefined");
		type = type.toLowerCase()

		let newValue = this.getAreaValue();
		this.log("onInput: " + newValue);

		if (-1 !== newValue.indexOf("&nbsp;"))
			newValue = newValue.split("&nbsp;").join(" ");

		if (("compositionstart" === type) && this.IsComposition)
		{
			// не пришел end - пришлем сами
			this.compositeEnd();
		}

		if (("compositionstart" === type || "compositionupdate" === type) && !this.IsComposition)
		{
			// начался композитный ввод
			this.TextBeforeComposition = this.Text;

			this.log("compositionStart: " + this.TextBeforeComposition);
			this.compositeStart();
		}

		let lastSymbol = 0;
		let newTextLength = 0;

		let isAsyncInput = false;
		if (this.IsComposition)
		{
			if (newValue.length >= this.TextBeforeComposition.length)
			{
				let newText = newValue.substr(this.TextBeforeComposition.length);

				this.log("compositionText: " + newText);

				let codes = [];
				for (let iter = newText.getUnicodeIterator(); iter.check(); iter.next())
					codes.push(iter.value());

				newTextLength = codes.length;
				if (newTextLength > 0)
					lastSymbol = codes[newTextLength - 1];

				isAsyncInput = this.checkTextInput(codes);
			}
		}
		else
		{
			// текст может не только добавиться, но и замениться (например на маке зажать i - и выбрать вариант)
			let codesOld = [];
			for (let iter = this.Text.getUnicodeIterator(); iter.check(); iter.next())
				codesOld.push(iter.value());

			let codesNew = [];
			for (let iter = newValue.getUnicodeIterator(); iter.check(); iter.next())
				codesNew.push(iter.value());

			let oldLen = codesOld.length;
			let newLen = codesNew.length;
			let savedLen = (oldLen < newLen) ? oldLen : newLen;
			let equalsLen = 0;

			for (let i = 0; i < savedLen; i++)
			{
				if (codesOld[i] !== codesNew[i])
					break;
				++equalsLen;
			}

			newTextLength = newLen;

			// удаляем то, чего уже нет
			let codesRemove = undefined;
			if (oldLen > equalsLen)
				codesRemove = codesOld.slice(equalsLen);

			// удаляем старые из массива
			if (0 !== equalsLen)
				codesNew.splice(0, equalsLen);

			if (codesNew.length > 0)
				lastSymbol = codesNew[codesNew.length - 1];

			if (10 === lastSymbol)
			{
				// заглушка на интерфейс (если там enter был нажат - и сначала blur(), и только затем применение).
				this.clear();
				return;
			}

			// добавляем новые
			isAsyncInput = this.checkTextInput(codesNew, codesRemove);
		}

		if (("compositionend" === type) && this.IsComposition)
		{
			// закончился композитный ввод
			this.compositeEnd();

			this.log("compositionEnd: " + newValue);
		}

		if (!isAsyncInput)
		{
			// если асинхронно - то на коллбеке придет onInput - и текст добавится позже
			this.Text = newValue;
		}

		if (window.g_asc_plugins)
			window.g_asc_plugins.onPluginEvent("onInputHelperInput", { "text" : this.Text });

		if (!this.IsComposition && lastSymbol !== 0 && !isAsyncInput)
		{
			let isClear = false;
			switch (lastSymbol)
			{
				case 32: // пробел
				case 46: // точка
				case 44: // запятая
				//case 12290: // азиатская точка
				//case 65292: // азиатская запятая
				{
					isClear = true;
					break;
				}
				default:
				{
					// надеемся, что при вводе все-таки будут точки/пробелы/запятые
					// если нет - то не даем копить до бесконечности.
					let currentTextLenMax = this.Api.isMobileVersion ? 20 : 100;
					if (newTextLength > currentTextLenMax)
						isClear = true;
					break;
				}
			}
			if (isClear)
				this.clear();
		}
	};
	CTextInputPrototype.addText = function(text)
	{
		this.setAreaValue(this.getAreaValue() + text);

		this.onInput({
			type : "input",
			preventDefault : function() {},
			stopPropagation : function() {}
		});
	};
	CTextInputPrototype.compositeStart = function()
	{
		if (this.IsComposition)
			return;

		this.IsComposition = true;
		this.Api.Begin_CompositeInput();
	};
	CTextInputPrototype.compositeReplace = function(codes)
	{
		this.Api.Replace_CompositeText(codes);
	};
	CTextInputPrototype.compositeEnd = function()
	{
		if (!this.IsComposition)
			return;

		this.IsComposition = false;
		this.Api.End_CompositeInput();

		this.TextBeforeComposition = "";
	};
	CTextInputPrototype.apiCompositeEnd = function()
	{
		if (!this.IsComposition)
			return;

		this.compositeEnd();
		this.clear();
	};
	CTextInputPrototype.checkTextInput = function(codes, codesRemove)
	{
		var isAsync = AscFonts.FontPickerByCharacter.checkTextLight(codes, true);

		if (!isAsync)
		{
			window.LOCK_DRAW = true;

			if (this.IsComposition)
			{
				this.compositeReplace(codes);
			}
			else
			{
				this.addTextCodes(codes, codesRemove);
			}
		}
		else
		{
			AscFonts.FontPickerByCharacter.loadFonts(this, function ()
			{
				this.onInput({
					type : this.IsComposition ? "compositionupdate" : "input",
					preventDefault : function() {},
					stopPropagation : function() {}
				});

				//this.setReadOnly(false);
			});

			//this.setReadOnly(true);
		}
		return isAsync;
	};

	CTextInputPrototype.addTextCodes = function(codes, codesRemove)
	{
		if (codesRemove && codesRemove.length !== 0)
		{
			// old version (cells??).
			//this.removeText(codesRemove.length);

			let resultCorrection = this.Api.asc_correctEnterText(codesRemove, codes);
			if (true !== resultCorrection)
				this.Api.asc_enterText(codes);
		}
		else
		{
			this.Api.asc_enterText(codes);
		}
	};

	/* Old version
	CTextInputPrototype.addTextCodes = function(codes)
	{
		for (let i = 0, len = codes.length; i < len; i++)
		{
			this.addTextCode(codes[i]);
		}
	};
	CTextInputPrototype.addTextCode = function(code)
	{
		if (code === 32)
		{
			if (!this.isSpaceOnKeyDown)
			{
				// иначе пробел добавился на onKeyDown
				let keyObject = this.getKeyboardEventObject(code);
				this.Api.onKeyDown(keyObject);
				this.Api.onKeyUp(keyObject);
			}
			return;
		}
		else
		{
			// TODO: отдельный метод в апи
			// пока имитируем через keyCode - для keyDown/Up - сделаем такой код,
			// который ни на что не влияет. код для буквы 'a' - 65
			let keyObject = this.getKeyboardEventObject(code);
			let keyObjectUpDown = this.getKeyboardEventObject(65);

			this.Api.onKeyDown(keyObjectUpDown);
			this.Api.onKeyPress(keyObject);
			this.Api.onKeyUp(keyObjectUpDown);
		}
	};
	*/
	CTextInputPrototype.removeText = function(length)
	{
		for (let i = 0; i < length; i++)
		{
			// backspace
			let keyObject = this.getKeyboardEventObject(8);
			this.Api.onKeyDown(keyObject);
			this.Api.onKeyUp(keyObject);
		}
	};
	CTextInputPrototype.emulateKeyDownApi = function(code)
	{
		let keyObject = this.getKeyboardEventObject(code);

		this.Api.onKeyDown(keyObject);
		this.Api.onKeyUp(keyObject);
	};

	// keyboard
	CTextInputPrototype.getKeyboardEventObject = function(code)
	{
		return {
			altKey : false,
			ctrlKey : false,
			shiftKey : false,
			target : null,
			charCode : 0,
			which : code,
			keyCode : code,
			code : "",
			emulated: true,

			preventDefault : function() {},
			stopPropagation : function() {}
		};
	};
	CTextInputPrototype.emulateNativeKeyDown = function(e, target)
	{
		var oEvent = document.createEvent('KeyboardEvent');

		/*
		 var _event = new KeyboardEvent("keydown", {
		 bubbles : true,
		 cancelable : true,
		 char : e.charCode,
		 shiftKey : e.shiftKey,
		 ctrlKey : e.ctrlKey,
		 metaKey : e.metaKey,
		 altKey : e.altKey,
		 keyCode : e.keyCode,
		 which : e.which,
		 key : e.key
		 });
		 */

		// Chromium Hack
		Object.defineProperty(oEvent, 'keyCode', {
			get : function()
			{
				return this.keyCodeVal;
			}
		});
		Object.defineProperty(oEvent, 'which', {
			get : function()
			{
				return this.keyCodeVal;
			}
		});
		Object.defineProperty(oEvent, 'shiftKey', {
			get : function()
			{
				return this.shiftKeyVal;
			}
		});
		Object.defineProperty(oEvent, 'altKey', {
			get : function()
			{
				return this.altKeyVal;
			}
		});
		Object.defineProperty(oEvent, 'metaKey', {
			get : function()
			{
				return this.metaKeyVal;
			}
		});
		Object.defineProperty(oEvent, 'ctrlKey', {
			get : function()
			{
				return this.ctrlKeyVal;
			}
		});

		if (AscCommon.AscBrowser.isIE)
		{
			oEvent.preventDefault = function ()
			{
				try
				{
					Object.defineProperty(this, "defaultPrevented", {
						get: function ()
						{
							return true;
						}
					});
				}
				catch(err)
				{
				}
			};
		}

		var k = e.keyCode;
		if (oEvent.initKeyboardEvent)
		{
			oEvent.initKeyboardEvent("keydown", true, true, window, false, false, false, false, k, k);
		}
		else
		{
			oEvent.initKeyEvent("keydown", true, true, window, false, false, false, false, k, 0);
		}

		oEvent.keyCodeVal = k;
		oEvent.shiftKeyVal = e.shiftKey;
		oEvent.altKeyVal = e.altKey;
		oEvent.metaKeyVal = e.metaKey;
		oEvent.ctrlKeyVal = e.ctrlKey;

		var _elem = target ? target : _getElementKeyboardDown(this.nativeFocusElement, 3);
		_elem.dispatchEvent(oEvent);

		return oEvent.defaultPrevented;
	};

	//
	CTextInputPrototype.getAreaPos = function()
	{
		var _offset = 0;
		if (this.ElementType === InputTextElementType.TextArea)
		{
			_offset = this.HtmlArea.selectionEnd;
		}
		else
		{
			var sel = window.getSelection();
			if (sel.rangeCount > 0)
			{
				var range = sel.getRangeAt(0);
				_offset = range.endOffset;
			}
		}
		return _offset;
	};
	CTextInputPrototype.checkTargetPosition = function(isCorrect)
	{
		var _offset = this.getAreaPos();

		if (false !== isCorrect)
		{
			var _value = this.getAreaValue();
			_offset -= (_value.length - this.compositionValue.length);
		}

		if (!this.IsLockTargetMode)
		{
			// никакого смысла прыгать курсором туда-сюда
			if (_offset == 0 && this.compositionValue.length == 1)
				_offset = 1;
		}

		this.Api.Set_CursorPosInCompositeText(_offset);

		this.unlockTarget();
	};
	CTextInputPrototype.clear = function(isFromFocus)
	{
		this.log("clear");

		this.TextBeforeComposition = "";
		this.Text = "";

		this.compositeEnd();
		this.clearAreaValue();

		if (isFromFocus !== true)
			focusHtmlElement(this.HtmlArea);

		if (window.g_asc_plugins)
			window.g_asc_plugins.onPluginEvent("onInputHelperClear");
	};
	CTextInputPrototype.getAreaValue = function()
	{
		return (this.ElementType === InputTextElementType.TextArea) ? this.HtmlArea.value : this.HtmlArea.innerText;
	};
	CTextInputPrototype.clearAreaValue = function()
	{
		if (this.ElementType === InputTextElementType.TextArea)
			this.HtmlArea.value = "";
		else
			this.HtmlArea.innerHTML = "";
	};
	CTextInputPrototype.setAreaValue = function(value)
	{
		if (this.ElementType === InputTextElementType.TextArea)
			this.HtmlArea.value = value;
		else
			this.HtmlArea.innerHTML = value;
	};
	CTextInputPrototype.setReadOnly = function(isLock)
	{
		if (isLock)
			this.ReadOnlyCounter++;
		else
			this.ReadOnlyCounter--;

		// при синхронной загрузке шрифтов (десктоп)
		// может вызываться и в обратном порядке (setReadOnly(false), setReadOnly(true))
		// поэтому сравнение с нулем неверно. отрицательные значение могут быть.

		this.setReadOnlyWrapper((0 >= this.ReadOnlyCounter) ? false : true);
	};
	CTextInputPrototype.setReadOnlyWrapper = function(val)
	{
		this.HtmlArea.readOnly = this.isDisableKeyboard ? true : val;
	};
	CTextInputPrototype.setInterfaceEnableKeyEvents = function(value)
	{
		this.InterfaceEnableKeyEvents = value;
		if (true === this.InterfaceEnableKeyEvents)
		{
			if (document.activeElement)
			{
				var _id = document.activeElement.id;
				if (_id == "area_id" || (window.g_asc_plugins && window.g_asc_plugins.checkRunnedFrameId(_id)))
					return;
			}

			if (!this.isGlobalDisableFocus)
				focusHtmlElement(this.HtmlArea);
		}
	};
	CTextInputPrototype.externalEndCompositeInput = function()
	{
		this.clear();
	};
	CTextInputPrototype.externalChangeFocus = function()
	{
		return;
		if (!this.IsComposition)
			return false;

		setTimeout(function() {
			window['AscCommon'].g_inputContext.clear();
		}, 10);

		return true;
	};

	// html element
	CTextInputPrototype.init = function(target_id, parent_id)
	{
		this.TargetId   = target_id;

		this.HtmlDiv                  = document.createElement("div");
		this.HtmlDiv.id               = "area_id_parent";
		this.HtmlDiv.style.background = "transparent";
		this.HtmlDiv.style.border     = "none";

		// в хроме скроллируется редактор, когда курсор текстового поля выходит за пределы окна
		if (AscCommon.AscBrowser.isChrome && !TEXT_INPUT_DEBUG)
			this.HtmlDiv.style.position = "fixed";
		else
			this.HtmlDiv.style.position   = "absolute";
		this.HtmlDiv.style.zIndex     = 10;
		this.HtmlDiv.style.width      = TEXT_INPUT_DEBUG ? "200px" : "20px";
		this.HtmlDiv.style.height     = "50px";
		this.HtmlDiv.style.overflow   = "hidden";

		this.HtmlDiv.style.boxSizing 		= "content-box";
		this.HtmlDiv.style.webkitBoxSizing 	= "content-box";
		this.HtmlDiv.style.MozBoxSizing 	= "content-box";

		if (this.ElementType === InputTextElementType.TextArea)
		{
			this.HtmlArea = document.createElement("textarea");
		}
		else
		{
			this.HtmlArea = document.createElement("div");
			this.HtmlArea.setAttribute("contentEditable", true);
		}
		this.HtmlArea.id = "area_id";

		if (this.Api.isViewMode && this.Api.isMobileVersion)
			this.setReadOnlyWrapper(true);

		var _style = "";
		if (!TEXT_INPUT_DEBUG)
		{
			_style = ("left:-" + (this.HtmlAreaWidth >> 1) + "px;top:" + (-this.HtmlAreaOffset) + "px;");
			_style += "color:transparent;caret-color:transparent;background:transparent;";

			if (this.Api.isUseOldMobileVersion())
				_style += (AscCommon.AscBrowser.isAppleDevices && !AscCommon.AscBrowser.isTelegramWebView && (AscCommon.AscBrowser.maxTouchPoints > 0)) ? "font-size:0px;" : "font-size:8px;";
			else
				_style += "font-size:8px;";
		}
		else
		{
			_style = "left:0px;top:0px;color:black;caret-color:black;font-size:16px;background:transparent;";
		}
		_style += ("border:none;position:absolute;text-shadow:0 0 0 #000;outline:none;width:" + this.HtmlAreaWidth + "px;height:50px;");
		_style += "overflow:hidden;padding:0px;margin:0px;font-family:arial;resize:none;font-weight:normal;box-sizing:content-box;-moz-box-sizing:content-box;-webkit-box-sizing:content-box;";
		_style += "touch-action: none;-webkit-touch-callout: none;";

		this.HtmlArea.setAttribute("style", _style);
		this.HtmlArea.setAttribute("spellcheck", false);

		this.HtmlArea.setAttribute("autocapitalize", "none");
		if(AscCommon.AscBrowser.isChrome)
		{
			//Bug in Chrome. Autofill does not respect autocomplete="off"
			//https://bugs.chromium.org/p/chromium/issues/detail?id=914451
			this.HtmlArea.setAttribute("autocomplete", "extremely_off");
		}
		else
		{
			this.HtmlArea.setAttribute("autocomplete", "off");
		}
		this.HtmlArea.setAttribute("autocorrect", "off");

		this.HtmlDiv.appendChild(this.HtmlArea);

		this.appendInputToCanvas(parent_id);

		// events:
		var oThis                   = this;
		this.HtmlArea["onkeydown"]  = function(e)
		{
			if (AscCommon.AscBrowser.isSafariMacOs)
			{
				var cmdButton = (e.ctrlKey || e.metaKey) ? true : false;
				var buttonCode = ((e.keyCode == 67) || (e.keyCode == 88) || (e.keyCode == 86));
				if (cmdButton && buttonCode)
					oThis.IsDisableKeyPress = true;
				else
					oThis.IsDisableKeyPress = false;
			}
			return oThis.onKeyDown(e);
		};
		this.HtmlArea["onkeypress"] = function(e)
		{
			if (oThis.IsDisableKeyPress == true)
			{
				// macOS Sierra send keypress before copy event
				oThis.IsDisableKeyPress = false;
				var cmdButton = (e.ctrlKey || e.metaKey) ? true : false;
				if (cmdButton)
					return;
			}
			return oThis.onKeyPress(e);
		};
		this.HtmlArea["onkeyup"]    = function(e)
		{
			oThis.IsDisableKeyPress = false;
			return oThis.onKeyUp(e);
		};

		var inputEvents = ["input", /*"textInput", */"compositionstart", "compositionupdate", "compositionend"];
		for (let i = 0, len = inputEvents.length; i < len; i++)
		{
			this.HtmlArea.addEventListener(inputEvents[i], function(e)
			{
				return oThis.onInput(e);
			}, false);
		}

		this.Api.Input_UpdatePos();

		this.checkViewMode();
	};
	CTextInputPrototype.appendInputToCanvas = function(parent_id)
	{
		let oHtmlParent;
		if (undefined === parent_id)
			oHtmlParent = document.getElementById(this.TargetId).parentNode;
		else
			oHtmlParent = document.getElementById(parent_id);

		// нужен еще один родитель. чтобы скроллился он, а не oHtmlParent
		var oHtmlDivScrollable = document.createElement("div");
		oHtmlDivScrollable.id = "area_id_main";
		let styleZIndex = TEXT_INPUT_DEBUG ? "z-index:50;" : "z-index:0;";
		oHtmlDivScrollable.setAttribute("style", "background:transparent;border:none;position:absolute;padding:0px;margin:0px;pointer-events:none;" + styleZIndex);
		var parentStyle = getComputedStyle(oHtmlParent);
		oHtmlDivScrollable.style.left = parentStyle.left;
		oHtmlDivScrollable.style.top = parentStyle.top;
		oHtmlDivScrollable.style.width = parentStyle.width;
		oHtmlDivScrollable.style.height = parentStyle.height;
		oHtmlDivScrollable.style.overflow = "hidden";
		oHtmlDivScrollable.appendChild(this.HtmlDiv);
		oHtmlParent.parentNode.appendChild(oHtmlDivScrollable);
	};
	CTextInputPrototype.onResize = function(editorContainerId)
	{
		var _elem    = document.getElementById("area_id_main");
		var _elemSrc = document.getElementById(editorContainerId);

		if (!_elem || !_elemSrc)
			return;

		if (AscCommon.AscBrowser.isChrome)
		{
			var rectObject = AscCommon.UI.getBoundingClientRect(_elemSrc);
			this.FixedPosCheckElementX = rectObject.left;
			this.FixedPosCheckElementY = rectObject.top;
		}

		var _width = _elemSrc.style.width;
		if ((null == _width || "" == _width) && window.getComputedStyle)
		{
			var _s = window.getComputedStyle(_elemSrc);
			_elem.style.left   = _s.left;
			_elem.style.top    = _s.top;
			_elem.style.width  = _s.width;
			_elem.style.height = _s.height;
		}
		else
		{
			_elem.style.left   = _elemSrc.style.left;
			_elem.style.top    = _elemSrc.style.top;
			_elem.style.width  = _width;
			_elem.style.height = _elemSrc.style.height;
		}

		if (this.Api.isUseOldMobileVersion())
		{
			var _elem1 = document.getElementById("area_id_parent");
			var _elem2 = document.getElementById("area_id");

			_elem1.parentNode.style.pointerEvents = "";


			_elem1.style.left = "0px";
			let topStyle = "-1000px";

			if (AscCommon.AscBrowser.isTelegramWebView)
			{
				if (!AscCommon.AscBrowser.isAndroid && !AscCommon.AscBrowser.isAppleDevices)
					topStyle = "0px";
				else if (AscCommon.AscBrowser.isAppleDevices && navigator.maxTouchPoints === 0)
					topStyle = "0px";
			}

			_elem1.style.top = topStyle;
			_elem1.style.right = "0px";
			_elem1.style.bottom = "-100px";
			_elem1.style.width = "auto";
			_elem1.style.height = "auto";

			_elem2.style.left = "0px";
			_elem2.style.top = "0px";
			_elem2.style.right = "0px";
			_elem2.style.bottom = "0px";
			_elem2.style.width = "100%";
			_elem2.style.height = "100%";

			if (AscCommon.AscBrowser.isIE)
			{
				document.body.style["msTouchAction"] = "none";
				document.body.style["touchAction"] = "none";
			}
		}

		var _editorSdk = document.getElementById("editor_sdk");
		this.editorSdkW = _editorSdk.clientWidth;
		this.editorSdkH = _editorSdk.clientHeight;
	};
	CTextInputPrototype.checkFocus = function()
	{
		if (this.Api.asc_IsFocus() && !AscCommon.g_clipboardBase.IsFocus() && !AscCommon.g_clipboardBase.IsWorking())
		{
			if (document.activeElement != this.HtmlArea)
				focusHtmlElement(this.HtmlArea);
		}
	};
	CTextInputPrototype.moveAccurate = function(x, y)
	{
		if (!this.moveAccurateFunc)
		{
			this.moveAccurateFunc = function() {
				let ctx = AscCommon.g_inputContext;
				ctx.move(ctx.moveAccurateInfo.x, ctx.moveAccurateInfo.y);
				ctx.moveAccurateInfo.id = -1;
			};
		}

		if (-1 !== this.moveAccurateInfo.id)
			clearTimeout(this.moveAccurateInfo.id);

		this.moveAccurateInfo.x = x;
		this.moveAccurateInfo.y = y;
		this.moveAccurateInfo.id = setTimeout(this.moveAccurateFunc, 20);
	};
	CTextInputPrototype.move = function(x, y)
	{
		if (this.Api.isUseOldMobileVersion())
			return;

		var oTarget = document.getElementById(this.TargetId);
		if (!oTarget)
			return;

		var xPos = x ? x : parseInt(oTarget.style.left);
		var yPos = (y ? y : parseInt(oTarget.style.top)) + parseInt(oTarget.style.height);

		if (AscCommon.AscBrowser.isSafari && AscCommon.AscBrowser.isMobile)
			xPos = -100;

		if (this.Api.editorId === AscCommon.c_oEditorId.Presentation)
		{
			let offset = this.Api.getTextInputOffset();
			xPos += offset.X;
			yPos += offset.Y;
		}

		this.HtmlDiv.style.left = xPos + this.FixedPosCheckElementX + "px";
		this.HtmlDiv.style.top  = yPos + this.FixedPosCheckElementY + this.TargetOffsetY + this.HtmlAreaOffset + "px";

		this.HtmlArea.scrollTop = this.HtmlArea.scrollHeight;
		//this.log("" + this.HtmlArea.scrollTop + ", " + this.HtmlArea.scrollHeight);

		if (window.g_asc_plugins)
			window.g_asc_plugins.onPluginEvent("onTargetPositionChanged");
	};

	// virtual keyboard
	CTextInputPrototype.preventVirtualKeyboard = function(e)
	{
		if (this.isHardCheckKeyboard)
			return;
		//AscCommon.stopEvent(e);

		if (AscCommon.AscBrowser.isAndroid)
		{
			this.setReadOnlyWrapper(true);

			if (this.virtualKeyboardReadOnly_ShowKeyboard)
				return;

			this.virtualKeyboardClickTimeout = setTimeout(function ()
			{
				window['AscCommon'].g_inputContext.setReadOnlyWrapper(false);
				window['AscCommon'].g_inputContext.virtualKeyboardClickTimeout = -1;
			}, 1);
		}
	};
	CTextInputPrototype.enableVirtualKeyboard = function()
	{
		if (this.isHardCheckKeyboard)
			return;

		if (AscCommon.AscBrowser.isAndroid)
		{
			if (this.virtualKeyboardReadOnly_ShowKeyboard)
				return;

			if (-1 != this.virtualKeyboardClickTimeout)
			{
				clearTimeout(this.virtualKeyboardClickTimeout);
				this.virtualKeyboardClickTimeout = -1;
			}

			this.setReadOnlyWrapper(false);
		}
	};
	CTextInputPrototype.preventVirtualKeyboard_Hard = function()
	{
		this.setReadOnlyWrapper(true);
	};
	CTextInputPrototype.enableVirtualKeyboard_Hard = function()
	{
		this.setReadOnlyWrapper(false);
	};

	CTextInputPrototype.showKeyboard = function()
	{
		if (this.virtualKeyboardReadOnly_ShowKeyboard)
		{
			if (this.HtmlArea.readOnly === true)
			{
				this.setReadOnlyWrapper(false);
			}
		}

		if (!this.Api.asc_IsFocus())
			this.Api.asc_enableKeyEvents(true);
		else
			focusHtmlElement(this.HtmlArea);
	}

	CTextInputPrototype.checkViewMode = function()
	{
		let oldDisableKeyboard = this.isDisableKeyboard;
		this.isDisableKeyboard = this.Api.isViewMode;

		if (!this.isDisableKeyboard)
		{
			switch (this.Api.editorId)
			{
				case AscCommon.c_oEditorId.Word:
				{
					// use canEnterText instead this
					break;
				}
				case AscCommon.c_oEditorId.Presentation:
				case AscCommon.c_oEditorId.Spreadsheet:
				{
					if (this.Api.isRestrictionView() && !this.Api.isRestrictionForms())
						this.isDisableKeyboard = true;
					break;
				}
				default:
					break;
			}
		}

		if (oldDisableKeyboard !== this.isDisableKeyboard)
		{
			this.setReadOnlyWrapper(false);
		}
	};

	function _getAttirbute(_elem, _attr, _depth)
	{
		var _elemTest = _elem;
		for (var _level = 0; _elemTest && (_level < _depth); ++_level, _elemTest = _elemTest.parentNode)
		{
			var _res = _elemTest.getAttribute ? _elemTest.getAttribute(_attr) : null;
			if (null != _res)
				return _res;
		}
		return null;
	}
	function _getElementKeyboardDown(_elem, _depth)
	{
		var _elemTest = _elem;
		for (var _level = 0; _elemTest && (_level < _depth); ++_level, _elemTest = _elemTest.parentNode)
		{
			var _res = _elemTest.getAttribute ? _elemTest.getAttribute("oo_editor_keyboard") : null;
			if (null != _res)
				return _elemTest;
		}
		return null;
	}
	function _getDefaultKeyboardInput(_elem, _depth)
	{
		var _elemTest = _elem;
		for (var _level = 0; _elemTest && (_level < _depth); ++_level, _elemTest = _elemTest.parentNode)
		{
			var _name = " " + _elemTest.className + " ";
			if (_name.indexOf(" dropdown-menu" ) > -1 ||
				_name.indexOf(" dropdown-toggle ") > -1 ||
				_name.indexOf(" dropdown-submenu ") > -1 ||
				_name.indexOf(" canfocused ") > -1)
			{
				return "true";
			}
		}
		return null;
	}

	window['AscCommon']            = window['AscCommon'] || {};
	window['AscCommon'].CTextInput = CTextInput2;

	window['AscCommon'].InitBrowserInputContext = function(api, target_id, parent_id)
	{
		if (window['AscCommon'].g_inputContext)
			return;

		window['AscCommon'].g_inputContext = new CTextInput2(api);
		window['AscCommon'].g_inputContext.init(target_id, parent_id);
		window['AscCommon'].g_clipboardBase.Init(api);
		window['AscCommon'].g_clipboardBase.inputContext = window['AscCommon'].g_inputContext;

		window['AscCommon'].inputMethodCheckInitEvents();

		if (window['AscCommon'].TextBoxInputMode === true)
		{
			window['AscCommon'].g_inputContext.systemInputEnable(true);
		}

		//window["SetInputDebugMode"]();

		document.addEventListener("focus", function(e)
		{
			var t                = window['AscCommon'].g_inputContext;
			var _oldNativeFE	 = t.nativeFocusElement;
			t.nativeFocusElement = e.target;

			t.log("focus");

			if (t.IsComposition)
			{
				t.compositeEnd();
				t.externalEndCompositeInput();
			}

			t.onFocusInputText();

			/*
			if (!t.isNoClearOnFocus)
				t.clear(true);

            t.isNoClearOnFocus = false;
			 */
			t.Api.isBlurEditor = false;

			var _nativeFocusElementNoRemoveOnElementFocus = t.nativeFocusElementNoRemoveOnElementFocus;
			t.nativeFocusElementNoRemoveOnElementFocus = false;

			if (t.InterfaceEnableKeyEvents == false)
			{
				t.nativeFocusElement = null;
				return;
			}

			if (t.nativeFocusElement && (t.nativeFocusElement.id == t.HtmlArea.id))
			{
				t.Api.asc_enableKeyEvents(true, true);

				if (_nativeFocusElementNoRemoveOnElementFocus)
					t.nativeFocusElement = _oldNativeFE;
				else
					t.nativeFocusElement = null;

				return;
			}
			if (t.nativeFocusElement && (t.nativeFocusElement.id == window['AscCommon'].g_clipboardBase.CommonDivId))
			{
				t.nativeFocusElement = null;
				return;
			}

			t.nativeFocusElementNoRemoveOnElementFocus = false;

			var _isElementEditable = false;
			if (t.nativeFocusElement)
			{
				// detect _isElementEditable
				var _name = t.nativeFocusElement.nodeName;
				if (_name)
					_name = _name.toUpperCase();

				if ("INPUT" == _name || "TEXTAREA" == _name)
					_isElementEditable = true;
				else if ("DIV" == _name)
				{
					if (t.nativeFocusElement.getAttribute("contenteditable") == "true")
						_isElementEditable = true;
				}
			}
			if ("IFRAME" == _name)
			{
				// перехват клавиатуры
				t.Api.asc_enableKeyEvents(false, true);
				t.nativeFocusElement = null;
				return;
			}

			// перехватывает ли элемент ввод
			var _oo_editor_input    = _getAttirbute(t.nativeFocusElement, "oo_editor_input", 3);
			// нужно ли прокидывать нажатие клавиш элементу (ТОЛЬКО keyDown)
			var _oo_editor_keyboard = _getAttirbute(t.nativeFocusElement, "oo_editor_keyboard", 3);

			if (!_oo_editor_input && !_oo_editor_keyboard)
				_oo_editor_input = _getDefaultKeyboardInput(t.nativeFocusElement, 3);

			if (_oo_editor_keyboard == "true")
				_oo_editor_input = undefined;

			if (_oo_editor_input == "true")
			{
				// перехват клавиатуры
				t.Api.asc_enableKeyEvents(false, true);
				t.nativeFocusElement = null;
				return;
			}

			if (_isElementEditable && (_oo_editor_input != "false"))
			{
				// перехват клавиатуры
				t.Api.asc_enableKeyEvents(false, true);
				t.nativeFocusElement = null;
				return;
			}

			// итак, ввод у нас. теперь определяем, нужна ли клавиатура элементу
			if (_oo_editor_keyboard != "true")
				t.nativeFocusElement = null;

			var _elem = t.nativeFocusElement;
			t.nativeFocusElementNoRemoveOnElementFocus = true; // ie focus async
			AscCommon.AscBrowser.isMozilla ? setTimeout(function(){ focusHtmlElement(t.HtmlArea); }, 0) : focusHtmlElement(t.HtmlArea);
			t.nativeFocusElement = _elem;
			t.Api.asc_enableKeyEvents(true, true);
		}, true);

		// send focus
		if (!api.isMobileVersion && !api.isEmbedVersion)
			focusHtmlElement(window['AscCommon'].g_inputContext.HtmlArea);
	};

	function focusHtmlElement(element)
	{
		element.focus();
		/*
		var api = window['AscCommon'].g_inputContext.Api;
		if (api.isMobileVersion)
			element.focus();
		else
			element.focus({ "preventScroll" : true });
		*/
	};

	window["SetInputDebugMode"] = function()
	{
		if (!window['AscCommon'].g_inputContext)
			return;

		window['AscCommon'].g_inputContext.debugInputEnable(true);
		window['AscCommon'].g_inputContext.show();
	};

	window['AscCommon'].StartIntervalDrawText = function (isStart) {
		if (isStart) {
			window.renderIntervalId = setInterval(function(){

				window.LOCK_DRAW = false;

				if (undefined !== window.TEXT_DRAW_INSTANCE)
					window.TEXT_DRAW_INSTANCE._renderText(window.TEXT_DRAW_INSTANCE_POS);

				window.TEXT_DRAW_INSTANCE = undefined;
				window.TEXT_DRAW_INSTANCE_POS = 0;

			}, 20);
		} else {
			clearInterval(window.renderIntervalId);
		}
	};
})(window);

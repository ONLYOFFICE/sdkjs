/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function(window)
{
	/**
	 * @constructor
	 */
	function FFData()
	{
		this.calcOnExit = undefined; // bool
		this.checkBox   = undefined; // CheckBox
		this.ddList     = undefined; // DDList
		this.enabled    = undefined; // bool
		this.entryMacro = undefined; // string
		this.exitMacro  = undefined; // string
		this.helpText   = undefined; // FFDataText
		this.label      = undefined; // int
		this.name       = undefined; // string
		this.statusText = undefined; // FFDataText
		this.tabIndex   = undefined; // int
		this.textInput  = undefined; // TextInput
	}
	FFData.prototype.Copy = function()
	{
		let ffData = new FFData();
		
		ffData.calcOnExit = this.calcOnExit;
		ffData.checkBox   = this.checkBox ? this.checkBox.Copy() : undefined;
		ffData.ddList     = this.ddList ? this.ddList.Copy() : undefined;
		ffData.enabled    = this.enabled;
		ffData.entryMacro = this.entryMacro;
		ffData.exitMacro  = this.exitMacro;
		ffData.helpText   = this.helpText ? this.helpText.Copy() : undefined;
		ffData.label      = this.label;
		ffData.name       = this.name;
		ffData.statusText = this.statusText ? this.statusText.Copy() : undefined;
		ffData.tabIndex   = this.tabIndex;
		ffData.textInput  = this.textInput ? this.textInput.Copy() : undefined;
		return ffData;
	};
	FFData.prototype.initCheckBox = function()
	{
		this.checkBox = new CheckBox();
		return this.checkBox;
	};
	FFData.prototype.initDDList = function()
	{
		this.ddList = new DDList();
		return this.ddList;
	};
	FFData.prototype.initHelpText = function()
	{
		this.helpText = new FFDataText();
		return this.helpText;
	};
	FFData.prototype.initStatusText = function()
	{
		this.statusText = new FFDataText();
		return this.statusText;
	};
	FFData.prototype.initTextInput = function()
	{
		this.textInput = new TextInput();
		return this.textInput;
	};
	FFData.prototype.isCheckBoxAutoSize = function()
	{
		return !this.checkBox || undefined === this.checkBox.size || this.checkBox.sizeAuto;
	};
	FFData.prototype.getCheckBoxSize = function()
	{
		return this.checkBox && undefined !== this.checkBox.size ? this.checkBox.size / 2 : 10;
	};
	FFData.prototype.isCheckBoxChecked = function()
	{
		if (!this.checkBox)
			return false;
		
		if (undefined !== this.checkBox.checked)
			return this.checkBox.checked;
		
		return !!(this.checkBox.default);
	};
	FFData.prototype.isEnabled = function()
	{
		return (false !== this.enabled);
	};
	FFData.prototype.toBinary = function(writer)
	{
		let startPos = writer.GetCurPosition();
		writer.Skip(4);
		
		let flags = 0;
		if (undefined !== this.calcOnExit)
		{
			writer.WriteBool(this.calcOnExit);
			flags |= 1;
		}
		
		if (undefined !== this.checkBox)
		{
			this.checkBox.toBinary(writer);
			flags |= 2;
		}
		
		if (undefined !== this.ddList)
		{
			this.ddList.toBinary(writer);
			flags |= 4;
		}
		
		if (undefined !== this.enabled)
		{
			writer.WriteBool(this.enabled);
			flags |= 8;
		}
		
		if (undefined !== this.entryMacro)
		{
			writer.WriteString2(this.entryMacro);
			flags |= 16;
		}
		
		if (undefined !== this.exitMacro)
		{
			writer.WriteString2(this.exitMacro);
			flags |= 32;
		}
		
		if (undefined !== this.helpText)
		{
			this.helpText.toBinary(writer);
			flags |= 64;
		}
		
		if (undefined !== this.label)
		{
			writer.WriteLong(this.label);
			flags |= 128;
		}
		
		if (undefined !== this.name)
		{
			writer.WriteString2(this.name);
			flags |= 256;
		}
		
		if (undefined !== this.statusText)
		{
			this.statusText.toBinary(writer);
			flags |= 512;
		}
		
		if (undefined !== this.tabIndex)
		{
			writer.WriteLong(this.tabIndex);
			flags |= 1024;
		}
		
		if (undefined !== this.textInput)
		{
			this.textInput.toBinary(writer);
			flags |= 2048;
		}
		
		let endPos = writer.GetCurPosition();
		writer.Seek(startPos);
		writer.WriteLong(flags);
		writer.Seek(endPos);
	};
	FFData.prototype.isCheckBox = function()
	{
		return (!!this.checkBox);
	};
	FFData.fromBinary = function(reader)
	{
		let ffData = new FFData();
		
		let flags = reader.GetLong();
		
		if (flags & 1)
			ffData.calcOnExit = reader.GetBool();
		
		if (flags & 2)
			ffData.checkBox = CheckBox.fromBinary(reader);
		
		if (flags & 4)
			ffData.ddList = DDList.fromBinary(reader);
		
		if (flags & 8)
			ffData.enabled = reader.GetBool();
		
		if (flags & 16)
			ffData.entryMacro = reader.GetString2();
		
		if (flags & 32)
			ffData.exitMacro = reader.GetString2();
		
		if (ffData & 64)
			ffData.helpText = FFDataText.fromBinary(reader);
		
		if (ffData & 128)
			ffData.label = reader.GetLong();
		
		if (ffData & 256)
			ffData.name = reader.GetString2();
		
		if (ffData & 512)
			ffData.statusText = FFDataText.fromBinary(reader);
		
		if (ffData & 1024)
			ffData.tabIndex = reader.GetLong();
		
		if (ffData & 2048)
			ffData.textInput = TextInput.fromBinary(reader);
		
		return ffData;
	};
	
	/**
	 * ffData.checkBox
	 * @constructor
	 */
	function CheckBox()
	{
		this.checked  = undefined;
		this.default  = undefined;
		this.size     = undefined;
		this.sizeAuto = undefined;
	}
	CheckBox.prototype.Copy = function()
	{
		let cb = new CheckBox();
		
		cb.checked  = this.checked;
		cb.default  = this.default;
		cb.size     = this.size;
		cb.sizeAuto = this.sizeAuto;
		return cb;
	};
	CheckBox.prototype.toBinary = function(writer)
	{
		let startPos = writer.GetCurPosition();
		writer.Skip(4);
		
		let flags = 0;
		if (undefined !== this.checked)
		{
			writer.WriteBool(this.checked);
			flags |= 1;
		}
		
		if (undefined !== this.default)
		{
			writer.WriteBool(this.default);
			flags |= 2;
		}
		if (undefined !== this.size)
		{
			writer.WriteLong(this.size);
			flags |= 4;
		}
		
		if (undefined !== this.sizeAuto)
		{
			writer.WriteBool(this.sizeAuto);
			flags |= 8;
		}
		
		let endPos = writer.GetCurPosition();
		writer.Seek(startPos);
		writer.WriteLong(flags);
		writer.Seek(endPos);
	};
	CheckBox.fromBinary = function(reader)
	{
		let checkBox = new CheckBox();
		
		let flags = reader.GetLong();
		if (flags & 1)
			checkBox.checked = reader.GetBool();
		
		if (flags & 2)
			checkBox.default = reader.GetBool();
		
		if (flags & 4)
			checkBox.size = reader.GetLong();
		
		if (flags & 8)
			checkBox.sizeAuto = reader.GetBool();
		
		return checkBox;
	};
	
	/**
	 * ffData.ddList
	 * @constructor
	 */
	function DDList()
	{
		this.default = undefined;
		this.list    = [];
		this.result  = undefined;
	}
	DDList.prototype.Copy = function()
	{
		let ddList = new DDList();
		
		ddList.default = this.default;
		ddList.result  = this.result;
		ddList.list    = this.list.slice();
		return ddList;
	};
	DDList.prototype.toBinary = function(writer)
	{
		let flags = 0;
		if (undefined !== this.default)
			flags |= 1;
		if (undefined !== this.result)
			flags |= 2;
		writer.WriteLong(flags);
		
		if (undefined !== this.default)
			writer.WriteLong(this.default);
		
		if (undefined !== this.result)
			writer.WriteLong(this.result);
		
		writer.WriteLong(this.list.length);
		for (let i = 0; i < this.list.length; ++i)
		{
			writer.WriteString2(this.list[i]);
		}
	};
	DDList.fromBinary = function(reader)
	{
		let ddList = new DDList();
		
		let flags = reader.GetBool();
		if (flags & 1)
			ddList.default = reader.GetLong();
		if (flags & 2)
			ddList.result = reader.GetLong();
		
		let count = reader.GetLong();
		for (let i = 0; i < count; ++i)
		{
			ddList.list.push(reader.GetString2());
		}
		
		return ddList;
	};
	
	/**
	 * ffData.helpText or ffData.statusText
	 * @constructor
	 */
	function FFDataText()
	{
		this.type = undefined;
		this.val  = undefined;
	}
	FFDataText.prototype.Copy = function()
	{
		let ht = new FFDataText();
		
		ht.type = this.type;
		ht.val  = this.val;
		return ht;
	};
	FFDataText.prototype.toBinary = function(writer)
	{
		let flags = 0;
		if (undefined !== this.type)
			flags |= 1;
		if (undefined !== this.val)
			flags |= 2;
		
		writer.WriteByte(flags);
		
		if (undefined !== this.type)
			writer.WriteByte(this.type);
		if (undefined !== this.val)
			writer.WriteString2(this.val);
	};
	FFDataText.fromBinary = function(reader)
	{
		let text = new FFDataText();
		
		let flags = reader.GetByte();
		if (flags & 1)
			text.type = reader.GetByte();
		if (flags & 2)
			text.val = reader.GetString2();
		
		return text;
	};
	
	/**
	 * ffData.textInput
	 * @constructor
	 */
	function TextInput()
	{
		this.default   = undefined; // string
		this.format    = undefined; // string
		this.type      = undefined;
		this.maxLength = undefined;
	}
	TextInput.prototype.Copy = function()
	{
		let ti = new TextInput();
		
		ti.default   = this.default;
		ti.format    = this.format;
		ti.type      = this.type;
		ti.maxLength = this.maxLength;
		return ti;
	};
	TextInput.prototype.toBinary = function(writer)
	{
		let startPos = writer.GetCurPosition();
		writer.Skip(4);
		
		let flags = 0;
		if (undefined !== this.default)
		{
			writer.WriteString2(this.default);
			flags |= 1;
		}
		
		if (undefined !== this.format)
		{
			writer.WriteString2(this.format);
			flags |= 2;
		}
		if (undefined !== this.type)
		{
			writer.WriteByte(this.type);
			flags |= 4;
		}
		
		if (undefined !== this.maxLength)
		{
			writer.WriteLong(this.maxLength);
			flags |= 8;
		}
		
		let endPos = writer.GetCurPosition();
		writer.Seek(startPos);
		writer.WriteLong(flags);
		writer.Seek(endPos);
	};
	TextInput.fromBinary = function(reader)
	{
		let textInput = new TextInput();
		
		let flags = reader.GetByte();
		if (flags & 1)
			textInput.default = reader.GetString2();
		if (flags & 2)
			textInput.format = reader.GetString2();
		if (flags & 4)
			textInput.type = reader.GetByte();
		if (flags & 8)
			textInput.maxLength = reader.GetLong();
		
		return textInput;
	};
	
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].FFData = FFData;
	
})(window);


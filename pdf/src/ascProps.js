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

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {

	const Asc = window['Asc'];
	const AscCommon = window['AscCommon'];

	/** @constructor */
	function asc_CAnnotProperty() {
		this.ids = null;

		this.type = null;

		this.fill	= null;
		this.stroke	= null;
		this.opacity= undefined;
		this.subject= undefined;

		this.annotProps = null;
	}

	asc_CAnnotProperty.prototype.asc_getIds = function () {
		return this.ids;
	};
	asc_CAnnotProperty.prototype.asc_putIds = function (v) {
		this.ids = v;
	};
	asc_CAnnotProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CAnnotProperty.prototype.asc_putType = function (v) {
		this.type = v;
	};
	asc_CAnnotProperty.prototype.asc_getFill = function () {
		return this.fill;
	};
	asc_CAnnotProperty.prototype.asc_putFill = function (v) {
		this.fill = v;
	};
	asc_CAnnotProperty.prototype.asc_getStroke = function () {
		return this.stroke;
	};
	asc_CAnnotProperty.prototype.asc_putStroke = function (v) {
		this.stroke = v;
	};
	asc_CAnnotProperty.prototype.asc_getOpacity = function () {
		return this.opacity;
	};
	asc_CAnnotProperty.prototype.asc_putOpacity = function (v) {
		this.opacity = v;
	};
	asc_CAnnotProperty.prototype.asc_getSubject = function () {
		return this.subject;
	};
	asc_CAnnotProperty.prototype.asc_putSubject = function (v) {
		this.subject = v;
	};
	asc_CAnnotProperty.prototype.asc_getAnnotProps = function () {
		return this.annotProps;
	};
	asc_CAnnotProperty.prototype.asc_putAnnotProps = function (v) {
		this.annotProps = v;
	};
	asc_CAnnotProperty.prototype.compare = function(pr) {
		if (this.type !== pr.type) {
			this.type = null;
		}

		if (!this.fill || !pr.fill || this.fill.r !== pr.fill.r || this.fill.g !== pr.fill.g || this.fill.b !== pr.fill.b) {
			this.fill = null;
		}
		if (!this.stroke || !pr.stroke || this.stroke.r !== pr.stroke.r || this.stroke.g !== pr.stroke.g || this.stroke.b !== pr.stroke.b) {
			this.stroke = null;
		}
		if (this.opacity !== pr.opacity) {
			this.opacity = null;
		}
		if (this.subject !== pr.subject) {
			this.subject = null;
		}
		if (this.type && this.annotProps) {
			this.annotProps.compare(pr.annotProps);
		}
		else {
			this.annotProps = null;
		}
	};

	// free text
	function asc_CFreeTextAnnotProperty() {
		this.borderWidth		= undefined;
		this.borderStyle		= undefined;
		this.lineEnd			= null;
		this.canEditText		= undefined;
	}

	asc_CFreeTextAnnotProperty.prototype.asc_getBorderWidth = function () {
		return this.borderWidth;
	};
	asc_CFreeTextAnnotProperty.prototype.asc_putBorderWidth = function (v) {
		this.borderWidth = v;
	};
	asc_CFreeTextAnnotProperty.prototype.asc_getBorderStyle = function () {
		return this.borderStyle;
	};
	asc_CFreeTextAnnotProperty.prototype.asc_putBorderStyle = function (v) {
		this.borderStyle = v;
	};
	asc_CFreeTextAnnotProperty.prototype.asc_getLineEnd = function () {
		return this.lineEnd;
	};
	asc_CFreeTextAnnotProperty.prototype.asc_putLineEnd = function (v) {
		this.lineEnd = v;
	};
	asc_CFreeTextAnnotProperty.prototype.asc_getCanEditText = function () {
		return this.canEditText;
	};
	asc_CFreeTextAnnotProperty.prototype.asc_putCanEditText = function (v) {
		this.canEditText = v;
	};
	asc_CFreeTextAnnotProperty.prototype.compare = function(pr) {
		if (this.borderWidth !== pr.borderWidth) {
			this.borderWidth = null;
		}
		if (this.borderStyle !== pr.borderStyle) {
			this.borderStyle = null;
		}
		if (this.lineEnd !== pr.lineEnd) {
			this.lineEnd = null;
		}
		if (this.canEditText !== pr.canEditText) {
			this.canEditText = null;
		}
	};

	// line
	function asc_CLineAnnotProperty() {
		this.borderStyle= undefined;
		this.borderWidth= undefined;
		this.lineEnd	= undefined;
		this.lineStart	= undefined;
		this.canEditText= undefined;
	}

	asc_CLineAnnotProperty.prototype.asc_getBorderStyle = function () {
		return this.borderStyle;
	};
	asc_CLineAnnotProperty.prototype.asc_putBorderStyle = function (v) {
		this.borderStyle = v;
	};
	asc_CLineAnnotProperty.prototype.asc_getBorderWidth = function () {
		return this.borderWidth;
	};
	asc_CLineAnnotProperty.prototype.asc_putBorderWidth = function (v) {
		this.borderWidth = v;
	};
	asc_CLineAnnotProperty.prototype.asc_getLineEnd = function () {
		return this.lineEnd;
	};
	asc_CLineAnnotProperty.prototype.asc_putLineEnd = function (v) {
		this.lineEnd = v;
	};
	asc_CLineAnnotProperty.prototype.asc_getLineStart = function () {
		return this.lineStart;
	};
	asc_CLineAnnotProperty.prototype.asc_putLineStart = function (v) {
		this.lineStart = v;
	};
	asc_CLineAnnotProperty.prototype.asc_getCanEditText = function () {
		return this.canEditText;
	};
	asc_CLineAnnotProperty.prototype.asc_putCanEditText = function (v) {
		this.canEditText = v;
	};
	asc_CLineAnnotProperty.prototype.compare = function(pr) {
		if (this.borderStyle !== pr.borderStyle) {
			this.borderStyle = null;
		}
		if (this.borderWidth !== pr.borderWidth) {
			this.borderWidth = null;
		}
		if (this.lineEnd !== pr.lineEnd) {
			this.lineEnd = null;
		}
		if (this.lineStart !== pr.lineStart) {
			this.lineStart = null;
		}
		if (this.canEditText !== pr.canEditText) {
			this.canEditText = null;
		}
	};

	// polyline
	function asc_CPolyLineAnnotProperty() {
		this.borderStyle= undefined;
		this.borderWidth= undefined;
		this.lineEnd	= undefined;
		this.lineStart	= undefined;
	}

	asc_CPolyLineAnnotProperty.prototype.asc_getBorderStyle = function () {
		return this.borderStyle;
	};
	asc_CPolyLineAnnotProperty.prototype.asc_putBorderStyle = function (v) {
		this.borderStyle = v;
	};
	asc_CPolyLineAnnotProperty.prototype.asc_getBorderWidth = function () {
		return this.borderWidth;
	};
	asc_CPolyLineAnnotProperty.prototype.asc_putBorderWidth = function (v) {
		this.borderWidth = v;
	};
	asc_CPolyLineAnnotProperty.prototype.asc_getLineEnd = function () {
		return this.lineEnd;
	};
	asc_CPolyLineAnnotProperty.prototype.asc_putLineEnd = function (v) {
		this.lineEnd = v;
	};
	asc_CPolyLineAnnotProperty.prototype.asc_getLineStart = function () {
		return this.lineStart;
	};
	asc_CPolyLineAnnotProperty.prototype.asc_putLineStart = function (v) {
		this.lineStart = v;
	};
	asc_CPolyLineAnnotProperty.prototype.compare = function(pr) {
		if (this.borderStyle !== pr.borderStyle) {
			this.borderStyle = null;
		}
		if (this.borderWidth !== pr.borderWidth) {
			this.borderWidth = null;
		}
		if (this.lineEnd !== pr.lineEnd) {
			this.lineEnd = null;
		}
		if (this.lineStart !== pr.lineStart) {
			this.lineStart = null;
		}
	};

	// polygon/square/circle
	function asc_CClosedAnnotProperty() {
		this.borderStyle= undefined;
		this.borderWidth= undefined;
	}

	asc_CClosedAnnotProperty.prototype.asc_getBorderStyle = function () {
		return this.borderStyle;
	};
	asc_CClosedAnnotProperty.prototype.asc_putBorderStyle = function (v) {
		this.borderStyle = v;
	};
	asc_CClosedAnnotProperty.prototype.asc_getBorderWidth = function () {
		return this.borderWidth;
	};
	asc_CClosedAnnotProperty.prototype.asc_putBorderWidth = function (v) {
		this.borderWidth = v;
	};
	asc_CClosedAnnotProperty.prototype.compare = function(pr) {
		if (this.borderStyle !== pr.borderStyle) {
			this.borderStyle = null;
		}
		if (this.borderWidth !== pr.borderWidth) {
			this.borderWidth = null;
		}
	};

	//////////////////////////////////////////////////////////////////
	///// Common field
	/////////////////////////////////////////////////////////////////
	function asc_CBaseFieldProperty() {
		// common
		this.type				= undefined;

		this.name				= undefined;
		this.required			= undefined;
		this.readOnly			= undefined;
		this.rot				= undefined;
		this.display			= undefined;
		this.fill				= null;
		this.stroke				= null;
		this.strokeWidth		= undefined;
		this.strokeStyle		= undefined;
		this.tooltip			= undefined;
		this.digitsType			= undefined;
		this.locked				= false;

		this.actions	= null;
		this.fieldProps	= null;
	}
	asc_CBaseFieldProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CBaseFieldProperty.prototype.asc_putType = function (v) {
		this.type = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getName = function () {
		return this.name;
	};
	asc_CBaseFieldProperty.prototype.asc_putName = function (v) {
		this.name = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getRequired = function () {
		return this.required;
	};
	asc_CBaseFieldProperty.prototype.asc_putRequired = function (v) {
		this.required = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getReadOnly = function () {
		return this.readOnly;
	};
	asc_CBaseFieldProperty.prototype.asc_putReadOnly = function (v) {
		this.readOnly = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getRot = function () {
		return this.rot;
	};
	asc_CBaseFieldProperty.prototype.asc_putRot = function (v) {
		this.rot = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getDisplay = function () {
		return this.display;
	};
	asc_CBaseFieldProperty.prototype.asc_putDisplay = function (v) {
		this.display = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getFill = function () {
		return this.fill;
	};
	asc_CBaseFieldProperty.prototype.asc_putFill = function (v) {
		this.fill = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getStroke = function () {
		return this.stroke;
	};
	asc_CBaseFieldProperty.prototype.asc_putStroke = function (v) {
		this.stroke = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getStrokeWidth = function () {
		return this.strokeWidth;
	};
	asc_CBaseFieldProperty.prototype.asc_putStrokeWidth = function (v) {
		this.strokeWidth = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getStrokeStyle = function () {
		return this.strokeStyle;
	};
	asc_CBaseFieldProperty.prototype.asc_putStrokeStyle = function (v) {
		this.strokeStyle = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getPropLocked = function () {
		return this.locked;
	};
	asc_CBaseFieldProperty.prototype.asc_putPropLocked = function (v) {
		this.locked = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getTooltip = function () {
		return this.tooltip;
	};
	asc_CBaseFieldProperty.prototype.asc_putTooltip = function (v) {
		this.tooltip = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getDigitsType = function () {
		return this.digitsType;
	};
	asc_CBaseFieldProperty.prototype.asc_putDigitsType = function (v) {
		this.digitsType = v;
	};
	asc_CBaseFieldProperty.prototype.get_Locked = function () {
		return this.coEditLocked;
	};
	asc_CBaseFieldProperty.prototype.put_Locked = function (v) {
		this.coEditLocked = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getActionsProps = function () {
		return this.actions;
	};
	asc_CBaseFieldProperty.prototype.asc_putActionsProps = function (v) {
		this.actions = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getFieldProps = function () {
		return this.fieldProps;
	};
	asc_CBaseFieldProperty.prototype.asc_putFieldProps = function (v) {
		this.fieldProps = v;
	};
	asc_CBaseFieldProperty.prototype.compare = function (pr) {
		if (this.type !== pr.type) {
			this.type = null;
		}
		if (this.name !== pr.name) {
			this.name = null;
		}
		if (this.required !== pr.required) {
			this.required = null;
		}
		if (this.readOnly !== pr.readOnly) {
			this.readOnly = null;
		}
		if (this.rot !== pr.rot) {
			this.rot = null;
		}
		if (this.display !== pr.display) {
			this.display = null;
		}
		if (!this.fill || !pr.fill || this.fill.r !== pr.fill.r || this.fill.g !== pr.fill.g || this.fill.b !== pr.fill.b) {
			this.fill = null;
		}
		if (!this.stroke || !pr.stroke || this.stroke.r !== pr.stroke.r || this.stroke.g !== pr.stroke.g || this.stroke.b !== pr.stroke.b) {
			this.stroke = null;
		}
		if (this.strokeWidth !== pr.strokeWidth) {
			this.strokeWidth = null;
		}
		if (this.strokeStyle !== pr.strokeStyle) {
			this.strokeStyle = null;
		}
		if (this.tooltip !== pr.tooltip) {
			this.tooltip = null;
		}
		if (this.hindiDigits !== pr.hindiDigits) {
			this.hindiDigits = null;
		}
		if (this.locked !== pr.locked) {
			this.locked = null;
		}

		if (this.type != undefined && this.fieldProps && pr.fieldProps) {
			this.fieldProps.compare(pr.fieldProps);
		}
		else {
			this.fieldProps = null;
		}
	};
	//////////////////////////////////////////////////////////////////
	///// Text field
	//////////////////////////////////////////////////////////////////
	function asc_CTextFieldProperty() {
		// format
		this.format				= null;

		// text
		this.defaultValue		= undefined;
		this.multiline			= undefined;
		this.scrollLongText		= undefined;
		this.charLimit			= undefined;
		this.comb				= undefined;
		this.placeholder		= undefined;
		this.autoFit			= undefined;
		this.password			= undefined;
	}
	asc_CTextFieldProperty.prototype.asc_getDefaultValue = function () {
		return this.defaultValue;
	};
	asc_CTextFieldProperty.prototype.asc_putDefaultValue = function (v) {
		this.defaultValue = v;
	};
	asc_CTextFieldProperty.prototype.asc_getMultiline = function () {
		return this.multiline;
	};
	asc_CTextFieldProperty.prototype.asc_putMultiline = function (v) {
		this.multiline = v;
	};
	asc_CTextFieldProperty.prototype.asc_getScrollLongText = function () {
		return this.scrollLongText;
	};
	asc_CTextFieldProperty.prototype.asc_putScrollLongText = function (v) {
		this.scrollLongText = v;
	};
	asc_CTextFieldProperty.prototype.asc_getCharLimit = function () {
		return this.charLimit;
	};
	asc_CTextFieldProperty.prototype.asc_putCharLimit = function (v) {
		this.charLimit = v;
	};
	asc_CTextFieldProperty.prototype.asc_getComb = function () {
		return this.comb;
	};
	asc_CTextFieldProperty.prototype.asc_putComb = function (v) {
		this.comb = v;
	};
	asc_CTextFieldProperty.prototype.asc_getPlaceholder = function () {
		return this.placeholder;
	};
	asc_CTextFieldProperty.prototype.asc_putPlaceholder = function (v) {
		this.placeholder = v;
	};
	asc_CTextFieldProperty.prototype.asc_getAutoFit = function () {
		return this.autoFit;
	};
	asc_CTextFieldProperty.prototype.asc_putAutoFit = function (v) {
		this.autoFit = v;
	};
	asc_CTextFieldProperty.prototype.asc_getPassword = function () {
		return this.password;
	};
	asc_CTextFieldProperty.prototype.asc_putPassword = function (v) {
		this.password = v;
	};
	asc_CTextFieldProperty.prototype.compare = function (pr) {
		if (this.defaultValue !== pr.defaultValue) {
			this.defaultValue = null;
		}
		if (this.multiline !== pr.multiline) {
			this.multiline = null;
		}
		if (this.scrollLongText !== pr.scrollLongText) {
			this.scrollLongText = null;
		}
		if (this.charLimit !== pr.charLimit) {
			this.charLimit = null;
		}
		if (this.comb !== pr.comb) {
			this.comb = null;
		}
		if (this.placeholder !== pr.placeholder) {
			this.placeholder = null;
		}
		if (this.autoFit !== pr.autoFit) {
			this.autoFit = null;
		}
	};

	//////////////////////////////////////////////////////////////////
	///// Combobox field
	//////////////////////////////////////////////////////////////////
	function asc_CComboboxFieldProperty() {
		// format
		this.options			= [];
		this.commitOnSelChange	= undefined;
		this.editable			= undefined;
		this.placeholder		= undefined;
		this.autoFit			= undefined;
	}
	asc_CComboboxFieldProperty.prototype.asc_getOptions = function () {
		return this.options;
	};
	asc_CComboboxFieldProperty.prototype.asc_putOptions = function (v) {
		this.options = v;
	};
	asc_CComboboxFieldProperty.prototype.asc_getCommitOnSelChange = function () {
		return this.commitOnSelChange;
	};
	asc_CComboboxFieldProperty.prototype.asc_putCommitOnSelChange = function (v) {
		this.commitOnSelChange = v;
	};
	asc_CComboboxFieldProperty.prototype.asc_getEditable = function () {
		return this.editable;
	};
	asc_CComboboxFieldProperty.prototype.asc_putEditable = function (v) {
		this.editable = v;
	};
	asc_CComboboxFieldProperty.prototype.asc_getPlaceholder = function () {
		return this.placeholder;
	};
	asc_CComboboxFieldProperty.prototype.asc_putPlaceholder = function (v) {
		this.placeholder = v;
	};
	asc_CComboboxFieldProperty.prototype.asc_getAutoFit = function () {
		return this.autoFit;
	};
	asc_CComboboxFieldProperty.prototype.asc_putAutoFit = function (v) {
		this.autoFit = v;
	};
	asc_CComboboxFieldProperty.prototype.compare = function (pr) {
		if (!AscCommon.isEqualSortedArrays(this.options, pr.options)) {
			this.options = [];
		}

		if (this.commitOnSelChange !== pr.commitOnSelChange) {
			this.commitOnSelChange = null;
		}
		if (this.editable !== pr.editable) {
			this.editable = null;
		}
		if (this.placeholder !== pr.placeholder) {
			this.placeholder = null;
		}
		if (this.autoFit !== pr.autoFit) {
			this.autoFit = null;
		}
	};

	//////////////////////////////////////////////////////////////////
	///// Listbox field
	//////////////////////////////////////////////////////////////////
	function asc_CListboxFieldProperty() {
		this.options		 	= [];
		this.commitOnSelChange	= undefined;
		this.multipleSelection	= undefined;
	}
	asc_CListboxFieldProperty.prototype.asc_getOptions = function () {
		return this.options;
	};
	asc_CListboxFieldProperty.prototype.asc_putOptions = function (v) {
		this.options = v;
	};
	asc_CListboxFieldProperty.prototype.asc_getCommitOnSelChange = function () {
		return this.commitOnSelChange;
	};
	asc_CListboxFieldProperty.prototype.asc_putCommitOnSelChange = function (v) {
		this.commitOnSelChange = v;
	};
	asc_CListboxFieldProperty.prototype.asc_getMultipleSelection = function () {
		return this.multipleSelection;
	};
	asc_CListboxFieldProperty.prototype.asc_putMultipleSelection = function (v) {
		this.multipleSelection = v;
	};
	asc_CListboxFieldProperty.prototype.compare = function (pr) {
		if (!AscCommon.isEqualSortedArrays(this.options, pr.options)) {
			this.options = [];
		}

		if (this.commitOnSelChange !== pr.commitOnSelChange) {
			this.commitOnSelChange = null;
		}
		if (this.multipleSelection !== pr.multipleSelection) {
			this.multipleSelection = null;
		}
	};
	//////////////////////////////////////////////////////////////////
	///// Checkbox
	//////////////////////////////////////////////////////////////////
	function asc_CCheckboxFieldProperty() {
		this.checkboxStyle	= undefined;
		this.exportValue	= undefined;
		this.defaultChecked	= undefined;
		this.toggleToOff	= undefined;
	}
	asc_CCheckboxFieldProperty.prototype.asc_getCheckboxStyle = function () {
		return this.checkboxStyle;
	};
	asc_CCheckboxFieldProperty.prototype.asc_putCheckboxStyle = function (v) {
		this.checkboxStyle = v;
	};
	asc_CCheckboxFieldProperty.prototype.asc_getExportValue = function () {
		return this.exportValue;
	};
	asc_CCheckboxFieldProperty.prototype.asc_putExportValue = function (v) {
		this.exportValue = v;
	};
	asc_CCheckboxFieldProperty.prototype.asc_getDefaultChecked = function () {
		return this.defaultChecked;
	};
	asc_CCheckboxFieldProperty.prototype.asc_putDefaultChecked = function (v) {
		this.defaultChecked = v;
	};
	asc_CCheckboxFieldProperty.prototype.asc_getToggleToOff = function () {
		return this.toggleToOff;
	};
	asc_CCheckboxFieldProperty.prototype.asc_putToggleToOff = function (v) {
		this.toggleToOff = v;
	};
	asc_CCheckboxFieldProperty.prototype.compare = function (pr) {
		if (this.checkboxStyle !== pr.checkboxStyle) {
			this.checkboxStyle = null;
		}
		if (this.exportValue !== pr.exportValue) {
			this.exportValue = null;
		}
		if (this.defaultChecked !== pr.defaultChecked) {
			this.defaultChecked = null;
		}
		if (this.toggleToOff !== pr.toggleToOff) {
			this.toggleToOff = null;
		}
	};

	//////////////////////////////////////////////////////////////////
	///// Radiobutton
	//////////////////////////////////////////////////////////////////
	function asc_CRadiobuttonFieldProperty() {
		asc_CCheckboxFieldProperty.call(this);
		this.radiosInUnison	= undefined;
	}
	asc_CRadiobuttonFieldProperty.prototype = Object.create(asc_CCheckboxFieldProperty.prototype);
	asc_CRadiobuttonFieldProperty.prototype.constructor = asc_CRadiobuttonFieldProperty;

	asc_CRadiobuttonFieldProperty.prototype.asc_getRadiosInUnison = function () {
		return this.radiosInUnison;
	};
	asc_CRadiobuttonFieldProperty.prototype.asc_putRadiosInUnison = function (v) {
		this.radiosInUnison = v;
	};
	asc_CRadiobuttonFieldProperty.prototype.compare = function (pr) {
		asc_CCheckboxFieldProperty.prototype.compare.call(this, pr);

		if (this.radiosInUnison !== pr.radiosInUnison) {
			this.radiosInUnison = null;
		}
	};
	//////////////////////////////////////////////////////////////////
	///// Pushbutton
	//////////////////////////////////////////////////////////////////
	function asc_CButtonFieldProperty(buttonField) {
		this.parentFields	= [buttonField];

		this.highlight		= undefined;
		this.layout			= undefined;
		this.scaleWhen		= undefined;
		this.scaleHow		= undefined;
		this.fitBounds		= undefined;
		this.iconPos		= null;
		this.behavior		= undefined;
		this.currentState	= undefined;
		this.normalCaption	= undefined;
		this.hoverCaption	= undefined;
		this.downCaption	= undefined;
		this.normalImage	= undefined;
		this.hoverImage		= undefined;
		this.downImage		= undefined;

		this.DivId			= undefined;
	}
	asc_CButtonFieldProperty.prototype.getParentFields = function () {
		return this.parentFields;
	};
	asc_CButtonFieldProperty.prototype.asc_getHighlight = function () {
		return this.highlight;
	};
	asc_CButtonFieldProperty.prototype.asc_putHighlight = function (v) {
		this.highlight = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getLayout = function () {
		return this.layout;
	};
	asc_CButtonFieldProperty.prototype.asc_putLayout = function (v) {
		this.layout = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getScaleWhen = function () {
		return this.scaleWhen;
	};
	asc_CButtonFieldProperty.prototype.asc_putScaleWhen = function (v) {
		this.scaleWhen = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getScaleHow = function () {
		return this.scaleHow;
	};
	asc_CButtonFieldProperty.prototype.asc_putScaleHow = function (v) {
		this.scaleHow = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getFitBounds = function () {
		return this.fitBounds;
	};
	asc_CButtonFieldProperty.prototype.asc_putFitBounds = function (v) {
		this.fitBounds = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getIconPos = function () {
		return this.iconPos;
	};
	asc_CButtonFieldProperty.prototype.asc_putIconPos = function (v) {
		this.iconPos = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getBehavior = function () {
		return this.behavior;
	};
	asc_CButtonFieldProperty.prototype.asc_putBehavior = function (v) {
		this.behavior = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getNormalCaption = function () {
		return this.normalCaption;
	};
	asc_CButtonFieldProperty.prototype.asc_putNormalCaption = function (v) {
		this.normalCaption = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getNormalImage = function () {
		return this.normalImage;
	};
	asc_CButtonFieldProperty.prototype.asc_putNormalImage = function (v) {
		this.normalImage = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getHoverCaption = function () {
		return this.hoverCaption;
	};
	asc_CButtonFieldProperty.prototype.asc_putHoverCaption = function (v) {
		this.hoverCaption = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getHoverImage = function () {
		return this.hoverImage;
	};
	asc_CButtonFieldProperty.prototype.asc_putHoverImage = function (v) {
		this.hoverImage = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getDownCaption = function () {
		return this.downCaption;
	};
	asc_CButtonFieldProperty.prototype.asc_putDownCaption = function (v) {
		this.downCaption = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getDownImage = function () {
		return this.downImage;
	};
	asc_CButtonFieldProperty.prototype.asc_putDownImage = function (v) {
		this.downImage = v;
	};
	asc_CButtonFieldProperty.prototype.asc_putCurrentState = function(v) {
		this.currentState = v;

		// set to fields state to add image (will clear after set image)
		this.getParentFields().forEach(function(field) {
			field.asc_curImageState = v;
		});

		this.drawTexture(v);
	};
	asc_CButtonFieldProperty.prototype.asc_getCurrentState = function(v) {
		return this.currentState;
	};
	asc_CButtonFieldProperty.prototype.put_DivId = function (v) {
		this.DivId = v;
		this.drawTexture(this.currentState);
	};
	asc_CButtonFieldProperty.prototype.drawTexture = function (nState) {
		let sImageRasterId;
		switch (nState) {
			case AscPDF.APPEARANCE_TYPES.normal:
				sImageRasterId = this.normalImage;
				break;
			case AscPDF.APPEARANCE_TYPES.mouseDown:
				sImageRasterId = this.downImage;
				break;
			case AscPDF.APPEARANCE_TYPES.rollover:
				sImageRasterId = this.hoverImage;
				break;
		}

		var oDiv = document.getElementById(this.DivId);
		if(!oDiv){
			return;
		}

		var aChildren = oDiv.children;
		var oCanvas = null;
		for(var i = 0; i < aChildren.length; ++i){
			if(aChildren[i].nodeName && aChildren[i].nodeName.toUpperCase() === 'CANVAS'){
				oCanvas = aChildren[i];
				break;
			}
		}
		var nWidth = oDiv.clientWidth;
		var nHeight = oDiv.clientHeight;
		if(null === oCanvas){
			oCanvas = document.createElement('canvas');
			oCanvas.width = parseInt(nWidth);
			oCanvas.height = parseInt(nHeight);
			oDiv.appendChild(oCanvas);
		}
		var oContext = oCanvas.getContext('2d');
		oContext.clearRect(0, 0, oCanvas.width, oCanvas.height);
		if (!sImageRasterId) {
			return;
		}

		var _img = Asc.editor.ImageLoader.map_image_index[AscCommon.getFullImageSrc2(sImageRasterId)];
		if (_img != undefined && _img.Image != null && _img.Status != AscFonts.ImageLoadStatus.Loading)
		{
			var _x = 0;
			var _y = 0;
			var _w = Math.max(_img.Image.width, 1);
			var _h = Math.max(_img.Image.height, 1);

			var dAspect1 = nWidth / nHeight;
			var dAspect2 = _w / _h;

			_w = nWidth;
			_h = nHeight;
			if (dAspect1 >= dAspect2)
			{
				_w = dAspect2 * nHeight;
				_x = (nWidth - _w) / 2;
			}
			else
			{
				_h = _w / dAspect2;
				_y = (nHeight - _h) / 2;
			}
			oContext.drawImage(_img.Image, _x, _y, _w, _h);
		}
		else if (!_img || !_img.Image)
		{
			oContext.lineWidth = 1;

			oContext.beginPath();
			oContext.moveTo(0, 0);
			oContext.lineTo(nWidth, nHeight);
			oContext.moveTo(nWidth, 0);
			oContext.lineTo(0, nHeight);
			oContext.strokeStyle = "#FF0000";
			oContext.stroke();

			oContext.beginPath();
			oContext.moveTo(0, 0);
			oContext.lineTo(nWidth, 0);
			oContext.lineTo(nWidth, nHeight);
			oContext.lineTo(0, nHeight);
			oContext.closePath();

			oContext.strokeStyle = "#000000";
			oContext.stroke();
			oContext.beginPath();
		}
	};
	asc_CButtonFieldProperty.prototype.put_ImageUrl = function (sUrl, nState) {
		if (!this.DivId){
			return;
		}
		let Api = Asc.editor;

		Api._addImageUrl([sUrl], this.getParentFields());
	};
	asc_CButtonFieldProperty.prototype.showFileDialog = function (nState) {
		if (!this.DivId){
			return;
		}
		let Api = Asc.editor;

		// set to field state to add image (will clear after set image)
		let aFields = this.getParentFields();
		let oDoc = Asc.editor.getPDFDoc();
		let oActionsQueue = oDoc.GetActionsQueue();

		if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]()) {
			window["AscDesktopEditor"]["OpenFilenameDialog"]("images", false, function(_file) {
				var file = _file;
				if (Array.isArray(file))
					file = file[0];

				var _url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](file);
				Asc.editor._addImageUrl([AscCommon.g_oDocumentUrls.getImageUrl(_url)], aFields);
			});
		}
		else {
			AscCommon.ShowImageFileDialog(Api.documentId, Api.documentUserId, undefined, Api.documentShardKey, Api.documentWopiSrc, Api.documentUserSessionId, function(error, files) {
				if (error.canceled !== true) {
					Api._uploadCallback(error, files, aFields);
				}

				AscCommon.global_mouseEvent.UnLockMouse();

			}, function(error) {
				if (c_oAscError.ID.No !== error) {
					Api.sendEvent("asc_onError", error, c_oAscError.Level.NoCritical);
				}

				Api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
				AscCommon.global_mouseEvent.UnLockMouse();
			});
		}
	};
	asc_CButtonFieldProperty.prototype.compare = function (pr) {
		let aFullNames = this.parentFields.map(function(field) {
			return field.GetFullName();
		});

		let _t = this;
		pr.parentFields.forEach(function(field) {
			if (!aFullNames.includes(field.GetFullName())) {
				_t.parentFields.push(field);
			}
		});

		if (this.highlight !== pr.highlight) {
			this.highlight = null;
		}
		if (this.layout !== pr.layout) {
			this.layout = null;
		}
		if (this.scaleWhen !== pr.scaleWhen) {
			this.scaleWhen = null;
		}
		if (this.scaleHow !== pr.scaleHow) {
			this.scaleHow = null;
		}
		if (this.fitBounds !== pr.fitBounds) {
			this.fitBounds = null;
		}
		if (!this.iconPos || !pr.iconPos || this.iconPos.X !== pr.iconPos.X || this.iconPos.Y !== pr.iconPos.Y) {
			this.iconPos = null;
		}
		if (this.behavior !== pr.behavior) {
			this.behavior = null;
		}
		if (this.currentState !== pr.currentState) {
			this.currentState = null;
		}
		if (this.normalCaption !== pr.normalCaption) {
			this.normalCaption = null;
		}
		if (this.hoverCaption !== pr.hoverCaption) {
			this.hoverCaption = null;
		}
		if (this.downCaption !== pr.downCaption) {
			this.downCaption = null;
		}
		if (this.normalImage !== pr.normalImage) {
			this.normalImage = null;
		}
		if (this.hoverImage !== pr.hoverImage) {
			this.hoverImage = null;
		}
		if (this.downImage !== pr.downImage) {
			this.downImage = null;
		}
		if (this.radiosInUnison !== pr.radiosInUnison) {
			this.radiosInUnison = null;
		}
	};
	//////////////////////////////////////////////////////////////////
	///// Field Actions
	//////////////////////////////////////////////////////////////////
	function asc_CFieldActionsProperty() {
		this.mouseUp	= undefined;
		this.mouseDown	= undefined;
		this.mouseEnter	= undefined;
		this.mouseExit	= undefined;
		this.onFocus	= undefined;
		this.onBlur		= undefined;

		// text/combobox
		this.format		= undefined;
		this.keystroke	= undefined;
		this.validate	= undefined;
		this.calculate	= undefined;
	};

	asc_CFieldActionsProperty.prototype.asc_getMouseUp = function () {
		return this.mouseUp;
	};
	asc_CFieldActionsProperty.prototype.asc_putMouseUp = function (v) {
		this.mouseUp = v;
	};
	asc_CFieldActionsProperty.prototype.asc_getMouseDown = function () {
		return this.mouseDown;
	};
	asc_CFieldActionsProperty.prototype.asc_putMouseDown = function (v) {
		this.mouseDown = v;
	};
	asc_CFieldActionsProperty.prototype.asc_getMouseEnter = function () {
		return this.mouseEnter;
	};
	asc_CFieldActionsProperty.prototype.asc_putMouseEnter = function (v) {
		this.mouseEnter = v;
	};
	asc_CFieldActionsProperty.prototype.asc_getMouseExit = function () {
		return this.mouseExit;
	};
	asc_CFieldActionsProperty.prototype.asc_putMouseExit = function (v) {
		this.mouseExit = v;
	};
	asc_CFieldActionsProperty.prototype.asc_getOnFocus = function () {
		return this.onFocus;
	};
	asc_CFieldActionsProperty.prototype.asc_putOnFocus = function (v) {
		this.onFocus = v;
	};
	asc_CFieldActionsProperty.prototype.asc_getOnBlur = function () {
		return this.onBlur;
	};
	asc_CFieldActionsProperty.prototype.asc_putOnBlur = function (v) {
		this.onBlur = v;
	};

	asc_CFieldActionsProperty.prototype.asc_getFormat = function () {
		return this.format;
	};
	asc_CFieldActionsProperty.prototype.asc_putFormat = function (v) {
		this.format = v;
	};
	asc_CFieldActionsProperty.prototype.asc_getKeystroke = function () {
		return this.keystroke;
	};
	asc_CFieldActionsProperty.prototype.asc_putKeystroke = function (v) {
		this.keystroke = v;
	};
	asc_CFieldActionsProperty.prototype.asc_getValidate = function () {
		return this.validate;
	};
	asc_CFieldActionsProperty.prototype.asc_putValidate = function (v) {
		this.validate = v;
	};
	asc_CFieldActionsProperty.prototype.asc_getCalculate = function () {
		return this.calculate;
	};
	asc_CFieldActionsProperty.prototype.asc_putCalculate = function (v) {
		this.calculate = v;
	};
	asc_CFieldActionsProperty.prototype.compare = function (pr) {
		if (!this.mouseUp.isEqual(pr.mouseUp)) {
			this.mouseUp = null;
		}
		if (!this.mouseDown.isEqual(pr.mouseDown)) {
			this.mouseDown = null;
		}
		if (!this.mouseEnter.isEqual(pr.mouseEnter)) {
			this.mouseEnter = null;
		}
		if (!this.mouseExit.isEqual(pr.mouseExit)) {
			this.mouseExit = null;
		}
		if (!this.onFocus.isEqual(pr.onFocus)) {
			this.onFocus = null;
		}
		if (!this.onBlur.isEqual(pr.onBlur)) {
			this.onBlur = null;
		}

		if (!this.format.isEqual(pr.format)) {
			this.format = null;
		}
		if (!this.keystroke.isEqual(pr.keystroke)) {
			this.keystroke = null;
		}
		if (!this.validate.isEqual(pr.validate)) {
			this.validate = null;
		}
		if (!this.calculate.isEqual(pr.calculate)) {
			this.calculate = null;
		}
	};

	
	//////////////////////////////////////////////////////////////////
	///// Pdf action collection
	//////////////////////////////////////////////////////////////////
	function asc_CPdfActionCollectionProperty() {
		this.actions = undefined;
	};

	asc_CPdfActionCollectionProperty.prototype.asc_getActions = function () {
		return this.actions;
	};
	asc_CPdfActionCollectionProperty.prototype.asc_putActions = function (v) {
		this.actions = v;
	};
	asc_CPdfActionCollectionProperty.prototype.isEqual = function (v) {
		if (!v || !Array.isArray(this.actions) || !Array.isArray(v.actions)) return false;
		if (this.actions.length !== v.actions.length) return false;

		for (let i = 0; i < this.actions.length; i++) {
			const a = this.actions[i];
			const b = v.actions[i];

			if (!a.isEqual(b)) {
				return false;
			}
		}

		return true;
	};
	asc_CPdfActionCollectionProperty.prototype.getJsonActionInfo = function (v) {
		return this.actions.map(function(action) {
			return action.getJsonActionInfo();
		});
	};

	//////////////////////////////////////////////////////////////////
	///// Pdf JS action
	//////////////////////////////////////////////////////////////////
	function asc_CPdfActionJsProperty() {
		this.type	= AscPDF.ACTIONS_TYPES.JavaScript;
		this.script	= undefined;
	};

	asc_CPdfActionJsProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CPdfActionJsProperty.prototype.asc_getScript = function () {
		return this.script;
	};
	asc_CPdfActionJsProperty.prototype.asc_putScript = function (v) {
		this.script = v;
	};
	asc_CPdfActionJsProperty.prototype.isEqual = function (v) {
		return this.script == v;
	};
	asc_CPdfActionJsProperty.prototype.getJsonActionInfo = function (v) {
		return {
			"S":  AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": this.script
		}
	};

	//////////////////////////////////////////////////////////////////
	///// Pdf reset action
	//////////////////////////////////////////////////////////////////
	function asc_CPdfActionResetProperty() {
		this.type			= AscPDF.ACTIONS_TYPES.ResetForm;
		this.isAllExcept	= undefined;
		this.names			= [];
	};

	asc_CPdfActionResetProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CPdfActionResetProperty.prototype.asc_getIsAllExcept = function () {
		return this.isAllExcept;
	};
	asc_CPdfActionResetProperty.prototype.asc_putIsAllExcept = function (v) {
		this.isAllExcept = v;
	};
	asc_CPdfActionResetProperty.prototype.asc_getNames = function () {
		return this.names;
	};
	asc_CPdfActionResetProperty.prototype.asc_putNames = function (v) {
		this.names = v;
	};
	asc_CPdfActionResetProperty.prototype.isEqual = function (v) {
		return AscCommon.isEqualSortedArrays(this.names, v) && this.isAllExcept == v.isAllExcept;
	};
	asc_CPdfActionResetProperty.prototype.getJsonActionInfo = function (v) {
		return {
			"S": AscPDF.ACTIONS_TYPES.ResetForm,
			"Fields": this.names.slice(),
			"Flags": Number(this.isAllExcept)
		}
	};

	//////////////////////////////////////////////////////////////////
	///// Pdf URI action
	//////////////////////////////////////////////////////////////////
	function asc_CPdfActionUriProperty() {
		this.type				= AscPDF.ACTIONS_TYPES.URI;
		this.uri				= undefined;
	};

	asc_CPdfActionUriProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CPdfActionUriProperty.prototype.asc_getUri = function () {
		return this.uri;
	};
	asc_CPdfActionUriProperty.prototype.asc_putUri = function (v) {
		this.uri = v;
	};
	asc_CPdfActionUriProperty.prototype.isEqual = function (v) {
		return this.uri == v.uri;
	};
	asc_CPdfActionUriProperty.prototype.getJsonActionInfo = function (v) {
		return {
			"S": AscPDF.ACTIONS_TYPES.URI,
			"URI": this.uri
		}
	};

	//////////////////////////////////////////////////////////////////
	///// Pdf GoTo action
	//////////////////////////////////////////////////////////////////
	function asc_CPdfActionGoToProperty() {
		this.type		= AscPDF.ACTIONS_TYPES.GoTo;
		this.goToType	= undefined;
		this.page		= undefined;
		this.zoom		= undefined;
		this.rect		= undefined;
	};

	asc_CPdfActionGoToProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CPdfActionGoToProperty.prototype.asc_getGoToType = function () {
		return this.goToType;
	};
	asc_CPdfActionGoToProperty.prototype.asc_putGoToType = function (v) {
		this.goToType = v;
	};
	asc_CPdfActionGoToProperty.prototype.asc_getPage = function () {
		return this.page;
	};
	asc_CPdfActionGoToProperty.prototype.asc_putPage = function (v) {
		this.page = v;
	};
	asc_CPdfActionGoToProperty.prototype.asc_getZoom = function () {
		return this.zoom;
	};
	asc_CPdfActionGoToProperty.prototype.asc_putZoom = function (v) {
		this.zoom = v;
	};
	asc_CPdfActionGoToProperty.prototype.asc_getRect = function () {
		return this.rect;
	};
	asc_CPdfActionGoToProperty.prototype.asc_putRect = function (v) {
		this.rect = v;
	};
	asc_CPdfActionGoToProperty.prototype.isEqual = function (v) {
		return this.page == v.page && this.goToType == v.goToType
		&& this.zoom == v.zoom && this.top == v.top && this.right == v.right
		&& this.bottom == v.bottom && this.left == v.left;
	};
	asc_CPdfActionGoToProperty.prototype.getJsonActionInfo = function (v) {
		return {
			"S": AscPDF.ACTIONS_TYPES.GoTo,
			"page": this.page,
			"kind": this.goToType,
			"zoom": this.zoom,

			"top": this.rect["top"],
			"right": this.rect["right"],
			"bottom": this.rect["bottom"],
			"left": this.rect["left"],
		}
	};

	//////////////////////////////////////////////////////////////////
	///// Pdf HideShow action
	//////////////////////////////////////////////////////////////////
	function asc_CPdfActionHideShowProperty() {
		this.type	= AscPDF.ACTIONS_TYPES.HideShow;
		this.isHide	= undefined;
		this.names	= undefined;
	};

	asc_CPdfActionHideShowProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CPdfActionHideShowProperty.prototype.asc_getIsHide = function () {
		return this.isHide;
	};
	asc_CPdfActionHideShowProperty.prototype.asc_putIsHide = function (v) {
		this.isHide = v;
	};
	asc_CPdfActionHideShowProperty.prototype.asc_getNames = function () {
		return this.names;
	};
	asc_CPdfActionHideShowProperty.prototype.asc_putNames = function (v) {
		this.names = v;
	};
	asc_CPdfActionHideShowProperty.prototype.isEqual = function (v) {
		return AscCommon.isEqualSortedArrays(this.names, v.names) && this.isHide == v.isHide;
	};
	asc_CPdfActionHideShowProperty.prototype.getJsonActionInfo = function (v) {
		return {
			"S": AscPDF.ACTIONS_TYPES.HideShow,
			"H": this.isHide,
			"T": this.names.slice()
		}
	};

	//////////////////////////////////////////////////////////////////
	///// Pdf Named action
	//////////////////////////////////////////////////////////////////
	function asc_CPdfActionNamedProperty() {
		this.type = AscPDF.ACTIONS_TYPES.Named;
		this.name = undefined;
	};

	asc_CPdfActionNamedProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CPdfActionNamedProperty.prototype.asc_getName = function () {
		return this.name;
	};
	asc_CPdfActionNamedProperty.prototype.asc_putName = function (v) {
		this.name = v;
	};
	asc_CPdfActionNamedProperty.prototype.isEqual = function (v) {
		return this.name == v.name;
	};
	asc_CPdfActionNamedProperty.prototype.getJsonActionInfo = function (v) {
		return {
			"S": AscPDF.ACTIONS_TYPES.Named,
			"N": this.name
		}
	};

	//////////////////////////////////////////////////////////////////
	///// Number format
	//////////////////////////////////////////////////////////////////
	function asc_CFieldNumberFormatProperty() {
		this.type				= AscPDF.FormatType.NUMBER;
		this.decimals			= undefined;
		this.sepStyle			= undefined;
		this.negStyle			= undefined;
		this.currency			= undefined;
		this.currencyPrepend	= undefined;
	};

	asc_CFieldNumberFormatProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_getDecimals = function () {
		return this.decimals;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_putDecimals = function (v) {
		this.decimals = v;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_getSepStyle = function () {
		return this.sepStyle;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_putSepStyle = function (v) {
		this.sepStyle = v;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_getNegStyle = function () {
		return this.negStyle;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_putNegStyle = function (v) {
		this.negStyle = v;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_getCurrency = function () {
		return this.currency;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_putCurrency = function (v) {
		this.currency = v;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_getCurrencyPrepend = function () {
		return this.currencyPrepend;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_putCurrencyPrepend = function (v) {
		this.currencyPrepend = v;
	};
	asc_CFieldNumberFormatProperty.prototype.isEqual = function (pr) {
		if (!pr) {
			return false;
		}
		
		if (this.decimals !== pr.decimals) {
			return false;
		}
		if (this.sepStyle !== pr.sepStyle) {
			return false;
		}
		if (this.negStyle !== pr.negStyle) {
			return false;
		}
		if (this.currency !== pr.currency) {
			return false;
		}
		if (this.currencyPrepend !== pr.currencyPrepend) {
			return false;
		}

		return true;
	};
	asc_CFieldNumberFormatProperty.prototype.getJsonActionInfo = function (isKeystroke) {
		const decimals = this.decimals;
		const sepStyle = this.sepStyle;
		const negStyle = this.negStyle;
		const currStyle = 0;
		const currency = this.currency;
		const currencyPrepend = this.currencyPrepend;

		return [{
			S: 14,
			JS: (isKeystroke ? 'AFNumber_Keystroke(' : 'AFNumber_Format(') +
				decimals + ', ' +
				sepStyle + ', ' +
				negStyle + ', ' +
				currStyle + ', ' +
				JSON.stringify(currency) + ', ' +
				currencyPrepend +
			');'
		}];
	};

	//////////////////////////////////////////////////////////////////
	///// Percentage format
	//////////////////////////////////////////////////////////////////
	function asc_CFieldPercentageFormatProperty() {
		this.type		= AscPDF.FormatType.PERCENTAGE;
		this.decimals	= undefined;
		this.sepStyle	= undefined;
	};

	asc_CFieldPercentageFormatProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldPercentageFormatProperty.prototype.asc_getDecimals = function () {
		return this.decimals;
	};
	asc_CFieldPercentageFormatProperty.prototype.asc_putDecimals = function (v) {
		this.decimals = v;
	};
	asc_CFieldPercentageFormatProperty.prototype.asc_getSepStyle = function () {
		return this.sepStyle;
	};
	asc_CFieldPercentageFormatProperty.prototype.asc_putSepStyle = function (v) {
		this.sepStyle = v;
	};
	asc_CFieldPercentageFormatProperty.prototype.isEqual = function (pr) {
		if (!pr) {
			return false;
		}

		if (this.decimals !== pr.decimals) {
			return false;
		}
		if (this.sepStyle !== pr.sepStyle) {
			return false;
		}

		return true;
	};
	asc_CFieldPercentageFormatProperty.prototype.getJsonActionInfo = function (isKeystroke) {
		const decimals = this.decimals;
		const sepStyle = this.sepStyle;

		return [{
			S: 14,
			JS: (isKeystroke ? 'AFPercent_Keystroke(' : 'AFPercent_Format(') +
				decimals + ', ' +
				sepStyle +
			');'
		}];
	};

	//////////////////////////////////////////////////////////////////
	///// Date format
	//////////////////////////////////////////////////////////////////
	function asc_CFieldDateFormatProperty() {
		this.type		= AscPDF.FormatType.DATE;
		this.format		= undefined;
	};

	asc_CFieldDateFormatProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldDateFormatProperty.prototype.asc_getFormat = function () {
		return this.format;
	};
	asc_CFieldDateFormatProperty.prototype.asc_putFormat = function (v) {
		this.format = v;
	};
	asc_CFieldDateFormatProperty.prototype.isEqual = function (pr) {
		if (!pr) {
			return false;
		}

		if (this.format !== pr.format) {
			return false;
		}

		return true;
	};
	asc_CFieldDateFormatProperty.prototype.getJsonActionInfo = function (isKeystroke) {
		const format = this.format;

		return [{
			S: 14,
			JS: (isKeystroke ? 'AFDate_Keystroke(' : 'AFDate_Format(') + 
				JSON.stringify(format) + ');'
		}];
	};
	//////////////////////////////////////////////////////////////////
	///// Time format
	//////////////////////////////////////////////////////////////////
	function asc_CFieldTimeFormatProperty() {
		this.type		= AscPDF.FormatType.TIME;
		this.format		= undefined;
	};

	asc_CFieldTimeFormatProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldTimeFormatProperty.prototype.asc_getFormat = function () {
		return this.format;
	};
	asc_CFieldTimeFormatProperty.prototype.asc_putFormat = function (v) {
		this.format = v;
	};
	asc_CFieldTimeFormatProperty.prototype.isEqual = function (pr) {
		if (!pr) {
			return false;
		}

		if (this.format !== pr.format) {
			return false;
		}

		return true;
	};
	asc_CFieldTimeFormatProperty.prototype.getJsonActionInfo = function (isKeystroke) {
		const format = this.format;

		return [{
			S: 14,
			JS: (isKeystroke ? 'AFTime_Keystroke(' : 'AFTime_Format(') +
				format + ');'
		}];
	};

	//////////////////////////////////////////////////////////////////
	///// Special format
	//////////////////////////////////////////////////////////////////
	function asc_CFieldSpecialFormatProperty() {
		this.type		= AscPDF.FormatType.SPECIAL;
		this.format		= undefined;
		this.mask		= undefined;
	};

	asc_CFieldSpecialFormatProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldSpecialFormatProperty.prototype.asc_getFormat = function () {
		return this.format;
	};
	asc_CFieldSpecialFormatProperty.prototype.asc_putFormat = function (v) {
		this.format = v;
	};
	asc_CFieldSpecialFormatProperty.prototype.asc_getMask = function () {
		return this.mask;
	};
	asc_CFieldSpecialFormatProperty.prototype.asc_putMask = function (v) {
		this.mask = v;
	};
	asc_CFieldSpecialFormatProperty.prototype.isEqual = function (pr) {
		if (!pr) {
			return false;
		}

		if (this.format !== pr.format) {
			return false;
		}
		if (this.mask !== pr.mask) {
			return false;
		}

		return true;
	};
	asc_CFieldSpecialFormatProperty.prototype.getJsonActionInfo = function (isKeystroke) {
		const format = this.format;
		const mask = this.mask;

		if (mask !== undefined) {
			return [{
				S: 14,
				JS: 'AFSpecial_KeystrokeEx(' + mask + ');'
			}];
		}

		return [{
			S: 14,
			JS: (isKeystroke ? 'AFSpecial_Keystroke(' : 'AFSpecial_Format(') +
				format + ');'
		}];
	};

	//////////////////////////////////////////////////////////////////
	///// Custom format
	//////////////////////////////////////////////////////////////////
	function asc_CFieldCustomFormatProperty() {
		this.type		= AscPDF.FormatType.CUSTOM;
		this.script		= undefined;
	};

	asc_CFieldCustomFormatProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldCustomFormatProperty.prototype.asc_getScript = function () {
		return this.script;
	};
	asc_CFieldCustomFormatProperty.prototype.asc_putScript = function (v) {
		this.script = v;
	};
	asc_CFieldCustomFormatProperty.prototype.isEqual = function (pr) {
		if (!pr) {
			return false;
		}

		if (this.script !== pr.script) {
			return false;
		}

		return true;
	};
	asc_CFieldCustomFormatProperty.prototype.getJsonActionInfo = function () {
		const script = this.script;
		if (!script) {
			return [];
		}

		return [{
			S: 14,
			JS: script
		}];
	};

	//////////////////////////////////////////////////////////////////
	///// Validate
	//////////////////////////////////////////////////////////////////
	function asc_CFieldValidateProperty() {
		this.greaterThen	= undefined;
		this.lessThen		= undefined;
		this.script			= undefined;
	};

	asc_CFieldValidateProperty.prototype.asc_getGreaterThen = function () {
		return this.greaterThen;
	};
	asc_CFieldValidateProperty.prototype.asc_putGreaterThen = function (v) {
		this.greaterThen = v;
	};
	asc_CFieldValidateProperty.prototype.asc_getLessThen = function () {
		return this.lessThen;
	};
	asc_CFieldValidateProperty.prototype.asc_putLessThen = function (v) {
		this.lessThen = v;
	};
	asc_CFieldValidateProperty.prototype.asc_getScript = function () {
		return this.script;
	};
	asc_CFieldValidateProperty.prototype.asc_putScript = function (v) {
		this.script = v;
	};
	asc_CFieldValidateProperty.prototype.isEqual = function (pr) {
		if (!pr) {
			return false;
		}

		if (this.greaterThen !== pr.greaterThen) {
			return false;
		}
		if (this.lessThen !== pr.lessThen) {
			return false;
		}
		if (this.script !== pr.script) {
			return false;
		}

		return true;
	};
	asc_CFieldValidateProperty.prototype.getJsonActionInfo = function () {
		let script = this.script;

		if (!script) {
			if (this.lessThen == undefined && this.greaterThen == undefined) {
				return [];
			}

			const hasGreater = this.greaterThen !== undefined && this.greaterThen !== null;
			const hasLess = this.lessThen !== undefined && this.lessThen !== null;

			const greater = hasGreater ? this.greaterThen : 0;
			const less = hasLess ? this.lessThen : 0;

			script = 'AFRange_Validate(' +
				hasGreater + ', ' +
				greater + ', ' +
				hasLess + ', ' +
				less +
			');';
		}

		return [{
			S: 14,
			JS: script
		}];
	};

	//////////////////////////////////////////////////////////////////
	///// Calculate
	//////////////////////////////////////////////////////////////////
	function asc_CFieldCalculateProperty() {
		this.type			= undefined;
		this.names			= undefined;
		this.script			= undefined;
	};

	asc_CFieldCalculateProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldCalculateProperty.prototype.asc_putType = function (v) {
		this.type = v;
	};
	asc_CFieldCalculateProperty.prototype.asc_getNames = function () {
		return this.names;
	};
	asc_CFieldCalculateProperty.prototype.asc_putNames = function (v) {
		this.names = v;
	};
	asc_CFieldCalculateProperty.prototype.asc_getScript = function () {
		return this.script;
	};
	asc_CFieldCalculateProperty.prototype.asc_putScript = function (v) {
		this.script = v;
	};
	asc_CFieldCalculateProperty.prototype.isEqual = function (pr) {
		if (!pr) {
			return false;
		}

		if (this.type != pr.type) {
			return false;
		}
		if (AscCommon.isEqualSortedArrays(this.names, pr.names)) {
			return false;
		}
		if (this.script !== pr.script) {
			return false;
		}

		return true;
	};
	asc_CFieldCalculateProperty.prototype.getJsonActionInfo = function () {
		let script = this.script;

		if (!script) {
			const type = this.type;
			const names = this.names;

			script = 'AFSimple_Calculate(' +
				JSON.stringify(type) + ', ' +
				JSON.stringify(names) +
			');';
		}

		return [{
			S: 14,
			JS: script
		}];
	};

	/** @constructor */
	function asc_CPdfPageProperty() {
		this.deleteLock	= false;
		this.rotateLock	= false;
		this.editLock	= false;
	}

	asc_CPdfPageProperty.prototype.constructor = asc_CPdfPageProperty;
	asc_CPdfPageProperty.prototype.asc_getDeleteLock = function () {
		return this.deleteLock;
	};
	asc_CPdfPageProperty.prototype.asc_putDeleteLock = function (v) {
		this.deleteLock = v;
	};
	asc_CPdfPageProperty.prototype.asc_getRotateLock = function () {
		return this.rotateLock;
	};
	asc_CPdfPageProperty.prototype.asc_putRotateLock = function (v) {
		this.rotateLock = v;
	};
	asc_CPdfPageProperty.prototype.asc_getEditLock = function () {
		return this.editLock;
	};
	asc_CPdfPageProperty.prototype.asc_putEditLock = function (v) {
		this.editLock = v;
	};

	window["Asc"]["asc_CAnnotProperty"] = window["Asc"].asc_CAnnotProperty = asc_CAnnotProperty;
	prot = asc_CAnnotProperty.prototype;
	prot["asc_getIds"]				= prot.asc_getIds;
	prot["asc_putIds"]				= prot.asc_putIds;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_putType"]				= prot.asc_putType;
	prot["asc_getFill"]				= prot.asc_getFill;
	prot["asc_putFill"]				= prot.asc_putFill;
	prot["asc_getStroke"]			= prot.asc_getStroke;
	prot["asc_getOpacity"]			= prot.asc_getOpacity;
	prot["asc_putOpacity"]			= prot.asc_putOpacity;
	prot["asc_putStroke"]			= prot.asc_putStroke;
	prot["asc_getSubject"]			= prot.asc_getSubject;
	prot["asc_putSubject"]			= prot.asc_putSubject;
	prot["asc_getAnnotProps"]		= prot.asc_getAnnotProps;
	prot["asc_putAnnotProps"]		= prot.asc_putAnnotProps;

	window["Asc"]["asc_CFreeTextAnnotProperty"] = window["Asc"].asc_CFreeTextAnnotProperty = asc_CFreeTextAnnotProperty;
	prot = asc_CFreeTextAnnotProperty.prototype;
	prot["asc_getBorderWidth"]	= prot.asc_getBorderWidth;
	prot["asc_putBorderWidth"]	= prot.asc_putBorderWidth;
	prot["asc_getLineEnd"]		= prot.asc_getLineEnd;
	prot["asc_putLineEnd"]		= prot.asc_putLineEnd;
	prot["asc_getBorderStyle"]	= prot.asc_getBorderStyle;
	prot["asc_putBorderStyle"]	= prot.asc_putBorderStyle;
	prot["asc_getCanEditText"]	= prot.asc_getCanEditText;
	prot["asc_putCanEditText"]	= prot.asc_putCanEditText;

	window["Asc"]["asc_CLineAnnotProperty"] = window["Asc"].asc_CLineAnnotProperty = asc_CLineAnnotProperty;
	prot = asc_CLineAnnotProperty.prototype;
	prot["asc_getBorderStyle"]	= prot.asc_getBorderStyle;
	prot["asc_putBorderStyle"]	= prot.asc_putBorderStyle;
	prot["asc_getBorderWidth"]	= prot.asc_getBorderWidth;
	prot["asc_putBorderWidth"]	= prot.asc_putBorderWidth;
	prot["asc_getLineEnd"]		= prot.asc_getLineEnd;
	prot["asc_putLineEnd"]		= prot.asc_putLineEnd;
	prot["asc_getLineStart"]	= prot.asc_getLineStart;
	prot["asc_putLineStart"]	= prot.asc_putLineStart;
	prot["asc_getCanEditText"]	= prot.asc_getCanEditText;
	prot["asc_putCanEditText"]	= prot.asc_putCanEditText;

	window["Asc"]["asc_CPolyLineAnnotProperty"] = window["Asc"].asc_CPolyLineAnnotProperty = asc_CPolyLineAnnotProperty;
	prot = asc_CPolyLineAnnotProperty.prototype;
	prot["asc_getBorderStyle"]	= prot.asc_getBorderStyle;
	prot["asc_putBorderStyle"]	= prot.asc_putBorderStyle;
	prot["asc_getBorderWidth"]	= prot.asc_getBorderWidth;
	prot["asc_putBorderWidth"]	= prot.asc_putBorderWidth;
	prot["asc_getLineEnd"]		= prot.asc_getLineEnd;
	prot["asc_putLineEnd"]		= prot.asc_putLineEnd;
	prot["asc_getLineStart"]	= prot.asc_getLineStart;
	prot["asc_putLineStart"]	= prot.asc_putLineStart;

	window["Asc"]["asc_CClosedAnnotProperty"] = window["Asc"].asc_CClosedAnnotProperty = asc_CClosedAnnotProperty;
	prot = asc_CClosedAnnotProperty.prototype;
	prot["asc_getBorderStyle"]	= prot.asc_getBorderStyle;
	prot["asc_putBorderStyle"]	= prot.asc_putBorderStyle;
	prot["asc_getBorderWidth"]	= prot.asc_getBorderWidth;
	prot["asc_putBorderWidth"]	= prot.asc_putBorderWidth;

	window["Asc"]["asc_CBaseFieldProperty"] = window["Asc"].asc_CBaseFieldProperty = asc_CBaseFieldProperty;
	prot = asc_CBaseFieldProperty.prototype;
	prot["asc_getType"]			= prot.asc_getType;
	prot["asc_putType"]			= prot.asc_putType;
	prot["asc_getName"]			= prot.asc_getName;
	prot["asc_putName"]			= prot.asc_putName;
	prot["asc_getRequired"]		= prot.asc_getRequired;
	prot["asc_putRequired"]		= prot.asc_putRequired;
	prot["asc_getReadOnly"]		= prot.asc_getReadOnly;
	prot["asc_putReadOnly"]		= prot.asc_putReadOnly;
	prot["asc_getRot"]			= prot.asc_getRot;
	prot["asc_putRot"]			= prot.asc_putRot;
	prot["asc_getDisplay"]		= prot.asc_getDisplay;
	prot["asc_putDisplay"]		= prot.asc_putDisplay;
	prot["asc_getFill"]			= prot.asc_getFill;
	prot["asc_putFill"]			= prot.asc_putFill;
	prot["asc_getStroke"]		= prot.asc_getStroke;
	prot["asc_putStroke"]		= prot.asc_putStroke;
	prot["asc_getStrokeWidth"]	= prot.asc_getStrokeWidth;
	prot["asc_putStrokeWidth"]	= prot.asc_putStrokeWidth;
	prot["asc_getStrokeStyle"]	= prot.asc_getStrokeStyle;
	prot["asc_putStrokeStyle"]	= prot.asc_putStrokeStyle;
	prot["asc_getPropLocked"]	= prot.asc_getPropLocked;
	prot["asc_putPropLocked"]	= prot.asc_putPropLocked;
	prot["asc_getTooltip"]		= prot.asc_getTooltip;
	prot["asc_putTooltip"]		= prot.asc_putTooltip;
	prot["asc_getDigitsType"]	= prot.asc_getDigitsType;
	prot["asc_putDigitsType"]	= prot.asc_putDigitsType;
	prot["get_Locked"]			= prot.get_Locked;
	prot["put_Locked"]			= prot.put_Locked;
	prot["asc_getActionsProps"]	= prot.asc_getActionsProps;
	prot["asc_putActionsProps"]	= prot.asc_putActionsProps;
	prot["asc_getFieldProps"]	= prot.asc_getFieldProps;
	prot["asc_putFieldProps"]	= prot.asc_putFieldProps;

	window["Asc"]["asc_CTextFieldProperty"] = window["Asc"].asc_CTextFieldProperty = asc_CTextFieldProperty;
	prot = asc_CTextFieldProperty.prototype;
	prot["asc_getDefaultValue"]			= prot.asc_getDefaultValue;
	prot["asc_putDefaultValue"]			= prot.asc_putDefaultValue;
	prot["asc_getMultiline"]			= prot.asc_getMultiline;
	prot["asc_putMultiline"]			= prot.asc_putMultiline;
	prot["asc_getScrollLongText"]		= prot.asc_getScrollLongText;
	prot["asc_putScrollLongText"]		= prot.asc_putScrollLongText;
	prot["asc_getCharLimit"]			= prot.asc_getCharLimit;
	prot["asc_putCharLimit"]			= prot.asc_putCharLimit;
	prot["asc_getComb"]					= prot.asc_getComb;
	prot["asc_putComb"]					= prot.asc_putComb;
	prot["asc_getPlaceholder"]			= prot.asc_getPlaceholder;
	prot["asc_putPlaceholder"]			= prot.asc_putPlaceholder;
	prot["asc_getAutoFit"]				= prot.asc_getAutoFit;
	prot["asc_putAutoFit"]				= prot.asc_putAutoFit;
	prot["asc_getPassword"]				= prot.asc_getPassword;
	prot["asc_putPassword"]				= prot.asc_putPassword;

	window["Asc"]["asc_CComboboxFieldProperty"] = window["Asc"].asc_CComboboxFieldProperty = asc_CComboboxFieldProperty;
	prot = asc_CComboboxFieldProperty.prototype;
	prot["asc_getOptions"]				= prot.asc_getOptions;
	prot["asc_putOptions"]				= prot.asc_putOptions;
	prot["asc_getCommitOnSelChange"]	= prot.asc_getCommitOnSelChange;
	prot["asc_putCommitOnSelChange"]	= prot.asc_putCommitOnSelChange;
	prot["asc_getEditable"]				= prot.asc_getEditable;
	prot["asc_putEditable"]				= prot.asc_putEditable;
	prot["asc_getPlaceholder"]			= prot.asc_getPlaceholder;
	prot["asc_putPlaceholder"]			= prot.asc_putPlaceholder;
	prot["asc_getAutoFit"]				= prot.asc_getAutoFit;
	prot["asc_putAutoFit"]				= prot.asc_putAutoFit;

	window["Asc"]["asc_CListboxFieldProperty"] = window["Asc"].asc_CListboxFieldProperty = asc_CListboxFieldProperty;
	prot = asc_CListboxFieldProperty.prototype;
	prot["asc_getOptions"]				= prot.asc_getOptions;
	prot["asc_putOptions"]				= prot.asc_putOptions;
	prot["asc_getCommitOnSelChange"]	= prot.asc_getCommitOnSelChange;
	prot["asc_putCommitOnSelChange"]	= prot.asc_putCommitOnSelChange;
	prot["asc_getMultipleSelection"]	= prot.asc_getMultipleSelection;
	prot["asc_putMultipleSelection"]	= prot.asc_putMultipleSelection;

	window["Asc"]["asc_CCheckboxFieldProperty"] = window["Asc"].asc_CCheckboxFieldProperty = asc_CCheckboxFieldProperty;
	prot = asc_CCheckboxFieldProperty.prototype;
	prot["asc_getCheckboxStyle"]		= prot.asc_getCheckboxStyle;
	prot["asc_putCheckboxStyle"]		= prot.asc_putCheckboxStyle;
	prot["asc_getExportValue"]		= prot.asc_getExportValue;
	prot["asc_putExportValue"]		= prot.asc_putExportValue;
	prot["asc_getDefaultChecked"]	= prot.asc_getDefaultChecked;
	prot["asc_putDefaultChecked"]	= prot.asc_putDefaultChecked;
	prot["asc_getToggleToOff"]		= prot.asc_getToggleToOff;
	prot["asc_putToggleToOff"]		= prot.asc_putToggleToOff;

	window["Asc"]["asc_CRadiobuttonFieldProperty"] = window["Asc"].asc_CRadiobuttonFieldProperty = asc_CRadiobuttonFieldProperty;
	prot = asc_CRadiobuttonFieldProperty.prototype;
	prot["asc_getRadiosInUnison"]	= prot.asc_getRadiosInUnison;
	prot["asc_putRadiosInUnison"]	= prot.asc_putRadiosInUnison;

	window["Asc"]["asc_CButtonFieldProperty"] = window["Asc"].asc_CButtonFieldProperty = asc_CButtonFieldProperty;
	prot = asc_CButtonFieldProperty.prototype;
	prot["asc_getHighlight"]	= prot.asc_getHighlight;
	prot["asc_putHighlight"]	= prot.asc_putHighlight;
	prot["asc_getLayout"]		= prot.asc_getLayout;
	prot["asc_putLayout"]		= prot.asc_putLayout;
	prot["asc_getScaleWhen"]	= prot.asc_getScaleWhen;
	prot["asc_putScaleWhen"]	= prot.asc_putScaleWhen;
	prot["asc_getScaleHow"]		= prot.asc_getScaleHow;
	prot["asc_putScaleHow"]		= prot.asc_putScaleHow;
	prot["asc_getFitBounds"]	= prot.asc_getFitBounds;
	prot["asc_putFitBounds"]	= prot.asc_putFitBounds;
	prot["asc_getIconPos"]		= prot.asc_getIconPos;
	prot["asc_putIconPos"]		= prot.asc_putIconPos;
	prot["asc_getBehavior"]		= prot.asc_getBehavior;
	prot["asc_putBehavior"]		= prot.asc_putBehavior;
	prot["asc_getNormalCaption"]= prot.asc_getNormalCaption;
	prot["asc_putNormalCaption"]= prot.asc_putNormalCaption;
	prot["asc_getNormalImage"]	= prot.asc_getNormalImage;
	prot["asc_putNormalImage"]	= prot.asc_putNormalImage;
	prot["asc_getHoverCaption"]	= prot.asc_getHoverCaption;
	prot["asc_putHoverCaption"]	= prot.asc_putHoverCaption;
	prot["asc_getHoverImage"]	= prot.asc_getHoverImage;
	prot["asc_putHoverImage"]	= prot.asc_putHoverImage;
	prot["asc_getDownCaption"]	= prot.asc_getDownCaption;
	prot["asc_putDownCaption"]	= prot.asc_putDownCaption;
	prot["asc_getDownImage"]	= prot.asc_getDownImage;
	prot["asc_putDownImage"]	= prot.asc_putDownImage;
	prot["asc_putCurrentState"]	= prot.asc_putCurrentState;
	prot["asc_getCurrentState"]	= prot.asc_getCurrentState;
	prot["put_DivId"]			= prot.put_DivId;
	prot["drawTexture"]			= prot.drawTexture;
	prot["put_ImageUrl"]		= prot.put_ImageUrl;
	prot["showFileDialog"]		= prot.showFileDialog;

	window["Asc"]["asc_CFieldActionsProperty"] = window["Asc"].asc_CFieldActionsProperty = asc_CFieldActionsProperty;
	prot = asc_CFieldActionsProperty.prototype;
	prot["asc_getFormat"]		= prot.asc_getFormat;
	prot["asc_putFormat"]		= prot.asc_putFormat;
	prot["asc_getKeystroke"]	= prot.asc_getKeystroke;
	prot["asc_putKeystroke"]	= prot.asc_putKeystroke;
	prot["asc_getValidate"]		= prot.asc_getValidate;
	prot["asc_putValidate"]		= prot.asc_putValidate;
	prot["asc_getCalculate"]	= prot.asc_getCalculate;
	prot["asc_putCalculate"]	= prot.asc_putCalculate;

	window["Asc"]["asc_CPdfActionCollectionProperty"] = window["Asc"].asc_CPdfActionCollectionProperty = asc_CPdfActionCollectionProperty;
	prot = asc_CPdfActionCollectionProperty.prototype;
	prot["asc_getActions"]		= prot.asc_getActions;
	prot["asc_putActions"]		= prot.asc_putActions;

	window["Asc"]["asc_CPdfActionJsProperty"] = window["Asc"].asc_CPdfActionJsProperty = asc_CPdfActionJsProperty;
	prot = asc_CPdfActionJsProperty.prototype;
	prot["asc_getType"]			= prot.asc_getType;
	prot["asc_getScript"]		= prot.asc_getScript;
	prot["asc_putScript"]		= prot.asc_putScript;

	window["Asc"]["asc_CPdfActionResetProperty"] = window["Asc"].asc_CPdfActionResetProperty = asc_CPdfActionResetProperty;
	prot = asc_CPdfActionResetProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getIsAllExcept"]		= prot.asc_getIsAllExcept;
	prot["asc_putIsAllExcept"]		= prot.asc_putIsAllExcept;
	prot["asc_getNames"]			= prot.asc_getNames;
	prot["asc_putNames"]			= prot.asc_putNames;

	window["Asc"]["asc_CPdfActionUriProperty"] = window["Asc"].asc_CPdfActionUriProperty = asc_CPdfActionUriProperty;
	prot = asc_CPdfActionUriProperty.prototype;
	prot["asc_getType"]			= prot.asc_getType;
	prot["asc_getUri"]			= prot.asc_getUri;
	prot["asc_putUri"]			= prot.asc_putUri;

	window["Asc"]["asc_CPdfActionGoToProperty"] = window["Asc"].asc_CPdfActionGoToProperty = asc_CPdfActionGoToProperty;
	prot = asc_CPdfActionGoToProperty.prototype;
	prot["asc_getType"]			= prot.asc_getType;
	prot["asc_getGoToType"]		= prot.asc_getGoToType;
	prot["asc_putGoToType"]		= prot.asc_putGoToType;
	prot["asc_getPage"]			= prot.asc_getPage;
	prot["asc_putPage"]			= prot.asc_putPage;
	prot["asc_getZoom"]			= prot.asc_getZoom;
	prot["asc_putZoom"]			= prot.asc_putZoom;
	prot["asc_getRect"]			= prot.asc_getRect;
	prot["asc_putRect"]			= prot.asc_putRect;

	window["Asc"]["asc_CPdfActionHideShowProperty"] = window["Asc"].asc_CPdfActionHideShowProperty = asc_CPdfActionHideShowProperty;
	prot = asc_CPdfActionHideShowProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getIsHide"]			= prot.asc_getIsHide;
	prot["asc_putIsHide"]			= prot.asc_putIsHide;
	prot["asc_getNames"]			= prot.asc_getNames;
	prot["asc_putNames"]			= prot.asc_putNames;

	window["Asc"]["asc_CPdfActionNamedProperty"] = window["Asc"].asc_CPdfActionNamedProperty = asc_CPdfActionNamedProperty;
	prot = asc_CPdfActionNamedProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getName"]				= prot.asc_getName;
	prot["asc_putName"]				= prot.asc_putName;

	window["Asc"]["asc_CFieldNumberFormatProperty"] = window["Asc"].asc_CFieldNumberFormatProperty = asc_CFieldNumberFormatProperty;
	prot = asc_CFieldNumberFormatProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getDecimals"]			= prot.asc_getDecimals;
	prot["asc_putDecimals"]			= prot.asc_putDecimals;
	prot["asc_getSepStyle"]			= prot.asc_getSepStyle;
	prot["asc_putSepStyle"]			= prot.asc_putSepStyle;
	prot["asc_getNegStyle"]			= prot.asc_getNegStyle;
	prot["asc_putNegStyle"]			= prot.asc_putNegStyle;
	prot["asc_getCurrency"]			= prot.asc_getCurrency;
	prot["asc_putCurrency"]			= prot.asc_putCurrency;
	prot["asc_getCurrencyPrepend"]	= prot.asc_getCurrencyPrepend;
	prot["asc_putCurrencyPrepend"]	= prot.asc_putCurrencyPrepend;


	window["Asc"]["asc_CFieldPercentageFormatProperty"] = window["Asc"].asc_CFieldPercentageFormatProperty = asc_CFieldPercentageFormatProperty;
	prot = asc_CFieldPercentageFormatProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getDecimals"]			= prot.asc_getDecimals;
	prot["asc_putDecimals"]			= prot.asc_putDecimals;
	prot["asc_getSepStyle"]			= prot.asc_getSepStyle;
	prot["asc_putSepStyle"]			= prot.asc_putSepStyle;

	window["Asc"]["asc_CFieldDateFormatProperty"] = window["Asc"].asc_CFieldDateFormatProperty = asc_CFieldDateFormatProperty;
	prot = asc_CFieldDateFormatProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getFormat"]			= prot.asc_getFormat;
	prot["asc_putFormat"]			= prot.asc_putFormat;

	window["Asc"]["asc_CFieldTimeFormatProperty"] = window["Asc"].asc_CFieldTimeFormatProperty = asc_CFieldTimeFormatProperty;
	prot = asc_CFieldTimeFormatProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getFormat"]			= prot.asc_getFormat;
	prot["asc_putFormat"]			= prot.asc_putFormat;

	window["Asc"]["asc_CFieldSpecialFormatProperty"] = window["Asc"].asc_CFieldSpecialFormatProperty = asc_CFieldSpecialFormatProperty;
	prot = asc_CFieldSpecialFormatProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getFormat"]			= prot.asc_getFormat;
	prot["asc_putFormat"]			= prot.asc_putFormat;
	prot["asc_getMask"]				= prot.asc_getMask;
	prot["asc_putMask"]				= prot.asc_putMask;

	window["Asc"]["asc_CFieldCustomFormatProperty"] = window["Asc"].asc_CFieldCustomFormatProperty = asc_CFieldCustomFormatProperty;
	prot = asc_CFieldCustomFormatProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getScript"]			= prot.asc_getScript;
	prot["asc_putScript"]			= prot.asc_putScript;

	window["Asc"]["asc_CFieldValidateProperty"] = window["Asc"].asc_CFieldValidateProperty = asc_CFieldValidateProperty;
	prot = asc_CFieldValidateProperty.prototype;
	prot["asc_getGreaterThen"]		= prot.asc_getGreaterThen;
	prot["asc_putGreaterThen"]		= prot.asc_putGreaterThen;
	prot["asc_getLessThen"]			= prot.asc_getLessThen;
	prot["asc_putLessThen"]			= prot.asc_putLessThen;
	prot["asc_getScript"]			= prot.asc_getScript;
	prot["asc_putScript"]			= prot.asc_putScript;

	window["Asc"]["asc_CFieldCalculateProperty"] = window["Asc"].asc_CFieldCalculateProperty = asc_CFieldCalculateProperty;
	prot = asc_CFieldCalculateProperty.prototype;
	prot["asc_getType"]			= prot.asc_getType;
	prot["asc_putType"]			= prot.asc_putType;
	prot["asc_getNames"]		= prot.asc_getNames;
	prot["asc_putNames"]		= prot.asc_putNames;
	prot["asc_getScript"]		= prot.asc_getScript;
	prot["asc_putScript"]		= prot.asc_putScript;

	window["Asc"]["asc_CPdfPageProperty"] = window["Asc"].asc_CPdfPageProperty = asc_CPdfPageProperty;
	prot = asc_CPdfPageProperty.prototype;
	prot["asc_getDeleteLock"]	= prot.asc_getDeleteLock;
	prot["asc_putDeleteLock"]	= prot.asc_putDeleteLock;
	prot["asc_getRotateLock"]	= prot.asc_getRotateLock;
	prot["asc_putRotateLock"]	= prot.asc_putRotateLock;
	prot["asc_getEditLock"]		= prot.asc_getEditLock;
	prot["asc_putEditLock"]		= prot.asc_putEditLock;
	
})(window);

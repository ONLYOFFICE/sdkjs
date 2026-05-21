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

(function(){

	/**
	 * Controls how the icon is scaled (if necessary) to fit inside the button face. The convenience scaleHow object defines all
	 * of the valid alternatives:
	 * @typedef {Object} scaleHow
	 * @property {number} proportional
	 * @property {number} anamorphic
	 */

	/**
	 * Controls when an icon is scaled to fit inside the button face. The convenience scaleWhen object defines all of the valid
	 * alternatives:
	 * @typedef {Object} scaleWhen
	 * @property {number} always
	 * @property {number} never
	 * @property {number} tooBig
	 * @property {number} tooSmall
	 */

	//------------------------------------------------------------------------------------------------------------------
	//
	// Internal
	//
	//------------------------------------------------------------------------------------------------------------------

	
	// types without source object
	let ALIGN_TYPE = {
		left:   "left",
		center: "center",
		right:  "right"
	}

	let LINE_WIDTH = {
		"none":   0,
		"thin":   1,
		"medium": 2,
		"thick":  3
	}

	const highlight     = AscPDF.Api.Types.highlight;
	const style         = AscPDF.Api.Types.style;
	const display       = AscPDF.Api.Types.display;
	const border        = AscPDF.Api.Types.border;
	const color         = AscPDF.Api.Types.color;

	/**
	 * A string that sets the trigger for the action. Values are:
	 * @typedef {"MouseUp" | "MouseDown" | "MouseEnter" | "MouseExit" | "OnFocus" | "OnBlur" | "Keystroke" | "Validate" | "Calculate" | "Format"} cTrigger
	 * For a list box, use the Keystroke trigger for the Selection Change event.
	 */

	const ERROR_SET_MESSAGES = {
		logic: "InvalidSetError: The field is not a logical root.",
		noProp: "InvalidSetError: field doesn't have this prop.",
		invalidParam: "InvalidSetError: Set not possible, invalid parameter.",
		actionInProgress: "InvalidSetError: Set not possible, action is in progress."
	}

	const ERROR_GET_MESSAGES = {
		logic: "InvalidGetError: Get not possible, The field is not a logical root.",
		noProp: "InvalidGetError: field doesn't have this prop.",
		invalidParam: "InvalidGetError: Get not possible, invalid parameter.",
	}
	
	// main class (this) in JS PDF scripts
	function ApiDocument(oDoc) {
		this.doc = oDoc;
	};

	/**
	 * Returns an interactive field by name.
	 * @memberof ApiDocument
	 * @param {string} sName - field name.
	 * @typeofeditors ["PDF"]
	 * @returns {ApiBaseField}
	 */
	ApiDocument.prototype.getField = function(sName) {
		let oField = this.doc.GetField(sName);
		if (oField)
			return oField.GetFormApi();

		return null;        
	};

	/**
	 * The base file name, with extension, of the document referenced by the Doc
	 * @memberof ApiDocument
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiDocument.prototype, "documentFileName", {
		get: function() {
			return Asc.editor.documentTitle;
		}
	});

	// base form class with attributes and method for all types of forms
	function ApiBaseField(oField)
	{
		this.field = oField;
	}

	/**
	 * The border style for a field. Valid border styles are solid/dashed/beveled/inset/underline.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "borderStyle", {
		set: function(sValue) {
			if (this.field.IsLogicalRoot()) {
				if (Object.values(border).includes(sValue)) {
					let aFields = this.field.GetAllWidgets();
					aFields.forEach(function(field) {
						field.SetBorderStyle(private_GetIntBorderStyle(sValue));
					});

					return this["borderStyle"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return private_GetStrBorderStyle(this.field.GetKid(0).GetBorderStyle());
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		}
	});

	/**
	 * The default value of a field—that is, the value that the field is set to when the form is reset. For combo boxes and list
	 * boxes, either an export or a user value can be used to set the default. A conflict can occur, for example, when the field has
	 * an export value and a user value with the same value but these apply to different items in the list. In such cases, the
	 * export value is matched against first.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "defaultValue", {
		set: function(value) {
			if (this.field.IsLogicalRoot()) {
				if (value && value.toString) {
					value = value.toString();
					this.field.SetDefaultValue(value);

					return this["defaultValue"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.GetDefaultValue();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Controls whether the field is hidden or visible on screen and in print. The values for the display property are listed in
	 * the table below.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "display", {
		set: function(nType) {
			if (this.field.IsLogicalRoot()) {
				if (Object.values(display).includes(nType)) {
					let aFields = this.field.GetAllWidgets();
					aFields.forEach(function(field) {
						field.SetDisplay(nType);
					});

					return this["display"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.GetKid(0).GetDisplay();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Note: This property has been superseded by the display property and its use is discouraged.
	 * If the value is false, the field is visible to the user; if true, the field is invisible. The default value is false.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "hidden", {
		set: function(bValue) {
			if (this.field.IsLogicalRoot()) {
				if (typeof bValue == "boolean") {
					let aFields = this.field.GetAllWidgets();
					aFields.forEach(function(field) {
						field.SetDisplay(bValue ? display["hidden"] : display["visible"]);
					});

					return this["hidden"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.GetKid(0).GetDisplay() == display["hidden"];
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Specifies the background color for a field. The background color is used to fill the rectangle of the field. Values are
	 * defined by using transparent, gray, RGB or CMYK color. See Color arrays for information on defining color arrays and
	 * how values are used with this property.
	 * In older versions of this specification, this property was named bgColor. The use of bgColor is now discouraged,
	 * although it is still valid for backward compatibility.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "fillColor", {
		set: function(value) {
			if (this.field.IsLogicalRoot()) {
				if (Array.isArray(value)) {
					let aFields = this.field.GetAllWidgets();
					let aColor  = private_correctApiColor(value).slice(1);
					aFields.forEach(function(field) {
						field.SetBackgroundColor(aColor);
					});

					return this["fillColor"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return private_getApiColor(this.field.GetKid(0).GetBackgroundColor());
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Specifies the background color for a field. The background color is used to fill the rectangle of the field. Values are
	 * defined by using transparent, gray, RGB or CMYK color. See Color arrays for information on defining color arrays and
	 * how values are used with this property.
	 * Note: The use of bgColor is now discouraged,
	 * although it is still valid for backward compatibility.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "bgColor", {
		set: function(value) {
			this["fillColor"] = value;
		},
		get: function() {
			return this["fillColor"];
		}
	});

	/**
	 * Returns the Doc of the document to which the field belongs.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "doc", {
		get: function() {
			return this.field.GetDocument().GetDocumentApi();
		}
	});

	/**
	 * Specifies the thickness of the border when stroking the perimeter of a field’s rectangle. If the stroke color is transparent,
	 * this parameter has no effect except in the case of a beveled border. Values are:
	 * 0 — none
	 * 1 — thin
	 * 2 — medium
	 * 3 — thick
	 * In older versions of this specification, this property was borderWidth. The use of borderWidth is now discouraged,
	 * although it is still valid for backward compatibility.
	 * The default value for lineWidth is 1 (thin). Any integer value can be used; however, values beyond 5 may distort the
	 * field’s appearance
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "lineWidth", {
		set: function(nValue) {
			if (this.field.IsLogicalRoot()) {
				nValue = parseInt(nValue);
				if (Object.values(LINE_WIDTH).includes(nValue)) {
					let aFields = this.field.GetAllWidgets();
					aFields.forEach(function(field) {
						field.SetBorderWidth(nValue);
					});

					return this["lineWidth"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.GetKid(0).GetBorderWidth();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	Object.defineProperty(ApiBaseField.prototype, "borderWidth", {
		set: function(nValue) {
			this["lineWidth"] = nValue;
		},
		get: function() {
			return this["lineWidth"];
		}
	});

	/**
	 * This property returns the fully qualified field name of the field as a string object.
	 * Beginning with Acrobat 6.0, if the Field object represents one individual widget, the returned name includes an
	 * appended '.' followed by the widget index.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "name", {
		get: function() {
			return this.field.GetFullName();
		}
	});

	/**
	 * The page number or an array of page numbers of a field. If the field has only one appearance in the document, the page
	 * property returns an integer representing the 0-based page number of the page on which the field appears. If the field
	 * has multiple appearances, it returns an array of integers, each member of which is a 0-based page number of an
	 * appearance of the field. The order in which the page numbers appear in the array is determined by the order in which
	 * the individual widgets of this field were created (and is unaffected by tab-order). If an appearance of the field is on a
	 * hidden template page, page returns a value of -1 for that appearance.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "page", {
		get: function() {
			if (this.field.IsWidget()) {
				return this.field.GetPage();
			}
			else {
				let aFields = this.field.GetAllWidgets();
				let aPages = aFields.map(function(field) {
					return field.GetPage();
				})

				return aPages;
			}
		}
	});

	/**
	 * The read-only characteristic of a field. If a field is read-only, the user can see the field but cannot change it.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "readonly", {
		set: function(bValue) {
			if (this.field.IsLogicalRoot()) {
				if (typeof(bValue) == "boolean") {
					this.field.SetReadOnly(bValue);
					
					return this["readonly"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.IsReadOnly();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * An array of four numbers in rotated user space that specify the size and placement of the form field. These four numbers
	 * are the coordinates of the bounding rectangle and are listed in the following order: upper-left x, upper-left y, lower-right
	 * x and lower-right y.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "rect", {
		set: function(aRect) {
			if (!private_IsValidRect(aRect)) {
				throw Error(ERROR_SET_MESSAGES.invalidParam);
			}

			if (this.field.IsLogicalRoot()) {
				this.field.GetKid(0).SetRect(aRect);
				
				return this["rect"];
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.GetKid(0).GetRect();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * An array of four numbers in rotated user space that specify the size and placement of the form field. These four numbers
	 * are the coordinates of the bounding rectangle and are listed in the following order: upper-left x, upper-left y, lower-right
	 * x and lower-right y.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "required", {
		set: function(bValue) {
			if (this.field.IsLogicalRoot()) {
				if (typeof(bValue) == "boolean") {
					if (this.field.GetType() == AscPDF.FIELD_TYPES.button) {
						throw Error(ERROR_SET_MESSAGES.noProp);
					}

					this.field.SetRequired(bValue);

					return this["required"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.GetType() != AscPDF.FIELD_TYPES.button) {
				if (this.field.IsLogicalRoot()) {
					return this.field.IsRequired();
				}
				else {
					throw Error(ERROR_GET_MESSAGES.logic);
				}
			}
			else {
				throw Error(ERROR_GET_MESSAGES.noProp);
			}
		}
	});

	/**
	 * Specifies the stroke color for a field that is used to stroke the rectangle of the field with a line as large as the line width.
	 * Values are defined by using transparent, gray, RGB or CMYK color. See Color arrays for information on defining color
	 * arrays and how values are used with this property.
	 * In older versions of this specification, this property was borderColor. The use of borderColor is now discouraged,
	 * although it is still valid for backward compatibility.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "strokeColor", {
		set: function(value) {
			if (this.field.IsLogicalRoot()) {
				if (Array.isArray(value)) {
					let aFields = this.field.GetAllWidgets();
					let aColor  = private_correctApiColor(value).slice(1);
					aFields.forEach(function(field) {
						field.SetBorderColor(aColor);
					});

					return this["strokeColor"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return private_getApiColor(this.field.GetKid(0).GetBorderColor());
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	Object.defineProperty(ApiBaseField.prototype, "borderColor", {
		set: function(value) {
			this["strokeColor"] = value;
		},
		get: function() {
			return this["strokeColor"];
		}
	});

	/**
	 * The foreground color of a field. It represents the text color for text, button, or list box fields and the check color for check
	 * box or radio button fields. Values are defined the same as the fillColor. See Color arrays for information on defining
	 * color arrays and how values are set and used with this property.
	 * In older versions of this specification, this property was fgColor. The use of fgColor is now discouraged, although it is
	 * still valid for backward compatibility.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "textColor", {
		set: function(aColor) {
			if (this.field.IsLogicalRoot()) {
				if (Array.isArray(aColor)) {
					let aFields = this.field.GetAllWidgets();
					aFields.forEach(function(field) {
						field.SetTextColor(aColor.slice(1));
					});

					return this["textColor"];
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return private_getApiColor(this.field.GetKid(0).GetTextColor());
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	Object.defineProperty(ApiBaseField.prototype, "fgColor", {
		set: function(aColor) {
			this["textColor"] = aColor;
		},
		get: function() {
			return this["textColor"];
		}
	});

	/**
	 * Specifies the text size (in points) to be used in all controls. In check box and radio button fields, the text size determines
	 * the size of the check. Valid text sizes range from 0 to 32767, inclusive. A value of zero means the largest point size that
	 * allows all text data to fit in the field’s rectangle.
	 * @memberof ApiBaseField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseField.prototype, "textSize", {
		set: function(nValue) {
			if (this.field.IsLogicalRoot()) {
				if (typeof(nValue) == "number" && nValue >= 0 && nValue < AscPDF.MAX_TEXT_SIZE) {
					let aFields = this.field.GetAllWidgets();
					aFields.forEach(function(field) {
						field.SetTextSize(Math.round(nValue));
					});

					return this["textSize"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.GetKid(0).GetTextSize();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Sets the JavaScript action of the field for a given trigger.
	 * Note: This method will overwrite any action already defined for the chosen trigger.
	 * @memberof ApiBaseField
	 * @param {cTrigger} cTrigger - A string that sets the trigger for the action.
	 * @param {string} cScript - The JavaScript code to be executed when the trigger is activated.
	 * @typeofeditors ["PDF"]
	 */
	ApiBaseField.prototype.setAction = function(cTrigger, cScript) {
		let aFields = this.field.GetAllWidgets();
		let nInternalType;
		switch (cTrigger) {
			case "MouseUp":
				nInternalType = AscPDF.PDF_TRIGGERS_TYPES.MouseUp;
				break;
			case "MouseDown":
				nInternalType = AscPDF.PDF_TRIGGERS_TYPES.MouseDown;
				break;
			case "MouseEnter":
				nInternalType = AscPDF.PDF_TRIGGERS_TYPES.MouseEnter;
				break;
			case "MouseExit":
				nInternalType = AscPDF.PDF_TRIGGERS_TYPES.MouseExit;
				break;
			case "OnFocus":
				nInternalType = AscPDF.PDF_TRIGGERS_TYPES.OnFocus;
				break;
			case "OnBlur":
				nInternalType = AscPDF.PDF_TRIGGERS_TYPES.OnBlur;
				break;
			case "Keystroke":
				nInternalType = AscPDF.PDF_TRIGGERS_TYPES.Keystroke;
				break;
			case "Validate":
				nInternalType = AscPDF.PDF_TRIGGERS_TYPES.Validate;
				break;
			case "Calculate":
				nInternalType = AscPDF.PDF_TRIGGERS_TYPES.Calculate;
				break;
			case "Format":
				nInternalType = AscPDF.PDF_TRIGGERS_TYPES.Format;
				break;
		}

		if (nInternalType != null) {
			aFields.forEach(function(field) {
				field.SetActions(nInternalType, [{"S": AscPDF.ACTIONS_TYPES.JavaScript, "JS": cScript}]);
			});
		}
	};

	function ApiPushButtonField(oField)
	{
		ApiBaseField.call(this, oField);
	}
	ApiPushButtonField.prototype = Object.create(ApiBaseField.prototype);
	ApiPushButtonField.prototype.constructor = ApiPushButtonField;

	/**
	 * Controls how space is distributed from the left of the button face with respect to the icon. It is expressed as a percentage
	 * between 0 and 100, inclusive. The default value is 50.
	 * If the icon is scaled anamorphically (which results in no space differences), this property is not used.
	 * @memberof ApiPushButtonField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiPushButtonField.prototype, "buttonAlignX", {
		set: function(nValue) {
			if (this.field.IsLogicalRoot()) {
				if (typeof(nValue) == "number") {
					nValue = Math.round(nValue);
					let aFields = this.field.GetAllWidgets();
					aFields.forEach(function(field) {
						field.SetIconPosition(nValue, field.GetIconPosition().Y);
					});

					return this["buttonAlignX"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.GetKid(0).GetIconPosition().X;
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});
	/**
	 * Controls how unused space is distributed from the bottom of the button face with respect to the icon. It is expressed as a
	 * percentage between 0 and 100, inclusive. The default value is 50.
	 * If the icon is scaled anamorphically (which results in no space differences), this property is not used.
	 * @memberof ApiPushButtonField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiPushButtonField.prototype, "buttonAlignY", {
		set: function(nValue) {
			if (this.field.IsLogicalRoot()) {
				if (typeof(nValue) == "number") {
					nValue = Math.round(nValue);
					let aFields = this.field.GetAllWidgets();
					aFields.forEach(function(field) {
						field.SetIconPosition(field.GetIconPosition().X, nValue);
					});

					return this["buttonAlignY"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.GetKid(0).GetIconPosition().Y;
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	ApiPushButtonField.prototype.buttonImportIcon = function() {
		this.field.buttonImportIcon();
	};

	function ApiBaseCheckBoxField(oField)
	{
		ApiBaseField.call(this, oField);
	}
	
	ApiBaseCheckBoxField.prototype = Object.create(ApiBaseField.prototype);
	ApiBaseCheckBoxField.prototype.constructor = ApiBaseCheckBoxField;

	/**
	 * An array of strings representing the export values for the field. The array has as many elements as there are annotations
	 * in the field. The elements are mapped to the annotations in the order of creation (unaffected by tab-order).
	 * For radio button fields, this property is required to make the field work properly as a group. The button that is checked at any time gives its value to the field as a whole.
	 * For check box fields, unless an export value is specified, “Yes” (or the corresponding localized string) is the default when the field is checked. “Off” is the default when the field is unchecked (the same as for a radio button field when none of its buttons are checked).
	 * @memberof ApiBaseCheckBoxField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseCheckBoxField.prototype, "exportValues", {
		set: function(arrValues) {
			if (this.field.IsLogicalRoot()) {
				if (Array.isArray(arrValues)) {
					this.field.SetOptions(arrValues);

					return this["exportValues"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.GetKids().map(function(field) { return field.GetExportValue()});
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Allows the user to set the glyph style of a check box or radio button. The glyph style is the graphic used to indicate that
	 * the item has been selected.
	 * The style values are associated with keywords as follows:
	 * check    - style.ch
	 * cross    - style.cr
	 * diamond  - style.di
	 * circle   - style.ci
	 * star     - style.st
	 * square   - style.sq
	 * @memberof ApiBaseCheckBoxField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseCheckBoxField.prototype, "style", {
		set: function(sStyle) {
			if (this.field.IsLogicalRoot()) {
				if (Object.values(style).includes(sStyle)) {
					let aFields = this.field.GetAllWidgets();
					aFields.forEach(function(field) {
						field.SetStyle(private_GetIntChStyle(sStyle));
					})

					return this["style"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return private_GetStrChStyle(this.field.GetKid(0).GetStyle());
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Determines whether the specified widget is checked.
	 * Note: For a set of radio buttons that do not have duplicate export values, you can get the value, which is equal to the
	 * export value of the individual widget that is currently checked (or returns an empty string, if none is).
	 * @memberof ApiBaseCheckBoxField
	 * @param {number} optionIdx - The 0-based index of an individual radio button or check box widget for this field.
	 * The index is determined by the order in which the individual widgets of this field
	 * were created (and is unaffected by tab-order).
	 * Every entry in the Fields panel has a suffix giving this index, for example, MyField #0.
	 * @typeofeditors ["PDF"]
	 * @returns {string}
	 */
	ApiBaseCheckBoxField.prototype.isBoxChecked = function(optionIdx) {
		if (this.field.IsLogicalRoot()) {
			let aFields = this.field.GetAllWidgets();
			let oField  = aFields[optionIdx];
			if (!oField) {
				throw Error(ERROR_GET_MESSAGES.invalidParam);
			}

			return oField.IsChecked();
		}
		else {
			throw Error(ERROR_GET_MESSAGES.logic);
		}
	};

	function ApiCheckBoxField(oField)
	{
		ApiBaseCheckBoxField.call(this, oField);
	}
	ApiCheckBoxField.prototype = Object.create(ApiBaseCheckBoxField.prototype);
	ApiCheckBoxField.prototype.constructor = ApiCheckBoxField;
	Object.defineProperties(ApiCheckBoxField.prototype, {
		"value": {
			set: function(value) {
				if (this.field.IsLogicalRoot()) {
					if (value != undefined) {
						value = String(value);

						let oDoc            = this.field.GetDocument();
						let oCalcInfo       = oDoc.GetCalculateInfo();
						let oSourceField    = oCalcInfo.GetSourceField();

						if (oCalcInfo.IsInProgress() && ((oSourceField && oSourceField.GetFullName() == this.field.GetFullName() && oCalcInfo.GetCurrentField() !== oSourceField) && oCalcInfo.GetCurrentField() !== oSourceField))
							throw Error(ERROR_SET_MESSAGES.actionInProgress);
						if (oDoc.isOnValidate)
							throw Error(ERROR_SET_MESSAGES.actionInProgress);

						let sApiValueToSet = value;
						let aOpt = this.field.GetOptions();
						if (aOpt) {
							let nIdx = aOpt.indexOf(value);
							if (nIdx != -1)
								sApiValueToSet = String(nIdx);
						}

						let oWidget = this.field.GetKid(0);
						oWidget.SetValue(sApiValueToSet);
						oWidget.Commit();

						if (oCalcInfo.IsInProgress() == false && oDoc.IsNeedDoCalculate()) {
							oDoc.DoCalculateFields(this.field);
							oDoc.CommitFields();
						}

						return this["value"];
					}
					else {
						throw Error(ERROR_SET_MESSAGES.invalidParam);
					}
				}
				else {
					throw Error(ERROR_SET_MESSAGES.logic);
				}
			},
			get: function() {
				if (this.field.IsLogicalRoot()) {
					return this.field.GetLogicValue();
				}
				else {
					throw Error(ERROR_GET_MESSAGES.logic);
				}
			}
		}
	});

	function ApiRadioButtonField(oField)
	{
		ApiBaseCheckBoxField.call(this, oField);
	}
	ApiRadioButtonField.prototype = Object.create(ApiBaseCheckBoxField.prototype);
	ApiRadioButtonField.prototype.constructor = ApiRadioButtonField;

	/**
	 * If false, even if a group of radio buttons have the same name and export value, they behave in a mutually exclusive
	 * fashion, like HTML radio buttons. The default for new radio buttons is false.
	 * If true, if a group of radio buttons have the same name and export value, they turn on and off in unison, as in Acrobat
	 * 4.0.
	 * @memberof ApiRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiRadioButtonField.prototype, "radiosInUnison", {
		set: function(bValue) {
			if (this.field.IsLogicalRoot()) {
				if (typeof(bValue) == "boolean") {
					this.field.SetRadiosInUnison(bValue);
					this.field.UpdateAll();

					return this["radiosInUnison"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.IsRadiosInUnison();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	Object.defineProperties(ApiRadioButtonField.prototype, {
		"value": {
			set: function(sValue) {
				if (this.field.IsLogicalRoot()) {
					if (sValue != undefined) {
						sValue = String(sValue);

						let oDoc            = this.field.GetDocument();
						let oCalcInfo       = oDoc.GetCalculateInfo();
						let oSourceField    = oCalcInfo.GetSourceField();
	
						if (oCalcInfo.IsInProgress() && (oSourceField && oSourceField.GetFullName() == this.field.GetFullName() && oCalcInfo.GetCurrentField() !== oSourceField))
							throw Error(ERROR_SET_MESSAGES.actionInProgress);
						if (oDoc.isOnValidate)
							throw Error(ERROR_SET_MESSAGES.actionInProgress);
	
						
						let sApiValueToSet = sValue;
						let aOpt = this.field.GetOptions();
						if (aOpt) {
							let nIdx = aOpt.indexOf(sValue);
							if (nIdx != -1)
								sApiValueToSet = String(nIdx);
						}

						let oWidget = this.field.GetKid(0);
						oWidget.SetValue(sApiValueToSet);
						oWidget.Commit();
	
						if (oCalcInfo.IsInProgress() == false && oDoc.IsNeedDoCalculate()) {
							oDoc.DoCalculateFields(this.field);
							oDoc.CommitFields();
						}

						return this["value"];
					}
					else {
						throw Error(ERROR_SET_MESSAGES.invalidParam);
					}
				}
				else {
					throw Error(ERROR_SET_MESSAGES.logic);
				}
			},
			get: function() {
				if (this.field.IsLogicalRoot()) {
					let aOpt = this.field.GetOptions();
					if (aOpt) {
						return aOpt[this.field.GetLogicValue()];
					}
					else {
						return this.field.GetLogicValue();
					}
				}
				else {
					throw Error(ERROR_GET_MESSAGES.logic);
				}
			}
		}
	});

	function ApiTextField(oField)
	{
		ApiBaseField.call(this, oField);
	}
	ApiTextField.prototype = Object.create(ApiBaseField.prototype);
	ApiTextField.prototype.constructor = ApiTextField;

	/**
	 * Controls how the text is laid out within the text field. Values are left/center/right.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiTextField.prototype, "alignment", {
		set: function(sValue) {
			if (this.field.IsLogicalRoot()) {
				if (Object.values(ALIGN_TYPE).includes(sValue)) {
					let aFields = this.field.GetAllWidgets();
					var nJcType = private_GetIntAlign(sValue);
					aFields.forEach(function(field) {
						field.SetAlign(nJcType);
					});

					return this["alignment"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return private_GetStrAlign(this.field.GetKid(0).GetAlign());
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Changes the calculation order of fields in the document. When a computable text or combo box field is added to a
	 * document, the field’s name is appended to the calculation order array. The calculation order array determines in what
	 * order the fields are calculated. The calcOrderIndex property works similarly to the Calculate tab used by the Acrobat
	 * Form tool.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiTextField.prototype, "calcOrderIndex", {
		set: function(nValue) {
			if (this.field.IsLogicalRoot()) {
				let nIdx = parseInt(nValue);
				if (isNaN(parseInt(nIdx)) == false) {
					this.field.SetCalcOrderIndex(nIdx);

					return this["calcOrderIndex"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.GetCalcOrderIndex();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Limits the number of characters that a user can type into a text field.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiTextField.prototype, "charLimit", {
		set: function(nValue) {
			if (this.field.IsLogicalRoot()) {
				let nCharLimit = parseInt(nValue);
				if (isNaN(nCharLimit) == false) {
					this.field.SetCharLimit(nCharLimit);

					return this["charLimit"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.GetCharLimit();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * If set to true, the field background is drawn as series of boxes (one for each character in the value of the field) and each
	 * character of the content is drawn within those boxes. The number of boxes drawn is determined from the charLimit
	 * property.
	 * It applies only to text fields. The setter will also raise if any of the following field properties are also set multiline,
	 * password, and fileSelect. A side-effect of setting this property is that the doNotScroll property is also set.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiTextField.prototype, "comb", {
		set: function(bValue) {
			if (this.field.IsLogicalRoot()) {
				if (typeof bValue == "boolean") {
					if (bValue) {
						this.field.SetCharLimit(10);
					}
					else {
						this.field.SetCharLimit(0);
					}

					field.SetComb(bValue);

					return this["comb"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.IsComb();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * If true, the text field does not scroll and the user, therefore, is limited by the rectangular region designed for the field.
	 * Setting this property to true or false corresponds to checking or unchecking the Scroll Long Text field in the Options
	 * tab of the field.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiTextField.prototype, "doNotScroll", {
		set: function(bDoNot) {
			if (this.field.IsLogicalRoot()) {
				if (typeof bDoNot == "boolean") {
					this.field.SetDoNotScroll(bDoNot);

					return this["doNotScroll"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.IsDoNotScroll();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * If true, spell checking is not performed on this editable text field. Setting this property to true or false corresponds
	 * to unchecking or checking the Check Spelling attribute in the Options tab of the Field Properties dialog box.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiTextField.prototype, "doNotSpellCheck", {
		set: function(bDoNot) {
			if (this.field.IsLogicalRoot()) {
				if (typeof bDoNot == "boolean") {
					this.field.SetDoNotSpellCheck(bDoNot);

					return this["doNotSpellCheck"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.IsDoNotSpellCheck();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Controls how text is wrapped within the field. If false (the default), the text field can be a single line only. If true,
	 * multiple lines are allowed and they wrap to field boundaries.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiTextField.prototype, "multiline", {
		set: function(bValue) {
			if (this.field.IsLogicalRoot()) {
				if (typeof bValue == "boolean") {
					this.field.SetMultiline(bValue);

					return this["multiline"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.IsMultiline();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	Object.defineProperties(ApiTextField.prototype, {
		"value": {
			set: function(value) {
				if (this.field.IsLogicalRoot()) {
					if (value != undefined) {
						value = String(value);

						let oDoc            = this.field.GetDocument();
						let oCalcInfo       = oDoc.GetCalculateInfo();
						let oSourceField    = oCalcInfo.GetSourceField();

						if (oCalcInfo.IsInProgress() && (oSourceField && oSourceField.GetFullName() == this.field.GetFullName() && oCalcInfo.GetCurrentField() !== oSourceField))
							throw Error(ERROR_SET_MESSAGES.actionInProgress);
						if (oDoc.isOnValidate)
							throw Error(ERROR_SET_MESSAGES.actionInProgress);

						if (value != null && value.toString)
							value = value.toString();
							
						if (this["value"] == value)
							return;

						let oWidget = this.field.GetKid(0);
						let isCanCommit = oWidget.IsCanCommit(value);
						if (isCanCommit) {
							oWidget.SetValue(value);
							oWidget.Commit();
							if (oCalcInfo.IsInProgress() == false) {
								if (oDoc.IsNeedDoCalculate()) {
									oDoc.DoCalculateFields(this.field);
									oDoc.AddFieldToCommit(oWidget);
									oDoc.CommitFields();
								}
							}
						}

						return this["value"];
					}
					else {
						throw Error(ERROR_SET_MESSAGES.invalidParam);
					}
				}
				else {
					throw Error(ERROR_SET_MESSAGES.logic);
				}
			},
			get: function() {
				let value = this.field.GetLogicValue();
				let isNumber = /^[+-]?\d+(\.\d+)?$/.test(value);
				return isNumber ? parseFloat(value) : (value != undefined ? value : "");
			}
		},
	});
	
	function ApiBaseListField(oField)
	{
		ApiBaseField.call(this, oField);
	}
	ApiBaseListField.prototype = Object.create(ApiBaseField.prototype);
	ApiBaseListField.prototype.constructor = ApiBaseListField;

	/**
	 * Controls whether a field value is committed after a selection change:
	 * If true, the field value is committed immediately when the selection is made.
	 * If false, the user can change the selection multiple times without committing the field value. The value is
	 * committed only when the field loses focus, that is, when the user clicks outside the field.
	 * @memberof ApiBaseListField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiBaseListField.prototype, "commitOnSelChange", {
		set: function(bValue) {
			if (this.field.IsLogicalRoot()) {
				if (typeof bValue == "boolean") {
					this.field.SetCommitOnSelChange(bValue);

					return this["commitOnSelChange"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.IsCommitOnSelChange();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Gets the internal value of an item in a combo box or a list box.
	 * @memberof CTextField
	 * @param {number} nIdx - The 0-based index of the item in the list or -1 for the last item in the list.
	 * @param {boolean} [bExportValue=true] - Specifies whether to return an export value.
	 * @typeofeditors ["PDF"]
	 * @returns {string}
	 */
	ApiBaseListField.prototype.getItemAt = function(nIdx, bExportValue) {
		if (this.field.IsLogicalRoot()) {
			let aOptions = this.field.GetOptions();
			if (aOptions[nIdx]) {
				if (bExportValue == false) {
					return Array.isArray(aOptions[nIdx]) ? aOptions[nIdx][0] : aOptions[nIdx];
				}
				else {
					return Array.isArray(aOptions[nIdx]) ? aOptions[nIdx][1] : aOptions[nIdx];
				}
			}
			else {
				throw Error(ERROR_GET_MESSAGES.invalidParam);
			}
		}
		else {
			throw Error(ERROR_GET_MESSAGES.logic);
		}
	};

	function ApiComboBoxField(oField)
	{
		ApiBaseListField.call(this, oField);
	};
	ApiComboBoxField.prototype = Object.create(ApiBaseListField.prototype);
	ApiComboBoxField.prototype.constructor = ApiComboBoxField;

	/**
	 * Controls how the text is laid out within the text field. Values are left/center/right.
	 * @memberof ApiComboBoxField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiComboBoxField.prototype, "alignment", {
		set: function(sValue) {
			if (this.field.IsLogicalRoot()) {
				if (Object.values(ALIGN_TYPE).includes(sValue)) {
					let aFields = this.field.GetAllWidgets();
					var nJcType = private_GetIntAlign(sValue);
					aFields.forEach(function(field) {
						field.SetAlign(nJcType);
					});

					return this["alignment"];
				}
				else
					throw Error(ERROR_SET_MESSAGES.invalidParam);
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return private_GetStrAlign(this.field.GetKid(0).GetAlign());
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Changes the calculation order of fields in the document. When a computable text or combo box field is added to a
	 * document, the field’s name is appended to the calculation order array. The calculation order array determines in what
	 * order the fields are calculated. The calcOrderIndex property works similarly to the Calculate tab used by the Acrobat
	 * Form tool.
	 * @memberof ApiComboBoxField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiComboBoxField.prototype, "calcOrderIndex", {
		set: function(nValue) {
			if (this.field.IsLogicalRoot()) {
				let nIdx = parseInt(nValue);
				if (isNaN(parseInt(nIdx)) == false) {
					this.field.SetCalcOrderIndex(nIdx);

					return this["calcOrderIndex"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.GetCalcOrderIndex();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Reads and writes value index of a combo box.
	 * @memberof ApiComboBoxField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiComboBoxField.prototype, "currentValueIndices", {
		set: function(value) {
			if (this.field.IsLogicalRoot()) {
				let oDoc            = this.field.GetDocument();
				let oCalcInfo       = oDoc.GetCalculateInfo();
				let oSourceField    = oCalcInfo.GetSourceField();

				if (oCalcInfo.IsInProgress() && (oSourceField && oSourceField.GetFullName() == this.field.GetFullName() && oCalcInfo.GetCurrentField() !== oSourceField))
					throw Error(ERROR_SET_MESSAGES.actionInProgress);
				if (oDoc.isOnValidate)
					throw Error(ERROR_SET_MESSAGES.actionInProgress);

				let nIdx;
				if (Array.isArray(value)) {
					nIdx = parseInt(value[0]);
				}
				else {
					nIdx = parseInt(value);
				}

				if (isNaN(nIdx)) {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}

				if (this.field.GetLogicCurIdxs()[0] == nIdx)
					return;

				let oWidget = this.field.GetKid(0);
				let sDisplayValue = this.getItemAt(nIdx, false);
				let isCanCommit = oWidget.IsCanCommit(sDisplayValue);

				if (isCanCommit) {
					oWidget.SetCurIdxs([nIdx]);
					oWidget.Commit();
					if (oCalcInfo.IsInProgress() == false) {
						if (oDoc.IsNeedDoCalculate()) {
							oDoc.DoCalculateFields(this.field);
							oDoc.AddFieldToCommit(oWidget);
							oDoc.CommitFields();
						}
					}
				}

				return this["currentValueIndices"];
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				let aCurIdxs = this.field.GetLogicCurIdxs();
				return aCurIdxs[0] != undefined ? aCurIdxs[0] : -1;
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Controls whether a combo box is editable. If true, the user can type in a selection. If false, the user must choose one
	 * of the provided selections.
	 * @memberof ApiComboBoxField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiComboBoxField.prototype, "editable", {
		set: function(bValue) {
			if (this.field.IsLogicalRoot()) {
				if (typeof bValue == "boolean") {
					this.field.SetEditable(bValue);

					return this["editable"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.IsEditable();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	Object.defineProperties(ApiComboBoxField.prototype, {
		"value": {
			set: function(value) {
				if (this.field.IsLogicalRoot()) {
					if (value != undefined) {
						value = String(value);

						let oDoc = this.field.GetDocument();
						let oCalcInfo = oDoc.GetCalculateInfo();
						let oSourceField = oCalcInfo.GetSourceField();

						if (oCalcInfo.IsInProgress() && (oSourceField && oSourceField.GetFullName() == this.field.GetFullName() && oCalcInfo.GetCurrentField() !== oSourceField))
							throw Error(ERROR_SET_MESSAGES.actionInProgress);
						if (oDoc.isOnValidate)
							throw Error(ERROR_SET_MESSAGES.actionInProgress);

						if (value != null && value.toString)
							value = value.toString();
							
						if (this["value"] == value)
							return;
						
						let aOptions = this.field.GetOptions();
						let sDisplayValue;

						aOptions.forEach(function(option) {
							if (Array.isArray(option)) {
								if (option[1] == value) {
									sDisplayValue = option[0];
								}
							}
							else if (option == value) {
								sDisplayValue = value;
							}
						});

						if (sDisplayValue == undefined) {
							sDisplayValue = value;
						}

						let oWidget = this.field.GetKid(0);
						let isCanCommit = oWidget.IsCanCommit(value);
						if (isCanCommit) {
							oWidget.SetValue(value);
							oWidget.Commit();
							if (oCalcInfo.IsInProgress() == false) {
								if (oDoc.IsNeedDoCalculate()) {
									oDoc.DoCalculateFields(this.field);
									oDoc.AddFieldToCommit(oWidget);
									oDoc.CommitFields();
								}
							}
						}

						return this["value"];
					}
					else {
						throw Error(ERROR_SET_MESSAGES.invalidParam);
					}
				}
				else {
					throw Error(ERROR_SET_MESSAGES.logic);
				}
			},
			get: function() {
				let value = this.field.GetLogicValue();
				let isNumber = /^\d+(\.\d+)?$/.test(value);
				return isNumber ? parseFloat(value) : (value != undefined ? value : "");
			}
		}
	});

	/**
	 * If true, spell checking is not performed on this editable text field. Setting this property to true or false corresponds
	 * to unchecking or checking the Check Spelling attribute in the Options tab of the Field Properties dialog box.
	 * @memberof ApiComboBoxField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiComboBoxField.prototype, "doNotSpellCheck", {
		set: function(bDoNot) {
			if (this.field.IsLogicalRoot()) {
				if (typeof bDoNot == "boolean") {
					this.field.SetDoNotSpellCheck(bDoNot);

					return this["doNotSpellCheck"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.IsDoNotSpellCheck();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	function ApiListBoxField(oField)
	{
		ApiBaseListField.call(this, oField);
	};
	
	ApiListBoxField.prototype = Object.create(ApiBaseListField.prototype);
	ApiListBoxField.prototype.constructor = ApiListBoxField;

	/**
	 * Controls how the text is laid out within the text field. Values are left/center/right.
	 * @memberof ApiListBoxField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiListBoxField.prototype, "alignment", {
		set: function(sValue) {
			if (this.field.IsLogicalRoot()) {
				if (Object.values(ALIGN_TYPE).includes(sValue)) {
					let aFields = this.field.GetAllWidgets();
					var nJcType = private_GetIntAlign(sValue);
					aFields.forEach(function(field) {
						field.SetAlign(nJcType);
					});

					return this["alignment"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return private_GetStrAlign(this.field.GetKid(0).GetAlign());
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * Reads and writes single or multiply value index of a listbox.
	 * @memberof ApiListBoxField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiListBoxField.prototype, "currentValueIndices", {
		set: function(aIdxs) {
			if (this.field.IsLogicalRoot()) {
				let oDoc            = this.field.GetDocument();
				let oCalcInfo       = oDoc.GetCalculateInfo();
				let oSourceField    = oCalcInfo.GetSourceField();
				let oWidget         = this.field.GetKid(0);

				if (oCalcInfo.IsInProgress() && (oSourceField && oSourceField.GetFullName() == this.field.GetFullName() && oCalcInfo.GetCurrentField() !== oSourceField))
					throw Error(ERROR_SET_MESSAGES.actionInProgress);
				if (oDoc.isOnValidate)
					throw Error(ERROR_SET_MESSAGES.actionInProgress);
				if (Array.isArray(aIdxs) && this.field.IsMultipleSelection() === false)
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				
				if (Array.isArray(aIdxs))
					oWidget.SetCurIdxs(aIdxs);
				else
					oWidget.SetCurIdxs([aIdxs]);

				oWidget.Commit();
				if (oCalcInfo.IsInProgress() == false) {
					if (oDoc.IsNeedDoCalculate()) {
						oDoc.DoCalculateFields(this.field);
						oDoc.AddFieldToCommit(oWidget);
						oDoc.CommitFields();
					}
				}

				return this["currentValueIndices"];
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				let aCurIdxs    = this.field.GetLogicCurIdxs();
				if (this.field.IsMultipleSelection() == false) {
					return aCurIdxs[0] != undefined ? aCurIdxs[0] : -1;
				}
				else {
					return aCurIdxs.slice();
				}
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	/**
	 * If true, indicates that a list box allows a multiple selection of items.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDF"]
	 */
	Object.defineProperty(ApiListBoxField.prototype, "multipleSelection", {
		set: function(bValue) {
			if (this.field.IsLogicalRoot()) {
				if (typeof bValue == "boolean") {
					this.field.SetMultipleSelection(bValue);

					return this["multipleSelection"];
				}
				else {
					throw Error(ERROR_SET_MESSAGES.invalidParam);
				}
			}
			else {
				throw Error(ERROR_SET_MESSAGES.logic);
			}
		},
		get: function() {
			if (this.field.IsLogicalRoot()) {
				return this.field.IsMultipleSelection();
			}
			else {
				throw Error(ERROR_GET_MESSAGES.logic);
			}
		}
	});

	Object.defineProperties(ApiListBoxField.prototype, {
		"value": {
			set: function(value) {
				if (this.field.IsLogicalRoot()) {
					if (value != undefined) {
						value = String(value);

						let oDoc = this.field.GetDocument();
						let oCalcInfo = oDoc.GetCalculateInfo();
						let oSourceField = oCalcInfo.GetSourceField();

						if (oCalcInfo.IsInProgress() && (oSourceField && oSourceField.GetFullName() == this.field.GetFullName() && oCalcInfo.GetCurrentField() !== oSourceField))
							throw Error(ERROR_SET_MESSAGES.actionInProgress);
						if (oDoc.isOnValidate)
							throw Error(ERROR_SET_MESSAGES.actionInProgress);

						if (value != null && value.toString)
							value = value.toString();
						
						if (this["value"] == value)
							return;
						
						let oWidget = this.field.GetKid(0);
						oWidget.SetValue(value);
						oWidget.Commit();

						if (oCalcInfo.IsInProgress() == false && oDoc.IsNeedDoCalculate()) {
							oDoc.DoCalculateFields(this.field);
							oDoc.CommitFields();
						}

						return this["value"];
					}
					else {
						throw Error(ERROR_SET_MESSAGES.invalidParam);
					}
				}
				else {
					throw Error(ERROR_SET_MESSAGES.logic);
				}
			},
			get: function() {
				if (this.field.IsLogicalRoot()) {
					let value = this.field.GetLogicValue();
					let isNumber = /^\d+(\.\d+)?$/.test(value);
					return isNumber ? parseFloat(value) : (value != undefined ? value : "");
				}
				else {
					throw Error(ERROR_GET_MESSAGES.logic);
				}
			}
		}
	});

	function ApiSignatureField(oField)
	{
		ApiBaseField.call(this, oField);
	};

	function private_GetIntAlign(sType)
	{
		if ("left" === sType)
			return AscPDF.ALIGN_TYPE.left;
		else if ("right" === sType)
			return AscPDF.ALIGN_TYPE.right;
		else if ("center" === sType)
			return AscPDF.ALIGN_TYPE.center;

		return undefined;
	}
	function private_GetStrAlign(nType) {
		if (AscPDF.ALIGN_TYPE.left === nType)
			return "left";
		else if (AscPDF.ALIGN_TYPE.right === nType)
			return "right";
		else if (AscPDF.ALIGN_TYPE.center === nType)
			return "center";

		return undefined;
	}

	function private_GetStrHighlight(nType) {
		switch (nType) {
			case AscPDF.BUTTON_HIGHLIGHT_TYPES.push:
				return highlight["p"];
			case AscPDF.BUTTON_HIGHLIGHT_TYPES.invert:
				return highlight["i"];
			case AscPDF.BUTTON_HIGHLIGHT_TYPES.outline:
				return highlight["o"];
			case AscPDF.BUTTON_HIGHLIGHT_TYPES.none:
				return highlight["n"];
		}

		return undefined;
	}

	function private_GetIntHighlight(sType) {
		switch (sType) {
			case highlight["p"]:
				return AscPDF.BUTTON_HIGHLIGHT_TYPES.push;
			case highlight["i"]:
				return AscPDF.BUTTON_HIGHLIGHT_TYPES.invert;
			case highlight["o"]:
				return AscPDF.BUTTON_HIGHLIGHT_TYPES.outline;
			case highlight["n"]:
				return AscPDF.BUTTON_HIGHLIGHT_TYPES.none;
		}

		return undefined;
	}
	function private_GetIntBorderStyle(sType) {
		switch (sType) {
			case "solid":
				return AscPDF.BORDER_TYPES.solid;
			case "dashed":
				return AscPDF.BORDER_TYPES.dashed;
			case "beveled":
				return AscPDF.BORDER_TYPES.beveled;
			case "inset":
				return AscPDF.BORDER_TYPES.inset;
			case "underline":
				return AscPDF.BORDER_TYPES.underline;
			
		}
	}
	function private_GetStrBorderStyle(nType) {
		switch (nType) {
			case AscPDF.BORDER_TYPES.solid:
				return "solid";
			case AscPDF.BORDER_TYPES.dashed:
				return "dashed";
			case AscPDF.BORDER_TYPES.beveled:
				return "beveled";
			case AscPDF.BORDER_TYPES.inset:
				return "inset";
			case AscPDF.BORDER_TYPES.underline:
				return "underline";
			
		}
	}

	function private_GetIntChStyle(sType) {
		switch (sType) {
			case "check":
				return AscPDF.CHECKBOX_STYLES.check;
			case "cross":
				return AscPDF.CHECKBOX_STYLES.cross;
			case "diamond":
				return AscPDF.CHECKBOX_STYLES.diamond;
			case "circle":
				return AscPDF.CHECKBOX_STYLES.circle;
			case "star":
				return AscPDF.CHECKBOX_STYLES.star;
			case "square":
				return AscPDF.CHECKBOX_STYLES.square;
		}
	}
	function private_GetStrChStyle(nType) {
		switch (nType) {
			case AscPDF.CHECKBOX_STYLES.check:
				return "check";
			case AscPDF.CHECKBOX_STYLES.cross:
				return "cross";
			case AscPDF.CHECKBOX_STYLES.diamond:
				return "diamond";
			case AscPDF.CHECKBOX_STYLES.circle:
				return "circle";
			case AscPDF.CHECKBOX_STYLES.star:
				return "star";
			case AscPDF.CHECKBOX_STYLES.square:
				return "square";
		}
	}

	function private_getApiColor(oInternalColor) {
		if (!oInternalColor)
			return ["T"];

		if (oInternalColor.length == 1)
			return ["G", oInternalColor[0]]
		else if (oInternalColor.length == 3)
			return ["RGB", oInternalColor[0], oInternalColor[1], oInternalColor[2]];
		else if (oInternalColor.length == 4)
			return ["CMYK", oInternalColor[0], oInternalColor[1], oInternalColor[2], oInternalColor[3]];

		return ["T"];
	}

	function private_correctApiColor(aApiColor) {
		let sColorSpace = aApiColor[0];
		let aComponents = aApiColor.slice(1);

		function correctComponent(component) {
			if (typeof(component) != "number" || component < 0)
				component = 0;
			else if (component > 1)
				component = 1;

			return component;
		}

		if (sColorSpace == "T")
			return ["T"];
		if (sColorSpace == "RGB") {
			aComponents[0] = correctComponent(aComponents[0]);
			aComponents[1] = correctComponent(aComponents[1]);
			aComponents[2] = correctComponent(aComponents[2]);

			return ["RGB", aComponents[0], aComponents[1], aComponents[2]];
		}
		if (sColorSpace == "G") {
			aComponents[0] = correctComponent[aComponents[0]];

			return ["G", aComponents[0]];
		}
		if (sColorSpace == "CMYK") {
			aComponents[0] = correctComponent(aComponents[0]);
			aComponents[1] = correctComponent(aComponents[1]);
			aComponents[2] = correctComponent(aComponents[2]);
			aComponents[3] = correctComponent(aComponents[3]);

			return ["CMYK", aComponents[0], aComponents[1], aComponents[2], aComponents[3]];
		}

		return ["T"];
	}

	function private_IsValidRect(value, isForStamp) {
		return (
			Array.isArray(value) &&
			value.length === 4 &&
			value.every(Number.isFinite) &&
			(isForStamp !== true ? value[0] < value[2] &&
			value[1] < value[3] : true)
		);
	}

	if (!window["AscPDF"])
		window["AscPDF"] = {};
	
	
	ApiDocument.prototype["getField"]                   = ApiDocument.prototype.getField;


	ApiBaseField.prototype["setAction"]                 = ApiBaseField.prototype.setAction;
	

	ApiPushButtonField.prototype["buttonImportIcon"]    = ApiPushButtonField.prototype.buttonImportIcon;
	

	ApiBaseCheckBoxField.prototype["isBoxChecked"]      = ApiBaseCheckBoxField.prototype.isBoxChecked;
	

	ApiBaseListField.prototype["getItemAt"]             = ApiBaseListField.prototype.getItemAt;
	
	
	window["AscPDF"].ApiDocument          = ApiDocument;
	window["AscPDF"].ApiTextField         = ApiTextField;
	window["AscPDF"].ApiPushButtonField   = ApiPushButtonField;
	window["AscPDF"].ApiCheckBoxField     = ApiCheckBoxField;
	window["AscPDF"].ApiRadioButtonField  = ApiRadioButtonField;
	window["AscPDF"].ApiComboBoxField     = ApiComboBoxField;
	window["AscPDF"].ApiListBoxField      = ApiListBoxField;
})();

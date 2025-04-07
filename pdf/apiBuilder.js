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
(function(window, builder)
{
	/**
	 * A point.
	 * @typedef {number} pt
	 */

	/**
	 * Any valid field element.
	 * @typedef {(ApiTextField | ApiComboboxField | ApiListboxField | ApiButtonField | ApiCheckboxField | ApiRadiobuttonField )} ApiField
	 */

	/**
	 * Any valid field element.
	 * @typedef {(ApiBaseWidget | ApiTextWidget | ApiCheckboxWidget | ApiButtonWidget )} ApiWidget
	 */

	/**
	 * @typedef {number[]} WidgetRect
	 * @property {number} 0 - x1
	 * @property {number} 1 - y1
	 * @property {number} 2 - x2
	 * @property {number} 3 - y2
	 */

	
	/**
	 * @typedef {Object} ListOptionTuple
	 * @property {string} 0 - displayed value
	 * @property {string} 1 - exported value
	 */

	/**
	 * @typedef {(string | ListOptionTuple)} ListOption
	 */

	/**
	 * The available check styles.
	 * @typedef {("check" | "cross" | "diamond" | "circle" | "star" | "square")} CheckStyle
	 */

	/**
	 * The available widget border width.
	 * @typedef {("none" | "thin" | "medium" | "thick")} WidgetBorderWidth
	 */

	/**
	 * The available widget border styles.
	 * @typedef {("solid" | "beveled" | "dashed" | "inset" | "underline")} WidgetBorderStyle
	 */

	/**
	 * The available widget border styles.
	 * @typedef {("solid" | "beveled" | "dashed" | "inset" | "underline")} WidgetBorderStyle
	 */

	/**
	 * The available button widget border appearances types.
	 * @typedef {("normal" | "down" | "hover")} ButtonAppearance
	 */

	/**
	 * The available button widget layout types.
	 * @typedef {("textOnly" | "iconOnly" | "iconTextV" | "textIconV" | "iconTextH" | "textIconH" | "overlay")} ButtonLayout
	 */

	/**
	 * The available button widget scale when types.
	 * @typedef {("always" | "never" | "tooBig" | "tooSmall")} ButtonScaleWhen
	 */

	/**
	 * The available button widget scale how types.
	 * @typedef {("proportional" | "anamorphic")} ButtonScaleHow
	 */

	/**
	 * The available button widget behavior types.
	 * @typedef {("none" | "invert" | "push" | "outline")} ButtonBehavior
	 */

	/**
	 * Value from 0 to 100.
	 * @typedef {number} percentage
	 */

	/**
	 * NumberSepStyle — defines number formatting style:
	 * - "us"        — 1,234.56   (English style)
	 * - "plain"     — 1234.56    (No separators)
	 * - "euro"      — 1.234,56   (European style)
	 * - "europlain" — 1234,56    (European without separators)
	 * - "ch"        — 1'234.56   (Swiss style)
	 * @typedef {("us" | "plain" | "euro" | "europlain" | "ch")} NumberSepStyle
	 */

	/**
	 * NumberNegStyle defines the formatting style for negative numbers:
	 *
	 * - "black-minus" — "-1,234.56" (black minus sign)
	 * - "red-minus"   — "-1,234.56" (red minus sign)
	 * - "black-parens" — "(1,234.56)"" (black parentheses)
	 * - "red-parens"   — "(1,234.56)"" (red parentheses)
	 *
	 * @typedef {"black-minus" | "red-minus" | "black-parens" | "red-parens"} NumberNegStyle
	 */

	/**
	 * PsfFormat defines the type of formatting to apply:
	 *
	 * - "zip"       — ZIP code (e.g., 12345)
	 * - "zip+4"     — ZIP+4 (e.g., 12345-6789)
	 * - "phone"     — Phone number (e.g., (123) 456-7890)
	 * - "ssn"       — Social Security Number (e.g., 123-45-6789)
	 *
	 * @typedef {"zip" | "zip+4" | "phone" | "ssn"} PsfFormat
	 */

	/**
	 * @typedef {'24HR:MM' | '12HR:MM' | '24HR:MM:SS' | '12HR:MM:SS'} TimeFormat
	 * Time format options:
	 * - "24HR_MM" — 24-hour format, hours and minutes (e.g., "14:30")
	 * - "12HR_MM" — 12-hour format with AM/PM, hours and minutes (e.g., "2:30 PM")
	 * - "24HR_MM_SS" — 24-hour format, hours, minutes, and seconds (e.g., "14:30:15")
	 * - "12HR_MM_SS" — 12-hour format with AM/PM, hours, minutes, and seconds (e.g., "2:30:15 PM")
	 */

	//------------------------------------------------------------------------------------------------------------------
	//
	// Api
	//
	//------------------------------------------------------------------------------------------------------------------

	let position    = AscPDF.Api.Types.position;
    let scaleWhen   = AscPDF.Api.Types.scaleWhen;
    let scaleHow    = AscPDF.Api.Types.scaleHow;

	/**
	 * Base class
	 * @global
	 * @class
	 * @name Api
	 */
	let Api = window["Asc"]["PDFEditorApi"];

	/**
	 * Creates a text field with the specified text field properties.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiDocument}
	 * @see office-js-api/Examples/PDF/Api/Methods/GetDocument.js
	 */
	Api.prototype.GetDocument = function() {
		return new ApiDocument(private_GetLogicDocument());
	};

	/**
	 * Creates an RGB color setting the appropriate values for the red, green and blue color components.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @returns {ApiRGBColor}
	 * @see office-js-api/Examples/PDF/Api/Methods/CreateRGBColor.js
	 */
	Api.prototype.CreateRGBColor = function(r, g, b) {
		return new ApiRGBColor(r, g, b);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiDocument
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a document.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 */
	function ApiDocument(oDoc) {
		this.Document = oDoc;
	}

	/**
	 * Returns a type of the ApiDocument class.
	 * @memberof ApiDocument
	 * @typeofeditors ["PDFE"]
	 * @returns {"page"}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/GetClassType.js
	 */
	ApiDocument.prototype.GetClassType = function() {
		return "document";
	};

	/**
	 * Adds a new page to document.
	 * @memberof ApiDocument
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPos - Text field properties.
	 * @param {pt} [nWidth] - Text field properties.
	 * @param {pt} [nHeight] - Text field properties.
	 * @returns {ApiPage}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/AddPage.js
	 */
	ApiDocument.prototype.AddPage = function(nPos, nWidth, nHeight) {
		let oDoc = private_GetLogicDocument();
		let oFile = oDoc.GetFile();

		let oPageToClone = oFile.pages[nPos - 1] || oFile.pages[nPos];

		let oPage = {
			fonts: [],
			Rotate: 0,
			Dpi: 72,
			W: nWidth || oPageToClone.W,
			H: nHeight || oPageToClone.H
		}

		this.Document.AddPage(nPos, oPage);

		return new ApiPage(this.Document.GetPageInfo(nPos));
	};

	/**
	 * Gets page by index from document.
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPos - page position
	 * @returns {ApiPage}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/GetPage.js
	 */
	ApiDocument.prototype.GetPage = function(nPos) {
		let oPageInfo = this.Document.GetPageInfo(nPos);
		if (!oPageInfo) {
			return null;
		}

		return new ApiPage(oPageInfo);
	};

	/**
	 * Removes page by index from document
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPos - page position
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/RemovePage.js
	 */
	ApiDocument.prototype.RemovePage = function(nPos) {
		let oFile = this.Document.Viewer.file;
		if (!oFile.pages[nPos]) {
			return false;
		}

		this.Document.RemovePage(nPos);
		return true;
	};

	/**
	 * Gets document pages count
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/GetPagesCount.js
	 */
	ApiDocument.prototype.GetPagesCount = function() {
		let oFile = this.Document.Viewer.file;
		return oFile.pages.length;
	};

	/**
	 * Creates a text field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPage - page index
	 * @param {WidgetRect} aRect - widget rect
	 * @returns {ApiTextField}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/AddTextField.js
	 */
	ApiDocument.prototype.AddTextField = function(nPage, aRect) {
		let oField = this.Document.CreateTextField();
		oField.SetRect(aRect);

		this.Document.AddField(oField, nPage);
		return new ApiTextField(oField);
	};

	/**
	 * Creates a text date field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPage - page index
	 * @param {WidgetRect} aRect - widget rect
	 * @returns {ApiTextField}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/AddDateField.js
	 */
	ApiDocument.prototype.AddDateField = function(nPage, aRect) {
		let oField = this.Document.CreateTextField(true);
		oField.SetRect(aRect);

		this.Document.AddField(oField, nPage);
		return new ApiTextField(oField);
	};

	/**
	 * Creates a image field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPage - page index
	 * @param {WidgetRect} aRect - widget rect
	 * @returns {ApiTextField}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/AddImageField.js
	 */
	ApiDocument.prototype.AddImageField = function(nPage, aRect) {
		let oField = this.Document.CreateButtonField(true);
		oField.SetRect(aRect);

		this.Document.AddField(oField, nPage);
		return new ApiButtonField(oField);
	};

	/**
	 * Creates a checkbox field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPage - page index
	 * @param {WidgetRect} aRect - widget rect
	 * @returns {ApiTextField}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/AddImageField.js
	 */
	ApiDocument.prototype.AddCheckboxField = function(nPage, aRect) {
		let oField = this.Document.CreateCheckboxField();
		oField.SetRect(aRect);

		this.Document.AddField(oField, nPage);
		return new ApiCheckboxField(oField);
	};

	/**
	 * Creates a radiobutton field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPage - page index
	 * @param {WidgetRect} aRect - widget rect
	 * @returns {ApiTextField}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/AddRadiobuttonField.js
	 */
	ApiDocument.prototype.AddRadiobuttonField = function(nPage, aRect) {
		let oField = this.Document.CreateRadiobuttonField();
		oField.SetRect(aRect);

		this.Document.AddField(oField, nPage);
		return new ApiRadiobuttonField(oField);
	};

	/**
	 * Creates a combobox field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPage - page index
	 * @param {WidgetRect} aRect - widget rect
	 * @returns {ApiTextField}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/AddComboboxField.js
	 */
	ApiDocument.prototype.AddComboboxField = function(nPage, aRect) {
		let oField = this.Document.CreateComboboxField();
		oField.SetRect(aRect);

		this.Document.AddField(oField, nPage);
		return new ApiComboboxField(oField);
	};

	/**
	 * Creates a listbox field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPage - page index
	 * @param {WidgetRect} aRect - widget rect
	 * @returns {ApiTextField}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/AddListboxField.js
	 */
	ApiDocument.prototype.AddListboxField = function(nPage, aRect) {
		let oField = this.Document.CreateListboxField();
		oField.SetRect(aRect);

		this.Document.AddField(oField, nPage);
		return new ApiListboxField(oField);
	};

	/**
	 * Gets list of all fields in document.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiField}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/GetAllFields.js
	 */
	ApiDocument.prototype.GetAllFields = function() {
		let aFields = [];
		
		for (let i = 0, nCount = this.Document.GetPagesCount(); i < nCount; i++) {
			let oPageInfo = this.Document.GetPageInfo(i);

			oPageInfo.fields.forEach(function(widget) {
				let oParent = widget.GetParent();
				
				if (oParent) {
					while (oParent) {
						if (!aFields.includes(oParent)) {
							aFields.push(oParent);
						}

						oParent = oParent.GetParent();
					}
				}
				else if (!aFields.includes(widget)) {
					aFields.push(widget);
				}
			});
		}

		return aFields.map(private_GetFieldApi);
	};

	/**
	 * Gets field by it's name.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiField}
	 * @see office-js-api/Examples/PDF/ApiDocument/Methods/GetFieldByName.js
	 */
	ApiDocument.prototype.GetFieldByName = function(sName) {
		let oField = this.Document.GetField(sName);
		if (false == oField.IsWidget() || !oField.GetParent())	{
			return private_GetFieldApi(oField);
		}
		else {
			return private_GetFieldApi(oField.GetParent());
		}
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiPage
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a document page.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 */
	function ApiPage(oPage) {
		this.Page = oPage;
	}

	/**
	 * Returns a type of the ApiPage class.
	 * @memberof ApiPage
	 * @typeofeditors ["PDFE"]
	 * @returns {"page"}
	 * @see office-js-api/Examples/PDF/ApiPage/Methods/GetClassType.js
	 */
	ApiPage.prototype.GetClassType = function() {
		return "page";
	};

	/**
	 * Sets page rotation angle
	 * @typeofeditors ["PDFE"]
	 * @param {number} nAngle
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiPage/Methods/SetRotate.js
	 */
	ApiPage.prototype.SetRotate = function(nAngle) {
		if (nAngle % 90 !== 0) {
			return false;
		}

		let oDoc = private_GetLogicDocument();
		oDoc.SetPageRotate(this.GetIndex(), nAngle);
		return true;
	};

	/**
	 * Gets page rotation angle
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/PDF/ApiPage/Methods/GetRotate.js
	 */
	ApiPage.prototype.GetRotate = function() {
		return this.Page.GetRotate();
	};

	/**
	 * Gets page index
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/PDF/ApiPage/Methods/GetIndex.js
	 */
	ApiPage.prototype.GetIndex = function() {
		return this.Page.GetIndex();
	};

	/**
	 * Gets page widgets
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/PDF/ApiPage/Methods/GetAllWidgets.js
	 */
	ApiPage.prototype.GetAllWidgets = function() {
		return this.Page.fields.map(private_GetWidgetApi);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiBaseField
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a base field.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 */
	function ApiBaseField(oField) {
		this.Field = oField;
	}

	/**
	 * Sets new field name if possible.
	 * @typeofeditors ["PDFE"]
	 * @param {string} sName
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseField/Methods/SetFullName.js
	 */
	ApiBaseField.prototype.SetFullName = function(sName) {
		return this.Field.SetName(sName);
	};

	/**
	 * Gets field full name.
	 * @typeofeditors ["PDFE"]
	 * @returns {string}
	 * @see office-js-api/Examples/PDF/ApiBaseField/Methods/GetFullName.js
	 */
	ApiBaseField.prototype.GetFullName = function() {
		return this.Field.GetFullName();
	};

	/**
	 * Sets new field partial name.
	 * @typeofeditors ["PDFE"]
	 * @param {string} sName
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseField/Methods/SetPartialName.js
	 */
	ApiBaseField.prototype.SetPartialName = function(sName) {
		return this.Field.SetPartialName(sName);
	};

	/**
	 * Gets field partial name.
	 * @typeofeditors ["PDFE"]
	 * @returns {string}
	 * @see office-js-api/Examples/PDF/ApiBaseField/Methods/GetPartialName.js
	 */
	ApiBaseField.prototype.GetPartialName = function() {
		return this.Field.GetPartialName();
	};
	
	/**
	 * Sets field required
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseField/Methods/SetRequired.js
	 */
	ApiBaseField.prototype.SetRequired = function(bRequired) {
		this.Field.SetRequired(bRequired);
		return true;
	};

	/**
	 * Checks if field is required
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseField/Methods/IsRequired.js
	 */
	ApiBaseField.prototype.IsRequired = function() {
		return this.Field.IsRequired();
	};

	/**
	 * Sets field read only
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseField/Methods/SetReadOnly.js
	 */
	ApiBaseField.prototype.SetReadOnly = function(bReadOnly) {
		this.Field.SetReadOnly(bReadOnly);
		return true;
	};

	/**
	 * Checks if field is read only
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseField/Methods/IsReadOnly.js
	 */
	ApiBaseField.prototype.IsReadOnly = function() {
		return this.Field.IsReadOnly();
	};

	/**
	 * Sets field value
	 * @typeofeditors ["PDFE"]
	 * @param {string} sValue
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseField/Methods/SetValue.js
	 */
	ApiBaseField.prototype.SetValue = function(sValue) {
		let oDoc = private_GetLogicDocument();

		let oFieldToCommit = this.Field.IsWidget() ? this.Field : this.Field.GetKid(0);

		if (sValue != undefined && sValue.toString) {
			sValue = sValue.toString();
		}

		oFieldToCommit.SetValue(sValue);
		return oDoc.CommitField(oFieldToCommit);
	};

	/**
	 * Gets field value
	 * @typeofeditors ["PDFE"]
	 * @returns {string}
	 * @see office-js-api/Examples/PDF/ApiBaseField/Methods/GetValue.js
	 */
	ApiBaseField.prototype.GetValue = function() {
		return this.Field.GetParentValue();
	};

	/**
	 * Adds new widget - visual representation for field
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPage - page to add widget
	 * @param {WidgetRect} aRect - field rect
	 * @returns {?ApiWidget}
	 * @see office-js-api/Examples/PDF/ApiBaseField/Methods/AddWidget.js
	 */
	ApiBaseField.prototype.AddWidget = function(nPage, aRect) {
		let oDoc		= private_GetLogicDocument();
		let oPage		= oDoc.GetPageInfo(nPage);
		let nFieldType	= this.Field.GetType();

		if (!oPage) {
			return null;
		}

		let oWidget = oDoc.CreateField(this.Field.GetFullName(), nFieldType, aRect);
		oDoc.AddField(oWidget, nPage);

		this.Field = oWidget.GetParent();

		return private_GetWidgetApi(oWidget);
	};

	/**
	 * Gets array with widgets of the current field.
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiWidget}
	 * @see office-js-api/Examples/PDF/ApiBaseField/Methods/GetAllWidgets.js
	 */
	ApiBaseField.prototype.GetAllWidgets = function() {
		return this.Field.GetAllWidgets().map(private_GetWidgetApi);
	};

	/**
	 * Class representing a base field widget.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 */
	function ApiBaseWidget(oField) {
		this.Field = oField;
	}

	/**
	 * Returns a type of the ApiBaseWidget class.
	 * @memberof ApiBaseWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {"page"}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/GetClassType.js
	 */
	ApiBaseWidget.prototype.GetClassType = function() {
		return "baseWidget";
	};

	/**
	 * Sets widget border color.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiRGBColor} oColor
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/SetBorderColor.js
	 */
	ApiBaseWidget.prototype.SetBorderColor = function(oColor) {
		if (!(oColor instanceof ApiRGBColor)) {
			return false;
		}

		this.Field.SetBorderColor(private_GetInnerColorByRGB(oColor.R, oColor.G, oColor.B));

		if (this.Field.GetBorderStyle() == undefined) {
			this.Field.SetBorderStyle(AscPDF.BORDER_TYPES.solid);
		}
		if (this.Field.GetBorderWidth() == undefined) {
			this.Field.SetBorderWidth(AscPDF.BORDER_WIDTH.thin);
		}
		
		return true;
	};

	/**
	 * Gets widget border color.
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiRGBColor}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/GetBorderColor.js
	 */
	ApiBaseWidget.prototype.GetBorderColor = function() {
		let aInnerColor = this.Field.GetBorderColor();
		if (!aInnerColor) {
			return null;
		}

		let oRGB = this.Field.GetRGBColor(aInnerColor);

		return new ApiRGBColor(oRGB.r, oRGB.g, oRGB.b);
	};

	/**
	 * Sets widget border width.
	 * @typeofeditors ["PDFE"]
	 * @param {WidgetBorderWidth} sBorderWidth
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/SetBorderWidth.js
	 */
	ApiBaseWidget.prototype.SetBorderWidth = function(sBorderWidth) {
		if (!Object.keys(AscPDF.BORDER_WIDTH).includes(sBorderWidth)) {
			return false;
		}

		this.Field.SetBorderWidth(private_GetInnerBorderWidth(sBorderWidth));
		return true;
	};

	/**
	 * Gets widget border width.
	 * @typeofeditors ["PDFE"]
	 * @returns {WidgetBorderWidth}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/GetBorderWidth.js
	 */
	ApiBaseWidget.prototype.GetBorderWidth = function() {
		return private_GetStrBorderWidth(this.Field.GetBorderWidth());
	};

	/**
	 * Sets widget border style.
	 * @typeofeditors ["PDFE"]
	 * @param {WidgetBorderStyle} sBorderStyle
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/SetBorderStyle.js
	 */
	ApiBaseWidget.prototype.SetBorderStyle = function(sBorderStyle) {
		if (!Object.keys(AscPDF.BORDER_TYPES).includes(sBorderStyle)) {
			return false;
		}

		this.Field.SetBorderStyle(private_GetInnerBorderStyle(sBorderStyle));
		return true;
	};

	/**
	 * Gets widget border style.
	 * @typeofeditors ["PDFE"]
	 * @returns {WidgetBorderStyle}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/GetBorderStyle.js
	 */
	ApiBaseWidget.prototype.GetBorderStyle = function() {
		return private_GetStrBorderStyle(this.Field.GetBorderStyle());
	};

	/**
	 * Sets widget background color.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiRGBColor} oColor
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/SetBackgroundColor.js
	 */
	ApiBaseWidget.prototype.SetBackgroundColor = function(oColor) {
		if (!(oColor instanceof ApiRGBColor)) {
			return false;
		}

		this.Field.SetBackgroundColor(private_GetInnerColorByRGB(oColor.R, oColor.G, oColor.B));
		return true;
	};

	/**
	 * Gets widget background color.
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiRGBColor}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/GetBackgroundColor.js
	 */
	ApiBaseWidget.prototype.GetBackgroundColor = function() {
		let aInnerColor = this.Field.GetBackgroundColor();
		if (!aInnerColor) {
			return null;
		}

		let oRGB = this.Field.GetRGBColor(aInnerColor);

		return new ApiRGBColor(oRGB.r, oRGB.g, oRGB.b);
	};

	/**
	 * Sets widget text color.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiRGBColor} oColor
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/SetTextColor.js
	 */
	ApiBaseWidget.prototype.SetTextColor = function(oColor) {
		if (!(oColor instanceof ApiRGBColor)) {
			return false;
		}

		this.Field.SetTextColor(private_GetInnerColorByRGB(oColor.R, oColor.G, oColor.B));
		return true;
	};

	/**
	 * Gets widget text color.
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiRGBColor}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/GetTextColor.js
	 */
	ApiBaseWidget.prototype.GetTextColor = function() {
		let aInnerColor = this.Field.GetTextColor();
		if (!aInnerColor) {
			return null;
		}

		let oRGB = this.Field.GetRGBColor(aInnerColor);

		return new ApiRGBColor(oRGB.r, oRGB.g, oRGB.b);
	};

	/**
	 * Sets widget text size.
	 * <note> Text size === 0 means autofit </note>
	 * @typeofeditors ["PDFE"]
	 * @param {pt} nSize
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/SetTextSize.js
	 */
	ApiBaseWidget.prototype.SetTextSize = function(nSize) {
		if (typeof(nSize) != 'number' || nSize < 0) {
			return false;
		}

		this.Field.SetTextSize(nSize);
		return true;
	};

	/**
	 * Gets widget text size.
	 * <note> Text size === 0 means autofit </note>
	 * @typeofeditors ["PDFE"]
	 * @returns {pt}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/GetTextSize.js
	 */
	ApiBaseWidget.prototype.GetTextSize = function() {
		return this.Field.GetTextSize();
	};

	/**
	 * Sets text autofit.
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bAuto
	 * @returns {?ApiRGBColor}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/SetAutoFit.js
	 */
	ApiBaseWidget.prototype.SetAutoFit = function(bAuto) {
		return this.Field.SetTextSize(bAuto ? 0 : 11);
	};

	/**
	 * Checks if text is autofit.
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiRGBColor}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/IsAutoFit.js
	 */
	ApiBaseWidget.prototype.IsAutoFit = function() {
		return this.Field.GetTextSize() == 0;
	};

	/**
	 * Removes widget from parent field.
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseWidget/Methods/Remove.js
	 */
	ApiBaseWidget.prototype.Remove = function() {
		let oDoc = private_GetLogicDocument();
		return oDoc.RemoveField(this.Field.GetId());
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTextField
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a text field.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseField}
	 */
	function ApiTextField(oField) {
		ApiBaseField.call(this, oField);
	}

	ApiTextField.prototype = Object.create(ApiBaseField.prototype);
	ApiTextField.prototype.constructor = ApiTextField;

	/**
	 * Returns a type of the ApiTextField class.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @returns {"textField"}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/GetClassType.js
	 */
	ApiTextField.prototype.GetClassType = function() {
		return "textField";
	};

	/**
	 * Sets text field multiline prop.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bMultiline - will the field be multiline
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/SetMultiline.js
	 */
	ApiTextField.prototype.SetMultiline = function(bMultiline) {
		return this.Field.SetMultiline(bMultiline)
	};

	/**
	 * Checks if text field is multiline.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/IsMultiline.js
	 */
	ApiTextField.prototype.IsMultiline = function() {
		return this.Field.IsMultiline()
	};

	/**
	 * Sets text field chars limit.
	 * <note> Char limit 0 means field doesn't have char limit
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @param {number} nChars - chars limit number
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/SetCharLimit.js
	 */
	ApiTextField.prototype.SetCharLimit = function(nChars) {
		return this.Field.SetCharLimit(nChars)
	};

	/**
	 * Gets text field chars limit.
	 * <note> Char limit 0 means field doesn't have char limit
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/GetCharLimit.js
	 */
	ApiTextField.prototype.GetCharLimit = function() {
		return this.Field.GetCharLimit()
	};

	/**
	 * Sets text field comb prop.
	 * <note> Should have char limit more then 0 </note>
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bComb - will the field be comb
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/SetComb.js
	 */
	ApiTextField.prototype.SetComb = function(bComb) {
		return this.Field.SetComb(bComb)
	};

	/**
	 * Checks if text field is comb.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/IsComb.js
	 */
	ApiTextField.prototype.IsComb = function() {
		return this.Field.IsComb()
	};

	/**
	 * Sets text field can scroll long text prop.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bScroll - can the field scroll long text 
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/SetScrollLongText.js
	 */
	ApiTextField.prototype.SetScrollLongText = function(bScroll) {
		return this.Field.SetDoNotScroll(!bScroll)
	};

	/**
	 * Checks if text field can scroll long text.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/IsScrollLongText.js
	 */
	ApiTextField.prototype.IsScrollLongText = function() {
		return !this.Field.IsDoNotScroll()
	};

	/**
	 * Sets number format for field.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @param {number} nDemical - number of decimals
	 * @param {NumberSepStyle} - number separate style
	 * @param {NumberNegStyle} - number negative style
	 * @param {string} sCurrency - currency sybmol
	 * @param {boolean} bCurrencyPrepend - If true, places the currency symbol before the number (e.g., $1,234.56); 
 	 * if false, places it after (e.g., 1,234.56$).
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/SetNumberFormat.js
	 */
	ApiTextField.prototype.SetNumberFormat = function(nDemical, sSepStyle, sNegStyle, sCurrency, bCurrencyPrepend) {
		this.Field.ClearFormat();

		let aActionsFormat = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": "AFNumber_Format(" + nDemical + "," + private_GetInnerNumberSeparateType(sSepStyle) + "," + private_GetInnerNumberNegType(sNegStyle) + "," + "0" + ',"' + sCurrency + '",' + bCurrencyPrepend + ");"
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Format, aActionsFormat);

		let aActionsKeystroke = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": "AFNumber_Keystroke(" + nDemical + "," + private_GetInnerNumberSeparateType(sSepStyle) + "," + private_GetInnerNumberNegType(sNegStyle) + "," + "0" + ',"' + sCurrency + '",' + bCurrencyPrepend + ");"
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke, aActionsKeystroke);

		this.Field.Commit();

		return true;
	};

	/**
	 * Sets percentage format for field.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @param {number} nDemical - number of decimals
	 * @param {NumberSepStyle} - number separate style
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/SetPercentageFormat.js
	 */
	ApiTextField.prototype.SetPercentageFormat = function(nDemical, sSepStyle) {
		this.Field.ClearFormat();

		let aActionsFormat = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": "AFPercent_Format(" + nDemical + "," + private_GetInnerNumberSeparateType(sSepStyle) + ");"
		}]
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Format, aActionsFormat);

		let aActionsKeystroke = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": "AFPercent_Keystroke(" + nDemical + "," + private_GetInnerNumberSeparateType(sSepStyle) + ");"
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke, aActionsKeystroke);
		this.Field.Commit();

		return true;
	};

	/**
	 * Sets date format for field.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @param {string} sFormat - date format (e.g. "dd.mm.yyyy")
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/SetDateFormat.js
	 */
	ApiTextField.prototype.SetDateFormat = function(sFormat) {
		this.Field.ClearFormat();

		let aActionsFormat = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": 'AFDate_Format("' + sFormat + '");'
		}]
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Format, aActionsFormat);

		let aActionsKeystroke = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": 'AFDate_Keystroke("' + sFormat + '");'
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke, aActionsKeystroke);
		this.Field.Commit();

		return true;
	};

	/**
	 * Sets time format for field.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @param {TimeFormat} sFormat - available time format
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/SetTimeFormat.js
	 */
	ApiTextField.prototype.SetTimeFormat = function(sFormat) {
		this.Field.ClearFormat();

		let aActionsFormat = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": 'AFTime_Format(' + private_GetInnerTimeFormatType(sFormat) + ');'
		}]
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Format, aActionsFormat);

		let aActionsKeystroke = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": 'AFTime_Keystroke(' + private_GetInnerTimeFormatType(sFormat) + ');'
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke, aActionsKeystroke);
		this.Field.Commit();

		return true;
	};

	/**
	 * Sets special format for field.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @param {PsfFormat} sFormat - the formatting style to apply to the value
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/SetSpecialFormat.js
	 */
	ApiTextField.prototype.SetSpecialFormat = function(sFormat) {
		this.Field.ClearFormat();
				
		let aActionsFormat = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": "AFSpecial_Format(" + private_GetInnerSpecialPsfType(sFormat) + ");"
		}]
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Format, aActionsFormat);

		let aActionsKeystroke = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": "AFSpecial_Keystroke(" + private_GetInnerSpecialPsfType(sFormat) + ");"
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke, aActionsKeystroke);
		this.Field.Commit();

		return true;
	};

	/**
	 * Sets mask for entered text for field.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @param {string} sMask - field mask (e.g. "(999)999-9999")
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/SetMask.js
	 */
	ApiTextField.prototype.SetMask = function(sMask) {
		this.Field.ClearFormat();
		this.Field.SetArbitaryMask(sMask);
		this.Field.Commit();

		return true;
	};

	/**
	 * Sets regular expression validate string for field.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @param {string} sReg - field regular expression (e.g. "\\S+@\\S+\\.\\S+")
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/SetRegularExp.js
	 */
	ApiTextField.prototype.SetRegularExp = function(sReg) {
		this.Field.ClearFormat();
		this.Field.SetRegularExp(sReg);
		this.Field.Commit();

		return true;
	};

	/**
	 * Clears format of field.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/ClearFormat.js
	 */
	ApiTextField.prototype.ClearFormat = function() {
		this.Field.ClearFormat();
		this.Field.Commit();

		return true;
	};

	/**
	 * Sets validate range for field.
	 * <note> Can only be applied to fields with a percentage or number format. </note>
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} [bGreaterThan=false] - If true, enables minimum value check using `nGreaterThan`.
	 * @param {number} nGreaterThan - Minimum allowed value (inclusive or exclusive based on implementation).
	 * @param {boolean} [bLessThan=false] - If true, enables maximum value check using `nLessThan`.
	 * @param {number} nLessThan - Maximum allowed value (inclusive or exclusive based on implementation).
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextField/Methods/SetValidateRange.js
	 */
	ApiTextField.prototype.SetValidateRange = function(bGreaterThan, nGreaterThan, bLessThan, nLessThan) {
		if (false == this.Field.IsNumberFormat()) {
			return false;
		}
		
		if (bGreaterThan == undefined) {
			bGreaterThan = false;
		}
		if (bLessThan == undefined) {
			bLessThan = false;
		}

		let aActionsValidate = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": 'AFRange_Validate(' + bGreaterThan +  ',' + nGreaterThan + ',' + bLessThan + ',' + nLessThan +  ');'
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Validate, aActionsValidate);

		return true;
	};

	/**
	 * Class representing a text field widget.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 */
	function ApiTextWidget(oField) {
		ApiBaseWidget.call(this, oField);
	}

	ApiTextWidget.prototype = Object.create(ApiBaseWidget.prototype);
	ApiTextWidget.prototype.constructor = ApiTextWidget;

	/**
	 * Returns a type of the ApiTextWidget class.
	 * @memberof ApiTextWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {"page"}
	 * @see office-js-api/Examples/PDF/ApiTextWidget/Methods/GetClassType.js
	 */
	ApiTextWidget.prototype.GetClassType = function() {
		return "textWidget";
	};

	/**
	 * Sets text field placeholder.
	 * @memberof ApiTextWidget
	 * @typeofeditors ["PDFE"]
	 * @param {string} sPlaceholder - field placeholder 
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextWidget/Methods/SetPlaceholder.js
	 */
	ApiTextWidget.prototype.SetPlaceholder = function(sText) {
		return this.Field.SetPlaceholder(sText)
	};

	/**
	 * Gets text field placeholder.
	 * @memberof ApiTextWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {string}
	 * @see office-js-api/Examples/PDF/ApiTextWidget/Methods/GetPlaceholder.js
	 */
	ApiTextWidget.prototype.GetPlaceholder = function() {
		return this.Field.GetPlaceholder()
	};

	/**
	 * Sets text widget regular validate expression.
	 * @memberof ApiTextWidget
	 * @typeofeditors ["PDFE"]
	 * @param {string} sReg - field regular exp 
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextWidget/Methods/SetRegularExp.js
	 */
	ApiTextWidget.prototype.SetRegularExp = function(sReg) {
		return this.Field.SetRegularExp(sReg)
	};

	/**
	 * Gets text widget regular validate expression.
	 * @memberof ApiTextWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiTextWidget/Methods/GetRegularExp.js
	 */
	ApiTextWidget.prototype.GetRegularExp = function() {
		return this.Field.GetRegularExp()
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiBaseListField
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a base list field.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseField}
	 */
	function ApiBaseListField(oField) {
		ApiBaseField.call(this, oField);
	}

	ApiBaseListField.prototype = Object.create(ApiBaseField.prototype);
	ApiBaseListField.prototype.constructor = ApiBaseListField;

	/**
	 * Adds new option to list options.
	 * @memberof ApiBaseListField
	 * @typeofeditors ["PDFE"]
	 * @param {ListOption} option - list option to add
	 * @param {number} [nPos=this.GetOptions().lenght] - pos to add option
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseListField/Methods/AddOption.js
	 */
	ApiBaseListField.prototype.AddOption = function(option, nPos) {
		return this.Field.AddOption(option, nPos);
	};

	/**
	 * Removes option from list options.
	 * @memberof ApiBaseListField
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPos - pos to remove option
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseListField/Methods/RemoveOption.js
	 */
	ApiBaseListField.prototype.RemoveOption = function(nPos) {
		return !!this.Field.RemoveOption(nPos);
	};

	/**
	 * Moves option to specified position in list options.
	 * @memberof ApiBaseListField
	 * @typeofeditors ["PDFE"]
	 * @param {number} nCurPos - index of moved option
	 * @param {number} nNewPos - new positon for option
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseListField/Methods/MoveOptionTo.js
	 */
	ApiBaseListField.prototype.MoveOptionTo = function(nCurPos, nNewPos) {
		let aOptions = this.GetOptions();
		if (nCurPos < 0 || nCurPos >= aOptions.length || nNewPos < 0) return false;
	
		let opt = this.Field.RemoveOption(nCurPos);
		if (!opt)
			return false;
	
		let nTargetPos = Math.min(nNewPos, aOptions.length);
		
		this.Field.AddOption(opt, nTargetPos);
		return true;
	};

	/**
	 * Gets option from list options.
	 * @memberof ApiBaseListField
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPos - option index to get
	 * @returns {ListOption}
	 * @see office-js-api/Examples/PDF/ApiBaseListField/Methods/GetOption.js
	 */
	ApiBaseListField.prototype.GetOption = function(nPos) {
		let aOptions = this.Field.GetOptions();
		if (aOptions) {
			return aOptions[nPos];
		}

		return null;
	};

	/**
	 * Gets all options from list options.
	 * @memberof ApiBaseListField
	 * @typeofeditors ["PDFE"]
	 * @returns {ListOption[]}
	 * @see office-js-api/Examples/PDF/ApiBaseListField/Methods/GetOptions.js
	 */
	ApiBaseListField.prototype.GetOptions = function() {
		let aOptions = this.Field.GetOptions();
		return aOptions;
	};

	/**
	 * Sets field commit on selection change prop.
	 * @memberof ApiBaseListField
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bCommit - will the field value be applied to all with the same name immediately after the change
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseListField/Methods/SetCommitOnSelChange.js
	 */
	ApiBaseListField.prototype.SetCommitOnSelChange = function(bCommit) {
		return this.Field.SetCommitOnSelChange(bCommit)
	};

	/**
	 * Checks if field can commit on selection change.
	 * @memberof ApiBaseListField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseListField/Methods/IsCommitOnSelChange.js
	 */
	ApiBaseListField.prototype.IsCommitOnSelChange = function() {
		return this.Field.IsCommitOnSelChange()
	};

	/**
	 * Sets selected value indexes.
	 * @memberof ApiBaseListField
	 * @typeofeditors ["PDFE"]
	 * @param {number[]} aIndexes - selected indexes
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiBaseListField/Methods/SetValueIndexes.js
	 */
	ApiBaseListField.prototype.SetValueIndexes = function(aIndexes) {
		let oDoc = private_GetLogicDocument();

		let oFieldToCommit = this.Field.IsWidget() ? this.Field : this.Field.GetKid(0);

		oFieldToCommit.SetCurIdxs(aIndexes);
		return oDoc.CommitField(oFieldToCommit);
	};

	/**
	 * Gets selected value indexes.
	 * @memberof ApiBaseListField
	 * @typeofeditors ["PDFE"]
	 * @returns {number[]}
	 * @see office-js-api/Examples/PDF/ApiBaseListField/Methods/GetValueIndexes.js
	 */
	ApiBaseListField.prototype.GetValueIndexes = function() {
		return this.Field.GetParentCurIdxs();
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiComboboxField
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a combobox field.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseListField}
	 */
	function ApiComboboxField(oField) {
		ApiBaseListField.call(this, oField);
	}

	ApiComboboxField.prototype = Object.create(ApiBaseListField.prototype);
	ApiComboboxField.prototype.constructor = ApiComboboxField;

	/**
	 * Returns a type of the ApiComboboxField class.
	 * @memberof ApiComboboxField
	 * @typeofeditors ["PDFE"]
	 * @returns {"comboboxField"}
	 * @see office-js-api/Examples/PDF/ApiComboboxField/Methods/GetClassType.js
	 */
	ApiComboboxField.prototype.GetClassType = function() {
		return "comboboxField";
	};

	/**
	 * Sets field editable prop.
	 * @memberof ApiComboboxField
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bEditable - allow user enter custom text
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiComboboxField/Methods/SetEditable.js
	 */
	ApiComboboxField.prototype.SetEditable = function(bCommit) {
		return this.Field.SetEditable(bCommit)
	};

	/**
	 * Checks if field is editable.
	 * @memberof ApiComboboxField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiComboboxField/Methods/IsEditable.js
	 */
	ApiComboboxField.prototype.IsEditable = function(bCommit) {
		return this.Field.IsEditable(bCommit)
	};

	/**
	 * Sets number format for field.
	 * @memberof ApiComboboxField
	 * @typeofeditors ["PDFE"]
	 * @param {number} nDemical - number of decimals
	 * @param {NumberSepStyle} - number separate style
	 * @param {NumberNegStyle} - number negative style
	 * @param {string} sCurrency - currency sybmol
	 * @param {boolean} bCurrencyPrepend - If true, places the currency symbol before the number (e.g., $1,234.56); 
 	 * if false, places it after (e.g., 1,234.56$).
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiComboboxField/Methods/SetNumberFormat.js
	 */
	ApiComboboxField.prototype.SetNumberFormat = function(nDemical, sSepStyle, sNegStyle, sCurrency, bCurrencyPrepend) {
		this.Field.ClearFormat();

		let aActionsFormat = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": "AFNumber_Format(" + nDemical + "," + private_GetInnerNumberSeparateType(sSepStyle) + "," + private_GetInnerNumberNegType(sNegStyle) + "," + "0" + ',"' + sCurrency + '",' + bCurrencyPrepend + ");"
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Format, aActionsFormat);

		let aActionsKeystroke = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": "AFNumber_Keystroke(" + nDemical + "," + private_GetInnerNumberSeparateType(sSepStyle) + "," + private_GetInnerNumberNegType(sNegStyle) + "," + "0" + ',"' + sCurrency + '",' + bCurrencyPrepend + ");"
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke, aActionsKeystroke);

		this.Field.Commit();

		return true;
	};

	/**
	 * Sets percentage format for field.
	 * @memberof ApiComboboxField
	 * @typeofeditors ["PDFE"]
	 * @param {number} nDemical - number of decimals
	 * @param {NumberSepStyle} - number separate style
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiComboboxField/Methods/SetPercentageFormat.js
	 */
	ApiComboboxField.prototype.SetPercentageFormat = function(nDemical, sSepStyle) {
		this.Field.ClearFormat();

		let aActionsFormat = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": "AFPercent_Format(" + nDemical + "," + private_GetInnerNumberSeparateType(sSepStyle) + ");"
		}]
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Format, aActionsFormat);

		let aActionsKeystroke = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": "AFPercent_Keystroke(" + nDemical + "," + private_GetInnerNumberSeparateType(sSepStyle) + ");"
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke, aActionsKeystroke);
		this.Field.Commit();

		return true;
	};

	/**
	 * Sets date format for field.
	 * @memberof ApiComboboxField
	 * @typeofeditors ["PDFE"]
	 * @param {string} sFormat - date format (e.g. "dd.mm.yyyy")
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiComboboxField/Methods/SetDateFormat.js
	 */
	ApiComboboxField.prototype.SetDateFormat = function(sFormat) {
		this.Field.ClearFormat();

		let aActionsFormat = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": 'AFDate_Format("' + sFormat + '");'
		}]
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Format, aActionsFormat);

		let aActionsKeystroke = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": 'AFDate_Keystroke("' + sFormat + '");'
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke, aActionsKeystroke);
		this.Field.Commit();

		return true;
	};

	/**
	 * Sets time format for field.
	 * @memberof ApiComboboxField
	 * @typeofeditors ["PDFE"]
	 * @param {TimeFormat} sFormat - available time format
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiComboboxField/Methods/SetTimeFormat.js
	 */
	ApiComboboxField.prototype.SetTimeFormat = function(sFormat) {
		this.Field.ClearFormat();

		let aActionsFormat = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": 'AFTime_Format(' + private_GetInnerTimeFormatType(sFormat) + ');'
		}]
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Format, aActionsFormat);

		let aActionsKeystroke = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": 'AFTime_Keystroke(' + private_GetInnerTimeFormatType(sFormat) + ');'
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke, aActionsKeystroke);
		this.Field.Commit();

		return true;
	};

	/**
	 * Sets special format for field.
	 * @memberof ApiComboboxField
	 * @typeofeditors ["PDFE"]
	 * @param {PsfFormat} sFormat - the formatting style to apply to the value
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiComboboxField/Methods/SetSpecialFormat.js
	 */
	ApiComboboxField.prototype.SetSpecialFormat = function(sFormat) {
		this.Field.ClearFormat();
				
		let aActionsFormat = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": "AFSpecial_Format(" + private_GetInnerSpecialPsfType(sFormat) + ");"
		}]
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Format, aActionsFormat);

		let aActionsKeystroke = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": "AFSpecial_Keystroke(" + private_GetInnerSpecialPsfType(sFormat) + ");"
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke, aActionsKeystroke);
		this.Field.Commit();

		return true;
	};

	/**
	 * Sets mask for field.
	 * @memberof ApiComboboxField
	 * @typeofeditors ["PDFE"]
	 * @param {string} sMask - field mask (e.g. "(999)999-9999")
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiComboboxField/Methods/SetMask.js
	 */
	ApiComboboxField.prototype.SetMask = function(sMask) {
		this.Field.ClearFormat();
		this.Field.SetArbitaryMask(sMask);
		this.Field.Commit();

		return true;
	};

	/**
	 * Sets regular expression for field.
	 * @memberof ApiComboboxField
	 * @typeofeditors ["PDFE"]
	 * @param {string} sReg - field regular expression (e.g. "\\S+@\\S+\\.\\S+")
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiComboboxField/Methods/SetRegularExp.js
	 */
	ApiComboboxField.prototype.SetRegularExp = function(sReg) {
		this.Field.ClearFormat();
		this.Field.SetRegularExp(sReg);
		this.Field.Commit();

		return true;
	};

	/**
	 * Clears format of field.
	 * @memberof ApiComboboxField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiComboboxField/Methods/ClearFormat.js
	 */
	ApiComboboxField.prototype.ClearFormat = function() {
		this.Field.ClearFormat();
		this.Field.Commit();

		return true;
	};

	/**
	 * Sets validate range for field.
	 * <note> Can only be applied to fields with a percentage or number format. </note>
	 * @memberof ApiComboboxField
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} [bGreaterThan=false] - If true, enables minimum value check using `nGreaterThan`.
	 * @param {number} nGreaterThan - Minimum allowed value (inclusive or exclusive based on implementation).
	 * @param {boolean} [bLessThan=false] - If true, enables maximum value check using `nLessThan`.
	 * @param {number} nLessThan - Maximum allowed value (inclusive or exclusive based on implementation).
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiComboboxField/Methods/SetValidateRange.js
	 */
	ApiComboboxField.prototype.SetValidateRange = function(bGreaterThan, nGreaterThan, bLessThan, nLessThan) {
		if (false == this.Field.IsNumberFormat()) {
			return false;
		}
		
		if (bGreaterThan == undefined) {
			bGreaterThan = false;
		}
		if (bLessThan == undefined) {
			bLessThan = false;
		}

		let aActionsValidate = [{
			"S": AscPDF.ACTIONS_TYPES.JavaScript,
			"JS": 'AFRange_Validate(' + bGreaterThan +  ',' + nGreaterThan + ',' + bLessThan + ',' + nLessThan +  ');'
		}];
		this.Field.SetActions(AscPDF.FORMS_TRIGGERS_TYPES.Validate, aActionsValidate);

		return true;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiListboxField
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a listbox field.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseListField}
	 */
	function ApiListboxField(oField) {
		ApiBaseListField.call(this, oField);
	}

	ApiListboxField.prototype = Object.create(ApiBaseListField.prototype);
	ApiListboxField.prototype.constructor = ApiListboxField;

	/**
	 * Returns a type of the ApiListboxField class.
	 * @memberof ApiListboxField
	 * @typeofeditors ["PDFE"]
	 * @returns {"listboxField"}
	 * @see office-js-api/Examples/PDF/ApiListboxField/Methods/GetClassType.js
	 */
	ApiListboxField.prototype.GetClassType = function() {
		return "listboxField";
	};

	/**
	 * Sets field multiselect prop.
	 * @memberof ApiListboxField
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bMulti - allow user select multi values
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiListboxField/Methods/SetMultipleSelection.js
	 */
	ApiListboxField.prototype.SetMultipleSelection = function(bMulti) {
		return this.Field.SetMultipleSelection(bMulti)
	};

	/**
	 * Checks if field is multiselect.
	 * @memberof ApiListboxField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiListboxField/Methods/IsMultipleSelection.js
	 */
	ApiListboxField.prototype.IsMultipleSelection = function(bMulti) {
		return this.Field.IsMultipleSelection(bMulti)
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiCheckboxField
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a checkbox field.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseField}
	 */
	function ApiCheckboxField(oField) {
		ApiBaseField.call(this, oField);
	}

	ApiCheckboxField.prototype = Object.create(ApiBaseField.prototype);
	ApiCheckboxField.prototype.constructor = ApiCheckboxField;

	/**
	 * Returns a type of the ApiCheckboxField class.
	 * @memberof ApiCheckboxField
	 * @typeofeditors ["PDFE"]
	 * @returns {"checkboxField"}
	 * @see office-js-api/Examples/PDF/ApiCheckboxField/Methods/GetClassType.js
	 */
	ApiCheckboxField.prototype.GetClassType = function() {
		return "checkboxField";
	};

	/**
	 * Sets field toggle to off prop.
	 * @memberof ApiCheckboxField
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bToggle - can toggle to off
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiCheckboxField/Methods/SetToggleToOff.js
	 */
	ApiCheckboxField.prototype.SetToggleToOff = function(bToggle) {
		return this.Field.SetNoToggleToOff(!bToggle);
	};

	/**
	 * Checks if field is toggle to off.
	 * @memberof ApiCheckboxField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiCheckboxField/Methods/IsToggleToOff.js
	 */
	ApiCheckboxField.prototype.IsToggleToOff = function() {
		return !this.Field.IsNoToggleToOff();
	};

	/**
	 * Adds options to checkbox group.
	 * @memberof ApiCheckboxField
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPage - page to add option
	 * @param {WidgetRect} - rect of new option
	 * @param {string} [sExportValue] - option checked value
	 * @returns {ApiCheckboxWidget}
	 * @see office-js-api/Examples/PDF/ApiCheckboxField/Methods/AddOption.js
	 */
	ApiCheckboxField.prototype.AddOption = function(nPage, aRect, sExportValue) {
		if (!sExportValue) {
			return null;
		}

		let oDoc = private_GetLogicDocument();

		let oField;
		if (this.GetClassType() == 'checkboxField') {
			oField = oDoc.CreateCheckboxField();
		}
		else {
			oField = oDoc.CreateRadiobuttonField();
		}

		oField.SetRect(aRect);
		oField.SetPartialName(this.GetFullName());
		oDoc.AddField(oField, nPage);

		if (sExportValue) {
			oField.SetExportValue(sExportValue);
		}

		return new ApiRadiobuttonField(oField);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiRadiobuttonField
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a radiobutton field.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiCheckboxField}
	 */
	function ApiRadiobuttonField(oField) {
		ApiCheckboxField.call(this, oField);
	}

	ApiRadiobuttonField.prototype = Object.create(ApiCheckboxField.prototype);
	ApiRadiobuttonField.prototype.constructor = ApiRadiobuttonField;

	/**
	 * Returns a type of the ApiRadiobuttonField class.
	 * @memberof ApiRadiobuttonField
	 * @typeofeditors ["PDFE"]
	 * @returns {"radiobuttonField"}
	 * @see office-js-api/Examples/PDF/ApiRadiobuttonField/Methods/GetClassType.js
	 */
	ApiRadiobuttonField.prototype.GetClassType = function() {
		return "radiobuttonField";
	};

	/**
	 * Sets field in unison prop.
	 * @memberof ApiRadiobuttonField
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bInUnison - will fields with the same export value be checked at the same time
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiRadiobuttonField/Methods/SetCheckInUnison.js
	 */
	ApiRadiobuttonField.prototype.SetCheckInUnison = function(bInUnison) {
		return this.Field.SetRadiosInUnison(bInUnison);
	};

	/**
	 * Checks if field will check in unison.
	 * @memberof ApiRadiobuttonField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiRadiobuttonField/Methods/IsCheckInUnison.js
	 */
	ApiRadiobuttonField.prototype.IsCheckInUnison = function() {
		return this.Field.SetRadiosInUnison();
	};

	/**
	 * Class representing a checkbox field widget.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 */
	function ApiCheckboxWidget(oField) {
		ApiBaseWidget.call(this, oField);
	}

	ApiCheckboxWidget.prototype = Object.create(ApiBaseWidget.prototype);
	ApiCheckboxWidget.prototype.constructor = ApiCheckboxWidget;

	/**
	 * Returns a type of the ApiCheckboxWidget class.
	 * @memberof ApiCheckboxWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {"page"}
	 * @see office-js-api/Examples/PDF/ApiCheckboxWidget/Methods/GetClassType.js
	 */
	ApiCheckboxWidget.prototype.GetClassType = function() {
		return "checkboxWidget";
	};

	/**
	 * Sets checkbox widget checked.
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bChecked
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiCheckboxWidget/Methods/SetChecked.js
	 */
	ApiCheckboxWidget.prototype.SetChecked = function(bChecked) {
		let oDoc = private_GetLogicDocument();
		if (this.Field.IsChecked() == bChecked) {
			return true;
		}

		this.Field.SetChecked(bChecked);
		this.Field.SetNeedCommit(true);
		oDoc.private_CommitField(this.Field);

		return true;
	};

	/**
	 * Checks if checkbox widget is checked.
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiCheckboxWidget/Methods/IsChecked.js
	 */
	ApiCheckboxWidget.prototype.IsChecked = function() {
		return this.Field.IsChecked();
	};

	/**
	 * Sets widget checkbox style.
	 * @typeofeditors ["PDFE"]
	 * @param {CheckStyle} sStyle
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiCheckboxWidget/Methods/SetCheckStyle.js
	 */
	ApiCheckboxWidget.prototype.SetCheckStyle = function(sStyle) {
		let nType = private_GetInnerCheckStyle(sStyle);
		if (undefined == nType) {
			return false;
		}

		this.Field.SetStyle(nType);

		return true;
	};

	/**
	 * Gets widget checkbox style.
	 * @typeofeditors ["PDFE"]
	 * @returns {CheckStyle}
	 * @see office-js-api/Examples/PDF/ApiCheckboxWidget/Methods/GetCheckStyle.js
	 */
	ApiCheckboxWidget.prototype.GetCheckStyle = function() {
		return private_GetStrCheckStyle(this.Field.GetStyle());
	};

	/**
	 * Sets widget export value.
	 * @typeofeditors ["PDFE"]
	 * @param {string} sValue
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiCheckboxWidget/Methods/SetExportValue.js
	 */
	ApiCheckboxWidget.prototype.SetExportValue = function(sValue) {
		if (!sValue) {
			return false;
		}

		this.Field.SetExportValue(sValue);
		return true;
	};

	/**
	 * Gets widget export value.
	 * @typeofeditors ["PDFE"]
	 * @returns {string}
	 * @see office-js-api/Examples/PDF/ApiCheckboxWidget/Methods/GetExportValue.js
	 */
	ApiCheckboxWidget.prototype.GetExportValue = function() {
		return this.Field.GetExportValue();
	};

	/**
	 * Sets widget checked by default.
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bChecked
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiCheckboxWidget/Methods/SetCheckedByDefault.js
	 */
	ApiCheckboxWidget.prototype.SetCheckedByDefault = function(bChecked) {
		if (bChecked) {
			this.Field.SetDefaultValue(this.Field.GetExportValue());
		}
		else {
			this.Field.SetDefaultValue(undefined);
		}
		
		return true;
	};

	/**
	 * Checks if widget is checked by default.
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiCheckboxWidget/Methods/IsCheckedByDefault.js
	 */
	ApiCheckboxWidget.prototype.IsCheckedByDefault = function() {
		return this.Field.GetDefaultValue() === this.Field.GetExportValue();
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiButtonField
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a button field.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseField}
	 */
	function ApiButtonField(oField) {
		ApiBaseField.call(this, oField);
	}

	ApiButtonField.prototype = Object.create(ApiBaseField.prototype);
	ApiButtonField.prototype.constructor = ApiButtonField;

	/**
	 * Returns a type of the ApiButtonField class.
	 * @memberof ApiButtonField
	 * @typeofeditors ["PDFE"]
	 * @returns {"buttonField"}
	 * @see office-js-api/Examples/PDF/ApiButtonField/Methods/GetClassType.js
	 */
	ApiButtonField.prototype.GetClassType = function() {
		return "buttonField";
	};

	/**
	 * Class representing a button widget.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseWidget}
	 */
	function ApiButtonWidget(oField) {
		ApiBaseWidget.call(this, oField);
	}

	ApiButtonWidget.prototype = Object.create(ApiBaseWidget.prototype);
	ApiButtonWidget.prototype.constructor = ApiButtonWidget;

	/**
	 * Returns a type of the ApiButtonWidget class.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {"page"}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/GetClassType.js
	 */
	ApiButtonWidget.prototype.GetClassType = function() {
		return "buttonWidget";
	};

	/**
	 * Sets button widget layout type
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {ButtonLayout} sType - button layout type
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/SetLayout.js
	 */
	ApiButtonWidget.prototype.SetLayout = function(sType) {
		if (false == Object.keys(position).includes(sType)) {
			return false;
		}

		this.Field.SetLayout(position[sType]);
		return true;
	};

	/**
	 * Gets button widget layout type
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {ButtonLayout}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/GetLayout.js
	 */
	ApiButtonWidget.prototype.GetLayout = function() {
		let nType = this.Field.GetLayout();
		return Object.keys(position).find(function(key) {
			return position[key] === nType;
		});
	};

	/**
	 * Sets button widget scale when type
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {ButtonScaleWhen} sType - button widget scale when type
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/SetScaleWhen.js
	 */
	ApiButtonWidget.prototype.SetScaleWhen = function(sType) {
		if (false == Object.keys(scaleWhen).includes(sType)) {
			return false;
		}

		this.Field.SetScaleWhen(scaleWhen[sType]);
		return true;
	};

	/**
	 * Gets button widget scale when type
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {ButtonScaleWhen}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/GetScaleWhen.js
	 */
	ApiButtonWidget.prototype.GetScaleWhen = function() {
		let nType = this.Field.GetScaleWhen();
		return Object.keys(scaleWhen).find(function(key) {
			return scaleWhen[key] === nType;
		});
	};

	/**
	 * Sets button widget scale how type
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {ButtonScaleHow} sType - button widget scale how type
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/SetScaleHow.js
	 */
	ApiButtonWidget.prototype.SetScaleHow = function(sType) {
		if (false == Object.keys(scaleHow).includes(sType)) {
			return false;
		}

		this.Field.SetScaleHow(scaleHow[sType]);
		return true;
	};

	/**
	 * Gets button widget scale when type
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {ButtonScaleHow}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/GetScaleHow.js
	 */
	ApiButtonWidget.prototype.GetScaleHow = function() {
		let nType = this.Field.GetScaleHow();
		return Object.keys(scaleHow).find(function(key) {
			return scaleHow[key] === nType;
		});
	};

	/**
	 * Sets button widget fit bounds.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bFit
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/SetFitBounds.js
	 */
	ApiButtonWidget.prototype.SetFitBounds = function(bFit) {
		this.Field.SetFitBounds(bFit);
		return true;
	};

	/**
	 * Checks if button widget is fit bounds.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/IsFitBounds.js
	 */
	ApiButtonWidget.prototype.IsFitBounds = function() {
		return this.Field.IsButtonFitBounds();
	};

	/**
	 * Sets button widget icon x position.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {percentage} nPosX
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/SetIconXPos.js
	 */
	ApiButtonWidget.prototype.SetIconXPos = function(nPosX) {
		if (typeof(nPosX) !== "number" || nPosX < 0) {
			return false;
		}

		let oCurPos = this.Field.GetIconPosition();

		this.Field.SetIconPosition(nPosX / 100, oCurPos.Y);
		return true;
	};

	/**
	 * Gets button widget icon x position.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {percentage}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/GetIconXPos.js
	 */
	ApiButtonWidget.prototype.GetIconXPos = function() {
		let oCurPos = this.Field.GetIconPosition();

		return oCurPos.X * 100;
	};

	/**
	 * Sets button widget icon y position.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {percentage} nPosY
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/SetIconYPos.js
	 */
	ApiButtonWidget.prototype.SetIconYPos = function(nPosY) {
		if (typeof(nPosY) !== "number" || nPosY < 0) {
			return false;
		}

		let oCurPos = this.Field.GetIconPosition();

		this.Field.SetIconPosition(oCurPos.X, nPosY / 100);
		return true;
	};

	/**
	 * Gets button widget icon y position.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {percentage}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/GetIconYPos.js
	 */
	ApiButtonWidget.prototype.GetIconYPos = function() {
		let oCurPos = this.Field.GetIconPosition();

		return oCurPos.Y * 100;
	};

	/**
	 * Sets button widget behavior.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {ButtonBehavior} sType
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/SetBehavior.js
	 */
	ApiButtonWidget.prototype.SetBehavior = function(sType) {
		if (false == Object.keys(AscPDF.BUTTON_HIGHLIGHT_TYPES).includes(sType)) {
			return false;
		}

		this.Field.SetHighlight(private_GetInnerButtonBehaviorType(sType));
		return true;
	};

	/**
	 * Gets button widget behavior.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {ButtonBehavior}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/GetBehavior.js
	 */
	ApiButtonWidget.prototype.GetBehavior = function() {
		return private_GetStrButtonBehaviorType(this.Field.GetHighlight());
	};

	/**
	 * Sets label to button widget field.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {string} sLabel - button label
	 * @param {ButtonAppearance} [sApType='normal'] - for what state is the label set 
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/SetLabel.js
	 */
	ApiButtonWidget.prototype.SetLabel = function(sLabel, sApType) {
		if (this.Field.GetLayout() == position["iconOnly"]) {
			return false;
		}

		if (undefined == sApType) {
			sApType = 'normal';
		}

		if (false == ['normal', 'down', 'hover'].includes(sApType)) {
			return false;
		}

		this.Field.SetCaption(sLabel, private_GetInnerButtonApType(sApType));
		return true;
	};

	/**
	 * Gets label from button widget field.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {ButtonAppearance} [sApType='normal'] - from what state is the label set 
	 * @returns {?string}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/GetLabel.js
	 */
	ApiButtonWidget.prototype.GetLabel = function(sApType) {
		if (this.Field.GetLayout() == position["iconOnly"]) {
			return null;
		}

		if (undefined == sApType) {
			sApType = 'normal';
		}

		if (false == ['normal', 'down', 'hover'].includes(sApType)) {
			return null;
		}

		return this.Field.GetCaption(private_GetInnerButtonApType(sApType));;
	};

	/**
	 * Sets image to button widget field.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {string} [sImageUrl=''] - image url
	 * @param {ButtonAppearance} [sApType='normal'] - for what state is the picture set 
	 * @returns {boolean}
	 * @see office-js-api/Examples/PDF/ApiButtonWidget/Methods/SetImage.js
	 */
	ApiButtonWidget.prototype.SetImage = function(sImageUrl, sApType) {
		if (this.Field.GetLayout() == position["textOnly"]) {
			return false;
		}

		if (undefined == sApType) {
			sApType = 'normal';
		}

		if (undefined == sImageUrl) {
			sImageUrl = '';
		}

		if (false == ['normal', 'down', 'hover'].includes(sApType)) {
			return false;
		}

		this.Field.SetImageRasterId(sImageUrl, private_GetInnerButtonApType(sApType));
		this.Field.SetNeedUpdateImage(true);

		return true;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiRGBColor
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing an RGB Color.
	 * @constructor
	 */
	function ApiRGBColor(r, g, b) {
		this.Color = AscFormat.CreateUniColorRGB(r, g, b);
	}

	/**
	 * Returns a type of the ApiRGBColor class.
	 * @memberof ApiRGBColor
	 * @typeofeditors ["PDFE"]
	 * @returns {"rgbColor"}
	 * @see office-js-api/Examples/PDF/ApiRGBColor/Methods/GetClassType.js
	 */
	ApiRGBColor.prototype.GetClassType = function() {
		return "rgbColor";
	};

	
	/**
	 * Represents R component of color.
	 * @memberof ApiRGBColor
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 */
	Object.defineProperty(ApiRGBColor.prototype, "R", {
		get: function() {
			if (!this.Color.color || !this.Color.color.RGBA)
				return 0;
			
			let c = this.Color.color.RGBA;
			return c.R;
		},
		set: function(r) {
			if (!this.Color.color || !this.Color.color.RGBA)
				return;
			
			this.Color.color.RGBA.R = r;
		}
	});

	/**
	 * Represents G component of color.
	 * @memberof ApiRGBColor
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 */
	Object.defineProperty(ApiRGBColor.prototype, "G", {
		get: function() {
			if (!this.Color.color || !this.Color.color.RGBA)
				return 0;
			
			let c = this.Color.color.RGBA;
			return c.G;
		},
		set: function(g) {
			if (!this.Color.color || !this.Color.color.RGBA)
				return;
			
			this.Color.color.RGBA.G = g;
		}
	});

	/**
	 * Represents B component of color.
	 * @memberof ApiRGBColor
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 */
	Object.defineProperty(ApiRGBColor.prototype, "B", {
		get: function() {
			if (!this.Color.color || !this.Color.color.RGBA)
				return 0;
			
			let c = this.Color.color.RGBA;
			return c.B;
		},
		set: function(b) {
			if (!this.Color.color || !this.Color.color.RGBA)
				return;
			
			this.Color.color.RGBA.B = b;
		}
	});
	
	function private_GetLogicDocument() {
		return Asc.editor.getPDFDoc();
	}

	function private_GetFieldApi(field) {
		if (!field) {
			return null;
		}

		switch (field.GetType()) {
			case AscPDF.FIELD_TYPES.button: {
				return new ApiButtonField(field);
			}
			case AscPDF.FIELD_TYPES.radiobutton: {
				return new ApiRadiobuttonField(field);
			}
			case AscPDF.FIELD_TYPES.checkbox: {
				return new ApiCheckboxField(field);
			}
			case AscPDF.FIELD_TYPES.text: {
				return new ApiTextField(field);
			}
			case AscPDF.FIELD_TYPES.combobox: {
				return new ApiComboboxField(field);
			}
			case AscPDF.FIELD_TYPES.listbox: {
				return new ApiListboxField(field);
			}
		}
	}

	function private_GetWidgetApi(field) {
		if (!field) {
			return null;
		}

		switch (field.GetType()) {
			case AscPDF.FIELD_TYPES.button: {
				return new ApiButtonWidget(field);
			}
			case AscPDF.FIELD_TYPES.radiobutton:
			case AscPDF.FIELD_TYPES.checkbox: {
				return new ApiCheckboxWidget(field);
			}
			case AscPDF.FIELD_TYPES.text:
			case AscPDF.FIELD_TYPES.combobox: {
				return new ApiTextWidget(field);
			}
			case AscPDF.FIELD_TYPES.listbox: {
				return new ApiBaseWidget(field);
			}
		}
	}

	function private_GetInnerCheckStyle(sStyle) {
		switch (sStyle) {
			case "check": {
				return AscPDF.CHECKBOX_STYLES.check;
			}
			case "cross": {
				return AscPDF.CHECKBOX_STYLES.cross;
			}
			case "diamond": {
				return AscPDF.CHECKBOX_STYLES.diamond;
			}
			case "circle": {
				return AscPDF.CHECKBOX_STYLES.circle;
			}
			case "star": {
				return AscPDF.CHECKBOX_STYLES.star;
			}
			case "square": {
				return AscPDF.CHECKBOX_STYLES.square;
			}
		}
	}

	function private_GetStrCheckStyle(nStyle) {
		switch (nStyle) {
			case AscPDF.CHECKBOX_STYLES.check: {
				return "check";
			}
			case AscPDF.CHECKBOX_STYLES.cross: {
				return "cross";
			}
			case AscPDF.CHECKBOX_STYLES.diamond: {
				return "diamond";
			}
			case AscPDF.CHECKBOX_STYLES.circle: {
				return "circle";
			}
			case AscPDF.CHECKBOX_STYLES.star: {
				return "star";
			}
			case AscPDF.CHECKBOX_STYLES.square: {
				return "square";
			}
		}
	}

	function private_GetInnerBorderWidth(sBorderWidth) {
		switch (sBorderWidth) {
			case "none": {
				return AscPDF.BORDER_WIDTH.none;
			}
			case "thin": {
				return AscPDF.BORDER_WIDTH.thin;
			}
			case "medium": {
				return AscPDF.BORDER_WIDTH.medium;
			}
			case "thick": {
				return AscPDF.BORDER_WIDTH.thick;
			}
		}
	}

	function private_GetStrBorderWidth(nBorderWidth) {
		switch (nBorderWidth) {
			case AscPDF.BORDER_WIDTH.none: {
				return "none";
			}
			case AscPDF.BORDER_WIDTH.thin: {
				return "thin";
			}
			case AscPDF.BORDER_WIDTH.medium: {
				return "medium";
			}
			case AscPDF.BORDER_WIDTH.thick: {
				return "thick";
			}
		}
	}

	function private_GetInnerBorderStyle(sBorderStyle) {
		switch (sBorderStyle) {
			case "solid": {
				return AscPDF.BORDER_TYPES.solid;
			}
			case "beveled": {
				return AscPDF.BORDER_TYPES.beveled;
			}
			case "dashed": {
				return AscPDF.BORDER_TYPES.dashed;
			}
			case "inset": {
				return AscPDF.BORDER_TYPES.inset;
			}
			case "underline": {
				return AscPDF.BORDER_TYPES.underline;
			}
		}
	}

	function private_GetStrBorderStyle(nBorderStyle) {
		switch (nBorderStyle) {
			case AscPDF.BORDER_TYPES.solid: {
				return "solid";
			}
			case AscPDF.BORDER_TYPES.beveled: {
				return "beveled";
			}
			case AscPDF.BORDER_TYPES.dashed: {
				return "dashed";
			}
			case AscPDF.BORDER_TYPES.inset: {
				return "inset";
			}
			case AscPDF.BORDER_TYPES.underline: {
				return "underline";
			}
		}
	}

	function private_GetInnerButtonApType(sApType) {
		switch (sApType) {
			case "normal": {
				return AscPDF.APPEARANCE_TYPES.normal;
			}
			case "down": {
				return AscPDF.APPEARANCE_TYPES.mouseDown;
			}
			case "hover": {
				return AscPDF.APPEARANCE_TYPES.rollover;
			}
		}
	}

	function private_GetStrButtonApType(nApType) {
		switch (nApType) {
			case AscPDF.APPEARANCE_TYPES.normal: {
				return "normal";
			}
			case AscPDF.APPEARANCE_TYPES.mouseDown: {
				return "down";
			}
			case AscPDF.APPEARANCE_TYPES.rollover: {
				return "hover";
			}
		}
	}

	function private_GetInnerButtonBehaviorType(sType) {
		switch (sType) {
			case "none": {
				return AscPDF.BUTTON_HIGHLIGHT_TYPES.none;
			}
			case "invert": {
				return AscPDF.BUTTON_HIGHLIGHT_TYPES.invert;
			}
			case "push": {
				return AscPDF.BUTTON_HIGHLIGHT_TYPES.push;
			}
			case "outline": {
				return AscPDF.BUTTON_HIGHLIGHT_TYPES.outline;
			}
		}
	}

	function private_GetStrButtonBehaviorType(sType) {
		switch (sType) {
			case AscPDF.BUTTON_HIGHLIGHT_TYPES.none: {
				return "none";
			}
			case AscPDF.BUTTON_HIGHLIGHT_TYPES.invert: {
				return "invert";
			}
			case AscPDF.BUTTON_HIGHLIGHT_TYPES.push: {
				return "push";
			}
			case AscPDF.BUTTON_HIGHLIGHT_TYPES.outline: {
				return "outline";
			}
		}
	}

	function private_GetInnerNumberSeparateType(sType) {
		switch (sType) {
			case "us": {
				return AscPDF.SeparatorStyle.COMMA_DOT;
			}
			case "plain": {
				return AscPDF.SeparatorStyle.NO_SEPARATOR;
			}
			case "euro": {
				return AscPDF.SeparatorStyle.DOT_COMMA;
			}
			case "europlain": {
				return AscPDF.SeparatorStyle.NO_SEPARATOR_COMMA;
			}
			case "ch": {
				return AscPDF.SeparatorStyle.APOSTROPHE_DOT;
			}
		}
	}

	function private_GetStrNumberSeparateType(nType) {
		switch (nType) {
			case AscPDF.SeparatorStyle.COMMA_DOT: {
				return "us";
			}
			case AscPDF.SeparatorStyle.NO_SEPARATOR: {
				return "plain";
			}
			case AscPDF.SeparatorStyle.DOT_COMMA: {
				return "euro";
			}
			case AscPDF.SeparatorStyle.NO_SEPARATOR_COMMA: {
				return "europlain";
			}
			case AscPDF.SeparatorStyle.APOSTROPHE_DOT: {
				return "ch";
			}
		}
	}

	function private_GetInnerNumberNegType(sType) {
		switch (sType) {
			case "black-minus": {
				return AscPDF.NegativeStyle.BLACK_MINUS;
			}
			case "red-minus": {
				return AscPDF.NegativeStyle.RED_MINUS;
			}
			case "black-parens": {
				return AscPDF.NegativeStyle.PARENS_BLACK;
			}
			case "red-parens": {
				return AscPDF.NegativeStyle.PARENS_RED;
			}
		}
	}

	function private_GetStrNumberNegType(nType) {
		switch (nType) {
			case AscPDF.NegativeStyle.BLACK_MINUS: {
				return "black-minus";
			}
			case AscPDF.NegativeStyle.RED_MINUS: {
				return "red-minus";
			}
			case AscPDF.NegativeStyle.PARENS_BLACK: {
				return "black-parens";
			}
			case AscPDF.NegativeStyle.PARENS_RED: {
				return "red-parens";
			}
		}
	}

	function private_GetInnerSpecialPsfType(sType) {
		switch (sType) {
			case "zip": {
				return AscPDF.SpecialFormatType.ZIP_CODE;
			}
			case "zip+4": {
				return AscPDF.SpecialFormatType.ZIP_PLUS_4;
			}
			case "phone": {
				return AscPDF.SpecialFormatType.PHONE;
			}
			case "ssn": {
				return AscPDF.SpecialFormatType.SSN;
			}
		}
	}

	function private_GetStrSpecialPsfType(nType) {
		switch (nType) {
			case AscPDF.SpecialFormatType.ZIP_CODE: {
				return "zip";
			}
			case AscPDF.SpecialFormatType.ZIP_PLUS_4: {
				return "zip+4";
			}
			case AscPDF.SpecialFormatType.PHONE: {
				return "phone";
			}
			case AscPDF.SpecialFormatType.SSN: {
				return "ssn";
			}
		}
	}

	function private_GetInnerTimeFormatType(sType) {
		switch (sType) {
			case "24HR:MM": {
				return AscPDF.TimeFormatType["24HR:MM"];
			}
			case "12HR:MM": {
				return AscPDF.TimeFormatType["12HR:MM"];
			}
			case "24HR:MM:SS": {
				return AscPDF.TimeFormatType["24HR:MM:SS"];
			}
			case "12HR:MM:SS": {
				return AscPDF.TimeFormatType["12HR:MM:SS"];
			}
		}
	}

	function private_GetStrTimeFormatType(nType) {
		switch (nType) {
			case AscPDF.TimeFormatType["24HR:MM"]: {
				return "24HR:MM";
			}
			case AscPDF.TimeFormatType["12HR:MM"]: {
				return "12HR:MM";
			}
			case AscPDF.TimeFormatType["24HR:MM:SS"]: {
				return "24HR:MM:SS";
			}
			case AscPDF.TimeFormatType["12HR:MM:SS"]: {
				return "12HR:MM:SS";
			}
		}
	}

	function private_GetInnerColorByRGB(r, g, b) {
		return [r / 255, g / 255, b / 255];
	}

	// Api
	Api.prototype["GetDocument"]						= Api.prototype.GetDocument
	Api.prototype["CreateRGBColor"]						= Api.prototype.CreateRGBColor

	// ApiDocument
	ApiDocument.prototype["GetClassType"]				= ApiDocument.prototype.GetClassType;
	ApiDocument.prototype["AddPage"]					= ApiDocument.prototype.AddPage;
	ApiDocument.prototype["GetPage"]					= ApiDocument.prototype.GetPage;
	ApiDocument.prototype["RemovePage"]					= ApiDocument.prototype.RemovePage;
	ApiDocument.prototype["GetPagesCount"]				= ApiDocument.prototype.GetPagesCount;
	ApiDocument.prototype["AddTextField"]				= ApiDocument.prototype.AddTextField;
	ApiDocument.prototype["AddDateField"]				= ApiDocument.prototype.AddDateField;
	ApiDocument.prototype["AddImageField"]				= ApiDocument.prototype.AddImageField;
	ApiDocument.prototype["AddCheckboxField"]			= ApiDocument.prototype.AddCheckboxField;
	ApiDocument.prototype["AddRadiobuttonField"]		= ApiDocument.prototype.AddRadiobuttonField;
	ApiDocument.prototype["AddComboboxField"]			= ApiDocument.prototype.AddComboboxField;
	ApiDocument.prototype["AddListboxField"]			= ApiDocument.prototype.AddListboxField;
	ApiDocument.prototype["GetAllFields"]				= ApiDocument.prototype.GetAllFields;
	ApiDocument.prototype["GetFieldByName"]				= ApiDocument.prototype.GetFieldByName;

	// ApiPage
	ApiPage.prototype["GetClassType"]					= ApiPage.prototype.GetClassType;
	ApiPage.prototype["SetRotate"]						= ApiPage.prototype.SetRotate;
	ApiPage.prototype["GetRotate"]						= ApiPage.prototype.GetRotate;
	ApiPage.prototype["GetIndex"]						= ApiPage.prototype.GetIndex;
	ApiPage.prototype["GetAllWidgets"]					= ApiPage.prototype.GetAllWidgets;

	// ApiBaseField
	ApiBaseField.prototype["SetFullName"]				= ApiBaseField.prototype.SetFullName;
	ApiBaseField.prototype["GetFullName"]				= ApiBaseField.prototype.GetFullName;
	ApiBaseField.prototype["SetPartialName"]			= ApiBaseField.prototype.SetPartialName;
	ApiBaseField.prototype["GetPartialName"]			= ApiBaseField.prototype.GetPartialName;
	ApiBaseField.prototype["SetRequired"]				= ApiBaseField.prototype.SetRequired;
	ApiBaseField.prototype["IsRequired"]				= ApiBaseField.prototype.IsRequired;
	ApiBaseField.prototype["SetReadOnly"]				= ApiBaseField.prototype.SetReadOnly;
	ApiBaseField.prototype["IsReadOnly"]				= ApiBaseField.prototype.IsReadOnly;
	ApiBaseField.prototype["SetValue"]					= ApiBaseField.prototype.SetValue;
	ApiBaseField.prototype["GetValue"]					= ApiBaseField.prototype.GetValue;
	ApiBaseField.prototype["AddWidget"]					= ApiBaseField.prototype.AddWidget;
	ApiBaseField.prototype["GetAllWidgets"]				= ApiBaseField.prototype.GetAllWidgets;

	// ApiBaseWidget
	ApiBaseWidget.prototype["GetClassType"]				= ApiBaseWidget.prototype.GetClassType;
	ApiBaseWidget.prototype["SetBorderColor"]			= ApiBaseWidget.prototype.SetBorderColor;
	ApiBaseWidget.prototype["GetBorderColor"]			= ApiBaseWidget.prototype.GetBorderColor;
	ApiBaseWidget.prototype["SetBorderWidth"]			= ApiBaseWidget.prototype.SetBorderWidth;
	ApiBaseWidget.prototype["GetBorderWidth"]			= ApiBaseWidget.prototype.GetBorderWidth;
	ApiBaseWidget.prototype["SetBorderStyle"]			= ApiBaseWidget.prototype.SetBorderStyle;
	ApiBaseWidget.prototype["GetBorderStyle"]			= ApiBaseWidget.prototype.GetBorderStyle;
	ApiBaseWidget.prototype["SetBackgroundColor"]		= ApiBaseWidget.prototype.SetBackgroundColor;
	ApiBaseWidget.prototype["GetBackgroundColor"]		= ApiBaseWidget.prototype.GetBackgroundColor;
	ApiBaseWidget.prototype["SetTextColor"]				= ApiBaseWidget.prototype.SetTextColor;
	ApiBaseWidget.prototype["GetTextColor"]				= ApiBaseWidget.prototype.GetTextColor;
	ApiBaseWidget.prototype["SetTextSize"]				= ApiBaseWidget.prototype.SetTextSize;
	ApiBaseWidget.prototype["GetTextSize"]				= ApiBaseWidget.prototype.GetTextSize;
	ApiBaseWidget.prototype["SetAutoFit"]				= ApiBaseWidget.prototype.SetAutoFit;
	ApiBaseWidget.prototype["IsAutoFit"]				= ApiBaseWidget.prototype.IsAutoFit;
	ApiBaseWidget.prototype["Remove"]					= ApiBaseWidget.prototype.Remove;

	// ApiTextField
	ApiTextField.prototype["GetClassType"]				= ApiTextField.prototype.GetClassType;
	ApiTextField.prototype["SetMultiline"]				= ApiTextField.prototype.SetMultiline;
	ApiTextField.prototype["IsMultiline"]				= ApiTextField.prototype.IsMultiline;
	ApiTextField.prototype["SetCharLimit"]				= ApiTextField.prototype.SetCharLimit;
	ApiTextField.prototype["GetCharLimit"]				= ApiTextField.prototype.GetCharLimit;
	ApiTextField.prototype["SetComb"]					= ApiTextField.prototype.SetComb;
	ApiTextField.prototype["IsComb"]					= ApiTextField.prototype.IsComb;
	ApiTextField.prototype["SetScrollLongText"]			= ApiTextField.prototype.SetScrollLongText;
	ApiTextField.prototype["IsScrollLongText"]			= ApiTextField.prototype.IsScrollLongText;
	ApiTextField.prototype["SetNumberFormat"]			= ApiTextField.prototype.SetNumberFormat;
	ApiTextField.prototype["SetPercentageFormat"]		= ApiTextField.prototype.SetPercentageFormat;
	ApiTextField.prototype["SetDateFormat"]				= ApiTextField.prototype.SetDateFormat;
	ApiTextField.prototype["SetTimeFormat"]				= ApiTextField.prototype.SetTimeFormat;
	ApiTextField.prototype["SetSpecialFormat"]			= ApiTextField.prototype.SetSpecialFormat;
	ApiTextField.prototype["SetMask"]					= ApiTextField.prototype.SetMask;
	ApiTextField.prototype["SetRegularExp"]				= ApiTextField.prototype.SetRegularExp;
	ApiTextField.prototype["ClearFormat"]				= ApiTextField.prototype.ClearFormat;
	ApiTextField.prototype["SetValidateRange"]			= ApiTextField.prototype.SetValidateRange;

	// ApiTextWidget
	ApiTextWidget.prototype["GetClassType"]				= ApiTextWidget.prototype.GetClassType;
	ApiTextWidget.prototype["SetPlaceholder"]			= ApiTextWidget.prototype.SetPlaceholder;
	ApiTextWidget.prototype["GetPlaceholder"]			= ApiTextWidget.prototype.GetPlaceholder;
	ApiTextWidget.prototype["SetRegularExp"]			= ApiTextWidget.prototype.SetRegularExp;
	ApiTextWidget.prototype["GetRegularExp"]			= ApiTextWidget.prototype.GetRegularExp;

	// ApiBaseListField
	ApiBaseListField.prototype["AddOption"]				= ApiBaseListField.prototype.AddOption;
	ApiBaseListField.prototype["RemoveOption"]			= ApiBaseListField.prototype.RemoveOption;
	ApiBaseListField.prototype["MoveOption"]			= ApiBaseListField.prototype.MoveOption;
	ApiBaseListField.prototype["GetOption"]				= ApiBaseListField.prototype.GetOption;
	ApiBaseListField.prototype["GetOptions"]			= ApiBaseListField.prototype.GetOptions;
	ApiBaseListField.prototype["SetCommitOnSelChange"]	= ApiBaseListField.prototype.SetCommitOnSelChange;
	ApiBaseListField.prototype["IsCommitOnSelChange"]	= ApiBaseListField.prototype.IsCommitOnSelChange;
	ApiBaseListField.prototype["SetValueIndexes"]		= ApiBaseListField.prototype.SetValueIndexes;
	ApiBaseListField.prototype["GetValueIndexes"]		= ApiBaseListField.prototype.GetValueIndexes;

	// ApiComboboxField
	ApiComboboxField.prototype["GetClassType"]			= ApiComboboxField.prototype.GetClassType;
	ApiComboboxField.prototype["SetEditable"]			= ApiComboboxField.prototype.SetEditable;
	ApiComboboxField.prototype["IsEditable"]			= ApiComboboxField.prototype.IsEditable;
	ApiComboboxField.prototype["SetNumberFormat"]		= ApiComboboxField.prototype.SetNumberFormat;
	ApiComboboxField.prototype["SetPercentageFormat"]	= ApiComboboxField.prototype.SetPercentageFormat;
	ApiComboboxField.prototype["SetDateFormat"]			= ApiComboboxField.prototype.SetDateFormat;
	ApiComboboxField.prototype["SetTimeFormat"]			= ApiComboboxField.prototype.SetTimeFormat;
	ApiComboboxField.prototype["SetSpecialFormat"]		= ApiComboboxField.prototype.SetSpecialFormat;
	ApiComboboxField.prototype["SetMask"]				= ApiComboboxField.prototype.SetMask;
	ApiComboboxField.prototype["SetRegularExp"]			= ApiComboboxField.prototype.SetRegularExp;
	ApiComboboxField.prototype["ClearFormat"]			= ApiComboboxField.prototype.ClearFormat;
	ApiComboboxField.prototype["SetValidateRange"]		= ApiComboboxField.prototype.SetValidateRange;

	// ApiListboxField
	ApiListboxField.prototype["GetClassType"]			= ApiListboxField.prototype.GetClassType;
	ApiListboxField.prototype["SetMultipleSelection"]	= ApiListboxField.prototype.SetMultipleSelection;
	ApiListboxField.prototype["IsMultipleSelection"]	= ApiListboxField.prototype.IsMultipleSelection;

	// ApiCheckboxField
	ApiCheckboxField.prototype["GetClassType"]			= ApiCheckboxField.prototype.GetClassType;
	ApiCheckboxField.prototype["SetToggleToOff"]		= ApiCheckboxField.prototype.SetToggleToOff;
	ApiCheckboxField.prototype["IsToggleToOff"]			= ApiCheckboxField.prototype.IsToggleToOff;
	ApiCheckboxField.prototype["AddOption"]				= ApiCheckboxField.prototype.AddOption;

	// ApiRadiobuttonField
	ApiRadiobuttonField.prototype["GetClassType"]		= ApiRadiobuttonField.prototype.GetClassType;
	ApiRadiobuttonField.prototype["SetCheckInUnison"]	= ApiRadiobuttonField.prototype.SetCheckInUnison;
	ApiRadiobuttonField.prototype["IsCheckInUnison"]	= ApiRadiobuttonField.prototype.IsCheckInUnison;

	// ApiCheckboxWidget
	ApiCheckboxWidget.prototype["GetClassType"]			= ApiCheckboxWidget.prototype.GetClassType;
	ApiCheckboxWidget.prototype["SetChecked"]			= ApiCheckboxWidget.prototype.SetChecked;
	ApiCheckboxWidget.prototype["IsChecked"]			= ApiCheckboxWidget.prototype.IsChecked;
	ApiCheckboxWidget.prototype["SetCheckStyle"]		= ApiCheckboxWidget.prototype.SetCheckStyle;
	ApiCheckboxWidget.prototype["GetCheckStyle"]		= ApiCheckboxWidget.prototype.GetCheckStyle;
	ApiCheckboxWidget.prototype["SetExportValue"]		= ApiCheckboxWidget.prototype.SetExportValue;
	ApiCheckboxWidget.prototype["GetExportValue"]		= ApiCheckboxWidget.prototype.GetExportValue;
	ApiCheckboxWidget.prototype["SetCheckedByDefault"]	= ApiCheckboxWidget.prototype.SetCheckedByDefault;
	ApiCheckboxWidget.prototype["IsCheckedByDefault"]	= ApiCheckboxWidget.prototype.IsCheckedByDefault;

	// ApiButtonField
	ApiButtonField.prototype["GetClassType"]			= ApiButtonField.prototype.GetClassType;

	// ApiButtonWidget
	ApiButtonWidget.prototype["GetClassType"]			= ApiButtonWidget.prototype.GetClassType;
	ApiButtonWidget.prototype["SetLayout"]				= ApiButtonWidget.prototype.SetLayout;
	ApiButtonWidget.prototype["GetLayout"]				= ApiButtonWidget.prototype.GetLayout;
	ApiButtonWidget.prototype["SetScaleWhen"]			= ApiButtonWidget.prototype.SetScaleWhen;
	ApiButtonWidget.prototype["GetScaleWhen"]			= ApiButtonWidget.prototype.GetScaleWhen;
	ApiButtonWidget.prototype["SetScaleHow"]			= ApiButtonWidget.prototype.SetScaleHow;
	ApiButtonWidget.prototype["GetScaleHow"]			= ApiButtonWidget.prototype.GetScaleHow;
	ApiButtonWidget.prototype["SetFitBounds"]			= ApiButtonWidget.prototype.SetFitBounds;
	ApiButtonWidget.prototype["IsFitBounds"]			= ApiButtonWidget.prototype.IsFitBounds;
	ApiButtonWidget.prototype["SetIconXPos"]			= ApiButtonWidget.prototype.SetIconXPos;
	ApiButtonWidget.prototype["GetIconXPos"]			= ApiButtonWidget.prototype.GetIconXPos;
	ApiButtonWidget.prototype["SetIconYPos"]			= ApiButtonWidget.prototype.SetIconYPos;
	ApiButtonWidget.prototype["GetIconYPos"]			= ApiButtonWidget.prototype.GetIconYPos;
	ApiButtonWidget.prototype["SetBehavior"]			= ApiButtonWidget.prototype.SetBehavior;
	ApiButtonWidget.prototype["GetBehavior"]			= ApiButtonWidget.prototype.GetBehavior;
	ApiButtonWidget.prototype["SetLabel"]				= ApiButtonWidget.prototype.SetLabel;
	ApiButtonWidget.prototype["GetLabel"]				= ApiButtonWidget.prototype.GetLabel;
	ApiButtonWidget.prototype["SetImage"]				= ApiButtonWidget.prototype.SetImage;

	// ApiRGBColor
	ApiRGBColor.prototype["GetClassType"]				= ApiRGBColor.prototype.GetClassType;
}(window, null));


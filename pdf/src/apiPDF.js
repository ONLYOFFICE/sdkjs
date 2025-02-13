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

    const highlight     = AscPDF.Api.Objects.highlight;
    const style         = AscPDF.Api.Objects.style;
    const display       = AscPDF.Api.Objects.display;
    const border        = AscPDF.Api.Objects.border;

    /**
	 * A string that sets the trigger for the action. Values are:
	 * @typedef {"MouseUp" | "MouseDown" | "MouseEnter" | "MouseExit" | "OnFocus" | "OnBlur" | "Keystroke" | "Validate" | "Calculate" | "Format"} cTrigger
	 * For a list box, use the Keystroke trigger for the Selection Change event.
     */

    
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
            if (Object.values(border).includes(sValue)) {
                let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                aFields.forEach(function(field) {
                    field.SetBorderStyle(private_GetIntBorderStyle(sValue));
                });
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return private_GetStrBorderStyle(this.field.GetBorderStyle());
            else if (this.field.IsAllKidsWidgets())
                return private_GetStrBorderStyle(this.field.GetKid(0).GetBorderStyle());
            else
                throw Error("InvalidGetError: Field is not a widget");
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
            if (value && value.toString) {
                value = value.toString();

                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    this.field.SetDefaultValue(value);
                }
                else {
                    throw Error("InvalidGetError: Field is not a widget");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                return this.field.GetDefaultValue();
            }
            else {
                throw Error("InvalidGetError: Field is not a widget");
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
            if (Object.values(display).includes(nType)) {
                let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                aFields.forEach(function(field) {
                    field.SetDisplay(nType);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
           
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.GetDisplay();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).GetDisplay();
            else
                throw Error("InvalidGetError: Field is not a widget");
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
            if (typeof bValue == "boolean") {
                let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                aFields.forEach(function(field) {
                    field.SetDisplay(bValue ? display["hidden"] : display["visible"]);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.GetDisplay() == display["hidden"];
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).GetDisplay() == display["hidden"];
            else
                throw Error("InvalidGetError: Field is not a widget");
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
            if (Array.isArray(value)) {
                let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                let aColor  = private_correctApiColor(value).slice(1);
                aFields.forEach(function(field) {
                    field.SetBackgroundColor(aColor);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
            
        },
        get: function() {
            if (this.field.IsWidget())
                return private_getApiColor(this.field.GetBackgroundColor());
            else if (this.field.IsAllKidsWidgets())
                return private_getApiColor(this.field.GetKid(0).GetBackgroundColor());
            else
                throw Error("InvalidGetError: Field is not a widget");
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
            nValue = parseInt(nValue);
            if (Object.values(LINE_WIDTH).includes(nValue)) {
                let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                aFields.forEach(function(field) {
                    field.SetBorderWidth(nValue);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.GetBorderWidth();
            else
                throw Error("InvalidGetError: Field is not a widget");
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
                let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
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
            if (typeof(bValue) == "boolean") {
                let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                aFields.forEach(function(field) {
                    field.SetReadOnly(bValue);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.IsReadOnly();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).IsReadOnly();
            else
                throw Error("InvalidGetError: Field is not a widget");
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
        // set: function(aRect) {
        //     if (Array.isArray(aRect)) {
        //         let isValidRect = true;
        //         for (let i = 0; i < 4; i++) {
        //             if (typeof(aRect[i]) != "number") {
        //                 isValidRect = false;
        //                 break;
        //             }
        //         }
              
        //         if (isValidRect) {
                    
        //         }
        //     }
        // },
        get: function() {
            if (this.field.IsWidget())
                return this.field.GetOrigRect();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).GetOrigRect();
            else
                throw Error("InvalidGetError: Field is not a widget");
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
            if (typeof(bValue) == "boolean") {
                if (this.field.GetType() == AscPDF.FIELD_TYPES.button) {
                    throw Error("InvalidSetError: button field doesn't have 'required' prop.");
                }

                let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                aFields.forEach(function(field) {
                    if (field.GetType() != AscPDF.FIELD_TYPES.button)
                        field.SetRequired(bValue);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.GetType() != AscPDF.FIELD_TYPES.button) {
                if (this.field.IsWidget())
                    return this.field.IsRequired();
                else if (this.field.IsAllKidsWidgets())
                    return this.field.GetKid(0).IsRequired();
                else
                    throw Error("InvalidGetError: Field is not a widget");
            }
            else {
                throw Error("InvalidGetError: button field doesn't have 'required' prop.");
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
            if (Array.isArray(value)) {
                let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                let aColor  = private_correctApiColor(value).slice(1);
                aFields.forEach(function(field) {
                    field.SetBorderColor(aColor);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
            
        },
        get: function() {
            if (this.field.IsWidget())
                return private_getApiColor(this.field.GetBorderColor());
            else if (this.field.IsAllKidsWidgets())
                return private_getApiColor(this.field.GetKid(0).GetBorderColor());
            else
                throw Error("InvalidGetError: Field is not a widget");
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
            if (Array.isArray(aColor)) {
                let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                aFields.forEach(function(field) {
                    field.SetApiTextColor(aColor);
                });
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return private_getApiColor(this.field.GetTextColor());
            else if (this.field.IsAllKidsWidgets())
                return private_getApiColor(this.field.GetKid(0).GetTextColor());
            else
                throw Error("InvalidGetError: Field is not a widget");
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
            if (typeof(nValue) == "number" && nValue >= 0 && nValue < AscPDF.MAX_TEXT_SIZE) {
                let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                aFields.forEach(function(field) {
                    field.SetTextSize(Math.round(nValue));
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
                
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.GetTextSize();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).GetTextSize();
            else
                throw Error("InvalidGetError: Field is not a widget");
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
        let aFields = this.field._doc.GetAllWidgets(this.field.GetFullName());
        let nInternalType;
        switch (cTrigger) {
            case "MouseUp":
                nInternalType = AscPDF.FORMS_TRIGGERS_TYPES.MouseUp;
                break;
            case "MouseDown":
                nInternalType = AscPDF.FORMS_TRIGGERS_TYPES.MouseDown;
                break;
            case "MouseEnter":
                nInternalType = AscPDF.FORMS_TRIGGERS_TYPES.MouseEnter;
                break;
            case "MouseExit":
                nInternalType = AscPDF.FORMS_TRIGGERS_TYPES.MouseExit;
                break;
            case "OnFocus":
                nInternalType = AscPDF.FORMS_TRIGGERS_TYPES.OnFocus;
                break;
            case "OnBlur":
                nInternalType = AscPDF.FORMS_TRIGGERS_TYPES.OnBlur;
                break;
            case "Keystroke":
                nInternalType = AscPDF.FORMS_TRIGGERS_TYPES.Keystroke;
                break;
            case "Validate":
                nInternalType = AscPDF.FORMS_TRIGGERS_TYPES.Validate;
                break;
            case "Calculate":
                nInternalType = AscPDF.FORMS_TRIGGERS_TYPES.Calculate;
                break;
            case "Format":
                nInternalType = AscPDF.FORMS_TRIGGERS_TYPES.Format;
                break;
        }

        if (nInternalType != null) {
            aFields.forEach(function(field) {
                field.SetAction(nInternalType, cScript);
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
            if (typeof(nValue) == "number") {
                nValue = Math.round(nValue);
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    aFields.forEach(function(field) {
                        field.SetIconPosition(nValue, field.GetIconPosition().Y);
                    });
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid value.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, field is not a widget.");
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.GetIconPosition().X;
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).GetIconPosition().X;
            else
                throw Error("InvalidGetError: Field is not a widget");
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
            if (typeof(nValue) == "number") {
                nValue = Math.round(nValue);
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    aFields.forEach(function(field) {
                        field.SetIconPosition(field.GetIconPosition().X, nValue);
                    });
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid value.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, field is not a widget.");
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.GetIconPosition().Y;
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).GetIconPosition().Y;
            else
                throw Error("InvalidGetError: Field is not a widget");
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
        // set: function(arrValues) {
        //     if (Array.isArray(arrValues)) {
        //         if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
        //             let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());

        //             let aOpt = this.field.GetOptions();
        //             let apiValueToSet;
        //             let apiExpValue = this["value"];

        //             for (let i = 0; i < aFields.length; i++) {
        //                 let oField  = aFields[i];
        //                 let sExport = undefined;
        //                 if (arrValues[i] != undefined && arrValues[i] !== "") {
        //                     sExport = String(arrValues[i]);
        //                 }

        //                 if (sExport != undefined) {
        //                     oField.SetExportValue(arrValues[i]);
        //                     if (aOpt) {
        //                         aOpt[i] = arrValues[i];
        //                     }
        //                     if (oField.GetExportValue() == apiExpValue) {
        //                         apiValueToSet = oField.GetParentValue();
        //                     }
        //                 }
        //             }
        //             if (apiValueToSet != undefined)
        //                 this["value"] = apiValueToSet;
        //             else
        //                 this["value"] = "Off";
        //         }
        //         else {
        //             throw Error("InvalidSetError: Set not possible, field is not a widget.");
        //         }
        //     }
        //     else {
        //         throw Error("InvalidSetError: Set not possible, invalid value.");
        //     }
            
        // },
        get: function() {
            if (this.field.IsWidget())
                return [this.field.GetExportValue()];
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKids().map(function(field) { return field.GetExportValue()});
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
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
            if (Object.values(style).includes(sStyle)) {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    aFields.forEach(function(field) {
                        field.SetStyle(private_GetIntChStyle(sStyle));
                    })
                }
                else {
                    throw Error("InvalidSetError: Set not possible, field is not a widget.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
            
        },
        get: function() {
            if (this.field.IsWidget())
                return private_GetStrChStyle(this.field.GetStyle());
            else if (this.field.IsAllKidsWidgets())
                return private_GetStrChStyle(this.field.GetKid(0).GetStyle());
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
        }
	});

    /**
	 * Determines whether the specified widget is checked.
     * Note: For a set of radio buttons that do not have duplicate export values, you can get the value, which is equal to the
     * export value of the individual widget that is currently checked (or returns an empty string, if none is).
	 * @memberof ApiBaseCheckBoxField
     * @param {number} nWidget - The 0-based index of an individual radio button or check box widget for this field.
     * The index is determined by the order in which the individual widgets of this field
     * were created (and is unaffected by tab-order).
     * Every entry in the Fields panel has a suffix giving this index, for example, MyField #0.
	 * @typeofeditors ["PDF"]
     * @returns {string}
	 */
    ApiBaseCheckBoxField.prototype.isBoxChecked = function(nWidget) {
        if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
            let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
            let oField  = aFields[nWidget];
            if (!oField)
                throw Error("InvalidGetError: checkbox with this index doesn't exist");

            return oField.IsChecked();
        }
        else {
            throw Error("InvalidGetError: Field doesn't have widgets");
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
                if (value != undefined) {
                    value = String(value);
                    if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                        let oDoc            = this.field.GetDocument();
                        let oCalcInfo       = oDoc.GetCalculateInfo();
                        let oSourceField    = oCalcInfo.GetSourceField();

                        if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.field.GetFullName())
                            throw Error('InvalidSetError: Set not possible, invalid or unknown.');
                        if (oDoc.isOnValidate)
                            throw Error('InvalidSetError: Set not possible, invalid or unknown.');

                        let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                        aFields.forEach(function(field) {
                            field.SetValue(value);
                        });

                        this.field.SetParentValue(value);

                        if (oCalcInfo.IsInProgress() == false && oDoc.IsNeedDoCalculate()) {
                            oDoc.DoCalculateFields(this.field);
                            oDoc.CommitFields();
                        }
                    }
                    else {
                        this.field.SetParentValue(value);
                    }
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid value.");
                }
                
            },
            get: function() {
                return this.field.GetParentValue();
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
            if (typeof(bValue) == "boolean") {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    aFields.forEach(function(field) {
                        field.SetRadiosInUnison(bValue);
                    });
                    aFields[0].UpdateAll();
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.IsRadiosInUnison();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).IsRadiosInUnison();
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
        }
	});

    Object.defineProperties(ApiRadioButtonField.prototype, {
        "value": {
            set: function(sValue) {
                if (sValue != undefined) {
                    sValue = String(sValue);
                    if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                        let oDoc            = this.field.GetDocument();
                        let oCalcInfo       = oDoc.GetCalculateInfo();
                        let oSourceField    = oCalcInfo.GetSourceField();
    
                        if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.field.GetFullName())
                            throw Error('InvalidSetError: Set not possible, invalid or unknown.');
                        if (oDoc.isOnValidate)
                            throw Error('InvalidSetError: Set not possible, invalid or unknown.');;
    
                        
                        let sApiValueToSet = sValue;
                        let aOpt = this.field.GetOptions();
                        if (aOpt) {
                            let nIdx = aOpt.indexOf(sValue);
                            if (nIdx != -1)
                                sApiValueToSet = String(nIdx);
                        }

                        this.field.SetParentValue(sApiValueToSet);
                        this.field.GetKid(0).UpdateAll();
    
                        if (oCalcInfo.IsInProgress() == false && oDoc.IsNeedDoCalculate()) {
                            oDoc.DoCalculateFields(this.field);
                            oDoc.CommitFields();
                        }
                    }
                    else {
                        this.field.SetParentValue(sValue);
                    }
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid value.");
                }
            },
            get: function() {
                let aOpt = this.field.GetOptions();
                if (aOpt) {
                    return aOpt[this.field.GetParentValue()];
                }
                else {
                    return this.field.GetParentValue();
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
            if (Object.values(ALIGN_TYPE).includes(sValue)) {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    var nJcType = private_GetIntAlign(sValue);
                    aFields.forEach(function(field) {
                        field.SetAlign(nJcType);
                    });
                }
                else {
                    throw Error("InvalidSetError: Set not possible, field is not a widget.");
                }
            }
            else
                throw Error("InvalidSetError: Set not possible, invalid value.");
        },
        get: function() {
            if (this.field.IsWidget())
                return private_GetStrAlign(this.field.GetAlign());
            else if (this.field.IsAllKidsWidgets())
                return private_GetStrAlign(this.field.GetKid(0).GetAlign());
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
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
            let nIdx = parseInt(nValue);
            if (isNaN(parseInt(nIdx)) == false) {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    this.field.SetCalcOrderIndex(nIdx)
                }
                else {
                    throw Error("InvalidGetError: Get not possible, field is not a widget.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                return this.field.GetCalcOrderIndex();
            }
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
        }
	});

    /**
	 * Limits the number of characters that a user can type into a text field.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDF"]
	 */
    Object.defineProperty(ApiTextField.prototype, "charLimit", {
        set: function(nValue) {
            let nCharLimit = parseInt(nValue);
            if (isNaN(nCharLimit) == false) {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    aFields.forEach(function(field) {
                        field.SetCharLimit(nCharLimit);
                    });
                }
                else {
                    throw Error("InvalidGetError: Get not possible, field is not a widget.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
            
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.GetCharLimit();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).GetCharLimit();
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
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
        // set: function(bComb) {
        //     if (typeof bComb == "boolean") {
        //         if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
        //             let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
        //             aFields.forEach(function(field) {
        //                 field.SetComb(bComb);
        //             });
        //         }
        //         else {
        //             throw Error("InvalidGetError: Get not possible, field is not a widget.");
        //         }
        //     }
        //     else {
        //         throw Error("InvalidSetError: Set not possible, invalid value.");
        //     }
        // },
        get: function() {
            if (this.field.IsWidget())
                return this.field.IsComb();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).IsComb();
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
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
            if (typeof bDoNot == "boolean") {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    aFields.forEach(function(field) {
                        field.SetDoNotScroll(bDoNot);
                    });
                }
                else {
                    throw Error("InvalidGetError: Get not possible, field is not a widget.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.IsDoNotScroll();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).IsDoNotScroll();
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
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
            if (typeof bDoNot == "boolean") {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    aFields.forEach(function(field) {
                        field.SetDoNotSpellCheck(bDoNot);
                    });
                }
                else {
                    throw Error("InvalidGetError: Get not possible, field is not a widget.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.IsDoNotSpellCheck();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).IsDoNotSpellCheck();
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
        }
    });

    /**
	 * Controls how text is wrapped within the field. If false (the default), the text field can be a single line only. If true,
     * multiple lines are allowed and they wrap to field boundaries.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDF"]
	 */
    Object.defineProperty(ApiTextField.prototype, "multiline", {
        // set: function(bValue) {
        //     if (typeof bValue == "boolean") {
        //         if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
        //             let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
        //             aFields.forEach(function(field) {
        //                 field.SetMultiline(bValue);
        //             });
        //         }
        //         else {
        //             throw Error("InvalidGetError: Get not possible, field is not a widget.");
        //         }
        //     }
        //     else {
        //         throw Error("InvalidSetError: Set not possible, invalid value.");
        //     }
        // },
        get: function() {
            if (this.field.IsWidget())
                return this.field.IsMultiline();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).IsMultiline();
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
        }
    });

    Object.defineProperties(ApiTextField.prototype, {
        "value": {
            set: function(value) {
                if (value != undefined) {
                    value = String(value);
                    if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                        let oDoc            = this.field.GetDocument();
                        let oCalcInfo       = oDoc.GetCalculateInfo();
                        let oSourceField    = oCalcInfo.GetSourceField();

                        if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.field.GetFullName())
                            throw Error('InvalidSetError: Set not possible, invalid or unknown.');
                        if (oDoc.isOnValidate)
                            throw Error('InvalidSetError: Set not possible, invalid or unknown.');

                        if (value != null && value.toString)
                            value = value.toString();
                            
                        if (this["value"] == value)
                            return;

                        let oWidget = this.field.GetKid(0);
                        let isValid = oWidget.DoValidateAction(value);
                        if (isValid) {
                            oWidget.SetValue(value);

                            oWidget.needValidate = false; 
                            oWidget.Commit();
                            if (oCalcInfo.IsInProgress() == false) {
                                if (oDoc.event["rc"] !== false && oDoc.IsNeedDoCalculate()) {
                                    oDoc.DoCalculateFields(this.field);
                                    oDoc.AddFieldToCommit(oWidget);
                                    oDoc.CommitFields();
                                }
                            }
                        }
                    }
                    else {
                        this.field.SetParentValue(value);
                    }
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid value.");
                }
            },
            get: function() {
                let value = this.field.GetParentValue();
                let isNumber = /^\d+(\.\d+)?$/.test(value);
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
            if (typeof bValue == "boolean") {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    aFields.forEach(function(field) {
                        field.SetCommitOnSelChange(bValue);
                    });
                }
                else {
                    throw Error("InvalidGetError: Get not possible, field is not a widget.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.IsCommitOnSelChange();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).IsCommitOnSelChange();
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
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
        if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
            let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
            let oWidget = aFields[0];
            let aOptions = oWidget.GetOptions();
            if (aOptions[nIdx]) {
                if (bExportValue == false) {
                    return Array.isArray(aOptions[nIdx]) ? aOptions[nIdx][0] : aOptions[nIdx];
                }
                else {
                    return Array.isArray(aOptions[nIdx]) ? aOptions[nIdx][1] : aOptions[nIdx];
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        }
        else {
            throw Error("InvalidGetError: Get not possible, field is not a widget.");
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
            if (Object.values(ALIGN_TYPE).includes(sValue)) {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    var nJcType = private_GetIntAlign(sValue);
                    aFields.forEach(function(field) {
                        field.SetAlign(nJcType);
                    });
                }
                else {
                    throw Error("InvalidSetError: Set not possible, field is not a widget.");
                }
            }
            else
                throw Error("InvalidSetError: Set not possible, invalid value.");
        },
        get: function() {
            if (this.field.IsWidget())
                return private_GetStrAlign(this.field.GetAlign());
            else if (this.field.IsAllKidsWidgets())
                return private_GetStrAlign(this.field.GetKid(0).GetAlign());
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
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
            let nIdx = parseInt(nValue);
            if (isNaN(parseInt(nIdx)) == false) {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    this.field.SetCalcOrderIndex(nIdx)
                }
                else {
                    throw Error("InvalidGetError: Get not possible, field is not a widget.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                this.field.GetCalcOrderIndex();
            }
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
        }
	});

    /**
	 * Reads and writes value index of a combo box.
	 * @memberof ApiComboBoxField
	 * @typeofeditors ["PDF"]
	 */
    Object.defineProperty(ApiComboBoxField.prototype, "currentValueIndices", {
        set: function(nValue) {
            let nIdx = parseInt(nValue);
            if (isNaN(parseInt(nIdx)) == false) {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let oDoc            = this.field.GetDocument();
                    let oCalcInfo       = oDoc.GetCalculateInfo();
                    let oSourceField    = oCalcInfo.GetSourceField();

                    if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.field.GetFullName())
                        throw Error('InvalidSetError: Set not possible, invalid or unknown.');
                    if (oDoc.isOnValidate)
                        throw Error('InvalidSetError: Set not possible, invalid or unknown.');

                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    let oWidget = aFields[0];
                    if (oWidget.GetCurIdxs()[0] == nIdx)
                        return;

                    let sDisplayValue = this.getItemAt(nIdx, false);
                    let isValid = oWidget.DoValidateAction(sDisplayValue);

                    if (isValid) {
                        oWidget.SetCurIdxs([nIdx]);
                        oWidget.needValidate = false; 
                        oWidget.Commit();
                        if (oCalcInfo.IsInProgress() == false) {
                            if (oDoc.event["rc"] !== false && oDoc.IsNeedDoCalculate()) {
                                oDoc.DoCalculateFields(this.field);
                                oDoc.AddFieldToCommit(oWidget);
                                oDoc.CommitFields();
                            }
                        }
                    }
                }
                else {
                    throw Error("InvalidGetError: Get not possible, field is not a widget.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                let aCurIdxs = this.field.GetCurIdxs(true);
                return aCurIdxs[0] != undefined ? aCurIdxs[0] : -1;
            }
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
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
            if (typeof bValue == "boolean") {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    aFields.forEach(function(field) {
                        field.SetEditable(bValue);
                    });
                }
                else {
                    throw Error("InvalidGetError: Get not possible, field is not a widget.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.IsEditable();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).IsEditable();
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
        }
	});

    Object.defineProperties(ApiComboBoxField.prototype, {
        "value": {
            set: function(value) {
                if (value != undefined) {
                    value = String(value);
                    if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                        let oDoc = this.field.GetDocument();
                        let oCalcInfo = oDoc.GetCalculateInfo();
                        let oSourceField = oCalcInfo.GetSourceField();

                        if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.field.GetFullName())
                            throw Error('InvalidSetError: Set not possible, invalid or unknown.');
                        if (oDoc.isOnValidate)
                            throw Error('InvalidSetError: Set not possible, invalid or unknown.');

                        if (value != null && value.toString)
                            value = value.toString();
                            
                        if (this["value"] == value)
                            return;
                        
                        let oWidget = this.field.GetKid(0);
                        let isValid = oWidget.DoValidateAction(value);

                        if (isValid) {
                            oWidget.SetValue(value);
                            oWidget.needValidate = false; 
                            oWidget.Commit();
                            if (oCalcInfo.IsInProgress() == false) {
                                if (oDoc.event["rc"] !== false && oDoc.IsNeedDoCalculate()) {
                                    oDoc.DoCalculateFields(this.field);
                                    oDoc.AddFieldToCommit(oWidget);
                                    oDoc.CommitFields();
                                }
                            }
                        }
                    }
                    else {
                        this.field.SetParentValue(value);
                    }
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid value.");
                }
            },
            get: function() {
                let value = this.field.GetParentValue();
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
            if (typeof bDoNot == "boolean") {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    aFields.forEach(function(field) {
                        field.SetDoNotSpellCheck(bDoNot);
                    });
                }
                else {
                    throw Error("InvalidGetError: Get not possible, field is not a widget.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.IsDoNotSpellCheck();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).IsDoNotSpellCheck();
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
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
            if (Object.values(ALIGN_TYPE).includes(sValue)) {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    var nJcType = private_GetIntAlign(sValue);
                    aFields.forEach(function(field) {
                        field.SetAlign(nJcType);
                    });
                }
                else {
                    throw Error("InvalidSetError: Set not possible, field is not a widget.");
                }
            }
            else
                throw Error("InvalidSetError: Set not possible, invalid value.");
        },
        get: function() {
            if (this.field.IsWidget())
                return private_GetStrAlign(this.field.GetAlign());
            else if (this.field.IsAllKidsWidgets())
                return private_GetStrAlign(this.field.GetKid(0).GetAlign());
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
        }
	});

    /**
	 * Reads and writes single or multiply value index of a listbox.
	 * @memberof ApiListBoxField
	 * @typeofeditors ["PDF"]
	 */
    Object.defineProperty(ApiListBoxField.prototype, "currentValueIndices", {
        set: function(aIdxs) {
            if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                let oDoc            = this.field.GetDocument();
                let oCalcInfo       = oDoc.GetCalculateInfo();
                let oSourceField    = oCalcInfo.GetSourceField();
                let aFields         = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                let oWidget         = aFields[0];

                if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.field.GetFullName())
                    throw Error('InvalidSetError: Set not possible, invalid or unknown.');
                if (oDoc.isOnValidate)
                    throw Error('InvalidSetError: Set not possible, invalid or unknown.');
                if (Array.isArray(aIdxs) && oWidget.IsMultipleSelection() === false)
                    throw Error("InvalidSetError: Set not possible, invalid value.");
                
                if (Array.isArray(aIdxs))
                    oWidget.SetCurIdxs(aIdxs);
                else
                    oWidget.SetCurIdxs([aIdxs]);

                oWidget.Commit();
                if (oCalcInfo.IsInProgress() == false) {
                    if (oDoc.event["rc"] !== false && oDoc.IsNeedDoCalculate()) {
                        oDoc.DoCalculateFields(this.field);
                        oDoc.AddFieldToCommit(oWidget);
                        oDoc.CommitFields();
                    }
                }
            }
            else {
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
            }
        },
        get: function() {
            if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                let aFields     = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                let aCurIdxs    = this.field.GetCurIdxs(true);
                if (aFields[0].IsMultipleSelection() == false) {
                    return aCurIdxs[0] != undefined ? aCurIdxs[0] : -1;
                }
                else
                    return aCurIdxs.slice();
            }
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
        }
	});

    /**
	 * If true, indicates that a list box allows a multiple selection of items.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDF"]
	 */
    Object.defineProperty(ApiListBoxField.prototype, "multipleSelection", {
        set: function(bValue) {
            if (typeof bValue == "boolean") {
                if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                    let aFields = this.field.GetDocument().GetAllWidgets(this.field.GetFullName());
                    aFields.forEach(function(field) {
                        field.SetMultipleSelection(bValue);
                    });
                }
                else {
                    throw Error("InvalidGetError: Get not possible, field is not a widget.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid value.");
            }
        },
        get: function() {
            if (this.field.IsWidget())
                return this.field.IsMultipleSelection();
            else if (this.field.IsAllKidsWidgets())
                return this.field.GetKid(0).IsMultipleSelection();
            else
                throw Error("InvalidGetError: Get not possible, field is not a widget.");
        }
    });

    Object.defineProperties(ApiListBoxField.prototype, {
        "value": {
            set: function(value) {
                if (value != undefined) {
                    value = String(value);
                    if (this.field.IsWidget() || this.field.IsAllKidsWidgets()) {
                        let oDoc = this.field.GetDocument();
                        let oCalcInfo = oDoc.GetCalculateInfo();
                        let oSourceField = oCalcInfo.GetSourceField();

                        if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.field.GetFullName())
                            throw Error('InvalidSetError: Set not possible, invalid or unknown.');
                        if (oDoc.isOnValidate)
                            throw Error('InvalidSetError: Set not possible, invalid or unknown.');

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
                    }
                    else {
                        this.field.SetParentValue(value);
                    }
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid value.");
                }
                
            },
            get: function() {
                let value = this.field.GetParentValue();
                let isNumber = /^\d+(\.\d+)?$/.test(value);
                return isNumber ? parseFloat(value) : (value != undefined ? value : "");
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

/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

    // please use copy of this object
    let DEFAULT_SPAN = {
        "alignment":        ALIGN_TYPE.left,
        "fontFamily":       ["sans-serif"],
        "fontStretch":      "normal",
        "fontStyle":        "normal",
        "fontWeight":       400,
        "strikethrough":    false,
        "subscript":        false,
        "superscript":      false,
        "text":             "",
        "color":            AscPDF.Api.Objects.color["black"],
        "textSize":         12.0,
        "underline":        false
    }

    let highlight   = AscPDF.Api.Objects.highlight;
    let style       = AscPDF.Api.Objects.style;

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
        sName = sName != null && sName.toString ? sName.toString() : undefined;
        if (!sName)
            return null;

        let aPartNames = sName.split('.').filter(function(item) {
            if (item != "")
                return item;
        })

        let oRootField = this.doc.rootFields.get(aPartNames[0]);
        if (oRootField) {
            for (let i = 1; i < aPartNames.length; i++) {
                if (!oRootField)
                    return null;
                
                oRootField = oRootField.GetField(aPartNames[i]);
            }
    
            return oRootField.GetFormApi();
        }
        else {
            for (let i = 0; i < this.doc.widgets.length; i++) {
                if (this.doc.widgets[i].GetFullName() == sName)
                    return this.doc.widgets[i].GetFormApi();
            }
        }
            
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
            if (Object.values(AscPDF.BORDER_TYPES).includes(sValue)) {
                if (this.field.IsWidget()) {
                    let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());
                    aFields.forEach(function(field) {
                        field.SetBorderStyle(private_GetIntBorderStyle(sValue));
                    });
                }
                else {
                    this.field.GetKids().forEach(function(field) {
                        field.GetFormApi()["borderStyle"] = sValue;
                    });
                }
            }
        },
        get: function() {
            if (this.IsWidget())
                return private_GetStrBorderStyle(this.field.GetBorderStyle());
            else
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
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
        set: function(bValue) {
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                aFields.forEach(function(field) {
                    field.SetDefaultValue(bValue);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.GetDefaultValue();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
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
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                aFields.forEach(function(field) {
                    field.SetDisplay(nType);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.GetType();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
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
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                aFields.forEach(function(field) {
                    field.SetDisplay(window["AscPDF"].Api.Objects.display["hidden"]);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.GetDisplay() == window["AscPDF"].Api.Objects.display["hidden"];
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
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
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                let aColor = private_correctApiColor(value).slice(1);
                aFields.forEach(function(field) {
                    field.SetBackgroundColor(aColor);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return private_getApiColor(oField.GetBackgroundColor());
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
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
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                let aColor = private_correctApiColor(value).slice(1);
                aFields.forEach(function(field) {
                    field.SetBackgroundColor(aColor);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return private_getApiColor(oField.GetBackgroundColor());
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
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

    Object.defineProperties(ApiBaseField.prototype, {
        // private
        "parent": {
            enumerable: false,
            writable: true,
            value: null
        },
        "pagePos": {
            writable: true,
            enumerable: false,
            value: {
                x: 0,
                y: 0,
                w: 0,
                h: 0
            }
        },
        "kids": {
            enumerable: false,
            value: [],
        },
        "partialName": {
            writable: true,
            enumerable: false
        },

        
        "delay": {
            set: function(bValue) {
                if (typeof(bValue) == "boolean")
                    this._delay = bValue;
            },
            get: function() {
                return this._delay;
            }
        },
        "lineWidth": {
            set: function(nValue) {
                nValue = parseInt(nValue);
                if (Object.values(LINE_WIDTH).includes(nValue)) {
                    let aFields = this._doc.GetFields(this.name);
                    aFields.forEach(function(field) {
                        field._lineWidth = nValue;
                        field.SetNeedRecalc(true);
                        field.content.GetElement(0).Content.forEach(function(run) {
                            run.RecalcInfo.Measure = true;
                        });
                    });
                }
            },
            get: function() {
                return this._lineWidth;
            }
        },
        "borderWidth": {
            set: function(nValue) {
                this.lineWidth = nValue;
            },
            get: function() {
                return this.lineWidth;
            }
        },
        "name": {
            get: function() {
                return this.field.GetFullName();
            }
        },
        "page": {
            get: function() {
                return this.GetPage();
            }
        },
        "print": {
            set: function(bValue) {
                if (typeof(bValue) == "boolean")
                    this._print = bValue;
            },
            get: function() {
                return this._print;
            }
        },
        "readonly": {
            set: function(bValue) {
                if (typeof(bValue) == "boolean") {
                    let aFields = this._doc.GetFields(this.name);
                    aFields.forEach(function(field) {
                        field.SetReadOnly(bValue);
                    })
                }
                    
            },
            get: function() {
                return this._readonly;
            }
        },
        "rect": {
            set: function(aRect) {
                if (Array.isArray(aRect)) {
                    let isValidRect = true;
                    for (let i = 0; i < 4; i++) {
                        if (typeof(aRect[i]) != "number") {
                            isValidRect = false;
                            break;
                        }
                    }
                  
                    if (isValidRect)
                        this._rect = aRect;
                }
            },
            get: function() {
                return this._rect;
            }
        },
        "required": {
            set: function(bValue) {
                if (typeof(bValue) == "boolean" && this.GetType() != AscPDF.FIELD_TYPES.button) {
                    let aFields = this._doc.GetFields(this.name);

                    aFields.forEach(function(field) {
                        field.SetRequired(bValue);
                    })
                }
            },
            get: function() {
                if (this.GetType() != AscPDF.FIELD_TYPES.button)
                    return this._required;

                return undefined;
            }
        },
        "rotation": {
            set: function(nValue) {
                if (AscPDF.VALID_ROTATIONS.includes(nValue))
                    this._rotation = nValue;
            },
            get: function() {
                return this._rotation;
            }
        },
        "strokeColor": {
            set: function(aColor) {
                if (Array.isArray(aColor))
                    this._strokeColor = aColor;
            },
            get: function() {
                return this._strokeColor;
            }
        },
        "borderColor": {
            set: function(aColor) {
                if (Array.isArray(aColor))
                    this._borderColor = aColor;
            },
            get: function() {
                return this._borderColor;
            }
        },
        "submitName": {
            set: function(sValue) {
                if (typeof(sValue) == "string")
                    this._submitName = sValue;
            },
            get: function() {
                return this._submitName;
            }
        },
        "textColor": {
            set: function(aColor) {
                if (Array.isArray(aColor)) {
                    let aFields = this.field.GetDocument().GetFields(this.name);
                    aFields.forEach(function(field) {
                        field.SetApiTextColor(aColor);
                    });
                }
            },
            get: function() {
                return private_getApiColor(this.field.GetTextColor());
            }
        },
        "fgColor": {
            set: function(aColor) {
                if (Array.isArray(aColor))
                    this._fgColor = aColor;
            },
            get: function() {
                return this._fgColor;
            }
        },
        "textSize": {
            set: function(nValue) {
                if (typeof(nValue) == "number" && nValue >= 0 && nValue < AscPDF.MAX_TEXT_SIZE) {
                    let aFields = this._doc.GetFields(this.name);
                    let oField;
                    for (var i = 0; i < aFields.length; i++) {
                        oField = aFields[i];
                        oField._textSize = nValue;

                        let aParas = oField.content.Content;
                        aParas.forEach(function(para) {
                           para.SetApplyToAll(true);
                           para.Add(new AscCommonWord.ParaTextPr({FontSize : nValue}));
                           para.SetApplyToAll(false);
                           para.private_CompileParaPr(true);
                        });
                    }
                }
                    
            },
            get: function() {
                return 
            }
        },
        "userName": {
            set: function(sValue) {
                if (typeof(sValue) == "string")
                    this._userName = sValue;
            },
            get: function() {
                return this._userName;
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
        let aFields = this.field._doc.GetFields(this.name);
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
                let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

                if (aFields[0] && aFields[0].IsWidget()) {
                    aFields.forEach(function(field) {
                        field.SetIconPosition(nValue, field.GetIconPosition().Y);
                    });
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid or unknown.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.GetIconPosition().X;
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
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
            if (typeof(nValue) == "number") {
                nValue = Math.round(nValue);
                let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

                if (aFields[0] && aFields[0].IsWidget()) {
                    aFields.forEach(function(field) {
                        field.SetIconPosition(field.GetIconPosition().X, nValue);
                    });
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid or unknown.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.GetIconPosition().Y;
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
	});

    /**
	 * If true, the extent to which the icon may be scaled is set to the bounds of the button field. The additional icon
     * placement properties are still used to scale and position the icon within the button face.
	 * @memberof ApiPushButtonField
	 * @typeofeditors ["PDF"]
	 */
    Object.defineProperty(ApiPushButtonField.prototype, "buttonFitBounds", {
        set: function(bValue) {
            if (typeof(bValue) == "boolean") {
                let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

                if (aFields[0] && aFields[0].IsWidget()) {
                    aFields.forEach(function(field) {
                        field.SetButtonFitBounds(bValue);
                    });
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid or unknown.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.IsButtonFitBounds();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
	});

    /**
	 * Controls how the text and the icon of the button are positioned with respect to each other within the button face. The
     * convenience position object defines all of the valid alternatives.
	 * @memberof ApiPushButtonField
	 * @typeofeditors ["PDF"]
	 */
    Object.defineProperty(ApiPushButtonField.prototype, "buttonPosition", {
        set: function(bValue) {
            if (typeof(bValue) == "boolean") {
                let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

                if (aFields[0] && aFields[0].IsWidget()) {
                    aFields.forEach(function(field) {
                        field.SetHeaderPosition(bValue);
                    });
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid or unknown.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.GetHeaderPosition();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
	});

    /**
	 * Controls how the icon is scaled (if necessary) to fit inside the button face. he convenience scaleHow object defines all
     * of the valid alternatives:
     * Proportionally:      scaleHow.proportional
     * Non-proportionally:  scaleHow.anamorphic
     * @param {number}
	 * @memberof ApiPushButtonField
	 * @typeofeditors ["PDF"]
	 */
    Object.defineProperty(ApiPushButtonField.prototype, "buttonScaleHow", {
        set: function(nType) {
            if (typeof(nType) == "number") {
                nType = Math.round(nType);
                let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

                if (aFields[0] && aFields[0].IsWidget()) {
                    aFields.forEach(function(field) {
                        field.SetScaleHow(nType);
                    });
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid or unknown.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.GetScaleHow();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
	});

    /**
	 * Controls when an icon is scaled to fit inside the button face. The convenience scaleWhen object defines all of the valid 
     * alternatives:
     * Always:                  scaleWhen.always
     * Never:                   scaleWhen.never
     * If icon is too big:      scaleWhen.tooBig
     * If icon is too small:    scaleWhen.tooSmall
     * @param {number} - scaleHow.proportional or scaleHow.anamorphic
	 * @memberof ApiPushButtonField
	 * @typeofeditors ["PDF"]
	 */
    Object.defineProperty(ApiPushButtonField.prototype, "buttonScaleWhen", {
        set: function(nType) {
            if (typeof(nType) == "number") {
                nType = Math.round(nType);
                let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

                if (aFields[0] && aFields[0].IsWidget()) {
                    aFields.forEach(function(field) {
                        field.SetScaleWhen(nType);
                    });
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid or unknown.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.GetScaleWhen();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
	});

    /**
	 * Defines how a button reacts when a user clicks it. The four highlight modes supported are:
	 * none — No visual indication that the button has been clicked.
	 * invert — The region encompassing the button’s rectangle inverts momentarily.
	 * push — The down face for the button (if any) is displayed momentarily.
	 * outline — The border of the rectangle inverts momentarily.
	 * The convenience highlight object defines each state, as follows:
     * none - highlight.n
     * invert - highlight.i
     * push - highlight.p
     * outline - highlight.o
	 * @memberof ApiPushButtonField
	 * @typeofeditors ["PDF"]
	 */
    Object.defineProperty(ApiPushButtonField.prototype, "highlight", {
        set: function(sType) {
            if (typeof(sType) == "string" && highlight.includes(sType)) {
                sType = Math.round(sType);
                let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

                if (aFields[0] && aFields[0].IsWidget()) {
                    aFields.forEach(function(field) {
                        field.SetHighlight(private_GetIntHighlight(sType));
                    });
                }
                else {
                    throw Error("InvalidSetError: Set not possible, invalid or unknown.");
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return private_GetStrHighlight(oField.GetHighlight());
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
	});

    Object.defineProperties(ApiPushButtonField.prototype, {
        "textFont": {
            set: function(sValue) {
                if (typeof(sValue) == "string" && sValue !== "")
                    this._textFont = sValue;
            },
            get: function() {
                return this.textFont;
            }
        },
        "value": {
            get: function() {
                return "";
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
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                for (let i = 0; i < arrValues.length; i++) {
                    if (typeof(arrValues[i]) !== "string")
                        arrValues[i] = String(arrValues[i]);
                }

                for (let i = 0; i < aFields.length; i++) {
                    if (arrValues[i] != "" && arrValues[i] != undefined) {
                        aFields[i].SetExportValue(arrValues[i]);
                        if (aFields[i].GetExportValue() == this.field.GetApiValue())
                            aFields[i].SetChecked(true);
                        else
                            aFields[i].SetChecked(false);
                    }
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());
            if (aFields[0] && aFields[0].IsWidget()) {
                let aExpValues = [];
                for (let i = 0; i < aFields.length; i++) {
                    aExpValues.push(aFields[i].GetExportValue())
                }

                return aExpValues;
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
	});

    Object.defineProperties(ApiBaseCheckBoxField.prototype, {
        "style": {
            set: function(sStyle) {
                if (Object.values(style).includes(sStyle))
                    this._style = sStyle;
            },
            get: function() {
                return this._style;
            }
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
        let aFields = this.field._doc.GetFields(this.name);
        let oField = aFields[nWidget];
        if (!oField)
            return false;

        if (oField._exportValue == oField._value)
            return true;

        return false;
    };

    // for radiobutton
    const CheckedSymbol   = 0x25C9;
	const UncheckedSymbol = 0x25CB;

    function ApiCheckBoxField(oField)
    {
        ApiBaseCheckBoxField.call(this, oField);
    }
    ApiCheckBoxField.prototype = Object.create(ApiBaseCheckBoxField.prototype);
	ApiCheckBoxField.prototype.constructor = ApiCheckBoxField;
    Object.defineProperties(ApiCheckBoxField.prototype, {
        "value": {
            set: function(sValue) {
                let oDoc = this.field.GetDocument();
                let oCalcInfo = oDoc.GetCalculateInfo();
                let oSourceField = oCalcInfo.GetSourceField();

                if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.name)
                    throw Error('InvalidSetError: Set not possible, invalid or unknown.');;

                if (oDoc.isOnValidate)
                    return;

                let aFields = this.field.GetDocument().GetFields(this.name);
                if (this.exportValues.includes(sValue)) {
                    aFields.forEach(function(field) {
                        field._value = sValue;
                    });
                }
                else {
                    aFields.forEach(function(field) {
                        field._value = "Off";
                    });
                }

                if (oCalcInfo.IsInProgress() == false && oDoc.IsNeedDoCalculate()) {
                    oDoc.DoCalculateFields(this.field);
                    oDoc.CommitFields();
                }
            },
            get: function() {
                return this.field._value;
            }
        }
    });

    function ApiRadioButtonField(oField)
    {
        ApiBaseCheckBoxField.call(this, oField);
    }
    ApiRadioButtonField.prototype = Object.create(ApiBaseCheckBoxField.prototype);
	ApiRadioButtonField.prototype.constructor = ApiRadioButtonField;
    Object.defineProperties(ApiRadioButtonField.prototype, {
        "radiosInUnison": {
            set: function(bValue) {
                if (typeof(bValue) == "boolean") {
                    let aFields = this._doc.GetFields(this.name);
                    aFields.forEach(function(field) {
                        field.SetRadiosInUnison(bValue);
                    });
                }
            },
            get: function() {
                return this.IsRadiosInUnison();
            }
        },
        "value": {
            set: function(sValue) {
                let oDoc = this.field.GetDocument();
                let oCalcInfo = oDoc.GetCalculateInfo();
                let oSourceField = oCalcInfo.GetSourceField();

                if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.name)
                    throw Error('InvalidSetError: Set not possible, invalid or unknown.');
 
                if (oDoc.isOnValidate)
                    return;

                let aFields = this._doc.GetFields(this.name);
                if (this._exportValues.includes(sValue)) {
                    aFields.forEach(function(field) {
                        field._value = sValue;
                    });
                }
                else {
                    aFields.forEach(function(field) {
                        field._value = "Off";
                    });
                }

                if (oCalcInfo.IsInProgress() == false && oDoc.IsNeedDoCalculate()) {
                    oDoc.DoCalculateFields(this.field);
                    oDoc.CommitFields();
                }
            },
            get: function() {
                let aFields = this.field.GetDocument().GetFields(this.name);
                for (let i = 0; i < aFields.length; i++) {
                    if (aFields[i]._value != "Off" && aFields[i].IsNeedCommit()) {
                        return aFields[i]._value;
                    }
                }
                for (let i = 0; i < aFields.length; i++) {
                    if (aFields[i]._value != "Off") {
                        return aFields[i]._value;
                    }
                }
                return "Off";
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
            if (Object.values(ALIGN_TYPE).includes(sValue) == false)
                return;

            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                var nJcType = private_GetIntAlign(sValue);
                aFields.forEach(function(field) {
                    field.SetAlign(nJcType);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return private_GetStrAlign(oField.GetAlign());
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
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
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                aFields[0].SetCalcOrderIndex(nValue);
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.GetCalcOrderIndex();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
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
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                aFields.forEach(function(field) {
                    field.SetCharLimit(nValue);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.GetCharLimit();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
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
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                aFields.forEach(function(field) {
                    field.SetComb(bValue);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.IsComb();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
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
        set: function(bValue) {
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                aFields.forEach(function(field) {
                    field.SetDoNotScroll(bValue);
                });

                if (editor.getDocumentRenderer().activeForm == aFields[0]) {
                    if (bValue == true)
                        editor.getDocumentRenderer().activeForm.UpdateScroll(false, false);
                    else
                        editor.getDocumentRenderer().activeForm.UpdateScroll();
                }
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.IsDoNotScroll();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
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
        set: function(bValue) {
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                aFields.forEach(function(field) {
                    field.SetDoNotSpellCheck(bValue);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.IsDoNotSpellCheck();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
    });

    /**
	 * If true, sets the file-select flag in the Options tab of the text field (Field is Used for File Selection). This indicates that the
	 * value of the field represents a path of a file whose contents may be submitted with the form.
	 * The path may be entered directly into the field by the user, or the user can browse for the file. (See the
	 * browseForFileToSubmit method.)
	 * Note: The file select flag is mutually exclusive with the multiline, charLimit, password, and defaultValue
	 * properties. Also, on the Mac OS platform, when setting the file select flag, the field gets treated as read-only.
	 * Therefore, the user must browse for the file to enter into the field. (See browseForFileToSubmit.)
	 * This property can only be set during a batch or console event. See Privileged versus non-privileged context for
	 * details. The event object contains a discussion of JavaScript events.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDF"]
	 */
    Object.defineProperty(ApiTextField.prototype, "fileSelect", {
        set: function(bValue) {
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                aFields.forEach(function(field) {
                    field.SetFileSelect(bValue);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.IsFileSelect();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
    });

    Object.defineProperties(ApiTextField.prototype, {
        "multiline": {
            set: function(bValue) {
                if (typeof(bValue) != "boolean")
                    return;

                let aFields = this._doc.GetFields(this.name);
                aFields.forEach(function(field) {
                    field.SetMultiline(bValue);
                });
            },
            get: function() {
                return this._multiline;
            }
        },
        "password": {
            set: function(bValue) {
                if (typeof(bValue) != "boolean")
                    return;

                let aFields = this._doc.GetFields(this.name);
                aFields.forEach(function(field) {
                    field.SetPassword(bValue);
                });
            },
            get: function() {
                return this._password;
            }
        },
        "richText": {
            set: function(bValue) {
                if (typeof(bValue) == "boolean") {
                    let aFields = this._doc.GetFields(this.name);
                    aFields.forEach(function(field) {
                        field.SetRichText(bValue);
                    });
                }
            },
            get: function() {
                return this._richText;
            }
        },
        "richValue": {
            set: function(aSpans) {
                if (Array.isArray(aSpans)) {
                    let aCorrectVals = aSpans.filter(function(item) {
                        if (Array.isArray(item) == false && typeof(item) == "object" && item != null)
                            return item;
                    });

                    let aFields = this._doc.GetFields(this.name);
                    aFields.forEach(function(field) {
                        field._richValue = aCorrectVals;
                    });
                }
            },
            get: function() {
                return this._richValue;
            }
        },
        "textFont": {
            set: function(sValue) {
                if (typeof(sValue) == "string" && sValue !== "") {
                    let aFields = this._doc.GetFields(this.name);
                    aFields.forEach(function(field) {
                        field._textFont = sValue;
                    });
                }
            },
            get: function() {
                return this.textFont;
            }
        },
        "value": {
            set: function(value) {
                let oDoc = this.field.GetDocument();
                let oCalcInfo = oDoc.GetCalculateInfo();
                let oSourceField = oCalcInfo.GetSourceField();

                if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.name)
                    throw Error('InvalidSetError: Set not possible, invalid or unknown.');

                if (oDoc.isOnValidate)
                    return;

                if (value != null && value.toString)
                    value = value.toString();
                    
                if (this.value == value)
                    return;

                let isValid = this.field.DoValidateAction(value);

                if (isValid) {
                    this.field.SetValue(value);
                    if (this.field.IsWidget() == false)
                        return;

                    this.field.needValidate = false; 
                    this.field.Commit();
                    if (oCalcInfo.IsInProgress() == false) {
                        if (oDoc.event["rc"] == true && oDoc.IsNeedDoCalculate()) {
                            oDoc.DoCalculateFields(this.field);
                            oDoc.AddFieldToCommit(this.field);
                            oDoc.CommitFields();
                        }
                    }
                }
            },
            get: function() {
                let value = this.field.GetApiValue();
                let isNumber = /^\d+$/.test(value);
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
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                aFields.forEach(function(field) {
                    field.SetCommitOnSelChange(bValue);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.IsCommitOnSelChange();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
	});

    Object.defineProperties(ApiBaseListField.prototype, {
        "numItems": {
            get: function() {
                return this._options.length;
            }
        },
        "textFont": {
            set: function(sValue) {
                if (typeof(sValue) == "string" && sValue !== "")
                    this._textFont = sValue;
            },
            get: function() {
                return this.textFont;
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
        if (typeof(bExportValue) != "boolean")
            bExportValue = true;

        if (this.field._options[nIdx]) {
            if (typeof(this.field._options[nIdx]) == "string")
                return this.field._options[nIdx];
            else {
                if (bExportValue)
                    return this.field._options[nIdx][1];

                return this.field._options[nIdx][0];
            } 
        }
    };

    function ApiComboBoxField(oField)
    {
        ApiBaseListField.call(this, oField);
    };
    ApiComboBoxField.prototype = Object.create(ApiBaseListField.prototype);
	ApiComboBoxField.prototype.constructor = ApiComboBoxField;

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
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                aFields[0].SetCalcOrderIndex(nValue);
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.GetCalcOrderIndex();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
	});

    /**
	 * Reads and writes value index of a combo box.
	 * @memberof ApiComboBoxField
	 * @typeofeditors ["PDF"]
	 */
    Object.defineProperty(ApiComboBoxField.prototype, "currentValueIndices", {
        set: function(nValue) {
            if (typeof(nValue) !== "number" || this.getItemAt(nValue, false) == undefined)
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");

            let oDoc = this.field.GetDocument();
            let oCalcInfo = oDoc.GetCalculateInfo();
            let oSourceField = oCalcInfo.GetSourceField();
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.field.GetFullName() || aFields[0].IsWidget() == false)
                throw Error('InvalidSetError: Set not possible, invalid or unknown.');

            aFields[0].SelectOption(nValue);
            aFields[0].Commit();

            if (oCalcInfo.IsInProgress() == false && oDoc.IsNeedDoCalculate()) {
                oDoc.DoCalculateFields(this.field);
                oDoc.AddFieldToCommit(this.field);
                oDoc.CommitFields();
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.GetCurIdxs(true);
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
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
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                aFields[0].SetEditable(bValue);
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.IsEditable();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
	});

    Object.defineProperties(ApiComboBoxField.prototype, {
        "value": {
            set: function(value) {
                let oDoc = this.field.GetDocument();
                let oCalcInfo = oDoc.GetCalculateInfo();
                let oSourceField = oCalcInfo.GetSourceField();

                if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.name)
                    throw Error('InvalidSetError: Set not possible, invalid or unknown.');

                if (oDoc.isOnValidate)
                    return;

                if (value != null && value.toString)
                    value = value.toString();
                    
                if (this.value == value)
                    return;
                    
                let isValid = this.field.DoValidateAction(value);

                if (isValid) {
                    this.field.SetValue(value);
                    if (this.field.IsWidget() == false)
                        return;

                    this.field.needValidate = false; 
                    this.field.Commit();
                    if (oCalcInfo.IsInProgress() == false) {
                        if (oDoc.event["rc"] == true && oDoc.IsNeedDoCalculate()) {
                            oDoc.DoCalculateFields(this.field);
                            oDoc.AddFieldToCommit(this.field);
                            oDoc.CommitFields();
                        }
                    }
                }
            },
            get: function() {
                let value = this.field.GetApiValue();
                let isNumber = /^\d+$/.test(value);
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
        set: function(bValue) {
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());

            if (aFields[0] && aFields[0].IsWidget()) {
                aFields.forEach(function(field) {
                    field.SetDoNotSpellCheck(bValue);
                });
            }
            else {
                throw Error("InvalidSetError: Set not possible, invalid or unknown.");
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.IsDoNotSpellCheck();
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
    });

    /**
	 * Sets the list of items for a combo box.
	 * @memberof ApiComboBoxField
     * @param {string[]} values - An array in which each element is either an object convertible to a string or another array:
        For an element that can be converted to a string, the user and export values for the list item are equal to the string.
        For an element that is an array, the array must have two subelements convertible to strings, where the first is the user value and the second is the export value.
	 * @typeofeditors ["PDF"]
	 */
    ApiComboBoxField.prototype.setItems = function(values) {
        let aOptToPush = [];
        let oThis = this;

        for (let i = 0; i < values.length; i++) {
            if (values[i] == null)
                continue;
            if (typeof(values[i]) == "string" && values[i] != "")
                aOptToPush.push(values[i]);
            else if (Array.isArray(values[i]) && values[i][0] != undefined && values[i][1] != undefined) {
                if (values[i][0].toString && values[i][1].toString) {
                    aOptToPush.push([values[i][0].toString(), values[i][1].toString()])
                }
            }
            else if (typeof(values[i]) != "string" && values[i].toString) {
                aOptToPush.push(values[i].toString());
            }
        }

        let aFields = this._doc.GetFields(this.name);
        aFields.forEach(function(field) {
            field._options = aOptToPush.slice();
            if (field == oThis) {
                field.SelectOption(0, true);
                field.UnionLastHistoryPoints();
                field.SetNeedCommit(false);
            }
            else
                field.SelectOption(0, false);
        });
    };

    function ApiListBoxField(oField)
    {
        ApiBaseListField.call(this, oField);
    };
    
    ApiListBoxField.prototype = Object.create(ApiBaseListField.prototype);
	ApiListBoxField.prototype.constructor = ApiListBoxField;

    /**
	 * Reads and writes single or multiply value index of a listbox.
	 * @memberof ApiListBoxField
	 * @typeofeditors ["PDF"]
	 */
    Object.defineProperty(ApiListBoxField.prototype, "currentValueIndices", {
        set: function(value) {
            let oDoc = this.field.GetDocument();
            let oCalcInfo = oDoc.GetCalculateInfo();
            let oSourceField = oCalcInfo.GetSourceField();
            let aFields = this.field.GetDocument().GetFields(this.field.GetFullName());
            let curValues = this.field.GetCurIdxs(true);

            if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.field.GetFullName() || aFields[0].IsWidget() == false)
                throw Error('InvalidSetError: Set not possible, invalid or unknown.');

            if (Array.isArray(value) && this.multipleSelection === true) {
                let isValid = true;
                for (let i = 0; i < value.length; i++) {
                    if (typeof(value[i]) != "number" || this.getItemAt(value[i], false) === undefined) {
                        isValid = false;
                        break;
                    }
                }

                if (isValid == false)
                    throw Error('InvalidSetError: Set not possible, invalid or unknown.');

                // снимаем выделение с тех, которые не присутсвуют в новых значениях (value)
                for (let i = 0; i < curValues.length; i++) {
                    if (value.includes(curValues[i]) == false) {
                        this.UnselectOption(curValues[i]);
                    }
                }
                
                for (let i = 0; i < value.length; i++) {
                    // добавляем выделение тем, которые не присутсвуют в текущем поле
                    if (this.curValues.includes(value[i]) == false) {
                        this.SelectOption(value[i], false);
                    }
                }

                this.Commit();
            }
            else if (this.multipleSelection === false && typeof(value) === "number" && this.getItemAt(value, false) !== undefined) {
                this.SelectOption(value, true);
                this.Commit();
            }
            else
                return;

            aFields[0].Commit();
            if (oCalcInfo.IsInProgress() == false && oDoc.IsNeedDoCalculate()) {
                oDoc.DoCalculateFields(this.field);
                oDoc.AddFieldToCommit(this.field);
                oDoc.CommitFields();
            }
        },
        get: function() {
            let oField = this.field.GetDocument().GetField(this.field.GetFullName());
            if (oField && oField.IsWidget()) {
                return oField.GetCurIdxs(true);
            }
            else {
                throw Error("InvalidGetError: Get not possible, invalid or unknown.");
            }
        }
	});

    Object.defineProperties(ApiListBoxField.prototype, {
        "multipleSelection": {
            set: function(bValue) {
                if (typeof(bValue) == "boolean") {
                    if (bValue == this.multipleSelection)
                        return;

                    let aFields = this._doc.GetFields(this.name);
                    aFields.forEach(function(field) {
                        field.SetMultipleSelection(bValue);
                    });
                }
            },
            get: function() {
                return this.field.IsMultipleSelection();
            }
        },
        "value": {
            set: function(value) {
                let oDoc = this.field.GetDocument();
                let oCalcInfo = oDoc.GetCalculateInfo();
                let oSourceField = oCalcInfo.GetSourceField();

                if (oCalcInfo.IsInProgress() && oSourceField && oSourceField.GetFullName() == this.name)
                    throw Error('InvalidSetError: Set not possible, invalid or unknown.');;

                if (oDoc.isOnValidate)
                    return;

                if (value != null && value.toString)
                    value = value.toString();
                
                if (this.value == value)
                    return;
                
                this.field.SetValue(value);
                this.field.Commit();

                if (oCalcInfo.IsInProgress() == false && oDoc.IsNeedDoCalculate()) {
                    oDoc.DoCalculateFields(this.field);
                    oDoc.CommitFields();
                }
            },
            get: function() {
                let value = this.field.GetApiValue();
                let isNumber = /^\d+$/.test(value);
                return isNumber ? parseFloat(value) : (value != undefined ? value : "");
            }
        }
    });

    /**
	 * Sets the list of items for a list box.
	 * @memberof ApiListBoxField
     * @param {string[]} values - An array in which each element is either an object convertible to a string or another array:
        For an element that can be converted to a string, the user and export values for the list item are equal to the string.
        For an element that is an array, the array must have two subelements convertible to strings, where the first is the user value and the second is the export value.
	 * @typeofeditors ["PDF"]
	 */
    ApiListBoxField.prototype.setItems = function(values) {
        let aFields = this._doc.GetFields(this.name);

        aFields.forEach(function(field) {
            field._options = [];
            field.content.Internal_Content_RemoveAll();
            let sCaption, oPara, oRun;
            
            for (let i = 0; i < values.length; i++) {
                if (values[i] == null)
                    continue;
                sCaption = "";
                if (typeof(values[i]) == "string" && values[i] != "") {
                    sCaption = values[i];
                    field._options.push(values[i]);
                }
                else if (Array.isArray(values[i]) && values[i][0] != undefined && values[i][1] != undefined) {
                    if (values[i][0].toString && values[i][1].toString) {
                        field._options.push([values[i][0].toString(), values[i][1].toString()]);
                        sCaption = values[i][0].toString();
                    }
                }
                else if (typeof(values[i]) != "string" && values[i].toString) {
                    field._options.push(values[i].toString());
                    sCaption = values[i].toString();
                }

                if (sCaption != "") {
                    oPara = new AscCommonWord.Paragraph(field.content.DrawingDocument, field.content, false);
                    oRun = new AscCommonWord.ParaRun(oPara, false);
                    field.content.Internal_Content_Add(i, oPara);
                    oPara.Add(oRun);
                    oRun.AddText(sCaption);
                }
            }

            field.content.Recalculate_Page(0, true);
            field._curShiftView.x = 0;
            field._curShiftView.y = 0;
        });

        this.SelectOption(0, true);
        this.UnionLastHistoryPoints();

        if (this.field.IsMultipleSelection())
            this._currentValueIndices = [0];
        else
            this._currentValueIndices = 0;

        if (aFields.length > 1)
            this.Commit(this);
    };
    /**
	 * Inserts a new item into a list box
	 * @memberof ApiListBoxField
     * @param {string} cName - The item name that will appear in the form.
     * @param {string} cExport - (optional) The export value of the field when this item is selected. If not provided, the
     * cName is used as the export value.
     * @param {number} nIdx - (optional) The index in the list at which to insert the item. If 0 (the default), the new
     *  item is inserted at the top of the list. If –1, the new item is inserted at the end of the
     *  list.
	 * @typeofeditors ["PDF"]
	 */
    ApiListBoxField.prototype.insertItemAt = function(cName, cExport, nIdx) {
        let aFields = this.field._doc.GetFields(this.name);

        aFields.forEach(function(field) {
            field.InsertOption(cName, cExport, nIdx);
        })
    };

    function ApiSignatureField(oField)
    {
        ApiBaseField.call(this, oField);
    };

    function CSpan()
    {
        this._alignment = ALIGN_TYPE.left;
        this._fontFamily = ["sans-serif"];
        this._fontStretch = "normal";
        this._fontStyle = "normal";
        this._fontWeight = 400;
        this._strikethrough = false;
        this._subscript = false;
        this._superscript = false;

        Object.defineProperties(this, {
            "alignment": {
                set: function(sValue) {
                    if (Object.values(ALIGN_TYPE).includes(sValue))
                        this._alignment = sValue;
                },
                get: function() {
                    return this._alignment;
                }
            },
            "fontFamily": {
                set: function(arrValue) {
                    if (Array.isArray(arrValue))
                    {
                        let aCorrectFonts = [];

                        if (arrValue[0] !== undefined && typeof(arrValue[0]) == "string" && arrValue[0] === "")
                            aCorrectFonts.push(arrValue[0]);
                        if (arrValue[1] !== undefined && typeof(arrValue[1]) == "string" && arrValue[1] === "")
                            aCorrectFonts.push(arrValue[1]);

                        this._fontFamily = aCorrectFonts;
                    }
                }
            },
            "fontStretch": {
                set: function(sValue) {
                    if (AscPDF.FONT_STRETCH.includes(sValue))
                        this._fontStretch = sValue;
                },
                get: function() {
                    return this._fontStretch;
                }
            },
            "fontStyle": {
                set: function(sValue) {
                    if (Object.values(AscPDF.FONT_STYLE).includes(sValue))
                        this._fontStyle = sValue;
                },
                get: function() {
                    return this._fontStyle;
                }
            },
            "fontWeight": {
                set: function(nValue) {
                    if (AscPDF.FONT_WEIGHT.includes(nValue))
                        this._fontWeight = nValue;
                },
                get: function() {
                    return this._fontWeight;
                }
            },
            "strikethrough": {
                set: function(bValue) {
                    if (typeof(bValue) == "boolean")
                        this._strikethrough = bValue;
                },
                get: function() {
                    return this._strikethrough;
                }
            },
            "subscript": {
                set: function(bValue) {
                    if (typeof(bValue) == "boolean")
                        this._subscript = bValue;
                },
                get: function() {
                    return this._subscript;
                }
            },
            "superscript": {
                set: function(bValue) {
                    if (typeof(bValue) == "boolean")
                        this._superscript = bValue;
                },
                get: function() {
                    return this._superscript;
                }
            },

        });
    }

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

    function private_getApiColor(oInternalColor) {
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
    
    
    ApiComboBoxField.prototype["setItems"]              = ApiComboBoxField.prototype.setItems;
    

    ApiListBoxField.prototype["setItems"]               = ApiListBoxField.prototype.setItems;
    ApiListBoxField.prototype["insertItemAt"]           = ApiListBoxField.prototype.insertItemAt;

	window["AscPDF"].ApiDocument          = ApiDocument;
	window["AscPDF"].ApiTextField         = ApiTextField;
	window["AscPDF"].ApiPushButtonField   = ApiPushButtonField;
	window["AscPDF"].ApiCheckBoxField     = ApiCheckBoxField;
	window["AscPDF"].ApiRadioButtonField  = ApiRadioButtonField;
	window["AscPDF"].ApiComboBoxField     = ApiComboBoxField;
	window["AscPDF"].ApiListBoxField      = ApiListBoxField;
})();

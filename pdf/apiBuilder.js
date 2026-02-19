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
	 * @typedef {'HH:MM' | 'h:MM tt' | 'HH:MM:ss' | 'h:MM:ss tt'} TimeFormat
	 * Time format options:
	 * - "24HR_MM" — 24-hour format, hours and minutes (e.g., "14:30")
	 * - "12HR_MM" — 12-hour format with AM/PM, hours and minutes (e.g., "2:30 PM")
	 * - "24HR_MM_SS" — 24-hour format, hours, minutes, and seconds (e.g., "14:30:15")
	 * - "12HR_MM_SS" — 12-hour format with AM/PM, hours, minutes, and seconds (e.g., "2:30:15 PM")
	 */

	/**
	 * The available annotation border styles.
	 * @typedef {("solid" | "dashed")} AnnotBorderStyle
	 */

	/**
	 * Axis-aligned rectangle represented as a tuple.
	 *
	 * Invariants:
	 *  - rect[0] < rect[2] (x1 < x2)
	 *  - rect[1] < rect[3] (y1 < y2)
	 *
	 * @typedef {[pt, pt, pt, pt]} Rect
	 * @property {pt} 0 - x1 (left)
	 * @property {pt} 1 - y1 (top)
	 * @property {pt} 2 - x2 (right)
	 * @property {pt} 3 - y2 (bottom)
	 */

	/**
	 * Axis-aligned rectangle difference represented as a tuple.
	 * Describes coordinate-wise deltas between two rectangles (B - A).
	 *
	 * Invariants:
	 *  - diff[0] = x1B - x1A
	 *  - diff[1] = y1B - y1A
	 *  - diff[2] = x2B - x2A
	 *  - diff[3] = y2B - y2A
	 *
	 * @typedef {[pt, pt, pt, pt]} RectDiff
	 * @property {pt} 0 - dx1 (left delta)
	 * @property {pt} 1 - dy1 (top delta)
	 * @property {pt} 2 - dx2 (right delta)
	 * @property {pt} 3 - dy2 (bottom delta)
	 */

	/**
	 * The available display types.
	 * @typedef {("visible" | "hidden" | "noPrint" | "noView")} DisplayType
	 */

	/**
	 * The available text annot icon types.
	 * @typedef {("check" | "circle" | "comment" | "cross" | "crossH" | "help" | "insert" | "key" | "newParagraph" | "note" | "paragraph" | "rightArrow" | "rightPointer" | "star" | "upArrow" | "upLeftArrow")} TextIconType
	 */

	/**
	 * The available annotation border effect style.
	 * @typedef {("none" | "cloud")} AnnotBorderEffectStyle
	 */

	/**
	 * Axis-aligned point.
	 * @typedef {object} Point
	 * @property {pt} x
	 * @property {pt} y
	 */

	/**
	 * The available line end styles.
	 * @typedef {("square" | "circle" | "diamond" | "openArrow" | "closedArrow" | "none" | "butt" | "rOpenArrow" | "rClosedArrow" | "slash")} LineEndStyle
	 */

	/**
	 * An array of points representing a continuous path.
	 * @typedef {Array<Point>} Path
	 */

	/**
	 * An array of InkPath paths.
	 * @typedef {Array<Path>} PathList
	 */

	/**
	 * The available stamp types.
	 * @typedef {("D_Approved" | "D_Revised" | "D_Reviewed" | "D_Received" | "SB_Approved" | "SB_NotApproved" | "SB_Revised" | "SB_Confidential" | "SB_ForComment" | "SB_ForPublicRelease" | "SB_NotForPublicRelease" | "SB_PreliminaryResults" | "SB_InformationOnly" | "SB_Draft" | "SB_Completed" | "SB_Final" | "SB_Void" | "SH_SignHere" | "SH_Witness" | "SH_InitialHere" | "Expired")} StampType
	 */

	/**
	 * Quadrilateral represented as a flat tuple of vertices.
	 * Vertices order is fixed:
	 *  · left-top → right-top → left-bottom → right-bottom
	 *
	 * Invariants:
	 *  · x1 <= x2 (top edge goes left → right)
	 *  · x3 <= x4 (bottom edge goes left → right)
	 *  · y1 <= y3 (left edge goes top → bottom)
	 *  · y2 <= y4 (right edge goes top → bottom)
	 *
	 * @typedef {[pt, pt, pt, pt, pt, pt, pt, pt]} Quad
	 * @property {pt} 0 - x1 (left top)
	 * @property {pt} 1 - y1 (left top)
	 * @property {pt} 2 - x2 (right top)
	 * @property {pt} 3 - y2 (right top)
	 * @property {pt} 4 - x3 (left bottom)
	 * @property {pt} 5 - y3 (left bottom)
	 * @property {pt} 6 - x4 (right bottom)
	 * @property {pt} 7 - y4 (right bottom)
	 */


	/**
	 * The available free text annot intent.
	 * @typedef {("freeText" | "freeTextCallout")} FreeTextIntent
	 */

	/**
	 * FreeText callout coordinates.
	 * @typedef {[Point, Point, Point]} FreeTextCallout
	 */
	/**
	 * Degree defines an angle in degrees.
	 * Can be any finite number (positive or negative).
	 *
	 * @typedef {number} Degree
	 */

	/**
	 * @typedef {Object} SearchProps
	 * @property {string} text - The text to search for.
	 * @property {boolean} matchCase - Whether the search is case-sensitive.
	 * @property {boolean} wholeWords - Whether to match whole words only.
	 */

	/**
	 * The available horizontal text alignment.
	 * @typedef {("left" | "right" | "both" | "center")} HorTextAlign
	 */

	/**
	 * The available text vertical alignment (used to align text in a shape with a placement for text inside it).
	 * @typedef {("top" | "center" | "bottom")} VerticalTextAlign
	 */

	/**
	 * The available vertical text alignment.
	 * @typedef {("baseline" | "subscript" | "superscript")} TextVertAlign
	 */

	/**
	 * The reading order (left-to-right or right-to-left).
	 * @typedef {("ltr" | "rtl")} ReadingOrder
	 */

	/**
	 * Any valid drawing object.
	 * @typedef {(ApiShape | ApiImage | ApiTable | ApiChart )} Drawing
	 */

	/**
	 * Any valid page float object.
	 * @typedef {(ApiBaseField | ApiBaseAnnotation | Drawing )} FloatObject
	 */

	/**
	 * Twentieths of a point (equivalent to 1/1440th of an inch).
	 * @typedef {number} twips
	 */

	/**
	 * English measure unit. 1 mm = 36000 EMUs, 1 inch = 914400 EMUs.
	 * @typedef {number} EMU
	 */

	/**
	 * 1 millimetre equals 1/10th of a centimetre.
	 * @typedef {number} mm
	 */

	/**
	 * The available text direction inside a drawing content.
	 * @typedef {("lrtb" | "tbrl" | "btlr")} TextFlowDirection
	 */

	/**
	 * This type specifies the preset shape geometry that will be used for a shape.
	 * @typedef {("accentBorderCallout1" | "accentBorderCallout2" | "accentBorderCallout3" | "accentCallout1" | "accentCallout2" | "accentCallout3" | "actionButtonBackPrevious" | "actionButtonBeginning" | "actionButtonBlank" | "actionButtonDocument" | "actionButtonEnd" | "actionButtonForwardNext" | "actionButtonHelp" | "actionButtonHome" | "actionButtonInformation" | "actionButtonMovie" | "actionButtonReturn" | "actionButtonSound" | "arc" | "bentArrow" | "bentConnector2" | "bentConnector3" | "bentConnector4" | "bentConnector5" | "bentUpArrow" | "bevel" | "blockArc" | "borderCallout1" | "borderCallout2" | "borderCallout3" | "bracePair" | "bracketPair" | "callout1" | "callout2" | "callout3" | "can" | "chartPlus" | "chartStar" | "chartX" | "chevron" | "chord" | "circularArrow" | "cloud" | "cloudCallout" | "corner" | "cornerTabs" | "cube" | "curvedConnector2" | "curvedConnector3" | "curvedConnector4" | "curvedConnector5" | "curvedDownArrow" | "curvedLeftArrow" | "curvedRightArrow" | "curvedUpArrow" | "decagon" | "diagStripe" | "diamond" | "dodecagon" | "donut" | "doubleWave" | "downArrow" | "downArrowCallout" | "ellipse" | "ellipseRibbon" | "ellipseRibbon2" | "flowChartAlternateProcess" | "flowChartCollate" | "flowChartConnector" | "flowChartDecision" | "flowChartDelay" | "flowChartDisplay" | "flowChartDocument" | "flowChartExtract" | "flowChartInputOutput" | "flowChartInternalStorage" | "flowChartMagneticDisk" | "flowChartMagneticDrum" | "flowChartMagneticTape" | "flowChartManualInput" | "flowChartManualOperation" | "flowChartMerge" | "flowChartMultidocument" | "flowChartOfflineStorage" | "flowChartOffpageConnector" | "flowChartOnlineStorage" | "flowChartOr" | "flowChartPredefinedProcess" | "flowChartPreparation" | "flowChartProcess" | "flowChartPunchedCard" | "flowChartPunchedTape" | "flowChartSort" | "flowChartSummingJunction" | "flowChartTerminator" | "foldedCorner" | "frame" | "funnel" | "gear6" | "gear9" | "halfFrame" | "heart" | "heptagon" | "hexagon" | "homePlate" | "horizontalScroll" | "irregularSeal1" | "irregularSeal2" | "leftArrow" | "leftArrowCallout" | "leftBrace" | "leftBracket" | "leftCircularArrow" | "leftRightArrow" | "leftRightArrowCallout" | "leftRightCircularArrow" | "leftRightRibbon" | "leftRightUpArrow" | "leftUpArrow" | "lightningBolt" | "line" | "lineInv" | "mathDivide" | "mathEqual" | "mathMinus" | "mathMultiply" | "mathNotEqual" | "mathPlus" | "moon" | "nonIsoscelesTrapezoid" | "noSmoking" | "notchedRightArrow" | "octagon" | "parallelogram" | "pentagon" | "pie" | "pieWedge" | "plaque" | "plaqueTabs" | "plus" | "quadArrow" | "quadArrowCallout" | "rect" | "ribbon" | "ribbon2" | "rightArrow" | "rightArrowCallout" | "rightBrace" | "rightBracket" | "round1Rect" | "round2DiagRect" | "round2SameRect" | "roundRect" | "rtTriangle" | "smileyFace" | "snip1Rect" | "snip2DiagRect" | "snip2SameRect" | "snipRoundRect" | "squareTabs" | "star10" | "star12" | "star16" | "star24" | "star32" | "star4" | "star5" | "star6" | "star7" | "star8" | "straightConnector1" | "stripedRightArrow" | "sun" | "swooshArrow" | "teardrop" | "trapezoid" | "triangle" | "upArrowCallout" | "upDownArrow" | "upDownArrow" | "upDownArrowCallout" | "uturnArrow" | "verticalScroll" | "wave" | "wedgeEllipseCallout" | "wedgeRectCallout" | "wedgeRoundRectCallout")} ShapeType
	 */

	/**
	 * This type specifies the available chart types which can be used to create a new chart.
	 * @typedef {(
	 *     "bar" | "barStacked" | "barStackedPercent" | "bar3D" | "barStacked3D" | "barStackedPercent3D" | "barStackedPercent3DPerspective" |
	 *     "horizontalBar" | "horizontalBarStacked" | "horizontalBarStackedPercent" | "horizontalBar3D" | "horizontalBarStacked3D" | "horizontalBarStackedPercent3D" |
	 *     "lineNormal" | "lineStacked" | "lineStackedPercent" | "lineNormalMarker" | "lineStackedMarker" | "lineStackedPerMarker" | "line3D" |
	 *     "pie" | "pie3D" | "doughnut" |
	 *     "scatter" | "scatterLine" | "scatterLineMarker" | "scatterSmooth" | "scatterSmoothMarker" |
	 *     "stock" |
	 *     "area" | "areaStacked" | "areaStackedPercent" |
	 *     "comboCustom" | "comboBarLine" | "comboBarLineSecondary" |
	 *     "radar" | "radarMarker" | "radarFilled" |
	 *     "unknown"
	 * )} ChartType
	 */

	/**
	 * Standard numeric format.
	 * @typedef {("General" | "0" | "0.00" | "#,##0" | "#,##0.00" | "0%" | "0.00%" |
	 * "0.00E+00" | "# ?/?" | "# ??/??" | "m/d/yyyy" | "d-mmm-yy" | "d-mmm" | "mmm-yy" | "h:mm AM/PM" |
	 * "h:mm:ss AM/PM" | "h:mm" | "h:mm:ss" | "m/d/yyyy h:mm" | "#,##0_\);(#,##0)" | "#,##0_\);\[Red\]\(#,##0)" | 
	 * "#,##0.00_\);\(#,##0.00\)" | "#,##0.00_\);\[Red\]\(#,##0.00\)" | "mm:ss" | "[h]:mm:ss" | "mm:ss.0" | "##0.0E+0" | "@")} NumFormat
	 */

	//------------------------------------------------------------------------------------------------------------------
	//
	// Api
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Base class
	 * @global
	 * @class
	 * @name Api
	 */
	let Api = {};
	
	/**
	 * Creates a text field with the specified text field properties.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiDocument}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetDocument.js
	 */
	Api.GetDocument = function() {
		return new ApiDocument(private_GetLogicDocument());
	};

	/**
	 * Creates a text field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} aRect - widget rect
	 * @returns {ApiTextField}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/AddTextField.js
	 */
	Api.CreateTextField = function(aRect) {
		let oDoc = private_GetLogicDocument();
		let oField = oDoc.CreateTextField();
		oField.SetRect(aRect);

		return new ApiTextField(oField);
	};

	/**
	 * Creates a text date field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} aRect - widget rect
	 * @returns {ApiTextField}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/AddDateField.js
	 */
	Api.CreateDateField = function(aRect) {
		let oDoc = private_GetLogicDocument();
		let oField = oDoc.CreateTextField(true);
		oField.SetRect(aRect);

		return new ApiTextField(oField);
	};

	/**
	 * Creates a image field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} aRect - widget rect
	 * @returns {ApiButtonField}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/AddImageField.js
	 */
	Api.CreateImageField = function(aRect) {
		let oDoc = private_GetLogicDocument();
		let oField = oDoc.CreateButtonField(true);
		oField.SetRect(aRect);

		return new ApiButtonField(oField);
	};

	/**
	 * Creates a checkbox field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} aRect - widget rect
	 * @returns {ApiCheckboxField}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/AddImageField.js
	 */
	Api.CreateCheckboxField = function(aRect) {
		let oDoc = private_GetLogicDocument();
		let oField = oDoc.CreateCheckboxField();
		oField.SetRect(aRect);

		return new ApiCheckboxField(oField);
	};

	/**
	 * Creates a radiobutton field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} aRect - widget rect
	 * @returns {ApiRadiobuttonField}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/AddRadiobuttonField.js
	 */
	Api.CreateRadiobuttonField = function(aRect) {
		let oDoc = private_GetLogicDocument();
		let oField = oDoc.CreateRadiobuttonField();
		oField.SetRect(aRect);

		return new ApiRadiobuttonField(oField);
	};

	/**
	 * Creates a combobox field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} aRect - widget rect
	 * @returns {ApiComboboxField}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/AddComboboxField.js
	 */
	Api.CreateComboboxField = function(aRect) {
		let oDoc = private_GetLogicDocument();
		let oField = oDoc.CreateComboboxField();
		oField.SetRect(aRect);

		return new ApiComboboxField(oField);
	};

	/**
	 * Creates a listbox field.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} aRect - widget rect
	 * @returns {ApiListboxField}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/AddListboxField.js
	 */
	Api.CreateListboxField = function(aRect) {
		let oDoc = private_GetLogicDocument();
		let oField = oDoc.CreateListboxField();
		oField.SetRect(aRect);

		return new ApiListboxField(oField);
	};

	/**
	 * Creates text annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} rect - annotation rect.
	 * @returns {ApiTextAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateTextAnnot.js
	 */
	Api.CreateTextAnnot = function(rect) {
		let oDoc = private_GetLogicDocument();

		if (!private_IsValidRect(rect)) {
			AscBuilder.throwException("The rect parameter must be a valid rect");
		}

		let oProps = {
			rect:           rect,
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.Text,
			creationDate:   new Date().getTime(),
			modDate:        new Date().getTime(),
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);

		return new ApiTextAnnotation(oAnnot);
	};

	/**
	 * Creates circle annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} rect - annotation rect.
	 * @returns {ApiCircleAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateCircleAnnot.js
	 */
	Api.CreateCircleAnnot = function(rect) {
		let oDoc = private_GetLogicDocument();

		if (!private_IsValidRect(rect)) {
			AscBuilder.throwException("The rect parameter must be a valid rect");
		}

		let oProps = {
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.Circle,
			creationDate:   new Date().getTime(),
			modDate:        new Date().getTime(),
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);
		oAnnot.SetBorderWidth(1);
		oAnnot.SetBorderStyle(AscPDF.BORDER_TYPES.solid);
		oAnnot.SetBorderColor([0, 0, 0]);

		let oApiAnnot = new ApiCircleAnnotation(oAnnot);
		oApiAnnot.private_UpdateRect(rect);

		return oApiAnnot;
	};
	
	/**
	 * Creates square annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} rect - annotation rect.
	 * @returns {ApiSquareAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateSquareAnnot.js
	 */
	Api.CreateSquareAnnot = function(rect) {
		let oDoc = private_GetLogicDocument();

		if (!private_IsValidRect(rect)) {
			AscBuilder.throwException("The rect parameter must be a valid rect");
		}

		let oProps = {
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.Square,
			creationDate:   new Date().getTime(),
			modDate:        new Date().getTime(),
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);
		oAnnot.SetBorderWidth(1);
		oAnnot.SetBorderStyle(AscPDF.BORDER_TYPES.solid);
		oAnnot.SetBorderColor([0, 0, 0]);

		let oApiAnnot = new ApiSquareAnnotation(oAnnot);
		oApiAnnot.private_UpdateRect(rect);

		return oApiAnnot;
	};

	/**
	 * Creates freeText annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} rect - annotation rect.
	 * @returns {ApiFreeTextAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateFreeTextAnnot.js
	 */
	Api.CreateFreeTextAnnot = function(rect) {
		let oDoc = private_GetLogicDocument();

		if (!private_IsValidRect(rect)) {
			AscBuilder.throwException("The rect parameter must be a valid rect");
		}

		let oProps = {
			rect:           rect,
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.FreeText,
			creationDate:   new Date().getTime(),
			modDate:        new Date().getTime(),
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);
		oAnnot.SetBorderWidth(1);
		oAnnot.SetBorderStyle(AscPDF.BORDER_TYPES.solid);
		oAnnot.SetBorderColor([0, 0, 0]);

		return new ApiFreeTextAnnotation(oAnnot);
	};
	
	/**
	 * Creates line annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} rect - annotation rect.
	 * @param {Point} startPoint - start line point
	 * @param {Point} endPoint - end line point
	 * @returns {ApiLineAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateLineAnnot.js
	 */
	Api.CreateLineAnnot = function(rect, startPoint, endPoint) {
		let oDoc = private_GetLogicDocument();

		if (!private_IsValidRect(rect)) {
			AscBuilder.throwException("The rect parameter must be a valid rect");
		}

		private_CheckPoint(startPoint);
		private_CheckPoint(endPoint);

		let oProps = {
			rect:           rect,
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.Line,
			creationDate:   new Date().getTime(),
			modDate:        new Date().getTime(),
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);
		oAnnot.SetLinePoints([startPoint['x'], startPoint['y'], endPoint['x'], endPoint['y']]);
		oAnnot.SetBorderWidth(1);
		oAnnot.SetBorderStyle(AscPDF.BORDER_TYPES.solid);
		oAnnot.SetBorderColor([0, 0, 0]);

		return new ApiLineAnnotation(oAnnot);
	};

	/**
	 * Creates ink annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} rect - annotation rect.
	 * @param {PathList} pathList - ink path list
	 * @returns {ApiInkAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateInkAnnot.js
	 */
	Api.CreateInkAnnot = function(rect, inkPaths) {
		let oDoc = private_GetLogicDocument();

		inkPaths = AscBuilder.GetArrayParameter(inkPaths, []);
		if (inkPaths.length == 0)
			AscBuilder.throwException("The inkPaths parameter must be a non empty array");

		inkPaths.forEach(function(path) {
			path = AscBuilder.GetArrayParameter(path, []);
			if (path.length == 0)
				AscBuilder.throwException("The ink path parameter must be a non empty array");

			path.forEach(function(point) {
				private_CheckPoint(point);
			});
		});

		if (!private_IsValidRect(rect)) {
			AscBuilder.throwException("The rect parameter must be a valid rect");
		}

		let oProps = {
			rect:           rect,
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.Ink,
			creationDate:   new Date().getTime(),
			modDate:        new Date().getTime(),
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);

		oAnnot.SetBorderWidth(1);
		oAnnot.SetBorderStyle(AscPDF.BORDER_TYPES.solid);
		oAnnot.SetBorderColor([0, 0, 0]);
		oAnnot.SetInkPoints(inkPaths.map(function(path) {
			let flatPath = [];
			path.forEach(function(point) {
				flatPath.push(point["x"], point["y"]);
			});

			return flatPath;
		}));

		return new ApiInkAnnotation(oAnnot);
	};

	/**
	 * Creates polygon annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} rect - annotation rect.
	 * @param {Path} path - polygon path
	 * @returns {ApiPolygonAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreatePolygonAnnot.js
	 */
	Api.CreatePolygonAnnot = function(rect, path) {
		let oDoc = private_GetLogicDocument();

		path = AscBuilder.GetArrayParameter(path, []);
		if (path.length == 0)
			AscBuilder.throwException("The path parameter must be a non empty array");

		path.forEach(function(point) {
			private_CheckPoint(point);
		});

		if (!private_IsValidRect(rect)) {
			AscBuilder.throwException("The rect parameter must be a valid rect");
		}

		let aVertices = [];
		path.forEach(function(point) {
			aVertices.push(point["x"], point["y"]);
		});

		let oProps = {
			rect:           rect,
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.Polygon,
			creationDate:   new Date().getTime(),
			modDate:        new Date().getTime(),
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);

		oAnnot.SetBorderWidth(1);
		oAnnot.SetBorderStyle(AscPDF.BORDER_TYPES.solid);
		oAnnot.SetBorderColor([0, 0, 0]);
		oAnnot.SetVertices(aVertices);

		return new ApiPolygonAnnotation(oAnnot);
	};

	/**
	 * Creates polyline annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} rect - annotation rect.
	 * @param {Path} path - polyline path
	 * @returns {ApiPolyLineAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreatePolyLineAnnot.js
	 */
	Api.CreatePolyLineAnnot = function(rect, path) {
		let oDoc = private_GetLogicDocument();

		path = AscBuilder.GetArrayParameter(path, []);
		if (path.length == 0)
			AscBuilder.throwException("The path parameter must be a non empty array");

		path.forEach(function(point) {
			private_CheckPoint(point);
		});

		if (!private_IsValidRect(rect)) {
			AscBuilder.throwException("The rect parameter must be a valid rect");
		}

		let aVertices = [];
		path.forEach(function(point) {
			aVertices.push(point["x"], point["y"]);
		});

		let oProps = {
			rect:           rect,
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.PolyLine,
			creationDate:   new Date().getTime(),
			modDate:        new Date().getTime(),
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);

		oAnnot.SetBorderWidth(1);
		oAnnot.SetBorderStyle(AscPDF.BORDER_TYPES.solid);
		oAnnot.SetBorderColor([0, 0, 0]);
		oAnnot.SetVertices(aVertices);

		return new ApiPolyLineAnnotation(oAnnot);
	};

	/**
	 * Creates stamp annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} rect - annotation rect (only x1, y1 coordinates will be used, since the stamp dimensions are reserved).
	 * @param {StampType} type - stamp type
	 * @param {string} [author] - name of the author
	 * @param {number} [creationDate] - creation date (timeStamp)
	 * @returns {ApiStampAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateStampAnnot.js
	 */
	Api.CreateStampAnnot = function(rect, type, author, creationDate) {
		let oDoc = private_GetLogicDocument();

		if (!private_IsValidRect(rect, true)) {
			AscBuilder.throwException("The rect parameter must be a valid rect");
		}

		if (!Object.values(AscPDF.STAMP_TYPES).includes(type)) {
			AscBuilder.throwException("The type parameter must be one of available");
		}

		if (author != null) {
			author = AscBuilder.GetStringParameter(author, null);
			if (!author) {
				AscBuilder.throwException("The author parameter must be a non emptry string");
			}
		}
		else {
			author = Asc.editor.User.asc_getUserName();
		}

		if (creationDate != null) {
			creationDate = AscBuilder.GetNumberParameter(creationDate, null);
			if (!creationDate) {
				AscBuilder.throwException("The creationDate parameter must be a number");
			}
		}
		
		let oStampRender = oDoc.CreateStampRender(type, author, creationDate);
		let nExtX = oStampRender.Width * g_dKoef_mm_to_pt;
		let nExtY = oStampRender.Height * g_dKoef_mm_to_pt;
		let nLineW = oStampRender.m_oPen.Size * g_dKoef_mm_to_pt;

		let X1 = rect[0];
		let Y1 = rect[1];
		let X2 = X1 + nExtX;
		let Y2 = Y1 + nExtY;

		let oProps = {
			rect:			[X1, Y1, X2, Y2],
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.Stamp,
			creationDate:   creationDate ? new Date().getTime() : creationDate,
			modDate:        creationDate ? new Date().getTime() : creationDate,
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);

		oAnnot.SetIconType(type);
		oAnnot.SetBorderWidth(nLineW);
		oAnnot.SetBorderColor([0, 0, 0]);
		oAnnot.SetRenderStructure(oStampRender.m_aStack[0]);
		oAnnot.SetInRect([X1, Y2, X1, Y1, X2, Y1, X2, Y2]);
		let oXfrm = oAnnot.getXfrm();
		oXfrm.setRot(0);

		return new ApiStampAnnotation(oAnnot);
	};

	/**
	 * Creates highlight annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect | Quad[]} rect - region to apply highlight.
	 * @returns {ApiHighlightAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateHighlightAnnot.js
	 */
	Api.CreateHighlightAnnot = function(rect) {
		let oDoc = private_GetLogicDocument();
		rect = AscBuilder.GetArrayParameter(rect, []);

		if (!private_IsValidRect(rect) && !rect.find(function(quad) {return private_IsValidQuad(quad)})) {
			AscBuilder.throwException("The rect parameter must be a valid rect or quad");
		}

		let aQuads;
		let _rect;
		if (private_IsValidRect(rect)) {
			aQuads = [private_ConvertRectToQuad(rect)];
			_rect = rect;
		}
		else {
			let minX = Infinity, maxX = -Infinity;
			let minY = Infinity, maxY = -Infinity;

			for (let i = 0; i < rect.length; i++) {
				for (let j = 0; j < rect[i].length; j += 2) {
					let x = rect[i][j];
					let y = rect[i][j + 1];

					if (x < minX) minX = x;
					if (x > maxX) maxX = x;
					if (y < minY) minY = y;
					if (y > maxY) maxY = y;
				}
			}

			aQuads = rect;
			_rect = [minX, minY, maxX, maxY];
		}

		let oProps = {
			rect:           _rect,
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.Highlight,
			creationDate:   new Date().getTime(),
			modDate:        new Date().getTime(),
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);

		oAnnot.SetBorderColor([1, 0, 0]);
		oAnnot.SetQuads(aQuads);

		return new ApiHighlightAnnotation(oAnnot);
	};

	/**
	 * Creates strikeout annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect | Quad[]} rect - region to apply strikeout.
	 * @returns {ApiStrikeoutAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateStrikeoutAnnot.js
	 */
	Api.CreateStrikeoutAnnot = function(rect) {
		let oDoc = private_GetLogicDocument();
		rect = AscBuilder.GetArrayParameter(rect, []);

		if (!private_IsValidRect(rect) && !rect.find(function(quad) {return private_IsValidQuad(quad)})) {
			AscBuilder.throwException("The rect parameter must be a valid rect or quad");
		}

		let aQuads;
		let _rect;
		if (private_IsValidRect(rect)) {
			aQuads = [private_ConvertRectToQuad(rect)];
			_rect = rect;
		}
		else {
			let minX = Infinity, maxX = -Infinity;
			let minY = Infinity, maxY = -Infinity;

			for (let i = 0; i < rect.length; i++) {
				for (let j = 0; j < rect[i].length; j += 2) {
					let x = rect[i][j];
					let y = rect[i][j + 1];

					if (x < minX) minX = x;
					if (x > maxX) maxX = x;
					if (y < minY) minY = y;
					if (y > maxY) maxY = y;
				}
			}

			aQuads = rect;
			_rect = [minX, minY, maxX, maxY];
		}

		let oProps = {
			rect:           _rect,
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.Strikeout,
			creationDate:   new Date().getTime(),
			modDate:        new Date().getTime(),
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);

		oAnnot.SetBorderColor([1, 0, 0]);
		oAnnot.SetQuads(aQuads);

		return new ApiStrikeoutAnnotation(oAnnot);
	};

	/**
	 * Creates underline annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect | Quad[]} rect - region to apply underline.
	 * @returns {ApiUnderlineAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateUnderlineAnnot.js
	 */
	Api.CreateUnderlineAnnot = function(rect) {
		let oDoc = private_GetLogicDocument();
		rect = AscBuilder.GetArrayParameter(rect, []);

		if (!private_IsValidRect(rect) && !rect.find(function(quad) {return private_IsValidQuad(quad)})) {
			AscBuilder.throwException("The rect parameter must be a valid rect or quad");
		}

		let aQuads;
		let _rect;
		if (private_IsValidRect(rect)) {
			aQuads = [private_ConvertRectToQuad(rect)];
			_rect = rect;
		}
		else {
			let minX = Infinity, maxX = -Infinity;
			let minY = Infinity, maxY = -Infinity;

			for (let i = 0; i < rect.length; i++) {
				for (let j = 0; j < rect[i].length; j += 2) {
					let x = rect[i][j];
					let y = rect[i][j + 1];

					if (x < minX) minX = x;
					if (x > maxX) maxX = x;
					if (y < minY) minY = y;
					if (y > maxY) maxY = y;
				}
			}

			aQuads = rect;
			_rect = [minX, minY, maxX, maxY];
		}

		let oProps = {
			rect:           _rect,
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.Underline,
			creationDate:   new Date().getTime(),
			modDate:        new Date().getTime(),
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);

		oAnnot.SetBorderColor([1, 0, 0]);
		oAnnot.SetQuads(aQuads);

		return new ApiUnderlineAnnotation(oAnnot);
	};

	/**
	 * Creates caret annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect | Quad[]} rect - region to apply caret.
	 * @returns {ApiCaretAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateCaretAnnot.js
	 */
	Api.CreateCaretAnnot = function(rect) {
		let oDoc = private_GetLogicDocument();
		rect = AscBuilder.GetArrayParameter(rect, []);

		if (!private_IsValidRect(rect) && !rect.find(function(quad) {return private_IsValidQuad(quad)})) {
			AscBuilder.throwException("The rect parameter must be a valid rect or quad");
		}

		let aQuads;
		let _rect;
		if (private_IsValidRect(rect)) {
			aQuads = [private_ConvertRectToQuad(rect)];
			_rect = rect;
		}
		else {
			let minX = Infinity, maxX = -Infinity;
			let minY = Infinity, maxY = -Infinity;

			for (let i = 0; i < rect.length; i++) {
				for (let j = 0; j < rect[i].length; j += 2) {
					let x = rect[i][j];
					let y = rect[i][j + 1];

					if (x < minX) minX = x;
					if (x > maxX) maxX = x;
					if (y < minY) minY = y;
					if (y > maxY) maxY = y;
				}
			}

			aQuads = rect;
			_rect = [minX, minY, maxX, maxY];
		}

		let oProps = {
			rect:           _rect,
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.Caret,
			creationDate:   new Date().getTime(),
			modDate:        new Date().getTime(),
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);

		oAnnot.SetBorderColor([1, 0, 0]);
		oAnnot.SetQuads(aQuads);

		return new ApiCaretAnnotation(oAnnot);
	};

	/**
	 * Creates redact annotation.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {Rect | Quad[]} rect - region to apply redact.
	 * @returns {ApiRedactAnnotation}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateRedactAnnot.js
	 */
	Api.CreateRedactAnnot = function(rect) {
		let oDoc = private_GetLogicDocument();
		rect = AscBuilder.GetArrayParameter(rect, []);

		if (!private_IsValidRect(rect) && !rect.find(function(quad) {return private_IsValidQuad(quad)})) {
			AscBuilder.throwException("The rect parameter must be a valid rect or quad");
		}

		let aQuads;
		let _rect;
		if (private_IsValidRect(rect)) {
			aQuads = [private_ConvertRectToQuad(rect)];
			_rect = rect;
		}
		else {
			let minX = Infinity, maxX = -Infinity;
			let minY = Infinity, maxY = -Infinity;

			for (let i = 0; i < rect.length; i++) {
				for (let j = 0; j < rect[i].length; j += 2) {
					let x = rect[i][j];
					let y = rect[i][j + 1];

					if (x < minX) minX = x;
					if (x > maxX) maxX = x;
					if (y < minY) minY = y;
					if (y > maxY) maxY = y;
				}
			}

			aQuads = rect;
			_rect = [minX, minY, maxX, maxY];
		}

		let oProps = {
			rect:           _rect,
			name:           AscCommon.CreateGUID(),
			type:           AscPDF.ANNOTATIONS_TYPES.Redact,
			creationDate:   new Date().getTime(),
			modDate:        new Date().getTime(),
			hidden:         false
		}

		let oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);

		oAnnot.SetQuads(aQuads);
		oAnnot.SetFillColor([0, 0, 0]);
		oAnnot.SetBorderColor([1, 0, 0]);

		return new ApiRedactAnnotation(oAnnot);
	};

	/**
	 * Creates a new paragraph.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiParagraph}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateParagraph.js
	 */
	Api.CreateParagraph = function() {
		return this.private_CreateApiParagraph(new AscWord.Paragraph(null, true));
	};

	/**
	 * Creates a new rich paragraph.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiRichParagraph}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateRichParagraph.js
	 */
	Api.CreateRichParagraph = function() {
		return new ApiRichParagraph(new AscWord.Paragraph(private_GetLogicDocument(), true));
	};

	/**
	 * Creates a new rich run.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiRichRun}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateRichRun.js
	 */
	Api.CreateRichRun = function() {
		return new ApiRichRun(new ParaRun(null, false));
	};

	/**
	 * Creates the empty rich text properties.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiTextPr}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateRichTextPr.js
	 */
	Api.CreateRichTextPr = function() {
		return new ApiRichTextPr(null, new AscCommonWord.CTextPr());
	};

	/**
	 * Creates the empty rich paragraph properties.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiParaPr}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateRichParaPr.js
	 */
	Api.CreateRichParaPr = function() {
		return new ApiRichParaPr(null, new AscCommonWord.CParaPr());
	};

	/**
	 * Creates a shape with the parameters specified.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {ShapeType} [sType= "rect"] - The shape type which specifies the preset shape geometry.
	 * @param {EMU} [nWidth = 72] - The shape width in English measure units.
	 * @param {EMU} [nHeight = 72] - The shape height in English measure units.
	 * @param {ApiFill} [oFill = Api.CreateNoFill()] - The color or pattern used to fill the shape.
	 * @param {ApiStroke} [oStroke = Api.CreateStroke(0, Api.CreateNoFill())] - The stroke used to create the element shadow.
	 * @returns {ApiShape}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateShape.js
	 */
	Api.CreateShape = function(sType, nWidth, nHeight, oFill, oStroke) {
		let oDoc = private_GetLogicDocument();
		sType = sType || "rect";
		nWidth = nWidth || 72;
		nHeight = nHeight || 72;
		oFill = oFill || Asc.editor.CreateNoFill();
		oStroke = oStroke || Asc.editor.CreateStroke(0, Asc.editor.CreateNoFill());

		return new ApiShape(AscFormat.builder_CreateShape(sType, private_EMU2MM(nWidth), private_EMU2MM(nHeight), oFill.UniFill, oStroke.Ln, null, oDoc.GetTheme(), private_GetDrawingDocument(), false));
	};

	/**
	 * Creates an image with the parameters specified.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {string} sImageSrc - The image source where the image to be inserted should be taken from (currently,
	 * only internet URL or Base64 encoded images are supported).
	 * @param {EMU} nWidth - The image width in English measure units.
	 * @param {EMU} nHeight - The image height in English measure units.
	 * @returns {ApiImage}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateImage.js
	 */
	Api.CreateImage = function(sImageSrc, nWidth, nHeight) {
		let oImage = AscFormat.DrawingObjectsController.prototype.createImage(sImageSrc, 0, 0, private_EMU2MM(nWidth), private_EMU2MM(nHeight));

		return new ApiImage(oImage);
	};

	/**
	 * Creates a table.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param nCols - Number of columns.
	 * @param nRows - Number of rows.
	 * @returns {ApiTable}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateTable.js
	 */
	Api.CreateTable = function(nCols, nRows) {
		let oDoc = private_GetLogicDocument();
		let nPage = oDoc.GetCurPage();
		if (nPage == -1) {
			nPage = 0;
		}

		let oGraphicFrame = oDoc.private_Create_TableGraphicFrame(nCols, nRows, oDoc.DefaultTableStyleId, undefined, undefined, undefined, undefined, nPage);
		
		let content = oGraphicFrame.graphicObject.Content;
		for (let i = 0; i < content.length; ++i) {
			content[i].Set_Height(0, Asc.linerule_AtLeast);
		}

		return new ApiTable(oGraphicFrame);
	};

	/**
	 * Creates a chart with the parameters specified.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @param {ChartType} [sType="bar"] - The chart type used for the chart display.
	 * @param {number[][]} aSeries - The array of the data used to build the chart from.
	 * @param {number[] | string[]} aSeriesNames - The array of the names (the source table column names) used for the data which the chart will be build from.
	 * @param {number[] | string[]} aCatNames - The array of the names (the source table row names) used for the data which the chart will be build from.
	 * @param {EMU} nWidth - The chart width in English measure units.
	 * @param {EMU} nHeight - The chart height in English measure units.
	 * @param {number} nStyleIndex - The chart color style index (can be <b>1 - 48</b>, as described in OOXML specification).
	 * @param {NumFormat[] | String[]} aNumFormats - Numeric formats which will be applied to the series (can be custom formats).
	 * The default numeric format is "General".
	 * @returns {ApiChart}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateChart.js
	 */
	Api.CreateChart = function(sType, aSeries, aSeriesNames, aCatNames, nWidth, nHeight, nStyleIndex, aNumFormats) {
		let oChartSpace = AscFormat.builder_CreateChart(private_EMU2MM(nWidth), private_EMU2MM(nHeight), sType, aCatNames, aSeriesNames, aSeries, nStyleIndex, aNumFormats);
		return new ApiChart(oChartSpace);
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
	 * @returns {"document"}
	 * @see office-js-api/Examples/{Editor}/ApiDocument/Methods/GetClassType.js
	 */
	ApiDocument.prototype.GetClassType = function() {
		return "document";
	};

	/**
	 * Adds a new page to document.
	 * @memberof ApiDocument
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPos - pos to add page.
	 * @param {pt} [nWidth] - page width.
	 * @param {pt} [nHeight] - page height.
	 * @returns {ApiPage}
	 * @see office-js-api/Examples/{Editor}/ApiDocument/Methods/AddPage.js
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
	 * @memberof ApiDocument
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPos - page position
	 * @returns {ApiPage}
	 * @see office-js-api/Examples/{Editor}/ApiDocument/Methods/GetPage.js
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
	 * @memberof ApiDocument
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPos - page position
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiDocument/Methods/RemovePage.js
	 */
	ApiDocument.prototype.RemovePage = function(nPos) {
		let oFile = this.Document.GetFile();
		if (!oFile.pages[nPos]) {
			return false;
		}

		this.Document.RemovePage(nPos);
		return true;
	};

	/**
	 * Gets document pages count
	 * @memberof ApiDocument
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiDocument/Methods/GetPagesCount.js
	 */
	ApiDocument.prototype.GetPagesCount = function() {
		let oFile = this.Document.GetFile();
		return oFile.pages.length;
	};

	/**
	 * Gets list of all fields in document.
	 * @memberof ApiDocument
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiField}
	 * @see office-js-api/Examples/{Editor}/ApiDocument/Methods/GetAllFields.js
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
	 * @memberof ApiDocument
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiField}
	 * @see office-js-api/Examples/{Editor}/ApiDocument/Methods/GetFieldByName.js
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

	/**
	 * Searchs words and adds redact to it.
	 * @memberof ApiDocument
	 * @typeofeditors ["PDFE"]
	 * @param {SearchProps} props
	 * @returns {ApiRedactAnnotation[]}
	 * @see office-js-api/Examples/{Editor}/ApiDocument/Methods/SearchAndRedact.js
	 */
	ApiDocument.prototype.SearchAndRedact = function(props) {
		if (!props || typeof(props) !== "object" || Array.isArray(props)) {
			AscBuilder.throwException("The props parameter must be a SearchProps object");
		}

		let text = AscBuilder.GetStringParameter(props['text'], null);
		if (!text) {
			AscBuilder.throwException("The text property must be a valid string");
		}

		let matchCase = AscBuilder.GetBoolParameter(props['matchCase'], false);
		let wholeWords = AscBuilder.GetBoolParameter(props['wholeWords'], false);

		let searchSettings = new AscCommon.CSearchSettings();
		searchSettings.put_Text(text);
		searchSettings.put_MatchCase(matchCase);
		searchSettings.put_WholeWords(wholeWords);

		this.Document.Search(searchSettings);
		return this.Document.MarkAllSearchElementsForRedact().map(private_GetAnnotApi);
	};

	/**
	 * Applies added redact.
	 * @memberof ApiDocument
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiDocument/Methods/ApplyRedact.js
	 */
	ApiDocument.prototype.ApplyRedact = function() {
		let hasRedact = !!this.Document.annots.find(function(annot) {
			return annot.IsRedact() && !annot.GetRedactId();
		});

		if (!hasRedact) {
			AscBuilder.throwException("Has no redact to apply");
		}

		this.Document.ApplyRedact();
		return true;
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
	 * @see office-js-api/Examples/{Editor}/ApiPage/Methods/GetClassType.js
	 */
	ApiPage.prototype.GetClassType = function() {
		return "page";
	};

	/**
	 * Sets page rotation angle
	 * @memberof ApiPage
	 * @typeofeditors ["PDFE"]
	 * @param {number} nAngle
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiPage/Methods/SetRotation.js
	 */
	ApiPage.prototype.SetRotation = function(nAngle) {
		if (nAngle % 90 !== 0) {
			return false;
		}

		let oDoc = private_GetLogicDocument();
		oDoc.SetPageRotate(this.GetIndex(), nAngle);
		return true;
	};

	/**
	 * Gets page rotation angle
	 * @memberof ApiPage
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiPage/Methods/GetRotation.js
	 */
	ApiPage.prototype.GetRotation = function() {
		return this.Page.GetRotate();
	};

	/**
	 * Gets page index
	 * @memberof ApiPage
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiPage/Methods/GetIndex.js
	 */
	ApiPage.prototype.GetIndex = function() {
		return this.Page.GetIndex();
	};

	/**
	 * Gets page widgets
	 * @memberof ApiPage
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiPage/Methods/GetAllWidgets.js
	 */
	ApiPage.prototype.GetAllWidgets = function() {
		return this.Page.fields.map(private_GetWidgetApi);
	};

	/**
	 * Adds an available object to a page.
	 * @memberof ApiPage
	 * @typeofeditors ["PDFE"]
	 * @param {FloatObject} object
	 * @returns {FloatObject}
	 * @see office-js-api/Examples/{Editor}/ApiPage/Methods/AddObject.js
	 */
	ApiPage.prototype.AddObject = function(object) {
		if (!(object instanceof ApiBaseAnnotation) && !(object instanceof ApiBaseField) && !(object instanceof ApiDrawing)) {
			AscBuilder.throwException("The annot parameter must be a valid FloatObject");
		}

		let oInnerObj = object.private_GetImpl();
		if (oInnerObj.IsUseInDocument()) {
			AscBuilder.throwException("The object already in the document");
		}

		let oDoc = private_GetLogicDocument();
		if (object instanceof ApiBaseAnnotation) {
			oDoc.AddAnnot(oInnerObj, this.GetIndex());
		}
		else if (object instanceof ApiBaseField) {
			oDoc.AddField(oInnerObj, this.GetIndex());
		}
		else if (object instanceof ApiDrawing) {
			oDoc.AddDrawing(oInnerObj, this.GetIndex());
		}
		
		return object;
	};

	/**
	 * Gets all annots on page
	 * @memberof ApiPage
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiBaseAnnotation}
	 * @see office-js-api/Examples/{Editor}/ApiPage/Methods/GetAllAnnots.js
	 */
	ApiPage.prototype.GetAllAnnots = function() {
		let aAnnots = this.Page.GetAnnots();

		let aResult = [];
		for (let i = 0; i < aAnnots.length; i++) {
			if (aAnnots[i].IsRedact() && aAnnots[i].GetRedactId()) {
				continue;
			}

			aResult.push(private_GetAnnotApi(aAnnots[i]));
		}
		
		return aResult;
	};

	/**
	 * Search words and returns their quads.
	 * @memberof ApiPage
	 * @typeofeditors ["PDFE"]
	 * @param {SearchProps} props
	 * @returns {Quads[]}
	 * @see office-js-api/Examples/{Editor}/ApiPage/Methods/Search.js
	 */
	ApiPage.prototype.Search = function(props) {
		if (!props || typeof(props) !== "object" || Array.isArray(props)) {
			AscBuilder.throwException("The props parameter must be a SearchProps object");
		}

		let text = AscBuilder.GetStringParameter(props['text'], null);
		if (!text) {
			AscBuilder.throwException("The text property must be a valid string");
		}

		let matchCase = AscBuilder.GetBoolParameter(props['matchCase'], false);
		let wholeWords = AscBuilder.GetBoolParameter(props['wholeWords'], false);

		let searchSettings = new AscCommon.CSearchSettings();
		searchSettings.put_Text(text);
		searchSettings.put_MatchCase(matchCase);
		searchSettings.put_WholeWords(wholeWords);

		let nPageIdx = this.GetIndex();

		let oDoc = private_GetLogicDocument();
		let oSearchEngine = oDoc.Search(searchSettings);
		let aResult = [];

		Object.values(oSearchEngine.Elements).forEach(function(pdfMatch, idx) {
			if (pdfMatch.GetAbsolutePage() == nPageIdx) {
				let aPageSelQuads = oDoc.GetSearchElementSelectionQuads(idx);
				aResult = aResult.concat(aPageSelQuads[0].quads);
			}
		});
		
		return aResult;
	};

	/**
	 * Gets page selection quads
	 * @typeofeditors ["PDFE"]
	 * @param {Point} startPoint
	 * @param {Point} endPoint
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiPage/Methods/SetSelection.js
	 */
	ApiPage.prototype.SetSelection = function(startPoint, endPoint) {
		private_CheckPoint(startPoint);
		private_CheckPoint(endPoint);

		let oDoc = private_GetLogicDocument();
		let oFile = oDoc.GetFile();
		let nPageIdx = this.GetIndex();

		oDoc.BlurActiveObject();

		let startNearestPos = oFile.getNearestPos(nPageIdx, startPoint['x'], startPoint['y']);
		let endNearestPos = oFile.getNearestPos(nPageIdx, endPoint['x'], endPoint['y']);

		oFile.Selection.IsSelection = true;

		oFile.Selection.Page1  = nPageIdx;
		oFile.Selection.Line1  = startNearestPos.Line;
		oFile.Selection.Glyph1 = startNearestPos.Glyph;

		oFile.Selection.Page2  = nPageIdx;
		oFile.Selection.Line2  = endNearestPos.Line;
		oFile.Selection.Glyph2 = endNearestPos.Glyph;

		oDoc.Action.UpdateSelection = true;

		return true;
	};

	/**
	 * Gets page selection quads
	 * @typeofeditors ["PDFE"]
	 * @returns {Quad[]}
	 * @see office-js-api/Examples/{Editor}/ApiPage/Methods/GetSelectionQuads.js
	 */
	ApiPage.prototype.GetSelectionQuads = function() {
		let oDoc = private_GetLogicDocument();
		let nPageIdx = this.GetIndex();
		let aDocQuads = oDoc.GetFile().getSelectionQuads();

		let aPageQuads = [];
		for (let i = 0; i < aDocQuads.length; i++) {
			if (aDocQuads[i].page == nPageIdx) {
				aPageQuads = aDocQuads[i].quads;
				break;
			}
		}

		return aPageQuads;
	};

	/**
	 * Gets selected text on page
	 * @typeofeditors ["PDFE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiPage/Methods/GetSelectedText.js
	 */
	ApiPage.prototype.GetSelectedText = function() {
		return this.Page.GetSelectedText();
	};

	/**
	 * Recognizes content on the page and returns an array of recognized objects.
	 * @typeofeditors ["PDFE"]
	 * @returns {Drawing[]}
	 * @see office-js-api/Examples/{Editor}/ApiPage/Methods/RecognizeContent.js
	 */
	ApiPage.prototype.RecognizeContent = function() {
		let oDoc = private_GetLogicDocument();
		let nCurLength = this.Page.drawings.length;
		
		oDoc.EditPage(this.GetIndex());

		return AscBuilder.GetApiDrawings(this.Page.drawings.slice(nCurLength));
	};

	/**
	 * Gets all drawing objects from the page.
	 * @typeofeditors ["PDFE"]
	 * @returns {Drawing[]}
	 * @see office-js-api/Examples/{Editor}/ApiPage/Methods/GetAllDrawings.js
	 */
	ApiPage.prototype.GetAllDrawings = function() {
		return AscBuilder.GetApiDrawings(this.Page.drawings);
	};

	private_WrapClassMethods(ApiPage, function(method, args) {
		if (this.Page.GetIndex() == -1) {
			AscBuilder.throwException("You can't change deleted page");
		}
	});

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

	ApiBaseField.prototype.private_GetImpl = function() {
		return this.Field;
	};

	/**
	 * Sets new field name if possible.
	 * @typeofeditors ["PDFE"]
	 * @param {string} sName
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseField/Methods/SetFullName.js
	 */
	ApiBaseField.prototype.SetFullName = function(sName) {
		return this.Field.SetName(sName);
	};

	/**
	 * Gets field full name.
	 * @typeofeditors ["PDFE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiBaseField/Methods/GetFullName.js
	 */
	ApiBaseField.prototype.GetFullName = function() {
		return this.Field.GetFullName();
	};

	/**
	 * Sets new field partial name.
	 * @typeofeditors ["PDFE"]
	 * @param {string} sName
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseField/Methods/SetPartialName.js
	 */
	ApiBaseField.prototype.SetPartialName = function(sName) {
		return this.Field.SetPartialName(sName);
	};

	/**
	 * Gets field partial name.
	 * @typeofeditors ["PDFE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiBaseField/Methods/GetPartialName.js
	 */
	ApiBaseField.prototype.GetPartialName = function() {
		return this.Field.GetPartialName();
	};
	
	/**
	 * Sets field required
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseField/Methods/SetRequired.js
	 */
	ApiBaseField.prototype.SetRequired = function(bRequired) {
		this.Field.SetRequired(bRequired);
		return true;
	};

	/**
	 * Checks if field is required
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseField/Methods/IsRequired.js
	 */
	ApiBaseField.prototype.IsRequired = function() {
		return this.Field.IsRequired();
	};

	/**
	 * Sets field read only
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseField/Methods/SetReadOnly.js
	 */
	ApiBaseField.prototype.SetReadOnly = function(bReadOnly) {
		this.Field.SetReadOnly(bReadOnly);
		return true;
	};

	/**
	 * Checks if field is read only
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseField/Methods/IsReadOnly.js
	 */
	ApiBaseField.prototype.IsReadOnly = function() {
		return this.Field.IsReadOnly();
	};

	/**
	 * Sets field value
	 * @typeofeditors ["PDFE"]
	 * @param {string} sValue
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseField/Methods/SetValue.js
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseField/Methods/GetValue.js
	 */
	ApiBaseField.prototype.GetValue = function() {
		return this.Field.GetParentValue();
	};

	/**
	 * Adds new widget - visual representation for field
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPage - page to add widget
	 * @param {Rect} aRect - field rect
	 * @returns {?ApiWidget}
	 * @see office-js-api/Examples/{Editor}/ApiBaseField/Methods/AddWidget.js
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseField/Methods/GetAllWidgets.js
	 */
	ApiBaseField.prototype.GetAllWidgets = function() {
		return this.Field.GetAllWidgets().map(private_GetWidgetApi);
	};

	/**
	 * Removes field from document.
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseField/Methods/Delete.js
	 */
	ApiBaseField.prototype.Delete = function() {
		let oDoc = private_GetLogicDocument();
		
		this.Field.GetAllWidgets().forEach(function(widget) {
			oDoc.RemoveField(widget.GetId());
		});

		return true;
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/GetClassType.js
	 */
	ApiBaseWidget.prototype.GetClassType = function() {
		return "baseWidget";
	};

	/**
	 * Sets field rect.
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} rect
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/SetRect.js
	 */
	ApiBaseWidget.prototype.SetRect = function(rect) {
		if (!private_IsValidRect(rect)) {
			AscBuilder.throwException("The rect parameter must be a valid rect");
		}

		this.Field.SetRect(rect);
		return true;
	};

	/**
	 * Sets field rect.
	 * @typeofeditors ["PDFE"]
	 * @returns {Rect}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/GetRect.js
	 */
	ApiBaseWidget.prototype.GetRect = function() {
		return this.Field.GetRect();
	};

	/**
	 * Sets widget position.
	 * @typeofeditors ["PDFE"]
	 * @param {Point} position
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/SetPosition.js
	 */
	ApiBaseWidget.prototype.SetPosition = function(position) {
		private_CheckPoint(position);

		this.Field.SetPosition(position["x"], position["y"]);
		return true;
	};

	/**
	 * Gets widget position.
	 * @typeofeditors ["PDFE"]
	 * @returns {Point}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/GetPosition.js
	 */
	ApiBaseWidget.prototype.GetPosition = function() {
		let aRect = this.Field.GetRect();
		return {
			"x": aRect[0],
			"y": aRect[1]
		}
	};

	/**
	 * Sets widget border color.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiColor} color
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/SetBorderColor.js
	 */
	ApiBaseWidget.prototype.SetBorderColor = function(color) {
		if (!(color instanceof AscBuilder.ApiColor)) {
			return false;
		}

		this.Field.SetBorderColor(private_GetInnerColorByRGB(color["r"], color["g"], color["b"]));

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
	 * @returns {?ApiColor}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/GetBorderColor.js
	 */
	ApiBaseWidget.prototype.GetBorderColor = function() {
		let aInnerColor = this.Field.GetBorderColor();
		if (!aInnerColor) {
			return null;
		}

		let oRGB = this.Field.GetRGBColor(aInnerColor);

		return new Asc.editor.RGB(oRGB.r, oRGB.g, oRGB.b);
	};

	/**
	 * Sets widget border width.
	 * @typeofeditors ["PDFE"]
	 * @param {WidgetBorderWidth} sBorderWidth
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/SetBorderWidth.js
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/GetBorderWidth.js
	 */
	ApiBaseWidget.prototype.GetBorderWidth = function() {
		return private_GetStrBorderWidth(this.Field.GetBorderWidth());
	};

	/**
	 * Sets widget border style.
	 * @typeofeditors ["PDFE"]
	 * @param {WidgetBorderStyle} sBorderStyle
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/SetBorderStyle.js
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/GetBorderStyle.js
	 */
	ApiBaseWidget.prototype.GetBorderStyle = function() {
		return private_GetStrBorderStyle(this.Field.GetBorderStyle());
	};

	/**
	 * Sets widget background color.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiColor} color
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/SetBackgroundColor.js
	 */
	ApiBaseWidget.prototype.SetBackgroundColor = function(color) {
		if (!(color instanceof AscBuilder.ApiColor)) {
			return false;
		}

		this.Field.SetBackgroundColor(private_GetInnerColorByRGB(color["r"], color["g"], color["b"]));
		return true;
	};

	/**
	 * Gets widget background color.
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiColor}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/GetBackgroundColor.js
	 */
	ApiBaseWidget.prototype.GetBackgroundColor = function() {
		let aInnerColor = this.Field.GetBackgroundColor();
		if (!aInnerColor) {
			return null;
		}

		let oRGB = this.Field.GetRGBColor(aInnerColor);

		return new Asc.editor.RGB(oRGB.r, oRGB.g, oRGB.b);
	};

	/**
	 * Sets widget text color.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiColor} color
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/SetTextColor.js
	 */
	ApiBaseWidget.prototype.SetTextColor = function(color) {
		if (!(color instanceof AscBuilder.ApiColor)) {
			return false;
		}

		this.Field.SetTextColor(private_GetInnerColorByRGB(color["r"], color["g"], color["b"]));
		return true;
	};

	/**
	 * Gets widget text color.
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiColor}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/GetTextColor.js
	 */
	ApiBaseWidget.prototype.GetTextColor = function() {
		let aInnerColor = this.Field.GetTextColor();
		if (!aInnerColor) {
			return null;
		}

		let oRGB = this.Field.GetRGBColor(aInnerColor);

		return new Asc.editor.RGB(oRGB.r, oRGB.g, oRGB.b);
	};

	/**
	 * Sets widget text size.
	 * <note> Text size === 0 means autofit </note>
	 * @typeofeditors ["PDFE"]
	 * @param {pt} nSize
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/SetTextSize.js
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/GetTextSize.js
	 */
	ApiBaseWidget.prototype.GetTextSize = function() {
		return this.Field.GetTextSize();
	};

	/**
	 * Sets text autofit.
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bAuto
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/SetAutoFit.js
	 */
	ApiBaseWidget.prototype.SetAutoFit = function(bAuto) {
		return this.Field.SetTextSize(bAuto ? 0 : 11);
	};

	/**
	 * Checks if text is autofit.
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/IsAutoFit.js
	 */
	ApiBaseWidget.prototype.IsAutoFit = function() {
		return this.Field.GetTextSize() == 0;
	};

	/**
	 * Removes widget from parent field.
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseWidget/Methods/Delete.js
	 */
	ApiBaseWidget.prototype.Delete = function() {
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/GetClassType.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/SetMultiline.js
	 */
	ApiTextField.prototype.SetMultiline = function(bMultiline) {
		return this.Field.SetMultiline(bMultiline)
	};

	/**
	 * Checks if text field is multiline.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/IsMultiline.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/SetCharLimit.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/GetCharLimit.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/SetComb.js
	 */
	ApiTextField.prototype.SetComb = function(bComb) {
		return this.Field.SetComb(bComb)
	};

	/**
	 * Checks if text field is comb.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/IsComb.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/SetScrollLongText.js
	 */
	ApiTextField.prototype.SetScrollLongText = function(bScroll) {
		return this.Field.SetDoNotScroll(!bScroll)
	};

	/**
	 * Checks if text field can scroll long text.
	 * @memberof ApiTextField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/IsScrollLongText.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/SetNumberFormat.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/SetPercentageFormat.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/SetDateFormat.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/SetTimeFormat.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/SetSpecialFormat.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/SetMask.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/SetRegularExp.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/ClearFormat.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextField/Methods/SetValidateRange.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextWidget/Methods/GetClassType.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextWidget/Methods/SetPlaceholder.js
	 */
	ApiTextWidget.prototype.SetPlaceholder = function(sText) {
		return this.Field.SetPlaceholder(sText)
	};

	/**
	 * Gets text field placeholder.
	 * @memberof ApiTextWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiTextWidget/Methods/GetPlaceholder.js
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
	 * @see office-js-api/Examples/{Editor}/ApiTextWidget/Methods/SetRegularExp.js
	 */
	ApiTextWidget.prototype.SetRegularExp = function(sReg) {
		return this.Field.SetRegularExp(sReg)
	};

	/**
	 * Gets text widget regular validate expression.
	 * @memberof ApiTextWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTextWidget/Methods/GetRegularExp.js
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseListField/Methods/AddOption.js
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseListField/Methods/RemoveOption.js
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseListField/Methods/MoveOptionTo.js
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseListField/Methods/GetOption.js
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseListField/Methods/GetOptions.js
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseListField/Methods/SetCommitOnSelChange.js
	 */
	ApiBaseListField.prototype.SetCommitOnSelChange = function(bCommit) {
		return this.Field.SetCommitOnSelChange(bCommit)
	};

	/**
	 * Checks if field can commit on selection change.
	 * @memberof ApiBaseListField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseListField/Methods/IsCommitOnSelChange.js
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseListField/Methods/SetValueIndexes.js
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
	 * @see office-js-api/Examples/{Editor}/ApiBaseListField/Methods/GetValueIndexes.js
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
	 * @see office-js-api/Examples/{Editor}/ApiComboboxField/Methods/GetClassType.js
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
	 * @see office-js-api/Examples/{Editor}/ApiComboboxField/Methods/SetEditable.js
	 */
	ApiComboboxField.prototype.SetEditable = function(bCommit) {
		return this.Field.SetEditable(bCommit)
	};

	/**
	 * Checks if field is editable.
	 * @memberof ApiComboboxField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiComboboxField/Methods/IsEditable.js
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
	 * @see office-js-api/Examples/{Editor}/ApiComboboxField/Methods/SetNumberFormat.js
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
	 * @see office-js-api/Examples/{Editor}/ApiComboboxField/Methods/SetPercentageFormat.js
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
	 * @see office-js-api/Examples/{Editor}/ApiComboboxField/Methods/SetDateFormat.js
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
	 * @see office-js-api/Examples/{Editor}/ApiComboboxField/Methods/SetTimeFormat.js
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
	 * @see office-js-api/Examples/{Editor}/ApiComboboxField/Methods/SetSpecialFormat.js
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
	 * @see office-js-api/Examples/{Editor}/ApiComboboxField/Methods/SetMask.js
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
	 * @see office-js-api/Examples/{Editor}/ApiComboboxField/Methods/SetRegularExp.js
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
	 * @see office-js-api/Examples/{Editor}/ApiComboboxField/Methods/ClearFormat.js
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
	 * @see office-js-api/Examples/{Editor}/ApiComboboxField/Methods/SetValidateRange.js
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
	 * @see office-js-api/Examples/{Editor}/ApiListboxField/Methods/GetClassType.js
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
	 * @see office-js-api/Examples/{Editor}/ApiListboxField/Methods/SetMultipleSelection.js
	 */
	ApiListboxField.prototype.SetMultipleSelection = function(bMulti) {
		return this.Field.SetMultipleSelection(bMulti)
	};

	/**
	 * Checks if field is multiselect.
	 * @memberof ApiListboxField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiListboxField/Methods/IsMultipleSelection.js
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
	 * @see office-js-api/Examples/{Editor}/ApiCheckboxField/Methods/GetClassType.js
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
	 * @see office-js-api/Examples/{Editor}/ApiCheckboxField/Methods/SetToggleToOff.js
	 */
	ApiCheckboxField.prototype.SetToggleToOff = function(bToggle) {
		return this.Field.SetNoToggleToOff(!bToggle);
	};

	/**
	 * Checks if field is toggle to off.
	 * @memberof ApiCheckboxField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiCheckboxField/Methods/IsToggleToOff.js
	 */
	ApiCheckboxField.prototype.IsToggleToOff = function() {
		return !this.Field.IsNoToggleToOff();
	};

	/**
	 * Adds options to checkbox group.
	 * @memberof ApiCheckboxField
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPage - page to add option
	 * @param {Rect} - rect of new option
	 * @param {string} [sExportValue] - option checked value
	 * @returns {ApiCheckboxWidget}
	 * @see office-js-api/Examples/{Editor}/ApiCheckboxField/Methods/AddOption.js
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
	 * @see office-js-api/Examples/{Editor}/ApiRadiobuttonField/Methods/GetClassType.js
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
	 * @see office-js-api/Examples/{Editor}/ApiRadiobuttonField/Methods/SetCheckInUnison.js
	 */
	ApiRadiobuttonField.prototype.SetCheckInUnison = function(bInUnison) {
		return this.Field.SetRadiosInUnison(bInUnison);
	};

	/**
	 * Checks if field will check in unison.
	 * @memberof ApiRadiobuttonField
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRadiobuttonField/Methods/IsCheckInUnison.js
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
	 * @see office-js-api/Examples/{Editor}/ApiCheckboxWidget/Methods/GetClassType.js
	 */
	ApiCheckboxWidget.prototype.GetClassType = function() {
		return "checkboxWidget";
	};

	/**
	 * Sets checkbox widget checked.
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bChecked
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiCheckboxWidget/Methods/SetChecked.js
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
	 * @see office-js-api/Examples/{Editor}/ApiCheckboxWidget/Methods/IsChecked.js
	 */
	ApiCheckboxWidget.prototype.IsChecked = function() {
		return this.Field.IsChecked();
	};

	/**
	 * Sets widget checkbox style.
	 * @typeofeditors ["PDFE"]
	 * @param {CheckStyle} sStyle
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiCheckboxWidget/Methods/SetCheckStyle.js
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
	 * @see office-js-api/Examples/{Editor}/ApiCheckboxWidget/Methods/GetCheckStyle.js
	 */
	ApiCheckboxWidget.prototype.GetCheckStyle = function() {
		return private_GetStrCheckStyle(this.Field.GetStyle());
	};

	/**
	 * Sets widget export value.
	 * @typeofeditors ["PDFE"]
	 * @param {string} sValue
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiCheckboxWidget/Methods/SetExportValue.js
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
	 * @see office-js-api/Examples/{Editor}/ApiCheckboxWidget/Methods/GetExportValue.js
	 */
	ApiCheckboxWidget.prototype.GetExportValue = function() {
		return this.Field.GetExportValue();
	};

	/**
	 * Sets widget checked by default.
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bChecked
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiCheckboxWidget/Methods/SetCheckedByDefault.js
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
	 * @see office-js-api/Examples/{Editor}/ApiCheckboxWidget/Methods/IsCheckedByDefault.js
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
	 * @see office-js-api/Examples/{Editor}/ApiButtonField/Methods/GetClassType.js
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
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/GetClassType.js
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
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/SetLayout.js
	 */
	ApiButtonWidget.prototype.SetLayout = function(sType) {
		if (false == Object.keys(AscPDF.Api.Types.position).includes(sType)) {
			return false;
		}

		this.Field.SetLayout(AscPDF.Api.Types.position[sType]);
		return true;
	};

	/**
	 * Gets button widget layout type
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {ButtonLayout}
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/GetLayout.js
	 */
	ApiButtonWidget.prototype.GetLayout = function() {
		let nType = this.Field.GetLayout();
		return Object.keys(AscPDF.Api.Types.position).find(function(key) {
			return AscPDF.Api.Types.position[key] === nType;
		});
	};

	/**
	 * Sets button widget scale when type
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {ButtonScaleWhen} sType - button widget scale when type
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/SetScaleWhen.js
	 */
	ApiButtonWidget.prototype.SetScaleWhen = function(sType) {
		if (false == Object.keys(AscPDF.Api.Types.scaleWhen).includes(sType)) {
			return false;
		}

		this.Field.SetScaleWhen(AscPDF.Api.Types.scaleWhen[sType]);
		return true;
	};

	/**
	 * Gets button widget scale when type
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {ButtonScaleWhen}
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/GetScaleWhen.js
	 */
	ApiButtonWidget.prototype.GetScaleWhen = function() {
		let nType = this.Field.GetScaleWhen();
		return Object.keys(AscPDF.Api.Types.scaleWhen).find(function(key) {
			return AscPDF.Api.Types.scaleWhen[key] === nType;
		});
	};

	/**
	 * Sets button widget scale how type
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {ButtonScaleHow} sType - button widget scale how type
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/SetScaleHow.js
	 */
	ApiButtonWidget.prototype.SetScaleHow = function(sType) {
		if (false == Object.keys(AscPDF.Api.Types.scaleHow).includes(sType)) {
			return false;
		}

		this.Field.SetScaleHow(AscPDF.Api.Types.scaleHow[sType]);
		return true;
	};

	/**
	 * Gets button widget scale when type
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @returns {ButtonScaleHow}
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/GetScaleHow.js
	 */
	ApiButtonWidget.prototype.GetScaleHow = function() {
		let nType = this.Field.GetScaleHow();
		return Object.keys(AscPDF.Api.Types.scaleHow).find(function(key) {
			return AscPDF.Api.Types.scaleHow[key] === nType;
		});
	};

	/**
	 * Sets button widget fit bounds.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} bFit
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/SetFitBounds.js
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
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/IsFitBounds.js
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
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/SetIconXPos.js
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
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/GetIconXPos.js
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
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/SetIconYPos.js
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
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/GetIconYPos.js
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
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/SetBehavior.js
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
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/GetBehavior.js
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
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/SetLabel.js
	 */
	ApiButtonWidget.prototype.SetLabel = function(sLabel, sApType) {
		if (this.Field.GetLayout() == AscPDF.Api.Types.position["iconOnly"]) {
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
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/GetLabel.js
	 */
	ApiButtonWidget.prototype.GetLabel = function(sApType) {
		if (this.Field.GetLayout() == AscPDF.Api.Types.position["iconOnly"]) {
			return null;
		}

		if (undefined == sApType) {
			sApType = 'normal';
		}

		if (false == ['normal', 'down', 'hover'].includes(sApType)) {
			return null;
		}

		return this.Field.GetCaption(private_GetInnerButtonApType(sApType));
	};

	/**
	 * Sets image to button widget field.
	 * @memberof ApiButtonWidget
	 * @typeofeditors ["PDFE"]
	 * @param {string} [sImageUrl=''] - image url
	 * @param {ButtonAppearance} [sApType='normal'] - for what state is the picture set 
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiButtonWidget/Methods/SetImage.js
	 */
	ApiButtonWidget.prototype.SetImage = function(sImageUrl, sApType) {
		if (this.Field.GetLayout() == AscPDF.Api.Types.position["textOnly"]) {
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
	// ApiBaseAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a base annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 */
	function ApiBaseAnnotation(oAnnot) {
		this.Annot = oAnnot;
	}

	ApiBaseAnnotation.prototype.private_GetImpl = function() {
		return this.Annot;
	};

	ApiBaseAnnotation.prototype.private_UpdateRect = function(rect) {
		if (rect) {
			this.Annot.SetRect(rect);
		}
	};

	/**
	 * Sets annotation rect.
	 * @typeofeditors ["PDFE"]
	 * @param {Rect} rect
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetRect.js
	 */
	ApiBaseAnnotation.prototype.SetRect = function(rect) {
		if (!private_IsValidRect(rect)) {
			AscBuilder.throwException("The rect parameter must be a valid rect");
		}

		this.private_UpdateRect(rect);
		return true;
	};

	/**
	 * Sets annotation rect.
	 * @typeofeditors ["PDFE"]
	 * @returns {Rect}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetRect.js
	 */
	ApiBaseAnnotation.prototype.GetRect = function() {
		let aRD = this.Annot.GetRectangleDiff() || [0, 0, 0, 0];
		let aRect = this.Annot.GetRect();

		return [aRect[0] + aRD[0], aRect[1] + aRD[1], aRect[2] - aRD[2], aRect[3] - aRD[3]];
	};

	/**
	 * Sets annotation position.
	 * @typeofeditors ["PDFE"]
	 * @param {Point} position
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetPosition.js
	 */
	ApiBaseAnnotation.prototype.SetPosition = function(position) {
		private_CheckPoint(position);

		this.Annot.SetPosition(position["x"], position["y"]);
		return true;
	};

	/**
	 * Gets annotation position.
	 * @typeofeditors ["PDFE"]
	 * @returns {Point}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetPosition.js
	 */
	ApiBaseAnnotation.prototype.GetPosition = function() {
		let aRect = this.Annot.GetRect();
		return {
			"x": aRect[0],
			"y": aRect[1]
		}
	};

	/**
	 * Sets annotation border color.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiColor} color
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetBorderColor.js
	 */
	ApiBaseAnnotation.prototype.SetBorderColor = function(color) {
		if (!(color instanceof AscBuilder.ApiColor)) {
			return false;
		}

		this.Annot.SetBorderColor(private_GetInnerColorByRGB(color["r"], color["g"], color["b"]));
		return true;
	};

	/**
	 * Gets annotation border color.
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiColor}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetBorderColor.js
	 */
	ApiBaseAnnotation.prototype.GetBorderColor = function() {
		let aInnerColor = this.Annot.GetBorderColor();
		if (!aInnerColor) {
			return null;
		}

		let oRGB = this.Annot.GetRGBColor(aInnerColor);

		return new Asc.editor.RGB(oRGB.r, oRGB.g, oRGB.b);
	};

	/**
	 * Sets annotation fill color.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiColor} color
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetFillColor.js
	 */
	ApiBaseAnnotation.prototype.SetFillColor = function(color) {
		if (!(color instanceof AscBuilder.ApiColor)) {
			return false;
		}

		this.Annot.SetFillColor(private_GetInnerColorByRGB(color["r"], color["g"], color["b"]));
		return true;
	};

	/**
	 * Gets annotation fill color.
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiColor}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetFillColor.js
	 */
	ApiBaseAnnotation.prototype.GetFillColor = function() {
		let aInnerColor = this.Annot.GetFillColor();
		if (!aInnerColor) {
			return null;
		}

		let oRGB = this.Annot.GetRGBColor(aInnerColor);

		return new Asc.editor.RGB(oRGB.r, oRGB.g, oRGB.b);
	};

	/**
	 * Sets annotation border width.
	 * @typeofeditors ["PDFE"]
	 * @param {pt} width
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetBorderWidth.js
	 */
	ApiBaseAnnotation.prototype.SetBorderWidth = function(width) {
		width = AscBuilder.GetNumberParameter(width, 0);
		this.Annot.SetBorderWidth(width);
		return true;
	};

	/**
	 * Gets annotation border width.
	 * @typeofeditors ["PDFE"]
	 * @returns {pt}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetBorderWidth.js
	 */
	ApiBaseAnnotation.prototype.GetBorderWidth = function() {
		return this.Annot.GetBorderWidth();
	};

	/**
	 * Sets annotation border style.
	 * @typeofeditors ["PDFE"]
	 * @param {AnnotBorderStyle} borderStyle
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetBorderStyle.js
	 */
	ApiBaseAnnotation.prototype.SetBorderStyle = function(borderStyle) {
		if (borderStyle !== "solid" && borderStyle !== "dashed") {
			return false;
		}

		this.Annot.SetBorderStyle(private_GetInnerBorderStyle(borderStyle));
		return true;
	};

	/**
	 * Gets annotation border style.
	 * @typeofeditors ["PDFE"]
	 * @returns {AnnotBorderStyle}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetBorderStyle.js
	 */
	ApiBaseAnnotation.prototype.GetBorderStyle = function() {
		return private_GetStrBorderStyle(this.Annot.GetBorderStyle());
	};
	
	/**
	 * Sets annotation author name.
	 * @typeofeditors ["PDFE"]
	 * @param {string} name
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetAuthorName.js
	 */
	ApiBaseAnnotation.prototype.SetAuthorName = function(name) {
		name = AscBuilder.GetStringParameter(name, null);
		if (!name) {
			AscBuilder.throwException("The name parameter must be a non emptry string");
		}

		this.Annot.SetAuthor(name);
		return true;
	};

	/**
	 * Gets annotation author name.
	 * @typeofeditors ["PDFE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetAuthorName.js
	 */
	ApiBaseAnnotation.prototype.GetAuthorName = function() {
		return this.Annot.GetAuthor();
	};

	/**
	 * Sets annotation contents.
	 * @typeofeditors ["PDFE"]
	 * @param {?string} contents
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetContents.js
	 */
	ApiBaseAnnotation.prototype.SetContents = function(contents) {
		contents = AscBuilder.GetStringParameter(contents, null);
		this.Annot.SetContents(contents);
		return true;
	};

	/**
	 * Gets annotation contents.
	 * @typeofeditors ["PDFE"]
	 * @returns {?string}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetContents.js
	 */
	ApiBaseAnnotation.prototype.GetContents = function() {
		return this.Annot.GetContents();
	};

	/**
	 * Sets annotation creation date.
	 * @typeofeditors ["PDFE"]
	 * @param {number} timeStamp
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetCreationDate.js
	 */
	ApiBaseAnnotation.prototype.SetCreationDate = function(timeStamp) {
		timeStamp = AscBuilder.GetNumberParameter(timeStamp, null);
		if (!timeStamp) {
			AscBuilder.throwException("The timeStamp parameter must be number");
		}

		this.Annot.SetCreationDate(timeStamp);
		return true;
	};

	/**
	 * Gets annotation creation date.
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetCreationDate.js
	 */
	ApiBaseAnnotation.prototype.GetCreationDate = function() {
		return this.Annot.GetCreationDate();
	};

	/**
	 * Sets annotation last modification date.
	 * @typeofeditors ["PDFE"]
	 * @param {number} timeStamp
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetModDate.js
	 */
	ApiBaseAnnotation.prototype.SetModDate = function(timeStamp) {
		timeStamp = AscBuilder.GetNumberParameter(timeStamp, null);
		if (!timeStamp) {
			AscBuilder.throwException("The timeStamp parameter must be number");
		}

		this.Annot.SetModDate(timeStamp);
		return true;
	};

	/**
	 * Gets annotation last modification date.
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetModDate.js
	 */
	ApiBaseAnnotation.prototype.GetModDate = function() {
		return this.Annot.GetModDate();
	};

	/**
	 * Sets annotation unique name.
	 * @typeofeditors ["PDFE"]
	 * @param {string} name
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetUniqueName.js
	 */
	ApiBaseAnnotation.prototype.SetUniqueName = function(name) {
		name = AscBuilder.GetStringParameter(name, null);
		if (!name) {
			AscBuilder.throwException("The name parameter must be a non empty string");
		}

		if (Object.values(AscCommon.g_oTableId.m_aPairs).find(function(obj) {return obj.IsAnnot && obj.IsAnnot() && obj.GetName() == name})) {
			AscBuilder.throwException("This unique name is busy");
		}

		this.Annot.SetName(name);
		return true;
	};

	/**
	 * Gets annotation unique name.
	 * @typeofeditors ["PDFE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetUniqueName.js
	 */
	ApiBaseAnnotation.prototype.GetUniqueName = function() {
		return this.Annot.GetName();
	};

	/**
	 * Sets annotation opacity.
	 * @typeofeditors ["PDFE"]
	 * @param {percentage} name
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetOpacity.js
	 */
	ApiBaseAnnotation.prototype.SetOpacity = function(value) {
		value = AscBuilder.GetNumberParameter(value, null);
		if (null == value || value < 0 || value > 100) {
			AscBuilder.throwException("The value parameter must be number from 0 to 100");
		}

		this.Annot.SetOpacity(value / 100);
		return true;
	};

	/**
	 * Gets annotation opacity.
	 * @typeofeditors ["PDFE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetOpacity.js
	 */
	ApiBaseAnnotation.prototype.GetOpacity = function() {
		return this.Annot.GetOpacity() * 100;
	};

	/**
	 * Sets annotation subject.
	 * @typeofeditors ["PDFE"]
	 * @param {?string} subject
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetSubject.js
	 */
	ApiBaseAnnotation.prototype.SetSubject = function(subject) {
		subject = AscBuilder.GetStringParameter(subject, null);
		if (!subject) {
			AscBuilder.throwException("The subject parameter must be a non empty string");
		}

		this.Annot.SetSubject(subject);
		return true;
	};

	/**
	 * Gets annotation subject.
	 * @typeofeditors ["PDFE"]
	 * @returns {?string}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetSubject.js
	 */
	ApiBaseAnnotation.prototype.GetSubject = function() {
		return this.Annot.GetSubject();
	};

	/**
	 * Sets annotation display type.
	 * @typeofeditors ["PDFE"]
	 * @param {DisplayType} display
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetDisplay.js
	 */
	ApiBaseAnnotation.prototype.SetDisplay = function(display) {
		if (AscPDF.Api.Types.display[display] == undefined) {
			AscBuilder.throwException("The display parameter must be a one of DisplayType");
		}

		this.Annot.SetDisplay(AscPDF.Api.Types.display[display]);
		return true;
	};

	/**
	 * Gets annotation display type.
	 * @typeofeditors ["PDFE"]
	 * @returns {DisplayType}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetDisplay.js
	 */
	ApiBaseAnnotation.prototype.GetDisplay = function() {
		let nDisplay = this.Annot.GetDisplay();

		switch (nDisplay) {
			case AscPDF.Api.Types.display.visible: {
				return "visible";
			}
			case AscPDF.Api.Types.display.hidden: {
				return "hidden";
			}
			case AscPDF.Api.Types.display.noPrint: {
				return "noPrint";
			}
			case AscPDF.Api.Types.display.noView: {
				return "noView";
			}
		}
	};

	/**
	 * Sets annotation dash pattern.
	 * <note> The border style property must be set to "dashed". </note>
	 * @typeofeditors ["PDFE"]
	 * @param {number[]} pattern - A dash array defining a pattern of dashes and gaps to be used in drawing a dashed border. For example, a value of [3, 2] specifies a border drawn with 3-point dashes alternating with 2-point gaps.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetDashPattern.js
	 */
	ApiBaseAnnotation.prototype.SetDashPattern = function(pattern) {
		if (pattern.find(function(value) { value = AscBuilder.GetNumberParameter(value, null); if (!value) return true})) {
			AscBuilder.throwException("The pattern parameter must be an array with numbers");
		}

		this.Annot.SetDashPattern(pattern);
		return true;
	};

	/**
	 * Gets annotation dash pattern.
	 * @typeofeditors ["PDFE"]
	 * @returns {number[]}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetDashPattern.js
	 */
	ApiBaseAnnotation.prototype.GetDashPattern = function() {
		return this.Annot.GetDashPattern();
	};

	/**
	 * Sets annotation border effect style.
	 * <note> Can be applied to circle, square, freeText and polygon annotations </note>
	 * @typeofeditors ["PDFE"]
	 * @param {AnnotBorderEffectStyle} style
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetBorderEffectStyle.js
	 */
	ApiBaseAnnotation.prototype.SetBorderEffectStyle = function(style) {
		if (undefined == AscPDF.BORDER_EFFECT_STYLES[style]) {
			AscBuilder.throwException("The style parameter must be one of available");
		}

		this.Annot.SetBorderEffectStyle(AscPDF.BORDER_EFFECT_STYLES[style]);
		this.private_UpdateRect();

		return true;
	};

	/**
	 * Gets annotation border effect style.
	 * @typeofeditors ["PDFE"]
	 * @returns {AnnotBorderEffectStyle}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetBorderEffectStyle.js
	 */
	ApiBaseAnnotation.prototype.GetBorderEffectStyle = function() {
		let nBorderEffectStyle = this.Annot.GetBorderEffectStyle();

		switch (nBorderEffectStyle) {
			case AscPDF.BORDER_EFFECT_STYLES.none: {
				return "none";
			}
			case AscPDF.BORDER_EFFECT_STYLES.cloud: {
				return "cloud";
			}
		}
	};

	/**
	 * Sets annotation border effect intensity.
	 * <note> Can be applied to circle, square, freeText and polygon annotations </note>
	 * @typeofeditors ["PDFE"]
	 * @param {number} value
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/SetBorderEffectIntensity.js
	 */
	ApiBaseAnnotation.prototype.SetBorderEffectIntensity = function(value) {
		value = AscBuilder.GetNumberParameter(value, null);

		if (null == value || value < 0) {
			AscBuilder.throwException("The value parameter must be number greater than 0");
		}

		this.Annot.SetBorderEffectIntensity(value);
		this.private_UpdateRect();

		return true;
	};

	/**
	 * Gets annotation border effect intensity.
	 * @typeofeditors ["PDFE"]
	 * @param {number} value
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetBorderEffectIntensity.js
	 */
	ApiBaseAnnotation.prototype.GetBorderEffectIntensity = function() {
		return this.Annot.GetBorderEffectIntensity();
	};
	
	/**
	 * Adds reply on this annot.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiTextAnnotation} textAnnot
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/AddReply.js
	 */
	ApiBaseAnnotation.prototype.AddReply = function(textAnnot) {
		if (!(textAnnot instanceof ApiTextAnnotation)) {
			AscBuilder.throwException("The textAnnot parameter must be an ApiTextAnnotation class object");
		}

		if (this.Annot.IsUseContentAsComment() && this.Annot.GetContents() == null) {
			AscBuilder.throwException("Before add reply you need to set the contents property");
		}

		this.Annot.AddReply(textAnnot.private_GetImpl());
		return true;
	};

	/**
	 * Gets replies on this annot.
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiTextAnnotation[]}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/GetReplies.js
	 */
	ApiBaseAnnotation.prototype.GetReplies = function() {
		return this.Annot.GetReplies().map(private_GetAnnotApi);
	};

	/**
	 * Removes annotation from document.
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseAnnotation/Methods/Delete.js
	 */
	ApiBaseAnnotation.prototype.Delete = function() {
		let oDoc = private_GetLogicDocument();
		
		oDoc.RemoveAnnot(this.Annot.GetId());
		return true;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTextAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a text annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseAnnotation}
	 */
	function ApiTextAnnotation(oAnnot) {
		ApiBaseAnnotation.call(this, oAnnot);
	}

	ApiTextAnnotation.prototype = Object.create(ApiBaseAnnotation.prototype);
	ApiTextAnnotation.prototype.constructor = ApiTextAnnotation;

	/**
	 * Returns a type of the ApiTextAnnotation class.
	 * @memberof ApiTextAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"textAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiTextAnnotation/Methods/GetClassType.js
	 */
	ApiTextAnnotation.prototype.GetClassType = function() {
		return "textAnnot";
	};

	/**
	 * Sets icon type for this annotation.
	 * @memberof ApiTextAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {TextIconType} iconType
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTextAnnotation/Methods/SetIconType.js
	 */
	ApiTextAnnotation.prototype.SetIconType = function(iconType) {
		if (undefined == AscPDF.TEXT_ICONS_TYPES[iconType]) {
			AscBuilder.throwException("The iconType parameter must be one of available");
		}

		this.Annot.SetIconType(AscPDF.TEXT_ICONS_TYPES[iconType]);
		return true;
	};

	/**
	 * Gets icon type of this annotation.
	 * @memberof ApiTextAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {TextIconType}
	 * @see office-js-api/Examples/{Editor}/ApiTextAnnotation/Methods/GetIconType.js
	 */
	ApiTextAnnotation.prototype.GetIconType = function() {
		let nIconType = this.Annot.GetIconType();

		switch (nIconType) {
			case AscPDF.TEXT_ICONS_TYPES.check1:
			case AscPDF.TEXT_ICONS_TYPES.check2: {
				return "check";
			}
			case AscPDF.TEXT_ICONS_TYPES.circle: {
				return "circle";
			}
			case AscPDF.TEXT_ICONS_TYPES.comment: {
				return "comment";
			}
			case AscPDF.TEXT_ICONS_TYPES.cross: {
				return "cross";
			}
			case AscPDF.TEXT_ICONS_TYPES.crossH: {
				return "crossH";
			}
			case AscPDF.TEXT_ICONS_TYPES.help: {
				return "help";
			}
			case AscPDF.TEXT_ICONS_TYPES.insert: {
				return "insert";
			}
			case AscPDF.TEXT_ICONS_TYPES.key: {
				return "key";
			}
			case AscPDF.TEXT_ICONS_TYPES.newParagraph: {
				return "newParagraph";
			}
			case AscPDF.TEXT_ICONS_TYPES.note: {
				return "note";
			}
			case AscPDF.TEXT_ICONS_TYPES.paragraph: {
				return "paragraph";
			}
			case AscPDF.TEXT_ICONS_TYPES.rightArrow: {
				return "rightArrow";
			}
			case AscPDF.TEXT_ICONS_TYPES.rightPointer: {
				return "rightPointer";
			}
			case AscPDF.TEXT_ICONS_TYPES.star: {
				return "star";
			}
			case AscPDF.TEXT_ICONS_TYPES.upArrow: {
				return "upArrow";
			}
			case AscPDF.TEXT_ICONS_TYPES.upLeftArrow: {
				return "upLeftArrow";
			}
		}
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiCircleAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a circle annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseAnnotation}
	 */
	function ApiCircleAnnotation(oAnnot) {
		ApiBaseAnnotation.call(this, oAnnot);
	}

	ApiCircleAnnotation.prototype = Object.create(ApiBaseAnnotation.prototype);
	ApiCircleAnnotation.prototype.constructor = ApiCircleAnnotation;

	ApiCircleAnnotation.prototype.private_UpdateRect = function(rect) {
		if (!rect) {
			rect = this.Annot.GetRect();
		}

		AscCommon.History.StartNoHistoryMode();
		let aCurRect = this.Annot.GetRect();
		let aCurRD = this.Annot.GetRectangleDiff().slice();
		let nLineW = this.Annot.GetBorderWidth() * g_dKoef_pt_to_mm;
		this.Annot.SetRect(rect);
		this.Annot.SetRectangleDiff([0, 0, 0, 0]);
		this.Annot.recalcBounds();
		this.Annot.recalcGeometry();
		this.Annot.Recalculate(true);
		
		AscCommon.History.EndNoHistoryMode();
		
		let oGrBounds = this.Annot.bounds;
		let oShapeBounds = this.Annot.getRectBounds();

		rect[0] = (oGrBounds.l - nLineW) * g_dKoef_mm_to_pt;
		rect[1] = (oGrBounds.t - nLineW) * g_dKoef_mm_to_pt;
		rect[2] = (oGrBounds.r + nLineW) * g_dKoef_mm_to_pt;
		rect[3] = (oGrBounds.b + nLineW) * g_dKoef_mm_to_pt;

		this.Annot._rect = aCurRect;
		this.Annot._rectDiff = aCurRD;

		this.Annot.SetRect(rect);
		this.Annot.SetRectangleDiff([
			(oShapeBounds.l - oGrBounds.l + nLineW) * g_dKoef_mm_to_pt,
			(oShapeBounds.t - oGrBounds.t + nLineW) * g_dKoef_mm_to_pt,
			(oGrBounds.r - oShapeBounds.r + nLineW) * g_dKoef_mm_to_pt,
			(oGrBounds.b - oShapeBounds.b + nLineW) * g_dKoef_mm_to_pt
		]);
	};

	/**
	 * Returns a type of the ApiCircleAnnotation class.
	 * @memberof ApiCircleAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"circleAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiCircleAnnotation/Methods/GetClassType.js
	 */
	ApiCircleAnnotation.prototype.GetClassType = function() {
		return "circleAnnot";
	};

	/**
	 * Sets annotation rect difference.
	 * @memberof ApiCircleAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {RectDiff} rectDiff - A set of four numbers that shall describe the numerical differences between two rectangles.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiCircleAnnotation/Methods/SetRectDiff.js
	 */
	ApiCircleAnnotation.prototype.SetRectDiff = function(rectDiff) {
		if (!private_IsValidRectDiff(rectDiff)) {
			AscBuilder.throwException("The rectDiff parameter must be a valid rect diff");
		}

		this.Annot.SetRectangleDiff(rectDiff);
		return true;
	};

	/**
	 * Gets annotation rect difference.
	 * @memberof ApiCircleAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {Rect}
	 * @see office-js-api/Examples/{Editor}/ApiCircleAnnotation/Methods/GetRectDiff.js
	 */
	ApiCircleAnnotation.prototype.GetRectDiff = function() {
		return this.Annot.GetRectangleDiff();
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiSquareAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a square annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseAnnotation}
	 */
	function ApiSquareAnnotation(oAnnot) {
		ApiBaseAnnotation.call(this, oAnnot);
	}

	ApiSquareAnnotation.prototype = Object.create(ApiBaseAnnotation.prototype);
	ApiSquareAnnotation.prototype.constructor = ApiSquareAnnotation;

	ApiSquareAnnotation.prototype.private_UpdateRect = function(rect) {
		if (!rect) {
			rect = this.Annot.GetRect();
		}

		AscCommon.History.StartNoHistoryMode();
		let aCurRect = this.Annot.GetRect();
		let aCurRD = this.Annot.GetRectangleDiff().slice();
		let nLineW = this.Annot.GetBorderWidth() * g_dKoef_pt_to_mm;
		this.Annot.SetRect(rect);
		this.Annot.SetRectangleDiff([0, 0, 0, 0]);
		this.Annot.recalcBounds();
		this.Annot.recalcGeometry();
		this.Annot.Recalculate(true);
		
		AscCommon.History.EndNoHistoryMode();
		
		let oGrBounds = this.Annot.bounds;
		let oShapeBounds = this.Annot.getRectBounds();

		rect[0] = (oGrBounds.l - nLineW) * g_dKoef_mm_to_pt;
		rect[1] = (oGrBounds.t - nLineW) * g_dKoef_mm_to_pt;
		rect[2] = (oGrBounds.r + nLineW) * g_dKoef_mm_to_pt;
		rect[3] = (oGrBounds.b + nLineW) * g_dKoef_mm_to_pt;

		this.Annot._rect = aCurRect;
		this.Annot._rectDiff = aCurRD;

		this.Annot.SetRect(rect);
		this.Annot.SetRectangleDiff([
			(oShapeBounds.l - oGrBounds.l + nLineW) * g_dKoef_mm_to_pt,
			(oShapeBounds.t - oGrBounds.t + nLineW) * g_dKoef_mm_to_pt,
			(oGrBounds.r - oShapeBounds.r + nLineW) * g_dKoef_mm_to_pt,
			(oGrBounds.b - oShapeBounds.b + nLineW) * g_dKoef_mm_to_pt
		]);
	};

	/**
	 * Returns a type of the ApiSquareAnnotation class.
	 * @memberof ApiSquareAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"squareAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiSquareAnnotation/Methods/GetClassType.js
	 */
	ApiSquareAnnotation.prototype.GetClassType = function() {
		return "squareAnnot";
	};

	/**
	 * Sets annotation rect difference.
	 * @memberof ApiSquareAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {RectDiff} rectDiff - A set of four numbers that shall describe the numerical differences between two rectangles.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiSquareAnnotation/Methods/SetRectDiff.js
	 */
	ApiSquareAnnotation.prototype.SetRectDiff = function(rectDiff) {
		if (!private_IsValidRectDiff(rectDiff)) {
			AscBuilder.throwException("The rectDiff parameter must be a valid rect diff");
		}

		this.Annot.SetRectangleDiff(rectDiff);
		return true;
	};

	/**
	 * Gets annotation rect difference.
	 * @memberof ApiSquareAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {Rect}
	 * @see office-js-api/Examples/{Editor}/ApiSquareAnnotation/Methods/GetRectDiff.js
	 */
	ApiSquareAnnotation.prototype.GetRectDiff = function() {
		return this.Annot.GetRectangleDiff();
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiFreeTextAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a freeText annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseAnnotation}
	 */
	function ApiFreeTextAnnotation(oAnnot) {
		ApiBaseAnnotation.call(this, oAnnot);
	}

	ApiFreeTextAnnotation.prototype = Object.create(ApiBaseAnnotation.prototype);
	ApiFreeTextAnnotation.prototype.constructor = ApiFreeTextAnnotation;

	/**
	 * Returns a type of the ApiFreeTextAnnotation class.
	 * @memberof ApiFreeTextAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"freeTextAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiFreeTextAnnotation/Methods/GetClassType.js
	 */
	ApiFreeTextAnnotation.prototype.GetClassType = function() {
		return "freeTextAnnot";
	};

	/**
	 * Sets intent type for this annotation.
	 * @memberof ApiFreeTextAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {FreeTextIntent} intentType
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiFreeTextAnnotation/Methods/SetIntent.js
	 */
	ApiFreeTextAnnotation.prototype.SetIntent = function(intentType) {
		if (undefined == AscPDF.FREE_TEXT_INTENT_TYPE[intentType]) {
			AscBuilder.throwException("The intentType parameter must be one of available");
		}

		this.Annot.SetIntent(AscPDF.FREE_TEXT_INTENT_TYPE[intentType]);
		return true;
	};

	/**
	 * Gets intent type of this annotation.
	 * @memberof ApiFreeTextAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {FreeTextIntent}
	 * @see office-js-api/Examples/{Editor}/ApiFreeTextAnnotation/Methods/GetIntent.js
	 */
	ApiFreeTextAnnotation.prototype.GetIntent = function() {
		let nIntentType = this.Annot.GetIntent();

		switch (nIntentType) {
			case AscPDF.FREE_TEXT_INTENT_TYPE.freeText: {
				return "check";
			}
			case AscPDF.FREE_TEXT_INTENT_TYPE.freeTextCallout: {
				return "circle";
			}
		}
	};
	
	/**
	 * Sets annot callout.
	 * @memberof ApiFreeTextAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {FreeTextCallout} callout
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiFreeTextAnnotation/Methods/SetCallout.js
	 */
	ApiFreeTextAnnotation.prototype.SetCallout = function(callout) {
		callout = AscBuilder.GetArrayParameter(callout, null);
		if (!callout || callout.length !== 3) {
			AscBuilder.throwException("The callout parameter must be an array with 3 Point elements.");
		}

		callout.forEach(function(point) {
			private_CheckPoint(point);
		});

		let flatCallout = [];
		callout.forEach(function(point) {
			flatCallout.push(point["x"], point["y"]);
		});

		this.Annot.SetCallout(flatCallout);
		return true;
	};

	/**
	 * Gets annot callout.
	 * @memberof ApiFreeTextAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {FreeTextCallout}
	 * @see office-js-api/Examples/{Editor}/ApiFreeTextAnnotation/Methods/GetCallout.js
	 */
	ApiFreeTextAnnotation.prototype.GetCallout = function() {
		let aCallout = this.Annot.GetCallout();

		let aResult = [];
		for (let i = 0; i < aCallout.length - 1; i++) {
			aResult.push({
				"x": aCallout[i],
				"y": aCallout[i + 1],
			});
		}

		return aResult;
	};

	/**
	 * Sets annotation rect difference.
	 * @memberof ApiFreeTextAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {RectDiff} rectDiff - A set of four numbers that shall describe the numerical differences between two rectangles.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiFreeTextAnnotation/Methods/SetRectDiff.js
	 */
	ApiFreeTextAnnotation.prototype.SetRectDiff = function(rectDiff) {
		if (!private_IsValidRectDiff(rectDiff)) {
			AscBuilder.throwException("The rectDiff parameter must be a valid rect diff");
		}

		this.Annot.SetRectangleDiff(rectDiff);
		return true;
	};

	/**
	 * Gets annotation rect difference.
	 * @memberof ApiFreeTextAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {Rect}
	 * @see office-js-api/Examples/{Editor}/ApiFreeTextAnnotation/Methods/GetRectDiff.js
	 */
	ApiFreeTextAnnotation.prototype.GetRectDiff = function() {
		return this.Annot.GetRectangleDiff();
	};

	/**
	 * Gets annotation rich content.
	 * @memberof ApiFreeTextAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiRichContent}
	 * @see office-js-api/Examples/{Editor}/ApiFreeTextAnnotation/Methods/GetContent.js
	 */
	ApiFreeTextAnnotation.prototype.GetContent = function() {
		return new ApiRichContent(this.Annot.GetDocContent());
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiLineAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a line annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseAnnotation}
	 */
	function ApiLineAnnotation(oAnnot) {
		ApiBaseAnnotation.call(this, oAnnot);
	}

	ApiLineAnnotation.prototype = Object.create(ApiBaseAnnotation.prototype);
	ApiLineAnnotation.prototype.constructor = ApiLineAnnotation;

	/**
	 * Returns a type of the ApiLineAnnotation class.
	 * @memberof ApiLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"lineAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiLineAnnotation/Methods/GetClassType.js
	 */
	ApiLineAnnotation.prototype.GetClassType = function() {
		return "lineAnnot";
	};

	/**
	 * Sets a line start point.
	 * @memberof ApiLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {Point} point
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiLineAnnotation/Methods/SetStartPoint.js
	 */
	ApiLineAnnotation.prototype.SetStartPoint = function(point) {
		private_CheckPoint(point);

		let aCurPoints = this.Annot.GetLinePoints();
		let aNewPoints = aCurPoints.slice();

		aNewPoints[0] = point['x'];
		aNewPoints[1] = point['y'];

		this.Annot.SetLinePoints(aNewPoints);
		return true;
	};

	/**
	 * Gets a line start point.
	 * @memberof ApiLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {Point}
	 * @see office-js-api/Examples/{Editor}/ApiLineAnnotation/Methods/GetStartPoint.js
	 */
	ApiLineAnnotation.prototype.GetStartPoint = function() {
		let aPoints = this.Annot.GetLinePoints();
		return {
			"x": aPoints[0],
			"y": aPoints[1]
		}
	};

	/**
	 * Sets a line end point.
	 * @memberof ApiLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {Point} point
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiLineAnnotation/Methods/SetEndPoint.js
	 */
	ApiLineAnnotation.prototype.SetEndPoint = function(point) {
		private_CheckPoint(point);

		let aCurPoints = this.Annot.GetLinePoints();
		let aNewPoints = aCurPoints.slice();

		aNewPoints[2] = point['x'];
		aNewPoints[3] = point['y'];

		this.Annot.SetLinePoints(aNewPoints);
		return true;
	};

	/**
	 * Gets a line end point.
	 * @memberof ApiLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {Point}
	 * @see office-js-api/Examples/{Editor}/ApiLineAnnotation/Methods/GetEndPoint.js
	 */
	ApiLineAnnotation.prototype.GetEndPoint = function() {
		let aPoints = this.Annot.GetLinePoints();
		return {
			"x": aPoints[2],
			"y": aPoints[3]
		}
	};

	/**
	 * Sets a line start style.
	 * @memberof ApiLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {LineEndStyle} style
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiLineAnnotation/Methods/SetStartStyle.js
	 */
	ApiLineAnnotation.prototype.SetStartStyle = function(style) {
		if (undefined == private_GetInnerLineEndType(style)) {
			AscBuilder.throwException("The style parameter must be one of available");
		}

		this.Annot.SetLineStart(private_GetInnerLineEndType(style));
		this.Annot.SetRect(this.Annot.private_CalcBoundingRect());
		return true;
	};

	/**
	 * Gets a line start style.
	 * @memberof ApiLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {LineEndStyle}
	 * @see office-js-api/Examples/{Editor}/ApiLineAnnotation/Methods/GetStartStyle.js
	 */
	ApiLineAnnotation.prototype.GetStartStyle = function() {
		let nStyle = this.Annot.GetLineStart();
		return private_GetStrLineEndType(nStyle);
	};

	/**
	 * Sets a line end style.
	 * @memberof ApiLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {LineEndStyle} style
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiLineAnnotation/Methods/SetEndStyle.js
	 */
	ApiLineAnnotation.prototype.SetEndStyle = function(style) {
		if (undefined == private_GetInnerLineEndType(style)) {
			AscBuilder.throwException("The style parameter must be one of available");
		}

		this.Annot.SetLineEnd(private_GetInnerLineEndType(style));
		this.Annot.SetRect(this.Annot.private_CalcBoundingRect());
		return true;
	};

	/**
	 * Gets a line end style.
	 * @memberof ApiLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {LineEndStyle}
	 * @see office-js-api/Examples/{Editor}/ApiLineAnnotation/Methods/GetEndStyle.js
	 */
	ApiLineAnnotation.prototype.GetEndStyle = function() {
		let nStyle = this.Annot.GetLineEnd();
		return private_GetStrLineEndType(nStyle);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiInkAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a ink annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseAnnotation}
	 */
	function ApiInkAnnotation(oAnnot) {
		ApiBaseAnnotation.call(this, oAnnot);
	}

	ApiInkAnnotation.prototype = Object.create(ApiBaseAnnotation.prototype);
	ApiInkAnnotation.prototype.constructor = ApiInkAnnotation;

	/**
	 * Returns a type of the ApiInkAnnotation class.
	 * @memberof ApiInkAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"inkAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiInkAnnotation/Methods/GetClassType.js
	 */
	ApiInkAnnotation.prototype.GetClassType = function() {
		return "inkAnnot";
	};

	/**
	 * Sets ink path list.
	 * @memberof ApiInkAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {PathList} inkPaths - ink path list
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiInkAnnotation/Methods/SetPathList.js
	 */
	ApiInkAnnotation.prototype.SetPathList = function(inkPaths) {
		inkPaths = AscBuilder.GetArrayParameter(inkPaths, []);
		if (inkPaths.length == 0)
			AscBuilder.throwException("The inkPaths parameter must be a non empty array");

		inkPaths.forEach(function(path) {
			path = AscBuilder.GetArrayParameter(path, []);
			if (path.length == 0)
				AscBuilder.throwException("The ink path parameter must be a non empty array");

			path.forEach(function(point) {
				private_CheckPoint(point);
			});
		});

		this.Annot.SetInkPoints(inkPaths.map(function(path) {
			let flatPath = [];
			path.forEach(function(point) {
				flatPath.push(point["x"], point["y"]);
			});

			return flatPath;
		}));

		return true;
	};

	/**
	 * Gets ink path list.
	 * @memberof ApiInkAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {PathList}
	 * @see office-js-api/Examples/{Editor}/ApiInkAnnotation/Methods/GetPathList.js
	 */
	ApiInkAnnotation.prototype.GetPathList = function() {
		let aInkPaths = this.Annot.GetInkPoints();

		return aInkPaths.map(function(path) {
			let aPath = [];
			for (let i = 0; i < path.length - 1; i+= 2) {
				aPath.push({
					"x": path[i],
					"y": path[i+1]
				});
			}

			return aPath;
		});
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiPolygonAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a polygon annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseAnnotation}
	 */
	function ApiPolygonAnnotation(oAnnot) {
		ApiBaseAnnotation.call(this, oAnnot);
	}

	ApiPolygonAnnotation.prototype = Object.create(ApiBaseAnnotation.prototype);
	ApiPolygonAnnotation.prototype.constructor = ApiPolygonAnnotation;

	/**
	 * Returns a type of the ApiPolygonAnnotation class.
	 * @memberof ApiPolygonAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"polygonAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiPolygonAnnotation/Methods/GetClassType.js
	 */
	ApiPolygonAnnotation.prototype.GetClassType = function() {
		return "polygonAnnot";
	};

	/**
	 * Sets vertices to polygon annot.
	 * @memberof ApiPolygonAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {Path} path - polygon path
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiPolygonAnnotation/Methods/SetVertices.js
	 */
	ApiPolygonAnnotation.prototype.SetVertices = function(path) {
		path = AscBuilder.GetArrayParameter(path, []);
		if (path.length == 0)
			AscBuilder.throwException("The path parameter must be a non empty array");

		path.forEach(function(point) {
			private_CheckPoint(point);
		});

		let aVertices = [];
		path.forEach(function(point) {
			aVertices.push(point["x"], point["y"]);
		});

		this.Annot.SetVertices(aVertices);
		return true;
	};

	/**
	 * Gets ink path list.
	 * @memberof ApiPolygonAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {Path}
	 * @see office-js-api/Examples/{Editor}/ApiPolygonAnnotation/Methods/GetVertices.js
	 */
	ApiPolygonAnnotation.prototype.GetVertices = function() {
		let aVertices = this.Annot.GetVertices();

		let aPath = [];
		for (let i = 0; i < aVertices.length - 1; i+= 2) {
			aPath.push({
				"x": aVertices[i],
				"y": aVertices[i+1]
			});
		}

		return aPath;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiPolyLineAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a polyline annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseAnnotation}
	 */
	function ApiPolyLineAnnotation(oAnnot) {
		ApiBaseAnnotation.call(this, oAnnot);
	}

	ApiPolyLineAnnotation.prototype = Object.create(ApiBaseAnnotation.prototype);
	ApiPolyLineAnnotation.prototype.constructor = ApiPolyLineAnnotation;

	/**
	 * Returns a type of the ApiPolyLineAnnotation class.
	 * @memberof ApiPolyLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"polyLineAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiPolyLineAnnotation/Methods/GetClassType.js
	 */
	ApiPolyLineAnnotation.prototype.GetClassType = function() {
		return "polyLineAnnot";
	};

	/**
	 * Sets vertices to polyline annot.
	 * @memberof ApiPolyLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {Path} path - polyline path
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiPolyLineAnnotation/Methods/SetVertices.js
	 */
	ApiPolyLineAnnotation.prototype.SetVertices = function(path) {
		path = AscBuilder.GetArrayParameter(path, []);
		if (path.length == 0)
			AscBuilder.throwException("The path parameter must be a non empty array");

		path.forEach(function(point) {
			private_CheckPoint(point);
		});

		let aVertices = [];
		path.forEach(function(point) {
			aVertices.push(point["x"], point["y"]);
		});

		this.Annot.SetVertices(aVertices);
		return true;
	};

	/**
	 * Gets ink path list.
	 * @memberof ApiPolyLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {Path}
	 * @see office-js-api/Examples/{Editor}/ApiPolyLineAnnotation/Methods/GetVertices.js
	 */
	ApiPolyLineAnnotation.prototype.GetVertices = function() {
		let aVertices = this.Annot.GetVertices();

		let aPath = [];
		for (let i = 0; i < aVertices.length - 1; i+= 2) {
			aPath.push({
				"x": aVertices[i],
				"y": aVertices[i+1]
			});
		}

		return aPath;
	};

	/**
	 * Sets a line start style.
	 * @memberof ApiPolyLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {LineEndStyle} style
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiPolyLineAnnotation/Methods/SetStartStyle.js
	 */
	ApiPolyLineAnnotation.prototype.SetStartStyle = function(style) {
		if (undefined == private_GetInnerLineEndType(style)) {
			AscBuilder.throwException("The style parameter must be one of available");
		}

		this.Annot.SetLineStart(private_GetInnerLineEndType(style));
		this.Annot.SetRect(this.Annot.private_CalcBoundingRect());
		return true;
	};

	/**
	 * Gets a line start style.
	 * @memberof ApiPolyLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {LineEndStyle}
	 * @see office-js-api/Examples/{Editor}/ApiPolyLineAnnotation/Methods/GetStartStyle.js
	 */
	ApiPolyLineAnnotation.prototype.GetStartStyle = function() {
		let nStyle = this.Annot.GetLineStart();
		return private_GetStrLineEndType(nStyle);
	};

	/**
	 * Sets a line end style.
	 * @memberof ApiPolyLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {LineEndStyle} style
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiPolyLineAnnotation/Methods/SetEndStyle.js
	 */
	ApiPolyLineAnnotation.prototype.SetEndStyle = function(style) {
		if (undefined == private_GetInnerLineEndType(style)) {
			AscBuilder.throwException("The style parameter must be one of available");
		}

		this.Annot.SetLineEnd(private_GetInnerLineEndType(style));
		this.Annot.SetRect(this.Annot.private_CalcBoundingRect());
		return true;
	};

	/**
	 * Gets a line end style.
	 * @memberof ApiPolyLineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {LineEndStyle}
	 * @see office-js-api/Examples/{Editor}/ApiPolyLineAnnotation/Methods/GetEndStyle.js
	 */
	ApiPolyLineAnnotation.prototype.GetEndStyle = function() {
		let nStyle = this.Annot.GetLineEnd();
		return private_GetStrLineEndType(nStyle);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiStampAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a stamp annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseAnnotation}
	 */
	function ApiStampAnnotation(oAnnot) {
		ApiBaseAnnotation.call(this, oAnnot);
	}

	ApiStampAnnotation.prototype = Object.create(ApiBaseAnnotation.prototype);
	ApiStampAnnotation.prototype.constructor = ApiStampAnnotation;

	/**
	 * Returns a type of the ApiStampAnnotation class.
	 * @memberof ApiStampAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"stampAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiStampAnnotation/Methods/GetClassType.js
	 */
	ApiStampAnnotation.prototype.GetClassType = function() {
		return "stampAnnot";
	};

	/**
	 * Gets stamp type.
	 * @memberof ApiStampAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {StampType}
	 * @see office-js-api/Examples/{Editor}/ApiStampAnnotation/Methods/GetType.js
	 */
	ApiStampAnnotation.prototype.GetType = function() {
		return this.Annot.GetIconType();
	};

	/**
	 * Sets stamp size scale.
	 * @memberof ApiStampAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {number} scale - size scale
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiStampAnnotation/Methods/SetScale.js
	 */
	ApiStampAnnotation.prototype.SetScale = function(scale) {
		let aInRect = this.Annot.GetInRect();
		let nInRectRot = AscPDF.getQuadsRot([aInRect[0], aInRect[3], aInRect[4], aInRect[3], aInRect[4], aInRect[1], aInRect[0], aInRect[1]]);
		let aInRectNoRot = AscPDF.rotateRect([aInRect[0], aInRect[3], aInRect[4], aInRect[3], aInRect[4], aInRect[1], aInRect[0], aInRect[1]], -nInRectRot);

		let minX = Infinity, maxX = -Infinity;
		let minY = Infinity, maxY = -Infinity;

		for (let i = 0; i < aInRectNoRot.length; i += 2) {
			let x = aInRectNoRot[i];
			let y = aInRectNoRot[i + 1];

			if (x < minX) minX = x;
			if (x > maxX) maxX = x;
			if (y < minY) minY = y;
			if (y > maxY) maxY = y;
		}

		let nOrigExtX = maxX - minX;
		let nOrigExtY = maxY - minY;

		AscCommon.History.StartNoHistoryMode();
		let aCurRect = this.Annot.GetRect();
		let oCurXfrm = this.Annot.getXfrm();

		let nCurExtX = this.Annot.getXfrmExtX();
		let nCurExtY = this.Annot.getXfrmExtY();
		let nCurOffX = this.Annot.getXfrmOffX();
		let nCurOffY = this.Annot.getXfrmOffY();

		let aNewRect = [minX, minY, minX + nOrigExtX * scale, minY + nOrigExtY * scale];

		this.Annot.recalcBounds();
		this.Annot.recalcGeometry();
		
		this.Annot.SetRect(aNewRect);
		AscPDF.CAnnotationBase.prototype.RecalcSizes.call(this.Annot);
		this.Annot.recalculate();
		
		AscCommon.History.EndNoHistoryMode();
		
		let oGrBounds = this.Annot.bounds;
		aNewRect[0] = oGrBounds.l * g_dKoef_mm_to_pt;
		aNewRect[1] = oGrBounds.t * g_dKoef_mm_to_pt;
		aNewRect[2] = oGrBounds.r * g_dKoef_mm_to_pt;
		aNewRect[3] = oGrBounds.b * g_dKoef_mm_to_pt;

		this.Annot._rect = aCurRect;
		oCurXfrm.extX = nCurExtX;
		oCurXfrm.extY = nCurExtY;
		oCurXfrm.offX = nCurOffX;
		oCurXfrm.offY = nCurOffY;

		this.Annot.SetRect(aNewRect);
		this.Annot.Recalculate();

		return true;
	};

	/**
	 * Gets stamp size scale.
	 * @memberof ApiStampAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiStampAnnotation/Methods/GetScale.js
	 */
	ApiStampAnnotation.prototype.GetScale = function() {
		return this.Annot.GetOriginViewScale();
	};

	/**
	 * Sets stamp rotate.
	 * @memberof ApiStampAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {Degree} angle
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiStampAnnotation/Methods/SetRotation.js
	 */
	ApiStampAnnotation.prototype.SetRotation = function(angle) {
		angle = AscBuilder.GetNumberParameter(angle, 0);
		angle = private_NormalizeDegree(angle);

		let oXfrm = this.Annot.getXfrm();
		oXfrm.setRot(-angle * (Math.PI / 180));

		return true;
	};

	/**
	 * Gets stamp rotate.
	 * @memberof ApiStampAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {Degree}
	 * @see office-js-api/Examples/{Editor}/ApiStampAnnotation/Methods/GetRotation.js
	 */
	ApiStampAnnotation.prototype.GetRotation = function() {
		return this.Annot.GetRotate();
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiBaseMarkupAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a base markup annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseAnnotation}
	 */
	function ApiBaseMarkupAnnotation(oAnnot) {
		ApiBaseAnnotation.call(this, oAnnot);
	}

	ApiBaseMarkupAnnotation.prototype = Object.create(ApiBaseAnnotation.prototype);
	ApiBaseMarkupAnnotation.prototype.constructor = ApiBaseMarkupAnnotation;

	/**
	 * Sets quads to current markup annotation.
	 * @memberof ApiBaseMarkupAnnotation
	 * @typeofeditors ["PDFE"]
	 * @param {Quad[]} quads
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiBaseMarkupAnnotation/Methods/SetQuads.js
	 */
	ApiBaseMarkupAnnotation.prototype.SetQuads = function(quads) {
		quads = AscBuilder.GetArrayParameter(quads, null);
		if (!quads) {
			AscBuilder.throwException("The quads parameter must be a valid array");
		}

		quads.forEach(function(quad) {
			if (!private_IsValidQuad(quad)) {
				AscBuilder.throwException("The quad must be a valid quad");
			}
		});

		let minX = Infinity, maxX = -Infinity;
		let minY = Infinity, maxY = -Infinity;

		for (let i = 0; i < quads.length; i++) {
			for (let j = 0; j < quads[i].length; j += 2) {
				let x = quads[i][j];
				let y = quads[i][j + 1];

				if (x < minX) minX = x;
				if (x > maxX) maxX = x;
				if (y < minY) minY = y;
				if (y > maxY) maxY = y;
			}
		}

		this.Annot.SetQuads(quads);
		this.Annot.SetRect([minX - 1, minY - 1, maxX + 1, maxY + 1]);
		return true;
	};

	/**
	 * Gets quads from current markup annotation.
	 * @memberof ApiBaseMarkupAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {Quad[]}
	 * @see office-js-api/Examples/{Editor}/ApiBaseMarkupAnnotation/Methods/GetQuads.js
	 */
	ApiBaseMarkupAnnotation.prototype.GetQuads = function() {
		return this.Annot.GetQuads();
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiHighlightAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a highlight annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseMarkupAnnotation}
	 */
	function ApiHighlightAnnotation(oAnnot) {
		ApiBaseMarkupAnnotation.call(this, oAnnot);
	}

	ApiHighlightAnnotation.prototype = Object.create(ApiBaseMarkupAnnotation.prototype);
	ApiHighlightAnnotation.prototype.constructor = ApiHighlightAnnotation;

	/**
	 * Returns a type of the ApiHighlightAnnotation class.
	 * @memberof ApiHighlightAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"highlightAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiHighlightAnnotation/Methods/GetClassType.js
	 */
	ApiHighlightAnnotation.prototype.GetClassType = function() {
		return "highlightAnnot";
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiStrikeoutAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a strikeout annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseMarkupAnnotation}
	 */
	function ApiStrikeoutAnnotation(oAnnot) {
		ApiBaseMarkupAnnotation.call(this, oAnnot);
	}

	ApiStrikeoutAnnotation.prototype = Object.create(ApiBaseMarkupAnnotation.prototype);
	ApiStrikeoutAnnotation.prototype.constructor = ApiStrikeoutAnnotation;

	/**
	 * Returns a type of the ApiStrikeoutAnnotation class.
	 * @memberof ApiStrikeoutAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"strikeoutAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiStrikeoutAnnotation/Methods/GetClassType.js
	 */
	ApiStrikeoutAnnotation.prototype.GetClassType = function() {
		return "strikeoutAnnot";
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiUnderlineAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a underline annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseMarkupAnnotation}
	 */
	function ApiUnderlineAnnotation(oAnnot) {
		ApiBaseMarkupAnnotation.call(this, oAnnot);
	}

	ApiUnderlineAnnotation.prototype = Object.create(ApiBaseMarkupAnnotation.prototype);
	ApiUnderlineAnnotation.prototype.constructor = ApiUnderlineAnnotation;

	/**
	 * Returns a type of the ApiUnderlineAnnotation class.
	 * @memberof ApiUnderlineAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"underlineAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiUnderlineAnnotation/Methods/GetClassType.js
	 */
	ApiUnderlineAnnotation.prototype.GetClassType = function() {
		return "underlineAnnot";
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiCaretAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a caret annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseMarkupAnnotation}
	 */
	function ApiCaretAnnotation(oAnnot) {
		ApiBaseMarkupAnnotation.call(this, oAnnot);
	}

	ApiCaretAnnotation.prototype = Object.create(ApiBaseMarkupAnnotation.prototype);
	ApiCaretAnnotation.prototype.constructor = ApiCaretAnnotation;

	/**
	 * Returns a type of the ApiCaretAnnotation class.
	 * @memberof ApiCaretAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"caretAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiCaretAnnotation/Methods/GetClassType.js
	 */
	ApiCaretAnnotation.prototype.GetClassType = function() {
		return "caretAnnot";
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiRedactAnnotation
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a redact annotation.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 * @extends {ApiBaseMarkupAnnotation}
	 */
	function ApiRedactAnnotation(oAnnot) {
		ApiBaseMarkupAnnotation.call(this, oAnnot);
	}

	ApiRedactAnnotation.prototype = Object.create(ApiBaseMarkupAnnotation.prototype);
	ApiRedactAnnotation.prototype.constructor = ApiRedactAnnotation;

	/**
	 * Returns a type of the ApiRedactAnnotation class.
	 * @memberof ApiRedactAnnotation
	 * @typeofeditors ["PDFE"]
	 * @returns {"redactAnnot"}
	 * @see office-js-api/Examples/{Editor}/ApiRedactAnnotation/Methods/GetClassType.js
	 */
	ApiRedactAnnotation.prototype.GetClassType = function() {
		return "redactAnnot";
	};

	private_WrapClassMethods(ApiRedactAnnotation, function(method, args) {
		if (this.Annot.GetRedactId()) {
			AscBuilder.throwException("This Redact annot is already applied and can't be used");
		}
	});

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiRichContent
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a rich content.
	 * @constructor
	 * @typeofeditors ["PDFE"]
	 */
	function ApiRichContent(oContent) {
		this.Document = oContent;
	}

	/**
	 * Returns a type of the ApiRichContent class. 
	 * @memberof ApiRichContent
	 * @typeofeditors ["PDFE"]
	 * @returns {"richContent"}
	 * @see office-js-api/Examples/{Editor}/ApiRichContent/Methods/GetClassType.js
	 */
	ApiRichContent.prototype.GetClassType = function() {
		return "richContent";
	};

	/**
	 * Returns a number of elements in the current document.
	 * @memberof ApiRichContent
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiRichContent/Methods/GetElementsCount.js
	 */
	ApiRichContent.prototype.GetElementsCount = AscBuilder.ApiDocumentContent.prototype.GetElementsCount;

	/**
	 * Returns an rich paragraph by its position in the content.
	 * @memberof ApiRichContent
	 * @param {number} pos - The element position that will be taken from the content.
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiRichParagraph}
	 * @see office-js-api/Examples/{Editor}/ApiRichContent/Methods/GetElement.js
	 */
	ApiRichContent.prototype.GetElement = function(pos) {
		if (!this.Document.Content[pos])
			return null;

		return new ApiRichParagraph(this.Document.Content[pos]);
	};

	/**
	 * Adds a rich paragraph using its position in rich content.
	 * @memberof ApiRichContent
	 * @typeofeditors ["PDFE"]
	 * @param {number} pos - The position where the rich paragraph will be added.
	 * @param {ApiRichParagraph} richPara - The rich paragraph which will be added at the current position.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichContent/Methods/AddElement.js
	 */
	ApiRichContent.prototype.AddElement = function(pos, richPara) {
		if (!(richPara instanceof ApiRichParagraph)) {
			AscBuilder.throwException("The richPara parameter must be an ApiRichParagraph object");
		}

		let nMaxPos = this.GetElementsCount();

		pos = AscBuilder.GetNumberParameter(pos, null);
		if (null == pos || pos < 0 || pos > nMaxPos) {
			AscBuilder.throwException("The pos parameter must be a valid position");
		}

		let oElm = richPara.private_GetImpl();
		if (oElm.IsUseInDocument()) {
			AscBuilder.throwException("The richPara already in the document");
		}

		this.Document.Internal_Content_Add(pos, oElm);
		return true;
	};

	/**
	 * Pushes a rich paragraph to a rich content.
	 * @memberof ApiRichContent
	 * @typeofeditors ["PDFE"]
	 * @param {ApiRichParagraph} richPara - The rich paragraph which will be pushed to the rich content.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichContent/Methods/Push.js
	 */
	ApiRichContent.prototype.Push = function(richPara) {
		return this.AddElement(this.GetElementsCount(), richPara);
	};

	/**
	 * Removes all the elements from the current rich content.
	 * <note>When all elements are removed, a new empty rich paragraph is automatically created.</note>
	 * @memberof ApiRichContent
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichContent/Methods/RemoveAllElements.js
	 */
	ApiRichContent.prototype.RemoveAllElements = AscBuilder.ApiDocumentContent.prototype.RemoveAllElements;

	/**
	 * Removes an element using the position specified.
	 * @memberof ApiRichContent
	 * @typeofeditors ["PDFE"]
	 * @param {number} pos - The element number (position) in the rich content.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichContent/Methods/RemoveElement.js
	 */
	ApiRichContent.prototype.RemoveElement = AscBuilder.ApiDocumentContent.prototype.RemoveElement;

	/**
	 * Returns an array of rich paragraphs from the current rich content object.
	 * @memberof ApiRichContent
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} getCopies - Specifies if the copies of the document elements will be returned or not.
	 * @returns {ApiRichParagraph[]}
	 * @see office-js-api/Examples/{Editor}/ApiRichContent/Methods/GetContent.js
	 */
	ApiRichContent.prototype.GetContent = function(getCopies) {
		let aContent = [];
		let oTempElm = null;

		for (var nElm = 0; nElm < this.Document.Content.length; nElm++) {
			oTempElm = this.Document.Content[nElm];

			if (getCopies)
				oTempElm = oTempElm.Copy();

			aContent.push(new ApiRichParagraph(oTempElm));
		}

		return aContent;
	};

	/**
	 * Returns the inner text of the current document content object.
	 * @memberof ApiRichContent
	 * @typeofeditors ["PDFE"]
	 * @param {object} [options] - Options for formatting the returned text.
	 * @param {string} [options.ParaSeparator='\r\n'] - Defines how the paragraph separator will be specified in the resulting string. Any symbol can be used. The default separator is "\r\n".
	 * @param {string} [options.TabSymbol='\t'] - Defines how the tab will be specified in the resulting string. Any symbol can be used. The default symbol is "\t".
	 * @param {string} [options.NewLineSeparator='\r'] - Defines how the line separator will be specified in the resulting string. Any symbol can be used. The default separator is "\r".
	 * @return {string}
	 * @see office-js-api/Examples/{Editor}/ApiRichContent/Methods/GetText.js
	 */
	ApiRichContent.prototype.GetText = AscBuilder.ApiDocumentContent.prototype.GetText;

	/**
	 * Returns the current paragraph where the cursor is located.
	 * @memberof ApiRichContent
	 * @typeofeditors ["PDFE"]
	 * @return {?ApiRichParagraph}
	 * @see office-js-api/Examples/{Editor}/ApiRichContent/Methods/GetCurrentParagraph.js
	 */
	ApiRichContent.prototype.GetCurrentParagraph = function() {
		let oPara = this.Document.GetCurrentParagraph();
		if (!oPara) {
			return null;
		}

		return new ApiRichParagraph(oPara);
	};

	/**
	 * Returns the current run where the cursor is located.
	 * @memberof ApiRichContent
	 * @typeofeditors ["PDFE"]
	 * @return {?ApiRichRun}
	 * @see office-js-api/Examples/{Editor}/ApiRichContent/Methods/GetCurrentRun.js
	 */
	ApiRichContent.prototype.GetCurrentRun = function() {
		let oRun = this.Document.GetCurrentRun();
		if (!oRun) {
			return null;
		}

		return new ApiRichRun(oRun);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiRichParaPr
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing the rich paragraph properties.
	 * @constructor
	 */
	function ApiRichParaPr(Parent, ParaPr) {
		this.Parent = Parent;
		this.ParaPr = ParaPr;
	}

	/**
	 * Returns a type of the ApiRichParaPr class.
	 * @memberof ApiRichParaPr
	 * @typeofeditors ["PDFE"]
	 * @returns {"richParaPr"}
	 * @see office-js-api/Examples/{Editor}/ApiRichParaPr/Methods/GetClassType.js
	 */
	ApiRichParaPr.prototype.GetClassType = function() {
		return "richParaPr";
	};

	/**
	 * Sets the rich paragraph contents justification.
	 * @memberof ApiRichParaPr
	 * @typeofeditors ["PDFE"]
	 * @param {HorTextAlign} jc - The justification type that will be applied to the rich paragraph contents.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichParaPr/Methods/SetJc.js
	 */
	ApiRichParaPr.prototype.SetJc = AscBuilder.ApiParaPr.prototype.SetJc;

	/**
	 * Returns the rich paragraph contents justification.
	 * @memberof ApiRichParaPr
	 * @typeofeditors ["PDFE"]
	 * @returns {HorTextAlign} 
	 * @see office-js-api/Examples/{Editor}/ApiRichParaPr/Methods/GetJc.js
	 */
	ApiRichParaPr.prototype.GetJc = AscBuilder.ApiParaPr.prototype.GetJc;

	ApiRichParaPr.prototype.private_OnChange = function() {
		if (this.Parent)
			this.Parent.OnChangeParaPr(this);
	};
	
	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiRichParagraph
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a rich paragraph.
	 * @constructor
	 * @extends {ApiRichParaPr}
	 */
	function ApiRichParagraph(Paragraph)
	{
		ApiRichParaPr.call(this, this, Paragraph.Pr.Copy());
		this.Paragraph = Paragraph;
	}
	ApiRichParagraph.prototype = Object.create(ApiRichParaPr.prototype);
	ApiRichParagraph.prototype.constructor = ApiRichParagraph;

	/**
	 * Returns a type of the ApiRichParagraph class.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @returns {"richParagraph"}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/GetClassType.js
	 */
	ApiRichParagraph.prototype.GetClassType = function() {
		return "richParagraph";
	};

	/**
	 * Adds some text to the current paragraph.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @param {string} text - The text that we want to insert into the current paragraph.
	 * @returns {ApiRichRun}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/AddText.js
	 */
	ApiRichParagraph.prototype.AddText = function(text) {
		let oApiRun = AscBuilder.ApiParagraph.prototype.AddText.call(this, text);
		return new ApiRichRun(oApiRun.private_GetImpl());
	};

	/**
	 * Sets the paragraph properties.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @param {ApiRichParaPr} paraPr
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/SetParaPr.js
	 */
	ApiRichParagraph.prototype.SetParaPr = function(paraPr) {
		if (!(paraPr instanceof ApiRichParaPr)) {
			AscBuilder.throwException('The paraPr parameter must be an ApiRichParaPr object');
		}

		this.ParaPr.Merge(paraPr.ParaPr);
		this.OnChangeParaPr(new ApiRichParaPr(this.Paragraph, this.ParaPr));
		return true;
	};

	/**
	 * Returns the paragraph properties.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiRichParaPr}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/GetParaPr.js
	 */
	ApiRichParagraph.prototype.GetParaPr = function() {
		return new ApiRichParaPr(this, this.Paragraph.Pr.Copy());
	};

	/**
	 * Returns a number of elements in the current paragraph.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/GetElementsCount.js
	 */
	ApiRichParagraph.prototype.GetElementsCount = AscBuilder.ApiParagraph.prototype.GetElementsCount;

	/**
	 * Adds an element to the current paragraph.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @param {ApiRichRun} richRun - The element which will be added at the current position.
	 * @param {number} [pos] - The position where the current element will be added. If this value is not
	 * specified, then the element will be added at the end of the current paragraph.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/AddElement.js
	 */
	ApiRichParagraph.prototype.AddElement = function(richRun, pos) {
		if (!(richRun instanceof ApiRichRun)) {
			AscBuilder.throwException("The richRun parameter must be an ApiRichRun object");
		}

		let nMaxPos = this.GetElementsCount();

		pos = AscBuilder.GetNumberParameter(pos, null);
		if (null == pos || pos < 0 || pos > nMaxPos) {
			AscBuilder.throwException("The pos parameter must be a valid position");
		}

		let oParaElement = richRun.private_GetImpl();
		if (oParaElement.IsUseInDocument()) {
			AscBuilder.throwException("The richRun already in the document");
		}

		this.Paragraph.Add_ToContent(pos, oParaElement);
		this.Paragraph.CorrectContent(undefined, undefined, true);

		return true;
	};

	/**
	 * Adds an element to the current paragraph.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @param {ParagraphContent} richRun
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/Push.js
	 */
	ApiRichParagraph.prototype.Push = function(richRun) {
		return this.AddElement(richRun, this.GetElementsCount());
	};

	/**
	 * Returns a paragraph element using the position specified.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @param {number} pos - The position where the element which content we want to get must be located.
	 * @returns {?ApiRichRun}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/GetElement.js
	 */
	ApiRichParagraph.prototype.GetElement = function(pos) {
		if (pos < 0 || pos >= this.Paragraph.Content.length - 1)
			return null;

		return new ApiRichRun(this.Paragraph.Content[pos]);
	};

	/**
	 * Removes an element using the position specified.
	 * <note>If the element you remove is the last paragraph element (i.e. all the elements are removed from the paragraph),
	 * a new empty run is automatically created.</note>
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @param {number} pos - The element position which we want to remove from the paragraph.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/RemoveElement.js
	 */
	ApiRichParagraph.prototype.RemoveElement = AscBuilder.ApiParagraph.prototype.RemoveElement;

	/**
	 * Removes all the elements from the current paragraph.
	 * <note>When all the elements are removed from the paragraph, a new empty run is automatically created.</note>
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/RemoveAllElements.js
	 */
	ApiRichParagraph.prototype.RemoveAllElements = AscBuilder.ApiParagraph.prototype.RemoveAllElements;

	/**
	 * Deletes the current paragraph.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean} - returns false if paragraph haven't parent.
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/Delete.js
	 */
	ApiRichParagraph.prototype.Delete = AscBuilder.ApiParagraph.prototype.Delete;

	/**
	 * Returns the next paragraph.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiRichParagraph} - returns null if paragraph is last.
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/GetNext.js
	 */
	ApiRichParagraph.prototype.GetNext = function() {
		let nextPara = this.Paragraph.GetNextParagraph();
		if (nextPara !== null)
			return new ApiRichParagraph(nextPara);

		return null;
	};

	/**
	 * Returns the previous paragraph.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiRichParagraph} - returns null if paragraph is first.
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/GetPrevious.js
	 */
	ApiRichParagraph.prototype.GetPrevious = function() {
		let prevPara = this.Paragraph.GetPrevParagraph();
		if (prevPara !== null)
			return new ApiRichParagraph(prevPara);

		return null;
	};

	/**
	 * Returns the last element of the paragraph.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiRichRun}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/Last.js
	 */
	ApiRichParagraph.prototype.Last = function() {
		let oApiRun = AscBuilder.ApiParagraph.prototype.Last.call(this);
		return new ApiRichRun(oApiRun.private_GetImpl());
	};

	/**
	 * Returns the last Run with text in the current paragraph.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiRichRun}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/GetLastRunWithText.js
	 */
	ApiRichParagraph.prototype.GetLastRunWithText = AscBuilder.ApiParagraph.prototype.GetLastRunWithText;

	/**
	 * Creates a paragraph copy.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiRichParagraph}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/Copy.js
	 */
	ApiRichParagraph.prototype.Copy = function() {
		let oParagraph = this.Paragraph.Copy(undefined, private_GetDrawingDocument(), {
			SkipComments          : true,
			SkipFootnoteReference : true,
			SkipComplexFields     : true
		});

		return new ApiRichParagraph(oParagraph);
	};
	
	/**
	 * Specifies the reading order for the current paragraph.
	 * Possible values are:
	 * <b>null</b> - use the standart direction parameter;
	 * <b>"ltr"</b> - left-to-right text direction;
	 * <b>"rtl"</b> - right-to-left text direction.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @param {?ReadingOrder} [readingOrder = undefined] - The reading order.
	 * @returns {ApiRichParagraph} - Returns the current paragraph itself (ApiRichParagraph).
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/SetReadingOrder.js
	 */
	ApiRichParagraph.prototype.SetReadingOrder = function(readingOrder) {
		let oApiPara = AscBuilder.ApiParagraph.prototype.SetReadingOrder.call(this, readingOrder);
		return new ApiRichParagraph(oApiPara.private_GetImpl());
	};

	/**
	 * Returns the paragraph text.
	 * @memberof ApiRichParagraph
	 * @param {object} [options] - Options for formatting the returned text.
	 * @param {string} [options.NewLineSeparator='\r'] - Defines how the line separator will be specified in the resulting string. Any string can be used. The default separator is "\r".
	 * @param {string} [options.TabSymbol='\t'] - Defines how the tab will be specified in the resulting string (does not apply to numbering). Any string can be used. The default symbol is "\t".
	 * @typeofeditors ["PDFE"]
	 * @return {string}
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/GetText.js
	 */
	ApiRichParagraph.prototype.GetText = AscBuilder.ApiParagraph.prototype.GetText;

	/**
	 * Returns the paragraph position within its parent element.
	 * @memberof ApiRichParagraph
	 * @typeofeditors ["PDFE"]
	 * @returns {Number} - returns -1 if the paragraph parent doesn't exist. 
	 * @see office-js-api/Examples/{Editor}/ApiRichParagraph/Methods/GetPosInParent.js
	 */
	ApiRichParagraph.prototype.GetPosInParent = AscBuilder.ApiParagraph.prototype.GetPosInParent;

	ApiRichParagraph.prototype.OnChangeParaPr = function(oApiParaPr) {
		AscBuilder.ApiParagraph.prototype.OnChangeParaPr.call(this, oApiParaPr);
		this.Paragraph.OnContentChange();
	};

	ApiRichParagraph.prototype.private_GetImpl = AscBuilder.ApiParagraph.prototype.private_GetImpl;

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiRichTextPr
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing the rich text properties.
	 * @constructor
	 */
	function ApiRichTextPr(Parent, TextPr) {
		this.Parent = Parent;
		this.TextPr = TextPr;
	}

	/**
	 * Returns a type of the ApiRichTextPr class.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @returns {"textPr"}
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/GetClassType.js
	 */
	ApiRichTextPr.prototype.GetClassType = function() {
		return "richTextPr";
	};

	/**
	 * Sets the bold property to the text character.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} isBold - Specifies that the contents of the run are displayed bold.
	 * @return {ApiRichTextPr} - this text properties.
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/SetBold.js
	 */
	ApiRichTextPr.prototype.SetBold = AscBuilder.ApiTextPr.prototype.SetBold;

	/**
	 * Gets the bold property from the current text properties.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @return {?boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/GetBold.js
	 */
	ApiRichTextPr.prototype.GetBold = AscBuilder.ApiTextPr.prototype.GetBold;

	/**
	 * Sets the italic property to the text character.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} isItalic - Specifies that the contents of the current run are displayed italicized.
	 * @return {ApiRichTextPr} - this text properties.
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/SetItalic.js
	 */
	ApiRichTextPr.prototype.SetItalic = AscBuilder.ApiTextPr.prototype.SetItalic;

	/**
	 * Gets the italic property from the current text properties.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @return {?boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/GetItalic.js
	 */
	ApiRichTextPr.prototype.GetItalic = AscBuilder.ApiTextPr.prototype.GetItalic;

	/**
	 * Specifies that the contents of the run are displayed with a single horizontal line through the center of the line.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} isStrikeout - Specifies that the contents of the current run are displayed struck through.
	 * @return {ApiRichTextPr} - this text properties.
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/SetStrikeout.js
	 */
	ApiRichTextPr.prototype.SetStrikeout = AscBuilder.ApiTextPr.prototype.SetStrikeout;

	/**
	 * Gets the strikeout property from the current text properties.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @return {?boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/GetStrikeout.js
	 */
	ApiRichTextPr.prototype.GetStrikeout = AscBuilder.ApiTextPr.prototype.GetStrikeout;

	/**
	 * Specifies that the contents of the run are displayed along with a line appearing directly below the character
	 * (less than all the spacing above and below the characters on the line).
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} isUnderline - Specifies that the contents of the current run are displayed underlined.
	 * @return {ApiRichTextPr} - this text properties.
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/SetUnderline.js
	 */
	ApiRichTextPr.prototype.SetUnderline = AscBuilder.ApiTextPr.prototype.SetUnderline;

	/**
	 * Gets the underline property from the current text properties.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @return {?boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/GetUnderline.js
	 */
	ApiRichTextPr.prototype.GetUnderline = AscBuilder.ApiTextPr.prototype.GetUnderline;

	/**
	 * Sets all 4 font slots with the specified font family.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @param {string} fontFamily - The font family or families used for the current text run.
	 * @return {ApiRichTextPr} - this text properties.
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/SetFontFamily.js
	 */
	ApiRichTextPr.prototype.SetFontFamily = AscBuilder.ApiTextPr.prototype.SetFontFamily;

	/**
	 * Returns the font family from the current text properties.
	 * The method automatically calculates the font from the theme if the font was set via the theme.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * param {undefined | "ascii" | "eastAsia" | "hAnsi" | "cs"} [fontSlot="ascii"] - The font slot.
	 * If this parameter is not specified, the "ascii" value is used.
	 * @return {?string}
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/GetFontFamily.js
	 */
	ApiRichTextPr.prototype.GetFontFamily = AscBuilder.ApiTextPr.prototype.GetFontFamily;

	/**
	 * Sets the font size to the characters of the current text run.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @param {hps} nSize - The text size value measured in half-points (1/144 of an inch).
	 * @return {ApiRichTextPr} - this text properties.
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/SetFontSize.js
	 */
	ApiRichTextPr.prototype.SetFontSize = AscBuilder.ApiTextPr.prototype.SetFontSize;

	/**
	 * Gets the font size from the current text properties.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @return {?hps}
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/GetFontSize.js
	 */
	ApiRichTextPr.prototype.GetFontSize = AscBuilder.ApiTextPr.prototype.GetFontSize;

	/**
	 * Sets the text color to the current text run.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @param {ApiColor} color
	 * @return {ApiRichTextPr} - this text properties.
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/SetColor.js
	 */
	ApiRichTextPr.prototype.SetColor = AscBuilder.ApiTextPr.prototype.SetColor;

	/**
	 * Gets the RGB color from the current text properties.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @return {?ApiColor}
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/GetColor.js
	 */
	ApiRichTextPr.prototype.GetColor = AscBuilder.ApiTextPr.prototype.GetColor;

	/**
	 * Specifies the alignment which will be applied to the contents of the run in relation to the default appearance of the run text:
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @param {TextVertAlign} type - The vertical alignment type applied to the text contents.
	 * @return {ApiRichTextPr} - this text properties.
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/SetVertAlign.js
	 */
	ApiRichTextPr.prototype.SetVertAlign = AscBuilder.ApiTextPr.prototype.SetVertAlign;

	/**
	 * Gets the vertical alignment type from the current text properties.
	 * @memberof ApiRichTextPr
	 * @typeofeditors ["PDFE"]
	 * @return {TextVertAlign}
	 * @see office-js-api/Examples/{Editor}/ApiRichTextPr/Methods/GetVertAlign.js
	 */
	ApiRichTextPr.prototype.GetVertAlign = AscBuilder.ApiTextPr.prototype.GetVertAlign;

	ApiRichTextPr.prototype.private_OnChange = AscBuilder.ApiTextPr.prototype.private_OnChange;

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiRichRun
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a small text block called 'run'.
	 * @constructor
	 * @extends {ApiRichTextPr}
	 */
	function ApiRichRun(Run) {
		ApiRichTextPr.call(this, this, Run.Pr.Copy());
		this.Run = Run;
	}
	ApiRichRun.prototype = Object.create(ApiRichTextPr.prototype);
	ApiRichRun.prototype.constructor = ApiRichRun;

	/**
	 * Returns a type of the ApiRichRun class.
	 * @memberof ApiRichRun
	 * @typeofeditors ["PDFE"]
	 * @returns {"richRun"}
	 * @see office-js-api/Examples/{Editor}/ApiRichRun/Methods/GetClassType.js
	 */
	ApiRichRun.prototype.GetClassType = function() {
		return "richRun";
	};

	/**
	 * Sets the text properties to the current run.
	 * @memberof ApiRichRun
	 * @typeofeditors ["PDFE"]
	 * @param {ApiRichTextPr} textPr - The text properties that will be set to the current run.
	 * @return {ApiRichTextPr}  
	 * @see office-js-api/Examples/{Editor}/ApiRichRun/Methods/SetTextPr.js
	 */
	ApiRichRun.prototype.SetTextPr = function(textPr) {
		if (!(textPr instanceof ApiRichTextPr)) {
			AscBuilder.throwException('The textPr parameter must be an ApiRichTextPr object');
		}

		let runTextPr = this.GetTextPr();
		runTextPr.TextPr.Merge(textPr.TextPr);
		runTextPr.private_OnChange();

		return runTextPr;
	};

	/**
	 * Returns the text properties of the current run.
	 * @memberof ApiRichRun
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiRichTextPr}
	 * @see office-js-api/Examples/{Editor}/ApiRichRun/Methods/GetTextPr.js
	 */
	ApiRichRun.prototype.GetTextPr = function() {
		return new ApiRichTextPr(this, this.TextPr);
	};

	/**
	 * Clears the content from the current run.
	 * @memberof ApiRichRun
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichRun/Methods/ClearContent.js
	 */
	ApiRichRun.prototype.ClearContent = AscBuilder.ApiRun.prototype.ClearContent;

	/**
	 * Removes all the elements from the current run.
	 * @memberof ApiRichRun
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichRun/Methods/RemoveAllElements.js
	 */
	ApiRichRun.prototype.RemoveAllElements = AscBuilder.ApiRun.prototype.RemoveAllElements;

	/**
	 * Deletes the current run.
	 * @memberof ApiRichRun
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichRun/Methods/Delete.js
	 */
	ApiRichRun.prototype.Delete = AscBuilder.ApiRun.prototype.Delete;

	/**
	 * Adds some text to the current run.
	 * @memberof ApiRichRun
	 * @typeofeditors ["PDFE"]
	 * @param {string} text - The text which will be added to the current run.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRichRun/Methods/AddText.js
	 */
	ApiRichRun.prototype.AddText = AscBuilder.ApiRun.prototype.AddText;
	
	/**
	 * Creates a copy of the current run.
	 * @memberof ApiRichRun
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiRichRun}
	 * @see office-js-api/Examples/{Editor}/ApiRichRun/Methods/Copy.js
	 */
	ApiRichRun.prototype.Copy = function() {
		let oRun = this.Run.Copy(false, {
			SkipComments          : true,
			SkipAnchors           : true,
			SkipFootnoteReference : true,
			SkipComplexFields     : true
		});

		return new ApiRichRun(oRun);
	};

	/**
	 * Returns a text from the text run.
	 * @memberof ApiRichRun
	 * @typeofeditors ["PDFE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiRichRun/Methods/GetText.js
	 */
	ApiRichRun.prototype.GetText = AscBuilder.ApiRun.prototype.GetText;

	/**
	 * Returns a parent paragraph of the current run.
	 * @memberof ApiRichRun
	 * @typeofeditors ["PDFE"]
	 * @return {?ApiRichParagraph}
	 * @see office-js-api/Examples/{Editor}/ApiRichRun/Methods/GetParentParagraph.js
	 */
	ApiRichRun.prototype.GetParentParagraph = function() {
		let oPara = this.Run.GetParagraph();
		if (!oPara)
			return null;

		return new ApiRichParagraph(oPara); 
	};

	ApiRichRun.prototype.OnChangeTextPr = function(oApiTextPr) {
		AscBuilder.ApiRun.prototype.OnChangeTextPr.call(this, oApiTextPr);

		let oPara = this.Run.GetParagraph();
		if (oPara) {
			oPara.OnContentChange();
		}
	};

	ApiRichRun.prototype.private_GetImpl = AscBuilder.ApiRun.prototype.private_GetImpl;
	
	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiDrawing
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a graphical object.
	 * @constructor
	 */
	function ApiDrawing(Drawing) {
		this.Drawing = Drawing;
	}

	ApiDrawing.prototype.private_GetImpl = function() {
		return this.Drawing;
	};

	/**
	 * Returns the type of the ApiDrawing class.
	 * @typeofeditors ["PDFE"]
	 * @returns {"drawing"}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/GetClassType.js
	 */
	ApiDrawing.prototype.GetClassType = function() {
		return "drawing";
	};

	/**
	 * Returns the type of the ApiDrawing class.
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiPage}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/GetClassType.js
	 */
	ApiDrawing.prototype.GetParentPage = function() {
		let oParentPage = this.Drawing.GetParentPage();
		if (!oParentPage) {
			return null;
		}

		return new ApiPage(oParentPage);
	};

	/**
	 * Sets the size of the object (image, shape, chart) bounding box.
	 * @typeofeditors ["PDFE"]
	 * @param {EMU} width - The object width measured in English measure units.
	 * @param {EMU} height - The object height measured in English measure units.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/SetSize.js
	 */
	ApiDrawing.prototype.SetSize = function(width, height) {
		let fWidth = private_EMU2MM(width);
		let fHeight = private_EMU2MM(height);

		this.Drawing.checkTransformBeforeApply();
		let xfrm = this.Drawing.getXfrm();
		if (xfrm) {
			xfrm.setExtX(fWidth);
			xfrm.setExtY(fHeight);
		}

		return true;
	};

	/**
	 * Sets the position of the drawing on the page.
	 * @typeofeditors ["PDFE"]
	 * @param {EMU} posX - The distance from the left side of the page to the left side of the drawing measured in English measure units.
	 * @param {EMU} posY - The distance from the top side of the page to the upper side of the drawing measured in English measure units.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/SetPosition.js
	 */
	ApiDrawing.prototype.SetPosition = function(posX, posY) {
		let fPosX = private_EMU2MM(posX);
		let fPosY = private_EMU2MM(posY);

		if (this.Drawing && this.Drawing.spPr && this.Drawing.spPr.xfrm) {
			this.Drawing.spPr.xfrm.setOffX(fPosX);
			this.Drawing.spPr.xfrm.setOffY(fPosY);
		}

		return true;
	};

	/**
	 * Sets the x position of the drawing on the page.
	 * @typeofeditors ["PDFE"]
	 * @param {EMU} posX - The distance from the left side of the page to the left side of the drawing measured in English measure units.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/SetPosX.js
	 */
	ApiDrawing.prototype.SetPosX = function(posX) {
        this.Drawing.checkTransformBeforeApply();
		let oXfrm = this.Drawing.getXfrm();
		oXfrm.setOffX(private_EMU2MM(posX));

		return true;
	};

	/**
	 * Gets the x position of the drawing on the page.
	 * @typeofeditors ["PDFE"]
	 * @returns {EMU}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/GetPosX.js
	 */
	ApiDrawing.prototype.GetPosX = function() {
		return private_MM2EMU(this.Drawing.GetPosX());
	};

	/**
	 * Sets the y position of the drawing on the page.
	 * @typeofeditors ["PDFE"]
	 * @param {EMU} posY - The distance from the top side of the page to the upper side of the drawing measured in English measure units.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/SetPosY.js
	 */
	ApiDrawing.prototype.SetPosY = function(posY) {
        this.Drawing.checkTransformBeforeApply();
		let oXfrm = this.Drawing.getXfrm();
		oXfrm.setOffY(private_EMU2MM(posY));

		return true;
	};

	/**
	 * Gets the y position of the drawing on the page.
	 * @typeofeditors ["PDFE"]
	 * @returns {EMU}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/GetPosY.js
	 */
	ApiDrawing.prototype.GetPosY = function() {
		return private_MM2EMU(this.Drawing.GetPosY());
	};

	/**
	 * Creates a copy of the specified drawing object.
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiDrawing} - return null if drawing doesn't exist.
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/Copy.js
	 */
	ApiDrawing.prototype.Copy = function() {
		return new ApiDrawing(this.Drawing.copy());
	};

	/**
	 * Deletes the specified drawing object from the parent.
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean} - false if drawing doesn't exist or drawing hasn't a parent.
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/Delete.js
	 */
	ApiDrawing.prototype.Delete = function() {
		let oDoc = private_GetLogicDocument();
		oDoc.RemoveDrawing(this.Drawing.GetId());
		return true;
	};

	/**
	 * Returns the width of the current drawing.
	 * @memberof ApiDrawing
	 * @typeofeditors ["PDFE"]
	 * @returns {EMU}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/GetWidth.js
	 */
	ApiDrawing.prototype.GetWidth = function() {
		return private_MM2EMU(this.Drawing.GetWidth());
	};

	/**
	 * Returns the height of the current drawing.
	 * @memberof ApiDrawing
	 * @typeofeditors ["PDFE"]
	 * @returns {EMU}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/GetHeight.js
	 */
	ApiDrawing.prototype.GetHeight = function() {
		return private_MM2EMU(this.Drawing.GetHeight());
	};

	/**
	 * Selects the current graphic object.
	 * @memberof ApiDrawing
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/Select.js
	 */
	ApiDrawing.prototype.Select = function() {
		this.Drawing.Set_CurrentElement(true, 0, true);

		let oController = this.Drawing.getDrawingObjectsController();
        oController.updateSelectionState();
        oController.updateOverlay();

		return true;
	};

	/**
	 * Sets the rotation angle to the current drawing object.
	 * @memberof ApiDrawing
	 * @param {number} rotAngle - New drawing rotation angle.
	 * @typeofeditors ["PDFE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/SetRotation.js
	 */
	ApiDrawing.prototype.SetRotation = function(rotAngle) {
		if (!this.Drawing.canRotate()) {
			return false;
		}

		this.Drawing.checkTransformBeforeApply();
		let oXfrm = this.Drawing.getXfrm();
		oXfrm.setRot(rotAngle * Math.PI / 180);

		return true;
	};

	/**
	 * Returns the rotation angle of the current drawing object.
	 * @memberof ApiDrawing
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/GetRotation.js
	 */
	ApiDrawing.prototype.GetRotation = function() {
		this.Drawing.checkRecalculateTransform();
		return this.Drawing.rot * 180 / Math.PI
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiShape
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a shape.
	 * @constructor
	 */
	function ApiShape(oShape) {
		ApiDrawing.call(this, oShape);
		this.Shape = oShape;
	}
	ApiShape.prototype = Object.create(ApiDrawing.prototype);
	ApiShape.prototype.constructor = ApiShape;

	/**
	 * Returns the type of the ApiShape class.
	 * @typeofeditors ["PDFE"]
	 * @returns {"shape"}
	 * @see office-js-api/Examples/{Editor}/ApiShape/Methods/GetClassType.js
	 */
	ApiShape.prototype.GetClassType = function() {
		return "shape";
	};

	/**
	 * Returns the shape inner contents where a paragraph or text runs can be inserted.
	 * @typeofeditors ["PDFE"]
	 * @returns {?ApiDocumentContent}
	 * @see office-js-api/Examples/{Editor}/ApiShape/Methods/GetContent.js
	 */
	ApiShape.prototype.GetContent = function() {
		let docContent = this.Drawing.getDocContent();
		if (!docContent) {
			this.Drawing.createTextBody();
		}

		docContent = this.Drawing.getDocContent();
		if (docContent) {
			return Asc.editor.private_CreateApiDocContent(docContent);
		}

		return null;
	};

	/**
	 * Sets the vertical alignment to the shape content where a paragraph or text runs can be inserted.
	 * @typeofeditors ["PDFE"]
	 * @param {VerticalTextAlign} verticalAlign - The type of the vertical alignment for the shape inner contents.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiShape/Methods/SetVerticalTextAlign.js
	 */
	ApiShape.prototype.SetVerticalTextAlign = function(verticalAlign) {
		switch (verticalAlign) {
			case "top": {
				this.Shape.setVerticalAlign(4);
				break;
			}
			case "center": {
				this.Shape.setVerticalAlign(1);
				break;
			}
			case "bottom": {
				this.Shape.setVerticalAlign(0);
				break;
			}
		}

		return true;
	};

	/**
	 * Returns the geometry object from the current shape.
	 * @memberof ApiShape
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiGeometry}
	 * @see office-js-api/Examples/{Editor}/ApiShape/Methods/GetGeometry.js
	 */
	ApiShape.prototype.GetGeometry = function() {
		if (this.Shape && this.Shape.spPr && this.Shape.spPr.geometry) {
			return Api.private_CreateGeometry(this.Shape.spPr.geometry);
		}

		return null;
	};

	/**
	 * Sets a custom geometry for the current shape.
	 * @memberof ApiShape
	 * @typeofeditors ["PDFE"]
	 * @param {ApiGeometry} oGeometry - The geometry to set.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiShape/Methods/SetGeometry.js
	 */
	ApiShape.prototype.SetGeometry = function(oGeometry) {
		if (this.Shape && this.Shape.spPr && oGeometry && oGeometry.geometry) {
			this.Shape.spPr.setGeometry(oGeometry.geometry);
			return true;
		}

		return false;
	};

	/**
	 * Sets the fill properties to the current shape.
	 * @memberof ApiShape
	 * @typeofeditors ["PDFE"]
	 * @param {ApiFill} oFill - The fill type used to fill the shape.
	 * @returns {boolean} - returns false if param is invalid.
	 * @see office-js-api/Examples/{Editor}/ApiShape/Methods/SetFill.js
	 */
	ApiShape.prototype.SetFill = function(oFill) {
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		if (this.Shape && this.Shape.spPr) {
			this.Shape.spPr.setFill(oFill.UniFill);
			return true;
		}

		return false;
	};

	/**
	 * Gets the fill properties from the current shape.
	 * @memberof ApiShape
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiFill | null}
	 * @see office-js-api/Examples/{Editor}/ApiShape/Methods/GetFill.js
	 */
	ApiShape.prototype.GetFill = function() {
		if (this.Shape) {
			if (this.Shape.recalcInfo && this.Shape.recalcInfo.recalculateBrush) {
				this.Shape.recalculateBrush();
			}
			if (this.Shape.brush) {
				return new AscBuilder.ApiFill(this.Shape.brush);
			}
		}

		return null;
	};

	/**
	 * Sets the outline properties to the current shape.
	 * @memberof ApiShape
	 * @typeofeditors ["PDFE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the shape outline.
	 * @returns {boolean} - returns false if param is invalid.
	 * @see office-js-api/Examples/{Editor}/ApiShape/Methods/SetLine.js
	 */
	ApiShape.prototype.SetLine = function(oStroke) {
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		if (this.Shape && this.Shape.spPr) {
			this.Shape.spPr.setLn(oStroke.Ln);
			return true;
		}

		return false;
	};

	/**
	 * Gets the outline properties from the current shape.
	 * @memberof ApiShape
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiStroke | null}
	 * @see office-js-api/Examples/{Editor}/ApiShape/Methods/GetLine.js
	 */
	ApiShape.prototype.GetLine = function() {
		if (this.Shape) {
			if (this.Shape.recalcInfo && this.Shape.recalcInfo.recalculatePen) {
				this.Shape.recalculatePen();
			}
			if (this.Shape.pen) {
				return new AscBuilder.ApiStroke(this.Shape.pen);
			}
		}

		return null;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiImage
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing an image.
	 * @constructor
	 */
	function ApiImage(oImage) {
		ApiDrawing.call(this, oImage);
	}
	ApiImage.prototype = Object.create(ApiDrawing.prototype);
	ApiImage.prototype.constructor = ApiImage;

	/**
	 * Returns the type of the ApiImage class.
	 * @typeofeditors ["PDFE"]
	 * @returns {"image"}
	 * @see office-js-api/Examples/{Editor}/ApiImage/Methods/GetClassType.js
	 */
	ApiImage.prototype.GetClassType = function() {
		return "image";
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTable
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a table.
	 * @param oGraphicFrame
	 * @constructor
	 */
	function ApiTable(oGraphicFrame) {
		this.Table = oGraphicFrame.graphicObject;
		ApiDrawing.call(this, oGraphicFrame);
	}

	ApiTable.prototype = Object.create(ApiDrawing.prototype);
	ApiTable.prototype.constructor = ApiTable;

	/**
	 * Returns the type of the ApiTable object.
	 * @typeofeditors ["PDFE"]
	 * @returns {"table"}
	 * @see office-js-api/Examples/{Editor}/ApiTable/Methods/GetClassType.js
	 */
	ApiTable.prototype.GetClassType = function() {
		return "table";
	};

	/**
	 * Adds a new row to the current table.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiTableCell} [oCell] - If not specified, a new row will be added to the end of the table.
	 * @param {boolean} [isBefore=false] - Adds a new row before or after the specified cell. If no cell is specified,
	 * then this parameter will be ignored.
	 * @returns {ApiTableRow}
	 * @see office-js-api/Examples/{Editor}/ApiTable/Methods/AddRow.js
	 */
	ApiTable.prototype.AddRow = function(oCell, isBefore) {
		this.private_PrepareTableForActions();

		let _isBefore = AscBuilder.GetBoolParameter(isBefore, false);
		let _oCell = (oCell instanceof ApiTableCell ? oCell.Cell : undefined);
		if (_oCell && this.Table !== _oCell.Row.Table)
			_oCell = undefined;

		if (!_oCell) {
			_oCell = this.Table.Content[this.Table.Content.length - 1].Get_Cell(0);
			_isBefore = false;
		}

		let nRowIndex = true === _isBefore ? _oCell.Row.Index : _oCell.Row.Index + 1;

		this.Table.RemoveSelection();
		this.Table.CurCell = _oCell;
		this.Table.AddTableRow(_isBefore);

		return new ApiTableRow(this.Table.Content[nRowIndex]);
	};

	/**
	 * Returns a row by its index.
	 * @typeofeditors ["PDFE"]
	 * @param nIndex {number} - The row index (position) in the table.
	 * @returns {?ApiTableRow}
	 * @see office-js-api/Examples/{Editor}/ApiTable/Methods/GetRow.js
	 */
	ApiTable.prototype.GetRow = function(nIndex) {
		if (!this.Drawing) {
			return null;
		}

		let aTableContent = this.Table.Content;
		if (!aTableContent[nIndex]) {
			return null;
		}

		return new ApiTableRow(aTableContent[nIndex]);
	};

	/**
	 * Removes a table row with the specified cell.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiTableCell} oCell - The table cell from the row which will be removed.
	 * @returns {boolean} - defines if the table is empty after removing or not.
	 * @see office-js-api/Examples/{Editor}/ApiTable/Methods/RemoveRow.js
	 */
	ApiTable.prototype.RemoveRow = function(oCell) {
		if (!(oCell instanceof ApiTableCell) || this.Table !== oCell.Cell.Row.Table)
			return false;

		this.private_PrepareTableForActions();
		this.Table.RemoveSelection();
		this.Table.CurCell = oCell.Cell;
		return !(this.Table.RemoveTableRow());
	};

	/**
	 * Merges an array of cells. If merge is successful, it will return merged cell, otherwise "null".
	 * <b>Warning</b>: The number of cells in any row and the number of rows in the current table may be changed.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiTableCell[]} aCells - The array of cells.
	 * @returns {?ApiTableCell}
	 * @see office-js-api/Examples/{Editor}/ApiTable/Methods/MergeCells.js
	 */
	ApiTable.prototype.MergeCells = function(aCells) {
		this.private_PrepareTableForActions();

		let oTable = this.Table;
		oTable.Selection.Use = true;
		oTable.Selection.Type = table_Selection_Cell;
		oTable.Selection.Data = [];

		for (let nPos = 0, nCount = aCells.length; nPos < nCount; ++nPos) {
			let oCell = aCells[nPos].Cell;
			let oPos = {
				Cell: oCell.Index,
				Row: oCell.Row.Index
			};

			let nResultPos = 0;
			let nResultLength = oTable.Selection.Data.length;
			for (nResultPos = 0; nResultPos < nResultLength; ++nResultPos) {
				let oCurPos = oTable.Selection.Data[nResultPos];
				if (oCurPos.Row < oPos.Row) {
					continue;
				} else if (oCurPos.Row > oPos.Row) {
					break;
				} else {
					if (oCurPos.Cell >= oPos.Cell)
						break;
				}
			}

			oTable.Selection.Data.splice(nResultPos, 0, oPos);
		}

		let isMerged = this.Table.MergeTableCells(true);
		let oMergedCell = this.Table.CurCell;
		oTable.RemoveSelection();


		if (true === isMerged)
			return new ApiTableCell(oMergedCell);

		return null;
	};

	/**
	 * Specifies the components of the conditional formatting of the referenced table style (if one exists)
	 * which shall be applied to the set of table rows with the current table-level property exceptions. A table style
	 * can specify up to six different optional conditional formats [Example: Different formatting for first column],
	 * which then can be applied or omitted from individual table rows in the parent table.
	 *
	 * The default setting is to apply the row and column banding formatting, but not the first row, last row, first
	 * column, or last column formatting.
	 * @typeofeditors ["PDFE"]
	 * @param {boolean} isFirstColumn - Specifies that the first column conditional formatting shall be applied to the
	 *     table.
	 * @param {boolean} isFirstRow - Specifies that the first row conditional formatting shall be applied to the table.
	 * @param {boolean} isLastColumn - Specifies that the last column conditional formatting shall be applied to the
	 *     table.
	 * @param {boolean} isLastRow - Specifies that the last row conditional formatting shall be applied to the table.
	 * @param {boolean} isHorBand - Specifies that the horizontal banding conditional formatting shall not be applied
	 *     to the table.
	 * @param {boolean} isVerBand - Specifies that the vertical banding conditional formatting shall not be applied to
	 *     the table.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTable/Methods/SetTableLook.js
	 */
	ApiTable.prototype.SetTableLook = function(isFirstColumn, isFirstRow, isLastColumn, isLastRow, isHorBand, isVerBand) {
		let oTableLook = new AscCommon.CTableLook(AscBuilder.GetBoolParameter(isFirstColumn),
			AscBuilder.GetBoolParameter(isFirstRow),
			AscBuilder.GetBoolParameter(isLastColumn),
			AscBuilder.GetBoolParameter(isLastRow),
			AscBuilder.GetBoolParameter(isHorBand),
			AscBuilder.GetBoolParameter(isVerBand));

		this.Table.Set_TableLook(oTableLook);

		return true;
	};

	/**
	 * Adds a new column to the end of the current table.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiTableCell} [oCell] - If not specified, a new column will be added to the end of the table.
	 * @param {boolean} [isBefore=false] - Add a new column before or after the specified cell. If no cell is specified,
	 * then this parameter will be ignored.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTable/Methods/AddColumn.js
	 */
	ApiTable.prototype.AddColumn = function(oCell, isBefore) {
		this.private_PrepareTableForActions();

		let _isBefore = AscBuilder.GetBoolParameter(isBefore, false);
		let _oCell = (oCell instanceof ApiTableCell ? oCell.Cell : undefined);
		if (_oCell && this.Table !== _oCell.Row.Table)
			_oCell = undefined;

		if (!_oCell) {
			_oCell = this.Table.Content[0].Get_Cell(this.Table.Content[0].Get_CellsCount() - 1);
			_isBefore = false;
		}

		this.Table.RemoveSelection();
		this.Table.CurCell = _oCell;
		this.Table.AddTableColumn(_isBefore);

		return true;
	};

	/**
	 * Removes a table column with the specified cell.
	 * @typeofeditors ["PDFE"]
	 * @param {ApiTableCell} oCell - The table cell from the column which will be removed.
	 * @returns {boolean} - defines if the table is empty after removing or not.
	 * @see office-js-api/Examples/{Editor}/ApiTable/Methods/RemoveColumn.js
	 */
	ApiTable.prototype.RemoveColumn = function(oCell) {
		if (!(oCell instanceof ApiTableCell) || this.Table !== oCell.Cell.Row.Table)
			return false;
		this.private_PrepareTableForActions();
		this.Table.RemoveSelection();
		this.Table.CurCell = oCell.Cell;
		return !(this.Table.RemoveTableColumn());
	};

	/**
	 * Specifies the shading which shall be applied to the extents of the current table.
	 * @typeofeditors ["PDFE"]
	 * @param {ShdType | ApiFill} sType - The shading type applied to the contents of the current table. Can be ShdType or ApiFill.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTable/Methods/SetShd.js
	 */
	ApiTable.prototype.SetShd = function(sType, r, g, b) {
		let oPr = this.Table.Pr.Copy();
		let color = new Asc.asc_CColor({
			r: r,
			g: g,
			b: b,
			Auto: false
		});
		let oShd = new CDocumentShd();
		let _Shd = null;

		if (sType === "nil") {
			_Shd = {
				Value: Asc.c_oAscShdNil
			};
			oShd.Set_FromObject(_Shd);
			oPr.Shd = oShd;
		} else if (sType === "clear") {

			let Unifill = new AscFormat.CUniFill();
			Unifill.fill = new AscFormat.CSolidFill();
			Unifill.fill.color = AscFormat.CorrectUniColor(color, Unifill.fill.color, 1);
			_Shd = {
				Value: Asc.c_oAscShdClear,
				Color: {
					r: color.asc_getR(),
					g: color.asc_getG(),
					b: color.asc_getB()
				},
				Unifill: Unifill
			};

			oShd.Set_FromObject(_Shd);
			oPr.Shd = oShd;
		} else if (sType.GetClassType && sType.GetClassType() === "fill") {
			oShd.Value = Asc.c_oAscShdClear;
			oShd.Unifill = sType.UniFill;
			oPr.Shd = oShd;
		} else
			oPr.Shd = null;

		this.Table.Set_Pr(oPr);

		return true;
	};

	ApiTable.prototype.private_PrepareTableForActions = function() {
		this.Table.private_RecalculateGrid();
		this.Table.private_UpdateCellsGrid();
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTableRow
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a table row.
	 * @param oTableRow
	 * @constructor
	 */

	function ApiTableRow(oTableRow) {
		this.Row = oTableRow;
	}

	/**
	 * Returns the type of the ApiTableRow class.
	 * @typeofeditors ["PDFE"]
	 * @returns {"tableRow"}
	 * @see office-js-api/Examples/{Editor}/ApiTableRow/Methods/GetClassType.js
	 */
	ApiTableRow.prototype.GetClassType = function() {
		return "tableRow";
	};

	/**
	 * Returns a number of cells in the current row.
	 * @typeofeditors ["PDFE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiTableRow/Methods/GetCellsCount.js
	 */
	ApiTableRow.prototype.GetCellsCount = function() {
		return this.Row.Content.length;
	};

	/**
	 * Returns a cell by its position in the current row.
	 * @typeofeditors ["PDFE"]
	 * @param {number} nPos - The cell position in the table row.
	 * @returns {ApiTableCell}
	 * @see office-js-api/Examples/{Editor}/ApiTableRow/Methods/GetCell.js
	 */
	ApiTableRow.prototype.GetCell = function(nPos) {
		if (nPos < 0 || nPos >= this.Row.Content.length)
			return null;

		return new ApiTableCell(this.Row.Content[nPos]);
	};

	/**
	 * Sets the height to the current table row.
	 * @typeofeditors ["PDFE"]
	 * @param {EMU} [nValue] - The row height in English measure units.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTableRow/Methods/SetHeight.js
	 */
	ApiTableRow.prototype.SetHeight = function(nValue) {
		let fMaxTopMargin = 0,
			fMaxBottomMargin = 0,
			fMaxTopBorder = 0,
			fMaxBottomBorder = 0;

		for (let i = 0; i < this.Row.Content.length; ++i) {
			let oCell = this.Row.Content[i];

			let oMargins = oCell.GetMargins();
			if (oMargins.Bottom.W > fMaxBottomMargin) {
				fMaxBottomMargin = oMargins.Bottom.W;
			}
			if (oMargins.Top.W > fMaxTopMargin) {
				fMaxTopMargin = oMargins.Top.W;
			}

			let oBorders = oCell.Get_Borders();
			if (oBorders.Top.Size > fMaxTopBorder) {
				fMaxTopBorder = oBorders.Top.Size;
			}
			if (oBorders.Bottom.Size > fMaxBottomBorder) {
				fMaxBottomBorder = oBorders.Bottom.Size;
			}
		}

		this.Row.Set_Height(Math.max(1, private_EMU2MM(nValue) - fMaxTopMargin - fMaxBottomMargin - fMaxTopBorder / 2 - fMaxBottomBorder / 2), Asc.linerule_AtLeast);
		return true;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTableCell
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a table cell.
	 * @param oCell
	 * @constructor
	 */
	function ApiTableCell(oCell) {
		this.Cell = oCell;
	}

	/**
	 * Returns the type of the ApiTableCell class.
	 * @typeofeditors ["PDFE"]
	 * @returns {"tableCell"}
	 * @see office-js-api/Examples/{Editor}/ApiTableCell/Methods/GetClassType.js
	 */
	ApiTableCell.prototype.GetClassType = function() {
		return "tableCell";
	};

	/**
	 * Returns the current cell content.
	 * @typeofeditors ["PDFE"]
	 * @returns {ApiDocumentContent}
	 * @see office-js-api/Examples/{Editor}/ApiTableCell/Methods/GetContent.js
	 */
	ApiTableCell.prototype.GetContent = function() {
		return Asc.editor.private_CreateApiDocContent(this.Cell.Content);
	};

	/**
	 * Specifies the shading which shall be applied to the extents of the current table cell.
	 * @typeofeditors ["PDFE"]
	 * @param {ShdType | ApiFill} sType - The shading type applied to the contents of the current table. Can be ShdType or ApiFill.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTableCell/Methods/SetShd.js
	 */
	ApiTableCell.prototype.SetShd = function(sType, r, g, b) {
		let oPr = this.Cell.Pr.Copy();
		let color = new Asc.asc_CColor({
			r: r,
			g: g,
			b: b,
			Auto: false
		});
		let oShd = new CDocumentShd();
		let _Shd = null;

		if (sType === "nil") {
			_Shd = {
				Value: Asc.c_oAscShdNil
			};
			oShd.Set_FromObject(_Shd);
			oPr.Shd = oShd;
		} else if (sType === "clear") {

			let Unifill = new AscFormat.CUniFill();
			Unifill.fill = new AscFormat.CSolidFill();
			Unifill.fill.color = AscFormat.CorrectUniColor(color, Unifill.fill.color, 1);
			_Shd = {
				Value: Asc.c_oAscShdClear,
				Color: {
					r: color.asc_getR(),
					g: color.asc_getG(),
					b: color.asc_getB()
				},
				Unifill: Unifill
			};

			oShd.Set_FromObject(_Shd);
			oPr.Shd = oShd;
		} else if (sType.GetClassType && sType.GetClassType() === "fill") {
			oShd.Value = Asc.c_oAscShdClear;
			oShd.Unifill = sType.UniFill;
			oPr.Shd = oShd;
		} else
			oPr.Shd = null;

		this.Cell.Set_Pr(oPr);
		return true;
	};

	/**
	 * Specifies an amount of space which shall be left between the bottom extent of the cell contents and the border
	 * of a specific individual table cell within a table.
	 * @typeofeditors ["PDFE"]
	 * @param {?twips} nValue - If this value is <code>null</code>, then default table cell bottom margin shall be used,
	 * otherwise override the table cell bottom margin with specified value for the current cell.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTableCell/Methods/SetCellMarginBottom.js
	 */
	ApiTableCell.prototype.SetCellMarginBottom = function(nValue) {
		let oPr = this.Cell.Pr.Copy();
		if (!oPr.TableCellMar) {
			oPr.TableCellMar = {
				Bottom: undefined,
				Left: undefined,
				Right: undefined,
				Top: undefined
			};
		}

		if (null === nValue)
			oPr.TableCellMar.Bottom = undefined;
		else
			oPr.TableCellMar.Bottom = private_GetTableMeasure("twips", nValue);

		this.Cell.Set_Pr(oPr);
		return true;
	};

	/**
	 * Specifies an amount of space which shall be left between the left extent of the current cell contents and the
	 * left edge border of a specific individual table cell within a table.
	 * @typeofeditors ["PDFE"]
	 * @param {?twips} nValue - If this value is <code>null</code>, then default table cell left margin shall be used,
	 * otherwise override the table cell left margin with specified value for the current cell.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTableCell/Methods/SetCellMarginLeft.js
	 */
	ApiTableCell.prototype.SetCellMarginLeft = function(nValue) {
		let oPr = this.Cell.Pr.Copy();
		if (!oPr.TableCellMar) {
			oPr.TableCellMar = {
				Bottom: undefined,
				Left: undefined,
				Right: undefined,
				Top: undefined
			};
		}

		if (null === nValue)
			oPr.TableCellMar.Left = undefined;
		else
			oPr.TableCellMar.Left = private_GetTableMeasure("twips", nValue);

		this.Cell.Set_Pr(oPr);
		return true;
	};

	/**
	 * Specifies an amount of space which shall be left between the right extent of the current cell contents and the
	 * right edge border of a specific individual table cell within a table.
	 * @typeofeditors ["PDFE"]
	 * @param {?twips} nValue - If this value is <code>null</code>, then default table cell right margin shall be used,
	 * otherwise override the table cell right margin with specified value for the current cell.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTableCell/Methods/SetCellMarginRight.js
	 */
	ApiTableCell.prototype.SetCellMarginRight = function(nValue) {
		let oPr = this.Cell.Pr.Copy();
		if (!oPr.TableCellMar) {
			oPr.TableCellMar = {
				Bottom: undefined,
				Left: undefined,
				Right: undefined,
				Top: undefined
			};
		}

		if (null === nValue)
			oPr.TableCellMar.Right = undefined;
		else
			oPr.TableCellMar.Right = private_GetTableMeasure("twips", nValue);

		this.Cell.Set_Pr(oPr);
		return true;
	};

	/**
	 * Specifies an amount of space which shall be left between the top extent of the current cell contents and the
	 * top edge border of a specific individual table cell within a table.
	 * @typeofeditors ["PDFE"]
	 * @param {?twips} nValue - If this value is <code>null</code>, then default table cell top margin shall be used,
	 * otherwise override the table cell top margin with specified value for the current cell.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTableCell/Methods/SetCellMarginTop.js
	 */
	ApiTableCell.prototype.SetCellMarginTop = function(nValue) {
		let oPr = this.Cell.Pr.Copy();
		if (!oPr.TableCellMar) {
			oPr.TableCellMar = {
				Bottom: undefined,
				Left: undefined,
				Right: undefined,
				Top: undefined
			};
		}

		if (null === nValue)
			oPr.TableCellMar.Top = undefined;
		else
			oPr.TableCellMar.Top = private_GetTableMeasure("twips", nValue);

		this.Cell.Set_Pr(oPr);
		return true;
	};

	/**
	 * Sets the border which shall be displayed at the bottom of the current table cell.
	 * @typeofeditors ["PDFE"]
	 * @param {mm} fSize - The width of the current border.
	 * @param {ApiFill} oApiFill - The color or pattern used to fill the current border.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTableCell/Methods/SetCellBorderBottom.js
	 */
	ApiTableCell.prototype.SetCellBorderBottom = function(fSize, oApiFill) {
		let oBorder = new CDocumentBorder();
		oBorder.Value = border_Single;
		oBorder.Size = fSize;
		oBorder.Space = 0;
		oBorder.Unifill = oApiFill.UniFill;

		let oPr = this.Cell.Pr.Copy();
		oPr.TableCellBorders.Bottom = oBorder;

		this.Cell.Set_Pr(oPr);
		return true;
	};

	/**
	 * Sets the border which shall be displayed at the left of the current table cell.
	 * @typeofeditors ["PDFE"]
	 * @param {mm} fSize - The width of the current border.
	 * @param {ApiFill} oApiFill - The color or pattern used to fill the current border.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTableCell/Methods/SetCellBorderLeft.js
	 */
	ApiTableCell.prototype.SetCellBorderLeft = function(fSize, oApiFill) {
		let oBorder = new CDocumentBorder();
		oBorder.Value = border_Single;
		oBorder.Size = fSize;
		oBorder.Space = 0;
		oBorder.Unifill = oApiFill.UniFill;

		let oPr = this.Cell.Pr.Copy();
		oPr.TableCellBorders.Left = oBorder;

		this.Cell.Set_Pr(oPr);
		return true;
	};

	/**
	 * Sets the border which shall be displayed at the right of the current table cell.
	 * @typeofeditors ["PDFE"]
	 * @param {mm} fSize - The width of the current border.
	 * @param {ApiFill} oApiFill - The color or pattern used to fill the current border.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTableCell/Methods/SetCellBorderRight.js
	 */
	ApiTableCell.prototype.SetCellBorderRight = function(fSize, oApiFill) {
		let oBorder = new CDocumentBorder();
		oBorder.Value = border_Single;
		oBorder.Size = fSize;
		oBorder.Space = 0;
		oBorder.Unifill = oApiFill.UniFill;

		let oPr = this.Cell.Pr.Copy();
		oPr.TableCellBorders.Right = oBorder;

		this.Cell.Set_Pr(oPr);
		return true;
	};

	/**
	 * Sets the border which shall be displayed at the top of the current table cell.
	 * @typeofeditors ["PDFE"]
	 * @param {mm} fSize - The width of the current border.
	 * @param {ApiFill} oApiFill - The color or pattern used to fill the current border.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTableCell/Methods/SetCellBorderTop.js
	 */
	ApiTableCell.prototype.SetCellBorderTop = function(fSize, oApiFill) {
		let oBorder = new CDocumentBorder();
		oBorder.Value = border_Single;
		oBorder.Size = fSize;
		oBorder.Space = 0;
		oBorder.Unifill = oApiFill.UniFill;

		let oPr = this.Cell.Pr.Copy();
		oPr.TableCellBorders.Top = oBorder;

		this.Cell.Set_Pr(oPr);
		return true;
	};

	/**
	 * Specifies the vertical alignment for text within the current table cell.
	 * @typeofeditors ["PDFE"]
	 * @param {VerticalTextAlign} sType - The type of the vertical alignment.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTableCell/Methods/SetVerticalAlign.js
	 */
	ApiTableCell.prototype.SetVerticalAlign = function(sType) {
		let oPr = this.Cell.Pr.Copy();
		if ("top" === sType)
			oPr.VAlign = vertalignjc_Top;
		else if ("bottom" === sType)
			oPr.VAlign = vertalignjc_Bottom;
		else if ("center" === sType)
			oPr.VAlign = vertalignjc_Center;

		this.Cell.Set_Pr(oPr);
		return true;
	};

	/**
	 * Specifies the direction of the text flow for the current table cell.
	 * @typeofeditors ["PDFE"]
	 * @param {TextFlowDirection} sType - The type of the text flow direction. 
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiTableCell/Methods/SetTextDirection.js
	 */
	ApiTableCell.prototype.SetTextDirection = function(sType) {
		let oPr = this.Cell.Pr.Copy();
		if ("lrtb" === sType)
			oPr.TextDirection = textdirection_LRTB;
		else if ("tbrl" === sType)
			oPr.TextDirection = textdirection_TBRL;
		else if ("btlr" === sType)
			oPr.TextDirection = textdirection_BTLR;

		this.Cell.Set_Pr(oPr);
		return true;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiChart
	//
	//------------------------------------------------------------------------------------------------------------------

	function ApiChart(Chart) {
		ApiDrawing.call(this, Chart);
		this.Chart = Chart;
	}
	ApiChart.prototype = Object.create(ApiDrawing.prototype);
	ApiChart.prototype.constructor = ApiChart;

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiGroup
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a group of drawings.
	 * @constructor
	 */
	function ApiGroup(oGroup) {
		ApiDrawing.call(this, oGroup);
	}
	ApiGroup.prototype = Object.create(ApiDrawing.prototype);
	ApiGroup.prototype.constructor = ApiGroup;

	/**
	 * Returns a type of the ApiGroup class.
	 * @memberof ApiGroup
	 * @typeofeditors ["PDFE"]
	 * @returns {"group"}
	 * @see office-js-api/Examples/{Editor}/ApiGroup/Methods/GetClassType.js
	 */
	ApiGroup.prototype.GetClassType = function() {
		return "group";
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiSmartArt
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Class representing a group of drawings.
	 * @constructor
	 */
	function ApiSmartArt(oGroup) {
		ApiDrawing.call(this, oGroup);
	}
	ApiSmartArt.prototype = Object.create(ApiDrawing.prototype);
	ApiSmartArt.prototype.constructor = ApiSmartArt;

	/**
	 * Returns a type of the ApiSmartArt class.
	 * @memberof ApiSmartArt
	 * @typeofeditors ["PDFE"]
	 * @returns {"group"}
	 * @see office-js-api/Examples/{Editor}/ApiSmartArt/Methods/GetClassType.js
	 */
	ApiSmartArt.prototype.GetClassType = function() {
		return "smartArt";
	};

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
			case "HH:MM": {
				return AscPDF.TimeFormatType["HH:MM"];
			}
			case "h:MM tt": {
				return AscPDF.TimeFormatType["h:MM tt"];
			}
			case "HH:MM:ss": {
				return AscPDF.TimeFormatType["HH:MM:ss"];
			}
			case "h:MM:ss tt": {
				return AscPDF.TimeFormatType["h:MM:ss tt"];
			}
		}
	}

	function private_GetStrTimeFormatType(nType) {
		switch (nType) {
			case AscPDF.TimeFormatType["HH:MM"]: {
				return "HH:MM";
			}
			case AscPDF.TimeFormatType["h:MM tt"]: {
				return "h:MM tt";
			}
			case AscPDF.TimeFormatType["HH:MM:ss"]: {
				return "HH:MM:ss";
			}
			case AscPDF.TimeFormatType["h:MM:ss tt"]: {
				return "h:MM:ss tt";
			}
		}
	}

	function private_GetInnerColorByRGB(r, g, b) {
		return [r / 255, g / 255, b / 255];
	}

	function private_GetAnnotApi(annot) {
		if (!annot) {
			return null;
		}

		switch (annot.GetType()) {
			case AscPDF.ANNOTATIONS_TYPES.Text: {
				return new ApiTextAnnotation(annot);
			}
			case AscPDF.ANNOTATIONS_TYPES.Circle: {
				return new ApiCircleAnnotation(annot);
			}
			case AscPDF.ANNOTATIONS_TYPES.Square: {
				return new ApiSquareAnnotation(annot);
			}
			case AscPDF.ANNOTATIONS_TYPES.FreeText: {
				return new ApiFreeTextAnnotation(annot);
			}
			case AscPDF.ANNOTATIONS_TYPES.Line: {
				return new ApiLineAnnotation(annot);
			}
			case AscPDF.ANNOTATIONS_TYPES.Ink: {
				return new ApiInkAnnotation(annot);
			}
			case AscPDF.ANNOTATIONS_TYPES.Polygon: {
				return new ApiPolygonAnnotation(annot);
			}
			case AscPDF.ANNOTATIONS_TYPES.PolyLine: {
				return new ApiPolyLineAnnotation(annot);
			}
			case AscPDF.ANNOTATIONS_TYPES.Stamp: {
				return new ApiStampAnnotation(annot);
			}
			case AscPDF.ANNOTATIONS_TYPES.Highlight: {
				return new ApiHighlightAnnotation(annot);
			}
			case AscPDF.ANNOTATIONS_TYPES.Strikeout: {
				return new ApiStrikeoutAnnotation(annot);
			}
			case AscPDF.ANNOTATIONS_TYPES.Underline: {
				return new ApiUnderlineAnnotation(annot);
			}
			case AscPDF.ANNOTATIONS_TYPES.Caret: {
				return new ApiCaretAnnotation(annot);
			}
			case AscPDF.ANNOTATIONS_TYPES.Redact: {
				return new ApiRedactAnnotation(annot);
			}
		}
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

	function private_IsValidRectDiff(value) {
		return (
			Array.isArray(value) &&
			value.length === 4 &&
			value.every(Number.isFinite)
		);
	}

	function private_CheckPoint(point) {
		if (!point) {
			AscBuilder.throwException("The point must be an object");
		}

		let x = AscBuilder.GetNumberParameter(point['x'], null);
		if (!x) {
			AscBuilder.throwException("The x coordinate of a point must be a number");
		}

		let y = AscBuilder.GetNumberParameter(point['y'], null);
		if (!y) {
			AscBuilder.throwException("The y coordinate of a point must be a number");
		}
	}

	function private_GetInnerLineEndType(type) {
		return AscPDF.LINE_END_TYPE[type];
	}

	function private_GetStrLineEndType(type) {
		switch (type) {
			case AscPDF.LINE_END_TYPE.square: {
				return "square";
			}
			case AscPDF.LINE_END_TYPE.circle: {
				return "circle";
			}
			case AscPDF.LINE_END_TYPE.diamond: {
				return "diamond";
			}
			case AscPDF.LINE_END_TYPE.openArrow: {
				return "openArrow";
			}
			case AscPDF.LINE_END_TYPE.closedArrow: {
				return "closedArrow";
			}
			case AscPDF.LINE_END_TYPE.none: {
				return "none";
			}
			case AscPDF.LINE_END_TYPE.butt: {
				return "butt";
			}
			case AscPDF.LINE_END_TYPE.rOpenArrow: {
				return "rOpenArrow";
			}
			case AscPDF.LINE_END_TYPE.rClosedArrow: {
				return "rClosedArrow";
			}
			case AscPDF.LINE_END_TYPE.slash: {
				return "slash";
			}
		}
	}

	function private_IsValidQuad(quad) {
		if (!quad || quad.length !== 8) return false;

		var i, v;

		for (i = 0; i < 8; i++) {
			v = quad[i];
			if (typeof v !== 'number' || !isFinite(v)) {
				return false;
			}
		}

		return (
			quad[0] <= quad[2] && // x1 <= x2
			quad[4] <= quad[6] && // x3 <= x4
			quad[1] <= quad[5] && // y1 <= y3
			quad[3] <= quad[7]    // y2 <= y4
		);
	}

	function private_ConvertRectToQuad(rect) {
		var x1 = rect[0];
		var y1 = rect[1];
		var x2 = rect[2];
		var y2 = rect[3];

		return [
			x1, y1, // left top
			x2, y1, // right top
			x1, y2, // left bottom
			x2, y2  // right bottom
		];
	}

	function private_WrapClassMethods(Class, before) {
		let target = Class.prototype;
		let proto = target;
		let seen = Object.create(null);

		while (proto && proto !== Object.prototype) {
			Object.getOwnPropertyNames(proto).forEach(function(key) {
				if (key === 'constructor') return;
				if (seen[key]) return;
				seen[key] = true;

				let desc = Object.getOwnPropertyDescriptor(proto, key);
				if (!desc || typeof desc.value !== 'function') return;

				let own = Object.getOwnPropertyDescriptor(target, key);
				if (own && typeof own.value === 'function' && own.value.__wrapped__) return;

				let original = desc.value;

				function wrapped() {
					if (before && before.call(this, key, arguments) === false) return;
					return original.apply(this, arguments);
				}
				wrapped.__wrapped__ = true;

				Object.defineProperty(target, key, {
					value: wrapped,
					writable: true,
					enumerable: desc.enumerable,
					configurable: true
				});
			});

			proto = Object.getPrototypeOf(proto);
		}
	}

	function private_NormalizeDegree(angle) {
		return ((angle % 360) + 360) % 360;
	}

	function private_GetDrawingDocument() {
		return Asc.editor.getDrawingDocument();
	}
	
	function private_PtToMM(pt) {
		return 25.4 / 72.0 * pt;
	}

	function private_MM2Pt(mm) {
		return mm / (25.4 / 72.0);
	}

	function private_EMU2MM(EMU) {
		return EMU / 36000.0;
	}

	function private_MM2EMU(mm) {
		return mm * 36000.0;
	}

	function private_GetTableMeasure(sType, nValue) {
		let nType = tblwidth_Auto;
		let nW    = 0;

		if ("auto" === sType) {
			nType = tblwidth_Auto;
			nW    = 0;
		}
		else if ("nil" === sType) {
			nType = tblwidth_Nil;
			nW    = 0;
		}
		else if ("percent" === sType) {
			nType = tblwidth_Pct;
			nW    = AscBuilder.private_GetInt(nValue, null, null);
		}
		else if ("twips" === sType) {
			nType = tblwidth_Mm;
			nW    = AscBuilder.private_Twips2MM(nValue);
		}

		return new CTableMeasurement(nType, nW);
	}


	// Api
	Api["GetDocument"]            = Api.GetDocument;
	Api["CreateTextField"]        = Api.CreateTextField;
	Api["CreateDateField"]        = Api.CreateDateField;
	Api["CreateImageField"]       = Api.CreateImageField;
	Api["CreateCheckboxField"]    = Api.CreateCheckboxField;
	Api["CreateRadiobuttonField"] = Api.CreateRadiobuttonField;
	Api["CreateComboboxField"]    = Api.CreateComboboxField;
	Api["CreateListboxField"]     = Api.CreateListboxField;
	Api["CreateTextAnnot"]        = Api.CreateTextAnnot;
	Api["CreateCircleAnnot"]      = Api.CreateCircleAnnot;
	Api["CreateSquareAnnot"]      = Api.CreateSquareAnnot;
	Api["CreateFreeTextAnnot"]    = Api.CreateFreeTextAnnot;
	Api["CreateLineAnnot"]        = Api.CreateLineAnnot;
	Api["CreateInkAnnot"]         = Api.CreateInkAnnot;
	Api["CreatePolygonAnnot"]     = Api.CreatePolygonAnnot;
	Api["CreatePolyLineAnnot"]    = Api.CreatePolyLineAnnot;
	Api["CreateStampAnnot"]       = Api.CreateStampAnnot;
	Api["CreateHighlightAnnot"]   = Api.CreateHighlightAnnot;
	Api["CreateStrikeoutAnnot"]   = Api.CreateStrikeoutAnnot;
	Api["CreateUnderlineAnnot"]   = Api.CreateUnderlineAnnot;
	Api["CreateCaretAnnot"]       = Api.CreateCaretAnnot;
	Api["CreateRedactAnnot"]      = Api.CreateRedactAnnot;
	Api["CreateParagraph"]        = Api.CreateParagraph;
	Api["CreateRichParagraph"]    = Api.CreateRichParagraph;
	Api["CreateRichRun"]          = Api.CreateRichRun;
	Api["CreateRichTextPr"]       = Api.CreateRichTextPr;
	Api["CreateRichParaPr"]       = Api.CreateRichParaPr;

	// ApiDocument
	ApiDocument.prototype["GetClassType"]					= ApiDocument.prototype.GetClassType;
	ApiDocument.prototype["AddPage"]						= ApiDocument.prototype.AddPage;
	ApiDocument.prototype["GetPage"]						= ApiDocument.prototype.GetPage;
	ApiDocument.prototype["RemovePage"]						= ApiDocument.prototype.RemovePage;
	ApiDocument.prototype["GetPagesCount"]					= ApiDocument.prototype.GetPagesCount;
	ApiDocument.prototype["GetAllFields"]					= ApiDocument.prototype.GetAllFields;
	ApiDocument.prototype["GetFieldByName"]					= ApiDocument.prototype.GetFieldByName;
	ApiDocument.prototype["SearchAndRedact"]				= ApiDocument.prototype.SearchAndRedact;
	ApiDocument.prototype["ApplyRedact"]					= ApiDocument.prototype.ApplyRedact;

	// ApiPage
	ApiPage.prototype["GetClassType"]						= ApiPage.prototype.GetClassType;
	ApiPage.prototype["SetRotation"]						= ApiPage.prototype.SetRotation;
	ApiPage.prototype["GetRotation"]						= ApiPage.prototype.GetRotation;
	ApiPage.prototype["GetIndex"]							= ApiPage.prototype.GetIndex;
	ApiPage.prototype["GetAllWidgets"]						= ApiPage.prototype.GetAllWidgets;
	ApiPage.prototype["GetAllAnnots"]						= ApiPage.prototype.GetAllAnnots;
	ApiPage.prototype["Search"]								= ApiPage.prototype.Search;
	ApiPage.prototype["SetSelection"]						= ApiPage.prototype.SetSelection;
	ApiPage.prototype["GetSelectionQuads"]					= ApiPage.prototype.GetSelectionQuads;
	ApiPage.prototype["GetSelectedText"]					= ApiPage.prototype.GetSelectedText;
	ApiPage.prototype["RecognizeContent"]					= ApiPage.prototype.RecognizeContent;
	ApiPage.prototype["GetAllDrawings"]						= ApiPage.prototype.GetAllDrawings;

	// ApiBaseField
	ApiBaseField.prototype["SetRect"]						= ApiBaseField.prototype.SetRect;
	ApiBaseField.prototype["GetRect"]						= ApiBaseField.prototype.GetRect;
	ApiBaseField.prototype["SetFullName"]					= ApiBaseField.prototype.SetFullName;
	ApiBaseField.prototype["GetFullName"]					= ApiBaseField.prototype.GetFullName;
	ApiBaseField.prototype["SetPartialName"]				= ApiBaseField.prototype.SetPartialName;
	ApiBaseField.prototype["GetPartialName"]				= ApiBaseField.prototype.GetPartialName;
	ApiBaseField.prototype["SetRequired"]					= ApiBaseField.prototype.SetRequired;
	ApiBaseField.prototype["IsRequired"]					= ApiBaseField.prototype.IsRequired;
	ApiBaseField.prototype["SetReadOnly"]					= ApiBaseField.prototype.SetReadOnly;
	ApiBaseField.prototype["IsReadOnly"]					= ApiBaseField.prototype.IsReadOnly;
	ApiBaseField.prototype["SetValue"]						= ApiBaseField.prototype.SetValue;
	ApiBaseField.prototype["GetValue"]						= ApiBaseField.prototype.GetValue;
	ApiBaseField.prototype["AddWidget"]						= ApiBaseField.prototype.AddWidget;
	ApiBaseField.prototype["GetAllWidgets"]					= ApiBaseField.prototype.GetAllWidgets;
	ApiBaseField.prototype["Delete"]						= ApiBaseField.prototype.Delete;

	// ApiBaseWidget
	ApiBaseWidget.prototype["GetClassType"]					= ApiBaseWidget.prototype.GetClassType;
	ApiBaseWidget.prototype["SetPosition"]					= ApiBaseWidget.prototype.SetPosition;
	ApiBaseWidget.prototype["GetPosition"]					= ApiBaseWidget.prototype.GetPosition;
	ApiBaseWidget.prototype["SetBorderColor"]				= ApiBaseWidget.prototype.SetBorderColor;
	ApiBaseWidget.prototype["GetBorderColor"]				= ApiBaseWidget.prototype.GetBorderColor;
	ApiBaseWidget.prototype["SetBorderWidth"]				= ApiBaseWidget.prototype.SetBorderWidth;
	ApiBaseWidget.prototype["GetBorderWidth"]				= ApiBaseWidget.prototype.GetBorderWidth;
	ApiBaseWidget.prototype["SetBorderStyle"]				= ApiBaseWidget.prototype.SetBorderStyle;
	ApiBaseWidget.prototype["GetBorderStyle"]				= ApiBaseWidget.prototype.GetBorderStyle;
	ApiBaseWidget.prototype["SetBackgroundColor"]			= ApiBaseWidget.prototype.SetBackgroundColor;
	ApiBaseWidget.prototype["GetBackgroundColor"]			= ApiBaseWidget.prototype.GetBackgroundColor;
	ApiBaseWidget.prototype["SetTextColor"]					= ApiBaseWidget.prototype.SetTextColor;
	ApiBaseWidget.prototype["GetTextColor"]					= ApiBaseWidget.prototype.GetTextColor;
	ApiBaseWidget.prototype["SetTextSize"]					= ApiBaseWidget.prototype.SetTextSize;
	ApiBaseWidget.prototype["GetTextSize"]					= ApiBaseWidget.prototype.GetTextSize;
	ApiBaseWidget.prototype["SetAutoFit"]					= ApiBaseWidget.prototype.SetAutoFit;
	ApiBaseWidget.prototype["IsAutoFit"]					= ApiBaseWidget.prototype.IsAutoFit;
	ApiBaseWidget.prototype["Delete"]						= ApiBaseWidget.prototype.Delete;

	// ApiTextField
	ApiTextField.prototype["GetClassType"]					= ApiTextField.prototype.GetClassType;
	ApiTextField.prototype["SetMultiline"]					= ApiTextField.prototype.SetMultiline;
	ApiTextField.prototype["IsMultiline"]					= ApiTextField.prototype.IsMultiline;
	ApiTextField.prototype["SetCharLimit"]					= ApiTextField.prototype.SetCharLimit;
	ApiTextField.prototype["GetCharLimit"]					= ApiTextField.prototype.GetCharLimit;
	ApiTextField.prototype["SetComb"]						= ApiTextField.prototype.SetComb;
	ApiTextField.prototype["IsComb"]						= ApiTextField.prototype.IsComb;
	ApiTextField.prototype["SetScrollLongText"]				= ApiTextField.prototype.SetScrollLongText;
	ApiTextField.prototype["IsScrollLongText"]				= ApiTextField.prototype.IsScrollLongText;
	ApiTextField.prototype["SetNumberFormat"]				= ApiTextField.prototype.SetNumberFormat;
	ApiTextField.prototype["SetPercentageFormat"]			= ApiTextField.prototype.SetPercentageFormat;
	ApiTextField.prototype["SetDateFormat"]					= ApiTextField.prototype.SetDateFormat;
	ApiTextField.prototype["SetTimeFormat"]					= ApiTextField.prototype.SetTimeFormat;
	ApiTextField.prototype["SetSpecialFormat"]				= ApiTextField.prototype.SetSpecialFormat;
	ApiTextField.prototype["SetMask"]						= ApiTextField.prototype.SetMask;
	ApiTextField.prototype["SetRegularExp"]					= ApiTextField.prototype.SetRegularExp;
	ApiTextField.prototype["ClearFormat"]					= ApiTextField.prototype.ClearFormat;
	ApiTextField.prototype["SetValidateRange"]				= ApiTextField.prototype.SetValidateRange;

	// ApiTextWidget
	ApiTextWidget.prototype["GetClassType"]					= ApiTextWidget.prototype.GetClassType;
	ApiTextWidget.prototype["SetPlaceholder"]				= ApiTextWidget.prototype.SetPlaceholder;
	ApiTextWidget.prototype["GetPlaceholder"]				= ApiTextWidget.prototype.GetPlaceholder;
	ApiTextWidget.prototype["SetRegularExp"]				= ApiTextWidget.prototype.SetRegularExp;
	ApiTextWidget.prototype["GetRegularExp"]				= ApiTextWidget.prototype.GetRegularExp;

	// ApiBaseListField
	ApiBaseListField.prototype["AddOption"]					= ApiBaseListField.prototype.AddOption;
	ApiBaseListField.prototype["RemoveOption"]				= ApiBaseListField.prototype.RemoveOption;
	ApiBaseListField.prototype["MoveOption"]				= ApiBaseListField.prototype.MoveOption;
	ApiBaseListField.prototype["GetOption"]					= ApiBaseListField.prototype.GetOption;
	ApiBaseListField.prototype["GetOptions"]				= ApiBaseListField.prototype.GetOptions;
	ApiBaseListField.prototype["SetCommitOnSelChange"]		= ApiBaseListField.prototype.SetCommitOnSelChange;
	ApiBaseListField.prototype["IsCommitOnSelChange"]		= ApiBaseListField.prototype.IsCommitOnSelChange;
	ApiBaseListField.prototype["SetValueIndexes"]			= ApiBaseListField.prototype.SetValueIndexes;
	ApiBaseListField.prototype["GetValueIndexes"]			= ApiBaseListField.prototype.GetValueIndexes;

	// ApiComboboxField
	ApiComboboxField.prototype["GetClassType"]				= ApiComboboxField.prototype.GetClassType;
	ApiComboboxField.prototype["SetEditable"]				= ApiComboboxField.prototype.SetEditable;
	ApiComboboxField.prototype["IsEditable"]				= ApiComboboxField.prototype.IsEditable;
	ApiComboboxField.prototype["SetNumberFormat"]			= ApiComboboxField.prototype.SetNumberFormat;
	ApiComboboxField.prototype["SetPercentageFormat"]		= ApiComboboxField.prototype.SetPercentageFormat;
	ApiComboboxField.prototype["SetDateFormat"]				= ApiComboboxField.prototype.SetDateFormat;
	ApiComboboxField.prototype["SetTimeFormat"]				= ApiComboboxField.prototype.SetTimeFormat;
	ApiComboboxField.prototype["SetSpecialFormat"]			= ApiComboboxField.prototype.SetSpecialFormat;
	ApiComboboxField.prototype["SetMask"]					= ApiComboboxField.prototype.SetMask;
	ApiComboboxField.prototype["SetRegularExp"]				= ApiComboboxField.prototype.SetRegularExp;
	ApiComboboxField.prototype["ClearFormat"]				= ApiComboboxField.prototype.ClearFormat;
	ApiComboboxField.prototype["SetValidateRange"]			= ApiComboboxField.prototype.SetValidateRange;

	// ApiListboxField
	ApiListboxField.prototype["GetClassType"]				= ApiListboxField.prototype.GetClassType;
	ApiListboxField.prototype["SetMultipleSelection"]		= ApiListboxField.prototype.SetMultipleSelection;
	ApiListboxField.prototype["IsMultipleSelection"]		= ApiListboxField.prototype.IsMultipleSelection;

	// ApiCheckboxField
	ApiCheckboxField.prototype["GetClassType"]				= ApiCheckboxField.prototype.GetClassType;
	ApiCheckboxField.prototype["SetToggleToOff"]			= ApiCheckboxField.prototype.SetToggleToOff;
	ApiCheckboxField.prototype["IsToggleToOff"]				= ApiCheckboxField.prototype.IsToggleToOff;
	ApiCheckboxField.prototype["AddOption"]					= ApiCheckboxField.prototype.AddOption;

	// ApiRadiobuttonField
	ApiRadiobuttonField.prototype["GetClassType"]			= ApiRadiobuttonField.prototype.GetClassType;
	ApiRadiobuttonField.prototype["SetCheckInUnison"]		= ApiRadiobuttonField.prototype.SetCheckInUnison;
	ApiRadiobuttonField.prototype["IsCheckInUnison"]		= ApiRadiobuttonField.prototype.IsCheckInUnison;

	// ApiCheckboxWidget
	ApiCheckboxWidget.prototype["GetClassType"]				= ApiCheckboxWidget.prototype.GetClassType;
	ApiCheckboxWidget.prototype["SetChecked"]				= ApiCheckboxWidget.prototype.SetChecked;
	ApiCheckboxWidget.prototype["IsChecked"]				= ApiCheckboxWidget.prototype.IsChecked;
	ApiCheckboxWidget.prototype["SetCheckStyle"]			= ApiCheckboxWidget.prototype.SetCheckStyle;
	ApiCheckboxWidget.prototype["GetCheckStyle"]			= ApiCheckboxWidget.prototype.GetCheckStyle;
	ApiCheckboxWidget.prototype["SetExportValue"]			= ApiCheckboxWidget.prototype.SetExportValue;
	ApiCheckboxWidget.prototype["GetExportValue"]			= ApiCheckboxWidget.prototype.GetExportValue;
	ApiCheckboxWidget.prototype["SetCheckedByDefault"]		= ApiCheckboxWidget.prototype.SetCheckedByDefault;
	ApiCheckboxWidget.prototype["IsCheckedByDefault"]		= ApiCheckboxWidget.prototype.IsCheckedByDefault;

	// ApiButtonField
	ApiButtonField.prototype["GetClassType"]				= ApiButtonField.prototype.GetClassType;

	// ApiButtonWidget
	ApiButtonWidget.prototype["GetClassType"]				= ApiButtonWidget.prototype.GetClassType;
	ApiButtonWidget.prototype["SetLayout"]					= ApiButtonWidget.prototype.SetLayout;
	ApiButtonWidget.prototype["GetLayout"]					= ApiButtonWidget.prototype.GetLayout;
	ApiButtonWidget.prototype["SetScaleWhen"]				= ApiButtonWidget.prototype.SetScaleWhen;
	ApiButtonWidget.prototype["GetScaleWhen"]				= ApiButtonWidget.prototype.GetScaleWhen;
	ApiButtonWidget.prototype["SetScaleHow"]				= ApiButtonWidget.prototype.SetScaleHow;
	ApiButtonWidget.prototype["GetScaleHow"]				= ApiButtonWidget.prototype.GetScaleHow;
	ApiButtonWidget.prototype["SetFitBounds"]				= ApiButtonWidget.prototype.SetFitBounds;
	ApiButtonWidget.prototype["IsFitBounds"]				= ApiButtonWidget.prototype.IsFitBounds;
	ApiButtonWidget.prototype["SetIconXPos"]				= ApiButtonWidget.prototype.SetIconXPos;
	ApiButtonWidget.prototype["GetIconXPos"]				= ApiButtonWidget.prototype.GetIconXPos;
	ApiButtonWidget.prototype["SetIconYPos"]				= ApiButtonWidget.prototype.SetIconYPos;
	ApiButtonWidget.prototype["GetIconYPos"]				= ApiButtonWidget.prototype.GetIconYPos;
	ApiButtonWidget.prototype["SetBehavior"]				= ApiButtonWidget.prototype.SetBehavior;
	ApiButtonWidget.prototype["GetBehavior"]				= ApiButtonWidget.prototype.GetBehavior;
	ApiButtonWidget.prototype["SetLabel"]					= ApiButtonWidget.prototype.SetLabel;
	ApiButtonWidget.prototype["GetLabel"]					= ApiButtonWidget.prototype.GetLabel;
	ApiButtonWidget.prototype["SetImage"]					= ApiButtonWidget.prototype.SetImage;

	// ApiBaseAnnotation
	ApiBaseAnnotation.prototype["SetRect"]					= ApiBaseAnnotation.prototype.SetRect;
	ApiBaseAnnotation.prototype["GetRect"]					= ApiBaseAnnotation.prototype.GetRect;
	ApiBaseAnnotation.prototype["SetPosition"]				= ApiBaseAnnotation.prototype.SetPosition;
	ApiBaseAnnotation.prototype["GetPosition"]				= ApiBaseAnnotation.prototype.GetPosition;
	ApiBaseAnnotation.prototype["SetBorderColor"]			= ApiBaseAnnotation.prototype.SetBorderColor;
	ApiBaseAnnotation.prototype["GetBorderColor"]			= ApiBaseAnnotation.prototype.GetBorderColor;
	ApiBaseAnnotation.prototype["SetFillColor"]				= ApiBaseAnnotation.prototype.SetFillColor;
	ApiBaseAnnotation.prototype["GetFillColor"]				= ApiBaseAnnotation.prototype.GetFillColor;
	ApiBaseAnnotation.prototype["SetBorderWidth"]			= ApiBaseAnnotation.prototype.SetBorderWidth;
	ApiBaseAnnotation.prototype["GetBorderWidth"]			= ApiBaseAnnotation.prototype.GetBorderWidth;
	ApiBaseAnnotation.prototype["SetBorderStyle"]			= ApiBaseAnnotation.prototype.SetBorderStyle;
	ApiBaseAnnotation.prototype["GetBorderStyle"]			= ApiBaseAnnotation.prototype.GetBorderStyle;
	ApiBaseAnnotation.prototype["SetAuthorName"]			= ApiBaseAnnotation.prototype.SetAuthorName;
	ApiBaseAnnotation.prototype["GetAuthorName"]			= ApiBaseAnnotation.prototype.GetAuthorName;
	ApiBaseAnnotation.prototype["SetContents"]				= ApiBaseAnnotation.prototype.SetContents;
	ApiBaseAnnotation.prototype["GetContents"]				= ApiBaseAnnotation.prototype.GetContents;
	ApiBaseAnnotation.prototype["SetCreationDate"]			= ApiBaseAnnotation.prototype.SetCreationDate;
	ApiBaseAnnotation.prototype["GetCreationDate"]			= ApiBaseAnnotation.prototype.GetCreationDate;
	ApiBaseAnnotation.prototype["SetModDate"]				= ApiBaseAnnotation.prototype.SetModDate;
	ApiBaseAnnotation.prototype["GetModDate"]				= ApiBaseAnnotation.prototype.GetModDate;
	ApiBaseAnnotation.prototype["SetUniqueName"]			= ApiBaseAnnotation.prototype.SetUniqueName;
	ApiBaseAnnotation.prototype["GetUniqueName"]			= ApiBaseAnnotation.prototype.GetUniqueName;
	ApiBaseAnnotation.prototype["SetOpacity"]				= ApiBaseAnnotation.prototype.SetOpacity;
	ApiBaseAnnotation.prototype["GetOpacity"]				= ApiBaseAnnotation.prototype.GetOpacity;
	ApiBaseAnnotation.prototype["SetSubject"]				= ApiBaseAnnotation.prototype.SetSubject;
	ApiBaseAnnotation.prototype["GetSubject"]				= ApiBaseAnnotation.prototype.GetSubject;
	ApiBaseAnnotation.prototype["SetDisplay"]				= ApiBaseAnnotation.prototype.SetDisplay;
	ApiBaseAnnotation.prototype["GetDisplay"]				= ApiBaseAnnotation.prototype.GetDisplay;
	ApiBaseAnnotation.prototype["SetDashPattern"]			= ApiBaseAnnotation.prototype.SetDashPattern;
	ApiBaseAnnotation.prototype["GetDashPattern"]			= ApiBaseAnnotation.prototype.GetDashPattern;
	ApiBaseAnnotation.prototype["SetBorderEffectStyle"]		= ApiBaseAnnotation.prototype.SetBorderEffectStyle;
	ApiBaseAnnotation.prototype["GetBorderEffectStyle"]		= ApiBaseAnnotation.prototype.GetBorderEffectStyle;
	ApiBaseAnnotation.prototype["SetBorderEffectIntensity"]	= ApiBaseAnnotation.prototype.SetBorderEffectIntensity;
	ApiBaseAnnotation.prototype["GetBorderEffectIntensity"]	= ApiBaseAnnotation.prototype.GetBorderEffectIntensity;
	ApiBaseAnnotation.prototype["AddReply"]					= ApiBaseAnnotation.prototype.AddReply;
	ApiBaseAnnotation.prototype["GetReplies"]				= ApiBaseAnnotation.prototype.GetReplies;
	ApiBaseAnnotation.prototype["Delete"]					= ApiBaseAnnotation.prototype.Delete;

	// ApiTextAnnotation
	ApiTextAnnotation.prototype["GetClassType"]				= ApiTextAnnotation.prototype.GetClassType;
	ApiTextAnnotation.prototype["SetIconType"]				= ApiTextAnnotation.prototype.SetIconType;
	ApiTextAnnotation.prototype["GetIconType"]				= ApiTextAnnotation.prototype.GetIconType;
	
	// ApiCircleAnnotation
	ApiCircleAnnotation.prototype["GetClassType"]			= ApiCircleAnnotation.prototype.GetClassType;
	ApiCircleAnnotation.prototype["SetRectDiff"]			= ApiCircleAnnotation.prototype.SetRectDiff;
	ApiCircleAnnotation.prototype["GetRectDiff"]			= ApiCircleAnnotation.prototype.GetRectDiff;

	// ApiSquareAnnotation
	ApiSquareAnnotation.prototype["GetClassType"]			= ApiSquareAnnotation.prototype.GetClassType;
	ApiSquareAnnotation.prototype["SetRectDiff"]			= ApiSquareAnnotation.prototype.SetRectDiff;
	ApiSquareAnnotation.prototype["GetRectDiff"]			= ApiSquareAnnotation.prototype.GetRectDiff;

	// ApiFreeTextAnnotation
	ApiFreeTextAnnotation.prototype["GetClassType"]			= ApiFreeTextAnnotation.prototype.GetClassType;
	ApiFreeTextAnnotation.prototype["SetIntent"]			= ApiFreeTextAnnotation.prototype.SetIntent;
	ApiFreeTextAnnotation.prototype["GetIntent"]			= ApiFreeTextAnnotation.prototype.GetIntent;
	ApiFreeTextAnnotation.prototype["SetCallout"]			= ApiFreeTextAnnotation.prototype.SetCallout;
	ApiFreeTextAnnotation.prototype["GetCallout"]			= ApiFreeTextAnnotation.prototype.GetCallout;
	ApiFreeTextAnnotation.prototype["SetRectDiff"]			= ApiFreeTextAnnotation.prototype.SetRectDiff;
	ApiFreeTextAnnotation.prototype["GetRectDiff"]			= ApiFreeTextAnnotation.prototype.GetRectDiff;
	ApiFreeTextAnnotation.prototype["GetContent"]			= ApiFreeTextAnnotation.prototype.GetContent;

	// ApiLineAnnotation
	ApiLineAnnotation.prototype["GetClassType"]				= ApiLineAnnotation.prototype.GetClassType;
	ApiLineAnnotation.prototype["SetStartPoint"]			= ApiLineAnnotation.prototype.SetStartPoint;
	ApiLineAnnotation.prototype["GetStartPoint"]			= ApiLineAnnotation.prototype.GetStartPoint;
	ApiLineAnnotation.prototype["SetEndPoint"]				= ApiLineAnnotation.prototype.SetEndPoint;
	ApiLineAnnotation.prototype["GetEndPoint"]				= ApiLineAnnotation.prototype.GetEndPoint;
	ApiLineAnnotation.prototype["SetStartStyle"]			= ApiLineAnnotation.prototype.SetStartStyle;
	ApiLineAnnotation.prototype["GetStartStyle"]			= ApiLineAnnotation.prototype.GetStartStyle;
	ApiLineAnnotation.prototype["SetEndStyle"]				= ApiLineAnnotation.prototype.SetEndStyle;
	ApiLineAnnotation.prototype["GetEndStyle"]				= ApiLineAnnotation.prototype.GetEndStyle;

	// ApiInkAnnotation
	ApiInkAnnotation.prototype["GetClassType"]				= ApiInkAnnotation.prototype.GetClassType;
	ApiInkAnnotation.prototype["SetPathList"]				= ApiInkAnnotation.prototype.SetPathList;
	ApiInkAnnotation.prototype["GetPathList"]				= ApiInkAnnotation.prototype.GetPathList;

	// ApiPolygonAnnotation
	ApiPolygonAnnotation.prototype["GetClassType"]			= ApiPolygonAnnotation.prototype.GetClassType;
	ApiPolygonAnnotation.prototype["SetVertices"]			= ApiPolygonAnnotation.prototype.SetVertices;
	ApiPolygonAnnotation.prototype["GetVertices"]			= ApiPolygonAnnotation.prototype.GetVertices;

	// ApiPolyLineAnnotation
	ApiPolyLineAnnotation.prototype["GetClassType"]			= ApiPolyLineAnnotation.prototype.GetClassType;
	ApiPolyLineAnnotation.prototype["SetVertices"]			= ApiPolyLineAnnotation.prototype.SetVertices;
	ApiPolyLineAnnotation.prototype["GetVertices"]			= ApiPolyLineAnnotation.prototype.GetVertices;
	ApiPolyLineAnnotation.prototype["SetStartStyle"]		= ApiPolyLineAnnotation.prototype.SetStartStyle;
	ApiPolyLineAnnotation.prototype["GetStartStyle"]		= ApiPolyLineAnnotation.prototype.GetStartStyle;
	ApiPolyLineAnnotation.prototype["SetEndStyle"]			= ApiPolyLineAnnotation.prototype.SetEndStyle;
	ApiPolyLineAnnotation.prototype["GetEndStyle"]			= ApiPolyLineAnnotation.prototype.GetEndStyle;

	// ApiStampAnnotation
	ApiStampAnnotation.prototype["GetClassType"]			= ApiStampAnnotation.prototype.GetClassType;
	ApiStampAnnotation.prototype["GetType"]					= ApiStampAnnotation.prototype.GetType;
	ApiStampAnnotation.prototype["SetScale"]				= ApiStampAnnotation.prototype.SetScale;
	ApiStampAnnotation.prototype["GetScale"]				= ApiStampAnnotation.prototype.GetScale;
	ApiStampAnnotation.prototype["SetRotation"]				= ApiStampAnnotation.prototype.SetRotation;
	ApiStampAnnotation.prototype["GetRotation"]				= ApiStampAnnotation.prototype.GetRotation;

	// ApiBaseMarkupAnnotation
	ApiBaseMarkupAnnotation.prototype["GetClassType"]		= ApiBaseMarkupAnnotation.prototype.GetClassType;
	ApiBaseMarkupAnnotation.prototype["SetQuads"]			= ApiBaseMarkupAnnotation.prototype.SetQuads;
	ApiBaseMarkupAnnotation.prototype["GetQuads"]			= ApiBaseMarkupAnnotation.prototype.GetQuads;

	// ApiHighlightAnnotation
	ApiHighlightAnnotation.prototype["GetClassType"]		= ApiHighlightAnnotation.prototype.GetClassType;

	// ApiStrikeoutAnnotation
	ApiStrikeoutAnnotation.prototype["GetClassType"]		= ApiStrikeoutAnnotation.prototype.GetClassType;

	// ApiUnderlineAnnotation
	ApiUnderlineAnnotation.prototype["GetClassType"]		= ApiUnderlineAnnotation.prototype.GetClassType;

	// ApiCaretAnnotation
	ApiCaretAnnotation.prototype["GetClassType"]			= ApiCaretAnnotation.prototype.GetClassType;

	// ApiRedactAnnotation
	ApiRedactAnnotation.prototype["GetClassType"]			= ApiRedactAnnotation.prototype.GetClassType;

	// ApiRichContent
	ApiRichContent.prototype["GetClassType"]				= ApiRichContent.prototype.GetClassType;
	ApiRichContent.prototype["GetElementsCount"]			= ApiRichContent.prototype.GetElementsCount;
	ApiRichContent.prototype["GetElement"]					= ApiRichContent.prototype.GetElement;
	ApiRichContent.prototype["AddElement"]					= ApiRichContent.prototype.AddElement;
	ApiRichContent.prototype["Push"]						= ApiRichContent.prototype.Push;
	ApiRichContent.prototype["RemoveElement"]				= ApiRichContent.prototype.RemoveElement;
	ApiRichContent.prototype["GetContent"]					= ApiRichContent.prototype.GetContent;
	ApiRichContent.prototype["GetText"]						= ApiRichContent.prototype.GetText;
	ApiRichContent.prototype["GetCurrentParagraph"]			= ApiRichContent.prototype.GetCurrentParagraph;
	ApiRichContent.prototype["GetCurrentRun"]				= ApiRichContent.prototype.GetCurrentRun;

	// ApiRichParaPr
	ApiRichParaPr.prototype["GetClassType"]					= ApiRichParaPr.prototype.GetClassType;
	ApiRichParaPr.prototype["SetJc"]						= ApiRichParaPr.prototype.SetJc;
	ApiRichParaPr.prototype["GetJc"]						= ApiRichParaPr.prototype.GetJc;

	// ApiRichParagraph
	ApiRichParagraph.prototype["GetClassType"]				= ApiRichParagraph.prototype.GetClassType;
	ApiRichParagraph.prototype["AddText"]					= ApiRichParagraph.prototype.AddText;
	ApiRichParagraph.prototype["SetParaPr"]					= ApiRichParagraph.prototype.SetParaPr;
	ApiRichParagraph.prototype["GetParaPr"]					= ApiRichParagraph.prototype.GetParaPr;
	ApiRichParagraph.prototype["GetElementsCount"]			= ApiRichParagraph.prototype.GetElementsCount;
	ApiRichParagraph.prototype["AddElement"]				= ApiRichParagraph.prototype.AddElement;
	ApiRichParagraph.prototype["Push"]						= ApiRichParagraph.prototype.Push;
	ApiRichParagraph.prototype["GetElement"]				= ApiRichParagraph.prototype.GetElement;
	ApiRichParagraph.prototype["RemoveElement"]				= ApiRichParagraph.prototype.RemoveElement;
	ApiRichParagraph.prototype["RemoveAllElements"]			= ApiRichParagraph.prototype.RemoveAllElements;
	ApiRichParagraph.prototype["Delete"]					= ApiRichParagraph.prototype.Delete;
	ApiRichParagraph.prototype["GetNext"]					= ApiRichParagraph.prototype.GetNext;
	ApiRichParagraph.prototype["GetPrevious"]				= ApiRichParagraph.prototype.GetPrevious;
	ApiRichParagraph.prototype["Last"]						= ApiRichParagraph.prototype.Last;
	ApiRichParagraph.prototype["Copy"]						= ApiRichParagraph.prototype.Copy;
	ApiRichParagraph.prototype["SetReadingOrder"]			= ApiRichParagraph.prototype.SetReadingOrder;
	ApiRichParagraph.prototype["GetText"]					= ApiRichParagraph.prototype.GetText;
	ApiRichParagraph.prototype["GetPosInParent"]			= ApiRichParagraph.prototype.GetPosInParent;

	// ApiRichTextPr
	ApiRichTextPr.prototype["GetClassType"]					= ApiRichTextPr.prototype.GetClassType;
	ApiRichTextPr.prototype["SetBold"]						= ApiRichTextPr.prototype.SetBold;
	ApiRichTextPr.prototype["GetBold"]						= ApiRichTextPr.prototype.GetBold;
	ApiRichTextPr.prototype["SetItalic"]					= ApiRichTextPr.prototype.SetItalic;
	ApiRichTextPr.prototype["GetItalic"]					= ApiRichTextPr.prototype.GetItalic;
	ApiRichTextPr.prototype["SetStrikeout"]					= ApiRichTextPr.prototype.SetStrikeout;
	ApiRichTextPr.prototype["GetStrikeout"]					= ApiRichTextPr.prototype.GetStrikeout;
	ApiRichTextPr.prototype["SetUnderline"]					= ApiRichTextPr.prototype.SetUnderline;
	ApiRichTextPr.prototype["GetUnderline"]					= ApiRichTextPr.prototype.GetUnderline;
	ApiRichTextPr.prototype["SetFontFamily"]				= ApiRichTextPr.prototype.SetFontFamily;
	ApiRichTextPr.prototype["GetFontFamily"]				= ApiRichTextPr.prototype.GetFontFamily;
	ApiRichTextPr.prototype["SetFontSize"]					= ApiRichTextPr.prototype.SetFontSize;
	ApiRichTextPr.prototype["GetFontSize"]					= ApiRichTextPr.prototype.GetFontSize;
	ApiRichTextPr.prototype["SetColor"]						= ApiRichTextPr.prototype.SetColor;
	ApiRichTextPr.prototype["GetColor"]						= ApiRichTextPr.prototype.GetColor;
	ApiRichTextPr.prototype["SetVertAlign"]					= ApiRichTextPr.prototype.SetVertAlign;
	ApiRichTextPr.prototype["GetVertAlign"]					= ApiRichTextPr.prototype.GetVertAlign;

	// ApiRichRun
	ApiRichRun.prototype["GetClassType"]					= ApiRichRun.prototype.GetClassType;
	ApiRichRun.prototype["SetTextPr"]						= ApiRichRun.prototype.SetTextPr;
	ApiRichRun.prototype["GetTextPr"]						= ApiRichRun.prototype.GetTextPr;
	ApiRichRun.prototype["ClearContent"]					= ApiRichRun.prototype.ClearContent;
	ApiRichRun.prototype["RemoveAllElements"]				= ApiRichRun.prototype.RemoveAllElements;
	ApiRichRun.prototype["Delete"]							= ApiRichRun.prototype.Delete;
	ApiRichRun.prototype["AddText"]							= ApiRichRun.prototype.AddText;
	ApiRichRun.prototype["Copy"]							= ApiRichRun.prototype.Copy;
	ApiRichRun.prototype["GetText"]							= ApiRichRun.prototype.GetText;
	ApiRichRun.prototype["GetParentParagraph"]				= ApiRichRun.prototype.GetParentParagraph;
	
	// ApiDrawing
	ApiDrawing.prototype["GetClassType"]					= ApiDrawing.prototype.GetClassType;
	ApiDrawing.prototype["GetParentPage"]					= ApiDrawing.prototype.GetParentPage;
	ApiDrawing.prototype["SetSize"]							= ApiDrawing.prototype.SetSize;
	ApiDrawing.prototype["SetPosition"]						= ApiDrawing.prototype.SetPosition;
	ApiDrawing.prototype["SetPosX"]							= ApiDrawing.prototype.SetPosX;
	ApiDrawing.prototype["GetPosX"]							= ApiDrawing.prototype.GetPosX;
	ApiDrawing.prototype["SetPosY"]							= ApiDrawing.prototype.SetPosY;
	ApiDrawing.prototype["GetPosY"]							= ApiDrawing.prototype.GetPosY;
	ApiDrawing.prototype["Copy"]							= ApiDrawing.prototype.Copy;
	ApiDrawing.prototype["Delete"]							= ApiDrawing.prototype.Delete;
	ApiDrawing.prototype["GetWidth"]						= ApiDrawing.prototype.GetWidth;
	ApiDrawing.prototype["GetHeight"]						= ApiDrawing.prototype.GetHeight;
	ApiDrawing.prototype["Select"]							= ApiDrawing.prototype.Select;
	ApiDrawing.prototype["SetRotation"]						= ApiDrawing.prototype.SetRotation;
	ApiDrawing.prototype["GetRotation"]						= ApiDrawing.prototype.GetRotation;

	// ApiShape
	ApiShape.prototype["GetClassType"]						= ApiShape.prototype.GetClassType;
	ApiShape.prototype["GetContent"]						= ApiShape.prototype.GetContent;
	ApiShape.prototype["SetVerticalTextAlign"]				= ApiShape.prototype.SetVerticalTextAlign;
	ApiShape.prototype["GetGeometry"]						= ApiShape.prototype.GetGeometry;
	ApiShape.prototype["SetGeometry"]						= ApiShape.prototype.SetGeometry;
	ApiShape.prototype["SetFill"]							= ApiShape.prototype.SetFill;
	ApiShape.prototype["GetFill"]							= ApiShape.prototype.GetFill;
	ApiShape.prototype["SetLine"]							= ApiShape.prototype.SetLine;
	ApiShape.prototype["GetLine"]							= ApiShape.prototype.GetLine;

	// ApiImage
	ApiImage.prototype["GetClassType"]						= ApiImage.prototype.GetClassType;

	// ApiTable
	ApiTable.prototype["GetClassType"]						= ApiTable.prototype.GetClassType;
	ApiTable.prototype["AddRow"]							= ApiTable.prototype.AddRow;
	ApiTable.prototype["GetRow"]							= ApiTable.prototype.GetRow;
	ApiTable.prototype["RemoveRow"]							= ApiTable.prototype.RemoveRow;
	ApiTable.prototype["MergeCells"]						= ApiTable.prototype.MergeCells;
	ApiTable.prototype["SetTableLook"]						= ApiTable.prototype.SetTableLook;
	ApiTable.prototype["AddColumn"]							= ApiTable.prototype.AddColumn;
	ApiTable.prototype["RemoveColumn"]						= ApiTable.prototype.RemoveColumn;
	ApiTable.prototype["SetShd"]							= ApiTable.prototype.SetShd;

	// ApiTableRow
	ApiTableRow.prototype["GetClassType"]					= ApiTableRow.prototype.GetClassType;
	ApiTableRow.prototype["GetCellsCount"]					= ApiTableRow.prototype.GetCellsCount;
	ApiTableRow.prototype["GetCell"]						= ApiTableRow.prototype.GetCell;
	ApiTableRow.prototype["SetHeight"]						= ApiTableRow.prototype.SetHeight;

	// ApiTableCell
	ApiTableCell.prototype["GetClassType"]					= ApiTableCell.prototype.GetClassType;
	ApiTableCell.prototype["GetContent"]					= ApiTableCell.prototype.GetContent;
	ApiTableCell.prototype["SetShd"]						= ApiTableCell.prototype.SetShd;
	ApiTableCell.prototype["SetCellMarginBottom"]			= ApiTableCell.prototype.SetCellMarginBottom;
	ApiTableCell.prototype["SetCellMarginLeft"]				= ApiTableCell.prototype.SetCellMarginLeft;
	ApiTableCell.prototype["SetCellMarginRight"]			= ApiTableCell.prototype.SetCellMarginRight;
	ApiTableCell.prototype["SetCellMarginTop"]				= ApiTableCell.prototype.SetCellMarginTop;
	ApiTableCell.prototype["SetCellBorderBottom"]			= ApiTableCell.prototype.SetCellBorderBottom;
	ApiTableCell.prototype["SetCellBorderLeft"]				= ApiTableCell.prototype.SetCellBorderLeft;
	ApiTableCell.prototype["SetCellBorderRight"]			= ApiTableCell.prototype.SetCellBorderRight;
	ApiTableCell.prototype["SetCellBorderTop"]				= ApiTableCell.prototype.SetCellBorderTop;
	ApiTableCell.prototype["SetVerticalAlign"]				= ApiTableCell.prototype.SetVerticalAlign;
	ApiTableCell.prototype["SetTextDirection"]				= ApiTableCell.prototype.SetTextDirection;

	// ApiChart
	ApiChart.prototype["GetClassType"]						= ApiChart.prototype.GetClassType = AscBuilder.ApiChart.prototype.GetClassType;
	ApiChart.prototype["GetChartType"]						= ApiChart.prototype.GetChartType = AscBuilder.ApiChart.prototype.GetChartType;
	ApiChart.prototype["SetTitle"]							= ApiChart.prototype.SetTitle = AscBuilder.ApiChart.prototype.SetTitle;
	ApiChart.prototype["SetHorAxisTitle"]					= ApiChart.prototype.SetHorAxisTitle = AscBuilder.ApiChart.prototype.SetHorAxisTitle;
	ApiChart.prototype["SetVerAxisTitle"]					= ApiChart.prototype.SetVerAxisTitle = AscBuilder.ApiChart.prototype.SetVerAxisTitle;
	ApiChart.prototype["SetVerAxisOrientation"]				= ApiChart.prototype.SetVerAxisOrientation = AscBuilder.ApiChart.prototype.SetVerAxisOrientation;
	ApiChart.prototype["SetHorAxisOrientation"]				= ApiChart.prototype.SetHorAxisOrientation = AscBuilder.ApiChart.prototype.SetHorAxisOrientation;
	ApiChart.prototype["SetLegendPos"]						= ApiChart.prototype.SetLegendPos = AscBuilder.ApiChart.prototype.SetLegendPos;
	ApiChart.prototype["SetLegendFontSize"]					= ApiChart.prototype.SetLegendFontSize = AscBuilder.ApiChart.prototype.SetLegendFontSize;
	ApiChart.prototype["SetShowDataLabels"]					= ApiChart.prototype.SetShowDataLabels = AscBuilder.ApiChart.prototype.SetShowDataLabels;
	ApiChart.prototype["SetShowPointDataLabel"]				= ApiChart.prototype.SetShowPointDataLabel = AscBuilder.ApiChart.prototype.SetShowPointDataLabel;
	ApiChart.prototype["SetVertAxisTickLabelPosition"]		= ApiChart.prototype.SetVertAxisTickLabelPosition = AscBuilder.ApiChart.prototype.SetVertAxisTickLabelPosition;
	ApiChart.prototype["SetHorAxisTickLabelPosition"]		= ApiChart.prototype.SetHorAxisTickLabelPosition = AscBuilder.ApiChart.prototype.SetHorAxisTickLabelPosition;
	ApiChart.prototype["SetHorAxisMajorTickMark"]			= ApiChart.prototype.SetHorAxisMajorTickMark = AscBuilder.ApiChart.prototype.SetHorAxisMajorTickMark;
	ApiChart.prototype["SetHorAxisMinorTickMark"]			= ApiChart.prototype.SetHorAxisMinorTickMark = AscBuilder.ApiChart.prototype.SetHorAxisMinorTickMark;
	ApiChart.prototype["SetVertAxisMajorTickMark"]			= ApiChart.prototype.SetVertAxisMajorTickMark = AscBuilder.ApiChart.prototype.SetVertAxisMajorTickMark;
	ApiChart.prototype["SetVertAxisMinorTickMark"]			= ApiChart.prototype.SetVertAxisMinorTickMark = AscBuilder.ApiChart.prototype.SetVertAxisMinorTickMark;
	ApiChart.prototype["SetMajorVerticalGridlines"]			= ApiChart.prototype.SetMajorVerticalGridlines = AscBuilder.ApiChart.prototype.SetMajorVerticalGridlines;
	ApiChart.prototype["SetMinorVerticalGridlines"]			= ApiChart.prototype.SetMinorVerticalGridlines = AscBuilder.ApiChart.prototype.SetMinorVerticalGridlines;
	ApiChart.prototype["SetMajorHorizontalGridlines"]		= ApiChart.prototype.SetMajorHorizontalGridlines = AscBuilder.ApiChart.prototype.SetMajorHorizontalGridlines;
	ApiChart.prototype["SetMinorHorizontalGridlines"]		= ApiChart.prototype.SetMinorHorizontalGridlines = AscBuilder.ApiChart.prototype.SetMinorHorizontalGridlines;
	ApiChart.prototype["SetHorAxisLabelsFontSize"]			= ApiChart.prototype.SetHorAxisLabelsFontSize = AscBuilder.ApiChart.prototype.SetHorAxisLabelsFontSize;
	ApiChart.prototype["SetVertAxisLabelsFontSize"]			= ApiChart.prototype.SetVertAxisLabelsFontSize = AscBuilder.ApiChart.prototype.SetVertAxisLabelsFontSize;
	ApiChart.prototype["RemoveSeria"]						= ApiChart.prototype.RemoveSeria = AscBuilder.ApiChart.prototype.RemoveSeria;
	ApiChart.prototype["SetSeriaValues"]					= ApiChart.prototype.SetSeriaValues = AscBuilder.ApiChart.prototype.SetSeriaValues;
	ApiChart.prototype["SetXValues"]						= ApiChart.prototype.SetXValues = AscBuilder.ApiChart.prototype.SetXValues;
	ApiChart.prototype["SetSeriaName"]						= ApiChart.prototype.SetSeriaName = AscBuilder.ApiChart.prototype.SetSeriaName;
	ApiChart.prototype["SetCategoryName"]					= ApiChart.prototype.SetCategoryName = AscBuilder.ApiChart.prototype.SetCategoryName;
	ApiChart.prototype["ApplyChartStyle"]					= ApiChart.prototype.ApplyChartStyle = AscBuilder.ApiChart.prototype.ApplyChartStyle;
	ApiChart.prototype["SetPlotAreaFill"]					= ApiChart.prototype.SetPlotAreaFill = AscBuilder.ApiChart.prototype.SetPlotAreaFill;
	ApiChart.prototype["SetPlotAreaOutLine"]				= ApiChart.prototype.SetPlotAreaOutLine = AscBuilder.ApiChart.prototype.SetPlotAreaOutLine;
	ApiChart.prototype["SetSeriesFill"]						= ApiChart.prototype.SetSeriesFill = AscBuilder.ApiChart.prototype.SetSeriesFill;
	ApiChart.prototype["SetSeriesOutLine"]					= ApiChart.prototype.SetSeriesOutLine = AscBuilder.ApiChart.prototype.SetSeriesOutLine;
	ApiChart.prototype["SetDataPointFill"]					= ApiChart.prototype.SetDataPointFill = AscBuilder.ApiChart.prototype.SetDataPointFill;
	ApiChart.prototype["SetDataPointOutLine"]				= ApiChart.prototype.SetDataPointOutLine = AscBuilder.ApiChart.prototype.SetDataPointOutLine;
	ApiChart.prototype["SetMarkerFill"]						= ApiChart.prototype.SetMarkerFill = AscBuilder.ApiChart.prototype.SetMarkerFill;
	ApiChart.prototype["SetMarkerOutLine"]					= ApiChart.prototype.SetMarkerOutLine = AscBuilder.ApiChart.prototype.SetMarkerOutLine;
	ApiChart.prototype["SetTitleFill"]						= ApiChart.prototype.SetTitleFill = AscBuilder.ApiChart.prototype.SetTitleFill;
	ApiChart.prototype["SetTitleOutLine"]					= ApiChart.prototype.SetTitleOutLine = AscBuilder.ApiChart.prototype.SetTitleOutLine;
	ApiChart.prototype["SetLegendFill"]						= ApiChart.prototype.SetLegendFill = AscBuilder.ApiChart.prototype.SetLegendFill;
	ApiChart.prototype["SetLegendOutLine"]					= ApiChart.prototype.SetLegendOutLine = AscBuilder.ApiChart.prototype.SetLegendOutLine;
	ApiChart.prototype["SetAxieNumFormat"]					= ApiChart.prototype.SetAxieNumFormat = AscBuilder.ApiChart.prototype.SetAxieNumFormat;
	ApiChart.prototype["SetSeriaNumFormat"]					= ApiChart.prototype.SetSeriaNumFormat = AscBuilder.ApiChart.prototype.SetSeriaNumFormat;
	ApiChart.prototype["SetDataPointNumFormat"]				= ApiChart.prototype.SetDataPointNumFormat = AscBuilder.ApiChart.prototype.SetDataPointNumFormat;
	ApiChart.prototype["GetAllSeries"]						= ApiChart.prototype.GetAllSeries = AscBuilder.ApiChart.prototype.GetAllSeries;
	ApiChart.prototype["GetSeries"]							= ApiChart.prototype.GetSeries = AscBuilder.ApiChart.prototype.GetSeries;
	
	// ApiGroup
	ApiGroup.prototype["GetClassType"]						= ApiGroup.prototype.GetClassType;
	
	// ApiSmartArt
	ApiSmartArt.prototype["GetClassType"]					= ApiSmartArt.prototype.GetClassType;
	
	window['AscBuilder']["Pdf"] = window['AscBuilder'].Pdf = window['AscBuilder'].Pdf || {};
	AscBuilder.Pdf["Api"] = AscBuilder.Pdf.Api = Api;
	
	AscBuilder.Pdf.init = function()
	{
		AscBuilder.ApiDrawing   = ApiDrawing;
		AscBuilder.ApiShape     = ApiShape;
		AscBuilder.ApiImage     = ApiImage;
		AscBuilder.ApiGroup     = ApiGroup;
		AscBuilder.ApiSmartArt  = ApiSmartArt;
		AscBuilder.ApiChart     = ApiChart;
	};

}(window, null));


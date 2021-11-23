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

"use strict";
(function (window, builder) {

    /**
     * @global
     * @class
     * @name Api
     */
    var Api = window["Asc"]["asc_docs_api"] || window["Asc"]["spreadsheet_api"];

    /**
     * Class representing a presentation.
     * @constructor
     */
    function ApiPresentation(oPresentation){
        this.Presentation = oPresentation;
    }

    /**
     * Class representing a slide master.
     * @constructor
     */
    function ApiMaster(oMaster){
        this.Master = oMaster;
    }

    /**
     * Class representing a slide layout.
     * @constructor
     */
    function ApiLayout(oLayout){
        this.Layout = oLayout;
    }

    /**
     * Class representing a placeholder.
     * @constructor
     */
    function ApiPlaceholder(oPh){
        this.Placeholder = oPh;
    }

    /**
     * Class representing a presentation theme.
     * @constructor
     */
    function ApiTheme(oThemeInfo){
        this.ThemeInfo = oThemeInfo;
    }
    
    /**
     * Class representing a theme's color scheme.
     * @constructor
     */
    function ApiThemeColorScheme(oClrScheme){
        this.ColorScheme = oClrScheme;
    }

    /**
     * Class representing a theme's format scheme.
     * @constructor
     */
    function ApiThemeFormatScheme(ofmtScheme){
        this.FormatScheme = ofmtScheme;
    }

    /**
     * Class representing a theme's font scheme.
     * @constructor
     */
    function ApiThemeFontScheme(ofontScheme){
        this.FontScheme = ofontScheme;
    }

    /**
     * Class representing a slide.
     * @constructor
     */
    function ApiSlide(oSlide){
        this.Slide = oSlide;
    }

    /**
     * Class representing a graphical object.
     * @constructor
     */
    function ApiDrawing(Drawing)
    {
        this.Drawing = Drawing;
    }

    /**
     * Class representing a shape.
     * @constructor
     */
    function ApiShape(oShape){
		ApiDrawing.call(this, oShape);
        this.Shape = oShape;
    }
	ApiShape.prototype = Object.create(ApiDrawing.prototype);
	ApiShape.prototype.constructor = ApiShape;

    /**
     * Class representing an image.
     * @constructor
     */
    function ApiImage(oImage){
		ApiDrawing.call(this, oImage);
    }
	ApiImage.prototype = Object.create(ApiDrawing.prototype);
	ApiImage.prototype.constructor = ApiImage;

    /**
     * Class representing a chart.
     * @constructor
     */
    function ApiChart(oChart){
		ApiDrawing.call(this, oChart);
        this.Chart = oChart;
    }
	ApiChart.prototype = Object.create(ApiDrawing.prototype);
	ApiChart.prototype.constructor    = ApiChart;

    /**
     * Class representing a group of drawings.
     * @constructor
     */
    function ApiGroup(oGroup){
		ApiDrawing.call(this, oGroup);
    }
	ApiGroup.prototype = Object.create(ApiDrawing.prototype);
	ApiGroup.prototype.constructor = ApiGroup;


	/**
     * Class representing a table.
     * @param oGraphicFrame
     * @constructor
     * */
	function ApiTable(oGraphicFrame){
	    this.Table = oGraphicFrame.graphicObject;
	    ApiDrawing.call(this, oGraphicFrame);
    }

    ApiTable.prototype = Object.create(ApiDrawing.prototype);
    ApiTable.prototype.constructor = ApiTable;




    /**
     * Class representing a table row.
     * @param oTableRow
     * @constructor
     */

    function ApiTableRow(oTableRow){
        this.Row = oTableRow;
    }


    /**
     * Class representing a table cell.
     * @param oCell
     * @constructor
     */
    function ApiTableCell(oCell){
        this.Cell = oCell;
    }


    /**
     * Twentieths of a point (equivalent to 1/1440th of an inch).
     * @typedef {number} twips
     */

    /**
     * 240ths of a line.
     * @typedef {number} line240
     */

    /**
     * Half-points (2 half-points = 1 point).
     * @typedef {number} hps
     */

    /**
     * A numeric value from 0 to 255.
     * @typedef {number} byte
     */

    /**
     * 60000th of a degree (5400000 = 90 degrees).
     * @typedef {number} PositiveFixedAngle
     * */

    /**
     * A border type.
     * @typedef {("none" | "single")} BorderType
     */

    /**
     * Types of custom tab.
     * @typedef {("clear" | "left" | "right" | "center")} TabJc
     */

    /**
     * Eighths of a point (24 eighths of a point = 3 points).
     * @typedef {number} pt_8
     */

    /**
     * A point.
     * @typedef {number} pt
     */


    /**
     * English measure unit. 1 mm = 36000 EMUs, 1 inch = 914400 EMUs.
     * @typedef {number} EMU
     */

    /**
     * This type specifies the preset shape geometry that will be used for a shape.
     * @typedef {("accentBorderCallout1" | "accentBorderCallout2" | "accentBorderCallout3" | "accentCallout1" | "accentCallout2" | "accentCallout3" | "actionButtonBackPrevious" | "actionButtonBeginning" | "actionButtonBlank" | "actionButtonDocument" | "actionButtonEnd" | "actionButtonForwardNext" | "actionButtonHelp" | "actionButtonHome" | "actionButtonInformation" | "actionButtonMovie" | "actionButtonReturn" | "actionButtonSound" | "arc" | "bentArrow" | "bentConnector2" | "bentConnector3" | "bentConnector4" | "bentConnector5" | "bentUpArrow" | "bevel" | "blockArc" | "borderCallout1" | "borderCallout2" | "borderCallout3" | "bracePair" | "bracketPair" | "callout1" | "callout2" | "callout3" | "can" | "chartPlus" | "chartStar" | "chartX" | "chevron" | "chord" | "circularArrow" | "cloud" | "cloudCallout" | "corner" | "cornerTabs" | "cube" | "curvedConnector2" | "curvedConnector3" | "curvedConnector4" | "curvedConnector5" | "curvedDownArrow" | "curvedLeftArrow" | "curvedRightArrow" | "curvedUpArrow" | "decagon" | "diagStripe" | "diamond" | "dodecagon" | "donut" | "doubleWave" | "downArrow" | "downArrowCallout" | "ellipse" | "ellipseRibbon" | "ellipseRibbon2" | "flowChartAlternateProcess" | "flowChartCollate" | "flowChartConnector" | "flowChartDecision" | "flowChartDelay" | "flowChartDisplay" | "flowChartDocument" | "flowChartExtract" | "flowChartInputOutput" | "flowChartInternalStorage" | "flowChartMagneticDisk" | "flowChartMagneticDrum" | "flowChartMagneticTape" | "flowChartManualInput" | "flowChartManualOperation" | "flowChartMerge" | "flowChartMultidocument" | "flowChartOfflineStorage" | "flowChartOffpageConnector" | "flowChartOnlineStorage" | "flowChartOr" | "flowChartPredefinedProcess" | "flowChartPreparation" | "flowChartProcess" | "flowChartPunchedCard" | "flowChartPunchedTape" | "flowChartSort" | "flowChartSummingJunction" | "flowChartTerminator" | "foldedCorner" | "frame" | "funnel" | "gear6" | "gear9" | "halfFrame" | "heart" | "heptagon" | "hexagon" | "homePlate" | "horizontalScroll" | "irregularSeal1" | "irregularSeal2" | "leftArrow" | "leftArrowCallout" | "leftBrace" | "leftBracket" | "leftCircularArrow" | "leftRightArrow" | "leftRightArrowCallout" | "leftRightCircularArrow" | "leftRightRibbon" | "leftRightUpArrow" | "leftUpArrow" | "lightningBolt" | "line" | "lineInv" | "mathDivide" | "mathEqual" | "mathMinus" | "mathMultiply" | "mathNotEqual" | "mathPlus" | "moon" | "nonIsoscelesTrapezoid" | "noSmoking" | "notchedRightArrow" | "octagon" | "parallelogram" | "pentagon" | "pie" | "pieWedge" | "plaque" | "plaqueTabs" | "plus" | "quadArrow" | "quadArrowCallout" | "rect" | "ribbon" | "ribbon2" | "rightArrow" | "rightArrowCallout" | "rightBrace" | "rightBracket" | "round1Rect" | "round2DiagRect" | "round2SameRect" | "roundRect" | "rtTriangle" | "smileyFace" | "snip1Rect" | "snip2DiagRect" | "snip2SameRect" | "snipRoundRect" | "squareTabs" | "star10" | "star12" | "star16" | "star24" | "star32" | "star4" | "star5" | "star6" | "star7" | "star8" | "straightConnector1" | "stripedRightArrow" | "sun" | "swooshArrow" | "teardrop" | "trapezoid" | "triangle" | "upArrowCallout" | "upDownArrow" | "upDownArrow" | "upDownArrowCallout" | "uturnArrow" | "verticalScroll" | "wave" | "wedgeEllipseCallout" | "wedgeRectCallout" | "wedgeRoundRectCallout")} ShapeType
     */

    /**
    * A bullet type which will be added to the paragraph in spreadsheet or presentation.
    * @typedef {("None" | "ArabicPeriod"  | "ArabicParenR"  | "RomanUcPeriod" | "RomanLcPeriod" | "AlphaLcParenR" | "AlphaLcPeriod" | "AlphaUcParenR" | "AlphaUcPeriod")} BulletType
    */


    /**
     * This type specifies the available chart types which can be used to create a new chart.
     * @typedef {("bar" | "barStacked" | "barStackedPercent" | "bar3D" | "barStacked3D" | "barStackedPercent3D" | "barStackedPercent3DPerspective" | "horizontalBar" | "horizontalBarStacked" | "horizontalBarStackedPercent" | "horizontalBar3D" | "horizontalBarStacked3D" | "horizontalBarStackedPercent3D" | "lineNormal" | "lineStacked" | "lineStackedPercent" | "line3D" | "pie" | "pie3D" | "doughnut" | "scatter" | "stock" | "area" | "areaStacked" | "areaStackedPercent")} ChartType
     */

    /**
     * The available text vertical alignment (used to align text in a shape with a placement for text inside it).
     * @typedef {("top" | "center" | "bottom")} VerticalTextAlign
     * */

    /**
     * The available color scheme identifiers.
     * @typedef {("accent1" | "accent2" | "accent3" | "accent4" | "accent5" | "accent6" | "bg1" | "bg2" | "dk1" | "dk2" | "lt1" | "lt2" | "tx1" | "tx2")} SchemeColorId
     * */

    /**
     * The available preset color names.
     * @typedef {("aliceBlue" | "antiqueWhite" | "aqua" | "aquamarine" | "azure" | "beige" | "bisque" | "black" | "blanchedAlmond" | "blue" | "blueViolet" | "brown" | "burlyWood" | "cadetBlue" | "chartreuse" | "chocolate" | "coral" | "cornflowerBlue" | "cornsilk" | "crimson" | "cyan" | "darkBlue" | "darkCyan" | "darkGoldenrod" | "darkGray" | "darkGreen" | "darkGrey" | "darkKhaki" | "darkMagenta" | "darkOliveGreen" | "darkOrange" | "darkOrchid" | "darkRed" | "darkSalmon" | "darkSeaGreen" | "darkSlateBlue" | "darkSlateGray" | "darkSlateGrey" | "darkTurquoise" | "darkViolet" | "deepPink" | "deepSkyBlue" | "dimGray" | "dimGrey" | "dkBlue" | "dkCyan" | "dkGoldenrod" | "dkGray" | "dkGreen" | "dkGrey" | "dkKhaki" | "dkMagenta" | "dkOliveGreen" | "dkOrange" | "dkOrchid" | "dkRed" | "dkSalmon" | "dkSeaGreen" | "dkSlateBlue" | "dkSlateGray" | "dkSlateGrey" | "dkTurquoise" | "dkViolet" | "dodgerBlue" | "firebrick" | "floralWhite" | "forestGreen" | "fuchsia" | "gainsboro" | "ghostWhite" | "gold" | "goldenrod" | "gray" | "green" | "greenYellow" | "grey" | "honeydew" | "hotPink" | "indianRed" | "indigo" | "ivory" | "khaki" | "lavender" | "lavenderBlush" | "lawnGreen" | "lemonChiffon" | "lightBlue" | "lightCoral" | "lightCyan" | "lightGoldenrodYellow" | "lightGray" | "lightGreen" | "lightGrey" | "lightPink" | "lightSalmon" | "lightSeaGreen" | "lightSkyBlue" | "lightSlateGray" | "lightSlateGrey" | "lightSteelBlue" | "lightYellow" | "lime" | "limeGreen" | "linen" | "ltBlue" | "ltCoral" | "ltCyan" | "ltGoldenrodYellow" | "ltGray" | "ltGreen" | "ltGrey" | "ltPink" | "ltSalmon" | "ltSeaGreen" | "ltSkyBlue" | "ltSlateGray" | "ltSlateGrey" | "ltSteelBlue" | "ltYellow" | "magenta" | "maroon" | "medAquamarine" | "medBlue" | "mediumAquamarine" | "mediumBlue" | "mediumOrchid" | "mediumPurple" | "mediumSeaGreen" | "mediumSlateBlue" | "mediumSpringGreen" | "mediumTurquoise" | "mediumVioletRed" | "medOrchid" | "medPurple" | "medSeaGreen" | "medSlateBlue" | "medSpringGreen" | "medTurquoise" | "medVioletRed" | "midnightBlue" | "mintCream" | "mistyRose" | "moccasin" | "navajoWhite" | "navy" | "oldLace" | "olive" | "oliveDrab" | "orange" | "orangeRed" | "orchid" | "paleGoldenrod" | "paleGreen" | "paleTurquoise" | "paleVioletRed" | "papayaWhip" | "peachPuff" | "peru" | "pink" | "plum" | "powderBlue" | "purple" | "red" | "rosyBrown" | "royalBlue" | "saddleBrown" | "salmon" | "sandyBrown" | "seaGreen" | "seaShell" | "sienna" | "silver" | "skyBlue" | "slateBlue" | "slateGray" | "slateGrey" | "snow" | "springGreen" | "steelBlue" | "tan" | "teal" | "thistle" | "tomato" | "turquoise" | "violet" | "wheat" | "white" | "whiteSmoke" | "yellow" | "yellowGreen")} PresetColor
     * */


    /**
     * Possible values for the position of chart tick labels (either horizontal or vertical).
     * * **"none"** - not display the selected tick labels.
     * * **"nextTo"** - set the position of the selected tick labels next to the main label.
     * * **"low"** - set the position of the selected tick labels in the part of the chart with lower values.
     * * **"high"** - set the position of the selected tick labels in the part of the chart with higher values.
     * @typedef {("none" | "nextTo" | "low" | "high")} TickLabelPosition
     * **/

    /**
     * The type of a fill which uses an image as a background.
     * * **"tile"** - if the image is smaller than the shape which is filled, the image will be tiled all over the created shape surface.
     * * **"stretch"** - if the image is smaller than the shape which is filled, the image will be stretched to fit the created shape surface.
     * @typedef {"tile" | "stretch"} BlipFillType
     * */

    /**
     * The available preset patterns which can be used for the fill.
     * @typedef {"cross" | "dashDnDiag" | "dashHorz" | "dashUpDiag" | "dashVert" | "diagBrick" | "diagCross" | "divot" | "dkDnDiag" | "dkHorz" | "dkUpDiag" | "dkVert" | "dnDiag" | "dotDmnd" | "dotGrid" | "horz" | "horzBrick" | "lgCheck" | "lgConfetti" | "lgGrid" | "ltDnDiag" | "ltHorz" | "ltUpDiag" | "ltVert" | "narHorz" | "narVert" | "openDmnd" | "pct10" | "pct20" | "pct25" | "pct30" | "pct40" | "pct5" | "pct50" | "pct60" | "pct70" | "pct75" | "pct80" | "pct90" | "plaid" | "shingle" | "smCheck" | "smConfetti" | "smGrid" | "solidDmnd" | "sphere" | "trellis" | "upDiag" | "vert" | "wave" | "wdDnDiag" | "wdUpDiag" | "weave" | "zigZag"} PatternType
     * */



    /**
     * The available types of tick mark appearance.
     * @typedef {("cross" | "in" | "none" | "out")} TickMark
     * */

    //------------------------------------------------------------------------------------------------------------------
    //
    // Base Api
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * The 1000th of a percent (100000 = 100%).
     * @typedef {number} PositivePercentage
     * */

    /**
     * Get the main presentation.
     * @typeofeditors ["CPE"]
     * @memberof Api
     * @returns {ApiPresentation}
     */
    Api.prototype.GetPresentation = function(){
        if(this.WordControl && this.WordControl.m_oLogicDocument){
            return new ApiPresentation(this.WordControl.m_oLogicDocument);
        }
        return null;
    };

    /**
     * Create a new slide master.
     * @typeofeditors ["CPE"]
     * @memberof Api
     * @param {ApiTheme} [oTheme    = ApiPresentation.GetMaster(0).GetTheme()] - ApiTheme object.
     * @returns {ApiMaster | null} - returns null if presentation theme doesn't exist.
     */
    Api.prototype.CreateMaster = function(oTheme){
        if (!oTheme || !oTheme.GetClassType || !oTheme.GetClassType() === "theme")
            if (editor.GetPresentation().GetMaster(0))
                oTheme = editor.GetPresentation().GetMaster(0).GetTheme();
        
        if (!oTheme)
            return null;

        var oThemeCopy = oTheme.ThemeInfo.Theme.createDuplicate();
        var oMaster = new ApiMaster(new AscCommonSlide.MasterSlide());
        oMaster.Master.setTheme(oThemeCopy);

        return oMaster;
    };

    /**
     * Create a new slide layout and adds it to the slide master if it specified.
     * @typeofeditors ["CPE"]
     * @memberof Api
     * @param {ApiMaster} [oMaster = null] - parent slide master.
     * @returns {ApiLayout}
     */
    Api.prototype.CreateLayout = function(oMaster){
        var oLayout = new ApiLayout(new AscCommonSlide.SlideLayout());

        if (oMaster && oMaster.GetClassType && oMaster.GetClassType() === "master")
            oMaster.AddLayout(undefined, oLayout);
        
        return oLayout;
    };

    /**
     * Create a new placeholder.
     * @typeofeditors ["CPE"]
     * @memberof Api
     * @param {string} sType - the placeholder type.
     * @returns {ApiPlaceholder}
     */
    Api.prototype.CreatePlaceholder = function(sType){
        
        if (typeof(sType) !== "string")
            sType = "body";

        var oPh = new ApiPlaceholder(new AscFormat.Ph());
        oPh.SetType(sType);

        return oPh;
    };

    /**
     * Create a new theme.
     * @typeofeditors ["CPE"]
     * @memberof Api
     * @param {string} sName - specified theme's name.
     * @param {ApiMaster} oMaster - slide master. Required parameter.
     * @param {ApiThemeColorScheme} oClrScheme - color scheme. Required parameter.
     * @param {ApiThemeFormatScheme} oFormatScheme - format scheme. Required parameter.
     * @param {ApiThemeFontScheme} oFontScheme - font scheme. Required parameter.
     * @returns {ApiTheme | null} 
     */
    Api.prototype.CreateTheme = function(sName, oMaster, oClrScheme, oFormatScheme, oFontScheme){
        if (typeof(sName) !== "string")
            sName = "";
        if (oMaster.GetClassType() !== "master" || oClrScheme.GetClassType() !== "themeColorScheme" ||
        oFormatScheme.GetClassType() !== "themeFormatScheme" || oFontScheme.GetClassType() !== "themeFontScheme")
            return null;

        var oPresentation      = editor.GetPresentation().Presentation;
        var oThemeLoadInfo     = new AscCommonSlide.CThemeLoadInfo();
        oThemeLoadInfo.Master  = oMaster.Master;
        oThemeLoadInfo.Layouts = oMaster.Master.sldLayoutLst;
        
        var oTheme = new AscFormat.CTheme();
        oTheme.setName(sName);

        var presentation              = {};
        presentation.ImageMap         = {};
        presentation.Fonts            = [];
        presentation.Masters          = [oMaster.Master];
        presentation.DrawingDocument  = editor.WordControl.m_oDrawingDocument;
        presentation.pres             = oPresentation;
        presentation.Width            = oPresentation.Width;
        presentation.Height           = oPresentation.Height;
        presentation.defaultTextStyle = oPresentation.defaultTextStyle;

        oTheme.presentation             = presentation;
        oTheme.themeElements.clrScheme  = oClrScheme.ColorScheme;
        oTheme.themeElements.fmtScheme  = oFormatScheme.FormatScheme;
        oTheme.themeElements.fontScheme = oFontScheme.FontScheme;
        oThemeLoadInfo.Theme            = oTheme;

        oMaster.Master.setTheme(oTheme);

        return new ApiTheme(oThemeLoadInfo);
    };

    /**
     * Create a new theme color scheme.
     * @typeofeditors ["CPE"]
     * @memberof Api
     * @param {string} sName - specified name of theme's color scheme.
     * @param {(ApiUniColor | ApiRGBColor)[]} arrColors - This array defines a set of colors which are referred to as a color scheme.
     * The color scheme is responsible for defining a list of twelve colors.
     * The array should contain a sequence of colors: 2 dark, 2 light,6 primary and a color for each of a hyperlink and followed hyperlink.     
     * @returns {?ApiThemeColorScheme} 
     */
    Api.prototype.CreateThemeColorScheme = function(arrColors, sName){
        if (typeof(sName) !== "string")
            sName = "New theme's color scheme";
        if (!Array.isArray(arrColors) || arrColors.length !== 12)
            return null;

        var oClrScheme = new AscFormat.ClrScheme();
        oClrScheme.setName(sName);

        oClrScheme.addColor(0, arrColors[0].Unicolor);   
        oClrScheme.addColor(2, arrColors[2].Unicolor); 
        oClrScheme.addColor(1, arrColors[1].Unicolor);   
        oClrScheme.addColor(3, arrColors[3].Unicolor);
        oClrScheme.addColor(4, arrColors[4].Unicolor);
        oClrScheme.addColor(5, arrColors[5].Unicolor);

        for (var nColor = 6; nColor < 12; nColor++)
        {
            oClrScheme.addColor(nColor + 2, arrColors[nColor].Unicolor);   
        }
        
        return new ApiThemeColorScheme(oClrScheme);
    };

    /**
     * Create a new theme format scheme.
     * @typeofeditors ["CPE"]
     * @memberof Api
     * @param {ApiFill[]} arrFill - This array contains fill styles. Should be composed of subtle, moderate and intense fills.
     * @param {ApiFill[]} arrBgFill - This array contains the background fill styles. Should be composed of subtle, moderate and intense fills.
     * @param {ApiStroke[]} arrLine - This array contains line styles. Should be composed of subtle, moderate and intense lines.
     * @param {string} sName
     * @returns {?ApiThemeFormatScheme} 
     */
    Api.prototype.CreateThemeFormatScheme = function(arrFill, arrBgFill, arrLine, sName){
        
        if (typeof(sName) !== "string")
            sName = "New format scheme";

        if (Array.isArray(arrFill) && Array.isArray(arrBgFill) && Array.isArray(arrLine) && 
        arrFill.length === 3 && arrBgFill.length === 3 && arrLine.length === 3) 
        {
            var oFormatScheme = new ApiThemeFormatScheme(new AscFormat.FmtScheme());
        
            oFormatScheme.SetSchemeName(sName);
            oFormatScheme.ChangeFillStyles(arrFill);
            oFormatScheme.ChangeBgFillStyles(arrBgFill);
            oFormatScheme.ChangeLineStyles(arrLine);
            //oFormatScheme.ChangeEffectStyles(arrEffect);
    
            return oFormatScheme;
        }

        return null;
    };

    
    /**
     * Create a new theme font scheme.
     * @typeofeditors ["CPE"]
     * @memberof Api
     * @param {string} mjLatin - Font's name. Is used to determine the major theme font applied to latit text.
     * @param {string} mjEa - Font's name. Is used to determine the major theme font applied to east asian text.
     * @param {string} mjCs - Font's name. Is used to determine the major theme font applied to complex script text.
     * @param {string} mnLatin - Font's name. Is used to determine the minor theme font applied to latit text.
     * @param {string} mnEa - Font's name. Is used to determine the minor theme font applied to east asian text.
     * @param {string} mnCs - Font's name. Is used to determine the minor theme font applied to complex script text.
     * @param {string} sName - name of this font scheme.
     * @returns {ApiThemeFontScheme} 
     */
    Api.prototype.CreateThemeFontScheme = function(mjLatin, mjEa, mjCs, mnLatin, mnEa, mnCs, sName){
        
        if (typeof(sName) !== "string")
            sName = "New format scheme";

        var oFontScheme          = new AscFormat.FontScheme();
        var oMajorFontCollection = new AscFormat.FontCollection(oFontScheme);
        var oMinorFontCollection = new AscFormat.FontCollection(oFontScheme);

        oFontScheme.setName(sName);
        oFontScheme.setMajorFont(oMajorFontCollection);
        oFontScheme.setMinorFont(oMinorFontCollection);

        oMajorFontCollection.setLatin(mjLatin);
        oMajorFontCollection.setEA(mjEa);
        oMajorFontCollection.setCS(mjCs);

        oMinorFontCollection.setLatin(mnLatin);
        oMinorFontCollection.setEA(mnEa);
        oMinorFontCollection.setCS(mnCs);

        return new ApiThemeFontScheme(oFontScheme);
    };
    
    /**
     * Create a new slide.
     * @typeofeditors ["CPE"]
     * @memberof Api
     * @returns {ApiSlide}
     */
    Api.prototype.CreateSlide = function(){
        var oPresentation = private_GetPresentation();
        var oSlide = new AscCommonSlide.Slide(oPresentation, oPresentation.slideMasters[0].sldLayoutLst[0], 0);
        oSlide.setNotes(AscCommonSlide.CreateNotes());
        oSlide.notes.setNotesMaster(oPresentation.notesMasters[0]);
        oSlide.notes.setSlide(oSlide);
        return new ApiSlide(oSlide);
    };

    /**
     * Create an image with the parameters specified.
     * @memberof Api
     * @typeofeditors ["CPE"]
     * @param {string} sImageSrc - The image source where the image to be inserted should be taken from (currently
     * only internet URL or Base64 encoded images are supported).
     * @param {EMU} nWidth - The image width in English measure units.
     * @param {EMU} nHeight - The image height in English measure units.
     * @returns {ApiImage}
     */
    Api.prototype.CreateImage = function(sImageSrc, nWidth, nHeight){
        var oImage = AscFormat.DrawingObjectsController.prototype.createImage(sImageSrc, 0, 0, nWidth/36000, nHeight/36000);
        oImage.setParent(private_GetCurrentSlide());
        return new ApiImage(AscFormat.DrawingObjectsController.prototype.createImage(sImageSrc, 0, 0, nWidth/36000, nHeight/36000));
    };

    /**
     * Create a shape with the parameters specified.
     * @memberof Api
     * @typeofeditors ["CPE"]
     * @param {ShapeType} [sType="rect"] - The shape type which specifies the preset shape geometry.
     * @param {EMU} [nWidth = 914400] - The shape width in English measure units.
	 * @param {EMU} [nHeight = 914400] - The shape height in English measure units.
	 * @param {ApiFill} [oFill    = Api.CreateNoFill()] - The color or pattern used to fill the shape.
	 * @param {ApiStroke} [oStroke    = Api.CreateStroke(0, Api.CreateNoFill())] - The stroke used to create the element shadow.
     * @returns {ApiShape}
     * */
    Api.prototype.CreateShape = function(sType, nWidth, nHeight, oFill, oStroke){
        var oCurrentSlide = private_GetCurrentSlide();
        sType   = sType   || "rect";
        nWidth  = nWidth  || 914400;
	    nHeight = nHeight || 914400;
	    oFill   = oFill   || editor.CreateNoFill();
	    oStroke = oStroke || editor.CreateStroke(0, editor.CreateNoFill());
        var oTheme = oCurrentSlide && oCurrentSlide.Layout && oCurrentSlide.Layout.Master && oCurrentSlide.Layout.Master.Theme;
        return new ApiShape(AscFormat.builder_CreateShape(sType, nWidth/36000, nHeight/36000, oFill.UniFill, oStroke.Ln, oCurrentSlide, oTheme, private_GetDrawingDocument(), false));
    };

    
    /**
     * Create a chart with the parameters specified.
     * @memberof Api
     * @typeofeditors ["CPE"]
     * @param {ChartType} [sType="bar"] - The chart type used for the chart display.
     * @param {Array} aSeries - The array of the data used to build the chart from.
     * @param {Array} aSeriesNames - The array of the names (the source table column names) used for the data which the chart will be build from.
     * @param {Array} aCatNames - The array of the names (the source table row names) used for the data which the chart will be build from.
     * @param {EMU} nWidth - The chart width in English measure units.
     * @param {EMU} nHeight - The chart height in English measure units.
     * @param {number} nStyleIndex - The chart color style index (can be <b>1 - 48</b>, as described in OOXML specification).
     * @returns {ApiChart}
     * */
    Api.prototype.CreateChart = function(sType, aSeries, aSeriesNames, aCatNames, nWidth, nHeight, nStyleIndex)
    {
        var oChartSpace = AscFormat.builder_CreateChart(nWidth/36000, nHeight/36000, sType, aCatNames, aSeriesNames, aSeries, nStyleIndex);
        oChartSpace.setParent(private_GetCurrentSlide());
        return new ApiChart(oChartSpace);
    };


    /**
     * Create a group of drawings.
     * @memberof Api
     * @param {Array} aDrawings - The array of drawings.
     * @returns {ApiGroup}
     * */
    Api.prototype.CreateGroup = function(aDrawings){
        var oSlide = private_GetCurrentSlide();
        if(oSlide){
            var oGroup = AscFormat.builder_CreateGroup(aDrawings, oSlide.graphicObjects);
            if(oGroup){
                return new ApiGroup(oGroup);
            }
        }
        return null;
    };


    /**
     * Create a table.
     * @param nCols - Number of columns.
     * @param nRows - Number of raws.
     * @returns {?ApiTable}
     */
    Api.prototype.CreateTable = function(nCols, nRows){
        var oPresentation = private_GetPresentation();
        var oSlide = private_GetCurrentSlide();
        if(oPresentation && oSlide){
            var oGraphicFrame = oPresentation.Create_TableGraphicFrame(nCols, nRows, oSlide, oPresentation.DefaultTableStyleId);
            var content = oGraphicFrame.graphicObject.Content, i, j;
            for(i = 0; i < content.length; ++i)
            {
                content[i].Set_Height(0, Asc.linerule_AtLeast );
            }
            return new ApiTable(oGraphicFrame);
        }
        return null;
    };

    /**
     * Create a new paragraph.
     * @memberof Api
     * @typeofeditors ["CPE"]
     * @returns {ApiParagraph}
     */
    Api.prototype.CreateParagraph = function()
    {
        return this.private_CreateApiParagraph(new Paragraph(private_GetDrawingDocument(), null, true));
    };

    /**
	 * Save changes to the specified document.
	 * @typeofeditors ["CPE"]
	 * @memberof Api
	 */
	Api.prototype.Save = function () {
		this.SaveAfterMacros = true;
	};

    Api.prototype.private_checkDrawingUniNvPr = function(oDrawing)
    {
        var nv_sp_pr;
        var drawing = oDrawing.Drawing;

        if (drawing)
        {
            switch (drawing.getObjectType())
            {
                case AscDFH.historyitem_type_ChartSpace:
                case AscDFH.historyitem_type_GraphicFrame:
                { 
                    if(!drawing.nvGraphicFramePr)
                    {
                        nv_sp_pr = new AscFormat.UniNvPr();
                        nv_sp_pr.cNvPr.setId(++this.maxId);
                        drawing.setNvSpPr(nv_sp_pr);
                    }
                    break;
                }
                case AscDFH.historyitem_type_GroupShape:
                {
                    if(!drawing.nvGrpSpPr)
                    {
                        nv_sp_pr = new AscFormat.UniNvPr();
                        nv_sp_pr.cNvPr.setId(++this.maxId);
                        drawing.setNvSpPr(nv_sp_pr);
                    }
                    for(var i = 0; i < drawing.spTree.length; ++i)
                    {
                        this.checkDrawingUniNvPr(drawing.spTree[i]);
                    }
                    break;
                }
                case AscDFH.historyitem_type_ImageShape:
                case AscDFH.historyitem_type_OleObject:
                {
                    if(!drawing.nvPicPr)
                    {
                        nv_sp_pr = new AscFormat.UniNvPr();
                        nv_sp_pr.cNvPr.setId(++this.maxId);
                        drawing.setNvSpPr(nv_sp_pr);
                    }
                    break;
                }
                case AscDFH.historyitem_type_Shape:
                case AscDFH.historyitem_type_Cnx:
                {
                    if(!drawing.nvSpPr)
                    {
                        nv_sp_pr = new AscFormat.UniNvPr();
                        nv_sp_pr.cNvPr.setId(++this.maxId);
                        drawing.setNvSpPr(nv_sp_pr);
                    }
                    break;
                }
            }
        }
    };

    /**
	 * Checks for duplicate placeholders and sets idx.
     * Called when a placeholder is added to a shape.
	 * @typeofeditors ["CPE"]
	 * @memberof Api
     * @param {ApiSlide | ApiLayout | ApiMaster} object - object in which placeholders will be checked.
     * @param {ApiPlaceholder}  - placeholder to be added. 
     * @return {bool} - return false if object insupported or oPlaceholder isn't placeholder.
	 */
    Api.prototype.private_checkPlaceholders = function(object, oPlaceholder)
    {
        if (object.GetClassType() !== "slide" && object.GetClassType() !== "layout" && object.GetClassType() !== "master" )
            return false;
        if (!oPlaceholder || !oPlaceholder.GetClassType || oPlaceholder.GetClassType() !== "placeholder")
            return false;
        
        var allDrawings          = object.GetAllDrawings();
        var maxIndex             = 0;
        var oTempPlaceholder     = null;
        var noIdxPlaceholders    = [];

        for (var nDrawing = 0; nDrawing < allDrawings.length; nDrawing++)
        {
            oTempPlaceholder = allDrawings[nDrawing].GetPlaceholder();

            if (oTempPlaceholder.Placeholder.type === oPlaceholder.Placeholder.type)
            {
                if (oTempPlaceholder.Placeholder.idx && +oTempPlaceholder.Placeholder.idx > maxIndex)
                    maxIndex  = +oTempPlaceholder.Placeholder.idx;
                else if (!oTempPlaceholder.Placeholder.idx)
                    noIdxPlaceholders.push(oTempPlaceholder.Placeholder);
            }
        }
    
        for (var nPh = 0; nPh < noIdxPlaceholders.length; nPh++)
        {
            noIdxPlaceholders[nPh].idx = String(maxIndex + 1);
            maxIndex++;
        };

        return true;
    };

    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiPresentation
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * Get the type of this class.
     * @typeofeditors ["CPE"]
     * @returns {"presentation"}
     */
    ApiPresentation.prototype.GetClassType = function()
    {
        return "presentation";
    };

    /**
     * Get the index for the current slide.
     * @typeofeditors ["CPE"]
     * @memberof ApiPresentation
     * @returns {number}
     */
    ApiPresentation.prototype.GetCurSlideIndex = function(){
        if(this.Presentation){
            return this.Presentation.CurPage;
        }
        return -1;
    };


    /**
     * Get the slide by its position in the presentation.
     * @memberof ApiPresentation
     * @param {number} nIndex - The slide number (position) in the presentation.
     * @returns {?ApiSlide}
     */
    ApiPresentation.prototype.GetSlideByIndex = function(nIndex){
        if(this.Presentation && this.Presentation.Slides[nIndex]){
            return new ApiSlide(this.Presentation.Slides[nIndex]);
        }
        return null;
    };

    /**
     * Get the current slide.
     * @typeofeditors ["CPE"]
     * @memberof ApiPresentation
     * @returns {?ApiSlide}
     */
    ApiPresentation.prototype.GetCurrentSlide = function () {
        return this.GetSlideByIndex(this.GetCurSlideIndex());
    };


    /**
     * Append a new slide to the end of the presentation.
     * @typeofeditors ["CPE"]
     * @memberof ApiPresentation
     * @param {ApiSlide} oSlide - The slide created using the {@link Api#CreateSlide} method.
     */
    ApiPresentation.prototype.AddSlide = function(oSlide) {
        if(this.Presentation){
            oSlide.Slide.setSlideNum(this.Presentation.Slides.length);
            this.Presentation.insertSlide(this.Presentation.Slides.length, oSlide.Slide);
        }
    };



    /**
     * Set the size for the current presentation.
     * @typeofeditors ["CPE"]
     * @memberof ApiPresentation
     * @param {EMU} nWidth - The presentation width in English measure units.
     * @param {EMU} nHeight - The presentation height in English measure units.
     */
    ApiPresentation.prototype.SetSizes = function(nWidth, nHeight) {
        if(this.Presentation){
            this.Presentation.internalChangeSizes(nWidth, nHeight);
        }
    };

    /**
     * Create new history point.
     */
    ApiPresentation.prototype.CreateNewHistoryPoint = function()
    {
        this.Presentation.Create_NewHistoryPoint(AscDFH.historydescription_Document_ApiBuilder);
    };


    /**
     * Replace current image.
     * @typeofeditors ["CPE"]
     * @memberof ApiPresentation
     * @param {string} sImageUrl - The image source where the image to be inserted should be taken from (currently only internet URL or Base64 encoded images are supported).
     * @param {EMU} Width - The image width in English measure units.
     * @param {EMU} Height - The image height in English measure units.
     */
    ApiPresentation.prototype.ReplaceCurrentImage = function(sImageUrl, Width, Height)
    {
        var oPr = this.Presentation;
        if(oPr.Slides[oPr.CurPage]){
            var _slide = oPr.Slides[oPr.CurPage];
            var oController = _slide.graphicObjects;
            var _w = Width/36000.0;
            var _h = Height/36000.0;
            var oImage = oController.createImage(sImageUrl, 0, 0, _w, _h);
            oImage.setParent(_slide);
            var selectedObjects, spTree;
            if(oController.selection.groupSelection){
                selectedObjects = oController.selection.groupSelection.selectedObjects;
            }
            else{
                selectedObjects = oController.selectedObjects;
            }
            if(selectedObjects.length > 0 && !oController.getTargetDocContent()){
                if(selectedObjects[0].group){
                    spTree = selectedObjects[0].group.spTree;
                }
                else{
                    spTree = _slide.cSld.spTree;
                }
                for(var i = 0; i < spTree.length; ++i){
                    if(spTree[i] === selectedObjects[0]){
                        var _xfrm = spTree[i].spPr && spTree[i].spPr.xfrm;
                        var _xfrm2 = oImage.spPr.xfrm;
                        if(_xfrm){
                            _xfrm2.setOffX(_xfrm.offX);
                            _xfrm2.setOffY(_xfrm.offY);
                            //_xfrm2.setRot(_xfrm.rot);
                        }
                        else{
                            if(AscFormat.isRealNumber(spTree[i].x) && AscFormat.isRealNumber(spTree[i].y)){
                                _xfrm2.setOffX(spTree[i].x);
                                _xfrm2.setOffY(spTree[i].y);
                            }
                        }
                        if(selectedObjects[0].group){
                            var _group = selectedObjects[0].group;
                            _group.removeFromSpTreeByPos(i);
                            _group.addToSpTree(i, oImage);
                            oImage.setGroup(_group);
                            oController.selection.groupSelection.resetInternalSelection();
                            _group.selectObject(oImage, oPr.CurPage);
                        }
                        else{
                            _slide.removeFromSpTreeByPos(i);
                            _slide.addToSpTreeToPos(i, oImage);
                            oController.resetSelection();
                            oController.selectObject(oImage, oPr.CurPage);
                        }
                        return;
                    }
                }
            }
            var _x = (this.Presentation.GetWidthMM() - _w)/2.0;
            var _y = (this.Presentation.GetHeightMM() - _h)/2.0;
            oImage.spPr.xfrm.setOffX(_x);
            oImage.spPr.xfrm.setOffY(_y);
            _slide.addToSpTreeToPos(_slide.cSld.spTree.length, oImage);
            oController.resetSelection();
            oController.selectObject(oImage, oPr.CurPage);
        }
    };

    /**
	 * Specify the languages which will be used to check spelling and grammar (if requested).
	 * @memberof ApiPresentation
	 * @typeofeditors ["CDE"]
	 * @param {string} sLangId - The possible value for this parameter is a language identifier as defined by
	 * RFC 4646/BCP 47. Example: "en-CA".
     * @returns {bool}
	 */
    ApiPresentation.prototype.SetLanguage = function(sLangId)
    {
        var nLcid = Asc.g_oLcidNameToIdMap[sLangId];
        if (nLcid === undefined)
            return false;

        this.Presentation.SetLanguage(nLcid);
        return true;
    };
    /**
     * Gets slides count.
     * @typeofeditors ["CPE"]
     * @returns {number}
     */
    ApiPresentation.prototype.GetSlidesCount = function()
    {
        return this.Presentation.Slides.length;
    };

    /**
     * Gets slide masters count.
     * @typeofeditors ["CPE"]
     * @returns {number}
     */
    ApiPresentation.prototype.GetMastersCount = function()
    {
        return this.Presentation.slideMasters.length;
    };

    /**
     * Gets a slide master.
     * @typeofeditors ["CPE"]
     * @param {number} nPos 
     * @returns {ApiMaster | null} - returns null if position is invalid.
     */
    ApiPresentation.prototype.GetMaster = function(nPos)
    {
        if (nPos < 0 || nPos >= this.Presentation.slideMasters.length)
            return null;

        return new ApiMaster(this.Presentation.slideMasters[nPos]);
    };

    /**
     * Adds the slide master to presentation slide masters collection.
     * @typeofeditors ["CPE"]
     * @param {number} [nPos    = ApiPresentation.GetMastersCount()]
     * @param {ApiMaster} oApiMaster
     * @returns {bool} - return false if position is invalid or oApiMaster does'n exist.
     */
    ApiPresentation.prototype.AddMaster = function(nPos, oApiMaster)
    {
        if (oApiMaster && oApiMaster.GetClassType && oApiMaster.GetClassType() !== "master")
        {
            if (!nPos || nPos < 0 || nPos > this.Presentation.slideMasters.length)
                nPos = this.Presentation.slideMasters.length;

            this.Presentation.addSlideMaster(nPos, oApiMaster.Master)

            return true;
        }

        return false;
    };

    /**
     * Applies the theme to the all slides in presentation.
     * @typeofeditors ["CPE"]
     * @param {ApiTheme} oApiTheme
     * @returns {bool} - returns false if param isn't theme or presentation doesn't exist.
     * */
    ApiPresentation.prototype.ApplyTheme = function(oApiTheme){
       if (!this.Presentation || !oApiTheme.GetClassType || !oApiTheme.GetClassType() !== "theme")
       {
           this.Presentation.changeTheme(oApiTheme.ThemeInfo);
           return true;
       }

       return false;
    };
    /**
	 * Get a set of default properties for the text run in the current presentation.
	 * @memberof ApiPresentation
	 * @typeofeditors ["CPE"]
	 * @returns {ApiTextPr}
	 */
    ApiPresentation.prototype.GetDefaultTextPr = function()
	{
		var oStyles = this.Presentation.Get_Styles();
		return editor.private_CreateTextPr(this, oStyles.styles.Get_DefaultTextPr().Copy());
	};
    /**
	 * Remove a range of slides from presentation.
     * Deletes all slides without parameters.
	 * @memberof ApiPresentation
     * @param {Number} [nStart=0] - beginning of the deletion range.
     * @param {Number} [nCount=ApiPresentation.GetSlidesCount()] - count of slides for deletion.
	 * @typeofeditors ["CPE"]
	 * @returns {bool}
	 */
    ApiPresentation.prototype.RemoveSlides = function(nStart, nCount)
	{
        nStart = nStart || 0;
        nCount = nCount || this.GetSlidesCount();
        if (AscFormat.isRealNumber(nStart) && nStart > -1 && nStart < this.GetSlidesCount())
        {
            if (AscFormat.isRealNumber(nCount) && nCount > 0)
            {
                nCount = Math.min(nCount, this.GetSlidesCount());
                for (var nSlide = 0; nSlide < nCount; nSlide++)
                    this.Presentation.removeSlide(nStart);

                return true;
            }
        }

        return false;
	};

    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiMaster
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * Get the type of this class.
     * @typeofeditors ["CPE"]
     * @returns {"master"}
     */
    ApiMaster.prototype.GetClassType = function()
    {
        return "master";
    };

    /**
     * Gets a layout of specified slide master by position .
     * @typeofeditors ["CPE"]
     * @param {number} nPos
     * @returns {ApiLayout | null} - returns null if position is invalid.
     */
    ApiMaster.prototype.GetLayout = function(nPos)
    {
        if (nPos < 0 | nPos > this.Master.sldLayoutLst.length)
            return null;
        
        return new ApiLayout(this.Master.sldLayoutLst[nPos])
    };

    /**
     * Adds the layout to specified slide master.
     * @typeofeditors ["CPE"]
     * @param {number} [nPos = ApiMaster.GetLayoutsCount()] - position to add.
     * @param {ApiLayout} oLayout
     * @returns {bool} - returns false if param isn't layout
     */
    ApiMaster.prototype.AddLayout = function(nPos, oLayout)
    {
        if (nPos < 0 || nPos > this.Master.sldLayoutLst.length)
            nPos = this.Master.sldLayoutLst.length;

        if (oLayout && oLayout.GetClassType && oLayout.GetClassType() === "layout")
            this.Master.addToSldLayoutLstToPos(nPos, oLayout.Layout);
        else 
            return false;

        return true;
    };

    /**
     * Removes the layouts from this slide master.
     * @typeofeditors ["CPE"]
     * @param {number} nPos - position from which to delete
     * @param {number} [nCount = 1] - count of elements for delete.
     * @param {ApiLayout} oLayout
     * @returns {bool} - return false if position is invalid.
     */
    ApiMaster.prototype.RemoveLayout = function(nPos, nCount)
    {
        if (this.Master && this.Master.sldLayoutLst.length > 0)
        {
            if (nPos >= 0 && nPos < this.Master.sldLayoutLst.length)
            {
                if (!nCount || nCount <= 0 || nCount > this.GetLayoutsCount())
                    nCount = 1;

                this.Master.removeFromSldLayoutLstByPos(nPos, nCount);
                return true;
            }
        }
        
        return false;
    };

    /**
     * Gets the count of layout objects.
     * @typeofeditors ["CPE"]
     * @returns {number}
     */
    ApiMaster.prototype.GetLayoutsCount = function()
    {
        return this.Master.sldLayoutLst.length;
    };

    /**
     * Adds an object (image, shape or chart) to the current slide master.
     * @typeofeditors ["CPE"]
     * @memberof ApiSlide
     * @param {ApiDrawing} oDrawing - The object which will be added to the current slide master.
     * @returns {bool} - returns false if slide master doesn't exist.
     */
    ApiMaster.prototype.AddObject = function(oDrawing)
    {
        if (this.Master) 
        {
            oDrawing.Drawing.setParent(this.Master);
            this.Master.shapeAdd(this.Master.cSld.spTree.length, oDrawing.Drawing);
            editor.private_checkPlaceholders(this, oDrawing.GetPlaceholder());

            return true;
        }

        return false;
    };

    /**
     * Removes objects (image, shape or chart) from the current slide master.
     * @typeofeditors ["CPE"]
     * @memberof ApiMaster
     * @param {number} nPos - position from which to delete
     * @param {number} [nCount = 1] - count of elements for delete. 
     * @returns {bool} - returns false if master doesn't exist or pos is invalid or master haven't objects.
     */
    ApiMaster.prototype.RemoveObject = function(nPos, nCount)
    {
        if (this.Master && this.Master.cSld.spTree.length > 0)
        {
            if (nPos >= 0 && nPos < this.Master.cSld.spTree.length)
            {
                if (!nCount || nCount <= 0 || nCount > this.Master.cSld.spTree.length)
                    nCount = 1;

                this.Master.shapeRemove(nPos, nCount);
                return true;
            }
        }
        
        return false;
    };
    

    /**
     * Sets the background to the current slide master.
     * @memberOf ApiSlide
     * @typeofeditors ["CPE"]
     * @param {ApiFill} oApiFill - The color or pattern used to fill the presentation slide master background.
     * */
    ApiMaster.prototype.SetBackground = function(oApiFill){
        if(oApiFill && oApiFill.GetClassType && oApiFill.GetClassType() === "fill" && this.Master){
            var bg       = new AscFormat.CBg();
            bg.bgPr      = new AscFormat.CBgPr();
            bg.bgPr.Fill = oApiFill.UniFill;
            this.Master.changeBackground(bg);
        }
    };

    /**
     * Clears the slide master background.
     * @typeofeditors ["CPE"]
     * @returns {bool} - return false if slide master does'n exist.
     * */
    ApiMaster.prototype.ClearBackground = function(){
        if (!this.Master)
            return false;
        
        var apiNoFill = editor.CreateNoFill();
        var bg        = new AscFormat.CBg();
        bg.bgPr       = new AscFormat.CBgPr();
        bg.bgPr.Fill  = apiNoFill.UniFill;
        this.Master.changeBackground(bg);

        return true;
    };

    /**
     * Creates a copy of the specified slide master object.
     * @typeofeditors ["CPE"]
     * @returns {ApiMaster | null} - returns new ApiMaster object that represents the copy of slide master. 
     * Returns null if slide doesn't exist.
     * */
    ApiMaster.prototype.Copy = function(){
        if (!this.Master)
            return null;
        
        var oMasterCopy = this.Master.createDuplicate();
        return new ApiMaster(oMasterCopy);
    };

    /**
     * Creates a duplicate of the specified slide master object, adds the new slide master to the slide masters collection.
     * @typeofeditors ["CPE"]
     * @param {number} [nPos    = ApiPresentation.GetMastersCount()]
     * @returns {ApiMaster | null} - returns new ApiMaster object that represents the copy of slide master. 
     * Returns null if slide master doesn't exist or not in presentation.
     * */
    ApiMaster.prototype.Duplicate = function(nPos){
        if (!this.Master)
            return null;
        
        var oPresentation       = editor.GetPresentation().Presentation;
        var oMasterCopy         = this.Master.createDuplicate();
        
        if (!nPos || nPos < 0 || nPos > oPresentation.Slides.length)
            nPos = oPresentation.slideMasters.length;

        oPresentation.addSlideMaster(nPos, oMasterCopy);

        return new ApiMaster(oMasterCopy);
    };

    /**
     * Deletes the specified object from parent if it exist.
     * @typeofeditors ["CPE"]
     * @returns {bool} - return false if master doesn't exist or not in presentation. 
     * */
    ApiMaster.prototype.Delete = function(){
        if (this.Master && this.Master.presentation)
        {
            for (var nMaster = 0; nMaster < this.Master.presentation.slideMasters.length; nMaster++)
            {
                if (this.Master.Id === this.Master.presentation.slideMasters[nMaster].Id)
                {
                    this.Master.presentation.removeSlideMaster(nMaster, 1);
                    return true;
                }
            }
        }
        
        return false;
    };

    /**
     * Gets the theme of slide master.
     * @typeofeditors ["CPE"]
     * @returns {ApiTheme | null} - returns null if theme doesn't exist. 
     * */
    ApiMaster.prototype.GetTheme = function(){
        if (this.Master && this.Master.Theme)
        {
            var oThemeLoadInfo     = new AscCommonSlide.CThemeLoadInfo();
            oThemeLoadInfo.Master  = this.Master;
            oThemeLoadInfo.Layouts = this.Master.sldLayoutLst;
            oThemeLoadInfo.Theme   = this.Master.Theme;

            return new ApiTheme(oThemeLoadInfo);
        }
           
        return null;
    };

    /**
     * Sets the theme to the slide master.
     * Sets a copy of the theme object.
     * @typeofeditors ["CPE"]
     * @param {ApiTheme} oTheme
     * @returns {bool} - return false if param isn't theme or slide master doesn't exist.
     * */
    ApiMaster.prototype.SetTheme = function(oTheme){
        if (this.Master && oTheme && oTheme.GetClassType && oTheme.GetClassType() === "theme")
        {
            var oThemeCopy = oTheme.ThemeInfo.Theme.createDuplicate();
            this.Master.setTheme(oThemeCopy);
            return true;
        }
           
        return false;
    };
    
    /**
     * Gets the array with all drawing objects in slide master.
     * @typeofeditors ["CPE"]
     * @returns {ApiDrawing[]} 
     * */
    ApiMaster.prototype.GetAllDrawings = function(){
        var apiDrawingObjects = [];
        if (this.Master)
        {
            var drawingObjects = this.Master.cSld.spTree;
            for (var nObject = 0; nObject < drawingObjects.length; nObject++)
                apiDrawingObjects.push(new ApiDrawing(drawingObjects[nObject]));
        }
           
        return apiDrawingObjects;
    };

    /**
     * Gets the array with all shape objects in slide master.
     * @typeofeditors ["CPE"]
     * @returns {ApiShape[]} 
     * */
    ApiMaster.prototype.GetAllShapes = function(){
        var apiShapes = [];
        if (this.Master)
        {
            var drawingObjects = this.Master.cSld.spTree;
            for (var nObject = 0; nObject < drawingObjects.length; nObject++)
            {
                if (drawingObjects[nObject].getObjectType() === AscDFH.historyitem_type_Shape)
                apiShapes.push(new ApiShape(drawingObjects[nObject]));
            }
        }
           
        return apiShapes;
    };

    /**
     * Gets the array with all image objects in slide master.
     * @typeofeditors ["CPE"]
     * @returns {ApiImage[]} 
     * */
    ApiMaster.prototype.GetAllImages = function(){
        var apiImages = [];
        if (this.Master)
        {
            var drawingObjects = this.Master.cSld.spTree;
            for (var nObject = 0; nObject < drawingObjects.length; nObject++)
            {
                if (drawingObjects[nObject].getObjectType() === AscDFH.historyitem_type_ImageShape)
                apiImages.push(new ApiImage(drawingObjects[nObject]));
            }
        }
           
        return apiImages;
    };

    /**
     * Gets the array with all chart objects in slide master.
     * @typeofeditors ["CPE"]
     * @returns {ApiChart[]} 
     * */
    ApiMaster.prototype.GetAllCharts = function(){
        var apiCharts = [];
        if (this.Master)
        {
            var drawingObjects = this.Master.cSld.spTree;
            for (var nObject = 0; nObject < drawingObjects.length; nObject++)
            {
                if (drawingObjects[nObject].getObjectType() === AscDFH.historyitem_type_ChartSpace)
                apiCharts.push(new ApiChart(drawingObjects[nObject]));
            }
        }
           
        return apiCharts;
    };

    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiLayout
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * Get the type of this class.
     * @typeofeditors ["CPE"]
     * @returns {"master"}
     */
    ApiLayout.prototype.GetClassType = function()
    {
        return "layout";
    };

    /**
     * Sets the name of this layout.
     * @typeofeditors ["CPE"]
     * @param {string} sName
     * @returns {bool}
     */
    ApiLayout.prototype.SetName = function(sName)
    {
        if (typeof(sName) !== "string")
            this.Layout.setCSldName(sName);
        else 
            return false;
        
        return true;
    };

    /**
     * Adds an object (image, shape or chart) to the current slide layout.
     * @typeofeditors ["CPE"]
     * @memberof ApiSlide
     * @param {ApiDrawing} oDrawing - The object which will be added to the current slide layout.
     * @returns {bool} - returns false if slide layout doesn't exist.
     */
    ApiLayout.prototype.AddObject = function(oDrawing)
    {
        if (this.Layout) 
        {
            oDrawing.Drawing.setParent(this.Layout);
            this.Layout.shapeAdd(this.Layout.cSld.spTree.length, oDrawing.Drawing);
            editor.private_checkPlaceholders(this, oDrawing.GetPlaceholder());

            return true;
        }

        return false;
    };

    /**
     * Removes objects (image, shape or chart) from the current slide layout.
     * @typeofeditors ["CPE"]
     * @memberof ApiLayout
     * @param {number} nPos - position from which to delete
     * @param {number} [nCount = 1] - count of elements for delete. 
     * @returns {bool} - returns false if layout doesn't exist or pos is invalid or layout haven't objects.
     */
    ApiLayout.prototype.RemoveObject = function(nPos, nCount)
    {
        if (this.Layout && this.Layout.cSld.spTree.length > 0)
        {
            if (nPos >= 0 && nPos < this.Layout.cSld.spTree.length)
            {
                if (!nCount || nCount <= 0 || nCount > this.Layout.cSld.spTree.length)
                    nCount = 1;

                this.Layout.shapeRemove(nPos, nCount);
                return true;
            }
        }
        
        return false;
    };

    /**
     * Sets the background to the current slide layout.
     * @memberOf ApiSlide
     * @typeofeditors ["CPE"]
     * @param {ApiFill} oApiFill - The color or pattern used to fill the presentation slide layout background.
     * */
    ApiLayout.prototype.SetBackground = function(oApiFill){
        if(oApiFill && oApiFill.GetClassType && oApiFill.GetClassType() === "fill" && this.Layout){
            var bg       = new AscFormat.CBg();
            bg.bgPr      = new AscFormat.CBgPr();
            bg.bgPr.Fill = oApiFill.UniFill;
            this.Layout.changeBackground(bg);
        }
    };

    /**
     * Clears the slide layout background.
     * @typeofeditors ["CPE"]
     * @returns {bool} - return false if slide layout does'n exist.
     * */
    ApiLayout.prototype.ClearBackground = function(){
        if (!this.Layout)
            return false;

        var apiNoFill = editor.CreateNoFill();
        var bg        = new AscFormat.CBg();
        bg.bgPr       = new AscFormat.CBgPr();
        bg.bgPr.Fill  = apiNoFill.UniFill;
        this.Layout.changeBackground(bg);

        return true;
    };

    /**
     * Determines whether the slide follows the slide master background.
     * @typeofeditors ["CPE"]
     * @returns {bool} - returns false if master is null or master haven't background.  
     * */
    ApiLayout.prototype.FollowMasterBackground = function(){
        if (!this.Layout)
            return false;
        
        var oMaster = this.Layout.Master;

        if (oMaster && oMaster.cSld.Bg)
        {
            this.Layout.changeBackground(oMaster.cSld.Bg);
            return true;
        }
        else 
            return false;
    };

    /**
     * Creates a copy of the specified slide layout object.
     * @typeofeditors ["CPE"]
     * @returns {ApiLayout | null} - returns new ApiLayout object that represents the copy of slide layout. 
     * Returns null if slide layout doesn't exist.
     * */
    ApiLayout.prototype.Copy = function(){
        if (!this.Layout)
            return null;
        
        var oLayoutCopy = this.Layout.createDuplicate();
        return new ApiLayout(oLayoutCopy);
    };

    /**
     * Deletes the specified object from parent slide master if it exist.
     * @typeofeditors ["CPE"]
     * @returns {bool} - return false if parent slide master does'n exist. 
     * */
    ApiLayout.prototype.Delete = function(){
        if (this.Layout && this.Layout.Master)
        {
            for (var nLayout = 0; nLayout < this.Layout.Master.sldLayoutLst.length; nLayout++)
            {
                if (this.Layout.Id === this.Layout.Master.sldLayoutLst[nLayout].Id)
                {
                    this.Layout.Master.removeFromSldLayoutLstByPos(nLayout, 1);
                    return true;
                }
            }
        }
            return false;
    };

    /**
     * Creates a duplicate of the specified slide layout object, adds the new slide layout to the slide layout collection.
     * @typeofeditors ["CPE"]
     * @param {number} [nPos = ApiMaster.GetLayoutsCount()]
     * @returns {ApiLayout | null} - returns new ApiLayout object that represents the copy of slide layout. 
     * Returns null if slide layout doesn't exist or not in slide master.
     * */
    ApiLayout.prototype.Duplicate = function(nPos){
        if (this.Layout && this.Layout.Master)
        {
            var oMaster     = this.Layout.Master;
            var oLayoutCopy = this.Layout.createDuplicate();
            
            if (nPos < 0 || nPos > this.Layout.Master.sldLayoutLst.length || !nPos)
                nPos = oMaster.sldLayoutLst.length;
    
            oMaster.addToSldLayoutLstToPos(nPos, oLayoutCopy);
    
            return new ApiLayout(oLayoutCopy);
        }
        return null;
    };

    /**
     * Moves the specified object to a specific location within the same collection.
     * @typeofeditors ["CPE"]
     * @returns {bool} - returns false if layout or parent slide master doesn't exist or position is invalid. 
     * */
    ApiLayout.prototype.MoveTo = function(nPos){
        if (!this.Layout || !this.Layout.Master)
            return false;
        if (nPos < 0 || nPos >= this.Layout.Master.sldLayoutLst.length)
            return false;

        for (var Index = 0; Index < this.Layout.Master.sldLayoutLst.length; Index++)
        {
            if (this.Layout.Id === this.Layout.Master.sldLayoutLst[Index].Id)
            {
                this.Layout.Master.moveLayouts([Index], nPos)
                return true;
            }
        }
    };

    /**
     * Gets the array with all drawing objects in slide layout.
     * @typeofeditors ["CPE"]
     * @returns {ApiDrawing[]} 
     * */
    ApiLayout.prototype.GetAllDrawings = function(){
        var apiDrawingObjects = [];
        if (this.Layout)
        {
            var drawingObjects = this.Layout.cSld.spTree;
            for (var nObject = 0; nObject < drawingObjects.length; nObject++)
                apiDrawingObjects.push(new ApiDrawing(drawingObjects[nObject]));
        }
           
        return apiDrawingObjects;
    };

    /**
     * Gets the array with all shape objects in slide layout.
     * @typeofeditors ["CPE"]
     * @returns {ApiShape[]} 
     * */
    ApiLayout.prototype.GetAllShapes = function(){
        var apiShapes = [];
        if (this.Layout)
        {
            var drawingObjects = this.Layout.cSld.spTree;
            for (var nObject = 0; nObject < drawingObjects.length; nObject++)
            {
                if (drawingObjects[nObject].getObjectType() === AscDFH.historyitem_type_Shape)
                apiShapes.push(new ApiShape(drawingObjects[nObject]));
            }
        }
           
        return apiShapes;
    };

    /**
     * Gets the array with all image objects in slide layout.
     * @typeofeditors ["CPE"]
     * @returns {ApiImage[]} 
     * */
    ApiLayout.prototype.GetAllImages = function(){
        var apiImages = [];
        if (this.Layout)
        {
            var drawingObjects = this.Layout.cSld.spTree;
            for (var nObject = 0; nObject < drawingObjects.length; nObject++)
            {
                if (drawingObjects[nObject].getObjectType() === AscDFH.historyitem_type_ImageShape)
                apiImages.push(new ApiImage(drawingObjects[nObject]));
            }
        }
           
        return apiImages;
    };

    /**
     * Gets the array with all chart objects in slide layout.
     * @typeofeditors ["CPE"]
     * @returns {ApiChart[]} 
     * */
    ApiLayout.prototype.GetAllCharts = function(){
        var apiCharts = [];
        if (this.Layout)
        {
            var drawingObjects = this.Layout.cSld.spTree;
            for (var nObject = 0; nObject < drawingObjects.length; nObject++)
            {
                if (drawingObjects[nObject].getObjectType() === AscDFH.historyitem_type_ChartSpace)
                apiCharts.push(new ApiChart(drawingObjects[nObject]));
            }
        }
           
        return apiCharts;
    };

    /**
     * Gets parent slide master.
     * @typeofeditors ["CPE"]
     * @returns {?ApiMaster} - returns null if parent slide master doesn't exist.
     * */
    ApiLayout.prototype.GetMaster = function(){
        if (this.Layout && this.Layout.Master)
            return new ApiMaster(this.Layout.Master);
           
        return null;
    };

    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiPlaceholder
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * Get the type of this class.
     * @typeofeditors ["CPE"]
     * @returns {"placeholder"}
     */
    ApiPlaceholder.prototype.GetClassType = function()
    {
        return "placeholder";
    };
    /**
     * Set the placeholder type.
     * @typeofeditors ["CPE"]
     * @param {string} sType - name of placeholder type.
     * @returns {bool} - returns false if placeholder type doesn't exist.
     */
    ApiPlaceholder.prototype.SetType = function(sType)
    {
        var nType   = null;
        switch (sType)
        {
            case "body":
                nType = 0;
                break;
            case "chart":
                nType = 1;
                break;
            case "clipArt":
                nType = 2;
                break;
            case "ctrTitle":
                nType = 3;
                break;
            case "diagram":
                nType = 4;
                break;
            case "date":
                nType = 5;
                break;
            case "footer":
                nType = 6;
                break;
            case "header":
                nType = 7;
                break;
            case "media":
                nType = 8;
                break;
            case "object":
                nType = 9;
                break;
            case "picture":
                nType = 10;
                break;
            case "sldImage":
                nType = 11;
                break;
            case "sldNumber":
                nType = 12;
                break;
            case "subTitle":
                nType = 13;
                break;
            case "table":
                nType = 14;
                break;
            case "title":
                nType = 15;
                break;
            default:
                nType = 0;
        }

        this.Placeholder.setType(nType);
    };

    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiTheme
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * Get the type of this class.
     * @typeofeditors ["CPE"]
     * @returns {"theme"}
     */
    ApiTheme.prototype.GetClassType = function()
    {
        return "theme";
    };

    /**
     * Get the slide master of this theme.
     * @typeofeditors ["CPE"]
     * @returns {ApiMaster | null} - returns null if slide master doesn't exist.
     */
    ApiTheme.prototype.GetMaster = function()
    {
        if (this.ThemeInfo && this.ThemeInfo.Master)
            return new ApiMaster(this.ThemeInfo.Master);

        return null;
    };

    /**
     * Sets the presentation color scheme.
     * @typeofeditors ["CPE"]
     * @param {ApiThemeColorScheme} oApiColorScheme
     * @returns {bool} - return false if color scheme doesn't exist.
     */
    ApiTheme.prototype.SetColorScheme = function(oApiColorScheme)
    {
        if (oApiColorScheme && oApiColorScheme.GetClassType && oApiColorScheme.GetClassType() === "themeColorScheme")
        {
            this.ThemeInfo.Theme.themeElements.clrScheme = oApiColorScheme.ColorScheme;
            return true;
        }

        return false;
    };

    /**
     * Gets the presentation color scheme.
     * @typeofeditors ["CPE"]
     * @returns {?ApiThemeColorScheme}
     */
    ApiTheme.prototype.GetColorScheme = function()
    {
        if (this.ThemeInfo && this.ThemeInfo.Theme && this.ThemeInfo.Theme.themeElements)
        {
            return new ApiThemeColorScheme(this.ThemeInfo.Theme.themeElements.clrScheme);
        }

        return null;
    };

    /**
     * Sets the presentation format scheme.
     * @typeofeditors ["CPE"]
     * @param {ApiThemeFormatScheme} oApiFormatScheme
     * @returns {bool} - return false if format scheme doesn't exist.
     */
    ApiTheme.prototype.SetFormatScheme = function(oApiFormatScheme)
    {
        if (oApiFormatScheme && oApiFormatScheme.GetClassType && oApiFormatScheme.GetClassType() === "themeColorScheme")
        {
            this.ThemeInfo.Theme.themeElements.fmtScheme = oApiFormatScheme.FormatScheme;
            return true;
        }

        return false;
    };

    /**
     * Gets the presentation format scheme.
     * @typeofeditors ["CPE"]
     * @returns {?ApiThemeFormatScheme}
     */
    ApiTheme.prototype.GetFormatScheme = function()
    {
        if (this.ThemeInfo && this.ThemeInfo.Theme && this.ThemeInfo.Theme.themeElements)
        {
            return new ApiThemeFormatScheme(this.ThemeInfo.Theme.themeElements.fmtScheme);
        }

        return null;
    };

    /**
     * Sets the presentation font scheme.
     * @typeofeditors ["CPE"]
     * @param {ApiThemeFontScheme} oApiFontScheme
     * @returns {bool} - return false if font scheme doesn't exist.
     */
    ApiTheme.prototype.SetFontScheme = function(oApiFontScheme)
    {
        if (oApiFontScheme && oApiFontScheme.GetClassType && oApiFontScheme.GetClassType() === "themeColorScheme")
        {
            this.ThemeInfo.Theme.themeElements.fontScheme = oApiFontScheme.FontScheme;
            return true;
        }

        return false;
    };

    /**
     * Gets the presentation font scheme.
     * @typeofeditors ["CPE"]
     * @returns {?ApiThemeFontScheme}
     */
    ApiTheme.prototype.GetFontScheme = function()
    {
        if (this.ThemeInfo && this.ThemeInfo.Theme && this.ThemeInfo.Theme.themeElements)
        {
            return new ApiThemeFontScheme(this.ThemeInfo.Theme.themeElements.fontScheme);
        }

        return null;
    };

    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiThemeColorScheme
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * Gets the type of this class.
     * @typeofeditors ["CPE"]
     * @returns {"themeColorScheme"}
     */
    ApiThemeColorScheme.prototype.GetClassType = function()
    {
        return "themeColorScheme";
    };

    /**
     * Sets the color scheme name.
     * @typeofeditors ["CPE"]
     * @param {string} sName
     * @returns {bool}
     */
    ApiThemeColorScheme.prototype.SetSchemeName = function(sName)
    {
        if (typeof(sName) !== "string")
            sName = "";

        this.ColorScheme.setName(sName);
    };

    /**
     * Change color in color scheme.
     * @typeofeditors ["CPE"]
     * @param {ApiUniColor | ApiRGBColor} oColor
     * @returns {bool}
     */
    ApiThemeColorScheme.prototype.ChangeColor = function(nPos, oColor)
    {
        if (nPos < 0 || nPos > 12 || (oColor.GetClassType() !== "rgbColor" && oColor.GetClassType() !== "uniColor"))
            return false;

        if (nPos <= 5)
            this.ColorScheme.addColor(nPos, oColor.Unicolor);
        else if (nPos > 5)
            this.ColorScheme.addColor(nPos + 2, oColor.Unicolor)

        return true;
    };

    /**
     * Creates a copy of the specified theme's color scheme.
     * @typeofeditors ["CPE"]
     * @returns {ApiThemeColorScheme}
     */
    ApiThemeColorScheme.prototype.Copy = function()
    {
        return new ApiThemeColorScheme(this.ColorScheme.createDuplicate());
    };

    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiThemeFormatScheme
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * Get the type of this class.
     * @typeofeditors ["CPE"]
     * @returns {"formatColorScheme"}
     */
    ApiThemeFormatScheme.prototype.GetClassType = function()
    {
        return "themeFormatScheme";
    };

    /**
     * Sets the format scheme name.
     * @typeofeditors ["CPE"]
     * @param {string} sName
     * @returns {bool}
     */
    ApiThemeFormatScheme.prototype.SetSchemeName = function(sName)
    {
        if (typeof(sName) !== "string")
            sName = "";

        this.FormatScheme.setName(sName);
    };

    /**
     * Sets the fill styles.
     * @typeofeditors ["CPE"]
     * @param {ApiFill[]} arrFill - the array of styles must contains 3 elements and consist of subtle, moderate and intense fills.
     * If there are empty  or no ApiFill elements in the array, it will be filled by Api.CreateNoFill() element.
     */
    ApiThemeFormatScheme.prototype.ChangeFillStyles = function(arrFill)
    {
        if (!arrFill)
            arrFill = [];

        this.FormatScheme.fillStyleLst = [];

        for (var nFill = 0; nFill < 3; nFill++)
        {
            if (arrFill[nFill] && arrFill[nFill].GetClassType() === "fill")
                this.FormatScheme.addFillToStyleLst(arrFill[nFill].UniFill);
            else 
                this.FormatScheme.addFillToStyleLst(editor.CreateNoFill().UniFill);
        }
    };

    /**
     * Sets the background fill styles.
     * @typeofeditors ["CPE"]
     * @param {ApiFill[]} arrBgFill - the array of background fill styles must contains 3 elements and consist of subtle, moderate and intense fills.
     * If there are empty or no ApiFill elements in the array, it will be filled by Api.CreateNoFill() element.
     */
    ApiThemeFormatScheme.prototype.ChangeBgFillStyles = function(arrBgFill)
    {
        if (!arrBgFill)
            arrBgFill = [];

        this.FormatScheme.bgFillStyleLst = [];

        for (var nFill = 0; nFill < 3; nFill++)
        {
            if (arrBgFill[nFill] && arrBgFill[nFill].GetClassType() === "fill")
                this.FormatScheme.addBgFillToStyleLst(arrBgFill[nFill].UniFill);
            else 
                this.FormatScheme.addBgFillToStyleLst(editor.CreateNoFill().UniFill);
        }
    };

    /**
     * Sets the line styles.
     * @typeofeditors ["CPE"]
     * @param {ApiStroke[]} arrLine - the array of line styles must contains 3 elements and consist of subtle, moderate and intense fills.
     * If there are empty or no ApiStroke elements in the array, it will be filled by Api.CreateStroke(0, Api.CreateNoFill()) element.
     */
    ApiThemeFormatScheme.prototype.ChangeLineStyles = function(arrLine)
    {
        if (!arrLine)
            arrLine = [];

        this.FormatScheme.lnStyleLst = [];

        for (var nLine = 0; nLine < 3; nLine++)
        {
            if (arrLine[nLine] && arrLine[nLine].GetClassType() === "stroke")
                this.FormatScheme.addLnToStyleLst(arrLine[nLine].Ln);
            else 
                this.FormatScheme.addLnToStyleLst(editor.CreateStroke(0, editor.CreateNoFill()).Ln);
        }
    };

    /**
     * **Need to do**
     * Sets the effect styles.
     * @typeofeditors ["CPE"]
     * @param {?[]} arrEffect - the array of line styles must contains 3 elements and consist of subtle, moderate and intense fills.
     * If there are empty or no ApiStroke elements in the array, it will be filled by Api.CreateStroke(0, Api.CreateNoFill()) element.
     * @returns {bool}
     */
    ApiThemeFormatScheme.prototype.ChangeEffectStyles = function(arrEffect)
    {
        // if (!arrEffect)
            // arrEffect = [];

        // this.FormatScheme.effectStyleLst = [];

        // for (var nFill = 0; nFill < 3; nFill++)
        // {
        //     if (arrEffect[nFill] && arrEffect[nFill].GetClassType() === "stroke")
        //         this.FormatScheme.addEffectToStyleLst(arrEffect[nFill].UniFill);
        //     else 
        //         this.FormatScheme.addEffectToStyleLst(editor.CreateNoFill().UniFill);
        // }

        // return true;
    };

    /**
     * Creates a copy of the specified theme's format scheme.
     * @typeofeditors ["CPE"]
     * @returns {ApiThemeFormatScheme}
     */
    ApiThemeFormatScheme.prototype.Copy = function()
    {
        return new ApiThemeFormatScheme(this.FormatScheme.createDuplicate());
    };

    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiThemeFontScheme
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * Get the type of this class.
     * @typeofeditors ["CPE"]
     * @returns {"fontScheme"}
     */
    ApiThemeFontScheme.prototype.GetClassType = function()
    {
        return "themeFontScheme";
    };
    
    /**
     * Sets the scheme's name.
     * @typeofeditors ["CPE"]
     * @param {string} sName
     * @returns {bool} - returns false if font scheme doesn't exist.
     */
    ApiThemeFontScheme.prototype.SetSchemeName = function(sName)
    {
        if (typeof(sName) !== "string")
            sName = "";

        if (this.FontScheme)
        {
            this.FontScheme.setName(sName);
            return true;
        }
        else 
            return false;
    };

    /**
     * Sets the fonts for font scheme.
     * @typeofeditors ["CPE"]
     * @memberof Api
     * @param {string} mjLatin - Font's name. Is used to determine the major theme font applied to latit text.
     * @param {string} mjEa - Font's name. Is used to determine the major theme font applied to east asian text.
     * @param {string} mjCs - Font's name. Is used to determine the major theme font applied to complex script text.
     * @param {string} mnLatin - Font's name. Is used to determine the minor theme font applied to latit text.
     * @param {string} mnEa - Font's name. Is used to determine the minor theme font applied to east asian text.
     * @param {string} mnCs - Font's name. Is used to determine the minor theme font applied to complex script text.
     */
    ApiThemeFontScheme.prototype.SetFonts = function(mjLatin, mjEa, mjCs, mnLatin, mnEa, mnCs){
        
        var oMajorFontCollection = this.FontScheme.majorFont;
        var oMinorFontCollection = this.FontScheme.minorFont;

        if (typeof(mjLatin) === "string")
            oMajorFontCollection.setLatin(mjLatin);
        if (typeof(mjEa) === "string")
            oMajorFontCollection.setEA(mjEa);
        if (typeof(mjCs) === "string")
            oMajorFontCollection.setCS(mjCs);

        if (typeof(mnLatin) === "string")
            oMinorFontCollection.setLatin(mnLatin);
        if (typeof(mnEa) === "string")
            oMinorFontCollection.setEA(mnEa);
        if (typeof(mnCs) === "string")
            oMinorFontCollection.setCS(mnCs);
    };

    /**
     * Creates a copy of the specified theme's font scheme.
     * @typeofeditors ["CPE"]
     * @returns {ApiThemeFontScheme}
     */
    ApiThemeFontScheme.prototype.Copy = function()
    {
        return new ApiThemeFontScheme(this.FontScheme.createDuplicate());
    };
    
    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiSlide
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * Get the type of this class.
     * @typeofeditors ["CPE"]
     * @returns {"slide"}
     */
    ApiSlide.prototype.GetClassType = function()
    {
        return "slide";
    };

    /**
     * Remove all the objects from the current slide.
     * @typeofeditors ["CPE"]
     * @memberof ApiSlide
     */
    ApiSlide.prototype.RemoveAllObjects =  function(){
        if(this.Slide){
            var spTree = this.Slide.cSld.spTree;
            for(var i = spTree.length - 1; i > -1; --i){
                this.Slide.removeFromSpTreeById(spTree[i].Get_Id());
            }
        }
    };

    /**
     * Add an object (image, shape or chart) to the current presentation slide.
     * @typeofeditors ["CPE"]
     * @memberof ApiSlide
     * @param {ApiDrawing} oDrawing - The object which will be added to the current presentation slide.
     * @returns {bool} - returns false if slide doesn't exist.
     */
    ApiSlide.prototype.AddObject = function(oDrawing){
        if(this.Slide){
            oDrawing.Drawing.setParent(this.Slide);
            this.Slide.shapeAdd(this.Slide.cSld.spTree.length, oDrawing.Drawing);
            editor.private_checkPlaceholders(this, oDrawing.GetPlaceholder());

            return true;
        }

        return false;
    };

    /**
     * Removes objects (image, shape or chart) from the current slide.
     * @typeofeditors ["CPE"]
     * @memberof ApiSlide
     * @param {number} nPos - position from which to delete
     * @param {number} [nCount = 1] - count of elements for delete. 
     * @returns {bool} - returns false if slide doesn't exist or pos is invalid or slide haven't objects.
     */
    ApiSlide.prototype.RemoveObject = function(nPos, nCount)
    {
        if (this.Slide && this.Slide.cSld.spTree.length > 0)
        {
            if (nPos >= 0 && nPos < this.Slide.cSld.spTree.length)
            {
                if (!nCount || nCount <= 0 || nCount > this.Slide.cSld.spTree.length)
                    nCount = 1;

                this.Slide.shapeRemove(nPos, nCount);
                return true;
            }
        }
        
        return false;
    };

    /**
     * Set the background to the current presentation slide.
     * @memberOf ApiSlide
     * @typeofeditors ["CPE"]
     * @param {ApiFill} oApiFill - The color or pattern used to fill the presentation slide background.
     * */
    ApiSlide.prototype.SetBackground = function(oApiFill){
        if(oApiFill && oApiFill.GetClassType && oApiFill.GetClassType() === "fill" && this.Slide){
            var bg       = new AscFormat.CBg();
            bg.bgPr      = new AscFormat.CBgPr();
            bg.bgPr.Fill = oApiFill.UniFill;
            this.Slide.changeBackground(bg);
            this.Slide.recalculateBackground();
        }
    };

    /**
     * Get the slide width in English measure units.
     * @typeofeditors ["CPE"]
     * @returns {EMU}
     * */
    ApiSlide.prototype.GetWidth = function(){
        if(this.Slide){
            return this.Slide.Width*36000;
        }
        return 0;
    };

    /**
     * Get the slide height in English measure units.
     * @typeofeditors ["CPE"]
     * @returns {EMU}
     * */
    ApiSlide.prototype.GetHeight = function(){
        if(this.Slide){
            return this.Slide.Height*36000;
        }
        return 0;
    };

    /**
     * Applies specified layout to the slide.
     * @typeofeditors ["CPE"]
     * @param {ApiLayout} oLayout
     * @returns {bool} - returns false if slide doesn't exist.
     * */
    ApiSlide.prototype.ApplyLayout = function(oLayout){
        if (!this.Slide)
            return false;

        this.RemoveAllObjects();
        this.Slide.setLayout(oLayout.Layout);

        var oPresentation = editor.GetPresentation().Presentation;
        var layout, i, _ph_type, sp, hf, bIsSpecialPh;

        layout = oLayout.Layout;
        hf = oLayout.Layout.Master.hf;
        this.Slide.setNotes(AscCommonSlide.CreateNotes());
        this.Slide.notes.setNotesMaster(oPresentation.notesMasters[0]);
        this.Slide.notes.setSlide(this.Slide);
        for (i = 0; i < layout.cSld.spTree.length; ++i) {
            if (layout.cSld.spTree[i].isPlaceholder()) {
                _ph_type = layout.cSld.spTree[i].getPhType();
                bIsSpecialPh = _ph_type === AscFormat.phType_dt || _ph_type === AscFormat.phType_ftr || _ph_type === AscFormat.phType_hdr || _ph_type === AscFormat.phType_sldNum;
                if (!bIsSpecialPh || hf && ((_ph_type === AscFormat.phType_dt && (hf.dt !== false)) ||
                    (_ph_type === AscFormat.phType_ftr && (hf.ftr !== false)) ||
                    (_ph_type === AscFormat.phType_hdr && (hf.hdr !== false)) ||
                    (_ph_type === AscFormat.phType_sldNum && (hf.sldNum !== false)))) {
                    sp = layout.cSld.spTree[i].copy(undefined);
                    sp.setParent(this.Slide);
                    !bIsSpecialPh && sp.clearContent && sp.clearContent();
                    this.Slide.addToSpTreeToPos(this.Slide.cSld.spTree.length, sp);
                }
            }
        }
        this.Slide.setSlideSize(oPresentation.Width, oPresentation.Height);

        return true;
    };

    /**
     * Deletes the specified slide from presentation.
     * @typeofeditors ["CPE"]
     * @returns {bool} - returns false if slide doesn't exist or not in presentation.
     * */
    ApiSlide.prototype.Delete = function(){
        if (!this.Slide)
            return false;
        
        var oPresentation = editor.GetPresentation().Presentation;
        var nPosToDelete  = this.GetSlideIndex();

        if (nPosToDelete > -1)
        {
            oPresentation.removeSlide(nPosToDelete);
            return true;
        }

        return false;
    };

    /**
     * Creates a copy of the specified slide object.
     * @typeofeditors ["CPE"]
     * @returns {ApiSlide | null} - returns new ApiSlide object that represents the duplicate slide. 
     * Returns null if slide doesn't exist.
     * */
    ApiSlide.prototype.Copy = function(){
        if (!this.Slide)
            return null;
        
        var oSlideCopy = this.Slide.createDuplicate();
        return new ApiSlide(oSlideCopy);
    };

    /**
     * Creates a duplicate of the specified slide object, adds the new slide to the slides collection.
     * @typeofeditors ["CPE"]
     * @param {number} [nPos    = ApiPresentation.GetSlidesCount()]
     * @returns {ApiSlide | null} - returns new ApiSlide object that represents the duplicate slide. 
     * Returns null if slide doesn't exist or not in presentation.
     * */
    ApiSlide.prototype.Duplicate = function(nPos){
        if (!this.Slide)
            return null;
        
        var oPresentation = editor.GetPresentation().Presentation;
        var oSlideCopy    = this.Slide.createDuplicate();
        
        if (nPos < 0 || nPos > oPresentation.Slides.length || !nPos)
            nPos = oPresentation.Slides.length;

        oPresentation.insertSlide(nPos, oSlideCopy);

        return new ApiSlide(oSlideCopy);
    };

    /**
     * Moves the specified slide to a specific location within the same collection.
     * @typeofeditors ["CPE"]
     * @returns {bool} - returns false if slide doesn't exist or position is invalid or slide not in presentation.  
     * */
    ApiSlide.prototype.MoveTo = function(nPos){
        var oPresentation = editor.GetPresentation().Presentation;

        if (!this.Slide || nPos < 0 || nPos > oPresentation.Slides.length)
            return false;

        for (var Index = 0; Index < oPresentation.Slides.length; Index++)
        {
            if (this.Slide.Id === oPresentation.Slides[Index].Id)
            {
                oPresentation.moveSlides([Index], nPos)
                return true;
            }
        }

        return false;
    };

    /**
     * Gets the positions of the spicified slide in a presentation.
     * @typeofeditors ["CPE"]
     * @returns {number} - returns -1 if slide doesn't exist or not in presentation.  
     * */
    ApiSlide.prototype.GetSlideIndex = function (){
        if (!this.Slide)
            return -1;
        
        var oPresentation = editor.GetPresentation().Presentation;

        for (var Index = 0; Index < oPresentation.Slides.length; Index++)
        {
            if (this.Slide.Id === oPresentation.Slides[Index].Id)
            {
                return Index;
            }
        }

        return -1;
    };

    /**
     * Clears the slide background.
     * @typeofeditors ["CPE"]
     * @returns {bool} - return false if slide does'n exist.
     * */
    ApiSlide.prototype.ClearBackground = function(){
        if (!this.Slide)
            return false;
        
        var apiNoFill = editor.CreateNoFill();
        var bg        = new AscFormat.CBg();
        bg.bgPr       = new AscFormat.CBgPr();
        bg.bgPr.Fill  = apiNoFill.UniFill;
        this.Slide.changeBackground(bg);
        this.Slide.recalculateBackground();

        return true;
    };

    /**
     * Sets the slide background as the background of slide layout. 
     * @typeofeditors ["CPE"]
     * @returns {bool} - returns false if layout is null or layout haven't background or slide does'n exist. 
     * */
    ApiSlide.prototype.FollowLayoutBackground = function(){
        if (!this.Slide)
            return false;
        
        var Layout = this.Slide.Layout;

        if (Layout && Layout.cSld.Bg)
        {
            this.Slide.changeBackground(Layout.cSld.Bg);
            this.Slide.recalculateBackground();
            return true;
        }
        else 
            return false;
    };

    /**
     * Sets the slide background as the background of slide master. 
     * @typeofeditors ["CPE"]
     * @returns {bool} - returns false if master is null or master haven't background or slide does'n exist.  
     * */
    ApiSlide.prototype.FollowMasterBackground = function(){
        if (!this.Slide)
            return false;
        
        var oMaster = this.Slide.Layout.Master;

        if (oMaster && oMaster.cSld.Bg)
        {
            this.Slide.changeBackground(oMaster.cSld.Bg);
            return true;
        }
        else 
            return false;
    };

    /**
     * Applies a specified theme to the slide.
     * @typeofeditors ["CPE"]
     * @param {ApiTheme} oApiTheme
     * @returns {bool} - returns false if master is null or master haven't background.  
     * */
    ApiSlide.prototype.ApplyTheme = function(oApiTheme){
        if (!this.Slide || !oApiTheme || !oApiTheme.GetClassType ||oApiTheme.GetClassType() !== "theme")
            return false;

        var oPresentation = editor.GetPresentation().Presentation;
        var i;
    
        oPresentation.clearThemeTimeouts();
        for (i = 0; i < oPresentation.slideMasters.length; ++i) {
            if (oPresentation.slideMasters[i] === oApiTheme.ThemeInfo.Master) {
                break;
            }
        }
        if (i === oPresentation.slideMasters.length) {
            oPresentation.addSlideMaster(oPresentation.slideMasters.length, oApiTheme.ThemeInfo.Master);
        }
        var oldMaster = this.Slide && this.Slide.Layout && this.Slide.Layout.Master;
        var _new_master = oApiTheme.ThemeInfo.Master;
        _new_master.presentation = oPresentation;
        oApiTheme.ThemeInfo.Master.changeSize(oPresentation.Width, oPresentation.Height);
        var oContent, oMasterSp, oMasterContent, oSp;
        if (oldMaster && oldMaster.hf) {
            oApiTheme.ThemeInfo.Master.setHF(oldMaster.hf.createDuplicate());
            if (oldMaster.hf.dt !== false) {
                oMasterSp = oldMaster.getMatchingShape(AscFormat.phType_dt, null, false, {});
                if (oMasterSp) {
                    oMasterContent = oMasterSp.getDocContent && oMasterSp.getDocContent();
                    if (oMasterContent) {
                        oSp = oApiTheme.ThemeInfo.Master.getMatchingShape(AscFormat.phType_dt, null, false, {});
                        if (oSp) {
                            oContent = oSp.getDocContent && oSp.getDocContent();
                            oContent.Copy2(oMasterContent);
                        }
                        for (i = 0; i < oApiTheme.ThemeInfo.Master.sldLayoutLst.length; ++i) {
                            oSp = oApiTheme.ThemeInfo.Master.sldLayoutLst[i].getMatchingShape(AscFormat.phType_dt, null, false, {});
                            if (oSp) {
                                oContent = oSp.getDocContent && oSp.getDocContent();
                                oContent.Copy2(oMasterContent);
                            }
                        }
                    }
                }
            }
            if (oldMaster.hf.hdr !== false) {
                oMasterSp = oldMaster.getMatchingShape(AscFormat.phType_hdr, null, false, {});
                if (oMasterSp) {
                    oMasterContent = oMasterSp.getDocContent && oMasterSp.getDocContent();
                    if (oMasterContent) {
                        oSp = oApiTheme.ThemeInfo.Master.getMatchingShape(AscFormat.phType_hdr, null, false, {});
                        if (oSp) {
                            oContent = oSp.getDocContent && oSp.getDocContent();
                            oContent.Copy2(oMasterContent);
                        }
                        for (i = 0; i < oApiTheme.ThemeInfo.Master.sldLayoutLst.length; ++i) {
                            oSp = oApiTheme.ThemeInfo.Master.sldLayoutLst[i].getMatchingShape(AscFormat.phType_hdr, null, false, {});
                            if (oSp) {
                                oContent = oSp.getDocContent && oSp.getDocContent();
                                oContent.Copy2(oMasterContent);
                            }
                        }
                    }
                }
            }
            if (oldMaster.hf.ftr !== false) {
                oMasterSp = oldMaster.getMatchingShape(AscFormat.phType_ftr, null, false, {});
                if (oMasterSp) {
                    oMasterContent = oMasterSp.getDocContent && oMasterSp.getDocContent();
                    if (oMasterContent) {
                        oSp = oApiTheme.ThemeInfo.Master.getMatchingShape(AscFormat.phType_ftr, null, false, {});
                        if (oSp) {
                            oContent = oSp.getDocContent && oSp.getDocContent();
                            oContent.Copy2(oMasterContent);
                        }
                        for (i = 0; i < oApiTheme.ThemeInfo.Master.sldLayoutLst.length; ++i) {
                            oSp = oApiTheme.ThemeInfo.Master.sldLayoutLst[i].getMatchingShape(AscFormat.phType_ftr, null, false, {});
                            if (oSp) {
                                oContent = oSp.getDocContent && oSp.getDocContent();
                                oContent.Copy2(oMasterContent);
                            }
                        }
                    }
                }
            }
        }
        for (i = 0; i < oApiTheme.ThemeInfo.Master.sldLayoutLst.length; ++i) {
            oApiTheme.ThemeInfo.Master.sldLayoutLst[i].changeSize(oPresentation.Width, oPresentation.Height);
        }
        
        var new_layout;
      
        if (this.Slide.Layout.calculatedType == null) {
            this.Slide.Layout.calculateType();
        }
        new_layout = _new_master.getMatchingLayout(this.Slide.Layout.type, this.Slide.Layout.matchingName, this.Slide.Layout.cSld.name, true);
        if (!isRealObject(new_layout)) {
            new_layout = _new_master.sldLayoutLst[0];
        }
        this.Slide.setLayout(new_layout);
        this.Slide.checkNoTransformPlaceholder();

        return true;
    };

    /**
     * Gets layout of the slide.
     * @typeofeditors ["CPE"]
     * @returns {ApiLayout | null} - returns null if slide or layout doesn't exist. 
     * */
    ApiSlide.prototype.GetLayout = function(){
        if (this.Slide && this.Slide.Layout)
            return new ApiLayout(this.Slide.Layout);
           
        return null;
    };

    /**
     * Gets the theme of slide.
     * @typeofeditors ["CPE"]
     * @returns {ApiTheme} - returns null if slide or layout or master or theme doesn't exist.
     * */
    ApiSlide.prototype.GetTheme = function(){
        if (this.Slide && this.Slide.Layout && this.Slide.Layout.Master && this.Slide.Layout.Master.Theme)
        {
            var oThemeLoadInfo     = new AscCommonSlide.CThemeLoadInfo();
            oThemeLoadInfo.Master  = this.Slide.Layout.Master;
            oThemeLoadInfo.Layouts = this.Slide.Layout.Master.sldLayoutLst;
            oThemeLoadInfo.Theme   = this.Slide.Layout.Master.Theme;

            return new ApiTheme(oThemeLoadInfo);
        }
           
        return null;
    };

    /**
     * Gets the array with all drawing objects in slide.
     * @typeofeditors ["CPE"]
     * @returns {ApiDrawing[]} 
     * */
    ApiSlide.prototype.GetAllDrawings = function(){
        var apiDrawingObjects = [];
        if (this.Slide)
        {
            var drawingObjects = this.Slide.getDrawingObjects();
            for (var nObject = 0; nObject < drawingObjects.length; nObject++)
                apiDrawingObjects.push(new ApiDrawing(drawingObjects[nObject]));
        }
           
        return apiDrawingObjects;
    };

    /**
     * Gets the array with all shape objects in slide.
     * @typeofeditors ["CPE"]
     * @returns {ApiShape[]} 
     * */
    ApiSlide.prototype.GetAllShapes = function(){
        var apiShapes = [];
        if (this.Slide)
        {
            var drawingObjects = this.Slide.getDrawingObjects();
            for (var nObject = 0; nObject < drawingObjects.length; nObject++)
            {
                if (drawingObjects[nObject].getObjectType() === AscDFH.historyitem_type_Shape)
                apiShapes.push(new ApiShape(drawingObjects[nObject]));
            }
        }
           
        return apiShapes;
    };

    /**
     * Gets the array with all image objects in slide.
     * @typeofeditors ["CPE"]
     * @returns {ApiImage[]} 
     * */
    ApiSlide.prototype.GetAllImages = function(){
        var apiImages = [];
        if (this.Slide)
        {
            var drawingObjects = this.Slide.getDrawingObjects();
            for (var nObject = 0; nObject < drawingObjects.length; nObject++)
            {
                if (drawingObjects[nObject].getObjectType() === AscDFH.historyitem_type_ImageShape)
                apiImages.push(new ApiImage(drawingObjects[nObject]));
            }
        }
           
        return apiImages;
    };

    /**
     * Gets the array with all chart objects in slide.
     * @typeofeditors ["CPE"]
     * @returns {ApiChart[]} 
     * */
    ApiSlide.prototype.GetAllCharts = function(){
        var apiCharts = [];
        if (this.Slide)
        {
            var drawingObjects = this.Slide.getDrawingObjects();
            for (var nObject = 0; nObject < drawingObjects.length; nObject++)
            {
                if (drawingObjects[nObject].getObjectType() === AscDFH.historyitem_type_ChartSpace)
                apiCharts.push(new ApiChart(drawingObjects[nObject]));
            }
        }
           
        return apiCharts;
    };

    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiDrawing
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * Get the type of the class based on this base class.
     * @returns {"drawing"}
     */
    ApiDrawing.prototype.GetClassType = function()
    {
        return "drawing";
    };
    /**
     * Set the size of the object (image, shape, chart) bounding box.
     * @param {EMU} nWidth - The object width measured in English measure units.
     * @param {EMU} nHeight - The object height measured in English measure units.
     */
    ApiDrawing.prototype.SetSize = function(nWidth, nHeight)
    {
        var fWidth = private_EMU2MM(nWidth);
        var fHeight = private_EMU2MM(nHeight);
        if(this.Drawing && this.Drawing.spPr && this.Drawing.spPr.xfrm)
        {
            this.Drawing.spPr.xfrm.setExtX(fWidth);
            this.Drawing.spPr.xfrm.setExtY(fHeight);
        }
    };

    /**
     * Set the position of the drawing on the slide.
     * @param {EMU} nPosX - The distance from the left side of the slide to the left side of the drawing measured in English measure units.
     * @param {EMU} nPosY - The distance from the top side of the slide to the upper side of the drawing measured in English measure units.
     */
    ApiDrawing.prototype.SetPosition = function(nPosX, nPosY)
    {
        var fPosX = private_EMU2MM(nPosX);
        var fPosY = private_EMU2MM(nPosY);
        if(this.Drawing && this.Drawing.spPr && this.Drawing.spPr.xfrm)
        {
            this.Drawing.spPr.xfrm.setOffX(fPosX);
            this.Drawing.spPr.xfrm.setOffY(fPosY);
        }
    };

    /**
     * Gets the parent of drawing.
     * @typeofeditors ["CPE"]
     * @returns {ApiSlide | ApiLayout | ApiMaster | null}
     */
    ApiDrawing.prototype.GetParent = function()
    {
        if (this.Drawing && this.Drawing.parent)
        {
            switch(this.Drawing.parent.getObjectType())
            {
                case AscDFH.historyitem_type_Slide:
                    return new ApiSlide(this.Drawing.parent);
                case AscDFH.historyitem_type_SlideLayout:
                    return new ApiLayout(this.Drawing.parent);
                case AscDFH.historyitem_type_SlideMaster:
                    return new ApiMaster(this.Drawing.parent);
            }
        }

        return null;
    };
    
    /**
     * Gets the parent slide.
     * @typeofeditors ["CPE"]
     * @returns {ApiSlide | null} - return null if parent ins't slide.
     */
    ApiDrawing.prototype.GetParentSlide = function()
    {
        if (this.Drawing && this.Drawing.parent && this.Drawing.parent.getObjectType() === AscDFH.historyitem_type_Slide)
        {
            return new ApiSlide(this.Drawing.parent);
        }

        return null;
    };

    /**
     * Gets the parent slide layout.
     * @typeofeditors ["CPE"]
     * @returns {ApiLayout | null} - return null if parent ins't slide layout.
     */
    ApiDrawing.prototype.GetParentLayout = function()
    {
        if (this.Drawing && this.Drawing.parent && this.Drawing.parent.getObjectType() === AscDFH.historyitem_type_SlideLayout)
        {
            return new ApiLayout(this.Drawing.parent);
        }

        return null;
    };

    /**
     * Gets the parent slide master.
     * @typeofeditors ["CPE"]
     * @returns {ApiMaster | null} - return null if parent ins't slide master.
     */
    ApiDrawing.prototype.GetParentMaster = function()
    {
        if (this.Drawing && this.Drawing.parent && this.Drawing.parent.getObjectType() === AscDFH.historyitem_type_SlideMaster)
        {
            return new ApiMaster(this.Drawing.parent);
        }

        return null;
    };

    /**
     * Creates a copy of specified drawing object.
     * @typeofeditors ["CPE"]
     * @returns {ApiDrawing} - return null if drawing doesn't exist.
     */
    ApiDrawing.prototype.Copy = function()
    {
        if (this.Drawing)
            return new ApiDrawing(this.Drawing.copy());

        return null;
    };

    /**
     * Deletes the specified drawing object from parent.
     * @typeofeditors ["CPE"]
     * @returns {bool} - false if drawing doesn't exist or drawing haven't parent.
     */
    ApiDrawing.prototype.Delete = function()
    {
        var oParent = this.GetParent();
        if (this.Drawing && oParent)
        {
            var drawingObjects = oParent.GetAllDrawings();
            for (var nDrawing = 0; nDrawing < drawingObjects.length; nDrawing++)
            {
                if (this.Drawing.Id === drawingObjects[nDrawing].Drawing.Id)
                {
                    oParent.RemoveObject(nDrawing, 1);
                    return true;
                }
            }
        }
        
        return false;
    };

    /**
     * Sets the specified placeholder to the shape.
     * @typeofeditors ["CPE"]
     * @param {ApiPlaceholder} oPlaceholder - The type of the vertical alignment for the shape inner contents.
     * @returns {bool} - returns false if param doesn't placeholder.
     */
    ApiDrawing.prototype.SetPlaceholder = function(oPlaceholder)
    {
        if (!this.Drawing || !oPlaceholder || !oPlaceholder.GetClassType || !oPlaceholder.GetClassType() === "placeholder")
            return false;

        var drawingNvPr = null;

        var drawingParent       = this.GetParent();
        var allDrawingsInParent = null;

        editor.private_checkDrawingUniNvPr(this);

        switch (this.Drawing.getObjectType())
        {
            case AscDFH.historyitem_type_ChartSpace:
            case AscDFH.historyitem_type_GraphicFrame:
                drawingNvPr = this.Drawing.nvGraphicFramePr.nvPr;
                break;
            case AscDFH.historyitem_type_GroupShape:
                drawingNvPr = this.Drawing.nvGrpSpPr;
                break;
            case AscDFH.historyitem_type_ImageShape:
            case AscDFH.historyitem_type_OleObject:
                drawingNvPr = this.Drawing.nvPicPr.nvPr;
                break;
            case AscDFH.historyitem_type_Shape:
            case AscDFH.historyitem_type_Cnx:
                drawingNvPr = this.Drawing.nvSpPr.nvPr;
                break;
        }

        if (!drawingNvPr)
            return false;

        drawingNvPr.setPh(oPlaceholder.Placeholder);
        if (drawingParent)
        {
            allDrawingsInParent = drawingParent.GetAllDrawings();
            for (var nDrawing = 0; nDrawing < allDrawingsInParent.length; nDrawing++)
            {
                if (allDrawingsInParent[nDrawing].Drawing.Id === this.Drawing.Id)
                {
                    editor.private_checkPlaceholders(drawingParent, oPlaceholder);
                    break;
                }
            }
        }

        return true;
    };

    /**
     * Gets the placeholder of the specified shape.
     * @typeofeditors ["CPE"]
     * @returns {ApiPlaceholder | null} - returns null if placeholder doesn't exist.
     */
    ApiDrawing.prototype.GetPlaceholder = function()
    {
        var oPh = null;

        if (this.Drawing)
        {
            editor.private_checkDrawingUniNvPr(this);
            switch (this.Drawing.getObjectType())
            {
                case AscDFH.historyitem_type_ChartSpace:
                case AscDFH.historyitem_type_GraphicFrame:
                    oPh = this.Drawing.nvGraphicFramePr.nvPr.ph;
                    break;
                case AscDFH.historyitem_type_GroupShape:
                    oPh = this.Drawing.nvGrpSpPr.ph;
                case AscDFH.historyitem_type_ImageShape:
                case AscDFH.historyitem_type_OleObject:
                    oPh = this.Drawing.nvPicPr.nvPr.ph;
                    break;
                case AscDFH.historyitem_type_Shape:
                case AscDFH.historyitem_type_Cnx:
                    oPh = this.Drawing.nvSpPr.nvPr.ph;
                    break;
            }
        }

        if (oPh)
            return new ApiPlaceholder(oPh);

        return null;
    };

    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiImage
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * Get the type of this class.
     * @returns {"image"}
     */
    ApiImage.prototype.GetClassType = function()
    {
        return "image";
    };

    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiShape
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * Get the type of this class.
     * @typeofeditors ["CPE"]
     * @returns {"shape"}
     */
    ApiShape.prototype.GetClassType = function()
    {
        return "shape";
    };


    /**
     * Deprecated in 6.2.
     * Get the shape inner contents where a paragraph or text runs can be inserted. 
     * @typeofeditors ["CPE"]
     * @returns {?ApiDocumentContent}
     */
    ApiShape.prototype.GetDocContent = function()
    {
        var oApi = private_GetApi();
        if(oApi && this.Drawing && this.Drawing.txBody && this.Drawing.txBody.content)
        {
            return oApi.private_CreateApiDocContent(this.Drawing.txBody.content);
        }
        return null;
    };
    
    /**
     * Get the shape inner contents where a paragraph or text runs can be inserted. 
     * @typeofeditors ["CPE"]
     * @returns {?ApiDocumentContent}
     */
    ApiShape.prototype.GetContent = function()
    {
        var oApi = private_GetApi();
        if(oApi && this.Drawing && this.Drawing.txBody && this.Drawing.txBody.content)
        {
            return oApi.private_CreateApiDocContent(this.Drawing.txBody.content);
        }
        return null;
    };

    /**
     * Set the vertical alignment for the shape content where a paragraph or text runs can be inserted.
     * @typeofeditors ["CPE"]
     * @param {VerticalTextAlign} VerticalAlign - The type of the vertical alignment for the shape inner contents.
     */
    ApiShape.prototype.SetVerticalTextAlign = function(VerticalAlign)
    {
        if(this.Shape)
        {
            switch(VerticalAlign)
            {
                case "top":
                {
                    this.Shape.setVerticalAlign(4);
                    break;
                }
                case "center":
                {
                    this.Shape.setVerticalAlign(1);
                    break;
                }
                case "bottom":
                {
                    this.Shape.setVerticalAlign(0);
                    break;
                }
            }
        }
    };

    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiChart
    //
    //------------------------------------------------------------------------------------------------------------------
    /**
     * Get the type of this class.
     * @typeofeditors ["CPE"]
     * @returns {"chart"}
     */
    ApiChart.prototype.GetClassType = function()
    {
        return "chart";
    };

    /**
     *  Specify the chart title.
     *  @typeofeditors ["CPE"]
     *  @param {string} sTitle - The title which will be displayed for the current chart.
     *  @param {pt} nFontSize - The text size value measured in points.
     *  @param {?bool} bIsBold - Specifies if the chart title is written in bold font or not.
     */
    ApiChart.prototype.SetTitle = function (sTitle, nFontSize, bIsBold)
    {
        AscFormat.builder_SetChartTitle(this.Chart, sTitle, nFontSize, bIsBold);
    };

    /**
     *  Specify the chart horizontal axis title.
     *  @typeofeditors ["CPE"]
     *  @param {string} sTitle - The title which will be displayed for the horizontal axis of the current chart.
     *  @param {pt} nFontSize - The text size value measured in points.
     *  @param {?bool} bIsBold - Specifies if the horizontal axis title is written in bold font or not.
     * */
    ApiChart.prototype.SetHorAxisTitle = function (sTitle, nFontSize, bIsBold)
    {
        AscFormat.builder_SetChartHorAxisTitle(this.Chart, sTitle, nFontSize, bIsBold);
    };

    /**
     *  Specify the chart vertical axis title.
     *  @typeofeditors ["CPE"]
     *  @param {string} sTitle - The title which will be displayed for the vertical axis of the current chart.
     *  @param {pt} nFontSize - The text size value measured in points.
     *  @param {?bool} bIsBold - Specifies if the vertical axis title is written in bold font or not.
     * */
    ApiChart.prototype.SetVerAxisTitle = function (sTitle, nFontSize, bIsBold)
    {
        AscFormat.builder_SetChartVertAxisTitle(this.Chart, sTitle, nFontSize, bIsBold);
    };

    /**
     * Specify the chart legend position.
     * @typeofeditors ["CPE"]
     * @param {"left" | "top" | "right" | "bottom" | "none"} sLegendPos - The position of the chart legend inside the chart window.
     * */
    ApiChart.prototype.SetLegendPos = function(sLegendPos)
    {
        AscFormat.builder_SetChartLegendPos(this.Chart, sLegendPos);
    };

    /**
     * Specify a legend font size.
     * @param {pt} nFontSize - The text size value measured in points.
     * */
    ApiChart.prototype.SetLegendFontSize = function(nFontSize)
    {
        AscFormat.builder_SetLegendFontSize(this.Chart, nFontSize);
    };

    /**
     * Specify a vertical axis orientation.
     * @param {bool} bIsMinMax - The <code>true</code> value will set the normal data direction for the horizontal axis
	 * (from minimum to maximum). The <code>false</code> value will set the inverted data direction for the horizontal axis (from maximum to minimum).
     * */
    ApiChart.prototype.SetVerAxisOrientation = function(bIsMinMax){
        AscFormat.builder_SetChartVertAxisOrientation(this.Chart, bIsMinMax);
    };

    /**
     * Specify a horizontal axis orientation.
     * @param {bool} bIsMinMax - The <code>true</code> value will set the normal data direction for the horizontal axis
	 * (from minimum to maximum). The <code>false</code> value will set the inverted data direction for the horizontal axis (from maximum to minimum).
     * */
    ApiChart.prototype.SetHorAxisOrientation = function(bIsMinMax){
        AscFormat.builder_SetChartHorAxisOrientation(this.Chart, bIsMinMax);
    };

    /**
     *  Specify which chart data labels are shown for the chart.
     * @typeofeditors ["CPE"]
     * @param {boolean} bShowSerName - Whether to show or hide the source table column names used for the data which the chart will be build from.
     * @param {boolean} bShowCatName - Whether to show or hide the source table row names used for the data which the chart will be build from.
     * @param {boolean} bShowVal - Whether to show or hide the chart data values.
     * @param {boolean} bShowPercent - Whether to show or hide the percent for the data values (works with stacked chart types).
     * */
    ApiChart.prototype.SetShowDataLabels = function(bShowSerName, bShowCatName, bShowVal, bShowPercent)
    {
        AscFormat.builder_SetShowDataLabels(this.Chart, bShowSerName, bShowCatName, bShowVal, bShowPercent);
    };

    /**
     * Spicify the show options for data labels.
     * @param {number} nSeriesIndex - The series index from the array of the data used to build the chart from.
     * @param {number} nPointIndex - The point index from this series.
     * @param {boolean} bShowSerName - Whether to show or hide the source table column names used for the data which the chart will be build from.
     * @param {boolean} bShowCatName - Whether to show or hide the source table row names used for the data which the chart will be build from.
     * @param {boolean} bShowVal - Whether to show or hide the chart data values.
     * @param {boolean} bShowPercent - Whether to show or hide the percent for the data values (works with stacked chart types).
     * */
    ApiChart.prototype.SetShowPointDataLabel = function(nSeriesIndex, nPointIndex, bShowSerName, bShowCatName, bShowVal, bShowPercent)
    {
        AscFormat.builder_SetShowPointDataLabel(this.Chart, nSeriesIndex, nPointIndex, bShowSerName, bShowCatName, bShowVal, bShowPercent);
    };

    /**
     * Spicify tick labels position for vertical axis.
     * @param {TickLabelPosition} sTickLabelPosition - The type for the position of chart vertical tick labels.
     * */
    ApiChart.prototype.SetVertAxisTickLabelPosition = function(sTickLabelPosition)
    {
        AscFormat.builder_SetChartVertAxisTickLablePosition(this.Chart, sTickLabelPosition);
    };
    /**
     * Spicify tick labels position for horizontal axis.
     * @param {TickLabelPosition} sTickLabelPosition - The type for the position of chart horizontal tick labels.
     * */
    ApiChart.prototype.SetHorAxisTickLabelPosition = function(sTickLabelPosition)
    {
        AscFormat.builder_SetChartHorAxisTickLablePosition(this.Chart, sTickLabelPosition);
    };




    /**
     * Specify major tick mark for horizontal axis.
     * @param {TickMark} sTickMark - The type of tick mark appearance.
     * */

    ApiChart.prototype.SetHorAxisMajorTickMark = function(sTickMark){
        AscFormat.builder_SetChartHorAxisMajorTickMark(this.Chart, sTickMark);
    };
    /**
     * Specify minor tick mark for horizontal axis.
     * @param {TickMark} sTickMark - The type of tick mark appearance.
     * */

    ApiChart.prototype.SetHorAxisMinorTickMark = function(sTickMark){
        AscFormat.builder_SetChartHorAxisMinorTickMark(this.Chart, sTickMark);
    };

    /**
     * Specify major tick mark for vertical axis.
     * @param {TickMark} sTickMark - The type of tick mark appearance.
     * */

    ApiChart.prototype.SetVertAxisMajorTickMark = function(sTickMark){
        AscFormat.builder_SetChartVerAxisMajorTickMark(this.Chart, sTickMark);
    };

    /**
     * Specify minor tick mark for vertical axis.
     * @param {TickMark} sTickMark - The type of tick mark appearance.
     * */
    ApiChart.prototype.SetVertAxisMinorTickMark = function(sTickMark){
        AscFormat.builder_SetChartVerAxisMinorTickMark(this.Chart, sTickMark);
    };




    /**
     * Specify major vertical gridline's visual properties.
     * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
     * */
    ApiChart.prototype.SetMajorVerticalGridlines = function(oStroke)
    {
        AscFormat.builder_SetVerAxisMajorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
    };

    /**
     * Specify minor vertical gridline's visual properties.
     * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
     * */
    ApiChart.prototype.SetMinorVerticalGridlines = function(oStroke)
    {
        AscFormat.builder_SetVerAxisMinorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
    };


    /**
     * Specify major horizontal gridline's visual properties.
     * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
     * */
    ApiChart.prototype.SetMajorHorizontalGridlines = function(oStroke)
    {
        AscFormat.builder_SetHorAxisMajorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
    };

    /**
     * Specify minor vertical gridline's visual properties.
     * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
     * */
    ApiChart.prototype.SetMinorHorizontalGridlines = function(oStroke)
    {
        AscFormat.builder_SetHorAxisMinorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
    };


    /**
     * Specify font size for labels of horizontal axis.
     * @param {pt} nFontSize - The text size value measured in points.
     */
    ApiChart.prototype.SetHorAxisLablesFontSize = function(nFontSize){
        AscFormat.builder_SetHorAxisFontSize(this.Chart, nFontSize);
    };

    /**
     * Specify font size for labels of vertical axis.
     * @param {pt} nFontSize - The text size value measured in points.
     */
    ApiChart.prototype.SetVertAxisLablesFontSize = function(nFontSize){
        AscFormat.builder_SetVerAxisFontSize(this.Chart, nFontSize);
    };


    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiTable
    //
    //------------------------------------------------------------------------------------------------------------------
    /**
     * Return type of object.
     * @returns {"table"}
     * */
    ApiTable.prototype.GetClassType = function(){
        return "table";
    };


    /**
     * Return row by index.
     * @param nIndex {number} - The row number (position) in the table.
     * @returns {?ApiTableRow}
     * */
    ApiTable.prototype.GetRow = function(nIndex){
        if(!this.Drawing){
            return null;
        }
        var aTableContent = this.Table.Content;
        if(!aTableContent[nIndex]){
            return null;
        }
        return new ApiTableRow(aTableContent[nIndex]);
    };

    /**
     * Merge array of cells. If merge was done successfully, it will return merged cell, otherwise "null".
     * <b>Warning</b>: The number of cells in any row and the number of rows in the current table may be changed.
     * @param {ApiTableCell[]} aCells - The array of cells.
     * @returns {?ApiTableCell}
     */
    ApiTable.prototype.MergeCells = function(aCells)
    {
        this.private_PrepareTableForActions();

        var oTable            = this.Table;
        oTable.Selection.Use  = true;
        oTable.Selection.Type = table_Selection_Cell;
        oTable.Selection.Data = [];

        for (var nPos = 0, nCount = aCells.length; nPos < nCount; ++nPos)
        {
            var oCell = aCells[nPos].Cell;
            var oPos  = {Cell : oCell.Index, Row : oCell.Row.Index};

            var nResultPos    = 0;
            var nResultLength = oTable.Selection.Data.length;
            for (nResultPos = 0; nResultPos < nResultLength; ++nResultPos)
            {
                var oCurPos = oTable.Selection.Data[nResultPos];
                if (oCurPos.Row < oPos.Row)
                {
                    continue;
                }
                else if (oCurPos.Row > oPos.Row)
                {
                    break;
                }
                else
                {
                    if (oCurPos.Cell < oPos.Cell)
                        continue;
                    else
                        break;
                }
            }

            oTable.Selection.Data.splice(nResultPos, 0, oPos);
        }

        var isMerged = this.Table.MergeTableCells(true);
        var oMergedCell = this.Table.CurCell;
        oTable.RemoveSelection();


        if (true === isMerged)
            return new ApiTableCell(oMergedCell);

        return null;
    };

    ApiTable.prototype.OnChangeTablePr = function(oApiTablePr)
    {
        this.Table.Set_Pr(oApiTablePr.TablePr);
        oApiTablePr.TablePr = this.Table.Pr.Copy();
    };
    ApiTable.prototype.private_PrepareTableForActions = function()
    {
        this.Table.private_RecalculateGrid();
        this.Table.private_UpdateCellsGrid();
    };
    /**
     * Specify the components of the conditional formatting of the referenced table style (if one exists)
     * which shall be applied to the set of table rows with the current table-level property exceptions. A table style
     * can specify up to six different optional conditional formats [Example: Different formatting for first column],
     * which then can be applied or omitted from individual table rows in the parent table.
     *
     * The default setting is to apply the row and column banding formatting, but not the first row, last row, first
     * column, or last column formatting.
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
     */
    ApiTable.prototype.SetTableLook = function(isFirstColumn, isFirstRow, isLastColumn, isLastRow, isHorBand, isVerBand)
    {
        var oTableLook = new CTableLook(private_GetBoolean(isFirstColumn),
            private_GetBoolean(isFirstRow),
            private_GetBoolean(isLastColumn),
            private_GetBoolean(isLastRow),
            private_GetBoolean(isHorBand),
            private_GetBoolean(isVerBand));
        this.Table.Set_TableLook(oTableLook);
    };
    /**
     * Add a new row to the current table.
     * @param {ApiTableCell} [oCell] - If not specified a new row will be added to the end of the table.
     * @param {boolean} [isBefore=false] - Add a new row before or after the specified cell. If no cell is specified,
     * then this parameter will be ignored.
     * @returns {ApiTableRow}
     */
    ApiTable.prototype.AddRow = function(oCell, isBefore)
    {
        this.private_PrepareTableForActions();

        var _isBefore = private_GetBoolean(isBefore, false);
        var _oCell = (oCell instanceof ApiTableCell ? oCell.Cell : undefined);
        if (_oCell && this.Table !== _oCell.Row.Table)
            _oCell = undefined;

        if (!_oCell)
        {
            _oCell = this.Table.Content[this.Table.Content.length - 1].Get_Cell(0);
            _isBefore = false;
        }

        var nRowIndex = true === _isBefore ? _oCell.Row.Index : _oCell.Row.Index + 1;

        this.Table.RemoveSelection();
        this.Table.CurCell = _oCell;
        this.Table.AddTableRow(_isBefore);

        return new ApiTableRow(this.Table.Content[nRowIndex]);
    };
    /**
     * Add a new column to the end of the current table.
     * @param {ApiTableCell} [oCell] - If not specified a new column will be added to the end of the table.
     * @param {boolean} [isBefore=false] - Add a new column before or after the specified cell. If no cell is specified,
     * then this parameter will be ignored.
     */
    ApiTable.prototype.AddColumn = function(oCell, isBefore)
    {
        this.private_PrepareTableForActions();

        var _isBefore = private_GetBoolean(isBefore, false);
        var _oCell = (oCell instanceof ApiTableCell ? oCell.Cell : undefined);
        if (_oCell && this.Table !== _oCell.Row.Table)
            _oCell = undefined;

        if (!_oCell)
        {
            _oCell = this.Table.Content[0].Get_Cell(this.Table.Content[0].Get_CellsCount() - 1);
            _isBefore = false;
        }

        this.Table.RemoveSelection();
        this.Table.CurCell = _oCell;
        this.Table.AddTableColumn(_isBefore);
    };
    /**
     * Remove the table row with a specified cell.
     * @param {ApiTableCell} oCell - The table cell specified.
     * @returns {boolean} - defines if the table is empty after removing or not.
     */
    ApiTable.prototype.RemoveRow = function(oCell)
    {
        if (!(oCell instanceof ApiTableCell) || this.Table !== oCell.Cell.Row.Table)
            return false;
        this.private_PrepareTableForActions();
        this.Table.RemoveSelection();
        this.Table.CurCell = oCell.Cell;
        var isEmpty = !(this.Table.RemoveTableRow());
        return isEmpty;
    };
    /**
     * Remove the table column with a specified cell.
     * @param {ApiTableCell} oCell - The table cell specified.
     * @returns {boolean} - defines if the table is empty after removing or not.
     */
    ApiTable.prototype.RemoveColumn = function(oCell)
    {
        if (!(oCell instanceof ApiTableCell) || this.Table !== oCell.Cell.Row.Table)
            return false;
        this.private_PrepareTableForActions();
        this.Table.RemoveSelection();
        this.Table.CurCell = oCell.Cell;
        var isEmpty = !(this.Table.RemoveTableColumn());

        return isEmpty;
    };

    /**
     * Specify the shading which shall be applied to the extents of the current table.
     * @typeofeditors ["CPE"]
	 * @param {ShdType | ApiFill} sType - The shading type applied to the contents of the current table. Can be ShdType or ApiFill.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
    ApiTable.prototype.SetShd = function(sType, r, g, b)
    {
        var oPr    = this.Table.Pr.Copy();
        var color  = new Asc.asc_CColor({r : r, g: g, b: b, Auto : false});
        var oShd   = new CDocumentShd();
        var _Shd   = null;

        if (sType === "nil") {
            _Shd = {Value : Asc.c_oAscShdNil};
            oShd.Set_FromObject(_Shd);
            oPr.Shd = oShd;
        }
        else if (sType === "clear") {

            var Unifill        = new AscFormat.CUniFill();
			Unifill.fill       = new AscFormat.CSolidFill();
			Unifill.fill.color = AscFormat.CorrectUniColor(color, Unifill.fill.color, 1);
			_Shd = {
				Value   : Asc.c_oAscShdClear,
				Color   : {
					r : color.asc_getR(),
					g : color.asc_getG(),
					b : color.asc_getB()
				},
				Unifill : Unifill
			};
			
			oShd.Set_FromObject(_Shd);
            oPr.Shd = oShd;
        }
        else if (sType.GetClassType && sType.GetClassType() === "fill") {
            oShd.Value = Asc.c_oAscShdClear;
            oShd.Unifill = sType.UniFill;
            oPr.Shd = oShd;
        }
        else 
            oPr.Shd = null;
        
        this.Table.Set_Pr(oPr);
    };


    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiTableRow
    //
    //------------------------------------------------------------------------------------------------------------------
    /**
     * Get the type of this class.
     * @returns {"tableRow"}
     */
    ApiTableRow.prototype.GetClassType = function()
    {
        return "tableRow";
    };
    /**
     * Get the number of cells in the current row.
     * @returns {number}
     */
    ApiTableRow.prototype.GetCellsCount = function()
    {
        return this.Row.Content.length;
    };
    /**
     * Get the cell by its position.
     * @param {number} nPos - The cell position in the table row.
     * @returns {ApiTableCell}
     */
    ApiTableRow.prototype.GetCell = function(nPos)
    {
        if (nPos < 0 || nPos >= this.Row.Content.length)
            return null;

        return new ApiTableCell(this.Row.Content[nPos]);
    };


    /**
     * Set the height of the current table row within the current table.
     * @param {EMU} [nValue] - The row height in English measure units.
     */
    ApiTableRow.prototype.SetHeight = function(nValue)
    {
        var fMaxTopMargin = 0, fMaxBottomMargin = 0, fMaxTopBorder = 0, fMaxBottomBorder = 0;

        for (var i = 0;  i < this.Row.Content.length; ++i){
            var oCell = this.Row.Content[i];
            var oMargins = oCell.GetMargins();
            if(oMargins.Bottom.W > fMaxBottomMargin){
                fMaxBottomMargin = oMargins.Bottom.W;
            }
            if(oMargins.Top.W > fMaxTopMargin){
                fMaxTopMargin = oMargins.Top.W;
            }
            var oBorders = oCell.Get_Borders();
            if(oBorders.Top.Size > fMaxTopBorder){
                fMaxTopBorder = oBorders.Top.Size;
            }
            if(oBorders.Bottom.Size > fMaxBottomBorder){
                fMaxBottomBorder = oBorders.Bottom.Size;
            }
        }
        this.Row.Set_Height(Math.max(1, nValue/36000 - fMaxTopMargin - fMaxBottomMargin - fMaxTopBorder/2 - fMaxBottomBorder/2), Asc.linerule_AtLeast);
    };



    //------------------------------------------------------------------------------------------------------------------
    //
    // ApiTableRow
    //
    //------------------------------------------------------------------------------------------------------------------

    /**
     * Get the type of this class.
     * @returns {"tableCell"}
     */
    ApiTableCell.prototype.GetClassType = function()
    {
        return "tableCell";
    };

    /**
     * Returns cell content.
     * @returns {ApiDocumentContent}
     */
    ApiTableCell.prototype.GetContent = function(){
        var oApi = private_GetApi();
        return oApi.private_CreateApiDocContent(this.Cell.Content);

    };


    /**
     * Specify the shading which shall be applied to the extents of the current table cell.
     * @typeofeditors ["CPE"]
	 * @param {ShdType | ApiFill} sType - The shading type applied to the contents of the current table. Can be ShdType or ApiFill.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
    ApiTableCell.prototype.SetShd = function(sType, r, g, b)
    {
        var oPr    = this.Cell.Pr.Copy();
        var color  = new Asc.asc_CColor({r : r, g: g, b: b, Auto : false});
        var oShd   = new CDocumentShd();
        var _Shd   = null;

        if (sType === "nil") {
            _Shd = {Value : Asc.c_oAscShdNil};
            oShd.Set_FromObject(_Shd);
            oPr.Shd = oShd;
        }
        else if (sType === "clear") {

            var Unifill        = new AscFormat.CUniFill();
			Unifill.fill       = new AscFormat.CSolidFill();
			Unifill.fill.color = AscFormat.CorrectUniColor(color, Unifill.fill.color, 1);
			_Shd = {
				Value   : Asc.c_oAscShdClear,
				Color   : {
					r : color.asc_getR(),
					g : color.asc_getG(),
					b : color.asc_getB()
				},
				Unifill : Unifill
			};
			
			oShd.Set_FromObject(_Shd);
            oPr.Shd = oShd;
        }
        else if (sType.GetClassType && sType.GetClassType() === "fill") {
            oShd.Value = Asc.c_oAscShdClear;
            oShd.Unifill = sType.UniFill;
            oPr.Shd = oShd;
        }
        else 
            oPr.Shd = null;

        this.Cell.Set_Pr(oPr);
    };


    /**
     * Specify the amount of space which shall be left between the bottom extent of the cell contents and the border
     * of a specific individual table cell within a table.
     * @param {?twips} nValue - If this value is <code>null</code>, then default table cell bottom margin shall be used,
     * otherwise override the table cell bottom margin with specified value for the current cell.
     */
    ApiTableCell.prototype.SetCellMarginBottom = function(nValue)
    {
        var oPr = this.Cell.Pr.Copy();
        if (!oPr.TableCellMar)
        {
            oPr.TableCellMar =
                {
                    Bottom : undefined,
                    Left   : undefined,
                    Right  : undefined,
                    Top    : undefined
                };
        }

        if (null === nValue)
            oPr.TableCellMar.Bottom = undefined;
        else
            oPr.TableCellMar.Bottom = private_GetTableMeasure("twips", nValue);
        this.Cell.Set_Pr(oPr);
    };
    /**
     * Specify the amount of space which shall be left between the left extent of the current cell contents and the
     * left edge border of a specific individual table cell within a table.
     * @param {?twips} nValue - If this value is <code>null</code>, then default table cell bottom margin shall be used,
     * otherwise override the table cell bottom margin with specified value for the current cell.
     */
    ApiTableCell.prototype.SetCellMarginLeft = function(nValue)
    {
        var oPr = this.Cell.Pr.Copy();
        if (!oPr.TableCellMar)
        {
            oPr.TableCellMar =
                {
                    Bottom : undefined,
                    Left   : undefined,
                    Right  : undefined,
                    Top    : undefined
                };
        }

        if (null === nValue)
            oPr.TableCellMar.Left = undefined;
        else
            oPr.TableCellMar.Left = private_GetTableMeasure("twips", nValue);
        this.Cell.Set_Pr(oPr);
    };
    /**
     * Specifies the amount of space which shall be left between the right extent of the current cell contents and the
     * right edge border of a specific individual table cell within a table.
     * @param {?twips} nValue - If this value is <code>null</code>, then default table cell bottom margin shall be used,
     * otherwise override the table cell bottom margin with specified value for the current cell.
     */
    ApiTableCell.prototype.SetCellMarginRight = function(nValue)
    {
        var oPr = this.Cell.Pr.Copy();
        if (!oPr.TableCellMar)
        {
            oPr.TableCellMar =
                {
                    Bottom : undefined,
                    Left   : undefined,
                    Right  : undefined,
                    Top    : undefined
                };
        }

        if (null === nValue)
            oPr.TableCellMar.Right = undefined;
        else
            oPr.TableCellMar.Right = private_GetTableMeasure("twips", nValue);
        this.Cell.Set_Pr(oPr);
    };
    /**
     * Specifies the amount of space which shall be left between the top extent of the current cell contents and the
     * top edge border of a specific individual table cell within a table.
     * @param {?twips} nValue - If this value is <code>null</code>, then default table cell bottom margin shall be used,
     * otherwise override the table cell bottom margin with specified value for the current cell.
     */
    ApiTableCell.prototype.SetCellMarginTop = function(nValue)
    {
        var oPr = this.Cell.Pr.Copy();
        if (!oPr.TableCellMar)
        {
            oPr.TableCellMar =
                {
                    Bottom : undefined,
                    Left   : undefined,
                    Right  : undefined,
                    Top    : undefined
                };
        }

        if (null === nValue)
            oPr.TableCellMar.Top = undefined;
        else
            oPr.TableCellMar.Top = private_GetTableMeasure("twips", nValue);
        this.Cell.Set_Pr(oPr);
    };
    /**
     * Set the border which shall be displayed at the bottom of the current table cell.
     * @param {mm} fSize - The width of the current border.
     * @param {ApiFill} oApiFill - The color or pattern used to fill the current border.
     */
    ApiTableCell.prototype.SetCellBorderBottom = function(fSize, oApiFill)
    {
        var oBorder = new CDocumentBorder();
        oBorder.Value = border_Single;
        oBorder.Size  = fSize;
        oBorder.Space = 0;
        oBorder.Unifill = oApiFill.UniFill;
        var oPr = this.Cell.Pr.Copy();
        oPr.TableCellBorders.Bottom = oBorder;
        this.Cell.Set_Pr(oPr);
    };

    /**
     * Set the border which shall be displayed at the left of the current table cell.
     * @param {mm} fSize - The width of the current border.
     * @param {ApiFill} oApiFill - The color or pattern used to fill the current border.
     */
    ApiTableCell.prototype.SetCellBorderLeft = function(fSize, oApiFill)
    {
        var oBorder = new CDocumentBorder();
        oBorder.Value = border_Single;
        oBorder.Size  = fSize;
        oBorder.Space = 0;
        oBorder.Unifill = oApiFill.UniFill;
        var oPr = this.Cell.Pr.Copy();
        oPr.TableCellBorders.Left = oBorder;
        this.Cell.Set_Pr(oPr);
    };

    /**
     * Set the border which shall be displayed at the right of the current table cell.
     * @param {mm} fSize - The width of the current border.
     * @param {ApiFill} oApiFill - The color or pattern used to fill the current border.
     */
    ApiTableCell.prototype.SetCellBorderRight = function(fSize, oApiFill)
    {
        var oBorder = new CDocumentBorder();
        oBorder.Value = border_Single;
        oBorder.Size  = fSize;
        oBorder.Space = 0;
        oBorder.Unifill = oApiFill.UniFill;
        var oPr = this.Cell.Pr.Copy();
        oPr.TableCellBorders.Right = oBorder;
        this.Cell.Set_Pr(oPr);
    };

    /**
     * Set the border which shall be displayed at the top of the current table cell.
     * @param {mm} fSize - The width of the current border.
     * @param {ApiFill} oApiFill - The color or pattern used to fill the current border.
     */
    ApiTableCell.prototype.SetCellBorderTop = function(fSize, oApiFill)
    {
        var oBorder = new CDocumentBorder();
        oBorder.Value = border_Single;
        oBorder.Size  = fSize;
        oBorder.Space = 0;
        oBorder.Unifill = oApiFill.UniFill;
        var oPr = this.Cell.Pr.Copy();
        oPr.TableCellBorders.Top = oBorder;
        this.Cell.Set_Pr(oPr);
    };

    /**
     * Specify the vertical alignment for text within the current table cell.
     * @param {("top" | "center" | "bottom")} sType - The type of the vertical alignment.
     */
    ApiTableCell.prototype.SetVerticalAlign = function(sType)
    {
        var oPr = this.Cell.Pr.Copy();
        if ("top" === sType)
            oPr.VAlign = vertalignjc_Top;
        else if ("bottom" === sType)
            oPr.VAlign = vertalignjc_Bottom;
        else if ("center" === sType)
            oPr.VAlign = vertalignjc_Center;
        this.Cell.Set_Pr(oPr);
    };
    /**
     * Specify the direction of the text flow for this table cell.
     * @param {("lrtb" | "tbrl" | "btlr")} sType - The type of the text flow direction. 
     */
    ApiTableCell.prototype.SetTextDirection = function(sType)
    {
        var oPr = this.Cell.Pr.Copy();
        if ("lrtb" === sType)
            oPr.TextDirection = textdirection_LRTB;
        else if ("tbrl" === sType)
            oPr.TextDirection = textdirection_TBRL;
        else if ("btlr" === sType)
            oPr.TextDirection = textdirection_BTLR;
        this.Cell.Set_Pr(oPr);
    };



    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Export
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Api.prototype["GetPresentation"]                      = Api.prototype.GetPresentation;
    Api.prototype["CreateSlide"]                          = Api.prototype.CreateSlide;
    Api.prototype["CreateImage"]                          = Api.prototype.CreateImage;
    Api.prototype["CreateShape"]                          = Api.prototype.CreateShape;
    Api.prototype["CreateChart"]                          = Api.prototype.CreateChart;
    Api.prototype["CreateGroup"]                          = Api.prototype.CreateGroup;
    Api.prototype["CreateTable"]                          = Api.prototype.CreateTable;
    Api.prototype["CreateParagraph"]                      = Api.prototype.CreateParagraph;
    Api.prototype["Save"]                                 = Api.prototype.Save;
    Api.prototype["CreateMaster"]                         = Api.prototype.CreateMaster;
    Api.prototype["CreateLayout"]                         = Api.prototype.CreateLayout;
    Api.prototype["CreatePlaceholder"]                    = Api.prototype.CreatePlaceholder;
    Api.prototype["CreateTheme"]                          = Api.prototype.CreateTheme;
    Api.prototype["CreateThemeColorScheme"]               = Api.prototype.CreateThemeColorScheme;
    Api.prototype["CreateThemeFormatScheme"]              = Api.prototype.CreateThemeFormatScheme;
    Api.prototype["CreateThemeFontScheme"]                = Api.prototype.CreateThemeFontScheme;

    ApiPresentation.prototype["GetClassType"]             = ApiPresentation.prototype.GetClassType;
    ApiPresentation.prototype["GetCurSlideIndex"]         = ApiPresentation.prototype.GetCurSlideIndex;
    ApiPresentation.prototype["GetSlideByIndex"]          = ApiPresentation.prototype.GetSlideByIndex;
    ApiPresentation.prototype["GetCurrentSlide"]          = ApiPresentation.prototype.GetCurrentSlide;
    ApiPresentation.prototype["AddSlide"]                 = ApiPresentation.prototype.AddSlide;
    ApiPresentation.prototype["CreateNewHistoryPoint"]    = ApiPresentation.prototype.CreateNewHistoryPoint;
    ApiPresentation.prototype["SetSizes"]                 = ApiPresentation.prototype.SetSizes;
    ApiPresentation.prototype["ReplaceCurrentImage"]      = ApiPresentation.prototype.ReplaceCurrentImage;
    ApiPresentation.prototype["GetSlidesCount"]           = ApiPresentation.prototype.GetSlidesCount;
    ApiPresentation.prototype["GetMastersCount"]          = ApiPresentation.prototype.GetMastersCount;
    ApiPresentation.prototype["GetMaster"]                = ApiPresentation.prototype.GetMaster;
    ApiPresentation.prototype["AddMaster"]                = ApiPresentation.prototype.AddMaster;
    ApiPresentation.prototype["ApplyTheme"]               = ApiPresentation.prototype.ApplyTheme;
    ApiPresentation.prototype["GetDefaultTextPr"]         = ApiPresentation.prototype.GetDefaultTextPr;
    ApiPresentation.prototype["RemoveSlides"]             = ApiPresentation.prototype.RemoveSlides;
    ApiPresentation.prototype["SetLanguage"]           = ApiPresentation.prototype.SetLanguage;

    ApiMaster.prototype["GetClassType"]                   = ApiMaster.prototype.GetClassType;
    ApiMaster.prototype["GetLayout"]                      = ApiMaster.prototype.GetLayout;
    ApiMaster.prototype["AddLayout"]                      = ApiMaster.prototype.AddLayout;
    ApiMaster.prototype["RemoveLayout"]                   = ApiMaster.prototype.RemoveLayout;
    ApiMaster.prototype["GetLayoutsCount"]                = ApiMaster.prototype.GetLayoutsCount;
    ApiMaster.prototype["AddObject"]                      = ApiMaster.prototype.AddObject;
    ApiMaster.prototype["RemoveObject"]                   = ApiMaster.prototype.RemoveObject;
    ApiMaster.prototype["SetBackground"]                  = ApiMaster.prototype.SetBackground;
    ApiMaster.prototype["ClearBackground"]                = ApiMaster.prototype.ClearBackground;
    ApiMaster.prototype["Copy"]                           = ApiMaster.prototype.Copy;
    ApiMaster.prototype["Duplicate"]                      = ApiMaster.prototype.Duplicate;
    ApiMaster.prototype["Delete"]                         = ApiMaster.prototype.Delete;
    ApiMaster.prototype["GetTheme"]                       = ApiMaster.prototype.GetTheme;
    ApiMaster.prototype["SetTheme"]                       = ApiMaster.prototype.SetTheme;
    ApiMaster.prototype["GetAllDrawings"]                 = ApiMaster.prototype.GetAllDrawings;
    ApiMaster.prototype["GetAllShapes"]                   = ApiMaster.prototype.GetAllShapes;
    ApiMaster.prototype["GetAllImages"]                   = ApiMaster.prototype.GetAllImages;
    ApiMaster.prototype["GetAllCharts"]                   = ApiMaster.prototype.GetAllCharts;
    
    
    ApiLayout.prototype["GetClassType"]                   = ApiLayout.prototype.GetClassType;
    ApiLayout.prototype["SetName"]                        = ApiLayout.prototype.SetName;
    ApiLayout.prototype["AddObject"]                      = ApiLayout.prototype.AddObject;
    ApiLayout.prototype["RemoveObject"]                   = ApiLayout.prototype.RemoveObject;
    ApiLayout.prototype["SetBackground"]                  = ApiLayout.prototype.SetBackground;
    ApiLayout.prototype["ClearBackground"]                = ApiLayout.prototype.ClearBackground;
    ApiLayout.prototype["FollowMasterBackground"]         = ApiLayout.prototype.FollowMasterBackground;
    ApiLayout.prototype["Copy"]                           = ApiLayout.prototype.Copy;
    ApiLayout.prototype["Delete"]                         = ApiLayout.prototype.Delete;
    ApiLayout.prototype["Duplicate"]                      = ApiLayout.prototype.Duplicate;
    ApiLayout.prototype["MoveTo"]                         = ApiLayout.prototype.MoveTo;
    ApiLayout.prototype["GetAllDrawings"]                 = ApiLayout.prototype.GetAllDrawings;
    ApiLayout.prototype["GetAllShapes"]                   = ApiLayout.prototype.GetAllShapes;
    ApiLayout.prototype["GetAllImages"]                   = ApiLayout.prototype.GetAllImages;
    ApiLayout.prototype["GetAllCharts"]                   = ApiLayout.prototype.GetAllCharts;
    ApiLayout.prototype["GetMaster"]                      = ApiLayout.prototype.GetMaster;

    ApiPlaceholder.prototype["GetClassType"]              = ApiPlaceholder.prototype.GetClassType;
    ApiPlaceholder.prototype["SetType"]                   = ApiPlaceholder.prototype.SetType;

    ApiTheme.prototype["GetClassType"]                    = ApiTheme.prototype.GetClassType;
    ApiTheme.prototype["GetMaster"]                       = ApiTheme.prototype.GetMaster;
    ApiTheme.prototype["SetColorScheme"]                  = ApiTheme.prototype.SetColorScheme;
    ApiTheme.prototype["GetColorScheme"]                  = ApiTheme.prototype.GetColorScheme;
    ApiTheme.prototype["SetFormatScheme"]                 = ApiTheme.prototype.SetFormatScheme;
    ApiTheme.prototype["GetFormatScheme"]                 = ApiTheme.prototype.GetFormatScheme;
    ApiTheme.prototype["SetFontScheme"]                   = ApiTheme.prototype.SetFontScheme;
    ApiTheme.prototype["GetFontScheme"]                   = ApiTheme.prototype.GetFontScheme;

    ApiThemeColorScheme.prototype["GetClassType"]         = ApiThemeColorScheme.prototype.GetClassType;
    ApiThemeColorScheme.prototype["SetSchemeName"]        = ApiThemeColorScheme.prototype.SetSchemeName;
    ApiThemeColorScheme.prototype["ChangeColor"]          = ApiThemeColorScheme.prototype.ChangeColor;
    ApiThemeColorScheme.prototype["Copy"]                 = ApiThemeColorScheme.prototype.Copy;

    ApiThemeFormatScheme.prototype["GetClassType"]        = ApiThemeFormatScheme.prototype.GetClassType;
    ApiThemeFormatScheme.prototype["SetSchemeName"]       = ApiThemeFormatScheme.prototype.SetSchemeName;
    ApiThemeFormatScheme.prototype["ChangeFillStyles"]    = ApiThemeFormatScheme.prototype.ChangeFillStyles;
    ApiThemeFormatScheme.prototype["ChangeBgFillStyles"]  = ApiThemeFormatScheme.prototype.ChangeBgFillStyles;
    ApiThemeFormatScheme.prototype["ChangeLineStyles"]    = ApiThemeFormatScheme.prototype.ChangeLineStyles;
    ApiThemeFormatScheme.prototype["Copy"]                = ApiThemeFormatScheme.prototype.Copy;

    ApiThemeFontScheme.prototype["GetClassType"]          = ApiThemeFontScheme.prototype.GetClassType;
    ApiThemeFontScheme.prototype["SetSchemeName"]         = ApiThemeFontScheme.prototype.SetSchemeName;
    ApiThemeFontScheme.prototype["SetFonts"]              = ApiThemeFontScheme.prototype.SetFonts;
    ApiThemeFontScheme.prototype["Copy"]                  = ApiThemeFontScheme.prototype.Copy;


    ApiSlide.prototype["GetClassType"]                    = ApiSlide.prototype.GetClassType;
    ApiSlide.prototype["RemoveAllObjects"]                = ApiSlide.prototype.RemoveAllObjects;
    ApiSlide.prototype["AddObject"]                       = ApiSlide.prototype.AddObject;
    ApiSlide.prototype["SetBackground"]                   = ApiSlide.prototype.SetBackground;
    ApiSlide.prototype["GetWidth"]                        = ApiSlide.prototype.GetWidth;
    ApiSlide.prototype["GetHeight"]                       = ApiSlide.prototype.GetHeight;
    ApiSlide.prototype["ApplyLayout"]                     = ApiSlide.prototype.ApplyLayout;
    ApiSlide.prototype["Delete"]                          = ApiSlide.prototype.Delete;
    ApiSlide.prototype["Copy"]                            = ApiSlide.prototype.Copy;
    ApiSlide.prototype["Duplicate"]                       = ApiSlide.prototype.Duplicate;
    ApiSlide.prototype["MoveTo"]                          = ApiSlide.prototype.MoveTo;
    ApiSlide.prototype["GetSlideIndex"]                   = ApiSlide.prototype.GetSlideIndex;
    ApiSlide.prototype["ClearBackground"]                 = ApiSlide.prototype.ClearBackground;
    ApiSlide.prototype["FollowLayoutBackground"]          = ApiSlide.prototype.FollowLayoutBackground;
    ApiSlide.prototype["FollowMasterBackground"]          = ApiSlide.prototype.FollowMasterBackground;
    ApiSlide.prototype["ApplyTheme"]                      = ApiSlide.prototype.ApplyTheme;
    ApiSlide.prototype["GetLayout"]                       = ApiSlide.prototype.GetLayout;
    ApiSlide.prototype["GetTheme"]                        = ApiSlide.prototype.GetTheme;
    ApiSlide.prototype["GetAllDrawings"]                  = ApiSlide.prototype.GetAllDrawings;
    ApiSlide.prototype["GetAllShapes"]                    = ApiSlide.prototype.GetAllShapes;
    ApiSlide.prototype["GetAllImages"]                    = ApiSlide.prototype.GetAllImages;
    ApiSlide.prototype["GetAllCharts"]                    = ApiSlide.prototype.GetAllCharts;

    ApiDrawing.prototype["GetClassType"]                  = ApiDrawing.prototype.GetClassType;
    ApiDrawing.prototype["SetSize"]                       = ApiDrawing.prototype.SetSize;
    ApiDrawing.prototype["SetPosition"]                   = ApiDrawing.prototype.SetPosition;
    ApiDrawing.prototype["GetParent"]                     = ApiDrawing.prototype.GetParent;
    ApiDrawing.prototype["GetParentSlide"]                = ApiDrawing.prototype.GetParentSlide;
    ApiDrawing.prototype["GetParentLayout"]               = ApiDrawing.prototype.GetParentLayout;
    ApiDrawing.prototype["GetParentMaster"]               = ApiDrawing.prototype.GetParentMaster;
    ApiDrawing.prototype["Copy"]                          = ApiDrawing.prototype.Copy;
    ApiDrawing.prototype["Delete"]                        = ApiDrawing.prototype.Delete;
    ApiDrawing.prototype["SetPlaceholder"]                = ApiDrawing.prototype.SetPlaceholder;
    ApiDrawing.prototype["GetPlaceholder"]                = ApiDrawing.prototype.GetPlaceholder;

    ApiImage.prototype["GetClassType"]                    = ApiImage.prototype.GetClassType;

    ApiShape.prototype["GetClassType"]                    = ApiShape.prototype.GetClassType;
    ApiShape.prototype["GetDocContent"]                   = ApiShape.prototype.GetDocContent;
    ApiShape.prototype["GetContent"]                      = ApiShape.prototype.GetContent;
    ApiShape.prototype["SetVerticalTextAlign"]            = ApiShape.prototype.SetVerticalTextAlign;
    
    ApiChart.prototype["GetClassType"]                    = ApiChart.prototype.GetClassType;
    ApiChart.prototype["SetTitle"]                        = ApiChart.prototype.SetTitle;
    ApiChart.prototype["SetHorAxisTitle"]                 = ApiChart.prototype.SetHorAxisTitle;
    ApiChart.prototype["SetVerAxisTitle"]                 = ApiChart.prototype.SetVerAxisTitle;
    ApiChart.prototype["SetVerAxisOrientation"]           = ApiChart.prototype.SetVerAxisOrientation;
    ApiChart.prototype["SetHorAxisOrientation"]           = ApiChart.prototype.SetHorAxisOrientation;
    ApiChart.prototype["SetLegendPos"]                    = ApiChart.prototype.SetLegendPos;
    ApiChart.prototype["SetLegendFontSize"]               = ApiChart.prototype.SetLegendFontSize;
    ApiChart.prototype["SetShowDataLabels"]               = ApiChart.prototype.SetShowDataLabels;
    ApiChart.prototype["SetShowPointDataLabel"]           = ApiChart.prototype.SetShowPointDataLabel;
    ApiChart.prototype["SetVertAxisTickLabelPosition"]    = ApiChart.prototype.SetVertAxisTickLabelPosition;
    ApiChart.prototype["SetHorAxisTickLabelPosition"]     = ApiChart.prototype.SetHorAxisTickLabelPosition;

    ApiChart.prototype["SetHorAxisMajorTickMark"]         = ApiChart.prototype.SetHorAxisMajorTickMark;
    ApiChart.prototype["SetHorAxisMinorTickMark"]         = ApiChart.prototype.SetHorAxisMinorTickMark;
    ApiChart.prototype["SetVertAxisMajorTickMark"]        = ApiChart.prototype.SetVertAxisMajorTickMark;
    ApiChart.prototype["SetVertAxisMinorTickMark"]        = ApiChart.prototype.SetVertAxisMinorTickMark;
    ApiChart.prototype["SetMajorVerticalGridlines"]       = ApiChart.prototype.SetMajorVerticalGridlines;
    ApiChart.prototype["SetMinorVerticalGridlines"]       = ApiChart.prototype.SetMinorVerticalGridlines;
    ApiChart.prototype["SetMajorHorizontalGridlines"]     = ApiChart.prototype.SetMajorHorizontalGridlines;
    ApiChart.prototype["SetMinorHorizontalGridlines"]     = ApiChart.prototype.SetMinorHorizontalGridlines;
    ApiChart.prototype["SetHorAxisLablesFontSize"]        = ApiChart.prototype.SetHorAxisLablesFontSize;
    ApiChart.prototype["SetVertAxisLablesFontSize"]       = ApiChart.prototype.SetVertAxisLablesFontSize;

    ApiTable.prototype["GetClassType"]                    = ApiTable.prototype.GetClassType;
    ApiTable.prototype["GetRow"]                          = ApiTable.prototype.GetRow;
    ApiTable.prototype["MergeCells"]                      = ApiTable.prototype.MergeCells;
    ApiTable.prototype["SetTableLook"]                    = ApiTable.prototype.SetTableLook;
    ApiTable.prototype["AddRow"]                          = ApiTable.prototype.AddRow;
    ApiTable.prototype["AddColumn"]                       = ApiTable.prototype.AddColumn;
    ApiTable.prototype["RemoveRow"]                       = ApiTable.prototype.RemoveRow;
    ApiTable.prototype["RemoveColumn"]                    = ApiTable.prototype.RemoveColumn;
    ApiTable.prototype["SetShd"]                          = ApiTable.prototype.SetShd;

    ApiTableRow.prototype["GetClassType"]                 = ApiTableRow.prototype.GetClassType;
    ApiTableRow.prototype["GetCellsCount"]                = ApiTableRow.prototype.GetCellsCount;
    ApiTableRow.prototype["GetCell"]                      = ApiTableRow.prototype.GetCell;
    ApiTableRow.prototype["SetHeight"]                    = ApiTableRow.prototype.SetHeight;

    ApiTableCell.prototype["GetClassType"]                = ApiTableRow.prototype.GetClassType;
    ApiTableCell.prototype["GetContent"]                  = ApiTableRow.prototype.GetContent;
    ApiTableCell.prototype["SetShd"]                      = ApiTableRow.prototype.SetShd;
    ApiTableCell.prototype["SetCellMarginBottom"]         = ApiTableRow.prototype.SetCellMarginBottom;
    ApiTableCell.prototype["SetCellMarginLeft"]           = ApiTableRow.prototype.SetCellMarginLeft;
    ApiTableCell.prototype["SetCellMarginRight"]          = ApiTableRow.prototype.SetCellMarginRight;
    ApiTableCell.prototype["SetCellMarginTop"]            = ApiTableRow.prototype.SetCellMarginTop;
    ApiTableCell.prototype["SetCellBorderBottom"]         = ApiTableRow.prototype.SetCellBorderBottom;
    ApiTableCell.prototype["SetCellBorderLeft"]           = ApiTableRow.prototype.SetCellBorderLeft;
    ApiTableCell.prototype["SetCellBorderRight"]          = ApiTableRow.prototype.SetCellBorderRight;
    ApiTableCell.prototype["SetCellBorderTop"]            = ApiTableRow.prototype.SetCellBorderTop;
    ApiTableCell.prototype["SetVerticalAlign"]            = ApiTableRow.prototype.SetVerticalAlign;
    ApiTableCell.prototype["SetTextDirection"]            = ApiTableRow.prototype.SetTextDirection;



    function private_GetCurrentSlide(){
        var oApiPresentation = editor.GetPresentation();
        if(oApiPresentation){
            var oApiSlide = oApiPresentation.GetCurrentSlide();
            if(oApiSlide){
                return oApiSlide.Slide;
            }
        }
        return null;
    }

    function private_GetDrawingDocument(){
        if(editor && editor.WordControl){
            return editor.WordControl.m_oDrawingDocument;
        }
        if(Asc["editor"] && Asc["editor"].wbModel) {
            return Asc["editor"].wbModel.DrawingDocument;
        }
        return null;
    }

    function private_GetPresentation(){
        return editor.WordControl.m_oLogicDocument;
    }

    function private_EMU2MM(EMU)
    {
        return EMU / 36000.0;
    }

    function private_GetApi(){
        return editor;
    }


    function private_GetBoolean(bValue, bDefValue)
    {
        if (true === bValue)
            return true;
        else if (false === bValue)
            return false;
        else
            return (undefined !== bDefValue ? bDefValue : false);
    }
    function private_Twips2MM(twips)
    {
        return 25.4 / 72.0 / 20 * twips;
    }
    function private_GetInt(nValue, nMin, nMax)
    {
        var nResult = nValue | 0;

        if (undefined !== nMin && null !== nMin)
            nResult = Math.max(nMin, nResult);

        if (undefined !== nMax && null !== nMax)
            nResult = Math.min(nMax, nResult);

        return nResult;
    }
    function private_GetTableMeasure(sType, nValue)
    {
        var nType = tblwidth_Auto;
        var nW    = 0;
        if ("auto" === sType)
        {
            nType = tblwidth_Auto;
            nW    = 0;
        }
        else if ("nil" === sType)
        {
            nType = tblwidth_Nil;
            nW    = 0;
        }
        else if ("percent" === sType)
        {
            nType = tblwidth_Pct;
            nW    = private_GetInt(nValue, null, null);
        }
        else if ("twips" === sType)
        {
            nType = tblwidth_Mm;
            nW    = private_Twips2MM(nValue);
        }

        return new CTableMeasurement(nType, nW);
    }
})(window, null);

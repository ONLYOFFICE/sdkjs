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
(function(window, undefined) {

      // Import
      var CellValueType = AscCommon.CellValueType;
      var c_oAscCellAnchorType = AscCommon.c_oAscCellAnchorType;
      var c_oAscBorderStyles = Asc.c_oAscBorderStyles;
      var gc_nMaxRow0 = AscCommon.gc_nMaxRow0;
      var gc_nMaxCol0 = AscCommon.gc_nMaxCol0;
      var Binary_CommonReader = AscCommon.Binary_CommonReader;
      var BinaryCommonWriter = AscCommon.BinaryCommonWriter;
      var c_oSerPropLenType = AscCommon.c_oSerPropLenType;
      var c_oSerConstants = AscCommon.c_oSerConstants;
    var History = AscCommon.History;
    var pptx_content_loader = AscCommon.pptx_content_loader;
    var pptx_content_writer = AscCommon.pptx_content_writer;

      var c_oAscPageOrientation = Asc.c_oAscPageOrientation;

    var g_oDefaultFormat = AscCommonExcel.g_oDefaultFormat;
	var g_StyleCache = AscCommonExcel.g_StyleCache;
    var g_cSharedWriteStreak = 64;//like Excel

    function decodeXmlPath(str, opt_prepare_external_data) {
        var res = str;

        if (opt_prepare_external_data) {
            res = res.replaceAll("\\'", "\"");
        }

        res = res.replace(/&apos;/g,"'");
        res = res.replace(/&amp;/g,'&');
        res = res.replace(/%20/g,' ');
        res = res.replace(/%23/g,'#');
        res = res.replace(/%25/g,'%');
        res = res.replace(/%5e/g,'^');

        return res;
    }

    function encodeXmlPath(str, opt_amp_encode, opt_prepare_external_data) {
        var res = str;
        if (opt_prepare_external_data) {
            res = res.replaceAll('\"',"\\'");
        }
        if (opt_amp_encode) {
            res = res.replace(/&/g,'&amp;');
        }
        res = res.replace(/%/g,'%25');
        res = res.replace(/ /g,'%20');
        res = res.replace(/#/g,'%23');
        res = res.replace(/\^/g,'%5e');

        return res;
    }

    function completePathForLocalLinks(str) {
        // if the result contains a path relative to the current file, then we add data from the local path
        // BookLink.xlsx => C:\Users\FileFolder\[BookLink.xlsx] - same folder
        // test/BookLink.xlsx => C:\Users\FileFolder\test\[BookLink.xlsx] - deep folder
        // Users/FileFolder/BookLink.xlsx => C:\Users\FileFolder\[BookLink.xlsx] - up folder
        // file:///D:\AnotherDiskFolder\BookLink.xlsx => D:\AnotherDiskFolder\[BookLink.xlsx] - file on another disk

        let res = str;
        if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]()) {
            /* replace file:// */
            res = res.replace(/^file:\/\/\//, '');
            res = res.replace(/^file:\/\//, '');

            const currentFilePath = window["AscDesktopEditor"]["LocalFileGetSourcePath"]();
            let currentPathParts = currentFilePath && currentFilePath.split(/[\\/]/).slice(0, -1);    // remove file name

            let receivedPathParts = res.split(/[\\/]/);

            let diskRegex = /^[a-zA-Z]:/;
            let isLinkHasDiskLetter = diskRegex.test(receivedPathParts[0]);

            if (currentPathParts && !isLinkHasDiskLetter && res.indexOf(currentPathParts[0]) === -1) {
                // incomplete link string, check the path

                // TODO all links should contain BookLink, not the full path
                if (res[0] === "/" || res[0] === "//") {
                    // link to other file up folder
                    // add only disk name to path
                    res = currentPathParts[0] + receivedPathParts.join("\\");
                } else {
                    // link to other file deep in the current folder
                    // add current folder to path
                    for (let i = 0; i < receivedPathParts.length; i++) {
                        let part = receivedPathParts[i];
                        currentPathParts.push(part);
                    }
    
                    res = currentPathParts.join("\\");
                }
            } 
        }
        return res;
    }

	var XLSB = {
		rt_ROW_HDR : 0,
		rt_CELL_BLANK : 1,
		rt_CELL_RK : 2,
		rt_CELL_ERROR : 3,
		rt_CELL_BOOL : 4,
		rt_CELL_REAL : 5,
		rt_CELL_ST : 6,
		rt_CELL_ISST : 7,
		rt_FMLA_STRING : 8,
		rt_FMLA_NUM : 9,
		rt_FMLA_BOOL: 10,
		rt_FMLA_ERROR: 11,
		rt_BEGIN_SHEET_DATA: 145,
		rt_END_SHEET_DATA: 146
	};
//dif:
//Version:2 добавлены свойства колонок и строк CustomWidth, CustomHeight(раньше считались true)
    /** @enum */
    var c_oSerTableTypes =
    {
        Other: 0,
        SharedStrings: 1,
        Styles: 2,
        Workbook: 3,
        Worksheets: 4,
        CalcChain: 5,
        App: 6,
        Core: 7,
        PersonList: 8,
        CustomProperties: 9,
        Customs: 10
    };
    /** @enum */
    var c_oSerStylesTypes =
    {
        Borders: 0,
        Border: 1,
        CellXfs: 2,
        Xfs: 3,
        Fills: 4,
        Fill: 5,
        Fonts: 6,
        Font: 7,
        NumFmts: 8,
        NumFmt: 9,
        Dxfs: 10,
        Dxf: 11,
        TableStyles: 12,
        CellStyleXfs: 14,
        CellStyles: 15,
        CellStyle: 16,
        SlicerStyles: 17,
        ExtDxfs: 18,
        TimelineStyles: 19
    };
    /** @enum */
    var c_oSerBorderTypes =
    {
        Bottom: 0,
        Diagonal: 1,
        End: 2,
        Horizontal: 3,
        Start: 4,
        Top: 5,
        Vertical: 6,
        DiagonalDown: 7,
        DiagonalUp: 8,
        Outline: 9
    };
    /** @enum */
    var c_oSerBorderPropTypes =
    {
        Color: 0,
        Style: 1
    };
    /** @enum */
    var c_oSerXfsTypes =
    {
        ApplyAlignment: 0,
        ApplyBorder: 1,
        ApplyFill: 2,
        ApplyFont: 3,
        ApplyNumberFormat: 4,
        ApplyProtection: 5,
        BorderId: 6,
        FillId: 7,
        FontId: 8,
        NumFmtId: 9,
        PivotButton: 10,
        QuotePrefix: 11,
        XfId: 12,
        Aligment: 13,
        Protection: 14
    };
	var c_oSerProtectionTypes =
    {
        Hidden: 0,
        Locked: 1
    };
    /** @enum */
    var c_oSerFillTypes =
    {
        Pattern: 0,
        PatternBgColor_deprecated: 1,
        PatternType : 2,
        PatternFgColor : 3,
        PatternBgColor : 4,
        Gradient : 5,
        GradientType : 6,
        GradientLeft : 7,
        GradientTop : 8,
        GradientRight : 9,
        GradientBottom : 10,
        GradientDegree : 11,
        GradientStop : 12,
        GradientStopPosition : 13,
        GradientStopColor : 14
    };
    /** @enum */
    var c_oSerFontTypes =
    {
        Bold: 0,
        Color: 1,
        Italic: 3,
        RFont: 4,
        Strike: 5,
        Sz: 6,
        Underline: 7,
        VertAlign: 8,
        Scheme: 9
    };
    /** @enum */
    var c_oSerNumFmtTypes =
    {
        FormatCode: 0,
        NumFmtId: 1
    };
    /** @enum */
    var c_oSerSharedStringTypes =
    {
        Si: 0,
        Run: 1,
        RPr: 2,
        Text: 3
    };
    /** @enum */
    var c_oSerWorkbookTypes = {
        WorkbookPr: 0,
        BookViews: 1,
        WorkbookView: 2,
        DefinedNames: 3,
        DefinedName: 4,
        ExternalReferences: 5,
        ExternalReference: 6,
        PivotCaches: 7,
        PivotCache: 8,
        ExternalBook: 9,
        OleLink: 10,
        DdeLink: 11,
        VbaProject: 12,
        JsaProject: 13,
        Comments: 14,
        CalcPr: 15,
        Connections: 16,
        AppName: 17,
        SlicerCaches: 18,
        SlicerCachesExt: 19,
        SlicerCache: 20,
        WorkbookProtection: 21,
        OleSize: 22,
        ExternalFileId: 23,
        ExternalPortalName: 24,
        FileSharing: 25,
        ExternalLinksAutoRefresh: 26,
        TimelineCaches: 27,
        TimelineCache: 28,
        Metadata: 29,
        XmlMap: 30
    };
    /** @enum */
    var c_oSerWorkbookPrTypes =
    {
        Date1904: 0,
        DateCompatibility: 1,
		HidePivotFieldList: 2,
		ShowPivotChartFilter: 3,
        UpdateLinks: 4
    };
    /** @enum */
    var c_oSerWorkbookViewTypes =
    {
        ActiveTab: 0,
        AutoFilterDateGrouping: 1,
        FirstSheet: 2,
        Minimized: 3,
        ShowHorizontalScroll: 4,
        ShowSheetTabs: 5,
        ShowVerticalScroll: 6,
        TabRatio: 7,
        Visibility: 8,
        WindowHeight: 9,
        WindowWidth: 10,
        XWindow: 11,
        YWindow: 12
    };
    /** @enum */
    var c_oSerDefinedNameTypes =
    {
        Name: 0,
        Ref: 1,
        LocalSheetId: 2,
        Hidden: 3
    };
	/** @enum */
	var c_oSerCalcPrTypes = {
		CalcId: 0,
		CalcMode: 1,
		FullCalcOnLoad: 2,
		RefMode: 3,
		Iterate: 4,
		IterateCount: 5,
		IterateDelta: 6,
		FullPrecision: 7,
		CalcCompleted: 8,
		CalcOnSave: 9,
		ConcurrentCalc: 10,
		ConcurrentManualCount: 11,
		ForceFullCalc: 12
	};
	/** @enum */
    var c_oSerWorksheetsTypes =
    {
        Worksheet: 0,
        WorksheetProp: 1,
        Cols: 2,
        Col: 3,
        Dimension: 4,
        Hyperlinks: 5,
        Hyperlink: 6,
        MergeCells: 7,
        MergeCell: 8,
        SheetData: 9,
        Row: 10,
        SheetFormatPr: 11,
        Drawings: 12,
        Drawing: 13,
        PageMargins: 14,
        PageSetup: 15,
        PrintOptions: 16,
        Autofilter: 17,
        TableParts: 18,
        Comments: 19,
        Comment: 20,
        ConditionalFormatting: 21,
        SheetViews: 22,
        SheetView: 23,
        SheetPr: 24,
        SparklineGroups: 25,
		PivotTable: 26,
        HeaderFooter: 27,
        LegacyDrawingHF: 28,
        Picture: 29,
        RowBreaks: 30,
		ColBreaks: 31,
		DataValidations: 32,
		QueryTable: 33,
		Controls: 34,
		XlsbPos: 35,
		SortState: 36,
        Slicers: 37,
        SlicersExt: 38,
        Slicer: 39,
        NamedSheetView: 40,
        ProtectionSheet: 41,
        ProtectedRanges: 42,
        ProtectedRange: 43,
        CellWatches: 44,
        CellWatch: 45,
        CellWatchR: 46,
        UserProtectedRanges: 47,
        TimelinesList: 48,
        Timelines: 49,
        Timeline: 50,
        TableSingleCells: 51
    };
    /** @enum */
    var c_oSerWorksheetPropTypes =
    {
        Name: 0,
        SheetId: 1,
        State: 2,
        Ref: 3
    };
    /** @enum */
    var c_oSerWorksheetColTypes =
    {
        BestFit: 0,
        Hidden: 1,
        Max: 2,
        Min: 3,
        Style: 4,
        Width: 5,
        CustomWidth: 6,
        OutLevel: 7,
        Collapsed: 8
    };
    /** @enum */
    var c_oSerHyperlinkTypes =
    {
        Ref: 0,
        Hyperlink: 1,
        Location: 2,
        Tooltip: 3
    };
    /** @enum */
    var c_oSerSheetFormatPrTypes =
    {
        DefaultColWidth		: 0,
        DefaultRowHeight	: 1,
        BaseColWidth		: 2,
        CustomHeight		: 3,
        ZeroHeight			: 4,
        OutlineLevelCol		: 5,
        OutlineLevelRow		: 6
    };
    /** @enum */
    var c_oSerRowTypes =
    {
        Row: 0,
        Style: 1,
        Height: 2,
        Hidden: 3,
        Cells: 4,
        Cell: 5,
        CustomHeight: 6,
        OutLevel: 7,
        Collapsed: 8
    };
    /** @enum */
    var c_oSerCellTypes =
    {
        Ref: 0,
        Style: 1,
        Type: 2,
        Value: 3,
        Formula: 4,
        RefRowCol: 5,
        ValueText: 6,
        ValueCache: 7,
        CellMetadata: 8,
        ValueMetadata: 9
    };
    /** @enum */
    var c_oSerFormulaTypes =
    {
        Aca: 0,
        Bx: 1,
        Ca: 2,
        Del1: 3,
        Del2: 4,
        Dt2D: 5,
        Dtr: 6,
        R1: 7,
        R2: 8,
        Ref: 9,
        Si: 10,
        T: 11,
        Text: 12
    };
    /** @enum */
    var c_oSer_DrawingFromToType =
    {
        Col: 0,
        ColOff: 1,
        Row: 2,
        RowOff: 3
    };
    /** @enum */
    var c_oSer_DrawingPosType =
    {
        X: 0,
        Y: 1
    };
    /** @enum */
    var c_oSer_DrawingExtType =
    {
        Cx: 0,
        Cy: 1
    };
    /** @enum */
    var c_oSer_OtherType =
    {
        Media: 0,
        MediaItem: 1,
        MediaId: 2,
        MediaSrc: 3,
        Theme: 5
    };
    /** @enum */
    var c_oSer_CalcChainType =
    {
        CalcChainItem: 0,
        Array: 1,
        SheetId: 2,
        DependencyLevel: 3,
        Ref: 4,
        ChildChain: 5,
        NewThread: 6
    };
    /** @enum */
    var  c_oSer_PageMargins =
    {
        Left: 0,
        Top: 1,
        Right: 2,
        Bottom: 3,
        Header: 4,
        Footer: 5
    };
    /** @enum */
    var  c_oSer_PageSetup =
    {
        Orientation: 0,
        PaperSize: 1,
        BlackAndWhite: 2,
        CellComments: 3,
        Copies: 4,
        Draft: 5,
        Errors: 6,
        FirstPageNumber: 7,
        FitToHeight: 8,
        FitToWidth: 9,
        HorizontalDpi: 10,
        PageOrder: 11,
        PaperHeight: 12,
        PaperWidth: 13,
        PaperUnits: 14,
        Scale: 15,
        UseFirstPageNumber: 16,
        UsePrinterDefaults: 17,
        VerticalDpi: 18
    };
    /** @enum */
    var  c_oSer_PrintOptions =
    {
        GridLines: 0,
        Headings: 1,
        GridLinesSet: 2,
        HorizontalCentered: 3,
        VerticalCentered: 4
    };
    /** @enum */
    var c_oSer_TablePart =
    {
        Table:0,
        Ref:1,
        TotalsRowCount:2,
        DisplayName:3,
        AutoFilter:4,
        SortState:5,
        TableColumns:6,
        TableStyleInfo:7,
		HeaderRowCount:8,
		AltTextTable: 9,
        Name: 10,
        Comment: 11,
        ConnectionId: 12,
        Id: 13,
        DataCellStyle: 14,
        DataDxfId: 15,
        HeaderRowBorderDxfId: 16,
        HeaderRowCellStyle: 17,
        HeaderRowDxfId: 18,
        InsertRow: 19,
        InsertRowShift: 20,
        Published: 21,
        TableBorderDxfId: 22,
        TableType: 23,
        TotalsRowBorderDxfId: 24,
        TotalsRowCellStyle: 25,
        TotalsRowDxfId: 26,
        TotalsRowShown: 27,
        QueryTable: 28
    };
    /** @enum */
    var c_oSer_TableStyleInfo =
    {
        Name:0,
        ShowColumnStripes:1,
        ShowRowStripes:2,
        ShowFirstColumn:3,
        ShowLastColumn:4
    };
    
    /** @enum */
    var c_oSer_TableColumns =
    {
        TableColumn:0,
        Name:1,
        DataDxfId:2,
        TotalsRowLabel:3,
        TotalsRowFunction:4,
        TotalsRowFormula:5,
        CalculatedColumnFormula:6,
        DataCellStyle: 7,
        HeaderRowCellStyle: 8,
        HeaderRowDxfId: 9,
        Id: 10,
        QueryTableFieldId: 11,
        TotalsRowCellStyle: 12,
        TotalsRowDxfId: 13,
        UniqueName: 14,
        XmlColumnPr: 15,
        MapId: 16,
        Xpath: 17,
        Denormalized: 18,
        XmlDataType: 19
    };
    /** @enum */
    var c_oSer_SortState =
    {
        Ref:0,
        CaseSensitive:1,
        SortConditions:2,
        SortCondition:3,
        ConditionRef:4,
        ConditionSortBy:5,
        ConditionDescending:6,
        ConditionDxfId:7,
        ColumnSort: 8,
        SortMethod: 9
    };
    /** @enum */
    var c_oSer_AutoFilter =
    {
        Ref:0,
        FilterColumns:1,
        FilterColumn:2,
        SortState:3
    };
    /** @enum */
    var c_oSer_FilterColumn =
    {
        ColId:0,
        Filters:1,
        Filter:2,
        DateGroupItem:3,
        CustomFilters:4,
        ColorFilter:5,
        Top10:6,
        DynamicFilter: 7,
        HiddenButton: 8,
        ShowButton: 9,
        FiltersBlank: 10
    };
    /** @enum */
    var c_oSer_Filter =
    {
        Val:0
    };
    /** @enum */
    var c_oSer_DateGroupItem =
    {
        DateTimeGrouping:0,
        Day:1,
        Hour:2,
        Minute:3,
        Month:4,
        Second:5,
        Year:6
    };
    /** @enum */
    var c_oSer_CustomFilters =
    {
        And:0,
        CustomFilters:1,
        CustomFilter:2,
        Operator:3,
        Val:4
    };
    /** @enum */
    var c_oSer_DynamicFilter =
    {
        Type: 0,
        Val: 1,
        MaxVal: 2
    };
    /** @enum */
    var c_oSer_ColorFilter =
    {
        CellColor:0,
        DxfId:1
    };
    /** @enum */
    var c_oSer_Top10 =
    {
        FilterVal:0,
        Percent:1,
        Top:2,
        Val:3
    };
	var c_oSer_QueryTable =
    {
		AdjustColumnWidth: 0,
		ApplyAlignmentFormats: 1,
		ApplyBorderFormats: 2,
		ApplyFontFormats: 3,
		ApplyNumberFormats: 4,
		ApplyPatternFormats: 5,
		ApplyWidthHeightFormats: 6,
		BackgroundRefresh: 7,
		AutoFormatId: 8,
		ConnectionId: 9,
		DisableEdit: 10,
		DisableRefresh: 11,
		FillFormulas: 12,
		FirstBackgroundRefresh: 13,
		GrowShrinkType: 14,
		Headers: 15,
		Intermediate: 16,
		Name: 17,
		PreserveFormatting: 18,
		RefreshOnLoad: 19,
		RemoveDataOnSave: 20,
		RowNumbers: 21,
		QueryTableRefresh: 22
    };
	var c_oSer_QueryTableRefresh =
    {
        NextId: 0,
        MinimumVersion: 1,
        FieldIdWrapped: 2,
        HeadersInLastRefresh: 3,
        PreserveSortFilterLayout: 4,
        UnboundColumnsLeft: 5,
        UnboundColumnsRight: 6,
        QueryTableFields: 7,
        QueryTableDeletedFields: 8,
        SortState: 9
    };

	var c_oSer_QueryTableField =
    {
        QueryTableField: 0,
        Id: 1,
        TableColumnId: 2,
        Name: 3,
        RowNumbers: 4,
        FillFormulas: 5,
        DataBound: 6,
        Clipped: 7
    };

	var c_oSer_QueryTableDeletedField =
    {
        QueryTableDeletedField: 0,
        Name: 1
    };

    /** @enum */
    var c_oSer_Dxf =
    {
        Alignment:0,
        Border:1,
        Fill:2,
        Font:3,
        NumFmt:4
    };
    /** @enum */
    var c_oSer_TableStyles = {
        DefaultTableStyle:0,
        DefaultPivotStyle:1,
        TableStyles: 2,
        TableStyle: 3
    };
    var c_oSer_TableStyle = {
        Name: 0,
        Pivot: 1,
        Table: 2,
        Elements: 3,
        Element: 4,
        DisplayName: 5
    };
    var c_oSer_TableStyleElement = {
        DxfId: 0,
        Size: 1,
        Type: 2
    };
    var c_oSer_Comments =
    {
        Row: 0,
        Col: 1,
        CommentDatas : 2,
        CommentData : 3,
        Left: 4,
        LeftOffset: 5,
        Top: 6,
        TopOffset: 7,
        Right: 8,
        RightOffset: 9,
        Bottom: 10,
        BottomOffset: 11,
        LeftMM: 12,
        TopMM: 13,
        WidthMM: 14,
        HeightMM: 15,
        MoveWithCells: 16,
        SizeWithCells: 17,
        ThreadedComment: 18
    };
    var c_oSer_CommentData =
    {
        Text : 0,
        Time : 1,
        UserId : 2,
        UserName : 3,
        QuoteText : 4,
        Solved : 5,
        Document : 6,
        Replies : 7,
        Reply : 8,
        OOTime : 9,
        Guid : 10,
        UserData : 11
    };
    var c_oSer_ThreadedComment =
    {
        dT: 0,
        personId: 1,
        id: 2,
        done: 3,
        text: 4,
        mention: 5,
        reply: 6,
        mentionpersonId: 7,
        mentionId: 8,
        startIndex: 9,
        length: 10
    };
    var c_oSer_Person =
    {
        person: 0,
        id: 1,
        providerId: 2,
        userId: 3,
        displayName: 4
    };
    var c_oSer_ConditionalFormatting = {
        Pivot						: 0,
        SqRef						: 1,
        ConditionalFormattingRule	: 2
    };
    var c_oSer_ConditionalFormattingRule = {
        AboveAverage	: 0,
        Bottom			: 1,
        DxfId			: 2,
        EqualAverage	: 3,
        Operator		: 4,
        Percent			: 5,
        Priority		: 6,
        Rank			: 7,
        StdDev			: 8,
        StopIfTrue		: 9,
        Text			: 10,
        TimePeriod		: 11,
        Type			: 12,
        ColorScale		: 14,
        DataBar			: 15,
        FormulaCF		: 16,
		IconSet			: 17,
		Dxf				: 18,
		isExt			: 19
    };
    var c_oSer_ConditionalFormattingRuleColorScale = {
        CFVO			: 0,
        Color			: 1
    };
    var c_oSer_ConditionalFormattingDataBar = {
        CFVO			: 0,
        Color			: 1,
        MaxLength		: 2,
        MinLength		: 3,
		ShowValue		: 4,
		NegativeColor	: 5,
		BorderColor		: 6,
		AxisColor		: 7,
		NegativeBorderColor: 8,
		AxisPosition	: 9,
		Direction		: 10,
		GradientEnabled	: 11,
		NegativeBarColorSameAsPositive: 12,
		NegativeBarBorderColorSameAsPositive: 13
    };
    var c_oSer_ConditionalFormattingIconSet = {
        CFVO			: 0,
        IconSet			: 1,
        Percent			: 2,
        Reverse			: 3,
		ShowValue		: 4,
		CFIcon			: 5
    };
    var c_oSer_ConditionalFormattingValueObject = {
        Gte				: 0,
        Type			: 1,
		Val				: 2,
		Formula			: 3
    };
	var c_oSer_ConditionalFormattingIcon = {
		iconSet : 0,
		iconId : 1
	};
	var c_oSer_DataValidation = {
		DataValidations: 0,
		DataValidation: 1,
		DisablePrompts: 2,
		XWindow: 3,
		YWindow: 4,
		Type: 5,
		AllowBlank: 6,
		Error: 7,
		ErrorTitle: 8,
		ErrorStyle: 9,
		ImeMode: 10,
		Operator: 11,
		Prompt: 12,
		PromptTitle: 13,
		ShowDropDown: 14,
		ShowErrorMessage: 15,
		ShowInputMessage: 16,
		SqRef: 17,
		Formula1: 18,
		Formula2: 19,
		List: 20
	};
    var c_oSer_SheetView = {
        ColorId						: 0,
        DefaultGridColor			: 1,
        RightToLeft					: 2,
        ShowFormulas				: 3,
        ShowGridLines				: 4,
        ShowOutlineSymbols			: 5,
        ShowRowColHeaders			: 6,
        ShowRuler					: 7,
        ShowWhiteSpace				: 8,
        ShowZeros					: 9,
        TabSelected					: 10,
        TopLeftCell					: 11,
        View						: 12,
        WindowProtection			: 13,
        WorkbookViewId				: 14,
        ZoomScale					: 15,
        ZoomScaleNormal				: 16,
        ZoomScalePageLayoutView		: 17,
        ZoomScaleSheetLayoutView	: 18,

		Pane						: 19,
		Selection					: 20
    };
    var c_oSer_DrawingType =
    {
        Type: 0,
        From: 1,
        To: 2,
        Pos: 3,
        Pic: 4,
        PicSrc: 5,
        GraphicFrame: 6,
        Chart: 7,
        Ext: 8,
        pptxDrawing: 9,
        Chart2: 10,
        ObjectName: 11,
        EditAs: 12,
        ClientData: 14,
        pptxDrawingAlternative: 0x99
    };

    var c_oSer_DrawingClientDataType =
    {
        fLocksWithSheet: 0,
        fPrintsWithSheet: 1
    };

    /** @enum */
    var c_oSer_Pane = {
        ActivePane	: 0,
		State		: 1,
        TopLeftCell	: 2,
        XSplit		: 3,
		YSplit		: 4
    };
	/** @enum */
	 var c_oSer_Selection = {
		ActiveCell: 0,
		ActiveCellId: 1,
		Sqref: 2,
		Pane: 3
	};
    /** @enum */
    var c_oSer_CellStyle = {
        BuiltinId		: 0,
        CustomBuiltin	: 1,
        Hidden			: 2,
        ILevel			: 3,
        Name			: 4,
        XfId			: 5
    };
    /** @enum */
    var c_oSer_SheetPr = {
        CodeName							: 0,
        EnableFormatConditionsCalculation	: 1,
        FilterMode							: 2,
        Published							: 3,
        SyncHorizontal						: 4,
        SyncRef								: 5,
        SyncVertical						: 6,
        TransitionEntry						: 7,
        TransitionEvaluation				: 8,

		TabColor							: 9,
		PageSetUpPr							: 10,
		AutoPageBreaks						: 11,
		FitToPage							: 12,
		OutlinePr							: 13,
		ApplyStyles							: 14,
		ShowOutlineSymbols					: 15,
		SummaryBelow						: 16,
		SummaryRight						: 17
    };
    /** @enum */
    var c_oSer_Sparkline = {
        SparklineGroup: 0,
        ManualMax: 1,
        ManualMin: 2,
        LineWeight: 3,
        Type: 4,
        DateAxis: 5,
        DisplayEmptyCellsAs: 6,
        Markers: 7,
        High: 8,
        Low: 9,
        First: 10,
        Last: 11,
        Negative: 12,
        DisplayXAxis: 13,
        DisplayHidden: 14,
        MinAxisType: 15,
        MaxAxisType: 16,
        RightToLeft: 17,
        ColorSeries: 18,
        ColorNegative: 19,
        ColorAxis: 20,
        ColorMarkers: 21,
        ColorFirst: 22,
        ColorLast: 23,
        ColorHigh: 24,
        ColorLow: 25,
        Ref: 26,
        Sparklines: 27,
        Sparkline: 28,
        SparklineRef: 29,
        SparklineSqRef: 30
    };
	/** @enum */
	var c_oSer_AltTextTable = {
		AltText: 0,
		AltTextSummary: 1
	};
	/** @enum */
	var c_oSer_PivotTypes = {
		id: 0,
		cache: 1,
		record: 2,
		cacheId: 3,
		table: 4
	};
	/** @enum */
	var c_oSer_ExternalLinkTypes = {
		Id: 0,
		SheetNames: 1,
		SheetName: 2,
		DefinedNames: 3,
		DefinedName: 4,
		DefinedNameName: 5,
		DefinedNameRefersTo: 6,
		DefinedNameSheetId: 7,
		SheetDataSet: 8,
		SheetData: 9,
		SheetDataSheetId: 10,
		SheetDataRefreshError: 11,
		SheetDataRow: 12,
		SheetDataRowR: 13,
		SheetDataRowCell: 14,
		SheetDataRowCellRef: 15,
		SheetDataRowCellType: 16,
		SheetDataRowCellValue: 17,
		AlternateUrls : 18,
		AbsoluteUrl : 19,
		RelativeUrl : 20,
		ExternalAlternateUrlsDriveId : 21,
		ExternalAlternateUrlsItemId : 22,
		ValueMetadata : 23
	};
    var c_oSer_HeaderFooter = {
        AlignWithMargins: 0,
        DifferentFirst: 1,
        DifferentOddEven: 2,
        ScaleWithDoc: 3,
        EvenFooter: 4,
        EvenHeader: 5,
        FirstFooter: 6,
        FirstHeader: 7,
        OddFooter: 8,
        OddHeader: 9
    };
    var c_oSer_RowColBreaks = {
        Count: 0,
        ManualBreakCount: 1,
        Break: 2,
        Id: 3,
        Man: 4,
        Max: 5,
        Min: 6,
        Pt: 7
    };
    var c_oSer_LegacyDrawingHF = {
        Drawings: 0,
        Drawing: 1,
        DrawingId: 2,
        DrawingShape: 3,
        Cfe: 4,
        Cff: 5,
        Cfo: 6,
        Che: 7,
        Chf: 8,
        Cho: 9,
        Lfe: 10,
        Lff: 11,
        Lfo: 12,
        Lhe: 13,
        Lhf: 14,
        Lho: 15,
        Rfe: 16,
        Rff: 17,
        Rfo: 18,
        Rhe: 19,
        Rhf: 20,
        Rho: 21
    };
	var c_oSerWorksheetProtection = {
		AlgorithmName: 0,
		SpinCount: 1,
		HashValue: 2,
		SaltValue: 3,
		Password: 4,//TODO ?
		AutoFilter: 5,
		Content: 6,
		DeleteColumns: 7,
		DeleteRows: 8,
		FormatCells: 9,
		FormatColumns: 10,
		FormatRows: 11,
		InsertColumns: 12,
		InsertHyperlinks: 13,
		InsertRows: 14,
		Objects: 15,
		PivotTables: 16,
		Scenarios: 17,
		SelectLockedCells: 18,
		SelectUnlockedCells: 19,
		Sheet: 20,
		Sort: 21
	};
	var c_oSerProtectedRangeTypes = {
		AlgorithmName: 0,
        SpinCount: 1,
        HashValue: 2,
        SaltValue: 3,
        Name: 4,
        SqRef: 5,
        SecurityDescriptor: 6
	};

    var c_oSerWorkbookProtection = {
		WorkbookAlgorithmName: 0,
		WorkbookSpinCount: 1,
		WorkbookHashValue: 2,
		WorkbookSaltValue: 3,
		LockStructure: 4,
		LockWindows: 5,
		Password: 6,
		RevisionsAlgorithmName: 7,
		RevisionsHashValue: 8,
		RevisionsSaltValue: 9,
		RevisionsSpinCount: 10,
		LockRevision: 11
	};
	var c_oSerFileSharing = {
		AlgorithmName: 0,
		SpinCount: 1,
		HashValue: 2,
		SaltValue: 3,
		UserName: 4,
		ReadOnly: 5,
		Password: 6
	};
    var c_oSerCustoms = {
        Custom: 0,
        ItemId: 1,
        Uri: 2,
        Content: 3
    };
    var c_oSerUserProtectedRange = {
        UserProtectedRange: 0,
        Sqref: 1,
        Name: 2,
        Text: 3,
        User: 4,
        UsersGroup: 5,
        Type: 6
    };
	var c_oSerUserProtectedRangeDesc = {
		Id: 0,
		Name: 1,
		Type: 2
	};
	var c_oSerUserProtectedRangeType = {
		notView: 0,
		view: 1,
		edit: 2
	};

    /** @enum */
    var c_oSer_Timeline = {
        Name: 0,
        Caption: 1,
        Uid: 2,
        Cache: 3,
        ShowHeader: 4,
        ShowTimeLevel: 5,
        ShowSelectionLabel: 6,
        ShowHorizontalScrollbar: 7,
        Level: 8,
        SelectionLevel: 9,
        ScrollPosition: 10,
        Style: 11
    };
    var c_oSer_TimelineCache = {
        Name: 0,
        SourceName: 1,
        Uid: 2,
        PivotTables: 3,
        PivotTable: 4,
        State: 5,
        PivotFilter: 6
    };
    var c_oSer_TimelineState = {
        Name: 0,
        FilterState: 1,
        PivotCacheId: 2,
        MinimalRefreshVersion: 3,
        LastRefreshVersion: 4,
        FilterType: 5,
        Selection: 6,
        Bounds: 7
    };
    var c_oSer_TimelinePivotFilter = {

        Name: 0,
        Description: 1,
        UseWholeDay: 2,
        Id: 3,
        Fld: 4,
        AutoFilter: 5
    };
    var c_oSer_TimelineRange = {
        StartDate: 0,
        EndDate: 1
    };
    var c_oSer_TimelineCachePivotTable = {
        Name: 0,
        TabId: 1
    };
    var c_oSer_TimelineStyles = {
        DefaultTimelineStyle : 0,
        TimelineStyle : 2,
        TimelineStyleName : 3,
        TimelineStyleElement : 4,
        TimelineStyleElementType : 5,
        TimelineStyleElementDxfId : 6
    };


	var c_oSer_Metadata =
	{
		MetadataTypes: 0,
		MetadataStrings: 1,
		MdxMetadata: 2,
		CellMetadata: 3,
		ValueMetadata: 4,
		FutureMetadata: 5
	};
	var c_oSer_MetadataType =
	{
		MetadataType: 0,
		Name: 1,
		MinSupportedVersion: 2,
		GhostRow: 3,
		GhostCol: 4,
		Edit: 5,
		Delete: 6,
		Copy: 7,
		PasteAll: 8,
		PasteFormulas: 9,
		PasteValues: 10,
		PasteFormats: 11,
		PasteComments: 12,
		PasteDataValidation: 13,
		PasteBorders: 14,
		PasteColWidths: 15,
		PasteNumberFormats: 16,
		Merge: 17,
		SplitFirst: 18,
		SplitAll: 19,
		RowColShift: 30,
		ClearAll: 21,
		ClearFormats: 22,
		ClearContents: 23,
		ClearComments: 24,
		Assign: 25,
		Coerce: 26,
		CellMeta: 27
	};
	var c_oSer_MetadataString =
	{
		MetadataString: 0

	};
	var c_oSer_MetadataBlock =
	{
		MetadataBlock: 0,
		MetadataRecord: 1,
		MetadataRecordType: 2,
		MetadataRecordValue: 3
	};
	var c_oSer_FutureMetadataBlock =
	{
		Name: 0,
		FutureMetadataBlock: 1,
		RichValueBlock: 2,
		DynamicArrayProperties: 3,
		DynamicArray: 4,
		CollapsedArray: 5
	};
	var c_oSer_MdxMetadata =
	{
		Mdx: 0,
		NameIndex: 1,
		FunctionTag: 2,
		MdxTuple: 3,
		MdxSet: 4,
		MdxKPI: 5,
		MdxMemeberProp: 6
	};
	var c_oSer_MetadataMdxTuple =
	{
		IndexCount: 0,
		CultureCurrency: 1,
		StringIndex: 2,
		NumFmtIndex: 3,
		BackColor: 4,
		ForeColor: 5,
		Italic: 6,
		Underline: 7,
		Strike: 8,
		Bold: 9,
		MetadataStringIndex: 10
	};
	var c_oSer_MetadataStringIndex =
	{
		StringIsSet: 0,
		IndexValue: 1
	};
	var c_oSer_MetadataMdxSet =
	{
		Count: 0,
		Index: 1,
		SortOrder: 2,
		MetadataStringIndex: 3
	};
	var c_oSer_MetadataMdxKPI =
	{
		NameIndex: 0,
		Index: 1,
		Property: 2
	};
	var c_oSer_MetadataMemberProperty =
	{
		NameIndex: 0,
		Index: 1
	};

    /** @enum */
    var EBorderStyle =
    {
        borderstyleDashDot:  0,
        borderstyleDashDotDot:  1,
        borderstyleDashed:  2,
        borderstyleDotted:  3,
        borderstyleDouble:  4,
        borderstyleHair:  5,
        borderstyleMedium:  6,
        borderstyleMediumDashDot:  7,
        borderstyleMediumDashDotDot:  8,
        borderstyleMediumDashed:  9,
        borderstyleNone: 10,
        borderstyleSlantDashDot: 11,
        borderstyleThick: 12,
        borderstyleThin: 13
    };
    /** @enum */
    var EUnderline =
    {
        underlineDouble:  0,
        underlineDoubleAccounting:  1,
        underlineNone:  2,
        underlineSingle:  3,
        underlineSingleAccounting:  4
    };
    /** @enum */
    var ECellAnchorType =
    {
        cellanchorAbsolute:  0,
        cellanchorOneCell:  1,
        cellanchorTwoCell:  2
    };
    /** @enum */
    var EVisibleType =
    {
        visibleHidden:  0,
        visibleVeryHidden:  1,
        visibleVisible:  2
    };
    /** @enum */
    var ECellTypeType =
    {
        celltypeBool:  0,
        celltypeDate:  1,
        celltypeError:  2,
        celltypeInlineStr:  3,
        celltypeNumber:  4,
        celltypeSharedString:  5,
        celltypeStr:  6
    };
    /** @enum */
    var ECellFormulaType =
    {
        cellformulatypeArray:  0,
        cellformulatypeDataTable:  1,
        cellformulatypeNormal:  2,
        cellformulatypeShared:  3
    };
    /** @enum */
    var EPageOrientation =
    {
        pageorientLandscape: 0,
        pageorientPortrait: 1
    };
    /** @enum */
    var EPageSize =
    {
        pagesizeLetterPaper:  1,
        pagesizeLetterSmall:  2,
        pagesizeTabloidPaper:  3,
        pagesizeLedgerPaper:  4,
        pagesizeLegalPaper:  5,
        pagesizeStatementPaper:  6,
        pagesizeExecutivePaper:  7,
        pagesizeA3Paper:  8,
        pagesizeA4Paper:  9,
        pagesizeA4SmallPaper:  10,
        pagesizeA5Paper:  11,
        pagesizeB4Paper:  12,
        pagesizeB5Paper:  13,
        pagesizeFolioPaper:  14,
        pagesizeQuartoPaper:  15,
        pagesizeStandardPaper1:  16,
        pagesizeStandardPaper2:  17,
        pagesizeNotePaper:  18,
        pagesize9Envelope:  19,
        pagesize10Envelope:  20,
        pagesize11Envelope:  21,
        pagesize12Envelope:  22,
        pagesize14Envelope:  23,
        pagesizeCPaper:  24,
        pagesizeDPaper:  25,
        pagesizeEPaper:  26,
        pagesizeDLEnvelope:  27,
        pagesizeC5Envelope:  28,
        pagesizeC3Envelope:  29,
        pagesizeC4Envelope:  30,
        pagesizeC6Envelope:  31,
        pagesizeC65Envelope:  32,
        pagesizeB4Envelope:  33,
        pagesizeB5Envelope:  34,
        pagesizeB6Envelope:  35,
        pagesizeItalyEnvelope:  36,
        pagesizeMonarchEnvelope:  37,
        pagesize6_3_4Envelope:  38,
        pagesizeUSStandardFanfold:  39,
        pagesizeGermanStandardFanfold:  40,
        pagesizeGermanLegalFanfold:  41,
        pagesizeISOB4:  42,
        pagesizeJapaneseDoublePostcard:  43,
        pagesizeStandardPaper3:  44,
        pagesizeStandardPaper4:  45,
        pagesizeStandardPaper5:  46,
        pagesizeInviteEnvelope:  47,
        pagesizeLetterExtraPaper:  50,
        pagesizeLegalExtraPaper:  51,
        pagesizeTabloidExtraPaper:  52,
        pagesizeA4ExtraPaper:  53,
        pagesizeLetterTransversePaper:  54,
        pagesizeA4TransversePaper:  55,
        pagesizeLetterExtraTransversePaper:  56,
        pagesizeSuperA_SuperA_A4Paper:  57,
        pagesizeSuperB_SuperB_A3Paper:  58,
        pagesizeLetterPlusPaper:  59,
        pagesizeA4PlusPaper:  60,
        pagesizeA5TransversePaper:  61,
        pagesizeJISB5TransversePaper:  62,
        pagesizeA3ExtraPaper:  63,
        pagesizeA5ExtraPaper:  64,
        pagesizeISOB5ExtraPaper:  65,
        pagesizeA2Paper:  66,
        pagesizeA3TransversePaper:  67,
        pagesizeA3ExtraTransversePaper:  68,

        pagesizeEnvelopeChoukei3: 73,
        pagesizeROC16K: 121

    };
    /** @enum */
    var ETotalsRowFunction =
    {
        totalrowfunctionAverage: 1,
        totalrowfunctionCount: 2,
        totalrowfunctionCountNums: 3,
        totalrowfunctionCustom: 4,
        totalrowfunctionMax: 5,
        totalrowfunctionMin: 6,
        totalrowfunctionNone: 7,
        totalrowfunctionStdDev: 8,
        totalrowfunctionSum: 9,
        totalrowfunctionVar: 10
    };
    /** @enum */
    var ESortBy =
    {
        sortbyCellColor: 1,
        sortbyFontColor: 2,
        sortbyIcon: 3,
        sortbyValue: 4
    };
    /** @enum */
    var ECustomFilter =
    {
        customfilterEqual: 1,
        customfilterGreaterThan: 2,
        customfilterGreaterThanOrEqual: 3,
        customfilterLessThan: 4,
        customfilterLessThanOrEqual: 5,
        customfilterNotEqual: 6
    };
    /** @enum */
    var EDateTimeGroup =
    {
        datetimegroupDay: 1,
        datetimegroupHour: 2,
        datetimegroupMinute: 3,
        datetimegroupMonth: 4,
        datetimegroupSecond: 5,
        datetimegroupYear: 6
    };
    /** @enum */
    var ETableStyleType =
    {
        tablestyletypeBlankRow: 0,
        tablestyletypeFirstColumn: 1,
        tablestyletypeFirstColumnStripe: 2,
        tablestyletypeFirstColumnSubheading: 3,
        tablestyletypeFirstHeaderCell: 4,
        tablestyletypeFirstRowStripe: 5,
        tablestyletypeFirstRowSubheading: 6,
        tablestyletypeFirstSubtotalColumn: 7,
        tablestyletypeFirstSubtotalRow: 8,
        tablestyletypeFirstTotalCell: 9,
        tablestyletypeHeaderRow: 10,
        tablestyletypeLastColumn: 11,
        tablestyletypeLastHeaderCell: 12,
        tablestyletypeLastTotalCell: 13,
        tablestyletypePageFieldLabels: 14,
        tablestyletypePageFieldValues: 15,
        tablestyletypeSecondColumnStripe: 16,
        tablestyletypeSecondColumnSubheading: 17,
        tablestyletypeSecondRowStripe: 18,
        tablestyletypeSecondRowSubheading: 19,
        tablestyletypeSecondSubtotalColumn: 20,
        tablestyletypeSecondSubtotalRow: 21,
        tablestyletypeThirdColumnSubheading: 22,
        tablestyletypeThirdRowSubheading: 23,
        tablestyletypeThirdSubtotalColumn: 24,
        tablestyletypeThirdSubtotalRow: 25,
        tablestyletypeTotalRow: 26,
        tablestyletypeWholeTable: 27
    };
    /** @enum */
    var EFontScheme =
    {
        fontschemeMajor: 0,
        fontschemeMinor: 1,
        fontschemeNone: 2
    };
    /** @enum */
    var ECfOperator =
    {
        Operator_beginsWith: 0,
        Operator_between: 1,
        Operator_containsText: 2,
        Operator_endsWith: 3,
        Operator_equal: 4,
        Operator_greaterThan: 5,
        Operator_greaterThanOrEqual: 6,
        Operator_lessThan: 7,
        Operator_lessThanOrEqual: 8,
        Operator_notBetween: 9,
        Operator_notContains: 10,
        Operator_notEqual: 11
    };
    /** @enum */
    var ECfType =
    {
        aboveAverage: 0,
        beginsWith: 1,
        cellIs: 2,
        colorScale: 3,
        containsBlanks: 4,
        containsErrors: 5,
        containsText: 6,
        dataBar: 7,
        duplicateValues: 8,
        expression: 9,
        iconSet: 10,
        notContainsBlanks: 11,
        notContainsErrors: 12,
        notContainsText: 13,
        timePeriod: 14,
        top10: 15,
        uniqueValues: 16,
        endsWith: 17
    };
    /** @enum */
    var EIconSetType =
    {
        Arrows3: 0,
        Arrows3Gray: 1,
        Flags3: 2,
        Signs3: 3,
        Symbols3: 4,
        Symbols3_2: 5,
        Traffic3Lights1: 6,
        Traffic3Lights2: 7,
        Arrows4: 8,
        Arrows4Gray: 9,
        Rating4: 10,
        RedToBlack4: 11,
        Traffic4Lights: 12,
        Arrows5: 13,
        Arrows5Gray: 14,
        Quarters5: 15,
		Rating5: 16,
		Triangles3 : 17,
		Stars3 : 18,
		Boxes5 : 19,
		NoIcons : 20
    };
    var ECfvoType =
    {
        Formula: 0,
        Maximum: 1,
        Minimum: 2,
        Number: 3,
        Percent: 4,
        Percentile: 5,
        AutoMin: 6,
        AutoMax: 7
    };
    var ST_TimePeriod = {
        last7Days : 'last7Days',
        lastMonth : 'lastMonth',
        lastWeek  : 'lastWeek',
        nextMonth : 'nextMonth',
        nextWeek  : 'nextWeek',
        thisMonth : 'thisMonth',
        thisWeek  : 'thisWeek',
        today     : 'today',
        tomorrow  : 'tomorrow',
        yesterday : 'yesterday'
    };
	var EDataBarAxisPosition = {
		automatic: 0,
		middle: 1,
		none: 2
	};
	var EDataBarDirection = {
		context: 0,
		leftToRight: 1,
		rightToLeft: 2
	};
	var EActivePane = {
		bottomLeft: 0,
		bottomRight: 1,
		topLeft: 2,
		topRight: 3
	};

    var ESortMethod = {
        sortmethodNone: 1,
        sortmethodPinYin: 2,
        sortmethodStroke: 3
    };

    var ST_CellComments = {
        none: 0,
        asDisplayed: 1,
        atEnd: 2
    };

    var ST_PrintError = {
        displayed: 0,
        blank: 1,
        dash: 2,
        NA: 3
    };

    var ST_TableType = {
        queryTable: 0,
        worksheet: 1,
        xml: 2
    };

    var ST_PageOrder = {
        downThenOver: 0,
        overThenDown: 1
    };

    var ESheetViewType = {
        normal: 0,
        pageBreakPreview: 1,
        pageLayout: 2
    };

     var EUpdateLinksType = {
        updatelinksAlways:  0,
        updatelinksNever:  1,
        updatelinksUserSet:  2
    };

    var g_nNumsMaxId = 164;

    var DocumentPageSize = new function() {
        this.oSizes = [
            {id:EPageSize.pagesizeLetterPaper, w_mm: 215.9, h_mm: 279.4},
            {id:EPageSize.pagesizeLetterSmall, w_mm: 215.9, h_mm: 279.4},
            {id:EPageSize.pagesizeTabloidPaper, w_mm: 279.4, h_mm: 431.8},
            {id:EPageSize.pagesizeLedgerPaper, w_mm: 431.8, h_mm: 279.4},
            {id:EPageSize.pagesizeLegalPaper, w_mm: 215.9, h_mm: 355.6},
            {id:EPageSize.pagesizeStatementPaper, w_mm: 495.3, h_mm: 215.9},
            {id:EPageSize.pagesizeExecutivePaper, w_mm: 184.2, h_mm: 266.7},
            {id:EPageSize.pagesizeA3Paper, w_mm: 297, h_mm: 420},
            {id:EPageSize.pagesizeA4Paper, w_mm: 210, h_mm: 297},
            {id:EPageSize.pagesizeA4SmallPaper, w_mm: 210, h_mm: 297},
            {id:EPageSize.pagesizeA5Paper, w_mm: 148, h_mm: 210},
            {id:EPageSize.pagesizeB4Paper, w_mm: 250, h_mm: 353},
            {id:EPageSize.pagesizeB5Paper, w_mm: 176, h_mm: 250},
            {id:EPageSize.pagesizeFolioPaper, w_mm: 215.9, h_mm: 330.2},
            {id:EPageSize.pagesizeQuartoPaper, w_mm: 215, h_mm: 275},
            {id:EPageSize.pagesizeStandardPaper1, w_mm: 254, h_mm: 355.6},
            {id:EPageSize.pagesizeStandardPaper2, w_mm: 279.4, h_mm: 431.8},
            {id:EPageSize.pagesizeNotePaper, w_mm: 215.9, h_mm: 279.4},
            {id:EPageSize.pagesize9Envelope, w_mm: 98.4, h_mm: 225.4},
            {id:EPageSize.pagesize10Envelope, w_mm: 104.8, h_mm: 241.3},
            {id:EPageSize.pagesize11Envelope, w_mm: 114.3, h_mm: 263.5},
            {id:EPageSize.pagesize12Envelope, w_mm: 120.7, h_mm: 279.4},
            {id:EPageSize.pagesize14Envelope, w_mm: 127, h_mm: 292.1},
            {id:EPageSize.pagesizeCPaper, w_mm: 431.8, h_mm: 558.8},
            {id:EPageSize.pagesizeDPaper, w_mm: 558.8, h_mm: 863.6},
            {id:EPageSize.pagesizeEPaper, w_mm: 863.6, h_mm: 1117.6},
            {id:EPageSize.pagesizeDLEnvelope, w_mm: 110, h_mm: 220},
            {id:EPageSize.pagesizeC5Envelope, w_mm: 162, h_mm: 229},
            {id:EPageSize.pagesizeC3Envelope, w_mm: 324, h_mm: 458},
            {id:EPageSize.pagesizeC4Envelope, w_mm: 229, h_mm: 324},
            {id:EPageSize.pagesizeC6Envelope, w_mm: 114, h_mm: 162},
            {id:EPageSize.pagesizeC65Envelope, w_mm: 114, h_mm: 229},
            {id:EPageSize.pagesizeB4Envelope, w_mm: 250, h_mm: 353},
            {id:EPageSize.pagesizeB5Envelope, w_mm: 176, h_mm: 250},
            {id:EPageSize.pagesizeB6Envelope, w_mm: 176, h_mm: 125},
            {id:EPageSize.pagesizeItalyEnvelope, w_mm: 110, h_mm: 230},
            {id:EPageSize.pagesizeMonarchEnvelope, w_mm: 98.4, h_mm: 190.5},
            {id:EPageSize.pagesize6_3_4Envelope, w_mm: 92.1, h_mm: 165.1},
            {id:EPageSize.pagesizeUSStandardFanfold, w_mm: 377.8, h_mm: 279.4},
            {id:EPageSize.pagesizeGermanStandardFanfold, w_mm: 215.9, h_mm: 304.8},
            {id:EPageSize.pagesizeGermanLegalFanfold, w_mm: 215.9, h_mm: 330.2},
            {id:EPageSize.pagesizeISOB4, w_mm: 250, h_mm: 353},
            {id:EPageSize.pagesizeJapaneseDoublePostcard, w_mm: 200, h_mm: 148},
            {id:EPageSize.pagesizeStandardPaper3, w_mm: 228.6, h_mm: 279.4},
            {id:EPageSize.pagesizeStandardPaper4, w_mm: 254, h_mm: 279.4},
            {id:EPageSize.pagesizeStandardPaper5, w_mm: 381, h_mm: 279.4},
            {id:EPageSize.pagesizeInviteEnvelope, w_mm: 220, h_mm: 220},
            {id:EPageSize.pagesizeLetterExtraPaper, w_mm: 235.6, h_mm: 304.8},
            {id:EPageSize.pagesizeLegalExtraPaper, w_mm: 235.6, h_mm: 381},
            {id:EPageSize.pagesizeTabloidExtraPaper, w_mm: 296.9, h_mm: 457.2},
            {id:EPageSize.pagesizeA4ExtraPaper, w_mm: 236, h_mm: 322},
            {id:EPageSize.pagesizeLetterTransversePaper, w_mm: 210.2, h_mm: 279.4},
            {id:EPageSize.pagesizeA4TransversePaper, w_mm: 210, h_mm: 297},
            {id:EPageSize.pagesizeLetterExtraTransversePaper, w_mm: 235.6, h_mm: 304.8},
            {id:EPageSize.pagesizeSuperA_SuperA_A4Paper, w_mm: 227, h_mm: 356},
            {id:EPageSize.pagesizeSuperB_SuperB_A3Paper, w_mm: 305, h_mm: 487},
            {id:EPageSize.pagesizeLetterPlusPaper, w_mm: 215.9, h_mm: 12.69},
            {id:EPageSize.pagesizeA4PlusPaper, w_mm: 210, h_mm: 330},
            {id:EPageSize.pagesizeA5TransversePaper, w_mm: 148, h_mm: 210},
            {id:EPageSize.pagesizeJISB5TransversePaper, w_mm: 182, h_mm: 257},
            {id:EPageSize.pagesizeA3ExtraPaper, w_mm: 322, h_mm: 445},
            {id:EPageSize.pagesizeA5ExtraPaper, w_mm: 174, h_mm: 235},
            {id:EPageSize.pagesizeISOB5ExtraPaper, w_mm: 201, h_mm: 276},
            {id:EPageSize.pagesizeA2Paper, w_mm: 420, h_mm: 594},
            {id:EPageSize.pagesizeA3TransversePaper, w_mm: 297, h_mm: 420},
            {id:EPageSize.pagesizeA3ExtraTransversePaper, w_mm: 322, h_mm: 445},
            {id:EPageSize.pagesizeEnvelopeChoukei3, w_mm: 120, h_mm: 235},
            {id:EPageSize.pagesizeROC16K, w_mm: 196.8, h_mm: 273}
        ];
        this.getSizeByWH = function(widthMm, heightMm)
        {
            for( var index in this.oSizes)
            {
                var item = this.oSizes[index];
                if(widthMm == item.w_mm && heightMm == item.h_mm)
                    return item;
            }
            return this.oSizes[8];//A4
        };
        this.getSizeById = function(id)
        {
            for( var index in this.oSizes)
            {
                var item = this.oSizes[index];
                if(id == item.id)
                    return item;
            }
            return this.oSizes[8];//A4
        };
    };

    function OpenColor() {
        this.rgb = null;
        this.auto = null;
        this.theme = null;
        this.tint = null;
    }

	function OpenFormula() {
		this.aca = null;
		this.bx = null;
		this.ca = null;
		this.del1 = null;
		this.del2 = null;
		this.dt2d = null;
		this.dtr = null;
		this.r1 = null;
		this.r2 = null;
		this.ref = null;
		this.si = null;
		this.t = null;
		this.v = null;
	}
	OpenFormula.prototype.clean = function(){
		this.aca = null;
		this.bx = null;
		this.ca = null;
		this.del1 = null;
		this.del2 = null;
		this.dt2d = null;
		this.dtr = null;
		this.r1 = null;
		this.r2 = null;
		this.ref = null;
		this.si = null;
		this.t = null;
		this.v = null;
	};
	function OpenColumnFormula(nRow, formula, parsed, refPos, base) {
		this.nRow = nRow;
		this.formula = formula;
		this.parsed = parsed;
		this.refPos = refPos;
		this.base = base;
	}

	function OpenXf(){
		this.ApplyAlignment = null;
		this.ApplyBorder = null;
		this.ApplyFill = null;
		this.ApplyFont = null;
		this.ApplyNumberFormat = null;
		this.borderid = null;
		this.fillid = null;
		this.fontid = null;
		this.numid = null;
		this.QuotePrefix = null;
		this.align = null;
		this.PivotButton = null;
		this.XfId = null;

		this.applyProtection = null;
		this.locked = null;
		this.hidden = null;
	}

	function ReadColorSpreadsheet2(bcr, length) {
		var output = null;
		var color = new OpenColor();
		var res = bcr.Read2Spreadsheet(length, function(t,l){
			return bcr.ReadColorSpreadsheet(t,l, color);
		});
		if(null != color.theme)
			output = AscCommonExcel.g_oColorManager.getThemeColor(color.theme, color.tint);
		else if(null != color.rgb)
			output = new AscCommonExcel.RgbColor(0x00ffffff & color.rgb);
		return output;
	}

	function getSqRefString(ranges) {
		var refs = [];
		for (var i = 0; i < ranges.length; ++i) {
			refs.push(ranges[i].getName());
		}
		return refs.join(' ');
	}

    function getDisjointMerged(wb, bboxes) {
        var res = [];
        var curY, elem;
        var error = false;
        var indexTop = 0;
        var indexBottom = 0;
        var rangesTop = bboxes;
        var rangesBottom = bboxes.concat();
        rangesTop.sort(Asc.Range.prototype.compareByLeftTop);
        rangesBottom.sort(Asc.Range.prototype.compareByRightBottom);
        var tree = new AscCommon.DataIntervalTree();
        while (indexBottom < rangesBottom.length) {
            //next curY
            if (indexTop < rangesTop.length) {
                curY = Math.min(rangesTop[indexTop].r1, rangesBottom[indexBottom].r2);
            } else {
                curY = rangesBottom[indexBottom].r2;
            }
            while (indexTop < rangesTop.length && curY === rangesTop[indexTop].r1) {
                elem = rangesTop[indexTop];
                if (!tree.searchAny(elem.c1, elem.c2)) {
                    tree.insert(elem.c1, elem.c2, elem);
                    res.push(elem);
                } else {
                    error = true;
                }
                indexTop++;
            }
            while (indexBottom < rangesBottom.length && curY === rangesBottom[indexBottom].r2) {
                elem = rangesBottom[indexBottom];
                tree.remove(elem.c1, elem.c2, elem);
                indexBottom++;
            }
        }
        if (error && wb.oApi && wb.oApi.CoAuthoringApi) {
            var msg = 'Error: intersection of merged areas';
            AscCommon.sendClientLog("error", "changesError: " + msg, wb.oApi);
        }
        return res;
	}

    function checkMaxCellLength(text) {
        if (text && text.length > Asc.c_oAscMaxCellOrCommentLength) {
            text = text.slice(0, Asc.c_oAscMaxCellOrCommentLength);
        }
        return text;
    }

    //TODO копия кода из serialize2
    function BinaryCustomsTableWriter(memory, CustomXmls)
    {
        this.memory = memory;
        //this.Document = doc;
        this.bs = new BinaryCommonWriter(this.memory);
        this.CustomXmls = CustomXmls;
        this.Write = function()
        {
            var oThis = this;
            this.bs.WriteItemWithLength(function(){oThis.WriteCustomXmls();});
        };
        this.WriteCustomXmls = function()
        {
            var oThis = this;
            for (var i = 0; i < this.CustomXmls.length; ++i) {
                this.bs.WriteItem(c_oSerCustoms.Custom, function() {oThis.WriteCustomXml(oThis.CustomXmls[i]);});
            }
        };
        this.WriteCustomXml = function(customXml) {
            var oThis = this;
            for(var i = 0; i < customXml.Uri.length; ++i){
                this.bs.WriteItem(c_oSerCustoms.Uri, function () {
                    oThis.memory.WriteString3(customXml.Uri[i]);
                });
            }
            if (null !== customXml.ItemId) {
                this.bs.WriteItem(c_oSerCustoms.ItemId, function() {
                    oThis.memory.WriteString3(customXml.ItemId);
                });
            }
            if (null !== customXml.Content) {
                this.bs.WriteItem(c_oSerCustoms.Content, function() {
                    oThis.memory.WriteBuffer(customXml.Content, 0, customXml.Content.length)
                });
            }
        };
    }

    /** @constructor */
    function BinaryTableWriter(memory, InitSaveManager, isCopyPaste, tableIds)
    {
        this.memory = memory;
        this.InitSaveManager = InitSaveManager;
        this.bs = new BinaryCommonWriter(this.memory);
        this.isCopyPaste = isCopyPaste;
        this.tableIds = tableIds;
        this.Write = function(aTables, ws)
        {
            var oThis = this;
            for(var i = 0, length = aTables.length; i < length; ++i)
            {
                var rangeTable = null;
                //get range for copy/paste
                if (this.isCopyPaste)
                    rangeTable = aTables[i].Ref;

                if(!this.isCopyPaste || (this.isCopyPaste && rangeTable && this.isCopyPaste.isIntersect(rangeTable) && !ws.bExcludeHiddenRows))
                    this.bs.WriteItem(c_oSer_TablePart.Table, function(){oThis.WriteTable(aTables[i], ws);});
            }
        };
        this.WriteTable = function(table, ws)
        {
            var oThis = this;
            //Ref
            if(null != table.Ref)
            {
                this.memory.WriteByte(c_oSer_TablePart.Ref);
                this.memory.WriteString2(table.Ref.getName());
            }
            //HeaderRowCount
            if(null != table.HeaderRowCount)
                this.bs.WriteItem(c_oSer_TablePart.HeaderRowCount, function(){oThis.memory.WriteLong(table.HeaderRowCount);});
            //TotalsRowCount
            if(null != table.TotalsRowCount)
                this.bs.WriteItem(c_oSer_TablePart.TotalsRowCount, function(){oThis.memory.WriteLong(table.TotalsRowCount);});
            //Display Name
            if(null != table.DisplayName)
            {
                this.memory.WriteByte(c_oSer_TablePart.DisplayName);
                this.memory.WriteString2(table.DisplayName);

                if (this.tableIds) {
                    var elem = this.tableIds[table.DisplayName];
                    if (elem) {
                        this.bs.WriteItem(c_oSer_TablePart.Id, function(){oThis.memory.WriteULong(elem.id);});
                    }
                }
            }
            //AutoFilter
            if(null != table.AutoFilter)
                this.bs.WriteItem(c_oSer_TablePart.AutoFilter, function(){oThis.WriteAutoFilter(table.AutoFilter);});
            //SortState
            if(null != table.SortState)
                this.bs.WriteItem(c_oSer_TablePart.SortState, function(){oThis.WriteSortState(table.SortState);});
            //TableColumns
            if(null != table.TableColumns) {
                if (ws) {
                    //TODO пока оставляю. проверить необходим ли до сих пор вызов данной функции
                    table.syncTotalLabels(ws);
                }
                this.bs.WriteItem(c_oSer_TablePart.TableColumns, function(){oThis.WriteTableColumns(table.TableColumns);});
			}
            //TableStyleInfo
            if(null != table.TableStyleInfo)
                this.bs.WriteItem(c_oSer_TablePart.TableStyleInfo, function(){oThis.WriteTableStyleInfo(table.TableStyleInfo);});
			if(null != table.altText || null != table.altTextSummary)
				this.bs.WriteItem(c_oSer_TablePart.AltTextTable, function(){oThis.WriteAltTextTable(table);});

			if(null != table.QueryTable)
				this.bs.WriteItem(c_oSer_TablePart.QueryTable, function(){oThis.WriteQueryTable(table.QueryTable, table);});

			if(null != table.tableType)
				this.bs.WriteItem(c_oSer_TablePart.TableType, function(){oThis.memory.WriteULong(table.tableType);});
        };
		this.WriteAltTextTable = function(table)
		{
			var oThis = this;
			if (null != table.altText) {
				this.memory.WriteByte(c_oSer_AltTextTable.AltText);
				this.memory.WriteString2(table.altText);
			}
			if (null != table.altTextSummary) {
				this.memory.WriteByte(c_oSer_AltTextTable.AltTextSummary);
				this.memory.WriteString2(table.altTextSummary);
			}
 		};
        this.WriteAutoFilter = function(autofilter)
        {
            var oThis = this;
            //Ref
            if(null != autofilter.Ref)
            {
				this.memory.WriteByte(c_oSer_AutoFilter.Ref);
                this.memory.WriteString2(autofilter.Ref.getName());
            }
            //FilterColumns
            if(null != autofilter.FilterColumns)
                this.bs.WriteItem(c_oSer_AutoFilter.FilterColumns, function(){oThis.WriteFilterColumns(autofilter.FilterColumns);});
            //SortState
            if(null != autofilter.SortState)
                this.bs.WriteItem(c_oSer_AutoFilter.SortState, function(){oThis.WriteSortState(autofilter.SortState);});
        };
        this.WriteFilterColumns = function(filterColumns)
        {
            var oThis = this;
            for(var i = 0, length = filterColumns.length; i < length; ++i)
                this.bs.WriteItem(c_oSer_AutoFilter.FilterColumn, function(){oThis.WriteFilterColumn(filterColumns[i]);});
        };
        this.WriteFilterColumn = function(filterColumn)
        {
            var oThis = this;
            //ColId
            if(null != filterColumn.ColId)
                this.bs.WriteItem(c_oSer_FilterColumn.ColId, function(){oThis.memory.WriteLong(filterColumn.ColId);});
            //Filters
            if(null != filterColumn.Filters)
                this.bs.WriteItem(c_oSer_FilterColumn.Filters, function(){oThis.WriteFilters(filterColumn.Filters);});
            //CustomFilters
            if(null != filterColumn.CustomFiltersObj)
                this.bs.WriteItem(c_oSer_FilterColumn.CustomFilters, function(){oThis.WriteCustomFilters(filterColumn.CustomFiltersObj);});
            //DynamicFilter
            if(null != filterColumn.DynamicFilter)
                this.bs.WriteItem(c_oSer_FilterColumn.DynamicFilter, function(){oThis.WriteDynamicFilter(filterColumn.DynamicFilter);});
            //ColorFilter
            if(null != filterColumn.ColorFilter)
                this.bs.WriteItem(c_oSer_FilterColumn.ColorFilter, function(){oThis.WriteColorFilter(filterColumn.ColorFilter);});
            //Top10
            if(null != filterColumn.Top10)
                this.bs.WriteItem(c_oSer_FilterColumn.Top10, function(){oThis.WriteTop10(filterColumn.Top10);});
            //ShowButton
            if(null != filterColumn.ShowButton)
                this.bs.WriteItem(c_oSer_FilterColumn.ShowButton, function(){oThis.memory.WriteBool(filterColumn.ShowButton);});
        };
        this.WriteFilters = function(filters)
        {
            var oThis = this;
            if(null != filters.Values)
            {
				for(var i in filters.Values)
					this.bs.WriteItem(c_oSer_FilterColumn.Filter, function(){oThis.WriteFilter(i);});
            }
            if(null != filters.Dates)
            {
                for(var i = 0, length = filters.Dates.length; i < length; ++i)
                    this.bs.WriteItem(c_oSer_FilterColumn.DateGroupItem, function(){oThis.WriteDateGroupItem(filters.Dates[i]);});
            }
            if(null != filters.Blank)
                this.bs.WriteItem(c_oSer_FilterColumn.FiltersBlank, function(){oThis.memory.WriteBool(filters.Blank);});
        };
        this.WriteFilter = function(val)
        {
            if(null != val)
            {
                this.memory.WriteByte(c_oSer_Filter.Val);
                this.memory.WriteString2(val);
            }
        };
        this.WriteDateGroupItem = function(dateGroupItem)
        {
			var oDateGroupItem = new AscCommonExcel.DateGroupItem();
			oDateGroupItem.convertRangeToDateGroupItem(dateGroupItem);
			dateGroupItem = oDateGroupItem;

			if(null != dateGroupItem.DateTimeGrouping)
            {
                this.memory.WriteByte(c_oSer_DateGroupItem.DateTimeGrouping);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(dateGroupItem.DateTimeGrouping);
            }
            if(null != dateGroupItem.Day)
            {
                this.memory.WriteByte(c_oSer_DateGroupItem.Day);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(dateGroupItem.Day);
            }
            if(null != dateGroupItem.Hour)
            {
                this.memory.WriteByte(c_oSer_DateGroupItem.Hour);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(dateGroupItem.Hour);
            }
            if(null != dateGroupItem.Minute)
            {
                this.memory.WriteByte(c_oSer_DateGroupItem.Minute);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(dateGroupItem.Minute);
            }
            if(null != dateGroupItem.Month)
            {
                this.memory.WriteByte(c_oSer_DateGroupItem.Month);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(dateGroupItem.Month);
            }
            if(null != dateGroupItem.Second)
            {
                this.memory.WriteByte(c_oSer_DateGroupItem.Second);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(dateGroupItem.Second);
            }
            if(null != dateGroupItem.Year)
            {
                this.memory.WriteByte(c_oSer_DateGroupItem.Year);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(dateGroupItem.Year);
            }
        };
        this.WriteCustomFilters = function(customFilters)
        {
            var oThis = this;
            if(null != customFilters.And)
                this.bs.WriteItem(c_oSer_CustomFilters.And, function(){oThis.memory.WriteBool(customFilters.And);});
            if(null != customFilters.CustomFilters && customFilters.CustomFilters.length > 0)
                this.bs.WriteItem(c_oSer_CustomFilters.CustomFilters, function(){oThis.WriteCustomFiltersItems(customFilters.CustomFilters);});
        };
        this.WriteCustomFiltersItems = function(aCustomFilters)
        {
            var oThis = this;
            for(var i = 0, length = aCustomFilters.length; i < length; ++i)
                this.bs.WriteItem(c_oSer_CustomFilters.CustomFilter, function(){oThis.WriteCustomFiltersItem(aCustomFilters[i]);});
        };
        this.WriteCustomFiltersItem = function(customFilter)
        {
            if(null != customFilter.Operator)
            {
                this.memory.WriteByte(c_oSer_CustomFilters.Operator);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(customFilter.Operator);
            }
            if(null != customFilter.Val)
            {
                this.memory.WriteByte(c_oSer_CustomFilters.Val);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(customFilter.Val);
            }
        };
        this.WriteDynamicFilter = function(dynamicFilter)
        {
            if(null != dynamicFilter.Type)
            {
                this.memory.WriteByte(c_oSer_DynamicFilter.Type);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(dynamicFilter.Type);
            }
            if(null != dynamicFilter.Val)
            {
                this.memory.WriteByte(c_oSer_DynamicFilter.Val);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(dynamicFilter.Val);
            }
            if(null != dynamicFilter.MaxVal)
            {
                this.memory.WriteByte(c_oSer_DynamicFilter.MaxVal);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(dynamicFilter.MaxVal);
            }
        };
        this.WriteColorFilter = function(colorFilter)
        {
            if(null != colorFilter.CellColor)
            {
                this.memory.WriteByte(c_oSer_ColorFilter.CellColor);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(colorFilter.CellColor);
            }
            if(null != colorFilter.dxf)
            {
                this.memory.WriteByte(c_oSer_ColorFilter.DxfId);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(this.InitSaveManager.aDxfs.length);
                this.InitSaveManager.aDxfs.push(colorFilter.dxf);
            }
        };
        this.WriteTop10 = function(top10)
        {
            if(null != top10.FilterVal)
            {
                this.memory.WriteByte(c_oSer_Top10.FilterVal);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(top10.FilterVal);
            }
            if(null != top10.Percent)
            {
                this.memory.WriteByte(c_oSer_Top10.Percent);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(top10.Percent);
            }
            if(null != top10.Top)
            {
                this.memory.WriteByte(c_oSer_Top10.Top);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(top10.Top);
            }
            if(null != top10.Val)
            {
                this.memory.WriteByte(c_oSer_Top10.Val);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(top10.Val);
            }
        };
        this.WriteSortState = function(sortState)
        {
            var oThis = this;
            if(null != sortState.Ref)
            {
                this.memory.WriteByte(c_oSer_SortState.Ref);
                this.memory.WriteString2(sortState.Ref.getName());
            }
            if(null != sortState.CaseSensitive)
                this.bs.WriteItem(c_oSer_SortState.CaseSensitive, function(){oThis.memory.WriteBool(sortState.CaseSensitive);});
            if(null != sortState.ColumnSort)
                this.bs.WriteItem(c_oSer_SortState.ColumnSort, function(){oThis.memory.WriteBool(sortState.ColumnSort);});
            if(null != sortState.SortMethod)
                this.bs.WriteItem(c_oSer_SortState.SortMethod, function(){oThis.memory.WriteByte(sortState.SortMethod);});
            if(null != sortState.SortConditions)
                this.bs.WriteItem(c_oSer_SortState.SortConditions, function(){oThis.WriteSortConditions(sortState.SortConditions);});
        };
        this.WriteSortConditions = function(sortConditions)
        {
            var oThis = this;
            for(var i = 0, length = sortConditions.length; i < length; ++i) {
                if (sortConditions[i]) {
                    this.bs.WriteItem(c_oSer_SortState.SortCondition, function(){oThis.WriteSortCondition(sortConditions[i]);});
                }
            }

        };
        this.WriteSortCondition = function(sortCondition)
        {
            if(null != sortCondition.Ref)
            {
                this.memory.WriteByte(c_oSer_SortState.ConditionRef);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(sortCondition.Ref.getName());
            }
            if(null != sortCondition.ConditionSortBy)
            {
                this.memory.WriteByte(c_oSer_SortState.ConditionSortBy);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(sortCondition.ConditionSortBy);
            }
            if(null != sortCondition.ConditionDescending)
            {
                this.memory.WriteByte(c_oSer_SortState.ConditionDescending);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(sortCondition.ConditionDescending);
            }
            if(null != sortCondition.dxf)
            {
                this.memory.WriteByte(c_oSer_SortState.ConditionDxfId);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(this.InitSaveManager.aDxfs.length);
                this.InitSaveManager.aDxfs.push(sortCondition.dxf);
            }
        };
        this.WriteTableColumns = function(tableColumns)
        {
            var oThis = this;
            for(var i = 0, length = tableColumns.length; i < length; ++i)
                this.bs.WriteItem(c_oSer_TableColumns.TableColumn, function(){oThis.WriteTableColumn(tableColumns[i]);});
        };
        this.WriteTableXmlColumnPr = function (xmlColumnPr) {
            if (xmlColumnPr.mapId != null) {
                this.memory.WriteByte(c_oSer_TableColumns.MapId);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(xmlColumnPr.mapId);
            }
            if (xmlColumnPr.xpath != null) {
                this.memory.WriteByte(c_oSer_TableColumns.Xpath);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(xmlColumnPr.xpath);
            }
            if (xmlColumnPr.denormalized != null) {
                this.memory.WriteByte(c_oSer_TableColumns.Denormalized);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(xmlColumnPr.denormalized);
            }
            if (xmlColumnPr.xmlDataType != null) {
                this.memory.WriteByte(c_oSer_TableColumns.XmlDataType);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(xmlColumnPr.xmlDataType);
            }
        };
        this.WriteTableColumn = function(tableColumn)
        {
            var oThis = this;
            if(null != tableColumn.Name)
            {
                var columnName = tableColumn.Name.replaceAll("\n", "_x000a_");
                this.memory.WriteByte(c_oSer_TableColumns.Name);
                this.memory.WriteString2(columnName);
            }
            if(null != tableColumn.TotalsRowLabel)
            {
                this.memory.WriteByte(c_oSer_TableColumns.TotalsRowLabel);
                this.memory.WriteString2(tableColumn.TotalsRowLabel);
            }
            if(null != tableColumn.TotalsRowFunction)
                this.bs.WriteItem(c_oSer_TableColumns.TotalsRowFunction, function(){oThis.memory.WriteByte(tableColumn.TotalsRowFunction);});

            if(null != tableColumn.TotalsRowFormula)
            {
                this.memory.WriteByte(c_oSer_TableColumns.TotalsRowFormula);
                this.memory.WriteString2(tableColumn.TotalsRowFormula.getFormula());
            }
            if(null != tableColumn.dxf)
            {
                this.bs.WriteItem(c_oSer_TableColumns.DataDxfId, function(){oThis.memory.WriteLong(oThis.InitSaveManager.aDxfs.length);});
                this.InitSaveManager.aDxfs.push(tableColumn.dxf);
            }
            if(null != tableColumn.CalculatedColumnFormula)
            {
                this.memory.WriteByte(c_oSer_TableColumns.CalculatedColumnFormula);
                this.memory.WriteString2(tableColumn.CalculatedColumnFormula);
            }
			if(null != tableColumn.uniqueName)
			{
				this.memory.WriteByte(c_oSer_TableColumns.UniqueName);
				this.memory.WriteString2(tableColumn.uniqueName);
			}
			if(null != tableColumn.queryTableFieldId)
			{
				this.bs.WriteItem(c_oSer_TableColumns.QueryTableFieldId, function () {
					oThis.memory.WriteLong(tableColumn.queryTableFieldId);
				});
			}
            if(null != tableColumn.xmlColumnPr)
            {
                this.bs.WriteItem(c_oSer_TableColumns.XmlColumnPr, function () {
                    oThis.WriteTableXmlColumnPr(tableColumn.xmlColumnPr);
                });
            }
        };
        this.WriteTableStyleInfo = function(tableStyleInfo)
        {
            if(null != tableStyleInfo.Name)
            {
                this.memory.WriteByte(c_oSer_TableStyleInfo.Name);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(tableStyleInfo.Name);
            }
            if(null != tableStyleInfo.ShowColumnStripes)
            {
                this.memory.WriteByte(c_oSer_TableStyleInfo.ShowColumnStripes);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(tableStyleInfo.ShowColumnStripes);
            }
            if(null != tableStyleInfo.ShowRowStripes)
            {
                this.memory.WriteByte(c_oSer_TableStyleInfo.ShowRowStripes);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(tableStyleInfo.ShowRowStripes);
            }
            if(null != tableStyleInfo.ShowFirstColumn)
            {
                this.memory.WriteByte(c_oSer_TableStyleInfo.ShowFirstColumn);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(tableStyleInfo.ShowFirstColumn);
            }
            if(null != tableStyleInfo.ShowLastColumn)
            {
                this.memory.WriteByte(c_oSer_TableStyleInfo.ShowLastColumn);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(tableStyleInfo.ShowLastColumn);
            }
        };
		this.WriteQueryTable = function(queryTable, table)
		{
			var oThis = this;

			if (null != queryTable.connectionId) {
				this.bs.WriteItem(c_oSer_QueryTable.ConnectionId, function () {
					oThis.memory.WriteLong(queryTable.connectionId);
				});
			}
			if (null != queryTable.name) {
				this.bs.WriteItem(c_oSer_QueryTable.Name, function () {
					oThis.memory.WriteString3(queryTable.name);
				});
			}
			if (null != queryTable.autoFormatId) {
				this.bs.WriteItem(c_oSer_QueryTable.AutoFormatId, function () {
					oThis.memory.WriteLong(queryTable.autoFormatId);
				});
			}
			if (null != queryTable.growShrinkType) {
				this.bs.WriteItem(c_oSer_QueryTable.GrowShrinkType, function () {
					oThis.memory.WriteString2(queryTable.growShrinkType);
				});
			}
			if (null != queryTable.adjustColumnWidth) {
				this.bs.WriteItem(c_oSer_QueryTable.AdjustColumnWidth, function () {
					oThis.memory.WriteBool(queryTable.adjustColumnWidth);
				});
			}
			if (null != queryTable.applyAlignmentFormats) {
				this.bs.WriteItem(c_oSer_QueryTable.ApplyAlignmentFormats, function () {
					oThis.memory.WriteBool(queryTable.applyAlignmentFormats);
				});
			}
			if (null != queryTable.applyBorderFormats) {
				this.bs.WriteItem(c_oSer_QueryTable.ApplyBorderFormats, function () {
					oThis.memory.WriteBool(queryTable.applyBorderFormats);
				});
			}
			if (null != queryTable.applyFontFormats) {
				this.bs.WriteItem(c_oSer_QueryTable.ApplyFontFormats, function () {
					oThis.memory.WriteBool(queryTable.applyFontFormats);
				});
			}
			if (null != queryTable.applyNumberFormats) {
				this.bs.WriteItem(c_oSer_QueryTable.ApplyNumberFormats, function () {
					oThis.memory.WriteBool(queryTable.applyNumberFormats);
				});
			}
			if (null != queryTable.applyPatternFormats) {
				this.bs.WriteItem(c_oSer_QueryTable.ApplyPatternFormats, function () {
					oThis.memory.WriteBool(queryTable.applyPatternFormats);
				});
			}
			if (null != queryTable.applyWidthHeightFormats) {
				this.bs.WriteItem(c_oSer_QueryTable.ApplyWidthHeightFormats, function () {
					oThis.memory.WriteBool(queryTable.applyWidthHeightFormats);
				});
			}
			if (null != queryTable.backgroundRefresh) {
				this.bs.WriteItem(c_oSer_QueryTable.BackgroundRefresh, function () {
					oThis.memory.WriteBool(queryTable.backgroundRefresh);
				});
			}
			if (null != queryTable.disableEdit) {
				this.bs.WriteItem(c_oSer_QueryTable.DisableEdit, function () {
					oThis.memory.WriteBool(queryTable.disableEdit);
				});
			}
			if (null != queryTable.disableRefresh) {
				this.bs.WriteItem(c_oSer_QueryTable.DisableRefresh, function () {
					oThis.memory.WriteBool(queryTable.disableRefresh);
				});
			}
			if (null != queryTable.fillFormulas) {
				this.bs.WriteItem(c_oSer_QueryTable.FillFormulas, function () {
					oThis.memory.WriteBool(queryTable.fillFormulas);
				});
			}
			if (null != queryTable.firstBackgroundRefresh) {
				this.bs.WriteItem(c_oSer_QueryTable.FirstBackgroundRefresh, function () {
					oThis.memory.WriteBool(queryTable.firstBackgroundRefresh);
				});
			}
			if (null != queryTable.headers) {
				this.bs.WriteItem(c_oSer_QueryTable.Headers, function () {
					oThis.memory.WriteBool(queryTable.headers);
				});
			}
			if (null != queryTable.preserveFormatting) {
				this.bs.WriteItem(c_oSer_QueryTable.PreserveFormatting, function () {
					oThis.memory.WriteBool(queryTable.preserveFormatting);
				});
			}
			if (null != queryTable.refreshOnLoad) {
				this.bs.WriteItem(c_oSer_QueryTable.RefreshOnLoad, function () {
					oThis.memory.WriteBool(queryTable.refreshOnLoad);
				});
			}
			if (null != queryTable.removeDataOnSave) {
				this.bs.WriteItem(c_oSer_QueryTable.RemoveDataOnSave, function () {
					oThis.memory.WriteBool(queryTable.removeDataOnSave);
				});
			}
			if (null != queryTable.rowNumbers) {
				this.bs.WriteItem(c_oSer_QueryTable.RowNumbers, function () {
					oThis.memory.WriteBool(queryTable.rowNumbers);
				});
			}
			if (null != queryTable.queryTableRefresh) {
				this.bs.WriteItem(c_oSer_QueryTable.QueryTableRefresh, function () {
					oThis.WriteQueryTableRefresh(queryTable.queryTableRefresh, table);
				});
			}
		};
		this.WriteQueryTableRefresh = function(queryTableRefresh, table)
		{
			var oThis = this;
			if (null != queryTableRefresh.nextId) {
				this.bs.WriteItem(c_oSer_QueryTableRefresh.NextId, function () {
					oThis.memory.WriteLong(queryTableRefresh.nextId);
				});
			}
			if (null != queryTableRefresh.minimumVersion) {
				this.bs.WriteItem(c_oSer_QueryTableRefresh.MinimumVersion, function () {
					oThis.memory.WriteLong(queryTableRefresh.minimumVersion);
				});
			}
			if (null != queryTableRefresh.unboundColumnsLeft) {
				this.bs.WriteItem(c_oSer_QueryTableRefresh.UnboundColumnsLeft, function () {
					oThis.memory.WriteLong(queryTableRefresh.unboundColumnsLeft);
				});
			}
			if (null != queryTableRefresh.unboundColumnsRight) {
				this.bs.WriteItem(c_oSer_QueryTableRefresh.UnboundColumnsRight, function () {
					oThis.memory.WriteLong(queryTableRefresh.unboundColumnsRight);
				});
			}
			if (null != queryTableRefresh.fieldIdWrapped) {
				this.bs.WriteItem(c_oSer_QueryTableRefresh.FieldIdWrapped, function () {
					oThis.memory.WriteBool(queryTableRefresh.feldIdWrapped);
				});
			}
			if (null != queryTableRefresh.headersInLastRefresh) {
				this.bs.WriteItem(c_oSer_QueryTableRefresh.HeadersInLastRefresh, function () {
					oThis.memory.WriteBool(queryTableRefresh.headersInLastRefresh);
				});
			}
			if (null != queryTableRefresh.preserveSortFilterLayout) {
				this.bs.WriteItem(c_oSer_QueryTableRefresh.PreserveSortFilterLayout, function () {
					oThis.memory.WriteBool(queryTableRefresh.preserveSortFilterLayout);
				});
			}

			if (null != queryTableRefresh.sortState) {
				this.bs.WriteItem(c_oSer_QueryTableRefresh.SortState, function () {
					oThis.WriteSortState(queryTableRefresh.sortState);
				});
			}
			if (null != queryTableRefresh.queryTableFields) {
				this.bs.WriteItem(c_oSer_QueryTableRefresh.QueryTableFields, function () {
					oThis.WriteQueryTableFields(queryTableRefresh.queryTableFields, table);
				});
			}
			if (null != queryTableRefresh.queryTableDeletedFields) {
				this.bs.WriteItem(c_oSer_QueryTableRefresh.QueryTableDeletedFields, function () {
					oThis.WriteQueryTableDeletedFields(queryTableRefresh.queryTableDeletedFields/*,table*/);
				});
			}
		};
		this.WriteQueryTableFields = function (queryTableFields, table) {
			var oThis = this;
			for (var i = 0, length = queryTableFields.length; i < length; ++i) {
                //нужна синхронизация по id таблицы. поскольку, id генерируется в x2t по индексу колонки, то отправляем на сохранение именно индекс колонки
                var tableColumnId = table.getIndexTableColumnById(queryTableFields[i].tableColumnId);
                if (tableColumnId !== null) {
                    this.bs.WriteItem(c_oSer_QueryTableField.QueryTableField, function () {
                        oThis.WriteQueryTableField(queryTableFields[i], tableColumnId);
                    });
                }
            }
        };
		this.WriteQueryTableField = function (queryTableField, tableColumnId) {

			var oThis = this;
			if (null != queryTableField.name) {
				this.bs.WriteItem(c_oSer_QueryTableField.Name, function () {
					oThis.memory.WriteString3(queryTableField.name);
				});
			}
			if (null != queryTableField.id) {
				this.bs.WriteItem(c_oSer_QueryTableField.Id, function () {
					oThis.memory.WriteLong(queryTableField.id);
				});
			}
			if (null != tableColumnId) {
				this.bs.WriteItem(c_oSer_QueryTableField.TableColumnId, function () {
					oThis.memory.WriteLong(tableColumnId);
				});
			}
			if (null != queryTableField.rowNumbers) {
				this.bs.WriteItem(c_oSer_QueryTableField.RowNumbers, function () {
					oThis.memory.WriteBool(queryTableField.rowNumbers);
				});
			}
			if (null != queryTableField.fillFormulas) {
				this.bs.WriteItem(c_oSer_QueryTableField.FillFormulas, function () {
					oThis.memory.WriteBool(queryTableField.fillFormulas);
				});
			}
			if (null != queryTableField.dataBound) {
				this.bs.WriteItem(c_oSer_QueryTableField.DataBound, function () {
					oThis.memory.WriteBool(queryTableField.dataBound);
				});
			}
			if (null != queryTableField.clipped) {
				this.bs.WriteItem(c_oSer_QueryTableField.Clipped, function () {
					oThis.memory.WriteBool(queryTableField.clipped);
				});
			}
		};
		this.WriteQueryTableDeletedFields = function (queryTableDeletedFields) {
			var oThis = this;
			for (var i = 0, length = queryTableDeletedFields.length; i < length; ++i) {
				this.bs.WriteItem(c_oSer_QueryTableDeletedField.QueryTableDeletedField, function () {
					oThis.WriteQueryTableDeletedField(queryTableDeletedFields[i]);
				});
			}
		};
		this.WriteQueryTableDeletedField = function (queryTableDeletedField) {
			var oThis = this;
			if (null != queryTableDeletedField.name) {
				this.memory.WriteByte(c_oSer_QueryTableDeletedField.Name);
				this.memory.WriteString2(queryTableDeletedField.name);
			}
		};
    }
    /** @constructor */
	function BinarySharedStringsTableWriter(memory, wb, bsw, initSaveManager)
    {
        this.memory = memory;
		this.wb = wb;
        this.bs = new BinaryCommonWriter(this.memory);
		this.bsw = bsw;
        this.InitSaveManager = initSaveManager;
        this.Write = function()
        {
            var oThis = this;
            this.bs.WriteItemWithLength(function(){oThis.WriteSharedStringsContent();});
        };
        this.WriteSharedStringsContent = function()
        {
            var oThis = this;
            var oSharedStrings = this.InitSaveManager.oSharedStrings;
            var aSharedStrings = [];
			for (var i in oSharedStrings.strings) {
				if (oSharedStrings.strings.hasOwnProperty(i)){
					var from = i - 0;
					var to = oSharedStrings.strings[i];
					aSharedStrings[to] = this.wb.sharedStrings.get(from);
                }
			}
			for (var i = 0; i < aSharedStrings.length; ++i) {
				this.bs.WriteItem(c_oSerSharedStringTypes.Si, function(){oThis.WriteSi(aSharedStrings[i]);});
			}
        };
        this.WriteSi = function(si)
        {
			var oThis = this;
			if (typeof si === 'string') {
				this.memory.WriteByte(c_oSerSharedStringTypes.Text);
				this.memory.WriteString2(si);
			} else {
				for (var i = 0, length = si.length; i < length; ++i) {
					this.bs.WriteItem(c_oSerSharedStringTypes.Run, function() {oThis.WriteRun(si[i]);});
				}
			}
        };
        this.WriteRun = function(run)
        {
            var oThis = this;
            if(null != run.format)
                this.bs.WriteItem(c_oSerSharedStringTypes.RPr, function(){oThis.bsw.WriteFont(run.format);});
            if(null != run.text)
            {
                this.memory.WriteByte(c_oSerSharedStringTypes.Text);
                this.memory.WriteString2(run.text);
            }
        };
    }

	function StyleWriteMap(action, prepare) {
		this.action = action;
		this.prepare = prepare;
		this.ids = {};
		this.elems = [];
	}

	StyleWriteMap.prototype.add = function(elem) {
		var index = 0;
		if (elem) {
			elem = this.action.call(g_StyleCache, elem);
			index = this.ids[elem.getIndexNumber()];
			if (undefined === index) {
				index = this.elems.length;
				this.ids[elem.getIndexNumber()] = index;
				this.elems.push(this.prepare ? this.prepare(elem) : elem);
			}
		}
		return index;
	};
	function XfForWrite(xf) {
		this.xf = xf;
		this.fontid = 0;
		this.fillid = 0;
		this.borderid = 0;
		this.numid = 0;
		this.XfId = null;
		this.ApplyAlignment = null;
		this.ApplyBorder = null;
		this.ApplyFill = null;
		this.ApplyFont = null;
		this.ApplyNumberFormat = null
	}

	function StylesForWrite() {
		var t = this;
		this.oXfsMap = new StyleWriteMap(g_StyleCache.addXf, function(xf) {
			return t._getElem(xf, null);
		});
		this.oFontMap = new StyleWriteMap(g_StyleCache.addFont);
		this.oFillMap = new StyleWriteMap(g_StyleCache.addFill);
		this.oBorderMap = new StyleWriteMap(g_StyleCache.addBorder);
		this.oNumMap = new StyleWriteMap(g_StyleCache.addNum);
		this.oXfsStylesMap = [];
	}

	StylesForWrite.prototype.init = function() {
		this.oFontMap.add(g_StyleCache.firstFont);
		this.oFillMap.add(g_StyleCache.firstFill);
		this.oFillMap.add(g_StyleCache.secondFill);
		this.oBorderMap.add(g_StyleCache.firstBorder);
		this.oXfsMap.add(g_StyleCache.firstXf);
	};
	StylesForWrite.prototype.add = function(xf) {
		return this.oXfsMap.add(xf);
	};
	StylesForWrite.prototype.addCellStyle = function(style) {
		this.oXfsStylesMap.push(this._getElem(style.xfs, style));
	};
	StylesForWrite.prototype.finalizeCellStyles = function() {
		//XfId это порядковый номер, поэтому сортируем
		this.oXfsStylesMap.sort(function(a, b) {
			return a.XfId - b.XfId;
		});
	};
	StylesForWrite.prototype.getNumIdByFormat = function(num) {
		var numid = null;
		if (null != num.id) {
			numid = num.id;
		} else {
			numid = AscCommonExcel.aStandartNumFormatsId[num.getFormat()];
		}

		if (null == numid) {
			numid = g_nNumsMaxId + this.oNumMap.add(num);
		}
		return numid;
	};
	StylesForWrite.prototype._getElem = function(xf, style) {
		var elem = new XfForWrite(xf);
		elem.fontid = this.oFontMap.add(xf.font);
		elem.fillid = this.oFillMap.add(xf.fill);
		elem.borderid = this.oBorderMap.add(xf.border);
		elem.numid = xf.num ? this.getNumIdByFormat(xf.num) : 0;
		if(null != xf.align) {
			elem.alignMinimized = xf.align.getDif(g_oDefaultFormat.AlignAbs);
		}
		if (!style) {
			elem.ApplyAlignment = null != elem.alignMinimized || null;
			elem.ApplyBorder = 0 != elem.borderid || null;
			elem.ApplyFill = 0 != elem.fillid || null;
			elem.ApplyFont = 0 != elem.fontid || null;
			elem.ApplyNumberFormat = 0 != elem.numid || null;
		} else {
			elem.ApplyAlignment = style.ApplyAlignment;
			elem.ApplyBorder = style.ApplyBorder;
			elem.ApplyFill = style.ApplyFill;
			elem.ApplyFont = style.ApplyFont;
			elem.ApplyNumberFormat = style.ApplyNumberFormat;
			elem.XfId = style.XfId;
		}
		return elem;
	};
    /** @constructor */
	function BinaryStylesTableWriter(memory, wb, initSaveManager)
    {
        this.memory = memory;
        this.bs = new BinaryCommonWriter(this.memory);
        this.wb = wb;
		this.InitSaveManager = initSaveManager;
		this.stylesForWrite = new StylesForWrite();
        this.Write = function()
        {
            var oThis = this;
            this.bs.WriteItemWithLength(function(){oThis.WriteStylesContent();});
        };
        this.WriteStylesContent = function()
        {
            var oThis = this;
            var wb = this.wb;
            //borders
            this.bs.WriteItem(c_oSerStylesTypes.Borders, function(){oThis.WriteBorders();});
            //fills
            this.bs.WriteItem(c_oSerStylesTypes.Fills, function(){oThis.WriteFills();});
            //fonts
            this.bs.WriteItem(c_oSerStylesTypes.Fonts, function(){oThis.WriteFonts();});
            //CellStyleXfs
            this.bs.WriteItem(c_oSerStylesTypes.CellStyleXfs, function(){oThis.WriteCellStyleXfs();});
            //cellxfs
            this.bs.WriteItem(c_oSerStylesTypes.CellXfs, function(){oThis.WriteCellXfs();});

            if (wb) {
                //CellStyles
                this.bs.WriteItem(c_oSerStylesTypes.CellStyles, function(){oThis.WriteCellStyles(wb.CellStyles.CustomStyles);});

                if(null != wb.TableStyles)
                    this.bs.WriteItem(c_oSerStylesTypes.TableStyles, function(){oThis.WriteTableStyles(wb.TableStyles);});
            }
            //Dxfs пишется после TableStyles, потому что Dxfs может пополниться при записи TableStyles
            var dxfs = this.InitSaveManager.getDxfs();
			if(null != dxfs && dxfs.length > 0) {
				this.bs.WriteItem(c_oSerStylesTypes.Dxfs, function(){oThis.WriteDxfs(dxfs);});
            }
            if (wb) {
                var aExtDxfs = [];
                var slicerStyles = this.InitSaveManager.PrepareSlicerStyles(wb.SlicerStyles, aExtDxfs);
                if(aExtDxfs.length > 0) {
                    this.bs.WriteItem(c_oSerStylesTypes.ExtDxfs, function(){oThis.WriteDxfs(aExtDxfs);});
                }
                this.bs.WriteItem(c_oSerStylesTypes.SlicerStyles, function(){oThis.WriteSlicerStyles(slicerStyles);});
            }
            //numfmts пишется в конце потому что они могут пополниться при записи Dxfs
            this.bs.WriteItem(c_oSerStylesTypes.NumFmts, function(){oThis.WriteNumFmts();});

            if (wb && wb.TimelineStyles) {
                this.bs.WriteItem(c_oSerStylesTypes.TimelineStyles, function(){oThis.WriteTimelineStyles(wb.TimelineStyles);});
            }
        };
        this.WriteBorders = function()
        {
            var oThis = this;
			var elems = this.stylesForWrite.oBorderMap.elems;
			for (var i = 0; i < elems.length; ++i) {
				this.bs.WriteItem(c_oSerStylesTypes.Border, function() {oThis.WriteBorder(elems[i])});
            }
        };
        this.WriteBorder = function(border)
        {
            if(null == border)
                return;
            var oThis = this;
            //Bottom
            if(null != border.b)
                this.bs.WriteItem(c_oSerBorderTypes.Bottom, function(){oThis.WriteBorderProp(border.b);});
            //Diagonal
            if(null != border.d)
                this.bs.WriteItem(c_oSerBorderTypes.Diagonal, function(){oThis.WriteBorderProp(border.d);});
            //End
            if(null != border.r)
                this.bs.WriteItem(c_oSerBorderTypes.End, function(){oThis.WriteBorderProp(border.r);});
            //Horizontal
            if(null != border.ih)
                this.bs.WriteItem(c_oSerBorderTypes.Horizontal, function(){oThis.WriteBorderProp(border.ih);});
            //Start
            if(null != border.l)
                this.bs.WriteItem(c_oSerBorderTypes.Start, function(){oThis.WriteBorderProp(border.l);});
            //Top
            if(null != border.t)
                this.bs.WriteItem(c_oSerBorderTypes.Top, function(){oThis.WriteBorderProp(border.t);});
            //Vertical
            if(null != border.iv)
                this.bs.WriteItem(c_oSerBorderTypes.Vertical, function(){oThis.WriteBorderProp(border.iv);});
            //DiagonalDown
            if(border.dd)
                this.bs.WriteItem(c_oSerBorderTypes.DiagonalDown, function(){oThis.memory.WriteBool(border.dd);});
            //DiagonalUp
            if(border.du)
                this.bs.WriteItem(c_oSerBorderTypes.DiagonalUp, function(){oThis.memory.WriteBool(border.du);});
        };
        this.WriteBorderProp = function(borderProp)
        {
            var oThis = this;
            if(null != borderProp.s)
            {
                var nStyle = EBorderStyle.borderstyleNone;
                switch(borderProp.s)
                {
                    case c_oAscBorderStyles.DashDot:			nStyle = EBorderStyle.borderstyleDashDot;break;
                    case c_oAscBorderStyles.DashDotDot:			nStyle = EBorderStyle.borderstyleDashDotDot;break;
                    case c_oAscBorderStyles.Dashed:				nStyle = EBorderStyle.borderstyleDashed;break;
                    case c_oAscBorderStyles.Dotted:				nStyle = EBorderStyle.borderstyleDotted;break;
                    case c_oAscBorderStyles.Double:				nStyle = EBorderStyle.borderstyleDouble;break;
                    case c_oAscBorderStyles.Hair:				nStyle = EBorderStyle.borderstyleHair;break;
                    case c_oAscBorderStyles.Medium:				nStyle = EBorderStyle.borderstyleMedium;break;
                    case c_oAscBorderStyles.MediumDashDot:		nStyle = EBorderStyle.borderstyleMediumDashDot;break;
                    case c_oAscBorderStyles.MediumDashDotDot:	nStyle = EBorderStyle.borderstyleMediumDashDotDot;break;
                    case c_oAscBorderStyles.MediumDashed:		nStyle = EBorderStyle.borderstyleMediumDashed;break;
                    case c_oAscBorderStyles.None:				nStyle = EBorderStyle.borderstyleNone;break;
                    case c_oAscBorderStyles.SlantDashDot:		nStyle = EBorderStyle.borderstyleSlantDashDot;break;
                    case c_oAscBorderStyles.Thick:				nStyle = EBorderStyle.borderstyleThick;break;
                    case c_oAscBorderStyles.Thin:				nStyle = EBorderStyle.borderstyleThin;break;
                }
                this.memory.WriteByte(c_oSerBorderPropTypes.Style);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(nStyle);

                if (EBorderStyle.borderstyleNone !== nStyle) {
                    this.memory.WriteByte(c_oSerBorderPropTypes.Color);
                    this.memory.WriteByte(c_oSerPropLenType.Variable);
                    this.bs.WriteItemWithLength(function(){oThis.bs.WriteColorSpreadsheet(borderProp.c);});
                }
            }
        };
        this.WriteFills = function()
        {
            var oThis = this;
			var elems = this.stylesForWrite.oFillMap.elems;
			for (var i = 0; i < elems.length; ++i) {
				this.bs.WriteItem(c_oSerStylesTypes.Fill, function() {oThis.WriteFill(elems[i]);});
            }
        };
        this.WriteFill = function(fill, fixDxf)
        {
            var oThis = this;
            fill.checkEmptyContent();
            if (fill.patternFill) {
                this.bs.WriteItem(c_oSerFillTypes.Pattern, function(){oThis.WritePatternFill(fill.patternFill, fixDxf);});
            }
            if (fill.gradientFill) {
                this.bs.WriteItem(c_oSerFillTypes.Gradient, function(){oThis.WriteGradientFill(fill.gradientFill);});
            }
        };
        this.WritePatternFill = function(patternFill, fixDxf)
        {
            var oThis = this;
            fixDxf = fixDxf && (AscCommonExcel.c_oAscPatternType.None === patternFill.patternType || AscCommonExcel.c_oAscPatternType.Solid === patternFill.patternType);
            var fgColor = fixDxf ? patternFill.bgColor : patternFill.fgColor;
            var bgColor = fixDxf ? patternFill.fgColor : patternFill.bgColor;
            if (null != patternFill.patternType) {
                this.bs.WriteItem(c_oSerFillTypes.PatternType, function(){oThis.memory.WriteByte(patternFill.patternType);});
            }
            if (null != fgColor) {
                this.bs.WriteItem(c_oSerFillTypes.PatternFgColor, function(){oThis.bs.WriteColorSpreadsheet(fgColor);});
            }
            if (null != bgColor) {
                this.bs.WriteItem(c_oSerFillTypes.PatternBgColor, function(){oThis.bs.WriteColorSpreadsheet(bgColor);});
            }
        };
        this.WriteGradientFill = function(gradientFill)
        {
            var oThis = this;
            if (null != gradientFill.type) {
                this.bs.WriteItem(c_oSerFillTypes.GradientType, function(){oThis.memory.WriteByte(gradientFill.type);});
            }
            if (null != gradientFill.left) {
                this.bs.WriteItem(c_oSerFillTypes.GradientLeft, function(){oThis.memory.WriteDouble2(gradientFill.left);});
            }
            if (null != gradientFill.top) {
                this.bs.WriteItem(c_oSerFillTypes.GradientTop, function(){oThis.memory.WriteDouble2(gradientFill.top);});
            }
            if (null != gradientFill.right) {
                this.bs.WriteItem(c_oSerFillTypes.GradientRight, function(){oThis.memory.WriteDouble2(gradientFill.right);});
            }
            if (null != gradientFill.bottom) {
                this.bs.WriteItem(c_oSerFillTypes.GradientBottom, function(){oThis.memory.WriteDouble2(gradientFill.bottom);});
            }
            if (null != gradientFill.degree) {
                this.bs.WriteItem(c_oSerFillTypes.GradientDegree, function(){oThis.memory.WriteDouble2(gradientFill.degree);});
            }
            for (var i = 0; i < gradientFill.stop.length; ++i) {
                this.bs.WriteItem(c_oSerFillTypes.GradientStop, function(){oThis.WriteGradientFillStop(gradientFill.stop[i]);});
            }
        };
        this.WriteGradientFillStop = function(gradientStop)
        {
            var oThis = this;
            if (null != gradientStop.position) {
                this.bs.WriteItem(c_oSerFillTypes.GradientStopPosition, function(){oThis.memory.WriteDouble2(gradientStop.position);});
            }
            if (null != gradientStop.color) {
                this.bs.WriteItem(c_oSerFillTypes.GradientStopColor, function(){oThis.bs.WriteColorSpreadsheet(gradientStop.color);});
            }
        };
        this.WriteFonts = function()
        {
            var oThis = this;
			var elems = this.stylesForWrite.oFontMap.elems;
			for (var i = 0; i < elems.length; ++i) {
				this.bs.WriteItem(c_oSerStylesTypes.Font, function() {oThis.WriteFont(elems[i]);});
            }
        };
        this.WriteFont = function(font)
        {
            var oThis = this;
            if(null != font.b)
            {
                this.memory.WriteByte(c_oSerFontTypes.Bold);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(font.b);
            }
            if(null != font.c)
            {
                this.memory.WriteByte(c_oSerFontTypes.Color);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.bs.WriteItemWithLength(function(){oThis.bs.WriteColorSpreadsheet(font.c);});
            }
            if(null != font.i)
            {
                this.memory.WriteByte(c_oSerFontTypes.Italic);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(font.i);
            }
            if(null != font.fn)
            {
                this.memory.WriteByte(c_oSerFontTypes.RFont);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(font.fn);
            }
            if(null != font.scheme)
            {
                this.memory.WriteByte(c_oSerFontTypes.Scheme);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(font.scheme);
            }
            if(null != font.s)
            {
                this.memory.WriteByte(c_oSerFontTypes.Strike);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(font.s);
            }
            if(null != font.fs)
            {
                this.memory.WriteByte(c_oSerFontTypes.Sz);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                //tood write double
                this.memory.WriteDouble2(font.fs);
            }
            if(null != font.u)
            {
                this.memory.WriteByte(c_oSerFontTypes.Underline);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(font.u);
            }
            if(null != font.va)
            {
                this.memory.WriteByte(c_oSerFontTypes.VertAlign);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(font.va);
            }
        };
        this.WriteNumFmts = function()
        {
            var oThis = this;
			var elems = this.stylesForWrite.oNumMap.elems;
			for (var i = 0; i < elems.length; ++i) {
				this.bs.WriteItem(c_oSerStylesTypes.NumFmt, function() {oThis.WriteNum(g_nNumsMaxId + i, elems[i].getFormat());});
            }
        };
        this.WriteNum = function(id, format)
        {
            if(null != format)
            {
                this.memory.WriteByte(c_oSerNumFmtTypes.FormatCode);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(format);
            }
            if(null != id)
            {
                this.memory.WriteByte(c_oSerNumFmtTypes.NumFmtId);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(id);
            }
        };
        this.WriteCellStyleXfs = function()
        {
            var oThis = this;
			var elems = this.stylesForWrite.oXfsStylesMap;
			for (var i = 0; i < elems.length; ++i) {
				this.bs.WriteItem(c_oSerStylesTypes.Xfs, function() {oThis.WriteXfs(elems[i], true);});
            }
        };
        this.WriteCellXfs = function()
        {
            var oThis = this;
			var elems = this.stylesForWrite.oXfsMap.elems;
			for (var i = 0; i < elems.length; ++i) {
				this.bs.WriteItem(c_oSerStylesTypes.Xfs, function() {oThis.WriteXfs(elems[i]);});
            }
        };
        this.WriteXfs = function(xfForWrite, isCellStyle)
        {
            var oThis = this;
            var xf = xfForWrite.xf;
			if(null != xfForWrite.ApplyBorder)
			{
				this.memory.WriteByte(c_oSerXfsTypes.ApplyBorder);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(xfForWrite.ApplyBorder);
			}
            if(null != xfForWrite.borderid)
            {
                this.memory.WriteByte(c_oSerXfsTypes.BorderId);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(xfForWrite.borderid);
            }
			if(null != xfForWrite.ApplyFill)
			{
				this.memory.WriteByte(c_oSerXfsTypes.ApplyFill);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(xfForWrite.ApplyFill);
			}
            if(null != xfForWrite.fillid)
            {
                this.memory.WriteByte(c_oSerXfsTypes.FillId);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(xfForWrite.fillid);
            }
			if(null != xfForWrite.ApplyFont)
			{
				this.memory.WriteByte(c_oSerXfsTypes.ApplyFont);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(xfForWrite.ApplyFont);
			}
            if(null != xfForWrite.fontid)
            {
                this.memory.WriteByte(c_oSerXfsTypes.FontId);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(xfForWrite.fontid);
            }
			if(null != xfForWrite.ApplyNumberFormat)
			{
				this.memory.WriteByte(c_oSerXfsTypes.ApplyNumberFormat);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(xfForWrite.ApplyNumberFormat);
			}
            if(null != xfForWrite.numid)
            {
                this.memory.WriteByte(c_oSerXfsTypes.NumFmtId);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(xfForWrite.numid);
            }
			if(null != xfForWrite.ApplyAlignment)
			{
				this.memory.WriteByte(c_oSerXfsTypes.ApplyAlignment);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(xfForWrite.ApplyAlignment);
			}
			if(null != xfForWrite.alignMinimized)
			{
				this.memory.WriteByte(c_oSerXfsTypes.Aligment);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.bs.WriteItemWithLength(function(){oThis.WriteAlign(xfForWrite.alignMinimized);});
			}

			if (xf) {
				if(null != xf.QuotePrefix)
				{
					this.memory.WriteByte(c_oSerXfsTypes.QuotePrefix);
					this.memory.WriteByte(c_oSerPropLenType.Byte);
					this.memory.WriteBool(xf.QuotePrefix);
				}
				if(null != xf.PivotButton)
				{
					this.memory.WriteByte(c_oSerXfsTypes.PivotButton);
					this.memory.WriteByte(c_oSerPropLenType.Byte);
					this.memory.WriteBool(xf.PivotButton);
				}
				if(!isCellStyle && null != xf.XfId)
				{
					this.memory.WriteByte(c_oSerXfsTypes.XfId);
					this.memory.WriteByte(c_oSerPropLenType.Long);
					this.memory.WriteLong(xf.XfId);
				}
                if(null != xf.applyProtection)
                {
                    this.memory.WriteByte(c_oSerXfsTypes.ApplyProtection);
                    this.memory.WriteByte(c_oSerPropLenType.Byte);
                    this.memory.WriteBool(xf.applyProtection);
                }
				if (null != xf.locked || null != xf.hidden)
				{
					this.memory.WriteByte(c_oSerXfsTypes.Protection);
					this.memory.WriteByte(c_oSerPropLenType.Variable);
					this.bs.WriteItemWithLength(function(){oThis.WriteProtection(xf);});
                }
            }
        };
		this.WriteProtection = function(xf)
		{
			if(null != xf.hidden)
			{
				this.memory.WriteByte(c_oSerProtectionTypes.Hidden);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(xf.hidden);
			}
			if(null != xf.locked)
			{
				this.memory.WriteByte(c_oSerProtectionTypes.Locked);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(xf.locked);
			}
		};
        this.WriteAlign = function(align)
        {
            if(null != align.hor)
            {
                var ha = 4;
                switch (align.hor) {
                    case AscCommon.align_Center :ha = 0;break;
                    case AscCommon.align_Justify :ha = 5;break;
                    case AscCommon.align_Left :ha = 6;break;
                    case AscCommon.align_Right :ha = 7;break;
                    case AscCommon.align_CenterContinuous :ha = 8;break;
                }
                this.memory.WriteByte(Asc.c_oSerAligmentTypes.Horizontal);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(ha);
            }
            if(null != align.indent)
            {
                this.memory.WriteByte(Asc.c_oSerAligmentTypes.Indent);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(align.indent);
            }
            if(null != align.ReadingOrder)
            {
                this.memory.WriteByte(Asc.c_oSerAligmentTypes.ReadingOrder);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(align.ReadingOrder);
            }
            if(null != align.RelativeIndent)
            {
                this.memory.WriteByte(Asc.c_oSerAligmentTypes.RelativeIndent);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(align.RelativeIndent);
            }
            if(null != align.shrink)
            {
                this.memory.WriteByte(Asc.c_oSerAligmentTypes.ShrinkToFit);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(align.shrink);
            }
            if(null != align.angle)
            {
                this.memory.WriteByte(Asc.c_oSerAligmentTypes.TextRotation);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(align.angle);
            }
            if(null != align.ver)
            {
                this.memory.WriteByte(Asc.c_oSerAligmentTypes.Vertical);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(align.ver);
            }
            if(null != align.wrap)
            {
                this.memory.WriteByte(Asc.c_oSerAligmentTypes.WrapText);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(align.wrap);
            }
        };
		this.WriteDxfs = function(Dxfs)
        {
            var oThis = this;
            for(var i = 0, length = Dxfs.length; i < length; ++i)
				this.bs.WriteItem(c_oSerStylesTypes.Dxf, function(){oThis.WriteDxf(Dxfs[i]);});
        };
		this.WriteDxf = function(Dxf)
        {
            var oThis = this;
            if(null != Dxf.align)
                this.bs.WriteItem(c_oSer_Dxf.Alignment, function(){oThis.WriteAlign(Dxf.align);});
            if(null != Dxf.border)
                this.bs.WriteItem(c_oSer_Dxf.Border, function(){oThis.WriteBorder(Dxf.border);});
            if(null != Dxf.fill)
                this.bs.WriteItem(c_oSer_Dxf.Fill, function(){oThis.WriteFill(Dxf.fill, true);});
            if(null != Dxf.font)
                this.bs.WriteItem(c_oSer_Dxf.Font, function(){oThis.WriteFont(Dxf.font);});
			if(null != Dxf.num)
            {
				var numId = this.stylesForWrite.getNumIdByFormat(Dxf.num);
                if(null != numId)
                    this.bs.WriteItem(c_oSer_Dxf.NumFmt, function(){oThis.WriteNum(numId, Dxf.num.getFormat());});
            }
        };
        this.WriteCellStyles = function (cellStyles) {
            var oThis = this;
            for(var i = 0, length = cellStyles.length; i < length; ++i)
            {
                var style = cellStyles[i];
                this.bs.WriteItem(c_oSerStylesTypes.CellStyle, function(){oThis.WriteCellStyle(style);});
            }
        };
        this.WriteCellStyle = function (oCellStyle) {
            var oThis = this;
            if (null != oCellStyle.BuiltinId)
                this.bs.WriteItem(c_oSer_CellStyle.BuiltinId, function(){oThis.memory.WriteLong(oCellStyle.BuiltinId);});
            if (null != oCellStyle.CustomBuiltin)
                this.bs.WriteItem(c_oSer_CellStyle.CustomBuiltin, function(){oThis.memory.WriteBool(oCellStyle.CustomBuiltin);});
            if (null != oCellStyle.Hidden)
                this.bs.WriteItem(c_oSer_CellStyle.Hidden, function(){oThis.memory.WriteBool(oCellStyle.Hidden);});
            if (null != oCellStyle.ILevel)
                this.bs.WriteItem(c_oSer_CellStyle.ILevel, function(){oThis.memory.WriteLong(oCellStyle.ILevel);});
            if (null != oCellStyle.Name) {
                this.memory.WriteByte(c_oSer_CellStyle.Name);
                this.memory.WriteString2(oCellStyle.Name);
            }
            if (null != oCellStyle.XfId)
                this.bs.WriteItem(c_oSer_CellStyle.XfId, function(){oThis.memory.WriteLong(oCellStyle.XfId);});
        };
        this.WriteSlicerStyles = function(slicerStyles)
        {
            var old = new AscCommon.CMemory(true);
            pptx_content_writer.BinaryFileWriter.ExportToMemory(old);
            pptx_content_writer.BinaryFileWriter.ImportFromMemory(this.memory);
            pptx_content_writer.BinaryFileWriter.WriteRecord4(0, slicerStyles);
            pptx_content_writer.BinaryFileWriter.ExportToMemory(this.memory);
            pptx_content_writer.BinaryFileWriter.ImportFromMemory(old);
        };
        this.WriteTableStyles = function(tableStyles)
        {
            var oThis = this;
            if(null != tableStyles.DefaultTableStyle)
            {
                this.memory.WriteByte(c_oSer_TableStyles.DefaultTableStyle);
                this.memory.WriteString2(tableStyles.DefaultTableStyle);
            }
            if(null != tableStyles.DefaultPivotStyle)
            {
                this.memory.WriteByte(c_oSer_TableStyles.DefaultPivotStyle);
                this.memory.WriteString2(tableStyles.DefaultPivotStyle);
            }
            var bEmptyCustom = true;
            for(var i in tableStyles.CustomStyles)
            {
                bEmptyCustom = false;
                break;
            }
            if(false == bEmptyCustom)
            {
                this.bs.WriteItem(c_oSer_TableStyles.TableStyles, function(){oThis.WriteTableCustomStyles(tableStyles.CustomStyles);});
            }
        };
        this.WriteTableCustomStyles = function(customStyles)
        {
            var oThis = this;
            for(var i in customStyles)
            {
                var style = customStyles[i];
                this.bs.WriteItem(c_oSer_TableStyles.TableStyle, function(){oThis.WriteTableCustomStyle(style);});
            }
        };
        this.WriteTableCustomStyle = function(customStyle)
        {
            var oThis = this;
            if(null != customStyle.name)
            {
                this.memory.WriteByte(c_oSer_TableStyle.Name);
                this.memory.WriteString2(customStyle.name);
            }
            if(false === customStyle.pivot)
                this.bs.WriteItem(c_oSer_TableStyle.Pivot, function(){oThis.memory.WriteBool(customStyle.pivot);});
            if(false === customStyle.table)
                this.bs.WriteItem(c_oSer_TableStyle.Table, function(){oThis.memory.WriteBool(customStyle.table);});

            this.bs.WriteItem(c_oSer_TableStyle.Elements, function(){oThis.WriteTableCustomStyleElements(customStyle);});
        };
        this.WriteTableCustomStyleElements = function(customStyle)
        {
            var oThis = this;
            this.InitSaveManager.WriteTableCustomStyleElements(customStyle, function (type, elem) {
                oThis.bs.WriteItem(c_oSer_TableStyle.Element, function(){oThis.WriteTableCustomStyleElement(type, elem);});
            });
        };
        this.WriteTableCustomStyleElement = function(type, customElement)
        {
            if(null != type)
            {
                this.memory.WriteByte(c_oSer_TableStyleElement.Type);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(type);
            }
            if(null != customElement.size)
            {
                this.memory.WriteByte(c_oSer_TableStyleElement.Size);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(customElement.size);
            }
            var dxfs = this.InitSaveManager.getDxfs();
            if(null != customElement.dxf && null != dxfs)
            {
                this.memory.WriteByte(c_oSer_TableStyleElement.DxfId);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(dxfs.length);
                dxfs.push(customElement.dxf);
            }
        };
        this.WriteTimelineStyles = function (oTimelineStyles) {
            if (!oTimelineStyles) {
                return;
            }

            var oThis = this;
            if (oTimelineStyles.defaultTimelineStyle != null) {
                this.bs.WriteItem(c_oSer_TimelineStyles.DefaultTimelineStyle, function () {
                    oThis.memory.WriteString3(oTimelineStyles.defaultTimelineStyle);
                });
            }
            if (oTimelineStyles.timelineStyles && oTimelineStyles.timelineStyles.length) {
                for (let i = 0; i < oTimelineStyles.timelineStyles.length; ++i) {
                    this.bs.WriteItem(c_oSer_TimelineStyles.DefaultTimelineStyle, function () {
                        oThis.WriteTimelineStyle(oTimelineStyles.timelineStyles[i]);
                    });
                }
            }
        };
        this.WriteTimelineStyle = function (oTimelineStyle) {
            if (!oTimelineStyle) {
                return;
            }

            var oThis = this;
            if (oTimelineStyle.name) {
                this.bs.WriteItem(c_oSer_TimelineStyles.Name, function () {
                    oThis.memory.WriteString3(oTimelineStyle.name);
                });

            }

            if (oTimelineStyle.timelineStyleElements && oTimelineStyle.timelineStyleElements.length) {
                for (let i = 0; i < oTimelineStyle.timelineStyleElements.length; ++i) {
                    this.bs.WriteItem(c_oSer_TimelineStyles.DefaultTimelineStyle, function () {
                        oThis.WriteTimelineStyleElement(oTimelineStyle.timelineStyleElements[i]);
                    });
                }
            }
        };
        this.WriteTimelineStyleElement = function (oTimelineStyleElement) {
            if (!oTimelineStyleElement) {
                return;
            }

            if (oTimelineStyleElement.type) {
                this.memory.WriteByte(c_oSer_TimelineStyles.TimelineStyleElementType);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(oTimelineStyleElement.type);
            }
            if (oTimelineStyleElement.DxfId) {
                this.memory.WriteByte(c_oSer_TimelineStyles.TimelineStyleElementDxfId);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oTimelineStyleElement.DxfId);
            }
        }
    }

    function BinaryWorkbookTableWriter(memory, wb, oBinaryWorksheetsTableWriter, isCopyPaste, initSaveManager/*, tableIds, sheetIds*/)
    {
        this.memory = memory;
        this.bs = new BinaryCommonWriter(this.memory);
        this.wb = wb;
        this.oBinaryWorksheetsTableWriter = oBinaryWorksheetsTableWriter;
        this.isCopyPaste = isCopyPaste;
        this.InitSaveManager = initSaveManager;
        //this.tableIds = tableIds;
        //this.sheetIds = sheetIds;
        this.Write = function()
        {
            var oThis = this;
            this.bs.WriteItemWithLength(function(){oThis.WriteWorkbookContent();});
        };
        /*this.PrepareTableIds = function()
        {
            var oThis = this;
            var index = 1;
            this.oBinaryWorksheetsTableWriter.wb.forEach(function(ws) {
                for (var i = 0; i < ws.TableParts.length; ++i) {
                    oThis.tableIds[ws.TableParts[i].DisplayName] = {id: index++, table: ws.TableParts[i]}
                }
            }, this.oBinaryWorksheetsTableWriter.isCopyPaste);
        };
        this.PrepareSheetIds = function()
        {
            var oThis = this;
            var index = 1;
            this.oBinaryWorksheetsTableWriter.wb.forEach(function(ws) {
                oThis.sheetIds[ws.getId()] = index++;
            }, this.oBinaryWorksheetsTableWriter.isCopyPaste);
        };*/
        this.WriteWorkbookContent = function()
        {
            var oThis = this;
            //WorkbookPr
            this.bs.WriteItem(c_oSerWorkbookTypes.WorkbookPr, function(){oThis.WriteWorkbookPr();});

            //BookViews
            this.bs.WriteItem(c_oSerWorkbookTypes.BookViews, function(){oThis.WriteBookViews();});

            //DefinedNames
            this.bs.WriteItem(c_oSerWorkbookTypes.DefinedNames, function(){oThis.WriteDefinedNames();});

			this.bs.WriteItem(c_oSerWorkbookTypes.CalcPr, function(){oThis.WriteCalcPr(oThis.wb.calcPr);});

			//PivotCaches
			let pivotCaches = {};
			let pivotCacheIndex = this.wb.preparePivotForSerialization(pivotCaches, oThis.isCopyPaste);
			if (pivotCacheIndex > 0) {
				this.bs.WriteItem(c_oSerWorkbookTypes.PivotCaches, function () {oThis.WritePivotCaches(pivotCaches);});
			}
			//slicerCaches
            //for copy/paste write string name table/column ?
            /*if (!this.oBinaryWorksheetsTableWriter.isCopyPaste) {
                this.PrepareTableIds();
            }
            this.PrepareSheetIds();*/

            /*var slicerCacheIndex = 0;
            var slicerCaches = {};
            var slicerCacheExtIndex = 0;
            var slicerCachesExt = {};
            this.oBinaryWorksheetsTableWriter.wb.forEach(function(ws) {
                for (var i = 0; i < ws.aSlicers.length; ++i) {
                    var slicerCache = ws.aSlicers[i].getSlicerCache();
                    if (slicerCache) {
                        if (ws.aSlicers[i].isExt()) {
                            slicerCachesExt[slicerCache.name] = slicerCache;
                            slicerCacheExtIndex++;
                        } else {
                            slicerCaches[slicerCache.name] = slicerCache;
                            slicerCacheIndex++;
                        }
                    }
                }
            }, this.oBinaryWorksheetsTableWriter.isCopyPaste);*/

            var slicerCaches = this.InitSaveManager.getSlicersCache();
            var slicerCachesExt = this.InitSaveManager.getSlicersCache(true);
            if (slicerCaches) {
                this.bs.WriteItem(c_oSerWorkbookTypes.SlicerCaches, function () {oThis.WriteSlicerCaches(slicerCaches/*, oThis.tableIds, oThis.sheetIds*/);});
            }
            if (slicerCachesExt) {
                this.bs.WriteItem(c_oSerWorkbookTypes.SlicerCachesExt, function () {oThis.WriteSlicerCaches(slicerCachesExt/*, oThis.tableIds, oThis.sheetIds*/);});
            }
			if (this.wb.externalReferences.length > 0) {

            /*<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
                    <externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">
                    <sheetNames>
                    <sheetName val="Sheet1"/>
                    </sheetNames>
                    <sheetDataSet>
                    <sheetData sheetId="0" refreshError="1"/>
                    </sheetDataSet>
                    </externalBook>
                    <extLst>
                    <ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E66}">
                    <externalReference data="testData"/>
                    </ext>
                    </extLst>
                    </externalLink>*/

				this.bs.WriteItem(c_oSerWorkbookTypes.ExternalReferences, function() {oThis.WriteExternalReferences();});
			}
			if (!this.isCopyPaste) {
                if (this.wb.oApi.vbaProject) {
                    this.bs.WriteItem(c_oSerWorkbookTypes.VbaProject, function() {
                        var old = new AscCommon.CMemory(true);
                        pptx_content_writer.BinaryFileWriter.ExportToMemory(old);
                        pptx_content_writer.BinaryFileWriter.ImportFromMemory(oThis.memory);
                        pptx_content_writer.BinaryFileWriter.WriteRecord4(0, oThis.wb.oApi.vbaProject);
                        pptx_content_writer.BinaryFileWriter.ExportToMemory(oThis.memory);
                        pptx_content_writer.BinaryFileWriter.ImportFromMemory(old);
                    });
                }
				var macros = this.wb.oApi.macros.GetData();
                let customFunctions = this.wb.oApi["pluginMethod_GetCustomFunctions"] && this.wb.oApi["pluginMethod_GetCustomFunctions"]();
                if (customFunctions) {
                    customFunctions = AscCommonExcel.mergeCustomFunctions(customFunctions, true);
                }
                if (customFunctions) {
                    if (macros) {
                        let _macros = AscCommonExcel.safeJsonParse(macros);
                        if (_macros) {
                            _macros["customFunctions"] = customFunctions;
                            macros = JSON.stringify(_macros);
                        }
                    } else {
                        macros = {"customFunctions": customFunctions}
                        macros = JSON.stringify(macros);
                    }
                }
				if (macros) {
					this.bs.WriteItem(c_oSerWorkbookTypes.JsaProject, function() {oThis.memory.WriteXmlString(macros);});
				}
                if (this.wb.aComments.length > 0) {
                    this.bs.WriteItem(c_oSerWorkbookTypes.Comments, function() {oThis.WriteComments(oThis.wb.aComments);});
                }
                //TODO при чтении на клиенте - здесь строка, не пишем пока
				if (this.wb.connections && Array.isArray(oThis.wb.connections)) {
					this.bs.WriteItem(c_oSerWorkbookTypes.Connections, function() {oThis.memory.WriteBuffer(oThis.wb.connections, 0, oThis.wb.connections.length)});
				}
			}
			if (this.wb.workbookProtection) {
				this.bs.WriteItem(c_oSerWorkbookTypes.WorkbookProtection, function(){oThis.WriteWorkbookProtection(oThis.wb.workbookProtection);});
            }
			//FileSharing
			if (this.wb.fileSharing) {
				this.bs.WriteItem(c_oSerWorkbookTypes.FileSharing, function(){oThis.WriteFileSharing(oThis.wb.fileSharing);});
			}
			if (this.wb.oleSize) {
				var sRange = this.wb.oleSize.getName();
				this.bs.WriteItem(c_oSerWorkbookTypes.OleSize, function () {oThis.memory.WriteString3(sRange)});
			}

            if (this.wb.timelineCaches) {
                this.bs.WriteItem(c_oSerWorkbookTypes.TimelineCaches, function () {oThis.WriteTimelineCaches(oThis.wb.timelineCaches);});
            }

			if (this.wb.metadata) {
				this.bs.WriteItem(c_oSerWorkbookTypes.Metadata, function () {oThis.WriteMetadata(oThis.wb.metadata);});
			}
            var xmlMaps =  this.wb.xmlMaps;
            if (xmlMaps) {
                let stream = pptx_content_writer.BinaryFileWriter;
                for (let i = 0; i < xmlMaps.length; i++) {
                    this.bs.WriteItem(c_oSerWorkbookTypes.XmlMap, function() {
                        var old = new AscCommon.CMemory(true);
                        stream.ExportToMemory(old);
                        stream.ImportFromMemory(oThis.memory);

                        stream.StartRecord(0);
                        stream.WriteRecord2(0, stream, function(){
                            xmlMaps[i].toPPTY(stream);
                        });
                        stream.EndRecord();

                        stream.ExportToMemory(oThis.memory);
                        stream.ImportFromMemory(old);
                    });
                }
            }

        };
        this.WriteWorkbookPr = function()
        {
            let oWorkbookPr = this.wb.workbookPr;
            if (null != oWorkbookPr) {
                if (null != oWorkbookPr.Date1904) {
                    this.memory.WriteByte(c_oSerWorkbookPrTypes.Date1904);
                    this.memory.WriteByte(c_oSerPropLenType.Byte);
                    this.memory.WriteBool(oWorkbookPr.Date1904);
                }
                if (null != oWorkbookPr.DateCompatibility) {
                    this.memory.WriteByte(c_oSerWorkbookPrTypes.DateCompatibility);
                    this.memory.WriteByte(c_oSerPropLenType.Byte);
                    this.memory.WriteBool(oWorkbookPr.DateCompatibility);
                }
                if (null != oWorkbookPr.HidePivotFieldList) {
                    this.memory.WriteByte(c_oSerWorkbookPrTypes.HidePivotFieldList);
                    this.memory.WriteByte(c_oSerPropLenType.Byte);
                    this.memory.WriteBool(oWorkbookPr.HidePivotFieldList);
                }
                if (null != oWorkbookPr.ShowPivotChartFilter) {
                    this.memory.WriteByte(c_oSerWorkbookPrTypes.ShowPivotChartFilter);
                    this.memory.WriteByte(c_oSerPropLenType.Byte);
                    this.memory.WriteBool(oWorkbookPr.ShowPivotChartFilter);
                }
                if (null != oWorkbookPr.UpdateLinks) {
                    this.memory.WriteByte(c_oSerWorkbookPrTypes.UpdateLinks);
                    this.memory.WriteByte(c_oSerPropLenType.Byte);
                    this.memory.WriteByte(oWorkbookPr.UpdateLinks);
                }
			}
        };
        this.WriteBookViews = function()
        {
            var oThis = this;
            this.bs.WriteItem(c_oSerWorkbookTypes.WorkbookView, function(){oThis.WriteWorkbookView();});
        };
        this.WriteWorkbookView = function () {
            if (null != this.wb.nActive) {
                this.memory.WriteByte(c_oSerWorkbookViewTypes.ActiveTab);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(this.wb.nActive);
            }
            if (null != this.wb.showVerticalScroll) {
                this.memory.WriteByte(c_oSerWorkbookViewTypes.ShowVerticalScroll);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(this.wb.showVerticalScroll);
            }
            if (null != this.wb.showHorizontalScroll) {
                this.memory.WriteByte(c_oSerWorkbookViewTypes.ShowHorizontalScroll);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(this.wb.showHorizontalScroll);
            }
        };
        this.WriteDefinedNames = function()
        {
            var oThis = this;
            if (this.InitSaveManager.defNameList) {
                var defNames = this.InitSaveManager.defNameList;
                for (var i = 0; i < defNames.length; i++) {
                    this.bs.WriteItem(c_oSerWorkbookTypes.DefinedName, function(){oThis.WriteDefinedName(defNames[i]);});
                }
            }
        };
        this.WriteDefinedName = function(oDefinedName, LocalSheetId)
        {
            var oThis = this;
            if (null != oDefinedName.Name)
            {
                this.memory.WriteByte(c_oSerDefinedNameTypes.Name);
                this.memory.WriteString2(oDefinedName.Name);
            }
            if (null != oDefinedName.Ref)
            {
                this.memory.WriteByte(c_oSerDefinedNameTypes.Ref);
                this.memory.WriteString2(oDefinedName.Ref);
            }
            if (null !== oDefinedName.LocalSheetId){
                var _localSheetId = oDefinedName.LocalSheetId;
                if (this.isCopyPaste === false) {
                    //при переносе листов пишем только один лист
                    // соответсвенно именованные диапазоны должны ссылаться на первый лист
                    _localSheetId = 0;
                }
                this.bs.WriteItem(c_oSerDefinedNameTypes.LocalSheetId, function(){oThis.memory.WriteLong(_localSheetId);});
            }
            if (null != oDefinedName.Hidden) {
                this.bs.WriteItem(c_oSerDefinedNameTypes.Hidden, function(){oThis.memory.WriteBool(oDefinedName.Hidden);});
            }
        };
		this.WriteCalcPr = function(calcPr)
		{
			var t = this;
			//calcId Specifies the version of the calculation engine used to calculate values in the workbook
			//do not pretend to be other editors
			// if (null != calcPr.calcId) {
				// this.bs.WriteItem(c_oSerCalcPrTypes.CalcId, function() {t.memory.WriteLong(calcPr.calcId)});
			// }
			if (null != calcPr.calcMode) {
				this.bs.WriteItem(c_oSerCalcPrTypes.CalcMode, function() {t.memory.WriteByte(calcPr.calcMode)});
			}
			if (null != calcPr.fullCalcOnLoad) {
				this.bs.WriteItem(c_oSerCalcPrTypes.FullCalcOnLoad, function() {t.memory.WriteBool(calcPr.fullCalcOnLoad)});
			}
			if (null != calcPr.refMode) {
				this.bs.WriteItem(c_oSerCalcPrTypes.RefMode, function() {t.memory.WriteByte(calcPr.refMode)});
			}
			if (null != calcPr.iterate) {
				this.bs.WriteItem(c_oSerCalcPrTypes.Iterate, function() {t.memory.WriteBool(calcPr.iterate)});
			}
			if (null != calcPr.iterateCount) {
				this.bs.WriteItem(c_oSerCalcPrTypes.IterateCount, function() {t.memory.WriteLong(calcPr.iterateCount)});
			}
			if (null != calcPr.iterateDelta) {
				this.bs.WriteItem(c_oSerCalcPrTypes.IterateDelta, function() {t.memory.WriteDouble2(calcPr.iterateDelta)});
			}
			if (null != calcPr.fullPrecision) {
				this.bs.WriteItem(c_oSerCalcPrTypes.FullPrecision, function() {t.memory.WriteBool(calcPr.fullPrecision)});
			}
			if (null != calcPr.calcCompleted) {
				this.bs.WriteItem(c_oSerCalcPrTypes.CalcCompleted, function() {t.memory.WriteBool(calcPr.calcCompleted)});
			}
			if (null != calcPr.calcOnSave) {
				this.bs.WriteItem(c_oSerCalcPrTypes.CalcOnSave, function() {t.memory.WriteBool(calcPr.calcOnSave)});
			}
			if (null != calcPr.concurrentCalc) {
				this.bs.WriteItem(c_oSerCalcPrTypes.ConcurrentCalc, function() {t.memory.WriteBool(calcPr.concurrentCalc)});
			}
			if (null != calcPr.concurrentManualCount) {
				this.bs.WriteItem(c_oSerCalcPrTypes.ConcurrentManualCount, function() {t.memory.WriteLong(calcPr.concurrentManualCount)});
			}
			if (null != calcPr.forceFullCalc) {
				this.bs.WriteItem(c_oSerCalcPrTypes.ForceFullCalc, function() {t.memory.WriteBool(calcPr.forceFullCalc)});
			}
		};
		this.WritePivotCaches = function(pivotCaches) {
			var oThis = this;
			for (var id in pivotCaches) {
				if (pivotCaches.hasOwnProperty(id)) {
					var elem = pivotCaches[id];
					this.bs.WriteItem(c_oSerWorkbookTypes.PivotCache, function(){oThis.WritePivotCache(elem.id, elem.cache);});
				}
			}
		};
		this.WritePivotCache = function(id, pivotCache) {
			var oThis = this;
			var oldId = pivotCache.id;
			pivotCache.id = "rId1";
			this.bs.WriteItem(c_oSer_PivotTypes.id, function() {
				oThis.memory.WriteLong(id - 0);
			});
			var stylesForWrite = oThis.isCopyPaste ? undefined : oThis.oBinaryWorksheetsTableWriter.stylesForWrite;
			this.bs.WriteItem(c_oSer_PivotTypes.cache, function() {
				pivotCache.toXml(oThis.memory, stylesForWrite);
			});
			if (pivotCache.cacheRecords) {
				this.bs.WriteItem(c_oSer_PivotTypes.record, function() {
					pivotCache.cacheRecords.toXml(oThis.memory);
				});
			}
			pivotCache.id = oldId;
		};
        this.WriteSlicerCaches = function(slicerCaches/*, tableIds, sheetIds*/) {
            var oThis = this;
            var stream = pptx_content_writer.BinaryFileWriter;
            for (var name in slicerCaches) {
                if (slicerCaches.hasOwnProperty(name)) {
                    this.bs.WriteItem(c_oSerWorkbookTypes.SlicerCache, function() {
                        var old = new AscCommon.CMemory(true);
                        stream.ExportToMemory(old);
                        stream.ImportFromMemory(oThis.memory);

                        stream.StartRecord(0);
                        slicerCaches[name].toStream(stream, oThis.InitSaveManager.getTableIds(), oThis.InitSaveManager.getSheetIds(), oThis.isCopyPaste || oThis.isCopyPaste === false);
                        stream.EndRecord();

                        stream.ExportToMemory(oThis.memory);
                        stream.ImportFromMemory(old);
                    });
                }
            }
        };

        this.WriteTimelineCaches = function (timelineCaches) {
            let oThis = this;
            for (let id in timelineCaches) {
                if (timelineCaches.hasOwnProperty(id)) {
                    let elem = timelineCaches[id];
                    this.bs.WriteItem(c_oSerWorkbookTypes.TimelineCache, function () {
                        oThis.WriteTimelineCache(elem);
                    });
                }
            }
        };
        this.WriteTimelineCache = function (oTimelineCache) {
            if (!oTimelineCache) {
                return;
            }

            let oThis = this;
            if (oTimelineCache.name != null) {
                //this.bs.WriteItem(c_oSerWorksheetsTypes.PageMargins, function(){oThis.WritePageMargins(ws.PagePrintOptions.asc_getPageMargins());});

                oThis.memory.WriteByte(c_oSer_TimelineCache.Name);
                oThis.memory.WriteString2(oTimelineCache.name);
            }
            if (oTimelineCache.sourceName != null) {
                oThis.memory.WriteByte(c_oSer_TimelineCache.SourceName);
                oThis.memory.WriteString2(oTimelineCache.sourceName);
            }
            if (oTimelineCache.uid != null) {
                oThis.memory.WriteByte(c_oSer_TimelineCache.Uid);
                oThis.memory.WriteString2(oTimelineCache.uid);
            }
            if (oTimelineCache.pivotTables != null && oTimelineCache.pivotTables.length > 0) {
                this.bs.WriteItem(c_oSer_TimelineCache.PivotTables, function () {
                    oThis.WriteTimelineCachePivotTables(oTimelineCache.pivotTables);
                });
            }
            //TODO PivotFilter!
            if (oTimelineCache.pivotFilter != null) {
                this.bs.WriteItem(c_oSer_TimelineCache.PivotFilter, function () {
                    oThis.WriteTimelinePivotFilter(oTimelineCache.pivotFilter);
                });
            }
            if (oTimelineCache.state != null) {
                this.bs.WriteItem(c_oSer_TimelineCache.State, function () {
                    oThis.WriteTimelineState(oTimelineCache.state);
                });
            }
        };
        this.WriteTimelineState = function (oTimelineState) {
            if (!oTimelineState) {
                return;
            }

            let oThis = this;
            if (oTimelineState.name != null) {
                this.bs.WriteItem(c_oSer_TimelineState.Name, function(){oThis.memory.WriteString3(oTimelineState.name);});
            }
            if (oTimelineState.singleRangeFilterState != null) {
                this.bs.WriteItem(c_oSer_TimelineState.FilterState, function(){oThis.memory.WriteBool(oTimelineState.singleRangeFilterState);});
            }
            if (oTimelineState.pivotCacheId != null) {
                this.bs.WriteItem(c_oSer_TimelineState.PivotCacheId, function(){oThis.memory.WriteLong(oTimelineState.pivotCacheId);});
            }
            if (oTimelineState.minimalRefreshVersion != null) {
                this.bs.WriteItem(c_oSer_TimelineState.MinimalRefreshVersion, function(){oThis.memory.WriteLong(oTimelineState.minimalRefreshVersion);});
            }
            if (oTimelineState.lastRefreshVersion != null) {
                this.bs.WriteItem(c_oSer_TimelineState.LastRefreshVersion, function(){oThis.memory.WriteLong(oTimelineState.lastRefreshVersion);});
            }
            if (oTimelineState.filterType != null) {
                this.bs.WriteItem(c_oSer_TimelineState.FilterType, function(){oThis.memory.WriteString3(oTimelineState.filterType);});
            }
            if (oTimelineState.selection != null) {
                this.bs.WriteItem(c_oSer_TimelineState.Selection, function () {
                    oThis.WriteTimelineRange(oTimelineState.selection);
                });
            }
            if (oTimelineState.bounds != null) {
                this.bs.WriteItem(c_oSer_TimelineState.Bounds, function () {
                    oThis.WriteTimelineRange(oTimelineState.bounds);
                });
            }
        };
        this.WriteTimelineRange = function (oTimelineRange) {
            if (!oTimelineRange) {
                return;
            }

            if (oTimelineRange.startDate != null) {
                this.memory.WriteByte(c_oSer_TimelineRange.StartDate);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(oTimelineRange.startDate);
            }
            if (oTimelineRange.endDate != null) {
                this.memory.WriteByte(c_oSer_TimelineRange.EndDate);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(oTimelineRange.endDate);
            }
        };
        this.WriteTimelinePivotFilter = function (oPivotFilter) {
            if (!oPivotFilter) {
                return;
            }

            let oThis = this;
            if (oPivotFilter.name != null) {
                this.bs.WriteItem(c_oSer_TimelinePivotFilter.Name, function () {
                    oThis.memory.WriteString2(oPivotFilter.name)
                });
            }
            if (oPivotFilter.description != null) {
                this.bs.WriteItem(c_oSer_TimelinePivotFilter.Description, function () {
                    oThis.memory.WriteString2(oPivotFilter.description)
                });
            }
            if (oPivotFilter.UseWholeDay != null) {
                this.bs.WriteItem(c_oSer_TimelinePivotFilter.UseWholeDay, function () {
                    oThis.memory.WriteBool(oPivotFilter.useWholeDay)
                });
            }
            if (oPivotFilter.id != null) {
                this.bs.WriteItem(c_oSer_TimelinePivotFilter.Id, function () {
                    oThis.memory.WriteULong(oPivotFilter.id)
                });
            }
            if (oPivotFilter.fld != null) {
                this.bs.WriteItem(c_oSer_TimelinePivotFilter.Fld, function () {
                    oThis.memory.WriteULong(oPivotFilter.fld)
                });
            }
            if (oPivotFilter.AutoFilter != null) {
                let oBinaryTableWriter = new BinaryTableWriter(this.memory, this.InitSaveManager, false, {});
                this.bs.WriteItem(c_oSer_TimelinePivotFilter.AutoFilter, function () {
                    oBinaryTableWriter.WriteAutoFilter(oPivotFilter.AutoFilter);
                });
            }
        };
        this.WriteTimelineCachePivotTables = function (oPivotTables) {
            if (!oPivotTables || !oPivotTables.length) {
                return;
            }

            let oThis = this;
            for (let i = 0; i < oPivotTables.length; ++i) {
                this.bs.WriteItem(c_oSer_TimelineCache.PivotTable, function () {
                    oThis.WriteTimelineCachePivotTable(oPivotTables[i]);
                });
            }
        };
        this.WriteTimelineCachePivotTable = function (oPivotTable) {
            if (!oPivotTable) {
                return;
            }

            if (oPivotTable.name != null) {
                this.memory.WriteByte(c_oSer_TimelineCachePivotTable.name);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(oPivotTable.name);
            }
            if (oPivotTable.tabId != null) {
                this.memory.WriteByte(c_oSer_TimelineCachePivotTable.TabId);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                //oThis.sheetIds
                this.memory.WriteULong(this.InitSaveManager.sheetIds[oPivotTable.tabId] || 1);
            }
        };

        this.WriteExternalReferences = function() {
			var oThis = this;

            for(var i = 0, length = this.wb.externalReferences.length; i < length; ++i) {
                this.bs.WriteItem( c_oSerWorkbookTypes.ExternalReference, function(){oThis.WriteExternalReference(oThis.wb.externalReferences[i]);});
            }
		};
        this.WriteExternalReference = function(externalReference) {
            var oThis = this;

            if (externalReference.referenceData) {
                 if (externalReference.referenceData["fileKey"]) {
                     oThis.memory.WriteByte(c_oSerWorkbookTypes.ExternalFileId);
                     var fileKey = externalReference.referenceData["fileKey"] + "";
                     oThis.memory.WriteString2(encodeXmlPath(fileKey, true, true));
                 }
                 if (externalReference.referenceData["instanceId"]) {
                     oThis.memory.WriteByte(c_oSerWorkbookTypes.ExternalPortalName);
                     oThis.memory.WriteString2(externalReference.referenceData["instanceId"]);
                 }
            }

            switch (externalReference.Type) {
                case 0:
                    this.bs.WriteItem( c_oSerWorkbookTypes.ExternalBook, function(){
                        oThis.WriteExternalBook(externalReference);
                    });
                    break;
                case 1:
                    this.bs.WriteItem( c_oSerWorkbookTypes.OleLink, function(){
                        oThis.memory.WriteBuffer(externalReference.Buffer, 0, externalReference.Buffer.length);
                    });
                    break;
                case 2:
                    this.bs.WriteItem( c_oSerWorkbookTypes.OleLink, function(){
                        oThis.memory.WriteBuffer(externalReference.Buffer, 0, externalReference.Buffer.length);
                    });
                    break;
            }
        };
		this.WriteExternalBook = function(externalReference) {
			var oThis = this;
			if (null != externalReference.Id) {
				oThis.memory.WriteByte(c_oSer_ExternalLinkTypes.Id);
				oThis.memory.WriteString2(encodeXmlPath(externalReference.Id));
			}
			if (externalReference.SheetNames.length > 0) {
				this.bs.WriteItem(c_oSer_ExternalLinkTypes.SheetNames, function() {
					oThis.WriteExternalSheetNames(externalReference.SheetNames);
				});
			}
			if (externalReference.DefinedNames.length > 0) {
				this.bs.WriteItem(c_oSer_ExternalLinkTypes.DefinedNames, function() {
					oThis.WriteExternalDefinedNames(externalReference.DefinedNames);
				});
			}
			if (externalReference.SheetDataSet.length > 0) {
				this.bs.WriteItem(c_oSer_ExternalLinkTypes.SheetDataSet, function() {
					oThis.WriteExternalSheetDataSet(externalReference.SheetDataSet);
				});
			}
		};
		this.WriteExternalSheetNames = function(sheetNames) {
            var oThis = this;
		    for (var i = 0; i < sheetNames.length; i++) {
				this.memory.WriteByte(c_oSer_ExternalLinkTypes.SheetName);
				this.memory.WriteString2(sheetNames[i]);
			}
		};
		this.WriteExternalDefinedNames = function(definedNames) {
			var oThis = this;
			for (var i = 0; i < definedNames.length; i++) {
				this.bs.WriteItem(c_oSer_ExternalLinkTypes.DefinedName, function() {
					oThis.WriteExternalDefinedName(definedNames[i]);
				});
			}
		};
		this.WriteExternalDefinedName = function(definedName) {
			var oThis = this;
			if (null != definedName.Name) {
                this.bs.WriteItem(c_oSer_ExternalLinkTypes.DefinedNameName, function() {
                    oThis.memory.WriteString3(definedName.Name);
                });
			}
			if (null != definedName.RefersTo) {
                this.bs.WriteItem(c_oSer_ExternalLinkTypes.DefinedNameRefersTo, function() {
                    oThis.memory.WriteString3(definedName.RefersTo);
                });
			}
			if (null != definedName.SheetId) {
				this.bs.WriteItem(c_oSer_ExternalLinkTypes.DefinedNameSheetId, function() {
					oThis.memory.WriteLong(definedName.SheetId);
				});
			}
		};
		this.WriteExternalSheetDataSet = function(sheetDataSet) {
			var oThis = this;
			for (var i = 0; i < sheetDataSet.length; i++) {
				this.bs.WriteItem(c_oSer_ExternalLinkTypes.SheetData, function() {
					oThis.WriteExternalSheetData(sheetDataSet[i]);
				});
			}
		};
		this.WriteExternalSheetData = function(sheetData) {
			var oThis = this;
			if (null != sheetData.SheetId) {
				this.bs.WriteItem(c_oSer_ExternalLinkTypes.SheetDataSheetId, function() {
					oThis.memory.WriteLong(sheetData.SheetId);
				});
			}
			if (null != sheetData.RefreshError) {
				this.bs.WriteItem(c_oSer_ExternalLinkTypes.SheetDataRefreshError, function() {
					oThis.memory.WriteBool(sheetData.RefreshError);
				});
			}
			if (sheetData.Row.length > 0) {
				for (var i = 0; i < sheetData.Row.length; i++) {
					this.bs.WriteItem(c_oSer_ExternalLinkTypes.SheetDataRow, function() {
						oThis.WriteExternalRow(sheetData.Row[i]);
					});
				}
			}
		};
		this.WriteExternalRow = function(row) {
			var oThis = this;
			if (null != row.R) {
				this.bs.WriteItem(c_oSer_ExternalLinkTypes.SheetDataRowR, function() {
					oThis.memory.WriteLong(row.R);
				});
			}
			if (row.Cell.length > 0) {
				for (var i = 0; i < row.Cell.length; i++) {
					this.bs.WriteItem(c_oSer_ExternalLinkTypes.SheetDataRowCell, function() {
						oThis.WriteExternalCell(row.Cell[i]);
					});
				}
			}
		};
		this.WriteExternalCell = function(cell) {
			var oThis = this;
			if (null != cell.Ref) {
				oThis.memory.WriteByte(c_oSer_ExternalLinkTypes.SheetDataRowCellRef);
				oThis.memory.WriteString2(cell.Ref);
			}
			if (null != cell.CellType) {
				this.bs.WriteItem(c_oSer_ExternalLinkTypes.SheetDataRowCellType, function() {
					oThis.memory.WriteByte(cell.CellType);
				});
			}
			if (null != cell.CellValue) {
				oThis.memory.WriteByte(c_oSer_ExternalLinkTypes.SheetDataRowCellValue);
				oThis.memory.WriteString2(cell.CellValue);
			}
			if (null != cell.vm) {
				this.bs.WriteItem(c_oSer_ExternalLinkTypes.ValueMetadata, function () {
					oThis.memory.WriteULong(cell.vm);
				});
			}
		};


		//****write metadata****
		this.WriteMetadata = function (pMetadata) {
			if (!pMetadata) {
				return;
			}

			var oThis = this;
			if (pMetadata.metadataTypes) {
				this.bs.WriteItem(c_oSer_Metadata.MetadataTypes, function () {
					oThis.WriteMetadataTypes(pMetadata.metadataTypes);
				});
			}
			if (pMetadata.metadataTypes) {
				this.bs.WriteItem(c_oSer_Metadata.MetadataStrings, function () {
					oThis.WriteMetadataStrings(pMetadata.metadataTypes);
				});
			}
			if (pMetadata.mdxMetadata) {
				this.bs.WriteItem(c_oSer_Metadata.MdxMetadata, function () {
					oThis.WriteMdxMetadata(pMetadata.mdxMetadata);
				});
			}
			if (pMetadata.cellMetadata) {
				this.bs.WriteItem(c_oSer_Metadata.CellMetadata, function () {
					oThis.WriteMetadataBlocks(pMetadata.cellMetadata);
				});
			}
			if (pMetadata.valueMetadata) {
				this.bs.WriteItem(c_oSer_Metadata.ValueMetadata, function () {
					oThis.WriteMetadataBlocks(pMetadata.valueMetadata);
				});
			}
			if (pMetadata.aFutureMetadata) {
				for (let i = 0; i < pMetadata.aFutureMetadata.length; ++i) {
					this.bs.WriteItem(c_oSer_Metadata.FutureMetadata, function () {
						oThis.WriteFutureMetadata(pMetadata.aFutureMetadata[i]);
					});
				}
			}
		};
		this.WriteMetadataTypes = function (pMetadataTypes) {
			if (!pMetadataTypes) {
				return;
			}

			var oThis = this;
			for (let i = 0; i < pMetadataTypes.length; ++i) {
				this.bs.WriteItem(c_oSer_MetadataType.MetadataType, function () {
					oThis.WriteMetadataType(pMetadataTypes[i]);
				});
			}
		};
		this.WriteMetadataType = function (pMetadataType) {
			if (!pMetadataType) {
				return;
			}

			var oThis = this;
			if (pMetadataType.name != null) {
				this.bs.WriteItem(c_oSer_MetadataType.Name, function () {
					oThis.memory.WriteString3(pMetadataType.name);
				});
			}
			if (pMetadataType.minSupportedVersion != null) {
				this.bs.WriteItem(c_oSer_MetadataType.MinSupportedVersion, function () {
					oThis.memory.WriteLong(pMetadataType.minSupportedVersion);
				});
			}
			if (pMetadataType.ghostRow != null) {
				this.bs.WriteItem(c_oSer_MetadataType.GhostRow, function () {
					oThis.memory.WriteBool(pMetadataType.ghostRow);
				});
			}
			if (pMetadataType.ghostCol != null) {
				this.bs.WriteItem(c_oSer_MetadataType.GhostCol, function () {
					oThis.memory.WriteBool(pMetadataType.ghostCol);
				});
			}
			if (pMetadataType.edit != null) {
				this.bs.WriteItem(c_oSer_MetadataType.Edit, function () {
					oThis.memory.WriteBool(pMetadataType.edit);
				});
			}
			if (pMetadataType.delete != null) {
				this.bs.WriteItem(c_oSer_MetadataType.Delete, function () {
					oThis.memory.WriteBool(pMetadataType.delete);
				});
			}
			if (pMetadataType.copy != null) {
				this.bs.WriteItem(c_oSer_MetadataType.Copy, function () {
					oThis.memory.WriteBool(pMetadataType.copy);
				});
			}
			if (pMetadataType.pasteAll != null) {
				this.bs.WriteItem(c_oSer_MetadataType.PasteAll, function () {
					oThis.memory.WriteBool(pMetadataType.pasteAll);
				});
			}
			if (pMetadataType.pasteFormulas != null) {
				this.bs.WriteItem(c_oSer_MetadataType.PasteFormulas, function () {
					oThis.memory.WriteBool(pMetadataType.pasteFormulas);
				});
			}
			if (pMetadataType.pasteValues != null) {
				this.bs.WriteItem(c_oSer_MetadataType.PasteValues, function () {
					oThis.memory.WriteBool(pMetadataType.pasteValues);
				});
			}
			if (pMetadataType.pasteFormats != null) {
				this.bs.WriteItem(c_oSer_MetadataType.PasteFormats, function () {
					oThis.memory.WriteBool(pMetadataType.pasteFormats);
				});
			}
			if (pMetadataType.pasteComments != null) {
				this.bs.WriteItem(c_oSer_MetadataType.PasteComments, function () {
					oThis.memory.WriteBool(pMetadataType.pasteComments);
				});
			}
			if (pMetadataType.pasteDataValidation != null) {
				this.bs.WriteItem(c_oSer_MetadataType.PasteDataValidation, function () {
					oThis.memory.WriteBool(pMetadataType.pasteDataValidation);
				});
			}
			if (pMetadataType.pasteBorders != null) {
				this.bs.WriteItem(c_oSer_MetadataType.PasteBorders, function () {
					oThis.memory.WriteBool(pMetadataType.pasteBorders);
				});
			}
			if (pMetadataType.pasteColWidths != null) {
				this.bs.WriteItem(c_oSer_MetadataType.PasteColWidths, function () {
					oThis.memory.WriteBool(pMetadataType.pasteColWidths);
				});
			}
			if (pMetadataType.pasteNumberFormats != null) {
				this.bs.WriteItem(c_oSer_MetadataType.PasteNumberFormats, function () {
					oThis.memory.WriteBool(pMetadataType.pasteNumberFormats);
				});
			}
			if (pMetadataType.merge != null) {
				this.bs.WriteItem(c_oSer_MetadataType.Merge, function () {
					oThis.memory.WriteBool(pMetadataType.merge);
				});
			}
			if (pMetadataType.splitFirst != null) {
				this.bs.WriteItem(c_oSer_MetadataType.SplitFirst, function () {
					oThis.memory.WriteBool(pMetadataType.splitFirst);
				});
			}
			if (pMetadataType.SplitAll != null) {
				this.bs.WriteItem(c_oSer_MetadataType.SplitAll, function () {
					oThis.memory.WriteBool(pMetadataType.splitAll);
				});
			}
			if (pMetadataType.rowColShift != null) {
				this.bs.WriteItem(c_oSer_MetadataType.RowColShift, function () {
					oThis.memory.WriteBool(pMetadataType.rowColShift);
				});
			}
			if (pMetadataType.clearAll != null) {
				this.bs.WriteItem(c_oSer_MetadataType.ClearAll, function () {
					oThis.memory.WriteBool(pMetadataType.clearAll);
				});
			}
			if (pMetadataType.clearFormats != null) {
				this.bs.WriteItem(c_oSer_MetadataType.ClearFormats, function () {
					oThis.memory.WriteBool(pMetadataType.clearFormats);
				});
			}
			if (pMetadataType.clearContents != null) {
				this.bs.WriteItem(c_oSer_MetadataType.ClearContents, function () {
					oThis.memory.WriteBool(pMetadataType.clearContents);
				});
			}
			if (pMetadataType.clearComments != null) {
				this.bs.WriteItem(c_oSer_MetadataType.ClearComments, function () {
					oThis.memory.WriteBool(pMetadataType.clearComments);
				});
			}
			if (pMetadataType.assign != null) {
				this.bs.WriteItem(c_oSer_MetadataType.Assign, function () {
					oThis.memory.WriteBool(pMetadataType.assign);
				});
			}
			if (pMetadataType.coerce != null) {
				this.bs.WriteItem(c_oSer_MetadataType.Coerce, function () {
					oThis.memory.WriteBool(pMetadataType.coerce);
				});
			}
			if (pMetadataType.cellMeta != null) {
				this.bs.WriteItem(c_oSer_MetadataType.CellMeta, function () {
					oThis.memory.WriteBool(pMetadataType.cellMeta);
				});
			}
		};
		this.WriteMetadataStrings = function (pMetadataStrings) {
			if (!pMetadataStrings) {
				return;
			}

			var oThis = this;
			for (let i = 0; i < pMetadataStrings.length; ++i) {
				if (pMetadataStrings[i] && pMetadataStrings[i].v) {
					this.bs.WriteItem(c_oSer_MetadataString.MetadataString, function () {
						oThis.memory.WriteString3(pMetadataStrings[i].v);
					});
				}
			}
		};
		this.WriteMdxMetadata = function (pMdxMetadata) {
			if (!pMdxMetadata) {
				return;
			}

			var oThis = this;
			for (let i = 0; i < pMdxMetadata.length; ++i) {
				if (!pMdxMetadata[i]) {
					continue;
				}

				this.bs.WriteItem(c_oSer_MdxMetadata.Mdx, function () {
					oThis.WriteMdx(pMdxMetadata[i]);
				});
			}
		};
		this.WriteMdx = function (pMdx) {
			if (!pMdx) {
				return;
			}

			var oThis = this;
			if (pMdx.n != null) {
				this.bs.WriteItem(c_oSer_MdxMetadata.NameIndex, function () {
					oThis.memory.WriteLong(pMdx.n);
				});
			}
			if (pMdx.f) {
				this.bs.WriteItem(c_oSer_MdxMetadata.FunctionTag, function () {
					oThis.memory.WriteByte(pMdx.f);
				});
			}
			if (pMdx.mdxTuple) {
				this.bs.WriteItem(c_oSer_MdxMetadata.MdxTuple, function () {
					oThis.WriteMdxTuple(pMdx.mdxTuple);
				});
			}
			if (pMdx.mdxSet) {
				this.bs.WriteItem(c_oSer_MdxMetadata.MdxSet, function () {
					oThis.WriteMdxSet(pMdx.mdxSet);
				});
			}
			if (pMdx.mdxKPI) {
				this.bs.WriteItem(c_oSer_MdxMetadata.MdxKPI, function () {
					oThis.WriteMdxKPI(pMdx.mdxKPI);
				});
			}
			if (pMdx.mdxMemeberProp) {
				this.bs.WriteItem(c_oSer_MdxMetadata.MdxMemeberProp, function () {
					oThis.WriteMdxMemeberProp(pMdx.mdxMemeberProp);
				});
			}
		};
		this.WriteMdxTuple = function (pMdxTuple) {
			if (!pMdxTuple) {
				return;
			}

			var oThis = this;
			if (pMdxTuple.c) {
				this.bs.WriteItem(c_oSer_MetadataMdxTuple.IndexCount, function () {
					oThis.memory.WriteLong(pMdxTuple.c);
				});
			}
			if (pMdxTuple.ct) {
				this.bs.WriteItem(c_oSer_MetadataMdxTuple.CultureCurrency, function () {
					oThis.memory.WriteString3(pMdxTuple.ct);
				});
			}
			if (pMdxTuple.si) {
				this.bs.WriteItem(c_oSer_MetadataMdxTuple.StringIndex, function () {
					oThis.memory.WriteLong(pMdxTuple.si);
				});
			}
			if (pMdxTuple.fi) {
				this.bs.WriteItem(c_oSer_MetadataMdxTuple.NumFmtIndex, function () {
					oThis.memory.WriteLong(pMdxTuple.fi);
				});
			}
			if (pMdxTuple.bc) {
				this.bs.WriteItem(c_oSer_MetadataMdxTuple.BackColor, function () {
					oThis.memory.WriteLong(pMdxTuple.bc);
				});
			}
			if (pMdxTuple.fc) {
				this.bs.WriteItem(c_oSer_MetadataMdxTuple.ForeColor, function () {
					oThis.memory.WriteLong(pMdxTuple.fc);
				});
			}
			if (pMdxTuple.i) {
				this.bs.WriteItem(c_oSer_MetadataMdxTuple.Italic, function () {
					oThis.memory.WriteLong(pMdxTuple.i);
				});
			}
			if (pMdxTuple.b) {
				this.bs.WriteItem(c_oSer_MetadataMdxTuple.Bold, function () {
					oThis.memory.WriteBool(pMdxTuple.b);
				});
			}
			if (pMdxTuple.u) {
				this.bs.WriteItem(c_oSer_MetadataMdxTuple.Underline, function () {
					oThis.memory.WriteBool(pMdxTuple.u);
				});
			}
			if (pMdxTuple.st) {
				this.bs.WriteItem(c_oSer_MetadataMdxTuple.Strike, function () {
					oThis.memory.WriteBool(pMdxTuple.st);
				});
			}
			if (pMdxTuple.metadataStringIndexes) {
				for (let i = 0; i < pMdxTuple.metadataStringIndexes.length; ++i) {
					if (!pMdxTuple.metadataStringIndexes[i]) {
						continue;
					}

					this.bs.WriteItem(c_oSer_MetadataMdxTuple.MetadataStringIndex, function () {
						oThis.WriteMetadataStringIndex(pMdxTuple.metadataStringIndexes[i]);
					});
				}
			}
		};
		this.WriteMetadataStringIndex = function (pStringIndex) {
			if (!pStringIndex) {
				return;
			}

			var oThis = this;
			if (pStringIndex.x) {
				this.bs.WriteItem(c_oSer_MetadataStringIndex.IndexValue, function () {
					oThis.memory.WriteLong(pStringIndex.x);
				});
			}
			if (pStringIndex.s) {
				this.bs.WriteItem(c_oSer_MetadataStringIndex.StringIsSet, function () {
					oThis.memory.WriteLong(pStringIndex.s);
				});
			}
		};
		this.WriteMdxSet = function (pMdxSet) {
			if (!pMdxSet) {
				return;
			}

			var oThis = this;
			if (pMdxSet.c) {
				this.bs.WriteItem(c_oSer_MetadataMdxSet.Count, function () {
					oThis.memory.WriteLong(pMdxSet.c);
				});
			}
			if (pMdxSet.ns) {
				this.bs.WriteItem(c_oSer_MetadataMdxSet.Index, function () {
					oThis.memory.WriteLong(pMdxSet.ns);
				});
			}
			if (pMdxSet.o) {
				this.bs.WriteItem(c_oSer_MetadataMdxSet.SortOrder, function () {
					oThis.memory.WriteByte(pMdxSet.o);
				});
			}
			if (pMdxSet.metadataStringIndexes) {
				for (let i = 0; i < pMdxSet.metadataStringIndexes.length; ++i) {
					if (!pMdxSet.metadataStringIndexes[i]) {
						continue;
					}

					this.bs.WriteItem(c_oSer_MetadataMdxSet.MetadataStringIndex, function () {
						oThis.WriteMetadataStringIndex(pMdxSet.metadataStringIndexes[i]);
					});
				}
			}
		};
		this.WriteMdxKPI = function (pMdxKPI) {
			if (!pMdxKPI) {
				return;
			}

			var oThis = this;
			if (pMdxKPI.n) {
				this.bs.WriteItem(c_oSer_MetadataMdxKPI.NameIndex, function () {
					oThis.memory.WriteLong(pMdxKPI.n);
				});
			}
			if (pMdxKPI.np) {
				this.bs.WriteItem(c_oSer_MetadataMdxKPI.Index, function () {
					oThis.memory.WriteLong(pMdxKPI.np);
				});
			}
			if (pMdxKPI.p) {
				this.bs.WriteItem(c_oSer_MetadataMdxKPI.Property, function () {
					oThis.memory.WriteByte(pMdxKPI.p);
				});

			}
		};
		this.WriteMdxMemeberProp = function (pMdxMemeberProp) {
			if (!pMdxMemeberProp) {
				return;
			}

			var oThis = this;
			if (pMdxMemeberProp.n) {
				this.bs.WriteItem(c_oSer_MetadataMemberProperty.NameIndex, function () {
					oThis.memory.WriteLong(pMdxMemeberProp.n);
				});
			}
			if (pMdxMemeberProp.np) {
				this.bs.WriteItem(c_oSer_MetadataMemberProperty.Index, function () {
					oThis.memory.WriteLong(pMdxMemeberProp.np);
				});
			}
		};
		this.WriteMetadataBlocks = function (aMetadataBlocks) {
			if (!aMetadataBlocks) {
				return;
			}

			var oThis = this;
			for (let i = 0; i < aMetadataBlocks.length; ++i) {
				if (!aMetadataBlocks[i]) {
					continue;
				}

				this.bs.WriteItem(c_oSer_MetadataBlock.MetadataBlock, function () {
					oThis.WriteMetadataBlock(aMetadataBlocks[i]);
				});
			}
		};
		this.WriteMetadataBlock = function (metadataBlock) {
			if (!metadataBlock) {
				return;
			}

			var oThis = this;
			this.bs.WriteItem(c_oSer_MetadataBlock.MetadataRecord, function () {
				oThis.WriteMetadataRecord(metadataBlock);
			});
		};
		this.WriteMetadataRecord = function (pMetadataRecord) {
			if (!pMetadataRecord) {
				return;
			}

			var oThis = this;
			if (pMetadataRecord.t) {
				this.bs.WriteItem(c_oSer_MetadataBlock.MetadataRecordType, function () {
					oThis.memory.WriteLong(pMetadataRecord.t);
				});
			}
			if (pMetadataRecord.v) {
				this.bs.WriteItem(c_oSer_MetadataBlock.MetadataRecordValue, function () {
					oThis.memory.WriteLong(pMetadataRecord.v);
				});
			}
		};
		this.WriteFutureMetadataBlock = function (pFutureMetadataBlock) {
			if (!pFutureMetadataBlock || !pFutureMetadataBlock.extLst) {
				return;
			}

			var oThis = this;
			for (let i = 0; i < pFutureMetadataBlock.extLst.length; ++i) {
				if (!pFutureMetadataBlock.extLst[i]) {
					continue;
				}

				if (pFutureMetadataBlock.extLst[i].dynamicArrayProperties) {

					this.bs.WriteItem(c_oSer_FutureMetadataBlock.DynamicArrayProperties, function () {

						if (pFutureMetadataBlock.extLst[i].dynamicArrayProperties.fDynamic) {
							oThis.bs.WriteItem(c_oSer_FutureMetadataBlock.DynamicArray, function () {
								oThis.memory.WriteBool(pFutureMetadataBlock.extLst[i].dynamicArrayProperties.fDynamic);
							});

						}
						if (pFutureMetadataBlock.extLst[i].dynamicArrayProperties.fCollapsed) {
							oThis.bs.WriteItem(c_oSer_FutureMetadataBlock.CollapsedArray, function () {
								oThis.memory.WriteBool(pFutureMetadataBlock.extLst[i].dynamicArrayProperties.fCollapsed);
							});
						}
					});
				}

				if ((pFutureMetadataBlock.extLst[i].richValueBlock) && (pFutureMetadataBlock.extLst[i].richValueBlock.i)) {
					oThis.bs.WriteItem(c_oSer_FutureMetadataBlock.RichValueBlock, function () {
						oThis.memory.WriteLong(pFutureMetadataBlock.extLst[i].richValueBlock.i);
					});
				}
			}
		};
		this.WriteFutureMetadata = function (pFutureMetadata) {
			if (!pFutureMetadata) {
				return;
			}

			var oThis = this;
			if (pFutureMetadata.name) {
				oThis.bs.WriteItem(c_oSer_FutureMetadataBlock.Name, function () {
					oThis.memory.WriteString3(pFutureMetadata.name);
				});
			}
			if (pFutureMetadata.futureMetadataBlocks) {
				for (let i = 0; i < pFutureMetadata.futureMetadataBlocks.length; ++i) {
					if (!pFutureMetadata.futureMetadataBlocks[i]) {
						continue;
					}

					this.bs.WriteItem(c_oSer_FutureMetadataBlock.FutureMetadataBlock, function () {
						oThis.WriteFutureMetadataBlock(pFutureMetadata.futureMetadataBlocks[i]);
					});
				}
			}
		};

        
        this.WriteComments = function(aComments) {
            var t = this;
            for (var i = 0; i < aComments.length; ++i) {
                this.bs.WriteItem( c_oSer_Comments.CommentData, function(){t.oBinaryWorksheetsTableWriter.WriteCommentData(aComments[i]);});
            }
        };
		this.WriteWorkbookProtection = function(workbookProtection)
		{
			if (null != workbookProtection.lockStructure) {
				this.memory.WriteByte(c_oSerWorkbookProtection.LockStructure);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(workbookProtection.lockStructure);
			}
			if (null != workbookProtection.lockWindows) {
				this.memory.WriteByte(c_oSerWorkbookProtection.LockWindows);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(workbookProtection.lockWindows);
			}
			if (null != workbookProtection.lockRevision) {
				this.memory.WriteByte(c_oSerWorkbookProtection.LockRevision);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(workbookProtection.lockRevision);
			}
			if (null != workbookProtection.workbookPassword) {
				this.memory.WriteByte(c_oSerWorkbookProtection.Password);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(workbookProtection.workbookPassword);
			}

		    if (null != workbookProtection.revisionsAlgorithmName) {
				this.memory.WriteByte(c_oSerWorkbookProtection.RevisionsAlgorithmName);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteByte(workbookProtection.revisionsAlgorithmName);
			}
			if (null != workbookProtection.revisionsSpinCount) {
				this.memory.WriteByte(c_oSerWorkbookProtection.RevisionsSpinCount);
				this.memory.WriteByte(c_oSerPropLenType.Long);
				this.memory.WriteLong(workbookProtection.revisionsSpinCount);
			}
			if (null != workbookProtection.revisionsHashValue) {
				this.memory.WriteByte(c_oSerWorkbookProtection.RevisionsHashValue);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(workbookProtection.revisionsHashValue);
			}
			if (null != workbookProtection.revisionsSaltValue) {
				this.memory.WriteByte(c_oSerWorkbookProtection.RevisionsSaltValue);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(workbookProtection.revisionsSaltValue);
			}

			if (null != workbookProtection.workbookAlgorithmName) {
				this.memory.WriteByte(c_oSerWorkbookProtection.WorkbookAlgorithmName);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteByte(workbookProtection.workbookAlgorithmName);
			}
			if (null != workbookProtection.workbookSpinCount) {
				this.memory.WriteByte(c_oSerWorkbookProtection.WorkbookSpinCount);
				this.memory.WriteByte(c_oSerPropLenType.Long);
				this.memory.WriteLong(workbookProtection.workbookSpinCount);
			}
			if (null != workbookProtection.workbookHashValue) {
				this.memory.WriteByte(c_oSerWorkbookProtection.WorkbookHashValue);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(workbookProtection.workbookHashValue);
			}
			if (null != workbookProtection.workbookSaltValue) {
				this.memory.WriteByte(c_oSerWorkbookProtection.WorkbookSaltValue);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(workbookProtection.workbookSaltValue);
			}

		};
		this.WriteFileSharing = function(fileSharing)
		{
			if (null != fileSharing.algorithmName) {
				this.memory.WriteByte(c_oSerFileSharing.AlgorithmName);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteByte(fileSharing.algorithmName);
			}
			if (null != fileSharing.spinCount) {
				this.memory.WriteByte(c_oSerFileSharing.SpinCount);
				this.memory.WriteByte(c_oSerPropLenType.Long);
				this.memory.WriteLong(fileSharing.spinCount);
			}
			if (null != fileSharing.hashValue) {
				this.memory.WriteByte(c_oSerFileSharing.HashValue);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(fileSharing.hashValue);
			}
			if (null != fileSharing.saltValue) {
				this.memory.WriteByte(c_oSerFileSharing.SaltValue);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(fileSharing.saltValue);
			}

			if (null != fileSharing.password) {
				this.memory.WriteByte(c_oSerFileSharing.Password);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(fileSharing.password);
			}
			if (null != fileSharing.userName) {
				this.memory.WriteByte(c_oSerFileSharing.UserName);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(fileSharing.userName);
			}
			if (null != fileSharing.readOnly) {
				this.memory.WriteByte(c_oSerFileSharing.ReadOnly);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(fileSharing.readOnly);
			}
		};
    }
	function BinaryWorksheetsTableWriter(memory, wb, isCopyPaste, bsw, saveThreadedComments, initSaveManager)
    {
        this.memory = memory;
        this.bs = new BinaryCommonWriter(this.memory);
		this.bsw = bsw;
        this.wb = wb;
		this.stylesForWrite = bsw.stylesForWrite;
        this.isCopyPaste = isCopyPaste;
        this.saveThreadedComments = saveThreadedComments;
        this.InitSaveManager = initSaveManager;
        /*this.tableIds = tableIds;
        this.sheetIds = sheetIds;*/
        this._getCrc32FromObjWithProperty = function(val)
        {
            return Asc.crc32(this._getStringFromObjWithProperty(val));
        };
        this._getStringFromObjWithProperty = function(val)
        {
            var sRes = "";
            if(val.getProperties)
            {
                var properties = val.getProperties();
                for(var i in properties)
                {
                    var oCurProp = val.getProperty(properties[i]);
                    if(null != oCurProp && oCurProp.getProperties)
                        sRes += this._getStringFromObjWithProperty(oCurProp);
                    else
                        sRes += oCurProp;
                }
            }
            return sRes;
        };
        this.Write = function()
        {
            var oThis = this;
            this.InitSaveManager._prepeareStyles(this.stylesForWrite);
            this.bs.WriteItemWithLength(function(){oThis.WriteWorksheetsContent();});
        };
        this.WriteWorksheetsContent = function()
        {
			var oThis = this;
			var hasColorFilter = false;
			this.wb.forEach(function (ws, index) {
				oThis.bs.WriteItem(c_oSerWorksheetsTypes.Worksheet, function () {
					oThis.WriteWorksheet(ws);
				});
				hasColorFilter = ws.aNamedSheetViews.some(function(namedSheetView){
					return namedSheetView.hasColorFilter();
				});
			}, this.isCopyPaste, this.InitSaveManager.writeOnlySelectedTabs);

			//Fix excel crash with colorFilter in NamedSheetView
            var dxfs = this.InitSaveManager.getDxfs();
			if(hasColorFilter && 0 === dxfs.length) {
                dxfs.push(new AscCommonExcel.CellXfs());
			}
        };
        this.WriteWorksheet = function(ws)
        {
            var i;
            var oThis = this;
            this.bs.WriteItem(c_oSerWorksheetsTypes.WorksheetProp, function(){oThis.WriteWorksheetProp(ws);});

            if(ws.aCols.length > 0 || null != ws.oAllCol)
                this.bs.WriteItem(c_oSerWorksheetsTypes.Cols, function(){oThis.WriteWorksheetCols(ws);});

            //if(!oThis.isCopyPaste)
               this.bs.WriteItem(c_oSerWorksheetsTypes.SheetViews, function(){oThis.WriteSheetViews(ws);});

            if (null !== ws.sheetPr)
                this.bs.WriteItem(c_oSerWorksheetsTypes.SheetPr, function () {oThis.WriteSheetPr(ws.sheetPr);});

            this.bs.WriteItem(c_oSerWorksheetsTypes.SheetFormatPr, function(){oThis.WriteSheetFormatPr(ws);});

            if(null != ws.PagePrintOptions)
            {
                this.bs.WriteItem(c_oSerWorksheetsTypes.PageMargins, function(){oThis.WritePageMargins(ws.PagePrintOptions.asc_getPageMargins());});

                this.bs.WriteItem(c_oSerWorksheetsTypes.PageSetup, function(){oThis.WritePageSetup(ws.PagePrintOptions.asc_getPageSetup());});

                this.bs.WriteItem(c_oSerWorksheetsTypes.PrintOptions, function(){oThis.WritePrintOptions(ws.PagePrintOptions);});
            }

			this.bs.WriteItem(c_oSerWorksheetsTypes.SheetData, function() {
				oThis.bs.WriteItem(c_oSerWorksheetsTypes.XlsbPos, function() {
					oThis.memory.WriteULong(oThis.memory.GetCurPosition() + 4);
					oThis.WriteSheetDataXLSB(ws);
				});
			});

            this.bs.WriteItem(c_oSerWorksheetsTypes.Hyperlinks, function(){oThis.WriteHyperlinks(ws);});

            this.bs.WriteItem(c_oSerWorksheetsTypes.MergeCells, function(){oThis.WriteMergeCells(ws);});

            if (ws.Drawings && (ws.Drawings.length))
                this.bs.WriteItem(c_oSerWorksheetsTypes.Drawings, function(){oThis.WriteDrawings(ws.Drawings);});

            if (ws.aComments.length > 0) {
                this.bs.WriteItem(c_oSerWorksheetsTypes.Comments, function () {
                    oThis.WriteComments(ws.aComments, ws);
                });
            }

            var oBinaryTableWriter;
            if(null != ws.AutoFilter && !this.isCopyPaste)
            {
                oBinaryTableWriter = new BinaryTableWriter(this.memory, this.InitSaveManager, false, {});
                this.bs.WriteItem(c_oSerWorksheetsTypes.Autofilter, function(){oBinaryTableWriter.WriteAutoFilter(ws.AutoFilter);});
            }
            if(null != ws.sortState && !this.isCopyPaste)
            {
                oBinaryTableWriter = new BinaryTableWriter(this.memory, this.InitSaveManager, false, {});
                this.bs.WriteItem(c_oSerWorksheetsTypes.SortState, function(){oBinaryTableWriter.WriteSortState(ws.sortState);});
            }
            if(null != ws.TableParts && ws.TableParts.length > 0)
            {
                oBinaryTableWriter = new BinaryTableWriter(this.memory, this.InitSaveManager, this.isCopyPaste, this.InitSaveManager.getTableIds());
                this.bs.WriteItem(c_oSerWorksheetsTypes.TableParts, function(){oBinaryTableWriter.Write(ws.TableParts, ws);});
            }
			if (ws.aSparklineGroups.length > 0) {
                this.bs.WriteItem(c_oSerWorksheetsTypes.SparklineGroups, function(){oThis.WriteSparklineGroups(ws.aSparklineGroups);});
            }
            // ToDo combine rules for matching ranges
            ws.forEachConditionalFormattingRules(function (elem) {
                oThis.bs.WriteItem(c_oSerWorksheetsTypes.ConditionalFormatting, function(){oThis.WriteConditionalFormatting(elem);});
            });
			for (i = 0; i < ws.pivotTables.length; ++i) {
				this.bs.WriteItem(c_oSerWorksheetsTypes.PivotTable, function(){oThis.WritePivotTable(ws.pivotTables[i], oThis.isCopyPaste)});
			}
            var slicers = new Asc.CT_slicers();
            var slicerExt = new Asc.CT_slicers();
            for (var i = 0; i < ws.aSlicers.length; ++i) {
                if (this.isCopyPaste) {
                    var _graphicObject = ws.workbook.getSlicerViewByName(ws.aSlicers[i].name);
                    if (!_graphicObject || !_graphicObject.selected) {
                        continue;
                    }
                }

                if (ws.aSlicers[i].isExt()) {
                    slicerExt.slicer.push(ws.aSlicers[i]);
                } else {
                    slicers.slicer.push(ws.aSlicers[i]);
                }
            }
            if (slicers.slicer.length > 0) {
                this.bs.WriteItem(c_oSerWorksheetsTypes.Slicers, function () {oThis.WriteSlicers(slicers);});
            }
            if (slicerExt.slicer.length > 0) {
                this.bs.WriteItem(c_oSerWorksheetsTypes.SlicersExt, function () {oThis.WriteSlicers(slicerExt);});
            }
            if (null !== ws.headerFooter) {
                this.bs.WriteItem(c_oSerWorksheetsTypes.HeaderFooter, function () {oThis.WriteHeaderFooter(ws.headerFooter);});
            }
            if (null !== ws.rowBreaks) {
                this.bs.WriteItem(c_oSerWorksheetsTypes.RowBreaks, function () {oThis.WriteRowColBreaks(ws.rowBreaks);});
            }
            if (null !== ws.colBreaks) {
                this.bs.WriteItem(c_oSerWorksheetsTypes.ColBreaks, function () {oThis.WriteRowColBreaks(ws.colBreaks);});
            }
            if (null !== ws.legacyDrawingHF) {
                this.bs.WriteItem(c_oSerWorksheetsTypes.LegacyDrawingHF, function () {oThis.WriteLegacyDrawingHF(ws.legacyDrawingHF);});
            }
            if (null !== ws.picture) {
                this.memory.WriteByte(c_oSerWorksheetsTypes.Picture);
                this.memory.WriteString2(ws.picture);
            }
			if (null !== ws.dataValidations) {
				this.bs.WriteItem(c_oSerWorksheetsTypes.DataValidations, function () {oThis.WriteDataValidations(ws.dataValidations);});
			}
			if (ws.aNamedSheetViews.length > 0) {
				this.bs.WriteItem(c_oSerWorksheetsTypes.NamedSheetView, function () {
					var namedSheetViews = new Asc.CT_NamedSheetViews();
					namedSheetViews.namedSheetView = ws.aNamedSheetViews;

					var old = new AscCommon.CMemory(true);
					pptx_content_writer.BinaryFileWriter.ExportToMemory(old);
					pptx_content_writer.BinaryFileWriter.ImportFromMemory(oThis.memory);

					pptx_content_writer.BinaryFileWriter.StartRecord(0);
					namedSheetViews.toStream(pptx_content_writer.BinaryFileWriter, oThis.InitSaveManager.getTableIds());
					pptx_content_writer.BinaryFileWriter.EndRecord();

					pptx_content_writer.BinaryFileWriter.ExportToMemory(oThis.memory);
					pptx_content_writer.BinaryFileWriter.ImportFromMemory(old);
				});
			}
			if (null !== ws.sheetProtection) {
				this.bs.WriteItem(c_oSerWorksheetsTypes.ProtectionSheet, function () {oThis.WriteSheetProtection(ws.sheetProtection);});
			}
			if (ws.aProtectedRanges && ws.aProtectedRanges.length > 0) {
				this.bs.WriteItem(c_oSerWorksheetsTypes.ProtectedRanges, function(){oThis.WriteProtectedRanges(ws.aProtectedRanges);});
			}
            if (ws.aCellWatches && ws.aCellWatches.length > 0) {
                this.bs.WriteItem(c_oSerWorksheetsTypes.CellWatches, function(){oThis.WriteCellWatches(ws.aCellWatches);});
            }

            if (ws.userProtectedRanges && ws.userProtectedRanges.length > 0) {
                this.bs.WriteItem(c_oSerWorksheetsTypes.UserProtectedRanges, function(){oThis.WriteUserProtectedRanges(ws.userProtectedRanges);});
            }

            if (ws.timelines && ws.timelines.length > 0) {
                this.bs.WriteItem(c_oSerWorksheetsTypes.TimelinesList, function(){oThis.WriteTimelines(ws.timelines);});
            }

            let pTableSingleCells = ws.oTableSingleCells;
            if (pTableSingleCells) {
                let stream = pptx_content_writer.BinaryFileWriter;
                for (let i = 0; i < pTableSingleCells.length; i++) {
                    this.bs.WriteItem(c_oSerWorksheetsTypes.TableSingleCells, function () {
                        var old = new AscCommon.CMemory(true);
                        stream.ExportToMemory(old);
                        stream.ImportFromMemory(oThis.memory);

                        stream.StartRecord(0);
                        stream.WriteRecord2(0, stream, function () {
                            pTableSingleCells[i].toPPTY(stream);
                        });
                        stream.EndRecord();

                        stream.ExportToMemory(oThis.memory);
                        stream.ImportFromMemory(old);
                    });
                }
            }
        };
		this.WriteDataValidations = function(dataValidations)
		{
			var oThis = this;
			//Name
			if (null != dataValidations.disablePrompts) {
				this.bs.WriteItem(c_oSer_DataValidation.DisablePrompts, function () {oThis.memory.WriteBool(dataValidations.disablePrompts);});
			}
			if (null != dataValidations.xWindow) {
				this.bs.WriteItem(c_oSer_DataValidation.XWindow, function () {oThis.memory.WriteLong(dataValidations.xWindow);});
			}
			if (null != dataValidations.yWindow) {
				this.bs.WriteItem(c_oSer_DataValidation.YWindow, function () {oThis.memory.WriteLong(dataValidations.yWindow);});
			}
			if (dataValidations.elems.length > 0) {
				this.bs.WriteItem(c_oSer_DataValidation.DataValidations, function () {
					for (var i = 0; i < dataValidations.elems.length; ++i) {
						oThis.bs.WriteItem(c_oSer_DataValidation.DataValidation, function () {oThis.WriteDataValidation(dataValidations.elems[i]);});
					}
				});
			}
		};
		this.WriteDataValidation = function(dataValidation)
		{
			//Name
			if (null != dataValidation.allowBlank) {
				this.memory.WriteByte(c_oSer_DataValidation.AllowBlank);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(dataValidation.allowBlank);
			}
			if (null != dataValidation.type) {
				this.memory.WriteByte(c_oSer_DataValidation.Type);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteByte(dataValidation.type);
			}
			if (null != dataValidation.error) {
				this.memory.WriteByte(c_oSer_DataValidation.Error);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(dataValidation.error);
			}
			if (null != dataValidation.errorTitle) {
				this.memory.WriteByte(c_oSer_DataValidation.ErrorTitle);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(dataValidation.errorTitle);
			}
			if (null != dataValidation.errorStyle) {
				this.memory.WriteByte(c_oSer_DataValidation.ErrorStyle);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteByte(dataValidation.errorStyle);
			}
			if (null != dataValidation.imeMode) {
				this.memory.WriteByte(c_oSer_DataValidation.ImeMode);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteByte(dataValidation.imeMode);
			}
			if (null != dataValidation.operator) {
				this.memory.WriteByte(c_oSer_DataValidation.Operator);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteByte(dataValidation.operator);
			}
			if (null != dataValidation.prompt) {
				this.memory.WriteByte(c_oSer_DataValidation.Prompt);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(dataValidation.prompt);
			}
			if (null != dataValidation.promptTitle) {
				this.memory.WriteByte(c_oSer_DataValidation.PromptTitle);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(dataValidation.promptTitle);
			}
			if (null != dataValidation.showDropDown) {
				this.memory.WriteByte(c_oSer_DataValidation.ShowDropDown);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(dataValidation.showDropDown);
			}
			if (null != dataValidation.showErrorMessage) {
				this.memory.WriteByte(c_oSer_DataValidation.ShowErrorMessage);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(dataValidation.showErrorMessage);
			}
			if (null != dataValidation.showInputMessage) {
				this.memory.WriteByte(c_oSer_DataValidation.ShowInputMessage);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(dataValidation.showInputMessage);
			}
			if (null != dataValidation.ranges) {
				this.memory.WriteByte(c_oSer_DataValidation.SqRef);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(getSqRefString(dataValidation.ranges));
			}
			if (null != dataValidation.formula1) {
				this.memory.WriteByte(c_oSer_DataValidation.Formula1);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(dataValidation.formula1.text);
			}
			if (null != dataValidation.formula2) {
				this.memory.WriteByte(c_oSer_DataValidation.Formula2);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(dataValidation.formula2.text);
			}
			if (null != dataValidation.list) {
				this.memory.WriteByte(c_oSer_DataValidation.List);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(dataValidation.list);
			}
		};
		this.WriteSheetProtection = function(sheetProtection)
		{
		    if (null != sheetProtection.algorithmName) {
				this.memory.WriteByte(c_oSerWorksheetProtection.AlgorithmName);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteByte(sheetProtection.algorithmName);
			}
			if (null != sheetProtection.spinCount) {
				this.memory.WriteByte(c_oSerWorksheetProtection.SpinCount);
				this.memory.WriteByte(c_oSerPropLenType.Long);
				this.memory.WriteLong(sheetProtection.spinCount);
			}
			if (null != sheetProtection.hashValue) {
				this.memory.WriteByte(c_oSerWorksheetProtection.HashValue);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(sheetProtection.hashValue);
			}
			if (null != sheetProtection.saltValue) {
				this.memory.WriteByte(c_oSerWorksheetProtection.SaltValue);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(sheetProtection.saltValue);
			}
			//todo password
			if (null != sheetProtection.password) {
				this.memory.WriteByte(c_oSerWorksheetProtection.Password);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(sheetProtection.password);
			}

			if (null != sheetProtection.autoFilter) {
				this.memory.WriteByte(c_oSerWorksheetProtection.AutoFilter);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.autoFilter);
			}
			if (null != sheetProtection.content) {
				this.memory.WriteByte(c_oSerWorksheetProtection.Content);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.content);
			}
			if (null != sheetProtection.deleteColumns) {
				this.memory.WriteByte(c_oSerWorksheetProtection.DeleteColumns);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.deleteColumns);
			}
			if (null != sheetProtection.deleteRows) {
				this.memory.WriteByte(c_oSerWorksheetProtection.DeleteRows);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.deleteRows);
			}
			if (null != sheetProtection.formatCells) {
				this.memory.WriteByte(c_oSerWorksheetProtection.FormatCells);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.formatCells);
			}
			if (null != sheetProtection.formatColumns) {
				this.memory.WriteByte(c_oSerWorksheetProtection.FormatColumns);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.formatColumns);
			}
			if (null != sheetProtection.formatRows) {
				this.memory.WriteByte(c_oSerWorksheetProtection.FormatRows);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.formatRows);
			}
			if (null != sheetProtection.insertColumns) {
				this.memory.WriteByte(c_oSerWorksheetProtection.InsertColumns);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.insertColumns);
			}
			if (null != sheetProtection.insertHyperlinks) {
				this.memory.WriteByte(c_oSerWorksheetProtection.InsertHyperlinks);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.insertHyperlinks);
			}
			if (null != sheetProtection.insertRows) {
				this.memory.WriteByte(c_oSerWorksheetProtection.InsertRows);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.insertRows);
			}
			if (null != sheetProtection.objects && sheetProtection.sheet) {
				this.memory.WriteByte(c_oSerWorksheetProtection.Objects);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.objects);
			}
			if (null != sheetProtection.pivotTables) {
				this.memory.WriteByte(c_oSerWorksheetProtection.PivotTables);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.pivotTables);
			}
			//scenarios - в MS может быть всталена в true если только лист защищен
			//если выставлена в true при незазищенном листе, то после сохранения и открытия в ms обратно в false возвращается
			if (null != sheetProtection.scenarios && sheetProtection.sheet) {
				this.memory.WriteByte(c_oSerWorksheetProtection.Scenarios);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.scenarios);
			}
			if (null != sheetProtection.selectLockedCells) {
				this.memory.WriteByte(c_oSerWorksheetProtection.SelectLockedCells);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.selectLockedCells);
			}
			if (null != sheetProtection.selectUnlockedCells) {
				this.memory.WriteByte(c_oSerWorksheetProtection.SelectUnlockedCells);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.selectUnlockedCells);
			}
			if (null != sheetProtection.sheet) {
				this.memory.WriteByte(c_oSerWorksheetProtection.Sheet);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.sheet);
			}
			if (null != sheetProtection.sort) {
				this.memory.WriteByte(c_oSerWorksheetProtection.Sort);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteBool(sheetProtection.sort);
			}
		};
		this.WriteProtectedRanges = function (aProtectedRanges) {
			var oThis = this;
			for (var i = 0, length = aProtectedRanges.length; i < length; ++i) {
				this.bs.WriteItem(c_oSerWorksheetsTypes.ProtectedRange, function () {
					oThis.WriteProtectedRange(aProtectedRanges[i]);
				});
			}
		};
		this.WriteProtectedRange = function (oProtectedRange) {
			if (null != oProtectedRange.algorithmName) {
				this.memory.WriteByte(c_oSerProtectedRangeTypes.AlgorithmName);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteByte(oProtectedRange.algorithmName);
			}
			if (null != oProtectedRange.spinCount) {
				this.memory.WriteByte(c_oSerProtectedRangeTypes.SpinCount);
				this.memory.WriteByte(c_oSerPropLenType.Long);
				this.memory.WriteLong(oProtectedRange.spinCount);
			}
			if (null != oProtectedRange.hashValue) {
				this.memory.WriteByte(c_oSerProtectedRangeTypes.HashValue);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(oProtectedRange.hashValue);
			}
			if (null != oProtectedRange.saltValue) {
				this.memory.WriteByte(c_oSerProtectedRangeTypes.SaltValue);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(oProtectedRange.saltValue);
			}
			if (null != oProtectedRange.name) {
				this.memory.WriteByte(c_oSerProtectedRangeTypes.Name);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(oProtectedRange.name);
			}

			if (null != oProtectedRange.sqref) {
				this.memory.WriteByte(c_oSerProtectedRangeTypes.SqRef);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				var sqRef = getSqRefString(oProtectedRange.sqref);
				this.memory.WriteString2(sqRef);
			}

			//TODO
			if (null != oProtectedRange.securityDescriptor) {
				this.memory.WriteByte(c_oSerProtectedRangeTypes.SecurityDescriptor);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(oProtectedRange.securityDescriptor);
			}
		};
        this.WriteCellWatches = function (aCellWatches) {
            var oThis = this;
            for (var i = 0, length = aCellWatches.length; i < length; ++i) {
                this.bs.WriteItem(c_oSerWorksheetsTypes.CellWatch, function () {
                    oThis.WriteCellWatch(aCellWatches[i]);
                });
            }
        };
        this.WriteCellWatch = function (oCellWatch) {
            if (null != oCellWatch.r) {
                this.memory.WriteByte(c_oSerWorksheetsTypes.CellWatchR);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(oCellWatch.r.getName());
            }
        };
        this.WriteWorksheetProp = function(ws)
        {
            var oThis = this;
            //Name
            this.memory.WriteByte(c_oSerWorksheetPropTypes.Name);
            this.memory.WriteByte(c_oSerPropLenType.Variable);
            this.memory.WriteString2(ws.sName);
            //SheetId
            this.memory.WriteByte(c_oSerWorksheetPropTypes.SheetId);
            this.memory.WriteByte(c_oSerPropLenType.Long);
            this.memory.WriteLong(oThis.InitSaveManager.sheetIds[ws.getId()] || 1);
            //Hidden
            if(null != ws.bHidden)
            {
                this.memory.WriteByte(c_oSerWorksheetPropTypes.State);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                if(true == ws.bHidden)
                    this.memory.WriteByte(EVisibleType.visibleHidden);
                else
                    this.memory.WriteByte(EVisibleType.visibleVisible);
            }
            //activeRange(serialize activeRange)
            if(oThis.isCopyPaste)
            {
                var activeRange = oThis.isCopyPaste;

                var newRange = null;
                if(ws.bExcludeHiddenRows)
                {
                    for(var i = activeRange.r1; i < activeRange.r2; i++)
                    {
                        if(ws.getRowHidden(i))
                        {
                            if(!newRange)
                            {
								newRange = activeRange.clone();
                            }
							newRange.r2--;
                        }
                    }

                    if(newRange)
                    {
						activeRange = newRange;
                    }
                }

                this.memory.WriteByte(c_oSerWorksheetPropTypes.Ref);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(activeRange.getName());
            }
        };
        this.WriteWorksheetCols = function(ws)
        {
            var oThis = this;
            this.InitSaveManager.writeCols(ws, this.stylesForWrite, function (oColToWrite) {
                oThis.bs.WriteItem(c_oSerWorksheetsTypes.Col, function(){oThis.WriteWorksheetCol(oColToWrite);});
            });
        };
        this.WriteWorksheetCol = function(oTmpCol)
        {
            var oCol = oTmpCol.col;
            if(null != oCol.BestFit)
            {
                this.memory.WriteByte(c_oSerWorksheetColTypes.BestFit);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(oCol.BestFit);
            }
            if(oCol.hd)
            {
                this.memory.WriteByte(c_oSerWorksheetColTypes.Hidden);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(oCol.hd);
            }
            if(null != oTmpCol.Max)
            {
                this.memory.WriteByte(c_oSerWorksheetColTypes.Max);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oTmpCol.Max);
            }
            if(null != oTmpCol.Min)
            {
                this.memory.WriteByte(c_oSerWorksheetColTypes.Min);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oTmpCol.Min);
            }
            if(null != oTmpCol.xfsid)
            {
                this.memory.WriteByte(c_oSerWorksheetColTypes.Style);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oTmpCol.xfsid);
            }
            if(null != oTmpCol.width)
            {
                this.memory.WriteByte(c_oSerWorksheetColTypes.Width);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(oTmpCol.width);
            }
            if(null != oCol.CustomWidth)
            {
                this.memory.WriteByte(c_oSerWorksheetColTypes.CustomWidth);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(oCol.CustomWidth);
            }
            if (oCol.outlineLevel > 0)
            {
                this.memory.WriteByte(c_oSerWorksheetColTypes.OutLevel);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oCol.outlineLevel);
            }
            if (oCol.collapsed)
            {
                this.memory.WriteByte(c_oSerWorksheetColTypes.Collapsed);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(true);
            }
        };
        this.WriteSheetViews = function (ws) {
            var oThis = this;
            for (var i = 0, length = ws.sheetViews.length; i < length; ++i) {
				this.bs.WriteItem(c_oSerWorksheetsTypes.SheetView, function(){oThis.WriteSheetView(ws, ws.sheetViews[i]);});
            }
        };
		this.WriteSheetView = function (ws, oSheetView) {
            var oThis = this;
            if (null !== oSheetView.showGridLines && !oThis.isCopyPaste)
                this.bs.WriteItem(c_oSer_SheetView.ShowGridLines, function(){oThis.memory.WriteBool(oSheetView.showGridLines);});
            if (null !== oSheetView.showRowColHeaders && !oThis.isCopyPaste)
                this.bs.WriteItem(c_oSer_SheetView.ShowRowColHeaders, function(){oThis.memory.WriteBool(oSheetView.showRowColHeaders);});
			if (null !== oSheetView.zoomScale && !oThis.isCopyPaste)
				this.bs.WriteItem(c_oSer_SheetView.ZoomScale, function(){oThis.memory.WriteLong(oSheetView.zoomScale);});
            if (null !== oSheetView.pane && oSheetView.pane.isInit() && !oThis.isCopyPaste)
                this.bs.WriteItem(c_oSer_SheetView.Pane, function(){oThis.WriteSheetViewPane(oSheetView.pane);});
			if (null !== ws.selectionRange)
				this.bs.WriteItem(c_oSer_SheetView.Selection, function(){oThis.WriteSheetViewSelection(ws.selectionRange);});
            if (null !== oSheetView.showZeros && !oThis.isCopyPaste)
                this.bs.WriteItem(c_oSer_SheetView.ShowZeros, function(){oThis.memory.WriteBool(oSheetView.showZeros);});
            if (null !== oSheetView.showFormulas && !oThis.isCopyPaste)
                this.bs.WriteItem(c_oSer_SheetView.ShowFormulas, function(){oThis.memory.WriteBool(oSheetView.showFormulas);});
            if (null !== oSheetView.topLeftCell && !oThis.isCopyPaste)
                this.bs.WriteItem(c_oSer_SheetView.TopLeftCell, function(){oThis.memory.WriteString3(oSheetView.topLeftCell.getName());});
            if (null !== oSheetView.view && !oThis.isCopyPaste)
                this.bs.WriteItem(c_oSer_SheetView.View, function(){oThis.memory.WriteByte(oSheetView.view);});
            if (null !== oSheetView.rightToLeft && !oThis.isCopyPaste)
                this.bs.WriteItem(c_oSer_SheetView.RightToLeft, function(){oThis.memory.WriteBool(oSheetView.rightToLeft);});
        };
        this.WriteSheetViewPane = function (oPane) {
            var oThis = this;
			var col = oPane.topLeftFrozenCell.getCol0();
			var row = oPane.topLeftFrozenCell.getRow0();

			var activePane = null;
			if (0 < col && 0 < row) {
				activePane = EActivePane.bottomRight;
            } else if (0 < row) {
				activePane = EActivePane.bottomLeft;
            } else if (0 < col) {
	            activePane = EActivePane.topRight;
            }
            if (null !== activePane) {
				this.bs.WriteItem(c_oSer_Pane.ActivePane, function(){oThis.memory.WriteByte(activePane);});
            }

            // Всегда пишем Frozen
            this.bs.WriteItem(c_oSer_Pane.State, function(){oThis.memory.WriteString3(AscCommonExcel.c_oAscPaneState.Frozen);});
            this.bs.WriteItem(c_oSer_Pane.TopLeftCell, function(){oThis.memory.WriteString3(oPane.topLeftFrozenCell.getID());});

            if (0 < col)
                this.bs.WriteItem(c_oSer_Pane.XSplit, function(){oThis.memory.WriteDouble2(col);});
            if (0 < row)
                this.bs.WriteItem(c_oSer_Pane.YSplit, function(){oThis.memory.WriteDouble2(row);});
        };
		this.WriteSheetViewSelection = function (selectionRange) {
			var oThis = this;
			if (null != selectionRange.activeCell) {
				this.bs.WriteItem(c_oSer_Selection.ActiveCell, function(){oThis.memory.WriteString3(selectionRange.activeCell.getName());});
			}
			if (null != selectionRange.activeCellId) {
				this.bs.WriteItem(c_oSer_Selection.ActiveCellId, function(){oThis.memory.WriteLong(selectionRange.activeCellId);});
			}
			//this.bs.WriteItem(c_oSer_Selection.Pane, function(){oThis.memory.WriteByte();});
			if (null != selectionRange.ranges) {
				var sqRef = getSqRefString(selectionRange.ranges);
				this.bs.WriteItem(c_oSer_Selection.Sqref, function(){oThis.memory.WriteString3(sqRef);});
			}
		};
        this.WriteSheetPr = function (sheetPr) {
            var oThis = this;
            if (null !== sheetPr.CodeName)
                this.bs.WriteItem(c_oSer_SheetPr.CodeName, function(){oThis.memory.WriteString3(sheetPr.CodeName);});
            if (null !== sheetPr.EnableFormatConditionsCalculation)
                this.bs.WriteItem(c_oSer_SheetPr.EnableFormatConditionsCalculation, function(){oThis.memory.WriteBool(sheetPr.EnableFormatConditionsCalculation);});
            if (null !== sheetPr.FilterMode)
                this.bs.WriteItem(c_oSer_SheetPr.FilterMode, function(){oThis.memory.WriteBool(sheetPr.FilterMode);});
            if (null !== sheetPr.Published)
                this.bs.WriteItem(c_oSer_SheetPr.Published, function(){oThis.memory.WriteBool(sheetPr.Published);});
            if (null !== sheetPr.SyncHorizontal)
                this.bs.WriteItem(c_oSer_SheetPr.SyncHorizontal, function(){oThis.memory.WriteBool(sheetPr.SyncHorizontal);});
            if (null !== sheetPr.SyncRef)
                this.bs.WriteItem(c_oSer_SheetPr.SyncRef, function(){oThis.memory.WriteString3(sheetPr.SyncRef);});
            if (null !== sheetPr.SyncVertical)
                this.bs.WriteItem(c_oSer_SheetPr.SyncVertical, function(){oThis.memory.WriteBool(sheetPr.SyncVertical);});
            if (null !== sheetPr.TransitionEntry)
                this.bs.WriteItem(c_oSer_SheetPr.TransitionEntry, function(){oThis.memory.WriteBool(sheetPr.TransitionEntry);});
            if (null !== sheetPr.TransitionEvaluation)
                this.bs.WriteItem(c_oSer_SheetPr.TransitionEvaluation, function(){oThis.memory.WriteBool(sheetPr.TransitionEvaluation);});
            if (null !== sheetPr.TabColor)
                this.bs.WriteItem(c_oSer_SheetPr.TabColor, function(){oThis.bs.WriteColorSpreadsheet(sheetPr.TabColor);});
			if (null !== sheetPr.AutoPageBreaks || null !== sheetPr.FitToPage)
				this.bs.WriteItem(c_oSer_SheetPr.PageSetUpPr, function(){oThis.WritePageSetUpPr(sheetPr);});
			if (null !== sheetPr.ApplyStyles || null !== sheetPr.ShowOutlineSymbols || null !== sheetPr.SummaryBelow || null !== sheetPr.SummaryRight)
				this.bs.WriteItem(c_oSer_SheetPr.OutlinePr, function(){oThis.WriteOutlinePr(sheetPr);});
        };
		this.WriteOutlinePr = function(sheetPr)
		{
			var oThis = this;
			if (null !== sheetPr.ApplyStyles) {
				this.bs.WriteItem(c_oSer_SheetPr.ApplyStyles, function(){oThis.memory.WriteBool(sheetPr.ApplyStyles);});
			}
			if (null !== sheetPr.ShowOutlineSymbols) {
				this.bs.WriteItem(c_oSer_SheetPr.ShowOutlineSymbols, function(){oThis.memory.WriteBool(sheetPr.ShowOutlineSymbols);});
			}
			if (null !== sheetPr.SummaryBelow) {
				this.bs.WriteItem(c_oSer_SheetPr.SummaryBelow, function(){oThis.memory.WriteBool(sheetPr.SummaryBelow);});
			}
			if (null !== sheetPr.SummaryRight) {
				this.bs.WriteItem(c_oSer_SheetPr.SummaryRight, function(){oThis.memory.WriteBool(sheetPr.SummaryRight);});
			}
		};
		this.WritePageSetUpPr = function(sheetPr)
		{
			var oThis = this;
			if (null !== sheetPr.AutoPageBreaks) {
				this.bs.WriteItem(c_oSer_SheetPr.AutoPageBreaks, function(){oThis.memory.WriteBool(sheetPr.AutoPageBreaks);});
			}
			if (null !== sheetPr.FitToPage) {
				this.bs.WriteItem(c_oSer_SheetPr.FitToPage, function(){oThis.memory.WriteBool(sheetPr.FitToPage);});
			}
		};
        this.WriteSheetFormatPr = function(ws)
        {
            if (null !== ws.oSheetFormatPr.nBaseColWidth) {
                this.memory.WriteByte(c_oSerSheetFormatPrTypes.BaseColWidth);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(ws.oSheetFormatPr.nBaseColWidth);
            }
            if(null !== ws.oSheetFormatPr.dDefaultColWidth) {
                this.memory.WriteByte(c_oSerSheetFormatPrTypes.DefaultColWidth);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(ws.oSheetFormatPr.dDefaultColWidth);
            }
			if (ws.oSheetFormatPr.nOutlineLevelCol > 0) {
				this.memory.WriteByte(c_oSerSheetFormatPrTypes.OutlineLevelCol);
				this.memory.WriteByte(c_oSerPropLenType.Long);
				this.memory.WriteLong(ws.oSheetFormatPr.nOutlineLevelCol);
			}
            if(null !== ws.oSheetFormatPr.oAllRow) {
                var oAllRow = ws.oSheetFormatPr.oAllRow;
                if(oAllRow.h)
                {
                    this.memory.WriteByte(c_oSerSheetFormatPrTypes.DefaultRowHeight);
                    this.memory.WriteByte(c_oSerPropLenType.Double);
                    this.memory.WriteDouble2(oAllRow.h);
                }
                if(oAllRow.getCustomHeight())
                {
                    this.memory.WriteByte(c_oSerSheetFormatPrTypes.CustomHeight);
                    this.memory.WriteByte(c_oSerPropLenType.Byte);
                    this.memory.WriteBool(true);
                }
                if(oAllRow.getHidden())
                {
                    this.memory.WriteByte(c_oSerSheetFormatPrTypes.ZeroHeight);
                    this.memory.WriteByte(c_oSerPropLenType.Byte);
                    this.memory.WriteBool(true);
                }
                if (oAllRow.getOutlineLevel() > 0) {
                    this.memory.WriteByte(c_oSerSheetFormatPrTypes.OutlineLevelRow);
                    this.memory.WriteByte(c_oSerPropLenType.Long);
                    this.memory.WriteLong(oAllRow.getOutlineLevel());
                }
            }
        };
        this.WritePageMargins = function(oMargins)
        {
            //Left
            var dLeft = oMargins.asc_getLeft();
            if(null != dLeft)
            {
                this.memory.WriteByte(c_oSer_PageMargins.Left);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(dLeft);
            }
            //Top
            var dTop = oMargins.asc_getTop();
            if(null != dTop)
            {
                this.memory.WriteByte(c_oSer_PageMargins.Top);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(dTop);
            }
            //Right
            var dRight = oMargins.asc_getRight();
            if(null != dRight)
            {
                this.memory.WriteByte(c_oSer_PageMargins.Right);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(dRight);
            }
            //Bottom
            var dBottom = oMargins.asc_getBottom();
            if(null != dBottom)
            {
                this.memory.WriteByte(c_oSer_PageMargins.Bottom);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(dBottom);
            }

			var dHeader = oMargins.asc_getHeader();
			if(null != dHeader) {
				this.memory.WriteByte(c_oSer_PageMargins.Header);
				this.memory.WriteByte(c_oSerPropLenType.Double);
				this.memory.WriteDouble2(dHeader);//0.5inch
				//this.memory.WriteDouble2(12.7);//0.5inch
			}

			var dFooter = oMargins.asc_getFooter();
			if(null != dFooter) {
				this.memory.WriteByte(c_oSer_PageMargins.Footer);
				this.memory.WriteByte(c_oSerPropLenType.Double);
				this.memory.WriteDouble2(dFooter);//0.5inch
				//this.memory.WriteDouble2(12.7);//0.5inch
			}
        };
        this.WritePageSetup = function(oPageSetup)
        {
            //PageSize
            var dWidth = oPageSetup.asc_getWidth();
            var dHeight = oPageSetup.asc_getHeight();
            if(null != dWidth && null != dHeight)
            {
                var item = DocumentPageSize.getSizeByWH(dWidth, dHeight);
                this.memory.WriteByte(c_oSer_PageSetup.PaperSize);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(item.id);
            }
            if (null != oPageSetup.blackAndWhite) {
                this.memory.WriteByte(c_oSer_PageSetup.BlackAndWhite);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(oPageSetup.blackAndWhite);
            }
            if (null != oPageSetup.cellComments) {
                this.memory.WriteByte(c_oSer_PageSetup.CellComments);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(oPageSetup.cellComments);
            }
            if (null != oPageSetup.copies) {
                this.memory.WriteByte(c_oSer_PageSetup.Copies);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oPageSetup.copies);
            }
            if (null != oPageSetup.draft) {
                this.memory.WriteByte(c_oSer_PageSetup.Draft);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(oPageSetup.draft);
            }
            if (null != oPageSetup.errors) {
                this.memory.WriteByte(c_oSer_PageSetup.Errors);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(oPageSetup.errors);
            }
            if (null != oPageSetup.firstPageNumber) {
                this.memory.WriteByte(c_oSer_PageSetup.FirstPageNumber);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oPageSetup.firstPageNumber);
            }
            if (null != oPageSetup.fitToHeight) {
                this.memory.WriteByte(c_oSer_PageSetup.FitToHeight);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oPageSetup.fitToHeight);
            }
            if (null != oPageSetup.fitToWidth) {
                this.memory.WriteByte(c_oSer_PageSetup.FitToWidth);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oPageSetup.fitToWidth);
            }
            if (null != oPageSetup.horizontalDpi) {
                this.memory.WriteByte(c_oSer_PageSetup.HorizontalDpi);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oPageSetup.horizontalDpi);
            }
            //Orientation
            var byteOrientation = oPageSetup.asc_getOrientation();
            if(null != byteOrientation)
            {
                var byteFormatOrientation = null;
                switch(byteOrientation)
                {
                    case c_oAscPageOrientation.PagePortrait: byteFormatOrientation = EPageOrientation.pageorientPortrait;break;
                    case c_oAscPageOrientation.PageLandscape: byteFormatOrientation = EPageOrientation.pageorientLandscape;break;
                }
                if(null != byteFormatOrientation)
                {
                    this.memory.WriteByte(c_oSer_PageSetup.Orientation);
                    this.memory.WriteByte(c_oSerPropLenType.Byte);
                    this.memory.WriteByte(byteFormatOrientation);
                }
            }
            if (null != oPageSetup.pageOrder) {
                this.memory.WriteByte(c_oSer_PageSetup.PageOrder);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteByte(oPageSetup.pageOrder);
            }
            // if (null != oPageSetup.height) {
            //     this.memory.WriteByte(c_oSer_PageSetup.PaperHeight);
            //     this.memory.WriteByte(c_oSerPropLenType.Double);
            //     this.memory.WriteDouble2(oPageSetup.height);
            // }
            // if (null != oPageSetup.width) {
            //     this.memory.WriteByte(c_oSer_PageSetup.PaperWidth);
            //     this.memory.WriteByte(c_oSerPropLenType.Double);
            //     this.memory.WriteDouble2(oPageSetup.width);
            // }
            // if (null != oPageSetup.paperUnits) {
            //     this.memory.WriteByte(c_oSer_PageSetup.PaperUnits);
            //     this.memory.WriteByte(c_oSerPropLenType.Byte);
            //     this.memory.WriteByte(oPageSetup.paperUnits);
            // }
            if (null != oPageSetup.scale) {
                this.memory.WriteByte(c_oSer_PageSetup.Scale);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oPageSetup.scale);
            }
            if (null != oPageSetup.useFirstPageNumber) {
                this.memory.WriteByte(c_oSer_PageSetup.UseFirstPageNumber);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(oPageSetup.useFirstPageNumber);
            }
            if (null != oPageSetup.usePrinterDefaults) {
                this.memory.WriteByte(c_oSer_PageSetup.UsePrinterDefaults);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(oPageSetup.usePrinterDefaults);
            }
            if (null != oPageSetup.verticalDpi) {
                this.memory.WriteByte(c_oSer_PageSetup.VerticalDpi);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oPageSetup.verticalDpi);
            }
        };
        this.WritePrintOptions = function (oPrintOptions) {
            //GridLines
            let bGridLines = oPrintOptions.asc_getGridLines();
            if (null != bGridLines) {
                this.memory.WriteByte(c_oSer_PrintOptions.GridLines);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(bGridLines);
            }
            //Headings
            let bHeadings = oPrintOptions.asc_getHeadings();
            if (null != bHeadings) {
                this.memory.WriteByte(c_oSer_PrintOptions.Headings);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(bHeadings);
            }
            //GridLinesSet
            let bGridLinesSet = oPrintOptions.asc_getGridLinesSet();
            if (null != bGridLinesSet) {
                this.memory.WriteByte(c_oSer_PrintOptions.GridLinesSet);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(bGridLinesSet);
            }
            //HorizontalCentered
            let bHorizontalCentered = oPrintOptions.asc_getHorizontalCentered();
            if (null != bHorizontalCentered) {
                this.memory.WriteByte(c_oSer_PrintOptions.HorizontalCentered);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(bHorizontalCentered);
            }
            //VerticalCentered
            let bVerticalCentered = oPrintOptions.asc_getVerticalCentered();
            if (null != bVerticalCentered) {
                this.memory.WriteByte(c_oSer_PrintOptions.VerticalCentered);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(bVerticalCentered);
            }
        };
        this.WriteHyperlinks = function(ws)
        {
            var oThis = this;
            var oHyperlinks = ws.hyperlinkManager.getAll();
            //todo sort
            for(var i in oHyperlinks)
            {
                var elem = oHyperlinks[i];
                //write only active hyperlink, if copy/paste
                if(!this.isCopyPaste || (this.isCopyPaste && elem && elem.bbox && this.isCopyPaste.containsRange(elem.bbox)))
                    this.bs.WriteItem(c_oSerWorksheetsTypes.Hyperlink, function(){oThis.WriteHyperlink(elem.data);});
            }
        };
        this.WriteHyperlink = function (oHyperlink) {
            if (null != oHyperlink.Ref) {
                this.memory.WriteByte(c_oSerHyperlinkTypes.Ref);
                this.memory.WriteString2(oHyperlink.Ref.getName());
            }
            if (null != oHyperlink.Hyperlink) {
                var sHyperlink = oHyperlink.Hyperlink.length > Asc.c_nMaxHyperlinkLength ? this.Hyperlink.substring(0, Asc.c_nMaxHyperlinkLength) : oHyperlink.Hyperlink;
                this.memory.WriteByte(c_oSerHyperlinkTypes.Hyperlink);
                this.memory.WriteString2(sHyperlink);
            }
            if (null !== oHyperlink.getLocation()) {
                this.memory.WriteByte(c_oSerHyperlinkTypes.Location);
                this.memory.WriteString2(oHyperlink.getLocation());
            }
            if (null != oHyperlink.Tooltip) {
                this.memory.WriteByte(c_oSerHyperlinkTypes.Tooltip);
                this.memory.WriteString2(oHyperlink.Tooltip);
            }
        };
        this.WriteMergeCells = function(ws)
        {
            var t = this;
            this.InitSaveManager.WriteMergeCells(ws, function (ref) {
                t.memory.WriteByte(c_oSerWorksheetsTypes.MergeCell);
                t.memory.WriteString2(ref);
            })
        };
        this.WriteDrawings = function(aDrawings)
        {
            var oThis = this;
            var oPPTXWriter = pptx_content_writer.BinaryFileWriter;
            for(var i = 0, length = aDrawings.length; i < length; ++i)
            {
                //write only active drawing, if copy/paste
                var oDrawing = aDrawings[i];
                if(!this.isCopyPaste)
                    this.bs.WriteItem(c_oSerWorksheetsTypes.Drawing, function(){oThis.WriteDrawing(oDrawing);});
                else if(this.isCopyPaste && oDrawing.graphicObject.selected)//for copy/paste
                {
                    if(oDrawing.isGroup() && oDrawing.graphicObject.selectedObjects && oDrawing.graphicObject.selectedObjects.length)
                    {
                        var oDrawingSelected = oDrawing.graphicObject.selectedObjects;
                        var curDrawing, graphicObject;
                        for(var selDr = 0; selDr < oDrawingSelected.length; selDr++)
                        {
                            curDrawing = oDrawingSelected[selDr];

							//меняем graphicObject на время записи
							graphicObject = oDrawing.graphicObject;
							oDrawing.graphicObject = curDrawing;

                            this.bs.WriteItem(c_oSerWorksheetsTypes.Drawing, function(){oThis.WriteDrawing(oDrawing, curDrawing);});

							//возвращаем graphicObject обратно
							oDrawing.graphicObject = graphicObject;
                        }
                    }
                    else
                    {
                        var oCurDrawingToWrite = AscFormat.ExecuteNoHistory(function()
                        {
                            var oRet = oDrawing.graphicObject.copy(undefined);
                            var oMetrics = oDrawing.getGraphicObjectMetrics();
                            AscFormat.SetXfrmFromMetrics(oRet, oMetrics);
                            return oRet;
                        }, this, []);
                        var oOldGrObject = oDrawing.graphicObject;
                        oDrawing.graphicObject = oCurDrawingToWrite;
                        this.bs.WriteItem(c_oSerWorksheetsTypes.Drawing, function(){oThis.WriteDrawing(oDrawing);});
                        oDrawing.graphicObject = oOldGrObject;

                    }
                }
            }
        };
        this.WriteDrawing = function(oDrawing, curDrawing)
        {
            var oThis = this;
            var oDrawingToWrite = curDrawing || oDrawing.graphicObject;

            var nTypeToWrite = oDrawing.Type;
            if(oDrawingToWrite.getObjectType() === AscDFH.historyitem_type_OleObject)
            {
                nTypeToWrite = c_oAscCellAnchorType.cellanchorTwoCell;
            }

            if(null != nTypeToWrite)
                this.bs.WriteItem(c_oSer_DrawingType.Type, function(){oThis.memory.WriteByte(nTypeToWrite);});
            switch(nTypeToWrite)
            {
                case c_oAscCellAnchorType.cellanchorTwoCell:
                {
                    this.bs.WriteItem(c_oSer_DrawingType.EditAs, function(){oThis.memory.WriteByte(oDrawing.editAs);});
                    this.bs.WriteItem(c_oSer_DrawingType.From, function(){oThis.WriteFromTo(oDrawing.from);});
                    this.bs.WriteItem(c_oSer_DrawingType.To, function(){oThis.WriteFromTo(oDrawing.to);});
                    break;
                }
                case c_oAscCellAnchorType.cellanchorOneCell:
                {
                    this.bs.WriteItem(c_oSer_DrawingType.From, function(){oThis.WriteFromTo(oDrawing.from);});
                    this.bs.WriteItem(c_oSer_DrawingType.Ext, function(){oThis.WriteExt(oDrawing.ext);});
                    break;
                }
                case c_oAscCellAnchorType.cellanchorAbsolute:
                {
                    this.bs.WriteItem(c_oSer_DrawingType.Pos, function(){oThis.WritePos(oDrawing.Pos);});
                    this.bs.WriteItem(c_oSer_DrawingType.Ext, function(){oThis.WriteExt(oDrawing.ext);});
                    break;
                }
            }
            var oDrawingForWriting = curDrawing || oDrawing.graphicObject;
            if(oDrawingForWriting)
            {
                if(oDrawingForWriting.clientData)
                {
                    this.bs.WriteItem(c_oSer_DrawingType.ClientData, function(){oThis.WriteClientData(oDrawingForWriting.clientData);});
                }
            }
            this.bs.WriteItem(c_oSer_DrawingType.pptxDrawing, function(){pptx_content_writer.WriteDrawing(oThis.memory, oDrawingToWrite, null, null, null);});
        };
        this.WriteFromTo = function(oFromTo)
        {
            if(null != oFromTo.col)
            {
                this.memory.WriteByte(c_oSer_DrawingFromToType.Col);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oFromTo.col);
            }
            if(null != oFromTo.colOff)
            {
                this.memory.WriteByte(c_oSer_DrawingFromToType.ColOff);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(oFromTo.colOff);
            }
            if(null != oFromTo.row)
            {
                this.memory.WriteByte(c_oSer_DrawingFromToType.Row);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oFromTo.row);
            }
            if(null != oFromTo.rowOff)
            {
                this.memory.WriteByte(c_oSer_DrawingFromToType.RowOff);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(oFromTo.rowOff);
            }
        };
        this.WritePos = function(oPos)
        {
            if(null != oPos.X)
            {
                this.memory.WriteByte(c_oSer_DrawingPosType.X);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(oPos.X);
            }
            if(null != oPos.Y)
            {
                this.memory.WriteByte(c_oSer_DrawingPosType.Y);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(oPos.Y);
            }
        };
        this.WriteExt = function(oExt)
        {
            if(null != oExt.cx)
            {
                this.memory.WriteByte(c_oSer_DrawingExtType.Cx);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(oExt.cx);
            }
            if(null != oExt.cy)
            {
                this.memory.WriteByte(c_oSer_DrawingExtType.Cy);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(oExt.cy);
            }
        };
        this.WriteClientData = function(oClientData)
        {
            if (oClientData.fLocksWithSheet !== null)
            {
                this.memory.WriteByte(c_oSer_DrawingClientDataType.fLocksWithSheet);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(oClientData.fLocksWithSheet);
            }
            if (oClientData.fPrintsWithSheet !== null)
            {
                this.memory.WriteByte(c_oSer_DrawingClientDataType.fPrintsWithSheet);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(oClientData.fPrintsWithSheet);
            }
        };
		this.WriteSheetDataXLSB = function(ws)
        {
            var oThis = this;
            var range;
            if(oThis.isCopyPaste ){
                range = ws.getRange3(oThis.isCopyPaste.r1, oThis.isCopyPaste.c1, oThis.isCopyPaste.r2, oThis.isCopyPaste.c2);
            } else {
                range = ws.getRange3(0, 0, gc_nMaxRow0, gc_nMaxCol0);
            }

            var curRowIndex = -1;
            var curRow = null;
            var curCol = null;
            var allRow = ws.getAllRowNoEmpty();
            var tempRow = new AscCommonExcel.Row(ws);
            if (allRow) {
                tempRow.copyFrom(allRow);
            }
			this.memory.XlsbStartRecord(AscCommonExcel.XLSB.rt_BEGIN_SHEET_DATA, 0);
			this.memory.XlsbEndRecord();

            range._foreachRowNoEmpty(function(row, excludedCount) {
                row.toXLSB(oThis.memory, -excludedCount, oThis.stylesForWrite);
                curRowIndex = row.getIndex();
                curRow = row;
            }, function(cell, nRow0, nCol0, nRowStart0, nColStart0, excludedCount) {
                if (curRowIndex != nRow0) {
                    tempRow.setIndex(nRow0);
					tempRow.toXLSB(oThis.memory, -excludedCount, oThis.stylesForWrite);
					curRowIndex = tempRow.getIndex();
                    curRow = tempRow;
                }
                //готовим ячейку к записи
                var nXfsId;
                var cellXfs = cell.xfs;
                nXfsId = oThis.stylesForWrite.add(cell.xfs);

                // save even an empty style like Excel (needed to remove row/column style)
                let needWrite = cellXfs || !cell.isNullText()
                    || (curRow && curRow.xfs) //override row style
                    || ((curCol = (ws.aCols[nCol0] || ws.oAllCol)) && curCol && curCol.xfs);//override col style
                if (needWrite) {
					var formulaToWrite;
					if (cell.isFormula() && !(oThis.isCopyPaste && cell.ws && cell.ws.bIgnoreWriteFormulas)) {
						formulaToWrite = oThis.InitSaveManager.PrepareFormulaToWrite(cell);
                    }
					cell.toXLSB(oThis.memory, nXfsId, formulaToWrite, oThis.InitSaveManager.oSharedStrings);
				}
            }, (ws.bExcludeHiddenRows && oThis.isCopyPaste));

			this.memory.XlsbStartRecord(AscCommonExcel.XLSB.rt_END_SHEET_DATA, 0);
			this.memory.XlsbEndRecord();
        };
        this.WriteComments = function(aComments, ws)
        {
            var oThis = this;
            var i, elem;
            for(i = 0; i < aComments.length; ++i)
            {
                elem = aComments[i];
                //write only active comments, if copy/paste
                if(this.isCopyPaste)
				{
					//ignore hidden rows if ws.bExcludeHiddenRows === true
					if(!this.isCopyPaste.contains(elem.nCol, elem.nRow) || (ws.bExcludeHiddenRows && ws.getRowHidden(elem.nRow)))
					{
						continue;
					}
				}
				if (elem.coords && elem.coords.isValid()) {
					this.bs.WriteItem(c_oSerWorksheetsTypes.Comment, function(){oThis.WriteComment(elem);});
				}
            }
        };
        this.checkCommentGuid = function(comment) {
            var sGuid = comment.asc_getGuid();
            while (!sGuid || this.InitSaveManager.commentUniqueGuids[sGuid]) {
                sGuid = AscCommon.CreateGUID();
                comment.asc_putGuid(sGuid);
            }
            this.InitSaveManager.commentUniqueGuids[sGuid] = 1;
            if (comment.aReplies) {
                for (var i = 0; i < comment.aReplies.length; ++i) {
                    this.checkCommentGuid(comment.aReplies[i]);
                }
            }
        };
        this.WriteComment = function(comment)
        {
            var oThis = this;
            this.checkCommentGuid(comment);
            var coords = comment.coords;
            if (null != coords.nRow) {
                this.memory.WriteByte(c_oSer_Comments.Row);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(coords.nRow);
            }
            if (null != coords.nCol) {
                this.memory.WriteByte(c_oSer_Comments.Col);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(coords.nCol);
            }

            this.memory.WriteByte(c_oSer_Comments.CommentDatas);
            this.memory.WriteByte(c_oSerPropLenType.Variable);
            this.bs.WriteItemWithLength(function(){oThis.WriteCommentDatas(comment);});

            if (null != coords.nLeft) {
                this.memory.WriteByte(c_oSer_Comments.Left);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(coords.nLeft);
            }
            if (null != coords.nTop) {
                this.memory.WriteByte(c_oSer_Comments.Top);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(coords.nTop);
            }
            if (null != coords.nRight) {
                this.memory.WriteByte(c_oSer_Comments.Right);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(coords.nRight);
            }
            if (null != coords.nBottom) {
                this.memory.WriteByte(c_oSer_Comments.Bottom);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(coords.nBottom);
            }
            if (null != coords.nLeftOffset) {
                this.memory.WriteByte(c_oSer_Comments.LeftOffset);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(coords.nLeftOffset);
            }
            if (null != coords.nTopOffset) {
                this.memory.WriteByte(c_oSer_Comments.TopOffset);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(coords.nTopOffset);
            }
            if (null != coords.nRightOffset) {
                this.memory.WriteByte(c_oSer_Comments.RightOffset);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(coords.nRightOffset);
            }
            if (null != coords.nBottomOffset) {
                this.memory.WriteByte(c_oSer_Comments.BottomOffset);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(coords.nBottomOffset);
            }
            if(null != coords.dLeftMM) {
                this.memory.WriteByte(c_oSer_Comments.LeftMM);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(coords.dLeftMM);
            }
            if(null != coords.dTopMM) {
                this.memory.WriteByte(c_oSer_Comments.TopMM);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(coords.dTopMM);
            }
            if (null != coords.dWidthMM) {
                this.memory.WriteByte(c_oSer_Comments.WidthMM);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(coords.dWidthMM);
            }
            if (null != coords.dHeightMM) {
                this.memory.WriteByte(c_oSer_Comments.HeightMM);
                this.memory.WriteByte(c_oSerPropLenType.Double);
                this.memory.WriteDouble2(coords.dHeightMM);
            }
            if (null != coords.bMoveWithCells) {
                this.memory.WriteByte(c_oSer_Comments.MoveWithCells);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(coords.bMoveWithCells);
            }
            if (null != coords.bSizeWithCells) {
                this.memory.WriteByte(c_oSer_Comments.SizeWithCells);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(coords.bSizeWithCells);
            }
            if (this.saveThreadedComments && comment.isValidThreadComment()) {
                this.memory.WriteByte(c_oSer_Comments.ThreadedComment);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.bs.WriteItemWithLength(function(){oThis.WriteThreadedComment(comment);});
            }
        };
        this.WriteCommentDatas = function(data)
        {
            var oThis = this;
            this.bs.WriteItem( c_oSer_Comments.CommentData, function(){oThis.WriteCommentData(data);});
        };
        this.WriteCommentData = function(oCommentData)
        {
            var oThis = this;
            var sText = oCommentData.asc_getText();
            if(null != sText)
            {
                this.memory.WriteByte(c_oSer_CommentData.Text);
                this.memory.WriteString2(sText);
            }
            var sTime = oCommentData.asc_getTime();
            if(null != sTime && "" !== sTime)
            {
                this.memory.WriteByte(c_oSer_CommentData.Time);
                this.memory.WriteString2(new Date(sTime - 0).toISOString().slice(0, 19) + 'Z');
            }
            var sOOTime = oCommentData.asc_getOnlyOfficeTime();
            if(null != sOOTime && "" !== sOOTime)
            {
                this.memory.WriteByte(c_oSer_CommentData.OOTime);
                this.memory.WriteString2(new Date(sOOTime - 0).toISOString().slice(0, 19) + 'Z');
            }
            var sUserId = oCommentData.asc_getUserId();
            if(null != sUserId)
            {
                this.memory.WriteByte(c_oSer_CommentData.UserId);
                this.memory.WriteString2(sUserId);
            }
            var sUserName = oCommentData.asc_getUserName();
            if(null != sUserName)
            {
                this.memory.WriteByte(c_oSer_CommentData.UserName);
                this.memory.WriteString2(sUserName);
            }
            var userData = oCommentData.asc_getUserData();
            if(userData)
                this.bs.WriteItem( c_oSer_CommentData.UserData, function(){oThis.memory.WriteString3(userData);});
            var sQuoteText = oCommentData.asc_getQuoteText();
            if(null != sQuoteText)
            {
                this.memory.WriteByte(c_oSer_CommentData.QuoteText);
                this.memory.WriteString2(sQuoteText);
            }
            var bSolved = oCommentData.asc_getSolved();
            if(null != bSolved)
                this.bs.WriteItem( c_oSer_CommentData.Solved, function(){oThis.memory.WriteBool(bSolved);});
            var bDocumentFlag = oCommentData.asc_getDocumentFlag();
            if(null != bDocumentFlag)
                this.bs.WriteItem( c_oSer_CommentData.Document, function(){oThis.memory.WriteBool(bDocumentFlag);});
            var sGuid = oCommentData.asc_getGuid();
            if(null != sGuid){
                this.bs.WriteItem( c_oSer_CommentData.Guid, function(){oThis.memory.WriteString3(sGuid);});
            }
            var aReplies = oCommentData.aReplies;
            if(null != aReplies && aReplies.length > 0)
                this.bs.WriteItem( c_oSer_CommentData.Replies, function(){oThis.WriteReplies(aReplies);});
        };
        this.WriteReplies = function(aReplies)
        {
            var oThis = this;
            for(var i = 0, length = aReplies.length; i < length; ++i)
                this.bs.WriteItem( c_oSer_CommentData.Reply, function(){oThis.WriteCommentData(aReplies[i]);});
        };
        this.WriteThreadedComment = function(oCommentData)
        {
            var oThis = this;
            var i;
            var sOOTime = oCommentData.asc_getOnlyOfficeTime();
            if (sOOTime) {
                this.bs.WriteItem( c_oSer_ThreadedComment.dT, function(){oThis.memory.WriteString3(new Date(sOOTime - 0).toISOString().slice(0, 22) + "Z");});
            }
            var userId = oCommentData.asc_getUserId();
            var displayName = oCommentData.asc_getUserName();
            var providerId = oCommentData.asc_getProviderId();
            var person = this.InitSaveManager.personList.find(function isPrime(element) {
                return userId === element.userId && displayName === element.displayName && providerId === element.providerId;
            });
            if (!person) {
                person = {id: AscCommon.CreateGUID(), userId: userId, displayName: displayName, providerId: providerId};
                this.InitSaveManager.personList.push(person);
            }
            this.bs.WriteItem( c_oSer_ThreadedComment.personId, function(){oThis.memory.WriteString3(person.id);});
            var guid = oCommentData.asc_getGuid();
            if (guid) {
                this.bs.WriteItem( c_oSer_ThreadedComment.id, function(){oThis.memory.WriteString3(guid);});
            }
            var solved = oCommentData.asc_getSolved();
            if (null != solved) {
                this.bs.WriteItem( c_oSer_ThreadedComment.done, function(){oThis.memory.WriteBool(solved);});
            }
            var text = oCommentData.asc_getText();
            if (text) {
                this.bs.WriteItem( c_oSer_ThreadedComment.text, function(){oThis.memory.WriteString3(text);});
            }
            // if (oCommentData.aMentions && oCommentData.aMentions.length > 0) {
            //     for (i = 0; i < oCommentData.aMentions.length; ++i) {
            //         this.bs.WriteItem( c_oSer_ThreadedComment.mention, function(){oThis.WriteThreadedCommentMention(oCommentData.aMentions[i]);});
            //     }
            // }
            if (oCommentData.aReplies && oCommentData.aReplies.length > 0) {
                for (i = 0; i < oCommentData.aReplies.length; ++i) {
                    this.bs.WriteItem( c_oSer_ThreadedComment.reply, function(){oThis.WriteThreadedComment(oCommentData.aReplies[i]);});
                }
            }
        };
        this.WriteThreadedCommentMention = function(mention)
        {
            var oThis = this;
            if(mention.mentionpersonId){
                this.bs.WriteItem( c_oSer_ThreadedComment.mentionpersonId, function(){oThis.memory.WriteString3(mention.mentionpersonId);});
            }
            if(mention.mentionId){
                this.bs.WriteItem( c_oSer_ThreadedComment.mentionId, function(){oThis.memory.WriteString3(mention.mentionId);});
            }
            if(mention.startIndex){
                this.bs.WriteItem( c_oSer_ThreadedComment.startIndex, function(){oThis.memory.WriteULong(mention.startIndex);});
            }
            if(mention.length){
                this.bs.WriteItem( c_oSer_ThreadedComment.length, function(){oThis.memory.WriteULong(mention.length);});
            }
        };
		this.WriteConditionalFormatting = function(oRule)
		{
			var oThis = this;
			if (null != oRule.pivot) {
				this.bs.WriteItem(c_oSer_ConditionalFormatting.Pivot, function() {oThis.memory.WriteBool(oRule.pivot);});
			}
			if (null != oRule.ranges) {
				var sqRef = getSqRefString(oRule.ranges);
				this.bs.WriteItem(c_oSer_ConditionalFormatting.SqRef, function() {oThis.memory.WriteString3(sqRef);});
			}
			this.bs.WriteItem(c_oSer_ConditionalFormatting.ConditionalFormattingRule, function() {oThis.WriteConditionalFormattingRule(oRule);});
		};
		this.WriteConditionalFormattingRule = function(rule) {
			var oThis = this;
			if (null != rule.aboveAverage) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingRule.AboveAverage, function() {oThis.memory.WriteBool(rule.aboveAverage);});
			}
			if (null != rule.bottom) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingRule.Bottom, function() {oThis.memory.WriteBool(rule.bottom);});
			}
			if (null != rule.dxf) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingRule.Dxf, function(){oThis.bsw.WriteDxf(rule.dxf);});
			}
			if (null != rule.equalAverage) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingRule.EqualAverage, function() {oThis.memory.WriteBool(rule.equalAverage);});
			}
			if (null != rule.operator) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingRule.Operator, function() {oThis.memory.WriteByte(rule.operator);});
			}
			if (null != rule.percent) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingRule.Percent, function() {oThis.memory.WriteBool(rule.percent);});
			}
			if (null != rule.priority) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingRule.Priority, function() {oThis.memory.WriteLong(rule.priority);});
			}
			if (null != rule.rank) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingRule.Rank, function() {oThis.memory.WriteLong(rule.rank);});
			}
			if (null != rule.stdDev) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingRule.StdDev, function() {oThis.memory.WriteLong(rule.stdDev);});
			}
			if (null != rule.stopIfTrue) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingRule.StopIfTrue, function() {oThis.memory.WriteBool(rule.stopIfTrue);});
			}
			if (null != rule.text) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingRule.Text, function() {oThis.memory.WriteString3(rule.text);});
			}
			if (null != rule.timePeriod) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingRule.TimePeriod, function() {oThis.memory.WriteString3(rule.timePeriod);});
			}
			if (null != rule.type) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingRule.Type, function() {oThis.memory.WriteByte(rule.type);});
			}
			for (var i = 0; i < rule.aRuleElements.length; ++i) {
				var elem = rule.aRuleElements[i];
				if (elem instanceof AscCommonExcel.CColorScale) {
					this.bs.WriteItem(c_oSer_ConditionalFormattingRule.ColorScale, function() {oThis.WriteColorScale(elem);});
				} else if (elem instanceof AscCommonExcel.CDataBar) {
					this.bs.WriteItem(c_oSer_ConditionalFormattingRule.DataBar, function() {oThis.WriteDataBar(elem);});
				} else if (elem instanceof AscCommonExcel.CFormulaCF) {
					this.bs.WriteItem(c_oSer_ConditionalFormattingRule.FormulaCF, function() {oThis.memory.WriteString3(elem.Text);});
				} else if (elem instanceof AscCommonExcel.CIconSet) {
					this.bs.WriteItem(c_oSer_ConditionalFormattingRule.IconSet, function() {oThis.WriteIconSet(elem);});
				}
			}
		};
		this.WriteColorScale = function(colorScale) {
			var oThis = this;
			var i, elem;
			for (i = 0; i < colorScale.aCFVOs.length; ++i) {
				elem = colorScale.aCFVOs[i];
				this.bs.WriteItem(c_oSer_ConditionalFormattingRuleColorScale.CFVO, function() {oThis.WriteCFVO(elem);});
			}
			for (i = 0; i < colorScale.aColors.length; ++i) {
				elem = colorScale.aColors[i];
				this.bs.WriteItem(c_oSer_ConditionalFormattingRuleColorScale.Color, function() {oThis.bs.WriteColorSpreadsheet(elem);});
			}
		};
		this.WriteDataBar = function(dataBar) {
			var oThis = this;
			var i, elem;
			if (null != dataBar.MaxLength) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.MaxLength, function() {oThis.memory.WriteLong(dataBar.MaxLength);});
			}
			if (null != dataBar.MinLength) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.MinLength, function() {oThis.memory.WriteLong(dataBar.MinLength);});
			}
			if (null != dataBar.ShowValue) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.ShowValue, function() {oThis.memory.WriteBool(dataBar.ShowValue);});
			}
			if (null != dataBar.AxisPosition) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.AxisPosition, function() {oThis.memory.WriteLong(dataBar.AxisPosition);});
			}
			if (null != dataBar.Gradient) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.GradientEnabled, function() {oThis.memory.WriteBool(dataBar.Gradient);});
			}
			if (null != dataBar.Direction) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.Direction, function() {oThis.memory.WriteLong(dataBar.Direction);});
			}
			if (null != dataBar.NegativeBarColorSameAsPositive) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.NegativeBarColorSameAsPositive, function() {oThis.memory.WriteBool(dataBar.NegativeBarColorSameAsPositive);});
			}
			if (null != dataBar.NegativeBarBorderColorSameAsPositive) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.NegativeBarBorderColorSameAsPositive, function() {oThis.memory.WriteBool(dataBar.NegativeBarBorderColorSameAsPositive);});
			}
			if (null != dataBar.Color) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.Color, function() {oThis.bs.WriteColorSpreadsheet(dataBar.Color);});
			}
			if (null != dataBar.NegativeColor) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.NegativeColor, function() {oThis.bs.WriteColorSpreadsheet(dataBar.NegativeColor);});
			}
			if (null != dataBar.BorderColor) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.BorderColor, function() {oThis.bs.WriteColorSpreadsheet(dataBar.BorderColor);});
			}
			if (null != dataBar.NegativeBorderColor) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.NegativeBorderColor, function() {oThis.bs.WriteColorSpreadsheet(dataBar.NegativeBorderColor);});
			}
			if (null != dataBar.AxisColor) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.AxisColor, function() {oThis.bs.WriteColorSpreadsheet(dataBar.AxisColor);});
			}
			for (i = 0; i < dataBar.aCFVOs.length; ++i) {
				elem = dataBar.aCFVOs[i];
				this.bs.WriteItem(c_oSer_ConditionalFormattingDataBar.CFVO, function() {oThis.WriteCFVO(elem);});
			}
		};
		this.WriteIconSet = function(iconSet) {
			var oThis = this;
			var i, elem;
			if (null != iconSet.IconSet) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingIconSet.IconSet, function() {oThis.memory.WriteByte(iconSet.IconSet);});
			}
			if (null != iconSet.Percent) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingIconSet.Percent, function() {oThis.memory.WriteBool(iconSet.Percent);});
			}
			if (null != iconSet.Reverse) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingIconSet.Reverse, function() {oThis.memory.WriteBool(iconSet.Reverse);});
			}
			if (null != iconSet.ShowValue) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingIconSet.ShowValue, function() {oThis.memory.WriteBool(iconSet.ShowValue);});
			}
			for (i = 0; i < iconSet.aCFVOs.length; ++i) {
				elem = iconSet.aCFVOs[i];
				this.bs.WriteItem(c_oSer_ConditionalFormattingIconSet.CFVO, function() {oThis.WriteCFVO(elem);});
			}
			for (i = 0; i < iconSet.aIconSets.length; ++i) {
				elem = iconSet.aIconSets[i];
				this.bs.WriteItem(c_oSer_ConditionalFormattingIconSet.CFIcon, function() {oThis.WriteCFIS(elem);});
			}
		};
		this.WriteCFVO = function(cfvo) {
			var oThis = this;
			if (null != cfvo.Gte) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingValueObject.Gte, function() {oThis.memory.WriteBool(cfvo.Gte);});
			}
			if (null != cfvo.Type) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingValueObject.Type, function() {oThis.memory.WriteByte(cfvo.Type);});
			}
			if (null != cfvo.Val) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingValueObject.Formula, function() {oThis.memory.WriteString3(cfvo.Val);});
			}
		};
		this.WriteCFIS = function(cfis) {
			var oThis = this;
			if (null != cfis.IconSet) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingIcon.iconSet, function() {oThis.memory.WriteLong(cfis.IconSet);});
			}
			if (null != cfis.IconId) {
				this.bs.WriteItem(c_oSer_ConditionalFormattingIcon.iconId, function() {oThis.memory.WriteLong(cfis.IconId);});
			}
		};
		this.WriteSparklineGroups = function(aSparklineGroups)
        {
            var oThis = this;
            for(var i = 0, length = aSparklineGroups.length; i < length; ++i)
                this.bs.WriteItem( c_oSer_Sparkline.SparklineGroup, function(){oThis.WriteSparklineGroup(aSparklineGroups[i]);});
        };
		this.WriteSparklineGroup = function(oSparklineGroup)
        {
			var oThis = this;
			if (null != oSparklineGroup.manualMax) {
                this.bs.WriteItem( c_oSer_Sparkline.ManualMax, function(){oThis.memory.WriteDouble2(oSparklineGroup.manualMax);});
			}
			if (null != oSparklineGroup.manualMin) {
                this.bs.WriteItem( c_oSer_Sparkline.ManualMin, function(){oThis.memory.WriteDouble2(oSparklineGroup.manualMin);});
			}
			if (null != oSparklineGroup.lineWeight) {
                this.bs.WriteItem( c_oSer_Sparkline.LineWeight, function(){oThis.memory.WriteDouble2(oSparklineGroup.lineWeight);});
			}
			if (null != oSparklineGroup.type) {
                this.bs.WriteItem( c_oSer_Sparkline.Type, function(){oThis.memory.WriteByte(oSparklineGroup.type);});
			}
			if (null != oSparklineGroup.dateAxis) {
                this.bs.WriteItem( c_oSer_Sparkline.DateAxis, function(){oThis.memory.WriteBool(oSparklineGroup.dateAxis);});
			}
			if (null != oSparklineGroup.displayEmptyCellsAs) {
                this.bs.WriteItem( c_oSer_Sparkline.DisplayEmptyCellsAs, function(){oThis.memory.WriteByte(oSparklineGroup.displayEmptyCellsAs);});
			}
			if (null != oSparklineGroup.markers) {
                this.bs.WriteItem( c_oSer_Sparkline.Markers, function(){oThis.memory.WriteBool(oSparklineGroup.markers);});
			}
			if (null != oSparklineGroup.high) {
                this.bs.WriteItem( c_oSer_Sparkline.High, function(){oThis.memory.WriteBool(oSparklineGroup.high);});
			}
			if (null != oSparklineGroup.low) {
                this.bs.WriteItem( c_oSer_Sparkline.Low, function(){oThis.memory.WriteBool(oSparklineGroup.low);});
			}
			if (null != oSparklineGroup.first) {
                this.bs.WriteItem( c_oSer_Sparkline.First, function(){oThis.memory.WriteBool(oSparklineGroup.first);});
			}
			if (null != oSparklineGroup.last) {
                this.bs.WriteItem( c_oSer_Sparkline.Last, function(){oThis.memory.WriteBool(oSparklineGroup.last);});
			}
			if (null != oSparklineGroup.negative) {
                this.bs.WriteItem( c_oSer_Sparkline.Negative, function(){oThis.memory.WriteBool(oSparklineGroup.negative);});
			}
			if (null != oSparklineGroup.displayXAxis) {
                this.bs.WriteItem( c_oSer_Sparkline.DisplayXAxis, function(){oThis.memory.WriteBool(oSparklineGroup.displayXAxis);});
			}
			if (null != oSparklineGroup.displayHidden) {
                this.bs.WriteItem( c_oSer_Sparkline.DisplayHidden, function(){oThis.memory.WriteBool(oSparklineGroup.displayHidden);});
			}
			if (null != oSparklineGroup.minAxisType) {
                this.bs.WriteItem( c_oSer_Sparkline.MinAxisType, function(){oThis.memory.WriteByte(oSparklineGroup.minAxisType);});
			}
			if (null != oSparklineGroup.maxAxisType) {
                this.bs.WriteItem( c_oSer_Sparkline.MaxAxisType, function(){oThis.memory.WriteByte(oSparklineGroup.maxAxisType);});
			}
			if (null != oSparklineGroup.rightToLeft) {
                this.bs.WriteItem( c_oSer_Sparkline.RightToLeft, function(){oThis.memory.WriteBool(oSparklineGroup.rightToLeft);});
			}
			if (null != oSparklineGroup.colorSeries) {
                this.bs.WriteItem(c_oSer_Sparkline.ColorSeries, function(){oThis.bs.WriteColorSpreadsheet(oSparklineGroup.colorSeries);});
			}
			if (null != oSparklineGroup.colorNegative) {
                this.bs.WriteItem(c_oSer_Sparkline.ColorNegative, function(){oThis.bs.WriteColorSpreadsheet(oSparklineGroup.colorNegative);});
			}
			if (null != oSparklineGroup.colorAxis) {
                this.bs.WriteItem(c_oSer_Sparkline.ColorAxis, function(){oThis.bs.WriteColorSpreadsheet(oSparklineGroup.colorAxis);});
			}
			if (null != oSparklineGroup.colorMarkers) {
                this.bs.WriteItem(c_oSer_Sparkline.ColorMarkers, function(){oThis.bs.WriteColorSpreadsheet(oSparklineGroup.colorMarkers);});
			}
			if (null != oSparklineGroup.colorFirst) {
                this.bs.WriteItem(c_oSer_Sparkline.ColorFirst, function(){oThis.bs.WriteColorSpreadsheet(oSparklineGroup.colorFirst);});
			}
			if (null != oSparklineGroup.colorLast) {
                this.bs.WriteItem(c_oSer_Sparkline.ColorLast, function(){oThis.bs.WriteColorSpreadsheet(oSparklineGroup.colorLast);});
			}
			if (null != oSparklineGroup.colorHigh) {
                this.bs.WriteItem(c_oSer_Sparkline.ColorHigh, function(){oThis.bs.WriteColorSpreadsheet(oSparklineGroup.colorHigh);});
			}
			if (null != oSparklineGroup.colorLow) {
                this.bs.WriteItem(c_oSer_Sparkline.ColorLow, function(){oThis.bs.WriteColorSpreadsheet(oSparklineGroup.colorLow);});
			}
			if (null != oSparklineGroup.f) {
                this.memory.WriteByte(c_oSer_Sparkline.Ref);
                this.memory.WriteString2(oSparklineGroup.f);
			}
			if (null != oSparklineGroup.arrSparklines) {
				this.bs.WriteItem(c_oSer_Sparkline.Sparklines, function(){oThis.WriteSparklines(oSparklineGroup);});
			}
		};
		this.WriteSparklines = function(oSparklineGroup)
        {
            var oThis = this;
            for(var i = 0, length = oSparklineGroup.arrSparklines.length; i < length; ++i)
                this.bs.WriteItem( c_oSer_Sparkline.Sparkline, function(){oThis.WriteSparkline(oSparklineGroup.arrSparklines[i]);});
        };
		this.WriteSparkline = function(oSparkline)
        {
			if (null != oSparkline.f) {
                this.memory.WriteByte(c_oSer_Sparkline.SparklineRef);
                this.memory.WriteString2(oSparkline.f);
			}
			if (null != oSparkline.sqRef) {
				this.memory.WriteByte(c_oSer_Sparkline.SparklineSqRef);
                this.memory.WriteString2(oSparkline.sqRef.getName());
			}
		}
        this.WriteSlicers = function(slicers)
        {
            var oThis = this;
            this.bs.WriteItem(c_oSerWorksheetsTypes.Slicer, function() {
                var old = new AscCommon.CMemory(true);
                pptx_content_writer.BinaryFileWriter.ExportToMemory(old);
                pptx_content_writer.BinaryFileWriter.ImportFromMemory(oThis.memory);
                pptx_content_writer.BinaryFileWriter.WriteRecord4(0, slicers);
                pptx_content_writer.BinaryFileWriter.ExportToMemory(oThis.memory);
                pptx_content_writer.BinaryFileWriter.ImportFromMemory(old);
            });
        };
		this.WritePivotTable = function(pivotTable, isCopyPaste)
		{
			if (isCopyPaste && !pivotTable.isInRange(this.isCopyPaste)) {
				return;
			}

			var oThis = this;
			if (null != pivotTable.cacheId) {
				this.bs.WriteItem(c_oSer_PivotTypes.cacheId, function() {oThis.memory.WriteLong(pivotTable.cacheId);});
			}
			var stylesForWrite = oThis.isCopyPaste ? undefined : oThis.stylesForWrite;
			var dxfs = oThis.isCopyPaste ? undefined : this.InitSaveManager.getDxfs();
			this.bs.WriteItem(c_oSer_PivotTypes.table, function() {pivotTable.toXml(oThis.memory, stylesForWrite, dxfs);});
		};
        this.WriteHeaderFooter = function(headerFooter)
        {
            var oThis = this;
            if (null !== headerFooter.alignWithMargins) {
                this.bs.WriteItem(c_oSer_HeaderFooter.AlignWithMargins, function() {oThis.memory.WriteBool(headerFooter.alignWithMargins);});
            }
            if (null !== headerFooter.differentFirst) {
                this.bs.WriteItem(c_oSer_HeaderFooter.DifferentFirst, function() {oThis.memory.WriteBool(headerFooter.differentFirst);});
            }
            if (null !== headerFooter.differentOddEven) {
                this.bs.WriteItem(c_oSer_HeaderFooter.DifferentOddEven, function() {oThis.memory.WriteBool(headerFooter.differentOddEven);});
            }
            if (true !== headerFooter.scaleWithDoc) {
                this.bs.WriteItem(c_oSer_HeaderFooter.ScaleWithDoc, function() {oThis.memory.WriteBool(headerFooter.scaleWithDoc);});
            }
            if (null !== headerFooter.evenFooter) {
                this.memory.WriteByte(c_oSer_HeaderFooter.EvenFooter);
                this.memory.WriteString2(headerFooter.evenFooter.getStr());
            }
            if (null !== headerFooter.evenHeader) {
                this.memory.WriteByte(c_oSer_HeaderFooter.EvenHeader);
                this.memory.WriteString2(headerFooter.evenHeader.getStr());
            }
            if (null !== headerFooter.firstFooter) {
                this.memory.WriteByte(c_oSer_HeaderFooter.FirstFooter);
                this.memory.WriteString2(headerFooter.firstFooter.getStr());
            }
            if (null !== headerFooter.firstHeader) {
                this.memory.WriteByte(c_oSer_HeaderFooter.FirstHeader);
                this.memory.WriteString2(headerFooter.firstHeader.getStr());
            }
            if (null !== headerFooter.oddFooter) {
                this.memory.WriteByte(c_oSer_HeaderFooter.OddFooter);
                this.memory.WriteString2(headerFooter.oddFooter.getStr());
            }
            if (null !== headerFooter.oddHeader) {
                this.memory.WriteByte(c_oSer_HeaderFooter.OddHeader);
                this.memory.WriteString2(headerFooter.oddHeader.getStr());
            }
        };
        this.WriteRowColBreaks = function(breaks)
        {
            var oThis = this;
            if (null !== breaks.count) {
                this.bs.WriteItem(c_oSer_RowColBreaks.Count, function() {oThis.memory.WriteLong(breaks.count);});
            }
            if (null !== breaks.manualBreakCount) {
                this.bs.WriteItem(c_oSer_RowColBreaks.ManualBreakCount, function() {oThis.memory.WriteLong(breaks.manualBreakCount);});
            }
            for(var i = 0; i < breaks.breaks.length; ++i){
                this.bs.WriteItem(c_oSer_RowColBreaks.Break, function() {oThis.WriteRowColBreak(breaks.breaks[i]);});
            }
        };
        this.WriteRowColBreak = function(brk)
        {
            var oThis = this;
            if (null !== brk.id) {
                this.bs.WriteItem(c_oSer_RowColBreaks.Id, function() {oThis.memory.WriteLong(brk.id);});
            }
            if (null !== brk.man) {
                this.bs.WriteItem(c_oSer_RowColBreaks.Man, function() {oThis.memory.WriteBool(brk.man);});
            }
            if (null !== brk.max) {
                this.bs.WriteItem(c_oSer_RowColBreaks.Max, function() {oThis.memory.WriteLong(brk.max);});
            }
            if (null !== brk.min) {
                this.bs.WriteItem(c_oSer_RowColBreaks.Min, function() {oThis.memory.WriteLong(brk.min);});
            }
            if (null !== brk.pt) {
                this.bs.WriteItem(c_oSer_RowColBreaks.Pt, function() {oThis.memory.WriteBool(brk.pt);});
            }
        };
        this.WriteLegacyDrawingHF = function(legacyDrawingHF)
        {
            var oThis = this;
            if (null !== legacyDrawingHF.cfe) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Cfe, function() {oThis.memory.WriteLong(legacyDrawingHF.cfe);});
            }
            if (null !== legacyDrawingHF.cff) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Cff, function() {oThis.memory.WriteLong(legacyDrawingHF.cff);});
            }
            if (null !== legacyDrawingHF.cfo) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Cfo, function() {oThis.memory.WriteLong(legacyDrawingHF.cfo);});
            }
            if (null !== legacyDrawingHF.che) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Che, function() {oThis.memory.WriteLong(legacyDrawingHF.che);});
            }
            if (null !== legacyDrawingHF.chf) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Chf, function() {oThis.memory.WriteLong(legacyDrawingHF.chf);});
            }
            if (null !== legacyDrawingHF.cho) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Cho, function() {oThis.memory.WriteLong(legacyDrawingHF.cho);});
            }
            if (null !== legacyDrawingHF.lfe) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Lfe, function() {oThis.memory.WriteLong(legacyDrawingHF.lfe);});
            }
            if (null !== legacyDrawingHF.lff) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Lff, function() {oThis.memory.WriteLong(legacyDrawingHF.lff);});
            }
            if (null !== legacyDrawingHF.lfo) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Lfo, function() {oThis.memory.WriteLong(legacyDrawingHF.lfo);});
            }
            if (null !== legacyDrawingHF.lhe) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Lhe, function() {oThis.memory.WriteLong(legacyDrawingHF.lhe);});
            }
            if (null !== legacyDrawingHF.lhf) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Lhf, function() {oThis.memory.WriteLong(legacyDrawingHF.lhf);});
            }
            if (null !== legacyDrawingHF.lho) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Lho, function() {oThis.memory.WriteLong(legacyDrawingHF.lho);});
            }
            if (null !== legacyDrawingHF.rfe) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Rfe, function() {oThis.memory.WriteLong(legacyDrawingHF.rfe);});
            }
            if (null !== legacyDrawingHF.rff) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Rff, function() {oThis.memory.WriteLong(legacyDrawingHF.rff);});
            }
            if (null !== legacyDrawingHF.rfo) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Rfo, function() {oThis.memory.WriteLong(legacyDrawingHF.rfo);});
            }
            if (null !== legacyDrawingHF.rhe) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Rhe, function() {oThis.memory.WriteLong(legacyDrawingHF.rhe);});
            }
            if (null !== legacyDrawingHF.rhf) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Rhf, function() {oThis.memory.WriteLong(legacyDrawingHF.rhf);});
            }
            if (null !== legacyDrawingHF.rho) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Rho, function() {oThis.memory.WriteLong(legacyDrawingHF.rho);});
            }
            this.bs.WriteItem(c_oSer_LegacyDrawingHF.Drawings, function() {oThis.WriteLegacyDrawingHFDrawings(legacyDrawingHF.drawings);});

        };
        this.WriteLegacyDrawingHFDrawings = function(drawings) {
            var oThis = this;
            for (var i = 0; i < drawings.length; ++i) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.Drawing, function() {oThis.WriteLegacyDrawingHFDrawing(drawings[i]);});
            }
        };
        this.WriteLegacyDrawingHFDrawing = function(drawing) {
            var oThis = this;
            if (null !== drawing.id) {
                this.memory.WriteByte(c_oSer_LegacyDrawingHF.DrawingId);
                this.memory.WriteString2(drawing.id);
            }
            if (null !== drawing.graphicObject) {
                this.bs.WriteItem(c_oSer_LegacyDrawingHF.DrawingShape, function(){pptx_content_writer.WriteDrawing(oThis.memory, drawing.graphicObject, null, null, null);});
            }
        };

        this.WriteUserProtectedRanges = function (aUserProtectedRanges) {
            var oThis = this;
            for (var i = 0, length = aUserProtectedRanges.length; i < length; ++i) {
                this.bs.WriteItem(c_oSerUserProtectedRange.UserProtectedRange, function () {
                    oThis.WriteUserProtectedRange(aUserProtectedRanges[i]);
                });
            }
        };
		this.WriteUserProtectedRangeDesc = function (desc) {
			if (desc.id) {
				this.memory.WriteByte(c_oSerUserProtectedRangeDesc.Id);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(desc.id);
			}
			if (desc.name) {
				this.memory.WriteByte(c_oSerUserProtectedRangeDesc.Name);
				this.memory.WriteByte(c_oSerPropLenType.Variable);
				this.memory.WriteString2(desc.name);
			}
			if (null != desc.type) {
				this.memory.WriteByte(c_oSerUserProtectedRangeDesc.Type);
				this.memory.WriteByte(c_oSerPropLenType.Byte);
				this.memory.WriteByte(desc.type);
			}
		};
        this.WriteUserProtectedRange = function (oUserProtectedRange) {

            var oThis = this;

            if (oUserProtectedRange.name) {
				this.bs.WriteItem(c_oSerUserProtectedRange.Name, function () {
                    oThis.memory.WriteString3(oUserProtectedRange.name);
                });
            }
            if (oUserProtectedRange.ref) {
				var sqRef = getSqRefString([oUserProtectedRange.ref]);
            	this.bs.WriteItem(c_oSerUserProtectedRange.Sqref, function () {
                    oThis.memory.WriteString3(sqRef);
                });
            }
            if (oUserProtectedRange.warningText) {
                this.bs.WriteItem(c_oSerUserProtectedRange.Text, function () {
                    oThis.memory.WriteString3(oUserProtectedRange.warningText);
                });
            }

            let i;
            let users = oUserProtectedRange.asc_getUsers();
            if (null != users) {
                for (i = 0; i < users.length; i++) {
                    this.bs.WriteItem(c_oSerUserProtectedRange.User, function () {
                        oThis.WriteUserProtectedRangeDesc(users[i]);
                    });
                }
            }
			let userGroups = oUserProtectedRange.asc_getUserGroups();
            if (null != userGroups) {
				for (i = 0; i < userGroups.length; i++) {
					this.bs.WriteItem(c_oSerUserProtectedRange.UsersGroup, function () {
						oThis.WriteUserProtectedRangeDesc(userGroups[i]);
					});
				}
            }
            if (null != oUserProtectedRange.type) {
                this.bs.WriteItem(c_oSerUserProtectedRange.Type, function () {
                    oThis.memory.WriteByte(oUserProtectedRange.type);
                });
            }
        };

        this.WriteTimelines = function (aTimelines) {
            var oThis = this;
            if (aTimelines.length > 0) {
                this.bs.WriteItem(c_oSerWorksheetsTypes.Timelines, function () {
                    for (var i = 0; i < aTimelines.length; ++i) {
                        oThis.bs.WriteItem(c_oSerWorksheetsTypes.Timeline, function () {oThis.WriteTimeline(aTimelines[i]);});
                    }
                });
            }
        };
        this.WriteTimeline = function (oTimeline) {

            if (!oTimeline) {
                return;
            }

            if (oTimeline.name != null) {
                this.memory.WriteByte(c_oSer_Timeline.Name);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(oTimeline.name);
            }
            if (oTimeline.caption != null) {
                this.memory.WriteByte(c_oSer_Timeline.Caption);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(oTimeline.caption);
            }
            if (oTimeline.uid != null) {
                this.memory.WriteByte(c_oSer_Timeline.Uid);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(oTimeline.uid);
            }
            if (oTimeline.scrollPosition != null) {
                this.memory.WriteByte(c_oSer_Timeline.ScrollPosition);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(oTimeline.scrollPosition);
            }
            if (oTimeline.cache != null) {
                this.memory.WriteByte(c_oSer_Timeline.Cache);
                this.memory.WriteByte(c_oSerPropLenType.Variable);
                this.memory.WriteString2(oTimeline.cache);
            }
            if (oTimeline.selectionLevel != null) {
                this.memory.WriteByte(c_oSer_Timeline.SelectionLevel);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oTimeline.selectionLevel);
            }
            if (oTimeline.level != null) {
                this.memory.WriteByte(c_oSer_Timeline.Level);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oTimeline.level);
            }
            if (oTimeline.showHeader != null) {
                this.memory.WriteByte(c_oSer_Timeline.ShowHeader);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(oTimeline.showHeader);
            }
            if (oTimeline.showSelectionLabel != null) {
                this.memory.WriteByte(c_oSer_Timeline.ShowSelectionLabel);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(oTimeline.showSelectionLabel);
            }
            if (oTimeline.showTimeLevel != null) {
                this.memory.WriteByte(c_oSer_Timeline.ShowTimeLevel);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(oTimeline.showTimeLevel);
            }
            if (oTimeline.showHorizontalScrollbar != null) {
                this.memory.WriteByte(c_oSer_Timeline.ShowHorizontalScrollbar);
                this.memory.WriteByte(c_oSerPropLenType.Byte);
                this.memory.WriteBool(oTimeline.showHorizontalScrollbar);
            }
            if (oTimeline.style != null) {
                this.memory.WriteByte(c_oSer_Timeline.Style);
                this.memory.WriteByte(c_oSerPropLenType.Long);
                this.memory.WriteLong(oTimeline.style);
            }
        };
    }

	/** @constructor */
	function BinaryOtherTableWriter(memory, wb)
	{
		this.memory = memory;
		this.wb = wb;
		this.bs = new BinaryCommonWriter(this.memory);
		this.Write = function()
		{
			var oThis = this;
			this.bs.WriteItemWithLength(function(){oThis.WriteOtherContent();});
		};
		this.WriteOtherContent = function()
		{
			var oThis = this;
			this.bs.WriteItem(c_oSer_OtherType.Theme, function(){pptx_content_writer.WriteTheme(oThis.memory, oThis.wb.theme);});
		};
	}
    function BinaryPersonTableWriter(memory, personList)
    {
        this.memory = memory;
        this.personList = personList;
        this.bs = new BinaryCommonWriter(this.memory);
        this.Write = function()
        {
            var oThis = this;
            this.bs.WriteItemWithLength(function(){oThis.WritePersonList();});
        };
        this.WritePersonList = function()
        {
            var oThis = this;
            for (var i = 0; i < this.personList.length; ++i) {
                this.bs.WriteItem(c_oSer_Person.person, function(){oThis.WritePerson(oThis.personList[i]);});
            }
        };
        this.WritePerson = function(person)
        {
            var oThis = this;
            if (person.id) {
                this.bs.WriteItem(c_oSer_Person.id, function(){oThis.memory.WriteString3(person.id);});
            }
            if (person.userId && person.providerId) {
                this.bs.WriteItem(c_oSer_Person.userId, function(){oThis.memory.WriteString3(person.userId);});
                this.bs.WriteItem(c_oSer_Person.providerId, function(){oThis.memory.WriteString3(person.providerId);});
            }
            if (person.displayName) {
                this.bs.WriteItem(c_oSer_Person.displayName, function(){oThis.memory.WriteString3(person.displayName);});
            }
        };
    }
    /** @constructor */
    function BinaryFileWriter(wb, isCopyPaste)
    {
        this.Memory = new AscCommon.CMemory();
        this.wb = wb;
        this.isCopyPaste = isCopyPaste;
        this.saveThreadedComments = true;
        this.nLastFilePos = 0;
        this.nRealTableCount = 0;
        this.InitSaveManager = new InitSaveManager(wb, isCopyPaste);
        this.bs = new BinaryCommonWriter(this.Memory);
        this.Write = function(noBase64, onlySaveBase64)
        {
            var t = this;
            pptx_content_writer._Start();
			if (noBase64) {
				this.Memory.WriteXmlString(this.WriteFileHeader(0, Asc.c_nVersionNoBase64));
			}
			AscCommonExcel.executeInR1C1Mode(false, function () {
				t.WriteMainTable();
			});
            pptx_content_writer._End();
			if (noBase64) {
			    if (onlySaveBase64)
                    return this.Memory.GetBase64Memory();
			    else
			        return this.Memory.GetData();
			} else {
				return this.WriteFileHeader(this.Memory.GetCurPosition(), AscCommon.c_oSerFormat.Version) + this.Memory.GetBase64Memory();
			}
        };
        this.WriteFileHeader = function(nDataSize, version)
        {
            return AscCommon.c_oSerFormat.Signature + ";v" + version + ";" + nDataSize  + ";";
        };
        this.WriteMainTable = function()
        {
            var t = this;
            var nTableCount = 128;//Специально ставим большое число, чтобы не увеличивать его при добавлении очередной таблицы.
            this.nRealTableCount = 0;//Специально ставим большое число, чтобы не увеличивать его при добавлении очередной таблицы.
            var nStart = this.Memory.GetCurPosition();
            //вычисляем с какой позиции можно писать таблицы
            var nmtItemSize = 5;//5 byte
            this.nLastFilePos = nStart + nTableCount * nmtItemSize;
            //Write mtLen
            this.Memory.WriteByte(0);
            if (this.wb.App) {
                this.WriteTable(c_oSerTableTypes.App, {Write: function(){
                    var old = new AscCommon.CMemory(true);
                    pptx_content_writer.BinaryFileWriter.ExportToMemory(old);
                    pptx_content_writer.BinaryFileWriter.ImportFromMemory(t.Memory);
                    t.wb.App.toStream(pptx_content_writer.BinaryFileWriter);
                    pptx_content_writer.BinaryFileWriter.ExportToMemory(t.Memory);
                    pptx_content_writer.BinaryFileWriter.ImportFromMemory(old);
                }});
            }
            if (this.wb.Core) {
                this.WriteTable(c_oSerTableTypes.Core, {Write: function(){
                    var old = new AscCommon.CMemory(true);
                    pptx_content_writer.BinaryFileWriter.ExportToMemory(old);
                    pptx_content_writer.BinaryFileWriter.ImportFromMemory(t.Memory);
                    t.wb.Core.toStream(pptx_content_writer.BinaryFileWriter, t.wb.oApi);
                    pptx_content_writer.BinaryFileWriter.ExportToMemory(t.Memory);
                    pptx_content_writer.BinaryFileWriter.ImportFromMemory(old);
                }});
            }
            if (this.wb.CustomProperties && this.wb.CustomProperties.hasProperties()) {
                this.WriteTable(c_oSerTableTypes.CustomProperties, {Write: function(){
                    var old = new AscCommon.CMemory(true);
                    pptx_content_writer.BinaryFileWriter.ExportToMemory(old);
                    pptx_content_writer.BinaryFileWriter.ImportFromMemory(t.Memory);
                    t.wb.CustomProperties.toStream(pptx_content_writer.BinaryFileWriter);
                    pptx_content_writer.BinaryFileWriter.ExportToMemory(t.Memory);
                    pptx_content_writer.BinaryFileWriter.ImportFromMemory(old);
                }});
            }
            if (t.wb.customXmls && t.wb.customXmls.length > 0) {
                this.WriteTable(c_oSerTableTypes.Customs, new BinaryCustomsTableWriter(this.Memory, t.wb.customXmls));
            }

            //var oSharedStrings = {index: 0, strings: {}};
            //Write SharedStrings
            var nSharedStringsPos = this.ReserveTable(c_oSerTableTypes.SharedStrings);
            //Write Styles
            var nStylesTablePos = this.ReserveTable(c_oSerTableTypes.Styles);
            //Workbook
            //var personList = [];
            //var commentUniqueGuids = {};
            /*var tableIds = {};
            var sheetIds = {};*/
			var oBinaryStylesTableWriter = new BinaryStylesTableWriter(this.Memory, this.wb, this.InitSaveManager);
			var oBinaryWorksheetsTableWriter = new BinaryWorksheetsTableWriter(this.Memory, this.wb, this.isCopyPaste, oBinaryStylesTableWriter, this.saveThreadedComments, this.InitSaveManager);
            this.WriteTable(c_oSerTableTypes.Workbook, new BinaryWorkbookTableWriter(this.Memory, this.wb, oBinaryWorksheetsTableWriter, this.isCopyPaste, this.InitSaveManager));
            //Worksheets
            this.WriteTable(c_oSerTableTypes.Worksheets, oBinaryWorksheetsTableWriter);
            if (this.InitSaveManager.personList.length > 0) {
                this.WriteTable(c_oSerTableTypes.PersonList, new BinaryPersonTableWriter(this.Memory, this.InitSaveManager.personList));
            }
			if(!this.isCopyPaste)
				this.WriteTable(c_oSerTableTypes.Other, new BinaryOtherTableWriter(this.Memory, this.wb));
            //Write SharedStrings
			this.WriteReserved(new BinarySharedStringsTableWriter(this.Memory, this.wb, oBinaryStylesTableWriter, this.InitSaveManager), nSharedStringsPos);
            //Write Styles
			this.WriteReserved(oBinaryStylesTableWriter, nStylesTablePos);
            //Пишем количество таблиц
            this.Memory.Seek(nStart);
            this.Memory.WriteByte(this.nRealTableCount);

            //seek в конец, потому что GetBase64Memory заканчивает запись на текущей позиции.
            this.Memory.Seek(this.nLastFilePos);
        };
        this.WriteTable = function(type, oTableSer)
        {
            //Write mtItem
            //Write mtiType
            this.Memory.WriteByte(type);
            //Write mtiOffBits
            this.Memory.WriteLong(this.nLastFilePos);

            //Write table
            //Запоминаем позицию в MainTable
            var nCurPos = this.Memory.GetCurPosition();
            //Seek в свободную область
            this.Memory.Seek(this.nLastFilePos);
            oTableSer.Write();
            //сдвигаем позицию куда можно следующую таблицу
            this.nLastFilePos = this.Memory.GetCurPosition();
            //Seek вобратно в MainTable
            this.Memory.Seek(nCurPos);

            this.nRealTableCount++;
        };
        this.ReserveTable = function(type)
        {
            var res = 0;
            //Write mtItem
            //Write mtiType
            this.Memory.WriteByte(type);
            res = this.Memory.GetCurPosition();
            //Write mtiOffBits
            this.Memory.WriteLong(this.nLastFilePos);
            return res;
        };
        this.WriteReserved = function(oTableSer, nPos)
        {
            this.Memory.Seek(nPos);
            this.Memory.WriteLong(this.nLastFilePos);

            //Write table
            //Запоминаем позицию в MainTable
            var nCurPos = this.Memory.GetCurPosition();
            //Seek в свободную область
            this.Memory.Seek(this.nLastFilePos);
            oTableSer.Write();
            //сдвигаем позицию куда можно следующую таблицу
            this.nLastFilePos = this.Memory.GetCurPosition();
            //Seek вобратно в MainTable
            this.Memory.Seek(nCurPos);

            this.nRealTableCount++;
        };
    }
    /** @constructor */
    function Binary_TableReader(stream, initOpenManager, ws)
    {
        this.stream = stream;
        this.ws = ws;
        this.bcr = new Binary_CommonReader(this.stream);
        this.initOpenManager = initOpenManager;
        this.Read = function(length, aTables)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            res = this.bcr.Read1(length, function(t,l){
                return oThis.ReadTables(t,l, aTables);
            });
            return res;
        };
        this.ReadTables = function(type, length, aTables)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_TablePart.Table == type )
            {
                var oNewTable = this.ws.createTablePart();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadTable(t,l, oNewTable);
                });
                if(null != oNewTable.Ref && null != oNewTable.DisplayName)
                    this.ws.workbook.dependencyFormulas.addTableName(this.ws, oNewTable, true);
                aTables.push(oNewTable);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadTable = function(type, length, oTable)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_TablePart.Ref == type )
                oTable.Ref = AscCommonExcel.g_oRangeCache.getAscRange(this.stream.GetString2LE(length));
            else if ( c_oSer_TablePart.HeaderRowCount == type )
                oTable.HeaderRowCount = this.stream.GetULongLE();
            else if ( c_oSer_TablePart.TotalsRowCount == type )
                oTable.TotalsRowCount = this.stream.GetULongLE();
            else if ( c_oSer_TablePart.DisplayName == type )
                oTable.DisplayName = this.stream.GetString2LE(length);
            else if ( c_oSer_TablePart.AutoFilter == type )
            {
                oTable.AutoFilter = new AscCommonExcel.AutoFilter();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadAutoFilter(t,l, oTable.AutoFilter);
                });
                if(!oTable.AutoFilter.Ref) {
					oTable.AutoFilter.Ref = oTable.generateAutoFilterRef();
                }
            }
            else if ( c_oSer_TablePart.SortState == type )
            {
                oTable.SortState = new AscCommonExcel.SortState();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadSortState(t,l, oTable.SortState);
                });
            }
            else if ( c_oSer_TablePart.TableColumns == type )
            {
                oTable.TableColumns = [];
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadTableColumns(t,l, oTable.TableColumns);
                });
            }
            else if ( c_oSer_TablePart.TableStyleInfo == type )
            {
                oTable.TableStyleInfo = new AscCommonExcel.TableStyleInfo();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadTableStyleInfo(t,l, oTable.TableStyleInfo);
                });
            }
			else if ( c_oSer_TablePart.AltTextTable == type )
			{
				res = this.bcr.Read1(length, function(t,l){
					return oThis.ReadAltTextTable(t,l, oTable);
				});
			}
            else if ( c_oSer_TablePart.Id == type )
            {
                this.initOpenManager.oReadResult.tableIds[this.stream.GetULongLE()] = oTable;
            }
			else if ( c_oSer_TablePart.QueryTable == type )
			{
				oTable.QueryTable = new AscCommonExcel.QueryTable();
				res = this.bcr.Read1(length, function(t,l){
					return oThis.ReadQueryTable(t, l, oTable.QueryTable);
				});
			}
			else if (c_oSer_TablePart.TableType == type) {
				oTable.tableType = this.stream.GetULongLE();
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadAltTextTable = function(type, length, oTable)
		{
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			if (c_oSer_AltTextTable.AltText == type) {
				oTable.altText = this.stream.GetString2LE(length);
			} else if ( c_oSer_AltTextTable.AltTextSummary == type ) {
				oTable.altTextSummary  = this.stream.GetString2LE(length);
			} else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
        this.ReadAutoFilter = function(type, length, oAutoFilter)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_AutoFilter.Ref == type )
			{
				oAutoFilter.setStringRef(this.stream.GetString2LE(length));
			}
            else if ( c_oSer_AutoFilter.FilterColumns == type )
            {
                oAutoFilter.FilterColumns = [];
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadFilterColumns(t,l, oAutoFilter.FilterColumns);
                });
            }
            else if ( c_oSer_AutoFilter.SortState == type )
            {
                oAutoFilter.SortState = new AscCommonExcel.SortState();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadSortState(t,l, oAutoFilter.SortState);
                });
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadFilterColumns = function(type, length, aFilterColumns)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_AutoFilter.FilterColumn == type )
            {
                var oFilterColumn = new AscCommonExcel.FilterColumn();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadFilterColumn(t,l, oFilterColumn);
                });
                aFilterColumns.push(oFilterColumn);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadFilterColumn = function(type, length, oFilterColumn)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_FilterColumn.ColId == type )
                oFilterColumn.ColId = this.stream.GetULongLE();
            else if ( c_oSer_FilterColumn.Filters == type )
            {
                oFilterColumn.Filters = new AscCommonExcel.Filters();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadFilters(t,l, oFilterColumn.Filters);
                });
				oFilterColumn.Filters.sortDate();
            }
            else if ( c_oSer_FilterColumn.CustomFilters == type )
            {
                oFilterColumn.CustomFiltersObj = new Asc.CustomFilters();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadCustomFilters(t,l, oFilterColumn.CustomFiltersObj);
                });
            }
            else if ( c_oSer_FilterColumn.DynamicFilter == type )
            {
                oFilterColumn.DynamicFilter = new Asc.DynamicFilter();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadDynamicFilter(t,l, oFilterColumn.DynamicFilter);
                });
            }else if ( c_oSer_FilterColumn.ColorFilter == type )
            {
                oFilterColumn.ColorFilter = new Asc.ColorFilter();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadColorFilter(t,l, oFilterColumn.ColorFilter);
                });
            }
            else if ( c_oSer_FilterColumn.Top10 == type )
            {
                oFilterColumn.Top10 = new Asc.Top10();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadTop10(t,l, oFilterColumn.Top10);
                });
            }
            else if ( c_oSer_FilterColumn.HiddenButton == type )
                oFilterColumn.ShowButton = !this.stream.GetBool();
            else if ( c_oSer_FilterColumn.ShowButton == type )
                oFilterColumn.ShowButton = this.stream.GetBool();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadFilterColumnExternal = function()
		{
			var oThis = this;
			var oFilterColumn = new AscCommonExcel.FilterColumn();
			var length = this.stream.GetULongLE();
			var res = this.bcr.Read1(length, function(t,l){
				return oThis.ReadFilterColumn(t,l, oFilterColumn);
			});
			return oFilterColumn;
		};
        this.ReadFilters = function(type, length, oFilters)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_FilterColumn.Filter == type )
            {
                var oFilterVal = new AscCommonExcel.Filter();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadFilter(t,l, oFilterVal);
                });
                if(null != oFilterVal.Val)
					oFilters.Values[oFilterVal.Val] = 1;
            }
            else if ( c_oSer_FilterColumn.DateGroupItem == type )
            {
                var oDateGroupItem = new AscCommonExcel.DateGroupItem();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadDateGroupItem(t,l, oDateGroupItem);
                });

				var autoFilterDateElem = new AscCommonExcel.AutoFilterDateElem();
				autoFilterDateElem.convertDateGroupItemToRange(oDateGroupItem);
				oFilters.Dates.push(autoFilterDateElem);
            }
            else if ( c_oSer_FilterColumn.FiltersBlank == type )
                oFilters.Blank = this.stream.GetBool();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadFilter = function(type, length, oFilter)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_Filter.Val == type )
                oFilter.Val = this.stream.GetString2LE(length);
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadDateGroupItem = function(type, length, oDateGroupItem)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_DateGroupItem.DateTimeGrouping == type )
                oDateGroupItem.DateTimeGrouping = this.stream.GetUChar();
            else if ( c_oSer_DateGroupItem.Day == type )
                oDateGroupItem.Day = this.stream.GetULongLE();
            else if ( c_oSer_DateGroupItem.Hour == type )
                oDateGroupItem.Hour = this.stream.GetULongLE();
            else if ( c_oSer_DateGroupItem.Minute == type )
                oDateGroupItem.Minute = this.stream.GetULongLE();
            else if ( c_oSer_DateGroupItem.Month == type )
                oDateGroupItem.Month = this.stream.GetULongLE();
            else if ( c_oSer_DateGroupItem.Second == type )
                oDateGroupItem.Second = this.stream.GetULongLE();
            else if ( c_oSer_DateGroupItem.Year == type )
                oDateGroupItem.Year = this.stream.GetULongLE();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadCustomFilters = function(type, length, oCustomFilters)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_CustomFilters.And == type )
                oCustomFilters.And = this.stream.GetBool();
            else if ( c_oSer_CustomFilters.CustomFilters == type )
            {
                oCustomFilters.CustomFilters = [];
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadCustomFiltersItems(t,l, oCustomFilters.CustomFilters);
                });
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadCustomFiltersItems = function(type, length, aCustomFilters)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_CustomFilters.CustomFilter == type )
            {
                var oCustomFiltersItem = new Asc.CustomFilter();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadCustomFiltersItem(t,l, oCustomFiltersItem);
                });
                aCustomFilters.push(oCustomFiltersItem);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadCustomFiltersItem = function(type, length, oCustomFiltersItem)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_CustomFilters.Operator == type )
                oCustomFiltersItem.Operator = this.stream.GetUChar();
            else if ( c_oSer_CustomFilters.Val == type )
                oCustomFiltersItem.Val = this.stream.GetString2LE(length);
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadDynamicFilter = function(type, length, oDynamicFilter)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_DynamicFilter.Type == type )
                oDynamicFilter.Type = this.stream.GetUChar();
            else if ( c_oSer_DynamicFilter.Val == type )
                oDynamicFilter.Val = this.stream.GetDoubleLE();
            else if ( c_oSer_DynamicFilter.MaxVal == type )
                oDynamicFilter.MaxVal = this.stream.GetDoubleLE();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadColorFilter = function(type, length, oColorFilter)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_ColorFilter.CellColor == type )
                oColorFilter.CellColor = this.stream.GetBool();
            else if ( c_oSer_ColorFilter.DxfId == type )
            {
                var DxfId = this.stream.GetULongLE();
                oColorFilter.dxf = this.initOpenManager.Dxfs[DxfId];
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadTop10 = function(type, length, oTop10)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_Top10.FilterVal == type )
                oTop10.FilterVal = this.stream.GetDoubleLE();
            else if ( c_oSer_Top10.Percent == type )
                oTop10.Percent = this.stream.GetBool();
            else if ( c_oSer_Top10.Top == type )
                oTop10.Top = this.stream.GetBool();
            else if ( c_oSer_Top10.Val == type )
                oTop10.Val = this.stream.GetDoubleLE();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadSortConditionContent = function(type, length, oSortCondition)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_SortState.ConditionRef == type )
                oSortCondition.Ref = AscCommonExcel.g_oRangeCache.getAscRange(this.stream.GetString2LE(length));
            else if ( c_oSer_SortState.ConditionSortBy == type )
                //TODO char? проверить
                oSortCondition.ConditionSortBy = this.stream.GetUChar();
            else if ( c_oSer_SortState.ConditionDescending == type )
                oSortCondition.ConditionDescending = this.stream.GetBool();
            else if ( c_oSer_SortState.ConditionDxfId == type )
            {
                var DxfId = this.stream.GetULongLE();
                oSortCondition.dxf = this.initOpenManager.Dxfs[DxfId];
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadSortCondition = function(type, length, aSortConditions)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_SortState.SortCondition == type )
            {
                var oSortCondition = new AscCommonExcel.SortCondition();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadSortConditionContent(t,l, oSortCondition);
                });
                aSortConditions.push(oSortCondition);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadSortConditionExternal = function()
		{
			var oThis = this;
			var oSortCondition = new AscCommonExcel.SortCondition();
			var length = this.stream.GetULongLE();
			var res = this.bcr.Read2Spreadsheet(length, function(t,l){
				return oThis.ReadSortConditionContent(t,l, oSortCondition);
			});
			return oSortCondition;
		};
        this.ReadSortState = function(type, length, oSortState)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_SortState.Ref == type )
                oSortState.Ref = AscCommonExcel.g_oRangeCache.getAscRange(this.stream.GetString2LE(length));
            else if ( c_oSer_SortState.CaseSensitive == type )
                oSortState.CaseSensitive = this.stream.GetBool();
            else if ( c_oSer_SortState.ColumnSort == type )
                oSortState.ColumnSort = this.stream.GetBool();
            else if ( c_oSer_SortState.SortMethod == type )
                oSortState.SortMethod = this.stream.GetUChar();
            else if ( c_oSer_SortState.SortConditions == type )
            {
                oSortState.SortConditions = [];
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadSortCondition(t,l, oSortState.SortConditions);
                });
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadTableColumn = function (type, length, oTableColumn) {
            var res = c_oSerConstants.ReadOk;
            const oThis = this;
            if (c_oSer_TableColumns.Name === type) {
                //replace only _x000a_ for fix bug(other spec. symbols didn't see in table columns)
                var columnName = this.stream.GetString2LE(length);
                oTableColumn.setTableColumnName(columnName.replaceAll("_x000a_", "\n"));
            } else if (c_oSer_TableColumns.TotalsRowLabel === type) {
                oTableColumn.TotalsRowLabel = this.stream.GetString2LE(length);
            } else if (c_oSer_TableColumns.TotalsRowFunction === type) {
                oTableColumn.TotalsRowFunction = this.stream.GetUChar();
            } else if (c_oSer_TableColumns.TotalsRowFormula === type) {
                var formula = this.stream.GetString2LE(length);
                this.initOpenManager.oReadResult.tableCustomFunc.push({formula: formula, column: oTableColumn, ws: this.ws});
            } else if (c_oSer_TableColumns.DataDxfId === type) {
                var DxfId = this.stream.GetULongLE();
                oTableColumn.dxf = this.initOpenManager.Dxfs[DxfId];
            }
            /*else if ( c_oSer_TableColumns.CalculatedColumnFormula == type )
			{
				oTableColumn.CalculatedColumnFormula = this.stream.GetString2LE(length);
			}*/ else if (c_oSer_TableColumns.QueryTableFieldId === type) {
                oTableColumn.queryTableFieldId = this.stream.GetULongLE();
            } else if (c_oSer_TableColumns.UniqueName === type) {
                oTableColumn.uniqueName = this.stream.GetString2LE(length);
            } else if (c_oSer_TableColumns.Id === type) {
                oTableColumn.id = this.stream.GetULongLE();
            } else if (c_oSer_TableColumns.XmlColumnPr === type) {
                if (!oTableColumn.xmlColumnPr) {
                    oTableColumn.xmlColumnPr = new AscCommonExcel.CXmlColumnPr();
                }
                this.bcr.Read2Spreadsheet(length, function(t, l) {
                    return oThis.ReadTableXmlColumnPr(t, l, oTableColumn.xmlColumnPr);
                }, this);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadTableXmlColumnPr = function(type, length, oXmlColumnPr) {
            var res = c_oSerConstants.ReadOk;

            if (c_oSer_TableColumns.MapId === type) {
                oXmlColumnPr.mapId = this.stream.GetLong();
            }
            else if (c_oSer_TableColumns.Xpath === type) {
                oXmlColumnPr.xpath = this.stream.GetString2LE(length);
            }
            else if (c_oSer_TableColumns.Denormalized === type) {
                oXmlColumnPr.denormalized = this.stream.GetBool();
            }
            else if (c_oSer_TableColumns.XmlDataType === type) {
                if (!oXmlColumnPr.xmlDataType) {
                    oXmlColumnPr.xmlDataType = {};
                }
                oXmlColumnPr.xmlDataType.val = this.stream.GetUChar();
            }
            else {
                res = c_oSerConstants.ReadUnknown;
            }

            return res;
        };
        this.ReadTableColumns = function(type, length, aTableColumns)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_TableColumns.TableColumn == type )
            {
                var oTableColumn = new AscCommonExcel.TableColumn();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadTableColumn(t,l, oTableColumn);
                });
                aTableColumns.push(oTableColumn);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadTableStyleInfo = function(type, length, oTableStyleInfo)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_TableStyleInfo.Name == type )
                oTableStyleInfo.Name = this.stream.GetString2LE(length);
            else if ( c_oSer_TableStyleInfo.ShowColumnStripes == type )
                oTableStyleInfo.ShowColumnStripes = this.stream.GetBool();
            else if ( c_oSer_TableStyleInfo.ShowRowStripes == type )
                oTableStyleInfo.ShowRowStripes = this.stream.GetBool();
            else if ( c_oSer_TableStyleInfo.ShowFirstColumn == type )
                oTableStyleInfo.ShowFirstColumn = this.stream.GetBool();
            else if ( c_oSer_TableStyleInfo.ShowLastColumn == type )
                oTableStyleInfo.ShowLastColumn = this.stream.GetBool();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadQueryTable = function(type, length, oQueryTable)
		{
			var oThis = this;
		    var res = c_oSerConstants.ReadOk;
			if(c_oSer_QueryTable.ConnectionId == type)
			{
				oQueryTable.connectionId = this.stream.GetULongLE();
			}
			else if(c_oSer_QueryTable.Name == type)
			{
				oQueryTable.name = this.stream.GetString2LE(length);
			}
			else if(c_oSer_QueryTable.AutoFormatId == type)
			{
				oQueryTable.autoFormatId = this.stream.GetULongLE();
			}
			else if(c_oSer_QueryTable.GrowShrinkType == type)
			{
				oQueryTable.growShrinkType = this.stream.GetString2LE(length);
			}
			else if(c_oSer_QueryTable.AdjustColumnWidth == type)
			{
				oQueryTable.adjustColumnWidth = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.ApplyAlignmentFormats == type)
			{
				oQueryTable.applyAlignmentFormats = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.ApplyBorderFormats == type)
			{
				oQueryTable.applyBorderFormats = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.ApplyFontFormats == type)
			{
				oQueryTable.applyFontFormats = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.ApplyNumberFormats == type)
			{
				oQueryTable.applyNumberFormats = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.ApplyPatternFormats == type)
			{
				oQueryTable.ApplyPatternFormats = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.ApplyWidthHeightFormats == type)
			{
				oQueryTable.applyWidthHeightFormats = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.BackgroundRefresh == type)
			{
				oQueryTable.backgroundRefresh = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.DisableEdit == type)
			{
				oQueryTable.disableEdit = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.DisableRefresh == type)
			{
				oQueryTable.disableRefresh = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.FillFormulas == type)
			{
				oQueryTable.fillFormulas = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.FirstBackgroundRefresh == type)
			{
				oQueryTable.firstBackgroundRefresh = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.Headers == type)
			{
				oQueryTable.headers = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.Intermediate == type)
			{
				oQueryTable.intermediate = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.PreserveFormatting == type)
			{
				oQueryTable.preserveFormatting = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.RefreshOnLoad == type)
			{
				oQueryTable.refreshOnLoad = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.RemoveDataOnSave == type)
			{
				oQueryTable.removeDataOnSave = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.RowNumbers == type)
			{
				oQueryTable.rowNumbers = this.stream.GetBool();
			}
			else if(c_oSer_QueryTable.QueryTableRefresh == type)
			{
				oQueryTable.queryTableRefresh = new AscCommonExcel.QueryTableRefresh();
				res = this.bcr.Read1(length, function(t,l){
					return oThis.ReadQueryTableRefresh(t,l, oQueryTable.queryTableRefresh);
				});
			}
			else
				res = c_oSerConstants.ReadUnknown;

			return res;
		};

		this.ReadQueryTableRefresh = function(type, length, queryTableRefresh)
		{
			var oThis = this;
			var res = c_oSerConstants.ReadOk;
			if(c_oSer_QueryTableRefresh.NextId == type)
			{
				queryTableRefresh.nextId = this.stream.GetULongLE();
			}
			else if(c_oSer_QueryTableRefresh.MinimumVersion == type)
			{
				queryTableRefresh.minimumVersion = this.stream.GetULongLE();
			}
			else if(c_oSer_QueryTableRefresh.UnboundColumnsLeft == type)
			{
				queryTableRefresh.unboundColumnsLeft = this.stream.GetULongLE();
			}
			else if(c_oSer_QueryTableRefresh.UnboundColumnsRight == type)
			{
				queryTableRefresh.unboundColumnsRight = this.stream.GetULongLE();
			}
			else if(c_oSer_QueryTableRefresh.FieldIdWrapped == type)
			{
				queryTableRefresh.fieldIdWrapped = this.stream.GetBool();
			}
			else if(c_oSer_QueryTableRefresh.HeadersInLastRefresh == type)
			{
				queryTableRefresh.headersInLastRefresh = this.stream.GetBool();
			}
			else if(c_oSer_QueryTableRefresh.PreserveSortFilterLayout == type)
			{
				queryTableRefresh.preserveSortFilterLayout = this.stream.GetBool();
			}
			else if(c_oSer_QueryTableRefresh.SortState == type)
			{
				queryTableRefresh.sortState = new AscCommonExcel.SortState();
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadSortState(t, l, queryTableRefresh.sortState);
				});
			}
			else if(c_oSer_QueryTableRefresh.QueryTableFields == type)
			{
				res = this.bcr.Read1(length, function (t, l) {
					return oThis.ReadQueryTableFields(t, l, queryTableRefresh);
				});
			}
			else if(c_oSer_QueryTableRefresh.QueryTableDeletedFields == type)
			{
				res = this.bcr.Read1(length, function (t, l) {
					return oThis.ReadQueryTableDeletedFields(t, l, queryTableRefresh);
				});
			}
			else
				res = c_oSerConstants.ReadUnknown;

			return res;
		};
		this.ReadQueryTableFields = function (type, length, queryTableRefresh) {
			var oThis = this;
			var res = c_oSerConstants.ReadOk;
			if (c_oSer_QueryTableField.QueryTableField === type) {
				var queryTableField = new AscCommonExcel.QueryTableField(true);
				res = this.bcr.Read1(length, function (t, l) {
					return oThis.ReadQueryTableField(t, l, queryTableField);
				});
				if (null == queryTableRefresh.queryTableFields) {
					queryTableRefresh.queryTableFields = [];
				}
				queryTableRefresh.queryTableFields.push(queryTableField);
			} else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
		this.ReadQueryTableField = function(type, length, queryTableField)
		{
			var res = c_oSerConstants.ReadOk;
			if(c_oSer_QueryTableField.Name == type)
			{
				queryTableField.name = this.stream.GetString2LE(length);
			}
			else if(c_oSer_QueryTableField.Id == type)
			{
				queryTableField.id = this.stream.GetULongLE(length);
			}
			else if(c_oSer_QueryTableField.TableColumnId == type)
			{
				queryTableField.tableColumnId = this.stream.GetULongLE(length);
			}
			else if(c_oSer_QueryTableField.RowNumbers == type)
			{
				queryTableField.rowNumbers = this.stream.GetBool();
			}
			else if(c_oSer_QueryTableField.FillFormulas == type)
			{
				queryTableField.fillFormulas = this.stream.GetBool();
			}
			else if(c_oSer_QueryTableField.DataBound == type)
			{
				queryTableField.dataBound = this.stream.GetBool();
			}
			else if(c_oSer_QueryTableField.Clipped == type)
			{
				queryTableField.clipped = this.stream.GetBool();
			}
			else
				res = c_oSerConstants.ReadUnknown;

			return res;
		};
		this.ReadQueryTableDeletedFields = function(type, length, queryTableRefresh)
		{
			var oThis = this;
			var res = c_oSerConstants.ReadOk;
			if(c_oSer_QueryTableDeletedField.QueryTableDeletedField == type)
			{
				var queryTableDeletedField = new AscCommonExcel.QueryTableDeletedField();
				res = this.bcr.Read1(length, function (t, l) {
					return oThis.ReadQueryTableDeletedField(t, l, queryTableDeletedField);
				});
				if (null == queryTableRefresh.queryTableDeletedFields) {
					queryTableRefresh.queryTableDeletedFields = [];
				}
				queryTableRefresh.queryTableDeletedFields.push(queryTableDeletedField);
			}
			else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
		this.ReadQueryTableDeletedField = function(type, length, pQueryTableDeletedField)
		{
			var res = c_oSerConstants.ReadOk;
			if(c_oSer_QueryTableDeletedField.Name == type)
			{
				pQueryTableDeletedField.name = this.stream.GetString2LE(length);
			}
			else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
    }
    /** @constructor */
    function Binary_SharedStringTableReader(stream, wb, aSharedStrings)
    {
        this.stream = stream;
        this.wb = wb;
        this.aSharedStrings = aSharedStrings;
        this.bcr = new Binary_CommonReader(this.stream);
        this.Read = function()
        {
            var oThis = this;
            var tempValue = {text: null, multiText: null};
            return this.bcr.ReadTable(function(t, l){
                return oThis.ReadSharedStringContent(t,l, tempValue);
            });
        };
        this.ReadSharedStringContent = function(type, length, tempValue)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerSharedStringTypes.Si === type )
            {
                var oThis = this;
                tempValue.text = null;
                tempValue.multiText = null;
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadSharedString(t,l, tempValue);
                });
                if(null != this.aSharedStrings) {
                    if (null != tempValue.multiText) {
                        let aMultiText = tempValue.multiText;
                        if (null != tempValue.text) {
                            let oElem = new AscCommonExcel.CMultiTextElem();
                            oElem.text = tempValue.text;
                            aMultiText.unshift(oElem);
                        }
                        this.aSharedStrings.push(aMultiText);
                    } else if (null != tempValue.text) {
                        this.aSharedStrings.push(tempValue.text);
                    } else {
                        this.aSharedStrings.push("");
                    }
                }
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadSharedString = function(type, length, tempValue)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerSharedStringTypes.Run == type )
            {
                var oThis = this;
                var oRun = new AscCommonExcel.CMultiTextElem();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadRun(t,l,oRun);
                });
                if(null == tempValue.multiText)
                    tempValue.multiText = [];
                tempValue.multiText.push(oRun);
            }
            else if ( c_oSerSharedStringTypes.Text == type )
            {
                if(null == tempValue.text)
                    tempValue.text = "";
                tempValue.text = checkMaxCellLength(this.stream.GetString2LE(length));
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadRun = function(type, length, oRun)
        {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerSharedStringTypes.RPr == type )
            {
                if(null == oRun.format)
                    oRun.format = new AscCommonExcel.Font();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadRPr(t,l, oRun.format);
                });
				oRun.format.checkSchemeFont(this.wb.theme);
            }
            else if ( c_oSerSharedStringTypes.Text == type )
            {
                if(null == oRun.text)
                    oRun.text = "";
                oRun.text = checkMaxCellLength(this.stream.GetString2LE(length));
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadRPr = function(type, length, rPr)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerFontTypes.Bold == type )
                rPr.b = this.stream.GetBool();
            else if ( c_oSerFontTypes.Color == type ){
				var color = ReadColorSpreadsheet2(this.bcr, length);
				if (color) {
					rPr.c = color;
				}
			} else if ( c_oSerFontTypes.Italic == type )
                rPr.i = this.stream.GetBool();
            else if ( c_oSerFontTypes.RFont == type )
                rPr.fn = this.stream.GetString2LE(length);
            else if ( c_oSerFontTypes.Strike == type )
                rPr.s = this.stream.GetBool();
            else if ( c_oSerFontTypes.Sz == type )
                rPr.fs = this.stream.GetDoubleLE();
            else if ( c_oSerFontTypes.Underline == type )
                rPr.u = this.stream.GetUChar();
            else if ( c_oSerFontTypes.VertAlign == type )
            {
                rPr.va = this.stream.GetUChar();
            }
            else if ( c_oSerFontTypes.Scheme == type )
                rPr.scheme = this.stream.GetUChar();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
    }
    /** @constructor */
    function Binary_StylesTableReader(stream, wb, useNumId/*, aCellXfs, isCopyPaste, useNumId*/)
    {
        this.stream = stream;
        this.wb = wb;
        //this.aCellXfs = aCellXfs;
        this.bcr = new Binary_CommonReader(this.stream);
        this.bssr = new Binary_SharedStringTableReader(this.stream, wb);
		//this.isCopyPaste = isCopyPaste;
		this.useNumId = useNumId;
        this.Read = function()
        {
            var oThis = this;
            var oStyleObject = {aBorders: [], aFills: [], aFonts: [], oNumFmts: {}, aCellStyleXfs: [],
                aCellXfs: [], aDxfs: [], aExtDxfs: [], aCellStyles: [], oCustomTableStyles: {}, oCustomSlicerStyles: null, oTimelineStyles: null};
            var res = this.bcr.ReadTable(function (t, l) {
                return oThis.ReadStylesContent(t, l, oStyleObject);
            });
            return oStyleObject;
        };
        this.ReadStylesContent = function (type, length, oStyleObject) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if (c_oSerStylesTypes.Borders === type) {
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadBorders(t, l, oStyleObject.aBorders);
                });
            } else if (c_oSerStylesTypes.Fills === type) {
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadFills(t, l, oStyleObject.aFills);
                });
            } else if (c_oSerStylesTypes.Fonts === type) {
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadFonts(t, l, oStyleObject.aFonts);
                });
            } else if (c_oSerStylesTypes.NumFmts === type) {
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadNumFmts(t, l, oStyleObject.oNumFmts);
                });
            } else if (c_oSerStylesTypes.CellStyleXfs === type) {
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadCellStyleXfs(t, l, oStyleObject.aCellStyleXfs);
                });
            } else if (c_oSerStylesTypes.CellXfs === type) {
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadCellXfs(t,l, oStyleObject.aCellXfs);
                });
            } else if (c_oSerStylesTypes.CellStyles === type) {
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadCellStyles(t, l, oStyleObject.aCellStyles);
                });
            } else if (c_oSerStylesTypes.Dxfs === type) {
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadDxfs(t, l, oStyleObject.aDxfs);
                });
            } else if (c_oSerStylesTypes.TableStyles === type && oThis.wb) {
                res = this.bcr.Read1(length, function (t, l){
                    return oThis.ReadTableStyles(t, l, oThis.wb.TableStyles, oStyleObject.oCustomTableStyles);
                });
            } else if (c_oSerStylesTypes.ExtDxfs === type) {
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadDxfs(t, l, oStyleObject.aExtDxfs);
                });
            } else if (c_oSerStylesTypes.SlicerStyles === type && typeof Asc.CT_slicerStyles != "undefined") {
                var fileStream = this.stream.ToFileStream();
                fileStream.GetUChar();
                oStyleObject.oCustomSlicerStyles = new Asc.CT_slicerStyles();
                oStyleObject.oCustomSlicerStyles.fromStream(fileStream);
                this.stream.FromFileStream(fileStream);
            } else if (c_oSerStylesTypes.TimelineStyles === type) {
                oStyleObject.oTimeLineStyles = new AscCommonExcel.CTimelineStyles();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadTimelineStyles(t, l, oStyleObject.oTimeLineStyles);
                });
                //while put in TimelineStyles
                this.wb.TimelineStyles = oStyleObject.oTimeLineStyles;
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadBorders = function(type, length, aBorders)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerStylesTypes.Border == type )
            {
                var oNewBorder = new AscCommonExcel.Border();
                //cell borders can not be null
                oNewBorder.initDefault();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadBorder(t,l,oNewBorder);
                });
                aBorders.push(oNewBorder);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadBorder = function(type, length, oNewBorder)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerBorderTypes.Bottom == type )
            {
                oNewBorder.b = new AscCommonExcel.BorderProp();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadBorderProp(t,l,oNewBorder.b);
                });
            }
            else if ( c_oSerBorderTypes.Diagonal == type )
            {
                oNewBorder.d = new AscCommonExcel.BorderProp();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadBorderProp(t,l,oNewBorder.d);
                });
            }
            else if ( c_oSerBorderTypes.End == type )
            {
                oNewBorder.r = new AscCommonExcel.BorderProp();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadBorderProp(t,l,oNewBorder.r);
                });
            }
            else if ( c_oSerBorderTypes.Horizontal == type )
            {
                oNewBorder.ih = new AscCommonExcel.BorderProp();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadBorderProp(t,l,oNewBorder.ih);
                });
            }
            else if ( c_oSerBorderTypes.Start == type )
            {
                oNewBorder.l = new AscCommonExcel.BorderProp();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadBorderProp(t,l,oNewBorder.l);
                });
            }
            else if ( c_oSerBorderTypes.Top == type )
            {
                oNewBorder.t = new AscCommonExcel.BorderProp();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadBorderProp(t,l,oNewBorder.t);
                });
            }
            else if ( c_oSerBorderTypes.Vertical == type )
            {
                oNewBorder.iv = new AscCommonExcel.BorderProp();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadBorderProp(t,l,oNewBorder.iv);
                });
            }
            else if ( c_oSerBorderTypes.DiagonalDown == type )
            {
                oNewBorder.dd = this.stream.GetBool();
            }
            else if ( c_oSerBorderTypes.DiagonalUp == type )
            {
                oNewBorder.du = this.stream.GetBool();
            }
            // else if ( c_oSerBorderTypes.Outline == type )
            // {
            // oNewBorder.outline = this.stream.GetBool();
            // }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadBorderProp = function(type, length, oBorderProp)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerBorderPropTypes.Style == type )
            {
                switch(this.stream.GetUChar())
                {
                    case EBorderStyle.borderstyleDashDot:			oBorderProp.setStyle(c_oAscBorderStyles.DashDot);break;
                    case EBorderStyle.borderstyleDashDotDot:		oBorderProp.setStyle(c_oAscBorderStyles.DashDotDot);break;
                    case EBorderStyle.borderstyleDashed:			oBorderProp.setStyle(c_oAscBorderStyles.Dashed);break;
                    case EBorderStyle.borderstyleDotted:			oBorderProp.setStyle(c_oAscBorderStyles.Dotted);break;
                    case EBorderStyle.borderstyleDouble:			oBorderProp.setStyle(c_oAscBorderStyles.Double);break;
                    case EBorderStyle.borderstyleHair:				oBorderProp.setStyle(c_oAscBorderStyles.Hair);break;
                    case EBorderStyle.borderstyleMedium:			oBorderProp.setStyle(c_oAscBorderStyles.Medium);break;
                    case EBorderStyle.borderstyleMediumDashDot:		oBorderProp.setStyle(c_oAscBorderStyles.MediumDashDot);break;
                    case EBorderStyle.borderstyleMediumDashDotDot:	oBorderProp.setStyle(c_oAscBorderStyles.MediumDashDotDot);break;
                    case EBorderStyle.borderstyleMediumDashed:		oBorderProp.setStyle(c_oAscBorderStyles.MediumDashed);break;
                    case EBorderStyle.borderstyleNone:				oBorderProp.setStyle(c_oAscBorderStyles.None);break;
                    case EBorderStyle.borderstyleSlantDashDot:		oBorderProp.setStyle(c_oAscBorderStyles.SlantDashDot);break;
                    case EBorderStyle.borderstyleThick:				oBorderProp.setStyle(c_oAscBorderStyles.Thick);break;
                    case EBorderStyle.borderstyleThin:				oBorderProp.setStyle(c_oAscBorderStyles.Thin);break;
                    default :										oBorderProp.setStyle(c_oAscBorderStyles.None);break;
                }
            }
            else if ( c_oSerBorderPropTypes.Color == type ) {
				oBorderProp.c = ReadColorSpreadsheet2(this.bcr, length);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadCellStyleXfs = function (type, length, aCellStyleXfs) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if (c_oSerStylesTypes.Xfs === type) {
				var oNewXfs = new OpenXf();
                res = this.bcr.Read2Spreadsheet(length, function (t, l) {
                    return oThis.ReadXfs(t, l, oNewXfs);
                });
                aCellStyleXfs.push(oNewXfs);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadCellXfs = function(type, length, aCellXfs)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerStylesTypes.Xfs == type )
            {
				var oNewXfs = new OpenXf();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadXfs(t,l,oNewXfs);
                });
                aCellXfs.push(oNewXfs);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadXfs = function(type, length, oXfs)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerXfsTypes.ApplyAlignment == type )
                oXfs.ApplyAlignment = this.stream.GetBool();
            else if ( c_oSerXfsTypes.ApplyBorder == type )
                oXfs.ApplyBorder = this.stream.GetBool();
            else if ( c_oSerXfsTypes.ApplyFill == type )
                oXfs.ApplyFill = this.stream.GetBool();
            else if ( c_oSerXfsTypes.ApplyFont == type )
                oXfs.ApplyFont = this.stream.GetBool();
            else if ( c_oSerXfsTypes.ApplyNumberFormat == type )
                oXfs.ApplyNumberFormat = this.stream.GetBool();
            else if ( c_oSerXfsTypes.BorderId == type )
                oXfs.borderid = this.stream.GetULongLE();
            else if ( c_oSerXfsTypes.FillId == type )
                oXfs.fillid = this.stream.GetULongLE();
            else if ( c_oSerXfsTypes.FontId == type )
                oXfs.fontid = this.stream.GetULongLE();
            else if ( c_oSerXfsTypes.NumFmtId == type )
                oXfs.numid = this.stream.GetULongLE();
            else if ( c_oSerXfsTypes.QuotePrefix == type )
                oXfs.QuotePrefix = this.stream.GetBool();
			else if ( c_oSerXfsTypes.PivotButton == type )
				oXfs.PivotButton = this.stream.GetBool();
            else if (c_oSerXfsTypes.XfId === type)
                oXfs.XfId = this.stream.GetULongLE();
            else if ( c_oSerXfsTypes.Aligment == type )
            {
                if(null == oXfs.align)
                    oXfs.align = new AscCommonExcel.Align();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadAligment(t,l,oXfs.align);
                });
            }
            else if (c_oSerXfsTypes.ApplyProtection == type) {
				oXfs.applyProtection = this.stream.GetBool();
            }
            else if ( c_oSerXfsTypes.Protection == type )
			{
				res = this.bcr.Read2Spreadsheet(length, function(t,l){
					return oThis.ReadProtection(t,l,oXfs);
				});
			}
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadProtection = function(type, length, oXfs)
		{
			var res = c_oSerConstants.ReadOk;
			if ( c_oSerProtectionTypes.Hidden == type )
				oXfs.hidden = this.stream.GetBool();
			else if ( c_oSerProtectionTypes.Locked == type )
				oXfs.locked = this.stream.GetBool();
			else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
        this.ReadAligment = function(type, length, oAligment)
        {
            var res = c_oSerConstants.ReadOk;
            if ( Asc.c_oSerAligmentTypes.Horizontal == type )
            {
                switch(this.stream.GetUChar())
                {
                    case 0 :
                    case 1 : oAligment.hor = AscCommon.align_Center;break;
                    case 2 :
                    case 3 :
                    case 5 : oAligment.hor = AscCommon.align_Justify;break;
                    case 4 : oAligment.hor = null;break;
                    case 6 : oAligment.hor = AscCommon.align_Left;break;
                    case 7 : oAligment.hor = AscCommon.align_Right;break;
                    case 8 : oAligment.hor = AscCommon.align_CenterContinuous;break;
                }
            }
            else if ( Asc.c_oSerAligmentTypes.Indent == type ) {
                oAligment.indent = this.stream.GetULongLE();
                if (oAligment.indent < 0) {
                    oAligment.indent = 0;
                }
            }
            else if ( Asc.c_oSerAligmentTypes.ReadingOrder == type )
                oAligment.ReadingOrder = this.stream.GetULongLE();
            else if ( Asc.c_oSerAligmentTypes.RelativeIndent == type )
                oAligment.RelativeIndent = this.stream.GetULongLE();
            else if ( Asc.c_oSerAligmentTypes.ShrinkToFit == type )
                oAligment.shrink = this.stream.GetBool();
            else if ( Asc.c_oSerAligmentTypes.TextRotation == type )
                oAligment.angle = this.stream.GetULongLE();
            else if ( Asc.c_oSerAligmentTypes.Vertical == type )
                oAligment.ver = this.stream.GetUChar();
            else if ( Asc.c_oSerAligmentTypes.WrapText == type )
                oAligment.wrap= this.stream.GetBool();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadFills = function(type, length, aFills)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerStylesTypes.Fill == type )
            {
                var oNewFill = new AscCommonExcel.Fill();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadFill(t,l,oNewFill);
                });
                aFills.push(oNewFill);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadFill = function(type, length, oFill)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerFillTypes.Pattern == type ) {
                var patternFill = new AscCommonExcel.PatternFill();
                res = this.bcr.Read1(length, function(t, l) {
                    return oThis.ReadPatternFill(t, l, patternFill);
                });
                oFill.patternFill = patternFill;
            } else if ( c_oSerFillTypes.Gradient == type ) {
                var gradientFill = new AscCommonExcel.GradientFill();
                res = this.bcr.Read1(length, function(t, l) {
                    return oThis.ReadGradientFill(t, l, gradientFill);
                });
                oFill.gradientFill = gradientFill;
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadPatternFill = function(type, length, patternFill)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerFillTypes.PatternBgColor_deprecated == type ) {
                patternFill.fromColor(ReadColorSpreadsheet2(this.bcr, length));
            } else if ( c_oSerFillTypes.PatternType == type ) {
                patternFill.patternType = this.stream.GetUChar();
            } else if ( c_oSerFillTypes.PatternFgColor == type ) {
                patternFill.fgColor = ReadColorSpreadsheet2(this.bcr, length);
            } else if ( c_oSerFillTypes.PatternBgColor == type ) {
                patternFill.bgColor = ReadColorSpreadsheet2(this.bcr, length);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadGradientFill = function(type, length, gradientFill)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerFillTypes.GradientType == type ) {
                gradientFill.type = this.stream.GetUChar();
            } else if ( c_oSerFillTypes.GradientLeft == type ) {
                gradientFill.left = this.stream.GetDoubleLE();
            } else if ( c_oSerFillTypes.GradientTop == type ) {
                gradientFill.top = this.stream.GetDoubleLE();
            } else if ( c_oSerFillTypes.GradientRight == type ) {
                gradientFill.right = this.stream.GetDoubleLE();
            } else if ( c_oSerFillTypes.GradientBottom == type ) {
                gradientFill.bottom = this.stream.GetDoubleLE();
            } else if ( c_oSerFillTypes.GradientDegree == type ) {
                gradientFill.degree = this.stream.GetDoubleLE();
            } else if ( c_oSerFillTypes.GradientStop == type ) {
                var gradientStop = new AscCommonExcel.GradientStop();
                res = this.bcr.Read1(length, function(t, l) {
                    return oThis.ReadGradientFillStop(t, l, gradientStop);
                });
                gradientFill.stop.push(gradientStop);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadGradientFillStop = function(type, length, gradientStop)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerFillTypes.GradientStopPosition == type ) {
                gradientStop.position = this.stream.GetDoubleLE();
            } else if ( c_oSerFillTypes.GradientStopColor == type ) {
                gradientStop.color = ReadColorSpreadsheet2(this.bcr, length);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadFonts = function(type, length, aFonts)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerStylesTypes.Font == type )
            {
                var oNewFont = new AscCommonExcel.Font();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.bssr.ReadRPr(t,l,oNewFont);
                });
                if (this.wb) {
                    oNewFont.checkSchemeFont(this.wb.theme);
                }
                aFonts.push(oNewFont);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadNumFmts = function(type, length, oNumFmts)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerStylesTypes.NumFmt == type )
            {
                var oNewNumFmt = {f: null, id: null};
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadNumFmt(t,l,oNewNumFmt);
                });
				if (null != oNewNumFmt.id) {
                    AscCommonExcel.InitOpenManager.prototype.ParseNum.call(this, oNewNumFmt, oNumFmts, this.useNumId);
				}
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadNumFmt = function(type, length, oNumFmt)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerNumFmtTypes.FormatCode == type )
            {
                oNumFmt.f = this.stream.GetString2LE(length);
            }
            else if ( c_oSerNumFmtTypes.NumFmtId == type )
            {
                oNumFmt.id = this.stream.GetULongLE();
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadCellStyles = function (type, length, aCellStyles) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            var oCellStyle = null;
            if (c_oSerStylesTypes.CellStyle === type) {
                oCellStyle = new AscCommonExcel.CCellStyle();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadCellStyle(t, l, oCellStyle);
                });
                aCellStyles.push(oCellStyle);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadCellStyle = function (type, length, oCellStyle) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_CellStyle.BuiltinId === type)
                oCellStyle.BuiltinId = this.stream.GetULongLE();
            else if (c_oSer_CellStyle.CustomBuiltin === type)
                oCellStyle.CustomBuiltin = this.stream.GetBool();
            else if (c_oSer_CellStyle.Hidden === type)
                oCellStyle.Hidden = this.stream.GetBool();
            else if (c_oSer_CellStyle.ILevel === type)
                oCellStyle.ILevel = this.stream.GetULongLE();
            else if (c_oSer_CellStyle.Name === type)
                oCellStyle.Name = this.stream.GetString2LE(length);
            else if (c_oSer_CellStyle.XfId === type)
                oCellStyle.XfId = this.stream.GetULongLE();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadDxfs = function(type, length, aDxfs)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerStylesTypes.Dxf == type )
            {
                var oDxf = new AscCommonExcel.CellXfs();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadDxf(t,l,oDxf);
                });
                aDxfs.push(oDxf);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadDxf = function(type, length, oDxf)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_Dxf.Alignment == type )
            {
                oDxf.align = new AscCommonExcel.Align();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadAligment(t,l,oDxf.align);
                });
            }
            else if ( c_oSer_Dxf.Border == type )
            {
                var oNewBorder = new AscCommonExcel.Border();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadBorder(t,l,oNewBorder);
                });
                oDxf.border = oNewBorder;
            }
            else if ( c_oSer_Dxf.Fill == type )
            {
                var oNewFill = new AscCommonExcel.Fill();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadFill(t,l,oNewFill);
                });
                oNewFill.fixForDxf();
                oDxf.fill = oNewFill;
            }
            else if ( c_oSer_Dxf.Font == type )
            {
                var oNewFont = new AscCommonExcel.Font();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.bssr.ReadRPr(t,l,oNewFont);
                });
                if (this.wb) {
                    oNewFont.checkSchemeFont(this.wb.theme);
                }
                oDxf.font = oNewFont;
            }
            else if ( c_oSer_Dxf.NumFmt == type )
            {
                var oNewNumFmt = {f: null, id: null};
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadNumFmt(t,l,oNewNumFmt);
                });
                if(null != oNewNumFmt.id)
                    oDxf.num = AscCommonExcel.InitOpenManager.prototype.ParseNum.call(this, oNewNumFmt, null, this.useNumId);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadDxfExternal = function () {
			var oThis = this;
			var dxf = new AscCommonExcel.CellXfs();
			var length = this.stream.GetULongLE();
			this.bcr.Read1(length, function (t, l) {
				return oThis.ReadDxf(t, l, dxf);
			});
			return dxf;
		};
        this.ReadTableStyles = function(type, length, oTableStyles, oCustomStyles)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_TableStyles.DefaultTableStyle == type )
                oTableStyles.DefaultTableStyle = this.stream.GetString2LE(length);
            else if ( c_oSer_TableStyles.DefaultPivotStyle == type )
                oTableStyles.DefaultPivotStyle = this.stream.GetString2LE(length);
            else if ( c_oSer_TableStyles.TableStyles == type )
            {
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadTableCustomStyles(t,l, oCustomStyles);
                });
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadTableCustomStyles = function(type, length, oCustomStyles)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if (c_oSer_TableStyles.TableStyle === type)
            {
                var oNewStyle = new CTableStyle();
                var aElements = [];
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadTableCustomStyle(t,l, oNewStyle, aElements);
                });
                if(null != oNewStyle.name) {
                    if (null === oNewStyle.displayName)
                        oNewStyle.displayName = oNewStyle.name;
                    oCustomStyles[oNewStyle.name] = {style : oNewStyle, elements: aElements};
                }
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadTableCustomStyle = function(type, length, oNewStyle, aElements)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if (c_oSer_TableStyle.Name === type)
                oNewStyle.name = this.stream.GetString2LE(length);
            else if (c_oSer_TableStyle.Pivot === type)
                oNewStyle.pivot = this.stream.GetBool();
            else if (c_oSer_TableStyle.Table === type)
                oNewStyle.table = this.stream.GetBool();
            else if (c_oSer_TableStyle.Elements === type) {
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadTableCustomStyleElements(t,l, aElements);
                });
            } else if (c_oSer_TableStyle.DisplayName === type)
                oNewStyle.displayName = this.stream.GetString2LE(length);
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadTableCustomStyleElements = function(type, length, aElements)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if (c_oSer_TableStyle.Element === type)
            {
                var oNewStyleElement = {Type: null, Size: null, DxfId: null};
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadTableCustomStyleElement(t,l, oNewStyleElement);
                });
                if(null != oNewStyleElement.Type && null != oNewStyleElement.DxfId)
                    aElements.push(oNewStyleElement);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadTableCustomStyleElement = function(type, length, oNewStyleElement)
        {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_TableStyleElement.Type === type)
                oNewStyleElement.Type = this.stream.GetUChar();
            else if (c_oSer_TableStyleElement.Size === type)
                oNewStyleElement.Size = this.stream.GetULongLE();
            else if (c_oSer_TableStyleElement.DxfId === type)
                oNewStyleElement.DxfId = this.stream.GetULongLE();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadTimelineStyles = function (type, length, oTimelineStyles) {
            var oThis = this;
            let res = c_oSerConstants.ReadOk;
            if (c_oSer_TimelineStyles.DefaultTimelineStyle === type) {
                oTimelineStyles.defaultTimelineStyle = this.stream.GetString2LE(length);
            } else if (c_oSer_TimelineStyles.TimelineStyle === type) {
                if (!oTimelineStyles.timelineStyles) {
                    oTimelineStyles.timelineStyles = [];
                }
                let newTimelineStyle = new AscCommonExcel.CTimelineStyle();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadTimelineStyle(t, l, newTimelineStyle);
                });
                oTimelineStyles.timelineStyles.push(newTimelineStyle);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadTimelineStyle = function (type, length, oTimelineStyle) {
            var oThis = this;
            let res = c_oSerConstants.ReadOk;
            if (c_oSer_TimelineStyles.TimelineStyleName === type) {
                oTimelineStyle.name = this.stream.GetString2LE(length);
            } else if (c_oSer_TimelineStyles.TimelineStyle === type) {
                if (!oTimelineStyle.timelineStyleElements) {
                    oTimelineStyle.timelineStyleElements = [];
                }
                let timelineStyleElement = new AscCommonExcel.CTimelineStyleElement();
                res = this.bcr.Read2Spreadsheet(length, function (t, l) {
                    return oThis.ReadTimelineStyleElement(t, l, timelineStyleElement);
                });
                oTimelineStyle.timelineStyleElements.push(timelineStyleElement);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadTimelineStyleElement = function (type, length, oTimelineStyleElement) {
            let res = c_oSerConstants.ReadOk;
            if (c_oSer_TimelineStyles.TimelineStyleElementType === type) {
                oTimelineStyleElement.type = this.stream.GetUChar();
            } else if (c_oSer_TimelineStyles.TimelineStyleElementDxfId === type) {
                oTimelineStyleElement.dxfId = this.stream.GetLong();
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };

    }
    /** @constructor */
    function Binary_WorkbookTableReader(stream, InitOpenManager, oWorkbook, bwtr)
    {
        this.stream = stream;
        this.InitOpenManager = InitOpenManager;
        this.oWorkbook = oWorkbook;
        this.bcr = new Binary_CommonReader(this.stream);
        this.bwtr = bwtr;
        this.Read = function()
        {
            var oThis = this;
            return this.bcr.ReadTable(function(t, l){
                return oThis.ReadWorkbookContent(t,l);
            });
        };
        this.ReadWorkbookContent = function(type, length)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerWorkbookTypes.WorkbookPr === type )
            {
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadWorkbookPr(t,l,oThis.oWorkbook.workbookPr);
                });
            }
            else if ( c_oSerWorkbookTypes.BookViews === type )
            {
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadBookViews(t,l);
                });
            }
            else if ( c_oSerWorkbookTypes.DefinedNames === type )
            {
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadDefinedNames(t,l);
                });
            }
			else if (c_oSerWorkbookTypes.CalcPr === type)
			{
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadCalcPr(t, l, oThis.oWorkbook.calcPr);
				});
			}
			else if ( c_oSerWorkbookTypes.ExternalReferences === type )
			{
                res = this.bcr.Read1(length, function(t,l){
					return oThis.ReadExternalReferences(t,l);
				});
			}
			else if ( c_oSerWorkbookTypes.OleSize === type )
			{
				var sRange = this.stream.GetString2LE(length);
				var parsedRange = AscCommonExcel.g_oRangeCache.getAscRange(sRange).clone();
				if (parsedRange) {
					this.oWorkbook.setOleSize(new AscCommonExcel.OleSizeSelectionRange(null, parsedRange));
				}
			}
            else if (c_oSerWorkbookTypes.VbaProject === type)
            {
                let _end_rec = this.stream.cur + length;
                while (this.stream.cur < _end_rec)
                {
                    var _at = this.stream.GetUChar();
                    switch (_at)
                    {
                        case 0:
                        {
                            var fileStream = this.stream.ToFileStream();
                            let vbaProject = new AscCommon.VbaProject();
                            vbaProject.fromStream(fileStream);
                            this.InitOpenManager.oReadResult.vbaProject = vbaProject;
                            this.stream.FromFileStream(fileStream);
                            break;
                        }
                        default:
                        {
                            this.stream.SkipRecord();
                            break;
                        }
                    }
                }
            }
			else if (c_oSerWorkbookTypes.JsaProject === type)
			{
				this.InitOpenManager.oReadResult.macros = AscCommon.GetStringUtf8(this.stream, length);
			}
            else if (c_oSerWorkbookTypes.Comments === type)
            {
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.bwtr.ReadCommentDatas(t,l, oThis.oWorkbook.aComments);
                });
            }
			else if (c_oSerWorkbookTypes.Connections === type)
			{
				this.oWorkbook.connections = this.stream.GetBuffer(length);
			}
			else if (c_oSerWorkbookTypes.PivotCaches === type && typeof Asc.CT_PivotCacheDefinition != "undefined")
			{
				res = this.bcr.Read1(length, function(t,l){
					return oThis.ReadPivotCaches(t,l);
				});
			}
            else if ((c_oSerWorkbookTypes.SlicerCaches === type || c_oSerWorkbookTypes.SlicerCachesExt === type) && typeof Asc.CT_slicerCacheDefinition != "undefined")
            {
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadSlicerCaches(t,l);
                });
            }
			else if (c_oSerWorkbookTypes.WorkbookProtection === type && typeof Asc.CWorkbookProtection != "undefined")
			{
                var workbookProtection = Asc.CWorkbookProtection ? new Asc.CWorkbookProtection(this.oWorkbook) : null;
                if (workbookProtection) {
                    this.oWorkbook.workbookProtection = workbookProtection;
                    res = this.bcr.Read2Spreadsheet(length, function (t, l) {
                        return oThis.ReadWorkbookProtection(t, l, oThis.oWorkbook.workbookProtection);
                    });
                } else {
                    res = c_oSerConstants.ReadUnknown;
                }
			}
			else if (c_oSerWorkbookTypes.FileSharing === type)
			{
				var fileSharing = Asc.CFileSharing ? new Asc.CFileSharing(this.oWorkbook) : null;
				if (fileSharing) {
					this.oWorkbook.fileSharing = fileSharing;
					res = this.bcr.Read2Spreadsheet(length, function (t, l) {
						return oThis.ReadFileSharing(t, l, oThis.oWorkbook.fileSharing);
					});
				} else {
					res = c_oSerConstants.ReadUnknown;
				}
			}
            else if (c_oSerWorkbookTypes.TimelineCaches === type)
            {
                this.oWorkbook.timelineCaches = [];
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadTimelineCaches(t, l, oThis.oWorkbook.timelineCaches);
                });
            }
            /*else if (c_oSerWorkbookTypes.Metadata === type)
            {
                this.oWorkbook.metadata = new AscCommonExcel.CMetadata();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMetadata(t, l, oThis.oWorkbook.metadata);
                });
            }*/
            else if (c_oSerWorkbookTypes.XmlMap === type) {
                //this.stream.Skip(1); //skip type





                /*LONG end = pReader->GetPos() + pReader->GetRecordSize() + 4;

                while (pReader->GetPos() < end)
                {
                    BYTE _rec = pReader->GetUChar();

                    switch (_rec)
                    {
                        case 0:
                        {
                            m_MapInfo.Init();
                            m_MapInfo->fromPPTY(pReader);
                        }break;
                        default:
                        {
                            pReader->SkipRecord();
                        }break;
                    }
                }
                pReader->Seek(end);*/

                this.stream.GetUChar()
                var _len = this.stream.GetULong();
                var _start_pos = this.stream.cur;
                var end = _len + _start_pos;

                let oXmlMap;
                while (this.stream.cur < end) {
                    let _rec = this.stream.GetUChar();

                    switch (_rec) {
                        case 0: {
                            oXmlMap = new AscCommonExcel.CMapInfo();
                            oXmlMap.fromPPTY(this.stream);
                            break;
                        }
                        default: {
                            this.stream.SkipRecord();
                            break;
                        }
                    }
                }

                this.stream.Seek(end);

                // Add to workbook
                if (this.oWorkbook.xmlMaps == null) {
                    this.oWorkbook.xmlMaps = [];
                }
                this.oWorkbook.xmlMaps.push(oXmlMap);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadSlicerCaches = function(type, length)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerWorkbookTypes.SlicerCache == type ) {
                var slicerCacheDefinition = new Asc.CT_slicerCacheDefinition();
                var fileStream = this.stream.ToFileStream();
                fileStream.GetUChar();
                slicerCacheDefinition.fromStream(fileStream, oThis.bwtr.InitOpenManager.copyPasteObj && oThis.bwtr.InitOpenManager.copyPasteObj.isCopyPaste);
                this.stream.FromFileStream(fileStream);
                this.InitOpenManager.oReadResult.slicerCaches[slicerCacheDefinition.name] = slicerCacheDefinition;
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadPivotCaches = function(type, length)
		{
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			if ( c_oSerWorkbookTypes.PivotCache == type ) {
				var pivotCache = new Asc.CT_PivotCacheDefinition();
				res = this.bcr.Read1(length, function(t,l){
					return oThis.ReadPivotCache(t,l, pivotCache);
				});
				this.InitOpenManager.oReadResult.pivotCacheDefinitions[pivotCache.id] = pivotCache;
			} else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
		this.ReadPivotCache = function(type, length, pivotCache)
		{
			var res = c_oSerConstants.ReadOk;
			if ( c_oSer_PivotTypes.id === type ) {
				pivotCache.id = this.stream.GetLong();
			} else if ( c_oSer_PivotTypes.cache === type ) {
				let idOld = pivotCache.id;
				new AscCommon.openXml.SaxParserBase().parse(AscCommon.GetStringUtf8(this.stream, length), pivotCache);
				pivotCache.id = idOld;
			} else if ( c_oSer_PivotTypes.record === type ) {
				var cacheRecords = new Asc.CT_PivotCacheRecords();
				new AscCommon.openXml.SaxParserBase().parse(AscCommon.GetStringUtf8(this.stream, length), cacheRecords);
				pivotCache.cacheRecords = cacheRecords;
			} else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};

        this.ReadTimelineCaches = function (type, length, aTimelineCaches) {
            let oThis = this;
            let res = c_oSerConstants.ReadOk;

            if (c_oSerWorkbookTypes.TimelineCache === type) {
                let oTimelineCache = new AscCommonExcel.CTimelineCacheDefinition();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadTimelineCache(t, l, oTimelineCache);
                });
                aTimelineCaches.push(oTimelineCache);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadTimelineCache = function (type, length, oTimelineCache) {
            let oThis = this;
            let res = c_oSerConstants.ReadOk;
            if (c_oSer_TimelineCache.Name === type) {
                oTimelineCache.name = this.stream.GetString2LE(length);
            } else if (c_oSer_TimelineCache.SourceName === type) {
                oTimelineCache.sourceName = this.stream.GetString2LE(length);
            } else if (c_oSer_TimelineCache.Uid === type) {
                oTimelineCache.uid = this.stream.GetString2LE(length);
            } else if (c_oSer_TimelineCache.PivotTables === type) {
                oTimelineCache.pivotTables = [];
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadTimelineCachePivotTables(t, l, oTimelineCache.pivotTables);
                });
            } else if (c_oSer_TimelineCache.State === type) {
                oTimelineCache.state = new AscCommonExcel.CTimelineState();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadTimelineState(t, l, oTimelineCache.state);
                });
            } else if (c_oSer_TimelineCache.PivotFilter === type) {
                oTimelineCache.pivotFilter = new AscCommonExcel.CTimelinePivotFilter();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadTimelinePivotFilter(t, l, oTimelineCache.pivotFilter);
                });
            } else {
                res = c_oSerConstants.ReadUnknown;
            }

            return res;
        }
        this.ReadTimelineCachePivotTables = function (type, length, aTimelineCachePivotTables) {
            let oThis = this;
            let res = c_oSerConstants.ReadOk;

            if (c_oSer_TimelineCache.PivotTable === type) {
                let oTimelineCachePivotTable = new AscCommonExcel.CTimelineCachePivotTable();
                res = this.bcr.Read2Spreadsheet(length, function (t, l) {
                    return oThis.ReadTimelineCachePivotTable(t, l, oTimelineCachePivotTable);
                });
                aTimelineCachePivotTables.push(oTimelineCachePivotTable);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };

        this.ReadTimelineCachePivotTable = function (type, length, oTimelineCachePivotTable) {
            let res = c_oSerConstants.ReadOk;

            if (c_oSer_TimelineCachePivotTable.Name === type) {
                oTimelineCachePivotTable.name = this.stream.GetString2LE(length);
            } else if (c_oSer_TimelineCachePivotTable.TabId === type) {
                oTimelineCachePivotTable.tabId = this.stream.GetLong(length);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        }

        this.ReadTimelineState = function (type, length, oTimelineState) {
            let oThis = this;
            let res = c_oSerConstants.ReadOk;
            if (c_oSer_TimelineState.Name === type) {
                oTimelineState.name = this.stream.GetString2LE(length);
            } else if (c_oSer_TimelineState.FilterState === type) {
                oTimelineState.singleRangeFilterState = this.stream.GetBool();
            } else if (c_oSer_TimelineState.PivotCacheId === type) {
                oTimelineState.pivotCacheId = this.stream.GetLong();
            } else if (c_oSer_TimelineState.MinimalRefreshVersion === type) {
                oTimelineState.minimalRefreshVersion = this.stream.GetLong();
            } else if (c_oSer_TimelineState.LastRefreshVersion === type) {
                oTimelineState.lastRefreshVersion = this.stream.GetLong();
            } else if (c_oSer_TimelineState.FilterType === type) {
                oTimelineState.filterType = this.stream.GetString2LE(length);
            } else if (c_oSer_TimelineState.Selection === type) {
                oTimelineState.selection = new AscCommonExcel.CTimelineRange();
                res = this.bcr.Read2Spreadsheet(length, function (t, l) {
                    return oThis.ReadTimelineRange(t, l, oTimelineState.selection);
                });
            } else if (c_oSer_TimelineState.Bounds === type) {
                oTimelineState.bounds = new AscCommonExcel.CTimelineRange();
                res = this.bcr.Read2Spreadsheet(length, function (t, l) {
                    return oThis.ReadTimelineRange(t, l, oTimelineState.bounds);
                });
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        }
        this.ReadTimelineRange = function (type, length, oTimelineRange) {
            let res = c_oSerConstants.ReadOk;

            if (c_oSer_TimelineRange.StartDate === type) {
                oTimelineRange.startDate = this.stream.GetString2LE(length);
            } else if (c_oSer_TimelineRange.EndDate === type) {
                oTimelineRange.endDate = this.stream.GetString2LE(length);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }

            return res;
        };
        this.ReadTimelinePivotFilter = function (type, length, oTimelinePivotFilter) {
            let res = c_oSerConstants.ReadOk;

            if (c_oSer_TimelinePivotFilter.Name === type) {
                oTimelinePivotFilter.name = this.stream.GetString2LE(length);
            } else if (c_oSer_TimelinePivotFilter.Description === type) {
                oTimelinePivotFilter.description = this.stream.GetString2LE(length);
            } else if (c_oSer_TimelinePivotFilter.UseWholeDay === type) {
                oTimelinePivotFilter.useWholeDay = this.stream.GetBool();
            } else if (c_oSer_TimelinePivotFilter.Id === type) {
                oTimelinePivotFilter.id = this.stream.Getlong();
            } else if (c_oSer_TimelinePivotFilter.Fld === type) {
                oTimelinePivotFilter.fld = this.stream.Getlong();
            } else if (c_oSer_TimelinePivotFilter.AutoFilter === type) {
                let oBinary_TableReader = new Binary_TableReader(this.stream, this.InitOpenManager, /*ws*/null);
                oTimelinePivotFilter.autoFilter = new AscCommonExcel.AutoFilter();
                res = this.bcr.Read1(length, function (t, l) {
                    return oBinary_TableReader.ReadAutoFilter(t, l, oTimelinePivotFilter.autoFilter);
                });
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };

        //****metadata****
        this.ReadMetadata = function (type, length, pMetadata) {
            var oThis = this;
            let res = c_oSerConstants.ReadOk;
            if (c_oSer_Metadata.MetadataTypes === type) {
                if (!pMetadata.metadataTypes) {
                    pMetadata.metadataTypes = [];
                }
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMetadataTypes(t, l, pMetadata.metadataTypes);
                });
            } else if (c_oSer_Metadata.MetadataStrings === type) {
                if (!pMetadata.metadataStrings) {
                    pMetadata.metadataStrings = [];
                }
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMetadataStrings(t, l, pMetadata.metadataStrings);
                });
            } else if (c_oSer_Metadata.MdxMetadata === type) {
                if (!pMetadata.mdxMetadata) {
                    pMetadata.mdxMetadata = [];
                }
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMdxMetadata(t, l, pMetadata.mdxMetadata);
                });
            } else if (c_oSer_Metadata.CellMetadata === type) {
                if (!pMetadata.cellMetadata) {
                    pMetadata.cellMetadata = [];
                }
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMetadataBlocks(t, l, pMetadata.cellMetadata);
                });
            } else if (c_oSer_Metadata.ValueMetadata === type) {
                if (!pMetadata.valueMetadata) {
                    pMetadata.valueMetadata = [];
                }
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMetadataBlocks(t, l, pMetadata.valueMetadata);
                });
            } else if (c_oSer_Metadata.FutureMetadata === type) {
                if (!pMetadata.aFutureMetadata) {
                    pMetadata.aFutureMetadata = [];
                }
                let pMetadataRecord = new AscCommonExcel.CFutureMetadata();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadFutureMetadata(t, l, pMetadataRecord);
                });
                pMetadata.aFutureMetadata.push(pMetadataRecord);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadMetadataTypes = function (type, length, aMetadataTypes) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_MetadataType.MetadataType === type) {
                var metadataType = new AscCommonExcel.CMetadataType();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMetadataType(t, l, metadataType);
                });
                aMetadataTypes.push(metadataType);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };

        this.ReadMetadataType = function (type, length, pMetadataType) {
            var res = c_oSerConstants.ReadOk;

            if (c_oSer_MetadataType.Name === type) {
                pMetadataType.name = this.stream.GetString2LE(length);
            } else if (c_oSer_MetadataType.MinSupportedVersion === type) {
                pMetadataType.minSupportedVersion = this.stream.GetULong();
            } else if (c_oSer_MetadataType.GhostRow === type) {
                pMetadataType.ghostRow = this.stream.GetBool();
            } else if (c_oSer_MetadataType.GhostCol === type) {
                pMetadataType.ghostCol = this.stream.GetBool();
            } else if (c_oSer_MetadataType.Edit === type) {
                pMetadataType.edit = this.stream.GetBool();
            } else if (c_oSer_MetadataType.Delete === type) {
                pMetadataType.delete = this.stream.GetBool();
            } else if (c_oSer_MetadataType.Copy === type) {
                pMetadataType.copy = this.stream.GetBool();
            } else if (c_oSer_MetadataType.PasteAll === type) {
                pMetadataType.pasteAll = this.stream.GetBool();
            } else if (c_oSer_MetadataType.PasteFormulas === type) {
                pMetadataType.pasteFormulas = this.stream.GetBool();
            } else if (c_oSer_MetadataType.PasteValues === type) {
                pMetadataType.pasteValues = this.stream.GetBool();
            } else if (c_oSer_MetadataType.PasteFormats === type) {
                pMetadataType.pasteFormats = this.stream.GetBool();
            } else if (c_oSer_MetadataType.PasteComments === type) {
                pMetadataType.pasteComments = this.stream.GetBool();
            } else if (c_oSer_MetadataType.PasteDataValidation === type) {
                pMetadataType.pasteDataValidation = this.stream.GetBool();
            } else if (c_oSer_MetadataType.PasteBorders === type) {
                pMetadataType.pasteBorders = this.stream.GetBool();
            } else if (c_oSer_MetadataType.PasteColWidths === type) {
                pMetadataType.pasteColWidths = this.stream.GetBool();
            } else if (c_oSer_MetadataType.PasteNumberFormats === type) {
                pMetadataType.pasteNumberFormats = this.stream.GetBool();
            } else if (c_oSer_MetadataType.Merge === type) {
                pMetadataType.merge = this.stream.GetBool();
            } else if (c_oSer_MetadataType.SplitFirst === type) {
                pMetadataType.splitFirst = this.stream.GetBool();
            } else if (c_oSer_MetadataType.SplitAll === type) {
                pMetadataType.splitAll = this.stream.GetBool();
            } else if (c_oSer_MetadataType.RowColShift === type) {
                pMetadataType.rowColShift = this.stream.GetBool();
            } else if (c_oSer_MetadataType.ClearAll === type) {
                pMetadataType.clearAll = this.stream.GetBool();
            } else if (c_oSer_MetadataType.ClearFormats === type) {
                pMetadataType.clearFormats = this.stream.GetBool();
            } else if (c_oSer_MetadataType.ClearContents === type) {
                pMetadataType.clearContents = this.stream.GetBool();
            } else if (c_oSer_MetadataType.ClearComments === type) {
                pMetadataType.clearComments = this.stream.GetBool();
            } else if (c_oSer_MetadataType.Assign === type) {
                pMetadataType.assign = this.stream.GetBool();
            } else if (c_oSer_MetadataType.Coerce === type) {
                pMetadataType.coerce = this.stream.GetBool();
            } else if (c_oSer_MetadataType.CellMeta === type) {
                pMetadataType.cellMeta = this.stream.GetBool();
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadMetadataStrings = function (type, length, aMetadataStrings) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_MetadataString.MetadataString === type) {
                let pMetadataString = new AscCommonExcel.CMetadataString();
                pMetadataString.v = this.stream.GetString2LE(length);
                aMetadataStrings.push(pMetadataString);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadMdxMetadata = function (type, length, aMdxMetadata) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_MdxMetadata.Mdx === type) {
                let pMdx = new AscCommonExcel.CMdx();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMdx(t, l, pMdx);
                });
                aMdxMetadata.push(pMdx);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadMdx = function (type, length, pMdx) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_MdxMetadata.NameIndex === type) {
                pMdx.n = this.stream.GetULong();
            } else if (c_oSer_MdxMetadata.FunctionTag === type) {
                //pMdx.F.SetValueFromByte(this.stream.GetUChar());
                pMdx.f = this.stream.GetUChar();
            } else if (c_oSer_MdxMetadata.MdxTuple === type) {
                //READ1_DEF(length, res, this.ReadMdxTuple, pMdx.MdxTuple.GetPovarer());
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMdx(t, l, pMdx.mdxTuple);
                });
            } else if (c_oSer_MdxMetadata.MdxSet === type) {
                //READ1_DEF(length, res, this.ReadMdxSet, pMdx.MdxSet.GetPovarer());
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMdx(t, l, pMdx.mdxSet);
                });
            } else if (c_oSer_MdxMetadata.MdxKPI === type) {
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMdx(t, l, pMdx.mdxKPI);
                });
            } else if (c_oSer_MdxMetadata.MdxMemeberProp === type) {
                //READ1_DEF(length, res, this.ReadMdxMemeberProp, pMdx.MdxMemeberProp.GetPovarer());
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMdx(t, l, pMdx.mdxMemeberProp);
                });
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        //TODO CMetadataBlock -> CMetadataRecord array leto array???
        this.ReadMetadataBlocks = function (type, length, aMetadataBlocks) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_MetadataBlock.MetadataBlock === type) {
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMetadataBlock(t, l, aMetadataBlocks);
                });
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadMetadataBlock = function (type, length, aMetadataBlocks) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_MetadataBlock.MetadataRecord === type) {
                let pMetadataRecord = new AscCommonExcel.CMetadataRecord();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMetadataRecord(t, l, pMetadataRecord);
                });
                aMetadataBlocks.push(pMetadataRecord);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadMetadataRecord = function (type, length, pMetadataRecord) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_MetadataBlock.MetadataRecordType === type) {
                pMetadataRecord.t = this.stream.GetULong();
            } else if (c_oSer_MetadataBlock.MetadataRecordValue === type) {
                pMetadataRecord.v = this.stream.GetULong();
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadDynamicArrayProperties = function (type, length, pDynamicArrayProperties) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_FutureMetadataBlock.DynamicArray === type) {
                pDynamicArrayProperties.fDynamic = this.stream.GetBool();
            } else if (c_oSer_FutureMetadataBlock.CollapsedArray === type) {
                pDynamicArrayProperties.fCollapsed = this.stream.GetBool();
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        }
        this.ReadMetadataStringIndex = function (type, length, pStringIndex) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_MetadataStringIndex.StringIsSet === type) {
                pStringIndex.s = this.stream.GetULong();
            } else if (c_oSer_MetadataStringIndex.IndexValue === type) {
                pStringIndex.x = this.stream.GetULong();
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        }
        this.ReadMdxMemeberProp = function (type, length, pMdxMemeberProp) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_MetadataMemberProperty.NameIndex === type) {
                pMdxMemeberProp.n = this.stream.GetULong();
            } else if (c_oSer_MetadataMemberProperty.Index === type) {
                pMdxMemeberProp.np = this.stream.GetULong();
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        }
        this.ReadMdxKPI = function (type, length, pMetadataRecord) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_MetadataMdxKPI.NameIndex === type) {
                pMetadataRecord.n = this.stream.GetULong();
            } else if (c_oSer_MetadataMdxKPI.Index === type) {
                pMetadataRecord.np = this.stream.GetULong();
            } else if (c_oSer_MetadataMdxKPI.Property === type) {
                //pMdxKPI.P.Init();
                //pMdxKPI.P.SetValueFromByte(this.stream.GetUChar());
                pMetadataRecord.op = this.stream.GetUChar();
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadMdxSet = function (type, length, pMdxSet)
        {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_MetadataMdxSet.Count === type)
            {
                pMdxSet.c = this.stream.GetULong();
            }
            else if (c_oSer_MetadataMdxSet.Index === type)
            {
                pMdxSet.ns = this.stream.GetULong();
            }
            else if (c_oSer_MetadataMdxSet.SortOrder === type)
            {
                //pMdxSet.O.Init();
                //pMdxSet.O.SetValueFromByte(this.stream.GetUChar());
                pMdxSet.o = this.stream.GetUChar();
            }
            else if (c_oSer_MetadataMdxSet.MetadataStringIndex === type)
            {
               let pMetadataStringIndex = new AscCommonExcel.CMetadataStringIndex();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMetadataStringIndex(t, l, pMetadataStringIndex);
                });
                pMdxSet.metadataStringIndexes.push(pMetadataStringIndex);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadMdxTuple = function (type, length, pMdxTuple) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_MetadataMdxTuple.IndexCount === type) {
                pMdxTuple.c = this.stream.GetULong();
            } else if (c_oSer_MetadataMdxTuple.StringIndex === type) {
                pMdxTuple.si = this.stream.GetULong();
            } else if (c_oSer_MetadataMdxTuple.CultureCurrency === type) {
                pMdxTuple.ct = this.stream.GetString2LE(length);
            } else if (c_oSer_MetadataMdxTuple.NumFmtIndex === type) {
                pMdxTuple.fi = this.stream.GetULong();
            } else if (c_oSer_MetadataMdxTuple.BackColor === type) {
                pMdxTuple.bc = this.stream.GetULong();
            } else if (c_oSer_MetadataMdxTuple.ForeColor === type) {
                pMdxTuple.fc = this.stream.GetULong();
            } else if (c_oSer_MetadataMdxTuple.Italic === type) {
                pMdxTuple.i = this.stream.GetBool();
            } else if (c_oSer_MetadataMdxTuple.Bold === type) {
                pMdxTuple.b = this.stream.GetBool();
            } else if (c_oSer_MetadataMdxTuple.Underline === type) {
                pMdxTuple.u = this.stream.GetBool();
            } else if (c_oSer_MetadataMdxTuple.Strike === type) {
                pMdxTuple.st = this.stream.GetBool();
            } else if (c_oSer_MetadataMdxTuple.MetadataStringIndex === type) {
                let pMetadataStringIndex = new AscCommonExcel.CMetadataStringIndex();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadMetadataStringIndex(t, l, pMetadataStringIndex);
                });
                pMdxTuple.metadataStringIndexes.push(pMetadataStringIndex);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadFutureMetadata = function (type, length, pCFutureMetadata)
        {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;

            if (c_oSer_FutureMetadataBlock.Name === type)
            {
                pCFutureMetadata.name = this.stream.GetString2LE(length);
            }
            else if (c_oSer_FutureMetadataBlock.FutureMetadataBlock === type)
            {
                if (!pCFutureMetadata.futureMetadataBlocks) {
                    pCFutureMetadata.futureMetadataBlocks = [];
                }
                let pFutureMetadataBlock = new AscCommonExcel.CFutureMetadataBlock();
				if (!pFutureMetadataBlock.extLst) {
					pFutureMetadataBlock.extLst = [];
				}
				let elemExtList = new AscCommonExcel.CMetadataBlockExt();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadFutureMetadataBlock(t, l, elemExtList);
                });
				pFutureMetadataBlock.extLst.push(elemExtList);
                pCFutureMetadata.futureMetadataBlocks.push(pFutureMetadataBlock);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadFutureMetadataBlock = function (type, length, pFutureMetadataBlock)
        {
        	var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_FutureMetadataBlock.RichValueBlock === type)
            {
                /*let pExt = new Asc.COfficeArtExtension();
                pExt.m_sUri = L"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}";
                pExt.RichValueBlock.Init();
                pExt.RichValueBlock.I = this.stream.GetULong();*/

                let richValueBlock = new AscCommonExcel.CRichValueBlock();
				richValueBlock.i = this.stream.GetULong();
				pFutureMetadataBlock.richValueBlock = richValueBlock;
            }
            else if (c_oSer_FutureMetadataBlock.DynamicArrayProperties === type)
            {

                /*OOX.Drawing.COfficeArtExtension* pExt = new OOX.Drawing.COfficeArtExtension();
                pExt.m_sUri = L"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}";
                pExt.DynamicArrayProperties.Init();

                READ1_DEF(length, res, this.ReadDynamicArrayProperties, pExt.DynamicArrayProperties.GetPovarer());
                pFutureMetadataBlock.ExtLst.m_arrExt.push_back(pExt);*/

                let pExt = new AscCommonExcel.CDynamicArrayProperties();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadDynamicArrayProperties(t, l, pExt);
                });
				pFutureMetadataBlock.dynamicArrayProperties = pExt;
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };

        this.ReadWorkbookPr = function(type, length, WorkbookPr)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerWorkbookPrTypes.Date1904 === type )
                WorkbookPr.setDate1904(this.stream.GetBool());
            else if ( c_oSerWorkbookPrTypes.DateCompatibility === type )
                WorkbookPr.setDateCompatibility(this.stream.GetBool());
			else if ( c_oSerWorkbookPrTypes.HidePivotFieldList === type ) {
				WorkbookPr.setHidePivotFieldList(this.stream.GetBool());
			} else if ( c_oSerWorkbookPrTypes.ShowPivotChartFilter === type ) {
				WorkbookPr.setShowPivotChartFilter(this.stream.GetBool());
			} else if ( c_oSerWorkbookPrTypes.UpdateLinks === type ) {
                WorkbookPr.setUpdateLinks(this.stream.GetUChar());
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadBookViews = function(type, length)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerWorkbookTypes.WorkbookView == type )
            {
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadWorkbookView(t,l);
                });
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadWorkbookView = function (type, length) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSerWorkbookViewTypes.ActiveTab === type) {
                this.oWorkbook.nActive = this.stream.GetULongLE();
            } else  if (c_oSerWorkbookViewTypes.ShowVerticalScroll === type) {
                this.oWorkbook.showVerticalScroll = this.stream.GetBool();
            } else  if (c_oSerWorkbookViewTypes.ShowHorizontalScroll === type) {
                this.oWorkbook.showHorizontalScroll = this.stream.GetBool();
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadDefinedNames = function(type, length)
        {
            var res = c_oSerConstants.ReadOk, LocalSheetId;
            var oThis = this;
            if ( c_oSerWorkbookTypes.DefinedName == type )
            {
                var oNewDefinedName = new Asc.asc_CDefName();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadDefinedName(t,l,oNewDefinedName);
                });
                this.InitOpenManager.oReadResult.defNames.push(oNewDefinedName);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadDefinedName = function(type, length, oDefinedName)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerDefinedNameTypes.Name == type )
                oDefinedName.Name = this.stream.GetString2LE(length);
            else if ( c_oSerDefinedNameTypes.Ref == type )
                oDefinedName.Ref = this.stream.GetString2LE(length);
            else if ( c_oSerDefinedNameTypes.LocalSheetId == type )
                oDefinedName.LocalSheetId = this.stream.GetULongLE();
            else if ( c_oSerDefinedNameTypes.Hidden == type )
                oDefinedName.Hidden = this.stream.GetBool();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadCalcPr = function(type, length, oCalcPr) {
			var res = c_oSerConstants.ReadOk;
			if (c_oSerCalcPrTypes.CalcId == type) {
				oCalcPr.calcId = this.stream.GetULongLE();
			} else if (c_oSerCalcPrTypes.CalcMode == type) {
				oCalcPr.calcMode = this.stream.GetUChar();
			} else if (c_oSerCalcPrTypes.FullCalcOnLoad == type) {
				oCalcPr.fullCalcOnLoad = this.stream.GetBool();
			} else if (c_oSerCalcPrTypes.RefMode == type) {
				oCalcPr.refMode = this.stream.GetUChar();
			} else if (c_oSerCalcPrTypes.Iterate == type) {
				oCalcPr.iterate = this.stream.GetBool();
			} else if (c_oSerCalcPrTypes.IterateCount == type) {
				oCalcPr.iterateCount = this.stream.GetULongLE();
			} else if (c_oSerCalcPrTypes.IterateDelta == type) {
				oCalcPr.iterateDelta = this.stream.GetDoubleLE();
			} else if (c_oSerCalcPrTypes.FullPrecision == type) {
				oCalcPr.fullPrecision = this.stream.GetBool();
			} else if (c_oSerCalcPrTypes.CalcCompleted == type) {
				oCalcPr.calcCompleted = this.stream.GetBool();
			} else if (c_oSerCalcPrTypes.CalcOnSave == type) {
				oCalcPr.calcOnSave = this.stream.GetBool();
			} else if (c_oSerCalcPrTypes.ConcurrentCalc == type) {
				oCalcPr.concurrentCalc = this.stream.GetBool();
			} else if (c_oSerCalcPrTypes.ConcurrentManualCount == type) {
				oCalcPr.concurrentManualCount = this.stream.GetULongLE();
			} else if (c_oSerCalcPrTypes.ForceFullCalc == type) {
				oCalcPr.forceFullCalc = this.stream.GetBool();
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
		this.ReadExternalReferences = function(type, length) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSerWorkbookTypes.ExternalReference === type) {
                var externalReferenceExt = {};
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadExternalReference(t, l, externalReferenceExt);
                });
                if (externalReferenceExt.externalReference && externalReferenceExt.externalFileId && externalReferenceExt.externalPortalName) {
                    var filId = externalReferenceExt.externalFileId;
                    if (filId) {
                        filId = decodeXmlPath(filId, true);
                    }

                    externalReferenceExt.externalReference.setReferenceData(filId, externalReferenceExt.externalPortalName);
                }
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
		};
        this.ReadExternalReference = function(type, length, externalReferenceExt) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if (c_oSerWorkbookTypes.ExternalBook === type) {
                var externalBook = new AscCommonExcel.ExternalReference();
                res = this.bcr.Read1(length, function(t, l) {
                    return oThis.ReadExternalBook(t, l, externalBook);
                });
                externalReferenceExt.externalReference = externalBook;
                this.oWorkbook.externalReferences.push(externalBook);
            } else if (c_oSerWorkbookTypes.OleLink === type) {
                this.oWorkbook.externalReferences.push({Type: 1, Buffer: this.stream.GetBuffer(length)});
            } else if (c_oSerWorkbookTypes.DdeLink === type) {
                this.oWorkbook.externalReferences.push({Type: 2, Buffer: this.stream.GetBuffer(length)});
            }  else if (c_oSerWorkbookTypes.ExternalFileId === type) {
                externalReferenceExt.externalFileId = this.stream.GetString2LE(length);
            } else if (c_oSerWorkbookTypes.ExternalPortalName === type) {
                externalReferenceExt.externalPortalName = this.stream.GetString2LE(length);
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
		this.ReadExternalBook = function(type, length, externalBook) {
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			if (c_oSer_ExternalLinkTypes.Id == type) {
			    var id = this.stream.GetString2LE(length);
			    if (id) {
                    id = decodeXmlPath(id);
                    /* TODO is it possible to transfer the id .replace when opening a file to another location?? */
                    id = completePathForLocalLinks(id);
                }
				externalBook.Id = id;
			} else if (c_oSer_ExternalLinkTypes.SheetNames == type) {
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadExternalSheetNames(t, l, externalBook.SheetNames);
				});
			} else if (c_oSer_ExternalLinkTypes.DefinedNames == type) {
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadExternalDefinedNames(t, l, externalBook.DefinedNames);
				});
			} else if (c_oSer_ExternalLinkTypes.SheetDataSet == type) {
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadExternalSheetDataSet(t, l, externalBook.SheetDataSet);
				});
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
		this.ReadExternalSheetNames = function(type, length, sheetNames) {
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			if (c_oSer_ExternalLinkTypes.SheetName == type) {
				sheetNames.push(this.stream.GetString2LE(length));
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
		this.ReadExternalDefinedNames = function(type, length, definedNames) {
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			if (c_oSer_ExternalLinkTypes.DefinedName === type) {
				var definedName = new AscCommonExcel.ExternalDefinedName();
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadExternalDefinedName(t, l, definedName);
				});
				definedNames.push(definedName);
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
		this.ReadExternalDefinedName = function(type, length, definedName) {
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			if (c_oSer_ExternalLinkTypes.DefinedNameName === type) {
				definedName.Name = this.stream.GetString2LE(length);
			} else if (c_oSer_ExternalLinkTypes.DefinedNameRefersTo === type) {
				definedName.RefersTo = this.stream.GetString2LE(length);
			} else if (c_oSer_ExternalLinkTypes.DefinedNameSheetId === type) {
				definedName.SheetId = this.stream.GetULongLE();
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
		this.ReadExternalSheetDataSet = function(type, length, sheetDataSet) {
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			if (c_oSer_ExternalLinkTypes.SheetData == type) {
				var sheetData = new AscCommonExcel.ExternalSheetDataSet();
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadExternalSheetData(t, l, sheetData);
				});
				sheetDataSet.push(sheetData);
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
		this.ReadExternalSheetData = function(type, length, sheetData) {
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			if (c_oSer_ExternalLinkTypes.SheetDataSheetId == type) {
				sheetData.SheetId = this.stream.GetULongLE();
			} else if (c_oSer_ExternalLinkTypes.SheetDataRefreshError == type) {
				sheetData.RefreshError = this.stream.GetBool();
			} else if (c_oSer_ExternalLinkTypes.SheetDataRow == type) {
				var row = new AscCommonExcel.ExternalRow();
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadExternalRow(t, l, row);
				});
				sheetData.Row.push(row);
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
		this.ReadExternalRow = function(type, length, row) {
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			if (c_oSer_ExternalLinkTypes.SheetDataRowR == type) {
				row.R = this.stream.GetULongLE();
			} else if (c_oSer_ExternalLinkTypes.SheetDataRowCell == type) {
				var cell = new AscCommonExcel.ExternalCell();
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadExternalCell(t, l, cell);
				});
				if (cell.CellType === Asc.ECellTypeType.celltypeError && cell.CellValue && AscCommon.rx_error && !cell.CellValue.match(AscCommon.rx_error)) {
					cell.CellValue = null;
				}
				row.Cell.push(cell);
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
		this.ReadExternalCell = function(type, length, cell) {
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			if (c_oSer_ExternalLinkTypes.SheetDataRowCellRef == type) {
				cell.Ref = this.stream.GetString2LE(length);
			} else if (c_oSer_ExternalLinkTypes.SheetDataRowCellType == type) {
				cell.CellType = this.stream.GetUChar();
			} else if (c_oSer_ExternalLinkTypes.SheetDataRowCellValue == type) {
				cell.CellValue = this.stream.GetString2LE(length);
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};

		this.ReadWorkbookProtection = function (type, length, workbookProtection) {
		    var res = c_oSerConstants.ReadOk;
			if (c_oSerWorkbookProtection.LockStructure == type) {
				workbookProtection.lockStructure = this.stream.GetBool();
			} else if (c_oSerWorkbookProtection.LockWindows == type) {
				workbookProtection.lockWindows = this.stream.GetBool();
			} else if (c_oSerWorkbookProtection.LockRevision == type) {
				workbookProtection.lockRevision = this.stream.GetBool();
			} else if (c_oSerWorkbookProtection.RevisionsAlgorithmName == type) {
				workbookProtection.revisionsAlgorithmName = this.stream.GetUChar();
			} else if (c_oSerWorkbookProtection.RevisionsSpinCount == type) {
				workbookProtection.revisionsSpinCount = this.stream.GetLong();
			} else if (c_oSerWorkbookProtection.RevisionsHashValue == type) {
				workbookProtection.revisionsHashValue = this.stream.GetString2LE(length);
			} else if (c_oSerWorkbookProtection.RevisionsSaltValue == type) {
				workbookProtection.revisionsSaltValue = this.stream.GetString2LE(length);
			} else if (c_oSerWorkbookProtection.WorkbookAlgorithmName == type) {
				workbookProtection.workbookAlgorithmName = this.stream.GetUChar();
			} else if (c_oSerWorkbookProtection.WorkbookSpinCount == type) {
				workbookProtection.workbookSpinCount = this.stream.GetLong();
			} else if (c_oSerWorkbookProtection.WorkbookHashValue == type) {
				workbookProtection.workbookHashValue = this.stream.GetString2LE(length);
			} else if (c_oSerWorkbookProtection.WorkbookSaltValue == type) {
				workbookProtection.workbookSaltValue = this.stream.GetString2LE(length);
			} else if (c_oSerWorkbookProtection.Password == type) {
				workbookProtection.workbookPassword = this.stream.GetString2LE(length);
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
		this.ReadFileSharing = function (type, length, fileSharing) {
			var res = c_oSerConstants.ReadOk;

			if (c_oSerFileSharing.AlgorithmName === type) {
				fileSharing.algorithmName = this.stream.GetUChar();
			} else if (c_oSerFileSharing.SpinCount === type) {
				fileSharing.spinCount = this.stream.GetLong();
			} else if (c_oSerFileSharing.HashValue === type) {
				fileSharing.hashValue = this.stream.GetString2LE(length);
			} else if (c_oSerFileSharing.SaltValue === type) {
				fileSharing.saltValue = this.stream.GetString2LE(length);
			} else if (c_oSerFileSharing.Password === type) {
				fileSharing.password = this.stream.GetString2LE(length);
			}  else if (c_oSerFileSharing.UserName === type) {
				fileSharing.userName = this.stream.GetString2LE(length);
			} else if (c_oSerFileSharing.ReadOnly === type) {
				fileSharing.readOnly = this.stream.GetBool();
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
    }
    /** @constructor */
    function Binary_WorksheetTableReader(stream, InitOpenManager, wb, aSharedStrings, aCellXfs, oMediaArray, personList)
    {
        this.stream = stream;
        this.wb = wb;
        this.aSharedStrings = aSharedStrings;
        this.oMediaArray = oMediaArray;
        this.aCellXfs = aCellXfs;
        this.bcr = new Binary_CommonReader(this.stream);
        this.aMerged = [];
        this.aHyperlinks = [];
        this.personList = personList;
        this.curWorksheet = null;
        this.InitOpenManager = InitOpenManager;
        this.Read = function()
        {
            var oThis = this;
            return this.bcr.ReadTable(function(t, l){
                return oThis.ReadWorksheetsContent(t,l);
            });
        };
		this.ReadSheetDataExternal = function(bNoBuildDep)
		{
			//console.profile('ReadSheetDataExternal');
			var oThis = this;
			var res = c_oSerConstants.ReadOk;
			var oldPos = this.stream.GetCurPos();
			for (var i = 0; i < this.InitOpenManager.oReadResult.sheetData.length; ++i) {
				var sheetDataElem = this.InitOpenManager.oReadResult.sheetData[i];
				var ws = sheetDataElem.ws;
				this.stream.Seek2(sheetDataElem.pos);

				var tmp = {
					pos: null, len: null, bNoBuildDep: bNoBuildDep, ws: ws, row: new AscCommonExcel.Row(ws),
					cell: new AscCommonExcel.Cell(ws), formula: new OpenFormula(), sharedFormulas: {},
					prevFormulas: {}, siFormulas: {}, prevRow: -1, prevCol: -1, formulaArray: []
				};


                res = this.bcr.Read1(sheetDataElem.len, function(t, l) {
                    return oThis.ReadSheetData(t, l, tmp);
                });

				if (!bNoBuildDep) {
					//TODO возможно стоит делать это в worksheet после полного чтения
					//***array-formula***
					//добавление ко всем ячейкам массива головной формулы
					for(var j = 0; j < tmp.formulaArray.length; j++) {
						var curFormula = tmp.formulaArray[j];
						var ref = curFormula.ref;
						if(ref) {
							var rangeFormulaArray = tmp.ws.getRange3(ref.r1, ref.c1, ref.r2, ref.c2);
							rangeFormulaArray._foreach(function(cell){
								cell.setFormulaInternal(curFormula);
								if (curFormula.ca || cell.isNullTextString()) {
									tmp.ws.workbook.dependencyFormulas.addToChangedCell(cell);
								}
							});
						}
					}
					for (var nCol in tmp.prevFormulas) {
						if (tmp.prevFormulas.hasOwnProperty(nCol)) {
							var prevFormula = tmp.prevFormulas[nCol];
							if (!tmp.siFormulas[prevFormula.parsed.getListenerId()]) {
								prevFormula.parsed.buildDependencies();
							}
						}
					}
					for (var listenerId in tmp.siFormulas) {
						if (tmp.siFormulas.hasOwnProperty(listenerId)) {
							tmp.siFormulas[listenerId].buildDependencies();
						}
					}
				}
				if(c_oSerConstants.ReadOk !== res)
					break;
			}
			this.stream.Seek2(oldPos);
			//console.profileEnd();
			return res;
		};
        this.ReadWorksheetsContent = function(type, length)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerWorksheetsTypes.Worksheet === type )
            {
                this.aMerged = [];
                this.aHyperlinks = [];
                var oNewWorksheet = new AscCommonExcel.Worksheet(this.wb, wb.aWorksheets.length);
                oNewWorksheet.aFormulaExt = [];
                var DrawingDocument = oNewWorksheet.getDrawingDocument();
				//TODO при copy/paste в word из excel необходимо подменить DrawingDocument из word - пересмотреть правку!
				if(typeof editor != "undefined" && editor && editor.WordControl && editor.WordControl.m_oLogicDocument && editor.WordControl.m_oLogicDocument.DrawingDocument) {
                   this.wb.DrawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;
                }


                this.curWorksheet = oNewWorksheet;
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadWorksheet(t,l, oNewWorksheet);
                });
                this.curWorksheet = null;

                //merged
                this.InitOpenManager.prepareAfterReadMergedCells(oNewWorksheet, this.aMerged);

                //hyperlinks
                this.InitOpenManager.prepareAfterReadHyperlinks(this.aHyperlinks);

                //put sheet
                this.InitOpenManager.putSheetAfterRead(this.wb, oNewWorksheet);

                this.wb.DrawingDocument = DrawingDocument;
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadWorksheet = function(type, length, oWorksheet)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            var oBinary_TableReader, oConditionalFormatting;
            if ( c_oSerWorksheetsTypes.WorksheetProp == type )
            {
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadWorksheetProp(t,l, oWorksheet);
                });
            }
            else if ( c_oSerWorksheetsTypes.Cols == type )
            {
                var aTempCols = [];
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadWorksheetCols(t,l, aTempCols, oWorksheet, oThis.aCellXfs);
                });

                this.InitOpenManager.prepareAfterReadCols(oWorksheet, aTempCols);
            }
            else if ( c_oSerWorksheetsTypes.SheetFormatPr == type )
            {
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadSheetFormatPr(t,l, oWorksheet);
                });
            }
            else if ( c_oSerWorksheetsTypes.PageMargins == type )
            {
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadPageMargins(t,l, oWorksheet.PagePrintOptions.pageMargins);
                });
            }
            else if ( c_oSerWorksheetsTypes.PageSetup == type )
            {
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadPageSetup(t,l, oWorksheet.PagePrintOptions.pageSetup);
                });
            }
            else if ( c_oSerWorksheetsTypes.PrintOptions == type )
            {
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadPrintOptions(t,l, oWorksheet.PagePrintOptions);
                });
            }
            else if ( c_oSerWorksheetsTypes.Hyperlinks == type )
            {
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadHyperlinks(t,l, oWorksheet);
                });
            }
            else if ( c_oSerWorksheetsTypes.MergeCells == type )
            {
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadMergeCells(t,l, oWorksheet);
                });
            }
            else if ( c_oSerWorksheetsTypes.SheetData == type )
            {
				this.InitOpenManager.oReadResult.sheetData.push({ws: oWorksheet, pos: this.stream.GetCurPos(), len: length});
				res = c_oSerConstants.ReadUnknown;
            }
            else if ( c_oSerWorksheetsTypes.Drawings == type )
            {
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadDrawings(t,l, oWorksheet.Drawings, oWorksheet);
                });
            }
            else if ( c_oSerWorksheetsTypes.Autofilter == type )
            {
                oBinary_TableReader = new Binary_TableReader(this.stream, this.InitOpenManager, oWorksheet);
                oWorksheet.AutoFilter = new AscCommonExcel.AutoFilter();
                res = this.bcr.Read1(length, function(t,l){
                    return oBinary_TableReader.ReadAutoFilter(t,l, oWorksheet.AutoFilter);
                });
            } else if (c_oSerWorksheetsTypes.SortState === type) {
                oBinary_TableReader = new Binary_TableReader(this.stream, this.InitOpenManager, oWorksheet);
                oWorksheet.sortState = new AscCommonExcel.SortState();
                res = this.bcr.Read1(length, function(t, l) {
                    return oBinary_TableReader.ReadSortState(t, l, oWorksheet.sortState);
                });
            } else if (c_oSerWorksheetsTypes.TableParts == type) {
                oBinary_TableReader = new Binary_TableReader(this.stream, this.InitOpenManager, oWorksheet);
                oBinary_TableReader.Read(length, oWorksheet.TableParts);
            } else if ( c_oSerWorksheetsTypes.Comments == type
                && !(typeof editor !== "undefined" && editor.WordControl && editor.WordControl.m_oLogicDocument && Array.isArray(editor.WordControl.m_oLogicDocument.Slides))) {
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadComments(t,l, oWorksheet);
                });
            } else if (c_oSerWorksheetsTypes.ConditionalFormatting === type && typeof AscCommonExcel.CConditionalFormatting != "undefined") {
                oConditionalFormatting = new AscCommonExcel.CConditionalFormatting();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadConditionalFormatting(t, l, oConditionalFormatting);
                });
                this.InitOpenManager.prepareConditionalFormatting(oWorksheet, oConditionalFormatting);
            } else if (c_oSerWorksheetsTypes.SheetViews === type) {
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadSheetViews(t, l, oWorksheet.sheetViews);
                });
            } else if (c_oSerWorksheetsTypes.SheetPr === type) {
                oWorksheet.sheetPr = new AscCommonExcel.asc_CSheetPr();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadSheetPr(t, l, oWorksheet.sheetPr);
                });
			} else if (c_oSerWorksheetsTypes.SparklineGroups === type) {
				res = this.bcr.Read1(length, function (t, l) {
					return oThis.ReadSparklineGroups(t, l, oWorksheet);
				});
			} else if (c_oSerWorksheetsTypes.HeaderFooter === type) {
                res = this.bcr.Read1(length, function(t, l) {
                    return oThis.ReadHeaderFooter(t, l, oWorksheet.headerFooter);
                });
            } else if (c_oSerWorksheetsTypes.RowBreaks === type) {
				oWorksheet.rowBreaks = new AscCommonExcel.CRowColBreaks();
				res = this.bcr.Read1(length, function (t, l) {
					return oThis.ReadRowColBreaks(t, l, oWorksheet.rowBreaks);
				});
			} else if (c_oSerWorksheetsTypes.ColBreaks === type) {
				oWorksheet.colBreaks = new AscCommonExcel.CRowColBreaks();
				res = this.bcr.Read1(length, function (t, l) {
					return oThis.ReadRowColBreaks(t, l, oWorksheet.colBreaks);
				});
            } else if (c_oSerWorksheetsTypes.LegacyDrawingHF === type) {
				oWorksheet.legacyDrawingHF = new AscCommonExcel.CLegacyDrawingHF();
				res = this.bcr.Read1(length, function (t, l) {
					return oThis.ReadLegacyDrawingHF(t, l, oWorksheet.legacyDrawingHF);
				});
			// } else if (c_oSerWorksheetsTypes.Picture === type) {
			//     oWorksheet.picture = this.stream.GetString2LE(length);
			} else if (c_oSerWorksheetsTypes.DataValidations === type && typeof AscCommonExcel.CDataValidations != "undefined") {
				oWorksheet.dataValidations = new AscCommonExcel.CDataValidations();
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadDataValidations(t, l, oWorksheet.dataValidations);
				});
			} else if (c_oSerWorksheetsTypes.PivotTable === type && typeof Asc.CT_pivotTableDefinition != "undefined") {
				var data = {table: null, cacheId: null};
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadPivotCopyPaste(t, l, data);
				});
				var cacheDefinition = this.InitOpenManager.oReadResult.pivotCacheDefinitions[data.cacheId];
				if (data.table && cacheDefinition) {
					//ignore duplicate pivot tables(from 8.3.2). dont compare by name, excel can change it
					const pivotIndex = oWorksheet.pivotTables.findIndex(function(elem){
						return elem.location && data.table.location && elem.location.isEqual(data.table.location);
					});
					if (pivotIndex === -1) {
						data.table.cacheDefinition = cacheDefinition;
						oWorksheet.insertPivotTable(data.table);
					}
				}
            } else if (c_oSerWorksheetsTypes.Slicers === type || c_oSerWorksheetsTypes.SlicersExt === type) {
                res = this.bcr.Read1(length, function(t, l) {
                    return oThis.ReadSlicers(t, l, oWorksheet);
                });
			} else if (c_oSerWorksheetsTypes.NamedSheetView === type) {
                var namedSheetViews = Asc.CT_NamedSheetViews ? new Asc.CT_NamedSheetViews() : null;
                if (namedSheetViews) {
                    var fileStream = this.stream.ToFileStream();
                    fileStream.GetUChar();
                    namedSheetViews.fromStream(fileStream, this.wb);
                    oWorksheet.aNamedSheetViews = namedSheetViews.namedSheetView;
                    this.stream.FromFileStream(fileStream);
                } else {
                    res = c_oSerConstants.ReadUnknown;
                }
			} else if (c_oSerWorksheetsTypes.ProtectionSheet === type && typeof Asc.CSheetProtection != "undefined") {
				var sheetProtection = Asc.CSheetProtection ? new Asc.CSheetProtection(oWorksheet) : null;
                if (sheetProtection) {
                    oWorksheet.sheetProtection = sheetProtection;
                    res = this.bcr.Read2Spreadsheet(length, function(t,l){
                        return oThis.ReadSheetProtection(t,l, oWorksheet.sheetProtection);
                    });
                } else {
                    res = c_oSerConstants.ReadUnknown;
                }
			} else if (c_oSerWorksheetsTypes.ProtectedRanges === type) {
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadProtectedRanges(t, l, oWorksheet.aProtectedRanges);
				});
			} else if (c_oSerWorksheetsTypes.CellWatches === type) {
                res = this.bcr.Read1(length, function(t, l) {
                    return oThis.ReadCellWatches(t, l, oWorksheet.aCellWatches);
                });
            } else if (c_oSerWorksheetsTypes.UserProtectedRanges === type) {
                res = this.bcr.Read1(length, function(t, l) {
                    return oThis.ReadUserProtectedRanges(t, l, oWorksheet.userProtectedRanges);
                });
            } else if (c_oSerWorksheetsTypes.TimelinesList === type) {
                res = this.bcr.Read1(length, function(t, l) {
                    return oThis.ReadTimelinesList(t, l, oWorksheet.timelines);
                });
            } else if (c_oSerWorksheetsTypes.TableSingleCells === type) {
                //this.stream.Skip(1); //skip type

                this.stream.GetUChar();

                var _len = this.stream.GetULong();
                var _start_pos = this.stream.cur;
                var end = _len + _start_pos;

                let oTableSingleCells;
                while (this.stream.cur < end) {
                    let _rec = this.stream.GetUChar();

                    switch (_rec) {
                        case 0: {
                            oTableSingleCells = new AscCommonExcel.CSingleXmlCells();
                            oTableSingleCells.fromPPTY(this.stream);
                            break;
                        }
                        default: {
                            this.stream.SkipRecord();
                            break;
                        }
                    }
                }

                this.stream.Seek(end);

                if (oTableSingleCells) {
                    if (oWorksheet.oTableSingleCells == null) {
                        oWorksheet.oTableSingleCells = [];
                    }
                    oWorksheet.oTableSingleCells.push(oTableSingleCells);
                }
            } else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
		this.ReadPivotCopyPaste = function(type, length, data)
		{
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			if (c_oSer_PivotTypes.cacheId == type) {
				data.cacheId = this.stream.GetLong();
			} else if (c_oSer_PivotTypes.table == type) {
				data.table = new Asc.CT_pivotTableDefinition(true);
				new AscCommon.openXml.SaxParserBase().parse(AscCommon.GetStringUtf8(this.stream, length), data.table);
			} else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
        this.ReadSlicers = function(type, length, oWorksheet)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if (c_oSerWorksheetsTypes.Slicer === type) {
                if (typeof Asc.CT_slicers == "undefined") {
                    if (this.InitOpenManager.copyPasteObj.isCopyPaste) {
                        oWorksheet.aSlicers.push(null);
                    }
                    return c_oSerConstants.ReadUnknown;
                }

                var slicers = new Asc.CT_slicers();
                slicers.slicer = oWorksheet.aSlicers;
                var fileStream = this.stream.ToFileStream();
                fileStream.GetUChar();
                slicers.fromStream(fileStream, oWorksheet, oThis.InitOpenManager.oReadResult.slicerCaches);
                this.stream.FromFileStream(fileStream);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadDataValidations = function(type, length, dataValidations)
		{
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			if (c_oSer_DataValidation.DataValidations == type) {
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadDataValidationsContent(t, l, dataValidations);
				});
			} else if (c_oSer_DataValidation.DisablePrompts == type) {
				dataValidations.disablePrompts = this.stream.GetBool();
			} else if (c_oSer_DataValidation.XWindow == type) {
				dataValidations.xWindow = this.stream.GetLong();
			} else if (c_oSer_DataValidation.YWindow == type) {
				dataValidations.yWindow = this.stream.GetLong();
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadDataValidationsContent = function(type, length, dataValidations)
		{
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			if (c_oSer_DataValidation.DataValidation == type) {
				var dataValidation = new AscCommonExcel.CDataValidation();
				res = this.bcr.Read2(length, function(t, l) {
					return oThis.ReadDataValidation(t, l, dataValidation);
				});
				if (dataValidation && dataValidation.ranges) {
					dataValidations.elems.push(dataValidation);
				}
			} else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
		this.ReadDataValidation = function(type, length, dataValidation)
		{
			var res = c_oSerConstants.ReadOk;
			if (c_oSer_DataValidation.AllowBlank == type) {
				dataValidation.allowBlank = this.stream.GetBool();
			} else if (c_oSer_DataValidation.Type == type) {
				dataValidation.type = this.stream.GetUChar();
			} else if (c_oSer_DataValidation.Error == type) {
				dataValidation.error = this.stream.GetString2LE(length);
			} else if (c_oSer_DataValidation.ErrorTitle == type) {
				dataValidation.errorTitle = this.stream.GetString2LE(length);
			} else if (c_oSer_DataValidation.ErrorStyle == type) {
				dataValidation.errorStyle = this.stream.GetUChar();
			} else if (c_oSer_DataValidation.ImeMode == type) {
				dataValidation.imeMode = this.stream.GetUChar();
			} else if (c_oSer_DataValidation.Operator == type) {
				dataValidation.operator = this.stream.GetUChar();
			} else if (c_oSer_DataValidation.Prompt == type) {
				dataValidation.prompt = this.stream.GetString2LE(length);
			} else if (c_oSer_DataValidation.PromptTitle == type) {
				dataValidation.promptTitle = this.stream.GetString2LE(length);
			} else if (c_oSer_DataValidation.ShowDropDown == type) {
				dataValidation.showDropDown = this.stream.GetBool();
			} else if (c_oSer_DataValidation.ShowErrorMessage == type) {
				dataValidation.showErrorMessage = this.stream.GetBool();
			} else if (c_oSer_DataValidation.ShowInputMessage == type) {
				dataValidation.showInputMessage = this.stream.GetBool();
			} else if (c_oSer_DataValidation.SqRef == type) {
				dataValidation.setSqRef(this.stream.GetString2LE(length));
			} else if (c_oSer_DataValidation.Formula1 == type) {
				dataValidation.formula1 = new Asc.CDataFormula(this.stream.GetString2LE(length));
			} else if (c_oSer_DataValidation.Formula2 == type) {
				dataValidation.formula2 = new Asc.CDataFormula(this.stream.GetString2LE(length));
			} else if (c_oSer_DataValidation.List == type) {
				dataValidation.list = new Asc.CDataFormula(this.stream.GetString2LE(length));
			} else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
		this.ReadSheetProtection = function (type, length, sheetProtection) {
			var res = c_oSerConstants.ReadOk;
			if (c_oSerWorksheetProtection.AlgorithmName == type) {
				sheetProtection.algorithmName = this.stream.GetUChar();
			} else if (c_oSerWorksheetProtection.SpinCount == type) {
				sheetProtection.spinCount = this.stream.GetLong();
			} else if (c_oSerWorksheetProtection.HashValue == type) {
				sheetProtection.hashValue = this.stream.GetString2LE(length);
			} else if (c_oSerWorksheetProtection.SaltValue == type) {
				sheetProtection.saltValue = this.stream.GetString2LE(length);
			} else if (c_oSerWorksheetProtection.Password == type) {
				sheetProtection.password = this.stream.GetString2LE(length);
			} else if (c_oSerWorksheetProtection.AutoFilter == type) {
				sheetProtection.autoFilter = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.Content == type) {
				sheetProtection.content = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.DeleteColumns == type) {
				sheetProtection.deleteColumns = this.stream.GetBool(length);
			} else if (c_oSerWorksheetProtection.DeleteRows == type) {
				sheetProtection.deleteRows = this.stream.GetBool(length);
			} else if (c_oSerWorksheetProtection.FormatCells == type) {
				sheetProtection.formatCells = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.FormatColumns == type) {
				sheetProtection.formatColumns = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.FormatRows == type) {
				sheetProtection.formatRows = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.InsertColumns == type) {
				sheetProtection.insertColumns = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.InsertHyperlinks == type) {
				sheetProtection.insertHyperlinks = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.InsertRows == type) {
				sheetProtection.insertRows = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.Objects == type) {
				sheetProtection.objects = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.PivotTables == type) {
				sheetProtection.pivotTables = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.Scenarios == type) {
				sheetProtection.scenarios = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.SelectLockedCells == type) {
				sheetProtection.selectLockedCells = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.SelectUnlockedCells == type) {
				sheetProtection.selectUnlockedCells = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.Sheet == type) {
				sheetProtection.sheet = this.stream.GetBool();
			} else if (c_oSerWorksheetProtection.Sort == type) {
				sheetProtection.sort = this.stream.GetBool();
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
		this.ReadProtectedRanges = function (type, length, aProtectedRanges) {
			var res = c_oSerConstants.ReadOk;
			var oThis = this;
			var oProtectedRange = null;

			if (c_oSerWorksheetsTypes.ProtectedRange === type) {
				oProtectedRange = Asc.CProtectedRange ? new Asc.CProtectedRange() : null;
                if (oProtectedRange) {
                    res = this.bcr.Read2(length, function (t, l) {
                        return oThis.ReadProtectedRange(t, l, oProtectedRange);
                    });
                    aProtectedRanges.push(oProtectedRange);
                } else {
                    res = c_oSerConstants.ReadUnknown;
                }
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
		this.ReadProtectedRange = function (type, length, oProtectedRange) {
			var res = c_oSerConstants.ReadOk;
			if (c_oSerProtectedRangeTypes.AlgorithmName === type) {
				oProtectedRange.algorithmName = this.stream.GetUChar();
			} else if (c_oSerProtectedRangeTypes.SpinCount === type) {
				oProtectedRange.spinCount = this.stream.GetLong();
			} else if (c_oSerProtectedRangeTypes.HashValue === type) {
				oProtectedRange.hashValue = this.stream.GetString2LE(length);
			} else if (c_oSerProtectedRangeTypes.SaltValue === type) {
				oProtectedRange.saltValue = this.stream.GetString2LE(length);
			} else if (c_oSerProtectedRangeTypes.Name === type) {
				oProtectedRange.name = this.stream.GetString2LE(length);
			} else if (c_oSerProtectedRangeTypes.SqRef === type) {
				var sqRef = this.stream.GetString2LE(length);
				var newSqRef = AscCommonExcel.g_oRangeCache.getRangesFromSqRef(sqRef);
				if (newSqRef.length > 0) {
					oProtectedRange.sqref = newSqRef;
				}
			} else if (c_oSerProtectedRangeTypes.SecurityDescriptor === type) {
				oProtectedRange.securityDescriptor = this.stream.GetString2LE(length);
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
        this.ReadCellWatches = function (type, length, aCellWatches) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            var oCellWatch = null;

            if (c_oSerWorksheetsTypes.CellWatch === type) {
                oCellWatch = AscCommonExcel.CCellWatch ? new AscCommonExcel.CCellWatch() : null;
                if (oCellWatch) {
                    res = this.bcr.Read2(length, function (t, l) {
                        return oThis.ReadCellWatch(t, l, oCellWatch);
                    });
                    aCellWatches.push(oCellWatch);
                } else {
                    res = c_oSerConstants.ReadUnknown;
                }
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadCellWatch = function (type, length, oCellWatch) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSerWorksheetsTypes.CellWatchR === type) {
                var range = AscCommonExcel.g_oRangeCache.getAscRange(this.stream.GetString2LE(length));
                if (range) {
                    oCellWatch.r = new Asc.Range(range.c1, range.r1, range.c1, range.r1);
                }
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadUserProtectedRanges = function (type, length, aUserProtectedRanges) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            var oUserProtectedRange = null;

            if (c_oSerUserProtectedRange.UserProtectedRange === type) {
                oUserProtectedRange = Asc.CUserProtectedRange ? new Asc.CUserProtectedRange() : null;
                if (oUserProtectedRange) {
                    res = this.bcr.Read1(length, function (t, l) {
                        return oThis.ReadUserProtectedRange(t, l, oUserProtectedRange);
                    });
                    aUserProtectedRanges.push(oUserProtectedRange);
                } else {
                    res = c_oSerConstants.ReadUnknown;
                }
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
		this.ReadUserProtectedRangeDesc = function (type, length, oUser) {
			var res = c_oSerConstants.ReadOk;

			if (c_oSerUserProtectedRangeDesc.Name === type) {
				oUser.name = this.stream.GetString2LE(length);
			} else if (c_oSerUserProtectedRangeDesc.Id === type) {
				oUser.id = this.stream.GetString2LE(length);
			} else if (c_oSerUserProtectedRangeDesc.Type === type) {
				oUser.type = this.stream.GetByte();
			} else {
				res = c_oSerConstants.ReadUnknown;
			}

			return res;
		};
        this.ReadUserProtectedRange = function (type, length, oUserProtectedRange) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if (c_oSerUserProtectedRange.Name === type) {
                oUserProtectedRange.name = this.stream.GetString2LE(length);
            } else if (c_oSerUserProtectedRange.Sqref === type) {
                var range = AscCommonExcel.g_oRangeCache.getAscRange(this.stream.GetString2LE(length));
                if (range) {
                    oUserProtectedRange.ref = range.clone();
                }
            } else if (c_oSerUserProtectedRange.Text === type) {
                oUserProtectedRange.warningText = this.stream.GetString2LE(length);
            } else if (c_oSerUserProtectedRange.Type === type) {
                oUserProtectedRange.type = this.stream.GetByte(length);
            } else if (c_oSerUserProtectedRange.User === type)
			{

				let oUser = Asc.CUserProtectedRangeUserInfo ? new Asc.CUserProtectedRangeUserInfo() : null;
				if (oUser) {
					res = this.bcr.Read2(length, function (t, l) {
						return oThis.ReadUserProtectedRangeDesc(t, l, oUser);
					});
					if (!oUserProtectedRange.users) {
						oUserProtectedRange.users = [];
					}
					oUserProtectedRange.users.push(oUser);
				} else {
					res = c_oSerConstants.ReadUnknown;
				}
			}
			else if (c_oSerUserProtectedRange.UsersGroup === type)
			{
				let oUser = Asc.CUserProtectedRangeUserInfo ? new Asc.CUserProtectedRangeUserInfo() : null;
				if (oUser) {
					res = this.bcr.Read2(length, function (t, l) {
						return oThis.ReadUserProtectedRangeDesc(t, l, oUser);
					});
					if (!oUserProtectedRange.usersGroups) {
						oUserProtectedRange.usersGroups = [];
					}
					oUserProtectedRange.usersGroups.push(oUser);
				} else {
					res = c_oSerConstants.ReadUnknown;
				}
			} else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadTimelinesList = function (type, length, aTimelines) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            var oTimeline = null;

            if (c_oSerWorksheetsTypes.Timelines === type) {
                oTimeline = AscCommonExcel.CTimeline ? new AscCommonExcel.CTimeline() : null;
                if (oTimeline) {
                    res = this.bcr.Read1(length, function (t, l) {
                        return oThis.ReadTimelines(t, l, oTimeline);
                    });
                    aTimelines.push(oTimeline);
                } else {
                    res = c_oSerConstants.ReadUnknown;
                }
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadTimelines = function (type, length, oTimeline) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;

            if (c_oSerWorksheetsTypes.Timeline === type) {
                res = this.bcr.Read2(length, function (t, l) {
                    return oThis.ReadTimeline(t, l, oTimeline);
                });
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadTimeline = function (type, length, oTimeline) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;

            if (c_oSer_Timeline.Name === type) {
                oTimeline.name = this.stream.GetString2LE(length);
            } else if (c_oSer_Timeline.Cache === type) {
                oTimeline.cache = this.stream.GetString2LE(length);
            } else if (c_oSer_Timeline.Caption === type) {
                oTimeline.caption = this.stream.GetString2LE(length);
            } else if (c_oSer_Timeline.ScrollPosition === type) {
                oTimeline.scrollPosition = this.stream.GetString2LE(length);
            } else if (c_oSer_Timeline.Uid === type) {
                oTimeline.uid = this.stream.GetString2LE(length);
            } else if (c_oSer_Timeline.Level === type) {
                oTimeline.level = this.stream.GetULong();
            } else if (c_oSer_Timeline.SelectionLevel === type) {
                oTimeline.selectionLevel = this.stream.GetULong();
            } else if (c_oSer_Timeline.ShowHeader === type) {
                oTimeline.showHeader = this.stream.GetBool();
            } else if (c_oSer_Timeline.ShowHorizontalScrollbar === type) {
                oTimeline.showHorizontalScrollbar = this.stream.GetBool();
            } else if (c_oSer_Timeline.ShowSelectionLabel === type) {
                oTimeline.showSelectionLabel = this.stream.GetBool();
            } else if (c_oSer_Timeline.ShowTimeLevel === type) {
                oTimeline.showTimeLevel = this.stream.GetBool();
            } else if (c_oSer_Timeline.Style === type) {
                //oTimeline.Style.Init();
                //oTimeline.Style->SetValueFromByte(this.stream.GetUChar());
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadWorksheetProp = function(type, length, oWorksheet)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerWorksheetPropTypes.Name == type )
            {
                oWorksheet.sName = this.stream.GetString2LE(length);
				AscFonts.FontPickerByCharacter.getFontsByString(oWorksheet.sName);
            }
            else if ( c_oSerWorksheetPropTypes.SheetId == type ) {
                this.InitOpenManager.oReadResult.sheetIds[this.stream.GetULongLE()] = oWorksheet;
            } else if ( c_oSerWorksheetPropTypes.State == type )
            {
                switch(this.stream.GetUChar())
                {
                    case EVisibleType.visibleHidden: oWorksheet.bHidden = true;break;
                    case EVisibleType.visibleVeryHidden: oWorksheet.bHidden = true;break;
                    case EVisibleType.visibleVisible: oWorksheet.bHidden = false;break;
                }
            }
            else if(this.InitOpenManager.copyPasteObj.isCopyPaste && c_oSerWorksheetPropTypes.Ref == type)
                this.InitOpenManager.copyPasteObj.activeRange = this.stream.GetString2LE(length);
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadWorksheetCols = function(type, length, aTempCols, oWorksheet, aCellXfs)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerWorksheetsTypes.Col == type )
            {
                var oTempCol = {Max: null, Min: null, col: new AscCommonExcel.Col(oWorksheet, 0)};
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadWorksheetCol(t,l, oTempCol, aCellXfs);
                });
                oTempCol.col.fixOnOpening();
                aTempCols.push(oTempCol);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadWorksheetCol = function(type, length, oTempCol, aCellXfs)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerWorksheetColTypes.BestFit == type )
                oTempCol.col.BestFit = this.stream.GetBool();
            else if ( c_oSerWorksheetColTypes.Hidden == type )
                oTempCol.col.setHidden(this.stream.GetBool());
            else if ( c_oSerWorksheetColTypes.Max == type )
                oTempCol.Max = this.stream.GetULongLE();
            else if ( c_oSerWorksheetColTypes.Min == type )
                oTempCol.Min = this.stream.GetULongLE();
            else if (c_oSerWorksheetColTypes.Style == type) {
                var xfs = aCellXfs[this.stream.GetULongLE()];
                if (xfs) {
                    oTempCol.col.setStyle(xfs);
                }
            } else if ( c_oSerWorksheetColTypes.Width == type )
                oTempCol.col.width = this.stream.GetDoubleLE();
            else if ( c_oSerWorksheetColTypes.CustomWidth == type )
                oTempCol.col.CustomWidth = this.stream.GetBool();
            else if ( c_oSerWorksheetColTypes.OutLevel == type )
                oTempCol.col.outlineLevel = this.stream.GetULongLE();
            else if ( c_oSerWorksheetColTypes.Collapsed == type )
                oTempCol.col.collapsed = this.stream.GetBool();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadSheetFormatPr = function(type, length, oWorksheet)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerSheetFormatPrTypes.DefaultColWidth == type )
                oWorksheet.oSheetFormatPr.dDefaultColWidth = this.stream.GetDoubleLE();
            else if (c_oSerSheetFormatPrTypes.BaseColWidth === type)
            {
                var _nBaseColWidth = this.stream.GetULongLE();
                if (_nBaseColWidth > 0) {
                    oWorksheet.oSheetFormatPr.nBaseColWidth = _nBaseColWidth;
                }
            }
            else if ( c_oSerSheetFormatPrTypes.DefaultRowHeight == type )
            {
                var oAllRow = oWorksheet.getAllRow();
                oAllRow.setHeight(this.stream.GetDoubleLE());
            }
            else if ( c_oSerSheetFormatPrTypes.CustomHeight == type )
            {
                var oAllRow = oWorksheet.getAllRow();
				var CustomHeight = this.stream.GetBool();
				if(CustomHeight)
					oAllRow.setCustomHeight(true);
            }
            else if ( c_oSerSheetFormatPrTypes.ZeroHeight == type )
            {
                var oAllRow = oWorksheet.getAllRow();
				var hd = this.stream.GetBool();
				if(hd)
					oAllRow.setHidden(true);
            }
            else if ( c_oSerSheetFormatPrTypes.OutlineLevelCol == type )
            {
				oWorksheet.oSheetFormatPr.nOutlineLevelCol = this.stream.GetULongLE();
            }
            else if ( c_oSerSheetFormatPrTypes.OutlineLevelRow == type )
            {
                var oAllRow = oWorksheet.getAllRow();
                oAllRow.setOutlineLevel(this.stream.GetULongLE());
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadPageMargins = function(type, length, oPageMargins)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_PageMargins.Left == type )
                oPageMargins.asc_setLeft(this.stream.GetDoubleLE());
            else if ( c_oSer_PageMargins.Top == type )
                oPageMargins.asc_setTop(this.stream.GetDoubleLE());
            else if ( c_oSer_PageMargins.Right == type )
                oPageMargins.asc_setRight(this.stream.GetDoubleLE());
            else if ( c_oSer_PageMargins.Bottom == type )
                oPageMargins.asc_setBottom(this.stream.GetDoubleLE());
			else if ( c_oSer_PageMargins.Header == type )
				oPageMargins.asc_setHeader(this.stream.GetDoubleLE());
			else if ( c_oSer_PageMargins.Footer == type )
				oPageMargins.asc_setFooter(this.stream.GetDoubleLE());
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadPageSetup = function(type, length, oPageSetup)
        {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_PageSetup.BlackAndWhite === type) {
                oPageSetup.blackAndWhite = this.stream.GetBool();
            } else if ( c_oSer_PageSetup.CellComments == type ) {
                oPageSetup.cellComments = this.stream.GetUChar();
            } else if ( c_oSer_PageSetup.Copies == type ) {
                oPageSetup.copies = this.stream.GetULongLE();
            } else if ( c_oSer_PageSetup.Draft == type ) {
                oPageSetup.draft = this.stream.GetBool();
            } else if ( c_oSer_PageSetup.Errors == type ) {
                oPageSetup.errors = this.stream.GetUChar();
            } else if ( c_oSer_PageSetup.FirstPageNumber == type ) {
                var _firstPageNumber = this.stream.GetULongLE();
                if (_firstPageNumber >= 0 && _firstPageNumber < 2147483647) {
                    oPageSetup.firstPageNumber = _firstPageNumber;
                }
            } else if ( c_oSer_PageSetup.FitToHeight == type ) {
                oPageSetup.fitToHeight = this.stream.GetULongLE();
            } else if ( c_oSer_PageSetup.FitToWidth == type ) {
                oPageSetup.fitToWidth = this.stream.GetULongLE();
            } else if ( c_oSer_PageSetup.HorizontalDpi == type ) {
                oPageSetup.horizontalDpi = this.stream.GetULongLE();
            } else if ( c_oSer_PageSetup.Orientation == type ) {
                var byteFormatOrientation = this.stream.GetUChar();
                var byteOrientation = null;
                switch(byteFormatOrientation)
                {
                    case EPageOrientation.pageorientPortrait: byteOrientation = c_oAscPageOrientation.PagePortrait;break;
                    case EPageOrientation.pageorientLandscape: byteOrientation = c_oAscPageOrientation.PageLandscape;break;
                }
                if(null != byteOrientation)
                    oPageSetup.asc_setOrientation(byteOrientation);
            } else if ( c_oSer_PageSetup.PageOrder == type ) {
                oPageSetup.pageOrder = this.stream.GetUChar();
            // } else if ( c_oSer_PageSetup.PaperHeight == type ) {
            //     oPageSetup.height = this.stream.GetDoubleLE();
            } else if ( c_oSer_PageSetup.PaperSize == type ) {
                var bytePaperSize = this.stream.GetUChar();
                var item = DocumentPageSize.getSizeById(bytePaperSize);
                oPageSetup.asc_setWidth(item.w_mm);
                oPageSetup.asc_setHeight(item.h_mm);
            // } else if ( c_oSer_PageSetup.PaperWidth == type ) {
            //     oPageSetup.width = this.stream.GetDoubleLE();
            // } else if ( c_oSer_PageSetup.PaperUnits == type ) {
            //     oPageSetup.paperUnits = this.stream.GetUChar();
            } else if ( c_oSer_PageSetup.Scale == type ) {
                oPageSetup.scale = this.stream.GetULongLE();
            } else if ( c_oSer_PageSetup.UseFirstPageNumber == type ) {
                oPageSetup.useFirstPageNumber = this.stream.GetBool();
            } else if ( c_oSer_PageSetup.UsePrinterDefaults == type ) {
                oPageSetup.usePrinterDefaults = this.stream.GetBool();
            } else if ( c_oSer_PageSetup.VerticalDpi == type ) {
                oPageSetup.verticalDpi = this.stream.GetULongLE();
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadPrintOptions = function (type, length, oPrintOptions) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_PrintOptions.GridLines === type) {
                oPrintOptions.asc_setGridLines(this.stream.GetBool());
            } else if (c_oSer_PrintOptions.Headings === type) {
                oPrintOptions.asc_setHeadings(this.stream.GetBool());
            } else if (c_oSer_PrintOptions.GridLinesSet === type) {
                oPrintOptions.asc_setGridLinesSet(this.stream.GetBool());
            } else if (c_oSer_PrintOptions.HorizontalCentered === type) {
                oPrintOptions.asc_setHorizontalCentered(this.stream.GetBool());
            } else if (c_oSer_PrintOptions.VerticalCentered === type) {
                oPrintOptions.asc_setVerticalCentered(this.stream.GetBool());
            } else {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadHyperlinks = function(type, length, ws)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerWorksheetsTypes.Hyperlink == type )
            {
                var oNewHyperlink = new AscCommonExcel.Hyperlink();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadHyperlink(t,l, ws, oNewHyperlink);
                });
                this.aHyperlinks.push(oNewHyperlink);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadHyperlink = function(type, length, ws, oHyperlink)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerHyperlinkTypes.Ref == type )
                oHyperlink.Ref = ws.getRange2(this.stream.GetString2LE(length));
            else if ( c_oSerHyperlinkTypes.Hyperlink == type )
                oHyperlink.Hyperlink = this.stream.GetString2LE(length);
            else if ( c_oSerHyperlinkTypes.Location == type )
                oHyperlink.setLocation(this.stream.GetString2LE(length));
            else if ( c_oSerHyperlinkTypes.Tooltip == type )
                oHyperlink.Tooltip = this.stream.GetString2LE(length);
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadMergeCells = function(type, length)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerWorksheetsTypes.MergeCell == type )
            {
                this.aMerged.push(this.stream.GetString2LE(length));
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadSheetData = function(type, length, tmp)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
			if ( c_oSerWorksheetsTypes.XlsbPos === type )
            {
				var oldPos = this.stream.GetCurPos();
				this.stream.Seek2(this.stream.GetULongLE());

				tmp.ws.fromXLSB(this.stream, this.stream.XlsbReadRecordType(), tmp, this.aCellXfs, this.aSharedStrings,
					function(tmp) {
						oThis.InitOpenManager.initCellAfterRead(tmp);
					});

				this.stream.Seek2(oldPos);
				res = c_oSerConstants.ReadUnknown;
			}
			else if ( c_oSerWorksheetsTypes.Row === type )
			{
				tmp.pos =  null;
				tmp.len = null;
				tmp.row.clear();
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadRow(t,l, tmp);
                });
				if(null === tmp.row.index) {
					tmp.row.index = tmp.prevRow + 1;
				}
				tmp.row.saveContent();
				tmp.ws.cellsByColRowsCount = Math.max(tmp.ws.cellsByColRowsCount, tmp.row.index + 1);
				tmp.ws.nRowsCount = Math.max(tmp.ws.nRowsCount, tmp.ws.cellsByColRowsCount);
				tmp.prevRow = tmp.row.index;
				tmp.prevCol = -1;
				//читаем ячейки
				if (null !== tmp.pos && null !== tmp.len) {
					var nOldPos = this.stream.GetCurPos();
					this.stream.Seek2(tmp.pos);
					res = this.bcr.Read1(tmp.len, function(t,l){
						return oThis.ReadCells(t,l, tmp);
					});
					this.stream.Seek2(nOldPos);
				}
			}
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadRow = function(type, length, tmp)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerRowTypes.Row == type )
            {
            	var index = this.stream.GetULongLE() - 1;
				tmp.row.setIndex(index);
            }
            else if ( c_oSerRowTypes.Style == type )
            {
                var xfs = this.aCellXfs[this.stream.GetULongLE()];
                if(xfs)
					tmp.row.setStyle(xfs);
            }
            else if ( c_oSerRowTypes.Height == type )
            {
            	var h = this.stream.GetDoubleLE();
				tmp.row.setHeight(h);
                if(AscCommon.CurFileVersion < 2)
					tmp.row.setCustomHeight(true);
            }
            else if ( c_oSerRowTypes.CustomHeight == type )
			{
				var CustomHeight = this.stream.GetBool();
				if(CustomHeight)
					tmp.row.setCustomHeight(true);
			}
            else if ( c_oSerRowTypes.Hidden == type )
			{
				var hd = this.stream.GetBool();
				if(hd)
					tmp.row.setHidden(true);
			}
            else if ( c_oSerRowTypes.OutLevel == type )
            {
                tmp.row.setOutlineLevel(this.stream.GetULongLE());
            }
            else if ( c_oSerRowTypes.Collapsed == type )
            {
                tmp.row.setCollapsed(this.stream.GetBool());
            }
            else if ( c_oSerRowTypes.Cells == type )
            {
				//запоминам место чтобы читать Cells в конце, когда уже зачитан oRow.index
				tmp.pos = this.stream.GetCurPos();
				tmp.len = length;
				res = c_oSerConstants.ReadUnknown;
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadCells = function(type, length, tmp)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerRowTypes.Cell === type )
            {
				tmp.cell.clear();
                tmp.formula.clean();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadCell(t,l, tmp, tmp.cell, tmp.prevRow);
                });
                if (tmp.cell.isNullTextString()) {
                    //set default value in case of empty cell value
                    tmp.cell.setTypeInternal(CellValueType.Number);
                }
                if (tmp.cell.hasRowCol()) {
                    tmp.prevCol = tmp.cell.nCol;
                } else {
                    tmp.prevCol++;
                    tmp.cell.setRowCol(tmp.prevRow, tmp.prevCol);
                }
				this.InitOpenManager.initCellAfterRead(tmp);
			}
			else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
        this.ReadCell = function(type, length, tmp, oCell, nRowIndex)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerCellTypes.Ref === type ){
				var oCellAddress = AscCommon.g_oCellAddressUtils.getCellAddress(this.stream.GetString2LE(length));
				oCell.setRowCol(nRowIndex, oCellAddress.getCol0());
			}
            else if ( c_oSerCellTypes.RefRowCol === type ){
				var nRow = this.stream.GetULongLE();//todo не используем можно убрать
				oCell.setRowCol(nRowIndex, this.stream.GetULongLE());
			}
            else if( c_oSerCellTypes.Style === type )
            {
                var nStyleIndex = this.stream.GetULongLE();
                if(0 != nStyleIndex)
                {
                    var xfs = this.aCellXfs[nStyleIndex];
                    if(null != xfs)
                        oCell.setStyle(xfs);
                }
            }
            else if( c_oSerCellTypes.Type === type )
            {
                switch(this.stream.GetUChar())
                {
                    case ECellTypeType.celltypeBool: oCell.setTypeInternal(CellValueType.Bool);break;
                    case ECellTypeType.celltypeError: oCell.setTypeInternal(CellValueType.Error);break;
                    case ECellTypeType.celltypeNumber: oCell.setTypeInternal(CellValueType.Number);break;
                    case ECellTypeType.celltypeSharedString: oCell.setTypeInternal(CellValueType.String);break;
                }
            }
            else if( c_oSerCellTypes.Formula === type )
            {
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadFormula(t,l, tmp.formula);
                });
            }
			else if (c_oSerCellTypes.Value === type) {
				var val = this.stream.GetDoubleLE();
				if (CellValueType.String === oCell.getType() || CellValueType.Error === oCell.getType()) {
					var ss = this.aSharedStrings[val];
                    if (undefined !== ss) {
                        if (typeof ss === 'string') {
                            oCell.setValueTextInternal(ss);
                        } else {
                            oCell.setValueMultiTextInternal(ss);
                        }
                    }
				} else {
                    oCell.setValueNumberInternal(val);
				}
            }   /*else if (c_oSerCellTypes.CellMetadata === type)
            {
                oCell.cm = this.stream.GetULong();
            }
            else if (c_oSerCellTypes.ValueMetadata === type)
            {
                oCell.vm = this.stream.GetULong();
            }*/
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadFormula = function(type, length, oFormula)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSerFormulaTypes.Aca === type )
                oFormula.aca = this.stream.GetBool();
            else if ( c_oSerFormulaTypes.Bx === type )
                oFormula.bx = this.stream.GetBool();
            else if ( c_oSerFormulaTypes.Ca === type )
                oFormula.ca = this.stream.GetBool();
            else if ( c_oSerFormulaTypes.Del1 === type )
                oFormula.del1 = this.stream.GetBool();
            else if ( c_oSerFormulaTypes.Del2 === type )
                oFormula.del2 = this.stream.GetBool();
            else if ( c_oSerFormulaTypes.Dt2D === type )
                oFormula.dt2d = this.stream.GetBool();
            else if ( c_oSerFormulaTypes.Dtr === type )
                oFormula.dtr = this.stream.GetBool();
            else if ( c_oSerFormulaTypes.R1 === type )
                oFormula.r1 = this.stream.GetString2LE(length);
            else if ( c_oSerFormulaTypes.R2 === type )
                oFormula.r2 = this.stream.GetString2LE(length);
            else if ( c_oSerFormulaTypes.Ref === type )
                oFormula.ref = this.stream.GetString2LE(length);
            else if ( c_oSerFormulaTypes.Si === type )
                oFormula.si = this.stream.GetULongLE();
            else if ( c_oSerFormulaTypes.T === type )
                oFormula.t = this.stream.GetUChar();
            else if ( c_oSerFormulaTypes.Text === type ) {
                oFormula.v = this.stream.GetString2LE(length);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadDrawings = function(type, length, aDrawings, ws)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerWorksheetsTypes.Drawing == type )
            {
                var objectRender = new AscFormat.DrawingObjects();
                var oFlags = {from: false, to: false, pos: false, ext: false, editAs: c_oAscCellAnchorType.cellanchorTwoCell};
                var oNewDrawing = objectRender.createDrawingObject();
                res = this.bcr.Read1(length, function(t, l) {
                    return oThis.ReadDrawing(t, l, oNewDrawing, oFlags);
                });
                if(false != oFlags.from && false != oFlags.to) {
                    oNewDrawing.Type = c_oAscCellAnchorType.cellanchorTwoCell;
                    oNewDrawing.editAs = oFlags.editAs;
                } else if(false != oFlags.from && false != oFlags.ext)
                    oNewDrawing.Type = c_oAscCellAnchorType.cellanchorOneCell;
                else if(false != oFlags.pos && false != oFlags.ext)
                    oNewDrawing.Type = c_oAscCellAnchorType.cellanchorAbsolute;
                oNewDrawing.initAfterSerialize(ws);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadDrawing = function(type, length, oDrawing, oFlags)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_DrawingType.Type == type )
                oDrawing.Type = this.stream.GetUChar();
            else if ( c_oSer_DrawingType.EditAs == type )
                oFlags.editAs = this.stream.GetUChar();
            else if ( c_oSer_DrawingType.From == type )
            {
                oFlags.from = true;
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadFromTo(t,l, oDrawing.from);
                });
            }
            else if ( c_oSer_DrawingType.To == type )
            {
                oFlags.to = true;
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadFromTo(t,l, oDrawing.to);
                });
            }
            else if ( c_oSer_DrawingType.Pos == type )
            {
                oFlags.pos = true;
                if(null == oDrawing.Pos)
                    oDrawing.Pos = {};
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadPos(t,l, oDrawing.Pos);
                });
            }
            else if ( c_oSer_DrawingType.Ext == type )
            {
                oFlags.ext = true;
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadExt(t,l, oDrawing.ext);
                });
            }
            else if ( c_oSer_DrawingType.Pic == type )
            {
                oDrawing.image = new Image();
                res = this.bcr.Read1(length, function(t,l){
                    //return oThis.ReadPic(t,l, oDrawing.Pic);
                    return oThis.ReadPic(t,l, oDrawing);
                });
            }
            /** proprietary begin **/
            else if ( c_oSer_DrawingType.GraphicFrame == type )
            {
                //todo удалить
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadGraphicFrame(t, l, oDrawing);
                });
            }
            /** proprietary end **/
            else if ( c_oSer_DrawingType.pptxDrawing == type )
            {
                oDrawing.graphicObject = this.ReadPptxDrawing();

                if(oDrawing.graphicObject && !oDrawing.graphicObject.isSupported())
                {
                    let nPos = this.bcr.stream.cur;
                    let type_ = this.bcr.stream.GetUChar();
                    let length_ = this.bcr.stream.GetULongLE();
                    this.bcr.stream.Seek2(nPos);
                    if(type_ === c_oSer_DrawingType.pptxDrawingAlternative)
                    {
                        res = oThis.bcr.Read1(length_, function(t,l){
                            if(t === c_oSer_DrawingType.pptxDrawingAlternative){
                                oDrawing.graphicObject = pptx_content_loader.ReadGraphicObject2(oThis.stream, oThis.curWorksheet, oThis.curWorksheet.getDrawingDocument());
                                return c_oSerConstants.ReadOk;
                            }
                            return c_oSerConstants.ReadUnknown;
                        });
                    }
                }
            }
            else if( c_oSer_DrawingType.ClientData == type ) {
                var oClientData = new AscFormat.CClientData();
                res = this.bcr.Read2Spreadsheet(length, function (t, l) {
                    return oThis.ReadClientData(t, l, oClientData);
                });
                oDrawing.clientData = oClientData;
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };

        this.ReadPptxDrawing = function () {
            return pptx_content_loader.ReadGraphicObject(this.stream, this.curWorksheet, this.curWorksheet.getDrawingDocument());
        };
        this.ReadGraphicFrame = function (type, length, oDrawing) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_DrawingType.Chart2 == type) {
                var oNewChartSpace = new AscFormat.CChartSpace();
                var oBinaryChartReader = new AscCommon.BinaryChartReader(this.stream);
                res = oBinaryChartReader.ExternalReadCT_ChartSpace(length, oNewChartSpace, this.curWorksheet);
                if(oNewChartSpace.hasCharts()) {
                    oDrawing.graphicObject = oNewChartSpace;
                    oNewChartSpace.setBDeleted(false);
                    if(oNewChartSpace.setDrawingBase)
                    {
                        oNewChartSpace.setDrawingBase(oDrawing);
                    }
                }
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadFromTo = function(type, length, oFromTo)
        {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_DrawingFromToType.Col === type)
            {
                oFromTo.col = this.stream.GetULongLE();
            }
            else if (c_oSer_DrawingFromToType.ColOff === type)
            {
                oFromTo.colOff = this.stream.GetDoubleLE();
            }
            else if (c_oSer_DrawingFromToType.Row === type)
            {
                oFromTo.row = this.stream.GetULongLE();
            }
            else if (c_oSer_DrawingFromToType.RowOff === type)
            {
                oFromTo.rowOff = this.stream.GetDoubleLE();
            }
            else
            {
                res = c_oSerConstants.ReadUnknown;
            }
            return res;
        };
        this.ReadPos = function(type, length, oPos)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_DrawingPosType.X == type )
                oPos.X = this.stream.GetDoubleLE();
            else if ( c_oSer_DrawingPosType.Y == type )
                oPos.Y = this.stream.GetDoubleLE();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadExt = function(type, length, oExt)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_DrawingExtType.Cx == type )
                oExt.cx = this.stream.GetDoubleLE();
            else if ( c_oSer_DrawingExtType.Cy == type )
                oExt.cy = this.stream.GetDoubleLE();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadClientData = function(type, length, oClientData)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_DrawingClientDataType.fLocksWithSheet == type )
                oClientData.fLocksWithSheet = this.stream.GetBool();
            else if ( c_oSer_DrawingClientDataType.fPrintsWithSheet == type )
                oClientData.fPrintsWithSheet = this.stream.GetBool();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadPic = function(type, length, oDrawing)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_DrawingType.PicSrc == type )
            {
                var nIndex = this.stream.GetULongLE();
                var src = this.oMediaArray[nIndex];
                if(null != src)
                {
                  oDrawing.image.src = src;
                }
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadComments = function(type, length, oWorksheet)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSerWorksheetsTypes.Comment == type && AscCommonExcel.asc_CCommentCoords)
            {
                var oCommentCoords = new AscCommonExcel.asc_CCommentCoords();
                var aCommentData = [];
                var oAdditionalData = {isThreadedComment: false};
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadComment(t,l, oCommentCoords, aCommentData, oAdditionalData);
                });

                this.InitOpenManager.prepareComments(oWorksheet, oCommentCoords, aCommentData, oAdditionalData);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadComment = function(type, length, oCommentCoords, aCommentData, oAdditionalData)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_Comments.Row == type )
                oCommentCoords.nRow = this.stream.GetULongLE();
            else if ( c_oSer_Comments.Col == type )
                oCommentCoords.nCol = this.stream.GetULongLE();
            else if ( c_oSer_Comments.CommentDatas == type )
            {
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadCommentDatas(t,l, aCommentData);
                });
            }
            else if ( c_oSer_Comments.Left == type )
                oCommentCoords.nLeft = this.stream.GetULongLE();
            else if ( c_oSer_Comments.LeftOffset == type )
                oCommentCoords.nLeftOffset = this.stream.GetULongLE();
            else if ( c_oSer_Comments.Top == type )
                oCommentCoords.nTop = this.stream.GetULongLE();
            else if ( c_oSer_Comments.TopOffset == type )
                oCommentCoords.nTopOffset = this.stream.GetULongLE();
            else if ( c_oSer_Comments.Right == type )
                oCommentCoords.nRight = this.stream.GetULongLE();
            else if ( c_oSer_Comments.RightOffset == type )
                oCommentCoords.nRightOffset = this.stream.GetULongLE();
            else if ( c_oSer_Comments.Bottom == type )
                oCommentCoords.nBottom = this.stream.GetULongLE();
            else if ( c_oSer_Comments.BottomOffset == type )
                oCommentCoords.nBottomOffset = this.stream.GetULongLE();
            else if ( c_oSer_Comments.LeftMM == type )
                oCommentCoords.dLeftMM = this.stream.GetDoubleLE();
            else if ( c_oSer_Comments.TopMM == type )
                oCommentCoords.dTopMM = this.stream.GetDoubleLE();
            else if ( c_oSer_Comments.WidthMM == type )
                oCommentCoords.dWidthMM = this.stream.GetDoubleLE();
            else if ( c_oSer_Comments.HeightMM == type )
                oCommentCoords.dHeightMM = this.stream.GetDoubleLE();
            else if ( c_oSer_Comments.MoveWithCells == type )
                oCommentCoords.bMoveWithCells = this.stream.GetBool();
            else if ( c_oSer_Comments.SizeWithCells == type )
                oCommentCoords.bSizeWithCells = this.stream.GetBool();
            else if ( c_oSer_Comments.ThreadedComment == type ) {
                if (aCommentData.length > 0) {
                    oAdditionalData.isThreadedComment = true;
                    var commentData = aCommentData[0];
                    commentData.asc_putSolved(false);
                    commentData.aReplies = [];
                    // commentData.aMentions = [];
                    res = this.bcr.Read1(length, function(t, l) {
                        return oThis.ReadThreadedComment(t, l, commentData);
                    });
                } else {
                    res = c_oSerConstants.ReadUnknown;
                }
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadCommentDatas = function(type, length, aCommentData)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_Comments.CommentData === type )
            {
                var oCommentData = new Asc.asc_CCommentData();
                oCommentData.asc_putDocumentFlag(false);
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadCommentData(t,l,oCommentData);
                });

				if (oCommentData.asc_getDocumentFlag()) {
					oCommentData.nId = "doc_" + (this.wb.aComments.length + 1);
				} else {
					oCommentData.wsId = this.curWorksheet.Id;
					oCommentData.nId = "sheet" + oCommentData.wsId + "_" + (this.curWorksheet.aComments.length + 1);
				}
                aCommentData.push(oCommentData);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadCommentData = function(type, length, oCommentData)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_CommentData.Text == type )
                oCommentData.asc_putText(this.stream.GetString2LE(length));
            else if ( c_oSer_CommentData.Time == type )
            {
                var dateMs = AscCommon.getTimeISO8601(this.stream.GetString2LE(length));
                if(!isNaN(dateMs))
                    oCommentData.asc_putTime(dateMs + "");
            }
            else if ( c_oSer_CommentData.OOTime == type )
            {
                var dateMs = AscCommon.getTimeISO8601(this.stream.GetString2LE(length));
                if(!isNaN(dateMs))
                    oCommentData.asc_putOnlyOfficeTime(dateMs + "");
            }
            else if ( c_oSer_CommentData.UserId == type )
                oCommentData.asc_putUserId(this.stream.GetString2LE(length));
            else if ( c_oSer_CommentData.UserName == type )
                oCommentData.asc_putUserName(this.stream.GetString2LE(length));
            else if ( c_oSer_CommentData.UserData == type )
                oCommentData.asc_putUserData(this.stream.GetString2LE(length));
            else if ( c_oSer_CommentData.QuoteText == type )
                oCommentData.asc_putQuoteText(this.stream.GetString2LE(length));
            else if ( c_oSer_CommentData.Replies == type )
            {
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadReplies(t,l, oCommentData);
                });
            }
            else if ( c_oSer_CommentData.Solved == type )
                oCommentData.asc_putSolved(this.stream.GetBool());
            else if ( c_oSer_CommentData.Document == type )
                oCommentData.asc_putDocumentFlag(this.stream.GetBool());
            else if ( c_oSer_CommentData.Guid == type )
                oCommentData.asc_putGuid(this.stream.GetString2LE(length));
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadConditionalFormatting = function (type, length, oConditionalFormatting) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            var oConditionalFormattingRule = null;
            if (c_oSer_ConditionalFormatting.Pivot === type)
                oConditionalFormatting.pivot = this.stream.GetBool();
            else if (c_oSer_ConditionalFormatting.SqRef === type) {
                oConditionalFormatting.setSqRef(this.stream.GetString2LE(length));
            }
            else if (c_oSer_ConditionalFormatting.ConditionalFormattingRule === type) {
                oConditionalFormattingRule = new AscCommonExcel.CConditionalFormattingRule();
                var ext = {isExt: false};
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadConditionalFormattingRule(t, l, oConditionalFormattingRule, ext);
                });
                oConditionalFormatting.aRules.push(oConditionalFormattingRule);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadConditionalFormattingRule = function (type, length, oConditionalFormattingRule, ext) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            var oConditionalFormattingRuleElement = null;

            if (c_oSer_ConditionalFormattingRule.AboveAverage === type)
                oConditionalFormattingRule.aboveAverage = this.stream.GetBool();
            else if (c_oSer_ConditionalFormattingRule.Bottom === type)
                oConditionalFormattingRule.bottom = this.stream.GetBool();
            else if (c_oSer_ConditionalFormattingRule.DxfId === type)
            {
                var DxfId = this.stream.GetULongLE();
                oConditionalFormattingRule.dxf = this.InitOpenManager.Dxfs[DxfId];
            }
			else if (c_oSer_ConditionalFormattingRule.Dxf === type)
			{
				var oDxf = new AscCommonExcel.CellXfs();
				res = this.bcr.Read1(length, function(t,l){
					return oThis.InitOpenManager.oReadResult.stylesTableReader.ReadDxf(t,l,oDxf);
				});
				oConditionalFormattingRule.dxf = oDxf;
			}
            else if (c_oSer_ConditionalFormattingRule.EqualAverage === type)
                oConditionalFormattingRule.equalAverage = this.stream.GetBool();
            else if (c_oSer_ConditionalFormattingRule.Operator === type)
                oConditionalFormattingRule.operator = this.stream.GetUChar();
            else if (c_oSer_ConditionalFormattingRule.Percent === type)
                oConditionalFormattingRule.percent = this.stream.GetBool();
            else if (c_oSer_ConditionalFormattingRule.Priority === type)
                oConditionalFormattingRule.priority = this.stream.GetULongLE();
            else if (c_oSer_ConditionalFormattingRule.Rank === type)
                oConditionalFormattingRule.rank = this.stream.GetULongLE();
            else if (c_oSer_ConditionalFormattingRule.StdDev === type)
                oConditionalFormattingRule.stdDev = this.stream.GetULongLE();
            else if (c_oSer_ConditionalFormattingRule.StopIfTrue === type)
                oConditionalFormattingRule.stopIfTrue = this.stream.GetBool();
            else if (c_oSer_ConditionalFormattingRule.Text === type)
                oConditionalFormattingRule.text = this.stream.GetString2LE(length);
            else if (c_oSer_ConditionalFormattingRule.TimePeriod === type)
                oConditionalFormattingRule.timePeriod = this.stream.GetString2LE(length);
            else if (c_oSer_ConditionalFormattingRule.Type === type)
                oConditionalFormattingRule.type = this.stream.GetUChar();
            else if (c_oSer_ConditionalFormattingRule.ColorScale === type) {
                oConditionalFormattingRuleElement = new AscCommonExcel.CColorScale();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadColorScale(t, l, oConditionalFormattingRuleElement);
                });
                oConditionalFormattingRule.aRuleElements.push(oConditionalFormattingRuleElement);
            } else if (c_oSer_ConditionalFormattingRule.DataBar === type) {
                oConditionalFormattingRuleElement = new AscCommonExcel.CDataBar();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadDataBar(t, l, oConditionalFormattingRuleElement);
                });
                oConditionalFormattingRule.aRuleElements.push(oConditionalFormattingRuleElement);
            } else if (c_oSer_ConditionalFormattingRule.FormulaCF === type) {
                oConditionalFormattingRuleElement = new AscCommonExcel.CFormulaCF();
                oConditionalFormattingRuleElement.Text = this.stream.GetString2LE(length);
                oConditionalFormattingRule.aRuleElements.push(oConditionalFormattingRuleElement);
            } else if (c_oSer_ConditionalFormattingRule.IconSet === type) {
                oConditionalFormattingRuleElement = new AscCommonExcel.CIconSet();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadIconSet(t, l, oConditionalFormattingRuleElement);
                });
                oConditionalFormattingRule.aRuleElements.push(oConditionalFormattingRuleElement);
            } else if (c_oSer_ConditionalFormattingRule.isExt === type) {
                ext.isExt = this.stream.GetBool();
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadColorScale = function (type, length, oColorScale) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            var oObject = null;
            if (c_oSer_ConditionalFormattingRuleColorScale.CFVO === type) {
                oObject = new AscCommonExcel.CConditionalFormatValueObject();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadCFVO(t, l, oObject);
                });
                oColorScale.aCFVOs.push(oObject);
            } else if (c_oSer_ConditionalFormattingRuleColorScale.Color === type) {
				var color = ReadColorSpreadsheet2(this.bcr, length);
				if (null != color) {
					oColorScale.aColors.push(color);
				}
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadDataBar = function (type, length, oDataBar) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            var oObject = null;
            if (c_oSer_ConditionalFormattingDataBar.MaxLength === type)
                oDataBar.MaxLength = this.stream.GetULongLE();
            else if (c_oSer_ConditionalFormattingDataBar.MinLength === type)
                oDataBar.MinLength = this.stream.GetULongLE();
            else if (c_oSer_ConditionalFormattingDataBar.ShowValue === type)
                oDataBar.ShowValue = this.stream.GetBool();
            else if (c_oSer_ConditionalFormattingDataBar.Color === type) {
				var color = ReadColorSpreadsheet2(this.bcr, length);
				if (color) {
					oDataBar.Color = color;
				}
			} else if (c_oSer_ConditionalFormattingDataBar.NegativeColor === type) {
				var color = ReadColorSpreadsheet2(this.bcr, length);
				if (color) {
					oDataBar.NegativeColor = color;
				}
			} else if (c_oSer_ConditionalFormattingDataBar.BorderColor === type) {
				var color = ReadColorSpreadsheet2(this.bcr, length);
				if (color) {
					oDataBar.BorderColor = color;
				}
			} else if (c_oSer_ConditionalFormattingDataBar.AxisColor === type) {
				var color = ReadColorSpreadsheet2(this.bcr, length);
				if (color) {
					oDataBar.AxisColor = color;
				} else {
					//TODO наличие оси определяется по наличию AxisColor при отрисовке. других маркеров нет. пересмотреть!
					oDataBar.AxisColor = new AscCommonExcel.RgbColor(0);
				}
			} else if (c_oSer_ConditionalFormattingDataBar.NegativeBorderColor === type) {
				var color = ReadColorSpreadsheet2(this.bcr, length);
				if (color) {
					oDataBar.NegativeBorderColor = color;
				}
			} else if (c_oSer_ConditionalFormattingDataBar.AxisPosition === type) {
				oDataBar.AxisPosition = this.stream.GetULongLE();
			} else if (c_oSer_ConditionalFormattingDataBar.Direction === type) {
				oDataBar.Direction = this.stream.GetULongLE();
			} else if (c_oSer_ConditionalFormattingDataBar.GradientEnabled === type) {
				oDataBar.Gradient = this.stream.GetBool();
			} else if (c_oSer_ConditionalFormattingDataBar.NegativeBarColorSameAsPositive === type) {
				oDataBar.NegativeBarColorSameAsPositive = this.stream.GetBool();
			} else if (c_oSer_ConditionalFormattingDataBar.NegativeBarBorderColorSameAsPositive === type) {
				oDataBar.NegativeBarBorderColorSameAsPositive = this.stream.GetBool();
            } else if (c_oSer_ConditionalFormattingDataBar.CFVO === type) {
                oObject = new AscCommonExcel.CConditionalFormatValueObject();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadCFVO(t, l, oObject);
                });
                oDataBar.aCFVOs.push(oObject);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadIconSet = function (type, length, oIconSet) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            var oObject = null;
            if (c_oSer_ConditionalFormattingIconSet.IconSet === type)
                oIconSet.IconSet = this.stream.GetUChar();
            else if (c_oSer_ConditionalFormattingIconSet.Percent === type)
                oIconSet.Percent = this.stream.GetBool();
            else if (c_oSer_ConditionalFormattingIconSet.Reverse === type)
                oIconSet.Reverse = this.stream.GetBool();
            else if (c_oSer_ConditionalFormattingIconSet.ShowValue === type)
                oIconSet.ShowValue = this.stream.GetBool();
            else if (c_oSer_ConditionalFormattingIconSet.CFVO === type) {
                oObject = new AscCommonExcel.CConditionalFormatValueObject();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadCFVO(t, l, oObject);
                });
                oIconSet.aCFVOs.push(oObject);
			} else if (c_oSer_ConditionalFormattingIconSet.CFIcon === type) {
				oObject = new AscCommonExcel.CConditionalFormatIconSet();
				res = this.bcr.Read1(length, function(t, l) {
					return oThis.ReadCFIS(t, l, oObject);
				});
				oIconSet.aIconSets.push(oObject);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadCFVO = function (type, length, oCFVO) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_ConditionalFormattingValueObject.Gte === type)
                oCFVO.Gte = this.stream.GetBool();
            else if (c_oSer_ConditionalFormattingValueObject.Type === type)
                oCFVO.Type = this.stream.GetUChar();
			else if (c_oSer_ConditionalFormattingValueObject.Val === type || c_oSer_ConditionalFormattingValueObject.Formula === type)
                oCFVO.Val = this.stream.GetString2LE(length);
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadCFIS = function (type, length, oCFVO) {
			var res = c_oSerConstants.ReadOk;
			if (c_oSer_ConditionalFormattingIcon.iconSet === type)
				oCFVO.IconSet = this.stream.GetLong();
			else if (c_oSer_ConditionalFormattingIcon.iconId === type)
				oCFVO.IconId = this.stream.GetLong();
			else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
        this.ReadSheetViews = function (type, length, aSheetViews) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            var oSheetView = null;

            if (c_oSerWorksheetsTypes.SheetView === type && 0 == aSheetViews.length) {
                oSheetView = new AscCommonExcel.asc_CSheetViewSettings();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadSheetView(t, l, oSheetView);
                });
                aSheetViews.push(oSheetView);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadSheetView = function (type, length, oSheetView) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
			if (c_oSer_SheetView.ColorId === type) {
				this.stream.GetLong();
			} else if (c_oSer_SheetView.DefaultGridColor === type) {
				this.stream.GetBool();
			} else if (c_oSer_SheetView.RightToLeft === type) {
                oSheetView.rightToLeft = this.stream.GetBool();
			} else if (c_oSer_SheetView.ShowFormulas === type) {
				oSheetView.showFormulas = this.stream.GetBool();
			} else if (c_oSer_SheetView.ShowGridLines === type) {
				oSheetView.showGridLines = this.stream.GetBool();
			} else if (c_oSer_SheetView.ShowOutlineSymbols === type) {
				this.stream.GetBool();
			} else if (c_oSer_SheetView.ShowRowColHeaders === type) {
				oSheetView.showRowColHeaders = this.stream.GetBool();
			} else if (c_oSer_SheetView.ShowRuler === type) {
				this.stream.GetBool();
			} else if (c_oSer_SheetView.ShowWhiteSpace === type) {
				this.stream.GetBool();
			} else if (c_oSer_SheetView.ShowZeros === type) {
                oSheetView.showZeros = this.stream.GetBool();
			} else if (c_oSer_SheetView.TabSelected === type) {
				this.stream.GetBool();
			} else if (c_oSer_SheetView.TopLeftCell === type) {
                var _topLeftCell = AscCommonExcel.g_oRangeCache.getAscRange(this.stream.GetString2LE(length));
                if (_topLeftCell) {
                    oSheetView.topLeftCell = new Asc.Range(_topLeftCell.c1, _topLeftCell.r1, _topLeftCell.c1, _topLeftCell.r1);
                }
			} else if (c_oSer_SheetView.View === type) {
                oSheetView.view = this.stream.GetUChar();
                //TODO while don't support other view type
                if (oSheetView.view > 1) {
                    oSheetView.view = Asc.c_oAscESheetViewType.normal;
                }
			} else if (c_oSer_SheetView.WindowProtection === type) {
				this.stream.GetBool();
			} else if (c_oSer_SheetView.WorkbookViewId === type) {
				this.stream.GetLong();
			} else if (c_oSer_SheetView.ZoomScale === type) {
				oSheetView.asc_setZoomScale(this.stream.GetLong());
			} else if (c_oSer_SheetView.ZoomScaleNormal === type) {
				this.stream.GetLong();
			} else if (c_oSer_SheetView.ZoomScalePageLayoutView === type) {
				this.stream.GetLong();
			} else if (c_oSer_SheetView.ZoomScaleSheetLayoutView === type) {
				this.stream.GetLong();
            } else if (c_oSer_SheetView.Pane === type) {
                oSheetView.pane = new AscCommonExcel.asc_CPane();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadPane(t, l, oSheetView.pane);
                });
                oSheetView.pane.init();
			} else if (c_oSer_SheetView.Selection === type) {
				this.curWorksheet.selectionRange.clean();
				res = this.bcr.Read1(length, function (t, l) {
					return oThis.ReadSelection(t, l, oThis.curWorksheet.selectionRange);
				});
				this.curWorksheet.selectionRange.update();
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadPane = function (type, length, oPane) {
            var res = c_oSerConstants.ReadOk;
			if (c_oSer_Pane.ActivePane === type)
				this.stream.GetUChar();
			else if (c_oSer_Pane.State === type)
				oPane.state = this.stream.GetString2LE(length);
            else if (c_oSer_Pane.TopLeftCell === type)
				oPane.topLeftCell = this.stream.GetString2LE(length);
			else if (c_oSer_Pane.XSplit === type)
				oPane.xSplit = this.stream.GetDoubleLE();
			else if (c_oSer_Pane.YSplit === type)
				oPane.ySplit = this.stream.GetDoubleLE();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadSelection = function (type, length, selectionRange) {
			var res = c_oSerConstants.ReadOk;
			if (c_oSer_Selection.ActiveCell === type) {
				var activeCell = AscCommonExcel.g_oRangeCache.getAscRange(this.stream.GetString2LE(length));
				if (activeCell) {
					selectionRange.activeCell = new AscCommon.CellBase(activeCell.r1, activeCell.c1);
				}
			} else if (c_oSer_Selection.ActiveCellId === type) {
				selectionRange.activeCellId = this.stream.GetLong();
			} else if (c_oSer_Selection.Sqref === type) {
				var sqRef = this.stream.GetString2LE(length);
				var selectionNew = AscCommonExcel.g_oRangeCache.getRangesFromSqRef(sqRef);
				if (selectionNew.length > 0) {
					selectionRange.ranges = selectionNew;
				}
			} else if (c_oSer_Selection.Pane === type) {
				this.stream.GetUChar();
			} else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
        this.ReadSheetPr = function (type, length, oSheetPr) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_SheetPr.CodeName === type)
                oSheetPr.CodeName = this.stream.GetString2LE(length);
            else if (c_oSer_SheetPr.EnableFormatConditionsCalculation === type)
                oSheetPr.EnableFormatConditionsCalculation = this.stream.GetBool();
            else if (c_oSer_SheetPr.FilterMode === type)
                oSheetPr.FilterMode = this.stream.GetBool();
            else if (c_oSer_SheetPr.Published === type)
                oSheetPr.Published = this.stream.GetBool();
            else if (c_oSer_SheetPr.SyncHorizontal === type)
                oSheetPr.SyncHorizontal = this.stream.GetBool();
            else if (c_oSer_SheetPr.SyncRef === type)
                oSheetPr.SyncRef = this.stream.GetString2LE(length);
            else if (c_oSer_SheetPr.SyncVertical === type)
                oSheetPr.SyncVertical = this.stream.GetBool();
            else if (c_oSer_SheetPr.TransitionEntry === type)
                oSheetPr.TransitionEntry = this.stream.GetBool();
            else if (c_oSer_SheetPr.TransitionEvaluation === type)
                oSheetPr.TransitionEvaluation = this.stream.GetBool();
            else if (c_oSer_SheetPr.TabColor === type) {
				var color = ReadColorSpreadsheet2(this.bcr, length);
				if (color) {
					oSheetPr.TabColor = color;
				}
			} else if (c_oSer_SheetPr.PageSetUpPr === type) {
				res = this.bcr.Read1(length, function (t, l) {
					return oThis.ReadPageSetUpPr(t, l, oSheetPr);
				});
			} else if (c_oSer_SheetPr.OutlinePr === type) {
				res = this.bcr.Read1(length, function (t, l) {
					return oThis.ReadOutlinePr(t, l, oSheetPr);
				});
            } else
                res = c_oSerConstants.ReadUnknown;

            return res;
        };
		this.ReadOutlinePr = function (type, length, oSheetPr) {
			var oThis = this;
			var res = c_oSerConstants.ReadOk;
			if (c_oSer_SheetPr.ApplyStyles === type) {
				oSheetPr.ApplyStyles = this.stream.GetBool();
			} else if (c_oSer_SheetPr.ShowOutlineSymbols === type) {
				oSheetPr.ShowOutlineSymbols = this.stream.GetBool();
			} else if (c_oSer_SheetPr.SummaryBelow === type) {
				oSheetPr.SummaryBelow = this.stream.GetBool();
			} else if (c_oSer_SheetPr.SummaryRight === type) {
				oSheetPr.SummaryRight = this.stream.GetBool();
			} else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
		this.ReadPageSetUpPr = function (type, length, oSheetPr) {
			var oThis = this;
			var res = c_oSerConstants.ReadOk;
			if (c_oSer_SheetPr.AutoPageBreaks === type) {
				oSheetPr.AutoPageBreaks = this.stream.GetBool();
			} else if (c_oSer_SheetPr.FitToPage === type) {
				oSheetPr.FitToPage = this.stream.GetBool();
			} else
				res = c_oSerConstants.ReadUnknown;
			return res;
		};
		this.ReadSparklineGroups = function (type, length, oWorksheet) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_Sparkline.SparklineGroup === type) {
				var newSparklineGroup = new AscCommonExcel.sparklineGroup(true);
                newSparklineGroup.setWorksheet(oWorksheet);
				res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadSparklineGroup(t, l, newSparklineGroup);
                });
                oWorksheet.aSparklineGroups.push(newSparklineGroup);
			} else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadHeaderFooter = function (type, length, headerFooter) {
            var res = c_oSerConstants.ReadOk;
            var sVal;
            if (c_oSer_HeaderFooter.AlignWithMargins === type) {
                headerFooter.setAlignWithMargins(this.stream.GetBool());
            } else if (c_oSer_HeaderFooter.DifferentFirst === type) {
                headerFooter.setDifferentFirst(this.stream.GetBool());
            } else if (c_oSer_HeaderFooter.DifferentOddEven === type) {
                headerFooter.setDifferentOddEven(this.stream.GetBool());
            } else if (c_oSer_HeaderFooter.ScaleWithDoc === type) {
                headerFooter.setScaleWithDoc(this.stream.GetBool());
            } else if (c_oSer_HeaderFooter.EvenFooter === type) {
				sVal = this.stream.GetString2LE(length);
				if(sVal) {
					headerFooter.setEvenFooter(sVal);
                }
            } else if (c_oSer_HeaderFooter.EvenHeader === type) {
				sVal = this.stream.GetString2LE(length);
				if(sVal) {
					headerFooter.setEvenHeader(sVal);
				}
            } else if (c_oSer_HeaderFooter.FirstFooter === type) {
				sVal = this.stream.GetString2LE(length);
				if(sVal) {
					headerFooter.setFirstFooter(sVal);
				}
            } else if (c_oSer_HeaderFooter.FirstHeader === type) {
				sVal = this.stream.GetString2LE(length);
				if(sVal) {
					headerFooter.setFirstHeader(sVal);
				}
            } else if (c_oSer_HeaderFooter.OddFooter === type) {
				sVal = this.stream.GetString2LE(length);
				if(sVal) {
					headerFooter.setOddFooter(sVal);
				}
            } else if (c_oSer_HeaderFooter.OddHeader === type) {
				sVal = this.stream.GetString2LE(length);
				if(sVal) {
					headerFooter.setOddHeader(sVal);
				}
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadRowColBreaks = function (type, length, breaks) {
            let oThis = this;
            let res = c_oSerConstants.ReadOk;
            let val;
            if (c_oSer_RowColBreaks.Count === type) {
				val = this.stream.GetLong();
                breaks.setCount(val);
            } else if (c_oSer_RowColBreaks.ManualBreakCount === type) {
				val = this.stream.GetLong();
                breaks.setManualBreakCount(val);
            } else if (c_oSer_RowColBreaks.Break === type) {
                var brk = new AscCommonExcel.CBreak();
                res = this.bcr.Read1(length, function(t, l) {
                    return oThis.ReadRowColBreak(t, l, brk);
                });
				breaks.pushBreak(brk);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadRowColBreak = function (type, length, brk) {
            let res = c_oSerConstants.ReadOk;
            let val;
            if (c_oSer_RowColBreaks.Id === type) {
            	val = this.stream.GetLong();
                brk.setId(val);
            } else if (c_oSer_RowColBreaks.Man === type) {
				val = this.stream.GetBool();
            	brk.setMan(val);
            } else if (c_oSer_RowColBreaks.Max === type) {
                val = this.stream.GetLong();
				brk.setMax(val);
            } else if (c_oSer_RowColBreaks.Min === type) {
                val = this.stream.GetLong();
				brk.setMin(val);
            } else if (c_oSer_RowColBreaks.Pt === type) {
				val = this.stream.GetBool();
				brk.setPt(val);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadLegacyDrawingHF = function (type, length, legacyDrawingHF) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_LegacyDrawingHF.Drawings === type) {
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadLegacyDrawingHFDrawings(t, l, legacyDrawingHF.drawings);
                });
            } else if (c_oSer_LegacyDrawingHF.Cfe === type) {
                legacyDrawingHF.cfe = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Cff === type) {
                legacyDrawingHF.cff = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Cfo === type) {
                legacyDrawingHF.cfo = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Che === type) {
                legacyDrawingHF.che = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Chf === type) {
                legacyDrawingHF.chf = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Cho === type) {
                legacyDrawingHF.cho = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Lfe === type) {
                legacyDrawingHF.lfe = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Lff === type) {
                legacyDrawingHF.lff = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Lfo === type) {
                legacyDrawingHF.lfo = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Lhe === type) {
                legacyDrawingHF.lhe = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Lhf === type) {
                legacyDrawingHF.lhf = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Lho === type) {
                legacyDrawingHF.lho = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Rfe === type) {
                legacyDrawingHF.rfe = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Rff === type) {
                legacyDrawingHF.rff = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Rfo === type) {
                legacyDrawingHF.rfo = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Rhe === type) {
                legacyDrawingHF.rhe = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Rhf === type) {
                legacyDrawingHF.rhf = this.stream.GetBool();
            } else if (c_oSer_LegacyDrawingHF.Rho === type) {
                legacyDrawingHF.rho = this.stream.GetBool();
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadLegacyDrawingHFDrawings = function (type, length, drawings) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_LegacyDrawingHF.Drawing === type) {
                var drawing = new AscCommonExcel.CLegacyDrawingHFDrawing();
                res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadLegacyDrawingHFDrawing(t, l, drawing);
                });
                if (null !== drawing.id && null !== drawing.graphicObject) {
                    drawings.push(drawing);
                }
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadLegacyDrawingHFDrawing = function (type, length, drawing) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_LegacyDrawingHF.DrawingId === type) {
                drawing.id = this.stream.GetString2LE(length);
            } else if (c_oSer_LegacyDrawingHF.DrawingShape === type) {
                var graphicObject = this.ReadPptxDrawing();
                if (graphicObject) {
                    drawing.graphicObject = graphicObject;
                }
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadSparklineGroup = function (type, length, oSparklineGroup) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_Sparkline.ManualMax === type) {
				oSparklineGroup.manualMax = this.stream.GetDoubleLE();
			} else if (c_oSer_Sparkline.ManualMin === type) {
				oSparklineGroup.manualMin = this.stream.GetDoubleLE();
			} else if (c_oSer_Sparkline.LineWeight === type) {
				oSparklineGroup.lineWeight = this.stream.GetDoubleLE();
			} else if (c_oSer_Sparkline.Type === type) {
				oSparklineGroup.type = this.stream.GetUChar();
			} else if (c_oSer_Sparkline.DateAxis === type) {
				oSparklineGroup.dateAxis = this.stream.GetBool();
			} else if (c_oSer_Sparkline.DisplayEmptyCellsAs === type) {
				oSparklineGroup.displayEmptyCellsAs = this.stream.GetUChar();
			} else if (c_oSer_Sparkline.Markers === type) {
				oSparklineGroup.markers = this.stream.GetBool();
			} else if (c_oSer_Sparkline.High === type) {
				oSparklineGroup.high = this.stream.GetBool();
			} else if (c_oSer_Sparkline.Low === type) {
				oSparklineGroup.low = this.stream.GetBool();
			} else if (c_oSer_Sparkline.First === type) {
				oSparklineGroup.first = this.stream.GetBool();
			} else if (c_oSer_Sparkline.Last === type) {
				oSparklineGroup.last = this.stream.GetBool();
			} else if (c_oSer_Sparkline.Negative === type) {
				oSparklineGroup.negative = this.stream.GetBool();
			} else if (c_oSer_Sparkline.DisplayXAxis === type) {
				oSparklineGroup.displayXAxis = this.stream.GetBool();
			} else if (c_oSer_Sparkline.DisplayHidden === type) {
				oSparklineGroup.displayHidden = this.stream.GetBool();
			} else if (c_oSer_Sparkline.MinAxisType === type) {
				oSparklineGroup.minAxisType = this.stream.GetUChar();
			} else if (c_oSer_Sparkline.MaxAxisType === type) {
				oSparklineGroup.maxAxisType = this.stream.GetUChar();
			} else if (c_oSer_Sparkline.RightToLeft === type) {
				oSparklineGroup.rightToLeft = this.stream.GetBool();
			} else if (c_oSer_Sparkline.ColorSeries === type) {
				oSparklineGroup.colorSeries = ReadColorSpreadsheet2(this.bcr, length);
			} else if (c_oSer_Sparkline.ColorNegative === type) {
				oSparklineGroup.colorNegative = ReadColorSpreadsheet2(this.bcr, length);
			} else if (c_oSer_Sparkline.ColorAxis === type) {
				oSparklineGroup.colorAxis = ReadColorSpreadsheet2(this.bcr, length);
			} else if (c_oSer_Sparkline.ColorMarkers === type) {
				oSparklineGroup.colorMarkers = ReadColorSpreadsheet2(this.bcr, length);
			} else if (c_oSer_Sparkline.ColorFirst === type) {
				oSparklineGroup.colorFirst = ReadColorSpreadsheet2(this.bcr, length);
			} else if (c_oSer_Sparkline.ColorLast === type) {
				oSparklineGroup.colorLast = ReadColorSpreadsheet2(this.bcr, length);
			} else if (c_oSer_Sparkline.ColorHigh === type) {
				oSparklineGroup.colorHigh = ReadColorSpreadsheet2(this.bcr, length);
			} else if (c_oSer_Sparkline.ColorLow === type) {
				oSparklineGroup.colorLow = ReadColorSpreadsheet2(this.bcr, length);
			} else if (c_oSer_Sparkline.Ref === type) {
				oSparklineGroup.f = this.stream.GetString2LE(length);
			} else if (c_oSer_Sparkline.Sparklines === type) {
				res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadSparklines(t, l, oSparklineGroup);
                });
			} else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadSparklines = function (type, length, oSparklineGroup) {
            var oThis = this;
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_Sparkline.Sparkline === type) {
				var newSparkline = new AscCommonExcel.sparkline();
				res = this.bcr.Read1(length, function (t, l) {
                    return oThis.ReadSparkline(t, l, newSparkline);
                });
				oSparklineGroup.arrSparklines.push(newSparkline);
			} else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
		this.ReadSparkline = function (type, length, oSparkline) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSer_Sparkline.SparklineRef === type) {
				oSparkline.setF(this.stream.GetString2LE(length));
			} else if (c_oSer_Sparkline.SparklineSqRef === type) {
				oSparkline.setSqRef(this.stream.GetString2LE(length));
			} else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadReplies = function(type, length, oCommentData)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_CommentData.Reply === type )
            {
                var oReplyData = new Asc.asc_CCommentData();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadCommentData(t,l,oReplyData);
                });
                oCommentData.asc_addReply(oReplyData);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadThreadedComment = function(type, length, oCommentData)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_ThreadedComment.dT === type ) {
                oCommentData.asc_putTime("");
                var dateMs =  AscCommon.getTimeISO8601(this.stream.GetString2LE(length));
                if(!isNaN(dateMs))
                    oCommentData.asc_putOnlyOfficeTime(dateMs + "");
            } else if ( c_oSer_ThreadedComment.personId === type ) {
                var person = this.personList[this.stream.GetString2LE(length)];
                if (person) {
                    oCommentData.asc_putUserName(person.displayName);
                    oCommentData.asc_putUserId(person.userId);
                    oCommentData.asc_putProviderId(person.providerId);
                }
            } else if ( c_oSer_ThreadedComment.id === type ) {
                oCommentData.asc_putGuid(this.stream.GetString2LE(length));
            } else if ( c_oSer_ThreadedComment.done === type ) {
                oCommentData.asc_putSolved(this.stream.GetBool());
            } else if ( c_oSer_ThreadedComment.text === type ) {
                oCommentData.asc_putText(this.stream.GetString2LE(length));
            // } else if ( c_oSer_ThreadedComment.mention === type ) {
            //     var mention = {mentionpersonId: undefined, mentionId: undefined, startIndex: undefined, length: undefined};
            //     res = this.bcr.Read1(length, function(t,l){
            //         return oThis.ReadThreadedCommentMention(t,l,mention);
            //     });
            //     oCommentData.aMentions.push(mention);
            } else if ( c_oSer_ThreadedComment.reply === type ) {
                var reply = new Asc.asc_CCommentData();
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadThreadedComment(t,l,reply);
                });
                oCommentData.asc_addReply(reply);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadThreadedCommentMention = function(type, length, mention)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_ThreadedComment.mentionpersonId === type ) {
                mention.mentionpersonId = this.stream.GetString2LE(length);
            } else if ( c_oSer_ThreadedComment.mentionId === type ) {
                mention.mentionId = this.stream.GetString2LE(length);
            } else if ( c_oSer_ThreadedComment.startIndex === type ) {
                mention.startIndex = this.stream.GetULong();
            } else if ( c_oSer_ThreadedComment.length === type ) {
                mention.length = this.stream.GetULong();
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
    }
    /** @constructor */
    function Binary_CalcChainTableReader(stream, aCalcChain)
    {
        this.stream = stream;
        this.aCalcChain = aCalcChain;
        this.bcr = new Binary_CommonReader(this.stream);
        this.Read = function()
        {
            var oThis = this;
            return this.bcr.ReadTable(function(t, l){
                return oThis.ReadCalcChainContent(t,l);
            });
        };
        this.ReadCalcChainContent = function(type, length)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_CalcChainType.CalcChainItem === type )
            {
                var oNewCalcChain = {};
                res = this.bcr.Read2Spreadsheet(length, function(t,l){
                    return oThis.ReadCalcChain(t,l, oNewCalcChain);
                });
                this.aCalcChain.push(oNewCalcChain);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadCalcChain = function(type, length, oCalcChain)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_CalcChainType.Array == type )
                oCalcChain.Array = this.stream.GetBool();
            else if ( c_oSer_CalcChainType.SheetId == type )
                oCalcChain.SheetId = this.stream.GetULongLE();
            else if ( c_oSer_CalcChainType.DependencyLevel == type )
                oCalcChain.DependencyLevel = this.stream.GetBool();
            else if ( c_oSer_CalcChainType.Ref == type )
                oCalcChain.Ref = this.stream.GetString2LE(length);
            else if ( c_oSer_CalcChainType.ChildChain == type )
                oCalcChain.ChildChain = this.stream.GetBool();
            else if ( c_oSer_CalcChainType.NewThread == type )
                oCalcChain.NewThread = this.stream.GetBool();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
    }
    /** @constructor */
    function Binary_OtherTableReader(stream, oMedia, wb)
    {
        this.stream = stream;
        this.oMedia = oMedia;
        this.wb = wb;
        this.bcr = new Binary_CommonReader(this.stream);
        this.Read = function()
        {
            var oThis = this;
            var oRes = this.bcr.ReadTable(function(t, l){
                return oThis.ReadOtherContent(t,l);
            });
            return oRes;
        };
        this.ReadOtherContent = function(type, length)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_OtherType.Media === type )
            {
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadMediaContent(t,l);
                });
            }
            else if ( c_oSer_OtherType.Theme === type )
            {
                this.wb.theme = pptx_content_loader.ReadTheme(this, this.stream);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadMediaContent = function(type, length)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_OtherType.MediaItem === type )
            {
                var oNewMedia = {};
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadMediaItem(t,l, oNewMedia);
                });
                if(null != oNewMedia.id && null != oNewMedia.src)
                    this.oMedia[oNewMedia.id] = oNewMedia.src;
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadMediaItem = function(type, length, oNewMedia)
        {
            var res = c_oSerConstants.ReadOk;
            if ( c_oSer_OtherType.MediaSrc === type )
            {
                var src = this.stream.GetString2LE(length);
                if(0 != src.indexOf("http:") && 0 != src.indexOf("data:") && 0 != src.indexOf("https:") && 0 != src.indexOf("ftp:") && 0 != src.indexOf("file:"))
                    oNewMedia.src = AscCommon.g_oDocumentUrls.getImageUrl(src);
                else
                    oNewMedia.src = src;
            }
            else if ( c_oSer_OtherType.MediaId === type )
                oNewMedia.id = this.stream.GetULongLE();
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
    }

    //TODO копия кода из serialize2
    function Binary_CustomsTableReader(stream, CustomXmls) {
        //this.Document = doc;
        //this.oReadResult = oReadResult;
        this.stream = stream;
        this.CustomXmls = CustomXmls;
        this.bcr = new Binary_CommonReader(this.stream);
        this.Read = function() {
            var oThis = this;
            return this.bcr.ReadTable(function(t, l) {
                return oThis.ReadCustom(t, l);
            });
        };
        this.ReadCustom = function(type, length) {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if (c_oSerCustoms.Custom === type) {
                var custom = {Uri: [], ItemId: null, Content: null};
                res = this.bcr.Read1(length, function(t, l) {
                    return oThis.ReadCustomContent(t, l, custom);
                });
                this.CustomXmls.push(custom);
            }
            else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadCustomContent = function(type, length, custom) {
            var res = c_oSerConstants.ReadOk;
            if (c_oSerCustoms.Uri === type) {
                custom.Uri.push(this.stream.GetString2LE(length));
            } else if (c_oSerCustoms.ItemId === type) {
                custom.ItemId = this.stream.GetString2LE(length);
            } else if (c_oSerCustoms.Content === type) {
                //поскольку данную строку не использую, храню в виде массива. если нужно будет строка, то не забыть про BOM
                custom.Content = this.stream.GetBuffer(length);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
    }

    function getBinaryOtherTableGVar(wb)
    {
        AscCommonExcel.g_oColorManager.setTheme(wb.theme);

        var sMinorFont = null;
        if(null != wb.theme.themeElements && null != wb.theme.themeElements.fontScheme && null != wb.theme.themeElements.fontScheme.minorFont)
            sMinorFont = wb.theme.themeElements.fontScheme.minorFont.latin;
        var sDefFont = "Calibri";
        if(null != sMinorFont && "" != sMinorFont)
            sDefFont = sMinorFont;
        g_oDefaultFormat.Font = new AscCommonExcel.Font();
		g_oDefaultFormat.Font.assignFromObject({
		    fn: sDefFont,
            scheme: EFontScheme.fontschemeMinor,
			fs: 11,
			c: AscCommonExcel.g_oColorManager.getThemeColor(AscCommonExcel.g_nColorTextDefault)
		});
        g_oDefaultFormat.Fill = new AscCommonExcel.Fill();
        g_oDefaultFormat.Border = new AscCommonExcel.Border();
        g_oDefaultFormat.Border.initDefault();
        g_oDefaultFormat.Num = g_oDefaultFormat.NumAbs = new AscCommonExcel.Num({f : "General"});
        g_oDefaultFormat.Align = g_oDefaultFormat.AlignAbs = new AscCommonExcel.Align({
            hor : null,
            indent : 0,
            RelativeIndent : 0,
            shrink : false,
            angle : 0,
            ver : Asc.c_oAscVAlign.Bottom,
            wrap : false
        });
    }

    function BinaryPersonReader(stream, personList)
    {
        this.stream = stream;
        this.personList = personList;
        this.bcr = new Binary_CommonReader(this.stream);
        this.Read = function()
        {
            var oThis = this;
            var oRes = this.bcr.ReadTable(function(t, l){
                return oThis.ReadPersonList(t,l);
            });
            return oRes;
        };
        this.ReadPersonList = function(type, length, persons)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_Person.person === type ) {
                var person = {providerId:"", userId:"", displayName:""};
                res = this.bcr.Read1(length, function(t,l){
                    return oThis.ReadPerson(t,l, person);
                });
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        this.ReadPerson = function(type, length, person)
        {
            var res = c_oSerConstants.ReadOk;
            var oThis = this;
            if ( c_oSer_Person.id === type ) {
                this.personList[this.stream.GetString2LE(length)] = person;
            } else if (c_oSer_Person.providerId === type) {
                person.providerId = this.stream.GetString2LE(length);
            } else if (c_oSer_Person.userId === type) {
                person.userId = this.stream.GetString2LE(length);
            } else if (c_oSer_Person.displayName === type) {
                person.displayName = this.stream.GetString2LE(length);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
    }

    /** @constructor */
    function BinaryFileReader(isCopyPaste)
    {
        this.stream = null;
        this.InitOpenManager = new InitOpenManager(isCopyPaste);

        this.getbase64DecodedData = function(szSrc)
        {
			var isBase64 = typeof szSrc === 'string';
			var srcLen = szSrc.length;
			var nWritten = 0;

            var nType = 0;
            var index = AscCommon.c_oSerFormat.Signature.length;
            var version = "";
            var dst_len = "";
            while (true)
            {
                index++;
                var _c = isBase64 ? szSrc.charCodeAt(index) : szSrc[index];
                if (_c == ";".charCodeAt(0))
                {

                    if(0 == nType)
                    {
                        nType = 1;
                        continue;
                    }
                    else
                    {
                        index++;
                        break;
                    }
                }
                if(0 == nType)
                    version += String.fromCharCode(_c);
                else
                    dst_len += String.fromCharCode(_c);
            }
			var nVersion = 0;
			if(version.length > 1)
			{
				var nTempVersion = version.substring(1) - 0;
				if(nTempVersion)
					AscCommon.CurFileVersion = nVersion = nTempVersion;
			}
			var stream;
			if (Asc.c_nVersionNoBase64 !== nVersion) {
				var dstLen = parseInt(dst_len);

				var memoryData = AscCommon.Base64.decode(szSrc, false, dstLen, index);
				stream = new AscCommon.FT_Stream2(memoryData, memoryData.length);

			} else {
				stream = new AscCommon.FT_Stream2(szSrc, szSrc.length);
				//skip header
				stream.EnterFrame(index);
				stream.Seek(index);
			}

            return stream;
        };
        this.getbase64DecodedData2 = function(szSrc, szSrcOffset, stream, streamOffset)
        {
            var srcLen = szSrc.length;
            var nWritten = streamOffset;
            var dstPx = stream.data;
            var index = szSrcOffset;

            return AscCommon.Base64.decodeData(szSrc, index, srcLen, dstPx, nWritten);
        };
        this.Read = function(data, wb)
        {
            var t = this;
            pptx_content_loader.Clear();
            this.InitOpenManager.wb = wb;
			var pasteBinaryFromExcel = false;
			if(this.InitOpenManager.copyPasteObj && this.InitOpenManager.copyPasteObj.isCopyPaste && typeof editor != "undefined" && editor)
				pasteBinaryFromExcel = true;

			this.stream = this.getbase64DecodedData(data);
			if(!pasteBinaryFromExcel)
				History.TurnOff();

			AscCommonExcel.executeInR1C1Mode(false, function () {
				t.ReadMainTable(wb);
			});

            if(!this.InitOpenManager.copyPasteObj.isCopyPaste)
            {
                this.InitOpenManager.readDefStyles(wb, wb.CellStyles.DefaultStyles);
                // ReadDefTableStyles(wb, wb.TableStyles.DefaultStyles);
                // wb.TableStyles.concatStyles();
            }
            if(pptx_content_loader.Reader)
            {
                pptx_content_loader.Reader.AssignConnectedObjects();
            }
			if(!pasteBinaryFromExcel)
				History.TurnOn();
            //чтобы удалялся stream с бинарником
			pptx_content_loader.Clear(true);
        };
        this.ReadMainTable = function(wb)
        {
            var res = c_oSerConstants.ReadOk;
            //mtLen
            res = this.stream.EnterFrame(1);
            if(c_oSerConstants.ReadOk != res)
                return res;
            var mtLen = this.stream.GetUChar();
            var aSeekTable = [];
            var nOtherTableOffset = null;
            var nSharedStringTableOffset = null;
            var nStyleTableOffset = null;
            var nWorkbookTableOffset = null;
            var nPersonListTableOffset = null;
            var fileStream;
            for(var i = 0; i < mtLen; ++i)
            {
                //mtItem
                res = this.stream.EnterFrame(5);
                if(c_oSerConstants.ReadOk != res)
                    return res;
                var mtiType = this.stream.GetUChar();
                var mtiOffBits = this.stream.GetULongLE();
                if(c_oSerTableTypes.Other == mtiType)
                    nOtherTableOffset = mtiOffBits;
                else if(c_oSerTableTypes.SharedStrings == mtiType)
                    nSharedStringTableOffset = mtiOffBits;
                else if(c_oSerTableTypes.Styles == mtiType)
                    nStyleTableOffset = mtiOffBits;
                else if(c_oSerTableTypes.Workbook == mtiType)
                    nWorkbookTableOffset = mtiOffBits;
                else if(c_oSerTableTypes.PersonList == mtiType)
                    nPersonListTableOffset = mtiOffBits;
                else
                    aSeekTable.push( {type: mtiType, offset: mtiOffBits} );
            }




            var aSharedStrings = [];
            var aCellXfs = [];
            var oMediaArray = {};


            //****TODO Не нахожу в файле****
            wb.aWorksheets = [];
            if(null != nOtherTableOffset)
            {
                res = this.stream.Seek(nOtherTableOffset);
                if(c_oSerConstants.ReadOk == res)
                    res = (new Binary_OtherTableReader(this.stream, oMediaArray, wb)).Read();
            }

            this.InitOpenManager.initSchemeAndTheme(wb);
            //****TODO Не нахожу в файле****




            if(null != nSharedStringTableOffset)
            {
                res = this.stream.Seek(nSharedStringTableOffset);
                if(c_oSerConstants.ReadOk == res)
                    res = (new Binary_SharedStringTableReader(this.stream, wb, aSharedStrings)).Read();
            }

            //aCellXfs - внутри уже не нужна, поскольку вынес функцию InitStyleManager в InitOpenManager
			this.InitOpenManager.oReadResult.stylesTableReader = new Binary_StylesTableReader(this.stream, wb/*, aCellXfs, this.InitOpenManager.copyPasteObj.isCopyPaste*/)
            if(null != nStyleTableOffset)
            {
                res = this.stream.Seek(nStyleTableOffset);
                if (c_oSerConstants.ReadOk == res) {
                    var oStyleObject = this.InitOpenManager.oReadResult.stylesTableReader.Read();
                    this.InitOpenManager.InitStyleManager(oStyleObject, aCellXfs);
                    this.InitOpenManager.Dxfs = oStyleObject.aDxfs;
                    wb.oNumFmtsOpen = oStyleObject.oNumFmts;
                    wb.dxfsOpen = oStyleObject.aDxfs;
                }
            }


            var personList = {};
            if(null != nPersonListTableOffset)
            {
                res = this.stream.Seek(nPersonListTableOffset);
                if(c_oSerConstants.ReadOk == res)
                    res = new BinaryPersonReader(this.stream, personList).Read();
            }







			var bwtr = new Binary_WorksheetTableReader(this.stream, this.InitOpenManager, wb, aSharedStrings, aCellXfs, oMediaArray, personList);
			if(null != nWorkbookTableOffset)
			{
				res = this.stream.Seek(nWorkbookTableOffset);
				if(c_oSerConstants.ReadOk == res)
					res = (new Binary_WorkbookTableReader(this.stream, this.InitOpenManager, wb, bwtr)).Read();
			}




            if(c_oSerConstants.ReadOk == res)
            {
                for(var i = 0; i < aSeekTable.length; ++i)
                {
                    var seek = aSeekTable[i];
                    var mtiType = seek.type;
                    var mtiOffBits = seek.offset;
                    res = this.stream.Seek(mtiOffBits);
                    if(c_oSerConstants.ReadOk != res)
                        break;
                    switch(mtiType)
                    {
                        // case c_oSerTableTypes.SharedStrings:
                        // res = (new Binary_SharedStringTableReader(this.stream, aSharedStrings)).Read();
                        // break;
                        // case c_oSerTableTypes.Styles:
                        // res = (new Binary_StylesTableReader(this.stream, wb.oStyleManager, aCellXfs)).Read();
                        // break;
                        // case c_oSerTableTypes.Workbook:
                        // res = (new Binary_WorkbookTableReader(this.stream, wb)).Read();
                        // break;
                        case c_oSerTableTypes.Worksheets:
                            res = bwtr.Read();
                            break;
                        // case c_oSerTableTypes.CalcChain:
                        //     res = (new Binary_CalcChainTableReader(this.stream, wb.calcChain)).Read();
                        //     break;
                        // case c_oSerTableTypes.Other:
                        // res = (new Binary_OtherTableReader(this.stream, oMediaArray)).Read();
                        // break;
                        case c_oSerTableTypes.App:
                            this.stream.Seek2(mtiOffBits);
                            fileStream = this.stream.ToFileStream();
                            wb.App = new AscCommon.CApp();
                            wb.App.fromStream(fileStream);
                            this.stream.FromFileStream(fileStream);
                            break;
                        case c_oSerTableTypes.Core:
                            this.stream.Seek2(mtiOffBits);
                            fileStream = this.stream.ToFileStream();
                            wb.Core = new AscCommon.CCore();
                            wb.Core.fromStream(fileStream);
                            this.stream.FromFileStream(fileStream);
                            break;
                        case c_oSerTableTypes.CustomProperties:
                            this.stream.Seek2(mtiOffBits);
                            fileStream = this.stream.ToFileStream();
                            wb.CustomProperties.fromStream(fileStream);
                            this.stream.FromFileStream(fileStream);
                            break;
                        case c_oSerTableTypes.Customs:
                            this.stream.Seek2(mtiOffBits);
                            wb.customXmls = [];
                            res = (new Binary_CustomsTableReader(this.stream, wb.customXmls)).Read();
                            break;
                    }
                    if(c_oSerConstants.ReadOk != res)
                        break;
                }
            }
			this.InitOpenManager.PostLoadPrepareDefNames(wb);
            //todo инициализация формул из-за именованных диапазонов перенесена в wb.init ее надо вызывать в любом случае(Rev: 61959)
            if(!this.InitOpenManager.copyPasteObj.isCopyPaste || this.InitOpenManager.copyPasteObj.selectAllSheet)
            {
				bwtr.ReadSheetDataExternal(false);
				if (!this.InitOpenManager.copyPasteObj.isCopyPaste) {
					this.InitOpenManager.PostLoadPrepare(wb);
				}
                wb.init(this.InitOpenManager.oReadResult, false, true);
            } else {
				bwtr.ReadSheetDataExternal(true);
				if(Asc["editor"] && Asc["editor"].wb) {
					wb.init(this.InitOpenManager.oReadResult, true);
				}
            }
            return res;
        };
		this.PostLoadPrepare = function(wb)
		{
			if (wb.workbookPr && null != wb.workbookPr.Date1904) {
				AscCommon.bDate1904 = wb.workbookPr.Date1904;
				AscCommonExcel.c_DateCorrectConst = AscCommon.bDate1904?AscCommonExcel.c_Date1904Const:AscCommonExcel.c_Date1900Const;
			}
			if (this.InitOpenManager.oReadResult.macros) {
                let parsedMacros = AscCommonExcel.safeJsonParse(this.InitOpenManager.oReadResult.macro);
                if (parsedMacros) {
                    if (parsedMacros["customFunctions"]) {
                        wb.fileCustomFunctions = parsedMacros["customFunctions"];
                    }
                    delete parsedMacros["customFunctions"];
                    wb.oApi.macros.SetData(JSON.stringify(parsedMacros));
                } else {
                    wb.oApi.macros.SetData(this.InitOpenManager.oReadResult.macros);
                }
            }
			if (this.InitOpenManager.oReadResult.vbaMacros) {
				wb.oApi.vbaMacros = this.InitOpenManager.oReadResult.vbaMacros;
			}
            if (this.InitOpenManager.oReadResult.vbaProject) {
                wb.oApi.vbaProject = this.InitOpenManager.oReadResult.vbaProject;
            }
            wb.checkCorrectTables();
		};
		this.PostLoadPrepareDefNames = function(wb)
		{
			this.InitOpenManager.oReadResult.defNames.forEach(function (defName) {
				if (null != defName.Name && null != defName.Ref) {
					var _type = Asc.c_oAscDefNameType.none;
					if (wb.getSlicerCacheByName(defName.Name)) {
						_type = Asc.c_oAscDefNameType.slicer;
					}
					wb.dependencyFormulas.addDefNameOpen(defName.Name, defName.Ref, defName.LocalSheetId, defName.Hidden, _type);
				}
			});
		}
	}
    function CSlicerStyles()
    {
        this.DefaultStyle = "SlicerStyleLight1";
        this.CustomStyles = {};
        this.DefaultStyles = {};
        this.AllStyles = {};
    }
    CSlicerStyles.prototype =
    {
        concatStyles : function()
        {
            for(var i in this.DefaultStyles)
                this.AllStyles[i] = this.DefaultStyles[i];
            for(var i in this.CustomStyles)
                this.AllStyles[i] = this.CustomStyles[i];
        },
        addCustomStylesAtOpening : function(styles, Dxfs)
        {
            if (!styles) {
                return;
            }
            this.DefaultStyle = styles.defaultSlicerStyle || this.DefaultStyle;
            for(var i = 0; i < styles.slicerStyle.length; ++i){
                this.addStyleAtOpening(this.CustomStyles, Dxfs, styles.slicerStyle[i]);
            }
        },
        addDefaultStylesAtOpening : function(styles, Dxfs)
        {
            for(var i = 0; i < styles.slicerStyle.length; ++i){
                this.addStyleAtOpening(this.DefaultStyles, Dxfs, styles.slicerStyle[i]);
            }
        },
        addStyleAtOpening : function(styles, Dxfs, style)
        {
            if (null === style.name) {
                return;
            }
            var elems = {};
            for (var i = 0; i < style.slicerStyleElements.length; ++i) {
                var elem = style.slicerStyleElements[i];
                if (null !== elem.type && null !== elem.dxfId) {
                    elems[elem.type] = Dxfs[elem.dxfId];
                }
            }
            styles[style.name] = elems;
        }
    };
    function CTableStyles()
    {
        this.DefaultTableStyle = "TableStyleMedium2";
        this.DefaultPivotStyle = "PivotStyleLight16";
        this.CustomStyles = {};
        this.DefaultStyles = {};
		this.DefaultStylesPivot = {};
        this.AllStyles = {};
    }
    CTableStyles.prototype =
    {
        concatStyles : function()
        {
            for(var i in this.DefaultStyles)
                this.AllStyles[i] = this.DefaultStyles[i];
			for(var i in this.DefaultStylesPivot)
				this.AllStyles[i] = this.DefaultStylesPivot[i];
            for(var i in this.CustomStyles)
                this.AllStyles[i] = this.CustomStyles[i];
        },
		readAttributes: function(attr, uq) {
			if (attr()) {
				var vals = attr();
				var val;
				val = vals["defaultTableStyle"];
				if (undefined !== val) {
					this.DefaultTableStyle = AscCommon.unleakString(uq(val));
				}
				val = vals["defaultPivotStyle"];
				if (undefined !== val) {
					this.DefaultPivotStyle = AscCommon.unleakString(uq(val));
				}
			}
		},
		onStartNode: function(elem, attr, uq) {
			var newContext = this;
			if ("tableStyle" === elem) {
				newContext = new CTableStyle();
				if (newContext.readAttributes) {
					newContext.readAttributes(attr, uq);
				}
				this.CustomStyles[newContext.name] = newContext;
				AscCommon.openXml.SaxParserDataTransfer.curTableStyle = newContext;
			} else {
				newContext = null;
			}
			return newContext;
		}
    };

	function CTableStyleStripe(size, offset, opt_row){
		this.size = size;
		this.offset = offset;
		this.row = opt_row;
	}
    function CTableStyle()
    {
        this.name = null;
        this.pivot = true;
        this.table = true;
        this.displayName = null; // Показываемое имя (для дефалтовых оно будет с пробелами, а для пользовательских совпадает с name)

        this.blankRow = null;
        this.firstColumn = null;
        this.firstColumnStripe = null;
        this.firstColumnSubheading = null;
        this.firstHeaderCell = null;
        this.firstRowStripe = null;
        this.firstRowSubheading = null;
        this.firstSubtotalColumn = null;
        this.firstSubtotalRow = null;
        this.firstTotalCell = null;
        this.headerRow = null;
        this.lastColumn = null;
        this.lastHeaderCell = null;
        this.lastTotalCell = null;
        this.pageFieldLabels = null;
        this.pageFieldValues = null;
        this.secondColumnStripe = null;
        this.secondColumnSubheading = null;
        this.secondRowStripe = null;
        this.secondRowSubheading = null;
        this.secondSubtotalColumn = null;
        this.secondSubtotalRow = null;
        this.thirdColumnSubheading = null;
        this.thirdRowSubheading = null;
        this.thirdSubtotalColumn = null;
        this.thirdSubtotalRow = null;
        this.totalRow = null;
        this.wholeTable = null;
    }
	CTableStyle.prototype.initStyle = function (sheetMergedStyles, bbox, options, headerRowCount, totalsRowCount) {
		var r1Data = bbox.r1 + headerRowCount;
		var r2Data = bbox.r2 - totalsRowCount;
		var bboxTmp;
		var offsetStripe;
		var stripe;
		if (this.wholeTable) {
			sheetMergedStyles.setTablePivotStyle(bbox, this.wholeTable.dxf);
		}
		if (r1Data <= r2Data) {
			if (options.ShowColumnStripes) {
				if (this.firstColumnStripe) {
					offsetStripe = this.secondColumnStripe ? this.secondColumnStripe.size : 1;
					stripe = new CTableStyleStripe(this.firstColumnStripe.size, offsetStripe);
					bboxTmp = new Asc.Range(bbox.c1, r1Data, bbox.c2, r2Data);
					sheetMergedStyles.setTablePivotStyle(bboxTmp, this.firstColumnStripe.dxf, stripe);
				}
				if (this.secondColumnStripe) {
					offsetStripe = this.firstColumnStripe ? this.firstColumnStripe.size : 1;
					stripe = new CTableStyleStripe(this.secondColumnStripe.size, offsetStripe);
					if (bbox.c1 + offsetStripe <= bbox.c2) {
						bboxTmp = new Asc.Range(bbox.c1 + offsetStripe, r1Data, bbox.c2, r2Data);
						sheetMergedStyles.setTablePivotStyle(bboxTmp, this.secondColumnStripe.dxf, stripe);
					}
				}
			}
			if (options.ShowRowStripes) {
				if (this.firstRowStripe) {
					offsetStripe = this.secondRowStripe ? this.secondRowStripe.size : 1;
					stripe = new CTableStyleStripe(this.firstRowStripe.size, offsetStripe, true);
					bboxTmp = new Asc.Range(bbox.c1, r1Data, bbox.c2, r2Data);
					sheetMergedStyles.setTablePivotStyle(bboxTmp, this.firstRowStripe.dxf, stripe);
				}
				if (this.secondRowStripe) {
					offsetStripe = this.firstRowStripe ? this.firstRowStripe.size : 1;
					stripe = new CTableStyleStripe(this.secondRowStripe.size, offsetStripe, true);
					if (r1Data + offsetStripe <= r2Data) {
						bboxTmp = new Asc.Range(bbox.c1, r1Data + offsetStripe, bbox.c2, r2Data);
						sheetMergedStyles.setTablePivotStyle(bboxTmp, this.secondRowStripe.dxf, stripe);
					}
				}
			}
		}
		if (options.ShowLastColumn && this.lastColumn) {
			bboxTmp = new Asc.Range(bbox.c2, bbox.r1, bbox.c2, bbox.r2);
			sheetMergedStyles.setTablePivotStyle(bboxTmp, this.lastColumn.dxf);
		}
		if (options.ShowFirstColumn && this.firstColumn) {
			bboxTmp = new Asc.Range(bbox.c1, bbox.r1, bbox.c1, bbox.r2);
			sheetMergedStyles.setTablePivotStyle(bboxTmp, this.firstColumn.dxf);
		}
		if (this.headerRow && headerRowCount > 0) {
			bboxTmp = new Asc.Range(bbox.c1, bbox.r1, bbox.c2, bbox.r1);
			sheetMergedStyles.setTablePivotStyle(bboxTmp, this.headerRow.dxf);
		}
		if (this.totalRow && totalsRowCount > 0) {
			bboxTmp = new Asc.Range(bbox.c1, bbox.r2, bbox.c2, bbox.r2);
			sheetMergedStyles.setTablePivotStyle(bboxTmp, this.totalRow.dxf);
		}
		if (this.firstHeaderCell && headerRowCount > 0) {
			bboxTmp = new Asc.Range(bbox.c1, bbox.r1, bbox.c1, bbox.r1);
			sheetMergedStyles.setTablePivotStyle(bboxTmp, this.firstHeaderCell.dxf);
		}
		if (this.lastHeaderCell && headerRowCount > 0) {
			bboxTmp = new Asc.Range(bbox.c2, bbox.r1, bbox.c2, bbox.r1);
			sheetMergedStyles.setTablePivotStyle(bboxTmp, this.lastHeaderCell.dxf);
		}
		if (this.firstTotalCell && totalsRowCount > 0) {
			bboxTmp = new Asc.Range(bbox.c1, bbox.r2, bbox.c1, bbox.r2);
			sheetMergedStyles.setTablePivotStyle(bboxTmp, this.firstTotalCell.dxf);
		}
		if (this.lastTotalCell && totalsRowCount > 0) {
			bboxTmp = new Asc.Range(bbox.c2, bbox.r2, bbox.c2, bbox.r2);
			sheetMergedStyles.setTablePivotStyle(bboxTmp, this.lastTotalCell.dxf);
		}
	};
	CTableStyle.prototype.readAttributes = function (attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["name"];
			if (undefined !== val) {
				this.name = AscCommon.unleakString(uq(val));
				this.displayName = this.name;
			}
			val = vals["displayName"];
			if (undefined !== val) {
				this.displayName = AscCommon.unleakString(uq(val));
			}
			val = vals["pivot"];
			if (undefined !== val) {
				this.pivot = AscCommon.getBoolFromXml(val);
			}
			val = vals["table"];
			if (undefined !== val) {
				this.table = AscCommon.getBoolFromXml(val);
			}
		}
	};
	CTableStyle.prototype.onStartNode = function (elem, attr, uq) {
		var newContext = this;
		if ("tableStyleElement" === elem) {
			newContext = new CTableStyleElement();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
		} else {
			newContext = null;
		}
		return newContext;
	};

    CTableStyle.prototype.getTableStyleElement = function () {
        if (null != this.blankRow) {
            return this.blankRow;
        }
        if (null != this.firstColumn) {
            return this.firstColumn;
        }
        if (null != this.firstColumnStripe) {
            return this.firstColumnStripe;
        }
        if (null != this.firstColumnSubheading) {
            return this.firstColumnSubheading;
        }
        if (null != this.firstHeaderCell) {
            return this.firstHeaderCell;
        }
        if (null != this.firstRowStripe) {
            return this.firstRowStripe;
        }
        if (null != this.firstRowSubheading) {
            return this.firstRowSubheading;
        }
        if (null != this.firstSubtotalColumn) {
            return this.firstSubtotalColumn;
        }
        if (null != this.firstSubtotalRow) {
            return this.firstSubtotalRow;
        }
        if (null != this.firstTotalCell) {
            return this.firstTotalCell;
        }
        if (null != this.headerRow) {
            return this.headerRow;
        }
        if (null != this.lastColumn) {
            return this.lastColumn;
        }
        if (null != this.lastHeaderCell) {
            return this.lastHeaderCell;
        }
        if (null != this.lastTotalCell) {
            return this.lastTotalCell;
        }
        if (null != this.pageFieldLabels) {
            return this.pageFieldLabels;
        }
        if (null != this.pageFieldValues) {
            return this.pageFieldValues;
        }
        if (null != this.secondColumnStripe) {
            return this.secondColumnStripe;
        }
        if (null != this.secondColumnSubheading) {
            return this.secondColumnSubheading;
        }
        if (null != this.secondRowStripe) {
            return this.secondRowStripe;
        }
        if (null != this.secondRowSubheading) {
            return this.secondRowSubheading;
        }
        if (null != this.secondSubtotalColumn) {
            return this.secondSubtotalColumn;
        }
        if (null != this.secondSubtotalRow) {
            return this.secondSubtotalRow;
        }
        if (null != this.thirdColumnSubheading) {
            return this.thirdColumnSubheading;
        }
        if (null != this.thirdRowSubheading) {
            return this.thirdRowSubheading;
        }
        if (null != this.thirdSubtotalColumn) {
            return this.thirdSubtotalColumn;
        }
        if (null != this.thirdSubtotalRow) {
            return this.thirdSubtotalRow;
        }
        if (null != this.totalRow) {
            return this.totalRow;
        }
        if (null != this.wholeTable) {
            return this.wholeTable;
        }
    }

    function CTableStyleElement()
    {
        this.size = 1;
        this.dxf = null;
    }
	CTableStyleElement.prototype.readAttributes = function(attr, uq) {
		if(attr()){
			var vals = attr();
			var val;
			val = vals["type"];
			if(undefined !== val){
				var tableStyle = AscCommon.openXml.SaxParserDataTransfer.curTableStyle;
				if("wholeTable"===val)
					tableStyle.wholeTable = this;
				else if("headerRow"===val)
					tableStyle.headerRow = this;
				else if("totalRow"===val)
					tableStyle.totalRow = this;
				else if("firstColumn"===val)
					tableStyle.firstColumn = this;
				else if("lastColumn"===val)
					tableStyle.lastColumn = this;
				else if("firstRowStripe"===val)
					tableStyle.firstRowStripe = this;
				else if("secondRowStripe"===val)
					tableStyle.secondRowStripe = this;
				else if("firstColumnStripe"===val)
					tableStyle.firstColumnStripe = this;
				else if("secondColumnStripe"===val)
					tableStyle.secondColumnStripe = this;
				else if("firstHeaderCell"===val)
					tableStyle.firstHeaderCell = this;
				else if("lastHeaderCell"===val)
					tableStyle.lastHeaderCell = this;
				else if("firstTotalCell"===val)
					tableStyle.firstTotalCell = this;
				else if("lastTotalCell"===val)
					tableStyle.lastTotalCell = this;
				else if("firstSubtotalColumn"===val)
					tableStyle.firstSubtotalColumn = this;
				else if("secondSubtotalColumn"===val)
					tableStyle.secondSubtotalColumn = this;
				else if("thirdSubtotalColumn"===val)
					tableStyle.thirdSubtotalColumn = this;
				else if("firstSubtotalRow"===val)
					tableStyle.firstSubtotalRow = this;
				else if("secondSubtotalRow"===val)
					tableStyle.secondSubtotalRow = this;
				else if("thirdSubtotalRow"===val)
					tableStyle.thirdSubtotalRow = this;
				else if("blankRow"===val)
					tableStyle.blankRow = this;
				else if("firstColumnSubheading"===val)
					tableStyle.firstColumnSubheading = this;
				else if("secondColumnSubheading"===val)
					tableStyle.secondColumnSubheading = this;
				else if("thirdColumnSubheading"===val)
					tableStyle.thirdColumnSubheading = this;
				else if("firstRowSubheading"===val)
					tableStyle.firstRowSubheading = this;
				else if("secondRowSubheading"===val)
					tableStyle.secondRowSubheading = this;
				else if("thirdRowSubheading"===val)
					tableStyle.thirdRowSubheading = this;
				else if("pageFieldLabels"===val)
					tableStyle.pageFieldLabels = this;
				else if("pageFieldValues"===val)
					tableStyle.pageFieldValues = this;
			}
			val = vals["size"];
			if(undefined !== val){
				this.size = val - 0;
			}
			val = vals["dxfId"];
			if (undefined !== val) {
				this.dxf = AscCommon.openXml.SaxParserDataTransfer.dxfs[tableStyle.pivot ? val - 0 : val - 1] || null;
			}
		}
	};
    function ReadDefTableStyles(wb)
    {
    	var stylesZip = "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";

        var dstLen = stylesZip.length;
        var pointer = g_memory.Alloc(dstLen);
        var stream = new AscCommon.FT_Stream2(pointer.data, dstLen);
        stream.obj = pointer.obj;
        var oBinaryFileReader = new AscCommonExcel.BinaryFileReader();
        oBinaryFileReader.getbase64DecodedData2(stylesZip, 0, stream, 0);

        let jsZlib = new AscCommon.ZLib();
        if (jsZlib.open(new Uint8Array(pointer.data))) {
            let contentBytes = jsZlib.getFile("presetTableStyles.xml");
            if (contentBytes) {
                let content = AscCommon.UTF8ArrayToString(contentBytes, 0, contentBytes.length);
                jsZlib.close();
                var stylesXml = new CT_PresetTableStyles(wb.TableStyles.DefaultStyles, wb.TableStyles.DefaultStylesPivot);
                new AscCommon.openXml.SaxParserBase().parse(content, stylesXml);
                wb.TableStyles.concatStyles();
            }
        }
    }
    function ReadDefCellStyles(wb, oOutput)
    {
        var Types = {
            Style		: 0,
            BuiltinId	: 1,
            Hidden		: 2,
            CellStyle	: 3,
            Xfs			: 4,
            Font		: 5,
            Fill		: 6,
            Border		: 7,
            NumFmts		: 8
        };
        // Пишем тип и размер (версию не пишем)
        var sStyles = "XLSY;;11499;5ywAAACHAAAAAQQAAAAAAAAAAyMAAAAABAAAAAAAAAAEDAAAAE4AbwByAG0AYQBsAAUEAAAAAAAAAAQYAAAABgQAAAAABwQAAAAACAQAAAAACQQAAAAABSoAAAABBgMAAAACAQEEBg4AAABDAGEAbABpAGIAcgBpAAkBAQYFAAAAAAAAJkAGAAAAAAcAAAAAAJwAAAABBAAAABwAAAADHAAAAAQOAAAATgBlAHUAdAByAGEAbAAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUtAAAAAQYGAAAAAAQAZZz/BAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhAAAAAACwAAAAEGAAAAAASc6///BwAAAAAAlAAAAAEEAAAAGwAAAAMUAAAABAYAAABCAGEAZAAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUtAAAAAQYGAAAAAAQGAJz/BAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhAAAAAACwAAAAEGAAAAAATOx///BwAAAAAAlgAAAAEEAAAAGgAAAAMWAAAABAgAAABHAG8AbwBkAAUEAAAAAQAAAAQhAAAAAAEAAQEABAEABgQAAAAABwQCAAAACAQBAAAACQQAAAAABS0AAAABBgYAAAAABABhAP8EBg4AAABDAGEAbABpAGIAcgBpAAkBAQYFAAAAAAAAJkAGEAAAAAALAAAAAQYAAAAABM7vxv8HAAAAAADlAAAAAQQAAAAUAAAAAxgAAAAECgAAAEkAbgBwAHUAdAAFBAAAAAEAAAAEHgAAAAABAAQBAAYEAQAAAAcEAgAAAAgEAQAAAAkEAAAAAAUtAAAAAQYGAAAAAAR2Pz//BAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhAAAAAACwAAAAEGAAAAAASZzP//B1AAAAAADwAAAAAGBgAAAAAEf39//wEBDQIPAAAAAAYGAAAAAAR/f3//AQENBA8AAAAABgYAAAAABH9/f/8BAQ0FDwAAAAAGBgAAAAAEf39//wEBDQDqAAAAAQQAAAAVAAAAAxoAAAAEDAAAAE8AdQB0AHAAdQB0AAUEAAAAAQAAAAQeAAAAAAEABAEABgQBAAAABwQCAAAACAQBAAAACQQAAAAABTAAAAAAAQEBBgYAAAAABD8/P/8EBg4AAABDAGEAbABpAGIAcgBpAAkBAQYFAAAAAAAAJkAGEAAAAAALAAAAAQYAAAAABPLy8v8HUAAAAAAPAAAAAAYGAAAAAAQ/Pz//AQENAg8AAAAABgYAAAAABD8/P/8BAQ0EDwAAAAAGBgAAAAAEPz8//wEBDQUPAAAAAAYGAAAAAAQ/Pz//AQENAPQAAAABBAAAABYAAAADJAAAAAQWAAAAQwBhAGwAYwB1AGwAYQB0AGkAbwBuAAUEAAAAAQAAAAQeAAAAAAEABAEABgQBAAAABwQCAAAACAQBAAAACQQAAAAABTAAAAAAAQEBBgYAAAAABAB9+v8EBg4AAABDAGEAbABpAGIAcgBpAAkBAQYFAAAAAAAAJkAGEAAAAAALAAAAAQYAAAAABPLy8v8HUAAAAAAPAAAAAAYGAAAAAAR/f3//AQENAg8AAAAABgYAAAAABH9/f/8BAQ0EDwAAAAAGBgAAAAAEf39//wEBDQUPAAAAAAYGAAAAAAR/f3//AQENAO8AAAABBAAAABcAAAADIgAAAAQUAAAAQwBoAGUAYwBrACAAQwBlAGwAbAAFBAAAAAEAAAAEHgAAAAABAAQBAAYEAQAAAAcEAgAAAAgEAQAAAAkEAAAAAAUtAAAAAAEBAQYDAAAAAgEABAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhAAAAAACwAAAAEGAAAAAASlpaX/B1AAAAAADwAAAAAGBgAAAAAEPz8//wEBBAIPAAAAAAYGAAAAAAQ/Pz//AQEEBA8AAAAABgYAAAAABD8/P/8BAQQFDwAAAAAGBgAAAAAEPz8//wEBBACkAAAAAQQAAAA1AAAAAy4AAAAEIAAAAEUAeABwAGwAYQBuAGEAdABvAHIAeQAgAFQAZQB4AHQABQQAAAABAAAABCQAAAAAAQABAQACAQAEAQAGBAAAAAAHBAAAAAAIBAEAAAAJBAAAAAAFMAAAAAEGBgAAAAAEf39//wMBAQQGDgAAAEMAYQBsAGkAYgByAGkACQEBBgUAAAAAAAAmQAYAAAAABwAAAAAA4wAAAAEEAAAACgAAAAMWAAAABAgAAABOAG8AdABlAAUEAAAAAQAAAAQhAAAAAAEAAwEABAEABgQBAAAABwQCAAAACAQBAAAACQQAAAAABSoAAAABBgMAAAACAQEEBg4AAABDAGEAbABpAGIAcgBpAAkBAQYFAAAAAAAAJkAGEAAAAAALAAAAAQYAAAAABMz///8HUAAAAAAPAAAAAAYGAAAAAASysrL/AQENAg8AAAAABgYAAAAABLKysv8BAQ0EDwAAAAAGBgAAAAAEsrKy/wEBDQUPAAAAAAYGAAAAAASysrL/AQENAKgAAAABBAAAABgAAAADJAAAAAQWAAAATABpAG4AawBlAGQAIABDAGUAbABsAAUEAAAAAQAAAAQhAAAAAAEAAgEABAEABgQBAAAABwQAAAAACAQBAAAACQQAAAAABS0AAAABBgYAAAAABAB9+v8EBg4AAABDAGEAbABpAGIAcgBpAAkBAQYFAAAAAAAAJkAGAAAAAAcUAAAAAA8AAAAABgYAAAAABAGA//8BAQQAmQAAAAEEAAAACwAAAAMmAAAABBgAAABXAGEAcgBuAGkAbgBnACAAVABlAHgAdAAFBAAAAAEAAAAEJAAAAAABAAEBAAIBAAQBAAYEAAAAAAcEAAAAAAgEAQAAAAkEAAAAAAUtAAAAAQYGAAAAAAQAAP//BAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABgAAAAAHAAAAAAChAAAAAQQAAAAQAAAAAyAAAAAEEgAAAEgAZQBhAGQAaQBuAGcAIAAxAAUEAAAAAQAAAAQhAAAAAAEAAgEABAEABgQBAAAABwQAAAAACAQBAAAACQQAAAAABS0AAAAAAQEBBgMAAAACAQMEBg4AAABDAGEAbABpAGIAcgBpAAkBAQYFAAAAAAAALkAGAAAAAAcRAAAAAAwAAAAABgMAAAACAQQBAQwAqwAAAAEEAAAAEQAAAAMgAAAABBIAAABIAGUAYQBkAGkAbgBnACAAMgAFBAAAAAEAAAAEIQAAAAABAAIBAAQBAAYEAQAAAAcEAAAAAAgEAQAAAAkEAAAAAAUtAAAAAAEBAQYDAAAAAgEDBAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACpABgAAAAAHGwAAAAAWAAAAAAYNAAAAAgEEAwUA/3//v//fPwEBDACrAAAAAQQAAAASAAAAAyAAAAAEEgAAAEgAZQBhAGQAaQBuAGcAIAAzAAUEAAAAAQAAAAQhAAAAAAEAAgEABAEABgQBAAAABwQAAAAACAQBAAAACQQAAAAABS0AAAAAAQEBBgMAAAACAQMEBg4AAABDAGEAbABpAGIAcgBpAAkBAQYFAAAAAAAAJkAGAAAAAAcbAAAAABYAAAAABg0AAAACAQQDBc1kZjIzmdk/AQEGAJMAAAABBAAAABMAAAADIAAAAAQSAAAASABlAGEAZABpAG4AZwAgADQABQQAAAABAAAABCQAAAAAAQABAQACAQAEAQAGBAAAAAAHBAAAAAAIBAEAAAAJBAAAAAAFLQAAAAABAQEGAwAAAAIBAwQGDgAAAEMAYQBsAGkAYgByAGkACQEBBgUAAAAAAAAmQAYAAAAABwAAAAAAqgAAAAEEAAAAGQAAAAMYAAAABAoAAABUAG8AdABhAGwABQQAAAABAAAABCEAAAAAAQACAQAEAQAGBAEAAAAHBAAAAAAIBAEAAAAJBAAAAAAFLQAAAAABAQEGAwAAAAIBAQQGDgAAAEMAYQBsAGkAYgByAGkACQEBBgUAAAAAAAAmQAYAAAAAByIAAAAADAAAAAAGAwAAAAIBBAEBBAUMAAAAAAYDAAAAAgEEAQENAIsAAAABBAAAAA8AAAADGAAAAAQKAAAAVABpAHQAbABlAAUEAAAAAQAAAAQkAAAAAAEAAQEAAgEABAEABgQAAAAABwQAAAAACAQBAAAACQQAAAAABS0AAAAAAQEBBgMAAAACAQMEBg4AAABDAGEAbABpAGIAcgBpAAkBAAYFAAAAAAAAMkAGAAAAAAcAAAAAAKwAAAABBAAAAB4AAAADKAAAAAQaAAAAMgAwACUAIAAtACAAQQBjAGMAZQBuAHQAMQAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEBBAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEEAwXNZeYyc5npPwcAAAAAAKwAAAABBAAAACIAAAADKAAAAAQaAAAAMgAwACUAIAAtACAAQQBjAGMAZQBuAHQAMgAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEBBAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEFAwXNZeYyc5npPwcAAAAAAKwAAAABBAAAACYAAAADKAAAAAQaAAAAMgAwACUAIAAtACAAQQBjAGMAZQBuAHQAMwAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEBBAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEGAwXNZeYyc5npPwcAAAAAAKwAAAABBAAAACoAAAADKAAAAAQaAAAAMgAwACUAIAAtACAAQQBjAGMAZQBuAHQANAAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEBBAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEHAwXNZeYyc5npPwcAAAAAAKwAAAABBAAAAC4AAAADKAAAAAQaAAAAMgAwACUAIAAtACAAQQBjAGMAZQBuAHQANQAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEBBAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEIAwXNZeYyc5npPwcAAAAAAKwAAAABBAAAADIAAAADKAAAAAQaAAAAMgAwACUAIAAtACAAQQBjAGMAZQBuAHQANgAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEBBAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEJAwXNZeYyc5npPwcAAAAAAKwAAAABBAAAAB8AAAADKAAAAAQaAAAANAAwACUAIAAtACAAQQBjAGMAZQBuAHQAMQAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEBBAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEEAwWazExmJjPjPwcAAAAAAKwAAAABBAAAACMAAAADKAAAAAQaAAAANAAwACUAIAAtACAAQQBjAGMAZQBuAHQAMgAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEBBAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEFAwWazExmJjPjPwcAAAAAAKwAAAABBAAAACcAAAADKAAAAAQaAAAANAAwACUAIAAtACAAQQBjAGMAZQBuAHQAMwAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEBBAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEGAwWazExmJjPjPwcAAAAAAKwAAAABBAAAACsAAAADKAAAAAQaAAAANAAwACUAIAAtACAAQQBjAGMAZQBuAHQANAAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEBBAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEHAwWazExmJjPjPwcAAAAAAKwAAAABBAAAAC8AAAADKAAAAAQaAAAANAAwACUAIAAtACAAQQBjAGMAZQBuAHQANQAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEBBAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEIAwWazExmJjPjPwcAAAAAAKwAAAABBAAAADMAAAADKAAAAAQaAAAANAAwACUAIAAtACAAQQBjAGMAZQBuAHQANgAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEBBAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEJAwWazExmJjPjPwcAAAAAAKwAAAABBAAAACAAAAADKAAAAAQaAAAANgAwACUAIAAtACAAQQBjAGMAZQBuAHQAMQAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEABAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEEAwXNZGYyM5nZPwcAAAAAAKwAAAABBAAAACQAAAADKAAAAAQaAAAANgAwACUAIAAtACAAQQBjAGMAZQBuAHQAMgAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEABAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEFAwXNZGYyM5nZPwcAAAAAAKwAAAABBAAAACgAAAADKAAAAAQaAAAANgAwACUAIAAtACAAQQBjAGMAZQBuAHQAMwAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEABAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEGAwXNZGYyM5nZPwcAAAAAAKwAAAABBAAAACwAAAADKAAAAAQaAAAANgAwACUAIAAtACAAQQBjAGMAZQBuAHQANAAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEABAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEHAwXNZGYyM5nZPwcAAAAAAKwAAAABBAAAADAAAAADKAAAAAQaAAAANgAwACUAIAAtACAAQQBjAGMAZQBuAHQANQAFBAAAAAEAAAAEIQAAAAABAAEBAAQBAAYEAAAAAAcEAgAAAAgEAQAAAAkEAAAAAAUqAAAAAQYDAAAAAgEABAYOAAAAQwBhAGwAaQBiAHIAaQAJAQEGBQAAAAAAACZABhcAAAAAEgAAAAENAAAAAgEIAwXNZGYy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";

        var oBinaryFileReader = new BinaryFileReader();
        var stream = oBinaryFileReader.getbase64DecodedData(sStyles);
        var bcr = new Binary_CommonReader(stream);
        var oBinary_StylesTableReader = new Binary_StylesTableReader(stream, wb, true/*, [], undefined, true*/);

        var length = stream.GetULongLE();

        var fReadStyle = function(type, length, oCellStyle, oStyleObject) {
            var res = c_oSerConstants.ReadOk;
            if (Types.BuiltinId === type) {
                oCellStyle.BuiltinId = stream.GetULongLE();
            } else if (Types.Hidden === type) {
                oCellStyle.Hidden = stream.GetBool();
            } else if (Types.CellStyle === type) {
                res = bcr.Read1(length, function(t, l) {
                    return oBinary_StylesTableReader.ReadCellStyle(t, l, oCellStyle);
                });
            } else if (Types.Xfs === type) {
				oStyleObject.xfs = new OpenXf();
                res = bcr.Read2Spreadsheet(length, function (t, l) {
                    return oBinary_StylesTableReader.ReadXfs(t, l, oStyleObject.xfs);
                });
            } else if (Types.Font === type) {
                oStyleObject.font = new AscCommonExcel.Font();
                res = bcr.Read2Spreadsheet(length, function (t, l) {
                    return oBinary_StylesTableReader.bssr.ReadRPr(t, l, oStyleObject.font);
                });
				oStyleObject.font.checkSchemeFont(wb.theme);
            } else if (Types.Fill === type) {
                oStyleObject.fill = new AscCommonExcel.Fill();
                res = bcr.Read1(length, function (t, l) {
                    return oBinary_StylesTableReader.ReadFill(t, l, oStyleObject.fill);
                });
            } else if (Types.Border === type) {
                oStyleObject.border = new AscCommonExcel.Border();
                oStyleObject.border.initDefault();
                res = bcr.Read1(length, function (t, l) {
                    return oBinary_StylesTableReader.ReadBorder(t, l, oStyleObject.border);
                });
            } else if (Types.NumFmts === type) {
                res = bcr.Read1(length, function (t, l) {
                    return oBinary_StylesTableReader.ReadNumFmts(t, l, oStyleObject.oNumFmts);
                });
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };
        var fReadStyles = function (type, length, oOutput) {
            var res = c_oSerConstants.ReadOk;
            var oStyleObject = {font: null, fill: null, border: null, oNumFmts: {}, xfs: null};
            if (Types.Style === type) {
                var oCellStyle = new AscCommonExcel.CCellStyle();
                res = bcr.Read1(length, function (t, l) {
                    return fReadStyle(t,l, oCellStyle, oStyleObject);
                });

                var newXf = new AscCommonExcel.CellXfs();
                // Border
                if (null !== oStyleObject.border)
					newXf.border = g_StyleCache.addBorder(oStyleObject.border);
                // Fill
                if (null !== oStyleObject.fill)
					newXf.fill = g_StyleCache.addFill(oStyleObject.fill);
                // Font
                if (null !== oStyleObject.font)
					newXf.font = g_StyleCache.addFont(oStyleObject.font);
                // NumFmt
                if (null !== oStyleObject.xfs.numid) {
                    var oCurNum = oStyleObject.oNumFmts[oStyleObject.xfs.numid];
                    if(null != oCurNum)
						newXf.num = g_StyleCache.addNum(oCurNum);
                    else
						newXf.num = g_StyleCache.addNum( AscCommonExcel.InitOpenManager.prototype.ParseNum.call(this, {id: oStyleObject.xfs.numid, f: null}, oStyleObject.oNumFmts, oBinary_StylesTableReader.useNumId));
                        //g_StyleCache.addNum(oBinary_StylesTableReader.ParseNum({id: oStyleObject.xfs.numid, f: null}, oStyleObject.oNumFmts))
                }
                // QuotePrefix
                if(null != oStyleObject.xfs.QuotePrefix)
					newXf.QuotePrefix = oStyleObject.xfs.QuotePrefix;
				// PivotButton
				if(null != oStyleObject.xfs.PivotButton)
					newXf.PivotButton = oStyleObject.xfs.PivotButton;
				// hidden
				if(null != oStyleObject.xfs.hidden)
					newXf.hidden = oStyleObject.xfs.hidden;
				// locked
				if(null != oStyleObject.xfs.locked)
					newXf.locked = oStyleObject.xfs.locked;
                if(null != oStyleObject.xfs.applyProtection)
                    newXf.applyProtection = oStyleObject.xfs.applyProtection;
                // align
                if(null != oStyleObject.xfs.align)
					newXf.align = g_StyleCache.addAlign(oStyleObject.xfs.align);
                // XfId
                if (null !== oStyleObject.xfs.XfId)
					newXf.XfId = oStyleObject.xfs.XfId;
                // ApplyBorder (ToDo возможно это свойство должно быть в xfs)
                if (null !== oStyleObject.xfs.ApplyBorder)
                    oCellStyle.ApplyBorder = oStyleObject.xfs.ApplyBorder;
                // ApplyFill (ToDo возможно это свойство должно быть в xfs)
                if (null !== oStyleObject.xfs.ApplyFill)
                    oCellStyle.ApplyFill = oStyleObject.xfs.ApplyFill;
                // ApplyFont (ToDo возможно это свойство должно быть в xfs)
                if (null !== oStyleObject.xfs.ApplyFont)
                    oCellStyle.ApplyFont = oStyleObject.xfs.ApplyFont;
                // ApplyNumberFormat (ToDo возможно это свойство должно быть в xfs)
                if (null !== oStyleObject.xfs.ApplyNumberFormat)
                    oCellStyle.ApplyNumberFormat = oStyleObject.xfs.ApplyNumberFormat;
                oCellStyle.xfs = g_StyleCache.addXf(newXf);

                oOutput.push(oCellStyle);
            } else
                res = c_oSerConstants.ReadUnknown;
            return res;
        };

        var res = bcr.Read1(length, function (t, l) {
            return fReadStyles(t, l, oOutput);
        });

        // Если нет стилей в документе, то добавим
        if (0 === wb.CellStyles.CustomStyles.length && 0 < oOutput.length) {
            wb.CellStyles.CustomStyles.push(oOutput[0].clone());
            wb.CellStyles.CustomStyles[0].XfId = 0;
        }
        // Если XfId не задан, то определим его
        if (null == g_oDefaultFormat.XfId) {
            g_oDefaultFormat.XfId = 0;
        }
    }
    function RenameDefSlicerStyle(oStyleObject){
        var tableStyles = oStyleObject.oCustomTableStyles;
        var tableStylesRenamed = {}, i;
        for(i in tableStyles)
        {
            var item = tableStyles[i];
            if(null != item)
            {
                item.style.name = item.style.name.slice(0, -2);
                item.style.displayName = item.style.displayName.slice(0, -2);
                tableStylesRenamed[item.style.name] = item;
            }
        }
        oStyleObject.oCustomTableStyles = tableStylesRenamed;
        if (oStyleObject.oCustomSlicerStyles) {
            var slicerStyles = oStyleObject.oCustomSlicerStyles.slicerStyle;
            for (i = 0; i < slicerStyles.length; ++i) {
                slicerStyles[i].name = slicerStyles[i].name.slice(0, -2);
            }
        }
    }
    function ReadDefSlicerStyles(wb, oOutput)
    {
        var sStyles = "XLSY;;25178;VmIAAAAFAAAAAQAAAAAEIAAAAAULAAAAAAYAAAACAQAAABEFCwAAAAAGAAAAAgEAAAAIBi8AAAAHKgAAAAEGAwAAAAIBAQQGDgAAAEMAYQBsAGkAYgByAGkACQEBBgUAAAAAAAAmQA4dAAAAAxgAAAAGBAAAAAAHBAAAAAAIBAAAAAAJBAAAAAACIwAAAAMeAAAABgQAAAAABwQAAAAACAQAAAAACQQAAAAADAQAAAAADygAAAAQIwAAAAAEAAAAAAAAAAQMAAAATgBvAHIAbQBhAGwABQQAAAAAAAAACrEHAAALKgAAAAEUAAAAAA8AAAAABgYAAAAABL2BT/8BAQ0DDAAAAAABAQEGAwAAAAIBAQtjAAAAAVAAAAAADwAAAAAGBgAAAAAEvYFP/wEBDQIPAAAAAAYGAAAAAAS9gU//AQENBA8AAAAABgYAAAAABL2BT/8BAQ0FDwAAAAAGBgAAAAAEvYFP/wEBDQMJAAAAAQYDAAAAAgEBCzEAAAABGwAAAAAWAAAAAAYNAAAAAgEAAwWzmFnMLGbWvwEBDQMMAAAAAAEBAQYDAAAAAgEBC38AAAABbAAAAAAWAAAAAAYNAAAAAgEAAwUA/3//v//fvwEBDQIWAAAAAAYNAAAAAgEAAwUA/3//v//fvwEBDQQWAAAAAAYNAAAAAgEAAwUA/3//v//fvwEBDQUWAAAAAAYNAAAAAgEAAwUA/3//v//fvwEBDQMJAAAAAQYDAAAAAgEBCycAAAABEQAAAAAMAAAAAAYDAAAAAgEJAQENAwwAAAAAAQEBBgMAAAACAQELVwAAAAFEAAAAAAwAAAAABgMAAAACAQkBAQ0CDAAAAAAGAwAAAAIBCQEBDQQMAAAAAAYDAAAAAgEJAQENBQwAAAAABgMAAAACAQkBAQ0DCQAAAAEGAwAAAAIBAQsnAAAAAREAAAAADAAAAAAGAwAAAAIBCAEBDQMMAAAAAAEBAQYDAAAAAgEBC1cAAAABRAAAAAAMAAAAAAYDAAAAAgEIAQENAgwAAAAABgMAAAACAQgBAQ0EDAAAAAAGAwAAAAIBCAEBDQUMAAAAAAYDAAAAAgEIAQENAwkAAAABBgMAAAACAQELJwAAAAERAAAAAAwAAAAABgMAAAACAQcBAQ0DDAAAAAABAQEGAwAAAAIBAQtXAAAAAUQAAAAADAAAAAAGAwAAAAIBBwEBDQIMAAAAAAYDAAAAAgEHAQENBAwAAAAABgMAAAACAQcBAQ0FDAAAAAAGAwAAAAIBBwEBDQMJAAAAAQYDAAAAAgEBCycAAAABEQAAAAAMAAAAAAYDAAAAAgEGAQENAwwAAAAAAQEBBgMAAAACAQELVwAAAAFEAAAAAAwAAAAABgMAAAACAQYBAQ0CDAAAAAAGAwAAAAIBBgEBDQQMAAAAAAYDAAAAAgEGAQENBQwAAAAABgMAAAACAQYBAQ0DCQAAAAEGAwAAAAIBAQsnAAAAAREAAAAADAAAAAAGAwAAAAIBBQEBDQMMAAAAAAEBAQYDAAAAAgEBC1cAAAABRAAAAAAMAAAAAAYDAAAAAgEFAQENAgwAAAAABgMAAAACAQUBAQ0EDAAAAAAGAwAAAAIBBQEBDQUMAAAAAAYDAAAAAgEFAQENAwkAAAABBgMAAAACAQELJwAAAAERAAAAAAwAAAAABgMAAAACAQQBAQ0DDAAAAAABAQEGAwAAAAIBAQtXAAAAAUQAAAAADAAAAAAGAwAAAAIBBAEBDQIMAAAAAAYDAAAAAgEEAQENBAwAAAAABgMAAAACAQQBAQ0FDAAAAAAGAwAAAAIBBAEBDQMJAAAAAQYDAAAAAgEBCycAAAABEQAAAAAMAAAAAAYDAAAAAgEJAQENAwwAAAAAAQEBBgMAAAACAQELVwAAAAFEAAAAAAwAAAAABgMAAAACAQkBAQ0CDAAAAAAGAwAAAAIBCQEBDQQMAAAAAAYDAAAAAgEJAQENBQwAAAAABgMAAAACAQkBAQ0DCQAAAAEGAwAAAAIBAQsnAAAAAREAAAAADAAAAAAGAwAAAAIBCAEBDQMMAAAAAAEBAQYDAAAAAgEBC1cAAAABRAAAAAAMAAAAAAYDAAAAAgEIAQENAgwAAAAABgMAAAACAQgBAQ0EDAAAAAAGAwAAAAIBCAEBDQUMAAAAAAYDAAAAAgEIAQENAwkAAAABBgMAAAACAQELJwAAAAERAAAAAAwAAAAABgMAAAACAQcBAQ0DDAAAAAABAQEGAwAAAAIBAQtXAAAAAUQAAAAADAAAAAAGAwAAAAIBBwEBDQIMAAAAAAYDAAAAAgEHAQENBAwAAAAABgMAAAACAQcBAQ0FDAAAAAAGAwAAAAIBBwEBDQMJAAAAAQYDAAAAAgEBCycAAAABEQAAAAAM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+gAGASkAAAD7AAkAAAD6AAcBKAAAAPsApgAAAPoAEwAAAFMAbABpAGMAZQByAFMAdAB5AGwAZQBMAGkAZwBoAHQANAAgADIA+wB0AAAACAAAAAAJAAAA+gAAAScAAAD7AAkAAAD6AAIBJgAAAPsACQAAAPoAAQElAAAA+wAJAAAA+gADASQAAAD7AAkAAAD6AAQBIwAAAPsACQAAAPoABQEiAAAA+wAJAAAA+gAGASEAAAD7AAkAAAD6AAcBIAAAAPsApgAAAPoAEwAAAFMAbABpAGMAZQByAFMAdAB5AGwAZQBMAGkAZwBoAHQANQAgADIA+wB0AAAACAAAAAAJAAAA+gAAAR8AAAD7AAkAAAD6AAIBHgAAAPsACQAAAPoAAQEdAAAA+wAJAAAA+gADARwAAAD7AAkAAAD6AAQBGwAAAPsACQAAAPoABQEaAAAA+wAJAAAA+gAGARkAAAD7AAkAAAD6AAcBGAAAAPsApgAAAPoAEwAAAFMAbABpAGMAZQByAFMAdAB5AGwAZQBMAGkAZwBoAHQANgAgADIA+wB0AAAACAAAAAAJAAAA+gAAARcAAAD7AAkAAAD6AAIBFgAAAPsACQAAAPoAAQEVAAAA+wAJAAAA+gADARQAAAD7AAkAAAD6AAQBEwAAAPsACQAAAPoABQESAAAA+wAJAAAA+gAGAREAAAD7AAkAAAD6AAcBEAAAAPsApgAAAPoAEwAAAFMAbABpAGMAZQByAFMAdAB5AGwAZQBPAHQAaABlAHIAMQAgADIA+wB0AAAACAAAAAAJAAAA+gAAAQ8AAAD7AAkAAAD6AAIBDgAAAPsACQAAAPoAAQENAAAA+wAJAAAA+gADAQwAAAD7AAkAAAD6AAQBCwAAAPsACQAAAPoABQEKAAAA+wAJAAAA+gAGAQkAAAD7AAkAAAD6AAcBCAAAAPsApgAAAPoAEwAAAFMAbABpAGMAZQByAFMAdAB5AGwAZQBPAHQAaABlAHIAMgAgADIA+wB0AAAACAAAAAAJAAAA+gAAAQcAAAD7AAkAAAD6AAIBBgAAAPsACQAAAPoAAQEFAAAA+wAJAAAA+gADAQQAAAD7AAkAAAD6AAQBAwAAAPsACQAAAPoABQECAAAA+wAJAAAA+gAGAQEAAAD7AAkAAAD6AAcBAAAAAPs=";

        var oBinaryFileReader = new BinaryFileReader();
        var stream = oBinaryFileReader.getbase64DecodedData(sStyles);
        new Binary_CommonReader(stream);
        var oBinary_StylesTableReader = new Binary_StylesTableReader(stream, wb, true/*, [], undefined, true*/);
        var oStyleObject = oBinary_StylesTableReader.Read();
        RenameDefSlicerStyle(oStyleObject);
        AscCommonExcel.InitOpenManager.prototype.InitDefSlicerStyles.call(AscCommonExcel.InitOpenManager.prototype, wb, oStyleObject)
    }

	function CT_PresetTableStyles(tableStyles, pivotStyles) {
		this.tableStyles = tableStyles;
		this.pivotStyles = pivotStyles;
	}

	CT_PresetTableStyles.prototype.onStartNode = function(elem, attr, uq) {
		var newContext = this;
		if ("presetTableStyles" === elem) {
		} else if (0 === elem.indexOf("TableStyle") || 0 === elem.indexOf("PivotStyle")) {
			newContext = new CT_Stylesheet(new Asc.CTableStyles());
		} else {
			newContext = null;
		}
		return newContext;
	};
	CT_PresetTableStyles.prototype.onEndNode = function(prevContext, elem) {
		if (0 === elem.indexOf("TableStyle")) {
			for (var i in prevContext.tableStyles.CustomStyles) {
				this.tableStyles[i] = prevContext.tableStyles.CustomStyles[i];
			}
		} else if (0 === elem.indexOf("PivotStyle")) {
			for (var i in prevContext.tableStyles.CustomStyles) {
				this.pivotStyles[i] = prevContext.tableStyles.CustomStyles[i];
			}
		}
	};

	function CT_Stylesheet(tableStyles) {
		//Members
		this.numFmts = [];
		this.fonts = [];
		this.fills = [];
		this.borders = [];
		this.cellStyleXfs = [];
		this.cellXfs = [];
		this.cellStyles = [];
		this.dxfs = [];
		this.tableStyles = tableStyles;

		this.oCustomSlicerStyles = null;
		this.oTimelineStyles = null;
	}

	CT_Stylesheet.prototype.onStartNode = function(elem, attr, uq) {
		var newContext = this;
		if ("styleSheet" === elem) {
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
		} else if ("numFmts" === elem) {
			;
		} else if ("numFmt" === elem) {
			newContext = new AscCommonExcel.Num();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.numFmts.push(newContext);
		} else if ("fonts" === elem) {
			;
		} else if ("font" === elem) {
			newContext = new AscCommonExcel.Font();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.fonts.push(newContext);
		} else if ("fills" === elem) {
			AscCommon.openXml.SaxParserDataTransfer.priorityBg = false;
		} else if ("fill" === elem) {
			newContext = new AscCommonExcel.Fill();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.fills.push(newContext);
		} else if ("borders" === elem) {
			;
		} else if ("border" === elem) {
			newContext = new AscCommonExcel.Border();
            newContext.initDefault();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.borders.push(newContext);
			// } else if("cellStyleXfs" === elem){
			// 	newContext = new CT_CellStyleXfs();
			// 	if(newContext.readAttributes){
			// 		newContext.readAttributes(attr, uq);
			// 	}
			// 	this.cellStyleXfs = newContext;
			// } else if("cellXfs" === elem){
			// 	newContext = new CT_CellXfs();
			// 	if(newContext.readAttributes){
			// 		newContext.readAttributes(attr, uq);
			// 	}
			// 	this.cellXfs = newContext;
			// } else if("cellStyles" === elem){
			// 	newContext = new CT_CellStyles();
			// 	if(newContext.readAttributes){
			// 		newContext.readAttributes(attr, uq);
			// 	}
			// 	this.cellStyles = newContext;
		} else if ("dxfs" === elem) {
			AscCommon.openXml.SaxParserDataTransfer.dxfs = this.dxfs;
			AscCommon.openXml.SaxParserDataTransfer.priorityBg = true;
		} else if ("dxf" === elem) {
			newContext = new CT_Dxf();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
		} else if ("tableStyles" === elem) {
			newContext = this.tableStyles;
		} else {
			newContext = null;
		}
		return newContext;
	};
	CT_Stylesheet.prototype.onEndNode = function(prevContext, elem) {
		if ("dxf" === elem) {
			this.dxfs.push(g_StyleCache.addXf(prevContext.xf));
		}
	};

	function CT_Dxf() {
		//Members
		this.xf = new AscCommonExcel.CellXfs();
	}

	CT_Dxf.prototype.onStartNode = function(elem, attr, uq) {
		var newContext = this;
		if ("font" === elem) {
			newContext = new AscCommonExcel.Font();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.xf.font = newContext;
		} else if ("numFmt" === elem) {
			newContext = new AscCommonExcel.Num();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.xf.num = newContext;
		} else if ("fill" === elem) {
			newContext = new AscCommonExcel.Fill();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.xf.fill = newContext;
		} else if ("alignment" === elem) {
			newContext = new AscCommonExcel.Align();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.xf.align = newContext;
		} else if ("border" === elem) {
			newContext = new AscCommonExcel.Border();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.xf.border = newContext;
		} else {
			newContext = null;
		}
		return newContext;
	};

    function InitOpenManager(isCopyPaste, wb) {
        this.copyPasteObj = {
            isCopyPaste: isCopyPaste, activeRange: null, selectAllSheet: null
        };
        this.oReadResult = {
            tableCustomFunc: [],
            sheetData: [],
            stylesTableReader: null,
            pivotCacheDefinitions: {},
            vbaMacros: null,
            vbaProject: null,
            macros: null,
            slicerCaches: {},
            tableIds: {},
            defNames: [],
            sheetIds: {},
            customFunctions: null
        };
        this.wb = wb;
        this.Dxfs = [];

        //при чтении из xml
        this.legacyDrawingId = null;
    }

    InitOpenManager.prototype.initSchemeAndTheme = function (wb) {
        if(!this.copyPasteObj || !this.copyPasteObj.isCopyPaste) {
            wb.clrSchemeMap = AscFormat.GenerateDefaultColorMap();
            if(null == wb.theme)
                wb.theme = AscFormat.GenerateDefaultTheme(wb, 'Calibri');

            Asc.getBinaryOtherTableGVar(wb);
        }
    };

    InitOpenManager.prototype.PostLoadPrepare = function(wb)
    {
        if (wb.workbookPr && null != wb.workbookPr.Date1904) {
            AscCommon.bDate1904 = wb.workbookPr.Date1904;
            AscCommonExcel.c_DateCorrectConst = AscCommon.bDate1904?AscCommonExcel.c_Date1904Const:AscCommonExcel.c_Date1900Const;
        }
        if (this.oReadResult.macros) {
            let parsedMacros = AscCommonExcel.safeJsonParse(this.oReadResult.macros);
            if (parsedMacros) {
                if (parsedMacros["customFunctions"]) {
                    wb.fileCustomFunctions = parsedMacros["customFunctions"];
                }
                delete parsedMacros["customFunctions"];
                wb.oApi.macros.SetData(JSON.stringify(parsedMacros));
            } else {
                wb.oApi.macros.SetData(this.oReadResult.macros);
            }
        }
        if (this.oReadResult.vbaMacros) {
            wb.oApi.vbaMacros = this.oReadResult.vbaMacros;
        }
        if (this.oReadResult.vbaProject) {
            wb.oApi.vbaProject = this.oReadResult.vbaProject;
        }

        wb.checkCorrectTables();
    };
    InitOpenManager.prototype.PostLoadPrepareDefNames = function(wb)
    {
        this.oReadResult.defNames.forEach(function (defName) {
            if (null != defName.Name && null != defName.Ref) {
                var _type = Asc.c_oAscDefNameType.none;
                if (wb.getSlicerCacheByName(defName.Name)) {
                    _type = Asc.c_oAscDefNameType.slicer;
                }
                wb.dependencyFormulas.addDefNameOpen(defName.Name, defName.Ref, defName.LocalSheetId, defName.Hidden, _type);
            }
        });
    };

    InitOpenManager.prototype.initCellAfterRead = function(tmp)
    {
        //use only excel
        if(!(this.copyPasteObj && this.copyPasteObj.isCopyPaste && typeof editor != "undefined" && editor)) {
            this.setFormulaOpen(tmp);
        }
        tmp.cell.saveContent();
        if (tmp.cell.nCol >= tmp.ws.nColsCount) {
            tmp.ws.nColsCount = tmp.cell.nCol + 1;
        }
    };

    InitOpenManager.prototype.setFormulaOpen = function(tmp) {
        var cell = tmp.cell;
        var formula = tmp.formula;
        var curFormula;
        var prevFormula = tmp.prevFormulas[cell.nCol];
        if (null !== formula.si && (curFormula = tmp.sharedFormulas[formula.si])) {
            curFormula.parsed.getShared().ref.union3(cell.nCol, cell.nRow);
            if (prevFormula !== curFormula) {
                if (prevFormula && !tmp.bNoBuildDep && !tmp.siFormulas[prevFormula.parsed.getListenerId()]) {
                    prevFormula.parsed.buildDependencies();
                }
                tmp.prevFormulas[cell.nCol] = curFormula;
            }
        } else if (formula.v && formula.v.length <= AscCommon.c_oAscMaxFormulaLength) {
            if (formula.v.startsWith("_xludf.")) {
                //при открытии подобных формул ms удаляет префикс
                //TODO так же он проставляет флаг ca - рассмотреть стоит ли его нам доблавлять
                formula.v = formula.v.replace("_xludf.", "");
            }
            if (formula.v.startsWith("IFERROR(__xludf.DUMMYFUNCTION(\"")) {
                formula.v = formula.v.replace('IFERROR(__xludf.DUMMYFUNCTION(\"', "");
                formula.v = formula.v.substr(0, formula.v.lastIndexOf('\"\)'));
                formula.v = formula.v.replace(/\"\"/g,"\"");
            }
            if (formula.v.startsWith("=")) {
                //LO write "=" to file
                formula.v = formula.v.replace("=", "");
            }

            var offsetRow;
            var shared;
            var sharedRef;
            if (prevFormula && (shared = prevFormula.parsed.getShared())) {
                offsetRow = cell.nRow - shared.ref.r1;
            } else {
                offsetRow = 1;
            }
            //проверка на ECellFormulaType.cellformulatypeArray нужна для:
            //1.формула массива не может быть шаренной
            //2.в случае, когда две ячейки в одном столбце - каждая формула массива
            //и далее они становятся двумя шаренными
            //после того, как формула становится шаренной, ref array у второй начинает ссылаться на первую ячейку
            //поэтому при изменении второй ячейки из двух шаренных в функции _saveCellValueAfterEdit
            //берём array ref и присваиваем ему введенные данные, и поэтому в первой ячейки появляются данные второй
            if (prevFormula && formula.t !== ECellFormulaType.cellformulatypeArray &&
                prevFormula.t !== ECellFormulaType.cellformulatypeArray &&
                prevFormula.nRow + offsetRow === cell.nRow &&
                AscCommonExcel.compareFormula(prevFormula.formula, prevFormula.refPos, formula.v, offsetRow)) {
                if (!(shared && shared.ref)) {
                    sharedRef = new Asc.Range(cell.nCol, prevFormula.nRow, cell.nCol, cell.nRow);
                    prevFormula.parsed.setShared(sharedRef, prevFormula.base);
                } else {
                    shared.ref.union3(cell.nCol, cell.nRow);
                }
                curFormula = prevFormula;
            } else {
                if (prevFormula && !tmp.bNoBuildDep && !tmp.siFormulas[prevFormula.parsed.getListenerId()]) {
                    prevFormula.parsed.buildDependencies();
                }
                var parseResult = new AscCommonExcel.ParseResult([]);
                var newFormulaParent = new AscCommonExcel.CCellWithFormula(cell.ws, cell.nRow, cell.nCol);
                var parsed = new AscCommonExcel.parserFormula(formula.v, newFormulaParent, cell.ws);
                parsed.ca = formula.ca;
                parsed.parse(undefined, undefined, parseResult);
                if (parseResult.error === Asc.c_oAscError.ID.FrmlMaxReference) {
                    tmp.ws.workbook.openErrors.push(cell.getName());
                    return;
                }

				if (parsed.importFunctionsRangeLinks) {
					for (let i in parsed.importFunctionsRangeLinks) {
						let eR = tmp.ws.workbook.getExternalLinkByName(i);
						if (eR) {
							eR.notUpdateId = true;
						}
					}
				}

                if (null !== formula.ref) {
                    let range;
                    if(formula.t === ECellFormulaType.cellformulatypeShared) {
                        range = AscCommonExcel.g_oRangeCache.getAscRange(formula.ref);
                        sharedRef = range && range.clone();
                        if (sharedRef) {
                            parsed.setShared(sharedRef, newFormulaParent);
                        }
                    } else if(formula.t === ECellFormulaType.cellformulatypeArray) {//***array-formula***
                        if(AscCommonExcel.bIsSupportArrayFormula) {
                            range = AscCommonExcel.g_oRangeCache.getAscRange(formula.ref);
                            range = range && range.clone();
                            if (range) {
                                parsed.setArrayFormulaRef(range);
                                tmp.formulaArray.push(parsed);
                            }
                        }
                    }
                }

                curFormula = new OpenColumnFormula(cell.nRow, formula.v, parsed, parseResult.refPos, newFormulaParent);
                tmp.prevFormulas[cell.nCol] = curFormula;
            }
            if (null !== formula.si && curFormula.parsed.getShared()) {
                tmp.sharedFormulas[formula.si] = curFormula;
                tmp.siFormulas[curFormula.parsed.getListenerId()] = curFormula.parsed;
            }
        } else if (formula.v && !(this.copyPasteObj && this.copyPasteObj.isCopyPaste)) {
            tmp.ws.workbook.openErrors.push(cell.getName());
        }
        if (curFormula) {
            cell.setFormulaInternal(curFormula.parsed);
            if (curFormula.parsed.ca || cell.isNullTextString()) {
                tmp.ws.workbook.dependencyFormulas.addToChangedCell(cell);
            }
        }
    };
    InitOpenManager.prototype.InitStyleManager = function (oStyleObject, aCellXfs)
    {
        var i, xf, firstFont, firstFill, secondFill, firstBorder, firstXf, newXf, oCellStyle;
        if (0 === oStyleObject.aFonts.length) {
            oStyleObject.aFonts[0] = new AscCommonExcel.Font();
            oStyleObject.aFonts[0].initDefault(this.wb);
        }
        if (0 === oStyleObject.aCellXfs.length) {
            xf = new OpenXf();
            xf.fontid = xf.fillid = xf.borderid = xf.numid = xf.XfId = 0;
            oStyleObject.aCellXfs[0] = xf;
        }
        if (0 === oStyleObject.aCellStyleXfs.length) {
            xf = new OpenXf();
            xf.fontid = xf.fillid = xf.borderid = xf.numid = 0;
            oStyleObject.aCellStyleXfs[0] = xf;
        }
        var hasNormalStyle = false;
        for (i = 0; i < oStyleObject.aCellStyles.length; ++i) {
            oCellStyle = oStyleObject.aCellStyles[i];
            if (0 === oCellStyle.BuiltinId) {
                hasNormalStyle = true;
                break;
            }
        }
        if (!hasNormalStyle) {
            oCellStyle = new AscCommonExcel.CCellStyle();
            oCellStyle.Name = "Normal";
            oCellStyle.BuiltinId = 0;
            oCellStyle.XfId = 0;
            oStyleObject.aCellStyles.push(oCellStyle);
        }

        var defFont = oStyleObject.aFonts[oStyleObject.aCellXfs[0].fontid];
        if (defFont) {
            defFont.initDefault(this.wb);
        }

        for (i = 0; i < oStyleObject.aFonts.length; ++i) {
            oStyleObject.aFonts[i] = g_StyleCache.addFont(oStyleObject.aFonts[i]);
        }
        firstFont = oStyleObject.aFonts[0];

        for (i = 2; i < oStyleObject.aFills.length; ++i) {
            oStyleObject.aFills[i] = g_StyleCache.addFill(oStyleObject.aFills[i]);
        }
        //addXf with force flag should be last operation
        firstFill = new AscCommonExcel.Fill();
        firstFill.fromPatternParams(AscCommonExcel.c_oAscPatternType.None, null);
        secondFill = new AscCommonExcel.Fill();
        secondFill.fromPatternParams(AscCommonExcel.c_oAscPatternType.Gray125, null);
        if (!this.copyPasteObj.isCopyPaste) {
            firstFill = g_StyleCache.addFill(firstFill, true);
            secondFill = g_StyleCache.addFill(secondFill, true);
        } else {
            firstFill = g_StyleCache.addFill(firstFill);
            secondFill = g_StyleCache.addFill(secondFill);
        }
        oStyleObject.aFills[0] = firstFill;
        oStyleObject.aFills[1] = secondFill;

        oStyleObject.aBorders[0] = new AscCommonExcel.Border();
        oStyleObject.aBorders[0].initDefault();
        for (i = 0; i < oStyleObject.aBorders.length; ++i) {
            oStyleObject.aBorders[i] = g_StyleCache.addBorder(oStyleObject.aBorders[i]);
        }
        firstBorder = oStyleObject.aBorders[0];
        for (i = 0; i < oStyleObject.aCellStyleXfs.length; ++i) {
            xf = oStyleObject.aCellStyleXfs[i];
            if (xf.align) {
                xf.align = g_StyleCache.addAlign(xf.align);
            }
        }
        for (i = 0; i < oStyleObject.aCellXfs.length; ++i) {
            xf = oStyleObject.aCellXfs[i];
            if (xf.align) {
                xf.align = g_StyleCache.addAlign(xf.align);
            }
        }
        this.InitDxfs(oStyleObject.aDxfs);
        this.InitDxfs(oStyleObject.aExtDxfs);

        // ToDo убрать - это заглушка
        var arrStyleMap = {};
        var nIndexStyleMap = 1;//0 reserver for Normal style
        var XfIdTmp;
        // Список имен для стилей
        var oCellStyleNames = {};
        var normalXf = null;

        for (i = 0; i < oStyleObject.aCellStyles.length; ++i) {
            oCellStyle = oStyleObject.aCellStyles[i];
            newXf = new AscCommonExcel.CellXfs();
            // XfId
            XfIdTmp = oCellStyle.XfId;
            if (null !== XfIdTmp) {
                if (0 === oCellStyle.BuiltinId) {
                    arrStyleMap[XfIdTmp] = 0;
                    if (!normalXf) {
                        XfIdTmp = oCellStyle.XfId = 0;
                        normalXf = newXf;
                        //default fontid is always 0
                        if (oStyleObject.aCellStyleXfs[XfIdTmp]) {
                            oStyleObject.aCellStyleXfs[XfIdTmp].fontid = 0;
                        }
                    } else {
                        continue;
                    }
                } else {
                    arrStyleMap[XfIdTmp] = nIndexStyleMap;
                    oCellStyle.XfId = nIndexStyleMap++;
                }
            } else
                continue;	// Если его нет, то это ошибка по спецификации

            var oCellStyleXfs = oStyleObject.aCellStyleXfs[XfIdTmp];
            // Если есть стиль, но нет описания, то уберем этот стиль (Excel делает также)
            if (null == oCellStyleXfs)
                continue;

            // Border
            if (null != oCellStyleXfs.borderid) {
                var borderCellStyle = oStyleObject.aBorders[oCellStyleXfs.borderid];
                if(null != borderCellStyle)
                    newXf.border = borderCellStyle;
            }
            // Fill
            if (null != oCellStyleXfs.fillid) {
                var fillCellStyle = oStyleObject.aFills[oCellStyleXfs.fillid];
                if(null != fillCellStyle)
                    newXf.fill = fillCellStyle;
            }
            // Font
            if(null != oCellStyleXfs.fontid) {
                var fontCellStyle = oStyleObject.aFonts[oCellStyleXfs.fontid];
                if(null != fontCellStyle)
                    newXf.font = fontCellStyle;
            }
            // NumFmt
            if(null != oCellStyleXfs.numid) {
                var oCurNumCellStyle = oStyleObject.oNumFmts[oCellStyleXfs.numid];
                if(null != oCurNumCellStyle)
                    newXf.num = g_StyleCache.addNum(oCurNumCellStyle);
                else
                    newXf.num = g_StyleCache.addNum(this.ParseNum({id: oCellStyleXfs.numid, f: null}, oStyleObject.oNumFmts));
            }
            // QuotePrefix
            if(null != oCellStyleXfs.QuotePrefix)
                newXf.QuotePrefix = oCellStyleXfs.QuotePrefix;
            //PivotButton
            if(null != oCellStyleXfs.PivotButton)
                newXf.PivotButton = oCellStyleXfs.PivotButton;
            // hidden
            if(null != oCellStyleXfs.hidden)
                newXf.hidden = oCellStyleXfs.hidden;
            // locked
            if(null != oCellStyleXfs.locked)
                newXf.locked = oCellStyleXfs.locked;
            if(null != oCellStyleXfs.applyProtection)
                newXf.applyProtection = oCellStyleXfs.applyProtection;
            // align
            if(null != oCellStyleXfs.align)
                newXf.align = oCellStyleXfs.align;
            // ApplyBorder (ToDo возможно это свойство должно быть в xfs)
            if (null !== oCellStyleXfs.ApplyBorder)
                oCellStyle.ApplyBorder = oCellStyleXfs.ApplyBorder;
            // ApplyFill (ToDo возможно это свойство должно быть в xfs)
            if (null !== oCellStyleXfs.ApplyFill)
                oCellStyle.ApplyFill = oCellStyleXfs.ApplyFill;
            // ApplyFont (ToDo возможно это свойство должно быть в xfs)
            if (null !== oCellStyleXfs.ApplyFont)
                oCellStyle.ApplyFont = oCellStyleXfs.ApplyFont;
            // ApplyNumberFormat (ToDo возможно это свойство должно быть в xfs)
            if (null !== oCellStyleXfs.ApplyNumberFormat)
                oCellStyle.ApplyNumberFormat = oCellStyleXfs.ApplyNumberFormat;

            oCellStyle.xfs = g_StyleCache.addXf(newXf);
            // ToDo при отсутствии имени все не очень хорошо будет!
            this.wb.CellStyles.CustomStyles.push(oCellStyle);
            if (null !== oCellStyle.Name)
                oCellStyleNames[oCellStyle.Name] = true;
        }

        // ToDo стоит это переделать в дальнейшем (пробежимся по именам, и у отсутствующих создадим имя)
        var nNewStyleIndex = 1, newStyleName;
        for (var i = 0, length = this.wb.CellStyles.CustomStyles.length; i < length; ++i) {
            if (null === this.wb.CellStyles.CustomStyles[i].Name) {
                do {
                    newStyleName = "Style" + nNewStyleIndex++;
                } while (oCellStyleNames[newStyleName])
                    ;
                this.wb.CellStyles.CustomStyles[i].Name = newStyleName;
            }
        }

        // ToDo это нужно будет переделать (проходимся по всем стилям и меняем у них XfId по порядку)

        for(var i = 0, length = oStyleObject.aCellXfs.length; i < length; ++i) {
            var xfs = oStyleObject.aCellXfs[i];
            newXf = new AscCommonExcel.CellXfs();

            if(null != xfs.borderid)
            {
                var border = oStyleObject.aBorders[xfs.borderid];
                if(null != border)
                    newXf.border = border;
            }
            if(null != xfs.fillid)
            {
                var fill = oStyleObject.aFills[xfs.fillid];
                if(null != fill)
                    newXf.fill = fill;
            }
            if(null != xfs.fontid)
            {
                var font = oStyleObject.aFonts[xfs.fontid];
                if(null != font)
                    newXf.font = font;
            }
            if(null != xfs.numid)
            {
                var oCurNum = oStyleObject.oNumFmts[xfs.numid];
                //todo
                if(null != oCurNum)
                    newXf.num = g_StyleCache.addNum(oCurNum);
                else
                    newXf.num = g_StyleCache.addNum(this.ParseNum({id: xfs.numid, f: null}, oStyleObject.oNumFmts));
            }
            if(null != xfs.QuotePrefix)
                newXf.QuotePrefix = xfs.QuotePrefix;
            if(null != xfs.PivotButton)
                newXf.PivotButton = xfs.PivotButton;
            // hidden
            if(null != xfs.hidden)
                newXf.hidden = xfs.hidden;
            // locked
            if(null != xfs.locked)
                newXf.locked = xfs.locked;
            if(null != xfs.applyProtection)
                newXf.applyProtection = xfs.applyProtection;
            if(null != xfs.align)
                newXf.align = xfs.align;
            if (null !== xfs.XfId) {
                XfIdTmp = arrStyleMap[xfs.XfId];
                if (null == XfIdTmp)
                    XfIdTmp = 0;
                newXf.XfId = XfIdTmp;
            }

            if (0 === aCellXfs.length && !this.copyPasteObj.isCopyPaste) {
                firstXf = newXf;
            } else {
                newXf = g_StyleCache.addXf(newXf);
            }
            aCellXfs.push(newXf);
        }
        if (firstXf && !this.copyPasteObj.isCopyPaste) {
            //addXf with force flag should be last operation
            firstXf = g_StyleCache.addXf(firstXf, true);
            this.wb.oStyleManager.init(firstXf, firstFont, firstFill, secondFill, firstBorder, normalXf);
        }
        this.InitTableStyles(this.wb.TableStyles.CustomStyles, oStyleObject.oCustomTableStyles, oStyleObject.aDxfs);
        this.wb.SlicerStyles.addCustomStylesAtOpening(oStyleObject.oCustomSlicerStyles, oStyleObject.aExtDxfs);
    };
    InitOpenManager.prototype.InitDefSlicerStyles = function (wb, oStyleObject)
    {
        this.InitDxfs(oStyleObject.aDxfs);
        this.InitDxfs(oStyleObject.aExtDxfs);
        this.InitTableStyles(wb.TableStyles.DefaultStyles, oStyleObject.oCustomTableStyles, oStyleObject.aDxfs);
        wb.SlicerStyles.addDefaultStylesAtOpening(oStyleObject.oCustomSlicerStyles, oStyleObject.aExtDxfs);
    };
    InitOpenManager.prototype.InitDxfs = function (Dxfs)
    {
        if (!Dxfs) {
            return;
        }
        for (var i = 0; i < Dxfs.length; ++i) {
            Dxfs[i] = g_StyleCache.addXf(Dxfs[i]);
        }
    };
    InitOpenManager.prototype.InitTableStyles = function (tableStyles, oCustomTableStyles, aDxfs)
    {
        for(var i in oCustomTableStyles)
        {
            var item = oCustomTableStyles[i];
            if(null != item)
            {
                var style = item.style;
                var elems = item.elements;
                this.initTableStyle(style, elems, aDxfs);
                tableStyles[i] = style;
            }
        }
    };
    InitOpenManager.prototype.initTableStyle = function(style, elems, Dxfs)
    {
        for(var j = 0, length2 = elems.length; j < length2; ++j)
        {
            var elem = elems[j];
            if(null != elem.DxfId)
            {
                var Dxf = Dxfs[elem.DxfId];
                if(null != Dxf)
                {
                    var oTableStyleElement = new CTableStyleElement();
                    oTableStyleElement.dxf = Dxf;
                    if(null != elem.Size)
                        oTableStyleElement.size = elem.Size;
                    switch(elem.Type)
                    {
                        case ETableStyleType.tablestyletypeBlankRow: style.blankRow = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeFirstColumn: style.firstColumn = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeFirstColumnStripe: style.firstColumnStripe = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeFirstColumnSubheading: style.firstColumnSubheading = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeFirstHeaderCell: style.firstHeaderCell = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeFirstRowStripe: style.firstRowStripe = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeFirstRowSubheading: style.firstRowSubheading = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeFirstSubtotalColumn: style.firstSubtotalColumn = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeFirstSubtotalRow: style.firstSubtotalRow = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeFirstTotalCell: style.firstTotalCell = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeHeaderRow: style.headerRow = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeLastColumn: style.lastColumn = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeLastHeaderCell: style.lastHeaderCell = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeLastTotalCell: style.lastTotalCell = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypePageFieldLabels: style.pageFieldLabels = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypePageFieldValues: style.pageFieldValues = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeSecondColumnStripe: style.secondColumnStripe = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeSecondColumnSubheading: style.secondColumnSubheading = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeSecondRowStripe: style.secondRowStripe = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeSecondRowSubheading: style.secondRowSubheading = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeSecondSubtotalColumn: style.secondSubtotalColumn = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeSecondSubtotalRow: style.secondSubtotalRow = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeThirdColumnSubheading: style.thirdColumnSubheading = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeThirdRowSubheading: style.thirdRowSubheading = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeThirdSubtotalColumn: style.thirdSubtotalColumn = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeThirdSubtotalRow: style.thirdSubtotalRow = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeTotalRow: style.totalRow = oTableStyleElement;break;
                        case ETableStyleType.tablestyletypeWholeTable: style.wholeTable = oTableStyleElement;break;
                    }
                }
            }
        }
    };
    InitOpenManager.prototype.ParseNum = function(oNum, oNumFmts, _useNumId) {
        var oRes = new AscCommonExcel.Num();
        var useNumId = false;
        if (null != oNum && oNum.f) {//Excel ignors empty format. bug 70667
            oRes.f = oNum.f;
        } else {
            var sStandartNumFormat = AscCommonExcel.aStandartNumFormats[oNum.id];
            if (null != sStandartNumFormat) {
                oRes.f = sStandartNumFormat;
            }
            if (null == oRes.f) {
                oRes.f = "General";
            }
            //format string is more priority then id. so, fill oRes.id only if format is empty
            useNumId = true;
        }
        if ((useNumId || _useNumId) && AscCommon.canGetFormatByStandardId(oNum.id)) {
            oRes.id = oNum.id;
        }
        var numFormat = AscCommon.oNumFormatCache.get(oRes.f);
        numFormat.checkCultureInfoFontPicker();
        if (null != oNumFmts) {
            oNumFmts[oNum.id] = oRes;
        }
        return oRes;
    };
    InitOpenManager.prototype.readDefStyles = function(wb, output) {
        ReadDefCellStyles(wb, output);
        ReadDefSlicerStyles(wb, output);
        wb.SlicerStyles.concatStyles();
    };
    InitOpenManager.prototype.prepareAfterReadCols = function(oWorksheet, aTempCols) {
        //если есть стиль последней колонки, назначаем его стилем всей таблицы и убираем из колонок
        var oAllCol = null;
        if(aTempCols.length > 0)
        {
            var oLast = aTempCols[aTempCols.length - 1];
            if(AscCommon.gc_nMaxCol == oLast.Max)
            {
                oAllCol = oWorksheet.getAllCol();
                oLast.col.cloneTo(oAllCol);
            }
        }
        for(var i = 0; i < aTempCols.length; ++i)
        {
            var elem = aTempCols[i];
            if(null != oAllCol && oAllCol.isEqual(elem.col))
                continue;
            if(elem.col.isUpdateScroll() && elem.Max >= oWorksheet.nColsCount)
                oWorksheet.nColsCount = elem.Max;

            for(var j = elem.Min; j <= elem.Max; j++){
                var oNewCol = new AscCommonExcel.Col(oWorksheet, j - 1);
                elem.col.cloneTo(oNewCol);
                oWorksheet.aCols[oNewCol.index] = oNewCol;
            }
        }
    };
    InitOpenManager.prototype.prepareAfterReadHyperlinks = function (aHyperlinks) {
        for (var i = 0, length = aHyperlinks.length; i < length; ++i) {
            var hyperlink = aHyperlinks[i];
            if (null !== hyperlink.Ref) {
                hyperlink.Ref.setHyperlinkOpen(hyperlink);
            }
        }
    };
    InitOpenManager.prototype.prepareAfterReadMergedCells = function (oNewWorksheet, aMerged) {
        oNewWorksheet.mergeManager.initData = aMerged.slice();
    };
    InitOpenManager.prototype.prepareConditionalFormatting = function (oWorksheet, oConditionalFormatting) {
        if (oConditionalFormatting && oConditionalFormatting.isValid()) {
            oConditionalFormatting.initRules();
            for (let i = 0; i < oConditionalFormatting.aRules.length; i++) {
                oWorksheet.addConditionalFormattingRule(oConditionalFormatting.aRules[i]);
            }
        }
    };
    InitOpenManager.prototype.putSheetAfterRead = function (wb, oNewWorksheet) {
        wb.aWorksheets.push(oNewWorksheet);
        wb.aWorksheetsById[oNewWorksheet.getId()] = oNewWorksheet;
    };

    InitOpenManager.prototype.prepareComments = function (oWorksheet, oCommentCoords, aCommentData, oAdditionalData) {
        if (oCommentCoords.isValid()) {
            for(var i = 0, length = aCommentData.length; i < length; ++i)
            {
                aCommentData[i].coords = oCommentCoords;

                var elem = aCommentData[i];
                elem.asc_putRow(oCommentCoords.nRow);
                elem.asc_putCol(oCommentCoords.nCol);

                if (!oAdditionalData.isThreadedComment) {
                    elem.convertToThreadedComment();
                }

                if (elem.asc_getDocumentFlag()) {
                    this.wb.aComments.push(elem);
                } else {
                    oWorksheet.aComments.push(elem);
                }
            }
        }
    };

    function InitSaveManager(wb, isCopyPaste) {
        this.tableIds = {};
        this.sheetIds = {};

        this.aDxfs = [];

        this.slicerCaches = null;
        this.slicerCachesExt = null;

        this.oSharedStrings = {index: 0, strings: {}};
        this.personList = [];
        this.commentUniqueGuids = {};

        this.defNameList;

        this.isCopyPaste = isCopyPaste;
        this.wb = wb;

        this.sharedFormulas = {};
        this.sharedFormulasIndex = 0;

        //option settings for write only active tabs
        this.writeOnlySelectedTabs = null;

        this.prepare();
    }

    InitSaveManager.prototype.prepare = function () {
        var oThis = this;
        var tablesIndex = 1;
        var sheetIndex = 1;

        if (!this.wb) {
            return;
        }

        this.wb.forEach(function(ws) {
            //prepare tables IDs
            var i;
            if (ws.TableParts) {
                for (i = 0; i < ws.TableParts.length; ++i) {
                    oThis.tableIds[ws.TableParts[i].DisplayName] = {id: tablesIndex++, table: ws.TableParts[i]}
                }
            }
            //prepare sheet IDs
            oThis.sheetIds[ws.getId()] = sheetIndex++;

            //break slicers on ext and standard
            var slicerCacheIndex = 0;
            var slicerCacheExtIndex = 0;
            for (i = 0; i < ws.aSlicers.length; ++i) {
                var slicerCache = ws.aSlicers[i].getSlicerCache();
                if (slicerCache) {
                    if (ws.aSlicers[i].isExt()) {
                        if (!oThis.slicerCachesExt) {
                            oThis.slicerCachesExt = {};
                        }
                        oThis.slicerCachesExt[slicerCache.name] = slicerCache;
                        slicerCacheExtIndex++;
                    } else {
                        if (!oThis.slicerCaches) {
                            oThis.slicerCaches = {};
                        }
                        oThis.slicerCaches[slicerCache.name] = slicerCache;
                        slicerCacheIndex++;
                    }
                }
            }
        }, oThis.isCopyPaste);


        //prepare defnames
        var defNameList = this.wb.dependencyFormulas.saveDefName(this.isCopyPaste === false);
        var filterDefName = "_xlnm._FilterDatabase";
        var tempMap = {};
        var printAreaDefName = "Print_Area";
        var prefix = "_xlnm.";

        if(null != defNameList ){
            for(var i = 0; i < defNameList.length; i++){
                if(defNameList[i].Name !== filterDefName) {
                    //TODO временная правка. на открытие может приходить _FilterDatabase. защищаемся от записи двух одинаковых именванных диапазона
                    if(defNameList[i].Name === "_FilterDatabase") {
                        tempMap[defNameList[i].LocalSheetId] = 1;
                    }
                    //на запись добавляем к области печати префикс
                    if(printAreaDefName === defNameList[i].Name && null != defNameList[i].LocalSheetId && true === defNameList[i].isXLNM) {
                        defNameList[i].Name = prefix + defNameList[i].Name;
                    }
                }
            }
        }

        //TODO пишем два одинаковых именованных диапазона с а/ф при открытии книги с а/ф
        //write filters defines name
        //TODO сделать добавление данных именованных диапазонов при добавлении а/ф
        var ws, ref, defNameRef, defName;
        for(var i = 0; i < this.wb.aWorksheets.length; i++) {
            ws = this.wb.aWorksheets[i];
            if(ws && ws.AutoFilter && ws.AutoFilter.Ref && !tempMap[ws.index]) {
                ref = ws.AutoFilter.Ref;
                defNameRef = AscCommon.parserHelp.get3DRef(ws.getName(), ref.getAbsName());
                defName = new Asc.asc_CDefName(filterDefName, defNameRef, ws.index, null, true);
                defNameList.push(defName);
            }
        }
        this.defNameList = defNameList;

        //TODO pivot
    };

    InitSaveManager.prototype.getSlicersCache = function (isExt) {
        return isExt ? this.slicerCachesExt : this.slicerCaches;
    };
    InitSaveManager.prototype.getTableIds = function () {
        return this.tableIds;
    };
    InitSaveManager.prototype.getSheetIds = function () {
        return this.sheetIds;
    };
    InitSaveManager.prototype.getDxfs = function () {
        return this.aDxfs;
    };
    InitSaveManager.prototype.WriteMergeCells = function (ws, func) {
        var i, names, bboxes;
        if (!this.isCopyPaste && ws.mergeManager.initData) {
            names = ws.mergeManager.initData;
            for (i = 0; i < names.length; ++i) {
                func(names[i]);
            }
        } else {
            if (this.isCopyPaste) {
                bboxes = ws.mergeManager.get(this.isCopyPaste).inner.map(function(elem){return elem.bbox});
            } else {
                bboxes = ws.mergeManager.getAll().map(function(elem){return elem.bbox});
            }
            bboxes = getDisjointMerged(this.wb, bboxes);
            for (i = 0; i < bboxes.length; ++i) {
                if (!bboxes[i].isOneCell()) {
                    func(bboxes[i].getName());
                }
            }
        }
    };
    InitSaveManager.prototype._prepeareStyles = function (stylesForWrite) {
        stylesForWrite.init();

        if (!this.wb) {
            return;
        }

        var styles = this.wb.CellStyles.CustomStyles;
        var style = null;
        for (var i = 0; i < styles.length; ++i) {
            style = styles[i];
            if (style.xfs) {
                stylesForWrite.addCellStyle(style);
            }
        }
        stylesForWrite.finalizeCellStyles();
    };

    InitSaveManager.prototype.PrepareSlicerStyles = function (slicerStyles, aDxfs) {
        var styles = new Asc.CT_slicerStyles();
        styles.defaultSlicerStyle = slicerStyles.DefaultStyle;
        for (var name in slicerStyles.CustomStyles) {
            if (slicerStyles.CustomStyles.hasOwnProperty(name)) {
                var slicerStyle = new Asc.CT_slicerStyle();
                slicerStyle.name = name;
                var elems = slicerStyles.CustomStyles[name];
                for (var type in elems) {
                    if (elems.hasOwnProperty(type)) {
                        var styleElement = new Asc.CT_slicerStyleElement();
                        styleElement.type = parseInt(type);
                        styleElement.dxfId = aDxfs.length;
                        aDxfs.push(elems[type]);
                        slicerStyle.slicerStyleElements.push(styleElement);
                    }
                }
                styles.slicerStyle.push(slicerStyle);
            }
        }
        return styles;
    };

    InitSaveManager.prototype.writeCols = function (ws, stylesForWrite, func) {
        var aCols = ws.aCols;
        var oPrevCol = null;
        var nPrevIndexStart = null;
        var nPrevIndex = null;
        var aIndexes = [];
        var oColToWrite;
        var i, length;

        for (i in aCols) {
            aIndexes.push(i - 0);
        }
        aIndexes.sort(AscCommon.fSortAscending);

        var fInitCol = function (col, nMin, nMax) {
            var oRes = {col: col, Max: nMax, Min: nMin, xfsid: null, width: col.width};
            if (null == oRes.width) {
                if (null != ws.oSheetFormatPr.dDefaultColWidth) {
                    oRes.width = ws.oSheetFormatPr.dDefaultColWidth;
                } else {
                    oRes.width = AscCommonExcel.oDefaultMetrics.ColWidthChars;
                }
            }
            if (null != col.xfs) {
                oRes.xfsid = stylesForWrite.add(col.xfs);
            }
            return oRes;
        };

        var oAllCol = null;
        if (null != ws.oAllCol) {
            oAllCol = fInitCol(ws.oAllCol, 0, gc_nMaxCol0);
        }

        for (i = 0, length = aIndexes.length; i < length; ++i) {
            var nIndex = aIndexes[i];
            var col = aCols[nIndex];
            if (null != col) {
                if (false == col.isEmpty()) {
                    if (null != oAllCol && null == nPrevIndex && nIndex > 0) {
                        oAllCol.Min = 1;
                        oAllCol.Max = nIndex;
                        func(oAllCol);
                    }

                    if (null != nPrevIndex && (nPrevIndex + 1 != nIndex || false == oPrevCol.isEqual(col))) {
                        oColToWrite = fInitCol(oPrevCol, nPrevIndexStart + 1, nPrevIndex + 1);
                        func(oColToWrite);
                        nPrevIndexStart = null;
                        if (null != oAllCol && nPrevIndex + 1 != nIndex) {
                            oAllCol.Min = nPrevIndex + 2;
                            oAllCol.Max = nIndex;
                            func(oAllCol);
                        }
                    }
                    oPrevCol = col;
                    nPrevIndex = nIndex;
                    if (null == nPrevIndexStart) {
                        nPrevIndexStart = nPrevIndex;
                    }
                }
            }
        }
        if (null != nPrevIndexStart && null != nPrevIndex && null != oPrevCol) {
            oColToWrite = fInitCol(oPrevCol, nPrevIndexStart + 1, nPrevIndex + 1);
            func(oColToWrite);
        }

        if (null != oAllCol) {
            if (null == nPrevIndex) {
                oAllCol.Min = 1;
                oAllCol.Max = gc_nMaxCol0 + 1;
                func(oAllCol);
            } else if (gc_nMaxCol0 !== nPrevIndex) {
                oAllCol.Min = nPrevIndex + 2;
                oAllCol.Max = gc_nMaxCol0 + 1;
                func(oAllCol);
            }
        }
    };

    InitSaveManager.prototype.WriteTableCustomStyleElements = function (customStyle, func) {
        if(null != customStyle.blankRow)
            func(ETableStyleType.tablestyletypeBlankRow, customStyle.blankRow);
        if(null != customStyle.firstColumn)
            func(ETableStyleType.tablestyletypeFirstColumn, customStyle.firstColumn);
        if(null != customStyle.firstColumnStripe)
            func(ETableStyleType.tablestyletypeFirstColumnStripe, customStyle.firstColumnStripe);
        if(null != customStyle.firstColumnSubheading)
            func(ETableStyleType.tablestyletypeFirstColumnSubheading, customStyle.firstColumnSubheading);
        if(null != customStyle.firstHeaderCell)
            func(ETableStyleType.tablestyletypeFirstHeaderCell, customStyle.firstHeaderCell);
        if(null != customStyle.firstRowStripe)
            func(ETableStyleType.tablestyletypeFirstRowStripe, customStyle.firstRowStripe);
        if(null != customStyle.firstRowSubheading)
            func(ETableStyleType.tablestyletypeFirstRowSubheading, customStyle.firstRowSubheading);
        if(null != customStyle.firstSubtotalColumn)
            func(ETableStyleType.tablestyletypeFirstSubtotalColumn, customStyle.firstSubtotalColumn);
        if(null != customStyle.firstSubtotalRow)
            func(ETableStyleType.tablestyletypeFirstSubtotalRow, customStyle.firstSubtotalRow);
        if(null != customStyle.firstTotalCell)
            func(ETableStyleType.tablestyletypeFirstTotalCell, customStyle.firstTotalCell);
        if(null != customStyle.headerRow)
            func(ETableStyleType.tablestyletypeHeaderRow, customStyle.headerRow);
        if(null != customStyle.lastColumn)
            func(ETableStyleType.tablestyletypeLastColumn, customStyle.lastColumn);
        if(null != customStyle.lastHeaderCell)
            func(ETableStyleType.tablestyletypeLastHeaderCell, customStyle.lastHeaderCell);
        if(null != customStyle.lastTotalCell)
            func(ETableStyleType.tablestyletypeLastTotalCell, customStyle.lastTotalCell);
        if(null != customStyle.pageFieldLabels)
            func(ETableStyleType.tablestyletypePageFieldLabels, customStyle.pageFieldLabels);
        if(null != customStyle.pageFieldValues)
            func(ETableStyleType.tablestyletypePageFieldValues, customStyle.pageFieldValues);
        if(null != customStyle.secondColumnStripe)
            func(ETableStyleType.tablestyletypeSecondColumnStripe, customStyle.secondColumnStripe);
        if(null != customStyle.secondColumnSubheading)
            func(ETableStyleType.tablestyletypeSecondColumnSubheading, customStyle.secondColumnSubheading);
        if(null != customStyle.secondRowStripe)
            func(ETableStyleType.tablestyletypeSecondRowStripe, customStyle.secondRowStripe);
        if(null != customStyle.secondRowSubheading)
            func(ETableStyleType.tablestyletypeSecondRowSubheading, customStyle.secondRowSubheading);
        if(null != customStyle.secondSubtotalColumn)
            func(ETableStyleType.tablestyletypeSecondSubtotalColumn, customStyle.secondSubtotalColumn);
        if(null != customStyle.secondSubtotalRow)
            func(ETableStyleType.tablestyletypeSecondSubtotalRow, customStyle.secondSubtotalRow);
        if(null != customStyle.thirdColumnSubheading)
            func(ETableStyleType.tablestyletypeThirdColumnSubheading, customStyle.thirdColumnSubheading);
        if(null != customStyle.thirdRowSubheading)
            func(ETableStyleType.tablestyletypeThirdRowSubheading, customStyle.thirdRowSubheading);
        if(null != customStyle.thirdSubtotalColumn)
            func(ETableStyleType.tablestyletypeThirdSubtotalColumn, customStyle.thirdSubtotalColumn);
        if(null != customStyle.thirdSubtotalRow)
            func(ETableStyleType.tablestyletypeThirdSubtotalRow, customStyle.thirdSubtotalRow);
        if(null != customStyle.totalRow)
            func(ETableStyleType.tablestyletypeTotalRow, customStyle.totalRow);
        if(null != customStyle.wholeTable)
            func(ETableStyleType.tablestyletypeWholeTable, customStyle.wholeTable);
    };
    InitSaveManager.prototype.PrepareFormulaToWrite = function(cell)
    {
        var parsed = cell.getFormulaParsed();
        var formula;
        var si;
        var ref;
        var type;
        var shared = parsed.getShared();
        var arrayFormula = parsed.getArrayFormulaRef();
        if (shared) {
            var sharedToWrite = this.sharedFormulas[parsed.getIndexNumber()];
            if (!sharedToWrite) {
                sharedToWrite = {saveShared: !shared.ref.isOneCell() && parsed.canSaveShared(), si: {}};
                this.sharedFormulas[parsed.getIndexNumber()] = sharedToWrite;
            }
            if (sharedToWrite.saveShared && shared.ref.contains(cell.nCol, cell.nRow)) {
                type = ECellFormulaType.cellformulatypeShared;
                var rowIndex = Math.floor((cell.nRow - shared.ref.r1) / g_cSharedWriteStreak);
                var row = sharedToWrite.si[rowIndex];
                if (!row) {
                    row = {};
                    sharedToWrite.si[rowIndex] = row;
                }
                var colIndex = Math.floor((cell.nCol - shared.ref.c1) / g_cSharedWriteStreak);
                si = row[colIndex];
                if (undefined === si) {
                    row[colIndex] = si = this.sharedFormulasIndex++;
                    if (!(cell.nRow === shared.base.nRow && cell.nCol === shared.base.nCol)) {
                        cell.processFormula(function(parsed) {
                            formula = parsed.getFormula();
                        });
                    } else {
                        formula = parsed.getFormula();
                    }
                    var r1 = shared.ref.r1 + rowIndex * g_cSharedWriteStreak;
                    var c1 = shared.ref.c1 + colIndex * g_cSharedWriteStreak;
                    ref = new Asc.Range(c1, r1,
                        Math.min(c1 + g_cSharedWriteStreak - 1, shared.ref.c2),
                        Math.min(r1 + g_cSharedWriteStreak - 1, shared.ref.r2));
                }
            } else {
                cell.processFormula(function(parsed) {
                    formula = parsed.getFormula();
                });
            }
        } else if(null !== arrayFormula) {
            //***array-formula***
            var bIsFirstCellArray = parsed.checkFirstCellArray(cell);
            if(bIsFirstCellArray) {
                ref = arrayFormula;
                type = ECellFormulaType.cellformulatypeArray;
                formula = parsed.getFormula();
            } else if(this.isCopyPaste) {
                //если выделена часть формулы, и первая ячейка формулы массива не входит в выделение
                var intersection = arrayFormula.intersection(this.isCopyPaste);
                if(intersection && intersection.r1 === cell.nRow && intersection.c1 === cell.nCol) {
                    ref = arrayFormula;
                    type = ECellFormulaType.cellformulatypeArray;
                    formula = parsed.getFormula();
                }
            }
        } else {
            formula = parsed.getFormula();
        }
        //TODO пока едиственный идентификатор, что внутри есть функция import - importFunctionsRangeLinks
        //обходить каждый раз колстек - не хотелось бы замедлять сохранение, так же как и искать в строке

        //view ->sum(IMPORTRANGE("https://","Sheet1!A1"))+cos(1)
        //file -> IFERROR(__xludf.DUMMYFUNCTION("sum(IMPORTRANGE(""https://"",""Sheet1!A1""))+cos(1)"),123)</f>
        if (formula && parsed && parsed.importFunctionsRangeLinks) {
            formula = "IFERROR(__xludf.DUMMYFUNCTION(\"" + formula.replace(/\"/g,"\"\"") + "\")" + "," + cell.getValue() + ")";
        }
        return {formula: formula, si: si, ref: ref, type: type, ca: parsed.ca};
    };

    function ReadWbComments (wb, contentWorkbookComment, InitOpenManager) {
        var stream = new AscCommon.FT_Stream2(contentWorkbookComment, contentWorkbookComment.length);
        var bwtr = new Binary_WorksheetTableReader(stream, InitOpenManager, wb);
        var bcr = new AscCommon.Binary_CommonReader(stream);
        bcr.Read1(contentWorkbookComment.length, function(t,l){
            return bwtr.ReadCommentDatas(t,l, wb.aComments);
        });
    }

    function WriteWbComments (wb) {
        var res = null;
        if (wb && wb.aComments && wb.aComments.length) {
            var memory = new AscCommon.CMemory();
            var bwtw = new BinaryWorkbookTableWriter(memory, wb);

            var oBinaryStylesTableWriter = new BinaryStylesTableWriter(memory, wb);
            bwtw.oBinaryWorksheetsTableWriter = new BinaryWorksheetsTableWriter(memory, wb, /*this.isCopyPaste*/null, oBinaryStylesTableWriter);

            //var bs = new AscCommon.BinaryCommonWriter(memory)
            //bs.WriteItem(c_oSerWorkbookTypes.Comments, function() {bwtw.WriteComments(wb.aComments);});

            bwtw.WriteComments(wb.aComments);
            res = memory.GetData();
        }
        return res;
    }

    function CT_Workbook(wb) {
        //Members
        this.wb = wb;
        this.sheets = null;
        this.pivotCaches = null;
        this.externalReferences = [];
        this.extLst = null;
        this.slicerCachesIds = [];
        this.newDefinedNames = [];
    }

    function CT_Sheets(wb) {
        this.wb = wb;
        this.sheets = [];
    }

    CT_Sheets.prototype.fromXml = function (reader) {
        var depth = reader.GetDepth();
        while (reader.ReadNextSiblingNode(depth)) {
            if ("sheet" === reader.GetNameNoNS()) {
                var sheet = new AscCommonExcel.CT_Sheet();
                sheet.fromXml(reader);
                this.sheets.push(sheet);
            }
        }
    };

    function CT_Sheet() {
        //Attributes
        this.name = null;
        this.sheetId = null;
        this.id = null;
        this.bHidden = null;
    }

    CT_Sheet.prototype.fromXml = function (reader) {
        this.readAttr(reader);
        reader.ReadTillEnd();
    };
    CT_Sheet.prototype.readAttributes = function (attr, uq) {
        if (attr()) {
            this.parseAttributes(attr());
        }
    };
    CT_Sheet.prototype.readAttr = function (reader) {
        var val;
        while (reader.MoveToNextAttribute()) {
            var name = reader.GetNameNoNS();
            if ("name" === name) {
                this.name = reader.GetValueDecodeXml();
            } else if ("sheetId" === name) {
                this.sheetId = reader.GetValueInt();
            } else if ("id" === name) {
                this.id = reader.GetValueDecodeXml();
            } else if ("state" === name) {
                val = reader.GetValue();
                if ("hidden" === val) {
                    this.bHidden = true;
                } else if ("veryHidden" === val) {
                    this.bHidden = true;
                } else if ("visible" === val) {
                    this.bHidden = false;
                }
            }
        }
    };
    CT_Sheet.prototype.parseAttributes = function (vals, uq) {
        var val;
        val = vals["r:id"];
        if (undefined !== val) {
            this.id = AscCommon.unleakString(uq(val));
        }
    };

    function CT_PivotCaches() {
        this.pivotCaches = [];
    }

    CT_PivotCaches.prototype.fromXml = function (reader) {
        var depth = reader.GetDepth();
        while (reader.ReadNextSiblingNode(depth)) {
            if ("pivotCache" === reader.GetNameNoNS()) {
                var pivotCache = new AscCommonExcel.CT_PivotCache();
                pivotCache.fromXml(reader);
                this.pivotCaches.push(pivotCache);
            }
        }
    };

    function CT_PivotCache() {
        //Attributes
        this.cacheId = null;
        this.id = null;
    }

    CT_PivotCache.prototype.fromXml = function (reader) {
        this.readAttr(reader);
        reader.ReadTillEnd();
    };
    CT_PivotCache.prototype.readAttributes = function (attr, uq) {
        if (attr()) {
            var vals = attr();
            this.parseAttributes(attr(), uq);
        }
    };
    CT_PivotCache.prototype.readAttr = function (reader) {
        while (reader.MoveToNextAttribute()) {
            var name = reader.GetNameNoNS();
            if ("id" === name) {
                this.id = reader.GetValueDecodeXml();
            } else if ("cacheId" === name) {
                this.cacheId = parseInt(reader.GetValue());
            }
        }
    };
    CT_PivotCache.prototype.parseAttributes = function (vals, uq) {
        var val;
        val = vals["cacheId"];
        if (undefined !== val) {
            this.cacheId = val - 0;
        }
        val = vals["r:id"];
        if (undefined !== val) {
            this.id = AscCommon.unleakString(uq(val));
        }
    };

    var prot;
    window['Asc'] = window['Asc'] || {};
    window['AscCommonExcel'] = window['AscCommonExcel'] || {};
    window["Asc"].EBorderStyle = EBorderStyle;
    window["Asc"].EUnderline = EUnderline;
    window["Asc"].ECellAnchorType = ECellAnchorType;
    window["Asc"].EVisibleType = EVisibleType;
    window["Asc"].ECellTypeType = ECellTypeType;
    window["Asc"].ECellFormulaType = ECellFormulaType;
    window["Asc"].EPageOrientation = EPageOrientation;
    window["Asc"].EPageSize = EPageSize;
    window["Asc"].ETotalsRowFunction = ETotalsRowFunction;
    window["Asc"].ESortBy = ESortBy;
    window["Asc"].ECustomFilter = ECustomFilter;

    window['Asc']['EDateTimeGroup'] = window["Asc"].EDateTimeGroup = EDateTimeGroup;
    prot = EDateTimeGroup;
    prot['datetimegroupDay'] = prot.datetimegroupDay;
    prot['datetimegroupHour'] = prot.datetimegroupHour;
    prot['datetimegroupMinute'] = prot.datetimegroupMinute;
    prot['datetimegroupMonth'] = prot.datetimegroupMonth;
    prot['datetimegroupSecond'] = prot.datetimegroupSecond;
    prot['datetimegroupYear'] = prot.datetimegroupYear;

    window["Asc"].ETableStyleType = ETableStyleType;
    window["Asc"].EFontScheme = EFontScheme;

	window['Asc']['EIconSetType'] = window["Asc"].EIconSetType = EIconSetType;
	prot = EIconSetType;
	prot['Arrows3'] = prot.Arrows3;
	prot['Arrows3Gray'] = prot.Arrows3Gray;
	prot['Flags3'] = prot.Flags3;
	prot['Signs3'] = prot.Signs3;
	prot['Symbols3'] = prot.Symbols3;
	prot['Symbols3_2'] = prot.Symbols3_2;
	prot['Traffic3Lights1'] = prot.Traffic3Lights1;
	prot['Traffic3Lights2'] = prot.Traffic3Lights2;
	prot['Arrows4'] = prot.Arrows4;
	prot['Arrows4Gray'] = prot.Arrows4Gray;
	prot['Rating4'] = prot.Rating4;
	prot['RedToBlack4'] = prot.RedToBlack4;
	prot['Traffic4Lights'] = prot.Traffic4Lights;
	prot['Arrows5'] = prot.Arrows5;
	prot['Arrows5Gray'] = prot.Arrows5Gray;
	prot['Quarters5'] = prot.Quarters5;
	prot['Rating5'] = prot.Rating5;
	prot['Triangles3'] = prot.Triangles3;
	prot['Stars3'] = prot.Stars3;
	prot['Boxes5'] = prot.Boxes5;
	prot['NoIcons'] = prot.NoIcons;

    window["Asc"].c_oSer_DrawingType = c_oSer_DrawingType;
    window["Asc"].c_oSer_DrawingPosType = c_oSer_DrawingPosType;

    window['Asc']['c_oAscCFOperator'] = window["AscCommonExcel"].ECfOperator = ECfOperator;
    prot = ECfOperator;
    prot['beginsWith'] = prot.Operator_beginsWith;
    prot['between'] = prot.Operator_between;
    prot['containsText'] = prot.Operator_containsText;
    prot['endsWith'] = prot.Operator_endsWith;
    prot['equal'] = prot.Operator_equal;
    prot['greaterThan'] = prot.Operator_greaterThan;
    prot['greaterThanOrEqual'] = prot.Operator_greaterThanOrEqual;
    prot['lessThan'] = prot.Operator_lessThan;
    prot['lessThanOrEqual'] = prot.Operator_lessThanOrEqual;
    prot['notBetween'] = prot.Operator_notBetween;
    prot['notContains'] = prot.Operator_notContains;
    prot['notEqual'] = prot.Operator_notEqual;

    window['Asc']['c_oAscCFType'] = window["Asc"].ECfType  = ECfType;
    prot = ECfType;
    prot['aboveAverage'] = prot.aboveAverage;
    prot['beginsWith'] = prot.beginsWith;
    prot['cellIs'] = prot.cellIs;
    prot['colorScale'] = prot.colorScale;
    prot['containsBlanks'] = prot.containsBlanks;
    prot['containsErrors'] = prot.containsErrors;
    prot['containsText'] = prot.containsText;
    prot['dataBar'] = prot.dataBar;
    prot['duplicateValues'] = prot.duplicateValues;
    prot['expression'] = prot.expression;
    prot['iconSet'] = prot.iconSet;
    prot['notContainsBlanks'] = prot.notContainsBlanks;
    prot['notContainsErrors'] = prot.notContainsErrors;
    prot['notContainsText'] = prot.notContainsText;
    prot['timePeriod'] = prot.timePeriod;
    prot['top10'] = prot.top10;
    prot['uniqueValues'] = prot.uniqueValues;
    prot['endsWith'] = prot.endsWith;

    window['Asc']['c_oAscCfvoType'] = window["AscCommonExcel"].ECfvoType = ECfvoType;
    prot = ECfvoType;
    prot['Formula'] = prot.Formula;
    prot['Maximum'] = prot.Maximum;
    prot['Minimum'] = prot.Minimum;
    prot['Number'] = prot.Number;
    prot['Percent'] = prot.Percent;
    prot['Percentile'] = prot.Percentile;
    prot['AutoMin'] = prot.AutoMin;
    prot['AutoMax'] = prot.AutoMax;

    window['Asc']['c_oAscTimePeriod'] = window["AscCommonExcel"].ST_TimePeriod = ST_TimePeriod;
    prot = ST_TimePeriod;
    prot['last7Days'] = prot.last7Days;
    prot['lastMonth'] = prot.lastMonth;
    prot['lastWeek'] = prot.lastWeek;
    prot['nextMonth'] = prot.nextMonth;
    prot['nextWeek'] = prot.nextWeek;
    prot['thisMonth'] = prot.thisMonth;
    prot['thisWeek'] = prot.thisWeek;
    prot['today'] = prot.today;
    prot['tomorrow'] = prot.tomorrow;
    prot['yesterday'] = prot.yesterday;

    window['Asc']['c_oAscDataBarAxisPosition'] = window['AscCommonExcel'].EDataBarAxisPosition = EDataBarAxisPosition;
    prot = EDataBarAxisPosition;
    prot['automatic'] = prot.automatic;
    prot['middle'] = prot.middle;
    prot['none'] = prot.none;

    window['Asc']['c_oAscDataBarDirection'] = window["AscCommonExcel"].EDataBarDirection = EDataBarDirection;
    prot = EDataBarDirection;
    prot['context'] = prot.context;
    prot['leftToRight'] = prot.leftToRight;
    prot['rightToLeft'] = prot.rightToLeft;

	window["AscCommonExcel"].XLSB = XLSB;

    window["Asc"].CSlicerStyles = CSlicerStyles;
    window["Asc"].CTableStyles = CTableStyles;
    window["Asc"].CTableStyle = CTableStyle;
    window["Asc"].CTableStyleElement = CTableStyleElement;
    window["Asc"].CTableStyleStripe = CTableStyleStripe;
    window["AscCommonExcel"].BinaryFileReader = BinaryFileReader;
    window["AscCommonExcel"].BinaryFileWriter = BinaryFileWriter;

    window["AscCommonExcel"].BinaryTableWriter = BinaryTableWriter;
    window["AscCommonExcel"].Binary_TableReader = Binary_TableReader;
	window["AscCommonExcel"].OpenFormula = OpenFormula;

    window["Asc"].getBinaryOtherTableGVar = getBinaryOtherTableGVar;
    window["Asc"].ReadDefTableStyles = ReadDefTableStyles;

	window["AscCommonExcel"].BinaryStylesTableWriter = BinaryStylesTableWriter;
	window["AscCommonExcel"].Binary_StylesTableReader = Binary_StylesTableReader;
    window["AscCommonExcel"].StylesForWrite = StylesForWrite;

    window['Asc']['ETotalsRowFunction'] = window['AscCommonExcel'].ETotalsRowFunction = ETotalsRowFunction;
    prot = ETotalsRowFunction;
    prot['totalrowfunctionNone'] = prot.totalrowfunctionNone;
    prot['totalrowfunctionAverage'] = prot.totalrowfunctionAverage;
    prot['totalrowfunctionCount'] = prot.totalrowfunctionCount;
    prot['totalrowfunctionCountNums'] = prot.totalrowfunctionCountNums;
    prot['totalrowfunctionCustom'] = prot.totalrowfunctionCustom;
    prot['totalrowfunctionMax'] = prot.totalrowfunctionMax;
    prot['totalrowfunctionMin'] = prot.totalrowfunctionMin;
    prot['totalrowfunctionStdDev'] = prot.totalrowfunctionStdDev;
    prot['totalrowfunctionSum'] = prot.totalrowfunctionSum;
    prot['totalrowfunctionVar'] = prot.totalrowfunctionVar;

    window["AscCommonExcel"].getSqRefString = getSqRefString;

    window["AscCommonExcel"].InitSaveManager = InitSaveManager;
    window["AscCommonExcel"].InitOpenManager = InitOpenManager;

    window["AscCommonExcel"].CT_Stylesheet = CT_Stylesheet;
    window["AscCommonExcel"].OpenXf = OpenXf;
    window["AscCommonExcel"].DocumentPageSize = DocumentPageSize;
    window["AscCommonExcel"].getSqRefString = getSqRefString;
    window["AscCommonExcel"].g_nNumsMaxId = g_nNumsMaxId;
    window["AscCommonExcel"].XfForWrite = XfForWrite;
    window["AscCommonExcel"].ESortMethod = ESortMethod;
    window["AscCommonExcel"].ST_CellComments = ST_CellComments;
    window["AscCommonExcel"].ST_PrintError = ST_PrintError;
    window["AscCommonExcel"].ST_TableType = ST_TableType;
    window["AscCommonExcel"].ST_PageOrder = ST_PageOrder;
    window["AscCommonExcel"].EActivePane = EActivePane;

    window["AscCommonExcel"].ReadWbComments = ReadWbComments;
    window["AscCommonExcel"].WriteWbComments = WriteWbComments;

    window['AscCommonExcel'].CT_Workbook = CT_Workbook;
    window['AscCommonExcel'].CT_Sheets = CT_Sheets;
    window['AscCommonExcel'].CT_Sheet = CT_Sheet;
    window['AscCommonExcel'].CT_PivotCaches = CT_PivotCaches;
    window['AscCommonExcel'].CT_PivotCache = CT_PivotCache;

    window['AscCommonExcel'].decodeXmlPath = decodeXmlPath;
    window['AscCommonExcel'].encodeXmlPath = encodeXmlPath;

    window['Asc']['c_oAscESheetViewType'] = window['Asc'].c_oAscESheetViewType = window['AscCommonExcel'].ESheetViewType = ESheetViewType;
    prot = ESheetViewType;
    prot['normal'] = prot.normal;
    prot['pageBreakPreview'] = prot.pageBreakPreview;
    prot['pageLayout'] = prot.pageLayout;

    window['Asc']['c_oSerUserProtectedRangeType'] = window['Asc'].c_oSerUserProtectedRangeType = c_oSerUserProtectedRangeType;
    prot = c_oSerUserProtectedRangeType;
    prot['notView'] = prot.notView;
    prot['view'] = prot.view;
    prot['edit'] = prot.edit;


    window['Asc']['EUpdateLinksType'] = window['Asc'].EUpdateLinksType = EUpdateLinksType;
    prot = EUpdateLinksType;
    prot['updatelinksAlways'] = prot.updatelinksAlways;
    prot['updatelinksNever'] = prot.updatelinksNever;
    prot['updatelinksUserSet'] = prot.updatelinksUserSet;

})(window);

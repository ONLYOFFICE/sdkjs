/*
 * (c) Copyright Ascensio System SIA 2010-2023
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

/**
 * User: Ilja.Kirillov
 * Date: 26.10.2016
 * Time: 18:45
 */

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
	function(window, undefined)
{
	function CTableId()
	{
		this.m_aPairs        = null;
		this.m_bTurnOff      = false;
		this.m_oFactoryClass = {};
		this.Id              = null;
		this.isInit          = false;
	}

	CTableId.prototype.checkInit = function()
	{
		return this.isInit;
	};
	CTableId.prototype.init = function()
	{
		this.m_aPairs        = {};
		this.m_bTurnOff      = false;
		this.m_oFactoryClass = {};
		this.Id              = AscCommon.g_oIdCounter.Get_NewId();
		this.Add(this, this.Id);
		this.private_InitFactoryClass();
		this.isInit = true;
	};
	CTableId.prototype.Add = function(Class, Id)
	{
		if (this.m_bTurnOff || !Class)
			return;
		
		Class.Id          = Id;
		this.m_aPairs[Id] = Class;
		
		AscCommon.History.Add(new AscCommon.CChangesTableIdAdd(this, Id, Class));
	};
	CTableId.prototype.TurnOff = function()
	{
		this.m_bTurnOff = true;
	};
	CTableId.prototype.TurnOn = function()
	{
		this.m_bTurnOff = false;
	};
	CTableId.prototype.IsOn = function()
	{
		return (!this.m_bTurnOff);
	};
	/**
	 * Получаем указатель на класс по Id
	 * @param Id
	 * @returns {*}
	 */
	CTableId.prototype.Get_ById = function(Id)
	{
		if ("" === Id)
			return null;

		if (this.m_aPairs[Id])
			return this.m_aPairs[Id];

		return null;
	};
	CTableId.prototype.GetById = function(id)
	{
		return this.GetClass(id);
	}
	CTableId.prototype.GetClass = function(id)
	{
		if (!id || !this.m_aPairs[id])
			return null;

		return this.m_aPairs[id];
	};
	/**
	 * Получаем Id, по классу (вообще, данную функцию лучше не использовать)
	 * @param Class
	 * @returns {*}
	 */
	CTableId.prototype.Get_ByClass = function(Class)
	{
		if (Class.Get_Id)
			return Class.Get_Id();

		if (Class.GetId)
			return Class.GetId();

		return null;
	};
	CTableId.prototype.Get_Id = function()
	{
		return this.Id;
	};
	CTableId.prototype.GetId = function()
	{
		return this.Id;
	};
	CTableId.prototype.Clear = function()
	{
		this.m_aPairs   = {};
		this.m_bTurnOff = false;
		this.Id         = AscCommon.g_oIdCounter.Get_NewId();
		this.Add(this, this.Id);
	};
	CTableId.prototype.private_InitFactoryClass = function()
	{
		this.m_oFactoryClass[AscDFH.historyitem_type_Paragraph]              = AscCommonWord.Paragraph;
		this.m_oFactoryClass[AscDFH.historyitem_type_TextPr]                 = AscCommonWord.ParaTextPr;
		this.m_oFactoryClass[AscDFH.historyitem_type_Hyperlink]              = AscCommonWord.ParaHyperlink;
		this.m_oFactoryClass[AscDFH.historyitem_type_Drawing]                = AscCommonWord.ParaDrawing;
		this.m_oFactoryClass[AscDFH.historyitem_type_Table]                  = AscCommonWord.CTable;
		this.m_oFactoryClass[AscDFH.historyitem_type_TableRow]               = AscCommonWord.CTableRow;
		this.m_oFactoryClass[AscDFH.historyitem_type_TableCell]              = AscCommonWord.CTableCell;
		this.m_oFactoryClass[AscDFH.historyitem_type_DocumentContent]        = AscCommonWord.CDocumentContent;
		this.m_oFactoryClass[AscDFH.historyitem_type_HdrFtr]                 = AscCommonWord.CHeaderFooter;
		this.m_oFactoryClass[AscDFH.historyitem_type_AbstractNum]            = AscCommonWord.CAbstractNum;
		this.m_oFactoryClass[AscDFH.historyitem_type_Comment]                = AscCommon.CComment;
		this.m_oFactoryClass[AscDFH.historyitem_type_Style]                  = AscCommonWord.CStyle;
		this.m_oFactoryClass[AscDFH.historyitem_type_CommentMark]            = AscCommon.ParaComment;
		this.m_oFactoryClass[AscDFH.historyitem_type_ParaRun]                = AscCommonWord.ParaRun;
		this.m_oFactoryClass[AscDFH.historyitem_type_Section]                = AscCommonWord.CSectionPr;
		this.m_oFactoryClass[AscDFH.historyitem_type_Field]                  = AscCommonWord.ParaField;
		this.m_oFactoryClass[AscDFH.historyitem_type_FootEndNote]            = AscCommonWord.CFootEndnote;
		this.m_oFactoryClass[AscDFH.historyitem_type_DefaultShapeDefinition] = AscFormat.DefaultShapeDefinition;
		this.m_oFactoryClass[AscDFH.historyitem_type_CNvPr]                  = AscFormat.CNvPr;
		this.m_oFactoryClass[AscDFH.historyitem_type_NvPr]                   = AscFormat.NvPr;
		this.m_oFactoryClass[AscDFH.historyitem_type_Ph]                     = AscFormat.Ph;
		this.m_oFactoryClass[AscDFH.historyitem_type_UniNvPr]                = AscFormat.UniNvPr;
		this.m_oFactoryClass[AscDFH.historyitem_type_StyleRef]               = AscFormat.StyleRef;
		this.m_oFactoryClass[AscDFH.historyitem_type_FontRef]                = AscFormat.FontRef;
		this.m_oFactoryClass[AscDFH.historyitem_type_Chart]                  = AscFormat.CChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_ChartSpace]             = AscFormat.CChartSpace;
		this.m_oFactoryClass[AscDFH.historyitem_type_Legend]                 = AscFormat.CLegend;
		this.m_oFactoryClass[AscDFH.historyitem_type_Layout]                 = AscFormat.CLayout;
		this.m_oFactoryClass[AscDFH.historyitem_type_LegendEntry]            = AscFormat.CLegendEntry;
		this.m_oFactoryClass[AscDFH.historyitem_type_PivotFmt]               = AscFormat.CPivotFmt;
		this.m_oFactoryClass[AscDFH.historyitem_type_DLbl]                   = AscFormat.CDLbl;
		this.m_oFactoryClass[AscDFH.historyitem_type_Marker]                 = AscFormat.CMarker;
		this.m_oFactoryClass[AscDFH.historyitem_type_PlotArea]               = AscFormat.CPlotArea;
		this.m_oFactoryClass[AscDFH.historyitem_type_NumFmt]                 = AscFormat.CNumFmt;
		this.m_oFactoryClass[AscDFH.historyitem_type_Scaling]                = AscFormat.CScaling;
		this.m_oFactoryClass[AscDFH.historyitem_type_DTable]                 = AscFormat.CDTable;
		this.m_oFactoryClass[AscDFH.historyitem_type_LineChart]              = AscFormat.CLineChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_DLbls]                  = AscFormat.CDLbls;
		this.m_oFactoryClass[AscDFH.historyitem_type_UpDownBars]             = AscFormat.CUpDownBars;
		this.m_oFactoryClass[AscDFH.historyitem_type_BarChart]               = AscFormat.CBarChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_BubbleChart]            = AscFormat.CBubbleChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_DoughnutChart]          = AscFormat.CDoughnutChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_OfPieChart]             = AscFormat.COfPieChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_PieChart]               = AscFormat.CPieChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_RadarChart]             = AscFormat.CRadarChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_ScatterChart]           = AscFormat.CScatterChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_StockChart]             = AscFormat.CStockChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_SurfaceChart]           = AscFormat.CSurfaceChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_BandFmt]                = AscFormat.CBandFmt;
		this.m_oFactoryClass[AscDFH.historyitem_type_AreaChart]              = AscFormat.CAreaChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_ScatterSer]             = AscFormat.CScatterSeries;
		this.m_oFactoryClass[AscDFH.historyitem_type_DPt]                    = AscFormat.CDPt;
		this.m_oFactoryClass[AscDFH.historyitem_type_ErrBars]                = AscFormat.CErrBars;
		this.m_oFactoryClass[AscDFH.historyitem_type_MinusPlus]              = AscFormat.CMinusPlus;
		this.m_oFactoryClass[AscDFH.historyitem_type_NumLit]                 = AscFormat.CNumLit;
		this.m_oFactoryClass[AscDFH.historyitem_type_NumericPoint]           = AscFormat.CNumericPoint;
		this.m_oFactoryClass[AscDFH.historyitem_type_NumRef]                 = AscFormat.CNumRef;
		this.m_oFactoryClass[AscDFH.historyitem_type_TrendLine]              = AscFormat.CTrendLine;
		this.m_oFactoryClass[AscDFH.historyitem_type_Tx]                     = AscFormat.CTx;
		this.m_oFactoryClass[AscDFH.historyitem_type_StrRef]                 = AscFormat.CStrRef;
		this.m_oFactoryClass[AscDFH.historyitem_type_StrCache]               = AscFormat.CStrCache;
		this.m_oFactoryClass[AscDFH.historyitem_type_StrPoint]               = AscFormat.CStringPoint;
		this.m_oFactoryClass[AscDFH.historyitem_type_MultiLvlStrRef]         = AscFormat.CMultiLvlStrRef;
		this.m_oFactoryClass[AscDFH.historyitem_type_MultiLvlStrCache]       = AscFormat.CMultiLvlStrCache;
		this.m_oFactoryClass[AscDFH.historyitem_type_YVal]                   = AscFormat.CYVal;
		this.m_oFactoryClass[AscDFH.historyitem_type_AreaSeries]             = AscFormat.CAreaSeries;
		this.m_oFactoryClass[AscDFH.historyitem_type_Cat]                    = AscFormat.CCat;
		this.m_oFactoryClass[AscDFH.historyitem_type_PictureOptions]         = AscFormat.CPictureOptions;
		this.m_oFactoryClass[AscDFH.historyitem_type_RadarSeries]            = AscFormat.CRadarSeries;
		this.m_oFactoryClass[AscDFH.historyitem_type_BarSeries]              = AscFormat.CBarSeries;
		this.m_oFactoryClass[AscDFH.historyitem_type_LineSeries]             = AscFormat.CLineSeries;
		this.m_oFactoryClass[AscDFH.historyitem_type_PieSeries]              = AscFormat.CPieSeries;
		this.m_oFactoryClass[AscDFH.historyitem_type_SurfaceSeries]          = AscFormat.CSurfaceSeries;
		this.m_oFactoryClass[AscDFH.historyitem_type_BubbleSeries]           = AscFormat.CBubbleSeries;
		this.m_oFactoryClass[AscDFH.historyitem_type_ExternalData]           = AscFormat.CExternalData;
		this.m_oFactoryClass[AscDFH.historyitem_type_PivotSource]            = AscFormat.CPivotSource;
		this.m_oFactoryClass[AscDFH.historyitem_type_Protection]             = AscFormat.CProtection;
		this.m_oFactoryClass[AscDFH.historyitem_type_ChartWall]              = AscFormat.CChartWall;
		this.m_oFactoryClass[AscDFH.historyitem_type_View3d]                 = AscFormat.CView3d;
		this.m_oFactoryClass[AscDFH.historyitem_type_ChartText]              = AscFormat.CChartText;
		this.m_oFactoryClass[AscDFH.historyitem_type_ShapeStyle]             = AscFormat.CShapeStyle;
		this.m_oFactoryClass[AscDFH.historyitem_type_Xfrm]                   = AscFormat.CXfrm;
		this.m_oFactoryClass[AscDFH.historyitem_type_SpPr]                   = AscFormat.CSpPr;
		this.m_oFactoryClass[AscDFH.historyitem_type_UniColor]               = AscFormat.CUniColor;
		this.m_oFactoryClass[AscDFH.historyitem_type_ClrScheme]              = AscFormat.ClrScheme;
		this.m_oFactoryClass[AscDFH.historyitem_type_ClrMap]                 = AscFormat.ClrMap;
		this.m_oFactoryClass[AscDFH.historyitem_type_ExtraClrScheme]         = AscFormat.ExtraClrScheme;
		this.m_oFactoryClass[AscDFH.historyitem_type_FontCollection]         = AscFormat.FontCollection;
		this.m_oFactoryClass[AscDFH.historyitem_type_FontScheme]             = AscFormat.FontScheme;
		this.m_oFactoryClass[AscDFH.historyitem_type_FormatScheme]           = AscFormat.FmtScheme;
		this.m_oFactoryClass[AscDFH.historyitem_type_ThemeElements]          = AscFormat.ThemeElements;
		this.m_oFactoryClass[AscDFH.historyitem_type_HF]                     = AscFormat.HF;
		this.m_oFactoryClass[AscDFH.historyitem_type_BgPr]                   = AscFormat.CBgPr;
		this.m_oFactoryClass[AscDFH.historyitem_type_Bg]                     = AscFormat.CBg;
		this.m_oFactoryClass[AscDFH.historyitem_type_PrintSettings]          = AscFormat.CPrintSettings;
		this.m_oFactoryClass[AscDFH.historyitem_type_HeaderFooterChart]      = AscFormat.CHeaderFooterChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_PageMarginsChart]       = AscFormat.CPageMarginsChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_PageSetup]              = AscFormat.CPageSetup;
		this.m_oFactoryClass[AscDFH.historyitem_type_Shape]                  = AscFormat.CShape;
		this.m_oFactoryClass[AscDFH.historyitem_type_DispUnits]              = AscFormat.CDispUnits;
		this.m_oFactoryClass[AscDFH.historyitem_type_GroupShape]             = AscFormat.CGroupShape;
		this.m_oFactoryClass[AscDFH.historyitem_type_LockedCanvas]           = AscFormat.CLockedCanvas;
		this.m_oFactoryClass[AscDFH.historyitem_type_ImageShape]             = AscFormat.CImageShape;
		this.m_oFactoryClass[AscDFH.historyitem_type_Geometry]               = AscFormat.Geometry;
		this.m_oFactoryClass[AscDFH.historyitem_type_Path]                   = AscFormat.Path;
		this.m_oFactoryClass[AscDFH.historyitem_type_TextBody]               = AscFormat.CTextBody;
		this.m_oFactoryClass[AscDFH.historyitem_type_CatAx]                  = AscFormat.CCatAx;
		this.m_oFactoryClass[AscDFH.historyitem_type_ValAx]                  = AscFormat.CValAx;
		this.m_oFactoryClass[AscDFH.historyitem_type_WrapPolygon]            = AscCommonWord.CWrapPolygon;
		this.m_oFactoryClass[AscDFH.historyitem_type_DateAx]                 = AscFormat.CDateAx;
		this.m_oFactoryClass[AscDFH.historyitem_type_SerAx]                  = AscFormat.CSerAx;
		this.m_oFactoryClass[AscDFH.historyitem_type_Title]                  = AscFormat.CTitle;
		this.m_oFactoryClass[AscDFH.historyitem_type_OleObject]              = AscFormat.COleObject;
        this.m_oFactoryClass[AscDFH.historyitem_type_Cnx]                    = AscFormat.CConnectionShape;
		this.m_oFactoryClass[AscDFH.historyitem_type_DrawingContent]         = AscFormat.CDrawingDocContent;
		this.m_oFactoryClass[AscDFH.historyitem_type_Math]                   = AscCommonWord.ParaMath;
		this.m_oFactoryClass[AscDFH.historyitem_type_MathContent]            = AscCommonWord.CMathContent;
		this.m_oFactoryClass[AscDFH.historyitem_type_acc]                    = AscCommonWord.CAccent;
		this.m_oFactoryClass[AscDFH.historyitem_type_bar]                    = AscCommonWord.CBar;
		this.m_oFactoryClass[AscDFH.historyitem_type_box]                    = AscCommonWord.CBox;
		this.m_oFactoryClass[AscDFH.historyitem_type_borderBox]              = AscCommonWord.CBorderBox;
		this.m_oFactoryClass[AscDFH.historyitem_type_delimiter]              = AscCommonWord.CDelimiter;
		this.m_oFactoryClass[AscDFH.historyitem_type_eqArr]                  = AscCommonWord.CEqArray;
		this.m_oFactoryClass[AscDFH.historyitem_type_frac]                   = AscCommonWord.CFraction;
		this.m_oFactoryClass[AscDFH.historyitem_type_mathFunc]               = AscCommonWord.CMathFunc;
		this.m_oFactoryClass[AscDFH.historyitem_type_groupChr]               = AscCommonWord.CGroupCharacter;
		this.m_oFactoryClass[AscDFH.historyitem_type_lim]                    = AscCommonWord.CLimit;
		this.m_oFactoryClass[AscDFH.historyitem_type_matrix]                 = AscCommonWord.CMathMatrix;
		this.m_oFactoryClass[AscDFH.historyitem_type_nary]                   = AscCommonWord.CNary;
		this.m_oFactoryClass[AscDFH.historyitem_type_phant]                  = AscCommonWord.CPhantom;
		this.m_oFactoryClass[AscDFH.historyitem_type_rad]                    = AscCommonWord.CRadical;
		this.m_oFactoryClass[AscDFH.historyitem_type_deg_subsup]             = AscCommonWord.CDegreeSubSup;
		this.m_oFactoryClass[AscDFH.historyitem_type_deg]                    = AscCommonWord.CDegree;
		this.m_oFactoryClass[AscDFH.historyitem_type_BlockLevelSdt]          = AscCommonWord.CBlockLevelSdt;
		this.m_oFactoryClass[AscDFH.historyitem_type_InlineLevelSdt]         = AscCommonWord.CInlineLevelSdt;
		this.m_oFactoryClass[AscDFH.historyitem_type_ParaBookmark]           = AscCommonWord.CParagraphBookmark;
		this.m_oFactoryClass[AscDFH.historyitem_type_Num]                    = AscCommonWord.CNum;
		this.m_oFactoryClass[AscDFH.historyitem_type_PresentationField]      = AscCommonWord.CPresentationField;
		this.m_oFactoryClass[AscDFH.historyitem_type_RelSizeAnchor]      	 = AscFormat.CRelSizeAnchor;
		this.m_oFactoryClass[AscDFH.historyitem_type_AbsSizeAnchor]      	 = AscFormat.CAbsSizeAnchor;
		this.m_oFactoryClass[AscDFH.historyitem_type_ParaRevisionMove]       = AscCommon.CParaRevisionMove;
		this.m_oFactoryClass[AscDFH.historyitem_type_RunRevisionMove]        = AscWord.CRunRevisionMove;
		this.m_oFactoryClass[AscDFH.historyitem_type_DocPart]                = AscCommon.CDocPart;
		this.m_oFactoryClass[AscDFH.historyitem_type_SlicerView]             = AscFormat.CSlicer;
		this.m_oFactoryClass[AscDFH.historyitem_type_ChartStyle]             = AscFormat.CChartStyle;
		this.m_oFactoryClass[AscDFH.historyitem_type_MarkerLayout]           = AscFormat.CMarkerLayout;
		this.m_oFactoryClass[AscDFH.historyitem_type_ChartStyleEntry]        = AscFormat.CStyleEntry;
		this.m_oFactoryClass[AscDFH.historyitem_type_PrSet             ]     = AscFormat.PrSet;
		this.m_oFactoryClass[AscDFH.historyitem_type_CCommonDataList   ]     = AscFormat.CCommonDataList;
		this.m_oFactoryClass[AscDFH.historyitem_type_Point             ]     = AscFormat.Point;
		this.m_oFactoryClass[AscDFH.historyitem_type_PtLst             ]     = AscFormat.PtLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_DataModel         ]     = AscFormat.DataModel;
		this.m_oFactoryClass[AscDFH.historyitem_type_CxnLst            ]     = AscFormat.CxnLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_ExtLst            ]     = AscFormat.ExtLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_BgFormat          ]     = AscFormat.BgFormat;
		this.m_oFactoryClass[AscDFH.historyitem_type_Whole             ]     = AscFormat.Whole;
		this.m_oFactoryClass[AscDFH.historyitem_type_Cxn               ]     = AscFormat.Cxn;
		this.m_oFactoryClass[AscDFH.historyitem_type_Ext               ]     = AscFormat.Ext;
		this.m_oFactoryClass[AscDFH.historyitem_type_LayoutDef         ]     = AscFormat.LayoutDef;
		this.m_oFactoryClass[AscDFH.historyitem_type_CatLst            ]     = AscFormat.CatLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_SCat              ]     = AscFormat.SCat;
		this.m_oFactoryClass[AscDFH.historyitem_type_ClrData           ]     = AscFormat.ClrData;
		this.m_oFactoryClass[AscDFH.historyitem_type_Desc              ]     = AscFormat.Desc;
		this.m_oFactoryClass[AscDFH.historyitem_type_LayoutNode        ]     = AscFormat.LayoutNode;
		this.m_oFactoryClass[AscDFH.historyitem_type_Alg               ]     = AscFormat.Alg;
		this.m_oFactoryClass[AscDFH.historyitem_type_Param             ]     = AscFormat.Param;
		this.m_oFactoryClass[AscDFH.historyitem_type_Choose            ]     = AscFormat.Choose;
		this.m_oFactoryClass[AscDFH.historyitem_type_IteratorAttributes]     = AscFormat.IteratorAttributes;
		this.m_oFactoryClass[AscDFH.historyitem_type_Else              ]     = AscFormat.Else;
		this.m_oFactoryClass[AscDFH.historyitem_type_AxisType          ]     = AscFormat.AxisType;
		this.m_oFactoryClass[AscDFH.historyitem_type_If                ]     = AscFormat.If;
		this.m_oFactoryClass[AscDFH.historyitem_type_ElementType       ]     = AscFormat.ElementType;
		this.m_oFactoryClass[AscDFH.historyitem_type_ConstrLst         ]     = AscFormat.ConstrLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_Constr            ]     = AscFormat.Constr;
		this.m_oFactoryClass[AscDFH.historyitem_type_PresOf            ]     = AscFormat.PresOf;
		this.m_oFactoryClass[AscDFH.historyitem_type_RuleLst           ]     = AscFormat.RuleLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_Rule              ]     = AscFormat.Rule;
		this.m_oFactoryClass[AscDFH.historyitem_type_SShape            ]     = AscFormat.SShape;
		this.m_oFactoryClass[AscDFH.historyitem_type_AdjLst            ]     = AscFormat.AdjLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_Adj               ]     = AscFormat.Adj;
		this.m_oFactoryClass[AscDFH.historyitem_type_AnimLvl           ]     = AscFormat.AnimLvl;
		this.m_oFactoryClass[AscDFH.historyitem_type_AnimOne           ]     = AscFormat.AnimOne;
		this.m_oFactoryClass[AscDFH.historyitem_type_BulletEnabled     ]     = AscFormat.BulletEnabled;
		this.m_oFactoryClass[AscDFH.historyitem_type_ChMax             ]     = AscFormat.ChMax;
		this.m_oFactoryClass[AscDFH.historyitem_type_ChPref            ]     = AscFormat.ChPref;
		this.m_oFactoryClass[AscDFH.historyitem_type_DiagramDirection  ]     = AscFormat.DiagramDirection;
		this.m_oFactoryClass[AscDFH.historyitem_type_DiagramTitle      ]     = AscFormat.DiagramTitle;
		this.m_oFactoryClass[AscDFH.historyitem_type_LayoutDefHdrLst   ]     = AscFormat.LayoutDefHdrLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_LayoutDefHdr      ]     = AscFormat.LayoutDefHdr;
		this.m_oFactoryClass[AscDFH.historyitem_type_RelIds            ]     = AscFormat.RelIds;
		this.m_oFactoryClass[AscDFH.historyitem_type_VarLst            ]     = AscFormat.VarLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_ColorsDef         ]     = AscFormat.ColorsDef;
		this.m_oFactoryClass[AscDFH.historyitem_type_ColorDefStyleLbl  ]     = AscFormat.ColorDefStyleLbl;
		this.m_oFactoryClass[AscDFH.historyitem_type_ClrLst            ]     = AscFormat.ClrLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_EffectClrLst      ]     = AscFormat.EffectClrLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_FillClrLst        ]     = AscFormat.FillClrLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_LinClrLst         ]     = AscFormat.LinClrLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_TxEffectClrLst    ]     = AscFormat.TxEffectClrLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_TxFillClrLst      ]     = AscFormat.TxFillClrLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_TxLinClrLst       ]     = AscFormat.TxLinClrLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_ColorsDefHdr      ]     = AscFormat.ColorsDefHdr;
		this.m_oFactoryClass[AscDFH.historyitem_type_ColorsDefHdrLst   ]     = AscFormat.ColorsDefHdrLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_StyleDef          ]     = AscFormat.StyleDef;
		this.m_oFactoryClass[AscDFH.historyitem_type_Scene3d           ]     = AscFormat.Scene3d;
		this.m_oFactoryClass[AscDFH.historyitem_type_StyleDefStyleLbl  ]     = AscFormat.StyleDefStyleLbl;
		this.m_oFactoryClass[AscDFH.historyitem_type_Scene3d           ]     = AscFormat.Scene3d;
		this.m_oFactoryClass[AscDFH.historyitem_type_Backdrop          ]     = AscFormat.Backdrop;
		this.m_oFactoryClass[AscDFH.historyitem_type_BackdropNorm      ]     = AscFormat.BackdropNorm;
		this.m_oFactoryClass[AscDFH.historyitem_type_BackdropUp        ]     = AscFormat.BackdropUp;
		this.m_oFactoryClass[AscDFH.historyitem_type_Camera            ]     = AscFormat.Camera;
		this.m_oFactoryClass[AscDFH.historyitem_type_Rot               ]     = AscFormat.Rot;
		this.m_oFactoryClass[AscDFH.historyitem_type_LightRig          ]     = AscFormat.LightRig;
		this.m_oFactoryClass[AscDFH.historyitem_type_Sp3d              ]     = AscFormat.Sp3d;
		this.m_oFactoryClass[AscDFH.historyitem_type_Bevel             ]     = AscFormat.Bevel;
		this.m_oFactoryClass[AscDFH.historyitem_type_BevelB            ]     = AscFormat.BevelB;
		this.m_oFactoryClass[AscDFH.historyitem_type_BevelT            ]     = AscFormat.BevelT;
		this.m_oFactoryClass[AscDFH.historyitem_type_TxPr              ]     = AscFormat.TxPr;
		this.m_oFactoryClass[AscDFH.historyitem_type_FlatTx            ]     = AscFormat.FlatTx;
		this.m_oFactoryClass[AscDFH.historyitem_type_StyleDefHdrLst    ]     = AscFormat.StyleDefHdrLst;
		this.m_oFactoryClass[AscDFH.historyitem_type_StyleDefHdr       ]     = AscFormat.StyleDefHdr;
		this.m_oFactoryClass[AscDFH.historyitem_type_BackdropAnchor    ]     = AscFormat.BackdropAnchor;
		this.m_oFactoryClass[AscDFH.historyitem_type_StyleData         ]     = AscFormat.StyleData;
		this.m_oFactoryClass[AscDFH.historyitem_type_SampData          ]     = AscFormat.SampData;
		this.m_oFactoryClass[AscDFH.historyitem_type_ForEach           ]     = AscFormat.ForEach;
		this.m_oFactoryClass[AscDFH.historyitem_type_ResizeHandles     ]     = AscFormat.ResizeHandles;
		this.m_oFactoryClass[AscDFH.historyitem_type_OrgChart          ]     = AscFormat.OrgChart;
		this.m_oFactoryClass[AscDFH.historyitem_type_HierBranch        ]     = AscFormat.HierBranch;
		this.m_oFactoryClass[AscDFH.historyitem_type_ParameterVal      ]     = AscFormat.ParameterVal;
		this.m_oFactoryClass[AscDFH.historyitem_type_Coordinate        ]     = AscFormat.Coordinate;
		this.m_oFactoryClass[AscDFH.historyitem_type_ExtrusionClr      ]     = AscFormat.ExtrusionClr;
		this.m_oFactoryClass[AscDFH.historyitem_type_ContourClr        ]     = AscFormat.ContourClr;
		this.m_oFactoryClass[AscDFH.historyitem_type_SmartArt          ]     = AscFormat.SmartArt;
		this.m_oFactoryClass[AscDFH.historyitem_type_CCommonDataClrList]     = AscFormat.CCommonDataClrList;
		this.m_oFactoryClass[AscDFH.historyitem_type_BuNone            ]     = AscFormat.BuNone;
		this.m_oFactoryClass[AscDFH.historyitem_type_SmartArtDrawing   ]     = AscFormat.Drawing;
		this.m_oFactoryClass[AscDFH.historyitem_type_DiagramData       ]     = AscFormat.DiagramData;
		this.m_oFactoryClass[AscDFH.historyitem_type_FunctionValue     ]     = AscFormat.FunctionValue;
		this.m_oFactoryClass[AscDFH.historyitem_type_ShapeSmartArtInfo ]     = AscFormat.ShapeSmartArtInfo;
		this.m_oFactoryClass[AscDFH.historyitem_type_SmartArtTree      ]     = AscFormat.SmartArtTree;
		this.m_oFactoryClass[AscDFH.historyitem_type_SmartArtNode      ]     = AscFormat.SmartArtNode;
		this.m_oFactoryClass[AscDFH.historyitem_type_SmartArtNodeData  ]     = AscFormat.SmartArtNodeData;
		this.m_oFactoryClass[AscDFH.historyitem_type_BuBlip            ]     = AscFormat.CBuBlip;

		if (window['AscCommonSlide'])
		{
			this.m_oFactoryClass[AscDFH.historyitem_type_Slide]               = AscCommonSlide.Slide;
			this.m_oFactoryClass[AscDFH.historyitem_type_SlideLayout]         = AscCommonSlide.SlideLayout;
			this.m_oFactoryClass[AscDFH.historyitem_type_SlideMaster]         = AscCommonSlide.MasterSlide;
			this.m_oFactoryClass[AscDFH.historyitem_type_SlideComments]       = AscCommonSlide.SlideComments;
			this.m_oFactoryClass[AscDFH.historyitem_type_PropLocker]          = AscCommonSlide.PropLocker;
			this.m_oFactoryClass[AscDFH.historyitem_type_NotesMaster]         = AscCommonSlide.CNotesMaster;
			this.m_oFactoryClass[AscDFH.historyitem_type_Notes]               = AscCommonSlide.CNotes;
			this.m_oFactoryClass[AscDFH.historyitem_type_PresentationSection] = AscCommonSlide.CPrSection;
			this.m_oFactoryClass[AscDFH.historyitem_type_SldSz]               = AscCommonSlide.CSlideSize;

			//Classes for animation
			this.m_oFactoryClass[AscDFH.historyitem_type_EmptyObject]         = AscFormat.CEmptyObject;
			this.m_oFactoryClass[AscDFH.historyitem_type_Timing]              = AscFormat.CTiming;
			this.m_oFactoryClass[AscDFH.historyitem_type_CommonTimingList]    = AscFormat.CCommonTimingList;
			this.m_oFactoryClass[AscDFH.historyitem_type_AttrNameLst]         = AscFormat.CAttrNameLst;
			this.m_oFactoryClass[AscDFH.historyitem_type_BldLst]              = AscFormat.CBldLst;
			this.m_oFactoryClass[AscDFH.historyitem_type_CondLst]             = AscFormat.CCondLst;
			this.m_oFactoryClass[AscDFH.historyitem_type_ChildTnLst]          = AscFormat.CChildTnLst;
			this.m_oFactoryClass[AscDFH.historyitem_type_TmplLst]             = AscFormat.CTmplLst;
			this.m_oFactoryClass[AscDFH.historyitem_type_TnLst]               = AscFormat.CTnLst;
			this.m_oFactoryClass[AscDFH.historyitem_type_TavLst]              = AscFormat.CTavLst;
			this.m_oFactoryClass[AscDFH.historyitem_type_ObjectTarget]        = AscFormat.CObjectTarget;
			this.m_oFactoryClass[AscDFH.historyitem_type_BldBase]             = AscFormat.CBldBase;
			this.m_oFactoryClass[AscDFH.historyitem_type_BldDgm]              = AscFormat.CBldDgm;
			this.m_oFactoryClass[AscDFH.historyitem_type_BldGraphic]          = AscFormat.CBldGraphic;
			this.m_oFactoryClass[AscDFH.historyitem_type_BldOleChart]         = AscFormat.CBldOleChart;
			this.m_oFactoryClass[AscDFH.historyitem_type_BldP]                = AscFormat.CBldP;
			this.m_oFactoryClass[AscDFH.historyitem_type_BldSub]              = AscFormat.CBldSub;
			this.m_oFactoryClass[AscDFH.historyitem_type_DirTransition]       = AscFormat.CDirTransition;
			this.m_oFactoryClass[AscDFH.historyitem_type_OptBlackTransition]  = AscFormat.COptionalBlackTransition;
			this.m_oFactoryClass[AscDFH.historyitem_type_GraphicEl]           = AscFormat.CGraphicEl;
			this.m_oFactoryClass[AscDFH.historyitem_type_IndexRg]             = AscFormat.CIndexRg;
			this.m_oFactoryClass[AscDFH.historyitem_type_Tmpl]                = AscFormat.CTmpl;
			this.m_oFactoryClass[AscDFH.historyitem_type_Anim]                = AscFormat.CAnim;
			this.m_oFactoryClass[AscDFH.historyitem_type_CBhvr]               = AscFormat.CCBhvr;
			this.m_oFactoryClass[AscDFH.historyitem_type_CTn]                 = AscFormat.CCTn;
			this.m_oFactoryClass[AscDFH.historyitem_type_Cond]                = AscFormat.CCond;
			this.m_oFactoryClass[AscDFH.historyitem_type_TgtEl]               = AscFormat.CTgtEl;
			this.m_oFactoryClass[AscDFH.historyitem_type_SndTgt]              = AscFormat.CSndTgt;
			this.m_oFactoryClass[AscDFH.historyitem_type_SpTgt]               = AscFormat.CSpTgt;
			this.m_oFactoryClass[AscDFH.historyitem_type_IterateData]         = AscFormat.CIterateData;
			this.m_oFactoryClass[AscDFH.historyitem_type_Tav]                 = AscFormat.CTav;
			this.m_oFactoryClass[AscDFH.historyitem_type_AnimVariant]         = AscFormat.CAnimVariant;
			this.m_oFactoryClass[AscDFH.historyitem_type_AnimClr]             = AscFormat.CAnimClr;
			this.m_oFactoryClass[AscDFH.historyitem_type_AnimEffect]          = AscFormat.CAnimEffect;
			this.m_oFactoryClass[AscDFH.historyitem_type_AnimMotion]          = AscFormat.CAnimMotion;
			this.m_oFactoryClass[AscDFH.historyitem_type_AnimRot]             = AscFormat.CAnimRot;
			this.m_oFactoryClass[AscDFH.historyitem_type_AnimScale]           = AscFormat.CAnimScale;
			this.m_oFactoryClass[AscDFH.historyitem_type_Audio]               = AscFormat.CAudio;
			this.m_oFactoryClass[AscDFH.historyitem_type_CMediaNode]          = AscFormat.CCMediaNode;
			this.m_oFactoryClass[AscDFH.historyitem_type_Cmd]                 = AscFormat.CCmd;
			this.m_oFactoryClass[AscDFH.historyitem_type_TimeNodeContainer]   = AscFormat.CTimeNodeContainer;
			this.m_oFactoryClass[AscDFH.historyitem_type_Par]                 = AscFormat.CPar;
			this.m_oFactoryClass[AscDFH.historyitem_type_Excl]                = AscFormat.CExcl;
			this.m_oFactoryClass[AscDFH.historyitem_type_Seq]                 = AscFormat.CSeq;
			this.m_oFactoryClass[AscDFH.historyitem_type_Set]                 = AscFormat.CSet;
			this.m_oFactoryClass[AscDFH.historyitem_type_Video]               = AscFormat.CVideo;
			this.m_oFactoryClass[AscDFH.historyitem_type_OleChartEl]          = AscFormat.COleChartEl;
			this.m_oFactoryClass[AscDFH.historyitem_type_TlPoint]             = AscFormat.CTLPoint;
			this.m_oFactoryClass[AscDFH.historyitem_type_SndAc]               = AscFormat.CSndAc;
			this.m_oFactoryClass[AscDFH.historyitem_type_StSnd]               = AscFormat.CStSnd;
			this.m_oFactoryClass[AscDFH.historyitem_type_TxEl]                = AscFormat.CTxEl;
			this.m_oFactoryClass[AscDFH.historyitem_type_Wheel]               = AscFormat.CWheel;
			this.m_oFactoryClass[AscDFH.historyitem_type_AttrName]            = AscFormat.CAttrName;
		}

		this.m_oFactoryClass[AscDFH.historyitem_type_Theme]                  = AscFormat.CTheme;
		this.m_oFactoryClass[AscDFH.historyitem_type_GraphicFrame]           = AscFormat.CGraphicFrame;

		if (window['AscCommonExcel'])
		{
			this.m_oFactoryClass[AscDFH.historyitem_type_Sparkline]            = AscCommonExcel.sparklineGroup;
			this.m_oFactoryClass[AscDFH.historyitem_type_PivotTableDefinition] = Asc.CT_pivotTableDefinition;
			this.m_oFactoryClass[AscDFH.historyitem_type_PivotWorksheetSource] = Asc.CT_WorksheetSource;
			this.m_oFactoryClass[AscDFH.historyitem_type_NamedSheetView]       = Asc.CT_NamedSheetView;
			this.m_oFactoryClass[AscDFH.historyitem_type_DataValidation]       = AscCommonExcel.CDataValidation;
			this.m_oFactoryClass[AscDFH.historyitem_type_OleSizeSelection  ]   = AscCommonExcel.OleSizeSelectionRange;
			this.m_oFactoryClass[AscDFH.historyitem_type_ViewPr]               = AscFormat.CViewPr;
			this.m_oFactoryClass[AscDFH.historyitem_type_CommonViewPr]         = AscFormat.CCommonViewPr;
			this.m_oFactoryClass[AscDFH.historyitem_type_CSldViewPr]           = AscFormat.CCSldViewPr;
			this.m_oFactoryClass[AscDFH.historyitem_type_CViewPr]              = AscFormat.CCViewPr;
			this.m_oFactoryClass[AscDFH.historyitem_type_ViewPrScale]          = AscFormat.CViewPrScale;
			this.m_oFactoryClass[AscDFH.historyitem_type_ViewPrGuide]          = AscFormat.CViewPrGuide;

		}

		this.m_oFactoryClass[AscDFH.historyitem_type_DocumentMacros] = AscCommon.CDocumentMacros;
		
		this.InitOFormClasses();
	};
	CTableId.prototype.GetClassFromFactory = function(nType)
	{
		if (this.m_oFactoryClass[nType])
			return new this.m_oFactoryClass[nType]();

		return null;
	};
	CTableId.prototype.Refresh_RecalcData = function(Data)
	{
	};
	CTableId.prototype.InitOFormClasses = function()
	{
		if (AscCommon.IsSupportOFormFeature())
		{
			this.m_oFactoryClass[AscDFH.historyitem_type_OForm_UserMaster]  = AscOForm.CUserMaster;
			this.m_oFactoryClass[AscDFH.historyitem_type_OForm_User]        = AscOForm.CUser;
			this.m_oFactoryClass[AscDFH.historyitem_type_OForm_FieldMaster] = AscOForm.CFieldMaster;
			this.m_oFactoryClass[AscDFH.historyitem_type_OForm_Document]    = AscOForm.CDocument;
			this.m_oFactoryClass[AscDFH.historyitem_type_OForm_FieldGroup]  = AscOForm.CFieldGroup;
		}
		else
		{
			delete this.m_oFactoryClass[AscDFH.historyitem_type_OForm_UserMaster];
			delete this.m_oFactoryClass[AscDFH.historyitem_type_OForm_User];
			delete this.m_oFactoryClass[AscDFH.historyitem_type_OForm_FieldMaster];
			delete this.m_oFactoryClass[AscDFH.historyitem_type_OForm_Document];
			delete this.m_oFactoryClass[AscDFH.historyitem_type_OForm_FieldGroup];
		}
	};
	//-----------------------------------------------------------------------------------
	// Функции для работы с совместным редактирования
	//-----------------------------------------------------------------------------------
	CTableId.prototype.Unlock = function(Data)
	{
		// Ничего не делаем
	};

	window["AscCommon"].g_oTableId = new CTableId();
	window["AscCommon"].CTableId = CTableId;
})(window);

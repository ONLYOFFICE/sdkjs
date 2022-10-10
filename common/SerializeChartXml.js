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

(function(window, undefined) {
	let CChartSpace = window['AscFormat'].CChartSpace;
	let CPivotSource = window['AscFormat'].CPivotSource;
	let CProtection = window['AscFormat'].CProtection;
	let CPrintSettings = window['AscFormat'].CPrintSettings;
	let CHeaderFooterChart = window['AscFormat'].CHeaderFooterChart;
	let CPageMarginsChart = window['AscFormat'].CPageMarginsChart;
	let CPageSetup = window['AscFormat'].CPageSetup;
	let CChart = window['AscFormat'].CChart;
	let CLayout = window['AscFormat'].CLayout;
	let CTitle = window['AscFormat'].CTitle;
	let CChartText = window['AscFormat'].CChartText;
	let CStrRef = window['AscFormat'].CStrRef;
	let CStrCache = window['AscFormat'].CStrCache;
	let CStringPoint = window['AscFormat'].CStringPoint;
	let CPivotFmt = window['AscFormat'].CPivotFmt;
	let CMarker = window['AscFormat'].CMarker;
	let CDLbl = window['AscFormat'].CDLbl;
	let CView3d = window['AscFormat'].CView3d;
	let CChartWall = window['AscFormat'].CChartWall;
	let CPictureOptions = window['AscFormat'].CPictureOptions;
	let CLegend = window['AscFormat'].CLegend;
	let CLegendEntry = window['AscFormat'].CLegendEntry;
	let CPlotArea = window['AscFormat'].CPlotArea;
	let CDTable = window['AscFormat'].CDTable;
	let CAreaChart = window['AscFormat'].CAreaChart;
	let CBarChart = window['AscFormat'].CBarChart;
	let CBubbleChart = window['AscFormat'].CBubbleChart;
	let CDoughnutChart = window['AscFormat'].CDoughnutChart;
	let CLineChart = window['AscFormat'].CLineChart;
	let COfPieChart = window['AscFormat'].COfPieChart;
	let CPieChart = window['AscFormat'].CPieChart;
	let CRadarChart = window['AscFormat'].CRadarChart;
	let CScatterChart = window['AscFormat'].CScatterChart;
	let CStockChart = window['AscFormat'].CStockChart;
	let CSurfaceChart = window['AscFormat'].CSurfaceChart;
	let CAreaSeries = window['AscFormat'].CAreaSeries;
	let CBarSeries = window['AscFormat'].CBarSeries;
	let CBubbleSeries = window['AscFormat'].CBubbleSeries;
	let CLineSeries = window['AscFormat'].CLineSeries;
	let CPieSeries = window['AscFormat'].CPieSeries;
	let CRadarSeries = window['AscFormat'].CRadarSeries;
	let CScatterSeries = window['AscFormat'].CScatterSeries;
	let CSurfaceSeries = window['AscFormat'].CSurfaceSeries;
	let CTx = window['AscFormat'].CTx;
	let CDPt = window['AscFormat'].CDPt;
	let CDLbls = window['AscFormat'].CDLbls;
	let CNumFmt = window['AscFormat'].CNumFmt;
	let CTrendLine = window['AscFormat'].CTrendLine;
	let CErrBars = window['AscFormat'].CErrBars;
	let CMinusPlus = window['AscFormat'].CMinusPlus;
	let CNumLit = window['AscFormat'].CNumLit;
	let CNumRef = window['AscFormat'].CNumRef;
	let CNumericPoint = window['AscFormat'].CNumericPoint;
	let CCat = window['AscFormat'].CCat;
	let CYVal = window['AscFormat'].CYVal;
	let CMultiLvlStrRef = window['AscFormat'].CMultiLvlStrRef;
	let CMultiLvlStrCache = window['AscFormat'].CMultiLvlStrCache;
	let CUpDownBars = window['AscFormat'].CUpDownBars;
	let CBandFmt = window['AscFormat'].CBandFmt;
	let CValAx = window['AscFormat'].CValAx;
	let CCatAx = window['AscFormat'].CCatAx;
	let CDateAx = window['AscFormat'].CDateAx;
	let CSerAx = window['AscFormat'].CSerAx;
	let CScaling = window['AscFormat'].CScaling;
	let CDispUnits = window['AscFormat'].CDispUnits;
	let CChartStyle = window['AscFormat'].CChartStyle;
	let CStyleEntry = window['AscFormat'].CStyleEntry;
	let CMarkerLayout = window['AscFormat'].CMarkerLayout;
	let CChartColors = window['AscFormat'].CChartColors;

	let CT_Bool = window['AscCommon'].CT_Bool;
	let CT_String = window['AscCommon'].CT_String;
	let CT_Int = window['AscCommon'].CT_Int;
	let CT_UInt = window['AscCommon'].CT_UInt;
	let CT_Double = window['AscCommon'].CT_Double;

	function readSpPrPlain(reader) {
		//todo CSpPr
		let elem = new AscFormat.CSpPr();
		elem.fromXml(reader);
		return elem;
	}

	function readSpPr(reader, parent) {
		let elem = readSpPrPlain(reader);
		elem.setParent(parent);
		parent.setSpPr(elem);
	}

	function readTxPrPlain(reader) {
		//todo CTextBody
		let elem = new AscFormat.CTextBody();
		elem.fromXml(reader);
		if(!elem.content) {
			elem.setContent(new AscFormat.CDrawingDocContent(elem, reader.context.DrawingDocument, 0, 0, 0, 0, 0, 0, true));
		}
		return elem;
	}

	function readTxPr(reader, parent) {
		let elem = readTxPrPlain(reader);
		elem.setParent(parent);
		parent.setTxPr(elem);
	}

	function readStyleRef(reader) {
		let oElem = new AscFormat.StyleRef();
		oElem.fromXml(reader);
		return oElem;
	}
	function readFontRef(reader) {
		let oElem = new AscFormat.FontRef();
		oElem.fromXml(reader);
		return oElem;
	}
	function readRPr(reader) {
		//todo
		return new CTextPr();
	}
	function readBodyPr(reader) {
		let oElem = new AscFormat.CBodyPr();
		oElem.fromXml(reader);
		return oElem;
	}
	function readUniColor(reader, sName) {
		let oElem = new AscFormat.CUniColor();
		oElem.fromXml(reader, sName);
		return oElem;
	}
	function readColorModifier(reader) {
		let oColor = new AscFormat.CRGBColor();
		oColor.fromXml(reader);
		let oElem = new AscFormat.CColorModifiers();
		if(oColor.Mods) {
			oElem.Mods = oColor.Mods;
		}
		return oElem;
	}

//CT_ChartSpace
	CChartSpace.prototype.fromXml = function(reader) {
		reader.context.curChart = this;
		var name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("chartSpace" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("chartSpace" === name) {
			var elem, depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				switch (reader.GetNameNoNS()) {
					case "date1904" : {
						this.setDate1904(CT_Bool.prototype.toVal(reader, this.date1904));
						break;
					}
					case "lang" : {
						this.setLang(CT_String.prototype.toVal(reader, this.lang));
						break;
					}
					case "roundedCorners" : {
						this.setRoundedCorners(CT_Bool.prototype.toVal(reader, this.roundedCorners));
						break;
					}
					case "AlternateContent" : {
						let elem = new CT_XmlNode();
						elem.fromXml(reader);
						let style;
						if (elem.members["Choice"] && elem.members["Choice"].members["style"]) {
							style = parseInt(elem.members["Choice"].members["style"].attributes["val"]) - 100;
						} else if (elem.members["Fallback"] && elem.members["style"]) {
							style = parseInt(elem.members["Fallback"].members["style"]).attributes["val"];
						}
						if (!isNaN(style)) {
							this.setStyle(style);
						}
						break;
					}
					case "style" : {
						this.setStyle(CT_UInt.prototype.toVal(reader, this.style));
						break;
					}
					case "clrMapOvr" : {
						let oClrMapOvr = new AscFormat.CClrMapOvr();
						oClrMapOvr.fromXml(reader);
						this.setClrMapOvr(oClrMapOvr.overrideClrMapping);
						break;
					}
					case "pivotSource" : {
						elem = new CPivotSource();
						elem.fromXml(reader);
						this.setPivotSource(elem);
						break;
					}
					case "protection" : {
						elem = new CProtection();
						elem.fromXml(reader);
						this.setPivotSource(elem);
						break;
					}
					case "chart" : {
						elem = new AscFormat.CChart();
						elem.fromXml(reader);
						this.setChart(elem);
						break;
					}
					case "spPr" : {
						readSpPr(reader, this);
						break;
					}
					case "txPr" : {
						readTxPr(reader, this);
						break;
					}
					// case "externalData" : {
					// 	this.setExternalData(CT_ExternalData.prototype.toVal(reader, this.externalData));
					// 	this.externalData = CT_ExternalData.prototype.toVal(reader, this.externalData);
					// 	this.externalData = new CT_ExternalData();
					// 	this.externalData.fromXml(reader);
					// 	break;
					// }
					case "printSettings" : {
						elem = new CPrintSettings();
						elem.fromXml(reader);
						this.setPrintSettings(elem);
						break;
					}
					case "userShapes" : {
						let oChartSpace = this;
						let oUserShapesEntry = new AscFormat.IdEntry();
						oUserShapesEntry.fromXml(reader);
						oUserShapesEntry.readItem(reader, function () {
							let oNode = new CT_XmlNode(
								function(reader, name) {
									if(name === "relSizeAnchor") {
										let oAnchor = new AscFormat.CRelSizeAnchor();
										oAnchor.fromXml(reader);
										oChartSpace.addUserShape(null, oAnchor);
									}
									else if(name === "absSizeAnchor") {
										let oAnchor = new AscFormat.CAbsSizeAnchor();
										oAnchor.fromXml(reader);
										oChartSpace.addUserShape(null, oAnchor);
									}
									return true;
								}
							);
							return oNode;
						});
						break;
					}
					//todo styles, colors
				}
			}
		}
		this.correctAxes();
	};
	CChartSpace.prototype.toXml = function(writer) {
		let nOldDocType = writer.context.docType;
		writer.context.docType = AscFormat.XMLWRITER_DOC_TYPE_CHART;
		var name = "c:chartSpace";

		let style;
		if(null !== this.style) {
			style = new CT_XmlNode();
			style.attributes["xmlns:mc"] = "http://schemas.openxmlformats.org/markup-compatibility/2006";
			style.members["mc:Choice"] =new CT_XmlNode();
			style.members["mc:Choice"].attributes["Requires"] = "c14";
			style.members["mc:Choice"].attributes["xmlns:c14"] = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";
			style.members["mc:Choice"].members["c14:style"] =new CT_XmlNode();
			style.members["mc:Choice"].members["c14:style"].attributes["val"] = (100 + this.style).toString();

			style.members["mc:Fallback"] =new CT_XmlNode();
			style.members["mc:Fallback"].members["c:style"] =new CT_XmlNode();
			style.members["mc:Fallback"].members["c:style"].attributes["val"] = this.style.toString();
		}

		writer.WriteXmlString(AscCommonWord.g_sXmlHeader);
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlString(AscCommonWord.g_sXmlChartNamespaces);
		writer.WriteXmlAttributesEnd();

		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.date1904), "c:date1904");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(this.lang), "c:lang");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.roundedCorners), "c:roundedCorners");
		writer.WriteXmlNullable(style, "mc:AlternateContent");
		writer.WriteXmlNullable(this.clrMapOvr, "c:clrMapOvr");
		writer.WriteXmlNullable(this.pivotSource, "c:pivotSource");
		writer.WriteXmlNullable(this.protection, "c:protection");
		writer.WriteXmlNullable(this.chart, "c:chart");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.txPr, "c:txPr");
		// writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.externalData), "c:externalData");
		writer.WriteXmlNullable(this.printSettings, "c:printSettings");
		if(this.userShapes.length > 0) {
			let userShapesPart = writer.context.part.addPart(AscCommon.openXml.Types.chartDrawing);
			let memory = new AscCommon.CMemory();
			memory.context = writer.context;
			memory.WriteXmlNodeStart("c:userShapes");
			memory.WriteXmlAttributeString("xmlns:c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
			memory.WriteXmlAttributeString("xmlns:cdr", "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing");
			memory.WriteXmlAttributeString("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
			memory.WriteXmlAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			memory.WriteXmlAttributesEnd();
			for(let nUSp = 0; nUSp < this.userShapes.length; ++nUSp) {
				this.userShapes[nUSp].toXml(memory);
			}
			memory.WriteXmlNodeEnd("c:userShapes");
			let userShapesData = memory.GetDataUint8();
			userShapesPart.part.setData(userShapesData);
			memory.Seek(0);
			let oEntry = new AscFormat.IdEntry("c:userShapes");
			oEntry.rId = userShapesPart.rId;
			oEntry.toXml(writer);
		}
		writer.WriteXmlNodeEnd(name);

		writer.context.docType = nOldDocType;
	};

	CPivotSource.prototype.fromXml = function(reader) {
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "name" : {
					this.setName(reader.GetTextDecodeXml());
					break;
				}
				case "fmtId" : {
					this.setFmtId(CT_UInt.prototype.toVal(reader, this.fmtId));
					break;
				}
				case "extLst" : {
					break;
				}
			}
		}
	};
	CPivotSource.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullableValueStringEncode("c:name", this.name);
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.fmtId), "c:fmtId");
		writer.WriteXmlNodeEnd(name);
	};

	CProtection.prototype.fromXml = function(reader) {
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "chartObject" : {
					this.setChartObject(CT_Bool.prototype.toVal(reader, this.chartObject));
					break;
				}
				case "data" : {
					this.setData(CT_Bool.prototype.toVal(reader, this.data));
					break;
				}
				case "formatting" : {
					this.setFormatting(CT_Bool.prototype.toVal(reader, this.formatting));
					break;
				}
				case "selection" : {
					this.setSelection(CT_Bool.prototype.toVal(reader, this.selection));
					break;
				}
				case "userInterface" : {
					this.setUserInterface(CT_Bool.prototype.toVal(reader, this.userInterface));
					break;
				}
			}
		}
	};
	CProtection.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.chartObject), "c:chartObject");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.data), "c:data");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.formatting), "c:formatting");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.selection), "c:selection");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.userInterface), "c:userInterface");
		writer.WriteXmlNodeEnd(name);
	};

	CPrintSettings.prototype.fromXml = function(reader) {
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "headerFooter" : {
					elem = new CHeaderFooterChart();
					elem.fromXml(reader);
					this.setHeaderFooter(elem);
					break;
				}
				case "pageMargins" : {
					elem = new CPageMarginsChart();
					elem.fromXml(reader);
					this.setPageMargins(elem);
					break;
				}
				case "pageSetup" : {
					elem = new CPageSetup();
					elem.fromXml(reader);
					this.setPageSetup(elem);
					break;
				}
			}
		}
	};
	CPrintSettings.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.headerFooter, "c:headerFooter");
		writer.WriteXmlNullable(this.pageMargins, "c:pageMargins");
		writer.WriteXmlNullable(this.pageSetup, "c:pageSetup");
		writer.WriteXmlNodeEnd(name);
	};

	CHeaderFooterChart.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "alignWithMargins": {
					this.setAlignWithMargins(reader.GetValueBool());
					break;
				}
				case "differentOddEven": {
					this.setDifferentOddEven(reader.GetValueBool());
					break;
				}
				case "differentFirst": {
					this.setDifferentFirst(reader.GetValueBool());
					break;
				}
			}
		}
	};
	CHeaderFooterChart.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "oddHeader" : {
					this.setOddHeader(reader.GetTextDecodeXml());
					break;
				}
				case "oddFooter" : {
					this.setOddHeader(reader.GetTextDecodeXml());
					break;
				}
				case "evenHeader" : {
					this.setEvenHeader(reader.GetTextDecodeXml());
					break;
				}
				case "evenFooter" : {
					this.setEvenFooter(reader.GetTextDecodeXml());
					break;
				}
				case "firstHeader" : {
					this.setFirstHeader(reader.GetTextDecodeXml());
					break;
				}
				case "firstFooter" : {
					this.setFirstFooter(reader.GetTextDecodeXml());
					break;
				}
			}
		}
	};
	CHeaderFooterChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeBool("alignWithMargins", this.alignWithMargins);
		writer.WriteXmlNullableAttributeBool("differentOddEven", this.differentOddEven);
		writer.WriteXmlNullableAttributeBool("differentFirst", this.differentFirst);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullableValueStringEncode("c:oddHeader", this.oddHeader);
		writer.WriteXmlNullableValueStringEncode("c:oddFooter", this.oddFooter);
		writer.WriteXmlNullableValueStringEncode("c:evenHeader", this.evenHeader);
		writer.WriteXmlNullableValueStringEncode("c:evenFooter", this.evenFooter);
		writer.WriteXmlNullableValueStringEncode("c:firstHeader", this.firstHeader);
		writer.WriteXmlNullableValueStringEncode("c:firstFooter", this.firstFooter);
		writer.WriteXmlNodeEnd(name);
	};

	CPageMarginsChart.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "b": {
					this.setB(reader.GetValueDouble(this.b));
					break;
				}
				case "l": {
					this.setL(reader.GetValueDouble(this.l));
					break;
				}
				case "r": {
					this.setR(reader.GetValueDouble(this.r));
					break;
				}
				case "t": {
					this.setT(reader.GetValueDouble(this.t));
					break;
				}
				case "header": {
					this.setHeader(reader.GetValueDouble(this.header));
					break;
				}
				case "footer": {
					this.setFooter(reader.GetValueDouble(this.footer));
					break;
				}
			}
		}
	};
	CPageMarginsChart.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CPageMarginsChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeDouble("b", this.b);
		writer.WriteXmlNullableAttributeDouble("l", this.l);
		writer.WriteXmlNullableAttributeDouble("r", this.r);
		writer.WriteXmlNullableAttributeDouble("t", this.t);
		writer.WriteXmlNullableAttributeDouble("header", this.header);
		writer.WriteXmlNullableAttributeDouble("footer", this.footer);
		writer.WriteXmlAttributesEnd(true);
	};

	CPageSetup.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "paperSize": {
					this.setPaperSize(reader.GetValueUInt(this.paperSize));
					break;
				}
				case "paperHeight": {
					this.setPaperHeight(AscCommon.universalMeasureToMm(reader.GetValue(), 1, this.paperHeight));
					break;
				}
				case "paperWidth": {
					this.setPaperWidth(AscCommon.universalMeasureToMm(reader.GetValue(), 1, this.paperWidth));
					break;
				}
				case "firstPageNumber": {
					this.setFirstPageNumber(reader.GetValueUInt(this.firstPageNumber));
					break;
				}
				case "orientation": {
					this.setOrientation(fromXml_ST_PageSetupOrientation(reader.GetValue(), this.orientation));
					break;
				}
				case "blackAndWhite": {
					this.setBlackAndWhite(reader.GetValueBool());
					break;
				}
				case "draft": {
					this.setDraft(reader.GetValueBool());
					break;
				}
				case "useFirstPageNumber": {
					this.setUseFirstPageNumb(reader.GetValueBool());
					break;
				}
				case "horizontalDpi": {
					this.setHorizontalDpi(reader.GetValueInt(this.horizontalDpi));
					break;
				}
				case "verticalDpi": {
					this.setVerticalDpi(reader.GetValueInt(this.verticalDpi));
					break;
				}
				case "copies": {
					this.setCopies(reader.GetValueUInt(this.copies));
					break;
				}
			}
		}
	};
	CPageSetup.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CPageSetup.prototype.toXml = function(writer, name) {
		let paperHeight = null;
		if (null !== this.paperHeight) {
			paperHeight = this.paperHeight + "mm";
		}
		let paperWidth = null;
		if (null !== this.paperWidth) {
			paperWidth = this.paperWidth + "mm";
		}
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeUInt("paperSize", this.paperSize);
		writer.WriteXmlNullableAttributeStringEncode("paperHeight", paperHeight);
		writer.WriteXmlNullableAttributeStringEncode("paperWidth", paperWidth);
		writer.WriteXmlNullableAttributeUInt("firstPageNumber", this.firstPageNumber);
		writer.WriteXmlNullableAttributeString("orientation", toXml_ST_PageSetupOrientation(this.orientation));
		writer.WriteXmlNullableAttributeBool("blackAndWhite", this.blackAndWhite);
		writer.WriteXmlNullableAttributeBool("draft", this.draft);
		writer.WriteXmlNullableAttributeBool("useFirstPageNumber", this.useFirstPageNumber);
		writer.WriteXmlNullableAttributeInt("horizontalDpi", this.horizontalDpi);
		writer.WriteXmlNullableAttributeInt("verticalDpi", this.verticalDpi);
		writer.WriteXmlNullableAttributeUInt("copies", this.copies);
		writer.WriteXmlAttributesEnd(true);
	};

//CT_Chart
	CChart.prototype.fromXml = function(reader) {
		let t = this;
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "title" : {
					elem = new CTitle();
					elem.fromXml(reader);
					this.setTitle(elem);
					break;
				}
				case "autoTitleDeleted" : {
					this.setAutoTitleDeleted(CT_Bool.prototype.toVal(reader, this.autoTitleDeleted));
					break;
				}
				case "pivotFmts" : {
					reader.readXmlArray("pivotFmt", function() {
						elem = new CPivotFmt();
						elem.fromXml(reader);
						t.setPivotFmts(elem);
					});
					break;
				}
				case "view3D" : {
					elem = new AscFormat.CView3d();
					elem.fromXml(reader);
					this.setView3D(elem);
					break;
				}
				case "floor" : {
					elem = new AscFormat.CChartWall();
					elem.fromXml(reader);
					this.setFloor(elem);
					break;
				}
				case "sideWall" : {
					elem = new AscFormat.CChartWall();
					elem.fromXml(reader);
					this.setSideWall(elem);
					break;
				}
				case "backWall" : {
					elem = new AscFormat.CChartWall();
					elem.fromXml(reader);
					this.setBackWall(elem);
					break;
				}
				case "plotArea" : {
					var aChartWithAxis = [];
					elem = new AscFormat.CPlotArea();
					elem.fromXml(reader, aChartWithAxis);
					elem.initPostOpen(aChartWithAxis);
					this.setPlotArea(elem);
					break;
				}
				case "legend" : {
					elem = new AscFormat.CLegend();
					elem.fromXml(reader);
					elem.updateLegendPos();
					this.setLegend(elem);
					break;
				}
				case "plotVisOnly" : {
					this.setPlotVisOnly(CT_Bool.prototype.toVal(reader, this.plotVisOnly));
					break;
				}
				case "dispBlanksAs" : {
					this.setDispBlanksAs(fromXml_ST_DispBlanksAs(CT_String.prototype.toVal(reader, this.dispBlanksAs), this.dispBlanksAs));
					break;
				}
				case "showDLblsOverMax" : {
					this.setShowDLblsOverMax(CT_Bool.prototype.toVal(reader, this.showDLblsOverMax));
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.title, "c:title");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.autoTitleDeleted), "c:autoTitleDeleted");
		writer.WriteXmlArray(this.pivotFmts, "c:pivotFmt", "c:pivotFmts");
		writer.WriteXmlNullable(this.view3D, "c:view3D");
		writer.WriteXmlNullable(this.floor, "c:floor");
		writer.WriteXmlNullable(this.sideWall, "c:sideWall");
		writer.WriteXmlNullable(this.backWall, "c:backWall");
		writer.WriteXmlNullable(this.plotArea, "c:plotArea");
		writer.WriteXmlNullable(this.legend, "c:legend");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.plotVisOnly), "c:plotVisOnly");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_DispBlanksAs(this.dispBlanksAs)),
			"c:dispBlanksAs");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showDLblsOverMax), "c:showDLblsOverMax");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CLayout.prototype.fromXml = function (reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "layoutTarget" : {
					this.setLayoutTarget(fromXml_ST_LayoutTarget(CT_String.prototype.toVal(reader, this.layoutTarget), this.layoutTarget));
					break;
				}
				case "xMode" : {
					this.setXMode(fromXml_ST_LayoutMode(CT_String.prototype.toVal(reader, this.xMode), this.xMode));
					break;
				}
				case "yMode" : {
					this.setYMode(fromXml_ST_LayoutMode(CT_String.prototype.toVal(reader, this.yMode), this.yMode));
					break;
				}
				case "wMode" : {
					this.setWMode(fromXml_ST_LayoutMode(CT_String.prototype.toVal(reader, this.wMode), this.wMode));
					break;
				}
				case "hMode" : {
					this.setHMode(fromXml_ST_LayoutMode(CT_String.prototype.toVal(reader, this.hMode), this.hMode));
					break;
				}
				case "x" : {
					this.setX(CT_Double.prototype.toVal(reader, this.x));
					break;
				}
				case "y" : {
					this.setY(CT_Double.prototype.toVal(reader, this.y));
					break;
				}
				case "w" : {
					this.setW(CT_Double.prototype.toVal(reader, this.w));
					break;
				}
				case "h" : {
					this.setH(CT_Double.prototype.toVal(reader, this.h));
					break;
				}
			}
		}
	};
	CLayout.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_LayoutTarget(this.layoutTarget)), "c:layoutTarget");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_LayoutMode(this.xMode)), "c:xMode");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_LayoutMode(this.yMode)), "c:yMode");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_LayoutMode(this.wMode)), "c:wMode");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_LayoutMode(this.hMode)), "c:hMode");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.x), "c:x");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.y), "c:y");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.w), "c:w");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.h), "c:h");
		writer.WriteXmlNodeEnd(name);
	};

	CTitle.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "tx" : {
					elem = new AscFormat.CChartText();
					elem.fromXml(reader);
					elem.setChart(reader.context.curChart);
					this.setTx(elem);
					break;
				}
				case "layout" : {
					readLayout(reader, this);
					break;
				}
				case "overlay" : {
					this.setOverlay(CT_Bool.prototype.toVal(reader, this.overlay));
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "txPr" : {
					readTxPr(reader, this);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CTitle.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.tx, "c:tx");
		writer.WriteXmlNullable(writeLayout(this.layout), "c:layout");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.overlay), "c:overlay");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.txPr, "c:txPr");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CChartText.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "rich" : {
					let elem = readTxPrPlain(reader);
					elem.setParent(this);
					this.setRich(elem);
					break;
				}
				case "strRef" : {
					elem = new AscFormat.CStrRef();
					elem.fromXml(reader);
					this.setStrRef(elem);
					break;
				}
			}
		}
	};
	CChartText.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.rich, "c:rich");
		writer.WriteXmlNullable(this.strRef, "c:strRef");
		writer.WriteXmlNodeEnd(name);
	};

	CStrRef.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "f" : {
					this.setF(reader.GetTextDecodeXml());
					break;
				}
				case "strCache" : {
					elem = new AscFormat.CStrCache();
					elem.fromXml(reader);
					this.setStrCache(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CStrRef.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullableValueStringEncode("c:f", this.f);
		writer.WriteXmlNullable(this.strCache, "c:strCache");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CStrCache.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "ptCount" : {
					this.setPtCount(CT_UInt.prototype.toVal(reader, this.ptCount));
					break;
				}
				case "pt" : {
					elem = new CStringPoint();
					elem.fromXml(reader);
					this.addPt(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CStrCache.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.ptCount), "c:ptCount");
		writer.WriteXmlArray(this.pts, "c:pt");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CStringPoint.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "idx": {
					this.setIdx(reader.GetValueUInt(this.idx));
					break;
				}
			}
		}
	};
	CStringPoint.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "v" : {
					this.setVal(reader.GetTextDecodeXml());
					break;
				}
			}
		}
	};
	CStringPoint.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeUInt("idx", this.idx);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullableValueStringEncode("c:v", this.val);
		writer.WriteXmlNodeEnd(name);
	};

	CPivotFmt.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "idx" : {
					this.setIdx(CT_UInt.prototype.toVal(reader, this.idx));
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "txPr" : {
					readTxPr(reader, this);
					break;
				}
				case "marker" : {
					elem = new AscFormat.CMarker();
					elem.fromXml(reader);
					this.setMarker(elem);
					break;
				}
				case "dLbl" : {
					elem = new AscFormat.CDLbl();
					elem.fromXml(reader);
					elem.correctValues();
					this.setLbl(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CPivotFmt.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.idx), "c:idx");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.txPr, "c:txPr");
		writer.WriteXmlNullable(this.marker, "c:marker");
		writer.WriteXmlNullable(this.dLbl, "c:dLbl");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CMarker.prototype.fromXml = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "symbol" : {
					this.setSymbol(
						fromXml_ST_MarkerStyle(CT_String.prototype.toVal(reader, this.symbol), this.symbol));
					break;
				}
				case "size" : {
					this.setSize(CT_UInt.prototype.toVal(reader, this.size));
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CMarker.prototype.toXml = function(writer, name) {
		let symbol = CT_String.prototype.fromVal(toXml_ST_MarkerStyle(this.symbol));
		let size = CT_UInt.prototype.fromVal(this.size);

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(symbol, "c:symbol");
		writer.WriteXmlNullable(size, "c:size");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};
	CDLbl.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "idx" : {
					this.setIdx(CT_UInt.prototype.toVal(reader, this.idx));
					break;
				}
				case "delete" : {
					this.setDelete(CT_Bool.prototype.toVal(reader, this.delete));
					break;
				}
				case "layout" : {
					readLayout(reader, this);
					break;
				}
				case "tx" : {
					elem = new AscFormat.CChartText();
					elem.fromXml(reader);
					elem.setChart(reader.context.curChart);
					this.setTx(elem);
					break;
				}
				case "numFmt" : {
					elem = new AscFormat.CNumFmt();
					elem.fromXml(reader);
					this.setNumFmt(elem);
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "txPr" : {
					readTxPr(reader, this);
					break;
				}
				case "dLblPos" : {
					this.setDLblPos(
						fromXml_ST_DLblPos(CT_String.prototype.toVal(reader, this.dLblPos), this.dLblPos));
					break;
				}
				case "showLegendKey" : {
					this.setShowLegendKey(CT_Bool.prototype.toVal(reader, this.showLegendKey));
					break;
				}
				case "showVal" : {
					this.setShowVal(CT_Bool.prototype.toVal(reader, this.showVal));
					break;
				}
				case "showCatName" : {
					this.setShowCatName(CT_Bool.prototype.toVal(reader, this.showCatName));
					break;
				}
				case "showSerName" : {
					this.setShowSerName(CT_Bool.prototype.toVal(reader, this.showSerName));
					break;
				}
				case "showPercent" : {
					this.setShowPercent(CT_Bool.prototype.toVal(reader, this.showPercent));
					break;
				}
				case "showBubbleSize" : {
					this.setShowBubbleSize(CT_Bool.prototype.toVal(reader, this.showBubbleSize));
					break;
				}
				case "separator" : {
					this.setSeparator(reader.GetTextDecodeXml());
					break;
				}
			}
		}
	};
	CDLbl.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.idx), "c:idx");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.bDelete), "c:delete");
		writer.WriteXmlNullable(writeLayout(this.layout), "c:layout");
		writer.WriteXmlNullable(this.tx, "c:tx");
		writer.WriteXmlNullable(this.numFmt, "c:numFmt");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.txPr, "c:txPr");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_DLblPos(this.dLblPos)), "c:dLblPos");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showLegendKey), "c:showLegendKey");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showVal), "c:showVal");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showCatName), "c:showCatName");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showSerName), "c:showSerName");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showPercent), "c:showPercent");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showBubbleSize), "c:showBubbleSize");
		writer.WriteXmlNullableValueStringEncode("c:separator", this.separator);
		writer.WriteXmlNodeEnd(name);
	};

	CView3d.prototype.fromXml = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "rotX" : {
					this.setRotX(CT_Int.prototype.toVal(reader, this.rotX));
					break;
				}
				case "hPercent" : {
					//todo percent
					this.setHPercent(CT_Double.prototype.toVal(reader, this.hPercent));
					break;
				}
				case "rotY" : {
					this.setRotY(CT_UInt.prototype.toVal(reader, this.rotY));
					break;
				}
				case "depthPercent" : {
					//todo percent
					this.setDepthPercent(CT_Int.prototype.toVal(reader, this.depthPercent));
					break;
				}
				case "rAngAx" : {
					this.setRAngAx(CT_Bool.prototype.toVal(reader, this.rAngAx));
					break;
				}
				case "perspective" : {
					this.setPerspective(CT_UInt.prototype.toVal(reader, this.perspective));
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CView3d.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_Int.prototype.fromVal(this.rotX), "c:rotX");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.hPercent), "c:hPercent");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.rotY), "c:rotY");
		writer.WriteXmlNullable(CT_Int.prototype.fromVal(this.depthPercent), "c:depthPercent");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.rAngAx), "c:rAngAx");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.perspective), "c:perspective");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CChartWall.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "thickness" : {
					this.setThickness(CT_Int.prototype.toVal(reader, this.thickness));
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "pictureOptions" : {
					elem = new AscFormat.CPictureOptions();
					elem.fromXml(reader);
					this.setPictureOptions(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CChartWall.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_Int.prototype.fromVal(this.thickness), "c:thickness");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.pictureOptions, "c:pictureOptions");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CPictureOptions.prototype.fromXml = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "applyToFront" : {
					this.setApplyToFront(CT_Bool.prototype.toVal(reader, this.applyToFront));
					break;
				}
				case "applyToSides" : {
					this.setApplyToSides(CT_Bool.prototype.toVal(reader, this.applyToSides));
					break;
				}
				case "applyToEnd" : {
					this.setApplyToEnd(CT_Bool.prototype.toVal(reader, this.applyToEnd));
					break;
				}
				case "pictureFormat" : {
					this.setPictureFormat(
						fromXml_ST_PictureFormat(CT_String.prototype.toVal(reader, this.pictureFormat),
							this.pictureFormat));
					break;
				}
				case "pictureStackUnit" : {
					this.setPictureStackUnit(CT_Double.prototype.toVal(reader, this.pictureStackUnit));
					break;
				}
			}
		}
	};
	CPictureOptions.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.applyToFront), "c:applyToFront");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.applyToSides), "c:applyToSides");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.applyToEnd), "c:applyToEnd");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_PictureFormat(this.pictureFormat)),
			"c:pictureFormat");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.pictureStackUnit), "c:pictureStackUnit");
		writer.WriteXmlNodeEnd(name);
	};

	CLegend.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "legendPos" : {
					this.setLegendPos(
						fromXml_ST_LegendPos(CT_String.prototype.toVal(reader, this.legendPos), this.legendPos));
					break;
				}
				case "legendEntry" : {
					elem = new CLegendEntry();
					elem.fromXml(reader);
					this.addLegendEntry(elem);
					break;
				}
				case "layout" : {
					readLayout(reader, this);
					break;
				}
				case "overlay" : {
					this.setOverlay(CT_Bool.prototype.toVal(reader, this.overlay));
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "txPr" : {
					readTxPr(reader, this);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CLegend.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_LegendPos(this.legendPos)), "c:legendPos");
		writer.WriteXmlArray(this.legendEntryes, "c:legendEntry");
		writer.WriteXmlNullable(writeLayout(this.layout), "c:layout");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.overlay), "c:overlay");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.txPr, "c:txPr");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CLegendEntry.prototype.fromXml = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "idx" : {
					this.setIdx(CT_UInt.prototype.toVal(reader, this.idx));
					break;
				}
				case "delete" : {
					this.setDelete(CT_Bool.prototype.toVal(reader, this.delete));
					break;
				}
				case "txPr" : {
					readTxPr(reader, this);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CLegendEntry.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.idx), "c:idx");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.bDelete), "c:delete");
		writer.WriteXmlNullable(this.txPr, "c:txPr");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};
//CT_PlotArea
	CPlotArea.prototype.fromXml = function(reader, aChartWithAxis) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "layout" : {
					readLayout(reader, this);
					break;
				}
				case "area3DChart" : {
					elem = new AscFormat.CAreaChart();
					elem.fromXml(reader, aChartWithAxis);
					this.addChart(elem);
					break;
				}
				case "areaChart" : {
					elem = new AscFormat.CAreaChart();
					elem.fromXml(reader, aChartWithAxis);
					this.addChart(elem);
					break;
				}
				case "bar3DChart" : {
					elem = new AscFormat.CBarChart();
					elem.fromXml(reader, aChartWithAxis);
					elem.set3D(true);
					this.addChart(elem);
					break;
				}
				case "barChart" : {
					elem = new AscFormat.CBarChart();
					elem.fromXml(reader, aChartWithAxis);
					this.addChart(elem);
					break;
				}
				case "bubbleChart" : {
					elem = new AscFormat.CBubbleChart();
					elem.fromXml(reader, aChartWithAxis);
					//bubble -> scatter
					let scatter = elem.convertToScutterChart();
					CorrectChartWithAxis(elem, scatter, aChartWithAxis);
					this.addChart(scatter);
					break;
				}
				case "doughnutChart" : {
					elem = new AscFormat.CDoughnutChart();
					elem.fromXml(reader, aChartWithAxis);
					this.addChart(elem);
					break;
				}
				case "line3DChart" : {
					elem = new AscFormat.CLineChart();
					elem.fromXml(reader, aChartWithAxis);
					elem.convert3Dto2D();
					this.addChart(elem);
					break;
				}
				case "lineChart" : {
					elem = new AscFormat.CLineChart();
					elem.fromXml(reader, aChartWithAxis);
					this.addChart(elem);
					break;
				}
				case "ofPieChart" : {
					elem = new AscFormat.COfPieChart();
					elem.fromXml(reader, aChartWithAxis);
					let pie = elem.convertToPieChart();
					CorrectChartWithAxis(elem, pie, aChartWithAxis);
					this.addChart(pie);
					break;
				}
				case "pie3DChart" : {
					elem = new AscFormat.CPieChart();
					elem.fromXml(reader, aChartWithAxis);
					elem.set3D(true);
					this.addChart(elem);
					break;
				}
				case "pieChart" : {
					elem = new AscFormat.CPieChart();
					elem.fromXml(reader, aChartWithAxis);
					this.addChart(elem);
					break;
				}
				case "radarChart" : {
					elem = new AscFormat.CRadarChart();
					elem.fromXml(reader, aChartWithAxis);
					var line = elem.convertToLineChart();
					CorrectChartWithAxis(elem, line, aChartWithAxis);
					this.addChart(line);
					break;
				}
				case "scatterChart" : {
					elem = new AscFormat.CScatterChart();
					elem.fromXml(reader, aChartWithAxis);
					this.addChart(elem);
					break;
				}
				case "stockChart" : {
					elem = new AscFormat.CStockChart();
					elem.fromXml(reader, aChartWithAxis);
					this.addChart(elem);
					break;
				}
				case "surface3DChart" : {
					elem = new AscFormat.CSurfaceChart();
					elem.fromXml(reader, aChartWithAxis);
					this.addChart(elem);
					break;
				}
				case "surfaceChart" : {
					elem = new AscFormat.CSurfaceChart();
					elem.fromXml(reader, aChartWithAxis);
					this.addChart(elem);
					break;
				}
				case "valAx" : {
					elem = new AscFormat.CValAx();
					elem.fromXml(reader);
					this.addAxis(elem);
					break;
				}
				case "catAx" : {
					elem = new AscFormat.CCatAx();
					elem.fromXml(reader);
					this.addAxis(elem);
					break;
				}
				case "dateAx" : {
					elem = new AscFormat.CDateAx();
					elem.fromXml(reader);
					this.addAxis(elem);
					break;
				}
				case "serAx" : {
					elem = new AscFormat.CSerAx();
					elem.fromXml(reader);
					this.addAxis(elem);
					break;
				}
				case "dTable" : {
					elem = new AscFormat.CDTable();
					elem.fromXml(reader);
					this.setDTable(elem);
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CPlotArea.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(writeLayout(this.layout), "c:layout");

		this.charts.forEach(function(chart) {
			let name;
			if (chart instanceof AscFormat.CAreaChart) {
				name = "c:areaChart";
			} else if (chart instanceof AscFormat.CBarChart) {
				name = chart.b3D ? "c:bar3DChart" : "c:barChart";
			} else if (chart instanceof AscFormat.CBubbleChart) {
				name = "c:bubbleChart";
			} else if (chart instanceof AscFormat.CDoughnutChart) {
				name = "c:doughnutChart";
			} else if (chart instanceof AscFormat.CLineChart) {
				name = "c:lineChart";
			} else if (chart instanceof AscFormat.COfPieChart) {
				name = "c:ofPieChart";
			} else if (chart instanceof AscFormat.CPieChart) {
				name = (!chart.view3D && !chart.b3D) ? "c:pieChart" : "c:pie3DChart";
			} else if (chart instanceof AscFormat.CRadarChart) {
				name = "c:radarChart";
			} else if (chart instanceof AscFormat.CScatterChart) {
				name = "c:scatterChart";
			} else if (chart instanceof AscFormat.CStockChart) {
				name = "c:stockChart";
			} else if (chart instanceof AscFormat.CSurfaceChart) {
				name = "c:surfaceChart";
			}
			if (name) {
				chart.toXml(writer, name);
			}
		});
		this.axId.forEach(function(axis) {
			let name;
			if (axis instanceof AscFormat.CCatAx) {
				name = "c:catAx";
			} else if (axis instanceof AscFormat.CValAx) {
				name = "c:valAx";
			} else if (axis instanceof AscFormat.CDateAx) {
				name = "c:dateAx";
			} else if (axis instanceof AscFormat.CSerAx) {
				name = "c:serAx";
			}
			if (name) {
				axis.toXml(writer, name);
			}
		});
		writer.WriteXmlNullable(this.dTable, "c:dTable");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNodeEnd(name);
	};

	CDTable.prototype.fromXml = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "showHorzBorder" : {
					this.setShowHorzBorder(CT_Bool.prototype.toVal(reader, this.showHorzBorder));
					break;
				}
				case "showVertBorder" : {
					this.setShowVertBorder(CT_Bool.prototype.toVal(reader, this.showVertBorder));
					break;
				}
				case "showOutline" : {
					this.setShowOutline(CT_Bool.prototype.toVal(reader, this.showOutline));
					break;
				}
				case "showKeys" : {
					this.setShowKeys(CT_Bool.prototype.toVal(reader, this.showKeys));
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "txPr" : {
					readTxPr(reader, this);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CDTable.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showHorzBorder), "c:showHorzBorder");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showVertBorder), "c:showVertBorder");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showOutline), "c:showOutline");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showKeys), "c:showKeys");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.txPr, "c:txPr");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CAreaChart.prototype.fromXml = function(reader, aChartWithAxis) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "grouping" : {
					this.setGrouping(
						fromXml_ST_Grouping(CT_String.prototype.toVal(reader, this.grouping), this.grouping));
					break;
				}
				case "varyColors" : {
					this.setVaryColors(CT_Bool.prototype.toVal(reader, this.varyColors));
					break;
				}
				case "ser" : {
					elem = new AscFormat.CAreaSeries();
					elem.fromXml(reader);
					this.addSer(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "dropLines" : {
					elem = readChartLines(reader);
					this.setDropLines(elem || new AscFormat.CSpPr());
					break;
				}
				case "gapDepth" : {
					//todo percent
					// this.setGapDepth(CT_Double.prototype.toVal(reader, this.gapDepth));
					break;
				}
				case "axId" : {
					aChartWithAxis.push({ axisId: CT_Int.prototype.toVal(reader, this.axId), chart: this });
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CAreaChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_Grouping(this.grouping)), "c:grouping");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.varyColors), "c:varyColors");
		writer.WriteXmlArray(this.series, "c:ser");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlNullable(writeChartLines(this.dropLines), "c:dropLines");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.gapDepth), "c:gapDepth");
		this.axId.forEach(function(axis) {
			writer.WriteXmlNullable(CT_Int.prototype.fromVal(axis.axId), "c:axId");
		});
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CAreaSeries.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "idx" : {
					this.setIdx(CT_UInt.prototype.toVal(reader, this.idx));
					break;
				}
				case "order" : {
					this.setOrder(CT_UInt.prototype.toVal(reader, this.order));
					break;
				}
				case "tx" : {
					elem = new AscFormat.CTx();
					elem.fromXml(reader);
					this.setTx(elem);
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "pictureOptions" : {
					elem = new AscFormat.CPictureOptions();
					elem.fromXml(reader);
					this.setPictureOptions(elem);
					break;
				}
				case "dPt" : {
					elem = new AscFormat.CDPt();
					elem.fromXml(reader);
					this.addDPt(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "trendline" : {
					elem = new AscFormat.CTrendLine();
					elem.fromXml(reader);
					this.setTrendline(elem);
					break;
				}
				case "errBars" : {
					elem = new AscFormat.CErrBars();
					elem.fromXml(reader);
					this.setErrBars(elem);
					break;
				}
				case "cat" : {
					elem = new AscFormat.CCat();
					elem.fromXml(reader);
					this.setCat(elem);
					break;
				}
				case "val" : {
					elem = new AscFormat.CYVal();
					elem.fromXml(reader);
					this.setVal(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CAreaSeries.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.idx), "c:idx");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.order), "c:order");
		writer.WriteXmlNullable(this.tx, "c:tx");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.pictureOptions, "c:pictureOptions");
		writer.WriteXmlArray(this.dPt, "c:dPt");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlArray(this.trendline, "c:trendline");
		writer.WriteXmlArray(this.errBars, "c:errBars");
		writer.WriteXmlNullable(this.cat, "c:cat");
		writer.WriteXmlNullable(this.val, "c:val");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CTx.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "strRef" : {
					elem = new AscFormat.CStrRef();
					elem.fromXml(reader);
					this.setStrRef(elem);
					break;
				}
				case "v" : {
					this.setVal(reader.GetTextDecodeXml());
					break;
				}
			}
		}
	};
	CTx.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.strRef, "c:strRef");
		writer.WriteXmlNullableValueStringEncode("c:v", this.val);
		writer.WriteXmlNodeEnd(name);
	};

	CDPt.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "idx" : {
					this.setIdx(CT_UInt.prototype.toVal(reader, this.idx));
					break;
				}
				case "invertIfNegative" : {
					this.setInvertIfNegative(CT_Bool.prototype.toVal(reader, this.invertIfNegative));
					break;
				}
				case "marker" : {
					elem = new AscFormat.CMarker();
					elem.fromXml(reader);
					this.setMarker(elem);
					break;
				}
				case "bubble3D" : {
					this.setBubble3D(CT_Bool.prototype.toVal(reader, this.bubble3D));
					break;
				}
				case "explosion" : {
					this.setExplosion(CT_UInt.prototype.toVal(reader, this.explosion));
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "pictureOptions" : {
					elem = new AscFormat.CPictureOptions();
					elem.fromXml(reader);
					this.setPictureOptions(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CDPt.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.idx), "c:idx");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.invertIfNegative), "c:invertIfNegative");
		writer.WriteXmlNullable(this.marker, "c:marker");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.bubble3D), "c:bubble3D");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.explosion), "c:explosion");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.pictureOptions, "c:pictureOptions");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};
	CDLbls.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "dLbl" : {
					elem = new CDLbl();
					elem.fromXml(reader);
					elem.correctValues();
					this.addDLbl(elem);
					break;
				}
				case "delete" : {
					this.setDelete(CT_Bool.prototype.toVal(reader, this.delete));
					break;
				}
				case "numFmt" : {
					elem = new AscFormat.CNumFmt();
					elem.fromXml(reader);
					this.setNumFmt(elem);
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "txPr" : {
					readTxPr(reader, this);
					break;
				}
				case "dLblPos" : {
					this.setDLblPos(
						fromXml_ST_DLblPos(CT_String.prototype.toVal(reader, this.dLblPos), this.dLblPos));
					break;
				}
				case "showLegendKey" : {
					this.setShowLegendKey(CT_Bool.prototype.toVal(reader, this.showLegendKey));
					break;
				}
				case "showVal" : {
					this.setShowVal(CT_Bool.prototype.toVal(reader, this.showVal));
					break;
				}
				case "showCatName" : {
					this.setShowCatName(CT_Bool.prototype.toVal(reader, this.showCatName));
					break;
				}
				case "showSerName" : {
					this.setShowSerName(CT_Bool.prototype.toVal(reader, this.showSerName));
					break;
				}
				case "showPercent" : {
					this.setShowPercent(CT_Bool.prototype.toVal(reader, this.showPercent));
					break;
				}
				case "showBubbleSize" : {
					this.setShowBubbleSize(CT_Bool.prototype.toVal(reader, this.showBubbleSize));
					break;
				}
				case "separator" : {
					this.setSeparator(reader.GetTextDecodeXml());
					break;
				}
				case "showLeaderLines" : {
					this.setShowLeaderLines(CT_Bool.prototype.toVal(reader, this.showLeaderLines));
					break;
				}
				case "leaderLines" : {
					elem = readChartLines(reader);
					this.setLeaderLines(elem || new AscFormat.CSpPr());
					break;
				}
			}
		}
	};
	CDLbls.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlArray(this.dLbl, "c:dLbl");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.bDelete), "c:delete");
		writer.WriteXmlNullable(this.numFmt, "c:numFmt");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.txPr, "c:txPr");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_DLblPos(this.dLblPos)), "c:dLblPos");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showLegendKey), "c:showLegendKey");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showVal), "c:showVal");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showCatName), "c:showCatName");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showSerName), "c:showSerName");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showPercent), "c:showPercent");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showBubbleSize), "c:showBubbleSize");
		writer.WriteXmlNullableValueStringEncode("c:separator", this.separator);
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showLeaderLines), "c:showLeaderLines");
		writer.WriteXmlNullable(writeChartLines(this.leaderLines), "c:leaderLines");
// 		writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CNumFmt.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "formatCode": {
					this.setFormatCode(reader.GetValueDecodeXml());
					break;
				}
				case "sourceLinked": {
					this.setSourceLinked(reader.GetValueBool());
					break;
				}
			}
		}
	};
	CNumFmt.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CNumFmt.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("formatCode", this.formatCode);
		writer.WriteXmlNullableAttributeBool("sourceLinked", this.sourceLinked);
		writer.WriteXmlAttributesEnd(true);
	};

	CTrendLine.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "name" : {
					this.setName(reader.GetTextDecodeXml());
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "trendlineType" : {
					this.setTrendlineType(
						fromXml_ST_TrendlineType(CT_String.prototype.toVal(reader, this.trendlineType),
							this.trendlineType));
					break;
				}
				case "order" : {
					this.setOrder(CT_UInt.prototype.toVal(reader, this.order));
					break;
				}
				case "period" : {
					this.setPeriod(CT_UInt.prototype.toVal(reader, this.order));
					break;
				}
				case "forward" : {
					this.setForward(CT_Double.prototype.toVal(reader, this.forward));
					break;
				}
				case "backward" : {
					this.setBackward(CT_Double.prototype.toVal(reader, this.backward));
					break;
				}
				case "intercept" : {
					this.setIntercept(CT_Double.prototype.toVal(reader, this.intercept));
					break;
				}
				case "dispRSqr" : {
					this.setDispRSqr(CT_Bool.prototype.toVal(reader, this.dispRSqr));
					break;
				}
				case "dispEq" : {
					this.setDispEq(CT_Bool.prototype.toVal(reader, this.dispEq));
					break;
				}
				case "trendlineLbl" : {
					elem = new AscFormat.CDLbl();
					elem.fromXml(reader);
					elem.correctValues();
					this.setTrendlineLbl(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CTrendLine.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullableValueStringEncode("c:name", this.name);
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_TrendlineType(this.trendlineType)),
			"c:trendlineType");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.order), "c:order");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.period), "c:period");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.forward), "c:forward");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.backward), "c:backward");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.intercept), "c:intercept");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.dispRSqr), "c:dispRSqr");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.dispEq), "c:dispEq");
		writer.WriteXmlNullable(this.trendlineLbl, "c:trendlineLbl");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CErrBars.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "errDir" : {
					this.setErrDir(fromXml_ST_ErrDir(CT_String.prototype.toVal(reader, this.errDir), this.errDir));
					break;
				}
				case "errBarType" : {
					this.setErrBarType(
						fromXml_ST_ErrBarType(CT_String.prototype.toVal(reader, this.errBarType), this.errBarType));
					break;
				}
				case "errValType" : {
					this.setErrBarType(
						fromXml_ST_ErrValType(CT_String.prototype.toVal(reader, this.errValType), this.errValType));
					break;
				}
				case "noEndCap" : {
					this.setNoEndCap(CT_Bool.prototype.toVal(reader, this.noEndCap));
					break;
				}
				case "plus" : {
					elem = new AscFormat.CMinusPlus();
					elem.fromXml(reader);
					this.setPlus(elem);
					break;
				}
				case "minus" : {
					elem = new AscFormat.CMinusPlus();
					elem.fromXml(reader);
					this.setMinus(elem);
					break;
				}
				case "val" : {
					this.setVal(CT_Double.prototype.toVal(reader, this.val));
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CErrBars.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_ErrDir(this.errDir)), "c:errDir");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_ErrBarType(this.errBarType)), "c:errBarType");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_ErrValType(this.errValType)), "c:errValType");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.noEndCap), "c:noEndCap");
		writer.WriteXmlNullable(this.plus, "c:plus");
		writer.WriteXmlNullable(this.minus, "c:minus");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.val), "c:val");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CMinusPlus.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "numRef" : {
					elem = new AscFormat.CNumRef();
					elem.fromXml(reader);
					this.setNumRef(elem);
					break;
				}
				case "numLit" : {
					elem = new AscFormat.CNumLit();
					elem.fromXml(reader);
					this.setNumLit(elem);
					break;
				}
			}
		}
	};
	CMinusPlus.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.numRef, "c:numRef");
		writer.WriteXmlNullable(this.numLit, "c:numLit");
		writer.WriteXmlNodeEnd(name);
	};

	CNumRef.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "f" : {
					this.setF(reader.GetTextDecodeXml());
					break;
				}
				case "numCache" : {
					elem = new AscFormat.CNumLit();
					elem.fromXml(reader);
					this.setNumCache(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CNumRef.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullableValueStringEncode("c:f", this.f);
		writer.WriteXmlNullable(this.numCache, "c:numCache");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CNumLit.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "formatCode" : {
					this.setFormatCode(reader.GetTextDecodeXml());
					break;
				}
				case "ptCount" : {
					this.setPtCount(CT_UInt.prototype.toVal(reader, this.ptCount));
					break;
				}
				case "pt" : {
					elem = new AscFormat.CNumericPoint();
					elem.fromXml(reader);
					this.addPt(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CNumLit.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullableValueStringEncode("c:formatCode", this.formatCode);
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.ptCount), "c:ptCount");
		writer.WriteXmlArray(this.pts, "c:pt");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CNumericPoint.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "idx": {
					this.setIdx(reader.GetValueUInt(this.idx));
					break;
				}
				case "formatCode": {
					this.setFormatCode(reader.GetValueDecodeXml());
					break;
				}
			}
		}
	};
	CNumericPoint.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "v" : {
					this.setVal(reader.GetTextDouble(0));
					break;
				}
			}
		}
	};
	CNumericPoint.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeUInt("idx", this.idx);
		writer.WriteXmlNullableAttributeStringEncode("formatCode", this.formatCode);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullableValueDouble("c:v", this.val);
		writer.WriteXmlNodeEnd(name);
	};

	CCat.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "multiLvlStrRef" : {
					elem = new AscFormat.CMultiLvlStrRef();
					elem.fromXml(reader);
					this.setMultiLvlStrRef(elem);
					break;
				}
				case "numRef" : {
					elem = new AscFormat.CNumRef();
					elem.fromXml(reader);
					this.setNumRef(elem);
					break;
				}
				case "numLit" : {
					elem = new AscFormat.CNumLit();
					elem.fromXml(reader);
					this.setNumLit(elem);
					break;
				}
				case "strRef" : {
					elem = new AscFormat.CStrRef();
					elem.fromXml(reader);
					this.setStrRef(elem);
					break;
				}
				case "strLit" : {
					elem = new AscFormat.CStrCache();
					elem.fromXml(reader);
					this.setStrLit(elem);
					break;
				}
			}
		}
	};
	CCat.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.multiLvlStrRef, "c:multiLvlStrRef");
		writer.WriteXmlNullable(this.numRef, "c:numRef");
		writer.WriteXmlNullable(this.numLit, "c:numLit");
		writer.WriteXmlNullable(this.strRef, "c:strRef");
		writer.WriteXmlNullable(this.strLit, "c:strLit");
		writer.WriteXmlNodeEnd(name);
	};

	CYVal.prototype.fromXml = CMinusPlus.prototype.fromXml;
	CYVal.prototype.toXml = CMinusPlus.prototype.toXml;

	CMultiLvlStrRef.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "f" : {
					this.setF(reader.GetTextDecodeXml());
					break;
				}
				case "multiLvlStrCache" : {
					elem = new AscFormat.CMultiLvlStrCache();
					elem.fromXml(reader);
					this.setMultiLvlStrCache(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CMultiLvlStrRef.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullableValueStringEncode("c:f", this.f);
		writer.WriteXmlNullable(this.multiLvlStrCache, "c:multiLvlStrCache");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CMultiLvlStrCache.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "ptCount" : {
					this.setPtCount(CT_UInt.prototype.toVal(reader, this.ptCount));
					break;
				}
				case "lvl" : {
					elem = new AscFormat.CStrCache();
					elem.fromXml(reader);
					this.addLvl(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CMultiLvlStrCache.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.ptCount), "c:ptCount");
		writer.WriteXmlArray(this.lvl, "c:lvl");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CBarChart.prototype.fromXml = function(reader, aChartWithAxis) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "barDir" : {
					this.setBarDir(fromXml_ST_BarDir(CT_String.prototype.toVal(reader, this.barDir), this.barDir));
					break;
				}
				case "grouping" : {
					this.setGrouping(
						fromXml_ST_BarGrouping(CT_String.prototype.toVal(reader, this.grouping), this.grouping));
					break;
				}
				case "varyColors" : {
					this.setVaryColors(CT_Bool.prototype.toVal(reader, this.varyColors));
					break;
				}
				case "ser" : {
					elem = new AscFormat.CBarSeries();
					elem.fromXml(reader);
					this.addSer(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "gapWidth" : {
					//todo percent
					this.setGapWidth(CT_Double.prototype.toVal(reader, this.gapWidth));
					break;
				}
				case "overlap" : {
					//todo percent
					this.setOverlap(CT_Double.prototype.toVal(reader, this.overlap));
					break;
				}
				case "gapDepth" : {
					//todo percent
					this.setGapDepth(CT_Double.prototype.toVal(reader, this.gapDepth));
					break;
				}
				case "shape" : {
					this.setShape(fromXml_ST_Shape(CT_String.prototype.toVal(reader, this.shape), this.shape));
					break;
				}
				case "axId" : {
					aChartWithAxis.push({ axisId: CT_Int.prototype.toVal(reader, this.axId), chart: this });
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CBarChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_BarDir(this.barDir)), "c:barDir");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_BarGrouping(this.grouping)), "c:grouping");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.varyColors), "c:varyColors");
		writer.WriteXmlArray(this.series, "c:ser");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.gapWidth), "c:gapWidth");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.overlap), "c:overlap");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.gapDepth), "c:gapDepth");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_Shape(this.shape)), "c:shape");
		this.axId.forEach(function(axis) {
			writer.WriteXmlNullable(CT_Int.prototype.fromVal(axis.axId), "c:axId");
		});
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CBarSeries.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "idx" : {
					this.setIdx(CT_UInt.prototype.toVal(reader, this.idx));
					break;
				}
				case "order" : {
					this.setOrder(CT_UInt.prototype.toVal(reader, this.order));
					break;
				}
				case "tx" : {
					elem = new AscFormat.CTx();
					elem.fromXml(reader);
					this.setTx(elem);
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "invertIfNegative" : {
					this.setInvertIfNegative(CT_Bool.prototype.toVal(reader, this.invertIfNegative));
					break;
				}
				case "pictureOptions" : {
					elem = new AscFormat.CPictureOptions();
					elem.fromXml(reader);
					this.setPictureOptions(elem);
					break;
				}
				case "dPt" : {
					elem = new AscFormat.CDPt();
					elem.fromXml(reader);
					this.addDPt(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "trendline" : {
					elem = new AscFormat.CTrendLine();
					elem.fromXml(reader);
					this.setTrendline(elem);
					break;
				}
				case "errBars" : {
					elem = new AscFormat.CErrBars();
					elem.fromXml(reader);
					this.setErrBars(elem);
					break;
				}
				case "cat" : {
					elem = new AscFormat.CCat();
					elem.fromXml(reader);
					this.setCat(elem);
					break;
				}
				case "val" : {
					elem = new AscFormat.CYVal();
					elem.fromXml(reader);
					this.setVal(elem);
					break;
				}
				case "shape" : {
					this.setShape(fromXml_ST_Shape(CT_String.prototype.toVal(reader, this.shape), this.shape));
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CBarSeries.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.idx), "c:idx");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.order), "c:order");
		writer.WriteXmlNullable(this.tx, "c:tx");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.invertIfNegative), "c:invertIfNegative");
		writer.WriteXmlNullable(this.pictureOptions, "c:pictureOptions");
		writer.WriteXmlArray(this.dPt, "c:dPt");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlArray(this.trendline, "c:trendline");
		writer.WriteXmlNullable(this.errBars, "c:errBars");
		writer.WriteXmlNullable(this.cat, "c:cat");
		writer.WriteXmlNullable(this.val, "c:val");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_Shape(this.shape)), "c:shape");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CBubbleChart.prototype.fromXml = function(reader, aChartWithAxis) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "varyColors" : {
					this.setVaryColors(CT_Bool.prototype.toVal(reader, this.varyColors));
					break;
				}
				case "ser" : {
					elem = new AscFormat.CBubbleSeries();
					elem.fromXml(reader);
					this.addSer(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "bubble3D" : {
					this.setBubble3D(CT_Bool.prototype.toVal(reader, this.bubble3D));
					break;
				}
				case "bubbleScale" : {
					//todo percent
					this.setBubbleScale(CT_Double.prototype.toVal(reader, this.bubbleScale));
					break;
				}
				case "showNegBubbles" : {
					this.setShowNegBubbles(CT_Bool.prototype.toVal(reader, this.showNegBubbles));
					break;
				}
				case "sizeRepresents" : {
					this.setShape(fromXml_ST_SizeRepresents(CT_String.prototype.toVal(reader, this.sizeRepresents),
						this.sizeRepresents));
					break;
				}
				case "axId" : {
					aChartWithAxis.push({ axisId: CT_Int.prototype.toVal(reader, this.axId), chart: this });
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CBubbleChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.varyColors), "c:varyColors");
		writer.WriteXmlArray(this.series, "c:ser");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.bubble3D), "c:bubble3D");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.bubbleScale), "c:bubbleScale");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.showNegBubbles), "c:showNegBubbles");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_SizeRepresents(this.sizeRepresents)),
			"c:sizeRepresents");
		this.axId.forEach(function(axis) {
			writer.WriteXmlNullable(CT_Int.prototype.fromVal(axis.axId), "c:axId");
		});
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CBubbleSeries.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "idx" : {
					this.setIdx(CT_UInt.prototype.toVal(reader, this.idx));
					break;
				}
				case "order" : {
					this.setOrder(CT_UInt.prototype.toVal(reader, this.order));
					break;
				}
				case "tx" : {
					elem = new AscFormat.CTx();
					elem.fromXml(reader);
					this.setTx(elem);
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "invertIfNegative" : {
					this.setInvertIfNegative(CT_Bool.prototype.toVal(reader, this.invertIfNegative));
					break;
				}
				case "dPt" : {
					elem = new AscFormat.CDPt();
					elem.fromXml(reader);
					this.addDPt(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "trendline" : {
					elem = new AscFormat.CTrendLine();
					elem.fromXml(reader);
					this.setTrendline(elem);
					break;
				}
				case "errBars" : {
					elem = new AscFormat.CErrBars();
					elem.fromXml(reader);
					this.setErrBars(elem);
					break;
				}
				case "xVal" : {
					elem = new AscFormat.CCat();
					elem.fromXml(reader);
					this.setXVal(elem);
					break;
				}
				case "yVal" : {
					elem = new AscFormat.CYVal();
					elem.fromXml(reader);
					this.setYVal(elem);
					break;
				}
				case "bubbleSize" : {
					elem = new AscFormat.CYVal();
					elem.fromXml(reader);
					this.setBubbleSize(elem);
					break;
				}
				case "bubble3D" : {
					this.setBubble3D(CT_Bool.prototype.toVal(reader, this.bubble3D));
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CBubbleSeries.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.idx), "c:idx");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.order), "c:order");
		writer.WriteXmlNullable(this.tx, "c:tx");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.invertIfNegative), "c:invertIfNegative");
		writer.WriteXmlArray(this.dPt, "c:dPt");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlArray(this.trendline, "c:trendline");
		writer.WriteXmlArray(this.errBars, "c:errBars");
		writer.WriteXmlNullable(this.xVal, "c:xVal");
		writer.WriteXmlNullable(this.yVal, "c:yVal");
		writer.WriteXmlNullable(this.bubbleSize, "c:bubbleSize");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.bubble3D), "c:bubble3D");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CDoughnutChart.prototype.fromXml = function(reader, aChartWithAxis) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "varyColors" : {
					this.setVaryColors(CT_Bool.prototype.toVal(reader, this.varyColors));
					break;
				}
				case "ser" : {
					elem = new CPieSeries();
					elem.fromXml(reader);
					this.addSer(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "firstSliceAng" : {
					this.setFirstSliceAng(CT_UInt.prototype.toVal(reader, this.firstSliceAng));
					break;
				}
				case "holeSize" : {
					//todo percent
					this.setHoleSize(CT_Double.prototype.toVal(reader, this.holeSize));
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CDoughnutChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.varyColors), "c:varyColors");
		writer.WriteXmlArray(this.series, "c:ser");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.firstSliceAng), "c:firstSliceAng");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.holeSize), "c:holeSize");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};
	CPieSeries.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "idx" : {
					this.setIdx(CT_UInt.prototype.toVal(reader, this.idx));
					break;
				}
				case "order" : {
					this.setOrder(CT_UInt.prototype.toVal(reader, this.order));
					break;
				}
				case "tx" : {
					elem = new AscFormat.CTx();
					elem.fromXml(reader);
					this.setTx(elem);
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "explosion" : {
					this.setExplosion(CT_UInt.prototype.toVal(reader, this.explosion));
					break;
				}
				case "dPt" : {
					elem = new AscFormat.CDPt();
					elem.fromXml(reader);
					this.addDPt(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "cat" : {
					elem = new AscFormat.CCat();
					elem.fromXml(reader);
					this.setCat(elem);
					break;
				}
				case "val" : {
					elem = new AscFormat.CYVal();
					elem.fromXml(reader);
					this.setVal(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CPieSeries.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.idx), "c:idx");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.order), "c:order");
		writer.WriteXmlNullable(this.tx, "c:tx");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.explosion), "c:explosion");
		writer.WriteXmlArray(this.dPt, "c:dPt");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlNullable(this.cat, "c:cat");
		writer.WriteXmlNullable(this.val, "c:val");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CLineChart.prototype.fromXml = function(reader, aChartWithAxis) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "grouping" : {
					this.setGrouping(
						fromXml_ST_Grouping(CT_String.prototype.toVal(reader, this.grouping), this.grouping));
					break;
				}
				case "varyColors" : {
					this.setVaryColors(CT_Bool.prototype.toVal(reader, this.varyColors));
					break;
				}
				case "ser" : {
					elem = new AscFormat.CLineSeries();
					elem.fromXml(reader);
					this.addSer(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "dropLines" : {
					elem = readChartLines(reader);
					this.setDropLines(elem || new AscFormat.CSpPr());
					break;
				}
				case "hiLowLines" : {
					elem = readChartLines(reader);
					this.setHiLowLines(elem || new AscFormat.CSpPr());
					break;
				}
				case "upDownBars" : {
					elem = new AscFormat.CUpDownBars();
					elem.fromXml(reader);
					this.setUpDownBars(elem);
					break;
				}
				case "marker" : {
					this.setMarker(CT_Bool.prototype.toVal(reader, this.marker));
					break;
				}
				case "smooth" : {
					this.setSmooth(CT_Bool.prototype.toVal(reader, true));
					break;
				}
				case "axId" : {
					aChartWithAxis.push({ axisId: CT_Int.prototype.toVal(reader, this.axId), chart: this });
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
		if(this.smooth === null){
			this.setSmooth(false);
		}
	};
	CLineChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_Grouping(this.grouping)), "c:grouping");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.varyColors), "c:varyColors");
		writer.WriteXmlArray(this.series, "c:ser");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlNullable(writeChartLines(this.dropLines), "c:dropLines");
		writer.WriteXmlNullable(writeChartLines(this.hiLowLines), "c:hiLowLines");
		writer.WriteXmlNullable(this.upDownBars, "c:upDownBars");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.marker), "c:marker");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.smooth), "c:smooth");
		this.axId.forEach(function(axis) {
			writer.WriteXmlNullable(CT_Int.prototype.fromVal(axis.axId), "c:axId");
		});
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};
	CLineSeries.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "idx" : {
					this.setIdx(CT_UInt.prototype.toVal(reader, this.idx));
					break;
				}
				case "order" : {
					this.setOrder(CT_UInt.prototype.toVal(reader, this.order));
					break;
				}
				case "tx" : {
					elem = new AscFormat.CTx();
					elem.fromXml(reader);
					this.setTx(elem);
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "marker" : {
					elem = new AscFormat.CMarker();
					elem.fromXml(reader);
					this.setMarker(elem);
					break;
				}
				case "dPt" : {
					elem = new AscFormat.CDPt();
					elem.fromXml(reader);
					this.addDPt(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "trendline" : {
					elem = new AscFormat.CTrendLine();
					elem.fromXml(reader);
					this.setTrendline(elem);
					break;
				}
				case "errBars" : {
					elem = new AscFormat.CErrBars();
					elem.fromXml(reader);
					this.setErrBars(elem);
					break;
				}
				case "cat" : {
					elem = new AscFormat.CCat();
					elem.fromXml(reader);
					this.setCat(elem);
					break;
				}
				case "val" : {
					elem = new AscFormat.CYVal();
					elem.fromXml(reader);
					this.setVal(elem);
					break;
				}
				case "smooth" : {
					this.setSmooth(CT_Bool.prototype.toVal(reader, false));
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}

		if(this.smooth === null){
			this.setSmooth(false);
		}
	};
	CLineSeries.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.idx), "c:idx");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.order), "c:order");
		writer.WriteXmlNullable(this.tx, "c:tx");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.marker, "c:marker");
		writer.WriteXmlArray(this.dPt, "c:dPt");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlArray(this.trendline, "c:trendline");
		writer.WriteXmlNullable(this.errBars, "c:errBars");
		writer.WriteXmlNullable(this.cat, "c:cat");
		writer.WriteXmlNullable(this.val, "c:val");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.smooth), "c:smooth");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CUpDownBars.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "gapWidth" : {
					//todo percent
					this.setGapWidth(CT_Double.prototype.toVal(reader, this.gapWidth));
					break;
				}
				case "upBars" : {
					elem = readChartLines(reader);
					this.setUpBars(elem || new AscFormat.CSpPr());
					break;
				}
				case "downBars" : {
					elem = readChartLines(reader);
					this.setDownBars(elem || new AscFormat.CSpPr());
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CUpDownBars.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.gapWidth), "c:gapWidth");
		writer.WriteXmlNullable(writeChartLines(this.upBars), "c:upBars");
		writer.WriteXmlNullable(writeChartLines(this.downBars), "c:downBars");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	COfPieChart.prototype.fromXml = function(reader, aChartWithAxis) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "ofPieType" : {
					this.setOfPieType(
						fromXml_ST_OfPieType(CT_String.prototype.toVal(reader, this.ofPieType), this.ofPieType));
					break;
				}
				case "varyColors" : {
					this.setVaryColors(CT_Bool.prototype.toVal(reader, this.varyColors));
					break;
				}
				case "ser" : {
					elem = new CPieSeries();
					elem.fromXml(reader);
					this.addSer(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "gapWidth" : {
					//todo percent
					// this.setGapDepth(CT_Double.prototype.toVal(reader, this.gapDepth));
					break;
				}
				case "splitType" : {
					this.setSplitType(fromXml_ST_SplitType(CT_String.prototype.toVal(reader, this.splitType), this.splitType));
					break;
				}
				case "splitPos" : {
					this.setSplitPos(CT_Double.prototype.toVal(reader, this.splitPos));
					break;
				}
				case "custSplit" : {
					let subDepth = reader.GetDepth();
					while (reader.ReadNextSiblingNode(subDepth)) {
						if ("secondPiePt" === reader.GetNameNoNS()) {
							elem = new CT_UInt.prototype.toVal(null);
							if (null !== elem) {
								this.addCustSplit(elem);
							}
						}
					}
					break;
				}
				case "secondPieSize" : {
					//todo percent
					this.setSecondPieSize(CT_Double.prototype.toVal(reader, this.secondPieSize));
					break;
				}
				case "serLines" : {
					elem = readChartLines(reader);
					this.setSerLines(elem || new AscFormat.CSpPr());
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	COfPieChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_OfPieType(this.ofPieType)), "c:ofPieType");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.varyColors), "c:varyColors");
		writer.WriteXmlArray(this.series, "c:ser");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.gapWidth), "c:gapWidth");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_SplitType(this.splitType)), "c:splitType");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.splitPos), "c:splitPos");
		if (this.custSplit.length > 0) {
			this.WriteXmlNodeStart("c:custSplit");
			this.WriteXmlAttributesEnd();
			this.custSplit.forEach(function(elem) {
				writer.WriteXmlNullable(CT_UInt.prototype.fromVal(elem), "c:splitPos");
			}, this);
			this.WriteXmlNodeEnd("c:custSplit");
		}
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.secondPieSize), "c:secondPieSize");
		writer.WriteXmlNullable(writeChartLines(this.serLines), "c:serLines");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CPieChart.prototype.fromXml = function(reader, aChartWithAxis) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "varyColors" : {
					this.setVaryColors(CT_Bool.prototype.toVal(reader, this.varyColors));
					break;
				}
				case "ser" : {
					elem = new CPieSeries();
					elem.fromXml(reader);
					this.addSer(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "firstSliceAng" : {
					this.setFirstSliceAng(CT_UInt.prototype.toVal(reader, this.firstSliceAng));
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CPieChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.varyColors), "c:varyColors");
		writer.WriteXmlArray(this.series, "c:ser");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.firstSliceAng), "c:firstSliceAng");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CRadarChart.prototype.fromXml = function(reader, aChartWithAxis) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "radarStyle" : {
					this.setRadarStyle(
						fromXml_ST_RadarStyle(CT_String.prototype.toVal(reader, this.radarStyle), this.radarStyle));
					break;
				}
				case "varyColors" : {
					this.setVaryColors(CT_Bool.prototype.toVal(reader, this.varyColors));
					break;
				}
				case "ser" : {
					elem = new CRadarSeries();
					elem.fromXml(reader);
					this.addSer(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "axId" : {
					aChartWithAxis.push({ axisId: CT_Int.prototype.toVal(reader, this.axId), chart: this });
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CRadarChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_RadarStyle(this.radarStyle)), "c:radarStyle");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.varyColors), "c:varyColors");
		writer.WriteXmlArray(this.series, "c:ser");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		this.axId.forEach(function(axis) {
			writer.WriteXmlNullable(CT_Int.prototype.fromVal(axis.axId), "c:axId");
		});
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CRadarSeries.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "idx" : {
					this.setIdx(CT_UInt.prototype.toVal(reader, this.idx));
					break;
				}
				case "order" : {
					this.setOrder(CT_UInt.prototype.toVal(reader, this.order));
					break;
				}
				case "tx" : {
					elem = new AscFormat.CTx();
					elem.fromXml(reader);
					this.setTx(elem);
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "marker" : {
					elem = new AscFormat.CMarker();
					elem.fromXml(reader);
					this.setMarker(elem);
					break;
				}
				case "dPt" : {
					elem = new AscFormat.CDPt();
					elem.fromXml(reader);
					this.addDPt(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "cat" : {
					elem = new AscFormat.CCat();
					elem.fromXml(reader);
					this.setCat(elem);
					break;
				}
				case "val" : {
					elem = new AscFormat.CYVal();
					elem.fromXml(reader);
					this.setVal(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CRadarSeries.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.idx), "c:idx");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.order), "c:order");
		writer.WriteXmlNullable(this.tx, "c:tx");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.marker, "c:marker");
		writer.WriteXmlArray(this.dPt, "c:dPt");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlNullable(this.cat, "c:cat");
		writer.WriteXmlNullable(this.val, "c:val");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};
	CScatterChart.prototype.fromXml = function(reader, aChartWithAxis) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "scatterStyle" : {
					this.setScatterStyle(
						fromXml_ST_ScatterStyle(CT_String.prototype.toVal(reader, this.scatterStyle),
							this.scatterStyle));
					break;
				}
				case "varyColors" : {
					this.setVaryColors(CT_Bool.prototype.toVal(reader, this.varyColors));
					break;
				}
				case "ser" : {
					elem = new CScatterSeries();
					elem.fromXml(reader);
					this.addSer(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "axId" : {
					aChartWithAxis.push({ axisId: CT_Int.prototype.toVal(reader, this.axId), chart: this });
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CScatterChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_ScatterStyle(this.scatterStyle)),
			"c:scatterStyle");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.varyColors), "c:varyColors");
		writer.WriteXmlArray(this.series, "c:ser");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		this.axId.forEach(function(axis) {
			writer.WriteXmlNullable(CT_Int.prototype.fromVal(axis.axId), "c:axId");
		});
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};

	CScatterSeries.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "idx" : {
					this.setIdx(CT_UInt.prototype.toVal(reader, this.idx));
					break;
				}
				case "order" : {
					this.setOrder(CT_UInt.prototype.toVal(reader, this.order));
					break;
				}
				case "tx" : {
					elem = new AscFormat.CTx();
					elem.fromXml(reader);
					this.setTx(elem);
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "marker" : {
					elem = new AscFormat.CMarker();
					elem.fromXml(reader);
					this.setMarker(elem);
					break;
				}
				case "dPt" : {
					elem = new AscFormat.CDPt();
					elem.fromXml(reader);
					this.addDPt(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "trendline" : {
					elem = new AscFormat.CTrendLine();
					elem.fromXml(reader);
					this.setTrendline(elem);
					break;
				}
				case "errBars" : {
					elem = new AscFormat.CErrBars();
					elem.fromXml(reader);
					this.setErrBars(elem);
					break;
				}
				case "xVal" : {
					elem = new AscFormat.CCat();
					elem.fromXml(reader);
					this.setXVal(elem);
					break;
				}
				case "yVal" : {
					elem = new AscFormat.CYVal();
					elem.fromXml(reader);
					this.setYVal(elem);
					break;
				}
				case "smooth" : {
					this.setSmooth(CT_Bool.prototype.toVal(reader, true));
					break;
				}
			}
		}
		if(this.smooth === null){
			this.setSmooth(false);
		}
	};
	CScatterSeries.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.idx), "c:idx");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.order), "c:order");
		writer.WriteXmlNullable(this.tx, "c:tx");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.marker, "c:marker");
		writer.WriteXmlArray(this.dPt, "c:dPt");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlArray(this.trendline, "c:trendline");
		writer.WriteXmlArray(this.errBars, "c:errBars");
		writer.WriteXmlNullable(this.xVal, "c:xVal");
		writer.WriteXmlNullable(this.yVal, "c:yVal");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.smooth), "c:smooth");
		writer.WriteXmlNodeEnd(name);
	};

	CStockChart.prototype.fromXml = function(reader, aChartWithAxis) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "ser" : {
					elem = new CLineSeries();
					elem.fromXml(reader);
					this.addSer(elem);
					break;
				}
				case "dLbls" : {
					elem = new AscFormat.CDLbls();
					elem.fromXml(reader);
					elem.correctValues();
					this.setDLbls(elem);
					break;
				}
				case "dropLines" : {
					elem = readChartLines(reader);
					this.setDropLines(elem || new AscFormat.CSpPr());
					break;
				}
				case "hiLowLines" : {
					elem = readChartLines(reader);
					this.setHiLowLines(elem || new AscFormat.CSpPr());
					break;
				}
				case "upDownBars" : {
					elem = new AscFormat.CUpDownBars();
					elem.fromXml(reader);
					this.setUpDownBars(elem);
					break;
				}
				case "axId" : {
					aChartWithAxis.push({ axisId: CT_Int.prototype.toVal(reader, this.axId), chart: this });
					break;
				}
			}
		}
	};
	CStockChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlArray(this.series, "c:ser");
		writer.WriteXmlNullable(this.dLbls, "c:dLbls");
		writer.WriteXmlNullable(writeChartLines(this.dropLines), "c:dropLines");
		writer.WriteXmlNullable(writeChartLines(this.hiLowLines), "c:hiLowLines");
		writer.WriteXmlNullable(this.upDownBars, "c:upDownBars");
		this.axId.forEach(function(axis) {
			writer.WriteXmlNullable(CT_Int.prototype.fromVal(axis.axId), "c:axId");
		});
		writer.WriteXmlNodeEnd(name);
	};

	CSurfaceChart.prototype.fromXml = function(reader, aChartWithAxis) {
		let t = this;
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "wireframe" : {
					this.setWireframe(CT_Bool.prototype.toVal(reader, this.wireframe));
					break;
				}
				case "ser" : {
					elem = new CSurfaceSeries();
					elem.fromXml(reader);
					this.addSer(elem);
					break;
				}
				case "bandFmts" : {
					reader.readXmlArray("bandFmt", function() {
						elem = new AscFormat.CBandFmt();
						elem.fromXml(reader);
						t.addBandFmt(elem);
					});
					break;
				}
				case "axId" : {
					aChartWithAxis.push({ axisId: CT_Int.prototype.toVal(reader, this.axId), chart: this });
					break;
				}
			}
		}
	};
	CSurfaceChart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.wireframe), "c:wireframe");
		writer.WriteXmlArray(this.series, "c:ser");
		writer.WriteXmlNodeStart("c:bandFmts");
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlArray(this.bandFmts, "c:bandFmt");
		writer.WriteXmlNodeEnd("c:bandFmts");
		this.axId.forEach(function(axis) {
			writer.WriteXmlNullable(CT_Int.prototype.fromVal(axis.axId), "c:axId");
		});
		writer.WriteXmlNodeEnd(name);
	};
	CSurfaceSeries.prototype.fromXml = function (reader) {

		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "idx" : {
					this.setIdx(CT_UInt.prototype.toVal(reader, this.idx));
					break;
				}
				case "order" : {
					this.setOrder(CT_UInt.prototype.toVal(reader, this.order));
					break;
				}
				case "tx" : {
					elem = new AscFormat.CTx();
					elem.fromXml(reader);
					this.setTx(elem);
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "cat" : {
					elem = new AscFormat.CCat();
					elem.fromXml(reader);
					this.setCat(elem);
					break;
				}
				case "val" : {
					elem = new AscFormat.CYVal();
					elem.fromXml(reader);
					this.setVal(elem);
					break;
				}
				// case "extLst" : {
				// 	let subDepth = reader.GetDepth();
				// 	while (reader.ReadNextSiblingNode(subDepth)) {
				// 		if ("ext" === reader.GetNameNoNS()) {
				// 			elem = new CT_Extension();
				// 			elem.fromXml(reader);
				// 			this.extLst.push(elem);
				// 		}
				// 	}
				// 	break;
				// }
			}
		}
	};
	CSurfaceSeries.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.idx), "c:idx");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.order), "c:order");
		writer.WriteXmlNullable(this.tx, "c:tx");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNullable(this.cat, "c:cat");
		writer.WriteXmlNullable(this.val, "c:val");
		// writer.WriteXmlArray(this.extLst, "c:extLst");
		writer.WriteXmlNodeEnd(name);
	};
	CBandFmt.prototype.fromXml = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "idx" : {
					this.setIdx(CT_UInt.prototype.toVal(reader, this.idx));
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
			}
		}
	};
	CBandFmt.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.idx), "c:idx");
		writer.WriteXmlNullable(this.spPr, "c:spPr");
		writer.WriteXmlNodeEnd(name);
	};

//Ax
	CValAx.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if (readAxBase(reader, name, this)) {
				continue;
			}
			switch (name) {
				case "majorUnit" : {
					this.setMajorUnit(CT_Double.prototype.toVal(reader, this.majorUnit));
					break;
				}
				case "minorUnit" : {
					this.setMinorUnit(CT_Double.prototype.toVal(reader, this.minorUnit));
					break;
				}
				case "dispUnits" : {
					elem = new AscFormat.CDispUnits();
					elem.fromXml(reader);
					this.setDispUnits(elem);
					break;
				}
			}
		}
	};
	CValAx.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writeAxBase(writer, this);
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.majorUnit), "c:majorUnit");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.minorUnit), "c:minorUnit");
		writer.WriteXmlNullable(this.dispUnits, "c:dispUnits");
		writer.WriteXmlNodeEnd(name);
	};

	CScaling.prototype.fromXml = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "logBase" : {
					this.setLogBase(CT_Double.prototype.toVal(reader, this.logBase));
					break;
				}
				case "orientation" : {
					this.setOrientation(fromXml_ST_Orientation(CT_String.prototype.toVal(reader, this.orientation),
						this.orientation));
					break;
				}
				case "max" : {
					this.setMax(CT_Double.prototype.toVal(reader, this.max));
					break;
				}
				case "min" : {
					this.setMin(CT_Double.prototype.toVal(reader, this.min));
					break;
				}
			}
		}
	};
	CScaling.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.logBase), "c:logBase");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_Orientation(this.orientation)),
			"c:orientation");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.max), "c:max");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.min), "c:min");
		writer.WriteXmlNodeEnd(name);
	};
	CDispUnits.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "custUnit" : {
					this.setCustUnit(CT_Double.prototype.toVal(reader, this.custUnit));
					break;
				}
				case "builtInUnit" : {
					this.setBuiltInUnit(fromXml_ST_BuiltInUnit(CT_String.prototype.toVal(reader, this.builtInUnit),
						this.builtInUnit));
					break;
				}
				case "dispUnitsLbl" : {
					elem = new AscFormat.CDLbl();
					elem.fromXml(reader);
					this.setDispUnitsLbl(elem);
					break;
				}
			}
		}
	};
	CDispUnits.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.custUnit), "c:custUnit");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_BuiltInUnit(this.builtInUnit)),
			"c:builtInUnit");
		writer.WriteXmlNullable(this.dispUnitsLbl, "c:dispUnitsLbl");
		writer.WriteXmlNodeEnd(name);
	};

	CCatAx.prototype.fromXml = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if (readAxBase(reader, name, this)) {
				continue;
			}
			switch (name) {
				case "auto" : {
					this.setAuto(CT_Bool.prototype.toVal(reader, this.auto));
					break;
				}
				case "lblAlgn" : {
					this.setLblAlgn(
						fromXml_ST_LblAlgn(CT_String.prototype.toVal(reader, this.lblAlgn), this.lblAlgn));
					break;
				}
				case "lblOffset" : {
					//todo percent
					this.setLblOffset(CT_Double.prototype.toVal(reader, this.lblOffset));
					break;
				}
				case "tickLblSkip" : {
					this.setTickLblSkip(CT_UInt.prototype.toVal(reader, this.tickLblSkip));
					break;
				}
				case "tickMarkSkip" : {
					this.setTickMarkSkip(CT_UInt.prototype.toVal(reader, this.tickMarkSkip));
					break;
				}
				case "noMultiLvlLbl" : {
					this.setNoMultiLvlLbl(CT_Bool.prototype.toVal(reader, this.noMultiLvlLbl));
					break;
				}
			}
		}
	};
	CCatAx.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writeAxBase(writer, this);
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.auto), "c:auto");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_LblAlgn(this.lblAlgn)), "c:lblAlgn");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.lblOffset), "c:lblOffset");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.tickLblSkip), "c:tickLblSkip");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.tickMarkSkip), "c:tickMarkSkip");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.noMultiLvlLbl), "c:noMultiLvlLbl");
		writer.WriteXmlNodeEnd(name);
	};

	CDateAx.prototype.fromXml = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if (readAxBase(reader, name, this)) {
				continue;
			}
			switch (name) {
				case "auto" : {
					this.setAuto(CT_Bool.prototype.toVal(reader, this.auto));
					break;
				}
				case "lblOffset" : {
					//todo percent
					this.setLblOffset(CT_Double.prototype.toVal(reader, this.lblOffset));
					break;
				}
				case "baseTimeUnit" : {
					this.setBaseTimeUnit(fromXml_ST_TimeUnit(CT_String.prototype.toVal(reader, this.baseTimeUnit),
						this.baseTimeUnit));
					break;
				}
				case "majorUnit" : {
					this.setMajorUnit(CT_Double.prototype.toVal(reader, this.majorUnit));
					break;
				}
				case "majorTimeUnit" : {
					this.setMajorTimeUnit(fromXml_ST_TimeUnit(CT_String.prototype.toVal(reader, this.majorTimeUnit),
						this.majorTimeUnit));
					break;
				}
				case "minorUnit" : {
					this.setMinorUnit(CT_Double.prototype.toVal(reader, this.minorUnit));
					break;
				}
				case "minorTimeUnit" : {
					this.setMinorTimeUnit(fromXml_ST_TimeUnit(CT_String.prototype.toVal(reader, this.minorTimeUnit),
						this.minorTimeUnit));
					break;
				}
			}
		}
	};
	CDateAx.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writeAxBase(writer, this);
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(this.auto), "c:auto");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.lblOffset), "c:lblOffset");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_TimeUnit(this.baseTimeUnit)),
			"c:baseTimeUnit");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.majorUnit), "c:majorUnit");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_TimeUnit(this.majorTimeUnit)),
			"c:majorTimeUnit");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(this.minorUnit), "c:minorUnit");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_TimeUnit(this.minorTimeUnit)),
			"c:minorTimeUnit");
		writer.WriteXmlNodeEnd(name);
	};

	CSerAx.prototype.fromXml = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if (readAxBase(reader, name, this)) {
				continue;
			}
			switch (name) {
				case "tickLblSkip" : {
					this.setTickLblSkip(CT_UInt.prototype.toVal(reader, this.tickLblSkip));
					break;
				}
				case "tickMarkSkip" : {
					this.setTickMarkSkip(CT_UInt.prototype.toVal(reader, this.tickMarkSkip));
					break;
				}
			}
		}
	};
	CSerAx.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writeAxBase(writer, this);
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.tickLblSkip), "c:tickLblSkip");
		writer.WriteXmlNullable(CT_UInt.prototype.fromVal(this.tickMarkSkip), "c:tickMarkSkip");
		writer.WriteXmlNodeEnd(name);
	};

//CChartStyle
	CChartStyle.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "id": {
					this.setId(reader.GetValueUInt(this.id));
					break;
				}
			}
		}
	};
	CChartStyle.prototype.fromXml = function(reader) {
		var name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("chartStyle" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("chartStyle" === name) {
			this.readAttr(reader);
			let elem, depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				switch (reader.GetNameNoNS()) {
					case "axisTitle" : {
						elem = new CStyleEntry();
						elem.setType(1);
						elem.fromXml(reader);
						this.setAxisTitle(elem);
						break;
					}
					case "categoryAxis" : {
						elem = new CStyleEntry();
						elem.setType(2);
						elem.fromXml(reader);
						this.setCategoryAxis(elem);
						break;
					}
					case "chartArea" : {
						elem = new CStyleEntry();
						elem.setType(3);
						elem.fromXml(reader);
						this.setChartArea(elem);
						break;
					}
					case "dataLabel" : {
						elem = new CStyleEntry();
						elem.setType(4);
						elem.fromXml(reader);
						this.setDataLabel(elem);
						break;
					}
					case "dataLabelCallout" : {
						elem = new CStyleEntry();
						elem.setType(5);
						elem.fromXml(reader);
						this.setDataLabelCallout(elem);
						break;
					}
					case "dataPoint" : {
						elem = new CStyleEntry();
						elem.setType(6);
						elem.fromXml(reader);
						this.setDataPoint(elem);
						break;
					}
					case "dataPoint3D" : {
						elem = new CStyleEntry();
						elem.setType(7);
						elem.fromXml(reader);
						this.setDataPoint3D(elem);
						break;
					}
					case "dataPointLine" : {
						elem = new CStyleEntry();
						elem.setType(8);
						elem.fromXml(reader);
						this.setDataPointLine(elem);
						break;
					}
					case "dataPointMarker" : {
						elem = new CStyleEntry();
						elem.setType(9);
						elem.fromXml(reader);
						this.setDataPointMarker(elem);
						break;
					}
					case "dataPointMarkerLayout" : {
						elem = new CMarkerLayout();
						elem.fromXml(reader);
						this.setMarkerLayout(elem);
						break;
					}
					case "dataPointWireframe" : {
						elem = new CStyleEntry();
						elem.setType(10);
						elem.fromXml(reader);
						this.setDataPointWireframe(elem);
						break;
					}
					case "dataTable" : {
						elem = new CStyleEntry();
						elem.setType(11);
						elem.fromXml(reader);
						this.setDataTable(elem);
						break;
					}
					case "downBar" : {
						elem = new CStyleEntry();
						elem.setType(12);
						elem.fromXml(reader);
						this.setDownBar(elem);
						break;
					}
					case "dropLine" : {
						elem = new CStyleEntry();
						elem.setType(13);
						elem.fromXml(reader);
						this.setDropLine(elem);
						break;
					}
					case "errorBar" : {
						elem = new CStyleEntry();
						elem.setType(14);
						elem.fromXml(reader);
						this.setErrorBar(elem);
						break;
					}
					case "floor" : {
						elem = new CStyleEntry();
						elem.setType(15);
						elem.fromXml(reader);
						this.setFloor(elem);
						break;
					}
					case "gridlineMajor" : {
						elem = new CStyleEntry();
						elem.setType(16);
						elem.fromXml(reader);
						this.setGridlineMajor(elem);
						break;
					}
					case "gridlineMinor" : {
						elem = new CStyleEntry();
						elem.setType(17);
						elem.fromXml(reader);
						this.setGridlineMinor(elem);
						break;
					}
					case "hiLoLine" : {
						elem = new CStyleEntry();
						elem.setType(18);
						elem.fromXml(reader);
						this.setHiLoLine(elem);
						break;
					}
					case "leaderLine" : {
						elem = new CStyleEntry();
						elem.setType(19);
						elem.fromXml(reader);
						this.setLeaderLine(elem);
						break;
					}
					case "legend" : {
						elem = new CStyleEntry();
						elem.setType(20);
						elem.fromXml(reader);
						this.setLegend(elem);
						break;
					}
					case "plotArea" : {
						elem = new CStyleEntry();
						elem.setType(21);
						elem.fromXml(reader);
						this.setPlotArea(elem);
						break;
					}
					case "plotArea3D" : {
						elem = new CStyleEntry();
						elem.setType(22);
						elem.fromXml(reader);
						this.setPlotArea3D(elem);
						break;
					}
					case "seriesAxis" : {
						elem = new CStyleEntry();
						elem.setType(23);
						elem.fromXml(reader);
						this.setSeriesAxis(elem);
						break;
					}
					case "seriesLine" : {
						elem = new CStyleEntry();
						elem.setType(24);
						elem.fromXml(reader);
						this.setSeriesLine(elem);
						break;
					}
					case "title" : {
						elem = new CStyleEntry();
						elem.setType(25);
						elem.fromXml(reader);
						this.setTitle(elem);
						break;
					}
					case "trendline" : {
						elem = new CStyleEntry();
						elem.setType(26);
						elem.fromXml(reader);
						this.setTrendline(elem);
						break;
					}
					case "trendlineLabel" : {
						elem = new CStyleEntry();
						elem.setType(27);
						elem.fromXml(reader);
						this.setTrendlineLabel(elem);
						break;
					}
					case "upBar" : {
						elem = new CStyleEntry();
						elem.setType(28);
						elem.fromXml(reader);
						this.setUpBar(elem);
						break;
					}
					case "valueAxis" : {
						elem = new CStyleEntry();
						elem.setType(29);
						elem.fromXml(reader);
						this.setValueAxis(elem);
						break;
					}
					case "wall" : {
						elem = new CStyleEntry();
						elem.setType(30);
						elem.fromXml(reader);
						this.setWall(elem);
						break;
					}
				}
			}
		}
	};
	CChartStyle.prototype.toXml = function(writer) {
		let name = "cs:chartStyle";

		let nOldDocType = writer.context.docType;
		writer.context.docType = AscFormat.XMLWRITER_DOC_TYPE_CHART_STYLE;
		writer.WriteXmlString(AscCommonWord.g_sXmlHeader);
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlString(AscCommonWord.g_sXmlChartStyleNamespaces);
		writer.WriteXmlNullableAttributeUInt("id", this.id);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.axisTitle, "cs:axisTitle");
		writer.WriteXmlNullable(this.categoryAxis, "cs:categoryAxis");
		writer.WriteXmlNullable(this.chartArea, "cs:chartArea");
		writer.WriteXmlNullable(this.dataLabel, "cs:dataLabel");
		writer.WriteXmlNullable(this.dataLabelCallout, "cs:dataLabelCallout");
		writer.WriteXmlNullable(this.dataPoint, "cs:dataPoint");
		writer.WriteXmlNullable(this.dataPoint3D, "cs:dataPoint3D");
		writer.WriteXmlNullable(this.dataPointLine, "cs:dataPointLine");
		writer.WriteXmlNullable(this.dataPointMarker, "cs:dataPointMarker");
		writer.WriteXmlNullable(this.markerLayout, "cs:dataPointMarkerLayout");
		writer.WriteXmlNullable(this.dataPointWireframe, "cs:dataPointWireframe");
		writer.WriteXmlNullable(this.dataTable, "cs:dataTable");
		writer.WriteXmlNullable(this.downBar, "cs:downBar");
		writer.WriteXmlNullable(this.dropLine, "cs:dropLine");
		writer.WriteXmlNullable(this.errorBar, "cs:errorBar");
		writer.WriteXmlNullable(this.floor, "cs:floor");
		writer.WriteXmlNullable(this.gridlineMajor, "cs:gridlineMajor");
		writer.WriteXmlNullable(this.gridlineMinor, "cs:gridlineMinor");
		writer.WriteXmlNullable(this.hiLoLine, "cs:hiLoLine");
		writer.WriteXmlNullable(this.leaderLine, "cs:leaderLine");
		writer.WriteXmlNullable(this.legend, "cs:legend");
		writer.WriteXmlNullable(this.plotArea, "cs:plotArea");
		writer.WriteXmlNullable(this.plotArea3D, "cs:plotArea3D");
		writer.WriteXmlNullable(this.seriesAxis, "cs:seriesAxis");
		writer.WriteXmlNullable(this.seriesLine, "cs:seriesLine");
		writer.WriteXmlNullable(this.title, "cs:title");
		writer.WriteXmlNullable(this.trendline, "cs:trendline");
		writer.WriteXmlNullable(this.trendlineLabel, "cs:trendlineLabel");
		writer.WriteXmlNullable(this.upBar, "cs:upBar");
		writer.WriteXmlNullable(this.valueAxis, "cs:valueAxis");
		writer.WriteXmlNullable(this.wall, "cs:wall");
		writer.WriteXmlNodeEnd(name);
		writer.context.docType = nOldDocType;
	};
	CStyleEntry.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "mods": {
					// this.setMods(reader.GetValueDecodeXml());
					break;
				}
			}
		}
	};
	CStyleEntry.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "lnRef" : {
					this.setLnRef(readStyleRef(reader));
					break;
				}
				case "lineWidthScale" : {
					this.setLineWidthScale(CT_Double.prototype.toVal(reader, this.lineWidthScale));
					break;
				}
				case "fillRef" : {
					this.setFillRef(readStyleRef(reader));
					break;
				}
				case "effectRef" : {
					this.setEffectRef(readStyleRef(reader));
					break;
				}
				case "fontRef" : {
					this.setFontRef(readFontRef(reader));
					break;
				}
				case "spPr" : {
					readSpPr(reader, this);
					break;
				}
				case "defRPr" : {
					this.setDefRPr(readRPr(reader));
					break;
				}
				case "bodyPr" : {
					this.setBodyPr(readBodyPr(reader));
					break;
				}
			}
		}
	};
	CStyleEntry.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		// writer.WriteXmlNullableAttributeStringEncode("mods", this.mods);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.lnRef, "cs:lnRef");
		writer.WriteXmlNullable(this.lineWidthScale, "cs:lineWidthScale");
		writer.WriteXmlNullable(this.fillRef, "cs:fillRef");
		writer.WriteXmlNullable(this.effectRef, "cs:effectRef");
		writer.WriteXmlNullable(this.fontRef, "cs:fontRef");
		writer.WriteXmlNullable(this.spPr, "cs:spPr");
		//todo
		// writer.WriteXmlNullable(this.defRPr, "cs:defRPr");
		// writer.WriteXmlNullable(this.bodyPr, "cs:bodyPr");
		writer.WriteXmlNodeEnd(name);
	};
	CMarkerLayout.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "symbol": {
					this.setSymbol(fromXml_ST_MarkerStyle(reader.GetValue(), this.symbol));
					break;
				}
				case "size": {
					this.setSize(reader.GetValueByte(this.size));
					break;
				}
			}
		}
	};
	CMarkerLayout.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CMarkerLayout.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("symbol", toXml_ST_MarkerStyle(this.symbol));
		writer.WriteXmlNullableAttributeByte("size", this.size);
		writer.WriteXmlAttributesEnd(true);
	};
//CChartColors
	CChartColors.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "meth": {
					this.setMeth(reader.GetValueDecodeXml());
					break;
				}
				case "id": {
					this.setId(reader.GetValueUInt(this.id));
					break;
				}
			}
		}
	};
	CChartColors.prototype.fromXml = function(reader) {
		var name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("colorStyle" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("colorStyle" === name) {
			this.readAttr(reader);
			let elem, depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				let sName  = reader.GetNameNoNS();
				if(sName === "variation") {
					elem = readColorModifier(reader);
					this.addItem(elem);
				}
				else {
					elem = readUniColor(reader, sName);
					this.addItem(elem);
				}
			}
		}
	};
	CChartColors.prototype.toXml = function(writer) {
		let name = "cs:colorStyle";
		writer.WriteXmlString(AscCommonWord.g_sXmlHeader);
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlString(AscCommonWord.g_sXmlChartColorNamespaces);
		writer.WriteXmlNullableAttributeStringEncode("meth", this.meth);
		writer.WriteXmlNullableAttributeUInt("id", this.id);
		writer.WriteXmlAttributesEnd();
		this.items.forEach(function(item){
			if (item instanceof AscFormat.CUniColor) {
				item.toXml(writer);
			}
			else {
				//todo
				// writer.WriteXmlNullable(item, "cs:variation");
			}
		});
		writer.WriteXmlNodeEnd(name);
	};

	function readLayout(reader, parent) {
		let elem = new CT_XmlNode(function(reader, name) {
			if ("manualLayout" === name) {
				var layout = new AscFormat.CLayout();
				layout.fromXml(reader);
				return layout;
			}
		});
		elem.fromXml(reader);
		if (elem.members["manualLayout"]) {
			parent.setLayout(elem.members["manualLayout"]);
		}
	}

	function writeLayout(layout) {
		let res = new CT_XmlNode();
		if (layout) {
			res.members["c:manualLayout"] = layout;
		}
		return res;
	}

	function readChartLines(reader) {
		let elem = new CT_XmlNode(function(reader, name) {
			if ("spPr" === name) {
				return readSpPrPlain(reader);
			}
		});
		elem.fromXml(reader);
		return elem.members["spPr"];
	}

	function writeChartLines(lines) {
		let res = null;
		if (lines) {
			res = new CT_XmlNode();
			res.members["c:spPr"] = lines;
		}
		return res;
	}

	function readAxBase(reader, name, ax) {
		let elem;
		switch (name) {
			case "axId" : {
				ax.setAxId(CT_Int.prototype.toVal(reader, ax.axId));
				break;
			}
			case "scaling" : {
				elem = new AscFormat.CScaling();
				elem.fromXml(reader);
				ax.setScaling(elem);
				break;
			}
			case "delete" : {
				ax.setDelete(CT_Bool.prototype.toVal(reader, ax.delete));
				break;
			}
			case "axPos" : {
				ax.setAxPos(fromXml_ST_AxPos(CT_String.prototype.toVal(reader, ax.axPos), ax.axPos));
				break;
			}
			case "majorGridlines" : {
				elem = readChartLines(reader);
				ax.setMajorGridlines(elem || new AscFormat.CSpPr());
				break;
			}
			case "minorGridlines" : {
				elem = readChartLines(reader);
				ax.setMinorGridlines(elem || new AscFormat.CSpPr());
				break;
			}
			case "title" : {
				elem = new AscFormat.CTitle();
				elem.fromXml(reader);
				ax.setTitle(elem);
				break;
			}
			case "numFmt" : {
				elem = new AscFormat.CNumFmt();
				elem.fromXml(reader);
				ax.setNumFmt(elem);
				break;
			}
			case "majorTickMark" : {
				ax.setMajorTickMark(
					fromXml_ST_TickMark(CT_String.prototype.toVal(reader, ax.majorTickMark), ax.majorTickMark));
				break;
			}
			case "minorTickMark" : {
				ax.setMinorTickMark(
					fromXml_ST_TickMark(CT_String.prototype.toVal(reader, ax.majorTickMark), ax.majorTickMark));
				break;
			}
			case "tickLblPos" : {
				ax.setTickLblPos(
					fromXml_ST_TickLblPos(CT_String.prototype.toVal(reader, ax.tickLblPos), ax.tickLblPos));
				break;
			}
			case "spPr" : {
				readSpPr(reader, ax);
				break;
			}
			case "txPr" : {
				readTxPr(reader, ax);
				break;
			}
			case "crossAx" : {
				ax.crossAxId = CT_Int.prototype.toVal(reader, ax.crossAxId);
				break;
			}
			case "crosses" : {
				ax.setCrosses(fromXml_ST_Crosses(CT_String.prototype.toVal(reader, ax.crosses), ax.crosses));
				break;
			}
			case "crossesAt" : {
				ax.setCrossesAt(CT_Double.prototype.toVal(reader, ax.crossesAt));
				break;
			}
			case "crossBetween" : {
				ax.setCrossBetween(
					fromXml_ST_CrossBetween(CT_String.prototype.toVal(reader, ax.crossBetween), ax.crossBetween));
				break;
			}
		}

	}

	function writeAxBase(writer, ax) {
		writer.WriteXmlNullable(CT_Int.prototype.fromVal(ax.axId), "c:axId");
		writer.WriteXmlNullable(ax.scaling, "c:scaling");
		writer.WriteXmlNullable(CT_Bool.prototype.fromVal(ax.bDelete), "c:delete");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_AxPos(ax.axPos)), "c:axPos");
		writer.WriteXmlNullable(writeChartLines(ax.majorGridlines), "c:majorGridlines");
		writer.WriteXmlNullable(writeChartLines(ax.minorGridlines), "c:minorGridlines");
		writer.WriteXmlNullable(ax.title, "c:title");
		writer.WriteXmlNullable(ax.numFmt, "c:numFmt");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_TickMark(ax.majorTickMark)),
			"c:majorTickMark");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_TickMark(ax.minorTickMark)),
			"c:minorTickMark");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_TickLblPos(ax.tickLblPos)), "c:tickLblPos");
		writer.WriteXmlNullable(ax.spPr, "c:spPr");
		writer.WriteXmlNullable(ax.txPr, "c:txPr");
		writer.WriteXmlNullable(CT_Int.prototype.fromVal(ax.crossAx && ax.crossAx.axId), "c:crossAx");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_Crosses(ax.crosses)), "c:crosses");
		writer.WriteXmlNullable(CT_Double.prototype.fromVal(ax.crossesAt), "c:crossesAt");
		writer.WriteXmlNullable(CT_String.prototype.fromVal(toXml_ST_CrossBetween(ax.crossBetween)),
			"c:crossBetween");
	}
	function CorrectChartWithAxis(chartOld, chartNew, aChartWithAxis) {
		for (var i = 0, length = aChartWithAxis.length; i < length; ++i) {
			var item = aChartWithAxis[i];
			if (item.chart == chartOld)
				item.chart = chartNew;
		}
	}

	function fromXml_ST_LayoutTarget(val, def) {
		switch (val) {
			case "inner":
				return AscFormat.LAYOUT_TARGET_INNER;
			case "outer":
				return AscFormat.LAYOUT_TARGET_OUTER;
		}
		return def;
	}

	function toXml_ST_LayoutTarget(val) {
		switch (val) {
			case AscFormat.LAYOUT_TARGET_INNER:
				return "inner";
			case AscFormat.LAYOUT_TARGET_OUTER:
				return "outer";
		}
		return null;
	}
	function fromXml_ST_LayoutMode(val, def) {
		switch (val) {
			case "edge":
				return AscFormat.LAYOUT_MODE_EDGE;
			case "factor":
				return AscFormat.LAYOUT_MODE_FACTOR;
		}
		return def;
	}

	function toXml_ST_LayoutMode(val) {
		switch (val) {
			case AscFormat.LAYOUT_MODE_EDGE:
				return "edge";
			case AscFormat.LAYOUT_MODE_FACTOR:
				return "factor";
		}
		return null;
	}
	function fromXml_ST_PageSetupOrientation(val, def) {
		switch (val) {
			case "default":
				return AscFormat.PAGE_SETUP_ORIENTATION_DEFAULT;
			case "portrait":
				return AscFormat.PAGE_SETUP_ORIENTATION_PORTRAIT;
			case "landscape":
				return AscFormat.PAGE_SETUP_ORIENTATION_LANDSCAPE;
		}
		return def;
	}

	function toXml_ST_PageSetupOrientation(val) {
		switch (val) {
			case AscFormat.PAGE_SETUP_ORIENTATION_DEFAULT:
				return "default";
			case AscFormat.PAGE_SETUP_ORIENTATION_PORTRAIT:
				return "portrait";
			case AscFormat.PAGE_SETUP_ORIENTATION_LANDSCAPE:
				return "landscape";
		}
		return null;
	}

	function fromXml_ST_MarkerStyle(val, def) {
		switch (val) {
			case "circle":
				return AscFormat.SYMBOL_CIRCLE;
			case "dash":
				return AscFormat.SYMBOL_DASH;
			case "diamond":
				return AscFormat.SYMBOL_DIAMOND;
			case "dot":
				return AscFormat.SYMBOL_DOT;
			case "none":
				return AscFormat.SYMBOL_NONE;
			case "picture":
				return AscFormat.SYMBOL_PICTURE;
			case "plus":
				return AscFormat.SYMBOL_PLUS;
			case "square":
				return AscFormat.SYMBOL_SQUARE;
			case "star":
				return AscFormat.SYMBOL_STAR;
			case "triangle":
				return AscFormat.SYMBOL_TRIANGLE;
			case "x":
				return AscFormat.SYMBOL_X;
			case "auto":
				return null;
		}
		return def;
	}

	function toXml_ST_MarkerStyle(val) {
		switch (val) {
			case AscFormat.SYMBOL_CIRCLE:
				return "circle";
			case AscFormat.SYMBOL_DASH:
				return "dash";
			case AscFormat.SYMBOL_DIAMOND:
				return "diamond";
			case AscFormat.SYMBOL_DOT:
				return "dot";
			case AscFormat.SYMBOL_NONE:
				return "none";
			case AscFormat.SYMBOL_PICTURE:
				return "picture";
			case AscFormat.SYMBOL_PLUS:
				return "plus";
			case AscFormat.SYMBOL_SQUARE:
				return "square";
			case AscFormat.SYMBOL_STAR:
				return "star";
			case AscFormat.SYMBOL_TRIANGLE:
				return "triangle";
			case AscFormat.SYMBOL_X:
				return "x";
			// case null:
			// 	return "auto";
		}
		return null;
	}

	function fromXml_ST_PictureFormat(val, def) {
		switch (val) {
			case "stretch":
				return AscFormat.PICTURE_FORMAT_STACK_STRETCH;
			case "stack":
				return AscFormat.PICTURE_FORMAT_STACK;
			case "stackScale":
				return AscFormat.PICTURE_FORMAT_STACK_SCALE;
		}
		return def;
	}

	function toXml_ST_PictureFormat(val) {
		switch (val) {
			case AscFormat.PICTURE_FORMAT_STACK_STRETCH:
				return "stretch";
			case AscFormat.PICTURE_FORMAT_STACK:
				return "stack";
			case AscFormat.PICTURE_FORMAT_STACK_SCALE:
				return "stackScale";
		}
		return null;
	}

	function fromXml_ST_LegendPos(val, def) {
		switch (val) {
			case "b":
				return Asc.c_oAscChartLegendShowSettings.bottom;
			case "tr":
				return Asc.c_oAscChartLegendShowSettings.topRight;
			case "l":
				return Asc.c_oAscChartLegendShowSettings.left;
			case "r":
				return Asc.c_oAscChartLegendShowSettings.right;
			case "t":
				return Asc.c_oAscChartLegendShowSettings.top;
		}
		return def;
	}

	function toXml_ST_LegendPos(val) {
		switch (val) {
			case Asc.c_oAscChartLegendShowSettings.bottom:
				return "b";
			case Asc.c_oAscChartLegendShowSettings.topRight:
				return "tr";
			case Asc.c_oAscChartLegendShowSettings.left:
				return "l";
			case Asc.c_oAscChartLegendShowSettings.right:
				return "r";
			case Asc.c_oAscChartLegendShowSettings.top:
				return "t";
		}
		return null;
	}

	function fromXml_ST_DispBlanksAs(val, def) {
		switch (val) {
			case "span":
				return AscFormat.DISP_BLANKS_AS_SPAN;
			case "gap":
				return AscFormat.DISP_BLANKS_AS_GAP;
			case "zero":
				return AscFormat.DISP_BLANKS_AS_ZERO;
		}
		return def;
	}

	function fromXml_ST_Grouping(val, def) {
		switch (val) {
			case "percentStacked":
				return AscFormat.GROUPING_PERCENT_STACKED;
			case "standard":
				return AscFormat.GROUPING_STANDARD;
			case "stacked":
				return AscFormat.GROUPING_STACKED;
		}
		return def;
	}

	function toXml_ST_Grouping(val) {
		switch (val) {
			case AscFormat.GROUPING_PERCENT_STACKED:
				return "percentStacked";
			case AscFormat.GROUPING_STANDARD:
				return "standard";
			case AscFormat.GROUPING_STACKED:
				return "stacked";
		}
		return null;
	}

	function toXml_ST_DispBlanksAs(val) {
		switch (val) {
			case AscFormat.DISP_BLANKS_AS_SPAN:
				return "span";
			case AscFormat.DISP_BLANKS_AS_GAP:
				return "gap";
			case AscFormat.DISP_BLANKS_AS_ZERO:
				return "zero";
		}
		return null;
	}

	function fromXml_ST_DLblPos(val, def) {
		switch (val) {
			case "bestFit":
				return Asc.c_oAscChartDataLabelsPos.bestFit;
			case "b":
				return Asc.c_oAscChartDataLabelsPos.b;
			case "ctr":
				return Asc.c_oAscChartDataLabelsPos.ctr;
			case "inBase":
				return Asc.c_oAscChartDataLabelsPos.inBase;
			case "inEnd":
				return Asc.c_oAscChartDataLabelsPos.inEnd;
			case "l":
				return Asc.c_oAscChartDataLabelsPos.l;
			case "outEnd":
				return Asc.c_oAscChartDataLabelsPos.outEnd;
			case "r":
				return Asc.c_oAscChartDataLabelsPos.r;
			case "t":
				return Asc.c_oAscChartDataLabelsPos.t;
		}
		return def;
	}

	function toXml_ST_DLblPos(val) {
		switch (val) {
			case Asc.c_oAscChartDataLabelsPos.bestFit:
				return "bestFit";
			case Asc.c_oAscChartDataLabelsPos.b:
				return "b";
			case Asc.c_oAscChartDataLabelsPos.ctr:
				return "ctr";
			case Asc.c_oAscChartDataLabelsPos.inBase:
				return "inBase";
			case Asc.c_oAscChartDataLabelsPos.inEnd:
				return "inEnd";
			case Asc.c_oAscChartDataLabelsPos.l:
				return "l";
			case Asc.c_oAscChartDataLabelsPos.outEnd:
				return "outEnd";
			case Asc.c_oAscChartDataLabelsPos.r:
				return "r";
			case Asc.c_oAscChartDataLabelsPos.t:
				return "t";
		}
		return null;
	}

	function fromXml_ST_TrendlineType(val, def) {
		switch (val) {
			case "exp":
				return AscFormat.TRENDLINE_TYPE_EXP;
			case "linear":
				return AscFormat.TRENDLINE_TYPE_LINEAR;
			case "log":
				return AscFormat.TRENDLINE_TYPE_LOG;
			case "movingAvg":
				return AscFormat.TRENDLINE_TYPE_MOVING_AVG;
			case "poly":
				return AscFormat.TRENDLINE_TYPE_POLY;
			case "power":
				return AscFormat.TRENDLINE_TYPE_POWER;
		}
		return def;
	}

	function toXml_ST_TrendlineType(val) {
		switch (val) {
			case AscFormat.TRENDLINE_TYPE_EXP:
				return "exp";
			case AscFormat.TRENDLINE_TYPE_LINEAR:
				return "linear";
			case AscFormat.TRENDLINE_TYPE_LOG:
				return "log";
			case AscFormat.TRENDLINE_TYPE_MOVING_AVG:
				return "movingAvg";
			case AscFormat.TRENDLINE_TYPE_POLY:
				return "poly";
			case AscFormat.TRENDLINE_TYPE_POWER:
				return "power";
		}
		return null;
	}

	function fromXml_ST_ErrDir(val, def) {
		switch (val) {
			case "x":
				return AscFormat.ERR_DIR_X;
			case "y":
				return AscFormat.ERR_DIR_Y;
		}
		return def;
	}

	function toXml_ST_ErrDir(val) {
		switch (val) {
			case AscFormat.ERR_DIR_X:
				return "x";
			case AscFormat.ERR_DIR_Y:
				return "y";
		}
		return null;
	}

	function fromXml_ST_ErrBarType(val, def) {
		switch (val) {
			case "both":
				return AscFormat.ERR_BAR_TYPE_BOTH;
			case "minus":
				return AscFormat.ERR_BAR_TYPE_MINUS;
			case "plus":
				return AscFormat.ERR_BAR_TYPE_PLUS;
		}
		return def;
	}

	function toXml_ST_ErrBarType(val) {
		switch (val) {
			case AscFormat.ERR_BAR_TYPE_BOTH:
				return "both";
			case AscFormat.ERR_BAR_TYPE_MINUS:
				return "minus";
			case AscFormat.ERR_BAR_TYPE_PLUS:
				return "plus";
		}
		return null;
	}

	function fromXml_ST_ErrValType(val, def) {
		switch (val) {
			case "cust":
				return AscFormat.ERR_VAL_TYPE_CUST;
			case "fixedVal":
				return AscFormat.ERR_VAL_TYPE_FIXED_VAL;
			case "percentage":
				return AscFormat.ERR_VAL_TYPE_PERCENTAGE;
			case "stdDev":
				return AscFormat.ERR_VAL_TYPE_STD_DEV;
			case "stdErr":
				return AscFormat.ERR_VAL_TYPE_STD_ERR;
		}
		return def;
	}

	function toXml_ST_ErrValType(val) {
		switch (val) {
			case AscFormat.ERR_VAL_TYPE_CUST:
				return "cust";
			case AscFormat.ERR_VAL_TYPE_FIXED_VAL:
				return "fixedVal";
			case AscFormat.ERR_VAL_TYPE_PERCENTAGE:
				return "percentage";
			case AscFormat.ERR_VAL_TYPE_STD_DEV:
				return "stdDev";
			case AscFormat.ERR_VAL_TYPE_STD_ERR:
				return "stdErr";
		}
		return null;
	}

	function fromXml_ST_BarDir(val, def) {
		switch (val) {
			case "bar":
				return AscFormat.BAR_DIR_BAR;
			case "col":
				return AscFormat.BAR_DIR_COL;
		}
		return def;
	}

	function toXml_ST_BarDir(val) {
		switch (val) {
			case AscFormat.BAR_DIR_BAR:
				return "bar";
			case AscFormat.BAR_DIR_COL:
				return "col";
		}
		return null;
	}

	function fromXml_ST_BarGrouping(val, def) {
		switch (val) {
			case "percentStacked":
				return AscFormat.BAR_GROUPING_PERCENT_STACKED;
			case "clustered":
				return AscFormat.BAR_GROUPING_CLUSTERED;
			case "standard":
				return AscFormat.BAR_GROUPING_STANDARD;
			case "stacked":
				return AscFormat.BAR_GROUPING_STACKED;
		}
		return def;
	}

	function toXml_ST_BarGrouping(val) {
		switch (val) {
			case AscFormat.BAR_GROUPING_PERCENT_STACKED:
				return "percentStacked";
			case AscFormat.BAR_GROUPING_CLUSTERED:
				return "clustered";
			case AscFormat.BAR_GROUPING_STANDARD:
				return "standard";
			case AscFormat.BAR_GROUPING_STACKED:
				return "stacked";
		}
		return null;
	}

	function fromXml_ST_Shape(val, def) {
		switch (val) {
			case "cone":
				return AscFormat.BAR_SHAPE_CONE;
			case "coneToMax":
				return AscFormat.BAR_SHAPE_CONETOMAX;
			case "box":
				return AscFormat.BAR_SHAPE_BOX;
			case "cylinder":
				return AscFormat.BAR_SHAPE_CYLINDER;
			case "pyramid":
				return AscFormat.BAR_SHAPE_PYRAMID;
			case "pyramidToMax":
				return AscFormat.BAR_SHAPE_PYRAMIDTOMAX;
		}
		return def;
	}

	function toXml_ST_Shape(val) {
		switch (val) {
			case AscFormat.BAR_SHAPE_CONE:
				return "cone";
			case AscFormat.BAR_SHAPE_CONETOMAX:
				return "coneToMax";
			case AscFormat.BAR_SHAPE_BOX:
				return "box";
			case AscFormat.BAR_SHAPE_CYLINDER:
				return "cylinder";
			case AscFormat.BAR_SHAPE_PYRAMID:
				return "pyramid";
			case AscFormat.BAR_SHAPE_PYRAMIDTOMAX:
				return "pyramidToMax";
		}
		return null;
	}

	function fromXml_ST_SizeRepresents(val, def) {
		switch (val) {
			case "area":
				return AscFormat.SIZE_REPRESENTS_AREA;
			case "w":
				return AscFormat.SIZE_REPRESENTS_W;
		}
		return def;
	}

	function toXml_ST_SizeRepresents(val) {
		switch (val) {
			case AscFormat.SIZE_REPRESENTS_AREA:
				return "area";
			case AscFormat.SIZE_REPRESENTS_W:
				return "w";
		}
		return null;
	}

	function fromXml_ST_OfPieType(val, def) {
		switch (val) {
			case "pie":
				return AscFormat.OF_PIE_TYPE_PIE;
			case "bar":
				return AscFormat.OF_PIE_TYPE_BAR;
		}
		return def;
	}

	function toXml_ST_OfPieType(val) {
		switch (val) {
			case AscFormat.OF_PIE_TYPE_PIE:
				return "pie";
			case AscFormat.OF_PIE_TYPE_BAR:
				return "bar";
		}
		return null;
	}

	function fromXml_ST_SplitType(val, def) {
		switch (val) {
			case "auto":
				return AscFormat.SPLIT_TYPE_AUTO;
			case "cust":
				return AscFormat.SPLIT_TYPE_CUST;
			case "percent":
				return AscFormat.SPLIT_TYPE_PERCENT;
			case "pos":
				return AscFormat.SPLIT_TYPE_POS;
			case "val":
				return AscFormat.SPLIT_TYPE_VAL;
		}
		return def;
	}

	function toXml_ST_SplitType(val) {
		switch (val) {
			case AscFormat.SPLIT_TYPE_AUTO:
				return "auto";
			case AscFormat.SPLIT_TYPE_CUST:
				return "cust";
			case AscFormat.SPLIT_TYPE_PERCENT:
				return "percent";
			case AscFormat.SPLIT_TYPE_POS:
				return "pos";
			case AscFormat.SPLIT_TYPE_VAL:
				return "val";
		}
		return null;
	}

	function fromXml_ST_RadarStyle(val, def) {
		switch (val) {
			case "standard":
				return AscFormat.RADAR_STYLE_STANDARD;
			case "marker":
				return AscFormat.RADAR_STYLE_MARKER;
			case "filled":
				return AscFormat.RADAR_STYLE_FILLED;
		}
		return def;
	}

	function toXml_ST_RadarStyle(val) {
		switch (val) {
			case AscFormat.RADAR_STYLE_STANDARD:
				return "standard";
			case AscFormat.RADAR_STYLE_MARKER:
				return "marker";
			case AscFormat.RADAR_STYLE_FILLED:
				return "filled";
		}
		return null;
	}

	function fromXml_ST_ScatterStyle(val, def) {
		switch (val) {
			case "none":
				return AscFormat.SCATTER_STYLE_NONE;
			case "line":
				return AscFormat.SCATTER_STYLE_LINE;
			case "lineMarker":
				return AscFormat.SCATTER_STYLE_LINE_MARKER;
			case "marker":
				return AscFormat.SCATTER_STYLE_MARKER;
			case "smooth":
				return AscFormat.SCATTER_STYLE_SMOOTH;
			case "smoothMarker":
				return AscFormat.SCATTER_STYLE_SMOOTH_MARKER;
		}
		return def;
	}

	function toXml_ST_ScatterStyle(val) {
		switch (val) {
			case AscFormat.SCATTER_STYLE_NONE:
				return "none";
			case AscFormat.SCATTER_STYLE_LINE:
				return "line";
			case AscFormat.SCATTER_STYLE_LINE_MARKER:
				return "lineMarker";
			case AscFormat.SCATTER_STYLE_MARKER:
				return "marker";
			case AscFormat.SCATTER_STYLE_SMOOTH:
				return "smooth";
			case AscFormat.SCATTER_STYLE_SMOOTH_MARKER:
				return "smoothMarker";
		}
		return null;
	}

	function fromXml_ST_Orientation(val, def) {
		switch (val) {
			case "maxMin":
				return AscFormat.ORIENTATION_MAX_MIN;
			case "minMax":
				return AscFormat.ORIENTATION_MIN_MAX;
		}
		return def;
	}

	function toXml_ST_Orientation(val) {
		switch (val) {
			case AscFormat.ORIENTATION_MAX_MIN:
				return "maxMin";
			case AscFormat.ORIENTATION_MIN_MAX:
				return "minMax";
		}
		return null;
	}

	function fromXml_ST_BuiltInUnit(val, def) {
		switch (val) {
			case "hundreds":
				return Asc.c_oAscValAxUnits.HUNDREDS;
			case "thousands":
				return Asc.c_oAscValAxUnits.THOUSANDS;
			case "tenThousands":
				return Asc.c_oAscValAxUnits.TEN_THOUSANDS;
			case "hundredThousands":
				return Asc.c_oAscValAxUnits.HUNDRED_THOUSANDS;
			case "millions":
				return Asc.c_oAscValAxUnits.MILLIONS;
			case "tenMillions":
				return Asc.c_oAscValAxUnits.TEN_MILLIONS;
			case "hundredMillions":
				return Asc.c_oAscValAxUnits.HUNDRED_MILLIONS;
			case "billions":
				return Asc.c_oAscValAxUnits.BILLIONS;
			case "trillions":
				return Asc.c_oAscValAxUnits.TRILLIONS;
		}
		return def;
	}

	function toXml_ST_BuiltInUnit(val) {
		switch (val) {
			case Asc.c_oAscValAxUnits.HUNDREDS:
				return "hundreds";
			case Asc.c_oAscValAxUnits.THOUSANDS:
				return "thousands";
			case Asc.c_oAscValAxUnits.TEN_THOUSANDS:
				return "tenThousands";
			case Asc.c_oAscValAxUnits.HUNDRED_THOUSANDS:
				return "hundredThousands";
			case Asc.c_oAscValAxUnits.MILLIONS:
				return "millions";
			case Asc.c_oAscValAxUnits.TEN_MILLIONS:
				return "tenMillions";
			case Asc.c_oAscValAxUnits.HUNDRED_MILLIONS:
				return "hundredMillions";
			case Asc.c_oAscValAxUnits.BILLIONS:
				return "billions";
			case Asc.c_oAscValAxUnits.TRILLIONS:
				return "trillions";
		}
		return null;
	}

	function fromXml_ST_LblAlgn(val, def) {
		switch (val) {
			case "ctr":
				return AscFormat.LBL_ALG_CTR;
			case "l":
				return AscFormat.LBL_ALG_L;
			case "r":
				return AscFormat.LBL_ALG_R;
		}
		return def;
	}

	function toXml_ST_LblAlgn(val) {
		switch (val) {
			case AscFormat.LBL_ALG_CTR:
				return "ctr";
			case AscFormat.LBL_ALG_L:
				return "l";
			case AscFormat.LBL_ALG_R:
				return "r";
		}
		return null;
	}

	function fromXml_ST_TimeUnit(val, def) {
		switch (val) {
			case "days":
				return AscFormat.TIME_UNIT_DAYS;
			case "months":
				return AscFormat.TIME_UNIT_MONTHS;
			case "years":
				return AscFormat.TIME_UNIT_YEARS;
		}
		return def;
	}

	function toXml_ST_TimeUnit(val) {
		switch (val) {
			case AscFormat.TIME_UNIT_DAYS:
				return "days";
			case AscFormat.TIME_UNIT_MONTHS:
				return "months";
			case AscFormat.TIME_UNIT_YEARS:
				return "years";
		}
		return null;
	}

	function fromXml_ST_AxPos(val, def) {
		switch (val) {
			case "b":
				return AscFormat.AX_POS_B;
			case "l":
				return AscFormat.AX_POS_L;
			case "r":
				return AscFormat.AX_POS_R;
			case "t":
				return AscFormat.AX_POS_T;
		}
		return def;
	}

	function toXml_ST_AxPos(val) {
		switch (val) {
			case AscFormat.AX_POS_B:
				return "b";
			case AscFormat.AX_POS_L:
				return "l";
			case AscFormat.AX_POS_R:
				return "r";
			case AscFormat.AX_POS_T:
				return "t";
		}
		return null;
	}

	function fromXml_ST_TickMark(val, def) {
		switch (val) {
			case "cross":
				return Asc.c_oAscTickMark.TICK_MARK_CROSS;
			case "in":
				return Asc.c_oAscTickMark.TICK_MARK_IN;
			case "none":
				return Asc.c_oAscTickMark.TICK_MARK_NONE;
			case "out":
				return Asc.c_oAscTickMark.TICK_MARK_OUT;
		}
		return def;
	}

	function toXml_ST_TickMark(val) {
		switch (val) {
			case Asc.c_oAscTickMark.TICK_MARK_CROSS:
				return "cross";
			case Asc.c_oAscTickMark.TICK_MARK_IN:
				return "in";
			case Asc.c_oAscTickMark.TICK_MARK_NONE:
				return "none";
			case Asc.c_oAscTickMark.TICK_MARK_OUT:
				return "out";
		}
		return null;
	}

	function fromXml_ST_TickLblPos(val, def) {
		switch (val) {
			case "high":
				return Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_HIGH;
			case "low":
				return Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_LOW;
			case "nextTo":
				return Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO;
			case "none":
				return Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE;
		}
		return def;
	}

	function toXml_ST_TickLblPos(val) {
		switch (val) {
			case Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_HIGH:
				return "high";
			case Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_LOW:
				return "low";
			case Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO:
				return "nextTo";
			case Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE:
				return "none";
		}
		return null;
	}

	function fromXml_ST_Crosses(val, def) {
		switch (val) {
			case "autoZero":
				return AscFormat.CROSSES_AUTO_ZERO;
			case "max":
				return AscFormat.CROSSES_MAX;
			case "min":
				return AscFormat.CROSSES_MIN;
		}
		return def;
	}

	function toXml_ST_Crosses(val) {
		switch (val) {
			case AscFormat.CROSSES_AUTO_ZERO:
				return "autoZero";
			case AscFormat.CROSSES_MAX:
				return "max";
			case AscFormat.CROSSES_MIN:
				return "min";
		}
		return null;
	}

	function fromXml_ST_CrossBetween(val, def) {
		switch (val) {
			case "between":
				return AscFormat.CROSS_BETWEEN_BETWEEN;
			case "midCat":
				return AscFormat.CROSS_BETWEEN_MID_CAT;
		}
		return def;
	}

	function toXml_ST_CrossBetween(val) {
		switch (val) {
			case AscFormat.CROSS_BETWEEN_BETWEEN:
				return "between";
			case AscFormat.CROSS_BETWEEN_MID_CAT:
				return "midCat";
		}
		return null;
	}
})(window);

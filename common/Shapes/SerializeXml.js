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

/**
 * @param {Window} window
 * @param {undefined} undefined
 */
(function (window, undefined) {


	let XMLWRITER_DOC_TYPE_PPTX				 = 0;
	let XMLWRITER_DOC_TYPE_DOCX				 = 1;
	let XMLWRITER_DOC_TYPE_XLSX				 = 2;
	let XMLWRITER_DOC_TYPE_CHART			 = 3;
	let XMLWRITER_DOC_TYPE_WORDART			 = 4;
	let XMLWRITER_DOC_TYPE_GRAPHICS			 = 5;
	let XMLWRITER_DOC_TYPE_CHART_DRAWING	 = 6;
	let XMLWRITER_DOC_TYPE_DOCX_GLOSSARY	 = 7;
	let XMLWRITER_DOC_TYPE_DIAGRAM			 = 8;
	let XMLWRITER_DOC_TYPE_DSP_DRAWING		 = 9;
	let XMLWRITER_DOC_TYPE_CHART_STYLE		 = 10;

	let XMLWRITER_RECORD_TYPE_SPPR			 = 0;
	let XMLWRITER_RECORD_TYPE_CLRMAPOVR		 = 1;
	let XMLWRITER_RECORD_TYPE_TEXT_OUTLINE	 = 2;
	let XMLWRITER_RECORD_TYPE_TEXT_FILL		 = 3;
	window["AscFormat"] = window["AscFormat"] || {};
	let AscFormat = window["AscFormat"]
	AscFormat.XMLWRITER_DOC_TYPE_PPTX = XMLWRITER_DOC_TYPE_PPTX;
	AscFormat.XMLWRITER_DOC_TYPE_DOCX = XMLWRITER_DOC_TYPE_DOCX;
	AscFormat.XMLWRITER_DOC_TYPE_XLSX = XMLWRITER_DOC_TYPE_XLSX;
	AscFormat.XMLWRITER_DOC_TYPE_CHART = XMLWRITER_DOC_TYPE_CHART;
	AscFormat.XMLWRITER_DOC_TYPE_WORDART = XMLWRITER_DOC_TYPE_WORDART;
	AscFormat.XMLWRITER_DOC_TYPE_GRAPHICS = XMLWRITER_DOC_TYPE_GRAPHICS;
	AscFormat.XMLWRITER_DOC_TYPE_CHART_DRAWING = XMLWRITER_DOC_TYPE_CHART_DRAWING;
	AscFormat.XMLWRITER_DOC_TYPE_DOCX_GLOSSARY = XMLWRITER_DOC_TYPE_DOCX_GLOSSARY;
	AscFormat.XMLWRITER_DOC_TYPE_DIAGRAM = XMLWRITER_DOC_TYPE_DIAGRAM;
	AscFormat.XMLWRITER_DOC_TYPE_DSP_DRAWING = XMLWRITER_DOC_TYPE_DSP_DRAWING;
	AscFormat.XMLWRITER_DOC_TYPE_CHART_STYLE = XMLWRITER_DOC_TYPE_CHART_STYLE;


	AscFormat.XMLWRITER_RECORD_TYPE_SPPR =  XMLWRITER_RECORD_TYPE_SPPR;
	AscFormat.XMLWRITER_RECORD_TYPE_CLRMAPOVR =  XMLWRITER_RECORD_TYPE_CLRMAPOVR;
	AscFormat.XMLWRITER_RECORD_TYPE_TEXT_OUTLINE =  XMLWRITER_RECORD_TYPE_TEXT_OUTLINE;
	AscFormat.XMLWRITER_RECORD_TYPE_TEXT_FILL =  XMLWRITER_RECORD_TYPE_TEXT_FILL;

	function CT_GraphicalObject(graphicFrame) {
		this.Namespace = null;
		this.GraphicData = null;

		this.GraphicFrame = graphicFrame;
		return this;
	}
	CT_GraphicalObject.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "graphicData" : {
					this.GraphicData = new CT_GraphicalObjectData(this.GraphicFrame);
					this.GraphicData.fromXml(reader);
					break;
				}
			}
		}
	};
	CT_GraphicalObject.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		if (this.Namespace) {
			writer.WriteXmlString(this.Namespace);
		}
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.GraphicData, "a:graphicData");
		writer.WriteXmlNodeEnd(name);
	};
	function CT_GraphicalObjectData(graphicFrame) {
		this.Uri = null;
		this.graphicObject = null;

		this.graphicFrame = graphicFrame;
	}
	CT_GraphicalObjectData.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "uri": {
					this.Uri = reader.GetValueDecodeXml();
					break;
				}
			}
		}
	};
	CT_GraphicalObjectData.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			this.graphicObject = AscFormat.CGraphicObjectBase.prototype.fromXmlElem(reader, name, this.graphicFrame);
		}
	};
	CT_GraphicalObjectData.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		//"http://schemas.openxmlformats.org/drawingml/2006/picture"
		writer.WriteXmlNullableAttributeStringEncode("uri", this.Uri);
		writer.WriteXmlAttributesEnd();
		AscFormat.CGraphicObjectBase.prototype.toXmlElem(writer, this.graphicObject, "pic");
		writer.WriteXmlNodeEnd(name);
	};

	window['AscFormat'].CGraphicObjectBase.prototype.fromXmlElem = function(reader, name, graphicFrame) {
		let res = null;
		if( "cxnSp" === name) {
			res = new AscFormat.CConnectionShape();
			res.setBDeleted(false);
			res.fromXml(reader);
		}
		else if("grpSp" === name || "wgp" === name) {
			res = new AscFormat.CGroupShape();
			res.setBDeleted(false);
			res.fromXml(reader);
		}
		else if("lockedCanvas" === name) {
			res = new AscFormat.CLockedCanvas();
			res.setBDeleted(false);
			res.fromXml(reader);
		}
		else if("sp" === name || "wsp" === name) {
			res = new AscFormat.CShape();
			res.setBDeleted(false);
			res.setWordShape("wsp" === name);
			res.fromXml(reader);
		}
		else if ("pic" === name) {
			res = new AscFormat.CImageShape();
			res.setBDeleted(false);
			res.fromXml(reader);
		} else if ("graphicFrame" === name) {
			res = new AscFormat.CGraphicFrame();
			res.fromXml(reader);
			res = res.graphicObject;
		} else if ("AlternateContent" === name) {
			let bRetNull = false;
			let elem = new CT_XmlNode(function(reader, name) {
				let oThis = this;
				if(!res) {
					if ("Choice" === name) {
						let elem = new CT_XmlNode(function(reader, name) {
							if(!res) {
								res = window['AscFormat'].CGraphicObjectBase.prototype.fromXmlElem.call(oThis, reader, name, graphicFrame);
							}
							return true;
						});
						elem.fromXml(reader);
						let oAttr = elem.attributes;
						if(oAttr["Requires"] === "cx" && oAttr["cx"] === "http://schemas.microsoft.com/office/drawing/2014/chartex") {
							bRetNull = true;
						}
						return elem;
					}
					else if("Fallback" === name) {
						let elem = new CT_XmlNode(function(reader, name) {
							if(!res) {
								res = window['AscFormat'].CGraphicObjectBase.prototype.fromXmlElem.call(oThis, reader, name, graphicFrame);
							}
							return true;
						});
						elem.fromXml(reader);
						return elem;
					}
				}
				return true;
			});
			elem.fromXml(reader);
			if(bRetNull) {
				if(res) {
					res = null;
				}
			}
		}  else if ("slicer" === name) {
			res = new AscFormat.CSlicer();
			res.fromXml(reader);
			let _xfrm = null;
			if(graphicFrame) {
				_xfrm = graphicFrame.spPr && graphicFrame.spPr.xfrm;
			}
			res.checkEmptySpPrAndXfrm(_xfrm);
		} else if ("chart" === name) {
			if (typeof AscFormat.CChartSpace !== "undefined") {
				let elem = new CT_XmlNode();
				elem.fromXml(reader);
				let rId = elem.attributes["id"];
				let rel = reader.rels.getRelationship(rId);
				if ("Internal" === rel.targetMode) {
					let chartPart = reader.rels.pkg.getPartByUri(rel.targetFullName);
					if (chartPart) {
						let chartContent = chartPart.getDocumentContent();
						let chartReader = new AscCommon.StaxParser(chartContent, chartPart, reader.context);
						res = new AscFormat.CChartSpace();
						res.fromXml(chartReader);

						res.setBDeleted(false);
						if (res.hasCharts()) {
							let chartStylePart = chartPart.getPartByRelationshipType(AscCommon.openXml.Types.chartStyle.relationType);
							if (chartStylePart) {
								let chartStyleContent = chartStylePart.getDocumentContent();
								if (chartStyleContent) {
									let chartStyle = new AscFormat.CChartStyle();
									let readerStyle = new AscCommon.StaxParser(chartStyleContent, chartStylePart, reader.context);
									chartStyle.fromXml(readerStyle);
									res.setChartStyle(chartStyle);
								}
							}
							let chartColorStylePart = chartPart.getPartByRelationshipType(AscCommon.openXml.Types.chartColorStyle.relationType);
							if (chartColorStylePart) {
								let chartColorStyleContent = chartColorStylePart.getDocumentContent();
								if (chartColorStyleContent) {
									let chartStyle = new AscFormat.CChartColors();
									let readerStyle = new AscCommon.StaxParser(chartColorStyleContent, chartColorStylePart, reader.context);
									chartStyle.fromXml(readerStyle);
									res.setChartColors(chartStyle);
								}
							}
							let themeOverridePart = chartPart.getPartByRelationshipType(AscCommon.openXml.Types.themeOverride.relationType);
							if (themeOverridePart) {
								let themeOverrideContent = themeOverridePart.getDocumentContent();
								if (themeOverrideContent) {
									let themeOverride = new AscFormat.CTheme();
									themeOverride.setIsThemeOverride(true);
									let readerThemeOverride = new AscCommon.StaxParser(themeOverrideContent, themeOverridePart, reader.context);
									themeOverride.themeElements.fromXml(readerThemeOverride, true);
									res.setThemeOverride(themeOverride);
								}
							}
							let _xfrm = null;
							if(graphicFrame) {
								_xfrm = graphicFrame.spPr && graphicFrame.spPr.xfrm;
							}
							res.checkEmptySpPrAndXfrm(_xfrm);
						} else {
							res = null;
						}
					}
				}
			}
		} else if("tbl" === name)  {
			if(graphicFrame) {
				let _xfrm = graphicFrame.spPr && graphicFrame.spPr.xfrm;
				let extX = 50;
				let extY = 50;
				if(_xfrm) {
					extX = _xfrm.extX;
					extY = _xfrm.extY;
				}
				let _table = CTable.prototype.static_readFromDrawingML(reader, graphicFrame, extX, extY);
				graphicFrame.setGraphicObject(_table);
				graphicFrame.checkEmptySpPrAndXfrm(_xfrm);

			}

		}
		else if("relIds" === name) {
			res = new AscFormat.SmartArt();
			res.fromXml(reader);
			let _xfrm = null;
			if(graphicFrame) {
				_xfrm = graphicFrame.spPr && graphicFrame.spPr.xfrm;
			}
			res.checkEmptySpPrAndXfrm(_xfrm);
		}
		else if("oleObj" === name) {
			if(reader.GetName() === "p:oleObj") {

				let oOleObject = new AscFormat.COleObject();
				oOleObject.m_sData = null

				let oOleNode = new CT_XmlNode(function(reader, name) {
					if(name === "pic") {
						oOleObject.fromXml(reader);
						res = oOleObject;
						return true;
					}
					return false;
				});
				oOleNode.fromXml(reader);
				let sAppName = oOleNode.attributes["progId"];
				if(sAppName) {
					oOleObject.setApplicationId(sAppName);
				}
				let isN = AscFormat.isRealNumber;
				let nImgH = reader.GetInt(oOleNode.attributes["imgH"]);
				let nImgW = reader.GetInt(oOleNode.attributes["imgW"]);
				if(isN(nImgW) && isN(nImgH)) {
					oOleObject.setPixSizes(AscFormat.Emu_To_Px(nImgW), AscFormat.Emu_To_Px(nImgH));
				}
				oOleObject.fillDataLink(oOleNode.attributes["id"], reader);
				let oVmlPart = reader.rels.getPartByRelationshipType(AscCommon.openXml.Types.vmlDrawing.relationType);
				if(oVmlPart) {
					let oVmlPartContent = oVmlPart.getDocumentContent();
					if(oVmlPartContent) {
						let oVmlReader = new AscCommon.StaxParser(oVmlPartContent, oVmlPart, reader.context);
						let oVmlDrawing = new AscFormat.CVMLDrawing();
						oVmlDrawing.fromXml(oVmlReader);
						let oVmlShape = oVmlDrawing.getShapeBySpId(oOleNode.attributes["spid"]);
						if(oVmlShape) {
							let oImageData = oVmlShape.getImageData();
							if(oImageData) {
								let oOldReader = reader.context.reader;
								reader.context.reader = oVmlReader;
								let oFill = oImageData.getOOXMLFill(reader.context);
								if(oFill && oFill.isBlipFill()) {
									oOleObject.setBlipFill(oFill.fill);
									res = oOleObject;
								}
								reader.context.reader = oOldReader;
							}
						}
					}
				}
				if(!oOleObject.spPr) {
					oOleObject.setSpPr(new AscFormat.CSpPr());
				}
				oOleObject.spPr.setGeometry(AscFormat.CreateGeometry("rect"));
			}
			else {
				let oOLEObj = new AscFormat.COLEObject();
				oOLEObj.fromXml(reader);
				if(oOLEObj.m_oPic) {
					res = new AscFormat.COleObject();
					oOLEObj.fillEditorOleObject(res, oOLEObj.m_oPic, reader);
				}
			}
		}
		return res;
	};
	window['AscFormat'].CGraphicObjectBase.prototype.toXmlElem = function(writer, graphicObject, ns) {
		let elemForWriting = graphicObject;
		let graphicObjectName;
		switch (graphicObject.getObjectType()) {
			case AscDFH.historyitem_type_Shape:
				elemForWriting = graphicObject;
				break;
			case AscDFH.historyitem_type_Cnx:
				elemForWriting = graphicObject;
				break;
			case AscDFH.historyitem_type_OleObject:
			case AscDFH.historyitem_type_ImageShape:
				graphicObjectName = ns + ":pic";
				elemForWriting = graphicObject;
				break;
			case AscDFH.historyitem_type_GroupShape:
			case AscDFH.historyitem_type_SmartArt:
				elemForWriting = graphicObject;
				break;
			case AscDFH.historyitem_type_GraphicFrame:
				graphicObjectName = ns + ":graphicFrame";
				break;
			case AscDFH.historyitem_type_ChartSpace:
			{
				let partType = AscCommon.c_oEditorId.Word === writer.context.editorId ? AscCommon.openXml.Types.chartWord: AscCommon.openXml.Types.chart;
				graphicObjectName = "c:chart";
				let chartPart = writer.context.part.addPart(partType);
				chartPart.part.setDataXml(graphicObject, writer);
				let elem = new CT_XmlNode();
				elem.attributes["xmlns:c"] = "http://schemas.openxmlformats.org/drawingml/2006/chart";
				elem.attributes["xmlns:r"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
				elem.attributes["r:id"] = chartPart.rId;
				elemForWriting = elem;
				if (graphicObject.chartStyle) {
					let chartStylePart = chartPart.part.addPart(AscCommon.openXml.Types.chartStyle);
					chartStylePart.part.setDataXml(graphicObject.chartStyle, writer);
				}
				if (graphicObject.chartColors) {
					let chartColorPart = chartPart.part.addPart(AscCommon.openXml.Types.chartColorStyle);
					chartColorPart.part.setDataXml(graphicObject.chartColors, writer);
				}
				if (graphicObject.themeOverride) {
					let themeOverridePart = chartPart.part.addPart(AscCommon.openXml.Types.themeOverride);
					let memory = new AscCommon.CMemory();
					memory.context = writer.context;
					writer.context.clearCurrentPartDataMaps();
					memory.WriteXmlString(AscCommonWord.g_sXmlHeader);
					graphicObject.themeOverride.themeElements.toXml(memory, "a:themeOverride");
					let commentAuthorsData = memory.GetDataUint8();
					themeOverridePart.part.setData(commentAuthorsData);
					memory.Seek(0);
				}
				break;
			}
			case AscDFH.historyitem_type_SlicerView:
				elemForWriting = graphicObject;
				break;
			case AscDFH.historyitem_type_Table:
				graphicObjectName = "a:tbl";
				elemForWriting = graphicObject;
				break;
		}
		if (elemForWriting) {
			if(elemForWriting.toDrawingML) {
				elemForWriting.toDrawingML(writer, graphicObjectName)
			}
			else {
				elemForWriting.toXml(writer, graphicObjectName);
			}
		}
	};



	// window['AscFormat'].CBlipFill.prototype.toXml = function(writer, name)
	// {
	// 	let context = writer.context;
	// 	let imagePart = context.imageMap[this.RasterImageId];
	// 	if (!imagePart) {
	// 		let ext = AscCommon.GetFileExtension(this.RasterImageId);
	// 		let type = context.editorId === AscCommon.c_oEditorId.Word? AscCommon.openXml.Types.imageWord : AscCommon.openXml.Types.image;
	// 		type = Object.assign({}, type);
	// 		type.filename += ext;
	// 		type.contentType = openXml.GetMimeType(ext);
	// 		imagePart = context.part.addPart(type);
	// 		context.imageMap[this.RasterImageId] = imagePart;
	// 	}

	// 	writer.WriteXmlNodeStart(name);
	// 	writer.WriteXmlAttributesEnd();

	// 	writer.WriteXmlNodeStart("a:blip");
	// 	writer.WriteXmlString(' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
	// 	writer.WriteXmlAttributeString("r:embed", imagePart.rId);
	// 	writer.WriteXmlAttributesEnd();
	// 	writer.WriteXmlString('<a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst>');
	// 	writer.WriteXmlNodeEnd("a:blip");

	// 	writer.WriteXmlNodeStart("a:stretch");
	// 	writer.WriteXmlAttributesEnd();
	// 	writer.WriteXmlNodeStart("a:fillRect");
	// 	writer.WriteXmlAttributesEnd(true);
	// 	writer.WriteXmlNodeEnd("a:stretch");

	// 	writer.WriteXmlNodeEnd(name);
	// };

	window['AscFormat'].CT_GraphicalObject = CT_GraphicalObject;
	window['AscFormat'].CT_GraphicalObjectData = CT_GraphicalObjectData;

	function CT_ComplexType() {
		this.val = null;
		return this;
	}
	CT_ComplexType.prototype.constructor = CT_ComplexType;
	CT_ComplexType.prototype.readAttrVal = function(reader) {
	};
	CT_ComplexType.prototype.writeAttrVal = function(writer, val) {
	};
	CT_ComplexType.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "val": {
					this.val = this.readAttrVal(reader);
					break;
				}
			}
		}
	};
	CT_ComplexType.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_ComplexType.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		this.writeAttrVal(writer, this.val);
		writer.WriteXmlAttributesEnd(true);
	};
	CT_ComplexType.prototype.fromVal = function(val) {
		let res = null;
		if (null !== val && undefined !== val) {
			res = new this.constructor();
			res.val = val;
		}
		return res;
	};
	CT_ComplexType.prototype.toVal = function(reader, def) {
		let elem = new this.constructor();
		elem.fromXml(reader);
		return elem.getVal(def);
	};
	CT_ComplexType.prototype.getVal = function(def) {
		return null !== this.val ? this.val : def;
	};

	function CT_String() {
		CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_String, CT_ComplexType, 0);
	CT_String.prototype.readAttrVal = function(reader) {
		return reader.GetValueDecodeXml();
	};
	CT_String.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeStringEncode("val", val);
	};

	function CT_Bool() {
		CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_Bool, CT_ComplexType, 0);
	CT_Bool.prototype.readAttrVal = function(reader) {
		return reader.GetValueBool();
	};
	CT_Bool.prototype.writeAttrVal = function(writer, val) {
		if (!val) {
			writer.WriteXmlNullableAttributeBool("val", val);
		}
	};
	CT_Bool.prototype.getVal = function() {
		return null !== this.val ? this.val : true;
	};

	function CT_Int() {
		CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_Int, CT_ComplexType, 0);
	CT_Int.prototype.readAttrVal = function(reader) {
		return reader.GetValueInt(this.val);
	};
	CT_Int.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeInt("val", this.val);
	};
	CT_Int.prototype.fromVal = function(val, koef) {
		if (koef && null !== val && undefined !== val ) {
			return CT_ComplexType.prototype.fromVal.call(this, val * koef);
		} else {
			return CT_ComplexType.prototype.fromVal.call(this, val);
		}
	};

	function CT_UInt() {
		CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_UInt, CT_ComplexType, 0);
	CT_UInt.prototype.readAttrVal = function(reader) {
		return reader.GetValueUInt(this.val);
	};
	CT_UInt.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeUInt("val", this.val);
	};
	CT_UInt.prototype.fromVal = function(val, koef) {
		if (koef && null !== val && undefined !== val ) {
			return CT_ComplexType.prototype.fromVal.call(this, val * koef);
		} else {
			return CT_ComplexType.prototype.fromVal.call(this, val);
		}
	};

	function CT_Double() {
		CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_Double, CT_ComplexType, 0);
	CT_Double.prototype.readAttrVal = function(reader) {
		return reader.GetValueDouble(this.val);
	};
	CT_Double.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeDouble("val", this.val);
	};

	function CT_StringW() {
		CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_StringW, CT_String, 0);
	CT_StringW.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeStringEncode("w:val", val);
	};

	function CT_BoolW() {
		CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_BoolW, CT_Bool, 0);
	CT_BoolW.prototype.writeAttrVal = function(writer, val) {
		if (!val) {
			writer.WriteXmlNullableAttributeBool("w:val", val);
		}
	};

	function CT_IntW() {
		CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_IntW, CT_Int, 0);
	CT_IntW.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeInt("w:val", this.val);
	};

	function CT_UIntW() {
		CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_UIntW, CT_UInt, 0);
	CT_UIntW.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeUInt("w:val", this.val);
	};

	function CT_DoubleW() {
		CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_DoubleW, CT_Double, 0);
	CT_DoubleW.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeDouble("w:val", this.val);
	};

	AscCommonWord.Paragraph.prototype.fromDrawingML = function(reader) {
		let depth = reader.GetDepth();
		let EndPos = 0;
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			switch(name) {
				case "br": {
					let oRun = new AscCommonWord.ParaRun(this, false);
					oRun.AddToContent( 0, new AscWord.CRunBreak(AscWord.break_Line));
					this.AddToContent(EndPos++, oRun);
					oRun.fromDrawingML(reader);//Read run properties
					break;
				}
				case "endParaRPr": {
					let oTextPr = new AscCommonWord.CTextPr();
					oTextPr.fromDrawingML(reader);
					this.TextPr.Apply_TextPr(oTextPr);
					let oEndRun = this.Content[this.Content.length - 1]; //TODO: check whether it needs to set properties for end run
					let oParaTextPrEnd = new CTextPr();
					oParaTextPrEnd.Set_FromObject(oTextPr);
					oEndRun.Set_Pr(oParaTextPrEnd);
					break;
				}
				case "fld": {
					let oFld = new AscCommonWord.CPresentationField(this);
					oFld.fromDrawingML(reader);
					this.AddToContent(EndPos++, new ParaRun(this, false));
					this.AddToContent(EndPos++, oFld);
					this.AddToContent(EndPos++, new ParaRun(this, false));
					break;
				}
				case "pPr": {
					let oParaPr = new AscCommonWord.CParaPr();
					oParaPr.fromDrawingML(reader);
					this.Set_Pr(oParaPr);
					break;
				}
				case "r": {
					let oRun = new AscCommonWord.ParaRun(this, false);
					oRun.fromDrawingML(reader);
					let oRunPr = oRun.Pr;
					let oHyperLink = oRunPr && (oRunPr.hlinkClick || oRunPr.hlinkMouseOver);
					if(oHyperLink) {
						let oParaHLink = new ParaHyperlink();
						oParaHLink.SetValue(oHyperLink.id);
						if (oHyperLink.tooltip) {
							oParaHLink.SetToolTip(oHyperLink.tooltip);
						}
						oRunPr.Underline = true;
						oParaHLink.Add_ToContent(0, oRun, false);
						this.AddToContent(EndPos++, oParaHLink);
					}
					else {
						this.AddToContent(EndPos++, oRun);
					}
					break;
				}
				case "m": {
					let oThis = this;
					let oMathNode = new CT_XmlNode(function (reader, name) {
						switch (name) {
							case "oMath":
								CParagraphContentWithParagraphLikeContent.prototype.fromXmlElem.call(oThis, reader, name);
								break;
							case "oMathPara":
								CParagraphContentWithParagraphLikeContent.prototype.fromXmlElem.call(oThis, reader, name);
								break;
						}
						return true;
					});
					oMathNode.fromXml(reader);
					break;
				}
			}
		}
		this.Correct_Content();
	};
	AscCommonWord.ParaHyperlink.prototype.toDrawingML = function(writer) {
		for(let nIdx = 0; nIdx < this.Content.length; ++nIdx) {
			let oElement = this.Content[nIdx];
			if(oElement.toDrawingML) {
				oElement.toDrawingML(writer, nIdx, this.Paragraph);
			}
		}
	};
	ParaMath.prototype.toDrawingML = function(writer, index, paragraph) {
		writer.WriteXmlNodeStart("mc:AlternateContent");
		writer.WriteXmlAttributeString("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
		writer.WriteXmlAttributeString("xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
		writer.WriteXmlAttributesEnd();

		writer.WriteXmlNodeStart("mc:Choice");
		writer.WriteXmlAttributeString("xmlns:a14", "http://schemas.microsoft.com/office/drawing/2010/main");
		writer.WriteXmlAttributeString("Requires", "a14");
		writer.WriteXmlAttributesEnd();

		writer.WriteXmlNodeStart("a14:m");
		writer.WriteXmlAttributesEnd();
		if (paragraph.CheckMathPara(index)) {
			let mathPara = new AscCommon.CT_OMathPara();
			mathPara.initMathParaPr(this);
			mathPara.toXml(writer, "m:oMathPara", this);
		} else {
			this.toXml(writer, "m:oMath");
		}
		writer.WriteXmlNodeEnd("a14:m");

		writer.WriteXmlNodeEnd("mc:Choice");


		writer.WriteXmlNodeStart("mc:Fallback");
		writer.WriteXmlAttributesEnd(true);

		writer.WriteXmlNodeEnd("mc:AlternateContent");
	};
	AscCommonWord.Paragraph.prototype.toDrawingML = function(writer) {
		writer.WriteXmlNodeStart("a:p");
		writer.WriteXmlAttributesEnd();

		if (this.Pr) {
			this.Pr.toDrawingML(writer, "a:pPr");
		}

		let nCount = this.Content.length;
		for (let i = 0; i < nCount; ++i)
			if(this.Content[i].toDrawingML) {
				this.Content[i].toDrawingML(writer, i, this);
			}

		if(this.TextPr && this.TextPr.Value) {
			this.TextPr.Value.toDrawingML(writer, "a:endParaRPr")
		}

		writer.WriteXmlNodeEnd("a:p");
	};
	AscCommonWord.ParaRun.prototype.fromDrawingML = function (reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			switch(name) {
				case "rPr": {
					let oTextPr = new AscCommonWord.CTextPr();
					oTextPr.fromDrawingML(reader);
					this.SetPr(oTextPr);
					break;
				}
				case "t": {
					this.AddText(reader.GetTextDecodeXml(), -1);
					break;
				}
			}
		}
	};
	AscCommonWord.ParaRun.prototype.toDrawingML = function (writer) {
		if(this.IsParaEndRun()) {
			return;
		}
		let nStart = -1;
		let nIdx = 0;
		for(; nIdx < this.Content.length; ++nIdx) {
			let oItem = this.Content[nIdx];
			if(oItem.Type === para_NewLine) {
				this.toDrawingMLContent(writer, nStart, nIdx - 1);
				nStart = -1;
				writer.WriteXmlString("<a:br/>");
			}
			else {
				if(nStart === -1) {
					nStart = nIdx;
				}
			}
		}
		this.toDrawingMLContent(writer, nStart, nIdx - 1);
	};
	AscCommonWord.ParaRun.prototype.toDrawingMLContent = function (writer, nStart, nEnd) {
		if(nStart < 0 || nEnd < nStart || nEnd >= this.Content.length || nStart < 0) {
			return;
		}
		if(this.Content.length === 0) {
			return;
		}
		writer.WriteXmlNodeStart("a:r");
		writer.WriteXmlAttributesEnd();

		if (this.Pr) {
			this.Pr.toDrawingML(writer, "a:rPr", this);
		}
		this.toDrawingMLText(writer, nStart, nEnd);
		writer.WriteXmlNodeEnd("a:r");
	};
	AscCommonWord.ParaRun.prototype.toDrawingMLText = function(writer, nStart, nEnd) {
		
		writer.WriteXmlString("<a:t>");
		for(let nIdx = nStart; nIdx <= nEnd; ++nIdx) {
			let oItem = this.Content[nIdx];
			switch(oItem.Type) {
				case para_Text:
				case para_Space:
				{
					writer.WriteXmlCharCode(oItem.Value);
					break;
				}
				case para_Tab :
				{
					writer.WriteXmlCharCode(0x09);
					break;
				}
				case para_NewLine :
				{
					//impossible situation. we split run by para_NewLine
					break;
				}
			}
		}
		writer.WriteXmlString("</a:t>");
	};
	AscCommonWord.CTextPr.prototype.fromDrawingML = function (reader) {
		let sName;
		while (reader.MoveToNextAttribute()) {
			sName = reader.GetNameNoNS();
			switch(sName) {
				case "altLang": {
					break;
				}
				case "b": {
					this.Bold = reader.GetValueBool();
					break;
				}
				case "baseline": {
					let nBaseline = AscFormat.getPercentageValue(reader.GetValue());
					if (nBaseline < 0)
						this.VertAlign = AscCommon.vertalign_SubScript;
					else if (nBaseline > 0)
						this.VertAlign = AscCommon.vertalign_SuperScript;
					break;
				}
				case "bmk": {
					break;
				}
				case "cap": {
					let sCap = reader.GetValue();
					if(sCap === "all") {
						this.Caps = true;
						this.SmallCaps = false;
					}
					else if(sCap === "small") {
						this.Caps = false;
						this.SmallCaps = true;
					}
					else {
						this.Caps = false;
						this.SmallCaps = false;
					}
 					break;
				}
				case "dirty": {
					break;
				}
				case "err": {
					break;
				}
				case "i": {
					this.Italic = reader.GetValueBool();
					break;
				}
				case "kern": {
					break;
				}
				case "kumimoji": {
					break;
				}
				case "lang": {
					let sLang = reader.GetValue();
					let nLcid = Asc.g_oLcidNameToIdMap[sLang];
					if(nLcid)
						this.Lang.Val = nLcid;
					break;
				}
				case "noProof": {
					break;
				}
				case "normalizeH": {
					break;
				}
				case "smtClean": {
					break;
				}
				case "smtId": {
					break;
				}
				case "spc": {
					this.Spacing = reader.GetValueInt() * 25.4 / 7200;
					break;
				}
				case "strike": {
					let sStrike = reader.GetValue();
					if(sStrike === "dblStrike") {
						this.Strikeout = false;
						this.DStrikeout = true;
					}
					else if(sStrike === "sngStrike") {
						this.Strikeout = true;
						this.DStrikeout = false;
					}
					else if(sStrike === "noStrike") {
						this.Strikeout = false;
						this.DStrikeout = false;
					}
					break;
				}
				case "sz": {
					let nSz = reader.GetValueInt() / 100;
					nSz = ((nSz * 2) + 0.5) >> 0;
					nSz /= 2;
					this.FontSize = nSz;
					this.FontSizeCS = nSz;
					break;
				}
				case "u": {
					this.Underline = reader.GetValue() !== "none";
					break;
				}
			}
		}
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if(AscFormat.CUniFill.prototype.isFillName(name)) {
				this.Unifill = new AscFormat.CUniFill();
				this.Unifill.fromXml(reader, name);
			}
			else if(name === "ln") {
				this.TextOutline = new AscFormat.CLn();
				this.TextOutline.fromXml(reader);
			}
			else if(name === "hlinkClick" || name === "hlinkMouseOver") {
				let oHLink = new AscFormat.CT_Hyperlink();
				oHLink.fromXml(reader);
				if(name === "hlinkClick") {
					this.hlinkClick = oHLink;
				}
				else {
					this.hlinkMouseOver = oHLink;
				}
			}
			else if(name === "highlight") {
				this.HighlightColor = readHighlightColor(reader);
			}
			else {
				readRFont(this, reader, name);
			}
		}
	};
	AscCommonWord.CTextPr.prototype.toDrawingML = function (writer, sName, oRun) {
		writer.WriteXmlNodeStart(sName);
		//writer.WriteXmlAttributeString("kumimoji", kumimoji);
		
		writer.WriteXmlNullableAttributeString("lang", Asc.g_oLcidIdToNameMap[this.Lang.Val]);
		//writer.WriteXmlAttributeString("altLang", altLang);
		if(this.FontSize !== null && this.FontSize !== undefined) {
			writer.WriteXmlAttributeUInt("sz", this.FontSize * 100 >> 0);
		}
		writer.WriteXmlNullableAttributeBool("b", this.Bold);
		writer.WriteXmlNullableAttributeBool("i", this.Italic);
		if(this.Underline !== undefined && this.Underline !== null) {
			if(!this.Underline) {
				writer.WriteXmlAttributeString("u", "none");
			}
			else {
				writer.WriteXmlAttributeString("u", "sng");
			}
		}

		if(this.Strikeout === false && this.DStrikeout === true) {
			writer.WriteXmlAttributeString("strike", "dblStrike");
		}
		else if(this.Strikeout === true && this.DStrikeout === false) {
			writer.WriteXmlAttributeString("strike", "sngStrike");
		}
		else if(this.Strikeout === false && this.DStrikeout === false) {
			writer.WriteXmlAttributeString("strike", "noStrike");
		}
		//writer.WriteXmlAttributeString("kern", kern);


		if(this.Caps === true && this.SmallCaps === false) {
			writer.WriteXmlAttributeString("cap", "all");
		}
		else if(this.Caps === false && this.SmallCaps === true) {
			writer.WriteXmlAttributeString("cap", "small");
		}
		else if(this.Caps === false && this.SmallCaps === false) {
			writer.WriteXmlAttributeString("cap", "none");
		}


		if(this.Spacing !== undefined && this.Spacing !== null) {
			writer.WriteXmlAttributeInt("spc", this.Spacing * 7200 / 25.4 + 0.5 >> 0);
		}
		//writer.WriteXmlAttributeString("normalizeH", normalizeH);


        if (AscCommon.vertalign_SubScript === this.VertAlign) {
			writer.WriteXmlAttributeInt("baseline", -25000);
		}
        else if (AscCommon.vertalign_SuperScript === this.VertAlign) {
			writer.WriteXmlAttributeInt("baseline", 30000);
		}

		//writer.WriteXmlAttributeString("noProof", noProof);
		//writer.WriteXmlAttributeString("dirty", dirty);
		//writer.WriteXmlAttributeString("err", err);
		//writer.WriteXmlAttributeString("smtClean", smtClean);
		//writer.WriteXmlAttributeString("smtId", smtId);
		//writer.WriteXmlAttributeString("bmk", bmk);

		let oParaHyperlink = null;
		if(oRun) {
			if(oRun.Parent instanceof ParaHyperlink) {
				oParaHyperlink = oRun.Parent;
			}
		}
		if(this.TextOutline || this.Unifill || this.HighlightColor ||
			(this.RFonts.Ascii && this.RFonts.Ascii.Name) ||
			(this.RFonts.EastAsia && this.RFonts.EastAsia.Name) ||
			(this.RFonts.CS && this.RFonts.CS.Name) ||
			oParaHyperlink) {

			writer.WriteXmlAttributesEnd();

			if(this.TextOutline) {
				this.TextOutline.toXml(writer, "a:ln");
			}
			if(this.Unifill) {
				this.Unifill.toXml(writer);
			}
			//EffectList.toXmlWriter(writer);
			if(this.HighlightColor) {
				writer.WriteXmlNodeStart("a:highlight");
				writer.WriteXmlAttributesEnd();
				this.HighlightColor.toXml(writer);
				writer.WriteXmlNodeEnd("a:highlight");
			}
			//writer.Write(uFill);
			//writer.Write(uFillTx);
			if(this.RFonts.Ascii)
				writeTypeface(writer, "a:latin", this.RFonts.Ascii.Name);
			if(this.RFonts.EastAsia)
				writeTypeface(writer, "a:ea", this.RFonts.EastAsia.Name);
			if(this.RFonts.CS)
				writeTypeface(writer, "a:cs", this.RFonts.CS.Name);
			//writer.Write(sym);


			if(oParaHyperlink) {
				let oHyperlink = new AscFormat.CT_Hyperlink();
				oHyperlink.id = oParaHyperlink.Value;
				if (oParaHyperlink.tooltip) {
					oHyperlink.tooltip = oParaHyperlink.tooltip;
				}
				oHyperlink.toXml(writer, "a:hlinkClick");
			}
			// if(this.hlinkClick) {
			// 	this.hlinkClick.toXml(writer, "a:hlinkClick");
			// }
			//
			// if(this.hlinkMouseOver) {
			// 	this.hlinkMouseOver.toXml(writer, "a:hlinkMouseOver");
			// }
			//writer.Write(rtl);

			writer.WriteXmlNodeEnd(sName);
		}
		else {
			writer.WriteXmlAttributesEnd(true);
		}
	};
	AscCommonWord.CParaPr.prototype.fromDrawingML = function (reader) {
		let sName;
		while (reader.MoveToNextAttribute()) {
			sName = reader.GetNameNoNS();
			switch(sName) {
				case "algn": {
					let sVal = reader.GetValue();
					switch (sVal) {
						case "ctr": {
							this.Jc = AscCommon.align_Center;
							break;
						}
						case "dist": {
							this.Jc = AscCommon.align_Justify;
							break;
						}
						case "just": {
							this.Jc = AscCommon.align_Justify;
							break;
						}
						case "justLow": {
							this.Jc = AscCommon.align_Justify;
							break;
						}
						case "l": {
							this.Jc = AscCommon.align_Left;
							break;
						}
						case "r": {
							this.Jc = AscCommon.align_Right;
							break;
						}
						case "thaiDist": {
							this.Jc = AscCommon.align_Justify;
							break;
						}
					}
					break;
				}
				case "defTabSz": {
					this.DefaultTab = reader.GetValueInt()/36000;
					break;
				}
				case "eaLnBrk": {
					break;
				}
				case "fontAlgn": {
					break;
				}
				case "hangingPunct": {
					break;
				}
				case "indent": {
					this.Ind.FirstLine = reader.GetValueInt()/36000;
					break;
				}
				case "latinLnBrk": {
					break;
				}
				case "lvl": {
					this.Lvl = reader.GetValueInt();
					break;
				}
				case "marL": {
					this.Ind.Left = reader.GetValueInt()/36000;
					break;
				}
				case "marR": {
					this.Ind.Right = reader.GetValueInt()/36000;
					break;
				}
				case "rtl": {
					break;
				}
			}
		}
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if(name.indexOf("bu") === 0) {
				if(!this.Bullet) {
					this.Bullet = new AscFormat.CBullet();
				}
				this.Bullet.readChildXml(name, reader);
			}
			else if(name === "defRPr") {
				let oRPr = new AscCommonWord.CTextPr();
				oRPr.fromDrawingML(reader);
				this.DefaultRunPr = oRPr;
			}
			else if(name === "extLst") {

			}
			else if(name === "lnSpc") {
				this.Spacing.lnSpcFromDrawingML(reader);
			}
			else if(name === "spcAft") {
				this.Spacing.spcAftFromDrawingML(reader);
			}
			else if(name === "spcBef") {
				this.Spacing.spcBefFromDrawingML(reader);
			}
			else if(name === "tabLst") {
				this.Tabs = new CParaTabs();
				this.Tabs.fromDrawingML(reader);
			}
		}
	};
	AscCommonWord.CParaPr.prototype.toDrawingML = function (writer,sName) {
		writer.WriteXmlNodeStart(sName);
		if(this.Ind) {
			if(this.Ind.Left !== null && this.Ind.Left !== undefined) {

				writer.WriteXmlAttributeInt("marL", this.Ind.Left * 36000 + 0.5 >> 0);
			}
			if(this.Ind.Right !== null && this.Ind.Right !== undefined) {

				writer.WriteXmlAttributeInt("marR", this.Ind.Right * 36000 + 0.5 >> 0);
			}
		}
		writer.WriteXmlNullableAttributeInt("lvl", this.Lvl);
		
		if(this.Ind) {
			if(this.Ind.FirstLine !== null && this.Ind.FirstLine !== undefined) {
				writer.WriteXmlAttributeInt("indent", (this.Ind.FirstLine * 36000 + 0.5) >> 0);
			}
		}
		if(this.Jc !== undefined && this.Jc !== null) {
			switch(this.Jc) {
				case AscCommon.align_Center: {
					writer.WriteXmlAttributeString("algn", "ctr");
					break;
				}
				case AscCommon.align_Justify: {
					writer.WriteXmlAttributeString("algn", "just");
					break;
				}
				case AscCommon.align_Left: {
					writer.WriteXmlAttributeString("algn", "l");
					break;
				}
				case AscCommon.align_Right: {
					writer.WriteXmlAttributeString("algn", "r");
					break;
				}
			}
		}
		
		if(this.DefaultTab !== null && this.DefaultTab !== undefined) {
			writer.WriteXmlAttributeInt("defTabSz", this.DefaultTab * 36000 + 0.5 >> 0);
		}
		//writer.WriteXmlAttributeString("rtl", rtl);
		//writer.WriteXmlAttributeString("eaLnBrk", eaLnBrk);
		//writer.WriteXmlAttributeString("fontAlgn", fontAlgn);
		//writer.WriteXmlAttributeString("latinLnBrk", latinLnBrk);
		//writer.WriteXmlAttributeString("hangingPunct", hangingPunct);
		writer.WriteXmlAttributesEnd();

		this.Spacing.lnSpcToDrawingML(writer);
		this.Spacing.spcBefToDrawingML(writer);
		this.Spacing.spcAftToDrawingML(writer);
		
		if(this.Bullet) {
			this.Bullet.toXml(writer);
		}
		if(this.Tabs) {
			this.Tabs.toDrawingML(writer);
		}
		if(this.DefaultRunPr) {
			this.DefaultRunPr.toDrawingML(writer, "a:defRPr");
		}
		else {
			writer.WriteXmlString("<a:defRPr/>");
		}
		writer.WriteXmlNodeEnd(sName);
	};
	AscCommonWord.CPresentationField.prototype.fromDrawingML = function(reader) {
		let sName;
		while (reader.MoveToNextAttribute()) {
			sName = reader.GetNameNoNS();
			switch(sName) {
				case "id": {
					this.SetGuid(reader.GetValue());
					break;
				}
				case "type": {
					this.SetFieldType(reader.GetValue());
					break;
				}
			}
		}
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if(name === "pPr") {
				let oParaPr = new AscCommonWord.CParaPr();
				oParaPr.fromDrawingML(reader);
				this.SetPPr(oParaPr);
			}
			else if(name === "rPr") {
				let oTextPr = new AscCommonWord.CTextPr();
				oTextPr.fromDrawingML(reader);
				this.SetPr(oTextPr);
			}
			else if(name === "t") {
				this.AddText(reader.GetTextDecodeXml(), -1);
			}
		}
	};
	AscCommonWord.CPresentationField.prototype.toDrawingML = function(writer) {
		writer.WriteXmlNodeStart("a:fld");
		writer.WriteXmlNullableAttributeString("id", this.Guid);
		writer.WriteXmlNullableAttributeString("type", this.FieldType);
		writer.WriteXmlAttributesEnd();
		if(this.Pr) {
			this.Pr.toDrawingML(writer, "a:rPr");
		}
		if(this.PPr) {
			this.PPr.toDrawingML(writer, "a:pPr");
		}
		if(this.Content.length > 0) {
			this.toDrawingMLText(writer, 0, this.Content.length - 1);
		}
		else {
			writer.WriteXmlString("<a:t></a:t>");
		}
		writer.WriteXmlNodeEnd("a:fld");
	};

	AscCommonWord.CTable.prototype.static_readFromDrawingML = function(reader, graphicFrame, extX, extY) {
		let nRows = 0;
		let aColsGrid = [];
		let oDrawingDocument = reader.context.DrawingDocument;
		let oParent = graphicFrame || null;
		let oStartState = reader.getState();
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if(name === "tblGrid") {
				let oPr = new CT_XmlNode(function (reader, name) {
					if(name === "gridCol") {
						let oPr = new CT_XmlNode(function (reader, name) {
							return true;
						});
						oPr.fromXml(reader);
						if(oPr.attributes["w"]) {
							let nW = parseInt(oPr.attributes["w"]);
							if(AscFormat.isRealNumber(nW)) {
								aColsGrid.push(nW / 36000);
							}
						}
					}
					return true;
				});
				oPr.fromXml(reader);
			}
			else if(name === "tblPr") {
			}
			else if(name === "tr") {
				++nRows;
			}
		}
		if(aColsGrid.length === 0) {
			aColsGrid.push(extX);
		}

		let oTable = new CTable(oDrawingDocument, oParent, false, nRows, aColsGrid.length, aColsGrid, true);
		oTable.Set_TableLook(new AscCommon.CTableLook(false, false, false, false, false, false));
		oTable.Reset(0, 0, extX, 100000, 0, 0, 1);
		reader.setState(oStartState);
		depth = reader.GetDepth();
		let nCurRow = 0;
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if(name === "tr") {
				let oRow = oTable.Content[nCurRow++];
				if(oRow) {
					oRow.fromDrawingML(reader)
				}
			}
			else if(name === "tblPr") {
				let oPr = new CTablePr();
				oPr.fromDrawingML(reader, oTable);
				oTable.Set_Pr(oPr)
			}
		}
		oTable.SetTableLayout(tbllayout_Fixed);
		return oTable;
	};

	AscCommonWord.CTable.prototype.fromDrawingML = function(reader) {
		let oTable = this;

		this.SetTableLayout(tbllayout_Fixed);
		reader.context.TablesMap[this.Get_Id()] = this;
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if(name === "tblGrid") {
				let aGrid = [];
				let oPr = new CT_XmlNode(function (reader, name) {
					if(name === "gridCol") {
						let oPr = new CT_XmlNode(function (reader, name) {
							return true;
						});
						oPr.fromXml(reader);
						if(oPr.attributes["w"]) {
							let nW = parseInt(oPr.attributes["w"]);
							if(AscFormat.isRealNumber(nW)) {
								aGrid.push(nW / 36000);
							}
						}
					}
					return true;
				});
				oPr.fromXml(reader);
				this.SetTableGrid(aGrid);
			}
			else if(name === "tblPr") {
				let oPr = new CTablePr();
				oPr.fromDrawingML(reader, this);
				this.Set_Pr(oPr);
			}
			else if(name === "tr") {
				let nIdx = this.Content.length;
				let oRow = new CTableRow(this, 0);
				oRow.Index = nIdx;
				oRow.fromDrawingML(reader);
				this.Content[nIdx] = oRow;
				this.Content[nIdx].Recalc_CompiledPr();
				History.Add(new CChangesTableAddRow(this, nIdx, [this.Content[nIdx]]));
			}
		}

		this.SetTableLayout(tbllayout_Fixed);

	};
	AscCommonWord.CTable.prototype.toDrawingML = function (writer) {

		writer.WriteXmlNodeStart("a:tbl");
		writer.WriteXmlAttributesEnd();

		this.Pr.toDrawingML(writer, this);

		writer.WriteXmlString("<a:tblGrid>");
		for (let nGridCol = 0; nGridCol < this.TableGrid.length; ++nGridCol) {
			writer.WriteXmlNodeStart("a:gridCol");
			writer.WriteXmlAttributeUInt("w", this.TableGrid[nGridCol] * 36000 >> 0);
			writer.WriteXmlAttributesEnd(true);
		}
		writer.WriteXmlString("</a:tblGrid>");
		let oTableRowGrid = AscCommon.GenerateTableWriteGrid(this);
		for (let nRow = 0; nRow < this.Content.length; ++nRow) {
			this.Content[nRow].toDrawingML(writer, oTableRowGrid.Rows[nRow]);
		}
		writer.WriteXmlNodeEnd("a:tbl");
	};
	AscCommonWord.CTable.prototype.getObjectType = function() {
		return AscDFH.historyitem_type_Table;
	};
	CTablePr.prototype.fromDrawingML = function(reader, oTable) {
		let sName;
		let oTableLook = null;
		while (reader.MoveToNextAttribute()) {
			sName = reader.GetNameNoNS();
			switch(sName) {
				case "bandCol": {
					if(!oTableLook) {
						oTableLook = new AscCommon.CTableLook();
					}
					oTableLook.BandVer = reader.GetValueBool();
					break;
				}
				case "bandRow": {
					if(!oTableLook) {
						oTableLook = new AscCommon.CTableLook();
					}
					oTableLook.BandHor = reader.GetValueBool();
					break;
				}
				case "firstCol": {
					if(!oTableLook) {
						oTableLook = new AscCommon.CTableLook();
					}
					oTableLook.FirstCol = reader.GetValueBool();
					break;
				}
				case "firstRow": {
					if(!oTableLook) {
						oTableLook = new AscCommon.CTableLook();
					}
					oTableLook.FirstRow = reader.GetValueBool();
					break;
				}
				case "lastCol": {
					if(!oTableLook) {
						oTableLook = new AscCommon.CTableLook();
					}
					oTableLook.LastCol = reader.GetValueBool();
					break;
				}
				case "lastRow": {
					if(!oTableLook) {
						oTableLook = new AscCommon.CTableLook();
					}
					oTableLook.LastRow = reader.GetValueBool();
					break;
				}
				case "rtl": {
					//TODO
					break;
				}
			}
		}
		if(oTable && oTableLook) {
			oTable.Set_TableLook(oTableLook);
		}
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if(AscFormat.CUniFill.prototype.isFillName(name)) {
				let oFill = new AscFormat.CUniFill();
				oFill.fromXml(reader, name);
				this.Shd = new CDocumentShd();
				this.Shd.Value = Asc.c_oAscShdClear;
				this.Shd.Unifill = oFill;
			}
			if(name === "tableStyleId") {
				let sStyleGUID =  reader.GetTextDecodeXml();
				let oStyle = reader.context.getTableStyle(sStyleGUID);
				if(oStyle) {
					oTable.Set_TableStyle2(oStyle.Id);
				}
			}
		}
	};
	CTablePr.prototype.toDrawingML = function(writer, oTable) {
		writer.WriteXmlNodeStart(("a:tblPr"));


		//writer.WriteXmlAttributeString("rtl", Rtl);
		let oTableLook = oTable.TableLook;
		if(oTableLook) {
			writer.WriteXmlAttributeBool("firstRow", oTableLook.FirstRow);
			writer.WriteXmlAttributeBool("firstCol", oTableLook.FirstCol);
			writer.WriteXmlAttributeBool("lastRow", oTableLook.LastRow);
			writer.WriteXmlAttributeBool("lastCol", oTableLook.LastCol);
			writer.WriteXmlAttributeBool("bandRow", oTableLook.BandHor);
			writer.WriteXmlAttributeBool("bandCol", oTableLook.BandVer);
		}

		writer.WriteXmlAttributesEnd();

		let sStyleId = oTable.TableStyle;
		let sGUID = writer.context.tableStylesIdToGuid[sStyleId];
		if (sGUID)
			writer.WriteXmlString("<a:tableStyleId>" + sGUID + "</a:tableStyleId>");

		if(this.Shd && this.Shd.Unifill) {
			this.Shd.Unifill.toXml(writer);
		}
		writer.WriteXmlNodeEnd("a:tblPr");
	};
	AscCommonWord.CTableRow.prototype.fromDrawingML = function (reader) {
		let sName;
		let fRowHeight = 5;
		while (reader.MoveToNextAttribute()) {
			sName = reader.GetNameNoNS();
			switch(sName) {
				case "h": {
					fRowHeight = reader.GetValueInt() / 36000;
					break;
				}

			}
		}

		let nCellsCount = 0;
		let depth = reader.GetDepth();
		let aCellStates = [];
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if(name === "tc") {
				++nCellsCount;
				aCellStates.push(reader.getState());
			}
		}

		let oEndState = reader.getState();

		let _count = aCellStates.length;
		_count = Math.min(_count, this.Content.length);
		for(let i = 0; i < aCellStates.length; ++i) {
			reader.setState(aCellStates[i]);
			let bIsNoHMerge = this.Content[i].fromDrawingML(reader);
			if (bIsNoHMerge === false) {
				this.Remove_Cell(i);
				aCellStates.splice(i, 1);
				i--;
				_count--;
			}
			let _gridCol = 1;
			if ("number" == typeof (this.Content[i].Pr.GridSpan)) {
				_gridCol = this.Content[i].Pr.GridSpan;
			}

			if (_gridCol > (_count - i)) {
				_gridCol = _count - i;
				this.Content[i].Pr.GridSpan = _gridCol;
				if (1 === this.Content[i].Pr.GridSpan)
					this.Content[i].Pr.GridSpan = undefined;
			}

			_gridCol--;
			while (_gridCol > 0) {
				i++;
				if (i >= _count)
					break;

				reader.setState(aCellStates[i]);
				this.Content[i].fromDrawingML(reader);

				this.Remove_Cell(i);
				aCellStates.splice(i, 1);
				i--;
				_count--;

				--_gridCol;
			}
		}
		this.updateHeightAfterOpen(fRowHeight);
		reader.setState(oEndState);
	};
	AscCommonWord.CTableRow.prototype.toDrawingML = function (writer, oInfo) {
		writer.WriteXmlNodeStart("a:tr");
		writer.WriteXmlAttributeUInt("h", AscCommon.GetTableRowHeight(this));
		writer.WriteXmlAttributesEnd();

		let nCellsCount = oInfo.Cells.length;
		for (let nCell = 0; nCell < nCellsCount; nCell++) {
			let oCellInfo = oInfo.Cells[nCell];
			if (oCellInfo.isEmpty) {
				this.writeEmptyCell(writer, oCellInfo.vMerge);
			}
			else {
				writer.WriteXmlNodeStart("a:tc");
				if (oCellInfo.vMerge === false && oCellInfo.row_span > 1) {
					writer.WriteXmlAttributeInt("rowSpan", oCellInfo.row_span);
				}
				if (oCellInfo.hMerge === false && oCellInfo.grid_span > 1) {
					writer.WriteXmlAttributeInt("gridSpan", oCellInfo.grid_span);
				}
				if (oCellInfo.hMerge === true) {
					writer.WriteXmlAttributeBool("hMerge", true);
				}
				if (oCellInfo.vMerge === true) {
					writer.WriteXmlAttributeBool("vMerge", true);
				}
				writer.WriteXmlAttributesEnd();
				writer.WriteXmlString("<a:txBody>");
				writer.WriteXmlString("<a:bodyPr/>");
				writer.WriteXmlString("<a:lstStyle/>");
				let nCount = oCellInfo.Cell.Content.Content.length;
				for (let i = 0; i < nCount; ++i)
					oCellInfo.Cell.Content.Content[i].toDrawingML(writer);
				writer.WriteXmlString("</a:txBody>");
				oCellInfo.Cell.Pr.toDrawingML(writer, this.Table.Pr);
				writer.WriteXmlNodeEnd("a:tc");


			}

		}
		writer.WriteXmlNodeEnd("a:tr");
	};
	AscCommonWord.CTableRow.prototype.writeEmptyCell = function(writer, bVMerge) {
		if(!bVMerge) {
			writer.WriteXmlString("<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></a:txBody><a:tcPr/></a:tc>");
		}
		else {
			writer.WriteXmlString("<a:tc vMerge=\"1\"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></a:txBody><a:tcPr/></a:tc>");
		}

	}
	AscCommonWord.CTableRow.prototype.updateHeightAfterOpen = function(fRowHeight) {
		let fMaxTopMargin = 0, fMaxBottomMargin = 0, fMaxTopBorder = 0, fMaxBottomBorder = 0;
		let bLoadVal = AscCommon.g_oIdCounter.m_bLoad;
		let bRead = AscCommon.g_oIdCounter.m_bRead;
		AscCommon.g_oIdCounter.m_bLoad = false;
		AscCommon.g_oIdCounter.m_bRead = false;
		for(let i = 0;  i < this.Content.length; ++i){
			let oCell = this.Content[i];
			let oMargins = oCell.GetMargins();
			if(oMargins.Bottom.W > fMaxBottomMargin){
				fMaxBottomMargin = oMargins.Bottom.W;
			}
			if(oMargins.Top.W > fMaxTopMargin){
				fMaxTopMargin = oMargins.Top.W;
			}
			let oBorders = oCell.Get_Borders();
			if(oBorders.Top.Size > fMaxTopBorder){
				fMaxTopBorder = oBorders.Top.Size;
			}
			if(oBorders.Bottom.Size > fMaxBottomBorder){
				fMaxBottomBorder = oBorders.Bottom.Size;
			}
		}
		AscCommon.g_oIdCounter.m_bLoad = bLoadVal;
		AscCommon.g_oIdCounter.m_bRead = bRead;
		this.Set_Height(Math.max(1, fRowHeight - fMaxTopMargin - fMaxBottomMargin - fMaxTopBorder/2 - fMaxBottomBorder/2), Asc.linerule_AtLeast);
	};
	AscCommonWord.CTableCell.prototype.fromDrawingML = function (reader) {
		let sName;
		while (reader.MoveToNextAttribute()) {
			sName = reader.GetNameNoNS();
			switch(sName) {
				case "gridSpan": {
					this.Set_GridSpan(reader.GetValueInt());
					break;
				}
				case "hMerge": {
					this.hMerge = reader.GetValueBool();
					if(this.hMerge) {
						return false;
					}
					break;
				}
				case "id": {
					let sId = reader.GetValue();
					break;
				}
				case "rowSpan": {
					let nVal = reader.GetValueInt();
					if(nVal > 1) {
						this.SetVMerge(vmerge_Restart);
					}
					break;
				}
				case "vMerge": {
					let bIsVMerge = reader.GetValueBool();
					if (bIsVMerge && this.Pr.VMerge != vmerge_Restart)
					{
						this.SetVMerge(vmerge_Continue);
					}
					break;
				}

			}
		}
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if(name === "tcPr") {
				let props = new CTableCellPr();
				props.fromDrawingML(reader);
				props.Merge(this.Pr);
				this.Set_Pr(props);
			}
			else if(name === "txBody") {
				AscFormat.CTextBody.prototype.fromXml(reader, false, this.Content);
			}
		}
		return true;
	};
	CTableCellPr.prototype.fromDrawingML = function (reader) {

		let sName;
		while (reader.MoveToNextAttribute()) {
			sName = reader.GetNameNoNS();
			switch(sName) {
				case "anchor": {
					let sVal = reader.GetValue();
					switch (sVal) {
						case "b": {
							this.VAlign = vertalignjc_Bottom;
							break;
						}
						case "ctr": {
							this.VAlign = vertalignjc_Center;
							break;
						}
						case "dist": {
							this.VAlign = vertalignjc_Center;
							break;
						}
						case "just": {
							this.VAlign = vertalignjc_Center;
							break;
						}
						case "t": {
							this.VAlign = vertalignjc_Top;
							break;
						}
					}
					break;
				}
				case "anchorCtr": {
					let sVal = reader.GetValue();
					break;
				}
				case "horzOverflow": {
					let sVal = reader.GetValue();
					break;
				}
				case "marB": {
					if(!this.TableCellMar)
						this.TableCellMar = {};
					this.TableCellMar.Bottom   = new CTableMeasurement(tblwidth_Mm, reader.GetValueInt() / 36000);
					break;
				}
				case "marL": {
					if(!this.TableCellMar)
						this.TableCellMar = {};
					this.TableCellMar.Left   = new CTableMeasurement(tblwidth_Mm, reader.GetValueInt() / 36000);
					break;
				}
				case "marR": {
					if(!this.TableCellMar)
						this.TableCellMar = {};
					this.TableCellMar.Right   = new CTableMeasurement(tblwidth_Mm, reader.GetValueInt() / 36000);
					break;
				}
				case "marT": {
					if(!this.TableCellMar)
						this.TableCellMar = {};
					this.TableCellMar.Top   = new CTableMeasurement(tblwidth_Mm, reader.GetValueInt() / 36000);
					break;
				}
				case "vert": {
					let sVal = reader.GetValue();
					switch (sVal) {
						case "eaVert": {
							this.TextDirection = Asc.c_oAscCellTextDirection.TBRL;
							break;
						}
						case "horz": {
							this.TextDirection = Asc.c_oAscCellTextDirection.LRTB;
							break;
						}
						case "mongolianVert": {
							this.TextDirection = Asc.c_oAscCellTextDirection.TBRL;
							break;
						}
						case "vert": {
							this.TextDirection = Asc.c_oAscCellTextDirection.BTLR;
							break;
						}
						case "vert270": {
							this.TextDirection = Asc.c_oAscCellTextDirection.BTLR;
							break;
						}
						case "wordArtVert": {
							this.TextDirection = Asc.c_oAscCellTextDirection.TBRL;
							break;
						}
						case "wordArtVertRtl": {
							this.TextDirection = Asc.c_oAscCellTextDirection.TBRL;
							break;
						}
					}
					break;
				}

			}
		}
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if (AscFormat.CUniFill.prototype.isFillName(name)) {
				let oFill = new AscFormat.CUniFill();
				oFill.fromXml(reader, name);
				this.Shd = new CDocumentShd();
				this.Shd.Value = Asc.c_oAscShdClear;
				this.Shd.Unifill = oFill;
			}
			else if (name === "cell3D") {

			}
			else if (name === "extLst") {

			}
			else if (name === "lnB") {
				if(!this.TableCellBorders) {
					this.TableCellBorders = {};
				}
				this.TableCellBorders.Bottom = readBorder(reader);
			}
			else if (name === "lnBlToTr") {

			}
			else if (name === "lnL") {
				if(!this.TableCellBorders) {
					this.TableCellBorders = {};
				}
				this.TableCellBorders.Left = readBorder(reader);
			}
			else if (name === "lnR") {
				if(!this.TableCellBorders) {
					this.TableCellBorders = {};
				}
				this.TableCellBorders.Right = readBorder(reader);
			}
			else if (name === "lnT") {
				if(!this.TableCellBorders) {
					this.TableCellBorders = {};
				}
				this.TableCellBorders.Top = readBorder(reader);
			}
			else if (name === "lnTlToBr") {

			}
		}
	};
	CTableCellPr.prototype.toDrawingML = function (writer, oTablePr) {
		writer.WriteXmlNodeStart("a:tcPr");

		let _marg = this.TableCellMar;
		let tableMar = oTablePr && oTablePr.TableCellMar;

		if(_marg && _marg.Left && AscFormat.isRealNumber(_marg.Left.W)) {
			writer.WriteXmlAttributeInt("marL", (_marg.Left.W * 36000) >> 0);
		}
		else if(tableMar && tableMar.Left && AscFormat.isRealNumber(tableMar.Left.W)) {
			writer.WriteXmlAttributeInt("marL", (tableMar.Left.W * 36000) >> 0);
		}

		if(_marg && _marg.Top && AscFormat.isRealNumber(_marg.Top.W)) {
			writer.WriteXmlAttributeInt("marT", (_marg.Top.W * 36000) >> 0);
		}
		else if(tableMar && tableMar.Top && AscFormat.isRealNumber(tableMar.Top.W)) {
			writer.WriteXmlAttributeInt("marT", (tableMar.Top.W * 36000) >> 0);
		}

		if(_marg && _marg.Right && AscFormat.isRealNumber(_marg.Right.W)) {
			writer.WriteXmlAttributeInt("marR", (_marg.Right.W * 36000) >> 0);
		}
		else if(tableMar && tableMar.Right && AscFormat.isRealNumber(tableMar.Right.W)) {
			writer.WriteXmlAttributeInt("marR", (tableMar.Right.W * 36000) >> 0);
		}

		if(_marg && _marg.Bottom && AscFormat.isRealNumber(_marg.Bottom.W)) {
			writer.WriteXmlAttributeInt("marB", (_marg.Bottom.W * 36000) >> 0);
		}
		else if(tableMar && tableMar.Bottom && AscFormat.isRealNumber(tableMar.Bottom.W)) {
			writer.WriteXmlAttributeInt("marB", (tableMar.Bottom.W * 36000) >> 0);
		}


		if(AscFormat.isRealNumber(this.TextDirection)) {
			switch (this.TextDirection) {
				case Asc.c_oAscCellTextDirection.LRTB: {
					writer.WriteXmlAttributeString("vert", "horz");
					break;
				}
				case Asc.c_oAscCellTextDirection.TBRL: {
					writer.WriteXmlAttributeString("vert", "eaVert");
					break;
				}
				case Asc.c_oAscCellTextDirection.BTLR: {
					writer.WriteXmlAttributeString("vert", "vert");
					break;
				}
				default: {
					writer.WriteXmlAttributeString("vert", "horz");
					break;
				}
			}
		}
		if(AscFormat.isRealNumber(this.VAlign)) {
			switch(this.VAlign) {
				case vertalignjc_Bottom: {
					writer.WriteXmlAttributeString("anchor", "b");
					break;
				}
				case vertalignjc_Center: {
					writer.WriteXmlAttributeString("anchor", "ctr");
					break;
				}
				case vertalignjc_Top: {
					writer.WriteXmlAttributeString("anchor", "t");
					break;
				}
			}
		}


		writer.WriteXmlAttributesEnd();

		let oBorders = this.TableCellBorders;

		if(oBorders.Left) {
			let oLn = new AscFormat.CLn();
			oLn.fromDocumentBorder(oBorders.Left);
			oLn.toXml(writer, "a:lnL");
		}
		if(oBorders.Right) {
			let oLn = new AscFormat.CLn();
			oLn.fromDocumentBorder(oBorders.Right);
			oLn.toXml(writer, "a:lnR");
		}
		if(oBorders.Top) {
			let oLn = new AscFormat.CLn();
			oLn.fromDocumentBorder(oBorders.Top);
			oLn.toXml(writer, "a:lnT");
		}
		if(oBorders.Bottom) {
			let oLn = new AscFormat.CLn();
			oLn.fromDocumentBorder(oBorders.Bottom);
			oLn.toXml(writer, "a:lnB");
		}
		if(this.Shd && this.Shd.Unifill) {
			this.Shd.Unifill.toXml(writer);
		}

		writer.WriteXmlNodeEnd("a:tcPr");
	};
	CParaSpacing.prototype.lnSpcFromDrawingML = function (reader) {
		let oSpc = readSpacing(reader);
		if(oSpc.valPct !== null) {
			this.Line = oSpc.valPct;
			this.LineRule = Asc.linerule_Auto;
		}
		else if(oSpc.val !== null) {
			this.Line = oSpc.val;
			this.LineRule = Asc.linerule_Exact;
		}
	};
	CParaSpacing.prototype.lnSpcToDrawingML = function (writer) {
		if(this.Line !== undefined && this.Line !== null) {
			if(this.LineRule === Asc.linerule_Exact) {
				writeSpacing(writer, "a:lnSpc", {val: this.Line});
			}
			else {
				writeSpacing(writer, "a:lnSpc", {valPct: this.Line});
			}
		}
	};
	CParaSpacing.prototype.spcAftFromDrawingML = function (reader) {
		let oSpc = readSpacing(reader);
		if(oSpc.valPct !== null) {
			this.AfterPct = oSpc.valPct;
			this.After = 0;
		}
		else if(oSpc.val !== null) {
			this.After = oSpc.val;
		}
	};
	CParaSpacing.prototype.spcAftToDrawingML = function (writer) {
		if(this.After !== null && this.After !== undefined) {
			writeSpacing(writer, "a:spcAft", {val: this.After})
		}
		else if(this.AfterPct !== null && this.AfterPct !== undefined) {
			writeSpacing(writer, "a:spcAft", {valPct: this.AfterPct});
		}
	};
	CParaSpacing.prototype.spcBefFromDrawingML = function (reader) {
		let oSpc = readSpacing(reader);
		if(oSpc.valPct !== null) {
			this.BeforePct = oSpc.valPct;
			this.Before = 0;
		}
		else if(oSpc.val !== null) {
			this.Before = oSpc.val;
		}
	};
	CParaSpacing.prototype.spcBefToDrawingML = function (writer) {
		if(this.Before !== null && this.Before !== undefined) {
			writeSpacing(writer, "a:spcBef", {val: this.Before})
		}
		else if(this.BeforePct !== null && this.BeforePct !== undefined) {
			writeSpacing(writer, "a:spcBef", {valPct: this.BeforePct});
		}
	};
	CParaTabs.prototype.fromDrawingML = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if(name === "tab") {
				let oTab = new CParaTab();
				oTab.fromDrawingML(reader);
				this.Tabs.push(oTab);
			}
		}
	};
	CParaTabs.prototype.toDrawingML = function(writer) {
		writer.WriteXmlNodeStart("a:tabLst");
		if(this.Tabs.length > 0) {
			writer.WriteXmlAttributesEnd();
			for(let nIdx = 0; nIdx < this.Tabs.length; ++nIdx) {
				this.Tabs[nIdx].toDrawingML(writer);
			}
			writer.WriteXmlNodeEnd("a:tabLst");
		}
		else {
			writer.WriteXmlAttributesEnd(true);
		}
	};
	CParaTab.prototype.fromDrawingML = function(reader) {
		let sName;
		while (reader.MoveToNextAttribute()) {
			sName = reader.GetNameNoNS();
			switch(sName) {
				case "algn": {
					let sVal = reader.GetValue();
					if(sVal === "ctr") {
						this.Value = tab_Center;
					}
					else if(sVal === "r") {
						this.Value = tab_Right;
					}
					else {
						this.Value = tab_Left;
					}
					break;
				}
				case "pos": {
					this.Pos = reader.GetValueInt() / 36000;
					break;
				}
			}
		}
	};
	CParaTab.prototype.toDrawingML = function(writer) {
		writer.WriteXmlNodeStart("a:tab");
		let sAlign = "l";
		if(this.Value === tab_Center) {
			sAlign = "ctr";
		}
		else if(this.Value === tab_Right) {
			sAlign = "r";
		}
		writer.WriteXmlAttributeString("algn", sAlign );
		writer.WriteXmlAttributeString("pos", this.Pos * 36000 + 0.5 >> 0 );
		writer.WriteXmlAttributesEnd(true);
	};
	CStyles.prototype.fromDrawingML = function(reader) {
		if (!reader.ReadNextNode()) {
			return null;
		}
		let sDefaultTableStyleGuid = null;
		while (reader.MoveToNextAttribute()) {
			let sName = reader.GetNameNoNS();
			if(sName === "def") {
				sDefaultTableStyleGuid = reader.GetValue();
			}
		}
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let sName = reader.GetNameNoNS();
			if(sName === "tblStyle") {
				let oTableStyle = new AscCommonWord.CStyle("", null, null, styletype_Table);
				oTableStyle.fromDrawingMLTableStyle(reader);
				this.Add(oTableStyle);
			}
			//this.readChildXml(name, reader);
		}
		return sDefaultTableStyleGuid;
	};
	CStyles.prototype.toDrawingML = function (writer, oIdMap, sDefault) {

		//generate guids for table styles
		let oTableStylesIdToGuid = writer.context.tableStylesIdToGuid;
		for(let sKey in oIdMap) {
			if(oIdMap.hasOwnProperty(sKey)) {
				oTableStylesIdToGuid[sKey] = "{" + AscCommon.GUID() + "}"
			}
		}

		writer.WriteXmlString(AscCommonWord.g_sXmlHeader);
		writer.WriteXmlNodeStart("a:tblStyleLst");
		writer.WriteXmlAttributeString("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
		writer.WriteXmlAttributeString("def", oTableStylesIdToGuid[sDefault]);
		writer.WriteXmlAttributesEnd();
		for(let sKey in oTableStylesIdToGuid) {
			if(oIdMap.hasOwnProperty(sKey)) {
				let oTableStyle = AscCommon.g_oTableId.Get_ById(sKey);
				oTableStyle.toDrawingML(writer, oTableStylesIdToGuid[sKey]);
			}
		}

		writer.WriteXmlNodeEnd("a:tblStyleLst");
	};
	AscCommonWord.CStyle.prototype.fromDrawingMLTableStyle = function(reader) {
		let sStyleGuid = null;
		while (reader.MoveToNextAttribute()) {
			let sName = reader.GetNameNoNS();
			if(sName === "styleId") {
				sStyleGuid = reader.GetValue();
				reader.context.addTableStyle(sStyleGuid, this);
			}
			else if(sName === "styleName") {
				this.Set_Name(reader.GetValue());
			}
		}
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let sName = reader.GetNameNoNS();
			switch (sName) {
				case "tblBg": {
					this.readTableCellProperties(reader, this.TablePr);
					break;
				}
				case "band1H": {
					this.TableBand1Horz = this.readDrawingMLTableStylePart(reader);
					break;
				}
				case "band1V": {
					this.TableBand1Vert = this.readDrawingMLTableStylePart(reader);
					break;
				}
				case "band2H": {
					this.TableBand2Horz = this.readDrawingMLTableStylePart(reader);
					break;
				}
				case "band2V": {
					this.TableBand2Vert = this.readDrawingMLTableStylePart(reader);
					break;
				}
				case "firstCol": {
					this.TableFirstCol = this.readDrawingMLTableStylePart(reader);
					break;
				}
				case "firstRow": {
					this.TableFirstRow = this.readDrawingMLTableStylePart(reader);
					break;
				}
				case "lastCol": {
					this.TableLastCol = this.readDrawingMLTableStylePart(reader);
					break;
				}
				case "lastRow": {
					this.TableLastRow = this.readDrawingMLTableStylePart(reader);
					break;
				}
				case "wholeTbl": {
					this.TableWholeTable = this.readDrawingMLTableStylePart(reader);
					break;
				}
				case "neCell": {
					this.TableTRCell = this.readDrawingMLTableStylePart(reader);
					break;
				}
				case "nwCell": {
					this.TableTLCell = this.readDrawingMLTableStylePart(reader);
					break;
				}
				case "seCell": {
					this.TableBRCell = this.readDrawingMLTableStylePart(reader);
					break;
				}
				case "swCell": {
					this.TableBLCell = this.readDrawingMLTableStylePart(reader);
					break;
				}
			}
		}
		this.wholeToTablePr();
	};
	AscCommonWord.CStyle.prototype.readDrawingMLTableStylePart = function(reader) {
		let oPart =  new CTableStylePr();
		let oCellPr = oPart.TableCellPr;
		let oTextPr = oPart.TextPr;
		let oTablePr = oPart.TablePr;
		let oCellBorders = oCellPr.TableCellBorders;
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let sName = reader.GetNameNoNS();
			switch (sName) {
				case "tcStyle": {
					this.readTableCellProperties(reader, oCellPr);
					oTablePr.TableBorders.InsideH = oCellBorders.InsideH;
					oTablePr.TableBorders.InsideV = oCellBorders.InsideV;
					delete oCellBorders.InsideH;
					delete oCellBorders.InsideV;
					break;
				}
				case "tcTxStyle": {
					let oTxStyleNode = new CT_XmlNode(function(reader, name) {
						if(AscFormat.CUniColor.prototype.isUnicolor(name)) {
							let oColor = new AscFormat.CUniColor();
							oColor.fromXml(reader, name);
							oTextPr.Unifill = AscFormat.CreateUniFillByUniColor(oColor);
						}
						else if(name === "fontRef") {
							let oFontRef = new AscFormat.FontRef();
							oFontRef.fromXml(reader);
							oTextPr.FontRef = oFontRef;
						}
						else if(name === "font") {
							let oFontNode = new CT_XmlNode(function(reader, name) {
								readRFont(oTextPr, reader, name);
								return true;
							});
							oFontNode.fromXml(reader);
						}
						return true;
					});
					oTxStyleNode.fromXml(reader);
					let sBoldAttr = oTxStyleNode.attributes["b"];
					oTextPr.Bold = reader.GetBool(sBoldAttr);
					let sItalicAttr = oTxStyleNode.attributes["i"];
					oTextPr.Italic = reader.GetBool(sItalicAttr);
					break;
				}
			}
		}
		return oPart;
	};
	AscCommonWord.CStyle.prototype.readTableCellProperties = function(reader, oCellPr) {
		let oBorders = oCellPr.TableCellBorders;
		let oNode = new CT_XmlNode(function(reader, name) {
			if(name === "cell3D") {
				//TODO: implement
			}
			else if(name === "fill") {
				let oFillNode = new CT_XmlNode(function(reader, name) {
					if(AscFormat.CUniFill.prototype.isFillName(name)) {
						let oFill = new AscFormat.CUniFill();
						oFill.fromXml(reader, name);
						if(!oCellPr.Shd) {
							oCellPr.Shd = new AscCommonWord.CDocumentShd();
							oCellPr.Shd.Value = c_oAscShdClear;
						}
						oCellPr.Shd.Unifill = oFill;
					}
					return true;
				});
				oFillNode.fromXml(reader);
			}
			else if(name === "fillRef") {
				let oStyleRef = new AscFormat.StyleRef();
				oStyleRef.fromXml(reader);
				if(!oCellPr.Shd) {
					oCellPr.Shd = new AscCommonWord.CDocumentShd();
					oCellPr.Shd.Value = c_oAscShdClear;
				}
				oCellPr.Shd.FillRef = oStyleRef;
			}
			else if(name === "tcBdr") {
				if(oBorders) {
					let oBordersNode = new CT_XmlNode(function(reader, name) {
						let oBorder = new CDocumentBorder();
						let oBorderNode = new CT_XmlNode(function(reader, name) {
							if(name === "ln") {
								let oLn = new AscFormat.CLn();
								oLn.fromXml(reader);
								oLn.fillDocumentBorder(oBorder)
							}
							else if(name === "lnRef") {
								let oStyleRef = new AscFormat.StyleRef();
								oStyleRef.fromXml(reader);
								oBorder.LineRef = oStyleRef;
								oBorder.Value = AscCommonWord.border_Single;
							}
							return true;
						});
						oBorderNode.fromXml(reader);
						switch (name) {
							case "bottom": {
								oBorders.Bottom = oBorder;
								break;
							}
							case "insideH": {
								oBorders.InsideH = oBorder;
								break;
							}
							case "insideV": {
								oBorders.InsideV = oBorder;
								break;
							}
							case "left": {
								oBorders.Left = oBorder;
								break;
							}
							case "right": {
								oBorders.Right = oBorder;
								break;
							}
							case "tl2br": {
								break;
							}
							case "top": {
								oBorders.Top = oBorder;
								break;
							}
							case "tr2bl": {
								break;
							}
						}
						return true;
					});
					oBordersNode.fromXml(reader);
				}
			}
			return true;
		});
		oNode.fromXml(reader);
	};
	AscCommonWord.CStyle.prototype.toDrawingML = function(writer, sGUID) {
		writer.WriteXmlNodeStart("a:tblStyle");
		writer.WriteXmlAttributeString("styleId", sGUID);
		writer.WriteXmlAttributeString("styleName", this.Name);
		writer.WriteXmlAttributesEnd();
		this.writeTcStyle(writer, this.TablePr, "a:tblBg");
		this.toDrawingMLStylePart(writer, this.TableWholeTable, "a:wholeTbl");
		this.toDrawingMLStylePart(writer, this.TableBand1Horz, "a:band1H");
		this.toDrawingMLStylePart(writer, this.TableBand2Horz, "a:band2H");
		this.toDrawingMLStylePart(writer, this.TableBand1Vert, "a:band1V");
		this.toDrawingMLStylePart(writer, this.TableBand2Vert, "a:band2V");
		this.toDrawingMLStylePart(writer, this.TableLastCol, "a:lastCol");
		this.toDrawingMLStylePart(writer, this.TableFirstCol, "a:firstCol");
		this.toDrawingMLStylePart(writer, this.TableLastRow, "a:lastRow");
		this.toDrawingMLStylePart(writer, this.TableBRCell, "a:seCell");
		this.toDrawingMLStylePart(writer, this.TableBLCell, "a:swCell");
		this.toDrawingMLStylePart(writer, this.TableFirstRow, "a:firstRow");
		this.toDrawingMLStylePart(writer, this.TableTRCell, "a:neCell");
		this.toDrawingMLStylePart(writer, this.TableTLCell, "a:nwCell");
		writer.WriteXmlNodeEnd("a:tblStyle");
	};
	AscCommonWord.CStyle.prototype.toDrawingMLStylePart = function(writer, oPart, sName) {
		writer.WriteXmlNodeStart(sName);
		writer.WriteXmlAttributesEnd();
		this.writeTcTxStyle(writer, oPart.TextPr);
		this.writeTcStyle(writer, oPart.TableCellPr, "a:tcStyle");
		writer.WriteXmlNodeEnd(sName);
	};
	AscCommonWord.CStyle.prototype.writeTcTxStyle = function (writer, oTextPr) {

		let oUnicolor = oTextPr.Unifill && oTextPr.Unifill.fill && oTextPr.Unifill.fill.color;
		if(oTextPr.Italic === true || oTextPr.Italic === false ||
			oTextPr.Bold === true || oTextPr.Bold === false ||
			oTextPr.FontRef || oUnicolor) {
			writer.WriteXmlNodeStart("a:tcTxStyle");
			if(oTextPr.Italic === true || oTextPr.Italic === false) {
				writer.WriteXmlAttributeString("i", oTextPr.Italic ? "on" : "off");
			}
			if(oTextPr.Bold === true || oTextPr.Bold === false) {
				writer.WriteXmlAttributeString("b", oTextPr.Bold ? "on" : "off");
			}
			writer.WriteXmlAttributesEnd();

			if(oTextPr.FontRef) {
				oTextPr.FontRef.toXml(writer, "a:fontRef")
			}
			if(oUnicolor) {
				oUnicolor.toXml(writer);
			}
			writer.WriteXmlNodeEnd("a:tcTxStyle");
		}

	};
	AscCommonWord.CStyle.prototype.writeTcStyle = function (writer, oCellPr, sName) {

		let oBorders = oCellPr.TableCellBorders;
		let oShd = oCellPr.Shd;
		let bBorders = oBorders && (oBorders.Left || oBorders.Right || oBorders.Top ||  oBorders.Bottom ||  oBorders.InsideH || oBorders.InsideV);
		let bBg = sName === "a:tblBg"
		if(bBorders ||
			oCellPr.Shd && (oShd.FillRef || oShd.Unifill)) {

			writer.WriteXmlNodeStart(sName);
			writer.WriteXmlAttributesEnd();
			if(!bBg) {
				if(bBorders) {
					writer.WriteXmlNodeStart("a:tcBdr");
					writer.WriteXmlAttributesEnd();
					this.writeTcBorder(writer, oBorders.Left, "a:left");
					this.writeTcBorder(writer, oBorders.Right, "a:right");
					this.writeTcBorder(writer, oBorders.Top, "a:top");
					this.writeTcBorder(writer, oBorders.Bottom, "a:bottom");
					this.writeTcBorder(writer, oBorders.InsideH, "a:insideH");
					this.writeTcBorder(writer, oBorders.InsideV, "a:insideV");
					writer.WriteXmlNodeEnd("a:tcBdr");
				}
				else {
					writer.WriteXmlNodeStart("a:tcBdr");
					writer.WriteXmlAttributesEnd(true);
				}
			}
			if(oShd) {
				if(oShd.FillRef) {
					oShd.FillRef.toXml(writer, "a:fillRef");
				}
				if(oShd.Unifill) {
					writer.WriteXmlNodeStart("a:fill");
					writer.WriteXmlAttributesEnd();
					oShd.Unifill.toXml(writer);
					writer.WriteXmlNodeEnd("a:fill");
				}
			}
			//writer.Write(cell3D);

			writer.WriteXmlNodeEnd(sName);
		}
		else {
			if(sName !== "a:tblBg") {
				writer.WriteXmlNodeStart(sName);
				writer.WriteXmlAttributesEnd();
				writer.WriteXmlNodeStart("a:tcBdr");
				writer.WriteXmlAttributesEnd(true);
				writer.WriteXmlNodeEnd(sName);
			}
		}

	};
	AscCommonWord.CStyle.prototype.writeTcBorder = function (writer, oBorder, sName) {
		if(oBorder) {
			writer.WriteXmlNodeStart(sName);
			writer.WriteXmlAttributesEnd();
			if(oBorder.LineRef) {
				oBorder.LineRef.toXml(writer, "a:lnRef");
			}
			else {
				let oLn = new AscFormat.CLn();
				oLn.fromDocumentBorder(oBorder);
				oLn.toXml(writer, "a:ln");
			}
			writer.WriteXmlNodeEnd(sName);
		}
	};
	AscCommonWord.CStyle.prototype.wholeToTablePr = function() {

		let oWhole = this.TableWholeTable;
		if(!oWhole) {
			return
		}
		let oWholeBorders = oWhole.TablePr && oWhole.TablePr.TableBorders;
		if(!oWholeBorders) {
			return;
		}
		let oWholeCellBorders = oWhole.TableCellPr && oWhole.TableCellPr.TableCellBorders;
		if(!oWholeCellBorders) {
			return;
		}

		let oTablePBorders = this.TablePr && this.TablePr.TableBorders;
		if(!oTablePBorders) {
			return;
		}

		if(oWholeBorders.InsideH) {
			oTablePBorders.InsideH = oWholeBorders.InsideH;
			delete oWholeBorders.InsideH;
		}
		if(oWholeBorders.InsideV) {
			oTablePBorders.InsideV = oWholeBorders.InsideV;
			delete oWholeBorders.InsideV;
		}
		if(oWholeCellBorders.Top) {
			oTablePBorders.Top = oWholeCellBorders.Top;
			delete oWholeCellBorders.Top;
		}
		if(oWholeCellBorders.Bottom) {
			oTablePBorders.Bottom = oWholeCellBorders.Bottom;
			delete oWholeCellBorders.Bottom;
		}
		if(oWholeCellBorders.Left) {
			oTablePBorders.Left = oWholeCellBorders.Left;
			delete oWholeCellBorders.Left;
		}
		if(oWholeCellBorders.Right) {
			oTablePBorders.Right = oWholeCellBorders.Right;
			delete oWholeCellBorders.Right;
		}
	};
	let SPACING_SCALE = 0.00352777778;
	function readBorder(reader) {
		let oLn = new AscFormat.CLn();
		oLn.fromXml(reader);
		let border = new CDocumentBorder();
		if(oLn.Fill)
		{
			border.Unifill = oLn.Fill;
		}
		border.Size = (oLn.w == null) ? 12700 : ((oLn.w) >> 0);
		border.Size /= 36000;

		border.Value = border_Single;
		return border;
	}
	function readSpacing(reader) {
		let sName;
		let depth = reader.GetDepth();
		let oRet = {val: null, valPct: null};
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if(name === "spcPct") {
				while (reader.MoveToNextAttribute()) {
					sName = reader.GetNameNoNS();
					switch(sName) {
						case "val": {
							let nVal = AscFormat.getPercentageValue(reader.GetValue());
							if(nVal !== null) {
								oRet.valPct = nVal / 100;
							}
							break;
						}
					}
				}
			}
			else if(name === "spcPts") {
				while (reader.MoveToNextAttribute()) {
					sName = reader.GetNameNoNS();
					switch(sName) {
						case "val": {
							let nVal = reader.GetValueInt()
							oRet.val = nVal * SPACING_SCALE;
							break;
						}
					}
				}
			}
		}
		return oRet;
	}
	function writeSpacing(writer, sName, oSpacing) {
		writer.WriteXmlNodeStart(sName);
		writer.WriteXmlAttributesEnd();
		if(oSpacing.valPct !== undefined && oSpacing.valPct !== null) {
			writer.WriteXmlNodeStart("a:spcPct");
			writer.WriteXmlAttributeString("val", (oSpacing.valPct * 100000 + 0.5 >> 0) + "");
			writer.WriteXmlAttributesEnd(true);
		}
		else if(oSpacing.val !== undefined && oSpacing.val !== null) {
			writer.WriteXmlNodeStart("a:spcPts");
			writer.WriteXmlAttributeString("val", (oSpacing.val / SPACING_SCALE + 0.5 >> 0) + "");
			writer.WriteXmlAttributesEnd(true);
		}
 		writer.WriteXmlNodeEnd(sName);
	}
	function readTypeface(reader) {
		let sName;
		while (reader.MoveToNextAttribute()) {
			sName = reader.GetNameNoNS();
			switch(sName) {
				case "typeface": {
					return reader.GetValue();
				}
			}
		}
		return null;
	}
	function readRFont(oTextPr, reader, name) {
		if(name === "cs") {
			let sName = readTypeface(reader);
			if(sName) {
				oTextPr.RFonts.CS = { Name: sName, Index : -1 };
			}
		}
		else if(name === "ea") {
			let sName = readTypeface(reader);
			if(sName) {
				oTextPr.RFonts.EastAsia = { Name: sName, Index : -1 };
			}
		}
		else if(name === "latin") {
			let sName = readTypeface(reader);
			if(sName) {
				oTextPr.RFonts.Ascii = { Name: sName, Index : -1 };
				oTextPr.RFonts.HAnsi = { Name: sName, Index : -1 };
			}
		}
	}
	function writeTypeface(writer, sName, sFont) {
		if(typeof sFont === "string") {
			AscFormat.FontCollection.prototype.writeFont(writer, sName, sFont)
		}
	}
	function readHighlightColor(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			if(AscFormat.CUniColor.prototype.isUnicolor(name)) {
				let oColor = new AscFormat.CUniColor();
				oColor.fromXml(reader, name);
				return oColor;
			}
		}
		return null;
	}

	window['AscCommon'].CT_ComplexType = CT_ComplexType;
	window['AscCommon'].CT_Bool = CT_Bool;
	window['AscCommon'].CT_String = CT_String;
	window['AscCommon'].CT_Int = CT_Int;
	window['AscCommon'].CT_UInt = CT_UInt;
	window['AscCommon'].CT_Double = CT_Double;

	window['AscCommon'].CT_BoolW = CT_BoolW;
	window['AscCommon'].CT_StringW = CT_StringW;
	window['AscCommon'].CT_IntW = CT_IntW;
	window['AscCommon'].CT_UIntW = CT_UIntW;
	window['AscCommon'].CT_DoubleW = CT_DoubleW;
})(window);

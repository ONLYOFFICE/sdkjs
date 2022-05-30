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

(
	/**
	 * @param {Window} window
	 * @param {undefined} undefined
	 */
	function (window, undefined) {

		let CBaseId = AscFormat.CBaseFormatObject;
		let CBaseNoId = AscFormat.CBaseNoIdObject;
		let IC = AscFormat.InitClass;

		function getBooleanTrueFalse(bVal) {
			if (bVal === true) {
				return "t";
			}
			if (bVal === false) {
				return "f";
			}
			return null;
		}

		//VmlOfficeDrawing
		function CStrokeChild(sType) {
			CBaseNoId.call(this);

			this.m_sType = sType;
			this.m_sAlthref = null;
			this.m_oColor = null;
			this.m_oColor2 = null;
			this.m_oDashStyle = null;
			this.m_oEndArrow = null;
			this.m_oEndArrowLength = null;
			this.m_oEndArrowWidth = null;
			this.m_oEndCap = null;
			this.m_oExt = null;
			this.m_oFillType = null;
			this.m_oForceDash = null;
			this.m_sHref = null;
			this.m_oImageAlignShape = null;
			this.m_oImageAspect = null;
			this.m_sImageSize = null;
			this.m_oInsetPen = null;
			this.m_oJoinStyle = null;
			this.m_oLineStyle = null;
			this.m_oMiterLimit = null;
			this.m_oOn = null;
			this.m_sOpacity = null;
			this.m_sSrc = null;
			this.m_oStartArrow = null;
			this.m_oStartArrowLength = null;
			this.m_oStartArrowWidth = null;
			this.m_sTitle = null;
			this.m_sWeight = null;
		}

		IC(CStrokeChild, CBaseNoId, 0);
		CStrokeChild.prototype.readAttrXml = function (name, reader) {
			if ("althref") this.m_sAlthref = reader.GetValue();
			else if ("color") this.m_oColor = readColorType(reader);
			else if ("color2") this.m_oColor2 = readColorType(reader);
			else if ("dashstyle") this.m_oDashStyle = readDashStyle(reader);
			else if ("endarrow") this.m_oEndArrow = readArrowType(reader);
			else if ("endarrowlength") this.m_oEndArrowLength = readArrowLength(reader);
			else if ("endarrowwidth") this.m_oEndArrowWidth = readArrowWidth(reader);
			else if ("endcap") this.m_oEndCap = readEndCap(reader);
			else if ("ext") this.m_oExt = readExt(reader);
			else if ("filltype") this.m_oFillType = readFillType(reader);
			else if ("forcedash") this.m_oForceDash = reader.GetValueBool();
			else if ("href") this.m_sHref = reader.GetValue();
			else if ("imagealignshape") this.m_oImageAlignShape = reader.GetValueBool();
			else if ("imageaspect") this.m_oImageAspect = readImageAspect(reader)
			else if ("imagesize") this.m_sImageSize = reader.GetValue();
			else if ("insetpen") this.m_oInsetPen = reader.GetValueBool();
			else if ("joinstyle") this.m_oJoinStyle = readJoinStyle(reader);
			else if ("linestyle") this.m_oLineStyle = readLineStyle(reader);
			else if ("miterlimit") this.m_oMiterLimit = reader.GetValueInt()
			else if ("on") this.m_oOn = reader.GetValueBool();
			else if ("opacity") this.m_sOpacity = reader.GetValue();
			else if ("src") this.m_sSrc = reader.GetValue();
			else if ("startarrow") this.m_oStartArrow = readArrowType(reader);
			else if ("startarrowlength") this.m_oStartArrowLength = readArrowLength(reader);
			else if ("startarrowwidth") this.m_oStartArrowWidth = readArrowWidth(reader);
			else if ("title") this.m_sTitle = reader.GetValue();
			else if ("weight") this.m_sWeight = reader.GetValue();
		};
		CStrokeChild.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
			writer.WriteXmlNullableAttributeString("on", getBooleanTrueFalse(this.m_oOn));
			writer.WriteXmlNullableAttributeString("weight", this.m_sWeight)
			writer.WriteXmlNullableAttributeString("color", getColorType(this.m_oColor))
			writer.WriteXmlNullableAttributeString("color2", getColorType(this.m_oColor2))
			writer.WriteXmlNullableAttributeString("opacity", this.m_sOpacity)
			writer.WriteXmlNullableAttributeString("linestyle", getLineStyle(this.m_oLineStyle))
			writer.WriteXmlNullableAttributeString("miterlimit", this.m_oMiterLimit + "")
			writer.WriteXmlNullableAttributeString("joinstyle", getJoinStyle(this.m_oJoinStyle))
			writer.WriteXmlNullableAttributeString("endcap", getEndCap(this.m_oEndCap))
			writer.WriteXmlNullableAttributeString("dashstyle", getDashStyle(this.m_oDashStyle))
			writer.WriteXmlNullableAttributeString("insetpen", getBooleanTrueFalse(this.m_oInsetPen))
			writer.WriteXmlNullableAttributeString("filltype", getFillType(this.m_oFillType))
			writer.WriteXmlNullableAttributeString("src", this.m_sSrc)
			writer.WriteXmlNullableAttributeString("imageaspect", getImageAspect(this.m_oImageAspect))
			writer.WriteXmlNullableAttributeString("imagesize", this.m_sImageSize)
			writer.WriteXmlNullableAttributeString("imagealignshape", getBooleanTrueFalse(this.m_oImageAlignShape))
			writer.WriteXmlNullableAttributeString("startarrow", getArrowType(this.m_oStartArrow))
			writer.WriteXmlNullableAttributeString("startarrowwidth", getArrowWidth(this.m_oStartArrowWidth))
			writer.WriteXmlNullableAttributeString("startarrowlength", getArrowLength(this.m_oStartArrowLength))
			writer.WriteXmlNullableAttributeString("endarrow", getArrowType(this.m_oEndArrow))
			writer.WriteXmlNullableAttributeString("endarrowwidth", getArrowWidth(this.m_oEndArrowWidth))
			writer.WriteXmlNullableAttributeString("endarrowlength", getArrowLength(this.m_oEndArrowLength))
			writer.WriteXmlNullableAttributeString("o:href", this.m_sHref)
			writer.WriteXmlNullableAttributeString("althref", this.m_sAlthref)
			writer.WriteXmlNullableAttributeString("o:title", this.m_sTitle)
			writer.WriteXmlNullableAttributeString("o:forcedash", getBooleanTrueFalse(this.m_oForceDash))
		};

		function CCallout() {
			CBaseNoId.call(this);
			this.m_oAccentbar = null;
			this.m_oAngle = null;
			this.m_oDistance = null;
			this.m_oDrop = null;
			this.m_oDropAuto = null;
			this.m_oExt = null;
			this.m_oGap = null;
			this.m_oLength = null;
			this.m_oLengthSpecified = null;
			this.m_oMinusX = null;
			this.m_oMinusY = null;
			this.m_oOn = null;
			this.m_oTextBorder = null;
			this.m_oType = null;
		}

		IC(CCallout, CBaseNoId, 0);
		CCallout.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "accentbar": {
					this.m_oAccentbar = reader.GetValueBool();
					break;
				}
				case "angle": {
					let sVal = reader.GetValue();
					switch (sVal) {
						case "30": {
							this.m_oAngle = EVmlAngle.vmlangle30;
							break
						}
						case "45": {
							this.m_oAngle = EVmlAngle.vmlangle45;
							break
						}
						case "60": {
							this.m_oAngle = EVmlAngle.vmlangle60;
							break
						}
						case "90": {
							this.m_oAngle = EVmlAngle.vmlangle90;
							break
						}
						case "any": {
							this.m_oAngle = EVmlAngle.vmlangleAny;
							break
						}
						case "auto": {
							this.m_oAngle = EVmlAngle.vmlangleAuto;
							break
						}
						default: {
							this.m_oAngle = EVmlAngle.vmlangleAuto;
							break
						}
					}
					break;
				}
				case "distance": {
					this.m_oDistance = reader.GetValueInt();
					break;
				}
				case "drop": {
					this.m_oDrop = reader.GetValue();
					break;
				}
				case "dropauto": {
					this.m_oDropAuto = reader.GetValueBool();
					break;
				}
				case "ext": {
					this.m_oExt = readExt(reader);
					break;
				}
				case "gap": {
					this.m_oGap = reader.GetValueInt();
					break;
				}
				case "length": {
					this.m_oLength = reader.GetValueInt();
					break;
				}
				case "lengthspecified": {
					this.m_oLengthSpecified = reader.GetValueBool();
					break;
				}
				case "minusx": {
					this.m_oMinusX = reader.GetValueBool();
					break;
				}
				case "minusy": {
					this.m_oMinusY = reader.GetValueBool();
					break;
				}
				case "on": {
					this.m_oOn = reader.GetValueBool();
					break;
				}
				case "textborder": {
					this.m_oTextBorder = reader.GetValueBool();
					break;
				}
				case "type": {
					let sVal = reader.GetValue();
					switch (sVal) {
						case "rectangle": {
							this.m_oType = EVmlCalloutType.vmlcallouttypeRectangle;
							break;
						}
						case "roundedrectangle": {
							this.m_oType = EVmlCalloutType.vmlcallouttypeRoundRectangle;
							break;
						}
						case "oval": {
							this.m_oType = EVmlCalloutType.vmlcallouttypeOval;
							break;
						}
						case "cloud": {
							this.m_oType = EVmlCalloutType.vmlcallouttypeCloud;
							break;
						}
					}
					break;
				}
			}
		};
		CCallout.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
			writer.WriteXmlNullableAttributeString("on", getBooleanTrueFalse(this.m_oOn));
			writer.WriteXmlNullableAttributeString("type", getCalloutType(this.m_oType));
			writer.WriteXmlNullableAttributeNumber("gap", this.m_oGap);
			writer.WriteXmlNullableAttributeString("angle", getAngle(this.m_oAngle));
			writer.WriteXmlNullableAttributeString("dropauto", getBooleanTrueFalse(this.m_oDropAuto));
			writer.WriteXmlNullableAttributeString("drop", this.m_oDrop);//TODO: use type
			writer.WriteXmlNullableAttributeNumber("distance", this.m_oDistance);
			writer.WriteXmlNullableAttributeString("lengthspecified", getBooleanTrueFalse(this.m_oLengthSpecified));
			writer.WriteXmlNullableAttributeNumber("length", this.m_oLength);
			writer.WriteXmlNullableAttributeString("accentbar", getBooleanTrueFalse(this.m_oAccentbar));
			writer.WriteXmlNullableAttributeString("textborder", getBooleanTrueFalse(this.m_oTextBorder));
			writer.WriteXmlNullableAttributeString("minusx", getBooleanTrueFalse(this.m_oMinusX));
			writer.WriteXmlNullableAttributeString("minusy", getBooleanTrueFalse(this.m_oMinusY));
		};

		function CClipPath() {
			CBaseNoId.call(this);
			this.m_oV = null;
		}

		IC(CClipPath, CBaseNoId, 0);
		CClipPath.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "v": {
					this.m_oV = reader.GetValue();
					break;
				}
			}
		};
		CClipPath.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("o:v", this.m_oV)
		};

		function CColorMenu() {
			CBaseNoId.call(this);
			this.m_oExt = null;
			this.m_oExtrusionColor = null;
			this.m_oFillColor = null;
			this.m_oShadowColor = null;
			this.m_oStrokeColor = null;
		}

		IC(CColorMenu, CBaseNoId, 0);
		CColorMenu.prototype.readAttrXml = function (name, reader) {
			if ("ext" === name) this.m_oExt = readExt(reader);
			if ("extrusioncolor" === name) this.m_oExtrusionColor = readColorType(reader);
			if ("fillcolor" === name) this.m_oFillColor = readColorType(reader);
			if ("shadowcolor" === name) this.m_oShadowColor = readColorType(reader);
			if ("strokecolor" === name) this.m_oStrokeColor = readColorType(reader);
		};
		CColorMenu.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
			writer.WriteXmlNullableAttributeString("extrusioncolor", getColorType(this.m_oExtrusionColor));
			writer.WriteXmlNullableAttributeString("fillcolor", getColorType(this.m_oFillColor));
			writer.WriteXmlNullableAttributeString("shadowcolor", getColorType(this.m_oShadowColor));
			writer.WriteXmlNullableAttributeString("strokecolor", getColorType(this.m_oStrokeColor));
		};

		function CColorMru() {
			CBaseNoId.call(this);
			this.m_arrColors = null;
			this.m_oExt = null;
		}

		IC(CColorMru, CBaseNoId, 0);
		CColorMru.prototype.readAttrXml = function (name, reader) {
			if (name === "ext") {
				this.m_oExt = readExt(reader);
			} else if (name === "colors") {
				let sColors = reader.GetValue();
				if ("" !== sColors) {
					let nStartPos = 0;
					let nEndPos = -1;
					let sColor;
					while (-1 !== (nEndPos = sColors.indexOf(",", nStartPos))) {
						sColor = sColors.substr(nStartPos, nEndPos - nStartPos);
						let oColor = new CColor();
						oColor.fromString(sColor);
						this.m_arrColors.push(oColor);
						nStartPos = nEndPos + 1;
					}
					nEndPos = sColors.length;
					sColor = sColors.substr(nStartPos, nEndPos - nStartPos);
					let oColor = new CColor();
					oColor.fromString(sColor);
					this.m_arrColors.push(oColor);
				}
			}
		};
		CColorMru.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
			let sResult = "";

			for (let nIndex = 0; nIndex < this.m_arrColors.length; nIndex++) {
				if (this.m_arrColors[nIndex]) {
					if (sResult.length > 0) {
						sResult += ",";
					}
					sResult += this.m_arrColors[nIndex].ToString();
				}
			}

			writer.WriteXmlNullableAttributeString("colors", sResult);
		};

		function CComplex() {
			CBaseNoId.call(this);
			this.m_oExt = null;
		}

		IC(CComplex, CBaseNoId, 0);
		CComplex.prototype.readAttrXml = function (name, reader) {
			if (name === "ext") {
				this.m_oExt = readExt(reader);
			}
		};
		CComplex.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
		};

		function CRelation() {
			CBaseNoId.call(this);
			this.m_oExt = null;
			this.m_sIdCntr = null;
			this.m_sIdDest = null;
			this.m_sIdSrc = null;
		}

		IC(CRelation, CBaseNoId, 0);
		CRelation.prototype.readAttrXml = function (name, reader) {
			if (name === "ext") {
				this.m_oExt = readExt(reader);
			} else if (name === "idcntr") this.m_sIdCntr = reader.GetValue();
			else if (name === "iddest") this.m_sIdDest = reader.GetValue();
			else if (name === "idsrc") this.m_sIdSrc = reader.GetValue();
		};
		CRelation.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
			writer.WriteXmlNullableAttributeString("idcntr", this.m_sIdCntr);
			writer.WriteXmlNullableAttributeString("iddest", this.m_sIdDest);
			writer.WriteXmlNullableAttributeString("idsrc", this.m_sIdSrc);
		};

		function CRelationTable() {
			CBaseNoId.call(this);
			this.m_oExt = null;
			this.m_arrRel = [];
		}

		IC(CRelationTable, CBaseNoId, 0);
		CRelationTable.prototype.readAttrXml = function (name, reader) {
			if (name === "ext") {
				this.m_oExt = readExt(reader);
			}
		};
		CRelationTable.prototype.readChildXml = function (name, reader) {
			if ("rel" === name) {
				let oRel = new CRelation();
				oRel.fromXml(reader);
				this.m_arrRel.push(oRel);
			}
		};
		CRelationTable.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
		};
		CRelationTable.prototype.writeChildren = function (writer) {
			for (let nIndex = 0; nIndex < this.m_arrRel.length; nIndex++) {
				if (this.m_arrRel[nIndex])
					this.m_arrRel[nIndex].toXml(writer, "o:rel");
			}
		};

		function CDiagram() {
			CBaseNoId.call(this);
			this.m_oAutoFormat = null;
			this.m_oAutoLayout = null;
			this.m_sConstrainbounds = null;
			this.m_oDmgBaseTextScale = null;
			this.m_oDmgFontSize = null;
			this.m_oDmgScaleX = null;
			this.m_oDmgScaleY = null;
			this.m_oDmgStyle = null;
			this.m_oExt = null;
			this.m_oReverse = null;

			this.m_oRelationTable = null;
		}

		IC(CDiagram, CBaseNoId, 0);
		CDiagram.prototype.readAttrXml = function (name, reader) {
			if ("autoformat" === name) {
				this.m_oAutoFormat = reader.GetValueBool();
				return;
			}
			if ("autolayout" === name) {
				this.m_oAutoLayout = reader.GetValueBool();
				return;
			}
			if ("constrainbounds" === name) {
				this.m_sConstrainbounds = reader.GetValue();
				return;
			}
			if ("dgmbasetextscale" === name) {
				this.m_oDmgBaseTextScale = reader.GetValueInt();
				return;
			}
			if ("dgmfontsize" === name) {
				this.m_oDmgFontSize = reader.GetValueInt();
				return;
			}
			if ("dgmscalex" === name) {
				this.m_oDmgScaleX = reader.GetValueInt();
				return;
			}
			if ("dgmscaley" === name) {
				this.m_oDmgScaleY = reader.GetValueInt();
				return;
			}
			if ("dgmstyle" === name) {
				this.m_oDmgStyle = reader.GetValueInt();
				return;
			}
			if ("ext" === name) {
				this.m_oExt = readExt(reader);
				return;
			}
			if ("reverse" === name) {
				this.m_oReverse = reader.GetValueBool();
			}
		};
		CDiagram.prototype.readChildXml = function (name, reader) {
			if ("relationtable" === name) {
				this.m_oRelationTable = new CRelationTable();
				this.m_oRelationTable.fromXml(reader);
			}
		};
		CDiagram.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt))
			writer.WriteXmlNullableAttributeUInt("dgmstyle", this.m_oDmgStyle)
			writer.WriteXmlNullableAttributeString("autoformat", getBooleanTrueFalse(this.m_oAutoFormat))
			writer.WriteXmlNullableAttributeString("reverse", getBooleanTrueFalse(this.m_oReverse))
			writer.WriteXmlNullableAttributeString("autolayout", getBooleanTrueFalse(this.m_oAutoLayout))
			writer.WriteXmlNullableAttributeUInt("dgmscalex", this.m_oDmgScaleX)
			writer.WriteXmlNullableAttributeUInt("dgmscaley", this.m_oDmgScaleY)
			writer.WriteXmlNullableAttributeUInt("dgmfontsize", this.m_oDmgFontSize)
			writer.WriteXmlNullableAttributeString("constrainbounds", this.m_sConstrainbounds)
			writer.WriteXmlNullableAttributeUInt("dgmbasetextscale", this.m_oDmgBaseTextScale)
		};
		CDiagram.prototype.writeChildren = function (writer) {
			if (this.m_oRelationTable) {
				this.m_oRelationTable.toXml(writer, "o:relationtable");
			}
		};

		function CEntry() {
			CBaseNoId.call(this);
			this.m_oNew = null;
			this.m_oOld = null;
		}

		IC(CEntry, CBaseNoId, 0);
		CEntry.prototype.readAttrXml = function (name, reader) {

			if ("new" === name) {
				this.m_oNew = reader.GetValue();
			}
			if ("old" === name) {
				this.m_oOld = reader.GetValue();
			}
		};
		CEntry.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("new", this.m_oNew);
			writer.WriteXmlNullableAttributeString("old", this.m_oOld);
		};

		function CEquationXml() {
			CBaseNoId.call(this);
			this.m_oContentType = null;
		}

		IC(CEquationXml, CBaseNoId, 0);
		CEquationXml.prototype.readAttrXml = function (name, reader) {
			if (name === "contentType") this.m_oContentType = readAlternateMathContentType(reader);
		};
		CEquationXml.prototype.readChildXml = function (name, reader) {
			//TODO
		};
		CEquationXml.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("contentType", getAlternateMathContentType(this.m_oContentType));
		};
		CEquationXml.prototype.writeChildren = function (writer) {
			//TODO:Implement in children
		};

		function CExtrusion() {
			CBaseNoId.call(this);
			this.m_oAutoRotationCenter = null;
			this.m_oBackDepth = null;
			this.m_oBrightness = null;
			this.m_oColor = null;
			this.m_oColorMode = null;
			this.m_oDiffusity = null;
			this.m_oEdge = null;
			this.m_oExt = null;
			this.m_oFacet = null;
			this.m_oForeDepth = null;
			this.m_oLightFace = null;
			this.m_oLightHarsh = null;
			this.m_oLightHarsh2 = null;
			this.m_oLightLevel = null;
			this.m_oLightLevel2 = null;
			this.m_oLightPosition = null;
			this.m_oLightPosition2 = null;
			this.m_oLockRotationCenter = null;
			this.m_oMetal = null;
			this.m_oOn = null;
			this.m_oOrientation = null;
			this.m_oOrientationAngle = null;
			this.m_oPlane = null;
			this.m_oRender = null;
			this.m_oRotationAngle = null;
			this.m_oRotationCenter = null;
			this.m_oShininess = null;
			this.m_oSkewAmt = null;
			this.m_oSkewAngle = null;
			this.m_oSpecularity = null;
			this.m_oType = null;
			this.m_oViewPoint = null;
			this.m_oViewPointOrigin = null;
		}

		IC(CExtrusion, CBaseNoId, 0);
		CExtrusion.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "autorotationcenter": {
					this.m_oAutoRotationCenter = reader.GetValueBool();
					break;
				}
				case "backdepth": {
					this.m_oBackDepth = reader.GetValue();
					break;
				}
				case "brightness": {
					this.m_oBrightness = readCVml_1_65536(reader);
					break;
				}
				case "color": {
					this.m_oColor = CVmlCommonElements.prototype.readColor.call(this, reader);
					break;
				}
				case "colormode": {
					this.m_oColorMode = readColorMode(reader);
					break;
				}
				case "diffusity": {
					this.m_oDiffusity = readCVml_1_65536(reader);
					break;
				}
				case "edge": {
					this.m_oEdge = reader.GetValue();
					break;
				}
				case "ext": {
					this.m_oExt = readExt(reader);
					break;
				}
				case "facet": {
					this.m_oFacet = reader.GetValueInt();
					break;
				}
				case "foredepth": {
					this.m_oForeDepth = reader.GetValue();
					break;
				}
				case "lightface": {
					this.m_oLightFace = reader.GetValueBool();
					break;
				}
				case "lightharsh": {
					this.m_oLightHarsh = reader.GetValueBool();
					break;
				}
				case "lightharsh2": {
					this.m_oLightHarsh2 = reader.GetValueBool();
					break;
				}
				case "lightlevel": {
					this.m_oLightLevel = readCVml_1_65536(reader);
					break;
				}
				case "lightlevel2": {
					this.m_oLightLevel2 = readCVml_1_65536(reader);
					break;
				}
				case "lightposition": {
					this.m_oLightPosition = new CVml_Vector3D_65536(reader.GetValue());
					break;
				}
				case "lightposition2": {
					this.m_oLightPosition2 = new CVml_Vector3D_65536(reader.GetValue());
					break;
				}
				case "lockrotationcenter": {
					this.m_oLockRotationCenter = reader.GetValueBool();
					break;
				}
				case "metal": {
					this.m_oMetal = reader.GetValueBool();
					break;
				}
				case "on": {
					this.m_oOn = reader.GetValueBool();
					break;
				}
				case "orientation": {
					this.m_oOrientation = new CVml_Vector3D(reader.GetValue());
					break;
				}
				case "orientationangle": {
					this.m_oOrientationAngle = reader.GetValueInt();
					break;
				}
				case "plane": {
					this.m_oPlane = reader.GetValue();
					break;
				}
				case "render": {
					this.m_oRender = reader.GetValue();
					break;
				}
				case "rotationangle": {
					this.m_oRotationAngle = new CVml_Vector2D(reader.GetValue());
					break;
				}
				case "rotationcenter": {
					this.m_oRotationCenter = new CVml_Vector3D(reader.GetValue());
					break;
				}
				case "shininess": {
					this.m_oShininess = reader.GetValueInt();
					break;
				}
				case "skewamt": {
					this.m_oSkewAmt = reader.GetValueInt();
					break;
				}
				case "skewangle": {
					this.m_oSkewAngle = reader.GetValueInt();
					break;
				}
				case "specularity": {
					this.m_oSpecularity = readCVml_1_65536(reader);
					break;
				}
				case "type": {
					this.m_oType = readExtrusionType(reader);
					break;
				}
				case "viewpoint": {
					this.m_oViewPoint = new CVml_Vector3D(reader.GetValue());
					break;
				}
				case "viewpointorigin": {
					this.m_oViewPointOrigin = new CVml_Vector2D_F(reader.GetValue());
					break;
				}
			}
		};
		CExtrusion.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
			if (false !== this.m_oOn)
				writer.WriteXmlNullableAttributeString("on", this.m_oOn);

			if (EExtrusionType.extrusiontypeParallel !== this.m_oType)
				writer.WriteXmlNullableAttributeString("type", getExtrusionType(this.m_oType));

			if ("solid" !== this.m_oRender)
				writer.WriteXmlNullableAttributeString("render", this.m_oRender);

			if (this.m_oViewPointOrigin && (0.5 !== this.m_oViewPointOrigin.GetX() || -0.5 !== this.m_oViewPointOrigin.GetY()))
				writer.WriteXmlNullableAttributeString("viewpointorigin", this.m_oRender.ToString());

			if (this.m_oViewPoint && (0 !== this.m_oViewPoint.GetX() || 0 !== this.m_oViewPoint.GetY() || 0 !== this.m_oViewPoint.GetZ()))
				writer.WriteXmlNullableAttributeString("viewpoint", this.m_oViewPoint.ToString());

			writer.WriteXmlNullableAttributeString("plane", this.m_oPlane);

			if (225 !== this.m_oSkewAngle)
				writer.WriteXmlNullableAttributeInt("skewangle", this.m_oSkewAngle);

			if (50 !== this.m_oSkewAmt)
				writer.WriteXmlNullableAttributeInt("skewamt", this.m_oSkewAmt);

			//if ( 0 !== this.m_oForeDepth )
			writer.WriteXmlNullableAttributeString("foredepth", this.m_oForeDepth);

			//if ( 36 !== this.m_oBackDepth )
			writer.WriteXmlNullableAttributeString("backdepth", this.m_oBackDepth);

			if (this.m_oOrientation && (100 !== this.m_oOrientation.GetX() || 0 !== this.m_oOrientation.GetY() || 0 !== this.m_oOrientation.GetZ()))
				writer.WriteXmlNullableAttributeString("orientation", this.m_oOrientation.ToString());

			if (0 !== this.m_oOrientationAngle)
				writer.WriteXmlNullableAttributeInt("orientationangle", this.m_oOrientationAngle);

			if (true !== this.m_oLockRotationCenter)
				writer.WriteXmlNullableAttributeString("lockrotationcenter", getBooleanTrueFalse(this.m_oLockRotationCenter));

			if (false !== this.m_oAutoRotationCenter)
				writer.WriteXmlNullableAttributeString("autorotationcenter", getBooleanTrueFalse(this.m_oAutoRotationCenter));

			if (this.m_oRotationCenter && (0 !== this.m_oRotationCenter.GetX() || 0 !== this.m_oRotationCenter.GetY() || 0 !== this.m_oRotationCenter.GetZ()))
				writer.WriteXmlNullableAttributeString("rotationcenter", this.m_oRotationCenter.ToString());

			if (this.m_oRotationAngle && (0 !== this.m_oRotationAngle.GetX() || 0 !== this.m_oRotationAngle.GetY()))
				writer.WriteXmlNullableAttributeString("rotationangle", this.m_oRotationAngle.ToString());

			if (EColorMode.colormodeAuto !== this.m_oColorMode)
				writer.WriteXmlNullableAttributeString("colormode", getColorMode(this.m_oColorMode));

			if (this.m_oColor !== null)
				writer.WriteXmlNullableAttributeString("color", getColorType(this.m_oColor));

			if (5 !== this.m_oShininess)
				writer.WriteXmlNullableAttributeInt("shininess", this.m_oShininess);

			if (0 !== this.m_oSpecularity)
				writer.WriteXmlNullableAttributeString("specularity", getCVml_1_65536(this.m_oSpecularity));

			if (1 !== this.m_oDiffusity)
				writer.WriteXmlNullableAttributeString("diffusity", getCVml_1_65536(this.m_oDiffusity));

			if (false !== this.m_oMetal)
				writer.WriteXmlNullableAttributeString("metal", getBooleanTrueFalse(this.m_oMetal));

			//if ( 1 !== this.m_oEdge )
			writer.WriteXmlNullableAttributeString("edge", this.m_oEdge);

			if (30000 !== this.m_oFacet)
				writer.WriteXmlNullableAttributeInt("facet", this.m_oFacet);

			if (true !== this.m_oLightFace)
				writer.WriteXmlNullableAttributeString("lightface", getBooleanTrueFalse(this.m_oLightFace));

			if (0.3 !== this.m_oBrightness)
				writer.WriteXmlNullableAttributeString("brightness", getCVml_1_65536(this.m_oBrightness));

			if (this.m_oLightPosition && (50000 !== this.m_oLightPosition.GetX() || 0 !== this.m_oLightPosition.GetY() || 10000 !== this.m_oLightPosition.GetZ()))
				writer.WriteXmlNullableAttributeString("lightposition", this.m_oLightPosition.ToString());

			if (0.6 !== this.m_oLightLevel)
				writer.WriteXmlNullableAttributeString("lightlevel", getCVml_1_65536(this.m_oLightLevel));

			if (true !== this.m_oLightHarsh)
				writer.WriteXmlNullableAttributeString("lightharsh", getBooleanTrueFalse(this.m_oLightHarsh));

			if (this.m_oLightPosition2 && (50000 !== this.m_oLightPosition2.GetX() || 0 !== this.m_oLightPosition2.GetY() || 10000 !== this.m_oLightPosition2.GetZ()))
				writer.WriteXmlNullableAttributeString("lightposition2", this.m_oLightPosition2.ToString());

			if (0.6 !== this.m_oLightLevel2)
				writer.WriteXmlNullableAttributeString("lightlevel2", getCVml_1_65536(this.m_oLightLevel2));

			if (false !== this.m_oLightHarsh2)
				writer.WriteXmlNullableAttributeString("lightharsh2", getBooleanTrueFalse(this.m_oLightHarsh2));
		};

		function CFieldCodes() {
			CBaseNoId.call(this);
			this.m_sText = null;
		}

		IC(CFieldCodes, CBaseNoId, 0);
		CFieldCodes.prototype.fromXml = function (reader) {
			this.m_sText = reader.GetTextDecodeXml();
		};
		CFieldCodes.prototype.toXml = function (writer) {
			writer.WriteXmlNullableValueStringEncode("o:FieldCodes", this.m_sText);
		};

		function CFill() {
			CBaseNoId.call(this);
			this.m_oExt = null;
			this.m_oType = null;
		}

		IC(CFill, CBaseNoId, 0);
		CFill.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "vext":
					this.m_oExt = readExt(reader);
					break;
				case "type":
					this.m_oType = readFillType(reader);
					break;
			}
		};
		CFill.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
			writer.WriteXmlNullableAttributeString("type", getFillType(this.m_oType));
		};


		function CIdMap() {
			CBaseNoId.call(this);
			this.m_sData = null;
			this.m_oExt = null;
		}

		IC(CIdMap, CBaseNoId, 0);
		CIdMap.prototype.readAttrXml = function (name, reader) {
			if (name === "ext") this.m_oExt = readExt(reader);
			if (name === "data") this.m_sData = reader.GetValue()
		};
		CIdMap.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
			writer.WriteXmlNullableAttributeString("data", this.m_sData);
		};

		function CInk() {
			CBaseNoId.call(this);

			this.m_oAnnotation = null;
			this.m_oContentType = null;
			this.m_sI = null;
		}

		IC(CInk, CBaseNoId, 0);
		CInk.prototype.readAttrXml = function (name, reader) {

			if (name === "annotation") this.m_oAnnotation = reader.GetValueBool();
			if (name === "contentType") this.m_oContentType = readContentType(reader);
			if (name === "i") this.m_sI = reader.GetValue();
		};
		CInk.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("i", this.m_sI);
			writer.WriteXmlNullableAttributeString("annotation", getBooleanTrueFalse(this.m_oAnnotation));
			writer.WriteXmlNullableAttributeString("contentType", getContentType(this.m_oContentType));
		};

		let EOLELinkType =
			{
				olelinktypeBitmap: 0,
				olelinktypeEmf: 1,
				olelinktypeJpeg: 2,
				olelinktypePicture: 3,
				olelinktypePng: 4
			};

		function readOLELinkType(reader) {
			let sValue = reader.GetValue();
			if ("Bitmap" === sValue) return EOLELinkType.olelinktypeBitmap;
			if ("EnhancedMetaFile" === sValue) return EOLELinkType.olelinktypeEmf;
			if ("Jpeg" === sValue) return EOLELinkType.olelinktypeJpeg;
			if ("Picture" === sValue) return EOLELinkType.olelinktypePicture;
			else if ("Png" === sValue) return EOLELinkType.olelinktypePng;
			return EOLELinkType.olelinktypeBitmap;
		}


		function getOLELinkType(nType) {
			if (EOLELinkType.olelinktypeBitmap === nType) return ("Bitmap");
			if (EOLELinkType.olelinktypeEmf === nType) return ("EnhancedMetaFile");
			if (EOLELinkType.olelinktypeJpeg === nType) return ("Jpeg");
			if (EOLELinkType.olelinktypePicture === nType) return ("Picture");
			if (EOLELinkType.olelinktypePng === nType) return ("Png");
			return null;
		}

		function CLinkType() {
			CBaseNoId.call(this);
			this.m_oValue = null;
		}

		IC(CLinkType, CBaseNoId, 0);
		CLinkType.prototype.fromXml = function (reader) {
			this.m_oValue = readOLELinkType(reader.GetTextDecodeXml());
		};
		CLinkType.prototype.toXml = function (writer) {
			writer.WriteXmlNullableValueStringEncode("o:LinkType", getOLELinkType(this.m_oValue));
		};


		function CLock() {
			CBaseNoId.call(this);
			this.m_oAdjustHandles = null;
			this.m_oAspectRatio = null;
			this.m_oCropping = null;
			this.m_oExt = null;
			this.m_oGrouping = null;
			this.m_oPosition = null;
			this.m_oRotation = null;
			this.m_oSelection = null;
			this.m_oShapeType = null;
			this.m_oText = null;
			this.m_oUnGrouping = null;
			this.m_oVerticies = null;
		}

		IC(CLock, CBaseNoId, 0);
		CLock.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "adjusthandles":
					this.m_oAdjustHandles = reader.GetValueBool();
					break;
				case "aspectratio":
					this.m_oAspectRatio = reader.GetValueBool();
					break;
				case "cropping":
					this.m_oCropping = reader.GetValueBool();
					break;
				case "ext":
					this.m_oExt = readExt(reader);
					break;
				case "grouping":
					this.m_oGrouping = reader.GetValueBool();
					break;
				case "position":
					this.m_oPosition = reader.GetValueBool();
					break;
				case "rotation":
					this.m_oRotation = reader.GetValueBool();
					break;
				case "selection":
					this.m_oSelection = reader.GetValueBool();
					break;
				case "shapetype":
					this.m_oShapeType = reader.GetValueBool();
					break;
				case "text":
					this.m_oText = reader.GetValueBool();
					break;
				case "ungrouping":
					this.m_oUnGrouping = reader.GetValueBool();
					break;
				case "verticies":
					this.m_oVerticies = reader.GetValueBool();
					break;
			}
		};
		CLock.prototype.writeAttrXmlImpl = function (writer) {

			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));

			if (false !== this.m_oPosition)
				writer.WriteXmlNullableAttributeString("position", getBooleanTrueFalse(this.m_oPosition));

			if (false !== this.m_oSelection)
				writer.WriteXmlNullableAttributeString("selection", getBooleanTrueFalse(this.m_oSelection));

			if (false !== this.m_oGrouping)
				writer.WriteXmlNullableAttributeString("grouping", getBooleanTrueFalse(this.m_oGrouping));

			if (false !== this.m_oUnGrouping)
				writer.WriteXmlNullableAttributeString("ungrouping", getBooleanTrueFalse(this.m_oUnGrouping));

			if (false !== this.m_oRotation)
				writer.WriteXmlNullableAttributeString("rotation", getBooleanTrueFalse(this.m_oRotation));

			if (false !== this.m_oCropping)
				writer.WriteXmlNullableAttributeString("cropping", getBooleanTrueFalse(this.m_oCropping));

			if (false !== this.m_oVerticies)
				writer.WriteXmlNullableAttributeString("verticies", getBooleanTrueFalse(this.m_oVerticies));

			if (false !== this.m_oAdjustHandles)
				writer.WriteXmlNullableAttributeString("adjusthandles", getBooleanTrueFalse(this.m_oAdjustHandles));

			if (false !== this.m_oText)
				writer.WriteXmlNullableAttributeString("text", getBooleanTrueFalse(this.m_oText));

			if (false !== this.m_oAspectRatio)
				writer.WriteXmlNullableAttributeString("aspectratio", getBooleanTrueFalse(this.m_oAspectRatio));

			if (false !== this.m_oShapeType)
				writer.WriteXmlNullableAttributeString("shapetype", getBooleanTrueFalse(this.m_oShapeType));

		};

		function CLockedField() {
			CBaseNoId.call(this);
			this.m_oValue = null;
		}

		IC(CLockedField, CBaseNoId, 0);
		CLockedField.prototype.fromXml = function (reader) {
			this.m_oValue = reader.GetBool(reader.GetTextDecodeXml());
		};
		CLockedField.prototype.toXml = function (writer) {
			writer.WriteXmlNullableValueStringEncode("o:LockedField", getBooleanTrueFalse(this.m_oValue));
		};


		let EOLEDrawAspect =
			{
				oledrawaspectContent: 0,
				oledrawaspectIcon: 1
			};

		function readOLEDrawAspect(reader) {
			let sVal = reader.GetValue();
			if (sVal === "Content") return EOLEDrawAspect.oledrawaspectContent;
			if (sVal === "Icon") return EOLEDrawAspect.oledrawaspectIcon;
			return EOLEDrawAspect.oledrawaspectContent;
		}

		function getOLEDrawAspect(nType) {
			if (nType === EOLEDrawAspect.oledrawaspectContent) return "Content";
			if (nType === EOLEDrawAspect.oledrawaspectIcon) return "Icon";
			return null;
		}


		let EOLEType =
			{
				oletypeEmbed: 0,
				oletypeLink: 1
			};

		function readOLEType(reader) {
			let sVal = reader.GetValue();
			if (sVal === "embed") return EOLEType.oletypeEmbed;
			if (sVal === "link") return EOLEType.oletypeLink;
			return EOLEType.oletypeEmbed;
		}

		function getOLEType(nType) {
			if (nType === EOLEType.oletypeEmbed) return "embed";
			if (nType === EOLEType.oletypeLink) return "link";
			return null;
		}


		let EOLEUpdateMode =
			{
				oleupdatemodeAlways: 0,
				oleupdatemodeOnCall: 1
			};

		function readOLEUpdateMode(reader) {
			let sVal = reader.GetValue();
			if (sVal === "Always") return EOLEUpdateMode.oleupdatemodeAlways;
			if (sVal === "OnCall") return EOLEUpdateMode.oleupdatemodeOnCall;
			return EOLEUpdateMode.oleupdatemodeAlways;
		}

		function getOLEUpdateMode(nType) {
			if (nType === EOLEUpdateMode.oleupdatemodeAlways) return "Always";
			if (nType === EOLEUpdateMode.oleupdatemodeOnCall) return "OnCall";
			return null;
		}

		function COLEObject() {
			CBaseNoId.call(this);
		}

		IC(COLEObject, CBaseNoId, 0);
		COLEObject.prototype.readAttrXml = function (name, reader) {

			if (name === "DrawAspect") this.m_oDrawAspect = readOLEDrawAspect(reader.GetValue());
			if (name === "id") this.m_oId = reader.GetValue();
			if (name === "ObjectID") this.m_sObjectId = reader.GetValue();
			if (name === "ProgID") this.m_sProgId = reader.GetValue();
			if (name === "ShapeID") this.m_sShapeId = reader.GetValue();
			if (name === "Type") this.m_oType = readOLEType(reader);
			if (name === "UpdateMode") this.m_oUpdateMode = readOLEUpdateMode(reader);

		};
		COLEObject.prototype.readChildXml = function (name, reader) {
			if ("FieldCodes" === name) {
				this.m_oFieldCodes = new CFieldCodes();
				this.m_oFieldCodes.fromXml(reader);
			} else if ("LinkType" === name) {

				this.m_oLinkType = new CLinkType();
				this.m_oLinkType.fromXml(reader);
			} else if ("LockedField" === name) {
				this.m_oLockedField = new CLockedField();
				this.m_oLockedField.fromXml(reader);
			}
		};
		COLEObject.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("Type", getOLEType(this.m_oType));
			writer.WriteXmlNullableAttributeString("ProgID", this.m_sProgId);
			writer.WriteXmlNullableAttributeString("ShapeID", this.m_sShapeId);
			writer.WriteXmlNullableAttributeString("DrawAspect", getOLEDrawAspect(this.m_oDrawAspect));
			writer.WriteXmlNullableAttributeString("ObjectID", this.m_sObjectId);
			writer.WriteXmlNullableAttributeString("r:id", this.m_oId);
			writer.WriteXmlNullableAttributeString("UpdateMode", getOLEUpdateMode(this.m_oUpdateMode));
		};
		COLEObject.prototype.writeChildren = function (writer) {

			if (this.m_oLinkType)
				this.m_oLinkType.toXml(writer, "o:LinkType");

			if (this.m_oLockedField)
				this.m_oLockedField.toXml(writer, "o:LockedField");

			if (this.m_oFieldCodes)
				this.m_oFieldCodes.toXml(writer, "o:FieldCodes");
		};

		function CProxy() {
			CBaseNoId.call(this);
			this.m_oConnectLoc = null;
			this.m_oEnd = null;
			this.m_sIdRef = null;
			this.m_oStart = null;
		}

		IC(CProxy, CBaseNoId, 0);
		CProxy.prototype.readAttrXml = function (name, reader) {
			if (name === "connectloc") this.m_oConnectLoc = reader.GetValueInt();
			if (name === "end") this.m_oEnd = reader.GetValueBool();
			if (name === "idref") this.m_sIdRef = reader.GetValue();
			if (name === "start") this.m_oStart = reader.GetValueBool();
		};
		CProxy.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("start", getBooleanTrueFalse(this.m_oStart));
			writer.WriteXmlNullableAttributeString("end", getBooleanTrueFalse(this.m_oEnd));
			writer.WriteXmlNullableAttributeString("idref", this.m_sIdRef);
			writer.WriteXmlNullableAttributeInt("connectloc", this.m_oConnectLoc);
		};

		let EHow =
			{
				howBottom: 0,
				howCenter: 1,
				howLeft: 2,
				howMiddle: 3,
				howRight: 4,
				howTop: 5
			};


		function readHow(reader) {
			let sVal = reader.GetValue();
			if (sVal === "bottom") return EHow.howBottom;
			if (sVal === "center") return EHow.howCenter;
			if (sVal === "left") return EHow.howLeft;
			if (sVal === "middle") return EHow.howMiddle;
			if (sVal === "right") return EHow.howRight;
			if (sVal === "top") return EHow.howTop;
			return EHow.howTop;
		}

		function getHow(nType) {
			if (nType === EHow.howBottom) return "bottom";
			if (nType === EHow.howCenter) return "center";
			if (nType === EHow.howLeft) return "left";
			if (nType === EHow.howMiddle) return "middle";
			if (nType === EHow.howRight) return "right";
			if (nType === EHow.howTop) return "top";
			return null;
		}

		let ERType =
			{
				rtypeAlign: 0,
				rtypeArc: 1,
				rtypeCallout: 2,
				rtypeConnector: 3
			};

		function readRType(reader) {
			let sVal = reader.GetValue();
			if (sVal === "align") return ERType.rtypeAlign;
			if (sVal === "arc") return ERType.rtypeArc;
			if (sVal === "callout") return ERType.rtypeCallout;
			if (sVal === "connector") return ERType.rtypeConnector;
			return ERType.rtypeAlign;
		}

		function getRType(nType) {
			if (nType === ERType.rtypeAlign) return "align";
			if (nType === ERType.rtypeAlign) return "arc";
			if (nType === ERType.rtypeAlign) return "callout";
			if (nType === ERType.rtypeAlign) return "connector";
			return null;
		}

		function CR() {
			CBaseNoId.call(this);
			this.m_oHow = null;
			this.m_sId = null;
			this.m_sIdRef = null;
			this.m_oType = null;

			this.m_arrProxy = [];
		}

		IC(CR, CBaseNoId, 0);
		CR.prototype.readAttrXml = function (name, reader) {
			if (name === "how") this.m_oHow = readHow(reader);
			if (name === "id") this.m_sId = reader.GetValue();
			if (name === "idref") this.m_sIdRef = reader.GetValue();
			if (name === "type") this.m_oType = readRType(reader);
		};
		CR.prototype.readChildXml = function (name, reader) {
			if ("proxy" === name) {
				let oProxy = new CProxy();
				oProxy.fromXml(reader);
				this.m_arrProxy.push(oProxy);
			}
		};
		CR.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("type", getRType(this.m_oType))
			writer.WriteXmlNullableAttributeString("how", getHow(this.m_oHow))
			writer.WriteXmlNullableAttributeString("idref", this.m_sIdRef)

		};
		CR.prototype.writeChildren = function (writer) {
			for (let nIndex = 0; nIndex < this.m_arrProxy.length; nIndex++) {
				if (this.m_arrProxy[nIndex])
					this.m_arrProxy[nIndex].toXml(writer, "o:proxy");
			}
		};

		function CRegroupTable() {
			CBaseNoId.call(this);
			this.m_oExt = null;
			this.m_arrEntry = [];
		}

		IC(CRegroupTable, CBaseNoId, 0);
		CRegroupTable.prototype.readAttrXml = function (name, reader) {
			if (name === "ext") {
				this.m_oExt = readExt(reader);
			}
		};
		CRegroupTable.prototype.readChildXml = function (name, reader) {
			if ("entry" === name) {
				let oEntry = new CEntry();
				oEntry.fromXml(reader);
				this.m_arrEntry.push(oEntry);
			}
		};
		CRegroupTable.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
		};
		CRegroupTable.prototype.writeChildren = function (writer) {
			for (let nIndex = 0; nIndex < this.m_arrEntry.length; nIndex++) {
				if (this.m_arrEntry[nIndex])
					this.m_arrEntry[nIndex].toXml(writer, "o:entry");
			}
		};

		function CRules() {
			CBaseNoId.call(this);
			this.m_oExt = null;
			this.m_arrR = [];
		}

		IC(CRules, CBaseNoId, 0);
		CRules.prototype.readAttrXml = function (name, reader) {
			if (name === "ext") {
				this.m_oExt = readExt(reader);
			}
		};
		CRules.prototype.readChildXml = function (name, reader) {
			if ("entry" === name) {
				let oPr = new CR();
				oPr.fromXml(reader);
				this.m_arrR.push(oPr);
			}
		};
		CRules.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
		};
		CRules.prototype.writeChildren = function (writer) {

			for (let nIndex = 0; nIndex < this.m_arrR.length; nIndex++) {
				if (this.m_arrR[nIndex])
					this.m_arrR[nIndex].toXml(writer, "o:r");
			}

		};

		function CShapeLayout() {
			CBaseNoId.call(this);


			this.m_oExt = null;

			// Childs
			this.m_oIdMap = null;
			this.m_oRegroupTable = null;
			this.m_oRules = null;
		}

		IC(CShapeLayout, CBaseNoId, 0);
		CShapeLayout.prototype.readAttrXml = function (name, reader) {
			if (name === "ext") {
				this.m_oExt = readExt(reader);
			}
		};
		CShapeLayout.prototype.readChildXml = function (name, reader) {
			if ("idmap" === name) {

				this.m_oIdMap = new CIdMap();
				this.m_oIdMap.fromXml(reader);
			} else if ("regrouptable" === name) {

				this.m_oRegroupTable = new CRegroupTable();
				this.m_oRegroupTable.fromXml(reader);
			} else if ("rules" === name) {

				this.m_oRules = new CRules();
				this.m_oRules.fromXml(reader);
			}
		};
		CShapeLayout.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
		};
		CShapeLayout.prototype.writeChildren = function (writer) {
			if (this.m_oIdMap !== null)
				this.m_oIdMap.toXml(writer, "o:idmap");

			if (this.m_oRegroupTable !== null)
				this.m_oRegroupTable.toXml(writer, "o:regrouptable");

			if (this.m_oRules !== null)
				this.m_oRules.toXml(writer, "o:rules");
		};

		function CSignatureLine() {
			CBaseNoId.call(this);
			this.m_sAddXml = null;
			this.m_oAllowComments = null;
			this.m_oExt = null;
			this.m_oId = null;
			this.m_oIsSignatureLine = null;
			this.m_oProvId = null;
			this.m_oShowSignDate = null;
			this.m_sSigningInstructions = null;
			this.m_oSigningInstructionsSet = null;
			this.m_sSigProvUrl = null;
			this.m_sSuggestedSigner = null;
			this.m_sSuggestedSigner2 = null;
			this.m_sSuggestedSignerEmail = null;
		}

		IC(CSignatureLine, CBaseNoId, 0);
		CSignatureLine.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "addlxml":
					this.m_sAddXml = reader.GetValue();
					break;
				case "allowcomments":
					this.m_oAllowComments = reader.GetValueBool();
					break;
				case "ext":
					this.m_oExt = readExt(reader);
					break;
				case "id":
					this.m_oId = reader.GetValue();
					break;
				case "issignatureline":
					this.m_oIsSignatureLine = reader.GetValueBool();
					break;
				case "provid":
					this.m_oProvId = reader.GetValue();
					break;
				case "showsigndate":
					this.m_oShowSignDate = reader.GetValueBool();
					break;
				case "signinginstructions":
					this.m_sSigningInstructions = reader.GetValue();
					break;
				case "signinginstructionsset":
					this.m_oSigningInstructionsSet = reader.GetValueBool();
					break;
				case "sigprovurl":
					this.m_sSigProvUrl = reader.GetValue();
					break;
				case "suggestedsigner":
					this.m_sSuggestedSigner = reader.GetValue();
					break;
				case "suggestedsigner2":
					this.m_sSuggestedSigner2 = reader.GetValue();
					break;
				case "suggestedsigneremail":
					this.m_sSuggestedSignerEmail = reader.GetValue();
					break;
			}
		};
		CSignatureLine.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));

			if (this.m_oIsSignatureLine !== null)
				writer.WriteXmlNullableAttributeString("issignatureline", getBooleanTrueFalse(this.m_oIsSignatureLine));

			writer.WriteXmlNullableAttributeString("id", this.m_oId);
			writer.WriteXmlNullableAttributeString("provid", this.m_oProvId);

			if (this.m_oSigningInstructionsSet !== null)
				writer.WriteXmlNullableAttributeString("signinginstructionsset", getBooleanTrueFalse(this.m_oSigningInstructionsSet));

			if (this.m_oAllowComments !== null)
				writer.WriteXmlNullableAttributeString("allowcomments", getBooleanTrueFalse(this.m_oAllowComments));

			if (this.m_oShowSignDate !== null)
				writer.WriteXmlNullableAttributeString("showsigndate", getBooleanTrueFalse(this.m_oShowSignDate));

			writer.WriteXmlNullableAttributeString("o:suggestedsigner", this.m_sSuggestedSigner);
			writer.WriteXmlNullableAttributeString("o:suggestedsigner2", this.m_sSuggestedSigner2);
			writer.WriteXmlNullableAttributeString("o:suggestedsigneremail", this.m_sSuggestedSignerEmail);
			writer.WriteXmlNullableAttributeString("o:signinginstructions", this.m_sSigningInstructions);
			writer.WriteXmlNullableAttributeString("o:addlxml", this.m_sAddXml);
			writer.WriteXmlNullableAttributeString("o:sigprovurl", this.m_sSigProvUrl);
		};

		function CSkew() {
			CBaseNoId.call(this);
			this.m_oExt = null;
			this.m_sId = null;
			this.m_sMatrix = null;
			this.m_sOffset = null;
			this.m_oOn = null;
			this.m_sOrigin = null;
		}

		IC(CSkew, CBaseNoId, 0);
		CSkew.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "ext":
					this.m_oExt = readExt(reader.GetValue());
					break;
				case "id":
					this.m_sId = reader.GetValue();
					break;
				case "matrix":
					this.m_sMatrix = reader.GetValue();
					break;
				case "offset":
					this.m_sOffset = reader.GetValue();
					break;
				case "on":
					this.m_oOn = reader.GetValueBool();
					break;
				case "origin":
					this.m_sOrigin = reader.GetValue();
					break;
			}
		};
		CSkew.prototype.writeAttrXmlImpl = function (writer) {

			writer.WriteXmlNullableAttributeString("v:ext", getExt(this.m_oExt));
			writer.WriteXmlNullableAttributeString("id", this.m_sId);
			writer.WriteXmlNullableAttributeString("on", getBooleanTrueFalse(this.m_oOn));
			writer.WriteXmlNullableAttributeString("offset", this.m_sOffset);
			writer.WriteXmlNullableAttributeString("origin", this.m_sOrigin);
			writer.WriteXmlNullableAttributeString("matrix", this.m_sMatrix);
		};


		//Vml


		function getBWMode(nType) {

			switch (nType) {
				case EBWMode.bwmodeAuto : {
					return "auto";
				}
				case EBWMode.bwmodeBlack : {
					return "black";
				}
				case EBWMode.bwmodeBlackTextAndLines : {
					return "blackTextAndLines";
				}
				case EBWMode.bwmodeColor : {
					return "color";
				}
				case EBWMode.bwmodeGrayOutline : {
					return "grayOutline";
				}
				case EBWMode.bwmodeGrayScale : {
					return "grayScale";
				}
				case EBWMode.bwmodeHide : {
					return "hide";
				}
				case EBWMode.bwmodeHighContrast : {
					return "highContrast";
				}
				case EBWMode.bwmodeInverseGray : {
					return "inverseGray";
				}
				case EBWMode.bwmodeLightGrayscale : {
					return "lightGrayscale";
				}
				case EBWMode.bwmodeUndrawn : {
					return "undrawn";
				}
				case EBWMode.bwmodeWhite : {
					return "white";
				}
			}
			return null;
		}

		function CVmlCommonElements() {
			CBaseNoId.call(this);
			// 1 AG_AllCoreAttributes
			// 1.1 AG_CoreAttributes
			this.m_sId = null;
			this.m_oStyle = null;
			this.m_sHref = null;
			this.m_sTarget = null;
			this.m_sClass = null;
			this.m_sTitle = null;
			this.m_sAlt = null;
			this.m_oCoordSize = null;
			this.m_oCoordOrigin = null;
			this.m_oWrapCoords = null;
			this.m_oPrlet = null;
			// 1.2 AG_OfficeCoreAttributes
			this.m_sSpId = null;
			this.m_oOned = null;
			this.m_oRegroupId = null;
			this.m_oDoubleClickNotify = null;
			this.m_oButton = null;
			this.m_oUserHidden = null;
			this.m_oBullet = null;
			this.m_oHr = null;
			this.m_oHrStd = null;
			this.m_oHrNoShade = null;
			this.m_oHrPct = null;
			this.m_oHrAlign = null;
			this.m_oAllowInCell = null;
			this.m_oAllowOverlap = null;
			this.m_oUserDrawn = null;
			this.m_oBorderTopColor = null;
			this.m_oBorderLeftColor = null;
			this.m_oBorderBottomColor = null;
			this.m_oBorderRightColor = null;
			this.m_oDgmLayout = null;
			this.m_oDgmNodeKind = null;
			this.m_oDgmLayoutMru = null;
			this.m_oInsetMode = null;
			// 2 AG_AllShapeAttributes
			// 2.1 AG_ShapeAttributes
			this.m_oChromaKey = null;
			this.m_oFilled = null;
			this.m_oFillColor = null;
			this.m_oOpacity = null;
			this.m_oStroked = null;
			this.m_oStrokeColor = null;
			this.m_oStrokeWeight = null;
			this.m_oInsetPen = null;
			// 2.2 AG_OfficeShapeAttributes
			this.m_oSpt = null;
			this.m_oConnectorType = null;
			this.m_oBwMode = null;
			this.m_oBwPure = null;
			this.m_oBwNormal = null;
			this.m_oForceDash = null;
			this.m_oOleIcon = null;
			this.m_oOle = null;
			this.m_oPreferRelative = null;
			this.m_oClipToWrap = null;
			this.m_oClip = null;


			this.items = [];

			this.m_bComment = null;

		}

		IC(CVmlCommonElements, CBaseNoId, 0);
		CVmlCommonElements.prototype.readColor = function (reader) {
			return readColorType(reader);
		};
		CVmlCommonElements.prototype.readVector2D = function (reader) {
			return new CVmlVector2D(reader.GetValue());
		};
		CVmlCommonElements.prototype.readPolygon2D = function (reader) {
			return new CVmlPolygon2D(reader.GetValue());
		};
		CVmlCommonElements.prototype.readBWMode = function (reader) {
			let sVal = reader.GetValue();
			switch (sVal) {
				case "auto": {
					return EBWMode.bwmodeAuto;
				}
				case "black": {
					return EBWMode.bwmodeBlack;
				}
				case "blackTextAndLines": {
					return EBWMode.bwmodeBlackTextAndLines;
				}
				case "color": {
					return EBWMode.bwmodeColor;
				}
				case "grayOutline": {
					return EBWMode.bwmodeGrayOutline;
				}
				case "grayScale": {
					return EBWMode.bwmodeGrayScale;
				}
				case "hide": {
					return EBWMode.bwmodeHide;
				}
				case "highContrast": {
					return EBWMode.bwmodeHighContrast;
				}
				case "inverseGray": {
					return EBWMode.bwmodeInverseGray;
				}
				case "lightGrayscale": {
					return EBWMode.bwmodeLightGrayscale;
				}
				case "undrawn": {
					return EBWMode.bwmodeUndrawn;
				}
				case "white": {
					return EBWMode.bwmodeWhite;
				}
			}
			return null;
		};
		CVmlCommonElements.prototype.readAttrXml = function (name, reader) {
			if ("alt" === name) this.m_sAlt = reader.GetValue();
			else if ("chromakey" === name) this.m_oChromaKey = this.readColor(reader);
			else if ("class" === name) this.m_sClass = reader.GetValue();
			else if ("coordorigin" === name) this.m_oCoordOrigin = this.readVector2D(reader);
			else if ("coordsize" === name) this.m_oCoordSize = this.readVector2D(reader);
			else if ("fillcolor" === name) this.m_oFillColor = this.readColor(reader);
			else if ("filled" === name) this.m_oFilled = reader.GetValueBool();
			else if ("href" === name) this.m_sHref = reader.GetValue();
			else if ("id" === name) this.m_sId = reader.GetValue();
			else if ("insetpen" === name) this.m_oInsetPen = reader.GetValueBool();
			else if ("allowincell" === name) this.m_oAllowInCell = reader.GetValueBool();
			else if ("allowoverlap" === name) this.m_oAllowOverlap = reader.GetValueBool();
			else if ("opacity" === name) {
				this.m_oOpacity = readCVml_1_65536(reader);
			} else if ("borderbottomcolor" === name) this.m_oBorderBottomColor = this.readColor(reader);
			else if ("borderleftcolor" === name) this.m_oBorderLeftColor = this.readColor(reader);
			else if ("borderrightcolor" === name) this.m_oBorderRightColor = this.readColor(reader);
			else if ("bordertopcolor" === name) this.m_oBorderTopColor = this.readColor(reader);
			else if ("bullet" === name) this.m_oBullet = reader.GetValueBool();
			else if ("button" === name) this.m_oButton = reader.GetValueBool();
			else if ("bwmode" === name) this.m_oBwMode = this.readBWMode(reader);
			else if ("bwnormal" === name) this.m_oBwNormal = this.readBWMode(reader);
			else if ("bwpure" === name) this.m_oBwPure = this.readBWMode(reader);
			else if ("clip" === name) this.m_oClip = reader.GetValueBool();
			else if ("cliptowrap" === name) this.m_oClipToWrap = reader.GetValueBool();
			else if ("connectortype" === name) this.m_oConnectorType = reader.GetValue();
			else if ("doubleclicknotify" === name) this.m_oDoubleClickNotify = reader.GetValueBool();
			else if ("dgmlayout" === name) this.m_oDgmLayout = reader.GetValueInt();
			else if ("dgmlayoutmru" === name) this.m_oDgmLayoutMru = reader.GetValueInt();
			else if ("dgmnodekind" === name) this.m_oDgmNodeKind = reader.GetValue();
			else if ("forcedash" === name) this.m_oForceDash = reader.GetValueBool();
			else if ("hr" === name) this.m_oHr = reader.GetValueBool();
			else if ("hralign" === name) {
				let sVal = reader.GetValue();
				switch (sVal) {
					case "center": {
						this.m_oHrAlign = AscCommon.align_Center;
						break;
					}
					case "left": {
						this.m_oHrAlign = AscCommon.align_Left;
						break;
					}
					case "right": {
						this.m_oHrAlign = AscCommon.align_Right;
						break;
					}
				}
			} else if ("hrnoshade" === name) this.m_oHrNoShade = reader.GetValueBool();
			else if ("hrpct" === name) this.m_oHrPct = reader.GetValueDouble();
			else if ("hrstd" === name) this.m_oHrStd = reader.GetValueBool();
			else if ("insetmode" === name) {
				this.m_oInsetMode = readInsetMode(reader);
			} else if ("ole" === name) this.m_oOle = reader.GetValueBool();
			else if ("oleicon" === name) this.m_oOleIcon = reader.GetValueBool();
			else if ("oned" === name) this.m_oOned = reader.GetValueBool();
			else if ("preferrelative" === name) this.m_oPreferRelative = reader.GetValueBool();
			else if ("regroupid" === name) this.m_oRegroupId = reader.GetValueInt();
			else if ("spid" === name) this.m_sSpId = reader.GetValue();
			else if ("spt" === name) this.m_oSpt = reader.GetValueInt();
			else if ("userdrawn" === name) this.m_oUserDrawn = reader.GetValueBool();
			else if ("userhidden" === name) this.m_oUserHidden = reader.GetValueBool();
			else if ("print" === name) this.m_oPrlet = reader.GetValueBool();
			else if ("strokecolor" === name) this.m_oStrokeColor = this.readColor(reader);
			else if ("stroked" === name) this.m_oStroked = reader.GetValueBool();
			else if ("strokeweight" === name) this.m_oStrokeWeight = reader.GetValueInt();
			else if ("style" === name) this.m_oStyle = new CCssStyle(reader.GetValue());
			else if ("target" === name) this.m_sTarget = reader.GetValue();
			else if ("title" === name) this.m_sTitle = reader.GetValue();
			else if ("wrapcoords" === name) this.m_oWrapCoords = this.readPolygon2D(reader);
		};
		CVmlCommonElements.prototype.readChildXml = function (name, reader) {
			let oItem = null;
			let client_data;
			if ("callout" === name)
				oItem = new CCallout();
			else if ("clippath" === name)
				oItem = new CClipPath();
			else if ("extrusion" === name)
				oItem = new CExtrusion();
			else if ("lock" === name)
				oItem = new CLock();
			else if ("signatureline" === name)
				oItem = new CSignatureLine();
			else if ("skew" === name)
				oItem = new CSkew();
			else if ("fill" === name)
				oItem = new CFillVml();
			else if ("formulas" === name)
				oItem = new CFormulas();
			else if ("handles" === name)
				oItem = new CHandles();
			else if ("imagedata" === name)
				oItem = new CImageData();
			else if ("path" === name)
				oItem = new CPath();
			else if ("shadow" === name)
				oItem = new CShadow();
			else if ("stroke" === name)
				oItem = new CStroke();
			else if ("textbox" === name)
				oItem = new CTextbox();
			else if ("textpath" === name)
				oItem = new CTextPath();
			else if ("anchorLock" === name)
				oItem = new CAnchorLock();
			else if ("borderbottom" === name)
				oItem = new CBorder();
			else if ("borderleft" === name)
				oItem = new CBorder();
			else if ("borderright" === name)
				oItem = new CBorder();
			else if ("bordertop" === name)
				oItem = new CBorder();
			else if ("wrap" === name)
				oItem = new CWrap();
			else if ("wrap" === name)
				oItem = new CWrap();
			else if ("ClientData" === name)
				client_data = oItem = new CClientData();
			if (oItem) {
				oItem.fromXml(reader);
			}
			this.items.push(oItem);
			if ((client_data) && client_data.m_oObjectType === EVmlClientDataObjectType.vmlclientdataobjecttypeNote)
					this.m_bComment = true;

		};
		CVmlCommonElements.prototype.writeAttrXmlImpl = function (writer) {

			writer.WriteXmlNullableAttributeString("id", this.m_sId);
			if (this.m_oStyle)
				writer.WriteXmlNullableAttributeString("style", this.m_oStyle.ToString());
			writer.WriteXmlNullableAttributeString("href", this.m_sHref);
			writer.WriteXmlNullableAttributeString("target", this.m_sTarget);
			writer.WriteXmlNullableAttributeString("class", this.m_sClass);
			writer.WriteXmlNullableAttributeString("title", this.m_sTitle);
			writer.WriteXmlNullableAttributeString("alt", this.m_sAlt);

			if (this.m_oCoordSize)
				writer.WriteXmlNullableAttributeString("coordsize", this.m_oCoordSize.ToString());
			if (this.m_oCoordOrigin)
				writer.WriteXmlNullableAttributeString("coordorigin", this.m_oCoordOrigin.ToString());
			if (this.m_oWrapCoords)
				writer.WriteXmlNullableAttributeString("wrapcoords", this.m_oWrapCoords.ToString());

			// if (true !== this.m_oPrint)
			// 	writer.WriteXmlNullableAttributeString("print", "false");

			writer.WriteXmlNullableAttributeString("o:spid", this.m_sSpId);

			if (false !== this.m_oOned)
				writer.WriteXmlNullableAttributeString("o:oned", "true");

			writer.WriteXmlNullableAttributeInt("o:regroupid", this.m_oRegroupId);

			if (false !== this.m_oDoubleClickNotify)
				writer.WriteXmlNullableAttributeString("o:doubleclicknotify", "true");

			if (false !== this.m_oButton)
				writer.WriteXmlNullableAttributeString("o:button", "true");

			if (false !== this.m_oUserHidden)
				writer.WriteXmlNullableAttributeString("o:userhidden", "true");

			if (false !== this.m_oBullet)
				writer.WriteXmlNullableAttributeString("o:bullet", "true");

			if (false !== this.m_oHr)
				writer.WriteXmlNullableAttributeString("o:hr", "true");

			if (false !== this.m_oHrStd)
				writer.WriteXmlNullableAttributeString("o:hrstd", "true");

			if (false !== this.m_oHrNoShade)
				writer.WriteXmlNullableAttributeString("o:hrnoshade", "true");

			if (0 !== this.m_oHrPct)
				writer.WriteXmlNullableAttributeDouble("o:hrpct", this.m_oHrPct);

			if (this.m_oHrAlign !== null && AscCommon.align_Left !== this.m_oHrAlign) {
				switch (this.m_oHrAlign) {
					case AscCommon.align_Center: {
						writer.WriteXmlNullableAttributeString("o:hralign", "center");
						break;
					}
					case AscCommon.align_Left: {
						writer.WriteXmlNullableAttributeString("o:hralign", "left");
						break;
					}
					case AscCommon.align_Right: {
						writer.WriteXmlNullableAttributeString("o:hralign", "right");
						break;
					}
				}
			}

			if (false !== this.m_oAllowInCell)
				writer.WriteXmlNullableAttributeString("o:allowincell", "true");

			if (false !== this.m_oAllowOverlap)
				writer.WriteXmlNullableAttributeString("o:allowoverlap", "true");

			if (false !== this.m_oUserDrawn)
				writer.WriteXmlNullableAttributeString("o:userdrawn", "true");

			writer.WriteXmlNullableAttributeString("o:bordertopcolor", getBooleanTrueFalse(this.m_oBorderTopColor));
			writer.WriteXmlNullableAttributeString("o:borderleftcolor", getBooleanTrueFalse(this.m_oBorderLeftColor));
			writer.WriteXmlNullableAttributeString("o:borderbottomcolor", getBooleanTrueFalse(this.m_oBorderBottomColor));
			writer.WriteXmlNullableAttributeString("o:borderrightcolor", getBooleanTrueFalse(this.m_oBorderRightColor));

			writer.WriteXmlNullableAttributeInt("o:dgmlayout", this.m_oDgmLayout);
			writer.WriteXmlNullableAttributeInt("o:dgmlayoutmru", this.m_oDgmLayoutMru);
			writer.WriteXmlNullableAttributeString("o:dgmnodekind", this.m_oDgmNodeKind);

			if (EInsetMode.insetmodeCustom !== this.m_oInsetMode)
				writer.WriteXmlNullableAttributeString("o:insetmode", getInsetMode(this.m_oInsetMode));

			writer.WriteXmlNullableAttributeString("chromakey", getColorType(this.m_oChromaKey));

			writer.WriteXmlNullableAttributeString("filled", getBooleanTrueFalse(this.m_oFilled));

			writer.WriteXmlNullableAttributeString("fillcolor", getColorType(this.m_oFillColor));
			writer.WriteXmlNullableAttributeString("opacity", getCVml_1_65536(this.m_oOpacity));

			writer.WriteXmlNullableAttributeString("stroked", getBooleanTrueFalse(this.m_oStroked));

			writer.WriteXmlNullableAttributeString("strokecolor", getColorType(this.m_oStrokeColor));

			if (this.m_oStrokeWeight !== null)
				writer.WriteXmlNullableAttributeInt("strokeweight", this.m_oStrokeWeight);

			writer.WriteXmlNullableAttributeString("insetpen", getBooleanTrueFalse(this.m_oInsetPen));

			if (this.m_oSpt !== null)
				writer.WriteXmlNullableAttributeInt("o:spt", this.m_oSpt);

			if (this.m_oConnectorType !== null) {
				writer.WriteXmlNullableAttributeString("o:connectortype", this.m_oConnectorType);
			}

			writer.WriteXmlNullableAttributeString("o:bwmode", getBWMode(this.m_oBwMode));
			writer.WriteXmlNullableAttributeString("o:bwpure", getBWMode(this.m_oBwPure));
			writer.WriteXmlNullableAttributeString("o:bwnormal", getBWMode(this.m_oBwNormal));

			if (false !== this.m_oForceDash)
				writer.WriteXmlNullableAttributeString("o:forcedash", "true");

			if (false !== this.m_oOleIcon)
				writer.WriteXmlNullableAttributeString("o:oleicon", "true");

			if (false !== this.m_oOle)
				writer.WriteXmlNullableAttributeString("o:ole", "true");

			if (false !== this.m_oPreferRelative)
				writer.WriteXmlNullableAttributeString("o:preferrelative", "true");

			if (false !== this.m_oClipToWrap)
				writer.WriteXmlNullableAttributeString("o:cliptowrap", "true");

			writer.WriteXmlNullableAttributeString("o:clip", getBooleanTrueFalse(this.m_oClip));
		};
		CVmlCommonElements.prototype.writeChildren = function (writer) {

			for (let i = 0; i < this.items.length; ++i) {
				let oItem = this.items[i];
				if (oItem) {
					let sName = null;
					if (oItem instanceof CCallout)
						sName = "o:callout";
					else if (oItem instanceof CClipPath)
						sName = "o:clippath";
					else if (oItem instanceof CExtrusion)
						sName = "o:extrusion";
					else if (oItem instanceof CLock)
						sName = "o:lock";
					else if (oItem instanceof CSignatureLine)
						sName = "o:signatureline";
					else if (oItem instanceof CSkew)
						sName = "o:skew";
					else if (oItem instanceof CFillVml)
						sName = "v:fill";
					else if (oItem instanceof CFormulas)
						sName = "v:formulas";
					else if (oItem instanceof CHandles)
						sName = "v:handles";
					else if (oItem instanceof CImageData)
						sName = "v:imagedata";
					else if (oItem instanceof CPath)
						sName = "v:path";
					else if (oItem instanceof CShadow)
						sName = "v:shadow";
					else if (oItem instanceof CStroke)
						sName = "v:stroke";
					else if (oItem instanceof CTextbox)
						sName = "v:textbox";
					else if (oItem instanceof CTextPath)
						sName = "v:textpath";
					else if (oItem instanceof CAnchorLock)
						sName = "wd:anchorLock";
					else if (oItem instanceof CBorder)
						sName = oItem.m_sType;
					else if (oItem instanceof CWrap)
						sName = "w10:wrap";
					else if (oItem instanceof CClientData)
						sName = "x:ClientData";
					oItem.toXml(writer, sName);
				}
			}
		};


		function CArc() {
			CVmlCommonElements.call(this);

			this.m_oEndAngle = null;
		}

		IC(CArc, CVmlCommonElements, AscDFH.historyitem_type_VMLArc);
		CArc.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "endangle":
				case "endAngle": {
					this.m_oEndAngle = reader.GetValueInt();
					return;
				}
				case "startangle":
				case "startAngle": {
					this.m_oStartAngle = reader.GetValueInt();
					return;
				}
			}
			CVmlCommonElements.prototype.readAttrXml.call(this, name, reader);
		};
		CArc.prototype.writeAttrXmlImpl = function (writer) {
			CVmlCommonElements.prototype.writeAttrXmlImpl.call(this, writer);
			writer.WriteXmlNullableAttributeInt("startangle", this.m_oStartAngle);
			writer.WriteXmlNullableAttributeInt("endangle", this.m_oEndAngle);
		};

		function CCurve() {
			CVmlCommonElements.call(this);
			this.m_oFrom = null;
			this.m_oControl1 = null;
			this.m_oControl2 = null;
			this.m_oTo = null;
		}

		IC(CCurve, CVmlCommonElements, AscDFH.historyitem_type_VMLCurve);
		CCurve.prototype.readAttrXml = function (name, reader) {
			if ("control1" === name) {
				this.m_oControl1 = new CVml_Vector2D_Units(reader.GetValue());
				return;
			} else if ("control2" === name) {
				this.m_oControl2 = new CVml_Vector2D_Units(reader.GetValue());
				return;
			} else if ("from" === name) {
				this.m_oFrom = new CVml_Vector2D_Units(reader.GetValue());
				return;
			} else if ("to" === name) {
				this.m_oTo = new CVml_Vector2D_Units(reader.GetValue());
				return;
			}
			CVmlCommonElements.prototype.readAttrXml.call(this, name, reader);
		};
		CCurve.prototype.writeAttrXmlImpl = function (writer) {
			CVmlCommonElements.prototype.writeAttrXmlImpl.call(this, writer);
			this.m_oFrom && writer.WriteXmlNullableAttributeString("from", this.m_oFrom.ToString());
			this.m_oControl1 && writer.WriteXmlNullableAttributeString("control1", this.m_oControl1.ToString());
			this.m_oControl2 && writer.WriteXmlNullableAttributeString("control2", this.m_oControl2.ToString());
			this.m_oTo && writer.WriteXmlNullableAttributeString("to", this.m_oTo.ToString());

		};

		function CF() {
			CBaseNoId.call(this);
			this.m_sEqn = null;
		}

		IC(CF, CBaseNoId, 0);
		CF.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "eqn": {
					this.m_sEqn = reader.GetValue();
					break;
				}
			}
		};
		CF.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("eqn", this.m_sEqn);

		};


		let EFillMethod =
			{
				fillmethodAny: 0,
				fillmethodLinear: 1,
				fillmethodLinearSigma: 2,
				fillmethodSigma: 3,
				fillmethodNone: 4
			};

		function readFillMethod(reader) {
			let sValue = reader.GetValue();
			if ("any" === sValue) return EFillMethod.fillmethodAny;
			else if ("linear" === sValue) return EFillMethod.fillmethodLinear;
			else if ("linear sigma" === sValue) return EFillMethod.fillmethodLinearSigma;
			else if ("sigma" === sValue) return EFillMethod.fillmethodSigma;
			else if ("none" === sValue) return EFillMethod.fillmethodNone;
			else return EFillMethod.fillmethodNone;
		}

		function getFillMethod(nType) {
			if (nType === EFillMethod.fillmethodAny) return "any";
			if (nType === EFillMethod.fillmethodLinear) return "linear";
			if (nType === EFillMethod.fillmethodLinearSigma) return "linear sigma";
			if (nType === EFillMethod.fillmethodSigma) return "sigma";
			if (nType === EFillMethod.fillmethodNone) return "none";
			return null;

		}

		function CFixedPercentage(sVal) {
			this.m_dValue = 0;
			if (sVal) {
				this.FromString(sVal);
			}
		}

		CFixedPercentage.prototype.GetValue = function () {
			return this.m_dValue;
		}

		CFixedPercentage.prototype.SetValue = function (dValue) {
			this.m_dValue = Math.min(100.0, Math.max(-100.0, dValue));
		}

		CFixedPercentage.prototype.FromString = function (sValue) {
			let nPos = sValue.indexOf('%');
			let nLen = sValue.length;
			if (-1 === nPos || nPos !== sValue.length - 1 || nLen <= 0) {
				if (-1 === nPos && nLen > 0) {
					let nValue = Math.min(100000, Math.max(-100000, parseInt(sValue)));
					this.m_dValue = nValue / 1000.0;
				} else
					this.m_dValue = 0;
			} else
				this.m_dValue = Math.min(100.0, Math.max(-100.0, (sValue.substr(0, nLen - 1))));

			return this.m_dValue;
		}

		CFixedPercentage.prototype.ToString = function () {
			return this.m_dValue + "%";
		};


		function CFillVml() {
			CBaseNoId.call(this);
			// Attributes
			this.m_oAlignShape = null;
			this.m_sAltHref = null;
			this.m_oAngle = null;
			this.m_oAspect = null;
			this.m_oColor = null;
			this.m_oColor2 = null;
			this.m_arrColors = null;
			this.m_oDetectMouseClick = null;
			this.m_oFocus = null;
			this.m_oFocusPosition = null;
			this.m_oFocusSize = null;
			this.m_sHref = null;
			this.m_rId = null;
			this.m_sId = null;
			this.m_oMethod = null;
			this.m_oOn = null;
			this.m_oOpacity = null;
			this.m_oOpacity2 = null;
			this.m_oOrigin = null;
			this.m_oPosition = null;
			this.m_oRecolor = null;
			this.m_oRelId = null;
			this.m_oRotate = null;
			this.m_oSize = null;
			this.m_sSrc = null;
			this.m_sTitle = null;
			this.m_oType = null;

			// Childs
			this.m_oFill = null;
		}

		IC(CFillVml, CBaseNoId, 0);
		CFillVml.prototype.readAttrXml = function (name, reader) {

			if ("aspect" === name) this.m_oAspect = readImageAspect(reader);
			else if ("angle" === name) this.m_oAngle = reader.GetValueInt();
			else if ("alignshape" === name) this.m_oAlignShape = reader.GetValueBool();
			else if ("color" === name) this.m_oColor = readColorType(reader);
			else if ("color2" === name) this.m_oColor2 = readColorType(reader);
			else if ("colors" === name) this.sColors = reader.GetValue();
			else if ("id" === name) this.m_sId = reader.GetValue();
			else if ("method" === name) this.m_oMethod = readFillMethod(reader);
			else if ("focus" === name) this.m_oFocus = new CFixedPercentage(reader.GetValue());
			else if ("focussize" === name) this.m_oFocusSize = new CVml_Vector2D_Units_Or_Percentage(reader.GetValue());
			else if ("focusposition" === name) this.m_oFocusPosition = new CVml_Vector2D_Units_Or_Percentage(reader.GetValue());
			else if ("on" === name) this.m_oOn = reader.GetValueBool();
			else if ("opacity" === name) this.m_oOpacity = readCVml_1_65536(reader);
			else if ("href" === name) this.m_sHref = reader.GetValue();
			else if ("althref" === name) this.m_sAltHref = reader.GetValue();
			else if ("origin" === name) this.m_oOrigin = new CVml_Vector2D_Units_Or_Percentage(reader.GetValue());
			else if ("detectmouseclick" === name) this.m_oDetectMouseClick = reader.GetValueBool();
			else if ("title" === name) this.m_sTitle = reader.GetValue();
			else if ("opacity2" === name) this.m_oOpacity2 = readCVml_1_65536(reader);
			else if ("relid" === name) this.m_oRelId = reader.GetValue();
			else if ("position" === name) this.m_oPosition = new CVml_Vector2D_Units_Or_Percentage(reader.GetValue());
			else if ("recolor" === name) this.m_oRecolor = reader.GetValueBool();
			else if ("rotate" === name) this.m_oRotate = reader.GetValueBool();
			else if ("id" === name) this.m_rId = reader.GetValue();
			else if ("src" === name) this.m_sSrc = reader.GetValue();
			else if ("size" === name) this.m_oSize = new CVml_Vector2D_Units_Or_Percentage(reader.GetValue());
			else if ("type" === name) this.m_oType = readFillType(reader);
		};
		CFillVml.prototype.readChildXml = function (name, reader) {
			if (name === "fill") {
				this.m_oFill = new CFill();
				this.m_oFill.fromXml(reader);
			}
		};
		CFillVml.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("id", this.m_sId);
			writer.WriteXmlNullableAttributeString("type", getFillType(this.m_oType));

			if ((this.m_oOn !== null) && (true !== this.m_oOn))
				writer.WriteXmlNullableAttributeString("on", "false");

			if (this.m_oOpacity !== null)
				writer.WriteXmlNullableAttributeString("opacity", getCVml_1_65536(this.m_oOpacity));

			writer.WriteXmlNullableAttributeString("color", getColorType(this.m_oColor));
			writer.WriteXmlNullableAttributeString("color2", getColorType(this.m_oColor2));
			writer.WriteXmlNullableAttributeString("src", this.m_sSrc);
			writer.WriteXmlNullableAttributeString("o:href", this.m_sHref);
			writer.WriteXmlNullableAttributeString("o:althref", this.m_sAltHref);
			this.m_oSize && writer.WriteXmlNullableAttributeString("size", this.m_oSize.ToString());
			this.m_oOrigin && writer.WriteXmlNullableAttributeString("origin", this.m_oOrigin.ToString());
			this.m_oPosition && writer.WriteXmlNullableAttributeString("position", this.m_oPosition.ToString());

			if ((this.m_oAspect !== null) && (EImageAspect.imageaspectIgnore !== this.m_oAspect))
				writer.WriteXmlNullableAttributeString("aspect", getImageAspect(this.m_oAspect));


			writer.WriteXmlNullableAttributeInt("angle", this.m_oAngle);

			if ((this.m_oAlignShape !== null) && (true !== this.m_oAlignShape))
				writer.WriteXmlNullableAttributeString("alignshape", "false");

			if ((this.m_oFocus !== null) && (0 !== this.m_oFocus))
				writer.WriteXmlNullableAttributeString("focus", this.m_oFocus.ToString());

			if ((this.m_oFocusPosition !== null) && (0 !== this.m_oFocusPosition.GetX() || 0 !== this.m_oFocusPosition.GetY()))
				writer.WriteXmlNullableAttributeString("focusposition", this.m_oFocusPosition.ToString());

			if ((this.m_oFocusSize !== null) && (0 !== this.m_oFocusSize.GetX() || 0 !== this.m_oFocusSize.GetY()))
				writer.WriteXmlNullableAttributeString("focussize", this.m_oFocusSize.ToString());

			if ((this.m_oMethod !== null) && (EFillMethod.fillmethodSigma !== this.m_oMethod))
				writer.WriteXmlNullableAttributeString("method", getFillMethod(this.m_oMethod));

			writer.WriteXmlNullableAttributeString("o:detectmouseclick", getBooleanTrueFalse(this.m_oDetectMouseClick));
			writer.WriteXmlNullableAttributeString("o:title", this.m_sTitle);

			if (this.m_oOpacity2 !== null)
				writer.WriteXmlNullableAttributeString("o:opacity2", getCVml_1_65536(this.m_oOpacity2));

			if ((this.m_oRecolor !== null) && (false !== this.m_oRecolor))
				writer.WriteXmlNullableAttributeString("recolor", "true");

			if ((this.m_oRotate !== null) && (false !== this.m_oRotate))
				writer.WriteXmlNullableAttributeString("rotate", "true");

			writer.WriteXmlNullableAttributeString("r:id", this.m_oRelId);
			writer.WriteXmlNullableAttributeString("o:relid", this.m_oRelId);
		};
		CFillVml.prototype.writeChildren = function (writer) {
			this.m_oFill && this.m_oFill.toXml(writer, "o:fill");
		};

		let EScreenSize =
			{
				screensize1024x768: 0,
				screensize1152x862: 1,
				screensize544x376: 2,
				screensize640x480: 3,
				screensize720x512: 4,
				screensize800x600: 5
			};

		function readScreenSize(reader) {
			let sValue = reader.GetValue();
			if ("1024,768" === sValue) return EScreenSize.screensize1024x768;
			else if ("1152,862" === sValue) return EScreenSize.screensize1152x862;
			else if ("544,376" === sValue) return EScreenSize.screensize544x376;
			else if ("640,480" === sValue) return EScreenSize.screensize640x480;
			else if ("720,512" === sValue) return EScreenSize.screensize720x512;
			else if ("800,600" === sValue) return EScreenSize.screensize800x600;
			return EScreenSize.screensize640x480;
		}

		function getScreenSize(nType) {
			if (nType === EScreenSize.screensize1024x768) return "1024,768";
			if (nType === EScreenSize.screensize1152x862) return "1152,862";
			if (nType === EScreenSize.screensize544x376) return "544,376";
			if (nType === EScreenSize.screensize640x480) return "640,480";
			if (nType === EScreenSize.screensize720x512) return "720,512";
			if (nType === EScreenSize.screensize800x600) return "800,600";
			return "640,480";
		}

		function CBackground() {
			CVmlCommonElements.call(this);
		}

		IC(CBackground, CVmlCommonElements, 0);
		CBackground.prototype.readAttrXml = function (name, reader) {
			if (name === "targetscreensize") {
				this.m_oTargetScreenSize = readScreenSize(reader);
				return;
			}
			CVmlCommonElements.prototype.readAttrXml.call(this, name, reader);
		};
		CBackground.prototype.writeAttrXmlImpl = function (writer) {
			CVmlCommonElements.prototype.writeAttrXmlImpl.call(this, writer);
			writer.WriteXmlNullableAttributeString("o:targetscreensize", getScreenSize(this.m_oTargetScreenSize));
		};

		function CFormulas() {
			CBaseNoId.call(this);
			this.items = [];
		}

		IC(CFormulas, CBaseNoId, 0);
		CFormulas.prototype.readChildXml = function (name, reader) {
			switch (name) {
				case "f": {
					let oF = new CF();
					oF.fromXml(reader);
					this.items.push(oF);
					break;
				}
			}
		};
		CFormulas.prototype.writeChildren = function (writer) {

			for (let i = 0; i < this.items.length; ++i) {
				if (this.items[i]) {
					this.items[i].toXml(writer, "v:f");
				}
			}
		};


		function CH() {
			CBaseNoId.call(this);
			this.m_oInvX = null;
			this.m_oInvY = null;
			this.m_oMap = null;
			this.m_oPolar = null;
			this.m_oPosition = null;
			this.m_oRadiusRange = null;
			this.m_oSwitch = null;
			this.m_oXRange = null;
			this.m_oYRange = null;
		}

		IC(CH, CBaseNoId, 0);
		CH.prototype.readAttrXml = function (name, reader) {
			if ("invx" === name) this.m_oInvX = reader.GetValueBool();
			else if ("invy" === name) this.m_oInvY = reader.GetValueBool();
			else if ("map" === name) this.m_oMap = new CVml_Vector2D(reader.GetValue());
			else if ("position" === name) this.m_oPosition = new CVml_Vector2D_Position(reader.GetValue());
			else if ("polar" === name) this.m_oPolar = new CVml_Vector2D(reader.GetValue());
			else if ("radiusrange" === name) this.m_oRadiusRange = new CVml_Vector2D(reader.GetValue());
			else if ("switch" === name) this.m_oSwitch = reader.GetValueBool();
			else if ("xrange" === name) this.m_oXRange = new CVml_Vector2D(reader.GetValue());
			else if ("yrange" === name) this.m_oYRange = new CVml_Vector2D(reader.GetValue());

		};
		CH.prototype.writeAttrXmlImpl = function (writer) {
			if (this.m_oPosition) {
				if (EVml_Vector2D_Position.vmlvector2dposConstant !== this.m_oPosition.GetTypeX() || EVml_Vector2D_Position.vmlvector2dposConstant !== this.m_oPosition.GetTypeY() || 0 !== this.m_oPosition.GetX() || 0 !== this.m_oPosition.GetY())
					writer.WriteXmlNullableAttributeString("position", this.m_oPosition.ToString());
			}

			this.m_oPolar && writer.WriteXmlNullableAttributeString("polar", this.m_oPolar.ToString());

			if (this.m_oMap) {
				if (0 !== this.m_oMap.GetX() || 1000 !== this.m_oMap.GetY())
					writer.WriteXmlNullableAttributeString("map", this.m_oMap.ToString());
			}

			if (false !== this.m_oInvX)
				writer.WriteXmlNullableAttributeString("invx", "true");

			if (false !== this.m_oInvY)
				writer.WriteXmlNullableAttributeString("invy", "true");

			if (false !== this.m_oSwitch)
				writer.WriteXmlNullableAttributeString("switch", "true");

			if (this.m_oXRange) {
				if (0 !== this.m_oXRange.GetX() || 0 !== this.m_oXRange.GetY())
					writer.WriteXmlNullableAttributeString("xrange", this.m_oXRange.ToString());
			}
			if (this.m_oYRange) {
				if (0 !== this.m_oYRange.GetX() || 0 !== this.m_oYRange.GetY())
					writer.WriteXmlNullableAttributeString("yrange", this.m_oYRange.ToString());
			}
			if (this.m_oRadiusRange) {
				if (0 !== this.m_oRadiusRange.GetX() || 0 !== this.m_oRadiusRange.GetY())
					writer.WriteXmlNullableAttributeString("radiusrange", this.m_oRadiusRange.ToString());
			}


		};


		function CHandles() {
			CBaseNoId.call(this);
			this.items = [];
		}

		IC(CHandles, CBaseNoId, 0);
		CHandles.prototype.readChildXml = function (name, reader) {
			switch (name) {
				case "h": {
					let oPr = new CH();
					oPr.fromXml(reader);
					this.items.push(oPr);
					break;
				}
			}
		};
		CHandles.prototype.writeChildren = function (writer) {
			for (let i = 0; i < this.items.length; ++i) {
				if (this.items[i]) {
					this.items[i].toXml(writer, "v:h");
				}
			}
		};

		function CImage() {
			CVmlCommonElements.call(this);
			this.m_sSrc = null;
			this.m_oCropLeft = null;
			this.m_oCropTop = null;
			this.m_oCropRight = null;
			this.m_oCropBottom = null;
			this.m_oGain = null;
			this.m_oBlackLevel = null;
			this.m_oGamma = null;
			this.m_oGrayscale = null;
			this.m_oBiLevel = null;
		}

		IC(CImage, CVmlCommonElements, AscDFH.historyitem_type_VMLImage);
		CImage.prototype.readAttrXml = function (name, reader) {
			if ("bilevel" === name) {
				this.m_oBiLevel = reader.GetValueBool();
				return;
			} else if ("blacklevel" === name) {
				this.m_oBlackLevel = reader.GetValueDouble();
				return;
			} else if ("cropleft" === name) {
				this.m_oCropLeft = readCVml_1_65536(reader);
				return;
			} else if ("croptop" === name) {
				this.m_oCropTop = readCVml_1_65536(reader);
				return;
			} else if ("cropright" === name) {
				this.m_oCropRight = readCVml_1_65536(reader);
				return;
			} else if ("cropbottom" === name) {
				this.m_oCropBottom = readCVml_1_65536(reader);
				return;
			} else if ("gain" === name) {
				this.m_oGain = reader.GetValueDouble();
				return;
			} else if ("gamma" === name) {
				this.m_oGamma = reader.GetValueDouble();
				return;
			} else if ("grayscale" === name) {
				this.m_oGrayscale = reader.GetValueBool();
				return;
			} else if ("src" === name) {
				this.m_sSrc = reader.GetValue();
				return;
			}
			CVmlCommonElements.prototype.readAttrXml.call(this, name, reader);
		};
		CImage.prototype.writeAttrXmlImpl = function (writer) {

			CVmlCommonElements.prototype.writeAttrXmlImpl.call(this, writer);

			if ("" !== this.m_sSrc)
				writer.WriteXmlNullableAttributeString("src", this.m_sSrc);

			if (0 !== this.m_oCropLeft)
				writer.WriteXmlNullableAttributeString("cropleft", getCVml_1_65536(this.m_oCropLeft));

			if (0 !== this.m_oCropTop)
				writer.WriteXmlNullableAttributeString("croptop", getCVml_1_65536(this.m_oCropTop));

			if (0 !== this.m_oCropRight)
				writer.WriteXmlNullableAttributeString("cropright", getCVml_1_65536(this.m_oCropRight));

			if (0 !== this.m_oCropBottom)
				writer.WriteXmlNullableAttributeString("cropbottom", getCVml_1_65536(this.m_oCropBottom));

			if (1 !== this.m_oGain)
				writer.WriteXmlNullableAttributeDouble("gain", this.m_oGain);

			if (0 !== this.m_oBlackLevel)
				writer.WriteXmlNullableAttributeDouble("blacklevel", this.m_oBlackLevel);

			if (1 !== this.m_oGamma)
				writer.WriteXmlNullableAttributeDouble("gamma", this.m_oGamma);

			if (false !== this.m_oGrayscale)
				writer.WriteXmlNullableAttributeString("grayscale", "true");

			if (false !== this.m_oBiLevel)
				writer.WriteXmlNullableAttributeString("bilevel", "true");
		};


		function CImageData() {
			CBaseNoId.call(this);
			this.m_sAltHref = null;
			this.m_oBiLevel = null;
			this.m_oBlackLevel = null;
			this.m_oChromaKey = null;
			this.m_oCropLeft = null;
			this.m_oCropTop = null;
			this.m_oCropRight = null;
			this.m_oCropBottom = null;
			this.m_oDetectMouseClick = null;
			this.m_oEmbossColor = null;
			this.m_oGain = null;
			this.m_oGamma = null;
			this.m_oGrayscale = null;
			this.m_rHref = null;
			this.m_oHref = null;
			this.m_rId = null;
			this.m_oId = null;
			this.m_oMovie = null;
			this.m_oOleId = null;
			this.m_rPict = null;
			this.m_oRecolorTarget = null;
			this.m_oRelId = null;
			this.m_sSrc = null;
			this.m_sTitle = null;
		}

		IC(CImageData, CBaseNoId, 0);
		CImageData.prototype.readAttrXml = function (name, reader) {
			if ("bilevel" === name) this.m_oBiLevel = reader.GetValueBool();
			else if ("blacklevel" === name) this.m_oBlackLevel = reader.GetValueDouble();
			else if ("cropleft" === name) this.m_oCropLeft = readCVml_1_65536(reader);
			else if ("croptop" === name) this.m_oCropTop = readCVml_1_65536(reader);
			else if ("cropright" === name) this.m_oCropRight = readCVml_1_65536(reader);
			else if ("cropbottom" === name) this.m_oCropBottom = readCVml_1_65536(reader);
			else if ("embosscolor" === name) this.m_oEmbossColor = readColorType(reader);
			else if ("gain" === name) this.m_oGain = reader.GetValueDouble();
			else if ("gamma" === name) this.m_oGamma = reader.GetValueDouble();
			else if ("grayscale" === name) this.m_oGrayscale = reader.GetValueBool();
			else if ("id" === name) {
				let sName = reader.GetName();
				if (sName === "id") {
					this.m_oId = reader.GetValue();
				} else if (sName === "r:id" || sName === "relationships:id") {
					this.m_rId = reader.GetValue();
				}
			} else if ("detectmouseclick" === name) this.m_oDetectMouseClick = reader.GetValue();
			else if ("href" === name) {
				let sName = reader.GetName();
				if (sName === "o:href") {
					this.m_oHref = reader.GetValue();
				} else if (sName === "r:href") {
					this.m_rHref = reader.GetValue();
				}
			} else if ("oleid" === name) this.m_oOleId = reader.GetValue();
			else if ("title" === name) this.m_sTitle = reader.GetValue();
			else if ("pict" === name) this.m_rPict = reader.GetValue();
			else if ("recolortarget" === name) this.m_oRecolorTarget = readColorType(reader);
			else if ("src" === name) this.m_sSrc = reader.GetValue();
		}
		CImageData.prototype.writeAttrXmlImpl = function (writer) {

			writer.WriteXmlNullableAttributeString("id", this.m_oId);

			if ("" !== this.m_sSrc)
				writer.WriteXmlNullableAttributeString("src", this.m_sSrc);

			writer.WriteXmlNullableAttributeString("cropleft", getCVml_1_65536(this.m_oCropLeft));
			writer.WriteXmlNullableAttributeString("croptop", getCVml_1_65536(this.m_oCropTop));
			writer.WriteXmlNullableAttributeString("cropright", getCVml_1_65536(this.m_oCropRight));
			writer.WriteXmlNullableAttributeString("cropbottom", getCVml_1_65536(this.m_oCropBottom));

			if (1 !== this.m_oGain)
				writer.WriteXmlNullableAttributeDouble("gain", this.m_oGain);

			if (0 !== this.m_oBlackLevel)
				writer.WriteXmlNullableAttributeDouble("blacklevel", this.m_oBlackLevel);

			if (1 !== this.m_oGamma)
				writer.WriteXmlNullableAttributeDouble("gamma", this.m_oGamma);

			if (false !== this.m_oGrayscale)
				writer.WriteXmlNullableAttributeString("grayscale", "true");

			if (false !== this.m_oBiLevel)
				writer.WriteXmlNullableAttributeString("bilevel", "true");

			//	writer.WriteXmlNullableAttributeString("chromakey",    this.m_oChromaKey );
			writer.WriteXmlNullableAttributeString("embosscolor", getColorType(this.m_oEmbossColor));

			writer.WriteXmlNullableAttributeString("o:href", this.m_oHref);
			writer.WriteXmlNullableAttributeString("o:althref", this.m_sAltHref);
			writer.WriteXmlNullableAttributeString("o:title", this.m_sTitle);
			writer.WriteXmlNullableAttributeString("o:oleid", this.m_oOleId);
			writer.WriteXmlNullableAttributeString("o:detectmouseclick", this.m_oDetectMouseClick);
			//writer.WriteXmlNullableAttributeString("o:movie",            this.m_oMovie );
			//writer.WriteXmlNullableAttributeString("o:relid",            this.m_oRelId );

			writer.WriteXmlNullableAttributeString("r:id", this.m_rId);
			writer.WriteXmlNullableAttributeString("r:pict", this.m_rPict);
			writer.WriteXmlNullableAttributeString("r:href", this.m_rHref);
		};


		function CLine() {
			CVmlCommonElements.call(this);
			this.m_oFrom = null;
			this.m_oTo = null;
		}

		IC(CLine, CVmlCommonElements, AscDFH.historyitem_type_VMLLine);
		CLine.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "from":
					this.m_oFrom = new CVml_Vector2D_Units(reader.GetValue());
					return;
				case "to":
					this.m_oTo = new CVml_Vector2D_Units(reader.GetValue());
					return;
			}
			CVmlCommonElements.prototype.readAttrXml.call(this, name, reader);
		};
		CLine.prototype.writeAttrXmlImpl = function (writer) {

			CVmlCommonElements.prototype.writeAttrXmlImpl.call(this, writer);
			this.m_oFrom && writer.WriteXmlNullableAttributeString("from", this.m_oFrom.ToString());
			this.m_oTo && writer.WriteXmlNullableAttributeString("to", this.m_oTo.ToString());

		};

		function COval() {
			CVmlCommonElements.call(this);
		}

		IC(COval, CVmlCommonElements, AscDFH.historyitem_type_VMLOval);

		function CPath() {
			CBaseNoId.call(this);
			this.m_oArrowOk = null;
			this.m_oConnectAngles = null;
			this.m_oConnectLocs = null;
			this.m_oConnectType = null;
			this.m_oExtrusionOk = null;
			this.m_oFillOk = null;
			this.m_oGradientShapeOk = null;
			this.m_oId = null;
			this.m_oInsetPenOk = null;
			this.m_oLimo = null;
			this.m_oShadowOk = null;
			this.m_oStrokeOk = null;
			this.m_oTextBoxRect = null;
			this.m_oTextPathOk = null;
			this.m_oV = null;
		}

		IC(CPath, CBaseNoId, 0);
		CPath.prototype.readAttrXml = function (name, reader) {
			if ("arrowok" === name) this.m_oArrowOk = reader.GetValueBool();
			else if ("fillok" === name) this.m_oFillOk = reader.GetValueBool();
			else if ("gradientshapeok" === name) this.m_oGradientShapeOk = reader.GetValueBool();
			else if ("id" === name) this.m_oId = reader.GetValue();
			else if ("insetpenok" === name) this.m_oInsetPenOk = reader.GetValueBool();
			else if ("limo" === name) this.m_oLimo = new CVml_Vector2D_Units(reader.GetValue());
			else if ("connectangles" === name) this.m_oConnectAngles = reader.GetValue();
			else if ("connectlocs" === name) this.m_oConnectLocs = reader.GetValue();
			else if ("connecttype" === name) this.m_oConnectType = readConnectType(reader);
			else if ("extrusionok" === name) this.m_oExtrusionOk = reader.GetValueBool();
			else if ("shadowok" === name) this.m_oShadowOk = reader.GetValueBool();
			else if ("strokeok" === name) this.m_oStrokeOk = reader.GetValueBool();
			else if ("textboxrect" === name) this.m_oTextBoxRect = new CVml_Polygon2D(reader.GetValue());
			else if ("textpathok" === name) this.m_oTextPathOk = reader.GetValue();
			else if ("v" === name) this.m_oV = reader.GetValue();
		};
		CPath.prototype.writeAttrXmlImpl = function (writer) {


			writer.WriteXmlNullableAttributeString("id", this.m_oId);
			writer.WriteXmlNullableAttributeString("v", this.m_oV);

			if (this.m_oLimo) {
				if (0 !== this.m_oLimo.GetX() || 0 !== this.m_oLimo.GetY())
					writer.WriteXmlNullableAttributeString("limo", this.m_oLimo.ToString());
			}

			if (this.m_oTextBoxRect) {
				writer.WriteXmlNullableAttributeString("textboxrect", this.m_oTextBoxRect.ToString());
			}

			if (true !== this.m_oFillOk)
				writer.WriteXmlNullableAttributeString("fillok", "false");

			if (true !== this.m_oStrokeOk)
				writer.WriteXmlNullableAttributeString("strokeok", "false");

			if (true !== this.m_oShadowOk)
				writer.WriteXmlNullableAttributeString("shadowok", "false");

			if (false !== this.m_oArrowOk)
				writer.WriteXmlNullableAttributeString("arrowok", "true");

			if (false !== this.m_oGradientShapeOk)
				writer.WriteXmlNullableAttributeString("gradientshapeok", "true");

			if (false !== this.m_oTextPathOk)
				writer.WriteXmlNullableAttributeString("textpathok", "true");

			if (false !== this.m_oInsetPenOk)
				writer.WriteXmlNullableAttributeString("insetpenok", "true");

			if (EConnectType.connecttypeNone !== this.m_oConnectType)
				writer.WriteXmlNullableAttributeString("o:connecttype", getConnectType(this.m_oConnectType));

			writer.WriteXmlNullableAttributeString("o:connectlocs", this.m_oConnectLocs);
			writer.WriteXmlNullableAttributeString("o:connectangles", this.m_oConnectAngles);

			if (true !== this.m_oExtrusionOk)
				writer.WriteXmlNullableAttributeString("extrusionok", "false");
		};

		function CPolyLine() {
			CVmlCommonElements.call(this);
			this.m_oPoints = null;
		}

		IC(CPolyLine, CVmlCommonElements, AscDFH.historyitem_type_VMLPolyLine);
		CPolyLine.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "points":
					this.m_oPoints = new CVml_Polygon2D_Units(reader.GetValue());
					return;
			}
			CVmlCommonElements.prototype.readAttrXml.call(this, name, reader);
		};
		CPolyLine.prototype.writeAttrXmlImpl = function (writer) {

			CVmlCommonElements.prototype.writeAttrXmlImpl.call(this, writer);

			this.m_oPoints && writer.WriteXmlNullableAttributeString("points", this.m_oPoints.ToString());

		};

		function CRect() {
			CVmlCommonElements.call(this);
		}

		IC(CRect, CVmlCommonElements, AscDFH.historyitem_type_VMLRect);

		function CRoundRect() {
			CVmlCommonElements.call(this);

			this.m_oArcSize = null;
		}

		IC(CRoundRect, CVmlCommonElements, AscDFH.historyitem_type_VMLRoundRect);
		CRoundRect.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "arcsize":
					this.m_oArcSize = new CVml_1_65536_Or_Percentage(reader.GetValue());
					return;
			}
			CVmlCommonElements.prototype.readAttrXml.call(this, name, reader);
		};
		CRoundRect.prototype.writeAttrXmlImpl = function (writer) {
			CVmlCommonElements.prototype.writeAttrXmlImpl.call(this, writer);

			this.m_oArcSize && writer.WriteXmlNullableAttributeString("arcsize", this.m_oArcSize.ToString());
		};

		function CShadow() {
			CBaseNoId.call(this);
			this.m_oColor = null;
			this.m_oColor2 = null;
			this.m_oId = null;
			this.m_oMatrix = null;
			this.m_oObscured = null;
			this.m_oOffset = null;
			this.m_oOffset2 = null;
			this.m_oOn = null;
			this.m_oOpacity = null;
			this.m_oOrigin = null;
			this.m_oType = null;
		}

		IC(CShadow, CBaseNoId, 0);
		CShadow.prototype.readAttrXml = function (name, reader) {
			if ("color" === name) this.m_oColor = readColorType(reader);
			else if ("color2" === name) this.m_oColor2 = readColorType(reader);
			else if ("id" === name) this.m_oId = reader.GetValue();
			else if ("matrix" === name) this.m_oMatrix = new CVml_Matrix(reader.GetValue());
			else if ("obscured" === name) this.m_oObscured = reader.GetValueBool();
			else if ("offset" === name) this.m_oOffset = new CVml_Vector2D_Units_Or_Percentage(reader.GetValue());
			else if ("offset2" === name) this.m_oOffset2 = new CVml_Vector2D_Units_Or_Percentage(reader.GetValue());
			else if ("on" === name) this.m_oOn = reader.GetValueBool();
			else if ("opacity" === name) this.m_oOpacity = readCVml_1_65536(reader);
			else if ("origin" === name) this.m_oOrigin = new CVml_Vector2D_Units_Or_Percentage(reader.GetValue());
			else if ("type" === name) this.m_oType = readShadowType(reader);
		};
		CShadow.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("id", this.m_oId);

			if (true !== this.m_oOn)
				writer.WriteXmlNullableAttributeString("on", "false");

			if (EShadowType.shadowtypeSingle !== this.m_oType)
				writer.WriteXmlNullableAttributeString("type", getShadowType(this.m_oType));

			if (false !== this.m_oObscured)
				writer.WriteXmlNullableAttributeString("obscured", "true");

			if (this.m_oColor) {
				if (128 !== this.m_oColor.Get_R() || 128 !== this.m_oColor.Get_G() || 128 !== this.m_oColor.Get_B())
					writer.WriteXmlNullableAttributeString("color", this.m_oColor.ToString());
			}

			if (this.m_oOpacity !== null)
				writer.WriteXmlNullableAttributeString("opacity", getCVml_1_65536(this.m_oOpacity));

			this.m_oOffset && writer.WriteXmlNullableAttributeString("offset", this.m_oOffset.ToString());

			if (this.m_oColor2) {
				if (203 !== this.m_oColor2.Get_R() || 203 !== this.m_oColor2.Get_G() || 203 !== this.m_oColor2.Get_B())
					writer.WriteXmlNullableAttributeString("color2", this.m_oColor2.ToString());
			}

			this.m_oOffset2 && writer.WriteXmlNullableAttributeString("offset2", this.m_oOffset2.ToString());

			if (this.m_oOrigin) {
				if (0 !== this.m_oOrigin.GetX() || 0 !== this.m_oOrigin.GetY())
					writer.WriteXmlNullableAttributeString("origin", this.m_oOrigin.ToString());
			}

			this.m_oMatrix && writer.WriteXmlNullableAttributeString("matrix", this.m_oMatrix.ToString());
		};

		function CShapeType() {
			CVmlCommonElements.call(this);
			this.m_sAdj = null;
			this.m_oPath = null;
			this.m_oMaster = null;
		}

		IC(CShapeType, CVmlCommonElements, AscDFH.historyitem_type_VMLShapeType);
		CShapeType.prototype.readAttrXml = function (name, reader) {

			if ("adj" === name) {
				this.m_sAdj = reader.GetValue();
				return;
			}
			if ("path" === name) {
				this.m_oPath = reader.GetValue();
				return;
			}
			if ("master" === name) {
				this.m_oMaster = reader.GetValueBool();
				return;
			}
			CVmlCommonElements.prototype.readAttrXml.call(this, name, reader);
		};
		CShapeType.prototype.writeAttrXmlImpl = function (writer) {

			CVmlCommonElements.prototype.writeAttrXmlImpl.call(this, writer);
			if (this.m_sAdj !== null) {
				writer.WriteXmlNullableAttributeString("adj", this.m_sAdj);
			}

			if (this.m_oPath !== null)
				writer.WriteXmlNullableAttributeString("path", this.m_oPath.ToString());

			if (false !== this.m_oMaster)
				writer.WriteXmlNullableAttributeString("o:master", "true");
		};

		function CShape() {
			CVmlCommonElements.call(this);
			this.m_sType = null;
			this.m_sAdj = null;
			this.m_oPath = null;
			this.m_sGfxData = null;
			this.m_sEquationXML = null;
		}

		IC(CShape, CVmlCommonElements, AscDFH.historyitem_type_VMLShape);
		CShape.prototype.readAttrXml = function (name, reader) {
			if ("adj" === name) {
				this.m_sAdj = reader.GetValue();
				return;
			}
			if ("equationxml" === name) {
				this.m_sEquationXML = reader.GetValue();
				return;
			}
			if ("gfxdata" === name) {
				this.m_sGfxData = reader.GetValue();
				return;
			}
			if ("path" === name) {
				this.m_oPath = reader.GetValue();
				return;
			}
			if ("type" === name) {
				this.m_sType = reader.GetValue();
				return;
			}

			CVmlCommonElements.prototype.readAttrXml.call(this, name, reader);
		};
		CShape.prototype.writeAttrXmlImpl = function (writer) {

			CVmlCommonElements.prototype.writeAttrXmlImpl.call(this, writer);

			if (this.m_sType !== null) {
				writer.WriteXmlNullableAttributeString("type", this.m_sType);
			}

			if (this.m_sAdj !== null) {
				writer.WriteXmlNullableAttributeString("adj", this.m_sAdj);
			}

			if (this.m_oPath !== null)
				writer.WriteXmlNullableAttributeString("path", this.m_oPath.ToString());


			if (this.m_sEquationXML !== null) {
				writer.WriteXmlNullableAttributeString("equationxml", this.m_sEquationXML);
			}

		};


		function parseBool(val) {
			return "1" === val || "true" === val || "t" === val || "on" === val;
		}


		let EDropStyle =
			{
				valCombo: 0,
				valComboedit: 1,
				valSimple: 2
			};

		function getDropStyle(sValue) {
			if ("combo" === sValue) return EDropStyle.valCombo;
			else if ("comboedit" === sValue) return EDropStyle.valComboedit;
			else if ("simple" === sValue) return EDropStyle.valSimple;
			return EDropStyle.valSimple;
		}

		function getDropStyleStringValue(nType) {
			if (EDropStyle.valCombo === nType) return "combo";
			else if (EDropStyle.valComboedit === nType) return "comboedit";
			else if (EDropStyle.valSimple === nType) return "simple";
			return null;
		}


		let EChecked =
			{
				valUnchecked: 0,
				valChecked: 1,
				valMixed: 2
			};

		function getChecked(sValue) {
			if ("Mixed" === sValue) return EChecked.valMixed;
			else if ("Checked" === sValue) return EChecked.valChecked;
			else if ("Unchecked" === sValue) return EChecked.valUnchecked;
			else if ("1" === sValue) return EChecked.valChecked;
			else if ("0" === sValue) return EChecked.valUnchecked;
			return EChecked.valUnchecked;
		}

		function getCheckedStringValue(nType) {
			if (EChecked.valMixed === nType) return "Mixed";
			else if (EChecked.valChecked === nType) return "Checked";
			else if (EChecked.valUnchecked === nType) return "Unchecked";
			else if (EChecked.valChecked === nType) return "1";
			else if (EChecked.valUnchecked === nType) return "0";
			return null;
		}

		let ESelType =
			{

				valSingle: 0,
				valMulti: 1,
				valExtended: 2
			};

		function getSelType(sValue) {
			if ("extended" === sValue) return ESelType.valSingle;
			else if ("multi" === sValue) return ESelType.valMulti;
			else if ("single" === sValue) return ESelType.valExtended;
			return ESelType.valSingle;
		}

		function getSelTypeStringValue(nType) {
			if (ESelType.valSingle === nType) return "extended";
			else if (ESelType.valMulti === nType) return "multi";
			else if (ESelType.valExtended === nType) return "single";
			return null;
		}

		let EHorizontalAlignment =
			{
				horizontalalignmentCenter: 0,
				horizontalalignmentContinuous: 1,
				horizontalalignmentDistributed: 2,
				horizontalalignmentFill: 3,
				horizontalalignmentGeneral: 4,
				horizontalalignmentJustify: 5,
				horizontalalignmentLeft: 6,
				horizontalalignmentRight: 7,
				horizontalalignmentCenterContinuous: 8
			};

		function getHorizontalAlignment(sValue) {
			if ("center" === sValue || "Center" === sValue)
				return EHorizontalAlignment.horizontalalignmentCenter;
			else if ("centerContinuous" === sValue || "CenterAcrossSelection" === sValue)
				return EHorizontalAlignment.horizontalalignmentCenterContinuous;
			else if ("continuous" === sValue)
				return EHorizontalAlignment.horizontalalignmentContinuous;
			else if ("distributed" === sValue || "Distributed" === sValue)
				return EHorizontalAlignment.horizontalalignmentDistributed;
			else if ("fill" === sValue || "Fill" === sValue)
				return EHorizontalAlignment.horizontalalignmentFill;
			else if ("general" === sValue || "Automatic" === sValue)
				return EHorizontalAlignment.horizontalalignmentGeneral;
			else if ("justify" === sValue || "Justify" === sValue)
				return EHorizontalAlignment.horizontalalignmentJustify;
			else if ("left" === sValue || "Left" === sValue)
				return EHorizontalAlignment.horizontalalignmentLeft;
			else if ("right" === sValue || "Right" === sValue)
				return EHorizontalAlignment.horizontalalignmentRight;
			else
				return EHorizontalAlignment.horizontalalignmentGeneral;
		}


		function getHorizontalAlignmentStringValue(nType) {
			if (EHorizontalAlignment.horizontalalignmentCenter === nType)
				return "center";
			else if (EHorizontalAlignment.horizontalalignmentCenterContinuous === nType)
				return "centerContinuous";
			else if (EHorizontalAlignment.horizontalalignmentContinuous === nType)
				return "continuous";
			else if (EHorizontalAlignment.horizontalalignmentDistributed === nType)
				return "distributed";
			else if (EHorizontalAlignment.horizontalalignmentFill === nType)
				return "fill";
			else if (EHorizontalAlignment.horizontalalignmentGeneral === nType)
				return "general";
			else if (EHorizontalAlignment.horizontalalignmentJustify === nType)
				return "justify";
			else if (EHorizontalAlignment.horizontalalignmentLeft === nType)
				return "left";
			else if (EHorizontalAlignment.horizontalalignmentRight === nType)
				return "right";
			else
				return null;
		}


		let EVerticalAlignment =
			{
				verticalalignmentBottom: 0,
				verticalalignmentCenter: 1,
				verticalalignmentDistributed: 2,
				verticalalignmentJustify: 3,
				verticalalignmentTop: 4
			};

		function getVerticalAlignment(sValue) {
			if ("bottom" === sValue || "Bottom" === sValue)
				return EVerticalAlignment.verticalalignmentBottom;
			else if ("center" === sValue || "Center" === sValue)
				return EVerticalAlignment.verticalalignmentCenter;
			else if ("distributed" === sValue || "Distributed" === sValue)
				return EVerticalAlignment.verticalalignmentDistributed;
			else if ("justify" === sValue || "Justify" === sValue)
				return EVerticalAlignment.verticalalignmentJustify;
			else if ("top" === sValue || "Top" === sValue)
				return EVerticalAlignment.verticalalignmentTop;
			else
				return EVerticalAlignment.verticalalignmentBottom;
		}


		function getVerticalAlignmentStringValue(nType) {
			if (EVerticalAlignment.verticalalignmentBottom === nType)
				return "bottom";
			else if (EVerticalAlignment.verticalalignmentCenter === nType)
				return "center";
			else if (EVerticalAlignment.verticalalignmentDistributed === nType)
				return "distributed";
			else if (EVerticalAlignment.verticalalignmentJustify === nType)
				return "justify";
			else if (EVerticalAlignment.verticalalignmentTop === nType)
				return "top";
			else
				return null;
		}

		function CClientData() {
			CBaseNoId.call(this);
			this.m_oObjectType = null;

			this.m_oMoveWithCells = null;
			this.m_oSizeWithCells = null;
			this.m_oAnchor = null;
			this.m_oRow = null;
			this.m_oColumn = null;
			this.m_oMin = null;
			this.m_oMax = null;
			this.m_oInc = null;
			this.m_oDx = null;
			this.m_oPage = null;
			this.m_oDropLines = null;
			this.m_oSel = null;
			this.m_oWidthMin = null;
			this.m_oDropStyle = null;
			this.m_oFirstButton = null;
			this.m_oDefaultSize = null;
			this.m_oAutoFill = null;
			this.m_oAutoScale = null;
			this.m_oAutoLine = null;
			this.m_oHoriz = null;
			this.m_oVScroll = null;
			this.m_oAutoPict = null;
			this.m_oColored = null;
			this.m_oMultiLine = null;
			this.m_oNoThreeD = null;
			this.m_oNoThreeD2 = null;
			this.m_oLockText = null;
			this.m_oJustLastX = null;
			this.m_oSecretEdit = null;
			this.m_oFmlaLink = null;
			this.m_oFmlaRange = null;
			this.m_oFmlaMacro = null;
			this.m_oFmlaTxbx = null;
			this.m_oFmlaGroup = null;
			this.m_oCf = null;
			this.m_oChecked = null;
			this.m_oMultiSel = null;
			this.m_oSelType = null;
			this.m_oVal = null;
			this.m_oTextHAlign = null;
			this.m_oTextVAlign = null;
			this.m_oVisible = null;
		}

		IC(CClientData, CBaseNoId, AscDFH.historyitem_type_VMLClientData);
		CClientData.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "ObjectType" :
					this.m_oObjectType = readClientDataObjectType(reader);
					break;
			}
		};
		CClientData.prototype.readChildXml = function (name, reader) {
			let sContent = reader.GetTextDecodeXml();
			if ("MoveWithCells" === name) this.m_oMoveWithCells = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("SizeWithCells" === name) this.m_oSizeWithCells = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("Anchor" === name) this.m_oAnchor = sContent;
			else if ("Row" === name) this.m_oRow = parseInt(sContent);
			else if ("Column" === name) this.m_oColumn = parseInt(sContent);
			else if ("DefaultSize" === name) this.m_oDefaultSize = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("AutoLine" === name) this.m_oAutoLine = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("AutoFill" === name) this.m_oAutoFill = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("AutoPict" === name) this.m_oAutoPict = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("AutoScale" === name) this.m_oAutoScale = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("FmlaLink" === name) this.m_oFmlaLink = sContent;
			else if ("FmlaRange" === name) this.m_oFmlaRange = sContent;
			else if ("FmlaMacro" === name) this.m_oFmlaMacro = sContent;
			else if ("FmlaTxbx" === name) this.m_oFmlaTxbx = sContent;
			else if ("FmlaGroup" === name) this.m_oFmlaGroup = sContent;
			else if ("CF" === name) this.m_oCf = sContent;
			else if ("Checked" === name) this.m_oChecked = getChecked(sContent);
			else if ("Min" === name) this.m_oMin = parseInt(sContent);
			else if ("Max" === name) this.m_oMax = parseInt(sContent);
			else if ("Val" === name) this.m_oVal = sContent;
			else if ("Inc" === name) this.m_oInc = parseInt(sContent);
			else if ("Sel" === name) this.m_oSel = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("WidthMin" === name) this.m_oWidthMin = parseInt(sContent);
			else if ("Dx" === name) this.m_oDx = parseInt(sContent);
			else if ("Page" === name) this.m_oPage = parseInt(sContent);
			else if ("DropLines" === name) this.m_oDropLines = parseInt(sContent);
			else if ("NoThreeD2" === name) this.m_oNoThreeD2 = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("NoThreeD" === name) this.m_oNoThreeD = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("DropStyle" === name) this.m_oDropStyle = getDropStyle(sContent);
			else if ("FirstButton" === name) this.m_oFirstButton = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("VScroll" === name) this.m_oVScroll = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("Horiz" === name) this.m_oHoriz = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("TextHAlign" === name) this.m_oTextHAlign = getHorizontalAlignment(sContent);
			else if ("TextVAlign" === name) this.m_oTextVAlign = getVerticalAlignment(sContent);
			else if ("Colored" === name) this.m_oColored = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("MultiLine" === name) this.m_oMultiLine = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("LockText" === name) this.m_oLockText = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("JustLastX" === name) this.m_oJustLastX = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("SecretEdit" === name) this.m_oSecretEdit = parseBool(sContent.length === 0 ? "t" : sContent);
			else if ("SelType" === name) this.m_oSelType = getSelType(sContent);
			else if ("Visible" === name) this.m_oVisible = parseBool(sContent.length === 0 ? "t" : sContent);
		};
		CClientData.prototype.writeAttrXmlImpl = function (writer) {
			if (this.m_oObjectType !== null) {
				writer.WriteXmlNullableAttributeString(" ObjectType", this.m_oObjectType.ToString());
			}
		};
		CClientData.prototype.writeChildren = function (writer) {
			if (this.m_oMoveWithCells !== null && this.m_oMoveWithCells)
				writer.WriteXmlValueString("x:MoveWithCells", "");
			if (this.m_oSizeWithCells !== null && this.m_oSizeWithCells)
				writer.WriteXmlValueString("x:SizeWithCells", "");
			if (this.m_oAnchor !== null) {
				writer.WriteXmlValueString("x:Anchor", this.m_oAnchor);
			}
			if (this.m_oDefaultSize !== null) {

				writer.WriteXmlValueString("x:DefaultSize", (this.m_oDefaultSize ? "True" : "False"));
			}
			if (this.m_oAutoLine !== null) {
				writer.WriteXmlValueString("x:AutoLine", (this.m_oAutoLine ? "True" : "False"));
			}
			if (this.m_oAutoScale !== null) {
				writer.WriteXmlValueString("x:AutoScale", (this.m_oAutoScale ? "True" : "False"));
			}
			if (this.m_oAutoPict !== null) {
				writer.WriteXmlValueString("x:AutoPict", (this.m_oAutoPict ? "True" : "False"));
			}
			if (this.m_oAutoFill !== null) {
				writer.WriteXmlValueString("x:AutoFill", (this.m_oAutoFill ? "True" : "False"));
			}
			if (this.m_oFirstButton !== null) {
				writer.WriteXmlValueString("x:FirstButton", (this.m_oFirstButton ? "True" : "False"));
			}
			if (this.m_oVScroll !== null) {
				writer.WriteXmlValueString("x:VScroll", (this.m_oVScroll ? "True" : "False"));
			}
			if (this.m_oHoriz !== null) {
				writer.WriteXmlValueString("x:Horiz", (this.m_oHoriz ? "True" : "False"));
			}
			if (this.m_oColored !== null) {
				writer.WriteXmlValueString("x:Colored", (this.m_oColored ? "True" : "False"));
			}
			if (this.m_oJustLastX !== null) {
				writer.WriteXmlValueString("x:JustLastX", (this.m_oJustLastX ? "True" : "False"));
			}
			if (this.m_oMultiLine !== null) {
				writer.WriteXmlValueString("x:MultiLine", (this.m_oMultiLine ? "True" : "False"));
			}
			if (this.m_oTextHAlign !== null) {
				writer.WriteXmlNullableValueString("x:TextHAlign", getHorizontalAlignmentStringValue(this.m_oTextHAlign));
			}
			if (this.m_oTextVAlign !== null) {
				writer.WriteXmlNullableValueString("x:TextVAlign", getVerticalAlignmentStringValue(this.m_oTextVAlign));
			}
			if (this.m_oLockText !== null) {
				writer.WriteXmlValueString("x:LockText", (this.m_oLockText ? "True" : "False"));
			}
			if (this.m_oRow !== null) {
				writer.WriteXmlValueNumber("x:Row", (this.m_oRow));
			}
			if (this.m_oColumn !== null) {
				writer.WriteXmlValueNumber("x:Column", (this.m_oColumn));
			}
			if (this.m_oVal !== null) {
				writer.WriteXmlValueString("x:Val", (this.m_oVal));
			}
			if (this.m_oCf !== null) {
				writer.WriteXmlValueString("x:CF", (this.m_oCf));
			}
			// if(this.m_oMultiSel !== null)
			// {
			// 	"<x:MultiSel>" + (this.m_oMultiSel) + "</x:MultiSel>";
			// 	writer.WriteXmlValueString("x:MultiSel", (this.m_oMultiSel));
			// }
			if (this.m_oChecked !== null) {
				writer.WriteXmlNullableValueString("x:Checked", getCheckedStringValue(this.m_oChecked));
			}
			if (this.m_oFmlaLink !== null) {
				writer.WriteXmlValueString("x:FmlaLink", (this.m_oFmlaLink));
			}
			if (this.m_oFmlaRange !== null) {
				writer.WriteXmlValueString("x:FmlaRange", (this.m_oFmlaRange));
			}
			if (this.m_oFmlaMacro !== null) {
				writer.WriteXmlValueString("x:FmlaMacro", (this.m_oFmlaMacro));
			}
			if (this.m_oFmlaTxbx !== null) {
				writer.WriteXmlValueString("x:FmlaTxbx", (this.m_oFmlaTxbx));
			}
			if (this.m_oFmlaGroup !== null) {
				writer.WriteXmlValueString("x:FmlaGroup", (this.m_oFmlaGroup));
			}
			if (this.m_oMin !== null) {
				writer.WriteXmlValueNumber("x:Min", (this.m_oMin));
			}
			if (this.m_oMax !== null) {
				writer.WriteXmlValueNumber("x:Max", (this.m_oMax));
			}
			if (this.m_oInc !== null) {
				writer.WriteXmlValueNumber("x:Inc", (this.m_oInc));
			}
			if (this.m_oSel !== null) {
				writer.WriteXmlValueString("x:Sel", getBooleanTrueFalse(this.m_oSel));
			}
			if (this.m_oSelType !== null) {
				writer.WriteXmlNullableValueString("x:SelType", getSelTypeStringValue(this.m_oSelType));
			}
			if (this.m_oDx !== null) {
				writer.WriteXmlValueNumber("x:Dx", (this.m_oDx));
			}
			if (this.m_oDropStyle !== null) {
				writer.WriteXmlNullableValueString("x:DropStyle", getDropStyleStringValue(this.m_oDropStyle));
			}
			if (this.m_oDropLines !== null) {
				writer.WriteXmlValueNumber("x:DropLines", (this.m_oDropLines));
			}
			if (this.m_oPage !== null) {
				writer.WriteXmlValueNumber("x:Page", (this.m_oPage));
			}
			if (this.m_oWidthMin !== null) {
				writer.WriteXmlValueNumber("x:WidthMin", (this.m_oWidthMin));
			}
			if (this.m_oNoThreeD !== null) {
				writer.WriteXmlValueString("x:NoThreeD", (this.m_oNoThreeD ? "True" : "False"));
			}
			if (this.m_oNoThreeD2 !== null) {
				writer.WriteXmlValueString("x:NoThreeD2", (this.m_oNoThreeD2 ? "True" : "False"));
			}
			if (this.m_oSecretEdit !== null) {
				writer.WriteXmlValueString("x:SecretEdit", (this.m_oSecretEdit ? "True" : "False"));
			}
		};

		CClientData.prototype.getAnchorArray = function (aAnchor) {
			aAnchor.length = 0;
			if (this.m_oAnchor) {
				let arSplit = this.m_oAnchor.split(",");
				for (let i = 0; i < arSplit.length; i++) {
					aAnchor.push(parseInt(arSplit[i]));
				}
			}
		};


		CClientData.prototype.toCellAnchor = function () {
			//TODO: implement
		};
		CClientData.prototype.toFormControlPr = function () {
			//TODO: implement
		};


		function CStroke() {
			CBaseNoId.call(this);

			this.m_oId = null;
			this.m_sAltHref = null;
			this.m_oColor = null;
			this.m_oColor2 = null;
			this.m_oDahsStyle = null;
			this.m_oEndArrow = null;
			this.m_oEndArrowLength = null;
			this.m_oEndArrowWidth = null;
			this.m_oEndCap = null;
			this.m_oFillType = null;
			this.m_oForceDash = null;
			this.m_sHref = null;
			this.m_rId = null;
			this.m_oImageAlignShape = null;
			this.m_oImageAspect = null;
			this.m_oImageSize = null;
			this.m_oInsetPen = null;
			this.m_oJoinStyle = null;
			this.m_oLineStyle = null;
			this.m_oMiterLimit = null;
			this.m_oOn = null;
			this.m_oOpacity = null;
			this.m_oRelId = null;
			this.m_sSrc = null;
			this.m_oStartArrow = null;
			this.m_oStartArrowLength = null;
			this.m_oStartArrowWidth = null;
			this.m_sTitle = null;
			this.m_oWeight = null;

			// Childs
			this.m_oLeft = null;
			this.m_oTop = null;
			this.m_oRight = null;
			this.m_oBottom = null;
			this.m_oColumn = null;
		}

		IC(CStroke, CBaseNoId, 0);
		CStroke.prototype.readAttrXml = function (name, reader) {
			if ("color" === name) this.m_oColor = readColorType(reader);
			else if ("color2" === name) this.m_oColor2 = readColorType(reader);
			else if ("dashstyle" === name) this.m_oDahsStyle = readDashStyle(reader);
			else if ("endarrow" === name) this.m_oEndArrow = readArrowType(reader);
			else if ("endarrowlength" === name) this.m_oEndArrowLength = readArrowLength(reader);
			else if ("endarrowwidth" === name) this.m_oEndArrowWidth = readArrowWidth(reader);
			else if ("endcap" === name) this.m_oEndCap = readEndCap(reader);
			else if ("filltype" === name) this.m_oFillType = readFillType(reader);
			else if ("id" === name) {
				let sName = reader.GetName();
				if (sName === "id") {
					this.m_oId = reader.GetValue();
				} else if (sName === "r:id" || sName === "relationships:id") {
					this.m_rId = reader.GetValue();
				}
			} else if ("imagealignshape" === name) this.m_oImageAlignShape = reader.GetValueBool();
			else if ("imageaspect" === name) this.m_oImageAspect = readImageAspect(reader);
			else if ("imagesize" === name) this.m_oImageSize = new CVml_Vector2D_Units_Or_Percentage(reader.GetValue());
			else if ("insetpen" === name) this.m_oInsetPen = reader.GetValueBool();
			else if ("joinstyle" === name) this.m_oJoinStyle = readJoinStyle(reader);
			else if ("linestyle" === name) this.m_oLineStyle = readLineStyle(reader);
			else if ("miterlimit" === name) this.m_oMiterLimit = reader.GetValueInt();
			else if ("althref" === name) this.m_sAltHref = reader.GetValue();
			else if ("forcedash" === name) this.m_oForceDash = reader.GetValueBool();
			else if ("href" === name) this.m_sHref = reader.GetValue();
			else if ("on" === name) this.m_oOn = reader.GetValueBool();
			else if ("opacity" === name) this.m_oOpacity = reader.GetValueDouble();
			else if ("relid" === name) this.m_oRelId = reader.GetValue();
			else if ("title" === name) this.m_sTitle = reader.GetValue();
			else if ("src" === name) this.m_sSrc = reader.GetValue();
			else if ("startarrow" === name) this.m_oStartArrow = readArrowType(reader);
			else if ("startarrowlength" === name) this.m_oStartArrowLength = readArrowLength(reader);
			else if ("startarrowwidth" === name) this.m_oStartArrowWidth = readArrowWidth(reader);
			else if ("weight" === name) this.m_oWeight = reader.GetValueDouble();
		};
		CStroke.prototype.readChildXml = function (name, reader) {
			if ("left" === name)
				this.m_oLeft = this.readChildSide(reader);
			else if ("top" === name)
				this.m_oTop = this.readChildSide(reader);
			else if ("right" === name)
				this.m_oRight = this.readChildSide(reader);
			else if ("bottom" === name)
				this.m_oBottom = this.readChildSide(reader);
			else if ("column" === name)
				this.m_oColumn = this.readChildSide(reader);
		};
		CStroke.prototype.readChildSide = function (reader) {

			let oSide = new CStrokeChild();
			oSide.fromXml(reader);
			return oSide;
		};
		CStroke.prototype.writeAttrXmlImpl = function (writer) {


			writer.WriteXmlNullableAttributeString("id", this.m_oId);
			if (true !== this.m_oOn)
				writer.WriteXmlNullableAttributeString("on", "false");

			if (1 !== this.m_oWeight)
				writer.WriteXmlNullableAttributeDouble("weight", this.m_oWeight);

			writer.WriteXmlNullableAttributeString("color", getColorType(this.m_oColor));

			if (this.m_oOpacity !== null)
				writer.WriteXmlNullableAttributeDouble("opacity", this.m_oOpacity);

			writer.WriteXmlNullableAttributeString("linestyle", getLineStyle(this.m_oLineStyle));

			if (8 !== this.m_oMiterLimit)
				writer.WriteXmlNullableAttributeInt("miterlimit", this.m_oMiterLimit);

			if (EStrokeJoinStyle.strokejoinstyleRound !== this.m_oJoinStyle)
				writer.WriteXmlNullableAttributeString("joinstyle", getJoinStyle(this.m_oJoinStyle));

			if (EStrokeEndCap.strokeendcapFlat !== this.m_oEndCap)
				writer.WriteXmlNullableAttributeString("endcap", getEndCap(this.m_oEndCap));

			if (EVmlDashStyle.vmldashstyleSolid !== this.m_oDahsStyle)
				writer.WriteXmlNullableAttributeString("dashstyle", getDashStyle(this.m_oDahsStyle));

			if (EFillType.filltypeSolid !== this.m_oFillType)
				writer.WriteXmlNullableAttributeString("filltype", getFillType(this.m_oFillType));

			writer.WriteXmlNullableAttributeString("src", this.m_sSrc);

			if (EImageAspect.imageaspectIgnore !== this.m_oImageAspect)
				writer.WriteXmlNullableAttributeString("imageaspect", getImageAspect(this.m_oImageAspect));

			this.m_oImageSize && writer.WriteXmlNullableAttributeString("imagesize", this.m_oImageSize.ToString());

			if (true !== this.m_oOn)
				writer.WriteXmlNullableAttributeString("imagealignshape", "false");

			writer.WriteXmlNullableAttributeString("color2", getColorType(this.m_oColor2));

			if (EStrokeArrowType.strokearrowtypeNone !== this.m_oStartArrow)
				writer.WriteXmlNullableAttributeString("startarrow", getArrowType(this.m_oStartArrow));

			if (EStrokeArrowWidth.strokearrowwidthMedium !== this.m_oStartArrowWidth)
				writer.WriteXmlNullableAttributeString("startarrowwidth", getArrowWidth(this.m_oStartArrowWidth));

			if (EStrokeArrowLength.strokearrowlengthMedium !== this.m_oStartArrowLength)
				writer.WriteXmlNullableAttributeString("startarrowlength", getArrowLength(this.m_oStartArrowLength));

			if (EStrokeArrowType.strokearrowtypeNone !== this.m_oEndArrow)
				writer.WriteXmlNullableAttributeString("endarrow", getArrowType(this.m_oEndArrow));

			if (EStrokeArrowWidth.strokearrowwidthMedium !== this.m_oEndArrowWidth)
				writer.WriteXmlNullableAttributeString("endarrowwidth", getArrowWidth(this.m_oEndArrowWidth));

			if (EStrokeArrowLength.strokearrowlengthMedium !== this.m_oEndArrowLength)
				writer.WriteXmlNullableAttributeString("endarrowlength", getArrowLength(this.m_oEndArrowLength));

			writer.WriteXmlNullableAttributeString("o:href", this.m_sHref);
			writer.WriteXmlNullableAttributeString("o:althref", this.m_sAltHref);
			writer.WriteXmlNullableAttributeString("o:title", this.m_sTitle);

			if (false !== this.m_oForceDash)
				writer.WriteXmlNullableAttributeString("o:forcedash", "true");

			writer.WriteXmlNullableAttributeString("r:id", this.m_rId);
			writer.WriteXmlNullableAttributeString("insetpen", getBooleanTrueFalse(this.m_oInsetPen));
			writer.WriteXmlNullableAttributeString("o:relid", this.m_oRelId);
		};
		CStroke.prototype.writeChildren = function (writer) {
			if (this.m_oLeft !== null)
				this.m_oLeft.toXml(writer, "o:left");

			if (this.m_oTop !== null)
				this.m_oTop.toXml(writer, "o:top");

			if (this.m_oRight !== null)
				this.m_oRight.toXml(writer, "o:right");

			if (this.m_oBottom !== null)
				this.m_oBottom.toXml(writer, "o:bottom");

			if (this.m_oColumn !== null)
				this.m_oColumn.toXml(writer, "o:column");

		};

		function CTextbox() {
			CBaseNoId.call(this);
			this.m_oId = null;
			this.m_oStyle = null;
			this.m_oInset = null;
			this.m_oSingleClick = null;
			this.m_oInsetMode = null;

			this.m_oTxtbxContent = null;
			this.m_oText = null;
		}

		IC(CTextbox, CBaseNoId, 0);
		CTextbox.prototype.readAttrXml = function (name, reader) {
			if ("id" === name) this.m_oId = reader.GetValue();
			else if ("inset" === name) this.m_oInset = new CVml_TextBoxInset(reader.GetValue());
			else if ("insetmode" === name) this.m_oInsetMode = readInsetMode(reader);
			else if ("singleclick" === name) this.m_oSingleClick = reader.GetValueBool();
			else if ("style" === name) this.m_oStyle = new CCssStyle(reader.GetValue());
		};
		CTextbox.prototype.readChildXml = function (name, reader) {
			if ("txbxContent" === name) {
				this.m_oTxtbxContent = AscFormat.fReadTxBoxContentXML(reader, this)
			} else if ("div" === name) {
				this.m_oText = reader.GetTextDecodeXml();
			}
		};
		CTextbox.prototype.writeAttrXmlImpl = function (writer) {

			writer.WriteXmlNullableAttributeString("id", this.m_oId);
			this.m_oStyle && writer.WriteXmlNullableAttributeString("style", this.m_oStyle.ToString());

			if (this.m_oInset !== null) {
				writer.WriteXmlNullableAttributeString("inset", this.m_oInset.ToString());
			}

			if (false !== this.m_oSingleClick)
				writer.WriteXmlNullableAttributeString("o:singleclick", "true");

			if (EInsetMode.insetmodeCustom !== this.m_oInsetMode)
				writer.WriteXmlNullableAttributeString("o:insetmode", getInsetMode(this.m_oInsetMode));
		};
		CTextbox.prototype.writeChildren = function (writer) {

			if (this.m_oTxtbxContent !== null)
				this.m_oTxtbxContent.toXml(writer, "txbxContent");

			if (this.m_oText !== null) {
				writer.WriteXmlValueString("div", this.m_oText);
			}

		};

		function CTextPath() {
			CBaseNoId.call(this);
			this.m_oFitPath = null;
			this.m_oFitShape = null;
			this.m_oId = null;
			this.m_oOn = null;
			this.m_sString = null;
			this.m_oStyle = null;
			this.m_oTrim = null;
			this.m_oXScale = null;
			this.m_sStringOriginal = null;
		}

		IC(CTextPath, CBaseNoId, 0);
		CTextPath.prototype.readAttrXml = function (name, reader) {
			let wsChar = name.charCodeAt(0);
			switch (wsChar) {
				case 'f':
					if ("fitpath" === name) this.m_oFitPath = reader.GetValueBool();
					else if ("fitshape" === name) this.m_oFitShape = reader.GetValueBool();
					break;
				case 'i':
					if ("id" === name) this.m_oId = reader.GetValue();
					break;
				case 'o':
					if ("on" === name) this.m_oOn = reader.GetValueBool();
					break;
				case 's':
					if ("string" === name) this.m_sString = reader.GetValue();
					else if ("style" === name) this.m_oStyle = new CCssStyle(reader.GetValue());
					break;
				case 't':
					if ("trim" === name) this.m_oTrim = reader.GetValueBool();
					break;
				case 'x':
					if ("xscale" === name) this.m_oXScale = reader.GetValueBool();
					break;
			}
		};
		CTextPath.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("id", this.m_oId);
			this.m_oStyle && writer.WriteXmlNullableAttributeString("style", this.m_oStyle.ToString());

			if ((this.m_oOn !== null) && (false !== this.m_oOn))
				writer.WriteXmlNullableAttributeString("on", "true");

			if ((this.m_oFitShape !== null) && (false !== this.m_oFitShape))
				writer.WriteXmlNullableAttributeString("fitshape", "true");

			if ((this.m_oFitPath !== null) && (false !== this.m_oFitPath))
				writer.WriteXmlNullableAttributeString("fitpath", "true");

			if ((this.m_oTrim !== null) && (false !== this.m_oTrim))
				writer.WriteXmlNullableAttributeString("trim", "true");

			if ((this.m_oXScale !== null) && (false !== this.m_oXScale))
				writer.WriteXmlNullableAttributeString("xscale", "true");

			writer.WriteXmlNullableAttributeString("string", this.m_sString);
		};

		function CGroup() {
			CVmlCommonElements.call(this);
			this.m_oEditAs = null;
			this.m_oTableLimits = null;
			this.m_oTableProperties = null;
		}

		IC(CGroup, CVmlCommonElements, AscDFH.historyitem_type_VMLGroup);

		CGroup.prototype.readAttrXml = function (name, reader) {
			if ("editas" === name) {
				this.m_oEditAs = readEditAs(reader);
				return;
			} else if ("tableproperties" === name) {
				this.m_oTableProperties = new CVml_TableProperties(reader.GetValue());
				return;
			} else if ("tablelimits" === name) {
				this.m_oTableLimits = new CVml_TableLimits(reader.GetValue());
				return;
			}
			CVmlCommonElements.prototype.readAttrXml.call(this, name, reader);
		};
		CGroup.prototype.readChildXml = function (name, reader) {
			let oItem = null;
			if ("callout" === name)
				oItem = new CCallout();
			else if ("clippath" === name)
				oItem = new CClipPath();
			else if ("diagram" === name)
				oItem = new CDiagram();
			else if ("extrusion" === name)
				oItem = new CExtrusion();
			else if ("lock" === name)
				oItem = new CLock();
			else if ("signatureline" === name)
				oItem = new CSignatureLine();
			else if ("skew" === name)
				oItem = new CSkew();
			else if ("arc" === name)
				oItem = new CArc();
			else if ("curve" === name)
				oItem = new CCurve();
			else if ("fill" === name)
				oItem = new CFillVml();
			else if ("formulas" === name)
				oItem = new CFormulas();
			else if ("group" === name)
				oItem = new CGroup();
			else if ("handles" === name)
				oItem = new CHandles();
			else if ("imagedata" === name)
				oItem = new CImageData();
			else if ("image" === name)
				oItem = new CImage();
			else if ("line" === name)
				oItem = new CLine();
			else if ("oval" === name)
				oItem = new COval();
			else if ("path" === name)
				oItem = new CPath();
			else if ("polyline" === name)
				oItem = new CPolyLine();
			else if ("rect" === name)
				oItem = new CRect();
			else if ("roundrect" === name)
				oItem = new CRoundRect();
			else if ("shadow" === name)
				oItem = new CShadow();
			else if ("shape" === name)
				oItem = new CShape();
			else if ("shapetype" === name)
				oItem = new CShapeType();
			else if ("stroke" === name)
				oItem = new CStroke();
			else if ("textbox" === name)
				oItem = new CTextbox();
			else if ("textpath" === name)
				oItem = new CTextPath();
			else if ("anchorLock" === name)
				oItem = new CAnchorLock();
			else if ("borderbottom" === name)
				oItem = new CBorder("borderbottom");
			else if ("borderleft" === name)
				oItem = new CBorder("borderleft");
			else if ("borderright" === name)
				oItem = new CBorder("borderright");
			else if ("bordertop" === name)
				oItem = new CBorder("bordertop");
			else if ("wrap" === name)
				oItem = new CWrap();
			else if ("ClientData" === name)
				oItem = new CClientData();

			if (oItem) {
				oItem.fromXml(reader);
				this.items.push(oItem);
			}
		};
		CGroup.prototype.writeAttrXmlImpl = function (writer) {
			CVmlCommonElements.prototype.writeAttrXmlImpl.call(this, writer);
			writer.WriteXmlNullableAttributeString("editas", getEditAs(this.m_oEditAs));
			if (this.m_oTableProperties)
				writer.WriteXmlNullableAttributeString("o:tableproperties", this.m_oTableProperties.ToString());
			this.m_oTableLimits && writer.WriteXmlNullableAttributeString("o:tablelimits", this.m_oTableLimits.ToString());
		};
		CGroup.prototype.writeChildren = function (writer) {

			for (let i = 0; i < this.items.length; ++i) {
				let oItem = this.items[i];
				if (oItem) {
					let sName = null;
					if (oItem instanceof CCallout)
						sName = ("o:callout");
					else if (oItem instanceof CClipPath)
						sName = ("o:clippath");
					else if (oItem instanceof CDiagram)
						sName = ("o:diagram");
					else if (oItem instanceof CExtrusion)
						sName = ("o:extrusion");
					else if (oItem instanceof CLock)
						sName = ("o:lock");
					else if (oItem instanceof CSignatureLine)
						sName = ("o:signatureline");
					else if (oItem instanceof CSkew)
						sName = ("o:skew");
					else if (oItem instanceof CArc)
						sName = ("v:arc");
					else if (oItem instanceof CCurve)
						sName = ("v:curve");
					else if (oItem instanceof CFillVml)
						sName = ("v:fill");
					else if (oItem instanceof CFormulas)
						sName = ("v:formulas");
					else if (oItem instanceof CGroup)
						sName = ("v:group");
					else if (oItem instanceof CHandles)
						sName = ("v:handles");
					else if (oItem instanceof CImageData)
						sName = ("v:imagedata");
					else if (oItem instanceof CImage)
						sName = ("v:image");
					else if (oItem instanceof CLine)
						sName = ("v:line");
					else if (oItem instanceof COval)
						sName = ("v:oval");
					else if (oItem instanceof CPath)
						sName = ("v:path");
					else if (oItem instanceof CPolyLine)
						sName = ("v:polyline");
					else if (oItem instanceof CRect)
						sName = ("v:rect");
					else if (oItem instanceof CRoundRect)
						sName = ("v:roundrect");
					else if (oItem instanceof CShadow)
						sName = ("v:shadow");
					else if (oItem instanceof CShape)
						sName = ("v:shape");
					else if (oItem instanceof CShapeType)
						sName = ("v:shapetype");
					else if (oItem instanceof CStroke)
						sName = ("v:stroke");
					else if (oItem instanceof CTextbox)
						sName = ("v:textbox");
					else if (oItem instanceof CTextPath)
						sName = ("v:textpath");
					else if (oItem instanceof CAnchorLock)
						sName = ("wd:anchorLock");
					else if (oItem instanceof CBorder)
						sName = "wd:" + oItem.m_sType;
					else if (oItem instanceof CWrap)
						sName = ("wd:wrap");
					else if (oItem instanceof CClientData)
						sName = ("x:ClientData");
					if (sName) {
						oItem.toXml(writer, sName);
					}
				}
			}
		};


		function CShapeDefaults() {
			CBaseNoId.call(this);
			this.m_oAllowInCell = null;
			this.m_oExt = null;
			this.m_oFill = null;
			this.m_oFillColor = null;
			this.m_oSpIdMax = null;
			this.m_oStroke = null;
			this.m_oStrokeColor = null;
			this.m_oStyle = null;

			this.m_oVmlFill = null;
			this.m_oVmlStroke = null;
			this.m_oVmlTextbox = null;
			this.m_oVmlShadow = null;

			this.m_oSkew = null;
			this.m_oExtrusion = null;
			this.m_oCallout = null;
			this.m_oLock = null;
			this.m_oColorMru = null;
			this.m_oColorMenu = null;
		}

		IC(CShapeDefaults, CBaseNoId, 0);
		CShapeDefaults.prototype.readAttrXml = function (name, reader) {
			if ("fill" === name) this.m_oFill = reader.GetValueBool();
			else if ("fillcolor" === name) this.m_oFillColor = readColorType(reader);
			else if ("allowincell" === name) this.m_oAllowInCell = reader.GetValueBool();
			else if ("spidmax" === name) this.m_oSpIdMax = reader.GetValueInt();
			else if ("style" === name) this.m_oStyle = new CCssStyle(reader.GetValue());
			else if ("stroke" === name) this.m_oStroke = reader.GetValueBool();
			else if ("strokecolor" === name) this.m_oStrokeColor = readColorType(reader);
			else if ("ext" === name) this.m_oExt = readExt(reader);
		};
		CShapeDefaults.prototype.readChildXml = function (name, reader) {
			if ("fill" === name) {

				this.m_oVmlFill = new CFillVml();
				this.m_oVmlFill.fromXml(reader);
			} else if ("stroke" === name) {

				this.m_oVmlStroke = new CStroke();
				this.m_oVmlStroke.fromXml(reader);
			} else if ("textbox" === name) {

				this.m_oVmlTextbox = new CTextbox();
				this.m_oVmlTextbox.fromXml(reader);
			} else if ("shadow" === name) {

				this.m_oVmlShadow = new CShadow();
				this.m_oVmlShadow.fromXml(reader);
			} else if ("skew" === name) {

				this.m_oSkew = new CSkew();
				this.m_oSkew.fromXml(reader);
			} else if ("extrusion" === name) {

				this.m_oExtrusion = new CExtrusion();
				this.m_oExtrusion.fromXml(reader);
			} else if ("callout" === name) {

				this.m_oCallout = new CCallout();
				this.m_oCallout.fromXml(reader);
			} else if ("lock" === name) {
				this.m_oLock = new CLock();
				this.m_oLock.fromXml(reader);
			} else if ("colormru" === name) {

				this.m_oColorMru = new CColorMru();
				this.m_oColorMru.fromXml(reader);
			} else if ("colormenu" === name) {

				this.m_oColorMenu = new CColorMenu();
				this.m_oColorMenu.fromXml(reader);
			}
		};
		CShapeDefaults.prototype.writeAttrXmlImpl = function (writer) {
			writer.WriteXmlNullableAttributeString("v:ext", this.m_oExt);

			if (0 !== this.m_oSpIdMax)
				writer.WriteXmlNullableAttributeInt("spidmax", this.m_oSpIdMax);

			this.m_oStyle && writer.WriteXmlNullableAttributeString("style", this.m_oStyle.ToString());

			if (true !== this.m_oFill)
				writer.WriteXmlNullableAttributeString("fill", "false");

			writer.WriteXmlNullableAttributeString("fillcolor", getColorType(this.m_oFillColor));

			if (true !== this.m_oStroke)
				writer.WriteXmlNullableAttributeString("stroke", "false");

			if (this.m_oStrokeColor && EColorType.colortypeBlack !== this.m_oStrokeColor.type)
				writer.WriteXmlNullableAttributeString("strokecolor", getColorType(this.m_oStrokeColor));

			if (false !== this.m_oAllowInCell)
				writer.WriteXmlNullableAttributeString("o:allowincell", "true");

		};
		CShapeDefaults.prototype.writeChildren = function (writer) {

			if (this.m_oVmlFill !== null)
				this.m_oVmlFill.toXml(writer, "v:fill");

			if (this.m_oVmlStroke !== null)
				this.m_oVmlStroke.toXml(writer, "v:stroke");

			if (this.m_oVmlTextbox !== null)
				this.m_oVmlTextbox.toXml(writer, "v:textbox");

			if (this.m_oVmlShadow !== null)
				this.m_oVmlShadow.toXml(writer, "v:shadow");

			if (this.m_oSkew !== null)
				this.m_oSkew.toXml(writer, "o:skew");

			if (this.m_oExtrusion !== null)
				this.m_oExtrusion.toXml(writer, "o:extrusion");

			if (this.m_oCallout !== null)
				this.m_oCallout.toXml(writer, "o:callout");

			if (this.m_oLock !== null)
				this.m_oLock.toXml(writer, "o:lock");

			if (this.m_oColorMru !== null)
				this.m_oColorMru.toXml(writer, "o:colormru");

			if (this.m_oColorMenu !== null)
				this.m_oColorMenu.toXml(writer, "o:colormenu");

		};

		//VmlWord


		function CAnchorLock() {
			CBaseNoId.call(this);
		}

		IC(CAnchorLock, CBaseNoId, 0);

		function CBorder(sType) {
			CBaseNoId.call(this);
			this.m_sType = sType;
			this.m_oShadow = null;
			this.m_oType = null;
			this.m_oWidth = null;
		}

		IC(CBorder, CBaseNoId, 0);
		CBorder.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "shadow":
					this.m_oShadow = readBorderShadow(reader);
					break;
				case "type":
					this.m_oType = readBorderType(reader);
					break;
				case "width":
					this.m_oWidth = reader.GetValueInt();
					break;
			}
		};
		CBorder.prototype.writeAttrXmlImpl = function (writer) {

			if (this.m_oType !== null)
				writer.WriteXmlNullableAttributeString("wd:type", getBorderType(this.m_oType));

			if (this.m_oWidth !== null)
				writer.WriteXmlNullableAttributeInt("wd:width", this.m_oWidth);

			if (this.m_oShadow !== null)
				writer.WriteXmlNullableAttributeString("wd:shadow", getBorderShadow(this.m_oShadow));
		};


		function CWrap() {
			CBaseNoId.call(this);

			this.m_oAnchorX = null;
			this.m_oAnchorY = null;
			this.m_oSide = null;
			this.m_oType = null;
		}

		IC(CWrap, CBaseNoId, 0);
		CWrap.prototype.readAttrXml = function (name, reader) {
			switch (name) {
				case "anchorx":
					this.m_oAnchorX = readHorizontalAnchor(reader);
					break;
				case "anchory":
					this.m_oAnchorY = readVerticalAnchor(reader);
					break;
				case "side":
					this.m_oSide = readWrapSide(reader);
					break;
				case "type":
					this.m_oType = readWrapType(reader);
					break;
			}
		};
		CWrap.prototype.writeAttrXmlImpl = function (writer) {

			if (this.m_oType !== null)
				writer.WriteXmlNullableAttributeString("type", getWrapType(this.m_oType));

			if (this.m_oSide !== null)
				writer.WriteXmlNullableAttributeString("side", getWrapSide(this.m_oSide));

			if (this.m_oAnchorX !== null)
				writer.WriteXmlNullableAttributeString("anchorx", getHorizontalAnchor(this.m_oAnchorX));

			if (this.m_oAnchorY !== null)
				writer.WriteXmlNullableAttributeString("anchory", getVerticalAnchor(this.m_oAnchorY));

		};


		function Cm_To_Mm(dValue) {
			return dValue * 10;
		}

		function Cm_To_Pt(dValue) {
			return dValue * 72 / 2.54;
		}

		function Cm_To_Px(dValue) {
			return dValue * 72 * 4 / 3 / 2.54;
		}

		function Cm_To_Inch(dValue) {
			return dValue / 2.54;
		}

		function Cm_To_Dx(dValue) {
			return dValue * 72 * 20 / 2.54;
		}

		function Cm_To_Sx(dValue) {
			return dValue * 72 * 100 * 1000 / 20;
		}

		function Cm_To_Multi(dValue) {
			return dValue * 72 * 20 / 2.54;
		}

		function Cm_To_Emu(dValue) {
			return dValue * 360000;
		}

		function Mm_To_Cm(dValue) {
			return dValue / 10;
		}

		function Mm_To_Pt(dValue) {
			return dValue * 72 / 10 / 2.54;
		}

		function Mm_To_Px(dValue) {
			return dValue * 72 * 4 / 3 / 10 / 2.54;
		}

		function Mm_To_Inch(dValue) {
			return dValue / 2.54 / 10;
		}

		function Mm_To_Dx(dValue) {
			return dValue * 72 * 20 / 10 / 2.54;
		}

		function Mm_To_Sx(dValue) {
			return dValue * 72 * 100 * 1000 / 10 / 20;
		}

		function Mm_To_Multi(dValue) {
			return dValue * 72 * 20 / 10 / 2.54;
		}

		function Mm_To_Emu(dValue) {
			return dValue * 36000;
		}

		function Pt_To_Cm(dValue) {
			return dValue * 2.54 / 72;
		}

		function Pt_To_Mm(dValue) {
			return dValue * 2.54 * 10 / 72;
		}

		function Pt_To_Px(dValue) {
			return dValue * 4 / 3;
		}

		function Pt_To_Inch(dValue) {
			return dValue / 72;
		}

		function Pt_To_Dx(dValue) {
			return dValue * 20;
		}

		function Pt_To_Sx(dValue) {
			return dValue * 2.54 * 100 * 1000 / 20;
		}

		function Pt_To_Multi(dValue) {
			return dValue * 20;
		}

		function Pt_To_Emu(dValue) {
			return dValue * 12700;
		}

		function Px_To_Cm(dValue) {
			return dValue * 2.54 * 3 / 72 / 4;
		}

		function Px_To_Mm(dValue) {
			return dValue * 2.54 * 10 * 3 / 72 / 4;
		}

		function Px_To_Pt(dValue) {
			return dValue * 3 / 4;
		}

		function Px_To_Inch(dValue) {
			return dValue * 3 / 72 / 4;
		}

		function Px_To_Dx(dValue) {
			return dValue * 20 * 3 / 4;
		}

		function Px_To_Sx(dValue) {
			return dValue * 2.54 * 100 * 1000 * 3 / 20 / 4;
		}

		function Px_To_Multi(dValue) {
			return dValue * 20 * 3 / 4;
		}

		function Px_To_Emu(dValue) {
			return dValue * 9525;
		}

		function Inch_To_Cm(dValue) {
			return dValue * 2.54;
		}

		function Inch_To_Mm(dValue) {
			return dValue * 2.54 * 10;
		}

		function Inch_To_Pt(dValue) {
			return dValue * 72;
		}

		function Inch_To_Px(dValue) {
			return dValue * 72 * 4 / 3;
		}

		function Inch_To_Dx(dValue) {
			return dValue * 72 * 20;
		}

		function Inch_To_Sx(dValue) {
			return dValue * 1000 * 100 * 2.54 * 72 / 20;
		}

		function Inch_To_Multi(dValue) {
			return dValue * 72 * 20;
		}

		function Inch_To_Emu(dValue) {
			return dValue * 914400;
		}

		function Dx_To_Cm(dValue) {
			return dValue * 2.54 / 72 / 20;
		}

		function Dx_To_Mm(dValue) {
			return dValue * 2.54 * 10 / 72 / 20;
		}

		function Dx_To_Pt(dValue) {
			return dValue / 20;
		}


		function Dx_To_Px(dValue) {
			return dValue * 4 / 3 / 20;
		}

		function Dx_To_Inch(dValue) {
			return dValue / 20 / 72;
		}

		function Dx_To_Sx(dValue) {
			return dValue * 635;
		}

		function Dx_To_Multi(dValue) {
			return dValue;
		}

		function Dx_To_Emu(dValue) {
			return dValue * 635;
		}

		function Sx_To_Cm(dValue) {
			return dValue * 20 / 72 / 100 / 1000;
		}

		function Sx_To_Mm(dValue) {
			return dValue * 20 / 72 / 100 / 1000 * 10;
		}

		function Sx_To_Pt(dValue) {
			return dValue * 20 / 100 / 1000 / 2.54;
		}

		function Sx_To_Px(dValue) {
			return dValue * 20 * 4 / 3 / 100 / 1000 / 2.54;
		}

		function Sx_To_Inch(dValue) {
			return dValue * 20 / 2.54 / 72 / 100 / 1000;
		}

		function Sx_To_Dx(dValue) {
			return dValue * 20 * 20 / 2.54 / 100 / 1000;
		}

		function Sx_To_Multi(dValue) {
			return dValue * 20 * 20 / 2.54 / 100 / 1000;
		}

		function Sx_To_Emu(dValue) {
			return dValue;
		}

		function Multi_To_Cm(dValue) {
			return dValue * 2.54 / 72 / 20;
		}

		function Multi_To_Mm(dValue) {
			return dValue * 2.54 * 10 / 72 / 20;
		}

		function Multi_To_Pt(dValue) {
			return dValue / 20;
		}

		function Multi_To_Px(dValue) {
			return dValue * 4 / 3 / 20;
		}

		function Multi_To_Inch(dValue) {
			return dValue / 20 / 72;
		}

		function Multi_To_Sx(dValue) {
			return dValue * 635;
		}

		function Multi_To_Dx(dValue) {
			return dValue;
		}

		function Multi_To_Emu(dValue) {
			return dValue * 635;
		}

		function Emu_To_Cm(dValue) {
			return dValue / 360000;
		}

		function Emu_To_Mm(dValue) {
			return dValue / 36000;
		}

		function Emu_To_Pt(dValue) {
			return dValue / 12700;
		}

		function Emu_To_Twips(dValue) {
			return dValue / 635;
		}

		function Emu_To_Px(dValue) {
			return dValue / 9525;
		}

		function Emu_To_Inch(dValue) {
			return dValue / 914400;
		}

		function Emu_To_Sx(dValue) {
			return dValue;
		}

		function Emu_To_Dx(dValue) {
			return dValue / 635;
		}

		function Emu_To_Multi(dValue) {
			return dValue / 635;
		}

		function CUniversalMeasure() {
			this.m_bUnit = null;
			this.m_dValue = null;
		}

		CUniversalMeasure.prototype.GetValue = function () {
			return this.m_dValue;
		};
		CUniversalMeasure.prototype.ToPoints = function () {
			return this.m_dValue;
		};
		CUniversalMeasure.prototype.ToInches = function () {
			return this.m_dValue / 72.0;
		};
		CUniversalMeasure.prototype.ToMm = function () {
			return this.m_dValue * 25.4 / 72;
		}
		CUniversalMeasure.prototype.ToTwips = function () {
			return Pt_To_Dx(this.m_dValue);
		}
		CUniversalMeasure.prototype.ToHps = function () {
			return (this.m_dValue * 2);
		}
		CUniversalMeasure.prototype.ToUnsignedTwips = function () {
			return Math.abs(Pt_To_Dx(this.m_dValue));
		}
		CUniversalMeasure.prototype.FromPoints = function (dValue) {
			this.m_dValue = dValue;
			return this.m_dValue;
		}
		CUniversalMeasure.prototype.FromTwips = function (dValue) {
			this.m_dValue = Dx_To_Pt(dValue);
			return this.m_dValue;
		}
		CUniversalMeasure.prototype.FromMm = function (dValue) {
			this.m_dValue = Mm_To_Pt(dValue);
			return this.m_dValue;
		}
		CUniversalMeasure.prototype.FromInches = function (dValue) {
			this.m_dValue = Inch_To_Pt(dValue);
			return this.m_dValue;
		}
		CUniversalMeasure.prototype.FromEmu = function (dValue) {
			this.m_dValue = Emu_To_Pt(dValue);
			return this.m_dValue;
		}
		CUniversalMeasure.prototype.IsUnits = function () {
			return this.m_bUnit;
		}
		CUniversalMeasure.prototype.Parse = function (sValue, dKoef) {
			this.m_bUnit = false;
			this.m_dValue = 0;

			if (sValue.length === 0) return;

			if (sValue.length <= 2) {
				this.m_dValue = parseFloat(sValue) / dKoef;
				return;
			}

			let sUnit = sValue.substr(sValue.length - 2, 2);
			this.m_bUnit = true;

			if ("cm" === sUnit) {
				let dValue = parseFloat(sValue.substr(0, sValue.length - 2));
				this.m_dValue = Cm_To_Pt(dValue);
			} else if ("mm" === sUnit) {
				let dValue = parseFloat(sValue.substr(0, sValue.length - 2));
				this.m_dValue = Mm_To_Pt(dValue);
			} else if ("in" === sUnit) {
				let dValue = parseFloat(sValue.substr(0, sValue.length - 2));
				this.m_dValue = Inch_To_Pt(dValue);
			} else if ("pt" === sUnit) {
				this.m_dValue = parseFloat(sValue.substr(0, sValue.length - 2));
			} else if ("pc" === sUnit) {
				let dValue = parseFloat(sValue.substr(0, sValue.length - 2));
				this.m_dValue = dValue * 12.0;
			} else if ("pi" === sUnit) {
				let dValue = parseFloat(sValue.substr(0, sValue.length - 2));
				this.m_dValue = dValue * 12.0;
			} else {
				this.m_bUnit = false;
				this.m_dValue = parseFloat(sValue) / dKoef;

			}
		}

		function CUniversalMeasureOrPercent() {
			CUniversalMeasure.call(this);
			this.m_bTrailingPercentSign = null;
		}

		IC(CUniversalMeasureOrPercent, CUniversalMeasure, 0);
		CUniversalMeasureOrPercent.prototype.SetValue = function (dValue) {
			this.m_bUnit = false;
			this.m_dValue = dValue;
		};
		CUniversalMeasureOrPercent.prototype.FromString = function (sValue) {
			this.m_bUnit = false;
			this.m_bTrailingPercentSign = false;
			if (sValue.length === 0) {
				this.m_dValue = 0;
				return this.m_dValue;
			}
			if ('%' === sValue[sValue.length - 1]) {
				this.m_bTrailingPercentSign = true;
				this.m_dValue = parseFloat(sValue.substr(0, sValue.length - 1));
			} else {
				this.Parse(sValue, 1);
			}
			return this.m_dValue;
		};
		CUniversalMeasureOrPercent.prototype.ToString = function () {
			let sResult;
			if (this.m_bUnit)
				sResult = this.m_dValue + "pt";
			else if (this.m_bTrailingPercentSign)
				sResult = this.m_dValue + "%";
			else
				sResult = this.m_dValue + "";
			return sResult;
		};
		CUniversalMeasureOrPercent.prototype.IsPercent = function () {
			return this.m_bTrailingPercentSign;
		};

		function readPoint(reader) {
			return new CPoint(reader.GetValue());
		}

		function CPoint(sValue) {
			CUniversalMeasure.call(this);
			if (sValue) {
				this.FromString(sValue);
			}
		}

		IC(CPoint, CUniversalMeasure, 0);
		CPoint.prototype.FromString = function (sValue) {
			this.Parse(sValue, 1);
			return this.m_dValue;
		};
		CPoint.prototype.SetValue = function (dValue) {
			this.m_bUnit = false;
			this.m_dValue = dValue;
		};
		CPoint.prototype.ToString = function () {
			return this.m_dValue + "pt";
		};
		CPoint.prototype.FromPoints = function (dValue) {
			this.m_dValue = dValue;
			return this.m_dValue;
		};
		CPoint.prototype.FromInches = function (dValue) {
			this.m_dValue = dValue * 72;
			return this.m_dValue;
		};
		CPoint.prototype.GetValue = function () {
			return this.m_dValue;
		};

		function getPoint(oPt) {
			if (oPt) {
				return oPt.ToString();
			}
			return null;
		}

		function CInch() {
			CUniversalMeasure.call(this);
		}

		IC(CInch, CUniversalMeasure, 0)
		CInch.prototype.FromString = function (sValue) {
			this.Parse(sValue, 1.0 / 72);
			return this.m_dValue;
		};
		CInch.prototype.SetValue = function (dValue) {
			this.m_bUnit = false;
			this.m_dValue = this.FromInches(dValue);
		};
		CInch.prototype.ToString = function () {
			return this.ToInches() + "in";
		};


		function _vml_shape()
		{
			this.nId = null;		// for comments
			this.sXml = null;		// for pptx
			this.pElement = null;	// for docx/xlsx
			this.bUsed = false;		// for single drawing
		}


		function CVMLDrawing() {
			CBaseNoId.call(this);
			this.items = [];

			this.m_oReadPath = null;
			this.m_mapShapes = {};
			this.m_arrShapeTypes = [];
//writing
			this.m_mapComments = {};
			this.m_arObjectXml = [];
			this.m_arControlXml = [];

			this.m_lObjectIdVML = null;
		}

		IC(CVMLDrawing, CBaseNoId, 0);
		CVMLDrawing.prototype.fromXml = function (reader, bSkipFirstNode) {
			if (bSkipFirstNode) {
				if (!reader.ReadNextNode()) {
					return;
				}
			}
			this.readAttr(reader);
			this.elementContent ="";
			this.bReadyElement = false;
			let depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				let name = reader.GetNameNoNS();
				this.readChildXml(name, reader);
			}
		};
		CVMLDrawing.prototype.readChildXml = function (name, reader) {
			let oItem;
			if ("arc" === name) {
				oItem = new CArc();
				 this.bReadyElement = true;
			}
			if ("curve" === name) {
				oItem = new CCurve();
				 this.bReadyElement = true;
			}
			if ("group" === name) {
				oItem = new CGroup();
				 this.bReadyElement = true;
			}
			if ("image" === name) {
				oItem = new CImage();
				 this.bReadyElement = true;
			}
			if ("line" === name) {
				oItem = new CLine();
				 this.bReadyElement = true;
			}
			if ("oval" === name) {
				oItem = new COval();
				 this.bReadyElement = true;
			}
			if ("polyline" === name) {
				oItem = new CPolyLine();
				 this.bReadyElement = true;
			}
			if ("rect" === name) {
				oItem = new CRect();
				 this.bReadyElement = true;
			}
			if ("roundrect" === name) {
				oItem = new CRoundRect();
				 this.bReadyElement = true;
			}
			if ("shape" === name) {
				oItem = new CShape();
				 this.bReadyElement = true;
			}
			if ("shapetype" === name) {
				oItem = new CShapeType();
			}
			if (oItem) {
				oItem.fromXml(reader);
				this.items.push(oItem);



				let common = oItem;
				let sSpid;
				let bComment = false;

				if (common)
				{
					if (common.m_sSpId !== null)	sSpid = common.m_sSpId;
				else if (common.m_sId !== null)sSpid = common.m_sId;

					bComment = common.m_bComment;
				}
				if ( this.bReadyElement)
				{
					if (sSpid && sSpid !== "")
					{
						let element = new _vml_shape();

						element.nId			= this.items.length - 1;
						element.sXml		= this.elementContent;
						element.pElement	= oItem;
						element.bUsed       = bComment;

						this.m_mapShapes[sSpid] = element;
					}
					this.elementContent = "";
					 this.bReadyElement   = false;
					bComment        = false;
				}
				else
				{
					let element = new _vml_shape();
					element.nId			= -1;
					element.sXml		= this.elementContent;
					element.pElement	= oItem;
					element.bUsed       = false;
					this.m_arrShapeTypes.push(element);
				}
			}
		};
		CVMLDrawing.prototype.writeChildren = function (writer) {
		};

		function isEmptyObject(oObj) {
			if(!oObj) {
				return false;
			}
			for(let sKey in oObj) {
				if(oObj.hasOwnProperty(sKey)) {
					return false;
				}
			}
			return true;
		}
		CVMLDrawing.prototype.getXmlString = function() {
			if((!this.m_mapComments || isEmptyObject(this.m_mapComments)) && isEmptyObject(this.m_arObjectXml) && isEmptyObject(this.m_arControlXml))
				return "";

			let sXml = "";
			sXml += "<xml \
xmlns:v=\"urn:schemas-microsoft-com:vml\" \
xmlns:o=\"urn:schemas-microsoft-com:office:office\" \
xmlns:x=\"urn:schemas-microsoft-com:office:excel\">";

			for (let i = 0; i < this.m_arObjectXml.length; ++i)
			{
				sXml += (this.m_arObjectXml[i]);
			}

			if (false ===isEmptyObject(this.m_arControlXml) || ((null !== this.m_mapComments) && (false === isEmptyObject(this.m_mapComments))))
			{
				sXml += ("<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>");
			}

			if (this.m_arControlXml.length > 0)
			{

				sXml += ("<v:shapetype id=\"_x0000_t201\" coordsize=\"21600,21600\" o:spt=\"201\"");
				sXml += (" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/>");
				sXml += ("<v:path shadowok=\"f\" o:extrusionok=\"f\" strokeok=\"f\" fillok=\"f\" o:connecttype=\"rect\"/>");
				sXml += ("<o:lock v:ext=\"edit\" shapetype=\"t\"/></v:shapetype>");

				for (let i = 0; i < this.m_arControlXml.length; ++i)
				{
					sXml += (this.m_arControlXml[i]);
				}
			}

			let nIndex = this.m_lObjectIdVML + 1;
			if ((null !== this.m_mapComments) && (false === isEmptyObject(this.m_mapComments)))
			{
				sXml += ("<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"");
				sXml += (" path=\"m,l,21600r21600,l21600,xe\">");
				sXml += ("<v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>");

				for (let sKey in this.m_mapComments)
				{
					let comment = this.m_mapComments[sKey];
					let oCoords = comment.coords;
					let sStyle = "";
					if(oCoords.dLeftMM !== null)
					{
						let oPoint = new CPoint(""); oPoint.FromMm(oCoords.dLeftMM);
						sStyle += "margin-left:" + oPoint.ToPoints() + "pt;";
					}
					if(oCoords.dTopMM !== null)
					{
						let oPoint = new CPoint(""); oPoint.FromMm(oCoords.dTopMM);
						sStyle += "margin-top:" + oPoint.ToPoints() + "pt;";
					}
					if(oCoords.dWidthMM !== null)
					{
						let oPoint = new CPoint(""); oPoint.FromMm(oCoords.dWidthMM);
						sStyle += "width:" + oPoint.ToPoints() + "pt;";
					}
					if(oCoords.dHeightMM !== null)
					{
						let oPoint = new CPoint(""); oPoint.FromMm(oCoords.dHeightMM);
						sStyle += "height:" + oPoint.ToPoints() + "pt;";
					}
					let sClientData = "<x:ClientData ObjectType=\"Note\">";

					if(oCoords.bMoveWithCells !== null && true === oCoords.bMoveWithCells)
						sClientData += "<x:MoveWithCells/>";

					if(oCoords.bSizeWithCells !== null && true === oCoords.bSizeWithCells)
						sClientData += "<x:SizeWithCells/>";

					if( oCoords.nLeft !== null   && oCoords.nLeftOffset !== null  &&
						oCoords.nTop !== null    && oCoords.nTopOffset !== null   &&
						oCoords.nRight !== null  && oCoords.nRightOffset !== null &&
						oCoords.nBottom !== null && oCoords.nBottomOffset !== null)
					{
						sClientData += "<x:Anchor>";
						sClientData += (oCoords.nLeft)          + ",";
						sClientData += (oCoords.nLeftOffset)    + ",";
						sClientData += (oCoords.nTop)           + ",";
						sClientData += (oCoords.nTopOffset)     + ",";
						sClientData += (oCoords.nRight)         + ",";
						sClientData += (oCoords.nRightOffset)   + ",";
						sClientData += (oCoords.nBottom)        + ",";
						sClientData += (oCoords.nBottomOffset);
						sClientData += "</x:Anchor>";
					}
					sClientData += "<x:AutoFill>False</x:AutoFill>";

					if(oCoords.nRow !== null)
						sClientData += "<x:Row>" + (oCoords.nRow) + "</x:Row>";

					if(oCoords.nCol !== null)
						sClientData += "<x:Column>" + (oCoords.nCol) + "</x:Column>";

					sClientData += "</x:ClientData>";

					let sGfxdata = "";
					if(comment.m_sGfxdata !== null)
						sGfxdata = "o:gfxdata=\"" + comment.m_sGfxdata + "\"";

					let sShape = "";
					sShape += "<v:shape id=\"_x0000_s" + (nIndex++) + " \" type=\"#_x0000_t202\" style='position:absolute;";
					sShape += sStyle;
					sShape += "z-index:4;visibility:hidden' ";
					sShape += sGfxdata;
					sShape += " fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div></v:textbox>";
					sShape += sClientData;
					sShape += "</v:shape>";

					sXml += (sShape);
				}
			}
			sXml += ("</xml>");
			return sXml;
		};
		CVMLDrawing.prototype.write = function (writer) {
			writer.WriteXmlString(this.getXmlString());
		};


		let shemeDefaultColor =
			[
				0x00000000, 0x00FFFFFF, 0x00FF0000, 0x0000FF00, 0x000000FF, 0x00FFFF00, 0x00FF00FF, 0x0000FFFF,
				0x00000000, 0x00FFFFFF, 0x00FF0000, 0x0000FF00, 0x000000FF, 0x00FFFF00, 0x00FF00FF, 0x0000FFFF,
				0x00800000, 0x00008000, 0x00000080, 0x00808000, 0x00800080, 0x00008080, 0x00C0C0C0, 0x00808080,
				0x009999FF, 0x00993366, 0x00FFFFCC, 0x00CCFFFF, 0x00660066, 0x00FF8080, 0x000066CC, 0x00CCCCFF,
				0x00000080, 0x00FF00FF, 0x00FFFF00, 0x0000FFFF, 0x00800080, 0x00800000, 0x00008080, 0x000000FF,
				0x0000CCFF, 0x00CCFFFF, 0x00CCFFCC, 0x00FFFF99, 0x0099CCFF, 0x00FF99CC, 0x00CC99FF, 0x00FFCC99,
				0x003366FF, 0x0033CCCC, 0x0099CC00, 0x00FFCC00, 0x00FF9900, 0x00FF6600, 0x00666699, 0x00969696,
				0x00003366, 0x00339966, 0x00003300, 0x00333300, 0x00993300, 0x00993366, 0x00333399, 0x00333333
			];
		let controlPanelColors2 =
			[
				0x00000000, 0x00FFFFFF, 0x00000000, 0x00FFFFFF,
				0x00000000, 0x00000000, 0x00000000, 0x00FFFFFF,
				0x00FFFFFF, 0x00000000, 0x00FFFFFF, 0x00FFFFFF,
				0x00000000, 0x00000000, 0x00000000, 0x00000000,
				0x00FFFFFF, 0x00FFFFFF, 0x00FFFFFF, 0x00000000,
				0x00FFFFFF, 0x00000000, 0x00000000, 0x00000000,
				0x00000000, 0x00000000, 0x00FFFFFF, 0x00FFFFFF
			];
		let controlPanelColors1 =
			[
				0x00FFFFFF, 0x00CCCCCC, 0x00FFFFFF, 0x006363CE,
				0x00DDDDDD, 0x00DDDDDD, 0x00888888, 0x00000000,
				0x00000000, 0x00808080, 0x00B5D5FF, 0x00000000,
				0x00FFFFFF, 0x00FFFFFF, 0x007F7F7F, 0x00FBFCC5,
				0x00000000, 0x00F7F7F7, 0x00000000, 0x00FFFFFF,
				0x00666666, 0x00C0C0C0, 0x00DDDDDD, 0x00C0C0C0,
				0x00888888, 0x00FFFFFF, 0x00CCCCCC, 0x00000000
			];

		let EColorType = {
			colortypeNone: 0,
			colortypeRGB: 1,
			colortypeAqua: 2,
			colortypeBlack: 3,
			colortypeBlue: 4,
			colortypeFuchsia: 5,
			colortypeGray: 6,
			colortypeGreen: 7,
			colortypeLime: 8,
			colortypeMaroon: 9,
			colortypeNavy: 10,
			colortypeOlive: 11,
			colortypePurple: 12,
			colortypeRed: 13,
			colortypeSilver: 14,
			colortypeTeal: 15,
			colortypeWhite: 16,
			colortypeYellow: 17
		};

		let EAlternateMathContentType =
			{
				alternatemathcontenttypeOfficeOpenXmlMath: 0,
				alternatemathcontenttypeMathMl: 1
			};

		function readAlternateMathContentType(reader) {
			let sVal = reader.GetValue();
			if (sVal === "officeopenxmlmath") {
				return EAlternateMathContentType.alternatemathcontenttypeOfficeOpenXmlMath
			}
			if (sVal === "mathml") {
				return EAlternateMathContentType.alternatemathcontenttypeMathMl
			}
			return EAlternateMathContentType.alternatemathcontenttypeOfficeOpenXmlMath;
		}

		function getAlternateMathContentType(nType) {
			if (nType === EAlternateMathContentType.alternatemathcontenttypeOfficeOpenXmlMath)
				return "officeopenxmlmath";
			if (nType === EAlternateMathContentType.alternatemathcontenttypeMathMl)
				return "mathml";
			return "officeopenxmlmath";
		}

		function readContentType(reader) {

			let sVal = reader.GetValue();
			return sVal;
		}

		function getContentType(type) {
			return type;
		}

		function CColor(sVal) {
			this.type = EColorType.colortypeRGB;
			this.val = sVal;
			this.r = 0;
			this.g = 0;
			this.b = 0;
			this.fromString(sVal);
		}

		CColor.prototype.fromString = function (sVal) {
			this.val = sVal;
			if (sVal.charAt(0) === '#') {
				this.byHexColor(sVal)
			} else {
				this.byColorName(sVal);
			}
		};

		CColor.prototype.toString = function () {
			return this.val;
		};
		CColor.prototype.Get_R = function () {
			return this.r;
		};
		CColor.prototype.Get_G = function () {
			return this.g;
		};
		CColor.prototype.Get_B = function () {
			return this.b;
		};
		CColor.prototype.setRGB = function () {

			if (this.type === EColorType.colortypeRGB) return;

			switch (this.type) {
				case EColorType.colortypeAqua: {
					this.r = 0x00;
					this.g = 0xff;
					this.b = 0xff;
				}
					break;
				case EColorType.colortypeBlack: {
					this.r = 0x00;
					this.g = 0x00;
					this.b = 0x00;
				}
					break;
				case EColorType.colortypeBlue: {
					this.r = 0x00;
					this.g = 0x00;
					this.b = 0xff;
				}
					break;
				case EColorType.colortypeFuchsia: {
					this.r = 0xff;
					this.g = 0x00;
					this.b = 0xff;
				}
					break;
				case EColorType.colortypeGray: {
					this.r = 0x80;
					this.g = 0x80;
					this.b = 0x80;
				}
					break;
				case EColorType.colortypeGreen: {
					this.r = 0x00;
					this.g = 0x80;
					this.b = 0x00;
				}
					break;
				case EColorType.colortypeLime: {
					this.r = 0x00;
					this.g = 0xff;
					this.b = 0x00;
				}
					break;
				case EColorType.colortypeMaroon: {
					this.r = 0x80;
					this.g = 0x00;
					this.b = 0x00;
				}
					break;
				case EColorType.colortypeNavy: {
					this.r = 0x00;
					this.g = 0x00;
					this.b = 0x80;
				}
					break;
				case EColorType.colortypeOlive: {
					this.r = 0x80;
					this.g = 0x80;
					this.b = 0x00;
				}
					break;
				case EColorType.colortypePurple: {
					this.r = 0x80;
					this.g = 0x00;
					this.b = 0x80;
				}
					break;
				case EColorType.colortypeRed: {
					this.r = 0xff;
					this.g = 0x00;
					this.b = 0x00;
				}
					break;
				case EColorType.colortypeSilver: {
					this.r = 0xc0;
					this.g = 0xc0;
					this.b = 0xc0;
				}
					break;
				case EColorType.colortypeTeal: {
					this.r = 0x00;
					this.g = 0x80;
					this.b = 0x80;
				}
					break;
				case EColorType.colortypeWhite: {
					this.r = 0xff;
					this.g = 0xff;
					this.b = 0xff;
				}
					break;
				case EColorType.colortypeYellow: {
					this.r = 0xff;
					this.g = 0xff;
					this.b = 0;
				}
					break;
				case EColorType.colortypeNone:
				default: {
					this.r = 0;
					this.g = 0;
					this.b = 0;
				}
					break;
			}
		};
		CColor.prototype.byHexColor = function (sVal) {
			this.type = EColorType.colortypeRGB;
			let nColor = parseInt(sVal.slice(1));
			this.r = (nColor >> 16) & 0xFF;
			this.g = (nColor >> 8) & 0xFF;
			this.b = nColor & 0xFF;
		};
		CColor.prototype.byColorName = function (sVal) {
			this.type = EColorType.colortypeNone;

			if (sVal.indexOf("aqua") > -1) this.type = EColorType.colortypeAqua;
			else if (sVal.indexOf("black") > -1) this.type = EColorType.colortypeBlack;
			else if (sVal.indexOf("blue") > -1) this.type = EColorType.colortypeBlue;
			else if (sVal.indexOf("fuchsia") > -1) this.type = EColorType.colortypeFuchsia;
			else if (sVal.indexOf("gray") > -1) this.type = EColorType.colortypeGray;
			else if (sVal.indexOf("green") > -1) this.type = EColorType.colortypeGreen;
			else if (sVal.indexOf("lime") > -1) this.type = EColorType.colortypeLime;
			else if (sVal.indexOf("maroon") > -1) this.type = EColorType.colortypeMaroon;
			else if (sVal.indexOf("navy") > -1) this.type = EColorType.colortypeNavy;
			else if (sVal.indexOf("olive") > -1) this.type = EColorType.colortypeOlive;
			else if (sVal.indexOf("purple") > -1) this.type = EColorType.colortypePurple;
			else if (sVal.indexOf("red") > -1) this.type = EColorType.colortypeRed;
			else if (sVal.indexOf("silver") > -1) this.type = EColorType.colortypeSilver;
			else if (sVal.indexOf("teal") > -1) this.type = EColorType.colortypeTeal;
			else if (sVal.indexOf("white") > -1) this.type = EColorType.colortypeWhite;
			else if (sVal.indexOf("yellow") > -1) this.type = EColorType.colortypeYellow;
			else if (sVal.indexOf("[") > -1 && sVal.indexOf("]") > -1) {
				let p1 = sVal.indexOf("[");
				let p2 = sVal.indexOf("]");
				let sIndex = p2 > p1 ? sVal.substr(p1 + 1, p2 - p1 - 1) : "";

				if (sIndex.length > 0) {
					let index = parseInt(sIndex);
					let nRGB = 0;
					if (index < 64) {
						nRGB = shemeDefaultColor[index];
					} else if (index > 64 && index < 92) {
						nRGB = controlPanelColors1[index - 65];
					}
					this.r = ((nRGB >> 16) & 0xff);
					this.g = ((nRGB >> 8) & 0xff);
					this.b = (nRGB & 0xff);
					this.type = EColorType.colortypeRGB;
				}
			}

			this.setRGB();

		};


		function CVmlVector2D(sVal) {
			this.x = 0;
			this.y = 0;

			if (sVal) {
				this.fromString(sVal);
			}
		}

		CVmlVector2D.prototype.fromString = function (sValue) {
			let nLen = sValue.length;
			if (nLen <= 0)
				return 0;

			let nPos = sValue.indexOf(",");
			let strX, strY;
			if (-1 === nPos) {//only x coord
				strX = sValue;
			} else {
				strX = sValue.substr(0, nPos);
				strY = sValue.substr(nPos + 1, nLen - nPos - 1);
			}
			strY.replace("@", "");
			strX.replace("@", "");
			this.x = strX.length === 0 ? 0 : parseInt(strX);
			this.y = strY.length === 0 ? 0 : parseInt(strY);
			return 0;
		};

		CVmlVector2D.prototype.ToString = function () {
			return "" + this.x + "," + this.y;
		};

		function CVmlPolygon2D(sVal) {
			this.pts = [];
			if (sVal) {
				this.fromString(sVal)
			}
		}

		CVmlPolygon2D.prototype.fromString = function (sValue) {
			this.pts.length = 0;

			let nLen = sValue.length;
			if (nLen <= 0)
				return 0;

			let nStartPos = 0;
			while (true) {
				let nMidPos = sValue.indexOf(",", nStartPos);
				let nEndPos = sValue.indexOf(",", nMidPos + 1);

				if (-1 === nMidPos)
					break;

				if (-1 === nEndPos)
					nEndPos = nLen;

				let strX = sValue.substr(nStartPos, nMidPos - nStartPos);
				let strY = sValue.substr(nStartPos, nMidPos - nStartPos);

				strX.replace("@", "");
				strY.replace("@", "");

				let nX = strX.length === 0 ? 0 : parseInt(strX);
				let nY = strY.length === 0 ? 0 : parseInt(strY);

				let oPr = new CVmlVector2D();
				oPr.x = nX;
				oPr.y = nY;
				this.pts.push(oPr);
				nStartPos = nEndPos + 1;
			}


			return 0;
		};


		function CVml_Matrix(sValue) {
			this.m_dSxx = null;
			this.m_dSxy = null;
			this.m_dSyx = null;
			this.m_dSyy = null;
			this.m_dPx = null;
			this.m_dPy = null;
			this.ResetMatrix();
			if (sValue) {
				this.FromString(sValue);
			}
		}

		CVml_Matrix.prototype.ResetMatrix = function () {
			this.m_dSxx = 1;
			this.m_dSxy = 0;
			this.m_dSyx = 0;
			this.m_dSyy = 1;
			this.m_dPx = 0;
			this.m_dPy = 0;
		}
		CVml_Matrix.prototype.SetMatrix = function (dSxx, dSxy, dSyx, dSyy, dPx, dPy) {
			this.m_dSxx = dSxx;
			this.m_dSxy = dSxy;
			this.m_dSyx = dSyx;
			this.m_dSyy = dSyy;
			this.m_dPx = dPx;
			this.m_dPy = dPy;
		}
		CVml_Matrix.prototype.Get_Sxx = function () {
			return this.m_dSxx;
		}
		CVml_Matrix.prototype.Get_Sxy = function () {
			return this.m_dSxy;
		}
		CVml_Matrix.prototype.Get_Syx = function () {
			return this.m_dSyx;
		}
		CVml_Matrix.prototype.Get_Syy = function () {
			return this.m_dSyy;
		}
		CVml_Matrix.prototype.Get_Px = function () {
			return this.m_dPx;
		}
		CVml_Matrix.prototype.Get_Py = function () {
			return this.m_dPy;
		}
		CVml_Matrix.prototype.FromString = function (sValue) {
			this.ResetMatrix();

			let nLen = sValue.length;
			if (nLen <= 0)
				return 0;

			// Sxx
			let nStartPos = 0;
			let nEndPos = sValue.indexOf(",", nStartPos);
			if (-1 === nEndPos)
				nEndPos = nLen;

			if (nEndPos - nStartPos > 0) {
				let strValue = sValue.substr(nStartPos, nEndPos - nStartPos);
				this.m_dSxx = strValue.length === 0 ? 0 : parseFloat(strValue);
			}

			// Sxy
			nStartPos = nEndPos + 1;
			nEndPos = sValue.indexOf(",", nStartPos);
			if (-1 === nEndPos)
				nEndPos = nLen;

			if (nEndPos - nStartPos > 0) {
				let strValue = sValue.substr(nStartPos, nEndPos - nStartPos);
				this.m_dSxy = strValue.length === 0 ? 0 : parseFloat(strValue);
			}

			// Syx
			nStartPos = nEndPos + 1;
			nEndPos = sValue.indexOf(",", nStartPos);
			if (-1 === nEndPos)
				nEndPos = nLen;

			if (nEndPos - nStartPos > 0) {
				let strValue = sValue.substr(nStartPos, nEndPos - nStartPos);
				this.m_dSyx = strValue.length === 0 ? 0 : parseFloat(strValue);
			}

			// Syy
			nStartPos = nEndPos + 1;
			nEndPos = sValue.indexOf(",", nStartPos);
			if (-1 === nEndPos)
				nEndPos = nLen;

			if (nEndPos - nStartPos > 0) {
				let strValue = sValue.substr(nStartPos, nEndPos - nStartPos);
				this.m_dSyy = strValue.length === 0 ? 0 : parseFloat(strValue);
			}

			// Px
			nStartPos = nEndPos + 1;
			nEndPos = sValue.indexOf(",", nStartPos);
			if (-1 === nEndPos)
				nEndPos = nLen;

			if (nEndPos - nStartPos > 0) {
				let strValue = sValue.substr(nStartPos, nEndPos - nStartPos);
				this.m_dPx = strValue.length === 0 ? 0 : parseFloat(strValue);
			}

			// Py
			nStartPos = nEndPos + 1;
			nEndPos = sValue.indexOf(",", nStartPos);
			if (-1 === nEndPos)
				nEndPos = nLen;

			if (nEndPos - nStartPos > 0) {
				let strValue = sValue.substr(nStartPos, nEndPos - nStartPos);
				this.m_dPy = strValue.length === 0 ? 0 : parseFloat(strValue);
			}

			nStartPos = nEndPos + 1;
			return 0;
		}
		CVml_Matrix.prototype.ToString = function () {
			return this.m_dSxx + "," + this.m_dSxy + "," + this.m_dSyx + "," + this.m_dSyy + "," + this.m_dPx + "," + this.m_dPy;
		}


		function CPercentage(sValue) {
			{
				this.m_dValue = 0;
				if (sValue) {
					this.FromString(sValue);
				}
			}

		}

		CPercentage.prototype.GetValue = function () {
			return this.m_dValue;
		};
		CPercentage.prototype.SetValue = function (dValue) {
			this.m_dValue = dValue;
		};
		CPercentage.prototype.FromString = function (sValue) {
			let nPos = sValue.indexOf('%');
			let nLen = sValue.length;
			if (-1 === nPos || nPos !== sValue.length - 1 || nLen <= 0) {
				if (-1 === nPos) {
					let dValue = sValue.length === 0 ? 0 : parseFloat(sValue);
					if (Math.abs(dValue) >= 0 && Math.abs(dValue) <= 1) {
						this.m_dValue = dValue;
					} else {
						this.m_dValue = dValue / 1000.0;
					}
				} else
					this.m_dValue = 0;
			} else {
				let strValue = sValue.substr(0, nLen - 1);
				this.m_dValue = strValue.length === 0 ? 0 : parseFloat(strValue);
			}

			return this.m_dValue;
		};
		CPercentage.prototype.ToString = function () {
			return this.m_dValue + "%";
		};
		CPercentage.prototype.ToStringDecimalNumber = function () {
			let sResult = this.m_dValue * 1000.0 >> 0;

			return sResult;
		};

		function CVml_Vector2D_Units_Or_Percentage(sValue) {
			this.m_dX = 0;
			this.m_dY = 0;
			this.m_bUnitsX = true;
			this.m_bUnitsY = true;
			if (sValue) {
				this.FromString(sValue);
			}
		}

		CVml_Vector2D_Units_Or_Percentage.prototype.GetX = function () {
			return this.m_dX;
		};
		CVml_Vector2D_Units_Or_Percentage.prototype.GetY = function () {
			return this.m_dY;
		};
		CVml_Vector2D_Units_Or_Percentage.prototype.IsXinPoints = function () {
			return this.m_bUnitsX;
		};
		CVml_Vector2D_Units_Or_Percentage.prototype.IsYinPoints = function () {
			return this.m_bUnitsY;
		};
		CVml_Vector2D_Units_Or_Percentage.prototype.SetValue_Points = function (dX, dY) {
			this.m_dX = dX;
			this.m_dY = dY;

			this.m_bUnitsX = true;
			this.m_bUnitsY = true;
		};
		CVml_Vector2D_Units_Or_Percentage.prototype.FromString = function (sValue) {
			this.m_dX = 0;
			this.m_dY = 0;
			this.m_bUnitsX = true;
			this.m_bUnitsY = true;

			let nLen = sValue.length;
			if (nLen < 1)
				return 0;

			let nPos = sValue.indexOf(",");
			if (-1 === nPos)
				return 0;

			let sTemp = sValue.substr(0, nPos);
			if (-1 !== sTemp.indexOf('%')) {
				let oPerc = new CPercentage(sTemp);
				this.m_dX = oPerc.GetValue();
				this.m_bUnitsX = false;
			} else {
				let oPt = new CPoint(sTemp);
				this.m_dX = oPt.GetValue();
				this.m_bUnitsX = true;
			}

			sTemp = sValue.substr(nPos + 1, nLen - nPos - 1);
			if (-1 !== sTemp.indexOf('%')) {
				let oPerc = new CPercentage(sTemp);
				this.m_dY = oPerc.GetValue();
				this.m_bUnitsY = false;
			} else {
				let oPt = new CPoint(sTemp);
				this.m_dY = oPt.GetValue();
				this.m_bUnitsY = true;
			}

			return 0;
		};
		CVml_Vector2D_Units_Or_Percentage.prototype.ToString = function () {
			let sResult = this.m_dX + "";

			if (this.m_bUnitsX) sResult += "pt,";
			else sResult += "%,";

			sResult += this.m_dY;

			if (this.m_bUnitsY) sResult += "pt";
			else sResult += "%";

			return sResult;
		};


		function CVml_1_65536_Or_Percentage(sVal) {
			this.m_dValue = 0;
			if (sVal) {
				this.FromString(sVal);
			}
		}

		CVml_1_65536_Or_Percentage.prototype.GetValue = function () {
			return this.m_dValue;
		};

		CVml_1_65536_Or_Percentage.prototype.SetValue = function (dValue) {
			this.m_dValue = Math.max(0.0, Math.min(1.0, dValue));
		};

		CVml_1_65536_Or_Percentage.prototype.SetValue = function (nValue) {
			this.m_dValue = Math.max(0.0, Math.min(65536.0, nValue)) / 65536.0;
		};
		CVml_1_65536_Or_Percentage.prototype.SetPercentage = function (dValue) {
			this.m_dValue = Math.max(0.0, Math.min(100.0, dValue)) / 100.0;
		};

		CVml_1_65536_Or_Percentage.prototype.FromString = function (sValue) {
			let nLen = sValue.length;
			if (nLen <= 0)
				return 0;

			let bFraction = ('f' === sValue.charAt(nLen - 1));
			let bPercentage = ('%' === sValue.charAt(nLen - 1));

			if (bFraction) {
				let strValue = sValue.substr(0, nLen - 1);
				let nValue = strValue.length === 0 ? 0 : parseInt(strValue);

				this.SetValue(nValue);
			} else if (bPercentage) {
				let strValue = sValue.substr(0, nLen - 1);
				let dValue = strValue.length === 0 ? 0 : parseFloat(strValue);
				this.SetPercentage(dValue);
			} else {
				let dValue = sValue.length === 0 ? 0 : parseFloat(sValue);
				this.SetValue(dValue);
			}

			return this.m_dValue;
		};

		CVml_1_65536_Or_Percentage.prototype.ToString = function () {
			return this.m_dValue + "";
		};


		function CVml_Vector3D_65536(sVal) {
			this.m_nX = 0;
			this.m_nY = 0;
			this.m_nZ = 0;
			if (sVal) {
				this.fromString(sVal);
			}
		}

		CVml_Vector3D_65536.prototype.fromString = function (sValue) {
			this.m_nX = 0;
			this.m_nY = 0;
			this.m_nZ = 0;

			let nLen = sValue.length;
			if (nLen <= 0)
				return 0;

			let nPos = sValue.indexOf(",");
			if (-1 === nPos) {//only x position
				let strX = sValue;
				strX.replace("@", "");

				this.m_nX = strX.length === 0 ? 0 : parseInt(strX);
				return 0;
			}
			let strX = sValue.substr(0, nPos);
			strX.replace("@", "");

			this.m_nX = strX.length === 0 ? 0 : parseInt(strX);

			let nPos2 = sValue.indexOf(",", nPos + 1);
			if (-1 === nPos2) {// only x, y position
				let strY = sValue.substr(nPos + 1);
				strY.replace("@", "");
				this.m_nY = strY.length === 0 ? 0 : parseInt(strY);
				return 0;
			}

			let strY = sValue.substr(nPos + 1, nPos2 - nPos - 1);
			let strZ = sValue.substr(nPos2 + 1, nLen - nPos2 - 1);

			strY.replace("@", "");
			strZ.replace("@", "");

			this.m_nY = strY.length === 0 ? 0 : parseInt(strY);
			this.m_nZ = strZ.length === 0 ? 0 : parseInt(strZ);
		};
		CVml_Vector3D_65536.prototype.ToString = function () {
			return ("" + (this.m_nX) + "," + (this.m_nX)) + ("" + this.m_nY) + "," + (this.m_nZ);
		};

		function CVml_Vector3D(sVal) {
			this.m_nX = 0;
			this.m_nY = 0;
			this.m_nZ = 0;
			if (sVal) {
				this.fromString(sVal);
			}
		}

		CVml_Vector3D.prototype.GetX = function () {
			return this.m_nX;
		};
		CVml_Vector3D.prototype.GetY = function () {
			return this.m_nY;
		};
		CVml_Vector3D.prototype.GetZ = function () {
			return this.m_nZ;
		};

		CVml_Vector3D.prototype.fromString = function (sValue) {
			this.m_nX = 0;
			this.m_nY = 0;
			this.m_nZ = 0;

			let nLen = sValue.length;
			if (nLen <= 0)
				return 0;

			let nPos = sValue.indexOf(",");
			if (-1 === nPos) {//only x position
				let strX = sValue;
				strX.replace("@", "");

				this.m_nX = strX.length === 0 ? 0 : parseInt(strX);
				return 0;
			}
			let strX = sValue.substr(0, nPos);
			strX.replace("@", "");

			this.m_nX = strX.length === 0 ? 0 : parseInt(strX);

			let nPos2 = sValue.indexOf(",", nPos + 1);
			if (-1 === nPos2) {// only x, y position
				let strY = sValue.substr(nPos + 1);
				strY.replace("@", "");
				this.m_nY = strY.length === 0 ? 0 : parseInt(strY);
				return 0;
			}

			let strY = sValue.substr(nPos + 1, nPos2 - nPos - 1);
			let strZ = sValue.substr(nPos2 + 1, nLen - nPos2 - 1);

			strY.replace("@", "");
			strZ.replace("@", "");

			this.m_nY = strY.length === 0 ? 0 : parseInt(strY);
			this.m_nZ = strZ.length === 0 ? 0 : parseInt(strZ);
		};
		CVml_Vector3D.prototype.ToString = function () {
			return (this.m_nX) + "," + (this.m_nY) + "," + (this.m_nZ);
		};


		function CVml_Vector2D(sVal) {
			this.m_nX = 0;
			this.m_nY = 0;
			if (sVal) {
				this.fromString(sVal);
			}
		}

		CVml_Vector2D.prototype.GetX = function () {
			return this.m_nX;
		};
		CVml_Vector2D.prototype.GetY = function () {
			return this.m_nY;
		};

		CVml_Vector2D.prototype.fromString = function (sValue) {
			this.m_nX = 0;
			this.m_nY = 0;

			let nLen = sValue.length;
			if (nLen <= 0)
				return 0;

			let nPos = sValue.indexOf(",");

			let strX, strY;
			if (-1 === nPos) {//only x coord
				strX = sValue;
			} else {
				strX = sValue.substr(0, nPos);
				strY = sValue.substr(nPos + 1, nLen - nPos - 1);
			}
			strY.replace("@", "");
			strX.replace("@", "");

			this.m_nX = strX.length === 0 ? 0 : parseInt(strX);
			this.m_nY = strY.length === 0 ? 0 : parseInt(strY);
		};
		CVml_Vector2D.prototype.ToString = function () {
			return (this.m_nX) + "," + (this.m_nY);
		};


		function CVml_Vector2D_F(sVal) {
			this.m_nX = 0;
			this.m_nY = 0;
			if (sVal) {
				this.fromString(sVal);
			}
		}

		CVml_Vector2D_F.prototype.fromString = function (sValue) {
			this.m_nX = 0;
			this.m_nY = 0;

			let nLen = sValue.length;
			if (nLen <= 0)
				return 0;

			let nPos = sValue.indexOf(",");

			let strX, strY;
			if (-1 === nPos) {//only x coord
				strX = sValue;
			} else {
				strX = sValue.substr(0, nPos);
				strY = sValue.substr(nPos + 1, nLen - nPos - 1);
			}
			strY.replace("@", "");
			strX.replace("@", "");

			this.m_nX = strX.length === 0 ? 0 : parseFloat(strX);
			this.m_nY = strY.length === 0 ? 0 : parseFloat(strY);
		};
		CVml_Vector2D_F.prototype.GetX = function () {
			return this.m_nX;
		};
		CVml_Vector2D_F.prototype.GetY = function () {
			return this.m_nY;
		};
		CVml_Vector2D_F.prototype.ToString = function () {
			return (this.m_nX) + "," + (this.m_nY);
		};

		let EInsetMode =
			{
				insetmodeAuto: 0,
				insetmodeCustom: 1
			};

		//--------------------------------------------------------------------------------
		let EBWMode =
			{
				bwmodeAuto: 0,
				bwmodeBlack: 1,
				bwmodeBlackTextAndLines: 2,
				bwmodeColor: 3,
				bwmodeGrayOutline: 4,
				bwmodeGrayScale: 5,
				bwmodeHide: 6,
				bwmodeHighContrast: 7,
				bwmodeInverseGray: 8,
				bwmodeLightGrayscale: 9,
				bwmodeUndrawn: 10,
				bwmodeWhite: 11
			};


		let EColorMode =
			{
				colormodeAuto: 0,
				colormodeCustom: 1
			};

		function readColorMode(reader) {
			let sVal = reader.GetValue()
			switch (sVal) {
				case "auto": {
					return EColorMode.colormodeAuto;
				}
				case "custom": {
					return EColorMode.colormodeCustom;
				}
			}
			return null;
		}

		function getColorMode(nType) {
			if (nType === EColorMode.colormodeAuto) {
				return "auto";
			}
			if (nType === EColorMode.colormodeCustom) {
				return "custom";
			}
			return null;
		}


		function readInsetMode(reader) {

			let sVal = reader.GetValue()
			switch (sVal) {
				case "auto": {
					return EInsetMode.insetmodeAuto;
				}
				case "custom": {
					return EInsetMode.insetmodeCustom;
				}
			}
			return null;
		}

		function getInsetMode(nType) {
			if (nType === EInsetMode.insetmodeAuto) {
				return "auto";
			}
			if (nType === EInsetMode.insetmodeCustom) {
				return "custom";
			}
			return "auto";
		}

		function readColorType(reader) {
			return new CColor(reader.GetValue());
		}

		function getColorType(oColorType) {
			if (oColorType) {
				return oColorType.toString();
			}
			return null;
		}

		function readCVml_1_65536(reader) {
			let sValue = reader.GetValue();
			let nLen = sValue.length;
			if (nLen <= 0)
				return 0;

			let bFraction = ('f' === sValue.charAt(nLen - 1));

			if (bFraction) {
				let strValue = sValue.substr(0, nLen - 1);
				let nValue = strValue.length === 0 ? 0 : parseInt(strValue);

				return Math.max(0.0, Math.min(65536.0, nValue)) / 65536.0;
			} else {
				let dValue = sValue.length === 0 ? 0 : parseFloat(sValue);
				return Math.max(0.0, Math.min(1.0, dValue));
			}
		}

		function getCVml_1_65536(dVal) {
			if (!AscFormat.isRealNumber(dVal)) {
				return null;
			}
			return dVal.toString();
		}


		let EEditAs =
			{
				editasBullseye: 0,
				editasCanvas: 1,
				editasCycle: 2,
				editasOrgchart: 3,
				editasRadial: 4,
				editasStacked: 5,
				editasVenn: 6
			};

		function readEditAs(reader) {
			let sValue = reader.GetValue();
			if ("bullseye" === sValue) return EEditAs.editasBullseye;
			else if ("canvas" === sValue) return EEditAs.editasCanvas;
			else if ("cycle" === sValue) return EEditAs.editasCycle;
			else if ("orgchart" === sValue) return EEditAs.editasOrgchart;
			else if ("radial" === sValue) return EEditAs.editasRadial;
			else if ("stacked" === sValue) return EEditAs.editasStacked;
			else if ("venn" === sValue) return EEditAs.editasVenn;
			return EEditAs.editasCanvas;
		}


		function getEditAs(nType) {
			if (nType === EEditAs.editasBullseye) return "bullseye";
			if (nType === EEditAs.editasCanvas) return "canvas";
			if (nType === EEditAs.editasCycle) return "cycle";
			if (nType === EEditAs.editasOrgchart) return "orgchart";
			if (nType === EEditAs.editasRadial) return "radial";
			if (nType === EEditAs.editasStacked) return "stacked";
			if (nType === EEditAs.editasVenn) return "venn";
			return "canvas";
		}


		function CVml_TableLimits(sValue) {
			this.m_arrLimits = [];
			if (sValue) {
				this.FromString(sValue);
			}
		}

		CVml_TableLimits.prototype.GetSize = function () {
			return this.m_arrLimits.length;
		};
		CVml_TableLimits.prototype.GetAt = function (nIndex) {
			if (nIndex < 0 || nIndex >= this.m_arrLimits.length)
				return 0;

			return this.m_arrLimits[nIndex];
		};
		CVml_TableLimits.prototype.AddValue = function (dValue) {
			this.m_arrLimits.push(dValue);
		};
		CVml_TableLimits.prototype.FromString = function (sValue) {
			let nPos = 0;
			let nLen = sValue.length;

			let nSpacePos = 0;
			let wChar;
			while (nPos < nLen) {
				while (' ' === (wChar = sValue.charAt(nPos))) {
					nPos++;
					if (nPos >= nLen)
						return 0;
				}

				nSpacePos = sValue.indexOf(" ", nPos);
				if (-1 === nSpacePos)
					nSpacePos = nLen;

				let oPoint = new CPoint(sValue.substr(nPos, nSpacePos - nPos));
				nPos = nSpacePos + 1;

				this.m_arrLimits.push(oPoint.ToPoints());
			}

			return 0;
		};
		CVml_TableLimits.prototype.ToString = function () {
			let sResult = "";

			for (let nIndex = 0; nIndex < this.m_arrLimits.length; nIndex++) {
				sResult += (this.m_arrLimits[nIndex] + "pt ");
			}

			return sResult;
		}

		function CVml_TableProperties(sValue) {

			this.m_eValue = 0;
			if (sValue) {
				this.FromString(sValue);
			}
		}


		CVml_TableProperties.prototype.FromString = function (sValue) {
			this.m_eValue = parseInt(sValue);
			return this.m_eValue;
		};

		CVml_TableProperties.prototype.ToString = function () {
			return this.m_eValue + "";
		};

		CVml_TableProperties.prototype.IsTable = function () {
			return (this.m_eValue & 1 ? true : false);
		};

		CVml_TableProperties.prototype.IsPlaceholder = function () {
			return (this.m_eValue & 2 ? true : false);
		};

		CVml_TableProperties.prototype.IsBiDirectionalText = function () {
			return (this.m_eValue & 4 ? true : false);
		};


		let EVmlAngle =
			{
				vmlangle30: 0,
				vmlangle45: 1,
				vmlangle60: 2,
				vmlangle90: 3,
				vmlangleAny: 4,
				vmlangleAuto: 5
			};

		function getAngle(nType) {
			if (nType === EVmlAngle.vmlangle30) return "30";
			if (nType === EVmlAngle.vmlangle45) return "45";
			if (nType === EVmlAngle.vmlangle60) return "60";
			if (nType === EVmlAngle.vmlangle90) return "90";
			if (nType === EVmlAngle.vmlangleAny) return "any";
			if (nType === EVmlAngle.vmlangleAuto) return "auto";
			return null;
		}


		let EExt =
			{
				extBackwardCompatible: 0,
				extEdit: 1,
				extView: 2,
			};

		let EVmlCalloutType =
			{
				vmlcallouttypeRectangle: 0,
				vmlcallouttypeRoundRectangle: 1,
				vmlcallouttypeOval: 2,
				vmlcallouttypeCloud: 3
			};


		function getCalloutType(nType) {
			if (nType === EVmlCalloutType.vmlcallouttypeRectangle) return "rectangle";
			if (nType === EVmlCalloutType.vmlcallouttypeRoundRectangle) return "roundrectangle";
			if (nType === EVmlCalloutType.vmlcallouttypeOval) return "oval";
			if (nType === EVmlCalloutType.vmlcallouttypeCloud) return "cloud";
			return null;
		}


		let EExtrusionType =
			{
				extrusiontypeParallel: 0,
				extrusiontypePerspective: 1
			};

		function readExt(reader) {
			let sVal = reader.GetValue();
			switch (sVal) {
				case "backwardCompatible": {
					return EExt.extBackwardCompatible;
				}
				case "edit": {
					return EExt.extEdit;
				}
				case "view": {
					return EExt.extView;
				}
			}
			return null;
		}

		function getExt(nType) {
			switch (nType) {
				case EExt.extBackwardCompatible: {
					return "backwardCompatible";
				}
				case EExt.extEdit: {
					return "edit";
				}
				case EExt.extView: {
					return "view";
				}
			}
			return null;
		}

		function readExtrusionType(reader) {
			let sVal = reader.GetValue();
			switch (sVal) {
				case "parallel":
					return EExtrusionType.extrusiontypeParallel;
				case "extrusiontypeParallel":
					return EExtrusionType.extrusiontypePerspective;
			}
			return EExtrusionType.extrusiontypeParallel;
		}

		function getExtrusionType(nType) {
			switch (nType) {
				case EExtrusionType.extrusiontypeParallel:
					return "parallel";
				case EExtrusionType.extrusiontypePerspective:
					return "extrusiontypeParallel";
			}
			return null;
		}


		let EFillType =
			{
				filltypeBackground: 0,
				filltypeFrame: 1,
				filltypeGradient: 2,
				filltypeGradientCenter: 3,
				filltypeGradientRadial: 4,
				filltypeGradientUnscaled: 5,
				filltypePattern: 6,
				filltypeSolid: 7,
				filltypeTile: 8
			};

		function readFillType(reader) {
			let sVal = reader.GetValue();
			switch (sVal) {
				case "background":
					return EFillType.filltypeBackground;
				case "frame":
					return EFillType.filltypeFrame;
				case "gradient":
					return EFillType.filltypeGradient;
				case "gradientCenter":
					return EFillType.filltypeGradientCenter;
				case "gradientRadial":
					return EFillType.filltypeGradientRadial;
				case "gradientUnscaled":
					return EFillType.filltypeGradientUnscaled;
				case "pattern":
					return EFillType.filltypePattern;
				case "solid":
					return EFillType.filltypeSolid;
				case "tile":
					return EFillType.filltypeTile;
			}
			return EFillType.filltypeSolid;
		}

		function getFillType(nType) {
			switch (nType) {
				case EFillType.filltypeBackground:
					return "background";
				case EFillType.filltypeFrame:
					return "frame";
				case EFillType.filltypeGradient:
					return "gradient";
				case EFillType.filltypeGradientCenter:
					return "gradientCenter";
				case EFillType.filltypeGradientRadial:
					return "gradientRadial";
				case EFillType.filltypeGradientUnscaled:
					return "gradientUnscaled";
				case EFillType.filltypePattern:
					return "pattern";
				case EFillType.filltypeSolid:
					return "solid";
				case EFillType.filltypeTile:
					return "tile";
			}
			return null;
		}


		let EVml_Vector2D_Position =
			{
				vmlvector2dposConstant: 0,
				vmlvector2dposFormula: 1,
				vmlvector2dposAdjValue: 2,
				vmlvector2dposCenter: 3,
				vmlvector2dposTopLeft: 4,
				vmlvector2dposBottomRight: 5
			};

		function CVml_Vector2D_Position(sVal) {
			this.m_eTypeX = null;
			this.m_eTypeY = null;

			this.m_sIdX = null;
			this.m_sIdY = null;

			this.m_dX = null;
			this.m_dY = null;
			if (sVal) {
				this.fromString(sVal);
			}
		}

		CVml_Vector2D_Position.prototype.fromString = function (sValue) {
			let nLen = sValue.length;
			if (nLen <= 0)
				return 0;

			let nPos = sValue.indexOf(",");
			if (-1 === nPos)
				return 0;

			let sFirst = sValue.substr(0, nPos);
			let sSecond = sValue.substr(nPos + 1, nLen - nPos - 1);

			this.parse(sFirst, true);
			this.parse(sSecond, false);

			return 0;
		};
		CVml_Vector2D_Position.prototype.parse = function (sValue, bFirst) {
			let eValue = EVml_Vector2D_Position.vmlvector2dposConstant;
			let dValue = 0.0;
			let sId;
			sValue.replace(" ", "");
			let nLen = sValue.length;
			if (nLen > 0) {

				let wChar = sValue[0];
				switch (wChar) {
					case '0':
					case '1':
					case '2':
					case '3':
					case '4':
					case '5':
					case '6':
					case '7':
					case '8':
					case '9':
					case '.':
						eValue = EVml_Vector2D_Position.vmlvector2dposConstant;
						dValue = sValue.length === 0 ? 0 : parseFloat(sValue);
						break;
					case 'c':

						if ("center" === sValue)
							eValue = EVml_Vector2D_Position.vmlvector2dposCenter;
						break;
					case 't':
						if ("topleft" === sValue)
							eValue = EVml_Vector2D_Position.vmlvector2dposTopLeft;
						break;
					case 'b':
						if ("bottomright" === sValue)
							eValue = EVml_Vector2D_Position.vmlvector2dposBottomRight;
						break;
					case '@':
						eValue = EVml_Vector2D_Position.vmlvector2dposFormula;
						sId = sValue.substr(1, nLen - 1);
						break;
					case '#':
						eValue = EVml_Vector2D_Position.vmlvector2dposAdjValue;
						sId = sValue.substr(1, nLen - 1);
						break;
				}
			}
			if (bFirst) {
				this.m_eTypeX = eValue;
				this.m_sIdX = sId;
				this.m_dX = dValue;
			} else {
				this.m_eTypeY = eValue;
				this.m_sIdY = sId;
				this.m_dY = dValue;
			}
		};
		CVml_Vector2D_Position.prototype.GetTypeX = function () {
			return this.m_eTypeX;
		};
		CVml_Vector2D_Position.prototype.GetTypeY = function () {
			return this.m_eTypeY;
		};
		CVml_Vector2D_Position.prototype.ToString = function () {
			let sResult = "";

			switch (this.m_eTypeX) {
				case EVml_Vector2D_Position.vmlvector2dposConstant    :
					sResult = this.m_dX;
					break;
				case EVml_Vector2D_Position.vmlvector2dposFormula     :
					sResult = ("@") + this.m_sIdX;
					break;
				case EVml_Vector2D_Position.vmlvector2dposAdjValue    :
					sResult = ("#") + this.m_sIdX;
					break;
				case EVml_Vector2D_Position.vmlvector2dposCenter      :
					sResult = ("center");
					break;
				case EVml_Vector2D_Position.vmlvector2dposTopLeft     :
					sResult = ("topleft");
					break;
				case EVml_Vector2D_Position.vmlvector2dposBottomRight :
					sResult = ("bottomright");
					break;
			}

			sResult += (",");

			switch (this.m_eTypeY) {
				case EVml_Vector2D_Position.vmlvector2dposConstant    : {
					let sTemp = "" + (this.m_dY);
					sResult += sTemp;
					break;
				}
				case EVml_Vector2D_Position.vmlvector2dposFormula     :
					sResult += ("@") + this.m_sIdY;
					break;
				case EVml_Vector2D_Position.vmlvector2dposAdjValue    :
					sResult += ("#") + this.m_sIdY;
					break;
				case EVml_Vector2D_Position.vmlvector2dposCenter      :
					sResult += ("center");
					break;
				case EVml_Vector2D_Position.vmlvector2dposTopLeft     :
					sResult += ("topleft");
					break;
				case EVml_Vector2D_Position.vmlvector2dposBottomRight :
					sResult += ("bottomright");
					break;
			}

			return sResult;
		}

		function CVml_Polygon2D_Units(val) {
			this.m_arrPoints = [];
			this.m_wcDelimiter = " ";
			if (val) {
				if (val instanceof CVml_Polygon2D_Units) {
					this.m_wcDelimiter = " ";
					this.FromString(val.ToString());
				} else {
					this.FromString(val);
				}
			}
		}

		CVml_Polygon2D_Units.prototype.SetDelimiter = function (wcNew) {
			this.m_wcDelimiter = wcNew;
		};
		CVml_Polygon2D_Units.prototype.GetX = function (nIndex) {
			if (nIndex < 0 || nIndex >= this.m_arrPoints.length)
				return 0;

			return this.m_arrPoints[nIndex].nX;
		};
		CVml_Polygon2D_Units.prototype.GetY = function (nIndex) {
			if (nIndex < 0 || nIndex >= this.m_arrPoints.length)
				return 0;

			return this.m_arrPoints[nIndex].nY;
		};

		CVml_Polygon2D_Units.prototype.AddPoint = function (dX, dY) {
			let oPt = new TPoint(dX, dY);
			this.m_arrPoints.push(oPt);
		};

		CVml_Polygon2D_Units.prototype.FromString = function (sValue) {
			this.m_arrPoints.length = 0;

			let nLen = sValue.length;
			if (nLen <= 0)
				return 0;

			let nStartPos = 0;
			while (true) {
				let nMidPos = sValue.indexOf(",", nStartPos);
				let nEndPos = sValue.indexOf(this.m_wcDelimiter, nMidPos + 1);

				if (-1 === nMidPos)
					break;

				if (-1 === nEndPos)
					nEndPos = nLen;

				let strX = sValue.substr(nStartPos, nMidPos - nStartPos);
				let strY = sValue.substr(nMidPos + 1, nEndPos - nMidPos - 1);

				strX.replace("@", "");
				strY.replace("@", "");

				let nX = strX.length === 0 ? 0 : parseInt(strX);
				let nY = strY.length === 0 ? 0 : parseInt(strY);

				this.m_arrPoints.push(new TPoint(nX, nY));

				nStartPos = nEndPos + 1;
			}


			return 0;
		};

		CVml_Polygon2D_Units.prototype.ToString = function () {
			let sResult = "";

			for (let nIndex = 0; nIndex < this.m_arrPoints.length; nIndex++) {
				sResult += this.m_arrPoints[nIndex].nX + "," + this.m_arrPoints[nIndex].nY;
				if (nIndex < this.m_arrPoints.length - 1)
					sResult += this.m_wcDelimiter;
			}

			return sResult;
		};


		function CVml_Polygon2D(sValue) {
			this.m_arrPoints = [];
			if (sValue) {
				this.FromString(sValue);
			}
		}


		CVml_Polygon2D.prototype.GetSize = function () {
			return this.m_arrPoints.length;
		}

		CVml_Polygon2D.prototype.GetX = function (nIndex) {
			if (nIndex < 0 || nIndex >= this.m_arrPoints.length)
				return 0;

			return this.m_arrPoints[nIndex].nX;
		}
		CVml_Polygon2D.prototype.GetY = function (nIndex) {
			if (nIndex < 0 || nIndex >= this.m_arrPoints.length)
				return 0;

			return this.m_arrPoints[nIndex].nY;
		}

		CVml_Polygon2D.prototype.AddPoint = function (nX, nY) {
			let oPt = new TPoint(nX, nY);
			this.m_arrPoints.push(oPt);
		}

		CVml_Polygon2D.prototype.FromString = function (sValue) {
			this.m_arrPoints.length = 0;

			let nLen = sValue.length;
			if (nLen <= 0)
				return 0;

			let nStartPos = 0;
			while (true) {
				let nMidPos = sValue.indexOf(",", nStartPos);
				let nEndPos = sValue.indexOf(",", nMidPos + 1);

				if (-1 === nMidPos)
					break;

				if (-1 === nEndPos)
					nEndPos = nLen;

				let strX = sValue.substr(nStartPos, nMidPos - nStartPos);
				let strY = sValue.substr(nStartPos, nMidPos - nStartPos);

				strX.replace("@", "");
				strY.replace("@", "");

				let nX = strX.length === 0 ? 0 : parseInt(strX);
				let nY = strY.length === 0 ? 0 : parseInt(strY);

				this.m_arrPoints.push(new TPoint(nX, nY));

				nStartPos = nEndPos + 1;
			}


			return 0;
		}

		CVml_Polygon2D.prototype.ToString = function () {
			let sResult = "";

			for (let nIndex = 0; nIndex < this.m_arrPoints.length; nIndex++) {
				let sTemp = this.m_arrPoints[nIndex].nX + "," + this.m_arrPoints[nIndex].nY;
				if (nIndex < this.m_arrPoints.length - 1) sTemp += ",";
				sResult += sTemp;
			}

			return sResult;
		};

		function TPoint(n_X, n_Y) {
			this.nX = n_X;
			this.nY = n_Y;
		}


		function CVml_Vector2D_Units(sValue) {
			this.m_dX = 0;
			this.m_dY = 0;
			if (sValue) {
				this.fromString(sValue);
			}
		}

		CVml_Vector2D_Units.prototype.fromString = function (sValue) {
			this.m_dX = 0;
			this.m_dY = 0;

			let nLen = sValue.length;
			if (nLen <= 0)
				return 0;

			let nPos = sValue.indexOf(",");
			if (-1 === nPos) {//only x position
				let oPt1 = new CPoint(sValue);
				this.m_dX = oPt1.GetValue();
				return 0;
			}


			let oPt1 = new CPoint(sValue.substr(0, nPos));
			this.m_dX = oPt1.GetValue();


			let oPt2 = new CPoint(sValue.substr(nPos + 1, nLen - nPos - 1));
			this.m_dY = oPt2.GetValue();
			return 0;
		};


		CVml_Vector2D_Units.prototype.ToString = function () {
			return "" + this.m_dX + "," + this.m_dY;
		};


		let EConnectType =
			{
				connecttypeCustom: 0,
				connecttypeNone: 1,
				connecttypeRect: 2,
				connecttypeSegments: 3
			};

		function readConnectType(reader) {
			let sVal = reader.GetValue();
			switch (sVal) {
				case "custom": {
					return EConnectType.connecttypeCustom;
				}
				case "none": {
					return EConnectType.connecttypeNone;
				}
				case "rect": {
					return EConnectType.connecttypeRect;
				}
				case "segments": {
					return EConnectType.connecttypeSegments;
				}
			}
		}

		function getConnectType(nType) {
			switch (nType) {
				case EConnectType.connecttypeCustom: {
					return "custom";
				}
				case EConnectType.connecttypeNone: {
					return "none";
				}
				case EConnectType.connecttypeRect: {
					return "rect";
				}
				case EConnectType.connecttypeSegments: {
					return "segments";
				}
			}
			return null;
		}


		let EShadowType =
			{
				shadowtypeDouble: 1,
				shadowtypeEmboss: 2,
				shadowtypePerspective: 3,
				shadowtypeSingle: 4
			};


		function readShadowType(reader) {
			let sVal = reader.GetValue();
			switch (sVal) {
				case "double" : {
					return EShadowType.shadowtypeDouble;
				}
				case "emboss" : {
					return EShadowType.shadowtypeEmboss;
				}
				case "perspective" : {
					return EShadowType.shadowtypePerspective;
				}
				case "single" : {
					return EShadowType.shadowtypeSingle;
				}
			}
			return EShadowType.shadowtypeSingle;
		}

		function getShadowType(nType) {
			switch (nType) {
				case EShadowType.shadowtypeDouble : {
					return "double";
				}
				case EShadowType.shadowtypeEmboss: {
					return "emboss";
				}
				case EShadowType.shadowtypePerspective: {
					return "perspective";
				}
				case  EShadowType.shadowtypeSingle: {
					return "single";
				}
			}
			return null;
		}


		let EVmlDashStyle =
			{
				vmldashstyleSolid: 0,
				vmldashstyleShortDash: 1,
				vmldashstyleShortDot: 2,
				vmldashstyleShortDashDot: 3,
				vmldashstyleShortDashDotDot: 4,
				vmldashstyleDot: 5,
				vmldashstyleDash: 6,
				vmldashstyleLongDash: 7,
				vmldashstyleDashDot: 8,
				vmldashstyleLongDashDot: 9,
				vmldashstyleLongDashDotDot: 10,
				vmldashstyleCustom: 11
			};

		function readDashStyle(reader) {
			let sValue = reader.GetValue();
			if ("solid" === sValue) return EVmlDashStyle.vmldashstyleSolid;
			else if ("shortdash" === sValue) return EVmlDashStyle.vmldashstyleShortDash;
			else if ("shortdot" === sValue) return EVmlDashStyle.vmldashstyleShortDot;
			else if ("shortdashdot" === sValue) return EVmlDashStyle.vmldashstyleShortDashDot;
			else if ("shortdashdotdot" === sValue) return EVmlDashStyle.vmldashstyleShortDashDotDot;
			else if ("dot" === sValue) return EVmlDashStyle.vmldashstyleDot;
			else if ("dash" === sValue) return EVmlDashStyle.vmldashstyleDash;
			else if ("dashdot" === sValue) return EVmlDashStyle.vmldashstyleDashDot;
			else if ("longdash" === sValue) return EVmlDashStyle.vmldashstyleLongDash;
			else if ("longdashdot" === sValue) return EVmlDashStyle.vmldashstyleLongDashDot;
			else if ("longdashdotdot" === sValue) return EVmlDashStyle.vmldashstyleLongDashDotDot;
			return EVmlDashStyle.vmldashstyleSolid;
		}

		function getDashStyle(nType) {
			if (EVmlDashStyle.vmldashstyleSolid === nType) return "solid";
			else if (EVmlDashStyle.vmldashstyleShortDash === nType) return "shortdash";
			else if (EVmlDashStyle.vmldashstyleShortDot === nType) return "shortdot";
			else if (EVmlDashStyle.vmldashstyleShortDashDot === nType) return "shortdashdot";
			else if (EVmlDashStyle.vmldashstyleShortDashDotDot === nType) return "shortdashdotdot";
			else if (EVmlDashStyle.vmldashstyleDot === nType) return "dot";
			else if (EVmlDashStyle.vmldashstyleDash === nType) return "dash";
			else if (EVmlDashStyle.vmldashstyleDashDot === nType) return "dashdot";
			else if (EVmlDashStyle.vmldashstyleLongDash === nType) return "longdash";
			else if (EVmlDashStyle.vmldashstyleLongDashDot === nType) return "longdashdot";
			else if (EVmlDashStyle.vmldashstyleLongDashDotDot === nType) return "longdashdotdot";
			return null;
		}

		let EStrokeArrowLength =
			{
				strokearrowlengthLong: 0,
				strokearrowlengthMedium: 1,
				strokearrowlengthShort: 2
			};

		function readArrowLength(reader) {
			let sValue = reader.GetValue();
			if ("long" === sValue) return EStrokeArrowLength.strokearrowlengthLong;
			else if ("medium" === sValue) return EStrokeArrowLength.strokearrowlengthMedium;
			else if ("short" === sValue) return EStrokeArrowLength.strokearrowlengthShort;
			return EStrokeArrowLength.strokearrowlengthMedium;
		}

		function getArrowLength(nType) {
			if (EStrokeArrowLength.strokearrowlengthLong === nType) return "long";
			else if (EStrokeArrowLength.strokearrowlengthMedium === nType) return "medium";
			else if (EStrokeArrowLength.strokearrowlengthShort === nType) return "short";
			return null;
		}


		let EStrokeArrowType =
			{
				strokearrowtypeBlock: 0,
				strokearrowtypeClassic: 1,
				strokearrowtypeDiamond: 2,
				strokearrowtypeNone: 3,
				strokearrowtypeOpen: 4,
				strokearrowtypeOval: 5
			};

		function readArrowType(reader) {
			let sValue = reader.GetValue();
			if ("block" === sValue) return EStrokeArrowType.strokearrowtypeBlock;
			else if ("classic" === sValue) return EStrokeArrowType.strokearrowtypeClassic;
			else if ("diamond" === sValue) return EStrokeArrowType.strokearrowtypeDiamond;
			else if ("none" === sValue) return EStrokeArrowType.strokearrowtypeNone;
			else if ("open" === sValue) return EStrokeArrowType.strokearrowtypeOpen;
			else if ("oval" === sValue) return EStrokeArrowType.strokearrowtypeOval;
			return EStrokeArrowType.strokearrowtypeNone;
		}

		function getArrowType(nType) {
			if (EStrokeArrowType.strokearrowtypeBlock === nType) return "block";
			else if (EStrokeArrowType.strokearrowtypeClassic === nType) return "classic";
			else if (EStrokeArrowType.strokearrowtypeDiamond === nType) return "diamond";
			else if (EStrokeArrowType.strokearrowtypeNone === nType) return "none";
			else if (EStrokeArrowType.strokearrowtypeOpen === nType) return "open";
			else if (EStrokeArrowType.strokearrowtypeOval === nType) return "oval";
			return null;
		}


		let EStrokeArrowWidth =
			{
				strokearrowwidthMedium: 0,
				strokearrowwidthNarrow: 1,
				strokearrowwidthWide: 2
			};

		function readArrowWidth(reader) {
			let sValue = reader.GetValue();
			if ("medium" === sValue) return EStrokeArrowWidth.strokearrowwidthMedium;
			else if ("narrow" === sValue) return EStrokeArrowWidth.strokearrowwidthNarrow;
			else if ("wide" === sValue) return EStrokeArrowWidth.strokearrowwidthWide;
			return EStrokeArrowWidth.strokearrowwidthMedium;
		}

		function getArrowWidth(nType) {
			if (EStrokeArrowWidth.strokearrowwidthMedium === nType) return "medium";
			else if (EStrokeArrowWidth.strokearrowwidthNarrow === nType) return "narrow";
			else if (EStrokeArrowWidth.strokearrowwidthWide === nType) return "wide";
			return null;
		}


		//--------------------------------------------------------------------------------
		let EStrokeEndCap =
			{
				strokeendcapFlat: 0,
				strokeendcapRound: 1,
				strokeendcapSqaure: 2
			};


		function readEndCap(reader) {
			let sValue = reader.GetValue();
			if ("flat" === sValue) return EStrokeEndCap.strokeendcapFlat;
			else if ("round" === sValue) return EStrokeEndCap.strokeendcapRound;
			else if ("square" === sValue) return EStrokeEndCap.strokeendcapSqaure;
			return EStrokeEndCap.strokeendcapRound;
		}

		function getEndCap(nType) {
			if (EStrokeEndCap.strokeendcapFlat === nType) return "flat";
			else if (EStrokeEndCap.strokeendcapRound === nType) return "round";
			else if (EStrokeEndCap.strokeendcapSqaure === nType) return "square";
			return null;
		}


		let EStrokeJoinStyle =
			{
				strokejoinstyleBevel: 0,
				strokejoinstyleMiter: 1,
				strokejoinstyleRound: 2
			};

		function readJoinStyle(reader) {
			let sValue = reader.GetValue();
			if ("round" === sValue) return EStrokeJoinStyle.strokejoinstyleRound;
			else if ("miter" === sValue) return EStrokeJoinStyle.strokejoinstyleMiter;
			else if ("bevel" === sValue) return EStrokeJoinStyle.strokejoinstyleBevel;
			return EStrokeJoinStyle.strokejoinstyleRound;
		}

		function getJoinStyle(nType) {
			if (EStrokeJoinStyle.strokejoinstyleRound === nType) return "round";
			else if (EStrokeJoinStyle.strokejoinstyleMiter === nType) return "miter";
			else if (EStrokeJoinStyle.strokejoinstyleBevel === nType) return "bevel";
			return null;
		}


		let EStrokeLineStyle =
			{
				strokelinestyleSingle: 0,
				strokelinestyleThickBetweenThin: 1,
				strokelinestyleThickThin: 2,
				strokelinestyleThinThick: 3,
				strokelinestyleThinThin: 4
			};

		function readLineStyle(reader) {
			let sValue = reader.GetValue();
			if ("single" === sValue) return EStrokeLineStyle.strokelinestyleSingle;
			else if ("thickBetweenThin" === sValue) return EStrokeLineStyle.strokelinestyleThickBetweenThin;
			else if ("thickThin" === sValue) return EStrokeLineStyle.strokelinestyleThickThin;
			else if ("thinThick" === sValue) return EStrokeLineStyle.strokelinestyleThinThick;
			else if ("thinThin" === sValue) return EStrokeLineStyle.strokelinestyleThinThin;
			return EStrokeLineStyle.strokelinestyleSingle;
		}

		function getLineStyle(nType) {
			if (EStrokeLineStyle.strokelinestyleSingle === nType) return "single";
			else if (EStrokeLineStyle.strokelinestyleThickBetweenThin === nType) return "thickBetweenThin";
			else if (EStrokeLineStyle.strokelinestyleThickThin === nType) return "thickThin";
			else if (EStrokeLineStyle.strokelinestyleThinThick === nType) return "thinThick";
			else if (EStrokeLineStyle.strokelinestyleThinThin === nType) return "thinThin";
			return null;
		}


		let EImageAspect =
			{
				imageaspectAtLeast: 0,
				imageaspectAtMost: 1,
				imageaspectIgnore: 2
			};

		function readImageAspect(reader) {
			let sValue = reader.GetValue();
			if ("atLeast" === sValue) return EImageAspect.imageaspectAtLeast;
			else if ("atMost" === sValue) return EImageAspect.imageaspectAtMost;
			else if ("ignore" === sValue) return EImageAspect.imageaspectIgnore;
			return EImageAspect.imageaspectIgnore;
		}

		function getImageAspect(nType) {
			if (EImageAspect.imageaspectAtLeast === nType) return "atLeast";
			else if (EImageAspect.imageaspectAtMost === nType) return "atMost";
			else if (EImageAspect.imageaspectIgnore === nType) return "ignore";
			return null;
		}


		function CVml_TextBoxInset(sValue) {

			this.m_dLeft = 0;
			this.m_dTop = 0;
			this.m_dRight = 0;
			this.m_dBottom = 0;
			if (sValue) {
				this.FromString(sValue);
			}
		}

		CVml_TextBoxInset.prototype.GetLeft = function (nIndex) {
			return this.m_dLeft;
		};
		CVml_TextBoxInset.prototype.GetTop = function (nIndex) {
			return this.m_dTop;
		};
		CVml_TextBoxInset.prototype.GetRight = function (nIndex) {
			return this.m_dRight;
		};
		CVml_TextBoxInset.prototype.GetBottom = function (nIndex) {
			return this.m_dBottom;
		};
		CVml_TextBoxInset.prototype.Set = function (dL, dT, dR, dB) {
			this.m_dLeft = dL;
			this.m_dTop = dT;
			this.m_dRight = dR;
			this.m_dBottom = dB;
		};
		CVml_TextBoxInset.prototype.FromString = function (sValue) {
			this.Set(0, 0, 0, 0);

			if (sValue.length === 0)
				return 0;

			let arSplit;

			sValue.replace("@", "");
			arSplit = sValue.split(", ");

			if (arSplit.length > 0) {
				let oPt = new CPoint(arSplit[0]);
				this.m_dLeft = oPt.GetValue();
			}

			if (arSplit.length > 1) {
				let oPt = new CPoint(arSplit[1]);
				this.m_dTop = oPt.GetValue();
			}

			if (arSplit.length > 2) {
				let oPt = new CPoint(arSplit[2]);
				this.m_dRight = oPt.GetValue();
			}

			if (arSplit.length > 3) {
				let oPt = new CPoint(arSplit[3]);
				this.m_dBottom = oPt.GetValue();
			}

			return 0;
		};
		CVml_TextBoxInset.prototype.ToString = function () {
			return "" + this.m_dLeft + "pt," + this.m_dTop + "pt," + this.m_dRight + "pt," + this.m_dBottom + "pt";
		};


		function CCssStyle(sValue) {
			this.m_arrProperties = [];
			this.m_sCss = null;
			if (sValue) {
				this.FromString(sValue);
			}
		}

		CCssStyle.prototype.Clear = function () {
			this.m_arrProperties.length = 0;
		}

		CCssStyle.prototype.FromString = function (sValue) {
			this.Clear();

			this.m_sCss = sValue;
			this.ParseProperties();

			return this.m_sCss;
		};

		CCssStyle.prototype.ToString = function () {
			return this.m_sCss;
		};

		CCssStyle.prototype.ParseProperties = function () {
			let sTemp = this.m_sCss;
			while (sTemp.length > 0) {
				let nPos = sTemp.indexOf(';');
				if (-1 === nPos) {
					let oProperty = new CCssProperty(sTemp);
					if ((ECssPropertyType.cssptUnknown !== oProperty.get_Type())) {
						this.m_arrProperties.push(oProperty);
					}
					sTemp = "";

				} else {
					let oProperty = new CCssProperty(sTemp.substr(0, nPos));
					if ((ECssPropertyType.cssptUnknown !== oProperty.get_Type())) {
						this.m_arrProperties.push(oProperty);
					}
					sTemp = sTemp.substr(nPos + 1, sTemp.length - nPos - 1);
				}
			}
			return true;
		};


		let ECssPropertyType =
			{
				cssptUnknown: 0,
				cssptFlip: 1000,
				cssptHeight: 1001,
				cssptLeft: 1002,
				cssptMarginBottom: 1003,
				cssptMarginLeft: 1004,
				cssptMarginRight: 1005,
				cssptMarginTop: 1006,
				cssptMsoPositionHorizontal: 1007,
				cssptMsoPositionHorizontalRelative: 1008,
				cssptMsoPositionVertical: 1009,
				cssptMsoPositionVerticalRelative: 1010,
				cssptMsoWrapDistanceBottom: 1011,
				cssptMsoWrapDistanceLeft: 1012,
				cssptMsoWrapDistanceRight: 1013,
				cssptMsoWrapDistanceTop: 1014,
				cssptMsoWrapEdited: 1015,
				cssptMsoWrapStyle: 1016,
				cssptPosition: 1017,
				cssptRotation: 1018,
				cssptTop: 1019,
				cssptVisibility: 1020,
				cssptWidth: 1021,
				cssptZIndex: 1022,
				csspctMsoWidthPercent: 1023,
				csspctMsoHeightPercent: 1024,

				cssptDirection: 1100,
				cssptLayoutFlow: 1101,
				cssptMsoDirectionAlt: 1102,
				cssptMsoFitShapeToText: 1103,
				cssptMsoFitTextToShape: 1104,
				cssptMsoLayoutFlowAlt: 1105,
				cssptMsoNextTextbox: 1106,
				cssptMsoRotate: 1107,
				cssptMsoTextScale: 1108,
				cssptVTextAnchor: 1109,

				cssptFont: 1200,
				cssptFontFamily: 1201,
				cssptFontSize: 1202,
				cssptFontStyle: 1203,
				cssptFontVariant: 1204,
				cssptFontWeight: 1205,
				cssptMsoTextShadow: 1206,
				cssptTextDecoration: 1207,
				cssptVRotateLetters: 1208,
				cssptVSameLetterHeights: 1209,
				cssptVTextAlign: 1210,
				cssptVTextKern: 1211,
				cssptVTextReverse: 1212,
				cssptVTextSpacingMode: 1213,
				cssptVTextSpacing: 1214,
				cssptHTextAlign: 1215
			};


		let ECssFlip =
			{
				cssflipX: 0,
				cssflipY: 1,
				cssflipXY: 2,
				cssflipYX: 3
			};
		let ECssUnitsType =
			{
				cssunitstypeAuto: 0,
				cssunitstypeUnits: 1,
				cssunitstypePerc: 2,
				cssunitstypeAbsolute: 3
			};


		let ECssMsoPosHorRel =
			{
				cssmsoposhorrelMargin: 0,
				cssmsoposhorrelPage: 1,
				cssmsoposhorrelText: 2,
				cssmsoposhorrelChar: 3,
				cssmsoposhorrelLeftMargin: 4,
				cssmsoposhorrelRightMargin: 5
			};
		let ECssMsoPosHor =
			{
				cssmsoposhorAbsolute: 0,
				cssmsoposhorLeft: 1,
				cssmsoposhorCenter: 2,
				cssmsoposhorRight: 3,
				cssmsoposhorInside: 4,
				cssmsoposhorOutside: 5
			};
		let ECssMsoPosVer =
			{
				cssmsoposverAbsolute: 0,
				cssmsoposverTop: 1,
				cssmsoposverCenter: 2,
				cssmsoposverBottom: 3,
				cssmsoposverInside: 4,
				cssmsoposverOutside: 5
			};
		let ECssMsoPosVerRel =
			{
				cssmsoposverrelMargin: 0,
				cssmsoposverrelPage: 1,
				cssmsoposverrelText: 2,
				cssmsoposverrelLine: 3,
				cssmsoposverrelTopMargin: 4,
				cssmsoposverrelBottomMargin: 5
			};
		let ECssMsoWrapStyle =
			{
				cssmsowrapstyleSqaure: 0,
				cssmsowrapstyleNone: 1,
			};
		let ECssPosition =
			{
				csspositionStatic: 0,
				csspositionAbsolute: 1,
				csspositionRelative: 2,
			};
		let ECssVisibility =
			{
				cssvisibilityHidden: 0,
				cssvisibilityInherit: 1,
			};
		let ECssZIndexType =
			{
				csszindextypeAuto: 0,
				csszindextypeOrder: 1
			};


		let ECssDirection =
			{
				cssdirectionLTR: 0,
				cssdirectionRTL: 1
			};
		let ECssLayoutFlow =
			{
				csslayoutflowHorizontal: 0,
				csslayoutflowVertical: 1,
				csslayoutflowVerticalIdeographic: 2,
				csslayoutflowHorizontalIdeographic: 3
			};
		let ECssDirectionAlt =
			{
				cssdirectionaltContext: 0,
			};
		let ECssLayoutFlowAlt =
			{
				csslayoutflowaltBottomToTop: 0,
			};

		let ECssMsoRotate =
			{
				cssmsorotate0: 0,
				cssmsorotate90: 90,
				cssmsorotate180: 180,
				cssmsorotate270: -90
			};
		let ECssVTextAnchor =
			{
				cssvtextanchorTop: 0,
				cssvtextanchorMiddle: 1,
				cssvtextanchorBottom: 2,
				cssvtextanchorTopCenter: 3,
				cssvtextanchorMiddleCenter: 4,
				cssvtextanchorBottomCenter: 5,
				cssvtextanchorTopBaseline: 6,
				cssvtextanchorBottomBaseline: 7,
				cssvtextanchorTopCenterBaseline: 8,
				cssvtextanchorBottomCenterBaseline: 9
			};
		let ECssFontStyle =
			{
				cssfontstyleNormal: 0,
				cssfontstyleItalic: 1,
				cssfontstyleOblique: 2
			};
		let ECssFontVarian =
			{
				cssfontvariantNormal: 0,
				cssfontvariantSmallCaps: 1,
			};
		let ECssFontWeight =
			{
				cssfontweightNormal: 0,
				cssfontweightLighter: 1,
				cssfontweight100: 100,
				cssfontweight200: 200,
				cssfontweight300: 300,
				cssfontweight400: 400,
				cssfontweightBold: 550,
				cssfontweightBolder: 750,
				cssfontweight500: 500,
				cssfontweight600: 600,
				cssfontweight700: 700,
				cssfontweight800: 800,
				cssfontweight900: 900
			};
		let ECssTextDecoration =
			{
				csstextdecorationNone: 0,
				csstextdecorationUnderline: 1,
				csstextdecorationOverline: 2,
				csstextdecorationLineThrough: 3,
				csstextdecorationBlink: 4,
			};
		let ECssVTextAlign =
			{
				cssvtextalignLeft: 0,
				cssvtextalignRight: 1,
				cssvtextalignCenter: 2,
				cssvtextalignJustify: 3,
				cssvtextalignLetterJustify: 4,
				cssvtextalignStretchJustify: 5,
			};
		let ECssVTextSpacingMode =
			{
				cssvtextspacingmodeTightening: 0,
				cssvtextspacingmodeTracking: 1
			};


		function TCssUnitsValue() {
			this.eType = null;
			this.dValue = null;
		}

		function TCssZIndexValue() {
			this.eType = null;
			this.nOrder = null;
		}

		function UCssValue() {
			this.eFlip = null;
			this.oValue = new TCssUnitsValue();
			this.eMsoPosHor = null;
			this.eMsoPosHorRel = null;
			this.eMsoPosVer = null;
			this.eMsoPosVerRel = null;
			this.bValue = null;
			this.eMsoWrapStyle = null;
			this.ePosition = null;
			this.eVisibility = null;
			this.oZIndex = new TCssZIndexValue();
			this.eDirection = null;
			this.eLayoutFlow = null;
			this.eDirectionAlt = null;
			this.eLayoutFlowAlt = null;
			this.wsValue = null;
			this.eRotate = null;
			this.eVTextAnchor = null;
			this.eFontStyle = null;
			this.eFontVariant = null;
			this.eFontWeight = null;
			this.eTextDecoration = null;
			this.eVTextAlign = null;
			this.eVTextSpacingMode = null;
			this.eHTextAlign = null;
		}

		function CCssProperty(sBuffer) {
			this.m_eType = ECssPropertyType.cssptUnknown;
			this.m_oValue = new UCssValue();
			if (sBuffer) {
				this.Parse(sBuffer);
			}
		}

		CCssProperty.prototype.get_Value = function () {
			return this.m_oValue;
		}

		CCssProperty.prototype.get_Type = function () {
			return this.m_eType;
		}


		CCssProperty.prototype.Parse = function (sBuffer) {
			let nPos = sBuffer.indexOf(':');
			let sValue;

			if (-1 === nPos) {
				this.m_eType = ECssPropertyType.cssptUnknown;
			} else {
				let sProperty = sBuffer.substr(0, nPos);
				sValue = sBuffer.substr(nPos + 1, sBuffer.length - (nPos + 1));

				sProperty = sProperty.replace(/\s/g, "");

				if (sProperty.length <= 2) {
					this.m_eType = ECssPropertyType.cssptUnknown;
					return;
				}

				if ("direction" === sProperty) this.m_eType = ECssPropertyType.cssptDirection;
				if ("flip" === sProperty) this.m_eType = ECssPropertyType.cssptFlip;
				if ("font" === sProperty) this.m_eType = ECssPropertyType.cssptFont;
				else if ("font-family" === sProperty) this.m_eType = ECssPropertyType.cssptFontFamily;
				else if ("font-size" === sProperty) this.m_eType = ECssPropertyType.cssptFontSize;
				else if ("font-style" === sProperty) this.m_eType = ECssPropertyType.cssptFontStyle;
				else if ("font-variant" === sProperty) this.m_eType = ECssPropertyType.cssptFontVariant;
				else if ("font-weight" === sProperty) this.m_eType = ECssPropertyType.cssptFontWeight;
				if ("height" === sProperty) this.m_eType = ECssPropertyType.cssptHeight;
				if ("layout-flow" === sProperty) this.m_eType = ECssPropertyType.cssptLayoutFlow;
				else if ("left" === sProperty) this.m_eType = ECssPropertyType.cssptLeft;
				if ("margin-bottom" === sProperty) this.m_eType = ECssPropertyType.cssptMarginBottom;
				else if ("margin-left" === sProperty) this.m_eType = ECssPropertyType.cssptMarginLeft;
				else if ("margin-right" === sProperty) this.m_eType = ECssPropertyType.cssptMarginRight;
				else if ("margin-top" === sProperty) this.m_eType = ECssPropertyType.cssptMarginTop;
				if ("mso-direction-alt" === sProperty) this.m_eType = ECssPropertyType.cssptMsoDirectionAlt;
				else if ("mso-fit-shape-to-text" === sProperty) this.m_eType = ECssPropertyType.cssptMsoFitShapeToText;
				else if ("mso-fit-text-to-shape" === sProperty) this.m_eType = ECssPropertyType.cssptMsoFitTextToShape;
				else if ("mso-layout-flow-alt" === sProperty) this.m_eType = ECssPropertyType.cssptMsoLayoutFlowAlt;
				else if ("mso-next-textbox" === sProperty) this.m_eType = ECssPropertyType.cssptMsoNextTextbox;
				else if ("mso-position-horizontal" === sProperty) this.m_eType = ECssPropertyType.cssptMsoPositionHorizontal;
				else if ("mso-position-horizontal-relative" === sProperty) this.m_eType = ECssPropertyType.cssptMsoPositionHorizontalRelative;
				else if ("mso-position-vertical" === sProperty) this.m_eType = ECssPropertyType.cssptMsoPositionVertical;
				else if ("mso-position-vertical-relative" === sProperty) this.m_eType = ECssPropertyType.cssptMsoPositionVerticalRelative;
				else if ("mso-rotate" === sProperty) this.m_eType = ECssPropertyType.cssptMsoRotate;
				else if ("mso-text-scale" === sProperty) this.m_eType = ECssPropertyType.cssptMsoTextScale;
				else if ("mso-text-shadow" === sProperty) this.m_eType = ECssPropertyType.cssptMsoTextShadow;
				else if ("mso-wrap-distance-bottom" === sProperty) this.m_eType = ECssPropertyType.cssptMsoWrapDistanceBottom;
				else if ("mso-wrap-distance-left" === sProperty) this.m_eType = ECssPropertyType.cssptMsoWrapDistanceLeft;
				else if ("mso-wrap-distance-right" === sProperty) this.m_eType = ECssPropertyType.cssptMsoWrapDistanceRight;
				else if ("mso-wrap-distance-top" === sProperty) this.m_eType = ECssPropertyType.cssptMsoWrapDistanceTop;
				else if ("mso-wrap-edited" === sProperty) this.m_eType = ECssPropertyType.cssptMsoWrapEdited;
				else if ("mso-wrap-style" === sProperty) this.m_eType = ECssPropertyType.cssptMsoWrapStyle;
				else if ("mso-height-percent" === sProperty) this.m_eType = ECssPropertyType.csspctMsoHeightPercent;
				else if ("mso-width-percent" === sProperty) this.m_eType = ECssPropertyType.csspctMsoWidthPercent;
				if ("position" === sProperty) this.m_eType = ECssPropertyType.cssptPosition;
				if ("rotation" === sProperty) this.m_eType = ECssPropertyType.cssptRotation;
				if ("text-decoration" === sProperty) this.m_eType = ECssPropertyType.cssptTextDecoration;
				else if ("top" === sProperty) this.m_eType = ECssPropertyType.cssptTop;
				else if ("text-align" === sProperty) this.m_eType = ECssPropertyType.cssptHTextAlign;
				if ("visibility" === sProperty) this.m_eType = ECssPropertyType.cssptVisibility;
				else if ("v-rotate-letters" === sProperty) this.m_eType = ECssPropertyType.cssptVRotateLetters;
				else if ("v-same-letter-heights" === sProperty) this.m_eType = ECssPropertyType.cssptVSameLetterHeights;
				else if ("v-text-align" === sProperty) this.m_eType = ECssPropertyType.cssptVTextAlign;
				else if ("v-text-anchor" === sProperty) this.m_eType = ECssPropertyType.cssptVTextAnchor;
				else if ("v-text-kern" === sProperty) this.m_eType = ECssPropertyType.cssptVTextKern;
				else if ("v-text-reverse" === sProperty) this.m_eType = ECssPropertyType.cssptVTextReverse;
				else if ("v-text-spacing-mode" === sProperty) this.m_eType = ECssPropertyType.cssptVTextSpacingMode;
				else if ("v-text-spacing" === sProperty) this.m_eType = ECssPropertyType.cssptVTextSpacing;
				if ("width" === sProperty) this.m_eType = ECssPropertyType.cssptWidth;
				if ("z-index" === sProperty) this.m_eType = ECssPropertyType.cssptZIndex;

				switch (this.m_eType) {
					case ECssPropertyType.cssptUnknown :
						this.ReadValue_Unknown(sValue);
						break;

					case ECssPropertyType.cssptFlip :
						this.ReadValue_Flip(sValue);
						break;
					case ECssPropertyType.cssptHeight :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptLeft :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptMarginBottom :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptMarginLeft :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptMarginRight :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptMarginTop :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptMsoPositionHorizontal :
						this.ReadValue_MsoPosHor(sValue);
						break;
					case ECssPropertyType.cssptMsoPositionHorizontalRelative :
						this.ReadValue_MsoPosHorRel(sValue);
						break;
					case ECssPropertyType.cssptMsoPositionVertical :
						this.ReadValue_MsoPosVer(sValue);
						break;
					case ECssPropertyType.cssptMsoPositionVerticalRelative :
						this.ReadValue_MsoPosVerRel(sValue);
						break;
					case ECssPropertyType.cssptMsoWrapDistanceBottom :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptMsoWrapDistanceLeft :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptMsoWrapDistanceRight :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptMsoWrapDistanceTop :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptMsoWrapEdited :
						this.ReadValue_Boolean(sValue);
						break;
					case ECssPropertyType.cssptMsoWrapStyle :
						this.ReadValue_MsoWrapStyle(sValue);
						break;
					case ECssPropertyType.cssptPosition :
						this.ReadValue_Position(sValue);
						break;
					case ECssPropertyType.cssptRotation :
						this.ReadValue_Rotation(sValue);
						break;
					case ECssPropertyType.cssptTop :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptVisibility :
						this.ReadValue_Visibility(sValue);
						break;
					case ECssPropertyType.cssptWidth :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptZIndex :
						this.ReadValue_ZIndex(sValue);
						break;

					case ECssPropertyType.cssptDirection :
						this.ReadValue_Direction(sValue);
						break;
					case ECssPropertyType.cssptLayoutFlow :
						this.ReadValue_LayoutFlow(sValue);
						break;
					case ECssPropertyType.cssptMsoDirectionAlt :
						this.ReadValue_DirectionAlt(sValue);
						break;
					case ECssPropertyType.cssptMsoFitShapeToText :
						this.ReadValue_Boolean(sValue);
						break;
					case ECssPropertyType.cssptMsoFitTextToShape :
						this.ReadValue_Boolean(sValue);
						break;
					case ECssPropertyType.cssptMsoLayoutFlowAlt :
						this.ReadValue_LayoutFlowAlt(sValue);
						break;
					case ECssPropertyType.cssptMsoNextTextbox :
						this.ReadValue_String(sValue);
						break;
					case ECssPropertyType.cssptMsoRotate :
						this.ReadValue_MsoRotate(sValue);
						break;
					case ECssPropertyType.cssptMsoTextScale :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptVTextAnchor :
						this.ReadValue_VTextAnchor(sValue);
						break;

					case ECssPropertyType.cssptFont :
						this.ReadValue_String(sValue);
						break;
					case ECssPropertyType.cssptFontFamily :
						this.ReadValue_String(sValue);
						break;
					case ECssPropertyType.cssptFontSize :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptFontStyle :
						this.ReadValue_FontStyle(sValue);
						break;
					case ECssPropertyType.cssptFontVariant :
						this.ReadValue_FontVariant(sValue);
						break;
					case ECssPropertyType.cssptFontWeight :
						this.ReadValue_FontWeight(sValue);
						break;
					case ECssPropertyType.cssptMsoTextShadow :
						this.ReadValue_Boolean(sValue);
						break;
					case ECssPropertyType.cssptTextDecoration :
						this.ReadValue_TextDecoration(sValue);
						break;
					case ECssPropertyType.cssptVRotateLetters :
						this.ReadValue_Boolean(sValue);
						break;
					case ECssPropertyType.cssptVSameLetterHeights :
						this.ReadValue_Boolean(sValue);
						break;
					case ECssPropertyType.cssptVTextAlign :
						this.ReadValue_VTextAlign(sValue);
						break;
					case ECssPropertyType.cssptVTextKern :
						this.ReadValue_Boolean(sValue);
						break;
					case ECssPropertyType.cssptVTextReverse :
						this.ReadValue_Boolean(sValue);
						break;
					case ECssPropertyType.cssptVTextSpacingMode :
						this.ReadValue_VTextSpacingMode(sValue);
						break;
					case ECssPropertyType.cssptVTextSpacing :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.csspctMsoWidthPercent :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.csspctMsoHeightPercent :
						this.ReadValue_Units(sValue);
						break;
					case ECssPropertyType.cssptHTextAlign :
						this.ReadValue_VTextAlign(sValue);
						break;
				}
			}
		};

		CCssProperty.prototype.ReadValue_Unknown = function (sValue) {
		};
		CCssProperty.prototype.ReadValue_Flip = function (sValue) {
			if ("x" === sValue) this.m_oValue.eFlip = ECssFlip.cssflipX;
			else if ("y" === sValue) this.m_oValue.eFlip = ECssFlip.cssflipY;
			else if ("xy" === sValue) this.m_oValue.eFlip = ECssFlip.cssflipXY;
			else if ("yx" === sValue) this.m_oValue.eFlip = ECssFlip.cssflipYX;
			else
				this.m_eType = ECssPropertyType.cssptUnknown;
		};
		CCssProperty.prototype.ReadValue_Units = function (sValue) {
			let nPos;
			if (-1 !== (nPos = sValue.indexOf("auto"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeAuto;
			} else if (-1 !== (nPos = sValue.indexOf("in"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeUnits;

				let strValue = sValue.substr(0, nPos);
				let dValue = parseFloat(sValue);

				this.m_oValue.oValue.dValue = Inch_To_Pt(dValue);
			} else if (-1 !== (nPos = sValue.indexOf("cm"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeUnits;

				let strValue = sValue.substr(0, nPos);
				let dValue = parseFloat(sValue);

				this.m_oValue.oValue.dValue = Cm_To_Pt(dValue);
			} else if (-1 !== (nPos = sValue.indexOf("mm"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeUnits;

				let strValue = sValue.substr(0, nPos);
				let dValue = parseFloat(sValue);

				this.m_oValue.oValue.dValue = Mm_To_Pt(dValue);
			} else if (-1 !== (nPos = sValue.indexOf("em"))) {
			} else if (-1 !== (nPos = sValue.indexOf("ex"))) {
			} else if (-1 !== (nPos = sValue.indexOf("pt"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeUnits;

				let strValue = sValue.substr(0, nPos);
				let dValue = parseFloat(sValue);

				this.m_oValue.oValue.dValue = dValue;
			} else if (-1 !== (nPos = sValue.indexOf("pc"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeUnits;

				let strValue = sValue.substr(0, nPos);
				let dValue = parseFloat(sValue);

				this.m_oValue.oValue.dValue = dValue * 12;
			} else if (-1 !== (nPos = sValue.indexOf("%"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypePerc;

				let strValue = sValue.substr(0, nPos);
				this.m_oValue.oValue.dValue = strValue.length === 0 ? 0 : parseFloat(strValue);
			} else if (-1 !== (nPos = sValue.indexOf("px"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeUnits;

				let strValue = sValue.substr(0, nPos);
				let dValue = parseFloat(sValue);

				this.m_oValue.oValue.dValue = Px_To_Pt(dValue);
			} else {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeAbsolute;
				this.m_oValue.oValue.dValue = parseFloat(sValue);
			}
		};

		CCssProperty.prototype.ReadValue_MsoPosHor = function (sValue) {
			if ("absolute" === sValue) this.m_oValue.eMsoPosHor = ECssMsoPosHor.cssmsoposhorAbsolute;
			else if ("left" === sValue) this.m_oValue.eMsoPosHor = ECssMsoPosHor.cssmsoposhorLeft;
			else if ("center" === sValue) this.m_oValue.eMsoPosHor = ECssMsoPosHor.cssmsoposhorCenter;
			else if ("right" === sValue) this.m_oValue.eMsoPosHor = ECssMsoPosHor.cssmsoposhorRight;
			else if ("inside" === sValue) this.m_oValue.eMsoPosHor = ECssMsoPosHor.cssmsoposhorInside;
			else if ("outside" === sValue) this.m_oValue.eMsoPosHor = ECssMsoPosHor.cssmsoposhorOutside;
			else
				this.m_oValue.eMsoPosHor = ECssMsoPosHor.cssmsoposhorAbsolute;
		};
		CCssProperty.prototype.ReadValue_MsoPosHorRel = function (sValue) {
			if ("left-margin-area" === sValue) this.m_oValue.eMsoPosHorRel = ECssMsoPosHorRel.cssmsoposhorrelLeftMargin;
			else if ("right-margin-area" === sValue) this.m_oValue.eMsoPosHorRel = ECssMsoPosHorRel.cssmsoposhorrelRightMargin;
			else if ("margin" === sValue) this.m_oValue.eMsoPosHorRel = ECssMsoPosHorRel.cssmsoposhorrelMargin;
			else if ("page" === sValue) this.m_oValue.eMsoPosHorRel = ECssMsoPosHorRel.cssmsoposhorrelPage;
			else if ("text" === sValue) this.m_oValue.eMsoPosHorRel = ECssMsoPosHorRel.cssmsoposhorrelText;
			else if ("char" === sValue) this.m_oValue.eMsoPosHorRel = ECssMsoPosHorRel.cssmsoposhorrelChar;
			else
				this.m_oValue.eMsoPosHorRel = ECssMsoPosHorRel.cssmsoposhorrelText;
		};
		CCssProperty.prototype.ReadValue_MsoPosVer = function (sValue) {
			if ("absolute" === sValue) this.m_oValue.eMsoPosVer = ECssMsoPosVer.cssmsoposverAbsolute;
			else if ("top" === sValue) this.m_oValue.eMsoPosVer = ECssMsoPosVer.cssmsoposverTop;
			else if ("center" === sValue) this.m_oValue.eMsoPosVer = ECssMsoPosVer.cssmsoposverCenter;
			else if ("bottom" === sValue) this.m_oValue.eMsoPosVer = ECssMsoPosVer.cssmsoposverBottom;
			else if ("inside" === sValue) this.m_oValue.eMsoPosVer = ECssMsoPosVer.cssmsoposverInside;
			else if ("outside" === sValue) this.m_oValue.eMsoPosVer = ECssMsoPosVer.cssmsoposverOutside;
			else
				this.m_oValue.eMsoPosVer = ECssMsoPosVer.cssmsoposverAbsolute;
		};
		CCssProperty.prototype.ReadValue_MsoPosVerRel = function (sValue) {
			if ("bottom-margin-area" === sValue) this.m_oValue.eMsoPosVerRel = ECssMsoPosVerRel.cssmsoposverrelBottomMargin;
			else if ("top-margin-area" === sValue) this.m_oValue.eMsoPosVerRel = ECssMsoPosVerRel.cssmsoposverrelTopMargin;
			else if ("margin" === sValue) this.m_oValue.eMsoPosVerRel = ECssMsoPosVerRel.cssmsoposverrelMargin;
			else if ("page" === sValue) this.m_oValue.eMsoPosVerRel = ECssMsoPosVerRel.cssmsoposverrelPage;
			else if ("text" === sValue) this.m_oValue.eMsoPosVerRel = ECssMsoPosVerRel.cssmsoposverrelText;
			else if ("line" === sValue) this.m_oValue.eMsoPosVerRel = ECssMsoPosVerRel.cssmsoposverrelLine;
			else
				this.m_oValue.eMsoPosVerRel = ECssMsoPosVerRel.cssmsoposverrelText;
		};

		CCssProperty.prototype.ReadValue_Rotation = function (sValue) {
			this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeAbsolute;
			this.m_oValue.oValue.dValue = sValue.length === 0 ? 0 : parseFloat(sValue);

			if (sValue.indexOf("fd") !== -1) {
				this.m_oValue.oValue.dValue /= 6000.;
			} else if (sValue.indexOf("f") === sValue.length - 1) {
				this.m_oValue.oValue.dValue /= 65536.;
			}
		};
		CCssProperty.prototype.ReadValue_Boolean = function (sValue) {
			if ("true" === sValue || "t" === sValue || "1" === sValue)
				this.m_oValue.bValue = true;
			else
				this.m_oValue.bValue = false;
		};
		CCssProperty.prototype.ReadValue_MsoWrapStyle = function (sValue) {
			if ("square" === sValue) this.m_oValue.eMsoWrapStyle = ECssMsoWrapStyle.cssmsowrapstyleSqaure;
			else if ("none" === sValue) this.m_oValue.eMsoWrapStyle = ECssMsoWrapStyle.cssmsowrapstyleNone;
			else
				this.m_oValue.eMsoWrapStyle = ECssMsoWrapStyle.cssmsowrapstyleSqaure;
		};
		CCssProperty.prototype.ReadValue_Position = function (sValue) {
			if ("static" === sValue) this.m_oValue.ePosition = ECssPosition.csspositionStatic;
			else if ("absolute" === sValue) this.m_oValue.ePosition = ECssPosition.csspositionAbsolute;
			else if ("relative" === sValue) this.m_oValue.ePosition = ECssPosition.csspositionRelative;
			else
				this.m_oValue.ePosition = ECssPosition.csspositionAbsolute;
		};
		CCssProperty.prototype.ReadValue_Visibility = function (sValue) {
			if ("hidden" === sValue) this.m_oValue.eVisibility = ECssVisibility.cssvisibilityHidden;
			else if ("inherit" === sValue) this.m_oValue.eVisibility = ECssVisibility.cssvisibilityInherit;
			else
				this.m_oValue.eVisibility = ECssVisibility.cssvisibilityInherit;
		};
		CCssProperty.prototype.ReadValue_ZIndex = function (sValue) {
			if ("auto" === sValue) this.m_oValue.oZIndex.m_eType = ECssZIndexType.csszindextypeAuto;
			else {
				this.m_oValue.oZIndex.m_eType = ECssZIndexType.csszindextypeOrder;
				this.m_oValue.oZIndex.nOrder = parseInt(sValue);

			}
		};
		CCssProperty.prototype.ReadValue_Direction = function (sValue) {
			if ("ltr" === sValue) this.m_oValue.eDirection = ECssDirection.cssdirectionLTR;
			else if ("rtl" === sValue) this.m_oValue.eDirection = ECssDirection.cssdirectionRTL;
			else
				this.m_oValue.eDirection = ECssDirection.cssdirectionLTR;
		};
		CCssProperty.prototype.ReadValue_LayoutFlow = function (sValue) {
			if ("horizontal" === sValue) this.m_oValue.eLayoutFlow = ECssLayoutFlow.csslayoutflowHorizontal;
			else if ("vertical" === sValue) this.m_oValue.eLayoutFlow = ECssLayoutFlow.csslayoutflowVertical;
			else if ("vertical-ideographic" === sValue) this.m_oValue.eLayoutFlow = ECssLayoutFlow.csslayoutflowVerticalIdeographic;
			else if ("horizontal-ideographic" === sValue) this.m_oValue.eLayoutFlow = ECssLayoutFlow.csslayoutflowHorizontalIdeographic;
			else
				this.m_oValue.eLayoutFlow = ECssLayoutFlow.csslayoutflowHorizontal;
		};
		CCssProperty.prototype.ReadValue_DirectionAlt = function (sValue) {
			this.m_oValue.eDirectionAlt = ECssDirectionAlt.cssdirectionaltContext;
		};
		CCssProperty.prototype.ReadValue_LayoutFlowAlt = function (sValue) {
			this.m_oValue.eLayoutFlowAlt = ECssLayoutFlowAlt.csslayoutflowaltBottomToTop;
		};
		CCssProperty.prototype.ReadValue_String = function (sValue) {
			this.m_oValue.wsValue = sValue;
		};

		CCssProperty.prototype.ReadValue_MsoRotate = function (sValue) {
			if ("0" === sValue) this.m_oValue.eRotate = ECssMsoRotate.cssmsorotate0;
			else if ("90" === sValue) this.m_oValue.eRotate = ECssMsoRotate.cssmsorotate90;
			else if ("180" === sValue) this.m_oValue.eRotate = ECssMsoRotate.cssmsorotate180;
			else if ("-90" === sValue) this.m_oValue.eRotate = ECssMsoRotate.cssmsorotate270;
			else
				this.m_oValue.eRotate = ECssMsoRotate.cssmsorotate0;
		};
		CCssProperty.prototype.ReadValue_VTextAnchor = function (sValue) {
			if ("top" === sValue) this.m_oValue.eVTextAnchor = ECssVTextAnchor.cssvtextanchorTop;
			else if ("middle" === sValue) this.m_oValue.eVTextAnchor = ECssVTextAnchor.cssvtextanchorMiddle;
			else if ("bottom" === sValue) this.m_oValue.eVTextAnchor = ECssVTextAnchor.cssvtextanchorBottom;
			else if ("top-center" === sValue) this.m_oValue.eVTextAnchor = ECssVTextAnchor.cssvtextanchorTopCenter;
			else if ("middle-center" === sValue) this.m_oValue.eVTextAnchor = ECssVTextAnchor.cssvtextanchorMiddleCenter;
			else if ("bottom-center" === sValue) this.m_oValue.eVTextAnchor = ECssVTextAnchor.cssvtextanchorBottomCenter;
			else if ("top-baseline" === sValue) this.m_oValue.eVTextAnchor = ECssVTextAnchor.cssvtextanchorTopBaseline;
			else if ("bottom-baseline" === sValue) this.m_oValue.eVTextAnchor = ECssVTextAnchor.cssvtextanchorBottomBaseline;
			else if ("top-center-baseline" === sValue) this.m_oValue.eVTextAnchor = ECssVTextAnchor.cssvtextanchorTopCenterBaseline;
			else if ("bottom-center-baseline" === sValue) this.m_oValue.eVTextAnchor = ECssVTextAnchor.cssvtextanchorBottomCenterBaseline;
			else
				this.m_oValue.eVTextAnchor = ECssVTextAnchor.cssvtextanchorTop;
		};
		CCssProperty.prototype.ReadValue_FontStyle = function (sValue) {
			if ("normal" === sValue) this.m_oValue.eFontStyle = ECssFontStyle.cssfontstyleNormal;
			else if ("italic" === sValue) this.m_oValue.eFontStyle = ECssFontStyle.cssfontstyleItalic;
			else if ("oblique" === sValue) this.m_oValue.eFontStyle = ECssFontStyle.cssfontstyleOblique;
			else
				this.m_oValue.eFontStyle = ECssFontStyle.cssfontstyleNormal;
		};
		CCssProperty.prototype.ReadValue_FontVariant = function (sValue) {
			if ("normal" === sValue) this.m_oValue.eFontVariant = ECssFontVarian.cssfontvariantNormal;
			else if ("small-caps" === sValue) this.m_oValue.eFontVariant = ECssFontVarian.cssfontvariantSmallCaps;
			else
				this.m_oValue.eFontVariant = ECssFontVarian.cssfontvariantNormal;
		};
		CCssProperty.prototype.ReadValue_FontWeight = function (sValue) {
			if ("normal" === sValue) this.m_oValue.eFontWeight = ECssFontWeight.cssfontweightNormal;
			else if ("lighter" === sValue) this.m_oValue.eFontWeight = ECssFontWeight.cssfontweightLighter;
			else if ("100" === sValue) this.m_oValue.eFontWeight = ECssFontWeight.cssfontweight100;
			else if ("200" === sValue) this.m_oValue.eFontWeight = ECssFontWeight.cssfontweight200;
			else if ("300" === sValue) this.m_oValue.eFontWeight = ECssFontWeight.cssfontweight300;
			else if ("400" === sValue) this.m_oValue.eFontWeight = ECssFontWeight.cssfontweight400;
			else if ("bold" === sValue) this.m_oValue.eFontWeight = ECssFontWeight.cssfontweightBold;
			else if ("bolder" === sValue) this.m_oValue.eFontWeight = ECssFontWeight.cssfontweightBolder;
			else if ("500" === sValue) this.m_oValue.eFontWeight = ECssFontWeight.cssfontweight500;
			else if ("600" === sValue) this.m_oValue.eFontWeight = ECssFontWeight.cssfontweight600;
			else if ("700" === sValue) this.m_oValue.eFontWeight = ECssFontWeight.cssfontweight700;
			else if ("800" === sValue) this.m_oValue.eFontWeight = ECssFontWeight.cssfontweight800;
			else if ("900" === sValue) this.m_oValue.eFontWeight = ECssFontWeight.cssfontweight900;
			else
				this.m_oValue.eFontWeight = ECssFontWeight.cssfontweightNormal;
		};
		CCssProperty.prototype.ReadValue_TextDecoration = function (sValue) {
			if ("none" === sValue) this.m_oValue.eTextDecoration = ECssTextDecoration.csstextdecorationNone;
			else if ("underline" === sValue) this.m_oValue.eTextDecoration = ECssTextDecoration.csstextdecorationUnderline;
			else if ("overline" === sValue) this.m_oValue.eTextDecoration = ECssTextDecoration.csstextdecorationOverline;
			else if ("line-through" === sValue) this.m_oValue.eTextDecoration = ECssTextDecoration.csstextdecorationLineThrough;
			else if ("blink" === sValue) this.m_oValue.eTextDecoration = ECssTextDecoration.csstextdecorationBlink;
			else
				this.m_oValue.eTextDecoration = ECssTextDecoration.csstextdecorationNone;
		};
		CCssProperty.prototype.ReadValue_VTextAlign = function (sValue) {
			if ("left" === sValue) this.m_oValue.eVTextAlign = ECssVTextAlign.cssvtextalignLeft;
			else if ("right" === sValue) this.m_oValue.eVTextAlign = ECssVTextAlign.cssvtextalignRight;
			else if ("center" === sValue) this.m_oValue.eVTextAlign = ECssVTextAlign.cssvtextalignCenter;
			else if ("justify" === sValue) this.m_oValue.eVTextAlign = ECssVTextAlign.cssvtextalignJustify;
			else if ("letter-justify" === sValue) this.m_oValue.eVTextAlign = ECssVTextAlign.cssvtextalignLetterJustify;
			else if ("stretch-justify" === sValue) this.m_oValue.eVTextAlign = ECssVTextAlign.cssvtextalignStretchJustify;
			else
				this.m_oValue.eVTextAlign = ECssVTextAlign.cssvtextalignLeft;
		};
		CCssProperty.prototype.ReadValue_VTextSpacingMode = function (sValue) {
			if ("tightening" === sValue) this.m_oValue.eVTextSpacingMode = ECssVTextSpacingMode.cssvtextspacingmodeTightening;
			else if ("tracking" === sValue) this.m_oValue.eVTextSpacingMode = ECssVTextSpacingMode.cssvtextspacingmodeTracking;
			else
				this.m_oValue.eVTextSpacingMode = ECssVTextSpacingMode.cssvtextspacingmodeTightening;
		};


		let EBorderType =
			{
				bordertypeDash: 0, // (pecifies a line border consisting of a dashed line around the parent object.)
				bordertypeDashDotDot: 1, // (Dash Dot Dot Border)
				bordertypeDashDotStroked: 2, // (Stroked Dash Dot Border)
				bordertypeDashedSmall: 3, // (Small Dash Border)
				bordertypeDot: 4, // (Dotted Border)
				bordertypeDotDash: 5, // (Dot Dash Border)
				bordertypeDouble: 6, // (Double Line Border)
				bordertypeDoubleWave: 7, // (Double Wavy Lines Border)
				bordertypeHairline: 8, // (Hairline Border)
				bordertypeHTMLInset: 9, // (Inset Border)
				bordertypeHTMLOutset: 10, // (Outset Border)
				bordertypeNone: 11, // (No Border)
				bordertypeSingle: 12, // (Single Line Border)
				bordertypeThick: 13, // (Thick Line Border)
				bordertypeThickBetweenThin: 14, // (Thin-thick-thin Border)
				bordertypeThickBetweenThinLarge: 15, // (Large thin-thick-thin Border)
				bordertypeThickBetweenThinSmall: 16, // (Small thin-thick-thin Lines Border)
				bordertypeThickThin: 17, // (Thick Thin Line Border)
				bordertypeThickThinLarge: 18, // (Thick Thin Large Gap Border)
				bordertypeThickThinSmall: 19, // (Small thick-thin lines border)
				bordertypeThinThick: 20, // (Thin Thick Line Border)
				bordertypeThinThickLarge: 21, // (Thin Thick Large Gap Border)
				bordertypeThinThickSmall: 22, // (Thin Thick Small Gap Border)
				bordertypeThreeDEmboss: 23, // (3D Embossed Border)
				bordertypeThreeDEngrave: 24, // (3D Engraved Border)
				bordertypeTriple: 25, // (Triple Line Border)
				bordertypeWave: 26 // (Wavy Border)
			};


		function readBorderType(reader) {
			let sValue = reader.GetValue();
			let wChar = sValue.charAt(0);
			switch (wChar) {
				case 'd':
					if ("dash" === sValue) return EBorderType.bordertypeDash;
					else if ("dashDotDot" === sValue) return EBorderType.bordertypeDashDotDot;
					else if ("dashDotStroked" === sValue) return EBorderType.bordertypeDashDotStroked;
					else if ("dashedSmall" === sValue) return EBorderType.bordertypeDashedSmall;
					else if ("dot" === sValue) return EBorderType.bordertypeDot;
					else if ("dotDash" === sValue) return EBorderType.bordertypeDotDash;
					else if ("double" === sValue) return EBorderType.bordertypeDouble;
					else if ("doubleWave" === sValue) return EBorderType.bordertypeDoubleWave;
					break;
				case 'h':
					if ("hairline" === sValue) return EBorderType.bordertypeHairline;
					break;

				case 'H':
					if ("HTMLInset" === sValue) return EBorderType.bordertypeHTMLInset;
					else if ("HTMLOutset" === sValue) return EBorderType.bordertypeHTMLOutset;
					break;

				case 'n':
					if ("none" === sValue) return EBorderType.bordertypeNone;
					break;

				case 's':
					if ("single" === sValue) return EBorderType.bordertypeSingle;
					break;

				case 't':
					if ("thick" === sValue) return EBorderType.bordertypeThick;
					else if ("thickBetweenThin" === sValue) return EBorderType.bordertypeThickBetweenThin;
					else if ("thickBetweenThinLarge" === sValue) return EBorderType.bordertypeThickBetweenThinLarge;
					else if ("thickBetweenThinSmall" === sValue) return EBorderType.bordertypeThickBetweenThinSmall;
					else if ("thickThin" === sValue) return EBorderType.bordertypeThickThin;
					else if ("thickThinLarge" === sValue) return EBorderType.bordertypeThickThinLarge;
					else if ("thickThinSmall" === sValue) return EBorderType.bordertypeThickThinSmall;
					else if ("thinThick" === sValue) return EBorderType.bordertypeThinThick;
					else if ("thinThickLarge" === sValue) return EBorderType.bordertypeThinThickLarge;
					else if ("thinThickSmall" === sValue) return EBorderType.bordertypeThinThickSmall;
					else if ("threeDEmboss" === sValue) return EBorderType.bordertypeThreeDEmboss;
					else if ("threeDEngrave" === sValue) return EBorderType.bordertypeThreeDEngrave;
					else if ("triple" === sValue) return EBorderType.bordertypeTriple;
					break;

				case 'w':
					if ("wave" === sValue) return EBorderType.bordertypeWave;
					break;
			}
			return EBorderType.bordertypeNone;
		}

		function getBorderType(nType) {
			if (EBorderType.bordertypeDash === nType) return "dash";
			else if (EBorderType.bordertypeDashDotDot === nType) return "dashDotDot";
			else if (EBorderType.bordertypeDashDotStroked === nType) return "dashDotStroked";
			else if (EBorderType.bordertypeDashedSmall === nType) return "dashedSmall";
			else if (EBorderType.bordertypeDot === nType) return "dot";
			else if (EBorderType.bordertypeDotDash === nType) return "dotDash";
			else if (EBorderType.bordertypeDouble === nType) return "double";
			else if (EBorderType.bordertypeDoubleWave === nType) return "doubleWave";
			if (EBorderType.bordertypeHairline === nType) return "hairline";
			if (EBorderType.bordertypeHTMLInset === nType) return "HTMLInset";
			else if (EBorderType.bordertypeHTMLOutset === nType) return "HTMLOutset";
			if (EBorderType.bordertypeNone === nType) return ("none");
			if (EBorderType.bordertypeSingle === nType) return ("single");
			if (EBorderType.bordertypeThick === nType) return "thick";
			else if (EBorderType.bordertypeThickBetweenThin === nType) return ("thickBetweenThin");
			else if (EBorderType.bordertypeThickBetweenThinLarge === nType) return ("thickBetweenThinLarge");
			else if (EBorderType.bordertypeThickBetweenThinSmall === nType) return "thickBetweenThinSmall";
			else if (EBorderType.bordertypeThickThin === nType) return ("thickThin");
			else if (EBorderType.bordertypeThickThinLarge === nType) return ("thickThinLarge");
			else if (EBorderType.bordertypeThickThinSmall === nType) return ("thickThinSmall");
			else if (EBorderType.bordertypeThinThick === nType) return ("thinThick");
			else if (EBorderType.bordertypeThinThickLarge === nType) return ("thinThickLarge");
			else if (EBorderType.bordertypeThinThickSmall === nType) return ("thinThickSmall");
			else if (EBorderType.bordertypeThreeDEmboss === nType) return ("threeDEmboss");
			else if (EBorderType.bordertypeThreeDEngrave === nType) return ("threeDEngrave");
			else if (EBorderType.bordertypeTriple === nType) return ("triple");
			if (EBorderType.bordertypeWave === nType) return ("wave");
			return null;
		}

		let EBorderShadow =
			{
				bordershadowFalse: 0,
				bordershadowTrue: 1
			};

		function readBorderShadow(reader) {
			let bVal = reader.GetValueBool();
			return bVal ? EBorderShadow.bordershadowTrue : EBorderShadow.bordershadowFalse;
		}

		function getBorderShadow(nType) {
			if (nType === EBorderShadow.bordershadowTrue) {
				return "t";
			}
			if (nType === EBorderShadow.bordershadowFalse) {
				return "f";
			}
			return false;
		}


		let EWrapSide =
			{
				wrapsideBoth: 0,
				wrapsideLargest: 1,
				wrapsideLeft: 2,
				wrapsideRight: 3
			};

		function readWrapSide(reader) {
			let sValue = reader.GetValue();

			if ("both" === sValue) return EWrapSide.wrapsideBoth;
			if ("largest" === sValue) return EWrapSide.wrapsideLargest;
			if ("left" === sValue) return EWrapSide.wrapsideLeft;
			if ("right" === sValue) return EWrapSide.wrapsideRight;
		}

		function getWrapSide(nType) {
			if (EWrapSide.wrapsideBoth === nType) return "both";
			if (EWrapSide.wrapsideLargest === nType) return "largest";
			if (EWrapSide.wrapsideLeft === nType) return "left";
			if (EWrapSide.wrapsideRight === nType) return "right";
			return null;
		}

		let EVerticalAnchor =
			{
				verticalanchorLine: 0,
				verticalanchorMargin: 1,
				verticalanchorPage: 2,
				verticalanchorText: 3
			};

		function readVerticalAnchor(reader) {
			let sVal = reader.GetValue();
			if (sVal === "line") return EVerticalAnchor.verticalanchorLine;
			if (sVal === "margin") return EVerticalAnchor.verticalanchorMargin;
			if (sVal === "page") return EVerticalAnchor.verticalanchorPage;
			if (sVal === "text") return EVerticalAnchor.verticalanchorText;
			return EVerticalAnchor.verticalanchorLine;
		}

		function getVerticalAnchor(nType) {
			if (nType === EVerticalAnchor.verticalanchorLine) return "line";
			if (nType === EVerticalAnchor.verticalanchorMargin) return "margin";
			if (nType === EVerticalAnchor.verticalanchorPage) return "page";
			if (nType === EVerticalAnchor.verticalanchorText) return "text";
			return null;
		}

		function readHorizontalAnchor(reader) {
			let sVal = reader.GetValue();
			if (sVal === "char") return EHorizontalAnchor.horizontalanchorChar;
			if (sVal === "margin") return EHorizontalAnchor.horizontalanchorMargin;
			if (sVal === "page") return EHorizontalAnchor.horizontalanchorPage;
			if (sVal === "text") return EHorizontalAnchor.horizontalanchorText;
			return null;
		}

		function getHorizontalAnchor(nType) {
			if (nType === EHorizontalAnchor.horizontalanchorChar) return "char";
			if (nType === EHorizontalAnchor.horizontalanchorMargin) return "margin";
			if (nType === EHorizontalAnchor.horizontalanchorPage) return "page";
			if (nType === EHorizontalAnchor.horizontalanchorText) return "text";
			return null;
		}


		let EHorizontalAnchor =
			{
				horizontalanchorChar: 0,
				horizontalanchorMargin: 1,
				horizontalanchorPage: 2,
				horizontalanchorText: 3
			};


		let EWrapType =
			{
				wraptypeNone: 0,
				wraptypeSquare: 1,
				wraptypeThrough: 2,
				wraptypeTight: 3,
				wraptypeTopAndBottom: 4
			};

		function readWrapType(reader) {
			let sVal = reader.GetValue();
			if (sVal === "none") return EWrapType.wraptypeNone;
			if (sVal === "square") return EWrapType.wraptypeSquare;
			if (sVal === "through") return EWrapType.wraptypeThrough;
			if (sVal === "tight") return EWrapType.wraptypeTight;
			if (sVal === "topAndBottom") return EWrapType.wraptypeTopAndBottom;
			return EWrapType.wraptypeNone;
		}

		function getWrapType(nType) {
			if (nType === EWrapType.wraptypeNone) return "none";
			if (nType === EWrapType.wraptypeSquare) return "square";
			if (nType === EWrapType.wraptypeThrough) return "through";
			if (nType === EWrapType.wraptypeTight) return "tight";
			if (nType === EWrapType.wraptypeTopAndBottom) return "topAndBottom";
			return null;
		}


		let EVmlClientDataObjectType =
			{
				vmlclientdataobjecttypeButton: 0,
				vmlclientdataobjecttypeCheckbox: 1,
				vmlclientdataobjecttypeDialog: 2,
				vmlclientdataobjecttypeDrop: 3,
				vmlclientdataobjecttypeEdit: 4,
				vmlclientdataobjecttypeGBox: 5,
				vmlclientdataobjecttypeGroup: 6,
				vmlclientdataobjecttypeLabel: 7,
				vmlclientdataobjecttypeLineA: 8,
				vmlclientdataobjecttypeList: 9,
				vmlclientdataobjecttypeMovie: 10,
				vmlclientdataobjecttypeNote: 11,
				vmlclientdataobjecttypePict: 12,
				vmlclientdataobjecttypeRadio: 13,
				vmlclientdataobjecttypeRect: 14,
				vmlclientdataobjecttypeRectA: 15,
				vmlclientdataobjecttypeScroll: 16,
				vmlclientdataobjecttypeShape: 17,
				vmlclientdataobjecttypeSpin: 18
			};

		function readClientDataObjectType(reader) {
			let sValue = reader.GetValue();
			if ("Button" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeButton;
			else if ("Checkbox" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeCheckbox;
			else if ("Dialog" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeDialog;
			else if ("Drop" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeDrop;
			else if ("Edit" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeEdit;
			else if ("GBox" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeGBox;
			else if ("Group" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeGroup;
			else if ("Label" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeLabel;
			else if ("LineA" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeLineA;
			else if ("List" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeList;
			else if ("Movie" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeMovie;
			else if ("Note" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeNote;
			else if ("Pict" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypePict;
			else if ("Radio" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeRadio;
			else if ("Rect" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeRect;
			else if ("RectA" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeRectA;
			else if ("Scroll" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeScroll;
			else if ("Shape" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeShape;
			else if ("Spin" === sValue) return EVmlClientDataObjectType.vmlclientdataobjecttypeSpin;
			return EVmlClientDataObjectType.vmlclientdataobjecttypeButton;
		}

		function getClientDataObjectType(nType) {
			if (EVmlClientDataObjectType.vmlclientdataobjecttypeButton === nType) return ("Button");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeCheckbox === nType) return ("Checkbox");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeDialog === nType) return ("Dialog");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeDrop === nType) return ("Drop");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeEdit === nType) return ("Edit");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeGBox === nType) return ("GBox");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeGroup === nType) return ("Group");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeLabel === nType) return ("Label");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeLineA === nType) return ("LineA");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeList === nType) return ("List");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeMovie === nType) return ("Movie");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeNote === nType) return ("Note");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypePict === nType) return ("Pict");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeRadio === nType) return ("Radio");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeRect === nType) return ("Rect");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeRectA === nType) return ("RectA");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeScroll === nType) return ("Scroll");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeShape === nType) return ("Shape");
			else if (EVmlClientDataObjectType.vmlclientdataobjecttypeSpin === nType) return ("Spin");
			return null;
		}

		window['AscFormat'].CVMLDrawing = CVMLDrawing;
		window['AscFormat'].ECssPropertyType = ECssPropertyType;
		window['AscFormat'].ECssUnitsType = ECssUnitsType;
		window['AscFormat'].CPoint = CPoint;

	})(window);

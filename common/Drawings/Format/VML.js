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
						sColor = sColors.substring(nStartPos, nEndPos);
						let oColor = new CColor();
						oColor.fromString(sColor);
						this.m_arrColors.push(oColor);
						nStartPos = nEndPos + 1;
					}
					nEndPos = sColors.length;
					sColor = sColors.substring(nStartPos, nEndPos);
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
		CRelationTable.prototype.writeChildrenXml = function (writer) {
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
		CDiagram.prototype.writeChildrenXml = function (writer) {
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
		CEquationXml.prototype.writeChildrenXml = function (writer) {
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
			this.m_oPic= null;
		}

		IC(COLEObject, CBaseNoId, 0);
		COLEObject.prototype.readAttrXml = function (name, reader) {

			if (name === "DrawAspect") this.m_oDrawAspect = readOLEDrawAspect(reader);
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
			} else if("pic" === name) {
				let oPic = new AscFormat.CImageShape();
				oPic.fromXml(reader);
				this.m_oPic = oPic;
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
		COLEObject.prototype.writeChildrenXml = function (writer) {

			if (this.m_oLinkType)
				this.m_oLinkType.toXml(writer, "o:LinkType");

			if (this.m_oLockedField)
				this.m_oLockedField.toXml(writer, "o:LockedField");

			if (this.m_oFieldCodes)
				this.m_oFieldCodes.toXml(writer, "o:FieldCodes");
		};
		COLEObject.prototype.fillEditorOleObject = function (oEditorObject, oOOXMLDrawing, reader) {
			oEditorObject.m_sData = null;
			oEditorObject.setApplicationId(this.m_sProgId);
			oEditorObject.fillDataLink(this.m_oId, reader);
			if(oOOXMLDrawing) {
				if(oOOXMLDrawing.nvPicPr)
				{
					oEditorObject.setNvPicPr(oOOXMLDrawing.nvPicPr.createDuplicate());
				}
				if(oOOXMLDrawing.spPr)
				{
					oEditorObject.setSpPr(oOOXMLDrawing.spPr.createDuplicate());
					oEditorObject.spPr.setParent(oEditorObject);
				}
				if(oOOXMLDrawing.blipFill)
				{
					oEditorObject.setBlipFill(oOOXMLDrawing.blipFill);
				}
				if(oOOXMLDrawing.style)
				{
					oEditorObject.setStyle(oOOXMLDrawing.style.createDuplicate());
				}
			}
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
		CR.prototype.writeChildrenXml = function (writer) {
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
		CRegroupTable.prototype.writeChildrenXml = function (writer) {
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
		CRules.prototype.writeChildrenXml = function (writer) {

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
		CShapeLayout.prototype.writeChildrenXml = function (writer) {
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
		
		let ODRAW = {
			lineJoinMiter: 0,
			lineJoinRound: 1
		};

		function CVmlHandle() {
			this.position = null;
			this.xrange = null;
			this.yrange = null;
			this.switchHandle = null;
			this.polar = null;
			this.radiusrange = null;
		}

		const ShapeSize = 43200.0;
		const ShapeSizeVML = 21600;
		const RadKoef = Math.PI/10800000.0;
		const pow3_16 = 60000;//?!! from core

		const PPTFormulaType = {
			ftSum: 0,
			ftProduct: 1,
			ftMid: 2,
			ftAbsolute: 3,
			ftMin: 4,
			ftMax: 5,
			ftIf: 6,
			ftSqrt: 7,
			ftMod: 8,
			ftSin: 9,
			ftCos: 10,
			ftTan: 11,
			ftAtan2: 12,
			ftSinatan2: 13,
			ftCosatan2: 14,
			ftSumangle: 15,
			ftEllipse: 16,
			ftVal: 17
		};

		const ParamType =
		{
			ptFormula	: 0,
			ptAdjust	: 1,
			ptValue		: 2
		};

		function GetValue(strParam, oRes, lShapeWidth, lShapeHeight) {
			oRes.ptType = ParamType.ptValue;
			oRes.bRes		= true;
			let val	= 0;
			let lShapeWidth_ = AscFormat.isRealNumber(lShapeWidth) ? lShapeWidth : ShapeSizeVML;
			let lShapeHeight_ = AscFormat.isRealNumber(lShapeHeight) ? lShapeHeight : ShapeSizeVML;
			if ('#' === strParam[0]) {
				oRes.ptType = ParamType.ptAdjust;
				val = parseInt(strParam.substring(1));
			}
			else if ('@' === strParam[0]) {
				oRes.ptType = ParamType.ptFormula;
				val = parseInt(strParam.substring(1));
			}
			else {
				let nNumVal = parseInt(strParam);
				if (!AscFormat.isRealNumber(nNumVal)) {
					if ("width" === strParam) {
						val = lShapeWidth_;
					}
					else if ("height" === strParam) {
						val = lShapeHeight_;
					}
					else if ("pixelWidth" === strParam) {
						val = lShapeWidth_;
					}
					else if ("pixelHeight" === strParam) {
						val = lShapeHeight_;
					}
					else if ("pixelLineWidth" === strParam || "lineDrawn" === strParam) {
						val = 1;
					}
					else {
						oRes.bRes = false;
						return 0;
					}
				}
				else {
					oRes.ptType = ParamType.ptValue;
					val = parseInt(strParam);
				}
			}
			return val;
		}


		function GetRuler(strName, oRes)
		{
			oRes.bRes = true;
			if		("moveTo"		=== strName)	return RulesType.rtOOXMLMoveTo;
			else if ("lnTo"			=== strName)	return RulesType.rtOOXMLLineTo;
			else if ("cubicBezTo"	=== strName)	return RulesType.rtOOXMLCubicBezTo;
			else if ("close"		=== strName)	return RulesType.rtOOXMLClose;
			else if ("end"			=== strName)	return RulesType.rtOOXMLEnd;
			else if ("arcTo"		=== strName)	return RulesType.rtOOXMLArcTo;
			else if ("quadBezTo"	=== strName)	return RulesType.rtOOXMLQuadBezTo;

			//bRes = true;
			else if	(("m" === strName) || ("M" === strName))		return RulesType.rtMoveTo;
			else if (("l" === strName) || ("L" === strName))		return RulesType.rtLineTo;
			else if (("c" === strName) || ("C" === strName))		return RulesType.rtCurveTo;
			else if (("x" === strName) || ("Z" === strName))		return RulesType.rtClose;
			else if (("e" === strName) || ("N" === strName))		return RulesType.rtEnd;
			else if ("t" === strName)								return RulesType.rtRMoveTo;
			else if ("r" === strName)								return RulesType.rtRLineTo;
			else if ("v" === strName)								return RulesType.rtRCurveTo;
			else if (("nf" === strName) || ("F" === strName))		return RulesType.rtNoFill;
			else if (("ns" === strName) || ("S" === strName))		return RulesType.rtNoStroke;
			else if (("ae" === strName) || ("T" === strName))		return RulesType.rtAngleEllipseTo;
			else if (("al" === strName) || ("U" === strName))		return RulesType.rtAngleEllipse;
			else if (("at" === strName) || ("A" === strName))		return RulesType.rtArcTo;
			else if (("ar" === strName) || ("B" === strName))		return RulesType.rtArc;
			else if (("wa" === strName) || ("W" === strName))		return RulesType.rtClockwiseArcTo;
			else if (("wr" === strName) || ("V" === strName))		return RulesType.rtClockwiseArc;
			else if (("qx" === strName) || ("X" === strName))		return RulesType.rtEllipticalQuadrX;
			else if (("qy" === strName) || ("Y" === strName))		return RulesType.rtEllipticalQuadrY;
			else if (("qb" === strName) || ("Q" === strName))		return RulesType.rtQuadrBesier;
			else oRes.bRes = false;

			return RulesType.rtEnd;
		}
		function CPPTFormula(sFormula) {
			this.m_eFormulaType = PPTFormulaType.ftSum;
			this.m_lIndex = 0;

			this.m_lParam1 = 0;
			this.m_eType1 = ParamType.ptValue;

			this.m_lParam2 = 0;
			this.m_eType2 = ParamType.ptValue;

			this.m_lParam3 = 0;
			this.m_eType3 = ParamType.ptValue;

			this.m_lCountRecurs = 0;

			this.FromString(sFormula);
		}
		CPPTFormula.prototype.FromString = function(sFormula, lShapeWidth, lShapeHeight) {
			if(!sFormula) {
				return;
			}
			let aParams = sFormula.split(" ");
			let nCount = aParams.length;
			let oRes = {};
			this.m_eFormulaType = this.GetFormula(aParams[0]);

			oRes.ptType = ParamType.ptValue;
			if (1 < nCount) {
				this.m_lParam1 = this.GetValue(aParams[1], oRes, lShapeWidth, lShapeHeight);
				this.m_eType1 = oRes.ptType;
			}
			if (2 < nCount) {
				this.m_lParam2 = this.GetValue(aParams[2], oRes, lShapeWidth, lShapeHeight);
				this.m_eType2 = oRes.ptType;
			}
			if (3 < nCount) {
				this.m_lParam3 = this.GetValue(aParams[3], oRes, lShapeWidth, lShapeHeight);
				this.m_eType3 = oRes.ptType;
			}
		};
		CPPTFormula.prototype.GetFormula = function(strName) {
			if		("sum" === strName)									return PPTFormulaType.ftSum;
			else if (("prod" === strName) || ("product" === strName)) return PPTFormulaType.ftProduct;
			else if ("mid" === strName)									return PPTFormulaType.ftMid;
			else if (("absolute" === strName) || ("abs" === strName)) return PPTFormulaType.ftAbsolute;
			else if ("min" === strName)									return PPTFormulaType.ftMin;
			else if ("max" === strName)									return PPTFormulaType.ftMax;
			else if ("if" === strName)									return PPTFormulaType.ftIf;
			else if ("sqrt" === strName)									return PPTFormulaType.ftSqrt;
			else if ("mod" === strName)									return PPTFormulaType.ftMod;
			else if ("sin" === strName)									return PPTFormulaType.ftSin;
			else if ("cos" === strName)									return PPTFormulaType.ftCos;
			else if ("tan" === strName)									return PPTFormulaType.ftTan;
			else if ("atan2" === strName)								return PPTFormulaType.ftAtan2;
			else if ("sinatan2" === strName)								return PPTFormulaType.ftSinatan2;
			else if ("cosatan2" === strName)								return PPTFormulaType.ftCosatan2;
			else if ("sumangle" === strName)								return PPTFormulaType.ftSumangle;
			else if ("ellipse" === strName)								return PPTFormulaType.ftEllipse;
			else if ("val" === strName)									return PPTFormulaType.ftVal;
		};
		CPPTFormula.prototype.GetValue = function(strParam, oRes, lShapeWidth, lShapeHeight) {
			return GetValue(strParam, oRes, lShapeWidth, lShapeHeight);
		};


		function IS_ALPHA(c) {
			return (((c >= 'a') && (c <= 'z')) || ((c >= 'A') && (c <= 'Z')));
		}
		function IS_DIGIT(c) {
			return (((c >= '0') && (c <= '9')) || (c === '-'));
		}

		const RulesType =
		{
			// VML
				rtLineTo			: 0,	// 2*
				rtCurveTo			: 1,	// 6*
				rtMoveTo			: 2,	// 2

				rtClose				: 3,	// 0
				rtEnd				: 4,	// 0

				rtRMoveTo			: 5,	// 2*
				rtRLineTo			: 6,	// 2*
				rtRCurveTo			: 7,	// 6*

				rtNoFill			: 8,	// 0
				rtNoStroke			: 9,	// 0

				rtAngleEllipseTo	: 10,	// 6*
				rtAngleEllipse		: 11,	// 6*

				rtArc				: 12,	// 8*
				rtArcTo				: 13,	// 8*

				rtClockwiseArcTo	: 14,	// 8*
				rtClockwiseArc		: 15,	// 8*

				rtEllipticalQuadrX	: 16,	// 2*
				rtEllipticalQuadrY	: 17,	// 2*

				rtQuadrBesier		: 18,	// 2 + 2*

				rtFillColor			: 20,
				rtLineColor			: 21,

				// OOXML
				rtOOXMLMoveTo		: 0 + 100,	// 2
				rtOOXMLLineTo		: 1 + 100,	// 2*
				rtOOXMLCubicBezTo	: 2 + 100,	// 6*
				rtOOXMLArcTo		: 3 + 100,	// 8*
				rtOOXMLQuadBezTo	: 4 + 100,	// 2 + 2*
				rtOOXMLClose		: 5 + 100,	// 0
				rtOOXMLEnd			: 6	+ 100	// 0
		};

		function CSlicePath(eRuler) {
			this.m_nCountElementsPoint = 0;
			this.m_lX = 0;
			this.m_lY = 0;
			this.m_eRuler = AscFormat.isRealNumber(eRuler) ? eRuler : RulesType.rtMoveTo;
			this.m_arPoints = [];
			this.m_arPointsType = [];
		}
		CSlicePath.prototype.AddParam = function(lParam, eParType) {
			let lPoint = this.m_nCountElementsPoint % 2;
			if (0 === lPoint)
			{
				let point = {};
				let pointType = {};
				point.x = lParam;
				if (this.m_eRuler !==  RulesType.rtRMoveTo && this.m_eRuler !==  RulesType.rtRLineTo && this.m_eRuler !==  RulesType.rtRCurveTo)
				{
					point.x -= this.m_lX;
				}
				point.y = 0;
				pointType.x = eParType;
				pointType.y = ParamType.ptValue;
				this.m_arPoints.push(point);
				this.m_arPointsType.push(pointType);
			}
			else
			{
				this.m_arPoints[this.m_arPoints.length - 1].y = lParam;
				if (this.m_eRuler !==  RulesType.rtRMoveTo && this.m_eRuler !==  RulesType.rtRLineTo && this.m_eRuler !==  RulesType.rtRCurveTo)
				{
					this.m_arPoints[this.m_arPoints.length - 1].y -= this.m_lY;
				}

				this.m_arPointsType[this.m_arPointsType.length - 1].y = eParType;
			}
			++this.m_nCountElementsPoint;
		};


		function SPointType()
		{
			this.x = 0;
			this.y = 0;
		}
		function SPointExist()
		{
			this.x = false;
			this.y = false;
		}

		function SHandle()
		{
			this.gdRef = new SPointType();
			this.gdRefType = new SPointType();
			this.bRefExist = new SPointExist();
			this.bRefPolarExist = new SPointExist();

			this.Max = new SPointType();
			this.MaxType = new SPointType();
			this.bMaxExist = new SPointExist();
			this.bMaxPolarExist = new SPointExist();

			this.Min  = new SPointType();
			this.MinType  = new SPointType();
			this.bMinExist = new SPointExist();
			this.bMinPolarExist = new SPointExist();

			this.Pos = new SPointType();
			this.PosType = new SPointType();

			this.PolarCentre = new SPointType();
			this.PolarCentreType = new SPointType();
		}

		function CVmlGeometryData() {
			this.type = null;
			this.adjustments = [];
			this.absMaxAdjustments = [];
			this.path = null;
			this.guides = [];
			this.connectors = [];
			this.textRect = null;
			this.handles = [];
			this.connectorsAngles = [];

			this.concentricFill = null;
			this.join = null;

			this.limoX = null;
			this.limoY = null;

			this.m_lWidth = 0;
			this.m_lHeight = 0;

			this.m_lIndexDst = 0;
			this.m_lIndexSrc = -1;
			this.m_arIndexDst = [];
			this.m_lMaxAdjUse = -1;

			this.m_lCoordSizeX = 21600;
			this.m_lCoordSizeY = 21600;

			this.pCurPoint = {};
			this.pCurPointType = {};
			this.pCurPoint1 = {};
			this.pCurPointType1 = {};
			this.pTmpPoint = {};
			this.m_oParam = {};
			this.m_oParam.m_lParam = 65536;
			this.m_oParam.m_lCoef = 65536;
			this.m_oParam.m_eType = ParamType.ptValue;
			this.geometry = null;
		}
		CVmlGeometryData.prototype.fillByType = function(nType) {
			this.type = nType;
			if(this.type === null) {
				return;
			}
			switch(this.type) {
				case ShapeType.sptCAccentBorderCallout90: this.fillCAccentBorderCallout90Type(); break;
				case ShapeType.sptCAccentBorderCallout1: this.fillCAccentBorderCallout1Type(); break;
				case ShapeType.sptCAccentBorderCallout2: this.fillCAccentBorderCallout2Type(); break;
				case ShapeType.sptCAccentBorderCallout3: this.fillCAccentBorderCallout3Type(); break;

				case ShapeType.sptCAccentCallout90: this.fillCAccentCallout90Type(); break;
				case ShapeType.sptCAccentCallout1: this.fillCAccentCallout1Type(); break;
				case ShapeType.sptCAccentCallout2: this.fillCAccentCallout2Type(); break;
				case ShapeType.sptCAccentCallout3: this.fillCAccentCallout3Type(); break;

				case ShapeType.sptCBorderCallout90: this.fillCBorderCallout90Type(); break;
				case ShapeType.sptCBorderCallout1: this.fillCBorderCallout1Type(); break;
				case ShapeType.sptCBorderCallout2: this.fillCBorderCallout2Type(); break;
				case ShapeType.sptCBorderCallout3: this.fillCBorderCallout3Type(); break;

				case ShapeType.sptCCallout90: this.fillCCallout90Type(); break;
				case ShapeType.sptCCallout1: this.fillCCallout1Type(); break;
				case ShapeType.sptCCallout2: this.fillCCallout2Type(); break;
				case ShapeType.sptCCallout3: this.fillCCallout3Type(); break;

				case ShapeType.sptCActionButtonBlank: this.fillCActionButtonBlankType(); break;
				case ShapeType.sptCActionButtonHome: this.fillCActionButtonHomeType(); break;
				case ShapeType.sptCActionButtonHelp: this.fillCActionButtonHelpType(); break;
				case ShapeType.sptCActionButtonInformation: this.fillCActionButtonInfoType(); break;
				case ShapeType.sptCActionButtonBackPrevious: this.fillCActionButtonBackType(); break;
				case ShapeType.sptCActionButtonForwardNext: this.fillCActionButtonNextType(); break;
				case ShapeType.sptCActionButtonBeginning: this.fillCActionButtonBeginType(); break;
				case ShapeType.sptCActionButtonEnd: this.fillCActionButtonEndType(); break;
				case ShapeType.sptCActionButtonReturn: this.fillCActionButtonReturnType(); break;
				case ShapeType.sptCActionButtonDocument: this.fillCActionButtonDocType(); break;
				case ShapeType.sptCActionButtonSound: this.fillCActionButtonSoundType(); break;
				case ShapeType.sptCActionButtonMovie: this.fillCActionButtonMovieType(); break;

				case ShapeType.sptCArc: this.fillCArcType(); break;
				case ShapeType.sptCLine: this.fillCLineType(); break;

				case ShapeType.sptCBentArrow: this.fillCBentArrowType(); break;
				case ShapeType.sptCBentUpArrow: this.fillCBentUpArrowType(); break;
				case ShapeType.sptCBevel: this.fillCBevelType(); break;
				case ShapeType.sptCBlockArc: this.fillCBlockArcType(); break;
				case ShapeType.sptCBracePair: this.fillCBracePairType(); break;
				case ShapeType.sptCBracketPair: this.fillCBracketPairType(); break;

				case ShapeType.sptCCan: this.fillCCanType(); break;
				case ShapeType.sptCChevron: this.fillCChevronType(); break;
				case ShapeType.sptCCircularArrow: this.fillCCircularArrowType(); break;
				case ShapeType.sptCCloudCallout: this.fillCCloudCalloutType(); break;
				case ShapeType.sptCCube: this.fillCCubeType(); break;
				case ShapeType.sptCCurvedDownArrow: this.fillCCurvedDownArrowType(); break;
				case ShapeType.sptCCurvedLeftArrow: this.fillCCurvedLeftArrowType(); break;
				case ShapeType.sptCCurvedRightArrow: this.fillCCurvedRightArrowType(); break;
				case ShapeType.sptCCurvedUpArrow: this.fillCCurvedUpArrowType(); break;

				case ShapeType.sptCDiamond: this.fillCDiamondType(); break;
				case ShapeType.sptCDonut: this.fillCDonutType(); break;
				case ShapeType.sptCDownArrowCallout: this.fillCDownArrowCalloutType(); break;
				case ShapeType.sptCDownArrow: this.fillCDownArrowType(); break;

				case ShapeType.sptCEllipse: this.fillCEllipseType(); break;
				case ShapeType.sptCEllipseRibbon: this.fillCEllipseRibbonType(); break;
				case ShapeType.sptCEllipseRibbon2: this.fillCEllipseRibbon2Type(); break;

				case ShapeType.sptCFlowChartAlternateProcess: this.fillCFlowChartAlternateProcessType(); break;
				case ShapeType.sptCFlowChartCollate: this.fillCFlowChartCollateType(); break;
				case ShapeType.sptCFlowChartConnector: this.fillCFlowChartConnectorType(); break;
				case ShapeType.sptCFlowChartDecision: this.fillCFlowChartDecisionType(); break;
				case ShapeType.sptCFlowChartDisplay: this.fillCFlowChartDisplayType(); break;
				case ShapeType.sptCFlowChartDelay: this.fillCFlowChartDelayType(); break;
				case ShapeType.sptCFlowChartDocument: this.fillCFlowChartDocumentType(); break;
				case ShapeType.sptCFlowChartExtract: this.fillCFlowChartExtractType(); break;
				case ShapeType.sptCFlowChartInputOutput: this.fillCFlowChartInputOutputType(); break;
				case ShapeType.sptCFlowChartInternalStorage: this.fillCFlowChartInternalStorageType(); break;
				case ShapeType.sptCFlowChartMagneticDisk: this.fillCFlowChartMagneticDiskType(); break;
				case ShapeType.sptCFlowChartMagneticDrum: this.fillCFlowChartMagneticDrumType(); break;
				case ShapeType.sptCFlowChartMagneticTape: this.fillCFlowChartMagneticTapeType(); break;
				case ShapeType.sptCFlowChartManualInput: this.fillCFlowChartManualInputType(); break;
				case ShapeType.sptCFlowChartManualOperation: this.fillCFlowChartManualOperationType(); break;
				case ShapeType.sptCFlowChartMerge: this.fillCFlowChartMergeType(); break;
				case ShapeType.sptCFlowChartMultidocument: this.fillCFlowChartMultidocumentType(); break;
				case ShapeType.sptCFlowChartOffpageConnector: this.fillCFlowChartOffpageConnectorType(); break;
				case ShapeType.sptCFlowChartOnlineStorage: this.fillCFlowChartOnlineStorageType(); break;
				case ShapeType.sptCFlowChartOr: this.fillCFlowChartOrType(); break;
				case ShapeType.sptCFlowChartPredefinedProcess: this.fillCFlowChartPredefinedProcessType(); break;
				case ShapeType.sptCFlowChartPreparation: this.fillCFlowChartPreparationType(); break;
				case ShapeType.sptCFlowChartProcess: this.fillCFlowChartProcessType(); break;
				case ShapeType.sptCFlowChartPunchedCard: this.fillCFlowChartPunchedCardType(); break;
				case ShapeType.sptCFlowChartPunchedTape: this.fillCFlowChartPunchedTapeType(); break;
				case ShapeType.sptCFlowChartSort: this.fillCFlowChartSortType(); break;
				case ShapeType.sptCFlowChartSummingJunction: this.fillCFlowChartSummingJunctionType(); break;
				case ShapeType.sptCFlowChartTerminator: this.fillCFlowChartTerminatorType(); break;
				case ShapeType.sptCFoldedCorner: this.fillCFoldedCornerType(); break;

				case ShapeType.sptCHeart: this.fillCHeartType(); break;
				case ShapeType.sptCHexagon: this.fillCHexagonType(); break;
				case ShapeType.sptCHomePlate: this.fillCHomePlateType(); break;

				case ShapeType.sptCIrregularSeal1: this.fillCIrregularSealOneType(); break;
				case ShapeType.sptCIrregularSeal2: this.fillCIrregularSealTwo(); break;
				case ShapeType.sptCIsocelesTriangle: this.fillCIsoscelesTriangleType(); break;

				case ShapeType.sptCLeftArrowCallout: this.fillCLeftArrowCalloutType(); break;
				case ShapeType.sptCLeftArrow: this.fillCLeftArrowType(); break;
				case ShapeType.sptCLeftBrace: this.fillCLeftBraceType(); break;
				case ShapeType.sptCLeftBracket: this.fillCLeftBracketType(); break;
				case ShapeType.sptCLeftRightArrowCallout: this.fillCLeftRightArrowCalloutType(); break;
				case ShapeType.sptCLeftRightArrow: this.fillCLeftRightArrowType(); break;
				case ShapeType.sptCLeftRightUpArrow: this.fillCLeftRightUpArrow(); break;
				case ShapeType.sptCLeftUpArrow: this.fillCLeftUpArrowType(); break;
				case ShapeType.sptCLightningBolt: this.fillCLightningBoltType(); break;

				case ShapeType.sptCMoon: this.fillCMoonType(); break;

				case ShapeType.sptCNoSmoking: this.fillCNoSmokingType(); break;
				case ShapeType.sptCNotchedRightArrow: this.fillCNotchedRightArrowType(); break;

				case ShapeType.sptCOctagon: this.fillCOctagonType(); break;

				case ShapeType.sptCParallelogram: this.fillCParallelogramType(); break;
				case ShapeType.sptCPentagon: this.fillCPentagonType(); break;
				case ShapeType.sptCPlaque: this.fillCPlaqueType(); break;
				case ShapeType.sptCPlus: this.fillCPlusType(); break;

				case ShapeType.sptCQuadArrowCallout: this.fillCQuadArrowCalloutType(); break;
				case ShapeType.sptCQuadArrow: this.fillCQuadArrowType(); break;

				case ShapeType.sptCRect: this.fillCRectangleType(); break;
				case ShapeType.sptCRibbon: this.fillCRibbonDownType(); break;
				case ShapeType.sptCRibbon2: this.fillCRibbonUpType(); break;
				case ShapeType.sptCRightArrowCallout: this.fillCRightArrowCalloutType(); break;
				case ShapeType.sptCRightArrow: this.fillCRightArrowType(); break;
				case ShapeType.sptCRightBrace: this.fillCRightBracetype(); break;
				case ShapeType.sptCRightBracket: this.fillCRightBracketType(); break;
				case ShapeType.sptCRtTriangle: this.fillCRightTriangleType(); break;
				case ShapeType.sptCRoundRect: this.fillCRoundedRectangleType(); break;

				case ShapeType.sptCStar16: this.fillCStar16Type(); break;
				case ShapeType.sptCStar24: this.fillCStar24Type(); break;
				case ShapeType.sptCStar32: this.fillCStar32Type(); break;
				case ShapeType.sptCStar4: this.fillCStar4Type(); break;
				case ShapeType.sptCStar8: this.fillCStar8Type(); break;
				case ShapeType.sptCSmileyFace: this.fillCSmileyFaceType(); break;
				case ShapeType.sptCStar5: this.fillCStar2Type(); break;
				case ShapeType.sptCStraightConnector1: this.fillCStraightConnectorType(); break;
				case ShapeType.sptCStripedRightArrow: this.fillCStripedRightArrowType(); break;
				case ShapeType.sptCSun: this.fillCSunType(); break;

				case ShapeType.sptCTextBox: this.fillCTextboxType(); break;
				case ShapeType.sptCTrapezoid: this.fillCTrapezoidType(); break;

				case ShapeType.sptCUpArrowCallout: this.fillCUpArrowCalloutType(); break;
				case ShapeType.sptCUpArrow: this.fillCUpArrowType(); break;
				case ShapeType.sptCUpDownArrowCallout: this.fillCUpDownArrowCalloutType(); break;
				case ShapeType.sptCUpDownArrow: this.fillCUpDownArrowType(); break;
				case ShapeType.sptCUturnArrow: this.fillCUturnArrowType(); break;

				case ShapeType.sptCVerticalScroll: this.fillCVerticalScrollType(); break;
				case ShapeType.sptCHorizontalScroll: this.fillCHorizontalScrollType(); break;

				case ShapeType.sptCWedgeEllipseCallout: this.fillCWedgeEllipseCalloutType(); break;
				case ShapeType.sptCWedgeRectCallout: this.fillCWedgeRectCalloutType(); break;
				case ShapeType.sptCWedgeRoundRectCallout: this.fillCWedgeRoundedRectCalloutType(); break;

				case ShapeType.sptCWave: this.fillCWaveType(); break;
				case ShapeType.sptCDoubleWave: this.fillCWaveDoubleType(); break;

				case ShapeType.sptCBentConnector2:
				case ShapeType.sptCBentConnector3:
				case ShapeType.sptCBentConnector4:
				case ShapeType.sptCBentConnector5:
				{
					this.fillCBentConnectorType();
					break;
				}
				case ShapeType.sptCCurvedConnector2:
				case ShapeType.sptCCurvedConnector3:
				case ShapeType.sptCCurvedConnector4:
				case ShapeType.sptCCurvedConnector5:
				{
					this.fillCCurvedConnectorType();
					break;
				}

				case ShapeType.sptCTextPlain:
				case ShapeType.sptCTextStop:
				case ShapeType.sptCTextTriangle:
				case ShapeType.sptCTextTriangleInverted:
				case ShapeType.sptCTextChevron:
				case ShapeType.sptCTextChevronInverted:
				case ShapeType.sptCTextRingInside:
				case ShapeType.sptCTextRingOutside:
				case ShapeType.sptCTextArchUp:
				case ShapeType.sptCTextArchDown:
				case ShapeType.sptCTextCircle:
				case ShapeType.sptCTextButton:
				case ShapeType.sptCTextArchUpPour:
				case ShapeType.sptCTextArchDownPour:
				case ShapeType.sptCTextCirclePour:
				case ShapeType.sptCTextButtonPour:
				case ShapeType.sptCTextCurveUp:
				case ShapeType.sptCTextCurveDown:
				case ShapeType.sptCTextCascadeUp:
				case ShapeType.sptCTextCascadeDown:
				case ShapeType.sptCTextWave1:
				case ShapeType.sptCTextWave2:
				case ShapeType.sptCTextWave3:
				case ShapeType.sptCTextWave4:
				case ShapeType.sptCTextInflate:
				case ShapeType.sptCTextDeflate:
				case ShapeType.sptCTextInflateBottom:
				case ShapeType.sptCTextDeflateBottom:
				case ShapeType.sptCTextInflateTop:
				case ShapeType.sptCTextDeflateTop:
				case ShapeType.sptCTextDeflateInflate:
				case ShapeType.sptCTextDeflateInflateDeflate:
				case ShapeType.sptCTextFadeRight:
				case ShapeType.sptCTextFadeLeft:
				case ShapeType.sptCTextFadeUp:
				case ShapeType.sptCTextFadeDown:
				case ShapeType.sptCTextSlantUp:
				case ShapeType.sptCTextSlantDown:
				case ShapeType.sptCTextCanUp:
				case ShapeType.sptCTextCanDown:
				{
					this.fillCTextboxType();
					break;
				}
			}
		};
		CVmlGeometryData.prototype.loadConnectorsList = function(sList) {
			this.connectors = sList.split(";");
		};
		CVmlGeometryData.prototype.loadTextRect = function(sRect) {
			this.textRect = sRect;
		};
		CVmlGeometryData.prototype.addConnectorAngle = function(nAngle) {
			this.connectorsAngles.push(nAngle);
		};
		CVmlGeometryData.prototype.addAdjustment = function(nAdj) {
			this.adjustments.push(nAdj);
		};
		CVmlGeometryData.prototype.addAbsMaxAdjustment = function(nAdj) {
			this.absMaxAdjustments.push(nAdj);
		};
		CVmlGeometryData.prototype.loadPath = function(sPath) {
			this.path = sPath;
		};
		CVmlGeometryData.prototype.loadCoordSize = function(oCoordSize) {
			if(oCoordSize && oCoordSize.x > 0 && oCoordSize.y > 0) {

				this.m_lCoordSizeX = oCoordSize.x;
				this.m_lCoordSizeY = oCoordSize.y;
			}
		};
		CVmlGeometryData.prototype.parseFormula = function(sFormula) {
			return new CPPTFormula(sFormula);
		};
		CVmlGeometryData.prototype.addGuide = function(sGuide) {
			this.guides.push(this.parseFormula(sGuide));
		};
		CVmlGeometryData.prototype.addHandle = function(oHandle) {
			this.handles.push(oHandle);
		};

		CVmlGeometryData.prototype.convertToOOXML = function() {
			if(this.type === ShapeType.sptCRect ||  this.type === ShapeType.sptCTextBox ) {
				return AscFormat.CreateGeometry("rect");
			}
			if(this.type === ShapeType.sptCLine) {
				return AscFormat.CreateGeometry("line");
			}
			if(!(typeof this.path === "string" && this.path.length > 0)) {
				return null;
			}

			this.m_lIndexDst = 0;
			this.geometry = new AscFormat.Geometry();
			this.convertGuides();
			this.convertPaths();
			this.convertTextRect();
			this.convertHandles();
			return this.geometry;
		};
		CVmlGeometryData.prototype.convertPaths = function() {
			let aStringPaths = this.path.split('e');
			let nPathsCount = aStringPaths.length;
			for(let nPath = 0; nPath < nPathsCount; nPath++) {
				this.convertPath(aStringPaths[nPath]);
			}
		};
		CVmlGeometryData.prototype.parsePath = function (strSource) {

			let pArrayResults = [];
			let strPath = strSource;
			let nLength = strPath.length;
			if (-1 !== strPath.indexOf('h'))
			{
				let pBuff = "";

				let nCur = 0;
				for (let i = 1; i < nLength; ++i)
				{
					let _c = strPath.charAt(i - 1);
					if (_c !== 'h') {
						pBuff += _c;
					}
					else
					{
						let _c1 = strPath.charAt(i);
						if (_c1 === 'a' ||
							_c1 === 'b' ||
							_c1 === 'c' ||
							_c1 === 'd' ||
							_c1 === 'e' ||
							_c1 === 'f' ||
							_c1 === 'g' ||
							_c1 === 'h' ||
							_c1 === 'i')
						{
							++i;
						}
					}
				}

				if (nLength > 0)
					pBuff += strPath.charAt(nLength - 1);


				strPath = pBuff;
				nLength = strPath.length;
			}

			if (nLength > 0 && strPath.charAt(nLength - 1) === ',') {
				strPath += "0";
				++nLength;
			}

			let nIndexOld = 0;
			for (let nIndex = 0; nIndex < nLength; ++nIndex)
			{
				if (nIndex === (nLength - 1))
				{
					pArrayResults.push(strPath.substring(nIndexOld));
					//continue;
				}

				let _c	= strPath[nIndex];
				let _c1	= strPath[nIndex + 1];

				if (_c1 === ',')
				{
					if (',' === _c)
					{
						pArrayResults.push("0");
					}
					else if (IS_ALPHA(_c))
					{
						pArrayResults.push(strPath.substring(nIndexOld, nIndex + 1));
						pArrayResults.push("0");
					}
					else if (IS_DIGIT(_c))
					{
						pArrayResults.push(strPath.substring(nIndexOld, nIndex + 1));
					}
				}
				else if (',' === _c)
				{
					if (IS_ALPHA(_c1))
					{
						pArrayResults.push("0");
						nIndexOld = nIndex + 1;
					}
					else if (IS_DIGIT(_c1))
					{
						nIndexOld = nIndex + 1;
					}
				}
				else
				{
					let _isA = IS_ALPHA(_c);
					let _isD = _isA ? false : IS_DIGIT(_c);

					let _isA1 = IS_ALPHA(_c1);
					let _isD1 = _isA1 ? false : IS_DIGIT(_c1);

					if (_isA && _isD1)
					{
						pArrayResults.push(strPath.substring(nIndexOld, nIndex + 1));
						nIndexOld = nIndex + 1;
					}
					else if (_isD && _isA1)
					{
						pArrayResults.push(strPath.substring(nIndexOld, nIndex + 1));
						nIndexOld = nIndex + 1;
					}
					else if (_isD && ('@' === _c1))
					{
						pArrayResults.push(strPath.substring(nIndexOld, nIndex + 1));

						++nIndex;
						nIndexOld = nIndex;
					}
					else if (_isD && ('#' === _c1))
					{
						pArrayResults.push(strPath.substring(nIndexOld, nIndex + 1));

						++nIndex;
						nIndexOld = nIndex;
					}
					else if (_isA && ('@' === _c1))
					{
						pArrayResults.push(strPath.substring(nIndexOld, nIndex + 1));

						++nIndex;
						nIndexOld = nIndex;
					}
					else if (_isA && ('#' === _c1))
					{
						pArrayResults.push(strPath.substring(nIndexOld, nIndex + 1));

						++nIndex;
						nIndexOld = nIndex;
					}
					else if (('x' === _c) && _isA1)
					{
						pArrayResults.push("x");
						nIndexOld = nIndex + 1;
					}
				}
			}
			return pArrayResults;
		};
		CVmlGeometryData.prototype.convertPath = function(sPath) {
			let oArray = this.parsePath(sPath);
			let oRes = {};
			let arSlicesPath = [];
			let bFill = true;
			let bStroke = true;
			let nIndex = sPath.indexOf("nf");
			if (-1 !== nIndex)
			{
				bFill = false;
			}
			nIndex = sPath.indexOf("ns");
			if (-1 !== nIndex)
			{
				bStroke = false;
			}

			nIndex = sPath.indexOf("F");
			if (-1 !== nIndex)
			{
				bFill = false;
			}

			nIndex = sPath.indexOf("S");
			if (-1 !== nIndex)
			{
				bStroke = false;
			}


			for (let nIndex = 0; nIndex < oArray.length; ++nIndex)
			{
				let str = oArray[nIndex];
				let lValue = GetValue(str, oRes);
				if (oRes.bRes)
				{
					if (0 !== arSlicesPath.length)
					{
						arSlicesPath[arSlicesPath.length - 1].AddParam(lValue, oRes.ptType);
					}
				}
				else
				{

					let eRuler = GetRuler(str, oRes);
					if (oRes.bRes)
					{
						if (RulesType.rtNoFill === eRuler)
						{
							bFill = false;
						}
						else if (RulesType.rtNoStroke === eRuler)
						{
							bStroke = false;
						}
						else
						{
							let oSlice = new CSlicePath(eRuler);
							arSlicesPath.push(oSlice);
						}
					}
				}
			}
			let oOOXMLPath = new AscFormat.Path();
			oOOXMLPath.setPathW(this.m_lCoordSizeX);
			oOOXMLPath.setPathH(this.m_lCoordSizeY);
			if(!bFill) {
				oOOXMLPath.setFill("none");
			}
			else {
				oOOXMLPath.setFill("norm");
			}
			if(!bStroke) {
				oOOXMLPath.setStroke(false);
			}
			else {
				oOOXMLPath.setStroke(true);
			}
			oOOXMLPath.setExtrusionOk(false);
			for(let nSlice = 0; nSlice < arSlicesPath.length; ++nSlice) {
				let oSlice = arSlicesPath[nSlice];
				switch (oSlice.m_eRuler)
				{
					case RulesType.rtMoveTo:
					{
						this.ConvertSlice_MoveTo(oSlice, oOOXMLPath);
						break;
					}
					case RulesType.rtRMoveTo:
					{
						this.ConvertSlice_RMoveTo(oSlice, oOOXMLPath);
						break;
					}
					case RulesType.rtClose:
					{
						oOOXMLPath.close();
						break;
					}
					case RulesType.rtLineTo:
					{
						this.ConvertSlice_LineTo(oSlice, oOOXMLPath);
						break;
					}
					case RulesType.rtRLineTo:
					{
						this.ConvertSlice_RLineTo(oSlice, oOOXMLPath);
						break;
					}
					case RulesType.rtArcTo:
					case RulesType.rtArc:
					{
						this.ConvertSlice_ArcTo(oSlice, oOOXMLPath);
						break;
					}
					case RulesType.rtClockwiseArcTo:
					case RulesType.rtClockwiseArc:
					{
						this.ConvertSlice_ClockwiseTo(oSlice, oOOXMLPath);
						break;
					}
					case RulesType.rtQuadrBesier:
					{
						this.ConvertSlice_QuadrBesier(oSlice, oOOXMLPath);
						break;
					}
					case RulesType.rtCurveTo:
					{
						this.ConvertSlice_CurveTo(oSlice, oOOXMLPath);
						break;
					}
					case RulesType.rtRCurveTo:
					{
						this.ConvertSlice_RCurveTo(oSlice, oOOXMLPath);
						break;
					}
					case RulesType.rtAngleEllipse:
					case RulesType.rtAngleEllipseTo:
					{
						this.ConvertSlice_AngleEllipse(oSlice, oOOXMLPath);
						break;
					}
					case RulesType.rtEllipticalQuadrX:
					{
						this.ConvertSlice_EllipticalQuadrX(oSlice, oOOXMLPath);
						break;
					}
					case RulesType.rtEllipticalQuadrY:
					{
						this.ConvertSlice_EllipticalQuadrY(oSlice, oOOXMLPath);
						break;
					}
					default:
						break;
				}
			}

			if(oOOXMLPath.ArrPathCommandInfo.length > 0) {
				this.geometry.AddPath(oOOXMLPath);
			}

 		};
		CVmlGeometryData.prototype.addGeomGuide = function(name, formula, x, y, z) {
			this.geometry.AddGuide(name, AscFormat.MAP_FMLA_TO_TYPE[formula], x, y, z)
		};
		CVmlGeometryData.prototype.convertProdString = function(strParam1, strParam2, lParam3)
		{
			let lRes;
			if ('#' === strParam2[0])
			{
				lRes = parseInt(strParam2.substring(1));
				strParam2 = this.GetValue(lRes, ParamType.ptAdjust, false);
			}
			else if ('@' === strParam2[0])
			{
				lRes = parseInt(strParam2.substring(1));
				strParam2 = this.GetValue(lRes, ParamType.ptFormula, false);
			}
			this.addGeomGuide(this.getNextGdName(), "*/", strParam1, strParam2, this.GetValue(lParam3, ParamType.ptValue, false));
		};
		CVmlGeometryData.prototype.convertTextRect = function() {
			let strRect = this.textRect;
			if (!strRect || strRect.length === 0)
				return;

			let arBorder = strRect.split(",");

			this.m_lIndexSrc++;
			if (arBorder.length > 0 && !arBorder[0].length !== 0)
				this.convertProdString("w",  arBorder[0], this.m_lCoordSizeX);
		else
			this.convertProdString("w",  "0", this.m_lCoordSizeX);

			if (arBorder.length > 1 && !arBorder[1].length !== 0)
				this.convertProdString("h",  arBorder[1], this.m_lCoordSizeY);
		else
			this.convertProdString("h",  "0", this.m_lCoordSizeY);

			if (arBorder.length > 2 && !arBorder[2].length !== 0)
				this.convertProdString("w",  arBorder[2], this.m_lCoordSizeX);
		else
			this.convertProdString("w",  "" + (this.m_lCoordSizeX), this.m_lCoordSizeX);

			if (arBorder.length > 3 && !arBorder[3].length !== 0)
				this.convertProdString("h",  arBorder[3], this.m_lCoordSizeY);
		else
			this.convertProdString("h",  "" + (this.m_lCoordSizeY), this.m_lCoordSizeY);

			this.m_arIndexDst.push(this.m_lIndexDst);

			let l = this.GetValue(this.m_lIndexDst - 4, ParamType.ptFormula, true) + "";
			let t = this.GetValue(this.m_lIndexDst - 3, ParamType.ptFormula, true) + "";
			let r = this.GetValue(this.m_lIndexDst - 2, ParamType.ptFormula, true) + "";
			let b = this.GetValue(this.m_lIndexDst - 1, ParamType.ptFormula, true) + "";
			this.geometry.AddRect(l, t, r, b);
		};

		CVmlGeometryData.prototype.GetHandleValue = function(strParam, lVal, oRes)
		{
			oRes.ptType = ParamType.ptValue;
			if ('#' === strParam[0])
			{
				oRes.ptType = ParamType.ptAdjust;
				return parseInt(strParam.substring(1));
			}
			else if ('@' === strParam[0])
			{
				oRes.ptType = ParamType.ptFormula;
				return parseInt(strParam.substr(1));
			}
			else if (!AscFormat.isRealNumber(parseInt(strParam)))
			{
				if ("center" === strParam)
				return lVal/2;
			else if ("bottomRight" === strParam)
				return lVal;
			else
				return 0;
			}
			else
			{
				oRes.ptType = ParamType.ptValue;
				return parseInt(strParam);
			}
		};
		CVmlGeometryData.prototype.GetHandlePos = function( strParam,  strBase,  lSize)
		{
			let lRes;
			let strSize;
			let strFrmla;

			strSize = " " + lSize;

			if ('#' === strParam[0])
			{
				lRes = parseInt(strParam.substring(1));
				strFrmla = this.GetValue(lRes, ParamType.ptAdjust, false) + strSize;
			}
			else if ('&' === strParam[0])
			{
				lRes = parseInt(strParam.substring(1));
				strFrmla = this.GetValue(lRes, ParamType.ptFormula, true) + strSize;
			}
			else if ('@' === strParam[0])
			{
				lRes = parseInt(strParam.substring(1));
				strFrmla = this.GetValue(lRes, ParamType.ptFormula, false) + strSize;
			}
			else if (!AscFormat.isRealNumber(parseInt(strParam)))
			{
				if ("center" === strParam)
				strFrmla = "1 2";
			else if ("topLeft" === strParam)
				strFrmla = "0 1";
			else if ("bottomRight" === strParam)
				strFrmla = "1 1";
			}
			else
			{
				lRes = parseInt(strParam);
				strFrmla = this.GetValue(lRes, ParamType.ptValue, false) + strSize;
			}
			let aParams = strFrmla.split(' ');
			let x = aParams[0];
			let y = aParams[1];
			this.addGeomGuide(this.getNextGdName(), "*/", strBase, x, y);
			this.m_arIndexDst.push(this.m_lIndexDst);
			//this.m_lIndexDst++;

			return this.m_lIndexDst-1;
		};
		CVmlGeometryData.prototype.convertHandles = function() {
			if (this.type === ShapeType.sptCArc) {
				return;
			}
			let nHandlesCount = this.handles.length;
			for (let i = 0; i < nHandlesCount; ++i)
			{
				let pHnPoint = this.handles[i];
				let arPos = [];
				let arRangeX = [];
				let arRangeY = [];
				let arPolar = [];
				let oHandle = new SHandle();
				let ptType;
				let oRes = {};

				if (pHnPoint.position !== null && pHnPoint.position !== "")
				{
					arPos = pHnPoint.position.split(",");

					oHandle.gdRef.x = this.GetHandleValue(arPos[0], this.m_lCoordSizeX, oRes);
					oHandle.gdRefType.x = oRes.ptType;
					if ( oHandle.gdRefType.x === ParamType.ptAdjust)
						oHandle.bRefExist.x = true;

					oHandle.gdRef.y = this.GetHandleValue(arPos[1], this.m_lCoordSizeX, oRes);
					oHandle.gdRefType.y = oRes.ptType;
					if ( oHandle.gdRefType.y === ParamType.ptAdjust)
						oHandle.bRefExist.y = true;


					if (pHnPoint.polar !== null && pHnPoint.polar !== "")
					{
						oHandle.bRefExist.y = false;

						oHandle.bRefPolarExist.y = true;
						oHandle.bMinPolarExist.y = true;
						arPolar = pHnPoint.polar.split(",");

						oHandle.PolarCentre.x = this.GetHandleValue(arPolar[0], this.m_lCoordSizeX, oRes);
						oHandle.PolarCentreType.x = oRes.ptType;

						oHandle.PolarCentre.y = this.GetHandleValue(arPolar[1], this.m_lCoordSizeX, oRes);
						oHandle.PolarCentreType.y = oRes.ptType;

						if (oHandle.gdRefType.y === ParamType.ptAdjust )
						{
							let strNewFmla = this.GetValue (this.m_lIndexDst, ParamType.ptFormula, true);
							let strOldFmla = this.GetValue (oHandle.gdRef.y, oHandle.gdRefType.y, false);
							let nIndex = this.m_lIndexDst;

							//replace guide to adjustment
							let oOldGeometry = this.geometry;
							let aOldGdLst = oOldGeometry.gdLstInfo;
							oOldGeometry.gdLstInfo.length = 0;
							let oNewGeometry = oOldGeometry.createDuplicate();
							this.geometry = oNewGeometry;
							this.m_lIndexSrc++;
							this.convertProd (oHandle.gdRef.y, oHandle.gdRefType.y, this.m_oParam.m_lCoef, ParamType.ptValue, pow3_16, ParamType.ptValue, false, true, false);
							this.m_arIndexDst.push(this.m_lIndexDst-1);
							for(let nGd = 0; nGd < aOldGdLst.length; ++nGd) {
								let oGd = aOldGdLst[nGd];
								let x = oGd.x;
								if(x === strOldFmla) {
									x = strNewFmla;
								}
								let y = oGd.y;
								if(y === strOldFmla) {
									y = strNewFmla;
								}
								let z = oGd.z;
								if(z === strOldFmla) {
									z = strNewFmla;
								}
								this.addGeomGuide(oGd.name, oGd.formula, x, y, z)
							}
							oOldGeometry.gdLstInfo = aOldGdLst;

							if (oHandle.gdRefType.y === ParamType.ptAdjust )
							{
								let lVal = this.adjustments[oHandle.gdRef.y];
								lVal = (lVal * pow3_16 / this.m_oParam.m_lCoef + 0.5) >> 0;
								this.adjustments[oHandle.gdRef.y] = lVal;
							}

							this.m_lIndexSrc++;
							this.convertCos(oHandle.gdRef.x, oHandle.gdRefType.x, oHandle.gdRef.y, oHandle.gdRefType.y, false, false);
							this.convertSin(oHandle.gdRef.x, oHandle.gdRefType.x, oHandle.gdRef.y, oHandle.gdRefType.y, false, false);
							this.convertSum(oHandle.PolarCentre.x, oHandle.PolarCentreType.x, this.m_lIndexDst-2, ParamType.ptFormula, 0, ParamType.ptValue, false, true, false);
							this.convertSum(oHandle.PolarCentre.y, oHandle.PolarCentreType.y, this.m_lIndexDst-2, ParamType.ptFormula, 0, ParamType.ptValue, false, true, false);
							this.m_arIndexDst.push(this.m_lIndexDst-1);

							oHandle.Pos.x = this.GetHandlePos ((this.m_lIndexDst-2) + "", "w",  this.m_lCoordSizeX);
							oHandle.PosType.x = ParamType.ptFormula;

							oHandle.Pos.y = this.GetHandlePos ((this.m_lIndexDst-2) + "", "h",  this.m_lCoordSizeY);
							oHandle.PosType.y = ParamType.ptFormula;
						}

					}
				else//если пришли обычные координаты
					{
						if ((pHnPoint.xrange !== null && pHnPoint.xrange !== "" && oHandle.gdRefType.x !== ParamType.ptAdjust && oHandle.gdRefType.y === ParamType.ptAdjust) ||
						(pHnPoint.yrange !== null && pHnPoint.yrange !== "" && oHandle.gdRefType.x === ParamType.ptAdjust && oHandle.gdRefType.y !== ParamType.ptAdjust))
						{
							oHandle.Pos.x = this.GetHandlePos(arPos[1], "w",  this.m_lCoordSizeX);
							oHandle.PosType.x = ParamType.ptFormula;

							oHandle.Pos.y = this.GetHandlePos(arPos[0], "h",  this.m_lCoordSizeY);
							oHandle.PosType.y = ParamType.ptFormula;
						}
					else
						{
							oHandle.Pos.x = this.GetHandlePos(arPos[0], "w",  this.m_lCoordSizeX);
							oHandle.PosType.x = ParamType.ptFormula;

							oHandle.Pos.y = this.GetHandlePos(arPos[1], "h",  this.m_lCoordSizeY);
							oHandle.PosType.y = ParamType.ptFormula;
						}
					}
				}

				if (pHnPoint.xrange !== null && pHnPoint.xrange !== "")
				{
					if ( oHandle.gdRefType.x !== ParamType.ptAdjust && oHandle.gdRefType.y === ParamType.ptAdjust)
					{
						oHandle.gdRef.x = this.GetHandleValue(arPos[1], this.m_lCoordSizeX, oRes);
						oHandle.gdRefType.x = oRes.ptType;
						oHandle.bRefExist.x = true;

						oHandle.gdRef.y = this.GetHandleValue(arPos[0], this.m_lCoordSizeY, oRes);
						oHandle.gdRefType.y = oRes.ptType;
						oHandle.bRefExist.y = false;
					}

					arRangeX = pHnPoint.xrange.split(",");
					oHandle.Min.x = this.GetHandleValue(arRangeX[0], this.m_lCoordSizeX, oRes);
					oHandle.MinType.x = oRes.ptType;
					if ( oHandle.bRefExist.x)
						oHandle.bMinExist.x = true;

					oHandle.Max.x = this.GetHandleValue(arRangeX[1], this.m_lCoordSizeX, oRes);
					oHandle.MaxType.x = oRes.ptType;
					if ( oHandle.bRefExist.x)
						oHandle.bMaxExist.x = true;
				}

				if (pHnPoint.yrange !== null &&  pHnPoint.yrange !== "")
				{
					if ( oHandle.gdRefType.x === ParamType.ptAdjust && oHandle.gdRefType.y !== ParamType.ptAdjust)
					{
						oHandle.gdRef.x = this.GetHandleValue(arPos[1], this.m_lCoordSizeX, oRes);
						oHandle.gdRefType.x = oRes.ptType;
						oHandle.bRefExist.x = false;

						oHandle.gdRef.y = this.GetHandleValue(arPos[0], this.m_lCoordSizeY, oRes);
						oHandle.gdRefType.y = oRes.ptType;
						oHandle.bRefExist.y = true;
					}

					arRangeY = pHnPoint.yrange.split(",");
					oHandle.Min.y = this.GetHandleValue(arRangeY[0], this.m_lCoordSizeY, oRes);
					oHandle.MinType.y = oRes.ptType;
					if ( oHandle.bRefExist.y )
						oHandle.bMinExist.y = true;

					oHandle.Max.y = this.GetHandleValue(arRangeY[1], this.m_lCoordSizeY, oRes);
					oHandle.MaxType.y = oRes.ptType;
					if ( oHandle.bRefExist.y )
						oHandle.bMaxExist.y = true;
				}

				if (pHnPoint.radiusrange !== null && pHnPoint.radiusrange !== "")
				{
					arPos = pHnPoint.radiusrange.split(",");
					oHandle.Min.x = this.GetHandleValue(arPos[0], this.m_lCoordSizeY, oRes);
					oHandle.MinType.x = oRes.ptType;
					oHandle.bMinPolarExist.x = true;

					oHandle.Max.x = this.GetHandleValue(arPos[1], this.m_lCoordSizeY, oRes);
					oHandle.MaxType.x = oRes.ptType;
					oHandle.bMinPolarExist.y = true;

					if ( oHandle.gdRefType.x === ptAdjust)
					{
						oHandle.bRefPolarExist.x = true;
						oHandle.bRefExist.x = false;
					}
				}

				this.CreateHandle(oHandle);

			}
			this.ConvertAdj ();
		};
		CVmlGeometryData.prototype.CreateHandle = function(oHnd) {
			if (oHnd.bRefPolarExist.x || oHnd.bRefPolarExist.y) {
				let gdRefR, gdRefAng, minR, maxR, minAng, maxAng, posX, posY;
				if (oHnd.bRefPolarExist.x)
				{
					gdRefR = this.GetValue(oHnd.gdRef.x, oHnd.gdRefType.x, false) + "";
				}
				if (oHnd.bRefPolarExist.y)
				{
					gdRefAng = this.GetValue(oHnd.gdRef.y, oHnd.gdRefType.y, true) + "";
				}
				if (oHnd.bMinPolarExist.x)
				{
					minR = this.GetValue(oHnd.Min.x, oHnd.MinType.x, false) + "";
					maxR = this.GetValue(oHnd.Max.x, oHnd.MaxType.x, false) + "";
				}
				if (oHnd.bMinPolarExist.y)
				{
					minAng = "0";
					maxAng = "21600000";
				}

				posX = this.GetValue(oHnd.Pos.x, oHnd.PosType.x, true) + "";
				posY = this.GetValue(oHnd.Pos.y, oHnd.PosType.y, true) + "";
				this.geometry.AddHandlePolar(gdRefAng, minAng, maxAng, gdRefR, minR, maxR, posX, posY)
			}
			else if (oHnd.bRefExist.x || oHnd.bRefExist.y) {

				let gdRefX, gdRefY, minX, maxX, minY, maxY, posX, posY;
				if (oHnd.bRefExist.x)
				{
					gdRefX = this.GetValue(oHnd.gdRef.x, oHnd.gdRefType.x, false) + "";
				}
				if (oHnd.bRefExist.y)
				{
					gdRefY = this.GetValue(oHnd.gdRef.y, oHnd.gdRefType.y, false) + "";
				}
				if (oHnd.bMinExist.x)
				{
					minX = this.GetValue(oHnd.Min.x, oHnd.MinType.x, false) + "";
					maxX = this.GetValue(oHnd.Max.x, oHnd.MaxType.x, false) + "";
				}
				else if (oHnd.bRefExist.x)
				{
					minX="-21474836";
					maxX="21474836";
				}
				if (oHnd.bMinExist.y)
				{
					minY = this.GetValue(oHnd.Min.y, oHnd.MinType.y, false) + "";
					maxY = this.GetValue(oHnd.Max.y, oHnd.MaxType.y, false) + "";
				}
				else if (oHnd.bRefExist.y)
				{
					minY="-21474836";
					maxY="21474836";
				}
				posX = this.GetValue(oHnd.Pos.x, oHnd.PosType.x, true) + "";
				posY = this.GetValue(oHnd.Pos.y, oHnd.PosType.y, true) + "";
				this.geometry.AddHandleXY(gdRefX, minX, maxX, gdRefY, minY, maxY, posX, posY);
			}

		};
		CVmlGeometryData.prototype.ConvertAdj = function() {
			for(let nAdj = 0; nAdj < this.adjustments.length; ++nAdj) {
				this.geometry.AddAdj("adj" + nAdj, "val", this.adjustments[nAdj] + "");
			}
		};
		CVmlGeometryData.prototype.GetValue = function(lParam, eParamType, bExtShape) {
			let strValue;

			switch (eParamType)
			{
				case ParamType.ptFormula:
				{
					if (bExtShape)
						strValue = "" + lParam;
					else
						strValue = "" + this.m_arIndexDst[lParam];
						strValue = "gd" + strValue;
						break;
				}
				case ParamType.ptAdjust:
				{
					strValue = "adj" + lParam;
					break;
				}
				case ParamType.ptValue:
				{
					strValue = "" + lParam;
					break;
				}
				default: break;
			}
			return strValue;
		};
		CVmlGeometryData.prototype.ConvertQuadrX = function(pPoint, pPointType, oOOXMLPath)
		{
			let nIndex = this.m_arIndexDst[this.m_lIndexSrc];

			this.m_lIndexSrc++;

			this.convertSum(pPoint.x, pPointType.x, 0, ParamType.ptValue, nIndex-1, ParamType.ptFormula, false, true, true);
			this.convertSum(pPoint.y, pPointType.y, 0, ParamType.ptValue, nIndex, ParamType.ptFormula, false, true, true);
			this.convertIf (this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, -1, ParamType.ptValue, true, true, true);
			this.convertIf (this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, -1, ParamType.ptValue, true, true, true);
			this.convertProd (this.m_lIndexDst-2, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
			this.convertIf (this.m_lIndexDst-4, ParamType.ptFormula, 16200000, ParamType.ptValue, 5400000, ParamType.ptValue, true, true, true);	//stAng
			this.convertIf (this.m_lIndexDst-2, ParamType.ptFormula, 5400000, ParamType.ptValue, -5400000, ParamType.ptValue, true, true, true);	//swAng
			this.convertProd (this.m_lIndexDst-7, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);
			this.convertProd (this.m_lIndexDst-7, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);
			this.convertIf (this.m_lIndexDst-9, ParamType.ptFormula, this.m_lIndexDst-9, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, true, true, true);//wR
			this.convertIf (this.m_lIndexDst-9, ParamType.ptFormula, this.m_lIndexDst-9, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, true, true, true);//hR

			this.m_arIndexDst.push(this.m_lIndexDst-1);

			nIndex = this.m_arIndexDst[this.m_lIndexSrc];

			let wR = this.GetValue(nIndex-1, ParamType.ptFormula, true) + "";
			let hR = this.GetValue(nIndex, ParamType.ptFormula, true) + "";
			let stAng = this.GetValue(nIndex-5, ParamType.ptFormula, true) + "";
			let swAng = this.GetValue(nIndex-4, ParamType.ptFormula, true) + "";
			oOOXMLPath.arcTo(wR, hR, stAng, swAng);
			this.m_lIndexSrc++;
			this.convertVal(pPoint.x, pPointType.x, false);
			this.convertVal(pPoint.y, pPointType.y, false);
			this.m_arIndexDst.push(this.m_lIndexDst-1);
			return;
		};

		CVmlGeometryData.prototype.ConvertQuadrY = function(pPoint, pPointType, oOOXMLPath)
		{
			let nIndex = this.m_arIndexDst[this.m_lIndexSrc];

			this.m_lIndexSrc++;

			this.convertSum(pPoint.x, pPointType.x, 0, ParamType.ptValue, nIndex-1, ParamType.ptFormula, false, true, true);
			this.convertSum(pPoint.y, pPointType.y, 0, ParamType.ptValue, nIndex, ParamType.ptFormula, false, true, true);
			this.convertIf(this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, -1, ParamType.ptValue, true, true, true);
			this.convertIf(this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, -1, ParamType.ptValue, true, true, true);
			this.convertProd(this.m_lIndexDst-2, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
			this.convertIf(this.m_lIndexDst-5, ParamType.ptFormula, 10800000, ParamType.ptValue, 0, ParamType.ptValue, true, true, true);	//stAng
			this.convertIf(this.m_lIndexDst-2, ParamType.ptFormula, -5400000, ParamType.ptValue, 5400000, ParamType.ptValue, true, true, true);	//swAng
			this.convertProd(this.m_lIndexDst-7, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);
			this.convertProd(this.m_lIndexDst-7, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);
			this.convertIf(this.m_lIndexDst-9, ParamType.ptFormula, this.m_lIndexDst-9, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, true, true, true);//wR
			this.convertIf(this.m_lIndexDst-9, ParamType.ptFormula, this.m_lIndexDst-9, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, true, true, true);//hR
			this.m_arIndexDst.push(this.m_lIndexDst-1);

			nIndex = this.m_arIndexDst[this.m_lIndexSrc];

			let wR = this.GetValue(nIndex-1, ParamType.ptFormula, true);
			let hR = this.GetValue(nIndex, ParamType.ptFormula, true);
			let stAng = this.GetValue(nIndex-5, ParamType.ptFormula, true);
			let swAng = this.GetValue(nIndex-4, ParamType.ptFormula, true);
			oOOXMLPath.arcTo(wR, hR, stAng, swAng);
			this.m_lIndexSrc++;
			this.convertVal(pPoint.x, pPointType.x, false);
			this.convertVal(pPoint.y, pPointType.y, false);
			this.m_arIndexDst.push(this.m_lIndexDst-1);
		};
		//-------------------------------------

		CVmlGeometryData.prototype.ConvertSlice_MoveTo = function(oSlice, oOOXMLPath)
		{
			for (let j = 0; j < oSlice.m_arPoints.length; ++j)
			{
				this.pCurPoint		= oSlice.m_arPoints[j];
				this.pCurPointType	= oSlice.m_arPointsType[j];

				this.m_lIndexSrc++;
				this.convertVal(this.pCurPoint.x, this.pCurPointType.x, false);
				this.convertVal(this.pCurPoint.y, this.pCurPointType.y, false);

				this.m_arIndexDst.push(this.m_lIndexDst-1);
				let x = this.GetValue(this.m_lIndexDst-2, ParamType.ptFormula, true) + "";
				let y = this.GetValue(this.m_lIndexDst-1, ParamType.ptFormula, true) + "";
				oOOXMLPath.moveTo(x, y);
			}
		};

		CVmlGeometryData.prototype.ConvertSlice_RMoveTo = function( oSlice, oOOXMLPath)
		{
			for (let j = 0; j < oSlice.m_arPoints.length; j++)
			{
				this.pCurPoint		= oSlice.m_arPoints[j];
				this.pCurPointType	= oSlice.m_arPointsType[j];
				this.m_lIndexSrc ++;
				this.convertSum(this.m_lIndexDst-2, ParamType.ptFormula, this.pCurPoint.x, this.pCurPointType.x, 0, ParamType.ptValue, true, false, true);
				this.convertSum(this.m_lIndexDst-2, ParamType.ptFormula, this.pCurPoint.y, this.pCurPointType.y, 0, ParamType.ptValue, true, false, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);
				let x = this.GetValue(this.m_lIndexDst-2, ParamType.ptFormula, true) + "";
				let y = this.GetValue(this.m_lIndexDst-1, ParamType.ptFormula, true) + "";
				oOOXMLPath.moveTo(x, y);
			}
		};

		CVmlGeometryData.prototype.ConvertSlice_LineTo = function( oSlice, oOOXMLPath)
		{
			for (let j = 0; j < oSlice.m_arPoints.length; j++)
			{
				this.pCurPoint = oSlice.m_arPoints[j];
				this.pCurPointType = oSlice.m_arPointsType[j];

				this.m_lIndexSrc++;
				this.convertVal(this.pCurPoint.x, this.pCurPointType.x, false);
				this.convertVal(this.pCurPoint.y, this.pCurPointType.y, false);

				this.m_arIndexDst.push(this.m_lIndexDst-1);

				let x = this.GetValue(this.m_lIndexDst-2, ParamType.ptFormula, true) + "";
				let y = this.GetValue(this.m_lIndexDst-1, ParamType.ptFormula, true) + "";
				oOOXMLPath.lnTo(x, y);
			}
		};

		CVmlGeometryData.prototype.ConvertSlice_RLineTo = function( oSlice, oOOXMLPath)
		{
			for (let j = 0; j < oSlice.m_arPoints.length; j++)
			{
				this.pCurPoint = oSlice.m_arPoints[j];
				this.pCurPointType = oSlice.m_arPointsType[j];
				this.m_lIndexSrc++;
				this.convertSum(this.m_lIndexDst-2, ParamType.ptFormula, this.pCurPoint.x, this.pCurPointType.x, 0, ParamType.ptValue, true, false, true);
				this.convertSum(this.m_lIndexDst-2, ParamType.ptFormula, this.pCurPoint.y, this.pCurPointType.y, 0, ParamType.ptValue, true, false, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);
				let x = this.GetValue(this.m_lIndexDst-2, ParamType.ptFormula, true) + "";
				let y = this.GetValue(this.m_lIndexDst-1, ParamType.ptFormula, true) + "";
				oOOXMLPath.lnTo(x, y);
			}
		};

		CVmlGeometryData.prototype.ConvertSlice_ArcTo = function(oSlice, oOOXMLPath)
		{
			let nIndex = 0;
			let nIndex1 = 0;
			let nIndex2 = 0;

			for (let j = 0; j < oSlice.m_arPoints.length; j+=4)
			{
				this.pCurPoint		= oSlice.m_arPoints[j];
				this.pCurPointType	= oSlice.m_arPointsType[j];
				this.pCurPoint1		= oSlice.m_arPoints[j+1];
				this.pCurPointType1	= oSlice.m_arPointsType[j+1];

				this.m_lIndexSrc++;
				this.convertSum(this.pCurPoint1.x, this.pCurPointType1.x, 0, ParamType.ptValue, this.pCurPoint.x, this.pCurPointType.x, false, true, false);
				this.convertSum(this.pCurPoint1.y, this.pCurPointType1.y, 0, ParamType.ptValue, this.pCurPoint.y, this.pCurPointType.y, false, true, false);
				this.m_arIndexDst.push(this.m_lIndexDst-1);
				nIndex = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc++;
				this.convertProd(nIndex-1, ParamType.ptFormula, 1, ParamType.ptValue, 2, ParamType.ptValue, true, true, true); //a=wR
				this.convertProd(nIndex, ParamType.ptFormula, 1, ParamType.ptValue, 2, ParamType.ptValue, true, true, true); //b=hR

				this.convertSum(this.pCurPoint.x, this.pCurPointType.x, this.m_lIndexDst-2, ParamType.ptFormula, 0, ParamType.ptValue, false, true, true);
				this.convertSum(this.pCurPoint.y, this.pCurPointType.y, this.m_lIndexDst-2, ParamType.ptFormula, 0, ParamType.ptValue, false, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				this.pCurPoint = oSlice.m_arPoints[j+2];
				this.pCurPointType = oSlice.m_arPointsType[j+2];
				nIndex = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc++;
				this.convertSum(this.pCurPoint.x, this.pCurPointType.x, 0, ParamType.ptValue, nIndex-1, ParamType.ptFormula, false, true, true);
				this.convertSum(this.pCurPoint.y, this.pCurPointType.y, 0, ParamType.ptValue, nIndex, ParamType.ptFormula, false, true, true);
				this.convertIf(this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, -1, ParamType.ptValue, true, true, true);
				this.convertIf(this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, -1, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);
				nIndex = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc++;
				this.convertProd(nIndex-2, ParamType.ptFormula, 1, ParamType.ptValue, nIndex-3, ParamType.ptFormula, true, true, true);
				this.convertAt2(1, ParamType.ptValue, this.m_lIndexDst-1, ParamType.ptFormula, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);
				nIndex = this.m_arIndexDst[this.m_lIndexSrc];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-2];

				this.m_lIndexSrc++;
				this.convertCos(1, ParamType.ptValue, nIndex, ParamType.ptFormula, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex1-2, ParamType.ptFormula, nIndex1-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);//(b*cos(u))^2

				this.convertSin(1, ParamType.ptValue, nIndex, ParamType.ptFormula, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex1-3, ParamType.ptFormula, nIndex1-3, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);//(a*sin(u))^2

				this.convertSum(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-5, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);
				this.convertSqrt(this.m_lIndexDst-1, ParamType.ptFormula, true);
				this.convertProd(nIndex1-3, ParamType.ptFormula, nIndex1-2, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);//r
				this.m_arIndexDst.push(this.m_lIndexDst-1);
				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];

				this.m_lIndexSrc++;
				this.convertIf(nIndex-3, ParamType.ptFormula, 0, ParamType.ptValue, 10800000, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);
				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-3];
				nIndex2 = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc ++;
				this.convertProd(nIndex1-1, ParamType.ptFormula, nIndex1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);
				this.convertIf(nIndex, ParamType.ptFormula, nIndex, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);

				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-3, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);

				this.convertSum(this.m_lIndexDst-1, ParamType.ptFormula, nIndex2, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);
				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-4];
				nIndex2 = this.m_arIndexDst[this.m_lIndexSrc-5];

				this.m_lIndexSrc++;
				this.convertProd(nIndex, ParamType.ptFormula, nIndex-10, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex1-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertSum(nIndex2-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true); //x

				this.convertProd(nIndex-6, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);
				this.convertIf(nIndex-6, ParamType.ptFormula, nIndex-6, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);

				this.convertProd(nIndex, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);//r*sin
				this.convertProd(nIndex1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertSum(nIndex2, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);//y
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				this.pCurPoint = oSlice.m_arPoints[j+3];
				this.pCurPointType = oSlice.m_arPointsType[j+3];

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-6];

				this.m_lIndexSrc++;
				this.convertSum(this.pCurPoint.x, this.pCurPointType.x, 0, ParamType.ptValue, nIndex-1, ParamType.ptFormula, false, true, true);
				this.convertSum(this.pCurPoint.y, this.pCurPointType.y, 0, ParamType.ptValue, nIndex, ParamType.ptFormula, false, true, true);
				this.convertIf(this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, -1, ParamType.ptValue, true, true, true);
				this.convertIf(this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, -1, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc++;
				this.convertProd(nIndex-2, ParamType.ptFormula, 1, ParamType.ptValue, nIndex-3, ParamType.ptFormula, true, true, true);
				this.convertAt2(1, ParamType.ptValue, this.m_lIndexDst-1, ParamType.ptFormula, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-8];

				this.m_lIndexSrc++;
				this.convertCos(1, ParamType.ptValue, nIndex, ParamType.ptFormula, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex1-2, ParamType.ptFormula, nIndex1-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);

				this.convertSin(1, ParamType.ptValue, nIndex, ParamType.ptFormula, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex1-3, ParamType.ptFormula, nIndex1-3, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);

				this.convertSum(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-5, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);
				this.convertSqrt(this.m_lIndexDst-1, ParamType.ptFormula, true);
				this.convertProd(nIndex1-3, ParamType.ptFormula, nIndex1-2, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];

				this.m_lIndexSrc++;
				this.convertIf(nIndex-3, ParamType.ptFormula, 0, ParamType.ptValue, 10800000, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-3];
				nIndex2 = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc++;
				this.convertProd(nIndex1-1, ParamType.ptFormula, nIndex1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);
				this.convertIf(nIndex, ParamType.ptFormula, nIndex, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);

				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-3, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);

				this.convertSum(this.m_lIndexDst-1, ParamType.ptFormula, nIndex2, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);
				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-4];
				nIndex2 = this.m_arIndexDst[this.m_lIndexSrc-11];

				this.m_lIndexSrc++;
				this.convertProd(nIndex, ParamType.ptFormula, nIndex-10, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);//r*cos(a)
				this.convertProd(nIndex1-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertSum(nIndex2-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);//x

				this.convertProd(nIndex-6, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);
				this.convertIf(nIndex-6, ParamType.ptFormula, nIndex-6, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);

				this.convertProd(nIndex, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);//r*sin(a)
				this.convertProd(nIndex1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertSum(nIndex2, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);//y
				this.m_arIndexDst.push(this.m_lIndexDst-1);
				nIndex = this.m_arIndexDst[this.m_lIndexSrc-1];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-7];

				this.m_lIndexSrc++;
				this.convertSum(nIndex1, ParamType.ptFormula, 0, ParamType.ptValue, nIndex, ParamType.ptFormula, true, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);

				this.convertSum(nIndex1, ParamType.ptFormula, 21600000, ParamType.ptValue, nIndex, ParamType.ptFormula, true, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true); // -1*((1)+360-(2))

				this.convertIf(this.m_lIndexDst-4, ParamType.ptFormula, this.m_lIndexDst-3, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);

				this.m_arIndexDst.push(this.m_lIndexDst-1);
				nIndex = this.m_arIndexDst[this.m_lIndexSrc-14];
				this.m_lIndexSrc++;
				this.convertProd(nIndex-1, ParamType.ptFormula, 1, ParamType.ptValue, 2, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex, ParamType.ptFormula, 1, ParamType.ptValue, 2, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				//---------------------------------------------------------
				nIndex = this.m_arIndexDst[this.m_lIndexSrc-8];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-9];
				nIndex2 = this.m_arIndexDst[this.m_lIndexSrc-1];

				if (oSlice.m_eRuler === RulesType.rtArc && j === 0)
				{
					let x = this.GetValue(nIndex-5, ParamType.ptFormula, true) + "";
					let y = this.GetValue(nIndex, ParamType.ptFormula, true) + "";
					oOOXMLPath.moveTo(x, y);
					let wR = this.GetValue(this.m_lIndexDst-2, ParamType.ptFormula, true) + "";
					let hR = this.GetValue(this.m_lIndexDst-1, ParamType.ptFormula, true) + "";
					let stAng = this.GetValue(nIndex1, ParamType.ptFormula, true) + "";
					let swAng = this.GetValue(nIndex2, ParamType.ptFormula, true) + "";
					oOOXMLPath.arcTo(wR, hR, stAng, swAng);
				}
				else
				{
					let x = this.GetValue(nIndex-5, ParamType.ptFormula, true) + "";
					let y = this.GetValue(nIndex, ParamType.ptFormula, true) + "";
					let wR = this.GetValue(this.m_lIndexDst-2, ParamType.ptFormula, true) + "";
					let hR = this.GetValue(this.m_lIndexDst-1, ParamType.ptFormula, true) + "";
					let stAng = this.GetValue(nIndex1, ParamType.ptFormula, true) + "";
					let swAng = this.GetValue(nIndex2, ParamType.ptFormula, true) + "";
					oOOXMLPath.lnTo(x, y);
					oOOXMLPath.arcTo(wR, hR, stAng, swAng);
				}
				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];
				this.convertVal(nIndex-5, ParamType.ptFormula, true);
				this.convertVal(nIndex, ParamType.ptFormula, true);
			}
		};

		CVmlGeometryData.prototype.ConvertSlice_ClockwiseTo = function( oSlice, oOOXMLPath)
		{
			let nIndex = 0;
			let nIndex1 = 0;
			let nIndex2 = 0;

			for (let j = 0; j < oSlice.m_arPoints.length; j += 4)
			{
				this.pCurPoint		= oSlice.m_arPoints[j];
				this.pCurPointType	= oSlice.m_arPointsType[j];
				this.pCurPoint1		= oSlice.m_arPoints[j+1];
				this.pCurPointType1	= oSlice.m_arPointsType[j+1];

				this.m_lIndexSrc++;
				this.convertSum(this.pCurPoint1.x, this.pCurPointType1.x, 0, ParamType.ptValue, this.pCurPoint.x, this.pCurPointType.x, false, true, false);
				this.convertSum(this.pCurPoint1.y, this.pCurPointType1.y, 0, ParamType.ptValue, this.pCurPoint.y, this.pCurPointType.y, false, true, false);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc++;
				this.convertProd(nIndex-1, ParamType.ptFormula, 1, ParamType.ptValue, 2, ParamType.ptValue, true, true, true); //a=wR
				this.convertProd(nIndex, ParamType.ptFormula, 1, ParamType.ptValue, 2, ParamType.ptValue, true, true, true); //b=hR

				this.convertSum(this.pCurPoint.x, this.pCurPointType.x, this.m_lIndexDst-2, ParamType.ptFormula, 0, ParamType.ptValue, false, true, true);
				this.convertSum(this.pCurPoint.y, this.pCurPointType.y, this.m_lIndexDst-2, ParamType.ptFormula, 0, ParamType.ptValue, false, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);


				this.pCurPoint = oSlice.m_arPoints[j+2];
				this.pCurPointType = oSlice.m_arPointsType[j+2];
				nIndex = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc++;
				this.convertSum(this.pCurPoint.x, this.pCurPointType.x, 0, ParamType.ptValue, nIndex-1, ParamType.ptFormula, false, true, true);
				this.convertSum(this.pCurPoint.y, this.pCurPointType.y, 0, ParamType.ptValue, nIndex, ParamType.ptFormula, false, true, true);
				this.convertIf(this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, -1, ParamType.ptValue, true, true, true);
				this.convertIf(this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, -1, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc++;
				this.convertProd(nIndex-2, ParamType.ptFormula, 1, ParamType.ptValue, nIndex-3, ParamType.ptFormula, true, true, true);
				this.convertAt2(1, ParamType.ptValue, this.m_lIndexDst-1, ParamType.ptFormula, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-2];

				this.m_lIndexSrc++;
				this.convertCos(1, ParamType.ptValue, nIndex, ParamType.ptFormula, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex1-2, ParamType.ptFormula, nIndex1-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);//(b*cos(u))^2

				this.convertSin(1, ParamType.ptValue, nIndex, ParamType.ptFormula, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex1-3, ParamType.ptFormula, nIndex1-3, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);//(a*sin(u))^2

				this.convertSum(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-5, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);
				this.convertSqrt(this.m_lIndexDst-1, ParamType.ptFormula, true);
				this.convertProd(nIndex1-3, ParamType.ptFormula, nIndex1-2, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);//r
				this.m_arIndexDst.push(this.m_lIndexDst-1);
				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];

				this.m_lIndexSrc++;
				this.convertIf(nIndex-3, ParamType.ptFormula, 0, ParamType.ptValue, 10800000, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-3];
				nIndex2 = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc++;
				this.convertProd(nIndex1-1, ParamType.ptFormula, nIndex1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);
				this.convertIf(nIndex, ParamType.ptFormula, nIndex, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);

				this.convertProd( this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-3, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);

				this.convertSum( this.m_lIndexDst-1, ParamType.ptFormula, nIndex2, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-4];
				nIndex2 = this.m_arIndexDst[this.m_lIndexSrc-5];

				this.m_lIndexSrc++;
				this.convertProd(nIndex, ParamType.ptFormula, nIndex-10, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);//r*cos(a)
				this.convertProd(nIndex1-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertSum(nIndex2-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true); //x

				this.convertProd(nIndex-6, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);
				this.convertIf(nIndex-6, ParamType.ptFormula, nIndex-6, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);

				this.convertProd(nIndex, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);//r*sin
				this.convertProd(nIndex1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertSum(nIndex2, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);//y
				this.m_arIndexDst.push(this.m_lIndexDst-1);


				this.pCurPoint = oSlice.m_arPoints[j+3];
				this.pCurPointType = oSlice.m_arPointsType[j+3];

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-6];

				this.m_lIndexSrc++;
				this.convertSum(this.pCurPoint.x, this.pCurPointType.x, 0, ParamType.ptValue, nIndex-1, ParamType.ptFormula, false, true, true);
				this.convertSum(this.pCurPoint.y, this.pCurPointType.y, 0, ParamType.ptValue, nIndex, ParamType.ptFormula, false, true, true);
				this.convertIf(this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, -1, ParamType.ptValue, true, true, true);
				this.convertIf(this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, -1, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc++;
				this.convertProd(nIndex-2, ParamType.ptFormula, 1, ParamType.ptValue, nIndex-3, ParamType.ptFormula, true, true, true);
				this.convertAt2(1, ParamType.ptValue, this.m_lIndexDst-1, ParamType.ptFormula, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-8];

				this.m_lIndexSrc++;
				this.convertCos(1, ParamType.ptValue, nIndex, ParamType.ptFormula, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex1-2, ParamType.ptFormula, nIndex1-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);

				this.convertSin(1, ParamType.ptValue, nIndex, ParamType.ptFormula, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex1-3, ParamType.ptFormula, nIndex1-3, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);

				this.convertSum(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-5, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);
				this.convertSqrt(this.m_lIndexDst-1, ParamType.ptFormula, true);
				this.convertProd(nIndex1-3, ParamType.ptFormula, nIndex1-2, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];

				this.m_lIndexSrc++;
				this.convertIf(nIndex-3, ParamType.ptFormula, 0, ParamType.ptValue, 10800000, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-3];
				nIndex2 = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc ++;
				this.convertProd(nIndex1-1, ParamType.ptFormula, nIndex1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);
				this.convertIf(nIndex, ParamType.ptFormula, nIndex, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);

				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-3, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);

				this.convertSum(this.m_lIndexDst-1, ParamType.ptFormula, nIndex2, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-4];
				nIndex2 = this.m_arIndexDst[this.m_lIndexSrc-11];

				this.m_lIndexSrc++;
				this.convertProd(nIndex, ParamType.ptFormula, nIndex-10, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);//r*cos(a)
				this.convertProd(nIndex1-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertSum(nIndex2-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);//x

				this.convertProd(nIndex-6, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);
				this.convertIf(nIndex-6, ParamType.ptFormula, nIndex-6, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);

				this.convertProd(nIndex, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);//r*sin(a)
				this.convertProd(nIndex1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertSum(nIndex2, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);//y
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-1];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-7];

				this.m_lIndexSrc++;
				this.convertSum(nIndex, ParamType.ptFormula, 0, ParamType.ptValue, nIndex1, ParamType.ptFormula, true, true, true);
				this.convertSum(21600000, ParamType.ptValue, this.m_lIndexDst-1, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);
				this.convertIf(this.m_lIndexDst-2, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);//swAng

				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-14];
				this.m_lIndexSrc++;
				this.convertProd(nIndex-1, ParamType.ptFormula, 1, ParamType.ptValue, 2, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex, ParamType.ptFormula, 1, ParamType.ptValue, 2, ParamType.ptValue, true, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				//---------------------------------------------------------
				nIndex = this.m_arIndexDst[this.m_lIndexSrc-8];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-9];
				nIndex2 = this.m_arIndexDst[this.m_lIndexSrc-1];

				if (oSlice.m_eRuler === RulesType.rtClockwiseArc && j === 0)
				{
					let x = this.GetValue(nIndex-5, ParamType.ptFormula, true) + "";
					let y = this.GetValue(nIndex, ParamType.ptFormula, true) + "";
					oOOXMLPath.moveTo(x, y);
					let wR = this.GetValue(this.m_lIndexDst-2, ParamType.ptFormula, true);
					let hR = this.GetValue(this.m_lIndexDst-1, ParamType.ptFormula, true);
					let stAng = this.GetValue(nIndex1, ParamType.ptFormula, true);
					let swAng = this.GetValue(nIndex2, ParamType.ptFormula, true);
					oOOXMLPath.arcTo(wR, hR, stAng, swAng);
				}
				else
				{
					let x = this.GetValue(nIndex-5, ParamType.ptFormula, true) + "";
					let y = this.GetValue(nIndex, ParamType.ptFormula, true) + "";
					oOOXMLPath.lnTo(x, y);
					let wR = this.GetValue(this.m_lIndexDst-2, ParamType.ptFormula, true) + "";
					let hR = this.GetValue(this.m_lIndexDst-1, ParamType.ptFormula, true) + "";
					let stAng = this.GetValue(nIndex1, ParamType.ptFormula, true) + "";
					let swAng = this.GetValue(nIndex2, ParamType.ptFormula, true) + "";
					oOOXMLPath.arcTo(wR, hR, stAng, swAng);
				}

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-2];
				this.convertVal(nIndex-5, ParamType.ptFormula, true);
				this.convertVal(nIndex, ParamType.ptFormula, true);
			}
		};

		CVmlGeometryData.prototype.ConvertSlice_QuadrBesier = function( oSlice, oOOXMLPath)
		{
			for (let j = 0; j < oSlice.m_arPoints.length; j += 2)
			{
				let l = (oSlice.m_arPoints.length - j - 3);
				if (l >= 0)
				{
					let aPoints = []
					for (let  k = 0; k < 2; ++k)
					{
						this.pCurPoint = oSlice.m_arPoints[j+k];
						this.pCurPointType = oSlice.m_arPointsType[j+k];
						let oPt = {x: "0", y: "0"};
						oPt.x = this.GetValue(this.pCurPoint.x, this.pCurPointType.x, false) + "";
						oPt.y = this.GetValue(this.pCurPoint.y, this.pCurPointType.y, false) + "";
						aPoints.push(oPt);
					}
					oOOXMLPath.quadBezTo(aPoints[0].x, aPoints[0].y, aPoints[1].x, aPoints[1].y);
				}
				else
				{
					for (let k = 0; k < oSlice.m_arPoints.length - j; ++k)
					{
						this.pCurPoint = oSlice.m_arPoints[j+k];
						this.pCurPointType = oSlice.m_arPointsType[j+k];
						let x = this.GetValue(this.pCurPoint.x, this.pCurPointType.x, false) + "";
						let y = this.GetValue(this.pCurPoint.y, this.pCurPointType.y, false) + "";
						oOOXMLPath.lnTo(x, y);
					}
				}

				this.m_lIndexSrc++;
				this.convertVal(this.pCurPoint.x, this.pCurPointType.x, false);
				this.convertVal(this.pCurPoint.y, this.pCurPointType.y, false);
				this.m_arIndexDst.push(this.m_lIndexDst-1);
			}
		};

		CVmlGeometryData.prototype.ConvertSlice_CurveTo = function(oSlice, oOOXMLPath)
		{
			let nIndex = 0;
			for (let j = 0; j < oSlice.m_arPoints.length; j += 3)
			{
				let l = (oSlice.m_arPoints.length - j - 3);
				if ( l >= 0 )
				{
					let aPoints = [];
					for (let k = 0; k < 3; ++k)
					{
						this.pCurPoint = oSlice.m_arPoints[j+k];
						this.pCurPointType = oSlice.m_arPointsType[j+k];

						let oPt = {};
						oPt.x = this.GetValue(this.pCurPoint.x, this.pCurPointType.x, false) + "";
						oPt.y = this.GetValue(this.pCurPoint.y, this.pCurPointType.y, false) + "";
						aPoints.push(oPt);
					}
					oOOXMLPath.cubicBezTo(aPoints[0].x, aPoints[0].y, aPoints[1].x, aPoints[1].y, aPoints[2].x, aPoints[2].y)
				}
				else
				{
					for (let k = 0; k < oSlice.m_arPoints.length - j; ++k)
					{
						this.pCurPoint = oSlice.m_arPoints[j+k];
						this.pCurPointType = oSlice.m_arPointsType[j+k];
						let x = this.GetValue(this.pCurPoint.x, this.pCurPointType.x, false) + "";
						let y = this.GetValue(this.pCurPoint.y, this.pCurPointType.y, false) + "";
						oOOXMLPath.lnTo(x, y);
					}
				}

				this.m_lIndexSrc++;
				this.convertVal(this.pCurPoint.x, this.pCurPointType.x, false);
				this.convertVal(this.pCurPoint.y, this.pCurPointType.y, false);
				this.m_arIndexDst.push(this.m_lIndexDst-1);
			}
		};

		CVmlGeometryData.prototype.ConvertSlice_RCurveTo = function( oSlice, oOOXMLPath)
		{
			let nIndex = 0;
			for (let j = 0; j < oSlice.m_arPoints.length; j += 3)
			{
				nIndex = this.m_arIndexDst[this.m_lIndexSrc];

				let l = (oSlice.m_arPoints.length - j - 3);
				if (l >= 0)
				{
					let aPoints = [];
					for (let k = 0; k < 3; ++k)
					{
						this.pCurPoint = oSlice.m_arPoints[j+k];
						this.pCurPointType = oSlice.m_arPointsType[j+k];

						this.m_lIndexSrc++;
						this.convertSum(nIndex-1, ParamType.ptFormula, this.pCurPoint.x, this.pCurPointType.x, 0, ParamType.ptValue, true, false, true);
						this.convertSum(nIndex, ParamType.ptFormula, this.pCurPoint.y, this.pCurPointType.y, 0, ParamType.ptValue, true, false, true);
						this.m_arIndexDst.push(this.m_lIndexDst-1);

						let oPt = {};
						oPt.x = this.GetValue(this.m_lIndexDst-2, ParamType.ptFormula, true) + "";
						oPt.y = this.GetValue(this.m_lIndexDst-1, ParamType.ptFormula, true) + "";
						aPoints.push(oPt);
					}
					oOOXMLPath.cubicBezTo(aPoints[0].x, aPoints[0].y, aPoints[1].x, aPoints[1].y, aPoints[2].x, aPoints[2].y)
				}
				else
				{
					for (let k=0; k < oSlice.m_arPoints.length - j; ++k)
					{
						this.pCurPoint		= oSlice.m_arPoints[j+k];
						this.pCurPointType	= oSlice.m_arPointsType[j+k];

						this.m_lIndexSrc++;
						this.convertSum(nIndex-1, ParamType.ptFormula, this.pCurPoint.x, this.pCurPointType.x, 0, ParamType.ptValue, true, false, true);
						this.convertSum(nIndex, ParamType.ptFormula, this.pCurPoint.y, this.pCurPointType.y, 0, ParamType.ptValue, true, false, true);
						this.m_arIndexDst.push(this.m_lIndexDst-1);
						let x = this.GetValue(this.m_lIndexDst-2, ParamType.ptFormula, true) + "";
						let y = this.GetValue(this.m_lIndexDst-1, ParamType.ptFormula, true) + "";
						oOOXMLPath.lnTo(x, y);
					}
				}

				nIndex = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc++;
				this.convertVal(nIndex-1, ParamType.ptFormula, true);
				this.convertVal(nIndex, ParamType.ptFormula, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);
			}
		};

		CVmlGeometryData.prototype.ConvertSlice_AngleEllipse = function( oSlice, oOOXMLPath)
		{
			let nIndex = 0;
			let nIndex1 = 0;
			let nIndex2 = 0;

			for (let j = 0; j < oSlice.m_arPoints.length; j += 3)
			{
				this.pCurPoint		= oSlice.m_arPoints[j+1];
				this.pCurPointType	= oSlice.m_arPointsType[j+1];
				this.pCurPoint1		= oSlice.m_arPoints[j+2];
				this.pCurPointType1	= oSlice.m_arPointsType[j+2];

				this.m_lIndexSrc++;
				this.convertProd(this.pCurPoint1.x, this.pCurPointType1.x, pow3_16, ParamType.ptValue, m_oParam.m_lParam, m_oParam.m_eType, false, true, true);
				this.convertProd(this.pCurPoint1.y, this.pCurPointType1.y, pow3_16, ParamType.ptValue, m_oParam.m_lParam, m_oParam.m_eType, false, true, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc++;
				this.convertVal(this.pCurPoint.x, this.pCurPointType.x, false);//wr=a
				this.convertVal(this.pCurPoint.y, this.pCurPointType.y, false);//hr=b

				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-1];

				this.m_lIndexSrc++;
				this.convertProd(nIndex-1, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true); //stAng
				this.convertSum(nIndex-1, ParamType.ptFormula, nIndex, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex, ParamType.ptFormula, -1, ParamType.ptValue, 1, ParamType.ptValue, true, true, true);//swAng
				this.m_arIndexDst.push(this.m_lIndexDst-1);

				nIndex = this.m_arIndexDst[this.m_lIndexSrc]; //stang
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-1]; //wr hr

				this.m_lIndexSrc++;
				this.convertCos(1, ParamType.ptValue, nIndex-2, ParamType.ptFormula, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex1, ParamType.ptFormula, nIndex1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);//(b*cos(u))^2

				this.convertSin(1, ParamType.ptValue, nIndex-2, ParamType.ptFormula, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(nIndex1-1, ParamType.ptFormula, nIndex1-1, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);
				this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);//(a*sin(u))^2

				this.convertSum(this.m_lIndexDst-1, ParamType.ptFormula, this.m_lIndexDst-5, ParamType.ptFormula, 0, ParamType.ptValue, true, true, true);
				this.convertSqrt(this.m_lIndexDst-1, ParamType.ptFormula, true);
				this.convertProd(nIndex1, ParamType.ptFormula, nIndex1-1, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);//r
				this.m_arIndexDst.push(this.m_lIndexDst-1);


				this.pCurPoint1 = oSlice.m_arPoints[j];
				this.pCurPointType1 = oSlice.m_arPointsType[j];

				nIndex = this.m_arIndexDst[this.m_lIndexSrc-1]; //stang
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc]; //r

				this.m_lIndexSrc++;

				this.convertProd(nIndex1, ParamType.ptFormula, nIndex1-10, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);  //r*cos
				this.convertSum(this.pCurPoint1.x, this.pCurPointType1.x, this.m_lIndexDst-1, ParamType.ptFormula, 0, ParamType.ptValue, false, true, true);//x

				this.convertProd(nIndex1, ParamType.ptFormula, nIndex1-6, ParamType.ptFormula, 1, ParamType.ptValue, true, true, true);// r*sin
				this.convertSum(this.pCurPoint1.y, this.pCurPointType1.y, this.m_lIndexDst-1, ParamType.ptFormula, 0, ParamType.ptValue, false, true, true);//y
				this.m_arIndexDst.push(this.m_lIndexDst-1);
				//---------------------

				nIndex = this.m_arIndexDst[this.m_lIndexSrc];
				nIndex1 = this.m_arIndexDst[this.m_lIndexSrc-2];
				nIndex2 = this.m_arIndexDst[this.m_lIndexSrc-3];


				let x = this.GetValue(nIndex-2, ParamType.ptFormula, true) + "";
				let y = this.GetValue(nIndex, ParamType.ptFormula, true) + "";
				let wR =this.GetValue(nIndex2-1, ParamType.ptFormula, true) + "";
				let hR = this.GetValue(nIndex2, ParamType.ptFormula, true) + "";
				let stAng = this.GetValue(nIndex1-2, ParamType.ptFormula, true) + "";
				let swAng = this.GetValue(nIndex1, ParamType.ptFormula, true) + "";
				if (j === 0)
				{
					oOOXMLPath.moveTo(x, y);
				}
				else
				{
					oOOXMLPath.lineTo(x, y);
				}
				oOOXMLPath.arcTo(wR, hR, swAng, swAng);


				nIndex = this.m_arIndexDst[this.m_lIndexSrc];

				this.m_lIndexSrc++;
				this.convertVal(nIndex-2, ParamType.ptFormula, true);
				this.convertVal(nIndex, ParamType.ptFormula, true);
				this.m_arIndexDst.push(this.m_lIndexDst-1);
			}
		};

		CVmlGeometryData.prototype.ConvertSlice_EllipticalQuadrX = function( oSlice, oOOXMLPath)
		{
			for (let j = 0; j < oSlice.m_arPoints.length; j += 2)
			{
				this.pCurPoint		= oSlice.m_arPoints[j];
				this.pCurPointType	= oSlice.m_arPointsType[j];

				this.ConvertQuadrX(this.pCurPoint, this.pCurPointType, oOOXMLPath);

				if (j + 1 < oSlice.m_arPoints.length)
				{
					this.pCurPoint1 = oSlice.m_arPoints[j+1];
					this.pCurPointType1 = oSlice.m_arPointsType[j+1];
					this.ConvertQuadrY(this.pCurPoint1, this.pCurPointType1, oOOXMLPath);
				}
			}
		};

		CVmlGeometryData.prototype.ConvertSlice_EllipticalQuadrY = function( oSlice, oOOXMLPath)
		{
			for (let j = 0; j < oSlice.m_arPoints.length; j += 2)
			{
				this.pCurPoint = oSlice.m_arPoints[j];
				this.pCurPointType = oSlice.m_arPointsType[j];
				this.ConvertQuadrY(this.pCurPoint, this.pCurPointType, oOOXMLPath);

				if (j + 1 < oSlice.m_arPoints.length)
				{
					this.pCurPoint1 = oSlice.m_arPoints[j+1];
					this.pCurPointType1 = oSlice.m_arPointsType[j+1];
					this.ConvertQuadrX(this.pCurPoint1, this.pCurPointType1, oOOXMLPath);
				}
			}
		};


		CVmlGeometryData.prototype.convertGuides = function() {

			this.convertVal(this.m_oParam.m_lCoef, ParamType.ptValue, false);
			for(let nGuide = 0; nGuide < this.guides.length; ++nGuide) {
				let oGuide = this.guides[nGuide];
				this.m_lIndexSrc++;
				switch (oGuide.m_eFormulaType) {
					case PPTFormulaType.ftVal: {
						this.convertVal(oGuide.m_lParam1, oGuide.m_eType1, false);
						break;
					}
					case PPTFormulaType.ftSum: {
						this.convertSum(oGuide.m_lParam1, oGuide.m_eType1, oGuide.m_lParam2, oGuide.m_eType2, oGuide.m_lParam3, oGuide.m_eType3, false, false, false);
						break;
					}
					case PPTFormulaType.ftProduct: {
						this.convertProd(oGuide.m_lParam1, oGuide.m_eType1, oGuide.m_lParam2, oGuide.m_eType2, oGuide.m_lParam3, oGuide.m_eType3, false, false, false);
						break;
					}
					case PPTFormulaType.ftMid: {
						this.convertSum(oGuide.m_lParam1, oGuide.m_eType1, oGuide.m_lParam2, oGuide.m_eType2, 0, ParamType.ptValue, false, false, false);
						this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, 2, ParamType.ptValue, true, false, false);
						break;
					}
					case PPTFormulaType.ftAbsolute: {
						this.convertAbs(oGuide.m_lParam1, oGuide.m_eType1, false);
						break;
					}
					case PPTFormulaType.ftMin: {
						this.convertMin(oGuide.m_lParam1, oGuide.m_eType1, oGuide.m_lParam2, oGuide.m_eType2, false, false);
						break;
					}
					case PPTFormulaType.ftMax: {
						this.convertMax(oGuide.m_lParam1, oGuide.m_eType1, oGuide.m_lParam2, oGuide.m_eType2, false, false);
						break;
					}
					case PPTFormulaType.ftIf: {
						this.convertIf(oGuide.m_lParam1, oGuide.m_eType1, oGuide.m_lParam2, oGuide.m_eType2, oGuide.m_lParam3, oGuide.m_eType3, false, false, false);
						break;
					}
					case PPTFormulaType.ftSqrt: {
						this.convertSqrt(oGuide.m_lParam1, oGuide.m_eType1, false);
						break;
					}
					case PPTFormulaType.ftMod: {
						this.convertMod(oGuide.m_lParam1, oGuide.m_eType1, oGuide.m_lParam2, oGuide.m_eType2, oGuide.m_lParam3, oGuide.m_eType3, false, false, false);
						break;
					}
					case PPTFormulaType.ftSin: {
						this.convertProd(oGuide.m_lParam2, oGuide.m_eType2, pow3_16, ParamType.ptValue, this.m_oParam.m_lParam, this.m_oParam.m_eType, false, false, true);
						this.convertSin(oGuide.m_lParam1, oGuide.m_eType1, this.m_lIndexDst-1, ParamType.ptFormula, false, true);
						break;
					}
					case PPTFormulaType.ftCos: {
						this.convertProd(oGuide.m_lParam2, oGuide.m_eType2, pow3_16, ParamType.ptValue, this.m_oParam.m_lParam, this.m_oParam.m_eType, false, false, true);
						this.convertCos(oGuide.m_lParam1, oGuide.m_eType1, this.m_lIndexDst-1, ParamType.ptFormula, false, true);
						break;
					}
					case PPTFormulaType.ftTan: {
						this.convertTan(oGuide.m_lParam1, oGuide.m_eType1, oGuide.m_lParam2, oGuide.m_eType2, false, false);
						break;
					}
					case PPTFormulaType.ftAtan2: {


						//ConvertAt2(pFormula.m_lParam1, pFormula.m_eType1, pFormula.m_lParam2, pFormula.m_eType2, false, false, m_oGuidsRes);
						//ConvertProd(m_lIndexDst-1, ptFormula, m_oParam.m_lParam, m_oParam.m_eType, pow3_16, ptValue, true, true, false, m_oGuidsRes);
						this.convertAt2(oGuide.m_lParam1, oGuide.m_eType1, oGuide.m_lParam2, oGuide.m_eType2, false, false, true);
						this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_oParam.m_lParam, this.m_oParam.m_eType, pow3_16, ParamType.ptValue, true, true, false);
						break;
					}
					case PPTFormulaType.ftSinatan2: {
						this.convertSat2(oGuide.m_lParam1, oGuide.m_eType1, oGuide.m_lParam2, oGuide.m_eType2, oGuide.m_lParam3, oGuide.m_eType3, false, false, false);
						this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_oParam.m_lParam, this.m_oParam.m_eType, pow3_16, ParamType.ptValue, true, true, false);
						break;
					}
					case PPTFormulaType.ftCosatan2: {
						this.convertCat2(oGuide.m_lParam1, oGuide.m_eType1, oGuide.m_lParam2, oGuide.m_eType2, oGuide.m_lParam3, oGuide.m_eType3, false, false, false);
						this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_oParam.m_lParam, this.m_oParam.m_eType, pow3_16, ParamType.ptValue, true, true, false);
						break;
					}
					case PPTFormulaType.ftSumangle: {
						this.convertProd(oGuide.m_lParam1, oGuide.m_eType1, pow3_16, ParamType.ptValue, this.m_oParam.m_lParam, this.m_oParam.m_eType, false, false, true);
						this.convertProd(pow3_16, ParamType.ptValue, oGuide.m_lParam2, oGuide.m_eType2, 1, ParamType.ptValue, false, false, false);
						this.convertProd(pow3_16, ParamType.ptValue, oGuide.m_lParam3, oGuide.m_eType3, 1, ParamType.ptValue, false, false, false);
						this.convertSum(this.m_lIndexDst-3, ParamType.ptFormula, this.m_lIndexDst-2, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, true, true, true);
						this.convertProd(this.m_lIndexDst-1, ParamType.ptFormula, this.m_oParam.m_lParam, this.m_oParam.m_eType, pow3_16, ParamType.ptValue, true, true, false);
						break;
					}
					case PPTFormulaType.ftEllipse: {
						this.convertProd(oGuide.m_lParam1, oGuide.m_eType1, oGuide.m_lParam1, oGuide.m_eType1, 1, ParamType.ptValue , false, false, false);
						this.convertProd(oGuide.m_lParam2, oGuide.m_eType2, oGuide.m_lParam2, oGuide.m_eType2, 1, ParamType.ptValue , false, false, false);
						this.convertProd(1, ParamType.ptValue, this.m_lIndexDst-2, ParamType.ptFormula, this.m_lIndexDst-1, ParamType.ptFormula, false, true, true);
						this.convertSum(0, ParamType.ptValue, 1, ParamType.ptValue, this.m_lIndexDst-1, ParamType.ptFormula, false, false, true);
						this.convertSqrt(this.m_lIndexDst-1, ParamType.ptFormula, true);
						this.convertProd(oGuide.m_lParam3, oGuide.m_eType3, this.m_lIndexDst-1, ParamType.ptFormula, 1, ParamType.ptValue, false, true, false);
						break;
					}
				}
				this.m_arIndexDst.push(this.m_lIndexDst - 1);
			}
		};
		CVmlGeometryData.prototype.getNextGdName = function() {
			return "gd" + (this.m_lIndexDst++);
		};
		CVmlGeometryData.prototype.getValue = function(lParam, eParamType, bExtShape) {
			switch (eParamType)
			{
				case ParamType.ptFormula: {
					if (bExtShape || lParam < this.m_arIndexDst.length) {
						if(bExtShape) {
							return "gd" + lParam;
						}
						else {
							return "gd" + this.m_arIndexDst[lParam];
						}
					}
					break;
				}
				case ParamType.ptAdjust: {
					if (lParam > this.m_lMaxAdjUse)
						this.m_lMaxAdjUse = lParam;
					return "adj" + lParam;
				}
				case ParamType.ptValue:
				{
					return "" + lParam;
				}
				return "";
			}
		};
		CVmlGeometryData.prototype.convert1ParamFmla = function(sOOXMLFmla, lParam1, eType1, bExtShape1) {
			let sParam1 = this.getValue(lParam1, eType1, bExtShape1);
			this.addGeomGuide(this.getNextGdName(), sOOXMLFmla, sParam1);
		};
		CVmlGeometryData.prototype.convert2ParamFmla = function(sOOXMLFmla, lParam1, eType1, lParam2, eType2, bExtShape1, bExtShape2) {
			let sParam1 = this.getValue(lParam1, eType1, bExtShape1);
			let sParam2 = this.getValue(lParam2, eType2, bExtShape2);
			this.addGeomGuide(this.getNextGdName(), sOOXMLFmla, sParam1, sParam2);
		};
		CVmlGeometryData.prototype.convert3ParamFmla = function(sOOXMLFmla, lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3) {
			let sParam1 = this.getValue(lParam1, eType1, bExtShape1);
			let sParam2 = this.getValue(lParam2, eType2, bExtShape2);
			let sParam3 = this.getValue(lParam3, eType3, bExtShape3);
			this.addGeomGuide(this.getNextGdName(), sOOXMLFmla, sParam1, sParam2, sParam3);
		};
		CVmlGeometryData.prototype.convertVal = function (lParam1, eType1, bExtShape1) {
			this.convert1ParamFmla("val", lParam1, eType1, bExtShape1);
		};
		CVmlGeometryData.prototype.convertSum = function(lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3) {
			this.convert3ParamFmla("+-", lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3);
		};
		CVmlGeometryData.prototype.convertProd = function(lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3) {
			this.convert3ParamFmla("*/", lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3);
		};
		CVmlGeometryData.prototype.convertAbs = function (lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3) {
			this.convert3ParamFmla("abs", lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3);
		};
		CVmlGeometryData.prototype.convertMin = function(lParam1, eType1, lParam2, eType2, bExtShape1, bExtShape2) {
			this.convert2ParamFmla("min", lParam1, eType1, lParam2, eType2, bExtShape1, bExtShape2);
		};
		CVmlGeometryData.prototype.convertMax = function(lParam1, eType1, lParam2, eType2, bExtShape1, bExtShape2) {
			this.convert2ParamFmla("min", lParam1, eType1, lParam2, eType2, bExtShape1, bExtShape2);
		};
		CVmlGeometryData.prototype.convertIf = function(lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3) {
			this.convert3ParamFmla("?:", lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3);
		};
		CVmlGeometryData.prototype.convertSqrt = function (lParam1, eType1, bExtShape1) {
			this.convert1ParamFmla("sqrt", lParam1, eType1, bExtShape1);
		};
		CVmlGeometryData.prototype.convertAt2 = function(lParam1, eType1, lParam2, eType2, bExtShape1, bExtShape2) {
			this.convert2ParamFmla("at2", lParam1, eType1, lParam2, eType2, bExtShape1, bExtShape2);
		};
		CVmlGeometryData.prototype.convertSin = function(lParam1, eType1, lParam2, eType2, bExtShape1, bExtShape2) {
			this.convert2ParamFmla("sin", lParam1, eType1, lParam2, eType2, bExtShape1, bExtShape2);
		};
		CVmlGeometryData.prototype.convertCos = function(lParam1, eType1, lParam2, eType2, bExtShape1, bExtShape2) {
			this.convert2ParamFmla("cos", lParam1, eType1, lParam2, eType2, bExtShape1, bExtShape2);
		};
		CVmlGeometryData.prototype.convertCat2 = function (lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3) {
			this.convert3ParamFmla("cat2", lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3);
		};
		CVmlGeometryData.prototype.convertSat2 = function (lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3) {
			this.convert3ParamFmla("sat2", lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3);
		};
		CVmlGeometryData.prototype.convertMod = function (lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3) {
			this.convert3ParamFmla("mod", lParam1, eType1, lParam2, eType2, lParam3, eType3, bExtShape1, bExtShape2, bExtShape3);
		};
		CVmlGeometryData.prototype.convertTan = function(lParam1, eType1, lParam2, eType2, bExtShape1, bExtShape2) {
			this.convert2ParamFmla("tan", lParam1, eType1, lParam2, eType2, bExtShape1, bExtShape2);
		};



		CVmlGeometryData.prototype.fillCAccentBorderCallout90Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3nfem,l21600,r,21600l,21600xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");

			this.addAdjustment(-1800);
			this.addAdjustment(24300);
			this.addAdjustment(-1800);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");
			let oHandle1 = new CVmlHandle();
			oHandle1.position = "#0,#1";
			this.addHandle(oHandle1);
			let oHandle2 = new CVmlHandle();
			oHandle2.position = "#2,#3";
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCAccentBorderCallout1Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3nfem@2,l@2,21600nfem,l21600,r,21600l,21600xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");

			this.addAdjustment(8280);
			this.addAdjustment(24300);
			this.addAdjustment(-1800);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCAccentBorderCallout2Type = function() {

			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3@4@5nfem@4,l@4,21600nfem,l21600,r,21600l,21600xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("val #4");
			this.addGuide("val #5");

			this.addAdjustment(-10080);
			this.addAdjustment(24300);
			this.addAdjustment(-3600);
			this.addAdjustment(4050);
			this.addAdjustment(-1800);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#4,#5");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCAccentBorderCallout3Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3@4@5@6@7nfem@6,l@6,21600nfem,l21600,r,21600l,21600xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("val #4");
			this.addGuide("val #5");
			this.addGuide("val #6");
			this.addGuide("val #7");

			this.addAdjustment(23400);
			this.addAdjustment(24400);
			this.addAdjustment(25200);
			this.addAdjustment(21600);
			this.addAdjustment(25200);
			this.addAdjustment(4050);
			this.addAdjustment(23400);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#4,#5");
			this.addHandle(oHandle3);

			let oHandle4 = new CVmlHandle();
			oHandle4.position = ("#6,#7");
			this.addHandle(oHandle4);
		};
		CVmlGeometryData.prototype.fillCAccentCallout90Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3nfem,l21600,r,21600l,21600xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");

			this.addAdjustment(-1800);
			this.addAdjustment(24300);
			this.addAdjustment(-1800);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCAccentCallout1Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3nfem@2,l@2,21600nfem,l21600,r,21600l,21600nsxe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");

			this.addAdjustment(-8280);
			this.addAdjustment(24300);
			this.addAdjustment(-1800);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCAccentCallout2Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3@4@5nfem@4,l@4,21600nfem,l21600,r,21600l,21600nsxe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("val #4");
			this.addGuide("val #5");

			this.addAdjustment(-10080);
			this.addAdjustment(24300);
			this.addAdjustment(-3600);
			this.addAdjustment(4050);
			this.addAdjustment(-1800);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#4,#5");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCAccentCallout3Type = function() {

			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3@4@5@6@7nfem@6,l@6,21600nfem,l21600,r,21600l,21600nsxe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("val #4");
			this.addGuide("val #5");
			this.addGuide("val #6");
			this.addGuide("val #7");

			this.addAdjustment(23400);
			this.addAdjustment(24400);
			this.addAdjustment(25200);
			this.addAdjustment(21600);
			this.addAdjustment(25200);
			this.addAdjustment(4050);
			this.addAdjustment(23400);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#4,#5");
			this.addHandle(oHandle3);

			let oHandle4 = new CVmlHandle();
			oHandle4.position = ("#6,#7");
			this.addHandle(oHandle4);
		};
		CVmlGeometryData.prototype.fillCBorderCallout90Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3nfem,l21600,r,21600l,21600nsxe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");

			this.addAdjustment(-1800);
			this.addAdjustment(24300);
			this.addAdjustment(-1800);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCBorderCallout1Type = function() {

			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3nfem,l21600,r,21600l,21600xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");

			this.addAdjustment(-8280);
			this.addAdjustment(24300);
			this.addAdjustment(-1800);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCBorderCallout2Type = function() {

			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3@4@5nfem,l21600,r,21600l,21600xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("val #4");
			this.addGuide("val #5");

			this.addAdjustment(-10080);
			this.addAdjustment(24300);
			this.addAdjustment(-3600);
			this.addAdjustment(4050);
			this.addAdjustment(-1800);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#4,#5");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCBorderCallout3Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat
			this.loadPath("m@0@1l@2@3@4@5@6@7nfem,l21600,r,21600l,21600xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("val #4");
			this.addGuide("val #5");
			this.addGuide("val #6");
			this.addGuide("val #7");

			this.addAdjustment(23400);
			this.addAdjustment(24400);
			this.addAdjustment(25200);
			this.addAdjustment(21600);
			this.addAdjustment(25200);
			this.addAdjustment(4050);
			this.addAdjustment(23400);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#4,#5");
			this.addHandle(oHandle3);

			let oHandle4 = new CVmlHandle();
			oHandle4.position = ("#6,#7");
			this.addHandle(oHandle4);
		};
		CVmlGeometryData.prototype.fillCCallout90Type = function() {

			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3nfem,l21600,r,21600l,21600xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");

			this.addAdjustment(-1800);
			this.addAdjustment(24300);
			this.addAdjustment(-1800);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCCallout1Type = function() {

			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3nfem,l21600,r,21600l,21600nsxe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");

			this.addAdjustment(8280);
			this.addAdjustment(24300);
			this.addAdjustment(-1800);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCCallout2Type = function() {

			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3@4@5nfem,l21600,r,21600l,21600nsxe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("val #4");
			this.addGuide("val #5");

			this.addAdjustment(-10080);
			this.addAdjustment(24300);
			this.addAdjustment(-3600);
			this.addAdjustment(4050);
			this.addAdjustment(-1800);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#4,#5");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCCallout3Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@0@1l@2@3@4@5@6@7nfem,l21600,r,21600l,21600nsxe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("val #4");
			this.addGuide("val #5");
			this.addGuide("val #6");
			this.addGuide("val #7");

			this.addAdjustment(23400);
			this.addAdjustment(24400);
			this.addAdjustment(25200);
			this.addAdjustment(21600);
			this.addAdjustment(25200);
			this.addAdjustment(4050);
			this.addAdjustment(23400);
			this.addAdjustment(4050);

			this.loadConnectorsList("@0,@1;10800,0;10800,21600;0,10800;21600,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#4,#5");
			this.addHandle(oHandle3);

			let oHandle4 = new CVmlHandle();
			oHandle4.position = ("#6,#7");
			this.addHandle(oHandle4);
		};
		CVmlGeometryData.prototype.fillCActionButtonBlankType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0e");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("prod #0 1 2");
			this.addGuide("prod #0 3 2");
			this.addGuide("sum @1 @5 0");
			this.addGuide("sum @2 @5 0");

			this.addAdjustment(1350);
			this.loadConnectorsList("0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0");
			this.loadTextRect("@0,@0,@1,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCActionButtonHomeType = function() {

			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0em@3@9nfl@11@4@28@4@28@10@33@10@33@4@12@4@32@26@32@24@31@24@31@25xem@31@25nfl@32@26em@28@4nfl@33@4em@29@10nfl@29@27@30@27@30@10e");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("prod #0 1 2");
			this.addGuide("prod #0 3 2");
			this.addGuide("sum @1 @5 0");
			this.addGuide("sum @2 @5 0");
			this.addGuide("sum @0 @4 8100");
			this.addGuide("sum @2 8100 @4");
			this.addGuide("sum @0 @3 8100");
			this.addGuide("sum @1 8100 @3");
			this.addGuide("sum @10 0 @9");
			this.addGuide("prod @13 1 16");
			this.addGuide("prod @13 1 8");
			this.addGuide("prod @13 3 16");
			this.addGuide("prod @13 5 16");
			this.addGuide("prod @13 7 16");
			this.addGuide("prod @13 9 16");
			this.addGuide("prod @13 11 16");
			this.addGuide("prod @13 3 4");
			this.addGuide("prod @13 13 16");
			this.addGuide("prod @13 7 8");
			this.addGuide("sum @9 @14 0");
			this.addGuide("sum @9 @16 0");
			this.addGuide("sum @9 @17 0");
			this.addGuide("sum @9 @21 0");
			this.addGuide("sum @11 @15 0");
			this.addGuide("sum @11 @18 0");
			this.addGuide("sum @11 @19 0");
			this.addGuide("sum @11 @20 0");
			this.addGuide("sum @11 @22 0");
			this.addGuide("sum @11 @23 0");
			this.addGuide("sum @3 @5 0");
			this.addGuide("sum @4 @5 0");
			this.addGuide("sum @9 @5 0");
			this.addGuide("sum @10 @5 0");
			this.addGuide("sum @11 @5 0");
			this.addGuide("sum @12 @5 0");
			this.addGuide("sum @24 @5 0");
			this.addGuide("sum @25 @5 0");
			this.addGuide("sum @26 @5 0");
			this.addGuide("sum @27 @5 0");
			this.addGuide("sum @28 @5 0");
			this.addGuide("sum @29 @5 0");
			this.addGuide("sum @30 @5 0");
			this.addGuide("sum @31 @5 0");
			this.addGuide("sum @32 @5 0");
			this.addGuide("sum @33 @5 0");

			this.addAdjustment(1350);
			this.loadConnectorsList("0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0");
			this.loadTextRect("@0,@0,@1,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCActionButtonHelpType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0em@33@27nfqy@3@9@40@27@39@4@37@29l@37@30@36@30@36@29qy@37@28@39@27@3@26@34@27xem@3@31nfqx@35@32@3@10@38@32@3@31xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("prod #0 1 2");
			this.addGuide("prod #0 3 2");
			this.addGuide("sum @1 @5 0");
			this.addGuide("sum @2 @5 0");
			this.addGuide("sum @0 @4 8100");
			this.addGuide("sum @2 8100 @4");
			this.addGuide("sum @0 @3 8100");
			this.addGuide("sum @1 8100 @3");
			this.addGuide("sum @10 0 @9");
			this.addGuide("prod @13 1 7");
			this.addGuide("prod @13 3 14");
			this.addGuide("prod @13 2 7");
			this.addGuide("prod @13 5 14");
			this.addGuide("prod @13 11 28");
			this.addGuide("prod @13 3 7");
			this.addGuide("prod @13 4 7");
			this.addGuide("prod @13 17 28");
			this.addGuide("prod @13 9 14");
			this.addGuide("prod @13 21 28");
			this.addGuide("prod @13 11 14");
			this.addGuide("prod @13 25 28");
			this.addGuide("sum @9 @14 0");
			this.addGuide("sum @9 @16 0");
			this.addGuide("sum @9 @18 0");
			this.addGuide("sum @9 @21 0");
			this.addGuide("sum @9 @23 0");
			this.addGuide("sum @9 @24 0");
			this.addGuide("sum @9 @25 0");
			this.addGuide("sum @11 @15 0");
			this.addGuide("sum @11 @17 0");
			this.addGuide("sum @11 @18 0");
			this.addGuide("sum @11 @19 0");
			this.addGuide("sum @11 @20 0");
			this.addGuide("sum @11 @21 0");
			this.addGuide("sum @11 @22 0");
			this.addGuide("sum @11 @24 0");
			this.addGuide("sum @3 @5 0");
			this.addGuide("sum @4 @5 0");
			this.addGuide("sum @9 @5 0");
			this.addGuide("sum @10 @5 0");
			this.addGuide("sum @11 @5 0");
			this.addGuide("sum @12 @5 0");
			this.addGuide("sum @26 @5 0");
			this.addGuide("sum @27 @5 0");
			this.addGuide("sum @28 @5 0");
			this.addGuide("sum @29 @5 0");
			this.addGuide("sum @30 @5 0");
			this.addGuide("sum @31 @5 0");
			this.addGuide("sum @32 @5 0");
			this.addGuide("sum @33 @5 0");
			this.addGuide("sum @34 @5 0");
			this.addGuide("sum @35 @5 0");
			this.addGuide("sum @36 @5 0");
			this.addGuide("sum @37 @5 0");
			this.addGuide("sum @38 @5 0");
			this.addGuide("sum @39 @5 0");
			this.addGuide("sum @40 @5 0");

			this.addAdjustment(1350);
			this.loadConnectorsList("0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0");
			this.loadTextRect("@0,@0,@1,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCActionButtonInfoType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0em@3@9nfqx@11@4@3@10@12@4@3@9xem@3@25nfqx@33@26@3@27@36@26@3@25xem@32@28nfl@32@29@34@29@34@30@32@30@32@31@37@31@37@30@35@30@35@28xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("prod #0 1 2");
			this.addGuide("prod #0 3 2");
			this.addGuide("sum @1 @5 0");
			this.addGuide("sum @2 @5 0");
			this.addGuide("sum @0 @4 8100");
			this.addGuide("sum @2 8100 @4");
			this.addGuide("sum @0 @3 8100");
			this.addGuide("sum @1 8100 @3");
			this.addGuide("sum @10 0 @9");
			this.addGuide("prod @13 1 32");
			this.addGuide("prod @13 5 32");
			this.addGuide("prod @13 9 32");
			this.addGuide("prod @13 5 16");
			this.addGuide("prod @13 3 8");
			this.addGuide("prod @13 13 32");
			this.addGuide("prod @13 19 32");
			this.addGuide("prod @13 5 8");
			this.addGuide("prod @13 11 16");
			this.addGuide("prod @13 13 16");
			this.addGuide("prod @13 7 8");
			this.addGuide("sum @9 @14 0");
			this.addGuide("sum @9 @15 0");
			this.addGuide("sum @9 @16 0");
			this.addGuide("sum @9 @17 0");
			this.addGuide("sum @9 @18 0");
			this.addGuide("sum @9 @23 0");
			this.addGuide("sum @9 @24 0");
			this.addGuide("sum @11 @17 0");
			this.addGuide("sum @11 @18 0");
			this.addGuide("sum @11 @19 0");
			this.addGuide("sum @11 @20 0");
			this.addGuide("sum @11 @21 0");
			this.addGuide("sum @11 @22 0");
			this.addGuide("sum @3 @5 0");
			this.addGuide("sum @4 @5 0");
			this.addGuide("sum @9 @5 0");
			this.addGuide("sum @10 @5 0");
			this.addGuide("sum @11 @5 0");
			this.addGuide("sum @12 @5 0");
			this.addGuide("sum @25 @5 0");
			this.addGuide("sum @26 @5 0");
			this.addGuide("sum @27 @5 0");
			this.addGuide("sum @28 @5 0");
			this.addGuide("sum @29 @5 0");
			this.addGuide("sum @30 @5 0");
			this.addGuide("sum @31 @5 0");
			this.addGuide("sum @32 @5 0");
			this.addGuide("sum @33 @5 0");
			this.addGuide("sum @34 @5 0");
			this.addGuide("sum @35 @5 0");
			this.addGuide("sum @36 @5 0");
			this.addGuide("sum @37 @5 0");

			this.addAdjustment(1350);
			this.loadConnectorsList("0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0");
			this.loadTextRect("@0,@0,@1,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCActionButtonBackType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0em@12@9nfl@11@4@12@10xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("prod #0 1 2");
			this.addGuide("prod #0 3 2");
			this.addGuide("sum @1 @5 0");
			this.addGuide("sum @2 @5 0");
			this.addGuide("sum @0 @4 8100");
			this.addGuide("sum @2 8100 @4");
			this.addGuide("sum @0 @3 8100");
			this.addGuide("sum @1 8100 @3");
			this.addGuide("sum @4 @5 0");
			this.addGuide("sum @9 @5 0");
			this.addGuide("sum @10 @5 0");
			this.addGuide("sum @11 @5 0");
			this.addGuide("sum @12 @5 0");

			this.addAdjustment(1350);
			this.loadConnectorsList("0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0");
			this.loadTextRect("@0,@0,@1,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCActionButtonNextType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0em@11@9nfl@12@4@11@10xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("prod #0 1 2");
			this.addGuide("prod #0 3 2");
			this.addGuide("sum @1 @5 0");
			this.addGuide("sum @2 @5 0");
			this.addGuide("sum @0 @4 8100");
			this.addGuide("sum @2 8100 @4");
			this.addGuide("sum @0 @3 8100");
			this.addGuide("sum @1 8100 @3");
			this.addGuide("sum @4 @5 0");
			this.addGuide("sum @9 @5 0");
			this.addGuide("sum @10 @5 0");
			this.addGuide("sum @11 @5 0");
			this.addGuide("sum @12 @5 0");

			this.addAdjustment(1350);
			this.loadConnectorsList("0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0");
			this.loadTextRect("@0,@0,@1,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCActionButtonBeginType = function() {this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0em@12@9l@17@4@12@10xem@11@9l@16@9@16@10@11@10xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("prod #0 1 2");
			this.addGuide("prod #0 3 2");
			this.addGuide("sum @1 @5 0");
			this.addGuide("sum @2 @5 0");
			this.addGuide("sum @0 @4 8100");
			this.addGuide("sum @2 8100 @4");
			this.addGuide("sum @0 @3 8100");
			this.addGuide("sum @1 8100 @3");
			this.addGuide("sum @10 0 @9");
			this.addGuide("prod @13 1 8");
			this.addGuide("prod @13 1 4");
			this.addGuide("sum @11 @14 0");
			this.addGuide("sum @11 @15 0");
			this.addGuide("sum @4 @5 0");
			this.addGuide("sum @9 @5 0");
			this.addGuide("sum @10 @5 0");
			this.addGuide("sum @11 @5 0");
			this.addGuide("sum @12 @5 0");
			this.addGuide("sum @16 @5 0");
			this.addGuide("sum @17 @5 0");

			this.addAdjustment(1350);
			this.loadConnectorsList("0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0");
			this.loadTextRect("@0,@0,@1,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCActionButtonEndType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0em@11@9l@16@4@11@10xem@17@9l@12@9@12@10@17@10xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("prod #0 1 2");
			this.addGuide("prod #0 3 2");
			this.addGuide("sum @1 @5 0");
			this.addGuide("sum @2 @5 0");
			this.addGuide("sum @0 @4 8100");
			this.addGuide("sum @2 8100 @4");
			this.addGuide("sum @0 @3 8100");
			this.addGuide("sum @1 8100 @3");
			this.addGuide("sum @10 0 @9");
			this.addGuide("prod @13 3 4");
			this.addGuide("prod @13 7 8");
			this.addGuide("sum @11 @14 0");
			this.addGuide("sum @11 @15 0");
			this.addGuide("sum @4 @5 0");
			this.addGuide("sum @9 @5 0");
			this.addGuide("sum @10 @5 0");
			this.addGuide("sum @11 @5 0");
			this.addGuide("sum @12 @5 0");
			this.addGuide("sum @16 @5 0");
			this.addGuide("sum @17 @5 0");

			this.addAdjustment(1350);
			this.loadConnectorsList("0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0");
			this.loadTextRect("@0,@0,@1,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCActionButtonReturnType = function() { this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0em@12@21nfl@23@9@3@21@24@21@24@20qy@3@19l@25@19qx@26@20l@26@21@11@21@11@20qy@25@10l@3@10qx@22@20l@22@21xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("prod #0 1 2");
			this.addGuide("prod #0 3 2");
			this.addGuide("sum @1 @5 0");
			this.addGuide("sum @2 @5 0");
			this.addGuide("sum @0 @4 8100");
			this.addGuide("sum @2 8100 @4");
			this.addGuide("sum @0 @3 8100");
			this.addGuide("sum @1 8100 @3");
			this.addGuide("sum @10 0 @9");
			this.addGuide("prod @13 7 8");
			this.addGuide("prod @13 3 4");
			this.addGuide("prod @13 5 8");
			this.addGuide("prod @13 3 8");
			this.addGuide("prod @13 1 4");
			this.addGuide("sum @9 @15 0");
			this.addGuide("sum @9 @16 0");
			this.addGuide("sum @9 @18 0");
			this.addGuide("sum @11 @14 0");
			this.addGuide("sum @11 @15 0");
			this.addGuide("sum @11 @16 0");
			this.addGuide("sum @11 @17 0");
			this.addGuide("sum @11 @18 0");
			this.addGuide("sum @3 @5 0");
			this.addGuide("sum @9 @5 0");
			this.addGuide("sum @10 @5 0");
			this.addGuide("sum @11 @5 0");
			this.addGuide("sum @12 @5 0");
			this.addGuide("sum @19 @5 0");
			this.addGuide("sum @20 @5 0");
			this.addGuide("sum @21 @5 0");
			this.addGuide("sum @22 @5 0");
			this.addGuide("sum @23 @5 0");
			this.addGuide("sum @24 @5 0");
			this.addGuide("sum @25 @5 0");
			this.addGuide("sum @26 @5 0");

			this.addAdjustment(1350);
			this.loadConnectorsList("0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0");
			this.loadTextRect("@0,@0,@1,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCActionButtonDocType = function() { this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0em@12@9nfl@12@10@13@10@13@14@15@9xem@15@9nfl@15@14@13@14e");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("prod #0 1 2");
			this.addGuide("prod #0 3 2");
			this.addGuide("sum @1 @5 0");
			this.addGuide("sum @2 @5 0");
			this.addGuide("sum @0 @4 8100");
			this.addGuide("sum @2 8100 @4");
			this.addGuide("prod #0 3 4");
			this.addGuide("sum @3 @11 6075");
			this.addGuide("sum @3 6075 @11");
			this.addGuide("sum @4 @5 4050");
			this.addGuide("sum @13 @5 4050");
			this.addGuide("sum @9 @5 0");
			this.addGuide("sum @10 @5 0");
			this.addGuide("sum @12 @5 0");
			this.addGuide("sum @13 @5 0");
			this.addGuide("sum @14 @5 0");
			this.addGuide("sum @15 @5 0");

			this.addAdjustment(1350);
			this.loadConnectorsList("0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0");
			this.loadTextRect("@0,@0,@1,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCActionButtonSoundType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0em@11@21nfl@11@22@24@22@25@10@25@9@24@21xem@26@21nfl@12@20em@26@4nfl@12@4em@26@22nfl@12@23e");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("prod #0 1 2");
			this.addGuide("prod #0 3 2");
			this.addGuide("sum @1 @5 0");
			this.addGuide("sum @2 @5 0");
			this.addGuide("sum @0 @4 8100");
			this.addGuide("sum @2 8100 @4");
			this.addGuide("sum @0 @3 8100");
			this.addGuide("sum @1 8100 @3");
			this.addGuide("sum @10 0 @9");
			this.addGuide("prod @13 1 8");
			this.addGuide("prod @13 5 16");
			this.addGuide("prod @13 5 8");
			this.addGuide("prod @13 11 16");
			this.addGuide("prod @13 3 4");
			this.addGuide("prod @13 7 8");
			this.addGuide("sum @9 @14 0");
			this.addGuide("sum @9 @15 0");
			this.addGuide("sum @9 @17 0");
			this.addGuide("sum @9 @19 0");
			this.addGuide("sum @11 @15 0");
			this.addGuide("sum @11 @16 0");
			this.addGuide("sum @11 @18 0");
			this.addGuide("sum @4 @5 0");
			this.addGuide("sum @9 @5 0");
			this.addGuide("sum @10 @5 0");
			this.addGuide("sum @11 @5 0");
			this.addGuide("sum @12 @5 0");
			this.addGuide("sum @20 @5 0");
			this.addGuide("sum @21 @5 0");
			this.addGuide("sum @22 @5 0");
			this.addGuide("sum @23 @5 0");
			this.addGuide("sum @24 @5 0");
			this.addGuide("sum @25 @5 0");
			this.addGuide("sum @26 @5 0");

			this.addAdjustment(1350);
			this.loadConnectorsList("0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0");
			this.loadTextRect("@0,@0,@1,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCActionButtonMovieType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0em@11@39nfl@11@44@31@44@32@43@33@43@33@47@35@47@35@45@36@45@38@46@12@46@12@41@38@41@37@42@35@42@35@41@34@40@32@40@31@39xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("prod #0 1 2");
			this.addGuide("prod #0 3 2");
			this.addGuide("sum @1 @5 0");
			this.addGuide("sum @2 @5 0");
			this.addGuide("sum @0 @4 8100");
			this.addGuide("sum @2 8100 @4");
			this.addGuide("sum @0 @3 8100");
			this.addGuide("sum @1 8100 @3");
			this.addGuide("sum @10 0 @9");
			this.addGuide("prod @13 1455 21600");
			this.addGuide("prod @13 1905 21600");
			this.addGuide("prod @13 2325 21600");
			this.addGuide("prod @13 16155 21600");
			this.addGuide("prod @13 17010 21600");
			this.addGuide("prod @13 19335 21600");
			this.addGuide("prod @13 19725 21600");
			this.addGuide("prod @13 20595 21600");
			this.addGuide("prod @13 5280 21600");
			this.addGuide("prod @13 5730 21600");
			this.addGuide("prod @13 6630 21600");
			this.addGuide("prod @13 7492 21600");
			this.addGuide("prod @13 9067 21600");
			this.addGuide("prod @13 9555 21600");
			this.addGuide("prod @13 13342 21600");
			this.addGuide("prod @13 14580 21600");
			this.addGuide("prod @13 15592 21600");
			this.addGuide("sum @11 @14 0");
			this.addGuide("sum @11 @15 0");
			this.addGuide("sum @11 @16 0");
			this.addGuide("sum @11 @17 0");
			this.addGuide("sum @11 @18 0");
			this.addGuide("sum @11 @19 0");
			this.addGuide("sum @11 @20 0");
			this.addGuide("sum @11 @21 0");
			this.addGuide("sum @9 @22 0");
			this.addGuide("sum @9 @23 0");
			this.addGuide("sum @9 @24 0");
			this.addGuide("sum @9 @25 0");
			this.addGuide("sum @9 @26 0");
			this.addGuide("sum @9 @27 0");
			this.addGuide("sum @9 @28 0");
			this.addGuide("sum @9 @29 0");
			this.addGuide("sum @9 @30 0");
			this.addGuide("sum @9 @31 0");
			this.addGuide("sum @4 @5 0");
			this.addGuide("sum @9 @5 0");
			this.addGuide("sum @10 @5 0");
			this.addGuide("sum @11 @5 0");
			this.addGuide("sum @12 @5 0");
			this.addGuide("sum @31 @5 0");
			this.addGuide("sum @32 @5 0");
			this.addGuide("sum @33 @5 0");
			this.addGuide("sum @34 @5 0");
			this.addGuide("sum @35 @5 0");
			this.addGuide("sum @36 @5 0");
			this.addGuide("sum @37 @5 0");
			this.addGuide("sum @38 @5 0");
			this.addGuide("sum @39 @5 0");
			this.addGuide("sum @40 @5 0");
			this.addGuide("sum @41 @5 0");
			this.addGuide("sum @42 @5 0");
			this.addGuide("sum @43 @5 0");
			this.addGuide("sum @44 @5 0");
			this.addGuide("sum @45 @5 0");
			this.addGuide("sum @46 @5 0");
			this.addGuide("sum @47 @5 0");

			this.addAdjustment(1350);
			this.loadConnectorsList("0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0");
			this.loadTextRect("@0,@0,@1,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCArcType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinRound;
			this.loadPath("wr-21600,,21600,43200,,,21600,21600nfewr-21600,,21600,43200,,,21600,21600l,21600nsxe");

			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("val #4");

			this.addAdjustment(-5898240);
			this.addAdjustment(0);
			this.addAdjustment(0);
			this.addAdjustment(ShapeSize);
			this.addAdjustment(ShapeSize);

			this.loadConnectorsList("0,0;21600,21600;0,21600");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("@2,#0");
			oHandle1.polar = ("@0,@1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("@2,#1");
			oHandle2.polar = ("@0,@1");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCLineType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinRound;
			this.loadPath("m0,0l21600,21600nfe");

			this.loadConnectorsList("0,0;21600,21600;10800,10800");
		};
		CVmlGeometryData.prototype.fillCBentArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m21600,6079l@0,0@0@1,12427@1qx,12158l,21600@4,21600@4,12158qy12427@2l@0@2@0,12158xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("sum 12158 0 #1");
			this.addGuide("sum @2 0 #1");
			this.addGuide("prod @3 32768 32059");
			this.addGuide("prod @4 1 2");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("prod @6 #1 6079");
			this.addGuide("sum @7 #0 0");

			this.addAdjustment(18000);
			this.addAdjustment(3000);

			this.loadConnectorsList("@0,0;@0,12158;@5,21600;21600,6079");
			this.addConnectorAngle(270);
			this.addConnectorAngle(90);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("12427,@1,@8,@2;0,12158,@4,21600");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			oHandle1.xrange = ("12427,21600");
			oHandle1.yrange = ("0,6079");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCBentUpArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m@4,l@0@2@5@2@5@12,0@12,,21600@1,21600@1@2,21600@2xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("prod #0 1 2");
			this.addGuide("sum @3 10800 0");
			this.addGuide("sum 21600 #0 #1");
			this.addGuide("sum #1 #2 0");
			this.addGuide("prod @6 1 2");
			this.addGuide("prod #1 2 1");
			this.addGuide("sum @8 0 21600");
			this.addGuide("prod 21600 @0 @1");
			this.addGuide("prod 21600 @4 @1");
			this.addGuide("prod 21600 @5 @1");
			this.addGuide("prod 21600 @7 @1");
			this.addGuide("prod #1 1 2");
			this.addGuide("sum @5 0 @4");
			this.addGuide("sum @0 0 @4");
			this.addGuide("prod @2 @15 @16");

			this.addAdjustment(9257);
			this.addAdjustment(18514);
			this.addAdjustment(7200);

			this.loadConnectorsList("@4,0;@0,@2;0,@11;@14,21600;@1,@13;21600,@2");
			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);
			this.addConnectorAngle(0);

			this.loadTextRect("0,@12,@1,21600;@5,@17,@1,21600");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("@2,@9");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#1,#2");
			oHandle2.xrange = ("@4,21600");
			oHandle2.yrange = ("0,@0");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCBevelType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0e");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("prod #0 1 2");
			this.addGuide("prod #0 3 2");
			this.addGuide("sum @1 @5 0");
			this.addGuide("sum @2 @5 0");

			this.addAdjustment(2700);

			this.loadConnectorsList("0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0");
			this.addConnectorAngle(270);
			this.addConnectorAngle(90);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@0,@0,@1,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCBlockArcType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("al10800,10800@0@0@2@14,10800,10800,10800,10800@3@15xe");

			this.addGuide("val #1");
			this.addGuide("val #0");
			this.addGuide("sum 0 0 #0");
			this.addGuide("sumangle #0 0 180");
			this.addGuide("sumangle #0 0 90");
			this.addGuide("prod @4 2 1");
			this.addGuide("sumangle #0 90 0");
			this.addGuide("prod @6 2 1");
			this.addGuide("abs #0");
			this.addGuide("sumangle @8 0 90");
			this.addGuide("if @9 @7 @5");
			this.addGuide("sumangle @10 0 360");
			this.addGuide("if @10 @11 @10");
			this.addGuide("sumangle @12 0 360");
			this.addGuide("if @12 @13 @12");
			this.addGuide("sum 0 0 @14");
			this.addGuide("val 10800");
			this.addGuide("sum 10800 0 #1");
			this.addGuide("prod #1 1 2");
			this.addGuide("sum @18 5400 0");
			this.addGuide("cos @19 #0");
			this.addGuide("sin @19 #0");
			this.addGuide("sum @20 10800 0");
			this.addGuide("sum @21 10800 0");
			this.addGuide("sum 10800 0 @20");
			this.addGuide("sum #1 10800 0");
			this.addGuide("if @9 @17 @25");
			this.addGuide("if @9 0 21600");
			this.addGuide("cos 10800 #0");
			this.addGuide("sin 10800 #0");
			this.addGuide("sin #1 #0");
			this.addGuide("sum @28 10800 0");
			this.addGuide("sum @29 10800 0");
			this.addGuide("sum @30 10800 0");
			this.addGuide("if @4 0 @31");
			this.addGuide("if #0 @34 0");
			this.addGuide("if @6 @35 @31");
			this.addGuide("sum 21600 0 @36");
			this.addGuide("if @4 0 @33");
			this.addGuide("if #0 @38 @32");
			this.addGuide("if @6 @39 0");
			this.addGuide("if @4 @32 21600");
			this.addGuide("if @6 @41 @33");

			this.addAdjustment(11796480);
			this.addAdjustment(5400);

			this.loadConnectorsList("10800,@27;@22,@23;10800,@26;@24,@23");

			this.loadTextRect("@36,@40,@37,@42");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#1,#0");
			oHandle1.polar = ("10800,10800");
			oHandle1.radiusrange = ("0,10800");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCBracePairType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@9,nfqx@0@0l@0@7qy0@4@0@8l@0@6qy@9,21600em@10,nfqx@5@0l@5@7qy21600@4@5@8l@5@6qy@10,21600em@9,nsqx@0@0l@0@7qy0@4@0@8l@0@6qy@9,21600l@10,21600qx@5@6l@5@8qy21600@4@5@7l@5@0qy@10,xe");

			this.addGuide("val #0");
			this.addGuide("val width");
			this.addGuide("val height");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("sum @4 0 #0");
			this.addGuide("sum @4 #0 0");
			this.addGuide("prod #0 2 1");
			this.addGuide("sum width 0 @9");
			this.addGuide("prod #0 9598 32768");
			this.addGuide("sum height 0 @11");
			this.addGuide("sum @11 #0 0");
			this.addGuide("sum width 0 @13");

			this.addAdjustment(1800);

			this.loadConnectorsList("@3,0;0,@4;@3,@2;@1,@4");
			this.loadTextRect("@13,@11,@14,@12");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("topLeft,#0");
			oHandle1.switchHandle = ("true");
			oHandle1.yrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCBracketPairType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinRound;
			//Encaps: Flat

			this.loadPath("m@0,nfqx0@0l0@2qy@0,21600em@1,nfqx21600@0l21600@2qy@1,21600em@0,nsqx0@0l0@2qy@0,21600l@1,21600qx21600@2l21600@0qy@1,xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod @0 2929 10000");
			this.addGuide("sum width 0 @3");
			this.addGuide("sum height 0 @3");
			this.addGuide("val width");
			this.addGuide("val height");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");

			this.addAdjustment(3600);

			this.loadConnectorsList("@8,0;0,@9;@8,@7;@6,@9");
			this.loadTextRect("@3,@3,@4,@5");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCCanType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinRound;
			this.loadPath("m10800,qx0@1l0@2qy10800,21600,21600@2l21600@1qy10800,xem0@1qy10800@0,21600@1nfe");

			this.addGuide("val #0");
			this.addGuide("prod #0 1 2");
			this.addGuide("sum height 0 @1");

			this.addAdjustment(5400);

			this.loadConnectorsList("10800,@0;10800,0;0,10800;10800,21600;21600,10800");
			this.addConnectorAngle(270);
			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("0,@0,21600,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("center,#0");
			oHandle1.yrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCChevronType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m@0,l,0@1,10800,,21600@0,21600,21600,10800xe");

			this.addGuide("val #0");
			this.addGuide("sum 21600 0 @0");
			this.addGuide("prod #0 1 2");

			this.addAdjustment(16200);
			this.loadConnectorsList("@2,0;@1,10800;@2,21600;21600,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("0,0,10800,21600;0,0,16200,21600;0,0,21600,21600");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("0,21600");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCCircularArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("al10800,10800@8@8@4@6,10800,10800,10800,10800@9@7l@30@31@17@18@24@25@15@16@32@33xe");

			this.addGuide("val #1");
			this.addGuide("val #0");
			this.addGuide("sum #1 0 #0");
			this.addGuide("val 10800");
			this.addGuide("sum 0 0 #1");
			this.addGuide("sumangle @2 360 0");
			this.addGuide("if @2 @2 @5");
			this.addGuide("sum 0 0 @6");
			this.addGuide("val #2");
			this.addGuide("sum 0 0 #0");
			this.addGuide("sum #2 0 2700");
			this.addGuide("cos @10 #1");
			this.addGuide("sin @10 #1");
			this.addGuide("cos 13500 #1");
			this.addGuide("sin 13500 #1");
			this.addGuide("sum @11 10800 0");
			this.addGuide("sum @12 10800 0");
			this.addGuide("sum @13 10800 0");
			this.addGuide("sum @14 10800 0");
			this.addGuide("prod #2 1 2 ");
			this.addGuide("sum @19 5400 0");
			this.addGuide("cos @20 #1");
			this.addGuide("sin @20 #1");
			this.addGuide("sum @21 10800 0");
			this.addGuide("sum @12 @23 @22");
			this.addGuide("sum @22 @23 @11");
			this.addGuide("cos 10800 #1");
			this.addGuide("sin 10800 #1");
			this.addGuide("cos #2 #1");
			this.addGuide("sin #2 #1");
			this.addGuide("sum @26 10800 0");
			this.addGuide("sum @27 10800 0");
			this.addGuide("sum @28 10800 0");
			this.addGuide("sum @29 10800 0");
			this.addGuide("sum @19 5400 0 ");
			this.addGuide("cos @34 #0");
			this.addGuide("sin @34 #0");
			this.addGuide("mid #0 #1");
			this.addGuide("sumangle @37 180 0");
			this.addGuide("if @2 @37 @38");
			this.addGuide("cos 10800 @39");
			this.addGuide("sin 10800 @39");
			this.addGuide("cos #2 @39");
			this.addGuide("sin #2 @39");
			this.addGuide("sum @40 10800 0");
			this.addGuide("sum @41 10800 0");
			this.addGuide("sum @42 10800 0");
			this.addGuide("sum @43 10800 0");
			this.addGuide("sum @35 10800 0");
			this.addGuide("sum @36 10800 0");

			this.addAdjustment(-11796480);
			this.addAdjustment(0);
			this.addAdjustment(5400);

			this.loadConnectorsList("@44,@45;@48,@49;@46,@47;@17,@18;@24,@25;@15,@16");
			this.loadTextRect("3163,3163,18437,18437");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("@3,#0");
			oHandle1.polar = ("10800,10800");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#2,#1");
			oHandle2.polar = ("10800,10800");
			oHandle2.radiusrange = ("0,10800");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCCloudCalloutType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinRound;

			this.loadPath("ar,7165,4345,13110,1950,7185,1080,12690,475,11732,4835,17650,1080,12690,2910,17640,2387,9757,10107,20300,2910,17640,8235,19545,7660,12382,14412,21597,8235,19545,14280,18330,12910,11080,18695,18947,14280,18330,18690,15045,14822,5862,21597,15082,18690,15045,20895,7665,15772,2592,21105,9865,20895,7665,19140,2715,14330,,19187,6595,19140,2715,14910,1170,10992,,15357,5945,14910,1170,11250,1665,6692,650,12025,7917,11250,1665,7005,2580,1912,1972,8665,11162,7005,2580,1950,7185xear,7165,4345,13110,1080,12690,2340,13080nfear475,11732,4835,17650,2910,17640,3465,17445nfear7660,12382,14412,21597,7905,18675,8235,19545nfear7660,12382,14412,21597,14280,18330,14400,17370nfear12910,11080,18695,18947,18690,15045,17070,11475nfear15772,2592,21105,9865,20175,9015,20895,7665nfear14330,,19187,6595,19200,3345,19140,2715nfear14330,,19187,6595,14910,1170,14550,1980nfear10992,,15357,5945,11250,1665,11040,2340nfear1912,1972,8665,11162,7650,3270,7005,2580nfear1912,1972,8665,11162,1950,7185,2070,7890nfem@23@37qx@35@24@23@36@34@24@23@37xem@16@33qx@31@17@16@32@30@17@16@33xem@38@29qx@27@39@38@28@26@39@38@29xe");

			this.addGuide("sum #0 0 10800");
			this.addGuide("sum #1 0 10800");
			this.addGuide("cosatan2 10800 @0 @1");
			this.addGuide("sinatan2 10800 @0 @1");
			this.addGuide("sum @2 10800 0");
			this.addGuide("sum @3 10800 0");
			this.addGuide("sum @4 0 #0");
			this.addGuide("sum @5 0 #1");
			this.addGuide("mod @6 @7 0");
			this.addGuide("prod 600 11 1");
			this.addGuide("sum @8 0 @9");
			this.addGuide("prod @10 1 3");
			this.addGuide("prod 600 3 1");
			this.addGuide("sum @11 @12 0");
			this.addGuide("prod @13 @6 @8");
			this.addGuide("prod @13 @7 @8");
			this.addGuide("sum @14 #0 0");
			this.addGuide("sum @15 #1 0");
			this.addGuide("prod 600 8 1");
			this.addGuide("prod @11 2 1");
			this.addGuide("sum @18 @19 0");
			this.addGuide("prod @20 @6 @8");
			this.addGuide("prod @20 @7 @8");
			this.addGuide("sum @21 #0 0");
			this.addGuide("sum @22 #1 0");
			this.addGuide("prod 600 2 1");
			this.addGuide("sum #0 600 0");
			this.addGuide("sum #0 0 600");
			this.addGuide("sum #1 600 0");
			this.addGuide("sum #1 0 600");
			this.addGuide("sum @16 @25 0");
			this.addGuide("sum @16 0 @25");
			this.addGuide("sum @17 @25 0");
			this.addGuide("sum @17 0 @25");
			this.addGuide("sum @23 @12 0");
			this.addGuide("sum @23 0 @12");
			this.addGuide("sum @24 @12 0");
			this.addGuide("sum @24 0 @12");
			this.addGuide("val #0");
			this.addGuide("val #1");

			this.addAdjustment(1350);
			this.addAdjustment(25920);

			this.loadConnectorsList("67,10800;10800,21577;21582,10800;10800,1235;@38,@39");
			this.loadTextRect("2977,3262,17087,17337");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCCubeType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m@0,l0@0,,21600@1,21600,21600@2,21600,xem0@0nfl@1@0,21600,em@1@0nfl@1,21600e");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("mid height #0");
			this.addGuide("prod @1 1 2");
			this.addGuide("prod @2 1 2");
			this.addGuide("mid width #0");

			this.addAdjustment(5400);

			this.loadConnectorsList("@6,0;@4,@0;0,@3;@4,21600;@1,@3;21600,@5");

			this.addConnectorAngle(270);
			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);
			this.addConnectorAngle(0);

			this.loadTextRect("0,@0,@1,21600");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("topLeft,#0");
			oHandle1.switchHandle = ("true");
			oHandle1.yrange = ("0,21600");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCCurvedDownArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("wr,0@3@23,0@22@4,0@15,0@1@23@7,0@13@2l@14@2@8@22@12@2at,0@3@23@11@2@17@26@15,0@1@23@17@26@15@22xewr,0@3@23@4,0@17@26nfe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("sum #0 width #1");
			this.addGuide("prod @3 1 2");
			this.addGuide("sum #1 #1 width");
			this.addGuide("sum @5 #1 #0");
			this.addGuide("prod @6 1 2");
			this.addGuide("mid width #0");
			this.addGuide("sum height 0 #2");
			this.addGuide("ellipse @9 height @4");
			this.addGuide("sum @4 @10 0");
			this.addGuide("sum @11 #1 width");
			this.addGuide("sum @7 @10 0");
			this.addGuide("sum @12 width #0 ");
			this.addGuide("sum @5 0 #0");
			this.addGuide("prod @15 1 2");
			this.addGuide("mid @4 @7");
			this.addGuide("sum #0 #1 width");
			this.addGuide("prod @18 1 2");
			this.addGuide("sum @17 0 @19");
			this.addGuide("val width");
			this.addGuide("val height");
			this.addGuide("prod height 2 1");
			this.addGuide("sum @17 0 @4 ");
			this.addGuide("ellipse @24 @4 height");
			this.addGuide("sum height 0 @25");
			this.addGuide("sum @8 128 0");
			this.addGuide("prod @5 1 2");
			this.addGuide("sum @5 0 128");
			this.addGuide("sum #0 @17 @12");
			this.addGuide("ellipse @20 @4 height");
			this.addGuide("sum width 0 #0");
			this.addGuide("prod @32 1 2");
			this.addGuide("prod height height 1");
			this.addGuide("prod @9 @9 1");
			this.addGuide("sum @34 0 @35");
			this.addGuide("sqrt @36");
			this.addGuide("sum @37 height 0");
			this.addGuide("prod width height @38");
			this.addGuide("sum @39 64 0");
			this.addGuide("prod #0 1 2");
			this.addGuide("ellipse @33 @41 height");
			this.addGuide("sum height 0 @42");
			this.addGuide("sum @43 64 0");
			this.addGuide("prod @4 1 2");
			this.addGuide("sum #1 0 @45");
			this.addGuide("prod height 4390 32768");
			this.addGuide("prod height 28378 32768");

			this.addAdjustment(12960);
			this.addAdjustment(19440);
			this.addAdjustment(14400);

			this.loadConnectorsList("@17,0;@16,@22;@12,@2;@8,@22;@14,@2");

			this.addConnectorAngle(270);
			this.addConnectorAngle(90);
			this.addConnectorAngle(90);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@45,@47,@46,@48");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,bottomRight");
			oHandle1.xrange = ("@40,@29");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle1.position = ("#1,bottomRight");
			oHandle1.xrange = ("@27,@21");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("bottomRight,#2");
			oHandle3.yrange = ("@44,@22");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCCurvedLeftArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("wr@22,0@21@3,,0@21@4@22@14@21@1@21@7@2@12l@2@13,0@8@2@11at@22,0@21@3@2@10@24@16@22@14@21@1@24@16,0@14xear@22@14@21@1@21@7@24@16nfe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("sum #0 width #1");
			this.addGuide("prod @3 1 2");
			this.addGuide("sum #1 #1 width");
			this.addGuide("sum @5 #1 #0");
			this.addGuide("prod @6 1 2");
			this.addGuide("mid width #0");
			this.addGuide("ellipse #2 height @4");
			this.addGuide("sum @4 @9 0");
			this.addGuide("sum @10 #1 width");
			this.addGuide("sum @7 @9 0");
			this.addGuide("sum @11 width #0");
			this.addGuide("sum @5 0 #0");
			this.addGuide("prod @14 1 2");
			this.addGuide("mid @4 @7");
			this.addGuide("sum #0 #1 width");
			this.addGuide("prod @17 1 2");
			this.addGuide("sum @16 0 @18");
			this.addGuide("val width");
			this.addGuide("val height");
			this.addGuide("sum 0 0 height");
			this.addGuide("sum @16 0 @4");
			this.addGuide("ellipse @23 @4 height");
			this.addGuide("sum @8 128 0");
			this.addGuide("prod @5 1 2");
			this.addGuide("sum @5 0 128");
			this.addGuide("sum #0 @16 @11");
			this.addGuide("sum width 0 #0");
			this.addGuide("prod @29 1 2");
			this.addGuide("prod height height 1");
			this.addGuide("prod #2 #2 1");
			this.addGuide("sum @31 0 @32");
			this.addGuide("sqrt @33");
			this.addGuide("sum @34 height 0");
			this.addGuide("prod width height @35");
			this.addGuide("sum @36 64 0");
			this.addGuide("prod #0 1 2");
			this.addGuide("ellipse @30 @38 height");
			this.addGuide("sum @39 0 64");
			this.addGuide("prod @4 1 2");
			this.addGuide("sum #1 0 @41");
			this.addGuide("prod height 4390 32768");
			this.addGuide("prod height 28378 32768");

			this.addAdjustment(12960);
			this.addAdjustment(19440);
			this.addAdjustment(7200);

			this.loadConnectorsList("0,@15;@2,@11;0,@8;@2,@13;@21,@16");

			this.addConnectorAngle(180);
			this.addConnectorAngle(180);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@43,@41,@44,@42");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("topLeft,#0");
			oHandle1.yrange = ("@37,@27");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle1.position = ("topLeft,#1");
			oHandle1.yrange = ("@25,@20");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#2,bottomRight");
			oHandle3.xrange = ("0,@40");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCCurvedRightArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("ar,0@23@3@22,,0@4,0@15@23@1,0@7@2@13l@2@14@22@8@2@12wa,0@23@3@2@11@26@17,0@15@23@1@26@17@22@15xear,0@23@3,0@4@26@17nfe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("sum #0 width #1");
			this.addGuide("prod @3 1 2");
			this.addGuide("sum #1 #1 width");
			this.addGuide("sum @5 #1 #0");
			this.addGuide("prod @6 1 2");
			this.addGuide("mid width #0");
			this.addGuide("sum height 0 #2");
			this.addGuide("ellipse @9 height @4");
			this.addGuide("sum @4 @10 0");
			this.addGuide("sum @11 #1 width");
			this.addGuide("sum @7 @10 0");
			this.addGuide("sum @12 width #0");
			this.addGuide("sum @5 0 #0");
			this.addGuide("prod @15 1 2");
			this.addGuide("mid @4 @7");
			this.addGuide("sum #0 #1 width");
			this.addGuide("prod @18 1 2");
			this.addGuide("sum @17 0 @19");
			this.addGuide("val width");
			this.addGuide("val height");
			this.addGuide("prod height 2 1");
			this.addGuide("sum @17 0 @4");
			this.addGuide("ellipse @24 @4 height");
			this.addGuide("sum height 0 @25");
			this.addGuide("sum @8 128 0");
			this.addGuide("prod @5 1 2");
			this.addGuide("sum @5 0 128");
			this.addGuide("sum #0 @17 @12");
			this.addGuide("ellipse @20 @4 height");
			this.addGuide("sum width 0 #0");
			this.addGuide("prod @32 1 2");
			this.addGuide("prod height height 1");
			this.addGuide("prod @9 @9 1");
			this.addGuide("sum @34 0 @35");
			this.addGuide("sqrt @36");
			this.addGuide("sum @37 height 0");
			this.addGuide("prod width height @38");
			this.addGuide("sum @39 64 0");
			this.addGuide("prod #0 1 2");
			this.addGuide("ellipse @33 @41 height");
			this.addGuide("sum height 0 @42");
			this.addGuide("sum @43 64 0");
			this.addGuide("prod @4 1 2");
			this.addGuide("sum #1 0 @45");
			this.addGuide("prod height 4390 32768");
			this.addGuide("prod height 28378 32768");

			this.addAdjustment(12960);
			this.addAdjustment(19440);
			this.addAdjustment(14400);

			this.loadConnectorsList("0,@17;@2,@14;@22,@8;@2,@12;@22,@16");

			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);
			this.addConnectorAngle(0);
			this.addConnectorAngle(0);

			this.loadTextRect("@47,@45,@48,@46");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("bottomRight,#0");
			oHandle1.yrange = ("@40,@29");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle1.position = ("bottomRight,#1");
			oHandle1.yrange = ("@27,@21");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#2,bottomRight");
			oHandle3.xrange = ("@44,@22");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCCurvedUpArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("ar0@22@3@21,,0@4@21@14@22@1@21@7@21@12@2l@13@2@8,0@11@2wa0@22@3@21@10@2@16@24@14@22@1@21@16@24@14,xewr@14@22@1@21@7@21@16@24nfe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("sum #0 width #1");
			this.addGuide("prod @3 1 2");
			this.addGuide("sum #1 #1 width ");
			this.addGuide("sum @5 #1 #0");
			this.addGuide("prod @6 1 2");
			this.addGuide("mid width #0");
			this.addGuide("ellipse #2 height @4");
			this.addGuide("sum @4 @9 0 ");
			this.addGuide("sum @10 #1 width");
			this.addGuide("sum @7 @9 0 ");
			this.addGuide("sum @11 width #0 ");
			this.addGuide("sum @5 0 #0 ");
			this.addGuide("prod @14 1 2 ");
			this.addGuide("mid @4 @7 ");
			this.addGuide("sum #0 #1 width ");
			this.addGuide("prod @17 1 2 ");
			this.addGuide("sum @16 0 @18 ");
			this.addGuide("val width ");
			this.addGuide("val height ");
			this.addGuide("sum 0 0 height");
			this.addGuide("sum @16 0 @4 ");
			this.addGuide("ellipse @23 @4 height ");
			this.addGuide("sum @8 128 0 ");
			this.addGuide("prod @5 1 2 ");
			this.addGuide("sum @5 0 128 ");
			this.addGuide("sum #0 @16 @11 ");
			this.addGuide("sum width 0 #0 ");
			this.addGuide("prod @29 1 2 ");
			this.addGuide("prod height height 1 ");
			this.addGuide("prod #2 #2 1 ");
			this.addGuide("sum @31 0 @32 ");
			this.addGuide("sqrt @33 ");
			this.addGuide("sum @34 height 0 ");
			this.addGuide("prod width height @35");
			this.addGuide("sum @36 64 0 ");
			this.addGuide("prod #0 1 2 ");
			this.addGuide("ellipse @30 @38 height ");
			this.addGuide("sum @39 0 64 ");
			this.addGuide("prod @4 1 2");
			this.addGuide("sum #1 0 @41 ");
			this.addGuide("prod height 4390 32768");
			this.addGuide("prod height 28378 32768");

			this.addAdjustment(12960);
			this.addAdjustment(19440);
			this.addAdjustment(7200);

			this.loadConnectorsList("@8,0;@11,@2;@15,0;@16,@21;@13,@2");

			this.addConnectorAngle(270);
			this.addConnectorAngle(270);
			this.addConnectorAngle(270);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@41,@43,@42,@44");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("@37,@27");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle1.position = ("#1,topLeft");
			oHandle1.xrange = ("@25,@20");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("bottomRight,#2");
			oHandle3.yrange = ("0,@40");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCDiamondType = function() {
			this.concentricFill = true;

			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m10800,l,10800,10800,21600,21600,10800xe");

			this.loadConnectorsList("Rectangle");
			this.loadTextRect("5400,5400,16200,16200");
		};
		CVmlGeometryData.prototype.fillCDonutType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinRound;
			this.loadPath("m,10800qy10800,,21600,10800,10800,21600,,10800xm@0,10800qy10800@2@1,10800,10800@0@0,10800xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod @0 2929 10000");
			this.addGuide("sum width 0 @3");
			this.addGuide("sum height 0 @3");

			this.addAdjustment(5400);
			this.loadConnectorsList("10800,0;3163,3163;0,10800;3163,18437;10800,21600;18437,18437;21600,10800;18437,3163");
			this.loadTextRect("3163,3163,18437,18437");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,center");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCDownArrowCalloutType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m,l21600,,21600@0@5@0@5@2@4@2,10800,21600@1@2@3@2@3@0,0@0xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("sum 21600 0 #3");
			this.addGuide("prod #0 1 2");

			this.addAdjustment(14400);
			this.addAdjustment(5400);
			this.addAdjustment(18000);
			this.addAdjustment(8100);

			this.loadConnectorsList("10800,0;0,@6;10800,21600;21600,@6");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("0,0,21600,@0");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("topLeft,#0");
			oHandle1.yrange = ("0,@2");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#1,bottomRight");
			oHandle2.xrange = ("0,@3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#3,#2");
			oHandle3.xrange = ("@1,10800");
			oHandle3.yrange = ("@0,21600");
			this.addHandle(oHandle3);

		};
		CVmlGeometryData.prototype.fillCDownArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m0@0l@1@0@1,0@2,0@2@0,21600@0,10800,21600xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("sum height 0 #1");
			this.addGuide("sum 10800 0 #1");
			this.addGuide("sum width 0 #0");
			this.addGuide("prod @4 @3 10800");
			this.addGuide("sum width 0 @5");

			this.addAdjustment(16200);
			this.addAdjustment(5400);

			this.loadConnectorsList("10800,0;0,@0;10800,21600;21600,@0");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@1,0,@2,@6");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#1,#0");
			oHandle1.xrange = ("0,10800");
			oHandle1.yrange = ("0,21600");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCEllipseType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m,10800qy10800,,21600,10800,10800,21600,,10800xe");
			this.loadTextRect("3233,3233,18367,18367");
		};
		CVmlGeometryData.prototype.fillCEllipseRibbonType = function() {this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("ar@9@38@8@37,0@27@0@26@9@13@8@4@0@25@22@25@9@38@8@37@22@26@3@27l@7@40@3,wa@9@35@8@10@3,0@21@33@9@36@8@1@21@31@20@31@9@35@8@10@20@33,,l@5@40xewr@9@36@8@1@20@31@0@32nfl@20@33ear@9@36@8@1@21@31@22@32nfl@21@33em@0@26nfl@0@32em@22@26nfl@22@32e");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val width");
			this.addGuide("val height");
			this.addGuide("prod width 1 8");
			this.addGuide("prod width 1 2");
			this.addGuide("prod width 7 8");
			this.addGuide("prod width 3 2");
			this.addGuide("sum 0 0 @6");
			this.addGuide("sum height 0 #2");
			this.addGuide("prod @10 30573 4096");
			this.addGuide("prod @11 2 1");
			this.addGuide("sum height 0 @12");
			this.addGuide("sum @11 #2 0");
			this.addGuide("sum @11 height #1");
			this.addGuide("sum height 0 #1");
			this.addGuide("prod @16 1 2");
			this.addGuide("sum @11 @17 0");
			this.addGuide("sum @14 #1 height");
			this.addGuide("sum #0 @5 0");
			this.addGuide("sum width 0 @20");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum @6 0 #0");
			this.addGuide("ellipse @23 width @11");
			this.addGuide("sum @24 height @11");
			this.addGuide("sum @25 @11 @19");
			this.addGuide("sum #2 @11 @19");
			this.addGuide("prod @11 2391 32768");
			this.addGuide("sum @6 0 @20");
			this.addGuide("ellipse @29 width @11");
			this.addGuide("sum #1 @30 @11");
			this.addGuide("sum @25 #1 height");
			this.addGuide("sum height @30 @14");
			this.addGuide("sum @11 @14 0");
			this.addGuide("sum height 0 @34");
			this.addGuide("sum @35 @19 @11");
			this.addGuide("sum @10 @15 @11");
			this.addGuide("sum @35 @15 @11");
			this.addGuide("sum @28 @14 @18");
			this.addGuide("sum height 0 @39");
			this.addGuide("sum @19 0 @18");
			this.addGuide("prod @41 2 3");
			this.addGuide("sum #1 0 @42");
			this.addGuide("sum #2 0 @42");
			this.addGuide("min @44 20925");
			this.addGuide("prod width 3 8");
			this.addGuide("sum @46 0 4");

			this.addAdjustment(5400);
			this.addAdjustment(5400);
			this.addAdjustment(18900);

			this.loadConnectorsList("@6,0;@5,@36;@6,@1;@7,@36");
			this.loadTextRect("@0,@22,@19,@1");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,bottomRight");
			oHandle1.xrange = ("@5,@47");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("center,#1");
			oHandle2.yrange = ("@10,@43");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("topLeft,#2");
			oHandle3.yrange = ("@27,@45");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCEllipseRibbon2Type = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("wr@9@34@8@35,0@24@0@23@9,0@8@11@0@22@19@22@9@34@8@35@19@23@3@24l@7@36@3@4at@9@31@8@32@3@4@18@30@9@1@8@33@18@28@17@28@9@31@8@32@17@30,0@4l@5@36xear@9@1@8@33@17@28@0@29nfl@17@30ewr@9@1@8@33@18@28@19@29nfl@18@30em@0@23nfl@0@29em@19@23nfl@19@29e");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val width");
			this.addGuide("val height");
			this.addGuide("prod width 1 8");
			this.addGuide("prod width 1 2");
			this.addGuide("prod width 7 8");
			this.addGuide("prod width 3 2");
			this.addGuide("sum 0 0 @6");
			this.addGuide("prod #2 30573 4096");
			this.addGuide("prod @10 2 1");
			this.addGuide("sum @10 height #2");
			this.addGuide("sum @10 #1 0");
			this.addGuide("prod #1 1 2");
			this.addGuide("sum @10 @14 0");
			this.addGuide("sum @12 0 #1");
			this.addGuide("sum #0 @5 0");
			this.addGuide("sum width 0 @17");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum @6 0 #0");
			this.addGuide("ellipse @20 width @10");
			this.addGuide("sum @10 0 @21");
			this.addGuide("sum @22 @16 @10");
			this.addGuide("sum #2 @16 @10");
			this.addGuide("prod @10 2391 32768");
			this.addGuide("sum @6 0 @17");
			this.addGuide("ellipse @26 width @10");
			this.addGuide("sum @10 #1 @27");
			this.addGuide("sum @22 #1 0");
			this.addGuide("sum @12 0 @27");
			this.addGuide("sum height 0 #2");
			this.addGuide("sum @10 @12 0");
			this.addGuide("sum @32 @10 @16");
			this.addGuide("sum @31 @10 @13");
			this.addGuide("sum @32 @10 @13");
			this.addGuide("sum @25 @12 @15");
			this.addGuide("sum @16 0 @15");
			this.addGuide("prod @37 2 3");
			this.addGuide("sum @1 @38 0");
			this.addGuide("sum #2 @38 0");
			this.addGuide("max @40 675");
			this.addGuide("prod width 3 8");
			this.addGuide("sum @42 0 4");

			this.addAdjustment(5400);
			this.addAdjustment(16200);
			this.addAdjustment(2700);

			this.loadConnectorsList("@6,0;@5,@36;@6,@1;@7,@36");
			this.loadTextRect("@0,@22,@19,@1");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("topLeft,#0");
			oHandle1.xrange = ("@5,@43");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("center,#1");
			oHandle2.yrange = ("@39,@31");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("topLeft,#2");
			oHandle3.yrange = ("@41,@24 ");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCFlowChartAlternateProcessType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m@0,qx0@0l0@2qy@0,21600l@1,21600qx21600@2l21600@0qy@1,xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod @0 2929 10000");
			this.addGuide("sum width 0 @3");
			this.addGuide("sum height 0 @3");
			this.addGuide("val width");
			this.addGuide("val height");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");

			this.addAdjustment(2700);

			this.loadConnectorsList("@8,0;0,@9;@8,@7;@6,@9");
			this.loadTextRect("@3,@3,@4,@5");

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCFlowChartCollateType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m21600,21600l,21600,21600,,,xe");
			this.loadConnectorsList("10800,0;10800,10800;10800,21600");
			this.loadTextRect("5400,5400,16200,16200");
		};
		CVmlGeometryData.prototype.fillCFlowChartConnectorType = function() {
			this.concentricFill = true;

			this.join = ODRAW.lineJoinRound;

			this.loadPath("m10800,qx,10800,10800,21600,21600,10800,10800,xe");

			this.loadConnectorsList("10800,0;3163,3163;0,10800;3163,18437;10800,21600;18437,18437;21600,10800;18437,3163");
			this.loadTextRect("3163,3163,18437,18437");
		};
		CVmlGeometryData.prototype.fillCFlowChartDecisionType = function() {
			this.concentricFill = true;

			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m10800,l,10800,10800,21600,21600,10800xe");

			this.loadConnectorsList("Rectangle");
			this.loadTextRect("5400,5400,16200,16200");
		};
		CVmlGeometryData.prototype.fillCFlowChartDisplayType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m17955,v862,282,1877,1410,2477,3045c21035,5357,21372,7895,21597,10827v-225,2763,-562,5300,-1165,7613c19832,20132,18817,21260,17955,21597r-14388,l,10827,3567,xe");
			this.loadConnectorsList("Rectangle");
			this.loadTextRect("3567,0,17955,21600");
		};
		CVmlGeometryData.prototype.fillCFlowChartDelayType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m10800,qx21600,10800,10800,21600l,21600,,xe");
			this.loadConnectorsList("Rectangle");
			this.loadTextRect("0,3163,18437,18437");
		};
		CVmlGeometryData.prototype.fillCFlowChartDocumentType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,20172v945,400,1887,628,2795,913c3587,21312,4342,21370,5060,21597v2037,,2567,-227,3095,-285c8722,21197,9325,20970,9855,20800v490,-228,945,-400,1472,-740c11817,19887,12347,19660,12875,19375v567,-228,1095,-513,1700,-740c15177,18462,15782,18122,16537,17950v718,-113,1398,-398,2228,-513c19635,17437,20577,17322,21597,17322l21597,,,xe");

			this.loadConnectorsList("10800,0;0,10800;10800,20400;21600,10800");
			this.loadTextRect("0,0,21600,17322");
		};
		CVmlGeometryData.prototype.fillCFlowChartExtractType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m10800,l21600,21600,,21600xe");
			this.loadConnectorsList("10800,0;5400,10800;10800,21600;16200,10800");
			this.loadTextRect("5400,10800,16200,21600");
		};
		CVmlGeometryData.prototype.fillCFlowChartInputOutputType = function() {
			this.concentricFill = true;

			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m4321,l21600,,17204,21600,,21600xe");

			this.loadConnectorsList("2961,0;10800,0;2161,10800;8602,21600;10800,21600;19402,10800");
			this.loadTextRect("4321,0,17204,21600");
		};
		CVmlGeometryData.prototype.fillCFlowChartInternalStorageType = function() {
			this.concentricFill = true;

			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem4236,nfl4236,21600em,4236nfl21600,4236e");

			this.loadConnectorsList("Rectangle");
			this.loadTextRect("4236,4236,21600,21600");
		};
		CVmlGeometryData.prototype.fillCFlowChartMagneticDiskType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m10800,qx,3391l,18209qy10800,21600,21600,18209l21600,3391qy10800,xem,3391nfqy10800,6782,21600,3391e");
			this.loadConnectorsList("10800,6782;10800,0;0,10800;10800,21600;21600,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("0,6782,21600,18209");
		};
		CVmlGeometryData.prototype.fillCFlowChartMagneticDrumType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m21600,10800qy18019,21600l3581,21600qx,10800,3581,l18019,qx21600,10800xem18019,21600nfqx14438,10800,18019,e");
			this.loadConnectorsList("10800,0;0,10800;10800,21600;14438,10800;21600,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);
			this.addConnectorAngle(0);

			this.loadTextRect("3581,0,14438,21600");
		};
		CVmlGeometryData.prototype.fillCFlowChartMagneticTapeType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("ar,,21600,21600,18685,18165,10677,21597l20990,21597r,-3432xe");
			this.loadConnectorsList("Rectangle");
			this.loadTextRect("3163,3163,18437,18437");
		};
		CVmlGeometryData.prototype.fillCFlowChartManualInputType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,4292l21600,r,21600l,21600xe");

			this.loadConnectorsList("10800,2146;0,10800;10800,21600;21600,10800");
			this.loadTextRect("0,4291,21600,21600");
		};
		CVmlGeometryData.prototype.fillCFlowChartManualOperationType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l21600,,17240,21600r-12880,xe");

			this.loadConnectorsList("10800,0;2180,10800;10800,21600;19420,10800");
			this.loadTextRect("4321,0,17204,21600");
		};
		CVmlGeometryData.prototype.fillCFlowChartMergeType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m,l21600,,10800,21600xe");
			this.loadConnectorsList("10800,0;5400,10800;10800,21600;16200,10800");
			this.loadTextRect("5400,0,16200,10800");
		};
		CVmlGeometryData.prototype.fillCFlowChartMultidocumentType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,20465v810,317,1620,452,2397,725c3077,21325,3790,21417,4405,21597v1620,,2202,-180,2657,-272c7580,21280,8002,21010,8455,20917v422,-135,810,-405,1327,-542c10205,20150,10657,19967,11080,19742v517,-182,970,-407,1425,-590c13087,19017,13605,18745,14255,18610v615,-180,1262,-318,1942,-408c16975,18202,17785,18022,18595,18022r,-1670l19192,16252r808,l20000,14467r722,-75l21597,14392,21597,,2972,r,1815l1532,1815r,1860l,3675,,20465xem1532,3675nfl18595,3675r,12677em2972,1815nfl20000,1815r,12652e");

			this.loadConnectorsList("10800,0;0,10800;10800,19890;21600,10800");
			this.loadTextRect("0,3675,18595,18022");
		};
		CVmlGeometryData.prototype.fillCFlowChartOffpageConnectorType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m,l21600,r,17255l10800,21600,,17255xe");
			this.loadConnectorsList("Rectangle");
			this.loadTextRect("0,0,21600,17255");
		};
		CVmlGeometryData.prototype.fillCFlowChartOnlineStorageType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m3600,21597c2662,21202,1837,20075,1087,18440,487,16240,75,13590,,10770,75,8007,487,5412,1087,3045,1837,1465,2662,337,3600,l21597,v-937,337,-1687,1465,-2512,3045c18485,5412,18072,8007,17997,10770v75,2820,488,5470,1088,7670c19910,20075,20660,21202,21597,21597xe");
			this.loadConnectorsList("10800,0;0,10800;10800,21600;17997,10800");
			this.loadTextRect("3600,0,17997,21600");
		};
		CVmlGeometryData.prototype.fillCFlowChartOrType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m10800,qx,10800,10800,21600,21600,10800,10800,xem,10800nfl21600,10800em10800,nfl10800,21600e");
			this.loadConnectorsList("10800,0;3163,3163;0,10800;3163,18437;10800,21600;18437,18437;21600,10800;18437,3163");
			this.loadTextRect("3163,3163,18437,18437");
		};
		CVmlGeometryData.prototype.fillCFlowChartPredefinedProcessType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xem2610,nfl2610,21600em18990,nfl18990,21600e");

			this.loadConnectorsList("Rectangle");
			this.loadTextRect("2610,0,18990,21600");
		};
		CVmlGeometryData.prototype.fillCFlowChartPreparationType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m4353,l17214,r4386,10800l17214,21600r-12861,l,10800xe");

			this.loadConnectorsList("Rectangle");
			this.loadTextRect("4353,0,17214,21600");
		};
		CVmlGeometryData.prototype.fillCFlowChartProcessType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,l21600,xe");
			this.loadConnectorsList("Rectangle");
		};
		CVmlGeometryData.prototype.fillCFlowChartPunchedCardType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m4321,l21600,r,21600l,21600,,4338xe");
			this.loadConnectorsList("Rectangle");
			this.loadTextRect("0,4321,21600,21600");
		};
		CVmlGeometryData.prototype.fillCFlowChartPunchedTapeType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m21597,19450v-225,-558,-750,-1073,-1650,-1545c18897,17605,17585,17347,16197,17260v-1500,87,-2700,345,-3787,645c11472,18377,10910,18892,10800,19450v-188,515,-750,1075,-1613,1460c8100,21210,6825,21425,5400,21597,3937,21425,2700,21210,1612,20910,675,20525,150,19965,,19450l,2147v150,558,675,1073,1612,1460c2700,3950,3937,4165,5400,4337,6825,4165,8100,3950,9187,3607v863,-387,1425,-902,1613,-1460c10910,1632,11472,1072,12410,600,13497,300,14697,85,16197,v1388,85,2700,300,3750,600c20847,1072,21372,1632,21597,2147xe");
			this.loadConnectorsList("10800,2147;0,10800;10800,19450;21600,10800");
			this.loadTextRect("0,4337,21600,17260");
		};
		CVmlGeometryData.prototype.fillCFlowChartSortType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m10800,l,10800,10800,21600,21600,10800xem,10800nfl21600,10800e");
			this.loadConnectorsList("Rectangle");
			this.loadTextRect("5400,5400,16200,16200");
		};
		CVmlGeometryData.prototype.fillCFlowChartSummingJunctionType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m10800,qx,10800,10800,21600,21600,10800,10800,xem3163,3163nfl18437,18437em3163,18437nfl18437,3163e");
			this.loadConnectorsList("10800,0;3163,3163;0,10800;3163,18437;10800,21600;18437,18437;21600,10800;18437,3163");
			this.loadTextRect("3163,3163,18437,18437");
		};
		CVmlGeometryData.prototype.fillCFlowChartTerminatorType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinRound;

			this.loadPath("m3475,qx,10800,3475,21600l18125,21600qx21600,10800,18125,xe");

			this.loadConnectorsList("Rectangle");
			this.loadTextRect("1018,3163,20582,18437");
		};
		CVmlGeometryData.prototype.fillCFoldedCornerType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600@0,21600,21600@0,21600,xem@0,21600nfl@3@5c@7@9@11@13,21600@0e");

			this.addGuide("val #0");
			this.addGuide("sum 21600 0 @0");
			this.addGuide("prod @1 8481 32768");
			this.addGuide("sum @2 @0 0");
			this.addGuide("prod @1 1117 32768");
			this.addGuide("sum @4 @0 0");
			this.addGuide("prod @1 11764 32768");
			this.addGuide("sum @6 @0 0");
			this.addGuide("prod @1 6144 32768");
			this.addGuide("sum @8 @0 0");
			this.addGuide("prod @1 20480 32768");
			this.addGuide("sum @10 @0 0");
			this.addGuide("prod @1 6144 32768");
			this.addGuide("sum @12 @0 0");

			this.addAdjustment(18900);

			this.loadConnectorsList("Rectangle");
			this.loadTextRect("0,0,21600,@13");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,bottomRight");
			oHandle1.xrange = ("10800,21600");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCHeartType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m10860,2187c10451,1746,9529,1018,9015,730,7865,152,6685,,5415,,4175,152,2995,575,1967,1305,1150,2187,575,3222,242,4220,,5410,242,6560,575,7597l10860,21600,20995,7597v485,-1037,605,-2187,485,-3377c21115,3222,20420,2187,19632,1305,18575,575,17425,152,16275,,15005,,13735,152,12705,730v-529,288,-1451,1016,-1845,1457xe");

			this.loadConnectorsList("10860,2187;2928,10800;10860,21600;18672,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("5037,2277,16557,13677");
		};
		CVmlGeometryData.prototype.fillCHexagonType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m@0,l,10800@0,21600@1,21600,21600,10800@1,xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod @0 2929 10000");
			this.addGuide("sum width 0 @3");
			this.addGuide("sum height 0 @3");

			this.addAdjustment(5400);

			this.loadConnectorsList("Rectangle");
			this.loadTextRect("1800,1800,19800,19800;3600,3600,18000,18000;6300,6300,15300,15300");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCHomePlateType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m@0,l,,,21600@0,21600,21600,10800xe");

			this.addGuide("val #0");
			this.addGuide("prod #0 1 2");

			this.addAdjustment(16200);

			this.loadConnectorsList("@1,0;0,10800;@1,21600;21600,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("0,0,10800,21600;0,0,16200,21600;0,0,21600,21600");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("0,21600");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCIrregularSealOneType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m10800,5800l8352,2295,7312,6320,370,2295,4627,7617,,8615r3722,3160l135,14587r5532,-650l4762,17617,7715,15627r770,5973l10532,14935r2715,4802l14020,14457r4125,3638l16837,12942r4763,348l17607,10475,21097,8137,16702,7315,18380,4457r-4225,868l14522,xe");

			this.loadConnectorsList("14522,0;0,8615;8485,21600;21600,13290");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("4627,6320,16702,13937");
		};
		CVmlGeometryData.prototype.fillCIrregularSealTwo = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m11462,4342l9722,1887,8550,6382,4502,3625r870,4192l1172,8270r2763,3322l,12877r3330,2493l1285,17825r3520,415l4917,21600,7527,18125r1173,1587l9872,17370r1740,1472l12180,15935r2762,1435l14640,14350r4237,1282l16380,12310r1890,-1020l16985,9402,21600,6645,16380,6532,18007,3172,14525,5777,14790,xe");

			this.loadConnectorsList("9722,1887;0,12877;11612,18842;21600,6645");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("5372,6382,14640,15935");
		};
		CVmlGeometryData.prototype.fillCIsoscelesTriangleType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m@0,l,21600r21600,xe");

			this.addGuide("val #0");
			this.addGuide("prod #0 1 2");
			this.addGuide("sum @1 10800 0");

			this.addAdjustment(10800);

			this.loadConnectorsList("@0,0;@1,10800;0,21600;10800,21600;21600,21600;@2,10800");

			this.loadTextRect("0,10800,10800,18000;5400,10800,16200,18000;10800,10800,21600,18000;0,7200,7200,21600;7200,7200,14400,21600;14400,7200,21600,21600");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("0,21600");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCLeftArrowCalloutType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m@0,l@0@3@2@3@2@1,,10800@2@4@2@5@0@5@0,21600,21600,21600,21600,xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("sum 21600 0 #3");
			this.addGuide("sum #0 21600 0");

			this.addAdjustment(7200);
			this.addAdjustment(5400);
			this.addAdjustment(3600);
			this.addAdjustment(8100);

			this.loadConnectorsList("@7,0;0,10800;@7,21600;21600,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@0,0,21600,21600");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("@2,21600");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("topLeft,#1");
			oHandle2.yrange = ("0,@3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#2,#3");
			oHandle3.xrange = ("0,@0");
			oHandle3.yrange = ("@1,10800");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCLeftArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m@0,l@0@1,0@1,0@2@0@2@0,21600,21600,10800xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("sum height 0 #1");
			this.addGuide("sum 10800 0 #1");
			this.addGuide("sum width 0 #0");
			this.addGuide("prod @4 @3 10800");
			this.addGuide("sum width 0 @5");

			this.addAdjustment(16200);
			this.addAdjustment(5400);

			this.loadConnectorsList("@0,0;0,10800;@0,21600;21600,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("0,@1,@6,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			oHandle1.xrange = ("0,21600");
			oHandle1.yrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCLeftBraceType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			//Endcaps: Flat

			this.loadPath("m21600,qx10800@0l10800@2qy0@11,10800@3l10800@1qy21600,21600e");

			this.addGuide("val #0");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("sum #1 0 #0");
			this.addGuide("sum #1 #0 0");
			this.addGuide("prod #0 9598 32768");
			this.addGuide("sum 21600 0 @4");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("min #1 @6");
			this.addGuide("prod @7 1 2");
			this.addGuide("prod #0 2 1");
			this.addGuide("sum 21600 0 @9");
			this.addGuide("val #1");

			this.addAdjustment(1800);
			this.addAdjustment(10800);

			this.loadConnectorsList("21600,0;0,10800;21600,21600");
			this.loadTextRect("13963,@4,21600,@5");

			let oHandle1 = new CVmlHandle();
			let oHandle2 = new CVmlHandle();

			oHandle1.position = ("center,#0");
			oHandle1.yrange = ("0,@8");
			oHandle2.position = ("topLeft,#1");
			oHandle2.yrange = ("@9,@10");

			this.addHandle(oHandle1);
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCLeftBracketType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinRound;
			//Endcaps: Flat

			this.loadPath("m21600,qx0@0l0@1qy21600,21600e");

			this.addGuide("val #0");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("prod #0 9598 32768");
			this.addGuide("sum 21600 0 @2");

			this.addAdjustment(1800);

			this.loadConnectorsList("21600,0;0,10800;21600,21600");
			this.loadTextRect("6326,@2,21600,@3");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("topLeft,#0");
			oHandle1.yrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCLeftRightArrowCalloutType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m@0,l@0@3@2@3@2@1,,10800@2@4@2@5@0@5@0,21600@8,21600@8@5@9@5@9@4,21600,10800@9@1@9@3@8@3@8,xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("sum 21600 0 #3");
			this.addGuide("sum #0 21600 0");
			this.addGuide("prod @6 1 2");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("sum 21600 0 #2");

			this.addAdjustment(5400);
			this.addAdjustment(5400);
			this.addAdjustment(2700);
			this.addAdjustment(8100);

			this.loadConnectorsList("10800,0;0,10800;10800,21600;21600,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@0,0,@8,21600");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("@2,10800");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("topLeft,#1");
			oHandle2.yrange = ("0,@3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#2,#3");
			oHandle3.xrange = ("0,@0");
			oHandle3.yrange = ("@1,10800");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCLeftRightArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,10800l@0,21600@0@3@2@3@2,21600,21600,10800@2,0@2@1@0@1@0,xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("prod #0 #1 10800");
			this.addGuide("sum #0 0 @4");
			this.addGuide("sum 21600 0 @5");

			this.addAdjustment(4320);
			this.addAdjustment(5400);

			this.loadConnectorsList("@2,0;10800,@1;@0,0;0,10800;@0,21600;10800,@3;@2,21600;21600,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(270);
			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(90);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@5,@1,@6,@3");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			oHandle1.xrange = ("0,10800");
			oHandle1.yrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCLeftRightUpArrow = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m10800,l@0@2@1@2@1@6@7@6@7@5,0@8@7,21600@7@9@10@9@10,21600,21600@8@10@5@10@6@4@6@4@2@3@2xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("prod @0 21600 @3");
			this.addGuide("prod @1 21600 @3");
			this.addGuide("prod @2 @3 21600");
			this.addGuide("prod 10800 21600 @3");
			this.addGuide("prod @4 21600 @3");
			this.addGuide("sum 21600 0 @7");
			this.addGuide("sum @5 0 @8");
			this.addGuide("sum @6 0 @8");
			this.addGuide("prod @12 @7 @11");
			this.addGuide("sum 21600 0 @13");
			this.addGuide("sum @0 0 10800");
			this.addGuide("sum @1 0 10800");
			this.addGuide("prod @2 @16 @15");

			this.addAdjustment(6480);
			this.addAdjustment(8640);
			this.addAdjustment(6171);

			this.loadConnectorsList("10800,0;0,@8;10800,@9;21600,@8");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@13,@6,@14,@9;@1,@17,@4,@9");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("@2,@1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#1,#2");
			oHandle2.xrange = ("@0,10800");
			oHandle2.yrange = ("0,@5");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCLeftUpArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m@4,l@0@2@5@2@5@5@2@5@2@0,0@4@2,21600@2@1@1@1@1@2,21600@2xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("prod #0 1 2");
			this.addGuide("sum @3 10800 0");
			this.addGuide("sum 21600 #0 #1");
			this.addGuide("sum #1 #2 0");
			this.addGuide("prod @6 1 2");
			this.addGuide("prod #1 2 1");
			this.addGuide("sum @8 0 21600");
			this.addGuide("sum @5 0 @4");
			this.addGuide("sum #0 0 @4");
			this.addGuide("prod @2 @10 @11");

			this.addAdjustment(9257);
			this.addAdjustment(18514);
			this.addAdjustment(6171);

			this.loadConnectorsList("@4,0;@0,@2;@2,@0;0,@4;@2,21600;@7,@1;@1,@7;21600,@2");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);
			this.addConnectorAngle(0);

			this.loadTextRect("@12,@5,@1,@1;@5,@12,@1,@1");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("@2,@9");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#1,#2");
			oHandle2.xrange = ("@4,21600");
			oHandle2.yrange = ("0,@0");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCLightningBoltType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m8472,l,3890,7602,8382,5022,9705r7200,4192l10012,14915r11588,6685l14767,12877r1810,-870l11050,6797r1810,-717xe");

			this.loadConnectorsList("8472,0;0,3890;5022,9705;10012,14915;21600,21600;16577,12007;12860,6080");

			this.addConnectorAngle(270);
			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);
			this.addConnectorAngle(0);

			this.loadTextRect("8757,7437,13917,14277");
		};
		CVmlGeometryData.prototype.fillCMoonType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m21600,qx,10800,21600,21600wa@0@10@6@11,21600,21600,21600,xe");

			this.addGuide("val #0");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("prod #0 #0 @1");
			this.addGuide("prod 21600 21600 @1");
			this.addGuide("prod @3 2 1");
			this.addGuide("sum @4 0 @2");
			this.addGuide("sum @5 0 #0");
			this.addGuide("prod @5 1 2");
			this.addGuide("sum @7 0 #0");
			this.addGuide("prod @8 1 2");
			this.addGuide("sum 10800 0 @9");
			this.addGuide("sum @9 10800 0");
			this.addGuide("prod #0 9598 32768");
			this.addGuide("sum 21600 0 @12");
			this.addGuide("ellipse @13 21600 10800");
			this.addGuide("sum 10800 0 @14");
			this.addGuide("sum @14 10800 0");

			this.addAdjustment(10800);

			this.loadConnectorsList("21600,0;0,10800;21600,21600;@0,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@12,@15,@0,@16");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,center");
			oHandle1.xrange = ("0,18900");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCNoSmokingType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m,10800qy10800,,21600,10800,10800,21600,,10800xar@0@0@16@16@12@14@15@13xar@0@0@16@16@13@15@14@12xe");

			this.addGuide("val #0");
			this.addGuide("prod @0 2 1");
			this.addGuide("sum 21600 0 @1");
			this.addGuide("prod @2 @2 1");
			this.addGuide("prod @0 @0 1");
			this.addGuide("sum @3 0 @4");
			this.addGuide("prod @5 1 8");
			this.addGuide("sqrt @6");
			this.addGuide("prod @4 1 8");
			this.addGuide("sqrt @8");
			this.addGuide("sum @7 @9 0");
			this.addGuide("sum @7 0 @9");
			this.addGuide("sum @10 10800 0");
			this.addGuide("sum 10800 0 @10");
			this.addGuide("sum @11 10800 0");
			this.addGuide("sum 10800 0 @11");
			this.addGuide("sum 21600 0 @0");

			this.addAdjustment(2700);

			this.loadConnectorsList("10800,0;3163,3163;0,10800;3163,18437;10800,21600;18437,18437;21600,10800;18437,3163");

			this.loadTextRect("3163,3163,18437,18437");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,center");
			oHandle1.xrange = ("0,7200");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCNotchedRightArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m@0,l@0@1,0@1@5,10800,0@2@0@2@0,21600,21600,10800xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("sum height 0 #1");
			this.addGuide("sum 10800 0 #1");
			this.addGuide("sum width 0 #0");
			this.addGuide("prod @4 @3 10800");
			this.addGuide("sum width 0 @5");

			this.addAdjustment(16200);
			this.addAdjustment(5400);

			this.loadConnectorsList("@0,0;@5,10800;@0,21600;21600,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@5,@1,@6,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			oHandle1.xrange = ("0,21600");
			oHandle1.yrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCOctagonType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m@0,l0@0,0@2@0,21600@1,21600,21600@2,21600@0@1,xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod @0 2929 10000");
			this.addGuide("sum width 0 @3");
			this.addGuide("sum height 0 @3");
			this.addGuide("val width");
			this.addGuide("val height");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");

			this.addAdjustment(6326);

			this.loadConnectorsList("@8,0;0,@9;@8,@7;@6,@9");
			this.loadTextRect("0,0,21600,21600;2700,2700,18900,18900;5400,5400,16200,16200");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCParallelogramType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m@0,l,21600@1,21600,21600,xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("prod #0 1 2");
			this.addGuide("sum width 0 @2");
			this.addGuide("mid #0 width");
			this.addGuide("mid @1 0");
			this.addGuide("prod height width #0");
			this.addGuide("prod @6 1 2");
			this.addGuide("sum height 0 @7");
			this.addGuide("prod width 1 2");
			this.addGuide("sum #0 0 @9");
			this.addGuide("if @10 @8 0");
			this.addGuide("if @10 @7 height");

			this.addAdjustment(5400);

			this.loadConnectorsList("@4,0;10800,@11;@3,10800;@5,21600;10800,@12;@2,10800");
			this.loadTextRect("1800,1800,19800,19800;8100,8100,13500,13500;10800,10800,10800,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("0,21600");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCPentagonType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m10800,l,8259,4200,21600r13200,l21600,8259xe");
			this.loadConnectorsList("10800,0;0,8259;4200,21600;10800,21600;17400,21600;21600,8259");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(90);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("4200,5077,17400,21600");
		};
		CVmlGeometryData.prototype.fillCPlaqueType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m@0,qy0@0l0@2qx@0,21600l@1,21600qy21600@2l21600@0qx@1,xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod @0 7071 10000");
			this.addGuide("sum width 0 @3");
			this.addGuide("sum height 0 @3");
			this.addGuide("val width");
			this.addGuide("val height");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");

			this.addAdjustment(3600);

			this.loadConnectorsList("@8,0;0,@9;@8,@7;@6,@9");
			this.loadTextRect("@3,@3,@4,@5");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCPlusType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m@0,l@0@0,0@0,0@2@0@2@0,21600@1,21600@1@2,21600@2,21600@0@1@0@1,xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("sum height 0 #0");
			this.addGuide("prod @0 2929 10000");
			this.addGuide("sum width 0 @3");
			this.addGuide("sum height 0 @3");
			this.addGuide("val width");
			this.addGuide("val height");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");

			this.addAdjustment(5400);

			this.loadConnectorsList("@8,0;0,@9;@8,@7;@6,@9");
			this.loadTextRect("0,0,21600,21600;5400,5400,16200,16200;10800,10800,10800,10800");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.switchHandle = ("true");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCQuadArrowCalloutType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m@0@0l@3@0@3@2@1@2,10800,0@4@2@5@2@5@0@8@0@8@3@9@3@9@1,21600,10800@9@4@9@5@8@5@8@8@5@8@5@9@4@9,10800,21600@1@9@3@9@3@8@0@8@0@5@2@5@2@4,,10800@2@1@2@3@0@3xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("sum 21600 0 #3");
			this.addGuide("sum #0 21600 0");
			this.addGuide("prod @6 1 2");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("sum 21600 0 #2");

			this.addAdjustment(5400);
			this.addAdjustment(8100);
			this.addAdjustment(2700);
			this.addAdjustment(9450);

			this.loadConnectorsList("Rectangle");

			this.loadTextRect("@0,@0,@8,@8");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("topLeft,#0");
			oHandle1.yrange = ("@2,@1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#1,topLeft");
			oHandle2.xrange = ("@0,@3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#3,#2");
			oHandle3.xrange = ("@1,10800");
			oHandle3.yrange = ("0,@0");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCQuadArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m10800,l@0@2@1@2@1@1@2@1@2@0,,10800@2@3@2@4@1@4@1@5@0@5,10800,21600@3@5@4@5@4@4@5@4@5@3,21600,10800@5@0@5@1@4@1@4@2@3@2xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("sum 21600 0 #2");
			this.addGuide("sum #0 0 10800");
			this.addGuide("sum #1 0 10800");
			this.addGuide("prod @7 #2 @6");
			this.addGuide("sum 21600 0 @8");

			this.addAdjustment(6480);
			this.addAdjustment(8640);
			this.addAdjustment(4320);

			this.loadConnectorsList("Rectangle");

			this.loadTextRect("@8,@1,@9,@4;@1,@8,@4,@9");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("@2,@1");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#1,#2");
			oHandle2.xrange = ("@0,10800");
			oHandle2.yrange = ("0,@0");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCRectangleType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m,l,21600r21600,l21600,xe");
		};
		CVmlGeometryData.prototype.fillCRibbonDownType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l@3,qx@4@11l@4@10@5@10@5@11qy@6,l@21,0@19@15@21@16@9@16@9@17qy@8@22l@1@22qx@0@17l@0@16,0@16,2700@15xem@4@11nfqy@3@12l@1@12qx@0@13@1@10l@4@10em@5@11nfqy@6@12l@8@12qx@9@13@8@10l@5@10em@0@13nfl@0@16em@9@13nfl@9@16e");

			this.addGuide("val #0");
			this.addGuide("sum @0 675 0");
			this.addGuide("sum @1 675 0");
			this.addGuide("sum @2 675 0");
			this.addGuide("sum @3 675 0");
			this.addGuide("sum width 0 @4");
			this.addGuide("sum width 0 @3");
			this.addGuide("sum width 0 @2");
			this.addGuide("sum width 0 @1");
			this.addGuide("sum width 0 @0");
			this.addGuide("val #1");
			this.addGuide("prod @10 1 4");
			this.addGuide("prod @11 2 1");
			this.addGuide("prod @11 3 1");
			this.addGuide("prod height 1 2");
			this.addGuide("sum @14 0 @12");
			this.addGuide("sum height 0 @10");
			this.addGuide("sum height 0 @11");
			this.addGuide("prod width 1 2");
			this.addGuide("sum width 0 2700");
			this.addGuide("sum @18 0 2700");
			this.addGuide("val width");
			this.addGuide("val height");

			this.addAdjustment(5400);
			this.addAdjustment(2700);

			this.loadConnectorsList("@18,@10;2700,@15;@18,21600;@19,@15");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@0,@10,@9,21600");

			let oHandle1 = new CVmlHandle();
			let oHandle2 = new CVmlHandle();

			oHandle1.position = ("#0,bottomRight");
			oHandle1.xrange = ("2700,8100");
			oHandle2.position = ("center,#1");
			oHandle2.yrange = ("0,7200");

			this.addHandle(oHandle1);
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCRibbonUpType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m0@29l@3@29qx@4@19l@4@10@5@10@5@19qy@6@29l@28@29@26@22@28@23@9@23@9@24qy@8,l@1,qx@0@24l@0@23,0@23,2700@22xem@4@19nfqy@3@20l@1@20qx@0@21@1@10l@4@10em@5@19nfqy@6@20l@8@20qx@9@21@8@10l@5@10em@0@21nfl@0@23em@9@21nfl@9@23e");

			this.addGuide("val #0");
			this.addGuide("sum @0 675 0");
			this.addGuide("sum @1 675 0");
			this.addGuide("sum @2 675 0");
			this.addGuide("sum @3 675 0");
			this.addGuide("sum width 0 @4");
			this.addGuide("sum width 0 @3");
			this.addGuide("sum width 0 @2");
			this.addGuide("sum width 0 @1");
			this.addGuide("sum width 0 @0");
			this.addGuide("val #1");
			this.addGuide("prod @10 1 4");
			this.addGuide("prod @10 1 2");
			this.addGuide("prod @10 3 4");
			this.addGuide("prod height 3 4");
			this.addGuide("prod height 1 2");
			this.addGuide("prod height 1 4");
			this.addGuide("prod height 3 2");
			this.addGuide("prod height 2 3");
			this.addGuide("sum @11 @14 0");
			this.addGuide("sum @12 @15 0");
			this.addGuide("sum @13 @16 0");
			this.addGuide("sum @17 0 @20");
			this.addGuide("sum height 0 @10");
			this.addGuide("sum height 0 @19");
			this.addGuide("prod width 1 2");
			this.addGuide("sum width 0 2700");
			this.addGuide("sum @25 0 2700");
			this.addGuide("val width");
			this.addGuide("val height");

			this.addAdjustment(5400);
			this.addAdjustment(18900);

			this.loadConnectorsList("@25,0;2700,@22;@25,@10;@26,@22");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@0,0,@9,@10");

			let oHandle1 = new CVmlHandle();
			let oHandle2 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("2700,8100");
			oHandle2.position = ("center,#1");
			oHandle2.yrange = ("14400,21600");
			this.addHandle(oHandle1);
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCRightArrowCalloutType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m,l,21600@0,21600@0@5@2@5@2@4,21600,10800@2@1@2@3@0@3@0,x");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("sum 21600 0 #3");
			this.addGuide("prod #0 1 2");

			this.addAdjustment(14400);
			this.addAdjustment(5400);
			this.addAdjustment(18000);
			this.addAdjustment(8100);

			this.loadConnectorsList("@6,0;0,10800;@6,21600;21600,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("0,0,@0,21600");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("0,@2");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("bottomRight,#1");
			oHandle2.yrange = ("0,@3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#2,#3");
			oHandle3.xrange = ("@0,21600");
			oHandle3.yrange = ("@1,10800");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCRightArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m@0,l@0@1,0@1,0@2@0@2@0,21600,21600,10800xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("sum height 0 #1");
			this.addGuide("sum 10800 0 #1");
			this.addGuide("sum width 0 #0");
			this.addGuide("prod @4 @3 10800");
			this.addGuide("sum width 0 @5");

			this.addAdjustment(16200);
			this.addAdjustment(5400);

			this.loadConnectorsList("@0,0;0,10800;@0,21600;21600,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("0,@1,@6,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			oHandle1.xrange = ("0,21600");
			oHandle1.yrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCRightBracetype = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,qx10800@0l10800@2qy21600@11,10800@3l10800@1qy,21600e");

			this.addGuide("val #0");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("sum #1 0 #0");
			this.addGuide("sum #1 #0 0");
			this.addGuide("prod #0 9598 32768");
			this.addGuide("sum 21600 0 @4");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("min #1 @6");
			this.addGuide("prod @7 1 2");
			this.addGuide("prod #0 2 1");
			this.addGuide("sum 21600 0 @9");
			this.addGuide("val #1");

			this.addAdjustment(1800);
			this.addAdjustment(10800);

			this.loadConnectorsList("0,0;21600,@11;0,21600");
			this.loadTextRect("0,@4,7637,@5");

			let oHandle1 = new CVmlHandle();
			let oHandle2 = new CVmlHandle();
			oHandle1.position = ("center,#0");
			oHandle1.yrange = ("0,@8");
			oHandle2.position = ("bottomRight,#1");
			oHandle2.yrange = ("@9,@10");
			this.addHandle(oHandle1);
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCRightBracketType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinRound;
			//Endcaps: Flat

			this.loadPath("m,qx21600@0l21600@1qy,21600e");

			this.addGuide("val #0");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("prod #0 9598 32768");
			this.addGuide("sum 21600 0 @2");

			this.addAdjustment(1800);

			this.loadConnectorsList("0,0;0,21600;21600,10800");
			this.loadTextRect("0,@2,15274,@3");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("bottomRight,#0");
			oHandle1.yrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCRightTriangleType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m,l,21600r21600,xe");

			this.loadConnectorsList("0,0;0,10800;0,21600;10800,21600;21600,21600;10800,10800");
			this.loadTextRect("1800,12600,12600,19800");
		};
		CVmlGeometryData.prototype.fillCRoundedRectangleType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m0@0qy@0,0l@1,0qx21600@0l21600@1qy@1,21600l@0,21600qx0@1xe");

			this.addGuide("val #0");
			this.addGuide("sum 21600 0 #0");

			this.addAdjustment(5400);

			this.loadTextRect("0,0,21600,21600");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,topLeft");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCStar16Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m21600,10800l@5@10,20777,6667@7@12,18436,3163@8@11,14932,822@6@9,10800,0@10@9,6667,822@12@11,3163,3163@11@12,822,6667@9@10,,10800@9@6,822,14932@11@8,3163,18436@12@7,6667,20777@10@5,10800,21600@6@5,14932,20777@8@7,18436,18436@7@8,20777,14932@5@6xe");

			this.addGuide("sum 10800 0 #0");
			this.addGuide("prod @0 32138 32768");
			this.addGuide("prod @0 6393 32768");
			this.addGuide("prod @0 27246 32768");
			this.addGuide("prod @0 18205 32768");
			this.addGuide("sum @1 10800 0");
			this.addGuide("sum @2 10800 0");
			this.addGuide("sum @3 10800 0");
			this.addGuide("sum @4 10800 0");
			this.addGuide("sum 10800 0 @1");
			this.addGuide("sum 10800 0 @2");
			this.addGuide("sum 10800 0 @3");
			this.addGuide("sum 10800 0 @4");
			this.addGuide("prod @0 23170 32768");
			this.addGuide("sum @13 10800 0");
			this.addGuide("sum 10800 0 @13");

			this.addAdjustment(2700);

			this.loadConnectorsList("Rectangle");
			this.loadTextRect("@15,@15,@14,@14");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,center");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCStar24Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m21600,10800l@7@14,21232,8005@9@16,20153,5400@11@18,18437,3163@12@17,16200,1447@10@15,13595,368@8@13,10800,0@14@13,8005,368@16@15,5400,1447@18@17,3163,3163@17@18,1447,5400@15@16,368,8005@13@14,,10800@13@8,368,13595@15@10,1447,16200@17@12,3163,18437@18@11,5400,20153@16@9,8005,21232@14@7,10800,21600@8@7,13595,21232@10@9,16200,20153@12@11,18437,18437@11@12,20153,16200@9@10,21232,13595@7@8xe");

			this.addGuide("sum 10800 0 #0");
			this.addGuide("prod @0 32488 32768");
			this.addGuide("prod @0 4277 32768");
			this.addGuide("prod @0 30274 32768");
			this.addGuide("prod @0 12540 32768");
			this.addGuide("prod @0 25997 32768");
			this.addGuide("prod @0 19948 32768");
			this.addGuide("sum @1 10800 0");
			this.addGuide("sum @2 10800 0");
			this.addGuide("sum @3 10800 0");
			this.addGuide("sum @4 10800 0");
			this.addGuide("sum @5 10800 0");
			this.addGuide("sum @6 10800 0");
			this.addGuide("sum 10800 0 @1");
			this.addGuide("sum 10800 0 @2");
			this.addGuide("sum 10800 0 @3");
			this.addGuide("sum 10800 0 @4");
			this.addGuide("sum 10800 0 @5");
			this.addGuide("sum 10800 0 @6");
			this.addGuide("prod @0 23170 32768");
			this.addGuide("sum @19 10800 0");
			this.addGuide("sum 10800 0 @19");

			this.addAdjustment(2700);

			this.loadConnectorsList("Rectangle");
			this.loadTextRect("@21,@21,@20,@20");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,center");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCStar32Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m21600,10800l@9@18,21392,8693@11@20,20777,6667@13@22,19780,4800@15@24,18436,3163@16@23,16800,1820@14@21,14932,822@12@19,12907,208@10@17,10800,0@18@17,8693,208@20@19,6667,822@22@21,4800,1820@24@23,3163,3163@23@24,1820,4800@21@22,822,6667@19@20,208,8693@17@18,,10800@17@10,208,12907@19@12,822,14932@21@14,1820,16800@23@16,3163,18436@24@15,4800,19780@22@13,6667,20777@20@11,8693,21392@18@9,10800,21600@10@9,12907,21392@12@11,14932,20777@14@13,16800,19780@16@15,18436,18436@15@16,19780,16800@13@14,20777,14932@11@12,21392,12907@9@10xe");

			this.addGuide("sum 10800 0 #0");
			this.addGuide("prod @0 32610 32768");
			this.addGuide("prod @0 3212 32768");
			this.addGuide("prod @0 31357 32768");
			this.addGuide("prod @0 9512 32768");
			this.addGuide("prod @0 28899 32768");
			this.addGuide("prod @0 15447 32768");
			this.addGuide("prod @0 25330 32768");
			this.addGuide("prod @0 20788 32768");
			this.addGuide("sum @1 10800 0");
			this.addGuide("sum @2 10800 0");
			this.addGuide("sum @3 10800 0");
			this.addGuide("sum @4 10800 0");
			this.addGuide("sum @5 10800 0");
			this.addGuide("sum @6 10800 0");
			this.addGuide("sum @7 10800 0");
			this.addGuide("sum @8 10800 0");
			this.addGuide("sum 10800 0 @1");
			this.addGuide("sum 10800 0 @2");
			this.addGuide("sum 10800 0 @3");
			this.addGuide("sum 10800 0 @4");
			this.addGuide("sum 10800 0 @5");
			this.addGuide("sum 10800 0 @6");
			this.addGuide("sum 10800 0 @7");
			this.addGuide("sum 10800 0 @8");
			this.addGuide("prod @0 23170 32768");
			this.addGuide("sum @25 10800 0");
			this.addGuide("sum 10800 0 @25");

			this.addAdjustment(2700);

			this.loadConnectorsList("Rectangle");
			this.loadTextRect("@27,@27,@26,@26");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,center");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCStar4Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m21600,10800l@2@3,10800,0@3@3,,10800@3@2,10800,21600@2@2xe");

			this.addGuide("sum 10800 0 #0");
			this.addGuide("prod @0 23170 32768");
			this.addGuide("sum @1 10800 0");
			this.addGuide("sum 10800 0 @1");

			this.addAdjustment(8100);

			this.loadConnectorsList("Rectangle");
			this.loadTextRect("@3,@3,@2,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,center");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCStar8Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m21600,10800l@3@6,18436,3163@4@5,10800,0@6@5,3163,3163@5@6,,10800@5@4,3163,18436@6@3,10800,21600@4@3,18436,18436@3@4xe");

			this.addGuide("sum 10800 0 #0");
			this.addGuide("prod @0 30274 32768");
			this.addGuide("prod @0 12540 32768");
			this.addGuide("sum @1 10800 0");
			this.addGuide("sum @2 10800 0");
			this.addGuide("sum 10800 0 @1");
			this.addGuide("sum 10800 0 @2");
			this.addGuide("prod @0 23170 32768");
			this.addGuide("sum @7 10800 0");
			this.addGuide("sum 10800 0 @7");

			this.addAdjustment(2538);

			this.loadConnectorsList("Rectangle");
			this.loadTextRect("@9,@9,@8,@8");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,center");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCSmileyFaceType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinRound;

			this.loadPath("m10800,qx,10800,10800,21600,21600,10800,10800,xem7340,6445qx6215,7570,7340,8695,8465,7570,7340,6445xnfem14260,6445qx13135,7570,14260,8695,15385,7570,14260,6445xnfem4960@0c8853@3,12747@3,16640@0nfe");

			this.addGuide("sum 33030 0 #0");
			this.addGuide("prod #0 4 3");
			this.addGuide("prod @0 1 3");
			this.addGuide("sum @1 0 @2");

			this.addAdjustment(17520);

			this.loadConnectorsList("10800,0;3163,3163;0,10800;3163,18437;10800,21600;18437,18437;21600,10800;18437,3163");
			this.loadTextRect("3163,3163,18437,18437");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("center,#0");
			oHandle1.yrange = ("15510,17520");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCStar2Type = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;

			this.loadPath("m10800,l8280,8259,,8259r6720,5146l4200,21600r6600,-5019l17400,21600,14880,13405,21600,8259r-8280,xe");

			this.loadConnectorsList("10800,0;0,8259;4200,21600;17400,21600;21600,8259");
			this.loadTextRect("6720,8259,14880,15628");
		};
		CVmlGeometryData.prototype.fillCStraightConnectorType = function() {
			this.loadPath("m,l21600,21600e");
			this.loadConnectorsList("0,0;21600,21600");
		};
		CVmlGeometryData.prototype.fillCStripedRightArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m@0,l@0@1,3375@1,3375@2@0@2@0,21600,21600,10800xem1350@1l1350@2,2700@2,2700@1xem0@1l0@2,675@2,675@1xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("sum height 0 #1");
			this.addGuide("sum 10800 0 #1");
			this.addGuide("sum width 0 #0");
			this.addGuide("prod @4 @3 10800");
			this.addGuide("sum width 0 @5");

			this.addAdjustment(16200);
			this.addAdjustment(5400);

			this.loadConnectorsList("@0,0;0,10800;@0,21600;21600,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);


			this.loadTextRect("3375,@1,@6,@2");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			oHandle1.xrange = ("3375,21600");
			oHandle1.yrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCSunType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m21600,10800l@15@14@15@18xem18436,3163l@17@12@16@13xem10800,l@14@10@18@10xem3163,3163l@12@13@13@12xem,10800l@10@18@10@14xem3163,18436l@13@16@12@17xem10800,21600l@18@15@14@15xem18436,18436l@16@17@17@16xem10800@19qx@19,10800,10800@20@20,10800,10800@19xe");

			this.addGuide("sum 10800 0 #0");
			this.addGuide("prod @0 30274 32768");
			this.addGuide("prod @0 12540 32768");
			this.addGuide("sum @1 10800 0");
			this.addGuide("sum @2 10800 0");
			this.addGuide("sum 10800 0 @1");
			this.addGuide("sum 10800 0 @2");
			this.addGuide("prod @0 23170 32768");
			this.addGuide("sum @7 10800 0");
			this.addGuide("sum 10800 0 @7");
			this.addGuide("prod @5 3 4");
			this.addGuide("prod @6 3 4");
			this.addGuide("sum @10 791 0");
			this.addGuide("sum @11 791 0");
			this.addGuide("sum @11 2700 0");
			this.addGuide("sum 21600 0 @10");
			this.addGuide("sum 21600 0 @12");
			this.addGuide("sum 21600 0 @13");
			this.addGuide("sum 21600 0 @14");
			this.addGuide("val #0");
			this.addGuide("sum 21600 0 #0");

			this.addAdjustment(5400);

			this.loadTextRect("@9,@9,@8,@8");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,center");
			oHandle1.xrange = ("2700,10125");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCTextboxType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m,l,21600r21600,l21600,xe");
		};
		CVmlGeometryData.prototype.fillCTrapezoidType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m,l@0,21600@1,21600,21600,xe");

			this.addGuide("val #0");
			this.addGuide("sum width 0 #0");
			this.addGuide("prod #0 1 2");
			this.addGuide("sum width 0 @2");
			this.addGuide("mid #0 width");
			this.addGuide("mid @1 0");
			this.addGuide("prod height width #0");
			this.addGuide("prod @6 1 2");
			this.addGuide("sum height 0 @7");
			this.addGuide("prod width 1 2");
			this.addGuide("sum #0 0 @9");
			this.addGuide("if @10 @8 0");
			this.addGuide("if @10 @7 height");

			this.addAdjustment(5400);

			this.loadConnectorsList("@3,10800;10800,21600;@2,10800;10800,0");
			this.loadTextRect("1800,1800,19800,19800;4500,4500,17100,17100;7200,7200,14400,14400");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,bottomRight");
			oHandle1.xrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCUpArrowCalloutType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m0@0l@3@0@3@2@1@2,10800,0@4@2@5@2@5@0,21600@0,21600,21600,,21600xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("sum 21600 0 #3");
			this.addGuide("sum #0 21600 0");
			this.addGuide("prod @6 1 2");

			this.addAdjustment(7200);
			this.addAdjustment(5400);
			this.addAdjustment(3600);
			this.addAdjustment(8100);

			this.loadConnectorsList("10800,0;0,@7;10800,21600;21600,@7");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("0,@0,21600,21600");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("topLeft,#0");
			oHandle1.yrange = ("@2,21600");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#1,topLeft");
			oHandle2.xrange = ("0,@3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#3,#2");
			oHandle3.xrange = ("@1,10800");
			oHandle3.yrange = ("0,@0");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCUpArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m0@0l@1@0@1,21600@2,21600@2@0,21600@0,10800,xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("prod #0 #1 10800");
			this.addGuide("sum #0 0 @3");

			this.addAdjustment(5400);
			this.addAdjustment(5400);

			this.loadConnectorsList("10800,0;0,@0;10800,21600;21600,@0");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("@1,@4,@2,21600");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#1,#0");
			oHandle1.xrange = ("0,10800");
			oHandle1.yrange = ("0,21600");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCUpDownArrowCalloutType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m0@0l@3@0@3@2@1@2,10800,0@4@2@5@2@5@0,21600@0,21600@8@5@8@5@9@4@9,10800,21600@1@9@3@9@3@8,0@8xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("val #2");
			this.addGuide("val #3");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("sum 21600 0 #3");
			this.addGuide("sum #0 21600 0");
			this.addGuide("prod @6 1 2");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("sum 21600 0 #2");

			this.addAdjustment(5400);
			this.addAdjustment(5400);
			this.addAdjustment(2700);
			this.addAdjustment(8100);

			this.loadConnectorsList("10800,0;0,10800;10800,21600;21600,10800");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("0,@0,21600,@8");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("topLeft,#0");
			oHandle1.yrange = ("@2,10800");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("#1,topLeft");
			oHandle2.xrange = ("0,@3");
			this.addHandle(oHandle2);

			let oHandle3 = new CVmlHandle();
			oHandle3.position = ("#3,#2");
			oHandle3.xrange = ("@1,10800");
			oHandle3.yrange = ("0,@0");
			this.addHandle(oHandle3);
		};
		CVmlGeometryData.prototype.fillCUpDownArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m10800,l21600@0@3@0@3@2,21600@2,10800,21600,0@2@1@2@1@0,0@0xe");

			this.addGuide("val #1");
			this.addGuide("val #0");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("prod #1 #0 10800");
			this.addGuide("sum #1 0 @4");
			this.addGuide("sum 21600 0 @5");

			this.addAdjustment(5400);
			this.addAdjustment(4320);

			this.loadConnectorsList("10800,0;0,@0;@1,10800;0,@2;10800,21600;21600,@2;@3,10800;21600,@0");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(180);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);
			this.addConnectorAngle(0);
			this.addConnectorAngle(0);

			this.loadTextRect("@1,@5,@3,@6");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			oHandle1.xrange = ("0,10800");
			oHandle1.yrange = ("0,10800");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCUturnArrowType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m15662,14285l21600,8310r-2970,qy9250,,,8485l,21600r6110,l6110,8310qy8907,5842l9725,5842qx12520,8310l9725,8310xe");

			this.loadConnectorsList("9250,0;3055,21600;9725,8310;15662,14285;21600,8310");

			this.addConnectorAngle(270);
			this.addConnectorAngle(90);
			this.addConnectorAngle(90);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			this.loadTextRect("0,8310,6110,21600");
		};
		CVmlGeometryData.prototype.fillCVerticalScrollType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@5,qx@1@2l@1@0@2@0qx0@7@2,21600l@9,21600qx@10@7l@10@1@11@1qx21600@2@11,xem@5,nfqx@6@2@5@1@4@3@5@2l@6@2em@5@1nfl@10@1em@2,21600nfqx@1@7l@1@0em@2@0nfqx@3@8@2@7l@1@7e");

			this.addGuide("sum height 0 #0");
			this.addGuide("val #0");
			this.addGuide("prod @1 1 2");
			this.addGuide("prod @1 3 4");
			this.addGuide("prod @1 5 4");
			this.addGuide("prod @1 3 2");
			this.addGuide("prod @1 2 1");
			this.addGuide("sum height 0 @2");
			this.addGuide("sum height 0 @3");
			this.addGuide("sum width 0 @5");
			this.addGuide("sum width 0 @1");
			this.addGuide("sum width 0 @2");
			this.addGuide("val height");
			this.addGuide("prod height 1 2");
			this.addGuide("prod width 1 2");

			this.addAdjustment(2700);

			this.loadConnectorsList("@14,0;@1,@13;@14,@12;@10,@13");
			this.loadTextRect("@1,@1,@10,@7");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("topLeft,#0");
			oHandle1.yrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCHorizontalScrollType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m0@5qy@2@1l@0@1@0@2qy@7,,21600@2l21600@9qy@7@10l@1@10@1@11qy@2,21600,0@11xem0@5nfqy@2@6@1@5@3@4@2@5l@2@6em@1@5nfl@1@10em21600@2nfqy@7@1l@0@1em@0@2nfqy@8@3@7@2l@7@1e");

			this.addGuide("sum width 0 #0");
			this.addGuide("val #0");
			this.addGuide("prod @1 1 2");
			this.addGuide("prod @1 3 4");
			this.addGuide("prod @1 5 4");
			this.addGuide("prod @1 3 2");
			this.addGuide("prod @1 2 1");
			this.addGuide("sum width 0 @2");
			this.addGuide("sum width 0 @3");
			this.addGuide("sum height 0 @5");
			this.addGuide("sum height 0 @1");
			this.addGuide("sum height 0 @2");
			this.addGuide("val width");
			this.addGuide("prod width 1 2");
			this.addGuide("prod height 1 2");

			this.addAdjustment(2700);

			this.loadConnectorsList("@13,@1;0,@14;@13,@10;@12,@14");
			this.loadTextRect("@1,@1,@7,@10");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("topLeft,#0");
			oHandle1.yrange = ("0,5400");
			this.addHandle(oHandle1);

			this.limoX = 10800;
			this.limoY = 10800;
		};
		CVmlGeometryData.prototype.fillCWedgeEllipseCalloutType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("wr,,21600,21600@15@16@17@18l@21@22xe");

			this.addGuide("val #0");
			this.addGuide("val #1");
			this.addGuide("sum 10800 0 #0");
			this.addGuide("sum 10800 0 #1");
			this.addGuide("atan2 @2 @3");
			this.addGuide("sumangle @4 11 0");
			this.addGuide("sumangle @4 0 11");
			this.addGuide("cos 10800 @4");
			this.addGuide("sin 10800 @4");
			this.addGuide("cos 10800 @5");
			this.addGuide("sin 10800 @5");
			this.addGuide("cos 10800 @6");
			this.addGuide("sin 10800 @6");
			this.addGuide("sum 10800 0 @7");
			this.addGuide("sum 10800 0 @8");
			this.addGuide("sum 10800 0 @9");
			this.addGuide("sum 10800 0 @10");
			this.addGuide("sum 10800 0 @11");
			this.addGuide("sum 10800 0 @12");
			this.addGuide("mod @2 @3 0");
			this.addGuide("sum @19 0 10800");
			this.addGuide("if @20 #0 @13");
			this.addGuide("if @20 #1 @14");

			this.addAdjustment(1350);
			this.addAdjustment(25920);

			this.loadConnectorsList("10800,0;3163,3163;0,10800;3163,18437;10800,21600;18437,18437;21600,10800;18437,3163;@21,@22");
			this.loadTextRect("3163,3163,18437,18437");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCWedgeRectCalloutType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m,l0@8@12@24,0@9,,21600@6,21600@15@27@7,21600,21600,21600,21600@9@18@30,21600@8,21600,0@7,0@21@33@6,xe");

			this.addGuide("sum 10800 0 #0");
			this.addGuide("sum 10800 0 #1");
			this.addGuide("sum #0 0 #1");
			this.addGuide("sum @0 @1 0");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("if @0 3600 12600");
			this.addGuide("if @0 9000 18000");
			this.addGuide("if @1 3600 12600");
			this.addGuide("if @1 9000 18000");
			this.addGuide("if @2 0 #0");
			this.addGuide("if @3 @10 0");
			this.addGuide("if #0 0 @11");
			this.addGuide("if @2 @6 #0");
			this.addGuide("if @3 @6 @13");
			this.addGuide("if @5 @6 @14");
			this.addGuide("if @2 #0 21600");
			this.addGuide("if @3 21600 @16");
			this.addGuide("if @4 21600 @17");
			this.addGuide("if @2 #0 @6");
			this.addGuide("if @3 @19 @6");
			this.addGuide("if #1 @6 @20");
			this.addGuide("if @2 @8 #1");
			this.addGuide("if @3 @22 @8");
			this.addGuide("if #0 @8 @23");
			this.addGuide("if @2 21600 #1");
			this.addGuide("if @3 21600 @25");
			this.addGuide("if @5 21600 @26");
			this.addGuide("if @2 #1 @8");
			this.addGuide("if @3 @8 @28");
			this.addGuide("if @4 @8 @29");
			this.addGuide("if @2 #1 0");
			this.addGuide("if @3 @31 0");
			this.addGuide("if #1 0 @32");
			this.addGuide("val #0");
			this.addGuide("val #1");

			this.addAdjustment(1350);
			this.addAdjustment(25920);

			this.loadConnectorsList("10800,0;0,10800;10800,21600;21600,10800;@34,@35");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCWedgeRoundedRectCalloutType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			this.loadPath("m3600,qx,3600l0@8@12@24,0@9,,18000qy3600,21600l@6,21600@15@27@7,21600,18000,21600qx21600,18000l21600@9@18@30,21600@8,21600,3600qy18000,l@7,0@21@33@6,xe");

			this.addGuide("sum 10800 0 #0");
			this.addGuide("sum 10800 0 #1");
			this.addGuide("sum #0 0 #1");
			this.addGuide("sum @0 @1 0");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("if @0 3600 12600");
			this.addGuide("if @0 9000 18000");
			this.addGuide("if @1 3600 12600");
			this.addGuide("if @1 9000 18000");
			this.addGuide("if @2 0 #0");
			this.addGuide("if @3 @10 0");
			this.addGuide("if #0 0 @11");
			this.addGuide("if @2 @6 #0");
			this.addGuide("if @3 @6 @13");
			this.addGuide("if @5 @6 @14");
			this.addGuide("if @2 #0 21600");
			this.addGuide("if @3 21600 @16");
			this.addGuide("if @4 21600 @17");
			this.addGuide("if @2 #0 @6");
			this.addGuide("if @3 @19 @6");
			this.addGuide("if #1 @6 @20");
			this.addGuide("if @2 @8 #1");
			this.addGuide("if @3 @22 @8");
			this.addGuide("if #0 @8 @23");
			this.addGuide("if @2 21600 #1");
			this.addGuide("if @3 21600 @25");
			this.addGuide("if @5 21600 @26");
			this.addGuide("if @2 #1 @8");
			this.addGuide("if @3 @8 @28");
			this.addGuide("if @4 @8 @29");
			this.addGuide("if @2 #1 0");
			this.addGuide("if @3 @31 0");
			this.addGuide("if #1 0 @32");
			this.addGuide("val #0");
			this.addGuide("val #1");

			this.addAdjustment(1350);
			this.addAdjustment(25920);

			this.loadConnectorsList("10800,0;0,10800;10800,21600;21600,10800;@34,@35");
			this.loadTextRect("791,791,20809,20809");

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("#0,#1");
			this.addHandle(oHandle1);
		};
		CVmlGeometryData.prototype.fillCWaveType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@28@0c@27@1@26@3@25@0l@21@4c@22@5@23@6@24@4xe");

			this.addGuide("val #0");
			this.addGuide("prod @0 41 9");
			this.addGuide("prod @0 23 9");
			this.addGuide("sum 0 0 @2");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("sum 21600 0 @1");
			this.addGuide("sum 21600 0 @3");
			this.addGuide("sum #1 0 10800");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("prod @8 2 3");
			this.addGuide("prod @8 4 3");
			this.addGuide("prod @8 2 1");
			this.addGuide("sum 21600 0 @9");
			this.addGuide("sum 21600 0 @10");
			this.addGuide("sum 21600 0 @11");
			this.addGuide("prod #1 2 3");
			this.addGuide("prod #1 4 3");
			this.addGuide("prod #1 2 1");
			this.addGuide("sum 21600 0 @15");
			this.addGuide("sum 21600 0 @16");
			this.addGuide("sum 21600 0 @17");
			this.addGuide("if @7 @14 0");
			this.addGuide("if @7 @13 @15");
			this.addGuide("if @7 @12 @16");
			this.addGuide("if @7 21600 @17");
			this.addGuide("if @7 0 @20");
			this.addGuide("if @7 @9 @19");
			this.addGuide("if @7 @10 @18");
			this.addGuide("if @7 @11 21600");
			this.addGuide("sum @24 0 @21");
			this.addGuide("sum @4 0 @0");
			this.addGuide("max @21 @25");
			this.addGuide("min @24 @28");
			this.addGuide("prod @0 2 1");
			this.addGuide("sum 21600 0 @33");
			this.addGuide("mid @26 @27");
			this.addGuide("mid @24 @28");
			this.addGuide("mid @22 @23");
			this.addGuide("mid @21 @25");

			this.addAdjustment(2809);
			this.addAdjustment(10800);

			this.loadConnectorsList("@35,@0;@38,10800;@37,@4;@36,10800");
			this.loadTextRect("@31,@33,@32,@34");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("topLeft,#0");
			oHandle1.yrange = ("0,4459");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("bottomRight,#1");
			oHandle2.xrange = ("8640,12960");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCWaveDoubleType = function() {
			this.concentricFill = false;
			this.join = ODRAW.lineJoinMiter;
			//Encaps: Flat

			this.loadPath("m@43@0c@42@1@41@3@40@0@39@1@38@3@37@0l@30@4c@31@5@32@6@33@4@34@5@35@6@36@4xe");

			this.addGuide("val #0");
			this.addGuide("prod @0 41 9");
			this.addGuide("prod @0 23 9");
			this.addGuide("sum 0 0 @2");
			this.addGuide("sum 21600 0 #0");
			this.addGuide("sum 21600 0 @1");
			this.addGuide("sum 21600 0 @3");
			this.addGuide("sum #1 0 10800");
			this.addGuide("sum 21600 0 #1");
			this.addGuide("prod @8 1 3");
			this.addGuide("prod @8 2 3");
			this.addGuide("prod @8 4 3");
			this.addGuide("prod @8 5 3");
			this.addGuide("prod @8 2 1");
			this.addGuide("sum 21600 0 @9");
			this.addGuide("sum 21600 0 @10");
			this.addGuide("sum 21600 0 @8");
			this.addGuide("sum 21600 0 @11");
			this.addGuide("sum 21600 0 @12");
			this.addGuide("sum 21600 0 @13");
			this.addGuide("prod #1 1 3");
			this.addGuide("prod #1 2 3");
			this.addGuide("prod #1 4 3");
			this.addGuide("prod #1 5 3");
			this.addGuide("prod #1 2 1");
			this.addGuide("sum 21600 0 @20");
			this.addGuide("sum 21600 0 @21");
			this.addGuide("sum 21600 0 @22");
			this.addGuide("sum 21600 0 @23");
			this.addGuide("sum 21600 0 @24");
			this.addGuide("if @7 @19 0");
			this.addGuide("if @7 @18 @20");
			this.addGuide("if @7 @17 @21");
			this.addGuide("if @7 @16 #1");
			this.addGuide("if @7 @15 @22");
			this.addGuide("if @7 @14 @23");
			this.addGuide("if @7 21600 @24");
			this.addGuide("if @7 0 @29");
			this.addGuide("if @7 @9 @28");
			this.addGuide("if @7 @10 @27");
			this.addGuide("if @7 @8 @8");
			this.addGuide("if @7 @11 @26");
			this.addGuide("if @7 @12 @25");
			this.addGuide("if @7 @13 21600");
			this.addGuide("sum @36 0 @30");
			this.addGuide("sum @4 0 @0");
			this.addGuide("max @30 @37");
			this.addGuide("min @36 @43");
			this.addGuide("prod @0 2 1");
			this.addGuide("sum 21600 0 @48");
			this.addGuide("mid @36 @43");
			this.addGuide("mid @30 @37");

			this.addAdjustment(1404);
			this.addAdjustment(10800);

			this.loadConnectorsList("@40,@0;@51,10800;@33,@4;@50,10800");
			this.loadTextRect("@46,@48,@47,@49");

			this.addConnectorAngle(270);
			this.addConnectorAngle(180);
			this.addConnectorAngle(90);
			this.addConnectorAngle(0);

			let oHandle1 = new CVmlHandle();
			oHandle1.position = ("topLeft,#0");
			oHandle1.yrange = ("0,2229");
			this.addHandle(oHandle1);

			let oHandle2 = new CVmlHandle();
			oHandle2.position = ("bottomRight,#1");
			oHandle2.xrange = ("8640,12960");
			this.addHandle(oHandle2);
		};
		CVmlGeometryData.prototype.fillCBentConnectorType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinRound;

			//m_strPathLimoX = ("m0,0l@0,0r0,21600l21600,21600nfe");
			//m_strPathLimoY = ("m0,0l,@0r21600,0l21600,21600nfe");

			this.loadPath("m,l@0,0@0,21600,21600,21600e");//m_strPathLimoX;

			this.addGuide("val #0");
			this.addAdjustment(10800);

			this.addAbsMaxAdjustment(82000);

			this.loadConnectorsList("0,0;21600,21600");
		};
		CVmlGeometryData.prototype.fillCCurvedConnectorType = function() {
			this.concentricFill = true;
			this.join = ODRAW.lineJoinRound;

			//m_strPathLimoX = ("m0,0qx@0,10800qy21600,21600nfe");
			//m_strPathLimoY = ("m0,0qy@0,10800qx21600,21600nfe");

			this.loadPath("m,c@0,0@1,5400@1,10800@1,16200@2,21600,21600,21600e");//m_strPathLimoX;

			this.addGuide("mid #0 0");
			this.addGuide("val #0");
			this.addGuide("mid #0 21600");

			//this.addGuide("val #0");
			this.addAdjustment(10800);

			this.loadConnectorsList("0,0;21600,21600");
		};

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
			else if ("strokeweight" === name) {
				let sValue = reader.GetValue();
				let oVal = new CUniversalMeasure();
				oVal.Parse(sValue, 1);
				this.m_oStrokeWeight = oVal.m_dValue;
			}
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
				oItem = new CBorder(name);
			else if ("borderleft" === name)
				oItem = new CBorder(name);
			else if ("borderright" === name)
				oItem = new CBorder(name);
			else if ("bordertop" === name)
				oItem = new CBorder(name);
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
				writer.WriteXmlNullableAttributeString("strokeweight", this.m_oStrokeWeight + "pt");

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
		CVmlCommonElements.prototype.writeChildrenXml = function (writer) {

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
		CVmlCommonElements.prototype.getShapeType = function () {
			return null;
		};
		CVmlCommonElements.prototype.findItemByConstructor = function(fConstructor) {
			for(let nItem = 0; nItem < this.items.length; ++nItem) {
				let oItem = this.items[nItem];
				if(oItem instanceof fConstructor) {
					return oItem;
				}
			}
			return null;
		};
		CVmlCommonElements.prototype.getShadow = function() {
			return this.findItemByConstructor(CShadow);
		};
		CVmlCommonElements.prototype.getWrap = function() {
			return this.findItemByConstructor(CWrap);
		};
		CVmlCommonElements.prototype.getImageData = function() {
			return this.findItemByConstructor(CImageData);
		};
		CVmlCommonElements.prototype.getFill = function() {
			return this.findItemByConstructor(CFillVml);
		};

		CVmlCommonElements.prototype.isSignatureLine = function() {
			if(this instanceof  CShape) {
				let oSL = this.getSignatureLine();
				if(oSL) {
					if(oSL.m_oIsSignatureLine === true) {
						return true;
					}
				}
			}
			return false;
		};
		CVmlCommonElements.prototype.getSignatureLine = function() {
			return this.findItemByConstructor(CSignatureLine);
		};
		CVmlCommonElements.prototype.getStroke = function() {
			return this.findItemByConstructor(CStroke);
		};
		CVmlCommonElements.prototype.getTextbox = function() {
			return this.findItemByConstructor(CTextbox);
		};
		CVmlCommonElements.prototype.getTextPath = function() {
			return this.findItemByConstructor(CTextPath);
		};
		CVmlCommonElements.prototype.getClientData = function() {
			return this.findItemByConstructor(CClientData);
		};
		CVmlCommonElements.prototype.getLeftBorder = function() {
			for(let nItem = 0; nItem < this.items.length; ++nItem) {
				let oItem = this.items[nItem];
				if(oItem instanceof CBorder && oItem.m_sType === "borderleft") {
					return oItem;
				}
			}
			return null;
		};
		CVmlCommonElements.prototype.correctFillOpacity = function(oFill) {
			if(this.m_oOpacity !== null) {
				oFill.addAlpha(this.m_oOpacity);
			}
			else {
				let oVMLFill = this.getFill();
				if(oVMLFill && oVMLFill.m_oOpacity !== null) {
					oFill.addAlpha(oVMLFill.m_oOpacity);
				}
			}
		};
		CVmlCommonElements.prototype.getOOXMLFill = function(oContext) {
			if(this.m_oFilled === false) {
				return AscFormat.CreateNoFillUniFill();
			}

			let oFill = null;
			//imagedata
			let oImageData = this.getImageData();
			if(oImageData) {
				oFill = oImageData.getOOXMLFill(oContext);
				if(oFill) {
					this.correctFillOpacity(oFill);
					return oFill;
				}
			}
			if(this.m_oFillColor) {
				oFill = this.m_oFillColor.getOOXMLFill(oContext);
				let oFillVML = this.getFill();
				if(oFillVML) {
					if(oFillVML.isGradient()) {
						oFill = oFillVML.getOOXMLFill(oContext, this.m_oFillColor)
					}
				}
				this.correctFillOpacity(oFill);
				return oFill;
			}
			let oFillVML = this.getFill();
			if(oFillVML) {
				return oFillVML.getOOXMLFill(oContext);
			}
			return AscFormat.CreateSolidFillRGB(0xFF, 0xFF, 0xFF);
		};
		CVmlCommonElements.prototype.getOOXMLStroke = function() {
			let oStroke = null;
			let oVMLStroke = this.getStroke();
			if(oVMLStroke) {
				oStroke = oVMLStroke.getOOXMLStroke();
				if(this.m_oStrokeWeight !== null) {
					oStroke.w = Pt_To_Emu(this.m_oStrokeWeight);
				}
				if(oVMLStroke.m_oColor === null) {
					if(this.m_oStrokeColor !== null) {
						oStroke.Fill = this.m_oStrokeColor.getOOXMLFill();
					}
					else {
						oStroke.Fill = AscFormat.CreateSolidFillRGB(0, 0, 0);
					}
				}
			}
			else {
				if(this.m_oStroked === false) {
					oStroke = AscFormat.CreateNoFillLine();
					if(this.m_oStrokeWeight !== null) {
						oStroke.w = Pt_To_Emu(this.m_oStrokeWeight);
					}
				}
				else {
					if(this.m_oStrokeWeight !== null || this.m_oStrokeColor !== null) {
						oStroke = new AscFormat.CLn();
						if(this.m_oStrokeColor !== null) {
							oStroke.Fill = this.m_oStrokeColor.getOOXMLFill();
						}
						else {
							oStroke.Fill = AscFormat.CreateSolidFillRGB(0, 0, 0);
						}
						if(this.m_oStrokeWeight !== null) {
							oStroke.w = Pt_To_Emu(this.m_oStrokeWeight);
						}
					}
				}
			}
			if(!oStroke || !oStroke.isVisible()) {
				//take line from left border
				let oLeftBorder = this.getLeftBorder();
				if(oLeftBorder) {
					if(oLeftBorder.m_oType !== EBorderType.bordertypeNone) {
						oStroke = new AscFormat.CLn();
						if(oLeftBorder.m_oWidth !== null) {
							oStroke.w = Pt_To_Emu(oLeftBorder.m_oWidth);
						}
						switch (oLeftBorder.m_oType) {
							case EBorderType.bordertypeDash:			oStroke.prstDash = 3; break;
							case EBorderType.bordertypeDashDotDot:		oStroke.prstDash = 5; break;
							case EBorderType.bordertypeDashDotStroked:	oStroke.prstDash = 1; break;
							case EBorderType.bordertypeDashedSmall:	oStroke.prstDash = 0; break;
							case EBorderType.bordertypeDot:			oStroke.prstDash = 2; break;
							case EBorderType.bordertypeDotDash:		oStroke.prstDash = 1; break;
						}
						let oLeftBorderColor = this.m_oBorderLeftColor;
						if(oLeftBorderColor) {
							oStroke.Fill = oLeftBorderColor.getOOXMLFill();
						}
					}
				}
			}
			if(!oStroke) {
				oStroke = new AscFormat.CLn();
				oStroke.setFill(AscFormat.CreateSolidFillRGB(0, 0, 0));
			}
			// if(!oStroke.isVisible()) {
			// 	return null;
			// }
			return oStroke;
		};
		CVmlCommonElements.prototype.createSpPrIfNoPresent = function(oSpPr) {
			let oWorkSpPr = oSpPr;
			if(!oWorkSpPr) {
				oWorkSpPr = new AscFormat.CSpPr();
			}
			return oWorkSpPr;
		};
		CVmlCommonElements.prototype.convertFillToOOXML = function(oSpPr, oContext) {
			let oWorkSpPr = this.createSpPrIfNoPresent(oSpPr);
			oWorkSpPr.setFill(this.getOOXMLFill(oContext));
			return oWorkSpPr;
		};
		CVmlCommonElements.prototype.convertStrokeToOOXML = function(oSpPr) {
			let oWorkSpPr = this.createSpPrIfNoPresent(oSpPr);
			oWorkSpPr.setLn(this.getOOXMLStroke());
			return oWorkSpPr;
		};
		CVmlCommonElements.prototype.convertFillStrokeToOOXML = function(oSpPr, oContext) {
			return this.convertFillToOOXML(this.convertStrokeToOOXML(oSpPr), oContext);
		};
		CVmlCommonElements.prototype.convertFlipRot = function(oXfrm) {

			if(this.m_oStyle) {
				let sFlip = this.m_oStyle.GetPropertyValueString("flip");
				if (sFlip !== null)
				{
					if (sFlip === "x") {
						oXfrm.setFlipH(true);
					}
					else if (sFlip === "y") {
						oXfrm.setFlipV(true);
					}
					else if ((sFlip === "xy") || (sFlip === "yx") || (sFlip === "x y") || (sFlip === "y x")
						|| (sFlip === "y,x") || (sFlip === "x,y"))
					{
						oXfrm.setFlipH(true);
						oXfrm.setFlipV(true);
					}
				}
				let sRot = this.m_oStyle.GetPropertyValueString("rotation");
				if (sRot) {
					oXfrm.setRot(getRotateAngle(sRot, oXfrm.flipH, oXfrm.flipV));
				}
			}
		};
		CVmlCommonElements.prototype.getMainProperties = function(oContext) {
			let oProps = {IsTop: oContext.bIsTopDrawing};
			let oOldItem = oContext.sourceItem;
			oContext.sourceItem = this;
			CLegacyDrawing.prototype.GetDrawingMainProps(null, oContext, oProps);
			oContext.sourceItem = oOldItem;
			return oProps;
		};
		CVmlCommonElements.prototype.convertToOOXML = function(aOtherElements, oOOXMLGroup, oContext) {
			let oShapeType = CLegacyDrawing.prototype.static_GetShapeTypeForShape(this, aOtherElements);
			let oSpPr = new AscFormat.CSpPr();
			this.convertFillToOOXML(oSpPr, oContext);
			this.convertStrokeToOOXML(oSpPr);


			let bIsTop = oContext.bIsTopDrawing;

			let sStyleAdvanced = null;
			let oGeometryData = new CVmlGeometryData();
			let nType = this.getFinalShapeType(aOtherElements);
			oGeometryData.fillByType(nType);
			let sAdj = this.m_sAdj || oShapeType && oShapeType.m_sAdj;
			if(sAdj) {
				let aAdj = sAdj.split(",");
				for(let nAdj = 0; nAdj < aAdj.length; ++nAdj) {
					let nAdjVal = parseInt(aAdj[nAdj]);
					if(AscFormat.isRealNumber(nAdjVal)) {
						oGeometryData.adjustments[nAdj] = nAdjVal;
					}
				}
			}
			if(this instanceof CRoundRect) {
				if(this.m_oArcSize && AscFormat.isRealNumber(this.m_oArcSize.m_dValue)) {
					oGeometryData.adjustments[0] = (this.m_oArcSize.m_dValue * 65536 / 10.0 + 0.5) >> 0;
				}
 			}
			let sPath = this.m_oPath || oShapeType && oShapeType.m_oPath;
			let bNeeLoadCoordSize = true;
			if(this instanceof CLine) {
				if(this.m_oFrom && this.m_oTo) {
					let x1, y1, x2, y2;
					x1 = this.m_oFrom.m_dX;
					y1 = this.m_oFrom.m_dY;
					x2 = this.m_oTo.m_dX;
					y2 = this.m_oTo.m_dY;
					if (x1 > x2)
					{
						let tmp = x1;
						x1 = x2;
						x2 = tmp;
					}
					if (y1 > y2)
					{
						let tmp = y1;
						y1 = y2;
						y2 = tmp;
					}
					sStyleAdvanced =  ";left:" + x1
					+ ";top:"     + y1
					+ ";width:"   + (x2-x1)
					+ ";height:"  + (y2-y1)
					+ ";";
				}
			}
			if(this instanceof CPolyLine) {
				bNeeLoadCoordSize = false;
				if(this.m_oPoints) {
					let oPath = this.m_oPoints.ToSVGPath();
					if(oPath) {
						sPath = oPath.path;
						let oBounds = oPath.bounds;
						sStyleAdvanced += ";margin-left:" + oBounds.l + ";margin-top:" + oBounds.t
						+ ";width:" + (oBounds.r - oBounds.l) + ";height:" + (oBounds.b - oBounds.t) + ";polyline_correct:true;";
					}
				}
			}
			if(sPath) {
				oGeometryData.loadPath(sPath);
			}
			if(bNeeLoadCoordSize) {
				oGeometryData.loadCoordSize(this.m_oCoordSize);
			}
			oSpPr.setGeometry(oGeometryData.convertToOOXML());

			let oOldCSS = this.m_oStyle;
			if(sStyleAdvanced) {
				let sNewCSS = sStyleAdvanced + (this.m_oStyle && this.m_oStyle.m_sCss || "");
				this.m_oStyle = new CCssStyle(sNewCSS);
			}
			let oProps = this.getMainProperties(oContext);
			let oXfrm = new AscFormat.CXfrm();
			if (bIsTop)
			{
				oXfrm.setOffX(0);
				oXfrm.setOffY(0);
			}
			else
			{
				oXfrm.setOffX(oProps.X);
				oXfrm.setOffY(oProps.Y);
			}

			oXfrm.setExtX(oProps.Width);
			oXfrm.setExtY(oProps.Height);
			this.convertFlipRot(oXfrm);
			oSpPr.setXfrm(oXfrm);

			//this.m_oStyle = oOldCSS;


			let oOOXMLDrawing;
			let oSignatureLine = this.getSignatureLine();
			let bIsPicture = false;
			if(nType === ShapeType.sptCFrame && !oSignatureLine) {
				oOOXMLDrawing = new AscFormat.CImageShape();
				bIsPicture = true;
				if(oSpPr.Fill && oSpPr.Fill.isBlipFill()) {
					oOOXMLDrawing.setBlipFill(oSpPr.Fill.fill);
					oSpPr.setFill(null);
					oSpPr.setLn(null);
				}
				else {
					return null;
				}
			}
			else {
				oOOXMLDrawing = new AscFormat.CShape();
				oOOXMLDrawing.setWordShape(true);
				if(oSignatureLine && oSignatureLine.m_oIsSignatureLine) {
					let oOOXMLSignatureLine = new AscFormat.CSignatureLine();
					oOOXMLSignatureLine.id = oSignatureLine.m_oId;
					oOOXMLSignatureLine.signer = oSignatureLine.m_sSuggestedSigner;
					oOOXMLSignatureLine.signer2 = oSignatureLine.m_sSuggestedSigner2;
					oOOXMLSignatureLine.email = oSignatureLine.m_sSuggestedSignerEmail;
					oOOXMLSignatureLine.showDate = oSignatureLine.m_oShowSignDate;
					oOOXMLSignatureLine.instructions = oSignatureLine.m_sSigningInstructions;
					oOOXMLDrawing.setSignature(oOOXMLSignatureLine);
				}
			}
			let oNvPr = new AscFormat.UniNvPr();
			oOOXMLDrawing.setNvSpPr(oNvPr);
			oOOXMLDrawing.setSpPr(oSpPr);
			if(!bIsPicture) {
				let bIsWordArt = this.isWordArt(oContext.aOtherElements);
				let oDocContent = null;
				let oCSSStyle = null;
				let sFontName = "Arial";
				let bBold = false;
				let bItalic = false;
				let nFontSize = 11;
				let sText = "";
				let oTextFill, oTextStroke;
				let oBodyPr = new AscFormat.CBodyPr();
				if(bIsWordArt) {
					let eTextShapeType;
					let oTextPath = this.getTextPath();
					switch (nType)
					{
						case ShapeType.sptCTextPlain:					eTextShapeType = "textPlain"; break;
						case ShapeType.sptCTextArchUp:					eTextShapeType = "textArchUp"; break;
						case ShapeType.sptCTextArchDown:				eTextShapeType = "textArchDown"; break;
						case ShapeType.sptCTextButton:					eTextShapeType = "textButton"; break;
						case ShapeType.sptCTextCurveUp:					eTextShapeType = "textCurveUp"; break;
						case ShapeType.sptCTextCurveDown:				eTextShapeType = "textCurveDown"; break;
						case ShapeType.sptCTextCanUp:					eTextShapeType = "textCanUp"; break;
						case ShapeType.sptCTextCanDown:					eTextShapeType = "textCanDown"; break;
						case ShapeType.sptCTextWave1:					eTextShapeType = "textWave1"; break;
						case ShapeType.sptCTextWave2:					eTextShapeType = "textWave2"; break;
						case ShapeType.sptCTextWave3:					eTextShapeType = "textDoubleWave1"; break;
						case ShapeType.sptCTextWave4:					eTextShapeType = "textWave4"; break;
						case ShapeType.sptCTextInflate:					eTextShapeType = "textInflate"; break;
						case ShapeType.sptCTextDeflate:					eTextShapeType = "textDeflate"; break;
						case ShapeType.sptCTextInflateBottom:			eTextShapeType = "textInflateBottom"; break;
						case ShapeType.sptCTextDeflateBottom:			eTextShapeType = "textDeflateBottom"; break;
						case ShapeType.sptCTextInflateTop:				eTextShapeType = "textInflateTop"; break;
						case ShapeType.sptCTextDeflateTop:				eTextShapeType = "textDeflateTop"; break;
						case ShapeType.sptCTextDeflateInflate:			eTextShapeType = "textDeflateInflate"; break;
						case ShapeType.sptCTextDeflateInflateDeflate:	eTextShapeType = "textDeflateInflateDeflate"; break;
						case ShapeType.sptCTextFadeRight:				eTextShapeType = "textFadeRight"; break;
						case ShapeType.sptCTextFadeLeft:				eTextShapeType = "textFadeLeft"; break;
						case ShapeType.sptCTextFadeUp:					eTextShapeType = "textFadeUp"; break;
						case ShapeType.sptCTextFadeDown:				eTextShapeType = "textFadeDown"; break;
						case ShapeType.sptCTextSlantUp:					eTextShapeType = "textSlantUp"; break;
						case ShapeType.sptCTextSlantDown:				eTextShapeType = "textSlantDown"; break;
						case ShapeType.sptCTextCascadeUp:				eTextShapeType = "textCascadeUp"; break;
						case ShapeType.sptCTextCascadeDown:				eTextShapeType = "textCascadeDown"; break;
						case ShapeType.sptCTextButtonPour:				eTextShapeType = "textButtonPour"; break;
						case ShapeType.sptCTextStop:					eTextShapeType = "textStop"; break;
						case ShapeType.sptCTextTriangle:				eTextShapeType = "textTriangle"; break;
						case ShapeType.sptCTextTriangleInverted:		eTextShapeType = "textTriangleInverted"; break;
						case ShapeType.sptCTextChevron:					eTextShapeType = "textChevron"; break;
						case ShapeType.sptCTextChevronInverted:			eTextShapeType = "textChevronInverted"; break;
						case ShapeType.sptCTextRingInside:				eTextShapeType = "textRingInside"; break;
						case ShapeType.sptCTextRingOutside:				eTextShapeType = "textRingOutside"; break;
						case ShapeType.sptCTextCirclePour:				eTextShapeType = "textCirclePour"; break;
						case ShapeType.sptCTextArchUpPour:				eTextShapeType = "textArchUpPour"; break;
						case ShapeType.sptCTextArchDownPour:			eTextShapeType = "textArchDownPour"; break;
						default:										eTextShapeType = "textNoShape"; break;
					}
					oBodyPr.prstTxWarp = AscFormat.CreatePrstTxWarpGeometry(eTextShapeType);
					oTextFill = this.getOOXMLFill(oContext);
					oTextStroke = this.getOOXMLStroke();
					oSpPr.setGeometry(AscFormat.CreateGeometry("rect"));
					oSpPr.setFill(AscFormat.CreateNoFillUniFill());
					oSpPr.setLn(null);
					oBodyPr.lIns = 0;
					oBodyPr.tIns = 0;
					oBodyPr.rIns = 0;
					oBodyPr.bIns = 0;


					if(oTextPath) {
						sText = oTextPath.m_sString || "";
						oCSSStyle = oTextPath.m_oStyle;

					}


					if (oTextPath && (oTextPath.m_oFitShape || oTextPath.m_oFitPath)) {
						oBodyPr.textFit = new AscFormat.CTextFit();
						oBodyPr.textFit.type = AscFormat.text_fit_NormAuto;
					}
					oBodyPr.wrap = AscFormat.nTWTSquare;
					oBodyPr.fromWordArt = true;
					oOOXMLDrawing.setBodyPr(oBodyPr);
					oDocContent = new CDocumentContent(oOOXMLDrawing, oContext.DrawingDocument, 0, 0, 0, 0, false, false, false)
					oDocContent.MoveCursorToStartPos(false);
					oDocContent.AddText(sText);
					oOOXMLDrawing.setTextBoxContent(oDocContent);
				}
				else {
					oBodyPr.wrap = AscFormat.nTWTSquare;
					oBodyPr.upright = true;
					let oTextbox = this.getTextbox();
					if(oTextbox) {
						if(oTextbox.m_oTxtbxContent) {
							//oBodyPr.setAnchor(1);
							let oInset = oTextbox.m_oInset;
							if(oInset) {
								if(oInset.m_dLeft !== null) {
									oBodyPr.lIns = Pt_To_Mm(oInset.m_dLeft);
								}
								if(oInset.m_dTop !== null) {
									oBodyPr.tIns = Pt_To_Mm(oInset.m_dTop);
								}
								if(oInset.m_dRight !== null) {
									oBodyPr.rIns = Pt_To_Mm(oInset.m_dRight);
								}
								if(oInset.m_dBottom !== null) {
									oBodyPr.bIns = Pt_To_Mm(oInset.m_dBottom);
								}
							}
							
							let oCssStyle = oTextbox.m_oStyle;
							if(oCssStyle) {
								let oProperty = oCssStyle.GetProperty(ECssPropertyType.cssptLayoutFlow);
								if (oProperty) {
									if (oProperty.m_oValue.eLayoutFlow === ECssLayoutFlow.csslayoutflowVertical) {
										oBodyPr.vert = AscFormat.nVertTTvert;
									}
								}
								oProperty = oCssStyle.GetProperty(ECssPropertyType.cssptMsoLayoutFlowAlt);
								if (oProperty) {
									if (oProperty.m_oValue.eLayoutFlowAlt === ECssLayoutFlowAlt.csslayoutflowaltBottomToTop) {
										oBodyPr.vert = AscFormat.nVertTTvert270;
									}
								}
								oProperty = oCssStyle.GetProperty(ECssPropertyType.cssptMsoRotate);
								if (oProperty) {
									let val = 0;
									switch (oProperty.m_oValue.eRotate) {
										case ECssMsoRotate.cssmsorotate90:	val = 90; break;
										case ECssMsoRotate.cssmsorotate180:	val = 180; break;
										case ECssMsoRotate.cssmsorotate270:	val = 270; break;
									}
									oBodyPr.rot = val * 60000;
								}
								oProperty = oCssStyle.GetProperty(ECssPropertyType.cssptMsoFitShapeToText);
								if (oProperty) {
									if (oProperty.m_oValue.bValue) {
										oBodyPr.textFit = new AscFormat.CTextFit(AscFormat.text_fit_Auto);
									}
								}
								oProperty = oCssStyle.GetProperty(ECssPropertyType.cssptMsoFitTextToShape);
								if (oProperty) {
									if (oProperty.m_oValue.bValue) {
										oBodyPr.textFit = new AscFormat.CTextFit(AscFormat.text_fit_No);
									}
								}
								oProperty = oCssStyle.GetProperty(ECssPropertyType.cssptMsoTextScale);
								if (oProperty) {
									if (oProperty.m_oValue.oValue.eType === ECssUnitsType.cssunitstypeUnits) {
										oBodyPr.textFit = new AscFormat.CTextFit(AscFormat.text_fit_NormAuto);
										oBodyPr.textFit.fontScale = (100 * oProperty.m_oValue.oValue.dValue + 0.5) >> 0;
									}
								}
							}
							oOOXMLDrawing.setBodyPr(oBodyPr);
							oOOXMLDrawing.setTextBoxContent(oTextbox.m_oTxtbxContent);
							oTextbox.m_oTxtbxContent.SetParent(oOOXMLDrawing);
						}
						else if(oTextbox.m_oText) {
							//oBodyPr.setAnchor(1);
							oOOXMLDrawing.setBodyPr(oBodyPr);
							oDocContent = new CDocumentContent(oOOXMLDrawing, oContext.DrawingDocument, 0, 0, 0, 0, false, false, false)
							oDocContent.MoveCursorToStartPos(false);
							oDocContent.AddText(oTextbox.m_oText);
							oOOXMLDrawing.setTextBoxContent(oDocContent);
							if(oTextbox.m_oTextStyle) {
								oCSSStyle = oTextbox.m_oTextStyle;
							}
						}
					}
				}

				if(this.m_oStyle) {
					let sCSSAnchor = this.m_oStyle.GetPropertyValueString("v-text-anchor");
					if(sCSSAnchor) {
						if (sCSSAnchor === "middle")					oBodyPr.setAnchor(oBodyPr.GetAnchorCode("ctr"));
						if (sCSSAnchor === "bottom")					oBodyPr.setAnchor(oBodyPr.GetAnchorCode("b"));
						if (sCSSAnchor === "top-center")				oBodyPr.setAnchor(oBodyPr.GetAnchorCode("t"));
						if (sCSSAnchor === "middle-center")			    oBodyPr.setAnchor(oBodyPr.GetAnchorCode("ctr"));
						if (sCSSAnchor === "bottom-center")			    oBodyPr.setAnchor(oBodyPr.GetAnchorCode("b"));
						if (sCSSAnchor === "top-baseline")			    oBodyPr.setAnchor(oBodyPr.GetAnchorCode("t"));
						if (sCSSAnchor === "bottom-baseline")		    oBodyPr.setAnchor(oBodyPr.GetAnchorCode("b"));
						if (sCSSAnchor === "top-center-baseline")	    oBodyPr.setAnchor(oBodyPr.GetAnchorCode("t"));
						if (sCSSAnchor === "bottom-center-baseline")	oBodyPr.setAnchor(oBodyPr.GetAnchorCode("b"));
					}
				}
				if(oDocContent) {
					if(oCSSStyle) {
						let sCSSFont = oCSSStyle.GetStringValue(ECssPropertyType.cssptFontFamily);
						if(typeof sCSSFont === "string" && sCSSFont.length > 0) {
							sFontName = sCSSFont.replace(new RegExp("\"", 'g'), "");
						}
						let nCSSFontSize = oCSSStyle.GetNumberValue(ECssPropertyType.cssptFontSize);
						if(nCSSFontSize !== null) {
							nFontSize = nCSSFontSize;
						}
						let oFontStylePr = oCSSStyle.GetProperty(ECssPropertyType.cssptFontStyle);
						if(oFontStylePr) {
							let oValue = oFontStylePr.m_oValue;
							if(oValue.eFontStyle === ECssFontStyle.cssfontstyleItalic) {
								bItalic = true;
							}
						}
						let oFontWeightPr = oCSSStyle.GetProperty(ECssPropertyType.cssptFontWeight);
						if(oFontWeightPr) {
							let oValue = oFontWeightPr.m_oValue;
							if(oValue.eFontWeight >= ECssFontWeight.cssfontweight400) {
								bBold = true;
							}
						}
					}
					let oParaPr = new AscCommonWord.CParaPr();
					let oTextPr = new AscCommonWord.CTextPr();
					oParaPr.Jc = AscCommon.align_Center;
					oTextPr.RFonts.Ascii = { Name: sFontName, Index : -1 };
					oTextPr.RFonts.HAnsi = { Name: sFontName, Index : -1 };
					if(bBold) {
						oTextPr.Bold = bBold;
					}
					if(bItalic) {
						oTextPr.Italic = bItalic;
					}
					oTextPr.TextFill = oTextFill;
					oTextPr.TextOutline = oTextStroke;
					oTextPr.FontSize = nFontSize;
					oDocContent.SetApplyToAll(true);
					oDocContent.SetParagraphPr(oParaPr);
					oDocContent.AddToParagraph(new AscCommonWord.ParaTextPr(oTextPr));
					oDocContent.SetApplyToAll(false);
				}
			}
			oOOXMLDrawing.setBDeleted(false);
			return oOOXMLDrawing;
		};
		CVmlCommonElements.prototype.getFinalShapeType = function(aOtherElements) {
			let oShapeType = null;
			if(this instanceof CShape) {
				oShapeType = CLegacyDrawing.prototype.static_GetShapeTypeForShape(this, aOtherElements);
			}
			let nShapeType = null;

			if(this instanceof CBackground ) {
				nShapeType = ShapeType.sptCRect;
			}
			if(this instanceof CRect) {
				nShapeType = ShapeType.sptCRect;
			}
			if(this instanceof CRoundRect) {
				nShapeType = ShapeType.sptCRoundRect;
				//TODO: adjustment
			}
			if(this instanceof COval) {
				nShapeType = ShapeType.sptCEllipse;
			}
			if(this instanceof CImage) {
				nShapeType = ShapeType.sptCFrame;
			}
			if(this instanceof CLine) {
				nShapeType = ShapeType.sptCLine;
				//TODO: creates advanced css style with coordinates
			}
			if(this instanceof CPolyLine) {
				nShapeType = ShapeType.sptCustom;
				//TODO: creates advanced css style with coordinates
			}
			if(this instanceof CShape) {
				if(oShapeType) {
					//TODO: copy properties from shapetype
					nShapeType = oShapeType.getShapeType();
				}
				else {
					nShapeType = this.getShapeType();
				}
			}
			if(nShapeType === null) {
				let sConnectorType = this.m_oConnectorType;
				if (sConnectorType === "elbow") nShapeType = ShapeType.sptCBentConnector2;
				else if (sConnectorType === "straight")	nShapeType = ShapeType.sptCStraightConnector1;
				else if (sConnectorType === "curved") nShapeType = ShapeType.sptCCurvedConnector2;
			}
			return nShapeType;
		};
		CVmlCommonElements.prototype.isWordArt = function(aOtherElements) {
			let nType = this.getFinalShapeType(aOtherElements);
			if (nType >= ShapeType.sptCTextPlain && nType <= ShapeType.sptCTextCanDown) {
				return true;
			}
			let oTextPath = this.getTextPath();
			if(oTextPath) {
				return true;
			}
			return false;
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
		CArc.prototype.getShapeType = function () {
			return ShapeType.sptCArc;
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
				this.m_dValue = Math.min(100.0, Math.max(-100.0, (sValue.substring(0, nLen - 1))));

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
			else if ("id" === name) {
				let sName = reader.GetName();
				if(sName === "id") {
					this.m_sId = reader.GetValue();
				}
				else if(sName === "r:id" || "relationships:id") {
					this.m_rId = reader.GetValue();
				}
			}
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

			writer.WriteXmlNullableAttributeString("r:id", this.m_rId);
			writer.WriteXmlNullableAttributeString("o:relid", this.m_rId);
		};
		CFillVml.prototype.writeChildrenXml = function (writer) {
			this.m_oFill && this.m_oFill.toXml(writer, "o:fill");
		};
		CFillVml.prototype.isGradient = function() {
			return (this.m_oType === EFillType.filltypeGradient || this.m_oType === EFillType.filltypeGradientRadial);
		};
		CFillVml.prototype.getOOXMLFill = function(oContext, oFirstColor) {
			let oFill = null;
			if(this.isGradient()) {
				oFill = new AscFormat.CUniFill();
				let oGradFill = new AscFormat.CGradFill();
				oFill.fill = oGradFill;
				let oGs;
				let oStartColor = this.m_oColor || oFirstColor;
				if (oStartColor) {
					oGs = new AscFormat.CGs();
					oGs.setColor(oStartColor.getOOXMLColor());
					oGradFill.addColor(oGs);
				}
				if(this.m_oColor2) {

					oGs = new AscFormat.CGs();
					oGs.pos = 100000;
					oGs.setColor(this.m_oColor2.getOOXMLColor());
					oGradFill.addColor(oGs);
				}
				if (oGradFill.getColorsCount() === 1) {//Sindicatum.docx

					let oGs_ = new AscFormat.CGs();
					oGs_.setColor(AscFormat.CreateUniColorRGB(0xFF, 0xFF, 0xFF));
					if (oGs.pos === 0) {
						oGs_.pos = 100 * 1000;
					}
					oGradFill.addColor(oGs_);
				}
				if(this.m_oRotate === true) {
					oGradFill.rotateWithShape = true;
				}
				let nAngle = 90;
				let nFocus = 0;
				if(this.m_oAngle !== null) {
					nAngle =  (-1) * this.m_oAngle + 90;
				}
				if(this.m_oFocus && this.m_oFocus.m_dValue) {
					nFocus = parseInt(this.m_oFocus / 100);
				}
				let oGradLin = new AscFormat.GradLin();
				oGradLin.angle = (nAngle * 60000 + 0.5) >> 0;
				oGradFill.lin = oGradLin;
			}
			else if (typeof this.m_rId === "string" && this.m_rId.length > 0) {
				oFill = new AscFormat.CreateBlipFillUniFillFromUrl("");

				AscFormat.fReadXmlRasterImageId(oContext.reader, this.m_rId, oFill.fill);
				if(EFillType.filltypeTile === this.m_oType || EFillType.filltypePattern) {
					oFill.fill.tile = new AscFormat.CBlipFillTile();
				}
				else {
					oFill.fill.stretch = true;
				}
			}
			else if (this.m_oColor) {
				return this.m_oColor.getOOXMLFill();
			}
			if(!oFill) {
				oFill = AscFormat.CreateSolidFillRGB(0xFF, 0xFF, 0xFF);
			}
			if(this.m_oOpacity !== null) {
				oFill.addAlpha(this.m_oOpacity);
			}
			return oFill;
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
		CBackground.prototype.getShapeType = function () {
			return ShapeType.sptCRect;
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
		CFormulas.prototype.writeChildrenXml = function (writer) {

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
		CHandles.prototype.writeChildrenXml = function (writer) {
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
		CImage.prototype.getShapeType = function () {
			return ShapeType.sptCFrame;
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
			else if ("relid" === name) this.m_oRelId = reader.GetValue();
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
		CImageData.prototype.getOOXMLFill = function(oContext) {
			let sRid = this.m_rId;
			let sRelid = this.m_oRelId;
			let sPictId = this.m_rPict;
			let sRasterImageId = sRid || sRelid || sPictId;
			if(sRasterImageId) {
				let oFill = new AscFormat.CreateBlipFillUniFillFromUrl("");

				AscFormat.fReadXmlRasterImageId(oContext.reader, sRasterImageId, oFill.fill);
				//TODO: check tile
				//crop
				if(this.m_oCropLeft !== null || this.m_oCropTop !== null ||
					this.m_oCropRight !== null || this.m_oCropBottom !== null) {
					oFill.fill.srcRect = new AscFormat.CSrcRect();
					oFill.fill.srcRect.l = (this.m_oCropLeft !== null) ? this.m_oCropLeft * 100 : 0;
					oFill.fill.srcRect.t = (this.m_oCropTop !== null) ? this.m_oCropTop * 100 : 0;
					oFill.fill.srcRect.r = 100 - ((this.m_oCropRight !== null) ? this.m_oCropRight * 100 : 0);
					oFill.fill.srcRect.b = 100 - ((this.m_oCropBottom !== null) ? this.m_oCropBottom * 100 : 0);
				}
				return oFill;
			}
			return null;
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
		CLine.prototype.getShapeType = function () {
			return ShapeType.sptCLine;
		};

		function COval() {
			CVmlCommonElements.call(this);
		}

		IC(COval, CVmlCommonElements, AscDFH.historyitem_type_VMLOval);
		COval.prototype.getShapeType = function () {
			return ShapeType.sptCEllipse;
		};

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
		CPolyLine.prototype.getShapeType = function () {
			return ShapeType.sptCustom;
		};

		function CRect() {
			CVmlCommonElements.call(this);
		}

		IC(CRect, CVmlCommonElements, AscDFH.historyitem_type_VMLRect);
		CRect.prototype.getShapeType = function () {
			return ShapeType.sptCRect;
		};

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
		CRoundRect.prototype.getShapeType = function () {
			return ShapeType.sptCRoundRect;
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
		CShapeType.prototype.getShapeType = function() {
			return this.m_oSpt;
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
		CShape.prototype.getShapeType = function() {
			return parseShapeType(this.m_sType);
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

		function parseShapeType(sValue) {
			if(!(typeof sValue === "string" && sValue.length > 0)) {
				return null;
			}
			let sWorkingValue;
			if(sValue.indexOf("#") === 0) {
				sWorkingValue = sValue.substring(1);
			}
			else {
				sWorkingValue = sValue;
			}
			let aVals = sWorkingValue.split("_");
			for(let nVal = 0; nVal < aVals.length; ++nVal) {
				let sVal = aVals[nVal];
				if(sVal.charAt(0) === 't') {
					let nParsedVal = parseInt(sVal.substring(1));
					if(!isNaN(nParsedVal) && nParsedVal > ShapeType.sptMin && nParsedVal < ShapeType.sptMax) {
						return nParsedVal;
					}
				}
			}
			return null;
		}

		let ShapeType  = {};
			ShapeType.sptMin = 0;
			ShapeType.sptCustom = 0;
			ShapeType.sptNotPrimitive = 0;
			ShapeType.sptCRect = 1;
			ShapeType.sptCRoundRect = 2;
			ShapeType.sptCEllipse = 3;
			ShapeType.sptCDiamond = 4;
			ShapeType.sptCIsocelesTriangle = 5;
			ShapeType.sptCRtTriangle = 6;
			ShapeType.sptCParallelogram = 7;
			ShapeType.sptCTrapezoid = 8;
			ShapeType.sptCHexagon = 9;
			ShapeType.sptCOctagon = 10;
			ShapeType.sptCPlus = 11;
			ShapeType.sptCStar5 = 12;
			ShapeType.sptCRightArrow = 13;
			ShapeType.sptCThickArrow = 14;
			ShapeType.sptCHomePlate = 15;
			ShapeType.sptCCube = 16;
			ShapeType.sptCBalloon = 17;
			ShapeType.sptCSeal = 18;
			ShapeType.sptCArc = 19;
			ShapeType.sptCLine = 20;
			ShapeType.sptCPlaque = 21;
			ShapeType.sptCCan = 22;
			ShapeType.sptCDonut = 23;
			ShapeType.sptCTextSimple = 24;
			ShapeType.sptCTextOctagon = 25;
			ShapeType.sptCTextHexagon = 26;
			ShapeType.sptCTextCurve = 27;
			ShapeType.sptCTextWave = 28;
			ShapeType.sptCTextRing = 29;
			ShapeType.sptCTextOnCurve = 30;
			ShapeType.sptCTextOnRing = 31;
			ShapeType.sptCStraightConnector1 = 32;
			ShapeType.sptCBentConnector2 = 33;
			ShapeType.sptCBentConnector3 = 34;
			ShapeType.sptCBentConnector4 = 35;
			ShapeType.sptCBentConnector5 = 36;
			ShapeType.sptCCurvedConnector2 = 37;
			ShapeType.sptCCurvedConnector3 = 38;
			ShapeType.sptCCurvedConnector4 = 39;
			ShapeType.sptCCurvedConnector5 = 40;
			ShapeType.sptCCallout1 = 41;
			ShapeType.sptCCallout2 = 42;
			ShapeType.sptCCallout3 = 43;
			ShapeType.sptCAccentCallout1 = 44;
			ShapeType.sptCAccentCallout2 = 45;
			ShapeType.sptCAccentCallout3 = 46;
			ShapeType.sptCBorderCallout1 = 47;
			ShapeType.sptCBorderCallout2 = 48;
			ShapeType.sptCBorderCallout3 = 49;
			ShapeType.sptCAccentBorderCallout1 = 50;
			ShapeType.sptCAccentBorderCallout2 = 51;
			ShapeType.sptCAccentBorderCallout3 = 52;
			ShapeType.sptCRibbon = 53;
			ShapeType.sptCRibbon2 = 54;
			ShapeType.sptCChevron = 55;
			ShapeType.sptCPentagon = 56;
			ShapeType.sptCNoSmoking = 57;
			ShapeType.sptCStar8 = 58;
			ShapeType.sptCStar16 = 59;
			ShapeType.sptCStar32 = 60;
			ShapeType.sptCWedgeRectCallout = 61;
			ShapeType.sptCWedgeRoundRectCallout = 62;
			ShapeType.sptCWedgeEllipseCallout = 63;
			ShapeType.sptCWave = 64;
			ShapeType.sptCFoldedCorner = 65;
			ShapeType.sptCLeftArrow = 66;
			ShapeType.sptCDownArrow = 67;
			ShapeType.sptCUpArrow = 68;
			ShapeType.sptCLeftRightArrow = 69;
			ShapeType.sptCUpDownArrow = 70;
			ShapeType.sptCIrregularSeal1 = 71;
			ShapeType.sptCIrregularSeal2 = 72;
			ShapeType.sptCLightningBolt = 73;
			ShapeType.sptCHeart = 74;
			ShapeType.sptCFrame = 75;
			ShapeType.sptCQuadArrow = 76;
			ShapeType.sptCLeftArrowCallout = 77;
			ShapeType.sptCRightArrowCallout = 78;
			ShapeType.sptCUpArrowCallout = 79;
			ShapeType.sptCDownArrowCallout = 80;
			ShapeType.sptCLeftRightArrowCallout = 81;
			ShapeType.sptCUpDownArrowCallout = 82;
			ShapeType.sptCQuadArrowCallout = 83;
			ShapeType.sptCBevel = 84;
			ShapeType.sptCLeftBracket = 85;
			ShapeType.sptCRightBracket = 86;
			ShapeType.sptCLeftBrace = 87;
			ShapeType.sptCRightBrace = 88;
			ShapeType.sptCLeftUpArrow = 89;
			ShapeType.sptCBentUpArrow = 90;
			ShapeType.sptCBentArrow = 91;
			ShapeType.sptCStar24 = 92;
			ShapeType.sptCStripedRightArrow = 93;
			ShapeType.sptCNotchedRightArrow = 94;
			ShapeType.sptCBlockArc = 95;
			ShapeType.sptCSmileyFace = 96;
			ShapeType.sptCVerticalScroll = 97;
			ShapeType.sptCHorizontalScroll = 98;
			ShapeType.sptCCircularArrow = 99;
			ShapeType.sptCNotchedCircularArrow = 100;
			ShapeType.sptCUturnArrow = 101;
			ShapeType.sptCCurvedRightArrow = 102;
			ShapeType.sptCCurvedLeftArrow = 103;
			ShapeType.sptCCurvedUpArrow = 104;
			ShapeType.sptCCurvedDownArrow = 105;
			ShapeType.sptCCloudCallout = 106;
			ShapeType.sptCEllipseRibbon = 107;
			ShapeType.sptCEllipseRibbon2 = 108;
			ShapeType.sptCFlowChartProcess = 109;
			ShapeType.sptCFlowChartDecision = 110;
			ShapeType.sptCFlowChartInputOutput = 111;
			ShapeType.sptCFlowChartPredefinedProcess = 112;
			ShapeType.sptCFlowChartInternalStorage = 113;
			ShapeType.sptCFlowChartDocument = 114;
			ShapeType.sptCFlowChartMultidocument = 115;
			ShapeType.sptCFlowChartTerminator = 116;
			ShapeType.sptCFlowChartPreparation = 117;
			ShapeType.sptCFlowChartManualInput = 118;
			ShapeType.sptCFlowChartManualOperation = 119;
			ShapeType.sptCFlowChartConnector = 120;
			ShapeType.sptCFlowChartPunchedCard = 121;
			ShapeType.sptCFlowChartPunchedTape = 122;
			ShapeType.sptCFlowChartSummingJunction = 123;
			ShapeType.sptCFlowChartOr = 124;
			ShapeType.sptCFlowChartCollate = 125;
			ShapeType.sptCFlowChartSort = 126;
			ShapeType.sptCFlowChartExtract = 127;
			ShapeType.sptCFlowChartMerge = 128;
			ShapeType.sptCFlowChartOfflineStorage = 129;
			ShapeType.sptCFlowChartOnlineStorage = 130;
			ShapeType.sptCFlowChartMagneticTape = 131;
			ShapeType.sptCFlowChartMagneticDisk = 132;
			ShapeType.sptCFlowChartMagneticDrum = 133;
			ShapeType.sptCFlowChartDisplay = 134;
			ShapeType.sptCFlowChartDelay = 135;
			ShapeType.sptCTextPlain = 136;
			ShapeType.sptCTextStop = 137;
			ShapeType.sptCTextTriangle = 138;
			ShapeType.sptCTextTriangleInverted = 139;
			ShapeType.sptCTextChevron = 140;
			ShapeType.sptCTextChevronInverted = 141;
			ShapeType.sptCTextRingInside = 142;
			ShapeType.sptCTextRingOutside = 143;
			ShapeType.sptCTextArchUp = 144;
			ShapeType.sptCTextArchDown = 145;
			ShapeType.sptCTextCircle = 146;
			ShapeType.sptCTextButton = 147;
			ShapeType.sptCTextArchUpPour = 148;
			ShapeType.sptCTextArchDownPour = 149;
			ShapeType.sptCTextCirclePour = 150;
			ShapeType.sptCTextButtonPour = 151;
			ShapeType.sptCTextCurveUp = 152;
			ShapeType.sptCTextCurveDown = 153;
			ShapeType.sptCTextCascadeUp = 154;
			ShapeType.sptCTextCascadeDown = 155;
			ShapeType.sptCTextWave1 = 156;
			ShapeType.sptCTextWave2 = 157;
			ShapeType.sptCTextWave3 = 158;
			ShapeType.sptCTextWave4 = 159;
			ShapeType.sptCTextInflate = 160;
			ShapeType.sptCTextDeflate = 161;
			ShapeType.sptCTextInflateBottom = 162;
			ShapeType.sptCTextDeflateBottom = 163;
			ShapeType.sptCTextInflateTop = 164;
			ShapeType.sptCTextDeflateTop = 165;
			ShapeType.sptCTextDeflateInflate = 166;
			ShapeType.sptCTextDeflateInflateDeflate = 167;
			ShapeType.sptCTextFadeRight = 168;
			ShapeType.sptCTextFadeLeft = 169;
			ShapeType.sptCTextFadeUp = 170;
			ShapeType.sptCTextFadeDown = 171;
			ShapeType.sptCTextSlantUp = 172;
			ShapeType.sptCTextSlantDown = 173;
			ShapeType.sptCTextCanUp = 174;
			ShapeType.sptCTextCanDown = 175;
			ShapeType.sptCFlowChartAlternateProcess = 176;
			ShapeType.sptCFlowChartOffpageConnector = 177;
			ShapeType.sptCCallout90 = 178;
			ShapeType.sptCAccentCallout90 = 179;
			ShapeType.sptCBorderCallout90 = 180;
			ShapeType.sptCAccentBorderCallout90 = 181;
			ShapeType.sptCLeftRightUpArrow = 182;
			ShapeType.sptCSun = 183;
			ShapeType.sptCMoon = 184;
			ShapeType.sptCBracketPair = 185;
			ShapeType.sptCBracePair = 186;
			ShapeType.sptCStar4 = 187;
			ShapeType.sptCDoubleWave = 188;
			ShapeType.sptCActionButtonBlank = 189;
			ShapeType.sptCActionButtonHome = 190;
			ShapeType.sptCActionButtonHelp = 191;
			ShapeType.sptCActionButtonInformation = 192;
			ShapeType.sptCActionButtonForwardNext = 193;
			ShapeType.sptCActionButtonBackPrevious = 194;
			ShapeType.sptCActionButtonEnd = 195;
			ShapeType.sptCActionButtonBeginning = 196;
			ShapeType.sptCActionButtonReturn = 197;
			ShapeType.sptCActionButtonDocument = 198;
			ShapeType.sptCActionButtonSound = 199;
			ShapeType.sptCActionButtonMovie = 200;
			ShapeType.sptCHostControl = 201;
			ShapeType.sptCTextBox = 202;
			ShapeType.sptCChartPlus = 203;
			ShapeType.sptMax = 204;

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
				writer.WriteXmlNullableAttributeString(" ObjectType", getClientDataObjectType(this.m_oObjectType));
			}
		};
		CClientData.prototype.writeChildrenXml = function (writer) {
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
		CStroke.prototype.writeChildrenXml = function (writer) {
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
		CStroke.prototype.getOOXMLArrow = function (nArrowType, nArrowLen, nArrowWidth) {
			if(nArrowType !== null || nArrowLen !== null || nArrowWidth !== null) {
				let oArrow = new AscFormat.EndArrow();

				if( nArrowType === EStrokeArrowType.strokearrowtypeBlock) oArrow.type = oArrow.GetTypeCode("triangle");
				else if( nArrowType === EStrokeArrowType.strokearrowtypeClassic) oArrow.type = oArrow.GetTypeCode("stealth");
				else if ("none" === EStrokeArrowType.strokearrowtypeNone) oArrow.type = oArrow.GetTypeCode("none");
				else if ("open" === EStrokeArrowType.strokearrowtypeOpen) oArrow.type = oArrow.GetTypeCode("arrow");
				else if ("oval" === EStrokeArrowType.strokearrowtypeOval) oArrow.type = oArrow.GetTypeCode("oval");
				else oArrow.type = oArrow.GetTypeCode("none");

				if (nArrowLen === EStrokeArrowLength.strokearrowlengthLong) oArrow.len = oArrow.GetSizeCode("lg");
				else if (nArrowLen === EStrokeArrowLength.strokearrowlengthMedium) oArrow.len = oArrow.GetSizeCode("med");
				else if (nArrowLen === EStrokeArrowLength.strokearrowlengthShort) oArrow.len = oArrow.GetSizeCode("sm");
				else oArrow.len = oArrow.GetSizeCode("med");


				if (nArrowWidth === EStrokeArrowWidth.strokearrowwidthMedium) oArrow.w = oArrow.GetSizeCode("med");
				else if (nArrowWidth === EStrokeArrowWidth.strokearrowwidthNarrow) oArrow.w = oArrow.GetSizeCode("sm");
				else if (nArrowWidth === EStrokeArrowWidth.strokearrowwidthWide) oArrow.w = oArrow.GetSizeCode("lg");
				else oArrow.w = oArrow.GetSizeCode("med");
				return oArrow;
			}
			else {
				return null;
			}
		};
		CStroke.prototype.getOOXMLStroke = function() {
			if(this.m_oOn === false) {
				return AscFormat.CreateNoFillLine();
			}
			let oStroke = null;
			if(this.m_oColor !== null ||
				this.m_oDahsStyle !== null ||
			this.m_oEndArrow !== null || this.m_oEndArrowLength !== null || this.m_oEndArrowWidth !== null ||
			this.m_oStartArrow !== null || this.m_oStartArrowLength !== null || this.m_oStartArrowWidth !== null ||
				this.m_oEndCap !== null ||
			this.m_oJoinStyle !== null) {
				oStroke = new AscFormat.CLn();
				if(this.m_oColor) {
					oStroke.Fill = this.m_oColor.getOOXMLFill();
				}
				else {
					oStroke.Fill = AscFormat.CreateSolidFillRGB(0, 0, 0);
				}
				if(this.m_oDahsStyle !== null) {
					switch (this.m_oDahsStyle) {
						case EVmlDashStyle.vmldashstyleSolid: {oStroke.prstDash = oStroke.GetDashCode("solid"); break;}
						case EVmlDashStyle.vmldashstyleShortDash: {oStroke.prstDash = oStroke.GetDashCode("sysDash"); break;}
						case EVmlDashStyle.vmldashstyleShortDot: {oStroke.prstDash = oStroke.GetDashCode("sysDot"); break;}
						case EVmlDashStyle.vmldashstyleShortDashDot: {oStroke.prstDash = oStroke.GetDashCode("sysDashDot"); break;}
						case EVmlDashStyle.vmldashstyleShortDashDotDot: {oStroke.prstDash = oStroke.GetDashCode("sysDashDotDot"); break;}
						case EVmlDashStyle.vmldashstyleDot: {oStroke.prstDash = oStroke.GetDashCode("dot"); break;}
						case EVmlDashStyle.vmldashstyleDash: {oStroke.prstDash = oStroke.GetDashCode("dash"); break;}
						case EVmlDashStyle.vmldashstyleDashDot: {oStroke.prstDash = oStroke.GetDashCode("lgDash"); break;}
						case EVmlDashStyle.vmldashstyleLongDash: {oStroke.prstDash = oStroke.GetDashCode("dashDot"); break;}
						case EVmlDashStyle.vmldashstyleLongDashDot: {oStroke.prstDash = oStroke.GetDashCode("lgDashDot"); break;}
						case EVmlDashStyle.vmldashstyleLongDashDotDot: {oStroke.prstDash = oStroke.GetDashCode("lgDashDotDot"); break;}
						default: {oStroke.prstDash = oStroke.GetDashCode("solid"); break;}
					}
				}
				oStroke.tailEnd = this.getOOXMLArrow(this.m_oEndArrow, this.m_oEndArrowLength, this.m_oEndArrowWidth);
				oStroke.headEnd = this.getOOXMLArrow(this.m_oStartArrow, this.m_oStartArrowLength, this.m_oStartArrowWidth);
				if(this.m_oEndCap !== null) {
					switch (this.m_oEndCap) {
						case EStrokeEndCap.strokeendcapFlat: {oStroke.cap = oStroke.GetCapCode("flat"); break;}
						case EStrokeEndCap.strokeendcapRound: {oStroke.cap = oStroke.GetCapCode("rnd"); break;}
						case EStrokeEndCap.strokeendcapSqaure: {oStroke.cap = oStroke.GetCapCode("sq"); break;}
					}
				}
				if(this.m_oJoinStyle !== null) {
					oStroke.Join = new AscFormat.LineJoin();
					switch (this.m_oJoinStyle) {
						case EStrokeJoinStyle.strokejoinstyleRound: {oStroke.Join.type = AscFormat.LineJoinType.Round;}
						case EStrokeJoinStyle.strokejoinstyleMiter: {oStroke.Join.type = AscFormat.LineJoinType.Miter;}
						case EStrokeJoinStyle.strokejoinstyleBevel: {oStroke.Join.type = AscFormat.LineJoinType.Bevel;}
						default: {oStroke.Join.type = AscFormat.LineJoinType.Round;}
					}
				}
			}
			else {
				oStroke = AscFormat.CreateNoFillLine();
			}
			return oStroke;
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
			this.m_oTextStyle = null;
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
				while (reader.MoveToNextAttribute()) {
					if ("style" === reader.GetNameNoNS()) {
						this.m_oTextStyle = new CCssStyle(reader.GetValue());
					}
				}
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
		CTextbox.prototype.writeChildrenXml = function (writer) {

			if (this.m_oTxtbxContent !== null)
				this.m_oTxtbxContent.toXml(writer, "txbxContent");

			if (this.m_oText !== null) {
				writer.WriteXmlNodeStart("div");
				if(this.m_oTextStyle !== null) {
					writer.WriteXmlNullableAttributeString("style", this.m_oTextStyle.ToString());
				}
				writer.WriteXmlAttributesEnd();
				writer.WriteXmlString(this.m_oText);
				writer.WriteXmlNodeEnd("div");
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
			let wsChar = name.charAt(0);
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
		CGroup.prototype.writeChildrenXml = function (writer) {

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
		CGroup.prototype.convertToOOXML = function(aOtherItems, oOOXMLGroup, oContext) {
			let oGroup = new AscFormat.CGroupShape();
			let bIsTop = oContext.bIsTopDrawing;
			oContext.bIsTopDrawing = false;
			let aOldOtherElemensts = oContext.aOtherElements;
			oContext.aOtherElements = this.items;
			for(let nItem = 0; nItem < this.items.length; ++nItem) {
				let oItem = this.items[nItem];
				if(oItem instanceof CShape ||
				oItem instanceof CRect ||
				oItem instanceof COval ||
				oItem instanceof CLine ||
				oItem instanceof CPolyLine ||
				oItem instanceof CBackground ||
				oItem instanceof CRoundRect ||
				oItem instanceof CGroup) {
					let oOOXMLDrawing = oItem.convertToOOXML(this.items, oGroup, oContext);
					if(oOOXMLDrawing) {
						oGroup.addToSpTree(undefined, oOOXMLDrawing);
					}
				}
			}
			oContext.bIsTopDrawing = bIsTop;
			oContext.aOtherElements = aOldOtherElemensts;
			if(oGroup.getSpCount() > 0) {
				if(oOOXMLGroup) {
					oGroup.setGroup(oOOXMLGroup);
				}
				let oSpPr = new AscFormat.CSpPr();
				oGroup.setSpPr(oSpPr);
				let oXfrm = new AscFormat.CXfrm();
				oSpPr.setXfrm(oXfrm);
				let lCoordSizeW = 0, lCoordSizeH = 0;
				let lCoordOriginX = 0, lCoordOriginY = 0;
				if (this.m_oCoordSize) {
					lCoordSizeW = Emu_To_Mm(this.m_oCoordSize.x);
					lCoordSizeH = Emu_To_Mm(this.m_oCoordSize.y);
				}
				if (this.m_oCoordOrigin) {
					lCoordOriginX = Emu_To_Mm(this.m_oCoordOrigin.x);
					lCoordOriginY = Emu_To_Mm(this.m_oCoordOrigin.y);
				}

				let oProps = this.getMainProperties(oContext);
				if (bIsTop)
				{
					oXfrm.setOffX(0);
					oXfrm.setOffY(0);
				}
				else
				{
					oXfrm.setOffX(oProps.X);
					oXfrm.setOffY(oProps.Y);
				}

				oXfrm.setExtX(oProps.Width);
				oXfrm.setExtY(oProps.Height);
				oXfrm.setChOffX(lCoordOriginX);
				oXfrm.setChOffY(lCoordOriginY);
				oXfrm.setChExtX(lCoordSizeW);
				oXfrm.setChExtY(lCoordSizeH);

				this.convertFlipRot(oXfrm);
				let oNvPr = new AscFormat.UniNvPr();
				oGroup.setNvSpPr(oNvPr);
				oGroup.setBDeleted(false);
				return oGroup;
			}
			return null;
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
		CShapeDefaults.prototype.writeChildrenXml = function (writer) {

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
			return dValue * 36000 + 0.5 >> 0;
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
			return (dValue * 9525 + 0.5) >> 0;
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

			let sUnit = sValue.substring(sValue.length - 2);
			this.m_bUnit = true;

			if ("cm" === sUnit) {
				let dValue = parseFloat(sValue.substring(0, sValue.length - 2));
				this.m_dValue = Cm_To_Pt(dValue);
			} else if ("mm" === sUnit) {
				let dValue = parseFloat(sValue.substring(0, sValue.length - 2));
				this.m_dValue = Mm_To_Pt(dValue);
			} else if ("in" === sUnit) {
				let dValue = parseFloat(sValue.substring(0, sValue.length - 2));
				this.m_dValue = Inch_To_Pt(dValue);
			} else if ("pt" === sUnit) {
				this.m_dValue = parseFloat(sValue.substring(0, sValue.length - 2));
			} else if ("pc" === sUnit) {
				let dValue = parseFloat(sValue.substring(0, sValue.length - 2));
				this.m_dValue = dValue * 12.0;
			} else if ("pi" === sUnit) {
				let dValue = parseFloat(sValue.substring(0, sValue.length - 2));
				this.m_dValue = dValue * 12.0;
			} else {
				this.m_bUnit = false;
				this.m_dValue = parseFloat(sValue) / dKoef;

			}
			return this.m_dValue;
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
				this.m_dValue = parseFloat(sValue.substring(0, sValue.length - 1));
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
			return oItem;
		};
		CVMLDrawing.prototype.writeChildrenXml = function (writer) {
		};
		CVMLDrawing.prototype.getShape = function (nId) {
			let sId = "_x0000_s" + nId;
			return this.getShapeById(sId);
		};
		CVMLDrawing.prototype.getShapeById = function (sId) {
			for(let nItem = 0; nItem < this.items.length; ++nItem) {
				let oItem = this.items[nItem];
				if(oItem instanceof CShape) {
					if(oItem.m_sId === sId) {
						return oItem;
					}
				}
			}
			return null;
		};
		CVMLDrawing.prototype.getShapeBySpId = function (sId) {
			for(let nItem = 0; nItem < this.items.length; ++nItem) {
				let oItem = this.items[nItem];
				if(oItem instanceof CShape) {
					if(oItem.m_sSpId === sId) {
						return oItem;
					}
				}
			}
			return null;
		};
		CVMLDrawing.prototype.getXmlString = function() {
			if((!this.m_mapComments || isEmptyObject(this.m_mapComments)) && isEmptyObject(this.m_arObjectXml) && isEmptyObject(this.m_arControlXml))
				return "";

			let sXml = "";

			for (let i = 0; i < this.m_arObjectXml.length; ++i)
			{
				sXml += (this.m_arObjectXml[i]);
			}

			if (false === isEmptyObject(this.m_arControlXml) || ((null !== this.m_mapComments) && (false === isEmptyObject(this.m_mapComments))))
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
		CVMLDrawing.prototype.getSignatureLines = function() {
			let aSL = [];
			for(let nItem = 0; nItem < this.items.length; ++nItem) {
				let oItem = this.items[nItem];
				if(oItem.isSignatureLine()){
					aSL.push(oItem);
				}
			}
			return aSL;
		};
		CVMLDrawing.prototype.convertSignatureLines = function(oContext) {
			let aSL = this.getSignatureLines();
			let aOOXMLSl = [];
			for(let nSL = 0; nSL < aSL.length; ++nSL) {
				let oSL = aSL[nSL];
				let oOOXMLSL = oSL.convertToOOXML(this.items, null, oContext);
				aOOXMLSl.push(oOOXMLSL);
			}
			return aOOXMLSl;
		};


		function CLegacyDrawing(sType) {
			CVMLDrawing.call(this);
			this.type = sType;
			this.element = null;
			this.ole = null;

			this.dxaOrig = null;
			this.dyaOrig = null;
		}
		IC(CLegacyDrawing, CVMLDrawing, 0);
		CLegacyDrawing.prototype.readAttrXml = function (name, reader) {
			CVMLDrawing.prototype.readAttrXml.call(this, name, reader);
			if(name === "dxaOrig") {
				this.dxaOrig = reader.GetValueInt();
			}
			else if(name === "dyaOrig") {
				this.dyaOrig = reader.GetValueInt();
			}
		};
		CLegacyDrawing.prototype.readChildXml = function(name, reader) {
			let oItem = CVMLDrawing.prototype.readChildXml.call(this, name, reader);
			if(!oItem) {
				if(name === "OLEObject" || name === "objectEmbed") {
					oItem = new COLEObject();
					oItem.fromXml(reader);
					this.items.push(oItem);
				}
			}
		};
		CLegacyDrawing.prototype.getOLEObject = function() {
			for(let nItem = 0; nItem < this.items.length; ++nItem) {
				let oItem = this.items[nItem];
				if(oItem instanceof COLEObject) {
					return oItem;
				}
			}
			return null;
		};
		CLegacyDrawing.prototype.convertToDrawingML = function(reader) {
			let oOOXMLDrawing = null;
			let oContext = reader.context;
			oContext.reader = reader;
			oContext.bIsTopDrawing = true;
			oContext.aOtherElements = this.items;
			for(let nItem = 0; nItem < this.items.length; ++nItem) {
				let oItem = this.items[nItem];
				if(oItem instanceof CGroup) {
					oContext.sourceItem = oItem;
					oOOXMLDrawing = this.static_ConvertGroup(oItem, this.items, null, oContext);
				}
				else if(oItem instanceof CShadow) {
					oContext.sourceShadow = oItem;
					continue;
				}
				else if(oItem instanceof CShape){
					oContext.sourceItem = oItem;
					oOOXMLDrawing = this.static_ConvertShape(oItem, this.items, null, oContext);
				}
				else if(oItem instanceof CShapeType){
					continue;
				}
				else {
					if(oItem.convertToOOXML) {
						oContext.sourceItem = oItem;
						oOOXMLDrawing = this.static_ConvertShape(oItem, this.items, null, oContext);
					}
				}
				if(oOOXMLDrawing) {
					let oOleObject = this.getOLEObject();
					if(oOleObject && oContext.sourceItem.m_sId === oOleObject.m_sShapeId && oOOXMLDrawing instanceof AscFormat.CImageShape) {
						let oOOXMLImage = oOOXMLDrawing;
						oOOXMLDrawing = new AscFormat.COleObject();
						oOOXMLDrawing.setBDeleted(false);
						if(this.dxaOrig !== null && this.dyaOrig !== null) {
							oOOXMLDrawing.setPixSizes(this.dxaOrig, this.dyaOrig);
						}
						oOleObject.fillEditorOleObject(oOOXMLDrawing, oOOXMLImage,  reader);
					}
					else {
						oOOXMLDrawing.setBDeleted(false);
					}
					return oOOXMLDrawing;
				}
			}
			return oOOXMLDrawing;
		};
		CLegacyDrawing.prototype.GetDrawingMainProps = function(oParaDrawing, oReaderContext, oProps_)
		{
				let oNode, oCssStyles, oProps;
				oNode = oReaderContext.sourceItem;
				if(!oNode) {
					return;
				}
			oCssStyles = oNode.m_oStyle;
			if(!oCssStyles) {
				return;
			}
			let EM = Emu_To_Mm;
			 oProps = oProps_ || {IsTop: true};
				let pFind;

			let bIsInline = false;
			let bIsMargin = false;
			let bZIndex = false;
			let nZIndex = oCssStyles.GetZIndex();
			if (nZIndex !== null)
				bZIndex = true;
			
			let dLeft = oCssStyles.GetLeftInMM();
			let dMarginLeft = oCssStyles.GetMarginLeftInMM();
			let dTop = oCssStyles.GetTopInMM();
			let dMarginTop = oCssStyles.GetMarginTopInMM();

			if (oProps.IsTop === true)
			{
				if (dLeft === null && dMarginLeft === null && dTop === null && dMarginTop === null) 
				{
					bIsInline = true;
				}
				if (dMarginLeft !== null && dMarginTop !== null || dLeft !== null	&& dTop !== null)
				{
					bIsMargin = true;
				}
				pFind = oCssStyles.GetPropertyValueString("mso-position-horizontal-relative");
				if (null !== pFind && ((pFind === "text" && !bIsMargin) || pFind === "char"))
				{
					pFind = oCssStyles.GetPropertyValueString("mso-position-vertical-relative");
					if (null !== pFind && ((pFind === "text" && !bIsMargin) || pFind === "line"))
					{
						if (!bZIndex || !bIsMargin) //Liturgie Homberg 2017 mit Abendmahlsteil.docx
							bIsInline = true;
					}
				}

				if (!bIsInline)
				{
					pFind = oCssStyles.GetPropertyValueString("position");
					if (null !== pFind && pFind === "static")
					{
						bIsInline = true;
					}
				}
			}

			let parserPoint = new CPoint();
			let dKoef = 25.4 * 36000 / 72.0;
			let dKoefSize = oProps.IsTop ? dKoef : 1;
			//let dKoefSize = dKoef;

			let left	= 0;
			let top	= 0;
			let width	= 0;
			let height = 0;

			pFind = oCssStyles.GetPropertyValueString("polyline_correct");
			let bIsPolyCorrect = (null !== pFind);
			if (bIsPolyCorrect)
				dKoefSize = 1;

			if (!bIsInline)
			{
				pFind = oCssStyles.GetPropertyValueString("margin-left");
				if (null === pFind)
					pFind = oCssStyles.GetPropertyValueString("left");

				if (null !== pFind)
				{
					let oArray1 = pFind.split(",");

					for (let i = 0; i < oArray1.length; i++)
					{
						left += (dKoefSize * parserPoint.FromString(oArray1[i]) + 0.5) >> 0;
					}
				}

				pFind = oCssStyles.GetPropertyValueString("margin-top");

				if (null === pFind)
					pFind = oCssStyles.GetPropertyValueString("top");

				if (null !== pFind)
				{
					let oArray1 = pFind.split(",");
					for (let i = 0; i < oArray1.length; i++)
					{
						top += (dKoefSize * parserPoint.FromString(oArray1[i]) + 0.5) >> 0;
					}
				}
			}

			pFind = oCssStyles.GetPropertyValueString("width");
			if (null !== pFind)
			{
				width = (dKoefSize * parserPoint.FromString(pFind) + 0.5) >> 0;
			}
			else
			{
				pFind = oCssStyles.GetPropertyValueString("margin-right");
				if (null !== pFind)
					width = ((dKoefSize * parserPoint.FromString(pFind) + 0.5) >> 0) - left;
			}

			pFind = oCssStyles.GetPropertyValueString("height");
			if (null !== pFind)
			{
				height = (dKoefSize * parserPoint.FromString(pFind) + 0.5) >> 0;
			}
			else
			{
				pFind = oCssStyles.GetPropertyValueString("margin-bottom");
				if (null !== pFind)
					height = ((dKoefSize * parserPoint.FromString(pFind) + 0.5) >> 0) - top;
			}

			let margL = (9 * dKoef + 0.5) >> 0;
			let margT = 0;
			let margR = (9 * dKoef + 0.5) >> 0;
			let margB = 0;

			pFind = oCssStyles.GetPropertyValueString("mso-wrap-distance-left");
			if (null !== pFind)
				margL = (dKoef * parserPoint.FromString(pFind) + 0.5) >> 0;

			pFind = oCssStyles.GetPropertyValueString("mso-wrap-distance-top");
			if (null !== pFind)
				margT = (dKoef * parserPoint.FromString(pFind) + 0.5) >> 0;

			pFind = oCssStyles.GetPropertyValueString("mso-wrap-distance-right");
			if (null !== pFind)
				margR = (dKoef * parserPoint.FromString(pFind) + 0.5) >> 0;

			pFind = oCssStyles.GetPropertyValueString("mso-wrap-distance-bottom");
			if (null !== pFind)
				margB = (dKoef * parserPoint.FromString(pFind) + 0.5) >> 0;

			let rel_width = null;
			let rel_height = null;
			let rel_top = null;
			let rel_left = null;

			pFind = oCssStyles.GetPropertyValueString("mso-width-percent");
			if (null !== pFind)
			{
				rel_width = parserPoint.FromString(pFind) / 1000.;
			}
			pFind = oCssStyles.GetPropertyValueString("mso-height-percent");
			if (null !== pFind)
			{
				rel_height = parserPoint.FromString(pFind) / 1000.;
			}
			pFind = oCssStyles.GetPropertyValueString("mso-top-percent");
			if (null !== pFind)
			{
				rel_top = parserPoint.FromString(pFind) / 1000.;
			}
			pFind = oCssStyles.GetPropertyValueString("mso-left-percent");
			if (null !== pFind)
			{
				rel_left = parserPoint.FromString(pFind) / 1000.;
			}

			oProps.X		= EM(left);
			oProps.Y		= EM(top);
			oProps.Width	= EM(width);
			oProps.Height	= EM(height);
			if(!oParaDrawing) {
				return oProps;
			}


			let bExtendedSize = false;
			let shadow = oNode.getShadow();
			if (shadow)
			{
				if (shadow.m_oOn)
				{
					bExtendedSize = true;
				}

			}
			if (bIsInline)
			{
				oParaDrawing.Set_DrawingType(drawing_Inline);
				oParaDrawing.Set_Distance(EM(margL), EM(margT), EM(margR), EM(margB));
				oParaDrawing.setExtent(EM(width), EM(height));

				if(bExtendedSize) {
					oParaDrawing.setEffectExtent(EM(10795), EM(5080), EM(28575), EM(26670));
				}
				else {
					oParaDrawing.setEffectExtent(0, 0, 0, 0);
				}


				if (rel_width !== null)
				{
					oParaDrawing.SetSizeRelH({RelativeFrom : AscCommon.c_oAscSizeRelFromH.sizerelfromhPage, Percent : rel_width});
				}
				if (rel_height !== null)
				{
					oParaDrawing.SetSizeRelV({RelativeFrom : AscCommon.c_oAscSizeRelFromH.sizerelfromhPage, Percent : rel_width});
				}
				return;
			}
//------------------------------------------------------------------------------------


			oParaDrawing.Set_DrawingType(drawing_Anchor);
			oParaDrawing.Set_Distance(EM(margL), EM(margT), EM(margR), EM(margB));

			pFind = oCssStyles.GetPropertyValueString("z-index");
			let zIndex = null;

			if (null !== pFind)
			{
				zIndex = parserPoint.FromString(pFind);

				let zIndex_ = zIndex;// >= 0 ? *zIndex : -*zIndex;

				if (oReaderContext.maxZIndex === 0 && ((zIndex_ < 0xF000000 && zIndex_ > 0x80000) ||
					(zIndex_ > -0xF000000 && zIndex_ < -0x80000)))
				{
					zIndex_ = 0xF000000 - 0x80000 + zIndex_;
				}
				else
				{
					zIndex_ = Math.abs(zIndex_);
				}
				oParaDrawing.Set_RelativeHeight(zIndex_);
			}

			let isAllowInCell = null;
			let isAllowOverlap = null;
			if(oNode.m_oAllowInCell !== null) {
				isAllowInCell = oNode.m_oAllowInCell;
			}
			if(oNode.m_oAllowOverlap !== null) {
				isAllowOverlap = oNode.m_oAllowOverlap;
			}
			let bWrapPolygon = false;
			if(oNode.m_oWrapCoords) {
				let oWrapPolygon = oParaDrawing.wrappingPolygon;
				oWrapPolygon.setEdited(true);
				oWrapPolygon.setArrRelPoints(oNode.m_oWrapCoords.pts);
				bWrapPolygon = oNode.m_oWrapCoords.pts.length > 0;
			}
			if (zIndex !== null)
			{
				if (zIndex > 0)
				{
					oParaDrawing.Set_BehindDoc(false);
				}
			else if (zIndex < 0)
				{
					oParaDrawing.Set_BehindDoc(true);
				}
			}
			if (isAllowOverlap !== null)
			{
				oParaDrawing.Set_AllowOverlap(isAllowOverlap);
			}
			if (isAllowInCell !== null)
			{
				oParaDrawing.Set_LayoutInCell(isAllowInCell);
			}
			pFind = oCssStyles.GetPropertyValueString("mso-position-horizontal-relative");

			let nHRelativeFrom;
			let nVRelativeFrom;

			if (pFind !== null)
			{
				if ("char" === pFind)				nHRelativeFrom = Asc.c_oAscRelativeFromH.Character;
			else if ("page" === pFind)				nHRelativeFrom = Asc.c_oAscRelativeFromH.Page;
			else if ("margin" === pFind)			nHRelativeFrom = Asc.c_oAscRelativeFromH.Margin;
			else if ("left-margin-area" === pFind)	nHRelativeFrom = Asc.c_oAscRelativeFromH.LeftMargin;
			else if ("right-margin-area" === pFind)	nHRelativeFrom = Asc.c_oAscRelativeFromH.RightMargin;
			else if ("inner-margin-area" === pFind)	nHRelativeFrom = Asc.c_oAscRelativeFromH.InsideMargin;
			else if ("outer-margin-area" === pFind)	nHRelativeFrom = Asc.c_oAscRelativeFromH.OutsideMargin;
			else
				nHRelativeFrom = Asc.c_oAscRelativeFromH.Column;
			}
			else
			{
				nHRelativeFrom = Asc.c_oAscRelativeFromH.Column;
			}
			let strPosH = "absolute";
			pFind = oCssStyles.GetPropertyValueString("mso-position-horizontal");
			if (null !== pFind)
				strPosH = pFind;

			if (strPosH === "absolute")
			{
				if (rel_left !== null)
				{
					oParaDrawing.Set_PositionH(nHRelativeFrom, false, EM(rel_left), true);
				}
				else
				{
					oParaDrawing.Set_PositionH(nHRelativeFrom , false , EM(left), false);
				}
			}
		else
			{
				switch (strPosH) {
					case "left": {
						oParaDrawing.Set_PositionH(nHRelativeFrom , true, Asc.c_oAscAlignH.Left, false);
						break;
					}
					case "right": {
						oParaDrawing.Set_PositionH(nHRelativeFrom , true, Asc.c_oAscAlignH.Right, false);
						break;
					}
					case "center": {
						oParaDrawing.Set_PositionH(nHRelativeFrom , true, Asc.c_oAscAlignH.Center, false);
						break;
					}
					case "inside": {
						oParaDrawing.Set_PositionH(nHRelativeFrom , true, Asc.c_oAscAlignH.Inside, false);
						break;
					}
					case "outside": {
						oParaDrawing.Set_PositionH(nHRelativeFrom , true, Asc.c_oAscAlignH.Outside, false);
						break;
					}
				}
			}
			pFind = oCssStyles.GetPropertyValueString("mso-position-vertical-relative");
			if (pFind !== null)
			{
				if ("margin" === pFind)				nVRelativeFrom = Asc.c_oAscRelativeFromV.Margin;
			else if ("text" === pFind)					nVRelativeFrom = Asc.c_oAscRelativeFromV.Paragraph;
			else if ("page" === pFind)					nVRelativeFrom = Asc.c_oAscRelativeFromV.Page;
			else if ("top-margin-area" === pFind)		nVRelativeFrom = Asc.c_oAscRelativeFromV.TopMargin;
			else if ("bottom-margin-area" === pFind)	nVRelativeFrom = Asc.c_oAscRelativeFromV.BottomMargin;
			else if ("inner-margin-area" === pFind)		nVRelativeFrom = Asc.c_oAscRelativeFromV.InsideMargin;
			else if ("outer-margin-area" === pFind)		nVRelativeFrom = Asc.c_oAscRelativeFromV.OutsideMargin;
			else
				nVRelativeFrom = Asc.c_oAscRelativeFromV.Line;
			}
			else
			{
				nVRelativeFrom = Asc.c_oAscRelativeFromV.Paragraph;
			}

			let strPosV = "absolute";
			pFind = oCssStyles.GetPropertyValueString("mso-position-vertical");
			if (null !== pFind)
				strPosV = pFind;

			if (strPosV === "absolute")
			{
				if (rel_left !== null)
				{
					oParaDrawing.Set_PositionV(nVRelativeFrom, false, EM(rel_top), true);
				}
				else
				{
					oParaDrawing.Set_PositionV(nVRelativeFrom , false , EM(top), false);
				}
			}
		else
			{

				switch (strPosV) {
					case "top": {
						oParaDrawing.Set_PositionV(nVRelativeFrom , true, Asc.c_oAscAlignV.Top, false);
						break;
					}
					case "bottom": {
						oParaDrawing.Set_PositionV(nVRelativeFrom , true, Asc.c_oAscAlignV.Bottom, false);
						break;
					}
					case "center": {
						oParaDrawing.Set_PositionV(nVRelativeFrom , true, Asc.c_oAscAlignV.Center, false);
						break;
					}
					case "inside": {
						oParaDrawing.Set_PositionV(nVRelativeFrom , true, Asc.c_oAscAlignV.Inside, false);
						break;
					}
					case "outside": {
						oParaDrawing.Set_PositionV(nVRelativeFrom , true, Asc.c_oAscAlignV.Outside, false);
						break;
					}
				}
			}

			oParaDrawing.setExtent(EM(width), EM(height));
			if(bExtendedSize) {
				oParaDrawing.setEffectExtent(EM(10795), EM(5080), EM(28575), EM(26670));
			}
			else {
				oParaDrawing.setEffectExtent(0, 0, 0, 0);
			}

			let oWrap = oNode.getWrap();
			if (oWrap)
			{
				if (oWrap.m_oType === EWrapType.wraptypeNone || oWrap.m_oType === null) {
					oParaDrawing.Set_WrappingType(WRAPPING_TYPE_NONE);
				}
				else if (oWrap.m_oType === EWrapType.wraptypeSquare) {
					oParaDrawing.Set_WrappingType(WRAPPING_TYPE_SQUARE);
				}
				else if (oWrap.m_oType === EWrapType.wraptypeTopAndBottom) {
					oParaDrawing.Set_WrappingType(WRAPPING_TYPE_TOP_AND_BOTTOM);
				}
				else if (oWrap.m_oType === EWrapType.wraptypeTight)
				{
					oParaDrawing.Set_WrappingType(WRAPPING_TYPE_TIGHT);
				}
			else if (oWrap.m_oType === EWrapType.wraptypeThrough)
				{
					oParaDrawing.Set_WrappingType(WRAPPING_TYPE_THROUGH);
				}
			else {

					oParaDrawing.Set_WrappingType(WRAPPING_TYPE_SQUARE);
				}
			}
			else
			{
				oParaDrawing.Set_WrappingType(WRAPPING_TYPE_NONE);
			}
			let bHidden = false;
			pFind = oCssStyles.GetPropertyValueString("visibility");
			if (null !== pFind)
			{
				if ("hidden" === pFind)
					bHidden = true;
			}
			if (rel_width !== null)
			{
				oParaDrawing.SetSizeRelH({RelativeFrom: nHRelativeFrom, Percent: rel_width});
			}
			if (rel_height !== null)
			{
				oParaDrawing.SetSizeRelV({RelativeFrom: nVRelativeFrom, Percent: rel_height});
			}
			oParaDrawing.docPr.setIsHidden(bHidden);
			return oProps;
		}


		CLegacyDrawing.prototype.static_GetShapeTypeForShape = function(oShape, aOtherItems) {
			if(oShape instanceof CShape) {
				return null;
			}
			let nType = oShape.getShapeType();
			if(nType !== null) {
				for(let nItem = 0; nItem < aOtherItems.length; ++nItem) {
					let oItem = aOtherItems[nItem];
					if(oItem instanceof CShapeType) {
						if(oItem.getShapeType() === nType) {
							return oItem;
						}
					}
				}
			}
			return null;
		};
		CLegacyDrawing.prototype.static_ConvertShape = function(oElement, aOtherItems, oOOXMLGroup, oContext) {
			return oElement.convertToOOXML(aOtherItems, oOOXMLGroup, oContext);
			// let nShapeType = null;
			// let bSignatureLine = false;
			// let bTextBox = false;
			// let bHidden = false;
			//
			// if(oElement instanceof CBackground ||
			// oElement instanceof CRect ||
			// oElement instanceof CRoundRect ||
			// oElement instanceof COval ||
			// oElement instanceof CImage ||
			// oElement instanceof CLine ||
			// oElement instanceof CPolyLine) {
			// 	oElement.convertToOOXML(oOOXMLGroup);
			// }
			//
			// if(oElement instanceof CBackground ) {
			// 	nShapeType = ShapeType.sptCRect;
			// }
			// if(oElement instanceof CRect) {
			// 	nShapeType = ShapeType.sptCRect;
			// }
			// if(oElement instanceof CRoundRect) {
			// 	nShapeType = ShapeType.sptCRoundRect;
			// 	//TODO: adjustment
			// }
			// if(oElement instanceof COval) {
			// 	nShapeType = ShapeType.sptCEllipse;
			// }
			// if(oElement instanceof CImage) {
			// 	nShapeType = ShapeType.sptCFrame;
			// }
			// if(oElement instanceof CLine) {
			// 	nShapeType = ShapeType.sptCLine;
			// 	//TODO: creates advanced css style with coordinates
			// }
			// if(oElement instanceof CPolyLine) {
			// 	nShapeType = ShapeType.sptCustom;
			// 	//TODO: creates advanced css style with coordinates
			// }
			// let oShapeType = null;
			// if(oElement instanceof CShape) {
			// 	let oShapeType = this.static_GetShapeTypeForShape(oElement, aOtherItems);
			// 	if(oShapeType) {
			// 		//TODO: copy properties from shapetype
			// 		nShapeType = oShapeType.getShapeType();
			// 	}
			// 	else {
			// 		if(oElement.getShapeType() === ShapeType.sptCFrame) {
			// 			nShapeType = ShapeType.sptCFrame;
			// 		}
			// 	}
			// }
			//
			// if(nShapeType === null) {
			// 	let sConnectorType = oElement.m_oConnectorType;
			// 	if (sConnectorType === "elbow") nShapeType = ShapeType.sptCBentConnector2;
			// 	else if (sConnectorType === "straight")	nShapeType = ShapeType.sptCStraightConnector1;
			// 	else if (sConnectorType === "curved") nShapeType = ShapeType.sptCCurvedConnector2;
			// }
			// //picture or shape
			// let oGraphicObject = null;
			// if(nShapeType !== null) {
			// 	let oSpPr = new AscFormat.CSpPr();
			// 	let oUniNvPr = new AscFormat.UniNvPr();
			// 	let oCNvPr = oUniNvPr.cNvPr;
			// 	if(nShapeType === ShapeType.sptCFrame && !bSignatureLine) {
			// 		oGraphicObject = new AscFormat.CImageShape();
			// 		oGraphicObject.setSpPr(oSpPr);
			// 		oGraphicObject.setNvPicPr(oUniNvPr);
			//
			// 	}
			// 	else {
			// 		oGraphicObject = new AscFormat.CShape();
			// 		oGraphicObject.setSpPr(oSpPr);
			// 		oGraphicObject.setNvSpPr(oUniNvPr);
			// 	}
			// 	//Fill
			// 	if(oElement.m_oFilled === false) {
			// 		oSpPr.setFill(AscFormat.CreateNoFillUniFill());
			// 	}
			// 	else if(oElement.m_oFillColor) {
			// 		let oC = oElement.m_oFillColor;
			// 		oSpPr.setFill(AscFormat.CreateSolidFillRGBA(oC.r, oC.g, oC.b, oC.a))
			// 	}
			//
			// }
			// return oGraphicObject;

		};
		CLegacyDrawing.prototype.static_ConvertGroup = function(oElement, aOtherItems, oOOXMLGroup, oContext) {
			return oElement.convertToOOXML(aOtherItems, oOOXMLGroup, oContext);
		};
		CLegacyDrawing.prototype.static_ConvertOle = function(oElement, oContext) {
			return null;
		};
		let LD_PROTOTYPE = CLegacyDrawing.prototype;

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
			let oRGBA = AscCommon.RgbaHexToRGBA(sVal);
			this.r = oRGBA.R;
			this.g = oRGBA.G;
			this.b = oRGBA.B;
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
			else if (sVal.indexOf("fill") > -1) {
				this.type = EColorType.colortypeRGB;

				let sColorEffect = sVal;
				let sColor = sVal;
				if (sColorEffect.length > 5)
					sColorEffect = sColorEffect.substring(5);

				let resR, resG, resB;

				resR = this.r;
				resG = this.g;
				resB = this.b;

				let param = 0;
				let pos1 = sColor.indexOf('(');
				let pos2 = sColor.indexOf(')');
				if (pos1 === -1 || pos2 === -1)
					return;
				if (pos2 < (pos1 + 2))
					return;

				let s = sColor.substring(pos1 + 1, pos2);
				param = parseInt(s);
				let isEffect = false;

				if (0 === sColorEffect.indexOf("darken")) {
					resR = (this.r * param / 255);
					resG = (this.g * param / 255);
					resB = (this.b * param / 255);
					isEffect = true;
				}
				else if (0 === sColorEffect.indexOf("lighten")) {
					resR = 255 - ((255 - this.r) * param / 255);
					resG = 255 - ((255 - this.g) * param / 255);
					resB = 255 - ((255 - this.b) * param / 255);
					isEffect = true;
				}
				else if (0 === sColorEffect.indexOf("add")) {
					resR = this.r + param;
					resG = this.g + param;
					resB = this.b + param;
					isEffect = true;
				}
				else if (0 === sColorEffect.indexOf("subtract")) {
					resR = this.r - param;
					resG = this.g - param;
					resB = this.b - param;
					isEffect = true;
				}
				else if (0 === sColorEffect.indexOf("reversesubtract")) {
					resR = param - this.r;
					resG = param - this.g;
					resB = param - this.b;
					isEffect = true;
				}
				else if (0 === sColorEffect.indexOf("blackwhite")) {
					resR = (this.r < param) ? 0 : 255;
					resG = (this.g < param) ? 0 : 255;
					resB = (this.b < param) ? 0 : 255;
					isEffect = true;
				}

				if (isEffect) {
					resR = (resR < 0) ? 0 : resR;
					resR = (resR > 255) ? 255 : resR;

					resG = (resG < 0) ? 0 : resG;
					resG = (resG > 255) ? 255 : resG;

					resB = (resB < 0) ? 0 : resB;
					resB = (resB > 255) ? 255 : resB;
				}
				this.r = resR;
				this.g = resG;
				this.b = resB;
			}
			else if (sVal.indexOf("[") > -1 && sVal.indexOf("]") > -1) {
				let p1 = sVal.indexOf("[");
				let p2 = sVal.indexOf("]");
				let sIndex = p2 > p1 ? sVal.substring(p1 + 1, p2) : "";

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
		CColor.prototype.getOOXMLColor = function() {
			return AscFormat.CreateUniColorRGB(this.r, this.g, this.b);
			//todo: fill
			// if (sColor2->find("fill") != -1)
			// {
			// 	std::wstring sColorEffect = *sColor2;
			// 	if (sColorEffect.length() > 5)
			// 		sColorEffect = sColorEffect.substr(5);
			//
			// 	int resR, resG, resB;
			// 	GetColorWithEffect(sColorEffect, R, G, B, resR, resG, resB);
			//
			// 	Gs_.color.Color->SetRGB(resR, resG, resB);
		};
		CColor.prototype.getOOXMLFill = function () {
			return AscFormat.CreateUniFillByUniColor(this.getOOXMLColor());
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
				strX = sValue.substring(0, nPos);
				strY = sValue.substring(nPos + 1);
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
				let nMidPos = sValue.indexOf(" ", nStartPos);
				let nEndPos = sValue.indexOf(" ", nMidPos + 1);

				if (-1 === nMidPos)
					break;

				if (-1 === nEndPos)
					nEndPos = nLen;

				let strX = sValue.substring(nStartPos, nMidPos);
				let strY = sValue.substring(nMidPos + 1, nEndPos);

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
				let strValue = sValue.substring(nStartPos, nEndPos);
				this.m_dSxx = strValue.length === 0 ? 0 : parseFloat(strValue);
			}

			// Sxy
			nStartPos = nEndPos + 1;
			nEndPos = sValue.indexOf(",", nStartPos);
			if (-1 === nEndPos)
				nEndPos = nLen;

			if (nEndPos - nStartPos > 0) {
				let strValue = sValue.substring(nStartPos, nEndPos);
				this.m_dSxy = strValue.length === 0 ? 0 : parseFloat(strValue);
			}

			// Syx
			nStartPos = nEndPos + 1;
			nEndPos = sValue.indexOf(",", nStartPos);
			if (-1 === nEndPos)
				nEndPos = nLen;

			if (nEndPos - nStartPos > 0) {
				let strValue = sValue.substring(nStartPos, nEndPos);
				this.m_dSyx = strValue.length === 0 ? 0 : parseFloat(strValue);
			}

			// Syy
			nStartPos = nEndPos + 1;
			nEndPos = sValue.indexOf(",", nStartPos);
			if (-1 === nEndPos)
				nEndPos = nLen;

			if (nEndPos - nStartPos > 0) {
				let strValue = sValue.substring(nStartPos, nEndPos);
				this.m_dSyy = strValue.length === 0 ? 0 : parseFloat(strValue);
			}

			// Px
			nStartPos = nEndPos + 1;
			nEndPos = sValue.indexOf(",", nStartPos);
			if (-1 === nEndPos)
				nEndPos = nLen;

			if (nEndPos - nStartPos > 0) {
				let strValue = sValue.substring(nStartPos, nEndPos);
				this.m_dPx = strValue.length === 0 ? 0 : parseFloat(strValue);
			}

			// Py
			nStartPos = nEndPos + 1;
			nEndPos = sValue.indexOf(",", nStartPos);
			if (-1 === nEndPos)
				nEndPos = nLen;

			if (nEndPos - nStartPos > 0) {
				let strValue = sValue.substring(nStartPos, nEndPos);
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
				let strValue = sValue.substring(0, nLen - 1);
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

			let sTemp = sValue.substring(0, nPos);
			if (-1 !== sTemp.indexOf('%')) {
				let oPerc = new CPercentage(sTemp);
				this.m_dX = oPerc.GetValue();
				this.m_bUnitsX = false;
			} else {
				let oPt = new CPoint(sTemp);
				this.m_dX = oPt.GetValue();
				this.m_bUnitsX = true;
			}

			sTemp = sValue.substring(nPos + 1);
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
				let strValue = sValue.substring(0, nLen - 1);
				let nValue = strValue.length === 0 ? 0 : parseInt(strValue);

				this.SetValue(nValue);
			} else if (bPercentage) {
				let strValue = sValue.substring(0, nLen - 1);
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
			let strX = sValue.substring(0, nPos);
			strX.replace("@", "");

			this.m_nX = strX.length === 0 ? 0 : parseInt(strX);

			let nPos2 = sValue.indexOf(",", nPos + 1);
			if (-1 === nPos2) {// only x, y position
				let strY = sValue.substring(nPos + 1);
				strY.replace("@", "");
				this.m_nY = strY.length === 0 ? 0 : parseInt(strY);
				return 0;
			}

			let strY = sValue.substring(nPos + 1, nPos2);
			let strZ = sValue.substring(nPos2 + 1);

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
			let strX = sValue.substring(0, nPos);
			strX.replace("@", "");

			this.m_nX = strX.length === 0 ? 0 : parseInt(strX);

			let nPos2 = sValue.indexOf(",", nPos + 1);
			if (-1 === nPos2) {// only x, y position
				let strY = sValue.substring(nPos + 1);
				strY.replace("@", "");
				this.m_nY = strY.length === 0 ? 0 : parseInt(strY);
				return 0;
			}

			let strY = sValue.substring(nPos + 1, nPos2);
			let strZ = sValue.substring(nPos2 + 1);

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
				strX = sValue.substring(0, nPos);
				strY = sValue.substring(nPos + 1);
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

			let strX, strY = "";
			if (-1 === nPos) {//only x coord
				strX = sValue;
			} else {
				strX = sValue.substring(0, nPos);
				strY = sValue.substring(nPos + 1);
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
				let strValue = sValue.substring(0, nLen - 1);
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

				let oPoint = new CPoint(sValue.substring(nPos, nSpacePos));
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
				case "gradientradial":
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

			let sFirst = sValue.substring(0, nPos);
			let sSecond = sValue.substring(nPos + 1);

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
						sId = sValue.substring(1);
						break;
					case '#':
						eValue = EVml_Vector2D_Position.vmlvector2dposAdjValue;
						sId = sValue.substring(1);
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

				let strX = sValue.substring(nStartPos, nMidPos);
				let strY = sValue.substring(nMidPos + 1, nEndPos);

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
		CVml_Polygon2D_Units.prototype.ToSVGPath = function() {
			let _x, _y, _r, _b;
			let oPt;
			oPt = this.m_arrPoints[0];
			if(!oPt) {
				return null;
			}
			_x = oPt.nX;
			_y = oPt.nY;
			_r = oPt.nX;
			_b = oPt.nY;
			for(let nPt = 1; nPt < this.m_arrPoints.length; ++nPt) {
				let oPt = this.m_arrPoints[nPt];
				_x = Math.min(_x, oPt.nX);
				_y = Math.min(_y, oPt.nY);
				_r = Math.max(_r, oPt.nX);
				_b = Math.max(_b, oPt.nY);
			}
			let dKoefX = 21600.0 / Math.max((_r - _x), 1);
			let dKoefY = 21600.0 / Math.max((_b - _y), 1);
			let strPath = "";
			for(let nPt = 0; nPt < this.m_arrPoints.length; ++nPt) {
				let oPt = this.m_arrPoints[nPt];
				let _s;
				if (nPt === 0) _s = "m";
				else _s = "l";
				strPath += _s + (dKoefX * (oPt.nX - _x)) + "," + (dKoefY * (oPt.nY - _y));
			}

			strPath += "e";
			return {path: strPath, bounds: {l: _x, t: _y, r: _r, b: _b}};
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

		CVml_Polygon2D.prototype.AddPolet = function (nX, nY) {
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

				let strX = sValue.substring(nStartPos, nMidPos);
				let strY = sValue.substring(nMidPos + 1, nEndPos);

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


			let oPt1 = new CPoint(sValue.substring(0, nPos));
			this.m_dX = oPt1.GetValue();


			let oPt2 = new CPoint(sValue.substring(nPos + 1));
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
			sValue.replace(" ", "");
			arSplit = sValue.split(",");

			if (arSplit.length > 0) {
				let oPt = new CPoint(arSplit[0]);
				this.m_dLeft = oPt.GetValue();
			}

			if (arSplit.length > 1) {
				let oPt = new CUniversalMeasure();
				oPt.Parse(arSplit[1], 1)
				this.m_dTop = oPt.GetValue();
			}

			if (arSplit.length > 2) {
				let oPt = new CUniversalMeasure();
				oPt.Parse(arSplit[2], 1)
				this.m_dRight = oPt.GetValue();
			}

			if (arSplit.length > 3) {
				let oPt = new CUniversalMeasure();
				oPt.Parse(arSplit[3], 1)
				this.m_dBottom = oPt.GetValue();
			}

			return 0;
		};
		CVml_TextBoxInset.prototype.ToString = function () {
			return "" + this.m_dLeft + "pt," + this.m_dTop + "pt," + this.m_dRight + "pt," + this.m_dBottom + "pt";
		};


		function CorrectXmlString2(strText) {
			strText = strText.replace(new RegExp("&apos;", 'g'), "'");
			strText = strText.replace(new RegExp("&lt;", 'g'), "<");
			strText = strText.replace(new RegExp("&gt;", 'g'), ">");
			strText = strText.replace(new RegExp("&quot;", 'g'), "\"");
			strText = strText.replace(new RegExp("&amp;", 'g'), "&");
			return strText;
		}

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

			let sValue_ = CorrectXmlString2(sValue);
			this.m_sCss = sValue_;
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
					this.m_arrProperties.push(oProperty);
					sTemp = "";
				} else {
					let oProperty = new CCssProperty(sTemp.substring(0, nPos));
					this.m_arrProperties.push(oProperty);
					sTemp = sTemp.substring(nPos + 1);
				}
			}
			return true;
		};
		CCssStyle.prototype.GetProperty = function(nType) {
			for(let nPr = 0; nPr < this.m_arrProperties.length; ++nPr) {
				if(this.m_arrProperties[nPr].m_eType === nType) {
					return this.m_arrProperties[nPr];
				}
			}
			return null;
		};
		
		CCssStyle.prototype.GetPropertyByStringType = function(sType) {
			for(let nPr = 0; nPr < this.m_arrProperties.length; ++nPr) {
				if(this.m_arrProperties[nPr].m_sType === sType) {
					return this.m_arrProperties[nPr];
				}
			}
			return null;
		};
		CCssStyle.prototype.GetPropertyValueString = function(sType) {
			let oPr = this.GetPropertyByStringType(sType);
			if(oPr) {
				return oPr.m_sValue;
			}
			return null;
		};
		CCssStyle.prototype.GetZIndex = function() {
			let oPr = this.GetProperty(ECssPropertyType.cssptZIndex);
			if(oPr === null) {
				return null;
			}

			let oZIndex = oPr && oPr.m_oValue && oPr.m_oValue.oZIndex;
			if(oZIndex && oZIndex && AscFormat.isRealNumber(oZIndex.nOrder)) {
				return oZIndex.nOrder;
			}
		};


		CCssStyle.prototype.GetNumberValueInMM = function(nType) {
			let dNumVal = this.GetNumberValue(nType)
			if(AscFormat.isRealNumber(dNumVal)) {
				return Pt_To_Mm(dNumVal/36000);
			}
			return null;
		};
		CCssStyle.prototype.GetNumberValue = function(nType) {
			let oPr = this.GetProperty(nType);
			if(oPr === null) {
				return null;
			}
			let oValue = oPr && oPr.m_oValue && oPr.m_oValue.oValue;
			if(oValue && AscFormat.isRealNumber(oValue.dValue)) {
				return oValue.dValue;
			}
			return null;
		};
		CCssStyle.prototype.GetStringValue = function(nType) {
			let oPr = this.GetProperty(nType);
			if(oPr === null) {
				return null;
			}
			let oValue = oPr && oPr.m_oValue;
			if(oValue && oValue.wsValue) {
				return oValue.wsValue;
			}
			return null;
		};

		CCssStyle.prototype.GetMarginLeftInMM = function() {
			return this.GetNumberValueInMM(ECssPropertyType.cssptMarginLeft);
		};
		CCssStyle.prototype.GetMarginTopInMM = function() {
			return this.GetNumberValueInMM(ECssPropertyType.cssptMarginTop);
		};
		CCssStyle.prototype.GetMarginRightInMM = function() {
			return this.GetNumberValueInMM(ECssPropertyType.cssptMarginRight);
		};
		CCssStyle.prototype.GetMarginBottomInMM = function() {
			return this.GetNumberValueInMM(ECssPropertyType.cssptMarginBottom);
		};
		CCssStyle.prototype.GetLeftInMM = function() {
			return this.GetNumberValueInMM(ECssPropertyType.cssptLeft);
		};
		CCssStyle.prototype.GetTopInMM = function() {
			return this.GetNumberValueInMM(ECssPropertyType.cssptTop);
		};
		CCssStyle.prototype.GetWidthInMM = function() {
			return this.GetNumberValueInMM(ECssPropertyType.cssptWidth);
		};
		CCssStyle.prototype.GetHeightInMM = function() {
			return this.GetNumberValueInMM(ECssPropertyType.cssptHeight);
		};
		CCssStyle.prototype.GetFontStyle = function() {
			return this.GetStringValue(ECssPropertyType.cssptFont);
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
			this.m_sType = "unknown";
			this.m_sValue = "";
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
				let sProperty = sBuffer.substring(0, nPos);
				sValue = sBuffer.substring(nPos + 1);

				sProperty = sProperty.replace(/\s/g, "");

				if (sProperty.length <= 2) {
					this.m_eType = ECssPropertyType.cssptUnknown;
					return;
				}
				this.m_sType = sProperty;
				this.m_sValue = sValue;

				if ("direction" === sProperty) this.m_eType = ECssPropertyType.cssptDirection;
				else if ("flip" === sProperty) this.m_eType = ECssPropertyType.cssptFlip;
				else if ("font" === sProperty) this.m_eType = ECssPropertyType.cssptFont;
				else if ("font-family" === sProperty) this.m_eType = ECssPropertyType.cssptFontFamily;
				else if ("font-size" === sProperty) this.m_eType = ECssPropertyType.cssptFontSize;
				else if ("font-style" === sProperty) this.m_eType = ECssPropertyType.cssptFontStyle;
				else if ("font-variant" === sProperty) this.m_eType = ECssPropertyType.cssptFontVariant;
				else if ("font-weight" === sProperty) this.m_eType = ECssPropertyType.cssptFontWeight;
				else if ("height" === sProperty) this.m_eType = ECssPropertyType.cssptHeight;
				else if ("layout-flow" === sProperty) this.m_eType = ECssPropertyType.cssptLayoutFlow;
				else if ("left" === sProperty) this.m_eType = ECssPropertyType.cssptLeft;
				else if ("margin-bottom" === sProperty) this.m_eType = ECssPropertyType.cssptMarginBottom;
				else if ("margin-left" === sProperty) this.m_eType = ECssPropertyType.cssptMarginLeft;
				else if ("margin-right" === sProperty) this.m_eType = ECssPropertyType.cssptMarginRight;
				else if ("margin-top" === sProperty) this.m_eType = ECssPropertyType.cssptMarginTop;
				else if ("mso-direction-alt" === sProperty) this.m_eType = ECssPropertyType.cssptMsoDirectionAlt;
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
				else if ("position" === sProperty) this.m_eType = ECssPropertyType.cssptPosition;
				else if ("rotation" === sProperty) this.m_eType = ECssPropertyType.cssptRotation;
				else if ("text-decoration" === sProperty) this.m_eType = ECssPropertyType.cssptTextDecoration;
				else if ("top" === sProperty) this.m_eType = ECssPropertyType.cssptTop;
				else if ("text-align" === sProperty) this.m_eType = ECssPropertyType.cssptHTextAlign;
				else if ("visibility" === sProperty) this.m_eType = ECssPropertyType.cssptVisibility;
				else if ("v-rotate-letters" === sProperty) this.m_eType = ECssPropertyType.cssptVRotateLetters;
				else if ("v-same-letter-heights" === sProperty) this.m_eType = ECssPropertyType.cssptVSameLetterHeights;
				else if ("v-text-align" === sProperty) this.m_eType = ECssPropertyType.cssptVTextAlign;
				else if ("v-text-anchor" === sProperty) this.m_eType = ECssPropertyType.cssptVTextAnchor;
				else if ("v-text-kern" === sProperty) this.m_eType = ECssPropertyType.cssptVTextKern;
				else if ("v-text-reverse" === sProperty) this.m_eType = ECssPropertyType.cssptVTextReverse;
				else if ("v-text-spacing-mode" === sProperty) this.m_eType = ECssPropertyType.cssptVTextSpacingMode;
				else if ("v-text-spacing" === sProperty) this.m_eType = ECssPropertyType.cssptVTextSpacing;
				else if ("width" === sProperty) this.m_eType = ECssPropertyType.cssptWidth;
				else if ("z-index" === sProperty) this.m_eType = ECssPropertyType.cssptZIndex;

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

				let strValue = sValue.substring(0, nPos);
				let dValue = parseFloat(sValue);

				this.m_oValue.oValue.dValue = Inch_To_Pt(dValue);
			} else if (-1 !== (nPos = sValue.indexOf("cm"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeUnits;

				let strValue = sValue.substring(0, nPos);
				let dValue = parseFloat(sValue);

				this.m_oValue.oValue.dValue = Cm_To_Pt(dValue);
			} else if (-1 !== (nPos = sValue.indexOf("mm"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeUnits;

				let strValue = sValue.substring(0, nPos);
				let dValue = parseFloat(sValue);

				this.m_oValue.oValue.dValue = Mm_To_Pt(dValue);
			} else if (-1 !== (nPos = sValue.indexOf("em"))) {
			} else if (-1 !== (nPos = sValue.indexOf("ex"))) {
			} else if (-1 !== (nPos = sValue.indexOf("pt"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeUnits;

				let strValue = sValue.substring(0, nPos);
				let dValue = parseFloat(sValue);

				this.m_oValue.oValue.dValue = dValue;
			} else if (-1 !== (nPos = sValue.indexOf("pc"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeUnits;

				let strValue = sValue.substring(0, nPos);
				let dValue = parseFloat(sValue);

				this.m_oValue.oValue.dValue = dValue * 12;
			} else if (-1 !== (nPos = sValue.indexOf("%"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypePerc;

				let strValue = sValue.substring(0, nPos);
				this.m_oValue.oValue.dValue = strValue.length === 0 ? 0 : parseFloat(strValue);
			} else if (-1 !== (nPos = sValue.indexOf("px"))) {
				this.m_oValue.oValue.m_eType = ECssUnitsType.cssunitstypeUnits;

				let strValue = sValue.substring(0, nPos);
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

		function CVMLToDrawingMLConverter() {
		}
		CVMLToDrawingMLConverter.prototype.convertBackground = function(reader) {
			let oShape = new AscFormat.CShape();
			let oSpPr = new AscFormat.CSpPr();
			oSpPr.setGeometry(AscFormat.CreateGeometry("rect"));
			oShape.setSpPr(oSpPr);
			let oDrawingConverter = this;
			let oNode = new CT_XmlNode(function(reader, name) {
				if(name === "fill") {
					oSpPr.setFill(oDrawingConverter.convertFill(reader));
				}
				return true;
			});
			oNode.fromXml(reader);
			return oShape;
		};
		CVMLToDrawingMLConverter.prototype.convertObject = function(reader) {
			let oMapShapeTypes = {};
			let oGraphicObject = null;
			let oVMLDrawing = new CVMLDrawing();
			oVMLDrawing.fromXml(reader);


			let oShape = new AscFormat.CShape();
			let oSpPr = new AscFormat.CSpPr();
			oSpPr.setGeometry(AscFormat.CreateGeometry("rect"));
			oShape.setSpPr(oSpPr);
			let oDrawingConverter = this;
			let oNode = new CT_XmlNode(function(reader, name) {
				if(name === "fill") {
					oSpPr.setFill(oDrawingConverter.convertFill(reader));
				}
				return true;
			});
			oNode.fromXml(reader);
			return oShape;
		};
		CVMLToDrawingMLConverter.prototype.convertOleObject = function(reader) {
			let oVMLOleObj = new COLEObject();
			oVMLOleObj.fromXml(reader);
			let oEditorOLEObject = new AscFormat.COLEObject();
			oVMLOleObj.fillEditorOleObject(oEditorOLEObject, oVMLOleObj.m_oPic,  reader);
			return oEditorOLEObject;
		};
		CVMLToDrawingMLConverter.prototype.convertFill = function(reader) {
			return null;
		};
		CVMLToDrawingMLConverter.prototype.convertObjectRecursive = function(name, reader, paraDrawing) {
			switch (name) {
				case "background": {
					return this.convertBackground(reader);
				}
				case "pict":
				case "object": {
					return this.convertObject(reader);
				}
				case "oleObj": {
					return this.convertOleObject(reader);
				}
				case "drawing": {
					break;
				}
				case "Choice":
				case "Fallback":
				case "AlternateContent": {
					let oDrawing = null;
					let oDrawingConverter = this;
					let oNode = new CT_XmlNode(function(reader, name) {
						if(oDrawing) {
							return true;
						}
						oDrawing = oDrawingConverter.convertObjectRecursive(name, reader, paraDrawing);
						return true;
					});
					oNode.fromXml(reader);
					if(oDrawing) {
						return oDrawing;
					}
					break;
				}
			}
			return null;
		};
		CVMLToDrawingMLConverter.prototype.createParaDrawingMLFromVMLNode = function(reader, name, paragraph) {
			let oOOXMLDrawing = null;
			if(name === "pict" || name === "object") {
				let oLegacyDrawing = new CLegacyDrawing(name);
				oLegacyDrawing.fromXml(reader);
				oOOXMLDrawing = oLegacyDrawing.convertToDrawingML(reader);
			}
			else if(name === "oleObj") {
				let oOleObject = new AscFormat.COLEObject();
				oOleObject.fromXml(reader);
				oOOXMLDrawing = new AscFormat.COLEObject();
				oOleObject.fillEditorOleObject(oOOXMLDrawing, null, reader);
			}
			if(oOOXMLDrawing) {
				let oParaDrawing = new AscCommonWord.ParaDrawing(0, 0, oOOXMLDrawing, reader.context.DrawingDocument, paragraph.Parent, paragraph);
				oOOXMLDrawing.setParent(oParaDrawing);
				let oXfrm = oOOXMLDrawing.spPr && oOOXMLDrawing.spPr.xfrm;
				if(oXfrm) {
					oParaDrawing.setExtent(oXfrm.extX, oXfrm.extY);
					oXfrm.setOffX(0);
					oXfrm.setOffY(0);
				}
				CLegacyDrawing.prototype.GetDrawingMainProps(oParaDrawing, reader.context);
				return oParaDrawing;
			}
			return null;
		};


		function getRotateAngle(sRot, flipX, flipY) {
			let nCheckInvert = 0;

			if (flipX === true)
				nCheckInvert += 1;
			if (flipY === true)
				nCheckInvert += 1;

			let nRot = parseInt(sRot);
			if (sRot.indexOf('f') !== -1) {
				let dVal = nRot;
				dVal /= 65536;

				if (nCheckInvert === 1) {
					dVal = -dVal;
				}

				if (dVal > 360) {
					let nPart = (dVal / 360 + 0.5) >> 0 ;
					dVal = dVal - nPart * 360;
				}
				else if (dVal < 0)
				{
					let nPart = (dVal / 360 + 0.5) >> 0;
					nPart = 1 - nPart;
					dVal = dVal + nPart * 360;
				}

				nRot = (dVal * 60000 + 0.5) >> 0;
			}
			else {
				if (nCheckInvert === 1) {
					nRot = -nRot;
				}

				if (nRot > 360) {
					let nPart = (nRot / 360 + 0.5) >> 0.5;
					nRot = nRot - nPart * 360;
				}
				else if (nRot < 0) {
					let nPart = (nRot / 360 + 0.5) >> 0.5;
					nPart = 1 - nPart;
					nRot = nRot + nPart * 360;
				}

				nRot *= 60000;
			}

			nRot  *= AscFormat.cToRad;
			return nRot;
		}


		window['AscFormat'].CVMLDrawing = CVMLDrawing;
		window['AscFormat'].ECssPropertyType = ECssPropertyType;
		window['AscFormat'].ECssUnitsType = ECssUnitsType;
		window['AscFormat'].CPoint = CPoint;
		window['AscFormat'].CVMLToDrawingMLConverter = CVMLToDrawingMLConverter;
		window['AscFormat'].COLEObject = COLEObject;
		window['AscFormat'].CVMLClientData = CClientData;
		window['AscFormat'].CVMLSignatureLine = CSignatureLine;
		window['AscFormat'].EOLEDrawAspect = EOLEDrawAspect;
		window['AscFormat'].EOLEType = EOLEType;
		window['AscFormat'].EVmlClientDataObjectType = EVmlClientDataObjectType;
		window['AscFormat'].Pt_To_Px = Pt_To_Px;
		window['AscFormat'].Emu_To_Px = Emu_To_Px;
		window['AscFormat'].Mm_To_Px = Mm_To_Px;
		window['AscFormat'].Px_To_Mm = Px_To_Mm;
		window['AscFormat'].Emu_To_Mm = Emu_To_Mm;
		window['AscFormat'].Mm_To_Emu = Mm_To_Emu;
		window['AscFormat'].Emu_To_Twips = Emu_To_Twips;
		window['AscFormat'].Px_To_Emu = Px_To_Emu;

	})(window);

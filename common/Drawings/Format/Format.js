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

(
	/**
	 * @param {Window} window
	 * @param {undefined} undefined
	 */
	function (window, undefined) {


		var recalcSlideInterval = 30;

// Import
		var CreateAscColor = AscCommon.CreateAscColor;
		var g_oIdCounter = AscCommon.g_oIdCounter;
		var g_oTableId = AscCommon.g_oTableId;
		var isRealObject = AscCommon.isRealObject;
		
		var c_oAscColor = Asc.c_oAscColor;
		var c_oAscFill = Asc.c_oAscFill;
		var asc_CShapeFill = Asc.asc_CShapeFill;
		var c_oAscFillGradType = Asc.c_oAscFillGradType;
		var c_oAscFillBlipType = Asc.c_oAscFillBlipType;
		var c_oAscStrokeType = Asc.c_oAscStrokeType;
		var asc_CShapeProperty = Asc.asc_CShapeProperty;

		var g_nodeAttributeStart = AscCommon.g_nodeAttributeStart;
		var g_nodeAttributeEnd = AscCommon.g_nodeAttributeEnd;


		var CChangesDrawingsBool = AscDFH.CChangesDrawingsBool;
		var CChangesDrawingsLong = AscDFH.CChangesDrawingsLong;
		var CChangesDrawingsDouble = AscDFH.CChangesDrawingsDouble;
		var CChangesDrawingsString = AscDFH.CChangesDrawingsString;
		var CChangesDrawingsObjectNoId = AscDFH.CChangesDrawingsObjectNoId;
		var CChangesDrawingsObject = AscDFH.CChangesDrawingsObject;
		var CChangesDrawingsContentNoId = AscDFH.CChangesDrawingsContentNoId;
		var CChangesDrawingsContentLong = AscDFH.CChangesDrawingsContentLong;
		var CChangesDrawingsContentLongMap = AscDFH.CChangesDrawingsContentLongMap;
		var CChangesDrawingsContent = AscDFH.CChangesDrawingsContent;
		var CChangesDrawingsDouble2 = AscDFH.CChangesDrawingsDouble2;



		function CBaseNoIdObject() {
		}


		CBaseNoIdObject.prototype.classType = AscDFH.historyitem_type_Unknown;
		CBaseNoIdObject.prototype.notAllowedWithoutId = function () {
			return false;
		};
		CBaseNoIdObject.prototype.getObjectType = function () {
			return this.classType;
		};
		CBaseNoIdObject.prototype.Get_Id = function () {
			return this.Id;
		};
		CBaseNoIdObject.prototype.GetId = function () {
			return this.Id;
		};
		CBaseNoIdObject.prototype.Write_ToBinary2 = function (oWriter) {
			oWriter.WriteLong(this.getObjectType());
			oWriter.WriteString2(this.Get_Id());
		};
		CBaseNoIdObject.prototype.Read_FromBinary2 = function (oReader) {
			this.Id = oReader.GetString2();
		};
		CBaseNoIdObject.prototype.Refresh_RecalcData = function (oChange) {
		};
		CBaseNoIdObject.prototype.readAttr = function (reader) {
			while (reader.MoveToNextAttribute()) {
				this.readAttrXml(reader.GetNameNoNS(), reader);
			}
		};
		CBaseNoIdObject.prototype.readAttrXml = function (name, reader) {
			//Implement in children
		};
		CBaseNoIdObject.prototype.readChildXml = function (name, reader) {
			//Implement in children
		};
		CBaseNoIdObject.prototype.writeAttrXmlImpl = function (writer) {
			//Implement in children
		};
		CBaseNoIdObject.prototype.writeChildrenXml = function (writer) {
			//Implement in children
		};
		CBaseNoIdObject.prototype.fromXml = function (reader, bSkipFirstNode) {
			if (bSkipFirstNode) {
				if (!reader.ReadNextNode()) {
					return;
				}
			}
			this.readAttr(reader);
			let depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				let name = reader.GetNameNoNS();
				this.readChildXml(name, reader);
			}
		};
		CBaseNoIdObject.prototype.toXml = function (writer, name) {
			writer.WriteXmlNodeStart(name);
			this.writeAttrXml(writer);
			this.writeChildrenXml(writer);
			writer.WriteXmlNodeEnd(name);
		};
		CBaseNoIdObject.prototype.writeAttrXml = function (writer) {
			this.writeAttrXmlImpl(writer);
			writer.WriteXmlAttributesEnd();
		};


		function CBaseObject() {
			CBaseNoIdObject.call(this);
			this.Id = null;
			this.initId();
		}
		InitClass(CBaseObject, CBaseNoIdObject, AscDFH.historyitem_type_Unknown);
		CBaseObject.prototype.initId = function() {
			if ((AscCommon.g_oIdCounter.m_bLoad || AscCommon.History.CanAddChanges() || this.notAllowedWithoutId()) && !this.isGlobalSkipAddId()) {
				this.Id = AscCommon.g_oIdCounter.Get_NewId();
				AscCommon.g_oTableId.Add(this, this.Id);
			}
		};
		CBaseObject.prototype.notAllowedWithoutId = function () {
			return true;
		};
		CBaseObject.prototype.isGlobalSkipAddId = function () {
			const oApi = window.editor || Asc.editor;
			return !!(oApi && oApi.isSkipAddIdToBaseObject);
		};
		//Method for debug
		CBaseObject.prototype.compareTypes = function (oOther) {
			if (!oOther || !oOther.compareTypes) {
				debugger;
			}
			for (var sKey in oOther) {
				if ((oOther[sKey] === null || oOther[sKey] === undefined) &&
					(this[sKey] !== null && this[sKey] !== undefined)
					|| (this[sKey] === null || this[sKey] === undefined) &&
					(oOther[sKey] !== null && oOther[sKey] !== undefined)
					|| (typeof this[sKey]) !== (typeof oOther[sKey])) {
					debugger;
				}
				if (this[sKey] !== this.parent && this[sKey] !== this.group && typeof this[sKey] === "object" && this[sKey] && this[sKey].compareTypes) {
					this[sKey].compareTypes(oOther[sKey]);
				}
				if (Array.isArray(this[sKey])) {
					if (!Array.isArray(oOther[sKey])) {
						debugger;
					} else {
						var a1 = this[sKey];
						var a2 = oOther[sKey];
						if (a1.length !== a2.length) {
							debugger;
						} else {
							for (var i = 0; i < a1.length; ++i) {
								if (!a1[i] || !a2[i]) {
									debugger;
								}
								if (typeof a1[i] === "object" && a1[i] && a1[i].compareTypes) {
									a1[i].compareTypes(a2[i]);
								}
							}
						}
					}
				}
			}
		};

		function InitClassWithoutType(fClass, fBase) {
			fClass.prototype = Object.create(fBase.prototype);
			fClass.prototype.superclass = fBase;
			fClass.prototype.constructor = fClass;
		}

		function InitClass(fClass, fBase, nType) {
			InitClassWithoutType(fClass, fBase);
			fClass.prototype.classType = nType;
		}

		function CBaseFormatNoIdObject() {
			CBaseNoIdObject.call(this);
		}
		InitClass(CBaseFormatNoIdObject, CBaseNoIdObject, AscDFH.historyitem_type_Unknown);
		CBaseFormatNoIdObject.prototype.getImageFromBulletsMap = function(oImages) {};
		CBaseFormatNoIdObject.prototype.getDocContentsWithImageBullets = function (arrContents) {};
		CBaseFormatNoIdObject.prototype.createDuplicate = function (oIdMap) {
			var oCopy = new this.constructor();
			this.fillObject(oCopy, oIdMap);
			return oCopy;
		};
		CBaseFormatNoIdObject.prototype.fillObject = function (oCopy, oIdMap) {
		};
		CBaseFormatNoIdObject.prototype.fromPPTY = function (pReader) {
			var oStream = pReader.stream;
			var nStart = oStream.cur;
			var nEnd = nStart + oStream.GetULong() + 4;
			if (this.readAttribute) {
				this.readAttributes(pReader);
			}
			this.readChildren(nEnd, pReader);
			oStream.Seek2(nEnd);
		};
		CBaseFormatNoIdObject.prototype.readAttributes = function (pReader) {
			var oStream = pReader.stream;
			oStream.Skip2(1); // start attributes
			while (true) {
				var nType = oStream.GetUChar();
				if (nType == g_nodeAttributeEnd)
					break;
				if (this.readAttribute) {
					this.readAttribute(nType, pReader);
				}
			}
		};
		CBaseFormatNoIdObject.prototype.readAttribute = function (nType, pReader) {//todo return undefined by default(check pptx)
		};
		CBaseFormatNoIdObject.prototype.readChildren = function (nEnd, pReader) {
			var oStream = pReader.stream;
			while (oStream.cur < nEnd) {
				var nType = oStream.GetUChar();
				const res = this.readChild(nType, pReader);
				if (false === res) {
					oStream.SkipRecord();
				}
			}
		};
		CBaseFormatNoIdObject.prototype.readChild = function (nType, pReader) {
			pReader.stream.SkipRecord();
		};
		CBaseFormatNoIdObject.prototype.toPPTY = function (pWriter) {
			if (this.privateWriteAttributes) {
				this.writeAttributes(pWriter);
			}
			this.writeChildren(pWriter);
		};
		CBaseFormatNoIdObject.prototype.writeAttributes = function (pWriter) {
			pWriter.WriteUChar(g_nodeAttributeStart);
			if (this.privateWriteAttributes) {
				this.privateWriteAttributes(pWriter);
			}
			pWriter.WriteUChar(g_nodeAttributeEnd);
		};
		CBaseFormatNoIdObject.prototype.privateWriteAttributes = function (pWriter) {//todo return undefined by default(check pptx)
		};
		CBaseFormatNoIdObject.prototype.writeChildren = function (pWriter) {
		};
		CBaseFormatNoIdObject.prototype.writeRecord1 = function (pWriter, nType, oChild) {
			if (AscCommon.isRealObject(oChild)) {
				pWriter.WriteRecord1(nType, oChild, function (oChild) {
					oChild.toPPTY(pWriter);
				});
			} else {
				//TODO: throw an error
			}
		};
		CBaseFormatNoIdObject.prototype.writeRecord2 = function (pWriter, nType, oChild) {
			if (AscCommon.isRealObject(oChild)) {
				this.writeRecord1(pWriter, nType, oChild);
			}
		};
		CBaseFormatNoIdObject.prototype.getChildren = function () {
			return [];
		};
		CBaseFormatNoIdObject.prototype.traverse = function (fCallback, bReverseOrder) {
			if (fCallback(this)) {
				return true;
			}
			let aChildren = this.getChildren();
			if(bReverseOrder === undefined || bReverseOrder === true) {
				for (let nChild = aChildren.length - 1; nChild > -1; --nChild) {
					let oChild = aChildren[nChild];
					if (oChild && oChild.traverse) {
						if (oChild.traverse(fCallback, bReverseOrder)) {
							return true;
						}
					}
				}
			}
			else {
				for (let nChild = 0; nChild < aChildren.length; ++nChild) {
					let oChild = aChildren[nChild];
					if (oChild && oChild.traverse) {
						if (oChild.traverse(fCallback, bReverseOrder)) {
							return true;
						}
					}
				}
			}
			return false;
		};
		CBaseFormatNoIdObject.prototype.isEqual = function (oOther) {
			if (!oOther) {
				return false;
			}
			if (this.getObjectType() !== oOther.getObjectType()) {
				return false;
			}
			var aThisChildren = this.getChildren();
			var aOtherChildren = oOther.getChildren();
			if (aThisChildren.length !== aOtherChildren.length) {
				return false;
			}
			for (var nChild = 0; nChild < aThisChildren.length; ++nChild) {
				var oThisChild = aThisChildren[nChild];
				var oOtherChild = aOtherChildren[nChild];
				if (oThisChild !== this.checkEqualChild(oThisChild, oOtherChild)) {
					return false;
				}
			}
			return true;
		};
		CBaseFormatNoIdObject.prototype.checkEqualChild = function (oThisChild, oOtherChild) {
			if (AscCommon.isRealObject(oThisChild) && oThisChild.isEqual) {
				if (!oThisChild.isEqual(oOtherChild)) {
					return undefined;
				}
			} else {
				if (oThisChild !== oOtherChild) {
					return undefined;
				}
			}
			return oThisChild;
		};
		function CBaseFormatObject() {
			CBaseFormatNoIdObject.call(this);
			this.Id = null;
			this.initId();
			this.parent = null;
		}
		InitClass(CBaseFormatObject, CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);
		CBaseFormatObject.prototype.notAllowedWithoutId = CBaseObject.prototype.notAllowedWithoutId;
		CBaseFormatObject.prototype.isGlobalSkipAddId = CBaseObject.prototype.isGlobalSkipAddId;
		CBaseFormatObject.prototype.initId = CBaseObject.prototype.initId;
		CBaseFormatObject.prototype.compareTypes = CBaseObject.prototype.compareTypes;

		CBaseFormatObject.prototype.setParent = function (oParent) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_CommonChartFormat_SetParent, this.parent, oParent));
			this.parent = oParent;
		};
		CBaseFormatObject.prototype.setParentToChild = function (oChild) {
			if (oChild && oChild.setParent) {
				oChild.setParent(this);
			}
		};
		CBaseFormatObject.prototype.handleRemoveObject = function (sObjectId) {
			return false;
		};
		CBaseFormatObject.prototype.onRemoveChild = function (oChild) {
			if (this.parent) {
				this.parent.onRemoveChild(this);
			}
		};


		// Visio Extensions
		function CClrSchemeExtLst() {
			CBaseNoIdObject.call(this);

			/**
			 * @type {CVariationClrScheme[]}
			 */
			this.variationClrSchemeLst = [];
			/**
			 * from vt:bkgnd tag. Not visible background but something else
			 * @type {?CVarColor}
			 */
			this.background = null;
			/**
			 * attribute schemeEnum from vt:schemeID tag
			 * @type {?string}
			 */
			this.schemeEnum = null;
		}
		InitClass(CClrSchemeExtLst, CBaseNoIdObject, 0);

		/**
		 * props container class
		 * @constructor
		 */
		function CFontProps() {
			CBaseFormatNoIdObject.call(this);

			/**
			 * number read as string
			 * @type {?string}
			 */
			this.style = null;

			/**
			 * concrete props object
			 * vt:color with a:schemeClr or something else
			 * See CUniColor for color reference.
			 * @type {{color: (CSchemeColor | CSysColor
			 * | CPrstColor | CRGBColor | any)}}
			 */
			this.fontPropsObject = {
				color: null
			};
		}
		InitClass(CFontProps, CBaseFormatNoIdObject, 0);
		/**
		 * @param [bSaveFormatting = false] - made by example false is set in other cases by default
		 * @return {CFontProps}
		 */
		CFontProps.prototype.createDuplicate = function (bSaveFormatting) {
			if (bSaveFormatting === undefined) {
				bSaveFormatting = false;
			}

			var duplicate = new CFontProps();

			duplicate.style = this.style;

			if (null !== this.fontPropsObject.color) {
				if (bSaveFormatting === true) {
					duplicate.fontPropsObject.color = this.fontPropsObject.color.saveSourceFormatting();
				} else {
					duplicate.fontPropsObject.color = this.fontPropsObject.color.createDuplicate();
				}
			}

			return duplicate;
		};
		/**
		 * Read attributes from stream for CFontProps
		 *
		 * @param {number} attrType - The attribute type
		 * @param {CBinaryFileReader} pReader - The binary reader
		 * @return {boolean} true if the attribute was read, false otherwise
		 */
		CFontProps.prototype.readAttribute = function(attrType, pReader) {
			if (attrType === 0) {
				this.style = pReader.stream.GetULong().toString();
			} else {
				return false;
			}
			return true;
		};
		/**
		 * Read children from stream for CFontProps
		 *
		 * @param {number} elementType - The element type
		 * @param {CBinaryFileReader} pReader - The binary reader
		 * @return {boolean} true if the child was read, false otherwise
		 */
		CFontProps.prototype.readChild = function(elementType, pReader) {
			let handled = true;
			switch (elementType) {
				case 0:
					this.fontPropsObject.color = pReader.ReadUniColor();
					break;
				default:
					handled = false;
			}
			return handled;
		}
		/**
		 * Write attributes to stream for CFontProps
		 * 
		 * @param {CBinaryFileWriter} pWriter - The binary writer
		 */
		CFontProps.prototype.privateWriteAttributes = function(pWriter) {
			pWriter._WriteUInt2(0, parseInt(this.style));
		};

		/**
		 * Write children to stream for CFontProps
		 *
		 * @param {CBinaryFileWriter} pWriter - The binary writer
		 */
		CFontProps.prototype.writeChildren = function(pWriter) {
			pWriter.WriteRecord1(0, this.fontPropsObject.color, pWriter.WriteUniColor);
		};

		/**
		 * line props container class. child element of fmtConnectorSchemeLineStyles or fmtSchemeLineStyles.
		 * see 2.3.4.2.20	CT_LineStyle
		 * @constructor
		 */
		function CLineStyle() {
			CBaseFormatNoIdObject.call(this);

			/**
			 * 2.3.4.2.19	CT_LineEx.
			 * default is
			 * <vt:lineEx rndg="0" start="0" startSize="2" end="0" endSize="2" pattern="1"/>
			 * @type {Object}
			 */
			this.lineEx = {rndg: 0, start: 0, startSize: 2, end: 0, endSize: 2, pattern: 1};

			/**
			 * 2.3.4.2.25	CT_Sketch
			 * @type {Object}
			 */
			this.sketch = {};
		}
		InitClass(CLineStyle, CBaseFormatNoIdObject, 0);
		/**
		 * @param [bSaveFormatting = false] - made by example false is set in other cases by default
		 * @return {CLineStyle}
		 */
		CLineStyle.prototype.createDuplicate = function (bSaveFormatting) {
			if (bSaveFormatting === undefined) {
				bSaveFormatting = false;
			}

			var duplicate = new CLineStyle();

			if (null !== this.lineEx) {
				// duplicate.lineEx = this.lineEx.createDuplicate();
				duplicate.lineEx.rndg = this.lineEx.rndg;
				duplicate.lineEx.start = this.lineEx.start;
				duplicate.lineEx.startSize = this.lineEx.startSize;
				duplicate.lineEx.end = this.lineEx.end;
				duplicate.lineEx.endSize = this.lineEx.endSize;
				duplicate.lineEx.pattern = this.lineEx.pattern;
			}

			if (null !== this.sketch) {
				// 2.3.4.2.25	CT_Sketch
				// duplicate.sketch = this.sketch.createDuplicate();
				duplicate.sketch.lnAmp = this.sketch.lnAmp;
				duplicate.sketch.fillAmp = this.sketch.fillAmp;
				duplicate.sketch.lnWeight = this.sketch.lnWeight;
				duplicate.sketch.numPts = this.sketch.numPts;
			}

			return duplicate;
		};
		CLineStyle.prototype.readAttribute = undefined;
		/**
		 * Read children from stream for CLineStyle
		 *
		 * @param {number} elementType - The type of the element to read
		 * @param {CBinaryFileReader} pReader - The binary reader
		 * @return {boolean} True if the element was read, false otherwise
		 */
		CLineStyle.prototype.readChild = function(elementType, pReader) {
			const t = this;
			let handled = true;
			switch (elementType) {
				case 0:
					AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
						readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
						readAttributes: AscFormat.CBaseFormatNoIdObject.prototype.readAttributes,
						readAttribute: function(attrType, pReader) {
							if (attrType === 0) {
								t.lineEx.rndg = pReader.stream.GetULong();
							} else if (attrType === 1) {
								t.lineEx.start = pReader.stream.GetULong();
							} else if (attrType === 2) {
								t.lineEx.startSize = pReader.stream.GetULong();
							} else if (attrType === 3) {
								t.lineEx.end = pReader.stream.GetULong();
							} else if (attrType === 4) {
								t.lineEx.endSize = pReader.stream.GetULong();
							} else if (attrType === 5) {
								t.lineEx.pattern = pReader.stream.GetULong();
							} else {
								return false;
							}
							return true;
						}
					}, pReader);
					break;
				case 1:
					AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
						readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
						readAttributes: AscFormat.CBaseFormatNoIdObject.prototype.readAttributes,
						readAttribute: function(attrType, pReader) {
							if (attrType === 0) {
								t.sketch.lnAmp = pReader.stream.GetULong();
							} else if (attrType === 1) {
								t.lineEx.fillAmp = pReader.stream.GetULong();
							} else if (attrType === 2) {
								t.lineEx.lnWeight = pReader.stream.GetULong();
							} else if (attrType === 3) {
								t.lineEx.numPts = pReader.stream.GetULong();
							} else {
								return false;
							}
							return true;
						}
					}, pReader);
					break;
				default:
					handled = false;
					break;
			}
			return handled;
		}
		CLineStyle.prototype.privateWriteAttributes = undefined;
		/**
		 * Write children to stream for CLineStyle
		 *
		 * @param {CBinaryFileWriter} pWriter - The binary writer
		 */
		CLineStyle.prototype.writeChildren = function(pWriter) {
			pWriter.StartRecord(0);
			pWriter.WriteUChar(AscCommon.g_nodeAttributeStart);
			pWriter._WriteUInt2(0, this.lineEx.rndg);
			pWriter._WriteUInt2(1, this.lineEx.start);
			pWriter._WriteUInt2(2, this.lineEx.startSize);
			pWriter._WriteUInt2(3, this.lineEx.end);
			pWriter._WriteUInt2(4, this.lineEx.endSize);
			pWriter._WriteUInt2(5, this.lineEx.pattern);
			pWriter.WriteUChar(AscCommon.g_nodeAttributeEnd);
			pWriter.EndRecord();
			if (Object.keys(this.sketch).length > 0) {
				pWriter.StartRecord(1);
				pWriter.WriteUChar(AscCommon.g_nodeAttributeStart);
				pWriter._WriteUInt2(0, this.sketch.lnAmp);
				pWriter._WriteUInt2(1, this.sketch.fillAmp);
				pWriter._WriteUInt2(2, this.sketch.lnWeight);
				pWriter._WriteUInt2(3, this.sketch.numPts);
				pWriter.WriteUChar(AscCommon.g_nodeAttributeEnd);
				pWriter.EndRecord();
			}
		};

		function CVariationClrScheme() {
			CBaseFormatNoIdObject.call(this);
			/**
			 *
			 * @type {*}
			 */
			this.monotone = null;
			/**
			 * read from varColor1, varColor2, varColor3, ...
			 * @type {CVarColor[]}
			 */
			this.varColor = [];
		}
		InitClass(CVariationClrScheme, CBaseFormatNoIdObject, 0);
		CVariationClrScheme.prototype.readAttribute = undefined;
		/**
		 * Read children from stream for CVariationClrScheme
		 *
		 * @param {number} elementType - The type of the element to read
		 * @param {CBinaryFileReader} pReader - The binary reader
		 * @return {boolean} True if the element was read, false otherwise
		 */
		CVariationClrScheme.prototype.readChild = function(elementType, pReader) {
			let handled = true;
			if (0 <= elementType && elementType < 7) {
				let varColor = new CVarColor();
				varColor.unicolor = pReader.ReadUniColor();
				this.varColor[elementType] = varColor;
			} else {
				handled = false;
			}
			return handled;
		}
		CVariationClrScheme.prototype.privateWriteAttributes = undefined;
		/**
		 * Write children to stream for CLineStyle
		 *
		 * @param {CBinaryFileWriter} pWriter - The binary writer
		 */
		CVariationClrScheme.prototype.writeChildren = function(pWriter) {
			for (let i = 0; i < this.varColor.length; i++) {
				if (!this.varColor[i] || !this.varColor[i].unicolor) {
					continue;
				}
				pWriter.WriteRecord1(i, this.varColor[i].unicolor, pWriter.WriteUniColor);
			}
		};

		function CVarColor() {
			CBaseNoIdObject.call(this);
			/**
			 * @type {CUniColor}
			 */
			this.unicolor = null;
		}
		InitClass(CVarColor, CBaseNoIdObject, 0);

		/**
		 * General theme extensions. Tag inside themeElements.
		 * @constructor
		 */
		function CThemeExt() {
			CBaseNoIdObject.call(this);

			/**
			 * @type {?FmtScheme}
			 */
			this.fmtConnectorScheme = null;

			/**
			 * The only value from themeScheme tag
			 * @type {?string}
			 */
			this.themeSchemeSchemeEnum = null;

			/**
			 * The only value from fmtSchemeEx tag
			 * @type {?string}
			 */
			this.fmtSchemeExSchemeEnum = null;

			/**
			 * The only value from fmtConnectorSchemeEx tag
			 * @type {?string}
			 */
			this.fmtConnectorSchemeExSchemeEnum = null;

			/**
			 * vt:fillStyles
			 * @type {{pattern: number}[]}
			 */
			this.fillStyles = [];

			/**
			 * @type {{fmtConnectorSchemeLineStyles: CLineStyle[], fmtSchemeLineStyles: CLineStyle[] }}
			 */
			this.lineStyles = {
				fmtConnectorSchemeLineStyles: [],
				fmtSchemeLineStyles: []
			};

			/**
			 * @type {{connectorFontStyles: CFontProps[], fontStyles: CFontProps[] }}
			 */
			this.fontStylesGroup = {
				connectorFontStyles: [],
				fontStyles: []
			};
			this.variationStyleSchemeLst = [];
		}
		InitClass(CThemeExt, CBaseNoIdObject, 0);

		function CVariationStyleScheme() {
			CBaseFormatNoIdObject.call(this);

			this.embellishment = null;
			/**
			 * @type {CVarStyle[]}
			 */
			this.varStyle = [];
		}
		InitClass(CVariationStyleScheme, CBaseFormatNoIdObject, 0);
		/**
		 * Read attributes from stream for DocumentSettings_Type
		 * 
		 * @param  {number} attrType - The type of attribute
		 * @param {BinaryVSDYLoader} pReader - The binary reader
		 * @returns {boolean} - True if attribute was handled, false otherwise
		 */
		CVariationStyleScheme.prototype.readAttribute = function(attrType, pReader) {
			if (attrType === 0) {
				this.embellishment = pReader.stream.GetULong();
			} else {
				return false;
			}
			return true;
		}
		/**
		 * Read child elements from stream for CVariationStyleScheme
		 *
		 * @param {number} elementType - The type of child element
		 * @param {BinaryVSDYLoader} pReader - The binary reader
		 * @returns {boolean} - True if element was handled, false otherwise
		 */
		CVariationStyleScheme.prototype.readChild = function(elementType, pReader) {
			switch (elementType) {
				case 0: {
					const varStyle = new CVarStyle();
					varStyle.fromPPTY(pReader);
					this.varStyle.push(varStyle);
					break;
				}
				default:
					return false;
			}

			return true;
		};
		/**
		 * Write attributes to stream for CVariationStyleScheme
		 * 
		 * @param {CBinaryFileWriter} pWriter - The binary writer
		 */
		CVariationStyleScheme.prototype.privateWriteAttributes = function(pWriter) {
			pWriter._WriteUInt2(0, this.embellishment);
		};

		/**
		 * Write children to stream for CVariationStyleScheme
		 *
		 * @param {CBinaryFileWriter} pWriter - The binary writer
		 */
		CVariationStyleScheme.prototype.writeChildren = function(pWriter) {
			for (let i = 0; i < this.varStyle.length; i++) {
				pWriter.WriteRecordPPTY(0, this.varStyle[i]);
			}
		};

		function CVarStyle() {
			CBaseFormatNoIdObject.call(this);

			this.fillIdx = null;
			this.lineIdx = null;
			this.effectIdx = null;
			this.fontIdx = null;
		}
		InitClass(CVarStyle, CBaseFormatNoIdObject, 0);
		/**
		 * Read attributes from stream for DocumentSettings_Type
		 * 
		 * @param  {number} attrType - The type of attribute
		 * @param {BinaryVSDYLoader} pReader - The binary reader
		 * @returns {boolean} - True if attribute was handled, false otherwise
		 */
		CVarStyle.prototype.readAttribute = function(attrType, pReader) {
			if (attrType === 0) {
				this.fillIdx = pReader.stream.GetULong();
			} else if (attrType === 1) {
				this.lineIdx = pReader.stream.GetULong();
			} else if (attrType === 2) {
				this.effectIdx = pReader.stream.GetULong();
			} else if (attrType === 3) {
				this.fontIdx = pReader.stream.GetULong();
			} else {
				return false;
			}
			return true;
		}
		/**
		 * Write attributes to stream for CVarStyle
		 * 
		 * @param {CBinaryFileWriter} pWriter - The binary writer
		 */
		CVarStyle.prototype.privateWriteAttributes = function(pWriter) {
			pWriter._WriteUInt2(0, this.fillIdx);
			pWriter._WriteUInt2(1, this.lineIdx);
			pWriter._WriteUInt2(2, this.effectIdx);
			pWriter._WriteUInt2(3, this.fontIdx);
		};				
		// Theme visio extensions end


		function CT_Hyperlink() {
			CBaseNoIdObject.call(this);
			this.snd = null;
			this.id = null;
			this.invalidUrl = null;
			this.action = null;
			this.tgtFrame = null;
			this.tooltip = null;
			this.history = null;
			this.highlightClick = null;
			this.endSnd = null;
		}

		InitClass(CT_Hyperlink, CBaseNoIdObject, 0);
		CT_Hyperlink.prototype.Write_ToBinary = function (w) {
			var nStartPos = w.GetCurPosition();
			var nFlags = 0;
			w.WriteLong(0);

			if (null !== this.snd) {
				nFlags |= 1;
				w.WriteString2(this.snd);
			}
			if (null !== this.id) {
				nFlags |= 2;
				w.WriteString2(this.id);
			}
			if (null !== this.invalidUrl) {
				nFlags |= 4;
				w.WriteString2(this.invalidUrl);
			}
			if (null !== this.action) {
				nFlags |= 8;
				w.WriteString2(this.action);
			}
			if (null !== this.tgtFrame) {
				nFlags |= 16;
				w.WriteString2(this.tgtFrame);
			}
			if (null !== this.tooltip) {
				nFlags |= 32;
				w.WriteString2(this.tooltip);
			}
			if (null !== this.history) {
				nFlags |= 64;
				w.WriteBool(this.history);
			}
			if (null !== this.highlightClick) {
				nFlags |= 128;
				w.WriteBool(this.highlightClick);
			}
			if (null !== this.endSnd) {
				nFlags |= 256;
				w.WriteBool(this.endSnd);
			}
			var nEndPos = w.GetCurPosition();
			w.Seek(nStartPos);
			w.WriteLong(nFlags);
			w.Seek(nEndPos);
		};
		CT_Hyperlink.prototype.Read_FromBinary = function (r) {
			var nFlags = r.GetLong();
			if (nFlags & 1) {
				this.snd = r.GetString2();
			}
			if (nFlags & 2) {
				this.id = r.GetString2();
			}
			if (nFlags & 4) {
				this.invalidUrl = r.GetString2();
			}
			if (nFlags & 8) {
				this.action = r.GetString2();
			}
			if (nFlags & 16) {
				this.tgtFrame = r.GetString2();
			}
			if (nFlags & 32) {
				this.tooltip = r.GetString2();
			}
			if (nFlags & 64) {
				this.history = r.GetBool();
			}
			if (nFlags & 128) {
				this.highlightClick = r.GetBool();
			}
			if (nFlags & 256) {
				this.endSnd = r.GetBool();
			}
		};
		CT_Hyperlink.prototype.createDuplicate = function () {
			var ret = new CT_Hyperlink();
			ret.snd = this.snd;
			ret.id = this.id;
			ret.invalidUrl = this.invalidUrl;
			ret.action = this.action;
			ret.tgtFrame = this.tgtFrame;
			ret.tooltip = this.tooltip;
			ret.history = this.history;
			ret.highlightClick = this.highlightClick;
			ret.endSnd = this.endSnd;
			return ret;
		};


		var drawingsChangesMap = window['AscDFH'].drawingsChangesMap;
		var drawingConstructorsMap = window['AscDFH'].drawingsConstructorsMap;
		var drawingContentChanges = window['AscDFH'].drawingContentChanges;


		drawingsChangesMap[AscDFH.historyitem_DefaultShapeDefinition_SetSpPr] = function (oClass, value) {
			oClass.spPr = value;
		};
		drawingsChangesMap[AscDFH.historyitem_DefaultShapeDefinition_SetBodyPr] = function (oClass, value) {
			oClass.bodyPr = value;
		};
		drawingsChangesMap[AscDFH.historyitem_DefaultShapeDefinition_SetLstStyle] = function (oClass, value) {
			oClass.lstStyle = value;
		};
		drawingsChangesMap[AscDFH.historyitem_DefaultShapeDefinition_SetStyle] = function (oClass, value) {
			oClass.style = value;
		};
		drawingsChangesMap[AscDFH.historyitem_CNvPr_SetId] = function (oClass, value) {
			oClass.id = value;
		};
		drawingsChangesMap[AscDFH.historyitem_CNvPr_SetName] = function (oClass, value) {
			oClass.name = value;
		};
		drawingsChangesMap[AscDFH.historyitem_CNvPr_SetIsHidden] = function (oClass, value) {
			oClass.isHidden = value;
		};
		drawingsChangesMap[AscDFH.historyitem_CNvPr_SetDescr] = function (oClass, value) {
			oClass.descr = value;
		};
		drawingsChangesMap[AscDFH.historyitem_CNvPr_SetTitle] = function (oClass, value) {
			oClass.title = value;
		};
		drawingsChangesMap[AscDFH.historyitem_CNvPr_SetHlinkClick] = function (oClass, value) {
			oClass.hlinkClick = value;
		};
		drawingsChangesMap[AscDFH.historyitem_CNvPr_SetHlinkHover] = function (oClass, value) {
			oClass.hlinkHover = value;
		};
		drawingsChangesMap[AscDFH.historyitem_NvPr_SetIsPhoto] = function (oClass, value) {
			oClass.isPhoto = value;
		};
		drawingsChangesMap[AscDFH.historyitem_NvPr_SetUserDrawn] = function (oClass, value) {
			oClass.userDrawn = value;
		};
		drawingsChangesMap[AscDFH.historyitem_NvPr_SetPh] = function (oClass, value) {
			oClass.ph = value;
		};
		drawingsChangesMap[AscDFH.historyitem_Ph_SetHasCustomPrompt] = function (oClass, value) {
			oClass.hasCustomPrompt = value;
		};
		drawingsChangesMap[AscDFH.historyitem_Ph_SetIdx] = function (oClass, value) {
			oClass.idx = value;
		};
		drawingsChangesMap[AscDFH.historyitem_Ph_SetOrient] = function (oClass, value) {
			oClass.orient = value;
		};
		drawingsChangesMap[AscDFH.historyitem_Ph_SetSz] = function (oClass, value) {
			oClass.sz = value;
		};
		drawingsChangesMap[AscDFH.historyitem_Ph_SetType] = function (oClass, value) {
			oClass.type = value;
		};
		drawingsChangesMap[AscDFH.historyitem_UniNvPr_SetCNvPr] = function (oClass, value) {
			oClass.cNvPr = value;
		};
		drawingsChangesMap[AscDFH.historyitem_UniNvPr_SetUniPr] = function (oClass, value) {
			oClass.uniPr = value;
		};
		drawingsChangesMap[AscDFH.historyitem_UniNvPr_SetNvPr] = function (oClass, value) {
			oClass.nvPr = value;
		};
		drawingsChangesMap[AscDFH.historyitem_ShapeStyle_SetLnRef] = function (oClass, value) {
			oClass.lnRef = value;
		};
		drawingsChangesMap[AscDFH.historyitem_ShapeStyle_SetFillRef] = function (oClass, value) {
			oClass.fillRef = value;
		};
		drawingsChangesMap[AscDFH.historyitem_ShapeStyle_SetFontRef] = function (oClass, value) {
			oClass.fontRef = value;
		};
		drawingsChangesMap[AscDFH.historyitem_ShapeStyle_SetEffectRef] = function (oClass, value) {
			oClass.effectRef = value;
		};
		drawingsChangesMap[AscDFH.historyitem_Xfrm_SetParent] = function (oClass, value) {
			oClass.parent = value;
		};
		drawingsChangesMap[AscDFH.historyitem_Xfrm_SetOffX] = function (oClass, value) {
			oClass.offX = value;
			oClass.handleUpdatePosition();
		};
		drawingsChangesMap[AscDFH.historyitem_Xfrm_SetOffY] = function (oClass, value) {
			oClass.offY = value;
			oClass.handleUpdatePosition();
		};
		drawingsChangesMap[AscDFH.historyitem_Xfrm_SetExtX] = function (oClass, value) {
			oClass.extX = value;
			oClass.handleUpdateExtents();
		};
		drawingsChangesMap[AscDFH.historyitem_Xfrm_SetExtY] = function (oClass, value) {
			oClass.extY = value;
			oClass.handleUpdateExtents();
		};
		drawingsChangesMap[AscDFH.historyitem_Xfrm_SetChOffX] = function (oClass, value) {
			oClass.chOffX = value;
			oClass.handleUpdateChildOffset();
		};
		drawingsChangesMap[AscDFH.historyitem_Xfrm_SetChOffY] = function (oClass, value) {
			oClass.chOffY = value;
			oClass.handleUpdateChildOffset();
		};
		drawingsChangesMap[AscDFH.historyitem_Xfrm_SetChExtX] = function (oClass, value) {
			oClass.chExtX = value;
			oClass.handleUpdateChildExtents();
		};
		drawingsChangesMap[AscDFH.historyitem_Xfrm_SetChExtY] = function (oClass, value) {
			oClass.chExtY = value;
			oClass.handleUpdateChildExtents();
		};
		drawingsChangesMap[AscDFH.historyitem_Xfrm_SetFlipH] = function (oClass, value) {
			oClass.flipH = value;
			oClass.handleUpdateFlip();
		};
		drawingsChangesMap[AscDFH.historyitem_Xfrm_SetFlipV] = function (oClass, value) {
			oClass.flipV = value;
			oClass.handleUpdateFlip();
		};
		drawingsChangesMap[AscDFH.historyitem_Xfrm_SetRot] = function (oClass, value) {
			oClass.rot = value;
			oClass.handleUpdateRot();
		};
		drawingsChangesMap[AscDFH.historyitem_SpPr_SetParent] = function (oClass, value) {
			oClass.parent = value;
		};
		drawingsChangesMap[AscDFH.historyitem_SpPr_SetBwMode] = function (oClass, value) {
			oClass.bwMode = value;
		};
		drawingsChangesMap[AscDFH.historyitem_SpPr_SetXfrm] = function (oClass, value) {
			oClass.xfrm = value;
		};
		drawingsChangesMap[AscDFH.historyitem_SpPr_SetGeometry] = function (oClass, value) {
			oClass.geometry = value;
			oClass.handleUpdateGeometry();
		};
		drawingsChangesMap[AscDFH.historyitem_SpPr_SetFill] = function (oClass, value, FromLoad) {
			oClass.Fill = value;
			oClass.handleUpdateFill();
			if (FromLoad) {
				if (typeof AscCommon.CollaborativeEditing !== "undefined") {
					if (oClass.Fill && oClass.Fill.fill && oClass.Fill.fill.type === c_oAscFill.FILL_TYPE_BLIP && typeof oClass.Fill.fill.RasterImageId === "string" && oClass.Fill.fill.RasterImageId.length > 0) {
						AscCommon.CollaborativeEditing.Add_NewImage(oClass.Fill.fill.RasterImageId);
					}
				}
			}
		};
		drawingsChangesMap[AscDFH.historyitem_SpPr_SetLn] = function (oClass, value) {
			oClass.ln = value;
			oClass.handleUpdateLn();
		};
		drawingsChangesMap[AscDFH.historyitem_SpPr_SetEffectPr] = function (oClass, value) {
			oClass.effectProps = value;
			oClass.handleUpdateGeometry();
		};
		drawingsChangesMap[AscDFH.historyitem_ExtraClrScheme_SetClrScheme] = function (oClass, value) {
			oClass.clrScheme = value;
		};
		drawingsChangesMap[AscDFH.historyitem_ExtraClrScheme_SetClrMap] = function (oClass, value) {
			oClass.clrMap = value;
		};
		drawingsChangesMap[AscDFH.historyitem_ThemeSetColorScheme] = function (oClass, value) {
			oClass.themeElements.clrScheme = value;
			var oWordGraphicObjects = oClass.GetWordDrawingObjects();
			if (oWordGraphicObjects) {
				oWordGraphicObjects.drawingDocument.CheckGuiControlColors();
				oWordGraphicObjects.document.Api.chartPreviewManager.clearPreviews();
				oWordGraphicObjects.document.Api.textArtPreviewManager.clear();
			}
		};
		drawingsChangesMap[AscDFH.historyitem_ThemeSetFontScheme] = function (oClass, value) {
			oClass.themeElements.fontScheme = value;
		};
		drawingsChangesMap[AscDFH.historyitem_ThemeSetFmtScheme] = function (oClass, value, bFromLoad) {
			oClass.themeElements.fmtScheme = value;
			if(bFromLoad) {
				if(typeof AscCommon.CollaborativeEditing !== "undefined") {
					if(value) {
						let aImages = [];
						value.getAllRasterImages(aImages);
						for(let nImage = 0; nImage < aImages.length; ++nImage) {
							AscCommon.CollaborativeEditing.Add_NewImage(aImages[nImage]);
						}
					}
				}
			}
		};
		drawingsChangesMap[AscDFH.historyitem_ThemeSetName] = function (oClass, value) {
			oClass.name = value;
		};
		drawingsChangesMap[AscDFH.historyitem_ThemeSetIsThemeOverride] = function (oClass, value) {
			oClass.isThemeOverride = value;
		};
		drawingsChangesMap[AscDFH.historyitem_ThemeSetSpDef] = function (oClass, value) {
			oClass.spDef = value;
		};
		drawingsChangesMap[AscDFH.historyitem_ThemeSetLnDef] = function (oClass, value) {
			oClass.lnDef = value;
		};
		drawingsChangesMap[AscDFH.historyitem_ThemeSetTxDef] = function (oClass, value) {
			oClass.txDef = value;
		};
		drawingsChangesMap[AscDFH.historyitem_HF_SetDt] = function (oClass, value) {
			oClass.dt = value;
		};
		drawingsChangesMap[AscDFH.historyitem_HF_SetFtr] = function (oClass, value) {
			oClass.ftr = value;
		};
		drawingsChangesMap[AscDFH.historyitem_HF_SetHdr] = function (oClass, value) {
			oClass.hdr = value;
		};
		drawingsChangesMap[AscDFH.historyitem_HF_SetSldNum] = function (oClass, value) {
			oClass.sldNum = value;
		};
		drawingsChangesMap[AscDFH.historyitem_UniNvPr_SetUniSpPr] = function (oClass, value) {
			oClass.nvUniSpPr = value;
		};
		drawingsChangesMap[AscDFH.historyitem_NvPr_SetUniMedia] = function (oClass, value) {
			oClass.unimedia = value;
		};

		drawingContentChanges[AscDFH.historyitem_ClrMap_SetClr] = function (oClass) {
			return oClass.color_map
		};
		drawingContentChanges[AscDFH.historyitem_ThemeAddExtraClrScheme] = function (oClass) {
			return oClass.extraClrSchemeLst;
		};
		drawingContentChanges[AscDFH.historyitem_ThemeRemoveExtraClrScheme] = function (oClass) {
			return oClass.extraClrSchemeLst;
		};
		drawingContentChanges[AscDFH.historyitem_CustomPropertiesAddProperty] = function (oClass) {
			return oClass.properties;
		};


		drawingConstructorsMap[AscDFH.historyitem_ClrMap_SetClr] = CUniColor;
		drawingConstructorsMap[AscDFH.historyitem_DefaultShapeDefinition_SetBodyPr] = CBodyPr;
		drawingConstructorsMap[AscDFH.historyitem_DefaultShapeDefinition_SetLstStyle] = TextListStyle;
		drawingConstructorsMap[AscDFH.historyitem_ShapeStyle_SetLnRef] =
			drawingConstructorsMap[AscDFH.historyitem_ShapeStyle_SetFillRef] =
				drawingConstructorsMap[AscDFH.historyitem_ShapeStyle_SetEffectRef] = StyleRef;
		drawingConstructorsMap[AscDFH.historyitem_ShapeStyle_SetFontRef] = FontRef;
		drawingConstructorsMap[AscDFH.historyitem_SpPr_SetFill] = CUniFill;
		drawingConstructorsMap[AscDFH.historyitem_SpPr_SetLn] = CLn;
		drawingConstructorsMap[AscDFH.historyitem_SpPr_SetEffectPr] = CEffectProperties;
		drawingConstructorsMap[AscDFH.historyitem_ThemeSetColorScheme] = ClrScheme;
		drawingConstructorsMap[AscDFH.historyitem_ThemeSetFontScheme] = FontScheme;
		drawingConstructorsMap[AscDFH.historyitem_ThemeSetFmtScheme] = FmtScheme;
		drawingConstructorsMap[AscDFH.historyitem_UniNvPr_SetUniSpPr] = CNvUniSpPr;

		drawingConstructorsMap[AscDFH.historyitem_CNvPr_SetHlinkClick] = CT_Hyperlink;
		drawingConstructorsMap[AscDFH.historyitem_CNvPr_SetHlinkHover] = CT_Hyperlink;
		drawingConstructorsMap[AscDFH.historyitem_CustomPropertiesAddProperty] = CCustomProperty;

		AscDFH.changesFactory[AscDFH.historyitem_DefaultShapeDefinition_SetSpPr] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_DefaultShapeDefinition_SetBodyPr] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_DefaultShapeDefinition_SetLstStyle] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_DefaultShapeDefinition_SetStyle] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_CNvPr_SetId] = CChangesDrawingsLong;
		AscDFH.changesFactory[AscDFH.historyitem_CNvPr_SetName] = CChangesDrawingsString;
		AscDFH.changesFactory[AscDFH.historyitem_CNvPr_SetIsHidden] = CChangesDrawingsBool;
		AscDFH.changesFactory[AscDFH.historyitem_CNvPr_SetDescr] = CChangesDrawingsString;
		AscDFH.changesFactory[AscDFH.historyitem_CNvPr_SetTitle] = CChangesDrawingsString;
		AscDFH.changesFactory[AscDFH.historyitem_CNvPr_SetHlinkClick] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_CNvPr_SetHlinkHover] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_NvPr_SetIsPhoto] = CChangesDrawingsBool;
		AscDFH.changesFactory[AscDFH.historyitem_NvPr_SetUserDrawn] = CChangesDrawingsBool;
		AscDFH.changesFactory[AscDFH.historyitem_NvPr_SetPh] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_NvPr_AddExt] = CChangesDrawingsContentNoId;
		AscDFH.changesFactory[AscDFH.historyitem_NvPr_SetUniMedia] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_Ph_SetHasCustomPrompt] = CChangesDrawingsBool;
		AscDFH.changesFactory[AscDFH.historyitem_Ph_SetIdx] = CChangesDrawingsString;
		AscDFH.changesFactory[AscDFH.historyitem_Ph_SetOrient] = CChangesDrawingsLong;
		AscDFH.changesFactory[AscDFH.historyitem_Ph_SetSz] = CChangesDrawingsLong;
		AscDFH.changesFactory[AscDFH.historyitem_Ph_SetType] = CChangesDrawingsLong;
		AscDFH.changesFactory[AscDFH.historyitem_UniNvPr_SetCNvPr] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_UniNvPr_SetUniPr] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_UniNvPr_SetNvPr] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_UniNvPr_SetUniSpPr] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeStyle_SetLnRef] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeStyle_SetFillRef] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeStyle_SetFontRef] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeStyle_SetEffectRef] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_Xfrm_SetParent] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_Xfrm_SetOffX] = CChangesDrawingsDouble;
		AscDFH.changesFactory[AscDFH.historyitem_Xfrm_SetOffY] = CChangesDrawingsDouble;
		AscDFH.changesFactory[AscDFH.historyitem_Xfrm_SetExtX] = CChangesDrawingsDouble;
		AscDFH.changesFactory[AscDFH.historyitem_Xfrm_SetExtY] = CChangesDrawingsDouble;
		AscDFH.changesFactory[AscDFH.historyitem_Xfrm_SetChOffX] = CChangesDrawingsDouble;
		AscDFH.changesFactory[AscDFH.historyitem_Xfrm_SetChOffY] = CChangesDrawingsDouble;
		AscDFH.changesFactory[AscDFH.historyitem_Xfrm_SetChExtX] = CChangesDrawingsDouble;
		AscDFH.changesFactory[AscDFH.historyitem_Xfrm_SetChExtY] = CChangesDrawingsDouble;
		AscDFH.changesFactory[AscDFH.historyitem_Xfrm_SetFlipH] = CChangesDrawingsBool;
		AscDFH.changesFactory[AscDFH.historyitem_Xfrm_SetFlipV] = CChangesDrawingsBool;
		AscDFH.changesFactory[AscDFH.historyitem_Xfrm_SetRot] = CChangesDrawingsDouble;
		AscDFH.changesFactory[AscDFH.historyitem_SpPr_SetParent] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_SpPr_SetBwMode] = CChangesDrawingsLong;
		AscDFH.changesFactory[AscDFH.historyitem_SpPr_SetXfrm] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_SpPr_SetGeometry] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_SpPr_SetFill] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_SpPr_SetLn] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_SpPr_SetEffectPr] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_ClrMap_SetClr] = CChangesDrawingsContentLongMap;
		AscDFH.changesFactory[AscDFH.historyitem_ExtraClrScheme_SetClrScheme] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_ExtraClrScheme_SetClrMap] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ThemeSetColorScheme] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_ThemeSetFontScheme] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_ThemeSetFmtScheme] = CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_ThemeSetName] = CChangesDrawingsString;
		AscDFH.changesFactory[AscDFH.historyitem_ThemeSetIsThemeOverride] = CChangesDrawingsBool;
		AscDFH.changesFactory[AscDFH.historyitem_ThemeSetSpDef] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ThemeSetLnDef] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ThemeSetTxDef] = CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ThemeAddExtraClrScheme] = CChangesDrawingsContent;
		AscDFH.changesFactory[AscDFH.historyitem_ThemeRemoveExtraClrScheme] = CChangesDrawingsContent;
		AscDFH.changesFactory[AscDFH.historyitem_HF_SetDt] = CChangesDrawingsBool;
		AscDFH.changesFactory[AscDFH.historyitem_HF_SetFtr] = CChangesDrawingsBool;
		AscDFH.changesFactory[AscDFH.historyitem_HF_SetHdr] = CChangesDrawingsBool;
		AscDFH.changesFactory[AscDFH.historyitem_HF_SetSldNum] = CChangesDrawingsBool;
		AscDFH.changesFactory[AscDFH.historyitem_CustomPropertiesAddProperty] = CChangesDrawingsContentNoId;




// COLOR -----------------------
		/*
 var map_color_scheme = {};
 map_color_scheme["accent1"] = 0;
 map_color_scheme["accent2"] = 1;
 map_color_scheme["accent3"] = 2;
 map_color_scheme["accent4"] = 3;
 map_color_scheme["accent5"] = 4;
 map_color_scheme["accent6"] = 5;
 map_color_scheme["bg1"]     = 6;
 map_color_scheme["bg2"]     = 7;
 map_color_scheme["dk1"]     = 8;
 map_color_scheme["dk2"]     = 9;
 map_color_scheme["folHlink"] = 10;
 map_color_scheme["hlink"]   = 11;
 map_color_scheme["lt1"]     = 12;
 map_color_scheme["lt2"]     = 13;
 map_color_scheme["phClr"]   = 14;
 map_color_scheme["tx1"]     = 15;
 map_color_scheme["tx2"]     = 16;
 */

//Типы изменений в классе CTheme

		function CreateFontRef(idx, color) {
			var ret = new FontRef();
			ret.idx = idx;
			ret.Color = color;
			return ret;
		}

		function CreateStyleRef(idx, color) {
			var ret = new StyleRef();
			ret.idx = idx;
			ret.Color = color;
			return ret;
		}

		function CreatePresetColor(id) {
			var ret = new CPrstColor();
			ret.id = id;
			return ret;
		}

		function sRGB_to_scRGB(value) {
			if (value < 0)
				return 0;
			if (value <= 0.04045)
				return value / 12.92;
			if (value <= 1)
				return Math.pow(((value + 0.055) / 1.055), 2.4);
			return 1;
		}

		function scRGB_to_sRGB(value) {
			if (value < 0)
				return 0;
			if (value <= 0.0031308)
				return value * 12.92;
			if (value < 1)
				return 1.055 * (Math.pow(value, (1 / 2.4))) - 0.055;
			return 1;
		}

		function checkRasterImageId(rasterImageId) {
			var imageLocal = AscCommon.g_oDocumentUrls.getImageLocal(rasterImageId);
			return imageLocal ? imageLocal : rasterImageId;
		}


		var g_oThemeFontsName = {};
		g_oThemeFontsName["+mj-cs"] = true;
		g_oThemeFontsName["+mj-ea"] = true;
		g_oThemeFontsName["+mj-lt"] = true;
		g_oThemeFontsName["+mn-cs"] = true;
		g_oThemeFontsName["+mn-ea"] = true;
		g_oThemeFontsName["+mn-lt"] = true;
		g_oThemeFontsName["majorAscii"] = true;
		g_oThemeFontsName["majorBidi"] = true;
		g_oThemeFontsName["majorEastAsia"] = true;
		g_oThemeFontsName["majorHAnsi"] = true;
		g_oThemeFontsName["minorAscii"] = true;
		g_oThemeFontsName["minorBidi"] = true;
		g_oThemeFontsName["minorEastAsia"] = true;
		g_oThemeFontsName["minorHAnsi"] = true;

		function isRealNumber(n) {
			return typeof n === "number" && !isNaN(n) && isFinite(n);
		}

		function isRealBool(b) {
			return b === true || b === false;
		}

		// Own realization of Object.groupBy for IE11 compatibility
		function groupBy(arr, callback) {
			return arr.reduce(function (storage, item) {
				let group = callback(item);
				storage[group] = storage[group] || [];
				storage[group].splice(0, 0, item);
				return storage;
			}, {});
		}

		function writeLong(w, val) {
			w.WriteBool(isRealNumber(val));
			if (isRealNumber(val)) {
				w.WriteLong(val);
			}
		}

		function readLong(r) {
			var ret;
			if (r.GetBool()) {
				ret = r.GetLong();
			} else {
				ret = null;
			}
			return ret;
		}

		function writeDouble(w, val) {
			w.WriteBool(isRealNumber(val));
			if (isRealNumber(val)) {
				w.WriteDouble(val);
			}
		}

		function readDouble(r) {
			var ret;
			if (r.GetBool()) {
				ret = r.GetDouble();
			} else {
				ret = null;
			}
			return ret;
		}

		function writeDouble2(w, val) {
			const isNumber = typeof val === "number";
			w.WriteBool(isNumber);
			if (isNumber) {
				w.WriteDouble2(val);
			}
		}

		function readDouble2(r) {
			var ret;
			if (r.GetBool()) {
				ret = r.GetDoubleLE();
			} else {
				ret = null;
			}
			return ret;
		}

		function writeBool(w, val) {
			w.WriteBool(isRealBool(val));
			if (isRealBool(val)) {
				w.WriteBool(val);
			}
		}

		function readBool(r) {
			var ret;
			if (r.GetBool()) {
				ret = r.GetBool();
			} else {
				ret = null;
			}
			return ret;
		}

		function writeString(w, val) {
			w.WriteBool(typeof val === "string");
			if (typeof val === "string") {
				w.WriteString2(val);
			}
		}

		function readString(r) {
			var ret;
			if (r.GetBool()) {
				ret = r.GetString2();
			} else {
				ret = null;
			}
			return ret;
		}

		function writeObject(w, val) {
			w.WriteBool(isRealObject(val));
			if (isRealObject(val)) {
				w.WriteString2(val.Get_Id());
			}
		}

		function readObject(r) {
			var ret;
			if (r.GetBool()) {
				ret = g_oTableId.Get_ById(r.GetString2());
			} else {
				ret = null;
			}
			return ret;
		}

		function writeObjectNoId(w, val) {
			const bIsRealObject = isRealObject(val);
			w.WriteBool(bIsRealObject);
			if (bIsRealObject) {
				w.WriteLong(val.getObjectType());
				val.Write_ToBinary(w);
			}
		}

		function readObjectNoId(r) {
			if (r.GetBool()) {
				const val = AscCommon.g_oTableId.GetClassFromFactory(r.GetLong());
				val.Read_FromBinary(r);
				return val;
			}
			return null;
		}
		function writeObjectNoIdNoType(w, val) {
			const bIsRealObject = isRealObject(val);
			w.WriteBool(bIsRealObject);
			if (bIsRealObject) {
				val.Write_ToBinary(w);
			}
		}

		function readObjectNoIdNoType(r, classConstructor) {
			if (r.GetBool()) {
				const val = new classConstructor();
				val.Read_FromBinary(r);
				return val;
			}
			return null;
		}


		function checkThemeFonts(oFontMap, font_scheme) {
			if (oFontMap["+mj-lt"]) {
				if (font_scheme.majorFont && typeof font_scheme.majorFont.latin === "string" && font_scheme.majorFont.latin.length > 0)
					oFontMap[font_scheme.majorFont.latin] = 1;
				delete oFontMap["+mj-lt"];
			}
			if (oFontMap["+mj-ea"]) {
				if (font_scheme.majorFont && typeof font_scheme.majorFont.ea === "string" && font_scheme.majorFont.ea.length > 0)
					oFontMap[font_scheme.majorFont.ea] = 1;
				delete oFontMap["+mj-ea"];
			}
			if (oFontMap["+mj-cs"]) {
				if (font_scheme.majorFont && typeof font_scheme.majorFont.cs === "string" && font_scheme.majorFont.cs.length > 0)
					oFontMap[font_scheme.majorFont.cs] = 1;
				delete oFontMap["+mj-cs"];
			}

			if (oFontMap["+mn-lt"]) {
				if (font_scheme.minorFont && typeof font_scheme.minorFont.latin === "string" && font_scheme.minorFont.latin.length > 0)
					oFontMap[font_scheme.minorFont.latin] = 1;
				delete oFontMap["+mn-lt"];
			}
			if (oFontMap["+mn-ea"]) {
				if (font_scheme.minorFont && typeof font_scheme.minorFont.ea === "string" && font_scheme.minorFont.ea.length > 0)
					oFontMap[font_scheme.minorFont.ea] = 1;
				delete oFontMap["+mn-ea"];
			}
			if (oFontMap["+mn-cs"]) {
				if (font_scheme.minorFont && typeof font_scheme.minorFont.cs === "string" && font_scheme.minorFont.cs.length > 0)
					oFontMap[font_scheme.minorFont.cs] = 1;
				delete oFontMap["+mn-cs"];
			}
		}

		function ExecuteNoHistory(f, oThis, args, notOffTableId) {
			AscCommon.History.TurnOff && AscCommon.History.TurnOff();

			var b_table_id = false;
			if (!notOffTableId && g_oTableId && !g_oTableId.m_bTurnOff) {
				g_oTableId.m_bTurnOff = true;
				b_table_id = true;
			}

			var ret = f.apply(oThis, args);
			AscCommon.History.TurnOn && AscCommon.History.TurnOn();
			if (b_table_id) {
				g_oTableId.m_bTurnOff = false;
			}
			return ret;
		}


		function checkObjectUnifill(obj, theme, colorMap) {
			if (obj && obj.Unifill) {
				obj.Unifill.check(theme, colorMap);
				var rgba = obj.Unifill.getRGBAColor();
				obj.Color = new CDocumentColor(rgba.R, rgba.G, rgba.B, false);
			}
		}

		function checkTableCellPr(cellPr, slide, layout, master, theme) {
			cellPr.Check_PresentationPr(theme);
			var color_map, rgba;
			if (slide.clrMap) {
				color_map = slide.clrMap;
			} else if (layout.clrMap) {
				color_map = layout.clrMap;
			} else if (master.clrMap) {
				color_map = master.clrMap;
			} else {
				color_map = AscFormat.GetDefaultColorMap();
			}

			checkObjectUnifill(cellPr.Shd, theme, color_map);
			if (cellPr.TableCellBorders) {
				checkObjectUnifill(cellPr.TableCellBorders.Left, theme, color_map);
				checkObjectUnifill(cellPr.TableCellBorders.Top, theme, color_map);
				checkObjectUnifill(cellPr.TableCellBorders.Right, theme, color_map);
				checkObjectUnifill(cellPr.TableCellBorders.Bottom, theme, color_map);
				checkObjectUnifill(cellPr.TableCellBorders.InsideH, theme, color_map);
				checkObjectUnifill(cellPr.TableCellBorders.InsideV, theme, color_map);
			}
			return cellPr;
		}

		var Ax_Counter = {
			GLOBAL_AX_ID_COUNTER: 1000
		};
		var TYPE_TRACK = {
			SHAPE: 0,
			GROUP: 0,
			GROUP_PASSIVE: 1,
			TEXT: 2,
			EMPTY_PH: 3,
			CHART_TEXT: 4,
			CROP: 5,
			FORM: 6,
			ANNOT_STAMP: 7
		};
		var TYPE_KIND = {
			SLIDE: 0,
			LAYOUT: 1,
			MASTER: 2,
			NOTES: 3,
			NOTES_MASTER: 4
		};



		var map_hightlight = {};
		map_hightlight["black"] = 0x000000;
		map_hightlight["blue"] = 0x0000FF;
		map_hightlight["cyan"] = 0x00FFFF;
		map_hightlight["darkBlue"] = 0x00008B;
		map_hightlight["darkCyan"] = 0x008B8B;
		map_hightlight["darkGray"] = 0x0A9A9A9;
		map_hightlight["darkGreen"] = 0x006400;
		map_hightlight["darkMagenta"] = 0x800080;
		map_hightlight["darkRed"] = 0x8B0000;
		map_hightlight["darkYellow"] = 0x808000;
		map_hightlight["green"] = 0x00FF00;
		map_hightlight["lightGray"] = 0xD3D3D3;
		map_hightlight["magenta"] = 0xFF00FF;
		map_hightlight["none"] = 0x000000;
		map_hightlight["red"] = 0xFF0000;
		map_hightlight["white"] = 0xFFFFFF;
		map_hightlight["yellow"] = 0xFFFF00;

		var map_prst_color = {};
		map_prst_color["aliceBlue"] = 0xF0F8FF;
		map_prst_color["antiqueWhite"] = 0xFAEBD7;
		map_prst_color["aqua"] = 0x00FFFF;
		map_prst_color["aquamarine"] = 0x7FFFD4;
		map_prst_color["azure"] = 0xF0FFFF;
		map_prst_color["beige"] = 0xF5F5DC;
		map_prst_color["bisque"] = 0xFFE4C4;
		map_prst_color["black"] = 0x000000;
		map_prst_color["blanchedAlmond"] = 0xFFEBCD;
		map_prst_color["blue"] = 0x0000FF;
		map_prst_color["blueViolet"] = 0x8A2BE2;
		map_prst_color["brown"] = 0xA52A2A;
		map_prst_color["burlyWood"] = 0xDEB887;
		map_prst_color["cadetBlue"] = 0x5F9EA0;
		map_prst_color["chartreuse"] = 0x7FFF00;
		map_prst_color["chocolate"] = 0xD2691E;
		map_prst_color["coral"] = 0xFF7F50;
		map_prst_color["cornflowerBlue"] = 0x6495ED;
		map_prst_color["cornsilk"] = 0xFFF8DC;
		map_prst_color["crimson"] = 0xDC143C;
		map_prst_color["cyan"] = 0x00FFFF;
		map_prst_color["darkBlue"] = 0x00008B;
		map_prst_color["darkCyan"] = 0x008B8B;
		map_prst_color["darkGoldenrod"] = 0xB8860B;
		map_prst_color["darkGray"] = 0xA9A9A9;
		map_prst_color["darkGreen"] = 0x006400;
		map_prst_color["darkGrey"] = 0xA9A9A9;
		map_prst_color["darkKhaki"] = 0xBDB76B;
		map_prst_color["darkMagenta"] = 0x8B008B;
		map_prst_color["darkOliveGreen"] = 0x556B2F;
		map_prst_color["darkOrange"] = 0xFF8C00;
		map_prst_color["darkOrchid"] = 0x9932CC;
		map_prst_color["darkRed"] = 0x8B0000;
		map_prst_color["darkSalmon"] = 0xE9967A;
		map_prst_color["darkSeaGreen"] = 0x8FBC8F;
		map_prst_color["darkSlateBlue"] = 0x483D8B;
		map_prst_color["darkSlateGray"] = 0x2F4F4F;
		map_prst_color["darkSlateGrey"] = 0x2F4F4F;
		map_prst_color["darkTurquoise"] = 0x00CED1;
		map_prst_color["darkViolet"] = 0x9400D3;
		map_prst_color["deepPink"] = 0xFF1493;
		map_prst_color["deepSkyBlue"] = 0x00BFFF;
		map_prst_color["dimGray"] = 0x696969;
		map_prst_color["dimGrey"] = 0x696969;
		map_prst_color["dkBlue"] = 0x00008B;
		map_prst_color["dkCyan"] = 0x008B8B;
		map_prst_color["dkGoldenrod"] = 0xB8860B;
		map_prst_color["dkGray"] = 0xA9A9A9;
		map_prst_color["dkGreen"] = 0x006400;
		map_prst_color["dkGrey"] = 0xA9A9A9;
		map_prst_color["dkKhaki"] = 0xBDB76B;
		map_prst_color["dkMagenta"] = 0x8B008B;
		map_prst_color["dkOliveGreen"] = 0x556B2F;
		map_prst_color["dkOrange"] = 0xFF8C00;
		map_prst_color["dkOrchid"] = 0x9932CC;
		map_prst_color["dkRed"] = 0x8B0000;
		map_prst_color["dkSalmon"] = 0xE9967A;
		map_prst_color["dkSeaGreen"] = 0x8FBC8B;
		map_prst_color["dkSlateBlue"] = 0x483D8B;
		map_prst_color["dkSlateGray"] = 0x2F4F4F;
		map_prst_color["dkSlateGrey"] = 0x2F4F4F;
		map_prst_color["dkTurquoise"] = 0x00CED1;
		map_prst_color["dkViolet"] = 0x9400D3;
		map_prst_color["dodgerBlue"] = 0x1E90FF;
		map_prst_color["firebrick"] = 0xB22222;
		map_prst_color["floralWhite"] = 0xFFFAF0;
		map_prst_color["forestGreen"] = 0x228B22;
		map_prst_color["fuchsia"] = 0xFF00FF;
		map_prst_color["gainsboro"] = 0xDCDCDC;
		map_prst_color["ghostWhite"] = 0xF8F8FF;
		map_prst_color["gold"] = 0xFFD700;
		map_prst_color["goldenrod"] = 0xDAA520;
		map_prst_color["gray"] = 0x808080;
		map_prst_color["green"] = 0x008000;
		map_prst_color["greenYellow"] = 0xADFF2F;
		map_prst_color["grey"] = 0x808080;
		map_prst_color["honeydew"] = 0xF0FFF0;
		map_prst_color["hotPink"] = 0xFF69B4;
		map_prst_color["indianRed"] = 0xCD5C5C;
		map_prst_color["indigo"] = 0x4B0082;
		map_prst_color["ivory"] = 0xFFFFF0;
		map_prst_color["khaki"] = 0xF0E68C;
		map_prst_color["lavender"] = 0xE6E6FA;
		map_prst_color["lavenderBlush"] = 0xFFF0F5;
		map_prst_color["lawnGreen"] = 0x7CFC00;
		map_prst_color["lemonChiffon"] = 0xFFFACD;
		map_prst_color["lightBlue"] = 0xADD8E6;
		map_prst_color["lightCoral"] = 0xF08080;
		map_prst_color["lightCyan"] = 0xE0FFFF;
		map_prst_color["lightGoldenrodYellow"] = 0xFAFAD2;
		map_prst_color["lightGray"] = 0xD3D3D3;
		map_prst_color["lightGreen"] = 0x90EE90;
		map_prst_color["lightGrey"] = 0xD3D3D3;
		map_prst_color["lightPink"] = 0xFFB6C1;
		map_prst_color["lightSalmon"] = 0xFFA07A;
		map_prst_color["lightSeaGreen"] = 0x20B2AA;
		map_prst_color["lightSkyBlue"] = 0x87CEFA;
		map_prst_color["lightSlateGray"] = 0x778899;
		map_prst_color["lightSlateGrey"] = 0x778899;
		map_prst_color["lightSteelBlue"] = 0xB0C4DE;
		map_prst_color["lightYellow"] = 0xFFFFE0;
		map_prst_color["lime"] = 0x00FF00;
		map_prst_color["limeGreen"] = 0x32CD32;
		map_prst_color["linen"] = 0xFAF0E6;
		map_prst_color["ltBlue"] = 0xADD8E6;
		map_prst_color["ltCoral"] = 0xF08080;
		map_prst_color["ltCyan"] = 0xE0FFFF;
		map_prst_color["ltGoldenrodYellow"] = 0xFAFA78;
		map_prst_color["ltGray"] = 0xD3D3D3;
		map_prst_color["ltGreen"] = 0x90EE90;
		map_prst_color["ltGrey"] = 0xD3D3D3;
		map_prst_color["ltPink"] = 0xFFB6C1;
		map_prst_color["ltSalmon"] = 0xFFA07A;
		map_prst_color["ltSeaGreen"] = 0x20B2AA;
		map_prst_color["ltSkyBlue"] = 0x87CEFA;
		map_prst_color["ltSlateGray"] = 0x778899;
		map_prst_color["ltSlateGrey"] = 0x778899;
		map_prst_color["ltSteelBlue"] = 0xB0C4DE;
		map_prst_color["ltYellow"] = 0xFFFFE0;
		map_prst_color["magenta"] = 0xFF00FF;
		map_prst_color["maroon"] = 0x800000;
		map_prst_color["medAquamarine"] = 0x66CDAA;
		map_prst_color["medBlue"] = 0x0000CD;
		map_prst_color["mediumAquamarine"] = 0x66CDAA;
		map_prst_color["mediumBlue"] = 0x0000CD;
		map_prst_color["mediumOrchid"] = 0xBA55D3;
		map_prst_color["mediumPurple"] = 0x9370DB;
		map_prst_color["mediumSeaGreen"] = 0x3CB371;
		map_prst_color["mediumSlateBlue"] = 0x7B68EE;
		map_prst_color["mediumSpringGreen"] = 0x00FA9A;
		map_prst_color["mediumTurquoise"] = 0x48D1CC;
		map_prst_color["mediumVioletRed"] = 0xC71585;
		map_prst_color["medOrchid"] = 0xBA55D3;
		map_prst_color["medPurple"] = 0x9370DB;
		map_prst_color["medSeaGreen"] = 0x3CB371;
		map_prst_color["medSlateBlue"] = 0x7B68EE;
		map_prst_color["medSpringGreen"] = 0x00FA9A;
		map_prst_color["medTurquoise"] = 0x48D1CC;
		map_prst_color["medVioletRed"] = 0xC71585;
		map_prst_color["midnightBlue"] = 0x191970;
		map_prst_color["mintCream"] = 0xF5FFFA;
		map_prst_color["mistyRose"] = 0xFFE4FF;
		map_prst_color["moccasin"] = 0xFFE4B5;
		map_prst_color["navajoWhite"] = 0xFFDEAD;
		map_prst_color["navy"] = 0x000080;
		map_prst_color["oldLace"] = 0xFDF5E6;
		map_prst_color["olive"] = 0x808000;
		map_prst_color["oliveDrab"] = 0x6B8E23;
		map_prst_color["orange"] = 0xFFA500;
		map_prst_color["orangeRed"] = 0xFF4500;
		map_prst_color["orchid"] = 0xDA70D6;
		map_prst_color["paleGoldenrod"] = 0xEEE8AA;
		map_prst_color["paleGreen"] = 0x98FB98;
		map_prst_color["paleTurquoise"] = 0xAFEEEE;
		map_prst_color["paleVioletRed"] = 0xDB7093;
		map_prst_color["papayaWhip"] = 0xFFEFD5;
		map_prst_color["peachPuff"] = 0xFFDAB9;
		map_prst_color["peru"] = 0xCD853F;
		map_prst_color["pink"] = 0xFFC0CB;
		map_prst_color["plum"] = 0xD3A0D3;
		map_prst_color["powderBlue"] = 0xB0E0E6;
		map_prst_color["purple"] = 0x800080;
		map_prst_color["red"] = 0xFF0000;
		map_prst_color["rosyBrown"] = 0xBC8F8F;
		map_prst_color["royalBlue"] = 0x4169E1;
		map_prst_color["saddleBrown"] = 0x8B4513;
		map_prst_color["salmon"] = 0xFA8072;
		map_prst_color["sandyBrown"] = 0xF4A460;
		map_prst_color["seaGreen"] = 0x2E8B57;
		map_prst_color["seaShell"] = 0xFFF5EE;
		map_prst_color["sienna"] = 0xA0522D;
		map_prst_color["silver"] = 0xC0C0C0;
		map_prst_color["skyBlue"] = 0x87CEEB;
		map_prst_color["slateBlue"] = 0x6A5AEB;
		map_prst_color["slateGray"] = 0x708090;
		map_prst_color["slateGrey"] = 0x708090;
		map_prst_color["snow"] = 0xFFFAFA;
		map_prst_color["springGreen"] = 0x00FF7F;
		map_prst_color["steelBlue"] = 0x4682B4;
		map_prst_color["tan"] = 0xD2B48C;
		map_prst_color["teal"] = 0x008080;
		map_prst_color["thistle"] = 0xD8BFD8;
		map_prst_color["tomato"] = 0xFF7347;
		map_prst_color["turquoise"] = 0x40E0D0;
		map_prst_color["violet"] = 0xEE82EE;
		map_prst_color["wheat"] = 0xF5DEB3;
		map_prst_color["white"] = 0xFFFFFF;
		map_prst_color["whiteSmoke"] = 0xF5F5F5;
		map_prst_color["yellow"] = 0xFFFF00;
		map_prst_color["yellowGreen"] = 0x9ACD32;

		function CColorMod(sName, nVal) {
			this.name = sName ? sName : "";
			this.val = AscFormat.isRealNumber(nVal) ? nVal : 0;
		}

		CColorMod.prototype.setName = function (name) {
			this.name = name;
		};
		CColorMod.prototype.setVal = function (val) {
			this.val = val;
		};
		CColorMod.prototype.createDuplicate = function () {
			var duplicate = new CColorMod();
			duplicate.name = this.name;
			duplicate.val = this.val;
			return duplicate;
		};
		var cd16 = 1.0 / 6.0;
		var cd13 = 1.0 / 3.0;
		var cd23 = 2.0 / 3.0;
		var max_hls = 255.0;

		var DEC_GAMMA = 2.3;
		var INC_GAMMA = 1.0 / DEC_GAMMA;
		var MAX_PERCENT = 100000;

		function CColorModifiers() {
			this.Mods = [];
		}

		CColorModifiers.prototype.isUsePow = (!AscCommon.AscBrowser.isSailfish || !AscCommon.AscBrowser.isEmulateDevicePixelRatio);
		CColorModifiers.prototype.getModValue = function (sName) {
			let oMod = this.getMod(sName);
			if(oMod) {
				return oMod.val;
			}
			return null;
		};
		CColorModifiers.prototype.getMod = function (sName) {
			if (Array.isArray(this.Mods)) {
				for (let nMod = 0; nMod < this.Mods.length; ++nMod) {
					let oMod = this.Mods[nMod];
					if (oMod && oMod.name === sName) {
						return oMod;
					}
				}
			}
			return null;
		};
		CColorModifiers.prototype.Write_ToBinary = function (w) {
			w.WriteLong(this.Mods.length);
			for (var i = 0; i < this.Mods.length; ++i) {
				w.WriteString2(this.Mods[i].name);
				w.WriteLong(this.Mods[i].val);
			}
		};
		CColorModifiers.prototype.Read_FromBinary = function (r) {
			var len = r.GetLong();
			for (var i = 0; i < len; ++i) {
				var mod = new CColorMod();
				mod.name = r.GetString2();
				mod.val = r.GetLong();
				this.Mods.push(mod);
			}
		};
		CColorModifiers.prototype.addMod = function (mod) {
			let oModForAdd;
			if(arguments.length === 1) {
				if(arguments[0] instanceof CColorMod) {
					oModForAdd = arguments[0];
				}
			}
			else if(arguments.length === 2) {
				if(typeof arguments[0] === "string" && AscFormat.isRealNumber(arguments[1])) {
					oModForAdd = new CColorMod(arguments[0], arguments[1]);
				}
			}
			if(oModForAdd) {
				this.Mods.push(oModForAdd);
			}
		};
		CColorModifiers.prototype.removeMod = function (pos) {
			this.Mods.splice(pos, 1)[0];
		};
		CColorModifiers.prototype.IsIdentical = function (mods) {
			if (mods == null) {
				return false
			}
			if (mods.Mods == null || this.Mods.length !== mods.Mods.length) {
				return false;
			}

			for (var i = 0; i < this.Mods.length; ++i) {
				if (this.Mods[i].name !== mods.Mods[i].name
					|| this.Mods[i].val !== mods.Mods[i].val) {
					return false;
				}
			}
			return true;

		};
		CColorModifiers.prototype.createDuplicate = function () {
			var duplicate = new CColorModifiers();
			for (var i = 0; i < this.Mods.length; ++i) {
				duplicate.Mods[i] = this.Mods[i].createDuplicate();
			}
			return duplicate;
		};
		CColorModifiers.prototype.RGB2HSL = function (R, G, B, HLS) {
			var iMin = (R < G ? R : G);
			iMin = iMin < B ? iMin : B;//Math.min(R, G, B);
			var iMax = (R > G ? R : G);
			iMax = iMax > B ? iMax : B;//Math.max(R, G, B);
			var iDelta = iMax - iMin;
			var dMax = (iMax + iMin) / 255.0;
			var dDelta = iDelta / 255.0;
			var H = 0;
			var S = 0;
			var L = dMax / 2.0;

			if (iDelta != 0) {
				if (L < 0.5) S = dDelta / dMax;
				else S = dDelta / (2.0 - dMax);

				dDelta = dDelta * 1530.0;
				var dR = (iMax - R) / dDelta;
				var dG = (iMax - G) / dDelta;
				var dB = (iMax - B) / dDelta;

				if (R == iMax) H = dB - dG;
				else if (G == iMax) H = cd13 + dR - dB;
				else if (B == iMax) H = cd23 + dG - dR;

				if (H < 0.0) H += 1.0;
				if (H > 1.0) H -= 1.0;
			}

			H = H * max_hls;
			if (H < 0)
				H = 0;
			if (H > 255)
				H = 255;

			S = S * max_hls;
			if (S < 0)
				S = 0;
			if (S > 255)
				S = 255;

			L = L * max_hls;
			if (L < 0)
				L = 0;
			if (L > 255)
				L = 255;

			HLS.H = H;
			HLS.S = S;
			HLS.L = L;
		};
		CColorModifiers.prototype.HSL2RGB = function (HSL, RGB) {
			if (HSL.S == 0) {
				const clampL = AscFormat.ClampColor(HSL.L);
				RGB.R = clampL;
				RGB.G = clampL;
				RGB.B = clampL;
			} else {
				var H = HSL.H / max_hls;
				var S = HSL.S / max_hls;
				var L = HSL.L / max_hls;
				var v2 = 0;
				if (L < 0.5)
					v2 = L * (1.0 + S);
				else
					v2 = L + S - S * L;

				var v1 = 2.0 * L - v2;

				var R = (255 * this.Hue_2_RGB(v1, v2, H + cd13));
				var G = (255 * this.Hue_2_RGB(v1, v2, H));
				var B = (255 * this.Hue_2_RGB(v1, v2, H - cd13));

				RGB.R = AscFormat.ClampColor(R);
				RGB.G = AscFormat.ClampColor(G);
				RGB.B = AscFormat.ClampColor(B);
			}
		};
		CColorModifiers.prototype.Hue_2_RGB = function (v1, v2, vH) {
			if (vH < 0.0)
				vH += 1.0;
			if (vH > 1.0)
				vH -= 1.0;
			if (vH < cd16)
				return v1 + (v2 - v1) * 6.0 * vH;
			if (vH < 0.5)
				return v2;
			if (vH < cd23)
				return v1 + (v2 - v1) * (cd23 - vH) * 6.0;
			return v1;
		};
		CColorModifiers.prototype.lclRgbCompToCrgbComp = function (value) {
			return (value * MAX_PERCENT / 255);
		};
		CColorModifiers.prototype.lclCrgbCompToRgbComp = function (value) {
			return (value * 255 / MAX_PERCENT);
		};
		CColorModifiers.prototype.lclGamma = function (nComp, fGamma) {
			return (Math.pow(nComp / MAX_PERCENT, fGamma) * MAX_PERCENT + 0.5) >> 0;
		};
		CColorModifiers.prototype.RgbtoCrgb = function (RGBA) {
			//RGBA.R = this.lclGamma(this.lclRgbCompToCrgbComp(RGBA.R), DEC_GAMMA);
			//RGBA.G = this.lclGamma(this.lclRgbCompToCrgbComp(RGBA.G), DEC_//GAMMA);
			//RGBA.B = this.lclGamma(this.lclRgbCompToCrgbComp(RGBA.B), DEC_GAMMA);

			if (this.isUsePow) {
				RGBA.R = (Math.pow(RGBA.R / 255, DEC_GAMMA) * MAX_PERCENT + 0.5) >> 0;
				RGBA.G = (Math.pow(RGBA.G / 255, DEC_GAMMA) * MAX_PERCENT + 0.5) >> 0;
				RGBA.B = (Math.pow(RGBA.B / 255, DEC_GAMMA) * MAX_PERCENT + 0.5) >> 0;
			}
		};
		CColorModifiers.prototype.CrgbtoRgb = function (RGBA) {
			//RGBA.R = (this.lclCrgbCompToRgbComp(this.lclGamma(RGBA.R, INC_GAMMA)) + 0.5) >> 0;
			//RGBA.G = (this.lclCrgbCompToRgbComp(this.lclGamma(RGBA.G, INC_GAMMA)) + 0.5) >> 0;
			//RGBA.B = (this.lclCrgbCompToRgbComp(this.lclGamma(RGBA.B, INC_GAMMA)) + 0.5) >> 0;

			if (this.isUsePow) {
				RGBA.R = Math.pow(RGBA.R / 100000, INC_GAMMA) * 255;
				RGBA.G = Math.pow(RGBA.G / 100000, INC_GAMMA) * 255;
				RGBA.B = Math.pow(RGBA.B / 100000, INC_GAMMA) * 255;
			}
			RGBA.R = AscFormat.ClampColor(RGBA.R);
			RGBA.G = AscFormat.ClampColor(RGBA.G);
			RGBA.B = AscFormat.ClampColor(RGBA.B);
		};
		CColorModifiers.prototype.Apply = function (RGBA) {
			if (null == this.Mods)
				return;

			const _len = this.Mods.length;
			for (let i = 0; i < _len; i++) {
				const colorMod = this.Mods[i];
				let val = colorMod.val / 100000.0;

				if (colorMod.name === "alpha") {
					RGBA.A = AscFormat.ClampColor(255 * val);
				} else if (colorMod.name === "blue") {
					RGBA.B = AscFormat.ClampColor(255 * val);
				} else if (colorMod.name === "blueMod") {
					RGBA.B = AscFormat.ClampColor(RGBA.B * val);
				} else if (colorMod.name === "blueOff") {
					RGBA.B = AscFormat.ClampColor(RGBA.B + val * 255);
				} else if (colorMod.name === "green") {
					RGBA.G = AscFormat.ClampColor(255 * val);
				} else if (colorMod.name === "greenMod") {
					RGBA.G = AscFormat.ClampColor(RGBA.G * val);
				} else if (colorMod.name === "greenOff") {
					RGBA.G = AscFormat.ClampColor(RGBA.G + val * 255);
				} else if (colorMod.name === "red") {
					RGBA.R = AscFormat.ClampColor(255 * val);
				} else if (colorMod.name === "redMod") {
					RGBA.R = AscFormat.ClampColor(RGBA.R * val);
				} else if (colorMod.name === "redOff") {
					RGBA.R = AscFormat.ClampColor(RGBA.R + val * 255);
				} else if (colorMod.name === "hueMod") {
					if (val === 1) {
						continue;
					}
					const HSL = {H: 0, S: 0, L: 0};
					this.RGB2HSL(RGBA.R, RGBA.G, RGBA.B, HSL);
					HSL.H = AscCommon.trimMinMaxValue(HSL.H * val, 0, max_hls);

					this.HSL2RGB(HSL, RGBA);
				} else if (colorMod.name === "hueOff") {
					if (val === 0) {
						continue;
					}
					const HSL = {H: 0, S: 0, L: 0};
					this.RGB2HSL(RGBA.R, RGBA.G, RGBA.B, HSL);
					val = (colorMod.val / 60000) * (max_hls / 360);
					const res = HSL.H + val;
					HSL.H = AscCommon.trimMinMaxValue(res, 0, max_hls);

					this.HSL2RGB(HSL, RGBA);
				} else if (colorMod.name === "inv") {
					RGBA.R ^= 0xFF;
					RGBA.G ^= 0xFF;
					RGBA.B ^= 0xFF;
				} else if (colorMod.name === "lumMod") {
					if (val === 1) {
						continue;
					}
					const HSL = {H: 0, S: 0, L: 0};
					this.RGB2HSL(RGBA.R, RGBA.G, RGBA.B, HSL);

					HSL.L = AscCommon.trimMinMaxValue(HSL.L * val, 0, max_hls);
					this.HSL2RGB(HSL, RGBA);
				} else if (colorMod.name === "lumOff") {
					if (val === 0) {
						continue;
					}
					const HSL = {H: 0, S: 0, L: 0};
					this.RGB2HSL(RGBA.R, RGBA.G, RGBA.B, HSL);

					const res = HSL.L + val * max_hls;
					HSL.L = AscCommon.trimMinMaxValue(res, 0, max_hls);

					this.HSL2RGB(HSL, RGBA);
				} else if (colorMod.name === "satMod") {
					if (val === 1) {
						continue;
					}
					const HSL = {H: 0, S: 0, L: 0};
					this.RGB2HSL(RGBA.R, RGBA.G, RGBA.B, HSL);

					HSL.S = AscCommon.trimMinMaxValue(HSL.S * val, 0, max_hls);
					this.HSL2RGB(HSL, RGBA);
				} else if (colorMod.name === "satOff") {
					if (val === 0) {
						continue;
					}
					const HSL = {H: 0, S: 0, L: 0};
					this.RGB2HSL(RGBA.R, RGBA.G, RGBA.B, HSL);
					const res = HSL.S + val * max_hls;
					HSL.S = AscCommon.trimMinMaxValue(res, 0, max_hls);
					this.HSL2RGB(HSL, RGBA);
				} else if (colorMod.name === "wordShade") {
					if (colorMod.val === 255) {
						continue;
					}
					const val_ = colorMod.val / 255;
					//GBA.R = Math.max(0, (RGBA.R * (1 - val_)) >> 0);
					//GBA.G = Math.max(0, (RGBA.G * (1 - val_)) >> 0);
					//GBA.B = Math.max(0, (RGBA.B * (1 - val_)) >> 0);


					//RGBA.R = Math.max(0,  ((1 - val_)*(- RGBA.R) + RGBA.R) >> 0);
					//RGBA.G = Math.max(0,  ((1 - val_)*(- RGBA.G) + RGBA.G) >> 0);
					//RGBA.B = Math.max(0,  ((1 - val_)*(- RGBA.B) + RGBA.B) >> 0);

					const HSL = {H: 0, S: 0, L: 0};
					this.RGB2HSL(RGBA.R, RGBA.G, RGBA.B, HSL);

					HSL.L = AscCommon.trimMinMaxValue(HSL.L * val_, 0, max_hls);
					this.HSL2RGB(HSL, RGBA);
				} else if (colorMod.name === "wordTint") {
					if (colorMod.val === 255) {
						continue;
					}
					const _val = colorMod.val / 255;
					//RGBA.R = Math.max(0,  ((1 - _val)*(255 - RGBA.R) + RGBA.R) >> 0);
					//RGBA.G = Math.max(0,  ((1 - _val)*(255 - RGBA.G) + RGBA.G) >> 0);
					//RGBA.B = Math.max(0,  ((1 - _val)*(255 - RGBA.B) + RGBA.B) >> 0);

					const HSL = {H: 0, S: 0, L: 0};
					this.RGB2HSL(RGBA.R, RGBA.G, RGBA.B, HSL);

					const L_ = HSL.L * _val + (255 - colorMod.val);
					HSL.L = AscCommon.trimMinMaxValue(L_, 0, max_hls);
					this.HSL2RGB(HSL, RGBA);
				} else if (colorMod.name === "shade") {
					this.RgbtoCrgb(RGBA);
					if (val < 0) val = 0;
					if (val > 1) val = 1;
					RGBA.R = (RGBA.R * val);
					RGBA.G = (RGBA.G * val);
					RGBA.B = (RGBA.B * val);
					this.CrgbtoRgb(RGBA);
				} else if (colorMod.name === "tint") {
					this.RgbtoCrgb(RGBA);
					if (val < 0) val = 0;
					if (val > 1) val = 1;
					RGBA.R = (MAX_PERCENT - (MAX_PERCENT - RGBA.R) * val);
					RGBA.G = (MAX_PERCENT - (MAX_PERCENT - RGBA.G) * val);
					RGBA.B = (MAX_PERCENT - (MAX_PERCENT - RGBA.B) * val);
					this.CrgbtoRgb(RGBA);
				} else if (colorMod.name === "gamma") {
					this.RgbtoCrgb(RGBA);
					RGBA.R = this.lclGamma(RGBA.R, INC_GAMMA);
					RGBA.G = this.lclGamma(RGBA.G, INC_GAMMA);
					RGBA.B = this.lclGamma(RGBA.B, INC_GAMMA);
					this.CrgbtoRgb(RGBA);
				} else if (colorMod.name === "invGamma") {
					this.RgbtoCrgb(RGBA);
					RGBA.R = this.lclGamma(RGBA.R, DEC_GAMMA);
					RGBA.G = this.lclGamma(RGBA.G, DEC_GAMMA);
					RGBA.B = this.lclGamma(RGBA.B, DEC_GAMMA);
					this.CrgbtoRgb(RGBA);
				}
			}
		};
		CColorModifiers.prototype.Merge = function (oOther) {
			if (!oOther) {
				return;
			}
			this.Mods = oOther.Mods.concat(this.Mods);
		};
		CColorModifiers.prototype.getShadeOrTint = function() {
			const M = this.Mods;
			if(M.length === 1 && M[0].name === "lumMod" && M[0].val > 0) {//shade
				return -(100000 - M[0].val);
			}
			if(M.length === 2 && M[0].name === "lumMod" && M[0].val > 0 && M[1].name === "lumOff" && M[1].val > 0) {
				return (100000 - M[0].val);
			}
			return null;
		};
		CColorModifiers.prototype.canGetShadeOrTint = function() {
			return this.getShadeOrTint() !== null;
		};
		CColorModifiers.prototype.getEffectValue = function () {
			if(this.Mods.length === 1) {
				let oMod = this.Mods[0];
				if(oMod.name === "wordTint") {
					return (255 - oMod.val) / 255;
				}
				if(oMod.name === "wordShade") {
					return -(255 - oMod.val) / 255;
				}
			}
			let nVal = this.getShadeOrTint();
			if(nVal !== null) {
				return nVal / 100000;
			}
			return 0;
		};


		function getPercentageValue(sVal) {
			var _len = sVal.length;
			if (_len === 0)
				return null;

			var _ret = null;
			if ((_len - 1) === sVal.indexOf("%")) {
				sVal.substring(0, _len - 1);
				_ret = parseFloat(sVal);
				if (isNaN(_ret))
					_ret = null;
			} else {
				_ret = parseFloat(sVal);
				if (isNaN(_ret))
					_ret = null;
				else
					_ret /= 1000;
			}
			return _ret;

			// let nVal = 0;
			// if (sVal.indexOf("%") > -1) {
			// 	let nPct = parseInt(sVal.slice(0, sVal.length - 1));
			// 	if (AscFormat.isRealNumber(nPct)) {
			// 		return ((100000 * nPct / 100 + 0.5 >> 0) - 1);
			// 	}
			// 	return 0;
			// }
			// let nValPct = parseInt(sVal);
			// if (AscFormat.isRealNumber(nValPct)) {
			// 	return nValPct;
			// }
			// return 0;
		}

		function getPercentageValueForWrite(dVal) {
			if (!AscFormat.isRealNumber(dVal)) {
				return null;
			}
			return (dVal * 1000 + 0.5 >> 0);
		}

		function CBaseColor() {
			CBaseNoIdObject.call(this);
			this.RGBA = {
				R: 0,
				G: 0,
				B: 0,
				A: 255,
				needRecalc: true
			};
			this.Mods = null; //[];
		}

		InitClass(CBaseColor, CBaseNoIdObject, 0);
		CBaseColor.prototype.type = c_oAscColor.COLOR_TYPE_NONE;
		CBaseColor.prototype.setR = function (pr) {
			this.RGBA.R = pr;
		};
		CBaseColor.prototype.setG = function (pr) {
			this.RGBA.G = pr;
		};
		CBaseColor.prototype.setB = function (pr) {
			this.RGBA.B = pr;
		};
		CBaseColor.prototype.readModifier = function (name, reader) {
			if (MODS_MAP[name]) {
				var oMod = new CColorMod();
				oMod.name = name;
				while (reader.MoveToNextAttribute()) {
					if (reader.GetNameNoNS() === "val") {
						let sVal = reader.GetValue();
						let nLen = sVal.length;
						if(typeof sVal === "string" && nLen > 0) {
							if ((nLen - 1) === sVal.indexOf("%")) {
								sVal.substring(0, nLen - 1);
								let dVal = parseFloat(sVal);
								if (AscFormat.isRealNumber(dVal))
									oMod.val = dVal * 1000 + 0.5 >> 0;
							}
							else {
								let dVal = parseFloat(sVal);
								if(AscFormat.isRealNumber(dVal)) {
									oMod.val = dVal + 0.5 >> 0;
								}
							}
						}
						break;
					}
				}
				if (!Array.isArray(this.Mods)) {
					this.Mods = [];
				}
				this.Mods.push(oMod);
			}
			return false;
		};
		CBaseColor.prototype.getChannelValue = function (sVal) {
			let nValPct = getPercentageValue(sVal);
			return (255 * nValPct / 100000 + 0.5 >> 0);
		};
		CBaseColor.prototype.getTypeName = function () {
			return "";
		};


		const COLOR_3DDKSHADOW               = 21;
		const COLOR_3DFACE                   = 15;
		const COLOR_3DHIGHLIGHT              = 20;
		const COLOR_3DHILIGHT                = 20;
		const COLOR_3DLIGHT                  = 22;
		const COLOR_3DSHADOW                 = 16;
		const COLOR_ACTIVEBORDER             = 10;
		const COLOR_ACTIVECAPTION            = 2;
		const COLOR_APPWORKSPACE             = 12;
		const COLOR_BACKGROUND               = 1;
		const COLOR_BTNFACE                  = 15;
		const COLOR_BTNHIGHLIGHT             = 20;
		const COLOR_BTNHILIGHT               = 20;
		const COLOR_BTNSHADOW                = 16;
		const COLOR_BTNTEXT                  = 18;
		const COLOR_CAPTIONTEXT              = 9;
		const COLOR_DESKTOP                  = 1;
		const COLOR_GRAYTEXT                 = 17;
		const COLOR_HIGHLIGHT                = 13;
		const COLOR_HIGHLIGHTTEXT            = 14;
		const COLOR_HOTLIGHT                 = 26;
		const COLOR_INACTIVEBORDER           = 11;
		const COLOR_INACTIVECAPTION          = 3;
		const COLOR_INACTIVECAPTIONTEXT      = 19;
		const COLOR_INFOBK                   = 24;
		const COLOR_INFOTEXT                 = 23;
		const COLOR_MENU                     = 4;
		const COLOR_GRADIENTACTIVECAPTION    = 27;
		const COLOR_GRADIENTINACTIVECAPTION  = 28;
		const COLOR_MENUHILIGHT              = 29;
		const COLOR_MENUBAR                  = 30;
		const COLOR_MENUTEXT                 = 7;
		const COLOR_SCROLLBAR                = 0;
		const COLOR_WINDOW                   = 5;
		const COLOR_WINDOWFRAME              = 6;
		const COLOR_WINDOWTEXT               = 8;


		function GetSysColor(nIndex)
		{
			// get color values from any windows theme
			// http://msdn.microsoft.com/en-us/library/windows/desktop/ms724371(v=vs.85).aspx
			//***************** GetSysColor values begin (Win7 x64) *****************
			let nValue = 0x0;
			//***************** GetSysColor values begin (Win7 x64) *****************
			switch (nIndex) {
				case COLOR_3DDKSHADOW: nValue = 0x696969; break;
				case COLOR_3DFACE: nValue = 0xf0f0f0; break;
				case COLOR_3DHIGHLIGHT: nValue = 0xffffff; break;
				// case COLOR_3DHILIGHT: nValue = 0xffffff; break; // is COLOR_3DHIGHLIGHT
				case COLOR_3DLIGHT: nValue = 0xe3e3e3; break;
				case COLOR_3DSHADOW: nValue = 0xa0a0a0; break;
				case COLOR_ACTIVEBORDER: nValue = 0xb4b4b4; break;
				case COLOR_ACTIVECAPTION: nValue = 0xd1b499; break;
				case COLOR_APPWORKSPACE: nValue = 0xababab; break;
				case COLOR_BACKGROUND: nValue = 0x0; break;
				// case COLOR_BTNFACE: nValue = 0xf0f0f0; break; // is COLOR_3DFACE
				// case COLOR_BTNHIGHLIGHT: nValue = 0xffffff; break; // is COLOR_3DHIGHLIGHT
				// case COLOR_BTNHILIGHT: nValue = 0xffffff; break; // is COLOR_3DHIGHLIGHT
				// case COLOR_BTNSHADOW: nValue = 0xa0a0a0; break; // is COLOR_3DSHADOW
				case COLOR_BTNTEXT: nValue = 0x0; break;
				case COLOR_CAPTIONTEXT: nValue = 0x0; break;
				// case COLOR_DESKTOP: nValue = 0x0; break; // is COLOR_BACKGROUND
				case COLOR_GRADIENTACTIVECAPTION: nValue = 0xead1b9; break;
				case COLOR_GRADIENTINACTIVECAPTION: nValue = 0xf2e4d7; break;
				case COLOR_GRAYTEXT: nValue = 0x6d6d6d; break;
				case COLOR_HIGHLIGHT: nValue = 0xff9933; break;
				case COLOR_HIGHLIGHTTEXT: nValue = 0xffffff; break;
				case COLOR_HOTLIGHT: nValue = 0xcc6600; break;
				case COLOR_INACTIVEBORDER: nValue = 0xfcf7f4; break;
				case COLOR_INACTIVECAPTION: nValue = 0xdbcdbf; break;
				case COLOR_INACTIVECAPTIONTEXT: nValue = 0x544e43; break;
				case COLOR_INFOBK: nValue = 0xe1ffff; break;
				case COLOR_INFOTEXT: nValue = 0x0; break;
				case COLOR_MENU: nValue = 0xf0f0f0; break;
				case COLOR_MENUHILIGHT: nValue = 0xff9933; break;
				case COLOR_MENUBAR: nValue = 0xf0f0f0; break;
				case COLOR_MENUTEXT: nValue = 0x0; break;
				case COLOR_SCROLLBAR: nValue = 0xc8c8c8; break;
				case COLOR_WINDOW: nValue = 0xffffff; break;
				case COLOR_WINDOWFRAME: nValue = 0x646464; break;
				case COLOR_WINDOWTEXT: nValue = 0x0; break;
				default: nValue = 0x0; break;
			} // switch (nIndex)
			//***************** GetSysColor values end *****************
			return nValue;
		}

		function CSysColor() {
			CBaseColor.call(this);
			this.id = "";
		}

		InitClass(CSysColor, CBaseColor, 0);
		CSysColor.prototype.type = c_oAscColor.COLOR_TYPE_SYS;
		CSysColor.prototype.check = function () {
			var ret = this.RGBA.needRecalc;
			this.RGBA.needRecalc = false;
			return ret;
		};
		CSysColor.prototype.Write_ToBinary = function (w) {
			w.WriteLong(this.type);
			w.WriteString2(this.id);
			w.WriteLong(((this.RGBA.R << 16) & 0xFF0000) + ((this.RGBA.G << 8) & 0xFF00) + this.RGBA.B);
		};
		CSysColor.prototype.Read_FromBinary = function (r) {
			this.id = r.GetString2();

			var RGB = r.GetLong();
			this.RGBA.R = (RGB >> 16) & 0xFF;
			this.RGBA.G = (RGB >> 8) & 0xFF;
			this.RGBA.B = RGB & 0xFF;
		};
		CSysColor.prototype.setId = function (id) {
			this.id = id;
		};
		CSysColor.prototype.IsIdentical = function (color) {
			return color && color.type === this.type && color.id === this.id;
		};
		CSysColor.prototype.Calculate = function (obj) {
		};
		CSysColor.prototype.createDuplicate = function () {
			var duplicate = new CSysColor();
			duplicate.id = this.id;
			duplicate.RGBA.R = this.RGBA.R;
			duplicate.RGBA.G = this.RGBA.G;
			duplicate.RGBA.B = this.RGBA.B;
			duplicate.RGBA.A = this.RGBA.A;
			return duplicate;
		};
		CSysColor.prototype.FillRGBFromVal = function(str) {
			let RGB = 0;
			if (str && str !== "") {
				switch (str.charAt(0)) {
					case '3':
						if (str === ("3dDkShadow")) {
							RGB = GetSysColor(COLOR_3DDKSHADOW);
							break;
						}
						if (str === ("3dLight")) {
							RGB = GetSysColor(COLOR_3DLIGHT);
							break;
						}
						break;
					case 'a':
						if (str === ("activeBorder")) {
							RGB = GetSysColor(COLOR_ACTIVEBORDER);
							break;
						}
						if (str === ("activeCaption")) {
							RGB = GetSysColor(COLOR_ACTIVECAPTION);
							break;
						}
						if (str === ("appWorkspace")) {
							RGB = GetSysColor(COLOR_APPWORKSPACE);
							break;
						}
						break;
					case 'b':
						if (str === ("background")) {
							RGB = GetSysColor(COLOR_BACKGROUND);
							break;
						}
						if (str === ("btnFace")) {
							RGB = GetSysColor(COLOR_BTNFACE);
							break;
						}
						if (str === ("btnHighlight")) {
							RGB = GetSysColor(COLOR_BTNHIGHLIGHT);
							break;
						}
						if (str === ("btnShadow")) {
							RGB = GetSysColor(COLOR_BTNSHADOW);
							break;
						}
						if (str === ("btnText")) {
							RGB = GetSysColor(COLOR_BTNTEXT);
							break;
						}
						break;
					case 'c':
						if (str === ("captionText")) {
							RGB = GetSysColor(COLOR_CAPTIONTEXT);
							break;
						}
						break;
					case 'g':
						if (str === ("gradientActiveCaption")) {
							RGB = GetSysColor(COLOR_GRADIENTACTIVECAPTION);
							break;
						}
						if (str === ("gradientInactiveCaption")) {
							RGB = GetSysColor(COLOR_GRADIENTINACTIVECAPTION);
							break;
						}
						if (str === ("grayText")) {
							RGB = GetSysColor(COLOR_GRAYTEXT);
							break;
						}
						break;
					case 'h':
						if (str === ("highlight")) {
							RGB = GetSysColor(COLOR_HIGHLIGHT);
							break;
						}
						if (str === ("highlightText")) {
							RGB = GetSysColor(COLOR_HIGHLIGHTTEXT);
							break;
						}
						if (str === ("hotLight")) {
							RGB = GetSysColor(COLOR_HOTLIGHT);
							break;
						}
						break;
					case 'i':
						if (str === ("inactiveBorder")) {
							RGB = GetSysColor(COLOR_INACTIVEBORDER);
							break;
						}
						if (str === ("inactiveCaption")) {
							RGB = GetSysColor(COLOR_INACTIVECAPTION);
							break;
						}
						if (str === ("inactiveCaptionText")) {
							RGB = GetSysColor(COLOR_INACTIVECAPTIONTEXT);
							break;
						}
						if (str === ("infoBk")) {
							RGB = GetSysColor(COLOR_INFOBK);
							break;
						}
						if (str === ("infoText")) {
							RGB = GetSysColor(COLOR_INFOTEXT);
							break;
						}
						break;
					case 'm':
						if (str === ("menu")) {
							RGB = GetSysColor(COLOR_MENU);
							break;
						}
						if (str === ("menuBar")) {
							RGB = GetSysColor(COLOR_MENUBAR);
							break;
						}
						if (str === ("menuHighlight")) {
							RGB = GetSysColor(COLOR_MENUHILIGHT);
							break;
						}
						if (str === ("menuText")) {
							RGB = GetSysColor(COLOR_MENUTEXT);
							break;
						}
						break;
					case 's':
						if (str === ("scrollBar")) {
							RGB = GetSysColor(COLOR_SCROLLBAR);
							break;
						}
						break;
					case 'w':
						if (str === ("window")) {
							RGB = GetSysColor(COLOR_WINDOW);
							break;
						}
						if (str === ("windowFrame")) {
							RGB = GetSysColor(COLOR_WINDOWFRAME);
							break;
						}
						if (str === ("windowText")) {
							RGB = GetSysColor(COLOR_WINDOWTEXT);
							break;
						}
						break;
				}
			}


			this.RGBA.R = (RGB >> 16) & 0xFF;
			this.RGBA.G = (RGB >> 8) & 0xFF;
			this.RGBA.B = RGB & 0xFF;
		}

		function CPrstColor() {
			CBaseColor.call(this);
			this.id = "";
		}

		InitClass(CPrstColor, CBaseColor, 0);
		CPrstColor.prototype.type = c_oAscColor.COLOR_TYPE_PRST;
		CPrstColor.prototype.Write_ToBinary = function (w) {
			w.WriteLong(this.type);
			w.WriteString2(this.id);
		};
		CPrstColor.prototype.Read_FromBinary = function (r) {
			this.id = r.GetString2();
		};
		CPrstColor.prototype.setId = function (id) {
			this.id = id;
		};
		CPrstColor.prototype.IsIdentical = function (color) {
			return color && color.type === this.type && color.id === this.id;
		};
		CPrstColor.prototype.createDuplicate = function () {
			var duplicate = new CPrstColor();
			duplicate.id = this.id;
			duplicate.RGBA.R = this.RGBA.R;
			duplicate.RGBA.G = this.RGBA.G;
			duplicate.RGBA.B = this.RGBA.B;
			duplicate.RGBA.A = this.RGBA.A;
			return duplicate;
		};
		CPrstColor.prototype.Calculate = function (obj) {
			var RGB = map_prst_color[this.id];
			this.RGBA.R = (RGB >> 16) & 0xFF;
			this.RGBA.G = (RGB >> 8) & 0xFF;
			this.RGBA.B = RGB & 0xFF;
		};
		CPrstColor.prototype.check = function () {
			var r, g, b, rgb;
			rgb = map_prst_color[this.id];
			r = (rgb >> 16) & 0xFF;
			g = (rgb >> 8) & 0xFF;
			b = rgb & 0xFF;

			var RGBA = this.RGBA;
			if (RGBA.needRecalc) {
				RGBA.R = r;
				RGBA.G = g;
				RGBA.B = b;
				RGBA.needRecalc = false;
				return true;
			} else {
				if (RGBA.R === r && RGBA.G === g && RGBA.B === b)
					return false;
				else {
					RGBA.R = r;
					RGBA.G = g;
					RGBA.B = b;
					return true;
				}
			}
		};
		var MODS_MAP = {};
		MODS_MAP["alpha"] = true;
		MODS_MAP["alphaMod"] = true;
		MODS_MAP["alphaOff"] = true;
		MODS_MAP["blue"] = true;
		MODS_MAP["blueMod"] = true;
		MODS_MAP["blueOff"] = true;
		MODS_MAP["comp"] = true;
		MODS_MAP["gamma"] = true;
		MODS_MAP["gray"] = true;
		MODS_MAP["green"] = true;
		MODS_MAP["greenMod"] = true;
		MODS_MAP["greenOff"] = true;
		MODS_MAP["hue"] = true;
		MODS_MAP["hueMod"] = true;
		MODS_MAP["hueOff"] = true;
		MODS_MAP["inv"] = true;
		MODS_MAP["invGamma"] = true;
		MODS_MAP["lum"] = true;
		MODS_MAP["lumMod"] = true;
		MODS_MAP["lumOff"] = true;
		MODS_MAP["red"] = true;
		MODS_MAP["redMod"] = true;
		MODS_MAP["redOff"] = true;
		MODS_MAP["sat"] = true;
		MODS_MAP["satMod"] = true;
		MODS_MAP["satOff"] = true;
		MODS_MAP["shade"] = true;
		MODS_MAP["tint"] = true;

		function toHex(c) {
			var res = Number(c).toString(16).toUpperCase();
			return res.length === 1 ? "0" + res : res;
		}

		function fRGBAToHexString(oRGBA) {
			return "" + toHex(oRGBA.R) + toHex(oRGBA.G) + toHex(oRGBA.B);
		}

		function CRGBColor() {
			CBaseColor.call(this);

			this.h = null;
			this.s = null;
			this.l = null;
		}

		InitClass(CRGBColor, CBaseColor, 0);
		CRGBColor.prototype.type = c_oAscColor.COLOR_TYPE_SRGB;
		CRGBColor.prototype.check = function () {
			var ret = this.RGBA.needRecalc;
			this.RGBA.needRecalc = false;
			return ret;
		};
		CRGBColor.prototype.writeToBinaryLong = function (w) {
			w.WriteLong(((this.RGBA.R << 16) & 0xFF0000) + ((this.RGBA.G << 8) & 0xFF00) + this.RGBA.B);
		};
		CRGBColor.prototype.readFromBinaryLong = function (r) {
			var RGB = r.GetLong();
			this.RGBA.R = (RGB >> 16) & 0xFF;
			this.RGBA.G = (RGB >> 8) & 0xFF;
			this.RGBA.B = RGB & 0xFF;
		};
		CRGBColor.prototype.Write_ToBinary = function (w) {
			w.WriteLong(this.type);
			w.WriteLong(((this.RGBA.R << 16) & 0xFF0000) + ((this.RGBA.G << 8) & 0xFF00) + this.RGBA.B);
		};
		CRGBColor.prototype.Read_FromBinary = function (r) {
			var RGB = r.GetLong();
			this.RGBA.R = (RGB >> 16) & 0xFF;
			this.RGBA.G = (RGB >> 8) & 0xFF;
			this.RGBA.B = RGB & 0xFF;
		};
		CRGBColor.prototype.setColor = function (r, g, b) {
			this.RGBA.R = r;
			this.RGBA.G = g;
			this.RGBA.B = b;
		};
		CRGBColor.prototype.IsIdentical = function (color) {
			return color && color.type === this.type && color.RGBA.R === this.RGBA.R && color.RGBA.G === this.RGBA.G && color.RGBA.B === this.RGBA.B && color.RGBA.A === this.RGBA.A;
		};
		CRGBColor.prototype.createDuplicate = function () {
			var duplicate = new CRGBColor();
			duplicate.id = this.id;
			duplicate.RGBA.R = this.RGBA.R;
			duplicate.RGBA.G = this.RGBA.G;
			duplicate.RGBA.B = this.RGBA.B;
			duplicate.RGBA.A = this.RGBA.A;
			return duplicate;
		};
		CRGBColor.prototype.Calculate = function (obj) {
		};
		CRGBColor.prototype.checkHSL = function () {
			if (this.h !== null && this.s !== null && this.l !== null) {
				CColorModifiers.prototype.HSL2RGB.call(this, {H: this.h, S: this.s, L: this.l}, this.RGBA);
				this.h = null;
				this.s = null;
				this.l = null;
			}
		};
		CRGBColor.prototype.fromScRgb = function() {
			this.RGBA.R = 255 * this.scRGB_to_sRGB(this.RGBA.R);
			this.RGBA.G = 255 * this.scRGB_to_sRGB(this.RGBA.G);
			this.RGBA.B = 255 * this.scRGB_to_sRGB(this.RGBA.B);
		};
		CRGBColor.prototype.scRGB_to_sRGB = function(value) {
			if( value < 0)
				return 0;
			if(value <= 0.0031308)
				return value * 12.92;
			if(value < 1)
				return 1.055 * (Math.pow(value , (1 / 2.4))) - 0.055;
			return 1;
		};
		function CSchemeColor() {
			CBaseColor.call(this);
			this.id = 0;
		}

		InitClass(CSchemeColor, CBaseColor, 0);
		CSchemeColor.prototype.type = c_oAscColor.COLOR_TYPE_SCHEME;
		CSchemeColor.prototype.check = function (theme, colorMap) {
			var RGBA, colors = theme.themeElements.clrScheme.colors;
			if (colorMap[this.id] != null && colors[colorMap[this.id]] != null)
				RGBA = colors[colorMap[this.id]].color.RGBA;
			else if (colors[this.id] != null)
				RGBA = colors[this.id].color.RGBA;
			if (!RGBA) {
				RGBA = {R: 0, G: 0, B: 0, A: 255};
			}
			var _RGBA = this.RGBA;
			if (this.RGBA.needRecalc) {
				_RGBA.R = RGBA.R;
				_RGBA.G = RGBA.G;
				_RGBA.B = RGBA.B;
				_RGBA.A = RGBA.A;
				this.RGBA.needRecalc = false;
				return true;
			} else {
				if (_RGBA.R === RGBA.R && _RGBA.G === RGBA.G && _RGBA.B === RGBA.B && _RGBA.A === RGBA.A) {
					return false;
				} else {
					_RGBA.R = RGBA.R;
					_RGBA.G = RGBA.G;
					_RGBA.B = RGBA.B;
					_RGBA.A = RGBA.A;
					return true;
				}
			}
		};
		CSchemeColor.prototype.getObjectType = function () {
			return AscDFH.historyitem_type_SchemeColor;
		};
		CSchemeColor.prototype.Write_ToBinary = function (w) {
			w.WriteLong(this.type);
			w.WriteLong(this.id);
		};
		CSchemeColor.prototype.Read_FromBinary = function (r) {
			this.id = r.GetLong();
		};
		CSchemeColor.prototype.setId = function (id) {
			this.id = id;
		};
		CSchemeColor.prototype.IsIdentical = function (color) {
			return color && color.type === this.type && color.id === this.id;
		};
		CSchemeColor.prototype.createDuplicate = function () {
			var duplicate = new CSchemeColor();
			duplicate.id = this.id;
			duplicate.RGBA.R = this.RGBA.R;
			duplicate.RGBA.G = this.RGBA.G;
			duplicate.RGBA.B = this.RGBA.B;
			duplicate.RGBA.A = this.RGBA.A;
			return duplicate;
		};
		CSchemeColor.prototype.Calculate = function (theme, slide, layout, masterSlide, RGBA, colorMap) {
			if (theme && theme.themeElements && theme.themeElements.clrScheme) {
				if (this.id === phClr) {
					this.RGBA = RGBA;
				} else {
					var clrMap;
					if (colorMap && colorMap.color_map) {
						clrMap = colorMap.color_map;
					} else if (slide != null && slide.clrMap != null) {
						clrMap = slide.clrMap.color_map;
					} else if (layout != null && layout.clrMap != null) {
						clrMap = layout.clrMap.color_map;
					} else if (masterSlide != null && masterSlide.clrMap != null) {
						clrMap = masterSlide.clrMap.color_map;
					} else {
						clrMap = AscFormat.GetDefaultColorMap().color_map;
					}
					if (clrMap[this.id] != null && theme.themeElements.clrScheme.colors[clrMap[this.id]] != null && theme.themeElements.clrScheme.colors[clrMap[this.id]].color != null)
						this.RGBA = theme.themeElements.clrScheme.colors[clrMap[this.id]].color.RGBA;
					else if (theme.themeElements.clrScheme.colors[this.id] != null && theme.themeElements.clrScheme.colors[this.id].color != null)
						this.RGBA = theme.themeElements.clrScheme.colors[this.id].color.RGBA;
				}
			}
		};

		function CStyleColor() {
			CBaseColor.call(this);
			this.bAuto = false;
			this.val = null;
		}

		InitClass(CStyleColor, CBaseColor, 0);
		CStyleColor.prototype.type = c_oAscColor.COLOR_TYPE_STYLE;
		CStyleColor.prototype.check = function (theme, colorMap) {
		};
		CStyleColor.prototype.Write_ToBinary = function (w) {
			w.WriteLong(this.type);
			writeBool(w, this.bAuto);
			writeLong(w, this.val);
		};
		CStyleColor.prototype.Read_FromBinary = function (r) {
			this.bAuto = readBool(r);
			this.val = readLong(r);
		};
		CStyleColor.prototype.IsIdentical = function (color) {
			return color && color.type === this.type && color.bAuto === this.bAuto && this.val === color.val;
		};
		CStyleColor.prototype.createDuplicate = function () {
			var duplicate = new CStyleColor();
			duplicate.bAuto = this.bAuto;
			duplicate.val = this.val;
			return duplicate;
		};
		CStyleColor.prototype.Calculate = function (theme, slide, layout, masterSlide, RGBA, colorMap) {
		};
		CStyleColor.prototype.getNoStyleUnicolor = function (nIdx, aColors) {
			if (this.bAuto || this.val === null) {
				return aColors[nIdx % aColors.length];
			} else {
				return aColors[this.val % aColors.length];
			}
		};

		/**
		 * @constructor
		 */
		function CUniColor() {
			CBaseNoIdObject.call(this);
			this.color = null;
			this.Mods = null;//new CColorModifiers();
			this.RGBA = {
				R: 0,
				G: 0,
				B: 0,
				A: 255
			};
		}

		InitClass(CUniColor, CBaseNoIdObject, 0);
		CUniColor.prototype.checkPhColor = function (unicolor) {
			if (this.color && this.color.type === c_oAscColor.COLOR_TYPE_SCHEME && this.color.id === 14) {
				if (unicolor) {
					if (unicolor.color) {
						this.color = unicolor.color.createDuplicate();
					}
					if (unicolor.Mods) {
						if (!this.Mods || this.Mods.Mods.length === 0) {
							this.Mods = unicolor.Mods.createDuplicate();
						} else {
							this.Mods.Merge(unicolor.Mods);
						}
					}
				}
			}
		};
		CUniColor.prototype.saveSourceFormatting = function () {
			var _ret = new CUniColor();
			_ret.color = new CRGBColor();
			_ret.color.RGBA.R = this.RGBA.R;
			_ret.color.RGBA.G = this.RGBA.G;
			_ret.color.RGBA.B = this.RGBA.B;
			return _ret;
		};
		CUniColor.prototype.addColorMod = function (mod) {
			if (!this.Mods) {
				this.Mods = new CColorModifiers();
			}
			this.Mods.addMod(mod.createDuplicate());
		};
		CUniColor.prototype.check = function (theme, colorMap) {
			if (this.color && this.color.check(theme, colorMap.color_map)/*возвращает был ли изменен RGBA*/) {
				this.RGBA.R = this.color.RGBA.R;
				this.RGBA.G = this.color.RGBA.G;
				this.RGBA.B = this.color.RGBA.B;
				if (this.Mods)
					this.Mods.Apply(this.RGBA);
			}
		};
		CUniColor.prototype.getModValue = function (sModName) {
			if (this.Mods && this.Mods.getModValue) {
				return this.Mods.getModValue(sModName);
			}
			return null;
		};
		CUniColor.prototype.getMod = function (sModName) {
			if (this.Mods && this.Mods.getMod) {
				return this.Mods.getMod(sModName);
			}
			return null;
		};
		CUniColor.prototype.getTransparency = function () {
			let nAlphaVal = this.getModValue("alpha");
			if(nAlphaVal === null) {
				return 0;
			}
			return (100000 - nAlphaVal) / 1000;
		};
		CUniColor.prototype.checkWordMods = function () {
			return this.Mods && this.Mods.Mods.length === 1
				&& (this.Mods.Mods[0].name === "wordTint" || this.Mods.Mods[0].name === "wordShade");
		};
		CUniColor.prototype.convertToPPTXMods = function () {
			if (this.checkWordMods()) {
				var val_, mod_;
				if (this.Mods.Mods[0].name === "wordShade") {
					mod_ = new CColorMod("lumMod", ((this.Mods.Mods[0].val / 255) * 100000) >> 0);
					this.Mods.Mods.splice(0, this.Mods.Mods.length);
					this.Mods.Mods.push(mod_);
				} else {
					val_ = ((this.Mods.Mods[0].val / 255) * 100000) >> 0;
					this.Mods.Mods.splice(0, this.Mods.Mods.length);
					this.Mods.addMod("lumMod", val_);
					this.Mods.addMod("lumOff", 100000 - val_);
				}
			}
		};
		CUniColor.prototype.canConvertPPTXModsToWord = function () {
			return this.Mods && this.Mods.canGetShadeOrTint();
		};
		CUniColor.prototype.convertToWordMods = function () {
			if (this.canConvertPPTXModsToWord()) {
				var mod_ = new CColorMod();
				mod_.setName(this.Mods.Mods.length === 1 ? "wordShade" : "wordTint");
				mod_.setVal(((this.Mods.Mods[0].val * 255) / 100000) >> 0);
				this.Mods.Mods.splice(0, this.Mods.Mods.length);
				this.Mods.Mods.push(mod_);
			}
		};
		CUniColor.prototype.setColor = function (color) {
			this.color = color;
		};
		CUniColor.prototype.setMods = function (mods) {
			this.Mods = mods;
		};
		CUniColor.prototype.Write_ToBinary = function (w) {
			if (this.color) {
				w.WriteBool(true);
				this.color.Write_ToBinary(w);
			} else {
				w.WriteBool(false);
			}
			if (this.Mods) {
				w.WriteBool(true);
				this.Mods.Write_ToBinary(w);
			} else {
				w.WriteBool(false);
			}
			if(w.bWriteCompiledColors) {
				w.WriteLong(((this.RGBA.R << 16) & 0xFF0000) + ((this.RGBA.G << 8) & 0xFF00) + this.RGBA.B);
			}
		};
		CUniColor.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				var type = r.GetLong();
				switch (type) {
					case c_oAscColor.COLOR_TYPE_NONE: {
						break;
					}
					case c_oAscColor.COLOR_TYPE_SRGB: {
						this.color = new CRGBColor();
						this.color.Read_FromBinary(r);
						break;
					}
					case c_oAscColor.COLOR_TYPE_PRST: {
						this.color = new CPrstColor();
						this.color.Read_FromBinary(r);
						break;
					}
					case c_oAscColor.COLOR_TYPE_SCHEME: {
						this.color = new CSchemeColor();
						this.color.Read_FromBinary(r);
						break;
					}
					case c_oAscColor.COLOR_TYPE_SYS: {
						this.color = new CSysColor();
						this.color.Read_FromBinary(r);
						break;
					}
					case c_oAscColor.COLOR_TYPE_STYLE: {
						this.color = new CStyleColor();
						this.color.Read_FromBinary(r);
						break;
					}
				}
			}
			if (r.GetBool()) {
				this.Mods = new CColorModifiers();
				this.Mods.Read_FromBinary(r);
			} else {
				this.Mods = null;
			}
			if(r.bReadCompiledColors) {
				let RGB = r.GetLong();
				this.RGBA.R = (RGB >> 16) & 0xFF;
				this.RGBA.G = (RGB >> 8) & 0xFF;
				this.RGBA.B = RGB & 0xFF;
			}
		};
		CUniColor.prototype.createDuplicate = function () {
			var duplicate = new CUniColor();
			if (this.color != null) {
				duplicate.color = this.color.createDuplicate();
			}
			if (this.Mods)
				duplicate.Mods = this.Mods.createDuplicate();

			duplicate.RGBA.R = this.RGBA.R;
			duplicate.RGBA.G = this.RGBA.G;
			duplicate.RGBA.B = this.RGBA.B;
			duplicate.RGBA.A = this.RGBA.A;
			return duplicate;
		};
		CUniColor.prototype.IsIdentical = function (unicolor) {
			if (!isRealObject(unicolor)) {
				return false;
			}
			if (!isRealObject(unicolor.color) && isRealObject(this.color)
				|| !isRealObject(this.color) && isRealObject(unicolor.color)
				|| isRealObject(this.color) && !this.color.IsIdentical(unicolor.color)) {
				return false;
			}
			if (!isRealObject(unicolor.Mods) && isRealObject(this.Mods) && this.Mods.Mods.length > 0
				|| !isRealObject(this.Mods) && isRealObject(unicolor.Mods) && unicolor.Mods.Mods.length > 0
				|| isRealObject(this.Mods) && !this.Mods.IsIdentical(unicolor.Mods)) {
				return false;
			}
			return true;
		};
		CUniColor.prototype.isEqual = CUniColor.prototype.IsIdentical;
		CUniColor.prototype.Calculate = function (theme, slide, layout, masterSlide, RGBA, colorMap) {
			if (this.color == null)
				return this.RGBA;

			this.color.Calculate(theme, slide, layout, masterSlide, RGBA, colorMap);

			this.RGBA = {R: this.color.RGBA.R, G: this.color.RGBA.G, B: this.color.RGBA.B, A: this.color.RGBA.A};
			if (this.Mods)
				this.Mods.Apply(this.RGBA);
		};
		CUniColor.prototype.compare = function (unicolor) {
			if (unicolor == null) {
				return null;
			}
			var _ret = new CUniColor();
			if (this.color == null || unicolor.color == null ||
				this.color.type !== unicolor.color.type) {
				return _ret;
			}

			if (this.Mods && unicolor.Mods) {
				var aMods = this.Mods.Mods;
				var aMods2 = unicolor.Mods.Mods;
				if (aMods.length === aMods2.length) {
					for (var i = 0; i < aMods.length; ++i) {
						if (aMods2[i].name !== aMods[i].name || aMods2[i].val !== aMods[i].val) {
							break;
						}
					}
					if (i === aMods.length) {
						_ret.Mods = this.Mods.createDuplicate();
					}
				}
			}
			switch (this.color.type) {
				case c_oAscColor.COLOR_TYPE_NONE: {
					break;
				}
				case c_oAscColor.COLOR_TYPE_PRST: {
					_ret.color = new CPrstColor();
					if (unicolor.color.id == this.color.id) {
						_ret.color.id = this.color.id;
						_ret.color.RGBA.R = this.color.RGBA.R;
						_ret.color.RGBA.G = this.color.RGBA.G;
						_ret.color.RGBA.B = this.color.RGBA.B;
						_ret.color.RGBA.A = this.color.RGBA.A;
						_ret.RGBA.R = this.RGBA.R;
						_ret.RGBA.G = this.RGBA.G;
						_ret.RGBA.B = this.RGBA.B;
						_ret.RGBA.A = this.RGBA.A;
					}
					break;
				}
				case c_oAscColor.COLOR_TYPE_SCHEME: {
					_ret.color = new CSchemeColor();
					if (unicolor.color.id == this.color.id) {
						_ret.color.id = this.color.id;
						_ret.color.RGBA.R = this.color.RGBA.R;
						_ret.color.RGBA.G = this.color.RGBA.G;
						_ret.color.RGBA.B = this.color.RGBA.B;
						_ret.color.RGBA.A = this.color.RGBA.A;
						_ret.RGBA.R = this.RGBA.R;
						_ret.RGBA.G = this.RGBA.G;
						_ret.RGBA.B = this.RGBA.B;
						_ret.RGBA.A = this.RGBA.A;
					}
					break;
				}
				case c_oAscColor.COLOR_TYPE_SRGB: {
					_ret.color = new CRGBColor();
					var _RGBA1 = this.color.RGBA;
					var _RGBA2 = this.color.RGBA;
					if (_RGBA1.R === _RGBA2.R
						&& _RGBA1.G === _RGBA2.G
						&& _RGBA1.B === _RGBA2.B) {
						_ret.color.RGBA.R = this.color.RGBA.R;
						_ret.color.RGBA.G = this.color.RGBA.G;
						_ret.color.RGBA.B = this.color.RGBA.B;
						_ret.RGBA.R = this.RGBA.R;
						_ret.RGBA.G = this.RGBA.G;
						_ret.RGBA.B = this.RGBA.B;
					}
					if (_RGBA1.A === _RGBA2.A) {
						_ret.color.RGBA.A = this.color.RGBA.A;
					}

					break;
				}
				case c_oAscColor.COLOR_TYPE_SYS: {
					_ret.color = new CSysColor();
					if (unicolor.color.id == this.color.id) {
						_ret.color.id = this.color.id;
						_ret.color.RGBA.R = this.color.RGBA.R;
						_ret.color.RGBA.G = this.color.RGBA.G;
						_ret.color.RGBA.B = this.color.RGBA.B;
						_ret.color.RGBA.A = this.color.RGBA.A;
						_ret.RGBA.R = this.RGBA.R;
						_ret.RGBA.G = this.RGBA.G;
						_ret.RGBA.B = this.RGBA.B;
						_ret.RGBA.A = this.RGBA.A;
					}
					break;
				}
			}
			return _ret;
		};
		CUniColor.prototype.getCSSColor = function (transparent) {
			if (transparent != null) {
				return this.getCSSWithTransparent(1);
			}
			return this.getCSSWithTransparent(this.RGBA.A / 255);
		};
		CUniColor.prototype.getCSSValue = function (r, g, b, a) {
			return "rgba(" + r + "," + g + "," + b + ","+ a +")";
		};
		CUniColor.prototype.getCSSWithTransparent = function(dTransparent) {
			const oC = this.RGBA;
			return this.getCSSValue(oC.R, oC.G, oC.B, dTransparent);
		};
		CUniColor.prototype.isCorrect = function () {
			if (this.color !== null && this.color !== undefined) {
				return true;
			}
			return false;
		};
		CUniColor.prototype.getNoStyleUnicolor = function (nIdx, aColors) {
			if (!this.color) {
				return null;
			}
			if (this.color.type !== c_oAscColor.COLOR_TYPE_STYLE) {
				return this;
			}
			return this.color.getNoStyleUnicolor(nIdx, aColors);
		};
		CUniColor.prototype.UNICOLOR_MAP = {
			"hslClr": true,
			"scrgbClr": true,
			"srgbClr": true,
			"prstClr": true,
			"schemeClr": true,
			"sysClr": true
		};
		CUniColor.prototype.isUnicolor = function (sName) {
			return !!CUniColor.prototype.UNICOLOR_MAP[sName];
		};

		function CreateUniColorRGB(r, g, b) {
			var ret = new CUniColor();
			ret.setColor(new CRGBColor());
			ret.color.setColor(r, g, b);
			return ret;
		}

		function CreateUniColorRGB2(color) {
			var ret = new CUniColor();
			ret.setColor(new CRGBColor());
			ret.color.setColor(ret.RGBA.R = color.getR(), ret.RGBA.G = color.getG(), ret.RGBA.B = color.getB());
			return ret;
		}

		function CreateSolidFillRGB(r, g, b) {
			return AscFormat.CreateUniFillByUniColor(CreateUniColorRGB(r, g, b));
		}

		function CreateSolidFillRGBA(r, g, b, a) {
			var ret = new CUniFill();
			ret.setFill(new CSolidFill());
			ret.fill.setColor(new CUniColor());
			var _uni_color = ret.fill.color;
			_uni_color.setColor(new CRGBColor());
			_uni_color.color.setColor(r, g, b);
			_uni_color.RGBA.R = r;
			_uni_color.RGBA.G = g;
			_uni_color.RGBA.B = b;
			_uni_color.RGBA.A = a;
			return ret;
		}

// -----------------------------

		function CSrcRect() {
			CBaseNoIdObject.call(this);
			this.l = null;
			this.t = null;
			this.r = null;
			this.b = null;
		}

		InitClass(CSrcRect, CBaseNoIdObject, 0);
		CSrcRect.prototype.setLTRB = function (l, t, r, b) {
			this.l = l;
			this.t = t;
			this.r = r;
			this.b = b;
		};
		CSrcRect.prototype.setValueForFitBlipFill = function (shapeWidth, shapeHeight, imageWidth, imageHeight) {
			if ((imageHeight / imageWidth) > (shapeHeight / shapeWidth)) {
				this.l = 0;
				this.r = 100;
				var widthAspectRatio = imageWidth / shapeWidth;
				var heightAspectRatio = shapeHeight / imageHeight;
				var stretchPercentage = ((1 - widthAspectRatio * heightAspectRatio) / 2) * 100;
				this.t = stretchPercentage;
				this.b = 100 - stretchPercentage;
			} else {
				this.t = 0;
				this.b = 100;
				heightAspectRatio = imageHeight / shapeHeight;
				widthAspectRatio = shapeWidth / imageWidth;
				stretchPercentage = ((1 - heightAspectRatio * widthAspectRatio) / 2) * 100;
				this.l = stretchPercentage;
				this.r = 100 - stretchPercentage;
			}
		};
		CSrcRect.prototype.Write_ToBinary = function (w) {
			writeDouble(w, this.l);
			writeDouble(w, this.t);
			writeDouble(w, this.r);
			writeDouble(w, this.b);
		};
		CSrcRect.prototype.Read_FromBinary = function (r) {
			this.l = readDouble(r);
			this.t = readDouble(r);
			this.r = readDouble(r);
			this.b = readDouble(r);
		};
		CSrcRect.prototype.createDublicate = function () {
			var _ret = new CSrcRect();
			_ret.l = this.l;
			_ret.t = this.t;
			_ret.r = this.r;
			_ret.b = this.b;
			return _ret;
		};
		CSrcRect.prototype.isFullRect = function() {
			let r = this;
			let fAE = AscFormat.fApproxEqual;
			if(fAE(r.l, 0) && fAE(r.t, 0) && fAE(r.r, 100) && fAE(r.b, 100)) {
				return true;
			}
			return false;
		};
		CSrcRect.prototype.isEqual = function(r) {
			if(!r) {
				return false;
			}
			if(r.l !== this.l) {
				return false;
			}
			if(r.t !== this.t) {
				return false;
			}
			if(r.r !== this.r) {
				return false;
			}
			if(r.b !== this.v) {
				return false;
			}
			return true;
		};

		function CBlipFillTile() {
			CBaseNoIdObject.call(this)
			this.tx = null;
			this.ty = null;
			this.sx = null;
			this.sy = null;
			this.flip = null;
			this.algn = null;
		}

		InitClass(CBlipFillTile, CBaseNoIdObject, 0);
		CBlipFillTile.prototype.Write_ToBinary = function (w) {
			writeLong(w, this.tx);
			writeLong(w, this.ty);
			writeLong(w, this.sx);
			writeLong(w, this.sy);
			writeLong(w, this.flip);
			writeLong(w, this.algn);
		};
		CBlipFillTile.prototype.Read_FromBinary = function (r) {
			this.tx = readLong(r);
			this.ty = readLong(r);
			this.sx = readLong(r);
			this.sy = readLong(r);
			this.flip = readLong(r);
			this.algn = readLong(r);
		};
		CBlipFillTile.prototype.createDuplicate = function () {
			var ret = new CBlipFillTile();
			ret.tx = this.tx;
			ret.ty = this.ty;
			ret.sx = this.sx;
			ret.sy = this.sy;
			ret.flip = this.flip;
			ret.algn = this.algn;
			return ret;
		};
		CBlipFillTile.prototype.IsIdentical = function (o) {
			if (!o) {
				return false;
			}
			return (o.tx == this.tx &&
				o.ty == this.ty &&
				o.sx == this.sx &&
				o.sy == this.sy &&
				o.flip == this.flip &&
				o.algn == this.algn)
		};

		function CBaseFill() {
			CBaseNoIdObject.call(this);
		}

		InitClass(CBaseFill, CBaseNoIdObject, 0);
		CBaseFill.prototype.type = c_oAscFill.FILL_TYPE_NONE;




		function CBlip(oBlipFill) {
			CBaseNoIdObject.call(this);
			this.blipFill = oBlipFill;
			this.link = null;
		}

		InitClass(CBlip, CBaseNoIdObject, 0);


		function CBlipFill() {
			CBaseFill.call(this);
			this.RasterImageId = "";
			this.srcRect = null;
			this.stretch = null;
			this.tile = null;
			this.rotWithShape = null;
			this.Effects = [];
		}

		InitClass(CBlipFill, CBaseFill, 0);
		CBlipFill.prototype.type = c_oAscFill.FILL_TYPE_BLIP;
		CBlipFill.prototype.saveSourceFormatting = function () {
			return this.createDuplicate();
		};
		CBlipFill.prototype.Write_ToBinary = function (w) {
			writeString(w, this.RasterImageId);
			if (this.srcRect) {
				writeBool(w, true);
				writeDouble(w, this.srcRect.l);
				writeDouble(w, this.srcRect.t);
				writeDouble(w, this.srcRect.r);
				writeDouble(w, this.srcRect.b);
			} else {
				writeBool(w, false);
			}
			writeBool(w, this.stretch);
			if (isRealObject(this.tile)) {
				w.WriteBool(true);
				this.tile.Write_ToBinary(w);
			} else {
				w.WriteBool(false);
			}
			writeBool(w, this.rotWithShape);

			w.WriteLong(this.Effects.length);
			for (var i = 0; i < this.Effects.length; ++i) {
				this.Effects[i].Write_ToBinary(w);
			}
		};
		CBlipFill.prototype.Read_FromBinary = function (r) {
			this.RasterImageId = readString(r);

			var _correct_id = AscCommon.getImageFromChanges(this.RasterImageId);
			if (null != _correct_id)
				this.RasterImageId = _correct_id;

			//var srcUrl = readString(r);
			//if(srcUrl) {
			//    AscCommon.g_oDocumentUrls.addImageUrl(this.RasterImageId, srcUrl);
			//}

			if (readBool(r)) {

				this.srcRect = new CSrcRect();
				this.srcRect.l = readDouble(r);
				this.srcRect.t = readDouble(r);
				this.srcRect.r = readDouble(r);
				this.srcRect.b = readDouble(r);
			} else {
				this.srcRect = null;
			}
			this.stretch = readBool(r);
			if (r.GetBool()) {
				this.tile = new CBlipFillTile();
				this.tile.Read_FromBinary(r);
			} else {
				this.tile = null;
			}
			this.rotWithShape = readBool(r);
			var count = r.GetLong();
			for (var i = 0; i < count; ++i) {
				var effect = fReadEffect(r);
				if (!effect) {
					break;
				}
				this.Effects.push(effect);
			}
		};
		CBlipFill.prototype.Refresh_RecalcData = function () {
		};
		CBlipFill.prototype.check = function () {
		};
		CBlipFill.prototype.checkWordMods = function () {
			return false;
		};
		CBlipFill.prototype.convertToPPTXMods = function () {
		};
		CBlipFill.prototype.canConvertPPTXModsToWord = function () {
			return false;
		};
		CBlipFill.prototype.convertToWordMods = function () {
		};
		CBlipFill.prototype.getObjectType = function () {
			return AscDFH.historyitem_type_BlipFill;
		};
		CBlipFill.prototype.setRasterImageId = function (rasterImageId) {
			this.RasterImageId = checkRasterImageId(rasterImageId);
		};
		CBlipFill.prototype.setSrcRect = function (srcRect) {
			this.srcRect = srcRect;
		};
		CBlipFill.prototype.setStretch = function (stretch) {
			this.stretch = stretch;
		};
		CBlipFill.prototype.setTile = function (tile) {
			this.tile = tile;
		};
		CBlipFill.prototype.setRotWithShape = function (rotWithShape) {
			this.rotWithShape = rotWithShape;
		};
		CBlipFill.prototype.createDuplicate = function () {
			var duplicate = new CBlipFill();
			duplicate.RasterImageId = this.RasterImageId;

			duplicate.stretch = this.stretch;
			if (isRealObject(this.tile)) {
				duplicate.tile = this.tile.createDuplicate();
			}

			if (null != this.srcRect)
				duplicate.srcRect = this.srcRect.createDublicate();

			duplicate.rotWithShape = this.rotWithShape;

			if (Array.isArray(this.Effects)) {
				for (var i = 0; i < this.Effects.length; ++i) {
					if (this.Effects[i] && this.Effects[i].createDuplicate) {
						duplicate.Effects.push(this.Effects[i].createDuplicate());
					}
				}
			}
			return duplicate;
		};
		CBlipFill.prototype.IsIdentical = function (fill) {
			if (fill == null) {
				return false;
			}
			if (fill.type !== c_oAscFill.FILL_TYPE_BLIP) {
				return false;
			}

			if (fill.RasterImageId !== this.RasterImageId) {
				return false;
			}

			/*if(fill.VectorImageBin !=  this.VectorImageBin)
         {
         return false;
         }    */

			if (fill.stretch != this.stretch) {
				return false;
			}

			if (isRealObject(this.tile)) {
				if (!this.tile.IsIdentical(fill.tile)) {
					return false;
				}
			} else {
				if (fill.tile) {
					return false;
				}
			}
			if(fill.srcRect && !this.srcRect || this.srcRect && !fill.srcRect) {
				return false;
			}
			if(fill.srcRect) {
				if(!fill.srcRect.isEqual(this.srcRect)) {
					return false;
				}
			}
			if (fill.Effects.length !== this.Effects.length) {
				return false;
			}
			for (let i = 0; i < fill.Effects.length; i += 1) {
				if (!fill.Effects[i].IsIdentical(this.Effects[i])) {
					return false;
				}
			}

			/*
         if(fill.rotWithShape !=  this.rotWithShape)
         {
         return false;
         }
         */
			return true;

		};
		CBlipFill.prototype.compare = function (fill) {
			if (fill == null || fill.type !== c_oAscFill.FILL_TYPE_BLIP) {
				return null;
			}
			var _ret = new CBlipFill();
			if (this.RasterImageId == fill.RasterImageId) {
				_ret.RasterImageId = this.RasterImageId;
			}
			if (fill.stretch == this.stretch) {
				_ret.stretch = this.stretch;
			}
			if (isRealObject(fill.tile)) {
				if (fill.tile.IsIdentical(this.tile)) {
					_ret.tile = this.tile.createDuplicate();
				} else {
					_ret.tile = new CBlipFillTile();
				}
			}
			if (fill.rotWithShape === this.rotWithShape) {
				_ret.rotWithShape = this.rotWithShape;
			}
			return _ret;
		};
		CBlipFill.prototype.reduceSize = function (nW, nH, nMaxSize) {
			let dWK = nW / nMaxSize;
			let dHK = nH / nMaxSize;
			let dK = Math.max(dWK, dHK);
			let oResult = {W: nW, H: nH};
			if (dK > 1) {
				oResult.W = ((nW / dK) + 0.5 >> 0);
				oResult.H = ((nH / dK) + 0.5 >> 0);
			}
			return oResult;
		};
		CBlipFill.prototype.getBase64Data = function (bReduce, bReturnOrigIfCantDraw) {
			let sRasterImageId = this.RasterImageId;
			if (typeof sRasterImageId !== "string" || sRasterImageId.length === 0) {
				return null;
			}
			let oApi = Asc.editor || editor;
			let sDefaultResult = sRasterImageId;
			const nMaxSize = 640;
			if(bReturnOrigIfCantDraw === false) {
				sDefaultResult = null;
			}
			if (window["NATIVE_EDITOR_ENJINE"] && window["native"]) {
				let sSrc = AscCommon.getFullImageSrc2(sRasterImageId);
				let oImageSize = AscCommon.getSourceImageSize(sSrc);
				let nW = Math.max(oImageSize.width, 1);
				let nH = Math.max(oImageSize.height, 1);
				if (bReduce) {
					let oReducedSize = this.reduceSize(nW, nH, nMaxSize);
					nW = oReducedSize.W;
					nH = oReducedSize.H;
				}
				let oCanvas = new AscCommon.CNativeGraphics();
				oCanvas.width = nW;
				oCanvas.height = nH;
				oCanvas.create(window["native"], oCanvas.width, oCanvas.height, oCanvas.width, oCanvas.height);
				oCanvas.transform(1, 0, 0, 1, 0, 0);
				oCanvas.drawImage(sSrc, 0, 0, oCanvas.width, oCanvas.height);
				let sResult;
				try {
					sResult = oCanvas.toDataURL("image/png");
				} catch (err) {
					sResult = sDefaultResult;
				}
				return {img: sResult, w: oCanvas.width, h: oCanvas.height};
			}
			if (!oApi) {
				return {img: sDefaultResult, w: null, h: null};
			}
			let oImageLoader = oApi.ImageLoader;
			if (!oImageLoader) {
				return {img: sDefaultResult, w: null, h: null};
			}
			let oImage = oImageLoader.map_image_index[AscCommon.getFullImageSrc2(sRasterImageId)];
			if (!oImage || !oImage.Image || oImage.Status !== AscFonts.ImageLoadStatus.Complete) {
				return {img: sDefaultResult, w: null, h: null};
			}

			if (sRasterImageId.indexOf("data:") === 0 && sRasterImageId.indexOf("base64") > 0) {
				return {img: sRasterImageId, w: oImage.Image.width, h: oImage.Image.height};
			}
			let sResult = sDefaultResult;
			if (!window["NATIVE_EDITOR_ENJINE"]) {
				let oCanvas = document.createElement("canvas");
				let nW = Math.max(oImage.Image.width, 1);
				let nH = Math.max(oImage.Image.height, 1);
				if (bReduce) {
					let oReducedSize = this.reduceSize(nW, nH, nMaxSize);
					nW = oReducedSize.W;
					nH = oReducedSize.H;
				}
				oCanvas.width = nW;
				oCanvas.height = nH;
				var oCtx = oCanvas.getContext("2d");
				oCtx.drawImage(oImage.Image, 0, 0, oCanvas.width, oCanvas.height);
				try {
					sResult = oCanvas.toDataURL("image/png");
				} catch (err) {
					sResult = sDefaultResult;
				}
				return {img: sResult, w: oCanvas.width, h: oCanvas.height};
			}
			return {img: sRasterImageId, w: null, h: null};

		};
		CBlipFill.prototype.getBase64RasterImageId = function (bReduce, bReturnOrigIfCantDraw) {
			return this.getBase64Data(bReduce, bReturnOrigIfCantDraw).img;
		};
		CBlipFill.prototype.getTransparent = function() {
			for (let nEffect = 0; nEffect < this.Effects.length; ++nEffect) {
				let oEffect = this.Effects[nEffect];
				if (oEffect && oEffect instanceof AscFormat.CAlphaModFix && AscFormat.isRealNumber(oEffect.amt)) {
					return 255 * oEffect.amt / 100000;
				}
			}
			return null;
		};
		CBlipFill.prototype.setTransparent = function(transparent) {
			if (transparent != null) {
				let i;
				for (i = 0; i < this.Effects.length; ++i) {
					if (this.Effects[i].Type === EFFECT_TYPE_ALPHAMODFIX) {
						this.Effects[i].amt = ((transparent * 100000 / 255) >> 0);
						break;
					}
				}
				if (i === this.Effects.length) {
					let oEffect = new CAlphaModFix();
					oEffect.amt = ((transparent * 100000 / 255) >> 0);
					this.Effects.push(oEffect);
				}
			}
		};
		CBlipFill.prototype.createDuplicateNoRaster = function(transparent) {
			let sId = this.RasterImageId;
			this.RasterImageId = null;
			let copy = this.createDuplicate();
			this.RasterImageId = sId;
			return copy;
		};

//-----Effects-----
		var EFFECT_TYPE_NONE = 0;
		var EFFECT_TYPE_OUTERSHDW = 1;
		var EFFECT_TYPE_GLOW = 2;
		var EFFECT_TYPE_DUOTONE = 3;
		var EFFECT_TYPE_XFRM = 4;
		var EFFECT_TYPE_BLUR = 5;
		var EFFECT_TYPE_PRSTSHDW = 6;
		var EFFECT_TYPE_INNERSHDW = 7;
		var EFFECT_TYPE_REFLECTION = 8;
		var EFFECT_TYPE_SOFTEDGE = 9;
		var EFFECT_TYPE_FILLOVERLAY = 10;
		var EFFECT_TYPE_ALPHACEILING = 11;
		var EFFECT_TYPE_ALPHAFLOOR = 12;
		var EFFECT_TYPE_TINTEFFECT = 13;
		var EFFECT_TYPE_RELOFF = 14;
		var EFFECT_TYPE_LUM = 15;
		var EFFECT_TYPE_HSL = 16;
		var EFFECT_TYPE_GRAYSCL = 17;
		var EFFECT_TYPE_ELEMENT = 18;
		var EFFECT_TYPE_ALPHAREPL = 19;
		var EFFECT_TYPE_ALPHAOUTSET = 20;
		var EFFECT_TYPE_ALPHAMODFIX = 21;
		var EFFECT_TYPE_ALPHABILEVEL = 22;
		var EFFECT_TYPE_BILEVEL = 23;
		var EFFECT_TYPE_DAG = 24;
		var EFFECT_TYPE_FILL = 25;
		var EFFECT_TYPE_CLRREPL = 26;
		var EFFECT_TYPE_CLRCHANGE = 27;
		var EFFECT_TYPE_ALPHAINV = 28;
		var EFFECT_TYPE_ALPHAMOD = 29;
		var EFFECT_TYPE_BLEND = 30;

		function fCreateEffectByType(type) {
			var ret = null;
			switch (type) {
				case EFFECT_TYPE_NONE: {
					break;
				}
				case EFFECT_TYPE_OUTERSHDW: {
					ret = new COuterShdw();
					break;
				}
				case EFFECT_TYPE_GLOW: {
					ret = new CGlow();
					break;
				}
				case EFFECT_TYPE_DUOTONE: {
					ret = new CDuotone();
					break;
				}
				case EFFECT_TYPE_XFRM: {
					ret = new CXfrmEffect();
					break;
				}
				case EFFECT_TYPE_BLUR: {
					ret = new CBlur();
					break;
				}
				case EFFECT_TYPE_PRSTSHDW: {
					ret = new CPrstShdw();
					break;
				}
				case EFFECT_TYPE_INNERSHDW: {
					ret = new CInnerShdw();
					break;
				}
				case EFFECT_TYPE_REFLECTION: {
					ret = new CReflection();
					break;
				}
				case EFFECT_TYPE_SOFTEDGE: {
					ret = new CSoftEdge();
					break;
				}
				case EFFECT_TYPE_FILLOVERLAY: {
					ret = new CFillOverlay();
					break;
				}
				case EFFECT_TYPE_ALPHACEILING: {
					ret = new CAlphaCeiling();
					break;
				}
				case EFFECT_TYPE_ALPHAFLOOR: {
					ret = new CAlphaFloor();
					break;
				}
				case EFFECT_TYPE_TINTEFFECT: {
					ret = new CTintEffect();
					break;
				}
				case EFFECT_TYPE_RELOFF: {
					ret = new CRelOff();
					break;
				}
				case EFFECT_TYPE_LUM: {
					ret = new CLumEffect();
					break;
				}
				case EFFECT_TYPE_HSL: {
					ret = new CHslEffect();
					break;
				}
				case EFFECT_TYPE_GRAYSCL: {
					ret = new CGrayscl();
					break;
				}
				case EFFECT_TYPE_ELEMENT: {
					ret = new CEffectElement();
					break;
				}
				case EFFECT_TYPE_ALPHAREPL: {
					ret = new CAlphaRepl();
					break;
				}
				case EFFECT_TYPE_ALPHAOUTSET: {
					ret = new CAlphaOutset();
					break;
				}
				case EFFECT_TYPE_ALPHAMODFIX: {
					ret = new CAlphaModFix();
					break;
				}
				case EFFECT_TYPE_ALPHABILEVEL: {
					ret = new CAlphaBiLevel();
					break;
				}
				case EFFECT_TYPE_BILEVEL: {
					ret = new CBiLevel();
					break;
				}
				case EFFECT_TYPE_DAG: {
					ret = new CEffectContainer();
					break;
				}
				case EFFECT_TYPE_FILL: {
					ret = new CFillEffect();
					break;
				}
				case EFFECT_TYPE_CLRREPL: {
					ret = new CClrRepl();
					break;
				}
				case EFFECT_TYPE_CLRCHANGE: {
					ret = new CClrChange();
					break;
				}
				case EFFECT_TYPE_ALPHAINV: {
					ret = new CAlphaInv();
					break;
				}
				case EFFECT_TYPE_ALPHAMOD: {
					ret = new CAlphaMod();
					break;
				}
				case EFFECT_TYPE_BLEND: {
					ret = new CBlend();
					break;
				}
			}
			return ret;
		}

		function fReadEffect(r) {
			var type = r.GetLong();
			var ret = fCreateEffectByType(type);
			ret.Read_FromBinary(r);
			return ret;
		}

		function CAlphaBiLevel() {
			CBaseNoIdObject.call(this);
			this.tresh = 0;
		}

		InitClass(CAlphaBiLevel, CBaseNoIdObject, 0);

		CAlphaBiLevel.prototype.Type = EFFECT_TYPE_ALPHABILEVEL;
		CAlphaBiLevel.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_ALPHABILEVEL);
			w.WriteLong(this.tresh);
		};
		CAlphaBiLevel.prototype.Read_FromBinary = function (r) {
			this.tresh = r.GetLong();
		};
		CAlphaBiLevel.prototype.createDuplicate = function () {
			var oCopy = new CAlphaBiLevel();
			oCopy.tresh = this.tresh;
			return oCopy;
		};
		CAlphaBiLevel.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (this.tresh !== oEffect.tresh) {
				return false;
			}

			return true;
		};

		function CAlphaCeiling() {
			CBaseNoIdObject.call(this);
		}

		InitClass(CAlphaCeiling, CBaseNoIdObject, 0);
		CAlphaCeiling.prototype.Type = EFFECT_TYPE_ALPHACEILING;
		CAlphaCeiling.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_ALPHACEILING);
		};
		CAlphaCeiling.prototype.Read_FromBinary = function (r) {
		};
		CAlphaCeiling.prototype.createDuplicate = function () {
			var oCopy = new CAlphaCeiling();
			return oCopy;
		};
		CAlphaCeiling.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			return true;
		};
		function CAlphaFloor() {
			CBaseNoIdObject.call(this);
		}

		InitClass(CBaseNoIdObject, CBaseNoIdObject, 0);
		CAlphaFloor.prototype.Type = EFFECT_TYPE_ALPHAFLOOR;
		CAlphaFloor.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_ALPHAFLOOR);
		};
		CAlphaFloor.prototype.Read_FromBinary = function (r) {
		};
		CAlphaFloor.prototype.createDuplicate = function () {
			var oCopy = new CAlphaFloor();
			return oCopy;
		};
		CAlphaFloor.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			return true;
		};
		function CAlphaInv() {
			CBaseNoIdObject.call(this);
			this.unicolor = null;
		}

		InitClass(CAlphaInv, CBaseNoIdObject, 0);
		CAlphaInv.prototype.Type = EFFECT_TYPE_ALPHAINV;
		CAlphaInv.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_ALPHAINV);
			if (this.unicolor) {
				w.WriteBool(true);
				this.unicolor.Write_ToBinary(w);
			} else {
				w.WriteBool(false);
			}
		};
		CAlphaInv.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				this.unicolor = new CUniColor();
				this.unicolor.Read_FromBinary(r);
			}
		};
		CAlphaInv.prototype.createDuplicate = function () {
			var oCopy = new CAlphaInv();
			if (this.unicolor) {
				oCopy.unicolor = this.unicolor.createDuplicate();
			}
			return oCopy;
		};
		CAlphaInv.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (!(this.unicolor && oEffect.unicolor && this.unicolor.IsIdentical(oEffect.unicolor) || !this.unicolor && !oEffect.unicolor)) {
				return false;
			}
			return true;
		};
		var effectcontainertypeSib = 0;
		var effectcontainertypeTree = 1;

		function CEffectContainer() {
			CBaseNoIdObject.call(this);
			this.type = null;
			this.name = null;
			this.effectList = [];
		}

		InitClass(CEffectContainer, CBaseNoIdObject, 0);
		CEffectContainer.prototype.Type = EFFECT_TYPE_DAG;
		CEffectContainer.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_DAG);
			writeLong(w, this.type);
			writeString(w, this.name);
			w.WriteLong(this.effectList.length);
			for (var i = 0; i < this.effectList.length; ++i) {
				this.effectList[i].Write_ToBinary(w);
			}
		};
		CEffectContainer.prototype.Read_FromBinary = function (r) {
			this.type = readLong(r);
			this.name = readString(r);
			var count = r.GetLong();
			for (var i = 0; i < count; ++i) {
				var effect = fReadEffect(r);
				if (!effect) {
					//error
					break;
				}
				this.effectList.push(effect);
			}
		};
		CEffectContainer.prototype.createDuplicate = function () {
			var oCopy = new CEffectContainer();
			oCopy.type = this.type;
			oCopy.name = this.name;
			for (var i = 0; i < this.effectList.length; ++i) {
				oCopy.effectList.push(this.effectList[i].createDuplicate());
			}
			return oCopy;
		};
		CEffectContainer.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (this.type !== oEffect.type) {
				return false;
			}
			if (this.effectList.length !== oEffect.effectList.length) {
				return false;
			}
			for (let i = 0; i < this.effectList.length; i += 1) {
				if (!this.effectList[i].IsIdentical(oEffect.effectList[i])) {
					return false;
				}
			}
			// todo ?
			/*if (this.name !== oEffect.name) {
				return false;
			}*/
			return true;
		};

		function CAlphaMod() {
			CBaseNoIdObject.call(this);
			this.cont = new CEffectContainer();
		}

		InitClass(CAlphaMod, CBaseNoIdObject, 0);
		CAlphaMod.prototype.Type = EFFECT_TYPE_ALPHAMOD;
		CAlphaMod.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_ALPHAMOD);
			this.cont.Write_ToBinary(w);
		};
		CAlphaMod.prototype.Read_FromBinary = function (r) {
			this.cont.Read_FromBinary(r);
		};
		CAlphaMod.prototype.createDuplicate = function () {
			var oCopy = new CAlphaMod();
			oCopy.cont = this.cont.createDuplicate();
			return oCopy;
		};
		CAlphaMod.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (!(this.cont && oEffect.cont && this.cont.IsIdentical(oEffect.cont) || !this.cont && !oEffect.cont)) {
				return false;
			}
			return true;
		};

		function CAlphaModFix() {
			CBaseNoIdObject.call(this);
			this.amt = null;
		}

		InitClass(CAlphaModFix, CBaseNoIdObject, 0);
		CAlphaModFix.prototype.Type = EFFECT_TYPE_ALPHAMODFIX;
		CAlphaModFix.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_ALPHAMODFIX);
			writeLong(w, this.amt);
		};
		CAlphaModFix.prototype.Read_FromBinary = function (r) {
			this.amt = readLong(r);
		};
		CAlphaModFix.prototype.createDuplicate = function () {
			var oCopy = new CAlphaModFix();
			oCopy.amt = this.amt;
			return oCopy;
		};
		CAlphaModFix.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (this.amt !== oEffect.amt) {
				return false;
			}
			return true;
		};

		function CAlphaOutset() {
			CBaseNoIdObject.call(this);
			this.rad = null;
		}

		InitClass(CAlphaOutset, CBaseNoIdObject, 0);
		CAlphaOutset.prototype.Type = EFFECT_TYPE_ALPHAOUTSET;
		CAlphaOutset.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_ALPHAOUTSET);
			writeLong(w, this.rad);
		};
		CAlphaOutset.prototype.Read_FromBinary = function (r) {
			this.rad = readLong(r);
		};
		CAlphaOutset.prototype.createDuplicate = function () {
			var oCopy = new CAlphaOutset();
			oCopy.rad = this.rad;
			return oCopy;
		};
		CAlphaOutset.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (this.rad !== oEffect.rad) {
				return false;
			}

			return true;
		};

		function CAlphaRepl() {
			CBaseNoIdObject.call(this);
			this.a = null;
		}

		InitClass(CAlphaRepl, CBaseNoIdObject, 0);
		CAlphaRepl.prototype.Type = EFFECT_TYPE_ALPHAREPL;
		CAlphaRepl.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_ALPHAREPL);
			writeLong(w, this.a);
		};
		CAlphaRepl.prototype.Read_FromBinary = function (r) {
			this.a = readLong(r);
		};
		CAlphaRepl.prototype.createDuplicate = function () {
			var oCopy = new CAlphaRepl();
			oCopy.a = this.a;
			return oCopy;
		};
		CAlphaRepl.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (this.a !== oEffect.a) {
				return false;
			}

			return true;
		};

		function CBiLevel() {
			CBaseNoIdObject.call(this);
			this.thresh = null;
		}

		InitClass(CBiLevel, CBaseNoIdObject, 0);
		CBiLevel.prototype.Type = EFFECT_TYPE_BILEVEL;
		CBiLevel.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_BILEVEL);
			writeLong(w, this.thresh);
		};
		CBiLevel.prototype.Read_FromBinary = function (r) {
			this.thresh = readLong(r);
		};
		CBiLevel.prototype.createDuplicate = function () {
			var oCopy = new CBiLevel();
			oCopy.thresh = this.thresh;
			return oCopy;
		};
		CBiLevel.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}

			if (this.thresh !== oEffect.thresh) {
				return false;
			}

			return true;
		};

		var blendmodeDarken = 0;
		var blendmodeLighten = 1;
		var blendmodeMult = 2;
		var blendmodeOver = 3;
		var blendmodeScreen = 4;

		function CBlend() {
			CBaseNoIdObject.call(this);
			this.blend = null;
			this.cont = new CEffectContainer();
		}

		InitClass(CBlend, CBaseNoIdObject, 0);
		CBlend.prototype.Type = EFFECT_TYPE_BLEND;
		CBlend.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_BLEND);
			writeLong(w, this.blend);
			this.cont.Write_ToBinary(w);
		};
		CBlend.prototype.Read_FromBinary = function (r) {
			this.blend = readLong(r);
			this.cont.Read_FromBinary(r);
		};
		CBlend.prototype.createDuplicate = function () {
			var oCopy = new CBlend();
			oCopy.blend = this.blend;
			oCopy.cont = this.cont.createDuplicate();
			return oCopy;
		};
		CBlend.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (this.blend !== oEffect.blend) {
				return false;
			}
			if (!(this.cont && oEffect.cont && this.cont.IsIdentical(oEffect.cont) || !this.cont && !oEffect.cont)) {
				return false;
			}
			return true;
		};

		function CBlur() {
			CBaseNoIdObject.call(this);
			this.rad = null;
			this.grow = null;
		}

		InitClass(CBlur, CBaseNoIdObject, 0);
		CBlur.prototype.Type = EFFECT_TYPE_BLUR;
		CBlur.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_BLUR);
			writeLong(w, this.rad);
			writeBool(w, this.grow);
		};
		CBlur.prototype.Read_FromBinary = function (r) {
			this.rad = readLong(r);
			this.grow = readBool(r);
		};
		CBlur.prototype.createDuplicate = function () {
			var oCopy = new CBlur();
			oCopy.rad = this.rad;
			oCopy.grow = this.grow;
			return oCopy;
		};
		CBlur.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (this.rad !== oEffect.rad) {
				return false;
			}
			if (this.grow !== oEffect.grow) {
				return false;
			}

			return true;
		};

		function CClrChange() {
			CBaseNoIdObject.call(this);
			this.clrFrom = new CUniColor();
			this.clrTo = new CUniColor();
			this.useA = null;
		}

		InitClass(CClrChange, CBaseNoIdObject, 0);
		CClrChange.prototype.Type = EFFECT_TYPE_CLRCHANGE;
		CClrChange.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_CLRCHANGE);
			this.clrFrom.Write_ToBinary(w);
			this.clrTo.Write_ToBinary(w);
			writeBool(w, this.useA);

		};
		CClrChange.prototype.Read_FromBinary = function (r) {
			this.clrFrom.Read_FromBinary(r);
			this.clrTo.Read_FromBinary(r);
			this.useA = readBool(r);
		};
		CClrChange.prototype.createDuplicate = function () {
			var oCopy = new CClrChange();
			oCopy.clrFrom = this.clrFrom.createDuplicate();
			oCopy.clrTo = this.clrTo.createDuplicate();
			oCopy.useA = this.useA;
			return oCopy;
		};
		CClrChange.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (!(this.clrFrom && oEffect.clrFrom && this.clrFrom.IsIdentical(oEffect.clrFrom) || !this.clrFrom && !oEffect.clrFrom)) {
				return false;
			}
			if (!(this.clrTo && oEffect.clrTo && this.clrTo.IsIdentical(oEffect.clrTo) || !this.clrTo && !oEffect.clrTo)) {
				return false;
			}
			if (this.useA !== oEffect.useA) {
				return false;
			}
			return true;
		};

		function CClrRepl() {
			CBaseNoIdObject.call(this);
			this.color = new CUniColor();
		}

		InitClass(CClrRepl, CBaseNoIdObject, 0);
		CClrRepl.prototype.Type = EFFECT_TYPE_CLRREPL;
		CClrRepl.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_CLRREPL);
			this.color.Write_ToBinary(w);

		};
		CClrRepl.prototype.Read_FromBinary = function (r) {
			this.color.Read_FromBinary(r);
		};
		CClrRepl.prototype.createDuplicate = function () {
			var oCopy = new CClrRepl();
			oCopy.color = this.color.createDuplicate();
			return oCopy;
		};
		CClrRepl.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (!(this.color && oEffect.color && this.color.IsIdentical(oEffect.color) || !this.color && !oEffect.color)) {
				return false;
			}
			return true;
		};

		function CDuotone() {
			CBaseNoIdObject.call(this);
			this.colors = [];
		}

		InitClass(CDuotone, CBaseNoIdObject, 0);
		CDuotone.prototype.Type = EFFECT_TYPE_DUOTONE;
		CDuotone.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_DUOTONE);
			w.WriteLong(this.colors.length);
			for (var i = 0; i < this.colors.length; ++i) {
				this.colors[i].Write_ToBinary(w);
			}

		};
		CDuotone.prototype.Read_FromBinary = function (r) {
			var count = r.GetLong();
			for (var i = 0; i < count; ++i) {
				this.colors[i] = new CUniColor();
				this.colors[i].Read_FromBinary(r);
			}
		};
		CDuotone.prototype.createDuplicate = function () {
			var oCopy = new CDuotone();
			for (var i = 0; i < this.colors.length; ++i) {
				oCopy.colors[i] = this.colors[i].createDuplicate();
			}
			return oCopy;
		};
		CDuotone.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}

			if (this.colors.length !== oEffect.colors.length) {
				return false;
			}

			for (let i = 0; i < this.colors.length; i += 1) {
				if (!this.colors[i].IsIdentical(oEffect.colors[i])) {
					return false;
				}
			}

			return true;
		};


		function CEffectElement() {
			CBaseNoIdObject.call(this);
			this.ref = null;
		}

		InitClass(CEffectElement, CBaseNoIdObject, 0);
		CEffectElement.prototype.Type = EFFECT_TYPE_ELEMENT;
		CEffectElement.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_ELEMENT);
			writeString(w, this.ref);
		};
		CEffectElement.prototype.Read_FromBinary = function (r) {
			this.ref = readString(r);
		};
		CEffectElement.prototype.createDuplicate = function () {
			var oCopy = new CEffectElement();
			oCopy.ref = this.ref;
			return oCopy;
		};
		CEffectElement.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (this.ref !== oEffect.ref) {
				return false;
			}
			return true;
		};

		function CFillEffect() {
			CBaseNoIdObject.call(this);
			this.fill = new CUniFill();
		}

		InitClass(CFillEffect, CBaseNoIdObject, 0);
		CFillEffect.prototype.Type = EFFECT_TYPE_FILL;
		CFillEffect.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_FILL);
			this.fill.Write_ToBinary(w);
		};
		CFillEffect.prototype.Read_FromBinary = function (r) {
			this.fill.Read_FromBinary(r);
		};
		CFillEffect.prototype.createDuplicate = function () {
			var oCopy = new CFillEffect();
			oCopy.fill = this.fill.createDuplicate();
			return oCopy;
		};
		CFillEffect.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (!(this.fill && oEffect.fill && this.fill.IsIdentical(oEffect.fill) || !this.fill && !oEffect.fill)) {
				return false;
			}

			return true;
		};

		function CFillOverlay() {
			CBaseNoIdObject.call(this);
			this.fill = new CUniFill();
			this.blend = 0;
		}

		InitClass(CFillOverlay, CBaseNoIdObject, 0);
		CFillOverlay.prototype.Type = EFFECT_TYPE_FILLOVERLAY;
		CFillOverlay.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_FILLOVERLAY);
			this.fill.Write_ToBinary(w);
			w.WriteLong(this.blend);
		};
		CFillOverlay.prototype.Read_FromBinary = function (r) {
			this.fill.Read_FromBinary(r);
			this.blend = r.GetLong();
		};
		CFillOverlay.prototype.createDuplicate = function () {
			var oCopy = new CFillOverlay();
			oCopy.fill = this.fill.createDuplicate();
			oCopy.blend = this.blend;
			return oCopy;
		};
		CFillOverlay.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}

			if (!(this.fill && oEffect.fill && this.fill.IsIdentical(oEffect) || !this.fill && !oEffect.fill)) {
				return false;
			}

			if (this.blend !== oEffect.blend) {
				return false;
			}

			return true;
		};

		function CGlow() {
			CBaseNoIdObject.call(this);
			this.color = new CUniColor();
			this.rad = null;
		}

		InitClass(CGlow, CBaseNoIdObject, 0);
		CGlow.prototype.Type = EFFECT_TYPE_GLOW;
		CGlow.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_GLOW);
			this.color.Write_ToBinary(w);
			writeLong(w, this.rad);
		};
		CGlow.prototype.Read_FromBinary = function (r) {
			this.color.Read_FromBinary(r);
			this.rad = readLong(r);
		};
		CGlow.prototype.createDuplicate = function () {
			var oCopy = new CGlow();
			oCopy.color = this.color.createDuplicate();
			oCopy.rad = this.rad;
			return oCopy;
		};
		CGlow.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (!(this.color && oEffect.color && this.color.IsIdentical(oEffect.color) || !this.color && !oEffect.color)) {
				return false;
			}
			if (this.rad !== oEffect.rad) {
				return false;
			}
			return true;
		};

		function CGrayscl() {
			CBaseNoIdObject.call(this);
		}

		InitClass(CGrayscl, CBaseNoIdObject, 0);
		CGrayscl.prototype.Type = EFFECT_TYPE_GRAYSCL;
		CGrayscl.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_GRAYSCL);
		};
		CGrayscl.prototype.Read_FromBinary = function (r) {
		};
		CGrayscl.prototype.createDuplicate = function () {
			var oCopy = new CGrayscl();
			return oCopy;
		};
		CGrayscl.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}

			return true;
		};

		function CHslEffect() {
			CBaseNoIdObject.call(this);
			this.h = null;
			this.s = null;
			this.l = null;
		}

		InitClass(CHslEffect, CBaseNoIdObject, 0);
		CHslEffect.prototype.Type = EFFECT_TYPE_HSL;
		CHslEffect.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_HSL);
			writeLong(w, this.h);
			writeLong(w, this.s);
			writeLong(w, this.l);
		};
		CHslEffect.prototype.Read_FromBinary = function (r) {
			this.h = readLong(r);
			this.s = readLong(r);
			this.l = readLong(r);
		};
		CHslEffect.prototype.createDuplicate = function () {
			var oCopy = new CHslEffect();
			oCopy.h = this.h;
			oCopy.s = this.s;
			oCopy.l = this.l;
			return oCopy;
		};
		CHslEffect.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (this.h !== oEffect.h) {
				return false;
			}
			if (this.s !== oEffect.s) {
				return false;
			}
			if (this.l !== oEffect.l) {
				return false;
			}
			return true;
		};

		function CInnerShdw() {
			CBaseNoIdObject.call(this);
			this.color = new CUniColor();
			this.blurRad = null;
			this.dir = null;
			this.dist = null;
		}

		InitClass(CInnerShdw, CBaseNoIdObject, 0);
		CInnerShdw.prototype.Type = EFFECT_TYPE_INNERSHDW;
		CInnerShdw.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_INNERSHDW);
			this.color.Write_ToBinary(w);
			writeLong(w, this.blurRad);
			writeLong(w, this.dir);
			writeLong(w, this.dist);
		};
		CInnerShdw.prototype.Read_FromBinary = function (r) {
			this.color.Read_FromBinary(r);
			this.blurRad = readLong(r);
			this.dir = readLong(r);
			this.dist = readLong(r);
		};
		CInnerShdw.prototype.createDuplicate = function () {
			var oCopy = new CInnerShdw();
			oCopy.color = this.color.createDuplicate();
			oCopy.blurRad = this.blurRad;
			oCopy.dir = this.dir;
			oCopy.dist = this.dist;
			return oCopy;
		};
		CInnerShdw.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}

			if (!(this.color && oEffect.color && this.color.IsIdentical(oEffect.color) || !this.color && !oEffect.color)) {
				return false;
			}

			if (this.blurRad !== oEffect.blurRad) {
				return false;
			}

			if (this.dir !== oEffect.dir) {
				return false;
			}

			if (this.dist !== oEffect.dist) {
				return false;
			}

			return true;
		};
		function CLumEffect() {
			CBaseNoIdObject.call(this);
			this.bright = null;
			this.contrast = null;
		}

		InitClass(CLumEffect, CBaseNoIdObject, 0);
		CLumEffect.prototype.Type = EFFECT_TYPE_LUM;
		CLumEffect.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_LUM);
			writeLong(w, this.bright);
			writeLong(w, this.contrast);
		};
		CLumEffect.prototype.Read_FromBinary = function (r) {
			this.bright = readLong(r);
			this.contrast = readLong(r);
		};
		CLumEffect.prototype.createDuplicate = function () {
			var oCopy = new CLumEffect();
			oCopy.bright = this.bright;
			oCopy.contrast = this.contrast;
			return oCopy;
		};
		CLumEffect.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (this.bright !== oEffect.bright) {
				return false;
			}
			if (this.contrast !== oEffect.contrast) {
				return false;
			}
			return true;
		};

		function COuterShdw() {
			CBaseNoIdObject.call(this);
			this.color = new CUniColor();
			this.algn = null;
			this.blurRad = null;
			this.dir = null;
			this.dist = null;
			this.kx = null;
			this.ky = null;
			this.rotWithShape = null;
			this.sx = null;
			this.sy = null;
		}

		InitClass(COuterShdw, CBaseNoIdObject, 0);
		COuterShdw.prototype.Type = EFFECT_TYPE_OUTERSHDW;
		COuterShdw.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_OUTERSHDW);
			this.color.Write_ToBinary(w);
			writeLong(w, this.algn);
			writeLong(w, this.blurRad);
			writeLong(w, this.dir);
			writeLong(w, this.dist);
			writeLong(w, this.kx);
			writeLong(w, this.ky);
			writeBool(w, this.rotWithShape);
			writeLong(w, this.sx);
			writeLong(w, this.sy);
		};
		COuterShdw.prototype.Read_FromBinary = function (r) {
			this.color.Read_FromBinary(r);
			this.algn = readLong(r);
			this.blurRad = readLong(r);
			this.dir = readLong(r);
			this.dist = readLong(r);
			this.kx = readLong(r);
			this.ky = readLong(r);
			this.rotWithShape = readBool(r);
			this.sx = readLong(r);
			this.sy = readLong(r);
		};
		COuterShdw.prototype.createDuplicate = function () {
			var oCopy = new COuterShdw();
			this.fillObject(oCopy);
			return oCopy;
		};

		COuterShdw.prototype.fillObject = function (oCopy) {
			oCopy.color = this.color.createDuplicate();
			oCopy.algn = this.algn;
			oCopy.blurRad = this.blurRad;
			oCopy.dir = this.dir;
			oCopy.dist = this.dist;
			oCopy.kx = this.kx;
			oCopy.ky = this.ky;
			oCopy.rotWithShape = this.rotWithShape;
			oCopy.sx = this.sx;
			oCopy.sy = this.sy;
		};
		COuterShdw.prototype.IsIdentical = function (other) {
			if (!other) {
				return false;
			}
			if (!this.color && other.color || this.color && !other.color || !this.color.IsIdentical(other.color)) {
				return false;
			}
			if (other.algn !== this.algn ||
				other.blurRad !== this.blurRad ||
				other.dir !== this.dir ||
				other.dist !== this.dist ||
				other.kx !== this.kx ||
				other.ky !== this.ky ||
				other.rotWithShape !== this.rotWithShape ||
				other.sx !== this.sx ||
				other.sy !== this.sy) {
				return false;
			}
			return true;
		};
		COuterShdw.prototype.getAscShdw = function() {
			const oCopy = new asc_CShadowProperty();
			this.fillObject(oCopy);
			return oCopy;
		};

		const RECT_ALIGN_B = 0;
		const RECT_ALIGN_BL = 1;
		const RECT_ALIGN_BR = 2;
		const RECT_ALIGN_CTR = 3;
		const RECT_ALIGN_L = 4;
		const RECT_ALIGN_R = 5;
		const RECT_ALIGN_T = 6;
		const RECT_ALIGN_TL = 7;
		const RECT_ALIGN_TR = 8;
		function asc_CShadowProperty() {
			COuterShdw.call(this);
			this.setDefault();
		}

		InitClass(asc_CShadowProperty, COuterShdw, 0);
		asc_CShadowProperty.prototype.setDefault = function() {
			this.algn = 7;
			this.blurRad = 50800;
			this.color = new CUniColor();
			this.color.color = new CPrstColor();
			this.color.color.id = "black";
			this.putTransparency(60);
			this.dir = 2700000;
			this.dist = 38100;
			this.rotWithShape = false;
		};
		asc_CShadowProperty.prototype.putPreset = function (sAlgn) {
			this.setDefault();
			switch (sAlgn) {
				case "l": {
					this.algn = RECT_ALIGN_L;
					this.dir = null;
					break;
				}
				case "t": {
					this.algn = RECT_ALIGN_T;
					this.dir = 5400000;
					break;
				}
				case "r": {
					this.algn = RECT_ALIGN_R;
					this.dir = 10800000;
					break;
				}
				case "b": {
					this.algn = null;
					this.dir = 16200000;
					break;
				}
				case "tl": {
					this.algn = RECT_ALIGN_TL;
					this.dir = 2700000;
					break;
				}
				case "tr": {
					this.algn = RECT_ALIGN_TR;
					this.dir = 8100000;
					break;
				}
				case "bl": {
					this.algn = RECT_ALIGN_BL;
					this.dir = 18900000;
					break;
				}
				case "br": {
					this.algn = RECT_ALIGN_BR;
					this.dir = 13500000;
					break;
				}
				case "ctr": {
					this.algn = RECT_ALIGN_CTR;
					this.dir = null;
					this.sx = 102000;
					this.sy = 102000;
					this.dist = null;
					break;
				}
			}
		};
		asc_CShadowProperty.prototype.getPreset = function() {
			const aPresets = ["l", "t", "r", "b", "tl", "tr", "bl", "br", "ctr"];
			for(let nPrst = 0; nPrst < aPresets.length; ++nPrst) {
				let oShd = new asc_CShadowProperty();
				let sPrst = aPresets[nPrst];
				oShd.putPreset(sPrst);
				if(this.IsIdentical(oShd)) {
					return sPrst;
				}
			}
			return null;
		};
		asc_CShadowProperty.prototype.createDuplicate = function () {
			var oCopy = new asc_CShadowProperty();
			this.fillObject(oCopy);
			return oCopy;
		};
		asc_CShadowProperty.prototype.getTransparency = function() {
			if(!this.color) {
				return 0;
			}
			return this.color.getTransparency();
		};
		asc_CShadowProperty.prototype.putTransparency = function(nVal) {
			if(!this.color) {
				return;
			}
			if(!this.color.Mods) {
				this.color.Mods = new CColorModifiers();
			}
			let oMod = this.color.Mods.getMod("alpha");
			if(!oMod) {
				oMod = new CColorMod("alpha", (100 - nVal) * 1000 + 0.5 >> 0);
				this.color.Mods.addMod(oMod);
			}
			else {
				oMod.setVal((100 - nVal) * 1000 + 0.5 >> 0);
			}
		};
		asc_CShadowProperty.prototype.getSize = function() {
			let nSX = this.sx !== null ? this.sx : 100000;
			let nSY = this.sy !== null ? this.sy : 100000;
			return Math.max(nSX, nSY) / 1000;
		};
		asc_CShadowProperty.prototype.putSize = function(nVal) {
			this.sx = nVal * 1000 + 0.5 >> 0;
			this.sy = this.sx;
		};

		asc_CShadowProperty.prototype.getAngle = function() {
			let nAngle = this.dir || 0;
			return nAngle / 60000;
		};
		asc_CShadowProperty.prototype.putAngle = function(nVal) {
			this.dir = nVal * 60000 + 0.5 >> 0;
		};
		asc_CShadowProperty.prototype.getDistance = function() {
			let nDist = this.dist || 0;
			return nDist / 36000;
		};
		asc_CShadowProperty.prototype.putDistance = function(nVal) {
			this.dist = nVal * 36000 + 0.5 >> 0;
		};
		asc_CShadowProperty.prototype.getColor = function() {
			return AscCommon.CreateAscColor(this.color);
		};
		asc_CShadowProperty.prototype.putColor = function(oAscColor) {
			let nTransparency = this.getTransparency();
			let nFlag;
			if(Asc.editor.editorId === AscCommon.c_oEditorId.Word) {
				nFlag = 1;
			}
			else {
				nFlag = 0;
			}
			this.color = AscFormat.CorrectUniColor(oAscColor, this.color, nFlag);
			this.putTransparency(nTransparency);
		};
		asc_CShadowProperty.prototype.checkColor = function(oTheme, oColorMap) {
			if(this.color) {
				if(oTheme && oColorMap) {
					this.color.check(oTheme, oColorMap);
				}
			}
		};


		asc_CShadowProperty.prototype["getTransparency"] = asc_CShadowProperty.prototype.getTransparency;
		asc_CShadowProperty.prototype["putTransparency"] = asc_CShadowProperty.prototype.putTransparency;
		asc_CShadowProperty.prototype["getSize"] = asc_CShadowProperty.prototype.getSize;
		asc_CShadowProperty.prototype["putSize"] = asc_CShadowProperty.prototype.putSize;
		asc_CShadowProperty.prototype["getAngle"] = asc_CShadowProperty.prototype.getAngle;
		asc_CShadowProperty.prototype["putAngle"] = asc_CShadowProperty.prototype.putAngle;
		asc_CShadowProperty.prototype["getDistance"] = asc_CShadowProperty.prototype.getDistance;
		asc_CShadowProperty.prototype["putDistance"] = asc_CShadowProperty.prototype.putDistance;
		asc_CShadowProperty.prototype["getColor"] = asc_CShadowProperty.prototype.getColor;
		asc_CShadowProperty.prototype["putColor"] = asc_CShadowProperty.prototype.putColor;
		asc_CShadowProperty.prototype["putPreset"] = asc_CShadowProperty.prototype.putPreset;
		asc_CShadowProperty.prototype["getPreset"] = asc_CShadowProperty.prototype.getPreset;


		window['Asc'] = window['Asc'] || {};
		window['Asc']['asc_CShadowProperty'] = window['Asc'].asc_CShadowProperty = asc_CShadowProperty;

		function CPrstShdw() {
			CBaseNoIdObject.call(this);
			this.color = new CUniColor();
			this.prst = null;
			this.dir = null;
			this.dist = null;
		}

		InitClass(CPrstShdw, CBaseNoIdObject, 0);
		CPrstShdw.prototype.Type = EFFECT_TYPE_PRSTSHDW;
		CPrstShdw.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_PRSTSHDW);
			this.color.Write_ToBinary(w);
			writeLong(w, this.prst);
			writeLong(w, this.dir);
			writeLong(w, this.dist);
		};
		CPrstShdw.prototype.Read_FromBinary = function (r) {
			this.color.Read_FromBinary(r);
			this.prst = readLong(r);
			this.dir = readLong(r);
			this.dist = readLong(r);
		};
		CPrstShdw.prototype.createDuplicate = function () {
			var oCopy = new CPrstShdw();
			oCopy.color = this.color.createDuplicate();
			oCopy.prst = this.prst;
			oCopy.dir = this.dir;
			oCopy.dist = this.dist;
			return oCopy;
		};
		CPrstShdw.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (!(this.color && oEffect.color && this.color.IsIdentical(oEffect.color) || !this.color && !oEffect.color)) {
				return false;
			}
			if (this.prst !== oEffect.prst) {
				return false;
			}
			if (this.dir !== oEffect.dir) {
				return false;
			}
			if (this.dist !== oEffect.dist) {
				return false;
			}
			return true;
		};

		function CReflection() {
			CBaseNoIdObject.call(this);
			this.algn = null;
			this.blurRad = null;
			this.stA = null;
			this.endA = null;
			this.stPos = null;
			this.endPos = null;
			this.dir = null;
			this.fadeDir = null;
			this.dist = null;
			this.kx = null;
			this.ky = null;
			this.rotWithShape = null;
			this.sx = null;
			this.sy = null;
		}

		InitClass(CReflection, CBaseNoIdObject, 0);
		CReflection.prototype.Type = EFFECT_TYPE_REFLECTION;
		CReflection.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_REFLECTION);
			writeLong(w, this.algn);
			writeLong(w, this.blurRad);
			writeLong(w, this.stA);
			writeLong(w, this.endA);
			writeLong(w, this.stPos);
			writeLong(w, this.endPos);
			writeLong(w, this.dir);
			writeLong(w, this.fadeDir);
			writeLong(w, this.dist);
			writeLong(w, this.kx);
			writeLong(w, this.ky);
			writeBool(w, this.rotWithShape);
			writeLong(w, this.sx);
			writeLong(w, this.sy);
		};
		CReflection.prototype.Read_FromBinary = function (r) {
			this.algn = readLong(r);
			this.blurRad = readLong(r);
			this.stA = readLong(r);
			this.endA = readLong(r);
			this.stPos = readLong(r);
			this.endPos = readLong(r);
			this.dir = readLong(r);
			this.fadeDir = readLong(r);
			this.dist = readLong(r);
			this.kx = readLong(r);
			this.ky = readLong(r);
			this.rotWithShape = readBool(r);
			this.sx = readLong(r);
			this.sy = readLong(r);
		};
		CReflection.prototype.createDuplicate = function () {
			var oCopy = new CReflection();
			oCopy.algn = this.algn;
			oCopy.blurRad = this.blurRad;
			oCopy.stA = this.stA;
			oCopy.endA = this.endA;
			oCopy.stPos = this.stPos;
			oCopy.endPos = this.endPos;
			oCopy.dir = this.dir;
			oCopy.fadeDir = this.fadeDir;
			oCopy.dist = this.dist;
			oCopy.kx = this.kx;
			oCopy.ky = this.ky;
			oCopy.rotWithShape = this.rotWithShape;
			oCopy.sx = this.sx;
			oCopy.sy = this.sy;
			return oCopy;
		};
		CReflection.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (this.algn !== oEffect.algn) {
				return false;
			}
			if (this.blurRad !== oEffect.blurRad) {
				return false;
			}
			if (this.stA !== oEffect.stA) {
				return false;
			}
			if (this.endA !== oEffect.endA) {
				return false;
			}
			if (this.stPos !== oEffect.stPos) {
				return false;
			}
			if (this.endPos !== oEffect.endPos) {
				return false;
			}
			if (this.dir !== oEffect.dir) {
				return false;
			}
			if (this.fadeDir !== oEffect.fadeDir) {
				return false;
			}
			if (this.dist !== oEffect.dist) {
				return false;
			}
			if (this.kx !== oEffect.kx) {
				return false;
			}
			if (this.ky !== oEffect.ky) {
				return false;
			}
			if (this.rotWithShape !== oEffect.rotWithShape) {
				return false;
			}
			if (this.sx !== oEffect.sx) {
				return false;
			}
			if (this.sy !== oEffect.sy) {
				return false;
			}
			return true;
		};
		function CRelOff() {
			CBaseNoIdObject.call(this);
			this.tx = null;
			this.ty = null;
		}

		InitClass(CRelOff, CBaseNoIdObject, 0);
		CRelOff.prototype.Type = EFFECT_TYPE_RELOFF;
		CRelOff.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_RELOFF);
			writeLong(w, this.tx);
			writeLong(w, this.ty);
		};
		CRelOff.prototype.Read_FromBinary = function (r) {
			this.tx = readLong(r);
			this.ty = readLong(r);
		};
		CRelOff.prototype.createDuplicate = function () {
			var oCopy = new CRelOff();
			oCopy.tx = this.tx;
			oCopy.ty = this.ty;
			return oCopy;
		};
		CRelOff.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (this.tx !== oEffect.tx) {
				return false;
			}
			if (this.ty !== oEffect.ty) {
				return false;
			}

			return true;
		};
		function CSoftEdge() {
			CBaseNoIdObject.call(this);
			this.rad = null;
		}

		InitClass(CSoftEdge, CBaseNoIdObject, 0);
		CSoftEdge.prototype.Type = EFFECT_TYPE_SOFTEDGE;
		CSoftEdge.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_SOFTEDGE);
			writeLong(w, this.rad);
		};
		CSoftEdge.prototype.Read_FromBinary = function (r) {
			this.rad = readLong(r);
		};
		CSoftEdge.prototype.createDuplicate = function () {
			var oCopy = new CSoftEdge();
			oCopy.rad = this.rad;
			return oCopy;
		};
		CSoftEdge.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}

			if (this.rad !== oEffect.rad) {
				return false;
			}

			return true;
		};

		function CTintEffect() {
			CBaseNoIdObject.call(this);
			this.amt = null;
			this.hue = null;
		}

		InitClass(CTintEffect, CBaseNoIdObject, 0);
		CTintEffect.prototype.Type = EFFECT_TYPE_TINTEFFECT;
		CTintEffect.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_TINTEFFECT);
			writeLong(w, this.amt);
			writeLong(w, this.hue);
		};
		CTintEffect.prototype.Read_FromBinary = function (r) {
			this.amt = readLong(r);
			this.hue = readLong(r);
		};
		CTintEffect.prototype.createDuplicate = function () {
			var oCopy = new CTintEffect();
			oCopy.amt = this.amt;
			oCopy.hue = this.hue;
			return oCopy;
		};
		CTintEffect.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}

			if (this.amt !== oEffect.amt) {
				return false;
			}

			if (this.hue !== oEffect.hue) {
				return false;
			}

			return true;
		};
		function CXfrmEffect() {
			CBaseNoIdObject.call(this);
			this.kx = null;
			this.ky = null;
			this.sx = null;
			this.sy = null;
			this.tx = null;
			this.ty = null;
		}

		InitClass(CXfrmEffect, CBaseNoIdObject, 0);
		CXfrmEffect.prototype.Type = EFFECT_TYPE_XFRM;
		CXfrmEffect.prototype.Write_ToBinary = function (w) {
			w.WriteLong(EFFECT_TYPE_XFRM);
			writeLong(w, this.kx);
			writeLong(w, this.ky);
			writeLong(w, this.sx);
			writeLong(w, this.sy);
			writeLong(w, this.tx);
			writeLong(w, this.ty);
		};
		CXfrmEffect.prototype.Read_FromBinary = function (r) {
			this.kx = readLong(r);
			this.ky = readLong(r);
			this.sx = readLong(r);
			this.sy = readLong(r);
			this.tx = readLong(r);
			this.ty = readLong(r);
		};
		CXfrmEffect.prototype.createDuplicate = function () {
			var oCopy = new CXfrmEffect();
			oCopy.kx = this.kx;
			oCopy.ky = this.ky;
			oCopy.sx = this.sx;
			oCopy.sy = this.sy;
			oCopy.tx = this.tx;
			oCopy.ty = this.ty;
			return oCopy;
		};
		CXfrmEffect.prototype.IsIdentical = function (oEffect) {
			if (this.Type !== oEffect.Type) {
				return false;
			}
			if (this.kx !== oEffect.kx) {
				return false;
			}
			if (this.ky !== oEffect.ky) {
				return false;
			}
			if (this.sx !== oEffect.sx) {
				return false;
			}
			if (this.sy !== oEffect.sy) {
				return false;
			}
			if (this.tx !== oEffect.tx) {
				return false;
			}
			if (this.ty !== oEffect.ty) {
				return false;
			}

			return true;
		};

//-----------------


		function CSolidFill() {
			CBaseNoIdObject.call(this);
			this.type = c_oAscFill.FILL_TYPE_SOLID;
			this.color = null;
		}

		InitClass(CSolidFill, CBaseNoIdObject, 0);
		CSolidFill.prototype.check = function (theme, colorMap) {
			if (this.color) {
				this.color.check(theme, colorMap);
			}
		};
		CSolidFill.prototype.saveSourceFormatting = function () {
			var _ret = new CSolidFill();
			if (this.color) {
				_ret.color = this.color.saveSourceFormatting();
			}
			return _ret;
		};
		CSolidFill.prototype.setColor = function (color) {
			this.color = color;
		};
		CSolidFill.prototype.Write_ToBinary = function (w) {
			if (this.color) {
				w.WriteBool(true);
				this.color.Write_ToBinary(w);
			} else {
				w.WriteBool(false);
			}
		};
		CSolidFill.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				this.color = new CUniColor();
				this.color.Read_FromBinary(r);
			}
		};
		CSolidFill.prototype.checkWordMods = function () {
			return this.color && this.color.checkWordMods();
		};
		CSolidFill.prototype.convertToPPTXMods = function () {
			this.color && this.color.convertToPPTXMods();
		};
		CSolidFill.prototype.canConvertPPTXModsToWord = function () {
			return this.color && this.color.canConvertPPTXModsToWord();
		};
		CSolidFill.prototype.convertToWordMods = function () {
			this.color && this.color.convertToWordMods();
		};
		CSolidFill.prototype.IsIdentical = function (fill) {
			if (fill == null) {
				return false;
			}
			if (fill.type !== c_oAscFill.FILL_TYPE_SOLID) {
				return false;
			}

			if (this.color) {
				return this.color.IsIdentical(fill.color);
			}
			return (fill.color === null);

		};
		CSolidFill.prototype.createDuplicate = function () {
			var duplicate = new CSolidFill();
			if (this.color) {
				duplicate.color = this.color.createDuplicate();
			}
			return duplicate;
		};
		CSolidFill.prototype.compare = function (fill) {
			if (fill == null || fill.type !== c_oAscFill.FILL_TYPE_SOLID) {
				return null;
			}
			if (this.color && fill.color) {
				var _ret = new CSolidFill();
				_ret.color = this.color.compare(fill.color);
				return _ret;
			}
			return null;
		};

		function CGs() {
			CBaseNoIdObject.call(this);
			this.color = null;
			this.pos = 0;
		}

		InitClass(CGs, CBaseNoIdObject, 0);
		CGs.prototype.setColor = function (color) {
			this.color = color;
		};
		CGs.prototype.setPos = function (pos) {
			this.pos = pos;
		};
		CGs.prototype.saveSourceFormatting = function () {
			var _ret = new CGs();
			_ret.pos = this.pos;
			if (this.color) {
				_ret.color = this.color.saveSourceFormatting();
			}
			return _ret;
		};
		CGs.prototype.Write_ToBinary = function (w) {
			w.WriteBool(isRealObject(this.color));
			if (isRealObject(this.color)) {
				this.color.Write_ToBinary(w);
			}
			writeLong(w, this.pos);
		};
		CGs.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				this.color = new CUniColor();
				this.color.Read_FromBinary(r);
			} else {
				this.color = null;
			}
			this.pos = readLong(r);
		};
		CGs.prototype.IsIdentical = function (fill) {
			if (!fill)
				return false;
			if (this.pos !== fill.pos)
				return false;

			if (!this.color && fill.color || this.color && !fill.color || (this.color && fill.color && !this.color.IsIdentical(fill.color)))
				return false;
			return true;
		};
		CGs.prototype.createDuplicate = function () {
			var duplicate = new CGs();
			duplicate.pos = this.pos;
			if (this.color) {
				duplicate.color = this.color.createDuplicate();
			}
			return duplicate;
		};
		CGs.prototype.compare = function (gs) {
			if (gs.pos !== this.pos) {
				return null;
			}
			var compare_unicolor = this.color.compare(gs.color);
			if (!isRealObject(compare_unicolor)) {
				return null;
			}
			var ret = new CGs();
			ret.color = compare_unicolor;
			ret.pos = gs.pos === this.pos ? this.pos : 0;
			return ret;
		};

		function GradLin() {
			CBaseNoIdObject.call(this);
			this.angle = 5400000;
			this.scale = true;
		}

		InitClass(GradLin, CBaseNoIdObject, 0);
		GradLin.prototype.setAngle = function (angle) {
			this.angle = angle;
		};
		GradLin.prototype.setScale = function (scale) {
			this.scale = scale;
		};
		GradLin.prototype.Write_ToBinary = function (w) {
			writeLong(w, this.angle);
			writeBool(w, this.scale);
		};
		GradLin.prototype.Read_FromBinary = function (r) {
			this.angle = readLong(r);
			this.scale = readBool(r);
		};
		GradLin.prototype.IsIdentical = function (lin) {
			if (this.angle != lin.angle)
				return false;
			if (this.scale != lin.scale)
				return false;

			return true;
		};
		GradLin.prototype.createDuplicate = function () {
			var duplicate = new GradLin();
			duplicate.angle = this.angle;
			duplicate.scale = this.scale;
			return duplicate;
		};
		GradLin.prototype.compare = function (lin) {
			return null;
		};

		function GradPath() {
			CBaseNoIdObject.call(this);
			this.path = 0;
			this.rect = null;
		}

		InitClass(GradPath, CBaseNoIdObject, 0);
		GradPath.prototype.setPath = function (path) {
			this.path = path;
		};
		GradPath.prototype.setRect = function (rect) {
			this.rect = rect;
		};
		GradPath.prototype.Write_ToBinary = function (w) {
			writeLong(w, this.path);
			w.WriteBool(isRealObject(this.rect));
			if (isRealObject(this.rect)) {
				this.rect.Write_ToBinary(w);
			}
		};
		GradPath.prototype.Read_FromBinary = function (r) {
			this.path = readLong(r);
			if (r.GetBool()) {
				this.rect = new CSrcRect();
				this.rect.Read_FromBinary(r);
			}
		};
		GradPath.prototype.IsIdentical = function (path) {
			if (this.path !== path.path)
				return false;
			return true;
		};
		GradPath.prototype.createDuplicate = function () {
			var duplicate = new GradPath();
			duplicate.path = this.path;
			return duplicate;
		};
		GradPath.prototype.compare = function (path) {
			return null;
		};

		function CGradFill() {
			CBaseFill.call(this);
			// пока просто front color
			this.colors = [];

			this.lin = null;
			this.path = null;

			this.rotateWithShape = null;
		}

		InitClass(CGradFill, CBaseFill, 0);
		CGradFill.prototype.type = c_oAscFill.FILL_TYPE_GRAD;
		CGradFill.prototype.saveSourceFormatting = function () {
			var _ret = new CGradFill();
			if (this.lin) {
				_ret.lin = this.lin.createDuplicate();
			}
			if (this.path) {
				_ret.path = this.path.createDuplicate();
			}
			for (var i = 0; i < this.colors.length; ++i) {
				_ret.colors.push(this.colors[i].saveSourceFormatting());
			}
			return _ret;
		};
		CGradFill.prototype.check = function (theme, colorMap) {
			for (var i = 0; i < this.colors.length; ++i) {
				if (this.colors[i].color) {
					this.colors[i].color.check(theme, colorMap);
				}
			}
		};
		CGradFill.prototype.checkWordMods = function () {
			for (var i = 0; i < this.colors.length; ++i) {
				if (this.colors[i] && this.colors[i].color && this.colors[i].color.checkWordMods()) {
					return true;
				}
			}
			return false;
		};
		CGradFill.prototype.convertToPPTXMods = function () {
			for (var i = 0; i < this.colors.length; ++i) {
				this.colors[i] && this.colors[i].color && this.colors[i].color.convertToPPTXMods();

			}
		};
		CGradFill.prototype.canConvertPPTXModsToWord = function () {
			for (var i = 0; i < this.colors.length; ++i) {
				if (this.colors[i] && this.colors[i].color && this.colors[i].color.canConvertPPTXModsToWord()) {
					return true;
				}
			}
			return false;
		};
		CGradFill.prototype.convertToWordMods = function () {
			for (var i = 0; i < this.colors.length; ++i) {
				this.colors[i] && this.colors[i].color && this.colors[i].color.convertToWordMods();

			}
		};
		CGradFill.prototype.addColor = function (color) {
			this.colors.push(color);
		};
		CGradFill.prototype.setLin = function (lin) {
			this.lin = lin;
		};
		CGradFill.prototype.setPath = function (path) {
			this.path = path;
		};
		CGradFill.prototype.Write_ToBinary = function (w) {
			w.WriteLong(this.colors.length);
			for (var i = 0; i < this.colors.length; ++i) {
				this.colors[i].Write_ToBinary(w);
			}
			w.WriteBool(isRealObject(this.lin));
			if (isRealObject(this.lin)) {
				this.lin.Write_ToBinary(w);
			}
			w.WriteBool(isRealObject(this.path));
			if (isRealObject(this.path)) {
				this.path.Write_ToBinary(w);
			}
			writeBool(w, this.rotateWithShape);
		};
		CGradFill.prototype.Read_FromBinary = function (r) {
			var len = r.GetLong();
			for (var i = 0; i < len; ++i) {
				this.colors[i] = new CGs();
				this.colors[i].Read_FromBinary(r);
			}
			if (r.GetBool()) {
				this.lin = new GradLin();
				this.lin.Read_FromBinary(r);
			} else {
				this.lin = null;
			}
			if (r.GetBool()) {
				this.path = new GradPath();
				this.path.Read_FromBinary(r);
			} else {
				this.path = null;
			}
			this.rotateWithShape = readBool(r);
		};
		CGradFill.prototype.IsIdentical = function (fill) {
			if (fill == null) {
				return false;
			}
			if (fill.type !== c_oAscFill.FILL_TYPE_GRAD) {
				return false;
			}
			if (fill.colors.length !== this.colors.length) {
				return false;
			}
			for (var i = 0; i < this.colors.length; ++i) {
				if (!this.colors[i].IsIdentical(fill.colors[i])) {
					return false;
				}
			}

			if (!this.path && fill.path || this.path && !fill.path || (this.path && fill.path && !this.path.IsIdentical(fill.path)))
				return false;

			if (!this.lin && fill.lin || !fill.lin && this.lin || (this.lin && fill.lin && !this.lin.IsIdentical(fill.lin)))
				return false;

			if(this.rotateWithShape !== fill.rotateWithShape) {
				return false;
			}

			return true;
		};
		CGradFill.prototype.createDuplicate = function () {
			var duplicate = new CGradFill();
			for (var i = 0; i < this.colors.length; ++i) {
				duplicate.colors[i] = this.colors[i].createDuplicate();
			}

			if (this.lin)
				duplicate.lin = this.lin.createDuplicate();

			if (this.path)
				duplicate.path = this.path.createDuplicate();

			if (this.rotateWithShape != null)
				duplicate.rotateWithShape = this.rotateWithShape;

			return duplicate;
		};
		CGradFill.prototype.compare = function (fill) {
			if (fill == null || fill.type !== c_oAscFill.FILL_TYPE_GRAD) {
				return null;
			}
			var _ret = new CGradFill();
			if (this.lin == null || fill.lin == null)
				_ret.lin = null;
			else {
				_ret.lin = new GradLin();
				_ret.lin.angle = this.lin && this.lin.angle === fill.lin.angle ? fill.lin.angle : 5400000;
				_ret.lin.scale = this.lin && this.lin.scale === fill.lin.scale ? fill.lin.scale : true;
			}
			if (this.path == null || fill.path == null) {
				_ret.path = null;
			} else {
				_ret.path = new GradPath();
			}

			if (this.colors.length === fill.colors.length) {
				for (var i = 0; i < this.colors.length; ++i) {
					var compare_unicolor = this.colors[i].compare(fill.colors[i]);
					if (!isRealObject(compare_unicolor)) {
						return null;
					}
					_ret.colors[i] = compare_unicolor;
				}
			}
			if(this.rotateWithShape === fill.rotateWithShape) {
				_ret.rotateWithShape = this.rotateWithShape;
			}
			return _ret;
		};
		CGradFill.prototype.getColorsCount = function() {
			return this.colors.length;
		};

		function CPattFill() {
			CBaseFill.call(this);
			this.ftype = 0;
			this.fgClr = null;//new CUniColor();
			this.bgClr = null;//new CUniColor();
		}

		InitClass(CPattFill, CBaseFill, 0);
		CPattFill.prototype.type = c_oAscFill.FILL_TYPE_PATT;
		CPattFill.prototype.check = function (theme, colorMap) {
			if (this.fgClr)
				this.fgClr.check(theme, colorMap);
			if (this.bgClr)
				this.bgClr.check(theme, colorMap);
		};
		CPattFill.prototype.checkWordMods = function () {
			if (this.fgClr && this.fgClr.checkWordMods()) {
				return true;
			}
			return this.bgClr && this.bgClr.checkWordMods();

		};
		CPattFill.prototype.saveSourceFormatting = function () {
			var _ret = new CPattFill();
			if (this.fgClr) {
				_ret.fgClr = this.fgClr.saveSourceFormatting();
			}
			if (this.bgClr) {
				_ret.bgClr = this.bgClr.saveSourceFormatting();
			}
			_ret.ftype = this.ftype;
			return _ret;
		};
		CPattFill.prototype.convertToPPTXMods = function () {
			this.fgClr && this.fgClr.convertToPPTXMods();
			this.bgClr && this.bgClr.convertToPPTXMods();
		};
		CPattFill.prototype.canConvertPPTXModsToWord = function () {
			if (this.fgClr && this.fgClr.canConvertPPTXModsToWord()) {
				return true;
			}
			return this.bgClr && this.bgClr.canConvertPPTXModsToWord();
		};
		CPattFill.prototype.convertToWordMods = function () {

			this.fgClr && this.fgClr.convertToWordMods();
			this.bgClr && this.bgClr.convertToWordMods();
		};
		CPattFill.prototype.setFType = function (fType) {
			this.ftype = fType;
		};
		CPattFill.prototype.setFgColor = function (fgClr) {
			this.fgClr = fgClr;
		};
		CPattFill.prototype.setBgColor = function (bgClr) {
			this.bgClr = bgClr;
		};
		CPattFill.prototype.Write_ToBinary = function (w) {
			writeLong(w, this.ftype);
			w.WriteBool(isRealObject(this.fgClr));
			if (isRealObject(this.fgClr)) {
				this.fgClr.Write_ToBinary(w);
			}
			w.WriteBool(isRealObject(this.bgClr));
			if (isRealObject(this.bgClr)) {
				this.bgClr.Write_ToBinary(w);
			}
		};
		CPattFill.prototype.Read_FromBinary = function (r) {
			this.ftype = readLong(r);
			if (r.GetBool()) {
				this.fgClr = new CUniColor();
				this.fgClr.Read_FromBinary(r);
			}
			if (r.GetBool()) {
				this.bgClr = new CUniColor();
				this.bgClr.Read_FromBinary(r);
			}
		};
		CPattFill.prototype.IsIdentical = function (fill) {
			if (fill == null) {
				return false;
			}
			if (fill.type !== c_oAscFill.FILL_TYPE_PATT && this.ftype !== fill.ftype) {
				return false;
			}

			return this.fgClr.IsIdentical(fill.fgClr) && this.bgClr.IsIdentical(fill.bgClr) && this.ftype === fill.ftype;
		};
		CPattFill.prototype.createDuplicate = function () {
			var duplicate = new CPattFill();
			duplicate.ftype = this.ftype;
			if (this.fgClr) {
				duplicate.fgClr = this.fgClr.createDuplicate();
			}
			if (this.bgClr) {
				duplicate.bgClr = this.bgClr.createDuplicate();
			}
			return duplicate;
		};
		CPattFill.prototype.compare = function (fill) {
			if (fill == null) {
				return null;
			}
			if (fill.type !== c_oAscFill.FILL_TYPE_PATT) {
				return null;
			}
			var _ret = new CPattFill();
			if (fill.ftype == this.ftype) {
				_ret.ftype = this.ftype;
			}
			if (this.fgClr) {
				_ret.fgClr = this.fgClr.compare(fill.fgClr);
			} else {
				_ret.fgClr = null;
			}
			if (this.bgClr) {
				_ret.bgClr = this.bgClr.compare(fill.bgClr);
			} else {
				_ret.bgClr = null;
			}
			if (!_ret.bgClr && !_ret.fgClr) {
				return null;
			}
			return _ret;
		};

		function CNoFill() {
			CBaseFill.call(this);
		}

		InitClass(CNoFill, CBaseFill, 0);
		CNoFill.prototype.type = c_oAscFill.FILL_TYPE_NOFILL;
		CNoFill.prototype.check = function () {
		};
		CNoFill.prototype.saveSourceFormatting = function () {
			return this.createDuplicate();
		};
		CNoFill.prototype.Write_ToBinary = function (w) {

		};
		CNoFill.prototype.Read_FromBinary = function (r) {
		};
		CNoFill.prototype.checkWordMods = function () {
			return false;

		};
		CNoFill.prototype.convertToPPTXMods = function () {
		};
		CNoFill.prototype.canConvertPPTXModsToWord = function () {
			return false;
		};
		CNoFill.prototype.convertToWordMods = function () {
		};
		CNoFill.prototype.createDuplicate = function () {
			return new CNoFill();
		};
		CNoFill.prototype.IsIdentical = function (fill) {
			if (fill == null) {
				return false;
			}
			return fill.type === c_oAscFill.FILL_TYPE_NOFILL;
		};
		CNoFill.prototype.compare = function (nofill) {
			if (nofill == null) {
				return null;
			}
			if (nofill.type === this.type) {
				return new CNoFill();
			}
			return null;
		};

		function CGrpFill() {
			CBaseFill.call(this);
		}

		InitClass(CGrpFill, CBaseFill, 0);
		CGrpFill.prototype.type = c_oAscFill.FILL_TYPE_GRP;
		CGrpFill.prototype.check = function () {
		};
		CGrpFill.prototype.getObjectType = function () {
			return AscDFH.historyitem_type_GrpFill;
		};
		CGrpFill.prototype.Write_ToBinary = function (w) {
		};
		CGrpFill.prototype.Read_FromBinary = function (r) {
		};
		CGrpFill.prototype.checkWordMods = function () {
			return false;

		};
		CGrpFill.prototype.convertToPPTXMods = function () {
		};
		CGrpFill.prototype.canConvertPPTXModsToWord = function () {
			return false;
		};
		CGrpFill.prototype.convertToWordMods = function () {
		};
		CGrpFill.prototype.createDuplicate = function () {
			return new CGrpFill();
		};
		CGrpFill.prototype.IsIdentical = function (fill) {
			if (fill == null) {
				return false;
			}
			return fill.type === c_oAscFill.FILL_TYPE_GRP;
		};
		CGrpFill.prototype.compare = function (grpfill) {
			if (grpfill == null) {
				return null;
			}
			if (grpfill.type === this.type) {
				return new CGrpFill();
			}
			return null;
		};


		function getGrayscaleValue(color) {
			return color.R * 0.2126 + color.G * 0.7152 + color.B * 0.0722;
		}

		function FormatRGBAColor(r, g, b) {
			this.R = r || 0;
			this.G = g || 0;
			this.B = b || 0;
			this.A = 255;
		}

		FormatRGBAColor.prototype.getGrayscaleValue = function () {
			return getGrayscaleValue(this);
		};

		function checkUniFillRasterImageId(oUnifill) {
			if (oUnifill) {
				return oUnifill.checkRasterImageId();
			}
			return null;
		}

		/** @constructor */
		function CUniFill() {
			CBaseNoIdObject.call(this);
			this.fill = null;
			this.transparent = null;
		}

		InitClass(CUniFill, CBaseNoIdObject, 0);
		CUniFill.prototype.check = function (theme, colorMap) {
			if (this.fill) {
				this.fill.check(theme, colorMap);
			}
		};
		CUniFill.prototype.addColorMod = function (mod) {
			if (this.fill) {
				switch (this.fill.type) {
					case c_oAscFill.FILL_TYPE_BLIP:
					case c_oAscFill.FILL_TYPE_NOFILL:
					case c_oAscFill.FILL_TYPE_GRP: {
						break;
					}
					case c_oAscFill.FILL_TYPE_SOLID: {
						if (this.fill.color && this.fill.color) {
							this.fill.color.addColorMod(mod);
						}
						break;
					}
					case c_oAscFill.FILL_TYPE_GRAD: {
						for (var i = 0; i < this.fill.colors.length; ++i) {
							if (this.fill.colors[i] && this.fill.colors[i].color) {
								this.fill.colors[i].color.addColorMod(mod);
							}
						}
						break;
					}
					case c_oAscFill.FILL_TYPE_PATT: {
						if (this.fill.bgClr) {
							this.fill.bgClr.addColorMod(mod);
						}
						if (this.fill.fgClr) {
							this.fill.fgClr.addColorMod(mod);
						}
						break;
					}
				}
			}
		};
		CUniFill.prototype.checkPhColor = function (unicolor) {
			if (this.fill) {
				switch (this.fill.type) {
					case c_oAscFill.FILL_TYPE_BLIP:
					case c_oAscFill.FILL_TYPE_NOFILL:
					case c_oAscFill.FILL_TYPE_GRP: {
						break;
					}
					case c_oAscFill.FILL_TYPE_SOLID: {
						if (this.fill.color && this.fill.color) {
							this.fill.color.checkPhColor(unicolor);
						}
						break;
					}
					case c_oAscFill.FILL_TYPE_GRAD: {
						for (var i = 0; i < this.fill.colors.length; ++i) {
							if (this.fill.colors[i] && this.fill.colors[i].color) {
								this.fill.colors[i].color.checkPhColor(unicolor);
							}
						}
						break;
					}
					case c_oAscFill.FILL_TYPE_PATT: {
						if (this.fill.bgClr) {
							this.fill.bgClr.checkPhColor(unicolor);
						}
						if (this.fill.fgClr) {
							this.fill.fgClr.checkPhColor(unicolor);
						}
						break;
					}
				}
			}
		};
		CUniFill.prototype.checkPatternType = function (nType) {
			if (this.fill) {
				if (this.fill.type === c_oAscFill.FILL_TYPE_PATT) {
					this.fill.ftype = nType;
				}
			}
		};
		CUniFill.prototype.Get_TextBackGroundColor = function () {
			if (!this.fill) {
				return undefined;
			}
			var oColor = undefined, RGBA;
			switch (this.fill.type) {
				case c_oAscFill.FILL_TYPE_SOLID: {
					if (this.fill.color) {
						RGBA = this.fill.color.RGBA;
						if (RGBA) {
							oColor = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, false);
						}
					}
					break;
				}
				case c_oAscFill.FILL_TYPE_PATT: {
					var oClr;
					if (this.fill.ftype === 38) {
						oClr = this.fill.fgClr;
					} else {
						oClr = this.fill.bgClr;
					}
					if (oClr) {
						RGBA = oClr.RGBA;
						if (RGBA) {
							oColor = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, false);
						}
					}
					break;
				}
			}
			return oColor;
		};
		CUniFill.prototype.checkWordMods = function () {
			return this.fill && this.fill.checkWordMods();
		};
		CUniFill.prototype.convertToPPTXMods = function () {
			this.fill && this.fill.convertToPPTXMods();
		};
		CUniFill.prototype.canConvertPPTXModsToWord = function () {
			return this.fill && this.fill.canConvertPPTXModsToWord();
		};
		CUniFill.prototype.convertToWordMods = function () {
			this.fill && this.fill.convertToWordMods();
		};
		CUniFill.prototype.setFill = function (fill) {
			this.fill = fill;
		};
		CUniFill.prototype.setTransparent = function (transparent) {
			this.transparent = transparent;
		};
		CUniFill.prototype.Set_FromObject = function (o) {
			//TODO:
		};
		CUniFill.prototype.Write_ToBinary = function (w) {
			writeDouble(w, this.transparent);
			w.WriteBool(isRealObject(this.fill));

			if (isRealObject(this.fill)) {
				w.WriteLong(this.fill.type);
				this.fill.Write_ToBinary(w);
			}

		};
		CUniFill.prototype.Read_FromBinary = function (r) {
			this.transparent = readDouble(r);
			if (r.GetBool()) {
				var type = r.GetLong();
				switch (type) {
					case c_oAscFill.FILL_TYPE_BLIP: {
						this.fill = new CBlipFill();
						this.fill.Read_FromBinary(r);
						break;
					}
					case c_oAscFill.FILL_TYPE_NOFILL: {
						this.fill = new CNoFill();
						this.fill.Read_FromBinary(r);
						break;
					}
					case c_oAscFill.FILL_TYPE_SOLID: {
						this.fill = new CSolidFill();
						this.fill.Read_FromBinary(r);
						break;
					}
					case c_oAscFill.FILL_TYPE_GRAD: {
						this.fill = new CGradFill();
						this.fill.Read_FromBinary(r);
						break;
					}
					case c_oAscFill.FILL_TYPE_PATT: {
						this.fill = new CPattFill();
						this.fill.Read_FromBinary(r);
						break;
					}
					case c_oAscFill.FILL_TYPE_GRP: {
						this.fill = new CGrpFill();
						this.fill.Read_FromBinary(r);
						break;
					}
				}
			}
		};
		CUniFill.prototype.calculate = function (theme, slide, layout, masterSlide, RGBA, colorMap) {
			if (this.fill) {
				if (this.fill.color) {
					this.fill.color.Calculate(theme, slide, layout, masterSlide, RGBA, colorMap);
				}
				if (this.fill.colors) {
					for (var i = 0; i < this.fill.colors.length; ++i) {
						this.fill.colors[i].color.Calculate(theme, slide, layout, masterSlide, RGBA, colorMap);
					}
				}
				if (this.fill.fgClr)
					this.fill.fgClr.Calculate(theme, slide, layout, masterSlide, RGBA, colorMap);
				if (this.fill.bgClr)
					this.fill.bgClr.Calculate(theme, slide, layout, masterSlide, RGBA, colorMap);
			}
		};
		CUniFill.prototype.getGrayscaleValue = function () {
			const RGBAColor = this.getRGBAColor();
			return getGrayscaleValue(RGBAColor);
		};
		CUniFill.prototype.getRGBAColor = function () {
			if (this.fill) {
				if (this.fill.type === c_oAscFill.FILL_TYPE_SOLID) {
					if (this.fill.color) {
						return this.fill.color.RGBA;
					} else {
						return new FormatRGBAColor();
					}
				}
				if (this.fill.type === c_oAscFill.FILL_TYPE_GRAD) {
					var RGBA = new FormatRGBAColor();
					var _colors = this.fill.colors;
					var _len = _colors.length;

					if (0 === _len)
						return RGBA;

					for (var i = 0; i < _len; i++) {
						RGBA.R += _colors[i].color.RGBA.R;
						RGBA.G += _colors[i].color.RGBA.G;
						RGBA.B += _colors[i].color.RGBA.B;
					}

					RGBA.R = (RGBA.R / _len) >> 0;
					RGBA.G = (RGBA.G / _len) >> 0;
					RGBA.B = (RGBA.B / _len) >> 0;

					return RGBA;
				}
				if (this.fill.type === c_oAscFill.FILL_TYPE_PATT) {
					return this.fill.fgClr.RGBA;
				}
				if (this.fill.type === c_oAscFill.FILL_TYPE_NOFILL) {
					return {R: 0, G: 0, B: 0};
				}
			}
			return new FormatRGBAColor();
		};

		CUniFill.prototype.getStartAnimRGBA = function () {
			let oFill = this.fill;
			if(!oFill) {
				return new FormatRGBAColor(255, 255, 255);
			}
			switch (oFill.type) {
				case c_oAscFill.FILL_TYPE_SOLID: {
					if (oFill.color) {
						return this.fill.color.RGBA;
					}
					else {
						return new FormatRGBAColor(255, 255, 255);
					}
				}
				case c_oAscFill.FILL_TYPE_GRAD: {
					let _colors = this.fill.colors;
					let _len = _colors.length;

					if (0 === _len) {
						return new FormatRGBAColor(255, 255, 255);
					}

					let oFirstColor = _colors[0].color;
					if(!oFirstColor) {
						return new FormatRGBAColor(255, 255, 255);
					}
					return oFirstColor.RGBA;
				}
				case c_oAscFill.FILL_TYPE_PATT: {
					if(oFill.fgClr) {
						return oFill.fgClr.RGBA
					}
					return new FormatRGBAColor(255, 255, 255);
				}
				case c_oAscFill.FILL_TYPE_NOFILL: {
					return new FormatRGBAColor(255, 255, 255);
				}
			}
			return new FormatRGBAColor(255, 255, 255);
		};
		CUniFill.prototype.createDuplicate = function () {
			var duplicate = new CUniFill();
			if (this.fill != null) {
				duplicate.fill = this.fill.createDuplicate();
			}
			duplicate.transparent = this.transparent;
			return duplicate;
		};
		CUniFill.prototype.saveSourceFormatting = function () {
			var duplicate = new CUniFill();
			if (this.fill) {
				if (this.fill.saveSourceFormatting) {
					duplicate.fill = this.fill.saveSourceFormatting();
				} else {
					duplicate.fill = this.fill.createDuplicate();
				}

			}
			duplicate.transparent = this.transparent;
			return duplicate;
		};
		CUniFill.prototype.merge = function (unifill) {
			if (unifill) {
				if (unifill.fill != null) {
					this.fill = unifill.fill.createDuplicate();
					if (this.fill.type === c_oAscFill.FILL_TYPE_PATT) {
						var _patt_fill = this.fill;
						if (!_patt_fill.fgClr) {
							_patt_fill.setFgColor(CreateUniColorRGB(0, 0, 0));
						}
						if (!_patt_fill.bgClr) {
							_patt_fill.bgClr = CreateUniColorRGB(255, 255, 255);
						}
					}
				}
				if (unifill.transparent != null) {
					this.transparent = unifill.transparent;
				}
			}
		};
		CUniFill.prototype.IsIdentical = function (unifill) {
			if (unifill == null) {
				return false;
			}

			if (isRealNumber(this.transparent) !== isRealNumber(unifill.transparent)
				|| isRealNumber(this.transparent) && this.transparent !== unifill.transparent) {
				return false;
			}

			if (this.fill == null && unifill.fill == null) {
				return true;
			}
			if (this.fill != null) {
				return this.fill.IsIdentical(unifill.fill);
			} else {
				return false;
			}
		};
		CUniFill.prototype.Is_Equal = function (unfill) {
			return this.IsIdentical(unfill);
		};
		CUniFill.prototype.isEqual = function (unfill) {
			return this.IsIdentical(unfill);
		};
		CUniFill.prototype.compare = function (unifill) {
			if (unifill == null) {
				return null;
			}
			var _ret = new CUniFill();
			if (!(this.fill == null || unifill.fill == null)) {
				if (this.fill.compare) {
					_ret.fill = this.fill.compare(unifill.fill);
				}
			}
			return _ret.fill;
		};
		CUniFill.prototype.isSolidFillRGB = function () {
			return (this.isSolidFill() && this.fill.color.color.type === window['Asc'].c_oAscColor.COLOR_TYPE_SRGB)
		};
		CUniFill.prototype.isSolidFill = function () {
			return !!(this.fill && this.fill.color && this.fill.color.color);
		};
		CUniFill.prototype.isSolidFillScheme = function () {
			return (this.isSolidFill() && this.fill.color.color.type === window['Asc'].c_oAscColor.COLOR_TYPE_SCHEME)
		};
		CUniFill.prototype.isNoFill = function () {
			return this.fill && this.fill.type === window['Asc'].c_oAscFill.FILL_TYPE_NOFILL;
		};
		CUniFill.prototype.isVisible = function () {
			return this.fill && this.fill.type !== window['Asc'].c_oAscFill.FILL_TYPE_NOFILL;
		};
		CUniFill.prototype.FILL_NAMES = {
			"blipFill": true,
			"gradFill": true,
			"grpFill": true,
			"noFill": true,
			"pattFill": true,
			"solidFill": true
		};
		CUniFill.prototype.isFillName = function (sName) {
			return !!CUniFill.prototype.FILL_NAMES[sName];
		};
		CUniFill.prototype.addAlpha = function(dValue) {
			this.setTransparent(Math.max(0, Math.min(255, (dValue * 255 + 0.5) >> 0)));

			// let oMod = new CColorMod();
			// oMod.name = "alpha";
			// oMod.val = nPctValue;
			// this.addColorMod(oMod);
		};
		CUniFill.prototype.isBlipFill = function() {
			if(this.fill && this.fill.type === c_oAscFill.FILL_TYPE_BLIP) {
				return true;
			}
		};
		CUniFill.prototype.isGradientFill = function() {
			if(this.fill && this.fill.type === c_oAscFill.FILL_TYPE_GRAD) {
				return true;
			}
		};
		CUniFill.prototype.checkTransparent = function() {
			let oFill = this.fill;
			if(oFill) {
				switch (oFill.type) {
					case c_oAscFill.FILL_TYPE_BLIP: {
						let aEffects = oFill.Effects;
						for(let nEffect = 0; nEffect < aEffects.length; ++nEffect) {
							let oEffect = aEffects[nEffect];
							if(oEffect instanceof AscFormat.CAlphaModFix && AscFormat.isRealNumber(oEffect.amt)) {
								this.setTransparent(255 * oEffect.amt / 100000);
							}
						}
						break;
					}
					case c_oAscFill.FILL_TYPE_SOLID: {
						let oColor = oFill.color;
						if(oColor) {
							let fAlphaVal = oColor.getModValue("alpha");
							if(fAlphaVal !== null) {
								this.setTransparent(255 * fAlphaVal / 100000);
								let aMods = oColor.Mods && oColor.Mods.Mods;
								if(Array.isArray(aMods)) {
									for(let nMod = aMods.length -1; nMod > -1; nMod--) {
										let oMod = aMods[nMod];
										if(oMod && oMod.name === "alpha") {
											aMods.splice(nMod, 1);
										}
									}
								}
							}
						}
						break;
					}
					case c_oAscFill.FILL_TYPE_PATT: {
						if(oFill.fgClr && oFill.bgClr) {
							let fAlphaVal = oFill.fgClr.getModValue("alpha");
							if(fAlphaVal !== null) {
								if(fAlphaVal === oFill.bgClr.getModValue("alpha")) {
									this.setTransparent(255 * fAlphaVal / 100000)
								}
							}
						}
						break;
					}
				}
			}
		};
		CUniFill.prototype.checkRasterImageId = function() {
			if (this.fill && typeof this.fill.RasterImageId === "string" && this.fill.RasterImageId.length > 0)
				return this.fill.RasterImageId;
			return null;
		}

		function CBuBlip() {
			CBaseNoIdObject.call(this);
			this.blip = null;
		}

		InitClass(CBuBlip, CBaseNoIdObject, 0);
		CBuBlip.prototype.setBlip = function (oPr) {
			this.blip = oPr;
		};
		CBuBlip.prototype.fillObject = function (oCopy, oIdMap) {
			if (this.blip) {
				oCopy.setBlip(this.blip.createDuplicate(oIdMap));
			}
		};
		CBuBlip.prototype.createDuplicate = function () {
			var oCopy = new CBuBlip();
			this.fillObject(oCopy, {});
			return oCopy;
		};
		CBuBlip.prototype.getChildren = function () {
			return [this.blip];
		};
		CBuBlip.prototype.isEqual = function (oBlip) {
			return this.blip.isEqual(oBlip.blip);
		};
		CBuBlip.prototype.toPPTY = function (pWriter) {
			var _src = this.blip.fill.RasterImageId;
			var imageLocal = AscCommon.g_oDocumentUrls.getImageLocal(_src);
			if (imageLocal)
				_src = imageLocal;

			pWriter.image_map[_src] = true;

			_src = pWriter.prepareRasterImageIdForWrite(_src);
			pWriter.WriteBlip(this.blip.fill, _src);
		};
		CBuBlip.prototype.fromPPTY = function (pReader, oParagraph, oBullet) {
			this.setBlip(new AscFormat.CUniFill());
			this.blip.setFill(new AscFormat.CBlipFill());
			pReader.ReadBlip(this.blip, undefined, undefined, undefined, oParagraph, oBullet);
		};
		CBuBlip.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				this.blip = new CUniFill();
				this.blip.Read_FromBinary(r);
			}
		};
		CBuBlip.prototype.Write_ToBinary = function (w) {
			w.WriteBool(isRealObject(this.blip));
			if (isRealObject(this.blip)) {
				this.blip.Write_ToBinary(w);
			}
		};
		CBuBlip.prototype.compare = function (compareObj) {
			var ret = null;
			if (compareObj instanceof CBuBlip) {
				ret = new CBuBlip();
				if (this.blip) {
					ret.blip = CompareUniFill(this.blip, compareObj.blip);
				}
			}
			return ret;
		};

		function CompareUniFill(unifill_1, unifill_2) {

			if (unifill_1 == null || unifill_2 == null) {
				return null;
			}

			var _ret = new CUniFill();

			if (!(unifill_1.transparent === null || unifill_2.transparent === null
				|| unifill_1.transparent !== unifill_2.transparent)) {
				_ret.transparent = unifill_1.transparent;
			}

			if (unifill_1.fill == null || unifill_2.fill == null
				|| unifill_1.fill.type != unifill_2.fill.type) {
				return _ret;
			}
			_ret.fill = unifill_1.compare(unifill_2);
			return _ret;
		}

		function CompareBlipTiles(tile1, tile2) {
			if (isRealObject(tile1)) {
				return tile1.IsIdentical(tile2);
			}
			return tile1 === tile2;
		}

		function CompareUnifillBool(u1, u2) {
			if (!u1 && !u2)
				return true;
			if (!u1 && u2 || u1 && !u2)
				return false;

			if (isRealNumber(u1.transparent) !== isRealNumber(u2.transparent)
				|| isRealNumber(u1.transparent) && u1.transparent !== u2.transparent) {
				return false;
			}
			if (!u1.fill && !u2.fill)
				return true;
			if (!u1.fill && u2.fill || u1.fill && !u2.fill)
				return false;

			if (u1.fill.type !== u2.fill.type)
				return false;
			switch (u1.fill.type) {
				case c_oAscFill.FILL_TYPE_BLIP: {
					if (u1.fill.RasterImageId && !u2.fill.RasterImageId || u2.fill.RasterImageId && !u1.fill.RasterImageId)
						return false;

					if (typeof u1.fill.RasterImageId === "string" && typeof u2.fill.RasterImageId === "string"
						&& AscCommon.getFullImageSrc2(u1.fill.RasterImageId) !== AscCommon.getFullImageSrc2(u2.fill.RasterImageId))
						return false;


					if (u1.fill.srcRect && !u2.fill.srcRect || !u1.fill.srcRect && u2.fill.srcRect)
						return false;

					if (u1.fill.srcRect && u2.fill.srcRect) {
						if (u1.fill.srcRect.l !== u2.fill.srcRect.l ||
							u1.fill.srcRect.t !== u2.fill.srcRect.t ||
							u1.fill.srcRect.r !== u2.fill.srcRect.r ||
							u1.fill.srcRect.b !== u2.fill.srcRect.b)
							return false;
					}

					if (u1.fill.stretch !== u2.fill.stretch || !CompareBlipTiles(u1.fill.tile, u2.fill.tile) || u1.fill.rotWithShape !== u2.fill.rotWithShape)
						return false;
					break;
				}
				case c_oAscFill.FILL_TYPE_SOLID: {
					if (u1.fill.color && u2.fill.color) {
						return CompareUniColor(u1.fill.color, u2.fill.color)
					}
					break;
				}
				case c_oAscFill.FILL_TYPE_GRAD: {
					if (u1.fill.colors.length !== u2.fill.colors.length)
						return false;
					if (isRealObject(u1.fill.path) !== isRealObject(u2.fill.path))
						return false;
					if (u1.fill.path && !u1.fill.path.IsIdentical(u2.fill.path))
						return false;

					if (isRealObject(u1.fill.lin) !== isRealObject(u2.fill.lin))
						return false;

					if (u1.fill.lin && !u1.fill.lin.IsIdentical(u2.fill.lin))
						return false;

					for (var i = 0; i < u1.fill.colors.length; ++i) {
						if (u1.fill.colors[i].pos !== u2.fill.colors[i].pos
							|| !CompareUniColor(u1.fill.colors[i].color, u2.fill.colors[i].color))
							return false;
					}
					break;
				}
				case c_oAscFill.FILL_TYPE_PATT: {
					if (u1.fill.ftype !== u2.fill.ftype
						|| !CompareUniColor(u1.fill.fgClr, u2.fill.fgClr)
						|| !CompareUniColor(u1.fill.bgClr, u2.fill.bgClr))
						return false;
					break;
				}
			}
			return true;
		}

		function CompareUniColor(u1, u2) {
			if (!u1 && !u2)
				return true;
			if (!u1 && u2 || u1 && !u2)
				return false;
			if (!u1.color && u2.color || u1.color && !u2.color)
				return false;
			if (u1.color && u2.color) {
				if (u1.color.type !== u2.color.type)
					return false;
				switch (u1.color.type) {
					case c_oAscColor.COLOR_TYPE_NONE: {
						break;
					}
					case c_oAscColor.COLOR_TYPE_SRGB: {
						if (u1.color.RGBA.R !== u2.color.RGBA.R
							|| u1.color.RGBA.G !== u2.color.RGBA.G
							|| u1.color.RGBA.B !== u2.color.RGBA.B
							|| u1.color.RGBA.A !== u2.color.RGBA.A) {
							return false;
						}
						break;
					}
					case c_oAscColor.COLOR_TYPE_PRST:
					case c_oAscColor.COLOR_TYPE_SCHEME: {
						if (u1.color.id !== u2.color.id)
							return false;
						break;
					}
					case c_oAscColor.COLOR_TYPE_SYS: {
						if (u1.color.RGBA.R !== u2.color.RGBA.R
							|| u1.color.RGBA.G !== u2.color.RGBA.G
							|| u1.color.RGBA.B !== u2.color.RGBA.B
							|| u1.color.RGBA.A !== u2.color.RGBA.A
							|| u1.color.id !== u2.color.id) {
							return false;
						}
						break;
					}
					case c_oAscColor.COLOR_TYPE_STYLE: {
						if (u1.bAuto !== u2.bAuto || u1.val !== u2.val) {
							return false;
						}
						break;
					}
				}
			}
			if (!u1.Mods && u2.Mods || !u2.Mods && u1.Mods)
				return false;

			if (u1.Mods && u2.Mods) {
				if (u1.Mods.Mods.length !== u2.Mods.Mods.length)
					return false;
				for (var i = 0; i < u1.Mods.Mods.length; ++i) {
					if (u1.Mods.Mods[i].name !== u2.Mods.Mods[i].name
						|| u1.Mods.Mods[i].val !== u2.Mods.Mods[i].val)
						return false;
				}
			}
			return true;
		}

// -----------------------------

		function CompareShapeProperties(shapeProp1, shapeProp2) {
			var _result_shape_prop = {};
			if (shapeProp1.type === shapeProp2.type) {
				_result_shape_prop.type = shapeProp1.type;
			} else {
				_result_shape_prop.type = null;
			}

			if (shapeProp1.h === shapeProp2.h) {
				_result_shape_prop.h = shapeProp1.h;
			} else {
				_result_shape_prop.h = null;
			}

			if (shapeProp1.w === shapeProp2.w) {
				_result_shape_prop.w = shapeProp1.w;
			} else {
				_result_shape_prop.w = null;
			}
			if (shapeProp1.x === shapeProp2.x) {
				_result_shape_prop.x = shapeProp1.x;
			} else {
				_result_shape_prop.x = null;
			}
			if (shapeProp1.y === shapeProp2.y) {
				_result_shape_prop.y = shapeProp1.y;
			} else {
				_result_shape_prop.y = null;
			}
			if (shapeProp1.rot === shapeProp2.rot) {
				_result_shape_prop.rot = shapeProp1.rot;
			} else {
				_result_shape_prop.rot = null;
			}
			if (shapeProp1.flipH === shapeProp2.flipH) {
				_result_shape_prop.flipH = shapeProp1.flipH;
			} else {
				_result_shape_prop.flipH = null;
			}
			if (shapeProp1.flipV === shapeProp2.flipV) {
				_result_shape_prop.flipV = shapeProp1.flipV;
			} else {
				_result_shape_prop.flipV = null;
			}

			if (shapeProp1.anchor === shapeProp2.anchor) {
				_result_shape_prop.anchor = shapeProp1.anchor;
			} else {
				_result_shape_prop.anchor = null;
			}

			if (shapeProp1.stroke == null || shapeProp2.stroke == null) {
				_result_shape_prop.stroke = null;
			} else {
				_result_shape_prop.stroke = shapeProp1.stroke.compare(shapeProp2.stroke)
			}

			/* if(shapeProp1.verticalTextAlign === shapeProp2.verticalTextAlign)
     {
     _result_shape_prop.verticalTextAlign = shapeProp1.verticalTextAlign;
     }
     else */
			{
				_result_shape_prop.verticalTextAlign = null;
				_result_shape_prop.vert = null;
			}

			if (shapeProp1.canChangeArrows !== true || shapeProp2.canChangeArrows !== true)
				_result_shape_prop.canChangeArrows = false;
			else
				_result_shape_prop.canChangeArrows = true;

			_result_shape_prop.fill = CompareUniFill(shapeProp1.fill, shapeProp2.fill);
			_result_shape_prop.IsLocked = shapeProp1.IsLocked === true || shapeProp2.IsLocked === true;
			if (isRealObject(shapeProp1.paddings) && isRealObject(shapeProp2.paddings)) {
				_result_shape_prop.paddings = new Asc.asc_CPaddings();
				_result_shape_prop.paddings.Left = isRealNumber(shapeProp1.paddings.Left) ? (shapeProp1.paddings.Left === shapeProp2.paddings.Left ? shapeProp1.paddings.Left : undefined) : undefined;
				_result_shape_prop.paddings.Top = isRealNumber(shapeProp1.paddings.Top) ? (shapeProp1.paddings.Top === shapeProp2.paddings.Top ? shapeProp1.paddings.Top : undefined) : undefined;
				_result_shape_prop.paddings.Right = isRealNumber(shapeProp1.paddings.Right) ? (shapeProp1.paddings.Right === shapeProp2.paddings.Right ? shapeProp1.paddings.Right : undefined) : undefined;
				_result_shape_prop.paddings.Bottom = isRealNumber(shapeProp1.paddings.Bottom) ? (shapeProp1.paddings.Bottom === shapeProp2.paddings.Bottom ? shapeProp1.paddings.Bottom : undefined) : undefined;
			}
			_result_shape_prop.canFill = shapeProp1.canFill === true || shapeProp2.canFill === true;

			if (shapeProp1.bFromChart || shapeProp2.bFromChart) {
				_result_shape_prop.bFromChart = true;
			} else {
				_result_shape_prop.bFromChart = false;
			}
			if (shapeProp1.bFromSmartArt || shapeProp2.bFromSmartArt) {
				_result_shape_prop.bFromSmartArt = true;
			} else {
				_result_shape_prop.bFromSmartArt = false;
			}
			if (shapeProp1.bFromSmartArtInternal || shapeProp2.bFromSmartArtInternal) {
				_result_shape_prop.bFromSmartArtInternal = true;
			} else {
				_result_shape_prop.bFromSmartArtInternal = false;
			}
			if (shapeProp1.bFromGroup || shapeProp2.bFromGroup) {
				_result_shape_prop.bFromGroup = true;
			} else {
				_result_shape_prop.bFromGroup = false;
			}
			if (!shapeProp1.bFromImage || !shapeProp2.bFromImage) {
				_result_shape_prop.bFromImage = false;
			} else {
				_result_shape_prop.bFromImage = true;
			}
			if (shapeProp1.locked || shapeProp2.locked) {
				_result_shape_prop.locked = true;
			}
			_result_shape_prop.lockAspect = !!(shapeProp1.lockAspect && shapeProp2.lockAspect);
			_result_shape_prop.textArtProperties = CompareTextArtProperties(shapeProp1.textArtProperties, shapeProp2.textArtProperties);
			if (shapeProp1.bFromSmartArtInternal && !shapeProp2.bFromSmartArtInternal || !shapeProp1.bFromSmartArtInternal && shapeProp2.bFromSmartArtInternal) {
				_result_shape_prop.textArtProperties = null;
			}
			if (shapeProp1.title === shapeProp2.title) {
				_result_shape_prop.title = shapeProp1.title;
			}
			if (shapeProp1.description === shapeProp2.description) {
				_result_shape_prop.description = shapeProp1.description;
			}
			if (shapeProp1.name === shapeProp2.name) {
				_result_shape_prop.name = shapeProp1.name;
			}
			if (shapeProp1.columnNumber === shapeProp2.columnNumber) {
				_result_shape_prop.columnNumber = shapeProp1.columnNumber;
			}
			if (shapeProp1.columnSpace === shapeProp2.columnSpace) {
				_result_shape_prop.columnSpace = shapeProp1.columnSpace;
			}
			if (shapeProp1.textFitType === shapeProp2.textFitType) {
				_result_shape_prop.textFitType = shapeProp1.textFitType;
			}

			if (shapeProp1.vertOverflowType === shapeProp2.vertOverflowType) {
				_result_shape_prop.vertOverflowType = shapeProp1.vertOverflowType;
			}

			if (!shapeProp1.shadow && !shapeProp2.shadow) {
				_result_shape_prop.shadow = null;
			} else if (shapeProp1.shadow && !shapeProp2.shadow) {
				_result_shape_prop.shadow = null;
			} else if (!shapeProp1.shadow && shapeProp2.shadow) {
				_result_shape_prop.shadow = null;
			} else if (shapeProp1.shadow.IsIdentical(shapeProp2.shadow)) {
				_result_shape_prop.shadow = shapeProp1.shadow.createDuplicate();
			} else {
				_result_shape_prop.shadow = null;
			}

			_result_shape_prop.protectionLockText = CompareProtectionFlags(shapeProp1.protectionLockText, shapeProp2.protectionLockText);
			_result_shape_prop.protectionLocked = CompareProtectionFlags(shapeProp1.protectionLocked, shapeProp2.protectionLocked);
			_result_shape_prop.protectionPrint = CompareProtectionFlags(shapeProp1.protectionPrint, shapeProp2.protectionPrint);

			return _result_shape_prop;
		}

		function CompareProtectionFlags(bFlag1, bFlag2) {
			if (bFlag1 === null || bFlag2 === null) {
				return null;
			} else if (bFlag1 === bFlag2) {
				return bFlag1;
			}
			return undefined;
		}

		function CompareTextArtProperties(oProps1, oProps2) {
			if (!oProps1 || !oProps2)
				return null;
			var oRet = {Fill: undefined, Line: undefined, Form: undefined};
			if (oProps1.Form === oProps2.Form) {
				oRet.From = oProps1.Form;
			}
			if (oProps1.Fill && oProps2.Fill) {
				oRet.Fill = CompareUniFill(oProps1.Fill, oProps2.Fill);
			}
			if (oProps1.Line && oProps2.Line) {
				oRet.Line = oProps1.Line.compare(oProps2.Line);
			}
			return oRet;
		}

// LN --------------------------
// размеры стрелок;
		var lg = 500, mid = 300, sm = 200;
//типы стрелок
		var ar_arrow = 0, ar_diamond = 1, ar_none = 2, ar_oval = 3, ar_stealth = 4, ar_triangle = 5;

		var LineEndType = {
			None:				0,
			Arrow:				1,
			Diamond:			2,
			Oval:				3,
			Stealth:			4,
			Triangle:			5,
			ReverseArrow:		6,
			ReverseTriangle:	7,
			Butt:				8,
			Square:				9,
			Slash:				10,

			vsdxNone: 11,
			vsdxOpen90Arrow: 12,
			vsdxFilled90Arrow: 13,
			vsdxOpenSharpArrow: 14,
			vsdxFilledSharpArrow: 15,
			vsdxIndentedFilledArrow: 16,
			vsdxOutdentedFilledArrow: 17,
			vsdxOpenFletch: 18,
			vsdxFilledFletch: 19,
			vsdxDimensionLine: 20,
			vsdxFilledDot: 21,
			vsdxFilledSquare: 22,
			vsdxOpenASMEArrow: 23,
			vsdxFilledASMEArrow: 24,
			vsdxClosedASMEArrow: 25,
			vsdxClosed90Arrow: 26,
			vsdxClosedSharpArrow: 27,
			vsdxIndentedClosedArrow: 28,
			vsdxOutdentedClosedArrow: 29,
			vsdxClosedFletch: 30,
			vsdxClosedDot: 31,
			vsdxClosedSquare: 32,
			vsdxDiamond: 33,
			vsdxBackslash: 34,
			vsdxOpenOneDash: 35,
			vsdxOpenTwoDash: 36,
			vsdxOpenThreeDash: 37,
			vsdxFork: 38,
			vsdxDashFork: 39,
			vsdxClosedFork: 40,
			vsdxClosedPlus: 41,
			vsdxClosedOneDash: 42,
			vsdxClosedTwoDash: 43,
			vsdxClosedThreeDash: 44,
			vsdxClosedDiamond: 45,
			vsdxFilledOneDash: 46,
			vsdxFilledTwoDash: 47,
			vsdxFilledThreeDash: 48,
			vsdxFilledDiamond: 49,
			vsdxFilledDoubleArrow: 50,
			vsdxClosedDoubleArrow: 51,
			vsdxClosedNoDash: 52,
			vsdxFilledNoDash: 53,
			vsdxOpenDoubleArrow: 54,
			vsdxOpenArrowSingleDash: 55,
			vsdxOpenDoubleArrowSingleDash: 56
		};
		var LineEndSize = {
			Large: 0,
			Mid: 1,
			Small: 2,

			vsdxVerySmall : 3,
			vsdxSmall: 		4,
			vsdxMedium: 	5,
			vsdxLarge:		6,
			vsdxExtraLarge:	7,
			vsdxJumbo:		8,
			vsdxColossal:	9
		};

		var LineJoinType = {
			Empty: 0,
			Round: 1,
			Bevel: 2,
			Miter: 3
		};
		let lineWidthInfluenceCoef = 0.028;

		function EndArrow() {
			CBaseNoIdObject.call(this);
			this.type = null;
			this.len = null;
			this.w = null;

		}

		InitClass(EndArrow, CBaseNoIdObject, 0);
		EndArrow.prototype.compare = function (end_arrow) {
			if (end_arrow == null) {
				return null;
			}
			var _ret = new EndArrow();
			if (this.type === end_arrow.type) {
				_ret.type = this.type;
			}
			if (this.len === end_arrow.len) {
				_ret.len = this.len;
			}
			if (this.w === end_arrow) {
				_ret.w = this.w;
			}
			return _ret;
		};
		EndArrow.prototype.createDuplicate = function () {
			var duplicate = new EndArrow();
			duplicate.type = this.type;
			duplicate.len = this.len;
			duplicate.w = this.w;
			return duplicate;
		};
		EndArrow.prototype.IsIdentical = function (arrow) {
			return arrow && arrow.type === this.type && arrow.len === this.len && arrow.w === this.w;
		};
		EndArrow.prototype.GetWidth = function (_size, _max) {
			var size = Math.max(_size, _max ? _max : 2);
			var _ret = 3 * size;
			let startSizeInch;
			let inchSize;

			if (null != this.w) {
				switch (this.w) {
					case LineEndSize.Large:
						_ret = 5 * size;
						break;
					case LineEndSize.Small:
						_ret = 2 * size;
						break;

					case LineEndSize.vsdxVerySmall:
						startSizeInch = 0.0391;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;
					case LineEndSize.vsdxSmall:
						startSizeInch = 0.0508;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;
					case LineEndSize.vsdxMedium:
						startSizeInch = 0.0693;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;
					case LineEndSize.vsdxLarge:
						startSizeInch = 0.0898;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;
					case LineEndSize.vsdxExtraLarge:
						startSizeInch = 0.1094;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;
					case LineEndSize.vsdxJumbo:
						startSizeInch = 0.2473;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;
					case LineEndSize.vsdxColossal:
						startSizeInch = 0.499;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;
					default:
						break;
				}
			}
			return _ret;
		};
		EndArrow.prototype.GetLen = function (_size, _max) {
			var size = Math.max(_size, _max ? _max : 2);
			var _ret = 3 * size;
			let startSizeInch;
			let inchSize;

			if (null != this.len) {
				switch (this.len) {
					case LineEndSize.Large:
						_ret = 5 * size;
						break;
					case LineEndSize.Small:
						_ret = 2 * size;
						break;

					case LineEndSize.vsdxVerySmall:
						startSizeInch = 0.0391;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;
					case LineEndSize.vsdxSmall:
						startSizeInch = 0.0508;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;
					case LineEndSize.vsdxMedium:
						startSizeInch = 0.0693;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;
					case LineEndSize.vsdxLarge:
						startSizeInch = 0.0898;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;
					case LineEndSize.vsdxExtraLarge:
						startSizeInch = 0.1094;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;
					case LineEndSize.vsdxJumbo:
						startSizeInch = 0.2473;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;
					case LineEndSize.vsdxColossal:
						startSizeInch = 0.499;
						inchSize = startSizeInch + lineWidthInfluenceCoef * _size *  AscCommonWord.g_dKoef_mm_to_pt;
						_ret = inchSize * AscCommonWord.g_dKoef_in_to_mm;
						break;

					default:
						break;
				}
			}
			return _ret;
		};
		EndArrow.prototype.setType = function (type) {
			this.type = type;
		};
		EndArrow.prototype.setLen = function (len) {
			this.len = len;
		};
		EndArrow.prototype.setW = function (w) {
			this.w = w;
		};
		EndArrow.prototype.Write_ToBinary = function (w) {
			writeLong(w, this.type);
			writeLong(w, this.len);
			writeLong(w, this.w);
		};
		EndArrow.prototype.Read_FromBinary = function (r) {
			this.type = readLong(r);
			this.len = readLong(r);
			this.w = readLong(r);
		};
		EndArrow.prototype.GetSizeCode = function (sVal) {

			switch (sVal) {
				case "lg": {
					return LineEndSize.Large;
				}
				case "med": {
					return LineEndSize.Mid;
				}
				case "sm": {
					return LineEndSize.Small;
				}

				// not used for visio end types. use types enum directly
			}
			return LineEndSize.Mid;
		};

		EndArrow.prototype.GetSizeByCode = function (nCode) {

			switch (nCode) {
				case LineEndSize.Large: {
					return "lg";
				}
				case LineEndSize.Mid: {
					return "med";
				}
				case LineEndSize.Small: {
					return "sm";
				}

				// not used for visio end types. use types enum directly
			}
			return "med";
		};

		EndArrow.prototype.GetTypeCode = function (sVal) {
			switch (sVal) {
				case "arrow": {
					return LineEndType.Arrow;
				}
				case "diamond": {
					return LineEndType.Diamond;
				}
				case "none": {
					return LineEndType.None;
				}
				case "oval": {
					return LineEndType.Oval;
				}
				case "stealth": {
					return LineEndType.Stealth;
				}
				case "triangle": {
					return LineEndType.Triangle;
				}

				// not used for visio end types. use types enum directly
			}
			return LineEndType.Arrow;
		};


		EndArrow.prototype.GetTypeByCode = function (nCode) {
			switch (nCode) {
				case LineEndType.Arrow : {
					return "arrow";
				}
				case LineEndType.Diamond: {
					return "diamond";
				}
				case LineEndType.None: {
					return "none";
				}
				case LineEndType.Oval: {
					return "oval";
				}
				case LineEndType.Stealth: {
					return "stealth";
				}
				case LineEndType.Triangle: {
					return "triangle";
				}

				// not used for visio end types. use types enum directly
			}
			return "arrow";
		};


		function ConvertJoinAggType(_type) {
			switch (_type) {
				case LineJoinType.Round:
					return 2;
				case LineJoinType.Bevel:
					return 1;
				case LineJoinType.Miter:
					return 0;
				default:
					break;
			}
			return 2;
		}

		function LineJoin(type) {
			CBaseNoIdObject.call(this);
			this.type = AscFormat.isRealNumber(type) ? type : null;
			this.limit = null;
		}

		InitClass(LineJoin, CBaseNoIdObject, 0);
		LineJoin.prototype.IsIdentical = function (oJoin) {
			if (!oJoin)
				return false;
			if (this.type !== oJoin.type) {
				return false;
			}
			if (this.limit !== oJoin.limit)
				return false;
			return true;
		};
		LineJoin.prototype.createDuplicate = function () {
			var duplicate = new LineJoin();
			duplicate.type = this.type;
			duplicate.limit = this.limit;
			return duplicate;
		};
		LineJoin.prototype.setType = function (type) {
			this.type = type;
		};
		LineJoin.prototype.setLimit = function (limit) {
			this.limit = limit;
		};
		LineJoin.prototype.Write_ToBinary = function (w) {
			writeLong(w, this.type);
			writeLong(w, this.limit);
		};
		LineJoin.prototype.Read_FromBinary = function (r) {
			this.type = readLong(r);
			this.limit = readLong(r);
		};

		function CLn() {
			CBaseNoIdObject.call(this);
			this.Fill = null;//new CUniFill();
			this.prstDash = null;

			this.Join = null;
			this.headEnd = null;
			this.tailEnd = null;

			this.algn = null;
			this.cap = null;
			this.cmpd = null;
			this.w = null;
		}

		InitClass(CLn, CBaseNoIdObject, 0);
		CLn.prototype.compare = function (line) {
			if (line == null) {
				return null;
			}
			var _ret = new CLn();
			if (this.Fill != null) {
				_ret.Fill = CompareUniFill(this.Fill, line.Fill);
			}
			if (this.prstDash === line.prstDash) {
				_ret.prstDash = this.prstDash;
			} else {
				_ret.prstDash = undefined;
			}
			if (this.Join === line.Join) {
				_ret.Join = this.Join;
			}
			if (this.tailEnd != null) {
				_ret.tailEnd = this.tailEnd.compare(line.tailEnd);
			}
			if (this.headEnd != null) {
				_ret.headEnd = this.headEnd.compare(line.headEnd);
			}

			if (this.algn === line.algn) {
				_ret.algn = this.algn;
			}
			if (this.cap === line.cap) {
				_ret.cap = this.cap;
			}
			if (this.cmpd === line.cmpd) {
				_ret.cmpd = this.cmpd;
			}
			if (this.w === line.w) {
				_ret.w = this.w;
			}
			return _ret;
		};
		CLn.prototype.merge = function (ln) {
			if (ln == null) {
				return;
			}
			if (ln.Fill != null && ln.Fill.fill != null) {
				this.setFill(ln.Fill.createDuplicate());
			}

			if (ln.prstDash != null) {
				this.setPrstDash(ln.prstDash);
			}

			if (ln.Join != null) {
				this.setJoin(ln.Join.createDuplicate());
			}

			if (ln.headEnd != null) {
				this.setHeadEnd(ln.headEnd.createDuplicate());
			}

			if (ln.tailEnd != null) {
				this.setTailEnd(ln.tailEnd.createDuplicate());
			}

			if (ln.algn != null) {
				this.setAlgn(ln.algn);
			}

			if (ln.cap != null) {
				this.setCap(ln.cap);
			}

			if (ln.cmpd != null) {
				this.setCmpd(ln.cmpd);
			}

			if (ln.w != null) {
				this.setW(ln.w);
			}
			else if(ln.isNoFillLine && ln.isNoFillLine()) {
				this.setW(0);
			}
		};
		CLn.prototype.calculate = function (theme, slide, layout, master, RGBA, colorMap) {
			if (isRealObject(this.Fill)) {
				this.Fill.calculate(theme, slide, layout, master, RGBA, colorMap);
			}
		};
		CLn.prototype.createDuplicate = function (bSaveFormatting) {
			var duplicate = new CLn();

			if (null != this.Fill) {
				if (bSaveFormatting === true) {
					duplicate.Fill = this.Fill.saveSourceFormatting();
				} else {
					duplicate.Fill = this.Fill.createDuplicate();
				}

			}

			duplicate.prstDash = this.prstDash;
			duplicate.Join = this.Join;
			if (this.headEnd != null) {
				duplicate.headEnd = this.headEnd.createDuplicate();
			}

			if (this.tailEnd != null) {
				duplicate.tailEnd = this.tailEnd.createDuplicate();
			}

			duplicate.algn = this.algn;
			duplicate.cap = this.cap;
			duplicate.cmpd = this.cmpd;
			duplicate.w = this.w;
			return duplicate;
		};
		CLn.prototype.IsIdentical = function (ln) {
			return ln && (this.Fill == null ? ln.Fill == null : this.Fill.IsIdentical(ln.Fill)) && (this.Join == null ? ln.Join == null : this.Join.IsIdentical(ln.Join))
				&& (this.headEnd == null ? ln.headEnd == null : this.headEnd.IsIdentical(ln.headEnd))
				&& (this.tailEnd == null ? ln.tailEnd == null : this.tailEnd.IsIdentical(ln.tailEnd)) &&
				this.algn == ln.algn && this.cap == ln.cap && this.cmpd == ln.cmpd && this.w == ln.w && this.prstDash === ln.prstDash;
		};
		CLn.prototype.isEqual = function (ln) {
			return this.IsIdentical(ln);
		};
		CLn.prototype.setFill = function (fill) {
			this.Fill = fill;
		};
		CLn.prototype.setPrstDash = function (prstDash) {
			this.prstDash = prstDash;
		};
		CLn.prototype.setJoin = function (join) {
			this.Join = join;
		};
		CLn.prototype.setHeadEnd = function (headEnd) {
			this.headEnd = headEnd;
		};
		CLn.prototype.setTailEnd = function (tailEnd) {
			this.tailEnd = tailEnd;
		};
		CLn.prototype.setAlgn = function (algn) {
			this.algn = algn;
		};
		CLn.prototype.setCap = function (cap) {
			this.cap = cap;
		};
		CLn.prototype.setCmpd = function (cmpd) {
			this.cmpd = cmpd;
		};
		CLn.prototype.setW = function (w) {
			this.w = w;
		};
		CLn.prototype.isVisible = function () {
			return this.Fill && this.Fill.isVisible();
		};
		CLn.prototype.Write_ToBinary = function (w) {
			w.WriteBool(isRealObject(this.Fill));
			if (isRealObject(this.Fill)) {
				this.Fill.Write_ToBinary(w);
			}
			writeLong(w, this.prstDash);

			w.WriteBool(isRealObject(this.Join));
			if (isRealObject(this.Join)) {
				this.Join.Write_ToBinary(w);
			}

			w.WriteBool(isRealObject(this.headEnd));
			if (isRealObject(this.headEnd)) {
				this.headEnd.Write_ToBinary(w);
			}
			w.WriteBool(isRealObject(this.tailEnd));
			if (isRealObject(this.tailEnd)) {
				this.tailEnd.Write_ToBinary(w);
			}
			writeLong(w, this.algn);
			writeLong(w, this.cap);
			writeLong(w, this.cmpd);
			writeLong(w, this.w);
		};
		CLn.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				this.Fill = new CUniFill();
				this.Fill.Read_FromBinary(r);
			} else {
				this.Fill = null;
			}
			this.prstDash = readLong(r);

			if (r.GetBool()) {
				this.Join = new LineJoin();
				this.Join.Read_FromBinary(r);
			}

			if (r.GetBool()) {
				this.headEnd = new EndArrow();
				this.headEnd.Read_FromBinary(r);
			}
			if (r.GetBool()) {
				this.tailEnd = new EndArrow();
				this.tailEnd.Read_FromBinary(r);
			}
			this.algn = readLong(r);
			this.cap = readLong(r);
			this.cmpd = readLong(r);
			this.w = readLong(r);
		};
		CLn.prototype.isNoFillLine = function () {
			if (this.Fill) {
				return this.Fill.isNoFill();
			}
			return false;
		};
		CLn.prototype.GetCapCode = function (sVal) {
			switch (sVal) {
				case "flat": {
					return 0;
				}
				case "rnd": {
					return 1;
				}
				case "sq": {
					return 2;
				}
			}
			return 0;
		};
		CLn.prototype.GetCapByCode = function (nCode) {
			switch (nCode) {
				case 0: {
					return "flat";
				}
				case 1: {
					return "rnd";
				}
				case 2: {
					return "sq";
				}
			}
			return null;
		};
		CLn.prototype.GetAlgnCode = function (sVal) {
			switch (sVal) {
				case "ctr": {
					return 0;
				}
				case "in": {
					return 1;
				}
			}
			return 0;
		};
		CLn.prototype.GetAlgnByCode = function (sVal) {
			switch (sVal) {
				case 0: {
					return "ctr";
				}
				case 1: {
					return "in";
				}
			}
			return null;
		};
		CLn.prototype.GetCmpdCode = function (sVal) {
			switch (sVal) {
				case "dbl": {
					return 0;
				}
				case "sng": {
					return 1;
				}
				case "thickThin": {
					return 2;
				}
				case "thinThick": {
					return 3;
				}
				case "tri": {
					return 4;
				}
			}
			return 1;
		};
		CLn.prototype.GetCmpdByCode = function (sVal) {
			switch (sVal) {
				case 0: {
					return "dbl";
				}
				case 1: {
					return "sng";
				}
				case 2: {
					return "thickThin";
				}
				case 3: {
					return "thinThick";
				}
				case 4: {
					return "tri";
				}
			}
			return null;
		};
		CLn.prototype.GetDashCode = function (sVal) {
			switch (sVal) {
				case "dash": {
					return 0;
				}
				case "dashDot": {
					return 1;
				}
				case "dot": {
					return 2;
				}
				case "lgDash": {
					return 3;
				}
				case "lgDashDot": {
					return 4;
				}
				case "lgDashDotDot": {
					return 5;
				}
				case "solid": {
					return 6;
				}
				case "sysDash": {
					return 7;
				}
				case "sysDashDot": {
					return 8;
				}
				case "sysDashDotDot": {
					return 9;
				}
				case "sysDot": {
					return 10;
				}
				case "vsdxTransparent": { return 11; }
				case "vsdxSolid": { return 12; }
				case "vsdxDash": { return 13; }
				case "vsdxDot": { return 14; }
				case "vsdxDashDot": { return 15; }
				case "vsdxDashDotDot": { return 16; }
				case "vsdxDashDashDot": { return 17; }
				case "vsdxLongDashShortDash": { return 18; }
				case "vsdxLongDashShortDashShortDash": { return 19; }
				case "vsdxHalfDash": { return 20; }
				case "vsdxHalfDot": { return 21; }
				case "vsdxHalfDashDot": { return 22; }
				case "vsdxHalfDashDotDot": { return 23; }
				case "vsdxHalfDashDashDot": { return 24; }
				case "vsdxHalfLongDashShortDash": { return 25; }
				case "vsdxHalfLongDashShortDashShortDash": { return 26; }
				case "vsdxDoubleDash": { return 27; }
				case "vsdxDoubleDot": { return 28; }
				case "vsdxDoubleDashDot": { return 29; }
				case "vsdxDoubleDashDotDot": { return 30; }
				case "vsdxDoubleDashDashDot": { return 31; }
				case "vsdxDoubleLongDashShortDash": { return 32; }
				case "vsdxDoubleLongDashShortDashShortDash": { return 33; }
				case "vsdxHalfHalfDash": { return 34; }
			}
			return 6;
		};
		CLn.prototype.GetDashByCode = function (sVal) {
			switch (sVal) {
				case 0: {
					return "dash";
				}
				case 1 : {
					return "dashDot";
				}
				case 2 : {
					return "dot";
				}
				case 3 : {
					return "lgDash";
				}
				case 4 : {
					return "lgDashDot";
				}
				case 5 : {
					return "lgDashDotDot";
				}
				case 6 : {
					return "solid";
				}
				case 7 : {
					return "sysDash";
				}
				case 8: {
					return "sysDashDot";
				}
				case 9 : {
					return "sysDashDotDot";
				}
				case 10 : {
					return "sysDot";
				}
				case 11: { return "vsdxTransparent"; }
				case 12: { return "vsdxSolid"; }
				case 13: { return "vsdxDash"; }
				case 14: { return "vsdxDot"; }
				case 15: { return "vsdxDashDot"; }
				case 16: { return "vsdxDashDotDot"; }
				case 17: { return "vsdxDashDashDot"; }
				case 18: { return "vsdxLongDashShortDash"; }
				case 19: { return "vsdxLongDashShortDashShortDash"; }
				case 20: { return "vsdxHalfDash"; }
				case 21: { return "vsdxHalfDot"; }
				case 22: { return "vsdxHalfDashDot"; }
				case 23: { return "vsdxHalfDashDotDot"; }
				case 24: { return "vsdxHalfDashDashDot"; }
				case 25: { return "vsdxHalfLongDashShortDash"; }
				case 26: { return "vsdxHalfLongDashShortDashShortDash"; }
				case 27: { return "vsdxDoubleDash"; }
				case 28: { return "vsdxDoubleDot"; }
				case 29: { return "vsdxDoubleDashDot"; }
				case 30: { return "vsdxDoubleDashDotDot"; }
				case 31: { return "vsdxDoubleDashDashDot"; }
				case 32: { return "vsdxDoubleLongDashShortDash"; }
				case 33: { return "vsdxDoubleLongDashShortDashShortDash"; }
				case 34: { return "vsdxHalfHalfDash"; }
			}
			return null;
		};
		CLn.prototype.fillDocumentBorder = function(oBorder) {
			if(this.Fill) {
				oBorder.Unifill = this.Fill;
			}
			oBorder.Size = (this.w === null) ? 12700 : ((this.w) >> 0);
			oBorder.Size /= 36000;
			oBorder.Value = AscCommonWord.border_Single;
		};
		CLn.prototype.fromDocumentBorder = function(oBorder) {
			this.Fill = oBorder.Unifill;
			this.w = null;
			if(AscFormat.isRealNumber(oBorder.Size)) {
				this.w = oBorder.Size * 36000 >> 0;
			}
			this.cmpd = 1;
		};
		CLn.prototype.getWidthMM = function () {
			const nEmu = AscFormat.isRealNumber(this.w) ? this.w : 12700;
			return nEmu * AscCommonWord.g_dKoef_emu_to_mm;
		};
// -----------------------------

// SHAPE ----------------------------

		function DefaultShapeDefinition() {
			CBaseFormatObject.call(this);
			this.spPr = new CSpPr();
			this.bodyPr = new CBodyPr();
			this.lstStyle = new TextListStyle();
			this.style = null;
		}

		InitClass(DefaultShapeDefinition, CBaseFormatObject, AscDFH.historyitem_type_DefaultShapeDefinition);
		DefaultShapeDefinition.prototype.setSpPr = function (spPr) {
			AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DefaultShapeDefinition_SetSpPr, this.spPr, spPr));
			this.spPr = spPr;
		};
		DefaultShapeDefinition.prototype.setBodyPr = function (bodyPr) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DefaultShapeDefinition_SetBodyPr, this.bodyPr, bodyPr));
			this.bodyPr = bodyPr;
		};
		DefaultShapeDefinition.prototype.setLstStyle = function (lstStyle) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DefaultShapeDefinition_SetLstStyle, this.lstStyle, lstStyle));
			this.lstStyle = lstStyle;
		};
		DefaultShapeDefinition.prototype.setStyle = function (style) {
			AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DefaultShapeDefinition_SetStyle, this.style, style));
			this.style = style;
		};
		DefaultShapeDefinition.prototype.createDuplicate = function () {
			var ret = new DefaultShapeDefinition();
			if (this.spPr) {
				ret.setSpPr(this.spPr.createDuplicate());
			}
			if (this.bodyPr) {
				ret.setBodyPr(this.bodyPr.createDuplicate());
			}
			if (this.lstStyle) {
				ret.setLstStyle(this.lstStyle.createDuplicate());
			}
			if (this.style) {
				ret.setStyle(this.style.createDuplicate());
			}
			return ret;
		};


		function CNvPr() {

			CBaseFormatObject.call(this);
			this.id = 0;
			this.name = "";
			this.isHidden = null;
			this.descr = null;
			this.title = null;

			this.hlinkClick = null;
			this.hlinkHover = null;

			this.form = null;

			this.setId(AscCommon.CreateDurableId());
		}

		InitClass(CNvPr, CBaseFormatObject, AscDFH.historyitem_type_CNvPr);
		CNvPr.prototype.createDuplicate = function () {
			var duplicate = new CNvPr();
			duplicate.setName(this.name);
			duplicate.setIsHidden(this.isHidden);
			duplicate.setDescr(this.descr);
			duplicate.setTitle(this.title);
			if (this.hlinkClick) {
				duplicate.setHlinkClick(this.hlinkClick.createDuplicate());
			}
			if (this.hlinkHover) {
				duplicate.setHlinkHover(this.hlinkHover.createDuplicate());
			}
			return duplicate;
		};
		CNvPr.prototype.setId = function (id) {
			AscCommon.History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_CNvPr_SetId, this.id, id));
			this.id = id;
		};
		CNvPr.prototype.setName = function (name) {
			AscCommon.History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_CNvPr_SetName, this.name, name));
			this.name = name;
		};
		CNvPr.prototype.setIsHidden = function (isHidden) {
			AscCommon.History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_CNvPr_SetIsHidden, this.isHidden, isHidden));
			this.isHidden = isHidden;
		};
		CNvPr.prototype.setDescr = function (descr) {
			AscCommon.History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_CNvPr_SetDescr, this.descr, descr));
			this.descr = descr;
		};
		CNvPr.prototype.setHlinkClick = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_CNvPr_SetHlinkClick, this.hlinkClick, pr));
			this.hlinkClick = pr;
		};
		CNvPr.prototype.setHlinkHover = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_CNvPr_SetHlinkHover, this.hlinkHover, pr));
			this.hlinkHover = pr;
		};
		CNvPr.prototype.setTitle = function (title) {
			AscCommon.History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_CNvPr_SetTitle, this.title, title));
			this.title = title;
		};
		CNvPr.prototype.setFromOther = function (oOther) {
			if (!oOther) {
				return;
			}
			if (oOther.name) {
				this.setName(oOther.name);
			}
			if (oOther.descr) {
				this.setDescr(oOther.descr);
			}
			if (oOther.title) {
				this.setTitle(oOther.title);
			}
		};
		CNvPr.prototype.hasSameNameAndId = function (oPr) {
			if (!oPr) {
				return false;
			}
			return this.id === oPr.id && this.name === oPr.name;
		};


		const AUDIO_CD = 0;
		const WAV_AUDIO_FILE = 1;
		const AUDIO_FILE = 2;
		const VIDEO_FILE = 3;
		const QUICK_TIME_FILE = 4;


		const DRAW_TYPE_PEN = 0;
		const DRAW_TYPE_PENCIL = 1;
		const DRAW_TYPE_HIGHLITER = 2;



		function UniMedia() {
			CBaseNoIdObject.call(this);
			this.type = null;
			this.media = null;
		}

		InitClass(UniMedia, CBaseNoIdObject, 0);
		UniMedia.prototype.Write_ToBinary = function (w) {
			var bType = this.type !== null && this.type !== undefined;
			var bMedia = typeof this.media === 'string';
			var nFlags = 0;
			bType && (nFlags |= 1);
			bMedia && (nFlags |= 2);
			w.WriteLong(nFlags);
			bType && w.WriteLong(this.type);
			bMedia && w.WriteString2(this.media);
		};
		UniMedia.prototype.Read_FromBinary = function (r) {
			var nFlags = r.GetLong();
			if (nFlags & 1) {
				this.type = r.GetLong();
			}
			if (nFlags & 2) {
				this.media = r.GetString2();
			}
		};
		UniMedia.prototype.createDuplicate = function () {
			var _ret = new UniMedia();
			_ret.type = this.type;
			_ret.media = this.media;
			return _ret;
		};

		drawingConstructorsMap[AscDFH.historyitem_NvPr_SetUniMedia] = UniMedia;

		function NvPr() {
			CBaseFormatObject.call(this);
			this.isPhoto = null;
			this.userDrawn = null;
			this.ph = null;
			this.unimedia = null;

			this.extLst = [];
		}

		InitClass(NvPr, CBaseFormatObject, AscDFH.historyitem_type_NvPr);
		NvPr.prototype.setIsPhoto = function (isPhoto) {
			AscCommon.History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_NvPr_SetIsPhoto, this.isPhoto, isPhoto));
			this.isPhoto = isPhoto;
		};
		NvPr.prototype.setUserDrawn = function (userDrawn) {
			AscCommon.History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_NvPr_SetUserDrawn, this.userDrawn, userDrawn));
			this.userDrawn = userDrawn;
		};
		NvPr.prototype.setPh = function (ph) {
			AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_NvPr_SetPh, this.ph, ph));
			this.ph = ph;
		};
		NvPr.prototype.addExt = function (ext) {
			History.Add(new CChangesDrawingsContentNoId(this, AscDFH.historyitem_NvPr_AddExt, this.extLst.length, [ext], true));
			this.extLst.push(ext);
		};
		NvPr.prototype.setUniMedia = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_NvPr_SetUniMedia, this.unimedia, pr));
			this.unimedia = pr;
		};
		NvPr.prototype.createDuplicate = function () {
			var duplicate = new NvPr();
			duplicate.setIsPhoto(this.isPhoto);
			duplicate.setUserDrawn(this.userDrawn);
			if (this.ph != null) {
				duplicate.setPh(this.ph.createDuplicate());
			}
			if (this.unimedia != null) {
				duplicate.setUniMedia(this.unimedia.createDuplicate());
			}
			return duplicate;
		};

		var szPh_full = 0,
			szPh_half = 1,
			szPh_quarter = 2;

		var orientPh_horz = 0,
			orientPh_vert = 1;


		function Ph() {
			CBaseFormatObject.call(this);
			this.hasCustomPrompt = null;
			this.idx = null;
			this.orient = null;
			this.sz = null;
			this.type = null;
		}

		InitClass(Ph, CBaseFormatObject, AscDFH.historyitem_type_Ph);
		Ph.prototype.createDuplicate = function () {
			var duplicate = new Ph();
			duplicate.setHasCustomPrompt(this.hasCustomPrompt);
			duplicate.setIdx(this.idx);
			duplicate.setOrient(this.orient);
			duplicate.setSz(this.sz);
			duplicate.setType(this.type);
			return duplicate;
		};
		Ph.prototype.setHasCustomPrompt = function (hasCustomPrompt) {
			AscCommon.History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Ph_SetHasCustomPrompt, this.hasCustomPrompt, hasCustomPrompt));
			this.hasCustomPrompt = hasCustomPrompt;
		};
		Ph.prototype.setIdx = function (idx) {
			AscCommon.History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Ph_SetIdx, this.idx, idx));
			this.idx = idx;
		};
		Ph.prototype.getIdx = function() {
			return this.idx;
		}
		Ph.prototype.setOrient = function (orient) {
			AscCommon.History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Ph_SetOrient, this.orient, orient));
			this.orient = orient;
		};
		Ph.prototype.setSz = function (sz) {
			AscCommon.History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Ph_SetSz, this.sz, sz));
			this.sz = sz;
		};
		Ph.prototype.setType = function (type) {
			AscCommon.History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Ph_SetType, this.type, type));
			this.type = type;
		};
		Ph.prototype.getType = function() {
			return this.type;
		};

		function fUpdateLocksValue(nLocks, nMask, bValue) {
			nLocks |= nMask;
			if (bValue) {
				nLocks |= (nMask << 1)
			} else {
				nLocks &= ~(nMask << 1)
			}
			return nLocks;
		}

		function fGetLockValue(nLocks, nMask) {
			if (nLocks & nMask) {
				return !!(nLocks & (nMask << 1));
			}
			return undefined;
		}

		window['AscFormat'] = window['AscFormat'] || {};
		window['AscFormat'].fUpdateLocksValue = fUpdateLocksValue;
		window['AscFormat'].fGetLockValue = fGetLockValue;

		function CNvUniSpPr() {
			CBaseNoIdObject.call(this);
			this.locks = null;

			this.stCnxIdx = null;
			this.stCnxId = null;

			this.endCnxIdx = null;
			this.endCnxId = null;
		}

		InitClass(CNvUniSpPr, CBaseNoIdObject, 0);
		CNvUniSpPr.prototype.Write_ToBinary = function (w) {
			if (AscFormat.isRealNumber(this.locks)) {
				w.WriteBool(true);
				w.WriteLong(this.locks);
			} else {
				w.WriteBool(false);
			}
			if (AscFormat.isRealNumber(this.stCnxIdx) && typeof (this.stCnxId) === "string" && this.stCnxId.length > 0) {
				w.WriteBool(true);
				w.WriteLong(this.stCnxIdx);
				w.WriteString2(this.stCnxId);
			} else {
				w.WriteBool(false);
			}
			if (AscFormat.isRealNumber(this.endCnxIdx) && typeof (this.endCnxId) === "string" && this.endCnxId.length > 0) {
				w.WriteBool(true);
				w.WriteLong(this.endCnxIdx);
				w.WriteString2(this.endCnxId);
			} else {
				w.WriteBool(false);
			}
		};
		CNvUniSpPr.prototype.Read_FromBinary = function (r) {
			var bCnx = r.GetBool();
			if (bCnx) {
				this.locks = r.GetLong();
			} else {
				this.locks = null;
			}
			bCnx = r.GetBool();
			if (bCnx) {
				this.stCnxIdx = r.GetLong();
				this.stCnxId = r.GetString2();
			} else {
				this.stCnxIdx = null;
				this.stCnxId = null;
			}
			bCnx = r.GetBool();
			if (bCnx) {
				this.endCnxIdx = r.GetLong();
				this.endCnxId = r.GetString2();
			} else {
				this.endCnxIdx = null;
				this.endCnxId = null;
			}
		};
		CNvUniSpPr.prototype.assignConnectors = function(aSpTree) {
			let bNeedSetStart = AscFormat.isRealNumber(this.stCnxIdFormat);
			let bNeedSetEnd = AscFormat.isRealNumber(this.endCnxIdFormat);
			if(bNeedSetStart || bNeedSetEnd) {
				for(let nSp = 0; nSp < aSpTree.length && (bNeedSetEnd || bNeedSetStart); ++nSp) {
					let oSp = aSpTree[nSp];
					if(bNeedSetStart && oSp.getFormatId() === this.stCnxIdFormat) {
						this.stCnxId = oSp.Get_Id();
						this.stCnxIdFormat = undefined;
						bNeedSetStart = false;
					}
					if(bNeedSetEnd && oSp.getFormatId() === this.endCnxIdFormat) {
						this.endCnxId = oSp.Get_Id();
						this.endCnxIdFormat = undefined;
						bNeedSetEnd = false;
					}
				}
			}
		};
		CNvUniSpPr.prototype.copy = function () {
			var _ret = new CNvUniSpPr();
			_ret.locks = this.locks;
			_ret.stCnxId = this.stCnxId;
			_ret.stCnxIdx = this.stCnxIdx;
			_ret.endCnxId = this.endCnxId;
			_ret.endCnxIdx = this.endCnxIdx;
			return _ret;
		};
		CNvUniSpPr.prototype.getLocks = function() {
			if(!AscFormat.isRealNumber(this.locks)) {
				return 0;
			}
			return this.locks;
		};
		function UniNvPr() {

			CBaseFormatObject.call(this);
			this.cNvPr = null;
			this.UniPr = null;
			this.nvPr = null;
			this.nvUniSpPr = null;

			this.setCNvPr(new CNvPr());
			this.setNvPr(new NvPr());
			this.setUniSpPr(new CNvUniSpPr());

		}

		InitClass(UniNvPr, CBaseFormatObject, AscDFH.historyitem_type_UniNvPr);
		UniNvPr.prototype.setCNvPr = function (cNvPr) {
			AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_UniNvPr_SetCNvPr, this.cNvPr, cNvPr));
			this.cNvPr = cNvPr;
		};
		UniNvPr.prototype.setUniSpPr = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_UniNvPr_SetUniSpPr, this.nvUniSpPr, pr));
			this.nvUniSpPr = pr;
		};
		UniNvPr.prototype.setUniPr = function (uniPr) {
			AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_UniNvPr_SetUniPr, this.UniPr, uniPr));
			this.UniPr = uniPr;
		};
		UniNvPr.prototype.setNvPr = function (nvPr) {
			AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_UniNvPr_SetNvPr, this.nvPr, nvPr));
			this.nvPr = nvPr;
		};
		UniNvPr.prototype.createDuplicate = function () {
			var duplicate = new UniNvPr();
			this.cNvPr && duplicate.setCNvPr(this.cNvPr.createDuplicate());
			this.nvPr && duplicate.setNvPr(this.nvPr.createDuplicate());
			this.nvUniSpPr && duplicate.setUniSpPr(this.nvUniSpPr.copy());
			return duplicate;
		};
		UniNvPr.prototype.Write_ToBinary2 = function (w) {
			w.WriteLong(this.getObjectType());
			w.WriteString2(this.Id);
			writeObject(w, this.cNvPr);
			writeObject(w, this.nvPr);
		};
		UniNvPr.prototype.Read_FromBinary2 = function (r) {
			this.Id = r.GetString2();
			this.cNvPr = readObject(r);
			this.nvPr = readObject(r);
		};
		UniNvPr.prototype.getLocks = function() {
			if(this.nvUniSpPr) {
				return this.nvUniSpPr.getLocks();
			}
			return 0;
		};


		function StyleRef() {
			CBaseNoIdObject.call(this);
			this.idx = 0;
			this.Color = new CUniColor();
			this.styleClr = null;
		}

		InitClass(StyleRef, CBaseNoIdObject, 0);
		StyleRef.prototype.isIdentical = function (styleRef) {
			if (styleRef == null) {
				return false;
			}
			if (this.idx !== styleRef.idx) {
				return false;
			}
			if(this.Color && !styleRef.Color || !this.Color && styleRef.Color) {
				return false;
			}
			if (!this.Color.IsIdentical(styleRef.Color)) {
				return false;
			}
			return true;
		};
		StyleRef.prototype.getObjectType = function () {
			return AscDFH.historyitem_type_StyleRef;
		};
		StyleRef.prototype.setIdx = function (idx) {
			this.idx = idx;
		};
		StyleRef.prototype.setStyleClr = function (styleClr) {
			this.styleClr = styleClr;
		};
		StyleRef.prototype.setColor = function (color) {
			this.Color = color;
		};
		StyleRef.prototype.createDuplicate = function () {
			var duplicate = new StyleRef();
			duplicate.setIdx(this.idx);
			if (this.Color)
				duplicate.setColor(this.Color.createDuplicate());
			return duplicate;
		};
		StyleRef.prototype.Refresh_RecalcData = function () {
		};
		StyleRef.prototype.Write_ToBinary = function (w) {
			writeLong(w, this.idx);
			w.WriteBool(isRealObject(this.Color));
			if (isRealObject(this.Color)) {
				this.Color.Write_ToBinary(w);
			}
		};
		StyleRef.prototype.Read_FromBinary = function (r) {
			this.idx = readLong(r);
			if (r.GetBool()) {
				this.Color = new CUniColor();
				this.Color.Read_FromBinary(r);
			}
		};
		StyleRef.prototype.getNoStyleUnicolor = function (nIdx, aColors) {
			if (this.Color && this.Color.isCorrect()) {
				return this.Color.getNoStyleUnicolor(nIdx, aColors);
			}
			return null;
		};

		function FontRef() {
			CBaseNoIdObject.call(this);
			this.idx = AscFormat.fntStyleInd_none;
			this.Color = null;
		}

		InitClass(FontRef, CBaseNoIdObject, 0);
		FontRef.prototype.getObjectType = function() {
			return AscDFH.historyitem_type_FontRef;
		};
		FontRef.prototype.setIdx = function (idx) {
			this.idx = idx;
		};
		FontRef.prototype.setColor = function (color) {
			this.Color = color;
		};
		FontRef.prototype.createDuplicate = function () {
			var duplicate = new FontRef();
			duplicate.setIdx(this.idx);
			if (this.Color)
				duplicate.setColor(this.Color.createDuplicate());
			return duplicate;
		};
		FontRef.prototype.Write_ToBinary = function (w) {
			writeLong(w, this.idx);
			w.WriteBool(isRealObject(this.Color));
			if (isRealObject(this.Color)) {
				this.Color.Write_ToBinary(w);
			}
		};
		FontRef.prototype.Read_FromBinary = function (r) {
			this.idx = readLong(r);
			if (r.GetBool()) {
				this.Color = new CUniColor();
				this.Color.Read_FromBinary(r);
			}
		};
		FontRef.prototype.getNoStyleUnicolor = function (nIdx, aColors) {
			if (this.Color && this.Color.isCorrect()) {
				return this.Color.getNoStyleUnicolor(nIdx, aColors);
			}
			return null;
		};
		FontRef.prototype.getFirstPartThemeName = function () {
			if (this.idx === AscFormat.fntStyleInd_major) {
				return "+mj-";
			}
			return "+mn-";
		};

		function CShapeStyle() {

			CBaseFormatObject.call(this);
			this.lnRef = null;
			this.fillRef = null;
			this.effectRef = null;
			this.fontRef = null;
		}

		InitClass(CShapeStyle, CBaseFormatObject, AscDFH.historyitem_type_ShapeStyle);
		CShapeStyle.prototype.merge = function (style) {
			if (style != null) {
				if (style.lnRef != null) {
					this.setLnRef(style.lnRef.createDuplicate());
				}

				if (style.fillRef != null) {
					this.setFillRef(style.fillRef.createDuplicate());
				}

				if (style.effectRef != null) {
					this.setEffectRef(style.effectRef.createDuplicate());
				}

				if (style.fontRef != null) {
					this.setFontRef(style.fontRef.createDuplicate());
				}
			}
		};
		CShapeStyle.prototype.createDuplicate = function () {
			var duplicate = new CShapeStyle();
			if (this.lnRef != null) {
				duplicate.setLnRef(this.lnRef.createDuplicate());
			}
			if (this.fillRef != null) {
				duplicate.setFillRef(this.fillRef.createDuplicate());
			}
			if (this.effectRef != null) {
				duplicate.setEffectRef(this.effectRef.createDuplicate());
			}
			if (this.fontRef != null) {
				duplicate.setFontRef(this.fontRef.createDuplicate());
			}
			return duplicate;
		};
		CShapeStyle.prototype.setLnRef = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ShapeStyle_SetLnRef, this.lnRef, pr));
			this.lnRef = pr;
		};
		CShapeStyle.prototype.setFillRef = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ShapeStyle_SetFillRef, this.fillRef, pr));
			this.fillRef = pr;
		};
		CShapeStyle.prototype.setFontRef = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ShapeStyle_SetFontRef, this.fontRef, pr));
			this.fontRef = pr;
		};
		CShapeStyle.prototype.setEffectRef = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ShapeStyle_SetEffectRef, this.effectRef, pr));
			this.effectRef = pr;
		};
		CShapeStyle.prototype.Write_ToBinary = function(w) {
			AscFormat.writeObjectNoId(w, this.lnRef);
			AscFormat.writeObjectNoId(w, this.fillRef);
			AscFormat.writeObjectNoId(w, this.fontRef);
			AscFormat.writeObjectNoId(w, this.effectRef);
		};
		CShapeStyle.prototype.Read_FromBinary = function(r) {
			this.lnRef = AscFormat.readObjectNoId(r);
			this.fillRef = AscFormat.readObjectNoId(r);
			this.fontRef = AscFormat.readObjectNoId(r);
			this.effectRef = AscFormat.readObjectNoId(r);
		};

		var LINE_PRESETS_MAP = {};

		LINE_PRESETS_MAP["line"] = true;
		LINE_PRESETS_MAP["bracePair"] = true;
		LINE_PRESETS_MAP["leftBrace"] = true;
		LINE_PRESETS_MAP["rightBrace"] = true;
		LINE_PRESETS_MAP["bracketPair"] = true;
		LINE_PRESETS_MAP["leftBracket"] = true;
		LINE_PRESETS_MAP["rightBracket"] = true;
		LINE_PRESETS_MAP["bentConnector2"] = true;
		LINE_PRESETS_MAP["bentConnector3"] = true;
		LINE_PRESETS_MAP["bentConnector4"] = true;
		LINE_PRESETS_MAP["bentConnector5"] = true;
		LINE_PRESETS_MAP["curvedConnector2"] = true;
		LINE_PRESETS_MAP["curvedConnector3"] = true;
		LINE_PRESETS_MAP["curvedConnector4"] = true;
		LINE_PRESETS_MAP["curvedConnector5"] = true;
		LINE_PRESETS_MAP["straightConnector1"] = true;
		LINE_PRESETS_MAP["arc"] = true;

		function CreateDefaultShapeStyle(preset) {

			var b_line = typeof preset === "string" && LINE_PRESETS_MAP[preset];
			var tx_color = b_line;
			var unicolor;
			var style = new CShapeStyle();
			var lnRef = new StyleRef();
			lnRef.setIdx(b_line ? 1 : 2);

			unicolor = new CUniColor();
			unicolor.setColor(new CSchemeColor());
			unicolor.color.setId(g_clr_accent1);
			unicolor.setMods(new CColorModifiers());
			unicolor.Mods.addMod("shade", 50000);
			lnRef.setColor(unicolor);

			style.setLnRef(lnRef);


			var fillRef = new StyleRef();
			unicolor = new CUniColor();
			unicolor.setColor(new CSchemeColor());
			unicolor.color.setId(g_clr_accent1);
			fillRef.setIdx(b_line ? 0 : 1);
			fillRef.setColor(unicolor);
			style.setFillRef(fillRef);


			var effectRef = new StyleRef();
			unicolor = new CUniColor();
			unicolor.setColor(new CSchemeColor());
			unicolor.color.setId(g_clr_accent1);
			effectRef.setIdx(0);
			effectRef.setColor(unicolor);
			style.setEffectRef(effectRef);

			var fontRef = new FontRef();
			unicolor = new CUniColor();
			unicolor.setColor(new CSchemeColor());
			unicolor.color.setId(tx_color ? 15 : 12);
			fontRef.setIdx(AscFormat.fntStyleInd_minor);
			fontRef.setColor(unicolor);
			style.setFontRef(fontRef);
			return style;
		}


		function CXfrm() {
			CBaseFormatObject.call(this);
			this.offX = null;
			this.offY = null;
			this.extX = null;
			this.extY = null;
			this.chOffX = null;
			this.chOffY = null;
			this.chExtX = null;
			this.chExtY = null;

			this.flipH = null;
			this.flipV = null;
			this.rot = null;
		}

		InitClass(CXfrm, CBaseFormatObject, AscDFH.historyitem_type_Xfrm);
		CXfrm.prototype.isNotNull = function () {
			return isRealNumber(this.offX) && isRealNumber(this.offY) && isRealNumber(this.extX) && isRealNumber(this.extY);
		};

		CXfrm.prototype.isNull = function () {
			return !isRealNumber(this.offX) && !isRealNumber(this.offY) && !isRealNumber(this.extX) && !isRealNumber(this.extY);
		};

		CXfrm.prototype.isNotNullForGroup = function () {
			return isRealNumber(this.offX) && isRealNumber(this.offY)
				&& isRealNumber(this.chOffX) && isRealNumber(this.chOffY)
				&& isRealNumber(this.extX) && isRealNumber(this.extY)
				&& isRealNumber(this.chExtX) && isRealNumber(this.chExtY);
		};


		CXfrm.prototype.isZero = function () {
			return (
				this.offX === 0 &&
				this.offY === 0 &&
				this.extX === 0 &&
				this.extY === 0
			);
		};
		CXfrm.prototype.isZeroCh = function () {
			return (
				this.chOffX === 0 &&
				this.chOffY === 0 &&
				this.chExtX === 0 &&
				this.chExtY === 0
			);
		};

		CXfrm.prototype.isZeroInGroup = function () {
			return this.isZero() && this.isZeroCh();
		};
		CXfrm.prototype.isEqual = function (xfrm) {
			return xfrm && this.offX === xfrm.offX && this.offY === xfrm.offY && this.extX === xfrm.extX &&
				this.extY === xfrm.extY && this.chOffX === xfrm.chOffX && this.chOffY === xfrm.chOffY && this.chExtX === xfrm.chExtX &&
				this.chExtY === xfrm.chExtY;
		};
		CXfrm.prototype.merge = function (xfrm) {
			if(!xfrm) {
				return;
			}
			if (xfrm.offX != null) {
				this.setOffX(xfrm.offX);
			}
			if (xfrm.offY != null) {
				this.setOffY(xfrm.offY);
			}


			if (xfrm.extX != null) {
				this.setExtX(xfrm.extX);
			}
			if (xfrm.extY != null) {
				this.setExtY(xfrm.extY);
			}


			if (xfrm.chOffX != null) {
				this.setChOffX(xfrm.chOffX);
			}
			if (xfrm.chOffY != null) {
				this.setChOffY(xfrm.chOffY);
			}


			if (xfrm.chExtX != null) {
				this.setChExtX(xfrm.chExtX);
			}
			if (xfrm.chExtY != null) {
				this.setChExtY(xfrm.chExtY);
			}


			if (xfrm.flipH != null) {
				this.setFlipH(xfrm.flipH);
			}
			if (xfrm.flipV != null) {
				this.setFlipV(xfrm.flipV);
			}

			if (xfrm.rot != null) {
				this.setRot(xfrm.rot);
			}
		};
		CXfrm.prototype.createDuplicate = function () {
			var duplicate = new CXfrm();
			duplicate.setOffX(this.offX);
			duplicate.setOffY(this.offY);
			duplicate.setExtX(this.extX);
			duplicate.setExtY(this.extY);
			duplicate.setChOffX(this.chOffX);
			duplicate.setChOffY(this.chOffY);
			duplicate.setChExtX(this.chExtX);
			duplicate.setChExtY(this.chExtY);

			duplicate.setFlipH(this.flipH);
			duplicate.setFlipV(this.flipV);
			duplicate.setRot(this.rot);
			return duplicate;
		};
		CXfrm.prototype.setParent = function (pr) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Xfrm_SetParent, this.parent, pr));
			this.parent = pr;
		};
		CXfrm.prototype.setOffX = function (pr) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Xfrm_SetOffX, this.offX, pr));
			this.offX = pr;
			this.handleUpdatePosition();
		};
		CXfrm.prototype.setOffY = function (pr) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Xfrm_SetOffY, this.offY, pr));
			this.offY = pr;
			this.handleUpdatePosition();
		};
		CXfrm.prototype.setExtX = function (pr) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Xfrm_SetExtX, this.extX, pr));
			this.extX = pr;
			this.handleUpdateExtents(true);
		};
		CXfrm.prototype.setExtY = function (pr) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Xfrm_SetExtY, this.extY, pr));
			this.extY = pr;
			this.handleUpdateExtents(false);
		};
		CXfrm.prototype.setChOffX = function (pr) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Xfrm_SetChOffX, this.chOffX, pr));
			this.chOffX = pr;
			this.handleUpdateChildOffset();
		};
		CXfrm.prototype.setChOffY = function (pr) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Xfrm_SetChOffY, this.chOffY, pr));
			this.chOffY = pr;
			this.handleUpdateChildOffset();
		};
		CXfrm.prototype.setChExtX = function (pr) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Xfrm_SetChExtX, this.chExtX, pr));
			this.chExtX = pr;
			this.handleUpdateChildExtents();
		};
		CXfrm.prototype.setChExtY = function (pr) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Xfrm_SetChExtY, this.chExtY, pr));
			this.chExtY = pr;
			this.handleUpdateChildExtents();
		};
		CXfrm.prototype.setFlipH = function (pr) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Xfrm_SetFlipH, this.flipH, pr));
			this.flipH = pr;
			this.handleUpdateFlip();
		};
		CXfrm.prototype.setFlipV = function (pr) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Xfrm_SetFlipV, this.flipV, pr));
			this.flipV = pr;
			this.handleUpdateFlip();
		};
		CXfrm.prototype.setRot = function (pr) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Xfrm_SetRot, this.rot, pr));
			this.rot = pr;
			this.handleUpdateRot();
		};
		CXfrm.prototype.getRot = function() {
			return this.rot;
		};
		CXfrm.prototype.shift = function(dDX, dDY) {
			if(this.offX !== null && this.offY !== null) {
				this.setOffX(this.offX + dDX);
				this.setOffY(this.offY + dDY);
			}
		};
		CXfrm.prototype.handleUpdatePosition = function () {
			if (this.parent && this.parent.handleUpdatePosition) {
				this.parent.handleUpdatePosition();
			}
		};
		CXfrm.prototype.handleUpdateExtents = function (bExtX) {
			if (this.parent && this.parent.handleUpdateExtents) {
				this.parent.handleUpdateExtents(bExtX);
			}
		};
		CXfrm.prototype.handleUpdateChildOffset = function () {
			if (this.parent && this.parent.handleUpdateChildOffset) {
				this.parent.handleUpdateChildOffset();
			}
		};
		CXfrm.prototype.handleUpdateChildExtents = function () {
			if (this.parent && this.parent.handleUpdateChildExtents) {
				this.parent.handleUpdateChildExtents();
			}
		};
		CXfrm.prototype.handleUpdateFlip = function () {
			if (this.parent && this.parent.handleUpdateFlip) {
				this.parent.handleUpdateFlip();
			}
		};
		CXfrm.prototype.handleUpdateRot = function () {
			if (this.parent && this.parent.handleUpdateRot) {
				this.parent.handleUpdateRot();
			}
		};
		CXfrm.prototype.Refresh_RecalcData = function (data) {
			switch (data.Type) {
				case AscDFH.historyitem_Xfrm_SetOffX: {
					this.handleUpdatePosition();
					break;
				}
				case AscDFH.historyitem_Xfrm_SetOffY: {
					this.handleUpdatePosition();
					break;
				}
				case AscDFH.historyitem_Xfrm_SetExtX: {
					this.handleUpdateExtents();
					break;
				}
				case AscDFH.historyitem_Xfrm_SetExtY: {
					this.handleUpdateExtents();
					break;
				}
				case AscDFH.historyitem_Xfrm_SetChOffX: {
					this.handleUpdateChildOffset();
					break;
				}
				case AscDFH.historyitem_Xfrm_SetChOffY: {
					this.handleUpdateChildOffset();
					break;
				}
				case AscDFH.historyitem_Xfrm_SetChExtX: {
					this.handleUpdateChildExtents();
					break;
				}
				case AscDFH.historyitem_Xfrm_SetChExtY: {
					this.handleUpdateChildExtents();
					break;
				}
				case AscDFH.historyitem_Xfrm_SetFlipH: {
					this.handleUpdateFlip();
					break;
				}
				case AscDFH.historyitem_Xfrm_SetFlipV: {
					this.handleUpdateFlip();
					break;
				}
				case AscDFH.historyitem_Xfrm_SetRot: {
					this.handleUpdateRot();
					break;
				}
			}
		};

		CXfrm.prototype.fillStandardSmartArtXfrm = function () {
			const oApi = Asc.editor || editor;
			if (oApi && oApi.isDocumentEditor) {
				this.setOffX(0);
				this.setOffY(0);
			} else {
				this.setOffX(2032000 * AscCommonWord.g_dKoef_emu_to_mm);
				this.setOffY(719666 * AscCommonWord.g_dKoef_emu_to_mm);
			}
			this.setExtX(8128000 * AscCommonWord.g_dKoef_emu_to_mm);
			this.setExtY(5418667 * AscCommonWord.g_dKoef_emu_to_mm);
			this.setChOffX(0);
			this.setChOffY(0);
			this.setChExtX(this.extX);
			this.setChExtY(this.extY);
		};

		function CEffectProperties() {
			CBaseNoIdObject.call(this);
			this.EffectDag = null;
			this.EffectLst = null;
		}

		InitClass(CEffectProperties, CBaseNoIdObject, AscDFH.historyitem_type_CEffectProperties);
		CEffectProperties.prototype.createDuplicate = function () {
			var oCopy = new CEffectProperties();
			if (this.EffectDag) {
				oCopy.EffectDag = this.EffectDag.createDuplicate();
			}
			if (this.EffectLst) {
				oCopy.EffectLst = this.EffectLst.createDuplicate();
			}
			return oCopy;
		};
		CEffectProperties.prototype.merge = function (effectPr) {
			// if (effectPr.EffectDag) {
			// 	if (!this.EffectDag) {
			// 		this.EffectDag = new CEffectContainer();
			// 	}
			// 	this.EffectDag.merge(effectPr.EffectDag);
			// }
			if (effectPr.EffectLst) {
				if (!this.EffectLst) {
					this.EffectLst = new CEffectLst();
				}
				this.EffectLst.merge(effectPr.EffectLst);
			}
		};
		CEffectProperties.prototype.Write_ToBinary = function (w) {
			var nFlags = 0;
			if (this.EffectDag) {
				nFlags |= 1;
			}
			if (this.EffectLst) {
				nFlags |= 2;
			}
			w.WriteLong(nFlags);
			if (this.EffectDag) {
				this.EffectDag.Write_ToBinary(w);
			}
			if (this.EffectLst) {
				this.EffectLst.Write_ToBinary(w);
			}
		};
		CEffectProperties.prototype.Read_FromBinary = function (r) {
			var nFlags = r.GetLong();
			if (nFlags & 1) {
				this.EffectDag = new CEffectContainer();
				this.EffectDag.Read_FromBinary(r);
			}
			if (nFlags & 2) {
				this.EffectLst = new CEffectLst();
				this.EffectLst.Read_FromBinary(r);
			}
		};

		function CEffectLst() {
			CBaseNoIdObject.call(this);
			this.blur = null;
			this.fillOverlay = null;
			this.glow = null;
			this.innerShdw = null;
			this.outerShdw = null;
			this.prstShdw = null;
			this.reflection = null;
			this.softEdge = null;
		}

		InitClass(CEffectLst, CBaseNoIdObject, 0);
		CEffectLst.prototype.createDuplicate = function () {
			var oCopy = new CEffectLst();
			if (this.blur) {
				oCopy.blur = this.blur.createDuplicate();
			}
			if (this.fillOverlay) {
				oCopy.fillOverlay = this.fillOverlay.createDuplicate();
			}
			if (this.glow) {
				oCopy.glow = this.glow.createDuplicate();
			}
			if (this.innerShdw) {
				oCopy.innerShdw = this.innerShdw.createDuplicate();
			}
			if (this.outerShdw) {
				oCopy.outerShdw = this.outerShdw.createDuplicate();
			}
			if (this.prstShdw) {
				oCopy.prstShdw = this.prstShdw.createDuplicate();
			}
			if (this.reflection) {
				oCopy.reflection = this.reflection.createDuplicate();
			}
			if (this.softEdge) {
				oCopy.softEdge = this.softEdge.createDuplicate();
			}
			return oCopy;
		};
		CEffectLst.prototype.merge = function (effectLst) {
			if (effectLst.blur) {
				this.blur = effectLst.blur.createDuplicate();
			}
			if (effectLst.fillOverlay) {
				this.fillOverlay = effectLst.fillOverlay.createDuplicate();
			}
			if (effectLst.glow) {
				this.glow = effectLst.glow.createDuplicate();
			}
			if (effectLst.innerShdw) {
				this.innerShdw = effectLst.innerShdw.createDuplicate();
			}
			if (effectLst.outerShdw) {
				this.outerShdw = effectLst.outerShdw.createDuplicate();
			}
			if (effectLst.prstShdw) {
				this.prstShdw = effectLst.prstShdw.createDuplicate();
			}
			if (effectLst.reflection) {
				this.reflection = effectLst.reflection.createDuplicate();
			}
			if (effectLst.softEdge) {
				this.softEdge = effectLst.softEdge.createDuplicate();
			}
		};
		CEffectLst.prototype.Write_ToBinary = function (w) {
			var nFlags = 0;
			if (this.blur) {
				nFlags |= 1;
			}
			if (this.fillOverlay) {
				nFlags |= 2;
			}
			if (this.glow) {
				nFlags |= 4;
			}
			if (this.innerShdw) {
				nFlags |= 8;
			}
			if (this.outerShdw) {
				nFlags |= 16;
			}
			if (this.prstShdw) {
				nFlags |= 32;
			}
			if (this.reflection) {
				nFlags |= 64;
			}
			if (this.softEdge) {
				nFlags |= 128;
			}
			w.WriteLong(nFlags);
			if (this.blur) {
				this.blur.Write_ToBinary(w);
			}
			if (this.fillOverlay) {
				this.fillOverlay.Write_ToBinary(w);
			}
			if (this.glow) {
				this.glow.Write_ToBinary(w);
			}
			if (this.innerShdw) {
				this.innerShdw.Write_ToBinary(w);
			}
			if (this.outerShdw) {
				this.outerShdw.Write_ToBinary(w);
			}
			if (this.prstShdw) {
				this.prstShdw.Write_ToBinary(w);
			}
			if (this.reflection) {
				this.reflection.Write_ToBinary(w);
			}
			if (this.softEdge) {
				this.softEdge.Write_ToBinary(w);
			}
		};
		CEffectLst.prototype.Read_FromBinary = function (r) {
			var nFlags = r.GetLong();
			if (nFlags & 1) {
				this.blur = new CBlur();
				r.GetLong();
				this.blur.Read_FromBinary(r);
			}
			if (nFlags & 2) {
				this.fillOverlay = new CFillOverlay();
				r.GetLong();
				this.fillOverlay.Read_FromBinary(r);
			}
			if (nFlags & 4) {
				this.glow = new CGlow();
				r.GetLong();
				this.glow.Read_FromBinary(r);
			}
			if (nFlags & 8) {
				this.innerShdw = new CInnerShdw();
				r.GetLong();
				this.innerShdw.Read_FromBinary(r);
			}
			if (nFlags & 16) {
				this.outerShdw = new COuterShdw();
				r.GetLong();
				this.outerShdw.Read_FromBinary(r);
			}
			if (nFlags & 32) {
				this.prstShdw = new CPrstShdw();
				r.GetLong();
				this.prstShdw.Read_FromBinary(r);
			}
			if (nFlags & 64) {
				this.reflection = new CReflection();
				r.GetLong();
				this.reflection.Read_FromBinary(r);
			}
			if (nFlags & 128) {
				this.softEdge = new CSoftEdge();
				r.GetLong();
				this.softEdge.Read_FromBinary(r);
			}
		};

		function CSpPr() {

			CBaseFormatObject.call(this);
			this.bwMode = 0;

			this.xfrm = null;
			this.geometry = null;
			this.Fill = null;
			this.ln = null;
			this.parent = null;

			this.effectProps = null;
		}

		InitClass(CSpPr, CBaseFormatObject, AscDFH.historyitem_type_SpPr);
		CSpPr.prototype.Refresh_RecalcData = function (data) {
			switch (data.Type) {
				case AscDFH.historyitem_SpPr_SetParent: {
					break;
				}
				case AscDFH.historyitem_SpPr_SetBwMode: {
					break;
				}
				case AscDFH.historyitem_SpPr_SetXfrm: {
					this.handleUpdateExtents();
					break;
				}
				case AscDFH.historyitem_SpPr_SetGeometry:
				case AscDFH.historyitem_SpPr_SetEffectPr: {
					this.handleUpdateGeometry();
					break;
				}
				case AscDFH.historyitem_SpPr_SetFill: {
					this.handleUpdateFill();
					break;
				}
				case AscDFH.historyitem_SpPr_SetLn: {
					this.handleUpdateLn();
					break;
				}
			}
		};
		CSpPr.prototype.Refresh_RecalcData2 = function (data) {
		};
		CSpPr.prototype.createDuplicate = function () {
			var duplicate = new CSpPr();
			duplicate.setBwMode(this.bwMode);
			if (this.xfrm) {
				duplicate.setXfrm(this.xfrm.createDuplicate());
				duplicate.xfrm.setParent(duplicate);
			}
			if (this.geometry != null) {
				duplicate.setGeometry(this.geometry.createDuplicate());
			}
			if (this.Fill != null) {
				duplicate.setFill(this.Fill.createDuplicate());
			}
			if (this.ln != null) {
				duplicate.setLn(this.ln.createDuplicate());
			}
			if (this.effectProps) {
				duplicate.setEffectPr(this.effectProps.createDuplicate());
			}
			return duplicate;
		};
		CSpPr.prototype.createDuplicateForSmartArt = function () {
			var duplicate = new CSpPr();
			if (this.Fill != null) {
				duplicate.setFill(this.Fill.createDuplicate());
			}
			return duplicate;
		};
		CSpPr.prototype.hasRGBFill = function () {
			return this.Fill && this.Fill.fill && this.Fill.fill.color
				&& this.Fill.fill.color.color && this.Fill.fill.color.color.type === c_oAscColor.COLOR_TYPE_SRGB;
		};
		CSpPr.prototype.hasNoFill = function () {
			if (this.Fill) {
				return this.Fill.isNoFill();
			}
			return false;
		};
		CSpPr.prototype.hasNoFillLine = function () {
			if (this.ln) {
				return this.ln.isNoFillLine();
			}
			return false;
		};
		CSpPr.prototype.checkUniFillRasterImageId = function (unifill) {
			return checkUniFillRasterImageId(unifill);
		};
		CSpPr.prototype.checkBlipFillRasterImage = function (images) {
			var fill_image_id = this.checkUniFillRasterImageId(this.Fill);
			if (fill_image_id !== null)
				images.push(fill_image_id);
			if (this.ln) {
				var line_image_id = this.checkUniFillRasterImageId(this.ln.Fill);
				if (line_image_id)
					images.push(line_image_id);
			}
		};
		CSpPr.prototype.changeShadow = function (oShadow) {
			if (oShadow) {
				var oEffectProps = this.effectProps ? this.effectProps.createDuplicate() : new AscFormat.CEffectProperties();
				if (!oEffectProps.EffectLst) {
					oEffectProps.EffectLst = new CEffectLst();
				}
				oEffectProps.EffectLst.outerShdw = oShadow.createDuplicate();
				this.setEffectPr(oEffectProps);
			} else {
				if (this.effectProps) {
					if (this.effectProps.EffectLst) {
						if (this.effectProps.EffectLst.outerShdw) {
							var oEffectProps = this.effectProps.createDuplicate();
							oEffectProps.EffectLst.outerShdw = null;
							this.setEffectPr(oEffectProps);
						}
					}
				}
			}
		};
		CSpPr.prototype.merge = function (spPr) {
			/*if(spPr.xfrm != null)
         {
         this.xfrm.merge(spPr.xfrm);
         }  */
			if (spPr.geometry != null) {
				this.geometry = spPr.geometry.createDuplicate();
			}

			if (spPr.Fill != null && spPr.Fill.fill != null) {
				//this.Fill = spPr.Fill.createDuplicate();
			}

			/*if(spPr.ln!=null)
         {
         if(this.ln == null)
         this.ln = new CLn();
         this.ln.merge(spPr.ln);
         }  */
		};
		CSpPr.prototype.fullMerge = function (spPr) {
			if (spPr.xfrm != null) {
				this.xfrm.merge(spPr.xfrm);
			}
			if (spPr.geometry !== null) {
				this.setGeometry(spPr.geometry.createDuplicate());
			}

			if (spPr.Fill !== null && spPr.Fill.fill !== null) {
				this.setFill(spPr.Fill.createDuplicate());
			}

			if (spPr.ln != null) {
				if (this.ln == null)
					this.setLn(new CLn());

				this.ln.merge(spPr.ln);
			}
			if (spPr.effectProps !== null) {
				if (this.effectProps === null) {
					this.setEffectPr(new CEffectProperties());
				}
				this.effectProps.merge(spPr.effectProps);
			}
		};
		CSpPr.prototype.setParent = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_SpPr_SetParent, this.parent, pr));
			this.parent = pr;
		};
		CSpPr.prototype.setBwMode = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_SpPr_SetBwMode, this.bwMode, pr));
			this.bwMode = pr;
		};
		CSpPr.prototype.setXfrm = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_SpPr_SetXfrm, this.xfrm, pr));
			this.xfrm = pr;
			if(pr) {
				pr.setParent(this);
			}
		};
		CSpPr.prototype.setGeometry = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_SpPr_SetGeometry, this.geometry, pr));
			this.geometry = pr;
			if (this.geometry) {
				this.geometry.setParent(this);
			}
			this.handleUpdateGeometry();
		};
		CSpPr.prototype.setFill = function (pr) {
			if(!pr || !pr.isBlipFill() || !Asc.editor.evalCommand)
				AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SpPr_SetFill, this.Fill, pr));
			this.Fill = pr;
			if (this.parent && this.parent.handleUpdateFill) {
				this.parent.handleUpdateFill();
			}
		};
		CSpPr.prototype.setLn = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SpPr_SetLn, this.ln, pr));
			this.ln = pr;
			if (this.parent && this.parent.handleUpdateLn) {
				this.parent.handleUpdateLn();
			}
		};
		CSpPr.prototype.setEffectPr = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SpPr_SetEffectPr, this.effectProps, pr));
			this.effectProps = pr;
		};
		CSpPr.prototype.handleUpdatePosition = function () {
			if (this.parent && this.parent.handleUpdatePosition) {
				this.parent.handleUpdatePosition();
			}
		};
		CSpPr.prototype.handleUpdateExtents = function (bExtX) {
			if (this.parent && this.parent.handleUpdateExtents) {
				this.parent.handleUpdateExtents(bExtX);
			}
		};
		CSpPr.prototype.handleUpdateChildOffset = function () {
			if (this.parent && this.parent.handleUpdateChildOffset) {
				this.parent.handleUpdateChildOffset();
			}
		};
		CSpPr.prototype.handleUpdateChildExtents = function () {
			if (this.parent && this.parent.handleUpdateChildExtents) {
				this.parent.handleUpdateChildExtents();
			}
		};
		CSpPr.prototype.handleUpdateFlip = function () {
			if (this.parent && this.parent.handleUpdateFlip) {
				this.parent.handleUpdateFlip();
			}
		};
		CSpPr.prototype.handleUpdateRot = function () {
			if (this.parent && this.parent.handleUpdateRot) {
				this.parent.handleUpdateRot();
			}
		};
		CSpPr.prototype.handleUpdateGeometry = function () {
			if (this.parent && this.parent.handleUpdateGeometry) {
				this.parent.handleUpdateGeometry();
			}
		};
		CSpPr.prototype.handleUpdateFill = function () {
			if (this.parent && this.parent.handleUpdateFill) {
				this.parent.handleUpdateFill();
			}
		};
		CSpPr.prototype.handleUpdateLn = function () {
			if (this.parent && this.parent.handleUpdateLn) {
				this.parent.handleUpdateLn();
			}
		};
		CSpPr.prototype.setLineFill = function () {
			if (this.ln && this.ln.Fill) {
				this.setFill(this.ln.Fill.createDuplicate());
			}
		};
		CSpPr.prototype.getOuterShdw = function () {
			if (this.effectProps && this.effectProps.EffectLst && this.effectProps.EffectLst.outerShdw) {
				return this.effectProps.EffectLst.outerShdw;
			}
			return null;
		};
// ----------------------------------

// THEME ----------------------------

		var g_clr_MIN = 0;
		var g_clr_accent1 = 0;
		var g_clr_accent2 = 1;
		var g_clr_accent3 = 2;
		var g_clr_accent4 = 3;
		var g_clr_accent5 = 4;
		var g_clr_accent6 = 5;
		var g_clr_dk1 = 6;
		var g_clr_dk2 = 7;
		var g_clr_folHlink = 8;
		var g_clr_hlink = 9;
		var g_clr_lt1 = 10;
		var g_clr_lt2 = 11;
		var g_clr_MAX = 11;

		var g_clr_bg1 = g_clr_lt1;
		var g_clr_bg2 = g_clr_lt2;
		var g_clr_tx1 = g_clr_dk1;
		var g_clr_tx2 = g_clr_dk2;

		var phClr = 14;
		var tx1 = 15;
		var tx2 = 16;


		let CLR_IDX_MAP = {};
		CLR_IDX_MAP["dk1"] = 8;
		CLR_IDX_MAP["lt1"] = 12;
		CLR_IDX_MAP["dk2"] = 9;
		CLR_IDX_MAP["lt2"] = 13;
		CLR_IDX_MAP["accent1"] = 0;
		CLR_IDX_MAP["accent2"] = 1;
		CLR_IDX_MAP["accent3"] = 2;
		CLR_IDX_MAP["accent4"] = 3;
		CLR_IDX_MAP["accent5"] = 4;
		CLR_IDX_MAP["accent6"] = 5;
		CLR_IDX_MAP["hlink"] = 11;
		CLR_IDX_MAP["folHlink"] = 10;


		let CLR_NAME_MAP = {};
		CLR_NAME_MAP[8] = "dk1";
		CLR_NAME_MAP[12] = "lt1";
		CLR_NAME_MAP[9] = "dk2";
		CLR_NAME_MAP[13] = "lt2";
		CLR_NAME_MAP[0] = "accent1";
		CLR_NAME_MAP[1] = "accent2";
		CLR_NAME_MAP[2] = "accent3";
		CLR_NAME_MAP[3] = "accent4";
		CLR_NAME_MAP[4] = "accent5";
		CLR_NAME_MAP[5] = "accent6";
		CLR_NAME_MAP[11] = "hlink";
		CLR_NAME_MAP[10] = "folHlink";

		function ClrScheme() {
			CBaseNoIdObject.call(this);
			this.name = "";
			this.colors = [];
			/**
			 * Used to store xml extensions (e.g. visio variant colors)
			 * set in sdkjs-ooxml/common/Drawings/Format.js ClrScheme.prototype.readChildXml
			 */
			this.clrSchemeExtLst = null;

			for (var i = g_clr_MIN; i <= g_clr_MAX; i++)
				this.colors[i] = null;

		}

		InitClass(ClrScheme, CBaseNoIdObject, 0);
		ClrScheme.prototype.isIdentical = function (clrScheme) {
			if (!(clrScheme instanceof ClrScheme)) {
				return false;
			}
			if (clrScheme.name !== this.name) {
				return false;
			}
			for (var _clr_index = g_clr_MIN; _clr_index <= g_clr_MAX; ++_clr_index) {
				if (this.colors[_clr_index]) {
					if (!this.colors[_clr_index].IsIdentical(clrScheme.colors[_clr_index])) {
						return false;
					}
				} else {
					if (clrScheme.colors[_clr_index]) {
						return false;
					}
				}
			}
			return true;
		};
		ClrScheme.prototype.createDuplicate = function () {
			var _duplicate = new ClrScheme();
			_duplicate.name = this.name;
			for (var _clr_index = 0; _clr_index <= this.colors.length; ++_clr_index) {
				if (this.colors[_clr_index]) {
					_duplicate.colors[_clr_index] = this.colors[_clr_index].createDuplicate();
				}
			}
			return _duplicate;
		};
		ClrScheme.prototype.Write_ToBinary = function (w) {
			w.WriteLong(this.colors.length);
			w.WriteString2(this.name);
			for (var i = 0; i < this.colors.length; ++i) {
				w.WriteBool(isRealObject(this.colors[i]));
				if (isRealObject(this.colors[i])) {
					this.colors[i].Write_ToBinary(w);
				}
			}

		};
		ClrScheme.prototype.Read_FromBinary = function (r) {
			var len = r.GetLong();
			this.name = r.GetString2();
			for (var i = 0; i < len; ++i) {
				if (r.GetBool()) {
					this.colors[i] = new CUniColor();
					this.colors[i].Read_FromBinary(r);
				} else {
					this.colors[i] = null;
				}
			}

		};
		ClrScheme.prototype.setName = function (name) {
			this.name = name;
		};
		ClrScheme.prototype.addColor = function (index, color) {
			this.colors[index] = color;
		};

		function ClrMap() {
			CBaseFormatObject.call(this);
			this.color_map = [];

			for (var i = g_clr_MIN; i <= g_clr_MAX; i++)
				this.color_map[i] = null;
		}

		InitClass(ClrMap, CBaseFormatObject, AscDFH.historyitem_type_ClrMap);
		ClrMap.prototype.Refresh_RecalcData = function () {
		};
		ClrMap.prototype.notAllowedWithoutId = function () {
			return false;
		};
		ClrMap.prototype.createDuplicate = function () {
			var _copy = new ClrMap();
			for (var _color_index = g_clr_MIN; _color_index <= this.color_map.length; ++_color_index) {
				_copy.setClr(_color_index, this.color_map[_color_index]);
			}
			return _copy;
		};
		ClrMap.prototype.compare = function (other) {
			if (!other)
				return false;
			for (var i = g_clr_MIN; i < this.color_map.length; ++i) {
				if (this.color_map[i] !== other.color_map[i]) {
					return false;
				}
			}
			return true;
		};
		ClrMap.prototype.setClr = function (index, clr) {
			AscCommon.History.Add(new CChangesDrawingsContentLongMap(this, AscDFH.historyitem_ClrMap_SetClr, index, [clr], true));
			this.color_map[index] = clr;
		};


		ClrMap.prototype.SchemeClr_GetBYTECode = function (sValue) {
			if ("accent1" === sValue)
				return 0;
			if ("accent2" === sValue)
				return 1;
			if ("accent3" === sValue)
				return 2;
			if ("accent4" === sValue)
				return 3;
			if ("accent5" === sValue)
				return 4;
			if ("accent6" === sValue)
				return 5;
			if ("bg1" === sValue)
				return 6;
			if ("bg2" === sValue)
				return 7;
			if ("dk1" === sValue)
				return 8;
			if ("dk2" === sValue)
				return 9;
			if ("folHlink" === sValue)
				return 10;
			if ("hlink" === sValue)
				return 11;
			if ("lt1" === sValue)
				return 12;
			if ("lt2" === sValue)
				return 13;
			if ("phClr" === sValue)
				return 14;
			if ("tx1" === sValue)
				return 15;
			if ("tx2" === sValue)
				return 16;
			return 0;
		};
		ClrMap.prototype.SchemeClr_GetStringCode = function (val) {
			switch (val) {
				case 0:
					return ("accent1");
				case 1:
					return ("accent2");
				case 2:
					return ("accent3");
				case 3:
					return ("accent4");
				case 4:
					return ("accent5");
				case 5:
					return ("accent6");
				case 6:
					return ("bg1");
				case 7:
					return ("bg2");
				case 8:
					return ("dk1");
				case 9:
					return ("dk2");
				case 10:
					return ("folHlink");
				case 11:
					return ("hlink");
				case 12:
					return ("lt1");
				case 13:
					return ("lt2");
				case 14:
					return ("phClr");
				case 15:
					return ("tx1");
				case 16:
					return ("tx2");
			}
			return ("accent1");
		}

		ClrMap.prototype.getColorIdx = function (name) {
			if ("accent1" === name)
				return 0;
			if ("accent2" === name)
				return 1;
			if ("accent3" === name)
				return 2;
			if ("accent4" === name)
				return 3;
			if ("accent5" === name)
				return 4;
			if ("accent6" === name)
				return 5;
			if ("bg1" === name)
				return 6;
			if ("bg2" === name)
				return 7;
			if ("dk1" === name)
				return 8;
			if ("dk2" === name)
				return 9;
			if ("folHlink" === name)
				return 10;
			if ("hlink" === name)
				return 11;
			if ("lt1" === name)
				return 12;
			if ("lt2" === name)
				return 13;
			if ("phClr" === name)
				return 14;
			if ("tx1" === name)
				return 15;
			if ("tx2" === name)
				return 16;

			return null;
		};
		ClrMap.prototype.getColorName = function (nIdx) {
			if (0 === nIdx)
				return "accent1";
			if (1 === nIdx)
				return "accent2";
			if (2 === nIdx)
				return "accent3";
			if (3 === nIdx)
				return "accent4";
			if (4 === nIdx)
				return "accent5";
			if (5 === nIdx)
				return "accent6";
			if (6 === nIdx)
				return "bg1";
			if (7 === nIdx)
				return "bg2";
			if (8 === nIdx)
				return "dk1";
			if (9 === nIdx)
				return "dk2";
			if (10 === nIdx)
				return "folHlink";
			if (11 === nIdx)
				return "hlink";
			if (12 === nIdx)
				return "lt1";
			if (13 === nIdx)
				return "lt2";
			if (14 === nIdx)
				return "phClr";
			if (15 === nIdx)
				return "tx1";
			if (16 === nIdx)
				return "tx2";

			return null;
		};
		ClrMap.prototype.getColorIdxWord = function (name) {
			if ("accent1" === name)
				return 0;
			if ("accent2" === name)
				return 1;
			if ("accent3" === name)
				return 2;
			if ("accent4" === name)
				return 3;
			if ("accent5" === name)
				return 4;
			if ("accent6" === name)
				return 5;
			if ("dark1" === name)
				return 8;
			if ("dark2" === name)
				return 9;
			if ("followedHyperlink" === name)
				return 10;
			if ("hyperlink" === name)
				return 11;
			if ("light1" === name)
				return 12;
			if ("light2" === name)
				return 13;

			return null;
		};
		ClrMap.prototype.getColorNameWord = function (val) {
			switch (val) {
				case 0:
					return ("accent1");
				case 1:
					return ("accent2");
				case 2:
					return ("accent3");
				case 3:
					return ("accent4");
				case 4:
					return ("accent5");
				case 5:
					return ("accent6");
				case 8:
					return ("dark1");
				case 9:
					return ("dark2");
				case 10:
					return ("followedHyperlink");
				case 11:
					return ("hyperlink");
				case 12:
					return ("light1");
				case 13:
					return ("light2");
			}
			return ("");
		};

		function ExtraClrScheme() {
			CBaseFormatObject.call(this);
			this.clrScheme = null;
			this.clrMap = null;

		}

		InitClass(ExtraClrScheme, CBaseFormatObject, AscDFH.historyitem_type_ExtraClrScheme);
		ExtraClrScheme.prototype.Refresh_RecalcData = function () {
		};
		ExtraClrScheme.prototype.setClrScheme = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ExtraClrScheme_SetClrScheme, this.clrScheme, pr));
			this.clrScheme = pr;
		};
		ExtraClrScheme.prototype.setClrMap = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ExtraClrScheme_SetClrMap, this.clrMap, pr));
			this.clrMap = pr;
		};
		ExtraClrScheme.prototype.createDuplicate = function () {
			var ret = new ExtraClrScheme();
			if (this.clrScheme) {
				ret.setClrScheme(this.clrScheme.createDuplicate())
			}
			if (this.clrMap) {
				ret.setClrMap(this.clrMap.createDuplicate());
			}
			return ret;
		};

		drawingConstructorsMap[AscDFH.historyitem_ExtraClrScheme_SetClrScheme] = ClrScheme;
		drawingConstructorsMap[AscDFH.historyitem_NvPr_AddExt] = CExtP;

		function FontCollection(fontScheme) {
			CBaseNoIdObject.call(this);
			this.latin = null;
			this.ea = null;
			this.cs = null;
			if (fontScheme) {
				this.setFontScheme(fontScheme);
			}
		}

		InitClass(FontCollection, CBaseNoIdObject, 0);
		FontCollection.prototype.Refresh_RecalcData = function () {
		};
		FontCollection.prototype.setFontScheme = function (fontScheme) {
			this.fontScheme = fontScheme;
		};
		FontCollection.prototype.setLatin = function (pr) {
			this.latin = pr;
			if (this.fontScheme)
				this.fontScheme.checkFromFontCollection(pr, this, FONT_REGION_LT);
		};
		FontCollection.prototype.setEA = function (pr) {
			this.ea = pr;
			if (this.fontScheme)
				this.fontScheme.checkFromFontCollection(pr, this, FONT_REGION_EA);
		};
		FontCollection.prototype.setCS = function (pr) {
			this.cs = pr;
			if (this.fontScheme)
				this.fontScheme.checkFromFontCollection(pr, this, FONT_REGION_CS);
		};
		FontCollection.prototype.Write_ToBinary = function (w) {
			writeString(w, this.latin);
			writeString(w, this.ea);
			writeString(w, this.cs);
		};
		FontCollection.prototype.Read_FromBinary = function (r) {
			this.latin = readString(r);
			this.ea = readString(r);
			this.cs = readString(r);

			if (this.fontScheme) {
				this.fontScheme.checkFromFontCollection(this.latin, this, FONT_REGION_LT);
				this.fontScheme.checkFromFontCollection(this.ea, this, FONT_REGION_EA);
				this.fontScheme.checkFromFontCollection(this.cs, this, FONT_REGION_CS);
			}
		};

		var FONT_REGION_LT = 0x00;
		var FONT_REGION_EA = 0x01;
		var FONT_REGION_CS = 0x02;

		function FontScheme() {
			CBaseNoIdObject.call(this)
			this.name = "";
			this.majorFont = new FontCollection(this);
			this.minorFont = new FontCollection(this);
			this.fontMap = {
				"+mj-lt": undefined,
				"+mj-ea": undefined,
				"+mj-cs": undefined,
				"+mn-lt": undefined,
				"+mn-ea": undefined,
				"+mn-cs": undefined,
				"majorAscii": undefined,
				"majorBidi": undefined,
				"majorEastAsia": undefined,
				"majorHAnsi": undefined,
				"minorAscii": undefined,
				"minorBidi": undefined,
				"minorEastAsia": undefined,
				"minorHAnsi": undefined
			};
			/**
			 * visio extension from tag vt:schemeID which is stored in a:extLst
			 * @type {string | null}
			 */
			this.schemeEnum = null;
		}

		InitClass(FontScheme, CBaseNoIdObject, 0);
		FontScheme.prototype.createDuplicate = function () {
			var oCopy = new FontScheme();
			oCopy.majorFont.setLatin(this.majorFont.latin);
			oCopy.majorFont.setEA(this.majorFont.ea);
			oCopy.majorFont.setCS(this.majorFont.cs);

			oCopy.minorFont.setLatin(this.minorFont.latin);
			oCopy.minorFont.setEA(this.minorFont.ea);
			oCopy.minorFont.setCS(this.minorFont.cs);
			return oCopy;
		};
		FontScheme.prototype.Refresh_RecalcData = function () {
		};
		FontScheme.prototype.Write_ToBinary = function (w) {
			this.majorFont.Write_ToBinary(w);
			this.minorFont.Write_ToBinary(w);
			writeString(w, this.name);
		};
		FontScheme.prototype.Read_FromBinary = function (r) {
			this.majorFont.Read_FromBinary(r);
			this.minorFont.Read_FromBinary(r);
			this.name = readString(r);
		};
		FontScheme.prototype.checkFromFontCollection = function (font, fontCollection, region) {
			if (fontCollection === this.majorFont) {
				switch (region) {
					case FONT_REGION_LT: {
						this.fontMap["+mj-lt"] = font;
						this.fontMap["majorAscii"] = font;
						this.fontMap["majorHAnsi"] = font;
						break;
					}
					case FONT_REGION_EA: {
						this.fontMap["+mj-ea"] = font;
						this.fontMap["majorEastAsia"] = font;
						break;
					}
					case FONT_REGION_CS: {
						this.fontMap["+mj-cs"] = font;
						this.fontMap["majorBidi"] = font;
						break;
					}
				}
			} else if (fontCollection === this.minorFont) {
				switch (region) {
					case FONT_REGION_LT: {
						this.fontMap["+mn-lt"] = font;
						this.fontMap["minorAscii"] = font;
						this.fontMap["minorHAnsi"] = font;
						break;
					}
					case FONT_REGION_EA: {
						this.fontMap["+mn-ea"] = font;
						this.fontMap["minorEastAsia"] = font;
						break;
					}
					case FONT_REGION_CS: {
						this.fontMap["+mn-cs"] = font;
						this.fontMap["minorBidi"] = font;
						break;
					}
				}
			}
		};
		FontScheme.prototype.checkFont = function (font) {
			if (g_oThemeFontsName[font]) {
				if (this.fontMap[font]) {
					return this.fontMap[font];
				} else if (this.fontMap["+mn-lt"]) {
					return this.fontMap["+mn-lt"];
				} else {
					return "Arial";
				}
			}
			return font;
		};
		FontScheme.prototype.getObjectType = function () {
			return AscDFH.historyitem_type_FontScheme;
		};
		FontScheme.prototype.setName = function (pr) {
			this.name = pr;
		};
		FontScheme.prototype.setMajorFont = function (pr) {
			this.majorFont = pr;
		};
		FontScheme.prototype.setMinorFont = function (pr) {
			this.minorFont = pr;
		};

		function FmtScheme() {
			CBaseNoIdObject.call(this);
			this.name = "";
			this.fillStyleLst = [];
			this.lnStyleLst = [];
			this.effectStyleLst = [];
			this.bgFillStyleLst = [];
		}

		InitClass(FmtScheme, CBaseNoIdObject, 0);
		FmtScheme.prototype.GetFillStyle = function (number, unicolor) {
			if (number >= 1 && number <= 999) {
				var ret = this.fillStyleLst[number - 1];
				if (!ret)
					return null;
				var ret2 = ret.createDuplicate();
				ret2.checkPhColor(unicolor);
				return ret2;
			} else if (number >= 1001) {
				var ret = this.bgFillStyleLst[number - 1001];
				if (!ret)
					return null;
				var ret2 = ret.createDuplicate();
				ret2.checkPhColor(unicolor);
				return ret2;
			}
			return null;
		};
		FmtScheme.prototype.GetOuterShdw = function (number) {
			if (this.effectStyleLst[number - 1]) {
				const effectStyle = this.effectStyleLst[number - 1];
				const effectLst = effectStyle.effectProperties && effectStyle.effectProperties.EffectLst;
				if (effectLst && effectLst.outerShdw) {
					return effectLst.outerShdw.createDuplicate();
				}
			}
			return null;
		};
		FmtScheme.prototype.Write_ToBinary = function (w) {
			writeString(w, this.name);
			var i;
			w.WriteLong(this.fillStyleLst.length);
			for (i = 0; i < this.fillStyleLst.length; ++i) {
				this.fillStyleLst[i].Write_ToBinary(w);
			}

			w.WriteLong(this.lnStyleLst.length);
			for (i = 0; i < this.lnStyleLst.length; ++i) {
				this.lnStyleLst[i].Write_ToBinary(w);
			}

			w.WriteLong(this.effectStyleLst.length);
			for (i = 0; i < this.effectStyleLst.length; ++i) {
				this.effectStyleLst[i].Write_ToBinary(w);
			}

			w.WriteLong(this.bgFillStyleLst.length);
			for (i = 0; i < this.bgFillStyleLst.length; ++i) {
				this.bgFillStyleLst[i].Write_ToBinary(w);
			}
		};
		FmtScheme.prototype.Read_FromBinary = function (r) {
			this.name = readString(r);
			var _len = r.GetLong(), i;
			for (i = 0; i < _len; ++i) {
				this.fillStyleLst[i] = new CUniFill();
				this.fillStyleLst[i].Read_FromBinary(r);
			}

			_len = r.GetLong();
			for (i = 0; i < _len; ++i) {
				this.lnStyleLst[i] = new CLn();
				this.lnStyleLst[i].Read_FromBinary(r);
			}

			_len = r.GetLong();
			for (i = 0; i < _len; ++i) {
				this.effectStyleLst[i] = new CEffectStyle();
				this.effectStyleLst[i].Read_FromBinary(r);
			}

			_len = r.GetLong();
			for (i = 0; i < _len; ++i) {
				this.bgFillStyleLst[i] = new CUniFill();
				this.bgFillStyleLst[i].Read_FromBinary(r);
			}
		};
		FmtScheme.prototype.createDuplicate = function () {
			var oCopy = new FmtScheme();
			oCopy.name = this.name;
			var i;
			for (i = 0; i < this.fillStyleLst.length; ++i) {
				oCopy.fillStyleLst[i] = this.fillStyleLst[i].createDuplicate();
			}
			for (i = 0; i < this.lnStyleLst.length; ++i) {
				oCopy.lnStyleLst[i] = this.lnStyleLst[i].createDuplicate();
			}

			for (i = 0; i < this.effectStyleLst.length; ++i) {
				oCopy.effectStyleLst[i] = this.effectStyleLst[i].createDuplicate();
			}

			for (i = 0; i < this.bgFillStyleLst.length; ++i) {
				oCopy.bgFillStyleLst[i] = this.bgFillStyleLst[i].createDuplicate();
			}
			return oCopy;
		};
		FmtScheme.prototype.setName = function (pr) {
			this.name = pr;
		};
		FmtScheme.prototype.addFillToStyleLst = function (pr) {
			this.fillStyleLst.push(pr);
		};
		FmtScheme.prototype.addLnToStyleLst = function (pr) {
			this.lnStyleLst.push(pr);
		};
		FmtScheme.prototype.addEffectToStyleLst = function (pr) {
			this.effectStyleLst.push(pr);
		};
		FmtScheme.prototype.addBgFillToStyleLst = function (pr) {
			this.bgFillStyleLst.push(pr);
		};
		FmtScheme.prototype.getImageFromBulletsMap = function(oImages) {};
		FmtScheme.prototype.getDocContentsWithImageBullets = function (arrContents) {};
		FmtScheme.prototype.getAllRasterImages = function(aImages) {
			for(let nIdx = 0; nIdx < this.fillStyleLst.length; ++nIdx) {
				let oUnifill = this.fillStyleLst[nIdx];
				let sRasterImageId = oUnifill && oUnifill.fill && oUnifill.fill.RasterImageId;
				if(sRasterImageId) {
					aImages.push(sRasterImageId);
				}
			}
			for(let nIdx = 0; nIdx < this.bgFillStyleLst.length; ++nIdx) {
				let oUnifill = this.bgFillStyleLst[nIdx];
				let sRasterImageId = oUnifill && oUnifill.fill && oUnifill.fill.RasterImageId;
				if(sRasterImageId) {
					aImages.push(sRasterImageId);
				}
			}
		};
		FmtScheme.prototype.Reassign_ImageUrls = function(oImageMap) {
			for(let nIdx = 0; nIdx < this.fillStyleLst.length; ++nIdx) {
				let oUnifill = this.fillStyleLst[nIdx];
				let sRasterImageId = oUnifill && oUnifill.fill && oUnifill.fill.RasterImageId;
				if(sRasterImageId && oImageMap[sRasterImageId]) {
					oUnifill.fill.RasterImageId = oImageMap[sRasterImageId]
				}
			}
			for(let nIdx = 0; nIdx < this.bgFillStyleLst.length; ++nIdx) {
				let oUnifill = this.bgFillStyleLst[nIdx];
				let sRasterImageId = oUnifill && oUnifill.fill && oUnifill.fill.RasterImageId;
				if(sRasterImageId && oImageMap[sRasterImageId]) {
					oUnifill.fill.RasterImageId = oImageMap[sRasterImageId]
				}
			}
		};
		
		function CEffectStyle() {
			CBaseFormatNoIdObject.call(this);
			this.effectProperties = null;
			this.scene3d = null;
			this.sp3d = null;
		}
		InitClass(CEffectStyle, CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

		CEffectStyle.prototype.Write_ToBinary = function (w) {
			writeObjectNoId(w, this.effectProperties);
			writeObjectNoId(w, this.scene3d);
			writeObjectNoId(w, this.sp3d);
		};
		CEffectStyle.prototype.Read_FromBinary = function (r) {
			this.setEffectPr(readObjectNoId(r));
			this.setScene3d(readObjectNoId(r));
			this.setSp3d(readObjectNoId(r));
		};
		CEffectStyle.prototype.setEffectPr = function (pr) {
			this.effectProperties = pr;
		};
		CEffectStyle.prototype.setScene3d = function (pr) {
			this.scene3d = pr;
		};
		CEffectStyle.prototype.setSp3d = function (pr) {
			this.sp3d = pr;
		};
		CEffectStyle.prototype.createDuplicate = function () {
			const copy = new CEffectStyle();
			if (this.effectProperties) {
				copy.setEffectPr(this.effectProperties.createDuplicate());
			}
			if (this.scene3d) {
				copy.setScene3d(this.scene3d.createDuplicate());
			}
			if (this.sp3d) {
				copy.setSp3d(this.sp3d.createDuplicate());
			}
			return copy;
		};
		CEffectStyle.prototype.readAttribute = null;
		CEffectStyle.prototype.privateWriteAttributes = null;
		CEffectStyle.prototype.readChild = function(nType, pReader) {
			switch (nType) {
				case 0:
					this.setEffectPr(pReader.ReadEffectProperties());
					break;
				case 1:
					this.setScene3d(new AscFormat.Scene3d());
					this.scene3d.fromPPTY(pReader);
					break;
				case 2:
					this.setSp3d(new AscFormat.Sp3d());
					this.sp3d.fromPPTY(pReader);
					break;
				default:
					return false;
			}
			return true;
		};
		CEffectStyle.prototype.writeChildren = function(pWriter) {
			var oEffectPr = this.effectProperties;
			if(oEffectPr)
			{
				if(oEffectPr.EffectLst)
				{
					pWriter.WriteRecord1(0, oEffectPr.EffectLst, pWriter.WriteEffectLst);
				}
				else if(oEffectPr.EffectDag)
				{
					pWriter.WriteRecord1(0, oEffectPr.EffectDag, pWriter.WriteEffectDag)
				}
			}
			this.writeRecord2(pWriter, 1, this.scene3d);
			this.writeRecord2(pWriter, 2, this.sp3d);
		};

		function ThemeElements(oTheme) {
			CBaseNoIdObject.call(this);
			this.theme = oTheme;
			this.clrScheme = new ClrScheme();
			this.fontScheme = new FontScheme();
			this.fmtScheme = new FmtScheme();

			/**
			 *
			 * @type {CThemeExt}
			 */
			this.themeExt = null;
		}

		InitClass(ThemeElements, CBaseNoIdObject, 0);

		/**
		 *
		 * @constructor
		 */
		function CTheme() {
			CBaseFormatObject.call(this);
			this.name = "";
			this.themeElements = new ThemeElements(this);
			this.spDef = null;
			this.lnDef = null;
			this.txDef = null;

			this.extraClrSchemeLst = [];

			this.isThemeOverride = false;

			// pointers
			this.presentation = null;
			this.clrMap = null;

		}

		InitClass(CTheme, CBaseFormatObject, 0);

		CTheme.prototype.notAllowedWithoutId = function () {
			return false;
		};
		CTheme.prototype.createDuplicate = function () {
			var oTheme = new CTheme();
			oTheme.setName(this.name);
			oTheme.setColorScheme(this.themeElements.clrScheme.createDuplicate());
			oTheme.setFontScheme(this.themeElements.fontScheme.createDuplicate());
			oTheme.setFormatScheme(this.themeElements.fmtScheme.createDuplicate());
			if (this.spDef) {
				oTheme.setSpDef(this.spDef.createDuplicate());
			}
			if (this.lnDef) {
				oTheme.setLnDef(this.lnDef.createDuplicate());
			}
			if (this.txDef) {
				oTheme.setTxDef(this.txDef.createDuplicate());
			}
			for (var i = 0; i < this.extraClrSchemeLst.length; ++i) {
				oTheme.addExtraClrSceme(this.extraClrSchemeLst[i].createDuplicate());
			}
			return oTheme;
		};
		CTheme.prototype.Document_Get_AllFontNames = function (AllFonts) {
			var font_scheme = this.themeElements.fontScheme;
			var major_font = font_scheme.majorFont;
			typeof major_font.latin === "string" && major_font.latin.length > 0 && (AllFonts[major_font.latin] = 1);
			typeof major_font.ea === "string" && major_font.ea.length > 0 && (AllFonts[major_font.ea] = 1);
			typeof major_font.cs === "string" && major_font.cs.length > 0 && (AllFonts[major_font.cs] = 1);
			var minor_font = font_scheme.minorFont;
			typeof minor_font.latin === "string" && minor_font.latin.length > 0 && (AllFonts[minor_font.latin] = 1);
			typeof minor_font.ea === "string" && minor_font.ea.length > 0 && (AllFonts[minor_font.ea] = 1);
			typeof minor_font.cs === "string" && minor_font.cs.length > 0 && (AllFonts[minor_font.cs] = 1);
		};
		CTheme.prototype.getOuterShdw = function (idx) {
			return this.themeElements.fmtScheme.GetOuterShdw(idx);
		};
		CTheme.prototype.getFillStyle = function (idx, unicolor) {
			if (idx === 0 || idx === 1000) {
				return AscFormat.CreateNoFillUniFill();
			}
			var ret;
			if (idx >= 1 && idx <= 999) {
				if (this.themeElements.fmtScheme.fillStyleLst[idx - 1]) {
					ret = this.themeElements.fmtScheme.fillStyleLst[idx - 1].createDuplicate();
					if (ret) {
						ret.checkPhColor(unicolor);
						return ret;
					}
				}
			} else if (idx >= 1001) {
				if (this.themeElements.fmtScheme.bgFillStyleLst[idx - 1001]) {
					ret = this.themeElements.fmtScheme.bgFillStyleLst[idx - 1001].createDuplicate();
					if (ret) {
						ret.checkPhColor(unicolor);
						return ret;
					}
				}
			}
			return CreateSolidFillRGBA(0, 0, 0, 255);
		};
		CTheme.prototype.getLnStyle = function (idx, unicolor) {
			if (idx === 0) {
				return AscFormat.CreateNoFillLine();
			}
			if (this.themeElements.fmtScheme.lnStyleLst[idx - 1]) {
				var ret = this.themeElements.fmtScheme.lnStyleLst[idx - 1].createDuplicate();
				if (ret.Fill) {
					ret.Fill.checkPhColor(unicolor);
				}
				return ret;
			}
			return new CLn();
		};
		/**
		 * Made for visio editor
		 * @memberof CTheme
		 * @param {number} idx
		 * @param unicolor
		 * @param {Boolean} getConnectorStyle
		 * @return {CFontProps}
		 */
		CTheme.prototype.getFontStyle = function (idx, unicolor, getConnectorStyle) {
			if (idx === 0 || isNaN(idx)) {
				AscCommon.consoleLog("idx getFontStyle argument is 0 or isNaN(idx) is true");
				return AscFormat.CreateNoFillLine();
			}
			let fontProp;
			if (this.themeElements.themeExt) {
				if (getConnectorStyle) {
					fontProp = this.themeElements.themeExt.fontStylesGroup.connectorFontStyles[idx - 1];
				} else {
					fontProp = this.themeElements.themeExt.fontStylesGroup.fontStyles[idx - 1];
				}
			}
			if (fontProp) {
				var ret = fontProp.createDuplicate();
				if (ret.fontPropsObject.color) {
					ret.fontPropsObject.color.checkPhColor(unicolor, false);
				}
				return ret;
			}
			AscCommon.consoleLog("no fontProp object found for idx. Return empty obj", idx);
			return new CFontProps();
		};
		/**
		 * Made for visio editor. Get line end props - lineStyles extension.
		 * @memberof CTheme
		 * @param {number} idx
		 * @param {Boolean} getConnectorStyle
		 * @return {CLineStyle} lineStyle
		 */
		CTheme.prototype.getLineEndStyle = function (idx, getConnectorStyle) {
			if (idx === 0 || isNaN(idx)) {
				AscCommon.consoleLog("idx getFontStyle argument is 0 or isNaN(idx) is true");
				return new CLineStyle();
			}
			let lineEndProp;

			if (this.themeElements.themeExt) {
				// not using idx - 1. Seems like visio bug here. See file https://disk.yandex.ru/d/OQVR9m1U255B1Q
				if (getConnectorStyle) {
					lineEndProp = this.themeElements.themeExt.lineStyles.fmtConnectorSchemeLineStyles[idx];
				} else {
					lineEndProp = this.themeElements.themeExt.lineStyles.fmtSchemeLineStyles[idx];
				}
			} else {
				AscCommon.consoleLog("no this.themeElements.themeExt (this = CTheme) found with info about line endings. " +
					"Its ok sometimes. Set end:4 for default lineEndProp");
				lineEndProp = new CLineStyle();
				lineEndProp.lineEx = {rndg: 0, start: 0, startSize: 2, end: 4, endSize: 2, pattern: 1};
			}

			if (lineEndProp) {
				var ret = lineEndProp.createDuplicate();
				return ret;
			}
			AscCommon.consoleLog("no fontProp object found for idx. Its ok sometimes." +
				" Return new CLineStyle()", idx);
			return new CLineStyle();
		};

		/**
		 * Made for visio editor. Get fill pattern from <vt:fillStyles...
		 * @memberof CTheme
		 * @param {number} idx
		 * @return {{pattern: number}} lineStyle
		 */
		CTheme.prototype.getFillProp = function (idx) {
			if (idx === 0 || isNaN(idx)) {
				AscCommon.consoleLog("idx getFillProp argument is 0 or isNaN(idx) is true");
				return {pattern: 1}; // solid
			}

			let fillProp;
			if (this.themeElements.themeExt) {
				fillProp = this.themeElements.themeExt.fillStyles[idx - 1];
			}

			if (fillProp) {
				var ret = {pattern: fillProp.pattern};
				return ret;
			}
			AscCommon.consoleLog("no FillProp object found for idx. Its ok sometimes." +
				" Return default FillProp", idx);
			return {pattern: 1};
		};

		CTheme.prototype.getExtraClrScheme = function (sName) {
			for (var i = 0; i < this.extraClrSchemeLst.length; ++i) {
				if (this.extraClrSchemeLst[i].clrScheme && this.extraClrSchemeLst[i].clrScheme.name === sName) {
					return this.extraClrSchemeLst[i].clrScheme.createDuplicate();
				}
			}
			return null;
		};
		CTheme.prototype.changeColorScheme = function (clrScheme) {
			var oCurClrScheme = this.themeElements.clrScheme;
			this.setColorScheme(clrScheme);
			var oOldAscColorScheme = AscCommon.getAscColorScheme(oCurClrScheme, this),
				aExtraAscClrSchemes = this.getExtraAscColorSchemes();
			var oNewAscColorScheme = AscCommon.getAscColorScheme(clrScheme, this);
			if (AscCommon.getIndexColorSchemeInArray(AscCommon.g_oUserColorScheme, oOldAscColorScheme) === -1) {
				if (AscCommon.getIndexColorSchemeInArray(aExtraAscClrSchemes, oOldAscColorScheme) === -1) {
					var oExtraClrScheme = new ExtraClrScheme();
					if (this.clrMap) {
						oExtraClrScheme.setClrMap(this.clrMap.createDuplicate());
					}
					oExtraClrScheme.setClrScheme(oCurClrScheme.createDuplicate());
					this.addExtraClrSceme(oExtraClrScheme, 0);
					aExtraAscClrSchemes = this.getExtraAscColorSchemes();
				}
			}
			var nIndex = AscCommon.getIndexColorSchemeInArray(aExtraAscClrSchemes, oNewAscColorScheme);
			if (nIndex > -1) {
				this.removeExtraClrScheme(nIndex);
			}
		};
		CTheme.prototype.getExtraAscColorSchemes = function () {
			var asc_color_scheme;
			var aCustomSchemes = [];
			var _extra = this.extraClrSchemeLst;
			var _count = _extra.length;
			for (var i = 0; i < _count; ++i) {
				var _scheme = _extra[i].clrScheme;
				asc_color_scheme = AscCommon.getAscColorScheme(_scheme, this);
				aCustomSchemes.push(asc_color_scheme);
			}
			return aCustomSchemes;
		};
		CTheme.prototype.setColorScheme = function (clrScheme) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ThemeSetColorScheme, this.themeElements.clrScheme, clrScheme));
			this.themeElements.clrScheme = clrScheme;
		};
		CTheme.prototype.setFontScheme = function (fontScheme) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ThemeSetFontScheme, this.themeElements.fontScheme, fontScheme));
			this.themeElements.fontScheme = fontScheme;
		};

		/**
		 * made for Visio editor
		 * @return {FontScheme}
		 */
		CTheme.prototype.getFontScheme = function () {
			return this.themeElements.fontScheme;
		};

		CTheme.prototype.setThemeExt = function (themeExt) {
			this.themeElements.themeExt = themeExt;
		};
		CTheme.prototype.changeFontScheme = function (fontScheme) {
			this.setFontScheme(fontScheme);
			let aIndexes = this.GetAllSlideIndexes();
			let aSlides = this.GetPresentationSlides();
			if(aIndexes && aSlides) {
				for (let i = 0; i < aIndexes.length; ++i) {
					aSlides[aIndexes[i]] && aSlides[aIndexes[i]].checkSlideTheme();
				}
			}
		};
		CTheme.prototype.setFormatScheme = function (fmtScheme) {
			AscCommon.History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ThemeSetFmtScheme, this.themeElements.fmtScheme, fmtScheme));
			this.themeElements.fmtScheme = fmtScheme;
		};
		CTheme.prototype.setName = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_ThemeSetName, this.name, pr));
			this.name = pr;
		};
		CTheme.prototype.setIsThemeOverride = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_ThemeSetIsThemeOverride, this.isThemeOverride, pr));
			this.isThemeOverride = pr;
		};
		CTheme.prototype.setSpDef = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ThemeSetSpDef, this.spDef, pr));
			this.spDef = pr;
		};
		CTheme.prototype.setLnDef = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ThemeSetLnDef, this.lnDef, pr));
			this.lnDef = pr;
		};
		CTheme.prototype.setTxDef = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ThemeSetTxDef, this.txDef, pr));
			this.txDef = pr;
		};
		CTheme.prototype.addExtraClrSceme = function (pr, idx) {
			var pos;
			if (AscFormat.isRealNumber(idx))
				pos = idx;
			else
				pos = this.extraClrSchemeLst.length;
			AscCommon.History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_ThemeAddExtraClrScheme, pos, [pr], true));
			this.extraClrSchemeLst.splice(pos, 0, pr);
		};
		CTheme.prototype.removeExtraClrScheme = function (idx) {
			if (idx > -1 && idx < this.extraClrSchemeLst.length) {
				AscCommon.History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_ThemeRemoveExtraClrScheme, idx, this.extraClrSchemeLst.splice(idx, 1), false));
			}
		};
		CTheme.prototype.GetLogicDocument = function() {
			let oRet = typeof editor !== "undefined" &&
				editor.WordControl &&
				editor.WordControl.m_oLogicDocument;
			return AscCommon.isRealObject(oRet) ? oRet : null;
		};
		CTheme.prototype.GetWordDrawingObjects = function () {
			let oLogicDocument = this.GetLogicDocument();
			let oRet = oLogicDocument && oLogicDocument.DrawingObjects;
			return AscCommon.isRealObject(oRet) ? oRet : null;
		};
		CTheme.prototype.GetPresentationSlides = function () {
			let oLogicDocument = this.GetLogicDocument();
			if(oLogicDocument && Array.isArray(oLogicDocument.Slides)) {
				return oLogicDocument.Slides;
			}
			return null;
		};
		CTheme.prototype.GetPresentation = function () {
			let oLogicDoc = this.GetLogicDocument();
			if(AscCommonSlide.CPresentation && oLogicDoc instanceof AscCommonSlide.CPresentation) {
				return oLogicDoc;
			}
			return null;
		};
		CTheme.prototype.GetAllSlideIndexes = function () {
			let oPresentation = this.GetLogicDocument();
			let aSlides = this.GetPresentationSlides();
			if(oPresentation && aSlides) {
				let aIndexes = [];
				for(let nSlide = 0; nSlide < aSlides.length; ++nSlide) {
					let oSlide = aSlides[nSlide];
					let oTheme = oSlide.getTheme();
					if(oTheme === this) {
						aIndexes.push(nSlide);
					}
				}
				return aIndexes;
			}
			return null;
		};
		CTheme.prototype.Refresh_RecalcData = function (oData) {
			if (oData) {
				if (oData.Type === AscDFH.historyitem_ThemeSetColorScheme) {
					let oWordGraphicObject = this.GetWordDrawingObjects();
					if (oWordGraphicObject) {
						AscCommon.History.RecalcData_Add({All: true});
						let aDrawings = oWordGraphicObject.drawingObjects;
						for (let nDrawing = 0; nDrawing < aDrawings.length; ++nDrawing) {
							let oGrObject = aDrawings[nDrawing].GraphicObj;
							if (oGrObject) {
								oGrObject.handleUpdateFill();
								oGrObject.handleUpdateLn();
							}
						}
						let oApi = oWordGraphicObject.document.Api;
						oApi.chartPreviewManager.clearPreviews();
						oApi.textArtPreviewManager.clear();
					}
					else {
						let oPresentation = this.GetPresentation();
						if(oPresentation) {
							oPresentation.bNeedUpdateThemes = true;
							let oThemedObjects = oPresentation.GetSlideObjectsWithTheme(this);
							for(let nIdx = 0; nIdx < oThemedObjects.masters.length; ++nIdx) {
								oThemedObjects.masters[nIdx].checkSlideTheme();
							}
							for(let nIdx = 0; nIdx < oThemedObjects.layouts.length; ++nIdx) {
								oThemedObjects.layouts[nIdx].checkSlideTheme();
							}
							for(let nIdx = 0; nIdx < oThemedObjects.slides.length; ++nIdx) {
								oThemedObjects.slides[nIdx].checkSlideTheme();
							}
							AscCommon.History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, Object: this});
						}
					}
				}
				else if(oData.Type === AscDFH.historyitem_ThemeSetFontScheme) {
					let oPresentation = this.GetLogicDocument();
					let aSlideIndexes = this.GetAllSlideIndexes();
					if(oPresentation && aSlideIndexes && aSlideIndexes.length > 0) {
						oPresentation.Refresh_RecalcData2({Type: AscDFH.historyitem_ThemeSetFontScheme, aIndexes: aSlideIndexes});
					}
				} else if (oData.Type === AscDFH.historyitem_ThemeSetName) {
					let oPresentation = this.GetPresentation();
					if (oPresentation) {
						oPresentation.Refresh_RecalcData2({Type: AscDFH.historyitem_ThemeSetName});
					}
				}
			}
		};
		CTheme.prototype.getAllRasterImages = function(aImages) {
			if(this.themeElements && this.themeElements.fmtScheme) {
				this.themeElements.fmtScheme.getAllRasterImages(aImages);
			}
		};
		CTheme.prototype.getImageFromBulletsMap = function(oImages) {};
		CTheme.prototype.getDocContentsWithImageBullets = function (arrContents) {};
		CTheme.prototype.Reassign_ImageUrls = function(images_rename) {
			if(this.themeElements && this.themeElements.fmtScheme) {
				let aImages = [];
				this.themeElements.fmtScheme.getAllRasterImages(aImages);
				let bReassign = false;
				for(let nImage = 0; nImage < aImages.length; ++nImage) {
					if(images_rename[aImages[nImage]]) {
						bReassign = true;
						break;
					}
				}
				if(bReassign) {
					let oNewFmtScheme = this.themeElements.fmtScheme.createDuplicate();
					oNewFmtScheme.Reassign_ImageUrls(images_rename);
					this.setFormatScheme(oNewFmtScheme);
				}
			}
		};
		CTheme.prototype.getObjectType = function () {
			return AscDFH.historyitem_type_Theme;
		};
		CTheme.prototype.Write_ToBinary2 = function (w) {
			w.WriteLong(AscDFH.historyitem_type_Theme);
			w.WriteString2(this.Id);
		};
		CTheme.prototype.Read_FromBinary2 = function (r) {
			this.Id = r.GetString2();
		};

		/**
		 * @memberOf CTheme
		 * @return {boolean}
		 */
		CTheme.prototype.isVariationClrSchemeLstExists = function() {
			let clrScheme = this.themeElements.clrScheme;
			let variationClrSchemeLst = clrScheme && clrScheme.clrSchemeExtLst
				&& clrScheme.clrSchemeExtLst.variationClrSchemeLst;
			return variationClrSchemeLst && variationClrSchemeLst.length > 0;
		}

		/**
		 * @memberOf CTheme
		 * @param variationIndex
		 * @param colorIndex
		 * @return {CUniColor|null}
		 */
		CTheme.prototype.getVariationClrSchemeColor = function (variationIndex, colorIndex) {
			let clrScheme = this.themeElements.clrScheme;
			if (clrScheme.clrSchemeExtLst) {
				let variationClrSchemeLst = clrScheme.clrSchemeExtLst.variationClrSchemeLst;
				if (variationClrSchemeLst) {
					let variation = variationClrSchemeLst[variationIndex];
					if (variation) {
						let colorObj = variation.varColor[colorIndex];
						if (colorObj) {
							return colorObj.unicolor;
						}
					}
				}
			}
			return null;
		};
		/**
		 * @param variationIndex
		 * @param styleIndex
		 * @return {CVarStyle|null}
		 */
		CTheme.prototype.getVariationStyleScheme = function (variationIndex, styleIndex) {
			let themeExt = this.themeElements.themeExt;
			if (themeExt && themeExt.variationStyleSchemeLst[variationIndex]) {
				return themeExt.variationStyleSchemeLst[variationIndex].varStyle[styleIndex] || null;
			}
			return null;
		};
// ----------------------------------

// CSLD -----------------------------

		function HF() {
			CBaseFormatObject.call(this);
			this.dt = null;
			this.ftr = null;
			this.hdr = null;
			this.sldNum = null;
		}

		InitClass(HF, CBaseFormatObject, AscDFH.historyitem_type_HF);
		HF.prototype.Refresh_RecalcData = function () {
		};
		HF.prototype.createDuplicate = function () {
			var ret = new HF();
			if (ret.dt !== this.dt) {
				ret.setDt(this.dt);
			}
			if (ret.ftr !== this.ftr) {
				ret.setFtr(this.ftr);
			}
			if (ret.hdr !== this.hdr) {
				ret.setHdr(this.hdr);
			}
			if (ret.sldNum !== this.sldNum) {
				ret.setSldNum(this.sldNum);
			}
			return ret;
		};
		HF.prototype.setDt = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_HF_SetDt, this.dt, pr));
			this.dt = pr;
		};
		HF.prototype.setFtr = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_HF_SetFtr, this.ftr, pr));
			this.ftr = pr;
		};
		HF.prototype.setHdr = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_HF_SetHdr, this.hdr, pr));
			this.hdr = pr;
		};
		HF.prototype.setSldNum = function (pr) {
			AscCommon.History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_HF_SetSldNum, this.sldNum, pr));
			this.sldNum = pr;
		};
		HF.prototype.applySettings = function (oSettings) {
			if(!oSettings) return;
			if (oSettings.get_ShowSlideNum()) {
				if (this.sldNum !== null) {
					this.setSldNum(null);
				}
			}
			else {
				if (this.sldNum !== false) {
					this.setSldNum(false);
				}
			}

			if (oSettings.get_ShowFooter()) {
				if (this.ftr !== null) {
					this.setFtr(null);
				}
			}
			else {
				if (this.ftr !== false) {
					this.setFtr(false);
				}
			}

			if (oSettings.get_ShowHeader()) {
				if (this.hdr !== null) {
					this.setHdr(null);
				}
			}
			else {
				if (this.hdr !== false) {
					this.setHdr(false);
				}
			}

			if (oSettings.get_ShowDateTime()) {
				if (this.dt !== null) {
					this.setDt(null);
				}
			}
			else {
				if (this.dt !== false) {
					this.setDt(false);
				}
			}
		};

		function CBgPr() {
			CBaseNoIdObject.call(this)
			this.Fill = null;
			this.shadeToTitle = false;
			this.EffectProperties = null;
		}

		InitClass(CBgPr, CBaseNoIdObject, 0);
		CBgPr.prototype.merge = function (bgPr) {
			if (this.Fill == null) {
				this.Fill = new CUniFill();
				if (bgPr.Fill != null) {
					this.Fill.merge(bgPr.Fill)
				}
			}
		};
		CBgPr.prototype.createFullCopy = function () {
			var _copy = new CBgPr();
			if (this.Fill != null) {
				_copy.Fill = this.Fill.createDuplicate();
			}
			_copy.shadeToTitle = this.shadeToTitle;
			return _copy;
		};
		CBgPr.prototype.setFill = function (pr) {
			this.Fill = pr;
		};
		CBgPr.prototype.setShadeToTitle = function (pr) {
			this.shadeToTitle = pr;
		};
		CBgPr.prototype.Write_ToBinary = function (w) {
			w.WriteBool(isRealObject(this.Fill));
			if (isRealObject(this.Fill)) {
				this.Fill.Write_ToBinary(w);
			}
			w.WriteBool(this.shadeToTitle);
		};
		CBgPr.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				this.Fill = new CUniFill();
				this.Fill.Read_FromBinary(r);
			}
			this.shadeToTitle = r.GetBool();
		};
		CBgPr.prototype.checkBlipFillRasterImage = function (images) {
			let fill_image_id = checkUniFillRasterImageId(this.Fill);
			if (fill_image_id !== null)
				images.push(fill_image_id);
		};



		function CBg() {
			CBaseNoIdObject.call(this);
			this.bwMode = null;
			this.bgPr = null;
			this.bgRef = null;
		}

		InitClass(CBg, CBaseNoIdObject, 0);
		CBg.prototype.setBwMode = function (pr) {
			this.bwMode = pr;
		};
		CBg.prototype.setBgPr = function (pr) {
			this.bgPr = pr;
		};
		CBg.prototype.setBgRef = function (pr) {
			this.bgRef = pr;
		};
		CBg.prototype.merge = function (bg) {
			if (this.bgPr == null) {
				this.bgPr = new CBgPr();
				if (bg.bgPr != null) {
					this.bgPr.merge(bg.bgPr);
				}
			}
		};
		CBg.prototype.createFullCopy = function () {
			var _copy = new CBg();
			_copy.bwMode = this.bwMode;
			if (this.bgPr != null) {
				_copy.bgPr = this.bgPr.createFullCopy();
			}
			if (this.bgRef != null) {
				_copy.bgRef = this.bgRef.createDuplicate();
			}
			return _copy;
		};
		CBg.prototype.Write_ToBinary = function (w) {
			w.WriteBool(isRealObject(this.bgPr));
			if (isRealObject(this.bgPr)) {
				this.bgPr.Write_ToBinary(w);
			}
			w.WriteBool(isRealObject(this.bgRef));
			if (isRealObject(this.bgRef)) {
				this.bgRef.Write_ToBinary(w);
			}
		};
		CBg.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				this.bgPr = new CBgPr();
				this.bgPr.Read_FromBinary(r);
			}
			if (r.GetBool()) {
				this.bgRef = new StyleRef();
				this.bgRef.Read_FromBinary(r);
			}
		};

		function CSld(parent) {
			CBaseNoIdObject.call(this);
			this.name = "";
			this.Bg = null;
			this.spTree = [];//new GroupShape();
			this.parent = parent;
		}

		InitClass(CSld, CBaseNoIdObject, 0);
		CSld.prototype.removeAllInks = function () {
			const oController = this.parent && this.parent.graphicObjects;
			if (!oController) {
				return;
			}

			const arrInks = this.getAllInks();
			oController.removeAllInks(arrInks);
		};
		CSld.prototype.getAllInks = function (arrInks) {
			arrInks = arrInks || [];
			const arrSpTree = this.spTree;
			for (let i = arrSpTree.length - 1; i >= 0; i -= 1) {
				const oDrawing = arrSpTree[i];
				if (oDrawing.isInk() || oDrawing.isHaveOnlyInks()) {
					arrInks.push(oDrawing);
				} else {
					oDrawing.getAllInks(arrInks);
				}
			}
			return arrInks;
		};
		CSld.prototype.getObjectsNamesPairs = function () {
			var aPairs = [];
			var aSpTree = this.spTree;
			for (var nSp = 0; nSp < aSpTree.length; ++nSp) {
				var oSp = aSpTree[nSp];
				if (!oSp.isEmptyPlaceholder()) {
					aPairs.push({object: oSp, name: oSp.getObjectName()});
				}
			}
			return aPairs;
		};
		CSld.prototype.getObjectsNames = function () {
			var aPairs = this.getObjectsNamesPairs();
			var aNames = [];
			for (var nPair = 0; nPair < aPairs.length; ++nPair) {
				var oPair = aPairs[nPair];
				aNames.push(oPair.name);
			}
			return aNames;
		};
		CSld.prototype.getObjectByName = function (sName) {
			var aSpTree = this.spTree;
			for (var nSp = 0; nSp < aSpTree.length; ++nSp) {
				var oSp = aSpTree[nSp];
				if (oSp.getObjectName() === sName) {
					return oSp;
				}
			}
			return null;
		};
		CSld.prototype.forEachSp = function(fCallback) {
			for(let nSp = 0; nSp < this.spTree.length; ++nSp) {
				fCallback(this.spTree[nSp]);
			}
		};
		CSld.prototype.handleAllContents = function(fCallback) {
			this.forEachSp(function(oSp) {
				if (oSp.handleAllContents) {
					oSp.handleAllContents(fCallback);
				}
			});
		};
		CSld.prototype.refreshAllContentsFields = function(bNoHistory) {
			let bOldUpdate = AscCommon.History.RecalculateData.Update;
			if(bNoHistory) {
				AscCommon.History.RecalculateData.Update = false;
			}
			this.handleAllContents(RefreshContentAllFields);
			if(bNoHistory) {
				AscCommon.History.RecalculateData.Update = bOldUpdate;
			}
		};

		function RefreshContentAllFields(oContent) {
			if(!oContent) {
				return;
			}
			if(!oContent.RecalcAllFields) {
				return;
			}
			oContent.RecalcAllFields();
		}


		function CSpTree(oSlideObject) {
			CBaseNoIdObject.call(this);
			this.spTree = [];
			this.slideObject = oSlideObject;

		}

		InitClass(CSpTree, CBaseNoIdObject, 0);

// ----------------------------------

// MASTERSLIDE ----------------------


		function CTextStyles() {
			CBaseNoIdObject.call(this);
			this.titleStyle = null;
			this.bodyStyle = null;
			this.otherStyle = null;
		}

		InitClass(CTextStyles, CBaseNoIdObject, 0);
		CTextStyles.prototype.getStyleByPhType = function (phType) {
			switch (phType) {
				case AscFormat.phType_ctrTitle:
				case AscFormat.phType_title: {
					return this.titleStyle;
				}
				case AscFormat.phType_body:
				case AscFormat.phType_subTitle:
				case AscFormat.phType_obj:
				case null: {
					return this.bodyStyle;
				}
				default: {
					break;
				}
			}
			return this.otherStyle;
		};
		CTextStyles.prototype.createDuplicate = function () {
			var ret = new CTextStyles();
			if (isRealObject(this.titleStyle)) {
				ret.titleStyle = this.titleStyle.createDuplicate();
			}

			if (isRealObject(this.bodyStyle)) {
				ret.bodyStyle = this.bodyStyle.createDuplicate();
			}

			if (isRealObject(this.otherStyle)) {
				ret.otherStyle = this.otherStyle.createDuplicate();
			}
			return ret;
		};
		CTextStyles.prototype.Refresh_RecalcData = function () {
		};
		CTextStyles.prototype.Write_ToBinary = function (w) {
			w.WriteBool(isRealObject(this.titleStyle));
			if (isRealObject(this.titleStyle)) {
				this.titleStyle.Write_ToBinary(w);
			}


			w.WriteBool(isRealObject(this.bodyStyle));
			if (isRealObject(this.bodyStyle)) {
				this.bodyStyle.Write_ToBinary(w);
			}

			w.WriteBool(isRealObject(this.otherStyle));
			if (isRealObject(this.otherStyle)) {
				this.otherStyle.Write_ToBinary(w);
			}
		};
		CTextStyles.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				this.titleStyle = new TextListStyle();
				this.titleStyle.Read_FromBinary(r);
			} else {
				this.titleStyle = null;
			}


			if (r.GetBool()) {
				this.bodyStyle = new TextListStyle();
				this.bodyStyle.Read_FromBinary(r);
			} else {
				this.bodyStyle = null;
			}

			if (r.GetBool()) {
				this.otherStyle = new TextListStyle();
				this.otherStyle.Read_FromBinary(r);
			} else {
				this.otherStyle = null;
			}
		};
		CTextStyles.prototype.Document_Get_AllFontNames = function (AllFonts) {
			if (this.titleStyle) {
				this.titleStyle.Document_Get_AllFontNames(AllFonts);
			}
			if (this.bodyStyle) {
				this.bodyStyle.Document_Get_AllFontNames(AllFonts);
			}
			if (this.otherStyle) {
				this.otherStyle.Document_Get_AllFontNames(AllFonts);
			}
		};

//---------------------------

// SLIDELAYOUT ----------------------

//Layout types
		var nSldLtTBlank = 0; // Blank ))
		var nSldLtTChart = 1; //Chart)
		var nSldLtTChartAndTx = 2; //( Chart and Text ))
		var nSldLtTClipArtAndTx = 3; //Clip Art and Text)
		var nSldLtTClipArtAndVertTx = 4; //Clip Art and Vertical Text)
		var nSldLtTCust = 5; // Custom ))
		var nSldLtTDgm = 6; //Diagram ))
		var nSldLtTFourObj = 7; //Four Objects)
		var nSldLtTMediaAndTx = 8; // ( Media and Text ))
		var nSldLtTObj = 9; //Title and Object)
		var nSldLtTObjAndTwoObj = 10; //Object and Two Object)
		var nSldLtTObjAndTx = 11; // ( Object and Text ))
		var nSldLtTObjOnly = 12; //Object)
		var nSldLtTObjOverTx = 13; // ( Object over Text))
		var nSldLtTObjTx = 14; //Title, Object, and Caption)
		var nSldLtTPicTx = 15; //Picture and Caption)
		var nSldLtTSecHead = 16; //Section Header)
		var nSldLtTTbl = 17; // ( Table ))
		var nSldLtTTitle = 18; // ( Title ))
		var nSldLtTTitleOnly = 19; // ( Title Only ))
		var nSldLtTTwoColTx = 20; // ( Two Column Text ))
		var nSldLtTTwoObj = 21; //Two Objects)
		var nSldLtTTwoObjAndObj = 22; //Two Objects and Object)
		var nSldLtTTwoObjAndTx = 23; //Two Objects and Text)
		var nSldLtTTwoObjOverTx = 24; //Two Objects over Text)
		var nSldLtTTwoTxTwoObj = 25; //Two Text and Two Objects)
		var nSldLtTTx = 26; // ( Text ))
		var nSldLtTTxAndChart = 27; // ( Text and Chart ))
		var nSldLtTTxAndClipArt = 28; //Text and Clip Art)
		var nSldLtTTxAndMedia = 29; // ( Text and Media ))
		var nSldLtTTxAndObj = 30; // ( Text and Object ))
		var nSldLtTTxAndTwoObj = 31; //Text and Two Objects)
		var nSldLtTTxOverObj = 32; // ( Text over Object))
		var nSldLtTVertTitleAndTx = 33; //Vertical Title and Text)
		var nSldLtTVertTitleAndTxOverChart = 34; //Vertical Title and Text Over Chart)
		var nSldLtTVertTx = 35; //Vertical Text)

		AscFormat.nSldLtTBlank = nSldLtTBlank; // Blank ))
		AscFormat.nSldLtTChart = nSldLtTChart; //Chart)
		AscFormat.nSldLtTChartAndTx = nSldLtTChartAndTx; //( Chart and Text ))
		AscFormat.nSldLtTClipArtAndTx = nSldLtTClipArtAndTx; //Clip Art and Text)
		AscFormat.nSldLtTClipArtAndVertTx = nSldLtTClipArtAndVertTx; //Clip Art and Vertical Text)
		AscFormat.nSldLtTCust = nSldLtTCust; // Custom ))
		AscFormat.nSldLtTDgm = nSldLtTDgm; //Diagram ))
		AscFormat.nSldLtTFourObj = nSldLtTFourObj; //Four Objects)
		AscFormat.nSldLtTMediaAndTx = nSldLtTMediaAndTx; // ( Media and Text ))
		AscFormat.nSldLtTObj = nSldLtTObj; //Title and Object)
		AscFormat.nSldLtTObjAndTwoObj = nSldLtTObjAndTwoObj; //Object and Two Object)
		AscFormat.nSldLtTObjAndTx = nSldLtTObjAndTx; // ( Object and Text ))
		AscFormat.nSldLtTObjOnly = nSldLtTObjOnly; //Object)
		AscFormat.nSldLtTObjOverTx = nSldLtTObjOverTx; // ( Object over Text))
		AscFormat.nSldLtTObjTx = nSldLtTObjTx; //Title, Object, and Caption)
		AscFormat.nSldLtTPicTx = nSldLtTPicTx; //Picture and Caption)
		AscFormat.nSldLtTSecHead = nSldLtTSecHead; //Section Header)
		AscFormat.nSldLtTTbl = nSldLtTTbl; // ( Table ))
		AscFormat.nSldLtTTitle = nSldLtTTitle; // ( Title ))
		AscFormat.nSldLtTTitleOnly = nSldLtTTitleOnly; // ( Title Only ))
		AscFormat.nSldLtTTwoColTx = nSldLtTTwoColTx; // ( Two Column Text ))
		AscFormat.nSldLtTTwoObj = nSldLtTTwoObj; //Two Objects)
		AscFormat.nSldLtTTwoObjAndObj = nSldLtTTwoObjAndObj; //Two Objects and Object)
		AscFormat.nSldLtTTwoObjAndTx = nSldLtTTwoObjAndTx; //Two Objects and Text)
		AscFormat.nSldLtTTwoObjOverTx = nSldLtTTwoObjOverTx; //Two Objects over Text)
		AscFormat.nSldLtTTwoTxTwoObj = nSldLtTTwoTxTwoObj; //Two Text and Two Objects)
		AscFormat.nSldLtTTx = nSldLtTTx; // ( Text ))
		AscFormat.nSldLtTTxAndChart = nSldLtTTxAndChart; // ( Text and Chart ))
		AscFormat.nSldLtTTxAndClipArt = nSldLtTTxAndClipArt; //Text and Clip Art)
		AscFormat.nSldLtTTxAndMedia = nSldLtTTxAndMedia; // ( Text and Media ))
		AscFormat.nSldLtTTxAndObj = nSldLtTTxAndObj; // ( Text and Object ))
		AscFormat.nSldLtTTxAndTwoObj = nSldLtTTxAndTwoObj; //Text and Two Objects)
		AscFormat.nSldLtTTxOverObj = nSldLtTTxOverObj; // ( Text over Object))
		AscFormat.nSldLtTVertTitleAndTx = nSldLtTVertTitleAndTx; //Vertical Title and Text)
		AscFormat.nSldLtTVertTitleAndTxOverChart = nSldLtTVertTitleAndTxOverChart; //Vertical Title and Text Over Chart)
		AscFormat.nSldLtTVertTx = nSldLtTVertTx; //Vertical Text)

		var _ph_multiplier = 4;

		var _weight_body = 9;
		var _weight_chart = 5;
		var _weight_clipArt = 2;
		var _weight_ctrTitle = 11;
		var _weight_dgm = 4;
		var _weight_media = 3;
		var _weight_obj = 8;
		var _weight_pic = 7;
		var _weight_subTitle = 10;
		var _weight_tbl = 6;
		var _weight_title = 11;

		var _ph_summ_blank = 0;
		var _ph_summ_chart = Math.pow(_ph_multiplier, _weight_title) + Math.pow(_ph_multiplier, _weight_chart);
		var _ph_summ_chart_and_tx = Math.pow(_ph_multiplier, _weight_title) + Math.pow(_ph_multiplier, _weight_chart) + Math.pow(_ph_multiplier, _weight_body);
		var _ph_summ_dgm = Math.pow(_ph_multiplier, _weight_title) + Math.pow(_ph_multiplier, _weight_dgm);
		var _ph_summ_four_obj = Math.pow(_ph_multiplier, _weight_title) + 4 * Math.pow(_ph_multiplier, _weight_obj);
		var _ph_summ__media_and_tx = Math.pow(_ph_multiplier, _weight_title) + Math.pow(_ph_multiplier, _weight_media) + Math.pow(_ph_multiplier, _weight_body);
		var _ph_summ__obj = Math.pow(_ph_multiplier, _weight_title) + Math.pow(_ph_multiplier, _weight_obj);
		var _ph_summ__obj_and_two_obj = Math.pow(_ph_multiplier, _weight_title) + 3 * Math.pow(_ph_multiplier, _weight_obj);
		var _ph_summ__obj_and_tx = Math.pow(_ph_multiplier, _weight_title) + Math.pow(_ph_multiplier, _weight_obj) + Math.pow(_ph_multiplier, _weight_body);
		var _ph_summ__obj_only = Math.pow(_ph_multiplier, _weight_obj);
		var _ph_summ__pic_tx = Math.pow(_ph_multiplier, _weight_title) + Math.pow(_ph_multiplier, _weight_pic) + Math.pow(_ph_multiplier, _weight_body);
		var _ph_summ__sec_head = Math.pow(_ph_multiplier, _weight_title) + Math.pow(_ph_multiplier, _weight_subTitle);
		var _ph_summ__tbl = Math.pow(_ph_multiplier, _weight_title) + Math.pow(_ph_multiplier, _weight_tbl);
		var _ph_summ__title_only = Math.pow(_ph_multiplier, _weight_title);
		var _ph_summ__two_col_tx = Math.pow(_ph_multiplier, _weight_title) + 2 * Math.pow(_ph_multiplier, _weight_body);
		var _ph_summ__two_obj_and_tx = Math.pow(_ph_multiplier, _weight_title) + 2 * Math.pow(_ph_multiplier, _weight_obj) + Math.pow(_ph_multiplier, _weight_body);
		var _ph_summ__two_obj_and_two_tx = Math.pow(_ph_multiplier, _weight_title) + 2 * Math.pow(_ph_multiplier, _weight_obj) + 2 * Math.pow(_ph_multiplier, _weight_body);
		var _ph_summ__tx = Math.pow(_ph_multiplier, _weight_title) + Math.pow(_ph_multiplier, _weight_body);
		var _ph_summ__tx_and_clip_art = Math.pow(_ph_multiplier, _weight_title) + Math.pow(_ph_multiplier, _weight_body) + +Math.pow(_ph_multiplier, _weight_clipArt);

		var _arr_lt_types_weight = [];
		_arr_lt_types_weight[0] = _ph_summ_blank;
		_arr_lt_types_weight[1] = _ph_summ_chart;
		_arr_lt_types_weight[2] = _ph_summ_chart_and_tx;
		_arr_lt_types_weight[3] = _ph_summ_dgm;
		_arr_lt_types_weight[4] = _ph_summ_four_obj;
		_arr_lt_types_weight[5] = _ph_summ__media_and_tx;
		_arr_lt_types_weight[6] = _ph_summ__obj;
		_arr_lt_types_weight[7] = _ph_summ__obj_and_two_obj;
		_arr_lt_types_weight[8] = _ph_summ__obj_and_tx;
		_arr_lt_types_weight[9] = _ph_summ__obj_only;
		_arr_lt_types_weight[10] = _ph_summ__pic_tx;
		_arr_lt_types_weight[11] = _ph_summ__sec_head;
		_arr_lt_types_weight[12] = _ph_summ__tbl;
		_arr_lt_types_weight[13] = _ph_summ__title_only;
		_arr_lt_types_weight[14] = _ph_summ__two_col_tx;
		_arr_lt_types_weight[15] = _ph_summ__two_obj_and_tx;
		_arr_lt_types_weight[16] = _ph_summ__two_obj_and_two_tx;
		_arr_lt_types_weight[17] = _ph_summ__tx;
		_arr_lt_types_weight[18] = _ph_summ__tx_and_clip_art;

		_arr_lt_types_weight.sort(AscCommon.fSortAscending);


		var _global_layout_summs_array = {};
		_global_layout_summs_array["_" + _ph_summ_blank] = nSldLtTBlank;
		_global_layout_summs_array["_" + _ph_summ_chart] = nSldLtTChart;
		_global_layout_summs_array["_" + _ph_summ_chart_and_tx] = nSldLtTChartAndTx;
		_global_layout_summs_array["_" + _ph_summ_dgm] = nSldLtTDgm;
		_global_layout_summs_array["_" + _ph_summ_four_obj] = nSldLtTFourObj;
		_global_layout_summs_array["_" + _ph_summ__media_and_tx] = nSldLtTMediaAndTx;
		_global_layout_summs_array["_" + _ph_summ__obj] = nSldLtTObj;
		_global_layout_summs_array["_" + _ph_summ__obj_and_two_obj] = nSldLtTObjAndTwoObj;
		_global_layout_summs_array["_" + _ph_summ__obj_and_tx] = nSldLtTObjAndTx;
		_global_layout_summs_array["_" + _ph_summ__obj_only] = nSldLtTObjOnly;
		_global_layout_summs_array["_" + _ph_summ__pic_tx] = nSldLtTPicTx;
		_global_layout_summs_array["_" + _ph_summ__sec_head] = nSldLtTSecHead;
		_global_layout_summs_array["_" + _ph_summ__tbl] = nSldLtTTbl;
		_global_layout_summs_array["_" + _ph_summ__title_only] = nSldLtTTitleOnly;
		_global_layout_summs_array["_" + _ph_summ__two_col_tx] = nSldLtTTwoColTx;
		_global_layout_summs_array["_" + _ph_summ__two_obj_and_tx] = nSldLtTTwoObjAndTx;
		_global_layout_summs_array["_" + _ph_summ__two_obj_and_two_tx] = nSldLtTTwoTxTwoObj;
		_global_layout_summs_array["_" + _ph_summ__tx] = nSldLtTTx;
		_global_layout_summs_array["_" + _ph_summ__tx_and_clip_art] = nSldLtTTxAndClipArt;


// SLIDE ----------------------------
		function redrawSlide(slide, presentation, arrInd, pos, direction, arr_slides) {
			if (slide) {
				slide.recalculate();
				presentation.DrawingDocument.OnRecalculateSlide(presentation.GetSlideIndex(slide));
			}
			if (direction === 0) {
				if (pos > 0) {
					presentation.backChangeThemeTimeOutId = setTimeout(function () {
						redrawSlide(arr_slides[arrInd[pos - 1]], presentation, arrInd, pos - 1, -1, arr_slides)
					}, recalcSlideInterval);
				} else {
					presentation.backChangeThemeTimeOutId = null;
				}
				if (pos < arrInd.length - 1) {
					presentation.forwardChangeThemeTimeOutId = setTimeout(function () {
						redrawSlide(arr_slides[arrInd[pos + 1]], presentation, arrInd, pos + 1, +1, arr_slides)
					}, recalcSlideInterval);
				} else {
					presentation.forwardChangeThemeTimeOutId = null;
				}
				presentation.startChangeThemeTimeOutId = null;
			}
			if (direction > 0) {
				if (pos < arrInd.length - 1) {
					presentation.forwardChangeThemeTimeOutId = setTimeout(function () {
						redrawSlide(arr_slides[arrInd[pos + 1]], presentation, arrInd, pos + 1, +1, arr_slides)
					}, recalcSlideInterval);
				} else {
					presentation.forwardChangeThemeTimeOutId = null;
				}
			}
			if (direction < 0) {
				if (pos > 0) {
					presentation.backChangeThemeTimeOutId = setTimeout(function () {
						redrawSlide(arr_slides[arrInd[pos - 1]], presentation, arrInd, pos - 1, -1, arr_slides)
					}, recalcSlideInterval);
				} else {
					presentation.backChangeThemeTimeOutId = null;
				}
			}
		}


		function CTextFit(nType) {
			CBaseNoIdObject.call(this);
			this.type = nType !== undefined && nType !== null ? nType : 0;
			this.fontScale = null;
			this.lnSpcReduction = null;
		}

		InitClass(CTextFit, CBaseNoIdObject, 0);
		CTextFit.prototype.CreateDublicate = function () {
			var d = new CTextFit();
			d.type = this.type;
			d.fontScale = this.fontScale;
			d.lnSpcReduction = this.lnSpcReduction;
			return d;
		};
		CTextFit.prototype.Write_ToBinary = function (w) {
			writeLong(w, this.type);
			writeLong(w, this.fontScale);
			writeLong(w, this.lnSpcReduction);
		};
		CTextFit.prototype.Read_FromBinary = function (r) {
			this.type = readLong(r);
			this.fontScale = readLong(r);
			this.lnSpcReduction = readLong(r);
		};
		CTextFit.prototype.Get_Id = function () {
			return this.Id;
		};
		CTextFit.prototype.Refresh_RecalcData = function () {
		};
//-----------------------------

//Text Anchoring Types
		var nTextATB = 0;// (Text Anchor Enum ( Bottom ))
		var nTextATCtr = 1;// (Text Anchor Enum ( Center ))
		var nTextATDist = 2;// (Text Anchor Enum ( Distributed ))
		var nTextATJust = 3;// (Text Anchor Enum ( Justified ))
		var nTextATT = 4;// Top

		function CBodyPr() {
			CBaseNoIdObject.call(this);
			this.flatTx = null;
			this.anchor = null;
			this.anchorCtr = null;
			this.bIns = null;
			this.compatLnSpc = null;
			this.forceAA = null;
			this.fromWordArt = null;
			this.horzOverflow = null;
			this.lIns = null;
			this.numCol = null;
			this.rIns = null;
			this.rot = null;
			this.rtlCol = null;
			this.spcCol = null;
			this.spcFirstLastPara = null;
			this.tIns = null;
			this.upright = null;
			this.vert = null;
			this.vertOverflow = null;
			this.wrap = null;
			this.textFit = null;
			this.prstTxWarp = null;
			this.parent = null;
		}

		InitClass(CBodyPr, CBaseNoIdObject, 0);
		CBodyPr.prototype.Get_Id = function () {
			return this.Id;
		};
		CBodyPr.prototype.Refresh_RecalcData = function () {
		};
		CBodyPr.prototype.setVertOpen = function (nVert) {
			let nVert_ = nVert;
			if(nVert === AscFormat.nVertTTwordArtVert) {
				nVert_ = AscFormat.nVertTTvert;
			}
			this.vert = nVert_;
		};
		CBodyPr.prototype.getLnSpcReduction = function () {
			if (this.textFit
				&& this.textFit.type === AscFormat.text_fit_NormAuto
				&& AscFormat.isRealNumber(this.textFit.lnSpcReduction)) {
				return this.textFit.lnSpcReduction / 100000.0;
			}
			return undefined;
		};
		CBodyPr.prototype.getFontScale = function () {
			if (this.textFit
				&& this.textFit.type === AscFormat.text_fit_NormAuto
				&& AscFormat.isRealNumber(this.textFit.fontScale)) {
				return this.textFit.fontScale / 100000.0;
			}
			return undefined;
		};
		CBodyPr.prototype.isNotNull = function () {
			return this.flatTx !== null ||
				this.anchor !== null ||
				this.anchorCtr !== null ||
				this.bIns !== null ||
				this.compatLnSpc !== null ||
				this.forceAA !== null ||
				this.fromWordArt !== null ||
				this.horzOverflow !== null ||
				this.lIns !== null ||
				this.numCol !== null ||
				this.rIns !== null ||
				this.rot !== null ||
				this.rtlCol !== null ||
				this.spcCol !== null ||
				this.spcFirstLastPara !== null ||
				this.tIns !== null ||
				this.upright !== null ||
				this.vert !== null ||
				this.vertOverflow !== null ||
				this.wrap !== null ||
				this.textFit !== null ||
				this.prstTxWarp !== null;
		};
		CBodyPr.prototype.setAnchor = function (val) {
			this.anchor = val;
		};
		CBodyPr.prototype.setVert = function (val) {
			this.vert = val;
		};
		CBodyPr.prototype.setRot = function (val) {
			this.rot = val;
		};
		CBodyPr.prototype.reset = function () {
			this.flatTx = null;
			this.anchor = null;
			this.anchorCtr = null;
			this.bIns = null;
			this.compatLnSpc = null;
			this.forceAA = null;
			this.fromWordArt = null;
			this.horzOverflow = null;
			this.lIns = null;
			this.numCol = null;
			this.rIns = null;
			this.rot = null;
			this.rtlCol = null;
			this.spcCol = null;
			this.spcFirstLastPara = null;
			this.tIns = null;
			this.upright = null;
			this.vert = null;
			this.vertOverflow = null;
			this.wrap = null;
			this.textFit = null;
			this.prstTxWarp = null;
		};
		CBodyPr.prototype.WritePrstTxWarp = function (w) {
			w.WriteBool(isRealObject(this.prstTxWarp));
			if (isRealObject(this.prstTxWarp)) {
				writeString(w, this.prstTxWarp.preset);
				var startPos = w.GetCurPosition(), countAv = 0;
				w.Skip(4);
				for (var key in this.prstTxWarp.avLst) {
					if (this.prstTxWarp.avLst.hasOwnProperty(key)) {
						++countAv;
						w.WriteString2(key);
						w.WriteLong(this.prstTxWarp.gdLst[key]);
					}
				}
				var endPos = w.GetCurPosition();
				w.Seek(startPos);
				w.WriteLong(countAv);
				w.Seek(endPos);
			}
		};
		CBodyPr.prototype.ReadPrstTxWarp = function (r) {
			ExecuteNoHistory(function () {
				if (r.GetBool()) {
					this.prstTxWarp = AscFormat.CreatePrstTxWarpGeometry(readString(r));
					var count = r.GetLong();
					for (var i = 0; i < count; ++i) {
						var sAdj = r.GetString2();
						var nVal = r.GetLong();
						this.prstTxWarp.AddAdj(sAdj, 15, nVal + "", undefined, undefined);
					}
				}
			}, this, []);
		};
		CBodyPr.prototype.Write_ToBinary2 = function (w) {
			var flag = this.flatTx != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.flatTx);
			}

			flag = this.anchor != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.anchor);
			}

			flag = this.anchorCtr != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.anchorCtr);
			}

			flag = this.bIns != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteDouble(this.bIns);
			}

			flag = this.compatLnSpc != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.compatLnSpc);
			}

			flag = this.forceAA != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.forceAA);
			}

			flag = this.fromWordArt != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.fromWordArt);
			}

			flag = this.horzOverflow != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.horzOverflow);
			}

			flag = this.lIns != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteDouble(this.lIns);
			}

			flag = this.numCol != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.numCol);
			}

			flag = this.rIns != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteDouble(this.rIns);
			}


			flag = this.rot != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.rot);
			}

			flag = this.rtlCol != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.rtlCol);
			}

			flag = this.spcCol != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteDouble(this.spcCol);
			}

			flag = this.spcFirstLastPara != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.spcFirstLastPara);
			}

			flag = this.tIns != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteDouble(this.tIns);
			}

			flag = this.upright != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.upright);
			}

			flag = this.vert != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.vert);
			}


			flag = this.vertOverflow != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.vertOverflow);
			}

			flag = this.wrap != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.wrap);
			}

			this.WritePrstTxWarp(w);

			w.WriteBool(isRealObject(this.textFit));
			if (this.textFit) {
				this.textFit.Write_ToBinary(w);
			}
		};
		CBodyPr.prototype.Read_FromBinary2 = function (r) {
			var flag = r.GetBool();
			if (flag) {
				this.flatTx = r.GetLong();
			}

			flag = r.GetBool();
			if (flag) {
				this.anchor = r.GetLong();
			}


			flag = r.GetBool();
			if (flag) {
				this.anchorCtr = r.GetBool();
			}


			flag = r.GetBool();
			if (flag) {
				this.bIns = r.GetDouble();
			}


			flag = r.GetBool();
			if (flag) {
				this.compatLnSpc = r.GetBool();
			}


			flag = r.GetBool();
			if (flag) {
				this.forceAA = r.GetBool();
			}


			flag = r.GetBool();
			if (flag) {
				this.fromWordArt = r.GetBool();
			}


			flag = r.GetBool();
			if (flag) {
				this.horzOverflow = r.GetLong();
			}


			flag = r.GetBool();
			if (flag) {
				this.lIns = r.GetDouble();
			}


			flag = r.GetBool();
			if (flag) {
				this.numCol = r.GetLong();
			}


			flag = r.GetBool();
			if (flag) {
				this.rIns = r.GetDouble();
			}


			flag = r.GetBool();
			if (flag) {
				this.rot = r.GetLong();
			}


			flag = r.GetBool();
			if (flag) {
				this.rtlCol = r.GetBool();
			}

			flag = r.GetBool();
			if (flag) {
				this.spcCol = r.GetDouble();
			}

			flag = r.GetBool();
			if (flag) {
				this.spcFirstLastPara = r.GetBool();
			}


			flag = r.GetBool();
			if (flag) {
				this.tIns = r.GetDouble();
			}


			flag = r.GetBool();
			if (flag) {
				this.upright = r.GetBool();
			}


			flag = r.GetBool();
			if (flag) {
				this.vert = r.GetLong();
			}


			flag = r.GetBool();
			if (flag) {
				this.vertOverflow = r.GetLong();
			}


			flag = r.GetBool();
			if (flag) {
				this.wrap = r.GetLong();
			}
			this.ReadPrstTxWarp(r);

			if (r.GetBool()) {
				this.textFit = new CTextFit();
				this.textFit.Read_FromBinary(r);
			}
		};
		CBodyPr.prototype.setDefaultInsets = function() {
			this.bIns = 45720 / 36000;
			this.tIns = 45720 / 36000;
			this.lIns = 91440 / 36000;
			this.rIns = 91440 / 36000;
		};
		CBodyPr.prototype.setInsets = function(l, t, r, b) {
			this.lIns = l;
			this.tIns = t;
			this.rIns = r;
			this.bIns = b;
		};
		CBodyPr.prototype.resetInsets = function() {
			this.setInsets(0, 0, 0, 0);
		};
		CBodyPr.prototype.setDefault = function () {
			this.setDefaultInsets();
			this.flatTx = null;
			this.anchor = 4;
			this.anchorCtr = false;
			this.compatLnSpc = false;
			this.forceAA = false;
			this.fromWordArt = false;
			this.horzOverflow = AscFormat.nHOTOverflow;
			this.numCol = 1;
			this.rot = null;
			this.rtlCol = false;
			this.spcCol = false;
			this.spcFirstLastPara = null;
			this.upright = false;
			this.vert = AscFormat.nVertTThorz;
			this.vertOverflow = AscFormat.nVOTOverflow;
			this.wrap = AscFormat.nTWTSquare;
			this.prstTxWarp = null;
			this.textFit = null;
		};
		CBodyPr.prototype.createDuplicate = function () {
			var duplicate = new CBodyPr();
			duplicate.flatTx = this.flatTx;
			duplicate.anchor = this.anchor;
			duplicate.anchorCtr = this.anchorCtr;
			duplicate.bIns = this.bIns;
			duplicate.compatLnSpc = this.compatLnSpc;
			duplicate.forceAA = this.forceAA;
			duplicate.fromWordArt = this.fromWordArt;
			duplicate.horzOverflow = this.horzOverflow;
			duplicate.lIns = this.lIns;
			duplicate.numCol = this.numCol;
			duplicate.rIns = this.rIns;
			duplicate.rot = this.rot;
			duplicate.rtlCol = this.rtlCol;
			duplicate.spcCol = this.spcCol;
			duplicate.spcFirstLastPara = this.spcFirstLastPara;
			duplicate.tIns = this.tIns;
			duplicate.upright = this.upright;
			duplicate.vert = this.vert;
			duplicate.vertOverflow = this.vertOverflow;
			duplicate.wrap = this.wrap;
			if (this.prstTxWarp) {
				duplicate.prstTxWarp = ExecuteNoHistory(function () {
					return this.prstTxWarp.createDuplicate();
				}, this, []);
			}
			if (this.textFit) {
				duplicate.textFit = this.textFit.CreateDublicate();
			}
			return duplicate;
		};
		CBodyPr.prototype.createDuplicateForSmartArt = function (oPr) {
			var duplicate = new CBodyPr();
			duplicate.anchor = this.anchor;
			duplicate.vert = this.vert;
			duplicate.rot = this.rot;
			duplicate.vertOverflow = this.vertOverflow;
			duplicate.horzOverflow = this.horzOverflow;
			duplicate.upright = this.upright;
			duplicate.rtlCol = this.rtlCol;
			duplicate.fromWordArt = this.fromWordArt;
			duplicate.compatLnSpc = this.compatLnSpc;
			duplicate.forceAA = this.forceAA;
			if (oPr.lIns) {
				duplicate.lIns = this.lIns;
			}
			if (oPr.rIns) {
				duplicate.rIns = this.rIns;
			}
			if (oPr.tIns) {
				duplicate.tIns = this.tIns;
			}
			if (oPr.bIns) {
				duplicate.bIns = this.bIns;
			}
			return duplicate;
		};
		CBodyPr.prototype.merge = function (bodyPr) {
			if (!bodyPr)
				return;
			if (bodyPr.flatTx != null) {
				this.flatTx = bodyPr.flatTx;
			}
			if (bodyPr.anchor != null) {
				this.anchor = bodyPr.anchor;
			}
			if (bodyPr.anchorCtr != null) {
				this.anchorCtr = bodyPr.anchorCtr;
			}
			if (bodyPr.bIns != null) {
				this.bIns = bodyPr.bIns;
			}
			if (bodyPr.compatLnSpc != null) {
				this.compatLnSpc = bodyPr.compatLnSpc;
			}
			if (bodyPr.forceAA != null) {
				this.forceAA = bodyPr.forceAA;
			}
			if (bodyPr.fromWordArt != null) {
				this.fromWordArt = bodyPr.fromWordArt;
			}

			if (bodyPr.horzOverflow != null) {
				this.horzOverflow = bodyPr.horzOverflow;
			}

			if (bodyPr.lIns != null) {
				this.lIns = bodyPr.lIns;
			}

			if (bodyPr.numCol != null) {
				this.numCol = bodyPr.numCol;
			}
			if (bodyPr.rIns != null) {
				this.rIns = bodyPr.rIns;
			}

			if (bodyPr.rot != null) {
				this.rot = bodyPr.rot;
			}

			if (bodyPr.rtlCol != null) {
				this.rtlCol = bodyPr.rtlCol;
			}

			if (bodyPr.spcCol != null) {
				this.spcCol = bodyPr.spcCol;
			}
			if (bodyPr.spcFirstLastPara != null) {
				this.spcFirstLastPara = bodyPr.spcFirstLastPara;
			}

			if (bodyPr.tIns != null) {
				this.tIns = bodyPr.tIns;
			}
			if (bodyPr.upright != null) {
				this.upright = bodyPr.upright;
			}

			if (bodyPr.vert != null) {
				this.vert = bodyPr.vert;
			}
			if (bodyPr.vertOverflow != null) {
				this.vertOverflow = bodyPr.vertOverflow;
			}
			if (bodyPr.wrap != null) {
				this.wrap = bodyPr.wrap;
			}
			if (bodyPr.prstTxWarp) {
				this.prstTxWarp = ExecuteNoHistory(function () {
					return bodyPr.prstTxWarp.createDuplicate();
				}, this, []);
			}
			if (bodyPr.textFit) {
				this.textFit = bodyPr.textFit.CreateDublicate();
			}
			if (bodyPr.numCol != null) {
				this.numCol = bodyPr.numCol;
			}
			return this;
		};
		CBodyPr.prototype.Write_ToBinary = function (w) {
			var flag = this.flatTx != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.flatTx);
			}

			flag = this.anchor != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.anchor);
			}

			flag = this.anchorCtr != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.anchorCtr);
			}

			flag = this.bIns != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteDouble(this.bIns);
			}

			flag = this.compatLnSpc != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.compatLnSpc);
			}

			flag = this.forceAA != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.forceAA);
			}

			flag = this.fromWordArt != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.fromWordArt);
			}

			flag = this.horzOverflow != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.horzOverflow);
			}

			flag = this.lIns != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteDouble(this.lIns);
			}

			flag = this.numCol != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.numCol);
			}

			flag = this.rIns != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteDouble(this.rIns);
			}


			flag = this.rot != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.rot);
			}

			flag = this.rtlCol != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.rtlCol);
			}

			flag = this.spcCol != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteDouble(this.spcCol);
			}

			flag = this.spcFirstLastPara != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.spcFirstLastPara);
			}

			flag = this.tIns != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteDouble(this.tIns);
			}

			flag = this.upright != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteBool(this.upright);
			}

			flag = this.vert != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.vert);
			}


			flag = this.vertOverflow != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.vertOverflow);
			}

			flag = this.wrap != null;
			w.WriteBool(flag);
			if (flag) {
				w.WriteLong(this.wrap);
			}
			this.WritePrstTxWarp(w);
			w.WriteBool(isRealObject(this.textFit));
			if (this.textFit) {
				this.textFit.Write_ToBinary(w);
			}
		};
		CBodyPr.prototype.Read_FromBinary = function (r) {
			var flag = r.GetBool();
			if (flag) {
				this.flatTx = r.GetLong();
			}

			flag = r.GetBool();
			if (flag) {
				this.anchor = r.GetLong();
			}


			flag = r.GetBool();
			if (flag) {
				this.anchorCtr = r.GetBool();
			}


			flag = r.GetBool();
			if (flag) {
				this.bIns = r.GetDouble();
			}


			flag = r.GetBool();
			if (flag) {
				this.compatLnSpc = r.GetBool();
			}


			flag = r.GetBool();
			if (flag) {
				this.forceAA = r.GetBool();
			}


			flag = r.GetBool();
			if (flag) {
				this.fromWordArt = r.GetBool();
			}


			flag = r.GetBool();
			if (flag) {
				this.horzOverflow = r.GetLong();
			}


			flag = r.GetBool();
			if (flag) {
				this.lIns = r.GetDouble();
			}


			flag = r.GetBool();
			if (flag) {
				this.numCol = r.GetLong();
			}


			flag = r.GetBool();
			if (flag) {
				this.rIns = r.GetDouble();
			}


			flag = r.GetBool();
			if (flag) {
				this.rot = r.GetLong();
			}


			flag = r.GetBool();
			if (flag) {
				this.rtlCol = r.GetBool();
			}

			flag = r.GetBool();
			if (flag) {
				this.spcCol = r.GetDouble();
			}

			flag = r.GetBool();
			if (flag) {
				this.spcFirstLastPara = r.GetBool();
			}


			flag = r.GetBool();
			if (flag) {
				this.tIns = r.GetDouble();
			}


			flag = r.GetBool();
			if (flag) {
				this.upright = r.GetBool();
			}


			flag = r.GetBool();
			if (flag) {
				this.vert = r.GetLong();
			}


			flag = r.GetBool();
			if (flag) {
				this.vertOverflow = r.GetLong();
			}


			flag = r.GetBool();
			if (flag) {
				this.wrap = r.GetLong();
			}
			this.ReadPrstTxWarp(r);

			if (r.GetBool()) {
				this.textFit = new CTextFit();
				this.textFit.Read_FromBinary(r)
			}
		};


		function CTextParagraphPr() {
			this.bullet = new CBullet();
			this.lvl = null;
			this.pPr = new CParaPr();
			this.rPr = new CTextPr();
		}

		function CreateNoneBullet() {
			var ret = new CBullet();
			ret.bulletType = new CBulletType();
			ret.bulletType.type = AscFormat.BULLET_TYPE_BULLET_NONE;
			return ret;
		}

		function CompareBullets(bullet1, bullet2) {

			//TODO: пока будем сравнивать только bulletType, т. к. эта функция используется для мержа свойств при отдаче в интерфейс, а для интерфейса bulletTyp'a достаточно. Если понадобится нужно сделать полное сравнение.
			//
			if (bullet1.bulletType && bullet2.bulletType
				&& bullet1.bulletType.type === bullet2.bulletType.type) {
				var ret = new CBullet();
				ret.bulletType = new CBulletType();
				switch (bullet1.bulletType.type) {
					case AscFormat.BULLET_TYPE_BULLET_CHAR: {
						ret.bulletType.type = AscFormat.BULLET_TYPE_BULLET_CHAR;
						if (bullet1.bulletType.Char === bullet2.bulletType.Char) {
							ret.bulletType.Char = bullet1.bulletType.Char;
						}
						break;
					}
					case AscFormat.BULLET_TYPE_BULLET_BLIP: {
						ret.bulletType.type = AscFormat.BULLET_TYPE_BULLET_BLIP;
						var compareBlip = bullet1.bulletType.Blip && bullet2.bulletType.Blip && bullet1.bulletType.Blip.compare(bullet2.bulletType.Blip);
						ret.bulletType.Blip = compareBlip;
						break;
					}
					case AscFormat.BULLET_TYPE_BULLET_AUTONUM: {
						if (bullet1.bulletType.AutoNumType === bullet2.bulletType.AutoNumType) {
							ret.bulletType.AutoNumType = bullet1.bulletType.AutoNumType;
						}
						if (bullet1.bulletType.startAt === bullet2.bulletType.startAt) {
							ret.bulletType.startAt = bullet1.bulletType.startAt;
						} else {
							ret.bulletType.startAt = undefined;
						}
						if (bullet1.bulletType.type === bullet2.bulletType.type) {
							ret.bulletType.type = bullet1.bulletType.type;
						}
						break;
					}
				}

				if (bullet1.bulletSize && bullet2.bulletSize
					&& bullet1.bulletSize.val === bullet2.bulletSize.val
					&& bullet1.bulletSize.type === bullet2.bulletSize.type) {
					ret.bulletSize = bullet1.bulletSize;
				}
				if (bullet1.bulletColor && bullet2.bulletColor
					&& bullet1.bulletColor.type === bullet2.bulletColor.type) {
					ret.bulletColor = new CBulletColor();
					ret.bulletColor.type = bullet2.bulletColor.type;
					if (bullet1.bulletColor.UniColor) {
						ret.bulletColor.UniColor = bullet1.bulletColor.UniColor.compare(bullet2.bulletColor.UniColor);
					}
					if (!ret.bulletColor.UniColor || !ret.bulletColor.UniColor.color) {
						ret.bulletColor = null;
					}
				}
				return ret;
			} else {
				return undefined;
			}
		}

		function CBullet() {
			CBaseNoIdObject.call(this)
			this.bulletColor = null;
			this.bulletSize = null;
			this.bulletTypeface = null;
			this.bulletType = null;
			this.Bullet = null;

			//used to get properties for interface
			this.FirstTextPr = null;
		}

		InitClass(CBullet, CBaseNoIdObject, 0);
		CBullet.prototype.Set_FromObject = function (obj) {
			if (obj) {
				if (obj.bulletColor) {
					this.bulletColor = new CBulletColor();
					this.bulletColor.Set_FromObject(obj.bulletColor);
				} else
					this.bulletColor = null;

				if (obj.bulletSize) {
					this.bulletSize = new CBulletSize();
					this.bulletSize.Set_FromObject(obj.bulletSize);
				} else
					this.bulletSize = null;

				if (obj.bulletTypeface) {
					this.bulletTypeface = new CBulletTypeface();
					this.bulletTypeface.Set_FromObject(obj.bulletTypeface);
				} else
					this.bulletTypeface = null;
			}
		};
		CBullet.prototype.merge = function (oBullet) {
			if (!oBullet) {
				return;
			}
			if (oBullet.bulletColor) {
				if (!this.bulletColor) {
					this.bulletColor = oBullet.bulletColor.createDuplicate();
				} else {
					this.bulletColor.merge(oBullet.bulletColor);
				}
			}
			if (oBullet.bulletSize) {
				if (!this.bulletSize) {
					this.bulletSize = oBullet.bulletSize.createDuplicate();
				} else {
					this.bulletSize.merge(oBullet.bulletSize);
				}
			}
			if (oBullet.bulletTypeface) {
				if (!this.bulletTypeface) {
					this.bulletTypeface = oBullet.bulletTypeface.createDuplicate();
				} else {
					this.bulletTypeface.merge(oBullet.bulletTypeface);
				}
			}
			if (oBullet.bulletType) {
				if (!this.bulletType) {
					this.bulletType = oBullet.bulletType.createDuplicate();
				} else {
					this.bulletType.merge(oBullet.bulletType);
				}
			}
		};
		CBullet.prototype.createDuplicate = function () {
			var duplicate = new CBullet();
			if (this.bulletColor) {
				duplicate.bulletColor = this.bulletColor.createDuplicate();
			}
			if (this.bulletSize) {
				duplicate.bulletSize = this.bulletSize.createDuplicate();
			}
			if (this.bulletTypeface) {
				duplicate.bulletTypeface = this.bulletTypeface.createDuplicate();
			}

			if (this.bulletType) {
				duplicate.bulletType = this.bulletType.createDuplicate();
			}

			duplicate.Bullet = this.Bullet;
			return duplicate;
		};
		CBullet.prototype.isBullet = function () {
			return this.bulletType != null && this.bulletType.type != null;
		};
		CBullet.prototype.isNone = function() {
			if (!this.bulletType)
				return true;
			
			return this.bulletType.type === AscFormat.BULLET_TYPE_TYPEFACE_NONE;
		};
		CBullet.prototype.getPresentationBullet = function (theme, color) {
			var para_pr = new CParaPr();
			para_pr.Bullet = this;
			return para_pr.Get_PresentationBullet(theme, color);
		};
		CBullet.prototype.getBulletType = function (theme, color) {
			return this.getPresentationBullet(theme, color).m_nType;
		};
		CBullet.prototype.Write_ToBinary = function (w) {
			w.WriteBool(isRealObject(this.bulletColor));
			if (isRealObject(this.bulletColor)) {
				this.bulletColor.Write_ToBinary(w);
			}

			w.WriteBool(isRealObject(this.bulletSize));
			if (isRealObject(this.bulletSize)) {
				this.bulletSize.Write_ToBinary(w);
			}


			w.WriteBool(isRealObject(this.bulletTypeface));
			if (isRealObject(this.bulletTypeface)) {
				this.bulletTypeface.Write_ToBinary(w);
			}


			w.WriteBool(isRealObject(this.bulletType));
			if (isRealObject(this.bulletType)) {
				this.bulletType.Write_ToBinary(w);
			}

		};
		CBullet.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				this.bulletColor = new CBulletColor();
				this.bulletColor.Read_FromBinary(r);
			}

			if (r.GetBool()) {
				this.bulletSize = new CBulletSize();
				this.bulletSize.Read_FromBinary(r);
			}

			if (r.GetBool()) {
				this.bulletTypeface = new CBulletTypeface();
				this.bulletTypeface.Read_FromBinary(r);
			}

			if (r.GetBool()) {
				this.bulletType = new CBulletType();
				this.bulletType.Read_FromBinary(r);
			}
		};
		CBullet.prototype.Get_AllFontNames = function (AllFonts) {
			if (this.bulletTypeface && typeof this.bulletTypeface.typeface === "string" && this.bulletTypeface.typeface.length > 0) {
				AllFonts[this.bulletTypeface.typeface] = true;
			}
		};
		CBullet.prototype.putNumStartAt = function (NumStartAt) {
			if (!this.bulletType) {
				this.bulletType = new CBulletType();
			}
			this.bulletType.type = AscFormat.BULLET_TYPE_BULLET_AUTONUM;
			this.bulletType.startAt = NumStartAt;
		};
		CBullet.prototype.getNumStartAt = function () {
			if (this.bulletType) {
				if (AscFormat.isRealNumber(this.bulletType.startAt)) {
					return Math.max(1, this.bulletType.startAt);
				}
			}
			return undefined;
		};
		CBullet.prototype.isEqual = function (oBullet) {
			if (!oBullet) {
				return false;
			}
			if (!this.bulletColor && oBullet.bulletColor
				|| !oBullet.bulletColor && this.bulletColor) {
				return false;
			}
			if (this.bulletColor && oBullet.bulletColor) {
				if (!this.bulletColor.IsIdentical(oBullet.bulletColor)) {
					return false;
				}
			}
			if (!this.bulletSize && oBullet.bulletSize
				|| this.bulletSize && !oBullet.bulletSize) {
				return false;
			}
			if (this.bulletSize && oBullet.bulletSize) {
				if (!this.bulletSize.IsIdentical(oBullet.bulletSize)) {
					return false;
				}
			}
			if (!this.bulletTypeface && oBullet.bulletTypeface
				|| this.bulletTypeface && !oBullet.bulletTypeface) {
				return false;
			}
			if (this.bulletTypeface && oBullet.bulletTypeface) {
				if (!this.bulletTypeface.IsIdentical(oBullet.bulletTypeface)) {
					return false;
				}
			}
			if (!this.bulletType && oBullet.bulletType
				|| this.bulletType && !oBullet.bulletType) {
				return false;
			}
			if (this.bulletType && oBullet.bulletType) {
				if (!this.bulletType.IsIdentical(oBullet.bulletType)) {
					return false;
				}
			}
			return true;
		};
		CBullet.prototype.fillBulletImage = function (url) {
			this.bulletType = new CBulletType();
			this.bulletType.Blip = new AscFormat.CBuBlip();
			this.bulletType.type = AscFormat.BULLET_TYPE_BULLET_BLIP;
			this.bulletType.Blip.setBlip(AscFormat.CreateBlipFillUniFillFromUrl(url));

		};
		CBullet.prototype.fillBulletFromCharAndFont = function (char, font) {
			this.bulletType = new AscFormat.CBulletType();
			this.bulletTypeface = new AscFormat.CBulletTypeface();
			this.bulletTypeface.type = AscFormat.BULLET_TYPE_TYPEFACE_BUFONT;
			this.bulletTypeface.typeface = font || AscFonts.FontPickerByCharacter.getFontBySymbol(char.getUnicodeIterator().value());
			this.bulletType.type = AscFormat.BULLET_TYPE_BULLET_CHAR;
			this.bulletType.Char = char;
		};
		CBullet.prototype.getImageBulletURL = function () {
			var res = (this.bulletType
				&& this.bulletType.Blip
				&& this.bulletType.Blip.blip
				&& this.bulletType.Blip.blip.fill
				&& this.bulletType.Blip.blip.fill.RasterImageId);
			return res ? res : null;
		};
		CBullet.prototype.setImageBulletURL = function (url) {
			var blipFill = (this.bulletType
				&& this.bulletType.Blip
				&& this.bulletType.Blip.blip
				&& this.bulletType.Blip.blip.fill);
			if (blipFill) {
				blipFill.setRasterImageId(url);
			}
		};
		CBullet.prototype.drawSquareImage = function (sDivId, nRelativeIndent) {
			nRelativeIndent = nRelativeIndent || 0;
			const sImageUrl = this.getImageBulletURL();
			const oApi = editor || Asc.editor;
			if (!sImageUrl || !oApi) {
				return;
			}

			const oDiv = document.getElementById(sDivId);
			if (!oDiv) {
				return;
			}
			const nWidth = oDiv.clientWidth;
			const nHeight = oDiv.clientHeight;
			const nRPR = AscCommon.AscBrowser.retinaPixelRatio;
			const nCanvasSide = Math.min(nWidth, nHeight) * nRPR;

			let oCanvas = oDiv.firstChild;
			if (!oCanvas) {
				oCanvas = document.createElement('canvas');
				oCanvas.style.cssText = "padding:0;margin:0;user-select:none;";
				oCanvas.style.width = oDiv.clientWidth + 'px';
				oCanvas.style.height = oDiv.clientHeight + 'px';
				oCanvas.width = nCanvasSide;
				oCanvas.height = nCanvasSide;
				oDiv.appendChild(oCanvas);
			}

			const oContext = oCanvas.getContext('2d');
			oContext.fillStyle = "white";
			oContext.fillRect(0, 0, oCanvas.width, oCanvas.height);
			const oImage = oApi.ImageLoader.map_image_index[AscCommon.getFullImageSrc2(sImageUrl)];
			if (oImage && oImage.Image && oImage.Status !== AscFonts.ImageLoadStatus.Loading) {
				const nImageWidth = oImage.Image.width;
				const nImageHeight = oImage.Image.height;
				const absoluteIndent = nCanvasSide * nRelativeIndent;
				const nSideSizeWithoutIndent = nCanvasSide - 2 * absoluteIndent;
				const nAdaptCoefficient = Math.max(nImageWidth / nSideSizeWithoutIndent, nImageHeight / nSideSizeWithoutIndent);
				const nImageAdaptWidth = nImageWidth / nAdaptCoefficient;
				const nImageAdaptHeight = nImageHeight / nAdaptCoefficient;
				const nX = (nCanvasSide - nImageAdaptWidth) / 2;
				const nY = (nCanvasSide - nImageAdaptHeight) / 2;
				oContext.drawImage(oImage.Image, nX, nY, nImageAdaptWidth, nImageAdaptHeight);
			}
		};
		//interface methods
		var prot = CBullet.prototype;
		prot["fillBulletImage"] = prot["asc_fillBulletImage"] = CBullet.prototype.fillBulletImage;
		prot["fillBulletFromCharAndFont"] = prot["asc_fillBulletFromCharAndFont"] = CBullet.prototype.fillBulletFromCharAndFont;
		prot["drawSquareImage"] = prot["asc_drawSquareImage"] = CBullet.prototype.drawSquareImage;
		prot.getImageId = function () {
			return this.getImageBulletURL();
		}
		prot["getImageId"] = prot["asc_getImageId"] = CBullet.prototype.getImageId;
		prot.getJsonBullet = prot["asc_getJsonBullet"] = function () {
			const sUrlId = this.getImageBulletURL();
			const oRes = window['AscJsonConverter'].WriterToJSON.prototype.SerBullet(this);
			if (sUrlId) {
				const oBuBlip = oRes["bulletType"] && oRes["bulletType"]["buBlip"] && oRes["bulletType"]["buBlip"]["blip"] && oRes["bulletType"]["buBlip"]["blip"]["fill"];
				if (oBuBlip) {
					oBuBlip["rasterImageId"] = sUrlId;
				}
			}
			return oRes;
		}
		prot.put_ImageUrl = function (sUrl, token) {
			var _this = this;
			var Api = editor || Asc.editor;
			if (!Api) {
				return;
			}
			AscCommon.sendImgUrls(Api, [sUrl], function (data) {
				if (data && data[0] && data[0].url !== "error") {
					var url = AscCommon.g_oDocumentUrls.imagePath2Local(data[0].path);
					Api.ImageLoader.LoadImagesWithCallback([AscCommon.getFullImageSrc2(url)], function () {
						_this.fillBulletImage(url);
						//_this.drawSquareImage();
						Api.sendEvent("asc_onBulletImageLoaded", _this);
					});
				}
			}, false, token);
		};
		prot["put_ImageUrl"] = prot["asc_putImageUrl"] = CBullet.prototype.put_ImageUrl;
		prot.showFileDialog = function () {

			var Api = editor || Asc.editor;

			if (!Api) {
				return;
			}
			var _this = this;
			AscCommon.ShowImageFileDialog(Api.documentId, Api.documentUserId, Api.CoAuthoringApi.get_jwt(), Api.documentShardKey, Api.documentWopiSrc, Api.documentUserSessionId, function (error, files) {
					if (Asc.c_oAscError.ID.No !== error) {
						Api.sendEvent("asc_onError", error, Asc.c_oAscError.Level.NoCritical);
					} else {
						Api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
						AscCommon.UploadImageFiles(files, Api.documentId, Api.documentUserId, Api.CoAuthoringApi.get_jwt(), Api.documentShardKey, Api.documentWopiSrc, Api.documentUserSessionId, function (error, urls) {
							if (Asc.c_oAscError.ID.No !== error) {
								Api.sendEvent("asc_onError", error, Asc.c_oAscError.Level.NoCritical);
								Api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
							} else {
								Api.ImageLoader.LoadImagesWithCallback(urls, function () {
									if (urls.length > 0) {
										_this.fillBulletImage(urls[0]);
										//_this.drawSquareImage();
										Api.sendEvent("asc_onBulletImageLoaded", _this);
									}
									Api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
								});
							}
						});
					}
				},
				function (error) {
					if (Asc.c_oAscError.ID.No !== error) {
						Api.sendEvent("asc_onError", error, Asc.c_oAscError.Level.NoCritical);
					}
					Api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
				});
		};
		prot["showFileDialog"] = prot["asc_showFileDialog"] = CBullet.prototype.showFileDialog;
		prot.asc_getSize = function () {
			var nRet = 100;
			if (this.bulletSize) {
				switch (this.bulletSize.type) {
					case AscFormat.BULLET_TYPE_SIZE_NONE: {
						break;
					}
					case AscFormat.BULLET_TYPE_SIZE_TX: {
						break;
					}
					case AscFormat.BULLET_TYPE_SIZE_PCT: {
						nRet = this.bulletSize.val / 1000.0;
						break;
					}
					case AscFormat.BULLET_TYPE_SIZE_PTS: {
						break;
					}
				}
			}
			return nRet;
		};
		prot["get_Size"] = prot["asc_getSize"] = CBullet.prototype.asc_getSize;
		prot.asc_putSize = function (Size) {
			if (AscFormat.isRealNumber(Size)) {
				this.bulletSize = new AscFormat.CBulletSize();
				this.bulletSize.type = AscFormat.BULLET_TYPE_SIZE_PCT;
				this.bulletSize.val = (Size * 1000) >> 0;
			}
		};
		prot["put_Size"] = prot["asc_putSize"] = CBullet.prototype.asc_putSize;
		prot.asc_getColor = function () {
			if (this.bulletColor) {
				if (this.bulletColor.UniColor) {
					return AscCommon.CreateAscColor(this.bulletColor.UniColor);
				}
			} else {
				var FirstTextPr = this.FirstTextPr;
				if (FirstTextPr && FirstTextPr.Unifill) {
					if (FirstTextPr.Unifill.fill instanceof AscFormat.CSolidFill && FirstTextPr.Unifill.fill.color) {
						return AscCommon.CreateAscColor(FirstTextPr.Unifill.fill.color);
					} else {
						var RGBA = FirstTextPr.Unifill.getRGBAColor();
						return AscCommon.CreateAscColorCustom(RGBA.R, RGBA.G, RGBA.B);
					}
				}
			}
			return AscCommon.CreateAscColorCustom(0, 0, 0);
		};
		prot["get_Color"] = prot["asc_getColor"] = prot.asc_getColor;
		prot.asc_putColor = function (color) {
			this.bulletColor = new AscFormat.CBulletColor();
			this.bulletColor.type = AscFormat.BULLET_TYPE_COLOR_CLR;
			this.bulletColor.UniColor = AscFormat.CorrectUniColor(color, this.bulletColor.UniColor, 0);
		};
		prot["put_Color"] = prot["asc_putColor"] = prot.asc_putColor;
		prot.asc_getFont = function () {
			var sRet = "";
			if (this.bulletTypeface
				&& this.bulletTypeface.type === AscFormat.BULLET_TYPE_TYPEFACE_BUFONT
				&& typeof this.bulletTypeface.typeface === "string"
				&& this.bulletTypeface.typeface.length > 0) {
				sRet = this.bulletTypeface.typeface;
			} else {
				var FirstTextPr = this.FirstTextPr;
				if (FirstTextPr && FirstTextPr.FontFamily && typeof FirstTextPr.FontFamily.Name === "string"
					&& FirstTextPr.FontFamily.Name.length > 0) {
					sRet = FirstTextPr.FontFamily.Name;
				}
			}
			return sRet;
		};
		prot["get_Font"] = prot["asc_getFont"] = prot.asc_getFont;
		prot.asc_putFont = function (val) {
			if (typeof val === "string" && val.length > 0) {
				this.bulletTypeface = new AscFormat.CBulletTypeface();
				this.bulletTypeface.type = AscFormat.BULLET_TYPE_TYPEFACE_BUFONT;
				this.bulletTypeface.typeface = val;
			}
		};
		prot["put_Font"] = prot["asc_putFont"] = prot.asc_putFont;
		prot.asc_putNumStartAt = function (NumStartAt) {
			this.putNumStartAt(NumStartAt);
		};
		prot["put_NumStartAt"] = prot["asc_putNumStartAt"] = prot.asc_putNumStartAt;
		prot.asc_getNumStartAt = function () {
			return this.getNumStartAt();
		};
		prot["get_NumStartAt"] = prot["asc_getNumStartAt"] = prot.asc_getNumStartAt;
		prot.asc_getSymbol = function () {
			if (this.bulletType && this.bulletType.type === AscFormat.BULLET_TYPE_BULLET_CHAR) {
				return this.bulletType.Char;
			}
			return undefined;
		};
		prot["get_Symbol"] = prot["asc_getSymbol"] = prot.asc_getSymbol;
		prot.asc_putSymbol = function (v) {
			if (!this.bulletType) {
				this.bulletType = new CBulletType();
			}
			this.bulletType.AutoNumType = 0;
			this.bulletType.type = AscFormat.BULLET_TYPE_BULLET_CHAR;
			this.bulletType.Char = v;
		};
		prot["put_Symbol"] = prot["asc_putSymbol"] = prot.asc_putSymbol;
		prot.asc_putAutoNumType = function (val) {
			if (!this.bulletType) {
				this.bulletType = new CBulletType();
			}
			this.bulletType.type = AscFormat.BULLET_TYPE_BULLET_AUTONUM;
			this.bulletType.AutoNumType = AscFormat.getNumberingType(val);
		};
		prot["put_AutoNumType"] = prot["asc_putAutoNumType"] = prot.asc_putAutoNumType;
		prot.asc_getAutoNumType = function () {
			if (this.bulletType && this.bulletType.type === AscFormat.BULLET_TYPE_BULLET_AUTONUM) {
				return AscFormat.fGetListTypeFromBullet(this).SubType;
			}
			return -1;
		};
		prot["get_AutoNumType"] = prot["asc_getAutoNumType"] = prot.asc_getAutoNumType;
		prot.asc_putListType = function (type, subtype, custom) {
			var NumberInfo =
				{
					Type: type,
					SubType: subtype,
					Custom: custom
				};
			AscFormat.fFillBullet(NumberInfo, this);
		};
		prot["put_ListType"] = prot["asc_putListType"] = prot.asc_putListType;
		prot.asc_getListType = function () {
			return new AscCommon.asc_CListType(AscFormat.fGetListTypeFromBullet(this));
		};
		prot.asc_getType = function () {
			return this.bulletType && this.bulletType.type;
		};
		prot["get_Type"] = prot["asc_getType"] = prot.asc_getType;
		window["Asc"]["asc_CBullet"] = window["Asc"].asc_CBullet = CBullet;


		function CBulletColor(nType) {
			CBaseNoIdObject.call(this);
			this.type = AscFormat.isRealNumber(nType) ? nType : AscFormat.BULLET_TYPE_COLOR_CLRTX;
			this.UniColor = null;
		}

		InitClass(CBulletColor, CBaseNoIdObject, 0);
		CBulletColor.prototype.Set_FromObject = function (o) {
			this.merge(o);
		};
		CBulletColor.prototype.merge = function (oBulletColor) {
			if (!oBulletColor) {
				return;
			}
			if (oBulletColor.UniColor) {
				this.type = oBulletColor.type;
				this.UniColor = oBulletColor.UniColor.createDuplicate();
			}
		};
		CBulletColor.prototype.IsIdentical = function (oBulletColor) {
			if (!oBulletColor) {
				return false;
			}
			if (this.type !== oBulletColor.type) {
				return false;
			}
			if (this.UniColor && !oBulletColor.UniColor || oBulletColor.UniColor && !this.UniColor) {
				return false;
			}
			if (this.UniColor) {
				if (!this.UniColor.IsIdentical(oBulletColor.UniColor)) {
					return false;
				}
			}
			return true;

		};
		CBulletColor.prototype.createDuplicate = function () {
			var duplicate = new CBulletColor();
			duplicate.type = this.type;
			if (this.UniColor != null) {
				duplicate.UniColor = this.UniColor.createDuplicate();
			}
			return duplicate;
		};
		CBulletColor.prototype.Write_ToBinary = function (w) {
			w.WriteBool(isRealNumber(this.type));
			if (isRealNumber(this.type)) {
				w.WriteLong(this.type);
			}
			w.WriteBool(isRealObject(this.UniColor));
			if (isRealObject(this.UniColor)) {
				this.UniColor.Write_ToBinary(w);
			}

		};
		CBulletColor.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				(this.type) = r.GetLong();
			}

			if (r.GetBool()) {
				this.UniColor = new CUniColor();
				this.UniColor.Read_FromBinary(r);
			}
		};

		function CBulletSize(nType) {
			CBaseNoIdObject.call(this);
			this.type = AscFormat.isRealNumber(nType) ? nType : AscFormat.BULLET_TYPE_SIZE_NONE;
			this.val = 0;
		}

		InitClass(CBulletSize, CBaseNoIdObject, 0);
		CBulletSize.prototype.Set_FromObject = function (o) {
			this.merge(o);
		};
		CBulletSize.prototype.merge = function (oBulletSize) {
			if (!oBulletSize) {
				return;
			}
			this.type = oBulletSize.type;
			this.val = oBulletSize.val;
		};
		CBulletSize.prototype.createDuplicate = function () {
			var d = new CBulletSize();
			d.type = this.type;
			d.val = this.val;
			return d;
		};
		CBulletSize.prototype.IsIdentical = function (oBulletSize) {
			if (!oBulletSize) {
				return false;
			}
			return this.type === oBulletSize.type && this.val === oBulletSize.val;
		};
		CBulletSize.prototype.Write_ToBinary = function (w) {
			w.WriteBool(isRealNumber(this.type));
			if (isRealNumber(this.type)) {
				w.WriteLong(this.type);
			}

			w.WriteBool(isRealNumber(this.val));
			if (isRealNumber(this.val)) {
				w.WriteLong(this.val);
			}

		};
		CBulletSize.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				(this.type) = r.GetLong();
			}
			if (r.GetBool()) {
				(this.val) = r.GetLong();
			}
		};

		function CBulletTypeface(nType) {
			CBaseNoIdObject.call(this);
			this.type = AscFormat.isRealNumber(nType) ? nType : AscFormat.BULLET_TYPE_TYPEFACE_NONE;
			this.typeface = "";
		}

		InitClass(CBulletTypeface, CBaseNoIdObject, 0);
		CBulletTypeface.prototype.Set_FromObject = function (o) {
			this.merge(o);
		};
		CBulletTypeface.prototype.createDuplicate = function () {
			var d = new CBulletTypeface();
			d.type = this.type;
			d.typeface = this.typeface;
			return d;
		};
		CBulletTypeface.prototype.merge = function (oBulletTypeface) {
			if (!oBulletTypeface) {
				return;
			}
			this.type = oBulletTypeface.type;
			this.typeface = oBulletTypeface.typeface;
		};
		CBulletTypeface.prototype.IsIdentical = function (oBulletTypeface) {
			if (!oBulletTypeface) {
				return false;
			}
			return this.type === oBulletTypeface.type && this.typeface === oBulletTypeface.typeface;
		};
		CBulletTypeface.prototype.Write_ToBinary = function (w) {
			w.WriteBool(isRealNumber(this.type));
			if (isRealNumber(this.type)) {
				w.WriteLong(this.type);
			}

			w.WriteBool(typeof this.typeface === "string");
			if (typeof this.typeface === "string") {
				w.WriteString2(this.typeface);
			}

		};
		CBulletTypeface.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				(this.type) = r.GetLong();
			}
			if (r.GetBool()) {
				(this.typeface) = r.GetString2();
			}
		};


		var numbering_presentationnumfrmt_AlphaLcParenBoth = 0;
		var numbering_presentationnumfrmt_AlphaLcParenR = 1;
		var numbering_presentationnumfrmt_AlphaLcPeriod = 2;
		var numbering_presentationnumfrmt_AlphaUcParenBoth = 3;
		var numbering_presentationnumfrmt_AlphaUcParenR = 4;
		var numbering_presentationnumfrmt_AlphaUcPeriod = 5;
		var numbering_presentationnumfrmt_Arabic1Minus = 6;
		var numbering_presentationnumfrmt_Arabic2Minus = 7;
		var numbering_presentationnumfrmt_ArabicDbPeriod = 8;
		var numbering_presentationnumfrmt_ArabicDbPlain = 9;
		var numbering_presentationnumfrmt_ArabicParenBoth = 10;
		var numbering_presentationnumfrmt_ArabicParenR = 11;
		var numbering_presentationnumfrmt_ArabicPeriod = 12;
		var numbering_presentationnumfrmt_ArabicPlain = 13;
		var numbering_presentationnumfrmt_CircleNumDbPlain = 14;
		var numbering_presentationnumfrmt_CircleNumWdBlackPlain = 15;
		var numbering_presentationnumfrmt_CircleNumWdWhitePlain = 16;
		var numbering_presentationnumfrmt_Ea1ChsPeriod = 17;
		var numbering_presentationnumfrmt_Ea1ChsPlain = 18;
		var numbering_presentationnumfrmt_Ea1ChtPeriod = 19;
		var numbering_presentationnumfrmt_Ea1ChtPlain = 20;
		var numbering_presentationnumfrmt_Ea1JpnChsDbPeriod = 21;
		var numbering_presentationnumfrmt_Ea1JpnKorPeriod = 22;
		var numbering_presentationnumfrmt_Ea1JpnKorPlain = 23;
		var numbering_presentationnumfrmt_Hebrew2Minus = 24;
		var numbering_presentationnumfrmt_HindiAlpha1Period = 25;
		var numbering_presentationnumfrmt_HindiAlphaPeriod = 26;
		var numbering_presentationnumfrmt_HindiNumParenR = 27;
		var numbering_presentationnumfrmt_HindiNumPeriod = 28;
		var numbering_presentationnumfrmt_RomanLcParenBoth = 29;
		var numbering_presentationnumfrmt_RomanLcParenR = 30;
		var numbering_presentationnumfrmt_RomanLcPeriod = 31;
		var numbering_presentationnumfrmt_RomanUcParenBoth = 32;
		var numbering_presentationnumfrmt_RomanUcParenR = 33;
		var numbering_presentationnumfrmt_RomanUcPeriod = 34;
		var numbering_presentationnumfrmt_ThaiAlphaParenBoth = 35;
		var numbering_presentationnumfrmt_ThaiAlphaParenR = 36;
		var numbering_presentationnumfrmt_ThaiAlphaPeriod = 37;
		var numbering_presentationnumfrmt_ThaiNumParenBoth = 38;
		var numbering_presentationnumfrmt_ThaiNumParenR = 39;
		var numbering_presentationnumfrmt_ThaiNumPeriod = 40;

		var numbering_presentationnumfrmt_None = 100;
		var numbering_presentationnumfrmt_Char = 101;
		var numbering_presentationnumfrmt_Blip = 102;

		AscFormat.numbering_presentationnumfrmt_AlphaLcParenBoth = numbering_presentationnumfrmt_AlphaLcParenBoth;
		AscFormat.numbering_presentationnumfrmt_AlphaLcParenR = numbering_presentationnumfrmt_AlphaLcParenR;
		AscFormat.numbering_presentationnumfrmt_AlphaLcPeriod = numbering_presentationnumfrmt_AlphaLcPeriod;
		AscFormat.numbering_presentationnumfrmt_AlphaUcParenBoth = numbering_presentationnumfrmt_AlphaUcParenBoth;
		AscFormat.numbering_presentationnumfrmt_AlphaUcParenR = numbering_presentationnumfrmt_AlphaUcParenR;
		AscFormat.numbering_presentationnumfrmt_AlphaUcPeriod = numbering_presentationnumfrmt_AlphaUcPeriod;
		AscFormat.numbering_presentationnumfrmt_Arabic1Minus = numbering_presentationnumfrmt_Arabic1Minus;
		AscFormat.numbering_presentationnumfrmt_Arabic2Minus = numbering_presentationnumfrmt_Arabic2Minus;
		AscFormat.numbering_presentationnumfrmt_ArabicDbPeriod = numbering_presentationnumfrmt_ArabicDbPeriod;
		AscFormat.numbering_presentationnumfrmt_ArabicDbPlain = numbering_presentationnumfrmt_ArabicDbPlain;
		AscFormat.numbering_presentationnumfrmt_ArabicParenBoth = numbering_presentationnumfrmt_ArabicParenBoth;
		AscFormat.numbering_presentationnumfrmt_ArabicParenR = numbering_presentationnumfrmt_ArabicParenR;
		AscFormat.numbering_presentationnumfrmt_ArabicPeriod = numbering_presentationnumfrmt_ArabicPeriod;
		AscFormat.numbering_presentationnumfrmt_ArabicPlain = numbering_presentationnumfrmt_ArabicPlain;
		AscFormat.numbering_presentationnumfrmt_CircleNumDbPlain = numbering_presentationnumfrmt_CircleNumDbPlain;
		AscFormat.numbering_presentationnumfrmt_CircleNumWdBlackPlain = numbering_presentationnumfrmt_CircleNumWdBlackPlain;
		AscFormat.numbering_presentationnumfrmt_CircleNumWdWhitePlain = numbering_presentationnumfrmt_CircleNumWdWhitePlain;
		AscFormat.numbering_presentationnumfrmt_Ea1ChsPeriod = numbering_presentationnumfrmt_Ea1ChsPeriod;
		AscFormat.numbering_presentationnumfrmt_Ea1ChsPlain = numbering_presentationnumfrmt_Ea1ChsPlain;
		AscFormat.numbering_presentationnumfrmt_Ea1ChtPeriod = numbering_presentationnumfrmt_Ea1ChtPeriod;
		AscFormat.numbering_presentationnumfrmt_Ea1ChtPlain = numbering_presentationnumfrmt_Ea1ChtPlain;
		AscFormat.numbering_presentationnumfrmt_Ea1JpnChsDbPeriod = numbering_presentationnumfrmt_Ea1JpnChsDbPeriod;
		AscFormat.numbering_presentationnumfrmt_Ea1JpnKorPeriod = numbering_presentationnumfrmt_Ea1JpnKorPeriod;
		AscFormat.numbering_presentationnumfrmt_Ea1JpnKorPlain = numbering_presentationnumfrmt_Ea1JpnKorPlain;
		AscFormat.numbering_presentationnumfrmt_Hebrew2Minus = numbering_presentationnumfrmt_Hebrew2Minus;
		AscFormat.numbering_presentationnumfrmt_HindiAlpha1Period = numbering_presentationnumfrmt_HindiAlpha1Period;
		AscFormat.numbering_presentationnumfrmt_HindiAlphaPeriod = numbering_presentationnumfrmt_HindiAlphaPeriod;
		AscFormat.numbering_presentationnumfrmt_HindiNumParenR = numbering_presentationnumfrmt_HindiNumParenR;
		AscFormat.numbering_presentationnumfrmt_HindiNumPeriod = numbering_presentationnumfrmt_HindiNumPeriod;
		AscFormat.numbering_presentationnumfrmt_RomanLcParenBoth = numbering_presentationnumfrmt_RomanLcParenBoth;
		AscFormat.numbering_presentationnumfrmt_RomanLcParenR = numbering_presentationnumfrmt_RomanLcParenR;
		AscFormat.numbering_presentationnumfrmt_RomanLcPeriod = numbering_presentationnumfrmt_RomanLcPeriod;
		AscFormat.numbering_presentationnumfrmt_RomanUcParenBoth = numbering_presentationnumfrmt_RomanUcParenBoth;
		AscFormat.numbering_presentationnumfrmt_RomanUcParenR = numbering_presentationnumfrmt_RomanUcParenR;
		AscFormat.numbering_presentationnumfrmt_RomanUcPeriod = numbering_presentationnumfrmt_RomanUcPeriod;
		AscFormat.numbering_presentationnumfrmt_ThaiAlphaParenBoth = numbering_presentationnumfrmt_ThaiAlphaParenBoth;
		AscFormat.numbering_presentationnumfrmt_ThaiAlphaParenR = numbering_presentationnumfrmt_ThaiAlphaParenR;
		AscFormat.numbering_presentationnumfrmt_ThaiAlphaPeriod = numbering_presentationnumfrmt_ThaiAlphaPeriod;
		AscFormat.numbering_presentationnumfrmt_ThaiNumParenBoth = numbering_presentationnumfrmt_ThaiNumParenBoth;
		AscFormat.numbering_presentationnumfrmt_ThaiNumParenR = numbering_presentationnumfrmt_ThaiNumParenR;
		AscFormat.numbering_presentationnumfrmt_ThaiNumPeriod = numbering_presentationnumfrmt_ThaiNumPeriod;

		AscFormat.numbering_presentationnumfrmt_None = numbering_presentationnumfrmt_None;
		AscFormat.numbering_presentationnumfrmt_Char = numbering_presentationnumfrmt_Char;
		AscFormat.numbering_presentationnumfrmt_Blip = numbering_presentationnumfrmt_Blip;

		var MAP_AUTONUM_TYPES = {};
		MAP_AUTONUM_TYPES["alphaLcParenBot"] = numbering_presentationnumfrmt_AlphaLcParenBoth;
		MAP_AUTONUM_TYPES["alphaLcParen"] = numbering_presentationnumfrmt_AlphaLcParenR;
		MAP_AUTONUM_TYPES["alphaLcPerio"] = numbering_presentationnumfrmt_AlphaLcPeriod;
		MAP_AUTONUM_TYPES["alphaUcParenBot"] = numbering_presentationnumfrmt_AlphaUcParenBoth;
		MAP_AUTONUM_TYPES["alphaUcParen"] = numbering_presentationnumfrmt_AlphaUcParenR;
		MAP_AUTONUM_TYPES["alphaUcPerio"] = numbering_presentationnumfrmt_AlphaUcPeriod;
		MAP_AUTONUM_TYPES["arabic1Minu"] = numbering_presentationnumfrmt_Arabic1Minus;
		MAP_AUTONUM_TYPES["arabic2Minu"] = numbering_presentationnumfrmt_Arabic2Minus;
		MAP_AUTONUM_TYPES["arabicDbPerio"] = numbering_presentationnumfrmt_ArabicDbPeriod;
		MAP_AUTONUM_TYPES["arabicDbPlai"] = numbering_presentationnumfrmt_ArabicDbPlain;
		MAP_AUTONUM_TYPES["arabicParenBoth"] = numbering_presentationnumfrmt_ArabicParenBoth;
		MAP_AUTONUM_TYPES["arabicParenR"] = numbering_presentationnumfrmt_ArabicParenR;
		MAP_AUTONUM_TYPES["arabicPeriod"] = numbering_presentationnumfrmt_ArabicPeriod;
		MAP_AUTONUM_TYPES["arabicPlain"] = numbering_presentationnumfrmt_ArabicPlain;
		MAP_AUTONUM_TYPES["circleNumDbPlain"] = numbering_presentationnumfrmt_CircleNumDbPlain;
		MAP_AUTONUM_TYPES["circleNumWdBlackPlain"] = numbering_presentationnumfrmt_CircleNumWdBlackPlain;
		MAP_AUTONUM_TYPES["circleNumWdWhitePlain"] = numbering_presentationnumfrmt_CircleNumWdWhitePlain;
		MAP_AUTONUM_TYPES["ea1ChsPeriod"] = numbering_presentationnumfrmt_Ea1ChsPeriod;
		MAP_AUTONUM_TYPES["ea1ChsPlain"] = numbering_presentationnumfrmt_Ea1ChsPlain;
		MAP_AUTONUM_TYPES["ea1ChtPeriod"] = numbering_presentationnumfrmt_Ea1ChtPeriod;
		MAP_AUTONUM_TYPES["ea1ChtPlain"] = numbering_presentationnumfrmt_Ea1ChtPlain;
		MAP_AUTONUM_TYPES["ea1JpnChsDbPeriod"] = numbering_presentationnumfrmt_Ea1JpnChsDbPeriod;
		MAP_AUTONUM_TYPES["ea1JpnKorPeriod"] = numbering_presentationnumfrmt_Ea1JpnKorPeriod;
		MAP_AUTONUM_TYPES["ea1JpnKorPlain"] = numbering_presentationnumfrmt_Ea1JpnKorPlain;
		MAP_AUTONUM_TYPES["hebrew2Minus"] = numbering_presentationnumfrmt_Hebrew2Minus;
		MAP_AUTONUM_TYPES["hindiAlpha1Period"] = numbering_presentationnumfrmt_HindiAlpha1Period;
		MAP_AUTONUM_TYPES["hindiAlphaPeriod"] = numbering_presentationnumfrmt_HindiAlphaPeriod;
		MAP_AUTONUM_TYPES["hindiNumParenR"] = numbering_presentationnumfrmt_HindiNumParenR;
		MAP_AUTONUM_TYPES["hindiNumPeriod"] = numbering_presentationnumfrmt_HindiNumPeriod;
		MAP_AUTONUM_TYPES["romanLcParenBoth"] = numbering_presentationnumfrmt_RomanLcParenBoth;
		MAP_AUTONUM_TYPES["romanLcParenR"] = numbering_presentationnumfrmt_RomanLcParenR;
		MAP_AUTONUM_TYPES["romanLcPeriod"] = numbering_presentationnumfrmt_RomanLcPeriod;
		MAP_AUTONUM_TYPES["romanUcParenBoth"] = numbering_presentationnumfrmt_RomanUcParenBoth;
		MAP_AUTONUM_TYPES["romanUcParenR"] = numbering_presentationnumfrmt_RomanUcParenR;
		MAP_AUTONUM_TYPES["romanUcPeriod"] = numbering_presentationnumfrmt_RomanUcPeriod;
		MAP_AUTONUM_TYPES["thaiAlphaParenBoth"] = numbering_presentationnumfrmt_ThaiAlphaParenBoth;
		MAP_AUTONUM_TYPES["thaiAlphaParenR"] = numbering_presentationnumfrmt_ThaiAlphaParenR;
		MAP_AUTONUM_TYPES["thaiAlphaPeriod"] = numbering_presentationnumfrmt_ThaiAlphaPeriod;
		MAP_AUTONUM_TYPES["thaiNumParenBoth"] = numbering_presentationnumfrmt_ThaiNumParenBoth;
		MAP_AUTONUM_TYPES["thaiNumParenR"] = numbering_presentationnumfrmt_ThaiNumParenR;
		MAP_AUTONUM_TYPES["thaiNumPeriod"] = numbering_presentationnumfrmt_ThaiNumPeriod;

		function CBulletType(nType) {
			CBaseNoIdObject.call(this);
			this.type = AscFormat.isRealNumber(nType) ? nType : null;//BULLET_TYPE_BULLET_NONE;
			this.Char = null;
			this.AutoNumType = null;
			this.Blip = null;
			this.startAt = null;
		}

		InitClass(CBulletType, CBaseNoIdObject, 0);
		CBulletType.prototype.Set_FromObject = function (o) {
			this.merge(o);
		};
		CBulletType.prototype.IsIdentical = function (oBulletType) {
			if (!oBulletType) {
				return false;
			}
			return this.type === oBulletType.type
				&& this.Char === oBulletType.Char
				&& this.AutoNumType === oBulletType.AutoNumType
				&& this.startAt === oBulletType.startAt
				&& ((this.Blip && this.Blip.isEqual(oBulletType.Blip)) || this.Blip === oBulletType.Blip);
		};
		CBulletType.prototype.merge = function (oBulletType) {
			if (!oBulletType) {
				return;
			}
			if (oBulletType.type !== null && this.type !== oBulletType.type) {
				this.type = oBulletType.type;
				this.Char = oBulletType.Char;
				this.AutoNumType = oBulletType.AutoNumType;
				this.startAt = oBulletType.startAt;
				if (oBulletType.Blip) {
					this.Blip = oBulletType.Blip.createDuplicate();
				}
			} else {
				if (this.type === AscFormat.BULLET_TYPE_BULLET_CHAR) {
					if (typeof oBulletType.Char === "string"
						&& oBulletType.Char.length > 0) {
						if (this.Char !== oBulletType.Char) {
							this.Char = oBulletType.Char;
						}
					}
				}
				if (this.type === AscFormat.BULLET_TYPE_BULLET_BLIP) {
					if (this.Blip instanceof AscFormat.CBuBlip && this.Blip !== oBulletType.Blip) {
						this.Blip = oBulletType.Blip.createDuplicate();
					}
				}
				if (this.type === AscFormat.BULLET_TYPE_BULLET_AUTONUM) {
					if (oBulletType.AutoNumType !== null && this.AutoNumType !== oBulletType.AutoNumType) {
						this.AutoNumType = oBulletType.AutoNumType;
					}
					if (oBulletType.startAt !== null && this.startAt !== oBulletType.startAt) {
						this.startAt = oBulletType.startAt;
					}
				}

			}
		};
		CBulletType.prototype.createDuplicate = function () {
			var d = new CBulletType();
			d.type = this.type;
			d.Char = this.Char;
			d.AutoNumType = this.AutoNumType;
			d.startAt = this.startAt;
			if (this.Blip) {
				d.Blip = this.Blip.createDuplicate();
			}
			return d;
		};
		CBulletType.prototype.setBlip = function (oPr) {
			this.Blip = oPr;
		};
		CBulletType.prototype.Write_ToBinary = function (w) {
			w.WriteBool(isRealNumber(this.type));
			if (isRealNumber(this.type)) {
				w.WriteLong(this.type);
			}

			w.WriteBool(typeof this.Char === "string");
			if (typeof this.Char === "string") {
				w.WriteString2(this.Char);
			}

			w.WriteBool(isRealNumber(this.AutoNumType));
			if (isRealNumber(this.AutoNumType)) {
				w.WriteLong(this.AutoNumType);
			}

			w.WriteBool(isRealNumber(this.startAt));
			if (isRealNumber(this.startAt)) {
				w.WriteLong(this.startAt);
			}
			w.WriteBool(isRealObject(this.Blip));
			if (isRealObject(this.Blip)) {
				this.Blip.Write_ToBinary(w);
			}
		};
		CBulletType.prototype.Read_FromBinary = function (r) {
			if (r.GetBool()) {
				(this.type) = r.GetLong();
			}
			if (r.GetBool()) {
				(this.Char) = r.GetString2();
				if (AscFonts.IsCheckSymbols)
					AscFonts.FontPickerByCharacter.getFontsByString(this.Char);
			}

			if (r.GetBool()) {
				(this.AutoNumType) = r.GetLong();
			}

			if (r.GetBool()) {
				(this.startAt) = r.GetLong();
			}
			if (r.GetBool()) {
				this.Blip = new CBuBlip();
				this.Blip.Read_FromBinary(r);

				var oUnifill = this.Blip.blip;
				var sRasterImageId = oUnifill && oUnifill.fill && oUnifill.fill.RasterImageId;
				if (typeof AscCommon.CollaborativeEditing !== "undefined") {
					if (typeof sRasterImageId === "string" && sRasterImageId.length > 0) {
						AscCommon.CollaborativeEditing.Add_NewImage(sRasterImageId);
					}
				}
			}
		};

		function TextListStyle() {
			CBaseNoIdObject.call(this);
			this.levels = new Array(10);

			for (var i = 0; i < 10; i++)
				this.levels[i] = null;
		}

		InitClass(TextListStyle, CBaseNoIdObject, 0);
		TextListStyle.prototype.Get_Id = function () {
			return this.Id;
		};
		TextListStyle.prototype.Refresh_RecalcData = function () {
		};
		TextListStyle.prototype.createDuplicate = function () {
			var duplicate = new TextListStyle();
			for (var i = 0; i < 10; ++i) {
				if (this.levels[i] != null) {
					duplicate.levels[i] = this.levels[i].Copy();
				}
			}
			return duplicate;
		};
		TextListStyle.prototype.Write_ToBinary = function (w) {
			for (var i = 0; i < 10; ++i) {
				w.WriteBool(isRealObject(this.levels[i]));
				if (isRealObject(this.levels[i])) {
					this.levels[i].Write_ToBinary(w);
				}
			}
		};
		TextListStyle.prototype.Read_FromBinary = function (r) {
			for (var i = 0; i < 10; ++i) {
				if (r.GetBool()) {
					this.levels[i] = new CParaPr();
					this.levels[i].Read_FromBinary(r);
				} else {
					this.levels[i] = null;
				}
			}
		};
		TextListStyle.prototype.merge = function (oTextListStyle) {
			if (!oTextListStyle) {
				return;
			}
			for (var i = 0; i < this.levels.length; ++i) {
				if (oTextListStyle.levels[i]) {
					if (this.levels[i]) {
						this.levels[i].Merge(oTextListStyle.levels[i]);
					} else {
						this.levels[i] = oTextListStyle.levels[i].Copy();
					}
				}
			}
		};
		TextListStyle.prototype.Document_Get_AllFontNames = function (AllFonts) {
			for (var i = 0; i < 10; ++i) {
				if (this.levels[i]) {
					if (this.levels[i].DefaultRunPr) {
						this.levels[i].DefaultRunPr.Document_Get_AllFontNames(AllFonts);
					}
					if (this.levels[i].Bullet) {
						this.levels[i].Bullet.Get_AllFontNames(AllFonts);
					}
				}
			}
		};
		TextListStyle.prototype.applyParaPr = function (nLvl, ParaPr, bIncreaseFontSize, oSp) {

			let iN = AscFormat.isRealNumber;
			if(bIncreaseFontSize === true || bIncreaseFontSize === false) {
				if(!this.levels[nLvl]) {
					this.levels[nLvl] = new AscWord.CParaPr();
				}
				let oLvl = this.levels[nLvl];
				if(oLvl) {
					if(!oLvl.DefaultRunPr) {
						oLvl.DefaultRunPr = new AscWord.CTextPr();
					}
					let TextPr = oLvl.DefaultRunPr;
					let Pr;
					if(iN(TextPr.FontSize)) {
						TextPr.FontSize = TextPr.GetIncDecFontSize(bIncreaseFontSize)
					}
					else {
						let oSpStyles = oSp.Get_Styles(nLvl);
						let Styles      = oSpStyles.styles;
						Pr = Styles.Get_Pr(oSpStyles.lastId, styletype_Paragraph, null);
						TextPr.FontSize = Pr.TextPr.GetIncDecFontSize(bIncreaseFontSize)
					}
					if(iN(TextPr.FontSizeCS)) {
						TextPr.FontSizeCS = TextPr.GetIncDecFontSizeCS(bIncreaseFontSize)
					}

					else {
						if(!Pr) {
							let oSpStyles = oSp.Get_Styles(nLvl);
							let Styles      = oSpStyles.styles;
							Pr = Styles.Get_Pr(oSpStyles.lastId, styletype_Paragraph, null);
						}
						TextPr.FontSize = Pr.TextPr.GetIncDecFontSize(bIncreaseFontSize)
					}
				}
				return;
			}
			if(!ParaPr) {
				this.levels[nLvl] = null;
				return;
			}
			if(!this.levels[nLvl]) {
				this.levels[nLvl] = new AscWord.CParaPr();
			}
			let oLvl = this.levels[nLvl];
			oLvl.Merge(ParaPr);
			if(!oLvl.DefaultRunPr) {
				oLvl.DefaultRunPr = new AscWord.CTextPr();
			}

			if(ParaPr.DefaultRunPr) {
				let TextPr = ParaPr.DefaultRunPr;
				if (TextPr.FontFamily) {
					let FName = TextPr.FontFamily.Name;
					let FIndex = TextPr.FontFamily.Index;
					TextPr.RFonts = new CRFonts();
					TextPr.RFonts.SetAll(FName, FIndex);
				}
				oLvl.DefaultRunPr.Apply(ParaPr.DefaultRunPr);
			}
		};
		TextListStyle.prototype.changeFontSize = function (nLvl, bIncrease) {
			if(!this.levels[nLvl]) {
				this.levels[nLvl] = new AscWord.CParaPr();
			}
			let oLvl = this.levels[nLvl];
			if(oLvl.DefaultRunPr) {
				oLvl.DefaultRunPr = new AscWord.CTextPr();
			}

		};

		function CBaseAttrObject() {
			CBaseNoIdObject.call(this);
			this.attr = {};
		}

		InitClass(CBaseAttrObject, CBaseNoIdObject, 0);



		function CChangesCorePr(Class, Old, New, Color) {
			AscDFH.CChangesBase.call(this, Class, Old, New, Color);
			if (Old && New) {
				this.OldCategory = Old.category;
				this.OldContentStatus = Old.contentStatus;
				this.OldCreated = Old.created;
				this.OldCreator = Old.creator;
				this.OldDescription = Old.description;
				this.OldIdentifier = Old.identifier;
				this.OldKeywords = Old.keywords;
				this.OldLanguage = Old.language;
				this.OldLastModifiedBy = Old.lastModifiedBy;
				this.OldLastPrinted = Old.lastPrinted;
				this.OldModified = Old.modified;
				this.OldRevision = Old.revision;
				this.OldSubject = Old.subject;
				this.OldTitle = Old.title;
				this.OldVersion = Old.version;

				this.NewCategory = New.category === Old.category ? undefined : New.category;
				this.NewContentStatus = New.contentStatus === Old.contentStatus ? undefined : New.contentStatus;
				this.NewCreated = New.created === Old.created ? undefined : New.created;
				this.NewCreator = New.creator === Old.creator ? undefined : New.creator;
				this.NewDescription = New.description === Old.description ? undefined : New.description;
				this.NewIdentifier = New.identifier === Old.identifier ? undefined : New.identifier;
				this.NewKeywords = New.keywords === Old.keywords ? undefined : New.keywords;
				this.NewLanguage = New.language === Old.language ? undefined : New.language;
				this.NewLastModifiedBy = New.lastModifiedBy === Old.lastModifiedBy ? undefined : New.lastModifiedBy;
				this.NewLastPrinted = New.lastPrinted === Old.lastPrinted ? undefined : New.lastPrinted;
				this.NewModified = New.modified === Old.modified ? undefined : New.modified;
				this.NewRevision = New.revision === Old.revision ? undefined : New.revision;
				this.NewSubject = New.subject === Old.subject ? undefined : New.subject;
				this.NewTitle = New.title === Old.title ? undefined : New.title;
				this.NewVersion = New.version === Old.version ? undefined : New.version;
			} else {
				this.OldCategory = undefined;
				this.OldContentStatus = undefined;
				this.OldCreated = undefined;
				this.OldCreator = undefined;
				this.OldDescription = undefined;
				this.OldIdentifier = undefined;
				this.OldKeywords = undefined;
				this.OldLanguage = undefined;
				this.OldLastModifiedBy = undefined;
				this.OldLastPrinted= undefined;
				this.OldModified = undefined;
				this.OldRevision = undefined;
				this.OldSubject = undefined;
				this.OldTitle = undefined;
				this.OldVersion = undefined;

				this.NewCategory = undefined;
				this.NewContentStatus = undefined;
				this.NewCreated = undefined;
				this.NewCreator = undefined;
				this.NewDescription = undefined;
				this.NewIdentifier = undefined;
				this.NewKeywords = undefined;
				this.NewLanguage = undefined;
				this.NewLastModifiedBy = undefined;
				this.NewLastPrinted = undefined;
				this.NewModified = undefined;
				this.NewRevision = undefined;
				this.NewSubject = undefined;
				this.NewTitle = undefined;
				this.NewVersion = undefined;
			}
		}

		CChangesCorePr.prototype = Object.create(AscDFH.CChangesBase.prototype);
		CChangesCorePr.prototype.constructor = CChangesCorePr;
		CChangesCorePr.prototype.Type = AscDFH.historyitem_CoreProperties;
		CChangesCorePr.prototype.Undo = function () {
			if (!this.Class) {
				return;
			}
			this.Class.category = this.OldCategory;
			this.Class.contentStatus = this.OldContentStatus;
			this.Class.created = this.OldCreated;
			this.Class.creator = this.OldCreator;
			this.Class.description = this.OldDescription;
			this.Class.identifier = this.OldIdentifier;
			this.Class.keywords = this.OldKeywords;
			this.Class.language = this.OldLanguage;
			this.Class.lastModifiedBy = this.OldLastModifiedBy;
			this.Class.lastPrinted = this.OldLastPrinted;
			this.Class.modified = this.OldModified;
			this.Class.revision = this.OldRevision;
			this.Class.subject = this.OldSubject;
			this.Class.title = this.OldTitle;
			this.Class.version = this.OldVersion;
		};
		CChangesCorePr.prototype.Redo = function () {
			if (!this.Class) {
				return;
			}
			if (this.NewCategory !== undefined) {
				this.Class.category = this.NewCategory;
			}
			if (this.NewContentStatus !== undefined) {
				this.Class.contentStatus = this.NewContentStatus;
			}
			if (this.NewCreated !== undefined) {
				this.Class.created = this.NewCreated;
			}
			if (this.NewCreator !== undefined) {
				this.Class.creator = this.NewCreator;
			}
			if (this.NewDescription !== undefined) {
				this.Class.description = this.NewDescription;
			}
			if (this.NewIdentifier !== undefined) {
				this.Class.identifier = this.NewIdentifier;
			}
			if (this.NewKeywords !== undefined) {
				this.Class.keywords = this.NewKeywords;
			}
			if (this.NewLanguage !== undefined) {
				this.Class.language = this.NewLanguage;
			}
			if (this.NewLastModifiedBy !== undefined) {
				this.Class.lastModifiedBy = this.NewLastModifiedBy;
			}
			if (this.NewLastPrinted !== undefined) {
				this.Class.lastPrinted = this.NewLastPrinted;
			}
			if (this.NewModified !== undefined) {
				this.Class.modified = this.NewModified;
			}
			if (this.NewRevision !== undefined) {
				this.Class.revision = this.NewRevision;
			}
			if (this.NewSubject !== undefined) {
				this.Class.subject = this.NewSubject;
			}
			if (this.NewTitle !== undefined) {
				this.Class.title = this.NewTitle;
			}
			if (this.NewVersion !== undefined) {
				this.Class.version = this.NewVersion;
			}
		};
		CChangesCorePr.prototype.WriteToBinary = function (Writer) {
			var nFlags = 0;
			if (undefined !== this.NewTitle) {
				nFlags |= (1 << 0);
			}
			if (undefined !== this.NewCreator) {
				nFlags |= (1 << 1);
			}
			if (undefined !== this.NewDescription) {
				nFlags |= (1 << 2);
			}
			if (undefined !== this.NewSubject) {
				nFlags |= (1 << 3);
			}
			if (undefined !== this.NewCategory) {
				nFlags |= (1 << 4);
			}
			if (undefined !== this.NewContentStatus) {
				nFlags |= (1 << 5);
			}
			if (undefined !== this.NewCreated) {
				nFlags |= (1 << 6);
			}
			if (undefined !== this.NewIdentifier) {
				nFlags |= (1 << 7);
			}
			if (undefined !== this.NewKeywords) {
				nFlags |= (1 << 8);
			}
			if (undefined !== this.NewLanguage) {
				nFlags |= (1 << 9);
			}
			if (undefined !== this.NewLastModifiedBy) {
				nFlags |= (1 << 10);
			}
			if (undefined !== this.NewLastPrinted) {
				nFlags |= (1 << 11);
			}
			if (undefined !== this.NewModified) {
				nFlags |= (1 << 12);
			}
			if (undefined !== this.NewRevision) {
				nFlags |= (1 << 13);
			}
			if (undefined !== this.NewVersion) {
				nFlags |= (1 << 14);
			}

			Writer.WriteLong(nFlags);
			var bIsField;
			if (nFlags & (1 << 0)) {
				bIsField = typeof this.NewTitle === "string";
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewTitle);
				}
			}
			if (nFlags & (1 << 1)) {
				bIsField = typeof this.NewCreator === "string";
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewCreator);
				}
			}
			if (nFlags & (1 << 2)) {
				bIsField = typeof this.NewDescription === "string";
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewDescription);
				}
			}
			if (nFlags & (1 << 3)) {
				bIsField = typeof this.NewSubject === "string";
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewSubject);
				}
			}
			if (nFlags & (1 << 4)) {
				bIsField = typeof this.NewCategory === "string";
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewCategory);
				}
			}
			if (nFlags & (1 << 5)) {
				bIsField = typeof this.NewContentStatus === "string";
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewContentStatus);
				}
			}
			if (nFlags & (1 << 6)) {
				bIsField = this.NewCreated && this.NewCreated instanceof Date;
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewCreated.toISOString().slice(0, 19) + 'Z');
				}
			}
			if (nFlags & (1 << 7)) {
				bIsField = typeof this.NewIdentifier === "string";
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewIdentifier);
				}
			}
			if (nFlags & (1 << 8)) {
				bIsField = typeof this.NewKeywords === "string";
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewKeywords);
				}
			}
			if (nFlags & (1 << 9)) {
				bIsField = typeof this.NewLanguage === "string";
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewLanguage);
				}
			}
			if (nFlags & (1 << 10)) {
				bIsField = typeof this.NewLastModifiedBy === "string";
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewLastModifiedBy);
				}
			}
			if (nFlags & (1 << 11)) {
				bIsField = this.NewLastPrinted && this.NewLastPrinted instanceof Date;
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewLastPrinted.toISOString().slice(0, 19) + 'Z');
				}
			}
			if (nFlags & (1 << 12)) {
				bIsField = this.NewModified && this.NewModified instanceof Date;
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewModified.toISOString().slice(0, 19) + 'Z');
				}
			}
			if (nFlags & (1 << 13)) {
				bIsField = typeof this.NewRevision === "string";
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewRevision);
				}
			}
			if (nFlags & (1 << 14)) {
				bIsField = typeof this.NewVersion === "string";
				Writer.WriteBool(bIsField);
				if (bIsField) {
					Writer.WriteString2(this.NewVersion);
				}
			}
		};

		CChangesCorePr.prototype.ReadFromBinary = function (Reader) {
			var nFlags = Reader.GetLong();
			var bIsField;
			if (nFlags & (1 << 0)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewTitle = Reader.GetString2();
				} else {
					this.NewTitle = null;
				}
			}
			if (nFlags & (1 << 1)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewCreator = Reader.GetString2();
				} else {
					this.NewCreator = null;
				}
			}
			if (nFlags & (1 << 2)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewDescription = Reader.GetString2();
				} else {
					this.NewDescription = null;
				}
			}
			if (nFlags & (1 << 3)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewSubject = Reader.GetString2();
				} else {
					this.NewSubject = null;
				}
			}
			if (nFlags & (1 << 4)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewCategory = Reader.GetString2();
				} else {
					this.NewCategory = null;
				}
			}
			if (nFlags & (1 << 5)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewContentStatus = Reader.GetString2();
				} else {
					this.NewContentStatus = null;
				}
			}
			if (nFlags & (1 << 6)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewCreated = new Date(Reader.GetString2());
				} else {
					this.NewCreated = null;
				}
			}
			if (nFlags & (1 << 7)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewIdentifier = Reader.GetString2();
				} else {
					this.NewIdentifier = null;
				}
			}
			if (nFlags & (1 << 8)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewKeywords = Reader.GetString2();
				} else {
					this.NewKeywords = null;
				}
			}
			if (nFlags & (1 << 9)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewLanguage = Reader.GetString2();
				} else {
					this.NewLanguage = null;
				}
			}
			if (nFlags & (1 << 10)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewLastModifiedBy = Reader.GetString2();
				} else {
					this.NewLastModifiedBy = null;
				}
			}
			if (nFlags & (1 << 11)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewLastPrinted = new Date(Reader.GetString2());
				} else {
					this.NewLastPrinted = null;
				}
			}
			if (nFlags & (1 << 12)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewModified = new Date(Reader.GetString2());
				} else {
					this.NewModified = null;
				}
			}
			if (nFlags & (1 << 13)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewRevision = Reader.GetString2();
				} else {
					this.NewRevision = null;
				}
			}
			if (nFlags & (1 << 14)) {
				bIsField = Reader.GetBool();
				if (bIsField) {
					this.NewVersion = Reader.GetString2();
				} else {
					this.NewVersion = null;
				}
			}
		};
		CChangesCorePr.prototype.CreateReverseChange = function () {
			var ret = new CChangesCorePr(this.Class);
			ret.OldCategory = this.NewCategory;
			ret.OldContentStatus = this.NewContentStatus;
			ret.OldCreated = this.NewCreated;
			ret.OldCreator = this.NewCreator;
			ret.OldDescription = this.NewDescription;
			ret.OldIdentifier = this.NewIdentifier;
			ret.OldKeywords = this.NewKeywords;
			ret.OldLanguage = this.NewLanguage;
			ret.OldLastModifiedBy = this.NewLastModifiedBy;
			ret.OldLastPrinted = this.NewLastPrinted;
			ret.OldModified = this.NewModified;
			ret.OldRevision = this.NewRevision;
			ret.OldSubject = this.NewSubject;
			ret.OldTitle = this.NewTitle;
			ret.OldVersion = this.NewVersion;

			ret.NewCategory = this.OldCategory;
			ret.NewContentStatus = this.OldContentStatus;
			ret.NewCreated = this.OldCreated;
			ret.NewCreator = this.OldCreator;
			ret.NewDescription = this.OldDescription;
			ret.NewIdentifier = this.OldIdentifier;
			ret.NewKeywords = this.OldKeywords;
			ret.NewLanguage = this.OldLanguage;
			ret.NewLastModifiedBy = this.OldLastModifiedBy;
			ret.NewLastPrinted = this.OldLastPrinted;
			ret.NewModified = this.OldModified;
			ret.NewRevision = this.OldRevision;
			ret.NewSubject = this.OldSubject;
			ret.NewTitle = this.OldTitle;
			ret.NewVersion = this.OldVersion;
			return ret;
		};

		AscDFH.changesFactory[AscDFH.historyitem_CoreProperties] = CChangesCorePr;

		function CCore() {
			AscFormat.CBaseFormatObject.call(this);
			this.category = null;
			this.contentStatus = null;//Status in menu
			this.created = null;
			this.creator = null;// Authors in menu
			this.description = null;//Comments in menu
			this.identifier = null;
			this.keywords = null;
			this.language = null;
			this.lastModifiedBy = null;
			this.lastPrinted = null;
			this.modified = null;
			this.revision = null;
			this.subject = null;
			this.title = null;
			this.version = null;

			this.Lock = new AscCommon.CLock();
			this.lockType = AscCommon.c_oAscLockTypes.kLockTypeNone;
		}

		InitClass(CCore, CBaseFormatObject, AscDFH.historyitem_type_Core);
		CCore.prototype.fromStream = function (s) {
			var _type = s.GetUChar();
			var _len = s.GetULong();
			var _start_pos = s.cur;
			var _end_pos = _len + _start_pos;
			var _at;

			// attributes
			var _sa = s.GetUChar();

			while (true) {
				_at = s.GetUChar();

				if (_at === g_nodeAttributeEnd)
					break;

				switch (_at) {
					case 0: {
						this.title = s.GetString2();
						break;
					}
					case 1: {
						this.creator = s.GetString2();
						break;
					}
					case 2: {
						this.lastModifiedBy = s.GetString2();
						break;
					}
					case 3: {
						this.revision = s.GetString2();
						break;
					}
					case 4: {
						this.created = this.readDate(s.GetString2());
						break;
					}
					case 5: {
						this.modified = this.readDate(s.GetString2());
						break;
					}
					default:
						return;
				}
			}
			while (true) {
				if (s.cur >= _end_pos)
					break;

				_type = s.GetUChar();
				switch (_type) {
					case 0: {
						var _end_rec2 = s.cur + s.GetLong() + 4;
						s.Skip2(1); // start attributes
						while (true) {
							_at = s.GetUChar();
							if (_at === g_nodeAttributeEnd)
								break;

							switch (_at) {
								case 6: {
									this.category = s.GetString2();
									break;
								}
								case 7: {
									this.contentStatus = s.GetString2();
									break;
								}
								case 8: {
									this.description = s.GetString2();
									break;
								}
								case 9: {
									this.identifier = s.GetString2();
									break;
								}
								case 10: {
									this.keywords = s.GetString2();
									break;
								}
								case 11: {
									this.language = s.GetString2();
									break;
								}
								case 12: {
									this.lastPrinted = this.readDate(s.GetString2());
									break;
								}
								case 13: {
									this.subject = s.GetString2();
									break;
								}
								case 14: {
									this.version = s.GetString2();
									break;
								}
								default:
									return;
							}
						}
						s.Seek2(_end_rec2);
						break;
					}
					default: {
						s.SkipRecord();
						break;
					}
				}
			}
			s.Seek2(_end_pos);
		};
		CCore.prototype.readDate = function (val) {
			val = new Date(val);
			return val instanceof Date && !isNaN(val) ? val : null;
		};
		CCore.prototype.toStream = function (s, api) {
			s.StartRecord(AscCommon.c_oMainTables.Core);

			s.WriteUChar(AscCommon.g_nodeAttributeStart);

			s._WriteString2(0, this.title);
			s._WriteString2(1, this.creator);
			if (api && api.DocInfo) {
				s._WriteString2(2, api.DocInfo.get_UserName());
			}
			var revision = 0;
			if (this.revision) {
				var rev = parseInt(this.revision);
				if (!isNaN(rev)) {
					revision = rev;
				}
			}
			s._WriteString2(3, (revision + 1).toString());

			if (this.created) {
				s._WriteString2(4, this.created.toISOString().slice(0, 19) + 'Z');
			}
			s._WriteString2(5, new Date().toISOString().slice(0, 19) + 'Z');

			s.WriteUChar(g_nodeAttributeEnd);

			s.StartRecord(0);

			s.WriteUChar(AscCommon.g_nodeAttributeStart);

			s._WriteString2(6, this.category);
			s._WriteString2(7, this.contentStatus);
			s._WriteString2(8, this.description);
			s._WriteString2(9, this.identifier);
			s._WriteString2(10, this.keywords);
			s._WriteString2(11, this.language);
			// we don't track it
			// if (this.lastPrinted) {
			//     s._WriteString1(12, this.lastPrinted.toISOString().slice(0, 19) + 'Z');
			// }
			s._WriteString2(13, this.subject);
			s._WriteString2(14, this.version);

			s.WriteUChar(g_nodeAttributeEnd);

			s.EndRecord();

			s.EndRecord();
		};
		CCore.prototype.asc_getTitle = function () {
			return this.title;
		};
		CCore.prototype.asc_getCreator = function () {
			return this.creator;
		};
		CCore.prototype.asc_getLastModifiedBy = function () {
			return this.lastModifiedBy;
		};
		CCore.prototype.asc_getRevision = function () {
			return this.revision;
		};
		CCore.prototype.asc_getCreated = function () {
			return this.created;
		};
		CCore.prototype.asc_getModified = function () {
			return this.modified;
		};
		CCore.prototype.asc_getCategory = function () {
			return this.category;
		};
		CCore.prototype.asc_getContentStatus = function () {
			return this.contentStatus;
		};
		CCore.prototype.asc_getDescription = function () {
			return this.description;
		};
		CCore.prototype.asc_getIdentifier = function () {
			return this.identifier;
		};
		CCore.prototype.asc_getKeywords = function () {
			return this.keywords;
		};
		CCore.prototype.asc_getLanguage = function () {
			return this.language;
		};
		CCore.prototype.asc_getLastPrinted = function () {
			return this.lastPrinted;
		};
		CCore.prototype.asc_getSubject = function () {
			return this.subject;
		};
		CCore.prototype.asc_getVersion = function () {
			return this.version;
		};
		CCore.prototype.asc_putTitle = function (v) {
			this.title = v;
		};
		CCore.prototype.asc_putCreator = function (v) {
			this.creator = v;
		};
		CCore.prototype.asc_putLastModifiedBy = function (v) {
			this.lastModifiedBy = v;
		};
		CCore.prototype.asc_putRevision = function (v) {
			this.revision = v;
		};
		CCore.prototype.asc_putCreated = function (v) {
			this.created = v;
		};
		CCore.prototype.asc_putModified = function (v) {
			this.modified = v;
		};
		CCore.prototype.asc_putCategory = function (v) {
			this.category = v;
		};
		CCore.prototype.asc_putContentStatus = function (v) {
			this.contentStatus = v;
		};
		CCore.prototype.asc_putDescription = function (v) {
			this.description = v;
		};
		CCore.prototype.asc_putIdentifier = function (v) {
			this.identifier = v;
		};
		CCore.prototype.asc_putKeywords = function (v) {
			this.keywords = v;
		};
		CCore.prototype.asc_putLanguage = function (v) {
			this.language = v;
		};
		CCore.prototype.asc_putLastPrinted = function (v) {
			this.lastPrinted = v;
		};
		CCore.prototype.asc_putSubject = function (v) {
			this.subject = v;
		};
		CCore.prototype.asc_putVersion = function (v) {
			this.version = v;
		};
		CCore.prototype.setProps = function (oProps) {
			AscCommon.History.Add(new CChangesCorePr(this, this, oProps, null));
			this.category = oProps.category;
			this.contentStatus = oProps.contentStatus;
			this.created = oProps.created;
			this.creator = oProps.creator;
			this.description = oProps.description;
			this.identifier = oProps.identifier;
			this.keywords = oProps.keywords;
			this.language = oProps.language;
			this.lastModifiedBy = oProps.lastModifiedBy;
			this.lastPrinted = oProps.lastPrinted;
			this.modified = oProps.modified;
			this.revision = oProps.revision;
			this.subject = oProps.subject;
			this.title = oProps.title;
			this.version = oProps.version;
		};

		CCore.prototype.setCategory = function (sCategory) {
			const coreCopy = this.copy();
			coreCopy.asc_putCategory(sCategory);
			this.setProps(coreCopy);
		};
		CCore.prototype.setContentStatus = function (sContentStatus) {
			const coreCopy = this.copy();
			coreCopy.asc_putContentStatus(sContentStatus);
			this.setProps(coreCopy);
		};
		CCore.prototype.setCreated = function (oCreatedDate) {
			if (oCreatedDate instanceof Date && !isNaN(oCreatedDate)) {
				const coreCopy = this.copy();
				coreCopy.asc_putCreated(oCreatedDate);
				this.setProps(coreCopy);
			}
		};
		CCore.prototype.setCreator = function (sCreator) {
			const coreCopy = this.copy();
			coreCopy.asc_putCreator(sCreator);
			this.setProps(coreCopy);
		};
		CCore.prototype.setDescription = function (sDescription) {
			const coreCopy = this.copy();
			coreCopy.asc_putDescription(sDescription);
			this.setProps(coreCopy);
		};
		CCore.prototype.setIdentifier = function (sIdentifier) {
			const coreCopy = this.copy();
			coreCopy.asc_putIdentifier(sIdentifier);
			this.setProps(coreCopy);
		};
		CCore.prototype.setKeywords = function (sKeywords) {
			const coreCopy = this.copy();
			coreCopy.asc_putKeywords(sKeywords);
			this.setProps(coreCopy);
		};
		CCore.prototype.setLanguage = function (sLanguage) {
			const coreCopy = this.copy();
			coreCopy.asc_putLanguage(sLanguage);
			this.setProps(coreCopy);
		};
		CCore.prototype.setLastModifiedBy = function (sLastModifiedBy) {
			const coreCopy = this.copy();
			coreCopy.asc_putLastModifiedBy(sLastModifiedBy);
			this.setProps(coreCopy);
		};
		CCore.prototype.setLastPrinted = function (oLastPrintedDate) {
			if (oLastPrintedDate instanceof Date && !isNaN(oLastPrintedDate)) {
				const coreCopy = this.copy();
				coreCopy.asc_putLastPrinted(oLastPrintedDate);
				this.setProps(coreCopy);
			}
		};
		CCore.prototype.setModified = function (oModifiedDate) {
			if (oModifiedDate instanceof Date && !isNaN(oModifiedDate)) {
				const coreCopy = this.copy();
				coreCopy.asc_putModified(oModifiedDate);
				this.setProps(coreCopy);
			}
		};
		CCore.prototype.setRevision = function (sRevision) {
			const coreCopy = this.copy();
			coreCopy.asc_putRevision(sRevision);
			this.setProps(coreCopy);
		};
		CCore.prototype.setSubject = function (sSubject) {
			const coreCopy = this.copy();
			coreCopy.asc_putSubject(sSubject);
			this.setProps(coreCopy);
		};
		CCore.prototype.setTitle = function (sTitle) {
			const coreCopy = this.copy();
			coreCopy.asc_putTitle(sTitle);
			this.setProps(coreCopy);
		};
		CCore.prototype.setVersion = function (sVersion) {
			const coreCopy = this.copy();
			coreCopy.asc_putVersion(sVersion);
			this.setProps(coreCopy);
		};

		CCore.prototype.Refresh_RecalcData = function () {
		};
		CCore.prototype.Refresh_RecalcData2 = function () {
		};
		CCore.prototype.copy = function () {
			return AscFormat.ExecuteNoHistory(function () {
				var oCopy = new CCore();
				oCopy.category = this.category;
				oCopy.contentStatus = this.contentStatus;
				oCopy.created = this.created;
				oCopy.creator = this.creator;
				oCopy.description = this.description;
				oCopy.identifier = this.identifier;
				oCopy.keywords = this.keywords;
				oCopy.language = this.language;
				oCopy.lastModifiedBy = this.lastModifiedBy;
				oCopy.lastPrinted = this.lastPrinted;
				oCopy.modified = this.modified;
				oCopy.revision = this.revision;
				oCopy.subject = this.subject;
				oCopy.title = this.title;
				oCopy.version = this.version;
				return oCopy;
			}, this, []);
		};
		CCore.prototype.createDefaultPresentationEditor = function() {
			this.lastModifiedBy = "";
		};

		let DEFAULT_CREATOR = "CREATOR";
		let DEFAULT_LAST_MODIFIED_BY = "CREATOR";
		CCore.prototype.setRequiredDefaultsPresentationEditor = function() {
			if(!this.creator) {
				this.creator = DEFAULT_CREATOR;
			}
			if(!this.lastModifiedBy) {
				this.lastModifiedBy = DEFAULT_LAST_MODIFIED_BY;
			}
		};


		window['AscCommon'].CCore = CCore;
		prot = CCore.prototype;
		prot["asc_getTitle"] = prot.asc_getTitle;
		prot["asc_getCreator"] = prot.asc_getCreator;
		prot["asc_getLastModifiedBy"] = prot.asc_getLastModifiedBy;
		prot["asc_getRevision"] = prot.asc_getRevision;
		prot["asc_getCreated"] = prot.asc_getCreated;
		prot["asc_getModified"] = prot.asc_getModified;
		prot["asc_getCategory"] = prot.asc_getCategory;
		prot["asc_getContentStatus"] = prot.asc_getContentStatus;
		prot["asc_getDescription"] = prot.asc_getDescription;
		prot["asc_getIdentifier"] = prot.asc_getIdentifier;
		prot["asc_getKeywords"] = prot.asc_getKeywords;
		prot["asc_getLanguage"] = prot.asc_getLanguage;
		prot["asc_getLastPrinted"] = prot.asc_getLastPrinted;
		prot["asc_getSubject"] = prot.asc_getSubject;
		prot["asc_getVersion"] = prot.asc_getVersion;

		prot["asc_putTitle"] = prot.asc_putTitle;
		prot["asc_putCreator"] = prot.asc_putCreator;
		prot["asc_putLastModifiedBy"] = prot.asc_putLastModifiedBy;
		prot["asc_putRevision"] = prot.asc_putRevision;
		prot["asc_putCreated"] = prot.asc_putCreated;
		prot["asc_putModified"] = prot.asc_putModified;
		prot["asc_putCategory"] = prot.asc_putCategory;
		prot["asc_putContentStatus"] = prot.asc_putContentStatus;
		prot["asc_putDescription"] = prot.asc_putDescription;
		prot["asc_putIdentifier"] = prot.asc_putIdentifier;
		prot["asc_putKeywords"] = prot.asc_putKeywords;
		prot["asc_putLanguage"] = prot.asc_putLanguage;
		prot["asc_putLastPrinted"] = prot.asc_putLastPrinted;
		prot["asc_putSubject"] = prot.asc_putSubject;
		prot["asc_putVersion"] = prot.asc_putVersion;

		prot["setCategory"] = prot.setCategory;
		prot["setContentStatus"] = prot.setContentStatus;
		prot["setCreated"] = prot.setCreated;
		prot["setCreator"] = prot.setCreator;
		prot["setDescription"] = prot.setDescription;
		prot["setIdentifier"] = prot.setIdentifier;
		prot["setKeywords"] = prot.setKeywords;
		prot["setLanguage"] = prot.setLanguage;
		prot["setLastModifiedBy"] = prot.setLastModifiedBy;
		prot["setLastPrinted"] = prot.setLastPrinted;
		prot["setModified"] = prot.setModified;
		prot["setRevision"] = prot.setRevision;
		prot["setSubject"] = prot.setSubject;
		prot["setTitle"] = prot.setTitle;
		prot["setVersion"] = prot.setVersion;

		function PartTitle() {
			CBaseNoIdObject.call(this);
			this.title = null;
		}
		InitClass(PartTitle, CBaseNoIdObject, 0);

		function CApp() {
			CBaseNoIdObject.call(this);
			this.Template = null;
			this.TotalTime = null;
			this.Words = null;
			this.Application = null;
			this.PresentationFormat = null;
			this.Paragraphs = null;
			this.Slides = null;
			this.Notes = null;
			this.HiddenSlides = null;
			this.MMClips = null;
			this.ScaleCrop = null;
			this.HeadingPairs = [];
			this.TitlesOfParts = [];
			this.Company = null;
			this.LinksUpToDate = null;
			this.SharedDoc = null;
			this.HyperlinksChanged = null;
			this.AppVersion = null;

			this.Characters = null;
			this.CharactersWithSpaces = null;
			this.DocSecurity = null;
			this.HyperlinkBase = null;
			this.Lines = null;
			this.Manager = null;
			this.Pages = null;
		}

		InitClass(CApp, CBaseNoIdObject, 0);
		CApp.prototype.getAppName = function() {
			return AscCommon.g_cCompanyName + "/" + AscCommon.g_cProductVersion;
		};
		CApp.prototype.setRequiredDefaults = function() {
			this.Application = this.getAppName();
		};
		CApp.prototype.merge = function(oOtherApp) {
			oOtherApp.Template !== null && (this.Template = oOtherApp.Template);
			oOtherApp.TotalTime !== null && (this.TotalTime = oOtherApp.TotalTime);
			oOtherApp.Words !== null && (this.Words = oOtherApp.Words);
			oOtherApp.Application !== null && (this.Application = oOtherApp.Application);
			oOtherApp.PresentationFormat !== null && (this.PresentationFormat = oOtherApp.PresentationFormat);
			oOtherApp.Paragraphs !== null && (this.Paragraphs = oOtherApp.Paragraphs);
			//oOtherApp.Slides !== null && (this.Slides = oOtherApp.Slides);
			//oOtherApp.Notes !== null && (this.Notes = oOtherApp.Notes);
			oOtherApp.HiddenSlides !== null && (this.HiddenSlides = oOtherApp.HiddenSlides);
			oOtherApp.MMClips !== null && (this.MMClips = oOtherApp.MMClips);
			oOtherApp.ScaleCrop !== null && (this.ScaleCrop = oOtherApp.ScaleCrop);
			oOtherApp.Company !== null && (this.Company = oOtherApp.Company);
			oOtherApp.LinksUpToDate !== null && (this.LinksUpToDate = oOtherApp.LinksUpToDate);
			oOtherApp.SharedDoc !== null && (this.SharedDoc = oOtherApp.SharedDoc);
			oOtherApp.HyperlinksChanged !== null && (this.HyperlinksChanged = oOtherApp.HyperlinksChanged);
			oOtherApp.AppVersion !== null && (this.AppVersion = oOtherApp.AppVersion);

			oOtherApp.Characters !== null && (this.Characters = oOtherApp.Characters);
			oOtherApp.CharactersWithSpaces !== null && (this.CharactersWithSpaces = oOtherApp.CharactersWithSpaces);
			oOtherApp.DocSecurity !== null && (this.DocSecurity = oOtherApp.DocSecurity);
			oOtherApp.HyperlinkBase !== null && (this.HyperlinkBase = oOtherApp.HyperlinkBase);
			oOtherApp.Lines !== null && (this.Lines = oOtherApp.Lines);
			oOtherApp.Manager !== null && (this.Manager = oOtherApp.Manager);
			oOtherApp.Pages !== null && (this.Pages = oOtherApp.Pages);
		};
		CApp.prototype.createDefaultPresentationEditor = function(nCountSlides, nCountThemes) {
			this.TotalTime = 0;
			this.Words = 0;
			this.setRequiredDefaults();
			this.PresentationFormat = "On-screen Show (4:3)";
			this.Paragraphs = 0;
			this.Slides = nCountSlides;
			this.Notes = nCountSlides;
			this.HiddenSlides = 0;
			this.MMClips = 2;
			this.ScaleCrop = false;

			this.HeadingPairs.push(new CVariant());
			this.HeadingPairs[0].type = c_oVariantTypes.vtLpstr;
			this.HeadingPairs[0].strContent = "Theme";
			this.HeadingPairs.push(new CVariant());
			this.HeadingPairs[1].type = c_oVariantTypes.vtI4;
			this.HeadingPairs[1].iContent = nCountThemes;
			this.HeadingPairs.push(new CVariant());
			this.HeadingPairs[2].type = c_oVariantTypes.vtLpstr;
			this.HeadingPairs[2].strContent = "Slide Titles";
			this.HeadingPairs.push(new CVariant());
			this.HeadingPairs[3].type = c_oVariantTypes.vtI4;
			this.HeadingPairs[3].iContent = nCountSlides;

			for (let i = 0; i < nCountThemes; ++i) {
				let s = "Theme " + ( i + 1);
				this.TitlesOfParts.push(new PartTitle());
				this.TitlesOfParts[i].title = s;
			}

			for (let i = 0; i < nCountSlides; ++i) {
				let s = "Slide " + (i + 1);
				this.TitlesOfParts.push( new PartTitle());
				this.TitlesOfParts[nCountThemes + i].title = s;
			}

			this.LinksUpToDate = false;
			this.SharedDoc = false;
			this.HyperlinksChanged = false;
		};
		CApp.prototype.fromStream = function (s) {
			var _type = s.GetUChar();
			var _len = s.GetULong();
			var _start_pos = s.cur;
			var _end_pos = _len + _start_pos;
			var _at;

			// attributes
			var _sa = s.GetUChar();

			while (true) {
				_at = s.GetUChar();

				if (_at === g_nodeAttributeEnd)
					break;

				switch (_at) {
					case 0: {
						this.Template = s.GetString2();
						break;
					}
					case 1: {
						this.Application = s.GetString2();
						break;
					}
					case 2: {
						this.PresentationFormat = s.GetString2();
						break;
					}
					case 3: {
						this.Company = s.GetString2();
						break;
					}
					case 4: {
						this.AppVersion = s.GetString2();
						break;
					}

					case 5: {
						this.TotalTime = s.GetLong();
						break;
					}
					case 6: {
						this.Words = s.GetLong();
						break;
					}
					case 7: {
						this.Paragraphs = s.GetLong();
						break;
					}
					case 8: {
						this.Slides = s.GetLong();
						break;
					}
					case 9: {
						this.Notes = s.GetLong();
						break;
					}
					case 10: {
						this.HiddenSlides = s.GetLong();
						break;
					}
					case 11: {
						this.MMClips = s.GetLong();
						break;
					}

					case 12: {
						this.ScaleCrop = s.GetBool();
						break;
					}
					case 13: {
						this.LinksUpToDate = s.GetBool();
						break;
					}
					case 14: {
						this.SharedDoc = s.GetBool();
						break;
					}
					case 15: {
						this.HyperlinksChanged = s.GetBool();
						break;
					}
					default:
						return;
				}
			}
			while (true) {
				if (s.cur >= _end_pos)
					break;

				_type = s.GetUChar();
				switch (_type) {
					case 0: {
						var _end_rec2 = s.cur + s.GetLong() + 4;
						s.Skip2(1); // start attributes
						while (true) {
							_at = s.GetUChar();
							if (_at === g_nodeAttributeEnd)
								break;

							switch (_at) {
								case 16: {
									this.Characters = s.GetLong();
									break;
								}
								case 17: {
									this.CharactersWithSpaces = s.GetLong();
									break;
								}
								case 18: {
									this.DocSecurity = s.GetLong();
									break;
								}
								case 19: {
									this.HyperlinkBase = s.GetString2();
									break;
								}
								case 20: {
									this.Lines = s.GetLong();
									break;
								}
								case 21: {
									this.Manager = s.GetString2();
									break;
								}
								case 22: {
									this.Pages = s.GetLong();
									break;
								}
								default:
									return;
							}
						}
						s.Seek2(_end_rec2);
						break;
					}
					default: {
						s.SkipRecord();
						break;
					}
				}
			}
			s.Seek2(_end_pos);
		};
		CApp.prototype.toStream = function (s) {
			s.StartRecord(AscCommon.c_oMainTables.App);

			s.WriteUChar(AscCommon.g_nodeAttributeStart);

			s._WriteString2(0, this.Template);
			// just in case
			// s._WriteString2(1, this.Application);
			s._WriteString2(2, this.PresentationFormat);
			s._WriteString2(3, this.Company);
			// just in case
			// s._WriteString2(4, this.AppVersion);

			//we don't count these stats
			// s._WriteInt2(5, this.TotalTime);
			// s._WriteInt2(6, this.Words);
			// s._WriteInt2(7, this.Paragraphs);
			// s._WriteInt2(8, this.Slides);
			// s._WriteInt2(9, this.Notes);
			// s._WriteInt2(10, this.HiddenSlides);
			// s._WriteInt2(11, this.MMClips);

			s._WriteBool2(12, this.ScaleCrop);
			s._WriteBool2(13, this.LinksUpToDate);
			s._WriteBool2(14, this.SharedDoc);
			s._WriteBool2(15, this.HyperlinksChanged);

			s.WriteUChar(g_nodeAttributeEnd);

			s.StartRecord(0);

			s.WriteUChar(AscCommon.g_nodeAttributeStart);

			// s._WriteInt2(16, this.Characters);
			// s._WriteInt2(17, this.CharactersWithSpaces);
			s._WriteInt2(18, this.DocSecurity);
			s._WriteString2(19, this.HyperlinkBase);
			// s._WriteInt2(20, this.Lines);
			s._WriteString2(21, this.Manager);
			// s._WriteInt2(22, this.Pages);

			s.WriteUChar(g_nodeAttributeEnd);

			s.EndRecord();

			s.EndRecord();
		};
		CApp.prototype.asc_getTemplate = function () {
			return this.Template;
		};
		CApp.prototype.asc_getTotalTime = function () {
			return this.TotalTime;
		};
		CApp.prototype.asc_getWords = function () {
			return this.Words;
		};
		CApp.prototype.asc_getApplication = function () {
			return this.Application;
		};
		CApp.prototype.asc_getPresentationFormat = function () {
			return this.PresentationFormat;
		};
		CApp.prototype.asc_getParagraphs = function () {
			return this.Paragraphs;
		};
		CApp.prototype.asc_getSlides = function () {
			return this.Slides;
		};
		CApp.prototype.asc_getNotes = function () {
			return this.Notes;
		};
		CApp.prototype.asc_getHiddenSlides = function () {
			return this.HiddenSlides;
		};
		CApp.prototype.asc_getMMClips = function () {
			return this.MMClips;
		};
		CApp.prototype.asc_getScaleCrop = function () {
			return this.ScaleCrop;
		};
		CApp.prototype.asc_getCompany = function () {
			return this.Company;
		};
		CApp.prototype.asc_getLinksUpToDate = function () {
			return this.LinksUpToDate;
		};
		CApp.prototype.asc_getSharedDoc = function () {
			return this.SharedDoc;
		};
		CApp.prototype.asc_getHyperlinksChanged = function () {
			return this.HyperlinksChanged;
		};
		CApp.prototype.asc_getAppVersion = function () {
			return this.AppVersion;
		};
		CApp.prototype.asc_getCharacters = function () {
			return this.Characters;
		};
		CApp.prototype.asc_getCharactersWithSpaces = function () {
			return this.CharactersWithSpaces;
		};
		CApp.prototype.asc_getDocSecurity = function () {
			return this.DocSecurity;
		};
		CApp.prototype.asc_getHyperlinkBase = function () {
			return this.HyperlinkBase;
		};
		CApp.prototype.asc_getLines = function () {
			return this.Lines;
		};
		CApp.prototype.asc_getManager = function () {
			return this.Manager;
		};
		CApp.prototype.asc_getPages = function () {
			return this.Pages;
		};
		window['AscCommon'].CApp = CApp;
		prot = CApp.prototype;
		prot["asc_getTemplate"] = prot.asc_getTemplate;
		prot["asc_getTotalTime"] = prot.asc_getTotalTime;
		prot["asc_getWords"] = prot.asc_getWords;
		prot["asc_getApplication"] = prot.asc_getApplication;
		prot["asc_getPresentationFormat"] = prot.asc_getPresentationFormat;
		prot["asc_getParagraphs"] = prot.asc_getParagraphs;
		prot["asc_getSlides"] = prot.asc_getSlides;
		prot["asc_getNotes"] = prot.asc_getNotes;
		prot["asc_getHiddenSlides"] = prot.asc_getHiddenSlides;
		prot["asc_getMMClips"] = prot.asc_getMMClips;
		prot["asc_getScaleCrop"] = prot.asc_getScaleCrop;
		prot["asc_getCompany"] = prot.asc_getCompany;
		prot["asc_getLinksUpToDate"] = prot.asc_getLinksUpToDate;
		prot["asc_getSharedDoc"] = prot.asc_getSharedDoc;
		prot["asc_getHyperlinksChanged"] = prot.asc_getHyperlinksChanged;
		prot["asc_getAppVersion"] = prot.asc_getAppVersion;
		prot["asc_getCharacters"] = prot.asc_getCharacters;
		prot["asc_getCharactersWithSpaces"] = prot.asc_getCharactersWithSpaces;
		prot["asc_getDocSecurity"] = prot.asc_getDocSecurity;
		prot["asc_getHyperlinkBase"] = prot.asc_getHyperlinkBase;
		prot["asc_getLines"] = prot.asc_getLines;
		prot["asc_getManager"] = prot.asc_getManager;
		prot["asc_getPages"] = prot.asc_getPages;

		function CCustomProperties() {
			CBaseObject.call(this);
			this.properties = [];
			this.Lock = new AscCommon.CLock();
		}

		InitClass(CCustomProperties, CBaseObject, AscDFH.historyitem_type_CustomProperties);
		CCustomProperties.prototype.fromStream = function (s) {
			var _type = s.GetUChar();
			var _len = s.GetULong();
			var _start_pos = s.cur;
			var _end_pos = _len + _start_pos;
			var _at;

			// attributes
			var _sa = s.GetUChar();

			while (true) {
				_at = s.GetUChar();

				if (_at === g_nodeAttributeEnd)
					break;
			}
			while (true) {
				if (s.cur >= _end_pos)
					break;

				_type = s.GetUChar();
				switch (_type) {
					case 0: {
						s.Skip2(4);
						var _c = s.GetULong();
						for (var i = 0; i < _c; ++i) {
							s.Skip2(1); // type
							var tmp = new CCustomProperty();
							tmp.fromStream(s);
							this.properties.push(tmp);
						}
						break;
					}
					default: {
						s.SkipRecord();
						break;
					}
				}
			}
			s.Seek2(_end_pos);
		};
		CCustomProperties.prototype.toStream = function (s) {
			s.StartRecord(AscCommon.c_oMainTables.CustomProperties);
			s.WriteUChar(AscCommon.g_nodeAttributeStart);
			s.WriteUChar(g_nodeAttributeEnd);

			this.fillNewPid();
			s.WriteRecordArray4(0, 0, this.properties);
			s.EndRecord();
		};
		CCustomProperties.prototype.fillNewPid = function (s) {
			var index = 2;
			this.properties.forEach(function (property) {
				property.pid = index++;
			});
		};
		CCustomProperties.prototype.add = function (name, variant, opt_linkTarget) {
			this.addProperty(this.properties.length, this.createPropertyWithVariant(name, variant, opt_linkTarget));
		};
		CCustomProperties.prototype.createPropertyWithVariant = function (name, variant, opt_linkTarget) {
			var newProperty = new CCustomProperty();
			newProperty.fmtid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
			newProperty.pid = null;
			newProperty.name = name;
			newProperty.linkTarget = opt_linkTarget || null;
			newProperty.content = variant;
			return newProperty;
		};
		CCustomProperties.prototype.createProperty = function (name, type, value, opt_linkTarget) {
			let oVariant = new CVariant();
			oVariant.setType(type);
			oVariant.setValue(value);
			return this.createPropertyWithVariant(name, oVariant, opt_linkTarget);
		};
		CCustomProperties.prototype.getAllProperties = function () {
			return [].concat(this.properties);
		};
		CCustomProperties.prototype.addProperty = function (idx, pr) {
			let nInsertIdx = Math.min(this.properties.length, Math.max(0, idx));
			AscCommon.History.Add(new CChangesDrawingsContentNoId(this, AscDFH.historyitem_CustomPropertiesAddProperty, nInsertIdx, [pr], true));
			this.properties.splice(nInsertIdx, 0, pr);
		};
		CCustomProperties.prototype.removeProperty = function (idx) {
			if(idx < 0 || idx >= this.properties.length)
				return;
			let aDeleteElems = this.properties.splice(idx, 1);
			AscCommon.History.Add(new CChangesDrawingsContentNoId(this, AscDFH.historyitem_CustomPropertiesAddProperty, idx, aDeleteElems, false));
		};
		CCustomProperties.prototype.modifyProperty = function (idx, pr) {
			if(idx < 0 || idx >= this.properties.length)
				return;
			this.removeProperty(idx);
			this.addProperty(idx, pr);
		};
		CCustomProperties.prototype.AddProperty = function (name, type, value) {
			for(let nIdx = 0; nIdx < this.properties.length; ++nIdx) {
				let oPr = this.properties[nIdx];
				if(oPr.name === name) {
					this.ModifyProperty(nIdx, name, type, value);
					return;
				}
			}
			this.addProperty(this.properties.length, this.createProperty(name, type, value, null));
		};
		CCustomProperties.prototype.ModifyProperty = function (idx, name, type, value) {
			this.modifyProperty(idx, this.createProperty(name, type, value, null));
		};
		CCustomProperties.prototype.RemoveProperty = function (idx) {
			return this.removeProperty(idx);
		};
		CCustomProperties.prototype.hasProperties = function () {
			return this.properties.length > 0;
		};

		window['AscCommon'].CCustomProperties = CCustomProperties;
		prot = CCustomProperties.prototype;
		prot["add"] = prot.add;


		function CCustomProperty() {
			CBaseNoIdObject.call(this);
			this.fmtid = null;
			this.pid = null;
			this.name = null;
			this.linkTarget = null;

			this.content = null;
		}

		InitClass(CCustomProperty, CBaseNoIdObject, 0);
		CCustomProperty.prototype.fromStream = function (s) {
			var _type;
			var _len = s.GetULong();
			var _start_pos = s.cur;
			var _end_pos = _len + _start_pos;
			var _at;

			// attributes
			var _sa = s.GetUChar();

			while (true) {
				_at = s.GetUChar();

				if (_at === g_nodeAttributeEnd)
					break;

				switch (_at) {
					case 0: {
						this.fmtid = s.GetString2();
						break;
					}
					case 1: {
						this.pid = s.GetLong();
						break;
					}
					case 2: {
						this.name = s.GetString2();
						break;
					}
					case 3: {
						this.linkTarget = s.GetString2();
						break;
					}
					default:
						return;
				}
			}
			while (true) {
				if (s.cur >= _end_pos)
					break;

				_type = s.GetUChar();
				switch (_type) {
					case 0: {
						this.content = new CVariant(this);
						this.content.fromStream(s);
						break;
					}
					default: {
						s.SkipRecord();
						break;
					}
				}
			}
			s.Seek2(_end_pos);
		};
		CCustomProperty.prototype.toStream = function (s) {
			s.WriteUChar(AscCommon.g_nodeAttributeStart);
			s._WriteString2(0, this.fmtid);
			s._WriteInt2(1, this.pid);
			s._WriteString2(2, this.name);
			s._WriteString2(3, this.linkTarget);
			s.WriteUChar(g_nodeAttributeEnd);

			s.WriteRecord4(0, this.content);
		};
		CCustomProperty.prototype.setContent = function (v) {
			this.content = v;
		};
		CCustomProperty.prototype.asc_getName = function() {
			return this.name;
		};
		CCustomProperty.prototype.asc_getType = function() {
			if(this.content) {
				return this.content.getVariantType();
			}
			return c_oVariantTypes.vtEmpty;
		};
		CCustomProperty.prototype.asc_getValue = function() {
			if(this.content) {
				return this.content.getValue();
			}
			return null;
		};
		CCustomProperty.prototype.Write_ToBinary = function(w) {
			let oStream = AscCommon.pptx_content_writer.BinaryFileWriter;
			var old = new AscCommon.CMemory(true);
            oStream.ExportToMemory(old);
            oStream.ImportFromMemory(w);
            oStream.WriteRecord4(0, this);
            oStream.ExportToMemory(w);
            oStream.ImportFromMemory(old);
		};
		CCustomProperty.prototype.Read_FromBinary = function(r) {
			let fileStream = r.ToFileStream();
			fileStream.GetUChar();
			this.fromStream(fileStream);
			r.FromFileStream(fileStream);
		};
		
		CCustomProperty.prototype["asc_getName"] = CCustomProperty.prototype.asc_getName;
		CCustomProperty.prototype["asc_getType"] = CCustomProperty.prototype.asc_getType;
		CCustomProperty.prototype["asc_getValue"] = CCustomProperty.prototype.asc_getValue;

		function CVariantVector() {
			CBaseNoIdObject.call(this);
			this.baseType = null;
			this.size = null;

			this.variants = [];
		}

		InitClass(CVariantVector, CBaseNoIdObject, 0);
		CVariantVector.prototype.fromStream = function (s) {
			var _type;
			var _len = s.GetULong();
			var _start_pos = s.cur;
			var _end_pos = _len + _start_pos;
			var _at;

			// attributes
			var _sa = s.GetUChar();

			while (true) {
				_at = s.GetUChar();

				if (_at === g_nodeAttributeEnd)
					break;

				switch (_at) {
					case 0: {
						this.baseType = s.GetUChar();
						break;
					}
					case 1: {
						this.size = s.GetLong();
						break;
					}
					default:
						return;
				}
			}
			while (true) {
				if (s.cur >= _end_pos)
					break;

				_type = s.GetUChar();
				switch (_type) {
					case 0: {
						s.Skip2(4);
						var _c = s.GetULong();
						for (var i = 0; i < _c; ++i) {
							s.Skip2(1); // type
							var tmp = new CVariant(this);
							tmp.fromStream(s);
							this.variants.push(tmp);
						}
						break;
					}
					default: {
						s.SkipRecord();
						break;
					}
				}
			}
			s.Seek2(_end_pos);
		};
		CVariantVector.prototype.toStream = function (s) {
			s.WriteUChar(AscCommon.g_nodeAttributeStart);
			s._WriteUChar2(0, this.baseType);
			s._WriteInt2(1, this.size);
			s.WriteUChar(g_nodeAttributeEnd);

			s.WriteRecordArray4(0, 0, this.variants);
		};

		CVariantVector.prototype.getVariantType = function () {
			return AscFormat.isRealNumber(this.baseType) ? this.baseType : c_oVariantTypes.vtEmpty;
		};

		function CVariantArray() {
			CBaseNoIdObject.call(this);
			this.baseType = null;
			this.lBounds = null;
			this.uBounds = null;

			this.variants = [];
		}

		InitClass(CVariantArray, CBaseNoIdObject, 0);
		CVariantArray.prototype.fromStream = function (s) {
			var _type;
			var _len = s.GetULong();
			var _start_pos = s.cur;
			var _end_pos = _len + _start_pos;
			var _at;

			// attributes
			var _sa = s.GetUChar();

			while (true) {
				_at = s.GetUChar();

				if (_at === g_nodeAttributeEnd)
					break;

				switch (_at) {
					case 0: {
						this.baseType = s.GetUChar();
						break;
					}
					case 1: {
						this.lBounds = s.GetString2();
						break;
					}
					case 2: {
						this.uBounds = s.GetString2();
						break;
					}
					default:
						return;
				}
			}
			while (true) {
				if (s.cur >= _end_pos)
					break;

				_type = s.GetUChar();
				switch (_type) {
					case 0: {
						s.Skip2(4);
						var _c = s.GetULong();
						for (var i = 0; i < _c; ++i) {
							s.Skip2(1); // type
							var tmp = new CVariant();
							tmp.fromStream(s);
							this.variants.push(tmp);
						}
						break;
					}
					default: {
						s.SkipRecord();
						break;
					}
				}
			}
			s.Seek2(_end_pos);
		};
		CVariantArray.prototype.toStream = function (s) {
			s.WriteUChar(AscCommon.g_nodeAttributeStart);
			s._WriteUChar2(0, this.baseType);
			s._WriteString2(1, this.lBounds);
			s._WriteString2(2, this.uBounds);
			s.WriteUChar(g_nodeAttributeEnd);

			s.WriteRecordArray4(0, 0, this.variants);
		};
		CVariantArray.prototype.getVariantType = function () {
			return AscFormat.isRealNumber(this.baseType) ? this.baseType : c_oVariantTypes.vtEmpty;
		};

		function CVariantVStream() {
			CBaseNoIdObject.call(this);
			this.version = null;

			this.strContent = null;
		}

		InitClass(CVariantVStream, CBaseNoIdObject, 0);
		CVariantVStream.prototype.fromStream = function (s) {
			var _type;
			var _len = s.GetULong();
			var _start_pos = s.cur;
			var _end_pos = _len + _start_pos;
			var _at;

			// attributes
			var _sa = s.GetUChar();

			while (true) {
				_at = s.GetUChar();

				if (_at === g_nodeAttributeEnd)
					break;

				switch (_at) {
					case 0: {
						this.version = s.GetString2();
						break;
					}
					default:
						return;
				}
			}
			while (true) {
				if (s.cur >= _end_pos)
					break;

				_type = s.GetUChar();
				switch (_type) {
					case 0: {
						this.strContent = s.GetString2();
						break;
					}
					default: {
						s.SkipRecord();
						break;
					}
				}
			}
			s.Seek2(_end_pos);
		};
		CVariantVStream.prototype.toStream = function (s) {
			s.WriteUChar(AscCommon.g_nodeAttributeStart);
			s._WriteString2(0, this.version);
			s.WriteUChar(g_nodeAttributeEnd);

			s._WriteString2(0, this.strContent);
		};

		function CVariant(parent) {
			this.type = null;
			this.strContent = null;
			this.iContent = null;
			this.uContent = null;
			this.dContent = null;
			this.bContent = null;
			this.variant = null;
			this.vector = null;
			this.array = null;
			this.vStream = null;
			this.parent = parent;
		}
		CVariant.prototype.fromStream = function (s) {
			var _type;
			var _len = s.GetULong();
			var _start_pos = s.cur;
			var _end_pos = _len + _start_pos;
			var _at;

			// attributes
			var _sa = s.GetUChar();

			while (true) {
				_at = s.GetUChar();

				if (_at === g_nodeAttributeEnd)
					break;

				switch (_at) {
					case 0: {
						this.type = s.GetUChar();
						break;
					}
					default:
						return;
				}
			}
			while (true) {
				if (s.cur >= _end_pos)
					break;

				_type = s.GetUChar();
				switch (_type) {
					case 0: {
						this.strContent = s.GetString2();
						break;
					}
					case 1: {
						this.iContent = s.GetLong();
						break;
					}
					case 2: {
						this.iContent = s.GetULong();
						break;
					}
					case 3: {
						this.dContent = s.GetDouble();
						break;
					}
					case 4: {
						this.bContent = s.GetBool();
						break;
					}
					case 5: {
						this.variant = new CVariant(this);
						this.variant.fromStream(s);
						break;
					}
					case 6: {
						this.vector = new CVariantVector();
						this.vector.fromStream(s);
						break;
					}
					case 7: {
						this.array = new CVariantArray();
						this.array.fromStream(s);
						break;
					}
					case 8: {
						this.vStream = new CVariantVStream();
						this.vStream.fromStream(s);
						break;
					}
					default: {
						s.SkipRecord();
						break;
					}
				}
			}
			s.Seek2(_end_pos);
		};
		CVariant.prototype.toStream = function (s) {
			s.WriteUChar(AscCommon.g_nodeAttributeStart);
			s._WriteUChar2(0, this.type);
			s.WriteUChar(g_nodeAttributeEnd);

			s._WriteString2(0, this.strContent);
			s._WriteInt2(1, this.iContent);
			s._WriteUInt2(2, this.uContent);
			s._WriteDoubleReal2(3, this.dContent);
			s._WriteBool2(4, this.bContent);
			s.WriteRecord4(5, this.variant);
			s.WriteRecord4(6, this.vector);
			s.WriteRecord4(7, this.array);
			s.WriteRecord4(8, this.vStream);
		};
		CVariant.prototype.setParent = function(pr) {
			this.parent = pr;
		};
		CVariant.prototype.setType = function(pr) {
			this.type = pr;
		};
		CVariant.prototype.setStrContent = function(pr) {
			this.strContent = pr;
		};
		CVariant.prototype.setIContent = function(pr) {
			this.iContent = pr;
		};
		CVariant.prototype.setUContent = function(pr) {
			this.uContent = pr;
		};
		CVariant.prototype.setDContent = function(pr) {
			this.dContent = pr;
		};
		CVariant.prototype.setBContent = function(pr) {
			this.bContent = pr;
		};
		CVariant.prototype.setVariant = function(pr) {
			this.variant = pr;
		};
		CVariant.prototype.setVector = function(pr) {
			this.vector = pr;
		};
		CVariant.prototype.setArray = function(pr) {
			this.array = pr;
		};
		CVariant.prototype.setVStream = function(pr) {
			this.vStream = pr;
		};
		CVariant.prototype.setText = function (val) {
			this.setType(c_oVariantTypes.vtLpwstr);
			this.setStrContent(val);
		};
		CVariant.prototype.setNumber = function (val) {
			this.setType(c_oVariantTypes.vtI4);
			this.setIContent(val);
		};
		CVariant.prototype.setDate = function (val) {
			this.setDate(c_oVariantTypes.vtFiletime);
			this.setStrContent(val.toISOString().slice(0, 19) + 'Z');
		};
		CVariant.prototype.setBool = function (val) {
			this.setType(c_oVariantTypes.vtBool);
			this.setBContent(val);
		};
		CVariant.prototype.typeStrToEnum = function (name) {
			switch (name) {
				case "vector": {
					return c_oVariantTypes.vtVector;
					break;
				}
				case "array": {
					return c_oVariantTypes.vtArray;
					break;
				}
				case "blob": {
					return c_oVariantTypes.vtBlob;
					break;
				}
				case "oblob": {
					return c_oVariantTypes.vtOBlob;
					break;
				}
				case "empty": {
					return c_oVariantTypes.vtEmpty;
					break;
				}
				case "null": {
					return c_oVariantTypes.vtNull;
					break;
				}
				case "i1": {
					return c_oVariantTypes.vtI1;
					break;
				}
				case "i2": {
					return c_oVariantTypes.vtI2;
					break;
				}
				case "i4": {
					return c_oVariantTypes.vtI4;
					break;
				}
				case "i8": {
					return c_oVariantTypes.vtI8;
					break;
				}
				case "int": {
					return c_oVariantTypes.vtInt;
					break;
				}
				case "ui1": {
					return c_oVariantTypes.vtUi1;
					break;
				}
				case "ui2": {
					return c_oVariantTypes.vtUi2;
					break;
				}
				case "ui4": {
					return c_oVariantTypes.vtUi4;
					break;
				}
				case "ui8": {
					return c_oVariantTypes.vtUi8;
					break;
				}
				case "uint": {
					return c_oVariantTypes.vtUint;
					break;
				}
				case "r4": {
					return c_oVariantTypes.vtR4;
					break;
				}
				case "r8": {
					return c_oVariantTypes.vtR8;
					break;
				}
				case "decimal": {
					return c_oVariantTypes.vtDecimal;
					break;
				}
				case "lpstr": {
					return c_oVariantTypes.vtLpstr;
					break;
				}
				case "lpwstr": {
					return c_oVariantTypes.vtLpwstr;
					break;
				}
				case "bstr": {
					return c_oVariantTypes.vtBstr;
					break;
				}
				case "date": {
					return c_oVariantTypes.vtDate;
					break;
				}
				case "filetime": {
					return c_oVariantTypes.vtFiletime;
					break;
				}
				case "bool": {
					return c_oVariantTypes.vtBool;
					break;
				}
				case "cy": {
					return c_oVariantTypes.vtCy;
					break;
				}
				case "error": {
					return c_oVariantTypes.vtError;
					break;
				}
				case "stream": {
					return c_oVariantTypes.vtStream;
					break;
				}
				case "ostream": {
					return c_oVariantTypes.vtOStream;
					break;
				}
				case "storage": {
					return c_oVariantTypes.vtStorage;
					break;
				}
				case "ostorage": {
					return c_oVariantTypes.vtOStorage;
					break;
				}
				case "vstream": {
					return c_oVariantTypes.vtVStream;
					break;
				}
				case "clsid": {
					return c_oVariantTypes.vtClsid;
					break;
				}
			}
			return null;
		};
		CVariant.prototype.getStringByType = function (eType) {
			if (c_oVariantTypes.vtEmpty === eType)
				return "empty";
			else if (c_oVariantTypes.vtNull === eType)
				return "null";
			else if (c_oVariantTypes.vtVariant === eType)
				return "variant";
			else if (c_oVariantTypes.vtVector === eType)
				return "vector";
			else if (c_oVariantTypes.vtArray === eType)
				return "array";
			else if (c_oVariantTypes.vtVStream === eType)
				return "vstream";
			else if (c_oVariantTypes.vtBlob === eType)
				return "blob";
			else if (c_oVariantTypes.vtOBlob === eType)
				return "oblob";
			else if (c_oVariantTypes.vtI1 === eType)
				return "i1";
			else if (c_oVariantTypes.vtI2 === eType)
				return "i2";
			else if (c_oVariantTypes.vtI4 === eType)
				return "i4";
			else if (c_oVariantTypes.vtI8 === eType)
				return "i8";
			else if (c_oVariantTypes.vtInt === eType)
				return "int";
			else if (c_oVariantTypes.vtUi1 === eType)
				return "ui1";
			else if (c_oVariantTypes.vtUi2 === eType)
				return "ui2";
			else if (c_oVariantTypes.vtUi4 === eType)
				return "ui4";
			else if (c_oVariantTypes.vtUi8 === eType)
				return "ui8";
			else if (c_oVariantTypes.vtUint === eType)
				return "uint";
			else if (c_oVariantTypes.vtR4 === eType)
				return "r4";
			else if (c_oVariantTypes.vtR8 === eType)
				return "r8";
			else if (c_oVariantTypes.vtDecimal === eType)
				return "decimal";
			else if (c_oVariantTypes.vtLpstr === eType)
				return "lpstr";
			else if (c_oVariantTypes.vtLpwstr === eType)
				return "lpwstr";
			else if (c_oVariantTypes.vtBstr === eType)
				return "bstr";
			else if (c_oVariantTypes.vtDate === eType)
				return "date";
			else if (c_oVariantTypes.vtFiletime === eType)
				return "filetime";
			else if (c_oVariantTypes.vtBool === eType)
				return "bool";
			else if (c_oVariantTypes.vtCy === eType)
				return "cy";
			else if (c_oVariantTypes.vtError === eType)
				return "error";
			else if (c_oVariantTypes.vtStream === eType)
				return "stream";
			else if (c_oVariantTypes.vtOStream === eType)
				return "ostream";
			else if (c_oVariantTypes.vtStorage === eType)
				return "storage";
			else if (c_oVariantTypes.vtOStorage === eType)
				return "ostorage";
			else if (c_oVariantTypes.vtClsid === eType)
				return "clsid";
			return "";
		};
		CVariant.prototype.getVariantType = function () {
			return AscFormat.isRealNumber(this.type) ? this.type : c_oVariantTypes.vtEmpty;
		};
		CVariant.prototype.getValue = function () {
			switch (this.type) {
				case c_oVariantTypes.vtClsid:
				case c_oVariantTypes.vtOStorage:
				case c_oVariantTypes.vtStorage:
				case c_oVariantTypes.vtOStream:
				case c_oVariantTypes.vtStream:
				case c_oVariantTypes.vtError:
				case c_oVariantTypes.vtCy:
				case c_oVariantTypes.vtBstr:
				case c_oVariantTypes.vtDecimal:
				case c_oVariantTypes.vtEmpty:
				case c_oVariantTypes.vtVariant:
				case c_oVariantTypes.vtVector:
				case c_oVariantTypes.vtArray:
				case c_oVariantTypes.vtVStream:
				case c_oVariantTypes.vtBlob:
				case c_oVariantTypes.vtOBlob:
				case c_oVariantTypes.vtNull: {
					return null;
				}
				case c_oVariantTypes.vtI2:
				case c_oVariantTypes.vtI4:
				case c_oVariantTypes.vtI8:
				case c_oVariantTypes.vtInt:
				case c_oVariantTypes.vtI1: {
					return this.iContent;
				}
				case c_oVariantTypes.vtUint:
				case c_oVariantTypes.vtUi8:
				case c_oVariantTypes.vtUi4:
				case c_oVariantTypes.vtUi2:
				case c_oVariantTypes.vtUi1: {
					return this.uContent;
				}
				case c_oVariantTypes.vtR8:
				case c_oVariantTypes.vtR4: {
					return this.dContent;
				}
				case c_oVariantTypes.vtLpwstr:
				case c_oVariantTypes.vtLpstr: {
					return this.strContent;
				}
				case c_oVariantTypes.vtDate:
				case c_oVariantTypes.vtFiletime: {
					return new Date(this.strContent);
				}
				case c_oVariantTypes.vtBool: {
					return this.bContent;
				}
			}
		};
		CVariant.prototype.setValue = function(v) {
			switch (this.type) {
				case c_oVariantTypes.vtClsid:
				case c_oVariantTypes.vtOStorage:
				case c_oVariantTypes.vtStorage:
				case c_oVariantTypes.vtOStream:
				case c_oVariantTypes.vtStream:
				case c_oVariantTypes.vtError:
				case c_oVariantTypes.vtCy:
				case c_oVariantTypes.vtBstr:
				case c_oVariantTypes.vtDecimal:
				case c_oVariantTypes.vtEmpty:
				case c_oVariantTypes.vtVariant:
				case c_oVariantTypes.vtVector:
				case c_oVariantTypes.vtArray:
				case c_oVariantTypes.vtVStream:
				case c_oVariantTypes.vtBlob:
				case c_oVariantTypes.vtOBlob:
				case c_oVariantTypes.vtNull: {
					break;
				}
				case c_oVariantTypes.vtI2:
				case c_oVariantTypes.vtI4:
				case c_oVariantTypes.vtI8:
				case c_oVariantTypes.vtInt:
				case c_oVariantTypes.vtI1: {
					this.setIContent(v);
					break;
				}
				case c_oVariantTypes.vtUint:
				case c_oVariantTypes.vtUi8:
				case c_oVariantTypes.vtUi4:
				case c_oVariantTypes.vtUi2:
				case c_oVariantTypes.vtUi1: {
					this.setUContent(v);
					break;
				}
				case c_oVariantTypes.vtR8:
				case c_oVariantTypes.vtR4: {
					this.setDContent(v);
					break;
				}
				case c_oVariantTypes.vtLpwstr:
				case c_oVariantTypes.vtLpstr: {
					this.setStrContent(v);
					break;
				}
				case c_oVariantTypes.vtDate:
				case c_oVariantTypes.vtFiletime: {
					this.setStrContent(v.toISOString().slice(0, 19) + 'Z');
					break;
				}
				case c_oVariantTypes.vtBool: {
					this.setBContent(v);
					break;
				}
			}
		};
		window['AscCommon'].CVariant = CVariant;

		prot = CVariant.prototype;
		prot["setText"] = prot.setText;
		prot["setNumber"] = prot.setNumber;
		prot["setDate"] = prot.setDate;
		prot["setBool"] = prot.setBool;


		var c_oVariantTypes = {
			vtEmpty: 0,
			vtNull: 1,
			vtVariant: 2,
			vtVector: 3,
			vtArray: 4,
			vtVStream: 5,
			vtBlob: 6,
			vtOBlob: 7,
			vtI1: 8,
			vtI2: 9,
			vtI4: 10,
			vtI8: 11,
			vtInt: 12,
			vtUi1: 13,
			vtUi2: 14,
			vtUi4: 15,
			vtUi8: 16,
			vtUint: 17,
			vtR4: 18,
			vtR8: 19,
			vtDecimal: 20,
			vtLpstr: 21,
			vtLpwstr: 22,
			vtBstr: 23,
			vtDate: 24,
			vtFiletime: 25,
			vtBool: 26,
			vtCy: 27,
			vtError: 28,
			vtStream: 29,
			vtOStream: 30,
			vtStorage: 31,
			vtOStorage: 32,
			vtClsid: 33
		};

		window['AscCommon'].c_oVariantTypes = c_oVariantTypes;
		window['AscCommon']["c_oVariantTypes"] = c_oVariantTypes;
		c_oVariantTypes["vtEmpty"] = c_oVariantTypes.vtEmpty;
		c_oVariantTypes["vtNull"] = c_oVariantTypes.vtNull;
		c_oVariantTypes["vtVariant"] = c_oVariantTypes.vtVariant;
		c_oVariantTypes["vtVector"] = c_oVariantTypes.vtVector;
		c_oVariantTypes["vtArray"] = c_oVariantTypes.vtArray;
		c_oVariantTypes["vtVStream"] = c_oVariantTypes.vtVStream;
		c_oVariantTypes["vtBlob"] = c_oVariantTypes.vtBlob;
		c_oVariantTypes["vtOBlob"] = c_oVariantTypes.vtOBlob;
		c_oVariantTypes["vtI1"] = c_oVariantTypes.vtI1;
		c_oVariantTypes["vtI2"] = c_oVariantTypes.vtI2;
		c_oVariantTypes["vtI4"] = c_oVariantTypes.vtI4;
		c_oVariantTypes["vtI8"] = c_oVariantTypes.vtI8;
		c_oVariantTypes["vtInt"] = c_oVariantTypes.vtInt;
		c_oVariantTypes["vtUi1"] = c_oVariantTypes.vtUi1;
		c_oVariantTypes["vtUi2"] = c_oVariantTypes.vtUi2;
		c_oVariantTypes["vtUi4"] = c_oVariantTypes.vtUi4;
		c_oVariantTypes["vtUi8"] = c_oVariantTypes.vtUi8;
		c_oVariantTypes["vtUint"] = c_oVariantTypes.vtUint;
		c_oVariantTypes["vtR4"] = c_oVariantTypes.vtR4;
		c_oVariantTypes["vtR8"] = c_oVariantTypes.vtR8;
		c_oVariantTypes["vtDecimal"] = c_oVariantTypes.vtDecimal;
		c_oVariantTypes["vtLpstr"] = c_oVariantTypes.vtLpstr;
		c_oVariantTypes["vtLpwstr"] = c_oVariantTypes.vtLpwstr;
		c_oVariantTypes["vtBstr"] = c_oVariantTypes.vtBstr;
		c_oVariantTypes["vtDate"] = c_oVariantTypes.vtDate;
		c_oVariantTypes["vtFiletime"] = c_oVariantTypes.vtFiletime;
		c_oVariantTypes["vtBool"] = c_oVariantTypes.vtBool;
		c_oVariantTypes["vtCy"] = c_oVariantTypes.vtCy;
		c_oVariantTypes["vtError"] = c_oVariantTypes.vtError;
		c_oVariantTypes["vtStream"] = c_oVariantTypes.vtStream;
		c_oVariantTypes["vtOStream"] = c_oVariantTypes.vtOStream;
		c_oVariantTypes["vtStorage"] = c_oVariantTypes.vtStorage;
		c_oVariantTypes["vtOStorage"] = c_oVariantTypes.vtOStorage;
		c_oVariantTypes["vtClsid"] = c_oVariantTypes.vtClsid;

		function CPres() {
			CBaseNoIdObject.call(this);
			this.defaultTextStyle = null;
			this.NotesSz = null;

			this.attrAutoCompressPictures = null;
			this.attrBookmarkIdSeed = null;
			this.attrCompatMode = null;
			this.attrConformance = null;
			this.attrEmbedTrueTypeFonts = null;
			this.attrFirstSlideNum = null;
			this.attrRemovePersonalInfoOnSave = null;
			this.attrRtl = null;
			this.attrSaveSubsetFonts = null;
			this.attrServerZoom = null;
			this.attrShowSpecialPlsOnTitleSld = null;
			this.attrStrictFirstAndLastChars = null;

		}

		InitClass(CPres, CBaseNoIdObject, 0);

		CPres.prototype.readSz = function(s) {
			const oSldSize = new AscCommonSlide.CSlideSize();
			s.Skip2(5); // len + start attributes

			while (true) {
				var _at = s.GetUChar();

				if (_at === g_nodeAttributeEnd)
					break;

				switch (_at) {
					case 0: {
						oSldSize.setCX(s.GetLong());
						break;
					}
					case 1: {
						oSldSize.setCY(s.GetLong());
						break;
					}
					case 2: {
						oSldSize.setType(s.GetUChar());
						break;
					}
					default:
						return;
				}
			}
			return oSldSize;
		};
		CPres.prototype.fromStream = function (s, reader) {
			var _type = s.GetUChar();
			var _len = s.GetULong();
			var _start_pos = s.cur;
			var _end_pos = _len + _start_pos;

			// attributes
			var _sa = s.GetUChar();

			var oPresentattion = reader.presentation;
			while (true) {
				var _at = s.GetUChar();

				if (_at === g_nodeAttributeEnd)
					break;

				switch (_at) {
					case 0: {
						this.attrAutoCompressPictures = s.GetBool();
						break;
					}
					case 1: {
						this.attrBookmarkIdSeed = s.GetLong();
						break;
					}
					case 2: {
						this.attrCompatMode = s.GetBool();
						break;
					}
					case 3: {
						this.attrConformance = s.GetUChar();
						break;
					}
					case 4: {
						this.attrEmbedTrueTypeFonts = s.GetBool();
						break;
					}
					case 5: {
						this.attrFirstSlideNum = s.GetLong();
						break;
					}
					case 6: {
						this.attrRemovePersonalInfoOnSave = s.GetBool();
						break;
					}
					case 7: {
						this.attrRtl = s.GetBool();
						break;
					}
					case 8: {
						this.attrSaveSubsetFonts = s.GetBool();
						break;
					}
					case 9: {
						this.attrServerZoom = s.GetString2();
						break;
					}
					case 10: {
						this.attrShowSpecialPlsOnTitleSld = s.GetBool();
						break;
					}
					case 11: {
						this.attrStrictFirstAndLastChars = s.GetBool();
						break;
					}
					default:
						return;
				}
			}

			while (true) {
				if (s.cur >= _end_pos)
					break;

				_type = s.GetUChar();
				switch (_type) {
					case 0: {
						this.defaultTextStyle = reader.ReadTextListStyle();
						break;
					}
					case 1: {
						s.SkipRecord();
						break;
					}
					case 2: {
						s.SkipRecord();
						break;
					}
					case 3: {
						let oNotesSize = this.readSz(s);
						if (oPresentattion.setNotesSz) {
							oPresentattion.setNotesSz(oNotesSize);
						}
						break;
					}
					case 4: {
						s.SkipRecord();
						break;
					}
					case 5: {
						let oSldSize = this.readSz(s);
						if (oPresentattion.setSldSz) {
							oPresentattion.setSldSz(oSldSize);
						}
						break;
					}
					case 6: {
						var _end_rec2 = s.cur + s.GetULong() + 4;
						while (s.cur < _end_rec2) {
							var _rec = s.GetUChar();

							switch (_rec) {
								case 0: {
									s.Skip2(4); // len
									var lCount = s.GetULong();

									for (var i = 0; i < lCount; i++) {
										s.Skip2(1);

										var _author = new AscCommon.CCommentAuthor();

										var _end_rec3 = s.cur + s.GetLong() + 4;
										s.Skip2(1); // start attributes

										while (true) {
											var _at2 = s.GetUChar();
											if (_at2 === g_nodeAttributeEnd)
												break;

											switch (_at2) {
												case 0:
													_author.Id = s.GetLong();
													break;
												case 1:
													_author.LastId = s.GetLong();
													break;
												case 2:
													var _clr_idx = s.GetLong();
													break;
												case 3:
													_author.Name = s.GetString2();
													break;
												case 4:
													_author.Initials = s.GetString2();
													break;
												default:
													break;
											}
										}

										s.Seek2(_end_rec3);

										oPresentattion.CommentAuthors[_author.Name] = _author;
									}

									break;
								}
								default: {
									s.SkipRecord();
									break;
								}
							}
						}

						s.Seek2(_end_rec2);
						break;
					}
					case 8: {
						oPresentattion.Api.vbaProject = new AscCommon.VbaProject();
						oPresentattion.Api.vbaProject.fromStream(s);
						break;
					}
					case 9: {
						var _length = s.GetULong();
						var _end_rec2 = s.cur + _length;

						oPresentattion.Api.macros.SetData(AscCommon.GetStringUtf8(s, _length));
						s.Seek2(_end_rec2);
						break;
					}
					case 10: {
						reader.ReadComments(oPresentattion.writecomments);
						break;
					}
					default: {
						s.SkipRecord();
						break;
					}
				}
			}
			if (oPresentattion.Load_Comments) {
				oPresentattion.Load_Comments(oPresentattion.CommentAuthors);
			}
			s.Seek2(_end_pos);
		};


		window['AscCommon'].CPres = CPres;

		function CClrMapOvr() {
			CBaseNoIdObject.call(this);
			this.overrideClrMapping = null;
		}
		InitClass(CClrMapOvr, CBaseNoIdObject, 0);


		function IdEntry(name) {
			AscFormat.CBaseNoIdObject.call(this);
			this.name = name;
			this.id = null;
			this.rId = null;
		}
		InitClass(IdEntry, CBaseNoIdObject, undefined);


		function CExtP() {
			this.st = null;
			this.end = null;
		}
		InitClass(CExtP, CBaseNoIdObject, undefined);

		CExtP.prototype.readAttribute = function (nType, pReader) {
			switch (nType) {
				case 0: {
					// id. embed / link
					pReader.stream.Skip2(4);
					break;
				}
				case 1: {
					this.st = pReader.stream.GetDouble();
					break;
				}
				case 2: {
					this.end = pReader.stream.GetDouble();
					break;
				}
			}
		};
		CExtP.prototype.Write_ToBinary = function (w) {
			writeDouble(w, this.st);
			writeDouble(w, this.end);
		};
		CExtP.prototype.Read_FromBinary = function (r) {
			this.st = readDouble(r);
			this.end = readDouble(r);
		};

		function CreateSchemeUnicolorWithMods(id, aMods) {
			let oColor = new CUniColor();
			oColor.color = new CSchemeColor();
			oColor.color.id = id;
			if (aMods) {
				for(let nMod = 0; nMod < aMods.length; ++nMod) {
					let oModObject = aMods[nMod];
					let oMod = new CColorMod(oModObject.name, oModObject.val);
					oColor.addColorMod(oMod);
				}
			}
			return oColor;
		}

// DEFAULT OBJECTS
		function GenerateDefaultTheme(presentation, opt_fontName) {
			return ExecuteNoHistory(function () {
				if (!opt_fontName) {
					opt_fontName = "Arial";
				}
				var theme = new CTheme();
				theme.presentation = presentation;
				theme.setFontScheme(new FontScheme());
				theme.themeElements.fontScheme.setMajorFont(new FontCollection(theme.themeElements.fontScheme));
				theme.themeElements.fontScheme.setMinorFont(new FontCollection(theme.themeElements.fontScheme));
				theme.themeElements.fontScheme.majorFont.setLatin(opt_fontName);
				theme.themeElements.fontScheme.minorFont.setLatin(opt_fontName);


				var scheme = theme.themeElements.clrScheme;

				scheme.colors[8] = CreateUniColorRGB(0, 0, 0);
				scheme.colors[12] = CreateUniColorRGB(255, 255, 255);
				scheme.colors[9] = CreateUniColorRGB(0x1F, 0x49, 0x7D);
				scheme.colors[13] = CreateUniColorRGB(0xEE, 0xEC, 0xE1);
				scheme.colors[0] = CreateUniColorRGB(0x4F, 0x81, 0xBD); //CreateUniColorRGB(0xFF, 0x81, 0xBD);//
				scheme.colors[1] = CreateUniColorRGB(0xC0, 0x50, 0x4D);
				scheme.colors[2] = CreateUniColorRGB(0x9B, 0xBB, 0x59);
				scheme.colors[3] = CreateUniColorRGB(0x80, 0x64, 0xA2);
				scheme.colors[4] = CreateUniColorRGB(0x4B, 0xAC, 0xC6);
				scheme.colors[5] = CreateUniColorRGB(0xF7, 0x96, 0x46);
				scheme.colors[11] = CreateUniColorRGB(0x00, 0x00, 0xFF);
				scheme.colors[10] = CreateUniColorRGB(0x80, 0x00, 0x80);

				// -------------- fill styles -------------------------
				var brush = new CUniFill();
				brush.setFill(new CSolidFill());
				brush.fill.setColor(new CUniColor());
				brush.fill.color.setColor(new CSchemeColor());
				brush.fill.color.color.setId(phClr);
				theme.themeElements.fmtScheme.fillStyleLst.push(brush);

				brush = new CUniFill();

				let oFill = new CGradFill();
				oFill.rotateWithShape = true;
				brush.setFill(oFill);
				let oLin = new GradLin();
				oFill.setLin(oLin);
				oLin.setAngle(16200000);
				oLin.setScale(true);
				let oGs = new CGs();
				oFill.addColor(oGs);
				oGs.setPos(0);

				let oColor = CreateSchemeUnicolorWithMods(14, [{name: "tint", val: 50000}, {name: "satMod", val: 300000}]);
				oGs.setColor(oColor);

				oGs = new CGs();
				oFill.addColor(oGs);
				oGs.setPos(35000);
				oColor = CreateSchemeUnicolorWithMods(14, [{name: "tint", val: 37000}, {name: "satMod", val: 300000}]);
				oGs.setColor(oColor);

				oGs = new CGs();
				oFill.addColor(oGs);
				oGs.setPos(100000);
				oColor = CreateSchemeUnicolorWithMods(14, [{name: "tint", val: 15000}, {name: "satMod", val: 350000}]);
				oGs.setColor(oColor);

				theme.themeElements.fmtScheme.fillStyleLst.push(brush);

				brush = new CUniFill();

				oFill = new CGradFill();
				brush.setFill(oFill);
				oFill.rotateWithShape = true;
				oLin = new GradLin();
				oFill.setLin(oLin);
				oLin.setAngle(16200000);
				oLin.setScale(false);
				oGs = new CGs();
				oFill.addColor(oGs);
				oGs.setPos(0);
				oColor = CreateSchemeUnicolorWithMods(14, [{name: "shade", val: 51000}, {name: "satMod", val: 130000}]);
				oGs.setColor(oColor);

				oGs = new CGs();
				oFill.addColor(oGs);
				oGs.setPos(80000);
				oColor = CreateSchemeUnicolorWithMods(14, [{name: "shade", val: 93000}, {name: "satMod", val: 130000}]);
				oGs.setColor(oColor);

				oGs = new CGs();
				oFill.addColor(oGs);
				oGs.setPos(100000);
				oColor = CreateSchemeUnicolorWithMods(14, [{name: "shade", val: 94000}, {name: "satMod", val: 135000}]);
				oGs.setColor(oColor);
				
				theme.themeElements.fmtScheme.fillStyleLst.push(brush);
				// ----------------------------------------------------

				// -------------- back styles -------------------------
				brush = new CUniFill();
				brush.setFill(new CSolidFill());
				brush.fill.setColor(new CUniColor());
				brush.fill.color.setColor(new CSchemeColor());
				brush.fill.color.color.setId(phClr);
				theme.themeElements.fmtScheme.bgFillStyleLst.push(brush);

				brush = AscFormat.CreateUniFillByUniColor(AscFormat.CreateUniColorRGB(0, 0, 0));
				theme.themeElements.fmtScheme.bgFillStyleLst.push(brush);

				brush = AscFormat.CreateUniFillByUniColor(AscFormat.CreateUniColorRGB(0, 0, 0));
				theme.themeElements.fmtScheme.bgFillStyleLst.push(brush);
				// ----------------------------------------------------

				var pen = new CLn();
				pen.setW(9525);
				pen.setFill(new CUniFill());
				pen.Fill.setFill(new CSolidFill());
				pen.Fill.fill.setColor(new CUniColor());
				pen.Fill.fill.color.setColor(new CSchemeColor());
				pen.Fill.fill.color.color.setId(phClr);
				pen.Fill.fill.color.setMods(new CColorModifiers());

				pen.Fill.fill.color.Mods.addMod("shade", 95000);
				pen.Fill.fill.color.Mods.addMod("satMod", 105000);
				theme.themeElements.fmtScheme.lnStyleLst.push(pen);

				pen = new CLn();
				pen.setW(25400);
				pen.setFill(new CUniFill());
				pen.Fill.setFill(new CSolidFill());

				pen.Fill.fill.setColor(new CUniColor());
				pen.Fill.fill.color.setColor(new CSchemeColor());
				pen.Fill.fill.color.color.setId(phClr);
				theme.themeElements.fmtScheme.lnStyleLst.push(pen);

				pen = new CLn();
				pen.setW(38100);
				pen.setFill(new CUniFill());
				pen.Fill.setFill(new CSolidFill());
				pen.Fill.fill.setColor(new CUniColor());
				pen.Fill.fill.color.setColor(new CSchemeColor());
				pen.Fill.fill.color.color.setId(phClr);
				theme.themeElements.fmtScheme.lnStyleLst.push(pen);
				theme.extraClrSchemeLst = [];
				return theme;
			}, this, []);
		}

		function GetDefaultTheme() {
			if(!AscFormat.DEFAULT_THEME) {
				AscFormat.DEFAULT_THEME = GenerateDefaultTheme(null);
			}
			return AscFormat.DEFAULT_THEME;
		}

		function GenerateDefaultMasterSlide(theme) {
			var master = new MasterSlide(theme.presentation, theme);
			master.Theme = theme;

			master.sldLayoutLst[0] = GenerateDefaultSlideLayout(master);
			master.setPreserve(true);
			return master;
		}

		function GenerateDefaultSlideLayout(master) {
			var layout = new SlideLayout();
			layout.Theme = master.Theme;
			layout.Master = master;
			layout.setPreserve(true);
			return layout;
		}

		function GenerateDefaultSlide(layout) {
			var slide = new Slide(layout.Master.presentation, layout, 0);
			slide.Master = layout.Master;
			slide.Theme = layout.Master.Theme;
			slide.setNotes(AscCommonSlide.CreateNotes());
			slide.notes.setNotesMaster(layout.Master.presentation.notesMasters[0]);
			slide.notes.setSlide(slide);
			return slide;
		}

		function CreateDefaultTextRectStyle() {
			var style = new CShapeStyle();
			var lnRef = new StyleRef();
			lnRef.setIdx(0);
			var unicolor = new CUniColor();

			unicolor.setColor(new CSchemeColor());
			unicolor.color.setId(g_clr_accent1);
			unicolor.setMods(new CColorModifiers());
			unicolor.Mods.addMod("shade", 50000);
			lnRef.setColor(unicolor);
			style.setLnRef(lnRef);


			var fillRef = new StyleRef();
			fillRef.setIdx(0);
			unicolor = new CUniColor();
			unicolor.setColor(new CSchemeColor());
			unicolor.color.setId(g_clr_accent1);
			fillRef.setColor(unicolor);
			style.setFillRef(fillRef);


			var effectRef = new StyleRef();
			effectRef.setIdx(0);
			unicolor = new CUniColor();
			unicolor.setColor(new CSchemeColor());
			unicolor.color.setId(g_clr_accent1);
			effectRef.setColor(unicolor);
			style.setEffectRef(effectRef);

			var fontRef = new FontRef();
			fontRef.setIdx(AscFormat.fntStyleInd_minor);
			unicolor = new CUniColor();
			unicolor.setColor(new CSchemeColor());
			unicolor.color.setId(8);
			fontRef.setColor(unicolor);
			style.setFontRef(fontRef);

			return style;
		}

		function GenerateDefaultColorMap() {
			return AscFormat.ExecuteNoHistory(function () {
				var clrMap = new ClrMap();
				clrMap.color_map[0] = 0;
				clrMap.color_map[1] = 1;
				clrMap.color_map[2] = 2;
				clrMap.color_map[3] = 3;
				clrMap.color_map[4] = 4;
				clrMap.color_map[5] = 5;
				clrMap.color_map[10] = 10;
				clrMap.color_map[11] = 11;
				clrMap.color_map[6] = 12;
				clrMap.color_map[7] = 13;
				clrMap.color_map[15] = 8;
				clrMap.color_map[16] = 9;
				return clrMap;
			}, [], null);
		}

		function GetDefaultColorMap() {
			if(!AscFormat.DEFAULT_COLOR_MAP) {
				AscFormat.DEFAULT_COLOR_MAP = GenerateDefaultColorMap();
			}
			return AscFormat.DEFAULT_COLOR_MAP;
		}

		function CreateAscFill(unifill) {
			if (null == unifill || null == unifill.fill)
				return new asc_CShapeFill();

			var ret = new asc_CShapeFill();

			var _fill = unifill.fill;
			switch (_fill.type) {
				case c_oAscFill.FILL_TYPE_SOLID: {
					ret.type = c_oAscFill.FILL_TYPE_SOLID;
					ret.fill = new Asc.asc_CFillSolid();
					ret.fill.color = CreateAscColor(_fill.color);
					break;
				}
				case c_oAscFill.FILL_TYPE_PATT: {
					ret.type = c_oAscFill.FILL_TYPE_PATT;
					ret.fill = new Asc.asc_CFillHatch();
					ret.fill.PatternType = _fill.ftype;
					ret.fill.fgClr = CreateAscColor(_fill.fgClr);
					ret.fill.bgClr = CreateAscColor(_fill.bgClr);
					break;
				}
				case c_oAscFill.FILL_TYPE_GRAD: {
					ret.type = c_oAscFill.FILL_TYPE_GRAD;
					ret.fill = new Asc.asc_CFillGrad();
					var bCheckTransparent = true, nLastTransparent = null, nLastTempTransparent, j, aMods;
					for (var i = 0; i < _fill.colors.length; i++) {
						if (0 == i) {
							ret.fill.Colors = [];
							ret.fill.Positions = [];
						}
						if (bCheckTransparent) {
							if (_fill.colors[i].color.Mods) {
								aMods = _fill.colors[i].color.Mods.Mods;
								nLastTempTransparent = null;
								for (j = 0; j < aMods.length; ++j) {
									if (aMods[j].name === "alpha") {
										if (nLastTempTransparent === null) {
											nLastTempTransparent = aMods[j].val;
											if (nLastTransparent === null) {
												nLastTransparent = nLastTempTransparent;
											} else {
												if (nLastTransparent !== nLastTempTransparent) {
													bCheckTransparent = false;
													break;
												}
											}
										} else {
											bCheckTransparent = false;
											break;
										}
									}
								}
							} else {
								bCheckTransparent = false;
							}
						}
						ret.fill.Colors.push(CreateAscColor(_fill.colors[i].color));
						ret.fill.Positions.push(_fill.colors[i].pos);
					}
					if (bCheckTransparent && nLastTransparent !== null) {
						ret.transparent = (nLastTransparent / 100000) * 255;
					}

					if (_fill.lin) {
						ret.fill.GradType = c_oAscFillGradType.GRAD_LINEAR;
						ret.fill.LinearAngle = _fill.lin.angle;
						ret.fill.LinearScale = _fill.lin.scale;
					} else if (_fill.path) {
						ret.fill.GradType = c_oAscFillGradType.GRAD_PATH;
						ret.fill.PathType = 0;
					} else {
						ret.fill.GradType = c_oAscFillGradType.GRAD_LINEAR;
						ret.fill.LinearAngle = 0;
						ret.fill.LinearScale = false;
					}

					break;
				}
				case c_oAscFill.FILL_TYPE_BLIP: {
					ret.type = c_oAscFill.FILL_TYPE_BLIP;
					ret.fill = new Asc.asc_CFillBlip();

					ret.fill.url = _fill.RasterImageId;
					ret.fill.type = (_fill.tile == null) ? c_oAscFillBlipType.STRETCH : c_oAscFillBlipType.TILE;
					break;
				}
				case c_oAscFill.FILL_TYPE_NOFILL: {
					ret.type = c_oAscFill.FILL_TYPE_NOFILL;
					break;
				}
				default:
					break;
			}

			if (isRealNumber(unifill.transparent)) {
				ret.transparent = unifill.transparent;
			}
			return ret;
		}

		function CorrectUniFill(asc_fill, unifill, editorId) {

			if (asc_fill instanceof CUniFill) {
				return asc_fill;
			}
			if (null == asc_fill)
				return unifill;

			var ret = unifill;
			if (null == ret)
				ret = new CUniFill();

			var _fill = asc_fill.fill;
			var _type = asc_fill.type;

			if (null != _type) {
				switch (_type) {
					case c_oAscFill.FILL_TYPE_NOFILL: {
						ret.fill = new CNoFill();
						break;
					}
					case c_oAscFill.FILL_TYPE_GRP: {
						ret.fill = new CGrpFill();
						break;
					}
					case c_oAscFill.FILL_TYPE_BLIP: {

						var _url = _fill.url;
						var _tx_id = _fill.texture_id;
						if (null != _tx_id && (0 <= _tx_id) && (_tx_id < AscCommon.g_oUserTexturePresets.length)) {
							_url = AscCommon.g_oUserTexturePresets[_tx_id];
						}


						if (ret.fill == null) {
							ret.fill = new CBlipFill();
						}

						if (ret.fill.type !== c_oAscFill.FILL_TYPE_BLIP) {
							if (!(typeof (_url) === "string" && _url.length > 0) || !isRealNumber(_fill.type)) {
								break;
							}
							ret.fill = new CBlipFill();
						}

						if (_url != null && _url !== undefined && _url != "")
							ret.fill.setRasterImageId(_url);

						if (ret.fill.RasterImageId == null)
							ret.fill.RasterImageId = "";

						var tile = _fill.type;
						if (tile === c_oAscFillBlipType.STRETCH) {
							ret.fill.tile = null;
							ret.fill.srcRect = null;
							ret.fill.stretch = true;
						} else if (tile === c_oAscFillBlipType.TILE) {
							ret.fill.tile = new CBlipFillTile();
							ret.fill.stretch = false;
							ret.fill.srcRect = null;
						}
						break;
					}
					case c_oAscFill.FILL_TYPE_PATT: {
						if (ret.fill == null) {
							ret.fill = new CPattFill();
						}

						if (ret.fill.type != c_oAscFill.FILL_TYPE_PATT) {
							if (undefined != _fill.PatternType && undefined != _fill.fgClr && undefined != _fill.bgClr) {
								ret.fill = new CPattFill();
							} else {
								break;
							}
						}

						if (undefined != _fill.PatternType) {
							ret.fill.ftype = _fill.PatternType;
						}
						if (undefined != _fill.fgClr) {
							ret.fill.fgClr = CorrectUniColor(_fill.fgClr, ret.fill.fgClr, editorId);
						}
						if (!ret.fill.fgClr) {
							ret.fill.fgClr = CreateUniColorRGB(0, 0, 0);
						}
						if (undefined != _fill.bgClr) {
							ret.fill.bgClr = CorrectUniColor(_fill.bgClr, ret.fill.bgClr, editorId);
						}
						if (!ret.fill.bgClr) {
							ret.fill.bgClr = CreateUniColorRGB(0, 0, 0);
						}

						break;
					}
					case c_oAscFill.FILL_TYPE_GRAD: {
						if (ret.fill == null) {
							ret.fill = new CGradFill();
						}

						var _colors = _fill.Colors;
						var _positions = _fill.Positions;

						if (ret.fill.type != c_oAscFill.FILL_TYPE_GRAD) {
							if (undefined != _colors && undefined != _positions) {
								ret.fill = new CGradFill();
							} else {
								break;
							}
						}

						if (undefined != _colors && undefined != _positions) {
							if (_colors.length === _positions.length) {
								if (ret.fill.colors.length === _colors.length) {
									for (var i = 0; i < _colors.length; i++) {
										var _gs = ret.fill.colors[i] ? ret.fill.colors[i] : new CGs();
										_gs.color = CorrectUniColor(_colors[i], _gs.color, editorId);
										_gs.pos = _positions[i];
										ret.fill.colors[i] = _gs;
									}
								} else {
									ret.fill.colors.length = 0;
									for (var i = 0; i < _colors.length; i++) {
										var _gs = new CGs();
										_gs.color = CorrectUniColor(_colors[i], _gs.color, editorId);
										_gs.pos = _positions[i];

										ret.fill.colors.push(_gs);
									}
								}
							}
						} else if (undefined != _colors) {
							if (_colors.length === ret.fill.colors.length) {
								for (var i = 0; i < _colors.length; i++) {
									ret.fill.colors[i].color = CorrectUniColor(_colors[i], ret.fill.colors[i].color, editorId);
								}
							}
						} else if (undefined != _positions) {
							if (_positions.length <= ret.fill.colors.length) {
								if (_positions.length < ret.fill.colors.length) {
									ret.fill.colors.splice(_positions.length, ret.fill.colors.length - _positions.length);
								}
								for (var i = 0; i < _positions.length; i++) {
									ret.fill.colors[i].pos = _positions[i];
								}
							}
						}

						var _grad_type = _fill.GradType;

						if (c_oAscFillGradType.GRAD_LINEAR === _grad_type) {
							var _angle = _fill.LinearAngle;
							var _scale = _fill.LinearScale;

							if (!ret.fill.lin) {
								ret.fill.lin = new GradLin();
								ret.fill.lin.angle = 0;
								ret.fill.lin.scale = false;
							}

							if (undefined != _angle)
								ret.fill.lin.angle = _angle;
							if (undefined != _scale)
								ret.fill.lin.scale = _scale;
							ret.fill.path = null;
						} else if (c_oAscFillGradType.GRAD_PATH === _grad_type) {
							ret.fill.lin = null;
							ret.fill.path = new GradPath();
						}
						break;
					}
					default: {
						if (ret.fill == null || ret.fill.type !== c_oAscFill.FILL_TYPE_SOLID) {
							ret.fill = new CSolidFill();
						}
						ret.fill.color = CorrectUniColor(_fill.color, ret.fill.color, editorId);
					}
				}
			}

			var _alpha = asc_fill.transparent;
			if (null != _alpha) {
				ret.transparent = _alpha;


			}

			if (ret.transparent != null) {
				if (ret.fill && ret.fill.type === c_oAscFill.FILL_TYPE_BLIP) {
					ret.fill.setTransparent(ret.transparent);
				}
			}
			return ret;
		}

// эта функция ДОЛЖНА минимизироваться
		function CreateAscStroke(ln, _canChangeArrows) {
			if (null == ln || null == ln.Fill || ln.Fill.fill == null)
				return new Asc.asc_CStroke();

			var ret = new Asc.asc_CStroke();

			var _fill = ln.Fill.fill;
			if (_fill != null) {
				switch (_fill.type) {
					case c_oAscFill.FILL_TYPE_BLIP: {
						break;
					}
					case c_oAscFill.FILL_TYPE_SOLID: {
						ret.color = CreateAscColor(_fill.color);
						ret.type = c_oAscStrokeType.STROKE_COLOR;
						break;
					}
					case c_oAscFill.FILL_TYPE_GRAD: {
						var _c = _fill.colors;
						if (_c != 0) {
							ret.color = CreateAscColor(_fill.colors[0].color);
							ret.type = c_oAscStrokeType.STROKE_COLOR;
						}

						break;
					}
					case c_oAscFill.FILL_TYPE_PATT: {
						ret.color = CreateAscColor(_fill.fgClr);
						ret.type = c_oAscStrokeType.STROKE_COLOR;
						break;
					}
					case c_oAscFill.FILL_TYPE_NOFILL: {
						ret.color = null;
						ret.type = c_oAscStrokeType.STROKE_NONE;
						break;
					}
					default: {
						break;
					}
				}

			}


			ret.width = (ln.w == null) ? 12700 : (ln.w >> 0);
			ret.width /= 36000.0;

			if (ln.cap != null)
				ret.asc_putLinecap(ln.cap);

			if (ln.Join != null)
				ret.asc_putLinejoin(ln.Join.type);

			if (ln.headEnd != null) {
				ret.asc_putLinebeginstyle((ln.headEnd.type == null) ? LineEndType.None : ln.headEnd.type);

				var _len = (null == ln.headEnd.len) ? 1 : (2 - ln.headEnd.len);
				var _w = (null == ln.headEnd.w) ? 1 : (2 - ln.headEnd.w);

				ret.asc_putLinebeginsize(_w * 3 + _len);
			} else {
				ret.asc_putLinebeginstyle(LineEndType.None);
			}

			if (ln.tailEnd != null) {
				ret.asc_putLineendstyle((ln.tailEnd.type == null) ? LineEndType.None : ln.tailEnd.type);

				var _len = (null == ln.tailEnd.len) ? 1 : (2 - ln.tailEnd.len);
				var _w = (null == ln.tailEnd.w) ? 1 : (2 - ln.tailEnd.w);

				ret.asc_putLineendsize(_w * 3 + _len);
			} else {
				ret.asc_putLineendstyle(LineEndType.None);
			}
			if (AscFormat.isRealNumber(ln.prstDash)) {
				ret.prstDash = ln.prstDash;
			} else if (ln.prstDash === null) {
				ret.prstDash = Asc.c_oDashType.solid;
			}
			if (true === _canChangeArrows)
				ret.canChangeArrows = true;

			if (isRealNumber(ln.Fill.transparent)) {
				ret.transparent = ln.Fill.transparent;
			}
			return ret;
		}

		function CorrectUniStroke(asc_stroke, unistroke, flag) {
			if (null == asc_stroke)
				return unistroke;

			var ret = unistroke;
			if (null == ret)
				ret = new CLn();

			var _type = asc_stroke.type;
			var _w = asc_stroke.width;

			if (_w !== null && _w !== undefined)
				ret.w = _w * 36000.0;

			var _color = asc_stroke.color;
			if (_type === c_oAscStrokeType.STROKE_NONE) {
				ret.Fill = new CUniFill();
				ret.Fill.fill = new CNoFill();
			} else if (_type != null) {
				if (null !== _color && undefined !== _color) {
					ret.Fill = new CUniFill();
					ret.Fill.type = c_oAscFill.FILL_TYPE_SOLID;
					ret.Fill.fill = new CSolidFill();
					ret.Fill.fill.color = CorrectUniColor(_color, ret.Fill.fill.color, flag);
				}
			}

			var _alpha = asc_stroke.transparent;
			if (null != _alpha) {
				if(!ret.Fill) {
					ret.Fill = AscFormat.CreateSolidFillRGB(0, 0, 0);
					if(!ret.w) {
						ret.w = 12700;
					}
				}
				ret.Fill.transparent = _alpha;
			}

			var _join = asc_stroke.LineJoin;
			if (null != _join) {
				ret.Join = new LineJoin();
				ret.Join.type = _join;
			}

			var _cap = asc_stroke.LineCap;
			if (null != _cap) {
				ret.cap = _cap;
			}

			var _begin_style = asc_stroke.LineBeginStyle;
			if (null != _begin_style) {
				if (ret.headEnd == null)
					ret.headEnd = new EndArrow();

				ret.headEnd.type = _begin_style;
			}

			var _end_style = asc_stroke.LineEndStyle;
			if (null != _end_style) {
				if (ret.tailEnd == null)
					ret.tailEnd = new EndArrow();

				ret.tailEnd.type = _end_style;
			}

			var _begin_size = asc_stroke.LineBeginSize;
			if (null != _begin_size) {
				if (ret.headEnd == null)
					ret.headEnd = new EndArrow();

				ret.headEnd.w = 2 - ((_begin_size / 3) >> 0);
				ret.headEnd.len = 2 - (_begin_size % 3);
			}

			var _end_size = asc_stroke.LineEndSize;
			if (null != _end_size) {
				if (ret.tailEnd == null)
					ret.tailEnd = new EndArrow();

				ret.tailEnd.w = 2 - ((_end_size / 3) >> 0);
				ret.tailEnd.len = 2 - (_end_size % 3);
			}
			if (AscFormat.isRealNumber(asc_stroke.prstDash)) {
				ret.prstDash = asc_stroke.prstDash;
			}
			return ret;
		}

// эта функция ДОЛЖНА минимизироваться
		function CreateAscShapeProp(shape) {
			if (null == shape)
				return new asc_CShapeProperty();

			var ret = new asc_CShapeProperty();
			ret.fill = CreateAscFill(shape.brush);
			ret.stroke = CreateAscStroke(shape.pen);
			ret.lockAspect = shape.getNoChangeAspect();
			var paddings = null;
			if (shape.textBoxContent) {
				var body_pr = shape.bodyPr;
				paddings = new Asc.asc_CPaddings();
				if (typeof body_pr.lIns === "number")
					paddings.Left = body_pr.lIns;
				else
					paddings.Left = 2.54;

				if (typeof body_pr.tIns === "number")
					paddings.Top = body_pr.tIns;
				else
					paddings.Top = 1.27;

				if (typeof body_pr.rIns === "number")
					paddings.Right = body_pr.rIns;
				else
					paddings.Right = 2.54;

				if (typeof body_pr.bIns === "number")
					paddings.Bottom = body_pr.bIns;
				else
					paddings.Bottom = 1.27;
			}
			return ret;
		}

		function CreateAscShapePropFromProp(shapeProp) {
			var obj = new asc_CShapeProperty();
			if (!isRealObject(shapeProp))
				return obj;
			if (isRealBool(shapeProp.locked)) {
				obj.Locked = shapeProp.locked;
			}
			obj.lockAspect = shapeProp.lockAspect;
			if (typeof shapeProp.type === "string")
				obj.type = shapeProp.type;
			if (isRealObject(shapeProp.fill))
				obj.fill = CreateAscFill(shapeProp.fill);
			if (isRealObject(shapeProp.stroke))
				obj.stroke = CreateAscStroke(shapeProp.stroke, shapeProp.canChangeArrows);
			if (isRealObject(shapeProp.paddings))
				obj.paddings = shapeProp.paddings;
			if (shapeProp.canFill === true || shapeProp.canFill === false) {
				obj.canFill = shapeProp.canFill;
			}
			obj.bFromChart = shapeProp.bFromChart;
			obj.bFromSmartArt = shapeProp.bFromSmartArt;
			obj.bFromSmartArtInternal = shapeProp.bFromSmartArtInternal;
			obj.bFromGroup = shapeProp.bFromGroup;
			obj.bFromImage = shapeProp.bFromImage;
			obj.w = shapeProp.w;
			obj.h = shapeProp.h;
			obj.rot = shapeProp.rot;
			obj.flipH = shapeProp.flipH;
			obj.flipV = shapeProp.flipV;
			obj.vert = shapeProp.vert;
			obj.verticalTextAlign = shapeProp.verticalTextAlign;
			if (shapeProp.textArtProperties) {
				obj.textArtProperties = CreateAscTextArtProps(shapeProp.textArtProperties);
			}
			obj.title = shapeProp.title;
			obj.description = shapeProp.description;
			obj.name = shapeProp.name;
			obj.columnNumber = shapeProp.columnNumber;
			obj.columnSpace = shapeProp.columnSpace;
			obj.textFitType = shapeProp.textFitType;
			obj.vertOverflowType = shapeProp.vertOverflowType;
			obj.shadow = shapeProp.shadow;
			if (shapeProp.signatureId) {
				obj.signatureId = shapeProp.signatureId;
			}

			if (shapeProp.signatureId) {
				obj.signatureId = shapeProp.signatureId;
			}
			obj.Position = new Asc.CPosition({X: shapeProp.x, Y: shapeProp.y});
			obj.isMotionPath = !!shapeProp.isMotionPath;
			return obj;
		}

		function CorrectShapeProp(asc_shape_prop, shape) {
			if (null == shape || null == asc_shape_prop)
				return;

			shape.spPr.Fill = CorrectUniFill(asc_shape_prop.asc_getFill(), shape.spPr.Fill);
			shape.spPr.ln = CorrectUniFill(asc_shape_prop.asc_getStroke(), shape.spPr.ln);
		}


		function CreateAscTextArtProps(oTextArtProps) {
			if (!oTextArtProps) {
				return undefined;
			}
			var oRet = new Asc.asc_TextArtProperties();
			if (oTextArtProps.Fill) {
				oRet.asc_putFill(CreateAscFill(oTextArtProps.Fill));
			}
			if (oTextArtProps.Line) {
				oRet.asc_putLine(CreateAscStroke(oTextArtProps.Line, false));
			}
			oRet.asc_putForm(oTextArtProps.Form);
			return oRet;
		}

		function CreateUnifillFromAscColor(asc_color, editorId) {
			var Unifill = new CUniFill();
			Unifill.fill = new CSolidFill();
			Unifill.fill.color = CorrectUniColor(asc_color, Unifill.fill.color, editorId);
			return Unifill;
		}

		function CorrectUniColor(asc_color, unicolor, flag) {
			if (null == asc_color)
				return unicolor;

			var ret = unicolor;
			if (null == ret)
				ret = new CUniColor();

			var _type = asc_color.asc_getType();
			switch (_type) {
				case c_oAscColor.COLOR_TYPE_PRST: {
					if (ret.color == null || ret.color.type !== c_oAscColor.COLOR_TYPE_PRST) {
						ret.color = new CPrstColor();
					}
					ret.color.id = asc_color.value;

					if (ret.Mods.Mods.length != 0)
						ret.Mods.Mods.splice(0, ret.Mods.Mods.length);
					break;
				}
				case c_oAscColor.COLOR_TYPE_SCHEME: {
					if (ret.color == null || ret.color.type !== c_oAscColor.COLOR_TYPE_SCHEME) {
						ret.color = new CSchemeColor();
					}

					// тут выставляется ТОЛЬКО из меню. поэтому:
					var _index = parseInt(asc_color.value);
					if (isNaN(_index))
						break;
					var _id = (_index / 6) >> 0;
					var _pos = _index - _id * 6;

					var array_colors_types = [6, 15, 7, 16, 0, 1, 2, 3, 4, 5];
					ret.color.id = array_colors_types[_id];
					if (!ret.Mods) {
						ret.setMods(new CColorModifiers());
					}

					if (ret.Mods.Mods.length != 0)
						ret.Mods.Mods.splice(0, ret.Mods.Mods.length);

					var __mods = null;

					var _flag;
					if (editor && editor.WordControl && editor.WordControl.m_oDrawingDocument && editor.WordControl.m_oDrawingDocument.GuiControlColorsMap) {
						var _map = editor.WordControl.m_oDrawingDocument.GuiControlColorsMap;
						_flag = isRealNumber(flag) ? flag : 1;
						__mods = AscCommon.GetDefaultMods(_map[_id].r, _map[_id].g, _map[_id].b, _pos, _flag);
					} else {
						var _editor = window["Asc"] && window["Asc"]["editor"];
						if (_editor && _editor.wbModel) {
							var _theme = _editor.wbModel.theme;
							var _clrMap = _editor.wbModel.clrSchemeMap;

							if (_theme && _clrMap) {
								var _schemeClr = new CSchemeColor();
								_schemeClr.id = array_colors_types[_id];

								var _rgba = {R: 0, G: 0, B: 0, A: 255};
								_schemeClr.Calculate(_theme, _clrMap.color_map, _rgba);
								_flag = isRealNumber(flag) ? flag : 0;
								__mods = AscCommon.GetDefaultMods(_schemeClr.RGBA.R, _schemeClr.RGBA.G, _schemeClr.RGBA.B, _pos, _flag);
							}
						}
					}

					if (null != __mods) {
						ret.Mods.Mods = __mods;
					}

					break;
				}
				default: {
					if (ret.color == null || ret.color.type !== c_oAscColor.COLOR_TYPE_SRGB) {
						ret.color = new CRGBColor();
					}
					ret.color.RGBA.R = asc_color.r;
					ret.color.RGBA.G = asc_color.g;
					ret.color.RGBA.B = asc_color.b;
					ret.color.RGBA.A = asc_color.a;

					if (ret.Mods && ret.Mods.Mods.length !== 0)
						ret.Mods.Mods.splice(0, ret.Mods.Mods.length);
				}
			}
			return ret;
		}

		function deleteDrawingBase(aObjects, graphicId) {
			var position = null;
			for (var i = 0; i < aObjects.length; i++) {
				if (aObjects[i].graphicObject.Get_Id() == graphicId) {
					aObjects.splice(i, 1);
					position = i;
					break;
				}
			}
			return position;
		}



		/**
		 * Common Functions For Builder
		 * @return {CShape}
		 */
		function builder_CreateShape(sType, nWidth, nHeight, oFill, oStroke, oParent, oTheme, oDrawingDocument, bWord, worksheet) {
			var oShapeTrack = new AscFormat.NewShapeTrack(sType, 0, 0, oTheme, null, null, null, 0);
			oShapeTrack.track({}, nWidth, nHeight, true);
			var oShape = oShapeTrack.getShape(bWord === true, oDrawingDocument, null);
			oShape.setParent(oParent);
			if (worksheet) {
				oShape.setWorksheet(worksheet);
			}
			if (bWord) {
				oShape.createTextBoxContent();
			} else {
				oShape.createTextBody();
			}
			oShape.spPr.setFill(oFill);
			oShape.spPr.setLn(oStroke);
			return oShape;
		}

		function ChartBuilderTypeToInternal(sType) {
			switch (sType) {
				case "bar" : {
					return Asc.c_oAscChartTypeSettings.barNormal;
				}
				case "barStacked": {
					return Asc.c_oAscChartTypeSettings.barStacked;
				}
				case "barStackedPercent": {
					return Asc.c_oAscChartTypeSettings.barStackedPer;
				}
				case "bar3D": {
					return Asc.c_oAscChartTypeSettings.barNormal3d;
				}
				case "barStacked3D": {
					return Asc.c_oAscChartTypeSettings.barStacked3d;
				}
				case "barStackedPercent3D": {
					return Asc.c_oAscChartTypeSettings.barStackedPer3d;
				}
				case "barStackedPercent3DPerspective": {
					return Asc.c_oAscChartTypeSettings.barNormal3dPerspective;
				}
				case "horizontalBar": {
					return Asc.c_oAscChartTypeSettings.hBarNormal;
				}
				case "horizontalBarStacked": {
					return Asc.c_oAscChartTypeSettings.hBarStacked;
				}
				case "horizontalBarStackedPercent": {
					return Asc.c_oAscChartTypeSettings.hBarStackedPer;
				}
				case "horizontalBar3D": {
					return Asc.c_oAscChartTypeSettings.hBarNormal3d;
				}
				case "horizontalBarStacked3D": {
					return Asc.c_oAscChartTypeSettings.hBarStacked3d;
				}
				case "horizontalBarStackedPercent3D": {
					return Asc.c_oAscChartTypeSettings.hBarStackedPer3d;
				}
				case "lineNormal": {
					return Asc.c_oAscChartTypeSettings.lineNormal;
				}
				case "lineStacked": {
					return Asc.c_oAscChartTypeSettings.lineStacked;
				}
				case "lineStackedPercent": {
					return Asc.c_oAscChartTypeSettings.lineStackedPer;
				}
				case "line3D": {
					return Asc.c_oAscChartTypeSettings.line3d;
				}
				case "pie": {
					return Asc.c_oAscChartTypeSettings.pie;
				}
				case "pie3D": {
					return Asc.c_oAscChartTypeSettings.pie3d;
				}
				case "doughnut": {
					return Asc.c_oAscChartTypeSettings.doughnut;
				}
				case "scatter": {
					return Asc.c_oAscChartTypeSettings.scatter;
				}
				case "stock": {
					return Asc.c_oAscChartTypeSettings.stock;
				}
				case "area": {
					return Asc.c_oAscChartTypeSettings.areaNormal;
				}
				case "areaStacked": {
					return Asc.c_oAscChartTypeSettings.areaStacked;
				}
				case "areaStackedPercent": {
					return Asc.c_oAscChartTypeSettings.areaStackedPer;
				}
				case "comboBarLine": {
					return Asc.c_oAscChartTypeSettings.comboBarLine;
				}
				case "comboBarLineSecondary": {
					return Asc.c_oAscChartTypeSettings.comboBarLineSecondary;
				}
				case "comboCustom": {
					return Asc.c_oAscChartTypeSettings.comboCustom;
				}
			}
			return null;
		}



		function builder_CreateChart(nW, nH, sType, aCatNames, aSeriesNames, aSeries, nStyleIndex, aNumFormats) {
			let settings = new Asc.asc_ChartSettings();
			settings.type = ChartBuilderTypeToInternal(sType);
			let aAscSeries = [];
			let aAlphaBet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
			let oCat, i;
			if (aCatNames.length > 0) {
				let aNumCache = [];
				for (i = 0; i < aCatNames.length; ++i) {
					aNumCache.push({val: aCatNames[i] + ""});
				}
				oCat = {
					Formula: "Sheet1!$B$1:$" + AscFormat.CalcLiterByLength(aAlphaBet, aCatNames.length) + "$1",
					NumCache: aNumCache
				};
			}
			for (i = 0; i < aSeries.length; ++i) {
				let oAscSeries = new AscFormat.asc_CChartSeria();
				oAscSeries.Val.NumCache = [];
				let aData = aSeries[i];
				let sEndLiter = AscFormat.CalcLiterByLength(aAlphaBet, aData.length);
				oAscSeries.Val.Formula = 'Sheet1!' + '$B$' + (i + 2) + ':$' + sEndLiter + '$' + (i + 2);
				if (aSeriesNames[i]) {
					oAscSeries.TxCache.Formula = 'Sheet1!' + '$A$' + (i + 2);
					oAscSeries.TxCache.NumCache = [{
						numFormatStr: "General",
						isDateTimeFormat: false,
						val: aSeriesNames[i],
						isHidden: false
					}];
				}
				if (oCat) {
					oAscSeries.Cat = oCat;
				}

				if (Array.isArray(aNumFormats) && typeof (aNumFormats[i]) === "string")
					oAscSeries.FormatCode = aNumFormats[i];

				for (let j = 0; j < aData.length; ++j) {
					oAscSeries.Val.NumCache.push({
						numFormatStr: oAscSeries.FormatCode !== "" ? null : "General",
						isDateTimeFormat: false,
						val: aData[j],
						isHidden: false
					});
				}
				aAscSeries.push(oAscSeries);
			}

			let oChartSpace = AscFormat.DrawingObjectsController.prototype._getChartSpace(aAscSeries, settings, true);
			if (!oChartSpace) {
				return null;
			}
			oChartSpace.setBDeleted(false);
			oChartSpace.extX = nW;
			oChartSpace.extY = nH;
			if (AscFormat.isRealNumber(nStyleIndex)) {
				oChartSpace.setStyle(nStyleIndex);
			}
			AscFormat.CheckSpPrXfrm(oChartSpace);
			return oChartSpace;
		}

		function builder_CreateGroup(aDrawings, oController) {
			if (!oController) {
				return null;
			}
			var aForGroup = [];
			for (var i = 0; i < aDrawings.length; ++i) {
				if (!aDrawings[i].Drawing || !aDrawings[i].Drawing.canGroup()) {
					return null;
				}
				aForGroup.push(aDrawings[i].Drawing);
			}
			return oController.getGroup(aForGroup);
		}

		function builder_CreateSchemeColor(sColorId) {
			var oUniColor = new AscFormat.CUniColor();
			oUniColor.setColor(new AscFormat.CSchemeColor());
			switch (sColorId) {
				case "accent1": {
					oUniColor.color.id = 0;
					break;
				}
				case "accent2": {
					oUniColor.color.id = 1;
					break;
				}
				case "accent3": {
					oUniColor.color.id = 2;
					break;
				}
				case "accent4": {
					oUniColor.color.id = 3;
					break;
				}
				case "accent5": {
					oUniColor.color.id = 4;
					break;
				}
				case "accent6": {
					oUniColor.color.id = 5;
					break;
				}
				case "bg1": {
					oUniColor.color.id = 6;
					break;
				}
				case "bg2": {
					oUniColor.color.id = 7;
					break;
				}
				case "dk1": {
					oUniColor.color.id = 8;
					break;
				}
				case "dk2": {
					oUniColor.color.id = 9;
					break;
				}
				case "lt1": {
					oUniColor.color.id = 12;
					break;
				}
				case "lt2": {
					oUniColor.color.id = 13;
					break;
				}
				case "tx1": {
					oUniColor.color.id = 15;
					break;
				}
				case "tx2": {
					oUniColor.color.id = 16;
					break;
				}
				default: {
					oUniColor.color.id = 16;
					break;
				}
			}
			return oUniColor;
		}

		function builder_CreatePresetColor(sPresetColor) {
			var oUniColor = new AscFormat.CUniColor();
			oUniColor.setColor(new AscFormat.CPrstColor());
			oUniColor.color.id = sPresetColor;
			return oUniColor;
		}

		function builder_CreateGradientStop(oUniColor, nPos) {
			var Gs = new AscFormat.CGs();
			Gs.pos = nPos;
			Gs.color = oUniColor;
			return Gs;
		}

		function builder_CreateGradient(aGradientStop) {
			var oUniFill = new AscFormat.CUniFill();
			oUniFill.fill = new AscFormat.CGradFill();
			for (var i = 0; i < aGradientStop.length; ++i) {
				oUniFill.fill.colors.push(aGradientStop[i].Gs);
			}
			return oUniFill;
		}

		function builder_CreateLinearGradient(aGradientStop, Angle) {
			var oUniFill = builder_CreateGradient(aGradientStop);
			oUniFill.fill.lin = new AscFormat.GradLin();
			if (!AscFormat.isRealNumber(Angle)) {
				oUniFill.fill.lin.angle = 0;
			} else {
				oUniFill.fill.lin.angle = Angle;
			}
			return oUniFill;
		}

		function builder_CreateRadialGradient(aGradientStop) {
			var oUniFill = builder_CreateGradient(aGradientStop);
			oUniFill.fill.path = new AscFormat.GradPath();
			return oUniFill;
		}

		function builder_CreatePatternFill(sPatternType, BgColor, FgColor) {
			var oUniFill = new AscFormat.CUniFill();
			oUniFill.fill = new AscFormat.CPattFill();
			oUniFill.fill.ftype = AscCommon.global_hatch_offsets[sPatternType];
			oUniFill.fill.fgClr = FgColor && FgColor.Unicolor;
			oUniFill.fill.bgClr = BgColor && BgColor.Unicolor;
			return oUniFill;
		}

		function builder_CreateBlipFill(sImageUrl, sBlipFillType) {
			var oUniFill = new AscFormat.CUniFill();
			oUniFill.fill = new AscFormat.CBlipFill();
			oUniFill.fill.RasterImageId = sImageUrl;
			if (sBlipFillType === "tile") {
				oUniFill.fill.tile = new AscFormat.CBlipFillTile();
			} else if (sBlipFillType === "stretch") {
				oUniFill.fill.stretch = true;
			}
			return oUniFill;
		}

		function builder_CreateLine(nWidth, oFill) {
			if (nWidth === 0) {
				return new AscFormat.CreateNoFillLine();
			}
			var oLn = new AscFormat.CLn();
			oLn.w = nWidth;
			oLn.Fill = oFill.UniFill;
			return oLn;
		}

		function builder_CreateChartTitle(sTitle, nFontSize, bIsBold, oDrawingDocument) {
			if (typeof sTitle === "string" && sTitle.length > 0) {
				var oTitle = new AscFormat.CTitle();
				oTitle.setOverlay(false);
				oTitle.setTx(new AscFormat.CChartText());
				var oTextBody = AscFormat.CreateTextBodyFromString(sTitle, oDrawingDocument, oTitle.tx);
				if (AscFormat.isRealNumber(nFontSize)) {
					oTextBody.content.SetApplyToAll(true);
					oTextBody.content.AddToParagraph(new ParaTextPr({FontSize: nFontSize, Bold: bIsBold}));
					oTextBody.content.SetApplyToAll(false);
				}
				oTitle.tx.setRich(oTextBody);
				return oTitle;
			}
			return null;
		}


		function builder_CreateTitle(sTitle, nFontSize, bIsBold, oChartSpace) {
			if (typeof sTitle === "string" && sTitle.length > 0) {
				var oTitle = new AscFormat.CTitle();
				oTitle.setOverlay(false);
				oTitle.setTx(new AscFormat.CChartText());
				var oTextBody = AscFormat.CreateTextBodyFromString(sTitle, oChartSpace.getDrawingDocument(), oTitle.tx);
				if (AscFormat.isRealNumber(nFontSize)) {
					oTextBody.content.SetApplyToAll(true);
					oTextBody.content.AddToParagraph(new ParaTextPr({FontSize: nFontSize, Bold: bIsBold}));
					oTextBody.content.SetApplyToAll(false);
				}
				oTitle.tx.setRich(oTextBody);
				return oTitle;
			}
			return null;
		}

		function builder_SetChartTitle(oChartSpace, sTitle, nFontSize, bIsBold) {
			if (oChartSpace) {
				oChartSpace.chart.setTitle(builder_CreateChartTitle(sTitle, nFontSize, bIsBold, oChartSpace.getDrawingDocument()));
			}
		}

		function builder_SetChartHorAxisTitle(oChartSpace, sTitle, nFontSize, bIsBold) {
			if (oChartSpace) {
				var horAxis = oChartSpace.chart.plotArea.getHorizontalAxis();
				if (horAxis) {
					horAxis.setTitle(builder_CreateTitle(sTitle, nFontSize, bIsBold, oChartSpace));
				}
			}
		}

		function builder_SetChartVertAxisTitle(oChartSpace, sTitle, nFontSize, bIsBold) {
			if (oChartSpace) {
				var verAxis = oChartSpace.chart.plotArea.getVerticalAxis();
				if (verAxis) {
					if (typeof sTitle === "string" && sTitle.length > 0) {
						verAxis.setTitle(builder_CreateTitle(sTitle, nFontSize, bIsBold, oChartSpace));
						if (verAxis.title) {
							var _body_pr = new AscFormat.CBodyPr();
							_body_pr.reset();
							if (!verAxis.title.txPr) {
								verAxis.title.setTxPr(AscFormat.CreateTextBodyFromString("", oChartSpace.getDrawingDocument(), verAxis.title));
							}
							var _text_body = verAxis.title.txPr;
							_text_body.setBodyPr(_body_pr);
							verAxis.title.setOverlay(false);
						}
					} else {
						verAxis.setTitle(null);
					}
				}
			}
		}

		function builder_SetChartVertAxisOrientation(oChartSpace, bIsMinMax) {
			if (oChartSpace) {
				var verAxis = oChartSpace.chart.plotArea.getVerticalAxis();
				if (verAxis) {
					if (!verAxis.scaling)
						verAxis.setScaling(new AscFormat.CScaling());
					var scaling = verAxis.scaling;
					if (bIsMinMax) {
						scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
					} else {
						scaling.setOrientation(AscFormat.ORIENTATION_MAX_MIN);
					}
				}
			}
		}

		function builder_SetChartHorAxisOrientation(oChartSpace, bIsMinMax) {
			if (oChartSpace) {
				var horAxis = oChartSpace.chart.plotArea.getHorizontalAxis();
				if (horAxis) {
					if (!horAxis.scaling)
						horAxis.setScaling(new AscFormat.CScaling());
					var scaling = horAxis.scaling;
					if (bIsMinMax) {
						scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
					} else {
						scaling.setOrientation(AscFormat.ORIENTATION_MAX_MIN);
					}
				}
			}
		}

		function builder_SetChartLegendPos(oChartSpace, sLegendPos) {

			if (oChartSpace && oChartSpace.chart) {
				if (sLegendPos === "none") {
					if (oChartSpace.chart.legend) {
						oChartSpace.chart.setLegend(null);
					}
				} else {
					var nLegendPos = null;
					switch (sLegendPos) {
						case "left": {
							nLegendPos = Asc.c_oAscChartLegendShowSettings.left;
							break;
						}
						case "top": {
							nLegendPos = Asc.c_oAscChartLegendShowSettings.top;
							break;
						}
						case "right": {
							nLegendPos = Asc.c_oAscChartLegendShowSettings.right;
							break;
						}
						case "bottom": {
							nLegendPos = Asc.c_oAscChartLegendShowSettings.bottom;
							break;
						}
					}
					if (null !== nLegendPos) {
						if (!oChartSpace.chart.legend) {
							oChartSpace.chart.setLegend(new AscFormat.CLegend());
						}
						if (oChartSpace.chart.legend.legendPos !== nLegendPos)
							oChartSpace.chart.legend.setLegendPos(nLegendPos);
						if (oChartSpace.chart.legend.overlay !== false) {
							oChartSpace.chart.legend.setOverlay(false);
						}
					}
				}
			}
		}

		function builder_SetObjectFontSize(oObject, nFontSize, oDrawingDocument) {
			if (!oObject) {
				return;
			}
			if (!oObject.txPr) {
				oObject.setTxPr(new AscFormat.CTextBody());
			}
			if (!oObject.txPr.bodyPr) {
				oObject.txPr.setBodyPr(new AscFormat.CBodyPr());
			}
			if (!oObject.txPr.content) {
				oObject.txPr.setContent(new AscFormat.CDrawingDocContent(oObject.txPr, oDrawingDocument, 0, 0, 100, 500, false, false, true));
			}
			var oPr = oObject.txPr.content.Content[0].Pr.Copy();
			if (!oPr.DefaultRunPr) {
				oPr.DefaultRunPr = new AscCommonWord.CTextPr();
			}
			oPr.DefaultRunPr.FontSize = nFontSize;
			oObject.txPr.content.Content[0].Set_Pr(oPr);
		}

		function builder_SetLegendFontSize(oChartSpace, nFontSize) {
			builder_SetObjectFontSize(oChartSpace.chart.legend, nFontSize, oChartSpace.getDrawingDocument());
		}

		function builder_SetHorAxisFontSize(oChartSpace, nFontSize) {
			builder_SetObjectFontSize(oChartSpace.chart.plotArea.getHorizontalAxis(), nFontSize, oChartSpace.getDrawingDocument());
		}

		function builder_SetVerAxisFontSize(oChartSpace, nFontSize) {
			builder_SetObjectFontSize(oChartSpace.chart.plotArea.getVerticalAxis(), nFontSize, oChartSpace.getDrawingDocument());
		}


		function builder_SetShowPointDataLabel(oChartSpace, nSeriesIndex, nPointIndex, bShowSerName, bShowCatName, bShowVal, bShowPerecent) {
			if (oChartSpace && oChartSpace.chart && oChartSpace.chart.plotArea && oChartSpace.chart.plotArea.charts[0]) {
				var oChart = oChartSpace.chart.plotArea.charts[0];
				var bPieChart = oChart.getObjectType() === AscDFH.historyitem_type_PieChart || oChart.getObjectType() === AscDFH.historyitem_type_DoughnutChart;
				var ser = oChart.series[nSeriesIndex];
				if (ser) {
					{
						if (!ser.dLbls) {
							if (oChart.dLbls) {
								ser.setDLbls(oChart.dLbls.createDuplicate());
							} else {
								ser.setDLbls(new AscFormat.CDLbls());
								ser.dLbls.setSeparator(",");
								ser.dLbls.setShowSerName(false);
								ser.dLbls.setShowCatName(false);
								ser.dLbls.setShowVal(false);
								ser.dLbls.setShowLegendKey(false);
								if (bPieChart) {
									ser.dLbls.setShowPercent(false);
								}
								ser.dLbls.setShowBubbleSize(false);
							}
						}
						var dLbl = ser.dLbls && ser.dLbls.findDLblByIdx(nPointIndex);
						if (!dLbl) {
							dLbl = new AscFormat.CDLbl();
							dLbl.setIdx(nPointIndex);
							if (ser.dLbls.txPr) {
								dLbl.merge(ser.dLbls);
							}
							ser.dLbls.addDLbl(dLbl);
						}
						dLbl.setSeparator(",");
						dLbl.setShowSerName(true == bShowSerName);
						dLbl.setShowCatName(true == bShowCatName);
						dLbl.setShowVal(true == bShowVal);
						dLbl.setShowLegendKey(false);
						if (bPieChart) {
							dLbl.setShowPercent(true === bShowPerecent);
						}
						dLbl.setShowBubbleSize(false);
					}
				}
			}
		}

		function builder_SetShowDataLabels(oChartSpace, bShowSerName, bShowCatName, bShowVal, bShowPerecent) {
			if (oChartSpace && oChartSpace.chart && oChartSpace.chart.plotArea && oChartSpace.chart.plotArea.charts[0]) {
				var oChart = oChartSpace.chart.plotArea.charts[0];
				var bPieChart = oChart.getObjectType() === AscDFH.historyitem_type_PieChart || oChart.getObjectType() === AscDFH.historyitem_type_DoughnutChart;
				if (false == bShowSerName && false == bShowCatName && false == bShowVal && (bPieChart && bShowPerecent === false)) {
					if (oChart.dLbls) {
						oChart.setDLbls(null);
					}
				}
				if (!oChart.dLbls) {
					oChart.setDLbls(new AscFormat.CDLbls());
				}
				oChart.dLbls.setSeparator(",");
				oChart.dLbls.setShowSerName(true == bShowSerName);
				oChart.dLbls.setShowCatName(true == bShowCatName);
				oChart.dLbls.setShowVal(true == bShowVal);
				oChart.dLbls.setShowLegendKey(false);
				if (bPieChart) {
					oChart.dLbls.setShowPercent(true === bShowPerecent);
				}

				oChart.dLbls.setShowBubbleSize(false);
			}
		}

		function builder_SetChartAxisLabelsPos(oAxis, sPosition) {
			if (!oAxis || !oAxis.setTickLblPos) {
				return;
			}
			var nPositionType = null;
			var c_oAscTickLabelsPos = window['Asc'].c_oAscTickLabelsPos;
			switch (sPosition) {
				case "high": {
					nPositionType = c_oAscTickLabelsPos.TICK_LABEL_POSITION_HIGH;
					break;
				}
				case "low": {
					nPositionType = c_oAscTickLabelsPos.TICK_LABEL_POSITION_LOW;
					break;
				}
				case "nextTo": {
					nPositionType = c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO;
					break;
				}
				case "none": {
					nPositionType = c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE;
					break;
				}
			}
			if (nPositionType !== null) {
				oAxis.setTickLblPos(nPositionType);
			}
		}

		function builder_SetChartVertAxisTickLablePosition(oChartSpace, sPosition) {
			if (oChartSpace) {
				builder_SetChartAxisLabelsPos(oChartSpace.chart.plotArea.getVerticalAxis(), sPosition);
			}
		}

		function builder_SetChartHorAxisTickLablePosition(oChartSpace, sPosition) {
			if (oChartSpace) {
				builder_SetChartAxisLabelsPos(oChartSpace.chart.plotArea.getHorizontalAxis(), sPosition);
			}
		}

		function builder_GetTickMark(sTickMark) {
			var nNewTickMark = null;
			switch (sTickMark) {
				case 'cross': {
					nNewTickMark = Asc.c_oAscTickMark.TICK_MARK_CROSS;
					break;
				}
				case 'in': {
					nNewTickMark = Asc.c_oAscTickMark.TICK_MARK_IN;
					break;
				}
				case 'none': {
					nNewTickMark = Asc.c_oAscTickMark.TICK_MARK_NONE;
					break;
				}
				case 'out': {
					nNewTickMark = Asc.c_oAscTickMark.TICK_MARK_OUT;
					break;
				}
			}
			return nNewTickMark;
		}

		function builder_SetChartAxisMajorTickMark(oAxis, sTickMark) {
			if (!oAxis) {
				return;
			}
			var nNewTickMark = builder_GetTickMark(sTickMark);
			if (nNewTickMark !== null) {
				oAxis.setMajorTickMark(nNewTickMark);
			}
		}

		function builder_SetChartAxisMinorTickMark(oAxis, sTickMark) {
			if (!oAxis) {
				return;
			}
			var nNewTickMark = builder_GetTickMark(sTickMark);
			if (nNewTickMark !== null) {
				oAxis.setMinorTickMark(nNewTickMark);
			}
		}

		function builder_SetChartHorAxisMajorTickMark(oChartSpace, sTickMark) {
			if (oChartSpace) {
				builder_SetChartAxisMajorTickMark(oChartSpace.chart.plotArea.getHorizontalAxis(), sTickMark);
			}
		}

		function builder_SetChartHorAxisMinorTickMark(oChartSpace, sTickMark) {
			if (oChartSpace) {
				builder_SetChartAxisMinorTickMark(oChartSpace.chart.plotArea.getHorizontalAxis(), sTickMark);
			}
		}

		function builder_SetChartVerAxisMajorTickMark(oChartSpace, sTickMark) {
			if (oChartSpace) {
				builder_SetChartAxisMajorTickMark(oChartSpace.chart.plotArea.getVerticalAxis(), sTickMark);
			}
		}

		function builder_SetChartVerAxisMinorTickMark(oChartSpace, sTickMark) {
			if (oChartSpace) {
				builder_SetChartAxisMinorTickMark(oChartSpace.chart.plotArea.getVerticalAxis(), sTickMark);
			}
		}

		function builder_SetAxisMajorGridlines(oAxis, oLn) {
			if (oAxis) {
				if (!oAxis.majorGridlines) {
					oAxis.setMajorGridlines(new AscFormat.CSpPr());
				}
				oAxis.majorGridlines.setLn(oLn);
				if (!oAxis.majorGridlines.Fill && !oAxis.majorGridlines.ln) {
					oAxis.setMajorGridlines(null);
				}
			}
		}

		function builder_SetAxisMinorGridlines(oAxis, oLn) {
			if (oAxis) {
				if (!oAxis.minorGridlines) {
					oAxis.setMinorGridlines(new AscFormat.CSpPr());
				}
				oAxis.minorGridlines.setLn(oLn);
				if (!oAxis.minorGridlines.Fill && !oAxis.minorGridlines.ln) {
					oAxis.setMinorGridlines(null);
				}
			}
		}

		function builder_SetHorAxisMajorGridlines(oChartSpace, oLn) {
			builder_SetAxisMajorGridlines(oChartSpace.chart.plotArea.getVerticalAxis(), oLn);
		}

		function builder_SetHorAxisMinorGridlines(oChartSpace, oLn) {
			builder_SetAxisMinorGridlines(oChartSpace.chart.plotArea.getVerticalAxis(), oLn);
		}

		function builder_SetVerAxisMajorGridlines(oChartSpace, oLn) {
			builder_SetAxisMajorGridlines(oChartSpace.chart.plotArea.getHorizontalAxis(), oLn);
		}

		function builder_SetVerAxisMinorGridlines(oChartSpace, oLn) {
			builder_SetAxisMinorGridlines(oChartSpace.chart.plotArea.getHorizontalAxis(), oLn);
		}

		function CreatePresentationTableStyles(Styles, IdMap) {
			function CreateThemedStyle1(schemeId) {
				if (schemeId == 8)
					var style = new CStyle("Themed Style 1", null, null, styletype_Table);
				else
					var style = new CStyle("Themed Style 1 - Accent " + (schemeId + 1), null, null, styletype_Table);
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideV:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.6)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.8)
								}
						}
				};
				style.TableBand1Horz.Set_FromObject(styleObject);
				style.TableBand1Vert.Set_FromObject(styleObject);
		
				styleObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
						},
					TableCellPr:
						{}
				};
				style.TableLastCol.Set_FromObject(styleObject);
				style.TableFirstCol.Set_FromObject(styleObject);
		
				var styleLastRowObject = {
					TableCellPr:
						{},
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
						}
				}
				styleLastRowObject.TableCellPr.TableCellBorders =
					{
						Right:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_None
							},
						Left:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_None
							},
		
					};
				style.TableLastRow.Set_FromObject(styleLastRowObject);
				styleObject.TableCellPr.Shd =
					{
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0)
					}
				styleObject.TableCellPr.TableCellBorders =
					{
						Bottom:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_Single
							}
					};
				styleObject.TextPr =
					{
						Bold: true,
						FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
					};
				style.TableFirstRow.Set_FromObject(styleObject);
				return style;
			}
		
			function CreateThemedStyle2(schemeId) {
				if (schemeId == 8)
					var style = new CStyle("Themed Style 2", null, null, styletype_Table);
				else
					var style = new CStyle("Themed Style 2 - Accent " + (schemeId + 1), null, null, styletype_Table);
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								InsideV:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.8)
								}
						}
				};
				style.TableBand1Horz.Set_FromObject(styleObject);
				style.TableBand1Vert.Set_FromObject(styleObject);
		
				styleObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
						},
					TableCellPr:
						{}
				};
				style.TableLastCol.Set_FromObject(styleObject);
				style.TableFirstCol.Set_FromObject(styleObject);
				styleObject.TableCellPr.Shd =
					{
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0)
					}
				styleObject.TableCellPr.TableCellBorders =
					{
						Top:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_Single
							}
					};
				style.TableLastRow.Set_FromObject(styleObject);
				styleObject.TableCellPr.TableCellBorders =
					{
						Bottom:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_Single
							}
					};
				styleObject.TextPr =
					{
						Bold: true,
						FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
					};
				style.TableFirstRow.Set_FromObject(styleObject);
				return style;
			}
		
			function CreateLightStyle1(schemeId) {
				if (schemeId == 8)
					var style = new CStyle("Light Style 1", null, null, styletype_Table);
				else
					var style = new CStyle("Light Style 1 - Accent " + (schemeId + 1), null, null, styletype_Table);
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								InsideV:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.4)
								}
						}
				};
				style.TableBand1Horz.Set_FromObject(styleObject);
				style.TableBand1Vert.Set_FromObject(styleObject);
		
				styleObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
						},
					TableCellPr:
						{}
				};
				style.TableLastCol.Set_FromObject(styleObject);
				style.TableFirstCol.Set_FromObject(styleObject);
		
				styleObject.TableCellPr.TableCellBorders =
					{
						Top:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_Single
							}
					};
				style.TableLastRow.Set_FromObject(styleObject);
				styleObject.TableCellPr.TableCellBorders =
					{
						Bottom:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_Single
							}
					};
				styleObject.TextPr =
					{
						Bold: true,
						FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
					};
				styleObject.TableCellPr.Shd =
					{
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
					}
				style.TableFirstRow.Set_FromObject(styleObject);
				return style;
			}
		
			function CreateLightStyle2(schemeId) {
				if (schemeId == 8)
					var style = new CStyle("Light Style 2", null, null, styletype_Table);
				else
					var style = new CStyle("Light Style 2 - Accent " + (schemeId + 1), null, null, styletype_Table);
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								InsideV:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
								},
							TableCellBorders:
								{
									Top:
										{
											Color: {r: 0, g: 0, b: 0},
											Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
											Space: 0,
											Size: 12700 / 36000,
											Value: border_Single
										},
									Bottom:
										{
											Color: {r: 0, g: 0, b: 0},
											Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
											Space: 0,
											Size: 12700 / 36000,
											Value: border_Single
										}
								},
						}
				};
				var styleTableBand1Vert =
					{
						TableCellPr:
							{
								TableCellBorders:
									{
										Left:
											{
												Color: {r: 0, g: 0, b: 0},
												Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
												Space: 0,
												Size: 12700 / 36000,
												Value: border_Single
											},
										Right:
											{
												Color: {r: 0, g: 0, b: 0},
												Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
												Space: 0,
												Size: 12700 / 36000,
												Value: border_Single
											}
									}
							}
					}
				style.TableBand1Horz.Set_FromObject(styleObject);
				style.TableBand1Vert.Set_FromObject(styleTableBand1Vert);
		
				styleObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
						},
					TableCellPr:
						{}
				};
				style.TableLastCol.Set_FromObject(styleObject);
				style.TableFirstCol.Set_FromObject(styleObject);
		
				styleObject.TableCellPr.TableCellBorders =
					{
						Top:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
								Space: 0,
								Size: 38100 / 36000,
								Value: border_Single
							}
					};
				styleObject.TableCellPr.Shd =
					{
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
					}
				style.TableLastRow.Set_FromObject(styleObject);
				styleObject.TableCellPr.TableCellBorders =
					{
						Bottom:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_Single
							}
					};
				styleObject.TextPr =
					{
						Bold: true,
						FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
					};
		
				styleObject.TableCellPr.Shd =
					{
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0)
					}
		
				style.TableFirstRow.Set_FromObject(styleObject);
				return style;
			}
		
			function CreateLightStyle3(schemeId) {
				if (schemeId == 8)
					var style = new CStyle("Light Style 3", null, null, styletype_Table);
				else
					var style = new CStyle("Light Style 3 - Accent " + (schemeId + 1), null, null, styletype_Table);
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideV:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.2)
								}
						}
				};
				style.TableBand1Horz.Set_FromObject(styleObject);
				style.TableBand1Vert.Set_FromObject(styleObject);
		
				styleObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
						},
					TableCellPr:
						{}
				};
				style.TableLastCol.Set_FromObject(styleObject);
				style.TableFirstCol.Set_FromObject(styleObject);
		
				styleObject.TableCellPr.TableCellBorders =
					{
						Top:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
								Space: 0,
								Size: 38100 / 36000,
								Value: border_Single
							}
					};
		
				styleObject.TableCellPr.Shd =
					{
						Shd:
							{
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
							}
					};
		
				style.TableLastRow.Set_FromObject(styleObject);
		
				styleObject.TableCellPr.TableCellBorders =
					{
						Bottom:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
								Space: 0,
								Size: 38100 / 36000,
								Value: border_Single
							}
					};
		
				styleObject.TextPr =
					{
						Bold: true,
						FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
					};
				style.TableFirstRow.Set_FromObject(styleObject);
				return style;
			}
		
			function CreateMediumStyle1(schemeId) {
		
				if (schemeId == 8)
					var style = new CStyle("Medium Style 1", null, null, styletype_Table);
				else
					var style = new CStyle("Medium Style 1 - Accent " + (schemeId + 1), null, null, styletype_Table);
		
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideV:
									{
										Value: border_None
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.2)
								}
						}
				};
				style.TableBand1Horz.Set_FromObject(styleObject);
				style.TableBand1Vert.Set_FromObject(styleObject);
		
				styleObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
						},
					TableCellPr:
						{}
				};
				style.TableLastCol.Set_FromObject(styleObject);
				style.TableFirstCol.Set_FromObject(styleObject);
		
				styleObject.TableCellPr.TableCellBorders =
					{
						Top:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
								Space: 0,
								Size: 38100 / 36000,
								Value: border_Single
							}
					};
				styleObject.TableCellPr.Shd =
					{
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
					}
				style.TableLastRow.Set_FromObject(styleObject);
		
				styleObject.TableCellPr.TableCellBorders =
					{
						Top:
							{
								Value: border_None
							}
					};
				styleObject.TableCellPr.Shd =
					{
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0)
					}
				styleObject.TableCellPr.TableCellBorders =
					{
						Bottom:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
								Space: 0,
								Size: 38100 / 36000,
								Value: border_Single
							}
					};
				styleObject.TextPr =
					{
						Bold: true,
						FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
					};
				style.TableFirstRow.Set_FromObject(styleObject);
				return style;
			}
		
			function CreateMediumStyle2(schemeId) {
				if (schemeId == 8)
					var style = new CStyle("Medium Style 2", null, null, styletype_Table);
				else
					var style = new CStyle("Medium Style 2 - Accent " + (schemeId + 1), null, null, styletype_Table);
				//style.Id = "{" + GUID() + "}";
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideV:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.2)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.4)
								}
						}
				};
				style.TableBand1Horz.Set_FromObject(styleObject);
				style.TableBand1Vert.Set_FromObject(styleObject);
		
				styleObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
						},
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0)
								}
						}
				};
		
				style.TableLastCol.Set_FromObject(styleObject);
				style.TableFirstCol.Set_FromObject(styleObject);
		
				styleObject.TableCellPr.TableCellBorders =
					{
						Top:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
								Space: 0,
								Size: 38100 / 36000,
								Value: border_Single
							}
					};
		
				style.TableLastRow.Set_FromObject(styleObject);
				styleObject.TableCellPr.TableCellBorders =
					{
						Bottom:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
								Space: 0,
								Size: 38100 / 36000,
								Value: border_Single
							}
					};
		
				style.TableFirstRow.Set_FromObject(styleObject);
				return style;
			}
		
			function CreateMediumStyle3(schemeId) {
				if (schemeId == 8)
					var style = new CStyle("Medium Style 3", null, null, styletype_Table);
				else
					var style = new CStyle("Medium Style 3 - Accent " + (schemeId + 1), null, null, styletype_Table);
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0),
										Space: 0,
										Size: 38100 / 36000,
										Value: border_Single
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0),
										Space: 0,
										Size: 38100 / 36000,
										Value: border_Single
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								InsideV:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(2, 0.2)
								}
						}
				};
				style.TableBand1Horz.Set_FromObject(styleObject);
				style.TableBand1Vert.Set_FromObject(styleObject);
		
				styleObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
						},
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0)
								}
						}
				};
				style.TableLastCol.Set_FromObject(styleObject);
				style.TableFirstCol.Set_FromObject(styleObject);
				styleObject.TableCellPr.Shd =
					{
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
					}
				styleObject.TableCellPr.TableCellBorders =
					{
						Top:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0),
								Space: 0,
								Size: 38100 / 36000,
								Value: border_Single
							}
					};
				styleObject.TextPr =
					{
						Bold: true,
						FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
					};
				style.TableLastRow.Set_FromObject(styleObject);
				styleObject.TableCellPr.TableCellBorders =
					{
						Bottom:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0),
								Space: 0,
								Size: 38100 / 36000,
								Value: border_Single
							}
					};
				styleObject.TextPr =
					{
						Bold: true,
						FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
					};
				styleObject.TableCellPr.Shd =
					{
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0)
					}
				style.TableFirstRow.Set_FromObject(styleObject);
				return style;
			}
		
			function CreateMediumStyle4(schemeId) {
				if (schemeId == 8)
					var style = new CStyle("Medium Style 4", null, null, styletype_Table);
				else
					var style = new CStyle("Medium Style 4 - Accent " + (schemeId + 1), null, null, styletype_Table);
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideV:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.2)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.4)
								}
						}
				};
				style.TableBand1Horz.Set_FromObject(styleObject);
				style.TableBand1Vert.Set_FromObject(styleObject);
		
				styleObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
						},
					TableCellPr:
						{}
				};
		
				style.TableFirstCol.Set_FromObject(styleObject);
				style.TableLastCol.Set_FromObject(styleObject)
		
				styleObject.TableCellPr.TableCellBorders =
					{
						Top:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
								Space: 0,
								Size: 38100 / 36000,
								Value: border_Single
							}
					};
				styleObject.TableCellPr.Shd =
					{
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.2)
					}
				style.TableLastRow.Set_FromObject(styleObject);
				styleObject.TableCellPr.TableCellBorders =
					{
						Bottom:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_Single
							}
					};
				styleObject.TableCellPr.Shd =
					{
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.2)
					}
				styleObject.TextPr =
					{
						Bold: true,
						FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
					};
				style.TableFirstRow.Set_FromObject(styleObject);
				return style;
			}
		
			function CreateNoStyle1(schemeId) {
		
				var style = new CStyle("No Style, No Grid", null, null, styletype_Table);
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								InsideV:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
								}
						}
				};
				style.TableBand1Horz.Set_FromObject(styleObject);
				style.TableBand1Vert.Set_FromObject(styleObject);
		
				styleObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
						},
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
								}
						}
				};
				style.TableLastCol.Set_FromObject(styleObject);
				style.TableFirstCol.Set_FromObject(styleObject);
		
				styleObject.TableCellPr.TableCellBorders =
					{
						Top:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_None
							}
					};
				style.TableLastRow.Set_FromObject(styleObject);
				styleObject.TableCellPr.TableCellBorders =
					{
						Bottom:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_None
							}
					};
				styleObject.TextPr =
					{
						Bold: true,
						FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
					};
				style.TableFirstRow.Set_FromObject(styleObject);
				return style;
			}
		
			function CreateDarkStyle1(schemeId) {
				if (schemeId == 8)
					var style = new CStyle("Dark Style 1", null, null, styletype_Table);
				else
					var style = new CStyle("Dark Style 1 - Accent " + (schemeId + 1), null, null, styletype_Table);
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideV:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, -0.6)
								}
						}
				};
				style.TableBand1Vert.Set_FromObject(styleObject);
				style.TableBand1Horz.Set_FromObject(styleObject);
		
		
				var styleFirstLastColumnObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
						},
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, -0.6)
								},
							TableCellBorders:
								{
									Right:
										{
											Color: {r: 0, g: 0, b: 0},
											Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
											Space: 0,
											Size: 38100 / 36000,
											Value: border_Single
										}
								}
						}
		
				};
				style.TableFirstCol.Set_FromObject(styleFirstLastColumnObject);
				styleFirstLastColumnObject.TableCellPr.TableCellBorders = {
					Left:
						{
							Color: {r: 0, g: 0, b: 0},
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
							Space: 0,
							Size: 38100 / 36000,
							Value: border_Single
						},
					Right:
						{
							Value: border_None
						}
				};
				style.TableLastCol.Set_FromObject(styleFirstLastColumnObject);
		
				var styleLastRowObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
						},
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, -0.4)
								},
						}
				};
				styleLastRowObject.TableCellPr.TableCellBorders = {
					Top:
						{
							Color: {r: 0, g: 0, b: 0},
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
							Space: 0,
							Size: 38100 / 36000,
							Value: border_Single
						}
				}
				style.TableLastRow.Set_FromObject(styleLastRowObject);
		
				var styleFirstRowObject = {
					TableCellPr:
						{
							TextPr:
								{
									Bold: true,
									FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
								},
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
								},
							TableCellBorders:
								{
									Bottom:
										{
											Color: {r: 0, g: 0, b: 0},
											Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
											Space: 0,
											Size: 38100 / 36000,
											Value: border_Single
										}
								}
						}
				}
				style.TableFirstRow.Set_FromObject(styleFirstRowObject);
				return style;
			}
		
			function CreateNoStyle2(schemeId) {
				var style = new CStyle("No Style, Table Grid", null, null, styletype_Table);
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									},
		
								InsideV:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_Single
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
								}
						}
				};
				style.TableBand1Horz.Set_FromObject(styleObject);
				style.TableBand1Vert.Set_FromObject(styleObject);
		
				styleObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
						},
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
								}
						}
				};
				style.TableLastCol.Set_FromObject(styleObject);
				style.TableFirstCol.Set_FromObject(styleObject);
		
				styleObject.TableCellPr.TableCellBorders =
					{
						Top:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_Single
							}
					};
				style.TableLastRow.Set_FromObject(styleObject);
				styleObject.TableCellPr.TableCellBorders =
					{
						Bottom:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_Single
							}
					};
				styleObject.TextPr =
					{
						Bold: true,
						FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
					};
				style.TableFirstRow.Set_FromObject(styleObject);
				return style;
			}
		
			function CreateBlackStyle(schemeId) {
				var style = new CStyle("Dark Style 1", null, null, styletype_Table);
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								InsideV:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0.2)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0.4)
								}
						}
				};
				style.TableBand1Horz.Set_FromObject(styleObject);
				style.TableBand1Vert.Set_FromObject(styleObject);
		
				styleObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
						},
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
								}
						}
				};
				style.TableLastCol.Set_FromObject(styleObject);
				style.TableFirstCol.Set_FromObject(styleObject);
		
				styleObject.TableCellPr.TableCellBorders =
					{
						Top:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_None
							}
					};
				style.TableLastRow.Set_FromObject(styleObject);
				styleObject.TableCellPr.TableCellBorders =
					{
						Bottom:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_Single
							}
					};
				styleObject.TextPr =
					{
						Bold: true,
						FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
					};
				style.TableFirstRow.Set_FromObject(styleObject);
				return style;
			}
		
			function CreateDarkStyle2(schemeId) {
				var style = new CStyle("Dark Style 2 - Accent " + (schemeId + 1) + "/" + "Accent " + (schemeId + 2), null, null, styletype_Table);
				style.TablePr.Set_FromObject(
					{
						TableBorders:
							{
								Left:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Right:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Top:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								Bottom:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								InsideH:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									},
		
								InsideV:
									{
										Color: {r: 0, g: 0, b: 0},
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0),
										Space: 0,
										Size: 12700 / 36000,
										Value: border_None
									}
							}
					}
				);
				style.TableWholeTable.Set_FromObject(
					{
						TextPr:
							{
								FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
							},
						TableCellPr:
							{
								Shd:
									{
										Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.2)
									}
							}
					}
				);
				var styleObject = {
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.4)
								}
						}
				};
				style.TableBand1Horz.Set_FromObject(styleObject);
				style.TableBand1Vert.Set_FromObject(styleObject);
		
				styleObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
						},
					TableCellPr:
						{
							Shd:
								{
									Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId + 1, 0)
								}
						}
				};
				var styleLastObject = {
					TextPr:
						{
							Bold: true,
							FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
							Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0)
						},
					TableCellPr:
						{}
				}
				style.TableLastCol.Set_FromObject(styleLastObject);
				style.TableFirstCol.Set_FromObject(styleLastObject);
		
				styleLastObject.TableCellPr.Shd =
					{
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(schemeId, 0.2)
					}
				styleLastObject.TableCellPr.TableCellBorders =
					{
						Top:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(8, 0),
								Space: 0,
								Size: 38100 / 36000,
								Value: border_Single
							}
					};
		
				style.TableLastRow.Set_FromObject(styleLastObject);
				styleObject.TableCellPr.TableCellBorders =
					{
						Bottom:
							{
								Color: {r: 0, g: 0, b: 0},
								Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0),
								Space: 0,
								Size: 12700 / 36000,
								Value: border_None
							}
					};
				styleObject.TextPr =
					{
						Bold: true,
						FontRef: AscFormat.CreateFontRef(AscFormat.fntStyleInd_minor, AscFormat.builder_CreatePresetColor("black")),
						Unifill: AscFormat.CreateUnifillSolidFillSchemeColor(12, 0)
					};
				style.TableFirstRow.Set_FromObject(styleObject);
				return style;
			}
		
			var def, style;
		
			style = CreateNoStyle1(8);
			Styles.Add(style);
			IdMap[style.Id] = true;
		
			for (var i = 0; i < 6; ++i) {
				style = CreateThemedStyle1(i);
				Styles.Add(style);
				IdMap[style.Id] = true;
			}
			style = CreateNoStyle2(8);
			Styles.Add(style);
			IdMap[style.Id] = true;
		
			for (var i = 0; i < 6; ++i) {
				style = CreateThemedStyle2(i);
				Styles.Add(style);
				IdMap[style.Id] = true;
			}
		
			style = CreateLightStyle1(8);
			Styles.Add(style);
			IdMap[style.Id] = true;
		
			for (var i = 0; i < 6; ++i) {
				style = CreateLightStyle1(i);
				Styles.Add(style);
				IdMap[style.Id] = true;
			}
			style = CreateLightStyle2(8);
			Styles.Add(style);
			IdMap[style.Id] = true;
		
			for (var i = 0; i < 6; ++i) {
				style = CreateLightStyle2(i);
				Styles.Add(style);
				IdMap[style.Id] = true;
			}
			style = CreateLightStyle3(8);
			Styles.Add(style);
			IdMap[style.Id] = true;
		
			for (var i = 0; i < 6; ++i) {
				style = CreateLightStyle3(i);
				Styles.Add(style);
				IdMap[style.Id] = true;
			}
		
			style = CreateMediumStyle1(8);
			Styles.Add(style);
			IdMap[style.Id] = true;
		
			for (var i = 0; i < 6; ++i) {
				style = CreateMediumStyle1(i);
				Styles.Add(style);
				IdMap[style.Id] = true;
			}
			style = CreateMediumStyle2(8);
			Styles.Add(style);
			IdMap[style.Id] = true;
		
			def = CreateMediumStyle2(0);
			Styles.Add(def);
			IdMap[def.Id] = true;
		
			for (var i = 1; i < 6; ++i) {
				style = CreateMediumStyle2(i);
				Styles.Add(style);
				IdMap[style.Id] = true;
			}
			style = CreateMediumStyle3(8);
			Styles.Add(style);
			IdMap[style.Id] = true;
		
			for (var i = 0; i < 6; ++i) {
				style = CreateMediumStyle3(i);
				Styles.Add(style);
				IdMap[style.Id] = true;
			}
			style = CreateMediumStyle4(8);
			Styles.Add(style);
			IdMap[style.Id] = true;
		
			for (var i = 0; i < 6; ++i) {
				style = CreateMediumStyle4(i);
				Styles.Add(style);
				IdMap[style.Id] = true;
			}
		
			style = CreateBlackStyle(2);
			Styles.Add(style);
			IdMap[style.Id] = true;
		
			for (var i = 0; i < 6; ++i) {
				style = CreateDarkStyle1(i);
				Styles.Add(style);
				IdMap[style.Id] = true;
			}
		
			for (var i = 0; i < 3; i++) {
				style = CreateDarkStyle2(2 * i);
				Styles.Add(style);
				IdMap[style.Id] = true;
			}
		
			const arrStylesId = Object.keys(Styles.Style);
			for (let i = 0; i < arrStylesId.length; i += 1) {
				const oStyle = Styles.Style[arrStylesId[i]];
				oStyle.SetStyleId(getDefaultGUIDTableStyleByName(oStyle.Get_Name()));
			}
		
			return def.Id;
		}
		function getDefaultGUIDTableStyleByName(sName) {
			switch (sName) {
				case "Medium Style 2 - Accent 1":
					return "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
				case "No Style, No Grid":
					return "{2D5ABB26-0587-4C30-8999-92F81FD0307C}";
				case "Themed Style 1 - Accent 1":
					return "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}";
				case "Themed Style 1 - Accent 2":
					return "{284E427A-3D55-4303-BF80-6455036E1DE7}";
				case "Themed Style 1 - Accent 3":
					return "{69C7853C-536D-4A76-A0AE-DD22124D55A5}";
				case "Themed Style 1 - Accent 4":
					return "{775DCB02-9BB8-47FD-8907-85C794F793BA}";
				case "Themed Style 1 - Accent 5":
					return "{35758FB7-9AC5-4552-8A53-C91805E547FA}";
				case "Themed Style 1 - Accent 6":
					return "{08FB837D-C827-4EFA-A057-4D05807E0F7C}";
				case "No Style, Table Grid":
					return "{5940675A-B579-460E-94D1-54222C63F5DA}";
				case "Themed Style 2 - Accent 1":
					return "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}";
				case "Themed Style 2 - Accent 2":
					return "{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}";
				case "Themed Style 2 - Accent 3":
					return "{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}";
				case "Themed Style 2 - Accent 4":
					return "{E269D01E-BC32-4049-B463-5C60D7B0CCD2}";
				case "Themed Style 2 - Accent 5":
					return "{327F97BB-C833-4FB7-BDE5-3F7075034690}";
				case "Themed Style 2 - Accent 6":
					return "{638B1855-1B75-4FBE-930C-398BA8C253C6}";
				case "Light Style 1":
					return "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}";
				case "Light Style 1 - Accent 1":
					return "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}";
				case "Light Style 1 - Accent 2":
					return "{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}";
				case "Light Style 1 - Accent 3":
					return "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}";
				case "Light Style 1 - Accent 4":
					return "{D27102A9-8310-4765-A935-A1911B00CA55}";
				case "Light Style 1 - Accent 5":
					return "{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}";
				case "Light Style 1 - Accent 6":
					return "{68D230F3-CF80-4859-8CE7-A43EE81993B5}";
				case "Light Style 2":
					return "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}";
				case "Light Style 2 - Accent 1":
					return "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}";
				case "Light Style 2 - Accent 2":
					return "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}";
				case "Light Style 2 - Accent 3":
					return "{F2DE63D5-997A-4646-A377-4702673A728D}";
				case "Light Style 2 - Accent 4":
					return "{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}";
				case "Light Style 2 - Accent 5":
					return "{5A111915-BE36-4E01-A7E5-04B1672EAD32}";
				case "Light Style 2 - Accent 6":
					return "{912C8C85-51F0-491E-9774-3900AFEF0FD7}";
				case "Light Style 3":
					return "{616DA210-FB5B-4158-B5E0-FEB733F419BA}";
				case "Light Style 3 - Accent 1":
					return "{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}";
				case "Light Style 3 - Accent 2":
					return "{5DA37D80-6434-44D0-A028-1B22A696006F}";
				case "Light Style 3 - Accent 3":
					return "{8799B23B-EC83-4686-B30A-512413B5E67A}";
				case "Light Style 3 - Accent 4":
					return "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}";
				case "Light Style 3 - Accent 5":
					return "{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}";
				case "Light Style 3 - Accent 6":
					return "{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}";
				case "Medium Style 1":
					return "{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}";
				case "Medium Style 1 - Accent 1":
					return "{B301B821-A1FF-4177-AEE7-76D212191A09}";
				case "Medium Style 1 - Accent 2":
					return "{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}";
				case "Medium Style 1 - Accent 3":
					return "{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}";
				case "Medium Style 1 - Accent 4":
					return "{1E171933-4619-4E11-9A3F-F7608DF75F80}";
				case "Medium Style 1 - Accent 5":
					return "{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}";
				case "Medium Style 1 - Accent 6":
					return "{10A1B5D5-9B99-4C35-A422-299274C87663}";
				case "Medium Style 2":
					return "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}";
				case "Medium Style 2 - Accent 2":
					return "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}";
				case "Medium Style 2 - Accent 3":
					return "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}";
				case "Medium Style 2 - Accent 4":
					return "{00A15C55-8517-42AA-B614-E9B94910E393}";
				case "Medium Style 2 - Accent 5":
					return "{7DF18680-E054-41AD-8BC1-D1AEF772440D}";
				case "Medium Style 2 - Accent 6":
					return "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}";
				case "Medium Style 3":
					return "{8EC20E35-A176-4012-BC5E-935CFFF8708E}";
				case "Medium Style 3 - Accent 1":
					return "{6E25E649-3F16-4E02-A733-19D2CDBF48F0}";
				case "Medium Style 3 - Accent 2":
					return "{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}";
				case "Medium Style 3 - Accent 3":
					return "{EB344D84-9AFB-497E-A393-DC336BA19D2E}";
				case "Medium Style 3 - Accent 4":
					return "{EB9631B5-78F2-41C9-869B-9F39066F8104}";
				case "Medium Style 3 - Accent 5":
					return "{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}";
				case "Medium Style 3 - Accent 6":
					return "{2A488322-F2BA-4B5B-9748-0D474271808F}";
				case "Medium Style 4":
					return "{D7AC3CCA-C797-4891-BE02-D94E43425B78}";
				case "Medium Style 4 - Accent 1":
					return "{69CF1AB2-1976-4502-BF36-3FF5EA218861}";
				case "Medium Style 4 - Accent 2":
					return "{8A107856-5554-42FB-B03E-39F5DBC370BA}";
				case "Medium Style 4 - Accent 3":
					return "{0505E3EF-67EA-436B-97B2-0124C06EBD24}";
				case "Medium Style 4 - Accent 4":
					return "{C4B1156A-380E-4F78-BDF5-A606A8083BF9}";
				case "Medium Style 4 - Accent 5":
					return "{22838BEF-8BB2-4498-84A7-C5851F593DF1}";
				case "Medium Style 4 - Accent 6":
					return "{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}";
				case "Dark Style 1":
					return "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}";
				case "Dark Style 1 - Accent 1":
					return "{125E5076-3810-47DD-B79F-674D7AD40C01}";
				case "Dark Style 1 - Accent 2":
					return "{37CE84F3-28C3-443E-9E96-99CF82512B78}";
				case "Dark Style 1 - Accent 3":
					return "{D03447BB-5D67-496B-8E87-E561075AD55C}";
				case "Dark Style 1 - Accent 4":
					return "{E929F9F4-4A8F-4326-A1B4-22849713DDAB}";
				case "Dark Style 1 - Accent 5":
					return "{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}";
				case "Dark Style 1 - Accent 6":
					return "{AF606853-7671-496A-8E4F-DF71F8EC918B}";
				case "Dark Style 2":
					return "{5202B0CA-FC54-4496-8BCA-5EF66A818D29}";
				case "Dark Style 2 - Accent 1/Accent 2":
					return "{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}";
				case "Dark Style 2 - Accent 3/Accent 4":
					return "{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}";
				case "Dark Style 2 - Accent 5/Accent 6":
					return "{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}";
				default:
					return AscCommon.CreateGUID();
			}
		}

		const OBJECT_MORPH_MARKER = "!!";

		function IsEqualObjects(obj1, obj2) {
			let iO = AscCommon.isRealObject;
			let iN = AscFormat.isRealNumber;
			let iEN = AscFormat.fApproxEqual;
			if(!iO(obj1) && iO(obj2) || iO(obj1) && !iO(obj2)) return false;
			for(let sKey in obj1) {
				if(obj1.hasOwnProperty(sKey)) {
					let pr1 = obj1[sKey];
					let pr2 = obj2[sKey];
					if(iO(pr1)) {
						if(!IsEqualObjects(pr1, pr2)) return false;
					}
					else if(iN(pr1)) {
						if(!iEN(pr1, pr2)) return false;
					}
					else {
						if(pr1 !== pr2) return false;
					}
				}
			}
			return true;
		}
//----------------------------------------------------------export----------------------------------------------------
		window['AscFormat'] = window['AscFormat'] || {};
		window['AscFormat'].CreateFontRef = CreateFontRef;
		window['AscFormat'].CreatePresetColor = CreatePresetColor;
		window['AscFormat'].isRealNumber = isRealNumber;
		window['AscFormat'].isRealBool = isRealBool;
		window['AscFormat'].groupBy = groupBy;
		window['AscFormat'].writeLong = writeLong;
		window['AscFormat'].readLong = readLong;
		window['AscFormat'].writeDouble = writeDouble;
		window['AscFormat'].readDouble = readDouble;
		window['AscFormat'].writeDouble2 = writeDouble2;
		window['AscFormat'].readDouble2 = readDouble2;
		window['AscFormat'].writeBool = writeBool;
		window['AscFormat'].readBool = readBool;
		window['AscFormat'].writeString = writeString;
		window['AscFormat'].readString = readString;
		window['AscFormat'].writeObject = writeObject;
		window['AscFormat'].readObject = readObject;
		window['AscFormat'].readObjectNoId = readObjectNoId;
		window['AscFormat'].writeObjectNoId = writeObjectNoId;
		window['AscFormat'].readObjectNoIdNoType = readObjectNoIdNoType;
		window['AscFormat'].writeObjectNoIdNoType = writeObjectNoIdNoType;
		window['AscFormat'].checkThemeFonts = checkThemeFonts;
		window['AscFormat'].ExecuteNoHistory = ExecuteNoHistory;
		window['AscFormat'].CreatePresentationTableStyles = CreatePresentationTableStyles;
		window['AscFormat'].getDefaultGUIDTableStyleByName = getDefaultGUIDTableStyleByName;
		window['AscFormat'].checkTableCellPr = checkTableCellPr;
		window['AscFormat'].CColorMod = CColorMod;
		window['AscFormat'].CColorModifiers = CColorModifiers;
		window['AscFormat'].CBaseColor = CBaseColor;
		window['AscFormat'].CSysColor = CSysColor;
		window['AscFormat'].CPrstColor = CPrstColor;
		window['AscFormat'].CRGBColor = CRGBColor;
		window['AscFormat'].CSchemeColor = CSchemeColor;
		window['AscFormat'].CStyleColor = CStyleColor;
		window['AscFormat'].CUniColor = CUniColor;
		window['AscFormat'].CreateUniColorRGB = CreateUniColorRGB;
		window['AscFormat'].CreateUniColorRGB2 = CreateUniColorRGB2;
		window['AscFormat'].CreateSolidFillRGB = CreateSolidFillRGB;
		window['AscFormat'].CreateSolidFillRGBA = CreateSolidFillRGBA;
		window['AscFormat'].getGrayscaleValue = getGrayscaleValue;
		window['AscFormat'].CSrcRect = CSrcRect;
		window['AscFormat'].CBlipFillTile = CBlipFillTile;
		window['AscFormat'].CBlipFill = CBlipFill;
		window['AscFormat'].CBlip = CBlip;
		window['AscFormat'].CSolidFill = CSolidFill;
		window['AscFormat'].CGs = CGs;
		window['AscFormat'].GradLin = GradLin;
		window['AscFormat'].GradPath = GradPath;
		window['AscFormat'].CGradFill = CGradFill;
		window['AscFormat'].CPattFill = CPattFill;
		window['AscFormat'].CNoFill = CNoFill;
		window['AscFormat'].CGrpFill = CGrpFill;
		window['AscFormat'].CUniFill = CUniFill;
		window['AscFormat'].CompareUniFill = CompareUniFill;
		window['AscFormat'].CompareUnifillBool = CompareUnifillBool;
		window['AscFormat'].CompareShapeProperties = CompareShapeProperties;
		window['AscFormat'].CompareProtectionFlags = CompareProtectionFlags;
		window['AscFormat'].EndArrow = EndArrow;
		window['AscFormat'].ConvertJoinAggType = ConvertJoinAggType;
		window['AscFormat'].LineJoin = LineJoin;
		window['AscFormat'].CLn = CLn;
		window['AscFormat'].DefaultShapeDefinition = DefaultShapeDefinition;
		window['AscFormat'].CNvPr = CNvPr;
		window['AscFormat'].NvPr = NvPr;
		window['AscFormat'].Ph = Ph;
		window['AscFormat'].UniNvPr = UniNvPr;
		window['AscFormat'].StyleRef = StyleRef;
		window['AscFormat'].FontRef = FontRef;
		window['AscFormat'].CShapeStyle = CShapeStyle;
		window['AscFormat'].CreateDefaultShapeStyle = CreateDefaultShapeStyle;
		window['AscFormat'].CXfrm = CXfrm;
		window['AscFormat'].CEffectProperties = CEffectProperties;
		window['AscFormat'].CEffectLst = CEffectLst;
		window['AscFormat'].CSpPr = CSpPr;
		window['AscFormat'].ClrScheme = ClrScheme;
		window['AscFormat'].ClrMap = ClrMap;
		window['AscFormat'].ExtraClrScheme = ExtraClrScheme;
		window['AscFormat'].FontCollection = FontCollection;
		window['AscFormat'].FontScheme = FontScheme;
		window['AscFormat'].FmtScheme = FmtScheme;
		window['AscFormat'].ThemeElements = ThemeElements;
		window['AscFormat'].CTheme = CTheme;
		window['AscFormat'].HF = HF;
		window['AscFormat'].CBgPr = CBgPr;
		window['AscFormat'].CBg = CBg;
		window['AscFormat'].CSld = CSld;
		window['AscFormat'].CTextStyles = CTextStyles;
		window['AscFormat'].redrawSlide = redrawSlide;
		window['AscFormat'].CTextFit = CTextFit;
		window['AscFormat'].CBodyPr = CBodyPr;
		window['AscFormat'].CTextParagraphPr = CTextParagraphPr;
		window['AscFormat'].CompareBullets = CompareBullets;
		window['AscFormat'].CBuBlip = CBuBlip;
		window['AscFormat'].CBullet = CBullet;
		window['AscFormat'].CBulletColor = CBulletColor;
		window['AscFormat'].CBulletSize = CBulletSize;
		window['AscFormat'].CBulletTypeface = CBulletTypeface;
		window['AscFormat'].CBulletType = CBulletType;
		window['AscFormat'].TextListStyle = TextListStyle;
		window['AscFormat'].GenerateDefaultTheme = GenerateDefaultTheme;
		window['AscFormat'].GenerateDefaultMasterSlide = GenerateDefaultMasterSlide;
		window['AscFormat'].GenerateDefaultSlideLayout = GenerateDefaultSlideLayout;
		window['AscFormat'].GenerateDefaultSlide = GenerateDefaultSlide;
		window['AscFormat'].CreateDefaultTextRectStyle = CreateDefaultTextRectStyle;
		window['AscFormat'].GenerateDefaultColorMap = GenerateDefaultColorMap;
		window['AscFormat'].CreateAscFill = CreateAscFill;
		window['AscFormat'].CorrectUniFill = CorrectUniFill;
		window['AscFormat'].CreateAscStroke = CreateAscStroke;
		window['AscFormat'].CorrectUniStroke = CorrectUniStroke;
		window['AscFormat'].CreateAscShapePropFromProp = CreateAscShapePropFromProp;
		window['AscFormat'].CreateAscTextArtProps = CreateAscTextArtProps;
		window['AscFormat'].CreateUnifillFromAscColor = CreateUnifillFromAscColor;
		window['AscFormat'].CorrectUniColor = CorrectUniColor;
		window['AscFormat'].deleteDrawingBase = deleteDrawingBase;
		window['AscFormat'].CNvUniSpPr = CNvUniSpPr;
		window['AscFormat'].UniMedia = UniMedia;
		window['AscFormat'].CT_Hyperlink = CT_Hyperlink;


		window['AscFormat'].builder_CreateShape = builder_CreateShape;
		window['AscFormat'].builder_CreateChart = builder_CreateChart;
		window['AscFormat'].builder_CreateGroup = builder_CreateGroup;
		window['AscFormat'].builder_CreateSchemeColor = builder_CreateSchemeColor;
		window['AscFormat'].builder_CreatePresetColor = builder_CreatePresetColor;
		window['AscFormat'].builder_CreateGradientStop = builder_CreateGradientStop;
		window['AscFormat'].builder_CreateLinearGradient = builder_CreateLinearGradient;
		window['AscFormat'].builder_CreateRadialGradient = builder_CreateRadialGradient;
		window['AscFormat'].builder_CreatePatternFill = builder_CreatePatternFill;
		window['AscFormat'].builder_CreateBlipFill = builder_CreateBlipFill;
		window['AscFormat'].builder_CreateLine = builder_CreateLine;
		window['AscFormat'].builder_SetChartTitle = builder_SetChartTitle;
		window['AscFormat'].builder_SetChartHorAxisTitle = builder_SetChartHorAxisTitle;
		window['AscFormat'].builder_SetChartVertAxisTitle = builder_SetChartVertAxisTitle;
		window['AscFormat'].builder_SetChartLegendPos = builder_SetChartLegendPos;
		window['AscFormat'].builder_SetShowDataLabels = builder_SetShowDataLabels;

		window['AscFormat'].builder_SetChartVertAxisOrientation = builder_SetChartVertAxisOrientation;
		window['AscFormat'].builder_SetChartHorAxisOrientation = builder_SetChartHorAxisOrientation;
		window['AscFormat'].builder_SetChartVertAxisTickLablePosition = builder_SetChartVertAxisTickLablePosition;
		window['AscFormat'].builder_SetChartHorAxisTickLablePosition = builder_SetChartHorAxisTickLablePosition;

		window['AscFormat'].builder_SetChartHorAxisMajorTickMark = builder_SetChartHorAxisMajorTickMark;
		window['AscFormat'].builder_SetChartHorAxisMinorTickMark = builder_SetChartHorAxisMinorTickMark;
		window['AscFormat'].builder_SetChartVerAxisMajorTickMark = builder_SetChartVerAxisMajorTickMark;
		window['AscFormat'].builder_SetChartVerAxisMinorTickMark = builder_SetChartVerAxisMinorTickMark;
		window['AscFormat'].builder_SetLegendFontSize = builder_SetLegendFontSize;
		window['AscFormat'].builder_SetHorAxisMajorGridlines = builder_SetHorAxisMajorGridlines;
		window['AscFormat'].builder_SetHorAxisMinorGridlines = builder_SetHorAxisMinorGridlines;
		window['AscFormat'].builder_SetVerAxisMajorGridlines = builder_SetVerAxisMajorGridlines;
		window['AscFormat'].builder_SetVerAxisMinorGridlines = builder_SetVerAxisMinorGridlines;
		window['AscFormat'].builder_SetHorAxisFontSize = builder_SetHorAxisFontSize;
		window['AscFormat'].builder_SetVerAxisFontSize = builder_SetVerAxisFontSize;
		window['AscFormat'].builder_SetShowPointDataLabel = builder_SetShowPointDataLabel;


		window['AscFormat'].Ax_Counter = Ax_Counter;
		window['AscFormat'].TYPE_TRACK = TYPE_TRACK;
		window['AscFormat'].TYPE_KIND = TYPE_KIND;
		window['AscFormat'].mapPrstColor = map_prst_color;
		window['AscFormat'].map_hightlight = map_hightlight;
		window['AscFormat'].ar_arrow = ar_arrow;
		window['AscFormat'].ar_diamond = ar_diamond;
		window['AscFormat'].ar_none = ar_none;
		window['AscFormat'].ar_oval = ar_oval;
		window['AscFormat'].ar_stealth = ar_stealth;
		window['AscFormat'].ar_triangle = ar_triangle;
		window['AscFormat'].LineEndType = LineEndType;
		window['AscFormat'].LineEndSize = LineEndSize;
		window['AscFormat'].LineJoinType = LineJoinType;

//типы плейсхолдеров
		window['AscFormat']["phType_body"] = window['AscFormat'].phType_body = 0;
		window['AscFormat']["phType_chart"] = window['AscFormat'].phType_chart = 1;
		window['AscFormat']["phType_clipArt"] = window['AscFormat'].phType_clipArt = 2;
		window['AscFormat']["phType_ctrTitle"] = window['AscFormat'].phType_ctrTitle = 3;
		window['AscFormat']["phType_dgm"] = window['AscFormat'].phType_dgm = 4;
		window['AscFormat']["phType_dt"] = window['AscFormat'].phType_dt = 5;
		window['AscFormat']["phType_ftr"] = window['AscFormat'].phType_ftr = 6;
		window['AscFormat']["phType_hdr"] = window['AscFormat'].phType_hdr = 7;
		window['AscFormat']["phType_media"] = window['AscFormat'].phType_media = 8;
		window['AscFormat']["phType_obj"] = window['AscFormat'].phType_obj = 9;
		window['AscFormat']["phType_pic"] = window['AscFormat'].phType_pic = 10;
		window['AscFormat']["phType_sldImg"] = window['AscFormat'].phType_sldImg = 11;
		window['AscFormat']["phType_sldNum"] = window['AscFormat'].phType_sldNum = 12;
		window['AscFormat']["phType_subTitle"] = window['AscFormat'].phType_subTitle = 13;
		window['AscFormat']["phType_tbl"] = window['AscFormat'].phType_tbl = 14;
		window['AscFormat']["phType_title"] = window['AscFormat'].phType_title = 15;

		window['AscFormat'].fntStyleInd_none = 2;
		window['AscFormat'].fntStyleInd_major = 0;
		window['AscFormat'].fntStyleInd_minor = 1;

		window['AscFormat'].VERTICAL_ANCHOR_TYPE_BOTTOM = 0;
		window['AscFormat'].VERTICAL_ANCHOR_TYPE_CENTER = 1;
		window['AscFormat'].VERTICAL_ANCHOR_TYPE_DISTRIBUTED = 2;
		window['AscFormat'].VERTICAL_ANCHOR_TYPE_JUSTIFIED = 3;
		window['AscFormat'].VERTICAL_ANCHOR_TYPE_TOP = 4;

//Vertical Text Types
		window['AscFormat'].nVertTTeaVert = 0; //( ( East Asian Vertical ))
		window['AscFormat'].nVertTThorz = 1; //( ( Horizontal ))
		window['AscFormat'].nVertTTmongolianVert = 2; //( ( Mongolian Vertical ))
		window['AscFormat'].nVertTTvert = 3; //( ( Vertical ))
		window['AscFormat'].nVertTTvert270 = 4;//( ( Vertical 270 ))
		window['AscFormat'].nVertTTwordArtVert = 5;//( ( WordArt Vertical ))
		window['AscFormat'].nVertTTwordArtVertRtl = 6;//(Vertical WordArt Right to Left)

//Text Wrapping Types
		window['AscFormat'].nTWTNone = 0;
		window['AscFormat'].nTWTSquare = 1;

		window['AscFormat'].BULLET_TYPE_COLOR_NONE = 0;
		window['AscFormat'].BULLET_TYPE_COLOR_CLRTX = 1;
		window['AscFormat'].BULLET_TYPE_COLOR_CLR = 2;

		window['AscFormat'].BULLET_TYPE_SIZE_NONE = 0;
		window['AscFormat'].BULLET_TYPE_SIZE_TX = 1;
		window['AscFormat'].BULLET_TYPE_SIZE_PCT = 2;
		window['AscFormat'].BULLET_TYPE_SIZE_PTS = 3;

		window['AscFormat'].BULLET_TYPE_TYPEFACE_NONE = 0;
		window['AscFormat'].BULLET_TYPE_TYPEFACE_TX = 1;
		window['AscFormat'].BULLET_TYPE_TYPEFACE_BUFONT = 2;

		window['AscFormat'].PARRUN_TYPE_NONE = 0;
		window['AscFormat'].PARRUN_TYPE_RUN = 1;
		window['AscFormat'].PARRUN_TYPE_FLD = 2;
		window['AscFormat'].PARRUN_TYPE_BR = 3;
		window['AscFormat'].PARRUN_TYPE_MATH = 4;
		window['AscFormat'].PARRUN_TYPE_MATHPARA = 5;

		window['AscFormat']._weight_body = _weight_body;
		window['AscFormat']._weight_chart = _weight_chart;
		window['AscFormat']._weight_clipArt = _weight_clipArt;
		window['AscFormat']._weight_ctrTitle = _weight_ctrTitle;
		window['AscFormat']._weight_dgm = _weight_dgm;
		window['AscFormat']._weight_media = _weight_media;
		window['AscFormat']._weight_obj = _weight_obj;
		window['AscFormat']._weight_pic = _weight_pic;
		window['AscFormat']._weight_subTitle = _weight_subTitle;
		window['AscFormat']._weight_tbl = _weight_tbl;
		window['AscFormat']._weight_title = _weight_title;

		window['AscFormat']._ph_multiplier = _ph_multiplier;

		window['AscFormat'].nSldLtTTitle = nSldLtTTitle;
		window['AscFormat'].nSldLtTObj = nSldLtTObj;
		window['AscFormat'].nSldLtTTx = nSldLtTTx;

		window['AscFormat']._arr_lt_types_weight = _arr_lt_types_weight;
		window['AscFormat']._global_layout_summs_array = _global_layout_summs_array;


		window['AscFormat'].AUDIO_CD = AUDIO_CD;
		window['AscFormat'].WAV_AUDIO_FILE = WAV_AUDIO_FILE;
		window['AscFormat'].AUDIO_FILE = AUDIO_FILE;
		window['AscFormat'].VIDEO_FILE = VIDEO_FILE;
		window['AscFormat'].QUICK_TIME_FILE = QUICK_TIME_FILE;




		window['AscFormat'].DRAW_TYPE_PEN = DRAW_TYPE_PEN;
		window['AscFormat'].DRAW_TYPE_PENCIL = DRAW_TYPE_PENCIL;
		window['AscFormat'].DRAW_TYPE_HIGHLITER = DRAW_TYPE_HIGHLITER;

		window['AscFormat'].EFFECT_TYPE_NONE = EFFECT_TYPE_NONE;
		window['AscFormat'].EFFECT_TYPE_OUTERSHDW = EFFECT_TYPE_OUTERSHDW;
		window['AscFormat'].EFFECT_TYPE_GLOW = EFFECT_TYPE_GLOW;
		window['AscFormat'].EFFECT_TYPE_DUOTONE = EFFECT_TYPE_DUOTONE;
		window['AscFormat'].EFFECT_TYPE_XFRM = EFFECT_TYPE_XFRM;
		window['AscFormat'].EFFECT_TYPE_BLUR = EFFECT_TYPE_BLUR;
		window['AscFormat'].EFFECT_TYPE_PRSTSHDW = EFFECT_TYPE_PRSTSHDW;
		window['AscFormat'].EFFECT_TYPE_INNERSHDW = EFFECT_TYPE_INNERSHDW;
		window['AscFormat'].EFFECT_TYPE_REFLECTION = EFFECT_TYPE_REFLECTION;
		window['AscFormat'].EFFECT_TYPE_SOFTEDGE = EFFECT_TYPE_SOFTEDGE;
		window['AscFormat'].EFFECT_TYPE_FILLOVERLAY = EFFECT_TYPE_FILLOVERLAY;
		window['AscFormat'].EFFECT_TYPE_ALPHACEILING = EFFECT_TYPE_ALPHACEILING;
		window['AscFormat'].EFFECT_TYPE_ALPHAFLOOR = EFFECT_TYPE_ALPHAFLOOR;
		window['AscFormat'].EFFECT_TYPE_TINTEFFECT = EFFECT_TYPE_TINTEFFECT;
		window['AscFormat'].EFFECT_TYPE_RELOFF = EFFECT_TYPE_RELOFF;
		window['AscFormat'].EFFECT_TYPE_LUM = EFFECT_TYPE_LUM;
		window['AscFormat'].EFFECT_TYPE_HSL = EFFECT_TYPE_HSL;
		window['AscFormat'].EFFECT_TYPE_GRAYSCL = EFFECT_TYPE_GRAYSCL;
		window['AscFormat'].EFFECT_TYPE_ELEMENT = EFFECT_TYPE_ELEMENT;
		window['AscFormat'].EFFECT_TYPE_ALPHAREPL = EFFECT_TYPE_ALPHAREPL;
		window['AscFormat'].EFFECT_TYPE_ALPHAOUTSET = EFFECT_TYPE_ALPHAOUTSET;
		window['AscFormat'].EFFECT_TYPE_ALPHAMODFIX = EFFECT_TYPE_ALPHAMODFIX;
		window['AscFormat'].EFFECT_TYPE_ALPHABILEVEL = EFFECT_TYPE_ALPHABILEVEL;
		window['AscFormat'].EFFECT_TYPE_BILEVEL = EFFECT_TYPE_BILEVEL;
		window['AscFormat'].EFFECT_TYPE_DAG = EFFECT_TYPE_DAG;
		window['AscFormat'].EFFECT_TYPE_FILL = EFFECT_TYPE_FILL;
		window['AscFormat'].EFFECT_TYPE_CLRREPL = EFFECT_TYPE_CLRREPL;
		window['AscFormat'].EFFECT_TYPE_CLRCHANGE = EFFECT_TYPE_CLRCHANGE;
		window['AscFormat'].EFFECT_TYPE_ALPHAINV = EFFECT_TYPE_ALPHAINV;
		window['AscFormat'].EFFECT_TYPE_ALPHAMOD = EFFECT_TYPE_ALPHAMOD;
		window['AscFormat'].EFFECT_TYPE_BLEND = EFFECT_TYPE_BLEND;


		window['AscFormat'].fCreateEffectByType = fCreateEffectByType;
		window['AscFormat'].COuterShdw = COuterShdw;
		window['AscFormat'].CGlow = CGlow;
		window['AscFormat'].CDuotone = CDuotone;
		window['AscFormat'].CXfrmEffect = CXfrmEffect;
		window['AscFormat'].CBlur = CBlur;
		window['AscFormat'].CPrstShdw = CPrstShdw;
		window['AscFormat'].CInnerShdw = CInnerShdw;
		window['AscFormat'].CReflection = CReflection;
		window['AscFormat'].CSoftEdge = CSoftEdge;
		window['AscFormat'].CFillOverlay = CFillOverlay;
		window['AscFormat'].CAlphaCeiling = CAlphaCeiling;
		window['AscFormat'].CAlphaFloor = CAlphaFloor;
		window['AscFormat'].CTintEffect = CTintEffect;
		window['AscFormat'].CRelOff = CRelOff;
		window['AscFormat'].CLumEffect = CLumEffect;
		window['AscFormat'].CHslEffect = CHslEffect;
		window['AscFormat'].CGrayscl = CGrayscl;
		window['AscFormat'].CEffectElement = CEffectElement;
		window['AscFormat'].CAlphaRepl = CAlphaRepl;
		window['AscFormat'].CAlphaOutset = CAlphaOutset;
		window['AscFormat'].CAlphaModFix = CAlphaModFix;
		window['AscFormat'].CAlphaBiLevel = CAlphaBiLevel;
		window['AscFormat'].CBiLevel = CBiLevel;
		window['AscFormat'].CEffectContainer = CEffectContainer;
		window['AscFormat'].CFillEffect = CFillEffect;
		window['AscFormat'].CClrRepl = CClrRepl;
		window['AscFormat'].CClrChange = CClrChange;
		window['AscFormat'].CAlphaInv = CAlphaInv;
		window['AscFormat'].CAlphaMod = CAlphaMod;
		window['AscFormat'].CBlend = CBlend;
		window['AscFormat'].CreateNoneBullet = CreateNoneBullet;
		window['AscFormat'].ChartBuilderTypeToInternal = ChartBuilderTypeToInternal;
		window['AscFormat'].InitClass = InitClass;
		window['AscFormat'].InitClassWithoutType = InitClassWithoutType;
		window['AscFormat'].CBaseObject = CBaseObject;
		window['AscFormat'].CBaseFormatObject = CBaseFormatObject;
		window['AscFormat'].CBaseFormatNoIdObject = CBaseFormatNoIdObject;
		window['AscFormat'].CBaseNoIdObject = CBaseNoIdObject;
		window['AscFormat'].checkRasterImageId = checkRasterImageId;
		window['AscFormat'].IdEntry = IdEntry;
		window['AscFormat'].CEffectStyle = CEffectStyle;

		window['AscFormat'].DEFAULT_COLOR_MAP = null;
		window['AscFormat'].DEFAULT_THEME = null;
		window['AscFormat'].GetDefaultColorMap = GetDefaultColorMap;
		window['AscFormat'].GetDefaultTheme = GetDefaultTheme;
		window['AscFormat'].getPercentageValue = getPercentageValue;
		window['AscFormat'].getPercentageValueForWrite = getPercentageValueForWrite;
		window['AscFormat'].CSpTree = CSpTree;
		window['AscFormat'].CClrMapOvr = CClrMapOvr;
		window['AscFormat'].CBaseAttrObject = CBaseAttrObject;
		window['AscFormat'].PartTitle = PartTitle;
		window['AscFormat'].CCustomProperty = CCustomProperty;
		window['AscFormat'].CVariantVector = CVariantVector;
		window['AscFormat'].CVariantArray = CVariantArray;
		window['AscFormat'].CVariantVStream = CVariantVStream;
		window['AscFormat'].fRGBAToHexString = fRGBAToHexString;
		window['AscFormat'].RefreshContentAllFields = RefreshContentAllFields;
		window['AscFormat'].IsEqualObjects = IsEqualObjects;
		window['AscFormat'].CreateSchemeUnicolorWithMods = CreateSchemeUnicolorWithMods;
		window['AscFormat'].szPh_full = szPh_full;
		window['AscFormat'].szPh_half = szPh_half;
		window['AscFormat'].szPh_quarter = szPh_quarter;
		window['AscFormat'].orientPh_horz = orientPh_horz;
		window['AscFormat'].orientPh_vert = orientPh_vert;
		window['AscFormat'].effectcontainertypeSib = effectcontainertypeSib;
		window['AscFormat'].effectcontainertypeTree = effectcontainertypeTree;
		window['AscFormat'].CLR_IDX_MAP = CLR_IDX_MAP;
		window['AscFormat'].MAP_AUTONUM_TYPES = MAP_AUTONUM_TYPES;
		window['AscFormat'].CLR_NAME_MAP = CLR_NAME_MAP;
		window['AscFormat'].LINE_PRESETS_MAP = LINE_PRESETS_MAP;
		window['AscFormat'].OBJECT_MORPH_MARKER = OBJECT_MORPH_MARKER;

		// Visio extensions export
		window['AscFormat'].CClrSchemeExtLst = CClrSchemeExtLst;
		window['AscFormat'].CVariationClrScheme = CVariationClrScheme;
		window['AscFormat'].CVarColor = CVarColor;
		window['AscFormat'].CThemeExt = CThemeExt;
		window['AscFormat'].CVariationStyleScheme = CVariationStyleScheme;
		window['AscFormat'].CVarStyle = CVarStyle;
		window['AscFormat'].CFontProps = CFontProps;
		window['AscFormat'].CLineStyle = CLineStyle;

		window["AscFormat"].RECT_ALIGN_B = RECT_ALIGN_B;
		window["AscFormat"].RECT_ALIGN_BL = RECT_ALIGN_BL;
		window["AscFormat"].RECT_ALIGN_BR = RECT_ALIGN_BR;
		window["AscFormat"].RECT_ALIGN_CTR = RECT_ALIGN_CTR;
		window["AscFormat"].RECT_ALIGN_L = RECT_ALIGN_L;
		window["AscFormat"].RECT_ALIGN_R = RECT_ALIGN_R;
		window["AscFormat"].RECT_ALIGN_T = RECT_ALIGN_T;
		window["AscFormat"].RECT_ALIGN_TL = RECT_ALIGN_TL;
		window["AscFormat"].RECT_ALIGN_TR = RECT_ALIGN_TR;
	})
(window);

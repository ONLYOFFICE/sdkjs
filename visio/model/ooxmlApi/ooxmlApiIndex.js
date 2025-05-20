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

(function(window, document)
{
	var c_oVsdxSheetStorageKind = {
		Cell_Type : 0,
		Trigger_Type : 1,
		Row_Type : 2,
		Section_Type : 3,
		Text_Type : 4,
		Data_Type : 5,
		ForeignData_Type : 6
	};
	/**
	 *    // Docs old:
	 * // Text_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/text_type-complextypevisio-xml
	 * @returns {Text_Type}
	 * @constructor
	 */
	function Text_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);

		/**
		 * if text is inherited (for calculate presentation field)
		 */
		this.isInherited = false;

		this.elements = []
		// array to store elems below. see ShapeSheet element
		// this.cp = [];
		// this.pp = [];
		// this.tp = [];
		// this.fld = [];

		// but we can have text among elements

		// <Text><cp IX='0'/><pp IX='0'/>Page <fld IX='0'>21</fld>\r\n</Text>
		// <Text><cp IX='0'/><fld IX='0'>2/10</fld>\r\n</Text>
		// <Text><cp IX='0'/><pp IX='0'/><fld IX='0'>3/10/2013</fld> - <fld IX='1'>3/17/2013</fld>\r\n</Text>

		// notice \r\n. \r is CR symbol and \n is LF symbol. So \r\n gives us line drop used in .xml files we work with.
		// \r\n only happens after xml declaration and in text. So in text it is important, it is a part of text.
		// here is an example of 1 2 3 4 5 each on new line
		// <Text><cp IX='0'/>1\r\n2\r\n3\r\n4\r\n5\r\n\r\n</Text>
		// if you want to see \r\n set proper settings in your editor
		return this;
	}
	AscFormat.InitClass(Text_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);
	Text_Type.prototype.kind = c_oVsdxSheetStorageKind.Text_Type;

	/**
	 * // Docs old:
	 * @returns {Data_Type}
	 * @constructor
	 */
	function Data_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);

		this.value = null;
		// to serialize in function writeShapeSheetElementsXml
		this.tagName = null;
		return this;
	}
	AscFormat.InitClass(Data_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);
	Data_Type.prototype.kind = c_oVsdxSheetStorageKind.Data_Type;

	/**
	 *    // Docs old:
	 * // Элемент Rel (ForeignData_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/rel-element-foreigndata_type-complextypevisio-xml
	 * // ForeignData_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/foreigndata_type-complextypevisio-xml
	 * @returns {ForeignData_Type}
	 * @constructor
	 */
	function ForeignData_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);

		this.foreignType = null;
		this.objectType = null;
		this.showAsIcon = null;
		this.objectWidth = null;
		this.objectHeight = null;
		this.mappingMode = null;
		this.extentX = null;
		this.extentY = null;
		this.compressionType = null;
		this.compressionLevel = null;
		this.rel = null;
		this.mediaFilename = null;
		this.oleFilename = null;
		return this;
	}
	AscFormat.InitClass(ForeignData_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);
	ForeignData_Type.prototype.kind = c_oVsdxSheetStorageKind.ForeignData_Type;


	/**
	 *    // https://learn.microsoft.com/ru-ru/office/client-developer/visio/trigger_type-complextypevisio-xml
	 * @constructor
	 */
	function Trigger_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);

		this.refBy = [];
		this.n = null;
	}
	AscFormat.InitClass(Trigger_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);
	Trigger_Type.prototype.kind = c_oVsdxSheetStorageKind.Trigger_Type;
	/**
	 * Abstract class. For all Cell containers: ShapeSheet_Type (Sheet_Type) descendents and
	 * sections and rows also.
	 * @constructor
	 */
	function SheetStorage() {
		AscFormat.CBaseFormatNoIdObject.call(this);


		// setSheetClassMembers

		/**
		 * When working with shapes use getElements/setElements methods
		 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
		 * elements is used bcs text can appear here maybe
		 * @type {{}}
		 */
		this.elements = {};

		/**
		 * @type {{}}
		 */
		this.inheritedElements = {};

		// elements below are stored in elements to support new schema

		// // 3 arrays below inherited from Sheet_Type
		// this.cells = [];
		// this.triggers = [];
		// this.sections = [];
		// also rows

		// // new attributes inherited from ShapeSheet_Type
		// this.text = null;
		// this.data1 = null;
		// this.data2 = null;
		// this.data3 = null;
		// this.foreignData = null;
	}
	AscFormat.InitClass(SheetStorage, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	/**
	 * clone master elements (sections, rows, cells) to shapeElements.
	 * For Sections and Rows merge is recursive: we compare inner cells by their names
	 * @memberof SheetStorage
	 * @param masterElements - cells rows sections
	 * @param {string[]?} elementsToMerge - cells rows sections list we can merge
	 * @param {boolean?} isParentInList
	 */
	SheetStorage.prototype.mergeElementArrays = function mergeElementArrays(masterElements, elementsToMerge, isParentInList) {
		/**
		 * find index of cell row or section
		 * @param {SheetStorage} elementsObject
		 * @param elementToFind
		 * @returns {*}
		 */
		function findObjectIn(elementsObject, elementToFind) {
			let objKey = AscVisio.createKeyFromSheetObject(elementToFind);
			let element = elementsObject.getElement(objKey);
			return element;
		}

		/**
		 * if text is inherited so we consider that text fields in it have wrong values
		 * and we recalculate values them
		 * @param masterElement
		 */
		function setIsInheritedForText(masterElement) {
			if (masterElement.kind === c_oVsdxSheetStorageKind.Text_Type) {
				masterElement.isInherited = true;
			}
		}

		let mergeAll = false;

		if (elementsToMerge === undefined) {
			mergeAll = true;
		}

		for (const key in masterElements) {
			const masterElement = masterElements[key];

			let overrideObject = findObjectIn(this, masterElement);
			let elementExistsAlready = overrideObject !== undefined;

			let isElementInList = elementsToMerge !== undefined && elementsToMerge.includes(masterElement.n);
			let listCheck = mergeAll || isParentInList || isElementInList;

			if (!elementExistsAlready) {
				if (listCheck) {
					// TODO fix order
					// now Section sort is realized in getSections,
					// rowsSort is not needed see getRow findObject call

					// mb lets not add cell after section
					// let elementCopy = clone(masterElement);
					setIsInheritedForText(masterElement);

					let elementLink = masterElement;
					this.inheritedElements[key] = elementLink;
				}
			} else {
				// merge inner elements recursive if not cell
				if (masterElement.kind !== c_oVsdxSheetStorageKind.Cell_Type) {
					// if Section or Row
					let shapeElement = overrideObject;
					if (masterElement.kind === c_oVsdxSheetStorageKind.Section_Type || masterElement.kind === c_oVsdxSheetStorageKind.Row_Type) {
						// for future checks
						isParentInList = isElementInList || isParentInList;
						// recursive calls
						overrideObject.mergeElementArrays(masterElement.getElements(), elementsToMerge, isParentInList);
					}
				}
			}
		}
	}


	/**
	 * Abstract class for ShapeSheet_Type (Sheet_Type) descendents only.
	 * @constructor
	 * @extends SheetStorage
	 */
	function SheetStorageAndStyles() {
		// for ooxml classes

		// setShapeSheetClassMembers

		// 3 attr below inherited from Sheet_Type using old schema
		// or from ShapeSheet_Type using new schema
		this.lineStyle = null;
		this.fillStyle = null;
		this.textStyle = null;

		/**
		 * @type {string | number | null}
		 */
		this.inheritedLineStyle = null;
		/**
		 * @type {string | number | null}
		 */
		this.inheritedFillStyle = null;
		/**
		 * @type {string | number | null}
		 */
		this.inheritedTextStyle = null;

		// call parent class constructor
		let parentClassConstructor = SheetStorage;
		parentClassConstructor.call(this);
	}
	// inherit parent class methods
	// https://developer.mozilla.org/ru/docs/Web/JavaScript/Reference/Global_Objects/Object/create#%D0%BF%D1%80%D0%B8%D0%BC%D0%B5%D1%80_%D0%BA%D0%BB%D0%B0%D1%81%D1%81%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%BE%D0%B5_%D0%BD%D0%B0%D1%81%D0%BB%D0%B5%D0%B4%D0%BE%D0%B2%D0%B0%D0%BD%D0%B8%D0%B5_%D1%81_object.create
	SheetStorageAndStyles.prototype = Object.create(SheetStorage.prototype);
	SheetStorageAndStyles.prototype.constructor = SheetStorageAndStyles;

	/**
	 * @memberof SheetStorageAndStyles
	 * @return {*|string|number|null}
	 */
	SheetStorageAndStyles.prototype.getLineStyle = function getLineStyle() {
		if (this.lineStyle === null || this.lineStyle === undefined) {
			return this.inheritedLineStyle;
		}
		return this.lineStyle;
	}

	/**
	 * @memberof SheetStorageAndStyles
	 * @return {*|string|number|null}
	 */
	SheetStorageAndStyles.prototype.getFillStyle = function getFillStyle() {
		if (this.fillStyle === null || this.fillStyle === undefined) {
			return this.inheritedFillStyle;
		}
		return this.fillStyle;
	}

	/**
	 * @memberof SheetStorageAndStyles
	 * @return {*|string|number|null}
	 */
	SheetStorageAndStyles.prototype.getTextStyle = function getTextStyle() {
		if (this.textStyle === null || this.textStyle === undefined) {
			return this.inheritedTextStyle;
		}
		return this.textStyle;
	}


	// inheritance from ShapeSheetType for
	// StyleSheet_Type, DocumentSheet_Type, PageSheet_Type and Shape_Type
	// consider using with new schema and support old schema

	// 2.3.4.2.5	Cell_Type
	// N attribute MUST be unique amongst all of the Cell_Type elements of the containing Row_Type element

	// 2.3.4.2.81	Row_Type
	// If a Row_Type element specifies an N attribute, then it MUST NOT specify an IX attribute.
	// This attribute MUST be unique amongst all of the Row_Type elements of the containing Section_Type
	// element.

	// If a Row_Type element specifies an IX attribute, then it MUST NOT specify an N attribute.
	// The IX attribute of a Row_Type element MUST be unique amongst all of the Row_Type elements
	// of the containing Section_Type element.

	// 2.3.4.2.85	Section_Type
	// N attribute MUST be unique amongst all of the Section_Type elements of the containing Sheet_Type element
	// unless it is equal to "Geometry"!

	// IX attribute MUST be unique amongst all of the Section_Type elements with the same N attribute
	// of the containing Sheet_Type.

	// When the IX attribute is not present, the index of the element is calculated implicitly
	// by counting the number of  preceding Section_Type elements with the same N attribute in the containing
	// Sheet_Type.
	let lastIx = 0;
	function createKeyFromSheetObject(object) {
		let key;
		if (object.kind === c_oVsdxSheetStorageKind.Cell_Type) {
			key = object.n;
		} else if (object.kind === c_oVsdxSheetStorageKind.Row_Type) {
			if (object.n !== null) {
				key = object.n;
			} else if (object.ix !== null) {
				key = object.ix;
				lastIx = object.ix;
			} else {
				key = lastIx;
				lastIx++;
				AscCommon.consoleLog("Ix key is generated to store object", object);
			}
		} else if (object.kind === c_oVsdxSheetStorageKind.Section_Type)	{
			if (object.n === "Geometry") {
				key = "Geometry_" + object.ix;
			} else if (object.ix !== null) {
				key = object.ix;
			} else if (object.n !== null) {
				key = object.n;
			} else {
				AscCommon.consoleLog("Cant calculate key to store object", object);
			}
		} else if (object.kind === c_oVsdxSheetStorageKind.Text_Type) {
			key = "Text";
		} else if (object.kind === c_oVsdxSheetStorageKind.Data_Type) {
			key = object.tagName;
		} else if (object.kind === c_oVsdxSheetStorageKind.ForeignData_Type) {
			key = "ForeignData";
		} else if (object.kind === c_oVsdxSheetStorageKind.Trigger_Type) {
			key = object.n;
		} else {
			AscCommon.consoleLog("Unknown object in SheetElementsStorage", object);
		}

		return key;
	}

	/**
	 * for ooxml read
	 * @memberOf SheetStorage
	 * @param tagName
	 * @param reader
	 */
	SheetStorage.prototype.readInheritedElements = function readInheritedElements(tagName, reader) {
		let elem;
		switch (tagName) {
			case "Cell" : {
				elem = new Cell_Type();
				elem.fromXml(reader);
				let key = createKeyFromSheetObject(elem);
				this.elements[key] = elem;
				break;
			}
			case "Trigger" : {
				elem = new Trigger_Type();
				elem.fromXml(reader);
				let key = createKeyFromSheetObject(elem);
				this.elements[key] = elem;
				break;
			}
			case "Row" : {
				elem = new Row_Type();
				elem.fromXml(reader);
				let key = createKeyFromSheetObject(elem);
				this.elements[key] = elem;
				break;
			}
			case "Section" : {
				elem = new Section_Type();
				elem.fromXml(reader);
				let key = createKeyFromSheetObject(elem);
				this.elements[key] = elem;
				break;
			}
			case "Text" : {
				elem = new Text_Type();
				elem.fromXml(reader);
				let key = createKeyFromSheetObject(elem);
				this.elements[key] = elem;
				break;
			}
			case "Data1" : {
				elem = new Data_Type();
				elem.tagName = "Data1";
				elem.fromXml(reader);
				let key = createKeyFromSheetObject(elem);
				this.elements[key] = elem;
				break;
			}
			case "Data2" : {
				elem = new Data_Type();
				elem.tagName = "Data2";
				elem.fromXml(reader);
				let key = createKeyFromSheetObject(elem);
				this.elements[key] = elem;
				break;
			}
			case "Data3" : {
				elem = new Data_Type();
				elem.tagName = "Data3";
				elem.fromXml(reader);
				let key = createKeyFromSheetObject(elem);
				this.elements[key] = elem;
				break;
			}
			case "ForeignData" : {
				elem = new ForeignData_Type();
				elem.fromXml(reader);
				let key = createKeyFromSheetObject(elem);
				this.elements[key] = elem;
				break;
			}
		}
	}

	/**
	 * @memberOf SheetStorageAndStyles
	 * @param attrName
	 * @param reader
	 */
	SheetStorageAndStyles.prototype.readInheritedAttributes = function readInheritedAttributes(attrName, reader) {
		switch (attrName) {
			case "LineStyle": {
				this.lineStyle = reader.GetValueUInt(this.lineStyle);
				break;
			}
			case "FillStyle": {
				this.fillStyle = reader.GetValueUInt(this.fillStyle);
				break;
			}
			case "TextStyle": {
				this.textStyle = reader.GetValueUInt(this.textStyle);
				break;
			}
		}
	}

	/**
	 * @memberOf SheetStorageAndStyles
	 * @param writer
	 */
	SheetStorageAndStyles.prototype.writeInheritedAttributes = function writeInheritedAttributes(writer) {
		writer.WriteXmlNullableAttributeUInt("LineStyle", this.lineStyle);
		writer.WriteXmlNullableAttributeUInt("FillStyle", this.fillStyle);
		writer.WriteXmlNullableAttributeUInt("TextStyle", this.textStyle);
	}

	//New schema
	// <xsd:complexType name="ShapeSheet_Type">
	// 	<xsd:choice minOccurs="0" maxOccurs="unbounded">
	// 		<xsd:element name="Text" type="Text_Type" minOccurs="0" maxOccurs="1"/>
	// 		<xsd:element name="Data1" type="Data1_Type" minOccurs="0" maxOccurs="1"/>
	// 		<xsd:element name="Data2" type="Data2_Type" minOccurs="0" maxOccurs="1"/>
	// 		<xsd:element name="Data3" type="Data3_Type" minOccurs="0" maxOccurs="1"/>
	// 		<xsd:element name="ForeignData" type="ForeignData_Type" minOccurs="0" maxOccurs="1"/>
	// 		<xsd:element name="Cell" type="Cell_Type" minOccurs="0" maxOccurs="unbounded">
	// 		</xsd:element>
	// 		<xsd:element name="Section" type="Section_Type" minOccurs="0" maxOccurs="unbounded">
	// 		</xsd:element>
	// 		<xsd:any namespace="##any" processContents="lax" minOccurs="0" maxOccurs="unbounded"/>
	// 	</xsd:choice>
	// 	<xsd:attribute name="LineStyle" type="xsd:unsignedInt"/>
	// 	<xsd:attribute name="FillStyle" type="xsd:unsignedInt"/>
	// 	<xsd:attribute name="TextStyle" type="xsd:unsignedInt"/>
	// 	<xsd:anyAttribute namespace="##other" processContents="lax"/>
	// </xsd:complexType>

	// Consider choice we can have any order of elements with any occurrence
	// for elements in choice: Text, Data1, Data2, Data3, ForeignData, Cell, Section.
	// So store them in one array to save order and save all the elements
	// lets add triggers from old schema to this array too
	/**
	 * @memberOf SheetStorage
	 * @param writer
	 */
	SheetStorage.prototype.writeInheritedElements = function writeInheritedElements(writer) {
		for (const key in this.elements) {
			const elem = this.elements[key];
			switch (elem.kind) {
				case c_oVsdxSheetStorageKind.Cell_Type:
					writer.WriteXmlNullable(elem, "Cell");
					break;
				case c_oVsdxSheetStorageKind.Trigger_Type:
					writer.WriteXmlNullable(elem, "Trigger");
					break;
				case c_oVsdxSheetStorageKind.Section_Type:
					writer.WriteXmlNullable(elem, "Section");
					break;
				case c_oVsdxSheetStorageKind.Row_Type:
					writer.WriteXmlNullable(elem, "Row");
					break;
				case c_oVsdxSheetStorageKind.Text_Type:
					writer.WriteXmlNullable(elem, "Text");
					break;
				case c_oVsdxSheetStorageKind.Data_Type:
					writer.WriteXmlNullable(elem, elem.tagName);
					break;
				case c_oVsdxSheetStorageKind.ForeignData_Type:
					writer.WriteXmlNullable(elem, "ForeignData");
					break;
			}
		}
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * Finds shape section by formula. Compares N with string argument. For Geometry use find sections.
	 * @param {String} formula
	 * @memberof SheetStorage
	 * @returns {Section_Type | null}
	 */
	SheetStorage.prototype.getSection = function getSection(formula) {
		let section = this.inheritedElements[formula];
		if (section === undefined) {
			section = this.elements[formula];
		}
		return section;
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * Returns link to object not copy.
	 * @param {String} formula
	 * @memberof SheetStorage
	 * @returns {Row_Type | null}
	 */
	SheetStorage.prototype.getRow = function getRow(formula) {
		let row = this.inheritedElements[formula];
		if (row === undefined) {
			row = this.elements[formula];
		}
		return row;
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * Finds shape cell by fromula.
	 *
	 * Shape can have cells directly in them or inside sections. Lets just search for cells
	 * directly inside shapes for now.
	 *
	 * Returns object of shape not copy!
	 *
	 * visio can use formulas like Geometry1.X2 - first geometrySection, second Row, X column
	 * but Row for example can be found by IX N.
	 *
	 * Cells in Section (section is like table in visio) can exist by itself, directy in section like here
	 * https://disk.yandex.ru/d/Ud6-wmVjNnOyUA
	 * with their names like Width, Height, Angle (for Shape Thansform section)
	 * or they can be in Rows like here
	 * https://disk.yandex.ru/i/4ASd_5KHYIlXKw
	 * with names X, Y, A, B (for Geometry SplineStart Row).
	 *
	 * Let's search cells only directly in Section for now (if called on Section).
	 * @param {String} formula
	 * @memberof SheetStorage
	 * @returns {Cell_Type|null}
	 */
	SheetStorage.prototype.getCell = function getCell(formula) {
		// Cells can have N only no IX
		let cell = this.inheritedElements[formula];
		if (cell === undefined) {
			cell = this.elements[formula];
		}
		if (cell !== undefined && !(cell instanceof Cell_Type)) {
			AscCommon.consoleLog("ERR: Tried to get cell but got other object!");
		}
		return cell;
	}

	/**
	 * Calls getCell on object and tries to parse as Number(cell.v) if cell exists otherwise return undefined.
	 * @param {String} formula
	 * @param {number?} defaultValue
	 * @return {Number | undefined} number
	 */
	SheetStorage.prototype.getCellNumberValue = function (formula, defaultValue) {
		let cell = this.getCell(formula);
		let result;
		if (cell !== undefined) {
			result = Number(cell.v);
		} else {
			result = undefined;
		}
		if (defaultValue !== undefined) {
			result = result === undefined ? defaultValue : result;
		}
		return result;
	}
	/**
	 * Calls getCell on object and tries to parse as Number(cell.v) if cell exists otherwise return undefined.
	 * @param {String} formula
	 * @param {Number} pageScale
	 * @return {Number | undefined} number
	 */
	SheetStorage.prototype.getCellNumberValueWithScale = function (formula, pageScale) {
		let cell = this.getCell(formula);
		if (cell !== undefined) {
			return Number(cell.v) / pageScale;
		} else {
			return undefined;
		}
	}

	/**
	 * Calls getCell on object and tries to parse as String(cell.v) if cell exists otherwise return undefined.
	 * @param {String} formula
	 * @return {String | undefined} string
	 */
	SheetStorage.prototype.getCellStringValue = function (formula) {
		let cell = this.getCell(formula);
		if (cell !== undefined) {
			return String(cell.v);
		} else {
			return undefined;
		}
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * if in formula we have both ix and n we should use findSection instead.
	 * or if we use it with number in formula
	 * low performance function! use if can't use get section
	 * @param {String} formula
	 * @memberof SheetStorage
	 * @returns {Section_Type[] | null}
	 */
	SheetStorage.prototype.getSections = function(formula) {
		// TODO check may be optimized. maybe use getGeometrySections

		function getSections(elements, resultArr) {
			if (/^\d+$/.test(formula)) {
				// if number
				AscCommon.consoleLog('strange findSections use (with number)');
				for (const key in elements) {
					const element = elements[key];
					if (element.kind === c_oVsdxSheetStorageKind.Section_Type && String(element.ix) === formula) {
						resultArr.push(element);
					}
				}
			} else {
				for (const key in elements) {
					const element = elements[key];
					if (element.kind === c_oVsdxSheetStorageKind.Section_Type && element.n === formula) {
						resultArr.push(element);
					}
				}
			}
		}

		let resultArr = [];

		getSections(this.elements, resultArr);
		getSections(this.inheritedElements, resultArr);

		resultArr.sort(function (a, b) {
			return a.ix - b.ix;
		});

		return resultArr;
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * Used with no argument to get all rows
	 * low performance function! use if can't use get row
	 * @memberof SheetStorage
	 * @returns {Row_Type[]}
	 */
	SheetStorage.prototype.getRows = function() {
		// TODO check may be optimized. maybe use binary search for elements with maximum number as index bcs geometry
		//  rows have Row.ix as index and it is number.
		let resultArr = [];

		function getRows(elements, resultArr) {
			for (const key in elements) {
				const element = elements[key];
				if (element.kind === c_oVsdxSheetStorageKind.Row_Type) {
					resultArr.push(element);
				}
			}
		}

		getRows(this.elements, resultArr);
		getRows(this.inheritedElements, resultArr);

		resultArr.sort(function (a, b) {
			return a.ix - b.ix;
		});
		return resultArr;
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * get elements inherited from shape sheet type and own.
	 * low performance function! use if can't use getElement
	 * @memberOf SheetStorage
	 * @return {{}}
	 */
	SheetStorage.prototype.getElements = function () {
		return Object.assign({}, this.elements, this.inheritedElements);
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * @memberOf SheetStorage
	 * @param {string} formula
	 * @return {*}
	 */
	SheetStorage.prototype.getElement = function (formula) {
		// Cells can have N only no IX
		let element = this.inheritedElements[formula];
		if (element === undefined) {
			element = this.elements[formula];
		}
		return element;
	}


	/**
	 *    // Docs old:
	 *    // Section_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/section_type-complextypevisio-xml
	 * @return {Section_Type}
	 * @constructor
	 * @extends SheetStorage
	 */
	function Section_Type() {
		this.n = null;
		this.del = null;
		this.ix = null;

		// always use getter setter methods
		// Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
		// call parent class constructor
		let parentClassConstructor = SheetStorage;
		parentClassConstructor.call(this);
		return this;
	}
	// inherit parent class methods
	// https://developer.mozilla.org/ru/docs/Web/JavaScript/Reference/Global_Objects/Object/create#%D0%BF%D1%80%D0%B8%D0%BC%D0%B5%D1%80_%D0%BA%D0%BB%D0%B0%D1%81%D1%81%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%BE%D0%B5_%D0%BD%D0%B0%D1%81%D0%BB%D0%B5%D0%B4%D0%BE%D0%B2%D0%B0%D0%BD%D0%B8%D0%B5_%D1%81_object.create
	Section_Type.prototype = Object.create(SheetStorage.prototype);
	Section_Type.prototype.constructor = Section_Type;
	Section_Type.prototype.kind = c_oVsdxSheetStorageKind.Section_Type;


	/**
	 *    // Docs old:
	 *    // Row_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/row_type-complextypevisio-xml
	 * @return {Row_Type}
	 * @constructor
	 * @extends SheetStorage
	 */
	function Row_Type() {
		this.n = null;
		this.localName = null;
		this.ix = null;
		this.t = null;
		this.del = null;

		// always use getter setter methods
		// Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
		// call parent class constructor
		let parentClassConstructor = SheetStorage;
		parentClassConstructor.call(this);
		return this;
	}
	// inherit parent class methods
	// https://developer.mozilla.org/ru/docs/Web/JavaScript/Reference/Global_Objects/Object/create#%D0%BF%D1%80%D0%B8%D0%BC%D0%B5%D1%80_%D0%BA%D0%BB%D0%B0%D1%81%D1%81%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%BE%D0%B5_%D0%BD%D0%B0%D1%81%D0%BB%D0%B5%D0%B4%D0%BE%D0%B2%D0%B0%D0%BD%D0%B8%D0%B5_%D1%81_object.create
	Row_Type.prototype = Object.create(SheetStorage.prototype);
	Row_Type.prototype.constructor = Row_Type;
	Row_Type.prototype.kind = c_oVsdxSheetStorageKind.Row_Type;


	/**
	 * Docs old:
	 * Элемент RefBy (Cell_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/refby-element-cell_type-complextypevisio-xml
	 * Cell_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/cell_type-complextypevisio-xml
	 * @return {Cell_Type}
	 * @constructor
	 */
	function Cell_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);

		// read all as strings
		/**
		 * read as string
		 * @type {string}
		 */
		this.n = null;
		/**
		 * read as string
		 * @type {string}
		 */
		this.u = null;
		/**
		 * read as string
		 * @type {string}
		 */
		this.e = null;
		/**
		 * read as string
		 * @type {string}
		 */
		this.f = null;
		/**
		 * read as string
		 * @type {string}
		 */
		this.v = null;
		this.refBy = [];
		this.textContent = null;

		// not same case like in Text_Type
		// I suppose text cant go along with inner text
		// There is either textContent or refBy
		// I dont make it like it Text_Type to
		// left separate attributes refBy  and textContent and dont replace both by elements
		return this;
	}
	AscFormat.InitClass(Cell_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);
	Cell_Type.prototype.kind = c_oVsdxSheetStorageKind.Cell_Type;
	/**
	 * get String(cell.v)
	 * @memberOf Cell_Type
	 * @return {undefined | string}
	 */
	Cell_Type.prototype.getStringValue = function () {
		let cell = this;
		if (cell !== undefined) {
			return String(cell.v);
		} else {
			return undefined;
		}
	}

	/**
	 * get Number(cell.v)
	 * @memberOf Cell_Type
	 * @return {undefined | number}
	 */
	Cell_Type.prototype.getNumberValue = function () {
		let cell = this;
		if (cell !== undefined) {
			return Number(cell.v);
		} else {
			return undefined;
		}
	}

	/**
	 * try parse string "0" then get Boolean(cell.v)
	 * @memberOf Cell_Type
	 * @return {undefined | boolean}
	 */
	Cell_Type.prototype.getBooleanValue = function () {
		let cell = this;
		if (cell !== undefined) {
			if (typeof cell.v === "string") {
				if (cell.v === "1") {
					return true;
				} else if (cell.v === "0") {
					return false;
				} else {
					// unknown string
					return true;
				}
			} else {
				return Boolean(cell.v);
			}
		} else {
			return undefined;
		}
	}

	/**
	 * Can parse themeval
	 * @param {Shape_Type} shape
	 * @param {Page_Type} pageInfo
	 * @param {CTheme[]} themes
	 * @param {{fontColor?: boolean, lineUniFill?: boolean, uniFillForegnd?: boolean}?} themeValWasUsedFor - changes
	 * during function. use only for font Color LineColor and FillColor cells otherwise undefined
	 * @param {boolean?} gradientEnabled
	 * @param {number?}  themedColorsRow
	 * @return {(CUniFill | CUniColor | boolean | *)}
	 */
	Cell_Type.prototype.calculateValue = function calculateCellValue(shape, pageInfo,
																		 themes, themeValWasUsedFor,
																		 gradientEnabled, themedColorsRow) {
		let cellValue = this.v;
		let cellName = this.n;
		let cellFunction = this.f;

		let returnValue;

		// supported cells
		let fillResultCells = ["LineColor", "FillForegnd", "FillBkgnd"];
		let fillColorResultCells = ["Color", "GradientStopColor"];
		let numberResultCells = ["LinePattern", "LineWeight", "GradientStopColorTrans", "GradientStopPosition",
		"FillGradientAngle", "EndArrowSize", "BeginArrowSize", "FillPattern", "LineCap"];
		let stringResultCells = ["EndArrow", "BeginArrow", "Font"];
		let booleanResultCells = ["FillGradientEnabled"];

		// TODO handle 2.2.7.5	Fixed Theme

		if (cellValue === "Themed" || cellFunction === "THEMEVAL()") {
			// equal to THEMEVAL() call
			// add themeval support for every supported cell
			returnValue = AscVisio.themeval(this, shape, pageInfo, themes, undefined,
				undefined, gradientEnabled, themedColorsRow);

			if (themeValWasUsedFor) {
				if (cellName === "LineColor") {
					themeValWasUsedFor.lineUniFill = true;
				} else if (cellName === "FillForegnd") {
					themeValWasUsedFor.uniFillForegnd = true;
				} else if (cellName === "Color") {
					// for text color
					themeValWasUsedFor.fontColor = true;
				}
			}
		} else if (fillResultCells.includes(cellName) || fillColorResultCells.includes(cellName)) {
			let rgba = null;
			if (/#\w{6}/.test(cellValue)) {
				// check if hex
				rgba = AscCommon.RgbaHexToRGBA(cellValue);
			} else {
				let colorIndex = parseInt(cellValue);
				if (!isNaN(colorIndex)) {
					switch (colorIndex) {
						case 0:
							rgba = AscCommon.RgbaHexToRGBA('#000000');
							break;
						case 1:
							rgba = AscCommon.RgbaHexToRGBA('#FFFFFF');
							break;
						case 2:
							rgba = AscCommon.RgbaHexToRGBA('#FF0000');
							break;
						case 3:
							rgba = AscCommon.RgbaHexToRGBA('#00FF00');
							break;
						case 4:
							rgba = AscCommon.RgbaHexToRGBA('#0000FF');
							break;
						case 5:
							rgba = AscCommon.RgbaHexToRGBA('#FFFF00');
							break;
						case 6:
							rgba = AscCommon.RgbaHexToRGBA('#FF00FF');
							break;
						case 7:
							rgba = AscCommon.RgbaHexToRGBA('#00FFFF');
							break;
						case 8:
							rgba = AscCommon.RgbaHexToRGBA('#800000');
							break;
						case 9:
							rgba = AscCommon.RgbaHexToRGBA('#008000');
							break;
						case 10:
							rgba = AscCommon.RgbaHexToRGBA('#000080');
							break;
						case 11:
							rgba = AscCommon.RgbaHexToRGBA('#808000');
							break;
						case 12:
							rgba = AscCommon.RgbaHexToRGBA('#800080');
							break;
						case 13:
							rgba = AscCommon.RgbaHexToRGBA('#008080');
							break;
						case 14:
							rgba = AscCommon.RgbaHexToRGBA('#C0C0C0');
							break;
						case 15:
							rgba = AscCommon.RgbaHexToRGBA('#E6E6E6');
							break;
						case 16:
							rgba = AscCommon.RgbaHexToRGBA('#CDCDCD');
							break;
						case 17:
							rgba = AscCommon.RgbaHexToRGBA('#B3B3B3');
							break;
						case 18:
							rgba = AscCommon.RgbaHexToRGBA('#9A9A9A');
							break;
						case 19:
							rgba = AscCommon.RgbaHexToRGBA('#808080');
							break;
						case 20:
							rgba = AscCommon.RgbaHexToRGBA('#666666');
							break;
						case 21:
							rgba = AscCommon.RgbaHexToRGBA('#4D4D4D');
							break;
						case 22:
							rgba = AscCommon.RgbaHexToRGBA('#333333');
							break;
						case 23:
							rgba = AscCommon.RgbaHexToRGBA('#1A1A1A');
							break;
						default:
							AscCommon.consoleLog("error: unknown color index");
							rgba = AscCommon.RgbaHexToRGBA('#000000');
							break;
					}
				} else {
					AscCommon.consoleLog("error: color index is null");
					rgba = AscCommon.RgbaHexToRGBA('#000000');
				}
			}

			if (fillResultCells.includes(cellName)) {
				returnValue = AscFormat.CreateUnfilFromRGB(rgba.R, rgba.G, rgba.B);
			} else if (fillColorResultCells.includes(cellName)) {
				// for text color
				returnValue = AscFormat.CreateUnfilFromRGB(rgba.R, rgba.G, rgba.B).fill.color;
			} else {
				AscCommon.consoleLog("wrong calculateCellValue argument cell. Cell unsupported. return null");
				return null;
			}
		} else if (numberResultCells.includes(cellName)) {
			let cellNumberValue = this.getNumberValue();
			if (!isNaN(cellNumberValue)) {
				if (cellName === "GradientStopPosition") {
					cellNumberValue *= 100000;
				} else if (cellName === "FillGradientAngle") {
					let angleRads = cellNumberValue;
					let angle = null;
					// 20.1.10.3 ST_Angle (Angle)
					// This simple type represents an angle in 60,000ths of a degree. Positive angles are clockwise (i.e., towards the
					// positive y axis); negative angles are counter-clockwise (i.e., towards the negative y axis)
					// direction is considered in global transform
					let stAngle = angleRads / Math.PI * 180 * AscFormat.degToC;
					if (!isNaN(stAngle)) {
						angle = stAngle;
					} else {
						angle = 5400000;
					}
					cellNumberValue = angle;
				} else if (cellName === "LineCap") {
					switch (cellNumberValue) {
						case 0:
							cellNumberValue = 1;
							break;
						case 1:
							cellNumberValue = 0;
							break;
						case 2:
							cellNumberValue = 2;
							break;
					}
				}
				return cellNumberValue;
			}
		} else if (booleanResultCells.includes(cellName)) {
			let cellBooleanValue = this.getBooleanValue();
			return cellBooleanValue;
		} else if (stringResultCells.includes(cellName)) {
			let cellStringValue = this.getStringValue();
			return cellStringValue;
		} else {
			AscCommon.consoleLog("Cell was not calculated in calculate cell value");
		}

		// code below is unused because we dont need default values here. Cell value is either parsed well or
		// default value is used in themeval function.
		// if (returnValue === null || returnValue === undefined) {
		// 	if (cellName === "LineColor" || cellName === "FillForegnd" || cellName === "FillBkgnd") {
		// 		AscCommon.consoleLog("no color found. so painting lt1.");
		// 		returnValue = AscFormat.CreateUniFillByUniColor(AscFormat.builder_CreateSchemeColor("lt1"));
		// 	} else if (cellName === "Color") {
		// 		// cellName === "Color" for text color
		// 		AscCommon.consoleLog("no text color found. so painting dk1.");
		// 		returnValue = AscFormat.builder_CreateSchemeColor("dk1");
		// 	} else if (cellName === "GradientStopColor") {
		// 		AscCommon.consoleLog("no GradientStopColor color found. so painting lk1.");
		// 		returnValue = AscFormat.builder_CreateSchemeColor("lt1");
		// 	} else {
		// 		AscCommon.consoleLog("no calculateCellValue result return undefined");
		// 	}
		// }
		return returnValue;
	}


	// /**
	//  * @memberOf Cell_Type
	//  * @return {number}
	//  */
	// Cell_Type.prototype.getValueInMM = function () {
	// 	let res;
	// 	//todo all units
	// 	switch (this.u) {
	// 		case "DL":
	// 		case "IN":
	// 		case "IN_F":
	// 			res = parseFloat(this.v) * g_dKoef_in_to_mm;
	// 			break;
	// 		case "FT":
	// 			res = parseFloat(this.v) * 12 * g_dKoef_in_to_mm;
	// 			break;
	// 		case "F_I":
	// 			res = parseFloat(this.v);
	// 			let intPart = Math.floor(res);
	// 			res = (intPart * 12 + (res - intPart)) * g_dKoef_in_to_mm;
	// 			break;
	// 		case "KM":
	// 			res = parseFloat(this.v) * 1000000;
	// 			break;
	// 		case "M":
	// 			res = parseFloat(this.v) * 1000;
	// 			break;
	// 		case "CM":
	// 			res = parseFloat(this.v) * 10;
	// 			break;
	// 		case "MM":
	// 			res = parseFloat(this.v);
	// 			break;
	// 		default:
	// 			res = parseFloat(this.v) * g_dKoef_in_to_mm;
	// 			break;
	// 	}
	// 	return res;
	// };

	// /**
	//  * @memberOf Cell_Type
	//  * @return {number}
	//  */
	// Cell_Type.prototype.getValueInInch = function () {
	// 	let res = this.getValueInMM() / g_dKoef_in_to_mm;
	// 	return res;
	// }


	/**
	 *    // Docs old:
	 *    // Useless see ShapeSheet_Type - old
	 *    // Inherites(extends) ShapeSheet_Type in new schema
	 * @return {Shape_Type}
	 * @constructor
	 * @extends SheetStorageAndStyles
	 */
	function Shape_Type() {
		this.id = null;
		this.originalID = null;
		this.del = null;
		this.masterShape = null;
		this.uniqueID = null;
		this.name = null;
		this.nameU = null;
		this.isCustomName = null;
		this.isCustomNameU = null;
		this.master = null;
		this.type = null;
		/**
		 * use get subshapes method
		 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
		 * @type {Shape_Type[]}
		 */
		this.shapes = [];

		/**
		 *	Own shapes and inherited
		 * @type {Shape_Type[]}
		 */
		this.inheritedShapes = [];

		/**
		 * Shape_Type.prototype.toGeometryAndTextCShapes creates CShape from Shape_Type but for image as an
		 * exception we make variable to store CImageShape in advance. We read CImageShape on parsing in
		 * AscVisio.Shape_Type.prototype.fromXml. Because CImageShape needs StaxParser reader object for init.
		 * @type {CImageShape}
		 */
		this.cImageShape = null;

		/**
		 * see MS-VSDX 2.2.7.4.9	Connector. if true shape is connector
		 * @type {boolean}
		 */
		this.isConnectorStyleIherited = false;

		// call parent class constructor
		let parentClassConstructor = SheetStorageAndStyles;
		parentClassConstructor.call(this);

		// return this;
	}
	// inherit parent class methods
	// https://developer.mozilla.org/ru/docs/Web/JavaScript/Reference/Global_Objects/Object/create#%D0%BF%D1%80%D0%B8%D0%BC%D0%B5%D1%80_%D0%BA%D0%BB%D0%B0%D1%81%D1%81%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%BE%D0%B5_%D0%BD%D0%B0%D1%81%D0%BB%D0%B5%D0%B4%D0%BE%D0%B2%D0%B0%D0%BD%D0%B8%D0%B5_%D1%81_object.create
	Shape_Type.prototype = Object.create(SheetStorageAndStyles.prototype);
	Shape_Type.prototype.constructor = Shape_Type;


	/**
	 * @memberOf Shape_Type
	 * @return {*} masterId
	 */
	Shape_Type.prototype.getMasterID = function() {
		return this.master;
	}

	/**
	 * get Shape properties that come from layers.
	 * if shape has multiple layers attached only equal layer properties applied (and added to object).
	 * @memberOf Shape_Type
	 * @param {Page_Type} pageInfo
	 * @return {{ [key: string]: Cell_Type }} An object mapping string identifiers to Cell_Type instances
	 */
	Shape_Type.prototype.getLayerProperties = function getLayerProperties(pageInfo) {
		let layerMemberString = this.getCellStringValue("LayerMember");
		if (layerMemberString === undefined || layerMemberString === "") {
			return {};
		}
		let layersArray = layerMemberString.split(";");
		let layersInfo = pageInfo.pageSheet.getSection("Layer");
		if (layersInfo === undefined) {
			return {};
		}
		let previousLayerElements = undefined;
		/** @type {Set<string>} */
		let unEqualProperties = new Set();
		layersArray.forEach(function (layerIndexString) {
			let layerIndex = Number(layerIndexString);
			let layerInfo = layersInfo.getRow(layerIndex);
			let layerElements = layerInfo.getElements();

			// Unlink original array
			let layerElementsCopy= {};
			for (const key in layerElements) {
				if (layerElements.hasOwnProperty(key)) {
					layerElementsCopy[key] = layerElements[key];
				}
			}

			// Set default color
			if (layerElementsCopy["Color"] === undefined) {
				let defaultColorCell = new Cell_Type();
				defaultColorCell.n = "Color";
				defaultColorCell.v = "0";
				layerElementsCopy["Color"] = defaultColorCell;
			}

			if (previousLayerElements === undefined) {
				previousLayerElements = layerElementsCopy;
			} else {
				// compare with previous shape layer
				for (const cellKey in layerElementsCopy) {
					const cell = layerElementsCopy[cellKey];
					let previousLayerCell = previousLayerElements[cell.n];
					if (previousLayerCell.v !== cell.v) {
						unEqualProperties.add(cell.n);
					}
				}
				previousLayerElements = layerElementsCopy;
			}
		});
		// layers have the same set of properties so lets take any of them
		// and remove unEqualProperties
		/** @type {{ [key: string]: Cell_Type }} */
		let resultObject = {};
		if (previousLayerElements === undefined) {
			return resultObject;
		}
		for (const cellKey in previousLayerElements) {
			const cell = previousLayerElements[cellKey];
			if (!unEqualProperties.has(cell.n)) {
				resultObject[cell.n] = cell;
			}
		}
		return resultObject;
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * @memberof Shape_Type
	 * @return {Shape_Type[]}
	 */
	Shape_Type.prototype.getSubshapes = function () {
		return this.shapes.concat(this.inheritedShapes);
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * Finds shape text element.
	 *
	 * Returns object of shape not copy!
	 *
	 * @memberof Shape_Type
	 * @returns {Text_Type | null}
	 */
	Shape_Type.prototype.getTextElement = function getTextElement() {
		let text = this.inheritedElements["Text"];
		if (text === undefined) {
			text = this.elements["Text"];
		}
		return text;
	}

	/**
	 * returns this shape and subshapes array without cloning so objects are linked.
	 * @param [resultArray = []]
	 * @memberof Shape_Type
	 * @return {Shape_Type[]}
	 */
	Shape_Type.prototype.collectSubshapesRecursive = function(resultArray) {
		if (resultArray === undefined) {
			resultArray = [];
		}

		// ! Don't change cell.v bcs you may change master values

		resultArray.push(this);

		let subShapes = this.getSubshapes();
		for (let j = 0; j < subShapes.length; j++) {
			const subShape = subShapes[j];

			subShape.collectSubshapesRecursive(resultArray);
		}

		return resultArray;
	}

	/**
	 * clones masters shapes to given shape.
	 * Uses MasterShapeAttributes to find shapes to insert.
	 * @param {Shape_Type[]} masterSubshapes
	 * @param masters - result from joinMastersInfoAndContents()
	 */
	Shape_Type.prototype.cloneSubshapes = function cloneSubshapes(masterSubshapes, masters) {
		function findIndexComparingByMasterShapeAttribute(shapeSubshapes, masterSubshape) {
			return shapeSubshapes.findIndex(function (element) {
				return element.masterShape === masterSubshape.id;
			});
		}

		// If subshape has Master attribute with id of any master: call realizeMasterToShapeInheritance
		// If subshape has MasterShape attribute with id of any parents shape masters subshapes:
		// 	call mergeElementArrays NO RECURSION HERE
		// If there is a shape in master but there is no such local subshape then it should be inherited (copied)
		// If there is a shape in master but there is no such local subshape then it should be inherited (copied)
		// lets check if it exists locally only by MasterShape attribute
		// examples it the bottom of the function

		// handle subshapes MasterShape attribute
		let thisContext = this;
		masterSubshapes.forEach(function(masterSubshape) {
			let mergeElementIndex = findIndexComparingByMasterShapeAttribute(thisContext.getSubshapes(),
					masterSubshape);
			let elementExistsAlready = mergeElementIndex !== -1;

			// 2.2.5.4.1	Master-to-Shape Inheritance
			// "subshapes not specified in the instance are inherited from the master." (from its master)
			if (!elementExistsAlready) {
				// maybe add masterShape attribute to new shape - lets dont do it because:
				// of recursive iterations of inheritance we will try to inherit because we will se masterShape but
				// there is no need because it is copy pasted element no need in inheritance
				// maybe consider id to insert in ascending order
				thisContext.inheritedShapes.push(clone(masterSubshape));
			} else {
				// 2.2.5.4.1	Master-to-Shape Inheritance
				// "if an instance contains a subshape whose ShapeSheet_Type element has a MasterShape attribute that matches
				// the ID attribute of a subshape of the master, the local properties specified in this subshape will
				// override those of the corresponding subshape in the master."

				// let masterElements = masterSubshape.elements;
				// let shapeElements = shapeSubshapes[mergeElementIndex].elements;
				// mergeElementArrays(masterElements);

				// it is done in realizeMasterToShapeInheritanceRecursive with subshapes handle
			}
		});
	}


	/**
	 * Realizes Master-To-Shape inheritance.
	 * Comes through the shape recursively through all subshapes
	 * and copies properties from their masters if master exist.
	 * Not uses clone function.
	 *
	 * This function erases original shape object and after function call we cant find what props are
	 * inherited.
	 *
	 * Also a memory problem. Master attributes copy to each shape.
	 * but on the most heavy file shapeClasses after inheritance is only 217 megabytes
	 *
	 * TODO:
	 * 	maybe like this
	 * 	rewrite: just set simple links to masters for shape and subshapes in this function BUT
	 * 	!!! always use function getCell/getRow/getSection/getSubshape/getElements which will search for element
	 * 	in this shape first and then in master. Also use getElements or getSubshapes functions to get all
	 * 	cells/rows/sections or subshapes, theese function will first merge master and shape elements/subshapes
	 * 	and then return elements/subshapes.
	 *  Also use setter functions.
	 *
	 * @param masters
	 * @param {Shape_Type[]?} ancestorMasterShapes
	 * @return {Shape_Type} shape with all its props
	 */
	Shape_Type.prototype.realizeMasterInheritanceRecursively = function(masters, ancestorMasterShapes) {
		// 2.2.5.4.1	Master-to-Shape Inheritance
		// If shape has master or masterShape that have 1 top level shape
		//  - inherit elements (sections, rows, cells) and subshapes
		// If shape has master or masterShape that have several top level shapes
		// - inherit only subshapes as shape subshapes
		//
		// subshapes inheritance:
		// 2.2.5.4.1	Master-to-Shape Inheritance
		// "if an instance contains a subshape whose ShapeSheet_Type (now it is Shape_Type)
		// element has a MasterShape attribute that matches the ID attribute of a subshape of the master,
		// the local properties specified in this subshape will override those of the corresponding subshape
		// in the master."
		// "subshapes not specified in the instance are inherited from the master." (from its master)
		//
		// So in main function come across all shapes and subshapes and call realizeMasterToShapeInheritanceRecursive
		//
		// realizeMasterToShapeInheritanceRecursive:
		// shape can have master or masterShape id and be it master or masterShape if master (master or masterShape) have
		// 1 top level shape it inherits its elements, then handle subshapes.
		//
		// handle subshapes:
		// come along master and check MasterShape attributes if we have no shape with this MasterShape copy it to
		// out shape.
		// What if we inherit with MasterShape id not with Master id do we check MasterShape attributes from our Shape?
		// - I think yes we compare subshape ids from masterShape with MasterShape ids of subshapes of shape that
		// is beiing inherited (check sub MasterShape ids)
		// We dont need to merge subshapes elements in currents step because we will make it in recursion iterations?
		// Its just a detail of realization but yes. For subshapes we merge elements on next steps on recursive calls.
		//
		// handle several top level subshapes:
		// just handle shape subshapes like above
		//
		// - can shape just consist of multiple shapes like master?
		// maybe dont memorize MasterShape shape. but when we inherit one level below its masterIds

		// - should we search for Master shape with the specified MasterShape recursively?
		// previously i set MasterShape as ancestorMasterShapes and didnt use collectSubshapesRecursive
		// to search for master but now code is code is more flexible but picture didnt change anyway
		// code is more readable

		// - should we compare nested shapes of shape and master recursively?
		// no bcs - realizeMasterToShapeInheritanceRecursive will call itself on subshapes so it will call
		// compare nested shapes recursively (mergeSubshapes)
		// but what if need to call merge subshapes recursively on subshapes relative to this master?
		// - no because if subshape at any deep have no Master/MasterShape in need no
		// call mergeSubshapes otherwise realizeMasterToShapeInheritanceRecursive will call mergeSubshapes because
		// shape have Master/MasterShape attribute.

		// - what about both master and masterShape inheritance?

		// - dont forget to inherit links to styles from master

		// Consider examples
		// <PageContents>
		//   <Shapes>
		//    	<Shape ID='22' NameU='Process' Name='Process' Type='Group' Master='4'>
		// 	  		<Cell N='PinX' V='2.75'/>
		//      	<Cell N='PinY' V='6.375'/>
		// 	  		...
		//       	<Shapes>
		//         		<Shape ID='98' NameU='Flags' Name='Flags' Type='Group' Master='26'>
		// 		  				<Cell N='PinX' V='0.875' U='IN'/>
		//          		<Cell N='PinY' V='0.9375' U='IN'/>
		// 		  				...
		//
		// <PageContents>
		//   <Shapes>
		//     <Shape ID='31' NameU='Headquarters' Name='Headquarters' Type='Group' Master='2'>
		//       <Cell N='PinX' V='5.5'/>
		//       <Cell N='PinY' V='7.097826086956522'/>
		// 	  		...
		// 	  		<Shapes>
		//         		<Shape ID='32' Type='Shape' MasterShape='6'>
		//           		<Cell N='PinX' V='1.113900526692464' F='Inh'/>
		//           		<Cell N='PinY' V='1.16976435446399' F='Inh'/>
		// 		  				...


		let masterShapesToInheritFrom = [];

		// lets create copy of this shape work with it then return it
		// let thisShapeCopy = clone(this);

		// check Master attribute and set shapes/shape
		// to inherit from: masterShapesToInheritFrom and ancestorMasterShapes
		let topShapeMasterId = this.getMasterID();
		if (topShapeMasterId !== null && topShapeMasterId !== undefined) {
			let topShapeMasterIndex = masters.findIndex(function(masterObject) {
				return masterObject.id === topShapeMasterId;
			});
			let topShapeMaster = masters[topShapeMasterIndex];

			if (topShapeMaster) {
				let masterShapes = topShapeMaster.content.shapes;
				masterShapesToInheritFrom = masterShapes;

				// all descendant shapes will inherit from that master
				ancestorMasterShapes = masterShapesToInheritFrom;
			}
		}

		// check MasterShape attribute and set shapes/shape
		// to inherit from: masterShapesToInheritFrom and ancestorMasterShapes
		let masterShapeId = this.masterShape;
		if (masterShapeId !== null && masterShapeId !== undefined) {
			if (ancestorMasterShapes === null || ancestorMasterShapes === undefined) {
				AscCommon.consoleLog("MasterShape attribute is set but Master is not set for ", this);
			} else {
				let masterIndex = -1;
				if (ancestorMasterShapes.length === 1) {
					// if master has one top level shape
					let masterSubshapes = ancestorMasterShapes[0].collectSubshapesRecursive();
					masterIndex = masterSubshapes.findIndex(function (masterSubshape) {
						return masterShapeId === masterSubshape.id;
					});
					let masterShape = masterSubshapes[masterIndex];
					masterShapesToInheritFrom = [masterShape];
				} else {
					let masterSubshapes = [];
					ancestorMasterShapes.forEach(function(ancestorMasterShape) {
						let masterSubshapesNth = ancestorMasterShape.collectSubshapesRecursive();
						masterSubshapes = masterSubshapes.concat(masterSubshapesNth);
					})
					masterIndex = masterSubshapes.findIndex(function (masterSubshape) {
						return masterShapeId === masterSubshape.id;
					});
					let masterShape = masterSubshapes[masterIndex];
					masterShapesToInheritFrom = [masterShape];
				}

				if (masterIndex === -1) {
					AscCommon.consoleLog('For MasterShape = ', masterShapeId, 'shape not found in master. Check shape: ', this);
				}
			}
		}

		// inherit: mergeElements and clone shapes to which elements will be merged on recursive calls
		if (masterShapesToInheritFrom.length === 1) {
			let masterShapeToInheritFrom = masterShapesToInheritFrom[0];

			// inherit link to styles
			if (!this.inheritedLineStyle) {
				this.inheritedLineStyle = masterShapeToInheritFrom.lineStyle;
			}
			if (!this.inheritedFillStyle) {
				this.inheritedFillStyle = masterShapeToInheritFrom.fillStyle;
			}
			if (!this.inheritedTextStyle) {
				this.inheritedTextStyle = masterShapeToInheritFrom.textStyle;
			}

			let masterElements = masterShapeToInheritFrom.getElements();
			this.mergeElementArrays(masterElements);
			if (masterShapeToInheritFrom.type === AscVisio.SHAPE_TYPES_FOREIGN) {
				if (masterShapeToInheritFrom.cImageShape) {
					this.cImageShape = clone(masterShapeToInheritFrom.cImageShape);
				}
			}

			let masterSubshapes = masterShapeToInheritFrom.getSubshapes();
			this.cloneSubshapes(masterSubshapes, masters);
		} else if (masterShapesToInheritFrom.length > 1) {
			// does it ever happens?
			// what about style inheritance?
			this.cloneSubshapes(masterShapesToInheritFrom, masters);
		}

		// call recursive on all subshapes
		let subshapes = this.getSubshapes();
		subshapes.forEach(function(shape) {
			shape.realizeMasterInheritanceRecursively(masters, ancestorMasterShapes);
		});

		// return thisShapeCopy;
		// end of method
	}

	/**
	 * Inherits style to master/shape or style to style. Process is the same.
	 * Inherits all style elements (sections, rows, cells).
	 * @param {Shape_Type | StyleSheet_Type} thisArgument
	 * @param {StyleSheet_Type[]} styles
	 * @param {?Set} stylesWithRealizedInheritance
	 */
	function realizeStyleToSheetObjInheritanceRecursive(thisArgument, styles, stylesWithRealizedInheritance) {
		if (stylesWithRealizedInheritance.has(thisArgument)) {
			// thisArgument is style not shape and it has realized inheritance already
			// AscCommon.consoleLog("style has realized inheritance already. return");
			return;
		}

		/**
		 * see MS-VSDX 2.2.7.4.9	Connector.
		 * @param {Shape_Type | StyleSheet_Type} object
		 * @param {StyleSheet_Type} style
		 */
		function setIsConnectorStyleInherited(object, style) {
			object.isConnectorStyleIherited = object.isConnectorStyleIherited ? true : style.nameU === "Connector";
		}

		if (!(thisArgument.getLineStyle() === thisArgument.getFillStyle()
				&& thisArgument.getLineStyle() === thisArgument.getTextStyle())) {
			// Attribute	Cell_Type elements

			// LineStyle	Specifies Cell_Type elements related to line properties except for Cell_Type child elements
			// of a FillGradient Section_Type.
			// Line property information in shapes, masters, and styles is specified by the LineColor, LinePattern, LineWeight,
			// LineCap, BeginArrow, EndArrow, LineColorTrans, CompoundType, BeginArrowSize, EndArrowSize, Rounding,
			// LineGradientDir, LineGradientAngle, and LineGradientEnabled Cell_Type elements, and the Cell_Type
			// elements belonging to the LineGradient Section_Type.

			// FillStyle	Specifies Cell_Type elements related to fill properties and effect properties
			// including Cell_Type child elements of a FillGradient Section_Type.
			// Fill property information in shapes, masters, and styles is specified by the FillForegnd,
			// FillForegndTrans, FillBkgnd, FillBkgndTrans, FillPattern,
			// FillGradientDir, FillGradientAngle, FillGradientEnabled,
			// RotateGradientWithShape, and UseGroupGradientCell_Type elements,
			// and the Cell_Type elements belonging to the FillGradient Section_Type.
			// Shadow effect set information in shapes, masters, and styles is specified by the ShdwForegnd, ShdwForegndTrans,
			// ShdwPattern, ShapeShdwType, ShapeShdwOffsetX, ShapeShdwOffsetY, ShapeShdwObliqueAngle, ShapeShdwScaleFactor,
			// and ShapeShdwBlur Cell_Type elements.

			// TextStyle	Specifies Cell_Type elements related to text.

			// What about Quick style cells?

			// cells rows sections
			// TODO check cells inside section LineGradient
			let lineStyleElements = ["LineColor", "LinePattern", "LineWeight", "LineCap", "BeginArrow", "EndArrow",
				"LineColorTrans", "CompoundType", "BeginArrowSize", "EndArrowSize", "Rounding",
				"LineGradientDir", "LineGradientAngle", "LineGradientEnabled", "LineGradient",
				"QuickStyleLineColor", "QuickStyleLineMatrix"];

			// TODO check cells inside section FillGradient
			let fillStyleElements = ["FillForegnd", "FillForegndTrans", "FillBkgnd", "FillBkgndTrans", "FillPattern",
					"FillGradientDir", "FillGradientAngle", "FillGradientEnabled",
					"RotateGradientWithShape", "UseGroupGradientCell_Type", "FillGradient",
					"ShdwForegnd", "ShdwForegndTrans", "ShdwPattern", "ShapeShdwType", "ShapeShdwOffsetX", "ShapeShdwOffsetY",
					"ShapeShdwObliqueAngle", "ShapeShdwScaleFactor", "ShapeShdwBlur",
				"QuickStyleFillColor", "QuickStyleFillMatrix"];

			let textStyleElements = ["TextBkgnd", "TextDirection", "TextBkgndTrans", "LockTextEdit", "HideText",
					"TheText", "IsTextEditTarget", "KeepTextFlat", "ReplaceLockText", "TextPosAfterBullet",
					"Character", "Paragraph", "Tabs", "DefaultTabStop", "VerticalAlign",
					"BottomMargin", "TopMargin", "RightMargin", "LeftMargin"];

			let commonElements = ["ColorSchemeIndex", "EffectSchemeIndex", "ConnectorSchemeIndex", "FontSchemeIndex",
					"ThemeIndex", "VariationColorIndex", "VariationStyleIndex", "EmbellishmentIndex",
			"QuickStyleLineColor", "QuickStyleFillColor", "QuickStyleShadowColor", "QuickStyleFontColor",
			"QuickStyleLineMatrix", "QuickStyleFillMatrix", "QuickStyleEffectsMatrix", "QuickStyleFontMatrix",
			"QuickStyleType", "QuickStyleVariation"];
			lineStyleElements = lineStyleElements.concat(commonElements);
			fillStyleElements = fillStyleElements.concat(commonElements);
			textStyleElements = textStyleElements.concat(commonElements);

			if (thisArgument.getLineStyle() !== null) {
				let styleId = Number(thisArgument.getLineStyle());
				let styleSheet = styles.find(function(style) {
					return style.id === styleId;
				});
				setIsConnectorStyleInherited(thisArgument, styleSheet);
				realizeStyleToSheetObjInheritanceRecursive(styleSheet, styles, stylesWithRealizedInheritance);
				thisArgument.mergeElementArrays(styleSheet.getElements(), lineStyleElements);
			}

			if (thisArgument.getFillStyle() !== null) {
				let styleId = Number(thisArgument.getFillStyle());
				let styleSheet = styles.find(function(style) {
					return style.id === styleId;
				});
				setIsConnectorStyleInherited(thisArgument, styleSheet);
				realizeStyleToSheetObjInheritanceRecursive(styleSheet, styles, stylesWithRealizedInheritance);
				thisArgument.mergeElementArrays(styleSheet.getElements(), fillStyleElements);
			}

			if (thisArgument.getTextStyle() !== null) {
				let styleId = Number(thisArgument.getTextStyle());
				let styleSheet = styles.find(function(style) {
					return style.id === styleId;
				});
				setIsConnectorStyleInherited(thisArgument, styleSheet);
				realizeStyleToSheetObjInheritanceRecursive(styleSheet, styles, stylesWithRealizedInheritance);
				thisArgument.mergeElementArrays(styleSheet.getElements(), textStyleElements);
			}
			if (thisArgument.constructor === AscVisio.StyleSheet_Type) {
				// memorize: that style has realized inheritance
				stylesWithRealizedInheritance.add(thisArgument);
			}

			return;
		}

		if (thisArgument.getLineStyle() === null && thisArgument.getFillStyle() === null
				&& thisArgument.getTextStyle() === null) {
			// AscCommon.consoleLog('Top parent style');
			return;
		}

		// if lineStyle === textStyle === fillStyle so let's take lineStyle
		let styleId = Number(thisArgument.getLineStyle());
		let styleSheet = styles.find(function(style) {
			return style.id === styleId;
		});
		setIsConnectorStyleInherited(thisArgument, styleSheet);

		realizeStyleToSheetObjInheritanceRecursive(styleSheet, styles, stylesWithRealizedInheritance);
		thisArgument.mergeElementArrays(styleSheet.getElements());
		if (thisArgument.constructor === AscVisio.StyleSheet_Type) {
			// memorize: that style has realized inheritance
			stylesWithRealizedInheritance.add(thisArgument);
		}
	}

	/**
	 * Style-To-Shape inheritance
	 * Copy all style elements (sections, rows, cells) to shape.
	 * (Doesn't take much memory < 300MB with master inheritance for the most large files).
	 * stylesWithRealizedInheritance was added for optimization.
	 * Check if this shape or sub-shapes have lineStyle/fillStyle/textStyle if so realize inheritance with
	 * recursive style inheritance
	 * @param styles
	 * @param {?Set} [stylesWithRealizedInheritance]
	 * @memberOf Shape_Type
	 */
	Shape_Type.prototype.realizeStyleInheritanceRecursively = function(styles, stylesWithRealizedInheritance) {
		if (stylesWithRealizedInheritance === undefined) {
			stylesWithRealizedInheritance = new Set();
		}
		realizeStyleToSheetObjInheritanceRecursive(this, styles, stylesWithRealizedInheritance);

		// call recursive on all subshapes
		let subshapes = this.getSubshapes();
		subshapes.forEach(function(shape) {
			shape.realizeStyleInheritanceRecursively(styles, stylesWithRealizedInheritance);
		});
	}

	/**
	 * @memberOf Shape_Type
	 * @return {ForeignData_Type | undefined}
	 */
	Shape_Type.prototype.getForeignDataObject = function getForeignDataObject() {
		let result = this.elements["ForeignData"];
		if (result === undefined) {
			result = this.inheritedElements["ForeignData"];
		}
		return result;
		// return this.elements.find(function findForeignData(element) {
		// 	return element.constructor.name === "ForeignData_Type";
		// });
	}

	/**
	 * get deep copy of object with prototypes
	 * @param object
	 * @return {any}
	 */
	function clone(object) {
		function cloneWithPrototypesRecursive(copyObject, originalObject) {

			// Iterate over object properties recursively
			for (const key in originalObject) {
				if (originalObject.hasOwnProperty(key)) {
					if (typeof originalObject[key] === 'object' && originalObject[key] !== null) {
						// after recursive call when we set array props using copyObject[key] = originalObject[key];
						// array length will change automatically
						copyObject[key] = Array.isArray(originalObject[key]) ? [] :
							Object.create(Object.getPrototypeOf(originalObject[key]));
						cloneWithPrototypesRecursive(copyObject[key], originalObject[key]);
					} else {
						copyObject[key] = originalObject[key];
					}
				}
			}
		}

		let copy = Object.create(Object.getPrototypeOf(object));
		// let copy = JSON.parse(JSON.stringify(object));
		cloneWithPrototypesRecursive(copy, object);
		// AscCommon.consoleLog("Clone function test.\nObject before: ", object, "\nAfter: ", copy);
		return copy;
	}


	/**
	 * // Docs old:
	 * // Элемент PageSheet (Page_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/pagesheet-element-page_type-complextypevisio-xml
	 * // Элемент Rel (Page_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/rel-element-page_type-complextypevisio-xml
	 * // Page_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/page_type-complextypevisio-xml
	 * @returns {Page_Type}
	 * @constructor
	 */
	function Page_Type() {
		this.id = null;
		this.name = null;
		this.nameU = null;
		this.isCustomName = null;
		this.isCustomNameU = null;
		this.background = null;
		this.backPage = null;
		this.viewScale = null;
		this.viewCenterX = null;
		this.viewCenterY = null;
		this.reviewerID = null;
		this.associatedPage = null;
		this.pageSheet = null;
		this.rel = null;

		//todo objectId
		this.deleteLock = new AscVisio.PropLocker(undefined);
		return this;
	}
	AscFormat.InitClass(Page_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	/**
	 * // Docs old:
	 * // PageSheet_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/pagesheet_type-complextypevisio-xml
	 * // In new schema inherits from ShapeSheet_Type
	 * @returns {PageSheet_Type}
	 * @constructor
	 * @extends SheetStorageAndStyles
	 */
	function PageSheet_Type() {
		this.uniqueID = null;

		// call parent class constructor
		let parentClassConstructor = SheetStorageAndStyles;
		parentClassConstructor.call(this);
		return this;
	}
	// inherit parent class methods
	// https://developer.mozilla.org/ru/docs/Web/JavaScript/Reference/Global_Objects/Object/create#%D0%BF%D1%80%D0%B8%D0%BC%D0%B5%D1%80_%D0%BA%D0%BB%D0%B0%D1%81%D1%81%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%BE%D0%B5_%D0%BD%D0%B0%D1%81%D0%BB%D0%B5%D0%B4%D0%BE%D0%B2%D0%B0%D0%BD%D0%B8%D0%B5_%D1%81_object.create
	PageSheet_Type.prototype = Object.create(SheetStorageAndStyles.prototype);
	PageSheet_Type.prototype.constructor = PageSheet_Type;

	/**
	 * // Docs old:
	 * // DocumentSheet_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/documentsheet_type-complextypevisio-xml
	 * // In new schema inherites from ShapeSheet_Type
	 * @returns {DocumentSheet_Type}
	 * @constructor
	 * @extends SheetStorageAndStyles
	 */
	function DocumentSheet_Type() {
		this.name = null;
		this.nameU = null;
		this.isCustomName = null;
		this.isCustomNameU = null;
		this.uniqueID = null;

		// call parent class constructor
		let parentClassConstructor = SheetStorageAndStyles;
		parentClassConstructor.call(this);
		return this;
	}
	// inherit parent class methods
	// https://developer.mozilla.org/ru/docs/Web/JavaScript/Reference/Global_Objects/Object/create#%D0%BF%D1%80%D0%B8%D0%BC%D0%B5%D1%80_%D0%BA%D0%BB%D0%B0%D1%81%D1%81%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%BE%D0%B5_%D0%BD%D0%B0%D1%81%D0%BB%D0%B5%D0%B4%D0%BE%D0%B2%D0%B0%D0%BD%D0%B8%D0%B5_%D1%81_object.create
	DocumentSheet_Type.prototype = Object.create(SheetStorageAndStyles.prototype);
	DocumentSheet_Type.prototype.constructor = DocumentSheet_Type;

	/**
	 * // Docs old:
	 *    // StyleSheet_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/stylesheet_type-complextypevisio-xml
	 *    //In new schema inherites from ShapeSheet_Type
	 * @returns {StyleSheet_Type}
	 * @constructor
	 * @extends SheetStorageAndStyles
	 */
	function StyleSheet_Type() {
		this.id = null;
		this.name = null;
		this.nameU = null;
		this.isCustomName = null;
		this.isCustomNameU = null;

		/**
		 * see MS-VSDX 2.2.7.4.9	Connector.
		 * @type {boolean}
		 */
		this.isConnectorStyleIherited = false;

		// call parent class constructor
		let parentClassConstructor = SheetStorageAndStyles;
		parentClassConstructor.call(this);
		return this;
	}
	// inherit parent class methods
	// https://developer.mozilla.org/ru/docs/Web/JavaScript/Reference/Global_Objects/Object/create#%D0%BF%D1%80%D0%B8%D0%BC%D0%B5%D1%80_%D0%BA%D0%BB%D0%B0%D1%81%D1%81%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%BE%D0%B5_%D0%BD%D0%B0%D1%81%D0%BB%D0%B5%D0%B4%D0%BE%D0%B2%D0%B0%D0%BD%D0%B8%D0%B5_%D1%81_object.create
	StyleSheet_Type.prototype = Object.create(SheetStorageAndStyles.prototype);
	StyleSheet_Type.prototype.constructor = StyleSheet_Type;

	/**
	 * // Docs old:
	 *    // Элемент Shapes (ShapeSheet_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/shapes-element-shapesheet_type-complextypevisio-xml
	 *    // ShapeSheet_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/shapesheet_type-complextypevisio-xml
	 * @returns {ShapeSheet_Type}
	 * @constructor
	 * @extends SheetStorageAndStyles
	 */
	function ShapeSheet_Type() {
		this.id = null;
		this.originalID = null;
		this.del = null;
		this.masterShape = null;
		this.uniqueID = null;
		this.name = null;
		this.nameU = null;
		this.isCustomName = null;
		this.isCustomNameU = null;
		this.master = null;
		this.type = null;

		this.shapes = [];
		this.inheritedShapes = [];

		this.items = null;
		this.anyAttr = null;

		// call parent class constructor
		let parentClassConstructor = SheetStorageAndStyles;
		parentClassConstructor.call(this);
		return this;
	}
	// inherit parent class methods
	// https://developer.mozilla.org/ru/docs/Web/JavaScript/Reference/Global_Objects/Object/create#%D0%BF%D1%80%D0%B8%D0%BC%D0%B5%D1%80_%D0%BA%D0%BB%D0%B0%D1%81%D1%81%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%BE%D0%B5_%D0%BD%D0%B0%D1%81%D0%BB%D0%B5%D0%B4%D0%BE%D0%B2%D0%B0%D0%BD%D0%B8%D0%B5_%D1%81_object.create
	ShapeSheet_Type.prototype = Object.create(SheetStorageAndStyles.prototype);
	ShapeSheet_Type.prototype.constructor = ShapeSheet_Type;

	// Docs old:
// Icon_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/icon_type-complextypevisio-xml
	function Icon_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);

		this.value = null;
		return this;
	}
	AscFormat.InitClass(Icon_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);


	//todo move to commons
	function PropLocker(objectId)
	{
		this.objectId = null;
		this.Lock = new AscCommon.CLock();
		this.Id = AscCommon.g_oIdCounter.Get_NewId();
		g_oTableId.Add(this, this.Id);

		if(typeof  objectId === "string")
		{
			this.setObjectId(objectId);
		}

	}

	PropLocker.prototype = {

		getObjectType: function()
		{
			return AscDFH.historyitem_type_PropLocker;
		},
		setObjectId: function(id)
		{
			//todo
			//History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_PropLockerSetId, this.objectId, id));
			this.objectId = id;
		},
		Get_Id: function()
		{
			return this.Id;
		},
		Write_ToBinary2: function(w)
		{
			w.WriteLong(AscDFH.historyitem_type_PropLocker);
			w.WriteString2(this.Id);
		},

		Read_FromBinary2: function(r)
		{
			this.Id = r.GetString2();
		},

		Refresh_RecalcData: function()
		{}

	};

	//-------------------------------------------------------------export---------------------------------------------------
	window['Asc']            = window['Asc'] || {};
	window['AscCommon']      = window['AscCommon'] || {};
	window['AscCommonWord']  = window['AscCommonWord'] || {};
	window['AscCommonSlide'] = window['AscCommonSlide'] || {};
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscVisio']  = window['AscVisio'] || {};
	window['AscFormat']  = window['AscFormat'] || {};
	window['AscWord'] = window['AscWord'] || {};


	window['AscVisio'].c_oVsdxSheetStorageKind = c_oVsdxSheetStorageKind;
	window['AscVisio'].SheetStorageAndStyles = SheetStorageAndStyles;
	window['AscVisio'].Text_Type = Text_Type;
	window['AscVisio'].Data_Type = Data_Type;
	window['AscVisio'].ForeignData_Type = ForeignData_Type;
	window['AscVisio'].Trigger_Type = Trigger_Type;
	window['AscVisio'].Row_Type = Row_Type;
	window['AscVisio'].Cell_Type = Cell_Type;
	window['AscVisio'].Shape_Type = Shape_Type;
	window['AscVisio'].Section_Type = Section_Type;
	window['AscVisio'].Page_Type = Page_Type;
	window['AscVisio'].PageSheet_Type = PageSheet_Type;
	window['AscVisio'].DocumentSheet_Type = DocumentSheet_Type;
	window['AscVisio'].StyleSheet_Type = StyleSheet_Type;
	window['AscVisio'].ShapeSheet_Type = ShapeSheet_Type;
	window['AscVisio'].Icon_Type = Icon_Type;
	window['AscVisio'].PropLocker = PropLocker;
	window['AscVisio'].createKeyFromSheetObject = createKeyFromSheetObject;

})(window, window.document);

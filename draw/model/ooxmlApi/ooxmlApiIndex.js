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

	/**
	 *    // Docs old:
	 *    // Useless see ShapeSheet_Type - old
	 *    // Inherites(extends) ShapeSheet_Type in new schema
	 * @return {Shape_Type}
	 * @constructor
	 */
	function Shape_Type() {
		this.iD = null;
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

		setShapeSheetClassMembers.call(this);
		return this;
	}


	/**
	 *    // Docs old:
	 *    // Section_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/section_type-complextypevisio-xml
	 * @return {Section_Type}
	 * @constructor
	 */
	function Section_Type() {
		this.n = null;
		this.del = null;
		this.iX = null;

		// always use getter setter methods
		// Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
		this.cells = [];
		this.triggers = [];
		this.rows = [];
		return this;
	}

	/**
	 * Docs old:
	 * Элемент RefBy (Cell_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/refby-element-cell_type-complextypevisio-xml
	 * Cell_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/cell_type-complextypevisio-xml
	 * @return {Cell_Type}
	 * @constructor
	 */
	function Cell_Type() {
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


	/**
	 *    // Docs old:
	 *    // Row_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/row_type-complextypevisio-xml
	 * @return {Row_Type}
	 * @constructor
	 */
	function Row_Type() {
		this.n = null;
		this.localName = null;
		this.iX = null;
		this.t = null;
		this.del = null;

		this.cells = [];
		this.triggers = [];
		return this;
	}


	Shape_Type.prototype.getMasterID = function() {
		return this.master;
	}


	/**
	 *
	 * @param obj
	 * @param {string} constructorName
	 * @param {string} [attributeName]
	 * @param attributeValue
	 * @return {*|undefined}
	 */
	function findObject(obj, constructorName, attributeName, attributeValue) {
		let attributeCheck = true;
		if (attributeName === undefined) {
			attributeCheck = false;
		}
		// Base case: if the object is null or undefined, or if it's not an object
		if (!obj || typeof obj !== 'object') {
			return null;
		}
		// Check if the current object has the desired attribute, value, and constructor name
		if (obj.constructor.name === constructorName && (attributeCheck ? obj[attributeName] === attributeValue : true)) {
			return obj;
		}
		// Iterate over object properties and recursively search for the attribute and constructor name
		for (const key in obj) {
			if (obj.hasOwnProperty(key)) {
				const result = findObject(obj[key], constructorName, attributeName, attributeValue);
				if (result) {
					return result;
				}
			}
		}
		// If the attribute was not found, return undefined
		return undefined;
	}

	function findObjects(obj, constructorName, attributeName, attributeValue, results) {
		if (typeof results === "undefined") {
			results = [];
		}
		// Base case: if the object is null or undefined, or if it's not an object
		if (!obj || typeof obj !== 'object') {
			return [];
		}
		// Check if the current object has the desired attribute, value, and constructor name
		if (obj.constructor.name === constructorName && obj[attributeName] === attributeValue) {
			results.push(obj);
		}
		// Iterate over object properties and recursively search for the attribute and constructor name
		for (const key in obj) {
			if (obj.hasOwnProperty(key)) {
				findObjects(obj[key], constructorName, attributeName, attributeValue, results);
			}
		}
		return results;
	}

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

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * Finds shape section by formula. Compares N with string argument. For Geometry use find sections.
	 * @param {String} formula
	 * @memberof Shape_Type
	 * @returns {Section_Type | null}
	 */
	Shape_Type.prototype.getSection = function getSection(formula) {
		return findObject(this.elements, "Section_Type", "n", formula);
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * @param {String} formula
	 * @memberof Shape_Type
	 * @returns {Row_Type | null}
	 */
	Shape_Type.prototype.getRow = function findRow(formula) {
		throw new Error("not realized");
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
	 * but Row for example can be found by IX N
	 * @param {String} formula
	 * @memberof Shape_Type
	 * @returns {Cell_Type|null}
	 */
	Shape_Type.prototype.getCell = function findCell(formula) {
		// Cells can have N only no IX
		return findObject(this.elements, "Cell_Type", "n", formula);
	}

	/**
	 * Calls getCell on Shape and tries to parse as Number(cell.v) if cell exists.
	 * @param {String} formula
	 * @return {?Number} number
	 */
	Shape_Type.prototype.getCellNumberValue = function (formula) {
		let cell = this.getCell(formula);
		if (cell !== undefined) {
			return Number(cell.v);
		} else {
			return undefined;
		}
	}

	/**
	 * Calls getCell on Shape and tries to parse as String(cell.v) if cell exists.
	 * @param {String} formula
	 * @return {?String} string
	 */
	Shape_Type.prototype.getCellStringValue = function (formula) {
		let cell = this.getCell(formula);
		if (cell !== undefined) {
			return String(cell.v);
		} else {
			return undefined;
		}
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
		return findObject(this.elements, "Text_Type");
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * if in formula we have both ix and n we should use findSection instead.
	 * or if we use it with number in formula
	 * @param {String} formula
	 * @memberof Shape_Type
	 * @returns {Section_Type[] | null}
	 */
	Shape_Type.prototype.getSections = function(formula) {
		if (/^\d+$/.test(formula)) {
			// if number
			console.log('strange findSections use (with number)');
			return findObjects(this.elements, "Section_Type", "iX", formula);
		}
		return findObjects(this.elements, "Section_Type", "n", formula);
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * @memberof Shape_Type
	 * @return {Shape_Type[]}
	 */
	Shape_Type.prototype.getSubshapes = function () {
		return this.shapes;
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * get elements inherited from shape sheet type
	 * @return {*}
	 */
	Shape_Type.prototype.getElements = function () {
		return this.elements;
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 *
	 * Cells in Section (section is like table in visio) can exist by itself, directy in section like here
	 * https://disk.yandex.ru/d/Ud6-wmVjNnOyUA
	 * with their names like Width, Height, Angle (for Shape Thansform section)
	 * or they can be in Rows like here
	 * https://disk.yandex.ru/i/4ASd_5KHYIlXKw
	 * with names X, Y, A, B (for Geometry SplineStart Row).
	 *
	 * Lets search cells only directly in Section for now.
	 *
	 * @param {String} formula
	 * @memberof Section_Type
	 * @returns {Cell_Type | null}
	 */
	Section_Type.prototype.getCell = function (formula) {
		// Cells can have N only no IX
		return findObject(this.cells, "Cell_Type", "n", formula);
	}

	/**
	 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
	 * Finds row in sections. Returns link to object not copy.
	 * @param {String | Number} formula
	 * @memberof Section_Type
	 * @returns {null | Row_Type}
	 */
	Section_Type.prototype.getRow = function (formula) {
		if (typeof formula === "number" || /^\d+$/.test(formula)) {
			return findObject(this.rows, "Row_Type", "iX", Number(formula));
		}
		return findObject(this.rows, "Row_Type", "n", formula);
	}


	/**
	 * returns this shape and subshapes array without cloning so objects are linked.
	 * @param {boolean} [recalculatePinCords] if true changes subshapes cords from parent shape relative to absolute
	 * @param [resultArray = []]
	 * @memberof Shape_Type
	 * @return {Shape_Type[]}
	 */
	Shape_Type.prototype.collectSubshapesRecursive = function(recalculatePinCords, resultArray) {
		if (resultArray === undefined) {
			resultArray = [];
		}

		if (recalculatePinCords === undefined || recalculatePinCords === null) {
			recalculatePinCords = false;
		}

		resultArray.push(this);

		let subShapes = this.getSubshapes();
		for (let j = 0; j < subShapes.length; j++) {
			const subShape = subShapes[j];

			if (recalculatePinCords) {
				let parentShapePinX = Number(this.getCell("PinX").v);
				let parentShapePinY = Number(this.getCell("PinY").v);
				let parentShapeLocPinX = Number(this.getCell("LocPinX").v);
				let parentShapeLocPinY = Number(this.getCell("LocPinY").v);

				let subShapePinX = subShape.getCell("PinX");
				let subShapePinY = subShape.getCell("PinY");

				// there was case with shape type group with no PinX and PinY
				if (subShapePinX !== null) {
					let sub_shape_pinX_inch = Number(subShapePinX.v);
					subShapePinX.v = parentShapePinX - parentShapeLocPinX + sub_shape_pinX_inch;
				}

				if (subShapePinY !== null) {
					let sub_shape_pinY_inch = Number(subShapePinY.v);
					subShapePinY.v = parentShapePinY - parentShapeLocPinY + sub_shape_pinY_inch;
				}
			}

			subShape.collectSubshapesRecursive(recalculatePinCords, resultArray);
		}

		return resultArray;
	}

	/**
	 * Realizes Master-To-Shape inheritance.
	 * Comes through the shape recursively through all subshapes
	 * and copies properties from their masters if master exist.
	 * Uses clone function so copied objects are unlinked from master.
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
				return masterObject.iD === topShapeMasterId;
			});
			let topShapeMaster = masters[topShapeMasterIndex];

			let masterShapes = topShapeMaster.content.shapes;
			masterShapesToInheritFrom = masterShapes;

			// all descendant shapes will inherit from that master
			ancestorMasterShapes = masterShapesToInheritFrom;
		}

		// check MasterShape attribute and set shapes/shape
		// to inherit from: masterShapesToInheritFrom and ancestorMasterShapes
		let masterShapeId = this.masterShape;
		if (masterShapeId !== null && masterShapeId !== undefined) {
			if (ancestorMasterShapes === null || ancestorMasterShapes === undefined) {
				console.log("MasterShape attribute is set but Master is not set for ", this);
			} else {
				let masterIndex = -1;
				if (ancestorMasterShapes.length === 1) {
					// if master has one top level shape
					let masterSubshapes = ancestorMasterShapes[0].collectSubshapesRecursive(false);
					masterIndex = masterSubshapes.findIndex(function (masterSubshape) {
						return masterShapeId === masterSubshape.iD;
					});
					let masterShape = masterSubshapes[masterIndex];
					masterShapesToInheritFrom = [masterShape];
				} else {
					let masterSubshapes = [];
					ancestorMasterShapes.forEach(function(ancestorMasterShape) {
						let masterSubshapesNth = ancestorMasterShape.collectSubshapesRecursive(false);
						masterSubshapes = masterSubshapes.concat(masterSubshapesNth);
					})
					masterIndex = masterSubshapes.findIndex(function (masterSubshape) {
						return masterShapeId === masterSubshape.iD;
					});
					let masterShape = masterSubshapes[masterIndex];
					masterShapesToInheritFrom = [masterShape];
				}

				if (masterIndex === -1) {
					console.log('For MasterShape = ', masterShapeId, 'shape not found in master. Check shape: ', this);
				}
			}
		}

		// inherit: mergeElements and clone shapes to which elements will be merged on recursive calls
		if (masterShapesToInheritFrom.length === 1) {
			let masterShapeToInheritFrom = masterShapesToInheritFrom[0];

			// inherit link to styles
			if (!this.lineStyle) {
				this.lineStyle = masterShapeToInheritFrom.lineStyle;
			}
			if (!this.fillStyle) {
				this.fillStyle = masterShapeToInheritFrom.fillStyle;
			}
			if (!this.textStyle) {
				this.textStyle = masterShapeToInheritFrom.textStyle;
			}

			let shapeElements = this.elements;
			let masterElements = masterShapeToInheritFrom.elements;
			mergeElementArrays(shapeElements, masterElements);

			let shapeSubshapes = this.shapes;
			let masterSubshapes = masterShapeToInheritFrom.shapes;
			cloneSubshapes(shapeSubshapes, masterSubshapes, masters);
		} else if (masterShapesToInheritFrom.length > 1) {
			// does it ever happens?
			// what about style inheritance?
			cloneSubshapes(this.shapes, masterShapesToInheritFrom, masters);
		}

		// call recursive on all subshapes
		let subshapes = this.shapes;
		subshapes.forEach(function(shape) {
			shape.realizeMasterInheritanceRecursively(masters, ancestorMasterShapes);
		});

		// return thisShapeCopy;
		// end of method

		/**
		 * clones masters shapes to given shape.
		 * Uses MasterShapeAttributes to find shapes to insert.
		 * @param {Shape_Type[]} shapeSubshapes
		 * @param {Shape_Type[]} masterSubshapes
		 * @param masters - result from joinMastersInfoAndContents()
		 */
		function cloneSubshapes(shapeSubshapes, masterSubshapes, masters) {
			// If subshape has Master attribute with id of any master: call realizeMasterToShapeInheritance
			// If subshape has MasterShape attribute with id of any parents shape masters subshapes:
			// 	call mergeElementArrays NO RECURSION HERE
			// If there is a shape in master but there is no such local subshape then it should be inherited (copied)
			// lets check if it exists locally only by MasterShape attribute
			// examples it the bottom of the function

			// handle subshapes MasterShape attribute
			masterSubshapes.forEach(function(masterSubshape) {
				let mergeElementIndex = findIndexComparingByMasterShapeAttribute(shapeSubshapes, masterSubshape);
				let elementExistsAlready = mergeElementIndex !== -1;

				// 2.2.5.4.1	Master-to-Shape Inheritance
				// "subshapes not specified in the instance are inherited from the master." (from its master)
				if (!elementExistsAlready) {
					// maybe add masterShape attribute to new shape - lets dont do it because:
					// of recursive iterations of inheritance we will try to inherit because we will se masterShape but
					// there is no need because it is copy pasted element no need in inheritance
					// maybe consider id to insert in ascending order
					shapeSubshapes.push(clone(masterSubshape));
				} else {
					// 2.2.5.4.1	Master-to-Shape Inheritance
					// "if an instance contains a subshape whose ShapeSheet_Type element has a MasterShape attribute that matches
					// the ID attribute of a subshape of the master, the local properties specified in this subshape will
					// override those of the corresponding subshape in the master."

					// let masterElements = masterSubshape.elements;
					// let shapeElements = shapeSubshapes[mergeElementIndex].elements;
					// mergeElementArrays(shapeElements, masterElements);

					// it is done in realizeMasterToShapeInheritanceRecursive with subshapes handle
				}
			});

			// handle subshapes Master attributes
			// shapeSubshapes.forEach(function(subShape) {
			// 	subShape.realizeMasterToShapeInheritanceRecursive(masters);
			// });
			// UPD: in realizeMasterToShapeInheritanceRecursive
		}

		function findIndexComparingByMasterShapeAttribute(shapeSubshapes, masterSubshape) {
			return shapeSubshapes.findIndex(function (element) {
				return element.masterShape === masterSubshape.iD;
			});
		}
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
			// console.log("style has realized inheritance already. return");
			return;
		}

		if (!(thisArgument.lineStyle === thisArgument.fillStyle && thisArgument.lineStyle === thisArgument.textStyle)) {
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
					"Character", "Paragraph", "Tabs"];

			let commonElements = ["ColorSchemeIndex", "EffectSchemeIndex", "ConnectorSchemeIndex", "FontSchemeIndex",
					"ThemeIndex", "VariationColorIndex", "VariationStyleIndex", "EmbellishmentIndex",
			"QuickStyleLineColor", "QuickStyleFillColor", "QuickStyleShadowColor", "QuickStyleFontColor",
			"QuickStyleLineMatrix", "QuickStyleFillMatrix", "QuickStyleEffectsMatrix", "QuickStyleFontMatrix",
			"QuickStyleType", "QuickStyleVariation"];
			lineStyleElements = lineStyleElements.concat(commonElements);
			fillStyleElements = fillStyleElements.concat(commonElements);
			textStyleElements = textStyleElements.concat(commonElements);

			if (thisArgument.lineStyle !== null) {
				let styleId = Number(thisArgument.lineStyle);
				let styleSheet = styles.find(function(style) {
					return style.iD === styleId;
				});
				realizeStyleToSheetObjInheritanceRecursive(styleSheet, styles, stylesWithRealizedInheritance);
				mergeElementArrays(thisArgument.elements, styleSheet.elements, lineStyleElements);
			}

			if (thisArgument.fillStyle !== null) {
				let styleId = Number(thisArgument.fillStyle);
				let styleSheet = styles.find(function(style) {
					return style.iD === styleId;
				});
				realizeStyleToSheetObjInheritanceRecursive(styleSheet, styles, stylesWithRealizedInheritance);
				mergeElementArrays(thisArgument.elements, styleSheet.elements, fillStyleElements);
			}

			if (thisArgument.textStyle !== null) {
				let styleId = Number(thisArgument.textStyle);
				let styleSheet = styles.find(function(style) {
					return style.iD === styleId;
				});
				realizeStyleToSheetObjInheritanceRecursive(styleSheet, styles, stylesWithRealizedInheritance);
				mergeElementArrays(thisArgument.elements, styleSheet.elements, textStyleElements);
			}

			if (thisArgument.constructor.name === "StyleSheet_Type") {
				// memorize: that style has realized inheritance
				stylesWithRealizedInheritance.add(thisArgument);
			}

			return;
		}

		if (thisArgument.lineStyle === null && thisArgument.fillStyle === null && thisArgument.textStyle === null) {
			// console.log('Top parent style');
			return;
		}

		let styleId = Number(thisArgument.lineStyle);
		let styleSheet = styles.find(function(style) {
			return style.iD === styleId;
		});

		realizeStyleToSheetObjInheritanceRecursive(styleSheet, styles, stylesWithRealizedInheritance);

		mergeElementArrays(thisArgument.elements, styleSheet.elements)

		if (thisArgument.constructor.name === "StyleSheet_Type") {
			// memorize: that style has realized inheritance
			stylesWithRealizedInheritance.add(thisArgument);
		}
	}

	/**
	 * Style-To-Shape inheritance
	 * Copy all style elements (sections, rows, cells) to shape.
	 * (Doesnt take much memory < 300MB with master inheritance for the most large files).
	 * stylesWithRealizedInheritance was added for optimization.
	 * @param styles
	 * @param {?Set} [stylesWithRealizedInheritance]
	 */
	Shape_Type.prototype.realizeStyleInheritanceRecursively = function(styles, stylesWithRealizedInheritance) {
		if (stylesWithRealizedInheritance === undefined) {
			stylesWithRealizedInheritance = new Set();
		}
		realizeStyleToSheetObjInheritanceRecursive(this, styles, stylesWithRealizedInheritance);

		// call recursive on all subshapes
		let subshapes = this.shapes;
		subshapes.forEach(function(shape) {
			shape.realizeStyleInheritanceRecursively(styles, stylesWithRealizedInheritance);
		});
	}

	/**
	 * clone master elements (sections, rows, cells) to shapeElements.
	 * For Sections and Rows merge is recursive: we compare inner cells by their names
	 * @param shapeElements - cells rows sections
	 * @param masterElements - cells rows sections
	 * @param {string[]?} elementsToMerge - cells rows sections list we can merge
	 * @param {boolean?} isParentInList
	 */
	function mergeElementArrays(shapeElements, masterElements, elementsToMerge, isParentInList) {
		/**
		 * find index of cell row or section
		 * @param elementsArray
		 * @param elementToFind
		 * @returns {*}
		 */
		function findIndexComparingByNorIX(elementsArray, elementToFind) {
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
			return elementsArray.findIndex(function (element) {
				let nAttributeExists = element.n !== null && typeof element.n !== 'undefined';
				let iXAttributeExists = element.iX !== null && typeof element.iX !== 'undefined';

				if (nAttributeExists && iXAttributeExists) {
					return element.iX == elementToFind.iX && element.n === elementToFind.n;
				}
				if (nAttributeExists) {
					return element.n === elementToFind.n;
				}
				if (iXAttributeExists) {
					return element.iX === elementToFind.iX;
				}
				// if smth wrong with nAttributeExists or iXAttributeExists
				return false;
			});
		}

		let mergeAll = false;

		if (elementsToMerge === undefined) {
			mergeAll = true;
		}

		masterElements.forEach(function(masterElement) {
			let mergeElementIndex = findIndexComparingByNorIX(shapeElements, masterElement);
			let elementExistsAlready = mergeElementIndex !== -1;

			let elementIsInList = elementsToMerge !== undefined && elementsToMerge.includes(masterElement.n);
			let listCheck = elementIsInList || isParentInList || mergeAll;

			if (!elementExistsAlready) {
				if (listCheck) {
					// TODO fix order
					// mb lets not add cell after section
					let elementCopy = clone(masterElement);
					shapeElements.push(elementCopy);
					// ix wrong order causes problems
					shapeElements.sort(function (a, b) {
						return a.iX - b.iX;
					});
				}
			} else {
				// merge inner elements recursive if not cell
				if (masterElement.constructor.name !== 'Cell_Type') {
					// if Section or Row
					let shapeElement = shapeElements[mergeElementIndex];
					if (masterElement.constructor.name == 'Section_Type') {
						// for future checks
						isParentInList = elementIsInList || isParentInList;
						// recursive calls
						mergeElementArrays(shapeElement.cells, masterElement.cells, elementsToMerge, isParentInList);
						mergeElementArrays(shapeElement.rows, masterElement.rows, elementsToMerge, isParentInList);
					} else if (masterElement.constructor.name == 'Row_Type') {
						// for future checks
						isParentInList = elementIsInList || isParentInList;
						// recursive calls
						mergeElementArrays(shapeElement.cells, masterElement.cells, elementsToMerge, isParentInList);
					}
				}
			}
		});
	}


	/**
	 * get deep copy of object with prototypes
	 * @param object
	 * @return {any}
	 */
	function clone(object) {
		function cloneWithPrototypesRecursive(copyObject, originalObject) {
			Object.setPrototypeOf(copyObject, Object.getPrototypeOf(originalObject));

			// Iterate over object properties and recursively
			for (const key in originalObject) {
				if (originalObject.hasOwnProperty(key)) {
					if (typeof originalObject[key] === 'object' && originalObject[key] !== null) {
						// so after line below when we set array props using copyObject[key] = originalObject[key];
						// array length will change automatically
						copyObject[key] = Array.isArray(originalObject[key]) ? [] : Object.create(null);
						cloneWithPrototypesRecursive(copyObject[key], originalObject[key]);
					} else {
						copyObject[key] = originalObject[key];
					}
				}
			}
		}

		let copy = Object.create(null);
		// let copy = JSON.parse(JSON.stringify(object));
		cloneWithPrototypesRecursive(copy, object);
		// console.log("Clone function test.\nObject before: ", object, "\nAfter: ", copy);
		return copy;
	}


	/**
	 * returns link not clone
	 * @param {String} formula
	 * @memberof Row_Type
	 * @returns {null | Cell_Type}
	 */
	Row_Type.prototype.getCell = function (formula) {
		// Cells can have N only no IX
		return findObject(this.cells, "Cell_Type", "n", formula);
	}


	/**
	 * @memberOf Cell_Type
	 * @return {number}
	 */
	Cell_Type.prototype.getValueInMM = function () {
		let res;
		//todo all units
		switch (this.u) {
			case "DL":
			case "IN":
			case "IN_F":
				res = parseFloat(this.v) * g_dKoef_in_to_mm;
				break;
			case "FT":
				res = parseFloat(this.v) * 12 * g_dKoef_in_to_mm;
				break;
			case "F_I":
				res = parseFloat(this.v);
				let intPart = Math.floor(res);
				res = (intPart * 12 + (res - intPart)) * g_dKoef_in_to_mm;
				break;
			case "KM":
				res = parseFloat(this.v) * 1000000;
				break;
			case "M":
				res = parseFloat(this.v) * 1000;
				break;
			case "CM":
				res = parseFloat(this.v) * 10;
				break;
			case "MM":
				res = parseFloat(this.v);
				break;
			default:
				res = parseFloat(this.v) * g_dKoef_in_to_mm;
				break;
		}
		return res;
	};
	/**
	 * @memberOf Cell_Type
	 * @return {number}
	 */
	Cell_Type.prototype.getValueInInch = function () {
		let res = this.getValueInMM() / g_dKoef_in_to_mm;
		return res;
	}

	// for ooxml classes
	function setShapeSheetClassMembers() {
		// 3 attr below inherited from Sheet_Type using old schema
		// or from ShapeSheet_Type using new schema
		this.lineStyle = null;
		this.fillStyle = null;
		this.textStyle = null;
		/**
		 * When working with shapes use getElements/setElements methods
		 * Always use it see Shape_Type.prototype.realizeMasterToShapeInheritanceRecursive js docs for explanation.
		 * elements is used bcs text can appear here maybe
		 * @type {*[]}
		 */
		this.elements = [];
		// elements below are stored in elements to support new schema

		// // 3 arrays below inherited from Sheet_Type
		// this.cells = [];
		// this.triggers = [];
		// this.sections = [];
		//
		//
		// // new attributes inherited from ShapeSheet_Type
		// this.text = null;
		// this.data1 = null;
		// this.data2 = null;
		// this.data3 = null;
		// this.foreignData = null;
	}

	//-------------------------------------------------------------export---------------------------------------------------
	window['Asc']            = window['Asc'] || {};
	window['AscCommon']      = window['AscCommon'] || {};
	window['AscCommonWord']  = window['AscCommonWord'] || {};
	window['AscCommonSlide'] = window['AscCommonSlide'] || {};
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonDraw']  = window['AscCommonDraw'] || {};
	window['AscFormat']  = window['AscFormat'] || {};
	window['AscWord'] = window['AscWord'] || {};

	window['AscCommonDraw'].Row_Type = Row_Type;
	window['AscCommonDraw'].Cell_Type = Cell_Type;
	window['AscCommonDraw'].Shape_Type = Shape_Type;
	window['AscCommonDraw'].Section_Type = Section_Type;

	window['AscCommonDraw'].setShapeSheetClassMembers = setShapeSheetClassMembers;

})(window, window.document);

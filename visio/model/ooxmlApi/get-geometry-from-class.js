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


(function geometryFromClassApi(window, document) {
	// works with visio class
	// NOT FINISHED

	const radToC = AscFormat.radToDeg * AscFormat.degToC;

	function getRandomPrst() {
		let types = AscCommon.g_oAutoShapesTypes[Math.floor(Math.random()*AscCommon.g_oAutoShapesTypes.length)];
		return types[Math.floor(Math.random()*types.length)].Type;
	}

	/**
	 *
	 * @param {number} mm
	 * @returns {number} emus
	 */
	function mmToEmu(mm) {
		// Conversion factor: 1 cm = 360000 EMUs, 1 cm = 10 mm
		const emusPerCm = 360000;
		const mmPerCm = 10;

		// Calculate EMUs from millimeters using the new conversion factor
		const emus = mm * (emusPerCm / mmPerCm);
		return emus;
	}

	/**
	 * Calls mmToEmu on value:
	 * value * additionalUnitKoef then convert to Emus
	 * @param {number} value - mmUnits value
	 * @param {number} additionalUnitKoef
	 * @returns {number} valueCorrectUnits
	 */
	function convertUnits(value, additionalUnitKoef) {
		let valueCorrectUnits = mmToEmu(value * additionalUnitKoef);
		return valueCorrectUnits;
	}

	/**
	 * afin rotate clockwise
	 * @param {number} x
	 * @param {number} y
	 * @param {number} radiansRotateAngle radians Rotate AntiClockWise Angle. E.g. 30 degrees rotates does DOWN.
	 * @returns {{x: number, y: number}} point
	 */
	function rotatePointAroundCordsStartClockWise(x, y, radiansRotateAngle) {
		let newX = x * Math.cos(radiansRotateAngle) + y * Math.sin(radiansRotateAngle);
		let newY = x * (-1) * Math.sin(radiansRotateAngle) + y * Math.cos(radiansRotateAngle);
		return {x : newX, y: newY};
	}

	/**
	 * get Geometry object from shape object reading shape elements
	 * @param {Shape_Type} shape
	 * @param {number} pageScale
	 * @param {number} isInvertCoords
	 * @returns {Geometry} geometry
	 */
	function getGeometryFromShape(shape, pageScale, isInvertCoords) {
		// init geometry
		let geometry = new AscFormat.Geometry();

		// in visio Geometry section represents Path:
		// A path is a collection of vertices and line or curve segments that specifies an enclosed area.
		// The geometry of a shape is specified by a collection of paths.
		// Each Geometry (Section_Type) element specifies a path.
		let geometrySections = shape.getSections("Geometry");

		// NEW VERSION: MAKE ONE TWO PATHS PER SHAPE:
		// One path for path with strokes for them to subtract but:
		// - it may not work well with more than 2 geometry path
		// - we cant set different setStroke for geometry paths now because it is one path now
		// see test folder https://disk.yandex.ru/d/_MQCkWbngkSz8w
		// 25.12.2023 commit
		let pathWithFill = new AscFormat.Path();
		let pathWithoutFill = new AscFormat.Path();

		pathWithFill.setExtrusionOk(false);
		pathWithFill.setFill("norm");
		pathWithFill.setStroke(true);
		pathWithFill.setPathW(undefined);
		pathWithFill.setPathH(undefined);

		pathWithoutFill.setExtrusionOk(false);
		pathWithoutFill.setFill("none");
		pathWithoutFill.setStroke(true);
		pathWithoutFill.setPathW(undefined);
		pathWithoutFill.setPathH(undefined);

		// for path overlap fix: If there is only equal geometry section NoLine values
		// we can set true NoLine value for path
		let filledPathNoLine = false;
		let unfilledPathNoLine = false;
		let allFilledPathNoLineValuesEqual = true;
		let allUnfilledPathNoLineValuesEqual = true;

		// set path objects - parts of geometry objects
		for (let i = 0; i < geometrySections.length; i++) {
			const geometrySection = geometrySections[i];
			// see [MS-VSDX]-220215 2.2.3.2.2.Geometry Path

			//	<xsd:complexType name="Geometry_Type">
			// 		<xsd:complexContent>
			// 			<xsd:extension base="Section_Type">
			// 				<xsd:choice minOccurs="0" maxOccurs="unbounded">
			// 					<xsd:element name="Cell" type="Cell_Type" minOccurs="0" maxOccurs="unbounded">
			// 						<xsd:alternative test="@N = 'NoFill'" type="NoFill_Type"/>
			// 						<xsd:alternative test="@N = 'NoLine'" type="NoLine_Type"/>
			// 						<xsd:alternative test="@N = 'NoShow'" type="NoShow_Type"/>
			// 						<xsd:alternative test="@N = 'NoSnap'" type="NoSnap_Type"/>
			// 						<xsd:alternative test="@N = 'NoQuickDrag'" type="NoQuickDrag_Type"/>
			// 						<xsd:alternative test="@N = 'Path'" type="Path_Type"/>
			// 					</xsd:element>
			// 					<xsd:element name="Row" type="GeometryRow_Type" minOccurs="0" maxOccurs="unbounded">
			// 						<xsd:alternative test="@T = 'MoveTo'" type="MoveTo_Type"/>
			// 						<xsd:alternative test="@T = 'RelMoveTo'" type="RelMoveTo_Type"/>
			// 						<xsd:alternative test="@T = 'LineTo'" type="LineTo_Type"/>
			// 						<xsd:alternative test="@T = 'RelLineTo'" type="RelLineTo_Type"/>
			// 						<xsd:alternative test="@T = 'ArcTo'" type="ArcTo_Type"/>
			// 						<xsd:alternative test="@T = 'InfiniteLine'" type="InfiniteLine_Type"/>
			// 						<xsd:alternative test="@T = 'Ellipse'" type="Ellipse_Type"/>
			// 						<xsd:alternative test="@T = 'EllipticalArcTo'" type="EllipticalArcTo_Type"/>
			// 						<xsd:alternative test="@T = 'RelEllipticalArcTo'" type="RelEllipticalArcTo_Type"/>
			// 						<xsd:alternative test="@T = 'SplineStart'" type="SplineStart_Type"/>
			// 						<xsd:alternative test="@T = 'SplineKnot'" type="SplineKnot_Type"/>
			// 						<xsd:alternative test="@T = 'PolylineTo'" type="PolylineTo_Type"/>
			// 						<xsd:alternative test="@T = 'NURBSTo'" type="NURBSTo_Type"/>
			// 						<xsd:alternative test="@T = 'RelCubBezTo'" type="RelCubBezTo_Type"/>
			// 						<xsd:alternative test="@T = 'RelQuadBezTo'" type="RelQuadBezTo_Type"/>
			// 					</xsd:element>
			// 				</xsd:choice>
			// 			</xsd:extension>
			// 		</xsd:complexContent>
			// 	</xsd:complexType>

			// The visibility of the path’s line and the visibility of the path’s fill are specified, respectively,
			// by the NoLine and NoFill Cell_Type child elements of the path’s Geometry Section_Type element.

			// While the format of the path’s line and the format of the path’s fill are specified, respectively,
			// by the line property and fill property of the shape containing the path.

			// For a path to be visible, the following conditions are necessary.
			// §	The shape containing the path is not on a layer whose Visible Cell_Type element has a value equal to zero.
			// §	The value of the NoShow Cell_Type child of the path’s Geometry Section_Type element is not equal to one.

			// NoSnap is not related to drawing.
			// NoQuickDrag is not related to drawing too. It doesn't allow us to select the shape.

			// The Path trigger is unused and MUST be ignored.

			// So Geometry section defines (for drawing): path (by rows), isShown (NoShow), NoLine, NoFill
			// So we can represent visio Geometry section using Path object
			// Use it below

			let noShowCell = geometrySection.getCell("NoShow");

			if (noShowCell === null || noShowCell === undefined) {
				AscCommon.consoleLog("noShowCell variable is null or undefined see shape: ", shape);
				continue;
			}
			if (Number(noShowCell.v) === 1) {
				continue;
			}

			let noFillCell = geometrySection.getCell("NoFill");
			let fillValue;
			if (noFillCell) {
				fillValue = Number(noFillCell.v) === 1 ? "none" : "norm";
			} else {
				fillValue = "norm";
			}

			let noLineCell = geometrySection.getCell("NoLine");
			let noLineValue = false;
			if (noLineCell) {
				noLineValue = Number(noLineCell.v) === 1;
			} else {
				noLineValue = false;
			}

			let path;

			// use one of two available path objects
			if (fillValue === "norm") {
				path = pathWithFill;
				if (i !== 0 && noLineValue !== filledPathNoLine) {
					// if new !== first
					allFilledPathNoLineValuesEqual = false;
				} else {
					filledPathNoLine = noLineValue;
				}
			} else {
				path = pathWithoutFill;
				if (i !== 0 && noLineValue !== unfilledPathNoLine) {
					allUnfilledPathNoLineValuesEqual = false;
				} else {
					unfilledPathNoLine = noLineValue;
				}
			}


			// imply that units were in mm until units parse realized
			// TODO parse formula and units
			// TODO parse line style fill style text style

			const additionalUnitCoefficient = g_dKoef_in_to_mm / pageScale;


			/* extrusionOk, fill, stroke, w, h*/
			// path.AddPathCommand(0, undefined, fillValue, undefined, undefined, undefined);

			//TODO maybe get shapeWidth and Height from outside
			//TODO shape with RelMoveTo and RelLineTo takes wrong position

			let shapeWidth = Number(shape.getCell("Width").v);
			let shapeHeight = Number(shape.getCell("Height").v);

			/**
			 *
			 * @type {{x: number, y: number}}
			 */
			let lastPoint = { x: 0, y : 0};

			/**
			 * https://learn.microsoft.com/en-us/office/client-developer/visio/splinestart-row-geometry-section
			 * fistControlPoint is taken from previous command
			 * @type {{
			 *   	firstControlPointX,
			 * 		firstControlPointY,
			 * 		secondControlPointX,
			 * 		secondControlPointY,
			 * 		secondKnot,
			 * 		firstKnot,
			 * 		lastKnot,
			 * 		degree
			 * }}
			 */
			let splineStartCommandData;
			/**
			 * https://learn.microsoft.com/en-us/office/client-developer/visio/splineknot-row-geometry-section
			 * @type {{controlPointX, controlPointY, knot}[]}
			 */
			let splineKnotCommandsData = [];
			let prevCommandName;

			let commandRows = geometrySection.getRows();

			for (let j = 0; j < commandRows.length; j++) {
				let commandRow = commandRows[j];
				if (commandRow.del) {
					continue;
				}
				let commandName = commandRow.t;

				switch (commandName) {
					case "MoveTo":
					{
						let moveToXValue = commandRow.getCellNumberValue("X", 0);
						let moveToYValue = commandRow.getCellNumberValue("Y", 0);

						if (isInvertCoords) {
							moveToYValue = shapeHeight - moveToYValue;
						}

						let newX = convertUnits(moveToXValue, additionalUnitCoefficient);
						let newY = convertUnits(moveToYValue, additionalUnitCoefficient);

						path.moveTo(newX, newY);
						lastPoint.x = newX;
						lastPoint.y = newY;
						break;
					}
					case "RelMoveTo":
					{
						let relMoveToXValue = commandRow.getCellNumberValue("X", 0);
						let relMoveToYValue = commandRow.getCellNumberValue("Y", 0);

						if (isInvertCoords) {
							relMoveToYValue = 1 - relMoveToYValue;
						}

						let newX = convertUnits(relMoveToXValue, additionalUnitCoefficient);
						let newY = convertUnits(relMoveToYValue, additionalUnitCoefficient);

						let relX = newX * shapeWidth;
						let relY = newY * shapeHeight;
						path.moveTo(relX, relY);
						lastPoint.x = relX;
						lastPoint.y = relY;
						break;
					}
					case "LineTo":
					{
						let lineToXValue = commandRow.getCellNumberValue("X", 0);
						let lineToYValue = commandRow.getCellNumberValue("Y", 0);

						if (isInvertCoords) {
							lineToYValue = shapeHeight - lineToYValue;
						}

						let newX = convertUnits(lineToXValue, additionalUnitCoefficient);
						let newY = convertUnits(lineToYValue, additionalUnitCoefficient);

						path.lnTo(newX, newY);
						lastPoint.x = newX;
						lastPoint.y = newY;
						break;
					}
					case "RelLineTo":
					{
						let relLineToXTextValue = commandRow.getCellNumberValue("X", 0);
						let relLineToYTextValue = commandRow.getCellNumberValue("Y", 0);

						if (isInvertCoords) {
							relLineToYTextValue = 1 - relLineToYTextValue;
						}

						let newX = convertUnits(relLineToXTextValue, additionalUnitCoefficient);
						let newY = convertUnits(relLineToYTextValue, additionalUnitCoefficient);

						let newXRel = newX * shapeWidth;
						let newYRel = newY * shapeHeight;
						path.lnTo(newXRel, newYRel);
						lastPoint.x = newXRel;
						lastPoint.y = newYRel;
						break;
					}
					case "EllipticalArcTo":
					{
						// https://learn.microsoft.com/en-us/office/client-developer/visio/ellipticalarcto-row-geometry-section
						let x = commandRow.getCellNumberValue("X", 0);
						let y = commandRow.getCellNumberValue("Y", 0);
						let a = commandRow.getCellNumberValue("A", 0);
						let b = commandRow.getCellNumberValue("B", 0);
						let c = commandRow.getCellNumberValue("C", 0);
						let d = commandRow.getCellNumberValue("D", 0);

						if (isInvertCoords) {
							y = shapeHeight - y;
							b = shapeHeight - b;
							c = -c;
						}

						let newX = convertUnits(x, additionalUnitCoefficient);
						let newY = convertUnits(y, additionalUnitCoefficient);
						let newA = convertUnits(a, additionalUnitCoefficient);
						let newB = convertUnits(b, additionalUnitCoefficient);
						let newC = c * radToC;
						let newD = d;

						// same but with a length in EMUs units and an angle in C-units, which will be expected clockwise
						// as in other sdkjs/common/Drawings/Format/Path.js functions.
						path.ellipticalArcTo(newX, newY, newA, newB, newC, newD);
						lastPoint.x = newX;
						lastPoint.y = newY;
						break;
					}
					case "Ellipse":
					{
						/**
						 * computes radii using center and two points on ellipse
						 * @param {number} cx
						 * @param {number} cy
						 * @param {number} x1
						 * @param {number} y1
						 * @param {number} x2
						 * @param {number} y2
						 * @return {{hR: number, wR: number} | false} radii
						 */
						function transformEllipseParams(cx, cy, x1, y1, x2, y2) {
							// subtract center from points
							// let rX1 = x1 - cx;
							// let rY1 = y1 - cy;
							// let rX2 = x2 - cx;
							// let rY2 = y2 - cy;

							// passing two points to (x1/a)^2 + (y1/b)^2 = 1 and (x2/a)^2 + (y2/b)^2 = 1 ellipse equations
							// we can wind radii: a and b values;
							// we can write it like below
							// let b = Math.sqrt((rX2*rX2*rY1*rY1-rY2*rY2*rX1*rX1)/(rX2*rX2-rX1*rX1));
							// let a = Math.sqrt(rX1*rX1*b*b/(b*b-rY1*rY1));
							// it is not useful I guess bcs ellipse is rotated so we need another equation

							// AscCommon.consoleLog('For ellipse with relative center (', cx, ', ', cy, ')');
							// AscCommon.consoleLog('point 1: (', x1, ', ', y1, '), point 2: (', x2, ', ', y2, ')');

							let rx = Math.hypot(x1 - cx, y1 - cy);
							let ry = Math.hypot(x2 - cx, y2 - cy);

							if ((x1 !== cx && y1 !== cy) || (x2 !== cx && y2 !== cy)) {
								// if some of points is not on the same vertical or horizontal as ellipse
								// ellipse command for rotated ellipse is not yet realized
								return false;
							}

							return {wR: rx, hR: ry};
						}

						let centerPointXValue = commandRow.getCellNumberValue("X", 0);
						let centerPointYValue = commandRow.getCellNumberValue("Y", 0);
						let somePointXValue = commandRow.getCellNumberValue("A", 0);
						let somePointYValue = commandRow.getCellNumberValue("B", 0);
						let anotherPointXValue = commandRow.getCellNumberValue("C", 0);
						let anotherPointYValue = commandRow.getCellNumberValue("D", 0);

						if (isInvertCoords) {
							centerPointYValue = shapeHeight - centerPointYValue;
							somePointYValue = shapeHeight - somePointYValue;
							anotherPointYValue = shapeHeight - anotherPointYValue;
						}

						let newX = convertUnits(centerPointXValue, additionalUnitCoefficient);
						let newY = convertUnits(centerPointYValue, additionalUnitCoefficient);
						let newA = convertUnits(somePointXValue, additionalUnitCoefficient);
						let newB = convertUnits(somePointYValue, additionalUnitCoefficient);
						let newC = convertUnits(anotherPointXValue, additionalUnitCoefficient);
						let newD = convertUnits(anotherPointYValue, additionalUnitCoefficient);

						let wRhR = transformEllipseParams(newX, newY, newA, newB, newC, newD);
						if (!wRhR) {
							AscCommon.consoleLog('Ellipse command for rotated ellipse is not yet realized');
						}
						// start to draw from ellipse right point

						// Check [MS-VSDX]-220215 2.2.3.2.2.Geometry Path
						// ellipse command implies moveTo
						path.moveTo(newX + wRhR.wR, newY);
						path.arcTo(wRhR.wR, wRhR.hR, 0, 180 * AscFormat.degToC);
						path.arcTo(wRhR.wR, wRhR.hR, 180 * AscFormat.degToC, 180 * AscFormat.degToC);
						// If the Row_Type element is of type Ellipse or InfiniteLine, it specifies the only segment of the path.
						path.moveTo(lastPoint.x, lastPoint.y);
						// lastPoint.x = newX;
						// lastPoint.y = newY;
						break;
					}
					case "ArcTo":
					{
						// https://learn.microsoft.com/en-us/office/client-developer/visio/arcto-row-geometry-section
						// circular arc

						// middleGap = a. can be negative which leads to opposite arc direction clockwise or anti-clockwise

						let x = commandRow.getCellNumberValue("X", 0);					// xEnd
						let y = commandRow.getCellNumberValue("Y", 0);					// yEnd
						let a = commandRow.getCellNumberValue("A", 0);					// middleGap

						if (isInvertCoords) {
							y = shapeHeight - y;
							a = -a;
						}

						let newX = convertUnits(x, additionalUnitCoefficient);
						let newY = convertUnits(y, additionalUnitCoefficient);
						let newA = convertUnits(a, additionalUnitCoefficient);

						// transform params for ellipticalArcTo
						let chordVector = {x: newX - lastPoint.x, y: newY - lastPoint.y };
						let chordVectorAngle = Math.atan2(chordVector.y, chordVector.x);
						let gapVectorAngle = chordVectorAngle - Math.PI / 2; // perpendicular clock wise
						let gapVector = {x: newA * Math.cos(gapVectorAngle), y: newA * Math.sin(gapVectorAngle)};
						let chordCenter = {x: chordVector.x / 2 + lastPoint.x, y: chordVector.y / 2 + lastPoint.y};
						let controlPoint = {x: chordCenter.x + gapVector.x, y: chordCenter.y + gapVector.y};

						if (a === 0) {
							// in fact it is line is a - arc gap = 0
							// ellipticalArcTo could catch cases line these but
							// there inaccuracy comes
							path.lnTo(newX, newY);
							AscCommon.consoleLog("tranform ellipticalArcTo to line",
								newX, newY, controlPoint.x, controlPoint.y, 0, 1);
							// path.ellipticalArcTo(newX, newY, controlPoint.x, controlPoint.y, 0, 1);
						} else  {
							path.ellipticalArcTo(newX, newY, controlPoint.x, controlPoint.y, 0, 1);
						}

						lastPoint.x = newX;
						lastPoint.y = newY;
						break;
					}
					case "PolylineTo":
					{
						// https://learn.microsoft.com/en-us/office/client-developer/visio/polylineto-row-geometry-section
						let x = commandRow.getCellNumberValue("X", 0);
						let y = commandRow.getCellNumberValue("Y", 0);

						if (isInvertCoords) {
							y = shapeHeight - y;
						}

						// formula: knotLast, degree, xType, yType, x1, y1, x2, y2, ...
						let formula = String(commandRow.getCell("A").v).trim();
						let formulaValues = formula.substring(9, formula.length - 1).split(",");

						let xType = parseInt(formulaValues[0]);
						let yType = parseInt(formulaValues[1]);

						//Convert units to EMUs
						let xEndPointNew = convertUnits(x, additionalUnitCoefficient);
						let yEndPointNew = convertUnits(y, additionalUnitCoefficient);
						for (let k = 2; k < formulaValues.length; k++) {
							// convert x and y

							if (isInvertCoords && (k + 1) % 2 === 0) {
								let maxY = yType === 0 ? 1 : shapeHeight;
								formulaValues[k] = maxY - Number(formulaValues[k]);
							}

							formulaValues[k] = convertUnits(Number(formulaValues[k]), additionalUnitCoefficient);
						}

						let xScale = 1;
						let yScale = 1;

						if (xType === 0)
							xScale = shapeWidth;
						if (yType === 0)
							yScale = shapeHeight;

						// scale x and y and draw line
						let groupsCount = (formulaValues.length - 2) / 2;
						for (let j = 0; j < groupsCount; j++) {
							let pointX = Number(formulaValues[2 + j * 2]);
							let pointY = Number(formulaValues[2 + j * 2 + 1]);

							// scale only in formula
							let scaledX = pointX * xScale;
							let scaledY = pointY * yScale;

							path.lnTo(scaledX, scaledY);
						}

						// then go to x y from command args
						path.lnTo(xEndPointNew, yEndPointNew);

						lastPoint.x = xEndPointNew;
						lastPoint.y = yEndPointNew;
						break;
					}
					case "NURBSTo":
					{
						// https://learn.microsoft.com/en-us/office/client-developer/visio/nurbsto-row-geometry-section
						let xEndPoint = commandRow.getCellNumberValue("X", 0);
						let yEndPoint = commandRow.getCellNumberValue("Y", 0);
						let preLastKnot = commandRow.getCellNumberValue("A", 0);
						let lastWeight = commandRow.getCellNumberValue("B", 0);
						let firstKnot = commandRow.getCellNumberValue("C", 0);
						let firstWeight = commandRow.getCellNumberValue("D", 0);
						// NURBS formula: knotLast, degree, xType, yType, x1, y1, knot1, weight1, x2, y2, knot2, weight2, ...
						let formula = commandRow.getCellStringValue("E");
						if (!formula) {
							AscCommon.consoleLog("!formula for NURBSTo");
							break;
						}
						let formulaValues = formula.trim().substring(6, formula.length - 1).split(",");

						if (isInvertCoords) {
							yEndPoint = shapeHeight - yEndPoint;
						}

						//Parse arguments
						let lastKnot = Number(formulaValues[0]);
						let degree = Number(formulaValues[1]);
						let xType =	parseInt(formulaValues[2]);
						let yType = parseInt(formulaValues[3]);

						//Convert units to EMUs
						let xEndPointNew = convertUnits(xEndPoint, additionalUnitCoefficient);
						let yEndPointNew = convertUnits(yEndPoint, additionalUnitCoefficient);
						for (let k = 4; k < formulaValues.length; k++) {
							if (k % 4 === 0 || k % 4 === 1) {
								// convert x and y

								if (isInvertCoords && (k % 4 === 1)) {
									let maxY = yType === 0 ? 1 : shapeHeight;
									formulaValues[k] = maxY - Number(formulaValues[k]);
								}

								formulaValues[k] = convertUnits(Number(formulaValues[k]), additionalUnitCoefficient);
							}
						}

						let prevLastX = lastPoint.x;
						let prevLastY = lastPoint.y;

						let xScale = 1;
						let yScale = 1;

						if (xType === 0)
							xScale = shapeWidth;
						if (yType === 0)
							yScale = shapeHeight;

						/** @type {{x: Number, y: Number}[]} */
						let controlPoints = [];
						/** @type {Number[]} */
						let weights = [];
						/** @type {Number[]} */
						let knots = [];

						knots.push(firstKnot);
						weights.push(firstWeight);
						controlPoints.push({x: prevLastX, y: prevLastY});

						// point + knot groups
						let groupsCount = (formulaValues.length - 4) / 4;
						for (let j = 0; j < groupsCount; j++) {
							let controlPointX = Number(formulaValues[4 + j * 4]);
							let controlPointY = Number(formulaValues[4 + j * 4 + 1]);
							let knot = Number(formulaValues[4 + j * 4 + 2]);
							let weight = Number(formulaValues[4 + j * 4 + 3]);

							// scale only in formula
							let scaledX = controlPointX * xScale;
							let scaledY = controlPointY * yScale;

							controlPoints.push({x: scaledX, y: scaledY});
							knots.push(knot);
							weights.push(weight);
						}

						knots.push(preLastKnot);
						knots.push(lastKnot);
						// add 3 more knots for 3 degree NURBS to clamp curve at end point
						// a clamped knot vector must have `degree + 1` equal knots
						for (let j = 0; j < degree; j++) {
							knots.push(lastKnot);
						}
						weights.push(lastWeight);
						controlPoints.push({x: xEndPointNew, y: yEndPointNew});

						path.nurbsTo(controlPoints, weights, knots, degree);

						lastPoint.x = xEndPointNew;
						lastPoint.y = yEndPointNew;
						break;
					}
					case "SplineStart":
					{
						// https://learn.microsoft.com/en-us/office/client-developer/visio/splinestart-row-geometry-section

						let secondControlPointY = commandRow.getCellNumberValue("Y", 0);
						let degree = commandRow.getCellNumberValue("D", 0); // not angle

						if (isInvertCoords) {
							secondControlPointY = shapeHeight - secondControlPointY;
						}
						splineStartCommandData = {
							firstControlPointX : lastPoint.x,
							firstControlPointY: lastPoint.y,
							secondControlPointX: convertUnits(commandRow.getCellNumberValue("X", 0),
								additionalUnitCoefficient),
							secondControlPointY: convertUnits(secondControlPointY,additionalUnitCoefficient),
							secondKnot: commandRow.getCellNumberValue("A", 0),
							firstKnot: commandRow.getCellNumberValue("B", 0),
							lastKnot: commandRow.getCellNumberValue("C", 0),
							degree: degree
						};
						break;
					}
					case "SplineKnot":
					{
						// https://learn.microsoft.com/en-us/office/client-developer/visio/splineknot-row-geometry-section
						let controlPointY = commandRow.getCellNumberValue("Y", 0);

						if (isInvertCoords) {
							controlPointY = shapeHeight - controlPointY;
						}
						splineKnotCommandsData.push({
							controlPointX: convertUnits(commandRow.getCellNumberValue("X", 0),
								additionalUnitCoefficient),
							controlPointY: convertUnits(controlPointY,additionalUnitCoefficient),
							knot: commandRow.getCellNumberValue("A", 0)
						});
						break;
					}
					case "InfiniteLine":
					{
						// https://learn.microsoft.com/en-us/office/client-developer/visio/infiniteline-row-geometry-section
						let x = commandRow.getCellNumberValue("X", 0);
						let y = commandRow.getCellNumberValue("Y", 0);
						let a = commandRow.getCellNumberValue("A", 0);
						let b = commandRow.getCellNumberValue("B", 0);

						if (isInvertCoords) {
							y = shapeHeight - y;
							b = shapeHeight - b;
						}

						let xNew = convertUnits(x, additionalUnitCoefficient);
						let yNew = convertUnits(y, additionalUnitCoefficient);
						let aNew = convertUnits(a, additionalUnitCoefficient);
						let bNew = convertUnits(b, additionalUnitCoefficient);

						let maxValue = 10000000000;
						if (xNew === aNew) {
							path.moveTo(xNew, -maxValue);
							path.lnTo(xNew, maxValue);
						} else if (yNew === bNew) {
							path.moveTo(-maxValue, yNew);
							path.lnTo(maxValue, yNew);
						} else {
							// visio doesnt draw diagonal infinite lines
							AscCommon.consoleLog("visio doesnt draw diagonal infinite lines");
						}
						// don't set last point here bcs it is always the only one element in geometry
						break;
					}
					case "RelCubBezTo":
					{
						// https://learn.microsoft.com/en-us/office/client-developer/visio/relcubbezto-row-geometry-section
						let x = commandRow.getCellNumberValue("X", 0);
						let y = commandRow.getCellNumberValue("Y", 0);
						let a = commandRow.getCellNumberValue("A", 0);
						let b = commandRow.getCellNumberValue("B", 0);
						let c = commandRow.getCellNumberValue("C", 0);
						let d = commandRow.getCellNumberValue("D", 0);

						if (isInvertCoords) {
							y = 1 - y;
							b = 1 - b;
							d = 1 - d;
						}

						let xNew = convertUnits(x, additionalUnitCoefficient) * shapeWidth;
						let yNew = convertUnits(y, additionalUnitCoefficient) * shapeHeight;
						let aNew = convertUnits(a, additionalUnitCoefficient) * shapeWidth;
						let bNew = convertUnits(b, additionalUnitCoefficient) * shapeHeight;
						let cNew = convertUnits(c, additionalUnitCoefficient) * shapeWidth;
						let dNew = convertUnits(d, additionalUnitCoefficient) * shapeHeight;

						path.cubicBezTo(aNew, bNew, cNew, dNew, xNew, yNew);

						lastPoint.x = xNew;
						lastPoint.y = yNew;
						break;
					}
					case "RelEllipticalArcTo":
					{
						let x = commandRow.getCellNumberValue("X", 0);
						let y = commandRow.getCellNumberValue("Y", 0);
						let a = commandRow.getCellNumberValue("A", 0);
						let b = commandRow.getCellNumberValue("B", 0);
						let c = commandRow.getCellNumberValue("C", 0);
						let d = commandRow.getCellNumberValue("D", 0);

						if (isInvertCoords) {
							y = 1 - y;
							b = 1 - b;
							c = -c;
						}

						let newX = convertUnits(x, additionalUnitCoefficient) * shapeWidth;
						let newY = convertUnits(y, additionalUnitCoefficient) * shapeHeight;
						let newA = convertUnits(a, additionalUnitCoefficient) * shapeWidth;
						let newB = convertUnits(b, additionalUnitCoefficient) * shapeHeight;
						let newC = c * radToC;
						let newD = d;

						// same but with a length in EMUs units and an angle in C-units, which will be expected clockwise
						// as in other sdkjs/common/Drawings/Format/Path.js functions.
						path.ellipticalArcTo(newX, newY, newA, newB, newC, newD);
						lastPoint.x = newX;
						lastPoint.y = newY;
						break;
					}
					case "RelQuadBezTo":
					{
						// https://learn.microsoft.com/en-us/office/client-developer/visio/relquadbezto-row-geometry-section
						let x = commandRow.getCellNumberValue("X", 0);
						let y = commandRow.getCellNumberValue("Y", 0);
						let a = commandRow.getCellNumberValue("A", 0);
						let b = commandRow.getCellNumberValue("B", 0);

						if (isInvertCoords) {
							y = 1 - y;
							b = 1 - b;
						}

						let xNew = convertUnits(x, additionalUnitCoefficient) * shapeWidth;
						let yNew = convertUnits(y, additionalUnitCoefficient) * shapeHeight;
						let aNew = convertUnits(a, additionalUnitCoefficient) * shapeWidth;
						let bNew = convertUnits(b, additionalUnitCoefficient) * shapeHeight;

						path.quadBezTo(aNew, bNew, xNew, yNew);

						lastPoint.x = xNew;
						lastPoint.y = yNew;
						break;
					}
				}
				if (prevCommandName === "SplineKnot" &&
					(commandName !== "SplineKnot" || j === geometrySection.getElements().length - 1) &&
					splineStartCommandData !== undefined) {
					// draw spline

					/** @type {{x: Number, y: Number}[]} */
					let controlPoints = [];
					/** @type {Number[]} */
					let knots = [];

					let degree = splineStartCommandData.degree;

					controlPoints.push({x: splineStartCommandData.firstControlPointX,
						y: splineStartCommandData.firstControlPointY});
					controlPoints.push({x: splineStartCommandData.secondControlPointX,
						y: splineStartCommandData.secondControlPointY});
					knots.push(splineStartCommandData.firstKnot);
					knots.push(splineStartCommandData.secondKnot);

					splineKnotCommandsData.forEach(function (splineKnotCommandData) {
						controlPoints.push({x: splineKnotCommandData.controlPointX, y: splineKnotCommandData.controlPointY});
						knots.push(splineKnotCommandData.knot);
					});

					knots.push(splineStartCommandData.lastKnot);
					// add 3 more knots for 3 degree NURBS to clamp curve at end point
					// a clamped knot vector must have `degree + 1` equal knots
					for (let j = 0; j < degree; j++) {
						knots.push(splineStartCommandData.lastKnot);
					}

					let weights = new Array(controlPoints.length).fill(1);

					path.nurbsTo(controlPoints, weights, knots, degree);

					lastPoint.x = controlPoints[controlPoints.length - 1].x;
					lastPoint.y = controlPoints[controlPoints.length - 1].y;
				}
				prevCommandName = commandName;
			}

			// path.close();

		}

		if (allFilledPathNoLineValuesEqual) {
			pathWithFill.setStroke(!filledPathNoLine);
		}

		if (allUnfilledPathNoLineValuesEqual) {
			pathWithoutFill.setStroke(!unfilledPathNoLine);
		}

		geometry.setPreset("Any");
		geometry.AddPath(pathWithFill);
		geometry.AddPath(pathWithoutFill);

		// TODO add connections
		// geometry.AddCnx('_3cd4', 'hc', 't');
		// geometry.AddCnx('cd2', 'l', 'vc');
		// geometry.AddCnx('cd4', 'hc', 'b');
		// geometry.AddCnx('0', 'r', 'vc');
		return geometry;
	}

	function getGeometryFromClass(shape) {
		// see ECMA-376-1_5th_edition and Geometry.js

		//TODO maybe get Zip so that we can find parts by relationships
		let master1shape1Geometry = getGeometryFromShape(shape)
		return master1shape1Geometry;
	}

	window['AscVisio'].getGeometryFromShape = getGeometryFromShape;
	window['AscVisio'].getGeometryFromClass = getGeometryFromClass;
})(window, window.document);

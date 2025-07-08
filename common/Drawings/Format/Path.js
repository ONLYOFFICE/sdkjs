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

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {
	const moveTo = 0;
	const lineTo = 1;
	const arcTo = 2;
	const bezier3 = 3;
	const bezier4 = 4;
	const close = 5;
	const ellipticalArcTo = 6;
	const nurbsTo = 7;

// Import
	const cToRad = AscFormat.cToRad;
	const HitToArc = AscFormat.HitToArc;
	const ArcToCurvers = AscFormat.ArcToCurvers;
	const ArcToOnCanvas = AscFormat.ArcToOnCanvas;
	const HitInLine = AscFormat.HitInLine;
	const HitInBezier4 = AscFormat.HitInBezier4;
	const HitInBezier3 = AscFormat.HitInBezier3;
	const MOVE_DELTA = AscFormat.MOVE_DELTA;

	const cToRad2 = (Math.PI / 60000) / 180;

	const degToC = 60000;
	const radToDeg = 180 / Math.PI;
	const cToDeg = cToRad2 * radToDeg;

	function CChangesDrawingsAddPathCommand(Class, oCommand, nIndex, bReverse) {
		this.Type = AscDFH.historyitem_PathAddPathCommand;
		this.Command = oCommand;
		this.Index = nIndex;
		this.bReverse = bReverse;
		AscDFH.CChangesBase.call(this, Class);
	}

	CChangesDrawingsAddPathCommand.prototype = Object.create(AscDFH.CChangesBase.prototype);
	CChangesDrawingsAddPathCommand.prototype.constructor = CChangesDrawingsAddPathCommand;

	CChangesDrawingsAddPathCommand.prototype.Undo = function () {
		if (this.bReverse) {
			this.Class.ArrPathCommandInfo.splice(this.Index, 0, this.Command);
		}
		else {
			this.Class.ArrPathCommandInfo.splice(this.Index, 1);
		}
	};
	CChangesDrawingsAddPathCommand.prototype.Redo = function () {
		if (this.bReverse) {
			this.Class.ArrPathCommandInfo.splice(this.Index, 1);
		}
		else {
			this.Class.ArrPathCommandInfo.splice(this.Index, 0, this.Command);
		}
	};


	CChangesDrawingsAddPathCommand.prototype.WriteToBinary = function (Writer) {
		Writer.WriteLong(this.Index);
		Writer.WriteLong(this.Command.id);
		Writer.WriteBool(!!this.bReverse);
		switch (this.Command.id) {
			case moveTo:
			case lineTo: {
				Writer.WriteString2(this.Command.X);
				Writer.WriteString2(this.Command.Y);
				break;
			}
			case bezier3: {
				Writer.WriteString2(this.Command.X0);
				Writer.WriteString2(this.Command.Y0);
				Writer.WriteString2(this.Command.X1);
				Writer.WriteString2(this.Command.Y1);
				break;
			}
			case bezier4: {
				Writer.WriteString2(this.Command.X0);
				Writer.WriteString2(this.Command.Y0);
				Writer.WriteString2(this.Command.X1);
				Writer.WriteString2(this.Command.Y1);
				Writer.WriteString2(this.Command.X2);
				Writer.WriteString2(this.Command.Y2);
				break;
			}
			case arcTo: {
				Writer.WriteString2(this.Command.hR);
				Writer.WriteString2(this.Command.wR);
				Writer.WriteString2(this.Command.stAng);
				Writer.WriteString2(this.Command.swAng);
				break;
			}
			case close: {
				break;
			}
		}
	};


	CChangesDrawingsAddPathCommand.prototype.ReadFromBinary = function (Reader) {
		this.Index = Reader.GetLong();
		this.Command = {};
		this.Command.id = Reader.GetLong();
		this.bReverse = Reader.GetBool();
		switch (this.Command.id) {
			case moveTo:
			case lineTo: {
				this.Command.X = Reader.GetString2();
				this.Command.Y = Reader.GetString2();
				break;
			}
			case bezier3: {
				this.Command.X0 = Reader.GetString2();
				this.Command.Y0 = Reader.GetString2();
				this.Command.X1 = Reader.GetString2();
				this.Command.Y1 = Reader.GetString2();
				break;
			}
			case bezier4: {
				this.Command.X0 = Reader.GetString2();
				this.Command.Y0 = Reader.GetString2();
				this.Command.X1 = Reader.GetString2();
				this.Command.Y1 = Reader.GetString2();
				this.Command.X2 = Reader.GetString2();
				this.Command.Y2 = Reader.GetString2();
				break;
			}
			case arcTo: {
				this.Command.hR = Reader.GetString2();
				this.Command.wR = Reader.GetString2();
				this.Command.stAng = Reader.GetString2();
				this.Command.swAng = Reader.GetString2();
				break;
			}
			case close: {
				break;
			}
		}
	};


	AscDFH.changesFactory[AscDFH.historyitem_PathSetStroke] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_PathSetExtrusionOk] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_PathSetFill] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_PathSetPathH] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_PathSetPathW] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_PathAddPathCommand] = CChangesDrawingsAddPathCommand;

	AscDFH.drawingsChangesMap[AscDFH.historyitem_PathSetStroke] = function (oClass, value) {
		oClass.stroke = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_PathSetExtrusionOk] = function (oClass, value) {
		oClass.extrusionOk = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_PathSetFill] = function (oClass, value) {
		oClass.fill = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_PathSetPathH] = function (oClass, value) {
		oClass.pathH = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_PathSetPathW] = function (oClass, value) {
		oClass.pathW = value;
	};


	function Path() {
		AscFormat.CBaseFormatObject.call(this);
		this.stroke = null;
		this.extrusionOk = null;
		this.fill = null;
		this.pathH = null;
		this.pathW = null;

		this.ArrPathCommandInfo = [];
		this.ArrPathCommand = [];
	}

	AscFormat.InitClass(Path, AscFormat.CBaseFormatObject, AscDFH.historyitem_type_Path);
	Path.prototypeRefresh_RecalcData = function () {
	};
	Path.prototype.createDuplicate = function () {
		let p = new Path();
		p.setStroke(this.stroke);
		p.setExtrusionOk(this.extrusionOk);
		p.setFill(this.fill);
		p.setPathH(this.pathH);
		p.setPathW(this.pathW);
		for (let i = 0; i < this.ArrPathCommandInfo.length; ++i) {
			let command = this.ArrPathCommandInfo[i];
			switch (command.id) {
				case moveTo:
				case lineTo: {
					let x = command.X;
					let y = command.Y;
					p.addPathCommand({id: command.id, X: x, Y: y});
					break;
				}
				case bezier3: {
					let X0 = command.X0;
					let Y0 = command.Y0;
					let X1 = command.X1;
					let Y1 = command.Y1;
					p.addPathCommand({id: bezier3, X0: X0, Y0: Y0, X1: X1, Y1: Y1});
					break;
				}
				case bezier4: {
					let X0 = command.X0;
					let Y0 = command.Y0;
					let X1 = command.X1;
					let Y1 = command.Y1;
					let X2 = command.X2;
					let Y2 = command.Y2;
					p.addPathCommand({id: bezier4, X0: X0, Y0: Y0, X1: X1, Y1: Y1, X2: X2, Y2: Y2});
					break;
				}
				case arcTo: {
					let hR = command.hR;
					let wR = command.wR;
					let stAng = command.stAng;
					let swAng = command.swAng;
					p.addPathCommand({id: arcTo, hR: hR, wR: wR, stAng: stAng, swAng: swAng});
					break;
				}
				case close: {
					p.addPathCommand({id: close});
					break;
				}
				case ellipticalArcTo: {
					let a = command.a;
					let b = command.b;
					let c = command.c;
					let d = command.d;
					let x = command.x;
					let y = command.y;
					p.addPathCommand({id: ellipticalArcTo, a: a, b: b, c: c, d: d, x: x,
						y: y});
					break;
				}
				case nurbsTo: {
					let nurbsToCopy = cloneSplineObject(command);
					p.addPathCommand(nurbsToCopy);
					break;
				}
			}
		}
		return p;


		/**
		 * Deep clones a NURBS/spline-like object with the structure shown in the screenshot
		 * Object contains:
		 * - controlPoints: Array of {x: Number, y: Number} points
		 * - degree: Number
		 * - id: Number/String
		 * - knots: Array of Numbers
		 * - weights: Array of Numbers
		 *
		 * @param {Object} splineObj - The spline object to clone
		 * @return {Object} A deep copy of the spline object
		 */
		function cloneSplineObject(splineObj) {
			let clone = {};

			// Clone controlPoints if they exist
			if (Array.isArray(splineObj.controlPoints)) {
				clone.controlPoints = [];
				for (let i = 0; i < splineObj.controlPoints.length; i++) {
					let point = splineObj.controlPoints[i];
					let pointClone = {};

					// Copy all properties from the point (typically x and y)
					for (let prop in point) {
						if (point.hasOwnProperty(prop)) {
							pointClone[prop] = point[prop];
						}
					}

					clone.controlPoints.push(pointClone);
				}
			}

			// Clone degree property
			if (splineObj.hasOwnProperty('degree')) {
				clone.degree = splineObj.degree;
			}

			// Clone id property
			if (splineObj.hasOwnProperty('id')) {
				clone.id = splineObj.id;
			}

			// Clone knots array if it exists
			if (Array.isArray(splineObj.knots)) {
				clone.knots = [];
				for (let i = 0; i < splineObj.knots.length; i++) {
					clone.knots.push(splineObj.knots[i]);
				}
			}

			// Clone weights array if it exists
			if (Array.isArray(splineObj.weights)) {
				clone.weights = [];
				for (let i = 0; i < splineObj.weights.length; i++) {
					clone.weights.push(splineObj.weights[i]);
				}
			}

			return clone;
		}
	};
	Path.prototype.setStroke = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_PathSetStroke, this.stroke, pr));
		this.stroke = pr;
	};
	Path.prototype.setExtrusionOk = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_PathSetExtrusionOk, this.extrusionOk, pr));
		this.extrusionOk = pr;
	};
	Path.prototype.setFill = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_PathSetFill, this.fill, pr));
		this.fill = pr;
	};
	Path.prototype.setPathH = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_PathSetPathH, this.pathH, pr));
		this.pathH = pr;
	};
	Path.prototype.setPathW = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_PathSetPathW, this.pathW, pr));
		this.pathW = pr;
	};
	Path.prototype.addPathCommand = function (cmd) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new CChangesDrawingsAddPathCommand(this, cmd, this.ArrPathCommandInfo.length));
		this.ArrPathCommandInfo.push(cmd);
	};
	Path.prototype.moveTo = function (x, y) {
		this.addPathCommand({id: moveTo, X: x, Y: y});
	};
	Path.prototype.lnTo = function (x, y) {
		this.addPathCommand({id: lineTo, X: x, Y: y});
	};
	Path.prototype.arcTo = function (wR, hR, stAng, swAng, ellipseRotation) {
		if (typeof ellipseRotation !== 'undefined') {
			// to be sure for backwards compatibility
			this.addPathCommand({
				id: arcTo,
				wR: wR,
				hR: hR,
				stAng: stAng,
				swAng: swAng,
				ellipseRotation: ellipseRotation
			});
		}
		else {
			this.addPathCommand({id: arcTo, wR: wR, hR: hR, stAng: stAng, swAng: swAng});
		}
	};
	Path.prototype.quadBezTo = function (x0, y0, x1, y1) {
		this.addPathCommand({id: bezier3, X0: x0, Y0: y0, X1: x1, Y1: y1});
	};
	Path.prototype.cubicBezTo = function (x0, y0, x1, y1, x2, y2) {
		this.addPathCommand({id: bezier4, X0: x0, Y0: y0, X1: x1, Y1: y1, X2: x2, Y2: y2});
	};
	Path.prototype.close = function () {
		if (this.ArrPathCommandInfo.length > 0) {
			this.addPathCommand({id: close});
		}
	};
	Path.prototype.ellipticalArcTo = function (x, y, a, b, c, d) {
		this.addPathCommand({id: ellipticalArcTo, x: x, y: y, a: a, b: b, c: c, d: d});
	};
	/**
	 * curveOrder (4 for 3 degree NURBS) + points count = knots count
	 * @param {{x, y}[]} controlPoints
	 * @param {[]} weights
	 * @param {[]} knots
	 * @param {Number} degree
	 */
	Path.prototype.nurbsTo = function (controlPoints, weights, knots, degree) {
		this.addPathCommand({
			id: nurbsTo,
			controlPoints: controlPoints,
			weights: weights,
			knots: knots,
			degree: degree
		});
	};
	Path.prototype.calculateCommandCoord = function (oGdLst, sFormula, dFormulaCoeff, dNumberCoeff) {
		let dVal;
		dVal = oGdLst[sFormula];
		if (dVal !== undefined) {
			return dVal * dFormulaCoeff;
		}
		return parseInt(sFormula, 10) * dNumberCoeff;
	};

	/**
	 * accepts angles in anti-clockwise system
	 * @param {number} startAngle from -180 to 180
	 * @param {number} endAngle from -180 to 180
	 * @param {number} ctrlAngle from -180 to 180
	 * @returns {number} sweep - positive sweep means anti-clockwise
	 */
	function computeSweep(startAngle, endAngle, ctrlAngle) {
		// Normalize angles from −180…180 to [0, 360) interval
		startAngle = (360 + startAngle) % 360;
		endAngle   = (360 + endAngle)   % 360;
		ctrlAngle  = (360 + ctrlAngle)  % 360;

		// different sweeps depending on where the control point is
		let sweep;
		if (startAngle < endAngle) {
			if (startAngle < ctrlAngle && ctrlAngle < endAngle) {
				// positive sweep - anti-clockwise
				sweep = endAngle - startAngle;
			}
			else {
				// negative sweep - clockwise
				sweep = (endAngle - startAngle) - 360;
			}
		}
		else {
			if (endAngle < ctrlAngle && ctrlAngle < startAngle) {
				// negative sweep - clockwise
				sweep = endAngle - startAngle;
			}
			else if (startAngle === endAngle && startAngle === ctrlAngle) {
				sweep = 0;
			}
			else {
				// positive sweep - anti-clockwise
				sweep = 360 - (startAngle - endAngle);
			}
		}

		return sweep;
	}

	/**
	 * afin rotate clockwise
	 * @param {number} x
	 * @param {number} y
	 * @param {number} radiansRotateAngle radians Rotate ClockWise Angle. E.g. 30 degrees rotates does DOWN.
	 * @returns {{x: number, y: number}} point
	 */
	function rotatePointAroundCordsStartClockWise(x, y, radiansRotateAngle) {
		let newX = x * Math.cos(radiansRotateAngle) + y * Math.sin(radiansRotateAngle);
		let newY = x * (-1) * Math.sin(radiansRotateAngle) + y * Math.cos(radiansRotateAngle);
		return {x: newX, y: newY};
	}

	/**
	 *  Accepts visio params
	 *  https://learn.microsoft.com/en-us/office/client-developer/visio/ellipticalarcto-row-geometry-section
	 * @param x0
	 * @param y0
	 * @param x
	 * @param y
	 * @param a
	 * @param b
	 * @param c
	 * @param d
	 * @returns {{stAng: number, ellipseRotation: number, hR: number, wR: number, swAng: number}}
	 * WARNING these are not ECMA params exactly, stAng and swAng angles are anti-clockwise
	 */
	function transformEllipticalArcParams(x0, y0, x, y, a, b, c, d) {
		// https://www.figma.com/file/hs43oiAUyuoqFULVoJ5lyZ/EllipticArcConvert?type=design&node-id=1-2&mode=design&t=QJu8MtR3JV62WiW9-0

		x0 = Number(x0);
		y0 = Number(y0);
		x = Number(x);
		y = Number(y);
		a = Number(a);
		b = Number(b);
		c = Number(c);
		d = Number(d);

		// translate points to ellipse angle
		let startPoint = rotatePointAroundCordsStartClockWise(x0, y0, c);
		let endPoint = rotatePointAroundCordsStartClockWise(x, y, c);
		let controlPoint = rotatePointAroundCordsStartClockWise(a, b, c);
		x0 = startPoint.x;
		y0 = startPoint.y;
		x = endPoint.x;
		y = endPoint.y;
		a = controlPoint.x;
		b = controlPoint.y;

		// http://visguy.com/vgforum/index.php?topic=2464.0
		let d2 = d * d;
		let cx = ((x0 - x) * (x0 + x) * (y - b) - (x - a) * (x + a) * (y0 - y) + d2 * (y0 - y) * (y - b) * (y0 - b)) / (2.0 * ((x0 - x) * (y - b) - (x - a) * (y0 - y)));
		let cy = ((x0 - x) * (x - a) * (x0 - a) / d2 + (x - a) * (y0 - y) * (y0 + y) - (x0 - x) * (y - b) * (y + b)) / (2.0 * ((x - a) * (y0 - y) - (x0 - x) * (y - b)));
		// can also be helpful https://stackoverflow.com/questions/6729056/mapping-svg-arcto-to-html-canvas-arcto

		let rx = Math.sqrt(Math.pow(x0 - cx, 2) + Math.pow(y0 - cy, 2) * d2);
		let ry = rx / d;

		let ctrlAngle = Math.atan2(b - cy, a - cx) * radToDeg;
		let startAngle = Math.atan2(y0 - cy, x0 - cx) * radToDeg;
		let endAngle = Math.atan2(y - cy, x - cx) * radToDeg;

		let sweep = computeSweep(startAngle, endAngle, ctrlAngle);

		let ellipseRotationAngle = c * radToDeg;

		// let mirrorVertically = false;
		// if (mirrorVertically) {
		// 	startAngle = 360 - startAngle;
		// 	sweep = -sweep;
		// 	ellipseRotationAngle = - ellipseRotationAngle;
		// }

		// WARNING these are not ECMA params exactly, stAng and swAng angles are anti-clockwise!
		// about ECMA cord system:
		// c is AntiClockwise so 30 deg go up in Visio
		// but in ECMA it should be another angle
		// because in ECMA angles are clockwise ang 30 deg go down.
		// convert from anticlockwise angle system to clockwise
		// angleEcma = 360 - angleVisio;
		// using visio angles here but still multiply to degToC
		// then cord system trasformed in sdkjs/draw/model/VisioDocument.js CVisioDocument.prototype.draw
		let swAng = sweep * degToC;
		let stAng = startAngle * degToC;
		let ellipseRotationInC = ellipseRotationAngle * degToC;

		let wR = rx;
		let hR = ry;

		return {wR: wR, hR: hR, stAng: stAng, swAng: swAng, ellipseRotation: ellipseRotationInC};
	}

	/**
	 * Determines whether three points are collinear using an adaptive tolerance proportional to the coordinate scale.
	 * @param {number} ax
	 * @param {number} ay
	 * @param {number} bx
	 * @param {number} by
	 * @param {number} cx
	 * @param {number} cy
	 * @returns {boolean}
	 */
	const COL_EPS_K = 3.6e-5;
	function isCollinear(ax, ay, bx, by, cx, cy) {
		// Fast path: any two points coincide – degenerate triangle considered collinear
		if ((ax === bx && ay === by) || (ax === cx && ay === cy) || (bx === cx && by === cy)) {
			return true;
		}

		// Twice the signed area of the triangle (cross product of edge vectors)
		const cross = (bx - ax) * (cy - ay) - (cx - ax) * (by - ay);

		// Adaptive tolerance: scales with the largest coordinate delta
		// Makes the check robust for very small and very large numbers.
		const scale = Math.max(
			Math.abs(bx - ax), Math.abs(by - ay),
			Math.abs(cx - ax), Math.abs(cy - ay)
		);
		const tol = COL_EPS_K * (scale || 1);
		return Math.abs(cross) <= tol;
	}

	Path.prototype.recalculate = function (gdLst, bResetPathsInfo) {
		let ch, cw;
		let dCustomPathCoeffW, dCustomPathCoeffH;
		if (this.pathW != undefined) {
			if (this.pathW > MOVE_DELTA) {
				cw = (gdLst["w"] / this.pathW);
			}
			else {
				cw = 0;
			}
			dCustomPathCoeffW = cw;
		}
		else {
			cw = 1;
			dCustomPathCoeffW = 1 / 36000;
		}
		if (this.pathH != undefined) {
			if (this.pathH > MOVE_DELTA) {
				ch = (gdLst["h"] / this.pathH);
			}
			else {
				ch = 0;
			}
			dCustomPathCoeffH = ch;
		}
		else {
			ch = 1;
			dCustomPathCoeffH = 1 / 36000;
		}
		let APCI = this.ArrPathCommandInfo, n = APCI.length, cmd;
		let x0, y0, x1, y1, x2, y2, wR, hR, stAng, swAng, ellipseRotation, lastX, lastY;
		this.ArrPathCommand.length = 0;
		for (let i = 0; i < n; ++i) {
			cmd = APCI[i];
			switch (cmd.id) {
				case moveTo:
				case lineTo: {
					x0 = this.calculateCommandCoord(gdLst, cmd.X, cw, dCustomPathCoeffW);
					y0 = this.calculateCommandCoord(gdLst, cmd.Y, ch, dCustomPathCoeffH);
					this.ArrPathCommand.push({id: cmd.id, X: x0, Y: y0});
					lastX = x0;
					lastY = y0;
					break;
				}
				case bezier3: {
					x0 = this.calculateCommandCoord(gdLst, cmd.X0, cw, dCustomPathCoeffW);
					y0 = this.calculateCommandCoord(gdLst, cmd.Y0, ch, dCustomPathCoeffH);
					x1 = this.calculateCommandCoord(gdLst, cmd.X1, cw, dCustomPathCoeffW);
					y1 = this.calculateCommandCoord(gdLst, cmd.Y1, ch, dCustomPathCoeffH);
					this.ArrPathCommand.push({id: bezier3, X0: x0, Y0: y0, X1: x1, Y1: y1});
					lastX = x1;
					lastY = y1;
					break;
				}
				case bezier4: {
					x0 = this.calculateCommandCoord(gdLst, cmd.X0, cw, dCustomPathCoeffW);
					y0 = this.calculateCommandCoord(gdLst, cmd.Y0, ch, dCustomPathCoeffH);
					x1 = this.calculateCommandCoord(gdLst, cmd.X1, cw, dCustomPathCoeffW);
					y1 = this.calculateCommandCoord(gdLst, cmd.Y1, ch, dCustomPathCoeffH);
					x2 = this.calculateCommandCoord(gdLst, cmd.X2, cw, dCustomPathCoeffW);
					y2 = this.calculateCommandCoord(gdLst, cmd.Y2, ch, dCustomPathCoeffH);
					this.ArrPathCommand.push({id: bezier4, X0: x0, Y0: y0, X1: x1, Y1: y1, X2: x2, Y2: y2});
					lastX = x2;
					lastY = y2;
					break;
				}
				case arcTo: {
					wR = this.calculateCommandCoord(gdLst, cmd.wR, cw, dCustomPathCoeffW);
					hR = this.calculateCommandCoord(gdLst, cmd.hR, ch, dCustomPathCoeffH);

					stAng = gdLst[cmd.stAng];
					if (stAng === undefined) {
						stAng = parseInt(cmd.stAng, 10);
					}

					swAng = gdLst[cmd.swAng];
					if (swAng === undefined) {
						swAng = parseInt(cmd.swAng, 10);
					}


					let a1 = stAng;
					let a2 = stAng + swAng;
					let a3 = swAng;

					stAng = Math.atan2(ch * Math.sin(a1 * cToRad), cw * Math.cos(a1 * cToRad)) / cToRad;
					swAng = Math.atan2(ch * Math.sin(a2 * cToRad), cw * Math.cos(a2 * cToRad)) / cToRad - stAng;

					if ((swAng > 0) && (a3 < 0)) swAng -= 21600000;
					if ((swAng < 0) && (a3 > 0)) swAng += 21600000;
					if (swAng == 0 && a3 != 0) swAng = 21600000;

					// https://www.figma.com/file/hs43oiAUyuoqFULVoJ5lyZ/EllipticArcConvert?type=design&node-id=291-2&mode=design&t=jLr0jZ6jdV6YhG2S-0
					let a = wR;
					let b = hR;
					let sin2 = Math.sin(stAng * cToRad);
					let cos2 = Math.cos(stAng * cToRad);
					let _xrad = cos2 / a;
					let _yrad = sin2 / b;
					let l = 1 / Math.sqrt(_xrad * _xrad + _yrad * _yrad);
					let xc = lastX - l * cos2;
					let yc = lastY - l * sin2;

					let sin1 = Math.sin((stAng + swAng) * cToRad);
					let cos1 = Math.cos((stAng + swAng) * cToRad);
					let _xrad1 = cos1 / a;
					let _yrad1 = sin1 / b;
					let l1 = 1 / Math.sqrt(_xrad1 * _xrad1 + _yrad1 * _yrad1);

					if (cmd.ellipseRotation === undefined) {
						this.ArrPathCommand.push({
							id: arcTo,
							stX: lastX,
							stY: lastY,
							wR: wR,
							hR: hR,
							stAng: stAng * cToRad,
							swAng: swAng * cToRad
						});

						lastX = xc + l1 * cos1;
						lastY = yc + l1 * sin1;
					}
					else {
						// do transformations with ellipseRotation by analogy. ellipseRotation is added later
						// then calculate new end point
						ellipseRotation = gdLst[cmd.ellipseRotation];
						if (ellipseRotation === undefined) {
							ellipseRotation = parseInt(cmd.ellipseRotation, 10);
						}

						let a4 = ellipseRotation;

						ellipseRotation = Math.atan2(ch * Math.sin(a4 * cToRad), cw * Math.cos(a4 * cToRad)) / cToRad;

						if ((ellipseRotation > 0) && (a4 < 0)) ellipseRotation -= 21600000;
						if ((ellipseRotation < 0) && (a4 > 0)) ellipseRotation += 21600000;
						if (ellipseRotation == 0 && a4 != 0) ellipseRotation = 21600000;

						this.ArrPathCommand.push({
							id: arcTo,
							stX: lastX,
							stY: lastY,
							wR: wR,
							hR: hR,
							stAng: stAng * cToRad,
							swAng: swAng * cToRad,
							ellipseRotation: ellipseRotation * cToRad
						});

						// https://www.figma.com/file/hs43oiAUyuoqFULVoJ5lyZ/EllipticArcConvert?type=design&node-id=291-34&mode=design&t=LKiEAjzKEzKacCBc-0

						// lets convert ECMA clockwise angle to trigonometrical
						// (anti clockwise) angle to correctly calculate sin and cos
						// of l1 angle to calculate l1 end point cords
						let l1AntiClockWiseAngle = (360 - (stAng + swAng)) * cToRad;

						//  new cord system center is ellipse center and its y does up
						let l1xNewCordSystem = l1 * Math.cos(l1AntiClockWiseAngle);
						let l1yNewCordSystem = l1 * Math.sin(l1AntiClockWiseAngle);
						// if no rotate it is:
						// lastX = xc + l1xNewCordSystem;
						// lastY = yc - l1yNewCordSystem;
						// we invert y because start and calculate point coordinate systems are different. see figma

						let l1xyRotatedEllipseNewCordSystem = rotatePointAroundCordsStartClockWise(l1xNewCordSystem, l1yNewCordSystem, ellipseRotation * cToRad);
						let l1xRotatedEllipseOldCordSystem = l1xyRotatedEllipseNewCordSystem.x + xc;
						let l1yRotatedEllipseOldCordSystem = yc - l1xyRotatedEllipseNewCordSystem.y;

						// calculate last point offset after ellipse rotate
						let lastXnewCordSystem = lastX - xc;
						let lastYnewCordSystem = -lastY + yc;
						// center of cord system now is the center of ellipse
						// blue point
						let rotatedLastXYnewCordSystem = rotatePointAroundCordsStartClockWise(lastXnewCordSystem, lastYnewCordSystem, ellipseRotation * cToRad);
						let rotatedLastXoldCordSystem = rotatedLastXYnewCordSystem.x + xc;
						let rotatedLastYoldCordSystem = yc - rotatedLastXYnewCordSystem.y;
						// calculate vector
						let lastPointOffsetVectorX = lastX - rotatedLastXoldCordSystem;
						let lastPointOffsetVectorY = lastY - rotatedLastYoldCordSystem;

						lastX = l1xRotatedEllipseOldCordSystem + lastPointOffsetVectorX;
						lastY = l1yRotatedEllipseOldCordSystem + lastPointOffsetVectorY;
					}

					break;
				}
				case close: {
					this.ArrPathCommand.push({id: close});
					break;
				}
				case ellipticalArcTo: {
					// https://learn.microsoft.com/en-us/office/client-developer/visio/ellipticalarcto-row-geometry-section
					// but with a length in EMUs units and an angle in C-units, which will be expected clockwise as in other functions.
					let x, y, a, b, c, d;
					x = this.calculateCommandCoord(gdLst, cmd.x, cw, dCustomPathCoeffW);
					y = this.calculateCommandCoord(gdLst, cmd.y, ch, dCustomPathCoeffH);
					a = this.calculateCommandCoord(gdLst, cmd.a, cw, dCustomPathCoeffW);
					b = this.calculateCommandCoord(gdLst, cmd.b, ch, dCustomPathCoeffH);

					c = gdLst[cmd.c];
					if (c === undefined) {
						c = parseInt(cmd.c, 10);
					}

					// d is fraction
					d = Number(cmd.d);

					// normalize angle (The values are given in degrees)
					// 360 * n —-> ~0
					// 270 --> 	-90
					// -290 -> 	70
					// -570 ->	150
					// -286.47888333333333 —-> 73.52111666666667
					// -572.9577833333333 —-> 147.04221666666672
					// so c in degrees is from -180 to 180
					c = Math.atan2(ch * Math.sin(c * cToRad), cw * Math.cos(c * cToRad)) / cToRad;

					let cRadians = c * cToRad2;

					if (isNaN(lastY)) {
						lastY = 0;
					}
					if (isNaN(lastX)) {
						lastX = 0;
					}

					// change ellipticalArcTo params to draw arc using old params
					let newParams = transformEllipticalArcParams(lastX, lastY, x, y, a, b, cRadians, d);

					// check if command arguments are wrong. Wrong arguments may refer to huge ellipse. It editor it can
					// break scroll bars. So as visio does let's transform elliptical arc to line.
					// see bug https://bugzilla.onlyoffice.com/show_bug.cgi?id=75317
					// files from 4 to 6 should not be caught

					// ~0 swing angle refers to bad arguments. check it
					// see files:
					// 1 simple lines and ellipses.vsdx
					// 2 swAng === 0 check  testFlatCurve Huge D.vsdx
					// 3 swAng ~=0 testFlatCurve.vsdx
					let swAngCheck = AscFormat.fApproxEqual(newParams.swAng, 0, 1e-7);

					// NaN in newParams refers to bad arguments.
					// see files:
					// 1 simple lines and ellipses.vsdx
					// 7 triangleSquare === 0 cehck test Diagonal.vsdx
					// 8 small square + NaN params in result, isCollinear cathces.vsdx
					let isNaNInParams = isNaN(newParams.swAng) || isNaN(newParams.stAng) ||
							isNaN(newParams.wR) || isNaN(newParams.hR) || isNaN(newParams.ellipseRotation);

					// three points on one line refers to bad arguments
					// see files:
					// 1 simple lines and ellipses.vsdx
					// 3 swAng ~=0 testFlatCurve.vsdx
					// 7 triangleSquare === 0 cehck test Diagonal.vsdx
					// 8 small square + NaN params in result, isCollinear cathces.vsdx
					let isCollinearCheck = isCollinear(lastX, lastY, a, b, x, y);

					if (swAngCheck || isNaNInParams || isCollinearCheck) {
						AscCommon.consoleLog("tranform ellipticalArcTo to line. 2 catch.");
						this.ArrPathCommand.push({id: lineTo, X: a, Y: b}); // go to control point first
						this.ArrPathCommand.push({id: lineTo, X: x, Y: y});
					}
					else {
						this.ArrPathCommand.push({
							id: ellipticalArcTo,
							stX: lastX,
							stY: lastY,
							wR: newParams.wR,
							hR: newParams.hR,
							stAng: newParams.stAng * cToRad,
							swAng: newParams.swAng * cToRad,
							ellipseRotation: newParams.ellipseRotation * cToRad
						});
					}

					lastX = x;
					lastY = y;
					break;
				}
				case nurbsTo: {
					//TODO
					// consider homogenous or euclidean weighted points weights

					/**
					 * for NURBS to Bezier convert https://math.stackexchange.com/questions/417859/convert-a-b-spline-into-bezier-curves
					 * @param {{x: Number, y: Number, z? :Number}[]} controlPoints
					 * @param {Number[]} weights
					 * @param {Number[]} knots
					 * @param {Number} multiplicity
					 * @returns {{controlPoints: {x: Number, y: Number, z? :Number}[], weights: Number[], knots: Number[]} || null} new bezier data
					 */
					function duplicateKnots(controlPoints, weights, knots, multiplicity) {
						/**
						 * http://preserve.mactech.com/articles/develop/issue_25/schneider.html
						 * can be found with pictures
						 * @param {{x: Number, y: Number, z? :Number}[]} controlPoints
						 * @param {Number[]} weights
						 * @param {Number[]} knots
						 * @param {Number} tNew
						 * @return {{controlPoints: {x: Number, y: Number, z? :Number}[], weights: Number[], knots: Number[]} || null} new bezier data
						 */
						function insertKnot(controlPoints, weights, knots, tNew) {
							let n = controlPoints.length;
							let order = knots.length - controlPoints.length;
							let k = order;

							let calculateZ = controlPoints[0].z !== undefined;

							let newKnots = [];
							let newWeights = [];
							let newControlPoints = [];

							// find index to insert tNew after
							let i = -1;
							for (let j = 0; j < n + k; j++) {
								if (tNew > knots[j] && tNew <= knots[j + 1]) {
									i = j;
									break;
								}
							}

							// insert tNew
							if (i === -1) {
								AscCommon.consoleLog("Not found position to insert new knot");
								return null;
							}
							else {
								// Copy knots to new array.
								for (let j = 0; j < n + k + 1; j++) {
									if (j <= i) {
										newKnots[j] = knots[j];
									}
									else if (j === i + 1) {
										newKnots[j] = tNew;
									}
									else {
										newKnots[j] = knots[j - 1];
									}
								}
							}

							// Compute position of new control point and new positions of
							// existing ones.
							let alpha;
							for (let j = 0; j < n + 1; j++) {
								if (j <= i - k + 1) {
									alpha = 1;
								}
								else if (i - k + 2 <= j && j <= i) {
									if (knots[j + k - 1] - knots[j] == 0) {
										alpha = 0;
									}
									else {
										alpha = (tNew - knots[j]) / (knots[j + k - 1] - knots[j]);
									}
								}
								else {
									alpha = 0;
								}
								if (alpha == 0) {
									newControlPoints[j] = controlPoints[j - 1];
									newWeights[j] = weights[j - 1];
								}
								else if (alpha == 1) {
									newControlPoints[j] = controlPoints[j];
									newWeights[j] = weights[j];
								}
								else {
									newControlPoints[j] = {};
									newControlPoints[j].x = (1 - alpha) * controlPoints[j - 1].x + alpha * controlPoints[j].x;
									newControlPoints[j].y = (1 - alpha) * controlPoints[j - 1].y + alpha * controlPoints[j].y;
									if (calculateZ) {
										newControlPoints[j].z = (1 - alpha) * controlPoints[j - 1].z + alpha * controlPoints[j].z;
									}
									newWeights[j] = (1 - alpha) * weights[j - 1] + alpha * weights[j];
								}
							}

							return {
								controlPoints: newControlPoints, weights: newWeights, knots: newKnots,
							};
						}

						/**
						 * Checks if all numbers in an array are in increasing order (allows duplicates)
						 * @param {number[]} arr - The array to check
						 * @returns {boolean} - True if array is in increasing order, false otherwise
						 */
						function isIncreasingOrder(arr) {
							if (arr.length < 2) {
								return true; // An empty array or array with one element is considered in order
							}

							for (let i = 0; i < arr.length - 1; i++) {
								if (arr[i] > arr[i + 1]) {
									return false; // Found an element that's decreasing
								}
							}

							return true;
						}

						if (multiplicity === undefined) {
							AscCommon.consoleLog("Error: multiplicity is undefined");
							return null;
						}

						if (!isIncreasingOrder(knots)) {
							AscCommon.consoleLog("Error: Knots with decreasing elements is not supported. Knots: " + knots);
							return null;
						}

						let knotValue = knots[0];
						let knotIndex;
						while (true) {
							let knotsCount = 0;
							for (let i = 0; i < knots.length; i++) {
								knotsCount += knots[i] === knotValue ? 1 : 0;
							}

							knotIndex = knots.indexOf(knotValue);
							let insertCount = multiplicity - knotsCount;
							insertCount = insertCount < 0 ? 0 : insertCount;
							for (let i = 0; i < insertCount; i++) {
								let newNURBSdata = insertKnot(controlPoints, weights, knots, knotValue);
								if (newNURBSdata == null) {
									AscCommon.consoleLog('Unknown error. unexpected t');
									return null;
								}

								controlPoints = newNURBSdata.controlPoints;
								weights = newNURBSdata.weights;
								knots = newNURBSdata.knots;
							}

							knotIndex = knotIndex + knotsCount + insertCount;
							if (knotIndex === knots.length) {
								// out of bounds
								break;
							}
							knotValue = knots[knotIndex];
						}
						return {controlPoints: controlPoints, weights: weights, knots: knots};
					}

					/**
					 *
					 * @param {{x: Number, y: Number, z? :Number}[]} controlPoints
					 * @param {Number} degree
					 * @returns {{
					 *              startPoint:     {x: Number, y: Number, z? :Number},
					 *              controlPoints:  {x: Number, y: Number, z? :Number}[],
					 *              endPoint:       {x: Number, y: Number, z? :Number}
					 *            }[]}
					 */
					function NURBSnormalizedToBezier(controlPoints, degree) {
						let bezierArray = [];
						// first Bezier
						let nthBezier = {
							startPoint: controlPoints[0], controlPoints: []
						};
						for (let i = 1; i < controlPoints.length; i++) {
							const point = controlPoints[i];
							if (i % degree === 0) {
								nthBezier.endPoint = point;
								let nthBezierCopy = JSON.parse(JSON.stringify(nthBezier));
								bezierArray.push(nthBezierCopy);
								nthBezier = {
									startPoint: point, controlPoints: []
								};
							}
							else {
								nthBezier.controlPoints.push(point);
							}
						}
						return bezierArray;
					}


					// Init arguments
					let controlPoints = cmd.controlPoints;
					let weights = cmd.weights;
					let knots = cmd.knots;
					let degree = cmd.degree;

					for (let j = 0; j < controlPoints.length; j++) {
						controlPoints[j].x = this.calculateCommandCoord(gdLst, controlPoints[j].x, cw, dCustomPathCoeffW);
						controlPoints[j].y = this.calculateCommandCoord(gdLst, controlPoints[j].y, ch, dCustomPathCoeffH);
					}

					// if degree === 0 just draw line to SplineStart like visio does
					if (degree === 0) {
						AscCommon.consoleLog("transform nurbsTo to line because degree is 0");
						this.ArrPathCommand.push({id: lineTo, X: controlPoints[1].x, Y: controlPoints[1].y});
						lastX = controlPoints[1].x;
						lastY = controlPoints[1].y;
						break;
					}

					if (degree + 1 + controlPoints.length !== knots.length) {
						AscCommon.consoleLog("Wrong arguments format.", "Degree + 1 + controlPoints.length !== knots.length", degree + 1 + controlPoints.length, "!==", knots.length);
						break;
					}

					// round knots maybe start is clamped
					// let precision = 10e13;
					// for (let j = 0; j < knots.length; j++) {
					//     knots[j] = Math.round(knots[j] * precision) / precision;
					// }

					let clampedStart = true;
					for (let j = 0; j < degree; j++) {
						// compare first degree + 1 knots
						clampedStart = clampedStart === false ? clampedStart : knots[j] === knots[j + 1];
					}
					if (!clampedStart) {
						AscCommon.consoleLog("first degree + 1 knots are not equal. Non clamped start is not yet supported", "Degree is", degree, "knots:", knots);
						break;
					}

					// Convert to Bezier
					let newNURBSform = duplicateKnots(controlPoints, weights, knots, degree);
					if (newNURBSform === null) {
						AscCommon.consoleLog("duplicateKnots() error");
						break;
					}
					let bezierArray = NURBSnormalizedToBezier(newNURBSform.controlPoints, degree);

					// change nurbsTo params to draw using bezier
					// nurbs degree is equal to each bezier degree
					this.ArrPathCommand.push({id: nurbsTo, degree: degree, bezierArray: bezierArray});

					lastX = bezierArray[bezierArray.length - 1].endPoint.x;
					lastY = bezierArray[bezierArray.length - 1].endPoint.y;
					break;
				}
				default: {
					break;
				}
			}
		}
		if (bResetPathsInfo) {
			delete this.ArrPathCommandInfo;
		}
	};
	Path.prototype.recalculate2 = function (gdLst, bResetPathsInfo) {
		let k = 10e-10;
		let APCI = this.ArrPathCommandInfo, n = APCI.length, cmd;
		let stAng, swAng, lastX, lastY;
		for (let i = 0; i < n; ++i) {
			cmd = APCI[i];
			switch (cmd.id) {
				case moveTo:
				case lineTo: {

					lastX = cmd.X * k;
					lastY = cmd.Y * k;
					this.ArrPathCommand[i] = {id: cmd.id, X: lastX, Y: lastY};
					break;
				}
				case bezier3: {

					lastX = cmd.X1;
					lastY = cmd.Y1;

					this.ArrPathCommand[i] = {id: bezier3, X0: cmd.X0 * k, Y0: cmd.Y0 * k, X1: lastX, Y1: lastY};

					break;
				}
				case bezier4: {
					lastX = cmd.X2;
					lastY = cmd.Y2;

					this.ArrPathCommand[i] = {
						id: bezier4,
						X0: cmd.X0 * k,
						Y0: cmd.Y0 * k,
						X1: cmd.X1 * k,
						Y1: cmd.Y1 * k,
						X2: lastX,
						Y2: lastY
					};
					break;
				}
				case arcTo: {


					let a1 = cmd.stAng;
					let a2 = cmd.stAng + cmd.swAng;
					let a3 = cmd.swAng;

					stAng = Math.atan2(k * Math.sin(a1 * cToRad), k * Math.cos(a1 * cToRad)) / cToRad;
					swAng = Math.atan2(k * Math.sin(a2 * cToRad), k * Math.cos(a2 * cToRad)) / cToRad - cmd.stAng;

					if ((swAng > 0) && (a3 < 0)) swAng -= 21600000;
					if ((swAng < 0) && (a3 > 0)) swAng += 21600000;
					if (swAng == 0 && a3 != 0) swAng = 21600000;

					let a = cmd.wR * k;
					let b = cmd.hR * k;
					let sin2 = Math.sin(stAng * cToRad);
					let cos2 = Math.cos(stAng * cToRad);
					let _xrad = cos2 / a;
					let _yrad = sin2 / b;
					let l = 1 / Math.sqrt(_xrad * _xrad + _yrad * _yrad);
					let xc = lastX - l * cos2;
					let yc = lastY - l * sin2;

					let sin1 = Math.sin((stAng + swAng) * cToRad);
					let cos1 = Math.cos((stAng + swAng) * cToRad);
					let _xrad1 = cos1 / a;
					let _yrad1 = sin1 / b;
					let l1 = 1 / Math.sqrt(_xrad1 * _xrad1 + _yrad1 * _yrad1);

					this.ArrPathCommand[i] = {
						id: arcTo,
						stX: lastX,
						stY: lastY,
						wR: cmd.wR * k,
						hR: cmd.hR * k,
						stAng: stAng * cToRad,
						swAng: swAng * cToRad
					};

					lastX = xc + l1 * cos1;
					lastY = yc + l1 * sin1;


					break;
				}
				case close: {
					this.ArrPathCommand[i] = {id: close};
					break;
				}
				default: {
					break;
				}
			}
		}
		// if(bResetPathsInfo)
		{
			delete this.ArrPathCommandInfo;
		}
	};
	Path.prototype.draw = function (shape_drawer) {
		if (shape_drawer.bIsCheckBounds === true && this.fill == "none") {
			// это для текстур
			return;
		}
		let bIsDrawLast = false;
		let path = this.ArrPathCommand;
		shape_drawer._s();
		for (let j = 0, l = path.length; j < l; ++j) {
			let cmd = path[j];
			switch (cmd.id) {
				case moveTo: {
					bIsDrawLast = true;
					shape_drawer._m(cmd.X, cmd.Y);
					break;
				}
				case lineTo: {
					bIsDrawLast = true;
					shape_drawer._l(cmd.X, cmd.Y);
					break;
				}
				case bezier3: {
					bIsDrawLast = true;
					shape_drawer._c2(cmd.X0, cmd.Y0, cmd.X1, cmd.Y1);
					break;
				}
				case bezier4: {
					bIsDrawLast = true;
					shape_drawer._c(cmd.X0, cmd.Y0, cmd.X1, cmd.Y1, cmd.X2, cmd.Y2);
					break;
				}
				case arcTo: {
					bIsDrawLast = true;
					ArcToCurvers(shape_drawer, cmd.stX, cmd.stY, cmd.wR, cmd.hR, cmd.stAng, cmd.swAng, cmd.ellipseRotation /*ellipseRotation added later*/);
					break;
				}
				case close: {
					shape_drawer._z();
					break;
				}
				case ellipticalArcTo: {
					bIsDrawLast = true;
					ArcToCurvers(shape_drawer, cmd.stX, cmd.stY, cmd.wR, cmd.hR, cmd.stAng, cmd.swAng, cmd.ellipseRotation);
					break;
				}
				case nurbsTo: {
					bIsDrawLast = true;
					cmd.bezierArray.forEach(function (bezier) {
						if (cmd.degree === 2) {
							let cp1x = bezier.controlPoints[0].x;
							let cp1y = bezier.controlPoints[0].y;
							let endx = bezier.endPoint.x;
							let endy = bezier.endPoint.y;
							shape_drawer._c2(cp1x, cp1y, endx, endy);
						}
						else if (cmd.degree === 3) {
							let cp1x = bezier.controlPoints[0].x;
							let cp1y = bezier.controlPoints[0].y;
							let cp2x = bezier.controlPoints[1].x;
							let cp2y = bezier.controlPoints[1].y;
							let endx = bezier.endPoint.x;
							let endy = bezier.endPoint.y;
							shape_drawer._c(cp1x, cp1y, cp2x, cp2y, endx, endy);
						}
						else {
							let startPoint = bezier.startPoint;
							let controlPoints = bezier.controlPoints;
							let endPoint = bezier.endPoint;
							// unlike in other commands pass start point bcs other commands use canvas system
							// end point for next command start point
							// (which we cant get explicitly for _cN calculation)
							shape_drawer._cN(startPoint, controlPoints, endPoint);
						}
					});
					break;
				}
			}
		}

		if (bIsDrawLast) {
			shape_drawer.drawFillStroke(true, this.fill, this.stroke && !shape_drawer.bIsNoStrokeAttack);
		}

		shape_drawer._e();
	};
	Path.prototype.check_bounds = function (checker, geom) {
		let path = this.ArrPathCommand;
		for (let j = 0, l = path.length; j < l; ++j) {
			let cmd = path[j];
			switch (cmd.id) {
				case moveTo: {
					checker._m(cmd.X, cmd.Y);
					break;
				}
				case lineTo: {
					checker._l(cmd.X, cmd.Y);
					break;
				}
				case bezier3: {
					checker._c2(cmd.X0, cmd.Y0, cmd.X1, cmd.Y1);
					break;
				}
				case bezier4: {
					checker._c(cmd.X0, cmd.Y0, cmd.X1, cmd.Y1, cmd.X2, cmd.Y2);
					break;
				}
				case arcTo: {
					ArcToCurvers(checker, cmd.stX, cmd.stY, cmd.wR, cmd.hR, cmd.stAng, cmd.swAng, cmd.ellipseRotation /*ellipseRotation added later*/);
					break;
				}
				case close: {
					checker._z();
					break;
				}
				case ellipticalArcTo: {
					ArcToCurvers(checker, cmd.stX, cmd.stY, cmd.wR, cmd.hR, cmd.stAng, cmd.swAng, cmd.ellipseRotation);
					break;
				}
				case nurbsTo: {
					cmd.bezierArray.forEach(function (bezier) {
						if (cmd.degree === 2) {
							let cp1x = bezier.controlPoints[0].x;
							let cp1y = bezier.controlPoints[0].y;
							let endx = bezier.endPoint.x;
							let endy = bezier.endPoint.y;
							checker._c2(cp1x, cp1y, endx, endy);
						}
						else if (cmd.degree === 3) {
							let cp1x = bezier.controlPoints[0].x;
							let cp1y = bezier.controlPoints[0].y;
							let cp2x = bezier.controlPoints[1].x;
							let cp2y = bezier.controlPoints[1].y;
							let endx = bezier.endPoint.x;
							let endy = bezier.endPoint.y;
							checker._c(cp1x, cp1y, cp2x, cp2y, endx, endy);
						}
						else {
							let startPoint = bezier.startPoint;
							let controlPoints = bezier.controlPoints;
							let endPoint = bezier.endPoint;
							let pointsToCheck = controlPoints.concat(endPoint);
							pointsToCheck = pointsToCheck.concat(startPoint);
							checker.checkPoints(pointsToCheck);
						}
					});
					break;
				}
			}
		}
	};
	Path.prototype.hitInInnerArea = function (canvasContext, x, y) {
		if (this.fill === "none") return false;

		let _arr_commands = this.ArrPathCommand;
		let _commands_count = _arr_commands.length;
		let _command_index;
		let _command;
		canvasContext.beginPath();
		for (_command_index = 0; _command_index < _commands_count; ++_command_index) {
			_command = _arr_commands[_command_index];
			switch (_command.id) {
				case moveTo: {
					canvasContext.moveTo(_command.X, _command.Y);
					break;
				}
				case lineTo: {
					canvasContext.lineTo(_command.X, _command.Y);
					break;
				}
				case arcTo: {
					ArcToOnCanvas(canvasContext, _command.stX, _command.stY, _command.wR, _command.hR, _command.stAng, _command.swAng);
					break;
				}
				case bezier3: {
					canvasContext.quadraticCurveTo(_command.X0, _command.Y0, _command.X1, _command.Y1);
					break;
				}
				case bezier4: {
					canvasContext.bezierCurveTo(_command.X0, _command.Y0, _command.X1, _command.Y1, _command.X2, _command.Y2);
					break;
				}
				case close: {
					canvasContext.closePath();
				}
			}
		}
		return !!canvasContext.isPointInPath(x, y);
	};
	Path.prototype.hitInPath = function (canvasContext, x, y, oAddingPoint, _path_index) {
		let _arr_commands = this.ArrPathCommand;
		let _commands_count = _arr_commands.length;
		let _command_index;
		let _command;
		let _last_x, _last_y;
		let _begin_x, _begin_y;
		for (_command_index = 0; _command_index < _commands_count; ++_command_index) {
			_command = _arr_commands[_command_index];
			switch (_command.id) {
				case moveTo: {
					_last_x = _command.X;
					_last_y = _command.Y;
					_begin_x = _command.X;
					_begin_y = _command.Y;
					break;
				}
				case lineTo: {
					if (HitInLine(canvasContext, x, y, _last_x, _last_y, _command.X, _command.Y)) return true;
					_last_x = _command.X;
					_last_y = _command.Y;
					break;
				}
				case arcTo: {
					if (HitToArc(canvasContext, x, y, _command.stX, _command.stY, _command.wR, _command.hR, _command.stAng, _command.swAng)) return true;
					_last_x = (_command.stX - _command.wR * Math.cos(_command.stAng) + _command.wR * Math.cos(_command.swAng));
					_last_y = (_command.stY - _command.hR * Math.sin(_command.stAng) + _command.hR * Math.sin(_command.swAng));
					break;
				}
				case bezier3: {
					if (HitInBezier3(canvasContext, x, y, _last_x, _last_y, _command.X0, _command.Y0, _command.X1, _command.Y1)) return true;
					_last_x = _command.X1;
					_last_y = _command.Y1;
					break;
				}
				case bezier4: {
					if (HitInBezier4(canvasContext, x, y, _last_x, _last_y, _command.X0, _command.Y0, _command.X1, _command.Y1, _command.X2, _command.Y2)) {
						if (oAddingPoint) {
							oAddingPoint.pathIndex = _path_index;
							oAddingPoint.commandIndex = _command_index;
						}
						return true;
					}
					_last_x = _command.X2;
					_last_y = _command.Y2;
					break;
				}
				case close: {
					if (HitInLine(canvasContext, x, y, _last_x, _last_y, _begin_x, _begin_y)) return true;
				}
			}
		}
		return false;
	};
	Path.prototype.isSmartLine = function () {
		if (this.ArrPathCommand.length != 2) return false;

		if (this.ArrPathCommand[0].id == moveTo && this.ArrPathCommand[1].id == lineTo) {
			if (Math.abs(this.ArrPathCommand[0].X - this.ArrPathCommand[1].X) < 0.0001) return true;

			if (Math.abs(this.ArrPathCommand[0].Y - this.ArrPathCommand[1].Y) < 0.0001) return true;
		}

		return false;
	};
	Path.prototype.isSmartRect = function () {
		if (this.ArrPathCommand.length != 5) return false;

		if (this.ArrPathCommand[0].id != moveTo || this.ArrPathCommand[1].id != lineTo || this.ArrPathCommand[2].id != lineTo || this.ArrPathCommand[3].id != lineTo || (this.ArrPathCommand[4].id != lineTo && this.ArrPathCommand[4].id != close)) return false;

		let _float_eps = 0.0001;
		if (Math.abs(this.ArrPathCommand[0].X - this.ArrPathCommand[1].X) < _float_eps) {
			if (Math.abs(this.ArrPathCommand[1].Y - this.ArrPathCommand[2].Y) < _float_eps) {
				if (Math.abs(this.ArrPathCommand[2].X - this.ArrPathCommand[3].X) < _float_eps && Math.abs(this.ArrPathCommand[3].Y - this.ArrPathCommand[0].Y) < _float_eps) {
					if (this.ArrPathCommand[4].id == close) return true;

					if (Math.abs(this.ArrPathCommand[0].X - this.ArrPathCommand[4].X) < _float_eps && Math.abs(this.ArrPathCommand[0].Y - this.ArrPathCommand[4].Y) < _float_eps) {
						return true;
					}
				}
			}
		}
		else if (Math.abs(this.ArrPathCommand[0].Y - this.ArrPathCommand[1].Y) < _float_eps) {
			if (Math.abs(this.ArrPathCommand[1].X - this.ArrPathCommand[2].X) < _float_eps) {
				if (Math.abs(this.ArrPathCommand[2].Y - this.ArrPathCommand[3].Y) < _float_eps && Math.abs(this.ArrPathCommand[3].X - this.ArrPathCommand[0].X) < _float_eps) {
					if (this.ArrPathCommand[4].id == close) return true;

					if (Math.abs(this.ArrPathCommand[0].X - this.ArrPathCommand[4].X) < _float_eps && Math.abs(this.ArrPathCommand[0].Y - this.ArrPathCommand[4].Y) < _float_eps) {
						return true;
					}
				}
			}
		}

		return false;
	};
	Path.prototype.drawSmart = function (shape_drawer) {
		let _graphics = shape_drawer.Graphics;
		let _full_trans = _graphics.m_oFullTransform;

		if (!_graphics || !_full_trans || undefined == _graphics.m_bIntegerGrid || true === shape_drawer.bIsNoSmartAttack) return this.draw(shape_drawer);

		let bIsTransformed = (_full_trans.shx == 0 && _full_trans.shy == 0) ? false : true;

		if (bIsTransformed) return this.draw(shape_drawer);

		let isLine = this.isSmartLine();
		let isRect = false;
		if (!isLine) isRect = this.isSmartRect();

		if (window["NATIVE_EDITOR_ENJINE"] || (!isLine && !isRect)) return this.draw(shape_drawer);

		let _old_int = _graphics.m_bIntegerGrid;

		if (false == _old_int) _graphics.SetIntegerGrid(true);

		let dKoefMMToPx = Math.max(_graphics.m_oCoordTransform.sx, 0.001);

		let _ctx = _graphics.m_oContext;
		let bIsStroke = (shape_drawer.bIsNoStrokeAttack || (this.stroke !== true)) ? false : true;
		let bIsEven = false;
		if (bIsStroke) {
			let _lineWidth = Math.max((shape_drawer.StrokeWidth * dKoefMMToPx + 0.5) >> 0, 1);
			_ctx.lineWidth = _lineWidth;

			if ((_lineWidth & 0x01) == 0x01) bIsEven = true;

			if (_graphics.dash_no_smart) {
				for (let index = 0; index < _graphics.dash_no_smart.length; index++) _graphics.dash_no_smart[index] = (_graphics.m_oCoordTransform.sx * _graphics.dash_no_smart[index] + 0.5) >> 0;

				_graphics.m_oContext.setLineDash(_graphics.dash_no_smart);
				_graphics.dash_no_smart = null;
			}
		}

		let bIsDrawLast = false;
		let path = this.ArrPathCommand;
		shape_drawer._s();

		if (!isRect) {
			for (let j = 0, l = path.length; j < l; ++j) {
				let cmd = path[j];
				switch (cmd.id) {
					case moveTo: {
						bIsDrawLast = true;

						let _x = (_full_trans.TransformPointX(cmd.X, cmd.Y)) >> 0;
						let _y = (_full_trans.TransformPointY(cmd.X, cmd.Y)) >> 0;
						if (bIsEven) {
							_x -= 0.5;
							_y -= 0.5;
						}
						_ctx.moveTo(_x, _y);
						break;
					}
					case lineTo: {
						bIsDrawLast = true;

						let _x = (_full_trans.TransformPointX(cmd.X, cmd.Y)) >> 0;
						let _y = (_full_trans.TransformPointY(cmd.X, cmd.Y)) >> 0;
						if (bIsEven) {
							_x -= 0.5;
							_y -= 0.5;
						}
						_ctx.lineTo(_x, _y);
						break;
					}
					case close: {
						_ctx.closePath();
						break;
					}
				}
			}
		}
		else {
			let minX = 100000;
			let minY = 100000;
			let maxX = -100000;
			let maxY = -100000;
			bIsDrawLast = true;
			for (let j = 0, l = path.length; j < l; ++j) {
				let cmd = path[j];
				switch (cmd.id) {
					case moveTo:
					case lineTo: {
						if (minX > cmd.X) minX = cmd.X;
						if (minY > cmd.Y) minY = cmd.Y;

						if (maxX < cmd.X) maxX = cmd.X;
						if (maxY < cmd.Y) maxY = cmd.Y;

						break;
					}
					default:
						break;
				}
			}

			let _x1 = (_full_trans.TransformPointX(minX, minY)) >> 0;
			let _y1 = (_full_trans.TransformPointY(minX, minY)) >> 0;
			let _x2 = (_full_trans.TransformPointX(maxX, maxY)) >> 0;
			let _y2 = (_full_trans.TransformPointY(maxX, maxY)) >> 0;

			if (bIsEven) _ctx.rect(_x1 + 0.5, _y1 + 0.5, _x2 - _x1, _y2 - _y1); else _ctx.rect(_x1, _y1, _x2 - _x1, _y2 - _y1);
		}

		if (bIsDrawLast) {
			shape_drawer.drawFillStroke(true, this.fill, bIsStroke);
		}

		shape_drawer._e();

		if (false == _old_int) _graphics.SetIntegerGrid(false);
	};
	Path.prototype.isEmpty = function () {
		return this.ArrPathCommandInfo.length <= 0;
	};
	Path.prototype.checkBetweenPolygons = function (oBoundsController, oPolygonWrapper1, oPolygonWrapper2) {
	};
	Path.prototype.checkByPolygon = function (oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds) {
	};
	Path.prototype.transform = function (oTransform, dKoeff) {
	};
	Path.prototype.getSVGPath = function (oTransform, dStartX, dStartY) {
		let aCmds = this.ArrPathCommand;
		let sSVG = "";
		let oPresentation = editor.WordControl.m_oLogicDocument;
		let dSlideWidth = oPresentation.GetWidthMM();
		let dSlideHeight = oPresentation.GetHeightMM();
		let calcX = function (dX, dY) {
			let dX_ = oTransform.TransformPointX(dX, dY);
			return ((((dX_ - dStartX) / dSlideWidth) * 1000 + 0.5 >> 0) / 1000) + "";
		}
		let calcY = function (dX, dY) {
			let dY_ = oTransform.TransformPointY(dX, dY);
			return ((((dY_ - dStartY) / dSlideHeight) * 1000 + 0.5 >> 0) / 1000) + "";
		}
		let nLastCmd = null, nLastX = null, nLastY = null;
		for (let nCmd = 0; nCmd < aCmds.length; ++nCmd) {
			let oCmd = aCmds[nCmd];
			if (sSVG.length > 0) {
				sSVG += " ";
			}
			switch (oCmd.id) {
				case moveTo: {
					if (nLastX !== null && nLastY !== null && AscFormat.fApproxEqual(nLastX, oCmd.X) && AscFormat.fApproxEqual(nLastY, oCmd.Y)) {
						break;
					}
					sSVG += "M ";
					sSVG += calcX(oCmd.X, oCmd.Y);
					sSVG += " ";
					sSVG += calcY(oCmd.X, oCmd.Y);
					nLastX = oCmd.X;
					nLastY = oCmd.Y;
					break;
				}
				case lineTo: {
					if (nLastX !== null && nLastY !== null && AscFormat.fApproxEqual(nLastX, oCmd.X) && AscFormat.fApproxEqual(nLastY, oCmd.Y)) {
						break;
					}
					sSVG += "L ";
					sSVG += calcX(oCmd.X, oCmd.Y);
					sSVG += " ";
					sSVG += calcY(oCmd.X, oCmd.Y);
					nLastX = oCmd.X;
					nLastY = oCmd.Y;
					break;
				}
				case bezier4: {
					sSVG += "C ";
					sSVG += calcX(oCmd.X0, oCmd.Y0);
					sSVG += " ";
					sSVG += calcY(oCmd.X0, oCmd.Y0);
					sSVG += " ";
					sSVG += calcX(oCmd.X1, oCmd.Y1);
					sSVG += " ";
					sSVG += calcY(oCmd.X1, oCmd.Y1);
					sSVG += " ";
					sSVG += calcX(oCmd.X2, oCmd.Y2);
					sSVG += " ";
					sSVG += calcY(oCmd.X2, oCmd.Y2);
					nLastX = oCmd.X2;
					nLastY = oCmd.Y2;
					break;
				}
				case close: {
					sSVG += "Z";
					nLastCmd = null;
					nLastX = null;
					nLastY = null;
					break;
				}
				default: {
					break;
				}
			}

			nLastCmd = oCmd.id;
		}
		return sSVG;
	};
	Path.prototype.isInk = function () {
		const nCmdCount = this.ArrPathCommandInfo.length;
		for (let nCmd = 0; nCmd < nCmdCount; ++nCmd) {
			let oCmd = this.ArrPathCommandInfo[nCmd];
			if (oCmd.id === close) {
				return false;
			}
			if (oCmd.id === arcTo) {
				return false;
			}
		}
		return true;
	};
	Path.prototype.convertToBezierCurves = function (oPath, oTransform, bConvertCurvesOnly) {
		const nCmdCount = this.ArrPathCommandInfo.length;
		let dX0, dY0, dX1, dY1, dX2, dY2;
		let oLastMoveTo = null;
		let dLastX, dLastY;
		let oFirstCmd = this.ArrPathCommand[0];
		if (!oFirstCmd) {
			return null;
		}
		if (oFirstCmd.id !== moveTo) {
			return null;
		}
		for (let nCmd = 0; nCmd < nCmdCount; ++nCmd) {
			let oCmd = this.ArrPathCommand[nCmd];
			switch (oCmd.id) {
				case moveTo: {
					dX0 = oTransform.TransformPointX(oCmd.X, oCmd.Y) * 36000 >> 0;
					dY0 = oTransform.TransformPointY(oCmd.X, oCmd.Y) * 36000 >> 0;
					oPath.moveTo(dX0, dY0);
					oLastMoveTo = oCmd;
					dLastX = oCmd.X;
					dLastY = oCmd.Y;
					break;
				}
				case lineTo: {
					dX0 = oTransform.TransformPointX(oCmd.X, oCmd.Y) * 36000 >> 0;
					dY0 = oTransform.TransformPointY(oCmd.X, oCmd.Y) * 36000 >> 0;
					dX0 = oTransform.TransformPointX(dLastX + (oCmd.X - dLastX) * (1 / 3), dLastY + (oCmd.Y - dLastY) * (1 / 3)) * 36000 >> 0;
					dY0 = oTransform.TransformPointY(dLastX + (oCmd.X - dLastX) * (1 / 3), dLastY + (oCmd.Y - dLastY) * (1 / 3)) * 36000 >> 0;
					dX1 = oTransform.TransformPointX(dLastX + (oCmd.X - dLastX) * (2 / 3), dLastY + (oCmd.Y - dLastY) * (2 / 3)) * 36000 >> 0;
					dY1 = oTransform.TransformPointY(dLastX + (oCmd.X - dLastX) * (2 / 3), dLastY + (oCmd.Y - dLastY) * (2 / 3)) * 36000 >> 0;
					dX2 = oTransform.TransformPointX(oCmd.X, oCmd.Y) * 36000 >> 0;
					dY2 = oTransform.TransformPointY(oCmd.X, oCmd.Y) * 36000 >> 0;
					(bConvertCurvesOnly) ? oPath.lnTo(dX2, dY2) : oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);
					dLastX = oCmd.X;
					dLastY = oCmd.Y;
					break;
				}
				case bezier3: {
					dX0 = oTransform.TransformPointX(oCmd.X0, oCmd.Y0) * 36000 >> 0;
					dY0 = oTransform.TransformPointY(oCmd.X0, oCmd.Y0) * 36000 >> 0;
					dX1 = oTransform.TransformPointX(oCmd.X1, oCmd.Y1) * 36000 >> 0;
					dY1 = oTransform.TransformPointY(oCmd.X1, oCmd.Y1) * 36000 >> 0;
					oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX1, dY1);
					dLastX = oCmd.X1;
					dLastY = oCmd.Y1;
					break;
				}
				case bezier4: {
					dX0 = oTransform.TransformPointX(oCmd.X0, oCmd.Y0) * 36000 >> 0;
					dY0 = oTransform.TransformPointY(oCmd.X0, oCmd.Y0) * 36000 >> 0;
					dX1 = oTransform.TransformPointX(oCmd.X1, oCmd.Y1) * 36000 >> 0;
					dY1 = oTransform.TransformPointY(oCmd.X1, oCmd.Y1) * 36000 >> 0;
					dX2 = oTransform.TransformPointX(oCmd.X2, oCmd.Y2) * 36000 >> 0;
					dY2 = oTransform.TransformPointY(oCmd.X2, oCmd.Y2) * 36000 >> 0;
					oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);
					dLastX = oCmd.X2;
					dLastY = oCmd.Y2;
					break;
				}
				case arcTo: {
					let oPathAccumulator = new AscFormat.PathAccumulator();
					ArcToCurvers(oPathAccumulator, oCmd.stX, oCmd.stY, oCmd.wR, oCmd.hR, oCmd.stAng, oCmd.swAng);
					let aArcToCommands = oPathAccumulator.pathCommand;
					for (let nArcCmd = 0; nArcCmd < aArcToCommands.length; ++nArcCmd) {
						let oArcToCmd = aArcToCommands[nArcCmd];
						switch (oArcToCmd.id) {
							case AscFormat.moveTo: {
								break;
							}
							case AscFormat.bezier4: {
								dX0 = oTransform.TransformPointX(oArcToCmd.X0, oArcToCmd.Y0) * 36000 >> 0;
								dY0 = oTransform.TransformPointY(oArcToCmd.X0, oArcToCmd.Y0) * 36000 >> 0;
								dX1 = oTransform.TransformPointX(oArcToCmd.X1, oArcToCmd.Y1) * 36000 >> 0;
								dY1 = oTransform.TransformPointY(oArcToCmd.X1, oArcToCmd.Y1) * 36000 >> 0;
								dX2 = oTransform.TransformPointX(oArcToCmd.X2, oArcToCmd.Y2) * 36000 >> 0;
								dY2 = oTransform.TransformPointY(oArcToCmd.X2, oArcToCmd.Y2) * 36000 >> 0;
								oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);

								dLastX = oArcToCmd.X2;
								dLastY = oArcToCmd.Y2;
								break;
							}
						}
					}
					break;
				}
				case close: {
					if (!bConvertCurvesOnly & oLastMoveTo) {
						let dXM = oTransform.TransformPointX(oLastMoveTo.X, oLastMoveTo.Y);
						let dYM = oTransform.TransformPointY(oLastMoveTo.X, oLastMoveTo.Y);
						let dLastXM = oTransform.TransformPointX(dLastX, dLastY);
						let dLastYM = oTransform.TransformPointY(dLastX, dLastY);
						dX0 = (dLastXM + (dXM - dLastXM) / 4) * 36000 >> 0;
						dY0 = (dLastYM + (dYM - dLastYM) / 4) * 36000 >> 0;
						dX1 = (dLastXM + 3 * (dXM - dLastXM) / 4) * 36000 >> 0;
						dY1 = (dLastYM + 3 * (dYM - dLastYM) / 4) * 36000 >> 0;
						dX2 = (dXM) * 36000 >> 0;
						dY2 = (dYM) * 36000 >> 0;
						oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);
					}
					oPath.close();
					break;
				}
			}
		}
		oPath.recalculate({}, true);
	};
	Path.prototype.isEqual = function (oPath) {
		if (!oPath) {
			return false;
		}
		if (this.stroke !== oPath.stroke) {
			return false;
		}
		if (this.extrusionOk !== oPath.extrusionOk) {
			return false;
		}
		if (this.fill !== oPath.fill) {
			return false;
		}
		if (this.pathH !== oPath.pathH) {
			return false;
		}
		if (this.pathW !== oPath.pathW) {
			return false;
		}
		if (this.ArrPathCommandInfo.length !== oPath.ArrPathCommandInfo.length) {
			return false;
		}
		for (let nCmd = 0; nCmd < this.ArrPathCommandInfo.length; ++nCmd) {
			let oCmd1 = this.ArrPathCommandInfo[nCmd];
			let oCmd2 = oPath.ArrPathCommandInfo[nCmd];
			if (oCmd1.id !== oCmd2.id) {
				return false;
			}
			switch (oCmd1.id) {
				case moveTo:
				case lineTo: {
					if (oCmd1.X !== oCmd2.X) {
						return false;
					}
					if (oCmd1.Y !== oCmd2.Y) {
						return false;
					}
					break;
				}
				case bezier3: {
					if (oCmd1.X0 !== oCmd2.X0) {
						return false;
					}
					if (oCmd1.X1 !== oCmd2.X1) {
						return false;
					}
					if (oCmd1.Y0 !== oCmd2.Y0) {
						return false;
					}
					if (oCmd1.Y1 !== oCmd2.Y1) {
						return false;
					}
					break;
				}
				case bezier4: {

					if (oCmd1.X0 !== oCmd2.X0) {
						return false;
					}
					if (oCmd1.X1 !== oCmd2.X1) {
						return false;
					}
					if (oCmd1.X2 !== oCmd2.X2) {
						return false;
					}
					if (oCmd1.Y0 !== oCmd2.Y0) {
						return false;
					}
					if (oCmd1.Y1 !== oCmd2.Y1) {
						return false;
					}
					if (oCmd1.Y2 !== oCmd2.Y2) {
						return false;
					}
					break;
				}
				case arcTo: {

					if (oCmd1.wR !== oCmd2.wR) {
						return false;
					}
					if (oCmd1.hR !== oCmd2.hR) {
						return false;
					}
					if (oCmd1.stAng !== oCmd2.stAng) {
						return false;
					}
					if (oCmd1.swAng !== oCmd2.swAng) {
						return false;
					}
					break;
				}
				case close: {
					break;
				}
				default: {
					break;
				}
			}
		}
		return true;
	};

	function CPathCmd() {
		AscFormat.CBaseNoIdObject.call(this);
		this.pts = [];
	}

	AscFormat.InitClass(CPathCmd, AscFormat.CBaseNoIdObject, 0);

	function CheckPointByPaths(dX, dY, dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2) {
		let cX, cY, point, topX, topY, bottomX, bottomY;
		cX = (dX - dMinX) / dWidth;
		cY = (dY - dMinY) / dHeight;
		if (cX > 1) {
			cX = 1;
		}
		if (cX < 0) {
			cX = 0;
		}
		point = oPolygonWrapper1.getPointOnPolygon(cX);
		topX = point.x;
		topY = point.y;
		point = oPolygonWrapper2.getPointOnPolygon(cX);
		bottomX = point.x;
		bottomY = point.y;
		return {x: topX + cY * (bottomX - topX), y: topY + cY * (bottomY - topY)};
	}

	function Path2(oPathMemory) {
		this.stroke = null;
		this.extrusionOk = null;
		this.fill = null;
		this.pathH = null;
		this.pathW = null;


		this.startPos = 0;

		this.PathMemory = oPathMemory;

		this.lastX = null;
		this.lastY = null;
	}

	Path2.prototype.getMemoryLength = function () {
		return this.getArrPathCommand()[this.startPos];
	};
	Path2.prototype.isEmpty = function () {
		return this.getMemoryLength() === 0;
	};
	Path2.prototype.isInk = function () {
		return false;
	};
	Path2.prototype.getArrPathCommand = function () {
		return this.PathMemory.ArrPathCommand;
	};
	Path2.prototype.writeNumber = function (dValue) {
		this.PathMemory.IncrementNumberInPos(this.startPos);
		this.PathMemory.WriteNumber(dValue);
	};
	Path2.prototype.setStroke = function (pr) {
		this.stroke = pr;
	};
	Path2.prototype.setExtrusionOk = function (pr) {

		this.extrusionOk = pr;
	};
	Path2.prototype.setFill = function (pr) {

		this.fill = pr;
	};
	Path2.prototype.setPathH = function (pr) {

		this.pathH = pr;
	};
	Path2.prototype.setPathW = function (pr) {

		this.pathW = pr;
	};
	Path2.prototype.addPathCommand = function (cmd) {
		this.getArrPathCommand().push(cmd);
	};
	Path2.prototype.writeCommandType = function (cmd) {
		this.writeNumber(cmd);
	};
	Path2.prototype.writePoint = function (x, y) {
		this.lastX = x * 10e-10;
		this.lastY = y * 10e-10;
		this.writeNumber(this.lastX);
		this.writeNumber(this.lastY);
	};
	Path2.prototype.moveTo = function (x, y) {
		this.writeCommandType(moveTo);
		this.writePoint(x, y);
	};
	Path2.prototype.lnTo = function (x, y) {
		this.writeCommandType(lineTo);
		this.writePoint(x, y);
	};
	Path2.prototype.arcTo = function (wR, hR, stAng, swAng) {
		let nSign = swAng < 0 ? -1 : 1;
		if (AscFormat.fApproxEqual(Math.abs(swAng), 21600000)) {
			swAng = nSign * (21600000 - 1);
		}
		let a1 = stAng;
		let a2 = stAng + swAng;
		let a3 = swAng;


		stAng = Math.atan2(10e-10 * Math.sin(a1 * cToRad), 10e-10 * Math.cos(a1 * cToRad)) / cToRad;
		swAng = Math.atan2(10e-10 * Math.sin(a2 * cToRad), 10e-10 * Math.cos(a2 * cToRad)) / cToRad - stAng;

		if ((swAng > 0) && (a3 < 0)) swAng -= 21600000;
		if ((swAng < 0) && (a3 > 0)) swAng += 21600000;
		if (swAng == 0 && a3 != 0) swAng = 21600000;

		this.writeCommandType(arcTo);
		//TODO: calculate correct the last point
		this.writeNumber(this.lastX);
		this.writeNumber(this.lastY);
		this.writeNumber(wR * 10e-10);
		this.writeNumber(hR * 10e-10);
		this.writeNumber(stAng * cToRad);
		this.writeNumber(swAng * cToRad);
	};
	Path2.prototype.quadBezTo = function (x0, y0, x1, y1) {
		this.writeCommandType(bezier3);
		this.writePoint(x0, y0);
		this.writePoint(x1, y1);
	};
	Path2.prototype.cubicBezTo = function (x0, y0, x1, y1, x2, y2) {
		this.writeCommandType(bezier4);
		this.writePoint(x0, y0);
		this.writePoint(x1, y1);
		this.writePoint(x2, y2);
	};
	Path2.prototype.close = function () {
		this.writeCommandType(close);
	};
	Path2.prototype.draw = function (shape_drawer) {
		if (this.isEmpty()) {
			return;
		}
		if (shape_drawer.bIsCheckBounds === true && this.fill == "none") {
			// это для текстур
			return;
		}
		let bIsDrawLast = false;
		let path = this.getArrPathCommand();
		shape_drawer._s();
		let i = 0;
		let len = this.getMemoryLength();
		while (i < len) {
			let cmd = path[this.startPos + i + 1];
			switch (cmd) {
				case moveTo: {
					bIsDrawLast = true;
					shape_drawer._m(path[this.startPos + i + 2], path[this.startPos + i + 3]);
					i += 3;
					break;
				}
				case lineTo: {
					bIsDrawLast = true;
					shape_drawer._l(path[this.startPos + i + 2], path[this.startPos + i + 3]);
					i += 3;
					break;
				}
				case bezier3: {
					bIsDrawLast = true;
					shape_drawer._c2(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5]);
					i += 5;
					break;
				}
				case bezier4: {
					bIsDrawLast = true;
					shape_drawer._c(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
					i += 7;
					break;
				}
				case arcTo: {
					bIsDrawLast = true;
					ArcToCurvers(shape_drawer, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
					i += 7;
					break;
				}
				case close: {
					shape_drawer._z();
					i += 1;
					break;
				}
			}
		}

		if (bIsDrawLast) {
			shape_drawer.drawFillStroke(true, "normal", this.stroke && !shape_drawer.bIsNoStrokeAttack);
		}

		shape_drawer._e();
	};
	Path2.prototype.hitInInnerArea = function (canvasContext, x, y) {
		let path = this.getArrPathCommand();
		canvasContext.beginPath();
		let i = 0;
		let len = this.getMemoryLength();
		while (i < len) {
			let cmd = path[this.startPos + i + 1];
			switch (cmd) {
				case moveTo: {
					canvasContext.moveTo(path[this.startPos + i + 2], path[this.startPos + i + 3]);
					i += 3;
					break;
				}
				case lineTo: {
					canvasContext.lineTo(path[this.startPos + i + 2], path[this.startPos + i + 3]);
					i += 3;
					break;
				}
				case bezier3: {
					canvasContext.quadraticCurveTo(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5]);
					i += 5;
					break;
				}
				case bezier4: {
					canvasContext.bezierCurveTo(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
					i += 7;
					break;
				}
				case arcTo: {
					ArcToOnCanvas(canvasContext, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
					i += 7;
					break;
				}
				case close: {
					canvasContext.closePath();
					i += 1;
					break;
				}
			}
		}
		canvasContext.closePath();
		return !!canvasContext.isPointInPath(x, y);
	};
	Path2.prototype.hitInPath = function (canvasContext, x, y) {
		let _last_x, _last_y;
		let _begin_x, _begin_y;

		let path = this.getArrPathCommand();
		let i = 0;
		let len = this.getMemoryLength();
		while (i < len) {
			let cmd = path[this.startPos + i + 1];
			switch (cmd) {
				case moveTo: {
					canvasContext.moveTo(path[this.startPos + i + 2], path[this.startPos + i + 3]);
					_last_x = path[this.startPos + i + 2];
					_last_y = path[this.startPos + i + 3];
					_begin_x = path[this.startPos + i + 2];
					_begin_y = path[this.startPos + i + 3];
					i += 3;
					break;
				}
				case lineTo: {
					if (HitInLine(canvasContext, x, y, _last_x, _last_y, path[this.startPos + i + 2], path[this.startPos + i + 3])) return true;
					_last_x = path[this.startPos + i + 2];
					_last_y = path[this.startPos + i + 3];
					i += 3;
					break;
				}
				case bezier3: {
					canvasContext.quadraticCurveTo(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5]);
					if (HitInBezier3(canvasContext, x, y, _last_x, _last_y, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5])) return true;
					_last_x = path[this.startPos + i + 4];
					_last_y = path[this.startPos + i + 5];
					i += 5;
					break;
				}
				case bezier4: {
					if (HitInBezier4(canvasContext, x, y, _last_x, _last_y, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7])) return true;
					_last_x = path[this.startPos + i + 6];
					_last_y = path[this.startPos + i + 7];
					i += 7;
					break;
				}
				case arcTo: {
					if (HitToArc(canvasContext, x, y, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7])) return true;
					_last_x = (path[this.startPos + i + 2] - path[this.startPos + i + 4] * Math.cos(path[this.startPos + i + 6]) + path[this.startPos + i + 4] * Math.cos(path[this.startPos + i + 7]));
					_last_y = (path[this.startPos + i + 3] - path[this.startPos + i + 5] * Math.sin(path[this.startPos + i + 6]) + path[this.startPos + i + 5] * Math.sin(path[this.startPos + i + 7]));
					i += 7;
					break;
				}
				case close: {
					if (HitInLine(canvasContext, x, y, _last_x, _last_y, _begin_x, _begin_y)) return true;
					i += 1;
					break;
				}
			}
		}
		return false;
	};
	Path2.prototype.drawTracks = function (drawingDocument, transform) {

		let oApi = Asc.editor || editor;
		let isDrawHandles = oApi ? oApi.isShowShapeAdjustments() : true;
		let i = 0;
		let len = this.getMemoryLength();

		let path = this.getArrPathCommand();
		let dDist = 0;
		while (i < len) {
			let cmd = path[this.startPos + i + 1];
			switch (cmd) {
				case moveTo: {
					drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i + 2] - dDist, path[this.startPos + i + 3] - dDist, 2 * dDist, 2 * dDist, false, false, undefined, isDrawHandles);
					i += 3;
					break;
				}
				case lineTo: {
					drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i + 2] - dDist, path[this.startPos + i + 3] - dDist, 2 * dDist, 2 * dDist, false, false, undefined, isDrawHandles);
					i += 3;
					break;
				}
				case bezier3: {
					//  drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i+2] - dDist, path[this.startPos + i + 3] - dDist, 2*dDist, 2*dDist, false, false);
					drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i + 4] - dDist, path[this.startPos + i + 5] - dDist, 2 * dDist, 2 * dDist, false, false, undefined, isDrawHandles);
					i += 5;
					break;
				}
				case bezier4: {
					//  drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i+2] - dDist, path[this.startPos + i + 3] - dDist, 2*dDist,2*dDist, false, false);
					// drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i+4] - dDist, path[this.startPos + i + 5] - dDist, 2*dDist,2*dDist, false, false); i+=7;
					drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i + 6] - dDist, path[this.startPos + i + 7] - dDist, 2 * dDist, 2 * dDist, false, false, undefined, isDrawHandles);
					i += 7;
					break;
				}
				case arcTo: {
					let path_accumulator = new AscFormat.PathAccumulator();
					ArcToCurvers(path_accumulator, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
					let arc_to_path_commands = path_accumulator.pathCommand;
					let lastX, lastY;
					for (let arc_to_path_index = 0; arc_to_path_index < arc_to_path_commands.length; ++arc_to_path_index) {
						let cur_arc_to_command = arc_to_path_commands[arc_to_path_index];
						switch (cur_arc_to_command.id) {
							case AscFormat.moveTo: {
								lastX = cur_arc_to_command.X;
								lastY = cur_arc_to_command.Y;
								//  drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, cur_arc_to_command.X - dDist, cur_arc_to_command.Y - dDist,  2*dDist, 2*dDist, false, false);
								break;
							}
							case AscFormat.bezier4: {

								lastX = cur_arc_to_command.X2;
								lastY = cur_arc_to_command.Y2;
								//drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, cur_arc_to_command.X0 - dDist, cur_arc_to_command.Y0 - dDist,  2*dDist, 2*dDist, false, false);
								//drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, cur_arc_to_command.X2 - dDist, cur_arc_to_command.Y2 - dDist,  2*dDist, 2*dDist, false, false);
								break;
							}
						}
					}
					drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, lastX - dDist, lastY - dDist, 2 * dDist, 2 * dDist, false, false, undefined, isDrawHandles);
					i += 7;
					break;
				}
				case close: {
					i += 1;
					break;
				}
			}
		}

	};
	Path2.prototype.getCommandByIndex = function (idx) {
		let i = 0;
		let path = this.getArrPathCommand();
		let len = this.getMemoryLength();
		let commandIndex = 0;
		while (i < len) {
			let cmd = path[this.startPos + i + 1];
			switch (cmd) {
				case moveTo: {
					if (idx === commandIndex) {
						return {id: moveTo, X: path[this.startPos + i + 2], Y: path[this.startPos + i + 3]};
					}
					i += 3;
					break;
				}
				case lineTo: {
					if (idx === commandIndex) {
						return {id: moveTo, X: path[this.startPos + i + 2], Y: path[this.startPos + i + 3]};
					}
					i += 3;
					break;
				}
				case bezier3: {
					if (idx === commandIndex) {
						return {
							id: bezier3,
							X0: path[this.startPos + i + 2],
							Y0: path[this.startPos + i + 3],
							X1: path[this.startPos + i + 4],
							Y1: path[this.startPos + i + 5]
						};
					}
					i += 5;
					break;
				}
				case bezier4: {
					if (idx === commandIndex) {
						return {
							id: bezier4,
							X0: path[this.startPos + i + 2],
							Y0: path[this.startPos + i + 3],
							X1: path[this.startPos + i + 4],
							Y1: path[this.startPos + i + 5],
							X2: path[this.startPos + i + 6],
							Y2: path[this.startPos + i + 7]
						};
					}
					i += 7;
					break;
				}
				case arcTo: {
					if (idx === commandIndex) {
						return {
							id: arcTo,
							stX: path[this.startPos + i + 2],
							stY: path[this.startPos + i + 3],
							wR: path[this.startPos + i + 4],
							hR: path[this.startPos + i + 5],
							stAng: path[this.startPos + i + 6],
							swAng: path[this.startPos + i + 7]
						};
					}

					i += 7;
					break;
				}
				case close: {

					if (idx === commandIndex) {
						return {id: close};
					}
					i += 1;
					break;
				}
			}
			++commandIndex;
		}
		return null;
	};
	Path2.prototype.check_bounds = function (shape_drawer) {
		let path = this.getArrPathCommand();

		let i = 0;
		let len = this.getMemoryLength();
		while (i < len) {
			let cmd = path[this.startPos + i + 1];
			switch (cmd) {
				case moveTo: {

					shape_drawer._m(path[this.startPos + i + 2], path[this.startPos + i + 3]);
					i += 3;
					break;
				}
				case lineTo: {

					shape_drawer._l(path[this.startPos + i + 2], path[this.startPos + i + 3]);
					i += 3;
					break;
				}
				case bezier3: {

					shape_drawer._c2(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5]);
					i += 5;
					break;
				}
				case bezier4: {

					shape_drawer._c(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
					i += 7;
					break;
				}
				case arcTo: {

					ArcToCurvers(shape_drawer, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
					i += 7;
					break;
				}
				case close: {
					shape_drawer._z();
					i += 1;
					break;
				}
			}
		}

	};
	Path2.prototype.isSmartLine = function () {
		let i = 0;
		let path = this.getArrPathCommand();
		let len = this.getMemoryLength();
		let commandIndex = 0;
		while (i < len) {
			if (commandIndex > 1) {
				return false;
			}
			let cmd = path[this.startPos + i + 1];
			switch (cmd) {
				case moveTo: {
					if (0 !== commandIndex) {
						return false;
					}
					i += 3;
					break;
				}
				case lineTo: {
					if (1 !== commandIndex) {
						return false;
					}
					i += 3;
					break;
				}
				default: {
					return false;
				}
			}
			++commandIndex;
		}

		return true;
	};
	Path2.prototype.isSmartRect = function () {
		let i = 0;
		let path = this.getArrPathCommand();
		let len = path[this.startPos];
		let commandIndex = 0;
		let x0, y0, x1, y1, x2, y2, x3, y3, x4, y4, isCommand4Close = false;
		while (i < len) {
			if (commandIndex > 4) {
				return false;
			}
			let cmd = path[this.startPos + i + 1];
			switch (cmd) {
				case moveTo: {
					if (0 !== commandIndex) {
						return false;
					}
					x0 = path[this.startPos + i + 2];
					y0 = path[this.startPos + i + 3];
					i += 3;
					break;
				}
				case lineTo: {
					if (commandIndex === 1) {
						x1 = path[this.startPos + i + 2];
						y1 = path[this.startPos + i + 3];
					}
					else if (commandIndex === 2) {
						x2 = path[this.startPos + i + 2];
						y2 = path[this.startPos + i + 3];
					}
					else if (commandIndex === 3) {
						x3 = path[this.startPos + i + 2];
						y3 = path[this.startPos + i + 3];
					}
					else if (commandIndex === 4) {
						x4 = path[this.startPos + i + 2];
						y4 = path[this.startPos + i + 3];
					}
					i += 3;
					break;
				}
				case close: {
					if (4 !== commandIndex) {
						return false;
					}
					isCommand4Close = true;
					break;
				}
				default: {
					return false;
				}
			}
			++commandIndex;
		}

		if (AscFormat.fApproxEqual(x0, x1)) {
			if (AscFormat.fApproxEqual(y1, y2)) {
				if (AscFormat.fApproxEqual(x2, x3) && AscFormat.fApproxEqual(y3, y0)) {
					if (isCommand4Close) return true;

					if (AscFormat.fApproxEqual(x0, x4) && AscFormat.fApproxEqual(y0, y4)) {
						return true;
					}
				}
			}
		}
		else if (AscFormat.fApproxEqual(y0, y1)) {
			if (AscFormat.fApproxEqual(x1, x2)) {
				if (AscFormat.fApproxEqual(y2, y3) && AscFormat.fApproxEqual(x3, x0)) {
					if (isCommand4Close) return true;

					if (AscFormat.fApproxEqual(x0, x4) && AscFormat.fApproxEqual(y0, y4)) {
						return true;
					}
				}
			}
		}

		return false;
	};
	Path2.prototype.drawSmart = function (shape_drawer) {
		let _graphics = shape_drawer.Graphics;
		let _full_trans = _graphics.m_oFullTransform;

		if (!_graphics || !_full_trans || undefined == _graphics.m_bIntegerGrid || true === shape_drawer.bIsNoSmartAttack) return this.draw(shape_drawer);

		let bIsTransformed = (_full_trans.shx == 0 && _full_trans.shy == 0) ? false : true;

		if (bIsTransformed) return this.draw(shape_drawer);

		let isLine = this.isSmartLine();
		let isRect = false;
		if (!isLine) isRect = this.isSmartRect();

		if (window["NATIVE_EDITOR_ENJINE"] || (!isLine && !isRect && !shape_drawer.bDrawSmartAttack)) return this.draw(shape_drawer);

		let _old_int = _graphics.m_bIntegerGrid;

		if (false == _old_int) _graphics.SetIntegerGrid(true);

		let dKoefMMToPx = Math.max(_graphics.m_oCoordTransform.sx, 0.001);

		let _ctx = _graphics.m_oContext;
		let bIsStroke = (shape_drawer.bIsNoStrokeAttack || (this.stroke !== true)) ? false : true;
		let bIsEven = false;
		if (bIsStroke) {
			let _lineWidth = Math.max((shape_drawer.StrokeWidth * dKoefMMToPx + 0.5) >> 0, 1);
			_ctx.lineWidth = _lineWidth;

			if ((_lineWidth & 0x01) == 0x01) bIsEven = true;
		}

		let bIsDrawLast = false;
		let path = this.getArrPathCommand();
		shape_drawer._s();

		if (!isRect) {
			let i = 0;
			let len = path[this.startPos];
			let X, Y;
			while (i < len) {
				let cmd = path[this.startPos + i + 1];
				switch (cmd) {
					case moveTo: {
						bIsDrawLast = true;
						X = path[this.startPos + i + 2];
						Y = path[this.startPos + i + 3];
						let _x = (_full_trans.TransformPointX(X, Y)) >> 0;
						let _y = (_full_trans.TransformPointY(X, Y)) >> 0;
						if (bIsEven) {
							_x -= 0.5;
							_y -= 0.5;
						}
						_ctx.moveTo(_x, _y);

						if (_graphics.ArrayPoints != null) {
							_graphics.ArrayPoints.push({x: X, y: Y});
						}

						i += 3;
						break;
					}
					case lineTo: {
						bIsDrawLast = true;
						X = path[this.startPos + i + 2];
						Y = path[this.startPos + i + 3];
						let _x = (_full_trans.TransformPointX(X, Y)) >> 0;
						let _y = (_full_trans.TransformPointY(X, Y)) >> 0;
						if (bIsEven) {
							_x -= 0.5;
							_y -= 0.5;
						}
						_ctx.lineTo(_x, _y);

						if (_graphics.ArrayPoints != null) {
							_graphics.ArrayPoints.push({x: X, y: Y});
						}

						i += 3;
						break;
					}
					case bezier3: {
						bIsDrawLast = true;

						i += 5;
						break;
					}
					case bezier4: {
						bIsDrawLast = true;
						i += 7;
						break;
					}
					case arcTo: {
						bIsDrawLast = true;
						i += 7;
						break;
					}
					case close: {
						_ctx.closePath();
						i += 1;
						break;
					}
				}
			}
		}
		else {
			let minX = 100000;
			let minY = 100000;
			let maxX = -100000;
			let maxY = -100000;
			bIsDrawLast = true;
			let i = 0;
			let len = path[this.startPos], X, Y;
			while (i < len) {
				let cmd = path[this.startPos + i + 1];
				switch (cmd) {
					case moveTo:
					case lineTo: {
						bIsDrawLast = true;
						X = path[this.startPos + i + 2];
						Y = path[this.startPos + i + 3];
						if (minX > X) minX = X;
						if (minY > Y) minY = Y;

						if (maxX < X) maxX = X;
						if (maxY < Y) maxY = Y;

						i += 3;
						break;
					}
					case bezier3: {
						bIsDrawLast = true;
						i += 5;
						break;
					}
					case bezier4: {
						bIsDrawLast = true;
						i += 7;
						break;
					}
					case arcTo: {
						bIsDrawLast = true;
						i += 7;
						break;
					}
					case close: {
						i += 1;
						break;
					}
				}
			}

			let _x1 = (_full_trans.TransformPointX(minX, minY)) >> 0;
			let _y1 = (_full_trans.TransformPointY(minX, minY)) >> 0;
			let _x2 = (_full_trans.TransformPointX(maxX, maxY)) >> 0;
			let _y2 = (_full_trans.TransformPointY(maxX, maxY)) >> 0;

			if (bIsEven) _ctx.rect(_x1 + 0.5, _y1 + 0.5, _x2 - _x1, _y2 - _y1); else _ctx.rect(_x1, _y1, _x2 - _x1, _y2 - _y1);
		}

		if (bIsDrawLast) {
			shape_drawer.isArrPix = true;
			shape_drawer.drawFillStroke(true, this.fill, bIsStroke);
			shape_drawer.isArrPix = false;
		}

		shape_drawer._e();

		if (false == _old_int) _graphics.SetIntegerGrid(false);
	};
	Path2.prototype.recalculate = function (gdLst, bResetPathsInfo) {
	};
	Path2.prototype.recalculate2 = function (gdLst, bResetPathsInfo) {
	};
	Path2.prototype.transformPointPolygon = function (x, y, oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds) {
		let oRet = {x: 0, y: 0}, y0, y1, cX, oPointOnPolygon, x1t, y1t, dX, dY, x0t, y0t;
		y0 = y;//dKoeff;
		if (bFlag) {
			y1 = 0;
			if (oBounds) {
				y0 -= oBounds.min_y;
			}
		}
		else {
			y1 = ContentHeight * dKoeff;
			if (oBounds) {
				y1 = (oBounds.max_y - oBounds.min_y);
				y0 -= oBounds.min_y;
			}
		}
		cX = x / XLimit;
		oPointOnPolygon = oPolygon.getPointOnPolygon(cX, true);
		x1t = oPointOnPolygon.x;
		y1t = oPointOnPolygon.y;
		dX = oPointOnPolygon.oP2.x - oPointOnPolygon.oP1.x;
		dY = oPointOnPolygon.oP2.y - oPointOnPolygon.oP1.y;

		if (bFlag) {
			dX = -dX;
			dY = -dY;
		}
		let dNorm = Math.sqrt(dX * dX + dY * dY);

		if (bFlag) {
			x0t = x1t + (dY / dNorm) * (y0);
			y0t = y1t - (dX / dNorm) * (y0);
		}
		else {

			x0t = x1t + (dY / dNorm) * (y1 - y0);
			y0t = y1t - (dX / dNorm) * (y1 - y0);
		}
		oRet.x = x0t;
		oRet.y = y0t;
		return oRet;
	};
	Path2.prototype.checkBetweenPolygons = function (oBoundsController, oPolygonWrapper1, oPolygonWrapper2) {

		let path = this.getArrPathCommand();
		let i = 0;
		let len = path[this.startPos];
		let p;
		let dMinX = oBoundsController.min_x, dMinY = oBoundsController.min_y,
			dWidth = oBoundsController.max_x - oBoundsController.min_x,
			dHeight = oBoundsController.max_y - oBoundsController.min_y;
		while (i < len) {
			let cmd = path[this.startPos + i + 1];
			switch (cmd) {
				case moveTo:
				case lineTo: {
					p = CheckPointByPaths(path[this.startPos + i + 2], path[this.startPos + i + 3], dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2);
					path[this.startPos + i + 2] = p.x;
					path[this.startPos + i + 3] = p.y;
					i += 3;
					break;
				}
				case bezier3: {

					p = CheckPointByPaths(path[this.startPos + i + 2], path[this.startPos + i + 3], dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2);
					path[this.startPos + i + 2] = p.x;
					path[this.startPos + i + 3] = p.y;
					p = CheckPointByPaths(path[this.startPos + i + 4], path[this.startPos + i + 5], dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2);
					path[this.startPos + i + 4] = p.x;
					path[this.startPos + i + 5] = p.y;
					i += 5;
					break;
				}
				case bezier4: {
					p = CheckPointByPaths(path[this.startPos + i + 2], path[this.startPos + i + 3], dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2);
					path[this.startPos + i + 2] = p.x;
					path[this.startPos + i + 3] = p.y;
					p = CheckPointByPaths(path[this.startPos + i + 4], path[this.startPos + i + 5], dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2);
					path[this.startPos + i + 4] = p.x;
					path[this.startPos + i + 5] = p.y;
					p = CheckPointByPaths(path[this.startPos + i + 6], path[this.startPos + i + 7], dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2);
					path[this.startPos + i + 6] = p.x;
					path[this.startPos + i + 7] = p.y;
					i += 7;
					break;
				}
				case arcTo: {
					i += 7;
					break;
				}
				case close: {
					i += 1;
					break;
				}
			}
		}
	};
	Path2.prototype.checkByPolygon = function (oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds) {

		let path = this.getArrPathCommand();
		let i = 0;
		let len = path[this.startPos];
		let p;
		while (i < len) {
			let cmd = path[this.startPos + i + 1];
			switch (cmd) {
				case moveTo:
				case lineTo: {
					p = this.transformPointPolygon(path[this.startPos + i + 2], path[this.startPos + i + 3], oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds);
					path[this.startPos + i + 2] = p.x;
					path[this.startPos + i + 3] = p.y;
					i += 3;
					break;
				}
				case bezier3: {

					p = this.transformPointPolygon(path[this.startPos + i + 2], path[this.startPos + i + 3], oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds);
					path[this.startPos + i + 2] = p.x;
					path[this.startPos + i + 3] = p.y;
					p = this.transformPointPolygon(path[this.startPos + i + 4], path[this.startPos + i + 5], oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds);
					path[this.startPos + i + 4] = p.x;
					path[this.startPos + i + 5] = p.y;
					i += 5;
					break;
				}
				case bezier4: {
					p = this.transformPointPolygon(path[this.startPos + i + 2], path[this.startPos + i + 3], oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds);
					path[this.startPos + i + 2] = p.x;
					path[this.startPos + i + 3] = p.y;
					p = this.transformPointPolygon(path[this.startPos + i + 4], path[this.startPos + i + 5], oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds);
					path[this.startPos + i + 4] = p.x;
					path[this.startPos + i + 5] = p.y;
					p = this.transformPointPolygon(path[this.startPos + i + 6], path[this.startPos + i + 7], oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds);
					path[this.startPos + i + 6] = p.x;
					path[this.startPos + i + 7] = p.y;
					i += 7;
					break;
				}
				case arcTo: {
					i += 7;
					break;
				}
				case close: {
					i += 1;
					break;
				}
			}
		}
	};
	Path2.prototype.transform = function (oTransform, dKoeff) {
		let path = this.getArrPathCommand();
		let i = 0;
		let len = path[this.startPos];
		let p, x, y;
		while (i < len) {
			let cmd = path[this.startPos + i + 1];
			switch (cmd) {
				case moveTo:
				case lineTo: {
					x = oTransform.TransformPointX(path[this.startPos + i + 2] * dKoeff, path[this.startPos + i + 3] * dKoeff);
					y = oTransform.TransformPointY(path[this.startPos + i + 2] * dKoeff, path[this.startPos + i + 3] * dKoeff);
					path[this.startPos + i + 2] = x;
					path[this.startPos + i + 3] = y;
					i += 3;
					break;
				}
				case bezier3: {
					x = oTransform.TransformPointX(path[this.startPos + i + 2] * dKoeff, path[this.startPos + i + 3] * dKoeff);
					y = oTransform.TransformPointY(path[this.startPos + i + 2] * dKoeff, path[this.startPos + i + 3] * dKoeff);
					path[this.startPos + i + 2] = x;
					path[this.startPos + i + 3] = y;
					x = oTransform.TransformPointX(path[this.startPos + i + 4] * dKoeff, path[this.startPos + i + 5] * dKoeff);
					y = oTransform.TransformPointY(path[this.startPos + i + 4] * dKoeff, path[this.startPos + i + 5] * dKoeff);
					path[this.startPos + i + 4] = x;
					path[this.startPos + i + 5] = y;
					i += 5;
					break;
				}
				case bezier4: {
					x = oTransform.TransformPointX(path[this.startPos + i + 2] * dKoeff, path[this.startPos + i + 3] * dKoeff);
					y = oTransform.TransformPointY(path[this.startPos + i + 2] * dKoeff, path[this.startPos + i + 3] * dKoeff);
					path[this.startPos + i + 2] = x;
					path[this.startPos + i + 3] = y;
					x = oTransform.TransformPointX(path[this.startPos + i + 4] * dKoeff, path[this.startPos + i + 5] * dKoeff);
					y = oTransform.TransformPointY(path[this.startPos + i + 4] * dKoeff, path[this.startPos + i + 5] * dKoeff);
					path[this.startPos + i + 4] = x;
					path[this.startPos + i + 5] = y;
					x = oTransform.TransformPointX(path[this.startPos + i + 6] * dKoeff, path[this.startPos + i + 7] * dKoeff);
					y = oTransform.TransformPointY(path[this.startPos + i + 6] * dKoeff, path[this.startPos + i + 7] * dKoeff);
					path[this.startPos + i + 6] = x;
					path[this.startPos + i + 7] = y;
					i += 7;
					break;
				}
				case arcTo: {
					i += 7;
					break;
				}
				case close: {
					i += 1;
					break;
				}
			}
		}
	};
	Path2.prototype.convertToBezierCurves = function (oPath, oTransform) {
		let dX0, dY0, dX1, dY1, dX2, dY2;
		let oLastMoveTo = null;
		let dLastX, dLastY;
		let path = this.getArrPathCommand();
		let i = 0;
		let len = path[this.startPos];
		if (len === 0) {
			return null;
		}
		let X, Y, X0, Y0, X1, Y1, X2, Y2, stX, stY, wR, hR, stAng, swAng;
		while (i < len) {
			let cmd = path[this.startPos + i + 1];
			switch (cmd) {
				case moveTo: {
					X = path[this.startPos + i + 2];
					Y = path[this.startPos + i + 3];
					dX0 = oTransform.TransformPointX(X, Y) * 36000 >> 0;
					dY0 = oTransform.TransformPointY(X, Y) * 36000 >> 0;
					oPath.moveTo(dX0, dY0);
					oLastMoveTo = i;
					dLastX = X;
					dLastY = Y;
					i += 3;
					break;
				}
				case lineTo: {
					X = path[this.startPos + i + 2];
					Y = path[this.startPos + i + 3];
					dX0 = oTransform.TransformPointX(X, Y) * 36000 >> 0;
					dY0 = oTransform.TransformPointY(X, Y) * 36000 >> 0;
					dX0 = oTransform.TransformPointX(dLastX + (X - dLastX) * (1 / 3), dLastY + (Y - dLastY) * (1 / 3)) * 36000 >> 0;
					dY0 = oTransform.TransformPointY(dLastX + (X - dLastX) * (1 / 3), dLastY + (Y - dLastY) * (1 / 3)) * 36000 >> 0;
					dX1 = oTransform.TransformPointX(dLastX + (X - dLastX) * (2 / 3), dLastY + (Y - dLastY) * (2 / 3)) * 36000 >> 0;
					dY1 = oTransform.TransformPointY(dLastX + (X - dLastX) * (2 / 3), dLastY + (Y - dLastY) * (2 / 3)) * 36000 >> 0;
					dX2 = oTransform.TransformPointX(X, Y) * 36000 >> 0;
					dY2 = oTransform.TransformPointY(X, Y) * 36000 >> 0;
					oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);
					dLastX = X;
					dLastY = Y;
					i += 3;
					break;
				}
				case bezier3: {
					X0 = path[this.startPos + i + 2];
					Y0 = path[this.startPos + i + 3];
					X1 = path[this.startPos + i + 4];
					Y1 = path[this.startPos + i + 5];
					dX0 = oTransform.TransformPointX(X0, Y0) * 36000 >> 0;
					dY0 = oTransform.TransformPointY(X0, Y0) * 36000 >> 0;
					dX1 = oTransform.TransformPointX(X1, Y1) * 36000 >> 0;
					dY1 = oTransform.TransformPointY(X1, Y1) * 36000 >> 0;
					oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX1, dY1);
					dLastX = X1;
					dLastY = Y1;

					i += 5;
					break;
				}
				case bezier4: {

					X0 = path[this.startPos + i + 2];
					Y0 = path[this.startPos + i + 3];
					X1 = path[this.startPos + i + 4];
					Y1 = path[this.startPos + i + 5];
					X2 = path[this.startPos + i + 6];
					Y2 = path[this.startPos + i + 7];

					dX0 = oTransform.TransformPointX(X0, Y0) * 36000 >> 0;
					dY0 = oTransform.TransformPointY(X0, Y0) * 36000 >> 0;
					dX1 = oTransform.TransformPointX(X1, Y1) * 36000 >> 0;
					dY1 = oTransform.TransformPointY(X1, Y1) * 36000 >> 0;
					dX2 = oTransform.TransformPointX(X2, Y2) * 36000 >> 0;
					dY2 = oTransform.TransformPointY(X2, Y2) * 36000 >> 0;
					oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);
					dLastX = X2;
					dLastY = Y2;
					i += 7;
					break;
				}
				case arcTo: {
					stX = path[this.startPos + i + 2];
					stY = path[this.startPos + i + 3]
					wR = path[this.startPos + i + 4];
					hR = path[this.startPos + i + 5];
					stAng = path[this.startPos + i + 6];
					swAng = path[this.startPos + i + 7];
					let oPathAccumulator = new AscFormat.PathAccumulator();
					ArcToCurvers(oPathAccumulator, stX, stY, wR, hR, stAng, swAng);
					let aArcToCommands = oPathAccumulator.pathCommand;
					for (let nArcCmd = 0; nArcCmd < aArcToCommands.length; ++nArcCmd) {
						let oArcToCmd = aArcToCommands[nArcCmd];
						switch (oArcToCmd.id) {
							case AscFormat.moveTo: {
								break;
							}
							case AscFormat.bezier4: {
								dX0 = oTransform.TransformPointX(oArcToCmd.X0, oArcToCmd.Y0) * 36000 >> 0;
								dY0 = oTransform.TransformPointY(oArcToCmd.X0, oArcToCmd.Y0) * 36000 >> 0;
								dX1 = oTransform.TransformPointX(oArcToCmd.X1, oArcToCmd.Y1) * 36000 >> 0;
								dY1 = oTransform.TransformPointY(oArcToCmd.X1, oArcToCmd.Y1) * 36000 >> 0;
								dX2 = oTransform.TransformPointX(oArcToCmd.X2, oArcToCmd.Y2) * 36000 >> 0;
								dY2 = oTransform.TransformPointY(oArcToCmd.X2, oArcToCmd.Y2) * 36000 >> 0;
								oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);

								dLastX = oArcToCmd.X2;
								dLastY = oArcToCmd.Y2;
								break;
							}
						}
					}
					i += 7;
					break;
				}
				case close: {

					if (AscFormat.isRealNumber(oLastMoveTo)) {
						X = path[this.startPos + oLastMoveTo + 2];
						Y = path[this.startPos + oLastMoveTo + 3];
						let dXM = oTransform.TransformPointX(X, Y);
						let dYM = oTransform.TransformPointY(X, Y);
						let dLastXM = oTransform.TransformPointX(dLastX, dLastY);
						let dLastYM = oTransform.TransformPointY(dLastX, dLastY);
						dX0 = (dLastXM + (dXM - dLastXM) / 4) * 36000 >> 0;
						dY0 = (dLastYM + (dYM - dLastYM) / 4) * 36000 >> 0;
						dX1 = (dLastXM + 3 * (dXM - dLastXM) / 4) * 36000 >> 0;
						dY1 = (dLastYM + 3 * (dYM - dLastYM) / 4) * 36000 >> 0;
						dX2 = (dXM) * 36000 >> 0;
						dY2 = (dYM) * 36000 >> 0;
						oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);
					}
					oPath.close();
					i += 1;
					break;
				}
			}
		}
		oPath.recalculate({}, true);
	};

	Path2.prototype.Write_ToBinary = function(writer) {
		AscFormat.writeBool(writer, this.extrusionOk);
		AscFormat.writeString(writer, this.fill);
		AscFormat.writeLong(writer, this.pathH);
		AscFormat.writeLong(writer, this.pathW);
		AscFormat.writeLong(writer, this.startPos);
		AscFormat.writeBool(writer, this.stroke);
	};
	Path2.prototype.Read_FromBinary = function(reader) {
		this.extrusionOk = AscFormat.readBool(reader);
		this.fill = AscFormat.readString(reader);
		this.pathH = AscFormat.readLong(reader);
		this.pathW = AscFormat.readLong(reader);
		this.startPos = AscFormat.readLong(reader);
		this.stroke = AscFormat.readBool(reader);
		this.PathMemory = reader.pathMemory;
	};

	function partition_bezier3(x0, y0, x1, y1, x2, y2, epsilon) {
		let dx01 = x1 - x0;
		let dy01 = y1 - y0;
		let dx12 = x2 - x1;
		let dy12 = y2 - y1;

		let r01 = Math.sqrt(dx01 * dx01 + dy01 * dy01);
		let r12 = Math.sqrt(dx12 * dx12 + dy12 * dy12);
		if (Math.max(r01, r12) < epsilon) {
			return [{x: x0, y: y0}, {x: x1, y: y1}, {x: x2, y: y2}];
		}

		let x01 = (x0 + x1) * 0.5;
		let y01 = (y0 + y1) * 0.5;

		let x12 = (x1 + x2) * 0.5;
		let y12 = (y1 + y2) * 0.5;

		let x012 = (x01 + x12) * 0.5;
		let y012 = (y01 + y12) * 0.5;

		return partition_bezier3(x0, y0, x01, y01, x012, y012, epsilon).concat(partition_bezier3(x012, y012, x12, y12, x2, y2, epsilon));
	}

	function partition_bezier4(x0, y0, x1, y1, x2, y2, x3, y3, epsilon) {
		let dx01 = x1 - x0;
		let dy01 = y1 - y0;
		let dx12 = x2 - x1;
		let dy12 = y2 - y1;
		let dx23 = x3 - x2;
		let dy23 = y3 - y2;

		let r01 = Math.sqrt(dx01 * dx01 + dy01 * dy01);
		let r12 = Math.sqrt(dx12 * dx12 + dy12 * dy12);
		let r23 = Math.sqrt(dx23 * dx23 + dy23 * dy23);

		if (Math.max(r01, r12, r23) < epsilon) return [{x: x0, y: y0}, {x: x1, y: y1}, {x: x2, y: y2}, {x: x3, y: y3}];


		let x01 = (x0 + x1) * 0.5;
		let y01 = (y0 + y1) * 0.5;

		let x12 = (x1 + x2) * 0.5;
		let y12 = (y1 + y2) * 0.5;

		let x23 = (x2 + x3) * 0.5;
		let y23 = (y2 + y3) * 0.5;

		let x012 = (x01 + x12) * 0.5;
		let y012 = (y01 + y12) * 0.5;

		let x123 = (x12 + x23) * 0.5;
		let y123 = (y12 + y23) * 0.5;

		let x0123 = (x012 + x123) * 0.5;
		let y0123 = (y012 + y123) * 0.5;

		return partition_bezier4(x0, y0, x01, y01, x012, y012, x0123, y0123, epsilon).concat(partition_bezier4(x0123, y0123, x123, y123, x23, y23, x3, y3, epsilon));
	}

	function splitBezier4(x0, y0, x1, y1, x2, y2, x3, y3, parameters) {
		const aResult = [[x0, y0, x1, y1, x2, y2, x3, y3]];
		if (!Array.isArray(parameters) || parameters.length === 0) {
			return aResult;
		}
		const aWorkingParameters = [].concat(parameters);
		aWorkingParameters.sort(function (a, b) {
			return a - b
		});
		const isN = AscFormat.isRealNumber;
		const isE = AscFormat.fApproxEqual;
		let dLastParam = 1.0;
		for (let nParamIdx = aWorkingParameters.length - 1; nParamIdx > -1; --nParamIdx) {
			let dParam = aWorkingParameters[nParamIdx];
			if (!isN(dParam) || dParam <= 0 || dParam >= 1.0 || isE(dParam, dLastParam)) {
				aWorkingParameters.splice(nParamIdx, 1);
			}
			else {
				dLastParam = dParam;
			}
		}

		dLastParam = 1.0;
		for (let nParamIdx = aWorkingParameters.length - 1; nParamIdx > -1; --nParamIdx) {
			let dParam = aWorkingParameters[nParamIdx];
			let dWorkingParam = dParam / dLastParam;
			let oCurrentCurve = aResult[0];

			let x0c = oCurrentCurve[0];
			let y0c = oCurrentCurve[1];
			let x1c = oCurrentCurve[2];
			let y1c = oCurrentCurve[3];
			let x2c = oCurrentCurve[4];
			let y2c = oCurrentCurve[5];
			let x3c = oCurrentCurve[6];
			let y3c = oCurrentCurve[7];
			let t = dWorkingParam;

			//De Casteljau's algorithm
			let x01 = x0c + (x1c - x0c) * t;
			let y01 = y0c + (y1c - y0c) * t;
			let x12 = x1c + (x2c - x1c) * t;
			let y12 = y1c + (y2c - y1c) * t;
			let x23 = x2c + (x3c - x2c) * t;
			let y23 = y2c + (y3c - y2c) * t;
			let x0112 = x01 + (x12 - x01) * t;
			let y0112 = y01 + (y12 - y01) * t;
			let x1223 = x12 + (x23 - x12) * t;
			let y1223 = y12 + (y23 - y12) * t;
			let x01121223 = x0112 + (x1223 - x0112) * t;
			let y01121223 = y0112 + (y1223 - y0112) * t;
			let oCurve1 = [x0c, y0c, x01, y01, x0112, y0112, x01121223, y01121223];
			let oCurve2 = [x01121223, y01121223, x1223, y1223, x23, y23, x3c, y3c];
			aResult.splice(0, 1, oCurve1, oCurve2);
			dLastParam = dParam;
		}
		return aResult;
	}

	function splitBezier4OnParts(x0, y0, x1, y1, x2, y2, x3, y3, nPartsCount) {
		if (!AscFormat.isRealNumber(nPartsCount) || nPartsCount < 2) {
			return splitBezier4(x0, y0, x1, y1, x2, y2, x3, y3, []);
		}
		let aParameters = [];
		const dDist = 1 / nPartsCount;
		for (let nPartIdx = 1; nPartIdx < nPartsCount; ++nPartIdx) {
			aParameters.push(nPartIdx * dDist);
		}
		return splitBezier4(x0, y0, x1, y1, x2, y2, x3, y3, aParameters);
	}


	function getEllipsePoint(dXCE, dYCE, dA, dB, dAlpha) {
		return {
			x: dXCE + dA * Math.cos(dAlpha), y: dYCE + dB * Math.sin(dAlpha)
		};
	}

	function getLineLength(oPt1, oPt2) {
		return Math.sqrt(Math.pow(oPt2.x - oPt1.x, 2) + Math.pow(oPt2.y - oPt1.y, 2));
	}

	function isInsideEllipse(dXCE, dYCE, dA, dB, p) {
		let dX = p.x;
		let dY = p.y;
		let dXVal = dX - dXCE;
		let dYVal = dY - dYCE;
		return (dXVal * dXVal) / (dA * dA) + (dYVal * dYVal) / (dB * dB) < 1;
	}

	function ellipseCircleIntersection(dXCE, dYCE, dA, dB, dStartAngle, dR) {
		const dTolerance = 0.001;
		let dAlpha1 = dStartAngle;
		let dAlpha2 = dStartAngle + Math.PI / 4;
		let pStart = getEllipsePoint(dXCE, dYCE, dA, dB, dStartAngle);
		let p1 = getEllipsePoint(dXCE, dYCE, dA, dB, dAlpha1);
		let p2 = getEllipsePoint(dXCE, dYCE, dA, dB, dAlpha2);

		let dAlphaMiddle = (dAlpha1 + dAlpha2) / 2.0;
		while (Math.abs(dAlpha1 - dAlpha2) > dTolerance) {
			dAlphaMiddle = (dAlpha1 + dAlpha2) / 2.0;
			let pMiddle = getEllipsePoint(dXCE, dYCE, dA, dB, dAlphaMiddle);
			if (isInsideEllipse(pStart.x, pStart.y, dR, dR, pMiddle)) {
				p1 = pMiddle;
				dAlpha1 = dAlphaMiddle;
			}
			else {
				p2 = pMiddle;
				dAlpha2 = dAlphaMiddle;
			}
		}
		return {p: p1, alpha: dAlphaMiddle};
	}


	function circlesIntersection(x1, y1, r1, x2, y2, r2) {
		let dDx = x1 - x2;
		let dDy = y1 - y2;
		let R = Math.sqrt(dDx * dDx + dDy * dDy);
		if (!(Math.abs(r1 - r2) <= R && R <= r1 + r2)) {
			return [];
		}

		let R2 = R * R;
		let R4 = R2 * R2;
		let a = (r1 * r1 - r2 * r2) / (2 * R2);
		let r2r2 = (r1 * r1 - r2 * r2);
		let c = Math.sqrt(2 * (r1 * r1 + r2 * r2) / R2 - (r2r2 * r2r2) / R4 - 1);

		let fx = (x1 + x2) / 2 + a * (x2 - x1);
		let gx = c * (y2 - y1) / 2;
		let ix1 = fx + gx;
		let ix2 = fx - gx;

		let fy = (y1 + y2) / 2 + a * (y2 - y1);
		let gy = c * (x1 - x2) / 2;
		let iy1 = fy + gy;
		let iy2 = fy - gy;

		return [{x: ix1, y: iy1}, {x: ix2, y: iy2}];
	}


	//--------------------------------------------------------export----------------------------------------------------
	window['AscFormat'] = window['AscFormat'] || {};
	window['AscFormat'].moveTo = moveTo;
	window['AscFormat'].lineTo = lineTo;
	window['AscFormat'].arcTo = arcTo;
	window['AscFormat'].bezier3 = bezier3;
	window['AscFormat'].bezier4 = bezier4;
	window['AscFormat'].close = close;
	window['AscFormat'].cToRad2 = cToRad2;
	window['AscFormat'].degToC = degToC;
	window['AscFormat'].radToDeg = radToDeg;
	window['AscFormat'].Path = Path;
	window['AscFormat'].Path2 = Path2;
	window['AscFormat'].CPathCmd = CPathCmd;
	window['AscFormat'].partition_bezier3 = partition_bezier3;
	window['AscFormat'].partition_bezier4 = partition_bezier4;
	window['AscFormat'].splitBezier4 = splitBezier4;
	window['AscFormat'].splitBezier4OnParts = splitBezier4OnParts;
	window['AscFormat'].ellipseCircleIntersection = ellipseCircleIntersection;
	window['AscFormat'].getLineLength = getLineLength;
	window['AscFormat'].getEllipsePoint = getEllipsePoint;
	window['AscFormat'].circlesIntersection = circlesIntersection;
})(window);

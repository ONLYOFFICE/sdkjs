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

(function (window) {
	// const IS_DEBUG_DRAWING = true;
	// const IS_ADD_HTML = false;
	AscCommon.IS_GENERATE_SMARTART_AND_TEXT_ON_OPEN = false;

	const LayoutNode = AscFormat.LayoutNode;
	const Choose = AscFormat.Choose;
	const If = AscFormat.If;
	const Else = AscFormat.Else;
	const Alg = AscFormat.Alg;
	const ForEach = AscFormat.ForEach;
	const Point = AscFormat.Point;
	const ConstrLst = AscFormat.ConstrLst;
	const SShape = AscFormat.SShape;
	const PresOf = AscFormat.PresOf;
	const RuleLst = AscFormat.RuleLst;
	const VarLst = AscFormat.VarLst;
	const LayoutBaseClass = AscFormat.LayoutBaseClass;
	const IteratorLayoutBase = AscFormat.IteratorLayoutBase;

	const degToRad = Math.PI / 180;
	const radToDeg = 1 / degToRad;
	const algDelta = 1e-10;
	const intDelta = 1e-2;

	const factRuleState = {
		default: 0,
		enabled: 1,
		disabled: 2
	};
	Math.imul =
		Math.imul ||
		function (a, b) {
			var ah = (a >>> 16) & 0xffff;
			var al = a & 0xffff;
			var bh = (b >>> 16) & 0xffff;
			var bl = b & 0xffff;
			return (al * bl + (((ah * bl + al * bh) << 16) >>> 0)) | 0;
		};
	function cyrb128(str) {
		let h1 = 1779033703, h2 = 3144134277,
			h3 = 1013904242, h4 = 2773480762;
		for (let i = 0, k; i < str.length; i++) {
			k = str.charCodeAt(i);
			h1 = h2 ^ Math.imul(h1 ^ k, 597399067);
			h2 = h3 ^ Math.imul(h2 ^ k, 2869860233);
			h3 = h4 ^ Math.imul(h3 ^ k, 951274213);
			h4 = h1 ^ Math.imul(h4 ^ k, 2716044179);
		}
		h1 = Math.imul(h3 ^ (h1 >>> 18), 597399067);
		h2 = Math.imul(h4 ^ (h2 >>> 22), 2869860233);
		h3 = Math.imul(h1 ^ (h3 >>> 17), 951274213);
		h4 = Math.imul(h2 ^ (h4 >>> 19), 2716044179);
		h1 ^= (h2 ^ h3 ^ h4), h2 ^= h1, h3 ^= h1, h4 ^= h1;
		return [h1>>>0, h2>>>0, h3>>>0, h4>>>0];
	}
	function sfc32(a, b, c, d) {
		return function() {
			a |= 0; b |= 0; c |= 0; d |= 0;
			let t = (a + b | 0) + d | 0;
			d = d + 1 | 0;
			a = b ^ b >>> 9;
			b = c + (c << 3) | 0;
			c = (c << 21 | c >>> 11);
			c = c + t | 0;
			return (t >>> 0) / 4294967296;
		}
	}
	function CreateTemplateGUID(randomFunc) {
		function s4() {
			return Math.floor((1 + randomFunc()) * 0x10000).toString(16).substring(1);
		}

		let val = '{' + s4() + s4() + '-' + s4() + '-' + s4() + '-' + s4() + '-' + s4() + s4() + s4() + '}';
		val = val.toUpperCase();
		return val;
	}
	function GUIDGenerator(str) {
		this.randomFunc = sfc32.apply(sfc32, cyrb128(str));
	}
	GUIDGenerator.prototype.getGUID = function() {
		return CreateTemplateGUID(this.randomFunc);
	};
	function getDefaultColorsStyleLbl() {
		return AscFormat.ExecuteNoHistory(function () {
			const styleLbl = new AscFormat.ColorDefStyleLbl();
			styleLbl.setName("noName");
			styleLbl.setFillClrLst(new AscFormat.ClrLst());
			styleLbl.fillClrLst.addToLst(0, AscFormat.CreateUniColorRGB(0, 0, 0));
			styleLbl.setLinClrLst(new AscFormat.ClrLst());
			styleLbl.linClrLst.addToLst(0, AscFormat.CreateUniColorRGB(0, 0, 0));
			return styleLbl;
		});
	}
	function isClockwisePoints(centerPoint, edgePoint1, edgePoint2) {
		const rectSum = (centerPoint.x - edgePoint1.x) * (edgePoint2.y - edgePoint1.y) -
			(edgePoint2.x - edgePoint1.x) * (centerPoint.y - edgePoint1.y);
		return rectSum < 0;
	}

	function checkBounds(firstBounds, secondBounds) {
		if (secondBounds.l < firstBounds.l) {
			firstBounds.l = secondBounds.l;
		}
		if (secondBounds.t < firstBounds.t) {
			firstBounds.t = secondBounds.t;
		}
		if (secondBounds.r > firstBounds.r) {
			firstBounds.r = secondBounds.r;
		}
		if (secondBounds.b > firstBounds.b) {
			firstBounds.b = secondBounds.b;
		}
	}

	function fAlgDeltaEqual(a, b) {
		return AscFormat.fApproxEqual(a, b, algDelta);
	}
	function fIntervalDeltaEqual(a, b) {
		return AscFormat.fApproxEqual(a, b, intDelta);
	}

	function CCoordPoint(x, y) {
		this.x = x;
		this.y = y;
	}

	LayoutBaseClass.prototype.executeLayoutAlgorithms = function(smartartAlgorithm) {
		if (this.shape) {
			this.shape.executeAlgorithm(smartartAlgorithm);
		}
		if (this.alg) {
			this.alg.executeAlgorithm(smartartAlgorithm);
		}
		for (let i = 0; i < this.list.length; i++) {
			this.list[i].executeAlgorithm(smartartAlgorithm);
		}
		if (this.presOf) {
			this.presOf.executeAlgorithm(smartartAlgorithm);
		}
		if (this.constrLst) {
			this.constrLst.executeAlgorithm(smartartAlgorithm);
		}
		if (this.ruleLst) {
			this.ruleLst.executeAlgorithm(smartartAlgorithm);
		}
		if (this.varLst) {
			this.varLst.executeAlgorithm(smartartAlgorithm);
		}
	};
	IteratorLayoutBase.prototype.executeLayoutAlgorithms = LayoutBaseClass.prototype.executeLayoutAlgorithms;
	CCoordPoint.prototype.getVector = function (point) {
		return new CVector(point.x - this.x, point.y - this.y);
	}
	function CVector(x, y) {
		this.x = x;
		this.y = y;
	}
	CVector.getVectorByAngle = function (angle) {
		return new CVector(Math.cos(angle), Math.sin(angle));
	};
	CVector.prototype.getDistance = function () {
		return Math.sqrt(this.x * this.x + this.y * this.y);
	};
	CVector.prototype.getDiffVector = function (vector) {
		return new CVector(this.x - vector.x, this.y - vector.y);
	};
	CVector.prototype.multiply = function (value) {
		this.x *= value;
		this.y *= value;
	};
	CVector.prototype.getAngle = function () {
		const x = this.x;
		const y = this.y;
		const vectorLength = Math.sqrt(x * x + y * y);
		if (vectorLength !== 0) {
			const angle = Math.acos(x / vectorLength);
			if (y > 0) {
				return angle;
			}
			return AscFormat.normalizeRotate(-angle);
		}
		return null;
	};

	CVector.prototype.getAngleBetween = function (vector, isAntiClockwise) {
		let currentAngle = this.getAngle();
		let anotherAngle = vector.getAngle();

		if (isAntiClockwise) {
			if (currentAngle < anotherAngle) {
				currentAngle += Math.PI * 2;
			}
			return currentAngle - anotherAngle;
		} else {
			if (anotherAngle < currentAngle || fAlgDeltaEqual(anotherAngle, currentAngle)) {
				anotherAngle += Math.PI * 2;
			}
			return anotherAngle - currentAngle;
		}
	};

	VarLst.prototype.executeAlgorithm = function (smartartAlgorithm) {};

	PresOf.prototype.executeAlgorithm = function (smartartAlgorithm) {
		const currentPresNode = smartartAlgorithm.getCurrentPresNode();
		const nodes = this.getNodesArray(smartartAlgorithm);
		for (let i = 0; i < nodes.length; i++) {
			const node = nodes[i];
			currentPresNode.contentNodes.push(node);
			if (currentPresNode.algorithm instanceof TextAlgorithm) {
				node.presOf.textNode = currentPresNode;
			} else if (currentPresNode.layoutInfo.shape && !currentPresNode.layoutInfo.shape.hideGeom && currentPresNode.layoutInfo.shape.type !== AscFormat.LayoutShapeType_outputShapeType_none) {
				node.presOf.contentNode = currentPresNode;
			}
		}

		if (nodes.length) {
			currentPresNode.node = nodes[0] || smartartAlgorithm.getCurrentNode();
		}
	};
	LayoutNode.prototype.executeAlgorithm = function (smartartAlgorithm) {
		const parentPresNode = smartartAlgorithm.getCurrentPresNode();
		const curPresNode = smartartAlgorithm.getPresNode(this);
		parentPresNode.addChild(curPresNode);
		smartartAlgorithm.addCurrentPresNode(curPresNode);
		this.executeLayoutAlgorithms(smartartAlgorithm);
		curPresNode.checkMoveWith();
		curPresNode.initPresShape();
		smartartAlgorithm.removeCurrentPresNode(curPresNode);
	}
	LayoutNode.prototype.getForEachMap = function () {
		const forEachMap = {};
		const elements = [this];
		while (elements.length) {
			const element = elements.pop();

			const list = element.list;
			if (list) {
				elements.push.apply(elements, list);
			}
			if (element instanceof ForEach) {
				forEachMap[element.name] = element;
			}
		}
		return forEachMap;
	}

	Choose.prototype.executeAlgorithm = function (smartartAlgorithm) {
		for (let i = 0; i < this.if.length; i++) {
			if (this.if[i].executeAlgorithm(smartartAlgorithm)) {
				return;
			}
		}
		this.else.executeAlgorithm(smartartAlgorithm);
	};


	If.prototype.executeAlgorithm = function (smartartAlgorithm) {
		if (this.checkCondition(smartartAlgorithm)) {
			this.executeLayoutAlgorithms(smartartAlgorithm);
			return true;
		}
		return false;
	}

	If.prototype.checkCondition = function (smartArtAlgorithm) {
		const node = smartArtAlgorithm.getCurrentNode();
		const nodes = this.getNodesArray(smartArtAlgorithm);

		switch (this.func) {
			case AscFormat.If_func_cnt:
				return this.funcCnt(nodes);
			case AscFormat.If_func_depth:
				return this.funcDepth(node);
			case AscFormat.If_func_maxDepth:
				const depth = node.isParNode() ? node.depth - 1 : node.depth;
				return this.funcMaxDepth(nodes, depth);
			case AscFormat.If_func_pos:
				return this.funcPos(nodes, node);
			case AscFormat.If_func_posEven:
				return this.funcPosEven(nodes, node);
			case AscFormat.If_func_posOdd:
				return this.funcPosOdd(nodes, node);
			case AscFormat.If_func_revPos:
				return this.funcRevPos(nodes, node);
			case AscFormat.If_func_var:
				return this.funcVar(smartArtAlgorithm);
			default:
				return false;
		}
	};

	If.prototype.check = function (expected, result) {
		switch (this.op) {
			case AscFormat.If_op_equ: {
				return expected === result;
			}
			case AscFormat.If_op_gt: {
				return result > expected;
			}
			case AscFormat.If_op_lt: {
				return result < expected;
			}
			case AscFormat.If_op_gte: {
				return result >= expected;
			}
			case AscFormat.If_op_lte: {
				return result <= expected;
			}
			case AscFormat.If_op_neq: {
				return result !== expected;
			}
			default: {
				return false;
			}
		}
	}
	If.prototype.funcPosEven = function (nodes, currentNode) {
		const conditionValue = parseInt(this.getConditionValue(), 10);
		for (let i = 0; i < nodes.length; i += 1) {
			if (nodes === currentNode) {
				return this.check(conditionValue, (i + 1) % 2 === 0 ? 1 : 0);
			}
		}
		return false;
	};
	If.prototype.funcPosOdd = function (nodes, currentNode) {
		const conditionValue = parseInt(this.getConditionValue(), 10);
		const position = currentNode.getPositionByParent() + 1;
		return this.check(conditionValue, position % 2);
	};
	If.prototype.funcPos = function (nodes, currentNode) {
		const conditionValue = parseInt(this.getConditionValue(), 10);
		const position = currentNode.getPositionByParent() + 1;
		return this.check(conditionValue, position);
	};
	If.prototype.funcRevPos = function (nodes, currentNode) {
		const conditionValue = parseInt(this.getConditionValue(), 10);
		const position = currentNode.getPositionByParent();
		const parentLength = currentNode.getParentChildLength();
		return this.check(conditionValue, parentLength - position);
	};
	If.prototype.funcDepth = function (currentNode) {
		const conditionValue = parseInt(this.getConditionValue(), 10);
		return this.check(conditionValue, currentNode.depth);
	};
	If.prototype.funcMaxDepth = function (nodes, curDepth) {
		const conditionValue = parseInt(this.getConditionValue(), 10);
		let maxDepth = curDepth;
		for (let i = 0; i < nodes.length; i++) {
			const depth = nodes[i].getChildDepth();
			if (depth > maxDepth) {
				maxDepth = depth;
			}
		}
		return this.check(conditionValue, maxDepth - curDepth);
	};
	If.prototype.getVarPoint = function (smartArtAlgorithm) {
		let node;
		switch (this.arg) {
			case AscFormat.If_arg_dir: {
				node = smartArtAlgorithm.getRootPresNode();
				break;
			}
			default:
				node = smartArtAlgorithm.getCurrentPresNode();
				while (node && node.presPoint && !(node.presPoint.prSet && node.presPoint.prSet.presLayoutVars)) {
					if (!(node.parent && node.parent.presPoint)) {
						break;
					}
					node = node.parent;
				}
				break;
		}
		return node && node.presPoint;
	};
	If.prototype.funcVar = function (smartArtAlgorithm) {
		const point = this.getVarPoint(smartArtAlgorithm);
		const nodeVal = point && point.getFuncVarValue(this.arg);
		return this.check(this.getConditionValue(), nodeVal);
	}
	If.prototype.funcCnt = function (nodes) {
		return this.check(parseInt(this.val, 10), nodes.length);
	}
	If.prototype.getConditionValue = function () {
		switch (this.arg) {
			case AscFormat.If_arg_dir: {
				return this.getConditionDirValue();
			}
			case AscFormat.If_arg_hierBranch: {
				return this.getConditionHierBranchValue();
			}
			default:
				return this.val;
		}
	};
	If.prototype.getConditionHierBranchValue = function () {
		switch (this.val) {
			case 'l':
				return AscFormat.HierBranch_val_l;
			case 'r':
				return AscFormat.HierBranch_val_r;
			case 'hang':
				return AscFormat.HierBranch_val_hang;
			case 'init':
				return AscFormat.HierBranch_val_init;
			case 'std':
				return AscFormat.HierBranch_val_std;
			default:
				break;
		}
	}
	If.prototype.getConditionDirValue = function () {
		switch (this.val) {
			case 'norm':
				return AscFormat.DiagramDirection_val_norm;
			case 'rev':
				return AscFormat.DiagramDirection_val_rev;
			default:
				break;
		}
	}
	ConstrLst.prototype.executeAlgorithm = function (smartartAlgorithm) {
		smartartAlgorithm.setConstraints(this.list);
	}

	RuleLst.prototype.executeAlgorithm = function (smartartAlgorithm) {
		smartartAlgorithm.setRules(this.list);
	}

	Else.prototype.executeAlgorithm = function (smartartAlgorithm) {
		this.executeLayoutAlgorithms(smartartAlgorithm);
	}

	ForEach.prototype.executeAlgorithm = function (smartartAlgorithm) {
		const refForEach = smartartAlgorithm.getForEach(this.ref);
		if (refForEach) {
			refForEach.executeAlgorithm(smartartAlgorithm);
			return;
		}
		const currentNode = smartartAlgorithm.getCurrentNode();

		let isHideLastTrans2 = this.getHideLastTrans(0);
		if (!isHideLastTrans2) {
			const parent = currentNode.parent || currentNode;
			const lastChild = parent.childs[parent.childs.length - 1];
			if (lastChild && lastChild.sibNode) {
				lastChild.sibNode.isHideLastTrans = false;
			}
		}
		const nodes = this.getNodesArray(smartartAlgorithm);
		for (let i = 0; i < nodes.length; i += 1) {
			const node = nodes[i];
			smartartAlgorithm.addCurrentNode(node);
			this.executeLayoutAlgorithms(smartartAlgorithm);
			smartartAlgorithm.removeCurrentNode();
		}
	};

	Alg.prototype.getAlgorithm = function (smartartAlgorithm) {
		let algorithm;
		switch (this.getType()) {
			case AscFormat.Alg_type_snake: {
				algorithm = new SnakeAlgorithm();
				break;
			}
			case AscFormat.Alg_type_tx: {
				algorithm = new TextAlgorithm();
				break;
			}
			case AscFormat.Alg_type_sp: {
				algorithm = new SpaceAlgorithm();
				break;
			}
			case AscFormat.Alg_type_composite: {
				algorithm = new CompositeAlgorithm();
				break;
			}
			case AscFormat.Alg_type_lin: {
				algorithm = new LinearAlgorithm();
				break;
			}
			case AscFormat.Alg_type_conn: {
				algorithm = new ConnectorAlgorithm();
				break;
			}
			case AscFormat.Alg_type_cycle: {
				algorithm = new CycleAlgorithm();
				break;
			}
			case AscFormat.Alg_type_pyra: {
				algorithm = new PyramidAlgorithm();
				break;
			}
			case AscFormat.Alg_type_hierRoot: {
				algorithm = new HierarchyRootAlgorithm();
				break;
			}
			case AscFormat.Alg_type_hierChild: {
				algorithm = new HierarchyChildAlgorithm();
				break;
			}
			default: {
				break;
			}
		}
		if (algorithm) {
			algorithm.setParams(this.param);
			algorithm.setParentNode(smartartAlgorithm.getCurrentPresNode());
		}
		return algorithm;
	};

	Alg.prototype.executeAlgorithm = function (smartartAlgorithm) {
		const node = smartartAlgorithm.getCurrentPresNode();
		node.setAlgorithm(this.getAlgorithm(smartartAlgorithm));
	}
	Point.prototype.getFuncVarValue = function(arg) {
		const varLst = this.prSet && this.prSet.presLayoutVars;
			switch (arg) {
				case AscFormat.If_arg_dir:
					if (varLst && varLst.dir !== null) {
						return varLst.dir;
					}
					return AscFormat.DiagramDirection_val_norm;
				case AscFormat.If_arg_hierBranch:
					if (varLst && varLst.hierBranch !== null) {
						return varLst.hierBranch;
					}
					return AscFormat.HierBranch_val_std;
				default:
					break;
			}
	};

	SShape.prototype.executeAlgorithm = function (smartartAlgoritm) {
		const presNode = smartartAlgoritm.getCurrentPresNode();
		presNode.layoutInfo.shape = this;
	}

	function SmartArtAlgorithm(smartart) {
		this.smartart = smartart;
		this.relations = this.smartart.getRelationOfContent2();
		this.dataRoot = null;
		this.presRoot = null;
		this.nodesStack = [];
		this.presNodesStack = [];
		this.presNodesInfoStack = [];
		this.connectorAlgorithmStack = [];
		this.moveShapeSettings = [];
		this.forEachMap = null;
		this.factRuleState = factRuleState.default;
		this.guidGenerator = new GUIDGenerator(this.getModelIdStr());
		this.initDataTree();
	}
	SmartArtAlgorithm.prototype.getModelIdStr = function() {
		const data = this.smartart.getDataModelFromData();
		if (data) {
			const ptLst = data.getPtLst();
			if (ptLst) {
				return ptLst.getModelIdStr();
			}
		}
		return "";
	}
	SmartArtAlgorithm.prototype.createPresNode = 	function (presName, styleLbl, contentNode) {
		presName = presName || "";
		const point = new Point();
		point.setType(AscFormat.Point_type_pres);
		const prSet = new AscFormat.PrSet();
		prSet.setPresName(presName);
		prSet.setPresStyleLbl(styleLbl || "node1");

		point.setPrSet(prSet);
		point.setModelId(this.getNewGUID());
		return new PresNode(point, contentNode);
	};
	SmartArtAlgorithm.prototype.forEachContentNode = function (callback) {
		const nodes = [this.dataRoot];
		callback(this.dataRoot, 0);
		while (nodes.length) {
			const node = nodes.pop();
			for (let i = 0; i < node.childs.length; i++) {
				const child = node.childs[i];
				callback(child, i);
				nodes.push(child);
			}
		}
	};
	SmartArtAlgorithm.prototype.getDataPointRelations = function () {
		const mapRelations = {};
		this.forEachContentNode(function (node) {
			mapRelations[node.point.getModelId()] = node;
			if (node.sibNode) {
				mapRelations[node.sibNode.point.getModelId()] = node.sibNode;
			}
			if (node.parNode) {
				mapRelations[node.parNode.point.getModelId()] = node.parNode;
			}
		});
		return mapRelations;
	};
	SmartArtAlgorithm.prototype.setFactRuleState = function (pr) {
		this.factRuleState = pr;
	};
	SmartArtAlgorithm.prototype.addSettingsForMove = function (settings) {
		this.moveShapeSettings.push(settings);
	};
	SmartArtAlgorithm.prototype.applySettingsForMove = function () {
		while (this.moveShapeSettings.length) {
			const settings = this.moveShapeSettings.pop();
			const coefficient = settings.coefficient;
			const shape = settings.shape;

			const presNode = shape.node;
			const prSet = presNode.getPrSet();
			let neighborWidth = null;
			let neighborHeight = null;
			const neighbor = presNode.getNeighbor();
			const neighborShape = neighbor.shape;
/*			if (neighborShape) {
				if (presNode.node.isSibNode()) {
					neighborHeight = neighborShape.cleanParams.height;
				} else {
					neighborWidth = neighborShape.cleanParams.width;
				}
			}*/
			if (prSet) {
				let offX = 0;
				let offY = 0;
				const radialInfo = shape.getCustomRadialInfo();
				if (radialInfo) {
					const custScaleRadius = prSet.custRadScaleRad === null ? 1 : prSet.custRadScaleRad;
					const custScaleAngle = prSet.custRadScaleInc === null ? 0 : prSet.custRadScaleInc;
					if (custScaleRadius !== 1 || custScaleAngle !== 0) {
						const shapeCenterPoint = new CCoordPoint(shape.x + shape.width / 2, shape.y + shape.height / 2);
						const centerPoint = new CCoordPoint(shapeCenterPoint.x - shape.radialVector.x, shapeCenterPoint.y - shape.radialVector.y);
						const customRadius = radialInfo.radius;
						const custAngle = radialInfo.angle;
						const custVector = CVector.getVectorByAngle(custAngle);
						custVector.multiply(customRadius);
						const custCenterPoint = new CCoordPoint(custVector.x + centerPoint.x, custVector.y + centerPoint.y);
						offX = custCenterPoint.x - shape.width / 2 - shape.x;
						offY = custCenterPoint.y - shape.height / 2 - shape.y;
						shape.custCenterPoint = centerPoint;
					}
				}

				if (prSet.custLinFactNeighborX) {
					const width = neighborWidth !== null ? neighborWidth : shape.cleanParams.width;
					offX += width * prSet.custLinFactNeighborX * coefficient;
				}
				if (prSet.custLinFactX) {
					offX += shape.cleanParams.width * prSet.custLinFactX * coefficient;
				}
				if (prSet.custLinFactNeighborY) {
					const height = neighborHeight !== null ? neighborHeight : shape.cleanParams.height;
					offY += height * prSet.custLinFactNeighborY * coefficient;
				}
				if (prSet.custLinFactY) {
					offY += shape.cleanParams.height * prSet.custLinFactY * coefficient;
				}

				const smartart = this.smartart;
				const smartartHeight = smartart.spPr.xfrm.extY;
				const smartartWidth = smartart.spPr.xfrm.extX;
				const shapeBounds = shape.getBounds();
				if (shapeBounds.l + offX < 0) {
					offX = -shapeBounds.l;
				} else if (shapeBounds.r + offX > smartartWidth) {
					offX = smartartWidth - shapeBounds.r;
				}
				if (shapeBounds.t + offY < 0) {
					offY = -shapeBounds.t;
				} else if (shapeBounds.b + offY > smartartHeight) {
					offY = smartartHeight - shapeBounds.b;
				}
				if (shape.radialVector) {
					shape.radialVector = new CVector(shape.radialVector.x + offX, shape.radialVector.y + offY);
				}
				shape.moveTo(offX, offY);
			}
		}
	};
	SmartArtAlgorithm.prototype.getForEachMap = function() {
		if (this.forEachMap === null) {
			const layout = this.smartart.getLayoutDef();
			const layoutNode = layout.getLayoutNode();
			this.forEachMap = layoutNode.getForEachMap();
		}
		return this.forEachMap;
	};
	SmartArtAlgorithm.prototype.getForEach = function (ref) {
		const forEachMap = this.getForEachMap();
		return forEachMap[ref];
	};
	SmartArtAlgorithm.prototype.addConnectorAlgorithm = function (algorithm) {
		this.connectorAlgorithmStack.push(algorithm);
	}
	SmartArtAlgorithm.prototype.getParentObjects = function () {
		const smartart = this.smartart;
		return smartart.getParentObjects();
	};
	SmartArtAlgorithm.prototype.getInitObjects = function () {
		const api = Asc.editor;
		if (!api) {
			return null;
		}
		const smartart = this.smartart;
		const initObjects = Object.create(this.getParentObjects());
		initObjects.worksheet = null;
		initObjects.parent = null;
		initObjects.drawingObjects = null;
		initObjects.drawingDocument = api.getDrawingDocument();
		const editorId = api.getEditorId();
		if (editorId === AscCommon.c_oEditorId.Spreadsheet) {
			initObjects.worksheet = smartart.worksheet;
			initObjects.drawingObjects = smartart.drawingObjects;
		} else if (editorId === AscCommon.c_oEditorId.Presentation) {
			initObjects.parent = smartart.parent;
		}
		return initObjects;
	};
	SmartArtAlgorithm.prototype.applyColorsDef = function (shadowShapes) {
		const colorsDef = this.smartart.getColorsDef();
		const stylesDef = this.smartart.getStyleDef();
		const styleLblsByName = stylesDef.styleLbl;
		const colorLblsByName = colorsDef.styleLbl;
		const parentObjects = this.getParentObjects();
		const shapesByStyleLbl = {};
		for (let i = 0; i < shadowShapes.length; i += 1) {
			const shadowShape = shadowShapes[i];
			const node = shadowShape.node;
			const styleLbl = node.getPresStyleLbl();
			if (styleLbl) {
				if (!shapesByStyleLbl[styleLbl]) {
					shapesByStyleLbl[styleLbl] = [];
				}
				shapesByStyleLbl[styleLbl].push(shadowShape.connectorShape || shadowShape);
			}
		}
		for (let styleLbl in shapesByStyleLbl) {
			const colorLbl = colorLblsByName[styleLbl] || getDefaultColorsStyleLbl();
			const shapes = shapesByStyleLbl[styleLbl];
			if (colorLbl) {
					colorLbl.setShapeFill(shapes, parentObjects);
					colorLbl.setShapeLn(shapes, parentObjects);
			}
			const quickStyleLbl = styleLblsByName[styleLbl];
			if (quickStyleLbl) {
				quickStyleLbl.setShapeStyle(shapes);
			}
		}
	};

	SmartArtAlgorithm.prototype.initDataTree = function () {
		const dataModel = this.smartart.getDataModel().getDataModel();
		const mainSmartArtPoint = dataModel.getMainPoint();
		const treeRoot = new SmartArtDataNode(mainSmartArtPoint, 0);
		this.dataRoot = treeRoot;
		const elements = [treeRoot];

		while (elements.length) {
			const root = elements.shift();
			const rootChildDepth = root.depth + 1;
			let connectionChildren = this.relations[[AscFormat.Cxn_type_parOf]][root.getModelId()];
			if (connectionChildren) {
				for (let i = 0; i < connectionChildren.length; i += 1) {
					const connectionChild = connectionChildren[i];
					const contentPoint = connectionChild.point;
					const sibPoint = connectionChild.sibPoint;
					const parPoint = connectionChild.parPoint;
					const node = new SmartArtDataNode(contentPoint, rootChildDepth);
					node.setSibNode(new SmartArtSibDataNode(sibPoint, rootChildDepth));
					node.setParNode(new SmartArtParDataNode(parPoint, rootChildDepth));
					root.addChild(node);
					elements.push(node);
				}
				if (connectionChildren.length) {
					const lastNode = elements[elements.length - 1];
					const sibTrans = lastNode.sibNode;
					sibTrans.isHideLastTrans = true;
				}
			}
		}
	}
	SmartArtAlgorithm.prototype.updateData = function() {
		const data = this.smartart.getDataModelFromData();
		const cxnLst = data.getCxnLst();
		const ptLst = data.getPtLst();
		const cxnMap = cxnLst.getCxnMap();
		const ptMap = ptLst.getPtMap();
		const dataRootId = this.dataRoot.point.getModelId();
		delete ptMap[dataRootId];
		this.checkPresOfCxn(this.presRoot, dataRootId, cxnMap, cxnLst);
		this.updatePresParOfData(cxnMap, ptMap);
		this.updatePresOfData(cxnMap);
		this.updateParOfData(cxnMap, ptMap);
		cxnLst.removeChildrenFromMap(cxnMap);
		ptLst.removeChildrenFromMap(ptMap);
		this.relations = this.smartart.getRelationOfContent2();
	};
	SmartArtAlgorithm.prototype.getNewGUID = function() {
		return this.guidGenerator.getGUID();
	};
	SmartArtAlgorithm.prototype.updateParOfData = function(cxnMap, ptMap) {
		const data = this.smartart.getDataModelFromData();
		const ptLst = data.getPtLst();
		const oThis = this;
		this.forEachContentNode(function(node, parentIndex) {
			const parent = node.parent;
			if (parent) {
				const parentModelId = parent.point.getModelId();
				const childModelId = node.point.getModelId();
				if (ptMap[childModelId]) {
					delete ptMap[childModelId];
				} else {
					ptLst.addToLst(ptLst.list.length, node.point);
				}
				let cxn;
				if (cxnMap[parentModelId][childModelId]) {
					cxn = cxnMap[parentModelId][childModelId];
					delete cxnMap[parentModelId][childModelId];
					if (cxn.srcOrd !== parentIndex) {
						cxn.setSrcOrd(parentIndex);
					}
				} else {
					cxn = new AscFormat.Cxn();
					cxn.setModelId(oThis.getNewGUID());
					cxn.setDestOrd(0);
					cxn.setSrcOrd(parentIndex);
					cxn.setSrcId(parentModelId);
					cxn.setDestId(childModelId);
					cxn.setType(AscFormat.Cxn_type_parOf);
				}
				const sibNode = node.sibNode;
				if (sibNode) {
					const sibModelId = sibNode.point.getModelId();
					if (ptMap[sibModelId]) {
						delete ptMap[sibModelId];
					} else {
						ptLst.addToLst(ptLst.list.length, sibNode.point);
					}
					if (cxn.sibTransId !== sibModelId) {
						cxn.setSibTransId(sibModelId);
					}
				}
				const parNode = node.parNode;
				if (parNode) {
					const parModelId = parNode.point.getModelId();
					if (ptMap[parModelId]) {
						delete ptMap[parModelId];
					} else {
						ptLst.addToLst(ptLst.list.length, parNode.point);
					}
					if (cxn.parTransId !== parModelId) {
						cxn.setParTransId(parModelId);
					}
				}
			}
		});
	};
	SmartArtAlgorithm.prototype.checkPresOfCxn = function(presOfNode, contentNodeId, cxnMap, cxnLst) {
		const presPoint = presOfNode.presPoint;
		const modelId = presPoint.getModelId();
		if (cxnMap[contentNodeId][modelId]) {
			delete cxnMap[contentNodeId][modelId];
		} else {
			const newCxn = new AscFormat.Cxn();
			newCxn.setModelId(this.getNewGUID());
			newCxn.setSrcId(contentNodeId);
			newCxn.setDestId(modelId);
			newCxn.setDestOrd(0);
			newCxn.setSrcOrd(0);
			newCxn.setType(AscFormat.Cxn_type_presOf);
			cxnLst.addToLst(cxnLst.list.length, newCxn);
		}
	};
	SmartArtAlgorithm.prototype.checkPresOfData = function(node, cxnMap, cxnLst) {
		if (!node) {
			return;
		}
		const nodeModelId = node.point.getModelId();
		const textNodes = node.getTextNodes();
		const textNode = textNodes.textNode;
		const contentNode = textNodes.contentNode;
		if (textNode) {
			this.checkPresOfCxn(textNode, nodeModelId, cxnMap, cxnLst);
		}
		if (contentNode) {
			this.checkPresOfCxn(contentNode, nodeModelId, cxnMap, cxnLst);
		}
	}
	SmartArtAlgorithm.prototype.updatePresOfData = function(cxnMap) {
		const data = this.smartart.getDataModelFromData();
		const cxnLst = data.getCxnLst();
		const oThis = this;
		this.forEachContentNode(function(node) {
			oThis.checkPresOfData(node, cxnMap, cxnLst);
			oThis.checkPresOfData(node.sibNode, cxnMap, cxnLst);
			oThis.checkPresOfData(node.parNode, cxnMap, cxnLst);
		});
	};

	SmartArtAlgorithm.prototype.updatePresParOfData = function(cxnMap, ptMap) {
		const data = this.smartart.getDataModelFromData();
		const cxnLst = data.getCxnLst();
		const ptLst = data.getPtLst();
		const oThis = this;
		this.forEachPresFromTop(function(node, parentIndex) {
			const presPoint = node.presPoint;
			const modelId = presPoint.getModelId();
			if (ptMap[modelId] && presPoint === ptMap[modelId]) {
				delete ptMap[modelId];
			} else {
				 ptLst.addToLst(ptLst.list.length, presPoint);
			}
			const parentPresPoint = node.parent && node.parent.presPoint;
			if (!parentPresPoint) {
				return;
			}

			const parentModelId = parentPresPoint.getModelId();
			const cxn = cxnMap[parentModelId] && cxnMap[parentModelId][modelId];
			if (cxn) {
				if (cxn.srcOrd !== parentIndex) {
					cxn.setSrcOrd(parentIndex);
				}
				delete cxnMap[parentModelId][modelId];
			} else {
				const newCxn = new AscFormat.Cxn();
				newCxn.setModelId(oThis.getNewGUID());
				newCxn.setSrcId(parentModelId);
				newCxn.setDestId(modelId);
				newCxn.setType(AscFormat.Cxn_type_presParOf);
				newCxn.setDestOrd(0);
				newCxn.setSrcOrd(parentIndex);
				cxnLst.addToLst(cxnLst.list.length, newCxn);
			}
		});
	};
	SmartArtAlgorithm.prototype.getPresNode = function (layoutNode) {
		const currentNode = this.getCurrentNode();
		const currentPresNode = this.getCurrentPresNode();
		const currentPresNodeInfo = this.getCurrentPresNodeInfo();
		const presRelations = this.relations[AscFormat.Cxn_type_presOf];
		const presParRelations = this.relations[AscFormat.Cxn_type_presParOf];
		let presNode;
		if (currentPresNode && currentPresNode.parent) {
			const children = presParRelations[currentPresNode.getModelId()];
			let child;
			if (children && children[layoutNode.name]) {
				if (currentPresNodeInfo.childrenIndexes[layoutNode.name] === undefined) {
					currentPresNodeInfo.childrenIndexes[layoutNode.name] = -1;
				}
				currentPresNodeInfo.childrenIndexes[layoutNode.name] += 1;
				const element = children[layoutNode.name][currentPresNodeInfo.childrenIndexes[layoutNode.name]];
				child = element && element.point;
			}
			if (child && child.getPresName() === layoutNode.name) {
				presNode = new PresNode(child, currentNode);
			} else {
				presNode = this.createPresNode(layoutNode.name, layoutNode.styleLbl, currentNode);
			}
		} else {
			const nodeModelId = currentNode.getModelId();
			const presPoints = presRelations[nodeModelId];
			const presPoint = presPoints && presPoints[layoutNode.name];
			if (presPoint) {
				presNode = new PresNode(presPoint, currentNode);
			} else {
				presNode = this.createPresNode(layoutNode.name, layoutNode.styleLbl, currentNode);
			}
		}
		presNode.moveWith = layoutNode.moveWith;
		return presNode;
	};

	SmartArtAlgorithm.prototype.addCurrentPresNode = function (presNode) {
		this.presNodesStack.push(presNode);
		this.presNodesInfoStack.push({childrenIndexes: {}});
	}
	SmartArtAlgorithm.prototype.removeCurrentPresNode = function () {
		this.presNodesStack.pop();
		this.presNodesInfoStack.pop();
	}
	SmartArtAlgorithm.prototype.getCurrentPresNode = function () {
		return this.presNodesStack[this.presNodesStack.length - 1];
	}
	SmartArtAlgorithm.prototype.getCurrentPresNodeInfo = function () {
		return this.presNodesInfoStack[this.presNodesStack.length - 1];
	}
	SmartArtAlgorithm.prototype.getRootPresNode = function () {
		return this.presNodesStack[1];
	}
	SmartArtAlgorithm.prototype.getShapes = function () {

		const algorithm = this.presRoot.algorithm;
		return algorithm ? algorithm.getShapes(this) : [];
	}
	SmartArtAlgorithm.prototype.connectShapeSmartArtInfo = function () {
		if (!this.presRoot) {
			this.startFromBegin();
		} else {
			this.cleanDrawingShapeInfo();
		}

		const spTree = this.smartart.drawing.spTree;
		const mapEditorShapes = {};
		for (let i = 0; i < spTree.length; i++) {
			const shape = spTree[i];
			mapEditorShapes[shape.getModelId()] = shape;
		}
		this.forEachPresFromBottom(function (presNode) {
			const point = presNode.presPoint;
			const editorShape = mapEditorShapes[point.getModelId()];
			if (editorShape) {
				const shadowShape = presNode.getShape();
				shadowShape.setShapeSmartArtInfo(editorShape);
				shadowShape.editorShape = editorShape;
				const textNodes = presNode.contentNodes[0] && presNode.contentNodes[0].getTextNodes();
				if (textNodes) {
					if (presNode === textNodes.textNode) {
						presNode.algorithm.applyFontRelations(editorShape);
					} else if (presNode === textNodes.contentNode && textNodes.textNode) {
						textNodes.textNode.algorithm.applyFontRelations(editorShape);
					}
				}
			}
		});
	};


	SmartArtAlgorithm.prototype.cleanDrawingShapeInfo = function () {
		const smartart = this.smartart;
		const drawing = smartart.drawing;
		for (let i = 0; i < drawing.spTree.length; i++) {
			const shape = drawing.spTree[i];
			shape.setShapeSmartArtInfo(null);
		}
	};
	SmartArtAlgorithm.prototype.checkDataModel = function() {
		this.executeLayoutAlgorithms();
		this.updateData();
	};
	SmartArtAlgorithm.prototype.cleanPresOf = function() {
		this.forEachContentNode(function(node) {
			node.cleanPresOf();
		});
	};
	SmartArtAlgorithm.prototype.executeLayoutAlgorithms = function() {
		this.cleanPresOf();
		this.addCurrentNode(this.dataRoot);
		const mockPresNode = new PresNode();
		this.addCurrentPresNode(mockPresNode);

		const layout = this.smartart.getLayoutDef();
		const layoutNode = layout.getLayoutNode();
		layoutNode.executeAlgorithm(this);
		this.presRoot = mockPresNode.childs[0];
		this.presRoot.parent = null;
		this.removeCurrentPresNode();
		this.removeCurrentNode();
	}
	SmartArtAlgorithm.prototype.startFromBegin = function () {
		this.cleanDrawingShapeInfo();
		this.executeLayoutAlgorithms();
		this.calculateShadowShapes();
	};
	SmartArtAlgorithm.prototype.calculateShadowShapes = function() {
		this.cleanCalcValues();
		this.calcConstraints();
		this.cleanRules();
		this.calcScaleCoefficients();
		if (this.factRuleState === factRuleState.disabled) {
			this.cleanCalcValues();
			this.calcConstraints();
			this.calcScaleCoefficients();
		}
		this.calcAdaptedConstraints();
		this.executeAlgorithms();
	};
	SmartArtAlgorithm.prototype.cleanCalcValues = function () {
		this.forEachPresFromTop(function (presNode) {
			presNode.cleanRules();
			presNode.cleanConstraints();
			presNode.cleanScales();
			presNode.cleanPosition();
		});
		this.presRoot.initRootConstraints(this.smartart, this);
	};


	SmartArtAlgorithm.prototype.calcScaleCoefficients = function () {
		const oThis = this;
		this.forEachPresFromBottom(function (presNode) {
			presNode.startAlgorithm(oThis, true);
		});
	};
	SmartArtAlgorithm.prototype.calcAdaptedConstraints = function () {
		const oThis = this;
		const isNeedSetRules = this.factRuleState !== factRuleState.disabled;
		this.forEachPresFromTop(function (presNode) {
			if (isNeedSetRules) {
				presNode.setRules(oThis);
			}
			presNode.setConstraints(true, oThis);
		});
	};
	SmartArtAlgorithm.prototype.calcConstraints = function () {
		const oThis = this;
		const isNeedSetRules = this.factRuleState !== factRuleState.disabled;
		this.forEachPresFromTop(function (presNode) {
			if (isNeedSetRules) {
				presNode.setRules(oThis);
			}
			presNode.setConstraints(false, oThis);
		});
	};
	SmartArtAlgorithm.prototype.cleanRules = function () {
		this.forEachPresFromTop(function (presNode) {
			presNode.cleanRules();
		});
	};
	SmartArtAlgorithm.prototype.executeAlgorithms = function () {
		const oThis = this;
		this.forEachPresFromBottom(function (presNode) {
			presNode.startAlgorithm(oThis);
		});
		this.applySettingsForMove();
		this.generateConnectors();
		this.applySettingsForMove();
	};
	SmartArtAlgorithm.prototype.generateConnectors = function () {
		while (this.connectorAlgorithmStack.length) {
			const connectorAlgorithm = this.connectorAlgorithmStack.pop();
			connectorAlgorithm.connectShapes(this);
		}
	};
	SmartArtAlgorithm.prototype.forEachPresFromBottom = function (callback) {
		const checkElements = [this.presRoot];
		while (checkElements.length) {
			const elem = checkElements.pop();
			if (elem.childs.length) {
				for (let i = 0; i < elem.childs.length; i += 1) {
					checkElements.push(elem.childs[i]);
				}
			} else {
				const callbackElements = [elem];
				while (callbackElements.length) {
					const elem = callbackElements.pop();
					callback(elem);
					const firstParentChild = elem.parent && elem.parent.childs[0];
					if (firstParentChild === elem) {
						callbackElements.push(elem.parent);
					}
				}
			}
		}
	};


	SmartArtAlgorithm.prototype.forEachPresFromTop = function (callback) {
		const elements = [this.presRoot];
		callback(this.presRoot, 0);
		while (elements.length) {
			const element = elements.pop();
			for (let i = 0; i < element.childs.length; i += 1) {
				const child = element.childs[i];
				callback(child, i);
				elements.push(child);
			}
		}
	};


	SmartArtAlgorithm.prototype.getCurrentNode = function () {
		return this.nodesStack[this.nodesStack.length - 1];
	}
	SmartArtAlgorithm.prototype.addCurrentNode = function (node) {
		this.nodesStack.push(node);
	}

	SmartArtAlgorithm.prototype.removeCurrentNode = function () {
		this.nodesStack.pop();
	}

	SmartArtAlgorithm.prototype.setConstraints = function (constr) {
		const node = this.getCurrentPresNode();
		node.setLayoutConstraints(constr);
	}
	SmartArtAlgorithm.prototype.setRules = function (rules) {
		const node = this.getCurrentPresNode();
		node.setLayoutRules(rules);
	}

	function SmartArtDataNodeBase(point, depth) {
		this.point = point;
		this.parent = null;
		this.childs = [];
		this.depth = AscFormat.isRealNumber(depth) ? depth : null;
		this.presOf = {
			contentNode: null,
			textNode: null
		};
	}
	SmartArtDataNodeBase.prototype.cleanPresOf = function() {
		this.presOf = {
			contentNode: null,
			textNode: null
		};
	};
	SmartArtDataNodeBase.prototype.getTextNode = function () {
		const textNodes = this.getTextNodes();
		return textNodes.textNode || textNodes.contentNode;
	};
	SmartArtDataNodeBase.prototype.getContentNode = function () {
		const textNodes = this.getTextNodes();
		return textNodes.contentNode || textNodes.textNode;
	};
	SmartArtDataNodeBase.prototype.getTextNodes = function () {
		return this.presOf;
	};
	SmartArtDataNodeBase.prototype.getPositionByParent = function () {
		return -1;
	};
	SmartArtDataNodeBase.prototype.getPresName = function () {
		const presNode = this.getPresNode();
		return presNode && presNode.getPresName();
	};
	SmartArtDataNodeBase.prototype.getPtType = function () {
		return this.point.type;
	}

	SmartArtDataNodeBase.prototype.getNodesByAxis = function (nodes, axis, ptType) {
		nodes = nodes || [];
		switch (axis) {
			case AscFormat.AxisType_value_root: {
				this.getNodesByRoot(nodes, ptType);
				break;
			}
			case AscFormat.AxisType_value_ch: {
				this.getNodesByCh(nodes, ptType);
				break;
			}
			case AscFormat.AxisType_value_par: {
				this.getNodesByParent(nodes, ptType);
				break;
			}
			case AscFormat.AxisType_value_self: {
				const needNode = this.getNodeByPtType(ptType);
				if (needNode) {
					nodes.push(needNode);
				}
				break;
			}
			case AscFormat.AxisType_value_follow: {
				this.getNodesByFollow(nodes, ptType);
				break;
			}
			case AscFormat.AxisType_value_followSib: {
				this.getNodesByFollowSib(nodes, ptType);
				break;
			}
			case AscFormat.AxisType_value_precedSib: {
				this.getNodesByPrecedSib(nodes, ptType);
				break;
			}
			case AscFormat.AxisType_value_des: {
				this.getNodesByDescendant(nodes, ptType);
				break;
			}
			case AscFormat.AxisType_value_desOrSelf: {
				this.getNodesByAxis(nodes, AscFormat.AxisType_value_self, ptType);
				this.getNodesByAxis(nodes, AscFormat.AxisType_value_des, ptType);
				break;
			}
			default: {
				break;
			}
		}
		return nodes;
	}
	SmartArtDataNodeBase.prototype.getParent = function () {
		return this.parent;
	}
	SmartArtDataNodeBase.prototype.getNodesByParent = function (nodes, ptType) {
		const parent = this.getParent();
		const needNode = parent && parent.getNodeByPtType(ptType);
		if (needNode) {
			nodes.push(needNode);
		}
	};
	SmartArtDataNodeBase.prototype.getNodesByRoot = function (nodes, ptType) {
		let curNode = this;
		while (curNode && !curNode.isRoot()) {
			curNode = curNode.parent;
		}
		if (curNode && curNode.isRoot()) {
			const needNode = curNode.getNodeByPtType(ptType);
			if (needNode) {
				nodes.push(needNode);
			}
		}
	}
	SmartArtDataNodeBase.prototype.getNodesByDescendant = function (nodes, ptType) {
		const elements = [].concat(this.childs);
		while (elements.length) {
			const child = elements.shift();
			const needNode = child.getNodeByPtType(ptType);
			if (needNode) {
				nodes.push(needNode);
			}
			for (let i = 0; i < child.childs.length; i++) {
				elements.push(child.childs[i]);
			}
		}
	};

	SmartArtDataNodeBase.prototype.forEachDes = function (callback) {
		const elements = [this];
		while (elements.length) {
			const element = elements.pop();
			for (let i = 0; i < element.childs.length; i++) {
				elements.push(element.childs[i]);
				callback(element.childs[i]);
			}
		}
	};

	SmartArtDataNodeBase.prototype.forEachDesOrSelf = function (callback) {
		callback(this);
		this.forEachDes(callback);
	};
	SmartArtDataNodeBase.prototype.getNodesByFollow = function (nodes, ptType) {
		const parent = this.getParent();
		if (parent) {
			let bAdd = false;
			for (let i = 0; i < parent.childs.length; i++) {
				const child = parent.childs[i];
				if (bAdd) {
					child.forEachDesOrSelf(function (node) {
						const needNode = node.getNodeByPtType(ptType);
						if (needNode) {
							nodes.push(needNode);
						}
					});
				} else if (child === this) {
					bAdd = true;
					if (ptType === AscFormat.ElementType_value_sibTrans) {
						child.forEachDesOrSelf(function (node) {
							const needNode = node.getNodeByPtType(ptType);
							if (needNode) {
								nodes.push(needNode);
							}
						});
					}
				} else if (child.sibNode === this) {
					bAdd = true;
				}
			}
		}
	}
	SmartArtDataNodeBase.prototype.getNodesByFollowSib = function (nodes, ptType) {
		const parent = this.getParent();
		if (parent) {
			let bAdd = false;
			for (let i = 0; i < parent.childs.length; i++) {
				const child = parent.childs[i];
				if (bAdd) {
					const needNode = parent.childs[i].getNodeByPtType(ptType);
					if (needNode) {
						nodes.push(needNode);
					}
				} else if (child === this) {
					bAdd = true;
					if (ptType === AscFormat.ElementType_value_sibTrans) {
						const needNode = parent.childs[i].getNodeByPtType(ptType);
						if (needNode) {
							nodes.push(needNode);
						}
					}
				} else if (child.sibNode === this) {
					bAdd = true;
				}
			}
		}
	};
	SmartArtDataNodeBase.prototype.getNodesByPrecedSib = function (nodes, ptType) {
		const parent = this.parent;
		if (parent) {
			for (let i = 0; i < parent.childs.length; i++) {
				if (parent.childs[i] === this) {
					if (ptType === AscFormat.ElementType_value_parTrans) {
						const needNode = parent.childs[i].getNodeByPtType(ptType);
						if (needNode) {
							nodes.push(needNode);
						}
					}
					break;
				} else {
					const needNode = parent.childs[i].getNodeByPtType(ptType);
					if (needNode) {
						nodes.push(needNode);
					}
				}
			}
		}
	};

	SmartArtDataNodeBase.prototype.getNodesByCh = function (nodes, ptType) {
		for (let i = 0; i < this.childs.length; i++) {
			const child = this.childs[i];
			const needNode = child.getNodeByPtType(ptType);
			if (needNode) {
				nodes.push(needNode);
			}
		}
	};

	SmartArtDataNodeBase.prototype.addChild = function (child, position) {
		position = AscFormat.isRealNumber(position) ? position : this.childs.length;
		this.childs.splice(position, 0, child);
		child.setParent(this);
	};
	SmartArtDataNodeBase.prototype.removeChilds = function (position, count) {
		this.childs.splice(position, count);
	};
	SmartArtDataNodeBase.prototype.setParent = function (parent) {
		this.parent = parent;
	};
	SmartArtDataNodeBase.prototype.getPresNode = function () {
		const textNodes = this.getTextNodes();
		return textNodes.textNode || textNodes.contentNode;
	};
	SmartArtDataNodeBase.prototype.getModelId = function () {
		return this.point.getModelId();
	};
	SmartArtDataNodeBase.prototype.isRoot = function () {
		return this.point.getType() === AscFormat.Point_type_doc;
	};
	SmartArtDataNodeBase.prototype.isNode = function () {
		return this.point.getType() === AscFormat.Point_type_node;
	};
	SmartArtDataNodeBase.prototype.isAsst = function () {
		return this.point.getType() === AscFormat.Point_type_asst;
	};
	SmartArtDataNodeBase.prototype.isContentNode = function () {
		return false;
	};
	SmartArtDataNodeBase.prototype.isSibNode = function () {
		return false;
	};
	SmartArtDataNodeBase.prototype.isParNode = function () {
		return false;
	};

	function SmartArtSibDataNode(mainPoint, depth) {
		SmartArtDataNodeBase.call(this, mainPoint, depth);
	}
	AscFormat.InitClassWithoutType(SmartArtSibDataNode, SmartArtDataNodeBase);
	SmartArtSibDataNode.prototype.isSibNode = function () {
		return true;
	};
	SmartArtSibDataNode.prototype.getParent = function () {
		return this.parent && this.parent.parent;
	}
	SmartArtSibDataNode.prototype.getNodeByPtType = function (elementTypeValue) {
		switch (elementTypeValue) {
			case AscFormat.ElementType_value_sibTrans:
				return this;
			case AscFormat.ElementType_value_node:
				return this;
			default:
				return this;
		}
	}

	function SmartArtParDataNode(mainPoint, depth) {
		SmartArtDataNodeBase.call(this, mainPoint, depth);
	}
	AscFormat.InitClassWithoutType(SmartArtParDataNode, SmartArtDataNodeBase);
	SmartArtParDataNode.prototype.isParNode = function () {
		return true;
	};
	SmartArtParDataNode.prototype.getNodeByPtType = function (elementTypeValue) {
		return this;
	}
	SmartArtParDataNode.prototype.getParent = function () {
		return this.parent && this.parent.parent;
	}
	function SmartArtDataNode(mainPoint, depth) {
		SmartArtDataNodeBase.call(this, mainPoint, depth);
		this.sibNode = null;
		this.parNode = null;
		this.childDepth = null;
	}
	AscFormat.InitClassWithoutType(SmartArtDataNode, SmartArtDataNodeBase);
	SmartArtDataNode.prototype.getPointType = function () {
		return this.point.getType();
	}
	SmartArtDataNode.prototype.getParentChildLength = function () {
		const parent = this.getParent();
		return parent.childs.length;
	}
	SmartArtDataNode.prototype.getPositionByParent = function () {
		const parent = this.getParent();
		for (let i = 0; i < parent.childs.length; i++) {
			if (parent.childs[i] === this) {
				return i;
			}
		}
		return -1;
	};

	SmartArtDataNode.prototype.getChildDepth = function () {
		if (this.childDepth === null) {
			let maxDepth = this.depth;
			const tempNodes = [this];
			while (tempNodes.length) {
				const node = tempNodes.pop();
				if (node.depth > maxDepth) {
					maxDepth = node.depth;
				}
				tempNodes.push.apply(tempNodes, node.childs);
			}
			this.childDepth = maxDepth;
		}
		return this.childDepth;
	};
	SmartArtDataNode.prototype.isContentNode = function () {
		return true;
	};

	SmartArtDataNode.prototype.setSibNode = function (node) {
		this.sibNode = node;
		node.setParent(this);
	};

	SmartArtDataNode.prototype.setParNode = function (node) {
		this.parNode = node;
		node.setParent(this);
	};

	SmartArtDataNode.prototype.getNodeByPtType = function (elementTypeValue) {
		switch (elementTypeValue) {
			case AscFormat.ElementType_value_sibTrans:
				return this.sibNode;
			case AscFormat.ElementType_value_node:
				if (this.isNode() || this.isAsst()) {
					return this;
				}
				break;
			case AscFormat.ElementType_value_parTrans:
				return this.parNode;
			case AscFormat.ElementType_value_all:
				return this;
			case AscFormat.ElementType_value_asst:
				if (this.isAsst()) {
					return this;
				}
				break;
			case AscFormat.ElementType_value_doc:
				if (this.isRoot()) {
					return this;
				}
				break;

			case AscFormat.ElementType_value_nonAsst:
				if (!this.isAsst()) {
					return this;
				}
				break;
			case AscFormat.ElementType_value_nonNorm:
				break;
			case AscFormat.ElementType_value_norm:
				break;
			case AscFormat.ElementType_value_pres:
				break;
			default:
				return this;
		}
	}

	SmartArtDataNode.prototype.getSibName = function () {
		return this.sibNode.getPresName();
	};
	SmartArtDataNode.prototype.getParName = function () {
		return this.parNode.getPresName();
	};

	SmartArtDataNode.prototype.checkName = function (name) {
		switch (name) {
			case this.getPresName():
				return this;
			case this.getSibName():
				return this.sibNode;
			case this.getParName():
				return this.parNode;
			default:
				return null;
		}
	}


	SmartArtDataNode.prototype.getPrSet = function () {
		const presNode = this.getPresNode();
		return presNode && presNode.getPrSet();
	}

	function Position(node) {
		this.rot = 0;
		this.x = 0;
		this.y = 0;
		this.width = 0;
		this.height = 0;
		this.node = node;
		this.cleanParams = {
			x: 0,
			y: 0,
			width: 0,
			height: 0
		};
	}
	Position.prototype.getMatrix = function () {
		const matrix = new AscCommon.CMatrix();
		matrix.RotateAt(this.rot * radToDeg, this.x + this.width / 2, this.y + this.height / 2);
		return matrix;
	};
	Position.prototype.getReverseMatrix = function () {
		const matrix = new AscCommon.CMatrix();
		matrix.RotateAt(-this.rot * radToDeg, this.x + this.width / 2, this.y + this.height / 2);
		return matrix;
	};
	Position.prototype.initFromShape = function () {

	};
	Position.prototype.moveTo = function (dx, dy) {
		this.node.moveTo(dx, dy, true);
	};
	Position.prototype.checkBounds = function (bounds, isClean) {
		const shapeBounds = this.getBounds(isClean);
		checkBounds(bounds, shapeBounds);
	};
	Position.prototype.getBounds = function (isClean, skipRotate) {
		const res = {
			isEllipse: this.node.layoutInfo.shape.type === AscFormat.LayoutShapeType_shapeType_ellipse
		};
		const pos = isClean ? this.cleanParams : this;
		let minX;
		let maxX;
		let minY;
		let maxY;
		if (skipRotate) {
			minX = pos.x;
			maxX = pos.x + pos.width;
			minY = pos.y;
			maxY = pos.y + pos.height;
		} else {
			const matrix = this.getMatrix();
			const x1 = matrix.TransformPointX(pos.x, pos.y);
			const x2 = matrix.TransformPointX(pos.x + pos.width, pos.y);
			const x3 = matrix.TransformPointX(pos.x, pos.y + pos.height);
			const x4 = matrix.TransformPointX(pos.x + pos.width, pos.y + pos.height);
			const y1 = matrix.TransformPointY(pos.x, pos.y);
			const y2 = matrix.TransformPointY(pos.x + pos.width, pos.y);
			const y3 = matrix.TransformPointY(pos.x, pos.y + pos.height);
			const y4 = matrix.TransformPointY(pos.x + pos.width, pos.y + pos.height);


			minX = Math.min(x1, x2, x3, x4);
			maxX = Math.max(x1, x2, x3, x4);
			minY = Math.min(y1, y2, y3, y4);
			maxY = Math.max(y1, y2, y3, y4);
		}
		if (pos.width < 0) {
			res.l = maxX;
			res.r = minX;
		} else {
			res.l = minX;
			res.r = maxX;
		}
		if (pos.height < 0) {
			res.t = maxY;
			res.b = minY;
		} else {
			res.t = minY;
			res.b = maxY;
		}
		return res;
	};
	function ShadowShape(node) {
		Position.call(this, node);
		this.rot = 0;
		this.type = AscFormat.LayoutShapeType_outputShapeType_none;
		this.ln = null;
		this.fill = null;
		this.tailLnArrow = null;
		this.headLnArrow = null;
		this.style = null;
		this.shape = null;
		this.connectorShape = null;
		this.customAdj = null;
		this.customGeom = [];
		this.radialVector = null;
		this.incAngle = null;
	}
	AscFormat.InitClassWithoutType(ShadowShape, Position);
	ShadowShape.prototype.setRadialInfo = function (radialVector, incAngle) {
		const nodes = [this.node];
		while (nodes.length) {
			const node = nodes.pop();
			if (node.algorithm instanceof CompositeAlgorithm) {
				nodes.push.apply(nodes, node.childs);
			} else {
				const shape = node.getShape();
				shape.radialVector = radialVector;
				shape.incAngle = incAngle;
			}
		}
	};
	ShadowShape.prototype.getCustomRadialInfo = function () {
		if (this.radialVector && AscFormat.isRealNumber(this.incAngle)) {
			const prSet = this.node.getPrSet();
			if (prSet) {
				const custScaleRadius = prSet.custRadScaleRad === null ? 1 : prSet.custRadScaleRad;
				const custScaleAngle = prSet.custRadScaleInc === null ? 0 : prSet.custRadScaleInc;
				const defaultRadius = this.radialVector.getDistance();
				const defaultAngle = this.radialVector.getAngle();
				if (custScaleRadius !== 1 || custScaleAngle !== 0) {
					const customRadius = defaultRadius * custScaleRadius;
					const custAngle = AscFormat.normalizeRotate(defaultAngle + custScaleAngle * this.incAngle);
					return {radius: customRadius, angle: custAngle};
				}
				return {radius: defaultRadius, angle: defaultAngle};
			}
		}
	};
	ShadowShape.prototype.moveTo = function (dX, dY) {
		this.node.moveTo(dX, dY, false);
	};
	ShadowShape.prototype.changeSize = function (coefficient, props) {
		props = props || {};
		this.x *= coefficient;
		this.width *= coefficient;
		if (props.changeHeight !== false) {
			this.y *= coefficient;
			this.height *= coefficient;
		}
	}
	ShadowShape.prototype.initSizesFromConstraints = function () {

	}
	ShadowShape.prototype.initFromShape = function (shape) {
		this.shape = shape;
		if (!(shape.hideGeom && (this.node.algorithm instanceof SpaceAlgorithm))) {
			this.type = shape.type;
		}

	}

	ShadowShape.prototype.setFill = function (fill) {
		this.fill = fill;
	};

	ShadowShape.prototype.setLn = function (ln) {
		this.ln = ln;
	}

	ShadowShape.prototype.setStyle = function (style) {
		this.style = style;
	}

	ShadowShape.prototype.getEditorLine = function (initObjects) {

		const shapeTrack = new AscFormat.NewShapeTrack("", this.x, this.y, initObjects.theme, initObjects.master, initObjects.layout, initObjects.slide, 0, undefined, undefined, undefined, true);
		shapeTrack.track({}, this.x + this.width, this.y + this.height);
		const shape = shapeTrack.getShape(false, initObjects.drawingDocument, initObjects.drawingObjects);
		const spPr = shape.spPr;
		spPr.setFill(null);
		spPr.setLn(null);
		spPr.xfrm.setExtX(this.width);
		spPr.xfrm.setExtY(this.height);
		shape.setBDeleted(false);
		shape.setParent(initObjects.parent);
		shape.setWorksheet(initObjects.worksheet);

		const geometry = spPr.geometry;
		for (let i = 0; i < this.customGeom.length; i += 1) {
			const custCommand = this.customGeom[i];
			for (let j = 1; j < custCommand.length; j++) {
				custCommand[j] = String(custCommand[j] >> 0);
			}
			geometry.AddPathCommand.apply(geometry, custCommand);
		}
		this.applyPostEditorSettings(shape);
		return shape;
	};
	ShadowShape.prototype.getEditorShape = function (isLine, initObjects) {
		if (this.connectorShape) {
			return this.connectorShape.getEditorShape(this.connectorShape.type === AscFormat.LayoutShapeType_outputShapeType_conn, initObjects);
		}

		if (isLine) {
			return this.getEditorLine(initObjects);
		} else {
			const shapeType = this.getEditorShapeType();
			if (typeof shapeType !== 'string') {
				return null;
			}

			const shapeTrack = new AscFormat.NewShapeTrack(this.getEditorShapeType(), this.x, this.y, initObjects.theme, initObjects.master, initObjects.layout, initObjects.slide, 0, undefined, undefined, undefined, true);
			shapeTrack.track({}, this.x + this.width, this.y + this.height);
			const shape = shapeTrack.getShape(false, initObjects.drawingDocument, initObjects.drawingObjects);
			const nvSpPr = new AscFormat.UniNvPr();
			nvSpPr.cNvPr.setId(0);
			shape.setNvSpPr(nvSpPr);
			const spPr = shape.spPr;
			spPr.setFill(null);
			spPr.setLn(null);
			spPr.xfrm.setExtX(this.width);
			spPr.xfrm.setExtY(this.height);
			shape.setBDeleted(false);
			shape.setParent(initObjects.parent);
			shape.setWorksheet(initObjects.worksheet);

			this.applyAdjLst(shape.spPr.geometry);
			this.applyPostEditorSettings(shape);
			return shape;
		}
	}
	ShadowShape.prototype.getAdjValueWithApplyFactor = function (adj) {
		const val = adj.val;
		const idx = adj.idx;
		switch (this.type) {
			case AscFormat.LayoutShapeType_shapeType_pie:
			case AscFormat.LayoutShapeType_shapeType_arc:
				return AscFormat.normalizeRotate(val * degToRad) * radToDeg * 60000;
			case AscFormat.LayoutShapeType_shapeType_blockArc:
				switch (idx) {
					case 1:
					case 2:
						return AscFormat.normalizeRotate(val * degToRad) * radToDeg * 60000;
					default:
						return val * 100000;
				}
			case AscFormat.LayoutShapeType_shapeType_circularArrow:
			case AscFormat.LayoutShapeType_shapeType_leftCircularArrow:
				switch (idx) {
					case 2:
					case 3:
					case 4:
						return AscFormat.normalizeRotate(val * degToRad) * radToDeg * 60000;
					default:
						return val * 100000;
				}
			case AscFormat.LayoutShapeType_shapeType_chord:
				return AscFormat.normalizeRotate(val * degToRad) * radToDeg * 60000;
			default:
				return val * 100000;
		}
	};
	function getAdjName(geometry, id) {
		switch (geometry.preset) {
			case "hexagon":
				if (id === 1) {
					return "adj"
				} else if (id === 2) {
					return "vf";
				}
				break;
			default:
				const singleAdjName = "adj";
				const adjName = singleAdjName + id;
				if (geometry.avLst[adjName]) {
					return adjName;
				} else if (geometry.avLst[singleAdjName]) {
					return singleAdjName;
				}
				break;
		}
		return null;
	}
	ShadowShape.prototype.applyAdjLst = function (geometry) {
		const adjLst = this.customAdj || (this.shape && this.shape.adjLst);
		if (adjLst) {
			for (let i = 0; i < adjLst.list.length; i += 1) {
				const adj = adjLst.list[i];
				const adjName = getAdjName(geometry, adj.idx);
				if (adjName) {
					geometry.AddAdj(adjName, 0, this.getAdjValueWithApplyFactor(adj));
				}
			}
		}

	}
	ShadowShape.prototype.setShapeSmartArtInfo = function (editorShape) {
		const shapeSmartArtInfo = new AscFormat.ShapeSmartArtInfo();
		editorShape.setShapeSmartArtInfo(shapeSmartArtInfo);

		const presNode = this.node;
		shapeSmartArtInfo.setShapePoint(presNode.presPoint);
		editorShape.setModelId(presNode.presPoint.getModelId());
		for (let i = presNode.contentNodes.length - 1; i >= 0; i -= 1) {
			const contentNode =  presNode.contentNodes[i];
			shapeSmartArtInfo.addToLstContentPoint(0, contentNode);
		}
	};

	ShadowShape.prototype.applyPostEditorSettings = function (editorShape) {
		this.setShapeSmartArtInfo(editorShape);
		this.applyShapeSettings(editorShape);
		this.applyTextSettings(editorShape);
		this.applyShapeRot(editorShape);
	};

	ShadowShape.prototype.applyShapeRot = function (editorShape) {
		const presNode = this.node;
		let sumRot = this.rot + presNode.moveRot;
		const prSet = presNode.getPrSet();
		if (prSet) {
			if (prSet.custAng) {
				sumRot += prSet.custAng;
			}
		}
		editorShape.spPr.xfrm.setRot(AscFormat.normalizeRotate(sumRot));
	};

	ShadowShape.prototype.applyTextSettings = function (editorShape) {
		const algorithm = this.node.algorithm;
		algorithm.applyTextSettings(editorShape);
	};
	ShadowShape.prototype.applyShapeSettings = function (editorShape) {
		const shapeSmartArtInfo = editorShape.getSmartArtInfo();
		const presPoint = shapeSmartArtInfo.shapePoint;
		const contentPoint = shapeSmartArtInfo.contentPoint[0];
		let skipSettingDefaultSpPr = false;
		if (this.style) {
			editorShape.setStyle(this.style);
			if (this.fill) {
				skipSettingDefaultSpPr = true;
				const fillRef = editorShape.style.fillRef ? editorShape.style.fillRef.createDuplicate() : new AscFormat.StyleRef();
				const color = this.fill.fill.color || null;
				fillRef.setColor(color);
				editorShape.style.setFillRef(fillRef);
			}
			if (this.ln) {
				skipSettingDefaultSpPr = true;
				const lnRef = editorShape.style.lnRef ? editorShape.style.lnRef.createDuplicate() : new AscFormat.StyleRef();
				const color = this.ln.Fill.fill.color || null;
				lnRef.setColor(color);
				editorShape.style.setLnRef(lnRef);
			}
		}
		if (!editorShape.spPr.ln) {
			editorShape.spPr.setLn(new AscFormat.CLn());
		}
		if (this.headLnArrow) {
			editorShape.spPr.ln.setHeadEnd(this.headLnArrow);
		}
		if (this.tailLnArrow) {
			editorShape.spPr.ln.setTailEnd(this.tailLnArrow);
		}


		const style = editorShape.style;
		if (contentPoint && contentPoint.point.prSet && contentPoint.point.prSet.style) {
			style.merge(contentPoint.point.prSet.style);
			skipSettingDefaultSpPr = true;
		}
		if (presPoint && presPoint.prSet && presPoint.prSet.style) {
			style.merge(presPoint.prSet.style);
			skipSettingDefaultSpPr = true;
		}

		const spPr = editorShape.spPr;
		if (!skipSettingDefaultSpPr) {
			if (this.fill) {
				spPr.setFill(this.fill);
			} else {
				spPr.setFill(AscFormat.CreateNoFillUniFill());
			}
			if (this.ln) {
				spPr.setLn(this.ln);
			} else {
				spPr.setLn(AscFormat.CreateNoFillLine());
			}
		}

		if (contentPoint && contentPoint.point.spPr) {
			spPr.fullMerge(contentPoint.point.spPr);
		}
		if (presPoint && presPoint.spPr) {
			spPr.fullMerge(presPoint.spPr);
		}
	};

	ShadowShape.prototype.getEditorShapeType = function () {
		if (this.type !== AscFormat.LayoutShapeType_outputShapeType_none && this.type !== AscFormat.LayoutShapeType_outputShapeType_conn) {
			return AscCommon.To_XML_ST_LayoutShapeType(this.type);
		}
	};


	function BaseAlgorithm() {
		this.params = {};
		this.parentNode = null;
		this._isHideLastChild = null;
		this.constraintSizes = null;
		this.constrClockwise = null;
		this.calcValues = {};
	}
	BaseAlgorithm.prototype.applyTextSettings = function (editorShape) {};
	BaseAlgorithm.prototype.setConstrClockwise = function (isClockwise) {
		this.constrClockwise = isClockwise;
	}
	BaseAlgorithm.prototype.isClockwise = function () {
		return true;
	};
	BaseAlgorithm.prototype.getConstrClockwise = function () {
		return this.constrClockwise;
	}
	BaseAlgorithm.prototype.getChildAlgorithmAlignBounds = function (isCalculateCoefficients, skipRotate) {
		let childShape = this.parentNode.getShape(isCalculateCoefficients);
		if (childShape.connectorShape) {
			childShape = childShape.connectorShape;
		}
		return childShape.getBounds(false, skipRotate);
	};
	BaseAlgorithm.prototype.getAlgorithmAlignBounds = function (isCalculateCoefficients) {
		const childs = this.parentNode.childs;
		let bounds;
		const isCycleAlgorithm = this instanceof CycleAlgorithm;
		for (let i = 0; i < childs.length; i += 1) {
			const node = childs[i];
			if (node.isSkipShape(isCalculateCoefficients, isCycleAlgorithm)) {
				continue;
			}
			const childBounds = node.algorithm.getChildAlgorithmAlignBounds(isCalculateCoefficients, isCycleAlgorithm);
			if (bounds) {
				checkBounds(bounds, childBounds);
			} else {
				bounds = childBounds;
			}
		}
		return bounds;
	};
	BaseAlgorithm.prototype.getHorizontalAlgorithmOffset = function (isCalculateCoefficients) {
		if (this.params[AscFormat.Param_type_horzAlign] === AscFormat.ParameterVal_horizontalAlignment_none) {
			return 0;
		}

		const constrBounds = this.getAlgorithmAlignBounds(isCalculateCoefficients);
		if (!constrBounds) {
			return 0;
		}

		switch (this.params[AscFormat.Param_type_horzAlign]) {
			case AscFormat.ParameterVal_horizontalAlignment_ctr: {
				const parentWidth = this.parentNode.getConstr(AscFormat.Constr_type_w, !isCalculateCoefficients);
				const boundsWidth = constrBounds.r - constrBounds.l;
				return parentWidth / 2 - (constrBounds.l + boundsWidth / 2);
			}
			case AscFormat.ParameterVal_horizontalAlignment_l: {
				return -constrBounds.l;
			}
			case AscFormat.ParameterVal_horizontalAlignment_r: {
				const parentWidth = this.parentNode.getConstr(AscFormat.Constr_type_w, !isCalculateCoefficients);
				return parentWidth - constrBounds.r;
			}
			default:
				return 0;

		}
	};
	BaseAlgorithm.prototype.getVerticalAlgorithmOffset = function (isCalculateCoefficients) {
		if (this.params[AscFormat.Param_type_vertAlign] === AscFormat.ParameterVal_horizontalAlignment_none) {
			return 0;
		}

		const constrBounds = this.getAlgorithmAlignBounds(isCalculateCoefficients);
		if (!constrBounds) {
			return 0;
		}

		switch (this.params[AscFormat.Param_type_vertAlign]) {
			case AscFormat.ParameterVal_verticalAlignment_mid: {
				const boundsHeight = constrBounds.b - constrBounds.t;
				const parentHeight = this.parentNode.getConstr(AscFormat.Constr_type_h, !isCalculateCoefficients);
				return parentHeight / 2 - (constrBounds.t + boundsHeight / 2);
			}
			case AscFormat.ParameterVal_verticalAlignment_t: {
				return -constrBounds.t;
			}
			case AscFormat.ParameterVal_verticalAlignment_b: {
				const parentHeight = this.parentNode.getConstr(AscFormat.Constr_type_h, !isCalculateCoefficients);
				return parentHeight - constrBounds.b;
			}
			default:
				return 0;
		}
	};
	BaseAlgorithm.prototype.applyAlgorithmAligns = function (isCalculateCoefficients) {
		const offX = this.getHorizontalAlgorithmOffset(isCalculateCoefficients);
		const offY = this.getVerticalAlgorithmOffset(isCalculateCoefficients);
		this.parentNode.forEachChild(function (node) {
			node.moveTo(offX, offY, isCalculateCoefficients, true);
		});
	};
	BaseAlgorithm.prototype.getParentConnectionNode = function () {
		return this.parentNode;
	};
	BaseAlgorithm.prototype.setParentConnection = function (connectorAlgorithm, childNode) {
		const parentConnNode = this.getParentConnectionNode();
		if (connectorAlgorithm && childNode.algorithm && parentConnNode) {
			let srcNode;
			let dstNode;
			if (connectorAlgorithm.params[AscFormat.Param_type_srcNode]) {
				srcNode = parentConnNode.getNamedNode(connectorAlgorithm.params[AscFormat.Param_type_srcNode]);
			} else {
				srcNode = parentConnNode.getDefaultConnectionNode();
			}
			if (connectorAlgorithm.params[AscFormat.Param_type_dstNode]) {
				dstNode = childNode.getNamedNode(connectorAlgorithm.params[AscFormat.Param_type_dstNode]);
			} else {
				dstNode = childNode.getDefaultConnectionNode();
			}
			if (srcNode && dstNode) {
				connectorAlgorithm.setParentAlgorithm(this);
				connectorAlgorithm.setFirstConnectorNode(srcNode);
				connectorAlgorithm.setLastConnectorNode(dstNode);
			}
		}
	};
	BaseAlgorithm.prototype.getBounds = function (isCalculateScaleCoefficients, bounds) {
		const shape = this.parentNode.getShape(isCalculateScaleCoefficients);
		if (!bounds) {
			bounds = shape.getBounds();
		} else {
			shape.checkBounds(bounds);
		}
		return bounds;
	};
	BaseAlgorithm.prototype.getRoot = function () {
		return this.parentNode;
	};
	BaseAlgorithm.prototype.getShapes = function (smartartAlgorithm) {
		const shapes = [];
		const initObjects = smartartAlgorithm.getInitObjects();
		if (!initObjects) {
			return shapes;
		}
		const shadowShapes = this.parentNode.getShadowShapesByZOrder();
		smartartAlgorithm.applyColorsDef(this.parentNode.getDescendantFilteredShapes().sort(function (shape1, shape2) {
			return shape2.shape.zOrderOff - shape1.shape.zOrderOff;
		}));
		for (let i = 0; i < shadowShapes.length; i++) {
			const shadowShape = shadowShapes[i];
			const editorShape = shadowShape.getEditorShape(false, initObjects);
			if (editorShape) {
				shadowShape.editorShape = editorShape;
				shapes.push(editorShape);
			}
		}
		return shapes;
	}
	BaseAlgorithm.prototype.applyPostAlgorithmSettings = function (smartartAlgorithm) {
		this.applyPostAlgorithmSettingsForShape(smartartAlgorithm, this.parentNode.getShape(false));
	};
	BaseAlgorithm.prototype.isHierarchy = function () {return false;};
	BaseAlgorithm.prototype.moveToHierarchyOffsets = function () {};
	BaseAlgorithm.prototype.isRootHierarchy = function () {
		return false;
	};
	BaseAlgorithm.prototype.setConstraintSizes = function (shape) {
		this.constraintSizes = {
			x     : shape.x,
			y     : shape.y,
			width : shape.width,
			height: shape.height
		}
	}
	BaseAlgorithm.prototype.setSibConnection = function (startNode, endNode, connectorAlgorithm) {
		let srcNode;
		let dstNode;
		const parentNode = this.parentNode;
		if (connectorAlgorithm.params[AscFormat.Param_type_srcNode]) {
			srcNode = startNode.getNamedNode(connectorAlgorithm.params[AscFormat.Param_type_srcNode]);
			if (!srcNode) {
				srcNode = parentNode.getNamedNode(connectorAlgorithm.params[AscFormat.Param_type_srcNode]);
			}
		} else {
			srcNode = startNode.getDefaultConnectionNode();
		}
		if (connectorAlgorithm.params[AscFormat.Param_type_dstNode]) {
			dstNode = endNode.getNamedNode(connectorAlgorithm.params[AscFormat.Param_type_dstNode]);
			if (!dstNode) {
				dstNode = parentNode.getNamedNode(connectorAlgorithm.params[AscFormat.Param_type_dstNode]);
			}
		} else {
			dstNode = endNode.getDefaultConnectionNode();
		}
		if (srcNode && dstNode) {
			connectorAlgorithm.setParentAlgorithm(this);
			connectorAlgorithm.setFirstConnectorNode(srcNode);
			connectorAlgorithm.setLastConnectorNode(dstNode);
		}
	};
	BaseAlgorithm.prototype.setConnections = function() {
		const nodes = this.parentNode.childs;
		let firstNode;
		let sibConnNode;
		for (let i = 0; i < nodes.length; i++) {
			const node = nodes[i];
			const shape = node.shape;
			if (nodes[i].isRealShapeType()) {
				sibConnNode = nodes[i];
			}
			if (shape.type === AscFormat.LayoutShapeType_outputShapeType_conn) {
				const algorithm = node.algorithm;
				if (node.isParNode() && this instanceof CompositeAlgorithm && sibConnNode) {
					this.setParentConnection(algorithm, sibConnNode);
				} else if (node.isParNode() || sibConnNode) {
					if (!firstNode) {
						firstNode = sibConnNode;
					}
					let nextIndex = i + 1;
					while (nextIndex < nodes.length && !nodes[nextIndex].isRealShapeType()) {
						nextIndex += 1;
					}
					const lastNode = nodes[nodes.length - 1].node;
					const nextNode = nextIndex === nodes.length && !lastNode.isHideLastTrans ? firstNode : nodes[nextIndex];
					if (nextNode) {
						if (node.isParNode()) {
							this.setParentConnection(algorithm, nextNode);
						} else {
							this.setSibConnection(sibConnNode, nextNode, algorithm);
						}
					}
				}
			}
		}
	};
	BaseAlgorithm.prototype.getRadialConnectionInfo = function () {};
	BaseAlgorithm.prototype.setParentAlgorithm = function (algorithm) {};
	BaseAlgorithm.prototype.isHideLastChild = function () {
		if (this._isHideLastChild !== null) {
			return this._isHideLastChild;
		}
		this._isHideLastChild = false;
		if (this.parentNode.isHideLastTrans) {
			const childs = this.parentNode.childs;
			const lastNode = childs[childs.length - 1];
			if (lastNode && lastNode.isSibNode()) {
				this._isHideLastChild = true;
			}
		}
		return this._isHideLastChild;
	};

	BaseAlgorithm.prototype.calcScaleCoefficients = function () {};
	BaseAlgorithm.prototype.setConnectionDistance = function (value, isStart) {

	};
	BaseAlgorithm.prototype.setStemThick = function (value) {

	};
	BaseAlgorithm.prototype.afterShape = function (smartartAlgorithm) {};
	BaseAlgorithm.prototype.createShadowShape = function (isCalculateScaleCoefficients) {
		return this.parentNode.createShadowShape(false, isCalculateScaleCoefficients);
	};
	BaseAlgorithm.prototype.calculateShapePositions = function (smartartAlgorithm, isCalculateScaleCoefficients) {
		this.createShadowShape(isCalculateScaleCoefficients);
	};
	BaseAlgorithm.prototype.setParams = function (params) {
		this.initParams(params);
	};
	BaseAlgorithm.prototype.initParams = function (params) {
		for (let i = 0; i < params.length; i++) {
			const param = params[i];
			this.params[param.type] = param.getValEnum();
		}
	};
	BaseAlgorithm.prototype.getAspectRatio = function () {
		return this.params[AscFormat.Param_type_ar] || 0;
	}


	BaseAlgorithm.prototype.setParentNode = function (node) {
		this.parentNode = node;
	};
	BaseAlgorithm.prototype.setFirstConnectorNode = function () {

	};
	BaseAlgorithm.prototype.setLastConnectorNode = function () {

	};
	BaseAlgorithm.prototype.applyPostAlgorithmSettingsForShape = function (smartartAlgorithm, shape, customCoefficient) {
		const presNode = shape.node;
		const prSet = presNode.getPrSet();
		if (prSet.custLinFactY || prSet.custLinFactX || prSet.custLinFactNeighborY || prSet.custLinFactNeighborX || prSet.custRadScaleInc || prSet.custRadScaleRad) {
			smartartAlgorithm.addSettingsForMove({shape: shape, coefficient: customCoefficient || 1});
		}
	};
	function PositionAlgorithm() {
		BaseAlgorithm.call(this);
		this.connector = null;
		this.shapeContainer = null;
		this.coefficientShapeContainer = null;
	}
	AscFormat.InitClassWithoutType(PositionAlgorithm, BaseAlgorithm);
	PositionAlgorithm.prototype.getParentNodeWidth = function (isAdapt) {
		return this.parentNode.getConstr(AscFormat.Constr_type_w, isAdapt) || this.parentNode.getParentWidth(isAdapt);
	};
	PositionAlgorithm.prototype.getParentNodeHeight = function (isAdapt) {
		return this.parentNode.getConstr(AscFormat.Constr_type_h, isAdapt) || this.parentNode.getParentHeight(isAdapt);
	};
	PositionAlgorithm.prototype.getShapeContainer = function (isCalculateScaleCoefficient) {
		return isCalculateScaleCoefficient ? this.coefficientShapeContainer : this.shapeContainer;
	};

	PositionAlgorithm.prototype.applyParamOffsets = function (isCalculateScaleCoefficient) {
		switch (this.params[AscFormat.Param_type_off]) {
			case AscFormat.ParameterVal_offset_ctr:
				this.applyCenterAlign(isCalculateScaleCoefficient);
				break;
			default:
				break;
		}
	};

	PositionAlgorithm.prototype.applyCenterAlign = function (isCalculateScaleCoefficient) {
		const parentHeight = this.parentNode.getConstr(AscFormat.Constr_type_h, !isCalculateScaleCoefficient)/* || this.parentNode.getParentHeight(!isCalculateScaleCoefficient)*/;
		const parentWidth = this.parentNode.getConstr(AscFormat.Constr_type_w, !isCalculateScaleCoefficient)/* || this.parentNode.getParentWidth(!isCalculateScaleCoefficient)*/;
		if (!(parentWidth && parentHeight)) {
			return
		}

		const shapeContainer = this.getShapeContainer(isCalculateScaleCoefficient);
		shapeContainer.applyCenterAlign(parentHeight, parentWidth, isCalculateScaleCoefficient, this);
	};

	PositionAlgorithm.prototype.calcScaleCoefficients = function () {

	};
	PositionAlgorithm.prototype.calculateShapePositions = function (smartartAlgorithm) {

	};


	function SnakeAlgorithm() {
		PositionAlgorithm.call(this);
		this.calcValues = {
			spacing: null
		};
	}

	AscFormat.InitClassWithoutType(SnakeAlgorithm, PositionAlgorithm);
	SnakeAlgorithm.prototype.initParams = function (params) {
		BaseAlgorithm.prototype.initParams.call(this, params);
		if (this.params[AscFormat.Param_type_flowDir] === undefined) {
			this.params[AscFormat.Param_type_flowDir] = AscFormat.ParameterVal_flowDirection_row;
		}
		if (this.params[AscFormat.Param_type_grDir] === undefined) {
			this.params[AscFormat.Param_type_grDir] = AscFormat.ParameterVal_growDirection_tL;
		}
		if (this.params[AscFormat.Param_type_bkpt] === undefined) {
			this.params[AscFormat.Param_type_bkpt] = AscFormat.ParameterVal_breakpoint_endCnv;
		}
		if (this.params[AscFormat.Param_type_vertAlign] === undefined) {
			this.params[AscFormat.Param_type_vertAlign] = AscFormat.ParameterVal_verticalAlignment_mid;
		}
		if (this.params[AscFormat.Param_type_horzAlign] === undefined) {
			this.params[AscFormat.Param_type_horzAlign] = AscFormat.ParameterVal_horizontalAlignment_ctr;
		}
		if (this.params[AscFormat.Param_type_contDir] === undefined) {
			this.params[AscFormat.Param_type_contDir] = AscFormat.ParameterVal_continueDirection_sameDir;
		}
	};
	SnakeAlgorithm.prototype.applyParamOffsets = function (isCalculateScaleCoefficients) {
		const shapeContainer = this.getShapeContainer(isCalculateScaleCoefficients);
		shapeContainer.applyCenterAlign(this.params[AscFormat.Param_type_off] === AscFormat.ParameterVal_offset_ctr);
	};
	SnakeAlgorithm.prototype.getStartValues = function (node) {
		const oRes = {coefficient: 1, width: 0, height: 0, prSpace: 0};
		if (node) {
			const shape = node.getShape(true);
			oRes.height = shape.height;
			oRes.width = shape.width;

			const smWidth = this.parentNode.getConstr(AscFormat.Constr_type_w);
			const smHeight = this.parentNode.getConstr(AscFormat.Constr_type_h);

			const widthKoef = smWidth / oRes.width;
			const heightKoef = smHeight / oRes.height;
			oRes.coefficient = Math.min(1, widthKoef, heightKoef);
		}
		return oRes;
	};


	SnakeAlgorithm.prototype.calculateFixedRowScaleCoefficient = function () {
		const parentNode = this.parentNode;
		const childs = parentNode.childs;

		const parentWidth = parentNode.getConstr(AscFormat.Constr_type_w);
		const parentHeight = parentNode.getConstr(AscFormat.Constr_type_h);
		const spaceConstr = parentNode.getConstr(AscFormat.Constr_type_sp);

		const countInRow = this.getBreakpointFixedValue();
		let maxWidth = 0;
		let height = -spaceConstr;
		let rowWidth = 0;
		let maxRowHeight = 0;
		let rowShapeCounter = 0;
		for (let i = 0; i < childs.length; i += 1) {
			const node = childs[i];
			const shape = node.getShape(true);
			if (!node.node.isSibNode()) {
				if (rowShapeCounter === countInRow) {
					rowShapeCounter = 0;
					if (maxWidth < rowWidth) {
						maxWidth = rowWidth;
					}
					height += maxRowHeight + spaceConstr;
					maxRowHeight = shape.height;
					rowWidth = shape.width;
				} else {
					rowWidth += shape.width;
					if (shape.height > maxRowHeight) {
						maxRowHeight = shape.height;
					}
				}
				rowShapeCounter += 1;
			} else if (rowShapeCounter !== 0 && rowShapeCounter !== countInRow) {
				rowWidth += shape.width;
			}
		}
		if (rowShapeCounter !== 0) {
			if (maxWidth < rowWidth) {
				maxWidth = rowWidth;
			}
			height += maxRowHeight + spaceConstr;
		}
		const coefficient = Math.min(1, parentWidth / maxWidth, parentHeight / height);
		this.setScaleCoefficient(coefficient);
	};
	SnakeAlgorithm.prototype.calculateFixedColumnScaleCoefficient = function () {
		const parentNode = this.parentNode;
		const childs = parentNode.childs;

		const parentWidth = parentNode.getConstr(AscFormat.Constr_type_w);
		const parentHeight = parentNode.getConstr(AscFormat.Constr_type_h);
		const spaceConstr = parentNode.getConstr(AscFormat.Constr_type_sp);

		const countInColumn = this.getBreakpointFixedValue();

		let maxHeight = 0;
		let width = -spaceConstr;
		let columnHeight = 0;
		let maxColumnWidth = 0;
		let columnShapeCounter = 0;
		for (let i = 0; i < childs.length; i += 1) {
			const node = childs[i];
			const shape = node.getShape(true);
			if (!node.node.isSibNode()) {
				if (columnShapeCounter === countInColumn) {
					columnShapeCounter = 0;
					if (maxHeight < columnHeight) {
						maxHeight = columnHeight;
					}
					width += maxColumnWidth + spaceConstr;
					maxColumnWidth = shape.width;
					columnHeight = shape.height;
				} else {
					columnHeight += shape.height;
					if (shape.width > maxColumnWidth) {
						maxColumnWidth = shape.width;
					}
				}
				columnShapeCounter += 1;
			} else if (columnShapeCounter !== 0 && columnShapeCounter !== countInColumn) {
				columnHeight += shape.height;
			}
		}
		if (columnShapeCounter !== 0) {
			if (maxHeight < columnHeight) {
				maxHeight = columnHeight;
			}
			width += maxColumnWidth + spaceConstr;
		}
		const coefficient = Math.min(1, parentWidth / width, parentHeight / maxHeight);
		this.setScaleCoefficient(coefficient);
	};
	SnakeAlgorithm.prototype.setScaleCoefficient = function (coefficient) {
		this.parentNode.forEachDes(function (node) {
			node.setSizesScale(coefficient, coefficient);
		});
		const space = this.parentNode.getConstr(AscFormat.Constr_type_sp);
		this.calcValues.spacing = coefficient * space;
	}
	SnakeAlgorithm.prototype.calculateCanvasRowScaleCoefficient = function () {
		const parentNode = this.parentNode;
		const childs = parentNode.childs;

		const parentWidth = parentNode.getConstr(AscFormat.Constr_type_w);
		const parentHeight = parentNode.getConstr(AscFormat.Constr_type_h);
		const spaceConstr = parentNode.getConstr(AscFormat.Constr_type_sp);
		const initValues = this.getStartValues(parentNode.childs[0]);
		let coefficient = initValues.coefficient;
		let i = 1;
		let nNeedRecalc = 0;
		function calculateAdaptCoefficient() {
			let prSpaceWidth = 0;
			let previousRowHeight = 0;
			let columnWidth = initValues.width;
			let rowHeight = initValues.height;
			let previousRowSpace = 0;
			let previousMaxWidthCoefficient = null;
			for (i; i < childs.length; i++) {
				const child = parentNode.childs[i];
				const childShape = child.getShape(true);
				if (child.node.isSibNode()) {
					prSpaceWidth = childShape.width;
				} else {
					const width = childShape.width;
					const height = childShape.height;

					const sumWidth = columnWidth + width + prSpaceWidth;
					const sumHeight = rowHeight + height + spaceConstr;

					const updatePreviousRowHeight = previousRowHeight + height + previousRowSpace;


					let widthCoefficient = 1;
					let heightCoefficient = 1;
					widthCoefficient = parentWidth / sumWidth;
					heightCoefficient = parentHeight / sumHeight;
					const tempCoefficient = Math.min(coefficient, Math.max(widthCoefficient, heightCoefficient));
					const nodeWidthCoefficient = parentWidth / width;
					const nodeHeightCoefficient = parentHeight / updatePreviousRowHeight;
					let addToWidth = false;


					if ((heightCoefficient < 1) && (heightCoefficient > widthCoefficient) && (nodeWidthCoefficient < tempCoefficient)) {
						coefficient = Math.min(coefficient, /*nodeWidthCoefficient,*/ nodeWidthCoefficient);
						addToWidth = true;
					} else if ((updatePreviousRowHeight > rowHeight) && (widthCoefficient < 1) && (widthCoefficient > nodeHeightCoefficient) && (nodeHeightCoefficient < tempCoefficient)) {
						if (previousRowHeight > height || nNeedRecalc >= i) {
							coefficient = Math.min(coefficient, nodeHeightCoefficient);
						} else {
							coefficient = Math.min(coefficient, parentHeight / height);
							return true;
						}
					} else {
						if (previousMaxWidthCoefficient !== null && previousMaxWidthCoefficient < coefficient && tempCoefficient < coefficient && previousMaxWidthCoefficient >= heightCoefficient) {
							coefficient = previousMaxWidthCoefficient;
							return true;
						} else if (widthCoefficient < coefficient && (previousMaxWidthCoefficient === null || previousMaxWidthCoefficient < widthCoefficient)) {
							previousMaxWidthCoefficient = widthCoefficient;
						}
						addToWidth = widthCoefficient >= tempCoefficient;
						coefficient = tempCoefficient;

					}
					if (addToWidth) {
						columnWidth = sumWidth;
						rowHeight = Math.max(rowHeight, updatePreviousRowHeight);
						// todo need optimize
						if (nNeedRecalc < i) {
							return true
						}
					} else {
						previousRowSpace = spaceConstr;
						previousRowHeight = rowHeight;
						rowHeight = sumHeight;
						columnWidth = width;
					}
				}
			}
			return false;
		}

		while (calculateAdaptCoefficient()) {
			nNeedRecalc = i;
			i = 1;
		}
		this.setScaleCoefficient(coefficient);
	};
	SnakeAlgorithm.prototype.createShadowShape = function (isCalculateScaleCoefficients) {
		return this.parentNode.createShadowShape(false, isCalculateScaleCoefficients);
	};
	SnakeAlgorithm.prototype.calculateShapePositions = function (smartartAlgorithm, isCalculateScaleCoefficients) {
		if (isCalculateScaleCoefficients) {
			this.calcScaleCoefficients();
		} else {
			this._calculateShapePositions();
			this.applyParamOffsets();
			this.applyAlgorithmAligns();
			this.applyPostAlgorithmSettings();
			this.setConnections();
		}
		this.createShadowShape(isCalculateScaleCoefficients);
	}
	SnakeAlgorithm.prototype._calculateShapePositions = function () {
		if (this.isColumn()) {
			this.calculateColumnSnakeShapePositions();
		} else {
			this.calculateRowSnakeShapePositions();
		}
	}
	SnakeAlgorithm.prototype.getMainElements = function () {
		const childs = this.parentNode.childs;
		const mainElements = [];
		for (let i = 0; i < childs.length; i += 1) {
			const node = childs[i];
			if (!node.node.isSibNode()) {
				mainElements.push(node);
			}
		}
		return mainElements;
	};
	SnakeAlgorithm.prototype.getBreakpointFixedValue = function () {
		if (this.params[AscFormat.Param_type_bkPtFixedVal] !== undefined) {
			return this.params[AscFormat.Param_type_bkPtFixedVal];
		}
		const mainElements = this.getMainElements();
		return Math.ceil(Math.sqrt(mainElements.length));

	};
	SnakeAlgorithm.prototype.calcScaleCoefficients = function () {
		if (this.isColumn()) {
			if (this.params[AscFormat.Param_type_bkpt] !== AscFormat.ParameterVal_breakpoint_endCnv) {
				this.calculateFixedColumnScaleCoefficient();
			}
		} else {
			if (this.params[AscFormat.Param_type_bkpt] === AscFormat.ParameterVal_breakpoint_endCnv) {
				this.calculateCanvasRowScaleCoefficient();
			} else {
				this.calculateFixedRowScaleCoefficient();
			}
		}
	};
	SnakeAlgorithm.prototype.isReverseNextLine = function () {
		return this.params[AscFormat.Param_type_contDir] === AscFormat.ParameterVal_continueDirection_revDir;
	};
	SnakeAlgorithm.prototype.addGridToParentContainer = function (parentContainer, gridContainer, isReverse) {
		if (isReverse) {
			gridContainer.reverse();
		}

		switch (this.params[AscFormat.Param_type_grDir]) {
			case AscFormat.ParameterVal_growDirection_tL:
				parentContainer.push(gridContainer);
				break;
			case AscFormat.ParameterVal_growDirection_bR:
				parentContainer.unshift(gridContainer);
				break;
			case AscFormat.ParameterVal_growDirection_tR:
				if (this.isColumn()) {
					parentContainer.unshift(gridContainer);
				} else {
					parentContainer.push(gridContainer);
				}
				break;
			case AscFormat.ParameterVal_growDirection_bL:
				if (this.isColumn()) {
					parentContainer.push(gridContainer);
				} else {
					parentContainer.unshift(gridContainer);
				}
				break;
			default:
				break;
		}
	};
	SnakeAlgorithm.prototype.isPushInCurrentLine = function (containerLength) {
		let defaultValue;
		switch (this.params[AscFormat.Param_type_grDir]) {
			case AscFormat.ParameterVal_growDirection_tR:
				defaultValue = this.isColumn();
				break;
			case AscFormat.ParameterVal_growDirection_tL:
				defaultValue = true;
				break;
			case AscFormat.ParameterVal_growDirection_bR:
				defaultValue = false;
				break;
			case AscFormat.ParameterVal_growDirection_bL:
				defaultValue = !this.isColumn();
				break;
			default:
				defaultValue = true;
				break;
		}
		return this.isReverseNextLine() && (containerLength % 2 === 1) ? !defaultValue : defaultValue;
	};
	SnakeAlgorithm.prototype.addShapeToRow = function (row, shape, isPush, offX) {
		row.push(shape);
		if (isPush) {
			shape.moveTo(offX - shape.x, -shape.y);
			return shape.x + shape.width;
		} else {
			shape.moveTo(offX - (shape.x + shape.width), -shape.y);
			return shape.x;
		}
	};
	SnakeAlgorithm.prototype.addShapeToColumn = function (column, shape, isPush, offY) {
		column.push(shape);
		if (isPush) {
			shape.moveTo(-shape.x, offY - shape.y);
			return shape.y + shape.height;
		} else {
			shape.moveTo(-shape.x, offY - (shape.y + shape.height));
			return shape.y;
		}
	};
	SnakeAlgorithm.prototype.calculateRowSnakeShapePositions = function () {
		const childs = this.parentNode.childs;
		const parentWidth = this.parentNode.getConstr(AscFormat.Constr_type_w, true);
		const constrSpace = this.calcValues.spacing;
		const rows = this.getShapeContainer();
		rows.setGap(constrSpace);
		let row = new ShapeContainer();
		let isPushInCurrentLine = this.isPushInCurrentLine(0);
		let sibSpacingShapes = [];
		let sibWidth = 0;
		let offX = 0;
		let startX = 0;
		for (let i = 0; i < childs.length; i += 1) {
			const child = childs[i];
			const shape = child.getShape(false);
			const newRowWidth = Math.abs(offX - startX) + shape.width + sibWidth;
			if (newRowWidth > parentWidth && !fAlgDeltaEqual(newRowWidth, parentWidth)) {
				this.addGridToParentContainer(rows, row, !isPushInCurrentLine);
				isPushInCurrentLine = this.isPushInCurrentLine(rows.getLength());
				if (isPushInCurrentLine) {
					const positionShape = row.shapes[0];
					offX = positionShape.x;
					startX = offX;
				} else {
					const positionShape = row.shapes[row.shapes.length - 1];
					offX = positionShape.x + positionShape.width;
					startX = offX;
				}

				row = new ShapeContainer();
				sibSpacingShapes = [];
				sibWidth = 0;
				if (!child.node.isSibNode()) {
					offX = this.addShapeToRow(row, shape, isPushInCurrentLine, offX);
				}
			} else {
				if (!child.node.isSibNode()) {
					for (let j = 0; j < sibSpacingShapes.length; j += 1) {
						 offX = this.addShapeToRow(row, sibSpacingShapes[j], isPushInCurrentLine, offX);
					}
					sibSpacingShapes = [];
					sibWidth = 0;
					offX = this.addShapeToRow(row, shape, isPushInCurrentLine, offX);
				} else if (row.shapes.length !== 0) {
					sibSpacingShapes.push(shape);
					sibWidth += shape.width;
				}
			}
		}
		if (row !== rows.rows[rows.rows.length - 1]) {
			this.addGridToParentContainer(rows, row);
		}
		rows.applyVerticalPositions();
	};

	SnakeAlgorithm.prototype.calculateColumnSnakeShapePositions = function () {
		const childs = this.parentNode.childs;
		const parentHeight = this.parentNode.getConstr(AscFormat.Constr_type_h, true);
		const constrSpace = this.calcValues.spacing;
		const columns = this.getShapeContainer();
		columns.setGap(constrSpace);
		let column = new ShapeContainer();
		let isPushInCurrentLine = this.isPushInCurrentLine(0);
		let sibSpacingShapes = [];
		let sibSpacingHeight = 0;
		let offY = 0;
		let startY = 0;
		for (let i = 0; i < childs.length; i += 1) {
			const child = childs[i];
			const shape = child.getShape(false);
			const newRowHeight = Math.abs(offY - startY) + shape.height + sibSpacingHeight;
			if (newRowHeight > parentHeight && !fAlgDeltaEqual(newRowHeight, parentHeight)) {
				this.addGridToParentContainer(columns, column, !isPushInCurrentLine);
				isPushInCurrentLine = this.isPushInCurrentLine(columns.getLength());
				if (isPushInCurrentLine) {
					const positionShape = column.shapes[0];
					offY = positionShape.y;
					startY = offY;
				} else {
					const positionShape = column.shapes[column.shapes.length - 1];
					offY = positionShape.y + positionShape.height;
					startY = offY;
				}

				column = new ShapeContainer();
				sibSpacingShapes = [];
				sibSpacingHeight = 0;
				if (!child.node.isSibNode()) {
					offY = this.addShapeToColumn(column, shape, isPushInCurrentLine, offY);
				}
			} else {
				if (!child.node.isSibNode()) {
					for (let j = 0; j < sibSpacingShapes.length; j += 1) {
						offY = this.addShapeToColumn(column, sibSpacingShapes[j], isPushInCurrentLine, offY);
					}
					sibSpacingShapes = [];
					sibSpacingHeight = 0;
					offY = this.addShapeToColumn(column, shape, isPushInCurrentLine, offY);
				} else if (column.shapes.length !== 0) {
					sibSpacingShapes.push(shape);
					sibSpacingHeight += shape.height;
				}
			}
		}
		if (column !== columns.columns[columns.columns.length - 1]) {
			this.addGridToParentContainer(columns, column);
		}
		columns.applyHorizontalPositions();
	};
	SnakeAlgorithm.prototype.isColumn = function () {
		return this.params[AscFormat.Param_type_flowDir] === AscFormat.ParameterVal_flowDirection_col;
	}
	SnakeAlgorithm.prototype.getShapeContainer = function (isCalculateScaleCoefficient) {
		if (isCalculateScaleCoefficient) {
			if (this.coefficientShapeContainer === null) {
				this.coefficientShapeContainer = this.isColumn() ? new ShapeColumns() : new ShapeRows();
			}
			return this.coefficientShapeContainer;
		}
		if (this.shapeContainer === null) {
			this.shapeContainer = this.isColumn() ? new ShapeColumns() : new ShapeRows();
		}
		return this.shapeContainer;
	};

	function ContainerBase() {

	}
	ContainerBase.prototype.forEachShape = function (callback) {
	};
	ContainerBase.prototype.applyCenterAlign = function (parentHeight, parentWidth, isCalculateScaleCoefficients, algorithm) {
	};

	function ShapeContainer() {
		ContainerBase.call(this);
		this.shapes = [];
		this.bounds = null;
	}
	AscFormat.InitClassWithoutType(ShapeContainer, ContainerBase);
	ShapeContainer.prototype.forEachShape = function (callback) {
		for (let i = 0; i < this.shapes.length; i += 1) {
			callback(this.shapes[i]);
		}
	};
	ShapeContainer.prototype.reverse = function () {
		return this.shapes.reverse();
	};
	ShapeContainer.prototype.push = function (shape) {
		this.shapes.push(shape);
	};
	ShapeContainer.prototype.unshift = function (shape) {
		this.shapes.unshift(shape);
	};
	ShapeContainer.prototype.getBounds = function (isCalculateScaleCoefficient, isClean) {
			if (this.shapes.length) {
				const firstShape = this.shapes[0];
				this.bounds = firstShape.getBounds(isClean);
				for (let i = 1; i < this.shapes.length; i += 1) {
					this.shapes[i].checkBounds(this.bounds, isClean);
				}
			} else {
				this.bounds = {l: 0, r: 0, t: 0, b: 0};
			}
		return this.bounds;
	};
	ShapeContainer.prototype.getOffsets = function (parentHeight, parentWidth, isCalculateScaleCoefficient) {
		const bounds = this.getBounds(isCalculateScaleCoefficient);
		const ctrX = bounds.l + (bounds.r - bounds.l) / 2;
		const ctrY = bounds.t + (bounds.b - bounds.t) / 2;
		const offX = parentWidth / 2 - ctrX;
		const offY = parentHeight / 2 - ctrY;
		return {offX: offX, offY: offY};
	}
	ShapeContainer.prototype.applyCenterAlign = function (parentHeight, parentWidth, isCalculateScaleCoefficient, algorithm) {
		const offsets = this.getOffsets(parentHeight, parentWidth, isCalculateScaleCoefficient);
		algorithm.moveToHierarchyOffsets(offsets.offX, offsets.offY);
		for (let i = 0; i < this.shapes.length; i++) {
			const shape = this.shapes[i];
			const node = shape.node;
			node.moveTo(offsets.offX, offsets.offY, isCalculateScaleCoefficient);
		}
	};
	function PyramidContainer() {
		ShapeContainer.call(this)
	}
	AscFormat.InitClassWithoutType(PyramidContainer, ShapeContainer);

	function HierarchyChildContainer() {
		ShapeContainer.call(this);
	}
	AscFormat.InitClassWithoutType(HierarchyChildContainer, ShapeContainer);
	HierarchyChildContainer.prototype.getBounds = function (isCalculateScaleCoefficient) {
		let bounds;
		for (let i = 0; i < this.shapes.length; i += 1) {
			const shape = this.shapes[i];
			const algorithm = shape.node.algorithm;
			bounds = algorithm.getBounds(isCalculateScaleCoefficient, bounds);
		}
		if (!bounds) {
			return {l: 0, r: 0, t: 0, b: 0};
		}
		return bounds;
	};
	function HierarchyRootContainer() {
		ShapeContainer.call(this);
	}
	AscFormat.InitClassWithoutType(HierarchyRootContainer, ShapeContainer);
	function CycleContainer() {
		ShapeContainer.call(this);
	}
	AscFormat.InitClassWithoutType(CycleContainer, ShapeContainer);
	CycleContainer.prototype.getBounds = function (isCalculateScaleCoefficient) {
			if (this.shapes.length) {
				const firstShape = this.shapes[0];
				const bounds = {
					l: firstShape.x,
					t: firstShape.y,
					b: firstShape.y + firstShape.height,
					r: firstShape.x + firstShape.width
				};
				this.bounds = bounds;
				for (let i = 0; i < this.shapes.length; i += 1) {
					const shape = this.shapes[i];
					if (shape.x < bounds.l) {
						bounds.l = shape.x;
					}
					if (shape.y < bounds.t) {
						bounds.t = shape.y;
					}

					const r = shape.x + shape.width;
					if (r > bounds.r) {
						bounds.r = r;
					}

					const b = shape.y + shape.height;
					if (b > bounds.b) {
						bounds.b = b;
					}
				}
			} else {
				this.bounds = {l: 0, r: 0, b: 0, t: 0};
			}
		return this.bounds;
	};


	function ShapeRows() {
		ContainerBase.call(this);
		this.bounds = null;
		this.rows = [];
		this.gap = 0;
	}
	AscFormat.InitClassWithoutType(ShapeRows, ContainerBase);
	ShapeRows.prototype.getLength = function () {
		return this.rows.length;
	};
	ShapeRows.prototype.setGap = function (value) {
		this.gap = value;
	};
	ShapeRows.prototype.applyVerticalPositions = function () {
		for (let i = 1; i < this.rows.length; i += 1) {
			const currentRow = this.rows[i];
			const currentRowBounds = currentRow.getBounds();
			const previousRowBounds = this.rows[i - 1].getBounds();

			const offY = previousRowBounds.b + this.gap - currentRowBounds.t;
			for (let j = 0; j < currentRow.shapes.length; j += 1) {
				const shape = currentRow.shapes[j];
				shape.moveTo(0, offY);
			}
		}
	};
	ShapeRows.prototype.push = function (elem) {
		this.rows.push(elem);
	}
	ShapeRows.prototype.unshift = function (elem) {
		this.rows.unshift(elem);
	}
	ShapeRows.prototype.getBounds = function (isCalculateScaleCoefficient, isClean) {
			if (this.rows.length) {
				const firstBounds = Object.assign({}, this.rows[0].getBounds(isCalculateScaleCoefficient, isClean));
				for (let i = 0; i < this.rows.length; i++) {
					const row = this.rows[i];
					const bounds = row.getBounds(isCalculateScaleCoefficient, isClean);
					checkBounds(firstBounds, bounds);
				}
				this.bounds = firstBounds;
			} else {
				this.bounds = {l: 0, r: 0, t: 0, b: 0};
			}
		return this.bounds;
	};

	ShapeRows.prototype.applyCenterAlign = function (isCenteringRows) {
		const bounds = this.getBounds();
		const width = bounds.r - bounds.l;

		for (let i = 0; i < this.rows.length; i++) {
			const row = this.rows[i];
			const rowBounds = row.getBounds();
			const rowHeight = rowBounds.b - rowBounds.t;
			let offRowX = 0;
			if (isCenteringRows) {
				const rowWidth = rowBounds.r - rowBounds.l;
				offRowX = bounds.l + width / 2 - (rowBounds.l + rowWidth / 2);
			}
			for (let j = 0; j < row.shapes.length; j++) {
				const shape = row.shapes[j];
				const offRowY = rowBounds.t + rowHeight / 2 - (shape.y + shape.height / 2);
				shape.moveTo(offRowX, offRowY);
			}
		}
	};
	ShapeRows.prototype.forEachShape = function (callback) {
		for (let i = 0; i < this.rows.length; i += 1) {
			this.rows[i].forEachShape(callback);
		}
	};

	function ShapeColumns() {
		ContainerBase.call(this);
		this.bounds = null;
		this.columns = [];
		this.gap = 0;
	}
	AscFormat.InitClassWithoutType(ShapeColumns, ContainerBase);

	ShapeColumns.prototype.getLength = function () {
		return this.columns.length;
	};
	ShapeColumns.prototype.push = function (elem) {
		this.columns.push(elem);
	}
	ShapeColumns.prototype.unshift = function (elem) {
		this.columns.unshift(elem);
	}
	ShapeColumns.prototype.setGap = function (value) {
		this.gap = value;
	}
	ShapeColumns.prototype.getBounds = function (isCalculateScaleCoefficient, isClean) {
			if (this.columns.length) {
				const firstBounds = Object.assign({}, this.columns[0].getBounds(isCalculateScaleCoefficient, isClean));
				for (let i = 0; i < this.columns.length; i++) {
					const column = this.columns[i];
					const bounds = column.getBounds(isCalculateScaleCoefficient, isClean);
					checkBounds(firstBounds, bounds);
				}
				this.bounds = firstBounds;
			} else {
				this.bounds = {l: 0, r: 0, t: 0, b: 0};
			}
		return this.bounds;
	};

	ShapeColumns.prototype.applyHorizontalPositions = function () {
		for (let i = 1; i < this.columns.length; i += 1) {
			const currentColumn = this.columns[i];
			const currentColumnBounds = currentColumn.getBounds();
			const previousColumnBounds = this.columns[i - 1].getBounds();

			const offX = previousColumnBounds.r + this.gap - currentColumnBounds.l;
			for (let j = 0; j < currentColumn.shapes.length; j += 1) {
				const shape = currentColumn.shapes[j];
				shape.moveTo(offX, 0);
			}
		}
	};

	ShapeColumns.prototype.applyCenterAlign = function (isCenteringColumns) {
		const bounds = this.getBounds();
		const height = bounds.b - bounds.t;

		for (let i = 0; i < this.columns.length; i++) {
			const column = this.columns[i];
			const columnBounds = column.getBounds();
			const columnWidth = columnBounds.r - columnBounds.l;

			let offColumnY = 0;
			if (isCenteringColumns) {
				const columnHeight = columnBounds.b - columnBounds.t;
				offColumnY = height / 2 - (columnBounds.t + columnHeight / 2);
			}
			for (let j = 0; j < column.shapes.length; j++) {
				const shape = column.shapes[j];
				const offColumnX = columnBounds.l + columnWidth / 2 - (shape.x + shape.width / 2);
				shape.moveTo(offColumnX, offColumnY);
			}
		}
	};
	ShapeColumns.prototype.forEachShape = function (callback) {
		for (let i = 0; i < this.columns.length; i += 1) {
			this.columns[i].forEachShape(callback);
		}
	};

function HierarchyAlgorithm() {
	PositionAlgorithm.call(this);
	this.levelPositions = [];
	this.calcValues = {
		mainChilds: null
	};
}
	AscFormat.InitClassWithoutType(HierarchyAlgorithm, PositionAlgorithm);
	HierarchyAlgorithm.prototype.moveToHierarchyOffsets = function (dx, dy) {
		for (let i = 0; i < this.levelPositions.length; i += 1) {
			const levelPosition = this.levelPositions[i];
			levelPosition.l += dx;
			levelPosition.r += dx;
			levelPosition.t += dy;
			levelPosition.b += dy;
		}
	};
	HierarchyAlgorithm.prototype.setLevelBounds = function (bounds) {
		this.levelPositions.push(Object.assign({}, bounds));
	};
	HierarchyAlgorithm.prototype.getHorizontalOffset = function (node, isCalculateScaleCoefficient, fromLeft) {
		const algorithm = node.algorithm;
		let maxSpace = 0;
		const shape = node.getShape(isCalculateScaleCoefficient);

		for (let i = 0; i < this.levelPositions.length; i += 1) {
			const startBounds = this.levelPositions[i];
			for (let j = 0; j < algorithm.levelPositions.length; j += 1) {
				const endBounds = algorithm.levelPositions[j];
				let levelDifference;
				if (fromLeft) {
					levelDifference = startBounds.r - endBounds.l;
				} else {
					levelDifference = endBounds.r - startBounds.l;
				}
				if (levelDifference > maxSpace && !!this.getVerticalIntersection(startBounds, endBounds)) {
					maxSpace = levelDifference;
				}
			}
			const shapeBounds = shape.getBounds();
			let nodeDifference;
			if (fromLeft) {
				nodeDifference = startBounds.r - shapeBounds.l;
			} else {
				nodeDifference = shapeBounds.r - startBounds.l;
			}
			if (nodeDifference > maxSpace && !!this.getVerticalIntersection(shapeBounds, startBounds)) {
				maxSpace = nodeDifference;
			}
		}
		return maxSpace;
	};
	HierarchyAlgorithm.prototype.getHorizontalIntersection = function (bounds1, bounds2) {
		return Math.max(Math.min(bounds1.r, bounds2.r) - Math.max(bounds1.l, bounds2.l), 0);
	};
	HierarchyAlgorithm.prototype.getVerticalIntersection = function (bounds1, bounds2) {
		return Math.max(Math.min(bounds1.b, bounds2.b) - Math.max(bounds1.t, bounds2.t), 0);
	};
	HierarchyAlgorithm.prototype.getVerticalOffset = function (node, isCalculateScaleCoefficient) {
		const algorithm = node.algorithm;
		let maxSpace = 0;
		const shape = node.getShape(isCalculateScaleCoefficient);
		for (let i = 0; i < this.levelPositions.length; i += 1) {
			const startBounds = this.levelPositions[i];
			for (let j = 0; j < algorithm.levelPositions.length; j += 1) {
				const endBounds = algorithm.levelPositions[j];
				const levelDifference = startBounds.b - endBounds.t;
				if (levelDifference > maxSpace && !!this.getHorizontalIntersection(startBounds, endBounds)) {
					maxSpace = levelDifference;
				}
			}
			const shapeBounds = shape.getBounds();
			const nodeDifference = startBounds.b - shapeBounds.t;
			if (nodeDifference > maxSpace && !!this.getHorizontalIntersection(shapeBounds, startBounds)) {
				maxSpace = nodeDifference;
			}
		}
		return maxSpace;
	};
	HierarchyAlgorithm.prototype.isHierarchy = function () {
		return true;
	};
	HierarchyAlgorithm.prototype.getMainChilds = function () {
		if (this.calcValues.mainChilds === null) {
			const childs = [];
			for (let i = 0; i < this.parentNode.childs.length; i += 1) {
				const child = this.parentNode.childs[i];
				if (child.isContentNode()) {
					childs.push(child);
				}
			}
			this.calcValues.mainChilds = childs;
		}
		return this.calcValues.mainChilds;
	};
	HierarchyAlgorithm.prototype._calculateShapePositions = function (isAdapt) {

	};
	HierarchyAlgorithm.prototype.setScaleCoefficient = function () {
		if (this.parentNode.parent && !(this.parentNode.parent.algorithm instanceof CompositeAlgorithm)) {
			return;
		}
		const childs = this.getMainChilds();
		const parentHeight = this.parentNode.getConstr(AscFormat.Constr_type_h);
		const parentWidth = this.parentNode.getConstr(AscFormat.Constr_type_w);
		if (!(parentHeight && parentWidth)) {
			return;
		}
		const shapeContainer = this.getShapeContainer(true);
		const bounds = shapeContainer.getBounds(true);
		const height = bounds.b - bounds.t;
		const width = bounds.r - bounds.l;

		const widthCoefficient = Math.min(parentWidth / width, 1);
		const heightCoefficient = Math.min(parentHeight / height, 1);
		const coefficient = Math.min(widthCoefficient, heightCoefficient, 1);
		if (!AscFormat.isRealNumber(coefficient)) {
			return
		}

		this.parentNode.forEachDes(function (node) {
				node.setSizesScale(coefficient, coefficient);
		});
	};

	HierarchyAlgorithm.prototype.updateLevelPositions = function (algorithm) {
		if (algorithm.levelPositions) {
			for (let i = 0; i < algorithm.levelPositions.length; i += 1) {
				this.setLevelBounds(algorithm.levelPositions[i]);
			}
		}
	};
	HierarchyAlgorithm.prototype.resetLevelPositions = function () {
		this.levelPositions = [];
	};
	HierarchyAlgorithm.prototype.collectHierarchyPositions = function () {};
	HierarchyAlgorithm.prototype.putShapesToShapeContainer = function (isCalculateScaleCoefficient) {};
	HierarchyAlgorithm.prototype.calculateShapePositions = function (smartartAlgorithm, isCalculateScaleCoefficients) {
		this.resetLevelPositions();
		this.putShapesToShapeContainer(isCalculateScaleCoefficients);
		this._calculateShapePositions(isCalculateScaleCoefficients);
		this.applyParamOffsets(isCalculateScaleCoefficients);
		this.collectHierarchyPositions();
		if (isCalculateScaleCoefficients) {
			this.setScaleCoefficient();
		} else {
			this.setConnections();
		}
		this.createShadowShape(isCalculateScaleCoefficients);
	};


	function HierarchyChildAlgorithm() {
		HierarchyAlgorithm.call(this);
	}
	AscFormat.InitClassWithoutType(HierarchyChildAlgorithm, HierarchyAlgorithm);

	HierarchyChildAlgorithm.prototype.getHorizontalSibSp = function (node, isCalculateScaleCoefficient, fromLeft) {
		const sibSp = this.parentNode.getConstr(AscFormat.Constr_type_sibSp, !isCalculateScaleCoefficient);
		const offset = this.getHorizontalOffset(node, isCalculateScaleCoefficient, fromLeft);
		if (fromLeft) {
			return sibSp + offset;
		} else {
			return - offset - sibSp;
		}
	};
	HierarchyChildAlgorithm.prototype.getVerticalSibSp = function (node, isCalculateScaleCoefficient) {
		const sibSp = this.parentNode.getConstr(AscFormat.Constr_type_sibSp, !isCalculateScaleCoefficient);
		const offset = this.getVerticalOffset(node, isCalculateScaleCoefficient);
		return sibSp + offset;
	};
	HierarchyChildAlgorithm.prototype.isHorizontalHierarchy = function () {
		const parent = this.parentNode.parent;
		if (parent && parent.algorithm.isHierarchy()) {
			return parent.algorithm.isHorizontalHierarchy();
		} else {
			switch (this.params[AscFormat.Param_type_chAlign]) {
				case AscFormat.ParameterVal_childAlignment_r:
				case AscFormat.ParameterVal_childAlignment_l:
					return true;
				default:
					return false;
			}
		}
	};
	HierarchyChildAlgorithm.prototype.createShadowShape = function (isCalculateScaleCoefficients) {
		return this.parentNode.createHierarchyChildShadowShape(isCalculateScaleCoefficients);
	};
	HierarchyChildAlgorithm.prototype.setParentConnection = function (connectorAlgorithm, childNode) {
		const parent = this.parentNode.parent;
		if (parent && parent.algorithm.isHierarchy()) {
			const root = parent.algorithm.getRoot();
			root.algorithm.setParentConnection(connectorAlgorithm, childNode.algorithm.getRoot());
		}
	};
	HierarchyChildAlgorithm.prototype.initParams = function (params) {
		HierarchyAlgorithm.prototype.initParams.call(this, params);
		if (this.params[AscFormat.Param_type_linDir] === undefined) {
			this.params[AscFormat.Param_type_linDir] = AscFormat.ParameterVal_linearDirection_fromL;
		}
		if (this.params[AscFormat.Param_type_chAlign] === undefined) {
			this.params[AscFormat.Param_type_chAlign] = AscFormat.ParameterVal_childAlignment_t;
		}
		if (this.params[AscFormat.Param_type_off] === undefined) {
			this.params[AscFormat.Param_type_off] = AscFormat.ParameterVal_offset_ctr;
		}
	}

	HierarchyChildAlgorithm.prototype.isHang = function () {
		return this.params[AscFormat.Param_type_secLinDir] !== undefined &&
			this.params[AscFormat.Param_type_linDir] !== undefined;
	}
	HierarchyChildAlgorithm.prototype.getCommonChildBounds = function (isCalculateScaleCoefficient) {
		const childs = this.getMainChilds();
		let bounds;
		for (let i = 0; i < childs.length; i += 1) {
			const child = childs[i];
			bounds = child.algorithm.getBounds(isCalculateScaleCoefficient, bounds);
		}
		return bounds;
	};
	HierarchyChildAlgorithm.prototype.getRootCenteringOffset = function (shape, isCalculateScaleCoefficient) {
		const result = {offX: 0, offY: 0};
		if (this.isHang()) {
			const isHorizontal = this.isHorizontalHierarchy();
			const sibSp = this.parentNode.getConstr(AscFormat.Constr_type_sibSp, !isCalculateScaleCoefficient);
			const bounds = this.getHangStructBounds(isCalculateScaleCoefficient);
			if (!bounds) {
				return result;
			}
			if (isHorizontal) {
				result.offY = shape.y + shape.height / 2 - bounds.b - sibSp / 2;
			} else {
				if (this.params[AscFormat.Param_type_linDir] === AscFormat.ParameterVal_linearDirection_fromR) {
					result.offX = shape.x + shape.width / 2 - bounds.l + sibSp / 2;
				} else {
					result.offX = shape.x + shape.width / 2 - bounds.r - sibSp / 2;
				}
			}
		}
		return result;
	};
	HierarchyChildAlgorithm.prototype.putShapesToShapeContainer = function (isCalculateScaleCoefficient) {
		const childs = this.getMainChilds();
		const shapeContainer = this.getShapeContainer(isCalculateScaleCoefficient);
		for (let i = 0; i < childs.length; i += 1) {
			shapeContainer.push(childs[i].getShape(isCalculateScaleCoefficient));
		}
	}
	HierarchyChildAlgorithm.prototype.calculateVerticalHierarchyVerticalShapePositions = function (isCalculateScaleCoefficient, fromTop) {
		const childs = fromTop ? this.getMainChilds() : this.getMainChilds().slice().reverse();
		const parentNode = this.parentNode;
		const sibSp = parentNode.getConstr(AscFormat.Constr_type_sibSp, !isCalculateScaleCoefficient);
		const commonBounds = this.getCommonChildBounds(isCalculateScaleCoefficient);
		const firstShape = childs[0].getShape(isCalculateScaleCoefficient);
		const firstBounds = childs[0].algorithm.getBounds(isCalculateScaleCoefficient);

		const firstChildOffsets = this.getChildAlignOffsets(commonBounds, firstBounds, this.params[AscFormat.Param_type_chAlign]);
		firstShape.moveTo(firstChildOffsets.offX, firstChildOffsets.offY);

		let offY = firstBounds.b - firstBounds.t + sibSp;
		this.setLevelBounds({l: firstShape.x, t: firstShape.y, b: firstShape.y + firstShape.height, r: firstShape.x + firstShape.width});
		this.updateLevelPositions(childs[0].algorithm);
		let previousShape = firstShape;
		let previousDesHeight = firstBounds.b - firstBounds.t;
		for (let i = 1; i < childs.length; i += 1) {
			const node = childs[i];
			const shape = node.getShape(isCalculateScaleCoefficient);
			const bounds = node.algorithm.getBounds(isCalculateScaleCoefficient);
			const childOffsets = this.getChildAlignOffsets(commonBounds, bounds, this.params[AscFormat.Param_type_chAlign]);
			const offY = previousShape.y + previousDesHeight + sibSp - shape.y;
			shape.moveTo(childOffsets.offX, offY + childOffsets.offY);
			this.setLevelBounds({l: shape.x, t: shape.y, b: shape.y + shape.height, r: shape.x + shape.width});
			this.updateLevelPositions(node.algorithm);
			previousShape = shape;
			previousDesHeight = bounds.b - bounds.t;
		}
	};
	HierarchyChildAlgorithm.prototype.calculateHorizontalHierarchyVerticalShapePositions = function (isCalculateScaleCoefficient, fromTop) {
		const childs = fromTop ? this.getMainChilds() : this.getMainChilds().slice().reverse();
		const commonBounds = this.getCommonChildBounds(isCalculateScaleCoefficient);
		const firstNode = childs[0];
		const firstShape = firstNode.getShape(isCalculateScaleCoefficient);
		const firstBounds = firstNode.algorithm.getBounds(isCalculateScaleCoefficient);
		const firstChildOffsets = this.getChildAlignOffsets(commonBounds, firstBounds, this.params[AscFormat.Param_type_chAlign]);
		firstShape.moveTo(firstChildOffsets.offX, firstChildOffsets.offY);

		this.updateLevelPositions(firstNode.algorithm);
		let previousShape = firstShape;
		this.setLevelBounds({l: firstShape.x, r: firstShape.x + firstShape.width, t: firstShape.y, b: firstShape.y + firstShape.height});
		for (let i = 1; i < childs.length; i += 1) {
			const node = childs[i];
			const shape = node.getShape(isCalculateScaleCoefficient);
			const bounds = node.algorithm.getBounds(isCalculateScaleCoefficient);
			const offY = previousShape.y + previousShape.height - shape.y;
			const childOffsets = this.getChildAlignOffsets(commonBounds, bounds, this.params[AscFormat.Param_type_chAlign]);
			shape.moveTo(childOffsets.offX, offY + childOffsets.offY);
			const sibSp = this.getVerticalSibSp(node, isCalculateScaleCoefficient);
			shape.moveTo(0, sibSp);
			this.updateLevelPositions(node.algorithm);
			this.setLevelBounds({l: shape.x, r: shape.x + shape.width, t: shape.y, b: shape.y + shape.height});
			previousShape = shape;
		}
	};
	HierarchyChildAlgorithm.prototype.setSibLevelBounds = function () {

	};
	HierarchyChildAlgorithm.prototype.calculateVerticalHierarchyHorizontalShapePositions = function (isCalculateScaleCoefficient, fromLeft) {
		const childs = this.getMainChilds();
		const commonBounds = this.getCommonChildBounds(isCalculateScaleCoefficient);
		const firstNode = childs[0];
		const firstShape = firstNode.getShape(isCalculateScaleCoefficient);
		const firstBounds = firstNode.algorithm.getBounds(isCalculateScaleCoefficient);
		const firstAlignOffsets = this.getChildAlignOffsets(commonBounds, firstBounds, this.params[AscFormat.Param_type_chAlign]);
		firstNode.moveTo(firstAlignOffsets.offX, firstAlignOffsets.offY, isCalculateScaleCoefficient);
		this.updateLevelPositions(childs[0].algorithm);
		let previousShape = firstShape;
		this.setLevelBounds({l: firstShape.x, t: firstShape.y, b: firstShape.y + firstShape.height, r: firstShape.x + firstShape.width});
		for (let i = 1; i < childs.length; i += 1) {
			const node = childs[i];
			const shape = node.getShape(isCalculateScaleCoefficient);

			let offX;
			if (fromLeft) {
				offX = previousShape.x + previousShape.width - shape.x;
			} else {
				offX = previousShape.x - (shape.x + shape.width);
			}
			const bounds = node.algorithm.getBounds(isCalculateScaleCoefficient);
			const alignOffsets = this.getChildAlignOffsets(commonBounds, bounds, this.params[AscFormat.Param_type_chAlign]);
			node.moveTo(offX + alignOffsets.offX, alignOffsets.offY, isCalculateScaleCoefficient);
			const offset = this.getHorizontalSibSp(node, isCalculateScaleCoefficient, fromLeft);
			node.moveTo(offset, 0, isCalculateScaleCoefficient);
			previousShape = shape;
			this.updateLevelPositions(node.algorithm);
			this.setLevelBounds({l: shape.x, t: shape.y, b: shape.y + shape.height, r: shape.x + shape.width});
		}
	};
	HierarchyChildAlgorithm.prototype.calculateHorizontalHierarchyHorizontalShapePositions = function (isCalculateScaleCoefficient, fromRight) {
		const childs = fromRight ? this.getMainChilds() : this.getMainChilds().slice().reverse();
		const parentNode = this.parentNode;
		const sibSp = parentNode.getConstr(AscFormat.Constr_type_sibSp, !isCalculateScaleCoefficient);
		const commonBounds = this.getCommonChildBounds(isCalculateScaleCoefficient);
		const firstShape = childs[0].getShape(isCalculateScaleCoefficient);
		const firstBounds = childs[0].algorithm.getBounds(isCalculateScaleCoefficient);

		const firstChildOffsets = this.getChildAlignOffsets(commonBounds, firstBounds, this.params[AscFormat.Param_type_chAlign]);
		firstShape.moveTo(firstChildOffsets.offX, firstChildOffsets.offY);

		this.setLevelBounds({l: firstShape.x, t: firstShape.y, b: firstShape.y + firstShape.height, r: firstShape.x + firstShape.width});
		this.updateLevelPositions(childs[0].algorithm);
		let previousShape = firstShape;
		let previousWidth = firstBounds.r - firstBounds.l;
		for (let i = 1; i < childs.length; i += 1) {
			const node = childs[i];
			const shape = node.getShape(isCalculateScaleCoefficient);
			const bounds = node.algorithm.getBounds(isCalculateScaleCoefficient);
			const childOffsets = this.getChildAlignOffsets(commonBounds, bounds, this.params[AscFormat.Param_type_chAlign]);
			const offX = previousShape.x + previousWidth + sibSp;
			shape.moveTo(childOffsets.offX + offX, childOffsets.offY);
			this.setLevelBounds({l: shape.x, t: shape.y, b: shape.y + shape.height, r: shape.x + shape.width});
			this.updateLevelPositions(node.algorithm);
			previousShape = shape;
			previousWidth = bounds.r - bounds.l;
		}
	};
	HierarchyChildAlgorithm.prototype.getShapeContainer = function (isCalculateScaleCoefficient) {
		if (isCalculateScaleCoefficient) {
			if (this.coefficientShapeContainer === null) {
				this.coefficientShapeContainer = new HierarchyChildContainer();
			}
			return this.coefficientShapeContainer;
		}
		if (this.shapeContainer === null) {
			this.shapeContainer = new HierarchyChildContainer();
		}
		return this.shapeContainer;
	}
	HierarchyChildAlgorithm.prototype.calculateShapePositionsHorizontalHang = function (isCalculateScaleCoefficient, fromRight) {
		const childs = this.getMainChilds();
		const sibSp = this.parentNode.getConstr(AscFormat.Constr_type_sibSp, !isCalculateScaleCoefficient);

		const topRow = [];
		const bottomRow = [];

		for (let i = 0; i < childs.length; i += 1) {
			const child = childs[i];
			if (i % 2 === 0) {
				topRow.push(child);
			} else {
				bottomRow.push(child);
			}
		}

		const topCalcBounds = this.applyMainChildAlign(topRow, isCalculateScaleCoefficient);
		const bottomCalcBounds = this.applyMainChildAlign(bottomRow, isCalculateScaleCoefficient);

		let bottomOffY = 0;
		if (topCalcBounds.commonBounds && bottomCalcBounds.commonBounds) {
			bottomOffY = topCalcBounds.commonBounds.b - bottomCalcBounds.commonBounds.t + sibSp;
		}

		let offX = 0;
		for (let i = 0; i < topRow.length; i += 1) {
			const topNode = topRow[i];
			const topShape = topNode.getShape(isCalculateScaleCoefficient);
			const topBounds = topCalcBounds.bounds[i];

			const commonChildBounds = Object.assign({}, topBounds);

			const bottomNode = bottomRow[i];
			let bottomShape;
			if (bottomNode) {
				bottomShape = bottomNode.getShape(isCalculateScaleCoefficient);
				const bottomBounds = bottomCalcBounds.bounds[i];
				checkBounds(commonChildBounds, bottomBounds);
				const bottomOffsets = this.getChildAlignOffsets(commonChildBounds, bottomBounds, this.params[AscFormat.Param_type_chAlign]);
				bottomNode.moveTo(bottomOffsets.offX, bottomOffsets.offY, isCalculateScaleCoefficient);
				if (fromRight) {
					bottomShape.moveTo(offX - (bottomShape.x + bottomShape.width), bottomOffY - bottomShape.y);
				} else {
					bottomShape.moveTo(offX - bottomShape.x, bottomOffY - bottomShape.y);
				}
				this.updateLevelPositions(bottomNode.algorithm);
				this.setLevelBounds({
					l: bottomShape.x,
					r: bottomShape.x + bottomShape.width,
					t: bottomShape.y,
					b: bottomShape.y + bottomShape.height
				});
			}
			const topOffsets = this.getChildAlignOffsets(commonChildBounds, topBounds, this.params[AscFormat.Param_type_chAlign]);
			topNode.moveTo(topOffsets.offX, topOffsets.offY, isCalculateScaleCoefficient);
			if (fromRight) {
				topShape.moveTo(offX - (topShape.x + topShape.width), 0);
			} else {
				topShape.moveTo(offX - topShape.x, 0);
			}
			this.updateLevelPositions(topNode.algorithm);
			this.setLevelBounds({
				l: topShape.x,
				r: topShape.x + topShape.width,
				t: topShape.y,
				b: topShape.y + topShape.height
			});
			if (topShape && bottomShape) {
				const bottomBounds = bottomCalcBounds.bounds[i];
				if (fromRight) {
					const topLeft = topShape.x + topShape.width - (topBounds.r - topBounds.l);
					const bottomLeft = bottomShape.x + bottomShape.width - (bottomBounds.r - bottomBounds.l);
					offX = Math.min(topLeft, bottomLeft) - sibSp;
				} else {
					const topRight = topShape.x + (topBounds.r - topBounds.l);
					const bottomRight = bottomShape.x + (bottomBounds.r - bottomBounds.l);
					offX = Math.max(topRight, bottomRight) + sibSp;
				}
			}
		}
	};
	HierarchyChildAlgorithm.prototype.getChildAlignOffsets = function (commonBounds, shapeBounds, childAlign) {
		const offsets = {offX: 0, offY: 0};

		switch (childAlign) {
			case AscFormat.ParameterVal_childAlignment_l:
				offsets.offX = commonBounds.l - shapeBounds.l;
				break;
			case AscFormat.ParameterVal_childAlignment_t:
				offsets.offY = commonBounds.t - shapeBounds.t;
				break;
			case AscFormat.ParameterVal_childAlignment_r:
				offsets.offX = commonBounds.r - shapeBounds.r;
				break;
			case AscFormat.ParameterVal_childAlignment_b:
				offsets.offY = commonBounds.b - shapeBounds.b;
				break;
			default:
				break;
		}
		return offsets;
	};
	HierarchyChildAlgorithm.prototype.calculateShapePositionsVerticalHangFromTop = function (isCalculateScaleCoefficient, fromRight) {
		const childs = this.getMainChilds();

		let leftCol = [];
		let rightCol = [];
		const sibSp = this.parentNode.getConstr(AscFormat.Constr_type_sibSp, !isCalculateScaleCoefficient);
		for (let i = 0; i < childs.length; i += 1) {
			const child = childs[i];
			if (i % 2 === 0) {
				leftCol.push(child);
			} else {
				rightCol.push(child);
			}
		}
		if (fromRight) {
			const tempCol = leftCol;
			leftCol = rightCol;
			rightCol = tempCol;
		}
		const leftCalcBounds = this.applyMainChildAlign(leftCol, isCalculateScaleCoefficient);
		const rightCalcBounds = this.applyMainChildAlign(rightCol, isCalculateScaleCoefficient);
		let rightOffX = 0;
		if (leftCalcBounds.commonBounds && rightCalcBounds.commonBounds) {
			rightOffX = leftCalcBounds.commonBounds.r - rightCalcBounds.commonBounds.l + sibSp;
		}
		let offY = 0;
		const columnLength = Math.ceil(childs.length / 2);
		for (let i = 0; i < columnLength; i += 1) {
			const leftNode = leftCol[i];
			if (leftNode) {
				const leftShape = leftNode.getShape(isCalculateScaleCoefficient);
				leftShape.moveTo(0, offY);
				this.updateLevelPositions(leftNode.algorithm);
				this.setLevelBounds({l: leftShape.x, r: leftShape.x + leftShape.width, t: leftShape.y, b: leftShape.y + leftShape.height});
			}
			const rightNode = rightCol[i];
			if (rightNode) {
				const rightShape = rightNode.getShape(isCalculateScaleCoefficient);
				rightShape.moveTo(rightOffX, offY);
				this.updateLevelPositions(rightNode.algorithm);
				this.setLevelBounds({l: rightShape.x, r: rightShape.x + rightShape.width, t: rightShape.y, b: rightShape.y + rightShape.height});
			}
			const leftBounds = leftCalcBounds.bounds[i];
			const rightBounds = rightCalcBounds.bounds[i];
			if (leftBounds && rightBounds) {
				const commonChildBounds = Object.assign({}, leftBounds);
				checkBounds(commonChildBounds, rightBounds);
/*				const leftOffsets = this.getChildAlignOffsets(commonChildBounds, leftBounds, this.params[AscFormat.Param_type_chAlign]);
				const rightOffsets = this.getChildAlignOffsets(commonChildBounds, rightBounds, this.params[AscFormat.Param_type_chAlign]);
				leftNode.moveTo(leftOffsets.offX, leftOffsets.offY, isCalculateScaleCoefficient);
				rightNode.moveTo(rightOffsets.offX, rightOffsets.offY, isCalculateScaleCoefficient);*/
				offY += Math.max(leftBounds.b - leftBounds.t, rightBounds.b - rightBounds.t) + sibSp;
			}
		}
	};
	HierarchyChildAlgorithm.prototype.applyMainChildAlign = function (childs, isCalculateScaleCoefficient) {
		const allBounds = [];
		const result = {
			commonBounds: null,
			bounds: allBounds
		};
		if (!childs.length) {
			return result;
		}
		const commonBounds = {};
		const firstChild = childs[0];
		const firstBounds = firstChild.algorithm.getBounds(isCalculateScaleCoefficient);
		allBounds.push(firstBounds);
		Object.assign(commonBounds, firstBounds);
		for (let i = 1; i < childs.length; i += 1) {
			const child = childs[i];
			const bounds = child.algorithm.getBounds(isCalculateScaleCoefficient);
			checkBounds(commonBounds, bounds);
			allBounds.push(bounds);
		}

		for (let i = 0; i < childs.length; i += 1) {
			const child = childs[i];
			const mainOffsets = this.getChildAlignOffsets(commonBounds, allBounds[i], this.params[AscFormat.Param_type_chAlign]);
			const secondaryOffsets = this.getChildAlignOffsets(commonBounds, allBounds[i], this.params[AscFormat.Param_type_secChAlign]);
			child.moveTo(mainOffsets.offX + secondaryOffsets.offX, mainOffsets.offY + secondaryOffsets.offY, isCalculateScaleCoefficient);
		}
		return {commonBounds: commonBounds, bounds: allBounds};
	};

	HierarchyChildAlgorithm.prototype.calculateShapePositionsVerticalHangFromRight = function (isCalculateScaleCoefficient) {

	};
	HierarchyChildAlgorithm.prototype._calculateShapePositions = function (isCalculateScaleCoefficient) {
		const childs = this.getMainChilds();
		if (!childs.length) {
			return;
		}
		if (this.isHang()) {
			switch (this.params[AscFormat.Param_type_secLinDir]) {
				case AscFormat.ParameterVal_linearDirection_fromT:
					const fromRight = this.params[AscFormat.Param_type_linDir] === AscFormat.ParameterVal_linearDirection_fromR;
					this.calculateShapePositionsVerticalHangFromTop(isCalculateScaleCoefficient, fromRight);
					break;
				case AscFormat.ParameterVal_linearDirection_fromL:
					this.calculateShapePositionsHorizontalHang(isCalculateScaleCoefficient, false);
					break;
				case AscFormat.ParameterVal_linearDirection_fromR:
					this.calculateShapePositionsHorizontalHang(isCalculateScaleCoefficient, true);
					break;
				default:
					break;
			}
		} else {
			if (this.isHorizontalHierarchy()) {
				switch (this.params[AscFormat.Param_type_linDir]) {
					case AscFormat.ParameterVal_linearDirection_fromB:
						this.calculateHorizontalHierarchyVerticalShapePositions(isCalculateScaleCoefficient, false);
						break;
					case AscFormat.ParameterVal_linearDirection_fromT:
						this.calculateHorizontalHierarchyVerticalShapePositions(isCalculateScaleCoefficient, true);
						break;
					case AscFormat.ParameterVal_linearDirection_fromR:
						this.calculateHorizontalHierarchyHorizontalShapePositions(isCalculateScaleCoefficient, false);
						break;
					case AscFormat.ParameterVal_linearDirection_fromL:
						this.calculateHorizontalHierarchyHorizontalShapePositions(isCalculateScaleCoefficient, true);
						break;
					default:
						break;
				}
			} else {
				switch (this.params[AscFormat.Param_type_linDir]) {
					case AscFormat.ParameterVal_linearDirection_fromB:
						this.calculateVerticalHierarchyVerticalShapePositions(isCalculateScaleCoefficient, false);
						break;
					case AscFormat.ParameterVal_linearDirection_fromT:
						this.calculateVerticalHierarchyVerticalShapePositions(isCalculateScaleCoefficient, true);
						break;
					case AscFormat.ParameterVal_linearDirection_fromR:
						this.calculateVerticalHierarchyHorizontalShapePositions(isCalculateScaleCoefficient, false);
						break;
					case AscFormat.ParameterVal_linearDirection_fromL:
						this.calculateVerticalHierarchyHorizontalShapePositions(isCalculateScaleCoefficient, true);
						break;
					default:
						break;
				}
			}
		}
	};

	HierarchyChildAlgorithm.prototype.getChildBounds = function (isCalculateScaleCoefficients) {
		const childs = this.getMainChilds();

		if (!childs.length) {
			return;
		}

		const firstChild = childs[0];
		const firstRoot = firstChild.algorithm.getRoot();
		const firstShape = firstRoot.getShape(isCalculateScaleCoefficients);
		const bounds = firstShape.getBounds();

		for (let i = 1; i < childs.length; i += 1) {
			const child = childs[i];
			const root = child.algorithm.getRoot();
			const shape = root.getShape(isCalculateScaleCoefficients);
			shape.checkBounds(bounds);
		}
		return bounds;
	};
	HierarchyChildAlgorithm.prototype.getHangStructBounds = function (isCalculateScaleCoefficients) {
		const childs = this.getMainChilds();
		if (!childs.length) {
			return;
		}
		const firstChild = childs[0];
		const bounds = firstChild.algorithm.getBounds(isCalculateScaleCoefficients);
		for (let i = 2; i < childs.length; i += 2) {
			const child = childs[i];
			child.algorithm.getBounds(isCalculateScaleCoefficients, bounds);
		}
		return bounds;
	};
	function HierarchyRootAlgorithm() {
		HierarchyAlgorithm.call(this);
	}
	AscFormat.InitClassWithoutType(HierarchyRootAlgorithm, HierarchyAlgorithm);
	HierarchyRootAlgorithm.prototype.isHorizontalHierarchy = function () {
		switch (this.params[AscFormat.Param_type_hierAlign]) {
			case AscFormat.ParameterVal_hierarchyAlignment_lB:
			case AscFormat.ParameterVal_hierarchyAlignment_lCtrCh:
			case AscFormat.ParameterVal_hierarchyAlignment_lCtrDes:
			case AscFormat.ParameterVal_hierarchyAlignment_lT:
			case AscFormat.ParameterVal_hierarchyAlignment_rB:
			case AscFormat.ParameterVal_hierarchyAlignment_rCtrCh:
			case AscFormat.ParameterVal_hierarchyAlignment_rCtrDes:
			case AscFormat.ParameterVal_hierarchyAlignment_rT:
				return true;
			default:
				return false;
		}
	};
	HierarchyRootAlgorithm.prototype.putShapesToShapeContainer = function (isCalculateScaleCoefficient) {
		const rootNode = this.getRoot();
		const asstNode = this.getAsstNode();
		const nonAsstNode = this.getNonAsstNode();
		const shapeContainer = this.getShapeContainer(isCalculateScaleCoefficient);
		shapeContainer.push(rootNode.getShape(isCalculateScaleCoefficient));
		if (asstNode && asstNode.algorithm.getMainChilds().length) {
			shapeContainer.push(asstNode.getShape(isCalculateScaleCoefficient));
		}
		if (nonAsstNode.algorithm.getMainChilds().length) {
			shapeContainer.push(nonAsstNode.getShape(isCalculateScaleCoefficient));
		}
	};
	HierarchyRootAlgorithm.prototype.getBounds = function (isCalculateScaleCoefficients, bounds) {
		const childs = this.getMainChilds();
		const firstChild = childs[0];
		const firstShape = firstChild.getShape(isCalculateScaleCoefficients);
		if (!bounds) {
			bounds = firstShape.getBounds();
		} else {
			firstShape.checkBounds(bounds);
		}
		const asstNode = this.getAsstNode();
		if (asstNode && asstNode.childs.length) {
			const shape = asstNode.getShape(isCalculateScaleCoefficients);
			shape.checkBounds(bounds);
		}
		const nonAsstNode = this.getNonAsstNode();
		if (nonAsstNode.childs.length) {
			const shape = nonAsstNode.getShape(isCalculateScaleCoefficients);
			shape.checkBounds(bounds);
		}
		return bounds;
	};
	HierarchyRootAlgorithm.prototype.createShadowShape = function (isCalculateScaleCoefficients) {
		return this.parentNode.createHierarchyRootShadowShape(isCalculateScaleCoefficients);
	};
	HierarchyRootAlgorithm.prototype.getNonAsstNode = function () {
		return this.parentNode.childs[1];
	};
	HierarchyRootAlgorithm.prototype.getAsstNode = function () {
		return this.parentNode.childs[2];
	};
	HierarchyRootAlgorithm.prototype.applyMainTopAlign = function (isCalculateScaleCoefficients) {
		const root = this.getRoot();
		if (!root) {
			return;
		}
		const parentNode = this.parentNode;
		const space = parentNode.getConstr(AscFormat.Constr_type_sp, !isCalculateScaleCoefficients);
		const rootShape = root.getShape(isCalculateScaleCoefficients);
		let offY = rootShape.y + rootShape.height + space;

		const asstNode = this.getAsstNode();
		if (asstNode && asstNode.algorithm.getMainChilds().length) {
			const asstShape = asstNode.getShape(isCalculateScaleCoefficients);
			asstShape.moveTo(/*todo: asstOffset*/0, offY - asstShape.y);
			offY = asstShape.y + asstShape.height + space;
		}

		const nonAsstNode = this.getNonAsstNode();
		if (nonAsstNode.algorithm.getMainChilds().length) {
			const nonAsstShape = nonAsstNode.getShape(isCalculateScaleCoefficients);
			nonAsstShape.moveTo(/*todo: nonAsstOffset*/0, offY - nonAsstShape.y);
		}
	};

	HierarchyRootAlgorithm.prototype.applyMainBottomAlign = function (isCalculateScaleCoefficients) {
		const rootNode = this.getRoot();
		if (!rootNode) {
			return;
		}
		const parentNode = this.parentNode;
		const space = parentNode.getConstr(AscFormat.Constr_type_sp, !isCalculateScaleCoefficients);

		let offY = 0;
		const nonAsstNode = this.getNonAsstNode();
		if (nonAsstNode.algorithm.getMainChilds().length) {
			const nonAsstShape = nonAsstNode.getShape(isCalculateScaleCoefficients);
			offY += nonAsstShape.y + nonAsstShape.height + space;
		}


		const asstNode = this.getAsstNode();
		if (asstNode && asstNode.algorithm.getMainChilds().length) {
			const asstShape = asstNode.getShape(isCalculateScaleCoefficients);
			asstShape.moveTo(/*todo: asstOffset*/0, offY - asstShape.y);
			offY = asstShape.y + asstShape.height + space;
		}

		const rootShape = rootNode.getShape(isCalculateScaleCoefficients);
		rootShape.moveTo(0, offY - rootShape.y);
	};

	HierarchyRootAlgorithm.prototype.applyMainLeftAlign = function (isCalculateScaleCoefficients) {
		const rootNode = this.getRoot();
		if (!rootNode) {
			return;
		}
		const space = this.parentNode.getConstr(AscFormat.Constr_type_sp, !isCalculateScaleCoefficients);
		const rootShape = rootNode.getShape(isCalculateScaleCoefficients);
		const bounds = rootShape.getBounds();
		let offX = bounds.r + space;

		const asstNode = this.getAsstNode();
		if (asstNode && asstNode.algorithm.getMainChilds().length) {
			const asstShape = asstNode.getShape(isCalculateScaleCoefficients);
			asstShape.moveTo(offX - asstShape.x, 0);
			offX = asstShape.x + asstShape.width + space;
		}

		const nonAsstNode = this.getNonAsstNode();
		if (nonAsstNode.algorithm.getMainChilds().length) {
			const nonAsstShape = nonAsstNode.getShape(isCalculateScaleCoefficients);
			nonAsstShape.moveTo(offX - nonAsstShape.x, 0);
		}
	};

	HierarchyRootAlgorithm.prototype.applyMainRightAlign = function (isCalculateScaleCoefficients) {
		const rootNode = this.getRoot();
		if (!rootNode) {
			return;
		}
		const space = this.parentNode.getConstr(AscFormat.Constr_type_sp, !isCalculateScaleCoefficients);
		let offX = 0;

		const nonAsstNode = this.getNonAsstNode();
		if (nonAsstNode.algorithm.getMainChilds().length) {
			const nonAsstShape = nonAsstNode.getShape(isCalculateScaleCoefficients);
			offX = nonAsstShape.x + nonAsstShape.width + space;
		}

		const asstNode = this.getAsstNode();
		if (asstNode && asstNode.algorithm.getMainChilds().length) {
			const asstShape = asstNode.getShape(isCalculateScaleCoefficients);
			asstShape.moveTo(offX - asstShape.x, 0);
			offX = asstShape.x + asstShape.width + space;
		}

		const rootShape = rootNode.getShape(isCalculateScaleCoefficients);
		const bounds = rootShape.getBounds();
		rootShape.moveTo(offX - bounds.l, 0);
	};

	HierarchyRootAlgorithm.prototype.applyMainAlign = function (isCalculateScaleCoefficients) {
		switch (this.params[AscFormat.Param_type_hierAlign]) {
			case AscFormat.ParameterVal_hierarchyAlignment_bCtrCh:
			case AscFormat.ParameterVal_hierarchyAlignment_bCtrDes:
			case AscFormat.ParameterVal_hierarchyAlignment_bL:
			case AscFormat.ParameterVal_hierarchyAlignment_bR:
				this.applyMainBottomAlign(isCalculateScaleCoefficients);
				break;
			case AscFormat.ParameterVal_hierarchyAlignment_lB:
			case AscFormat.ParameterVal_hierarchyAlignment_lT:
			case AscFormat.ParameterVal_hierarchyAlignment_lCtrCh:
			case AscFormat.ParameterVal_hierarchyAlignment_lCtrDes:
				this.applyMainLeftAlign(isCalculateScaleCoefficients);
				break;
			case AscFormat.ParameterVal_hierarchyAlignment_rB:
			case AscFormat.ParameterVal_hierarchyAlignment_rT:
			case AscFormat.ParameterVal_hierarchyAlignment_rCtrCh:
			case AscFormat.ParameterVal_hierarchyAlignment_rCtrDes:
				this.applyMainRightAlign(isCalculateScaleCoefficients);
				break;
			case AscFormat.ParameterVal_hierarchyAlignment_tCtrCh:
			case AscFormat.ParameterVal_hierarchyAlignment_tCtrDes:
			case AscFormat.ParameterVal_hierarchyAlignment_tL:
			case AscFormat.ParameterVal_hierarchyAlignment_tR:
				this.applyMainTopAlign(isCalculateScaleCoefficients);
				break;
			default:
				return;
		}
	};
	HierarchyRootAlgorithm.prototype.applySecondaryAlign = function (isCalculateScaleCoefficients) {
		const asstNode = this.getAsstNode();
		const nonAsstNode = this.getNonAsstNode();

		if (asstNode) {
			this.applySecondaryAlignForChild(isCalculateScaleCoefficients, asstNode);
		}
		this.applySecondaryAlignForChild(isCalculateScaleCoefficients, nonAsstNode);
	};
	HierarchyRootAlgorithm.prototype.applySecondaryAlignForChild = function (isCalculateScaleCoefficients, child) {
		const rootNode = this.getRoot();
		const rootShape = rootNode.getShape(isCalculateScaleCoefficients);
		const childShape = child.getShape(isCalculateScaleCoefficients);
		if (child.algorithm.isHang()) {
			const offsets = child.algorithm.getRootCenteringOffset(rootShape, isCalculateScaleCoefficients);
			childShape.moveTo(offsets.offX, offsets.offY);
			return;
		}

		const appendChildOffset = this.getOffsetAlign(isCalculateScaleCoefficients, child);
		switch (this.params[AscFormat.Param_type_hierAlign]) {
			case AscFormat.ParameterVal_hierarchyAlignment_lCtrCh:
			case AscFormat.ParameterVal_hierarchyAlignment_rCtrCh: {
				const childBounds = child.algorithm.getChildBounds(isCalculateScaleCoefficients);
				if (childBounds) {
					childShape.moveTo(0, (rootShape.y + rootShape.height / 2) - (childBounds.t + (childBounds.b - childBounds.t) / 2) + appendChildOffset);
				}
				break;
			}
			case AscFormat.ParameterVal_hierarchyAlignment_bCtrCh:
			case AscFormat.ParameterVal_hierarchyAlignment_tCtrCh: {
				const childBounds = child.algorithm.getChildBounds(isCalculateScaleCoefficients);
				if (childBounds) {
					childShape.moveTo((rootShape.x + rootShape.width / 2) - (childBounds.l + (childBounds.r - childBounds.l) / 2) + appendChildOffset, 0);
				}
				break;
			}
			case AscFormat.ParameterVal_hierarchyAlignment_lCtrDes:
			case AscFormat.ParameterVal_hierarchyAlignment_rCtrDes:
				childShape.moveTo(0, rootShape.y + rootShape.height / 2 - (childShape.y + childShape.height / 2) + appendChildOffset);
				break;
			case AscFormat.ParameterVal_hierarchyAlignment_bCtrDes:
			case AscFormat.ParameterVal_hierarchyAlignment_tCtrDes:
				childShape.moveTo(rootShape.x + rootShape.width / 2 - (childShape.x + childShape.width / 2) + appendChildOffset, 0);
				break;
			case AscFormat.ParameterVal_hierarchyAlignment_bL:
			case AscFormat.ParameterVal_hierarchyAlignment_tL:
				childShape.moveTo(rootShape.x - childShape.x + appendChildOffset, 0);
				break;
			case AscFormat.ParameterVal_hierarchyAlignment_bR:
			case AscFormat.ParameterVal_hierarchyAlignment_tR:
				childShape.moveTo((rootShape.x + rootShape.width) - (childShape.x + childShape.width) + appendChildOffset, 0);
				break;
			case AscFormat.ParameterVal_hierarchyAlignment_lB:
			case AscFormat.ParameterVal_hierarchyAlignment_rB:
				childShape.moveTo(0, (rootShape.y + rootShape.height) - (childShape.y + childShape.height) + appendChildOffset);
				break;
			case AscFormat.ParameterVal_hierarchyAlignment_lT:
			case AscFormat.ParameterVal_hierarchyAlignment_rT:
				childShape.moveTo(0, rootShape.y - childShape.y + appendChildOffset);
				break;
			default:
				return;
		}
	};

	HierarchyRootAlgorithm.prototype.getOffsetAlign = function (isCalculateScaleCoefficients, child) {
		const offsetFactor = child.algorithm.isHang() ? 0 : this.parentNode.getConstr(AscFormat.Constr_type_alignOff, !isCalculateScaleCoefficients);
		if (!offsetFactor) {
			return 0;
		}
		const rootNode = this.getRoot();
		const rootShape = rootNode.getShape(isCalculateScaleCoefficients);
		switch (this.params[AscFormat.Param_type_hierAlign]) {
			case AscFormat.ParameterVal_hierarchyAlignment_bCtrCh:
			case AscFormat.ParameterVal_hierarchyAlignment_lCtrCh:
			case AscFormat.ParameterVal_hierarchyAlignment_rCtrCh:
			case AscFormat.ParameterVal_hierarchyAlignment_tCtrCh:
			case AscFormat.ParameterVal_hierarchyAlignment_bCtrDes:
			case AscFormat.ParameterVal_hierarchyAlignment_lCtrDes:
			case AscFormat.ParameterVal_hierarchyAlignment_rCtrDes:
			case AscFormat.ParameterVal_hierarchyAlignment_tCtrDes:
				return 0;
			case AscFormat.ParameterVal_hierarchyAlignment_bL:
			case AscFormat.ParameterVal_hierarchyAlignment_tL:
				return rootShape.width * offsetFactor;
			case AscFormat.ParameterVal_hierarchyAlignment_bR:
			case AscFormat.ParameterVal_hierarchyAlignment_tR:
				return -rootShape.width * offsetFactor;
			case AscFormat.ParameterVal_hierarchyAlignment_lB:
			case AscFormat.ParameterVal_hierarchyAlignment_rB:
				return -rootShape.height * offsetFactor;
			case AscFormat.ParameterVal_hierarchyAlignment_lT:
			case AscFormat.ParameterVal_hierarchyAlignment_rT:
				return rootShape.height * offsetFactor;
			default:
				return 0;
		}
	};

	HierarchyRootAlgorithm.prototype.initParams = function (params) {
		HierarchyAlgorithm.prototype.initParams.call(this, params);
		if (this.params[AscFormat.Param_type_off] === undefined) {
			this.params[AscFormat.Param_type_off] = AscFormat.ParameterVal_offset_ctr;
		}
		if (this.params[AscFormat.Param_type_hierAlign] === undefined) {
			this.params[AscFormat.Param_type_hierAlign] = AscFormat.ParameterVal_hierarchyAlignment_tCtrCh;
		}
	}
	HierarchyRootAlgorithm.prototype.getRoot = function () {
		const childs = this.getMainChilds();
		if (childs.length) {
			return childs[0];
		}
	};
	HierarchyRootAlgorithm.prototype.getShapeContainer = function (isCalculateScaleCoefficients) {
		if (isCalculateScaleCoefficients) {
			if (this.coefficientShapeContainer === null) {
				this.coefficientShapeContainer = new HierarchyRootContainer();
			}
			return this.coefficientShapeContainer;
		} else {
			if (this.shapeContainer === null) {
				this.shapeContainer = new HierarchyRootContainer();
			}
			return this.shapeContainer;
		}
	};
	HierarchyRootAlgorithm.prototype._collectHierarchyPositions = function (isVertical) {
		const asstNode = this.getAsstNode();
		if (asstNode) {
			const asstAlgorithm = asstNode.algorithm;
			for (let i = 0; i < asstAlgorithm.levelPositions.length; i += 1) {
				this.setLevelBounds(asstAlgorithm.levelPositions[i]);
			}
		}
		const nonAsstNode = this.getNonAsstNode();
		const nonAsstAlgorithm = nonAsstNode.algorithm;
		for (let i = 0; i < nonAsstAlgorithm.levelPositions.length; i += 1) {
			this.setLevelBounds(nonAsstAlgorithm.levelPositions[i]);
		}
	};
	HierarchyRootAlgorithm.prototype.collectHierarchyPositions = function () {
		this._collectHierarchyPositions();
	};
	HierarchyRootAlgorithm.prototype._calculateShapePositions = function (isCalculateScaleCoefficients) {
		this.applyMainAlign(isCalculateScaleCoefficients);
		this.applySecondaryAlign(isCalculateScaleCoefficients);
	};


	function PyramidAlgorithm() {
		PositionAlgorithm.call(this);
		this.calcValues = {
			defaultBlockHeight: 0
		};
	}
	AscFormat.InitClassWithoutType(PyramidAlgorithm, PositionAlgorithm);
	PyramidAlgorithm.prototype.initParams = function (params) {
		PositionAlgorithm.prototype.initParams.call(this, params);
		if (this.params[AscFormat.Param_type_pyraAcctPos] === undefined) {
			this.params[AscFormat.Param_type_pyraAcctPos] = AscFormat.ParameterVal_pyramidAccentPosition_aft;
		}
		if (this.params[AscFormat.Param_type_off] === undefined) {
			this.params[AscFormat.Param_type_off] = AscFormat.ParameterVal_offset_ctr;
		}

	}
	PyramidAlgorithm.prototype.calcScaleCoefficients = function () {
		const parentHeight = this.parentNode.getConstr(AscFormat.Constr_type_h, false);
		const childs = this.parentNode.childs;
		const defaultBlockHeight = parentHeight / childs.length;
		let sumHeight = 0;
		for (let i = 0; i < childs.length; i++) {
			const child = this.getPyramidChildren(childs[i]).pyramid;
			const scaleBlockHeight = child.getHeightScale();
			sumHeight += scaleBlockHeight * defaultBlockHeight;
		}
		this.calcValues.defaultBlockHeight = defaultBlockHeight * (parentHeight / sumHeight);
	};
	PyramidAlgorithm.prototype.getTextNodeCoefficients = function (contentNode, textNode) {
		const resultScales = {width: 1, height: 1};
		const contentWidth = contentNode.getConstr(AscFormat.Constr_type_w, true, true);
		const contentHeight = contentNode.getConstr(AscFormat.Constr_type_h, true, true);
		const textWidth = textNode.getConstr(AscFormat.Constr_type_w, true, true);
		const textHeight = textNode.getConstr(AscFormat.Constr_type_h, true, true);
		if (contentWidth && textWidth) {
			resultScales.width = textWidth / contentWidth;
		}
		if (contentHeight && textHeight) {
			resultScales.height = textHeight / contentHeight;
		}
		return resultScales;
	};
	PyramidAlgorithm.prototype.setPyramidParametersForNode = function (child, x, y, height, width, cleanHeight, cleanWidth, adjValue) {
		const shape = child.getShape(false);
		if (shape) {
			if (width < height) {
				adjValue = adjValue * height / cleanWidth;
			}
			shape.height = height;
			shape.width = width;
			shape.x = x;
			shape.y = y;
			shape.cleanParams.height = cleanHeight;
			shape.cleanParams.width = cleanWidth;
			const adjLst = new AscFormat.AdjLst();
			const adj = new AscFormat.Adj();
			adj.setVal(adjValue);
			adj.setIdx(1);
			adjLst.addToLst(0, adj);
			shape.customAdj = adjLst;

			const textNodes = child.contentNodes[0] && child.contentNodes[0].getTextNodes();
			if (textNodes && textNodes.contentNode === child && textNodes.textNode) {
				const txNode = textNodes.textNode;
				const coefficients = this.getTextNodeCoefficients(child, txNode);
				const textShape = txNode.getShape();
				textShape.height = shape.height * coefficients.height;
				textShape.width = shape.width * coefficients.width;
				textShape.x = shape.x + (shape.width - textShape.width) / 2;
				textShape.y = shape.y + (shape.height - textShape.height) / 2;
			}
		}
	};
	PyramidAlgorithm.prototype.getPyramidChildren = function (node) {
		const pyramid = node.getNamedNode(this.params[AscFormat.Param_type_pyraLvlNode]);
		const acct = node.getNamedNode(this.params[AscFormat.Param_type_pyraAcctBkgdNode]);
		return {
			pyramid: pyramid,
			acct: acct
		};
	};
	PyramidAlgorithm.prototype.isReversedPyramid = function () {
		return this.params[AscFormat.Param_type_linDir] === AscFormat.ParameterVal_linearDirection_fromT;
	};
	PyramidAlgorithm.prototype.getStartDefaultBlockWidth = function () {
		const acctRatio = this.parentNode.getConstr(AscFormat.Constr_type_pyraAcctRatio, true);
		const parentWidth = this.parentNode.getConstr(AscFormat.Constr_type_w, true);
		return parentWidth * (1 - acctRatio);
	}
	PyramidAlgorithm.prototype.forEachChild = function (callback, startIndex) {
		const childs = this.parentNode.childs;
		if (this.isReversedPyramid()) {
			for (let i = startIndex; i < childs.length; i += 1) {
				callback(childs[i]);
			}
		} else {
			for (let i = childs.length - 1 - startIndex; i >= 0; i -= 1) {
				callback(childs[i]);
			}
		}
	}
	PyramidAlgorithm.prototype.getFirstPyramidComponents = function () {
		const childs = this.parentNode.childs;
		if (this.isReversedPyramid()) {
			return this.getPyramidChildren(childs[0]);
		}

		return this.getPyramidChildren(childs[childs.length - 1]);
	}
	PyramidAlgorithm.prototype.getShapeContainer = function (isCalculateScaleCoefficient) {
		if (isCalculateScaleCoefficient) {
			if (this.coefficientShapeContainer === null) {
				this.coefficientShapeContainer = new PyramidContainer();
			}
			return this.coefficientShapeContainer;
		}
		if (this.shapeContainer === null) {
			this.shapeContainer = new PyramidContainer();
		}
		return this.shapeContainer;
	};
	PyramidAlgorithm.prototype._calculateShapePositions = function () {
		const parentNode = this.parentNode;
		const shapeContainer = this.getShapeContainer(false);
		const parentHeight = parentNode.getConstr(AscFormat.Constr_type_h, true);
		const parentWidth = parentNode.getConstr(AscFormat.Constr_type_w, true);
		const defaultBlockHeight = this.calcValues.defaultBlockHeight;
		let previousBlockWidth = this.getStartDefaultBlockWidth();
		const defaultAdjValue = (previousBlockWidth / 2) / parentHeight;


		const firstPyramidComponents = this.getFirstPyramidComponents();
		const firstChild = firstPyramidComponents.pyramid;
		const firstShape = firstChild.getShape(false);
		const firstHeight = defaultBlockHeight * firstChild.getHeightScale();
		const firstWidth = previousBlockWidth * firstChild.getWidthScale();
		let ctrX = (previousBlockWidth - firstWidth) / 2 + firstWidth / 2;
		if (!this.isAfterAcct()) {
			ctrX = parentWidth - ctrX;
		}
		let previousY;
		if (this.isReversedPyramid()) {
			previousY = 0;
		} else {
			previousY = parentHeight - firstHeight;
		}
		let previousBlockHeight = firstHeight;
		this.setPyramidParametersForNode(firstChild, ctrX - firstWidth / 2, previousY,  firstHeight, firstWidth, defaultBlockHeight, previousBlockWidth, defaultAdjValue);
		const firstAcctOffset = defaultAdjValue * previousBlockHeight;
		if (this.isReversedPyramid()) {
			firstShape.rot = Math.PI;
			this.addAcctShape(firstChild, firstPyramidComponents.acct, defaultAdjValue, firstAcctOffset);
			shapeContainer.push(firstShape);
		} else {
			shapeContainer.push(firstShape);
			this.addAcctShape(firstChild, firstPyramidComponents.acct, defaultAdjValue, firstAcctOffset);
		}
		const oThis = this;
		this.forEachChild(function (node) {
			const pyramidComponents = oThis.getPyramidChildren(node);
			const child = pyramidComponents.pyramid;
			const shape = child.getShape(false);
			const blockHeightFactor = child.getHeightScale();
			const blockWidthFactor = child.getWidthScale();
			const scaledBlockHeight = blockHeightFactor * defaultBlockHeight;
			const curBlockWidth = previousBlockWidth - 2 * defaultAdjValue * previousBlockHeight;
			const scaledBlockWidth = curBlockWidth * blockWidthFactor;
			const x = ctrX - scaledBlockWidth / 2;
			let y;
			if (oThis.isReversedPyramid()) {
				y = previousY + previousBlockHeight;
			} else {
				y = previousY - scaledBlockHeight;
			}
			const acctOffset = defaultAdjValue * scaledBlockHeight;
			oThis.setPyramidParametersForNode(child, x, y,  scaledBlockHeight, scaledBlockWidth, defaultBlockHeight, curBlockWidth, defaultAdjValue);
			if (oThis.isReversedPyramid()) {
				shape.rot = Math.PI;
				oThis.addAcctShape(child, pyramidComponents.acct, defaultAdjValue, acctOffset);
				shapeContainer.push(shape);
			} else {
				shapeContainer.push(shape);
				oThis.addAcctShape(child, pyramidComponents.acct, defaultAdjValue, acctOffset);
			}
			previousBlockWidth = curBlockWidth;
			previousBlockHeight = scaledBlockHeight;
			previousY = y;
		}, 1);
		if (!this.isReversedPyramid()) {
			shapeContainer.reverse();
		}
	};
	PyramidAlgorithm.prototype.calculateShapePositions = function (smartartAlgorithm, isCalculateScaleCoefficients) {
		const parentNode = this.parentNode;
		const childs = parentNode.childs;
		if (!childs.length) {
			return;
		}
		if (isCalculateScaleCoefficients) {
			this.calcScaleCoefficients(smartartAlgorithm);
		} else {
			this._calculateShapePositions();
			this.applyParamOffsets();
			this.applyPostAlgorithmSettings();
			this.createShadowShape(false);
		}
	};
	PyramidAlgorithm.prototype.createShadowShape = function (isCalculateScaleCoefficients) {
		return this.parentNode.createShadowShape(true, isCalculateScaleCoefficients);
	};
	PyramidAlgorithm.prototype.getTemplateAdjAcctLst = function (firstAdj, secondAdj) {
		const adjLst = new AscFormat.AdjLst();
		const adj1 = new AscFormat.Adj();
		adjLst.addToLst(0, adj1);
		adj1.setIdx(1);

		const adj2 = new AscFormat.Adj();
		adj2.setIdx(2);
		adjLst.addToLst(1, adj2);

		adj1.setVal(firstAdj);
		adj2.setVal(secondAdj);
		return adjLst;
	};
	PyramidAlgorithm.prototype.isAfterAcct = function () {
		return this.params[AscFormat.Param_type_pyraAcctPos] === AscFormat.ParameterVal_pyramidAccentPosition_aft;
	}
	PyramidAlgorithm.prototype.addAcctShape = function (mainNode, acctNode, defaultAdjValue, acctHelper) {
		if (!(mainNode && acctNode)) {
			return;
		}
		const shapeContainer = this.getShapeContainer();
		const parentWidth = this.parentNode.getConstr(AscFormat.Constr_type_w, true);
		const mainShape = mainNode.getShape(false);
		const acctShape = acctNode.getShape(false);

		const heightScale = acctNode.getHeightScale();
		const widthScale = acctNode.getWidthScale();
		let defaultWidth;
		if (this.isAfterAcct()) {
			defaultWidth = parentWidth - (mainShape.x + mainShape.width - acctHelper);
		} else {
			defaultWidth = mainShape.x + acctHelper;
		}
		const defaultHeight = mainShape.height;
		acctShape.cleanParams.width = defaultWidth;
		acctShape.cleanParams.height = defaultHeight;
		acctShape.width = defaultWidth * widthScale;
		acctShape.height = defaultHeight * heightScale;
		if (defaultWidth < defaultHeight) {
			defaultAdjValue = defaultAdjValue * defaultHeight / defaultWidth;
		}

		if (this.isAfterAcct()) {
			acctShape.x = mainShape.x + mainShape.width - acctHelper + (defaultWidth - acctShape.width) / 2;
			if (this.isReversedPyramid()) {
				acctShape.customAdj = this.getTemplateAdjAcctLst(defaultAdjValue, 0);
			} else {
				acctShape.customAdj = this.getTemplateAdjAcctLst(0, defaultAdjValue);
			}
		} else {
			acctShape.x = (defaultWidth - acctShape.width) / 2;
			if (this.isReversedPyramid()) {
				acctShape.customAdj = this.getTemplateAdjAcctLst(0, defaultAdjValue);
			} else {
				acctShape.customAdj = this.getTemplateAdjAcctLst(defaultAdjValue, 0);
			}
		}
		acctShape.y = mainShape.y + (defaultHeight - acctShape.height) / 2;
		if (!this.isReversedPyramid()) {
			acctShape.rot = Math.PI;
		}
		shapeContainer.push(acctShape);

		const txNode = acctNode.contentNodes[0] && acctNode.contentNodes[0].getTextNode();
		if (txNode) {
			const textShape = txNode.getShape();
			textShape.type = AscFormat.LayoutShapeType_shapeType_rect;
			if (this.isAfterAcct()) {
				textShape.x = mainShape.x + mainShape.width;
				textShape.y = acctShape.y;
				textShape.height = acctShape.height;
				textShape.width = acctShape.width - (textShape.x - acctShape.x);
				textShape.rot = acctShape.rot;
			} else {
				textShape.x = acctShape.x;
				textShape.y = acctShape.y;
				textShape.height = acctShape.height;
				textShape.width = acctShape.width - acctHelper;
				textShape.rot = acctShape.rot;
			}

		}
	};

	function CycleAlgorithm() {
		PositionAlgorithm.call(this);
		this.calcValues = {
			radius: 0,
			startAngle: 0,
			stepAngle: 0,
			mainElements: [],
			centerNodeIndex: null,
			isInit: false
		};
	}
	AscFormat.InitClassWithoutType(CycleAlgorithm, PositionAlgorithm);
	CycleAlgorithm.prototype.getChildAlgorithmAlignBounds = function (isCalculateCoefficients, skipRotate) {
		const x = this.parentNode.getConstr(AscFormat.Constr_type_l, !isCalculateCoefficients);
		const y = this.parentNode.getConstr(AscFormat.Constr_type_t, !isCalculateCoefficients);
		const width = this.parentNode.getConstr(AscFormat.Constr_type_w, !isCalculateCoefficients);
		const height = this.parentNode.getConstr(AscFormat.Constr_type_h, !isCalculateCoefficients);
		return {
			l: x,
			t: y,
			r: x + width,
			b: y + height
		};
	};
	CycleAlgorithm.prototype.getCenterNode = function () {
		if (this.calcValues.centerNodeIndex !== null) {
			return this.parentNode.childs[this.calcValues.centerNodeIndex];
		}
	};
	CycleAlgorithm.prototype.getParentConnectionNode = function () {
		return this.getCenterNode();
	};
	CycleAlgorithm.prototype.isClockwise = function () {
		return this.calcValues.stepAngle > 0;
	}
	CycleAlgorithm.prototype.getShapeIndex = function (shape) {
		return this.calcValues.mainElements.indexOf(shape);
	};
	CycleAlgorithm.prototype.getRadialConnectionInfo = function (node) {
		const shape = node.getShape();
		const shapeCenterPoint = new CCoordPoint(shape.x + shape.width / 2, shape.y + shape.height / 2);
		let radius;
		let angle;
		let custVector;
		if (shape.custCenterPoint) {
			custVector = shape.custCenterPoint.getVector(shapeCenterPoint);
			radius = custVector.getDistance();
			angle = custVector.getAngle();
			return {
				point: shape.custCenterPoint,
				radius: radius,
				angle: angle,
				isClockwise: this.isClockwise()
			}
		} else if (shape.radialVector) {
				radius = shape.radialVector.getDistance();
				angle = shape.radialVector.getAngle();
				return {
					point: new CCoordPoint(shapeCenterPoint.x - shape.radialVector.x, shapeCenterPoint.y - shape.radialVector.y),
					radius: radius,
					angle: angle,
					isClockwise: this.isClockwise()
				};
		}
		return null;
	};
	CycleAlgorithm.prototype.initParams = function (params) {
		PositionAlgorithm.prototype.initParams.call(this, params);
		if (this.params[AscFormat.Param_type_stAng] === undefined) {
			this.params[AscFormat.Param_type_stAng] = 0;
		}
		if (this.params[AscFormat.Param_type_spanAng] === undefined) {
			this.params[AscFormat.Param_type_spanAng] = 360;
		}
		if (this.params[AscFormat.Param_type_vertAlign] === undefined) {
			this.params[AscFormat.Param_type_vertAlign] = AscFormat.ParameterVal_verticalAlignment_mid;
		}
		if (this.params[AscFormat.Param_type_horzAlign] === undefined) {
			this.params[AscFormat.Param_type_horzAlign] = AscFormat.ParameterVal_horizontalAlignment_ctr;
		}

		if (this.params[AscFormat.Param_type_ctrShpMap] === undefined) {
			this.params[AscFormat.Param_type_ctrShpMap] = AscFormat.ParameterVal_centerShapeMapping_none;
		}
	}
	CycleAlgorithm.prototype.getCenterShapeRadius = function (centerBounds, anotherBounds, guideVector) {
		if (!centerBounds || !anotherBounds) {
			return 0;
		}
		const centerPoint = getShapePoint(centerBounds);
		const anotherPoint = getShapePoint(anotherBounds);
		const centerEdgePoint = getMinShapeEdgePoint(centerBounds, guideVector);
		const anotherEdgePoint = getMinShapeEdgePoint(anotherBounds, new CVector(-guideVector.x, -guideVector.y));
		if (centerEdgePoint && anotherEdgePoint) {
			const centerDistance = centerPoint.getVector(centerEdgePoint).getDistance();
			const anotherDistance = anotherPoint.getVector(anotherEdgePoint).getDistance();
			const minPadding = this.parentNode.getConstr(AscFormat.Constr_type_sp);
			return centerDistance + anotherDistance + minPadding;
		}
		return 0;
	};
	CycleAlgorithm.prototype.getStartCycleBounds = function () {
		const centerNode = this.getCenterNode();
		if (centerNode) {
			const centerShape = centerNode.getShape(true);
			const bounds = centerShape.getBounds(true);
			const halfWidth = (bounds.r - bounds.l) / 2;
			const halfHeight = (bounds.b - bounds.t) / 2;
			return {
				l: -halfWidth,
				r: halfWidth,
				t: -halfHeight,
				b: halfHeight
			};
		}
		return {l: 0, r: 0, t: 0, b: 0};
	}
	CycleAlgorithm.prototype.initCenterShapeMap = function () {
		if (this.params[AscFormat.Param_type_ctrShpMap] === AscFormat.ParameterVal_centerShapeMapping_fNode) {
			const childs = this.parentNode.childs;
			for (let i = 0; i < childs.length; i += 1) {
				const child = childs[i];
				if (child.isContentNode()) {
					this.calcValues.centerNodeIndex = i;
					return i + 1;
				}
			}
		}
		return 0;
	};
	CycleAlgorithm.prototype.calculateStartValues = function () {
		if (!this.calcValues.isInit) {
			this.calcValues.isInit = true;
			const spanAngle = this.params[AscFormat.Param_type_spanAng];
			const startAngle = AscFormat.normalizeRotate(this.params[AscFormat.Param_type_stAng] * degToRad - Math.PI / 2);

			const mainElements = this.calcValues.mainElements;
			const childs = this.parentNode.childs;
			let startIndex = this.initCenterShapeMap();
			for (startIndex; startIndex < childs.length; startIndex += 1) {
				const child = childs[startIndex];
				if (child.isContentNode()) {
					mainElements.push(child);
				}
			}

			let stepAngle;
			if (Math.abs(spanAngle) === 360) {
				if (mainElements.length === 0) {
					stepAngle = 0;
				} else {
					stepAngle = (spanAngle / mainElements.length) * degToRad;
				}

			} else {
				if (mainElements.length === 0) {
					stepAngle = 0;
				} else {
					stepAngle = (spanAngle / (mainElements.length - 1)) * degToRad;
				}
			}
			this.calcValues.startAngle = startAngle;
			this.calcValues.stepAngle  = stepAngle;
		}
	};
	CycleAlgorithm.prototype.getNormalizeSibSp = function () {
		//todo think about how it is actually calculated
		const stepAngle = this.calcValues.stepAngle;
		const sibSp = this.parentNode.getConstr(AscFormat.Constr_type_sibSp);
		if (this.isClockwise()) {
			if (stepAngle > Math.PI / 2 && sibSp < 0) {
				return 0;
			}
		} else {
			if (stepAngle < -Math.PI / 2 && sibSp < 0) {
				return 0;
			}
		}
		return sibSp;
	}
	CycleAlgorithm.prototype.calcScaleCoefficients = function (smartartAlgorithm) {
		const startAngle = this.calcValues.startAngle;
		const stepAngle = this.calcValues.stepAngle;
		const mainElements = this.calcValues.mainElements;
		let previousAngle = startAngle;
		let currentAngle = AscFormat.normalizeRotate(startAngle + stepAngle);
		const divider = Math.sqrt(2 * (1 - Math.cos(Math.abs(stepAngle))));
		const sibSp = this.getNormalizeSibSp();

		const firstShape = mainElements[0] && mainElements[0].getShape(true)
		const firstElementBounds = firstShape && firstShape.getBounds(true);
		let maxRadius = 0;
		const centerNode = this.getCenterNode();
		let centerShapeBounds;
		if (centerNode) {
			const startGuideVector = CVector.getVectorByAngle(startAngle);
			const centerShape = centerNode.getShape(true);
			centerShapeBounds = centerShape.getBounds(true);
			maxRadius = this.getCenterShapeRadius(centerShapeBounds, firstElementBounds, startGuideVector);
		}
		if (divider !== 0) {
			let previousBounds = firstElementBounds;
			for (let i = 1; i < mainElements.length + 1; i++) {
				const curIndex = i === mainElements.length ? 0 : i;
				const shape = mainElements[curIndex].getShape(true);
				const currentBounds = shape.getBounds(true);
				const centerGuideVector = CVector.getVectorByAngle(currentAngle);
				const tempCenterRadius = this.getCenterShapeRadius(centerShapeBounds, currentBounds, centerGuideVector);

				let tempSibRadius = 0;

				const previousVector = CVector.getVectorByAngle(previousAngle);
				const currentVector = CVector.getVectorByAngle(currentAngle);
				const guideVector = currentVector.getDiffVector(previousVector);
				const currentEdgePoint = getMinShapeEdgePoint(currentBounds, guideVector);
				const previousEdgePoint = getMinShapeEdgePoint(previousBounds, new CVector(-guideVector.x, -guideVector.y));
				if (currentEdgePoint && previousEdgePoint) {
					const currentShapePoint = getShapePoint(currentBounds);
					const previousShapePoint = getShapePoint(previousBounds);
					const currentVector = currentShapePoint.getVector(currentEdgePoint);
					const previousVector = previousShapePoint.getVector(previousEdgePoint);
					const previousDistance = previousVector.getDistance();
					const currentDistance = currentVector.getDistance();

					tempSibRadius = (sibSp + previousDistance + currentDistance) / divider;
				}
				maxRadius = Math.max(maxRadius, tempSibRadius, tempCenterRadius);
				previousAngle = currentAngle;
				currentAngle = i === mainElements.length - 1 ? AscFormat.normalizeRotate(startAngle) : AscFormat.normalizeRotate(currentAngle + stepAngle);
				previousBounds = currentBounds;
			}
		}
		currentAngle = startAngle;
		const cycleBounds = this.getStartCycleBounds();
		const radiusBounds = {l: 0, r: 0, t: 0, b: 0};
		for (let i = 0; i < mainElements.length; i++) {
			const radiusVector = CVector.getVectorByAngle(currentAngle);
			radiusVector.multiply(maxRadius);
			const shape = mainElements[i].getShape(true);
			const currentBounds = shape.getBounds(true);
			const halfWidth = (currentBounds.r - currentBounds.l) / 2;
			const halfHeight = (currentBounds.b - currentBounds.t) / 2;
			const newL = radiusVector.x - halfWidth;
			const newR = radiusVector.x + halfWidth;
			const newT = radiusVector.y - halfHeight;
			const newB = radiusVector.y + halfHeight;

			if (newL < cycleBounds.l) {
				cycleBounds.l = newL;
				radiusBounds.l = radiusVector.x;
			}
			if (newT < cycleBounds.t) {
				cycleBounds.t = newT;
				radiusBounds.t = radiusVector.y;
			}
			if (newR > cycleBounds.r) {
				cycleBounds.r = newR;
				radiusBounds.r = radiusVector.x;
			}
			if (newB > cycleBounds.b) {
				cycleBounds.b = newB;
				radiusBounds.b = radiusVector.y;
			}

			currentAngle = AscFormat.normalizeRotate(currentAngle + stepAngle);
		}

		const parentWidth = this.getParentNodeWidth();
		const parentHeight = this.getParentNodeHeight();
		const cycleHeight = cycleBounds.b - cycleBounds.t;
		const cycleWidth = cycleBounds.r - cycleBounds.l;
		let coefficient = Math.min(1, parentWidth / cycleWidth, parentHeight / cycleHeight);
		if (coefficient === 1 && smartartAlgorithm.factRuleState === factRuleState.enabled) {
			smartartAlgorithm.setFactRuleState(factRuleState.disabled);
			return;
		}
		let radiusCoefficient = coefficient;
		const scaleFactor = this.getOffsetScaleFactor(radiusBounds, cycleBounds);
		if (scaleFactor > 1) {
			radiusCoefficient = Math.max(coefficient, scaleFactor);
		}
		const adaptRadius = maxRadius * radiusCoefficient;
		this.calcValues.radius = maxRadius;
		const constrRadius = this.getConstrRadius(true);
		if (constrRadius !== undefined && adaptRadius > constrRadius) {
			radiusCoefficient *= constrRadius / adaptRadius;
		}
		this.calcValues.radiusCoefficient = radiusCoefficient;
		for (let i = 0; i < this.parentNode.childs.length; i += 1) {
			const child = this.parentNode.childs[i];
			child.setSizesScale(coefficient, coefficient);
		}
	};
	CycleAlgorithm.prototype.getConstrRadius = function (isCalculateScaleCoefficients) {
		const constrObject = this.parentNode.getConstraints(!isCalculateScaleCoefficients);
		const diameter = constrObject[AscFormat.Constr_type_diam];
		if (diameter !== undefined) {
			return diameter / 2;
		}
	};
	CycleAlgorithm.prototype.getOffsetScaleFactor = function (radiusBounds, cycleBounds) {
		const parentNodeWidth = this.getParentNodeWidth();
		const parentNodeHeight = this.getParentNodeHeight();

		const radiusHeight = radiusBounds.b - radiusBounds.t;
		const radiusWidth = radiusBounds.r - radiusBounds.l;
		const parentWidth = parentNodeWidth - (radiusBounds.l - cycleBounds.l) - (cycleBounds.r - radiusBounds.r);
		const parentHeight = parentNodeHeight - (radiusBounds.t - cycleBounds.t) - (cycleBounds.b - radiusBounds.b);
		if (radiusHeight !== 0 && radiusWidth !== 0) {
			return Math.max(1, Math.min(parentWidth / radiusWidth, parentHeight / radiusHeight));
		} else if (radiusHeight === 0 && radiusWidth !== 0) {
			return Math.max(1, parentWidth / radiusWidth);
		} else if (radiusWidth === 0 && radiusHeight !== 0) {
			return Math.max(1, parentHeight / radiusHeight);
		}
		return 1;
	};

	CycleAlgorithm.prototype.generateCurveShapeConnectors = function (smartartAlgorithm) {
		for (let i = 0; i < this.parentNode.childs.length; i += 1) {
			const child = this.parentNode.childs[i];
			if (child.algorithm instanceof ConnectorAlgorithm && child.algorithm.isCurveShape()) {
				child.algorithm.connectShapes(smartartAlgorithm);
			}
		}
	};
	CycleAlgorithm.prototype.resetConnectors = function () {
		for (let i = 0; i < this.parentNode.childs.length; i += 1) {
			const child = this.parentNode.childs[i];
			if (child.algorithm instanceof ConnectorAlgorithm) {
				child.algorithm.reset();
			}
		}
	};
	CycleAlgorithm.prototype.calculateShapePositions = function (smartartAlgorithm, isCalculateScaleCoefficients) {
		this.calculateStartValues();
		if (isCalculateScaleCoefficients) {
			this.calcScaleCoefficients(smartartAlgorithm);
		} else {
			this._calculateShapePositions();
			this.setConnections();
			this.generateCurveShapeConnectors(smartartAlgorithm);
			this.applyAlgorithmAligns(isCalculateScaleCoefficients);
			this.resetConnectors();
			this.applyPostAlgorithmSettings();
			this.createShadowShape(isCalculateScaleCoefficients);
		}
	};
	CycleAlgorithm.prototype.isAlongPath = function () {
		return this.params[AscFormat.Param_type_rotPath] === AscFormat.ParameterVal_rotationPath_alongPath;
	}
	CycleAlgorithm.prototype.getAlongRot = function (currentAngle) {
		if (this.isAlongPath()) {
			return AscFormat.normalizeRotate(currentAngle + Math.PI / 2);
		}
		return 0;
	};
	CycleAlgorithm.prototype._calculateShapePositions = function () {
		const childs = this.parentNode.childs;
		const radius = this.calcValues.radius * this.calcValues.radiusCoefficient;
		let currentAngle = this.calcValues.startAngle;
		const stepAngle = this.calcValues.stepAngle;
		const container = this.getShapeContainer();
		let startIndex = 0;
		let incAngleDivider = childs.length;
		if (this.calcValues.centerNodeIndex !== null) {
			incAngleDivider -= 1;
			const centerNode = this.getCenterNode();
			const shape = centerNode && centerNode.getShape(false);
			if  (shape) {
				shape.moveTo(-(shape.x + shape.width / 2), -(shape.y + shape.height / 2));
				container.push(shape);
			}
			startIndex = this.calcValues.centerNodeIndex + 1;
		}
		let incAngle = 0;
		if (incAngleDivider > 0) {
			incAngle = 2 * Math.PI / incAngleDivider;
		}

		for (let i = startIndex; i < childs.length; i++) {
			const child = childs[i];
				const shape = child.getShape(false);
				if (child.isContentNode() && shape) {
					const radiusGuideVector = CVector.getVectorByAngle(currentAngle);
					radiusGuideVector.multiply(radius);
					shape.setRadialInfo(radiusGuideVector, incAngle);
					const bounds = shape.getBounds();
					const width = bounds.r - bounds.l;
					const height = bounds.b - bounds.t;
					const offX = radiusGuideVector.x - (bounds.l + width / 2);
					const offY = radiusGuideVector.y - (bounds.t + height / 2);

					shape.rot = this.getAlongRot(currentAngle);
					shape.moveTo(offX, offY);
					currentAngle = currentAngle + stepAngle;
					container.push(shape);
				}
		}
	};
	CycleAlgorithm.prototype.createShadowShape = function (isCalculateScaleCoefficients) {
		return this.parentNode.createShadowShape(true, isCalculateScaleCoefficients);
	};
	CycleAlgorithm.prototype.getShapeContainer = function (isCalculateScaleCoefficient) {
		if (isCalculateScaleCoefficient) {
			if (this.coefficientShapeContainer === null) {
				this.coefficientShapeContainer = new CycleContainer();
			}
			return this.coefficientShapeContainer;
		}
		if (this.shapeContainer === null) {
			this.shapeContainer = new CycleContainer();
		}
		return this.shapeContainer;
	};

	function LinearAlgorithm() {
		PositionAlgorithm.call(this);
	}
	AscFormat.InitClassWithoutType(LinearAlgorithm, PositionAlgorithm);

	LinearAlgorithm.prototype.isRow = function () {
		return this.params[AscFormat.Param_type_linDir] === AscFormat.ParameterVal_linearDirection_fromL ||
		this.params[AscFormat.Param_type_linDir] === AscFormat.ParameterVal_linearDirection_fromR;
	}
	LinearAlgorithm.prototype.initParams = function (params) {
		PositionAlgorithm.prototype.initParams.call(this, params);
		if (this.params[AscFormat.Param_type_linDir] === undefined) {
			this.params[AscFormat.Param_type_linDir] = AscFormat.ParameterVal_linearDirection_fromL;
		}

		if (this.params[AscFormat.Param_type_vertAlign] === undefined) {
			this.params[AscFormat.Param_type_vertAlign] = AscFormat.ParameterVal_verticalAlignment_mid;
		}
		if (this.params[AscFormat.Param_type_horzAlign] === undefined) {
			this.params[AscFormat.Param_type_horzAlign] = AscFormat.ParameterVal_horizontalAlignment_ctr;
		}

		if (this.isRow()) {
			if (this.params[AscFormat.Param_type_nodeVertAlign] === undefined) {
				this.params[AscFormat.Param_type_nodeVertAlign] = AscFormat.ParameterVal_nodeVerticalAlignment_mid;
			}
			this.params[AscFormat.Param_type_nodeHorzAlign] = undefined;
		} else {
			if (this.params[AscFormat.Param_type_nodeHorzAlign] === undefined) {
				this.params[AscFormat.Param_type_nodeHorzAlign] = AscFormat.ParameterVal_nodeHorizontalAlignment_ctr;
			}
			this.params[AscFormat.Param_type_nodeVertAlign] = undefined;
		}
	}
	LinearAlgorithm.prototype.applyLeftNodeAlign = function (isCalculateScaleCoefficients) {
		const shapeContainer = this.getShapeContainer(isCalculateScaleCoefficients);
		const bounds = shapeContainer.getBounds();

		shapeContainer.forEachShape(function (shape) {
			const shapeBounds = shape.getBounds();
			shape.moveTo(bounds.l - shapeBounds.l, 0);
		});
	};
	LinearAlgorithm.prototype.applyRightNodeAlign = function (isCalculateScaleCoefficients) {
		const shapeContainer = this.getShapeContainer(isCalculateScaleCoefficients);
		const bounds = shapeContainer.getBounds();

		shapeContainer.forEachShape(function (shape) {
			const shapeBounds = shape.getBounds();
			shape.moveTo(bounds.r - shapeBounds.r, 0);
		});
	};
	LinearAlgorithm.prototype.applyCenterNodeAlign = function (isCalculateScaleCoefficients) {
		const shapeContainer = this.getShapeContainer(isCalculateScaleCoefficients);
		const bounds = shapeContainer.getBounds();
		const boundsCenter = bounds.l + (bounds.r - bounds.l) / 2;
		shapeContainer.forEachShape(function (shape) {
			const shapeBounds = shape.getBounds();
			shape.moveTo(boundsCenter - (shapeBounds.l + (shapeBounds.r - shapeBounds.l) / 2), 0);
		});
	};
	LinearAlgorithm.prototype.applyTopNodeAlign = function (isCalculateScaleCoefficients) {
		const shapeContainer = this.getShapeContainer(isCalculateScaleCoefficients);
		const bounds = shapeContainer.getBounds();

		shapeContainer.forEachShape(function (shape) {
			const shapeBounds = shape.getBounds();
			shape.moveTo(0, bounds.t - shapeBounds.t);
		});
	};
	LinearAlgorithm.prototype.applyBottomNodeAlign = function (isCalculateScaleCoefficients) {
		const shapeContainer = this.getShapeContainer(isCalculateScaleCoefficients);
		const bounds = shapeContainer.getBounds();

		shapeContainer.forEachShape(function (shape) {
			const shapeBounds = shape.getBounds();
			shape.moveTo(0, bounds.b - shapeBounds.b);
		});
	};
	LinearAlgorithm.prototype.applyMidNodeAlign = function (isCalculateScaleCoefficients) {
		const shapeContainer = this.getShapeContainer(isCalculateScaleCoefficients);
		const bounds = shapeContainer.getBounds();
		const boundsCenter = bounds.t + (bounds.b - bounds.t) / 2;
		shapeContainer.forEachShape(function (shape) {
			const shapeBounds = shape.getBounds();
			shape.moveTo(0, boundsCenter - (shapeBounds.t + (shapeBounds.b - shapeBounds.t) / 2));
		});
	};
	LinearAlgorithm.prototype.applyNodeAligns = function (isCalculateScaleCoefficients) {
		switch (this.params[AscFormat.Param_type_nodeHorzAlign]) {
			case AscFormat.ParameterVal_nodeHorizontalAlignment_l:
				this.applyLeftNodeAlign(isCalculateScaleCoefficients);
				break;
			case AscFormat.ParameterVal_nodeHorizontalAlignment_ctr:
				this.applyCenterNodeAlign(isCalculateScaleCoefficients);
				break;
			case AscFormat.ParameterVal_nodeHorizontalAlignment_r:
				this.applyRightNodeAlign(isCalculateScaleCoefficients);
				break;
			default:
				break;
		}

		switch (this.params[AscFormat.Param_type_nodeVertAlign]) {
			case AscFormat.ParameterVal_nodeVerticalAlignment_t:
				this.applyTopNodeAlign(isCalculateScaleCoefficients);
				break;
			case AscFormat.ParameterVal_nodeVerticalAlignment_mid:
				this.applyMidNodeAlign(isCalculateScaleCoefficients);
				break;
			case AscFormat.ParameterVal_nodeVerticalAlignment_b:
				this.applyBottomNodeAlign(isCalculateScaleCoefficients);
				break;
			default:
				break;
		}
	}
	LinearAlgorithm.prototype.createShadowShape = function (isCalculateScaleCoefficients) {
		return this.parentNode.createShadowShape(true, isCalculateScaleCoefficients);
	};
	LinearAlgorithm.prototype.isFromLeft = function () {
		return this.params[AscFormat.Param_type_linDir] === AscFormat.ParameterVal_linearDirection_fromL;
	};
	LinearAlgorithm.prototype.isFromTop = function () {
		return this.params[AscFormat.Param_type_linDir] === AscFormat.ParameterVal_linearDirection_fromT;
	};
	LinearAlgorithm.prototype._calculateShapePositions = function (isCalculateScaleCoefficients) {
		switch (this.params[AscFormat.Param_type_linDir]) {
			case AscFormat.ParameterVal_linearDirection_fromL:
			case AscFormat.ParameterVal_linearDirection_fromR:
				this.calculateRowShapePositions(isCalculateScaleCoefficients);
				break;
			case AscFormat.ParameterVal_linearDirection_fromT:
			case AscFormat.ParameterVal_linearDirection_fromB:
				this.calculateColumnShapePositions(isCalculateScaleCoefficients);
				break;
		}
	};
	LinearAlgorithm.prototype.isHaveChildLinear = function () {
		for (let i = 0; i < this.parentNode.childs.length; i += 1) {
			if (this.parentNode.childs[i].algorithm instanceof LinearAlgorithm) {
				return true;
			}
		}
		return false;
	}
	LinearAlgorithm.prototype.is2DFallback = function () {
		return this.params[AscFormat.Param_type_fallback] === AscFormat.ParameterVal_fallbackDimension_2D || this.isHaveChildLinear();
	}
	LinearAlgorithm.prototype.setScaleCoefficient = function () {
		if (!(!this.parentNode.parent || this.parentNode.parent.algorithm instanceof CompositeAlgorithm)) {
			return;
		}

		const parentHeight = this.getParentNodeHeight(false);
		const parentWidth = this.getParentNodeWidth(false);
		if (!(parentHeight && parentWidth)) {
			return;
		}
		const childs = this.parentNode.childs;
		const length = this.isHideLastChild() ? childs.length - 1 : childs.length;
		if (length > 0) {
			const isRow = this.isRow();
			const shapeContainer = this.getShapeContainer(true);
			const bounds = shapeContainer.getBounds();
			const width = bounds.r - bounds.l;
			const height = bounds.b - bounds.t;
			let widthCoefficient = Math.min(parentWidth / width, 1);
			let heightCoefficient = Math.min(parentHeight / height, 1);
			if (this.is2DFallback()) {
				const commonCoefficient = Math.min(widthCoefficient, heightCoefficient);
				widthCoefficient = commonCoefficient;
				heightCoefficient = commonCoefficient;
			} else if (isRow) {
				//todo check the case when the height depends on the width
				const cleanBounds  = shapeContainer.getBounds(true, true);
				const cleanHeightCoefficient = parentHeight / (cleanBounds.b - cleanBounds.t);
				if (cleanHeightCoefficient >= 1 || fAlgDeltaEqual(cleanHeightCoefficient, 1)) {
					heightCoefficient = 1;
				}
			} else {
				const cleanBounds  = shapeContainer.getBounds(true, true);
				const cleanWidthCoefficient = parentWidth / (cleanBounds.r - cleanBounds.l);
				if (cleanWidthCoefficient >= 1 || fAlgDeltaEqual(cleanWidthCoefficient, 1)) {
					widthCoefficient = 1;
				}
			}

			const heightSizeScale = isRow ? 1 : heightCoefficient;
			const widthSizeScale = isRow ? widthCoefficient : 1;
			const nodes = [this.parentNode];
			while (nodes.length) {
				const node = nodes.pop();
				for (let i = 0; i < node.childs.length; i += 1) {
					const childNode = node.childs[i];

					childNode.setSizesScale(widthSizeScale, heightSizeScale);
					const shapeType = childNode.layoutInfo.shape.type;
					if (shapeType !== AscFormat.LayoutShapeType_outputShapeType_none &&
						shapeType !== AscFormat.LayoutShapeType_outputShapeType_conn) {
						if (isRow) {
							if (childNode.getAspectRatio()) {
								childNode.setParentScale(AscFormat.Constr_type_w, heightCoefficient);
								childNode.setParentScale(AscFormat.Constr_type_h, heightCoefficient);
							} else {
								childNode.setParentScale(AscFormat.Constr_type_h, heightCoefficient);
							}
						} else {
							if (childNode.getAspectRatio()) {
								childNode.setParentScale(AscFormat.Constr_type_w, widthCoefficient);
								childNode.setParentScale(AscFormat.Constr_type_h, widthCoefficient);
							} else {
								childNode.setParentScale(AscFormat.Constr_type_w, widthCoefficient);
							}
						}
					}
					nodes.push(childNode);
				}
			}
		}
	};
	LinearAlgorithm.prototype.calculateShapePositions = function (smartartAlgorithm, isCalculateScaleCoefficients) {
		this._calculateShapePositions(isCalculateScaleCoefficients);
		this.applyNodeAligns(isCalculateScaleCoefficients);
		this.applyAlgorithmAligns(isCalculateScaleCoefficients);
		if (isCalculateScaleCoefficients) {
			this.setScaleCoefficient();
		} else {
			this.setConnections();
		}
		this.createShadowShape(isCalculateScaleCoefficients);
	};
	LinearAlgorithm.prototype.calculateRowShapePositions = function (isCalculateScaleCoefficients) {
		const fromLeft = this.isFromLeft();
		const childs = this.parentNode.childs;
		const length = this.isHideLastChild() ? childs.length - 1 : childs.length;
		let offX = 0;
		const shapeContainer = this.getShapeContainer(isCalculateScaleCoefficients);
		const row = new ShapeContainer();
		shapeContainer.push(row);
		for (let i = 0; i < length; i++) {
			const node = childs[i];
			const shape = node.getShape(isCalculateScaleCoefficients);
			const bounds = shape.getBounds();
			row.push(shape);
			if (fromLeft) {
				shape.moveTo(offX - bounds.l, 0);
				offX += bounds.r - bounds.l;
			} else {
				shape.moveTo(offX - bounds.r, 0);
				offX -= bounds.r - bounds.l;
			}
		}
	};

	LinearAlgorithm.prototype.calculateColumnShapePositions = function (isCalculateScaleCoefficients) {
		const fromTop = this.isFromTop();
		const childs = this.parentNode.childs;
		const length = this.isHideLastChild() ? childs.length - 1 : childs.length;
		if (length === 0) {
			return;
		}
		const shapeContainer = this.getShapeContainer(isCalculateScaleCoefficients, true);
		const row = new ShapeContainer();
		shapeContainer.push(row);
		let offY = 0;
		for (let i = 0; i < length; i++) {
			const node = childs[i];
			const shape = node.getShape(isCalculateScaleCoefficients);
			row.push(shape);
			const bounds = shape.getBounds();
			if (fromTop) {
				shape.moveTo(0, offY - bounds.t);
				offY += bounds.b - bounds.t;
			} else {
				shape.moveTo(0, offY - bounds.b);
				offY -= bounds.b - bounds.t;
			}
		}
	};
	LinearAlgorithm.prototype.getShapeContainer = function (isCalculateScaleCoefficient, isColumn) {
		if (isCalculateScaleCoefficient) {
			if (this.coefficientShapeContainer === null) {
				if (isColumn) {
					this.coefficientShapeContainer = new ShapeColumns();
				} else {
					this.coefficientShapeContainer = new ShapeRows();
				}

			}
			return this.coefficientShapeContainer;
		}
		if (this.shapeContainer === null) {
			if (isColumn) {
				this.shapeContainer = new ShapeColumns();
			} else {
				this.shapeContainer = new ShapeRows();
			}
		}
		return this.shapeContainer;
	};

	function ConnectorAlgorithm() {
		BaseAlgorithm.call(this);
		this.startNode = null;
		this.endNode = null;
		this.connectionDistances = {
			begin: null,
			end: null
		};
		this.stemThick = 0.6;
		this.diameterScale = 1;
		this.parentAlgorithm = null;
		this.calcValues = {
			edgePoints: null,
			radiusCenterPoint: null,
			connectionPoints: null,
			pointPositions: null
		}
	}
	AscFormat.InitClassWithoutType(ConnectorAlgorithm, BaseAlgorithm);
	ConnectorAlgorithm.prototype.reset = function () {
		this.calcValues = {
			edgePoints: null,
			radiusCenterPoint: null,
			connectionPoints: null,
			pointPositions: null
		}
		const connectionDistanceResolver = this.parentAlgorithm.parentNode.connectionDistanceResolver;
		if (connectionDistanceResolver) {
			connectionDistanceResolver.connectionDistance = null;
		}
	};
	ConnectorAlgorithm.prototype.isClockwise = function () {
		if (this.constrClockwise !== null) {
			return this.constrClockwise;
		}
		return this.parentAlgorithm.isClockwise();
	};
	ConnectorAlgorithm.prototype.initParams = function (params) {
		BaseAlgorithm.prototype.initParams.call(this, params);
		if (this.params[AscFormat.Param_type_dim] === undefined) {
			this.params[AscFormat.Param_type_dim] = AscFormat.ParameterVal_connectorDimension_2D;
		}
		if (this.params[AscFormat.Param_type_begSty] === undefined) {
			this.params[AscFormat.Param_type_begSty] = AscFormat.ParameterVal_arrowheadStyle_noArr;
		}
		if (this.params[AscFormat.Param_type_endSty] === undefined) {
			this.params[AscFormat.Param_type_endSty] = AscFormat.ParameterVal_arrowheadStyle_arr;
		}
		if (this.params[AscFormat.Param_type_connRout] === undefined) {
			this.params[AscFormat.Param_type_connRout] = AscFormat.ParameterVal_connectorRouting_stra;
		}
		if (this.params[AscFormat.Param_type_connRout] === AscFormat.ParameterVal_connectorRouting_bend) {
			if (this.params[AscFormat.Param_type_bendPt] === undefined) {
				this.params[AscFormat.Param_type_bendPt] = AscFormat.ParameterVal_bendPoint_end;
			}
		}
	}

	ConnectorAlgorithm.prototype.getConnectionDistCoefficient = function (isBegin) {
		switch (this.params[AscFormat.Param_type_connRout]) {
			case AscFormat.ParameterVal_connectorRouting_bend:
				if (!this.connectionDistances.begin && !this.connectionDistances.end) {
					return isBegin ? 0 : 1;
				} else {
					if (isBegin) {
						return this.connectionDistances.begin || 0;
					} else {
						return this.connectionDistances.end || 1;
					}
				}
			case AscFormat.ParameterVal_connectorRouting_stra:
				if (!this.connectionDistances.begin && !this.connectionDistances.end) {
					if (this.params[AscFormat.Param_type_dim] === AscFormat.ParameterVal_connectorDimension_2D) {
						return isBegin ? 0.22 : 0.25;
					}
					return 0;
				}
				return isBegin ? this.connectionDistances.begin : this.connectionDistances.end;
			default:
				return isBegin ? this.connectionDistances.begin : this.connectionDistances.end;
		}
	}
	ConnectorAlgorithm.prototype.getEdgePoints = function () {
		if (!this.calcValues.edgePoints) {
			this.calcValues.edgePoints = {
				start: null,
				end: null
			};
			this.calcValues.pointPositions = {
				start: null,
				end: null
			};

			let calcStartEdgePoint;
			let calcEndEdgePoint;
			const startEdgePoints = this.getAvailableEdgePoints(true);
			const endEdgePoints = this.getAvailableEdgePoints();
			let minDistance = null;
			for (let i = 0; i < startEdgePoints.length; i += 1) {
				const startEdgePoint = startEdgePoints[i].point;
				for (let j = 0; j < endEdgePoints.length; j += 1) {
					const endEdgePoint = endEdgePoints[j].point;
					const distance = startEdgePoint.getVector(endEdgePoint).getDistance();
					if (minDistance === null || minDistance > distance) {
						minDistance = distance;
						calcStartEdgePoint = startEdgePoints[i];
						calcEndEdgePoint = endEdgePoints[j];
					}
				}
			}

			if (calcStartEdgePoint && calcEndEdgePoint) {
				this.calcValues.edgePoints.start = calcStartEdgePoint.point;
				this.calcValues.pointPositions.start = calcStartEdgePoint.type;
				this.calcValues.edgePoints.end = calcEndEdgePoint.point;
				this.calcValues.pointPositions.end = calcEndEdgePoint.type;
			}
			if (this.isCurveConnection()) {
				this.calcCurveEdgePoints(this.calcValues.edgePoints.start, this.calcValues.edgePoints.end);
			}
		}
		if (this.calcValues.edgePoints.start && this.calcValues.edgePoints.end) {
			return this.calcValues.edgePoints;
		}
		return null;
	};
	ConnectorAlgorithm.prototype.getBendConnectionPoints = function (startPoint, endPoint) {
			const vector = startPoint.getVector(endPoint);
			const startConnectionPoint = this.getBendConnectionPoint(startPoint, startPoint, vector, true);
			const endConnectionPoint = this.getBendConnectionPoint(endPoint, startPoint, vector, false);
			return {start: startConnectionPoint, end: endConnectionPoint};
	};

	ConnectorAlgorithm.prototype.getBendConnectionPoint = function (point, startPoint, vector, isStart) {
		const position = this.getPointPosition(isStart);
		const coefficient = this.getConnectionDistCoefficient(isStart);
		switch (position) {
			case AscFormat.ParameterVal_connectorPoint_bCtr:
			case AscFormat.ParameterVal_connectorPoint_tCtr:
				return new CCoordPoint(point.x, vector.y * coefficient + startPoint.y);
			case AscFormat.ParameterVal_connectorPoint_midR:
			case AscFormat.ParameterVal_connectorPoint_midL:
			case AscFormat.ParameterVal_connectorPoint_tL:
			case AscFormat.ParameterVal_connectorPoint_bR:
			case AscFormat.ParameterVal_connectorPoint_tR:
			case AscFormat.ParameterVal_connectorPoint_bL:
				return new CCoordPoint(vector.x * coefficient + startPoint.x, point.y);
			case AscFormat.ParameterVal_connectorPoint_radial:
			case AscFormat.ParameterVal_connectorPoint_auto:
			default:
				return null;
		}
	};
	ConnectorAlgorithm.prototype.getStraightConnectionPoints = function (startPoint, endPoint) {
		const startCoefficient = this.getConnectionDistCoefficient(true);
		const endCoefficient = 1 - this.getConnectionDistCoefficient(false);
		const endConnectionPoint = new CCoordPoint((endPoint.x - startPoint.x) * endCoefficient + startPoint.x, (endPoint.y - startPoint.y) * endCoefficient + startPoint.y);
		const startConnectionPoint = new CCoordPoint((endPoint.x - startPoint.x) * startCoefficient + startPoint.x, (endPoint.y - startPoint.y) * startCoefficient + startPoint.y);

		return {start: startConnectionPoint, end: endConnectionPoint};
	};
	ConnectorAlgorithm.prototype.getCenterPoint = function () {
		if (this.calcValues.radiusCenterPoint) {
			return this.calcValues.radiusCenterPoint;
		}
		if (this.parentAlgorithm instanceof CycleAlgorithm) {
			const startShape = this.getStartShape();
			if (startShape.radialVector) {
				const bounds = startShape.getBounds();
				const centerShapePoint = getShapePoint(bounds);
				return new CCoordPoint(centerShapePoint.x - startShape.radialVector.x, centerShapePoint.y - startShape.radialVector.y);
			}
		}
	};
	ConnectorAlgorithm.prototype.getCurveConnectionPoints = function (startPoint, endPoint) {
		const result = {start: null, end: null};
		const centerPoint =  this.getCenterPoint();
		if (centerPoint) {
			const startCoefficient = this.getConnectionDistCoefficient(true);
			const endCoefficient = this.getConnectionDistCoefficient(false);

			const startVector = centerPoint.getVector(startPoint);
			const radius = startVector.getDistance();
			const endVector = centerPoint.getVector(endPoint);
			const angle = startVector.getAngleBetween(endVector, !this.isClockwise());
			const startAngle = startVector.getAngle();
			const endAngle = endVector.getAngle();

			let startConnectionVector;
			let endConnectionVector;
			if (this.isClockwise() === this.isLongCurve()) {
				startConnectionVector = CVector.getVectorByAngle(startAngle - angle * startCoefficient);
				endConnectionVector = CVector.getVectorByAngle(endAngle + angle * endCoefficient);
			} else {
				startConnectionVector = CVector.getVectorByAngle(startAngle + angle * startCoefficient);
				endConnectionVector = CVector.getVectorByAngle(endAngle - angle * endCoefficient);
			}
			result.start = new CCoordPoint(centerPoint.x + startConnectionVector.x * radius, centerPoint.y + startConnectionVector.y * radius);
			result.end = new CCoordPoint(centerPoint.x + endConnectionVector.x * radius, centerPoint.y + endConnectionVector.y * radius);
		}
		return result;
	};
	ConnectorAlgorithm.prototype.getConnectionPoints = function () {
		if (!this.calcValues.connectionPoints) {
			this.calcValues.connectionPoints = {
				start: null,
				end : null
			};
			const edgePoints = this.getEdgePoints();
			if (edgePoints) {
				const startPoint = edgePoints.start;
				const endPoint = edgePoints.end;
				let connectionPoints;
				switch (this.params[AscFormat.Param_type_connRout]) {
					case AscFormat.ParameterVal_connectorRouting_bend:
						connectionPoints = this.getBendConnectionPoints(startPoint, endPoint);
						break;
					case AscFormat.ParameterVal_connectorRouting_stra:
						connectionPoints = this.getStraightConnectionPoints(startPoint, endPoint);
						break;
					case AscFormat.ParameterVal_connectorRouting_curve:
					case AscFormat.ParameterVal_connectorRouting_longCurve:
						connectionPoints = this.getCurveConnectionPoints(startPoint, endPoint);
						break;
					default:
						break;
				}

				if (connectionPoints) {
					this.calcValues.connectionPoints.start = connectionPoints.start;
					this.calcValues.connectionPoints.end = connectionPoints.end;
				}
			}
		}
		return this.calcValues.connectionPoints;
	}
	ConnectorAlgorithm.prototype.setParentAlgorithm = function (algorithm) {
		this.parentAlgorithm = algorithm;
	};
	ConnectorAlgorithm.prototype.setConnectionDistance = function (value, isStart) {
		if (isStart) {
			this.connectionDistances.begin = value;
		} else {
			this.connectionDistances.end = value;
		}
	};
	BaseAlgorithm.prototype.setStemThick = function (value) {
		this.stemThick = value;
	};
	ConnectorAlgorithm.prototype.getPointPosition = function (isStart) {
		if (!this.calcValues.pointPositions) {
			this.getEdgePoints();
		}
		return isStart ? this.calcValues.pointPositions.start : this.calcValues.pointPositions.end;
	};
	ConnectorAlgorithm.prototype.getStartShape = function () {
		return this.startNode.getShape(false);
	};
	ConnectorAlgorithm.prototype.getEndShape = function () {
		return this.endNode.getShape(false);
	};
	ConnectorAlgorithm.prototype.isAutoCenterPoint = function (isStart, isFromLayout) {
		const connCoefficient = this.getConnectionDistCoefficient(isStart);
		return connCoefficient === 0 && this.params[AscFormat.Param_type_dim] === AscFormat.ParameterVal_connectorDimension_2D &&
			isFromLayout;
	};
	ConnectorAlgorithm.prototype.getAutoEdgePoint = function (isStart, isFromLayout) {
		const startShape = this.getStartShape();
		const endShape = this.getEndShape();
		const startBounds = startShape.getBounds();
		const endBounds = endShape.getBounds();
		if (this.isAutoCenterPoint(isStart, isFromLayout)) {
			return isStart ? getShapePoint(startBounds) : getShapePoint(endBounds);
		}
		const startPoint = getShapePoint(startBounds);
		const endPoint = getShapePoint(endBounds);
		let guideVector;
		if (isStart) {
			guideVector = new CVector(endPoint.x - startPoint.x, endPoint.y - startPoint.y);
		} else {
			guideVector = new CVector(startPoint.x - endPoint.x, startPoint.y - endPoint.y);
		}
		if (fAlgDeltaEqual(guideVector.getDistance(), 0)) {
			return null;
		}
		const bounds = isStart ? startBounds : endBounds;
		return getMinShapeEdgePoint(bounds, guideVector);
	};
	ConnectorAlgorithm.prototype.getEllipseRadialEdgePoint = function (radialInfo, bounds, isStart) {
		const cycleAngle = radialInfo.angle;
		const centerPoint = radialInfo.point;
		const radius = radialInfo.radius;
		const shapeRadius = (bounds.r - bounds.l) / 2;
		const shapeAngle =  Math.acos(1 - ((shapeRadius * shapeRadius) / (2 * radius * radius)));
		let angle = cycleAngle;
		if (radialInfo.isClockwise) {
			if (isStart) {
				angle += shapeAngle;
			} else {
				angle -= shapeAngle;
			}
		} else {
			if (isStart) {
				angle -= shapeAngle;
			} else {
				angle += shapeAngle;
			}
		}


		return new CCoordPoint(Math.cos(angle) * radius + centerPoint.x, Math.sin(angle) * radius + centerPoint.y);
	};
	ConnectorAlgorithm.prototype.isPointOnInterval = function (point, startSegment, endSegment) {
		if (startSegment.x === endSegment.x) {
			return (point.y > startSegment.y) && (point.y < endSegment.y) && !(fIntervalDeltaEqual(point.y, startSegment.y) || fIntervalDeltaEqual(point.y, endSegment.y));
		} else {
			return (point.x > startSegment.x) && (point.x < endSegment.x) && !(fIntervalDeltaEqual(point.x, startSegment.x) || fIntervalDeltaEqual(point.x, endSegment.x));
		}
	};
	ConnectorAlgorithm.prototype.isPointOnSegment = function (point, startSegment, endSegment) {
		if (startSegment.x === endSegment.x) {
			return (point.y > startSegment.y || fIntervalDeltaEqual(point.y, startSegment.y)) && (point.y < endSegment.y || fIntervalDeltaEqual(point.y, endSegment.y));
		} else {
			return (point.x > startSegment.x || fIntervalDeltaEqual(point.x, startSegment.x)) && (point.x < endSegment.x || fIntervalDeltaEqual(point.x, endSegment.x));
		}
	}
	ConnectorAlgorithm.prototype.getCycleRadius = function () {
		const radiusCoefficient = this.parentAlgorithm.calcValues.radiusCoefficient || 1;
		if (this.parentAlgorithm instanceof CycleAlgorithm) {
			const diamScale = Math.abs(this.diameterScale);
			return this.parentAlgorithm.calcValues.radius * radiusCoefficient * diamScale;
		}
		const constrDiam = this.parentNode.getConstr(AscFormat.Constr_type_diam, true, true);
		if (constrDiam !== undefined) {
			return Math.abs(constrDiam) / 2;
		}

	};

	ConnectorAlgorithm.prototype.calcCurveEdgePoints = function (startPoint, endPoint) {
		if (!(startPoint && endPoint)) {
			startPoint = this.getAutoEdgePoint(true);
			endPoint = this.getAutoEdgePoint();
		}
		if (!(startPoint && endPoint)) {
			const startShape = this.getStartShape();
			const endShape = this.getEndShape();
			const startBounds = startShape.getBounds();
			const endBounds = endShape.getBounds();
			startPoint = getShapePoint(startBounds);
			endPoint = getShapePoint(endBounds);
		}
		if (startPoint && endPoint) {
			this.calcValues.edgePoints.start = startPoint;
			this.calcValues.edgePoints.end = endPoint;

			const cycleRadius = this.getCycleRadius();
			const isClockwise = this.isClockwise();
			const lineVector = startPoint.getVector(endPoint);
			const circlesIntersectionPoints = AscFormat.circlesIntersection(startPoint.x, startPoint.y, cycleRadius, endPoint.x, endPoint.y, cycleRadius);
			if (circlesIntersectionPoints.length) {
				const firstEllipsePoint = new CCoordPoint(circlesIntersectionPoints[0].x, circlesIntersectionPoints[0].y);
				const secondEllipsePoint = new CCoordPoint(circlesIntersectionPoints[1].x, circlesIntersectionPoints[1].y);
				let _isClockwisePoints = isClockwisePoints(firstEllipsePoint, startPoint, endPoint);

				if (_isClockwisePoints) {
					this.calcValues.radiusCenterPoint = isClockwise ? firstEllipsePoint : secondEllipsePoint;
				} else {
					this.calcValues.radiusCenterPoint = isClockwise ? secondEllipsePoint : firstEllipsePoint;
				}
			} else {
				this.calcValues.radiusCenterPoint = new CCoordPoint(startPoint.x + lineVector.x / 2, startPoint.y + lineVector.y / 2);
			}
		}
	};
	ConnectorAlgorithm.prototype.getRectRadialEdgePoint = function (radialInfo, bounds, isStart) {
		const centerPoint = radialInfo.point;
		const radius = radialInfo.radius;
		const isClockwise = radialInfo.isClockwise;

		const ellipseBounds = {
			l: centerPoint.x - radius,
			r: centerPoint.x + radius,
			t: centerPoint.y - radius,
			b: centerPoint.y + radius
		};

		const linePoints = [
			[new CCoordPoint(bounds.l, bounds.t), new CCoordPoint(bounds.r, bounds.t)],
			[new CCoordPoint(bounds.l, bounds.t), new CCoordPoint(bounds.l, bounds.b)],
			[new CCoordPoint(bounds.l, bounds.b), new CCoordPoint(bounds.r, bounds.b)],
			[new CCoordPoint(bounds.r, bounds.t), new CCoordPoint(bounds.r, bounds.b)]
		];
		const rectCenterPoint = getShapePoint(bounds);
		for (let i = 0; i < linePoints.length; i += 1) {
			const coords = linePoints[i];
			const paramLine = getParametricLinEquation(coords[0], new CVector(coords[1].x - coords[0].x, coords[1].y - coords[0].y));
			const answer = resolveParameterLineAndShapeEquation(ellipseBounds, paramLine);
			if (!answer.bError) {
				let point;
				const point1 = new CCoordPoint(paramLine.x + paramLine.ax * answer.x1, paramLine.y + paramLine.ay * answer.x1);
				const point2 = new CCoordPoint(paramLine.x + paramLine.ax * answer.x2, paramLine.y + paramLine.ay * answer.x2);

				if (this.isPointOnInterval(point1, coords[0], coords[1]) && this.isPointOnInterval(point2, coords[0], coords[1])) {
					continue;
				} else if (this.isPointOnSegment(point1, coords[0], coords[1])) {
					point = point1;
				} else if (this.isPointOnSegment(point2, coords[0], coords[1])) {
					point = point2;
				} else {
					continue;
				}
				if (isStart && isClockwise || !isStart && !isClockwise) {
					if (isClockwisePoints(centerPoint, rectCenterPoint, point)) {
						return point;
					}
				} else {
					if (isClockwisePoints(centerPoint, point, rectCenterPoint)) {
						return point;
					}
				}
			}
		}
		return null;
	};
	ConnectorAlgorithm.prototype.getRadialEdgePoint = function (isStart) {
		const shape = isStart ? this.getStartShape() : this.getEndShape();
		const radialInfo = this.parentAlgorithm.getRadialConnectionInfo(shape.node);
		if (!radialInfo || radialInfo.radius === 0) {
			return null;
		}

		const bounds = shape.getBounds();
		if (bounds.isEllipse) {
			return this.getEllipseRadialEdgePoint(radialInfo, bounds, isStart);
		}
		return this.getRectRadialEdgePoint(radialInfo, bounds, isStart);
	};
	ConnectorAlgorithm.prototype.getAvailablePointPositions = function (isStart) {
		const parameters = isStart ? this.params[AscFormat.Param_type_begPts] : this.params[AscFormat.Param_type_endPts];
		return parameters || [AscFormat.ParameterVal_connectorPoint_auto];

	}
	ConnectorAlgorithm.prototype.isCurveConnection = function () {
		return this.params[AscFormat.Param_type_connRout] === AscFormat.ParameterVal_connectorRouting_curve ||
		this.params[AscFormat.Param_type_connRout] === AscFormat.ParameterVal_connectorRouting_longCurve;
	};
	ConnectorAlgorithm.prototype.isCurveShape = function () {
		return this.params[AscFormat.Param_type_connRout] === AscFormat.ParameterVal_connectorRouting_curve &&
			this.params[AscFormat.Param_type_dim] === AscFormat.ParameterVal_connectorDimension_2D;
	};
	ConnectorAlgorithm.prototype.getAvailableEdgePoints = function (isStart) {
		const result = [];
		const pointPositions = this.getAvailablePointPositions(isStart);
		for (let i = 0; i < pointPositions.length; i += 1) {
			const type = pointPositions[i];
			let point;
			switch (type) {
				case AscFormat.ParameterVal_connectorPoint_radial:
					point = this.getRadialEdgePoint(isStart);
					break;
				case AscFormat.ParameterVal_connectorPoint_tL:
					point = this.getTopLeftEdgePoint(isStart);
					break;
				case AscFormat.ParameterVal_connectorPoint_tCtr:
					point = this.getTopCenterEdgePoint(isStart);
					break;
				case AscFormat.ParameterVal_connectorPoint_tR:
					point = this.getTopRightEdgePoint(isStart);
					break;
				case AscFormat.ParameterVal_connectorPoint_midR:
					point = this.getMidRightEdgePoint(isStart);
					break;
				case AscFormat.ParameterVal_connectorPoint_bR:
					point = this.getBottomRightEdgePoint(isStart);
					break;
				case AscFormat.ParameterVal_connectorPoint_bCtr:
					point = this.getBottomCenterEdgePoint(isStart);
					break;
				case AscFormat.ParameterVal_connectorPoint_bL:
					point = this.getBottomLeftEdgePoint(isStart);
					break;
				case AscFormat.ParameterVal_connectorPoint_midL:
					point = this.getMidLeftEdgePoint(isStart);
					break;
				case AscFormat.ParameterVal_connectorPoint_auto:
				default:
					point = this.getAutoEdgePoint(isStart, true);
					break;
			}
			if (point) {
				result.push({point: point, type: type});
			}
		}
		return result;
	};
	ConnectorAlgorithm.prototype.getTopLeftEdgePoint = function (isStart) {
		const shape = isStart ? this.getStartShape() : this.getEndShape();
		const bounds = shape.getBounds();
		return new CCoordPoint(bounds.l, bounds.t);
	};
	ConnectorAlgorithm.prototype.getTopCenterEdgePoint = function (isStart) {
		const shape = isStart ? this.getStartShape() : this.getEndShape();
		const bounds = shape.getBounds();
		return new CCoordPoint(bounds.l + (bounds.r - bounds.l) / 2, bounds.t);
	};
	ConnectorAlgorithm.prototype.getTopRightEdgePoint = function (isStart) {
		const shape = isStart ? this.getStartShape() : this.getEndShape();
		const bounds = shape.getBounds();
		return new CCoordPoint(bounds.r, bounds.t);
	};
	ConnectorAlgorithm.prototype.getMidRightEdgePoint = function (isStart) {
		const shape = isStart ? this.getStartShape() : this.getEndShape();
		const bounds = shape.getBounds();
		return new CCoordPoint(bounds.r, bounds.t + (bounds.b - bounds.t) / 2);
	};
	ConnectorAlgorithm.prototype.getBottomRightEdgePoint = function (isStart) {
		const shape = isStart ? this.getStartShape() : this.getEndShape();
		const bounds = shape.getBounds();
		return new CCoordPoint(bounds.r, bounds.b);
	};
	ConnectorAlgorithm.prototype.getBottomCenterEdgePoint = function (isStart) {
		const shape = isStart ? this.getStartShape() : this.getEndShape();
		const bounds = shape.getBounds();
		return new CCoordPoint(bounds.l + (bounds.r - bounds.l) / 2,  bounds.b);
	};
	ConnectorAlgorithm.prototype.getBottomLeftEdgePoint = function (isStart) {
		const shape = isStart ? this.getStartShape() : this.getEndShape();
		const bounds = shape.getBounds();
		return new CCoordPoint(bounds.l, bounds.b);
	};
	ConnectorAlgorithm.prototype.getMidLeftEdgePoint = function (isStart) {
		const shape = isStart ? this.getStartShape() : this.getEndShape();
		const bounds = shape.getBounds();
		return new CCoordPoint(bounds.l, bounds.t + (bounds.b - bounds.t) / 2);
	};

	ConnectorAlgorithm.prototype.createShadowShape = function (isCalculateScaleCoefficients) {
		return this.parentNode.createShadowShape(false, isCalculateScaleCoefficients);
	};
	ConnectorAlgorithm.prototype.calculateShapePositions = function (smartartAlgorithm, isCalculateScaleCoefficients) {
		this.createShadowShape(isCalculateScaleCoefficients);
		if (!isCalculateScaleCoefficients) {
			smartartAlgorithm.addConnectorAlgorithm(this);
		}
	}

	ConnectorAlgorithm.prototype.setFirstConnectorNode = function (node) {
		this.startNode = node;
	};
	ConnectorAlgorithm.prototype.setLastConnectorNode = function (node) {
		this.endNode = node;
	};
	ConnectorAlgorithm.prototype.connectShapes = function (smartartAlgorithm) {
		if (this.startNode && this.endNode) {
			if (this.params[AscFormat.Param_type_dim] === AscFormat.ParameterVal_connectorDimension_2D) {
				this.createShapeConnector(smartartAlgorithm);
			} else if (this.params[AscFormat.Param_type_dim] === AscFormat.ParameterVal_connectorDimension_1D) {
				this.createLineConnector();
			}
		}
	};
	ConnectorAlgorithm.prototype.getStraightAdjLst = function (shapeType) {
		if (shapeType !== AscFormat.LayoutShapeType_shapeType_rect) {
			const customAdjLst = new AscFormat.AdjLst();
			const adj1 = new AscFormat.Adj();
			const adj2 = new AscFormat.Adj();
			adj1.setIdx(1);
			adj2.setIdx(2);
			adj1.setVal(this.stemThick);
			adj2.setVal(0.5);
			customAdjLst.addToLst(0, adj1);
			customAdjLst.addToLst(0, adj2);
			return customAdjLst;
		}
	};
	ConnectorAlgorithm.prototype.isLongCurve = function () {
		return this.params[AscFormat.Param_type_connRout] === AscFormat.ParameterVal_connectorRouting_longCurve;
	}
	ConnectorAlgorithm.prototype.getConnectorShapeType = function () {
		const endStyle = this.params[AscFormat.Param_type_endSty];
		const beginStyle = this.params[AscFormat.Param_type_begSty];
		if (this.params[AscFormat.Param_type_connRout] === AscFormat.ParameterVal_connectorRouting_curve) {
			if (endStyle === AscFormat.ParameterVal_arrowheadStyle_arr && beginStyle === AscFormat.ParameterVal_arrowheadStyle_arr) {
				return AscFormat.LayoutShapeType_shapeType_leftRightCircularArrow;
			} else if (endStyle === AscFormat.ParameterVal_arrowheadStyle_arr) {
				if (this.isClockwise()) {
					return AscFormat.LayoutShapeType_shapeType_circularArrow;
				}
				return AscFormat.LayoutShapeType_shapeType_leftCircularArrow;
			} else if (beginStyle === AscFormat.ParameterVal_arrowheadStyle_arr) {
				if (this.isClockwise()) {
					return AscFormat.LayoutShapeType_shapeType_leftCircularArrow;
				}
				return AscFormat.LayoutShapeType_shapeType_circularArrow;
			}
			return AscFormat.LayoutShapeType_shapeType_blockArc;
		} else if (this.isLongCurve()) {
			if (endStyle === AscFormat.ParameterVal_arrowheadStyle_arr && beginStyle === AscFormat.ParameterVal_arrowheadStyle_arr) {
				return AscFormat.LayoutShapeType_shapeType_leftRightCircularArrow;
			} else if (endStyle === AscFormat.ParameterVal_arrowheadStyle_arr) {
				if (this.isClockwise()) {
					return AscFormat.LayoutShapeType_shapeType_leftCircularArrow;
				}
				return AscFormat.LayoutShapeType_shapeType_circularArrow;
			} else if (beginStyle === AscFormat.ParameterVal_arrowheadStyle_arr) {
				if (this.isClockwise()) {
					return AscFormat.LayoutShapeType_shapeType_circularArrow;
				}
				return AscFormat.LayoutShapeType_shapeType_leftCircularArrow;
			}
			return AscFormat.LayoutShapeType_shapeType_blockArc;
		} else {
			if (endStyle === AscFormat.ParameterVal_arrowheadStyle_arr && beginStyle === AscFormat.ParameterVal_arrowheadStyle_arr) {
				return AscFormat.LayoutShapeType_shapeType_leftRightArrow;
			} else if (endStyle === AscFormat.ParameterVal_arrowheadStyle_arr) {
				return AscFormat.LayoutShapeType_shapeType_rightArrow;
			} else if (beginStyle === AscFormat.ParameterVal_arrowheadStyle_arr) {
				return AscFormat.LayoutShapeType_shapeType_leftArrow;
			}
			return AscFormat.LayoutShapeType_shapeType_rect;
		}
	}
	ConnectorAlgorithm.prototype.getTemplateConnectorShape = function () {
		const shape = this.parentNode.shape;
		const connectorShape = new ShadowShape();
		connectorShape.shape = shape.shape;

		connectorShape.type = this.getConnectorShapeType();
		connectorShape.cleanParams = {};
		connectorShape.cleanParams.width = shape.cleanParams.width;
		connectorShape.cleanParams.height = shape.cleanParams.height;
		connectorShape.cleanParams.x = shape.cleanParams.x;
		connectorShape.cleanParams.y = shape.cleanParams.y;
		connectorShape.node = this.parentNode;
		return connectorShape;
	};

	ConnectorAlgorithm.prototype.createBlockArcCurveShapeConnector = function (startAngle, endAngle, radius) {
		const connectorShape = this.getTemplateConnectorShape();
		const connectorHeight = this.parentNode.getConstr(AscFormat.Constr_type_h, true);
		const stemThick = (connectorHeight / radius) / 3.4;
		const customAdjLst = new AscFormat.AdjLst();
		const adj1 = new AscFormat.Adj();
		const adj2 = new AscFormat.Adj();
		const adj3 = new AscFormat.Adj();
		adj1.setIdx(1);
		adj2.setIdx(2);
		adj3.setIdx(3);
		adj1.setVal(startAngle * radToDeg);
		adj2.setVal(endAngle * radToDeg);
		adj3.setVal(stemThick);
		customAdjLst.addToLst(0, adj1);
		customAdjLst.addToLst(0, adj2);
		customAdjLst.addToLst(0, adj3);
		connectorShape.customAdj = customAdjLst;

		const centerPoint = this.getCenterPoint();
		const absRadius = Math.abs(radius);
		const diameterSide = absRadius * 2;
		const truthSide = diameterSide / (1 - stemThick / 2);
		const truthRadius = truthSide / 2;
		connectorShape.width = truthSide;
		connectorShape.height = truthSide;
		connectorShape.x = centerPoint.x - truthRadius;
		connectorShape.y = centerPoint.y - truthRadius;
		const shape = this.parentNode.getShape();
		shape.connectorShape = connectorShape;
	};
	ConnectorAlgorithm.prototype.createCircularArrowCurveShapeConnector = function (startAngle, endAngle, radius) {
		const connectorShape = this.getTemplateConnectorShape();
		const connectorHeight = this.parentNode.getConstr(AscFormat.Constr_type_h, true, true);
		const constrArrowWidth = this.parentNode.getConstr(AscFormat.Constr_type_wArH, true, true);
		const constrStemThick = this.parentNode.getConstr(AscFormat.Constr_type_stemThick, true, true);
		const constrArrowHeight = this.parentNode.getConstr(AscFormat.Constr_type_hArH, true, true);
		let adjArrowWidth;
		let stemThick;
		let adjArrowHeight;
		let truthSide;
		// todo: increase accuracy in Continuous Cycle
		if (constrArrowWidth !== undefined && constrStemThick !== undefined && constrArrowHeight !== undefined) {
			const diam = radius * 2;
			const diamScale = Math.abs(this.diameterScale);
			if (connectorHeight !== undefined) {
				adjArrowWidth = constrArrowWidth / radius * 0.6195;
				stemThick = adjArrowWidth / 1.16666;
				adjArrowHeight = (constrArrowHeight / radius) * 28.65983;
				truthSide = (diam / (1 - adjArrowWidth * 2)) * 0.98999;
			} else {
				adjArrowWidth = 1 / (diamScale) * 0.0598728;
				stemThick = adjArrowWidth * 0.963;
				adjArrowHeight = (1 / diamScale) * 5.731787;
				truthSide = (diam / (1 - adjArrowWidth * 2)) * 0.9965;
			}
		} else if (connectorHeight !== undefined) {
			const addValue = connectorHeight * 0.15;
			truthSide = ((radius + addValue) * 2 + connectorHeight);
			adjArrowWidth = 0.5 - ((radius - addValue / 3) / truthSide);
			stemThick = adjArrowWidth * 0.85718;
			adjArrowHeight = (connectorHeight / radius) * radToDeg / 2;
		} else {
			return;
		}


		const customAdjLst = new AscFormat.AdjLst();
		const adj1 = new AscFormat.Adj();
		const adj2 = new AscFormat.Adj();
		const adj3 = new AscFormat.Adj();
		const adj4 = new AscFormat.Adj();
		const adj5 = new AscFormat.Adj();
		adj1.setIdx(1);
		adj2.setIdx(2);
		adj3.setIdx(3);
		adj4.setIdx(4);
		adj5.setIdx(5);
		adj1.setVal(stemThick);
		adj2.setVal(adjArrowHeight);
		if (this.getConnectorShapeType() === AscFormat.LayoutShapeType_shapeType_circularArrow) {
			adj3.setVal(endAngle * radToDeg - adjArrowHeight);
			adj4.setVal(startAngle * radToDeg);
		} else {
			adj3.setVal(startAngle * radToDeg + adjArrowHeight);
			adj4.setVal(endAngle * radToDeg);
		}
		adj5.setVal(adjArrowWidth);
		customAdjLst.addToLst(0, adj1);
		customAdjLst.addToLst(0, adj2);
		customAdjLst.addToLst(0, adj3);
		customAdjLst.addToLst(0, adj4);
		customAdjLst.addToLst(0, adj5);
		connectorShape.customAdj = customAdjLst;

		const centerPoint = this.getCenterPoint();
		const truthRadius = truthSide / 2;
		connectorShape.width = truthSide;
		connectorShape.height = truthSide;
		connectorShape.x = centerPoint.x - truthRadius;
		connectorShape.y = centerPoint.y - truthRadius;
		const shape = this.parentNode.getShape();
		shape.connectorShape = connectorShape;
	};
	ConnectorAlgorithm.prototype.createCurveShapeConnector = function (smartartAlgorithm, startPoint, endPoint) {
		const centerPoint = this.getCenterPoint();
		const startVector = centerPoint.getVector(startPoint);
		const endVector = centerPoint.getVector(endPoint);
		let startAngle = startVector.getAngle();
		let endAngle = endVector.getAngle();
		const isClockwise = this.isClockwise();
		if (fAlgDeltaEqual(startAngle, endAngle)) {
			endAngle = startAngle + 0.000001;
		} else if (isClockwise === this.isLongCurve()) {
			const temp = startAngle;
			startAngle = endAngle;
			endAngle = temp;
		}
		const radius = startVector.getDistance();
		switch (this.getConnectorShapeType()) {
			case AscFormat.LayoutShapeType_shapeType_blockArc:
				this.createBlockArcCurveShapeConnector(startAngle, endAngle, radius);
				break;
			case AscFormat.LayoutShapeType_shapeType_circularArrow:
			case AscFormat.LayoutShapeType_shapeType_leftCircularArrow:
				this.createCircularArrowCurveShapeConnector(startAngle, endAngle, radius);
				break;
			default:
				break;
		}

		const shape = this.parentNode.getShape();
		const connectorShape = shape.connectorShape;
		if (connectorShape) {
			const coefficient = connectorShape.width / shape.cleanParams.width;
			const heightScale = this.parentNode.getHeightScale(true);
			const widthScale = this.parentNode.getWidthScale(true);
			const newWidth = connectorShape.width * widthScale;
			const newHeight = connectorShape.height * heightScale;
			connectorShape.x += (connectorShape.width - newWidth) / 2;
			connectorShape.y += (connectorShape.height - newHeight) / 2;
			connectorShape.width = newWidth;
			connectorShape.height = newHeight;
			this.applyPostAlgorithmSettingsForShape(smartartAlgorithm, connectorShape, coefficient);
		}
	};
	ConnectorAlgorithm.prototype.createStraightShapeConnector = function (smartartAlgorithm, startPoint, endPoint) {
		const cx = (startPoint.x + endPoint.x) / 2;
		const cy = (startPoint.y + endPoint.y) / 2;

		const arrowVector = new CVector(endPoint.x - startPoint.x, endPoint.y - startPoint.y);

		let width;
		const connectionDistanceResolver = this.parentAlgorithm.parentNode.connectionDistanceResolver;
		const minConnectionDistance = connectionDistanceResolver && connectionDistanceResolver.getConnectionDistance();
		if (connectionDistanceResolver && minConnectionDistance !== -1) {
			width = minConnectionDistance;
		} else {
			width = arrowVector.getDistance();
		}
		const constrObj = this.parentNode.getConstraints(true);
		const height = constrObj[AscFormat.Constr_type_hArH] !== undefined ? constrObj[AscFormat.Constr_type_hArH] : this.parentNode.shape.height;

		const x = cx - width / 2;
		const y = cy - height / 2;
		const shape = this.parentNode.getShape();
		const connectorShape = this.getTemplateConnectorShape();
		connectorShape.customAdj = this.getStraightAdjLst(connectorShape.type);

		connectorShape.x = x;
		connectorShape.y = y;
		connectorShape.rot = arrowVector.getAngle();
		shape.connectorShape = connectorShape;

		const prSet = this.parentNode.getPrSet();
		if (!prSet.getPresStyleLbl()) {
			prSet.setPresStyleLbl("sibTrans2D1");
		}
		const coefficient = width / shape.cleanParams.width;
		this.applyPostAlgorithmSettings(smartartAlgorithm, connectorShape, coefficient);

		const heightScale = this.parentNode.getHeightScale(true);
		const widthScale = this.parentNode.getWidthScale(true);
		const scaleHeight = height * heightScale;
		const scaleWidth = width * widthScale;
		connectorShape.height = scaleHeight;
		connectorShape.width = scaleWidth;
		connectorShape.x += (width - scaleWidth) / 2;
		connectorShape.y += (height - scaleHeight) / 2;
	};
	ConnectorAlgorithm.prototype.createShapeConnector = function (smartartAlgorithm) {
		const connectionPoints = this.getConnectionPoints();
		const startArrowPoint = connectionPoints.start;
		const endArrowPoint = connectionPoints.end;
		if (startArrowPoint && endArrowPoint) {
			switch (this.params[AscFormat.Param_type_connRout]) {
				case AscFormat.ParameterVal_connectorRouting_longCurve:
				case AscFormat.ParameterVal_connectorRouting_curve:
					this.createCurveShapeConnector(smartartAlgorithm, startArrowPoint, endArrowPoint);
					break;
				default:
					this.createStraightShapeConnector(smartartAlgorithm, startArrowPoint, endArrowPoint);
					break;
			}
		}
	};

	ConnectorAlgorithm.prototype.createLineConnector = function () {
			const points = this.getConnectionPoints();
			if (points.start && points.end) {
				switch (this.params[AscFormat.Param_type_connRout]) {
					case AscFormat.ParameterVal_connectorRouting_stra:
						this.createStraightLineConnector(points.start, points.end);
						break;
					case AscFormat.ParameterVal_connectorRouting_bend:
						this.createBendLineConnector(points.start, points.end);
						break;
					case AscFormat.ParameterVal_connectorRouting_curve:
						this.createCurveLineConnector(points.start, points.end);
						break;
					default:
						break;
				}
			}
	};
	ConnectorAlgorithm.prototype.getStraightConnectionInfo = function (startPoint, endPoint) {
		const cx = (startPoint.x + endPoint.x) / 2;
		const cy = (startPoint.y + endPoint.y) / 2;

		const arrowVector = new CVector(endPoint.x - startPoint.x, endPoint.y - startPoint.y);

		let width;
		const connectionDistanceResolver = this.parentAlgorithm.parentNode.connectionDistanceResolver;
		const minConnectionDistance = connectionDistanceResolver && connectionDistanceResolver.getConnectionDistance();
		if (connectionDistanceResolver && minConnectionDistance !== -1) {
			width = minConnectionDistance;
		} else {
			width = arrowVector.getDistance();
		}
		const shape = this.parentNode.shape;
		const height = shape.height;

		const x = cx - width / 2;
		const y = cy - height / 2;

		const coefficient = width / shape.cleanParams.width;

		const heightScale = this.parentNode.getHeightScale(true);
		const widthScale = this.parentNode.getWidthScale(true);
		const scaleHeight = height * heightScale;
		const scaleWidth = width * widthScale;

		return {
			x: x,
			y: y,
			width: scaleWidth,
			height: scaleHeight,
			offX: (width - scaleWidth) / 2,
			offY: (height - scaleHeight) / 2,
			rot: arrowVector.getAngle(),
			coefficient: coefficient
		}
	}
	ConnectorAlgorithm.prototype.isDoubleBendHorizontalConnector = function() {
		const startPoint = this.getPointPosition(true);
		const endPoint = this.getPointPosition();
		return startPoint === AscFormat.ParameterVal_connectorPoint_midR && endPoint === AscFormat.ParameterVal_connectorPoint_midL ||
			startPoint === AscFormat.ParameterVal_connectorPoint_midL && endPoint === AscFormat.ParameterVal_connectorPoint_midR;
	};
	ConnectorAlgorithm.prototype.isDoubleBendVerticalConnector = function() {
		const startPoint = this.getPointPosition(true);
		const endPoint = this.getPointPosition();
		return startPoint === AscFormat.ParameterVal_connectorPoint_bCtr && endPoint === AscFormat.ParameterVal_connectorPoint_tCtr ||
			startPoint === AscFormat.ParameterVal_connectorPoint_tCtr && endPoint === AscFormat.ParameterVal_connectorPoint_bCtr;
	};
	ConnectorAlgorithm.prototype.isWrongBendPoints = function (startPoint, endPoint) {
		const endPointPosition = this.getPointPosition();
		const startPointPosition = this.getPointPosition(true);
		switch (endPointPosition) {
			case AscFormat.ParameterVal_connectorPoint_tCtr:
				if ((endPoint.y <= startPoint.y) ||
					(startPointPosition === AscFormat.ParameterVal_connectorPoint_midR && endPoint.x <= startPoint.x) ||
					(startPointPosition === AscFormat.ParameterVal_connectorPoint_midL && endPoint.x >= startPoint.x)) {
					return true;
				}
				break;
			case AscFormat.ParameterVal_connectorPoint_bCtr:
				if ((endPoint.y >= startPoint.y) ||
					(startPointPosition === AscFormat.ParameterVal_connectorPoint_midR && endPoint.x <= startPoint.x) ||
					(startPointPosition === AscFormat.ParameterVal_connectorPoint_midL && endPoint.x >= startPoint.x)) {
					return true;
				}
				break;
			case AscFormat.ParameterVal_connectorPoint_midR:
			case AscFormat.ParameterVal_connectorPoint_bR:
				if ((endPoint.x >= startPoint.x) ||
					((startPointPosition === AscFormat.ParameterVal_connectorPoint_bCtr ||
						startPointPosition === AscFormat.ParameterVal_connectorPoint_tCtr) && endPoint.y <= startPoint.y)) {
					return true;
				}
				break;
			case AscFormat.ParameterVal_connectorPoint_midL:
			case AscFormat.ParameterVal_connectorPoint_bL:
				if ((endPoint.x <= startPoint.x) ||
					((startPointPosition === AscFormat.ParameterVal_connectorPoint_bCtr ||
						startPointPosition === AscFormat.ParameterVal_connectorPoint_tCtr) && endPoint.y <= startPoint.y)) {
					return true;
				}
				break;
		}
		return false;
	};
	ConnectorAlgorithm.prototype.getBendPoints = function (startPoint, endPoint) {
		if (this.isWrongBendPoints(startPoint, endPoint)) {
			return [];
		}
		const isReverse = this.params[AscFormat.Param_type_bendPt] === AscFormat.ParameterVal_bendPoint_end;
		const edgePoints = this.getEdgePoints();
		let startEdgePoint = edgePoints.start;
		let endEdgePoint = edgePoints.end;
		if (isReverse) {
			const t = startPoint;
			startPoint = endPoint;
			endPoint = t;
			startEdgePoint = edgePoints.end;
			endEdgePoint = edgePoints.start;
		}
		let bendDist;
		if (this.isDoubleBendHorizontalConnector()) {
			bendDist = this.parentNode.adaptConstr[AscFormat.Constr_type_bendDist];
			if (bendDist === undefined) {
				bendDist = Math.abs(startEdgePoint.x - endEdgePoint.x) / 2;
			}
		} else if (this.isDoubleBendVerticalConnector()) {
			bendDist = this.parentNode.adaptConstr[AscFormat.Constr_type_bendDist];
			if (bendDist === undefined) {
				bendDist = Math.abs(startEdgePoint.y - endEdgePoint.y) / 2;
			}
		}
		const type = this.getPointPosition(!isReverse);
		let bendPoints;
		switch (type) {
			case AscFormat.ParameterVal_connectorPoint_auto:
				break;
			case AscFormat.ParameterVal_connectorPoint_bCtr: {
				if (bendDist === undefined) {
					bendPoints = [new CCoordPoint(startPoint.x, endPoint.y)];
				} else {
					let y = Math.min(endPoint.y, startEdgePoint.y + bendDist);
					bendPoints = [new CCoordPoint(startEdgePoint.x, y), new CCoordPoint(endEdgePoint.x, y)];
				}
				break;
			}
			case AscFormat.ParameterVal_connectorPoint_ctr:
				break;
			case AscFormat.ParameterVal_connectorPoint_bL:
			case AscFormat.ParameterVal_connectorPoint_midL: {
				if (bendDist === undefined) {
					bendPoints = [new CCoordPoint(endPoint.x, startPoint.y)];
				} else {
					let x = Math.max(endPoint.x, startEdgePoint.x - bendDist);
					bendPoints = [new CCoordPoint(x, startEdgePoint.y), new CCoordPoint(x, endEdgePoint.y)];
				}
				break;
			}
			case AscFormat.ParameterVal_connectorPoint_bR:
			case AscFormat.ParameterVal_connectorPoint_midR: {
				if (bendDist === undefined) {
					bendPoints = [new CCoordPoint(endPoint.x, startPoint.y)];
				} else {
					let x = Math.min(endPoint.x, startEdgePoint.x + bendDist);
					bendPoints = [new CCoordPoint(x, startEdgePoint.y), new CCoordPoint(x, endEdgePoint.y)];
				}
				break;
			}
			case AscFormat.ParameterVal_connectorPoint_radial:
				break;
			case AscFormat.ParameterVal_connectorPoint_tCtr: {
				if (bendDist === undefined) {
					bendPoints = [new CCoordPoint(startPoint.x, endPoint.y)];
				} else {
					let y = Math.max(endPoint.y, startEdgePoint.y - bendDist);
					bendPoints = [new CCoordPoint(startEdgePoint.x, y), new CCoordPoint(endEdgePoint.x, y)];
				}
				break;
			}
			case AscFormat.ParameterVal_connectorPoint_tL:
				break;
			case AscFormat.ParameterVal_connectorPoint_tR:
				break;
			default:
				break;
		}
		if (bendPoints) {
			return isReverse ? bendPoints.reverse() : bendPoints;
		}
	};
	ConnectorAlgorithm.prototype.getStartBendPoint = function (startPoint, endPoint) {

	};
	ConnectorAlgorithm.prototype.getEndBendPoint = function (startPoint, endPoint) {

	};
	ConnectorAlgorithm.prototype.getBendConnectionInfo = function (startPoint, endPoint) {
		const x = endPoint.x < startPoint.x ? endPoint.x : startPoint.x;
		const y = endPoint.y < startPoint.y ? endPoint.y : startPoint.y;
		const width = Math.abs(endPoint.x - startPoint.x);
		const height = Math.abs(endPoint.y - startPoint.y);
		const bendPoints = this.getBendPoints(startPoint, endPoint);
		return {
			x: x,
			y: y,
			width: width,
			height: height,
			bendPoints: bendPoints
		};
	}
	ConnectorAlgorithm.prototype.createBendLineConnector = function (startPoint, endPoint) {
		const shape = this.parentNode.getShape(false);
		shape.height = 0;
		const info = this.getBendConnectionInfo(startPoint, endPoint);
		const connectorShape = this.getTemplateConnectorLine();

		connectorShape.x = info.x;
		connectorShape.y = info.y;
		connectorShape.width = info.width;
		connectorShape.height = info.height;

		shape.connectorShape = connectorShape;

		const prSet = this.parentNode.getPrSet();
		if (!prSet.getPresStyleLbl()) {
			prSet.setPresStyleLbl("parChTrans1D2");
		}

		const bendPoints = info.bendPoints;
		const localEndPoint = new CCoordPoint(endPoint.x - connectorShape.x, endPoint.y - connectorShape.y);
		const localStartPoint = new CCoordPoint(startPoint.x - connectorShape.x, startPoint.y - connectorShape.y);
		if (bendPoints) {
			connectorShape.customGeom.push([0]);
				connectorShape.customGeom.push([1, localStartPoint.x * 36000, localStartPoint.y * 36000]);
				for (let i = 0; i < bendPoints.length; i++) {
					const bendPoint = bendPoints[i];
					const localBendPoint = new CCoordPoint(bendPoint.x - connectorShape.x, bendPoint.y - connectorShape.y);
					connectorShape.customGeom.push([2, localBendPoint.x * 36000, localBendPoint.y * 36000]);
				}
				connectorShape.customGeom.push([2, localEndPoint.x * 36000, localEndPoint.y * 36000]);

		} else {
			connectorShape.customGeom.push([0]);
			connectorShape.customGeom.push([1, localStartPoint.x * 36000, localStartPoint.y * 36000]);
			connectorShape.customGeom.push([2, localEndPoint.x * 36000, localEndPoint.y * 36000]);
		}
	};
	ConnectorAlgorithm.prototype.createCurveLineConnector = function (startPoint, endPoint) {
		const centerPoint = this.calcValues.radiusCenterPoint;
		if (centerPoint) {
			const shape = this.parentNode.getShape();
			const prSet = this.parentNode.getPrSet();
			if (!prSet.getPresStyleLbl()) {
				prSet.setPresStyleLbl("sibTrans1D1");
			}
			const startVector = centerPoint.getVector(startPoint);
			const endVector = centerPoint.getVector(endPoint);

			const radius = startVector.getDistance();
			const side = radius * 2;

			const isClockwise = this.isClockwise();
			const swAngle = isClockwise ? startVector.getAngleBetween(endVector) : -startVector.getAngleBetween(endVector, true);

			const connectorShape = this.getTemplateConnectorLine();
			connectorShape.width = side;
			connectorShape.height = side;
			connectorShape.x = centerPoint.x - radius;
			connectorShape.y = centerPoint.y - radius;

			shape.connectorShape = connectorShape;


			connectorShape.customGeom.push([0]);
			connectorShape.customGeom.push([1, (startPoint.x - connectorShape.x) * 36000, (startPoint.y - connectorShape.y) * 36000]);
			const startAngle = startVector.getAngle();
			connectorShape.customGeom.push([3, radius * 36000, radius * 36000, startAngle * radToDeg * 60000, swAngle * radToDeg * 60000]);
		}
	};
	ConnectorAlgorithm.prototype.createStraightLineConnector = function (startPoint, endPoint) {
		const shape = this.parentNode.shape;
		shape.height = 0;
		const info = this.getStraightConnectionInfo(startPoint, endPoint);
		const connectorShape = this.getTemplateConnectorLine();

		connectorShape.x = info.x;
		connectorShape.y = info.y;
		connectorShape.rot = info.rot;

		shape.connectorShape = connectorShape;

		const prSet = this.parentNode.getPrSet();
		if (!prSet.getPresStyleLbl()) {
			prSet.setPresStyleLbl("parChTrans1D2");
		}

		connectorShape.height = info.height;
		connectorShape.width = info.width;
		connectorShape.x += info.offX;
		connectorShape.y += info.offY;

		connectorShape.customGeom.push([0]);
		connectorShape.customGeom.push([1, 0, 0]);
		connectorShape.customGeom.push([2, connectorShape.width * 36000, 0]);
		connectorShape.customGeom.push([6]);
	};
	ConnectorAlgorithm.prototype.getTemplateConnectorLine = function () {
		const shape = this.parentNode.shape;
		const connectorShape = new ShadowShape();
		connectorShape.shape = shape.shape;

		connectorShape.type = AscFormat.LayoutShapeType_outputShapeType_conn;
		connectorShape.cleanParams = {};
		connectorShape.cleanParams.width = shape.cleanParams.width;
		connectorShape.cleanParams.height = shape.cleanParams.height;
		connectorShape.cleanParams.x = shape.cleanParams.x;
		connectorShape.cleanParams.y = shape.cleanParams.y;
		connectorShape.node = this.parentNode;
		if (this.params[AscFormat.Param_type_endSty] === AscFormat.ParameterVal_arrowheadStyle_arr) {
			const endArrow = new AscFormat.EndArrow();
			endArrow.type = AscFormat.LineEndType.Arrow;
			connectorShape.tailLnArrow = endArrow;
		}
		if (this.params[AscFormat.Param_type_begSty] === AscFormat.ParameterVal_arrowheadStyle_arr) {
			const endArrow = new AscFormat.EndArrow();
			endArrow.type = AscFormat.LineEndType.Arrow;
			connectorShape.headLnArrow = endArrow;
		}
		return connectorShape;
	}

	function SpaceAlgorithm() {
		BaseAlgorithm.call(this);
	}
	AscFormat.InitClassWithoutType(SpaceAlgorithm, BaseAlgorithm);

	SpaceAlgorithm.prototype.applyTextSettings = function (editorShape) {
		const node = this.parentNode.node;
		const textNodes = node.getTextNodes();

		const textNode = textNodes.textNode;
		if (textNode && textNodes.contentNode === this.parentNode) {
			textNodes.textNode.algorithm.applyTextSettings(editorShape, this.parentNode.getShape());
		}
	};
	SpaceAlgorithm.prototype.calculateShapePositions = function (smartartAlgorithm, isCalculateScaleCoefficients) {
		this.createShadowShape(isCalculateScaleCoefficients);
		if (!isCalculateScaleCoefficients) {
			this.applyPostAlgorithmSettings(smartartAlgorithm);
		}
	}
	SpaceAlgorithm.prototype.createShadowShape = function (isCalculateScaleCoefficients) {
		return this.parentNode.createShadowShape(false, isCalculateScaleCoefficients);
	};
	function TextAlgorithm() {
		BaseAlgorithm.call(this);
	}

	AscFormat.InitClassWithoutType(TextAlgorithm, BaseAlgorithm);
	TextAlgorithm.prototype.initParams = function (params) {
		BaseAlgorithm.prototype.initParams.call(this, params);
		if (this.params[AscFormat.Param_type_parTxLTRAlign] === undefined) {
			this.params[AscFormat.Param_type_parTxLTRAlign] = AscFormat.ParameterVal_horizontalAlignment_ctr;
		}
		if (this.params[AscFormat.Param_type_parTxRTLAlign] === undefined) {
			this.params[AscFormat.Param_type_parTxRTLAlign] = AscFormat.ParameterVal_horizontalAlignment_ctr;
		}
		if (this.params[AscFormat.Param_type_shpTxLTRAlignCh] === undefined) {
			this.params[AscFormat.Param_type_shpTxLTRAlignCh] = AscFormat.ParameterVal_horizontalAlignment_l;
		}
		if (this.params[AscFormat.Param_type_shpTxRTLAlignCh] === undefined) {
			this.params[AscFormat.Param_type_shpTxRTLAlignCh] = AscFormat.ParameterVal_horizontalAlignment_r;
		}
		if (this.params[AscFormat.Param_type_txAnchorVert] === undefined) {
			this.params[AscFormat.Param_type_txAnchorVert] = AscFormat.ParameterVal_textAnchorVertical_mid;
		}
		if (this.params[AscFormat.Param_type_txAnchorVertCh] === undefined) {
			this.params[AscFormat.Param_type_txAnchorVertCh] = AscFormat.ParameterVal_textAnchorVertical_t;
		}
		if (this.params[AscFormat.Param_type_autoTxRot] === undefined) {
			this.params[AscFormat.Param_type_autoTxRot] = AscFormat.ParameterVal_autoTextRotation_upr;
		}
		if (this.params[AscFormat.Param_type_stBulletLvl] === undefined) {
			this.params[AscFormat.Param_type_stBulletLvl] = 2;
		}
	};
	TextAlgorithm.prototype.applyContentFilling = function (editorShape) {
		const shapeSmartArtInfo = editorShape.getSmartArtInfo();
		const contentNodes = shapeSmartArtInfo.contentPoint;
		if (!contentNodes.length) {
			return;
		}


		const stBulletLvl = this.params[AscFormat.Param_type_stBulletLvl];
		editorShape.createTextBody();
		const shapeContent = editorShape.txBody.content;

		const arrParagraphs = [];
		let nBulletLevel = 0;

		const startDepth = contentNodes[0].depth;
		for (let i = 0; i < contentNodes.length; i += 1) {
			const contentNode = contentNodes[i];
			const mainPoint = contentNode.point;
			const deltaDepth = contentNode.depth - startDepth;
			if (contentNode.point) {
				const dataContent = mainPoint.t && mainPoint.t.content;
				if (dataContent) {
					const firstParagraph = dataContent.Content[0];
					if (firstParagraph) {
						const copyParagraph = firstParagraph.Copy(shapeContent, shapeContent.DrawingDocument);
						if (stBulletLvl === 1 || stBulletLvl === 2 && deltaDepth > 0) {
							const oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: 0, SubType: 0});
							oBullet.bulletTypeface.type = AscFormat.BULLET_TYPE_TYPEFACE_TX;

							copyParagraph.Add_PresentationNumbering(oBullet);
							copyParagraph.Set_PresentationLevel(nBulletLevel);
							nBulletLevel = Math.min(nBulletLevel + 1, 8);
							copyParagraph.Set_Ind({FirstLine: -7.9}, false);
						}
						copyParagraph.Set_Spacing({Line : 0.9, LineRule : Asc.linerule_Auto}, false);
						arrParagraphs.push(copyParagraph);
						for (let j = 1; j < dataContent.Content.length; j += 1) {
							const paragraph = dataContent.Content[j];
							const copyCurrentParagraph = paragraph.Copy();
							if (copyParagraph.Pr.Ind) {
								copyCurrentParagraph.Set_Ind({Left: copyParagraph.Pr.Ind.Left}, false);

							}
							if (copyParagraph.Pr.Spacing) {
								copyCurrentParagraph.Set_Spacing(copyParagraph.Pr.Spacing, false);
							}
							arrParagraphs.push(copyCurrentParagraph);
						}
					}
				}
			}
		}
		if (arrParagraphs.length) {
			shapeContent.Internal_Content_RemoveAll();
			for (let i = 0; i < arrParagraphs.length; i++) {
				shapeContent.AddToContent(shapeContent.Content.length, arrParagraphs[i]);
			}
		}
	};
	TextAlgorithm.prototype.applyDataBodyPr = function (editorShape) {
		const node = this.parentNode;
		const contentNode = node.contentNodes[0];
		const bodyPr = contentNode.point.t && contentNode.point.t.bodyPr;
		if (bodyPr) {
			const copyBodyPr = bodyPr.createDuplicate();
			editorShape.txBody.setBodyPr(copyBodyPr);
		}
	};
	TextAlgorithm.prototype.applyTextSettings = function (editorShape, contentShadowShape) {
		if (!contentShadowShape && this.parentNode.isTxXfrm()) {
			return;
		}
		const smartArtInfo = editorShape.getSmartArtInfo();
		this.applyFontRelations(editorShape);
		this.applyContentFilling(editorShape);
		if (smartArtInfo.contentPoint.length) {
			this.applyTxXfrmSettings(editorShape, contentShadowShape);
			this.applyDataBodyPr(editorShape);
			this.applyTextMargins(editorShape);
			this.applyHorizontalAlignment(editorShape);
			this.applyVerticalAlignment(editorShape);
			if (!editorShape.isCanFitFontSize()) {
				editorShape.applyCustTSettings();
			}
		}
	};
	TextAlgorithm.prototype.getTextRotate = function () {
		const shadowShape = this.parentNode.getShape();
		const shapeRotate = shadowShape.rot;
		switch (this.params[AscFormat.Param_type_autoTxRot]) {
			case AscFormat.ParameterVal_autoTextRotation_none:
				return 0;
			case AscFormat.ParameterVal_autoTextRotation_grav:
				if ((shapeRotate > Math.PI / 2) && (shapeRotate < Math.PI * 3 / 2)) {
					return Math.PI;
				}
				return 0;
			case AscFormat.ParameterVal_autoTextRotation_upr:
			default:
				let rot = 0;
				if (shapeRotate > Math.PI / 4) {
					rot -= Math.PI / 2;
				}
				if (shapeRotate >= Math.PI * 3 / 4) {
					rot -= Math.PI / 2;
				}
				if (shapeRotate > Math.PI * 5 / 4) {
					rot -= Math.PI / 2;
				}
				if (shapeRotate >= Math.PI * 7 / 4) {
					rot -= Math.PI / 2;
				}
				return AscFormat.normalizeRotate(rot);
		}
	};
	TextAlgorithm.prototype.applyTxXfrmSettings = function (editorShape, contentShadowShape) {
		const shadowShape = this.parentNode.getShape();
		const txXfrm = new AscFormat.CXfrm();
		const autoRot = this.getTextRotate();
		let geometry = editorShape.spPr.geometry;
		if (contentShadowShape) {
			geometry = AscFormat.ExecuteNoHistory(function(){return AscFormat.CreateGeometry(shadowShape.getEditorShapeType());},  this, []);
			shadowShape.applyAdjLst(geometry);
		}
		geometry.Recalculate(shadowShape.width, shadowShape.height);
		const geometryRect = geometry.rect;
		let extX;
		let extY;
		let offX;
		let offY;
		if (geometryRect) {
			offX = geometryRect.l + shadowShape.x;
			offY = geometryRect.t + shadowShape.y;
			extX = geometryRect.r - geometryRect.l;
			extY = geometryRect.b - geometryRect.t;
		} else {
			offX = shadowShape.x;
			offY = shadowShape.y;
			extX = shadowShape.width;
			extY = shadowShape.height;
		}
		const shapeMatrix = shadowShape.getReverseMatrix();
		const cx = shapeMatrix.TransformPointX(offX + extX / 2, offY + extY / 2);
		const cy = shapeMatrix.TransformPointY(offX + extX / 2, offY + extY / 2);
		offX = cx - extX / 2;
		offY = cy - extY / 2;

		const newSizes = AscFormat.fGetMaxInscribedRectangle(extX, extY, autoRot);
		offX += (extX - newSizes.width) / 2;
		offY += (extY - newSizes.height) / 2;
		txXfrm.setOffX(offX);
		txXfrm.setOffY(offY);
		txXfrm.setExtX(newSizes.width);
		txXfrm.setExtY(newSizes.height);
		const contentShadowShapeRot = contentShadowShape ? (shadowShape.rot - contentShadowShape.rot) : 0;
		txXfrm.setRot(AscFormat.normalizeRotate(autoRot + contentShadowShapeRot));

		editorShape.setTxXfrm(txXfrm);
	};
	TextAlgorithm.prototype.initDefaultMarginScales = function () {
		const node = this.parentNode;

		if (node.textConstraints[AscFormat.Constr_type_bMarg] === undefined) {
			const bMarg = node.getConstr(AscFormat.Constr_type_bMarg, true, true);
			if (bMarg === undefined) {
				node.textConstraints[AscFormat.Constr_type_bMarg] = 0.56;
			}
		}
		if (node.textConstraints[AscFormat.Constr_type_tMarg] === undefined) {
			const tMarg = node.getConstr(AscFormat.Constr_type_tMarg, true, true);
			if (tMarg === undefined) {
				node.textConstraints[AscFormat.Constr_type_tMarg] = 0.56;
			}
		}

		if (node.textConstraints[AscFormat.Constr_type_rMarg] === undefined) {
			const rMarg = node.getConstr(AscFormat.Constr_type_rMarg, true, true);
			if (rMarg === undefined) {
				node.textConstraints[AscFormat.Constr_type_rMarg] = 0.56;
			}
		}

		if (node.textConstraints[AscFormat.Constr_type_lMarg] === undefined) {
			const lMarg = node.getConstr(AscFormat.Constr_type_lMarg, true, true);
			if (lMarg === undefined) {
				node.textConstraints[AscFormat.Constr_type_lMarg] = 0.56;
			}
		}
	};
	TextAlgorithm.prototype.applyFontRelations = function (editorShape) {
		const node = this.parentNode;
		this.initDefaultMarginScales();

		const shapeSmartArtInfo = editorShape.getSmartArtInfo();
		shapeSmartArtInfo.textConstraintRelations = node.textConstraintRelations;
		shapeSmartArtInfo.textConstraints = node.textConstraints;
		shapeSmartArtInfo.params[AscFormat.Param_type_lnSpAfParP] = this.params[AscFormat.Param_type_lnSpAfParP];
		shapeSmartArtInfo.params[AscFormat.Param_type_lnSpAfChP] = this.params[AscFormat.Param_type_lnSpAfChP];
	};
	TextAlgorithm.prototype.applyTextMargins = function (editorShape) {
		const node = this.parentNode;
		const paddings = {};
		const isApplyInsets = node.contentNodes[0].point.isRecalculateInsets();
		if (isApplyInsets.Bottom && node.textConstraints[AscFormat.Constr_type_bMarg] === undefined) {
			paddings.Bottom = node.getConstr(AscFormat.Constr_type_bMarg, true) * g_dKoef_pt_to_mm;
		}

		if (isApplyInsets.Top && node.textConstraints[AscFormat.Constr_type_tMarg] === undefined) {
			paddings.Top = node.getConstr(AscFormat.Constr_type_tMarg, true) * g_dKoef_pt_to_mm;
		}

		if (isApplyInsets.Right && node.textConstraints[AscFormat.Constr_type_rMarg] === undefined) {
			paddings.Right = node.getConstr(AscFormat.Constr_type_rMarg, true) * g_dKoef_pt_to_mm;
		}

		if (isApplyInsets.Left && node.textConstraints[AscFormat.Constr_type_lMarg] === undefined) {
			paddings.Left = node.getConstr(AscFormat.Constr_type_lMarg, true) * g_dKoef_pt_to_mm;
		}

		editorShape.setPaddings(paddings, {bNotCopyToPoints: true});
	};
	function isParentWithChildren(nodes) {
		if (nodes.length) {
			const startDepth = nodes[0].depth;
			for (let i = 1; i < nodes.length; i += 1) {
				if (nodes[i].depth - startDepth > 0) {
					return true;
				}
			}
		}

		return false;
	}
	TextAlgorithm.prototype.applyVerticalAlignment = function (editorShape) {
		const node = this.parentNode;
		const bodyPr = editorShape.getBodyPr().createDuplicate();
		const stBulletLvl = this.params[AscFormat.Param_type_stBulletLvl];
		const firstNodeBodyPr = node.contentNodes[0].point.t.bodyPr;
		if (firstNodeBodyPr && typeof firstNodeBodyPr.anchor === "number") {
			bodyPr.setAnchor(firstNodeBodyPr.anchor);
		} else if (isParentWithChildren(node.contentNodes) || stBulletLvl === 1) {
			switch (this.params[AscFormat.Param_type_txAnchorVertCh]) {
				case AscFormat.ParameterVal_textAnchorVertical_b:
					bodyPr.setAnchor(AscFormat.VERTICAL_ANCHOR_TYPE_BOTTOM);
					break;
				case AscFormat.ParameterVal_textAnchorVertical_mid:
					bodyPr.setAnchor(AscFormat.VERTICAL_ANCHOR_TYPE_CENTER);
					break;
				case AscFormat.ParameterVal_textAnchorVertical_t:
				default:
					bodyPr.setAnchor(AscFormat.VERTICAL_ANCHOR_TYPE_TOP);
					break;
			}
		} else {
			switch (this.params[AscFormat.Param_type_txAnchorVert]) {
				case AscFormat.ParameterVal_textAnchorVertical_b:
					bodyPr.setAnchor(AscFormat.VERTICAL_ANCHOR_TYPE_BOTTOM);
					break;
				case AscFormat.ParameterVal_textAnchorVertical_t:
					bodyPr.setAnchor(AscFormat.VERTICAL_ANCHOR_TYPE_TOP);
					break;
				case AscFormat.ParameterVal_textAnchorVertical_mid:
				default:
					bodyPr.setAnchor(AscFormat.VERTICAL_ANCHOR_TYPE_CENTER);
					break;
			}
		}
		editorShape.txBody.setBodyPr(bodyPr);
	};
	TextAlgorithm.prototype.getTextAlignment = function (paramType) {
		switch (this.params[paramType]) {
			case AscFormat.ParameterVal_horizontalAlignment_l:
				return AscCommon.align_Left;
			case AscFormat.ParameterVal_horizontalAlignment_r:
				return AscCommon.align_Right;
			case AscFormat.ParameterVal_horizontalAlignment_ctr:
			default:
				return AscCommon.align_Center;
		}
	};
	TextAlgorithm.prototype.applyHorizontalAlignment = function (editorShape) {
		const drawingContent = editorShape.txBody.content;
		const node = this.parentNode;
		let paragraphLTRAlignment;
		let paragraphRTLAlignment;
		let isCtrHorzAlign;
		const stBulletLvl = this.params[AscFormat.Param_type_stBulletLvl];
		if (isParentWithChildren(node.contentNodes) || stBulletLvl === 1) {
			paragraphLTRAlignment = this.getTextAlignment(AscFormat.Param_type_shpTxLTRAlignCh);
			paragraphRTLAlignment = this.getTextAlignment(AscFormat.Param_type_shpTxRTLAlignCh);
			isCtrHorzAlign = !!this.params[AscFormat.Param_type_txAnchorHorzCh];
		} else {
			paragraphLTRAlignment = this.getTextAlignment(AscFormat.Param_type_parTxLTRAlign);
			paragraphRTLAlignment = this.getTextAlignment(AscFormat.Param_type_parTxRTLAlign);
			isCtrHorzAlign = !!this.params[AscFormat.Param_type_txAnchorHorz];
		}
		let skipSettingAlignment = false;
		for (let i = 0; i < drawingContent.Content.length; i++) {
			const item = drawingContent.Content[i];
			if (typeof item.Pr.Jc === "number") {
				skipSettingAlignment = true;
				break;
			}
		}
		if (!skipSettingAlignment) {
			for (let i = 0; i < drawingContent.Content.length; i++) {
				const item = drawingContent.Content[i];
				item.SetApplyToAll(true);
				if (item.isRtlDirection()) {
					item.SetParagraphAlign(paragraphRTLAlignment);
				} else {
					item.SetParagraphAlign(paragraphLTRAlignment);
				}
				item.SetApplyToAll(false);
			}
		}

		const bodyPr = editorShape.getBodyPr();
		const firstNodeBodyPr = node.contentNodes[0].point.t.bodyPr;
		if (firstNodeBodyPr && typeof firstNodeBodyPr.anchorCtr === "boolean") {
			const copyBodyPr = bodyPr.createDuplicate();
			bodyPr.anchorCtr = firstNodeBodyPr.anchorCtr;
			editorShape.txBody.setBodyPr(copyBodyPr);
		} else if (bodyPr.anchorCtr !== isCtrHorzAlign) {
			const copyBodyPr = bodyPr.createDuplicate();
			bodyPr.anchorCtr = isCtrHorzAlign;
			editorShape.txBody.setBodyPr(copyBodyPr);
		}
	};

	TextAlgorithm.prototype.calculateShapePositions = function (smartartAlgorithm, isCalculateScaleCoefficient) {
		this.createShadowShape(isCalculateScaleCoefficient);
		if (!isCalculateScaleCoefficient) {
			this.applyPostAlgorithmSettings(smartartAlgorithm);
		}
	};
	TextAlgorithm.prototype.createShadowShape = function (isCalculateScaleCoefficients) {
		return this.parentNode.createShadowShape(false, isCalculateScaleCoefficients);
	};

	function CompositeAlgorithm() {
		BaseAlgorithm.call(this);
	}
	AscFormat.InitClassWithoutType(CompositeAlgorithm, BaseAlgorithm);
	CompositeAlgorithm.prototype.initParams = function (params) {
		BaseAlgorithm.prototype.initParams.call(this, params);
		if (this.params[AscFormat.Param_type_horzAlign] === undefined) {
			this.params[AscFormat.Param_type_horzAlign] = AscFormat.ParameterVal_horizontalAlignment_ctr;
		}
		if (this.params[AscFormat.Param_type_vertAlign] === undefined) {
			this.params[AscFormat.Param_type_vertAlign] = AscFormat.ParameterVal_verticalAlignment_mid;
		}
	};

	CompositeAlgorithm.prototype.createShadowShape = function (isCalculateScaleCoefficients) {
		return this.parentNode.createShadowShape(true, isCalculateScaleCoefficients);
	};
	CompositeAlgorithm.prototype.calculateShapePositions = function (smartartAlgorithm, isCalculateCoefficients) {
		this.applyAlgorithmAligns(isCalculateCoefficients);
		if (!isCalculateCoefficients) {
			this.setConnections();
		}
		this.createShadowShape(isCalculateCoefficients);
	};

function PresNode(presPoint, contentNode) {
	this.parent = null;
	this.presPoint = presPoint || null;
	this.childs = [];
	this.algorithm = null;
	this.node = contentNode;
	this.contentNodes = [];
	this.layoutInfo = {
		constrLst: null,
		ruleLst: null,
		shape: null
	};

	this.shape = null;
	this.moveWithNodes = [];

	this.isHideLastTrans = true;
	this.connectionDistanceResolver = null;
	this.bounds = {
		constraints: null
	};
	this.namedNodes = null;
	this.moveWith = null;
	this._isTxXfrm = null;
	this.textConstraints = {};
	this.textConstraintRelations = [];

	this.cleanRules();
	this.cleanConstraints();
	this.cleanScales();
	this.cleanPosition();
}
	PresNode.prototype.cleanPosition = function () {
		this.position = null;
	};
PresNode.prototype.cleanScales = function () {
	this.moveScaleCoefficients = {
		width: 1,
		height: 1
	};
	this.moveRot = 0;
	this.parentScale = {};
};
PresNode.prototype.setSizesScale = function (widthCoefficient, heightCoefficient) {
	const aspectRatio = this.getAspectRatio();
	if (aspectRatio) {
		const commonCoefficient = Math.min(widthCoefficient, heightCoefficient);
		this.setWidthScale(commonCoefficient);
		this.setHeightScale(commonCoefficient);
	} else {
		this.setWidthScale(widthCoefficient);
		this.setHeightScale(heightCoefficient);
	}
};
	PresNode.prototype.initPresShape = function () {
		if (!this.layoutInfo.shape) {
			this.layoutInfo.shape = new AscFormat.SShape();
		}
	}
	PresNode.prototype.getParentScale = function (type) {
		return this.parentScale[type] === undefined ? 1 : this.parentScale[type];
	}
	PresNode.prototype.setParentScale = function (type, value) {
		if (value < 1 && (this.parentScale[type] === undefined || this.parentScale[type] > value)) {
			this.parentScale[type] = value;
		}
	}
PresNode.prototype.getDefaultConnectionNode = function() {
	if (this.childs.length) {
		return this.childs[0];
	}
	return this;
	};
	PresNode.prototype.getShape = function (isCalculateCoefficients) {
		if (isCalculateCoefficients) {
			if (!this.position) {
				this.position = new Position(this);
			}
			return this.position;
		}
		if (!this.shape) {
			this.shape = new ShadowShape(this);
		}
		return this.shape;
	};
	PresNode.prototype.getRelationConstr = function (type) {
		return this.relations[type];
	}
	PresNode.prototype.setRelationConstr = function (refNode, constr) {
		this.relations[constr.type] = {constr: constr, ref: refNode};
	};
	PresNode.prototype.setWidthScale = function (pr) {
		const relation = this.getRelationConstr(AscFormat.Constr_type_w);
		if (relation) {
			const relationConstr = relation.constr;
			const widthRef = relation.ref;
			const summaryWidthScale = this.getSummaryWidthScale();
			const newCoef = pr / summaryWidthScale;
			if (widthRef.getPresName() === this.getPresName() &&
				relationConstr.refType === AscFormat.Constr_type_h) {
				this.setParentScale(AscFormat.Constr_type_h, newCoef);
			} else {
				this.setParentScale(AscFormat.Constr_type_w, newCoef);
			}
		}
	}
	PresNode.prototype.getSummaryScale = function (scaleType) {
		let startCoefficient = 1;
		const startRelation = this.getRelationConstr(scaleType);
		if (!startRelation) {
			return startCoefficient;
		}
		let refNode = startRelation.ref;
		let relationConstr = startRelation.constr;
		const mapRelations = {};
		mapRelations[scaleType] = {};
		if (refNode && refNode.getPresName() === this.getPresName() && relationConstr.type !== relationConstr.refType) {
			mapRelations[scaleType][this.getPresName()] = true;
		}

		while (relationConstr && refNode) {
			const refType = relationConstr.refType;
			if (!mapRelations[refType]) {
				mapRelations[refType] = {};
			}
			const refPresName = refNode.getPresName();
			if (mapRelations[refType][refPresName]) {
				break;
			}
			startCoefficient *= refNode.getParentScale(refType);
			mapRelations[refType][refPresName] = true;
			const refRelation = refNode.getRelationConstr(refType);
			if (refRelation) {
				relationConstr = refRelation.constr;
				refNode = refRelation.ref;
			} else {
				break;
			}
		}
		return startCoefficient;
	}

	PresNode.prototype.getSummaryHeightScale = function () {
		return this.getSummaryScale(AscFormat.Constr_type_h);
	};

	PresNode.prototype.getSummaryWidthScale = function () {
		return this.getSummaryScale(AscFormat.Constr_type_w);
	};
	PresNode.prototype.setHeightScale = function (pr) {
		const relation = this.getRelationConstr(AscFormat.Constr_type_h);
		if (relation) {
			const relationConstr = relation.constr;
			const heightRef = relation.ref;
			const isSelfRelation = heightRef.getPresName() === this.getPresName();
			const summaryHeight = this.getSummaryHeightScale();
			const newCoef = pr / summaryHeight;
			if (isSelfRelation &&
				relationConstr.refType === AscFormat.Constr_type_w) {
				this.setParentScale(AscFormat.Constr_type_w, newCoef);
			} else {
				this.setParentScale(AscFormat.Constr_type_h, newCoef);
			}
		}
	}
PresNode.prototype.getNamedNode = function (name) {
		if (name === undefined || name === this.getPresName()) {
			return this;
		}
	if (this.namedNodes === null) {
		this.namedNodes = {};
		const childs = this.childs;
		for (let i = 0; i < childs.length; i++) {
			const child = childs[i];
			this.namedNodes[child.getPresName()] = child;
		}
	}
	return this.namedNodes[name];
};
	PresNode.prototype.isRealShapeType = function () {
		return !(this.algorithm instanceof ConnectorAlgorithm || (this.algorithm instanceof SpaceAlgorithm &&
			(this.layoutInfo.shape.type === AscFormat.LayoutShapeType_outputShapeType_conn ||
			this.layoutInfo.shape.type === AscFormat.LayoutShapeType_outputShapeType_none)));
	}
	PresNode.prototype.isSkipShape = function (isCalculateScaleCoefficient, isCycleAlgorithm) {
		let shape = this.getShape(isCalculateScaleCoefficient);
		if (shape.connectorShape) {
			shape = shape.connectorShape;
		}
		return shape.width <= 0 && shape.height <= 0 ||
			(this.algorithm instanceof ConnectorAlgorithm && !(isCycleAlgorithm && this.algorithm.isCurveShape())) ||
			this.isTxXfrm() ||
			!this.algorithm ||
			(this.algorithm instanceof SpaceAlgorithm &&
				!(this.parent && this.parent.algorithm instanceof CompositeAlgorithm) &&
				(this.layoutInfo.shape.hideGeom || !this.isRealShapeType()));
	};
	PresNode.prototype.isSibNode = function () {
		return this.node.isSibNode();
	};
	PresNode.prototype.isParNode = function () {
		return this.node.isParNode();
	};
	PresNode.prototype.isContentNode = function () {
		return this.node.isContentNode();
	};
	PresNode.prototype.getDescendantFilteredShapes = function () {
		const nodes = [this];
		const shapes = [];
		while (nodes.length) {
			const node = nodes.pop();
			shapes.push(node.getShape());

			for (let i = node.childs.length - 1; i >= 0; i -= 1) {
				nodes.push(node.childs[i]);
			}
		}
		return shapes.filter(function (shadowShape) {
			if (shadowShape.height === 0 && shadowShape.width === 0 && shadowShape.node.algorithm instanceof TextAlgorithm) {
				return false;
			}
			if (shadowShape.shape.hideGeom && shadowShape.shape.type === AscFormat.LayoutShapeType_shapeType_rect && shadowShape.node.algorithm instanceof TextAlgorithm && !shadowShape.node.isTxXfrm()) {
				return true;
			}
			return !shadowShape.shape.hideGeom && !shadowShape.node.isTxXfrm() && shadowShape.shape.type !== AscFormat.LayoutShapeType_outputShapeType_none;
		});
	}
	PresNode.prototype.getShadowShapesByZOrder = function () {
		const shapes = this.getDescendantFilteredShapes();
		const sortedShapes = shapes.slice();
		let sortedIndex = 0;
		for (let i = 0; i < shapes.length; i++) {
			const shape = shapes[i];
			const zOrderOff = shape.shape.zOrderOff;
			if (zOrderOff === 0) {
				continue;
			}
			while (sortedShapes[sortedIndex] !== shape) {
				sortedIndex += 1;
			}
			const sortedShape = sortedShapes[sortedIndex];
			sortedShapes.splice(sortedIndex, 1);
			const insertIndex = Math.min(Math.max(0, sortedIndex + zOrderOff), sortedShapes.length);
			sortedShapes.splice(insertIndex, 0, sortedShape);
		}

		return sortedShapes;
	};
	PresNode.prototype.getChildIndex = function (child) {
		for (let i = 0; i < this.childs.length; i++) {
			if (this.childs[i] === child) {
				return i;
			}
		}
	}
	PresNode.prototype.getNeighbor = function () {
		const parent = this.parent;
		const index = parent.getChildIndex(this);
		for (let i = index + 1; i < parent.childs.length; i += 1) {
			const child = parent.childs[i];
			const shape = child.shape;
			if (child.isRealShapeType()) {
				return child;
			}
		}

		for (let i = index - 1; i >= 0; i -= 1) {
			const child = parent.childs[i];
			const shape = child.shape;
			if (child.isRealShapeType()) {
				return child;
			}
		}

		return this;
	};
	PresNode.prototype.moveTo = function (deltaX, deltaY, isCalcScaleCoefficient, skipMoveWithNodes) {
		deltaX = deltaX || 0;
		deltaY = deltaY || 0;
		this.forEachDesOrSelf(function (node) {
			const shape = node.getShape(isCalcScaleCoefficient);
			if (shape) {
				shape.x += deltaX;
				shape.y += deltaY;
			}
		});
		if (!skipMoveWithNodes) {
			for (let i = 0; i < this.moveWithNodes.length; i++) {
				const moveNode = this.moveWithNodes[i];
				const moveShape = moveNode.getShape(isCalcScaleCoefficient);
				if (moveShape) {
					moveShape.x += deltaX;
					moveShape.y += deltaY;
				}
			}
			const textNodes = this.contentNodes[0] && this.contentNodes[0].getTextNodes();
			if (textNodes && textNodes.contentNode === this && textNodes.textNode) {
				const moveShape = textNodes.textNode.getShape(isCalcScaleCoefficient);
				if (moveShape) {
					moveShape.x += deltaX;
					moveShape.y += deltaY;
				}
			}
		}
		this.algorithm.moveToHierarchyOffsets(deltaX, deltaY);
	};

	PresNode.prototype.forEachChild = function (callback) {
		for (let i = 0; i < this.childs.length; i += 1) {
			callback(this.childs[i]);
		}
	};

	PresNode.prototype.forEachDes = function (callback) {
		const elements = [this];
		while (elements.length) {
			const element = elements.pop();
			for (let i = 0; i < element.childs.length; i++) {
				elements.push(element.childs[i]);
				callback(element.childs[i]);
			}
		}
	};

	PresNode.prototype.forEachDesOrSelf = function (callback) {
		callback(this);
		this.forEachDes(callback);
	};
	PresNode.prototype.getAspectRatio = function () {
		if (this.algorithm) {
			return this.algorithm.getAspectRatio();
		}
		return 0;
	}
	PresNode.prototype.getPresStyleLbl = function () {
		return this.presPoint.getPresStyleLbl();
	}
	PresNode.prototype.addMoveWithNode = function (node) {
		this.moveWithNodes.push(node);
	}
	PresNode.prototype.checkMoveWith = function () {
		for (let i = 0; i < this.childs.length; i += 1) {
			const node = this.childs[i];
			if (node.moveWith) {
				const moveNode = this.getNamedNode(node.moveWith);
				if (moveNode) {
					moveNode.addMoveWithNode(node);
				}
			}
		}
	};
PresNode.prototype.addChild = function (ch, pos) {
	if (!AscFormat.isRealNumber(pos)) {
		pos = this.childs.length;
	}
	this.childs.splice(pos, 0, ch);
	ch.parent = this;
};
	PresNode.prototype.removeChilds = function (pos, count) {
		this.childs.splice(pos, count);
	};
	PresNode.prototype.getPresName = function () {
		return this.presPoint.getPresName();
	};
	PresNode.prototype.getPrSet = function () {
		return this.presPoint.getPrSet();
	};

	PresNode.prototype.setShape = function (shape) {
		this.shape = shape;
	};

	PresNode.prototype.getNodesByAxis = function (nodes, constrType) {
		switch (constrType) {
			case AscFormat.Constr_for_self: {
				nodes.push(this);
				break;
			}
			case AscFormat.Constr_for_ch: {
				for (let i = 0; i < this.childs.length; i++) {
					nodes.push(this.childs[i]);
				}
				break;
			}
			case AscFormat.Constr_for_des: {
				const elements = [this];
				while (elements.length) {
					const element = elements.pop();
					for (let i = 0; i < element.childs.length; i++) {
						const child = element.childs[i];
						nodes.push(child);
						elements.push(child);
					}
				}
				break;
			}
			default: {
				break;
			}
		}
	};
	PresNode.prototype.setRules = function (smartartAlgorithm) {
		const ruleLst = this.layoutInfo.ruleLst;
		if (!ruleLst) {
			return;
		}
		let cacheFor = {};
		for (let i = 0; i < ruleLst.length; i++) {
			const rule = ruleLst[i];
			if (!cacheFor[rule.for]) {
				cacheFor[rule.for] = [];
				this.getNodesByAxis(cacheFor[rule.for], rule.for);
			}
			const nodes = cacheFor[rule.for];
			for (let j = 0; j < nodes.length; j++) {
				nodes[j].setRule(rule, smartartAlgorithm);
			}
		}
	};
	PresNode.prototype.getFactRule = function (constr) {
		const ruleFactor = this.factRules[constr.type];
		if (ruleFactor === undefined || ruleFactor > constr.fact) {
			return constr.fact;
		}
			return ruleFactor;
	};
	PresNode.prototype.cleanRules = function () {
		this.factRules = {};
		this.valRules = {};
	};
	PresNode.prototype.cleanConstraints = function () {
		this.relations = {};
		this.equationRelations = {
			adapt: {},
			nonAdapt: {}
		};

		this.equationRelations.nonAdapt[AscFormat.Constr_op_equ] = {};
		this.equationRelations.nonAdapt[AscFormat.Constr_op_gte] = {};
		this.equationRelations.nonAdapt[AscFormat.Constr_op_lte] = {};
		this.equationRelations.nonAdapt[AscFormat.Constr_op_none] = {};
		this.equationRelations.adapt[AscFormat.Constr_op_equ] = {};
		this.equationRelations.adapt[AscFormat.Constr_op_gte] = {};
		this.equationRelations.adapt[AscFormat.Constr_op_lte] = {};
		this.equationRelations.adapt[AscFormat.Constr_op_none] = {};
		this.adaptConstr = {};
		this.constr = {};
	};
	PresNode.prototype.setRule = function (rule, smartartAlgorithm) {
		const node = this.getConstraintNode(rule.forName, rule.ptType);
		if (node) {
			if (AscFormat.isRealNumber(rule.fact)) {
				if (rule.val !== rule.val) {
					node.factRules[rule.type] = rule.fact;
					smartartAlgorithm.setFactRuleState(factRuleState.enabled);
				}
			}
			if (typeof rule.val === "number" && rule.val === rule.val) {
				switch (rule.type) {
					case AscFormat.Constr_type_diam: {
						this.valRules[rule.type] = rule.val;
						break;
					}
					case AscFormat.Constr_type_primFontSz:
					case AscFormat.Constr_type_secFontSz: {
						if (!this.textConstraints[rule.type]) {
							this.textConstraints[rule.type] = new TextConstr();
						}
						const textConstr = this.textConstraints[rule.type];
						if (!textConstr.rule || rule.val < textConstr.rule.val) {
							textConstr.rule = rule;
						}
						break;
					}
					default:
						break;
				}
			}
		}
	};
	function TextConstr() {
		this.op = {};
		this.op[AscFormat.Constr_op_none] = [];
		this.op[AscFormat.Constr_op_equ] = [];
		this.op[AscFormat.Constr_op_gte] = [];
		this.op[AscFormat.Constr_op_lte] = [];
		this.rule = null;
	}
	TextConstr.prototype.getSecondaryFontSizeCoefficient = function () {
		const info = this.op[AscFormat.Constr_op_none];

		for (let i = 0; i < info.length; i += 1) {
			const constr = info[i].constr;
			if (constr.type === AscFormat.Constr_type_primFontSz && constr.refType === AscFormat.Constr_type_secFontSz) {
				return 1 / constr.fact;
			} else if (constr.refType === AscFormat.Constr_type_primFontSz && constr.type === AscFormat.Constr_type_secFontSz) {
				return constr.fact;
			}
		}
		return null;
	};
	TextConstr.prototype.collectShapeSmartArtInfo = function (opType, array, mapShapes, collectRefNodes) {
		const information = this.op[opType];
		for (let i = 0; i < information.length; i++) {
			const info = information[i];
			const nodes = collectRefNodes ? info.refNodes : info.nodes;
			for (let j = 0; j < nodes.length; j++) {
				const node = nodes[j];
				const editorShape = node.contentNodes[0] && node.contentNodes[0].getContentNode().getShape().editorShape;
				if (editorShape && !mapShapes[editorShape.GetId()]) {
					array.push(editorShape.getSmartArtInfo());
				}
			}
		}
	};

	TextConstr.prototype.getMaxFontSizeFromInfo = function (opType, isUseChildrenCoefficient, isReturnFirstFindFontSize) {
		const informations = this.op[opType];
		let fontSize = 65;
		//todo
		for (let i = 0; i < informations.length; i += 1) {
			const info = informations[i];
			const constr = info.constr;
			if (constr) {
				if (constr.refType === AscFormat.Constr_type_none && constr.op !== AscFormat.Constr_op_equ) {
					if (constr.val < fontSize) {
						fontSize = constr.val;
					}
					if (isReturnFirstFindFontSize) {
						break;
					}
				} else if (!(constr.for === constr.refFor && constr.forName === constr.refForName && constr.ptType === constr.refPtType)) {
					const refNodes = info.refNodes;
					for (let i = 0; i < refNodes.length; i += 1) {
						const refNode = refNodes[i];
						const editorShape = refNode.contentNodes[0] && refNode.contentNodes[0].getContentNode().getShape().editorShape;
						if (editorShape) {
							const shapeInfo = editorShape.getSmartArtInfo();
							const shapeFontSize = shapeInfo.getRelFitFontSize(isUseChildrenCoefficient);
							if (shapeFontSize !== null && shapeFontSize < fontSize) {
								fontSize = shapeFontSize;
							}
						}
					}
					if (isReturnFirstFindFontSize) {
						break;
					}
				}
			}
		}

		return fontSize;
	};
	TextConstr.prototype.getMaxFontSize = function (isUseChildrenCoefficient) {
		const noneFontSize = this.getMaxFontSizeFromInfo(AscFormat.Constr_op_none, isUseChildrenCoefficient, true);
		const lteFontSize = this.getMaxFontSizeFromInfo(AscFormat.Constr_op_lte, isUseChildrenCoefficient);
		return Math.min(noneFontSize, lteFontSize);
	};
	TextConstr.prototype.getMinFontSize = function () {
		if (this.rule) {
			return this.rule.val;
		}
		return 5;
	};
	function checkFontSizeConstraints(constr, refNodes, nodes) {
		const truthRefNodes = [];
		const truthNodes = [];
		for (let i = 0; i < nodes.length; i++) {
			const node = nodes[i].getConstraintNode(constr.forName, constr.ptType);
			if (node) {
				truthNodes.push(node);
			}
		}
		for (let i = 0; i < refNodes.length; i++) {
			const node = refNodes[i].getConstraintNode(constr.refForName, constr.refPtType);
			if (node) {
				truthRefNodes.push(node);
			}
		}
		for (let i = 0; i < truthNodes.length; i += 1) {
			const node = truthNodes[i];
			if (!node.textConstraints[constr.type]) {
				node.textConstraints[constr.type] = new TextConstr();
			}
			const calcConstr = node.textConstraints[constr.type];
			calcConstr.op[constr.op].push({constr: constr, refNodes: truthRefNodes, nodes: truthNodes});
		}
		if (constr.op === AscFormat.Constr_op_equ) {
			for (let i = 0; i < truthRefNodes.length; i++) {
				const node = truthRefNodes[i];
				if (!node.textConstraints[constr.type]) {
					node.textConstraints[constr.type] = new TextConstr();
				}
				const calcConstr = node.textConstraints[constr.type];
				calcConstr.op[constr.op].push({constr: constr, refNodes: truthNodes, nodes: truthRefNodes});
			}
		}
		if (constr.op === AscFormat.Constr_op_lte && !(constr.for === constr.refFor && constr.forName === constr.refForName && constr.ptType === constr.refPtType) && constr.refType !== AscFormat.Constr_type_none) {
			for (let i = 0; i < truthRefNodes.length; i += 1) {
				truthRefNodes[i].textConstraintRelations.push(truthNodes);
			}
		}
	}
	PresNode.prototype.setConstraints = function (isAdapt, smartartAlgorithm) {
		const constrLst = this.layoutInfo.constrLst;
		if (!constrLst) {
			return;
		}
		const valueCache = {};
		const cacheFor = {};
		for (let i = 0; i < constrLst.length; i++) {
			const constr = constrLst[i];
			if (!cacheFor[constr.for]) {
				cacheFor[constr.for] = [];
				this.getNodesByAxis(cacheFor[constr.for], constr.for);
			}
			const nodes = cacheFor[constr.for];

			if (!cacheFor[constr.refFor]) {
				cacheFor[constr.refFor] = [];
				this.getNodesByAxis(cacheFor[constr.refFor], constr.refFor);
			}
			const refNodes = cacheFor[constr.refFor];

			for (let j = 0; j < refNodes.length; j++) {
				const calcValue = refNodes[j].getCalcRefConstr(constr, isAdapt, valueCache);
				if (AscFormat.isRealNumber(calcValue)) {
					for (let k = 0; k < nodes.length; k++) {
						nodes[k].setConstraintByNode(constr, refNodes[j], calcValue, isAdapt);
					}
					break;
				}
			}
			if (isAdapt) {
				switch (constr.type) {
					case AscFormat.Constr_type_primFontSz:
					case AscFormat.Constr_type_secFontSz:
						checkFontSizeConstraints(constr, refNodes, nodes, smartartAlgorithm);
						break;
					default:
						break;
				}
			}
		}
	}
	PresNode.prototype.getPtType = function () {
		const node = this.node;
		return node.getPtType();
	}
	PresNode.prototype.checkPtType = function (elementType) {
		const ptType = this.getPtType();
		switch (elementType) {
			case AscFormat.ElementType_value_all:
				return this;
			case AscFormat.ElementType_value_sibTrans:
				if (ptType === AscFormat.Point_type_sibTrans) {
					return this;
				}
				break;
			case AscFormat.ElementType_value_parTrans:
				if (ptType === AscFormat.Point_type_parTrans) {
					return this;
				}
				break;
			case AscFormat.ElementType_value_node:
				if (ptType === AscFormat.Point_type_node || ptType === AscFormat.Point_type_asst) {
					return this;
				}
				break;
			case AscFormat.ElementType_value_asst:
				if (ptType === AscFormat.Point_type_asst) {
					return this;
				}
				break;
			case AscFormat.ElementType_value_nonAsst:
				if (ptType !== AscFormat.Point_type_asst) {
					return this;
				}
				break;
			default:
				break;
		}
	}

	PresNode.prototype.getConstraintNode = function (forName, ptType) {
		let node = this;
		if (forName) {
			node = this.checkName(forName);
		}
		return node && node.checkPtType(ptType);
	};

	PresNode.prototype.getCalcRefConstr = function (constr, isAdapt, valueCache) {
		if (!constr.refForName) {
			if (!valueCache[constr.refFor]) {
				valueCache[constr.refFor] = {};
			}
			if (!valueCache[constr.refFor][constr.refPtType]) {
				valueCache[constr.refFor][constr.refPtType] = {};
			}
			const cacheValue = valueCache[constr.refFor][constr.refPtType][constr.refType];
			if (AscFormat.isRealNumber(cacheValue)) {
				return cacheValue;
			}
		}

		const refNode = this.getConstraintNode(constr.refForName, constr.refPtType);
		if (!refNode) {
			return;
		}
		const calcValue = refNode.getRefConstr(constr, isAdapt);
		if (constr.refType !== AscFormat.Constr_type_none && !constr.refForName && constr.refFor !== AscFormat.Constr_for_self) {
			valueCache[constr.refFor][constr.refPtType][constr.refType] = calcValue;
		}
		return calcValue;
	};
	PresNode.prototype.setConstraintByNode = function (constr, refNode, calcValue, isAdapt) {
		const constrNode = this.getConstraintNode(constr.forName, constr.ptType);
		if (constrNode) {
			constrNode.addEqualRelation(refNode, constr, isAdapt);
			const isSettingConstraint = constrNode.setConstraint(constr, calcValue, isAdapt);
			constrNode.setParamConstraint(constr, refNode);
			if (isSettingConstraint) {
				constrNode.applyEqualRelations(isAdapt);
				setRelationConstraints(constrNode, refNode, constr);
			}
		}
	};

	function setRelationConstraints(constrNode, refNode, constr) {
		switch (constr.type) {
			case AscFormat.Constr_type_w:
			case AscFormat.Constr_type_h:
				constrNode.setRelationConstr(refNode, constr);
				break;
			default:
				break;
		}
	}

	PresNode.prototype.addEqualRelation = function (refNode, constr, isAdapt) {
		const relations = isAdapt ? this.equationRelations.adapt : this.equationRelations.nonAdapt;
		relations[constr.op][constr.type] = {constr: constr, ref: refNode};
	};
	PresNode.prototype.applyEqualRelations = function (isAdapt) {
		const relations = isAdapt ? this.equationRelations.adapt : this.equationRelations.nonAdapt;
		const equRelations = relations[AscFormat.Constr_op_equ];
		for (let constrType in equRelations) {
			const rel = equRelations[constrType];
			const constr = rel.constr;
			const refNode = rel.ref;

			const refConstrObject = isAdapt ? refNode.adaptConstr : refNode.constr;
			const curConstrObject = isAdapt ? this.adaptConstr : this.constr;

			let factor = this.getFactRule(constr);
			const refType = constr.refType === AscFormat.Constr_type_none && constr.refForName ? constr.type : constr.refType;
			if (refConstrObject[refType]) {
				if ( refNode === this && refConstrObject[refType] * factor !== curConstrObject[constr.type]) {
					refConstrObject[refType] = curConstrObject[constr.type];
				} else {
					curConstrObject[constr.type] = refConstrObject[refType] * factor;
				}
			}
		}
	};
	PresNode.prototype.setParamConstraint = function (constr, refNode) {
		switch (constr.type) {
			case AscFormat.Constr_type_connDist: {
				if (constr.for !== AscFormat.Constr_for_self) {
					if (this.algorithm) {
						if (!refNode.connectionDistanceResolver) {
							refNode.connectionDistanceResolver = new CConnectionDistanceResolver();
						}
						refNode.connectionDistanceResolver.addConnection(this.algorithm);
					}
				}
			}
		}
	}
	function isUserConstr(type) {
		switch (type) {
			case AscFormat.Constr_type_userA:
			case AscFormat.Constr_type_userB:
			case AscFormat.Constr_type_userC:
			case AscFormat.Constr_type_userD:
			case AscFormat.Constr_type_userE:
			case AscFormat.Constr_type_userF:
			case AscFormat.Constr_type_userG:
			case AscFormat.Constr_type_userH:
			case AscFormat.Constr_type_userI:
			case AscFormat.Constr_type_userJ:
			case AscFormat.Constr_type_userK:
			case AscFormat.Constr_type_userL:
			case AscFormat.Constr_type_userM:
			case AscFormat.Constr_type_userN:
			case AscFormat.Constr_type_userO:
			case AscFormat.Constr_type_userP:
			case AscFormat.Constr_type_userQ:
			case AscFormat.Constr_type_userR:
			case AscFormat.Constr_type_userS:
			case AscFormat.Constr_type_userT:
			case AscFormat.Constr_type_userU:
			case AscFormat.Constr_type_userV:
			case AscFormat.Constr_type_userW:
			case AscFormat.Constr_type_userX:
			case AscFormat.Constr_type_userY:
			case AscFormat.Constr_type_userZ:
				return true;
			default:
				return false;
		}
	}

	PresNode.prototype.setConstraint = function (constr, value, isAdapt) {
		if (!this.isCanAdapt(constr, isAdapt)) {
			return false;
		}
		let factor = this.getFactRule(constr);
		value *= factor;
		const constrObject = this.getConstraints(isAdapt);
		if (constrObject[constr.type] !== undefined && constr.refFor === AscFormat.Constr_for_self && constr.refType === AscFormat.Constr_type_none) {
			return false;
		}

		switch (constr.op) {
			case AscFormat.Constr_op_gte: {
				const oldValue = constrObject[constr.type];
				if (oldValue !== undefined && value < oldValue) {
					return false;
				}
				break;
			}
			case AscFormat.Constr_op_lte: {
				const oldValue = constrObject[constr.type];
				if (oldValue !== undefined && value > oldValue) {
					return false;
				}
				break;
			}
			default: {
				break;
			}
		}
		if (isAdapt) {
			switch (constr.type) {
				case AscFormat.Constr_type_w:
				case AscFormat.Constr_type_h: {
					const summaryScale = this.getSummaryScale(constr.type);
					const parentScale = Math.min(this.getParentScale(constr.type) / summaryScale, 1);
					value *= parentScale;
					break;
				}
				default:
					break;
			}
		}


		switch (constr.type) {
			case AscFormat.Constr_type_begPad:
			case AscFormat.Constr_type_endPad: {
					if (this.algorithm) {
						const padFactor = constr.refType === AscFormat.Constr_type_none ? constr.val : constr.fact;
						this.algorithm.setConnectionDistance(padFactor, constr.type === AscFormat.Constr_type_begPad);
					}
				break;
			}
			case AscFormat.Constr_type_stemThick: {
				if (this.algorithm) {
					this.algorithm.setStemThick(constr.fact);
				}
				break;
			}
			case AscFormat.Constr_type_diam:
				if (this.algorithm) {
					let val;
					if (value === 0) {
						val = constr.refType === AscFormat.Constr_type_none ? constr.val : constr.fact;
					} else {
						val = value;
					}
					this.algorithm.setConstrClockwise(val > 0);
					this.algorithm.diameterScale = constr.fact;
					break;
				}
				break;
			case AscFormat.Constr_type_bMarg:
			case AscFormat.Constr_type_tMarg:
			case AscFormat.Constr_type_rMarg:
			case AscFormat.Constr_type_lMarg:
				if (this.textConstraints[constr.type] === undefined && constrObject[constr.type] === undefined) {
					if (constr.refType === AscFormat.Constr_type_primFontSz || constr.refType === AscFormat.Constr_type_secFontSz) {
						this.textConstraints[constr.type] = constr.fact;
					} else if (AscFormat.isRealNumber(this.textConstraints[constr.refType])) {
						this.textConstraints[constr.type] = this.textConstraints[constr.refType];
					}
				} else if (constr.refType === AscFormat.Constr_type_primFontSz || constr.refType === AscFormat.Constr_type_secFontSz) {
					return;
				}
				break;
			default: {
				break;
			}
		}
		if (this.valRules[constr.type] !== Infinity) {
			constrObject[constr.type] = value;
		}
		return true;
	};
	PresNode.prototype.getParentWidth = function (isAdapt) {
		let node = this;
		while (node.isContentNode() && node.parent) {
			const parentConstrObject = isAdapt ? node.parent.adaptConstr : node.parent.constr;
			let parentWidth = parentConstrObject[AscFormat.Constr_type_w];
			let parentHeight = parentConstrObject[AscFormat.Constr_type_h];
			const aspectRatio = node.parent.getAspectRatio();
			if (aspectRatio) {
				const aspectWidth = parentHeight * aspectRatio;
				if (parentWidth > aspectWidth) {
					parentWidth = aspectWidth;
				}
			}
			if (parentWidth) {
				return parentWidth;
			}
			node = node.parent;
		}
	};
	PresNode.prototype.getParentHeight = function (isAdapt) {
		let node = this;
		while (node.isContentNode() && node.parent) {
			const parentConstrObject = isAdapt ? node.parent.adaptConstr : node.parent.constr;
			let parentWidth = parentConstrObject[AscFormat.Constr_type_w];
			let parentHeight = parentConstrObject[AscFormat.Constr_type_h];
			const aspectRatio = node.parent.getAspectRatio();
			if (aspectRatio) {
				const aspectWidth = parentHeight * aspectRatio;
				if (parentWidth <= aspectWidth) {
					parentHeight = parentWidth / aspectRatio;
				}
			}
			if (parentHeight) {
				return parentHeight;
			}
			node = node.parent;
		}
	};
	PresNode.prototype.isCanAdapt = function (constr, isAdapt) {
		if (isAdapt) {
			switch (constr.type) {
				case AscFormat.Constr_type_w:
				case AscFormat.Constr_type_h: {
					const relation = this.getRelationConstr(constr.type);
					return !!(relation && constr === relation.constr);
				}
				default:
					return true;
			}
		}
		return true;
	}
	PresNode.prototype.getRefConstr = function (constr, isAdapt) {
		let aspectRatio;
		if (!(constr.for === AscFormat.Constr_for_self && constr.refFor === AscFormat.Constr_for_self && !isUserConstr(constr.type))) {
			aspectRatio = this.getAspectRatio();
		}
		const constrObject = isAdapt ? this.adaptConstr : this.constr;
		let value;
		if (isUserConstr(constr.type) && constr.refType === AscFormat.Constr_type_none && constrObject[constr.type] !== undefined && constr.refFor === AscFormat.Constr_for_self) {
			value = constrObject[constr.type];
		} else if (constr.refType === AscFormat.Constr_type_none) {
			value = constr.val;
		} else if (constrObject[constr.refType]) {
			value = constrObject[constr.refType];
		}
		if (value !== undefined) {
			if (aspectRatio) {
				switch (constr.refType) {
					case AscFormat.Constr_type_h:
							const width = constrObject[AscFormat.Constr_type_w];
							if (width !== undefined) {
								const aspectWidth = width / aspectRatio;
								if (aspectWidth < value) {
									value = aspectWidth;
								}
							}

						break;
					case AscFormat.Constr_type_w:
							const height = constrObject[AscFormat.Constr_type_h];
							if (height !== undefined) {
								const aspectHeight = height * aspectRatio;
								if (aspectHeight < value) {
									value = aspectHeight;
								}
							}
						break;
					default:
						break;
				}
			}
		} else {
			value = this.getConstr(constr.refType, isAdapt, true);
			if (value === undefined) {
				switch (constr.refType) {
					case AscFormat.Constr_type_w: {
						const parentWidth = this.getParentWidth(isAdapt);
						constrObject[AscFormat.Constr_type_w] = parentWidth;
						return parentWidth;
					}
					case AscFormat.Constr_type_h: {
						const parentHeight = this.getParentHeight(isAdapt);
						constrObject[AscFormat.Constr_type_h] = parentHeight;
						return parentHeight;
					}
					default: {
						value = constr.val;
						break;
					}
				}
			}
		}
		return value;
	};

	PresNode.prototype.setAlgorithm = function (algorithm) {
		this.algorithm = algorithm;
	}

	PresNode.prototype.getAlgorithm = function () {
		return this.algorithm;
	}
	PresNode.prototype.getPositionConstrWithOffset = function (posConstrType, offsetConstrType, isAdapt) {
		const constrObject = this.getConstraints(isAdapt);
		const position = constrObject[posConstrType];
		let result;
		if (position !== undefined) {
			result = position;
			const offset = constrObject[offsetConstrType];
			if (offset !== undefined) {
				result += offset;
			}
		}
		return result;
	}
	PresNode.prototype.getConstrForCalculating = function (type, isAdapt) {
		switch (type) {
			case AscFormat.Constr_type_l:
				return this.getPositionConstrWithOffset(type, AscFormat.Constr_type_lOff, isAdapt);
			case AscFormat.Constr_type_r:
				return this.getPositionConstrWithOffset(type, AscFormat.Constr_type_rOff, isAdapt);
			case AscFormat.Constr_type_t:
				return this.getPositionConstrWithOffset(type, AscFormat.Constr_type_tOff, isAdapt);
			case AscFormat.Constr_type_b:
				return this.getPositionConstrWithOffset(type, AscFormat.Constr_type_bOff, isAdapt);
			case AscFormat.Constr_type_ctrX:
				return this.getPositionConstrWithOffset(type, AscFormat.Constr_type_ctrXOff, isAdapt);
			case AscFormat.Constr_type_ctrY:
				return this.getPositionConstrWithOffset(type, AscFormat.Constr_type_ctrYOff, isAdapt);
			case AscFormat.Constr_type_w:
				return this.getPositionConstrWithOffset(type, AscFormat.Constr_type_wOff, isAdapt);
			case AscFormat.Constr_type_h:
				return this.getPositionConstrWithOffset(type, AscFormat.Constr_type_hOff, isAdapt);
			default:
				const constrObj = this.getConstraints(isAdapt);
				return constrObj[type];
		}
	};
	PresNode.prototype.getConstr = function (type, isAdapt, skipDefaultValue, depth) {
		depth = depth || 0;
		const constrObj = isAdapt ? this.adaptConstr : this.constr;
		let result = this.getConstrForCalculating(type, isAdapt);
		if (depth < 2) {
			switch (type) {
				case AscFormat.Constr_type_l: {
					const width = this.getConstr(AscFormat.Constr_type_w, isAdapt, false, depth + 1);
					const right = this.getConstrForCalculating(AscFormat.Constr_type_r, isAdapt);
					const ctrX = this.getConstrForCalculating(AscFormat.Constr_type_ctrX, isAdapt);
					if (ctrX !== undefined) {
						result = ctrX - width / 2;
					} else if (right !== undefined) {
						result = right - width;
					}
					break;
				}
				case AscFormat.Constr_type_t: {
					const height = this.getConstrForCalculating(AscFormat.Constr_type_h, isAdapt);
					if (height !== undefined) {
						const ctrY = this.getConstrForCalculating(AscFormat.Constr_type_ctrY, isAdapt);
						const bottom = this.getConstrForCalculating(AscFormat.Constr_type_b, isAdapt);
						if (ctrY !== undefined) {
							result = ctrY - height / 2;
						} else if (bottom !== undefined) {
							result = bottom - height;
						}
					}
					break;
				}
				case AscFormat.Constr_type_w: {
					const right = this.getConstrForCalculating(AscFormat.Constr_type_r, isAdapt);
					const left = this.getConstrForCalculating(AscFormat.Constr_type_l, isAdapt);
					if (right !== undefined && left !== undefined) {
						result = right - left;
					} else if (this.algorithm instanceof TextAlgorithm && result === undefined) {
						const wOff = constrObj[AscFormat.Constr_type_wOff];
						result = this.getParentWidth(isAdapt);
						if (left !== undefined) {
							result -= left;
						} else if (wOff !== undefined) {
							result += wOff;
						}
					}
					break;
				}
				case AscFormat.Constr_type_h: {
					const bottom = this.getConstrForCalculating(AscFormat.Constr_type_b, isAdapt);
					const top = this.getConstr(AscFormat.Constr_type_t, isAdapt, true, depth + 1);
					if (bottom !== undefined && top !== undefined) {
						result = bottom - top;
					} else if (this.algorithm instanceof TextAlgorithm && result === undefined) {
						const hOff = constrObj[AscFormat.Constr_type_hOff];
						result = this.getParentHeight(isAdapt);
						if (top !== undefined) {
							result -= top;
						} else if (hOff !== undefined) {
							result += hOff;
						}
					}
					break;
				}
				case AscFormat.Constr_type_b: {
					const height = this.getConstrForCalculating(AscFormat.Constr_type_h, isAdapt);
					if (AscFormat.isRealNumber(height)) {
						const top = this.getConstrForCalculating(AscFormat.Constr_type_t, isAdapt);
						const ctrY = this.getConstrForCalculating(AscFormat.Constr_type_ctrY, isAdapt);
						if (AscFormat.isRealNumber(top)) {
							result = top + height;
						} else if (AscFormat.isRealNumber(ctrY)) {
							result = ctrY + height / 2;
						} else {
							result = height;
						}
					}
					break;
				}
				case AscFormat.Constr_type_r: {
					const width = this.getConstr(AscFormat.Constr_type_w, isAdapt, true, depth + 1);
					if (AscFormat.isRealNumber(width)) {
						const left = this.getConstr(AscFormat.Constr_type_l, isAdapt, false, depth + 1);
						result = left + width;
					}
					break;
				}
				case AscFormat.Constr_type_ctrY: {
					const height = this.getConstr(AscFormat.Constr_type_h, isAdapt, true, depth + 1);
					if (AscFormat.isRealNumber(height)) {
						const top = this.getConstr(AscFormat.Constr_type_t, isAdapt, false, depth + 1);
						result = top + height / 2;
					}
					break;
				}
				case AscFormat.Constr_type_ctrX: {
					const width = this.getConstr(AscFormat.Constr_type_w, isAdapt, true, depth + 1);
					if (AscFormat.isRealNumber(width)) {
						const left = this.getConstr(AscFormat.Constr_type_l, isAdapt, false, depth + 1);
						result = left + width / 2;
					}
					break;
				}
				default:
					break;
			}
		}

		return skipDefaultValue ? result : (result || 0);
	};

	PresNode.prototype.checkName = function (name) {
		if (this.getPresName() === name) {
			return this;
		}
	}

	PresNode.prototype.startAlgorithm = function (smartartAlgorithm, isCalculateScaleCoefficients) {
		if (this.algorithm) {
			this.algorithm.calculateShapePositions(smartartAlgorithm, isCalculateScaleCoefficients);
		}
	}

	PresNode.prototype.calcScaleCoefficients = function (smartartAlgorithm) {
		if (this.algorithm) {
			this.algorithm.calcScaleCoefficients(smartartAlgorithm);
		}
	}
	PresNode.prototype.setLayoutConstraints = function (lst) {
		this.layoutInfo.constrLst = lst;
	};
	PresNode.prototype.setLayoutRules = function (lst) {
		this.layoutInfo.ruleLst = lst;
	};
	PresNode.prototype.isTxXfrm = function () {
		const contentNode = this.node;
		const textNodes = contentNode.getTextNodes();
		return !!(textNodes.textNode && textNodes.contentNode && textNodes.textNode === this);
	};

	PresNode.prototype.updateCompositeSizes = function (isCalculateCoefficients) {
		const shape = this.getShape(isCalculateCoefficients);
		if (!this.childs.length) {
			shape.width = 0;
			shape.height = 0;
			return;
		}

		let shapeBounds;
		let cleanShapeBounds;
		for (let i = 0; i < this.childs.length; i += 1) {
			const node = this.childs[i];
			const childShape = node.getShape(isCalculateCoefficients);
			// todo: check this
			if (childShape.width <= 0 && childShape.height <= 0 || (node.algorithm instanceof ConnectorAlgorithm) || ((node.algorithm instanceof TextAlgorithm) && !node.isRealShapeType()) || node.isTxXfrm()) {
				continue;
			}
			if (shapeBounds) {
				checkBounds(shapeBounds, childShape.getBounds());
				checkBounds(cleanShapeBounds, childShape.getBounds(true));
			} else {
				shapeBounds = childShape.getBounds();
				cleanShapeBounds = childShape.getBounds(true);
			}
		}
		if (!shapeBounds) {
			shape.width = 0;
			shape.height = 0;
			return;
		}


			shape.x = shapeBounds.l;
			shape.y = shapeBounds.t;
			shape.cleanParams.x = cleanShapeBounds.l;
			shape.cleanParams.y = cleanShapeBounds.t;


		shape.width = shapeBounds.r - shapeBounds.l;
		shape.height = shapeBounds.b - shapeBounds.t;

		shape.cleanParams.width = cleanShapeBounds.r - cleanShapeBounds.l;
		shape.cleanParams.height = cleanShapeBounds.b - cleanShapeBounds.t;
		return shapeBounds;
	};

	PresNode.prototype.createShadowShape = function (isComposite, isCalculateCoefficients) {

		const shape = this.getShape(isCalculateCoefficients);
		shape.initFromShape(this.layoutInfo.shape);
		this.createShadowShapeFromConstraints(this.layoutInfo.shape, isCalculateCoefficients);
		if (isComposite) {
			this.forEachChild(function (node) {
				node.moveTo(shape.x, shape.y, isCalculateCoefficients, true);
			});
			this.algorithm.setConstraintSizes(shape);
			return this.updateCompositeSizes(isCalculateCoefficients);
		}
	};
	PresNode.prototype.createHierarchyRootShadowShape = function (isCalculateCoefficients) {
		const algorithm = this.algorithm;
		if (algorithm) {
			const root = algorithm.getRoot();
			const shape = this.getShape(isCalculateCoefficients);
			shape.initFromShape(this.layoutInfo.shape);
			// root.algorithm.getBounds();
			const rootShape = root.getShape(isCalculateCoefficients);
			shape.x = rootShape.x;
			shape.y = rootShape.y;
			shape.width = rootShape.width;
			shape.height = rootShape.height;
			shape.rot = rootShape.rot;
			shape.cleanParams = Object.assign({}, rootShape.cleanParams);
		}
	};
	PresNode.prototype.createHierarchyChildShadowShape = function (isCalculateCoefficients) {
		const shape = this.getShape(isCalculateCoefficients);
		shape.initFromShape(this.layoutInfo.shape);
		const childs = this.algorithm.getMainChilds();
		if (!childs.length) {
			return;
		}
		let bounds = childs[0].algorithm.getBounds(isCalculateCoefficients);
		for (let i = 1; i < childs.length; i += 1) {
			const child = childs[i];
			const algorithm = child.algorithm;
			bounds = algorithm.getBounds(isCalculateCoefficients, bounds);
		}
		if (bounds) {
			shape.x = bounds.l;
			shape.y = bounds.t;
			shape.width = bounds.r - bounds.l;
			shape.height = bounds.b - bounds.t;
		}
	};
	PresNode.prototype.getConstraints = function (isAdapt) {
		return isAdapt ? this.adaptConstr : this.constr;
	};
	PresNode.prototype.setMoveRot = function (rot) {
		this.moveRot = rot;
	}
	PresNode.prototype.createShadowShapeFromConstraints = function (layoutShape, isCalculateCoefficients) {
		const isChildInSpaceAlgorithm = this.parent && this.parent.algorithm instanceof SpaceAlgorithm;
		const constrNode = isChildInSpaceAlgorithm ? this.parent : this;

		const shape = this.getShape(isCalculateCoefficients);
		shape.rot = AscFormat.normalizeRotate(degToRad * layoutShape.rot);

		const widthCoef = this.getWidthScale();
		const heightCoef = this.getHeightScale();
		let x = constrNode.getConstr(AscFormat.Constr_type_l, !isCalculateCoefficients);
		let y = constrNode.getConstr(AscFormat.Constr_type_t, !isCalculateCoefficients);
		let width = constrNode.getConstr(AscFormat.Constr_type_w, !isCalculateCoefficients, !isChildInSpaceAlgorithm);
		let height = constrNode.getConstr(AscFormat.Constr_type_h, !isCalculateCoefficients, !isChildInSpaceAlgorithm);
		if (width === undefined && height === undefined) {
			const isSpaceAlgorithm = this.algorithm instanceof SpaceAlgorithm;
			width = isSpaceAlgorithm ? (constrNode.getParentWidth(!isCalculateCoefficients) || 0) : 0;
			height = isSpaceAlgorithm ? (constrNode.getParentHeight(!isCalculateCoefficients) || 0): 0;
		} else {
			width = width || 0;
			height = height || 0;
		}
		const newSizes = AscFormat.fGetMaxInscribedRectangle(width, height, shape.rot);
		const newScaledSizes = AscFormat.fGetMaxInscribedRectangle(width * widthCoef, height * heightCoef, shape.rot);
		x += (width - newSizes.width) / 2;
		y += (height - newSizes.height) / 2;
		width = newSizes.width;
		height = newSizes.height;
		const scaleWidth = newScaledSizes.width;
		const scaleHeight = newScaledSizes.height;
		shape.x = x;
		shape.y = y;


		shape.width = scaleWidth;
		shape.height = scaleHeight;
		const offX = (width - shape.width) / 2;
		const offY = (height - shape.height) / 2;
		shape.x += offX;
		shape.y += offY;

		let moveRot = 0;
		const prSet = this.getPrSet();

		if (prSet && prSet.custAng) {
			moveRot = prSet.custAng;
		}
		for (let i = 0; i < this.moveWithNodes.length; i += 1) {
			const moveNode = this.moveWithNodes[i];
			if (moveNode.isInitShape(isCalculateCoefficients)) {
				const moveShape = moveNode.getShape(isCalculateCoefficients);
				const scaleMoveWidth = moveShape.width * widthCoef;
				const scaleMoveHeight = moveShape.height * heightCoef;
				const moveOffX = (shape.x + (moveShape.x - x) * widthCoef) - moveShape.x;
				const moveOffY = (moveShape.height - scaleMoveHeight) / 2;
				moveShape.width = scaleMoveWidth;
				moveShape.height = scaleMoveHeight;
				moveNode.moveTo(moveOffX, moveOffY, isCalculateCoefficients);


			} else {
				moveNode.setMoveScaleCoefficients(widthCoef, heightCoef);
			}
			moveNode.setMoveRot(moveRot);
		}
		const textNodes = this.contentNodes[0] && this.contentNodes[0].getTextNodes();
		if (textNodes && textNodes.contentNode === this && textNodes.textNode) {
			if (textNodes.textNode.isInitShape(isCalculateCoefficients)) {
				const moveShape = textNodes.textNode.getShape(isCalculateCoefficients);
				const scaleMoveWidth = moveShape.width * widthCoef;
				const scaleMoveHeight = moveShape.height * heightCoef;
				const moveOffX = shape.x + shape.width / 2 + (moveShape.x + moveShape.width / 2 - (x + width / 2)) * widthCoef - (moveShape.x + scaleMoveWidth / 2);
				const moveOffY = shape.y + shape.height / 2 + (moveShape.y + moveShape.height / 2 - (y + height / 2)) * heightCoef - (moveShape.y + scaleMoveHeight / 2);
				moveShape.width = scaleMoveWidth;
				moveShape.height = scaleMoveHeight;
				textNodes.textNode.moveTo(moveOffX, moveOffY, isCalculateCoefficients);
			}
		}

		shape.cleanParams = {
			width: width,
			height: height,
			x: x,
			y: y
		};
	};
	PresNode.prototype.isInitShape = function (isCalculateScaleCoefficient) {
		return !!(isCalculateScaleCoefficient ? this.position : this.shape);
	};
	PresNode.prototype.setMoveScaleCoefficients = function (widthCoefficient, heightCoefficient) {
		this.moveScaleCoefficients.width = widthCoefficient;
		this.moveScaleCoefficients.height = heightCoefficient;
	};
	PresNode.prototype.getChildConstraintBounds = function (isCalculateCoefficients) {
		if (this.bounds.constraints === null) {
			const constrBounds = {
				l: 0,
				r: 0,
				t: 0,
				b: 0
			};
			this.bounds.constraints = constrBounds;
			if (this.childs.length) {
				const firstNodeConstraints = this.childs[0].algorithm && this.childs[0].algorithm.constraintSizes;
				if (firstNodeConstraints) {
					constrBounds.b = firstNodeConstraints.y + firstNodeConstraints.height;
					constrBounds.t = firstNodeConstraints.y;
					constrBounds.l = firstNodeConstraints.x;
					constrBounds.r = firstNodeConstraints.x + firstNodeConstraints.width;
				} else {
					const shape = this.childs[0].getShape(isCalculateCoefficients);
					constrBounds.b = shape.y + shape.height;
					constrBounds.t = shape.y;
					constrBounds.l = shape.x;
					constrBounds.r = shape.x + shape.width;
				}
				for (let i = 1; i < this.childs.length; i++) {
					this.childs[i].checkConstraintBounds(constrBounds, isCalculateCoefficients);
				}
			}
		}
		return this.bounds.constraints;
	};
	PresNode.prototype.checkConstraintBounds = function (bounds, isCalculateScaleCoefficient) {
		if (!this.isContentNode()) {
			return;
		}
		const constraintSizes = this.algorithm && this.algorithm.constraintSizes;
		if (constraintSizes) {
			if (constraintSizes.x < bounds.l) {
				bounds.l = constraintSizes.x;
			}
			if (constraintSizes.y < bounds.t) {
				bounds.t = constraintSizes.y;
			}
			const right = constraintSizes.x + constraintSizes.width;
			if (right > bounds.r) {
				bounds.r = right;
			}
			const bottom = constraintSizes.y + constraintSizes.height;
			if (bottom > bounds.b) {
				bounds.b = bottom;
			}
		} else {
			const shape = this.getShape(isCalculateScaleCoefficient);
			if (shape.x < bounds.l) {
				bounds.l = shape.x;
			}
			if (shape.y < bounds.t) {
				bounds.t = shape.y;
			}
			const right = shape.x + shape.width;
			if (right > bounds.r) {
				bounds.r = right;
			}
			const bottom = shape.y + shape.height;
			if (bottom > bounds.b) {
				bounds.b = bottom;
			}
		}
	};

	PresNode.prototype.getHeightScale = function (force) {
		if (!force && !this.isRealShapeType()) {
			return this.moveScaleCoefficients.height;
		}
		const prSet = this.getPrSet();
		if (prSet) {
			return (prSet.custScaleY || 1) * this.moveScaleCoefficients.height;
		}
		return this.moveScaleCoefficients.height;
	}

	PresNode.prototype.getWidthScale = function (force) {
		if (!force && !this.isRealShapeType()) {
			return this.moveScaleCoefficients.width;
		}
		const prSet = this.getPrSet();
		if (prSet) {
			return (prSet.custScaleX || 1) * this.moveScaleCoefficients.width;
		}
		return this.moveScaleCoefficients.width;
	}

	PresNode.prototype.initRootConstraints = function (smartArt, smartartAlgorithm) {
		const sizes = smartArt.getSizes();
		this.constr[AscFormat.Constr_type_w] = sizes.width;
		this.constr[AscFormat.Constr_type_h] = sizes.height;
		this.adaptConstr[AscFormat.Constr_type_w] = sizes.width;
		this.adaptConstr[AscFormat.Constr_type_h] = sizes.height;
	};
	PresNode.prototype.getModelId = function () {
		return this.presPoint.getModelId();
	};

function CConnectionDistanceResolver() {
	this.connectionAlgorithms = [];
	this.connectionDistance = null;
}

	CConnectionDistanceResolver.prototype.calcChildConnectionDistance = function () {
		this.connectionDistance = -1;
		const firstAlg = this.connectionAlgorithms[0];
		if (firstAlg) {
			const points = firstAlg.getConnectionPoints();
			if (points.end && points.start) {
				const v = new CVector(points.end.x - points.start.x, points.end.y - points.start.y);
				this.connectionDistance = v.getDistance();
			} else {
				return;
			}
		}
		for (let i = 1; i < this.connectionAlgorithms.length; i++) {
			const alg = this.connectionAlgorithms[i];
			const points = alg.getConnectionPoints();
			if (points.end && points.start) {
				const v = new CVector(points.end.x - points.start.x, points.end.y - points.start.y);
				const distance = v.getDistance();
				if (distance < this.connectionDistance) {
					this.connectionDistance = distance;
				}
			}
		}
	};
	CConnectionDistanceResolver.prototype.getConnectionDistance = function () {
		if (this.connectionDistance === null) {
			this.calcChildConnectionDistance();
		}
		return this.connectionDistance;
	};
	CConnectionDistanceResolver.prototype.addConnection = function (algorithm) {
		this.connectionAlgorithms.push(algorithm);
	}

	function getShapePoint(bounds) {
		return new CCoordPoint(bounds.l + (bounds.r - bounds.l) / 2, bounds.t + (bounds.b - bounds.t) / 2);
	}
	function getMinShapeEdgePoint(bounds, guideVector) {
		if (bounds.isEllipse) {
			return getMinCircleEdgePoint(bounds, guideVector);
		} else {
			return getMinRectEdgePoint(bounds, guideVector);
		}
	}
	function resolveParameterLineAndShapeEquation(ellipseBounds, paramLine) {
		const width = ellipseBounds.r - ellipseBounds.l;
		const height = ellipseBounds.b - ellipseBounds.t;
		const cw = width / 2;
		const ch = height / 2;
		const cx = cw + ellipseBounds.l;
		const cy = ch + ellipseBounds.t;

		const px = paramLine.ax;
		const py = paramLine.ay;
		const x1 = paramLine.x;
		const y1 = paramLine.y;
		const ch2 = ch * ch;
		const cw2 = cw * cw;
		const a = ch2 * px * px + cw2 * py * py;
		const b = 2 * ch2 * px * (x1 - cx) + 2 * cw2 * py * (y1 - cy);
		const c = ch2 * (cy * cy - 2 * cy * y1 + y1 * y1) + cw2 * (cx * cx - 2 * cx * x1 + x1 * x1) - cw2 * ch2;
		return AscFormat.fSolveQuadraticEquation(a, b, c);
	}

	function getMinCircleEdgePoint(bounds, guideVector) {
		const shapePoint = getShapePoint(bounds);
		const line = getParametricLinEquation(shapePoint, guideVector);
		const answer = resolveParameterLineAndShapeEquation(bounds, line);
		if (answer.bError) {
			return null;
		}
		const angle = guideVector.getAngle();

		const xt1 = line.x + line.ax * answer.x1;
		const yt1 = line.y + line.ay * answer.x1;

		let edgeAngle = new CVector(xt1 - shapePoint.x, yt1 - shapePoint.y).getAngle();
		if (AscFormat.fApproxEqual(edgeAngle, angle, algDelta)) {
			return new CCoordPoint(xt1, yt1);
		}

		const xt2 = line.x + line.ax * answer.x2;
		const yt2 = line.y + line.ay * answer.x2;

		edgeAngle = new CVector(xt2 - shapePoint.x, yt2 - shapePoint.y).getAngle();
		if (AscFormat.fApproxEqual(edgeAngle, angle, algDelta)) {
			return new CCoordPoint(xt2, yt2);
		}
	}

	function getMinRectEdgePoint(bounds, guideVector) {
		const shapePoint = getShapePoint(bounds);
		const centerAngle = guideVector.getAngle();
		let checkEdges = [
			[new CCoordPoint(bounds.l, bounds.t), new CCoordPoint(bounds.r, bounds.t)],
			[new CCoordPoint(bounds.r, bounds.t), new CCoordPoint(bounds.r, bounds.b)],
			[new CCoordPoint(bounds.l, bounds.t), new CCoordPoint(bounds.l, bounds.b)],
			[new CCoordPoint(bounds.l, bounds.b), new CCoordPoint(bounds.r, bounds.b)]
		];
		for (let i = 0; i < checkEdges.length; i += 1) {
			const edge = checkEdges[i];
			const point = getRectEdgePoint(shapePoint, guideVector, edge[0], edge[1]);
			if (point) {
				const edgeGuideVector = new CVector(point.x - shapePoint.x, point.y - shapePoint.y);
				const edgeAngle = edgeGuideVector.getAngle();
				if (AscFormat.fApproxEqual(edgeAngle, centerAngle, algDelta)) {
					return point;
				}
			}
		}
	}
	function getParametricLinEquation(startPoint, guideVector) {
		const len = guideVector.getDistance();
		return {
			x: startPoint.x,
			ax: guideVector.x / len,
			y: startPoint.y,
			ay: guideVector.y / len
		};
	}
	function getRectEdgePoint(linePoint, guideVector, rectEdgePoint1, rectEdgePoint2) {
		const line1 = getParametricLinEquation(linePoint, guideVector);
		const line2 = getParametricLinEquation(rectEdgePoint1, new CVector(rectEdgePoint2.x - rectEdgePoint1.x, rectEdgePoint2.y - rectEdgePoint1.y));
		return getIntersectionLines(line1, line2, rectEdgePoint1, rectEdgePoint2);
	}
	function getIntersectionLines(line1, line2, rectEdgePoint1, rectEdgePoint2) {
		const divider = line1.ay * line2.ax - line1.ax * line2.ay;
		if (divider === 0) {
			return null;
		}
		const parameter = (line1.ax * (line2.y - line1.y) - line1.ay * (line2.x - line1.x)) / divider;
		const x = line2.x + line2.ax * parameter;
		const y = line2.y + line2.ay * parameter;
		if (((x > rectEdgePoint1.x && x < rectEdgePoint2.x) || AscFormat.fApproxEqual(x, rectEdgePoint2.x, algDelta) || AscFormat.fApproxEqual(x, rectEdgePoint1.x, algDelta))
			&& ((y > rectEdgePoint1.y && y < rectEdgePoint2.y) || AscFormat.fApproxEqual(y, rectEdgePoint2.y, algDelta) || AscFormat.fApproxEqual(y, rectEdgePoint1.y, algDelta))) {
			return new CCoordPoint(x, y);
		}
		return null;
	}

		// document.body.addEventListener('keydown', function (e) {
		// 	if (e.ctrlKey && e.altKey && e.keyCode === 82) {
		// 			const oSM = editor.getGraphicController().selectedObjects[0];
		// 			const smartArtAlgorithm = new SmartArtAlgorithm(oSM);
		// 			smartArtAlgorithm.startFromBegin();
		// 			const drawing = oSM.spTree[0];
		// 			const shapeLength = drawing.spTree.length;
		// 			for (let i = 0; i < shapeLength; i++) {
		// 				drawing.removeFromSpTreeByPos(0);
		// 			}
		// 			const shapes = smartArtAlgorithm.getShapes();
		//
		// 			for (let i = shapes.length - 1; i >= 0; i -= 1) {
		// 				drawing.addToSpTree(0, shapes[i]);
		// 			}
		//
		// 			oSM.recalculate();
		//
		// 			editor.getLogicDocument().Recalculate();
		// 			oSM.fitFontSize();
		// 			editor.getLogicDocument().Recalculate();
		// 	}
		// });

	AscFormat.SmartArtAlgorithm = SmartArtAlgorithm;
	AscFormat.TextAlgorithm = TextAlgorithm;
})(window);


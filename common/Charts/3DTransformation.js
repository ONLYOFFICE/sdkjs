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

// Import
	let ORIENTATION_MIN_MAX = AscFormat.ORIENTATION_MIN_MAX;

	let globalBasePercent = 100;
	let global3DPersperctive = 30; // ToDo а нужна ли она в ChartsDrawer ?

	let c_oChartFloorPosition = {
		None: 0, Left: 1, Right: 2, Bottom: 3, Top: 4
	};

	/** @constructor */
	function Processor3D(width, height, left, right, bottom, top, chartSpace, chartsDrawer) {
		this.widthCanvas = width;
		this.heightCanvas = height;

		this.left = left ? left : 0;
		this.right = right ? right : 0;
		this.bottom = bottom ? bottom : 0;
		this.top = top ? top : 0;

		this.cameraDiffZ = 0;
		this.cameraDiffX = 0;
		this.cameraDiffY = 0;

		this.scaleY = 1;
		this.scaleX = 1;
		this.scaleZ = 1;

		this.aspectRatioY = 1;
		this.aspectRatioX = 1;
		this.aspectRatioZ = 1;

		this.specialStandardScaleX = 1;

		this.view3D = chartSpace.chart.getView3d();
		this.chartSpace = chartSpace;
		this.chartsDrawer = chartsDrawer;

		this.rPerspective = 0;

		this.hPercent = null;

		this.angleOx = this.view3D && this.view3D.rotX ? (-this.view3D.rotX / 360) * (Math.PI * 2) : 0;
		this.angleOy = this.view3D && this.view3D.rotY ? (-this.view3D.rotY / 360) * (Math.PI * 2) : 0;
		//this.angleOz = this.view3D && this.view3D.rotZ ? (- this.view3D.rotZ / 360) * (Math.PI * 2) : 0;

		if (!this.view3D.getRAngAx() && AscFormat.c_oChartTypes.Pie === this.chartsDrawer.calcProp.type) {
			this.angleOy = 0;
		}

		this.orientationCatAx = null;
		this.orientationValAx = null;

		this.matrixRotateOx = null;
		this.matrixRotateOy = null;
		this.matrixRotateAllAxis = null;
		this.matrixShearXY = null;
		this.projectMatrix = null;
	}

	Processor3D.prototype.calaculate3DProperties = function (baseDepth, gapDepth, bIsCheck) {
		this.calculateCommonOptions();

		//TODO baseDepth -  не универсальный параметр, позже переделать
		this._calculateAutoHPercent();

		//рассчёт коэффициэнта отношения ширины / высоты
		this._calcAspectRatio();

		//TODO рассчёт коэффицианты для диаграмм типа standard. позже необходимо отказаться
		//this._calcSpecialStandardScaleX();

		//глубина
		this.depthPerspective = this.view3D.getRAngAx() ? this._calculateDepth() : this._calculateDepthPerspective();

		//угол перспективы
		this._calculatePerspective(this.view3D);

		//после рассчета глубины меняются пропорции ширины и высоты
		if (this.view3D.getRAngAx()) {
			this._calculateScaleFromDepth();
		}

		//сдвиг камеры для того, чтобы попали все линии
		if (!bIsCheck) {
			this._calculateCameraDiff();

			if (this.view3D.getRAngAx()) {
				this._recalculateScaleWithMaxWidth();
			}
		}

		if (AscFormat.c_oChartTypes.Pie === this.chartsDrawer.calcProp.type && !this.view3D.getRAngAx()) {
			//TODO пересмотреть функцию
			this.tempChangeAspectRatioForPie();
		}
	};

	Processor3D.prototype.tempChangeAspectRatioForPie = function () {
		let perspectiveDepth = this.depthPerspective;

		let widthCanvas = this.widthCanvas;
		let originalWidthChart = widthCanvas - this.left - this.right;

		let heightCanvas = this.heightCanvas;
		let heightChart = heightCanvas - this.top - this.bottom;

		let points = [], faces = [];

		points.push(new Point3D(this.left + originalWidthChart / 2, this.top, perspectiveDepth, this));
		points.push(new Point3D(this.left, this.top, perspectiveDepth / 2, this));
		points.push(new Point3D(this.left + originalWidthChart, this.top, perspectiveDepth / 2, this));
		points.push(new Point3D(this.left + originalWidthChart / 2, this.top, 0, this));

		points.push(new Point3D(this.left + originalWidthChart / 2, this.top + heightChart, perspectiveDepth, this));
		points.push(new Point3D(this.left, this.top + heightChart, perspectiveDepth / 2, this));
		points.push(new Point3D(this.left + originalWidthChart, this.top + heightChart, perspectiveDepth / 2, this));
		points.push(new Point3D(this.left + originalWidthChart / 2, this.top + heightChart, 0, this));

		faces.push([0, 1, 2, 3]);
		faces.push([2, 5, 4, 3]);
		faces.push([1, 6, 7, 0]);
		faces.push([6, 5, 4, 7]);
		faces.push([7, 4, 3, 0]);
		faces.push([1, 6, 2, 5]);

		let minMaxOx = this._getMinMaxOx(points, faces);
		let minMaxOy = this._getMinMaxOy(points, faces);

		let kF = ((minMaxOx.right - minMaxOx.left) / originalWidthChart);
		if ((minMaxOy.bottom - minMaxOy.top) / kF > heightChart) {
			kF = ((minMaxOy.bottom - minMaxOy.top) / heightChart);
		}

		this.aspectRatioX = this.aspectRatioX * kF;
		this.aspectRatioY = this.aspectRatioY * kF;
		this.aspectRatioZ = this.aspectRatioZ * kF;
	};

Processor3D.prototype.calculateCommonOptions = function () {
	this.orientationCatAx = this.chartSpace && this.chartSpace.chart.plotArea.catAx ? this.chartSpace.chart.plotArea.catAx.getOrientation() : ORIENTATION_MIN_MAX;
	this.orientationValAx = this.chartSpace && this.chartSpace.chart.plotArea.valAx ? this.chartSpace.chart.plotArea.valAx.getOrientation() : ORIENTATION_MIN_MAX;
};

	Processor3D.prototype._calculateAutoHPercent = function () {
		let widthLine = this.widthCanvas - (this.left + this.right);
		let heightLine = this.heightCanvas - (this.top + this.bottom);

		if (this.hPercent == null) {
			this.hPercent = this.view3D.hPercent === null ? (heightLine / widthLine) : this.view3D.hPercent / 100;
			if (this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.HBar &&
				((this.view3D.hPercent === null && this.view3D.getRAngAx()) || (this.view3D.hPercent !== null && !this.view3D.getRAngAx()))) {
				this.hPercent = 1 / this.hPercent;
			}

			if (AscFormat.c_oChartTypes.Pie === this.chartsDrawer.calcProp.type) {
				this.hPercent = 0.12;
				//this.view3D.depthPercent = 180;
			}

		}
	};

	Processor3D.prototype._recalculateScaleWithMaxWidth = function () {
		let widthLine = this.widthCanvas - (this.left + this.right);
		let heightLine = this.heightCanvas - (this.top + this.bottom);
		let widthCanvas = this.widthCanvas;
		//todo оптимальную ширину нужно пересмотреть
		//оптимальная ширина -  ширина при которой не происходит масштабирования по ширине
		let optimalWidth = heightLine * 10;

		let subType = this.chartsDrawer.calcProp.subType;
		let type = this.chartsDrawer.calcProp.type;
		let isStandardType = subType === "standard" || type === AscFormat.c_oChartTypes.Line || (type === AscFormat.c_oChartTypes.Area && subType === "normal") || type ===
			AscFormat.c_oChartTypes.Surface;

		let optimalWidthLine, kF;
		if (!isStandardType) {
			this.widthCanvas = optimalWidth + (this.left + this.right);
			this._calculateScaleNStandard();
			optimalWidthLine = Math.abs(this.depthPerspective * Math.sin(-this.angleOy)) + ((this.widthCanvas - (this.left + this.right)) / this.aspectRatioX) / this.scaleX;
			kF = optimalWidthLine / widthLine;

			if (optimalWidthLine < widthLine) {
				this.widthCanvas = widthCanvas;
				this._calculateScaleNStandard();
			} else {
				this.aspectRatioX = widthLine / ((optimalWidthLine - Math.abs(this.depthPerspective * Math.sin(-this.angleOy))) / kF);
				this.scaleY = this.scaleY * kF;
				this.scaleZ = this.scaleZ * kF;

				this.widthCanvas = widthCanvas;
			}

			this._recalculateCameraDiff();
		} else {
			let scaleX = this.scaleX;
			let scaleY = this.scaleY;
			let scaleZ = this.scaleZ;

			let aspectRatioX = this.aspectRatioX;
			let aspectRatioY = this.aspectRatioY;
			let aspectRatioZ = this.aspectRatioZ;

			//если будут проблемы с поворотом standard диграмм, раскомментровать!
			//TODO протестировать, и если не будет проблем, то убрать if-else
			if (Math.abs(this.angleOy) > Math.PI) {
				//рассчитываем параметры диаграммы при оптимальной ширине
				this.widthCanvas = optimalWidth + (this.left + this.right);

				this.calaculate3DProperties(null, null, true);

				let newDepth = Math.abs(this.depthPerspective * Math.sin(-this.angleOy));
				optimalWidthLine = newDepth + ((this.widthCanvas - (this.left + this.right)) / this.aspectRatioX) / this.scaleX;
				kF = optimalWidthLine / widthLine;

				if (optimalWidthLine < widthLine) {
					this.widthCanvas = widthCanvas;
					this.scaleX = scaleX;
					this.scaleY = scaleY;
					this.scaleZ = scaleZ;

					this.aspectRatioX = aspectRatioX;
					this.aspectRatioY = aspectRatioY;
					this.aspectRatioZ = aspectRatioZ;

					return;
				}

				this.aspectRatioX = widthLine / ((optimalWidthLine - newDepth) / kF);
				this.scaleY = this.scaleY * kF;
				this.scaleZ = this.scaleZ * kF;

				this.widthCanvas = widthCanvas;

				this._recalculateCameraDiff();
			} else {
				//рассчитываем параметры диаграммы при оптимальной ширине
				this.widthCanvas = optimalWidth + (this.left + this.right);
				this.calaculate3DProperties(null, null, true);

				optimalWidthLine = this.depthPerspective * Math.sin(-this.angleOy) + ((this.widthCanvas - (this.left + this.right)) / this.aspectRatioX) / this.scaleX;

				if (optimalWidthLine < widthLine) {
					this.widthCanvas = widthCanvas;
					this.scaleX = scaleX;
					this.scaleY = scaleY;
					this.scaleZ = scaleZ;

					this.aspectRatioX = aspectRatioX;
					this.aspectRatioY = aspectRatioY;
					this.aspectRatioZ = aspectRatioZ;

					return;
				}

				kF = optimalWidthLine / widthLine;
				this.aspectRatioX = widthLine / ((optimalWidthLine - this.depthPerspective * Math.sin(-this.angleOy)) / kF);
				this.scaleY = this.scaleY * kF;
				this.scaleZ = this.scaleZ * kF;

				this.widthCanvas = widthCanvas;

				this._recalculateCameraDiff();
			}
		}
	};

	Processor3D.prototype._calculateScaleNStandard = function () {
		let seriesProps = this.chartsDrawer.calculateFirstChartCountSeries();
		let ptCount = seriesProps.points;

		let widthLine = this.widthCanvas - (this.left + this.right);
		let heightLine = this.heightCanvas - (this.top + this.bottom);

		let trueDepth = Math.abs(this.depthPerspective * Math.sin(-this.angleOx));
		let mustHeight = heightLine - trueDepth;
		let mustWidth = this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.HBar ? mustHeight * this.hPercent : mustHeight / this.hPercent;

		let areaStackedKf = this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Area && this.chartsDrawer.calcProp.subType !== "normal" ? (ptCount / ((ptCount - 1))) : 1;

		//без маштабирования
		if (this.angleOy === 0) {
			this.aspectRatioX = ((this.widthCanvas - (this.left + this.right)) / mustWidth) / areaStackedKf;
			this.scaleX = 1;
			this.scaleY = 1;
			this.aspectRatioY = heightLine / mustHeight;
		} else {
			this.aspectRatioX = ((this.widthCanvas - (this.left + this.right)) / mustWidth) / areaStackedKf;
			this.scaleX = 1;
			this.scaleY = 1;
			this.aspectRatioY = heightLine / mustHeight;
		}

		let optimalWidthLine = (widthLine / this.aspectRatioX) / this.scaleX;


		if (optimalWidthLine < widthLine) {
			return;
		}

		let kF = optimalWidthLine / (widthLine);
		this.aspectRatioX = kF * this.aspectRatioX;
		this.scaleY = this.scaleY * kF;
		this.scaleZ = this.scaleZ * kF;
	};

	Processor3D.prototype._recalculateCameraDiff = function () {
		this.cameraDiffX = 0;
		this.cameraDiffY = 0;
		this.cameraDiffZ = 0;

		this._calculateCameraDiff();
	};

	Processor3D.prototype.calculateZPositionValAxis = function () {
		let result = 0;
		if (!this.view3D.getRAngAx()) {
			result = this.orientationCatAx !== ORIENTATION_MIN_MAX ? this.depthPerspective : 0;
			if (this.chartsDrawer.calcProp.type !== AscFormat.c_oChartTypes.HBar) {
				let angleOyAbs = Math.abs(this.angleOy);
				if ((angleOyAbs >= Math.PI / 2 && angleOyAbs < Math.PI) || (angleOyAbs >= 3 * Math.PI / 2 && angleOyAbs < 2 * Math.PI)) {
					result = this.orientationCatAx !== ORIENTATION_MIN_MAX ? 0 : this.depthPerspective;
				}
			} else {
				if (this.orientationCatAx !== ORIENTATION_MIN_MAX) {
					result = Math.cos(this.angleOy) > 0 ? this.depthPerspective : 0;
				} else {
					result = Math.cos(this.angleOy) > 0 ? 0 : this.depthPerspective;
				}
			}
		} else if (this.chartsDrawer.calcProp.type !== AscFormat.c_oChartTypes.HBar && this.orientationCatAx !== ORIENTATION_MIN_MAX && this.depthPerspective !== undefined) {
			result = this.depthPerspective;
		} else if (this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.HBar && this.orientationCatAx !== ORIENTATION_MIN_MAX && this.depthPerspective !== undefined) {
			//if(this.chartSpace.chart.plotArea.valAx && this.chartSpace.chart.plotArea.valAx.yPoints && this.chartSpace.chart.plotArea.catAx.posY === this.chartSpace.chart.plotArea.valAx.yPoints[0].pos)
			result = this.depthPerspective;
		}

		return result;
	};

	Processor3D.prototype.calculateZPositionCatAxis = function () {
		let result = 0;
		if (!this.view3D.getRAngAx()) {
			if (this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.HBar && this.orientationValAx !== ORIENTATION_MIN_MAX) {
				result = Math.cos(this.angleOy) > 0 ? this.depthPerspective : 0;
			} else {
				result = Math.cos(this.angleOy) > 0 ? 0 : this.depthPerspective;
			}
		} else if (this.chartsDrawer.calcProp.type !== AscFormat.c_oChartTypes.HBar && this.orientationValAx !== ORIENTATION_MIN_MAX && this.depthPerspective !== undefined) {
			if (this.chartSpace.chart.plotArea.valAx && this.chartSpace.chart.plotArea.valAx.yPoints && this.chartSpace.chart.plotArea.catAx.posY ===
				this.chartSpace.chart.plotArea.valAx.yPoints[0].pos) {
				result = this.depthPerspective;
			}
		} else if (this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.HBar && this.orientationValAx !== ORIENTATION_MIN_MAX && this.depthPerspective !== undefined) {
			result = this.depthPerspective;
		}

		return result;
	};

	Processor3D.prototype.calculateXPositionSerAxis = function () {
		if (!this.view3D.getRAngAx()) {
			return Math.sin(this.angleOy) > 0;
		}
	};

	Processor3D.prototype.calculateFloorPosition = function () {
		let res, absOy;

		if (this.view3D.getRAngAx()) {
			if (this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.HBar) {
				absOy = Math.abs(this.angleOy);
				res = c_oChartFloorPosition.Left;
				if (absOy > Math.PI) {
					res = c_oChartFloorPosition.Right;
				}
			} else {
				res = c_oChartFloorPosition.Bottom;
				if (this.angleOx > 0) {
					res = c_oChartFloorPosition.None;
				}
			}

		} else {
			if (this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.HBar) {
				absOy = Math.abs(this.angleOy);
				res = c_oChartFloorPosition.Left;
				if (absOy > Math.PI) {
					res = c_oChartFloorPosition.Right;
				}
			} else {
				res = c_oChartFloorPosition.Bottom;
				if (this.angleOx > 0) {
					res = c_oChartFloorPosition.None;
				}
			}
		}

		return res;
	};

	//***functions for complete transformation point***
	Processor3D.prototype.convertAndTurnPoint = function (x, y, z, isNScale, isNRotate, isNProject) {
		let res = null;

		if (this.view3D.getRAngAx()) {
			res = this.convertAndTurnPointRAngAx(x, y, z);
		} else {
			res = this.convertAndTurnPointPerspective(x, y, z, isNScale, isNRotate, isNProject);
		}

		return res;
	};

	Processor3D.prototype.convertAndTurnPointRAngAx = function (x, y, z) {
		let heightChart = this.heightCanvas - this.top - this.bottom;

		let point3D = new Point3D(x, y, z, this);
		this.scale(point3D);

		//diff
		let centerXDiff = heightChart / 2 + this.left / 2;
		let centerYDiff = heightChart / 2 + this.top;
		let centerZDiff = this.depthPerspective / 2;

		point3D.offset(-centerXDiff, -centerYDiff, -centerZDiff);

		//rotate
		let matrixRotateAllAxis = this._shearXY();
		point3D.multiplyPointOnMatrix1(matrixRotateAllAxis);

		// diff camera for charts write into rect
		point3D.offset(this.cameraDiffX, this.cameraDiffY, this.cameraDiffZ);

		//undiff
		let specialReverseDiff = this.widthCanvas / 2 + (this.left - this.right) / 2;
		point3D.offset(specialReverseDiff, centerYDiff, centerZDiff);

		return {x: point3D.x, y: point3D.y, z: z};
	};

	Processor3D.prototype.convertAndTurnPointPerspective = function (x, y, z, isNScale, isNRotate, isNProject) {
		let point3D = new Point3D(x, y, z, this);

		if (!isNScale) {
			this.scale(point3D);
		}

		if (!isNRotate) {
			this.rotatePerspective(point3D);
		}

		if (!isNProject) {
			this.projectPerspective(point3D);
		}

		return {x: point3D.x, y: point3D.y, z: point3D.z};
	};

	Processor3D.prototype.scale = function (point3D) {
		//aspectRatio
		point3D.x = point3D.x / this.aspectRatioX;
		point3D.y = point3D.y / this.aspectRatioY;

		point3D.x = point3D.x / this.scaleX;
		point3D.y = point3D.y / this.scaleY;
		point3D.z = point3D.z / this.scaleZ;
	};

	Processor3D.prototype.rotatePerspective = function (point3D) {
		//diff
		point3D.offset((-this.widthCanvas / 2) / this.aspectRatioX, (-this.heightCanvas / 2) / this.aspectRatioY, 0);
		//rotate
		let matrixRotateAllAxis = this._getMatrixRotateAllAxis();
		point3D.multiplyPointOnMatrix1(matrixRotateAllAxis);

		point3D.offset((this.widthCanvas / 2) / this.aspectRatioX, (this.heightCanvas / 2) / this.aspectRatioY, 0);
	};

	Processor3D.prototype.projectPerspective = function (point3D) {
		//diff
		point3D.offset((-this.widthCanvas / 2) / this.aspectRatioX, (-this.heightCanvas / 2) / this.aspectRatioY /** aspectRatio*/, 0);

		// diff camera for charts write into rect
		point3D.offset(this.cameraDiffX, this.cameraDiffY, this.cameraDiffZ);

		//project
		let projectiveMatrix = this._getPerspectiveProjectionMatrix(1 / (this.rPerspective));
		point3D.project(projectiveMatrix);

		//undiff
		let specialReverseDiffX = this.widthCanvas / 2 + (this.left - this.right) / 2;
		let specialReverseDiffY = this.heightCanvas / 2 + (this.top - this.bottom) / 2;
		point3D.offset(specialReverseDiffX, specialReverseDiffY, 0);
	};

	//functions for step transformation point
	Processor3D.prototype.calculatePointManual = function (x, y, z, diffX, diffY, diffZ) {
		diffX = diffX !== undefined ? diffX : this.cameraDiffX;
		diffY = diffY !== undefined ? diffY : this.cameraDiffY;
		diffZ = diffZ !== undefined ? diffZ : this.cameraDiffZ;

		let diffAndScalePoints = this.diffAndScale(x, y, z);
		let x11 = diffAndScalePoints.x;
		let y11 = diffAndScalePoints.y;
		let z11 = diffAndScalePoints.z;


		let rotatePoints = this.rotate(x11, y11, z11);
		let x111 = rotatePoints.x;
		let y111 = rotatePoints.y;
		let z111 = rotatePoints.z;

		let x1111 = x111 + diffX;
		let y1111 = y111 + diffY;
		let z1111 = z111 + diffZ;


		let projectPoints = this.project(x1111, y1111, z1111);
		let x11111 = projectPoints.x;
		let y11111 = projectPoints.y;

		return {x: x11111, y: y11111};
	};

	Processor3D.prototype.diffAndScale = function (x, y, z) {
		let aRX = this.aspectRatioX;
		let aRY = this.aspectRatioY;

		let w = this.widthCanvas;
		let h = this.heightCanvas;

		let x1 = x / aRX;
		let x11 = x1 - (w / 2) / aRX;

		let z11 = z;

		let y1 = y / aRY;
		let y11 = y1 - (h / 2) / aRY;

		return {x: x11, y: y11, z: z11};
	};

	Processor3D.prototype.rotate = function (x11, y11, z11) {
		let sinOY = Math.sin(-this.angleOy);
		let cosOY = Math.cos(-this.angleOy);
		let sinOX = Math.sin(-this.angleOx);
		let cosOX = Math.cos(-this.angleOx);

		let x111 = z11 * sinOY + x11 * cosOY;
		let y111 = x11 * sinOX * sinOY + y11 * cosOX - z11 * sinOX * cosOY;
		let z111 = -x11 * sinOY * cosOX + y11 * sinOX + z11 * (cosOY * cosOX);

		return {x: x111, y: y111, z: z111};
	};

	Processor3D.prototype.project = function (x1111, y1111, z1111) {
		let fov = 1 / this.rPerspective;

		let w = this.widthCanvas;
		let h = this.heightCanvas;

		let x11111 = (fov * x1111) / (z1111 + fov) + w / 2;
		let y11111 = (fov * y1111) / (z1111 + fov) + h / 2;

		return {x: x11111, y: y11111};
	};

	//exception for pie charts
	Processor3D.prototype.convertAndTurnPointForPie = function (x, y, z) {
		let heightChart = this.heightCanvas - this.top - this.bottom;

		let point3D = new Point3D(x, y, z, this);

		//diff
		let centerXDiff = heightChart / 2 + this.left / 2;
		let centerYDiff = heightChart / 2 + this.top;
		let centerZDiff = this.depthPerspective / 2;


		point3D.offset(-centerXDiff, -centerYDiff, -centerZDiff);

		//rotate
		let matrixRotateAllAxis;
		if (!this.view3D.getRAngAx()) {
			matrixRotateAllAxis = this._getMatrixRotateAllAxis();
		} else {
			matrixRotateAllAxis = this._shearXY();
		}

		point3D.multiplyPointOnMatrix1(matrixRotateAllAxis);

		// diff camera for charts write into rect
		point3D.offset(this.cameraDiffX, this.cameraDiffY, this.cameraDiffZ);

		//project
		let projectionPoint = point3D;
		if (!this.view3D.getRAngAx()) {
			let projectiveMatrix = this._getPerspectiveProjectionMatrix(1 / (this.rPerspective));
			projectionPoint = point3D.project(projectiveMatrix);
		}

		//undiff
		let specialReverseDiff = this.widthCanvas / 2 + (this.left - this.right) / 2;
		projectionPoint.offset(specialReverseDiff, centerYDiff, centerZDiff);

		//console.log(" this.aspectRatioX: " + this.aspectRatioX + " this.aspectRatioY: " + this.aspectRatioY + " this.scaleX: " + this.scaleX + " this.scaleY: " + this.scaleY);

		return {x: projectionPoint.x, y: projectionPoint.y};
	};

	Processor3D.prototype.calculatePropertiesForPieCharts = function () {
		let sinAngleOx = Math.sin(-this.angleOx);
		let cosAngleOx = Math.cos(-this.angleOx);

		let widthCharts = this.widthCanvas - this.left - this.right;
		let heightCharts = this.heightCanvas - this.bottom - this.top;
		let radius1 = widthCharts / 2;
		let radius2 = radius1 * sinAngleOx;
		let depth = ((widthCharts + 4.89) / 8.37) * cosAngleOx;

		if ((radius2 * 2 + depth) > heightCharts) {
			let kF = (radius2 * 2 + depth) / heightCharts;
			radius1 = radius1 / kF;
			radius2 = radius2 / kF;
			depth = depth / kF;
		}

		//ToDo временная правка для круговой диграммы с углом поворота 0 градусов
		if (0 === radius2) {
			radius2 = 1;
		}

		return {radius1: radius1, radius2: radius2, depth: depth};
	};


	//***functions for matrix transformation***
	Processor3D.prototype._shearXY = function () {
		if (null === this.matrixShearXY) {
			this.matrixShearXY = new Matrix4D();

			this.matrixShearXY.a31 = Math.sin(-this.angleOy);
			this.matrixShearXY.a32 = Math.sin(this.angleOx);
			this.matrixShearXY.a33 = 0;
			this.matrixShearXY.a44 = 0;
		}

		return this.matrixShearXY;
	};
	Processor3D.prototype._shearXY2 = function () {
		if (null === this.matrixShearXY) {
			this.matrixShearXY = [[1, 0, 0, 0], [0, 1, 0, 0], [Math.sin(-this.angleOy), Math.sin(this.angleOx), 0, 0], [0, 0, 0, 0]];
		}
		return this.matrixShearXY;
	};

	Processor3D.prototype._getMatrixRotateAllAxis = function () {
		let matrixRotateOY = this._getMatrixRotateOy();
		let matrixRotateOX = this._getMatrixRotateOx();

		/*result matrix

		|cosOy               0           sinOy      0|
		|sinOx * sinOy     cosOx    -sinOx * cosOy  0|
		|-sinOy * cosOx     sinOx     cosOy * cosOx  0|
		|-sinOy              0        (cosOy + 1)   1|*/

		if (null === this.matrixRotateAllAxis) {
			this.matrixRotateAllAxis = matrixRotateOY.multiply(matrixRotateOX);
		}

		return this.matrixRotateAllAxis;
	};

	Processor3D.prototype._getMatrixRotateOx = function () {
		//todo array -> Float32Array ?
		if (null === this.matrixRotateOx) {
			this.matrixRotateOx =
				new Matrix4D()/*[[1, 0, 0, 0], [0, Math.cos(-this.angleOx), Math.sin(-this.angleOx), 0], [0, - Math.sin(-this.angleOx), Math.cos(-this.angleOx), 1], [0, 0, 0, 1]]*/;

			let cos = Math.cos(-this.angleOx);
			let sin = Math.sin(-this.angleOx);

			this.matrixRotateOx.a22 = cos;
			this.matrixRotateOx.a23 = sin;
			this.matrixRotateOx.a32 = -sin;
			this.matrixRotateOx.a33 = cos;
			this.matrixRotateOx.a34 = 1;
		}

		return this.matrixRotateOx;
	};

	Processor3D.prototype._getMatrixRotateOy = function () {
		if (null === this.matrixRotateOy) {
			this.matrixRotateOy =
				new Matrix4D()/*[[Math.cos(-this.angleOy), 0, -Math.sin(-this.angleOy), 0], [0, 1, 0, 0], [Math.sin(-this.angleOy), 0, Math.cos(-this.angleOy), 1], [0, 0, 0, 1]]*/;

			let cos = Math.cos(-this.angleOy);
			let sin = Math.sin(-this.angleOy);

			this.matrixRotateOy.a11 = cos;
			this.matrixRotateOy.a13 = -sin;
			this.matrixRotateOy.a31 = sin;
			this.matrixRotateOy.a33 = cos;
			this.matrixRotateOy.a34 = 1;
		}

		return this.matrixRotateOy;
	};

	Processor3D.prototype._getMatrixRotateOz = function () {
		return [[1, 0, 0, 0], [0, 1, 0, 0], [0, 0, 0, 1], [0, 0, 0, 1]];
	};

	Processor3D.prototype._getPerspectiveProjectionMatrix = function (fov) {
		/*let zf = this.rPerspective + this.depthPerspective;
		let zn = this.rPerspective;
		let q = zf / (zf - zn);
		return [[1 / Math.tan(this.rPerspective / 2), 0, 0, 0], [0, 1 / Math.tan(this.rPerspective / 2), 0, 0], [0, 0, q, 1], [0, 0, -q * zn, 0]];*/
		//[[1, 0, 0, 0], [0, 1, 0, 0], [0, 0, 0, 1 / fov], [0, 0, 0, 1]]
		if (null === this.projectMatrix) {
			this.projectMatrix = new Matrix4D();
			this.projectMatrix.a33 = 0;
			this.projectMatrix.a34 = 1 / fov;
		}

		return this.projectMatrix;
	};

	Processor3D.prototype.correctPointsPosition = function (chartSpace) {
		if (this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Pie) {
			return;
		}

		let pxToMM = 1 / AscCommon.g_dKoef_pix_to_mm;
		let t = this;
		//коррективы для подписей
		let xPoints = chartSpace.chart.plotArea && chartSpace.chart.plotArea.catAx ? chartSpace.chart.plotArea.catAx.xPoints : null;
		if (!xPoints) {
			xPoints = chartSpace.chart.plotArea && chartSpace.chart.plotArea.valAx ? chartSpace.chart.plotArea.valAx.xPoints : null;
		}

		let coordYAxisOx = chartSpace.chart.plotArea.catAx && chartSpace.chart.plotArea.catAx.posY;
		if (!coordYAxisOx && chartSpace.chart.plotArea.valAx) {
			coordYAxisOx = chartSpace.chart.plotArea.valAx.posY != undefined ? chartSpace.chart.plotArea.valAx.posY : chartSpace.chart.plotArea.valAx.posY;
		}

		let yPoints = chartSpace.chart.plotArea && chartSpace.chart.plotArea.valAx ? chartSpace.chart.plotArea.valAx.yPoints : null;
		if (!yPoints) {
			yPoints = chartSpace.chart.plotArea && chartSpace.chart.plotArea.catAx ? chartSpace.chart.plotArea.catAx.yPoints : null;
		}

		let coordXAxisOy;
		if (chartSpace.chart.plotArea.catAx) {
			coordXAxisOy = chartSpace.chart.plotArea.catAx.posX ? chartSpace.chart.plotArea.catAx.posX : chartSpace.chart.plotArea.catAx.xPos;
		}
		if (!coordXAxisOy && chartSpace.chart.plotArea.valAx) {
			coordXAxisOy = chartSpace.chart.plotArea.valAx.posX != undefined ? chartSpace.chart.plotArea.valAx.posX : null;
		}

		let correctPointsOx = function () {
			let z = t.calculateZPositionCatAxis();
			let valCatAx = chartSpace.chart.plotArea.catAx;
			if (!valCatAx.transformXPoints) {
				valCatAx.transformXPoints = [];
			}

			for (let i = 0; i < xPoints.length; i++) {
				let widthText = 0;
				if (valCatAx && valCatAx.labels && t.orientationValAx !== ORIENTATION_MIN_MAX) {
					widthText = valCatAx.labels.extY * pxToMM;
				}

				let point = t.convertAndTurnPoint(xPoints[i].pos * pxToMM, coordYAxisOx * pxToMM - widthText, z);

				valCatAx.transformXPoints[i] = {x: point.x / pxToMM, y: point.y / pxToMM};
			}
		};

		let correctPointsOxHBar = function () {
			let z = t.calculateZPositionValAxis();
			let valCatAx = chartSpace.chart.plotArea.valAx;
			if (!valCatAx.transformXPoints) {
				valCatAx.transformXPoints = [];
			}

			for (let i = 0; i < xPoints.length; i++) {
				let widthText = 0;
				if (t.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.HBar && valCatAx && valCatAx.labels && t.orientationCatAx !== ORIENTATION_MIN_MAX) {
					widthText = valCatAx.labels.extY * pxToMM;
				}

				let point = t.convertAndTurnPoint(xPoints[i].pos * pxToMM, coordYAxisOx * pxToMM - widthText, z);

				valCatAx.transformXPoints[i] = {x: point.x / pxToMM, y: point.y / pxToMM};
			}
		};

		let correctPointsOy = function () {
			let zPosition = t.calculateZPositionValAxis();
			let valCatAx = chartSpace.chart.plotArea.valAx;

			if (!valCatAx.transformYPoints) {
				valCatAx.transformYPoints = [];
			}

			for (let i = 0; i < yPoints.length; i++) {
				let point = t.convertAndTurnPoint(coordXAxisOy * pxToMM, yPoints[i].pos * (pxToMM), zPosition);

				//TODO значения высчитать
				let widthText = 5;
				if (valCatAx && valCatAx.labels) {
					widthText = valCatAx.labels.extX * pxToMM;
				}

				if (t.orientationCatAx !== ORIENTATION_MIN_MAX) {
					widthText = 0;
				}

				let diffXText = 0;
				let angleOyAbs = Math.abs(t.angleOy);
				if (!t.view3D.getRAngAx() && (angleOyAbs >= Math.PI / 2 && angleOyAbs < 3 * Math.PI / 2)) {
					diffXText = -diffXText;
				}

				valCatAx.transformYPoints[i] = {
					x: (point.x - (diffXText + widthText)) / pxToMM, y: point.y / pxToMM
				};
			}
		};

		let correctPointsOyHBar = function () {
			let zPosition = t.calculateZPositionCatAxis();
			let valCatAx = chartSpace.chart.plotArea.catAx;
			if (!valCatAx.transformYPoints) {
				valCatAx.transformYPoints = [];
			}

			for (let i = 0; i < yPoints.length; i++) {
				let point = t.convertAndTurnPoint(coordXAxisOy * pxToMM, yPoints[i].pos * (pxToMM), zPosition);

				//TODO значения высчитать
				let widthText = 0;
				if (valCatAx && valCatAx.labels) {
					widthText = valCatAx.labels.extX * pxToMM;
				}

				if (t.orientationValAx !== ORIENTATION_MIN_MAX) {
					widthText -= 10;
				} else {
					widthText += 5;
				}

				valCatAx.transformYPoints[i] = {x: (point.x - widthText) / pxToMM, y: point.y / pxToMM};
			}
		};


		if (xPoints) {
			if (this.chartsDrawer.calcProp.type !== AscFormat.c_oChartTypes.HBar) {
				correctPointsOx(xPoints);
			} else {
				correctPointsOxHBar(xPoints);
			}
		}

		if (yPoints) {
			if (this.chartsDrawer.calcProp.type !== AscFormat.c_oChartTypes.HBar) {
				correctPointsOy(yPoints);
			} else {
				correctPointsOyHBar(yPoints);
			}
		}
	};

	Processor3D.prototype._calculatePerspective = function (view3D) {
		let heightLine = this.heightCanvas - (this.top + this.bottom);

		let perspective = view3D && view3D.perspective ? view3D.perspective : global3DPersperctive;
		if (view3D && 0 === view3D.perspective && this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Surface) {
			perspective = 1;
		}

		let alpha = perspective / 4;//в xml проиходит двойной угол(в параметрах ms стоит 40, приходит в xml 80)
		//TODO this.top - this.bottom пересмотреть
		let catt = ((heightLine / 2 + (Math.abs(this.top - this.bottom)))) / Math.tan((alpha / 360) * (Math.PI * 2));
		let rPerspective;
		if (catt === 0) {
			rPerspective = 0;
		} else {
			rPerspective = 1 / catt;
		}

		this.rPerspective = rPerspective;
	};

	Processor3D.prototype._calculateDepth = function () {
		let widthOriginalChart = this.widthCanvas - (this.left + this.right);
		let heightOriginalChart = this.heightCanvas - (this.top + this.bottom);
		let subType = this.chartsDrawer.calcProp.subType;
		let type = this.chartsDrawer.calcProp.type;
		let defaultOverlap = (subType === "stacked" || subType === "stackedPer" || subType === "standard" || type === AscFormat.c_oChartTypes.Line || type ===
			AscFormat.c_oChartTypes.Area || type === AscFormat.c_oChartTypes.Surface) ? 100 : 0;
		let overlap = AscFormat.isRealNumber(this.chartSpace.chart.plotArea.chart.overlap) ? (this.chartSpace.chart.plotArea.chart.overlap / 100) : (defaultOverlap / 100);
		let gapWidth = this.chartSpace.chart.plotArea.chart.gapWidth != null ? (this.chartSpace.chart.plotArea.chart.gapWidth / 100) : (150 / 100);
		let gapDepth = this.chartSpace.chart.plotArea.chart.gapDepth != null ? (this.chartSpace.chart.plotArea.chart.gapDepth / 100) :
			type === AscFormat.c_oChartTypes.Area && subType !== "normal" ? 1 : (150 / 100);

		let seriesProps = this.chartsDrawer.calculateFirstChartCountSeries();
		let seriesCount = seriesProps.series;
		let ptCount = seriesProps.points;

		let sinOx = Math.abs(Math.sin(-this.angleOx));
		let sinOy = Math.sin(-this.angleOy);
		let hPercent = type === AscFormat.c_oChartTypes.HBar ? 1 : this.hPercent;
		let depthPercent = this._getDepthPercent();
		let t = this;

		let areaStackedKf = (type === AscFormat.c_oChartTypes.Area && subType !== "normal") || type === AscFormat.c_oChartTypes.Surface ? (ptCount / (2 * (ptCount - 1))) : 1;

		let depth = 0;
		let chartWidth = 0;

		let standardType = false;
		if (subType === "standard" || type === AscFormat.c_oChartTypes.Line || (type === AscFormat.c_oChartTypes.Area && subType === "normal") ||
			(type === AscFormat.c_oChartTypes.Surface)) {
			standardType = true;
		}

		let heightHPercent = heightOriginalChart / hPercent;
		let angleOxKf;
		if (!standardType) {
			let widthOneBar = ((heightHPercent / seriesCount) / (ptCount - (ptCount - 1) * (overlap) + gapWidth)) * sinOy;
			let a, b;

			if (this.angleOx === 0 && this.angleOy === 0)//withoutAngleNoAuto + withoutAngleAuto
			{
				chartWidth = widthOneBar + heightHPercent;
			} else if (this.angleOx !== 0)//AngleOYNoAut + AngleOYNoAutPerHeight + (ANGLEOX+ANGLEOY) + AngleOYOXNoAut + ANGLEOXANGLEOYHPerDPer(ANGLEOX+ANGLEOY HPercent)
			{
				//если выставить ширину 255 будет так же, как и в документе с расчётами
				b = (seriesCount - (seriesCount - 1) * overlap + gapWidth);
				a = (depthPercent / (ptCount * b)) / hPercent;
				let width = heightOriginalChart * areaStackedKf;
				depth = (width * a + gapDepth * width * a) / (1 / sinOx + (gapDepth) * a + a);

				chartWidth = heightHPercent - depth;
			} else if (this.angleOy !== 0)//angleOxNoAuto
			{
				//если выставить ширину = 321.25 будет так же, как и в документе с расчётам
				//TODO глубина с некоторыми графиками имеет различия, пересчитать!
				let widthChart = (widthOriginalChart / t.aspectRatioX) / t.specialStandardScaleX;

				b = (seriesCount - (seriesCount - 1) * overlap + gapWidth);
				if (subType === "standard" || type === AscFormat.c_oChartTypes.Line || type === AscFormat.c_oChartTypes.Area || type === AscFormat.c_oChartTypes.Surface) {
					b = b / seriesCount;
				}

				angleOxKf = sinOx === 0 ? 1 : sinOx;
				a = depthPercent / (ptCount * b);
				depth = (widthChart * a + gapDepth * widthChart * a) / (1 / angleOxKf + (gapDepth) * a + a);

				depth = depth / angleOxKf;
			}
		} else//allStandardDepth
		{
			angleOxKf = sinOx === 0 ? 0 : sinOx;

			if (type === AscFormat.c_oChartTypes.Area) {
				depth =
					(depthPercent / (angleOxKf * depthPercent + ((ptCount + (Math.floor((seriesCount - ptCount) / 2 - 0.5))) / seriesCount * hPercent))) * (heightOriginalChart);
			} else {
				depth = (depthPercent / (angleOxKf * depthPercent + ((ptCount + (Math.floor((seriesCount - ptCount) / 2))) / seriesCount * hPercent))) * (heightOriginalChart);
			}

			if ((this.angleOx !== 0)) {
				depth = depth * Math.sin(-this.angleOx);
			}
		}

		return sinOx !== 0 ? Math.abs(depth / sinOx) : Math.abs(depth);
	};

	Processor3D.prototype._calculateDepthPerspective = function () {
		let widthCanvas = this.widthCanvas;
		let widthOriginalChart = widthCanvas - (this.left + this.right);

		let aspectRatio = this.aspectRatioX / (this.specialStandardScaleX);
		let widthChart = widthOriginalChart / aspectRatio;
		widthChart = widthChart / this.scaleX;

		let countSeries = this.chartsDrawer.calculateFirstChartCountSeries();
		let seriesCount = countSeries.series;
		let ptCount = countSeries.points;

		let width = widthChart / ptCount;

		let isNormalArea = (this.chartsDrawer.calcProp.subType === "normal" && this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Area) ||
			this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Surface;

		let defaultOverlap = (this.chartsDrawer.calcProp.subType === "stacked" || this.chartsDrawer.calcProp.subType === "stackedPer" || this.chartsDrawer.calcProp.subType ===
			"standard" || this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Line || isNormalArea) ? 100 : 0;
		let overlap = AscFormat.isRealNumber(this.chartSpace.chart.plotArea.chart.overlap) ? (this.chartSpace.chart.plotArea.chart.overlap / 100) : (defaultOverlap / 100);

		let gapWidth = this.chartSpace.chart.plotArea.chart.gapWidth != null ? (this.chartSpace.chart.plotArea.chart.gapWidth / 100) : (150 / 100);
		let gapDepth = this.chartSpace.chart.plotArea.chart.gapDepth != null ? (this.chartSpace.chart.plotArea.chart.gapDepth / 100) : (150 / 100);

		if (AscFormat.c_oChartTypes.Area === this.chartsDrawer.calcProp.type || AscFormat.c_oChartTypes.Surface === this.chartsDrawer.calcProp.type) {
			gapWidth = 0;
			gapDepth = 0;
		}

		let baseDepth = width / (seriesCount - (seriesCount - 1) * overlap + gapWidth);
		if (this.chartsDrawer.calcProp.subType === "standard" || this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Line || isNormalArea) {
			baseDepth = (width / (seriesCount - (seriesCount - 1) * overlap + gapWidth)) * seriesCount;
		}

		let basePercent = this._getDepthPercent();
		let depth = baseDepth * basePercent;
		depth = depth + depth * gapDepth;

		if (this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.HBar && this.hPercent !== null) {
			depth = this.hPercent * depth;
		}

		//TODO глубина в некоторых случаях отличается(тип Standard)
		if (this.chartsDrawer.calcProp.subType === "standard" || this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Line || isNormalArea) {
			let b = 1 / seriesCount;

			let sinOx = Math.sin(-this.angleOx);
			let angleOxKf = 1;
			let a = basePercent / (ptCount * b);
			depth = (widthChart * a + widthChart * a) / (1 / angleOxKf + a);

			depth = depth / angleOxKf;
		}

		return depth;
	};

	Processor3D.prototype._calcSpecialStandardScaleX = function () {
		let type = this.chartsDrawer.calcProp.type;
		let subType = this.chartsDrawer.calcProp.subType;
		if (!(subType === "standard" || type === AscFormat.c_oChartTypes.Line || (type === AscFormat.c_oChartTypes.Area && subType === "normal") || type ===
			AscFormat.c_oChartTypes.Surface)) {
			return;
		}

		let countSeries = this.chartsDrawer.calculateFirstChartCountSeries();
		let seriesCount = countSeries.series;
		let ptCount = countSeries.points;

		//calculate width in 3d standard charts with rAngAx
		let n = Math.floor((seriesCount + ptCount) / 2);
		let kf = ptCount / n;

		this.specialStandardScaleX = 1 / kf;
	};

	Processor3D.prototype._calculateScaleFromDepth = function (/*isSkip*/) {
		//***Calculate scaleY***
		if (this.view3D.getRAngAx() && this.aspectRatioY === 1) {
			let heightCanvas = this.heightCanvas;
			let heightChart = heightCanvas - this.top - this.bottom;

			this.scaleY = heightChart / (-this.depthPerspective * Math.sin(Math.abs(this.angleOx)) + heightChart);

			let subType = this.chartsDrawer.calcProp.subType;
			let newDepth, newWidth;
			if (!(subType === "standard" || this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Line || this.chartsDrawer.calcProp.type ===
				AscFormat.c_oChartTypes.Surface || (this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Area && subType === "normal"))) {
				newDepth = this.depthPerspective * Math.sin(-this.angleOx);
				newWidth = heightChart - newDepth;
				this.scaleX = heightChart / newWidth;
			}
		}
	};

	Processor3D.prototype._calculateCameraDiff = function (/*isSkip*/) {
		//глубина по OZ
		let perspectiveDepth = this.depthPerspective;

		let widthCanvas = this.widthCanvas;
		let originalWidthChart = widthCanvas - this.left - this.right;

		let heightCanvas = this.heightCanvas;
		let heightChart = heightCanvas - this.top - this.bottom;

		//add test points for parallelepiped rect
		let points = [];
		let faces = [];

		/*if(AscFormat.c_oChartTypes.Pie === this.chartsDrawer.calcProp.type)
		{
			points.push(new Point3D(this.left + originalWidthChart / 2, this.top, perspectiveDepth, this));
			points.push(new Point3D(this.left, this.top, perspectiveDepth / 2, this));
			points.push(new Point3D(this.left + originalWidthChart, this.top, perspectiveDepth / 2, this));
			points.push(new Point3D(this.left + originalWidthChart / 2, this.top, 0, this));

			points.push(new Point3D(this.left + originalWidthChart / 2, this.top + heightChart, perspectiveDepth, this));
			points.push(new Point3D(this.left, this.top + heightChart, perspectiveDepth / 2, this));
			points.push(new Point3D(this.left + originalWidthChart, this.top + heightChart, perspectiveDepth / 2, this));
			points.push(new Point3D(this.left + originalWidthChart / 2, this.top + heightChart, 0, this));
		}
		else
		{*/
		points.push(new Point3D(this.left, this.top, perspectiveDepth, this));
		points.push(new Point3D(this.left, heightChart + this.top, perspectiveDepth, this));
		points.push(new Point3D(originalWidthChart + this.left, heightChart + this.top, perspectiveDepth, this));
		points.push(new Point3D(originalWidthChart + this.left, this.top, perspectiveDepth, this));
		points.push(new Point3D(originalWidthChart + this.left, this.top, 0, this));
		points.push(new Point3D(originalWidthChart + this.left, heightChart + this.top, 0, this));
		points.push(new Point3D(this.left, heightChart + this.top, 0, this));
		points.push(new Point3D(this.left, this.top, 0, this));
		//}

		faces.push([0, 1, 2, 3]);
		faces.push([2, 5, 4, 3]);
		faces.push([1, 6, 7, 0]);
		faces.push([6, 5, 4, 7]);
		faces.push([7, 4, 3, 0]);
		faces.push([1, 6, 2, 5]);


		//***Calculate cameraDiffZ***
		if (!this.view3D.getRAngAx()) {
			//быстрая функция поиска сдвигов камеры
			//console.time("sdf");
			this._calculateCameraDiffZX(points, faces);
			//console.timeEnd("sdf");
		}

		//***Calculate cameraDiffX***
		if (this.view3D.getRAngAx()) {
			let minMaxOx = this._getMinMaxOx(points, faces);
			this._calculateCameraDiffX(minMaxOx);

			//***Calculate cameraDiffY***
			let minMaxOy = this._getMinMaxOy(points, faces);
			this._calculateCameraDiffY(minMaxOy.top, minMaxOy.bottom);
		}
	};

	Processor3D.prototype._calculateCameraDiffZX = function (newPoints) {
		let heightChart = this.heightCanvas - this.top - this.bottom;
		let widthOriginalChart = this.widthCanvas - this.left - this.right;
		let heightOriginalChart = heightChart;

		let minX = null;
		let maxX = null;
		let minZ = null;

		let aspectRatio = (widthOriginalChart) / (heightOriginalChart);


		for (let i = 0; i < newPoints.length; i++) {
			let point3D = new Point3D(newPoints[i].x, newPoints[i].y, newPoints[i].z, this);

			point3D.scale(aspectRatio, 1, 1);
			point3D.offset((-this.widthCanvas / 2) / aspectRatio, (-this.heightCanvas / 2), 0);

			//rotate
			let matrixRotateAllAxis = this._getMatrixRotateAllAxis();
			point3D.multiplyPointOnMatrix1(matrixRotateAllAxis);


			if (minZ === null || point3D.z < minZ) {
				minZ = point3D.z;
			}

			if (minX === null || point3D.x < minX) {
				minX = point3D.x;
			}

			if (maxX === null || point3D.x > maxX) {
				maxX = point3D.x;
			}
		}

		//get min and max point's and diffX
		/*if(-t.angleOx === 0)
		{
			let minMaxOx = this._getMinMaxOxPoints(newPoints, minZ);
			this.cameraDiffZ = - minZ;
			this.cameraDiffX = - minMaxOx.diffX;

			let x111 = this.convertAndTurnPoint(minMaxOx.tempX1, minMaxOx.tempY1, minMaxOx.tempZ1);
			let x222 = this.convertAndTurnPoint(minMaxOx.tempX2, minMaxOx.tempY2, minMaxOx.tempZ2);

			if(Math.abs(x222.x - x111.x) > widthOriginalChart)
			{
				let correctOffset;
				if(x222.x > x111.x)
				{
					correctOffset = this._correctZPosition4(minMaxOx.tempX1, minMaxOx.tempX2, minMaxOx.tempZ1, minMaxOx.tempZ2, minZ, minMaxOx.tempY1, minMaxOx.tempY2);
				}
				else
				{
					correctOffset = this._correctZPosition4(minMaxOx.tempX1, minMaxOx.tempX2, minMaxOx.tempZ1, minMaxOx.tempZ2, minZ, minMaxOx.tempY1, minMaxOx.tempY2);
				}

				this.cameraDiffZ = Math.abs(correctOffset.minZ);
				this.cameraDiffX = correctOffset.diffX;
			}
		}
		else if(-t.angleOy === 0 && -t.angleOx !== 0)
		{
			let minMaxOy = this._getMinMaxOyPoints(newPoints, minZ);
			this.cameraDiffZ = - minZ;
			this.cameraDiffY = -minMaxOy.diffY;

			let y111 = this.convertAndTurnPoint(minMaxOy.tempX1, minMaxOy.tempY1, minMaxOy.tempZ1);
			let y222 = this.convertAndTurnPoint(minMaxOy.tempX2, minMaxOy.tempY2, minMaxOy.tempZ2);

			if(Math.abs(y111.y - y222.y) < heightOriginalChart)
			{
				let minMaxOx = this._getMinMaxOxPoints(newPoints, minZ);
				let x111 = this.convertAndTurnPoint(minMaxOx.tempX1, minMaxOx.tempY1, minMaxOx.tempZ1);
				let x222 = this.convertAndTurnPoint(minMaxOx.tempX2, minMaxOx.tempY2, minMaxOx.tempZ2);

				let correctOffset;
				if(x222.x > x111.x)
				{
					correctOffset = this._correctZPosition4(minMaxOx.tempX1, minMaxOx.tempX2, minMaxOx.tempZ1, minMaxOx.tempZ2, minZ, minMaxOx.tempY1, minMaxOx.tempY2);
				}
				else
				{
					correctOffset = this._correctZPosition4(minMaxOx.tempX1, minMaxOx.tempX2, minMaxOx.tempZ1, minMaxOx.tempZ2, minZ, minMaxOx.tempY1, minMaxOx.tempY2);
				}

				this.cameraDiffZ = Math.abs(correctOffset.minZ);
				this.cameraDiffX = correctOffset.diffX;


				let minMaxOy = this._getMinMaxOyPoints(newPoints, 0);
				//this.cameraDiffZ = - minZ;
				this.cameraDiffY = -minMaxOy.diffY;
			}
		}*/
		/*else if(-t.angleOy !== 0 && -t.angleOx !== 0)
		{
			let minMaxOx = this._getMinMaxOxPoints(newPoints, minZ);
			let x111 = this.convertAndTurnPoint(minMaxOx.tempX1, minMaxOx.tempY1, minMaxOx.tempZ1);
			let x222 = this.convertAndTurnPoint(minMaxOx.tempX2, minMaxOx.tempY2, minMaxOx.tempZ2);

			let correctOffset;
			if(x222.x > x111.x)
			{
				correctOffset = this._correctZPosition4(minMaxOx.tempX1, minMaxOx.tempX2, minMaxOx.tempZ1, minMaxOx.tempZ2, minZ, minMaxOx.tempY1, minMaxOx.tempY2);
			}
			else
			{
				correctOffset = this._correctZPosition4(minMaxOx.tempX2, minMaxOx.tempX1, minMaxOx.tempZ2, minMaxOx.tempZ1, minZ, minMaxOx.tempY2, minMaxOx.tempY1);
			}

			this.cameraDiffZ = correctOffset.minZ;
			this.cameraDiffX = correctOffset.diffX;

			let minMaxOy = this._getMinMaxOyPoints2(newPoints, this.cameraDiffZ);
			this.cameraDiffY = minMaxOy.diffY;

			console.log(" tempX1: " + minMaxOy.tempX1 + " tempX2: " + minMaxOy.tempX2 + " tempY1: " + minMaxOy.tempY1 + " tempY2: " + minMaxOy.tempY2 + " tempZ1: " + minMaxOy.tempZ1 + " tempZ2: " + minMaxOy.tempZ2);

			let y111 = this.convertAndTurnPoint(minMaxOy.tempX1, minMaxOy.tempY1, minMaxOy.tempZ1);
			let y222 = this.convertAndTurnPoint(minMaxOy.tempX2, minMaxOy.tempY2, minMaxOy.tempZ2);

			if(y111.y < this.top || y222.y > this.top + heightOriginalChart)
			{
				//correctOffset = this._correctZPositionOY(minMaxOy.tempX1, minMaxOy.tempX2, minMaxOy.tempZ1, minMaxOy.tempZ2, minZ, minMaxOy.tempY1, minMaxOy.tempY2);
				this.cameraDiffZ = minMaxOy.diffZ;
				//this.cameraDiffY = minMaxOy.diffY;

				let minMaxOx = this._getMinMaxOxPoints(newPoints, -this.cameraDiffZ);
				//correctOffset = this._correctZPosition4(minMaxOx.tempX1, minMaxOx.tempX2, minMaxOx.tempZ1, minMaxOx.tempZ2, -this.cameraDiffZ, minMaxOx.tempY1, minMaxOx.tempY2);
				this.cameraDiffX = -minMaxOx.diffX;
			}



			//TODO пока включаю для поворотов  по OX + по OY checkOutSideArea(медленная функция), затем нужно переделать, используя закомментированный код сверху
			this.cameraDiffZ = -minZ;
			//this.cameraDiffX = -minMaxOx.diffX;
			this.checkOutSideArea(newPoints);
		}*/

		//TODO пока включаю для ВСЕГО checkOutSideArea(медленная функция), затем нужно переделать, используя закомментированный код сверху
		this.cameraDiffZ = -minZ;
		this.checkOutSideArea2(newPoints);

	};

//TODO если будут проблемы при маштабировании, вернуть функцию checkOutSideArea2 вместо checkOutSideArea
	Processor3D.prototype.checkOutSideArea2 = function (newPoints) {
		let i = 0;
		let maxI = 1000;
		let t = this;

		let heightChart = this.heightCanvas - this.top - this.bottom;
		let widthChart = this.widthCanvas - this.left - this.right;

		let calculateZ = function (step) {
			let minMaxOx = t._getMinMaxOxPoints(newPoints, -t.cameraDiffZ);
			t.cameraDiffX = -minMaxOx.diffX;
			let x111 = t.convertAndTurnPoint(minMaxOx.tempX1, minMaxOx.tempY1, minMaxOx.tempZ1);
			let x222 = t.convertAndTurnPoint(minMaxOx.tempX2, minMaxOx.tempY2, minMaxOx.tempZ2);

			let diffX = Math.abs(x222.x - x111.x);

			let minMaxOy = t._getMinMaxOyPoints(newPoints, t.cameraDiffZ);
			t.cameraDiffY = -minMaxOy.diffY;
			let y111 = t.convertAndTurnPoint(minMaxOy.tempX1, minMaxOy.tempY1, minMaxOy.tempZ1);
			let y222 = t.convertAndTurnPoint(minMaxOy.tempX2, minMaxOy.tempY2, minMaxOy.tempZ2);

			let diffY = Math.abs(y222.y - y111.y);

			if (diffX < widthChart && diffY < heightChart)//if size less then width or height
			{
				while (diffX < widthChart && diffY < heightChart) {
					t.cameraDiffZ -= step;
					i++;
					minMaxOx = t._getMinMaxOxPoints(newPoints, -t.cameraDiffZ);
					t.cameraDiffX = -minMaxOx.diffX;

					x111 = t.convertAndTurnPoint(minMaxOx.tempX1, minMaxOx.tempY1, minMaxOx.tempZ1);
					x222 = t.convertAndTurnPoint(minMaxOx.tempX2, minMaxOx.tempY2, minMaxOx.tempZ2);
					diffX = Math.abs(x222.x - x111.x);

					minMaxOy = t._getMinMaxOyPoints(newPoints, t.cameraDiffZ);
					t.cameraDiffY = -minMaxOy.diffY;
					y111 = t.convertAndTurnPoint(minMaxOy.tempX1, minMaxOy.tempY1, minMaxOy.tempZ1);
					y222 = t.convertAndTurnPoint(minMaxOy.tempX2, minMaxOy.tempY2, minMaxOy.tempZ2);
					diffY = Math.abs(y222.y - y111.y);

					if (i > maxI) {
						break;
					}
				}
			} else if (diffX > widthChart || diffY > heightChart)//if size more then width or height
			{
				while (diffX > widthChart || diffY > heightChart) {
					t.cameraDiffZ += step;
					i++;
					minMaxOx = t._getMinMaxOxPoints(newPoints, -t.cameraDiffZ);
					t.cameraDiffX = -minMaxOx.diffX;

					x111 = t.convertAndTurnPoint(minMaxOx.tempX1, minMaxOx.tempY1, minMaxOx.tempZ1);
					x222 = t.convertAndTurnPoint(minMaxOx.tempX2, minMaxOx.tempY2, minMaxOx.tempZ2);
					diffX = Math.abs(x222.x - x111.x);

					minMaxOy = t._getMinMaxOyPoints(newPoints, t.cameraDiffZ);
					t.cameraDiffY = -minMaxOy.diffY;
					y111 = t.convertAndTurnPoint(minMaxOy.tempX1, minMaxOy.tempY1, minMaxOy.tempZ1);
					y222 = t.convertAndTurnPoint(minMaxOy.tempX2, minMaxOy.tempY2, minMaxOy.tempZ2);
					diffY = Math.abs(y222.y - y111.y);

					if (i > maxI) {
						break;
					}
				}
			}
		};

		calculateZ(100);
		calculateZ(10);
		calculateZ(1);
	};

	Processor3D.prototype.checkOutSideArea = function (newPoints) {
		let t = this;

		let heightChart = this.heightCanvas - this.top - this.bottom;
		let widthChart = this.widthCanvas - this.left - this.right;
		let DELTA = 3;
		let maxCount = 1000;
		let calculateZ = function () {
			let minMaxOx = t._getMinMaxOxPoints(newPoints, -t.cameraDiffZ);
			t.cameraDiffX = -minMaxOx.diffX;
			let x111 = t.convertAndTurnPoint(minMaxOx.tempX1, minMaxOx.tempY1, minMaxOx.tempZ1);
			let x222 = t.convertAndTurnPoint(minMaxOx.tempX2, minMaxOx.tempY2, minMaxOx.tempZ2);

			let diffX = Math.abs(x222.x - x111.x);

			let minMaxOy = t._getMinMaxOyPoints(newPoints, t.cameraDiffZ);
			t.cameraDiffY = -minMaxOy.diffY;
			let y111 = t.convertAndTurnPoint(minMaxOy.tempX1, minMaxOy.tempY1, minMaxOy.tempZ1);
			let y222 = t.convertAndTurnPoint(minMaxOy.tempX2, minMaxOy.tempY2, minMaxOy.tempZ2);

			let diffY = Math.abs(y222.y - y111.y);

			if (diffX <= widthChart && diffY <= heightChart && ((widthChart - diffX) < DELTA || (heightChart - diffY) < DELTA)) {
				return;
			}

			let count = 0;
			let fStart = t.cameraDiffZ ? t.cameraDiffZ : 1;
			let fLeftDiffZ, fRightDiffZ;
			if (diffX < widthChart && diffY < heightChart) {
				fLeftDiffZ = 0;
				fRightDiffZ = t.cameraDiffZ;
			} else {
				fLeftDiffZ = t.cameraDiffZ;
				while (diffX >= widthChart || diffY >= heightChart) {
					count++;
					t.cameraDiffZ += fStart / 2;
					minMaxOx = t._getMinMaxOxPoints(newPoints, -t.cameraDiffZ);
					t.cameraDiffX = -minMaxOx.diffX;
					x111 = t.convertAndTurnPoint(minMaxOx.tempX1, minMaxOx.tempY1, minMaxOx.tempZ1);
					x222 = t.convertAndTurnPoint(minMaxOx.tempX2, minMaxOx.tempY2, minMaxOx.tempZ2);
					diffX = Math.abs(x222.x - x111.x);
					minMaxOy = t._getMinMaxOyPoints(newPoints, t.cameraDiffZ);
					t.cameraDiffY = -minMaxOy.diffY;
					y111 = t.convertAndTurnPoint(minMaxOy.tempX1, minMaxOy.tempY1, minMaxOy.tempZ1);
					y222 = t.convertAndTurnPoint(minMaxOy.tempX2, minMaxOy.tempY2, minMaxOy.tempZ2);
					diffY = Math.abs(y222.y - y111.y);

					if (count > maxCount) {
						return;
					}
				}
				if (diffX <= widthChart && diffY <= heightChart && ((widthChart - diffX) < DELTA || (heightChart - diffY) < DELTA)) {
					return;
				}
				fRightDiffZ = t.cameraDiffZ;
			}


			while ((diffX > widthChart || diffY > heightChart) || (fRightDiffZ - fLeftDiffZ) > DELTA) {
				t.cameraDiffZ = fLeftDiffZ + ((fRightDiffZ - fLeftDiffZ) / 2);

				count++;
				minMaxOx = t._getMinMaxOxPoints(newPoints, -t.cameraDiffZ);
				t.cameraDiffX = -minMaxOx.diffX;
				let x111 = t.convertAndTurnPoint(minMaxOx.tempX1, minMaxOx.tempY1, minMaxOx.tempZ1);
				let x222 = t.convertAndTurnPoint(minMaxOx.tempX2, minMaxOx.tempY2, minMaxOx.tempZ2);


				let diffX = Math.abs(x222.x - x111.x);

				let minMaxOy = t._getMinMaxOyPoints(newPoints, t.cameraDiffZ);
				t.cameraDiffY = -minMaxOy.diffY;
				let y111 = t.convertAndTurnPoint(minMaxOy.tempX1, minMaxOy.tempY1, minMaxOy.tempZ1);
				let y222 = t.convertAndTurnPoint(minMaxOy.tempX2, minMaxOy.tempY2, minMaxOy.tempZ2);

				let diffY = Math.abs(y222.y - y111.y);
				if (diffX < widthChart && diffY < heightChart) {
					if (((widthChart - diffX) < DELTA) || ((heightChart - diffY) < DELTA)) {
						break;
					}
					fRightDiffZ = t.cameraDiffZ;
				} else {
					fLeftDiffZ = t.cameraDiffZ;
				}

				if (count > maxCount) {
					return;
				}
			}
		};

		calculateZ();

	};

	Processor3D.prototype._getMinMaxOxPoints = function (points, minZ) {
		let fov = 1 / this.rPerspective;

		let t = this;
		let aspectRatioX = this.aspectRatioX;
		let aspectRatioY = this.aspectRatioY;


		let diffAndRotatePoint = function (point) {
			let point3D = new Point3D(point.x, point.y, point.z, t);

			point3D.scale(aspectRatioX, aspectRatioY, 1);
			point3D.offset((-t.widthCanvas / 2) / aspectRatioX, (-t.heightCanvas / 2) / aspectRatioY /** aspectRatio*/, 0);

			//rotate
			let matrixRotateAllAxis = t._getMatrixRotateAllAxis();
			point3D.multiplyPointOnMatrix1(matrixRotateAllAxis);

			return point3D;
		};

		let calculateDiffX = function (x1, x2, z1, z2, minZ, y1, y2) {
			let diffAndScalePoints1 = t.diffAndScale(x1, y1, z1);
			x1 = diffAndScalePoints1.x;
			y1 = diffAndScalePoints1.y;
			z1 = diffAndScalePoints1.z;

			let rotatePoints1 = t.rotate(x1, y1, z1);

			let a1 = rotatePoints1.x;
			//let x1S = a1 - diffX;
			let z1S = rotatePoints1.z - minZ;
			//let x1SS = (fov * x1S / (z1S + fov)) + w / 2;


			let diffAndScalePoints2 = t.diffAndScale(x2, y2, z2);
			x2 = diffAndScalePoints2.x;
			y2 = diffAndScalePoints2.y;
			z2 = diffAndScalePoints2.z;

			let rotatePoints2 = t.rotate(x2, y2, z2);

			let a2 = rotatePoints2.x;
			//let x2S = a2 - diffX;
			let z2S = rotatePoints2.z - minZ;
			//let x2SS = (fov * x2S / (z2S + fov)) + w / 2;

			return (-a1 * z2S - a1 * fov - a2 * z1S - a2 * fov) / (-z2S - fov - z1S - fov);
		};


		let w = t.widthCanvas;

		let tempArray = this._getArrayAllVergeCube(points);

		let diffX;
		let x11, x22, y11, y22, z11, z22;
		let tempX1, tempY1, tempZ1, tempX2, tempY2, tempZ2, start, end;
		for (let i = 0; i < tempArray.length - 1; i++) {
			start = i;
			end = i + 1;
			x11 = tempArray[start].x - (t.widthCanvas / 2);
			x22 = tempArray[end].x - (t.widthCanvas / 2);

			y11 = tempArray[start].y - (t.heightCanvas / 2);
			y22 = tempArray[end].y - (t.heightCanvas / 2);

			z11 = tempArray[start].z;
			z22 = tempArray[end].z;

			tempX1 = tempArray[start].x;
			tempY1 = tempArray[start].y;
			tempZ1 = tempArray[start].z;

			tempX2 = tempArray[end].x;
			tempY2 = tempArray[end].y;
			tempZ2 = tempArray[end].z;

			if (x11 > x22) {
				start = i + 1;
				end = i;
				x11 = tempArray[start].x - (t.widthCanvas / 2);
				x22 = tempArray[end].x - (t.widthCanvas / 2);

				y11 = tempArray[start].y - (t.heightCanvas / 2);
				y22 = tempArray[end].y - (t.heightCanvas / 2);

				z11 = tempArray[start].z;
				z22 = tempArray[end].z;

				tempX1 = tempArray[start].x;
				tempY1 = tempArray[start].y;
				tempZ1 = tempArray[start].z;

				tempX2 = tempArray[end].x;
				tempY2 = tempArray[end].y;
				tempZ2 = tempArray[end].z;
			}


			diffX = calculateDiffX(tempX1, tempX2, tempZ1, tempZ2, minZ, tempY1, tempY2);

			let projectiveMatrix = t._getPerspectiveProjectionMatrix(1 / (t.rPerspective));

			let rotatePoint1 = diffAndRotatePoint(tempArray[start]);
			rotatePoint1.offset(-diffX, 0, -minZ);
			rotatePoint1 = rotatePoint1.project(projectiveMatrix);
			let x1 = Math.floor(rotatePoint1.x + t.widthCanvas / 2);


			let rotatePoint2 = diffAndRotatePoint(tempArray[end]);
			rotatePoint2.offset(-diffX, 0, -minZ);
			rotatePoint2 = rotatePoint2.project(projectiveMatrix);
			let x2 = Math.floor(rotatePoint2.x + t.widthCanvas / 2);

			let leftMargin = x1;
			let rightMargin = Math.floor(w - x2);
			if (!((leftMargin >= rightMargin - 1) && (leftMargin <= rightMargin + 1))) {
				continue;
			}

			let isTrue = true;
			for (let l = 0; l < tempArray.length - 1; l++) {
				let rotatePoint = diffAndRotatePoint(tempArray[l]);
				rotatePoint.offset(-diffX, 0, -minZ);
				rotatePoint = rotatePoint.project(projectiveMatrix);
				let tempX11 = Math.floor(rotatePoint.x + t.widthCanvas / 2);


				if (x1 < x2) {
					if (tempX11 < x1 || tempX11 > x2) {
						isTrue = false;
						break;
					}
				} else {
					if (tempX11 > x1 || tempX11 < x2) {
						isTrue = false;
						break;
					}
				}
			}

			if (isTrue) {
				break;
			}
		}

		return {
			diffX: diffX, tempX1: tempX1, tempY1: tempY1, tempZ1: tempZ1, tempX2: tempX2, tempY2: tempY2, tempZ2: tempZ2, x11: x11, z11: z11, x22: x22, z22: z22, y11: y11, y22: y22
		};
	};

	Processor3D.prototype._getMinMaxOyPoints = function (points, minZ) {
		let fov = 1 / this.rPerspective;
		let t = this;
		let h = t.heightCanvas;

		let calculateDiffY = function (x1, x2, y1, y2, z1, z2, minZ) {
			//let cos1 = Math.cos(-t.angleOy);
			//let sin1 = Math.sin(-t.angleOy);
			let diffY = 0;

			let diffAndScalePoints = t.diffAndScale(x1, y1, z1);
			x1 = diffAndScalePoints.x;
			y1 = diffAndScalePoints.y;
			z1 = diffAndScalePoints.z;


			let rotatePoints = t.rotate(x1, y1, z1);

			let a1 = rotatePoints.y;
			//let y1S = (a1 - diffY);
			let z1S = rotatePoints.z + minZ;

			//let x1SS = (fov * y1S / (z1S + fov)) + h / 2;


			diffAndScalePoints = t.diffAndScale(x2, y2, z2);
			x2 = diffAndScalePoints.x;
			y2 = diffAndScalePoints.y;
			z2 = diffAndScalePoints.z;

			rotatePoints = t.rotate(x2, y2, z2);

			let a2 = rotatePoints.y;
			//let y2S = (a2 - diffY);
			let z2S = rotatePoints.z + minZ;

			//let x2SS = (fov * y2S / (z2S + fov)) + h / 2;

			//let topMargin = (fov * y1S / (z1S + fov)) + h / 2;
			//let bottomMargin = h - ((fov * y2S / (z2S + fov)) + h / 2);


			//((a1 - diffY) / (z1S + fov)) =  - (((a2 - diffY) / (z2S + fov)))

			/*(a1 - diffY) * (z2S + fov) = -(z1S + fov) * (a2 - diffY)

			a1 * z2S + a1 * fov - diffY * z2S - diffY * fov = -z1S * a2 + z1S * diffY - fov * a2 + fov * diffY

			a1 * z2S + a1 * fov + fov * a2 + z1S * a2= diffY * z2S + diffY * fov + z1S * diffY + fov * diffY*/

			diffY = (a1 * z2S + a1 * fov + fov * a2 + z1S * a2) / (z2S + fov + z1S + fov);

			//let diffY = (-a1 * z2S - a1 * fov - a2 * z1S -a2 * fov) / ( -z2S - fov - z1S - fov);


			return diffY;

		};

		let tempArray = this._getArrayAllVergeCube(points);

		let diffY;
		let tempX1, tempY1, tempZ1, tempX2, tempY2, tempZ2;
		for (let i = 0; i < tempArray.length - 1; i++) {
			let start = i;
			let end = i + 1;

			tempX1 = tempArray[start].x;
			tempY1 = tempArray[start].y;
			tempZ1 = tempArray[start].z;

			tempX2 = tempArray[end].x;
			tempY2 = tempArray[end].y;
			tempZ2 = tempArray[end].z;

			if (tempY1 > tempY2) {
				start = i + 1;
				end = i;

				tempX1 = tempArray[start].x;
				tempY1 = tempArray[start].y;
				tempZ1 = tempArray[start].z;

				tempX2 = tempArray[end].x;
				tempY2 = tempArray[end].y;
				tempZ2 = tempArray[end].z;
			}


			diffY = calculateDiffY(tempX1, tempX2, tempY1, tempY2, tempZ1, tempZ2, minZ);

			let rotatePoint1 = this.calculatePointManual(tempX1, tempY1, tempZ1, this.cameraDiffX, -diffY, minZ);
			let y1 = rotatePoint1.y;


			let rotatePoint2 = this.calculatePointManual(tempX2, tempY2, tempZ2, this.cameraDiffX, -diffY, minZ);
			let y2 = rotatePoint2.y;

			let topMargin = y1;
			let bottomMargin = Math.floor(h - y2);
			if (!((topMargin >= bottomMargin - 1) && (topMargin <= bottomMargin + 1))) {
				continue;
			}

			let isTrue = true;
			for (let l = 0; l < tempArray.length - 1; l++) {
				let rotatePoint = this.calculatePointManual(tempArray[l].x, tempArray[l].y, tempArray[l].z, this.cameraDiffX, -diffY, minZ);
				let tempY11 = rotatePoint.y;

				if (y1 < y2) {
					if (tempY11 < y1 || tempY11 > y2) {
						isTrue = false;
						break;
					}
				} else {
					if (tempY11 > y1 || tempY11 < y2) {
						isTrue = false;
						break;
					}
				}
			}

			if (isTrue) {
				break;
			}
		}

		return {
			diffY: diffY, tempX1: tempX1, tempY1: tempY1, tempZ1: tempZ1, tempX2: tempX2, tempY2: tempY2, tempZ2: tempZ2
		};
	};

	Processor3D.prototype._getArrayAllVergeCube = function (points) {
		let res = [];

		res[0] = points[0];
		res[1] = points[1];
		res[2] = points[2];
		res[3] = points[3];
		res[4] = points[0];
		res[5] = points[4];
		res[6] = points[7];
		res[7] = points[6];
		res[8] = points[5];
		res[9] = points[2];
		res[10] = points[6];
		res[11] = points[1];
		res[12] = points[7];
		res[13] = points[0];
		res[14] = points[5];
		res[15] = points[1];
		res[16] = points[4];
		res[17] = points[3];
		res[18] = points[6];
		res[19] = points[0];
		res[20] = points[2];
		res[21] = points[7];
		res[22] = points[3];
		res[23] = points[5];
		res[24] = points[7];
		res[25] = points[4];
		res[26] = points[2];
		res[27] = points[3];
		res[28] = points[1];
		res[29] = points[4];
		res[30] = points[5];
		res[31] = points[6];
		res[32] = points[4];


		/*for(let i = 0; i < points.length; i++)
		{
			for(let j = 0; j < points.length; j++)
			{
				res.push(points[i]);
				res.push(points[j]);

			}
		}*/

		return res;
	};

	Processor3D.prototype._getMinMaxOyPoints2 = function (points, minZ) {
		let t = this;

		let h = t.heightCanvas;

		let tempArray = [];
		tempArray[0] = points[0];//leftNear
		tempArray[1] = points[1];//leftFar
		tempArray[2] = points[2];//rightFar
		tempArray[3] = points[3];//rightNear
		tempArray[4] = points[4];//leftNear
		tempArray[5] = points[5];//leftFar
		tempArray[6] = points[6];//leftFar
		tempArray[7] = points[7];//leftFar
		tempArray[8] = points[4];//leftFar
		tempArray[9] = points[1];//leftFar
		tempArray[10] = points[3];//leftFar
		tempArray[11] = points[0];//leftFar
		tempArray[12] = points[7];//leftFar
		tempArray[13] = points[2];//leftFar
		tempArray[14] = points[0];//leftFar
		tempArray[15] = points[6];//leftFar
		tempArray[16] = points[5];//leftFar
		tempArray[17] = points[0];//leftFar
		tempArray[18] = points[4];//leftFar


		let tempX1, tempY1, tempZ1, tempX2, tempY2, tempZ2;
		for (let i = 0; i < tempArray.length - 1; i++) {
			let start = i;
			let end = i + 1;

			tempX1 = tempArray[start].x;
			tempY1 = tempArray[start].y;
			tempZ1 = tempArray[start].z;

			tempX2 = tempArray[end].x;
			tempY2 = tempArray[end].y;
			tempZ2 = tempArray[end].z;

			if (tempY1 > tempY2) {
				start = i + 1;
				end = i;

				tempX1 = tempArray[start].x;
				tempY1 = tempArray[start].y;
				tempZ1 = tempArray[start].z;

				tempX2 = tempArray[end].x;
				tempY2 = tempArray[end].y;
				tempZ2 = tempArray[end].z;
			}


			//let diffY = calculateDiffY(tempX1, tempX2, tempY1, tempY2, tempZ1, tempZ2, minZ);


			let correctOffset = this._correctZPositionOY(tempX1, tempX2, tempZ1, tempZ2, minZ, tempY1, tempY2);
			let diffZ = correctOffset.minZ;
			let diffY = correctOffset.diffY;


			let rotatePoint1 = this.calculatePointManual(tempX1, tempY1, tempZ1, this.cameraDiffX, diffY, diffZ);
			let y1 = rotatePoint1.y;


			let rotatePoint2 = this.calculatePointManual(tempX2, tempY2, tempZ2, this.cameraDiffX, diffY, diffZ);
			let y2 = rotatePoint2.y;

			let topMargin = y1;
			let bottomMargin = Math.floor(h - y2);
			if (!((topMargin >= bottomMargin - 1) && (topMargin <= bottomMargin + 1))) {
				continue;
			}

			let isTrue = true;
			for (let l = 0; l < tempArray.length - 1; l++) {
				let rotatePoint = this.calculatePointManual(tempArray[l].x, tempArray[l].y, tempArray[l].z, this.cameraDiffX, diffY, diffZ);
				let tempY11 = rotatePoint.y;

				if (y1 < y2) {
					if (tempY11 < y1 || tempY11 > y2) {
						isTrue = false;
						break;
					}
				} else {
					if (tempY11 > y1 || tempY11 < y2) {
						isTrue = false;
						break;
					}
				}
			}

			if (isTrue) {
				break;
			}
		}

		return {
			tempX1: tempX1, tempY1: tempY1, tempZ1: tempZ1, tempX2: tempX2, tempY2: tempY2, tempZ2: tempZ2
		};
	};

	Processor3D.prototype._correctZPosition4 = function (x1, x2, z1, z2, minZ, y1, y2) {
		let t = this;

		let getDiffXZ3 = function (x1, x2, z1, z2, y1, y2) {
			let w = t.widthCanvas;

			let fov = 1 / t.rPerspective;

			let diffAndScalePoints = t.diffAndScale(x1, y1, z1);
			let x11 = diffAndScalePoints.x;
			let y11 = diffAndScalePoints.y;
			let z11 = diffAndScalePoints.z;


			let rotatePoints = t.rotate(x11, y11, z11);
			let x111 = rotatePoints.x;
			let z111 = rotatePoints.z;


			/*let x1111 = x111 + diffX;
			let y1111 = y111 + diffY;
			let z1111 = z111 + diffZ;
			let x11111 = (fov * x1111) / (z1111 + fov) + w / 2;
			let y11111 = (fov * y1111) / (z1111 + fov) + h / 2;*/


			//(fov * x1111) / (z1111 + fov) + w / 2 = t.left
			//(fov * (x111 + diffX)) / ((z111 + diffZ) + fov) + w / 2 = t.left
			let wL = t.left - w / 2;
			//fov * (x111 + diffX) = wL * ((z111 + diffZ) + fov)
			//fov * x111 + fov * diffX = wL * z111 + wL * diffZ + wL * fov


			diffAndScalePoints = t.diffAndScale(x2, y2, z2);
			let x22 = diffAndScalePoints.x;
			let y22 = diffAndScalePoints.y;
			let z22 = diffAndScalePoints.z;


			rotatePoints = t.rotate(x22, y22, z22);
			let x222 = rotatePoints.x;
			let z222 = rotatePoints.z;


			/*let x1111 = x111 + diffX;
			let y1111 = y111 + diffY;
			let z1111 = z111 + diffZ;
			let x11111 = (fov * x1111) / (z1111 + fov) + w / 2;
			let y11111 = (fov * y1111) / (z1111 + fov) + h / 2;*/


			let wR = w / 2 - t.left;
			//fov * (x222 + diffX) = wR * ((z222 + diffZ) + fov)
			//fov * x222 + fov * diffX = wR * z222 + wR * diffZ + wR * fov


			//итого
			//fov * x111 + fov * diffX = wL * z111 + wL * diffZ + wL * fov
			//fov * x222 + fov * diffX = wR * z222 + wR * diffZ + wR * fov


			/*diffX = (wL * z111 + wL * diffZ + wL * fov - fov * x111) / fov;

			fov * x222 + wL * z111 + wL * diffZ + wL * fov - fov * x111 = wR * z222 + wR * diffZ + wR * fov

			fov * x222 + wL * z111 + wL * fov - fov * x111 - wR * z222 - wR * fov =  wR * diffZ - wL * diffZ*/


			let diffZ = (fov * x222 + wL * z111 + wL * fov - fov * x111 - wR * z222 - wR * fov) / (wR - wL);
			let diffX = (wL * z111 + wL * diffZ + wL * fov - fov * x111) / fov;


			return {minZ: diffZ, diffX: diffX};
		};

		let diffXZ3 = getDiffXZ3(x1, x2, z1, z2, y1, y2);

		return {minZ: diffXZ3.minZ, diffX: diffXZ3.diffX};
	};

	Processor3D.prototype._correctZPositionOY = function (x1, x2, z1, z2, minZ, y1, y2) {
		let t = this;

		let getDiffXZ3 = function (x1, x2, z1, z2, y1, y2) {
			let h = t.heightCanvas;

			let fov = 1 / t.rPerspective;

			let diffAndScalePoints = t.diffAndScale(x1, y1, z1);
			let x11 = diffAndScalePoints.x;
			let y11 = diffAndScalePoints.y;
			let z11 = diffAndScalePoints.z;


			let rotatePoints = t.rotate(x11, y11, z11);
			let y111 = rotatePoints.y;
			let z111 = rotatePoints.z;


			/*let x1111 = x111 + diffX;
			let y1111 = y111 + diffY;
			let z1111 = z111 + diffZ;
			let x11111 = (fov * x1111) / (z1111 + fov) + w / 2;
			let y11111 = (fov * y1111) / (z1111 + fov) + h / 2;*/


			//(fov * x1111) / (z1111 + fov) + w / 2 = t.left
			//(fov * (x111 + diffX)) / ((z111 + diffZ) + fov) + w / 2 = t.left
			//let wL = t.left - w / 2;
			//fov * (x111 + diffX) = wL * ((z111 + diffZ) + fov)
			//fov * x111 + fov * diffX = wL * z111 + wL * diffZ + wL * fov


			diffAndScalePoints = t.diffAndScale(x2, y2, z2);
			let x22 = diffAndScalePoints.x;
			let y22 = diffAndScalePoints.y;
			let z22 = diffAndScalePoints.z;


			rotatePoints = t.rotate(x22, y22, z22);
			let y222 = rotatePoints.y;
			let z222 = rotatePoints.z;


			/*let x1111 = x111 + diffX;
			let y1111 = y111 + diffY;
			let z1111 = z111 + diffZ;
			let x11111 = (fov * x1111) / (z1111 + fov) + w / 2;
			let y11111 = (fov * y1111) / (z1111 + fov) + h / 2;*/


			//(fov * (y111 + diffY)) / ((z111 + diffZ) + fov) + h / 2 = t.top
			//(fov * (y222 + diffY)) / ((z222 + diffZ) + fov) + h / 2 = h - t.bottom
			let wL = t.top - h / 2;
			let wR = h / 2 - t.bottom;

			/*(fov * (y111 + diffY)) / ((z111 + diffZ) + fov) = wL
			(fov * (y222 + diffY)) / ((z222 + diffZ) + fov) = wR


			fov * (y111 + diffY) = wL * ((z111 + diffZ) + fov)

			fov * y111 + fov * diffY = wL * z111 + wL * diffZ + wL * fov
			fov * y222 + fov * diffY = wR * z222 + wR * diffZ + wR * fov

			diffY = (wR * z222 + wR * diffZ + wR * fov - fov * y222) / fov

			fov * y111 + wR * z222 + wR * diffZ + wR * fov - fov * y222 = wL * z111 + wL * diffZ + wL * fov

			fov * y111 + wR * z222 + wR * fov - fov * y222  - wL * z111 - wL * fov = wL * diffZ - wR * diffZ*/

			let diffZ = (fov * y111 + wR * z222 + wR * fov - fov * y222 - wL * z111 - wL * fov) / (wL - wR);
			let diffY = (wR * z222 + wR * diffZ + wR * fov - fov * y222) / fov;


			//let diffZ = (fov * x222 + wL * z111 + wL * fov - fov * x111 - wR * z222 - wR * fov) / (wR - wL);
			//let diffX = (wL * z111 + wL * diffZ + wL * fov - fov * x111) / fov;


			return {minZ: diffZ, diffY: diffY};
		};

		let diffXZ3 = getDiffXZ3(x1, x2, z1, z2, y1, y2);

		return {minZ: diffXZ3.minZ, diffY: diffXZ3.diffY};
	};

	Processor3D.prototype._calculateCameraDiffX = function (minMaxOx) {
		//test ровно по центру, но циклом
		let maxLeftPoint = minMaxOx.left;
		let maxRightPoint = minMaxOx.right;

		//так ближе к тому, как смещает excel
		let widthCanvas = this.widthCanvas;
		let originalWidthChart = widthCanvas - this.left - this.right;

		let diffLeft = maxLeftPoint - this.left;
		let diffRight = (this.left + originalWidthChart) - maxRightPoint;

		this.cameraDiffX = (((diffRight - diffLeft) / 2) * (1 / (this.rPerspective) + this.cameraDiffZ)) / (1 / (this.rPerspective));
	};

	Processor3D.prototype._calculateCameraDiffY = function (maxTopPoint, maxBottomPoint) {
		let top = this.top;
		let bottom = this.bottom;
		let heightCanvas = this.heightCanvas;
		let heightChart = heightCanvas - top - bottom;

		let diffTop = maxTopPoint - top;
		let diffBottom = (heightChart + top) - maxBottomPoint;

		//this.cameraDiffY = this.top - minMaxOy.top; - для rAngAx
		this.cameraDiffY = (diffBottom - diffTop) / 2;
	};

	Processor3D.prototype._getMinMaxOx = function (points, faces) {
		let mostLeftPointX;
		let mostRightPointX;
		let xRight;
		let xLeft;

		for (let i = 0; i < faces.length - 1; i++) {
			for (let k = 0; k <= 3; k++) {

				let point1 = this.convertAndTurnPoint(points[faces[i][k]].x, points[faces[i][k]].y, points[faces[i][k]].z);
				//let point2 = this.convertAndTurnPoint(points[faces[i][k + 1]].x, points[faces[i][k + 1]].y, points[faces[i][k + 1]].z, true);
				let x1 = point1.x;
				//let x2 = point2.x;

				if (!xLeft) {
					xLeft = x1;
					mostLeftPointX = new Point3D(points[faces[i][k]].x, points[faces[i][k]].y, points[faces[i][k]].z, this.cChartDrawer);

					xRight = x1;
					mostRightPointX = new Point3D(points[faces[i][k]].x, points[faces[i][k]].y, points[faces[i][k]].z, this.cChartDrawer);
				} else {
					if (x1 < xLeft) {
						xLeft = x1;
						mostLeftPointX = new Point3D(points[faces[i][k]].x, points[faces[i][k]].y, points[faces[i][k]].z, this.cChartDrawer);
					}

					if (x1 > xRight) {
						xRight = x1;
						mostRightPointX = new Point3D(points[faces[i][k]].x, points[faces[i][k]].y, points[faces[i][k]].z, this.cChartDrawer);
					}
				}

			}
		}

		return {left: xLeft, right: xRight, mostLeftPointX: mostLeftPointX, mostRightPointX: mostRightPointX};
	};

	Processor3D.prototype._getMinMaxOy = function (points, faces) {
		let mostTopPointY;
		let mostBottomPointY;
		let xTop = this.top + this.heightCanvas;
		let xBottom = 0;
		for (let i = 0; i < faces.length - 1; i++) {
			for (let k = 0; k < 3; k++) {

				let point1 = this.convertAndTurnPoint(points[faces[i][k]].x, points[faces[i][k]].y, points[faces[i][k]].z);
				let point2 = this.convertAndTurnPoint(points[faces[i][k + 1]].x, points[faces[i][k + 1]].y, points[faces[i][k + 1]].z);
				let y1 = point1.y;
				let y2 = point2.y;


				if (y1 <= xTop) {
					xTop = y1;
					mostTopPointY = new Point3D(points[faces[i][k]].x, points[faces[i][k]].y, points[faces[i][k]].z, this.cChartDrawer);
				}

				if (y2 <= xTop) {
					xTop = y2;
					mostTopPointY = new Point3D(points[faces[i][k + 1]].x, points[faces[i][k + 1]].y, points[faces[i][k + 1]].z, this.cChartDrawer);
				}

				if (y1 >= xBottom) {
					xBottom = y1;
					mostBottomPointY = new Point3D(points[faces[i][k]].x, points[faces[i][k]].y, points[faces[i][k]].z, this.cChartDrawer);
				}

				if (y2 >= xBottom) {
					xBottom = y2;
					mostBottomPointY = new Point3D(points[faces[i][k + 1]].x, points[faces[i][k + 1]].y, points[faces[i][k + 1]].z, this.cChartDrawer);
				}

			}
		}

		return {top: xTop, bottom: xBottom, mostTopPointY: mostTopPointY, mostBottomPointY: mostBottomPointY};
	};

	Processor3D.prototype._calcAspectRatio = function () {
		//return width / height
		let widthOriginalChart = this.widthCanvas - (this.left + this.right);
		let heightOriginalChart = this.heightCanvas - (this.top + this.bottom);
		//auto scale if hPercent == null
		let hPercent = this.hPercent;
		let depthPercent = this._getDepthPercent();

		let aspectRatioX = 1;
		let aspectRatioY = 1;

		let countSeries = this.chartsDrawer.calculateFirstChartCountSeries();
		let seriesCount = countSeries.series;
		let ptCount = countSeries.points;

		let subType = this.chartsDrawer.calcProp.subType;
		if ((subType === "standard" || this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Line ||
				(this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Area && subType === "normal") || this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Surface) &&
			!this.view3D.getRAngAx()) {
			this._calcSpecialStandardScaleX();

			aspectRatioX = (widthOriginalChart / (heightOriginalChart / hPercent)) * this.specialStandardScaleX;
		} else if (subType === "standard" || this.chartsDrawer.calcProp.type === AscFormat.c_oChartTypes.Line || this.chartsDrawer.calcProp.type ===
			AscFormat.c_oChartTypes.Surface) {

			let depth = this.view3D.getRAngAx() ? this._calculateDepth() : this._calculateDepthPerspective();
			let width = (depth / depthPercent) * (ptCount / seriesCount);

			aspectRatioX = (widthOriginalChart) / width;
		} else if (hPercent !== null)//auto scale height
		{
			aspectRatioX = widthOriginalChart / (heightOriginalChart / hPercent);
		}

		if (aspectRatioX < 1 && this.view3D.getRAngAx()) {
			aspectRatioY = 1 / aspectRatioX;
			aspectRatioX = 1;
		}

		this.aspectRatioX = aspectRatioX;
		this.aspectRatioY = aspectRatioY;
	};

	Processor3D.prototype._getDepthPercent = function () {
		if (AscFormat.c_oChartTypes.Pie === this.chartsDrawer.calcProp.type) {
			return globalBasePercent / 100;
		}
		return this.view3D && this.view3D.depthPercent ? this.view3D.depthPercent / 100 : globalBasePercent / 100;
	};


	function Point3D(x, y, z, chartsDrawer) {
		this.x = x;
		this.y = y;
		this.z = z;
		this.t = 1;

		if (chartsDrawer) {
			this.chartProp = chartsDrawer.calcProp;
			this.cChartDrawer = chartsDrawer;
			this.cChartSpace = chartsDrawer.cChartSpace;
		}
	}

	Point3D.prototype = {
		constructor: Point3D,

		rotate: function (angleOx, angleOy, angleOz) {
			let x = this.x;
			let y = this.y;
			let z = this.z;
			let cosy = Math.cos(angleOy);
			let siny = Math.sin(angleOy);
			let cosx = Math.cos(angleOx);
			let sinx = Math.sin(angleOx);
			let cosz = Math.cos(angleOz);
			let sinz = Math.sin(angleOz);


			let newX = cosy * (sinz * y + cosz * x) - siny * z;
			let newY = sinx * (cosy * z + siny * (sinz * y + cosz * x)) + cosx * (cosz * y - sinz * x);
			let newZ = cosx * (cosy * z + siny * (sinz * y + cosz * x)) - sinx * (cosz * y - sinz * x);

			this.x = newX;
			this.y = newY;
			this.z = newZ;

			return this;
		},

		project: function (matrix) {
			//умножаем
			let projectPoint = matrix.multiplyPoint(this);

			//делим на 4 коэффициэнт
			let newX = projectPoint.x / projectPoint.t;
			let newY = projectPoint.y / projectPoint.t;

			this.x = newX;
			this.y = newY;

			return this;
		},

		offset: function (offsetX, offsetY, offsetZ) {
			this.x = this.x + offsetX;
			this.y = this.y + offsetY;
			this.z = this.z + offsetZ;

			return this;
		},

		scale: function (aspectRatioOx, aspectRatioOy, aspectRatioOz) {
			this.x = this.x / aspectRatioOx;
			this.y = this.y / aspectRatioOy;
			this.z = this.z / aspectRatioOz;

			return this;
		},

		multiplyPointOnMatrix1: function (matrix) {
			let multiplyMatrix = matrix.multiplyPoint(this);
			this.x = multiplyMatrix.x;
			this.y = multiplyMatrix.y;
			this.z = multiplyMatrix.z;
		}
	};

	function Matrix4D() {
		this.a11 = 1.0;
		this.a12 = 0.0;
		this.a13 = 0.0;
		this.a14 = 0.0;

		this.a21 = 0.0;
		this.a22 = 1.0;
		this.a23 = 0.0;
		this.a24 = 0.0;

		this.a31 = 0.0;
		this.a32 = 0.0;
		this.a33 = 1.0;
		this.a34 = 0.0;

		this.a41 = 0.0;
		this.a42 = 0.0;
		this.a43 = 0.0;
		this.a44 = 1.0;

		return this;
	}

	Matrix4D.prototype.reset = function () {
		this.a11 = 1.0;
		this.a12 = 0.0;
		this.a13 = 0.0;
		this.a14 = 0.0;

		this.a21 = 0.0;
		this.a22 = 1.0;
		this.a23 = 0.0;
		this.a24 = 0.0;

		this.a31 = 0.0;
		this.a32 = 0.0;
		this.a33 = 1.0;
		this.a34 = 0.0;

		this.a41 = 0.0;
		this.a42 = 0.0;
		this.a43 = 0.0;
		this.a44 = 1.0;
	};
	Matrix4D.prototype.multiply = function (m) {
		let res = new Matrix4D();

		res.a11 = this.a11 * m.a11 + this.a12 * m.a21 + this.a13 * m.a31 + this.a14 * m.a41;
		res.a12 = this.a11 * m.a12 + this.a12 * m.a22 + this.a13 * m.a32 + this.a14 * m.a42;
		res.a13 = this.a11 * m.a13 + this.a12 * m.a23 + this.a13 * m.a33 + this.a14 * m.a43;
		res.a14 = this.a11 * m.a14 + this.a12 * m.a24 + this.a13 * m.a34 + this.a14 * m.a44;

		res.a21 = this.a21 * m.a11 + this.a22 * m.a21 + this.a23 * m.a31 + this.a24 * m.a41;
		res.a22 = this.a21 * m.a12 + this.a22 * m.a22 + this.a23 * m.a32 + this.a24 * m.a42;
		res.a23 = this.a21 * m.a13 + this.a22 * m.a23 + this.a23 * m.a33 + this.a24 * m.a43;
		res.a24 = this.a21 * m.a14 + this.a22 * m.a24 + this.a23 * m.a34 + this.a24 * m.a44;

		res.a31 = this.a31 * m.a11 + this.a32 * m.a21 + this.a33 * m.a31 + this.a34 * m.a41;
		res.a32 = this.a31 * m.a12 + this.a32 * m.a22 + this.a33 * m.a32 + this.a34 * m.a42;
		res.a33 = this.a31 * m.a13 + this.a32 * m.a23 + this.a33 * m.a33 + this.a34 * m.a43;
		res.a34 = this.a31 * m.a14 + this.a32 * m.a24 + this.a33 * m.a34 + this.a34 * m.a44;

		res.a41 = this.a41 * m.a11 + this.a42 * m.a21 + this.a43 * m.a31 + this.a44 * m.a41;
		res.a42 = this.a41 * m.a12 + this.a42 * m.a22 + this.a43 * m.a32 + this.a44 * m.a42;
		res.a43 = this.a41 * m.a13 + this.a42 * m.a23 + this.a43 * m.a33 + this.a44 * m.a43;
		res.a44 = this.a41 * m.a14 + this.a42 * m.a24 + this.a43 * m.a34 + this.a44 * m.a44;

		return res;
	};
	Matrix4D.prototype.multiplyPoint = function (p) {
		let res = new Point3D();

		res.x = p.x * this.a11 + p.y * this.a21 + p.z * this.a31 + this.a41;
		res.y = p.x * this.a12 + p.y * this.a22 + p.z * this.a32 + this.a42;
		res.z = p.x * this.a13 + p.y * this.a23 + p.z * this.a33 + this.a43;
		res.t = p.x * this.a14 + p.y * this.a24 + p.z * this.a34 + this.a44;

		return res;
	};


//sort parallalepiped faces
	function CSortFaces(cChartDrawer, centralViewPoint) {
		this.cChartDrawer = cChartDrawer;
		this.chartProp = cChartDrawer.calcProp;
		this.centralViewPoint = null;
		this._initProperties(centralViewPoint)
	}

	CSortFaces.prototype = {
		constructor: CSortFaces,

		_initProperties: function (centralViewPoint) {
			if (!centralViewPoint) {
				let top = this.chartProp.chartGutter._top;
				let bottom = this.chartProp.chartGutter._bottom;
				let left = this.chartProp.chartGutter._left;
				let right = this.chartProp.chartGutter._right;

				let widthCanvas = this.chartProp.widthCanvas;
				let heightCanvas = this.chartProp.heightCanvas;
				let heightChart = heightCanvas - top - bottom;
				let widthChart = widthCanvas - left - right;

				let centerChartY = top + heightChart / 2;
				let centerChartX = left + widthChart / 2;

				let diffX = centerChartX;
				let diffY = centerChartY;
				let diffZ = -1 / this.cChartDrawer.processor3D.rPerspective;
				//TODO протестировать на bar и затем сделать для всех перспективных диаграмм данный сдвиг
				if (!this.cChartDrawer.processor3D.view3D.getRAngAx() && this.chartProp.type === AscFormat.c_oChartTypes.Bar) {
					diffX -= this.cChartDrawer.processor3D.cameraDiffX;
					diffY -= this.cChartDrawer.processor3D.cameraDiffY;
					diffZ -= this.cChartDrawer.processor3D.cameraDiffZ;
				}

				//TODO пересмотреть правку!
				if (diffZ > 0 && this.chartProp.type === AscFormat.c_oChartTypes.Bar && this.cChartDrawer.processor3D.view3D.getRAngAx() &&
					(this.chartProp.subType === "stackedPer" || this.chartProp.subType === "stacked")) {
					diffZ -= 500;
				}

				this.centralViewPoint = {x: diffX, y: diffY, z: diffZ};
			} else {
				this.centralViewPoint = centralViewPoint;
			}
		},

		sortParallelepipeds: function (parallelepipeds) {
			let intersectionsParallelepipeds = this._getIntersectionsParallelepipeds(parallelepipeds);
			let revIntersections = intersectionsParallelepipeds.reverseIntersections;
			let countIntersection = intersectionsParallelepipeds.countIntersection;

			let startIndexes = [];
			for (let i = 0; i < countIntersection.length; i++) {
				if (countIntersection[i] === 0) {
					startIndexes.push({index: parseInt(i)});
				}
			}

			if (startIndexes.length === 0) {
				//TODO сделано для графиков типа stacked, когда пересечения найдены всех параллалеп.
				return this._getSortZIndexArray(parallelepipeds);
			}

			let g = revIntersections;
			let color = {}; // цвет вершины (0, 1, или 2)
			let time_in = {}, time_out = {}; // "времена" захода и выхода из вершины
			let dfs_timer = 0;

			let dfs = function (v) {
				time_in[v] = dfs_timer++;
				color[v] = 1;

				for (let i in g[v]) {
					if (!color[i]) {
						dfs(i);
					}
				}

				color[v] = 2;
				time_out[v] = dfs_timer++;
			};
			let getMax = function () {
				let max = 0;
				let res = null;
				for (let i in time_out) {
					if (!addIndexes[i] && time_out[i] > max) {
						max = time_out[i];
						res = i;
					}
				}
				return res;
			};

			for (let i = 0; i < startIndexes.length; i++) {
				dfs(startIndexes[i].index);
			}

			let addIndexes = {};
			let res = [];
			while (true) {
				let index = getMax();

				if (null === index) {
					break;
				}

				res.push({nextIndex: index});
				addIndexes[index] = 1;
			}
			if (res.length < parallelepipeds.length) {
				return this._getSortZIndexArray(parallelepipeds);
			}

			return res.reverse();
		},

		_getSortZIndexArray: function (arr) {
			arr.sort(function sortArr(a, b) {
				return b.z - a.z;
			});
			let result = [];
			for (let i = 0; i < arr.length; i++) {
				result.push({nextIndex: i});
			}
			return result;
		},

		_getIntersectionsParallelepipeds: function (parallelepipeds) {
			//TODO нужно по максимуму оптимизировать, при большом количестве значений работает медленно
			let usuallyIntersect = {};
			let usuallyIntersectRev = {};

			let intersections = [];
			let reverseIntersections = [];
			let countIntersection = [];
			for (let i = 0; i < parallelepipeds.length; i++) {
				//из каждой точки данного параллалепипеда строим прямые до точки наблюдателя
				let fromParallalepiped = parallelepipeds[i];
				countIntersection[i] = 0;

				for (let m = 0; m < parallelepipeds.length; m++) {
					if (m === i || usuallyIntersect[i] === m || usuallyIntersectRev[i] === m) {
						continue;
					}
					if (parallelepipeds[m].isValZero && this.cChartDrawer.calcProp.subType !== "normal") {
						continue;
					}

					for (let l = 0; l < fromParallalepiped.arrPoints.length; l++) {
						//перебираем точки параллалепипеда FROM
						let point = fromParallalepiped.arrPoints[l];

						//перебираем грани параллалепипеда TO и проверяем на пересечения с прямыми из параллалепипеда FROM
						let toParallalepiped = parallelepipeds[m];
						for (let n = 0; n < toParallalepiped.faces.length; n++) {
							let toVerge = toParallalepiped.faces[n];
							let lineEqucation = this.cChartDrawer.getLineEquation(point, this.centralViewPoint);
							let intersection = this._isIntersectionFaceAndLine(toVerge, lineEqucation, point);
							if (intersection) {
								usuallyIntersect[i] = m;
								usuallyIntersectRev[m] = i;

								if (!intersections[i]) {
									intersections[i] = [];
								}

								if (!reverseIntersections[m]) {
									reverseIntersections[m] = [];
								}

								countIntersection[i]++;

								intersections[i][m] = 1;
								reverseIntersections[m][i] = 1;

								l = point.length;
								//console.log("fromParallalepiped: " + i + " toParallalepiped " + m + " x: " + intersection.x  + " y: " + intersection.y  + " z: " + intersection.z);

								break;
							}
						}
					}
				}
			}

			return {
				intersections: intersections, reverseIntersections: reverseIntersections, countIntersection: countIntersection
			};
		},

		//смотрим есть ли пересечения точек, выходящих из вершин данной грани, с другими гранями
		_isIntersectionFacesPointLines: function (plainVerge, i, sortZIndexPaths) {
			let res = false;
			let t = this;

			for (let j = 0; j < plainVerge.points.length; j++) {
				let pointFromVerge = plainVerge.points[j];
				let lineEqucation = t.cChartDrawer.getLineEquation(pointFromVerge, this.centralViewPoint);

				//пересечение грани и прямой
				if (t._isIntersectionFacesAndLine(lineEqucation, pointFromVerge, i, j, sortZIndexPaths)) {
					res = true;
					break;
				}
			}

			return res;
		},

		//пересечение прямой с другими гранями
		_isIntersectionFacesAndLine: function (lineEqucation, pointFromVerge, i, j, sortZIndexPaths) {
			let res = false;

			for (let k = 0; k < sortZIndexPaths.length; k++) {
				if (k === i) {
					continue;
				}

				if (this._isIntersectionFaceAndLine(sortZIndexPaths[k], lineEqucation, pointFromVerge, i, j, k)) {
					res = true;
					break;
				}
			}

			return res;
		},

		//пересечение прямой гранью
		_isIntersectionFaceAndLine: function (plain, lineEquation, pointFromVerge) {
			let res = false;
			let t = this;

			//get intersection
			let nIntersectionPlainAndLine = t.cChartDrawer.isIntersectionPlainAndLine(plain.plainEquation, lineEquation);

			if (null === nIntersectionPlainAndLine) {
				return res;
			}

			let iSX = nIntersectionPlainAndLine.x;
			let iSY = nIntersectionPlainAndLine.y;
			let iSZ = nIntersectionPlainAndLine.z;

			if (Math.round(iSZ * 1000) / 1000 < Math.round(pointFromVerge.z * 1000) / 1000) {
				let minMaxpoints = t.cChartDrawer.getMinMaxPoints(plain.points);
				let minX = minMaxpoints.minX, maxX = minMaxpoints.maxX, minY = minMaxpoints.minY, maxY = minMaxpoints.maxY, minZ = minMaxpoints.minZ, maxZ = minMaxpoints.maxZ;

				let isBeetwenX = t._isBetweenPoint(iSX, minX, maxX);
				let isBeetwenY = t._isBetweenPoint(iSY, minY, maxY);
				let isBeetwenZ = t._isBetweenPoint(iSZ, minZ, maxZ);

				if (isBeetwenX && isBeetwenY && isBeetwenZ) {
					let point0 = plain.points2[0];
					let point1 = plain.points2[1];
					let point2 = plain.points2[2];
					let point3 = plain.points2[3];

					let projectIntersection = nIntersectionPlainAndLine;
					if (!this.cChartDrawer.processor3D.view3D.getRAngAx()) {
						projectIntersection =
							t.cChartDrawer._convertAndTurnPoint(nIntersectionPlainAndLine.x, nIntersectionPlainAndLine.y, nIntersectionPlainAndLine.z, true, true);
					}

					let areaQuadrilateral = plain.plainArea;
					let areaTriangle1 = t.cChartDrawer.getAreaTriangle(point0, projectIntersection, point1);
					let areaTriangle2 = t.cChartDrawer.getAreaTriangle(point1, projectIntersection, point2);
					let areaTriangle3 = t.cChartDrawer.getAreaTriangle(point2, projectIntersection, point3);
					let areaTriangle4 = t.cChartDrawer.getAreaTriangle(point3, projectIntersection, point0);

					if (parseInt(areaQuadrilateral) === parseInt(areaTriangle1 + areaTriangle2 + areaTriangle3 + areaTriangle4)) {
						//let pointFromVergeProject = t.cChartDrawer._convertAndTurnPoint(pointFromVerge.x, pointFromVerge.y, pointFromVerge.z, true, true);
						//console.log("x: " + projectIntersection.x + " ;y: " + projectIntersection.y + " ;fromX:" + pointFromVergeProject.x + " ;fromY:" + pointFromVergeProject.y);

						res = nIntersectionPlainAndLine;
					}
				}
			}

			return res;
		},

		_isBetweenPoint: function (point, start, end) {
			//TODO округление пересмотреть
			let res = false;
			point = Math.round(point * 100) / 100;
			start = Math.round(start * 100) / 100;
			end = Math.round(end * 100) / 100;

			if (point >= start && point <= end) {
				res = true;
			}

			return res;
		},

		_isEqualPoints: function (point1, point2) {
			//TODO округление пересмотреть
			let res = false;

			if (parseInt(point1.x) === parseInt(point2.x) && parseInt(point1.y) === parseInt(point2.y) && parseInt(point1.y) === parseInt(point2.y)) {
				res = true;
			}

			return res;
		}
	};

	//----------------------------------------------------------export----------------------------------------------------
	window['AscFormat'] = window['AscFormat'] || {};
	window['AscFormat'].Processor3D = Processor3D;
	window['AscFormat'].Point3D = Point3D;
	window['AscFormat'].CSortFaces = CSortFaces;
	window['AscFormat'].global3DPersperctive = global3DPersperctive;
	window['AscFormat'].globalBasePercent = globalBasePercent;
	window['AscCommon'].c_oChartFloorPosition = c_oChartFloorPosition;
})(window);

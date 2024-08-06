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

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {

	const drawingsChangesMap = window['AscDFH'].drawingsChangesMap;
	const drawingContentChanges = window['AscDFH'].drawingContentChanges;

	const CChangesDrawingsBool = AscDFH.CChangesDrawingsBool;
	const CChangesDrawingsLong = AscDFH.CChangesDrawingsLong;
	const CChangesDrawingsString = AscDFH.CChangesDrawingsString;
	const CChangesDrawingsContent = AscDFH.CChangesDrawingsContent;
	const CChangesDrawingsObject = AscDFH.CChangesDrawingsObject;
	const CChangesDrawingsObjectNoId = AscDFH.CChangesDrawingsObjectNoId;
	const CChangesDrawingsDouble2 = AscDFH.CChangesDrawingsDouble2;

	// Address
	drawingsChangesMap[AscDFH.historyitem_Address_SetAddress1] = function (oClass, value) {
		oClass.address1 = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Address_SetCountryRegion] = function (oClass, value) {
		oClass.countryRegion = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Address_SetAdminDistrict1] = function (oClass, value) {
		oClass.adminDistrict1 = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Address_SetAdminDistrict2] = function (oClass, value) {
		oClass.adminDistrict2 = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Address_SetPostalCode] = function (oClass, value) {
		oClass.postalCode = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Address_SetLocality] = function (oClass, value) {
		oClass.locality = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Address_SetISOCountryCode] = function (oClass, value) {
		oClass.isoCountryCode = value;
	};

	AscDFH.changesFactory[AscDFH.historyitem_Address_SetAddress1] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_Address_SetCountryRegion] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_Address_SetAdminDistrict1] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_Address_SetAdminDistrict2] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_Address_SetPostalCode] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_Address_SetLocality] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_Address_SetISOCountryCode] = window['AscDFH'].CChangesDrawingsString;

	const SERIES_LAYOUT_BOX_WHISKER = 0;
	const SERIES_LAYOUT_CLUSTERED_COLUMN = 1;
	const SERIES_LAYOUT_FUNNEL = 2;
	const SERIES_LAYOUT_PARETO_LINE = 3;
	const SERIES_LAYOUT_REGION_MAP = 4;
	const SERIES_LAYOUT_SUNBURST = 5;
	const SERIES_LAYOUT_TREEMAP = 6;
	const SERIES_LAYOUT_WATERFALL = 7;

	const DATA_LABEL_POS_BEST_FIT = 0;
	const DATA_LABEL_POS_B = 1;
	const DATA_LABEL_POS_CTR = 2;
	const DATA_LABEL_POS_IN_BASE = 3;
	const DATA_LABEL_POS_IN_END = 4;
	const DATA_LABEL_POS_L = 5;
	const DATA_LABEL_POS_OUT_END = 6;
	const DATA_LABEL_POS_R = 7;
	const DATA_LABEL_POS_T = 8;

	const PARENT_LABEL_LAYOUT_NONE = 0;
	const PARENT_LABEL_LAYOUT_BANNER = 1;
	const PARENT_LABEL_LAYOUT_OVERLAPPING = 2;

	const REGION_LABEL_LAYOUT_NONE = 0;
	const REGION_LABEL_LAYOUT_BEST_FIT_ONLY = 1;
	const REGION_LABEL_LAYOUT_SHOW_ALL = 2;

	const INTERVAL_CLOSED_SIDE_L = 0;
	const INTERVAL_CLOSED_SIDE_R = 1;

	const AXIS_UNIT_HUNDREDS = 0;
	const AXIS_UNIT_THOUSANDS = 1;
	const AXIS_UNIT_TEN_THOUSANDS = 2;
	const AXIS_UNIT_HUNDRED_THOUSANDS = 3;
	const AXIS_UNIT_MILLIONS = 4;
	const AXIS_UNIT_TEN_MILLIONS = 5;
	const AXIS_UNIT_HUNDRED_MILLIONS = 6;
	const AXIS_UNIT_BILLIONS = 7;
	const AXIS_UNIT_TRILLIONS = 8;
	const AXIS_UNIT_PERCENTAGE = 9;

	const SIDE_POS_L = 0;
	const SIDE_POS_T = 1;
	const SIDE_POS_R = 2;
	const SIDE_POS_B = 3;

	const POS_ALIGN_MIN = 0;
	const POS_ALIGN_CTR = 1;
	const POS_ALIGN_MAX = 2;

	const TICK_MARKS_TYPE_IN = 0;
	const TICK_MARKS_TYPE_OUT = 1;
	const TICK_MARKS_TYPE_CROSS = 2;
	const TICK_MARKS_TYPE_NONE = 3;

	const QUARTILE_METHOD_INCLUSIVE = 0;
	const QUARTILE_METHOD_EXCLUSIVE = 1;

	const STRING_DIMENSION_TYPE_CAT = 10;
	const STRING_DIMENSION_TYPE_COLOR_STR = 11;

	const NUMERIC_DIMENSION_TYPE_VAL = 0;
	const NUMERIC_DIMENSION_TYPE_X = 1;
	const NUMERIC_DIMENSION_TYPE_Y = 2;
	const NUMERIC_DIMENSION_TYPE_SIZE = 3;
	const NUMERIC_DIMENSION_TYPE_COLOR_VAL = 4;

	const FORMULA_DIRECTION_COL = 0;
	const FORMULA_DIRECTION_ROW = 1;

// Import
	const History = AscCommon.History;
	const InitClass = AscFormat.InitClass;
	const CBaseChartObject = AscFormat.CBaseChartObject;
	const CAxisBase = AscFormat.CAxisBase;

	function CAddress() {
		CBaseChartObject.call(this);
		this.address1 = null;
		this.countryRegion = null;
		this.adminDistrict1 = null;
		this.adminDistrict2 = null;
		this.postalCode = null;
		this.locality = null;
		this.isoCountryCode = null;
	}

	InitClass(CAddress, CBaseChartObject, AscDFH.historyitem_type_Address);

	CAddress.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setAddress1(this.adress1);
		oCopy.setCountryRegion(this.countryRegion);
		oCopy.setAdminDistrict1(this.adminDistrict1);
		oCopy.setAdminDistrict2(this.adminDistrict2);
		oCopy.setPostalCode(this.postalCode);
		oCopy.setLocality(this.locality);
		oCopy.setISOCountryCode(this.isoCountryCode);
	};

	CAddress.prototype.setAddress1 = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Address_SetAddress1, this.address1, pr));
		this.address1 = pr;
	};
	CAddress.prototype.setCountryRegion = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Address_SetCountryRegion, this.countryRegion, pr));
		this.countryRegion = pr;
	};
	CAddress.prototype.setAdminDistrict1 = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Address_SetAdminDistrict1, this.adminDistrict1, pr));
		this.adminDistrict1 = pr;
	};
	CAddress.prototype.setAdminDistrict2 = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Address_SetAdminDistrict2, this.adminDistrict2, pr));
		this.adminDistrict2 = pr;
	};
	CAddress.prototype.setPostalCode = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Address_SetPostalCode, this.postalCode, pr));
		this.postalCode = pr;
	};
	CAddress.prototype.setLocality = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Address_SetLocality, this.locality, pr));
		this.locality = pr;
	};
	CAddress.prototype.setISOCountryCode = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Address_SetISOCountryCode, this.isoCountryCode, pr));
		this.isoCountryCode = pr;
	};

	drawingsChangesMap[AscDFH.historyitem_Axis_SetUnits] = function (oClass, value) {
		oClass.units = value;
	};

	drawingsChangesMap[AscDFH.historyitem_Axis_SetTickLabels] = function (oClass, value) {
		oClass.tickLabels = value;
	};

	drawingsChangesMap[AscDFH.historyitem_Axis_SetHidden] = function (oClass, value) {
		oClass.hidden = value;
	};

	AscDFH.changesFactory[AscDFH.historyitem_Axis_SetUnits] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Axis_SetTickLabels] = window['AscDFH'].CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_Axis_SetHidden] = window['AscDFH'].CChangesDrawingsBool;

	function CAxis() {
		CAxisBase.call(this);
		this.units = null;
		this.hidden = null;

		this.majorTickMark = null;
		this.minorTickMark = null;
	}

	InitClass(CAxis, CAxisBase, AscDFH.historyitem_type_Axis);

	CAxis.prototype.fillObject = function (oCopy) {
		CAxisBase.prototype.fillObject.call(this, oCopy);
		if (this.units) {
			oCopy.setUnits(this.units.createDuplicate());
		}
		oCopy.setTickLabels(this.tickLabels);
		oCopy.setHidden(this.hidden);
	};

	// isVertAxis => shows if axis is vertical
	CAxis.prototype.setAxPos = function (isVertAxis) {
		const axPos = isVertAxis ? window['AscFormat'].AX_POS_L : window['AscFormat'].AX_POS_B;
		CAxisBase.prototype.setAxPos.call(this, axPos);
	}

	CAxis.prototype.setUnits = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Axis_SetUnits, this.units, pr));
		this.units = pr;
	};

	CAxis.prototype.getFormatCode = function () {
		var oNumFmt = this.numFmt;
		var sFormatCode = null;

		if (oNumFmt) {
		if (oNumFmt.sourceLinked) {
				return this.getSourceFormatCode();
		}
			sFormatCode = oNumFmt.formatCode;
		if (typeof sFormatCode === "string" && sFormatCode.length > 0) {
				return sFormatCode;
		}
		}
		return "General";
	};
	CAxis.prototype.setTickLabels = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Axis_SetTickLabels, this.tickLabels, pr));
		this.tickLabels = pr;
	};
	CAxis.prototype.setHidden = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Axis_SetHidden, this.hidden, pr));
		this.hidden = pr;
	};
	CAxis.prototype.isValuesAxis = function () {
		return (this.scaling instanceof CValueAxisScaling);
	};

	// AxisUnits
	drawingsChangesMap[AscDFH.historyitem_AxisUnits_SetUnitsLabel] = function (oClass, value) {
		oClass.unitsLabel = value;
	};
	drawingsChangesMap[AscDFH.historyitem_AxisUnits_SetUnit] = function (oClass, value) {
		oClass.unit = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_AxisUnits_SetUnitsLabel] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_AxisUnits_SetUnit] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CAxisUnits() {
		CBaseChartObject.call(this);
		this.unitsLabel = null;
		this.unit = null;
	}

	InitClass(CAxisUnits, CBaseChartObject, AscDFH.historyitem_type_AxisUnits);

	CAxisUnits.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.unitsLabel) {
			oCopy.setUnitsLabel(this.unitsLabel.createDuplicate());
		}
		oCopy.setUnit(this.unit);
	};
	CAxisUnits.prototype.setUnitsLabel = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_AxisUnits_SetUnitsLabel, this.unitsLabel, pr));
		this.unitsLabel = pr;
		this.setParentToChild(pr);
	};
	CAxisUnits.prototype.setUnit = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_AxisUnits_SetUnit, this.unit, pr));
		this.unit = pr;
	};

	// AxisUnitsLabel
	drawingsChangesMap[AscDFH.historyitem_AxisUnitsLabel_SetTx] = function (oClass, value) {
		oClass.tx = value;
	};
	drawingsChangesMap[AscDFH.historyitem_AxisUnitsLabel_SetSpPr] = function (oClass, value) {
		oClass.spPr = value;
	};
	drawingsChangesMap[AscDFH.historyitem_AxisUnitsLabel_SetTxPr] = function (oClass, value) {
		oClass.txPr = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_AxisUnitsLabel_SetTx] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_AxisUnitsLabel_SetSpPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_AxisUnitsLabel_SetTxPr] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CAxisUnitsLabel() {
		CBaseChartObject.call(this);
		this.tx = null;
		this.spPr = null;
		this.txPr = null;
	}

	InitClass(CAxisUnitsLabel, CBaseChartObject, AscDFH.historyitem_type_AxisUnitsLabel);

	CAxisUnitsLabel.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.tx) {
			oCopy.setTx(this.tx.createDuplicate());
		}
		if (this.spPr) {
			oCopy.setSpPr(this.spPr.createDuplicate());
		}
		if (this.txPr) {
			oCopy.setTxPr(this.txPr.createDuplicate());
		}
	};

	CAxisUnitsLabel.prototype.setTx = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_AxisUnitsLabel_SetTx, this.tx, pr));
		this.tx = pr;
		this.setParentToChild(pr);
	};
	CAxisUnitsLabel.prototype.setSpPr = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_AxisUnitsLabel_SetSpPr, this.spPr, pr));
		this.spPr = pr;
		this.setParentToChild(pr);
	};
	CAxisUnitsLabel.prototype.setTxPr = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_AxisUnitsLabel_SetTxPr, this.txPr, pr));
		this.txPr = pr;
		this.setParentToChild(pr);
	};


	// Binning
	drawingsChangesMap[AscDFH.historyitem_Binning_SetBinSize] = function (oClass, value) {
		oClass.binSize = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Binning_SetBinCount] = function (oClass, value) {
		oClass.binCount = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Binning_SetIntervalClosed] = function (oClass, value) {
		oClass.intervalClosed = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Binning_SetUnderflow] = function (oClass, value) {
		oClass.underflow = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Binning_SetOverflow] = function (oClass, value) {
		oClass.overflow = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_Binning_SetBinSize] = window['AscDFH'].CChangesDrawingsDouble2;
	AscDFH.changesFactory[AscDFH.historyitem_Binning_SetBinCount] = window['AscDFH'].CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_Binning_SetIntervalClosed] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Binning_SetUnderflow] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Binning_SetOverflow] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CBinning() {
		CBaseChartObject.call(this);
		this.binSize = null;
		this.binCount = null;
		this.intervalClosed = null;
		this.underflow = null;
		this.overflow = null;
	}

	InitClass(CBinning, CBaseChartObject, AscDFH.historyitem_type_Binning);

	CBinning.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setBinSize(this.binSize);
		oCopy.setBinCount(this.binCount);
		oCopy.setIntervalClosed(this.intervalClosed);
		oCopy.setUnderflow(this.underflow);
		oCopy.setOverflow(this.overflow);
	};

	CBinning.prototype.setBinSize = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_Binning_SetBinSize, this.binSize, pr));
		this.binSize = pr;
	};
	CBinning.prototype.setBinCount = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Binning_SetBinCount, this.binCount, pr));
		this.binCount = pr;
	};
	CBinning.prototype.setIntervalClosed = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Binning_SetIntervalClosed, this.intervalClosed, pr));
		this.intervalClosed = pr;
	};
	CBinning.prototype.setUnderflow = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Binning_SetUnderflow, this.underflow, pr));
		this.underflow = pr;
	};
	CBinning.prototype.setOverflow = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Binning_SetOverflow, this.overflow, pr));
		this.overflow = pr;
	};


	// CategoryAxisScaling
	drawingsChangesMap[AscDFH.historyitem_CategoryAxisScaling_SetGapWidth] = function (oClass, value) {
		oClass.gapWidth = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_CategoryAxisScaling_SetGapWidth] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CCategoryAxisScaling() {
		CBaseChartObject.call(this);
		this.gapWidth = null;
	}

	InitClass(CCategoryAxisScaling, CBaseChartObject, AscDFH.historyitem_type_CategoryAxisScaling);

	CCategoryAxisScaling.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setGapWidth(this.gapWidth);
	};

	CCategoryAxisScaling.prototype.setGapWidth = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_CategoryAxisScaling_SetGapWidth, this.gapWidth, pr));
		this.gapWidth = pr;

	};



	// ChartData
	drawingsChangesMap[AscDFH.historyitem_ChartData_SetExternalData] = function (oClass, value) {
		oClass.externalData = value;
	};

	drawingContentChanges[AscDFH.historyitem_ChartData_AddData] =
		drawingContentChanges[AscDFH.historyitem_ChartData_RemoveData] = function (oClass) {
			return oClass.data;
		};
	AscDFH.changesFactory[AscDFH.historyitem_ChartData_SetExternalData] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_ChartData_AddData] = window['AscDFH'].CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_ChartData_RemoveData] = window['AscDFH'].CChangesDrawingsContent;

	function CChartData() {
		CBaseChartObject.call(this);
		this.externalData = null;
		this.data = [];
	}

	InitClass(CChartData, CBaseChartObject, AscDFH.historyitem_type_ChartData);

	CChartData.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.externalData) {
			oCopy.setExternalData(this.externalData.createDuplicate());
		}
		if (this.data) {
			for (let i = 0; i < this.data.length; i++) {
				oCopy.addData(this.data[i].createDuplicate(), i);
			}
		}
	};

	CChartData.prototype.setExternalData = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartData_SetExternalData, this.externalData, pr));
		this.externalData = pr;
		this.setParentToChild(pr);
	};
	CChartData.prototype.addData = function (pr, idx) {
		let pos;
		if (AscFormat.isRealNumber(idx))
			pos = idx;
		else
			pos = this.data.length;
		History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_ChartData_AddData, pos, [pr], true));
		this.data.splice(pos, 0, pr);
		this.setParentToChild(pr);
	};
	CChartData.prototype.removeDataByPos = function (pos) {
		if (this.data[pos]) {
			let data = this.data.splice(pos, 1)[0];
			History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_ChartData_RemoveData, pos, [data], false));
		}
	};
	CChartData.prototype.updateReferences = function (bDisplayEmptyCellsAs, bDisplayHidden) {
		for (let nData = 0; nData < this.data.length; ++nData) {
			this.data[nData].updateReferences(bDisplayEmptyCellsAs, bDisplayHidden);
		}
	};
	CChartData.prototype.getData = function (id) {
		for (let nData = 0; nData < this.data.length; ++nData) {
			let oData = this.data[nData];
			if (oData.id === id) {
				return oData;
			}
		}
		return null;
	};

	// Clear (no in ChartSerializeEx.h)
	drawingsChangesMap[AscDFH.historyitem_Clear_SetGeoLocationQueryResults] = function (oClass, value) {
		oClass.geoLocationQueryResults = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Clear_SetGeoDataEntityQueryResults] = function (oClass, value) {
		oClass.geoDataEntityQueryResults = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Clear_SetGeoDataPointToEntityQueryResults] = function (oClass, value) {
		oClass.geoDataPointToEntityQueryResults = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Clear_SetGeoChildEntitiesQueryResults] = function (oClass, value) {
		oClass.geoChildEntitiesQueryResults = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_Clear_SetGeoLocationQueryResults] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Clear_SetGeoDataEntityQueryResults] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Clear_SetGeoDataPointToEntityQueryResults] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Clear_SetGeoChildEntitiesQueryResults] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CClear() {
		CBaseChartObject.call(this);
		this.geoLocationQueryResults = null;
		this.geoDataEntityQueryResults = null;
		this.geoDataPointToEntityQueryResults = null;
		this.geoChildEntitiesQueryResults = null;
	}

	InitClass(CClear, CBaseChartObject, AscDFH.historyitem_type_Clear);

	CChartData.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoLocationQueryResults) {
			oCopy.setGeoLocationQueryResults(this.geoLocationQueryResults.createDuplicate());
		}
		if (this.geoDataEntityQueryResults) {
			oCopy.setGeoDataEntityQueryResults(this.geoDataEntityQueryResults.createDuplicate());
		}
		if (this.geoDataPointToEntityQueryResults) {
			oCopy.setGeoDataPointToEntityQueryResults(this.geoDataPointToEntityQueryResults.createDuplicate());
		}
		if (this.geoChildEntitiesQueryResults) {
			oCopy.setGeoChildEntitiesQueryResults(this.geoChildEntitiesQueryResults.createDuplicate());
		}
	};

	CClear.prototype.setGeoLocationQueryResults = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Clear_SetGeoLocationQueryResults, this.geoLocationQueryResults, pr));
		this.geoLocationQueryResults = pr;
		this.setParentToChild(pr);
	};
	CClear.prototype.setGeoDataEntityQueryResults = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Clear_SetGeoDataEntityQueryResults, this.geoDataEntityQueryResults, pr));
		this.geoDataEntityQueryResults = pr;
		this.setParentToChild(pr);
	};
	CClear.prototype.setGeoDataPointToEntityQueryResults = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Clear_SetGeoDataPointToEntityQueryResults, this.geoDataPointToEntityQueryResults, pr));
		this.geoDataPointToEntityQueryResults = pr;
		this.setParentToChild(pr);
	};
	CClear.prototype.setGeoChildEntitiesQueryResults = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Clear_SetGeoChildEntitiesQueryResults, this.geoChildEntitiesQueryResults, pr));
		this.geoChildEntitiesQueryResults = pr;
		this.setParentToChild(pr);
	};


	// Copyrights (no in ChartSerializeEx.h)
	drawingsChangesMap[AscDFH.historyitem_Copyrights_SetCopyright] = function (oClass, value) {
		oClass.copyright = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_Copyrights_SetCopyright] = window['AscDFH'].CChangesDrawingsString;

	function CCopyrights() {
		CBaseChartObject.call(this);
		this.copyright = null;
	}

	InitClass(CCopyrights, CBaseChartObject, AscDFH.historyitem_type_Copyrights);

	CCopyrights.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setCopyright(this.copyright);
	};

	CCopyrights.prototype.setCopyright = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Copyrights_SetCopyright, this.copyright, pr));
		this.copyright = pr;
	};


	// Data

	drawingsChangesMap[AscDFH.historyitem_Data_SetDimension] = function (oClass, value) {
		oClass.dimension = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Data_SetId] = function (oClass, value) {
		oClass.id = value;
	};

	drawingContentChanges[AscDFH.historyitem_Data_AddDimension] =
		drawingContentChanges[AscDFH.historyitem_Data_RemoveDimension] = function (oClass) {
			return oClass.dimension;
		};
	// AscDFH.changesFactory[AscDFH.historyitem_Data_SetNumDim] = window['AscDFH'].CChangesDrawingsString;
	// AscDFH.changesFactory[AscDFH.historyitem_Data_SetStrDim] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_Data_SetDimension] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Data_SetId] = window['AscDFH'].CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_Data_AddDimension] = window['AscDFH'].CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_Data_RemoveDimension] = window['AscDFH'].CChangesDrawingsContent;

	function CData() {
		CBaseChartObject.call(this);
		// field dimension instead of numDim,strDim   (ChartSerializeEx.h)
		// this.numDim = null;
		// this.strDim = null;
		this.dimension = [];
		this.id = null;
	}

	InitClass(CData, CBaseChartObject, AscDFH.historyitem_type_Data);

	CData.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.dimension) {
			for (let i = 0; i < this.dimension.length; i++) {
				oCopy.addDimension(this.dimension[i].createDuplicate(), i);
			}
		}
		oCopy.setId(this.id);
	};

	// CData.prototype.setNumDim = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Data_SetNumDim, this.numDim, pr));
	//     this.numDim = pr;
	// };
	// CData.prototype.setStrDim = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Data_SetStrDim, this.strDim, pr));
	//     this.strDim = pr;
	// };
	CData.prototype.addDimension = function (pr, idx) {
		let pos;
		if (AscFormat.isRealNumber(idx))
			pos = idx;
		else
			pos = this.dimension.length;
		History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_Data_AddDimension, pos, [pr], true));
		this.dimension.splice(pos, 0, pr);
		this.setParentToChild(pr);
	};
	CData.prototype.removeDimensionByPos = function (pos) {
		if (this.dimension[pos]) {
			let dimension = this.dimension.splice(pos, 1)[0];
			History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_Data_RemoveDimension, pos, [dimension], false));
		}
	};
	CData.prototype.setId = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Data_SetId, this.id, pr));
		this.id = pr;
		this.setParentToChild(pr);
	};
	CData.prototype.updateReferences = function (bDisplayEmptyCellsAs, bDisplayHidden) {
		for (let nDim = 0; nDim < this.dimension.length; ++nDim) {
			this.dimension[nDim].updateReferences(bDisplayEmptyCellsAs, bDisplayHidden)
		}
	};
	CData.prototype.getValDimensions = function () {
		let aRes = [];
		for (let nDim = 0; nDim < this.dimension.length; ++nDim) {
			let oDim = this.dimension[nDim];
			if ((oDim instanceof  CNumericDimension) && oDim.type === AscFormat.NUMERIC_DIMENSION_TYPE_VAL) {
				aRes.push(oDim);
			}
		}
		return aRes;
	};
	CData.prototype.getCatDimensions = function () {
		let aRes = [];
		for (let nDim = 0; nDim < this.dimension.length; ++nDim) {
			let oDim = this.dimension[nDim];
			if ((oDim instanceof  CStringDimension) && oDim.type === AscFormat.STRING_DIMENSION_TYPE_CAT) {
				aRes.push(oDim);
			}
		}
		return aRes;
	};


	// // DataId (int field instead class in ChartSerializeEx.h)
	// drawingsChangesMap[AscDFH.historyitem_DataId_SetVal] = function (oClass, value) {
	//     oClass.val = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_DataId_SetVal] = window['AscDFH'].CChangesDrawingsString;
	// function CDataId() {
	//     CBaseChartObject.call(this);
	//     this.val = null;
	// }

	// InitClass(CDataId, CBaseChartObject, AscDFH.historyitem_type_DataId);

	// CDataId.prototype.setVal = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_DataId_SetVal, this.val, pr));
	//     this.val = pr;
	// };
	//

	// DataLabel
	drawingsChangesMap[AscDFH.historyitem_DataLabel_SetNumFmt] = function (oClass, value) {
		oClass.numFmt = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabel_SetSpPr] = function (oClass, value) {
		oClass.spPr = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabel_SetTxPr] = function (oClass, value) {
		oClass.txPr = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabel_SetVisibility] = function (oClass, value) {
		oClass.visibility = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabel_SetSeparator] = function (oClass, value) {
		oClass.separator = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabel_SetIdx] = function (oClass, value) {
		oClass.idx = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabel_SetPos] = function (oClass, value) {
		oClass.pos = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_DataLabel_SetNumFmt] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabel_SetSpPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabel_SetTxPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabel_SetVisibility] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabel_SetSeparator] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabel_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabel_SetPos] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CDataLabel() {
		CBaseChartObject.call(this);
		this.numFmt = null;
		this.spPr = null;
		this.txPr = null;
		this.visibility = null;
		this.separator = null;
		this.idx = null;
		this.pos = null;
	}

	InitClass(CDataLabel, CBaseChartObject, AscDFH.historyitem_type_DataLabel);

	CDataLabel.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.numFmt) {
			oCopy.setNumFmt(this.numFmt.createDuplicate());
		}
		if (this.spPr) {
			oCopy.setSpPr(this.spPr.createDuplicate());
		}
		if (this.txPr) {
			oCopy.setTxPr(this.txPr.createDuplicate());
		}
		if (this.visibility) {
			oCopy.setVisibility(this.visibility.createDuplicate());
		}
		oCopy.setSeparator(this.separator);
		oCopy.setPos(this.pos);
		oCopy.setIdx(this.idx);
	};

	CDataLabel.prototype.setNumFmt = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DataLabel_SetNumFmt, this.numFmt, pr));
		this.numFmt = pr;
		this.setParentToChild(pr);
	};
	CDataLabel.prototype.setSpPr = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DataLabel_SetSpPr, this.spPr, pr));
		this.spPr = pr;
		this.setParentToChild(pr);
	};
	CDataLabel.prototype.setTxPr = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DataLabel_SetTxPr, this.txPr, pr));
		this.txPr = pr;
		this.setParentToChild(pr);
	};
	CDataLabel.prototype.setVisibility = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DataLabel_SetVisibility, this.visibility, pr));
		this.visibility = pr;
		this.setParentToChild(pr);
	};
	CDataLabel.prototype.setSeparator = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_DataLabel_SetSeparator, this.separator, pr));
		this.separator = pr;
		this.setParentToChild(pr);
	};
	CDataLabel.prototype.setIdx = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DataLabel_SetIdx, this.idx, pr));
		this.idx = pr;
		this.setParentToChild(pr);
	};
	CDataLabel.prototype.setPos = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DataLabel_SetPos, this.pos, pr));
		this.pos = pr;
		this.setParentToChild(pr);
	};


	// DataLabelHidden
	drawingsChangesMap[AscDFH.historyitem_DataLabelHidden_SetIdx] = function (oClass, value) {
		oClass.idx = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_DataLabelHidden_SetIdx] = window['AscDFH'].CChangesDrawingsLong;

	function CDataLabelHidden() {
		CBaseChartObject.call(this);
		this.idx = null;
	}

	InitClass(CDataLabelHidden, CBaseChartObject, AscDFH.historyitem_type_DataLabelHidden);

	CDataLabelHidden.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setIdx(this.idx);
	};

	CDataLabelHidden.prototype.setIdx = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DataLabelHidden_SetIdx, this.idx, pr));
		this.idx = pr;
		this.setParentToChild(pr);
	};


	// DataLabels
	drawingsChangesMap[AscDFH.historyitem_DataLabels_SetNumFmt] = function (oClass, value) {
		oClass.numFmt = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabels_SetSpPr] = function (oClass, value) {
		oClass.spPr = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabels_SetTxPr] = function (oClass, value) {
		oClass.txPr = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabels_SetVisibility] = function (oClass, value) {
		oClass.visibility = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabels_SetSeparator] = function (oClass, value) {
		oClass.separator = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabels_SetDataLabel] = function (oClass, value) {
		oClass.dataLabel = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabels_SetDataLabelHidden] = function (oClass, value) {
		oClass.dataLabelHidden = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabels_SetPos] = function (oClass, value) {
		oClass.pos = value;
	};

	drawingContentChanges[AscDFH.historyitem_DataLabels_AddDataLabel] =
		drawingContentChanges[AscDFH.historyitem_DataLabels_RemoveDataLabel] = function (oClass) {
			return oClass.dataLabel;
		};

	drawingContentChanges[AscDFH.historyitem_DataLabels_AddDataLabelHidden] =
		drawingContentChanges[AscDFH.historyitem_DataLabels_RemoveDataLabelHidden] = function (oClass) {
			return oClass.dataLabelHidden;
		};
	AscDFH.changesFactory[AscDFH.historyitem_DataLabels_SetNumFmt] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabels_SetSpPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabels_SetTxPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabels_SetVisibility] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabels_SetSeparator] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabels_SetDataLabel] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabels_SetDataLabelHidden] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabels_SetPos] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabels_AddDataLabel] = window['AscDFH'].CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabels_RemoveDataLabel] = window['AscDFH'].CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabels_AddDataLabelHidden] = window['AscDFH'].CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabels_RemoveDataLabelHidden] = window['AscDFH'].CChangesDrawingsContent;

	function CDataLabels() {
		CBaseChartObject.call(this);
		this.numFmt = null;
		this.spPr = null;
		this.txPr = null;
		this.visibility = null;
		this.separator = null;
		this.dataLabel = [];
		this.dataLabelHidden = [];
		this.pos = null;
	}

	InitClass(CDataLabels, CBaseChartObject, AscDFH.historyitem_type_DataLabels);

	CDataLabels.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.numFmt) {
			oCopy.setNumFmt(this.numFmt.createDuplicate());
		}
		if (this.spPr) {
			oCopy.setSpPr(this.spPr.createDuplicate());
		}
		if (this.txPr) {
			oCopy.setTxPr(this.txPr.createDuplicate());
		}
		if (this.visibility) {
			oCopy.setVisibility(this.visibility.createDuplicate());
		}
		if (this.dataLabel) {
			for (let i = 0; i < this.dataLabel.length; i++) {
				oCopy.addDataLabel(this.dataLabel[i].createDuplicate(), i);
			}
		}
		if (this.dataLabelHidden) {
			for (let i = 0; i < this.dataLabelHidden.length; i++) {
				oCopy.addDataLabelHidden(this.dataLabelHidden[i].createDuplicate(), i);
			}
		}
		oCopy.setPos(this.pos);
	};

	CDataLabels.prototype.setNumFmt = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DataLabels_SetNumFmt, this.numFmt, pr));
		this.numFmt = pr;
		this.setParentToChild(pr);
	};
	CDataLabels.prototype.setSpPr = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DataLabels_SetSpPr, this.spPr, pr));
		this.spPr = pr;
		this.setParentToChild(pr);
	};
	CDataLabels.prototype.setTxPr = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DataLabels_SetTxPr, this.txPr, pr));
		this.txPr = pr;
		this.setParentToChild(pr);
	};
	CDataLabels.prototype.setVisibility = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DataLabels_SetVisibility, this.visibility, pr));
		this.visibility = pr;
		this.setParentToChild(pr);
	};
	CDataLabels.prototype.setSeparator = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_DataLabels_SetSeparator, this.separator, pr));
		this.separator = pr;
		this.setParentToChild(pr);
	};
	CDataLabels.prototype.setDataLabel = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DataLabels_SetDataLabel, this.dataLabel, pr));
		this.dataLabel = pr;
		this.setParentToChild(pr);
	};
	CDataLabels.prototype.addDataLabel = function (pr, idx) {
		let pos;
		if (AscFormat.isRealNumber(idx))
			pos = idx;
		else
			pos = this.dataLabel.length;
		History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_DataLabels_AddDataLabel, pos, [pr], true));
		this.dataLabel.splice(pos, 0, pr);
		this.setParentToChild(pr);
	};
	CDataLabels.prototype.removeDataLabelByPos = function (pos) {
		if (this.dataLabel[pos]) {
			let dataLabel = this.dataLabel.splice(pos, 1)[0];
			History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_DataLabels_RemoveDataLabel, pos, [dataLabel], false));
		}
	};
	CDataLabels.prototype.addDataLabelHidden = function (pr, idx) {
		let pos;
		if (AscFormat.isRealNumber(idx))
			pos = idx;
		else
			pos = this.dataLabelHidden.length;
		History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_DataLabels_AddDataLabelHidden, pos, [pr], true));
		this.dataLabelHidden.splice(pos, 0, pr);
		this.setParentToChild(pr);
	};
	CDataLabels.prototype.removeDataLabelHiddenByPos = function (pos) {
		if (this.dataLabelHidden[pos]) {
			let dataLabelHidden = this.dataLabelHidden.splice(pos, 1)[0];
			History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_DataLabels_RemoveDataLabelHidden, pos, [dataLabelHidden], false));
		}
	};
	CDataLabels.prototype.setPos = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DataLabels_SetPos, this.pos, pr));
		this.pos = pr;
		this.setParentToChild(pr);
	};


	// DataLabelVisibilities
	drawingsChangesMap[AscDFH.historyitem_DataLabelVisibilities_SetSeriesName] = function (oClass, value) {
		oClass.seriesName = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabelVisibilities_SetCategoryName] = function (oClass, value) {
		oClass.categoryName = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataLabelVisibilities_SetValue] = function (oClass, value) {
		oClass.value = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_DataLabelVisibilities_SetSeriesName] = window['AscDFH'].CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabelVisibilities_SetCategoryName] = window['AscDFH'].CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_DataLabelVisibilities_SetValue] = window['AscDFH'].CChangesDrawingsBool;

	function CDataLabelVisibilities() {
		CBaseChartObject.call(this);
		this.seriesName = null;
		this.categoryName = null;
		this.value = null;
	}

	InitClass(CDataLabelVisibilities, CBaseChartObject, AscDFH.historyitem_type_DataLabelVisibilities);

	CDataLabelVisibilities.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setSeriesName(this.seriesName);
		oCopy.setCategoryName(this.categoryName);
		oCopy.setValue(this.value);
	};

	CDataLabelVisibilities.prototype.setSeriesName = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DataLabelVisibilities_SetSeriesName, this.seriesName, pr));
		this.seriesName = pr;
		this.setParentToChild(pr);
	};
	CDataLabelVisibilities.prototype.setCategoryName = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DataLabelVisibilities_SetCategoryName, this.categoryName, pr));
		this.categoryName = pr;
		this.setParentToChild(pr);
	};
	CDataLabelVisibilities.prototype.setValue = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DataLabelVisibilities_SetValue, this.value, pr));
		this.value = pr;
		this.setParentToChild(pr);
	};


	// DataPoint
	drawingsChangesMap[AscDFH.historyitem_DataPoint_SetSpPr] = function (oClass, value) {
		oClass.spPr = value;
	};
	drawingsChangesMap[AscDFH.historyitem_DataPoint_SetIdx] = function (oClass, value) {
		oClass.idx = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_DataPoint_SetSpPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_DataPoint_SetIdx] = window['AscDFH'].CChangesDrawingsLong;

	function CDataPoint() {
		CBaseChartObject.call(this);
		this.spPr = null;
		this.idx = null;
	}

	InitClass(CDataPoint, CBaseChartObject, AscDFH.historyitem_type_DataPoint);

	CDataPoint.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.spPr) {
			oCopy.setSpPr(this.spPr.createDuplicate());
		}
		oCopy.setIdx(this.idx);
	};

	CDataPoint.prototype.setSpPr = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_DataPoint_SetSpPr, this.spPr, pr));
		this.spPr = pr;
		this.setParentToChild(pr);
	};
	CDataPoint.prototype.setIdx = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DataPoint_SetIdx, this.idx, pr));
		this.idx = pr;
		this.setParentToChild(pr);
	};


	// // Extension (Extensionlist is unused in ChartSerializeEx.h)
	// drawingsChangesMap[AscDFH.historyitem_Extension_SetUri] = function (oClass, value) {
	//     oClass.uri = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_Extension_SetUri] = window['AscDFH'].CChangesDrawingsString;
	// function CExtension() {
	//     CBaseChartObject.call(this);
	//     this.uri = null;
	// }

	// InitClass(CExtension, CBaseChartObject, AscDFH.historyitem_type_Extension);

	// CExtension.prototype.setUri = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Extension_SetUri, this.uri, pr));
	//     this.uri = pr;
	// };
	//

	// // ExtensionList (OOX::Drawing::COfficeArtExtensionList instead of CExtensionList)
	// drawingsChangesMap[AscDFH.historyitem_ExtensionList_SetExt] = function (oClass, value) {
	//     oClass.ext = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_ExtensionList_SetExt] = window['AscDFH'].CChangesDrawingsString;
	// function CExtensionList() {
	//     CBaseChartObject.call(this);
	//     this.ext = null;
	// }

	// InitClass(CExtensionList, CBaseChartObject, AscDFH.historyitem_type_ExtensionList);

	// CExtensionList.prototype.setExt = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_ExtensionList_SetExt, this.ext, pr));
	//     this.ext = pr;
	// };
	//

	// // ExternalData (contains in ChartFormat.js)
	// drawingsChangesMap[AscDFH.historyitem_ExternalData_SetId] = function (oClass, value) {
	//     oClass.id = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_ExternalData_SetAutoUpdate] = function (oClass, value) {
	//     oClass.autoUpdate = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_ExternalData_SetId] = window['AscDFH'].CChangesDrawingsString;
	// AscDFH.changesFactory[AscDFH.historyitem_ExternalData_SetAutoUpdate] = window['AscDFH'].CChangesDrawingsBool;
	// function CExternalData() {
	//     CBaseChartObject.call(this);
	//     this.id = null;
	//     this.autoUpdate = null;
	// }

	// InitClass(CExternalData, CBaseChartObject, AscDFH.historyitem_type_ExternalData);

	// CExternalData.prototype.setId = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_ExternalData_SetId, this.id, pr));
	//     this.id = pr;
	// };
	// CExternalData.prototype.setAutoUpdate = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_ExternalData_SetAutoUpdate, this.autoUpdate, pr));
	//     this.autoUpdate = pr;
	// };
	//

	// FormatOverride (no in ChartSerializeEx.h)
	drawingsChangesMap[AscDFH.historyitem_FormatOverride_SetSpPr] = function (oClass, value) {
		oClass.spPr = value;
	};
	drawingsChangesMap[AscDFH.historyitem_FormatOverride_SetIdx] = function (oClass, value) {
		oClass.idx = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_FormatOverride_SetSpPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_FormatOverride_SetIdx] = window['AscDFH'].CChangesDrawingsLong;

	function CFormatOverride() {
		CBaseChartObject.call(this);
		this.spPr = null;
		this.idx = null;
	}

	InitClass(CFormatOverride, CBaseChartObject, AscDFH.historyitem_type_FormatOverride);

	CFormatOverride.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.spPr) {
			oCopy.setSpPr(this.spPr.createDuplicate());
		}
		oCopy.setIdx(this.idx);
	};

	CFormatOverride.prototype.setSpPr = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_FormatOverride_SetSpPr, this.spPr, pr));
		this.spPr = pr;
		this.setParentToChild(pr);
	};
	CFormatOverride.prototype.setIdx = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_FormatOverride_SetIdx, this.idx, pr));
		this.idx = pr;
		this.setParentToChild(pr);
	};


	// FormatOverrides (no in ChartSerializeEx.h)
	drawingsChangesMap[AscDFH.historyitem_FormatOverrides_SetFmtOvr] = function (oClass, value) {
		oClass.fmtOvr = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_FormatOverrides_SetFmtOvr] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CFormatOverrides() {
		CBaseChartObject.call(this);
		this.fmtOvr = null;
	}

	InitClass(CFormatOverrides, CBaseChartObject, AscDFH.historyitem_type_FormatOverrides);

	CFormatOverrides.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.fmtOvr) {
			oCopy.setFmtOvr(this.fmtOvr.createDuplicate());
		}
	};

	CFormatOverrides.prototype.setFmtOvr = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_FormatOverrides_SetFmtOvr, this.fmtOvr, pr));
		this.fmtOvr = pr;
		this.setParentToChild(pr);
	};


	// Formula
	drawingsChangesMap[AscDFH.historyitem_Formula_SetDir] = function (oClass, value) {
		oClass.dir = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Formula_SetContent] = function (oClass, value) {
		oClass.content = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_Formula_SetDir] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Formula_SetContent] = window['AscDFH'].CChangesDrawingsString;

	function CFormula() {
		CBaseChartObject.call(this);
		this.dir = null;
		this.content = null;
	}

	InitClass(CFormula, CBaseChartObject, AscDFH.historyitem_type_Formula);

	CFormula.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setDir(this.dir);
		oCopy.setContent(this.content);
	};

	CFormula.prototype.setDir = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Formula_SetDir, this.dir, pr));
		this.dir = pr;
		this.setParentToChild(pr);
	};
	CFormula.prototype.setContent = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Formula_SetContent, this.content, pr));
		this.content = pr;
		this.setParentToChild(pr);
	};


	// GeoCache
	drawingsChangesMap[AscDFH.historyitem_GeoCache_SetBinary] = function (oClass, value) {
		oClass.binary = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoCache_SetClear] = function (oClass, value) {
		oClass.clear = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoCache_SetProvider] = function (oClass, value) {
		oClass.provider = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoCache_SetBinary] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_GeoCache_SetClear] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_GeoCache_SetProvider] = window['AscDFH'].CChangesDrawingsString;

	function CGeoCache() {
		CBaseChartObject.call(this);
		this.binary = null;
		this.clear = null;
		this.provider = null;
	}

	InitClass(CGeoCache, CBaseChartObject, AscDFH.historyitem_type_GeoCache);

	CGeoCache.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.clear) {
			oCopy.setClear(this.clear.createDuplicate());
		}
		oCopy.setBinary(this.binary);
		oCopy.setProvider(this.provider);
	};

	CGeoCache.prototype.setBinary = function (pr) { // todo base64binary type
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoCache_SetBinary, this.binary, pr));
		this.binary = pr;
		this.setParentToChild(pr);
	};
	CGeoCache.prototype.setClear = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoCache_SetClear, this.clear, pr));
		this.clear = pr;
		this.setParentToChild(pr);
	};
	CGeoCache.prototype.setProvider = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoCache_SetProvider, this.provider, pr));
		this.provider = pr;
		this.setParentToChild(pr);
	};


	// GeoChildEntities
	drawingsChangesMap[AscDFH.historyitem_GeoChildEntities_SetGeoHierarchyEntity] = function (oClass, value) {
		oClass.geoHierarchyEntity = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoChildEntities_SetGeoHierarchyEntity] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoChildEntities() {
		CBaseChartObject.call(this);
		this.geoHierarchyEntity = null;
	}

	InitClass(CGeoChildEntities, CBaseChartObject, AscDFH.historyitem_type_GeoChildEntities);

	CGeoChildEntities.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoHierarchyEntity) {
			oCopy.setGeoHierarchyEntity(this.geoHierarchyEntity.createDuplicate());
		}
	};

	CGeoChildEntities.prototype.setGeoHierarchyEntity = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoChildEntities_SetGeoHierarchyEntity, this.geoHierarchyEntity, pr));
		this.geoHierarchyEntity = pr;
		this.setParentToChild(pr);
	};


	// GeoChildEntitiesQuery
	drawingsChangesMap[AscDFH.historyitem_GeoChildEntitiesQuery_SetGeoChildTypes] = function (oClass, value) {
		oClass.geoChildTypes = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoChildEntitiesQuery_SetEntityId] = function (oClass, value) {
		oClass.entityId = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoChildEntitiesQuery_SetGeoChildTypes] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_GeoChildEntitiesQuery_SetEntityId] = window['AscDFH'].CChangesDrawingsString;

	function CGeoChildEntitiesQuery() {
		CBaseChartObject.call(this);
		this.geoChildTypes = null;
		this.entityId = null;
	}

	InitClass(CGeoChildEntitiesQuery, CBaseChartObject, AscDFH.historyitem_type_GeoChildEntitiesQuery);

	CGeoChildEntitiesQuery.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoChildTypes) {
			oCopy.setGeoChildTypes(this.geoChildTypes.createDuplicate());
		}
		oCopy.setEntityId(this.entityId);
	};

	CGeoChildEntitiesQuery.prototype.setGeoChildTypes = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoChildEntitiesQuery_SetGeoChildTypes, this.geoChildTypes, pr));
		this.geoChildTypes = pr;
		this.setParentToChild(pr);
	};
	CGeoChildEntitiesQuery.prototype.setEntityId = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoChildEntitiesQuery_SetEntityId, this.entityId, pr));
		this.entityId = pr;
		this.setParentToChild(pr);
	};


	// GeoChildEntitiesQueryResult
	drawingsChangesMap[AscDFH.historyitem_GeoChildEntitiesQueryResult_SetGeoChildEntitiesQuery] = function (oClass, value) {
		oClass.geoChildEntitiesQuery = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoChildEntitiesQueryResult_SetGeoChildEntities] = function (oClass, value) {
		oClass.geoChildEntities = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoChildEntitiesQueryResult_SetGeoChildEntitiesQuery] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_GeoChildEntitiesQueryResult_SetGeoChildEntities] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoChildEntitiesQueryResult() {
		CBaseChartObject.call(this);
		this.geoChildEntitiesQuery = null;
		this.geoChildEntities = null;
	}

	InitClass(CGeoChildEntitiesQueryResult, CBaseChartObject, AscDFH.historyitem_type_GeoChildEntitiesQueryResult);
	
	CGeoChildEntitiesQueryResult.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoChildEntities) {
			oCopy.setGeoChildEntities(this.geoChildEntities.createDuplicate());
		}
		if (this.geoChildEntitiesQuery) {
			oCopy.setGeoChildEntitiesQuery(this.geoChildEntitiesQuery.createDuplicate());
		}
	};

	CGeoChildEntitiesQueryResult.prototype.setGeoChildEntitiesQuery = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoChildEntitiesQueryResult_SetGeoChildEntitiesQuery, this.geoChildEntitiesQuery, pr));
		this.geoChildEntitiesQuery = pr;
		this.setParentToChild(pr);
	};
	CGeoChildEntitiesQueryResult.prototype.setGeoChildEntities = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoChildEntitiesQueryResult_SetGeoChildEntities, this.geoChildEntities, pr));
		this.geoChildEntities = pr;
		this.setParentToChild(pr);
	};


	// GeoChildEntitiesQueryResults
	drawingsChangesMap[AscDFH.historyitem_GeoChildEntitiesQueryResults_SetGeoChildEntitiesQueryResult] = function (oClass, value) {
		oClass.geoChildEntitiesQueryResult = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoChildEntitiesQueryResults_SetGeoChildEntitiesQueryResult] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoChildEntitiesQueryResults() {
		CBaseChartObject.call(this);
		this.geoChildEntitiesQueryResult = null;
	}

	InitClass(CGeoChildEntitiesQueryResults, CBaseChartObject, AscDFH.historyitem_type_GeoChildEntitiesQueryResults);

	CGeoChildEntitiesQueryResults.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoChildEntitiesQueryResult) {
			oCopy.setGeoChildEntitiesQueryResult(this.geoChildEntitiesQueryResult.createDuplicate());
		}
	};

	CGeoChildEntitiesQueryResults.prototype.setGeoChildEntitiesQueryResult = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoChildEntitiesQueryResults_SetGeoChildEntitiesQueryResult, this.geoChildEntitiesQueryResult, pr));
		this.geoChildEntitiesQueryResult = pr;
		this.setParentToChild(pr);
	};


	// GeoChildTypes
	drawingsChangesMap[AscDFH.historyitem_GeoChildTypes_SetEntityType] = function (oClass, value) {
		oClass.entityType = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoChildTypes_SetEntityType] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoChildTypes() {
		CBaseChartObject.call(this);
		this.entityType = null;
	}

	InitClass(CGeoChildTypes, CBaseChartObject, AscDFH.historyitem_type_GeoChildTypes);

	CGeoChildTypes.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setEntityType(this.entityType);
	};

	CGeoChildTypes.prototype.setEntityType = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoChildTypes_SetEntityType, this.entityType, pr));
		this.entityType = pr;
		this.setParentToChild(pr);
	};


	// GeoData
	drawingsChangesMap[AscDFH.historyitem_GeoData_SetGeoPolygons] = function (oClass, value) {
		oClass.geoPolygons = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoData_SetCopyrights] = function (oClass, value) {
		oClass.copyrights = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoData_SetEntityName] = function (oClass, value) {
		oClass.entityName = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoData_SetEntityId] = function (oClass, value) {
		oClass.entityId = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoData_SetEast] = function (oClass, value) {
		oClass.east = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoData_SetWest] = function (oClass, value) {
		oClass.west = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoData_SetNorth] = function (oClass, value) {
		oClass.north = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoData_SetSouth] = function (oClass, value) {
		oClass.south = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoData_SetGeoPolygons] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_GeoData_SetCopyrights] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_GeoData_SetEntityName] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_GeoData_SetEntityId] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_GeoData_SetEast] = window['AscDFH'].CChangesDrawingsDouble2;
	AscDFH.changesFactory[AscDFH.historyitem_GeoData_SetWest] = window['AscDFH'].CChangesDrawingsDouble2;
	AscDFH.changesFactory[AscDFH.historyitem_GeoData_SetNorth] = window['AscDFH'].CChangesDrawingsDouble2;
	AscDFH.changesFactory[AscDFH.historyitem_GeoData_SetSouth] = window['AscDFH'].CChangesDrawingsDouble2;

	function CGeoData() {
		CBaseChartObject.call(this);
		this.geoPolygons = null;
		this.copyrights = null;
		this.entityName = null;
		this.entityId = null;
		this.east = null;
		this.west = null;
		this.north = null;
		this.south = null;
	}

	InitClass(CGeoData, CBaseChartObject, AscDFH.historyitem_type_GeoData);

	CGeoData.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoPolygons) {
			oCopy.setGeoPolygons(this.geoPolygons.createDuplicate());
		}
		if (this.setCopyrights) {
			oCopy.setGeoChildEntitiesQuery(this.copyrights.createDuplicate());
		}
		oCopy.setEntityName(this.entityName);
		oCopy.setEntityId(this.entityId);
		oCopy.setEast(this.east);
		oCopy.setWest(this.west);
		oCopy.setNorth(this.north);
		oCopy.setSouth(this.south);
	};

	CGeoData.prototype.setGeoPolygons = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoData_SetGeoPolygons, this.geoPolygons, pr));
		this.geoPolygons = pr;
		this.setParentToChild(pr);
	};
	CGeoData.prototype.setCopyrights = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoData_SetCopyrights, this.copyrights, pr));
		this.copyrights = pr;
		this.setParentToChild(pr);
	};
	CGeoData.prototype.setEntityName = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoData_SetEntityName, this.entityName, pr));
		this.entityName = pr;
		this.setParentToChild(pr);
	};
	CGeoData.prototype.setEntityId = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoData_SetEntityId, this.entityId, pr));
		this.entityId = pr;
		this.setParentToChild(pr);
	};
	CGeoData.prototype.setEast = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_GeoData_SetEast, this.east, pr));
		this.east = pr;
		this.setParentToChild(pr);
	};
	CGeoData.prototype.setWest = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_GeoData_SetWest, this.west, pr));
		this.west = pr;
		this.setParentToChild(pr);
	};
	CGeoData.prototype.setNorth = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_GeoData_SetNorth, this.north, pr));
		this.north = pr;
		this.setParentToChild(pr);
	};
	CGeoData.prototype.setSouth = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_GeoData_SetSouth, this.south, pr));
		this.south = pr;
		this.setParentToChild(pr);
	};


	// GeoDataEntityQuery
	drawingsChangesMap[AscDFH.historyitem_GeoDataEntityQuery_SetEntityType] = function (oClass, value) {
		oClass.entityType = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoDataEntityQuery_SetEntityId] = function (oClass, value) {
		oClass.entityId = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoDataEntityQuery_SetEntityType] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_GeoDataEntityQuery_SetEntityId] = window['AscDFH'].CChangesDrawingsString;

	function CGeoDataEntityQuery() {
		CBaseChartObject.call(this);
		this.entityType = null;
		this.entityId = null;
	}

	InitClass(CGeoDataEntityQuery, CBaseChartObject, AscDFH.historyitem_type_GeoDataEntityQuery);

	CGeoDataEntityQuery.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setEntityType(this.entityType);
		oCopy.setEntityId(this.entityId);
	};

	CGeoDataEntityQuery.prototype.setEntityType = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoDataEntityQuery_SetEntityType, this.entityType, pr));
		this.entityType = pr;
		this.setParentToChild(pr);
	};
	CGeoDataEntityQuery.prototype.setEntityId = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoDataEntityQuery_SetEntityId, this.entityId, pr));
		this.entityId = pr;
		this.setParentToChild(pr);
	};


	// GeoDataEntityQueryResult
	drawingsChangesMap[AscDFH.historyitem_GeoDataEntityQueryResult_SetGeoDataEntityQuery] = function (oClass, value) {
		oClass.geoDataEntityQuery = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoDataEntityQueryResult_SetGeoData] = function (oClass, value) {
		oClass.geoData = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoDataEntityQueryResult_SetGeoDataEntityQuery] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_GeoDataEntityQueryResult_SetGeoData] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoDataEntityQueryResult() {
		CBaseChartObject.call(this);
		this.geoDataEntityQuery = null;
		this.geoData = null;
	}

	InitClass(CGeoDataEntityQueryResult, CBaseChartObject, AscDFH.historyitem_type_GeoDataEntityQueryResult);

	CGeoDataEntityQueryResult.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoDataEntityQuery) {
			oCopy.setGeoDataEntityQuery(this.geoDataEntityQuery.createDuplicate());
		}
		if (this.geoData) {
			oCopy.setGeoData(this.geoData.createDuplicate());
		}
	}

	CGeoDataEntityQueryResult.prototype.setGeoDataEntityQuery = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoDataEntityQueryResult_SetGeoDataEntityQuery, this.geoDataEntityQuery, pr));
		this.geoDataEntityQuery = pr;
		this.setParentToChild(pr);
	};
	CGeoDataEntityQueryResult.prototype.setGeoData = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoDataEntityQueryResult_SetGeoData, this.geoData, pr));
		this.geoData = pr;
		this.setParentToChild(pr);
	};


	// GeoDataEntityQueryResults
	drawingsChangesMap[AscDFH.historyitem_GeoDataEntityQueryResults_SetGeoDataEntityQueryResult] = function (oClass, value) {
		oClass.geoDataEntityQueryResult = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoDataEntityQueryResults_SetGeoDataEntityQueryResult] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoDataEntityQueryResults() {
		CBaseChartObject.call(this);
		this.geoDataEntityQueryResult = null;
	}

	InitClass(CGeoDataEntityQueryResults, CBaseChartObject, AscDFH.historyitem_type_GeoDataEntityQueryResults);

	CGeoDataEntityQueryResults.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoDataEntityQueryResult) {
			oCopy.setGeoDataEntityQueryResult(this.geoDataEntityQueryResult.createDuplicate());
		}
	}

	CGeoDataEntityQueryResults.prototype.setGeoDataEntityQueryResult = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoDataEntityQueryResults_SetGeoDataEntityQueryResult, this.geoDataEntityQueryResult, pr));
		this.geoDataEntityQueryResult = pr;
		this.setParentToChild(pr);
	};


	// GeoDataPointQuery
	drawingsChangesMap[AscDFH.historyitem_GeoDataPointQuery_SetEntityType] = function (oClass, value) {
		oClass.entityType = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoDataPointQuery_SetLatitude] = function (oClass, value) {
		oClass.latitude = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoDataPointQuery_SetLongitude] = function (oClass, value) {
		oClass.longitude = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoDataPointQuery_SetEntityType] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_GeoDataPointQuery_SetLatitude] = window['AscDFH'].CChangesDrawingsDouble2;
	AscDFH.changesFactory[AscDFH.historyitem_GeoDataPointQuery_SetLongitude] = window['AscDFH'].CChangesDrawingsDouble2;

	function CGeoDataPointQuery() {
		CBaseChartObject.call(this);
		this.entityType = null;
		this.latitude = null;
		this.longitude = null;
	}

	InitClass(CGeoDataPointQuery, CBaseChartObject, AscDFH.historyitem_type_GeoDataPointQuery);

	CGeoDataPointQuery.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setEntityType(this.entityType);
		oCopy.setLatitude(this.latitude);
		oCopy.setLongitude(this.longitude);
	}

	CGeoDataPointQuery.prototype.setEntityType = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoDataPointQuery_SetEntityType, this.entityType, pr));
		this.entityType = pr;
		this.setParentToChild(pr);
	};
	CGeoDataPointQuery.prototype.setLatitude = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_GeoDataPointQuery_SetLatitude, this.latitude, pr));
		this.latitude = pr;
		this.setParentToChild(pr);
	};
	CGeoDataPointQuery.prototype.setLongitude = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_GeoDataPointQuery_SetLongitude, this.longitude, pr));
		this.longitude = pr;
		this.setParentToChild(pr);
	};


	// GeoDataPointToEntityQuery
	drawingsChangesMap[AscDFH.historyitem_GeoDataPointToEntityQuery_SetEntityType] = function (oClass, value) {
		oClass.entityType = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoDataPointToEntityQuery_SetEntityId] = function (oClass, value) {
		oClass.entityId = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoDataPointToEntityQuery_SetEntityType] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_GeoDataPointToEntityQuery_SetEntityId] = window['AscDFH'].CChangesDrawingsString;

	function CGeoDataPointToEntityQuery() {
		CBaseChartObject.call(this);
		this.entityType = null;
		this.entityId = null;
	}

	InitClass(CGeoDataPointToEntityQuery, CBaseChartObject, AscDFH.historyitem_type_GeoDataPointToEntityQuery);

	CGeoDataPointToEntityQuery.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setEntityType(this.entityType);
		oCopy.setEntityId(this.entityId);
	}

	CGeoDataPointToEntityQuery.prototype.setEntityType = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoDataPointToEntityQuery_SetEntityType, this.entityType, pr));
		this.entityType = pr;
		this.setParentToChild(pr);
	};
	CGeoDataPointToEntityQuery.prototype.setEntityId = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoDataPointToEntityQuery_SetEntityId, this.entityId, pr));
		this.entityId = pr;
		this.setParentToChild(pr);
	};


	// GeoDataPointToEntityQueryResult
	drawingsChangesMap[AscDFH.historyitem_GeoDataPointToEntityQueryResult_SetGeoDataPointQuery] = function (oClass, value) {
		oClass.geoDataPointQuery = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoDataPointToEntityQueryResult_SetGeoDataPointToEntityQuery] = function (oClass, value) {
		oClass.geoDataPointToEntityQuery = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoDataPointToEntityQueryResult_SetGeoDataPointQuery] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_GeoDataPointToEntityQueryResult_SetGeoDataPointToEntityQuery] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoDataPointToEntityQueryResult() {
		CBaseChartObject.call(this);
		this.geoDataPointQuery = null;
		this.geoDataPointToEntityQuery = null;
	}

	InitClass(CGeoDataPointToEntityQueryResult, CBaseChartObject, AscDFH.historyitem_type_GeoDataPointToEntityQueryResult);

	CGeoDataPointToEntityQueryResult.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoDataPointQuery) {
			oCopy.setGeoDataPointQuery(this.geoDataPointQuery.createDuplicate());
		}
		if (this.geoDataPointToEntityQuery) {
			oCopy.setGeoDataPointToEntityQuery(this.geoDataPointToEntityQuery.createDuplicate());
		}
	}
	
	CGeoDataPointToEntityQueryResult.prototype.setGeoDataPointQuery = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoDataPointToEntityQueryResult_SetGeoDataPointQuery, this.geoDataPointQuery, pr));
		this.geoDataPointQuery = pr;
		this.setParentToChild(pr);
	};
	CGeoDataPointToEntityQueryResult.prototype.setGeoDataPointToEntityQuery = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoDataPointToEntityQueryResult_SetGeoDataPointToEntityQuery, this.geoDataPointToEntityQuery, pr));
		this.geoDataPointToEntityQuery = pr;
		this.setParentToChild(pr);
	};


	// GeoDataPointToEntityQueryResults
	drawingsChangesMap[AscDFH.historyitem_GeoDataPointToEntityQueryResults_SetGeoDataPointToEntityQueryResult] = function (oClass, value) {
		oClass.geoDataPointToEntityQueryResult = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoDataPointToEntityQueryResults_SetGeoDataPointToEntityQueryResult] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoDataPointToEntityQueryResults() {
		CBaseChartObject.call(this);
		this.geoDataPointToEntityQueryResult = null;
	}

	InitClass(CGeoDataPointToEntityQueryResults, CBaseChartObject, AscDFH.historyitem_type_GeoDataPointToEntityQueryResults);

	CGeoDataPointToEntityQueryResults.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoDataPointToEntityQueryResult) {
			oCopy.setGeoDataPointToEntityQueryResult(this.geoDataPointToEntityQueryResult.createDuplicate());
		}
	}
	
	CGeoDataPointToEntityQueryResults.prototype.setGeoDataPointToEntityQueryResult = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoDataPointToEntityQueryResults_SetGeoDataPointToEntityQueryResult, this.geoDataPointToEntityQueryResult, pr));
		this.geoDataPointToEntityQueryResult = pr;
		this.setParentToChild(pr);
	};


	// Geography
	drawingsChangesMap[AscDFH.historyitem_Geography_SetGeoCache] = function (oClass, value) {
		oClass.geoCache = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Geography_SetProjectionType] = function (oClass, value) {
		oClass.projectionType = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Geography_SetViewedRegionType] = function (oClass, value) {
		oClass.viewedRegionType = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Geography_SetCultureLanguage] = function (oClass, value) {
		oClass.cultureLanguage = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Geography_SetCultureRegion] = function (oClass, value) {
		oClass.cultureRegion = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Geography_SetAttribution] = function (oClass, value) {
		oClass.attribution = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_Geography_SetGeoCache] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Geography_SetProjectionType] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Geography_SetViewedRegionType] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Geography_SetCultureLanguage] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_Geography_SetCultureRegion] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_Geography_SetAttribution] = window['AscDFH'].CChangesDrawingsString;

	function CGeography() {
		CBaseChartObject.call(this);
		this.geoCache = null;
		this.projectionType = null;
		this.viewedRegionType = null;
		this.cultureLanguage = null;
		this.cultureRegion = null;
		this.attribution = null;
	}

	InitClass(CGeography, CBaseChartObject, AscDFH.historyitem_type_Geography);

	CGeography.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoCache) {
			oCopy.setGeoCache(this.geoCache.createDuplicate());
		}
		oCopy.setProjectionType(this.projectionType);
		oCopy.setViewedRegionType(this.viewedRegionType);
		oCopy.setCultureLanguage(this.cultureLanguage);
		oCopy.setCultureRegion(this.cultureRegion);
		oCopy.setAttribution(this.attribution);
	}

	CGeography.prototype.setGeoCache = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Geography_SetGeoCache, this.geoCache, pr));
		this.geoCache = pr;
		this.setParentToChild(pr);
	};
	CGeography.prototype.setProjectionType = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Geography_SetProjectionType, this.projectionType, pr));
		this.projectionType = pr;
		this.setParentToChild(pr);
	};
	CGeography.prototype.setViewedRegionType = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Geography_SetViewedRegionType, this.viewedRegionType, pr));
		this.viewedRegionType = pr;
		this.setParentToChild(pr);
	};
	CGeography.prototype.setCultureLanguage = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Geography_SetCultureLanguage, this.cultureLanguage, pr));
		this.cultureLanguage = pr;
		this.setParentToChild(pr);
	};
	CGeography.prototype.setCultureRegion = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Geography_SetCultureRegion, this.cultureRegion, pr));
		this.cultureRegion = pr;
		this.setParentToChild(pr);
	};
	CGeography.prototype.setAttribution = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Geography_SetAttribution, this.attribution, pr));
		this.attribution = pr;
		this.setParentToChild(pr);
	};


	// GeoHierarchyEntity
	drawingsChangesMap[AscDFH.historyitem_GeoHierarchyEntity_SetEntityName] = function (oClass, value) {
		oClass.entityName = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoHierarchyEntity_SetEntityId] = function (oClass, value) {
		oClass.entityId = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoHierarchyEntity_SetEntityType] = function (oClass, value) {
		oClass.entityType = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoHierarchyEntity_SetEntityName] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_GeoHierarchyEntity_SetEntityId] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_GeoHierarchyEntity_SetEntityType] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoHierarchyEntity() {
		CBaseChartObject.call(this);
		this.entityName = null;
		this.entityId = null;
		this.entityType = null;
	}

	InitClass(CGeoHierarchyEntity, CBaseChartObject, AscDFH.historyitem_type_GeoHierarchyEntity);

	CGeoHierarchyEntity.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setEntityName(this.entityName);
		oCopy.setEntityId(this.entityId);
		oCopy.setEntityType(this.entityType);
	}

	CGeoHierarchyEntity.prototype.setEntityName = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoHierarchyEntity_SetEntityName, this.entityName, pr));
		this.entityName = pr;
		this.setParentToChild(pr);
	};
	CGeoHierarchyEntity.prototype.setEntityId = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoHierarchyEntity_SetEntityId, this.entityId, pr));
		this.entityId = pr;
		this.setParentToChild(pr);
	};
	CGeoHierarchyEntity.prototype.setEntityType = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoHierarchyEntity_SetEntityType, this.entityType, pr));
		this.entityType = pr;
		this.setParentToChild(pr);
	};


	// GeoLocation
	drawingsChangesMap[AscDFH.historyitem_GeoLocation_SetAddress] = function (oClass, value) {
		oClass.address = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoLocation_SetLatitude] = function (oClass, value) {
		oClass.latitude = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoLocation_SetLongitude] = function (oClass, value) {
		oClass.longitude = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoLocation_SetEntityName] = function (oClass, value) {
		oClass.entityName = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoLocation_SetEntityType] = function (oClass, value) {
		oClass.entityType = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocation_SetAddress] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocation_SetLatitude] = window['AscDFH'].CChangesDrawingsDouble2;
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocation_SetLongitude] = window['AscDFH'].CChangesDrawingsDouble2;
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocation_SetEntityName] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocation_SetEntityType] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoLocation() {
		CBaseChartObject.call(this);
		this.address = null;
		this.latitude = null;
		this.longitude = null;
		this.entityName = null;
		this.entityType = null;
	}

	InitClass(CGeoLocation, CBaseChartObject, AscDFH.historyitem_type_GeoLocation);

	CGeoLocation.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.address) {
			oCopy.setAddress(this.address.createDuplicate());
		}
		oCopy.setLatitude(this.latitude);
		oCopy.setLongitude(this.longitude);
		oCopy.setEntityName(this.entityName);
		oCopy.setEntityType(this.entityType);
	}

	CGeoLocation.prototype.setAddress = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoLocation_SetAddress, this.address, pr));
		this.address = pr;
		this.setParentToChild(pr);
	};
	CGeoLocation.prototype.setLatitude = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_GeoLocation_SetLatitude, this.latitude, pr));
		this.latitude = pr;
		this.setParentToChild(pr);
	};
	CGeoLocation.prototype.setLongitude = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_GeoLocation_SetLongitude, this.longitude, pr));
		this.longitude = pr;
		this.setParentToChild(pr);
	};
	CGeoLocation.prototype.setEntityName = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoLocation_SetEntityName, this.entityName, pr));
		this.entityName = pr;
		this.setParentToChild(pr);
	};
	CGeoLocation.prototype.setEntityType = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoLocation_SetEntityType, this.entityType, pr));
		this.entityType = pr;
		this.setParentToChild(pr);
	};


	// GeoLocationQuery
	drawingsChangesMap[AscDFH.historyitem_GeoLocationQuery_SetCountryRegion] = function (oClass, value) {
		oClass.countryRegion = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoLocationQuery_SetAdminDistrict1] = function (oClass, value) {
		oClass.adminDistrict1 = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoLocationQuery_SetAdminDistrict2] = function (oClass, value) {
		oClass.adminDistrict2 = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoLocationQuery_SetPostalCode] = function (oClass, value) {
		oClass.postalCode = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoLocationQuery_SetEntityType] = function (oClass, value) {
		oClass.entityType = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocationQuery_SetCountryRegion] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocationQuery_SetAdminDistrict1] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocationQuery_SetAdminDistrict2] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocationQuery_SetPostalCode] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocationQuery_SetEntityType] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoLocationQuery() {
		CBaseChartObject.call(this);
		this.countryRegion = null;
		this.adminDistrict1 = null;
		this.adminDistrict2 = null;
		this.postalCode = null;
		this.entityType = null;
	}

	InitClass(CGeoLocationQuery, CBaseChartObject, AscDFH.historyitem_type_GeoLocationQuery);

	CGeoLocationQuery.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setCountryRegion(this.countryRegion);
		oCopy.setAdminDistrict1(this.adminDistrict1);
		oCopy.setAdminDistrict2(this.adminDistrict2);
		oCopy.setPostalCode(this.postalCode);
		oCopy.setEntityType(this.entityType);
	}

	CGeoLocationQuery.prototype.setCountryRegion = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoLocationQuery_SetCountryRegion, this.countryRegion, pr));
		this.countryRegion = pr;
		this.setParentToChild(pr);
	};
	CGeoLocationQuery.prototype.setAdminDistrict1 = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoLocationQuery_SetAdminDistrict1, this.adminDistrict1, pr));
		this.adminDistrict1 = pr;
		this.setParentToChild(pr);
	};
	CGeoLocationQuery.prototype.setAdminDistrict2 = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoLocationQuery_SetAdminDistrict2, this.adminDistrict2, pr));
		this.adminDistrict2 = pr;
		this.setParentToChild(pr);
	};
	CGeoLocationQuery.prototype.setPostalCode = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoLocationQuery_SetPostalCode, this.postalCode, pr));
		this.postalCode = pr;
		this.setParentToChild(pr);
	};
	CGeoLocationQuery.prototype.setEntityType = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoLocationQuery_SetEntityType, this.entityType, pr));
		this.entityType = pr;
		this.setParentToChild(pr);
	};


	// GeoLocationQueryResult
	drawingsChangesMap[AscDFH.historyitem_GeoLocationQueryResult_SetGeoLocationQuery] = function (oClass, value) {
		oClass.geoLocationQuery = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoLocationQueryResult_SetGeoLocations] = function (oClass, value) {
		oClass.geoLocations = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocationQueryResult_SetGeoLocationQuery] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocationQueryResult_SetGeoLocations] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoLocationQueryResult() {
		CBaseChartObject.call(this);
		this.geoLocationQuery = null;
		this.geoLocations = null;
	}

	InitClass(CGeoLocationQueryResult, CBaseChartObject, AscDFH.historyitem_type_GeoLocationQueryResult);

	CGeoLocationQueryResult.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoLocationQuery) {
			oCopy.setGeoLocationQuery(this.geoLocationQuery.createDuplicate());
		}
		if (this.geoLocations) {
			oCopy.setGeoLocations(this.geoLocations.createDuplicate());
		}
	}

	CGeoLocationQueryResult.prototype.setGeoLocationQuery = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoLocationQueryResult_SetGeoLocationQuery, this.geoLocationQuery, pr));
		this.geoLocationQuery = pr;
		this.setParentToChild(pr);
	};
	CGeoLocationQueryResult.prototype.setGeoLocations = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoLocationQueryResult_SetGeoLocations, this.geoLocations, pr));
		this.geoLocations = pr;
		this.setParentToChild(pr);
	};


	// GeoLocationQueryResults
	drawingsChangesMap[AscDFH.historyitem_GeoLocationQueryResults_SetGeoLocationQueryResult] = function (oClass, value) {
		oClass.geoLocationQueryResult = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocationQueryResults_SetGeoLocationQueryResult] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoLocationQueryResults() {
		CBaseChartObject.call(this);
		this.geoLocationQueryResult = null;
	}

	InitClass(CGeoLocationQueryResults, CBaseChartObject, AscDFH.historyitem_type_GeoLocationQueryResults);

	CGeoLocationQueryResults.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoLocationQueryResult) {
			oCopy.setGeoLocationQueryResult(this.geoLocationQueryResult.createDuplicate());
		}
	}

	CGeoLocationQueryResults.prototype.setGeoLocationQueryResult = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoLocationQueryResults_SetGeoLocationQueryResult, this.geoLocationQueryResult, pr));
		this.geoLocationQueryResult = pr;
		this.setParentToChild(pr);
	};


	// GeoLocations
	drawingsChangesMap[AscDFH.historyitem_GeoLocations_SetGeoLocation] = function (oClass, value) {
		oClass.geoLocation = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoLocations_SetGeoLocation] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoLocations() {
		CBaseChartObject.call(this);
		this.geoLocation = null;
	}

	InitClass(CGeoLocations, CBaseChartObject, AscDFH.historyitem_type_GeoLocations);

	CGeoLocations.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoLocation) {
			oCopy.setGeoLocation(this.geoLocation.createDuplicate());
		}
	}

	CGeoLocations.prototype.setGeoLocation = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoLocations_SetGeoLocation, this.geoLocation, pr));
		this.geoLocation = pr;
		this.setParentToChild(pr);
	};


	// GeoPolygon
	drawingsChangesMap[AscDFH.historyitem_GeoPolygon_SetPolygonId] = function (oClass, value) {
		oClass.polygonId = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoPolygon_SetNumPoints] = function (oClass, value) {
		oClass.numPoints = value;
	};
	drawingsChangesMap[AscDFH.historyitem_GeoPolygon_SetPcaRings] = function (oClass, value) {
		oClass.pcaRings = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoPolygon_SetPolygonId] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_GeoPolygon_SetNumPoints] = window['AscDFH'].CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_GeoPolygon_SetPcaRings] = window['AscDFH'].CChangesDrawingsString;

	function CGeoPolygon() {
		CBaseChartObject.call(this);
		this.polygonId = null;
		this.numPoints = null;
		this.pcaRings = null;
	}

	InitClass(CGeoPolygon, CBaseChartObject, AscDFH.historyitem_type_GeoPolygon);

	CGeoPolygon.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setPolygonId(this.polygonId);
		oCopy.setNumPoints(this.numPoints);
		oCopy.setPcaRings(this.pcaRings);
	}


	CGeoPolygon.prototype.setPolygonId = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoPolygon_SetPolygonId, this.polygonId, pr));
		this.polygonId = pr;
		this.setParentToChild(pr);
	};
	CGeoPolygon.prototype.setNumPoints = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_GeoPolygon_SetNumPoints, this.numPoints, pr));
		this.numPoints = pr;
		this.setParentToChild(pr);
	};
	CGeoPolygon.prototype.setPcaRings = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_GeoPolygon_SetPcaRings, this.pcaRings, pr));
		this.pcaRings = pr;
		this.setParentToChild(pr);
	};


	// GeoPolygons
	drawingsChangesMap[AscDFH.historyitem_GeoPolygons_SetGeoPolygon] = function (oClass, value) {
		oClass.geoPolygon = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_GeoPolygons_SetGeoPolygon] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CGeoPolygons() {
		CBaseChartObject.call(this);
		this.geoPolygon = null;
	}

	InitClass(CGeoPolygons, CBaseChartObject, AscDFH.historyitem_type_GeoPolygons);

	CGeoPolygons.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.geoPolygon) {
			oCopy.setGeoPolygon(this.geoPolygon.createDuplicate());
		}
	}

	CGeoPolygons.prototype.setGeoPolygon = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_GeoPolygons_SetGeoPolygon, this.geoPolygon, pr));
		this.geoPolygon = pr;
		this.setParentToChild(pr);
	};


	// Gridlines
	drawingsChangesMap[AscDFH.historyitem_Gridlines_SetSpPr] = function (oClass, value) {
		oClass.spPr = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Gridlines_SetName] = function (oClass, value) {
		oClass.name = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_Gridlines_SetSpPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Gridlines_SetName] = window['AscDFH'].CChangesDrawingsString;

	function CGridlines() {
		CBaseChartObject.call(this);
		this.spPr = null;
		this.name = null;
	}

	InitClass(CGridlines, CBaseChartObject, AscDFH.historyitem_type_Gridlines);

	CGridlines.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.spPr) {
			oCopy.setSpPr(this.spPr.createDuplicate());
		}
		oCopy.setName(this.name);
	}

	CGridlines.prototype.setSpPr = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Gridlines_SetSpPr, this.spPr, pr));
		this.spPr = pr;
		this.setParentToChild(pr);
	};
	CGridlines.prototype.setName = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Gridlines_SetName, this.name, pr));
		this.name = pr;
		this.setParentToChild(pr);
	};


	// // HeaderFooter (contains in ChartFormat.js)
	// drawingsChangesMap[AscDFH.historyitem_HeaderFooter_SetOddHeader] = function (oClass, value) {
	//     oClass.oddHeader = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_HeaderFooter_SetOddFooter] = function (oClass, value) {
	//     oClass.oddFooter = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_HeaderFooter_SetEvenHeader] = function (oClass, value) {
	//     oClass.evenHeader = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_HeaderFooter_SetEvenFooter] = function (oClass, value) {
	//     oClass.evenFooter = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_HeaderFooter_SetFirstHeader] = function (oClass, value) {
	//     oClass.firstHeader = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_HeaderFooter_SetFirstFooter] = function (oClass, value) {
	//     oClass.firstFooter = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_HeaderFooter_SetAlignWithMargins] = function (oClass, value) {
	//     oClass.alignWithMargins = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_HeaderFooter_SetDifferentOddEven] = function (oClass, value) {
	//     oClass.differentOddEven = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_HeaderFooter_SetDifferentFirst] = function (oClass, value) {
	//     oClass.differentFirst = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_HeaderFooter_SetOddHeader] = window['AscDFH'].CChangesDrawingsString;
	// AscDFH.changesFactory[AscDFH.historyitem_HeaderFooter_SetOddFooter] = window['AscDFH'].CChangesDrawingsString;
	// AscDFH.changesFactory[AscDFH.historyitem_HeaderFooter_SetEvenHeader] = window['AscDFH'].CChangesDrawingsString;
	// AscDFH.changesFactory[AscDFH.historyitem_HeaderFooter_SetEvenFooter] = window['AscDFH'].CChangesDrawingsString;
	// AscDFH.changesFactory[AscDFH.historyitem_HeaderFooter_SetFirstHeader] = window['AscDFH'].CChangesDrawingsString;
	// AscDFH.changesFactory[AscDFH.historyitem_HeaderFooter_SetFirstFooter] = window['AscDFH'].CChangesDrawingsString;
	// AscDFH.changesFactory[AscDFH.historyitem_HeaderFooter_SetAlignWithMargins] = window['AscDFH'].CChangesDrawingsBool;
	// AscDFH.changesFactory[AscDFH.historyitem_HeaderFooter_SetDifferentOddEven] = window['AscDFH'].CChangesDrawingsBool;
	// AscDFH.changesFactory[AscDFH.historyitem_HeaderFooter_SetDifferentFirst] = window['AscDFH'].CChangesDrawingsBool;
	// function CHeaderFooter() {
	//     CBaseChartObject.call(this);
	//     this.oddHeader = null;
	//     this.oddFooter = null;
	//     this.evenHeader = null;
	//     this.evenFooter = null;
	//     this.firstHeader = null;
	//     this.firstFooter = null;
	//     this.alignWithMargins = null;
	//     this.differentOddEven = null;
	//     this.differentFirst = null;
	// }

	// InitClass(CHeaderFooter, CBaseChartObject, AscDFH.historyitem_type_HeaderFooter);

	// CHeaderFooter.prototype.setOddHeader = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_HeaderFooter_SetOddHeader, this.oddHeader, pr));
	//     this.oddHeader = pr;
	// };
	// CHeaderFooter.prototype.setOddFooter = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_HeaderFooter_SetOddFooter, this.oddFooter, pr));
	//     this.oddFooter = pr;
	// };
	// CHeaderFooter.prototype.setEvenHeader = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_HeaderFooter_SetEvenHeader, this.evenHeader, pr));
	//     this.evenHeader = pr;
	// };
	// CHeaderFooter.prototype.setEvenFooter = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_HeaderFooter_SetEvenFooter, this.evenFooter, pr));
	//     this.evenFooter = pr;
	// };
	// CHeaderFooter.prototype.setFirstHeader = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_HeaderFooter_SetFirstHeader, this.firstHeader, pr));
	//     this.firstHeader = pr;
	// };
	// CHeaderFooter.prototype.setFirstFooter = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_HeaderFooter_SetFirstFooter, this.firstFooter, pr));
	//     this.firstFooter = pr;
	// };
	// CHeaderFooter.prototype.setAlignWithMargins = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_HeaderFooter_SetAlignWithMargins, this.alignWithMargins, pr));
	//     this.alignWithMargins = pr;
	// };
	// CHeaderFooter.prototype.setDifferentOddEven = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_HeaderFooter_SetDifferentOddEven, this.differentOddEven, pr));
	//     this.differentOddEven = pr;
	// };
	// CHeaderFooter.prototype.setDifferentFirst = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_HeaderFooter_SetDifferentFirst, this.differentFirst, pr));
	//     this.differentFirst = pr;
	// };
	//

	// // ChartExLegend (Legend contains in ChartFormat.js)
	// drawingsChangesMap[AscDFH.historyitem_ChartExLegend_SetSpPr] = function (oClass, value) {
	//     oClass.spPr = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_ChartExLegend_SetTxPr] = function (oClass, value) {
	//     oClass.txPr = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_ChartExLegend_SetPos] = function (oClass, value) {
	//     oClass.pos = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_ChartExLegend_SetAlign] = function (oClass, value) {
	//     oClass.align = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_ChartExLegend_SetOverlay] = function (oClass, value) {
	//     oClass.overlay = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_ChartExLegend_SetSpPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
	// AscDFH.changesFactory[AscDFH.historyitem_ChartExLegend_SetTxPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
	// AscDFH.changesFactory[AscDFH.historyitem_ChartExLegend_SetPos] = window['AscDFH'].CChangesDrawingsObjectNoId;
	// AscDFH.changesFactory[AscDFH.historyitem_ChartExLegend_SetAlign] = window['AscDFH'].CChangesDrawingsObjectNoId;
	// AscDFH.changesFactory[AscDFH.historyitem_ChartExLegend_SetOverlay] = window['AscDFH'].CChangesDrawingsBool;
	// function CChartExLegend() {
	//     CBaseChartObject.call(this);
	//     this.spPr = null;
	//     this.txPr = null;
	//     this.pos = null;
	//     this.align = null;
	//     this.overlay = null;
	// }

	// InitClass(CChartExLegend, CBaseChartObject, AscDFH.historyitem_type_ChartExLegend);

	// CChartExLegend.prototype.setSpPr = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartExLegend_SetSpPr, this.spPr, pr));
	//     this.spPr = pr;
	// };
	// CChartExLegend.prototype.setTxPr = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartExLegend_SetTxPr, this.txPr, pr));
	//     this.txPr = pr;
	// };
	// CChartExLegend.prototype.setPos = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartExLegend_SetPos, this.pos, pr));
	//     this.pos = pr;
	// };
	// CChartExLegend.prototype.setAlign = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartExLegend_SetAlign, this.align, pr));
	//     this.align = pr;
	// };
	// CChartExLegend.prototype.setOverlay = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_ChartExLegend_SetOverlay, this.overlay, pr));
	//     this.overlay = pr;
	// };


	// // NumberColorPosition (CNumericPoint replaces this class)
	// drawingsChangesMap[AscDFH.historyitem_NumberColorPosition_SetVal] = function (oClass, value) {
	//     oClass.val = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_NumberColorPosition_SetVal] = window['AscDFH'].CChangesDrawingsDouble2;
	// function CNumberColorPosition() {
	//     CBaseChartObject.call(this);
	//     this.val = null;
	// }

	// InitClass(CNumberColorPosition, CBaseChartObject, AscDFH.historyitem_type_NumberColorPosition);

	// CNumberColorPosition.prototype.setVal = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_NumberColorPosition_SetVal, this.val, pr));
	//     this.val = pr;
	// };
	//

	// // NumberFormat (CNumFmt contains in ChartFormat.js)
	// drawingsChangesMap[AscDFH.historyitem_NumberFormat_SetFormatCode] = function (oClass, value) {
	//     oClass.formatCode = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_NumberFormat_SetSourceLinked] = function (oClass, value) {
	//     oClass.sourceLinked = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_NumberFormat_SetFormatCode] = window['AscDFH'].CChangesDrawingsString;
	// AscDFH.changesFactory[AscDFH.historyitem_NumberFormat_SetSourceLinked] = window['AscDFH'].CChangesDrawingsBool;
	// function CNumberFormat() {
	//     CBaseChartObject.call(this);
	//     this.formatCode = null;
	//     this.sourceLinked = null;
	// }

	// InitClass(CNumberFormat, CBaseChartObject, AscDFH.historyitem_type_NumberFormat);

	// CNumberFormat.prototype.setFormatCode = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_NumberFormat_SetFormatCode, this.formatCode, pr));
	//     this.formatCode = pr;
	// };
	// CNumberFormat.prototype.setSourceLinked = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_NumberFormat_SetSourceLinked, this.sourceLinked, pr));
	//     this.sourceLinked = pr;
	// };
	//

	// Dimension
	drawingsChangesMap[AscDFH.historyitem_Dimension_SetF] = function (oClass, value) {
		oClass.f = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Dimension_SetNf] = function (oClass, value) {
		oClass.nf = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Dimension_SetType] = function (oClass, value) {
		oClass.type = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_Dimension_SetF] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Dimension_SetNf] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_Dimension_SetType] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CDimension() {
		CBaseChartObject.call(this);
		this.f = null;
		this.nf = null;
		this.type = null;
		this.levelData = [];
	}

	InitClass(CDimension, CBaseChartObject, AscDFH.historyitem_type_Dimension);

	CDimension.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.f) {
			oCopy.setF(this.f.createDuplicate());
		}
		if (this.nf) {
			oCopy.setNf(this.nf.createDuplicate());
		}
		oCopy.setType(this.type);
	}

	CDimension.prototype.setF = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Dimension_SetF, this.f, pr));
		this.f = pr;
		this.setParentToChild(pr);
	};
	CDimension.prototype.setNf = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Dimension_SetNf, this.nf, pr));
		this.nf = pr;
		this.setParentToChild(pr);
	};
	CDimension.prototype.setType = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Dimension_SetType, this.type, pr));
		this.type = pr;
		this.setParentToChild(pr);
	};
	CDimension.prototype.createLvl = function () {
		return null;
	};
	CDimension.prototype.fillCellVal = function (oCell, oLvl, nPtIdx) {
		return null;
	};
	CDimension.prototype.updateReferences = function (bDisplayEmptyCellsAs, bDisplayHidden) {


		if (!this.f) {
			return;
		}
		let sContent = this.f.content;

		let aParsedRef = AscFormat.fParseChartFormula(sContent);
		if (!Array.isArray(aParsedRef) || aParsedRef.length === 0) {
			return false;
		}
		let nPtCount = 0;
		if (aParsedRef.length > 0) {
			let nRows = 0, nRef, oRef, oBBox, nPtIdx, nCol, oWS, oCell, sVal, nCols = 0, nRow;
			let nLvl, oLvl;
			let bLvlsByRows;
			if (this.f.dir === AscFormat.FORMULA_DIRECTION_ROW) {
				bLvlsByRows = true;
			}
			else {
				bLvlsByRows = false;
			}

			if (bLvlsByRows) {
				for (nRef = 0; nRef < aParsedRef.length; ++nRef) {
					oRef = aParsedRef[nRef];
					oBBox = oRef.bbox;
					nPtCount += (oBBox.c2 - oBBox.c1 + 1);
					nRows = Math.max(nRows, oBBox.r2 - oBBox.r1 + 1);
				}
				for (nLvl = 0; nLvl < nRows; ++nLvl) {
					oLvl = this.createLvl();
					if (!oLvl) {
						return;
					}
					nPtIdx = 0;
					for (nRef = 0; nRef < aParsedRef.length; ++nRef) {
						oRef = aParsedRef[nRef];
						oBBox = oRef.bbox;
						oWS = oRef.worksheet;
						if (nLvl < (oBBox.r2 - oBBox.r1 + 1)) {
							for (nCol = oBBox.c1; nCol <= oBBox.c2; ++nCol) {
								oCell = oWS.getCell3(nLvl + oBBox.r1, nCol);
								this.fillCellVal(oCell, oLvl, nPtIdx);
								++nPtIdx;
							}
						}
						else {
							nPtIdx += (oBBox.c2 - oBBox.c1 + 1);
						}
					}
					nPtCount = Math.max(nPtCount, nPtIdx);
					oLvl.setPtCount(nPtIdx);
					this.addLevelData(oLvl);
				}
			}
			else {
				for (nRef = 0; nRef < aParsedRef.length; ++nRef) {
					oRef = aParsedRef[nRef];
					oBBox = oRef.bbox;
					nPtCount += (oBBox.r2 - oBBox.r1 + 1);
					nCols = Math.max(nCols, oBBox.c2 - oBBox.c1 + 1);
				}
				for (nLvl = 0; nLvl < nCols; ++nLvl) {
					oLvl = this.createLvl();
					nPtIdx = 0;
					for (nRef = 0; nRef < aParsedRef.length; ++nRef) {
						oRef = aParsedRef[nRef];
						oBBox = oRef.bbox;
						oWS = oRef.worksheet;
						if (nLvl < (oBBox.c2 - oBBox.c1 + 1)) {
							for (nRow = oBBox.r1; nRow <= oBBox.r2; ++nRow) {
								oCell = oWS.getCell3(nRow, nLvl + oBBox.c1);
								this.fillCellVal(oCell, oLvl, nPtIdx);
								++nPtIdx;
							}
						}
						else {
							nPtIdx += (oBBox.r2 - oBBox.r1 + 1);
						}
					}
					nPtCount = Math.max(nPtCount, nPtIdx);
					oLvl.setPtCount(nPtIdx);
					this.addLevelData(oLvl);
				}
			}
		}
	};

	// NumericDimension
	drawingsChangesMap[AscDFH.historyitem_NumericDimension_SetLevelData] = function (oClass, value) {
		oClass.levelData = value;
	};

	drawingContentChanges[AscDFH.historyitem_NumericDimension_AddLevelData] =
		drawingContentChanges[AscDFH.historyitem_NumericDimension_RemoveLevelData] = function (oClass) {
			return oClass.levelData;
		};

	AscDFH.changesFactory[AscDFH.historyitem_NumericDimension_SetLevelData] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_NumericDimension_AddLevelData] = window['AscDFH'].CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_NumericDimension_RemoveLevelData] = window['AscDFH'].CChangesDrawingsContent;

	function CNumericDimension() {
		CDimension.call(this);
	}

	InitClass(CNumericDimension, CDimension, AscDFH.historyitem_type_NumericDimension);

	CNumericDimension.prototype.fillObject = function (oCopy) {
		CDimension.prototype.fillObject.call(this, oCopy);
		if (this.levelData) {
			for (let i = 0; i < this.levelData.length; i++) {
				oCopy.addLevelData(this.levelData[i].createDuplicate(), i);
			}
		}
	}

	CNumericDimension.prototype.addLevelData = function (pr, idx) {
		let pos;
		if (AscFormat.isRealNumber(idx))
			pos = idx;
		else
			pos = this.levelData.length;
		History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_NumericDimension_AddLevelData, pos, [pr], true));
		this.levelData.splice(pos, 0, pr);
		this.setParentToChild(pr);
	};
	CNumericDimension.prototype.removeLevelDataByPos = function (pos) {
		if (this.levelData[pos]) {
			let levelData = this.levelData.splice(pos, 1)[0];
			History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_NumericDimension_RemoveLevelData, pos, [levelData], false));
		}
	};
	CNumericDimension.prototype.createLvl = function () {
		return new AscFormat.CNumLit();
	};
	CNumericDimension.prototype.fillCellVal = function (oCell, oLvl, nPtIdx) {
		let dVal = oCell.getNumberValue();
		if (!AscFormat.isRealNumber(dVal)) {
			let sVal = oCell.getValueForEdit();
			if ((typeof sVal === "string") && sVal.length > 0) {
				dVal = 0;
			}
		}
		if (AscFormat.isRealNumber(dVal)) {
			let oPt = new AscFormat.CNumericPoint();
			oPt.setIdx(nPtIdx);
			oPt.setVal(dVal);
			oLvl.addPt(oPt);
		}
	};



	// // PageMargins (CPageMarginsChart contains in ChartFormat.js)
	// drawingsChangesMap[AscDFH.historyitem_PageMargins_SetL] = function (oClass, value) {
	//     oClass.l = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PageMargins_SetR] = function (oClass, value) {
	//     oClass.r = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PageMargins_SetT] = function (oClass, value) {
	//     oClass.t = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PageMargins_SetB] = function (oClass, value) {
	//     oClass.b = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PageMargins_SetHeader] = function (oClass, value) {
	//     oClass.header = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PageMargins_SetFooter] = function (oClass, value) {
	//     oClass.footer = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_PageMargins_SetL] = window['AscDFH'].CChangesDrawingsDouble2;
	// AscDFH.changesFactory[AscDFH.historyitem_PageMargins_SetR] = window['AscDFH'].CChangesDrawingsDouble2;
	// AscDFH.changesFactory[AscDFH.historyitem_PageMargins_SetT] = window['AscDFH'].CChangesDrawingsDouble2;
	// AscDFH.changesFactory[AscDFH.historyitem_PageMargins_SetB] = window['AscDFH'].CChangesDrawingsDouble2;
	// AscDFH.changesFactory[AscDFH.historyitem_PageMargins_SetHeader] = window['AscDFH'].CChangesDrawingsDouble2;
	// AscDFH.changesFactory[AscDFH.historyitem_PageMargins_SetFooter] = window['AscDFH'].CChangesDrawingsDouble2;
	// function CPageMargins() {
	//     CBaseChartObject.call(this);
	//     this.l = null;
	//     this.r = null;
	//     this.t = null;
	//     this.b = null;
	//     this.header = null;
	//     this.footer = null;
	// }

	// InitClass(CPageMargins, CBaseChartObject, AscDFH.historyitem_type_PageMargins);

	// CPageMargins.prototype.setL = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_PageMargins_SetL, this.l, pr));
	//     this.l = pr;
	// };
	// CPageMargins.prototype.setR = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_PageMargins_SetR, this.r, pr));
	//     this.r = pr;
	// };
	// CPageMargins.prototype.setT = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_PageMargins_SetT, this.t, pr));
	//     this.t = pr;
	// };
	// CPageMargins.prototype.setB = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_PageMargins_SetB, this.b, pr));
	//     this.b = pr;
	// };
	// CPageMargins.prototype.setHeader = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_PageMargins_SetHeader, this.header, pr));
	//     this.header = pr;
	// };
	// CPageMargins.prototype.setFooter = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_PageMargins_SetFooter, this.footer, pr));
	//     this.footer = pr;
	// };
	//

	// // PageSetup (contains in ChartFormat.js)
	// drawingsChangesMap[AscDFH.historyitem_PageSetup_SetPaperSize] = function (oClass, value) {
	//     oClass.paperSize = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PageSetup_SetFirstPageNumber] = function (oClass, value) {
	//     oClass.firstPageNumber = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PageSetup_SetOrientation] = function (oClass, value) {
	//     oClass.orientation = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PageSetup_SetBlackAndWhite] = function (oClass, value) {
	//     oClass.blackAndWhite = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PageSetup_SetDraft] = function (oClass, value) {
	//     oClass.draft = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PageSetup_SetUseFirstPageNumber] = function (oClass, value) {
	//     oClass.useFirstPageNumber = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PageSetup_SetHorizontalDpi] = function (oClass, value) {
	//     oClass.horizontalDpi = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PageSetup_SetVerticalDpi] = function (oClass, value) {
	//     oClass.verticalDpi = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PageSetup_SetCopies] = function (oClass, value) {
	//     oClass.copies = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_PageSetup_SetPaperSize] = window['AscDFH'].CChangesDrawingsLong;
	// AscDFH.changesFactory[AscDFH.historyitem_PageSetup_SetFirstPageNumber] = window['AscDFH'].CChangesDrawingsLong;
	// AscDFH.changesFactory[AscDFH.historyitem_PageSetup_SetOrientation] = window['AscDFH'].CChangesDrawingsObjectNoId;
	// AscDFH.changesFactory[AscDFH.historyitem_PageSetup_SetBlackAndWhite] = window['AscDFH'].CChangesDrawingsBool;
	// AscDFH.changesFactory[AscDFH.historyitem_PageSetup_SetDraft] = window['AscDFH'].CChangesDrawingsBool;
	// AscDFH.changesFactory[AscDFH.historyitem_PageSetup_SetUseFirstPageNumber] = window['AscDFH'].CChangesDrawingsBool;
	// AscDFH.changesFactory[AscDFH.historyitem_PageSetup_SetHorizontalDpi] = window['AscDFH'].CChangesDrawingsLong;
	// AscDFH.changesFactory[AscDFH.historyitem_PageSetup_SetVerticalDpi] = window['AscDFH'].CChangesDrawingsLong;
	// AscDFH.changesFactory[AscDFH.historyitem_PageSetup_SetCopies] = window['AscDFH'].CChangesDrawingsLong;
	// function CPageSetup() {
	//     CBaseChartObject.call(this);
	//     this.paperSize = null;
	//     this.firstPageNumber = null;
	//     this.orientation = null;
	//     this.blackAndWhite = null;
	//     this.draft = null;
	//     this.useFirstPageNumber = null;
	//     this.horizontalDpi = null;
	//     this.verticalDpi = null;
	//     this.copies = null;
	// }

	// InitClass(CPageSetup, CBaseChartObject, AscDFH.historyitem_type_PageSetup);

	// CPageSetup.prototype.setPaperSize = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PageSetup_SetPaperSize, this.paperSize, pr));
	//     this.paperSize = pr;
	// };
	// CPageSetup.prototype.setFirstPageNumber = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PageSetup_SetFirstPageNumber, this.firstPageNumber, pr));
	//     this.firstPageNumber = pr;
	// };
	// CPageSetup.prototype.setOrientation = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_PageSetup_SetOrientation, this.orientation, pr));
	//     this.orientation = pr;
	// };
	// CPageSetup.prototype.setBlackAndWhite = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_PageSetup_SetBlackAndWhite, this.blackAndWhite, pr));
	//     this.blackAndWhite = pr;
	// };
	// CPageSetup.prototype.setDraft = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_PageSetup_SetDraft, this.draft, pr));
	//     this.draft = pr;
	// };
	// CPageSetup.prototype.setUseFirstPageNumber = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_PageSetup_SetUseFirstPageNumber, this.useFirstPageNumber, pr));
	//     this.useFirstPageNumber = pr;
	// };
	// CPageSetup.prototype.setHorizontalDpi = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PageSetup_SetHorizontalDpi, this.horizontalDpi, pr));
	//     this.horizontalDpi = pr;
	// };
	// CPageSetup.prototype.setVerticalDpi = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PageSetup_SetVerticalDpi, this.verticalDpi, pr));
	//     this.verticalDpi = pr;
	// };
	// CPageSetup.prototype.setCopies = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PageSetup_SetCopies, this.copies, pr));
	//     this.copies = pr;
	// };
	//

	// // ParentLabelLayout ( st instead ct_ParentLabelLayout )
	// drawingsChangesMap[AscDFH.historyitem_ParentLabelLayout_SetVal] = function (oClass, value) {
	//     oClass.val = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_ParentLabelLayout_SetVal] = window['AscDFH'].CChangesDrawingsObjectNoId;
	// function CParentLabelLayout() {
	//     CBaseChartObject.call(this);
	//     this.val = null;
	// }

	// InitClass(CParentLabelLayout, CBaseChartObject, AscDFH.historyitem_type_ParentLabelLayout);

	// CParentLabelLayout.prototype.setVal = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ParentLabelLayout_SetVal, this.val, pr));
	//     this.val = pr;
	// };


	// PercentageColorPosition
	drawingsChangesMap[AscDFH.historyitem_PercentageColorPosition_SetVal] = function (oClass, value) {
		oClass.val = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_PercentageColorPosition_SetVal] = window['AscDFH'].CChangesDrawingsDouble2;

	function CPercentageColorPosition() {
		CBaseChartObject.call(this);
		this.val = null;
	}

	InitClass(CPercentageColorPosition, CBaseChartObject, AscDFH.historyitem_type_PercentageColorPosition);

	CPercentageColorPosition.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setVal(this.val);
	}

	CPercentageColorPosition.prototype.setVal = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_PercentageColorPosition_SetVal, this.val, pr));
		this.val = pr;
		this.setParentToChild(pr);
	};


	// // PlotArea (contains in chartFormat.js but different fields)
	// drawingsChangesMap[AscDFH.historyitem_ChartExPlotArea_SetPlotAreaRegion] = function (oClass, value) {
	//     oClass.plotAreaRegion = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_ChartExPlotArea_SetAxis] = function (oClass, value) {
	//     oClass.axis = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_ChartExPlotArea_SetSpPr] = function (oClass, value) {
	//     oClass.spPr = value;
	// };
	// drawingContentChanges[AscDFH.historyitem_ChartExPlotArea_AddAxis] =
	//     drawingContentChanges[AscDFH.historyitem_ChartExPlotArea_RemoveAxis] = function (oClass) {
	//         return oClass.axis;
	//     };
	// AscDFH.changesFactory[AscDFH.historyitem_ChartExPlotArea_SetPlotAreaRegion] = window['AscDFH'].CChangesDrawingsObjectNoId;
	// AscDFH.changesFactory[AscDFH.historyitem_ChartExPlotArea_SetAxis] = window['AscDFH'].CChangesDrawingsObjectNoId;
	// AscDFH.changesFactory[AscDFH.historyitem_ChartExPlotArea_SetSpPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
	// AscDFH.changesFactory[AscDFH.historyitem_ChartExPlotArea_AddAxis] = window['AscDFH'].CChangesDrawingsContent;
	// AscDFH.changesFactory[AscDFH.historyitem_ChartExPlotArea_RemoveAxis] = window['AscDFH'].CChangesDrawingsContent;
	// function CChartExPlotArea() {
	//     CBaseChartObject.call(this);
	//     this.plotAreaRegion = null;
	//     this.axis = [];
	//     this.spPr = null;
	// }

	// InitClass(CChartExPlotArea, CBaseChartObject, AscDFH.historyitem_type_ChartExPlotArea);

	// CChartExPlotArea.prototype.setPlotAreaRegion = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartExPlotArea_SetPlotAreaRegion, this.plotAreaRegion, pr));
	//     this.plotAreaRegion = pr;
	// };
	// CChartExPlotArea.prototype.setAxis = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartExPlotArea_SetAxis, this.axis, pr));
	//     this.axis = pr;
	// };
	// CChartExPlotArea.prototype.setSpPr = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartExPlotArea_SetSpPr, this.spPr, pr));
	//     this.spPr = pr;
	// };
	// CChartExPlotArea.prototype.addAxis = function (pr, idx) {
	//     var pos;
	//     if (AscFormat.isRealNumber(idx))
	//         pos = idx;
	//     else
	//         pos = this.axis.length;
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_ChartExPlotArea_AddAxis, pos, [pr], true));
	//     this.axis.splice(pos, 0, pr);
	// };
	// CChartExPlotArea.prototype.removeAxisByPos = function (pos) {
	//     if (this.axis[pos]) {
	//         var axis = this.axis.splice(pos, 1)[0];
	//         History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_ChartExPlotArea_RemoveAxis, pos, [axis], false));
	//     }
	// };


	// PlotAreaRegion
	drawingsChangesMap[AscDFH.historyitem_PlotAreaRegion_SetPlotSurface] = function (oClass, value) {
		oClass.plotSurface = value;
	};
	drawingsChangesMap[AscDFH.historyitem_PlotAreaRegion_SetSeries] = function (oClass, value) {
		oClass.series = value;
	};

	drawingContentChanges[AscDFH.historyitem_PlotAreaRegion_AddSeries] =
		drawingContentChanges[AscDFH.historyitem_PlotAreaRegion_RemoveSeries] = function (oClass) {
			return oClass.series;
		};
	AscDFH.changesFactory[AscDFH.historyitem_PlotAreaRegion_SetPlotSurface] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_PlotAreaRegion_SetSeries] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_PlotAreaRegion_AddSeries] = window['AscDFH'].CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_PlotAreaRegion_RemoveSeries] = window['AscDFH'].CChangesDrawingsContent;

	function CPlotAreaRegion() {
		CBaseChartObject.call(this);
		this.plotSurface = null;
		this.series = [];
	}

	InitClass(CPlotAreaRegion, CBaseChartObject, AscDFH.historyitem_type_PlotAreaRegion);

	CPlotAreaRegion.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.plotSurface) {
			oCopy.setPlotSurface(this.plotSurface.createDuplicate());
		}
		if (this.series) {
			for (let i = 0; i < this.series.length; i++) {
				oCopy.addSeries(this.series[i].createDuplicate(), i);
			}
		}
	}

	CPlotAreaRegion.prototype.getMaxSeriesIdx = function () {
		return this.series.length;
	};
	CPlotAreaRegion.prototype.setPlotSurface = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_PlotAreaRegion_SetPlotSurface, this.plotSurface, pr));
		this.plotSurface = pr;
		this.setParentToChild(pr);
	};
	CPlotAreaRegion.prototype.addSeries = function (pr, idx) {
		let pos;
		if (AscFormat.isRealNumber(idx))
			pos = idx;
		else
			pos = this.series.length;
		History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_PlotAreaRegion_AddSeries, pos, [pr], true));
		this.series.splice(pos, 0, pr);
		this.setParentToChild(pr);
	};
	CPlotAreaRegion.prototype.removeSeriesByPos = function (pos) {
		if (this.series[pos]) {
			let series = this.series.splice(pos, 1)[0];
			History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_PlotAreaRegion_RemoveSeries, pos, [series], false));
		}
	};
	CPlotAreaRegion.prototype.updateReferences = function (bDisplayEmptyCellsAs, bDisplayHidden) {

	};
	CPlotAreaRegion.prototype.getAllSeries = function () {
		return [].concat(this.series);
	};

	// PlotSurface
	drawingsChangesMap[AscDFH.historyitem_PlotSurface_SetSpPr] = function (oClass, value) {
		oClass.spPr = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_PlotSurface_SetSpPr] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CPlotSurface() {
		CBaseChartObject.call(this);
		this.spPr = null;
	}

	InitClass(CPlotSurface, CBaseChartObject, AscDFH.historyitem_type_PlotSurface);

	CPlotSurface.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.spPr) {
			oCopy.setSpPr(this.spPr.createDuplicate());
		}
	}

	CPlotSurface.prototype.setSpPr = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_PlotSurface_SetSpPr, this.spPr, pr));
		this.spPr = pr;
		this.setParentToChild(pr);
	};


	// // PrintSettings (contains in ChartFormat.js)
	// drawingsChangesMap[AscDFH.historyitem_PrintSettings_SetHeaderFooter] = function (oClass, value) {
	//     oClass.headerFooter = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PrintSettings_SetPageMargins] = function (oClass, value) {
	//     oClass.pageMargins = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_PrintSettings_SetPageSetup] = function (oClass, value) {
	//     oClass.pageSetup = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_PrintSettings_SetHeaderFooter] = window['AscDFH'].CChangesDrawingsString;
	// AscDFH.changesFactory[AscDFH.historyitem_PrintSettings_SetPageMargins] = window['AscDFH'].CChangesDrawingsString;
	// AscDFH.changesFactory[AscDFH.historyitem_PrintSettings_SetPageSetup] = window['AscDFH'].CChangesDrawingsString;
	// function CPrintSettings() {
	//     CBaseChartObject.call(this);
	//     this.headerFooter = null;
	//     this.pageMargins = null;
	//     this.pageSetup = null;
	// }

	// InitClass(CPrintSettings, CBaseChartObject, AscDFH.historyitem_type_PrintSettings);

	// CPrintSettings.prototype.setHeaderFooter = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_PrintSettings_SetHeaderFooter, this.headerFooter, pr));
	//     this.headerFooter = pr;
	// };
	// CPrintSettings.prototype.setPageMargins = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_PrintSettings_SetPageMargins, this.pageMargins, pr));
	//     this.pageMargins = pr;
	// };
	// CPrintSettings.prototype.setPageSetup = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_PrintSettings_SetPageSetup, this.pageSetup, pr));
	//     this.pageSetup = pr;
	// };
	//

	// // RegionLabelLayout (st is used instead of CT_RegionLabelLayout)
	// drawingsChangesMap[AscDFH.historyitem_RegionLabelLayout_SetVal] = function (oClass, value) {
	//     oClass.val = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_RegionLabelLayout_SetVal] = window['AscDFH'].CChangesDrawingsString;
	// function CRegionLabelLayout() {
	//     CBaseChartObject.call(this);
	//     this.val = null;
	// }

	// InitClass(CRegionLabelLayout, CBaseChartObject, AscDFH.historyitem_type_RegionLabelLayout);

	// CRegionLabelLayout.prototype.setVal = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_RegionLabelLayout_SetVal, this.val, pr));
	//     this.val = pr;
	// };
	//

	// RelId
	drawingsChangesMap[AscDFH.historyitem_RelId_SetId] = function (oClass, value) {
		oClass.id = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_RelId_SetId] = window['AscDFH'].CChangesDrawingsObjectNoId;



	// Series

	drawingsChangesMap[AscDFH.historyitem_Series_SetDataLabels] = function (oClass, value) {
		oClass.dataLabels = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Series_SetDataId] = function (oClass, value) {
		oClass.dataId = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Series_SetLayoutPr] = function (oClass, value) {
		oClass.layoutPr = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Series_SetLayoutId] = function (oClass, value) {
		oClass.layoutId = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Series_SetHidden] = function (oClass, value) {
		oClass.hidden = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Series_SetOwnerIdx] = function (oClass, value) {
		oClass.ownerIdx = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Series_SetUniqueId] = function (oClass, value) {
		oClass.uniqueId = value;
	};
	drawingsChangesMap[AscDFH.historyitem_Series_SetFormatIdx] = function (oClass, value) {
		oClass.formatIdx = value;
	};

	drawingContentChanges[AscDFH.historyitem_Series_AddDataPt] =
		drawingContentChanges[AscDFH.historyitem_Series_RemoveDataPt] = function (oClass) {
			return oClass.dPt;
		};

	drawingContentChanges[AscDFH.historyitem_Series_AddAxisId] =
		drawingContentChanges[AscDFH.historyitem_Series_RemoveAxisId] = function (oClass) {
			return oClass.axisId;
		};

	AscDFH.changesFactory[AscDFH.historyitem_Series_SetDataLabels] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Series_SetDataId] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Series_SetLayoutPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Series_SetLayoutId] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Series_SetHidden] = window['AscDFH'].CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_Series_SetOwnerIdx] = window['AscDFH'].CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_Series_SetUniqueId] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_Series_SetFormatIdx] = window['AscDFH'].CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_Series_AddDataPt] = window['AscDFH'].CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_Series_RemoveDataPt] = window['AscDFH'].CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_Series_AddAxisId] = window['AscDFH'].CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_Series_RemoveAxisId] = window['AscDFH'].CChangesDrawingsContent;

	function CSeries() {
		AscFormat.CSeriesBase.call(this);
		// commented in ChartSerializeEx.h
		// this.valueColors = null;
		// this.valueColorPositions = null;
		this.dPt = [];
		this.dataLabels = null;
		this.dataId = null;
		this.layoutPr = null;
		this.axisId = [];
		this.layoutId = null;
		this.hidden = null;
		this.ownerIdx = null;
		this.uniqueId = null;
		this.formatIdx = null;
	}

	InitClass(CSeries, AscFormat.CSeriesBase, AscDFH.historyitem_type_Series);

	CSeries.prototype.fillObject = function (oCopy) {
		AscFormat.CSeriesBase.prototype.fillObject.call(this, oCopy);
		if (this.dataLabels) {
			oCopy.setDataLabels(this.dataLabels.createDuplicate());
		}
		if (this.dataId) {
			oCopy.setDataId(this.dataId.createDuplicate());
		}
		if (this.layoutPr) {
			oCopy.setLayoutPr(this.layoutPr.createDuplicate());
		}
		if (this.layoutId) {
			oCopy.setLayoutId(this.layoutId.createDuplicate());
		}
		if (this.dPt) {
			for (let i = 0; i < this.dPt.length; i++) {
				oCopy.addDataPt(this.dPt[i].createDuplicate(), i);
			}
		}
		if (this.axisId) {
			for (let i = 0; i < this.axisId.length; i++) {
				oCopy.addAxisId(this.axisId[i].createDuplicate(), i);
			}
		}
		oCopy.setHidden(this.hidden);
		oCopy.setOwnerIdx(this.ownerIdx);
		oCopy.setUniqueId(this.uniqueId);
		oCopy.setFormatIdx(this.formatIdx);
	}

	// CSeries.prototype.setValueColors = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Series_SetValueColors, this.valueColors, pr));
	//     this.valueColors = pr;
	// };
	// CSeries.prototype.setValueColorPositions = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Series_SetValueColorPositions, this.valueColorPositions, pr));
	//     this.valueColorPositions = pr;
	// };
	CSeries.prototype.addDataPt = function (pr, idx) {
		let pos;
		if (AscFormat.isRealNumber(idx))
			pos = idx;
		else
			pos = this.dPt.length;
		History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_Series_AddDataPt, pos, [pr], true));
		this.dPt.splice(pos, 0, pr);
		this.setParentToChild(pr);
	};
	CSeries.prototype.removeDataPtByPos = function (pos) {
		if (this.dPt[pos]) {
			let dPt = this.dPt.splice(pos, 1)[0];
			History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_Series_RemoveDataPt, pos, [dPt], false));
		}
	};
	CSeries.prototype.setDataLabels = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Series_SetDataLabels, this.dataLabels, pr));
		this.dataLabels = pr;
		this.setParentToChild(pr);
	};
	CSeries.prototype.setDataId = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Series_SetDataId, this.dataId, pr));
		this.dataId = pr;
		this.setParentToChild(pr);
	};
	CSeries.prototype.setLayoutPr = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Series_SetLayoutPr, this.layoutPr, pr));
		this.layoutPr = pr;
		this.setParentToChild(pr);
	};
	CSeries.prototype.addAxisId = function (pr, idx) {
		let pos;
		if (AscFormat.isRealNumber(idx))
			pos = idx;
		else
			pos = this.axisId.length;
		History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_Series_AddAxisId, pos, [pr], true));
		this.axisId.splice(pos, 0, pr);
		this.setParentToChild(pr);
	};
	CSeries.prototype.removeAxisIdByPos = function (pos) {
		if (this.axisId[pos]) {
			let axisId = this.axisId.splice(pos, 1)[0];
			History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_Series_RemoveAxisId, pos, [axisId], false));
		}
	};
	CSeries.prototype.setLayoutId = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Series_SetLayoutId, this.layoutId, pr));
		this.layoutId = pr;
		this.setParentToChild(pr);
	};
	CSeries.prototype.setHidden = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Series_SetHidden, this.hidden, pr));
		this.hidden = pr;
		this.setParentToChild(pr);
	};
	CSeries.prototype.setOwnerIdx = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Series_SetOwnerIdx, this.ownerIdx, pr));
		this.ownerIdx = pr;
		this.setParentToChild(pr);
	};
	CSeries.prototype.setUniqueId = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Series_SetUniqueId, this.uniqueId, pr));
		this.uniqueId = pr;
		this.setParentToChild(pr);
	};
	CSeries.prototype.setFormatIdx = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Series_SetFormatIdx, this.formatIdx, pr));
		this.formatIdx = pr;
		this.setParentToChild(pr);
	};
	CSeries.prototype.getData = function () {
		if (this.dataId === null) {
			return null;
		}
		let oChartSpace = this.getChartSpace();
		if (!oChartSpace) {
			return null;
		}
		let oChartData = oChartSpace.chartData;
		if (!oChartData) {
			return null;
		}
		let oSeriesData = oChartData.getData(this.dataId);
		if (!oSeriesData) {
			return null;
		}
		return oSeriesData;
	};
	CSeries.prototype.getValLit = function () {
		let oSeriesData = this.getData();
		if (!oSeriesData) {
			return null;
		}
		let aValDim = oSeriesData.getValDimensions();
		if (aValDim.length > 0) {
			let oDim = aValDim[0];
			if (oDim) {
				return oDim.levelData[0] || null;
			}
		}
		return null;
	};
	CSeries.prototype.getValPts = function () {
		const numLit = this.getValLit();
		return numLit ? numLit.pts : [];
	};
	CSeries.prototype.getCatLit = function (type) {
		let oSeriesData = this.getData();
		if (!oSeriesData) {
			return null;
		}
		let aCatDim = oSeriesData.getCatDimensions();
		if (aCatDim.length > 0) {
			let oDim = aCatDim[0];
			if (oDim) {
				let index = (type === AscFormat.SERIES_LAYOUT_WATERFALL || type === AscFormat.SERIES_LAYOUT_FUNNEL) ? oDim.levelData.length - 1 : 0;
				return oDim.levelData[index] || null;
			}
		}
		return null;
	};
	CSeries.prototype.getBrush = function () {
		return this.compiledSeriesBrush;
	};
	CSeries.prototype.getPen = function () {
		return this.compiledSeriesPen;
	};
	CSeries.prototype.getDptByIdx = function (idx) {
		for (let nDpt = 0; nDpt < this.dPt.length; ++nDpt) {
			if (this.dPt[nDpt].idx === idx) {
				return this.dPt[nDpt];
			}
		}
		return null;
	};
	CSeries.prototype.getPtPen = function (nIdx) {
		let oDpt = this.getDptByIdx(nIdx);
		if (oDpt && oDpt.pen) {
			return oDpt.pen;
		}
		return this.compiledSeriesPen;
	};

	CSeries.prototype.getPtBrush = function (nIdx) {
		let oDpt = this.getDptByIdx(nIdx);
		if (oDpt && oDpt.brush) {
			return oDpt.brush;
		}
		return this.compiledSeriesBrush;
	};

	CSeries.prototype.getIdx = function () {
		if (!this.parent) {
			return -1;
		}
		let aSeries = this.parent.series;
		for (let nSer = 0; nSer < aSeries.length; ++nSer) {
			if (this === aSeries[nSer]) {
				return nSer;
			}
		}
		return -1;
	};


	// SeriesElementVisibilities
	drawingsChangesMap[AscDFH.historyitem_SeriesElementVisibilities_SetConnectorLines] = function (oClass, value) {
		oClass.connectorLines = value;
	};
	drawingsChangesMap[AscDFH.historyitem_SeriesElementVisibilities_SetMeanLine] = function (oClass, value) {
		oClass.meanLine = value;
	};
	drawingsChangesMap[AscDFH.historyitem_SeriesElementVisibilities_SetMeanMarker] = function (oClass, value) {
		oClass.meanMarker = value;
	};
	drawingsChangesMap[AscDFH.historyitem_SeriesElementVisibilities_SetNonoutliers] = function (oClass, value) {
		oClass.nonoutliers = value;
	};
	drawingsChangesMap[AscDFH.historyitem_SeriesElementVisibilities_SetOutliers] = function (oClass, value) {
		oClass.outliers = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_SeriesElementVisibilities_SetConnectorLines] = window['AscDFH'].CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_SeriesElementVisibilities_SetMeanLine] = window['AscDFH'].CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_SeriesElementVisibilities_SetMeanMarker] = window['AscDFH'].CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_SeriesElementVisibilities_SetNonoutliers] = window['AscDFH'].CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_SeriesElementVisibilities_SetOutliers] = window['AscDFH'].CChangesDrawingsBool;

	function CSeriesElementVisibilities() {
		CBaseChartObject.call(this);
		this.connectorLines = null;
		this.meanLine = null;
		this.meanMarker = null;
		this.nonoutliers = null;
		this.outliers = null;
	}

	InitClass(CSeriesElementVisibilities, CBaseChartObject, AscDFH.historyitem_type_SeriesElementVisibilities);

	CSeriesElementVisibilities.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setConnectorLines(this.connectorLines);
		oCopy.setMeanLine(this.meanLine);
		oCopy.setMeanMarker(this.meanMarker);
		oCopy.setNonoutliers(this.nonoutliers);
		oCopy.setOutliers(this.outliers);
	}

	CSeriesElementVisibilities.prototype.setConnectorLines = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_SeriesElementVisibilities_SetConnectorLines, this.connectorLines, pr));
		this.connectorLines = pr;
		this.setParentToChild(pr);
	};
	CSeriesElementVisibilities.prototype.setMeanLine = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_SeriesElementVisibilities_SetMeanLine, this.meanLine, pr));
		this.meanLine = pr;
		this.setParentToChild(pr);
	};
	CSeriesElementVisibilities.prototype.setMeanMarker = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_SeriesElementVisibilities_SetMeanMarker, this.meanMarker, pr));
		this.meanMarker = pr;
		this.setParentToChild(pr);
	};
	CSeriesElementVisibilities.prototype.setNonoutliers = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_SeriesElementVisibilities_SetNonoutliers, this.nonoutliers, pr));
		this.nonoutliers = pr;
		this.setParentToChild(pr);
	};
	CSeriesElementVisibilities.prototype.setOutliers = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_SeriesElementVisibilities_SetOutliers, this.outliers, pr));
		this.outliers = pr;
		this.setParentToChild(pr);
	};


	// SeriesLayoutProperties
	drawingsChangesMap[AscDFH.historyitem_SeriesLayoutProperties_SetParentLabelLayout] = function (oClass, value) {
		oClass.parentLabelLayout = value;
	};
	drawingsChangesMap[AscDFH.historyitem_SeriesLayoutProperties_SetRegionLabelLayout] = function (oClass, value) {
		oClass.regionLabelLayout = value;
	};
	drawingsChangesMap[AscDFH.historyitem_SeriesLayoutProperties_SetVisibility] = function (oClass, value) {
		oClass.visibility = value;
	};
	drawingsChangesMap[AscDFH.historyitem_SeriesLayoutProperties_SetAggregation] = function (oClass, value) {
		oClass.aggregation = value;
	};
	drawingsChangesMap[AscDFH.historyitem_SeriesLayoutProperties_SetBinning] = function (oClass, value) {
		oClass.binning = value;
	};
	drawingsChangesMap[AscDFH.historyitem_SeriesLayoutProperties_SetGeography] = function (oClass, value) {
		oClass.geography = value;
	};
	drawingsChangesMap[AscDFH.historyitem_SeriesLayoutProperties_SetStatistics] = function (oClass, value) {
		oClass.statistics = value;
	};
	drawingsChangesMap[AscDFH.historyitem_SeriesLayoutProperties_SetSubtotals] = function (oClass, value) {
		oClass.subtotals = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_SeriesLayoutProperties_SetParentLabelLayout] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_SeriesLayoutProperties_SetRegionLabelLayout] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_SeriesLayoutProperties_SetVisibility] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_SeriesLayoutProperties_SetAggregation] = window['AscDFH'].CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_SeriesLayoutProperties_SetBinning] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_SeriesLayoutProperties_SetGeography] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_SeriesLayoutProperties_SetStatistics] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_SeriesLayoutProperties_SetSubtotals] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CSeriesLayoutProperties() {
		CBaseChartObject.call(this);
		this.parentLabelLayout = null;
		this.regionLabelLayout = null;
		this.visibility = null;
		this.aggregation = null;
		this.statistics = null;
		this.subtotals = null;
		this.binning = null;
		// todo in ChartSerializeEx.h
		this.geography = null;
	}

	InitClass(CSeriesLayoutProperties, CBaseChartObject, AscDFH.historyitem_type_SeriesLayoutProperties);

	CSeriesLayoutProperties.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.visibility) {
			oCopy.setVisibility(this.visibility.createDuplicate());
		}
		if (this.statistics) {
			oCopy.setStatistics(this.statistics.createDuplicate());
		}
		if (this.subtotals) {
			oCopy.setSubtotals(this.subtotals.createDuplicate());
		}
		if (this.binning) {
			oCopy.setBinning(this.binning.createDuplicate());
		}
		if (this.geography) {
			oCopy.setGeography(this.geography.createDuplicate());
		}
		oCopy.setParentLabelLayout(this.parentLabelLayout);
		oCopy.setRegionLabelLayout(this.regionLabelLayout);
		oCopy.setAggregation(this.aggregation);
	}

	CSeriesLayoutProperties.prototype.setParentLabelLayout = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SeriesLayoutProperties_SetParentLabelLayout, this.parentLabelLayout, pr));
		this.parentLabelLayout = pr;
		this.setParentToChild(pr);
	};
	CSeriesLayoutProperties.prototype.setRegionLabelLayout = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SeriesLayoutProperties_SetRegionLabelLayout, this.regionLabelLayout, pr));
		this.regionLabelLayout = pr;
		this.setParentToChild(pr);
	};
	CSeriesLayoutProperties.prototype.setVisibility = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SeriesLayoutProperties_SetVisibility, this.visibility, pr));
		this.visibility = pr;
		this.setParentToChild(pr);
	};
	CSeriesLayoutProperties.prototype.setAggregation = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_SeriesLayoutProperties_SetAggregation, this.aggregation, pr));
		this.aggregation = pr;
		this.setParentToChild(pr);
	};
	CSeriesLayoutProperties.prototype.setBinning = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SeriesLayoutProperties_SetBinning, this.binning, pr));
		this.binning = pr;
		this.setParentToChild(pr);
	};
	CSeriesLayoutProperties.prototype.setGeography = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SeriesLayoutProperties_SetGeography, this.geography, pr));
		this.geography = pr;
		this.setParentToChild(pr);
	};
	CSeriesLayoutProperties.prototype.setStatistics = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SeriesLayoutProperties_SetStatistics, this.statistics, pr));
		this.statistics = pr;
		this.setParentToChild(pr);
	};
	CSeriesLayoutProperties.prototype.setSubtotals = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SeriesLayoutProperties_SetSubtotals, this.subtotals, pr));
		this.subtotals = pr;
		this.setParentToChild(pr);
	};


	// Statistics
	drawingsChangesMap[AscDFH.historyitem_Statistics_SetQuartileMethod] = function (oClass, value) {
		oClass.quartileMethod = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_Statistics_SetQuartileMethod] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CStatistics() {
		CBaseChartObject.call(this);
		this.quartileMethod = null;
	}

	InitClass(CStatistics, CBaseChartObject, AscDFH.historyitem_type_Statistics);

	CStatistics.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setQuartileMethod(this.quartileMethod);
	}

	CStatistics.prototype.setQuartileMethod = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Statistics_SetQuartileMethod, this.quartileMethod, pr));
		this.quartileMethod = pr;
		this.setParentToChild(pr);
	};

	// StringDimension
	drawingsChangesMap[AscDFH.historyitem_StringDimension_SetLevelData] = function (oClass, value) {
		oClass.levelData = value;
	};

	drawingContentChanges[AscDFH.historyitem_StringDimension_AddLevelData] =
		drawingContentChanges[AscDFH.historyitem_StringDimension_RemoveLevelData] = function (oClass) {
			return oClass.levelData;
		};
	AscDFH.changesFactory[AscDFH.historyitem_StringDimension_SetLevelData] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_StringDimension_AddLevelData] = window['AscDFH'].CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_StringDimension_RemoveLevelData] = window['AscDFH'].CChangesDrawingsContent;

	function CStringDimension() {
		CDimension.call(this);
	}

	InitClass(CStringDimension, CDimension, AscDFH.historyitem_type_StringDimension);

	CStringDimension.prototype.fillObject = function (oCopy) {
		CDimension.prototype.fillObject.call(this, oCopy);
		if (this.levelData) {
			for (let i = 0; i < this.levelData.length; i++) {
				oCopy.addLevelData(this.levelData[i].createDuplicate(), i);
			}
		}
	}

	CStringDimension.prototype.addLevelData = function (pr, idx) {
		let pos;
		if (AscFormat.isRealNumber(idx))
			pos = idx;
		else
			pos = this.levelData.length;
		History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_StringDimension_AddLevelData, pos, [pr], true));
		this.levelData.splice(pos, 0, pr);
		this.setParentToChild(pr);
	};
	CStringDimension.prototype.removeLevelDataByPos = function (pos) {
		if (this.levelData[pos]) {
			let levelData = this.levelData.splice(pos, 1)[0];
			History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_StringDimension_RemoveLevelData, pos, [levelData], false));
		}
	};

	CStringDimension.prototype.createLvl = function () {
		return new AscFormat.CStrCache();
	};
	CStringDimension.prototype.fillCellVal = function (oCell, oLvl, nPtIdx) {
		let sVal = oCell.getValueWithFormat();
		if (typeof sVal === "string" && sVal.length > 0) {
			let oPt = new AscFormat.CStringPoint();
			oPt.setIdx(nPtIdx);
			oPt.setVal(sVal);
			oLvl.addPt(oPt);
		}
	};

	
	// Subtotals
	drawingsChangesMap[AscDFH.historyitem_Subtotals_SetIdx] = function (oClass, value) {
		oClass.idx = value;
	};

	drawingContentChanges[AscDFH.historyitem_Subtotals_AddIdx] =
		drawingContentChanges[AscDFH.historyitem_Subtotals_RemoveIdx] = function (oClass) {
			return oClass.idx;
		};
	AscDFH.changesFactory[AscDFH.historyitem_Subtotals_SetIdx] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Subtotals_AddIdx] = window['AscDFH'].CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_Subtotals_RemoveIdx] = window['AscDFH'].CChangesDrawingsContent;

	function CSubtotals() {
		CBaseChartObject.call(this);
		this.idx = [];
	}

	InitClass(CSubtotals, CBaseChartObject, AscDFH.historyitem_type_Subtotals);

	CSubtotals.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.idx) {
			for (let i = 0; i < this.idx.length; i++) {
				oCopy.addDataPt(this.idx[i], i);
			}
		}
	}

	CSubtotals.prototype.addIdx = function (pr, idx) {
		let pos;
		if (AscFormat.isRealNumber(idx))
			pos = idx;
		else
			pos = this.idx.length;
		History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_Subtotals_AddIdx, pos, [pr], true));
		this.idx.splice(pos, 0, pr);
		this.setParentToChild(pr);
	};
	CSubtotals.prototype.removeIdxByPos = function (pos) {
		if (this.idx[pos]) {
			let idx = this.idx.splice(pos, 1)[0];
			History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_Subtotals_RemoveIdx, pos, [idx], false));
		}
	};


	// // Text (CChartText instead CText)
	// drawingsChangesMap[AscDFH.historyitem_Text_SetTxData] = function (oClass, value) {
	//     oClass.txData = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_Text_SetRich] = function (oClass, value) {
	//     oClass.rich = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_Text_SetTxData] = window['AscDFH'].CChangesDrawingsObjectNoId;
	// AscDFH.changesFactory[AscDFH.historyitem_Text_SetRich] = window['AscDFH'].CChangesDrawingsObjectNoId;
	// function CText() {
	//     CBaseChartObject.call(this);
	//     this.txData = null;
	//     this.rich = null;
	// }

	// InitClass(CText, CBaseChartObject, AscDFH.historyitem_type_Text);

	// CText.prototype.setTxData = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Text_SetTxData, this.txData, pr));
	//     this.txData = pr;
	// };
	// CText.prototype.setRich = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Text_SetRich, this.rich, pr));
	//     this.rich = pr;
	// };


	// TextData
	drawingsChangesMap[AscDFH.historyitem_TextData_SetF] = function (oClass, value) {
		oClass.f = value;
	};
	drawingsChangesMap[AscDFH.historyitem_TextData_SetV] = function (oClass, value) {
		oClass.v = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_TextData_SetF] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_TextData_SetV] = window['AscDFH'].CChangesDrawingsString;

	function CTextData() {
		CBaseChartObject.call(this);
		this.f = null;
		this.v = null;
	}

	InitClass(CTextData, CBaseChartObject, AscDFH.historyitem_type_TextData);

	CTextData.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.f) {
			oCopy.setF(this.f.createDuplicate());
		}
		oCopy.setV(this.v);
	}

	CTextData.prototype.setF = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_TextData_SetF, this.f, pr));
		this.f = pr;
		this.setParentToChild(pr);
	};
	CTextData.prototype.setV = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_TextData_SetV, this.v, pr));
		this.v = pr;
		this.setParentToChild(pr);
	};


	// // TickLabels (unused, bool instead this class)
	// function CTickLabels() {
	//     CBaseChartObject.call(this);
	// }

	// InitClass(CTickLabels, CBaseChartObject, AscDFH.historyitem_type_TickLabels);
	//

	// TickMarks
	drawingsChangesMap[AscDFH.historyitem_TickMarks_SetType] = function (oClass, value) {
		oClass.type = value;
	};
	drawingsChangesMap[AscDFH.historyitem_TickMarks_SetName] = function (oClass, value) {
		oClass.name = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_TickMarks_SetType] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_TickMarks_SetName] = window['AscDFH'].CChangesDrawingsString;

	function CTickMarks() {
		CBaseChartObject.call(this);
		this.type = null;
		this.name = null;
	}

	InitClass(CTickMarks, CBaseChartObject, AscDFH.historyitem_type_TickMarks);

	CTickMarks.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setType(this.type);
		oCopy.setName(this.name);
	}

	CTickMarks.prototype.setType = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_TickMarks_SetType, this.type, pr));
		this.type = pr;
		this.setParentToChild(pr);
	};
	CTickMarks.prototype.setName = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_TickMarks_SetName, this.name, pr));
		this.name = pr;
		this.setParentToChild(pr);
	};


	// ValueAxisScaling
	drawingsChangesMap[AscDFH.historyitem_ValueAxisScaling_SetMax] = function (oClass, value) {
		oClass.max = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ValueAxisScaling_SetMin] = function (oClass, value) {
		oClass.min = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ValueAxisScaling_SetMajorUnit] = function (oClass, value) {
		oClass.majorUnit = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ValueAxisScaling_SetMinorUnit] = function (oClass, value) {
		oClass.minorUnit = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_ValueAxisScaling_SetMax] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_ValueAxisScaling_SetMin] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_ValueAxisScaling_SetMajorUnit] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_ValueAxisScaling_SetMinorUnit] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CValueAxisScaling() {
		CBaseChartObject.call(this);
		this.max = null;
		this.min = null;
		this.majorUnit = null;
		this.minorUnit = null;
	}

	InitClass(CValueAxisScaling, CBaseChartObject, AscDFH.historyitem_type_ValueAxisScaling);

	CValueAxisScaling.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setMax(this.max);
		oCopy.setMin(this.min);
		oCopy.setMajorUnit(this.majorUnit);
		oCopy.setMinorUnit(this.minorUnit);
	}

	CValueAxisScaling.prototype.setMax = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ValueAxisScaling_SetMax, this.max, pr));
		this.max = pr;
		this.setParentToChild(pr);
	};
	CValueAxisScaling.prototype.setMin = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_ValueAxisScaling_SetMin, this.min, pr));
		this.min = pr;
	};
	CValueAxisScaling.prototype.setMajorUnit = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_ValueAxisScaling_SetMajorUnit, this.majorUnit, pr));
		this.majorUnit = pr;
	};
	CValueAxisScaling.prototype.setMinorUnit = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_ValueAxisScaling_SetMinorUnit, this.minorUnit, pr));
		this.minorUnit = pr;
	};


	// ValueColorEndPosition
	drawingsChangesMap[AscDFH.historyitem_ValueColorEndPosition_SetExtremeValue] = function (oClass, value) {
		oClass.extremeValue = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ValueColorEndPosition_SetNumber] = function (oClass, value) {
		oClass.number = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ValueColorEndPosition_SetPercent] = function (oClass, value) {
		oClass.percent = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_ValueColorEndPosition_SetExtremeValue] = window['AscDFH'].CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ValueColorEndPosition_SetNumber] = window['AscDFH'].CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_ValueColorEndPosition_SetPercent] = window['AscDFH'].CChangesDrawingsString;

	function CValueColorEndPosition() {
		CBaseChartObject.call(this);
		this.extremeValue = null;
		this.number = null;
		this.percent = null;
	}

	InitClass(CValueColorEndPosition, CBaseChartObject, AscDFH.historyitem_type_ValueColorEndPosition);

	CValueColorEndPosition.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.percent) {
			oCopy.setPercent(this.percent.createDuplicate());
		}
		oCopy.setExtremeValue(this.extremeValue);
		oCopy.setNumber(this.number);
	}

	CValueColorEndPosition.prototype.setExtremeValue = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_ValueColorEndPosition_SetExtremeValue, this.extremeValue, pr));
		this.extremeValue = pr;
	};
	CValueColorEndPosition.prototype.setNumber = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_ValueColorEndPosition_SetNumber, this.number, pr));
		this.number = pr;
	};
	CValueColorEndPosition.prototype.setPercent = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_ValueColorEndPosition_SetPercent, this.percent, pr));
		this.percent = pr;
	};


	// ValueColorMiddlePosition
	drawingsChangesMap[AscDFH.historyitem_ValueColorMiddlePosition_SetNumber] = function (oClass, value) {
		oClass.number = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ValueColorMiddlePosition_SetPercent] = function (oClass, value) {
		oClass.percent = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_ValueColorMiddlePosition_SetNumber] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_ValueColorMiddlePosition_SetPercent] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CValueColorMiddlePosition() {
		CBaseChartObject.call(this);
		this.number = null;
		this.percent = null;
	}

	InitClass(CValueColorMiddlePosition, CBaseChartObject, AscDFH.historyitem_type_ValueColorMiddlePosition);

	CValueColorMiddlePosition.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.percent) {
			oCopy.setPercent(this.percent.createDuplicate());
		}
		oCopy.setNumber(this.number);
	}	

	CValueColorMiddlePosition.prototype.setNumber = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_ValueColorMiddlePosition_SetNumber, this.number, pr));
		this.number = pr;
		this.setParentToChild(pr);
	};
	CValueColorMiddlePosition.prototype.setPercent = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_ValueColorMiddlePosition_SetPercent, this.percent, pr));
		this.percent = pr;
		this.setParentToChild(pr);
	};


	// ValueColorPositions
	drawingsChangesMap[AscDFH.historyitem_ValueColorPositions_SetMin] = function (oClass, value) {
		oClass.min = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ValueColorPositions_SetMid] = function (oClass, value) {
		oClass.mid = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ValueColorPositions_SetMax] = function (oClass, value) {
		oClass.max = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ValueColorPositions_SetCount] = function (oClass, value) {
		oClass.count = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_ValueColorPositions_SetMin] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_ValueColorPositions_SetMid] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_ValueColorPositions_SetMax] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_ValueColorPositions_SetCount] = window['AscDFH'].CChangesDrawingsLong;

	function CValueColorPositions() {
		CBaseChartObject.call(this);
		this.min = null;
		this.mid = null;
		this.max = null;
		this.count = null;
	}

	InitClass(CValueColorPositions, CBaseChartObject, AscDFH.historyitem_type_ValueColorPositions);

	CValueColorPositions.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		if (this.min) {
			oCopy.setMin(this.min.createDuplicate());
		}
		if (this.mid) {
			oCopy.setMid(this.mid.createDuplicate());
		}
		if (this.max) {
			oCopy.setMax(this.max.createDuplicate());
		}
		oCopy.setCount(this.count);
	}	

	CValueColorPositions.prototype.setMin = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ValueColorPositions_SetMin, this.min, pr));
		this.min = pr;
		this.setParentToChild(pr);
	};
	CValueColorPositions.prototype.setMid = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ValueColorPositions_SetMid, this.mid, pr));
		this.mid = pr;
		this.setParentToChild(pr);
	};
	CValueColorPositions.prototype.setMax = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ValueColorPositions_SetMax, this.max, pr));
		this.max = pr;
		this.setParentToChild(pr);
	};
	CValueColorPositions.prototype.setCount = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_ValueColorPositions_SetCount, this.count, pr));
		this.count = pr;
	};


	// ValueColors
	drawingsChangesMap[AscDFH.historyitem_ValueColors_SetMinColor] = function (oClass, value) {
		oClass.minColor = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ValueColors_SetMidColor] = function (oClass, value) {
		oClass.midColor = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ValueColors_SetMaxColor] = function (oClass, value) {
		oClass.maxColor = value;
	};
	AscDFH.changesFactory[AscDFH.historyitem_ValueColors_SetMinColor] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_ValueColors_SetMidColor] = window['AscDFH'].CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_ValueColors_SetMaxColor] = window['AscDFH'].CChangesDrawingsObjectNoId;

	function CValueColors() {
		CBaseChartObject.call(this);
		this.minColor = null;
		this.midColor = null;
		this.maxColor = null;
	}

	InitClass(CValueColors, CBaseChartObject, AscDFH.historyitem_type_ValueColors);

	CValueColors.prototype.fillObject = function (oCopy) {
		CBaseChartObject.prototype.fillObject.call(this, oCopy);
		oCopy.setMinColor(this.minColor);
		oCopy.setMidColor(this.midColor);
		oCopy.setMaxColor(this.maxColor);
	}	

	CValueColors.prototype.setMinColor = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ValueColors_SetMinColor, this.minColor, pr));
		this.minColor = pr;
		this.setParentToChild(pr);
	};
	CValueColors.prototype.setMidColor = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ValueColors_SetMidColor, this.midColor, pr));
		this.midColor = pr;
		this.setParentToChild(pr);
	};
	CValueColors.prototype.setMaxColor = function (pr) {
		History.CanAddChanges() && History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ValueColors_SetMaxColor, this.maxColor, pr));
		this.maxColor = pr;
		this.setParentToChild(pr);
	};


	// Simple Types


	// // SidePos
	// drawingsChangesMap[AscDFH.historyitem_SidePos_SetSidePos] = function (oClass, value) {
	//     oClass.sidePos = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_SidePos_SetSidePos] = window['AscDFH'].CChangesDrawingsLong;
	// function CSidePos() {
	//     CBaseChartObject.call(this);
	//     this.sidePos = null;
	// }

	// InitClass(CSidePos, CBaseChartObject, AscDFH.historyitem_type_SidePos);

	// CSidePos.prototype.setSidePos = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_SidePos_SetSidePos, this.sidePos, pr));
	//     this.sidePos = pr;
	// };


	// // PosAlign
	// drawingsChangesMap[AscDFH.historyitem_PosAlign_SetPosAlign] = function (oClass, value) {
	//     oClass.posAlign = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_PosAlign_SetPosAlign] = window['AscDFH'].CChangesDrawingsLong;
	// function CPosAlign() {
	//     CBaseChartObject.call(this);
	//     this.posAlign = null;
	// }

	// InitClass(CPosAlign, CBaseChartObject, AscDFH.historyitem_type_PosAlign);

	// CPosAlign.prototype.setPosAlign = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PosAlign_SetPosAlign, this.posAlign, pr));
	//     this.posAlign = pr;
	// };


	// // AxisUnit
	// drawingsChangesMap[AscDFH.historyitem_AxisUnit_SetAxisUnit] = function (oClass, value) {
	//     oClass.axisUnit = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_AxisUnit_SetAxisUnit] = window['AscDFH'].CChangesDrawingsLong;
	// function CAxisUnit() {
	//     CBaseChartObject.call(this);
	//     this.axisUnit = null;
	// }

	// InitClass(CAxisUnit, CBaseChartObject, AscDFH.historyitem_type_AxisUnit);

	// CAxisUnit.prototype.setAxisUnit = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_AxisUnit_SetAxisUnit, this.axisUnit, pr));
	//     this.axisUnit = pr;
	// };


	// // FormulaDirection
	// drawingsChangesMap[AscDFH.historyitem_FormulaDirection_SetFormulaDirection] = function (oClass, value) {
	//     oClass.formulaDirection = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_FormulaDirection_SetFormulaDirection] = window['AscDFH'].CChangesDrawingsLong;
	// function CFormulaDirection() {
	//     CBaseChartObject.call(this);
	//     this.formulaDirection = null;
	// }

	// InitClass(CFormulaDirection, CBaseChartObject, AscDFH.historyitem_type_FormulaDirection);

	// CFormulaDirection.prototype.setFormulaDirection = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_FormulaDirection_SetFormulaDirection, this.formulaDirection, pr));
	//     this.formulaDirection = pr;
	// };


	// // IntervalClosedSide
	// drawingsChangesMap[AscDFH.historyitem_IntervalClosedSide_SetIntervalClosedSide] = function (oClass, value) {
	//     oClass.intervalClosedSide = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_IntervalClosedSide_SetIntervalClosedSide] = window['AscDFH'].CChangesDrawingsLong;
	// function CIntervalClosedSide() {
	//     CBaseChartObject.call(this);
	//     this.intervalClosedSide = null;
	// }

	// InitClass(CIntervalClosedSide, CBaseChartObject, AscDFH.historyitem_type_IntervalClosedSide);

	// CIntervalClosedSide.prototype.setIntervalClosedSide = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_IntervalClosedSide_SetIntervalClosedSide, this.intervalClosedSide, pr));
	//     this.intervalClosedSide = pr;
	// };


	// // DimensionType (NumericDimensionType)
	// drawingsChangesMap[AscDFH.historyitem_DimensionType_SetDimensionType] = function (oClass, value) {
	//     oClass.dimensionType = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_DimensionType_SetDimensionType] = window['AscDFH'].CChangesDrawingsLong;
	// function CDimensionType() {
	//     CBaseChartObject.call(this);
	//     this.dimensionType = null;
	// }

	// InitClass(CDimensionType, CBaseChartObject, AscDFH.historyitem_type_DimensionType);

	// CDimensionType.prototype.setDimensionType = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DimensionType_SetDimensionType, this.dimensionType, pr));
	//     this.dimensionType = pr;
	// };


	// // QuartileMethod
	// drawingsChangesMap[AscDFH.historyitem_QuartileMethod_SetQuartileMethod] = function (oClass, value) {
	//     oClass.quartileMethod = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_QuartileMethod_SetQuartileMethod] = window['AscDFH'].CChangesDrawingsLong;
	// function CQuartileMethod() {
	//     CBaseChartObject.call(this);
	//     this.quartileMethod = null;
	// }

	// InitClass(CQuartileMethod, CBaseChartObject, AscDFH.historyitem_type_QuartileMethod);

	// CQuartileMethod.prototype.setQuartileMethod = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_QuartileMethod_SetQuartileMethod, this.quartileMethod, pr));
	//     this.quartileMethod = pr;
	// };


	// // DataLabelPos
	// drawingsChangesMap[AscDFH.historyitem_DataLabelPos_SetDataLabelPos] = function (oClass, value) {
	//     oClass.dataLabelPos = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_DataLabelPos_SetDataLabelPos] = window['AscDFH'].CChangesDrawingsLong;
	// function CDataLabelPos() {
	//     CBaseChartObject.call(this);
	//     this.dataLabelPos = null;
	// }

	// InitClass(CDataLabelPos, CBaseChartObject, AscDFH.historyitem_type_DataLabelPos);

	// CDataLabelPos.prototype.setDataLabelPos = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DataLabelPos_SetDataLabelPos, this.dataLabelPos, pr));
	//     this.dataLabelPos = pr;
	// };


	// // SeriesLayout
	// drawingsChangesMap[AscDFH.historyitem_SeriesLayout_SetSeriesLayout] = function (oClass, value) {
	//     oClass.seriesLayout = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_SeriesLayout_SetSeriesLayout] = window['AscDFH'].CChangesDrawingsLong;
	// function CSeriesLayout() {
	//     CBaseChartObject.call(this);
	//     this.seriesLayout = null;
	// }

	// InitClass(CSeriesLayout, CBaseChartObject, AscDFH.historyitem_type_SeriesLayout);

	// CSeriesLayout.prototype.setSeriesLayout = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_SeriesLayout_SetSeriesLayout, this.seriesLayout, pr));
	//     this.seriesLayout = pr;
	// };


	// // TickMarksType
	// drawingsChangesMap[AscDFH.historyitem_TickMarksType_SetTickMarksType] = function (oClass, value) {
	//     oClass.tickMarksType = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_TickMarksType_SetTickMarksType] = window['AscDFH'].CChangesDrawingsLong;
	// function CTickMarksType() {
	//     CBaseChartObject.call(this);
	//     this.tickMarksType = null;
	// }

	// InitClass(CTickMarksType, CBaseChartObject, AscDFH.historyitem_type_TickMarksType);

	// CTickMarksType.prototype.setTickMarksType = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_TickMarksType_SetTickMarksType, this.tickMarksType, pr));
	//     this.tickMarksType = pr;
	// };


	// // EntityType
	// drawingsChangesMap[AscDFH.historyitem_EntityType_SetEntityType] = function (oClass, value) {
	//     oClass.entityType = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_EntityType_SetEntityType] = window['AscDFH'].CChangesDrawingsLong;
	// function CEntityType() {
	//     CBaseChartObject.call(this);
	//     this.entityType = null;
	// }

	// InitClass(CEntityType, CBaseChartObject, AscDFH.historyitem_type_EntityType);

	// CEntityType.prototype.setEntityType = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_EntityType_SetEntityType, this.entityType, pr));
	//     this.entityType = pr;
	// };


	// // GeoProjectionType
	// drawingsChangesMap[AscDFH.historyitem_GeoProjectionType_SetGeoProjectionType] = function (oClass, value) {
	//     oClass.geoProjectionType = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_GeoProjectionType_SetGeoProjectionType] = window['AscDFH'].CChangesDrawingsLong;
	// function CGeoProjectionType() {
	//     CBaseChartObject.call(this);
	//     this.geoProjectionType = null;
	// }

	// InitClass(CGeoProjectionType, CBaseChartObject, AscDFH.historyitem_type_GeoProjectionType);

	// CGeoProjectionType.prototype.setGeoProjectionType = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_GeoProjectionType_SetGeoProjectionType, this.geoProjectionType, pr));
	//     this.geoProjectionType = pr;
	// };


	// // GeoMappingLevel
	// drawingsChangesMap[AscDFH.historyitem_GeoMappingLevel_SetGeoMappingLevel] = function (oClass, value) {
	//     oClass.geoMappingLevel = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_GeoMappingLevel_SetGeoMappingLevel] = window['AscDFH'].CChangesDrawingsLong;
	// function CGeoMappingLevel() {
	//     CBaseChartObject.call(this);
	//     this.geoMappingLevel = null;
	// }

	// InitClass(CGeoMappingLevel, CBaseChartObject, AscDFH.historyitem_type_GeoMappingLevel);

	// CGeoMappingLevel.prototype.setGeoMappingLevel = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_GeoMappingLevel_SetGeoMappingLevel, this.geoMappingLevel, pr));
	//     this.geoMappingLevel = pr;
	// };


	// // PageOrientation
	// drawingsChangesMap[AscDFH.historyitem_PageOrientation_SetPageOrientation] = function (oClass, value) {
	//     oClass.pageOrientation = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_PageOrientation_SetPageOrientation] = window['AscDFH'].CChangesDrawingsLong;
	// function CPageOrientation() {
	//     CBaseChartObject.call(this);
	//     this.pageOrientation = null;
	// }

	// InitClass(CPageOrientation, CBaseChartObject, AscDFH.historyitem_type_PageOrientation);

	// CPageOrientation.prototype.setPageOrientation = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PageOrientation_SetPageOrientation, this.pageOrientation, pr));
	//     this.pageOrientation = pr;
	// };


	// // LabelLayout (ST_ParentLabelLayout)
	// drawingsChangesMap[AscDFH.historyitem_LabelLayout_SetLabelLayout] = function (oClass, value) {
	//     oClass.labelLayout = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_LabelLayout_SetLabelLayout] = window['AscDFH'].CChangesDrawingsLong;
	// function CLabelLayout() {
	//     CBaseChartObject.call(this);
	//     this.labelLayout = null;
	// }

	// InitClass(CLabelLayout, CBaseChartObject, AscDFH.historyitem_type_LabelLayout);

	// CLabelLayout.prototype.setPageOrientation = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_LabelLayout_SetLabelLayout, this.labelLayout, pr));
	//     this.labelLayout = pr;
	// };


	// // RegionLabelLayout (ST)
	// drawingsChangesMap[AscDFH.historyitem_RegionLabelLayout_SetRegionLabelLayout] = function (oClass, value) {
	//     oClass.regionLabelLayout = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_RegionLabelLayout_SetRegionLabelLayout] = window['AscDFH'].CChangesDrawingsLong;
	// function CRegionLabelLayout() {
	//     CBaseChartObject.call(this);
	//     this.regionLabelLayout = null;
	// }

	// InitClass(CRegionLabelLayout, CBaseChartObject, AscDFH.historyitem_type_RegionLabelLayout);

	// CRegionLabelLayout.prototype.setPageOrientation = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_RegionLabelLayout_SetRegionLabelLayout, this.regionLabelLayout, pr));
	//     this.regionLabelLayout = pr;
	// };


	// // DoubleOrAutomatic
	// drawingsChangesMap[AscDFH.historyitem_DoubleOrAutomatic_SetValue] = function (oClass, value) {
	//     oClass.value = value;
	// };
	// drawingsChangesMap[AscDFH.historyitem_DoubleOrAutomatic_SetType] = function (oClass, value) {
	//     oClass.type = value;
	// };
	// AscDFH.changesFactory[AscDFH.historyitem_DoubleOrAutomatic_SetValue] = window['AscDFH'].CChangesDrawingsLong;
	// AscDFH.changesFactory[AscDFH.historyitem_DoubleOrAutomatic_SetType] = window['AscDFH'].CChangesDrawingsDouble2;

	// const EDoubleOrAutomatic = { typeAuto: 0, typeDouble: 1};
	// function CDoubleOrAutomatic() {
	//     CBaseChartObject.call(this);
	//     this.value = 0;
	//     this.type = EDoubleOrAutomatic.typeAuto;
	// }

	// InitClass(CDoubleOrAutomatic, CBaseChartObject, AscDFH.historyitem_type_DoubleOrAutomatic);
	// CDoubleOrAutomatic.prototype.setValue = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DoubleOrAutomatic_SetValue, this.value, pr));
	//     this.value = pr;
	// };
	// CDoubleOrAutomatic.prototype.setType = function (pr) {
	//     History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_DoubleOrAutomatic_SetType, this.type, pr));
	//     this.type = pr;
	// };

	//--------------------------------------------------------export----------------------------------------------------
	window['AscFormat'] = window['AscFormat'] || {};
	window['AscFormat'].CAddress = CAddress;
	window['AscFormat'].CAxis = CAxis;
	// window['AscFormat'].CChartExTitle = CChartExTitle;
	// window['AscFormat'].CAxisTitle = CAxisTitle;
	window['AscFormat'].CAxisUnits = CAxisUnits;
	window['AscFormat'].CAxisUnitsLabel = CAxisUnitsLabel;
	window['AscFormat'].CBinning = CBinning;
	window['AscFormat'].CCategoryAxisScaling = CCategoryAxisScaling;
	window['AscFormat'].CChartData = CChartData;
	window['AscFormat'].CClear = CClear;
	window['AscFormat'].CCopyrights = CCopyrights;
	window['AscFormat'].CData = CData;
	window['AscFormat'].CDataLabel = CDataLabel;
	window['AscFormat'].CDataLabelHidden = CDataLabelHidden;
	window['AscFormat'].CDataLabels = CDataLabels;
	window['AscFormat'].CDataLabelVisibilities = CDataLabelVisibilities;
	window['AscFormat'].CDataPoint = CDataPoint;
	// window['AscFormat'].CExtension = CExtension;
	// window['AscFormat'].CExtensionList = CExtensionList;
	// window['AscFormat'].CExternalData = CExternalData;
	window['AscFormat'].CFormatOverride = CFormatOverride;
	window['AscFormat'].CFormatOverrides = CFormatOverrides;
	window['AscFormat'].CFormula = CFormula;
	window['AscFormat'].CGeoCache = CGeoCache;
	window['AscFormat'].CGeoChildEntities = CGeoChildEntities;
	window['AscFormat'].CGeoChildEntitiesQuery = CGeoChildEntitiesQuery;
	window['AscFormat'].CGeoChildEntitiesQueryResult = CGeoChildEntitiesQueryResult;
	window['AscFormat'].CGeoChildEntitiesQueryResults = CGeoChildEntitiesQueryResults;
	window['AscFormat'].CGeoChildTypes = CGeoChildTypes;
	window['AscFormat'].CGeoData = CGeoData;
	window['AscFormat'].CGeoDataEntityQuery = CGeoDataEntityQuery;
	window['AscFormat'].CGeoDataEntityQueryResult = CGeoDataEntityQueryResult;
	window['AscFormat'].CGeoDataEntityQueryResults = CGeoDataEntityQueryResults;
	window['AscFormat'].CGeoDataPointQuery = CGeoDataPointQuery;
	window['AscFormat'].CGeoDataPointToEntityQuery = CGeoDataPointToEntityQuery;
	window['AscFormat'].CGeoDataPointToEntityQueryResult = CGeoDataPointToEntityQueryResult;
	window['AscFormat'].CGeoDataPointToEntityQueryResults = CGeoDataPointToEntityQueryResults;
	window['AscFormat'].CGeography = CGeography;
	window['AscFormat'].CGeoHierarchyEntity = CGeoHierarchyEntity;
	window['AscFormat'].CGeoLocation = CGeoLocation;
	window['AscFormat'].CGeoLocationQuery = CGeoLocationQuery;
	window['AscFormat'].CGeoLocationQueryResult = CGeoLocationQueryResult;
	window['AscFormat'].CGeoLocationQueryResults = CGeoLocationQueryResults;
	window['AscFormat'].CGeoLocations = CGeoLocations;
	window['AscFormat'].CGeoPolygon = CGeoPolygon;
	window['AscFormat'].CGeoPolygons = CGeoPolygons;
	window['AscFormat'].CGridlines = CGridlines;
	// window['AscFormat'].CHeaderFooter = CHeaderFooter;
	// window['AscFormat'].CChartExLegend = CChartExLegend;
	// window['AscFormat'].CNumberColorPosition = CNumberColorPosition;
	// window['AscFormat'].CNumberFormat = CNumberFormat;

	// window['AscFormat'].CPageMargins = CPageMargins;
	// window['AscFormat'].CPageSetup = CPageSetup;
	// window['AscFormat'].CParentLabelLayout = CParentLabelLayout;
	window['AscFormat'].CPercentageColorPosition = CPercentageColorPosition;
	// window['AscFormat'].CChartExPlotArea = CChartExPlotArea;
	window['AscFormat'].CPlotAreaRegion = CPlotAreaRegion;
	window['AscFormat'].CPlotSurface = CPlotSurface;
	// window['AscFormat'].CPrintSettings = CPrintSettings;
	window['AscFormat'].CSeries = CSeries;
	window['AscFormat'].CSeriesElementVisibilities = CSeriesElementVisibilities;
	window['AscFormat'].CSeriesLayoutProperties = CSeriesLayoutProperties;
	window['AscFormat'].CStatistics = CStatistics;
	window['AscFormat'].CDimension = CDimension;
	window['AscFormat'].CNumericDimension = CNumericDimension;
	window['AscFormat'].CStringDimension = CStringDimension;
	window['AscFormat'].CSubtotals = CSubtotals;
	// window['AscFormat'].CText = CText;
	window['AscFormat'].CTextData = CTextData;
	// window['AscFormat'].CTickLabels = CTickLabels;
	window['AscFormat'].CTickMarks = CTickMarks;
	window['AscFormat'].CValueAxisScaling = CValueAxisScaling;
	window['AscFormat'].CValueColorEndPosition = CValueColorEndPosition;
	window['AscFormat'].CValueColorMiddlePosition = CValueColorMiddlePosition;
	window['AscFormat'].CValueColorPositions = CValueColorPositions;
	window['AscFormat'].CValueColors = CValueColors;
	// ---------------------------------------------
	// Simple Types
	// ---------------------------------------------
	// window['AscFormat'].CSidePos = CSidePos;
	// window['AscFormat'].CPosAlign = CPosAlign;
	// window['AscFormat'].CAxisUnit = CAxisUnit;
	// window['AscFormat'].CFormulaDirection = CFormulaDirection;
	// window['AscFormat'].CIntervalClosedSide = CIntervalClosedSide;
	// window['AscFormat'].CDimensionType = CDimensionType;
	// window['AscFormat'].CQuartileMethod = CQuartileMethod;
	// window['AscFormat'].CDataLabelPos = CDataLabelPos;
	// window['AscFormat'].CSeriesLayout = CSeriesLayout;
	// window['AscFormat'].CTickMarksType = CTickMarksType;
	// window['AscFormat'].CEntityType = CEntityType;
	// window['AscFormat'].CGeoProjectionType = CGeoProjectionType;
	// window['AscFormat'].CGeoMappingLevel = CGeoMappingLevel;
	// window['AscFormat'].CPageOrientation = CPageOrientation;
	// window['AscFormat'].CLabelLayout = CLabelLayout;
	// window['AscFormat'].CRegionLabelLayout = CRegionLabelLayout;
	// window['AscFormat'].CDoubleOrAutomatic = CDoubleOrAutomatic;

	window['AscFormat'].SERIES_LAYOUT_BOX_WHISKER = SERIES_LAYOUT_BOX_WHISKER;
	window['AscFormat'].SERIES_LAYOUT_CLUSTERED_COLUMN = SERIES_LAYOUT_CLUSTERED_COLUMN;
	window['AscFormat'].SERIES_LAYOUT_FUNNEL = SERIES_LAYOUT_FUNNEL;
	window['AscFormat'].SERIES_LAYOUT_PARETO_LINE = SERIES_LAYOUT_PARETO_LINE;
	window['AscFormat'].SERIES_LAYOUT_REGION_MAP = SERIES_LAYOUT_REGION_MAP;
	window['AscFormat'].SERIES_LAYOUT_SUNBURST = SERIES_LAYOUT_SUNBURST;
	window['AscFormat'].SERIES_LAYOUT_TREEMAP = SERIES_LAYOUT_TREEMAP;
	window['AscFormat'].SERIES_LAYOUT_WATERFALL = SERIES_LAYOUT_WATERFALL;

	window['AscFormat'].DATA_LABEL_POS_BEST_FIT = DATA_LABEL_POS_BEST_FIT;
	window['AscFormat'].DATA_LABEL_POS_B = DATA_LABEL_POS_B;
	window['AscFormat'].DATA_LABEL_POS_CTR = DATA_LABEL_POS_CTR;
	window['AscFormat'].DATA_LABEL_POS_IN_BASE = DATA_LABEL_POS_IN_BASE;
	window['AscFormat'].DATA_LABEL_POS_IN_END = DATA_LABEL_POS_IN_END;
	window['AscFormat'].DATA_LABEL_POS_L = DATA_LABEL_POS_L;
	window['AscFormat'].DATA_LABEL_POS_OUT_END = DATA_LABEL_POS_OUT_END;
	window['AscFormat'].DATA_LABEL_POS_R = DATA_LABEL_POS_R;
	window['AscFormat'].DATA_LABEL_POS_T = DATA_LABEL_POS_T;

	window['AscFormat'].PARENT_LABEL_LAYOUT_NONE = PARENT_LABEL_LAYOUT_NONE;
	window['AscFormat'].PARENT_LABEL_LAYOUT_BANNER = PARENT_LABEL_LAYOUT_BANNER;
	window['AscFormat'].PARENT_LABEL_LAYOUT_OVERLAPPING = PARENT_LABEL_LAYOUT_OVERLAPPING;

	window['AscFormat'].REGION_LABEL_LAYOUT_NONE = REGION_LABEL_LAYOUT_NONE;
	window['AscFormat'].REGION_LABEL_LAYOUT_BEST_FIT_ONLY = REGION_LABEL_LAYOUT_BEST_FIT_ONLY;
	window['AscFormat'].REGION_LABEL_LAYOUT_SHOW_ALL = REGION_LABEL_LAYOUT_SHOW_ALL;

	window['AscFormat'].INTERVAL_CLOSED_SIDE_L = INTERVAL_CLOSED_SIDE_L;
	window['AscFormat'].INTERVAL_CLOSED_SIDE_R = INTERVAL_CLOSED_SIDE_R;

	window['AscFormat'].AXIS_UNIT_HUNDREDS = AXIS_UNIT_HUNDREDS;
	window['AscFormat'].AXIS_UNIT_THOUSANDS = AXIS_UNIT_THOUSANDS;
	window['AscFormat'].AXIS_UNIT_TEN_THOUSANDS = AXIS_UNIT_TEN_THOUSANDS;
	window['AscFormat'].AXIS_UNIT_HUNDRED_THOUSANDS = AXIS_UNIT_HUNDRED_THOUSANDS;
	window['AscFormat'].AXIS_UNIT_MILLIONS = AXIS_UNIT_MILLIONS;
	window['AscFormat'].AXIS_UNIT_TEN_MILLIONS = AXIS_UNIT_TEN_MILLIONS;
	window['AscFormat'].AXIS_UNIT_HUNDRED_MILLIONS = AXIS_UNIT_HUNDRED_MILLIONS;
	window['AscFormat'].AXIS_UNIT_BILLIONS = AXIS_UNIT_BILLIONS;
	window['AscFormat'].AXIS_UNIT_TRILLIONS = AXIS_UNIT_TRILLIONS;
	window['AscFormat'].AXIS_UNIT_PERCENTAGE = AXIS_UNIT_PERCENTAGE;

	window['AscFormat'].SIDE_POS_L = SIDE_POS_L;
	window['AscFormat'].SIDE_POS_T = SIDE_POS_T;
	window['AscFormat'].SIDE_POS_R = SIDE_POS_R;
	window['AscFormat'].SIDE_POS_B = SIDE_POS_B;

	window['AscFormat'].POS_ALIGN_MIN = POS_ALIGN_MIN;
	window['AscFormat'].POS_ALIGN_CTR = POS_ALIGN_CTR;
	window['AscFormat'].POS_ALIGN_MAX = POS_ALIGN_MAX;

	window['AscFormat'].TICK_MARKS_TYPE_IN = TICK_MARKS_TYPE_IN;
	window['AscFormat'].TICK_MARKS_TYPE_OUT = TICK_MARKS_TYPE_OUT;
	window['AscFormat'].TICK_MARKS_TYPE_CROSS = TICK_MARKS_TYPE_CROSS;
	window['AscFormat'].TICK_MARKS_TYPE_NONE = TICK_MARKS_TYPE_NONE;

	window['AscFormat'].QUARTILE_METHOD_INCLUSIVE = QUARTILE_METHOD_INCLUSIVE;
	window['AscFormat'].QUARTILE_METHOD_EXCLUSIVE = QUARTILE_METHOD_EXCLUSIVE;

	window['AscFormat'].STRING_DIMENSION_TYPE_CAT = STRING_DIMENSION_TYPE_CAT;
	window['AscFormat'].STRING_DIMENSION_TYPE_COLOR_STR = STRING_DIMENSION_TYPE_COLOR_STR;

	window['AscFormat'].NUMERIC_DIMENSION_TYPE_VAL = NUMERIC_DIMENSION_TYPE_VAL;
	window['AscFormat'].NUMERIC_DIMENSION_TYPE_X = NUMERIC_DIMENSION_TYPE_X;
	window['AscFormat'].NUMERIC_DIMENSION_TYPE_Y = NUMERIC_DIMENSION_TYPE_Y;
	window['AscFormat'].NUMERIC_DIMENSION_TYPE_SIZE = NUMERIC_DIMENSION_TYPE_SIZE;
	window['AscFormat'].NUMERIC_DIMENSION_TYPE_COLOR_VAL = NUMERIC_DIMENSION_TYPE_COLOR_VAL;

	window['AscFormat'].FORMULA_DIRECTION_COL = FORMULA_DIRECTION_COL;
	window['AscFormat'].FORMULA_DIRECTION_ROW = FORMULA_DIRECTION_ROW;
})(window);

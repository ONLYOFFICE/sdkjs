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
//Generated code

var c_oAscSourceType = {
	Worksheet: 0,
	External: 1,
	Consolidation: 2,
	Scenario: 3
};
var c_oAscAxis = {
	AxisRow: 0,
	AxisCol: 1,
	AxisPage: 2,
	AxisValues: 3
};
var c_oAscFieldSortType = {
	Manual: 0,
	Ascending: 1,
	Descending: 2
};
var c_oAscDataConsolidateFunction = {
	Average: 1,
	CountNums: 2,
	Count: 3,
	Max: 4,
	Min: 5,
	Product: 6,
	StdDev: 7,
	StdDevp: 8,
	Sum: 9,
	Var: 10,
	Varp: 11
};
var c_oAscShowDataAs = {
	Normal: 0,
	Difference: 1,
	Percent: 2,
	PercentDiff: 3,
	RunTotal: 4,
	PercentOfRow: 5,
	PercentOfCol: 6,
	PercentOfTotal: 7,
	Index: 8
};
var c_oAscFormatAction = {
	Blank: 0,
	Formatting: 1,
	Drill: 2,
	Formula: 3
};
var c_oAscScope = {
	Selection: 0,
	Data: 1,
	Field: 2
};
var c_oAscType = {
	None: 0,
	All: 1,
	Row: 2,
	Column: 3
};
var c_oAscPivotFilterType = {
	Unknown: 0,
	Count: 1,
	Percent: 2,
	Sum: 3,
	CaptionEqual: 4,
	CaptionNotEqual: 5,
	CaptionBeginsWith: 6,
	CaptionNotBeginsWith: 7,
	CaptionEndsWith: 8,
	CaptionNotEndsWith: 9,
	CaptionContains: 10,
	CaptionNotContains: 11,
	CaptionGreaterThan: 12,
	CaptionGreaterThanOrEqual: 13,
	CaptionLessThan: 14,
	CaptionLessThanOrEqual: 15,
	CaptionBetween: 16,
	CaptionNotBetween: 17,
	ValueEqual: 18,
	ValueNotEqual: 19,
	ValueGreaterThan: 20,
	ValueGreaterThanOrEqual: 21,
	ValueLessThan: 22,
	ValueLessThanOrEqual: 23,
	ValueBetween: 24,
	ValueNotBetween: 25,
	DateEqual: 26,
	DateNotEqual: 27,
	DateOlderThan: 28,
	DateOlderThanOrEqual: 29,
	DateNewerThan: 30,
	DateNewerThanOrEqual: 31,
	DateBetween: 32,
	DateNotBetween: 33,
	Tomorrow: 34,
	Today: 35,
	Yesterday: 36,
	NextWeek: 37,
	ThisWeek: 38,
	LastWeek: 39,
	NextMonth: 40,
	ThisMonth: 41,
	LastMonth: 42,
	NextQuarter: 43,
	ThisQuarter: 44,
	LastQuarter: 45,
	NextYear: 46,
	ThisYear: 47,
	LastYear: 48,
	YearToDate: 49,
	Q1: 50,
	Q2: 51,
	Q3: 52,
	Q4: 53,
	M1: 54,
	M2: 55,
	M3: 56,
	M4: 57,
	M5: 58,
	M6: 59,
	M7: 60,
	M8: 61,
	M9: 62,
	M10: 63,
	M11: 64,
	M12: 65
};
var c_oAscSortType = {
	None: 0,
	Ascending: 1,
	Descending: 2,
	AscendingAlpha: 3,
	DescendingAlpha: 4,
	AscendingNatural: 5,
	DescendingNatural: 6
};
var c_oAscPivotAreaType = {
	None: 0,
	Normal: 1,
	Data: 2,
	All: 3,
	Origin: 4,
	Button: 5,
	TopEnd: 6
};
var c_oAscGroupBy = {
	Range: 0,
	Seconds: 1,
	Minutes: 2,
	Hours: 3,
	Days: 4,
	Months: 5,
	Quarters: 6,
	Years: 7
};
var c_oAscSortMethod = {
	Stroke: 0,
	PinYin: 1,
	None: 2
};
var c_oAscDynamicFilterType = {
	Null: 0,
	AboveAverage: 1,
	BelowAverage: 2,
	Tomorrow: 3,
	Today: 4,
	Yesterday: 5,
	NextWeek: 6,
	ThisWeek: 7,
	LastWeek: 8,
	NextMonth: 9,
	ThisMonth: 10,
	LastMonth: 11,
	NextQuarter: 12,
	ThisQuarter: 13,
	LastQuarter: 14,
	NextYear: 15,
	ThisYear: 16,
	LastYear: 17,
	YearToDate: 18,
	Q1: 19,
	Q2: 20,
	Q3: 21,
	Q4: 22,
	M1: 23,
	M2: 24,
	M3: 25,
	M4: 26,
	M5: 27,
	M6: 28,
	M7: 29,
	M8: 30,
	M9: 31,
	M10: 32,
	M11: 33,
	M12: 34
};
var c_oAscCalendarType = {
	Gregorian: 0,
	GregorianUs: 1,
	GregorianMeFrench: 2,
	GregorianArabic: 3,
	Hijri: 4,
	Hebrew: 5,
	Taiwan: 6,
	Japan: 7,
	Thai: 8,
	Korea: 9,
	Saka: 10,
	GregorianXlitEnglish: 11,
	GregorianXlitFrench: 12,
	None: 13
};
var c_oAscIconSetType = {
	ThreeArrows: 0,
	ThreeArrowsGray: 1,
	ThreeFlags: 2,
	ThreeTrafficLights1: 3,
	ThreeTrafficLights2: 4,
	ThreeSigns: 5,
	ThreeSymbols: 6,
	ThreeSymbols2: 7,
	FourArrows: 8,
	FourArrowsGray: 9,
	FourRedToBlack: 10,
	FourRating: 11,
	FourTrafficLights: 12,
	FiveArrows: 13,
	FiveArrowsGray: 14,
	FiveRating: 15,
	FiveQuarters: 16
};
var c_oAscSortBy = {
	Value: 0,
	CellColor: 1,
	FontColor: 2,
	Icon: 3
};
var c_oAscFilterOperator = {
	Equal: 0,
	LessThan: 1,
	LessThanOrEqual: 2,
	NotEqual: 3,
	GreaterThanOrEqual: 4,
	GreaterThan: 5
};
var c_oAscDateTimeGrouping = {
	Year: 0,
	Month: 1,
	Day: 2,
	Hour: 3,
	Minute: 4,
	Second: 5
};
var c_oAscAllocationMethod = {
	EqualAllocation: 0,
	EqualIncrement: 1,
	WeightedAllocation: 2,
	WeightedIncrement: 3
};

var st_VALUES = -2;
var DATA_CAPTION = 'Values';
var GRAND_TOTAL_CAPTION = 'Grand Total';
var ROW_HEADER_CAPTION = 'Row Labels';
var COL_HEADER_CAPTION = 'Column Labels';
var PAGE_ALL_CAPTION = '(All)';
var FIELD_CAPTION = '%1 of %2';
var NEW_PIVOT_LAST_COL_OFFSET = 2;
var NEW_PIVOT_LAST_ROW_OFFSET = 17;
var NEW_PIVOT_ROW = 2;
var NEW_PIVOT_COL = 0;

var cDate = Asc.cDate;

function ctItemSort(items, sharedItems) {
	items.sort(function(a, b) {
		var sharedItem = sharedItems.Items.get(a.x);
		var aType = sharedItem.type;
		var aVal = sharedItem.val;
		sharedItem = sharedItems.Items.get(b.x);
		var bType = sharedItem.type;
		var bVal = sharedItem.val;
		if (aType === bType) {
			if (c_oAscPivotRecType.String === aType) {
				if (aVal > bVal) {
					return 1;
				}
				if (aVal < bVal) {
					return -1;
				}
			} else if (c_oAscPivotRecType.Missing !== aType) {
				return aVal - bVal;
			}
		} else {
			return aType - bType;
		}
		return 0;
	});
}
function PivotDataElem(dataLength) {
	this.subtotal = this.vals = {};
	this.total = new Array(dataLength);
	for (var i = 0; i < dataLength; ++i) {
		this.total[i] = new AscCommonExcel.StatisticOnlineAlgorithm();
	}
}
PivotDataElem.prototype.unionTotal = function(val){
	for (var i = 0; i < this.total.length; ++i) {
		this.total[i].union(val.total[i]);
	}
}
function setTableProperty(pivot, oldVal, newVal, addToHistory, historyType, changeData) {
	if (oldVal === newVal) {
		return;
	}
	if (addToHistory) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, historyType, pivot.worksheet.getId(), null,
			new AscCommonExcel.UndoRedoData_PivotTable(pivot.Get_Id(), oldVal, newVal));
	}
	if (pivot && changeData) {
		pivot.setChanged(true);
	}
}
function setFieldProperty(pivot, index, oldVal, newVal, addToHistory, historyType, changeData) {
	if (oldVal === newVal) {
		return;
	}
	if (addToHistory) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotFields, historyType, pivot.worksheet.getId(), null,
			new AscCommonExcel.UndoRedoData_PivotField(pivot.Get_Id(), index, oldVal, newVal));
	}
	if (pivot && changeData) {
		pivot.setChanged(true);
	}
};
function toXmlWithLength(w, elem) {
	var StartPos = w.GetCurPosition();
	w.WriteLong(0);
	elem.toXml(w);
	var EndPos = w.GetCurPosition();
	w.Seek(StartPos);
	w.WriteLong(EndPos - StartPos - 4);
	w.Seek(EndPos);
}

function FromXml_ST_SourceType(val) {
	var res = -1;
	if ("worksheet" === val) {
		res = c_oAscSourceType.Worksheet;
	} else if ("external" === val) {
		res = c_oAscSourceType.External;
	} else if ("consolidation" === val) {
		res = c_oAscSourceType.Consolidation;
	} else if ("scenario" === val) {
		res = c_oAscSourceType.Scenario;
	}
	return res;
}
function ToXml_ST_SourceType(val) {
	var res = "";
	if (c_oAscSourceType.Worksheet === val) {
		res = "worksheet";
	} else if (c_oAscSourceType.External === val) {
		res = "external";
	} else if (c_oAscSourceType.Consolidation === val) {
		res = "consolidation";
	} else if (c_oAscSourceType.Scenario === val) {
		res = "scenario";
	}
	return res;
}

function FromXml_ST_Axis(val) {
	var res = -1;
	if ("axisRow" === val) {
		res = c_oAscAxis.AxisRow;
	} else if ("axisCol" === val) {
		res = c_oAscAxis.AxisCol;
	} else if ("axisPage" === val) {
		res = c_oAscAxis.AxisPage;
	} else if ("axisValues" === val) {
		res = c_oAscAxis.AxisValues;
	}
	return res;
}
function ToXml_ST_Axis(val) {
	var res = "";
	if (c_oAscAxis.AxisRow === val) {
		res = "axisRow";
	} else if (c_oAscAxis.AxisCol === val) {
		res = "axisCol";
	} else if (c_oAscAxis.AxisPage === val) {
		res = "axisPage";
	} else if (c_oAscAxis.AxisValues === val) {
		res = "axisValues";
	}
	return res;
}

function FromXml_ST_FieldSortType(val) {
	var res = -1;
	if ("manual" === val) {
		res = c_oAscFieldSortType.Manual;
	} else if ("ascending" === val) {
		res = c_oAscFieldSortType.Ascending;
	} else if ("descending" === val) {
		res = c_oAscFieldSortType.Descending;
	}
	return res;
}
function ToXml_ST_FieldSortType(val) {
	var res = "";
	if (c_oAscFieldSortType.Manual === val) {
		res = "manual";
	} else if (c_oAscFieldSortType.Ascending === val) {
		res = "ascending";
	} else if (c_oAscFieldSortType.Descending === val) {
		res = "descending";
	}
	return res;
}

function FromXml_ST_ItemType(val) {
	var res = -1;
	if ("data" === val) {
		res = Asc.c_oAscItemType.Data;
	} else if ("default" === val) {
		res = Asc.c_oAscItemType.Default;
	} else if ("sum" === val) {
		res = Asc.c_oAscItemType.Sum;
	} else if ("countA" === val) {
		res = Asc.c_oAscItemType.CountA;
	} else if ("avg" === val) {
		res = Asc.c_oAscItemType.Avg;
	} else if ("max" === val) {
		res = Asc.c_oAscItemType.Max;
	} else if ("min" === val) {
		res = Asc.c_oAscItemType.Min;
	} else if ("product" === val) {
		res = Asc.c_oAscItemType.Product;
	} else if ("count" === val) {
		res = Asc.c_oAscItemType.Count;
	} else if ("stdDev" === val) {
		res = Asc.c_oAscItemType.StdDev;
	} else if ("stdDevP" === val) {
		res = Asc.c_oAscItemType.StdDevP;
	} else if ("var" === val) {
		res = Asc.c_oAscItemType.Var;
	} else if ("varP" === val) {
		res = Asc.c_oAscItemType.VarP;
	} else if ("grand" === val) {
		res = Asc.c_oAscItemType.Grand;
	} else if ("blank" === val) {
		res = Asc.c_oAscItemType.Blank;
	}
	return res;
}
function ToXml_ST_ItemType(val) {
	var res = "";
	if (Asc.c_oAscItemType.Data === val) {
		res = "data";
	} else if (Asc.c_oAscItemType.Default === val) {
		res = "default";
	} else if (Asc.c_oAscItemType.Sum === val) {
		res = "sum";
	} else if (Asc.c_oAscItemType.CountA === val) {
		res = "countA";
	} else if (Asc.c_oAscItemType.Avg === val) {
		res = "avg";
	} else if (Asc.c_oAscItemType.Max === val) {
		res = "max";
	} else if (Asc.c_oAscItemType.Min === val) {
		res = "min";
	} else if (Asc.c_oAscItemType.Product === val) {
		res = "product";
	} else if (Asc.c_oAscItemType.Count === val) {
		res = "count";
	} else if (Asc.c_oAscItemType.StdDev === val) {
		res = "stdDev";
	} else if (Asc.c_oAscItemType.StdDevP === val) {
		res = "stdDevP";
	} else if (Asc.c_oAscItemType.Var === val) {
		res = "var";
	} else if (Asc.c_oAscItemType.VarP === val) {
		res = "varP";
	} else if (Asc.c_oAscItemType.Grand === val) {
		res = "grand";
	} else if (Asc.c_oAscItemType.Blank === val) {
		res = "blank";
	}
	return res;
}

function ToName_ST_ItemType(val) {
	var res = '';
	if (Asc.c_oAscItemType.Default === val) {
		res += 'Total';
	} else if (Asc.c_oAscItemType.Avg === val) {
		res += 'Average';
	} else if (Asc.c_oAscItemType.Count === val) {
		res += 'Count';
	} else if (Asc.c_oAscItemType.CountA === val) {
		res += 'Count';
	} else if (Asc.c_oAscItemType.Max === val) {
		res += 'Max';
	} else if (Asc.c_oAscItemType.Min === val) {
		res += 'Min';
	} else if (Asc.c_oAscItemType.Product === val) {
		res += 'Product';
	} else if (Asc.c_oAscItemType.StdDev === val) {
		res += 'StdDev';
	} else if (Asc.c_oAscItemType.StdDevP === val) {
		res += 'StdDevp';
	} else if (Asc.c_oAscItemType.Sum === val) {
		res += 'Sum';
	} else if (Asc.c_oAscItemType.Var === val) {
		res += 'Var';
	} else if (Asc.c_oAscItemType.VarP === val) {
		res += 'Varp';
	} else if (Asc.c_oAscItemType.Data === val) {
		res += 'Data';
	} else if (Asc.c_oAscItemType.Grand === val) {
		res += 'Total';
	} else if (Asc.c_oAscItemType.Blank === val) {
		res += 'Blank';
	}
	return res;
}

function FromXml_ST_DataConsolidateFunction(val) {
	var res = -1;
	if ("average" === val) {
		res = c_oAscDataConsolidateFunction.Average;
	} else if ("count" === val) {
		res = c_oAscDataConsolidateFunction.Count;
	} else if ("countNums" === val) {
		res = c_oAscDataConsolidateFunction.CountNums;
	} else if ("max" === val) {
		res = c_oAscDataConsolidateFunction.Max;
	} else if ("min" === val) {
		res = c_oAscDataConsolidateFunction.Min;
	} else if ("product" === val) {
		res = c_oAscDataConsolidateFunction.Product;
	} else if ("stdDev" === val) {
		res = c_oAscDataConsolidateFunction.StdDev;
	} else if ("stdDevp" === val) {
		res = c_oAscDataConsolidateFunction.StdDevp;
	} else if ("sum" === val) {
		res = c_oAscDataConsolidateFunction.Sum;
	} else if ("var" === val) {
		res = c_oAscDataConsolidateFunction.Var;
	} else if ("varp" === val) {
		res = c_oAscDataConsolidateFunction.Varp;
	}
	return res;
}
function ToXml_ST_DataConsolidateFunction(val) {
	var res = "";
	if (c_oAscDataConsolidateFunction.Average === val) {
		res = "average";
	} else if (c_oAscDataConsolidateFunction.Count === val) {
		res = "count";
	} else if (c_oAscDataConsolidateFunction.CountNums === val) {
		res = "countNums";
	} else if (c_oAscDataConsolidateFunction.Max === val) {
		res = "max";
	} else if (c_oAscDataConsolidateFunction.Min === val) {
		res = "min";
	} else if (c_oAscDataConsolidateFunction.Product === val) {
		res = "product";
	} else if (c_oAscDataConsolidateFunction.StdDev === val) {
		res = "stdDev";
	} else if (c_oAscDataConsolidateFunction.StdDevp === val) {
		res = "stdDevp";
	} else if (c_oAscDataConsolidateFunction.Sum === val) {
		res = "sum";
	} else if (c_oAscDataConsolidateFunction.Var === val) {
		res = "var";
	} else if (c_oAscDataConsolidateFunction.Varp === val) {
		res = "varp";
	}
	return res;
}

function ToName_ST_DataConsolidateFunction(val) {
	var res = "";
	if (c_oAscDataConsolidateFunction.Average === val) {
		res = "Average";
	} else if (c_oAscDataConsolidateFunction.Count === val) {
		res = "Count";
	} else if (c_oAscDataConsolidateFunction.CountNums === val) {
		res = "Count";
	} else if (c_oAscDataConsolidateFunction.Max === val) {
		res = "Max";
	} else if (c_oAscDataConsolidateFunction.Min === val) {
		res = "Min";
	} else if (c_oAscDataConsolidateFunction.Product === val) {
		res = "Product";
	} else if (c_oAscDataConsolidateFunction.StdDev === val) {
		res = "StdDev";
	} else if (c_oAscDataConsolidateFunction.StdDevp === val) {
		res = "StdDevp";
	} else if (c_oAscDataConsolidateFunction.Sum === val) {
		res = "Sum";
	} else if (c_oAscDataConsolidateFunction.Var === val) {
		res = "Var";
	} else if (c_oAscDataConsolidateFunction.Varp === val) {
		res = "Varp";
	}
	return res;
}

function FromXml_ST_ShowDataAs(val) {
	var res = -1;
	if ("normal" === val) {
		res = c_oAscShowDataAs.Normal;
	} else if ("difference" === val) {
		res = c_oAscShowDataAs.Difference;
	} else if ("percent" === val) {
		res = c_oAscShowDataAs.Percent;
	} else if ("percentDiff" === val) {
		res = c_oAscShowDataAs.PercentDiff;
	} else if ("runTotal" === val) {
		res = c_oAscShowDataAs.RunTotal;
	} else if ("percentOfRow" === val) {
		res = c_oAscShowDataAs.PercentOfRow;
	} else if ("percentOfCol" === val) {
		res = c_oAscShowDataAs.PercentOfCol;
	} else if ("percentOfTotal" === val) {
		res = c_oAscShowDataAs.PercentOfTotal;
	} else if ("index" === val) {
		res = c_oAscShowDataAs.Index;
	}
	return res;
}
function ToXml_ST_ShowDataAs(val) {
	var res = "";
	if (c_oAscShowDataAs.Normal === val) {
		res = "normal";
	} else if (c_oAscShowDataAs.Difference === val) {
		res = "difference";
	} else if (c_oAscShowDataAs.Percent === val) {
		res = "percent";
	} else if (c_oAscShowDataAs.PercentDiff === val) {
		res = "percentDiff";
	} else if (c_oAscShowDataAs.RunTotal === val) {
		res = "runTotal";
	} else if (c_oAscShowDataAs.PercentOfRow === val) {
		res = "percentOfRow";
	} else if (c_oAscShowDataAs.PercentOfCol === val) {
		res = "percentOfCol";
	} else if (c_oAscShowDataAs.PercentOfTotal === val) {
		res = "percentOfTotal";
	} else if (c_oAscShowDataAs.Index === val) {
		res = "index";
	}
	return res;
}

function FromXml_ST_FormatAction(val) {
	var res = -1;
	if ("blank" === val) {
		res = c_oAscFormatAction.Blank;
	} else if ("formatting" === val) {
		res = c_oAscFormatAction.Formatting;
	} else if ("drill" === val) {
		res = c_oAscFormatAction.Drill;
	} else if ("formula" === val) {
		res = c_oAscFormatAction.Formula;
	}
	return res;
}
function ToXml_ST_FormatAction(val) {
	var res = "";
	if (c_oAscFormatAction.Blank === val) {
		res = "blank";
	} else if (c_oAscFormatAction.Formatting === val) {
		res = "formatting";
	} else if (c_oAscFormatAction.Drill === val) {
		res = "drill";
	} else if (c_oAscFormatAction.Formula === val) {
		res = "formula";
	}
	return res;
}

function FromXml_ST_Scope(val) {
	var res = -1;
	if ("selection" === val) {
		res = c_oAscScope.Selection;
	} else if ("data" === val) {
		res = c_oAscScope.Data;
	} else if ("field" === val) {
		res = c_oAscScope.Field;
	}
	return res;
}
function ToXml_ST_Scope(val) {
	var res = "";
	if (c_oAscScope.Selection === val) {
		res = "selection";
	} else if (c_oAscScope.Data === val) {
		res = "data";
	} else if (c_oAscScope.Field === val) {
		res = "field";
	}
	return res;
}

function FromXml_ST_Type(val) {
	var res = -1;
	if ("none" === val) {
		res = c_oAscType.None;
	} else if ("all" === val) {
		res = c_oAscType.All;
	} else if ("row" === val) {
		res = c_oAscType.Row;
	} else if ("column" === val) {
		res = c_oAscType.Column;
	}
	return res;
}
function ToXml_ST_Type(val) {
	var res = "";
	if (c_oAscType.None === val) {
		res = "none";
	} else if (c_oAscType.All === val) {
		res = "all";
	} else if (c_oAscType.Row === val) {
		res = "row";
	} else if (c_oAscType.Column === val) {
		res = "column";
	}
	return res;
}

function FromXml_ST_PivotFilterType(val) {
	var res = -1;
	if ("unknown" === val) {
		res = c_oAscPivotFilterType.Unknown;
	} else if ("count" === val) {
		res = c_oAscPivotFilterType.Count;
	} else if ("percent" === val) {
		res = c_oAscPivotFilterType.Percent;
	} else if ("sum" === val) {
		res = c_oAscPivotFilterType.Sum;
	} else if ("captionEqual" === val) {
		res = c_oAscPivotFilterType.CaptionEqual;
	} else if ("captionNotEqual" === val) {
		res = c_oAscPivotFilterType.CaptionNotEqual;
	} else if ("captionBeginsWith" === val) {
		res = c_oAscPivotFilterType.CaptionBeginsWith;
	} else if ("captionNotBeginsWith" === val) {
		res = c_oAscPivotFilterType.CaptionNotBeginsWith;
	} else if ("captionEndsWith" === val) {
		res = c_oAscPivotFilterType.CaptionEndsWith;
	} else if ("captionNotEndsWith" === val) {
		res = c_oAscPivotFilterType.CaptionNotEndsWith;
	} else if ("captionContains" === val) {
		res = c_oAscPivotFilterType.CaptionContains;
	} else if ("captionNotContains" === val) {
		res = c_oAscPivotFilterType.CaptionNotContains;
	} else if ("captionGreaterThan" === val) {
		res = c_oAscPivotFilterType.CaptionGreaterThan;
	} else if ("captionGreaterThanOrEqual" === val) {
		res = c_oAscPivotFilterType.CaptionGreaterThanOrEqual;
	} else if ("captionLessThan" === val) {
		res = c_oAscPivotFilterType.CaptionLessThan;
	} else if ("captionLessThanOrEqual" === val) {
		res = c_oAscPivotFilterType.CaptionLessThanOrEqual;
	} else if ("captionBetween" === val) {
		res = c_oAscPivotFilterType.CaptionBetween;
	} else if ("captionNotBetween" === val) {
		res = c_oAscPivotFilterType.CaptionNotBetween;
	} else if ("valueEqual" === val) {
		res = c_oAscPivotFilterType.ValueEqual;
	} else if ("valueNotEqual" === val) {
		res = c_oAscPivotFilterType.ValueNotEqual;
	} else if ("valueGreaterThan" === val) {
		res = c_oAscPivotFilterType.ValueGreaterThan;
	} else if ("valueGreaterThanOrEqual" === val) {
		res = c_oAscPivotFilterType.ValueGreaterThanOrEqual;
	} else if ("valueLessThan" === val) {
		res = c_oAscPivotFilterType.ValueLessThan;
	} else if ("valueLessThanOrEqual" === val) {
		res = c_oAscPivotFilterType.ValueLessThanOrEqual;
	} else if ("valueBetween" === val) {
		res = c_oAscPivotFilterType.ValueBetween;
	} else if ("valueNotBetween" === val) {
		res = c_oAscPivotFilterType.ValueNotBetween;
	} else if ("dateEqual" === val) {
		res = c_oAscPivotFilterType.DateEqual;
	} else if ("dateNotEqual" === val) {
		res = c_oAscPivotFilterType.DateNotEqual;
	} else if ("dateOlderThan" === val) {
		res = c_oAscPivotFilterType.DateOlderThan;
	} else if ("dateOlderThanOrEqual" === val) {
		res = c_oAscPivotFilterType.DateOlderThanOrEqual;
	} else if ("dateNewerThan" === val) {
		res = c_oAscPivotFilterType.DateNewerThan;
	} else if ("dateNewerThanOrEqual" === val) {
		res = c_oAscPivotFilterType.DateNewerThanOrEqual;
	} else if ("dateBetween" === val) {
		res = c_oAscPivotFilterType.DateBetween;
	} else if ("dateNotBetween" === val) {
		res = c_oAscPivotFilterType.DateNotBetween;
	} else if ("tomorrow" === val) {
		res = c_oAscPivotFilterType.Tomorrow;
	} else if ("today" === val) {
		res = c_oAscPivotFilterType.Today;
	} else if ("yesterday" === val) {
		res = c_oAscPivotFilterType.Yesterday;
	} else if ("nextWeek" === val) {
		res = c_oAscPivotFilterType.NextWeek;
	} else if ("thisWeek" === val) {
		res = c_oAscPivotFilterType.ThisWeek;
	} else if ("lastWeek" === val) {
		res = c_oAscPivotFilterType.LastWeek;
	} else if ("nextMonth" === val) {
		res = c_oAscPivotFilterType.NextMonth;
	} else if ("thisMonth" === val) {
		res = c_oAscPivotFilterType.ThisMonth;
	} else if ("lastMonth" === val) {
		res = c_oAscPivotFilterType.LastMonth;
	} else if ("nextQuarter" === val) {
		res = c_oAscPivotFilterType.NextQuarter;
	} else if ("thisQuarter" === val) {
		res = c_oAscPivotFilterType.ThisQuarter;
	} else if ("lastQuarter" === val) {
		res = c_oAscPivotFilterType.LastQuarter;
	} else if ("nextYear" === val) {
		res = c_oAscPivotFilterType.NextYear;
	} else if ("thisYear" === val) {
		res = c_oAscPivotFilterType.ThisYear;
	} else if ("lastYear" === val) {
		res = c_oAscPivotFilterType.LastYear;
	} else if ("yearToDate" === val) {
		res = c_oAscPivotFilterType.YearToDate;
	} else if ("Q1" === val) {
		res = c_oAscPivotFilterType.Q1;
	} else if ("Q2" === val) {
		res = c_oAscPivotFilterType.Q2;
	} else if ("Q3" === val) {
		res = c_oAscPivotFilterType.Q3;
	} else if ("Q4" === val) {
		res = c_oAscPivotFilterType.Q4;
	} else if ("M1" === val) {
		res = c_oAscPivotFilterType.M1;
	} else if ("M2" === val) {
		res = c_oAscPivotFilterType.M2;
	} else if ("M3" === val) {
		res = c_oAscPivotFilterType.M3;
	} else if ("M4" === val) {
		res = c_oAscPivotFilterType.M4;
	} else if ("M5" === val) {
		res = c_oAscPivotFilterType.M5;
	} else if ("M6" === val) {
		res = c_oAscPivotFilterType.M6;
	} else if ("M7" === val) {
		res = c_oAscPivotFilterType.M7;
	} else if ("M8" === val) {
		res = c_oAscPivotFilterType.M8;
	} else if ("M9" === val) {
		res = c_oAscPivotFilterType.M9;
	} else if ("M10" === val) {
		res = c_oAscPivotFilterType.M10;
	} else if ("M11" === val) {
		res = c_oAscPivotFilterType.M11;
	} else if ("M12" === val) {
		res = c_oAscPivotFilterType.M12;
	}
	return res;
}
function ToXml_ST_PivotFilterType(val) {
	var res = "";
	if (c_oAscPivotFilterType.Unknown === val) {
		res = "unknown";
	} else if (c_oAscPivotFilterType.Count === val) {
		res = "count";
	} else if (c_oAscPivotFilterType.Percent === val) {
		res = "percent";
	} else if (c_oAscPivotFilterType.Sum === val) {
		res = "sum";
	} else if (c_oAscPivotFilterType.CaptionEqual === val) {
		res = "captionEqual";
	} else if (c_oAscPivotFilterType.CaptionNotEqual === val) {
		res = "captionNotEqual";
	} else if (c_oAscPivotFilterType.CaptionBeginsWith === val) {
		res = "captionBeginsWith";
	} else if (c_oAscPivotFilterType.CaptionNotBeginsWith === val) {
		res = "captionNotBeginsWith";
	} else if (c_oAscPivotFilterType.CaptionEndsWith === val) {
		res = "captionEndsWith";
	} else if (c_oAscPivotFilterType.CaptionNotEndsWith === val) {
		res = "captionNotEndsWith";
	} else if (c_oAscPivotFilterType.CaptionContains === val) {
		res = "captionContains";
	} else if (c_oAscPivotFilterType.CaptionNotContains === val) {
		res = "captionNotContains";
	} else if (c_oAscPivotFilterType.CaptionGreaterThan === val) {
		res = "captionGreaterThan";
	} else if (c_oAscPivotFilterType.CaptionGreaterThanOrEqual === val) {
		res = "captionGreaterThanOrEqual";
	} else if (c_oAscPivotFilterType.CaptionLessThan === val) {
		res = "captionLessThan";
	} else if (c_oAscPivotFilterType.CaptionLessThanOrEqual === val) {
		res = "captionLessThanOrEqual";
	} else if (c_oAscPivotFilterType.CaptionBetween === val) {
		res = "captionBetween";
	} else if (c_oAscPivotFilterType.CaptionNotBetween === val) {
		res = "captionNotBetween";
	} else if (c_oAscPivotFilterType.ValueEqual === val) {
		res = "valueEqual";
	} else if (c_oAscPivotFilterType.ValueNotEqual === val) {
		res = "valueNotEqual";
	} else if (c_oAscPivotFilterType.ValueGreaterThan === val) {
		res = "valueGreaterThan";
	} else if (c_oAscPivotFilterType.ValueGreaterThanOrEqual === val) {
		res = "valueGreaterThanOrEqual";
	} else if (c_oAscPivotFilterType.ValueLessThan === val) {
		res = "valueLessThan";
	} else if (c_oAscPivotFilterType.ValueLessThanOrEqual === val) {
		res = "valueLessThanOrEqual";
	} else if (c_oAscPivotFilterType.ValueBetween === val) {
		res = "valueBetween";
	} else if (c_oAscPivotFilterType.ValueNotBetween === val) {
		res = "valueNotBetween";
	} else if (c_oAscPivotFilterType.DateEqual === val) {
		res = "dateEqual";
	} else if (c_oAscPivotFilterType.DateNotEqual === val) {
		res = "dateNotEqual";
	} else if (c_oAscPivotFilterType.DateOlderThan === val) {
		res = "dateOlderThan";
	} else if (c_oAscPivotFilterType.DateOlderThanOrEqual === val) {
		res = "dateOlderThanOrEqual";
	} else if (c_oAscPivotFilterType.DateNewerThan === val) {
		res = "dateNewerThan";
	} else if (c_oAscPivotFilterType.DateNewerThanOrEqual === val) {
		res = "dateNewerThanOrEqual";
	} else if (c_oAscPivotFilterType.DateBetween === val) {
		res = "dateBetween";
	} else if (c_oAscPivotFilterType.DateNotBetween === val) {
		res = "dateNotBetween";
	} else if (c_oAscPivotFilterType.Tomorrow === val) {
		res = "tomorrow";
	} else if (c_oAscPivotFilterType.Today === val) {
		res = "today";
	} else if (c_oAscPivotFilterType.Yesterday === val) {
		res = "yesterday";
	} else if (c_oAscPivotFilterType.NextWeek === val) {
		res = "nextWeek";
	} else if (c_oAscPivotFilterType.ThisWeek === val) {
		res = "thisWeek";
	} else if (c_oAscPivotFilterType.LastWeek === val) {
		res = "lastWeek";
	} else if (c_oAscPivotFilterType.NextMonth === val) {
		res = "nextMonth";
	} else if (c_oAscPivotFilterType.ThisMonth === val) {
		res = "thisMonth";
	} else if (c_oAscPivotFilterType.LastMonth === val) {
		res = "lastMonth";
	} else if (c_oAscPivotFilterType.NextQuarter === val) {
		res = "nextQuarter";
	} else if (c_oAscPivotFilterType.ThisQuarter === val) {
		res = "thisQuarter";
	} else if (c_oAscPivotFilterType.LastQuarter === val) {
		res = "lastQuarter";
	} else if (c_oAscPivotFilterType.NextYear === val) {
		res = "nextYear";
	} else if (c_oAscPivotFilterType.ThisYear === val) {
		res = "thisYear";
	} else if (c_oAscPivotFilterType.LastYear === val) {
		res = "lastYear";
	} else if (c_oAscPivotFilterType.YearToDate === val) {
		res = "yearToDate";
	} else if (c_oAscPivotFilterType.Q1 === val) {
		res = "Q1";
	} else if (c_oAscPivotFilterType.Q2 === val) {
		res = "Q2";
	} else if (c_oAscPivotFilterType.Q3 === val) {
		res = "Q3";
	} else if (c_oAscPivotFilterType.Q4 === val) {
		res = "Q4";
	} else if (c_oAscPivotFilterType.M1 === val) {
		res = "M1";
	} else if (c_oAscPivotFilterType.M2 === val) {
		res = "M2";
	} else if (c_oAscPivotFilterType.M3 === val) {
		res = "M3";
	} else if (c_oAscPivotFilterType.M4 === val) {
		res = "M4";
	} else if (c_oAscPivotFilterType.M5 === val) {
		res = "M5";
	} else if (c_oAscPivotFilterType.M6 === val) {
		res = "M6";
	} else if (c_oAscPivotFilterType.M7 === val) {
		res = "M7";
	} else if (c_oAscPivotFilterType.M8 === val) {
		res = "M8";
	} else if (c_oAscPivotFilterType.M9 === val) {
		res = "M9";
	} else if (c_oAscPivotFilterType.M10 === val) {
		res = "M10";
	} else if (c_oAscPivotFilterType.M11 === val) {
		res = "M11";
	} else if (c_oAscPivotFilterType.M12 === val) {
		res = "M12";
	}
	return res;
}

function FromXml_ST_SortType(val) {
	var res = -1;
	if ("none" === val) {
		res = c_oAscSortType.None;
	} else if ("ascending" === val) {
		res = c_oAscSortType.Ascending;
	} else if ("descending" === val) {
		res = c_oAscSortType.Descending;
	} else if ("ascendingAlpha" === val) {
		res = c_oAscSortType.AscendingAlpha;
	} else if ("descendingAlpha" === val) {
		res = c_oAscSortType.DescendingAlpha;
	} else if ("ascendingNatural" === val) {
		res = c_oAscSortType.AscendingNatural;
	} else if ("descendingNatural" === val) {
		res = c_oAscSortType.DescendingNatural;
	}
	return res;
}
function ToXml_ST_SortType(val) {
	var res = "";
	if (c_oAscSortType.None === val) {
		res = "none";
	} else if (c_oAscSortType.Ascending === val) {
		res = "ascending";
	} else if (c_oAscSortType.Descending === val) {
		res = "descending";
	} else if (c_oAscSortType.AscendingAlpha === val) {
		res = "ascendingAlpha";
	} else if (c_oAscSortType.DescendingAlpha === val) {
		res = "descendingAlpha";
	} else if (c_oAscSortType.AscendingNatural === val) {
		res = "ascendingNatural";
	} else if (c_oAscSortType.DescendingNatural === val) {
		res = "descendingNatural";
	}
	return res;
}

function FromXml_ST_PivotAreaType(val) {
	var res = -1;
	if ("none" === val) {
		res = c_oAscPivotAreaType.None;
	} else if ("normal" === val) {
		res = c_oAscPivotAreaType.Normal;
	} else if ("data" === val) {
		res = c_oAscPivotAreaType.Data;
	} else if ("all" === val) {
		res = c_oAscPivotAreaType.All;
	} else if ("origin" === val) {
		res = c_oAscPivotAreaType.Origin;
	} else if ("button" === val) {
		res = c_oAscPivotAreaType.Button;
	} else if ("topEnd" === val) {
		res = c_oAscPivotAreaType.TopEnd;
	}
	return res;
}
function ToXml_ST_PivotAreaType(val) {
	var res = "";
	if (c_oAscPivotAreaType.None === val) {
		res = "none";
	} else if (c_oAscPivotAreaType.Normal === val) {
		res = "normal";
	} else if (c_oAscPivotAreaType.Data === val) {
		res = "data";
	} else if (c_oAscPivotAreaType.All === val) {
		res = "all";
	} else if (c_oAscPivotAreaType.Origin === val) {
		res = "origin";
	} else if (c_oAscPivotAreaType.Button === val) {
		res = "button";
	} else if (c_oAscPivotAreaType.TopEnd === val) {
		res = "topEnd";
	}
	return res;
}

function FromXml_ST_GroupBy(val) {
	var res = -1;
	if ("range" === val) {
		res = c_oAscGroupBy.Range;
	} else if ("seconds" === val) {
		res = c_oAscGroupBy.Seconds;
	} else if ("minutes" === val) {
		res = c_oAscGroupBy.Minutes;
	} else if ("hours" === val) {
		res = c_oAscGroupBy.Hours;
	} else if ("days" === val) {
		res = c_oAscGroupBy.Days;
	} else if ("months" === val) {
		res = c_oAscGroupBy.Months;
	} else if ("quarters" === val) {
		res = c_oAscGroupBy.Quarters;
	} else if ("years" === val) {
		res = c_oAscGroupBy.Years;
	}
	return res;
}
function ToXml_ST_GroupBy(val) {
	var res = "";
	if (c_oAscGroupBy.Range === val) {
		res = "range";
	} else if (c_oAscGroupBy.Seconds === val) {
		res = "seconds";
	} else if (c_oAscGroupBy.Minutes === val) {
		res = "minutes";
	} else if (c_oAscGroupBy.Hours === val) {
		res = "hours";
	} else if (c_oAscGroupBy.Days === val) {
		res = "days";
	} else if (c_oAscGroupBy.Months === val) {
		res = "months";
	} else if (c_oAscGroupBy.Quarters === val) {
		res = "quarters";
	} else if (c_oAscGroupBy.Years === val) {
		res = "years";
	}
	return res;
}

function FromXml_ST_SortMethod(val) {
	var res = -1;
	if ("stroke" === val) {
		res = c_oAscSortMethod.Stroke;
	} else if ("pinYin" === val) {
		res = c_oAscSortMethod.PinYin;
	} else if ("none" === val) {
		res = c_oAscSortMethod.None;
	}
	return res;
}
function ToXml_ST_SortMethod(val) {
	var res = "";
	if (c_oAscSortMethod.Stroke === val) {
		res = "stroke";
	} else if (c_oAscSortMethod.PinYin === val) {
		res = "pinYin";
	} else if (c_oAscSortMethod.None === val) {
		res = "none";
	}
	return res;
}

function FromXml_ST_DynamicFilterType(val) {
	var res = -1;
	if ("null" === val) {
		res = c_oAscDynamicFilterType.Null;
	} else if ("aboveAverage" === val) {
		res = c_oAscDynamicFilterType.AboveAverage;
	} else if ("belowAverage" === val) {
		res = c_oAscDynamicFilterType.BelowAverage;
	} else if ("tomorrow" === val) {
		res = c_oAscDynamicFilterType.Tomorrow;
	} else if ("today" === val) {
		res = c_oAscDynamicFilterType.Today;
	} else if ("yesterday" === val) {
		res = c_oAscDynamicFilterType.Yesterday;
	} else if ("nextWeek" === val) {
		res = c_oAscDynamicFilterType.NextWeek;
	} else if ("thisWeek" === val) {
		res = c_oAscDynamicFilterType.ThisWeek;
	} else if ("lastWeek" === val) {
		res = c_oAscDynamicFilterType.LastWeek;
	} else if ("nextMonth" === val) {
		res = c_oAscDynamicFilterType.NextMonth;
	} else if ("thisMonth" === val) {
		res = c_oAscDynamicFilterType.ThisMonth;
	} else if ("lastMonth" === val) {
		res = c_oAscDynamicFilterType.LastMonth;
	} else if ("nextQuarter" === val) {
		res = c_oAscDynamicFilterType.NextQuarter;
	} else if ("thisQuarter" === val) {
		res = c_oAscDynamicFilterType.ThisQuarter;
	} else if ("lastQuarter" === val) {
		res = c_oAscDynamicFilterType.LastQuarter;
	} else if ("nextYear" === val) {
		res = c_oAscDynamicFilterType.NextYear;
	} else if ("thisYear" === val) {
		res = c_oAscDynamicFilterType.ThisYear;
	} else if ("lastYear" === val) {
		res = c_oAscDynamicFilterType.LastYear;
	} else if ("yearToDate" === val) {
		res = c_oAscDynamicFilterType.YearToDate;
	} else if ("Q1" === val) {
		res = c_oAscDynamicFilterType.Q1;
	} else if ("Q2" === val) {
		res = c_oAscDynamicFilterType.Q2;
	} else if ("Q3" === val) {
		res = c_oAscDynamicFilterType.Q3;
	} else if ("Q4" === val) {
		res = c_oAscDynamicFilterType.Q4;
	} else if ("M1" === val) {
		res = c_oAscDynamicFilterType.M1;
	} else if ("M2" === val) {
		res = c_oAscDynamicFilterType.M2;
	} else if ("M3" === val) {
		res = c_oAscDynamicFilterType.M3;
	} else if ("M4" === val) {
		res = c_oAscDynamicFilterType.M4;
	} else if ("M5" === val) {
		res = c_oAscDynamicFilterType.M5;
	} else if ("M6" === val) {
		res = c_oAscDynamicFilterType.M6;
	} else if ("M7" === val) {
		res = c_oAscDynamicFilterType.M7;
	} else if ("M8" === val) {
		res = c_oAscDynamicFilterType.M8;
	} else if ("M9" === val) {
		res = c_oAscDynamicFilterType.M9;
	} else if ("M10" === val) {
		res = c_oAscDynamicFilterType.M10;
	} else if ("M11" === val) {
		res = c_oAscDynamicFilterType.M11;
	} else if ("M12" === val) {
		res = c_oAscDynamicFilterType.M12;
	}
	return res;
}
function ToXml_ST_DynamicFilterType(val) {
	var res = "";
	if (c_oAscDynamicFilterType.Null === val) {
		res = "null";
	} else if (c_oAscDynamicFilterType.AboveAverage === val) {
		res = "aboveAverage";
	} else if (c_oAscDynamicFilterType.BelowAverage === val) {
		res = "belowAverage";
	} else if (c_oAscDynamicFilterType.Tomorrow === val) {
		res = "tomorrow";
	} else if (c_oAscDynamicFilterType.Today === val) {
		res = "today";
	} else if (c_oAscDynamicFilterType.Yesterday === val) {
		res = "yesterday";
	} else if (c_oAscDynamicFilterType.NextWeek === val) {
		res = "nextWeek";
	} else if (c_oAscDynamicFilterType.ThisWeek === val) {
		res = "thisWeek";
	} else if (c_oAscDynamicFilterType.LastWeek === val) {
		res = "lastWeek";
	} else if (c_oAscDynamicFilterType.NextMonth === val) {
		res = "nextMonth";
	} else if (c_oAscDynamicFilterType.ThisMonth === val) {
		res = "thisMonth";
	} else if (c_oAscDynamicFilterType.LastMonth === val) {
		res = "lastMonth";
	} else if (c_oAscDynamicFilterType.NextQuarter === val) {
		res = "nextQuarter";
	} else if (c_oAscDynamicFilterType.ThisQuarter === val) {
		res = "thisQuarter";
	} else if (c_oAscDynamicFilterType.LastQuarter === val) {
		res = "lastQuarter";
	} else if (c_oAscDynamicFilterType.NextYear === val) {
		res = "nextYear";
	} else if (c_oAscDynamicFilterType.ThisYear === val) {
		res = "thisYear";
	} else if (c_oAscDynamicFilterType.LastYear === val) {
		res = "lastYear";
	} else if (c_oAscDynamicFilterType.YearToDate === val) {
		res = "yearToDate";
	} else if (c_oAscDynamicFilterType.Q1 === val) {
		res = "Q1";
	} else if (c_oAscDynamicFilterType.Q2 === val) {
		res = "Q2";
	} else if (c_oAscDynamicFilterType.Q3 === val) {
		res = "Q3";
	} else if (c_oAscDynamicFilterType.Q4 === val) {
		res = "Q4";
	} else if (c_oAscDynamicFilterType.M1 === val) {
		res = "M1";
	} else if (c_oAscDynamicFilterType.M2 === val) {
		res = "M2";
	} else if (c_oAscDynamicFilterType.M3 === val) {
		res = "M3";
	} else if (c_oAscDynamicFilterType.M4 === val) {
		res = "M4";
	} else if (c_oAscDynamicFilterType.M5 === val) {
		res = "M5";
	} else if (c_oAscDynamicFilterType.M6 === val) {
		res = "M6";
	} else if (c_oAscDynamicFilterType.M7 === val) {
		res = "M7";
	} else if (c_oAscDynamicFilterType.M8 === val) {
		res = "M8";
	} else if (c_oAscDynamicFilterType.M9 === val) {
		res = "M9";
	} else if (c_oAscDynamicFilterType.M10 === val) {
		res = "M10";
	} else if (c_oAscDynamicFilterType.M11 === val) {
		res = "M11";
	} else if (c_oAscDynamicFilterType.M12 === val) {
		res = "M12";
	}
	return res;
}

function FromXml_ST_CalendarType(val) {
	var res = -1;
	if ("gregorian" === val) {
		res = c_oAscCalendarType.Gregorian;
	} else if ("gregorianUs" === val) {
		res = c_oAscCalendarType.GregorianUs;
	} else if ("gregorianMeFrench" === val) {
		res = c_oAscCalendarType.GregorianMeFrench;
	} else if ("gregorianArabic" === val) {
		res = c_oAscCalendarType.GregorianArabic;
	} else if ("hijri" === val) {
		res = c_oAscCalendarType.Hijri;
	} else if ("hebrew" === val) {
		res = c_oAscCalendarType.Hebrew;
	} else if ("taiwan" === val) {
		res = c_oAscCalendarType.Taiwan;
	} else if ("japan" === val) {
		res = c_oAscCalendarType.Japan;
	} else if ("thai" === val) {
		res = c_oAscCalendarType.Thai;
	} else if ("korea" === val) {
		res = c_oAscCalendarType.Korea;
	} else if ("saka" === val) {
		res = c_oAscCalendarType.Saka;
	} else if ("gregorianXlitEnglish" === val) {
		res = c_oAscCalendarType.GregorianXlitEnglish;
	} else if ("gregorianXlitFrench" === val) {
		res = c_oAscCalendarType.GregorianXlitFrench;
	} else if ("none" === val) {
		res = c_oAscCalendarType.None;
	}
	return res;
}
function ToXml_ST_CalendarType(val) {
	var res = "";
	if (c_oAscCalendarType.Gregorian === val) {
		res = "gregorian";
	} else if (c_oAscCalendarType.GregorianUs === val) {
		res = "gregorianUs";
	} else if (c_oAscCalendarType.GregorianMeFrench === val) {
		res = "gregorianMeFrench";
	} else if (c_oAscCalendarType.GregorianArabic === val) {
		res = "gregorianArabic";
	} else if (c_oAscCalendarType.Hijri === val) {
		res = "hijri";
	} else if (c_oAscCalendarType.Hebrew === val) {
		res = "hebrew";
	} else if (c_oAscCalendarType.Taiwan === val) {
		res = "taiwan";
	} else if (c_oAscCalendarType.Japan === val) {
		res = "japan";
	} else if (c_oAscCalendarType.Thai === val) {
		res = "thai";
	} else if (c_oAscCalendarType.Korea === val) {
		res = "korea";
	} else if (c_oAscCalendarType.Saka === val) {
		res = "saka";
	} else if (c_oAscCalendarType.GregorianXlitEnglish === val) {
		res = "gregorianXlitEnglish";
	} else if (c_oAscCalendarType.GregorianXlitFrench === val) {
		res = "gregorianXlitFrench";
	} else if (c_oAscCalendarType.None === val) {
		res = "none";
	}
	return res;
}

function FromXml_ST_IconSetType(val) {
	var res = -1;
	if ("3Arrows" === val) {
		res = c_oAscIconSetType.ThreeArrows;
	} else if ("3ArrowsGray" === val) {
		res = c_oAscIconSetType.ThreeArrowsGray;
	} else if ("3Flags" === val) {
		res = c_oAscIconSetType.ThreeFlags;
	} else if ("3TrafficLights1" === val) {
		res = c_oAscIconSetType.ThreeTrafficLights1;
	} else if ("3TrafficLights2" === val) {
		res = c_oAscIconSetType.ThreeTrafficLights2;
	} else if ("3Signs" === val) {
		res = c_oAscIconSetType.ThreeSigns;
	} else if ("3Symbols" === val) {
		res = c_oAscIconSetType.ThreeSymbols;
	} else if ("3Symbols2" === val) {
		res = c_oAscIconSetType.ThreeSymbols2;
	} else if ("4Arrows" === val) {
		res = c_oAscIconSetType.FourArrows;
	} else if ("4ArrowsGray" === val) {
		res = c_oAscIconSetType.FourArrowsGray;
	} else if ("4RedToBlack" === val) {
		res = c_oAscIconSetType.FourRedToBlack;
	} else if ("4Rating" === val) {
		res = c_oAscIconSetType.FourRating;
	} else if ("4TrafficLights" === val) {
		res = c_oAscIconSetType.FourTrafficLights;
	} else if ("5Arrows" === val) {
		res = c_oAscIconSetType.FiveArrows;
	} else if ("5ArrowsGray" === val) {
		res = c_oAscIconSetType.FiveArrowsGray;
	} else if ("5Rating" === val) {
		res = c_oAscIconSetType.FiveRating;
	} else if ("5Quarters" === val) {
		res = c_oAscIconSetType.FiveQuarters;
	}
	return res;
}
function ToXml_ST_IconSetType(val) {
	var res = "";
	if (c_oAscIconSetType.ThreeArrows === val) {
		res = "3Arrows";
	} else if (c_oAscIconSetType.ThreeArrowsGray === val) {
		res = "3ArrowsGray";
	} else if (c_oAscIconSetType.ThreeFlags === val) {
		res = "3Flags";
	} else if (c_oAscIconSetType.ThreeTrafficLights1 === val) {
		res = "3TrafficLights1";
	} else if (c_oAscIconSetType.ThreeTrafficLights2 === val) {
		res = "3TrafficLights2";
	} else if (c_oAscIconSetType.ThreeSigns === val) {
		res = "3Signs";
	} else if (c_oAscIconSetType.ThreeSymbols === val) {
		res = "3Symbols";
	} else if (c_oAscIconSetType.ThreeSymbols2 === val) {
		res = "3Symbols2";
	} else if (c_oAscIconSetType.FourArrows === val) {
		res = "4Arrows";
	} else if (c_oAscIconSetType.FourArrowsGray === val) {
		res = "4ArrowsGray";
	} else if (c_oAscIconSetType.FourRedToBlack === val) {
		res = "4RedToBlack";
	} else if (c_oAscIconSetType.FourRating === val) {
		res = "4Rating";
	} else if (c_oAscIconSetType.FourTrafficLights === val) {
		res = "4TrafficLights";
	} else if (c_oAscIconSetType.FiveArrows === val) {
		res = "5Arrows";
	} else if (c_oAscIconSetType.FiveArrowsGray === val) {
		res = "5ArrowsGray";
	} else if (c_oAscIconSetType.FiveRating === val) {
		res = "5Rating";
	} else if (c_oAscIconSetType.FiveQuarters === val) {
		res = "5Quarters";
	}
	return res;
}

function FromXml_ST_SortBy(val) {
	var res = -1;
	if ("value" === val) {
		res = c_oAscSortBy.Value;
	} else if ("cellColor" === val) {
		res = c_oAscSortBy.CellColor;
	} else if ("fontColor" === val) {
		res = c_oAscSortBy.FontColor;
	} else if ("icon" === val) {
		res = c_oAscSortBy.Icon;
	}
	return res;
}
function ToXml_ST_SortBy(val) {
	var res = "";
	if (c_oAscSortBy.Value === val) {
		res = "value";
	} else if (c_oAscSortBy.CellColor === val) {
		res = "cellColor";
	} else if (c_oAscSortBy.FontColor === val) {
		res = "fontColor";
	} else if (c_oAscSortBy.Icon === val) {
		res = "icon";
	}
	return res;
}

function FromXml_ST_FilterOperator(val) {
	var res = -1;
	if ("equal" === val) {
		res = c_oAscFilterOperator.Equal;
	} else if ("lessThan" === val) {
		res = c_oAscFilterOperator.LessThan;
	} else if ("lessThanOrEqual" === val) {
		res = c_oAscFilterOperator.LessThanOrEqual;
	} else if ("notEqual" === val) {
		res = c_oAscFilterOperator.NotEqual;
	} else if ("greaterThanOrEqual" === val) {
		res = c_oAscFilterOperator.GreaterThanOrEqual;
	} else if ("greaterThan" === val) {
		res = c_oAscFilterOperator.GreaterThan;
	}
	return res;
}
function ToXml_ST_FilterOperator(val) {
	var res = "";
	if (c_oAscFilterOperator.Equal === val) {
		res = "equal";
	} else if (c_oAscFilterOperator.LessThan === val) {
		res = "lessThan";
	} else if (c_oAscFilterOperator.LessThanOrEqual === val) {
		res = "lessThanOrEqual";
	} else if (c_oAscFilterOperator.NotEqual === val) {
		res = "notEqual";
	} else if (c_oAscFilterOperator.GreaterThanOrEqual === val) {
		res = "greaterThanOrEqual";
	} else if (c_oAscFilterOperator.GreaterThan === val) {
		res = "greaterThan";
	}
	return res;
}

function FromXml_ST_DateTimeGrouping(val) {
	var res = -1;
	if ("year" === val) {
		res = c_oAscDateTimeGrouping.Year;
	} else if ("month" === val) {
		res = c_oAscDateTimeGrouping.Month;
	} else if ("day" === val) {
		res = c_oAscDateTimeGrouping.Day;
	} else if ("hour" === val) {
		res = c_oAscDateTimeGrouping.Hour;
	} else if ("minute" === val) {
		res = c_oAscDateTimeGrouping.Minute;
	} else if ("second" === val) {
		res = c_oAscDateTimeGrouping.Second;
	}
	return res;
}
function ToXml_ST_DateTimeGrouping(val) {
	var res = "";
	if (c_oAscDateTimeGrouping.Year === val) {
		res = "year";
	} else if (c_oAscDateTimeGrouping.Month === val) {
		res = "month";
	} else if (c_oAscDateTimeGrouping.Day === val) {
		res = "day";
	} else if (c_oAscDateTimeGrouping.Hour === val) {
		res = "hour";
	} else if (c_oAscDateTimeGrouping.Minute === val) {
		res = "minute";
	} else if (c_oAscDateTimeGrouping.Second === val) {
		res = "second";
	}
	return res;
}

function FromXml_ST_AllocationMethod(val) {
	var res = -1;
	if ("equalAllocation" === val) {
		res = c_oAscAllocationMethod.EqualAllocation;
	} else if ("equalIncrement" === val) {
		res = c_oAscAllocationMethod.EqualIncrement;
	} else if ("weightedAllocation" === val) {
		res = c_oAscAllocationMethod.WeightedAllocation;
	} else if ("weightedIncrement" === val) {
		res = c_oAscAllocationMethod.WeightedIncrement;
	}
	return res;
}
function ToXml_ST_AllocationMethod(val) {
	var res = "";
	if (c_oAscAllocationMethod.EqualAllocation === val) {
		res = "equalAllocation";
	} else if (c_oAscAllocationMethod.EqualIncrement === val) {
		res = "equalIncrement";
	} else if (c_oAscAllocationMethod.WeightedAllocation === val) {
		res = "weightedAllocation";
	} else if (c_oAscAllocationMethod.WeightedIncrement === val) {
		res = "weightedIncrement";
	}
	return res;
}

function CT_PivotCacheDefinition() {
//Attributes
	this.id = null;
	this.invalid = false;
	this.saveData = true;
	this.refreshOnLoad = false;
	this.optimizeMemory = false;
	this.enableRefresh = true;
	this.refreshedBy = null;
	this.refreshedDate = null;
	this.backgroundQuery = false;
	this.missingItemsLimit = null;
	this.createdVersion = 0;
	this.refreshedVersion = 0;
	this.minRefreshableVersion = 0;
	// this.recordCount = null;
	this.upgradeOnRefresh = false;
	this.hasTupleCache = false;
	this.supportSubquery = false;
	this.supportAdvancedDrill = false;
//Members
	this.cacheSource = null;
	this.cacheFields = null;
	this.cacheHierarchies = null;
	this.kpis = null;
	this.tupleCache = null;
	this.calculatedItems = null;
	this.calculatedMembers = null;
	this.dimensions = null;
	this.measureGroups = null;
	this.maps = null;
	this.extLst = null;
	//editor
	this.cacheRecords = null;
	this.Id = AscCommon.g_oIdCounter.Get_NewId();
}
CT_PivotCacheDefinition.prototype.getType = function() {
	return AscCommonExcel.UndoRedoDataTypes.PivotCacheDefinition;
};
CT_PivotCacheDefinition.prototype.Get_Id = function () {
	return this.Id;
};
CT_PivotCacheDefinition.prototype.Write_ToBinary2 = function(w) {
	toXmlWithLength(w, this);
	if (this.cacheRecords) {
		w.WriteBool(true);
		this.cacheRecords.Write_ToBinary2(w);
	} else {
		w.WriteBool(false);
	}
};
CT_PivotCacheDefinition.prototype.Read_FromBinary2 = function(r) {
	var len = r.GetLong();
	new openXml.SaxParserBase().parse(AscCommon.GetStringUtf8(r, len), this);
	if (r.GetBool()) {
		this.cacheRecords = new CT_PivotCacheRecords();
		this.cacheRecords.Read_FromBinary2(r);
	}
};
CT_PivotCacheDefinition.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["r:id"];
		if (undefined !== val) {
			this.id = AscCommon.unleakString(uq(val));
		}
		val = vals["invalid"];
		if (undefined !== val) {
			this.invalid = AscCommon.getBoolFromXml(val);
		}
		val = vals["saveData"];
		if (undefined !== val) {
			this.saveData = AscCommon.getBoolFromXml(val);
		}
		val = vals["refreshOnLoad"];
		if (undefined !== val) {
			this.refreshOnLoad = AscCommon.getBoolFromXml(val);
		}
		val = vals["optimizeMemory"];
		if (undefined !== val) {
			this.optimizeMemory = AscCommon.getBoolFromXml(val);
		}
		val = vals["enableRefresh"];
		if (undefined !== val) {
			this.enableRefresh = AscCommon.getBoolFromXml(val);
		}
		val = vals["refreshedBy"];
		if (undefined !== val) {
			this.refreshedBy = AscCommon.unleakString(uq(val));
		}
		val = vals["refreshedDate"];
		if (undefined !== val) {
			this.refreshedDate = val - 0;
		}
		val = vals["backgroundQuery"];
		if (undefined !== val) {
			this.backgroundQuery = AscCommon.getBoolFromXml(val);
		}
		val = vals["missingItemsLimit"];
		if (undefined !== val) {
			this.missingItemsLimit = val - 0;
		}
		val = vals["createdVersion"];
		if (undefined !== val) {
			this.createdVersion = val - 0;
		}
		val = vals["refreshedVersion"];
		if (undefined !== val) {
			this.refreshedVersion = val - 0;
		}
		val = vals["minRefreshableVersion"];
		if (undefined !== val) {
			this.minRefreshableVersion = val - 0;
		}
		val = vals["upgradeOnRefresh"];
		if (undefined !== val) {
			this.upgradeOnRefresh = AscCommon.getBoolFromXml(val);
		}
		val = vals["tupleCache"];
		if (undefined !== val) {
			this.hasTupleCache = AscCommon.getBoolFromXml(val);
		}
		val = vals["supportSubquery"];
		if (undefined !== val) {
			this.supportSubquery = AscCommon.getBoolFromXml(val);
		}
		val = vals["supportAdvancedDrill"];
		if (undefined !== val) {
			this.supportAdvancedDrill = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_PivotCacheDefinition.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pivotCacheDefinition" === elem) {
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
	} else if ("cacheSource" === elem) {
		newContext = new CT_CacheSource();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.cacheSource = newContext;
	} else if ("cacheFields" === elem) {
		newContext = new CT_CacheFields();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.cacheFields = newContext;
	} else if ("cacheHierarchies" === elem) {
		newContext = new CT_CacheHierarchies();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.cacheHierarchies = newContext;
	} else if ("kpis" === elem) {
		newContext = new CT_PCDKPIs();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.kpis = newContext;
	} else if ("tupleCache" === elem) {
		newContext = new CT_TupleCache();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.tupleCache = newContext;
	} else if ("calculatedItems" === elem) {
		newContext = new CT_CalculatedItems();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.calculatedItems = newContext;
	} else if ("calculatedMembers" === elem) {
		newContext = new CT_CalculatedMembers();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.calculatedMembers = newContext;
	} else if ("dimensions" === elem) {
		newContext = new CT_Dimensions();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.dimensions = newContext;
	} else if ("measureGroups" === elem) {
		newContext = new CT_MeasureGroups();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.measureGroups = newContext;
	} else if ("maps" === elem) {
		newContext = new CT_MeasureDimensionMaps();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.maps = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PivotCacheDefinition.prototype.toXml = function(writer) {
	writer.WriteXmlString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
	writer.WriteXmlNodeStart("pivotCacheDefinition");
	writer.WriteXmlString(
		" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
	//todo
	// if (null !== this.id) {
		// writer.WriteXmlAttributeStringEncode("r:id", this.id);
	// }
	if (false !== this.invalid) {
		writer.WriteXmlAttributeBool("invalid", this.invalid);
	}
	if (true !== this.saveData) {
		writer.WriteXmlAttributeBool("saveData", this.saveData);
	}
	if (false !== this.refreshOnLoad) {
		writer.WriteXmlAttributeBool("refreshOnLoad", this.refreshOnLoad);
	}
	if (false !== this.optimizeMemory) {
		writer.WriteXmlAttributeBool("optimizeMemory", this.optimizeMemory);
	}
	if (true !== this.enableRefresh) {
		writer.WriteXmlAttributeBool("enableRefresh", this.enableRefresh);
	}
	if (null !== this.refreshedBy) {
		writer.WriteXmlAttributeStringEncode("refreshedBy", this.refreshedBy);
	}
	if (null !== this.refreshedDate) {
		writer.WriteXmlAttributeNumber("refreshedDate", this.refreshedDate);
	}
	if (false !== this.backgroundQuery) {
		writer.WriteXmlAttributeBool("backgroundQuery", this.backgroundQuery);
	}
	if (null !== this.missingItemsLimit) {
		writer.WriteXmlAttributeNumber("missingItemsLimit", this.missingItemsLimit);
	}
	if (0 !== this.createdVersion) {
		writer.WriteXmlAttributeNumber("createdVersion", this.createdVersion);
	}
	if (0 !== this.refreshedVersion) {
		writer.WriteXmlAttributeNumber("refreshedVersion", this.refreshedVersion);
	}
	if (0 !== this.minRefreshableVersion) {
		writer.WriteXmlAttributeNumber("minRefreshableVersion", this.minRefreshableVersion);
	}
	if (null !== this.cacheRecords) {
		writer.WriteXmlAttributeNumber("recordCount", this.cacheRecords.getRowsCount());
	}
	if (false !== this.upgradeOnRefresh) {
		writer.WriteXmlAttributeBool("upgradeOnRefresh", this.upgradeOnRefresh);
	}
	if (false !== this.hasTupleCache) {
		writer.WriteXmlAttributeBool("tupleCache", this.hasTupleCache);
	}
	if (false !== this.supportSubquery) {
		writer.WriteXmlAttributeBool("supportSubquery", this.supportSubquery);
	}
	if (false !== this.supportAdvancedDrill) {
		writer.WriteXmlAttributeBool("supportAdvancedDrill", this.supportAdvancedDrill);
	}
	writer.WriteXmlNodeEnd("pivotCacheDefinition", true);
	if (null !== this.cacheSource) {
		this.cacheSource.toXml(writer, "cacheSource");
	}
	if (null !== this.cacheFields) {
		this.cacheFields.toXml(writer, "cacheFields");
	}
	if (null !== this.cacheHierarchies) {
		this.cacheHierarchies.toXml(writer, "cacheHierarchies");
	}
	if (null !== this.kpis) {
		this.kpis.toXml(writer, "kpis");
	}
	if (null !== this.tupleCache) {
		this.tupleCache.toXml(writer, "tupleCache");
	}
	if (null !== this.calculatedItems) {
		this.calculatedItems.toXml(writer, "calculatedItems");
	}
	if (null !== this.calculatedMembers) {
		this.calculatedMembers.toXml(writer, "calculatedMembers");
	}
	if (null !== this.dimensions) {
		this.dimensions.toXml(writer, "dimensions");
	}
	if (null !== this.measureGroups) {
		this.measureGroups.toXml(writer, "measureGroups");
	}
	if (null !== this.maps) {
		this.maps.toXml(writer, "maps");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd("pivotCacheDefinition");
};
CT_PivotCacheDefinition.prototype.getFields = function () {
	return this.cacheFields && this.cacheFields.cacheField;
};
CT_PivotCacheDefinition.prototype.getRecords = function () {
	return this.cacheRecords;
};
CT_PivotCacheDefinition.prototype.isValidCacheSource = function () {
	return this.cacheSource && this.cacheSource.type === c_oAscSourceType.Worksheet;
};
CT_PivotCacheDefinition.prototype.getWorksheetSource = function() {
	return this.cacheSource && this.cacheSource.worksheetSource;
};
CT_PivotCacheDefinition.prototype.fromDataRef = function(dataRef) {
	this.cacheSource = new CT_CacheSource();
	this.cacheSource.type = c_oAscSourceType.Worksheet;
	this.cacheSource.worksheetSource = new CT_WorksheetSource();
	this.cacheSource.worksheetSource.fromDataRef(dataRef);
	this.cacheRecords = new CT_PivotCacheRecords();
	var location = this.cacheSource.worksheetSource.getDataLocation();
	if (location) {
		this.cacheFields = new CT_CacheFields();
		this.cacheRecords.fromWorksheetRange(location, this.cacheFields);
	}
};
function CT_PivotCacheRecords() {
//Attributes
//	this.count = null;
//Members
	this.extLst = null;
//internal
	this.startCount = 0;
	this._cols = [];
	this._curColIndex = 0;
}
CT_PivotCacheRecords.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["count"];
		if (undefined !== val) {
			this.startCount = val - 0;
		}
	}
};
CT_PivotCacheRecords.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pivotCacheRecords" === elem) {
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
	} else if ("r" === elem) {
		this._curColIndex = 0;
	} else {
		var newContextCandidate = this._getCol(this._curColIndex).onStartNode(elem, attr, uq);
		if (newContextCandidate) {
			newContext = newContextCandidate;
		} else if ("extLst" === elem) {
			newContext = new CT_ExtensionList();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.extLst = newContext;
		} else {
			newContext = null;
		}
	}
	return newContext;
};
CT_PivotCacheRecords.prototype.onEndNode = function(prevContext, elem) {
	if (this._getCol(this._curColIndex).onEndNode(prevContext, elem)) {
		this._curColIndex++;
	}
};
CT_PivotCacheRecords.prototype.toXml = function(writer) {
	writer.WriteXmlString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
	writer.WriteXmlNodeStart("pivotCacheRecords");
	writer.WriteXmlString(
		" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
	var count = this.getRowsCount();
	if (count > 0) {
		writer.WriteXmlAttributeNumber("count", count);
	}

	writer.WriteXmlNodeEnd("pivotCacheRecords", true);

	for (var i = 0; i < count; ++i) {
		writer.WriteXmlNodeStart("r", true);
		for (var j = 0; j < this._cols.length; ++j) {
			this._cols[j].toXml(writer, i);
		}
		writer.WriteXmlNodeEnd("r");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd("pivotCacheRecords");
};
CT_PivotCacheRecords.prototype.getRowsCount = function() {
	return this._cols && this._cols[0] && this._cols[0].getSize() || 0;
};
CT_PivotCacheRecords.prototype.get = function(row, col) {
	var col = this._cols[col];
	if (col) {
		return col.get(row);
	}
};
CT_PivotCacheRecords.prototype._getCol = function(index) {
	var col = this._cols[index];
	if(!col){
		col = new PivotRecords();
		col.setStartCount(this.startCount);
		this._cols[index] = col;
	}
	return col;
};
CT_PivotCacheRecords.prototype._getDataMapFromFields = function(indexes, length, row, dataMap) {
	var i, val;
	for (i = 0; i < length; ++i) {
		val = this._cols[indexes[i]].get(row);
		dataMap = this._getDataMapAddElem(dataMap, val.val, dataMap.total.length);
	}
	return dataMap;
};
CT_PivotCacheRecords.prototype._getDataMapAddElem = function(dataMap, val, dataLen) {
	var elem = dataMap.vals[val];
	if (!elem) {
		elem = new PivotDataElem(dataLen);
		dataMap.vals[val] = elem;
	}
	return elem;
};
CT_PivotCacheRecords.prototype._getDataMapMergeSubtotal = function(rowMap1, rowMap2) {
	var isEnd = true;
	for (var i in rowMap2.subtotal) {
		if (rowMap2.subtotal.hasOwnProperty(i)) {
			isEnd =false;
			var sub1 = rowMap1.subtotal[i];
			if (!sub1) {
				sub1 = new PivotDataElem(rowMap1.total.length);
				rowMap1.subtotal[i] = sub1;
			}
			var sub2 = rowMap2.subtotal[i];
			this._getDataMapMergeSubtotal(sub1, sub2);
		}
	}
	if(isEnd){
		rowMap1.unionTotal(rowMap2);
	}
};
CT_PivotCacheRecords.prototype._getDataMapSubtotal = function(rowMap, rowCount) {
	if (rowCount > 0) {
		rowMap.subtotal = {};
		for (var i in rowMap.vals) {
			if (rowMap.vals.hasOwnProperty(i)) {
				var sub = rowMap.vals[i];
				this._getDataMapSubtotal(sub, rowCount - 1);
				this._getDataMapMergeSubtotal(rowMap, sub);
			}
		}
	} else {
		rowMap.subtotal = rowMap.vals;
	}
};
CT_PivotCacheRecords.prototype._getDataMapTotal = function(rowMap) {
	var i;
	for (i in rowMap.subtotal) {
		if (rowMap.subtotal.hasOwnProperty(i)) {
			this._getDataMapTotal(rowMap.subtotal[i]);
			rowMap.unionTotal(rowMap.subtotal[i]);
		}
	}
	if (rowMap.subtotal !== rowMap.vals) {
		for (i in rowMap.vals) {
			if (rowMap.vals.hasOwnProperty(i)) {
				this._getDataMapTotal(rowMap.vals[i]);
			}
		}
	}
};
CT_PivotCacheRecords.prototype._getDataMapRowToTotal = function(cacheFields, row, rowMapCur, dataIndexes) {
	var i, val, total;
	for (i = 0; i < rowMapCur.total.length; ++i) {
		total = rowMapCur.total[i];
		var dataIndex = dataIndexes[i];
		val = this._cols[dataIndex].get(row);
		if (c_oAscPivotRecType.Index === val.type && cacheFields) {
			var cacheField = cacheFields[dataIndex];
			val = cacheField && cacheField.sharedItems && cacheField.sharedItems.Items && cacheField.sharedItems.Items.get(val.val) || val;
		}
		if (c_oAscPivotRecType.Number === val.type || c_oAscPivotRecType.DateTime === val.type) {
			total.add(val.val);
		} else if (c_oAscPivotRecType.Error === val.type) {
			total.addError(val.val);
		} else if (c_oAscPivotRecType.Missing !== val.type) {
			total.addCount(val.val);
		}
	}
};
CT_PivotCacheRecords.prototype.getDataMap = function(cacheFields, filterMaps, rowIndexes, colIndexes, dataIndexes) {
	var res = new PivotDataElem(dataIndexes.length);
	var filter, val, i, row, isInFilter, rowMapCur;
	for (row = 0; row < this.getRowsCount(); ++row) {
		isInFilter = true;
		for (i = 0; i < filterMaps.length; ++i) {
			filter = filterMaps[i];
			val = this._cols[filter.index[i]].get(row);
			if (!filter.field.has(val.val)) {
				isInFilter = false;
				break;
			}
		}
		if (!isInFilter) {
			continue;
		}
		rowMapCur = res;
		if (rowIndexes.length > 0) {
			rowMapCur = this._getDataMapFromFields(rowIndexes, rowIndexes.length, row, rowMapCur);
		}
		if (colIndexes.length > 0) {
			rowMapCur = this._getDataMapFromFields(colIndexes, colIndexes.length, row, rowMapCur);
		}
		this._getDataMapRowToTotal(cacheFields, row, rowMapCur, dataIndexes);
	}
	this._getDataMapSubtotal(res, rowIndexes.length);
	this._getDataMapTotal(res);
	return res;
};
CT_PivotCacheRecords.prototype.convertToSharedItems = function(index, si) {
	var col = this._cols[index];
	if (col) {
		col.convertToSharedItems(si);
	}
};
CT_PivotCacheRecords.prototype.updateCacheData = function() {
	this.cacheRecords.updateCacheData();
};
CT_PivotCacheRecords.prototype.fromWorksheetRange = function(location, cacheFields) {
	var ws = location.ws;
	var bbox = location.bbox;
	var headings = location.headings;
	if (!headings) {
		headings = [];
		ws.getRange3(bbox.r1, bbox.c1, bbox.r1, bbox.c2)._foreachNoEmpty(function(cell) {
			if (!cell.isNullTextString()) {
				headings.push(cell.getValue());
			}
		});
	}
	if(bbox.getWidth() !== headings.length){
		return;
	}
	var nameDuplicateMap = new Map();
	for (var i = 0; i < headings.length; ++i) {
		var text = headings[i];
		var index = nameDuplicateMap.get(text);
		if (undefined !== index) {
			index++;
			text = text + index;
		} else {
			index = 1;
		}
		nameDuplicateMap.set(text, index);
		var cacheField = new CT_CacheField();
		cacheField.name = text;
		cacheField.numFmtId = 0;
		cacheField.sharedItems = new CT_SharedItems();
		cacheFields.cacheField.push(cacheField);
	}
	this._cols = [];
	for (var i = 0; i < cacheFields.cacheField.length; ++i) {
		var cacheField = cacheFields.cacheField[i];
		var si = cacheField.sharedItems;
		si.containsSemiMixedTypes = false;
		si.containsNonDate = false;
		si.containsString = false;
		si.containsInteger = true;
		si.maxValue = si.maxDate = Number.NEGATIVE_INFINITY;
		si.minValue = si.minDate = Number.POSITIVE_INFINITY;
		var records = new PivotRecords();
		var lastRow = location.headings ? bbox.r1 - 1 : bbox.r1;
		ws.getRange3(lastRow + 1, bbox.c1 + i, bbox.r2, bbox.c1 + i)._foreachNoEmptyByCol(function(cell) {
			if (!cell.isNullTextString()) {
				if (lastRow + 1 < cell.nRow) {
					records.addMissing(cell.nRow - lastRow - 1);
				}
				lastRow = cell.nRow;
				switch (cell.getType()) {
					case AscCommon.CellValueType.Number:
						var num = cell.getNumberValue();
						if (!cell.getNumFormat().isDateTimeFormat()) {
							records.addNumber(num);
							si.containsNumber = true;
							si.containsNonDate = true;
							si.minValue = Math.min(si.minValue, num);
							si.maxValue = Math.max(si.maxValue, num);
							if (si.containsInteger && !Number.isInteger(num)) {
								si.containsInteger = false;
							}
						} else {
							records.addDate(num);
							si.containsDate = true;
							si.minDate = Math.min(si.minDate, num);
							si.maxDate = Math.max(si.maxDate, num);
						}
						break;
					case AscCommon.CellValueType.String:
						var text = cell.getValueWithoutFormat();
						records.addString(text);
						if (text.length > 255) {
							si.longText = true;
						}
						si.containsString = true;
						si.containsNonDate = true;
						break;
					case AscCommon.CellValueType.Bool:
						records.addBool(cell.getBoolValue());
						si.containsString = true;
						si.containsNonDate = true;
						break;
					case AscCommon.CellValueType.Error:
						records.addError(cell.getErrorValue());
						si.containsString = true;
						si.containsNonDate = true;
						break;
				}
			} else {
				si.containsBlank = true;
			}
		});
		if (lastRow < bbox.r2) {
			records.addMissing(bbox.r2 - lastRow);
		}
		this._cols.push(records);
		if (si.containsString || si.containsBlank) {
			si.containsSemiMixedTypes = true;
		}
		if ((si.containsDate + si.containsNumber + si.containsString) > 1) {
			si.containsMixedTypes = true;
		}
		if (si.containsDate && si.containsNumber) {
			si.containsNumber = false;
		}
		if (!si.containsDate) {
			si.minDate = null;
			si.maxDate = null;
		}
		if (!si.containsNumber) {
			si.containsInteger = false;
			si.minValue = null;
			si.maxValue = null;
		}
	}
	//todo trim missing
};
CT_PivotCacheRecords.prototype.getType = function() {
	return AscCommonExcel.UndoRedoDataTypes.PivotCacheRecords;
};
CT_PivotCacheRecords.prototype.Write_ToBinary2 = function(w) {
	//todo write binary
	toXmlWithLength(w, this);
};
CT_PivotCacheRecords.prototype.Read_FromBinary2 = function(r) {
	var len = r.GetLong();
	new openXml.SaxParserBase().parse(AscCommon.GetStringUtf8(r, len), this);
};

function PivotTableChanged() {
	this.oldRanges = null;
	this.style = false;
	this.data = false;
}
function CT_pivotTableDefinition(setDefaults) {
//Attributes
	this.name = null;
	this.cacheId = null;
	this.dataOnRows = null;
	this.dataPosition = null;
	this.autoFormatId = null;
	this.applyNumberFormats = null;
	this.applyBorderFormats = null;
	this.applyFontFormats = null;
	this.applyPatternFormats = null;
	this.applyAlignmentFormats = null;
	this.applyWidthHeightFormats = null;
	this.dataCaption = null;
	this.grandTotalCaption = null;
	this.errorCaption = null;
	this.showError = null;
	this.missingCaption = null;
	this.showMissing = null;
	this.pageStyle = null;
	this.pivotTableStyle = null;
	this.vacatedStyle = null;
	this.tag = null;
	this.updatedVersion = null;
	this.minRefreshableVersion = null;
	this.asteriskTotals = null;
	this.showItems = null;
	this.editData = null;
	this.disableFieldList = null;
	this.showCalcMbrs = null;
	this.visualTotals = null;
	this.showMultipleLabel = null;
	this.showDataDropDown = null;
	this.showDrill = null;
	this.printDrill = null;
	this.showMemberPropertyTips = null;
	this.showDataTips = null;
	this.enableWizard = null;
	this.enableDrill = null;
	this.enableFieldProperties = null;
	this.preserveFormatting = null;
	this.useAutoFormatting = null;
	this.pageWrap = null;
	this.pageOverThenDown = null;
	this.subtotalHiddenItems = null;
	this.rowGrandTotals = null;
	this.colGrandTotals = null;
	this.fieldPrintTitles = null;
	this.itemPrintTitles = null;
	this.mergeItem = null;
	this.showDropZones = null;
	this.createdVersion = null;
	this.indent = null;
	this.showEmptyRow = null;
	this.showEmptyCol = null;
	this.showHeaders = null;
	this.compact = null;
	this.outline = null;
	this.outlineData = null;
	this.compactData = null;
	this.published = null;
	this.gridDropZones = null;
	this.immersive = null;
	this.multipleFieldFilters = null;
	this.chartFormat = null;
	this.rowHeaderCaption = null;
	this.colHeaderCaption = null;
	this.fieldListSortAscending = null;
	this.mdxSubqueries = null;
	this.customListSort = null;
//Members
	this.location = null;
	this.pivotFields = null;
	this.rowFields = null;
	this.rowItems = null;
	this.colFields = null;
	this.colItems = null;
	this.pageFields = null;
	this.dataFields = null;
	this.formats = null;
	this.conditionalFormats = null;
	this.chartFormats = null;
	this.pivotHierarchies = null;
	this.pivotTableStyleInfo = new CT_PivotTableStyle(); // May be absent in some cases, so create
	this.pivotTableStyleInfo.showRowHeaders = true; // When creating a new pivot table, initialize by default
	this.pivotTableStyleInfo.showColHeaders = true; // When creating a new pivot table, initialize by default
	this.filters = null;
	this.rowHierarchiesUsage = null;
	this.colHierarchiesUsage = null;
	//ext
	this.pivotTableDefinitionX14 = null;
	//editor
	this.cacheDefinition = null;

	this.isInit = false;
	this.changed = new PivotTableChanged();
	this.pageFieldsPositions = null;
	this.clearGrid = false;
	this.hasCompactField = true;
	this.ascInsertBlankRow = null;
	this.ascDefaultSubtotal = null;
	this.ascSubtotalTop = null;
	this.ascFillDownLabels = null;
	this.ascDataRef = null;
	this.ascAltText = null;
	this.ascAltTextSummary = null;

	if (setDefaults) {
		this.setDefaults();
	}

	this.worksheet = null;
	this.Id = AscCommon.g_oIdCounter.Get_NewId();
}
CT_pivotTableDefinition.prototype.setDefaults = function () {
	this.dataOnRows = false;
	this.showError = false;
	this.showMissing = true;
	this.updatedVersion = 0;
	this.minRefreshableVersion = 0;
	this.asteriskTotals = false;
	this.showItems = true;
	this.editData = false;
	this.disableFieldList = false;
	this.showCalcMbrs = true;
	this.visualTotals = true;
	this.showMultipleLabel = true;
	this.showDataDropDown = true;
	this.showDrill = true;
	this.printDrill = false;
	this.showMemberPropertyTips = true;
	this.showDataTips = true;
	this.enableWizard = true;
	this.enableDrill = true;
	this.enableFieldProperties = true;
	this.preserveFormatting = true;
	this.useAutoFormatting = false;
	this.pageWrap = 0;
	this.pageOverThenDown = false;
	this.subtotalHiddenItems = false;
	this.rowGrandTotals = true;
	this.colGrandTotals = true;
	this.fieldPrintTitles = false;
	this.itemPrintTitles = false;
	this.mergeItem = false;
	this.showDropZones = true;
	this.createdVersion = 0;
	this.indent = 1;
	this.showEmptyRow = false;
	this.showEmptyCol = false;
	this.showHeaders = true;
	this.compact = true;
	this.outline = false;
	this.outlineData = false;
	this.compactData = true;
	this.published = false;
	this.gridDropZones = false;
	this.immersive = true;
	this.multipleFieldFilters = true;
	this.chartFormat = 0;
	this.fieldListSortAscending = false;
	this.mdxSubqueries = false;
	this.customListSort = true;
};
CT_pivotTableDefinition.prototype.getObjectType = function () {
	return AscDFH.historyitem_type_PivotTableDefinition;
};
CT_pivotTableDefinition.prototype.Get_Id = function () {
	return this.Id;
};
CT_pivotTableDefinition.prototype.GetWS = function () {
	return this.worksheet;
};
CT_pivotTableDefinition.prototype.setWS = function (ws) {
	this.worksheet = ws;
};
CT_pivotTableDefinition.prototype.clone = function () {
	var data = new AscCommonExcel.UndoRedoData_BinaryWrapper(this);
	var pivot = data.getData();
	pivot.Id = AscCommon.g_oIdCounter.Get_NewId();
	pivot.init();
	return pivot;
};
CT_pivotTableDefinition.prototype.cloneShallow = function () {
	var oldCacheDefinition = this.cacheDefinition;
	this.cacheDefinition = null;
	var newPivot = this.clone();
	this.cacheDefinition = newPivot.cacheDefinition = oldCacheDefinition;
	return newPivot;
};
CT_pivotTableDefinition.prototype.stashCurReportRange = function () {
	var t = this;
	if (!this.changed.oldRanges) {
		this.changed.oldRanges = this.getReportRanges();
		this.changed.oldRanges.forEach(function(range){
			t.GetWS().getRange3(range.r1, range.c1, range.r2, range.c2).clearTableStyle();
		});
	}
};
CT_pivotTableDefinition.prototype.stashEmptyReportRange = function () {
	if (!this.changed.oldRanges) {
		this.changed.oldRanges = [];
	}
};
CT_pivotTableDefinition.prototype.setChanged = function (data, style) {
	this.changed.data = data || this.changed.data ;
	this.changed.style = style || this.changed.style;
};
CT_pivotTableDefinition.prototype.getAndCleanChanged = function () {
	var res = this.changed;
	this.changed = new PivotTableChanged();
	return res;
};
CT_pivotTableDefinition.prototype.Write_ToBinary2 = function (w) {
	w.WriteLong(this.getObjectType());
	w.WriteString2(this.worksheet ? this.worksheet.getId() : '-1');
	toXmlWithLength(w, this);
	if (this.cacheDefinition) {
		w.WriteBool(true);
		this.cacheDefinition.Write_ToBinary2(w);
	} else {
		w.WriteBool(false);
	}
};
CT_pivotTableDefinition.prototype.Read_FromBinary2 = function (r) {
	this.setDefaults();
	var api_sheet = Asc['editor'];
	this.worksheet = api_sheet.wbModel.getWorksheetById(r.GetString2());
	var len = r.GetLong();
	new openXml.SaxParserBase().parse(AscCommon.GetStringUtf8(r, len), this);
	if (r.GetBool()) {
		this.cacheDefinition = new CT_PivotCacheDefinition();
		this.cacheDefinition.Read_FromBinary2(r);
	}
};
CT_pivotTableDefinition.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["cacheId"];
		if (undefined !== val) {
			this.cacheId = val - 0;
		}
		val = vals["dataOnRows"];
		if (undefined !== val) {
			this.dataOnRows = AscCommon.getBoolFromXml(val);
		}
		val = vals["dataPosition"];
		if (undefined !== val) {
			this.dataPosition = val - 0;
		}
		val = vals["autoFormatId"];
		if (undefined !== val) {
			this.autoFormatId = val - 0;
		}
		val = vals["applyNumberFormats"];
		if (undefined !== val) {
			this.applyNumberFormats = AscCommon.getBoolFromXml(val);
		}
		val = vals["applyBorderFormats"];
		if (undefined !== val) {
			this.applyBorderFormats = AscCommon.getBoolFromXml(val);
		}
		val = vals["applyFontFormats"];
		if (undefined !== val) {
			this.applyFontFormats = AscCommon.getBoolFromXml(val);
		}
		val = vals["applyPatternFormats"];
		if (undefined !== val) {
			this.applyPatternFormats = AscCommon.getBoolFromXml(val);
		}
		val = vals["applyAlignmentFormats"];
		if (undefined !== val) {
			this.applyAlignmentFormats = AscCommon.getBoolFromXml(val);
		}
		val = vals["applyWidthHeightFormats"];
		if (undefined !== val) {
			this.applyWidthHeightFormats = AscCommon.getBoolFromXml(val);
		}
		val = vals["dataCaption"];
		if (undefined !== val) {
			this.dataCaption = AscCommon.unleakString(uq(val));
		}
		val = vals["grandTotalCaption"];
		if (undefined !== val) {
			this.grandTotalCaption = AscCommon.unleakString(uq(val));
		}
		val = vals["errorCaption"];
		if (undefined !== val) {
			this.errorCaption = AscCommon.unleakString(uq(val));
		}
		val = vals["showError"];
		if (undefined !== val) {
			this.showError = AscCommon.getBoolFromXml(val);
		}
		val = vals["missingCaption"];
		if (undefined !== val) {
			this.missingCaption = AscCommon.unleakString(uq(val));
		}
		val = vals["showMissing"];
		if (undefined !== val) {
			this.showMissing = AscCommon.getBoolFromXml(val);
		}
		val = vals["pageStyle"];
		if (undefined !== val) {
			this.pageStyle = AscCommon.unleakString(uq(val));
		}
		val = vals["pivotTableStyle"];
		if (undefined !== val) {
			this.pivotTableStyle = AscCommon.unleakString(uq(val));
		}
		val = vals["vacatedStyle"];
		if (undefined !== val) {
			this.vacatedStyle = AscCommon.unleakString(uq(val));
		}
		val = vals["tag"];
		if (undefined !== val) {
			this.tag = AscCommon.unleakString(uq(val));
		}
		val = vals["updatedVersion"];
		if (undefined !== val) {
			this.updatedVersion = val - 0;
		}
		val = vals["minRefreshableVersion"];
		if (undefined !== val) {
			this.minRefreshableVersion = val - 0;
		}
		val = vals["asteriskTotals"];
		if (undefined !== val) {
			this.asteriskTotals = AscCommon.getBoolFromXml(val);
		}
		val = vals["showItems"];
		if (undefined !== val) {
			this.showItems = AscCommon.getBoolFromXml(val);
		}
		val = vals["editData"];
		if (undefined !== val) {
			this.editData = AscCommon.getBoolFromXml(val);
		}
		val = vals["disableFieldList"];
		if (undefined !== val) {
			this.disableFieldList = AscCommon.getBoolFromXml(val);
		}
		val = vals["showCalcMbrs"];
		if (undefined !== val) {
			this.showCalcMbrs = AscCommon.getBoolFromXml(val);
		}
		val = vals["visualTotals"];
		if (undefined !== val) {
			this.visualTotals = AscCommon.getBoolFromXml(val);
		}
		val = vals["showMultipleLabel"];
		if (undefined !== val) {
			this.showMultipleLabel = AscCommon.getBoolFromXml(val);
		}
		val = vals["showDataDropDown"];
		if (undefined !== val) {
			this.showDataDropDown = AscCommon.getBoolFromXml(val);
		}
		val = vals["showDrill"];
		if (undefined !== val) {
			this.showDrill = AscCommon.getBoolFromXml(val);
		}
		val = vals["printDrill"];
		if (undefined !== val) {
			this.printDrill = AscCommon.getBoolFromXml(val);
		}
		val = vals["showMemberPropertyTips"];
		if (undefined !== val) {
			this.showMemberPropertyTips = AscCommon.getBoolFromXml(val);
		}
		val = vals["showDataTips"];
		if (undefined !== val) {
			this.showDataTips = AscCommon.getBoolFromXml(val);
		}
		val = vals["enableWizard"];
		if (undefined !== val) {
			this.enableWizard = AscCommon.getBoolFromXml(val);
		}
		val = vals["enableDrill"];
		if (undefined !== val) {
			this.enableDrill = AscCommon.getBoolFromXml(val);
		}
		val = vals["enableFieldProperties"];
		if (undefined !== val) {
			this.enableFieldProperties = AscCommon.getBoolFromXml(val);
		}
		val = vals["preserveFormatting"];
		if (undefined !== val) {
			this.preserveFormatting = AscCommon.getBoolFromXml(val);
		}
		val = vals["useAutoFormatting"];
		if (undefined !== val) {
			this.useAutoFormatting = AscCommon.getBoolFromXml(val);
		}
		val = vals["pageWrap"];
		if (undefined !== val) {
			this.pageWrap = val - 0;
		}
		val = vals["pageOverThenDown"];
		if (undefined !== val) {
			this.pageOverThenDown = AscCommon.getBoolFromXml(val);
		}
		val = vals["subtotalHiddenItems"];
		if (undefined !== val) {
			this.subtotalHiddenItems = AscCommon.getBoolFromXml(val);
		}
		val = vals["rowGrandTotals"];
		if (undefined !== val) {
			this.rowGrandTotals = AscCommon.getBoolFromXml(val);
		}
		val = vals["colGrandTotals"];
		if (undefined !== val) {
			this.colGrandTotals = AscCommon.getBoolFromXml(val);
		}
		val = vals["fieldPrintTitles"];
		if (undefined !== val) {
			this.fieldPrintTitles = AscCommon.getBoolFromXml(val);
		}
		val = vals["itemPrintTitles"];
		if (undefined !== val) {
			this.itemPrintTitles = AscCommon.getBoolFromXml(val);
		}
		val = vals["mergeItem"];
		if (undefined !== val) {
			this.mergeItem = AscCommon.getBoolFromXml(val);
		}
		val = vals["showDropZones"];
		if (undefined !== val) {
			this.showDropZones = AscCommon.getBoolFromXml(val);
		}
		val = vals["createdVersion"];
		if (undefined !== val) {
			this.createdVersion = val - 0;
		}
		val = vals["indent"];
		if (undefined !== val) {
			this.indent = val - 0;
		}
		val = vals["showEmptyRow"];
		if (undefined !== val) {
			this.showEmptyRow = AscCommon.getBoolFromXml(val);
		}
		val = vals["showEmptyCol"];
		if (undefined !== val) {
			this.showEmptyCol = AscCommon.getBoolFromXml(val);
		}
		val = vals["showHeaders"];
		if (undefined !== val) {
			this.showHeaders = AscCommon.getBoolFromXml(val);
		}
		val = vals["compact"];
		if (undefined !== val) {
			this.compact = AscCommon.getBoolFromXml(val);
		}
		val = vals["outline"];
		if (undefined !== val) {
			this.outline = AscCommon.getBoolFromXml(val);
		}
		val = vals["outlineData"];
		if (undefined !== val) {
			this.outlineData = AscCommon.getBoolFromXml(val);
		}
		val = vals["compactData"];
		if (undefined !== val) {
			this.compactData = AscCommon.getBoolFromXml(val);
		}
		val = vals["published"];
		if (undefined !== val) {
			this.published = AscCommon.getBoolFromXml(val);
		}
		val = vals["gridDropZones"];
		if (undefined !== val) {
			this.gridDropZones = AscCommon.getBoolFromXml(val);
		}
		val = vals["immersive"];
		if (undefined !== val) {
			this.immersive = AscCommon.getBoolFromXml(val);
		}
		val = vals["multipleFieldFilters"];
		if (undefined !== val) {
			this.multipleFieldFilters = AscCommon.getBoolFromXml(val);
		}
		val = vals["chartFormat"];
		if (undefined !== val) {
			this.chartFormat = val - 0;
		}
		val = vals["rowHeaderCaption"];
		if (undefined !== val) {
			this.rowHeaderCaption = AscCommon.unleakString(uq(val));
		}
		val = vals["colHeaderCaption"];
		if (undefined !== val) {
			this.colHeaderCaption = AscCommon.unleakString(uq(val));
		}
		val = vals["fieldListSortAscending"];
		if (undefined !== val) {
			this.fieldListSortAscending = AscCommon.getBoolFromXml(val);
		}
		val = vals["mdxSubqueries"];
		if (undefined !== val) {
			this.mdxSubqueries = AscCommon.getBoolFromXml(val);
		}
		val = vals["customListSort"];
		if (undefined !== val) {
			this.customListSort = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_pivotTableDefinition.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pivotTableDefinition" === elem) {
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
	} else if ("location" === elem) {
		newContext = new CT_Location();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.location = newContext;
	} else if ("pivotFields" === elem) {
		newContext = new CT_PivotFields();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pivotFields = newContext;
	} else if ("rowFields" === elem) {
		newContext = new CT_RowFields();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.rowFields = newContext;
	} else if ("rowItems" === elem) {
		newContext = new CT_rowItems();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.rowItems = newContext;
	} else if ("colFields" === elem) {
		newContext = new CT_ColFields();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.colFields = newContext;
	} else if ("colItems" === elem) {
		newContext = new CT_colItems();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.colItems = newContext;
	} else if ("pageFields" === elem) {
		newContext = new CT_PageFields();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pageFields = newContext;
	} else if ("dataFields" === elem) {
		newContext = new CT_DataFields();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.dataFields = newContext;
	} else if ("formats" === elem) {
		newContext = new CT_Formats();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.formats = newContext;
	} else if ("conditionalFormats" === elem) {
		newContext = new CT_ConditionalFormats();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.conditionalFormats = newContext;
	} else if ("chartFormats" === elem) {
		newContext = new CT_ChartFormats();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.chartFormats = newContext;
	} else if ("pivotHierarchies" === elem) {
		newContext = new CT_PivotHierarchies();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pivotHierarchies = newContext;
	} else if ("pivotTableStyleInfo" === elem) {
		newContext = new CT_PivotTableStyle();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pivotTableStyleInfo = newContext;
	} else if ("filters" === elem) {
		newContext = new CT_PivotFilters();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.filters = newContext;
	} else if ("rowHierarchiesUsage" === elem) {
		newContext = new CT_RowHierarchiesUsage();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.rowHierarchiesUsage = newContext;
	} else if ("colHierarchiesUsage" === elem) {
		newContext = new CT_ColHierarchiesUsage();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.colHierarchiesUsage = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
	} else {
		newContext = null;
	}
	return newContext;
};
CT_pivotTableDefinition.prototype.onEndNode = function(prevContext, elem) {
	if ("extLst" === elem) {
		for (var i = 0; i < prevContext.ext.length; ++i) {
			var ext = prevContext.ext[i];
			if ('{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}' == ext.uri) {
				this.pivotTableDefinitionX14 = ext.elem;
			}
		}
	}
};
CT_pivotTableDefinition.prototype.toXml = function(writer) {
	writer.WriteXmlString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
	writer.WriteXmlNodeStart("pivotTableDefinition");
	writer.WriteXmlString(
		" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (null !== this.cacheId) {
		writer.WriteXmlAttributeNumber("cacheId", this.cacheId);
	}
	if (false !== this.dataOnRows) {
		writer.WriteXmlAttributeBool("dataOnRows", this.dataOnRows);
	}
	if (null !== this.dataPosition) {
		writer.WriteXmlAttributeNumber("dataPosition", this.dataPosition);
	}
	if (null !== this.autoFormatId) {
		writer.WriteXmlAttributeNumber("autoFormatId", this.autoFormatId);
	}
	if (null !== this.applyNumberFormats) {
		writer.WriteXmlAttributeBool("applyNumberFormats", this.applyNumberFormats);
	}
	if (null !== this.applyBorderFormats) {
		writer.WriteXmlAttributeBool("applyBorderFormats", this.applyBorderFormats);
	}
	if (null !== this.applyFontFormats) {
		writer.WriteXmlAttributeBool("applyFontFormats", this.applyFontFormats);
	}
	if (null !== this.applyPatternFormats) {
		writer.WriteXmlAttributeBool("applyPatternFormats", this.applyPatternFormats);
	}
	if (null !== this.applyAlignmentFormats) {
		writer.WriteXmlAttributeBool("applyAlignmentFormats", this.applyAlignmentFormats);
	}
	if (null !== this.applyWidthHeightFormats) {
		writer.WriteXmlAttributeBool("applyWidthHeightFormats", this.applyWidthHeightFormats);
	}
	if (null !== this.dataCaption) {
		writer.WriteXmlAttributeStringEncode("dataCaption", this.dataCaption);
	}
	if (null !== this.grandTotalCaption) {
		writer.WriteXmlAttributeStringEncode("grandTotalCaption", this.grandTotalCaption);
	}
	if (null !== this.errorCaption) {
		writer.WriteXmlAttributeStringEncode("errorCaption", this.errorCaption);
	}
	if (false !== this.showError) {
		writer.WriteXmlAttributeBool("showError", this.showError);
	}
	if (null !== this.missingCaption) {
		writer.WriteXmlAttributeStringEncode("missingCaption", this.missingCaption);
	}
	if (true !== this.showMissing) {
		writer.WriteXmlAttributeBool("showMissing", this.showMissing);
	}
	if (null !== this.pageStyle) {
		writer.WriteXmlAttributeStringEncode("pageStyle", this.pageStyle);
	}
	if (null !== this.pivotTableStyle) {
		writer.WriteXmlAttributeStringEncode("pivotTableStyle", this.pivotTableStyle);
	}
	if (null !== this.vacatedStyle) {
		writer.WriteXmlAttributeStringEncode("vacatedStyle", this.vacatedStyle);
	}
	if (null !== this.tag) {
		writer.WriteXmlAttributeStringEncode("tag", this.tag);
	}
	if (0 !== this.updatedVersion) {
		writer.WriteXmlAttributeNumber("updatedVersion", this.updatedVersion);
	}
	if (0 !== this.minRefreshableVersion) {
		writer.WriteXmlAttributeNumber("minRefreshableVersion", this.minRefreshableVersion);
	}
	if (false !== this.asteriskTotals) {
		writer.WriteXmlAttributeBool("asteriskTotals", this.asteriskTotals);
	}
	if (true !== this.showItems) {
		writer.WriteXmlAttributeBool("showItems", this.showItems);
	}
	if (false !== this.editData) {
		writer.WriteXmlAttributeBool("editData", this.editData);
	}
	if (false !== this.disableFieldList) {
		writer.WriteXmlAttributeBool("disableFieldList", this.disableFieldList);
	}
	if (true !== this.showCalcMbrs) {
		writer.WriteXmlAttributeBool("showCalcMbrs", this.showCalcMbrs);
	}
	if (true !== this.visualTotals) {
		writer.WriteXmlAttributeBool("visualTotals", this.visualTotals);
	}
	if (true !== this.showMultipleLabel) {
		writer.WriteXmlAttributeBool("showMultipleLabel", this.showMultipleLabel);
	}
	if (true !== this.showDataDropDown) {
		writer.WriteXmlAttributeBool("showDataDropDown", this.showDataDropDown);
	}
	if (true !== this.showDrill) {
		writer.WriteXmlAttributeBool("showDrill", this.showDrill);
	}
	if (false !== this.printDrill) {
		writer.WriteXmlAttributeBool("printDrill", this.printDrill);
	}
	if (true !== this.showMemberPropertyTips) {
		writer.WriteXmlAttributeBool("showMemberPropertyTips", this.showMemberPropertyTips);
	}
	if (true !== this.showDataTips) {
		writer.WriteXmlAttributeBool("showDataTips", this.showDataTips);
	}
	if (true !== this.enableWizard) {
		writer.WriteXmlAttributeBool("enableWizard", this.enableWizard);
	}
	if (true !== this.enableDrill) {
		writer.WriteXmlAttributeBool("enableDrill", this.enableDrill);
	}
	if (true !== this.enableFieldProperties) {
		writer.WriteXmlAttributeBool("enableFieldProperties", this.enableFieldProperties);
	}
	if (true !== this.preserveFormatting) {
		writer.WriteXmlAttributeBool("preserveFormatting", this.preserveFormatting);
	}
	if (false !== this.useAutoFormatting) {
		writer.WriteXmlAttributeBool("useAutoFormatting", this.useAutoFormatting);
	}
	if (0 !== this.pageWrap) {
		writer.WriteXmlAttributeNumber("pageWrap", this.pageWrap);
	}
	if (false !== this.pageOverThenDown) {
		writer.WriteXmlAttributeBool("pageOverThenDown", this.pageOverThenDown);
	}
	if (false !== this.subtotalHiddenItems) {
		writer.WriteXmlAttributeBool("subtotalHiddenItems", this.subtotalHiddenItems);
	}
	if (true !== this.rowGrandTotals) {
		writer.WriteXmlAttributeBool("rowGrandTotals", this.rowGrandTotals);
	}
	if (true !== this.colGrandTotals) {
		writer.WriteXmlAttributeBool("colGrandTotals", this.colGrandTotals);
	}
	if (false !== this.fieldPrintTitles) {
		writer.WriteXmlAttributeBool("fieldPrintTitles", this.fieldPrintTitles);
	}
	if (false !== this.itemPrintTitles) {
		writer.WriteXmlAttributeBool("itemPrintTitles", this.itemPrintTitles);
	}
	if (false !== this.mergeItem) {
		writer.WriteXmlAttributeBool("mergeItem", this.mergeItem);
	}
	if (true !== this.showDropZones) {
		writer.WriteXmlAttributeBool("showDropZones", this.showDropZones);
	}
	if (0 !== this.createdVersion) {
		writer.WriteXmlAttributeNumber("createdVersion", this.createdVersion);
	}
	if (1 !== this.indent) {
		writer.WriteXmlAttributeNumber("indent", this.indent);
	}
	if (false !== this.showEmptyRow) {
		writer.WriteXmlAttributeBool("showEmptyRow", this.showEmptyRow);
	}
	if (false !== this.showEmptyCol) {
		writer.WriteXmlAttributeBool("showEmptyCol", this.showEmptyCol);
	}
	if (true !== this.showHeaders) {
		writer.WriteXmlAttributeBool("showHeaders", this.showHeaders);
	}
	if (true !== this.compact) {
		writer.WriteXmlAttributeBool("compact", this.compact);
	}
	if (false !== this.outline) {
		writer.WriteXmlAttributeBool("outline", this.outline);
	}
	if (false !== this.outlineData) {
		writer.WriteXmlAttributeBool("outlineData", this.outlineData);
	}
	if (true !== this.compactData) {
		writer.WriteXmlAttributeBool("compactData", this.compactData);
	}
	if (false !== this.published) {
		writer.WriteXmlAttributeBool("published", this.published);
	}
	if (false !== this.gridDropZones) {
		writer.WriteXmlAttributeBool("gridDropZones", this.gridDropZones);
	}
	if (true !== this.immersive) {
		writer.WriteXmlAttributeBool("immersive", this.immersive);
	}
	if (true !== this.multipleFieldFilters) {
		writer.WriteXmlAttributeBool("multipleFieldFilters", this.multipleFieldFilters);
	}
	if (0 !== this.chartFormat) {
		writer.WriteXmlAttributeNumber("chartFormat", this.chartFormat);
	}
	if (null !== this.rowHeaderCaption) {
		writer.WriteXmlAttributeStringEncode("rowHeaderCaption", this.rowHeaderCaption);
	}
	if (null !== this.colHeaderCaption) {
		writer.WriteXmlAttributeStringEncode("colHeaderCaption", this.colHeaderCaption);
	}
	if (false !== this.fieldListSortAscending) {
		writer.WriteXmlAttributeBool("fieldListSortAscending", this.fieldListSortAscending);
	}
	if (false !== this.mdxSubqueries) {
		writer.WriteXmlAttributeBool("mdxSubqueries", this.mdxSubqueries);
	}
	if (true !== this.customListSort) {
		writer.WriteXmlAttributeBool("customListSort", this.customListSort);
	}
	writer.WriteXmlNodeEnd("pivotTableDefinition", true);
	if (null !== this.location) {
		this.location.toXml(writer, "location");
	}
	if (null !== this.pivotFields) {
		this.pivotFields.toXml(writer, "pivotFields");
	}
	if (null !== this.rowFields) {
		this.rowFields.toXml(writer, "rowFields");
	}
	if (null !== this.rowItems) {
		this.rowItems.toXml(writer, "rowItems");
	}
	if (null !== this.colFields) {
		this.colFields.toXml(writer, "colFields");
	}
	if (null !== this.colItems) {
		this.colItems.toXml(writer, "colItems");
	}
	if (null !== this.pageFields) {
		this.pageFields.toXml(writer, "pageFields");
	}
	if (null !== this.dataFields) {
		this.dataFields.toXml(writer, "dataFields");
	}
	if (null !== this.formats) {
		this.formats.toXml(writer, "formats");
	}
	if (null !== this.conditionalFormats) {
		this.conditionalFormats.toXml(writer, "conditionalFormats");
	}
	if (null !== this.chartFormats) {
		this.chartFormats.toXml(writer, "chartFormats");
	}
	if (null !== this.pivotHierarchies) {
		this.pivotHierarchies.toXml(writer, "pivotHierarchies");
	}
	if (null !== this.pivotTableStyleInfo) {
		this.pivotTableStyleInfo.toXml(writer, "pivotTableStyleInfo");
	}
	if (null !== this.filters) {
		this.filters.toXml(writer, "filters");
	}
	if (null !== this.rowHierarchiesUsage) {
		this.rowHierarchiesUsage.toXml(writer, "rowHierarchiesUsage");
	}
	if (null !== this.colHierarchiesUsage) {
		this.colHierarchiesUsage.toXml(writer, "colHierarchiesUsage");
	}
	if (null !== this.pivotTableDefinitionX14) {
		var ext = new CT_Extension();
		ext.uri = "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}";
		ext.elem = this.pivotTableDefinitionX14;
		var extList = new CT_ExtensionList();
		extList.ext.push(ext);
		extList.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd("pivotTableDefinition");
};
CT_pivotTableDefinition.prototype.init = function () {
	this.isInit = true;
	this.pageFieldsPositions = [];
	var pageField = this.asc_getPageFields();
	var rowPageCount = 0, colPageCount = 0, r, c;
	if (pageField) {
		var pageFieldSize = this.getPageFieldSize();
		rowPageCount = pageFieldSize.row;
		colPageCount = pageFieldSize.col;
		var range = this.getRange();
		var baseCol = range.c1;
		var baseRow = range.r1 - pageFieldSize.row - 1;
		var pageWrap = this.pageWrap || Number.MAX_VALUE;
		for (var i = 0; i < pageField.length; ++i) {
			if (this.pageOverThenDown) {
				r = Math.floor(i / pageWrap);
				c = 3 * (i % pageWrap);
			} else {
				r = (i % pageWrap);
				c = 3 * Math.floor(i / pageWrap);
			}
			this.pageFieldsPositions.push(new AscCommon.CellBase(Math.max(0, baseRow + r), Math.max(0, baseCol + c)));
		}
	}

	if (this.location) {
		this.location.setPageCount(rowPageCount, colPageCount);
	}
	this.updatePivotType();
};
CT_pivotTableDefinition.prototype.updatePivotType = function () {
	this.clearGrid = false;
	this.hasCompactField = false;
	var pivotFields = this.asc_getPivotFields();
	var rowFields = this.asc_getRowFields();
	if (rowFields) {
		var i, index;
		for (i = 0; i < rowFields.length; ++i) {
			index = rowFields[i].asc_getIndex();
			if (st_VALUES !== index && false !== pivotFields[index].outline) {
				this.clearGrid = true;
				break;
			}
		}
		for (i = 0; i < pivotFields.length; ++i) {
			if (false !== pivotFields[i].compact) {
				this.hasCompactField = true;
				break;
			}
		}
	}
};
CT_pivotTableDefinition.prototype.hasCompact = function () {
	return false !== this.compactData || this.hasCompactField;
};
CT_pivotTableDefinition.prototype.intersection = function (range) {
	return (this.location && this.location.intersection(range)) || this.pageFieldsIntersection(range);
};
CT_pivotTableDefinition.prototype.isInRange = function(bbox) {
	var ranges = this.getReportRanges();
	return ranges.every(function(range) {
		return bbox.containsRange(range);
	});
};
CT_pivotTableDefinition.prototype.pageFieldsIntersection = function (range) {
	return this.pageFieldsPositions && this.pageFieldsPositions.some(function (element) {
			return Array.isArray(range) ? range.some(function (elementRange) {
				return (elementRange.contains(element.col, element.row) ||
				elementRange.contains(element.col + 1, element.row));
			}) : (range.contains(element.col, element.row) || range.contains(element.col + 1, element.row));
		});
};
CT_pivotTableDefinition.prototype.contains = function (col, row) {
	return (this.location && this.location.contains(col, row)) ||
		this.pageFieldsIntersection(new Asc.Range(col, row, col, row));
};
CT_pivotTableDefinition.prototype.getRange = function () {
	return this.location && this.location.ref;
};
CT_pivotTableDefinition.prototype.getReportRanges = function () {
	var res = [], pos;
	if (this.pageFieldsPositions) {
		for (var i = 0; i < this.pageFieldsPositions.length; ++i) {
			pos = this.pageFieldsPositions[i];
			res.push(new Asc.Range(pos.col, pos.row, pos.col + 1, pos.row));
		}
	}

	var pivotRange = this.getRange();
	res.push(new Asc.Range(pivotRange.c1, pivotRange.r1, pivotRange.c2, pivotRange.r2));
	return res;
};
CT_pivotTableDefinition.prototype.getFirstHeaderRow0 = function () {
	return this.location && (this.location.firstHeaderRow + this.getColumnFieldsCount() - 1);
};
CT_pivotTableDefinition.prototype.getFirstDataCol = function () {
	return this.location && this.location.firstDataCol;
};
CT_pivotTableDefinition.prototype.getColumnFieldsCount = function (withoutValues) {
	var res = 0;
	var colFields = this.asc_getColumnFields();
	if (colFields) {
		res = colFields.length;
		if (1 === res && withoutValues && st_VALUES === colFields[0].x) {
			res = 0;
		}
	}
	return res;
};
CT_pivotTableDefinition.prototype.getColumnFieldsValuesIndex = function () {
	var colFields = this.asc_getColumnFields();
	if (colFields) {
		return colFields && colFields.findIndex(function(element) {
				return element.asc_getIndex() === st_VALUES;
			});
	}
	return -1;
};
CT_pivotTableDefinition.prototype.getRowFieldsCount = function (compact) {
	var t = this, res = 0, l;
	var rowFields = this.asc_getRowFields();
	if (rowFields) {
		l = res = rowFields.length;
		if (compact) {
			this.getField(rowFields, function (element, i) {
				if (i !== l - 1) {
					var field = t.pivotFields.pivotField[element.asc_getIndex()];
					res -= (field && false !== field.outline && false !== field.compact) ? 1 : 0;
				}
			});
		}
	}
	return res;
};
CT_pivotTableDefinition.prototype.getRowFieldPos = function (index) {
	var res = 0;
	var rowFields = this.asc_getRowFields();
	if (rowFields) {
		var field;
		for (var i = 0; i < index && i < rowFields.length; ++i) {
			field = this.pivotFields.pivotField[rowFields[i].asc_getIndex()];
			res += (field && (false === field.outline || false === field.compact) && 1);
		}
	}

	return res;
};
CT_pivotTableDefinition.prototype.getRowFieldsValuesIndex = function() {
	var rowFields = this.asc_getRowFields();
	if (rowFields) {
		return rowFields.findIndex(function(element) {
				return element.asc_getIndex() === st_VALUES;
			});
	}
	return -1;
};
CT_pivotTableDefinition.prototype.getDataFieldsCount = function () {
	return (this.dataFields && this.dataFields.dataField.length) || 0;
};
CT_pivotTableDefinition.prototype.getField = function (arrFields, callback) {
	return arrFields && arrFields.map(callback, this);
};
CT_pivotTableDefinition.prototype.getRowItems = function () {
	return this.rowItems && this.rowItems.i;
};
CT_pivotTableDefinition.prototype.getColItems = function () {
	return this.colItems && this.colItems.i;
};
CT_pivotTableDefinition.prototype.getRecords = function () {
	return this.cacheDefinition.getRecords();
};
CT_pivotTableDefinition.prototype.getAllRange = function (ws) {
	var newSelection = new AscCommonExcel.SelectionRange(ws);
	newSelection.assign2(this.getRange());
	if (this.pageFieldsPositions && 0 < this.pageFieldsPositions.length) {
		this.pageFieldsPositions.forEach(function (element) {
			newSelection.addRange();
			newSelection.getLast().assign2(new Asc.Range(element.col, element.row, element.col + 1, element.row));
		});
		newSelection = newSelection.getUnion();
	}
	return newSelection;
};
CT_pivotTableDefinition.prototype.asc_getName = function () {
	return this.name;
};
CT_pivotTableDefinition.prototype.asc_getPageWrap = function () {
	return this.pageWrap || 0;
};
CT_pivotTableDefinition.prototype.asc_getPageOverThenDown = function () {
	return !!this.pageOverThenDown;
};
CT_pivotTableDefinition.prototype.asc_getRowGrandTotals = function () {
	return this.rowGrandTotals;
};
CT_pivotTableDefinition.prototype.asc_getColGrandTotals = function () {
	return this.colGrandTotals;
};
CT_pivotTableDefinition.prototype.asc_getShowHeaders = function () {
	return this.showHeaders;
};
CT_pivotTableDefinition.prototype.asc_getTitle = function () {
	return this.pivotTableDefinitionX14 && this.pivotTableDefinitionX14.altText;
};
CT_pivotTableDefinition.prototype.asc_getDescription = function () {
	return this.pivotTableDefinitionX14 && this.pivotTableDefinitionX14.altTextSummary;
};
CT_pivotTableDefinition.prototype.asc_getStyleInfo = function () {
	return this.pivotTableStyleInfo;
};
CT_pivotTableDefinition.prototype.asc_getCacheFields = function () {
	return this.cacheDefinition.getFields();
};
CT_pivotTableDefinition.prototype.asc_getPivotFields = function () {
	return this.pivotFields && this.pivotFields.pivotField;
};
CT_pivotTableDefinition.prototype.asc_getPageFields = function () {
	return this.pageFields && this.pageFields.pageField.length > 0 && this.pageFields.pageField;
};
CT_pivotTableDefinition.prototype.asc_getColumnFields = function () {
	return this.colFields && this.colFields.field.length > 0 && this.colFields.field;
};
CT_pivotTableDefinition.prototype.asc_getRowFields = function () {
	return this.rowFields && this.rowFields.field.length > 0 && this.rowFields.field;
};
CT_pivotTableDefinition.prototype.asc_getDataFields = function () {
	return this.dataFields && this.dataFields.dataField.length > 0 && this.dataFields.dataField;
};
CT_pivotTableDefinition.prototype.asc_select = function (api) {
	this.getAllRange(api.wbModel.getActiveWs()).Select();
};
CT_pivotTableDefinition.prototype.asc_set = function (api, newVal) {
	var t = this;
	if (null !== newVal.ascDataRef && newVal.ascDataRef !== t.asc_getDataRef() && !this.isValidDataRef(newVal.ascDataRef)) {
		api.sendEvent('asc_onError', c_oAscError.ID.PivotLabledColumns, c_oAscError.Level.NoCritical);
		return;
	}
	api._changePivotWithLock(this, function (ws) {
		if (null !== newVal.name) {
			t.asc_setName(newVal.name, true);
		}
		if (null !== newVal.rowGrandTotals) {
			t.asc_setRowGrandTotals(newVal.rowGrandTotals, true);
		}
		if (null !== newVal.colGrandTotals) {
			t.asc_setColGrandTotals(newVal.colGrandTotals, true);
		}
		if (null !== newVal.pageOverThenDown) {
			t.asc_setPageOverThenDown(newVal.pageOverThenDown, true);
		}
		if (null !== newVal.pageWrap) {
			t.asc_setPageWrap(newVal.pageWrap, true);
		}
		if (null !== newVal.showHeaders) {
			t.asc_setShowHeaders(newVal.showHeaders, true);
		}
		if (null !== newVal.compact) {
			t.setCompact(newVal.compact, true);
		}
		if (null !== newVal.outline) {
			t.setOutline(newVal.outline, true);
		}
		if (null != newVal.ascFillDownLabels) {
			t.setFillDownLabelsDefault(newVal.ascFillDownLabels, true);
		}
		if (null !== newVal.ascDataRef && newVal.ascDataRef !== t.asc_getDataRef()) {
			t.updateCacheData(newVal.ascDataRef);
		}
		if (null != newVal.ascAltText) {
			t.setTitle(newVal.ascAltText, true);
		}
		if (null != newVal.ascAltTextSummary) {
			t.setDescription(newVal.ascAltTextSummary, true);
		}
		if (null !== newVal.ascInsertBlankRow) {
			t.setInsertBlankRow(newVal.ascInsertBlankRow, true);
		}
		if (null !== newVal.ascDefaultSubtotal) {
			t.setDefaultSubtotal(newVal.ascDefaultSubtotal, true);
		}
		if (null !== newVal.ascSubtotalTop) {
			t.setSubtotalTop(newVal.ascSubtotalTop, true);
		}
	});
};
CT_pivotTableDefinition.prototype.asc_setName = function(newVal, addToHistory) {
	setTableProperty(this, this.name, newVal, addToHistory, AscCH.historyitem_PivotTable_SetName);
	this.name = newVal;
};
CT_pivotTableDefinition.prototype.asc_setRowGrandTotals = function(newVal, addToHistory) {
	setTableProperty(this, this.rowGrandTotals, newVal, addToHistory, AscCH.historyitem_PivotTable_SetRowGrandTotals, true);
	this.rowGrandTotals = newVal;
};
CT_pivotTableDefinition.prototype.asc_setColGrandTotals = function(newVal, addToHistory) {
	setTableProperty(this, this.colGrandTotals, newVal, addToHistory, AscCH.historyitem_PivotTable_SetColGrandTotals, true);
	this.colGrandTotals = newVal;
};
CT_pivotTableDefinition.prototype.asc_setPageOverThenDown = function(newVal, addToHistory) {
	setTableProperty(this, this.pageOverThenDown, newVal, addToHistory, AscCH.historyitem_PivotTable_SetPageOverThenDown, true);
	this.pageOverThenDown = newVal;
};
CT_pivotTableDefinition.prototype.asc_setPageWrap = function(newVal, addToHistory) {
	setTableProperty(this, this.pageWrap, newVal, addToHistory, AscCH.historyitem_PivotTable_SetPageWrap, true);
	this.pageWrap = newVal;
};
CT_pivotTableDefinition.prototype.asc_setShowHeaders = function(newVal, addToHistory) {
	setTableProperty(this, this.showHeaders, newVal, addToHistory, AscCH.historyitem_PivotTable_SetShowHeaders, true);
	this.showHeaders = newVal;
};
CT_pivotTableDefinition.prototype.asc_setCompact = function(newVal, addToHistory) {
	setTableProperty(this, this.compact, newVal, addToHistory, AscCH.historyitem_PivotTable_SetCompact, true);
	this.compact = newVal;
	this.compactData = this.compact;
};
CT_pivotTableDefinition.prototype.setCompact = function(newVal, addToHistory) {
	this.asc_setCompact(newVal, addToHistory);
	var pivotFields = this.asc_getPivotFields();
	if (addToHistory && pivotFields) {
		for (var i = 0; i < pivotFields.length; ++i) {
			pivotFields[i].asc_setCompact(newVal, this, i, addToHistory);
		}
	}
};
CT_pivotTableDefinition.prototype.asc_setOutline = function(newVal, addToHistory) {
	setTableProperty(this, this.outline, newVal, addToHistory, AscCH.historyitem_PivotTable_SetOutline, true);
	this.outline = newVal;
	this.outlineData = this.outline;
};
CT_pivotTableDefinition.prototype.setOutline = function(newVal, addToHistory) {
	this.asc_setOutline(newVal, addToHistory);
	var pivotFields = this.asc_getPivotFields();
	if (addToHistory && pivotFields) {
		for (var i = 0; i < pivotFields.length; ++i) {
			pivotFields[i].asc_setOutline(newVal, this, i, addToHistory);
		}
	}
};
CT_pivotTableDefinition.prototype.asc_setFillDownLabelsDefault = function(newVal) {
	this.ascFillDownLabels = newVal;
};
CT_pivotTableDefinition.prototype.setFillDownLabelsDefault = function(newVal, addToHistory) {
	if (!this.pivotTableDefinitionX14) {
		this.pivotTableDefinitionX14 = new CT_pivotTableDefinitionX14();
	}
	var oldVal = this.pivotTableDefinitionX14.fillDownLabelsDefault;
	setTableProperty(this, oldVal, newVal, addToHistory, AscCH.historyitem_PivotTable_SetFillDownLabelsDefault, true);
	this.pivotTableDefinitionX14.fillDownLabelsDefault = newVal;
	//todo
};
CT_pivotTableDefinition.prototype.asc_setDataRef = function(newVal) {
	this.ascDataRef = newVal;
};
CT_pivotTableDefinition.prototype.asc_setTitle = function(newVal) {
	this.ascAltText = newVal;
};
CT_pivotTableDefinition.prototype.setTitle = function(newVal, addToHistory) {
	if (!this.pivotTableDefinitionX14) {
		this.pivotTableDefinitionX14 = new CT_pivotTableDefinitionX14();
	}
	var oldVal = this.pivotTableDefinitionX14.altText;
	setTableProperty(this, oldVal, newVal, addToHistory, AscCH.historyitem_PivotTable_SetAltText, true);
	this.pivotTableDefinitionX14.altText = newVal;
};
CT_pivotTableDefinition.prototype.asc_setDescription = function(newVal, addToHistory) {
	this.ascAltTextSummary = newVal;
};
CT_pivotTableDefinition.prototype.setDescription = function(newVal, addToHistory) {
	if (!this.pivotTableDefinitionX14) {
		this.pivotTableDefinitionX14 = new CT_pivotTableDefinitionX14();
	}
	var oldVal = this.pivotTableDefinitionX14.altTextSummary;
	setTableProperty(this, oldVal, newVal, addToHistory, AscCH.historyitem_PivotTable_SetAltTextSummary, true);
	this.pivotTableDefinitionX14.altTextSummary = newVal;
};
CT_pivotTableDefinition.prototype.asc_setInsertBlankRow = function(newVal) {
	this.ascInsertBlankRow = newVal;
};
CT_pivotTableDefinition.prototype.setInsertBlankRow = function(newVal, addToHistory) {
	var pivotFields = this.asc_getPivotFields();
	if (pivotFields) {
		for (var i = 0; i < pivotFields.length; ++i) {
			pivotFields[i].asc_setInsertBlankRow(newVal, this, i, addToHistory);
		}
	}
};
CT_pivotTableDefinition.prototype.asc_setDefaultSubtotal = function(newVal) {
	this.ascDefaultSubtotal = newVal;
};
CT_pivotTableDefinition.prototype.setDefaultSubtotal = function(newVal, addToHistory) {
	var pivotFields = this.asc_getPivotFields();
	if (pivotFields) {
		for (var i = 0; i < pivotFields.length; ++i) {
			pivotFields[i].asc_setDefaultSubtotal(newVal, this, i, addToHistory);
			//todo
			pivotFields[i].checkSubtotal();
		}
	}
};
CT_pivotTableDefinition.prototype.asc_setSubtotalTop = function(newVal) {
	this.ascSubtotalTop = newVal;
};
CT_pivotTableDefinition.prototype.setSubtotalTop = function(newVal, addToHistory) {
	var pivotFields = this.asc_getPivotFields();
	if (pivotFields) {
		for (var i = 0; i < pivotFields.length; ++i) {
			pivotFields[i].asc_setSubtotalTop(newVal, this, i, addToHistory);
		}
	}
};
CT_pivotTableDefinition.prototype.asc_addDataField = function(api, pivotIndex, insertIndex) {
	var t = this;
	api._changePivotWithLock(this, function() {
		t.addDataField(pivotIndex, insertIndex, true);
	});
};
CT_pivotTableDefinition.prototype.addDataField = function(pivotIndex, insertIndex, addToHistory) {
	var pivotField = this.asc_getPivotFields()[pivotIndex];
	pivotField.dataField = true;
	var newField = new CT_DataField(true);
	newField.fld = pivotIndex;
	var cacheField = this.asc_getCacheFields()[pivotIndex];
	if (cacheField && cacheField.isSumSubtotal()) {
		newField.subtotal = c_oAscDataConsolidateFunction.Sum;
	} else {
		newField.subtotal = c_oAscDataConsolidateFunction.Count;
	}
	if (!this.dataFields) {
		this.dataFields = new CT_DataFields();
	}
	//todo translation
	var newName = AscCommon.translateManager.getValue(FIELD_CAPTION);
	newName = newName.replace("%1", ToName_ST_DataConsolidateFunction(newField.subtotal)).replace("%2", this.getPivotFieldName(pivotIndex));
	if (this.dataFields.checkDuplicateName(newName)) {
		var lastChar = newName.slice(-1);
		var delimiter = "";
		if ("0" <= lastChar && lastChar <= "9") {
			delimiter = "_";
		}
		var index = 2;
		while (this.dataFields.checkDuplicateName(newName + delimiter + index)) {
			index++;
		}
		newName = newName + delimiter + index;
	}
	newField.name = newName;
	newField.baseField = 0;
	newField.baseItem = 0;
	insertIndex = this.dataFields.add(newField, insertIndex);
	this.checkValuesField(addToHistory);
	if (addToHistory) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, AscCH.historyitem_PivotTable_AddDataField,
			this.worksheet ? this.worksheet.getId() : null, null,
			new AscCommonExcel.UndoRedoData_PivotTable(this.Get_Id(), pivotIndex, insertIndex));
	}
	this.setChanged(true);
};
CT_pivotTableDefinition.prototype.checkValuesField = function(addToHistory) {
	var newField;
	if(this.getDataFieldsCount() > 1){
		if (this.dataOnRows) {
			this.colFields && this.colFields.remove(st_VALUES);
			if (-1 === this.getRowFieldsValuesIndex()) {
				newField = new CT_Field();
				newField.x = st_VALUES;
				if (!this.rowFields) {
					this.rowFields = new CT_RowFields();
				}
				if(null !== this.dataPosition && this.dataPosition < this.rowFields.getCount()){
					this.rowFields.add(newField, this.dataPosition);
				} else {
					this.setDataPosition(null, addToHistory);
					this.rowFields.add(newField);
				}
			}
		} else {
			this.rowFields && this.rowFields.remove(st_VALUES);
			if (-1 === this.getColumnFieldsValuesIndex()) {
				newField = new CT_Field();
				newField.x = st_VALUES;
				if (!this.colFields) {
					this.colFields = new CT_ColFields();
				}
				if(null !== this.dataPosition && this.dataPosition < this.colFields.getCount()){
					this.colFields.add(newField, this.dataPosition);
				} else {
					this.setDataPosition(null, addToHistory);
					this.colFields.add(newField);
				}
			}
		}
	} else {
		this.colFields && this.colFields.remove(st_VALUES);
		this.rowFields && this.rowFields.remove(st_VALUES);
	}
};
CT_pivotTableDefinition.prototype.asc_addRowField = function(api, pivotIndex, insertIndex) {
	var t = this;
	api._changePivotWithLock(this, function() {
		var deleteIndex = t.removeNoDataField(pivotIndex, true);
		insertIndex = t.checkInsertIndex(insertIndex, deleteIndex);
		t.addRowField(pivotIndex, insertIndex, true);
	});
};
CT_pivotTableDefinition.prototype.addRowField = function(pivotIndex, insertIndex, addToHistory) {
	var pivotField = this.asc_getPivotFields()[pivotIndex];
	pivotField.axis = c_oAscAxis.AxisRow;
	var newField = new CT_Field();
	newField.x = pivotIndex;
	if (!this.rowFields) {
		this.rowFields = new CT_RowFields();
	}
	insertIndex = this.rowFields.add(newField, insertIndex);
	if (addToHistory) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, AscCH.historyitem_PivotTable_AddRowField,
			this.worksheet ? this.worksheet.getId() : null, null,
			new AscCommonExcel.UndoRedoData_PivotTable(this.Get_Id(), pivotIndex, insertIndex));
	}
	this.setChanged(true);
};
CT_pivotTableDefinition.prototype.asc_addColField = function(api, pivotIndex, insertIndex) {
	var t = this;
	api._changePivotWithLock(this, function() {
		var deleteIndex = t.removeNoDataField(pivotIndex, true);
		insertIndex = t.checkInsertIndex(insertIndex, deleteIndex);
		t.addColField(pivotIndex, insertIndex, true);
	});
};
CT_pivotTableDefinition.prototype.addColField = function(pivotIndex, insertIndex, addToHistory) {
	var pivotField = this.asc_getPivotFields()[pivotIndex];
	pivotField.axis = c_oAscAxis.AxisCol;
	var newField = new CT_Field();
	newField.x = pivotIndex;
	if (!this.colFields) {
		this.colFields = new CT_ColFields();
	}
	insertIndex = this.colFields.add(newField, insertIndex);
	if (addToHistory) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, AscCH.historyitem_PivotTable_AddColField,
			this.worksheet ? this.worksheet.getId() : null, null,
			new AscCommonExcel.UndoRedoData_PivotTable(this.Get_Id(), pivotIndex, insertIndex));
	}
	this.setChanged(true);
};
CT_pivotTableDefinition.prototype.asc_addPageField = function(api, pivotIndex, insertIndex) {
	var t = this;
	api._changePivotWithLock(this, function() {
		var deleteIndex = t.removeNoDataField(pivotIndex, true);
		insertIndex = t.checkInsertIndex(insertIndex, deleteIndex);
		t.addPageField(pivotIndex, insertIndex, true);
	});
};
CT_pivotTableDefinition.prototype.addPageField = function(pivotIndex, insertIndex, addToHistory) {
	var pivotField = this.asc_getPivotFields()[pivotIndex];
	pivotField.axis = c_oAscAxis.AxisPage;
	var newField = new CT_PageField();
	newField.fld = pivotIndex;
	newField.hier = -1;
	if (!this.pageFields) {
		this.pageFields = new CT_PageFields();
	}
	insertIndex = this.pageFields.add(newField, insertIndex);
	if (addToHistory) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, AscCH.historyitem_PivotTable_AddPageField,
			this.worksheet ? this.worksheet.getId() : null, null,
			new AscCommonExcel.UndoRedoData_PivotTable(this.Get_Id(), pivotIndex, insertIndex));
	}
	this.setChanged(true);
};
CT_pivotTableDefinition.prototype.asc_addField = function(api, index) {
	var cacheField = this.asc_getCacheFields()[index];
	if (cacheField && cacheField.isSumSubtotal()) {
		this.asc_addDataField(api, index);
	} else {
		this.asc_addRowField(api, index);
	}
};
CT_pivotTableDefinition.prototype.asc_removeField = function(api, pivotIndex) {
	if (st_VALUES === pivotIndex) {
		return;
	}
	var t = this;
	api._changePivotWithLock(this, function(ws) {
		t.removeNoDataField(pivotIndex, true);
		t.removeDataField(pivotIndex, undefined, true);
	});
};
CT_pivotTableDefinition.prototype.asc_removeNoDataField = function(api, pivotIndex) {
	var t = this;
	api._changePivotWithLock(this, function(ws) {
		if (st_VALUES === pivotIndex) {
			var dataFields = t.asc_getDataFields();
			if (dataFields) {
				for (var i = dataFields.length - 1; i >= 0; --i) {
					var dataField = dataFields[i];
					t.removeDataField(dataField.asc_getIndex(), i, true);
				}
			}
		} else {
			t.removeNoDataField(pivotIndex, true);
		}
	});
};
CT_pivotTableDefinition.prototype.asc_removeDataField = function(api, pivotIndex, dataIndex) {
	if (st_VALUES === pivotIndex) {
		return;
	}
	var t = this;
	api._changePivotWithLock(this, function(ws) {
		t.removeDataField(pivotIndex, dataIndex, true);
	});
};
CT_pivotTableDefinition.prototype.asc_moveToPageField = function(api, pivotIndex, dataIndex) {
	if (st_VALUES === pivotIndex) {
		return;
	}
	var t = this;
	api._changePivotWithLock(this, function(ws) {
		var deleteIndex = t.removeNoDataField(pivotIndex, true);
		if (undefined === deleteIndex && undefined !== dataIndex) {
			t.removeDataField(pivotIndex, dataIndex, true);
		}
		t.addPageField(pivotIndex, undefined, true);
	});
};
CT_pivotTableDefinition.prototype.asc_moveToRowField = function(api, pivotIndex, dataIndex) {
	var t = this;
	api._changePivotWithLock(this, function(ws) {
		if (st_VALUES === pivotIndex) {
			if (t.dataOnRows !== true) {
				t.setDataPosition(null, true);
			}
			t.setDataOnRows(true, true);
		} else {
			var deleteIndex = t.removeNoDataField(pivotIndex, true);
			if (undefined === deleteIndex && undefined !== dataIndex) {
				t.removeDataField(pivotIndex, dataIndex, true);
			}
			t.addRowField(pivotIndex, undefined, true);
		}
	});
};
CT_pivotTableDefinition.prototype.asc_moveToColField = function(api, pivotIndex, dataIndex) {
	var t = this;
	api._changePivotWithLock(this, function(ws) {
		if (st_VALUES === pivotIndex) {
			if (t.dataOnRows !== false) {
				t.setDataPosition(null, true);
			}
			t.setDataOnRows(false, true);
		} else {
			var deleteIndex = t.removeNoDataField(pivotIndex, true);
			if (undefined === deleteIndex && undefined !== dataIndex) {
				t.removeDataField(pivotIndex, dataIndex, true);
			}
			t.addColField(pivotIndex, undefined, true);
		}
	});
};
CT_pivotTableDefinition.prototype.asc_moveToDataField = function(api, pivotIndex, dataIndex) {
	if (st_VALUES === pivotIndex) {
		return;
	}
	var t = this;
	api._changePivotWithLock(this, function(ws) {
		var deleteIndex = t.removeNoDataField(pivotIndex, true);
		if (undefined === deleteIndex && undefined !== dataIndex) {
			t.removeDataField(pivotIndex, dataIndex, true);
		}
		t.addDataField(pivotIndex, undefined, true);
	});
};
CT_pivotTableDefinition.prototype.setDataOnRows = function(newVal, addToHistory) {
	setTableProperty(this, this.dataOnRows, newVal, addToHistory, AscCH.historyitem_PivotTable_SetDataOnRows, true);
	this.dataOnRows = newVal;
	this.checkValuesField(addToHistory);
};
CT_pivotTableDefinition.prototype.setDataPosition = function(newVal, addToHistory) {
	setTableProperty(this, this.dataPosition, newVal, addToHistory, AscCH.historyitem_PivotTable_SetDataPosition, true);
	this.dataPosition = newVal;
};
CT_pivotTableDefinition.prototype.asc_movePageField = function(api, from, to) {
	var t = this;
	api._changePivotWithLock(this, function(ws) {
		t.moveField(t.asc_getPageFields(), from, to, true, AscCH.historyitem_PivotTable_MovePageField);
	});
};
CT_pivotTableDefinition.prototype.asc_moveRowField = function(api, from, to) {
	var t = this;
	api._changePivotWithLock(this, function(ws) {
		t.moveDataPosition(t.asc_getRowFields(), from, to, true);
		t.moveField(t.asc_getRowFields(), from, to, true, AscCH.historyitem_PivotTable_MoveRowField);
	});
};
CT_pivotTableDefinition.prototype.asc_moveColField = function(api, from, to) {
	var t = this;
	api._changePivotWithLock(this, function(ws) {
		t.moveDataPosition(t.asc_getColumnFields(), from, to, true);
		t.moveField(t.asc_getColumnFields(), from, to, true, AscCH.historyitem_PivotTable_MoveColField);
	});
};
CT_pivotTableDefinition.prototype.asc_moveDataField = function(api, from, to) {
	var t = this;
	api._changePivotWithLock(this, function(ws) {
		t.moveField(t.asc_getDataFields(), from, to, true, AscCH.historyitem_PivotTable_MoveDataField);
	});
};
CT_pivotTableDefinition.prototype.moveDataPosition = function(arr, from, to, addToHistory) {
	if(st_VALUES === arr[from].x){
		if (to + 1 < arr.length) {
			this.setDataPosition(to, addToHistory);
		} else {
			this.setDataPosition(null, addToHistory);
		}
	}
};
CT_pivotTableDefinition.prototype.moveField = function(arr, from, to, addToHistory, historyType) {
	if (arr && 0 <= from && from < arr.length && 0 <= to && to < arr.length) {
		arr.splice(to, 0, arr.splice(from, 1)[0]);
		if (addToHistory) {
			History.Add(AscCommonExcel.g_oUndoRedoPivotTables, historyType,
						this.worksheet ? this.worksheet.getId() : null, null,
						new AscCommonExcel.UndoRedoData_PivotTable(this.Get_Id(), from, to));
		}
		this.setChanged(true);
	}
};
CT_pivotTableDefinition.prototype.asc_refresh = function(api) {
	var dataRef = this.asc_getDataRef();
	if (Asc.CT_pivotTableDefinition.prototype.isValidDataRef(dataRef)) {
		var t = this;
		api._changePivotWithLock(this, function (ws) {
			t.updateCacheData(dataRef);
		});
	} else {
		api.sendEvent('asc_onError', c_oAscError.ID.PivotLabledColumns, c_oAscError.Level.NoCritical);
	}
};
CT_pivotTableDefinition.prototype.asc_getDataRef = function() {
	return this.cacheDefinition && this.cacheDefinition.getWorksheetSource() && this.cacheDefinition.getWorksheetSource().getDataRef() || '';
};
CT_pivotTableDefinition.prototype.updateCacheData = function(dataRef) {
	var oldPivot = new AscCommonExcel.UndoRedoData_BinaryWrapper(this);
	var newCacheDefinition = new CT_PivotCacheDefinition();
	newCacheDefinition.fromDataRef(dataRef);

	var pivotFieldsMap = new Map();
	var newCTPivotFields = new CT_PivotFields();
	this._updateCacheDataUpdatePivotFieldsIndexes(newCacheDefinition, newCTPivotFields, pivotFieldsMap);
	var newCTPageFields = null;
	if (this.asc_getPageFields()) {
		newCTPageFields = new CT_PageFields();
		this._updateCacheDataUpdatePageDataFieldsIndexes(this.asc_getPageFields(), newCTPageFields.pageField, pivotFieldsMap);
	}
	var newCTDataFields = null;
	if (this.asc_getDataFields()) {
		newCTDataFields = new CT_DataFields();
		this._updateCacheDataUpdatePageDataFieldsIndexes(this.asc_getDataFields(), newCTDataFields.dataField, pivotFieldsMap);
	}
	var newCTRowFields = this._updateCacheDataUpdateRowColFieldsIndexes(this.asc_getRowFields(), new CT_RowFields(), newCTDataFields, pivotFieldsMap);
	var newCTColFields = this._updateCacheDataUpdateRowColFieldsIndexes(this.asc_getColumnFields(), new CT_ColFields(), newCTDataFields, pivotFieldsMap);

	this.cacheDefinition = newCacheDefinition;
	this.pivotFields = newCTPivotFields;
	this.pageFields = newCTPageFields;
	this.dataFields = newCTDataFields;
	this.rowFields = newCTRowFields;
	this.colFields = newCTColFields;
	this.setChanged(true);
	var newPivot = new AscCommonExcel.UndoRedoData_BinaryWrapper(this);
	History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_PivotReplace, this.worksheet.getId(),
		null, new AscCommonExcel.UndoRedoData_PivotTableRedo(this.Get_Id(), oldPivot, newPivot));
};
CT_pivotTableDefinition.prototype.setCacheDefinition = function(newCacheDefinition) {
	this.cacheDefinition = newCacheDefinition;
};
CT_pivotTableDefinition.prototype._updateCacheDataUpdatePivotFieldsIndexes = function(newCacheDefinition, newCTPivotFields, pivotFieldsMap) {
	var i;
	var cacheDefinitionMap = new Map();
	var newCacheFields = newCacheDefinition.getFields();
	for (i = 0; i < newCacheFields.length; ++i) {
		cacheDefinitionMap.set(newCacheFields[i].asc_getName(), i);
	}

	newCTPivotFields.fillWithEmpty(newCacheFields.length);
	var newPivotFields = newCTPivotFields.pivotField;

	var oldPivotFields = this.asc_getPivotFields();
	var oldCacheFields = this.cacheDefinition.getFields();
	for (i = 0; i < oldCacheFields.length; ++i) {
		var oldPivotField = oldPivotFields[i];
		var oldCacheField = oldCacheFields[i];
		var newIndex = cacheDefinitionMap.get(oldCacheField.asc_getName());
		if (undefined !== newIndex && oldPivotField) {
			this._updateCacheDataUpdatePivotFieldsIndexesItems(oldCacheField, oldPivotField, newIndex, newCacheDefinition);
			//oldPivotField.items = null;
			newPivotFields[newIndex] = oldPivotField;
			pivotFieldsMap.set(i, newIndex);
		}
	}
};
CT_pivotTableDefinition.prototype._updateCacheDataUpdatePivotFieldsIndexesItems = function(oldCacheField, oldPivotField, newIndex, newCacheDefinition) {
	var newCacheFields = newCacheDefinition.getFields();
	if (null !== oldPivotField.items && newCacheFields && newCacheFields[newIndex]) {
		this.checkPivotFieldItem(newIndex, oldPivotField, newCacheDefinition.cacheRecords, newCacheFields[newIndex], oldCacheField);
	}
};
CT_pivotTableDefinition.prototype._updateCacheDataUpdatePageDataFieldsIndexes = function(oldFields, newFields, pivotFieldsMap) {
	for (var i = 0; i < oldFields.length; ++i) {
		var oldDataField = oldFields[i];
		var newIndex = pivotFieldsMap.get(oldDataField.fld);
		if (undefined !== newIndex) {
			oldDataField.fld = newIndex;
			newFields.push(oldDataField);
		}
	}
};
CT_pivotTableDefinition.prototype._updateCacheDataUpdateRowColFieldsIndexes = function(oldFields, newCTFields, newCTDataFields, pivotFieldsMap) {
	if (!oldFields) {
		return null;
	}
	for (var i = 0; i < oldFields.length; ++i) {
		var oldField = oldFields[i];
		if (st_VALUES === oldField.x) {
			if (newCTDataFields.getCount() > 1) {
				newCTFields.field.push(oldField);
			}
		} else {
			var newIndex = pivotFieldsMap.get(oldField.x);
			if (undefined !== newIndex) {
				oldField.x = newIndex;
				newCTFields.field.push(oldField);
			}
		}
	}
	return newCTFields;
};
CT_pivotTableDefinition.prototype.asc_create = function(ws, name, cacheDefinition, bbox) {
	this.worksheet = ws;
	this.cacheDefinition = cacheDefinition;

	this.cacheId = null;
	this.name = name;
	this.applyNumberFormats = false;
	this.applyBorderFormats = false;
	this.applyFontFormats = false;
	this.applyPatternFormats = false;
	this.applyAlignmentFormats = false;
	this.applyWidthHeightFormats = true;
	this.dataCaption = AscCommon.translateManager.getValue(DATA_CAPTION);
	this.useAutoFormatting = true;
	this.itemPrintTitles = true;
	this.indent = 0;
	this.outline = true;
	this.outlineData = true;
	this.multipleFieldFilters = false;

	this.pivotFields = new CT_PivotFields();
	this.pivotFields.fillWithEmpty(this.cacheDefinition.getFields().length);

	this.pivotTableStyleInfo = new CT_PivotTableStyle();
	this.pivotTableStyleInfo.name = this.worksheet.workbook.TableStyles.DefaultPivotStyle;
	this.pivotTableStyleInfo.showRowHeaders = true;
	this.pivotTableStyleInfo.showColHeaders = true;
	this.pivotTableStyleInfo.showRowStripes = false;
	this.pivotTableStyleInfo.showColStripes = false;
	this.pivotTableStyleInfo.showLastColumn = true;

	this.location = new CT_Location();
	this.location.ref = bbox;
	this.updateLocation();

	this.init();
};
CT_pivotTableDefinition.prototype.checkPivotFieldItems = function(index) {
	var pivotField = this.asc_getPivotFields()[index];
	var cacheField, cacheRecords, i;
	if (pivotField && !pivotField.items) {
		cacheField = this.asc_getCacheFields()[index];
		cacheRecords = this.getRecords();
		if (cacheField && cacheRecords) {
			this.checkPivotFieldItem(index, pivotField, cacheRecords, cacheField);
		}
	}
};
CT_pivotTableDefinition.prototype.checkPivotFieldItem = function(index, pivotField, cacheRecords, cacheField, oldCacheField) {
	var item, i, j, newItem, equalMap = {};
	var newItems = new CT_Items();
	if (!(cacheField.sharedItems && cacheField.sharedItems.Items.getSize() > 0)) {
		if(!cacheField.sharedItems){
			cacheField.sharedItems = new CT_SharedItems();
		}
		cacheRecords.convertToSharedItems(index, cacheField.sharedItems);
	}
	//save old items order
	if (pivotField.items && oldCacheField) {
		for (i = 0; i < pivotField.items.item.length; ++i) {
			item = pivotField.items.item[i];
			if (Asc.c_oAscItemType.Data === item.t && !item.m) {
				var oldSharedItem = oldCacheField.getSharedItem(item.x);
				if (oldSharedItem) {
					for (j = 0; j < cacheField.getSharedSize(); ++j) {
						if (oldSharedItem.shallowEqual(cacheField.getSharedItem(j))) {
							//create new to lose other flags
							newItem = new CT_Item();
							newItem.x = j;
							newItems.item.push(newItem);
							equalMap[newItem.x] = 1;
							break;
						}
					}
				}
			}
		}
	}
	for (i = 0; i < cacheField.sharedItems.Items.getSize(); ++i) {
		if(!equalMap[i]){
			newItem = new CT_Item();
			newItem.x = i;
			newItems.item.push(newItem);
		}
	}
	if (!oldCacheField) {
		ctItemSort(newItems.item, cacheField.sharedItems);
	}
	pivotField.items = newItems;
	pivotField.checkSubtotal();
};
CT_pivotTableDefinition.prototype.getFilterMaps = function() {
	var filterFields = [];
	var i, j, map, hasHidden, pageField, pivotField, item;
	if (this.pageFields) {
		for (i = 0; i < this.pageFields.length; ++i) {
			pageField = this.pageFields[i];
			if (null !== pageField.fld) {
				this.checkPivotFieldItems(pageField.fld);
				if (null !== pageField.item) {
					map = new Map();
					map.set(pageField.item, 1);
					filterFields.push(pageField.fld, map);
				} else {
					pivotField = this.asc_getPivotFields()[pageField.fld];
					if (pivotField && pivotField.multipleItemSelectionAllowed && pivotField.items) {
						map = new Map();
						hasHidden = false;
						for (j = 0; j < pivotField.items.item.length; ++j) {
							item = pivotField.items.item[j];
							if (true === item.h) {
								hasHidden = true;
							} else {
								map.set(pageField.item, 1);
							}
						}
						if (hasHidden) {
							filterFields.push(pageField.fld, map);
						}
					}
				}
			}
		}
	}
	return filterFields;
};
CT_pivotTableDefinition.prototype.getRowColIndexes = function(fields, withoutValue) {
	var indexes = [];
	var i;
	if (fields) {
		for (i = 0; i < fields.length; ++i) {
			var index = fields[i].x;
			if (st_VALUES !== index || !withoutValue) {
				this.checkPivotFieldItems(index);
				indexes.push(index);
			}
		}
	}
	return indexes;
};
CT_pivotTableDefinition.prototype.getDataFieldIndexes = function() {
	var dataIndexes = [];
	var i;
	var dataFields = this.asc_getDataFields();
	if (dataFields) {
		for (i = 0; i < dataFields.length; ++i) {
			dataIndexes.push(dataFields[i].fld);
		}
	}
	return dataIndexes;
};
CT_pivotTableDefinition.prototype.getCacheFieldName = function(index) {
	return this.asc_getCacheFields()[index].name || "";
};
CT_pivotTableDefinition.prototype.getPivotFieldName = function(index) {
	return this.asc_getPivotFields()[index].name || this.getCacheFieldName(index);
};
CT_pivotTableDefinition.prototype.getDataFieldName = function(index) {
	var dataField = this.asc_getDataFields()[index];
	return dataField.name || this.getPivotFieldName(dataField.asc_getIndex());
};
CT_pivotTableDefinition.prototype.getPageFieldName = function(index) {
	var pageField = this.asc_getPageFields()[index];
	return pageField.name || this.getPivotFieldName(pageField.asc_getIndex());
};
CT_pivotTableDefinition.prototype.setRowItems = function(rowItems, addToHistory) {
	if (addToHistory) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, AscCH.historyitem_PivotTable_RowItems,
			this.worksheet ? this.worksheet.getId() : null, null,
			new AscCommonExcel.UndoRedoData_PivotTable(this.Get_Id(), this.rowItems, rowItems));
	}
	this.rowItems = rowItems;
};
CT_pivotTableDefinition.prototype.setColItems = function(colItems, addToHistory) {
	if (addToHistory) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, AscCH.historyitem_PivotTable_ColItems,
			this.worksheet ? this.worksheet.getId() : null, null,
			new AscCommonExcel.UndoRedoData_PivotTable(this.Get_Id(), this.colItems, colItems));
	}
	this.colItems = colItems;
};
CT_pivotTableDefinition.prototype.updateRowColItems = function () {
	var dataRow, pivotFields, rowFields, colFields, dataFields, cacheRecords, indexValues;
	cacheRecords = this.getRecords();
	if(!cacheRecords){
		return;
	}
	var rowIndexes = this.getRowColIndexes(this.asc_getRowFields(),true);
	var colIndexes = this.getRowColIndexes(this.asc_getColumnFields(),true);
	dataRow = cacheRecords.getDataMap(this.asc_getCacheFields(), this.getFilterMaps(), rowIndexes, colIndexes, this.getDataFieldIndexes());
	pivotFields = this.asc_getPivotFields();
	rowFields = this.asc_getRowFields();
	colFields = this.asc_getColumnFields();
	dataFields = this.asc_getDataFields();
	var rowItems = null;
	var colItems = null;
	if (rowFields) {
		rowItems = new CT_rowItems();
		indexValues = this.getRowFieldsValuesIndex();
		this._updateRowColItemsRecursively(0, dataRow, undefined, rowItems.i, rowFields, false, pivotFields, 0, dataFields, indexValues, false);
		this._updateRowColItemsGrandTotal(this.rowGrandTotals, indexValues, rowItems.i, rowFields, dataFields);
	}
	if (colFields) {
		colItems = new CT_colItems();
		indexValues = this.getColumnFieldsValuesIndex();
		this._updateRowColItemsRecursively(0, {vals: dataRow.subtotal}, undefined, colItems.i, colFields, true, pivotFields, 0, dataFields, indexValues, false);
		this._updateRowColItemsGrandTotal(this.colGrandTotals, indexValues, colItems.i, colFields, dataFields);
	}
	if (rowFields || colFields || dataFields) {
		if (!(rowItems && rowItems.i.length > 0)) {
			rowItems = new CT_rowItems();
			rowItems.i.push(new CT_I());
		}
		if (!(colItems && colItems.i.length > 0)) {
			colItems = new CT_colItems();
			colItems.i.push(new CT_I());
		}
	}
	this.setRowItems(rowItems, true);
	this.setColItems(colItems, true);
	return dataRow;
};
CT_pivotTableDefinition.prototype._updateRowColItemsRecursively = function(index, dataMap, parentI, items, fields, isCol, pivotFields, dataIndex, dataFields, indexValues, showAll) {
	if (index >= fields.length) {
		return;
	}
	var pivotField, indexItem, item, subDataMap, dataField;
	var x = fields[index].x;
	if (st_VALUES === x) {
		if (dataFields) {
			for (indexItem = 0; indexItem < dataFields.length; ++indexItem) {
				dataField = dataFields[indexItem];
				if (dataField) {
					pivotField = pivotFields[dataField.asc_getIndex()];
					if (pivotField) {
						this._updateRowColItemsRecursivelyElem(index, dataMap, items, fields, isCol, pivotField, pivotFields, indexItem, dataFields, indexItem, parentI, indexValues, showAll);
						parentI = null;
					}
				}
			}
		}
	} else {
		pivotField = pivotFields[x];
		if (pivotField && pivotField.items) {
			for (indexItem = 0; indexItem < pivotField.items.item.length; ++indexItem) {
				item = pivotField.items.item[indexItem];
				if (Asc.c_oAscItemType.Data === item.t) {
					subDataMap = dataMap.vals[item.x];
					if(!subDataMap && (showAll || pivotField.showAll)){
						showAll = showAll || pivotField.showAll;
						subDataMap = new PivotDataElem(dataFields.length);
					} else {
						showAll = false;
					}
					if (subDataMap) {
						this._updateRowColItemsRecursivelyElem(index, subDataMap, items, fields, isCol, pivotField, pivotFields, dataIndex, dataFields, indexItem, parentI, indexValues, showAll);
						parentI = null;
					}
				}
			}
		}
	}
};
CT_pivotTableDefinition.prototype._updateRowColItemsRecursivelyElem = function(index, dataMap, items, fields, isCol, parentPivotField, pivotFields, dataIndex, dataFields, indexItem, parentI, indexValues, showAll) {
	var newI, newParentI, i, j;
	var newX = new CT_X();
	newX.v = indexItem;
	var isTabular = isCol || !parentPivotField.outline;
	if (parentI) {
		parentI.x.push(newX);
		newParentI = isTabular ? parentI : undefined;
	} else {
		newI = new CT_I();
		newI.i = dataIndex;
		newI.r = index;
		newI.x.push(newX);
		items.push(newI);
		newParentI = isTabular ? newI : undefined;
	}
	this._updateRowColItemsRecursively(index + 1, dataMap, newParentI, items,
			fields,	isCol, pivotFields, dataIndex, dataFields, indexValues, showAll);
	var subtotals;
	var subtotalTop = true;
	var x = fields[index].x;
	if (st_VALUES !== x) {
		subtotals = pivotFields[x] && pivotFields[x].asc_getSubtotals(true);
		subtotalTop = pivotFields[x] && pivotFields[x].subtotalTop;
	}
	if (subtotals && subtotals.length > 0 && index !== indexValues && (index < fields.length - 2 || index === fields.length - 2 && index + 1 !== indexValues)) {
		var from = dataIndex;
		var to = dataIndex;
		if (index < indexValues && dataFields) {
			from = 0;
			to = dataFields.length - 1;
		}
		if (subtotals.length + to - from + 1 > 2 || isTabular || !subtotalTop) {
			for (i = 0; i < subtotals.length; ++i) {
				for (j = from; j <= to; ++j) {
					newX = new CT_X();
					newX.v = indexItem;
					newI = new CT_I();
					newI.i = j;
					newI.t = subtotals[i];
					newI.r = index;
					newI.x.push(newX);
					items.push(newI);
				}
			}
		}
	}
	if (!isCol && parentPivotField.insertBlankRow && index !== indexValues && index < fields.length - 1) {
		newX = new CT_X();
		newX.v = indexItem;
		newI = new CT_I();
		newI.t = Asc.c_oAscItemType.Blank;
		newI.r = index;
		newI.x.push(newX);
		items.push(newI);
	}
};
CT_pivotTableDefinition.prototype._updateRowColItemsGrandTotal = function(grandTotals, indexValues, items, fields, dataFields) {
	if (grandTotals && !(indexValues >= 0 && 1 === fields.length)) {
		var grandTotalsCount = indexValues >= 0 && dataFields ? dataFields.length : 1;
		for (var i = 0; i < grandTotalsCount; ++i) {
			var newI = new CT_I();
			newI.i = i;
			newI.t = Asc.c_oAscItemType.Grand;
			newI.x.push(new CT_X());
			items.push(newI);
		}
	}
};
CT_pivotTableDefinition.prototype.updateAfterEdit = function() {
	var dataRow = this.updateRowColItems();
	this.updateLocation();
	return dataRow;
};
CT_pivotTableDefinition.prototype.setLocation = function(location, addToHistory) {
	if (addToHistory) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, AscCH.historyitem_PivotTable_Location,
					this.worksheet ? this.worksheet.getId() : null, null,
					new AscCommonExcel.UndoRedoData_PivotTable(this.Get_Id(), this.location, location));
	}
	this.location = location;
	this.setChanged(false, true);
};
CT_pivotTableDefinition.prototype.setOffset = function(offset, addToHistory) {
	var location = this.location.clone();
	location.ref.setOffset(offset);
	this.setLocation(location, addToHistory);
};
CT_pivotTableDefinition.prototype.getPageFieldSize = function() {
	var res = new AscCommon.CellBase(0, 0);
	var pageField = this.asc_getPageFields();
	if (pageField) {
		var len = pageField.length;
		if (this.pageWrap) {
			res.row = this.pageOverThenDown ? Math.ceil(len / this.pageWrap) : Math.min(this.pageWrap, len);
			res.col = this.pageOverThenDown ? Math.min(this.pageWrap, len) : Math.ceil(len / this.pageWrap);
		} else {
			res.row = this.pageOverThenDown ? 1 : len;
			res.col = this.pageOverThenDown ? len : 1;
		}
	}
	return res;
};
CT_pivotTableDefinition.prototype.updateLocation = function() {
	//todo showHeaders
	var i;
	var location = this.location.clone();
	var pivotFields = this.asc_getPivotFields();
	var rowFields = this.asc_getRowFields();
	var colFields = this.asc_getColumnFields();
	var dataFields = this.asc_getDataFields();
	if (rowFields || colFields || dataFields) {
		location.firstHeaderRow = 1;
		var colFieldsCount = 0;
		if (colFields) {
			colFieldsCount = colFields.length;
			if (1 === colFieldsCount && st_VALUES === colFields[0].asc_getIndex()) {
				location.firstHeaderRow = 0;
			}
		}
		location.firstDataRow = location.firstHeaderRow + colFieldsCount;
		location.firstDataCol = 0;
		if (rowFields) {
			location.firstDataCol = 1;
			for (i = 0; i < rowFields.length - 1; ++i) {
				var index = rowFields[i].asc_getIndex();
				if (st_VALUES !== index) {
					var field = pivotFields[index];
					if (field && !(field.compact && field.outline)) {
						location.firstDataCol++;
					}
				} else {
					if (!(this.compact && this.outline)) {
						location.firstDataCol++;
					}
				}
			}
		} else if (colFields && dataFields && 1 === dataFields.length) {
			location.firstDataCol = 1;
		}
		var rowItemsCount = (rowFields || dataFields) ? this.rowItems.i.length : 0;
		var colItemsCount = (colFields || dataFields) ? this.colItems.i.length : 0;
		if (!this.compact) {
			colItemsCount = Math.max(colItemsCount, colFieldsCount);
		}
		location.ref.r2 = location.ref.r1 + location.firstDataRow + rowItemsCount - 1;
		location.ref.c2 = location.ref.c1 + location.firstDataCol + colItemsCount - 1;
	} else if (this.asc_getPageFields()) {
		location.ref.r2 = location.ref.r1;
		location.ref.c2 = location.ref.c1;
		location.firstHeaderRow = 0;
		location.firstDataRow = 0;
		location.firstDataCol = 0;
	} else {
		location.ref.r2 = location.ref.r1 + NEW_PIVOT_LAST_ROW_OFFSET;
		location.ref.c2 = location.ref.c1 + NEW_PIVOT_LAST_COL_OFFSET;
		location.firstHeaderRow = 1;
		location.firstDataRow = 1;
		location.firstDataCol = 0;
	}
	var pageFieldOffset = this.getPageFieldSize();
	if (location.ref.r1 - pageFieldOffset.row - 1 < 0) {
		location.ref.setOffset(new AscCommon.CellBase(-location.ref.r1 + pageFieldOffset.row + 1, 0));
	}
	this.setLocation(location, true);
};
CT_pivotTableDefinition.prototype.checkInsertIndex = function (insertIndex, deleteIndex) {
	if(undefined !== insertIndex && insertIndex >= deleteIndex){
		insertIndex--;
	}
	return insertIndex;
}
CT_pivotTableDefinition.prototype.removeNoDataField = function (pivotIndex, addToHistory) {
	var deleteIndex, t = this;
	var pivotField = this.asc_getPivotFields()[pivotIndex];
	var historyType;
	switch (pivotField.axis) {
		case c_oAscAxis.AxisRow:
			if (this.rowFields) {
				deleteIndex = this.rowFields.remove(pivotIndex);
				if (0 === this.rowFields.getCount()) {
					this.rowFields = null;
				}
			}
			historyType = AscCH.historyitem_PivotTable_RemoveRowField;
			break;
		case c_oAscAxis.AxisCol:
			if (this.colFields) {
				deleteIndex = this.colFields.remove(pivotIndex);
				if (0 === this.colFields.getCount()) {
					this.colFields = null;
				}
			}
			historyType = AscCH.historyitem_PivotTable_RemoveColField;
			break;
		case c_oAscAxis.AxisPage:
			if (this.pageFields) {
				deleteIndex = this.pageFields.remove(pivotIndex);
				if (0 === this.pageFields.getCount()) {
					this.pageFields = null;
				}
			}
			historyType = AscCH.historyitem_PivotTable_RemovePageField;
			break;
	}
	if (undefined !== deleteIndex) {
		if (addToHistory) {
			History.Add(AscCommonExcel.g_oUndoRedoPivotTables, historyType,
				this.worksheet ? this.worksheet.getId() : null, null,
				new AscCommonExcel.UndoRedoData_PivotTable(this.Get_Id(), pivotIndex, deleteIndex));
		}
		pivotField.axis = null;
		this.setChanged(true);
	}
	return deleteIndex;
};
CT_pivotTableDefinition.prototype.removeDataField = function (pivotIndex, dataIndex, addToHistory) {
	var pivotField = this.asc_getPivotFields()[pivotIndex];
	var removed;
	if (this.dataFields) {
		removed = this.dataFields.remove(pivotIndex, dataIndex);
		if (0 === this.dataFields.getCount()) {
			this.dataFields = null;
		}
	}
	pivotField.dataField = false;
	this.checkValuesField(addToHistory);
	if (addToHistory) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, AscCH.historyitem_PivotTable_RemoveDataField,
			this.worksheet ? this.worksheet.getId() : null, null,
			new AscCommonExcel.UndoRedoData_PivotTable(this.Get_Id(), pivotIndex, removed));
	}
	this.setChanged(true);
};
CT_pivotTableDefinition.prototype.parseDataRef = function(dataRef) {
	var worksheetSource = new CT_WorksheetSource();
	worksheetSource.fromDataRef(dataRef);
	return worksheetSource.getDataLocation();
};
CT_pivotTableDefinition.prototype.isValidDataRef = function(dataRef) {
	var location = this.parseDataRef(dataRef);
	if (location && location.bbox.getHeight() > 0) {
		if (location.headings) {
			return location.headings.length === location.bbox.getWidth();
		} else if (location.bbox.getHeight() > 1) {
			var header = this._prepareDataRange(location.ws, location.bbox.r1, location.bbox.c1, location.bbox.c2);
			return header.countCol === location.bbox.getWidth();
		}
	}
	return false;
};
CT_pivotTableDefinition.prototype.prepareDataRange = function(ws, range) {
	var header = this._prepareDataRange(ws, range.r1, range.c1, range.c2);
	if (header.minCol <= header.maxCol && header.countCol === header.maxCol - header.minCol + 1) {
		range.c1 = header.minCol;
		range.c2 = header.maxCol;
	} else if(range.r2 - range.r1 > 1){
		//test second row
		header = this._prepareDataRange(ws, range.r1 + 1, range.c1, range.c2);
		if (header.countCol === range.getWidth()) {
			range.r1 = range.r1 + 1;
		}
	}
};
CT_pivotTableDefinition.prototype._prepareDataRange = function(ws, row, c1, c2) {
	var res = {minCol: c2, maxCol: c1, countCol: 0};
	ws.getRange3(row, c1, row, c2)._foreachNoEmptyByCol(function(cell) {
		if (!cell.isNullTextString()) {
			res.minCol = Math.min(res.minCol, cell.nCol);
			res.maxCol = Math.max(res.maxCol, cell.nCol);
			res.countCol++;
		}
	});
	return res;
};

function CT_pivotTableDefinitionX14() {
//Attributes
	this.fillDownLabelsDefault = false;
	this.visualTotalsForSets = false;
	this.calculatedMembersInFilters = false;
	this.altText = null;
	this.altTextSummary = null;
	this.enableEdit = false;
	this.autoApply = false;
	this.allocationMethod = c_oAscAllocationMethod.EqualAllocation;
	this.weightExpression = null;
	this.hideValuesRow = false;
//Members
	//this.pivotEdits = null;
	//this.pivotChanges = null;
	//this.conditionalFormats = null;
}

CT_pivotTableDefinitionX14.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["fillDownLabelsDefault"];
		if (undefined !== val) {
			this.fillDownLabelsDefault = AscCommon.getBoolFromXml(val);
		}
		val = vals["visualTotalsForSets"];
		if (undefined !== val) {
			this.visualTotalsForSets = AscCommon.getBoolFromXml(val);
		}
		val = vals["calculatedMembersInFilters"];
		if (undefined !== val) {
			this.calculatedMembersInFilters = AscCommon.getBoolFromXml(val);
		}
		val = vals["altText"];
		if (undefined !== val) {
			this.altText = AscCommon.unleakString(uq(val));
		}
		val = vals["altTextSummary"];
		if (undefined !== val) {
			this.altTextSummary = AscCommon.unleakString(uq(val));
		}
		val = vals["enableEdit"];
		if (undefined !== val) {
			this.enableEdit = AscCommon.getBoolFromXml(val);
		}
		val = vals["autoApply"];
		if (undefined !== val) {
			this.autoApply = AscCommon.getBoolFromXml(val);
		}
		val = vals["allocationMethod"];
		if (undefined !== val) {
			val = FromXml_ST_AllocationMethod(val);
			if (-1 !== val) {
				this.allocationMethod = val;
			}
		}
		val = vals["weightExpression"];
		if (undefined !== val) {
			this.weightExpression = AscCommon.unleakString(uq(val));
		}
		val = vals["hideValuesRow"];
		if (undefined !== val) {
			this.hideValuesRow = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_pivotTableDefinitionX14.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pivotTableDefinition" === elem) {
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
	} else {
		newContext = null;
	}
	return newContext;
};
CT_pivotTableDefinitionX14.prototype.toXml = function(writer) {
	writer.WriteXmlNodeStart("x14:pivotTableDefinition");
	writer.WriteXmlString(" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"");
	if (false !== this.fillDownLabelsDefault) {
		writer.WriteXmlAttributeBool("fillDownLabelsDefault", this.fillDownLabelsDefault);
	}
	if (false !== this.visualTotalsForSets) {
		writer.WriteXmlAttributeBool("visualTotalsForSets", this.visualTotalsForSets);
	}
	if (false !== this.calculatedMembersInFilters) {
		writer.WriteXmlAttributeBool("calculatedMembersInFilters", this.calculatedMembersInFilters);
	}
	if (null !== this.altText) {
		writer.WriteXmlAttributeStringEncode("altText", this.altText);
	}
	if (null !== this.altTextSummary) {
		writer.WriteXmlAttributeStringEncode("altTextSummary", this.altTextSummary);
	}
	if (false !== this.enableEdit) {
		writer.WriteXmlAttributeBool("enableEdit", this.enableEdit);
	}
	if (false !== this.autoApply) {
		writer.WriteXmlAttributeBool("autoApply", this.autoApply);
	}
	if (c_oAscAllocationMethod.EqualAllocation !== this.allocationMethod) {
		writer.WriteXmlAttributeStringEncode("allocationMethod", ToXml_ST_AllocationMethod(this.allocationMethod));
	}
	if (null !== this.weightExpression) {
		writer.WriteXmlAttributeStringEncode("weightExpression", this.weightExpression);
	}
	if (false !== this.hideValuesRow) {
		writer.WriteXmlAttributeBool("hideValuesRow", this.hideValuesRow);
	}
	writer.WriteXmlNodeEnd("pivotTableDefinition", true, true);
};
function CT_CacheSource() {
//Attributes
	this.type = null;
	this.connectionId = 0;
//Members
	this.consolidation = null;
	this.extLst = null;
	this.worksheetSource = null;
}
CT_CacheSource.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["type"];
		if (undefined !== val) {
			val = FromXml_ST_SourceType(val);
			if (-1 !== val) {
				this.type = val;
			}
		}
		val = vals["connectionId"];
		if (undefined !== val) {
			this.connectionId = val - 0;
		}
	}
};
CT_CacheSource.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("consolidation" === elem) {
		newContext = new CT_Consolidation();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.consolidation = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else if ("worksheetSource" === elem) {
		newContext = new CT_WorksheetSource();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.worksheetSource = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_CacheSource.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.type) {
		writer.WriteXmlAttributeStringEncode("type", ToXml_ST_SourceType(this.type));
	}
	if (0 !== this.connectionId) {
		writer.WriteXmlAttributeNumber("connectionId", this.connectionId);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.consolidation) {
		this.consolidation.toXml(writer, "consolidation");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	if (null !== this.worksheetSource) {
		this.worksheetSource.toXml(writer, "worksheetSource");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_CacheFields() {
//Attributes
//	this.count = null;
//Members
	this.cacheField = [];
}
CT_CacheFields.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("cacheField" === elem) {
		newContext = new CT_CacheField();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.cacheField.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_CacheFields.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.cacheField.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.cacheField.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.cacheField.length; ++i) {
		var elem = this.cacheField[i];
		elem.toXml(writer, "cacheField");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_CacheHierarchies() {
//Attributes
// 	this.count = null;
//Members
	this.cacheHierarchy = [];
}
CT_CacheHierarchies.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("cacheHierarchy" === elem) {
		newContext = new CT_CacheHierarchy();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.cacheHierarchy.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_CacheHierarchies.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.cacheHierarchy.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.cacheHierarchy.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.cacheHierarchy.length; ++i) {
		var elem = this.cacheHierarchy[i];
		elem.toXml(writer, "cacheHierarchy");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_PCDKPIs() {
//Attributes
//	this.count = null;
//Members
	this.kpi = [];
}
CT_PCDKPIs.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("kpi" === elem) {
		newContext = new CT_PCDKPI();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.kpi.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PCDKPIs.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.kpi.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.kpi.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.kpi.length; ++i) {
		var elem = this.kpi[i];
		elem.toXml(writer, "kpi");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_TupleCache() {
//Members
	this.entries = null;
	this.sets = null;
	this.queryCache = null;
	this.serverFormats = null;
	this.extLst = null;
}
CT_TupleCache.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("entries" === elem) {
		newContext = new CT_PCDSDTCEntries();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.entries = newContext;
	} else if ("sets" === elem) {
		newContext = new CT_Sets();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.sets = newContext;
	} else if ("queryCache" === elem) {
		newContext = new CT_QueryCache();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.queryCache = newContext;
	} else if ("serverFormats" === elem) {
		newContext = new CT_ServerFormats();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.serverFormats = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_TupleCache.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.entries) {
		this.entries.toXml(writer, "entries");
	}
	if (null !== this.sets) {
		this.sets.toXml(writer, "sets");
	}
	if (null !== this.queryCache) {
		this.queryCache.toXml(writer, "queryCache");
	}
	if (null !== this.serverFormats) {
		this.serverFormats.toXml(writer, "serverFormats");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_CalculatedItems() {
//Attributes
//	this.count = null;
//Members
	this.calculatedItem = [];
}
CT_CalculatedItems.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("calculatedItem" === elem) {
		newContext = new CT_CalculatedItem();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.calculatedItem.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_CalculatedItems.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.calculatedItem.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.calculatedItem.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.calculatedItem.length; ++i) {
		var elem = this.calculatedItem[i];
		elem.toXml(writer, "calculatedItem");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_CalculatedMembers() {
//Attributes
//	this.count = null;
//Members
	this.calculatedMember = [];
}
CT_CalculatedMembers.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("calculatedMember" === elem) {
		newContext = new CT_CalculatedMember();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.calculatedMember.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_CalculatedMembers.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.calculatedMember.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.calculatedMember.length > 0);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.calculatedMember.length; ++i) {
		var elem = this.calculatedMember[i];
		elem.toXml(writer, "calculatedMember");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_Dimensions() {
//Attributes
//	this.count = null;
//Members
	this.dimension = [];
}
CT_Dimensions.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("dimension" === elem) {
		newContext = new CT_PivotDimension();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.dimension.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Dimensions.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.dimension.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.dimension.length > 0);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.dimension.length; ++i) {
		var elem = this.dimension[i];
		elem.toXml(writer, "dimension");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_MeasureGroups() {
//Attributes
//	this.count = null;
//Members
	this.measureGroup = [];
}
CT_MeasureGroups.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("measureGroup" === elem) {
		newContext = new CT_MeasureGroup();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.measureGroup.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_MeasureGroups.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.measureGroup.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.measureGroup.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.measureGroup.length; ++i) {
		var elem = this.measureGroup[i];
		elem.toXml(writer, "measureGroup");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_MeasureDimensionMaps() {
//Attributes
//	this.count = null;
//Members
	this.map = [];
}
CT_MeasureDimensionMaps.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("map" === elem) {
		newContext = new CT_MeasureDimensionMap();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.map.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_MeasureDimensionMaps.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.map.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.map.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.map.length; ++i) {
		var elem = this.map[i];
		elem.toXml(writer, "map");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_ExtensionList() {
//Members
	this.ext = [];
}
CT_ExtensionList.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("ext" === elem) {
		newContext = new CT_Extension();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.ext.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_ExtensionList.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.ext.length; ++i) {
		var elem = this.ext[i];
		elem.toXml(writer, "ext");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_Boolean() {
//Attributes
	this.v = null;
	this.u = null;
	this.f = null;
	this.c = null;
	this.cp = null;
//Members
	this.x = [];
}
CT_Boolean.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["v"];
		if (undefined !== val) {
			this.v = AscCommon.getBoolFromXml(val);
		}
		val = vals["u"];
		if (undefined !== val) {
			this.u = AscCommon.getBoolFromXml(val);
		}
		val = vals["f"];
		if (undefined !== val) {
			this.f = AscCommon.getBoolFromXml(val);
		}
		val = vals["c"];
		if (undefined !== val) {
			this.c = AscCommon.unleakString(uq(val));
		}
		val = vals["cp"];
		if (undefined !== val) {
			this.cp = val - 0;
		}
	}
};
CT_Boolean.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("x" === elem) {
		newContext = new CT_X();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.x.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Boolean.prototype.toXml = function(writer, name) {
	this.toXml2(writer, name, this.v, this);
};
CT_Boolean.prototype.toXml2 = function(writer, name, val, obj) {
	writer.WriteXmlNodeStart(name);
	if (null !== val) {
		writer.WriteXmlAttributeBool("v", val);
	}
	if (obj) {
		if (null !== obj.u) {
			writer.WriteXmlAttributeBool("u", obj.u);
		}
		if (null !== obj.f) {
			writer.WriteXmlAttributeBool("f", obj.f);
		}
		if (null !== obj.c) {
			writer.WriteXmlAttributeStringEncode("c", obj.c);
		}
		if (null !== obj.cp) {
			writer.WriteXmlAttributeNumber("cp", obj.cp);
		}
		writer.WriteXmlNodeEnd(name, true);
		for (var i = 0; i < obj.x.length; ++i) {
			var elem = obj.x[i];
			elem.toXml(writer, "x");
		}
	} else {
		writer.WriteXmlNodeEnd(name, true);
	}
	writer.WriteXmlNodeEnd(name);
};
CT_Boolean.prototype.isSimpleValue = function() {
	return null === this.u && null === this.f && null === this.c && null === this.cp && 0 === this.x.length;
};
CT_Boolean.prototype.clean = function() {
	this.v = null;
	this.u = null;
	this.f = null;
	this.c = null;
	this.cp = null;
	this.x = [];
};
function CT_DateTime() {
//Attributes
	this.v = null;
	this.u = null;
	this.f = null;
	this.c = null;
	this.cp = null;
//Members
	this.x = [];
}
CT_DateTime.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["v"];
		if (undefined !== val) {
			var d = new Asc.cDate(uq(val));
			this.v = new Asc.cDate(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(),
				d.getSeconds(),	d.getMilliseconds())).getExcelDateWithTime();
		}
		val = vals["u"];
		if (undefined !== val) {
			this.u = AscCommon.getBoolFromXml(val);
		}
		val = vals["f"];
		if (undefined !== val) {
			this.f = AscCommon.getBoolFromXml(val);
		}
		val = vals["c"];
		if (undefined !== val) {
			this.c = AscCommon.unleakString(uq(val));
		}
		val = vals["cp"];
		if (undefined !== val) {
			this.cp = val - 0;
		}
	}
};
CT_DateTime.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("x" === elem) {
		newContext = new CT_X();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.x.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_DateTime.prototype.toXml = function(writer, name) {
	this.toXml2(writer, name, this.v, this);
};
CT_DateTime.prototype.toXml2 = function(writer, name, val, obj) {
	writer.WriteXmlNodeStart(name);
	if (null !== val) {
		writer.WriteXmlAttributeStringEncode("v", Asc.cDate.prototype.getDateFromExcelWithTime(val).toISOString().slice(0, 19));
	}
	if (obj) {
		if (null !== obj.u) {
			writer.WriteXmlAttributeBool("u", obj.u);
		}
		if (null !== obj.f) {
			writer.WriteXmlAttributeBool("f", obj.f);
		}
		if (null !== obj.c) {
			writer.WriteXmlAttributeStringEncode("c", obj.c);
		}
		if (null !== obj.cp) {
			writer.WriteXmlAttributeNumber("cp", obj.cp);
		}
		writer.WriteXmlNodeEnd(name, true);
		for (var i = 0; i < obj.x.length; ++i) {
			var elem = obj.x[i];
			elem.toXml(writer, "x");
		}
	} else {
		writer.WriteXmlNodeEnd(name, true);
	}

	writer.WriteXmlNodeEnd(name);
};
CT_DateTime.prototype.isSimpleValue = function() {
	return null === this.u && null === this.f && null === this.c && null === this.cp && 0 === this.x.length;
};
CT_DateTime.prototype.clean = function() {
	this.v = null;
	this.u = null;
	this.f = null;
	this.c = null;
	this.cp = null;
	this.x = [];
};
function CT_Error() {
//Attributes
	this.v = null;
	this.u = null;
	this.f = null;
	this.c = null;
	this.cp = null;
	this.in = null;
	this.bc = null;
	this.fc = null;
	this.i = false;
	this.un = false;
	this.st = false;
	this.b = false;
//Members
	this.tpls = [];
	this.x = [];
}
CT_Error.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["v"];
		if (undefined !== val) {
			this.v =  AscCommonExcel.cError.prototype.getErrorTypeFromString(uq(val));
		}
		val = vals["u"];
		if (undefined !== val) {
			this.u = AscCommon.getBoolFromXml(val);
		}
		val = vals["f"];
		if (undefined !== val) {
			this.f = AscCommon.getBoolFromXml(val);
		}
		val = vals["c"];
		if (undefined !== val) {
			this.c = AscCommon.unleakString(uq(val));
		}
		val = vals["cp"];
		if (undefined !== val) {
			this.cp = val - 0;
		}
		val = vals["in"];
		if (undefined !== val) {
			this.in = val - 0;
		}
		val = vals["bc"];
		if (undefined !== val) {
			this.bc = val - 0;
		}
		val = vals["fc"];
		if (undefined !== val) {
			this.fc = val - 0;
		}
		val = vals["i"];
		if (undefined !== val) {
			this.i = AscCommon.getBoolFromXml(val);
		}
		val = vals["un"];
		if (undefined !== val) {
			this.un = AscCommon.getBoolFromXml(val);
		}
		val = vals["st"];
		if (undefined !== val) {
			this.st = AscCommon.getBoolFromXml(val);
		}
		val = vals["b"];
		if (undefined !== val) {
			this.b = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_Error.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("tpls" === elem) {
		newContext = new CT_Tuples();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.tpls.push(newContext);
	} else if ("x" === elem) {
		newContext = new CT_X();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.x.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Error.prototype.toXml = function(writer, name) {
	this.toXml2(writer, name, this.v, this);
};
CT_Error.prototype.toXml2 = function(writer, name, val, obj) {
	writer.WriteXmlNodeStart(name);
	if (null !== val) {
		writer.WriteXmlAttributeStringEncode("v", AscCommonExcel.cError.prototype.getStringFromErrorType(val));
	}
	if (obj) {
		if (null !== obj.u) {
			writer.WriteXmlAttributeBool("u", obj.u);
		}
		if (null !== obj.f) {
			writer.WriteXmlAttributeBool("f", obj.f);
		}
		if (null !== obj.c) {
			writer.WriteXmlAttributeStringEncode("c", obj.c);
		}
		if (null !== obj.cp) {
			writer.WriteXmlAttributeNumber("cp", obj.cp);
		}
		if (null !== obj.in) {
			writer.WriteXmlAttributeNumber("in", obj.in);
		}
		if (null !== obj.bc) {
			writer.WriteXmlAttributeNumber("bc", obj.bc);
		}
		if (null !== obj.fc) {
			writer.WriteXmlAttributeNumber("fc", obj.fc);
		}
		if (false !== obj.i) {
			writer.WriteXmlAttributeBool("i", obj.i);
		}
		if (false !== obj.un) {
			writer.WriteXmlAttributeBool("un", obj.un);
		}
		if (false !== obj.st) {
			writer.WriteXmlAttributeBool("st", obj.st);
		}
		if (false !== obj.b) {
			writer.WriteXmlAttributeBool("b", obj.b);
		}
		writer.WriteXmlNodeEnd(name, true);
		for (var i = 0; i < obj.tpls.length; ++i) {
			var elem = obj.tpls[i];
			elem.toXml(writer, "tpls");
		}
		for (var i = 0; i < obj.x.length; ++i) {
			var elem = obj.x[i];
			elem.toXml(writer, "x");
		}
	} else {
		writer.WriteXmlNodeEnd(name, true);
	}
	writer.WriteXmlNodeEnd(name);
};
CT_Error.prototype.isSimpleValue = function() {
	return null === this.u && null === this.f && null === this.c && null === this.cp && null === this.in &&
		null === this.bc && null === this.fc && null === this.i && null === this.un && null === this.st &&
		null === this.b && null === 0 === this.tpls.length && 0 === this.x.length;
};
CT_Error.prototype.clean = function() {
	this.v = null;
	this.u = null;
	this.f = null;
	this.c = null;
	this.cp = null;
	this.in = null;
	this.bc = null;
	this.fc = null;
	this.i = false;
	this.un = false;
	this.st = false;
	this.b = false;
	this.tpls = [];
	this.x = [];
};
function CT_Missing() {
//Attributes
	this.u = null;
	this.f = null;
	this.c = null;
	this.cp = null;
	this.in = null;
	this.bc = null;
	this.fc = null;
	this.i = false;
	this.un = false;
	this.st = false;
	this.b = false;
//Members
	this.tpls = [];
	this.x = [];
}
CT_Missing.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["u"];
		if (undefined !== val) {
			this.u = AscCommon.getBoolFromXml(val);
		}
		val = vals["f"];
		if (undefined !== val) {
			this.f = AscCommon.getBoolFromXml(val);
		}
		val = vals["c"];
		if (undefined !== val) {
			this.c = AscCommon.unleakString(uq(val));
		}
		val = vals["cp"];
		if (undefined !== val) {
			this.cp = val - 0;
		}
		val = vals["in"];
		if (undefined !== val) {
			this.in = val - 0;
		}
		val = vals["bc"];
		if (undefined !== val) {
			this.bc = val - 0;
		}
		val = vals["fc"];
		if (undefined !== val) {
			this.fc = val - 0;
		}
		val = vals["i"];
		if (undefined !== val) {
			this.i = AscCommon.getBoolFromXml(val);
		}
		val = vals["un"];
		if (undefined !== val) {
			this.un = AscCommon.getBoolFromXml(val);
		}
		val = vals["st"];
		if (undefined !== val) {
			this.st = AscCommon.getBoolFromXml(val);
		}
		val = vals["b"];
		if (undefined !== val) {
			this.b = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_Missing.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("tpls" === elem) {
		newContext = new CT_Tuples();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.tpls.push(newContext);
	} else if ("x" === elem) {
		newContext = new CT_X();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.x.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Missing.prototype.toXml = function(writer, name) {
	this.toXml2(writer, name, this);
};
CT_Missing.prototype.toXml2 = function(writer, name, obj) {
	writer.WriteXmlNodeStart(name);
	if (obj) {
		if (null !== obj.u) {
			writer.WriteXmlAttributeBool("u", obj.u);
		}
		if (null !== obj.f) {
			writer.WriteXmlAttributeBool("f", obj.f);
		}
		if (null !== obj.c) {
			writer.WriteXmlAttributeStringEncode("c", obj.c);
		}
		if (null !== obj.cp) {
			writer.WriteXmlAttributeNumber("cp", obj.cp);
		}
		if (null !== obj.in) {
			writer.WriteXmlAttributeNumber("in", obj.in);
		}
		if (null !== obj.bc) {
			writer.WriteXmlAttributeNumber("bc", obj.bc);
		}
		if (null !== obj.fc) {
			writer.WriteXmlAttributeNumber("fc", obj.fc);
		}
		if (false !== obj.i) {
			writer.WriteXmlAttributeBool("i", obj.i);
		}
		if (false !== obj.un) {
			writer.WriteXmlAttributeBool("un", obj.un);
		}
		if (false !== obj.st) {
			writer.WriteXmlAttributeBool("st", obj.st);
		}
		if (false !== obj.b) {
			writer.WriteXmlAttributeBool("b", obj.b);
		}
		writer.WriteXmlNodeEnd(name, true);
		for (var i = 0; i < obj.tpls.length; ++i) {
			var elem = obj.tpls[i];
			elem.toXml(writer, "tpls");
		}
		for (var i = 0; i < obj.x.length; ++i) {
			var elem = obj.x[i];
			elem.toXml(writer, "x");
		}
	} else {
		writer.WriteXmlNodeEnd(name, true);
	}
	writer.WriteXmlNodeEnd(name);
};
CT_Missing.prototype.isSimpleValue = function() {
	return null === this.u && null === this.f && null === this.c && null === this.cp && null === this.in &&
		null === this.bc && null === this.fc && null === this.i && null === this.un && null === this.st &&
		null === this.b && 0 === this.tpls.length && 0 === this.x.length;
};
CT_Missing.prototype.clean = function() {
	this.v = null;
	this.u = null;
	this.f = null;
	this.c = null;
	this.cp = null;
	this.in = null;
	this.bc = null;
	this.fc = null;
	this.i = false;
	this.un = false;
	this.st = false;
	this.b = false;
	this.tpls = [];
	this.x = [];
};
function CT_Number() {
//Attributes
	this.v = null;
	this.u = null;
	this.f = null;
	this.c = null;
	this.cp = null;
	this.in = null;
	this.bc = null;
	this.fc = null;
	this.i = false;
	this.un = false;
	this.st = false;
	this.b = false;
//Members
	this.tpls = [];
	this.x = [];
//internal
	this.realNumber = null;
}
CT_Number.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["v"];
		if (undefined !== val) {
			this.v = val - 0;
		}
		val = vals["u"];
		if (undefined !== val) {
			this.u = AscCommon.getBoolFromXml(val);
		}
		val = vals["f"];
		if (undefined !== val) {
			this.f = AscCommon.getBoolFromXml(val);
		}
		val = vals["c"];
		if (undefined !== val) {
			this.c = AscCommon.unleakString(uq(val));
		}
		val = vals["cp"];
		if (undefined !== val) {
			this.cp = val - 0;
		}
		val = vals["in"];
		if (undefined !== val) {
			this.in = val - 0;
		}
		val = vals["bc"];
		if (undefined !== val) {
			this.bc = val - 0;
		}
		val = vals["fc"];
		if (undefined !== val) {
			this.fc = val - 0;
		}
		val = vals["i"];
		if (undefined !== val) {
			this.i = AscCommon.getBoolFromXml(val);
		}
		val = vals["un"];
		if (undefined !== val) {
			this.un = AscCommon.getBoolFromXml(val);
		}
		val = vals["st"];
		if (undefined !== val) {
			this.st = AscCommon.getBoolFromXml(val);
		}
		val = vals["b"];
		if (undefined !== val) {
			this.b = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_Number.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("tpls" === elem) {
		newContext = new CT_Tuples();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.tpls.push(newContext);
	} else if ("x" === elem) {
		newContext = new CT_X();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.x.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Number.prototype.toXml = function(writer, name) {
	this.toXml2(writer, name, this.v, this);
};
CT_Number.prototype.toXml2 = function(writer, name, val, obj) {
	writer.WriteXmlNodeStart(name);
	if (null !== val) {
		writer.WriteXmlAttributeNumber("v", val);
	}
	if (obj) {
		if (null !== obj.u) {
			writer.WriteXmlAttributeBool("u", obj.u);
		}
		if (null !== obj.f) {
			writer.WriteXmlAttributeBool("f", obj.f);
		}
		if (null !== obj.c) {
			writer.WriteXmlAttributeStringEncode("c", obj.c);
		}
		if (null !== obj.cp) {
			writer.WriteXmlAttributeNumber("cp", obj.cp);
		}
		if (null !== obj.in) {
			writer.WriteXmlAttributeNumber("in", obj.in);
		}
		if (null !== obj.bc) {
			writer.WriteXmlAttributeNumber("bc", obj.bc);
		}
		if (null !== obj.fc) {
			writer.WriteXmlAttributeNumber("fc", obj.fc);
		}
		if (false !== obj.i) {
			writer.WriteXmlAttributeBool("i", obj.i);
		}
		if (false !== obj.un) {
			writer.WriteXmlAttributeBool("un", obj.un);
		}
		if (false !== obj.st) {
			writer.WriteXmlAttributeBool("st", obj.st);
		}
		if (false !== obj.b) {
			writer.WriteXmlAttributeBool("b", obj.b);
		}
		writer.WriteXmlNodeEnd(name, true);
		for (var i = 0; i < obj.tpls.length; ++i) {
			var elem = obj.tpls[i];
			elem.toXml(writer, "tpls");
		}
		for (var i = 0; i < obj.x.length; ++i) {
			var elem = obj.x[i];
			elem.toXml(writer, "x");
		}
	} else {
		writer.WriteXmlNodeEnd(name, true);
	}
	writer.WriteXmlNodeEnd(name);
};
CT_Number.prototype.isSimpleValue = function() {
	return null === this.u && null === this.f && null === this.c && null === this.cp && null === this.in &&
		null === this.bc && null === this.fc && null === this.i && null === this.un && null === this.st &&
		null === this.b && 0 === this.tpls.length && 0 === this.x.length;
};
CT_Number.prototype.clean = function() {
	this.v = null;
	this.u = null;
	this.f = null;
	this.c = null;
	this.cp = null;
	this.in = null;
	this.bc = null;
	this.fc = null;
	this.i = false;
	this.un = false;
	this.st = false;
	this.b = false;
	this.tpls = [];
	this.x = [];
};
function CT_String() {
//Attributes
	this.v = null;
	this.u = null;
	this.f = null;
	this.c = null;
	this.cp = null;
	this.in = null;
	this.bc = null;
	this.fc = null;
	this.i = false;
	this.un = false;
	this.st = false;
	this.b = false;
//Members
	this.tpls = [];
	this.x = [];
}
CT_String.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["v"];
		if (undefined !== val) {
			this.v = AscCommon.unleakString(uq(val));
		}
		val = vals["u"];
		if (undefined !== val) {
			this.u = AscCommon.getBoolFromXml(val);
		}
		val = vals["f"];
		if (undefined !== val) {
			this.f = AscCommon.getBoolFromXml(val);
		}
		val = vals["c"];
		if (undefined !== val) {
			this.c = AscCommon.unleakString(uq(val));
		}
		val = vals["cp"];
		if (undefined !== val) {
			this.cp = val - 0;
		}
		val = vals["in"];
		if (undefined !== val) {
			this.in = val - 0;
		}
		val = vals["bc"];
		if (undefined !== val) {
			this.bc = val - 0;
		}
		val = vals["fc"];
		if (undefined !== val) {
			this.fc = val - 0;
		}
		val = vals["i"];
		if (undefined !== val) {
			this.i = AscCommon.getBoolFromXml(val);
		}
		val = vals["un"];
		if (undefined !== val) {
			this.un = AscCommon.getBoolFromXml(val);
		}
		val = vals["st"];
		if (undefined !== val) {
			this.st = AscCommon.getBoolFromXml(val);
		}
		val = vals["b"];
		if (undefined !== val) {
			this.b = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_String.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("tpls" === elem) {
		newContext = new CT_Tuples();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.tpls.push(newContext);
	} else if ("x" === elem) {
		newContext = new CT_X();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.x.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_String.prototype.toXml = function(writer, name) {
	this.toXml2(writer, name, this.v, this);
};
CT_String.prototype.toXml2 = function(writer, name, val, obj) {
	writer.WriteXmlNodeStart(name);
	if (null !== val) {
		writer.WriteXmlAttributeStringEncode("v", val);
	}
	if (obj) {
		if (null !== obj.u) {
			writer.WriteXmlAttributeBool("u", obj.u);
		}
		if (null !== obj.f) {
			writer.WriteXmlAttributeBool("f", obj.f);
		}
		if (null !== obj.c) {
			writer.WriteXmlAttributeStringEncode("c", obj.c);
		}
		if (null !== obj.cp) {
			writer.WriteXmlAttributeNumber("cp", obj.cp);
		}
		if (null !== obj.in) {
			writer.WriteXmlAttributeNumber("in", obj.in);
		}
		if (null !== obj.bc) {
			writer.WriteXmlAttributeNumber("bc", obj.bc);
		}
		if (null !== obj.fc) {
			writer.WriteXmlAttributeNumber("fc", obj.fc);
		}
		if (false !== obj.i) {
			writer.WriteXmlAttributeBool("i", obj.i);
		}
		if (false !== obj.un) {
			writer.WriteXmlAttributeBool("un", obj.un);
		}
		if (false !== obj.st) {
			writer.WriteXmlAttributeBool("st", obj.st);
		}
		if (false !== obj.b) {
			writer.WriteXmlAttributeBool("b", obj.b);
		}
		writer.WriteXmlNodeEnd(name, true);
		for (var i = 0; i < obj.tpls.length; ++i) {
			var elem = obj.tpls[i];
			elem.toXml(writer, "tpls");
		}
		for (var i = 0; i < obj.x.length; ++i) {
			var elem = obj.x[i];
			elem.toXml(writer, "x");
		}
	} else {
		writer.WriteXmlNodeEnd(name, true);
	}
	writer.WriteXmlNodeEnd(name);
};
CT_String.prototype.isSimpleValue = function() {
	return null === this.u && null === this.f && null === this.c && null === this.cp && null === this.in &&
		null === this.bc && null === this.fc && null === this.i && null === this.un && null === this.st &&
		null === this.b && 0 === this.tpls.length && 0 === this.x.length;
};
CT_String.prototype.clean = function() {
	this.v = null;
	this.u = null;
	this.f = null;
	this.c = null;
	this.cp = null;
	this.in = null;
	this.bc = null;
	this.fc = null;
	this.i = false;
	this.un = false;
	this.st = false;
	this.b = false;
	this.tpls = [];
	this.x = [];
};
function CT_Index() {
//Attributes
	this.v = null;
}
CT_Index.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["v"];
		if (undefined !== val) {
			this.v = val - 0;
		}
	}
};
CT_Index.prototype.toXml = function(writer, name) {
	this.toXml2(writer, name, this.v);
};
CT_Index.prototype.toXml2 = function(writer, name, val) {
	writer.WriteXmlNodeStart(name);
	if (null !== val) {
		writer.WriteXmlAttributeNumber("v", val);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
CT_Index.prototype.isSimpleValue = function() {
	return true;
};
CT_Index.prototype.clean = function() {
	this.v = null;
};
function CT_Location() {
//Attributes
	this.ref = null;
	this.firstHeaderRow = null;
	this.firstDataRow = null;
	this.firstDataCol = null;
	this.rowPageCount = 0;
	this.colPageCount = 0;
}
CT_Location.prototype.clone = function() {
	var res = new CT_Location();
	res.ref = this.ref ? this.ref.clone() : this.ref;
	res.firstHeaderRow = this.firstHeaderRow;
	res.firstDataRow = this.firstDataRow;
	res.firstDataCol = this.firstDataCol;
	res.rowPageCount = this.rowPageCount;
	res.colPageCount = this.colPageCount;
	return res;
};
CT_Location.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["ref"];
		if (undefined !== val) {
			this.ref = AscCommonExcel.g_oRangeCache.getAscRange(uq(val));
		}
		val = vals["firstHeaderRow"];
		if (undefined !== val) {
			this.firstHeaderRow = val - 0;
		}
		val = vals["firstDataRow"];
		if (undefined !== val) {
			this.firstDataRow = val - 0;
		}
		val = vals["firstDataCol"];
		if (undefined !== val) {
			this.firstDataCol = val - 0;
		}
		val = vals["rowPageCount"];
		if (undefined !== val) {
			this.rowPageCount = val - 0;
		}
		val = vals["colPageCount"];
		if (undefined !== val) {
			this.colPageCount = val - 0;
		}
	}
};
CT_Location.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.ref) {
		writer.WriteXmlAttributeStringEncode("ref", this.ref.getName());
	}
	if (null !== this.firstHeaderRow) {
		writer.WriteXmlAttributeNumber("firstHeaderRow", this.firstHeaderRow);
	}
	if (null !== this.firstDataRow) {
		writer.WriteXmlAttributeNumber("firstDataRow", this.firstDataRow);
	}
	if (null !== this.firstDataCol) {
		writer.WriteXmlAttributeNumber("firstDataCol", this.firstDataCol);
	}
	if (0 !== this.rowPageCount) {
		writer.WriteXmlAttributeNumber("rowPageCount", this.rowPageCount);
	}
	if (0 !== this.colPageCount) {
		writer.WriteXmlAttributeNumber("colPageCount", this.colPageCount);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
CT_Location.prototype.intersection = function (range) {
	var t = this;
	return this.ref && (Array.isArray(range) ? range.some(function (element) {
			return t.ref.intersectionSimple(element);
		}) : this.ref.intersectionSimple(range));
};
CT_Location.prototype.contains = function (col, row) {
	return this.ref && this.ref.contains(col, row);
};
CT_Location.prototype.setPageCount = function (row, col) {
	this.rowPageCount = row;
	this.colPageCount = col;
};
CT_Location.prototype.getType = function() {
	return AscCommonExcel.UndoRedoDataTypes.PivotLocation;
};
CT_Location.prototype.Write_ToBinary2 = function(writer) {
	if (this.ref) {
		writer.WriteBool(true);
		writer.WriteLong(this.ref.c1);
		writer.WriteLong(this.ref.r1);
		writer.WriteLong(this.ref.c2);
		writer.WriteLong(this.ref.r2);
	} else {
		writer.WriteBool(false);
	}
	if (null !== this.firstHeaderRow) {
		writer.WriteBool(true);
		writer.WriteLong(this.firstHeaderRow);
	} else {
		writer.WriteBool(false);
	}
	if (null !== this.firstDataRow) {
		writer.WriteBool(true);
		writer.WriteLong(this.firstDataRow);
	} else {
		writer.WriteBool(false);
	}
	if (null !== this.firstDataCol) {
		writer.WriteBool(true);
		writer.WriteLong(this.firstDataCol);
	} else {
		writer.WriteBool(false);
	}
	writer.WriteLong(this.rowPageCount);
	writer.WriteLong(this.colPageCount);
};
CT_Location.prototype.Read_FromBinary2 = function(reader) {
	if (reader.GetBool()) {
		this.ref = new Asc.Range(reader.GetLong(), reader.GetLong(), reader.GetLong(), reader.GetLong());
	}
	if (reader.GetBool()) {
		this.firstHeaderRow = reader.GetLong();
	}
	if (reader.GetBool()) {
		this.firstDataRow = reader.GetLong();
	}
	if (reader.GetBool()) {
		this.firstDataCol = reader.GetLong();
	}
	this.rowPageCount = reader.GetLong();
	this.colPageCount = reader.GetLong();
};
function CT_PivotFields() {
//Attributes
//	this.count = null;
//Members
	this.pivotField = [];
}
CT_PivotFields.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pivotField" === elem) {
		newContext = new CT_PivotField(true);
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pivotField.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PivotFields.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.pivotField.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.pivotField.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.pivotField.length; ++i) {
		var elem = this.pivotField[i];
		elem.toXml(writer, "pivotField");
	}
	writer.WriteXmlNodeEnd(name);
};
CT_PivotFields.prototype.fillWithEmpty = function(count) {
	for (var i = 0; i < count; ++i) {
		var pivotField = new CT_PivotField(true);
		pivotField.showAll = false;
		this.pivotField.push(pivotField);
	}
};

function findFieldBase(index, array) {
	return array.findIndex(function(element) {
		return element.asc_getIndex() === index;
	});
}
function addFieldBase(index, array, val) {
	if (0 <= index && index <= array.length) {
		array.splice(index, 0, val);
	} else {
		index = array.length;
		array.push(val);
	}
	return index;
}
function removeFieldBase(index, array) {
	var deleteIndex = findFieldBase(index, array);
	if (-1 !== deleteIndex) {
		array.splice(deleteIndex, 1);
	}
	return deleteIndex;
}
function CT_RowFields() {
//Attributes
//	this.count = null;//0
//Members
	this.field = [];
}
CT_RowFields.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("field" === elem) {
		newContext = new CT_Field();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.field.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_RowFields.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.field.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.field.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.field.length; ++i) {
		var elem = this.field[i];
		elem.toXml(writer, "field");
	}
	writer.WriteXmlNodeEnd(name);
};
CT_RowFields.prototype.add = function (newContext, index) {
	return addFieldBase(index, this.field, newContext);
};
CT_RowFields.prototype.remove = function (index) {
	return removeFieldBase(index, this.field);
};
CT_RowFields.prototype.find = function (index) {
	return findFieldBase(index, this.field);
};
CT_RowFields.prototype.getCount = function () {
	return this.field.length;
};
function CT_rowItems() {
//Attributes
//	this.count = null;
//Members
	this.i = [];
}
CT_rowItems.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("i" === elem) {
		newContext = new CT_I();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.i.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_rowItems.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.i.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.i.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.i.length; ++i) {
		var elem = this.i[i];
		elem.toXml(writer, "i");
	}
	writer.WriteXmlNodeEnd(name);
};
CT_rowItems.prototype.getType = function() {
	return AscCommonExcel.UndoRedoDataTypes.PivotRowItems;
};
CT_rowItems.prototype.Write_ToBinary2 = function(writer) {
	writer.WriteLong(this.i.length);
	for (var i = 0; i < this.i.length; ++i) {
		var item = this.i[i];
		writer.WriteByte(item.t);
		writer.WriteLong(item.r);
		writer.WriteLong(item.i);
		writer.WriteLong(item.x.length);
		for (var j = 0; j < item.x.length; ++j) {
			writer.WriteLong(item.x[j].v);
		}
	}
};
CT_rowItems.prototype.Read_FromBinary2 = function(reader) {
	var len = reader.GetLong();
	while (--len >= 0) {
		var newI = new CT_I();
		newI.t = reader.GetByte();
		newI.r = reader.GetLong();
		newI.i = reader.GetLong();
		var len2 = reader.GetLong();
		while (--len2 >= 0) {
			var newX = new CT_X();
			newX.v = reader.GetLong();
			newI.x.push(newX);
		}
		this.i.push(newI);
	}
};

function CT_ColFields() {
//Attributes
//	this.count = null;//0
//Members
	this.field = [];
}
CT_ColFields.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("field" === elem) {
		newContext = new CT_Field();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.field.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_ColFields.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.field.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.field.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.field.length; ++i) {
		var elem = this.field[i];
		elem.toXml(writer, "field");
	}
	writer.WriteXmlNodeEnd(name);
};
CT_ColFields.prototype.add = function (newContext, index) {
	return addFieldBase(index, this.field, newContext);
};
CT_ColFields.prototype.remove = function (index) {
	return removeFieldBase(index, this.field);
};
CT_ColFields.prototype.find = function (index) {
	return findFieldBase(index, this.field);
};
CT_ColFields.prototype.getCount = function () {
	return this.field.length;
};
function CT_colItems() {
//Attributes
//	this.count = null;
//Members
	this.i = [];
}
CT_colItems.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("i" === elem) {
		newContext = new CT_I();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.i.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_colItems.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.i.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.i.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.i.length; ++i) {
		var elem = this.i[i];
		elem.toXml(writer, "i");
	}
	writer.WriteXmlNodeEnd(name);
};
CT_colItems.prototype.getType = function() {
	return AscCommonExcel.UndoRedoDataTypes.PivotColItems;
};
CT_colItems.prototype.Write_ToBinary2 = CT_rowItems.prototype.Write_ToBinary2;
CT_colItems.prototype.Read_FromBinary2 = CT_rowItems.prototype.Read_FromBinary2;
function CT_PageFields() {
//Attributes
//	this.count = null;
//Members
	this.pageField = [];
}
CT_PageFields.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pageField" === elem) {
		newContext = new CT_PageField();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pageField.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PageFields.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.pageField.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.pageField.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.pageField.length; ++i) {
		var elem = this.pageField[i];
		elem.toXml(writer, "pageField");
	}
	writer.WriteXmlNodeEnd(name);
};
CT_PageFields.prototype.add = function (newContext, index) {
	return addFieldBase(index, this.pageField, newContext);
};
CT_PageFields.prototype.remove = function (index) {
	return removeFieldBase(index, this.pageField);
};
CT_PageFields.prototype.getCount = function () {
	return this.pageField.length;
};
function CT_DataFields() {
//Attributes
//	this.count = null;
//Members
	this.dataField = [];
}
CT_DataFields.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("dataField" === elem) {
		newContext = new CT_DataField(true);
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.dataField.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_DataFields.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.dataField.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.dataField.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.dataField.length; ++i) {
		var elem = this.dataField[i];
		elem.toXml(writer, "dataField");
	}
	writer.WriteXmlNodeEnd(name);
};
CT_DataFields.prototype.add = function (newContext, index) {
	return addFieldBase(index, this.dataField, newContext);
};
CT_DataFields.prototype.remove = function(index, dataIndex) {
	if (0 <= dataIndex && dataIndex < this.dataField.length) {
		this.dataField.splice(dataIndex, 1);
		return [dataIndex];
	} else {
		var res = [];
		var deleteIndex = removeFieldBase(index, this.dataField);
		while (-1 !== deleteIndex) {
			res.push(deleteIndex);
			deleteIndex = removeFieldBase(index, this.dataField);
		}
		return res;
	}
};
CT_DataFields.prototype.getCount = function () {
	return this.dataField.length;
};
CT_DataFields.prototype.checkDuplicateName = function(name) {
	for (var i = 0; i < this.dataField.length; ++i) {
		if (name === this.dataField[i].name) {
			return true;
		}
	}
	return false;
};

function CT_Formats() {
//Attributes
//	this.count = null;//0
//Members
	this.format = [];
}
CT_Formats.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("format" === elem) {
		newContext = new CT_Format();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.format.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Formats.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.format.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.format.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.format.length; ++i) {
		var elem = this.format[i];
		elem.toXml(writer, "format");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_ConditionalFormats() {
//Attributes
//	this.count = null;//0
//Members
	this.conditionalFormat = [];
}
CT_ConditionalFormats.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("conditionalFormat" === elem) {
		newContext = new CT_ConditionalFormat();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.conditionalFormat.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_ConditionalFormats.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.conditionalFormat.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.conditionalFormat.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.conditionalFormat.length; ++i) {
		var elem = this.conditionalFormat[i];
		elem.toXml(writer, "conditionalFormat");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_ChartFormats() {
//Attributes
//	this.count = null;//0
//Members
	this.chartFormat = [];
}
CT_ChartFormats.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("chartFormat" === elem) {
		newContext = new CT_ChartFormat();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.chartFormat.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_ChartFormats.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.chartFormat.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.chartFormat.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.chartFormat.length; ++i) {
		var elem = this.chartFormat[i];
		elem.toXml(writer, "chartFormat");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_PivotHierarchies() {
//Attributes
//	this.count = null;
//Members
	this.pivotHierarchy = [];
}
CT_PivotHierarchies.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pivotHierarchy" === elem) {
		newContext = new CT_PivotHierarchy();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pivotHierarchy.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PivotHierarchies.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.pivotHierarchy.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.pivotHierarchy.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.pivotHierarchy.length; ++i) {
		var elem = this.pivotHierarchy[i];
		elem.toXml(writer, "pivotHierarchy");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_PivotTableStyle() {
//Attributes
	this.name = null;
	this.showRowHeaders = null;
	this.showColHeaders = null;
	this.showRowStripes = null;
	this.showColStripes = null;
	this.showLastColumn = null;
}
CT_PivotTableStyle.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["showRowHeaders"];
		if (undefined !== val) {
			this.showRowHeaders = AscCommon.getBoolFromXml(val);
		}
		val = vals["showColHeaders"];
		if (undefined !== val) {
			this.showColHeaders = AscCommon.getBoolFromXml(val);
		}
		val = vals["showRowStripes"];
		if (undefined !== val) {
			this.showRowStripes = AscCommon.getBoolFromXml(val);
		}
		val = vals["showColStripes"];
		if (undefined !== val) {
			this.showColStripes = AscCommon.getBoolFromXml(val);
		}
		val = vals["showLastColumn"];
		if (undefined !== val) {
			this.showLastColumn = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_PivotTableStyle.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (null !== this.showRowHeaders) {
		writer.WriteXmlAttributeBool("showRowHeaders", this.showRowHeaders);
	}
	if (null !== this.showColHeaders) {
		writer.WriteXmlAttributeBool("showColHeaders", this.showColHeaders);
	}
	if (null !== this.showRowStripes) {
		writer.WriteXmlAttributeBool("showRowStripes", this.showRowStripes);
	}
	if (null !== this.showColStripes) {
		writer.WriteXmlAttributeBool("showColStripes", this.showColStripes);
	}
	if (null !== this.showLastColumn) {
		writer.WriteXmlAttributeBool("showLastColumn", this.showLastColumn);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
CT_PivotTableStyle.prototype.set = function() {

};
CT_PivotTableStyle.prototype.asc_getName = function() {
	return this.name;
};
CT_PivotTableStyle.prototype.asc_getShowRowHeaders = function() {
	return this.showRowHeaders;
};
CT_PivotTableStyle.prototype.asc_getShowColHeaders = function() {
	return this.showColHeaders;
};
CT_PivotTableStyle.prototype.asc_getShowRowStripes = function() {
	return this.showRowStripes;
};
CT_PivotTableStyle.prototype.asc_getShowColStripes = function() {
	return this.showColStripes;
};
CT_PivotTableStyle.prototype.asc_setName = function(api, pivot, newVal) {
	if (newVal !== this.name) {
		var t = this;
		api._changePivotWithLock(pivot, function(ws) {t._setName(newVal, pivot, ws)});
	}
};
CT_PivotTableStyle.prototype.asc_setShowRowHeaders = function(api, pivot, newVal) {
	if (newVal !== this.showRowHeaders) {
		var t = this;
		api._changePivotWithLock(pivot, function(ws) {t._setShowRowHeaders(newVal, pivot, ws)});
	}
};
CT_PivotTableStyle.prototype.asc_setShowColHeaders = function(api, pivot, newVal) {
	if (newVal !== this.showColHeaders) {
		var t = this;
		api._changePivotWithLock(pivot, function(ws) {t._setShowColHeaders(newVal, pivot, ws)});
	}
};
CT_PivotTableStyle.prototype.asc_setShowRowStripes = function(api, pivot, newVal) {
	if (newVal !== this.showRowStripes) {
		var t = this;
		api._changePivotWithLock(pivot, function(ws) {t._setShowRowStripes(newVal, pivot, ws)});
	}
};
CT_PivotTableStyle.prototype.asc_setShowColStripes = function(api, pivot, newVal) {
	if (newVal !== this.showColStripes) {
		var t = this;
		api._changePivotWithLock(pivot, function(ws) {t._setShowColStripes(newVal, pivot, ws)});
	}
};
CT_PivotTableStyle.prototype._setName = function (newVal, pivot, ws) {
	if (History.Is_On() && this.name !== newVal) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, AscCH.historyitem_PivotTable_StyleName,
			ws ? ws.getId() : null, null,
			new AscCommonExcel.UndoRedoData_PivotTable(pivot && pivot.Get_Id(), this.name, newVal));
	}
	this.name = newVal;
	pivot.setChanged(false, true);
};
CT_PivotTableStyle.prototype._setShowRowHeaders = function (newVal, pivot, ws) {
	if (History.Is_On() && this.showRowHeaders !== newVal) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, AscCH.historyitem_PivotTable_StyleShowRowHeaders,
			ws ? ws.getId() : null, null,
			new AscCommonExcel.UndoRedoData_PivotTable(pivot && pivot.Get_Id(), this.showRowHeaders, newVal));
	}
	this.showRowHeaders = newVal;
	pivot.setChanged(false, true);
};
CT_PivotTableStyle.prototype._setShowColHeaders = function (newVal, pivot, ws) {
	if (History.Is_On() && this.showColHeaders !== newVal) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, AscCH.historyitem_PivotTable_StyleShowColHeaders,
			ws ? ws.getId() : null, null,
			new AscCommonExcel.UndoRedoData_PivotTable(pivot && pivot.Get_Id(), this.showColHeaders, newVal));
	}
	this.showColHeaders = newVal;
	pivot.setChanged(false, true);
};
CT_PivotTableStyle.prototype._setShowRowStripes = function (newVal, pivot, ws) {
	if (History.Is_On() && this.showRowStripes !== newVal) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, AscCH.historyitem_PivotTable_StyleShowRowStripes,
			ws ? ws.getId() : null, null,
			new AscCommonExcel.UndoRedoData_PivotTable(pivot && pivot.Get_Id(), this.showRowStripes, newVal));
	}
	this.showRowStripes = newVal;
	pivot.setChanged(false, true);
};
CT_PivotTableStyle.prototype._setShowColStripes = function (newVal, pivot, ws) {
	if (History.Is_On() && this.showColStripes !== newVal) {
		History.Add(AscCommonExcel.g_oUndoRedoPivotTables, AscCH.historyitem_PivotTable_StyleShowColStripes,
			ws ? ws.getId() : null, null,
			new AscCommonExcel.UndoRedoData_PivotTable(pivot && pivot.Get_Id(), this.showColStripes, newVal));
	}
	this.showColStripes = newVal;
	pivot.setChanged(false, true);
};
function CT_PivotFilters() {
//Attributes
//	this.count = null;//0
//Members
	this.filter = [];
}
CT_PivotFilters.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("filter" === elem) {
		newContext = new CT_PivotFilter();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.filter.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PivotFilters.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.filter.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.filter.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.filter.length; ++i) {
		var elem = this.filter[i];
		elem.toXml(writer, "filter");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_RowHierarchiesUsage() {
//Attributes
//	this.count = null;
//Members
	this.rowHierarchyUsage = [];
}
CT_RowHierarchiesUsage.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("rowHierarchyUsage" === elem) {
		newContext = new CT_HierarchyUsage();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.rowHierarchyUsage.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_RowHierarchiesUsage.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.rowHierarchyUsage.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.rowHierarchyUsage.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.rowHierarchyUsage.length; ++i) {
		var elem = this.rowHierarchyUsage[i];
		elem.toXml(writer, "rowHierarchyUsage");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_ColHierarchiesUsage() {
//Attributes
//	this.count = null;
//Members
	this.colHierarchyUsage = [];
}
CT_ColHierarchiesUsage.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("colHierarchyUsage" === elem) {
		newContext = new CT_HierarchyUsage();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.colHierarchyUsage.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_ColHierarchiesUsage.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.colHierarchyUsage.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.colHierarchyUsage.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.colHierarchyUsage.length; ++i) {
		var elem = this.colHierarchyUsage[i];
		elem.toXml(writer, "colHierarchyUsage");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_Consolidation() {
//Attributes
	this.autoPage = true;
//Members
	this.pages = null;
	this.rangeSets = null;
}
CT_Consolidation.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["autoPage"];
		if (undefined !== val) {
			this.autoPage = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_Consolidation.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pages" === elem) {
		newContext = new CT_Pages();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pages = newContext;
	} else if ("rangeSets" === elem) {
		newContext = new CT_RangeSets();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.rangeSets = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Consolidation.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (true !== this.autoPage) {
		writer.WriteXmlAttributeBool("autoPage", this.autoPage);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.pages) {
		this.pages.toXml(writer, "pages");
	}
	if (null !== this.rangeSets) {
		this.rangeSets.toXml(writer, "rangeSets");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_WorksheetSource() {
//Attributes
	this.ref = null;
	this.name = null;
	this.sheet = null;
	this.id = null;
//Private
	this.formula = null;
}
CT_WorksheetSource.prototype.onFormulaEvent = function (type, eventData) {
	if (AscCommon.c_oNotifyParentType.ChangeFormula === type) {
		// ToDo update formula with eventData.assemble;
	}
};
CT_WorksheetSource.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["ref"];
		if (undefined !== val) {
			this.ref = AscCommon.unleakString(uq(val));
		}
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["sheet"];
		if (undefined !== val) {
			this.sheet = AscCommon.unleakString(uq(val));
		}
		val = vals["r:id"];
		if (undefined !== val) {
			this.id = AscCommon.unleakString(uq(val));
		}

		var text;
		if (this.name) {
			if (this.sheet) {
				text = AscCommon.parserHelp.getEscapeSheetName(this.sheet) + "!" + this.name;
			} else {
				text = this.name;
			}
		} else if (this.ref && this.sheet) {
			text = AscCommon.parserHelp.get3DRef(this.sheet, this.ref);
		}
		if (text) {
			this.formula = new AscCommonExcel.parserFormula(text, this, AscCommonExcel.g_DefNameWorksheet);
			this.formula.parse();
			this.formula.buildDependencies();
		}
	}
};
CT_WorksheetSource.prototype.toXml = function(writer, name) {
	this._updateAttributes();
	writer.WriteXmlNodeStart(name);
	if (null !== this.ref) {
		writer.WriteXmlAttributeStringEncode("ref", this.ref);
	}
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (null !== this.sheet) {
		writer.WriteXmlAttributeStringEncode("sheet", this.sheet);
	}
	//todo
	// if (null !== this.id) {
		// writer.WriteXmlAttributeStringEncode("r:id", this.id);
	// }
	writer.WriteXmlNodeEnd(name, true, true);
};
CT_WorksheetSource.prototype.getDataLocation = function() {
	if (this.formula && 1 === this.formula.getOutStackSize()) {
		var elem = this.formula.getOutStackElem(0);
		var headings;
		if (elem.type === AscCommonExcel.cElementType.table) {
			headings = elem.geColumnHeadings();
		}
		var val = this.formula.calculate();
		if (val) {
			switch (val.type) {
				case AscCommonExcel.cElementType.cell:
				case AscCommonExcel.cElementType.cellsRange:
				case AscCommonExcel.cElementType.cell3D:
				case AscCommonExcel.cElementType.cellsRange3D:
					return {ws: val.getWS(), bbox: val.getBBox0(), headings: headings};
					break;
			}
		}
	}
};
CT_WorksheetSource.prototype.getDataRef = function() {
	if (this.formula) {
		return this.formula.getFormula();
	}
};
CT_WorksheetSource.prototype.fromDataRef = function(dataRef) {
	if (dataRef) {
		this.formula = new AscCommonExcel.parserFormula(dataRef, this, AscCommonExcel.g_DefNameWorksheet);
		this.formula.parse();
		this.formula.buildDependencies();
	}
};
CT_WorksheetSource.prototype._updateAttributes = function() {
	this.ref = null;
	this.name = null;
	this.sheet = null;
	if (this.formula && 1 === this.formula.getOutStackSize()) {
		var elem = this.formula.getOutStackElem(0);
		if (elem) {
			switch (elem.type) {
				case AscCommonExcel.cElementType.cell:
				case AscCommonExcel.cElementType.cellsRange:
				case AscCommonExcel.cElementType.cell3D:
				case AscCommonExcel.cElementType.cellsRange3D:
					this.sheet = elem.getWS().getName();
					this.ref = elem.getBBox0().getName();
					break;
				case AscCommonExcel.cElementType.name:
					this.name = elem.toString();
					break;
				case AscCommonExcel.cElementType.name3D:
					this.sheet = elem.getWS().getName();
					this.name = AscCommonExcel.cName.prototype.toString.call(elem);
					break;
				case AscCommonExcel.cElementType.table:
					//todo without '[]'
					this.name = elem.toString();
					break;
			}
		}
	}
};

function CT_CacheField() {
//Attributes
	this.name = null;
	this.caption = null;
	this.propertyName = null;
	this.serverField = false;
	this.uniqueList = true;
	this.numFmtId = null;
	this.formula = null;
	this.sqlType = 0;
	this.hierarchy = 0;
	this.level = 0;
	this.databaseField = true;
	this.mappingCount = null;
	this.memberPropertyField = false;
//Members
	this.sharedItems = null;
	this.fieldGroup = null;
	this.mpMap = [];
	this.extLst = null;
}
CT_CacheField.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["caption"];
		if (undefined !== val) {
			this.caption = AscCommon.unleakString(uq(val));
		}
		val = vals["propertyName"];
		if (undefined !== val) {
			this.propertyName = AscCommon.unleakString(uq(val));
		}
		val = vals["serverField"];
		if (undefined !== val) {
			this.serverField = AscCommon.getBoolFromXml(val);
		}
		val = vals["uniqueList"];
		if (undefined !== val) {
			this.uniqueList = AscCommon.getBoolFromXml(val);
		}
		val = vals["numFmtId"];
		if (undefined !== val) {
			this.numFmtId = val - 0;
		}
		val = vals["formula"];
		if (undefined !== val) {
			this.formula = AscCommon.unleakString(uq(val));
		}
		val = vals["sqlType"];
		if (undefined !== val) {
			this.sqlType = val - 0;
		}
		val = vals["hierarchy"];
		if (undefined !== val) {
			this.hierarchy = val - 0;
		}
		val = vals["level"];
		if (undefined !== val) {
			this.level = val - 0;
		}
		val = vals["databaseField"];
		if (undefined !== val) {
			this.databaseField = AscCommon.getBoolFromXml(val);
		}
		val = vals["mappingCount"];
		if (undefined !== val) {
			this.mappingCount = val - 0;
		}
		val = vals["memberPropertyField"];
		if (undefined !== val) {
			this.memberPropertyField = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_CacheField.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("sharedItems" === elem) {
		newContext = new CT_SharedItems();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.sharedItems = newContext;
	} else if ("fieldGroup" === elem) {
		newContext = new CT_FieldGroup();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.fieldGroup = newContext;
	} else if ("mpMap" === elem) {
		newContext = new CT_X();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.mpMap.push(newContext);
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_CacheField.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (null !== this.caption) {
		writer.WriteXmlAttributeStringEncode("caption", this.caption);
	}
	if (null !== this.propertyName) {
		writer.WriteXmlAttributeStringEncode("propertyName", this.propertyName);
	}
	if (false !== this.serverField) {
		writer.WriteXmlAttributeBool("serverField", this.serverField);
	}
	if (true !== this.uniqueList) {
		writer.WriteXmlAttributeBool("uniqueList", this.uniqueList);
	}
	if (null !== this.numFmtId) {
		writer.WriteXmlAttributeNumber("numFmtId", this.numFmtId);
	}
	if (null !== this.formula) {
		writer.WriteXmlAttributeStringEncode("formula", this.formula);
	}
	if (0 !== this.sqlType) {
		writer.WriteXmlAttributeNumber("sqlType", this.sqlType);
	}
	if (0 !== this.hierarchy) {
		writer.WriteXmlAttributeNumber("hierarchy", this.hierarchy);
	}
	if (0 !== this.level) {
		writer.WriteXmlAttributeNumber("level", this.level);
	}
	if (true !== this.databaseField) {
		writer.WriteXmlAttributeBool("databaseField", this.databaseField);
	}
	if (null !== this.mappingCount) {
		writer.WriteXmlAttributeNumber("mappingCount", this.mappingCount);
	}
	if (false !== this.memberPropertyField) {
		writer.WriteXmlAttributeBool("memberPropertyField", this.memberPropertyField);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.sharedItems) {
		this.sharedItems.toXml(writer, "sharedItems");
	}
	if (null !== this.fieldGroup) {
		this.fieldGroup.toXml(writer, "fieldGroup");
	}
	for (var i = 0; i < this.mpMap.length; ++i) {
		var elem = this.mpMap[i];
		elem.toXml(writer, "mpMap");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
CT_CacheField.prototype.asc_getName = function () {
	return this.name;
};
CT_CacheField.prototype.getSharedItem = function (index) {
	return this.sharedItems && this.sharedItems.Items.get(index);
};
CT_CacheField.prototype.getSharedSize = function () {
	return this.sharedItems && this.sharedItems.Items.getSize() || 0;
};
CT_CacheField.prototype.isSumSubtotal = function (index) {
	return this.sharedItems && false === this.sharedItems.containsSemiMixedTypes && true === this.sharedItems.containsNumber;
};

function CT_CacheHierarchy() {
//Attributes
	this.uniqueName = null;
	this.caption = null;
	this.measure = false;
	this.set = false;
	this.parentSet = null;
	this.iconSet = 0;
	this.attribute = false;
	this.time = false;
	this.keyAttribute = false;
	this.defaultMemberUniqueName = null;
	this.allUniqueName = null;
	this.allCaption = null;
	this.dimensionUniqueName = null;
	this.displayFolder = null;
	this.measureGroup = null;
	this.measures = false;
	this.count = null;
	this.oneField = false;
	this.memberValueDatatype = null;
	this.unbalanced = null;
	this.unbalancedGroup = null;
	this.hidden = false;
//Members
	this.fieldsUsage = null;
	this.groupLevels = null;
	this.extLst = null;
}
CT_CacheHierarchy.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["uniqueName"];
		if (undefined !== val) {
			this.uniqueName = AscCommon.unleakString(uq(val));
		}
		val = vals["caption"];
		if (undefined !== val) {
			this.caption = AscCommon.unleakString(uq(val));
		}
		val = vals["measure"];
		if (undefined !== val) {
			this.measure = AscCommon.getBoolFromXml(val);
		}
		val = vals["set"];
		if (undefined !== val) {
			this.set = AscCommon.getBoolFromXml(val);
		}
		val = vals["parentSet"];
		if (undefined !== val) {
			this.parentSet = val - 0;
		}
		val = vals["iconSet"];
		if (undefined !== val) {
			this.iconSet = val - 0;
		}
		val = vals["attribute"];
		if (undefined !== val) {
			this.attribute = AscCommon.getBoolFromXml(val);
		}
		val = vals["time"];
		if (undefined !== val) {
			this.time = AscCommon.getBoolFromXml(val);
		}
		val = vals["keyAttribute"];
		if (undefined !== val) {
			this.keyAttribute = AscCommon.getBoolFromXml(val);
		}
		val = vals["defaultMemberUniqueName"];
		if (undefined !== val) {
			this.defaultMemberUniqueName = AscCommon.unleakString(uq(val));
		}
		val = vals["allUniqueName"];
		if (undefined !== val) {
			this.allUniqueName = AscCommon.unleakString(uq(val));
		}
		val = vals["allCaption"];
		if (undefined !== val) {
			this.allCaption = AscCommon.unleakString(uq(val));
		}
		val = vals["dimensionUniqueName"];
		if (undefined !== val) {
			this.dimensionUniqueName = AscCommon.unleakString(uq(val));
		}
		val = vals["displayFolder"];
		if (undefined !== val) {
			this.displayFolder = AscCommon.unleakString(uq(val));
		}
		val = vals["measureGroup"];
		if (undefined !== val) {
			this.measureGroup = AscCommon.unleakString(uq(val));
		}
		val = vals["measures"];
		if (undefined !== val) {
			this.measures = AscCommon.getBoolFromXml(val);
		}
		val = vals["count"];
		if (undefined !== val) {
			this.count = val - 0;
		}
		val = vals["oneField"];
		if (undefined !== val) {
			this.oneField = AscCommon.getBoolFromXml(val);
		}
		val = vals["memberValueDatatype"];
		if (undefined !== val) {
			this.memberValueDatatype = val - 0;
		}
		val = vals["unbalanced"];
		if (undefined !== val) {
			this.unbalanced = AscCommon.getBoolFromXml(val);
		}
		val = vals["unbalancedGroup"];
		if (undefined !== val) {
			this.unbalancedGroup = AscCommon.getBoolFromXml(val);
		}
		val = vals["hidden"];
		if (undefined !== val) {
			this.hidden = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_CacheHierarchy.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("fieldsUsage" === elem) {
		newContext = new CT_FieldsUsage();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.fieldsUsage = newContext;
	} else if ("groupLevels" === elem) {
		newContext = new CT_GroupLevels();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.groupLevels = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_CacheHierarchy.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.uniqueName) {
		writer.WriteXmlAttributeStringEncode("uniqueName", this.uniqueName);
	}
	if (null !== this.caption) {
		writer.WriteXmlAttributeStringEncode("caption", this.caption);
	}
	if (false !== this.measure) {
		writer.WriteXmlAttributeBool("measure", this.measure);
	}
	if (false !== this.set) {
		writer.WriteXmlAttributeBool("set", this.set);
	}
	if (null !== this.parentSet) {
		writer.WriteXmlAttributeNumber("parentSet", this.parentSet);
	}
	if (0 !== this.iconSet) {
		writer.WriteXmlAttributeNumber("iconSet", this.iconSet);
	}
	if (false !== this.attribute) {
		writer.WriteXmlAttributeBool("attribute", this.attribute);
	}
	if (false !== this.time) {
		writer.WriteXmlAttributeBool("time", this.time);
	}
	if (false !== this.keyAttribute) {
		writer.WriteXmlAttributeBool("keyAttribute", this.keyAttribute);
	}
	if (null !== this.defaultMemberUniqueName) {
		writer.WriteXmlAttributeStringEncode("defaultMemberUniqueName", this.defaultMemberUniqueName);
	}
	if (null !== this.allUniqueName) {
		writer.WriteXmlAttributeStringEncode("allUniqueName", this.allUniqueName);
	}
	if (null !== this.allCaption) {
		writer.WriteXmlAttributeStringEncode("allCaption", this.allCaption);
	}
	if (null !== this.dimensionUniqueName) {
		writer.WriteXmlAttributeStringEncode("dimensionUniqueName", this.dimensionUniqueName);
	}
	if (null !== this.displayFolder) {
		writer.WriteXmlAttributeStringEncode("displayFolder", this.displayFolder);
	}
	if (null !== this.measureGroup) {
		writer.WriteXmlAttributeStringEncode("measureGroup", this.measureGroup);
	}
	if (false !== this.measures) {
		writer.WriteXmlAttributeBool("measures", this.measures);
	}
	if (null !== this.count) {
		writer.WriteXmlAttributeNumber("count", this.count);
	}
	if (false !== this.oneField) {
		writer.WriteXmlAttributeBool("oneField", this.oneField);
	}
	if (null !== this.memberValueDatatype) {
		writer.WriteXmlAttributeNumber("memberValueDatatype", this.memberValueDatatype);
	}
	if (null !== this.unbalanced) {
		writer.WriteXmlAttributeBool("unbalanced", this.unbalanced);
	}
	if (null !== this.unbalancedGroup) {
		writer.WriteXmlAttributeBool("unbalancedGroup", this.unbalancedGroup);
	}
	if (false !== this.hidden) {
		writer.WriteXmlAttributeBool("hidden", this.hidden);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.fieldsUsage) {
		this.fieldsUsage.toXml(writer, "fieldsUsage");
	}
	if (null !== this.groupLevels) {
		this.groupLevels.toXml(writer, "groupLevels");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_PCDKPI() {
//Attributes
	this.uniqueName = null;
	this.caption = null;
	this.displayFolder = null;
	this.measureGroup = null;
	this.parent = null;
	this.value = null;
	this.goal = null;
	this.status = null;
	this.trend = null;
	this.weight = null;
	this.time = null;
}
CT_PCDKPI.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["uniqueName"];
		if (undefined !== val) {
			this.uniqueName = AscCommon.unleakString(uq(val));
		}
		val = vals["caption"];
		if (undefined !== val) {
			this.caption = AscCommon.unleakString(uq(val));
		}
		val = vals["displayFolder"];
		if (undefined !== val) {
			this.displayFolder = AscCommon.unleakString(uq(val));
		}
		val = vals["measureGroup"];
		if (undefined !== val) {
			this.measureGroup = AscCommon.unleakString(uq(val));
		}
		val = vals["parent"];
		if (undefined !== val) {
			this.parent = AscCommon.unleakString(uq(val));
		}
		val = vals["value"];
		if (undefined !== val) {
			this.value = AscCommon.unleakString(uq(val));
		}
		val = vals["goal"];
		if (undefined !== val) {
			this.goal = AscCommon.unleakString(uq(val));
		}
		val = vals["status"];
		if (undefined !== val) {
			this.status = AscCommon.unleakString(uq(val));
		}
		val = vals["trend"];
		if (undefined !== val) {
			this.trend = AscCommon.unleakString(uq(val));
		}
		val = vals["weight"];
		if (undefined !== val) {
			this.weight = AscCommon.unleakString(uq(val));
		}
		val = vals["time"];
		if (undefined !== val) {
			this.time = AscCommon.unleakString(uq(val));
		}
	}
};
CT_PCDKPI.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.uniqueName) {
		writer.WriteXmlAttributeStringEncode("uniqueName", this.uniqueName);
	}
	if (null !== this.caption) {
		writer.WriteXmlAttributeStringEncode("caption", this.caption);
	}
	if (null !== this.displayFolder) {
		writer.WriteXmlAttributeStringEncode("displayFolder", this.displayFolder);
	}
	if (null !== this.measureGroup) {
		writer.WriteXmlAttributeStringEncode("measureGroup", this.measureGroup);
	}
	if (null !== this.parent) {
		writer.WriteXmlAttributeStringEncode("parent", this.parent);
	}
	if (null !== this.value) {
		writer.WriteXmlAttributeStringEncode("value", this.value);
	}
	if (null !== this.goal) {
		writer.WriteXmlAttributeStringEncode("goal", this.goal);
	}
	if (null !== this.status) {
		writer.WriteXmlAttributeStringEncode("status", this.status);
	}
	if (null !== this.trend) {
		writer.WriteXmlAttributeStringEncode("trend", this.trend);
	}
	if (null !== this.weight) {
		writer.WriteXmlAttributeStringEncode("weight", this.weight);
	}
	if (null !== this.time) {
		writer.WriteXmlAttributeStringEncode("time", this.time);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_PCDSDTCEntries() {
//Attributes
//	this.count = null;
//Members
	this.Items = [];
}
CT_PCDSDTCEntries.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("e" === elem) {
		newContext = new CT_Error();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.Items.push(newContext);
	} else if ("m" === elem) {
		newContext = new CT_Missing();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.Items.push(newContext);
	} else if ("n" === elem) {
		newContext = new CT_Number();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.Items.push(newContext);
	} else if ("s" === elem) {
		newContext = new CT_String();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.Items.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PCDSDTCEntries.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.Items.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.Items.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.Items.length; ++i) {
		var elem = this.Items[i];
		if (elem instanceof CT_Error) {
			elem.toXml(writer, "e");
		} else if (elem instanceof CT_Missing) {
			elem.toXml(writer, "m");
		} else if (elem instanceof CT_Number) {
			elem.toXml(writer, "n");
		} else if (elem instanceof CT_String) {
			elem.toXml(writer, "s");
		}
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_Sets() {
//Attributes
//	this.count = null;
//Members
	this.set = [];
}
CT_Sets.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("set" === elem) {
		newContext = new CT_Set();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.set.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Sets.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.set.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.set.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.set.length; ++i) {
		var elem = this.set[i];
		elem.toXml(writer, "set");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_QueryCache() {
//Attributes
//	this.count = null;
//Members
	this.query = [];
}
CT_QueryCache.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("query" === elem) {
		newContext = new CT_Query();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.query.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_QueryCache.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.query.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.query.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.query.length; ++i) {
		var elem = this.query[i];
		elem.toXml(writer, "query");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_ServerFormats() {
//Attributes
//	this.count = null;
//Members
	this.serverFormat = [];
}
CT_ServerFormats.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("serverFormat" === elem) {
		newContext = new CT_ServerFormat();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.serverFormat.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_ServerFormats.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.serverFormat.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.serverFormat.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.serverFormat.length; ++i) {
		var elem = this.serverFormat[i];
		elem.toXml(writer, "serverFormat");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_CalculatedItem() {
//Attributes
	this.field = null;
	this.formula = null;
//Members
	this.pivotArea = null;
	this.extLst = null;
}
CT_CalculatedItem.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["field"];
		if (undefined !== val) {
			this.field = val - 0;
		}
		val = vals["formula"];
		if (undefined !== val) {
			this.formula = AscCommon.unleakString(uq(val));
		}
	}
};
CT_CalculatedItem.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pivotArea" === elem) {
		newContext = new CT_PivotArea();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pivotArea = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_CalculatedItem.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.field) {
		writer.WriteXmlAttributeNumber("field", this.field);
	}
	if (null !== this.formula) {
		writer.WriteXmlAttributeStringEncode("formula", this.formula);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.pivotArea) {
		this.pivotArea.toXml(writer, "pivotArea");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_CalculatedMember() {
//Attributes
	this.name = null;
	this.mdx = null;
	this.memberName = null;
	this.hierarchy = null;
	this.parent = null;
	this.solveOrder = 0;
	this.set = false;
//Members
	this.extLst = null;
}
CT_CalculatedMember.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["mdx"];
		if (undefined !== val) {
			this.mdx = AscCommon.unleakString(uq(val));
		}
		val = vals["memberName"];
		if (undefined !== val) {
			this.memberName = AscCommon.unleakString(uq(val));
		}
		val = vals["hierarchy"];
		if (undefined !== val) {
			this.hierarchy = AscCommon.unleakString(uq(val));
		}
		val = vals["parent"];
		if (undefined !== val) {
			this.parent = AscCommon.unleakString(uq(val));
		}
		val = vals["solveOrder"];
		if (undefined !== val) {
			this.solveOrder = val - 0;
		}
		val = vals["set"];
		if (undefined !== val) {
			this.set = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_CalculatedMember.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_CalculatedMember.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (null !== this.mdx) {
		writer.WriteXmlAttributeStringEncode("mdx", this.mdx);
	}
	if (null !== this.memberName) {
		writer.WriteXmlAttributeStringEncode("memberName", this.memberName);
	}
	if (null !== this.hierarchy) {
		writer.WriteXmlAttributeStringEncode("hierarchy", this.hierarchy);
	}
	if (null !== this.parent) {
		writer.WriteXmlAttributeStringEncode("parent", this.parent);
	}
	if (0 !== this.solveOrder) {
		writer.WriteXmlAttributeNumber("solveOrder", this.solveOrder);
	}
	if (false !== this.set) {
		writer.WriteXmlAttributeBool("set", this.set);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_PivotDimension() {
//Attributes
	this.measure = false;
	this.name = null;
	this.uniqueName = null;
	this.caption = null;
}
CT_PivotDimension.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["measure"];
		if (undefined !== val) {
			this.measure = AscCommon.getBoolFromXml(val);
		}
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["uniqueName"];
		if (undefined !== val) {
			this.uniqueName = AscCommon.unleakString(uq(val));
		}
		val = vals["caption"];
		if (undefined !== val) {
			this.caption = AscCommon.unleakString(uq(val));
		}
	}
};
CT_PivotDimension.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (false !== this.measure) {
		writer.WriteXmlAttributeBool("measure", this.measure);
	}
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (null !== this.uniqueName) {
		writer.WriteXmlAttributeStringEncode("uniqueName", this.uniqueName);
	}
	if (null !== this.caption) {
		writer.WriteXmlAttributeStringEncode("caption", this.caption);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_MeasureGroup() {
//Attributes
	this.name = null;
	this.caption = null;
}
CT_MeasureGroup.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["caption"];
		if (undefined !== val) {
			this.caption = AscCommon.unleakString(uq(val));
		}
	}
};
CT_MeasureGroup.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (null !== this.caption) {
		writer.WriteXmlAttributeStringEncode("caption", this.caption);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_MeasureDimensionMap() {
//Attributes
	this.measureGroup = null;
	this.dimension = null;
}
CT_MeasureDimensionMap.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["measureGroup"];
		if (undefined !== val) {
			this.measureGroup = val - 0;
		}
		val = vals["dimension"];
		if (undefined !== val) {
			this.dimension = val - 0;
		}
	}
};
CT_MeasureDimensionMap.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.measureGroup) {
		writer.WriteXmlAttributeNumber("measureGroup", this.measureGroup);
	}
	if (null !== this.dimension) {
		writer.WriteXmlAttributeNumber("dimension", this.dimension);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_Extension() {
//Attributes
	this.uri = null;
//Members
	this.elem = null;
}
CT_Extension.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["uri"];
		if (undefined !== val) {
			this.uri = AscCommon.unleakString(uq(val));
		}
	}
};
CT_Extension.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("x14:pivotTableDefinition" === elem) {
		newContext = new CT_pivotTableDefinitionX14();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.elem = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Extension.prototype.toXml = function(writer, name) {
  if ("{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" === this.uri && this.elem) {
    writer.WriteXmlNodeStart(name);
    if (null !== this.uri) {
      writer.WriteXmlAttributeStringEncode("uri", this.uri);
      writer.WriteXmlString(" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"");
    }
    writer.WriteXmlNodeEnd(name, true);
    this.elem.toXml(writer, "x14:pivotTableDefinition");
    writer.WriteXmlNodeEnd(name);
  }
};
function CT_X() {
//Attributes
	this.v = 0;
}
CT_X.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["v"];
		if (undefined !== val) {
			this.v = val - 0;
		}
	}
};
CT_X.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (0 !== this.v) {
		writer.WriteXmlAttributeNumber("v", this.v);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
CT_X.prototype.getV = function () {
	return this.v;
};
function CT_Tuples() {
//Attributes
	this.c = null;
//Members
	this.tpl = [];
}
CT_Tuples.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["c"];
		if (undefined !== val) {
			this.c = val - 0;
		}
	}
};
CT_Tuples.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("tpl" === elem) {
		newContext = new CT_Tuple();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.tpl.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Tuples.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.c) {
		writer.WriteXmlAttributeNumber("c", this.c);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.tpl.length; ++i) {
		var elem = this.tpl[i];
		elem.toXml(writer, "tpl");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_PivotField(setDefaults) {
//Attributes
	this.name = null;
	this.axis = null;
	this.dataField = null;
	this.subtotalCaption = null;
	this.showDropDowns = null;
	this.hiddenLevel = null;
	this.uniqueMemberProperty = null;
	this.compact = null;
	this.allDrilled = null;
	this.numFmtId = null;
	this.outline = null;
	this.subtotalTop = null;
	this.dragToRow = null;
	this.dragToCol = null;
	this.multipleItemSelectionAllowed = null;
	this.dragToPage = null;
	this.dragToData = null;
	this.dragOff = null;
	this.showAll = null;
	this.insertBlankRow = null;
	this.serverField = null;
	this.insertPageBreak = null;
	this.autoShow = null;
	this.topAutoShow = null;
	this.hideNewItems = null;
	this.measureFilter = null;
	this.includeNewItemsInFilter = null;
	this.itemPageCount = null;
	this.sortType = null;
	this.dataSourceSort = null;
	this.nonAutoSortDefault = null;
	this.rankBy = null;
	this.defaultSubtotal = null;
	this.sumSubtotal = null;
	this.countASubtotal = null;
	this.avgSubtotal = null;
	this.maxSubtotal = null;
	this.minSubtotal = null;
	this.productSubtotal = null;
	this.countSubtotal = null;
	this.stdDevSubtotal = null;
	this.stdDevPSubtotal = null;
	this.varSubtotal = null;
	this.varPSubtotal = null;
	this.showPropCell = null;
	this.showPropTip = null;
	this.showPropAsCaption = null;
	this.defaultAttributeDrillState = null;
//Members
	this.items = null;
	this.autoSortScope = null;
	this.extLst = null;

	this.ascSubtotals = null;
	if (setDefaults) {
		this.setDefaults();
	}
}
CT_PivotField.prototype.setDefaults = function() {
	this.dataField = false;
	this.showDropDowns = true;
	this.hiddenLevel = false;
	this.compact = true;
	this.allDrilled = false;
	this.outline = true;
	this.subtotalTop = true;
	this.dragToRow = true;
	this.dragToCol = true;
	this.multipleItemSelectionAllowed = false;
	this.dragToPage = true;
	this.dragToData = true;
	this.dragOff = true;
	this.showAll = true;
	this.insertBlankRow = false;
	this.serverField = false;
	this.insertPageBreak = false;
	this.autoShow = false;
	this.topAutoShow = true;
	this.hideNewItems = false;
	this.measureFilter = false;
	this.includeNewItemsInFilter = false;
	this.itemPageCount = 10;
	this.sortType = c_oAscFieldSortType.Manual;
	this.nonAutoSortDefault = false;
	this.defaultSubtotal = true;
	this.sumSubtotal = false;
	this.countASubtotal = false;
	this.avgSubtotal = false;
	this.maxSubtotal = false;
	this.minSubtotal = false;
	this.productSubtotal = false;
	this.countSubtotal = false;
	this.stdDevSubtotal = false;
	this.stdDevPSubtotal = false;
	this.varSubtotal = false;
	this.varPSubtotal = false;
	this.showPropCell = false;
	this.showPropTip = false;
	this.showPropAsCaption = false;
	this.defaultAttributeDrillState = false;
};
CT_PivotField.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["axis"];
		if (undefined !== val) {
			val = FromXml_ST_Axis(val);
			if (-1 !== val) {
				this.axis = val;
			}
		}
		val = vals["dataField"];
		if (undefined !== val) {
			this.dataField = AscCommon.getBoolFromXml(val);
		}
		val = vals["subtotalCaption"];
		if (undefined !== val) {
			this.subtotalCaption = AscCommon.unleakString(uq(val));
		}
		val = vals["showDropDowns"];
		if (undefined !== val) {
			this.showDropDowns = AscCommon.getBoolFromXml(val);
		}
		val = vals["hiddenLevel"];
		if (undefined !== val) {
			this.hiddenLevel = AscCommon.getBoolFromXml(val);
		}
		val = vals["uniqueMemberProperty"];
		if (undefined !== val) {
			this.uniqueMemberProperty = AscCommon.unleakString(uq(val));
		}
		val = vals["compact"];
		if (undefined !== val) {
			this.compact = AscCommon.getBoolFromXml(val);
		}
		val = vals["allDrilled"];
		if (undefined !== val) {
			this.allDrilled = AscCommon.getBoolFromXml(val);
		}
		val = vals["numFmtId"];
		if (undefined !== val) {
			this.numFmtId = val - 0;
		}
		val = vals["outline"];
		if (undefined !== val) {
			this.outline = AscCommon.getBoolFromXml(val);
		}
		val = vals["subtotalTop"];
		if (undefined !== val) {
			this.subtotalTop = AscCommon.getBoolFromXml(val);
		}
		val = vals["dragToRow"];
		if (undefined !== val) {
			this.dragToRow = AscCommon.getBoolFromXml(val);
		}
		val = vals["dragToCol"];
		if (undefined !== val) {
			this.dragToCol = AscCommon.getBoolFromXml(val);
		}
		val = vals["multipleItemSelectionAllowed"];
		if (undefined !== val) {
			this.multipleItemSelectionAllowed = AscCommon.getBoolFromXml(val);
		}
		val = vals["dragToPage"];
		if (undefined !== val) {
			this.dragToPage = AscCommon.getBoolFromXml(val);
		}
		val = vals["dragToData"];
		if (undefined !== val) {
			this.dragToData = AscCommon.getBoolFromXml(val);
		}
		val = vals["dragOff"];
		if (undefined !== val) {
			this.dragOff = AscCommon.getBoolFromXml(val);
		}
		val = vals["showAll"];
		if (undefined !== val) {
			this.showAll = AscCommon.getBoolFromXml(val);
		}
		val = vals["insertBlankRow"];
		if (undefined !== val) {
			this.insertBlankRow = AscCommon.getBoolFromXml(val);
		}
		val = vals["serverField"];
		if (undefined !== val) {
			this.serverField = AscCommon.getBoolFromXml(val);
		}
		val = vals["insertPageBreak"];
		if (undefined !== val) {
			this.insertPageBreak = AscCommon.getBoolFromXml(val);
		}
		val = vals["autoShow"];
		if (undefined !== val) {
			this.autoShow = AscCommon.getBoolFromXml(val);
		}
		val = vals["topAutoShow"];
		if (undefined !== val) {
			this.topAutoShow = AscCommon.getBoolFromXml(val);
		}
		val = vals["hideNewItems"];
		if (undefined !== val) {
			this.hideNewItems = AscCommon.getBoolFromXml(val);
		}
		val = vals["measureFilter"];
		if (undefined !== val) {
			this.measureFilter = AscCommon.getBoolFromXml(val);
		}
		val = vals["includeNewItemsInFilter"];
		if (undefined !== val) {
			this.includeNewItemsInFilter = AscCommon.getBoolFromXml(val);
		}
		val = vals["itemPageCount"];
		if (undefined !== val) {
			this.itemPageCount = val - 0;
		}
		val = vals["sortType"];
		if (undefined !== val) {
			val = FromXml_ST_FieldSortType(val);
			if (-1 !== val) {
				this.sortType = val;
			}
		}
		val = vals["dataSourceSort"];
		if (undefined !== val) {
			this.dataSourceSort = AscCommon.getBoolFromXml(val);
		}
		val = vals["nonAutoSortDefault"];
		if (undefined !== val) {
			this.nonAutoSortDefault = AscCommon.getBoolFromXml(val);
		}
		val = vals["rankBy"];
		if (undefined !== val) {
			this.rankBy = val - 0;
		}
		val = vals["defaultSubtotal"];
		if (undefined !== val) {
			this.defaultSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["sumSubtotal"];
		if (undefined !== val) {
			this.sumSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["countASubtotal"];
		if (undefined !== val) {
			this.countASubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["avgSubtotal"];
		if (undefined !== val) {
			this.avgSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["maxSubtotal"];
		if (undefined !== val) {
			this.maxSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["minSubtotal"];
		if (undefined !== val) {
			this.minSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["productSubtotal"];
		if (undefined !== val) {
			this.productSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["countSubtotal"];
		if (undefined !== val) {
			this.countSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["stdDevSubtotal"];
		if (undefined !== val) {
			this.stdDevSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["stdDevPSubtotal"];
		if (undefined !== val) {
			this.stdDevPSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["varSubtotal"];
		if (undefined !== val) {
			this.varSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["varPSubtotal"];
		if (undefined !== val) {
			this.varPSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["showPropCell"];
		if (undefined !== val) {
			this.showPropCell = AscCommon.getBoolFromXml(val);
		}
		val = vals["showPropTip"];
		if (undefined !== val) {
			this.showPropTip = AscCommon.getBoolFromXml(val);
		}
		val = vals["showPropAsCaption"];
		if (undefined !== val) {
			this.showPropAsCaption = AscCommon.getBoolFromXml(val);
		}
		val = vals["defaultAttributeDrillState"];
		if (undefined !== val) {
			this.defaultAttributeDrillState = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_PivotField.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("items" === elem) {
		newContext = new CT_Items();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.items = newContext;
	} else if ("autoSortScope" === elem) {
		newContext = new CT_AutoSortScope();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.autoSortScope = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PivotField.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (null !== this.axis) {
		writer.WriteXmlAttributeStringEncode("axis", ToXml_ST_Axis(this.axis));
	}
	if (false !== this.dataField) {
		writer.WriteXmlAttributeBool("dataField", this.dataField);
	}
	if (null !== this.subtotalCaption) {
		writer.WriteXmlAttributeStringEncode("subtotalCaption", this.subtotalCaption);
	}
	if (true !== this.showDropDowns) {
		writer.WriteXmlAttributeBool("showDropDowns", this.showDropDowns);
	}
	if (false !== this.hiddenLevel) {
		writer.WriteXmlAttributeBool("hiddenLevel", this.hiddenLevel);
	}
	if (null !== this.uniqueMemberProperty) {
		writer.WriteXmlAttributeStringEncode("uniqueMemberProperty", this.uniqueMemberProperty);
	}
	if (true !== this.compact) {
		writer.WriteXmlAttributeBool("compact", this.compact);
	}
	if (false !== this.allDrilled) {
		writer.WriteXmlAttributeBool("allDrilled", this.allDrilled);
	}
	if (null !== this.numFmtId) {
		writer.WriteXmlAttributeNumber("numFmtId", this.numFmtId);
	}
	if (true !== this.outline) {
		writer.WriteXmlAttributeBool("outline", this.outline);
	}
	if (true !== this.subtotalTop) {
		writer.WriteXmlAttributeBool("subtotalTop", this.subtotalTop);
	}
	if (true !== this.dragToRow) {
		writer.WriteXmlAttributeBool("dragToRow", this.dragToRow);
	}
	if (true !== this.dragToCol) {
		writer.WriteXmlAttributeBool("dragToCol", this.dragToCol);
	}
	if (false !== this.multipleItemSelectionAllowed) {
		writer.WriteXmlAttributeBool("multipleItemSelectionAllowed", this.multipleItemSelectionAllowed);
	}
	if (true !== this.dragToPage) {
		writer.WriteXmlAttributeBool("dragToPage", this.dragToPage);
	}
	if (true !== this.dragToData) {
		writer.WriteXmlAttributeBool("dragToData", this.dragToData);
	}
	if (true !== this.dragOff) {
		writer.WriteXmlAttributeBool("dragOff", this.dragOff);
	}
	if (true !== this.showAll) {
		writer.WriteXmlAttributeBool("showAll", this.showAll);
	}
	if (false !== this.insertBlankRow) {
		writer.WriteXmlAttributeBool("insertBlankRow", this.insertBlankRow);
	}
	if (false !== this.serverField) {
		writer.WriteXmlAttributeBool("serverField", this.serverField);
	}
	if (false !== this.insertPageBreak) {
		writer.WriteXmlAttributeBool("insertPageBreak", this.insertPageBreak);
	}
	if (false !== this.autoShow) {
		writer.WriteXmlAttributeBool("autoShow", this.autoShow);
	}
	if (true !== this.topAutoShow) {
		writer.WriteXmlAttributeBool("topAutoShow", this.topAutoShow);
	}
	if (false !== this.hideNewItems) {
		writer.WriteXmlAttributeBool("hideNewItems", this.hideNewItems);
	}
	if (false !== this.measureFilter) {
		writer.WriteXmlAttributeBool("measureFilter", this.measureFilter);
	}
	if (false !== this.includeNewItemsInFilter) {
		writer.WriteXmlAttributeBool("includeNewItemsInFilter", this.includeNewItemsInFilter);
	}
	if (10 !== this.itemPageCount) {
		writer.WriteXmlAttributeNumber("itemPageCount", this.itemPageCount);
	}
	if (c_oAscFieldSortType.Manual !== this.sortType) {
		writer.WriteXmlAttributeStringEncode("sortType", ToXml_ST_FieldSortType(this.sortType));
	}
	if (null !== this.dataSourceSort) {
		writer.WriteXmlAttributeBool("dataSourceSort", this.dataSourceSort);
	}
	if (false !== this.nonAutoSortDefault) {
		writer.WriteXmlAttributeBool("nonAutoSortDefault", this.nonAutoSortDefault);
	}
	if (null !== this.rankBy) {
		writer.WriteXmlAttributeNumber("rankBy", this.rankBy);
	}
	if (true !== this.defaultSubtotal) {
		writer.WriteXmlAttributeBool("defaultSubtotal", this.defaultSubtotal);
	}
	if (false !== this.sumSubtotal) {
		writer.WriteXmlAttributeBool("sumSubtotal", this.sumSubtotal);
	}
	if (false !== this.countASubtotal) {
		writer.WriteXmlAttributeBool("countASubtotal", this.countASubtotal);
	}
	if (false !== this.avgSubtotal) {
		writer.WriteXmlAttributeBool("avgSubtotal", this.avgSubtotal);
	}
	if (false !== this.maxSubtotal) {
		writer.WriteXmlAttributeBool("maxSubtotal", this.maxSubtotal);
	}
	if (false !== this.minSubtotal) {
		writer.WriteXmlAttributeBool("minSubtotal", this.minSubtotal);
	}
	if (false !== this.productSubtotal) {
		writer.WriteXmlAttributeBool("productSubtotal", this.productSubtotal);
	}
	if (false !== this.countSubtotal) {
		writer.WriteXmlAttributeBool("countSubtotal", this.countSubtotal);
	}
	if (false !== this.stdDevSubtotal) {
		writer.WriteXmlAttributeBool("stdDevSubtotal", this.stdDevSubtotal);
	}
	if (false !== this.stdDevPSubtotal) {
		writer.WriteXmlAttributeBool("stdDevPSubtotal", this.stdDevPSubtotal);
	}
	if (false !== this.varSubtotal) {
		writer.WriteXmlAttributeBool("varSubtotal", this.varSubtotal);
	}
	if (false !== this.varPSubtotal) {
		writer.WriteXmlAttributeBool("varPSubtotal", this.varPSubtotal);
	}
	if (false !== this.showPropCell) {
		writer.WriteXmlAttributeBool("showPropCell", this.showPropCell);
	}
	if (false !== this.showPropTip) {
		writer.WriteXmlAttributeBool("showPropTip", this.showPropTip);
	}
	if (false !== this.showPropAsCaption) {
		writer.WriteXmlAttributeBool("showPropAsCaption", this.showPropAsCaption);
	}
	if (false !== this.defaultAttributeDrillState) {
		writer.WriteXmlAttributeBool("defaultAttributeDrillState", this.defaultAttributeDrillState);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.items) {
		this.items.toXml(writer, "items");
	}
	if (null !== this.autoSortScope) {
		this.autoSortScope.toXml(writer, "autoSortScope");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
CT_PivotField.prototype.asc_getName = function () {
	return this.name;
};
CT_PivotField.prototype.asc_getOutline = function () {
	return this.outline;
};
CT_PivotField.prototype.asc_getCompact = function () {
	return this.compact;
};
CT_PivotField.prototype.asc_getFillDownLabelsDefault = function () {
	return false;
};
CT_PivotField.prototype.asc_getInsertBlankRow = function () {
	return this.insertBlankRow;
};
CT_PivotField.prototype.asc_getDefaultSubtotal = function () {
	return this.defaultSubtotal;
};
CT_PivotField.prototype.asc_getSubtotalTop = function () {
	return this.subtotalTop;
};
CT_PivotField.prototype.asc_getShowAll = function () {
	return this.showAll;
};
CT_PivotField.prototype.asc_getSubtotals = function(withDefault) {
	var res = [];
	if (this.defaultSubtotal) {
		if (this.sumSubtotal) {
			res.push(Asc.c_oAscItemType.Sum);
		}
		if (this.countASubtotal) {
			res.push(Asc.c_oAscItemType.CountA);
		}
		if (this.avgSubtotal) {
			res.push(Asc.c_oAscItemType.Avg);
		}
		if (this.maxSubtotal) {
			res.push(Asc.c_oAscItemType.Max);
		}
		if (this.minSubtotal) {
			res.push(Asc.c_oAscItemType.Min);
		}
		if (this.productSubtotal) {
			res.push(Asc.c_oAscItemType.Product);
		}
		if (this.countSubtotal) {
			res.push(Asc.c_oAscItemType.Count);
		}
		if (this.stdDevSubtotal) {
			res.push(Asc.c_oAscItemType.StdDev);
		}
		if (this.stdDevPSubtotal) {
			res.push(Asc.c_oAscItemType.StdDevP);
		}
		if (this.varSubtotal) {
			res.push(Asc.c_oAscItemType.Var);
		}
		if (this.varPSubtotal) {
			res.push(Asc.c_oAscItemType.VarP);
		}
		if (withDefault && 0 == res.length) {
			res.push(Asc.c_oAscItemType.Default);
		}
	}
	return res;
};
CT_PivotField.prototype.getItem = function (index) {
	return this.items && this.items.item[index];
};
CT_PivotField.prototype.checkSubtotal = function() {
	if (!this.items) {
		return;
	}
	var newItem, i;
	var items = this.items.item;
	for (i = items.length - 1; i >= 0; --i) {
		var item = items[i];
		if (!(Asc.c_oAscItemType.Data === item.t || Asc.c_oAscItemType.Blank === item.t)) {
			items.splice(i, 1);
		}
	}
	var subtotals = this.asc_getSubtotals(true);
	for (i = 0; i < subtotals.length; ++i) {
		newItem = new CT_Item();
		newItem.t = subtotals[i];
		this.items.item.push(newItem);
	}
};
CT_PivotField.prototype.checkSubtotalTop = function() {
	return this.asc_getSubtotalTop() && 1 === this.asc_getSubtotals(true).length;
};
CT_PivotField.prototype.getSubtotalType = function() {
	var subtotals = this.asc_getSubtotals(true);
	return 1 === subtotals.length ? subtotals[0] : Asc.c_oAscItemType.Default;
};
CT_PivotField.prototype.asc_set = function (api, pivot, index, newVal) {
	var t = this;
	api._changePivotWithLock(pivot, function (ws) {
		if (null !== newVal.name) {
			t.asc_setName(newVal.name, pivot, index, true);
		}
		if (null !== newVal.outline) {
			t.asc_setOutline(newVal.outline, pivot, index, true);
		}
		if (null !== newVal.compact) {
			t.asc_setCompact(newVal.compact, pivot, index, true);
		}
		//todo repeat
		if (null !== newVal.insertBlankRow) {
			t.asc_setInsertBlankRow(newVal.insertBlankRow, pivot, index, true);
		}
		if (null !== newVal.defaultSubtotal) {
			t.asc_setDefaultSubtotal(newVal.defaultSubtotal, pivot, index, true);
		}
		if (null !== newVal.subtotalTop) {
			t.asc_setSubtotalTop(newVal.subtotalTop, pivot, index, true);
		}
		if (null !== newVal.showAll) {
			t.asc_setShowAll(newVal.showAll, pivot, index, true);
		}
		if (null !== newVal.ascSubtotals) {
			t.setSubtotals(newVal.ascSubtotals, pivot, index, true);
		}
		t.checkSubtotal();
	});
};
CT_PivotField.prototype.asc_setName = function (newVal, pivot, index, addToHistory) {
	setFieldProperty(pivot, index, this.name, newVal, addToHistory, AscCH.historyitem_PivotTable_PivotFieldSetName, true);
	this.name = newVal;
};
CT_PivotField.prototype.asc_setOutline = function (newVal, pivot, index, addToHistory) {
	setFieldProperty(pivot, index, this.outline, newVal, addToHistory, AscCH.historyitem_PivotTable_PivotFieldSetOutline, true);
	this.outline = newVal;
};
CT_PivotField.prototype.asc_setCompact = function (newVal, pivot, index, addToHistory) {
	setFieldProperty(pivot, index, this.compact, newVal, addToHistory, AscCH.historyitem_PivotTable_PivotFieldSetCompact, true);
	this.compact = newVal;
};
CT_PivotField.prototype.asc_setFillDownLabelsDefault = function (newVal, pivot, index, addToHistory) {
};
CT_PivotField.prototype.asc_setInsertBlankRow = function (newVal, pivot, index, addToHistory) {
	setFieldProperty(pivot, index, this.insertBlankRow, newVal, addToHistory, AscCH.historyitem_PivotTable_PivotFieldSetInsertBlankRow, true);
	this.insertBlankRow = newVal;
};
CT_PivotField.prototype.asc_setDefaultSubtotal = function (newVal, pivot, index, addToHistory) {
	setFieldProperty(pivot, index, this.defaultSubtotal, newVal, addToHistory, AscCH.historyitem_PivotTable_PivotFieldSetDefaultSubtotal, true);
	this.defaultSubtotal = newVal;
};
CT_PivotField.prototype.asc_setSubtotalTop = function (newVal, pivot, index, addToHistory) {
	setFieldProperty(pivot, index, this.subtotalTop, newVal, addToHistory, AscCH.historyitem_PivotTable_PivotFieldSetSubtotalTop, true);
	this.subtotalTop = newVal;
};
CT_PivotField.prototype.asc_setShowAll = function (newVal, pivot, index, addToHistory) {
	setFieldProperty(pivot, index, this.showAll, newVal, addToHistory, AscCH.historyitem_PivotTable_PivotFieldSetShowAll, true);
	this.showAll = newVal;
};
CT_PivotField.prototype.asc_setSubtotals = function (newVals) {
	this.ascSubtotals = newVals;
};
CT_PivotField.prototype.setSubtotals = function (newVals, pivot, index, addToHistory) {
	var oldVals = this.asc_getSubtotals();
	setFieldProperty(pivot, index, oldVals, newVals, addToHistory, AscCH.historyitem_PivotTable_PivotFieldSetSubtotals, true);
	this.sumSubtotal = this.countASubtotal = this.avgSubtotal = this.maxSubtotal = this.minSubtotal = false;
	this.productSubtotal = this.countSubtotal =	this.stdDevSubtotal = this.stdDevPSubtotal = this.varSubtotal = false;
	this.varPSubtotal = false;
	for (var i = 0; i < newVals.length; ++i) {
		var type = newVals[i];
		switch (type) {
			case Asc.c_oAscItemType.Sum:
				this.sumSubtotal = true;
				break;
			case Asc.c_oAscItemType.CountA:
				this.countASubtotal = true;
				break;
			case Asc.c_oAscItemType.Avg:
				this.avgSubtotal = true;
				break;
			case Asc.c_oAscItemType.Max:
				this.maxSubtotal = true;
				break;
			case Asc.c_oAscItemType.Min:
				this.minSubtotal = true;
				break;
			case Asc.c_oAscItemType.Product:
				this.productSubtotal = true;
				break;
			case Asc.c_oAscItemType.Count:
				this.countSubtotal = true;
				break;
			case Asc.c_oAscItemType.StdDev:
				this.stdDevSubtotal = true;
				break;
			case Asc.c_oAscItemType.StdDevP:
				this.stdDevPSubtotal = true;
				break;
			case Asc.c_oAscItemType.Var:
				this.varSubtotal = true;
				break;
			case Asc.c_oAscItemType.VarP:
				this.varPSubtotal = true;
				break;
		}
	}
};

function CT_Field() {
//Attributes
	this.x = null;
}
CT_Field.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["x"];
		if (undefined !== val) {
			this.x = val - 0;
		}
	}
};
CT_Field.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.x) {
		writer.WriteXmlAttributeNumber("x", this.x);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
CT_Field.prototype.asc_getIndex = function () {
	return this.x || 0;
};
function CT_I() {
//Attributes
	this.t = Asc.c_oAscItemType.Data;
	this.r = 0;
	this.i = 0;
//Members
	this.x = [];
}
CT_I.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["t"];
		if (undefined !== val) {
			val = FromXml_ST_ItemType(val);
			if (-1 !== val) {
				this.t = val;
			}
		}
		val = vals["r"];
		if (undefined !== val) {
			this.r = val - 0;
		}
		val = vals["i"];
		if (undefined !== val) {
			this.i = val - 0;
		}
	}
};
CT_I.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("x" === elem) {
		newContext = new CT_X();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.x.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_I.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (Asc.c_oAscItemType.Data !== this.t) {
		writer.WriteXmlAttributeStringEncode("t", ToXml_ST_ItemType(this.t));
	}
	if (0 !== this.r) {
		writer.WriteXmlAttributeNumber("r", this.r);
	}
	if (0 !== this.i) {
		writer.WriteXmlAttributeNumber("i", this.i);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.x.length; ++i) {
		var elem = this.x[i];
		elem.toXml(writer, "x");
	}
	writer.WriteXmlNodeEnd(name);
};
CT_I.prototype.getR = function () {
	return this.r;
};
function CT_PageField() {
//Attributes
	this.fld = null;
	this.item = null;
	this.hier = null;
	this.name = null;
	this.cap = null;
//Members
	this.extLst = null;
}
CT_PageField.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["fld"];
		if (undefined !== val) {
			this.fld = val - 0;
		}
		val = vals["item"];
		if (undefined !== val) {
			this.item = val - 0;
		}
		val = vals["hier"];
		if (undefined !== val) {
			this.hier = val - 0;
		}
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["cap"];
		if (undefined !== val) {
			this.cap = AscCommon.unleakString(uq(val));
		}
	}
};
CT_PageField.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PageField.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.fld) {
		writer.WriteXmlAttributeNumber("fld", this.fld);
	}
	if (null !== this.item) {
		writer.WriteXmlAttributeNumber("item", this.item);
	}
	if (null !== this.hier) {
		writer.WriteXmlAttributeNumber("hier", this.hier);
	}
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (null !== this.cap) {
		writer.WriteXmlAttributeStringEncode("cap", this.cap);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
CT_PageField.prototype.asc_getName = function () {
	return this.name;
};
CT_PageField.prototype.asc_getIndex = function () {
	return this.fld || 0;
};
function CT_DataField(setDefaults) {
//Attributes
	this.name = null;
	this.fld = null;
	this.subtotal = null;
	this.showDataAs = null;
	this.baseField = null;
	this.baseItem = null;
	this.numFmtId = null;
//Members
	this.extLst = null;

	if (setDefaults) {
		this.setDefaults();
	}
}
CT_DataField.prototype.setDefaults = function() {
	this.subtotal = c_oAscDataConsolidateFunction.Sum;
	this.showDataAs = c_oAscShowDataAs.Normal;
	this.baseField = -1;
	this.baseItem = 1048832;
};
CT_DataField.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["fld"];
		if (undefined !== val) {
			this.fld = val - 0;
		}
		val = vals["subtotal"];
		if (undefined !== val) {
			val = FromXml_ST_DataConsolidateFunction(val);
			if (-1 !== val) {
				this.subtotal = val;
			}
		}
		val = vals["showDataAs"];
		if (undefined !== val) {
			val = FromXml_ST_ShowDataAs(val);
			if (-1 !== val) {
				this.showDataAs = val;
			}
		}
		val = vals["baseField"];
		if (undefined !== val) {
			this.baseField = val - 0;
		}
		val = vals["baseItem"];
		if (undefined !== val) {
			this.baseItem = val - 0;
		}
		val = vals["numFmtId"];
		if (undefined !== val) {
			this.numFmtId = val - 0;
		}
	}
};
CT_DataField.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_DataField.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (null !== this.fld) {
		writer.WriteXmlAttributeNumber("fld", this.fld);
	}
	if (c_oAscDataConsolidateFunction.Sum !== this.subtotal) {
		writer.WriteXmlAttributeStringEncode("subtotal", ToXml_ST_DataConsolidateFunction(this.subtotal));
	}
	if (c_oAscShowDataAs.Normal !== this.showDataAs) {
		writer.WriteXmlAttributeStringEncode("showDataAs", ToXml_ST_ShowDataAs(this.showDataAs));
	}
	if (null !== this.baseField) {
		writer.WriteXmlAttributeNumber("baseField", this.baseField);
	}
	if (null !== this.baseItem) {
		writer.WriteXmlAttributeNumber("baseItem", this.baseItem);
	}
	if (null !== this.numFmtId) {
		writer.WriteXmlAttributeNumber("numFmtId", this.numFmtId);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
CT_DataField.prototype.asc_getName = function () {
	return this.name;
};
CT_DataField.prototype.asc_getIndex = function () {
	return this.fld || 0;
};
CT_DataField.prototype.asc_getSubtotal = function () {
	return this.subtotal;
};
CT_DataField.prototype.asc_getShowDataAs = function () {
	return this.showDataAs;
};
CT_DataField.prototype.asc_set = function (api, pivot, index, newVal) {
	var t = this;
	api._changePivotWithLock(pivot, function (ws) {
		if (null !== newVal.name) {
			t.asc_setName(newVal.name, pivot, index, true);
		}
		if (null !== newVal.subtotal) {
			t.asc_setSubtotal(newVal.subtotal, pivot, index, true);
		}
	});
};
CT_DataField.prototype.asc_setName = function(newVal, pivot, index, addToHistory) {
	setFieldProperty(pivot, index, this.name, newVal, addToHistory, AscCH.historyitem_PivotTable_DataFieldSetName, true);
	this.name = newVal;
};
CT_DataField.prototype.asc_setSubtotal = function(newVal, pivot, index, addToHistory) {
	setFieldProperty(pivot, index, this.subtotal, newVal, addToHistory, AscCH.historyitem_PivotTable_DataFieldSetSubtotal, true);
	this.subtotal = newVal;
};
function CT_Format() {
//Attributes
	this.action = c_oAscFormatAction.Formatting;
	this.dxfId = null;
//Members
	this.pivotArea = null;
	this.extLst = null;
}
CT_Format.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["action"];
		if (undefined !== val) {
			val = FromXml_ST_FormatAction(val);
			if (-1 !== val) {
				this.action = val;
			}
		}
		val = vals["dxfId"];
		if (undefined !== val) {
			this.dxfId = val - 0;
		}
	}
};
CT_Format.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pivotArea" === elem) {
		newContext = new CT_PivotArea();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pivotArea = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Format.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (c_oAscFormatAction.Formatting !== this.action) {
		writer.WriteXmlAttributeStringEncode("action", ToXml_ST_FormatAction(this.action));
	}
	//todo
	// if (null !== this.dxfId) {
	// 	writer.WriteXmlAttributeNumber("dxfId", this.dxfId);
	// }
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.pivotArea) {
		this.pivotArea.toXml(writer, "pivotArea");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_ConditionalFormat() {
//Attributes
	this.scope = c_oAscScope.Selection;
	this.type = c_oAscType.None;
	this.priority = null;
//Members
	this.pivotAreas = null;
	this.extLst = null;
}
CT_ConditionalFormat.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["scope"];
		if (undefined !== val) {
			val = FromXml_ST_Scope(val);
			if (-1 !== val) {
				this.scope = val;
			}
		}
		val = vals["type"];
		if (undefined !== val) {
			val = FromXml_ST_Type(val);
			if (-1 !== val) {
				this.type = val;
			}
		}
		val = vals["priority"];
		if (undefined !== val) {
			this.priority = val - 0;
		}
	}
};
CT_ConditionalFormat.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pivotAreas" === elem) {
		newContext = new CT_PivotAreas();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pivotAreas = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_ConditionalFormat.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (c_oAscScope.Selection !== this.scope) {
		writer.WriteXmlAttributeStringEncode("scope", ToXml_ST_Scope(this.scope));
	}
	if (c_oAscType.None !== this.type) {
		writer.WriteXmlAttributeStringEncode("type", ToXml_ST_Type(this.type));
	}
	if (null !== this.priority) {
		writer.WriteXmlAttributeNumber("priority", this.priority);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.pivotAreas) {
		this.pivotAreas.toXml(writer, "pivotAreas");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_ChartFormat() {
//Attributes
	this.chart = null;
	this.format = null;
	this.series = false;
//Members
	this.pivotArea = null;
}
CT_ChartFormat.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["chart"];
		if (undefined !== val) {
			this.chart = val - 0;
		}
		val = vals["format"];
		if (undefined !== val) {
			this.format = val - 0;
		}
		val = vals["series"];
		if (undefined !== val) {
			this.series = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_ChartFormat.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pivotArea" === elem) {
		newContext = new CT_PivotArea();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pivotArea = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_ChartFormat.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.chart) {
		writer.WriteXmlAttributeNumber("chart", this.chart);
	}
	if (null !== this.format) {
		writer.WriteXmlAttributeNumber("format", this.format);
	}
	if (false !== this.series) {
		writer.WriteXmlAttributeBool("series", this.series);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.pivotArea) {
		this.pivotArea.toXml(writer, "pivotArea");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_PivotHierarchy() {
//Attributes
	this.outline = false;
	this.multipleItemSelectionAllowed = false;
	this.subtotalTop = false;
	this.showInFieldList = true;
	this.dragToRow = true;
	this.dragToCol = true;
	this.dragToPage = true;
	this.dragToData = false;
	this.dragOff = true;
	this.includeNewItemsInFilter = false;
	this.caption = null;
//Members
	this.mps = null;
	this.members = [];
	this.extLst = null;
}
CT_PivotHierarchy.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["outline"];
		if (undefined !== val) {
			this.outline = AscCommon.getBoolFromXml(val);
		}
		val = vals["multipleItemSelectionAllowed"];
		if (undefined !== val) {
			this.multipleItemSelectionAllowed = AscCommon.getBoolFromXml(val);
		}
		val = vals["subtotalTop"];
		if (undefined !== val) {
			this.subtotalTop = AscCommon.getBoolFromXml(val);
		}
		val = vals["showInFieldList"];
		if (undefined !== val) {
			this.showInFieldList = AscCommon.getBoolFromXml(val);
		}
		val = vals["dragToRow"];
		if (undefined !== val) {
			this.dragToRow = AscCommon.getBoolFromXml(val);
		}
		val = vals["dragToCol"];
		if (undefined !== val) {
			this.dragToCol = AscCommon.getBoolFromXml(val);
		}
		val = vals["dragToPage"];
		if (undefined !== val) {
			this.dragToPage = AscCommon.getBoolFromXml(val);
		}
		val = vals["dragToData"];
		if (undefined !== val) {
			this.dragToData = AscCommon.getBoolFromXml(val);
		}
		val = vals["dragOff"];
		if (undefined !== val) {
			this.dragOff = AscCommon.getBoolFromXml(val);
		}
		val = vals["includeNewItemsInFilter"];
		if (undefined !== val) {
			this.includeNewItemsInFilter = AscCommon.getBoolFromXml(val);
		}
		val = vals["caption"];
		if (undefined !== val) {
			this.caption = AscCommon.unleakString(uq(val));
		}
	}
};
CT_PivotHierarchy.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("mps" === elem) {
		newContext = new CT_MemberProperties();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.mps = newContext;
	} else if ("members" === elem) {
		newContext = new CT_Members();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.members.push(newContext);
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PivotHierarchy.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (false !== this.outline) {
		writer.WriteXmlAttributeBool("outline", this.outline);
	}
	if (false !== this.multipleItemSelectionAllowed) {
		writer.WriteXmlAttributeBool("multipleItemSelectionAllowed", this.multipleItemSelectionAllowed);
	}
	if (false !== this.subtotalTop) {
		writer.WriteXmlAttributeBool("subtotalTop", this.subtotalTop);
	}
	if (true !== this.showInFieldList) {
		writer.WriteXmlAttributeBool("showInFieldList", this.showInFieldList);
	}
	if (true !== this.dragToRow) {
		writer.WriteXmlAttributeBool("dragToRow", this.dragToRow);
	}
	if (true !== this.dragToCol) {
		writer.WriteXmlAttributeBool("dragToCol", this.dragToCol);
	}
	if (true !== this.dragToPage) {
		writer.WriteXmlAttributeBool("dragToPage", this.dragToPage);
	}
	if (false !== this.dragToData) {
		writer.WriteXmlAttributeBool("dragToData", this.dragToData);
	}
	if (true !== this.dragOff) {
		writer.WriteXmlAttributeBool("dragOff", this.dragOff);
	}
	if (false !== this.includeNewItemsInFilter) {
		writer.WriteXmlAttributeBool("includeNewItemsInFilter", this.includeNewItemsInFilter);
	}
	if (null !== this.caption) {
		writer.WriteXmlAttributeStringEncode("caption", this.caption);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.mps) {
		this.mps.toXml(writer, "mps");
	}
	for (var i = 0; i < this.members.length; ++i) {
		var elem = this.members[i];
		elem.toXml(writer, "members");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_PivotFilter() {
//Attributes
	this.fld = null;
	this.mpFld = null;
	this.type = null;
	this.evalOrder = 0;
	this.id = null;
	this.iMeasureHier = null;
	this.iMeasureFld = null;
	this.name = null;
	this.description = null;
	this.stringValue1 = null;
	this.stringValue2 = null;
//Members
	this.autoFilter = null;
	this.extLst = null;
}
CT_PivotFilter.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["fld"];
		if (undefined !== val) {
			this.fld = val - 0;
		}
		val = vals["mpFld"];
		if (undefined !== val) {
			this.mpFld = val - 0;
		}
		val = vals["type"];
		if (undefined !== val) {
			val = FromXml_ST_PivotFilterType(val);
			if (-1 !== val) {
				this.type = val;
			}
		}
		val = vals["evalOrder"];
		if (undefined !== val) {
			this.evalOrder = val - 0;
		}
		val = vals["id"];
		if (undefined !== val) {
			this.id = val - 0;
		}
		val = vals["iMeasureHier"];
		if (undefined !== val) {
			this.iMeasureHier = val - 0;
		}
		val = vals["iMeasureFld"];
		if (undefined !== val) {
			this.iMeasureFld = val - 0;
		}
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["description"];
		if (undefined !== val) {
			this.description = AscCommon.unleakString(uq(val));
		}
		val = vals["stringValue1"];
		if (undefined !== val) {
			this.stringValue1 = AscCommon.unleakString(uq(val));
		}
		val = vals["stringValue2"];
		if (undefined !== val) {
			this.stringValue2 = AscCommon.unleakString(uq(val));
		}
	}
};
CT_PivotFilter.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("autoFilter" === elem) {
		newContext = new CT_AutoFilter();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.autoFilter = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PivotFilter.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.fld) {
		writer.WriteXmlAttributeNumber("fld", this.fld);
	}
	if (null !== this.mpFld) {
		writer.WriteXmlAttributeNumber("mpFld", this.mpFld);
	}
	if (null !== this.type) {
		writer.WriteXmlAttributeStringEncode("type", ToXml_ST_PivotFilterType(this.type));
	}
	if (0 !== this.evalOrder) {
		writer.WriteXmlAttributeNumber("evalOrder", this.evalOrder);
	}
	if (null !== this.id) {
		writer.WriteXmlAttributeNumber("id", this.id);
	}
	if (null !== this.iMeasureHier) {
		writer.WriteXmlAttributeNumber("iMeasureHier", this.iMeasureHier);
	}
	if (null !== this.iMeasureFld) {
		writer.WriteXmlAttributeNumber("iMeasureFld", this.iMeasureFld);
	}
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (null !== this.description) {
		writer.WriteXmlAttributeStringEncode("description", this.description);
	}
	if (null !== this.stringValue1) {
		writer.WriteXmlAttributeStringEncode("stringValue1", this.stringValue1);
	}
	if (null !== this.stringValue2) {
		writer.WriteXmlAttributeStringEncode("stringValue2", this.stringValue2);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.autoFilter) {
		this.autoFilter.toXml(writer, "autoFilter");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_HierarchyUsage() {
//Attributes
	this.hierarchyUsage = null;
}
CT_HierarchyUsage.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["hierarchyUsage"];
		if (undefined !== val) {
			this.hierarchyUsage = val - 0;
		}
	}
};
CT_HierarchyUsage.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.hierarchyUsage) {
		writer.WriteXmlAttributeNumber("hierarchyUsage", this.hierarchyUsage);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_Pages() {
//Attributes
//	this.count = null;
//Members
	this.page = [];
}
CT_Pages.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("page" === elem) {
		newContext = new CT_PCDSCPage();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.page.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Pages.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.page.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.page.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.page.length; ++i) {
		var elem = this.page[i];
		elem.toXml(writer, "page");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_RangeSets() {
//Attributes
//	this.count = null;
//Members
	this.rangeSet = [];
}
CT_RangeSets.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("rangeSet" === elem) {
		newContext = new CT_RangeSet();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.rangeSet.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_RangeSets.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.rangeSet.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.rangeSet.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.rangeSet.length; ++i) {
		var elem = this.rangeSet[i];
		elem.toXml(writer, "rangeSet");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_SharedItems() {
//Attributes
	this.containsSemiMixedTypes = true;
	this.containsNonDate = true;
	this.containsDate = false;
	this.containsString = true;
	this.containsBlank = false;
	this.containsMixedTypes = false;
	this.containsNumber = false;
	this.containsInteger = false;
	this.minValue = null;
	this.maxValue = null;
	this.minDate = null;
	this.maxDate = null;
//	this.count = null;
	this.longText = false;
//Members
	this.Items = new PivotRecords();
}
CT_SharedItems.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["containsSemiMixedTypes"];
		if (undefined !== val) {
			this.containsSemiMixedTypes = AscCommon.getBoolFromXml(val);
		}
		val = vals["containsNonDate"];
		if (undefined !== val) {
			this.containsNonDate = AscCommon.getBoolFromXml(val);
		}
		val = vals["containsDate"];
		if (undefined !== val) {
			this.containsDate = AscCommon.getBoolFromXml(val);
		}
		val = vals["containsString"];
		if (undefined !== val) {
			this.containsString = AscCommon.getBoolFromXml(val);
		}
		val = vals["containsBlank"];
		if (undefined !== val) {
			this.containsBlank = AscCommon.getBoolFromXml(val);
		}
		val = vals["containsMixedTypes"];
		if (undefined !== val) {
			this.containsMixedTypes = AscCommon.getBoolFromXml(val);
		}
		val = vals["containsNumber"];
		if (undefined !== val) {
			this.containsNumber = AscCommon.getBoolFromXml(val);
		}
		val = vals["containsInteger"];
		if (undefined !== val) {
			this.containsInteger = AscCommon.getBoolFromXml(val);
		}
		val = vals["minValue"];
		if (undefined !== val) {
			this.minValue = val - 0;
		}
		val = vals["maxValue"];
		if (undefined !== val) {
			this.maxValue = val - 0;
		}
		val = vals["minDate"];
		if (undefined !== val) {
			this.minDate = AscCommon.unleakString(uq(val));
		}
		val = vals["maxDate"];
		if (undefined !== val) {
			this.maxDate = AscCommon.unleakString(uq(val));
		}
		val = vals["longText"];
		if (undefined !== val) {
			this.longText = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_SharedItems.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	var newContextCandidate = this.Items.onStartNode(elem, attr, uq);
	if (newContextCandidate) {
		newContext = newContextCandidate;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_SharedItems.prototype.onEndNode = function(prevContext, elem) {
	this.Items.onEndNode(prevContext, elem);
};
CT_SharedItems.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (true !== this.containsSemiMixedTypes) {
		writer.WriteXmlAttributeBool("containsSemiMixedTypes", this.containsSemiMixedTypes);
	}
	if (true !== this.containsNonDate) {
		writer.WriteXmlAttributeBool("containsNonDate", this.containsNonDate);
	}
	if (false !== this.containsDate) {
		writer.WriteXmlAttributeBool("containsDate", this.containsDate);
	}
	if (true !== this.containsString) {
		writer.WriteXmlAttributeBool("containsString", this.containsString);
	}
	if (false !== this.containsBlank) {
		writer.WriteXmlAttributeBool("containsBlank", this.containsBlank);
	}
	if (false !== this.containsMixedTypes) {
		writer.WriteXmlAttributeBool("containsMixedTypes", this.containsMixedTypes);
	}
	if (false !== this.containsNumber) {
		writer.WriteXmlAttributeBool("containsNumber", this.containsNumber);
	}
	if (false !== this.containsInteger) {
		writer.WriteXmlAttributeBool("containsInteger", this.containsInteger);
	}
	if (null !== this.minValue) {
		writer.WriteXmlAttributeNumber("minValue", this.minValue);
	}
	if (null !== this.maxValue) {
		writer.WriteXmlAttributeNumber("maxValue", this.maxValue);
	}
	if (null !== this.minDate) {
		writer.WriteXmlAttributeStringEncode("minDate", this.minDate);
	}
	if (null !== this.maxDate) {
		writer.WriteXmlAttributeStringEncode("maxDate", this.maxDate);
	}
	var count = this.Items.getSize();
	if (count > 0) {
		writer.WriteXmlAttributeNumber("count", count);
	}
	if (false !== this.longText) {
		writer.WriteXmlAttributeBool("longText", this.longText);
	}
	writer.WriteXmlNodeEnd(name, true);
	this.Items.toXml(writer);
	writer.WriteXmlNodeEnd(name);
};
function CT_FieldGroup() {
//Attributes
	this.par = null;
	this.base = null;
//Members
	this.rangePr = null;
	this.discretePr = null;
	this.groupItems = null;
}
CT_FieldGroup.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["par"];
		if (undefined !== val) {
			this.par = val - 0;
		}
		val = vals["base"];
		if (undefined !== val) {
			this.base = val - 0;
		}
	}
};
CT_FieldGroup.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("rangePr" === elem) {
		newContext = new CT_RangePr();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.rangePr = newContext;
	} else if ("discretePr" === elem) {
		newContext = new CT_DiscretePr();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.discretePr = newContext;
	} else if ("groupItems" === elem) {
		newContext = new CT_GroupItems();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.groupItems = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_FieldGroup.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.par) {
		writer.WriteXmlAttributeNumber("par", this.par);
	}
	if (null !== this.base) {
		writer.WriteXmlAttributeNumber("base", this.base);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.rangePr) {
		this.rangePr.toXml(writer, "rangePr");
	}
	if (null !== this.discretePr) {
		this.discretePr.toXml(writer, "discretePr");
	}
	if (null !== this.groupItems) {
		this.groupItems.toXml(writer, "groupItems");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_FieldsUsage() {
//Attributes
//	this.count = null;
//Members
	this.fieldUsage = [];
}
CT_FieldsUsage.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("fieldUsage" === elem) {
		newContext = new CT_FieldUsage();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.fieldUsage.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_FieldsUsage.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.fieldUsage.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.fieldUsage.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.fieldUsage.length; ++i) {
		var elem = this.fieldUsage[i];
		elem.toXml(writer, "fieldUsage");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_GroupLevels() {
//Attributes
//	this.count = null;
//Members
	this.groupLevel = [];
}
CT_GroupLevels.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("groupLevel" === elem) {
		newContext = new CT_GroupLevel();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.groupLevel.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_GroupLevels.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.groupLevel.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.groupLevel.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.groupLevel.length; ++i) {
		var elem = this.groupLevel[i];
		elem.toXml(writer, "groupLevel");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_Set() {
//Attributes
//	this.count = null;
	this.maxRank = null;
	this.setDefinition = null;
	this.sortType = c_oAscSortType.None;
	this.queryFailed = false;
//Members
	this.tpls = [];
	this.sortByTuple = null;
}
CT_Set.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["maxRank"];
		if (undefined !== val) {
			this.maxRank = val - 0;
		}
		val = vals["setDefinition"];
		if (undefined !== val) {
			this.setDefinition = AscCommon.unleakString(uq(val));
		}
		val = vals["sortType"];
		if (undefined !== val) {
			val = FromXml_ST_SortType(val);
			if (-1 !== val) {
				this.sortType = val;
			}
		}
		val = vals["queryFailed"];
		if (undefined !== val) {
			this.queryFailed = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_Set.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("tpls" === elem) {
		newContext = new CT_Tuples();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.tpls.push(newContext);
	} else if ("sortByTuple" === elem) {
		newContext = new CT_Tuples();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.sortByTuple = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Set.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.tpls.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.tpls.length);
	}
	if (null !== this.maxRank) {
		writer.WriteXmlAttributeNumber("maxRank", this.maxRank);
	}
	if (null !== this.setDefinition) {
		writer.WriteXmlAttributeStringEncode("setDefinition", this.setDefinition);
	}
	if (c_oAscSortType.None !== this.sortType) {
		writer.WriteXmlAttributeStringEncode("sortType", ToXml_ST_SortType(this.sortType));
	}
	if (false !== this.queryFailed) {
		writer.WriteXmlAttributeBool("queryFailed", this.queryFailed);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.tpls.length; ++i) {
		var elem = this.tpls[i];
		elem.toXml(writer, "tpls");
	}
	if (null !== this.sortByTuple) {
		this.sortByTuple.toXml(writer, "sortByTuple");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_Query() {
//Attributes
	this.mdx = null;
//Members
	this.tpls = null;
}
CT_Query.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["mdx"];
		if (undefined !== val) {
			this.mdx = AscCommon.unleakString(uq(val));
		}
	}
};
CT_Query.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("tpls" === elem) {
		newContext = new CT_Tuples();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.tpls = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Query.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.mdx) {
		writer.WriteXmlAttributeStringEncode("mdx", this.mdx);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.tpls) {
		this.tpls.toXml(writer, "tpls");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_ServerFormat() {
//Attributes
	this.culture = null;
	this.format = null;
}
CT_ServerFormat.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["culture"];
		if (undefined !== val) {
			this.culture = AscCommon.unleakString(uq(val));
		}
		val = vals["format"];
		if (undefined !== val) {
			this.format = AscCommon.unleakString(uq(val));
		}
	}
};
CT_ServerFormat.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.culture) {
		writer.WriteXmlAttributeStringEncode("culture", this.culture);
	}
	if (null !== this.format) {
		writer.WriteXmlAttributeStringEncode("format", this.format);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_PivotArea() {
//Attributes
	this.field = null;
	this.type = c_oAscPivotAreaType.None;
	this.dataOnly = true;
	this.labelOnly = false;
	this.grandRow = false;
	this.grandCol = false;
	this.cacheIndex = false;
	this.outline = true;
	this.offset = null;
	this.collapsedLevelsAreSubtotals = false;
	this.axis = null;
	this.fieldPosition = null;
//Members
	this.references = null;
	this.extLst = null;
}
CT_PivotArea.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["field"];
		if (undefined !== val) {
			this.field = val - 0;
		}
		val = vals["type"];
		if (undefined !== val) {
			val = FromXml_ST_PivotAreaType(val);
			if (-1 !== val) {
				this.type = val;
			}
		}
		val = vals["dataOnly"];
		if (undefined !== val) {
			this.dataOnly = AscCommon.getBoolFromXml(val);
		}
		val = vals["labelOnly"];
		if (undefined !== val) {
			this.labelOnly = AscCommon.getBoolFromXml(val);
		}
		val = vals["grandRow"];
		if (undefined !== val) {
			this.grandRow = AscCommon.getBoolFromXml(val);
		}
		val = vals["grandCol"];
		if (undefined !== val) {
			this.grandCol = AscCommon.getBoolFromXml(val);
		}
		val = vals["cacheIndex"];
		if (undefined !== val) {
			this.cacheIndex = AscCommon.getBoolFromXml(val);
		}
		val = vals["outline"];
		if (undefined !== val) {
			this.outline = AscCommon.getBoolFromXml(val);
		}
		val = vals["offset"];
		if (undefined !== val) {
			this.offset = AscCommon.unleakString(uq(val));
		}
		val = vals["collapsedLevelsAreSubtotals"];
		if (undefined !== val) {
			this.collapsedLevelsAreSubtotals = AscCommon.getBoolFromXml(val);
		}
		val = vals["axis"];
		if (undefined !== val) {
			val = FromXml_ST_Axis(val);
			if (-1 !== val) {
				this.axis = val;
			}
		}
		val = vals["fieldPosition"];
		if (undefined !== val) {
			this.fieldPosition = val - 0;
		}
	}
};
CT_PivotArea.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("references" === elem) {
		newContext = new CT_PivotAreaReferences();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.references = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PivotArea.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.field) {
		writer.WriteXmlAttributeNumber("field", this.field);
	}
	if (c_oAscPivotAreaType.None !== this.type) {
		writer.WriteXmlAttributeStringEncode("type", ToXml_ST_PivotAreaType(this.type));
	}
	if (true !== this.dataOnly) {
		writer.WriteXmlAttributeBool("dataOnly", this.dataOnly);
	}
	if (false !== this.labelOnly) {
		writer.WriteXmlAttributeBool("labelOnly", this.labelOnly);
	}
	if (false !== this.grandRow) {
		writer.WriteXmlAttributeBool("grandRow", this.grandRow);
	}
	if (false !== this.grandCol) {
		writer.WriteXmlAttributeBool("grandCol", this.grandCol);
	}
	if (false !== this.cacheIndex) {
		writer.WriteXmlAttributeBool("cacheIndex", this.cacheIndex);
	}
	if (true !== this.outline) {
		writer.WriteXmlAttributeBool("outline", this.outline);
	}
	if (null !== this.offset) {
		writer.WriteXmlAttributeStringEncode("offset", this.offset);
	}
	if (false !== this.collapsedLevelsAreSubtotals) {
		writer.WriteXmlAttributeBool("collapsedLevelsAreSubtotals", this.collapsedLevelsAreSubtotals);
	}
	if (null !== this.axis) {
		writer.WriteXmlAttributeStringEncode("axis", ToXml_ST_Axis(this.axis));
	}
	if (null !== this.fieldPosition) {
		writer.WriteXmlAttributeNumber("fieldPosition", this.fieldPosition);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.references) {
		this.references.toXml(writer, "references");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_Tuple() {
//Attributes
	this.fld = null;
	this.hier = null;
	this.item = null;
}
CT_Tuple.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["fld"];
		if (undefined !== val) {
			this.fld = val - 0;
		}
		val = vals["hier"];
		if (undefined !== val) {
			this.hier = val - 0;
		}
		val = vals["item"];
		if (undefined !== val) {
			this.item = val - 0;
		}
	}
};
CT_Tuple.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.fld) {
		writer.WriteXmlAttributeNumber("fld", this.fld);
	}
	if (null !== this.hier) {
		writer.WriteXmlAttributeNumber("hier", this.hier);
	}
	if (null !== this.item) {
		writer.WriteXmlAttributeNumber("item", this.item);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_Items() {
//Attributes
//	this.count = null;
//Members
	this.item = [];
}
CT_Items.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("item" === elem) {
		newContext = new CT_Item();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.item.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Items.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.item.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.item.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.item.length; ++i) {
		var elem = this.item[i];
		elem.toXml(writer, "item");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_AutoSortScope() {
//Members
	this.pivotArea = null;
}
CT_AutoSortScope.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pivotArea" === elem) {
		newContext = new CT_PivotArea();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pivotArea = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_AutoSortScope.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.pivotArea) {
		this.pivotArea.toXml(writer, "pivotArea");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_PivotAreas() {
//Attributes
//	this.count = null;
//Members
	this.pivotArea = [];
}
CT_PivotAreas.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pivotArea" === elem) {
		newContext = new CT_PivotArea();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pivotArea.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PivotAreas.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.pivotArea.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.pivotArea.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.pivotArea.length; ++i) {
		var elem = this.pivotArea[i];
		elem.toXml(writer, "pivotArea");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_MemberProperties() {
//Attributes
//	this.count = null;
//Members
	this.mp = [];
}
CT_MemberProperties.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("mp" === elem) {
		newContext = new CT_MemberProperty();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.mp.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_MemberProperties.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.mp.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.mp.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.mp.length; ++i) {
		var elem = this.mp[i];
		elem.toXml(writer, "mp");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_Members() {
//Attributes
//	this.count = null;
	this.level = null;
//Members
	this.member = [];
}
CT_Members.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["level"];
		if (undefined !== val) {
			this.level = val - 0;
		}
	}
};
CT_Members.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("member" === elem) {
		newContext = new CT_Member();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.member.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Members.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.member.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.member.length);
	}
	if (null !== this.level) {
		writer.WriteXmlAttributeNumber("level", this.level);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.member.length; ++i) {
		var elem = this.member[i];
		elem.toXml(writer, "member");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_AutoFilter() {
//Attributes
	this.ref = null;
//Members
	this.filterColumn = [];
	this.sortState = null;
	this.extLst = null;
}
CT_AutoFilter.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["ref"];
		if (undefined !== val) {
			this.ref = AscCommon.unleakString(uq(val));
		}
	}
};
CT_AutoFilter.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("filterColumn" === elem) {
		newContext = new CT_FilterColumn();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.filterColumn.push(newContext);
	} else if ("sortState" === elem) {
		newContext = new CT_SortState();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.sortState = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_AutoFilter.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.ref) {
		writer.WriteXmlAttributeStringEncode("ref", this.ref);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.filterColumn.length; ++i) {
		var elem = this.filterColumn[i];
		elem.toXml(writer, "filterColumn");
	}
	if (null !== this.sortState) {
		this.sortState.toXml(writer, "sortState");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_PCDSCPage() {
//Attributes
//	this.count = null;
//Members
	this.pageItem = [];
}
CT_PCDSCPage.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("pageItem" === elem) {
		newContext = new CT_PageItem();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.pageItem.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PCDSCPage.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.pageItem.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.pageItem.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.pageItem.length; ++i) {
		var elem = this.pageItem[i];
		elem.toXml(writer, "pageItem");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_RangeSet() {
//Attributes
	this.i1 = null;
	this.i2 = null;
	this.i3 = null;
	this.i4 = null;
	this.ref = null;
	this.name = null;
	this.sheet = null;
	this.id = null;
}
CT_RangeSet.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["i1"];
		if (undefined !== val) {
			this.i1 = val - 0;
		}
		val = vals["i2"];
		if (undefined !== val) {
			this.i2 = val - 0;
		}
		val = vals["i3"];
		if (undefined !== val) {
			this.i3 = val - 0;
		}
		val = vals["i4"];
		if (undefined !== val) {
			this.i4 = val - 0;
		}
		val = vals["ref"];
		if (undefined !== val) {
			this.ref = AscCommon.unleakString(uq(val));
		}
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["sheet"];
		if (undefined !== val) {
			this.sheet = AscCommon.unleakString(uq(val));
		}
		val = vals["r:id"];
		if (undefined !== val) {
			this.id = AscCommon.unleakString(uq(val));
		}
	}
};
CT_RangeSet.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.i1) {
		writer.WriteXmlAttributeNumber("i1", this.i1);
	}
	if (null !== this.i2) {
		writer.WriteXmlAttributeNumber("i2", this.i2);
	}
	if (null !== this.i3) {
		writer.WriteXmlAttributeNumber("i3", this.i3);
	}
	if (null !== this.i4) {
		writer.WriteXmlAttributeNumber("i4", this.i4);
	}
	if (null !== this.ref) {
		writer.WriteXmlAttributeStringEncode("ref", this.ref);
	}
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (null !== this.sheet) {
		writer.WriteXmlAttributeStringEncode("sheet", this.sheet);
	}
	//todo
	// if (null !== this.id) {
		// writer.WriteXmlAttributeStringEncode("r:id", this.id);
	// }
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_RangePr() {
//Attributes
	this.autoStart = true;
	this.autoEnd = true;
	this.groupBy = c_oAscGroupBy.Range;
	this.startNum = null;
	this.endNum = null;
	this.startDate = null;
	this.endDate = null;
	this.groupInterval = 1;
}
CT_RangePr.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["autoStart"];
		if (undefined !== val) {
			this.autoStart = AscCommon.getBoolFromXml(val);
		}
		val = vals["autoEnd"];
		if (undefined !== val) {
			this.autoEnd = AscCommon.getBoolFromXml(val);
		}
		val = vals["groupBy"];
		if (undefined !== val) {
			val = FromXml_ST_GroupBy(val);
			if (-1 !== val) {
				this.groupBy = val;
			}
		}
		val = vals["startNum"];
		if (undefined !== val) {
			this.startNum = val - 0;
		}
		val = vals["endNum"];
		if (undefined !== val) {
			this.endNum = val - 0;
		}
		val = vals["startDate"];
		if (undefined !== val) {
			this.startDate = AscCommon.unleakString(uq(val));
		}
		val = vals["endDate"];
		if (undefined !== val) {
			this.endDate = AscCommon.unleakString(uq(val));
		}
		val = vals["groupInterval"];
		if (undefined !== val) {
			this.groupInterval = val - 0;
		}
	}
};
CT_RangePr.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (true !== this.autoStart) {
		writer.WriteXmlAttributeBool("autoStart", this.autoStart);
	}
	if (true !== this.autoEnd) {
		writer.WriteXmlAttributeBool("autoEnd", this.autoEnd);
	}
	if (c_oAscGroupBy.Range !== this.groupBy) {
		writer.WriteXmlAttributeStringEncode("groupBy", ToXml_ST_GroupBy(this.groupBy));
	}
	if (null !== this.startNum) {
		writer.WriteXmlAttributeNumber("startNum", this.startNum);
	}
	if (null !== this.endNum) {
		writer.WriteXmlAttributeNumber("endNum", this.endNum);
	}
	if (null !== this.startDate) {
		writer.WriteXmlAttributeStringEncode("startDate", this.startDate);
	}
	if (null !== this.endDate) {
		writer.WriteXmlAttributeStringEncode("endDate", this.endDate);
	}
	if (1 !== this.groupInterval) {
		writer.WriteXmlAttributeNumber("groupInterval", this.groupInterval);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_DiscretePr() {
//Attributes
//	this.count = null;
//Members
	this.x = [];
}
CT_DiscretePr.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("x" === elem) {
		newContext = new CT_Index();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.x.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_DiscretePr.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.x.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.x.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.x.length; ++i) {
		var elem = this.x[i];
		elem.toXml(writer, "x");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_GroupItems() {
//Attributes
//	this.count = null;
//Members
	this.Items = new PivotRecords();
}
CT_GroupItems.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	var newContextCandidate = this.Items.onStartNode(elem, attr, uq);
	if (newContextCandidate) {
		newContext = newContextCandidate;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_GroupItems.prototype.onEndNode = function(prevContext, elem) {
	this.Items.onEndNode(prevContext, elem);
};
CT_GroupItems.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	var count = this.Items.getSize();
	if (count > 0) {
		writer.WriteXmlAttributeNumber("count", count);
	}
	writer.WriteXmlNodeEnd(name, true);
	this.Items.toXml(writer);
	writer.WriteXmlNodeEnd(name);
};
function CT_FieldUsage() {
//Attributes
	this.x = null;
}
CT_FieldUsage.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["x"];
		if (undefined !== val) {
			this.x = val - 0;
		}
	}
};
CT_FieldUsage.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.x) {
		writer.WriteXmlAttributeNumber("x", this.x);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_GroupLevel() {
//Attributes
	this.uniqueName = null;
	this.caption = null;
	this.user = false;
	this.customRollUp = false;
//Members
	this.groups = null;
	this.extLst = null;
}
CT_GroupLevel.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["uniqueName"];
		if (undefined !== val) {
			this.uniqueName = AscCommon.unleakString(uq(val));
		}
		val = vals["caption"];
		if (undefined !== val) {
			this.caption = AscCommon.unleakString(uq(val));
		}
		val = vals["user"];
		if (undefined !== val) {
			this.user = AscCommon.getBoolFromXml(val);
		}
		val = vals["customRollUp"];
		if (undefined !== val) {
			this.customRollUp = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_GroupLevel.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("groups" === elem) {
		newContext = new CT_Groups();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.groups = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_GroupLevel.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.uniqueName) {
		writer.WriteXmlAttributeStringEncode("uniqueName", this.uniqueName);
	}
	if (null !== this.caption) {
		writer.WriteXmlAttributeStringEncode("caption", this.caption);
	}
	if (false !== this.user) {
		writer.WriteXmlAttributeBool("user", this.user);
	}
	if (false !== this.customRollUp) {
		writer.WriteXmlAttributeBool("customRollUp", this.customRollUp);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.groups) {
		this.groups.toXml(writer, "groups");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_PivotAreaReferences() {
//Attributes
//	this.count = null;
//Members
	this.reference = [];
}
CT_PivotAreaReferences.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("reference" === elem) {
		newContext = new CT_PivotAreaReference();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.reference.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PivotAreaReferences.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.reference.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.reference.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.reference.length; ++i) {
		var elem = this.reference[i];
		elem.toXml(writer, "reference");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_Item() {
//Attributes
	this.n = null;
	this.t = Asc.c_oAscItemType.Data;
	this.h = false;
	this.s = false;
	this.sd = true;
	this.f = false;
	this.m = false;
	this.c = false;
	this.x = null;
	this.d = false;
	this.e = true;
}
CT_Item.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["n"];
		if (undefined !== val) {
			this.n = AscCommon.unleakString(uq(val));
		}
		val = vals["t"];
		if (undefined !== val) {
			val = FromXml_ST_ItemType(val);
			if (-1 !== val) {
				this.t = val;
			}
		}
		val = vals["h"];
		if (undefined !== val) {
			this.h = AscCommon.getBoolFromXml(val);
		}
		val = vals["s"];
		if (undefined !== val) {
			this.s = AscCommon.getBoolFromXml(val);
		}
		val = vals["sd"];
		if (undefined !== val) {
			this.sd = AscCommon.getBoolFromXml(val);
		}
		val = vals["f"];
		if (undefined !== val) {
			this.f = AscCommon.getBoolFromXml(val);
		}
		val = vals["m"];
		if (undefined !== val) {
			this.m = AscCommon.getBoolFromXml(val);
		}
		val = vals["c"];
		if (undefined !== val) {
			this.c = AscCommon.getBoolFromXml(val);
		}
		val = vals["x"];
		if (undefined !== val) {
			this.x = val - 0;
		}
		val = vals["d"];
		if (undefined !== val) {
			this.d = AscCommon.getBoolFromXml(val);
		}
		val = vals["e"];
		if (undefined !== val) {
			this.e = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_Item.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.n) {
		writer.WriteXmlAttributeStringEncode("n", this.n);
	}
	if (Asc.c_oAscItemType.Data !== this.t) {
		writer.WriteXmlAttributeStringEncode("t", ToXml_ST_ItemType(this.t));
	}
	if (false !== this.h) {
		writer.WriteXmlAttributeBool("h", this.h);
	}
	if (false !== this.s) {
		writer.WriteXmlAttributeBool("s", this.s);
	}
	if (true !== this.sd) {
		writer.WriteXmlAttributeBool("sd", this.sd);
	}
	if (false !== this.f) {
		writer.WriteXmlAttributeBool("f", this.f);
	}
	if (false !== this.m) {
		writer.WriteXmlAttributeBool("m", this.m);
	}
	if (false !== this.c) {
		writer.WriteXmlAttributeBool("c", this.c);
	}
	if (null !== this.x) {
		writer.WriteXmlAttributeNumber("x", this.x);
	}
	if (false !== this.d) {
		writer.WriteXmlAttributeBool("d", this.d);
	}
	if (true !== this.e) {
		writer.WriteXmlAttributeBool("e", this.e);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_MemberProperty() {
//Attributes
	this.name = null;
	this.showCell = false;
	this.showTip = false;
	this.showAsCaption = false;
	this.nameLen = null;
	this.pPos = null;
	this.pLen = null;
	this.level = null;
	this.field = null;
}
CT_MemberProperty.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["showCell"];
		if (undefined !== val) {
			this.showCell = AscCommon.getBoolFromXml(val);
		}
		val = vals["showTip"];
		if (undefined !== val) {
			this.showTip = AscCommon.getBoolFromXml(val);
		}
		val = vals["showAsCaption"];
		if (undefined !== val) {
			this.showAsCaption = AscCommon.getBoolFromXml(val);
		}
		val = vals["nameLen"];
		if (undefined !== val) {
			this.nameLen = val - 0;
		}
		val = vals["pPos"];
		if (undefined !== val) {
			this.pPos = val - 0;
		}
		val = vals["pLen"];
		if (undefined !== val) {
			this.pLen = val - 0;
		}
		val = vals["level"];
		if (undefined !== val) {
			this.level = val - 0;
		}
		val = vals["field"];
		if (undefined !== val) {
			this.field = val - 0;
		}
	}
};
CT_MemberProperty.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (false !== this.showCell) {
		writer.WriteXmlAttributeBool("showCell", this.showCell);
	}
	if (false !== this.showTip) {
		writer.WriteXmlAttributeBool("showTip", this.showTip);
	}
	if (false !== this.showAsCaption) {
		writer.WriteXmlAttributeBool("showAsCaption", this.showAsCaption);
	}
	if (null !== this.nameLen) {
		writer.WriteXmlAttributeNumber("nameLen", this.nameLen);
	}
	if (null !== this.pPos) {
		writer.WriteXmlAttributeNumber("pPos", this.pPos);
	}
	if (null !== this.pLen) {
		writer.WriteXmlAttributeNumber("pLen", this.pLen);
	}
	if (null !== this.level) {
		writer.WriteXmlAttributeNumber("level", this.level);
	}
	if (null !== this.field) {
		writer.WriteXmlAttributeNumber("field", this.field);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_Member() {
//Attributes
	this.name = null;
}
CT_Member.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
	}
};
CT_Member.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_FilterColumn() {
//Attributes
	this.colId = null;
	this.hiddenButton = false;
	this.showButton = true;
//Members
	this.colorFilter = null;
	this.customFilters = null;
	this.dynamicFilter = null;
	this.extLst = null;
	this.filters = null;
	this.iconFilter = null;
	this.top10 = null;
}
CT_FilterColumn.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["colId"];
		if (undefined !== val) {
			this.colId = val - 0;
		}
		val = vals["hiddenButton"];
		if (undefined !== val) {
			this.hiddenButton = AscCommon.getBoolFromXml(val);
		}
		val = vals["showButton"];
		if (undefined !== val) {
			this.showButton = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_FilterColumn.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("colorFilter" === elem) {
		newContext = new CT_ColorFilter();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.colorFilter = newContext;
	} else if ("customFilters" === elem) {
		newContext = new CT_CustomFilters();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.customFilters = newContext;
	} else if ("dynamicFilter" === elem) {
		newContext = new CT_DynamicFilter();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.dynamicFilter = newContext;
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else if ("filters" === elem) {
		newContext = new CT_Filters();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.filters = newContext;
	} else if ("iconFilter" === elem) {
		newContext = new CT_IconFilter();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.iconFilter = newContext;
	} else if ("top10" === elem) {
		newContext = new CT_Top10();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.top10 = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_FilterColumn.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.colId) {
		writer.WriteXmlAttributeNumber("colId", this.colId);
	}
	if (false !== this.hiddenButton) {
		writer.WriteXmlAttributeBool("hiddenButton", this.hiddenButton);
	}
	if (true !== this.showButton) {
		writer.WriteXmlAttributeBool("showButton", this.showButton);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.colorFilter) {
		this.colorFilter.toXml(writer, "colorFilter");
	}
	if (null !== this.customFilters) {
		this.customFilters.toXml(writer, "customFilters");
	}
	if (null !== this.dynamicFilter) {
		this.dynamicFilter.toXml(writer, "dynamicFilter");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	if (null !== this.filters) {
		this.filters.toXml(writer, "filters");
	}
	if (null !== this.iconFilter) {
		this.iconFilter.toXml(writer, "iconFilter");
	}
	if (null !== this.top10) {
		this.top10.toXml(writer, "top10");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_SortState() {
//Attributes
	this.columnSort = false;
	this.caseSensitive = false;
	this.sortMethod = c_oAscSortMethod.None;
	this.ref = null;
//Members
	this.sortCondition = [];
	this.extLst = null;
}
CT_SortState.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["columnSort"];
		if (undefined !== val) {
			this.columnSort = AscCommon.getBoolFromXml(val);
		}
		val = vals["caseSensitive"];
		if (undefined !== val) {
			this.caseSensitive = AscCommon.getBoolFromXml(val);
		}
		val = vals["sortMethod"];
		if (undefined !== val) {
			val = FromXml_ST_SortMethod(val);
			if (-1 !== val) {
				this.sortMethod = val;
			}
		}
		val = vals["ref"];
		if (undefined !== val) {
			this.ref = AscCommon.unleakString(uq(val));
		}
	}
};
CT_SortState.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("sortCondition" === elem) {
		newContext = new CT_SortCondition();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.sortCondition.push(newContext);
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_SortState.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (false !== this.columnSort) {
		writer.WriteXmlAttributeBool("columnSort", this.columnSort);
	}
	if (false !== this.caseSensitive) {
		writer.WriteXmlAttributeBool("caseSensitive", this.caseSensitive);
	}
	if (c_oAscSortMethod.None !== this.sortMethod) {
		writer.WriteXmlAttributeStringEncode("sortMethod", ToXml_ST_SortMethod(this.sortMethod));
	}
	if (null !== this.ref) {
		writer.WriteXmlAttributeStringEncode("ref", this.ref);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.sortCondition.length; ++i) {
		var elem = this.sortCondition[i];
		elem.toXml(writer, "sortCondition");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_PageItem() {
//Attributes
	this.name = null;
}
CT_PageItem.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
	}
};
CT_PageItem.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_Groups() {
//Attributes
//	this.count = null;
//Members
	this.group = [];
}
CT_Groups.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("group" === elem) {
		newContext = new CT_LevelGroup();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.group.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Groups.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.group.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.group.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.group.length; ++i) {
		var elem = this.group[i];
		elem.toXml(writer, "group");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_PivotAreaReference() {
//Attributes
	this.field = null;
//	this.count = null;
	this.selected = true;
	this.byPosition = false;
	this.relative = false;
	this.defaultSubtotal = false;
	this.sumSubtotal = false;
	this.countASubtotal = false;
	this.avgSubtotal = false;
	this.maxSubtotal = false;
	this.minSubtotal = false;
	this.productSubtotal = false;
	this.countSubtotal = false;
	this.stdDevSubtotal = false;
	this.stdDevPSubtotal = false;
	this.varSubtotal = false;
	this.varPSubtotal = false;
//Members
	this.x = [];
	this.extLst = null;
}
CT_PivotAreaReference.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["field"];
		if (undefined !== val) {
			this.field = val - 0;
		}
		val = vals["selected"];
		if (undefined !== val) {
			this.selected = AscCommon.getBoolFromXml(val);
		}
		val = vals["byPosition"];
		if (undefined !== val) {
			this.byPosition = AscCommon.getBoolFromXml(val);
		}
		val = vals["relative"];
		if (undefined !== val) {
			this.relative = AscCommon.getBoolFromXml(val);
		}
		val = vals["defaultSubtotal"];
		if (undefined !== val) {
			this.defaultSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["sumSubtotal"];
		if (undefined !== val) {
			this.sumSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["countASubtotal"];
		if (undefined !== val) {
			this.countASubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["avgSubtotal"];
		if (undefined !== val) {
			this.avgSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["maxSubtotal"];
		if (undefined !== val) {
			this.maxSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["minSubtotal"];
		if (undefined !== val) {
			this.minSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["productSubtotal"];
		if (undefined !== val) {
			this.productSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["countSubtotal"];
		if (undefined !== val) {
			this.countSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["stdDevSubtotal"];
		if (undefined !== val) {
			this.stdDevSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["stdDevPSubtotal"];
		if (undefined !== val) {
			this.stdDevPSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["varSubtotal"];
		if (undefined !== val) {
			this.varSubtotal = AscCommon.getBoolFromXml(val);
		}
		val = vals["varPSubtotal"];
		if (undefined !== val) {
			this.varPSubtotal = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_PivotAreaReference.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("x" === elem) {
		newContext = new CT_Index();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.x.push(newContext);
	} else if ("extLst" === elem) {
		newContext = new CT_ExtensionList();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.extLst = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_PivotAreaReference.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.field) {
		writer.WriteXmlAttributeNumber("field", this.field);
	}
	if (this.x.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.x.length);
	}
	if (true !== this.selected) {
		writer.WriteXmlAttributeBool("selected", this.selected);
	}
	if (false !== this.byPosition) {
		writer.WriteXmlAttributeBool("byPosition", this.byPosition);
	}
	if (false !== this.relative) {
		writer.WriteXmlAttributeBool("relative", this.relative);
	}
	if (false !== this.defaultSubtotal) {
		writer.WriteXmlAttributeBool("defaultSubtotal", this.defaultSubtotal);
	}
	if (false !== this.sumSubtotal) {
		writer.WriteXmlAttributeBool("sumSubtotal", this.sumSubtotal);
	}
	if (false !== this.countASubtotal) {
		writer.WriteXmlAttributeBool("countASubtotal", this.countASubtotal);
	}
	if (false !== this.avgSubtotal) {
		writer.WriteXmlAttributeBool("avgSubtotal", this.avgSubtotal);
	}
	if (false !== this.maxSubtotal) {
		writer.WriteXmlAttributeBool("maxSubtotal", this.maxSubtotal);
	}
	if (false !== this.minSubtotal) {
		writer.WriteXmlAttributeBool("minSubtotal", this.minSubtotal);
	}
	if (false !== this.productSubtotal) {
		writer.WriteXmlAttributeBool("productSubtotal", this.productSubtotal);
	}
	if (false !== this.countSubtotal) {
		writer.WriteXmlAttributeBool("countSubtotal", this.countSubtotal);
	}
	if (false !== this.stdDevSubtotal) {
		writer.WriteXmlAttributeBool("stdDevSubtotal", this.stdDevSubtotal);
	}
	if (false !== this.stdDevPSubtotal) {
		writer.WriteXmlAttributeBool("stdDevPSubtotal", this.stdDevPSubtotal);
	}
	if (false !== this.varSubtotal) {
		writer.WriteXmlAttributeBool("varSubtotal", this.varSubtotal);
	}
	if (false !== this.varPSubtotal) {
		writer.WriteXmlAttributeBool("varPSubtotal", this.varPSubtotal);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.x.length; ++i) {
		var elem = this.x[i];
		elem.toXml(writer, "x");
	}
	if (null !== this.extLst) {
		this.extLst.toXml(writer, "extLst");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_ColorFilter() {
//Attributes
	this.dxfId = null;
	this.cellColor = true;
}
CT_ColorFilter.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["dxfId"];
		if (undefined !== val) {
			this.dxfId = val - 0;
		}
		val = vals["cellColor"];
		if (undefined !== val) {
			this.cellColor = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_ColorFilter.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	//todo
	// if (null !== this.dxfId) {
	// 	writer.WriteXmlAttributeNumber("dxfId", this.dxfId);
	// }
	if (true !== this.cellColor) {
		writer.WriteXmlAttributeBool("cellColor", this.cellColor);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_CustomFilters() {
//Attributes
	this.and = false;
//Members
	this.customFilter = [];
}
CT_CustomFilters.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["and"];
		if (undefined !== val) {
			this.and = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_CustomFilters.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("customFilter" === elem) {
		newContext = new CT_CustomFilter();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.customFilter.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_CustomFilters.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (false !== this.and) {
		writer.WriteXmlAttributeBool("and", this.and);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.customFilter.length; ++i) {
		var elem = this.customFilter[i];
		elem.toXml(writer, "customFilter");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_DynamicFilter() {
//Attributes
	this.type = null;
	this.val = null;
	this.valIso = null;
	this.maxValIso = null;
}
CT_DynamicFilter.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["type"];
		if (undefined !== val) {
			val = FromXml_ST_DynamicFilterType(val);
			if (-1 !== val) {
				this.type = val;
			}
		}
		val = vals["val"];
		if (undefined !== val) {
			this.val = val - 0;
		}
		val = vals["valIso"];
		if (undefined !== val) {
			this.valIso = AscCommon.unleakString(uq(val));
		}
		val = vals["maxValIso"];
		if (undefined !== val) {
			this.maxValIso = AscCommon.unleakString(uq(val));
		}
	}
};
CT_DynamicFilter.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.type) {
		writer.WriteXmlAttributeStringEncode("type", ToXml_ST_DynamicFilterType(this.type));
	}
	if (null !== this.val) {
		writer.WriteXmlAttributeNumber("val", this.val);
	}
	if (null !== this.valIso) {
		writer.WriteXmlAttributeStringEncode("valIso", this.valIso);
	}
	if (null !== this.maxValIso) {
		writer.WriteXmlAttributeStringEncode("maxValIso", this.maxValIso);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_Filters() {
//Attributes
	this.blank = false;
	this.calendarType = c_oAscCalendarType.None;
//Members
	this.filter = [];
	this.dateGroupItem = [];
}
CT_Filters.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["blank"];
		if (undefined !== val) {
			this.blank = AscCommon.getBoolFromXml(val);
		}
		val = vals["calendarType"];
		if (undefined !== val) {
			val = FromXml_ST_CalendarType(val);
			if (-1 !== val) {
				this.calendarType = val;
			}
		}
	}
};
CT_Filters.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("filter" === elem) {
		newContext = new CT_Filter();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.filter.push(newContext);
	} else if ("dateGroupItem" === elem) {
		newContext = new CT_DateGroupItem();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.dateGroupItem.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_Filters.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (false !== this.blank) {
		writer.WriteXmlAttributeBool("blank", this.blank);
	}
	if (c_oAscCalendarType.None !== this.calendarType) {
		writer.WriteXmlAttributeStringEncode("calendarType", ToXml_ST_CalendarType(this.calendarType));
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.filter.length; ++i) {
		var elem = this.filter[i];
		elem.toXml(writer, "filter");
	}
	for (var i = 0; i < this.dateGroupItem.length; ++i) {
		var elem = this.dateGroupItem[i];
		elem.toXml(writer, "dateGroupItem");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_IconFilter() {
//Attributes
	this.iconSet = null;
	this.iconId = null;
}
CT_IconFilter.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["iconSet"];
		if (undefined !== val) {
			val = FromXml_ST_IconSetType(val);
			if (-1 !== val) {
				this.iconSet = val;
			}
		}
		val = vals["iconId"];
		if (undefined !== val) {
			this.iconId = val - 0;
		}
	}
};
CT_IconFilter.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.iconSet) {
		writer.WriteXmlAttributeStringEncode("iconSet", ToXml_ST_IconSetType(this.iconSet));
	}
	if (null !== this.iconId) {
		writer.WriteXmlAttributeNumber("iconId", this.iconId);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_Top10() {
//Attributes
	this.top = true;
	this.percent = false;
	this.val = null;
	this.filterVal = null;
}
CT_Top10.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["top"];
		if (undefined !== val) {
			this.top = AscCommon.getBoolFromXml(val);
		}
		val = vals["percent"];
		if (undefined !== val) {
			this.percent = AscCommon.getBoolFromXml(val);
		}
		val = vals["val"];
		if (undefined !== val) {
			this.val = val - 0;
		}
		val = vals["filterVal"];
		if (undefined !== val) {
			this.filterVal = val - 0;
		}
	}
};
CT_Top10.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (true !== this.top) {
		writer.WriteXmlAttributeBool("top", this.top);
	}
	if (false !== this.percent) {
		writer.WriteXmlAttributeBool("percent", this.percent);
	}
	if (null !== this.val) {
		writer.WriteXmlAttributeNumber("val", this.val);
	}
	if (null !== this.filterVal) {
		writer.WriteXmlAttributeNumber("filterVal", this.filterVal);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_SortCondition() {
//Attributes
	this.descending = false;
	this.sortBy = c_oAscSortBy.Value;
	this.ref = null;
	this.customList = null;
	this.dxfId = null;
	this.iconSet = c_oAscIconSetType.ThreeArrows;
	this.iconId = null;
}
CT_SortCondition.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["descending"];
		if (undefined !== val) {
			this.descending = AscCommon.getBoolFromXml(val);
		}
		val = vals["sortBy"];
		if (undefined !== val) {
			val = FromXml_ST_SortBy(val);
			if (-1 !== val) {
				this.sortBy = val;
			}
		}
		val = vals["ref"];
		if (undefined !== val) {
			this.ref = AscCommon.unleakString(uq(val));
		}
		val = vals["customList"];
		if (undefined !== val) {
			this.customList = AscCommon.unleakString(uq(val));
		}
		val = vals["dxfId"];
		if (undefined !== val) {
			this.dxfId = val - 0;
		}
		val = vals["iconSet"];
		if (undefined !== val) {
			val = FromXml_ST_IconSetType(val);
			if (-1 !== val) {
				this.iconSet = val;
			}
		}
		val = vals["iconId"];
		if (undefined !== val) {
			this.iconId = val - 0;
		}
	}
};
CT_SortCondition.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (false !== this.descending) {
		writer.WriteXmlAttributeBool("descending", this.descending);
	}
	if (c_oAscSortBy.Value !== this.sortBy) {
		writer.WriteXmlAttributeStringEncode("sortBy", ToXml_ST_SortBy(this.sortBy));
	}
	if (null !== this.ref) {
		writer.WriteXmlAttributeStringEncode("ref", this.ref);
	}
	if (null !== this.customList) {
		writer.WriteXmlAttributeStringEncode("customList", this.customList);
	}
	//todo
	// if (null !== this.dxfId) {
	// 	writer.WriteXmlAttributeNumber("dxfId", this.dxfId);
	// }
	if (c_oAscIconSetType.ThreeArrows !== this.iconSet) {
		writer.WriteXmlAttributeStringEncode("iconSet", ToXml_ST_IconSetType(this.iconSet));
	}
	if (null !== this.iconId) {
		writer.WriteXmlAttributeNumber("iconId", this.iconId);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_LevelGroup() {
//Attributes
	this.name = null;
	this.uniqueName = null;
	this.caption = null;
	this.uniqueParent = null;
	this.id = null;
//Members
	this.groupMembers = null;
}
CT_LevelGroup.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["name"];
		if (undefined !== val) {
			this.name = AscCommon.unleakString(uq(val));
		}
		val = vals["uniqueName"];
		if (undefined !== val) {
			this.uniqueName = AscCommon.unleakString(uq(val));
		}
		val = vals["caption"];
		if (undefined !== val) {
			this.caption = AscCommon.unleakString(uq(val));
		}
		val = vals["uniqueParent"];
		if (undefined !== val) {
			this.uniqueParent = AscCommon.unleakString(uq(val));
		}
		val = vals["id"];
		if (undefined !== val) {
			this.id = val - 0;
		}
	}
};
CT_LevelGroup.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("groupMembers" === elem) {
		newContext = new CT_GroupMembers();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.groupMembers = newContext;
	} else {
		newContext = null;
	}
	return newContext;
};
CT_LevelGroup.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.name) {
		writer.WriteXmlAttributeStringEncode("name", this.name);
	}
	if (null !== this.uniqueName) {
		writer.WriteXmlAttributeStringEncode("uniqueName", this.uniqueName);
	}
	if (null !== this.caption) {
		writer.WriteXmlAttributeStringEncode("caption", this.caption);
	}
	if (null !== this.uniqueParent) {
		writer.WriteXmlAttributeStringEncode("uniqueParent", this.uniqueParent);
	}
	if (null !== this.id) {
		writer.WriteXmlAttributeNumber("id", this.id);
	}
	writer.WriteXmlNodeEnd(name, true);
	if (null !== this.groupMembers) {
		this.groupMembers.toXml(writer, "groupMembers");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_CustomFilter() {
//Attributes
	this.operator = c_oAscFilterOperator.Equal;
	this.val = null;
}
CT_CustomFilter.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["operator"];
		if (undefined !== val) {
			val = FromXml_ST_FilterOperator(val);
			if (-1 !== val) {
				this.operator = val;
			}
		}
		val = vals["val"];
		if (undefined !== val) {
			this.val = AscCommon.unleakString(uq(val));
		}
	}
};
CT_CustomFilter.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (c_oAscFilterOperator.Equal !== this.operator) {
		writer.WriteXmlAttributeStringEncode("operator", ToXml_ST_FilterOperator(this.operator));
	}
	if (null !== this.val) {
		writer.WriteXmlAttributeStringEncode("val", this.val);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_Filter() {
//Attributes
	this.val = null;
}
CT_Filter.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["val"];
		if (undefined !== val) {
			this.val = AscCommon.unleakString(uq(val));
		}
	}
};
CT_Filter.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.val) {
		writer.WriteXmlAttributeStringEncode("val", this.val);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_DateGroupItem() {
//Attributes
	this.year = null;
	this.month = null;
	this.day = null;
	this.hour = null;
	this.minute = null;
	this.second = null;
	this.dateTimeGrouping = null;
}
CT_DateGroupItem.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["year"];
		if (undefined !== val) {
			this.year = val - 0;
		}
		val = vals["month"];
		if (undefined !== val) {
			this.month = val - 0;
		}
		val = vals["day"];
		if (undefined !== val) {
			this.day = val - 0;
		}
		val = vals["hour"];
		if (undefined !== val) {
			this.hour = val - 0;
		}
		val = vals["minute"];
		if (undefined !== val) {
			this.minute = val - 0;
		}
		val = vals["second"];
		if (undefined !== val) {
			this.second = val - 0;
		}
		val = vals["dateTimeGrouping"];
		if (undefined !== val) {
			val = FromXml_ST_DateTimeGrouping(val);
			if (-1 !== val) {
				this.dateTimeGrouping = val;
			}
		}
	}
};
CT_DateGroupItem.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.year) {
		writer.WriteXmlAttributeNumber("year", this.year);
	}
	if (null !== this.month) {
		writer.WriteXmlAttributeNumber("month", this.month);
	}
	if (null !== this.day) {
		writer.WriteXmlAttributeNumber("day", this.day);
	}
	if (null !== this.hour) {
		writer.WriteXmlAttributeNumber("hour", this.hour);
	}
	if (null !== this.minute) {
		writer.WriteXmlAttributeNumber("minute", this.minute);
	}
	if (null !== this.second) {
		writer.WriteXmlAttributeNumber("second", this.second);
	}
	if (null !== this.dateTimeGrouping) {
		writer.WriteXmlAttributeStringEncode("dateTimeGrouping", ToXml_ST_DateTimeGrouping(this.dateTimeGrouping));
	}
	writer.WriteXmlNodeEnd(name, true, true);
};
function CT_GroupMembers() {
//Attributes
//	this.count = null;
//Members
	this.groupMember = [];
}
CT_GroupMembers.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = this;
	if ("groupMember" === elem) {
		newContext = new CT_GroupMember();
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
		this.groupMember.push(newContext);
	} else {
		newContext = null;
	}
	return newContext;
};
CT_GroupMembers.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (this.groupMember.length > 0) {
		writer.WriteXmlAttributeNumber("count", this.groupMember.length);
	}
	writer.WriteXmlNodeEnd(name, true);
	for (var i = 0; i < this.groupMember.length; ++i) {
		var elem = this.groupMember[i];
		elem.toXml(writer, "groupMember");
	}
	writer.WriteXmlNodeEnd(name);
};
function CT_GroupMember() {
//Attributes
	this.uniqueName = null;
	this.group = false;
}
CT_GroupMember.prototype.readAttributes = function(attr, uq) {
	if (attr()) {
		var vals = attr();
		var val;
		val = vals["uniqueName"];
		if (undefined !== val) {
			this.uniqueName = AscCommon.unleakString(uq(val));
		}
		val = vals["group"];
		if (undefined !== val) {
			this.group = AscCommon.getBoolFromXml(val);
		}
	}
};
CT_GroupMember.prototype.toXml = function(writer, name) {
	writer.WriteXmlNodeStart(name);
	if (null !== this.uniqueName) {
		writer.WriteXmlAttributeStringEncode("uniqueName", this.uniqueName);
	}
	if (false !== this.group) {
		writer.WriteXmlAttributeBool("group", this.group);
	}
	writer.WriteXmlNodeEnd(name, true, true);
};

var c_oAscPivotRecType = {
	Number: 1,
	String: 2,
	Boolean: 3,
	Error: 4,
	DateTime: 5,
	Missing: 6,
	Index: 7
};
var c_nNumberMissingValue =  2147483647;//Math.pow(2, 31) - 1

function PivotRecordValue() {
	this.clean();
};
PivotRecordValue.prototype.clean = function() {
	this.type = undefined;
	this.val = undefined;
	this.addition = undefined;
}
PivotRecordValue.prototype.getCellValue = function() {
	var res = new AscCommonExcel.CCellValue();
	switch (this.type) {
		case c_oAscPivotRecType.Boolean:
			res.type = AscCommon.CellValueType.Number;
			res.number = this.val ? 1 : 0;
			break;
		case c_oAscPivotRecType.DateTime:
			res.type = AscCommon.CellValueType.Number;
			res.number = this.val;
			break;
		case c_oAscPivotRecType.Error:
			res.type = AscCommon.CellValueType.String;
			res.text = AscCommonExcel.cError.prototype.getStringFromErrorType(this.val);
			break;
		case c_oAscPivotRecType.Number:
			res.type = AscCommon.CellValueType.Number;
			res.number = this.val;
			break;
		case c_oAscPivotRecType.String:
			res.type = AscCommon.CellValueType.String;
			res.text = this.val;
			break;
		default:
			var i = 0;
			break;
	}
	return res;
};
PivotRecordValue.prototype.shallowEqual = function(elem) {
	return this.type === elem.type && this.val === elem.val;
};

function PivotRecords() {
	this.chunks = [];
	this.addition = {};
	this.size = 0;
	this.startCount = 0;

//inner
	this._curBoolean = new CT_Boolean();
	this._curDateTime = new CT_DateTime();
	this._curError = new CT_Error();
	this._curMissing = new CT_Missing();
	this._curNumber = new CT_Number();
	this._curString = new CT_String();
	this._curIndex = new CT_Index();
	this.output = new PivotRecordValue();
}

PivotRecords.prototype.onStartNode = function(elem, attr, uq) {
	var newContext = null;
	if ("b" === elem) {
		this._curBoolean.clean();
		newContext = this._curBoolean;
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
	} else if ("d" === elem) {
		this._curDateTime.clean();
		newContext = this._curDateTime;
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
	} else if ("e" === elem) {
		this._curError.clean();
		newContext = this._curError;
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
	} else if ("m" === elem) {
		this._curMissing.clean();
		newContext = this._curMissing;
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
	} else if ("n" === elem) {
		this._curNumber.clean();
		newContext = this._curNumber;
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
	} else if ("s" === elem) {
		this._curString.clean();
		newContext = this._curString;
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
	} else if ("x" === elem) {
		this._curIndex.clean();
		newContext = this._curIndex;
		if (newContext.readAttributes) {
			newContext.readAttributes(attr, uq);
		}
	}
	return newContext;
};
PivotRecords.prototype.onEndNode = function(prevContext, elem) {
	var res = true;
	if ("b" === elem) {
		if (this._curBoolean.isSimpleValue()) {
			this.addBool(this._curBoolean.v);
		} else {
			this.addBool(this._curBoolean.v, this._curBoolean);
			this._curBoolean = new CT_Boolean();
		}
	} else if ("d" === elem) {
		if (this._curDateTime.isSimpleValue()) {
			this.addDate(this._curDateTime.v);
		} else {
			this.addDate(this._curDateTime.v, this._curDateTime);
			this._curDateTime = new CT_DateTime();
		}
	} else if ("e" === elem) {
		if (this._curError.isSimpleValue()) {
			this.addError(this._curError.v);
		} else {
			this.addError(this._curError.v, this._curError);
			this._curError = new CT_Error();
		}
	} else if ("m" === elem) {
		if (this._curMissing.isSimpleValue()) {
			this.addMissing(1);
		} else {
			this.addMissing(1, this._curMissing);
			this._curMissing = new CT_Missing();
		}
	} else if ("n" === elem) {
		if (this._curNumber.isSimpleValue()) {
			this.addNumber(this._curNumber.v);
		} else {
			this.addNumber(this._curNumber.v, this._curNumber);
			this._curNumber = new CT_Number();
		}
	} else if ("s" === elem) {
		if (this._curString.isSimpleValue()) {
			this.addString(this._curString.v);
		} else {
			this.addString(this._curString.v, this._curString);
			this._curString = new CT_String();
		}
	} else if ("x" === elem) {
		this.addIndex(this._curIndex.v);
	} else {
		res = false;
	}
	return res;
};
PivotRecords.prototype.toXml = function(writer, opt_index) {
	if (undefined !== opt_index) {
		this._toXml(writer, this.get(opt_index));
	} else {
		for (var i = 0; i < this.size; ++i) {
			this._toXml(writer, this.get(i));
		}
	}
};
PivotRecords.prototype.setStartCount = function(val) {
	this.startCount = val;
};
PivotRecords.prototype.getSize = function() {
	return this.size;
};
PivotRecords.prototype.addBool = function(val, addition) {
	this._add(c_oAscPivotRecType.Boolean, val, addition);
};
PivotRecords.prototype.addDate = function(val, addition) {
	this._add(c_oAscPivotRecType.DateTime, val, addition);
};
PivotRecords.prototype.addError = function(val, addition) {
	this._add(c_oAscPivotRecType.Error, val, addition);
};
PivotRecords.prototype.addMissing = function(count, addition) {
	//todo without cycle
	for (var i = 0; i < count; ++i) {
		this._add(c_oAscPivotRecType.Missing, undefined, addition);
	}
};
PivotRecords.prototype.addNumber = function(val, addition) {
	this._add(c_oAscPivotRecType.Number, val, addition);
};
PivotRecords.prototype.addString = function(val, addition) {
	//todo don't use global editor
	val = window["Asc"]["editor"].wbModel.sharedStrings.addText(val);
	this._add(c_oAscPivotRecType.String, val, addition);
};
PivotRecords.prototype.addIndex = function(val) {
	val++;//0 is missing value
	this._add(c_oAscPivotRecType.Index, val);
};
PivotRecords.prototype.get = function(index) {
	this.output.clean();
	for (var i = 0; i < this.chunks.length; ++i) {
		var chunk = this.chunks[i];
		if (chunk.from <= index && index < chunk.to) {
			this.output.type = chunk.type;
			this.output.addition = this.addition[index];
			if (chunk.data) {
				this.output.val = chunk.data[index - chunk.from];
				this._replaceMissingInOutput(chunk.type, this.output);
			}
			break;
		}
	}
	return this.output;
};
PivotRecords.prototype.convertToSharedItems = function(si) {
	var i, j, chunk, uniqueMap, index, val;
	var uniqueMaps = {missing: undefined, types: {}};
	var newChunk = {type: c_oAscPivotRecType.Index, data: new Uint32Array(this.size), capacity: this.size, from: 0, to: this.size};
	var newChunkIndex = 0;
	for (i = 0; i < this.chunks.length; ++i) {
		chunk = this.chunks[i];
		if (c_oAscPivotRecType.Missing !== chunk.type) {
			uniqueMap = uniqueMaps.types[chunk.type];
			if (!uniqueMap) {
				uniqueMap = new Map();
				uniqueMaps.types[chunk.type] = uniqueMap;
			}
			for (j = 0; j < chunk.to - chunk.from; ++j) {
				index = 0;
				if (chunk.data) {
					val = chunk.data[j];
					if (!this._isMissingValue(chunk.type, val)) {
						index = uniqueMap.get(val);
						if (undefined === index) {
							index = si.Items.getSize() + 1;//0 is missing value
							uniqueMap.set(val, index);
							si.Items._add(chunk.type, val);
						}
					} else {
						if (undefined === uniqueMaps.missing) {
							uniqueMaps.missing = si.Items.getSize() + 1;//0 is missing value
							si.Items._add(chunk.type, val);
						}
						index = uniqueMaps.missing;
					}

				}
				newChunk.data[newChunkIndex++] = index;
			}
		} else {
			if (undefined === uniqueMaps.missing) {
				uniqueMaps.missing = si.Items.getSize() + 1;//0 is missing value
				si.Items._add(c_oAscPivotRecType.Missing);
			}
			for (j = 0; j < chunk.to - chunk.from; ++j) {
				newChunk.data[newChunkIndex++] = uniqueMaps.missing;
			}
		}
	}
	this.chunks = [newChunk];
};
PivotRecords.prototype._add = function(type, val, addition) {
	var index = this.size;
	var prevChunk = this.chunks.length > 0 ? this.chunks[this.chunks.length - 1] : null;
	var chunk;
	var chunkIndex;
	if (prevChunk && (prevChunk.type === type || c_oAscPivotRecType.Missing === type)) {
		chunk = prevChunk;
		if (c_oAscPivotRecType.Missing === type) {
			val = this._getMissingFakeVal(chunk.type);
		} else if (c_oAscPivotRecType.Number === type && c_nNumberMissingValue === val) {
			//because there are much more Missing values than c_nNumberMissingValue.
			if(!addition){
				addition = new CT_Number();
			}
			addition.realNumber = true;
		}
	} else {
		if (prevChunk && prevChunk.type === c_oAscPivotRecType.Missing) {
			//convert Missing chunk
			chunk = prevChunk;
			chunk.type = type;
			chunkIndex = index - chunk.from;
			this._checkChunkSize(chunk, chunkIndex);
			var missingVal = this._getMissingFakeVal(chunk.type);
			chunk.data.fill(missingVal, 0, chunkIndex);
		} else {
			//todo shrink prevChunk
			var from = prevChunk ? prevChunk.to : 0;
			chunk = {type: type, data: null, capacity: 0, from: from, to: from + 1};
			this.chunks.push(chunk);
		}
	}
	chunkIndex = index - chunk.from;
	if (c_oAscPivotRecType.Missing !== chunk.type) {
		this._checkChunkSize(chunk, chunkIndex);
	}
	if (chunk.data) {
		chunk.data[chunkIndex] = val;
	}
	if (addition) {
		this.addition[index] = addition;
	}
	this.size++;
	chunk.to = this.size;
};
PivotRecords.prototype._checkChunkSize = function(chunk, chunkIndex) {
	if (chunkIndex >= chunk.capacity) {
		var oldData = chunk.data;
		var chunkStartCount = (0 === this.size && this.startCount) ? this.startCount - chunk.from : 0;
		var maxSize = Math.max((1.1 * chunk.capacity) >> 0, chunkIndex + 1, chunkStartCount);
		chunk.capacity = Math.min(maxSize, AscCommon.gc_nMaxRow0 + 1);
		switch (chunk.type) {
			case c_oAscPivotRecType.Boolean:
				chunk.data = new Uint8Array(chunk.capacity);
				break;
			case c_oAscPivotRecType.DateTime:
				chunk.data = new Float64Array(chunk.capacity);
				break;
			case c_oAscPivotRecType.Error:
				chunk.data = new Uint8Array(chunk.capacity);
				break;
			case c_oAscPivotRecType.Missing:
				break;
			case c_oAscPivotRecType.Number:
				chunk.data = new Float64Array(chunk.capacity);
				break;
			case c_oAscPivotRecType.String:
				chunk.data = new Uint32Array(chunk.capacity);
				break;
			case c_oAscPivotRecType.Index:
				chunk.data = new Uint32Array(chunk.capacity);
				break;
		}
		if (oldData) {
			chunk.data.set(oldData);
		}
	}
};
PivotRecords.prototype._getMissingFakeVal = function(type) {
	var res;
	switch (type) {
		case c_oAscPivotRecType.Boolean:
			res = 255;
			break;
		case c_oAscPivotRecType.DateTime:
			res = -1;
			break;
		case c_oAscPivotRecType.Error:
			res = 255;
			break;
		case c_oAscPivotRecType.Number:
			res = c_nNumberMissingValue;
			break;
		case c_oAscPivotRecType.String:
			res = 0;
			break;
		case c_oAscPivotRecType.Index:
			res = 0;
			break;
	}
	return res;
};
PivotRecords.prototype._replaceMissingInOutput = function(type, output) {
	switch (type) {
		case c_oAscPivotRecType.Boolean:
			if (255 === output.val) {
				output.type = c_oAscPivotRecType.Missing;
			} else {
				output.val = !!output.val;
			}
			break;
		case c_oAscPivotRecType.DateTime:
			if (-1 === output.val) {
				output.type = c_oAscPivotRecType.Missing;
			}
			break;
		case c_oAscPivotRecType.Error:
			if (255 === output.val) {
				output.type = c_oAscPivotRecType.Missing;
			}
			break;
		case c_oAscPivotRecType.Number:
			if (c_nNumberMissingValue === output.val && !(output.addition && output.addition.realNumber)) {
				output.type = c_oAscPivotRecType.Missing;
			}
			break;
		case c_oAscPivotRecType.String:
			if (0 === output.val) {
				output.type = c_oAscPivotRecType.Missing;
			} else {
				//todo don't use global editor
				output.val = window["Asc"]["editor"].wbModel.sharedStrings.get(output.val);
			}
			break;
		case c_oAscPivotRecType.Index:
			if (0 === output.val) {
				output.type = c_oAscPivotRecType.Missing;
			} else {
				--output.val;
			}
			break;
	}
};
PivotRecords.prototype._isMissingValue = function(type, val) {
	switch (type) {
		case c_oAscPivotRecType.Boolean:
			return 255 === val;
		case c_oAscPivotRecType.DateTime:
			return -1 === val;
		case c_oAscPivotRecType.Error:
			return 255 === val;
		case c_oAscPivotRecType.Number:
			return c_nNumberMissingValue === val;
		case c_oAscPivotRecType.String:
			return 0 === val;
		case c_oAscPivotRecType.Index:
			return 0 === val;
	}
	return false;
};
PivotRecords.prototype._toXml = function(writer, elem) {
	switch (elem.type) {
		case c_oAscPivotRecType.Boolean:
			CT_Boolean.prototype.toXml2(writer, "b", elem.val, elem.addition);
			break;
		case c_oAscPivotRecType.DateTime:
			CT_DateTime.prototype.toXml2(writer, "d", elem.val, elem.addition);
			break;
		case c_oAscPivotRecType.Error:
			CT_Error.prototype.toXml2(writer, "e", elem.val, elem.addition);
			break;
		case c_oAscPivotRecType.Missing:
			CT_Missing.prototype.toXml2(writer, "m", elem.addition);
			break;
		case c_oAscPivotRecType.Number:
			CT_Number.prototype.toXml2(writer, "n", elem.val, elem.addition);
			break;
		case c_oAscPivotRecType.String:
			CT_String.prototype.toXml2(writer, "s", elem.val, elem.addition);
			break;
		case c_oAscPivotRecType.Index:
			CT_Index.prototype.toXml2(writer, "x", elem.val);
			break;
	}
};

var prot;

window['Asc']['c_oAscSourceType'] = window['AscCommonExcel'].c_oAscSourceType = c_oAscSourceType;
prot = c_oAscSourceType;
prot['Worksheet'] = prot.Worksheet;
prot['External'] = prot.External;
prot['Consolidation'] = prot.Consolidation;
prot['Scenario'] = prot.Scenario;

window['Asc']['c_oAscAxis'] = window['AscCommonExcel'].c_oAscAxis = c_oAscAxis;
prot = c_oAscAxis;
prot['AxisRow'] = prot.AxisRow;
prot['AxisCol'] = prot.AxisCol;
prot['AxisPage'] = prot.AxisPage;
prot['AxisValues'] = prot.AxisValues;

window['Asc']['c_oAscFieldSortType'] = window['AscCommonExcel'].c_oAscFieldSortType = c_oAscFieldSortType;
prot = c_oAscFieldSortType;
prot['Manual'] = prot.Manual;
prot['Ascending'] = prot.Ascending;
prot['Descending'] = prot.Descending;

window['Asc']['c_oAscDataConsolidateFunction'] = window['AscCommonExcel'].c_oAscDataConsolidateFunction = c_oAscDataConsolidateFunction;
prot = c_oAscDataConsolidateFunction;
prot['Average'] = prot.Average;
prot['Count'] = prot.Count;
prot['CountNums'] = prot.CountNums;
prot['Max'] = prot.Max;
prot['Min'] = prot.Min;
prot['Product'] = prot.Product;
prot['StdDev'] = prot.StdDev;
prot['StdDevp'] = prot.StdDevp;
prot['Sum'] = prot.Sum;
prot['Var'] = prot.Var;
prot['Varp'] = prot.Varp;

window['Asc']['c_oAscShowDataAs'] = window['AscCommonExcel'].c_oAscShowDataAs = c_oAscShowDataAs;
prot = c_oAscShowDataAs;
prot['Normal'] = prot.Normal;
prot['Difference'] = prot.Difference;
prot['Percent'] = prot.Percent;
prot['PercentDiff'] = prot.PercentDiff;
prot['RunTotal'] = prot.RunTotal;
prot['PercentOfRow'] = prot.PercentOfRow;
prot['PercentOfCol'] = prot.PercentOfCol;
prot['PercentOfTotal'] = prot.PercentOfTotal;
prot['Index'] = prot.Index;

window['Asc']['c_oAscFormatAction'] = window['AscCommonExcel'].c_oAscFormatAction = c_oAscFormatAction;
prot = c_oAscFormatAction;
prot['Blank'] = prot.Blank;
prot['Formatting'] = prot.Formatting;
prot['Drill'] = prot.Drill;
prot['Formula'] = prot.Formula;

window['Asc']['c_oAscScope'] = window['AscCommonExcel'].c_oAscScope = c_oAscScope;
prot = c_oAscScope;
prot['Selection'] = prot.Selection;
prot['Data'] = prot.Data;
prot['Field'] = prot.Field;

window['Asc']['c_oAscType'] = window['AscCommonExcel'].c_oAscType = c_oAscType;
prot = c_oAscType;
prot['None'] = prot.None;
prot['All'] = prot.All;
prot['Row'] = prot.Row;
prot['Column'] = prot.Column;

window['Asc']['c_oAscPivotFilterType'] = window['AscCommonExcel'].c_oAscPivotFilterType = c_oAscPivotFilterType;
prot = c_oAscPivotFilterType;
prot['Unknown'] = prot.Unknown;
prot['Count'] = prot.Count;
prot['Percent'] = prot.Percent;
prot['Sum'] = prot.Sum;
prot['CaptionEqual'] = prot.CaptionEqual;
prot['CaptionNotEqual'] = prot.CaptionNotEqual;
prot['CaptionBeginsWith'] = prot.CaptionBeginsWith;
prot['CaptionNotBeginsWith'] = prot.CaptionNotBeginsWith;
prot['CaptionEndsWith'] = prot.CaptionEndsWith;
prot['CaptionNotEndsWith'] = prot.CaptionNotEndsWith;
prot['CaptionContains'] = prot.CaptionContains;
prot['CaptionNotContains'] = prot.CaptionNotContains;
prot['CaptionGreaterThan'] = prot.CaptionGreaterThan;
prot['CaptionGreaterThanOrEqual'] = prot.CaptionGreaterThanOrEqual;
prot['CaptionLessThan'] = prot.CaptionLessThan;
prot['CaptionLessThanOrEqual'] = prot.CaptionLessThanOrEqual;
prot['CaptionBetween'] = prot.CaptionBetween;
prot['CaptionNotBetween'] = prot.CaptionNotBetween;
prot['ValueEqual'] = prot.ValueEqual;
prot['ValueNotEqual'] = prot.ValueNotEqual;
prot['ValueGreaterThan'] = prot.ValueGreaterThan;
prot['ValueGreaterThanOrEqual'] = prot.ValueGreaterThanOrEqual;
prot['ValueLessThan'] = prot.ValueLessThan;
prot['ValueLessThanOrEqual'] = prot.ValueLessThanOrEqual;
prot['ValueBetween'] = prot.ValueBetween;
prot['ValueNotBetween'] = prot.ValueNotBetween;
prot['DateEqual'] = prot.DateEqual;
prot['DateNotEqual'] = prot.DateNotEqual;
prot['DateOlderThan'] = prot.DateOlderThan;
prot['DateOlderThanOrEqual'] = prot.DateOlderThanOrEqual;
prot['DateNewerThan'] = prot.DateNewerThan;
prot['DateNewerThanOrEqual'] = prot.DateNewerThanOrEqual;
prot['DateBetween'] = prot.DateBetween;
prot['DateNotBetween'] = prot.DateNotBetween;
prot['Tomorrow'] = prot.Tomorrow;
prot['Today'] = prot.Today;
prot['Yesterday'] = prot.Yesterday;
prot['NextWeek'] = prot.NextWeek;
prot['ThisWeek'] = prot.ThisWeek;
prot['LastWeek'] = prot.LastWeek;
prot['NextMonth'] = prot.NextMonth;
prot['ThisMonth'] = prot.ThisMonth;
prot['LastMonth'] = prot.LastMonth;
prot['NextQuarter'] = prot.NextQuarter;
prot['ThisQuarter'] = prot.ThisQuarter;
prot['LastQuarter'] = prot.LastQuarter;
prot['NextYear'] = prot.NextYear;
prot['ThisYear'] = prot.ThisYear;
prot['LastYear'] = prot.LastYear;
prot['YearToDate'] = prot.YearToDate;
prot['Q1'] = prot.Q1;
prot['Q2'] = prot.Q2;
prot['Q3'] = prot.Q3;
prot['Q4'] = prot.Q4;
prot['M1'] = prot.M1;
prot['M2'] = prot.M2;
prot['M3'] = prot.M3;
prot['M4'] = prot.M4;
prot['M5'] = prot.M5;
prot['M6'] = prot.M6;
prot['M7'] = prot.M7;
prot['M8'] = prot.M8;
prot['M9'] = prot.M9;
prot['M10'] = prot.M10;
prot['M11'] = prot.M11;
prot['M12'] = prot.M12;

window['Asc']['c_oAscSortType'] = window['AscCommonExcel'].c_oAscSortType = c_oAscSortType;
prot = c_oAscSortType;
prot['None'] = prot.None;
prot['Ascending'] = prot.Ascending;
prot['Descending'] = prot.Descending;
prot['AscendingAlpha'] = prot.AscendingAlpha;
prot['DescendingAlpha'] = prot.DescendingAlpha;
prot['AscendingNatural'] = prot.AscendingNatural;
prot['DescendingNatural'] = prot.DescendingNatural;

window['Asc']['c_oAscPivotAreaType'] = window['AscCommonExcel'].c_oAscPivotAreaType = c_oAscPivotAreaType;
prot = c_oAscPivotAreaType;
prot['None'] = prot.None;
prot['Normal'] = prot.Normal;
prot['Data'] = prot.Data;
prot['All'] = prot.All;
prot['Origin'] = prot.Origin;
prot['Button'] = prot.Button;
prot['TopEnd'] = prot.TopEnd;

window['Asc']['c_oAscGroupBy'] = window['AscCommonExcel'].c_oAscGroupBy = c_oAscGroupBy;
prot = c_oAscGroupBy;
prot['Range'] = prot.Range;
prot['Seconds'] = prot.Seconds;
prot['Minutes'] = prot.Minutes;
prot['Hours'] = prot.Hours;
prot['Days'] = prot.Days;
prot['Months'] = prot.Months;
prot['Quarters'] = prot.Quarters;
prot['Years'] = prot.Years;

window['Asc']['c_oAscSortMethod'] = window['AscCommonExcel'].c_oAscSortMethod = c_oAscSortMethod;
prot = c_oAscSortMethod;
prot['Stroke'] = prot.Stroke;
prot['PinYin'] = prot.PinYin;
prot['None'] = prot.None;

window['Asc']['c_oAscDynamicFilterType'] = window['AscCommonExcel'].c_oAscDynamicFilterType = c_oAscDynamicFilterType;
prot = c_oAscDynamicFilterType;
prot['Null'] = prot.Null;
prot['AboveAverage'] = prot.AboveAverage;
prot['BelowAverage'] = prot.BelowAverage;
prot['Tomorrow'] = prot.Tomorrow;
prot['Today'] = prot.Today;
prot['Yesterday'] = prot.Yesterday;
prot['NextWeek'] = prot.NextWeek;
prot['ThisWeek'] = prot.ThisWeek;
prot['LastWeek'] = prot.LastWeek;
prot['NextMonth'] = prot.NextMonth;
prot['ThisMonth'] = prot.ThisMonth;
prot['LastMonth'] = prot.LastMonth;
prot['NextQuarter'] = prot.NextQuarter;
prot['ThisQuarter'] = prot.ThisQuarter;
prot['LastQuarter'] = prot.LastQuarter;
prot['NextYear'] = prot.NextYear;
prot['ThisYear'] = prot.ThisYear;
prot['LastYear'] = prot.LastYear;
prot['YearToDate'] = prot.YearToDate;
prot['Q1'] = prot.Q1;
prot['Q2'] = prot.Q2;
prot['Q3'] = prot.Q3;
prot['Q4'] = prot.Q4;
prot['M1'] = prot.M1;
prot['M2'] = prot.M2;
prot['M3'] = prot.M3;
prot['M4'] = prot.M4;
prot['M5'] = prot.M5;
prot['M6'] = prot.M6;
prot['M7'] = prot.M7;
prot['M8'] = prot.M8;
prot['M9'] = prot.M9;
prot['M10'] = prot.M10;
prot['M11'] = prot.M11;
prot['M12'] = prot.M12;

window['Asc']['c_oAscCalendarType'] = window['AscCommonExcel'].c_oAscCalendarType = c_oAscCalendarType;
prot = c_oAscCalendarType;
prot['Gregorian'] = prot.Gregorian;
prot['GregorianUs'] = prot.GregorianUs;
prot['GregorianMeFrench'] = prot.GregorianMeFrench;
prot['GregorianArabic'] = prot.GregorianArabic;
prot['Hijri'] = prot.Hijri;
prot['Hebrew'] = prot.Hebrew;
prot['Taiwan'] = prot.Taiwan;
prot['Japan'] = prot.Japan;
prot['Thai'] = prot.Thai;
prot['Korea'] = prot.Korea;
prot['Saka'] = prot.Saka;
prot['GregorianXlitEnglish'] = prot.GregorianXlitEnglish;
prot['GregorianXlitFrench'] = prot.GregorianXlitFrench;
prot['None'] = prot.None;

window['Asc']['c_oAscIconSetType'] = window['AscCommonExcel'].c_oAscIconSetType = c_oAscIconSetType;
prot = c_oAscIconSetType;
prot['ThreeArrows'] = prot.ThreeArrows;
prot['ThreeArrowsGray'] = prot.ThreeArrowsGray;
prot['ThreeFlags'] = prot.ThreeFlags;
prot['ThreeTrafficLights1'] = prot.ThreeTrafficLights1;
prot['ThreeTrafficLights2'] = prot.ThreeTrafficLights2;
prot['ThreeSigns'] = prot.ThreeSigns;
prot['ThreeSymbols'] = prot.ThreeSymbols;
prot['ThreeSymbols2'] = prot.ThreeSymbols2;
prot['FourArrows'] = prot.FourArrows;
prot['FourArrowsGray'] = prot.FourArrowsGray;
prot['FourRedToBlack'] = prot.FourRedToBlack;
prot['FourRating'] = prot.FourRating;
prot['FourTrafficLights'] = prot.FourTrafficLights;
prot['FiveArrows'] = prot.FiveArrows;
prot['FiveArrowsGray'] = prot.FiveArrowsGray;
prot['FiveRating'] = prot.FiveRating;
prot['FiveQuarters'] = prot.FiveQuarters;

window['Asc']['c_oAscSortBy'] = window['AscCommonExcel'].c_oAscSortBy = c_oAscSortBy;
prot = c_oAscSortBy;
prot['Value'] = prot.Value;
prot['CellColor'] = prot.CellColor;
prot['FontColor'] = prot.FontColor;
prot['Icon'] = prot.Icon;

window['Asc']['c_oAscFilterOperator'] = window['AscCommonExcel'].c_oAscFilterOperator = c_oAscFilterOperator;
prot = c_oAscFilterOperator;
prot['Equal'] = prot.Equal;
prot['LessThan'] = prot.LessThan;
prot['LessThanOrEqual'] = prot.LessThanOrEqual;
prot['NotEqual'] = prot.NotEqual;
prot['GreaterThanOrEqual'] = prot.GreaterThanOrEqual;
prot['GreaterThan'] = prot.GreaterThan;

window['Asc']['c_oAscDateTimeGrouping'] = window['AscCommonExcel'].c_oAscDateTimeGrouping = c_oAscDateTimeGrouping;
prot = c_oAscDateTimeGrouping;
prot['Year'] = prot.Year;
prot['Month'] = prot.Month;
prot['Day'] = prot.Day;
prot['Hour'] = prot.Hour;
prot['Minute'] = prot.Minute;
prot['Second'] = prot.Second;

window['Asc']['st_VALUES'] = window['AscCommonExcel'].st_VALUES = st_VALUES;
window['AscCommonExcel'].DATA_CAPTION = DATA_CAPTION;
window['AscCommonExcel'].GRAND_TOTAL_CAPTION = GRAND_TOTAL_CAPTION;
window['AscCommonExcel'].ROW_HEADER_CAPTION = ROW_HEADER_CAPTION;
window['AscCommonExcel'].COL_HEADER_CAPTION = COL_HEADER_CAPTION;
window['AscCommonExcel'].PAGE_ALL_CAPTION = PAGE_ALL_CAPTION;
window['AscCommonExcel'].NEW_PIVOT_LAST_COL_OFFSET = NEW_PIVOT_LAST_COL_OFFSET;
window['AscCommonExcel'].NEW_PIVOT_LAST_ROW_OFFSET = NEW_PIVOT_LAST_ROW_OFFSET;
window['AscCommonExcel'].NEW_PIVOT_ROW = NEW_PIVOT_ROW;
window['AscCommonExcel'].NEW_PIVOT_COL = NEW_PIVOT_COL;

window['AscCommonExcel'].ToName_ST_ItemType = ToName_ST_ItemType;
window['AscCommonExcel'].ToName_ST_DataConsolidateFunction = ToName_ST_DataConsolidateFunction;

window['Asc']['CT_PivotCacheDefinition'] = window['Asc'].CT_PivotCacheDefinition = CT_PivotCacheDefinition;

window['Asc']['CT_PivotCacheRecords'] = window['Asc'].CT_PivotCacheRecords = CT_PivotCacheRecords;

window["Asc"]["CT_pivotTableDefinition"] = window['Asc'].CT_pivotTableDefinition = CT_pivotTableDefinition;
prot = CT_pivotTableDefinition.prototype;
prot["asc_getName"] = prot.asc_getName;
prot["asc_getPageWrap"] = prot.asc_getPageWrap;
prot["asc_getPageOverThenDown"] = prot.asc_getPageOverThenDown;
prot["asc_getRowGrandTotals"] = prot.asc_getRowGrandTotals;
prot["asc_getColGrandTotals"] = prot.asc_getColGrandTotals;
prot["asc_getShowHeaders"] = prot.asc_getShowHeaders;
prot["asc_getDataRef"] = prot.asc_getDataRef;
prot["asc_getTitle"] = prot.asc_getTitle;
prot["asc_getDescription"] = prot.asc_getDescription;
prot["asc_getStyleInfo"] = prot.asc_getStyleInfo;
prot["asc_getCacheFields"] = prot.asc_getCacheFields;
prot["asc_getPivotFields"] = prot.asc_getPivotFields;
prot["asc_getPageFields"] = prot.asc_getPageFields;
prot["asc_getColumnFields"] = prot.asc_getColumnFields;
prot["asc_getRowFields"] = prot.asc_getRowFields;
prot["asc_getDataFields"] = prot.asc_getDataFields;
prot["asc_select"] = prot.asc_select;
prot["asc_set"] = prot.asc_set;
prot["asc_setName"] = prot.asc_setName;
prot["asc_setRowGrandTotals"] = prot.asc_setRowGrandTotals;
prot["asc_setColGrandTotals"] = prot.asc_setColGrandTotals;
prot["asc_setPageOverThenDown"] = prot.asc_setPageOverThenDown;
prot["asc_setPageWrap"] = prot.asc_setPageWrap;
prot["asc_setShowHeaders"] = prot.asc_setShowHeaders;
prot["asc_setCompact"] = prot.asc_setCompact;
prot["asc_setOutline"] = prot.asc_setOutline;
prot["asc_setFillDownLabelsDefault"] = prot.asc_setFillDownLabelsDefault;
prot["asc_setDataRef"] = prot.asc_setDataRef;
prot["asc_setTitle"] = prot.asc_setTitle;
prot["asc_setDescription"] = prot.asc_setDescription;
prot["asc_setInsertBlankRow"] = prot.asc_setInsertBlankRow;
prot["asc_setDefaultSubtotal"] = prot.asc_setDefaultSubtotal;
prot["asc_setSubtotalTop"] = prot.asc_setSubtotalTop;
prot["asc_addPageField"] = prot.asc_addPageField;
prot["asc_addRowField"] = prot.asc_addRowField;
prot["asc_addColField"] = prot.asc_addColField;
prot["asc_addDataField"] = prot.asc_addDataField;
prot["asc_addField"] = prot.asc_addField;
prot["asc_removeField"] = prot.asc_removeField;
prot["asc_removeNoDataField"] = prot.asc_removeNoDataField;
prot["asc_removeDataField"] = prot.asc_removeDataField;
prot["asc_moveToPageField"] = prot.asc_moveToPageField;
prot["asc_moveToRowField"] = prot.asc_moveToRowField;
prot["asc_moveToColField"] = prot.asc_moveToColField;
prot["asc_moveToDataField"] = prot.asc_moveToDataField;
prot["asc_movePageField"] = prot.asc_movePageField;
prot["asc_moveRowField"] = prot.asc_moveRowField;
prot["asc_moveColField"] = prot.asc_moveColField;
prot["asc_moveDataField"] = prot.asc_moveDataField;
prot["asc_refresh"] = prot.asc_refresh;

prot = CT_PivotTableStyle.prototype;
prot["asc_getName"] = prot.asc_getName;
prot["asc_getShowRowHeaders"] = prot.asc_getShowRowHeaders;
prot["asc_getShowColHeaders"] = prot.asc_getShowColHeaders;
prot["asc_getShowRowStripes"] = prot.asc_getShowRowStripes;
prot["asc_getShowColStripes"] = prot.asc_getShowColStripes;
prot["asc_setName"] = prot.asc_setName;
prot["asc_setShowRowHeaders"] = prot.asc_setShowRowHeaders;
prot["asc_setShowColHeaders"] = prot.asc_setShowColHeaders;
prot["asc_setShowRowStripes"] = prot.asc_setShowRowStripes;
prot["asc_setShowColStripes"] = prot.asc_setShowColStripes;

prot = CT_CacheField.prototype;
prot["asc_getName"] = prot.asc_getName;

window["Asc"]["CT_PivotField"] = CT_PivotField;
prot = CT_PivotField.prototype;
prot["asc_getName"] = prot.asc_getName;
prot["asc_getOutline"] = prot.asc_getOutline;
prot["asc_getCompact"] = prot.asc_getCompact;
prot["asc_getFillDownLabelsDefault"] = prot.asc_getFillDownLabelsDefault;
prot["asc_getInsertBlankRow"] = prot.asc_getInsertBlankRow;
prot["asc_getDefaultSubtotal"] = prot.asc_getDefaultSubtotal;
prot["asc_getSubtotalTop"] = prot.asc_getSubtotalTop;
prot["asc_getShowAll"] = prot.asc_getShowAll;
prot["asc_getSubtotals"] = prot.asc_getSubtotals;
prot["asc_set"] = prot.asc_set;
prot["asc_setName"] = prot.asc_setName;
prot["asc_setOutline"] = prot.asc_setOutline;
prot["asc_setCompact"] = prot.asc_setCompact;
prot["asc_setInsertBlankRow"] = prot.asc_setInsertBlankRow;
prot["asc_setDefaultSubtotal"] = prot.asc_setDefaultSubtotal;
prot["asc_setSubtotalTop"] = prot.asc_setSubtotalTop;
prot["asc_setShowAll"] = prot.asc_setShowAll;
prot["asc_setSubtotals"] = prot.asc_setSubtotals;

prot = CT_Field.prototype;
prot["asc_getIndex"] = prot.asc_getIndex;

prot = CT_PageField.prototype;
prot["asc_getName"] = prot.asc_getName;
prot["asc_getIndex"] = prot.asc_getIndex;

window["Asc"]["CT_DataField"] = CT_DataField;
prot = CT_DataField.prototype;
prot["asc_getName"] = prot.asc_getName;
prot["asc_getIndex"] = prot.asc_getIndex;
prot["asc_getSubtotal"] = prot.asc_getSubtotal;
prot["asc_getShowDataAs"] = prot.asc_getShowDataAs;
prot["asc_set"] = prot.asc_set;
prot["asc_setName"] = prot.asc_setName;
prot["asc_setSubtotal"] = prot.asc_setSubtotal;

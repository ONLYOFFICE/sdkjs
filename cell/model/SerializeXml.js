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

/**
 * @param {Window} window
 * @param {undefined} undefined
 */
(function (window, undefined) {
	var CellValueType = AscCommon.CellValueType;

	//convert const functions
	function FromXml_ST_IconSetType(val) {
		//в пивотах есть функция FromXml_ST_IconSetType, но там корвенртирцем в другие константы. пока оставляю так, нужно сделать общие
		var res = undefined;
		switch (val) {
			case "3Arrows":
				res = Asc.EIconSetType.Arrows3;
				break;
			case "3ArrowsGray":
				res = Asc.EIconSetType.Arrows3Gray;
				break;
			case "3Flags":
				res = Asc.EIconSetType.Flags3;
				break;
			case "3TrafficLights1":
				res = Asc.EIconSetType.Traffic3Lights1;
				break;
			case "3TrafficLights2":
				res = Asc.EIconSetType.Traffic3Lights2;
				break;
			case "3Signs":
				res = Asc.EIconSetType.Signs3;
				break;
			case "3Symbols":
				res = Asc.EIconSetType.Symbols3;
				break;
			case "3Symbols2":
				res = Asc.EIconSetType.Symbols3_2;
				break;
			case "4Arrows":
				res = Asc.EIconSetType.Arrows4;
				break;
			case "4ArrowsGray":
				res = Asc.EIconSetType.Arrows4Gray;
				break;
			case "4TrafficLights":
				res = Asc.EIconSetType.Traffic4Lights;
				break;
			case "5Arrows":
				res = Asc.EIconSetType.Arrows5;
				break;
			case "5ArrowsGray":
				res = Asc.EIconSetType.Arrows5Gray;
				break;
			case "5Rating":
				res = Asc.EIconSetType.Rating5;
				break;
			case "5Quarters":
				res = Asc.EIconSetType.Quarters5;
				break;
			case "4RedToBlack":
				res = Asc.EIconSetType.RedToBlack4;
				break;
			case "4Rating":
				res = Asc.EIconSetType.Rating4;
				break;
			case "3Stars":
				res = Asc.EIconSetType.Stars3;
				break;
			case "3Triangles":
				res = Asc.EIconSetType.Triangles3;
				break;
			case "5Boxes":
				res = Asc.EIconSetType.Boxes5;
				break;
			case "NoIcons":
				res = Asc.EIconSetType.NoIcons;
				break;
		}
		return res;
	}

	function ToXml_ST_IconSetType(val) {
		//в пивотах есть функция ToXml_ST_IconSetType, но там корвенртирцем в другие константы. пока оставляю так, нужно сделать общие
		var res = undefined;
		switch (val) {
			case Asc.EIconSetType.Arrows3:
				res = "3Arrows";
				break;
			case Asc.EIconSetType.Arrows3Gray:
				res = "3ArrowsGray";
				break;
			case Asc.EIconSetType.Flags3:
				res = "3Flags";
				break;
			case Asc.EIconSetType.Traffic3Lights1:
				res = "3TrafficLights1";
				break;
			case Asc.EIconSetType.Traffic3Lights2:
				res = "3TrafficLights2";
				break;
			case Asc.EIconSetType.Signs3:
				res = "3Signs";
				break;
			case Asc.EIconSetType.Symbols3:
				res = "3Symbols";
				break;
			case Asc.EIconSetType.Symbols3_2:
				res = "3Symbols2";
				break;
			case Asc.EIconSetType.Arrows4:
				res = "4Arrows";
				break;
			case Asc.EIconSetType.Arrows4Gray:
				res = "4ArrowsGray";
				break;
			case Asc.EIconSetType.Traffic4Lights:
				res = "4TrafficLights";
				break;
			case Asc.EIconSetType.Arrows5:
				res = "5Arrows";
				break;
			case Asc.EIconSetType.Arrows5Gray:
				res = "5ArrowsGray";
				break;
			case Asc.EIconSetType.Rating5:
				res = "5Rating";
				break;
			case Asc.EIconSetType.Quarters5:
				res = "5Quarters";
				break;
			case Asc.EIconSetType.RedToBlack4:
				res = "4RedToBlack";
				break;
			case Asc.EIconSetType.Rating4:
				res = "4Rating";
				break;
			case Asc.EIconSetType.Stars3:
				res = "3Stars";
				break;
			case Asc.EIconSetType.Triangles3:
				res = "3Triangles";
				break;
			case Asc.EIconSetType.Boxes5:
				res = "5Boxes";
				break;
			case Asc.EIconSetType.NoIcons:
				res = "NoIcons";
				break;
		}
		return res;
	}

	function FromXml_ST_FontScheme(val) {
		var res = null;
		switch (val) {
			case  "major":
				res = Asc.EFontScheme.fontschemeMajor;
				break;
			case  "minor":
				res = Asc.EFontScheme.fontschemeMinor;
				break;
			case  Asc.EFontScheme.fontschemeNone:
				break;
		}
		return res;
	}

	function ToXml_ST_FontScheme(val) {
		var res = null;
		switch (val) {
			case  Asc.EFontScheme.fontschemeMajor:
				res = "major";
				break;
			case  Asc.EFontScheme.fontschemeMinor:
				res = "minor";
				break;
			case  Asc.EFontScheme.fontschemeNone:
				break;
		}
		return res;
	}

	function FromXml_ST_PageOrder(val) {
		var res = null;
		switch (val) {
			case  "downThenOver":
				res = 0;
				break;
			case  "overThenDown":
				res = 1;
				break;
		}
		return res;
	}

	function ToXml_ST_PageOrder(val) {
		var res = null;
		switch (val) {
			case  0:
				res = "downThenOver";
				break;
			case  1:
				res = "overThenDown";
				break;
		}
		return res;
	}

	function FromXml_ST_PageOrientation(val) {
		var res = null;
		switch (val) {
			case "portrait":
				res = Asc.c_oAscPageOrientation.PagePortrait;
				break;
			case "landscape":
				res = Asc.c_oAscPageOrientation.PageLandscape;
				break;
		}
		return res;
	}

	function ToXml_ST_PageOrientation(val) {
		var res = null;
		switch (val) {
			case Asc.c_oAscPageOrientation.PagePortrait:
				res = "portrait";
				break;
			case Asc.c_oAscPageOrientation.PageLandscape:
				res = "landscape";
				break;
		}
		return res;
	}

	function FromXml_ST_CellComments(val) {
		var res = 0;
		switch (val) {
			case "none":
				res = AscCommonExcel.ST_CellComments.none;
				break;
			case "asDisplayed":
				res = AscCommonExcel.ST_CellComments.asDisplayed;
				break;
			case "atEnd":
				res = AscCommonExcel.ST_CellComments.atEnd;
				break;
		}
		return res;
	}

	function ToXml_ST_CellComments(val) {
		var res = null;
		switch (val) {
			case AscCommonExcel.ST_CellComments.none:
				res = "none";
				break;
			case AscCommonExcel.ST_CellComments.asDisplayed:
				res = "asDisplayed";
				break;
			case AscCommonExcel.ST_CellComments.atEnd:
				res = "atEnd";
				break;
		}
		return res;
	}

	function FromXml_ST_PrintError(val) {
		var res = 0;
		switch (val) {
			case "displayed":
				res = AscCommonExcel.ST_PrintError.displayed;
				break;
			case "blank":
				res = AscCommonExcel.ST_PrintError.blank;
				break;
			case "dash":
				res = AscCommonExcel.ST_PrintError.dash;
				break;
			case "NA":
				res = AscCommonExcel.ST_PrintError.NA;
				break;
		}
		return res;
	}

	function ToXml_ST_PrintError(val) {
		var res = null;
		switch (val) {
			case AscCommonExcel.ST_PrintError.displayed:
				res = "displayed";
				break;
			case AscCommonExcel.ST_PrintError.blank:
				res = "blank";
				break;
			case AscCommonExcel.ST_PrintError.dash:
				res = "dash";
				break;
			case AscCommonExcel.ST_PrintError.NA:
				res = "NA";
				break;
		}
		return res;
	}

	function FromXml_ST_SortBy(val) {
		var res = 0;
		switch (val) {
			case "value":
				res = Asc.ESortBy.sortbyValue;
				break;
			case "cellColor":
				res = Asc.ESortBy.sortbyCellColor;
				break;
			case "fontColor":
				res = Asc.ESortBy.sortbyFontColor;
				break;
			case "icon":
				res = Asc.ESortBy.sortbyIcon;
				break;
		}
		return res;
	}

	function ToXml_ST_SortBy(val) {
		var res = null;
		switch (val) {
			case Asc.ESortBy.sortbyValue:
				res = "value";
				break;
			case Asc.ESortBy.sortbyCellColor:
				res = "cellColor";
				break;
			case Asc.ESortBy.sortbyFontColor:
				res = "fontColor";
				break;
			case Asc.ESortBy.sortbyIcon:
				res = "icon";
				break;
		}
		return res;
	}

	function FromXml_ST_CfType(val) {
		var res = null;
		switch (val) {
			case "aboveAverage":
				res = Asc.ECfType.aboveAverage;
				break;
			case "beginsWith":
				res = Asc.ECfType.beginsWith;
				break;
			case "cellIs":
				res = Asc.ECfType.cellIs;
				break;
			case "colorScale":
				res = Asc.ECfType.colorScale;
				break;
			case "containsBlanks":
				res = Asc.ECfType.containsBlanks;
				break;
			case "containsErrors":
				res = Asc.ECfType.containsErrors;
				break;
			case "containsText":
				res = Asc.ECfType.containsText;
				break;
			case "dataBar":
				res = Asc.ECfType.dataBar;
				break;
			case "duplicateValues":
				res = Asc.ECfType.duplicateValues;
				break;
			case "expression":
				res = Asc.ECfType.expression;
				break;
			case "notContainsBlanks":
				res = Asc.ECfType.notContainsBlanks;
				break;
			case "notContainsErrors":
				res = Asc.ECfType.notContainsErrors;
				break;
			case "notContainsText":
				res = Asc.ECfType.notContainsText;
				break;
			case "timePeriod":
				res = Asc.ECfType.timePeriod;
				break;
			case "top10":
				res = Asc.ECfType.top10;
				break;
			case "uniqueValues":
				res = Asc.ECfType.uniqueValues;
				break;
			case "endsWith":
				res = Asc.ECfType.endsWith;
				break;
			case "iconSet":
				res = Asc.ECfType.iconSet;
				break;
		}
		return res;
	}

	function ToXml_ST_CfType(val) {
		var res = null;
		switch (val) {
			case Asc.ECfType.aboveAverage:
				res = "aboveAverage";
				break;
			case Asc.ECfType.beginsWith:
				res = "beginsWith";
				break;
			case Asc.ECfType.cellIs:
				res = "cellIs";
				break;
			case Asc.ECfType.colorScale:
				res = "colorScale";
				break;
			case Asc.ECfType.containsBlanks:
				res = "containsBlanks";
				break;
			case Asc.ECfType.containsErrors:
				res = "containsErrors";
				break;
			case Asc.ECfType.containsText:
				res = "containsText";
				break;
			case Asc.ECfType.dataBar:
				res = "dataBar";
				break;
			case Asc.ECfType.duplicateValues:
				res = "duplicateValues";
				break;
			case Asc.ECfType.expression:
				res = "expression";
				break;
			case Asc.ECfType.notContainsBlanks:
				res = "notContainsBlanks";
				break;
			case Asc.ECfType.notContainsErrors:
				res = "notContainsErrors";
				break;
			case Asc.ECfType.notContainsText:
				res = "notContainsText";
				break;
			case Asc.ECfType.timePeriod:
				res = "timePeriod";
				break;
			case Asc.ECfType.top10:
				res = "top10";
				break;
			case Asc.ECfType.uniqueValues:
				res = "uniqueValues";
				break;
			case Asc.ECfType.endsWith:
				res = "endsWith";
				break;
			case Asc.ECfType.iconSet:
				res = "iconSet";
				break;
		}
		return res;
	}

	function FromXml_ST_DataValidationType(val) {
		var res = undefined;
		switch (val) {
			case "none":
				res = Asc.EDataValidationType.None;
				break;
			case "whole":
				res = Asc.EDataValidationType.Whole;
				break;
			case "decimal":
				res = Asc.EDataValidationType.Decimal;
				break;
			case "list":
				res = Asc.EDataValidationType.List;
				break;
			case "date":
				res = Asc.EDataValidationType.Date;
				break;
			case "time":
				res = Asc.EDataValidationType.Time;
				break;
			case "textLength":
				res = Asc.EDataValidationType.TextLength;
				break;
			case "custom":
				res = Asc.EDataValidationType.Custom;
				break;
		}
		return res;
	}

	function ToXml_ST_DataValidationType(val) {
		var res = undefined;
		switch (val) {
			case Asc.EDataValidationType.None:
				res = "none";
				break;
			case Asc.EDataValidationType.Whole:
				res = "whole";
				break;
			case Asc.EDataValidationType.Decimal:
				res = "decimal";
				break;
			case Asc.EDataValidationType.List:
				res = "list";
				break;
			case Asc.EDataValidationType.Date:
				res = "date";
				break;
			case Asc.EDataValidationType.Time:
				res = "time";
				break;
			case Asc.EDataValidationType.TextLength:
				res = "textLength";
				break;
			case Asc.EDataValidationType.Custom:
				res = "custom";
				break;
		}
		return res;
	}

	function FromXml_ST_DataValidationImeMode(val) {
		var res = undefined;
		switch (val) {
			case "noControl":
				res = Asc.EDataValidationImeMode.NoControl;
				break;
			case "off":
				res = Asc.EDataValidationImeMode.Off;
				break;
			case "on":
				res = Asc.EDataValidationImeMode.On;
				break;
			case "disabled":
				res = Asc.EDataValidationImeMode.Disabled;
				break;
			case "hiragana":
				res = Asc.EDataValidationImeMode.Hiragana;
				break;
			case "fullKatakana":
				res = Asc.EDataValidationImeMode.FullKatakana;
				break;
			case "halfKatakana":
				res = Asc.EDataValidationImeMode.HalfKatakana;
				break;
			case "fullAlpha":
				res = Asc.EDataValidationImeMode.FullAlpha;
				break;
			case "halfAlpha":
				res = Asc.EDataValidationImeMode.HalfAlpha;
				break;
			case "fullHangul":
				res = Asc.EDataValidationImeMode.FullHangul;
				break;
			case "halfHangul":
				res = Asc.EDataValidationImeMode.HalfHangul;
				break;
		}
		return res;
	}

	function ToXml_ST_DataValidationImeMode(val) {
		var res = undefined;
		switch (val) {
			case Asc.EDataValidationImeMode.NoControl:
				res = "noControl";
				break;
			case Asc.EDataValidationImeMode.Off:
				res = "off";
				break;
			case Asc.EDataValidationImeMode.On:
				res = "on";
				break;
			case Asc.EDataValidationImeMode.Disabled:
				res = "disabled";
				break;
			case Asc.EDataValidationImeMode.Hiragana:
				res = "hiragana";
				break;
			case Asc.EDataValidationImeMode.FullKatakana:
				res = "fullKatakana";
				break;
			case Asc.EDataValidationImeMode.HalfKatakana:
				res = "halfKatakana";
				break;
			case Asc.EDataValidationImeMode.FullAlpha:
				res = "fullAlpha";
				break;
			case Asc.EDataValidationImeMode.HalfAlpha:
				res = "halfAlpha";
				break;
			case Asc.EDataValidationImeMode.FullHangul:
				res = "fullHangul";
				break;
			case Asc.EDataValidationImeMode.HalfHangul:
				res = "halfHangul";
				break;
		}
		return res;
	}

	function FromXml_ST_DataValidationOperator(val) {
		var res = null;
		switch (val) {
			case "between":
				res = Asc.EDataValidationOperator.Between;
				break;
			case "notBetween":
				res = Asc.EDataValidationOperator.NotBetween;
				break;
			case "equal":
				res = Asc.EDataValidationOperator.Equal;
				break;
			case "notEqual":
				res = Asc.EDataValidationOperator.NotEqual;
				break;
			case "lessThan":
				res = Asc.EDataValidationOperator.LessThan;
				break;
			case "lessThanOrEqual":
				res = Asc.EDataValidationOperator.LessThanOrEqual;
				break;
			case "greaterThan":
				res = Asc.EDataValidationOperator.GreaterThan;
				break;
			case "greaterThanOrEq":
				res = Asc.EDataValidationOperator.GreaterThanOrEqual;
		}
		return res;
	}

	function ToXml_ST_DataValidationOperator(val) {
		var res = null;
		switch (val) {
			case Asc.EDataValidationOperator.Between:
				res = "between";
				break;
			case Asc.EDataValidationOperator.NotBetween:
				res = "notBetween";
				break;
			case Asc.EDataValidationOperator.Equal:
				res = "equal";
				break;
			case Asc.EDataValidationOperator.NotEqual:
				res = "notEqual";
				break;
			case Asc.EDataValidationOperator.LessThan:
				res = "lessThan";
				break;
			case Asc.EDataValidationOperator.LessThanOrEqual:
				res = "lessThanOrEqual";
				break;
			case Asc.EDataValidationOperator.GreaterThan:
				res = "greaterThan";
				break;
			case Asc.EDataValidationOperator.GreaterThanOrEqual:
				res = "greaterThanOrEq";
		}
		return res;
	}

	function ToXml_CFOperatorType(nType) {
		var sType = undefined;
		switch (nType) {
			case AscCommonExcel.ECfOperator.Operator_beginsWith:
				sType = "beginsWith";
				break;
			case AscCommonExcel.ECfOperator.Operator_between:
				sType = "between";
				break;
			case AscCommonExcel.ECfOperator.Operator_containsText:
				sType = "containsText";
				break;
			case AscCommonExcel.ECfOperator.Operator_endsWith:
				sType = "endsWith";
				break;
			case AscCommonExcel.ECfOperator.Operator_equal:
				sType = "equal";
				break;
			case AscCommonExcel.ECfOperator.Operator_greaterThan:
				sType = "greaterThan";
				break;
			case AscCommonExcel.ECfOperator.Operator_greaterThanOrEqual:
				sType = "greaterThanOrEqual";
				break;
			case AscCommonExcel.ECfOperator.Operator_lessThan:
				sType = "lessThan";
				break;
			case AscCommonExcel.ECfOperator.Operator_lessThanOrEqual:
				sType = "lessThanOrEqual";
				break;
			case AscCommonExcel.ECfOperator.Operator_notBetween:
				sType = "notBetween";
				break;
			case AscCommonExcel.ECfOperator.Operator_notContains:
				sType = "notContains";
				break;
			case AscCommonExcel.ECfOperator.Operator_notEqual:
				sType = "notEqual";
				break;
		}

		return sType;
	}

	function FromXml_CFOperatorType(sType) {
		var nType = undefined;
		switch (sType) {
			case "beginsWith":
				nType = AscCommonExcel.ECfOperator.Operator_beginsWith;
				break;
			case "between":
				nType = AscCommonExcel.ECfOperator.Operator_between;
				break;
			case "containsText":
				nType = AscCommonExcel.ECfOperator.Operator_containsText;
				break;
			case "endsWith":
				nType = AscCommonExcel.ECfOperator.Operator_endsWith;
				break;
			case "equal":
				nType = AscCommonExcel.ECfOperator.Operator_equal;
				break;
			case "greaterThan":
				nType = AscCommonExcel.ECfOperator.Operator_greaterThan;
				break;
			case "greaterThanOrEqual":
				nType = AscCommonExcel.ECfOperator.Operator_greaterThanOrEqual;
				break;
			case "lessThan":
				nType = AscCommonExcel.ECfOperator.Operator_lessThan;
				break;
			case "lessThanOrEqual":
				nType = AscCommonExcel.ECfOperator.Operator_lessThanOrEqual;
				break;
			case "notBetween":
				nType = AscCommonExcel.ECfOperator.Operator_notBetween;
				break;
			case "notContains":
				nType = AscCommonExcel.ECfOperator.Operator_notContains;
				break;
			case "notEqual":
				nType = AscCommonExcel.ECfOperator.Operator_notEqual;
				break;
		}

		return nType;
	}

	function ToXml_ST_TimePeriod(nType) {
		var sType = undefined;
		switch (nType) {
			case AscCommonExcel.ST_TimePeriod.last7Days:
				sType = "last7Days";
				break;
			case AscCommonExcel.ST_TimePeriod.lastMonth:
				sType = "lastMonth";
				break;
			case AscCommonExcel.ST_TimePeriod.lastWeek:
				sType = "lastWeek";
				break;
			case AscCommonExcel.ST_TimePeriod.nextMonth:
				sType = "nextMonth";
				break;
			case AscCommonExcel.ST_TimePeriod.nextWeek:
				sType = "nextWeek";
				break;
			case AscCommonExcel.ST_TimePeriod.thisMonth:
				sType = "thisMonth";
				break;
			case AscCommonExcel.ST_TimePeriod.thisWeek:
				sType = "thisWeek";
				break;
			case AscCommonExcel.ST_TimePeriod.today:
				sType = "today";
				break;
			case AscCommonExcel.ST_TimePeriod.tomorrow:
				sType = "tomorrow";
				break;
			case AscCommonExcel.ST_TimePeriod.yesterday:
				sType = "yesterday";
				break;
		}

		return sType;
	}

	function FromXml_ST_TimePeriod(sType) {
		var nType = undefined;
		switch (sType) {
			case "last7Days":
				nType = AscCommonExcel.ST_TimePeriod.last7Days;
				break;
			case "lastMonth":
				nType = AscCommonExcel.ST_TimePeriod.lastMonth;
				break;
			case "lastWeek":
				nType = AscCommonExcel.ST_TimePeriod.lastWeek;
				break;
			case "nextMonth":
				nType = AscCommonExcel.ST_TimePeriod.nextMonth;
				break;
			case "nextWeek":
				nType = AscCommonExcel.ST_TimePeriod.nextWeek;
				break;
			case "thisMonth":
				nType = AscCommonExcel.ST_TimePeriod.thisMonth;
				break;
			case "thisWeek":
				nType = AscCommonExcel.ST_TimePeriod.thisWeek;
				break;
			case "today":
				nType = AscCommonExcel.ST_TimePeriod.today;
				break;
			case "tomorrow":
				nType = AscCommonExcel.ST_TimePeriod.tomorrow;
				break;
			case "yesterday":
				nType = AscCommonExcel.ST_TimePeriod.yesterday;
				break;
		}

		return nType;
	}

	function FromXml_ST_TabularSlicerCacheSortOrder(val) {
		var res = null;
		switch (val) {
			case "ascending":
				res = Asc.ST_tabularSlicerCacheSortOrder.Ascending;
				break;
			case "descending":
				res = Asc.ST_tabularSlicerCacheSortOrder.Descending;
				break;
		}
		return res;
	}

	function ToXml_ST_TabularSlicerCacheSortOrder(val) {
		var res = null;
		switch (val) {
			case Asc.ST_tabularSlicerCacheSortOrder.Ascending:
				res = "ascending";
				break;
			case Asc.ST_tabularSlicerCacheSortOrder.Descending:
				res = "descending";
				break;
		}
		return res;
	}

	function FromXml_ST_SlicerCacheCrossFilter(val) {
		var res = null;
		switch (val) {
			case "showItemsWithDataAtTop":
				res = Asc.ST_slicerCacheCrossFilter.ShowItemsWithDataAtTop;
				break;
			case "showItemsWithNoData":
				res = Asc.ST_slicerCacheCrossFilter.ShowItemsWithNoData;
				break;
			default:
				res = Asc.ST_slicerCacheCrossFilter.None;
				break;
		}
		return res;
	}

	function ToXml_ST_SlicerCacheCrossFilter(val) {
		var res = null;
		switch (val) {
			case Asc.ST_slicerCacheCrossFilter.ShowItemsWithDataAtTop:
				res = "showItemsWithDataAtTop";
				break;
			case Asc.ST_slicerCacheCrossFilter.ShowItemsWithNoData:
				res = "showItemsWithNoData";
				break;
			default:
				res = "none";
				break;
		}
		return res;
	}

	function FromXml_ST_SortMethod(val) {
		var res = null;
		switch (val) {
			case "stroke":
				res = AscCommonExcel.ESortMethod.sortmethodStroke;
				break;
			case "pinYin":
				res = AscCommonExcel.ESortMethod.sortmethodPinYin;
				break;
		}
		return res;
	}

	function ToXml_ST_SortMethod(val) {
		var res = null;
		switch (val) {
			case AscCommonExcel.ESortMethod.sortmethodStroke:
				res = "stroke";
				break;
			case AscCommonExcel.ESortMethod.sortmethodPinYin:
				res = "pinYin";
				break;
		}
		return res;
	}

	function FromXML_ST_DispBlanksAs(val) {
		var res = null;
		switch (val) {
			case "span":
				res = Asc.c_oAscEDispBlanksAs.Span;
				break;
			case "gap":
				res = Asc.c_oAscEDispBlanksAs.Gap;
				break;
			case "zero":
				res = Asc.c_oAscEDispBlanksAs.Zero;
				break;
		}
		return res;
	}

	function ToXML_ST_DispBlanksAs(val) {
		var res = null;
		switch (val) {
			case Asc.c_oAscEDispBlanksAs.Span:
				res = "span";
				break;
			case Asc.c_oAscEDispBlanksAs.Gap:
				res = "gap";
				break;
			case Asc.c_oAscEDispBlanksAs.Zero:
				res = "zero";
				break;
		}
		return res;
	}

	function FromXml_ST_SparklineType(val) {
		var res = null;
		switch (val) {
			case "line":
				res = Asc.c_oAscSparklineType.Line;
				break;
			case "column":
				res = Asc.c_oAscSparklineType.Column;
				break;
			case "stacked":
				res = Asc.c_oAscSparklineType.Stacked;
				break;
		}
		return res;
	}

	function ToXml_ST_SparklineType(val) {
		var res = null;
		switch (val) {
			case Asc.c_oAscSparklineType.Line:
				res = "line";
				break;
			case Asc.c_oAscSparklineType.Column:
				res = "column";
				break;
			case Asc.c_oAscSparklineType.Stacked:
				res = "stacked";
				break;
		}
		return res;
	}

	function FromXml_ST_UnderlineValues(val) {
		var res = null;
		switch (val) {
			case "single":
				res = Asc.EUnderline.underlineSingle;
				break;
			case "double":
				res = Asc.EUnderline.underlineDouble;
				break;
			case "singleAccounting":
				res = Asc.EUnderline.underlineSingleAccounting;
				break;
			case "doubleAccounting":
				res = Asc.EUnderline.underlineDoubleAccounting;
				break;
			case "none":
				res = Asc.EUnderline.underlineNone;
				break;
		}
		return res;
	}

	function ToXml_ST_UnderlineValues(val) {
		var res = null;
		switch (val) {
			case Asc.EUnderline.underlineSingle:
				res = "single";
				break;
			case Asc.EUnderline.underlineDouble:
				res = "double";
				break;
			case Asc.EUnderline.underlineSingleAccounting:
				res = "singleAccounting";
				break;
			case Asc.EUnderline.underlineDoubleAccounting:
				res = "doubleAccounting";
				break;
			case Asc.EUnderline.underlineNone:
				res = "none";
				break;
		}
		return res;
	}

	function FromXml_ST_SparklineAxisMinMax(val) {
		var res = null;
		switch (val) {
			case "individual":
				res = Asc.c_oAscSparklineAxisMinMax.Individual;
				break;
			case "group":
				res = Asc.c_oAscSparklineAxisMinMax.Group;
				break;
			case "custom":
				res = Asc.c_oAscSparklineAxisMinMax.Custom;
				break;
		}
		return res;
	}

	function ToXml_ST_SparklineAxisMinMax(val) {
		var res = null;
		switch (val) {
			case Asc.c_oAscSparklineAxisMinMax.Individual:
				res = "individual";
				break;
			case Asc.c_oAscSparklineAxisMinMax.Group:
				res = "group";
				break;
			case Asc.c_oAscSparklineAxisMinMax.Custom:
				res = "custom";
				break;
		}
		return res;
	}

	function FromXml_ST_CellFormulaType(val) {
		var res = null;
		switch (val) {
			case "array":
				res = window["Asc"].ECellFormulaType.cellformulatypeArray;
				break;
			case "shared":
				res = window["Asc"].ECellFormulaType.cellformulatypeShared;
				break;
			case "dataTable":
				res = window["Asc"].ECellFormulaType.cellformulatypeDataTable;
				break;
		}
		return res;
	}

	function ToXml_ST_CellFormulaType(val) {
		var res = null;
		switch (val) {
			case window["Asc"].ECellFormulaType.cellformulatypeArray:
				res = "array";
				break;
			case window["Asc"].ECellFormulaType.cellformulatypeShared:
				res = "shared";
				break;
			case window["Asc"].ECellFormulaType.cellformulatypeDataTable:
				res = "dataTable";
				break;
		}
		return res;
	}

	function FromXml_ST_TableStyleType(val) {
		var res = null;
		switch (val) {
			case "wholeTable":
				res = Asc.ETableStyleType.tablestyletypeWholeTable;
				break;
			case "headerRow":
				res = Asc.ETableStyleType.tablestyletypeHeaderRow;
				break;
			case "totalRow":
				res = Asc.ETableStyleType.tablestyletypeTotalRow;
				break;
			case "firstColumn":
				res = Asc.ETableStyleType.tablestyletypeFirstColumn;
				break;
			case "lastColumn":
				res = Asc.ETableStyleType.tablestyletypeLastColumn;
				break;
			case "firstRowStripe":
				res = Asc.ETableStyleType.tablestyletypeFirstRowStripe;
				break;
			case "secondRowStripe":
				res = Asc.ETableStyleType.tablestyletypeSecondRowStripe;
				break;
			case "firstColumnStripe":
				res = Asc.ETableStyleType.tablestyletypeFirstColumnStripe;
				break;
			case "secondColumnStripe":
				res = Asc.ETableStyleType.tablestyletypeSecondColumnStripe;
				break;
			case "firstHeaderCell":
				res = Asc.ETableStyleType.tablestyletypeFirstHeaderCell;
				break;
			case "lastHeaderCell":
				res = Asc.ETableStyleType.tablestyletypeLastHeaderCell;
				break;
			case "firstTotalCell":
				res = Asc.ETableStyleType.tablestyletypeFirstTotalCell;
				break;
			case "lastTotalCell":
				res = Asc.ETableStyleType.tablestyletypeLastTotalCell;
				break;
			case "firstSubtotalColumn":
				res = Asc.ETableStyleType.tablestyletypeFirstSubtotalColumn;
				break;
			case "secondSubtotalColumn":
				res = Asc.ETableStyleType.tablestyletypeSecondSubtotalColumn;
				break;
			case "thirdSubtotalColumn":
				res = Asc.ETableStyleType.tablestyletypeThirdSubtotalColumn;
				break;
			case "firstSubtotalRow":
				res = Asc.ETableStyleType.tablestyletypeFirstSubtotalRow;
				break;
			case "secondSubtotalRow":
				res = Asc.ETableStyleType.tablestyletypeSecondSubtotalRow;
				break;
			case "thirdSubtotalRow":
				res = Asc.ETableStyleType.tablestyletypeThirdSubtotalRow;
				break;
			case "blankRow":
				res = Asc.ETableStyleType.tablestyletypeBlankRow;
				break;
			case "firstColumnSubheading":
				res = Asc.ETableStyleType.tablestyletypeFirstColumnSubheading;
				break;
			case "secondColumnSubheading":
				res = Asc.ETableStyleType.tablestyletypeSecondColumnSubheading;
				break;
			case "thirdColumnSubheading":
				res = Asc.ETableStyleType.tablestyletypeThirdColumnSubheading;
				break;
			case "firstRowSubheading":
				res = Asc.ETableStyleType.tablestyletypeFirstRowSubheading;
				break;
			case "secondRowSubheading":
				res = Asc.ETableStyleType.tablestyletypeSecondRowSubheading;
				break;
			case "thirdRowSubheading":
				res = Asc.ETableStyleType.tablestyletypeThirdRowSubheading;
				break;
			case "pageFieldLabels":
				res = Asc.ETableStyleType.tablestyletypePageFieldLabels;
				break;
			case "pageFieldValues":
				res = Asc.ETableStyleType.tablestyletypePageFieldValues;
				break;

		}
		return res;
	}

	function ToXml_ST_TableStyleType(val) {
		var res = null;
		switch (val) {
			case Asc.ETableStyleType.tablestyletypeWholeTable:
				res = "wholeTable";
				break;
			case Asc.ETableStyleType.tablestyletypeHeaderRow:
				res = "headerRow";
				break;
			case Asc.ETableStyleType.tablestyletypeTotalRow:
				res = "totalRow";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstColumn:
				res = "firstColumn";
				break;
			case Asc.ETableStyleType.tablestyletypeLastColumn:
				res = "lastColumn";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstRowStripe:
				res = "firstRowStripe";
				break;
			case Asc.ETableStyleType.tablestyletypeSecondRowStripe:
				res = "secondRowStripe";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstColumnStripe:
				res = "firstColumnStripe";
				break;
			case Asc.ETableStyleType.tablestyletypeSecondColumnStripe:
				res = "secondColumnStripe";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstHeaderCell:
				res = "firstHeaderCell";
				break;
			case Asc.ETableStyleType.tablestyletypeLastHeaderCell:
				res = "lastHeaderCell";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstTotalCell:
				res = "firstTotalCell";
				break;
			case Asc.ETableStyleType.tablestyletypeLastTotalCell:
				res = "lastTotalCell";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstSubtotalColumn:
				res = "firstSubtotalColumn";
				break;
			case Asc.ETableStyleType.tablestyletypeSecondSubtotalColumn:
				res = "secondSubtotalColumn";
				break;
			case Asc.ETableStyleType.tablestyletypeThirdSubtotalColumn:
				res = "thirdSubtotalColumn";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstSubtotalRow:
				res = "firstSubtotalRow";
				break;
			case Asc.ETableStyleType.tablestyletypeSecondSubtotalRow:
				res = "secondSubtotalRow";
				break;
			case Asc.ETableStyleType.tablestyletypeThirdSubtotalRow:
				res = "thirdSubtotalRow";
				break;
			case Asc.ETableStyleType.tablestyletypeBlankRow:
				res = "blankRow";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstColumnSubheading:
				res = "firstColumnSubheading";
				break;
			case Asc.ETableStyleType.tablestyletypeSecondColumnSubheading:
				res = "secondColumnSubheading";
				break;
			case Asc.ETableStyleType.tablestyletypeThirdColumnSubheading:
				res = "thirdColumnSubheading";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstRowSubheading:
				res = "firstRowSubheading";
				break;
			case Asc.ETableStyleType.tablestyletypeSecondRowSubheading:
				res = "secondRowSubheading";
				break;
			case Asc.ETableStyleType.tablestyletypeThirdRowSubheading:
				res = "thirdRowSubheading";
				break;
			case Asc.ETableStyleType.tablestyletypePageFieldLabels:
				res = "pageFieldLabels";
				break;
			case Asc.ETableStyleType.tablestyletypePageFieldValues:
				res = "pageFieldValues";
				break;

		}
		return res;
	}

	function FromXml_ST_FilterOperator(val) {
		var res = -1;
		if ("equal" === val) {
			res = Asc.c_oAscCustomAutoFilter.equals;
		} else if ("lessThan" === val) {
			res = Asc.c_oAscCustomAutoFilter.isLessThan;
		} else if ("lessThanOrEqual" === val) {
			res = Asc.c_oAscCustomAutoFilter.isLessThanOrEqualTo;
		} else if ("notEqual" === val) {
			res = Asc.c_oAscCustomAutoFilter.doesNotEqual;
		} else if ("greaterThanOrEqual" === val) {
			res = Asc.c_oAscCustomAutoFilter.isGreaterThanOrEqualTo;
		} else if ("greaterThan" === val) {
			res = Asc.c_oAscCustomAutoFilter.isGreaterThan;
		}
		return res;
	}

	function ToXml_ST_FilterOperator(val) {
		var res = "";
		if (Asc.c_oAscCustomAutoFilter.equals === val) {
			res = "equal";
		} else if (Asc.c_oAscCustomAutoFilter.isLessThan === val) {
			res = "lessThan";
		} else if (Asc.c_oAscCustomAutoFilter.isLessThanOrEqualTo === val) {
			res = "lessThanOrEqual";
		} else if (Asc.c_oAscCustomAutoFilter.doesNotEqual === val) {
			res = "notEqual";
		} else if (Asc.c_oAscCustomAutoFilter.isGreaterThanOrEqualTo === val) {
			res = "greaterThanOrEqual";
		} else if (Asc.c_oAscCustomAutoFilter.isGreaterThan === val) {
			res = "greaterThan";
		}
		return res;
	}

	function FromXml_ST_DynamicFilterType(val) {
		var res = -1;
		if ("null" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nullType;
		} else if ("aboveAverage" === val) {
			res = Asc.c_oAscDynamicAutoFilter.aboveAverage;
		} else if ("belowAverage" === val) {
			res = Asc.c_oAscDynamicAutoFilter.belowAverage;
		} else if ("tomorrow" === val) {
			res = Asc.c_oAscDynamicAutoFilter.tomorrow;
		} else if ("today" === val) {
			res = Asc.c_oAscDynamicAutoFilter.today;
		} else if ("yesterday" === val) {
			res = Asc.c_oAscDynamicAutoFilter.yesterday;
		} else if ("nextWeek" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nextWeek;
		} else if ("thisWeek" === val) {
			res = Asc.c_oAscDynamicAutoFilter.thisWeek;
		} else if ("lastWeek" === val) {
			res = Asc.c_oAscDynamicAutoFilter.lastWeek;
		} else if ("nextMonth" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nextMonth;
		} else if ("thisMonth" === val) {
			res = Asc.c_oAscDynamicAutoFilter.thisMonth;
		} else if ("lastMonth" === val) {
			res = Asc.c_oAscDynamicAutoFilter.lastMonth;
		} else if ("nextQuarter" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nextQuarter;
		} else if ("thisQuarter" === val) {
			res = Asc.c_oAscDynamicAutoFilter.thisQuarter;
		} else if ("lastQuarter" === val) {
			res = Asc.c_oAscDynamicAutoFilter.lastQuarter;
		} else if ("nextYear" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nextYear;
		} else if ("thisYear" === val) {
			res = Asc.c_oAscDynamicAutoFilter.thisYear;
		} else if ("lastYear" === val) {
			res = Asc.c_oAscDynamicAutoFilter.lastYear;
		} else if ("yearToDate" === val) {
			res = Asc.c_oAscDynamicAutoFilter.yearToDate;
		} else if ("Q1" === val) {
			res = Asc.c_oAscDynamicAutoFilter.q1;
		} else if ("Q2" === val) {
			res = Asc.c_oAscDynamicAutoFilter.q2;
		} else if ("Q3" === val) {
			res = Asc.c_oAscDynamicAutoFilter.q3;
		} else if ("Q4" === val) {
			res = Asc.c_oAscDynamicAutoFilter.q4;
		} else if ("M1" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m1;
		} else if ("M2" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m2;
		} else if ("M3" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m3;
		} else if ("M4" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m4;
		} else if ("M5" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m5;
		} else if ("M6" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m6;
		} else if ("M7" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m7;
		} else if ("M8" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m8;
		} else if ("M9" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m9;
		} else if ("M10" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m10;
		} else if ("M11" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m11;
		} else if ("M12" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m12;
		}
		return res;
	}

	function ToXml_ST_DynamicFilterType(val) {
		var res = "";
		if (Asc.c_oAscDynamicAutoFilter.nullType === val) {
			res = "null";
		} else if (Asc.c_oAscDynamicAutoFilter.aboveAverage === val) {
			res = "aboveAverage";
		} else if (Asc.c_oAscDynamicAutoFilter.belowAverage === val) {
			res = "belowAverage";
		} else if (Asc.c_oAscDynamicAutoFilter.tomorrow === val) {
			res = "tomorrow";
		} else if (Asc.c_oAscDynamicAutoFilter.today === val) {
			res = "today";
		} else if (Asc.c_oAscDynamicAutoFilter.yesterday === val) {
			res = "yesterday";
		} else if (Asc.c_oAscDynamicAutoFilter.nextWeek === val) {
			res = "nextWeek";
		} else if (Asc.c_oAscDynamicAutoFilter.thisWeek === val) {
			res = "thisWeek";
		} else if (Asc.c_oAscDynamicAutoFilter.lastWeek === val) {
			res = "lastWeek";
		} else if (Asc.c_oAscDynamicAutoFilter.nextMonth === val) {
			res = "nextMonth";
		} else if (Asc.c_oAscDynamicAutoFilter.thisMonth === val) {
			res = "thisMonth";
		} else if (Asc.c_oAscDynamicAutoFilter.lastMonth === val) {
			res = "lastMonth";
		} else if (Asc.c_oAscDynamicAutoFilter.nextQuarter === val) {
			res = "nextQuarter";
		} else if (Asc.c_oAscDynamicAutoFilter.thisQuarter === val) {
			res = "thisQuarter";
		} else if (Asc.c_oAscDynamicAutoFilter.lastQuarter === val) {
			res = "lastQuarter";
		} else if (Asc.c_oAscDynamicAutoFilter.nextYear === val) {
			res = "nextYear";
		} else if (Asc.c_oAscDynamicAutoFilter.thisYear === val) {
			res = "thisYear";
		} else if (Asc.c_oAscDynamicAutoFilter.lastYear === val) {
			res = "lastYear";
		} else if (Asc.c_oAscDynamicAutoFilter.yearToDate === val) {
			res = "yearToDate";
		} else if (Asc.c_oAscDynamicAutoFilter.q1 === val) {
			res = "Q1";
		} else if (Asc.c_oAscDynamicAutoFilter.q2 === val) {
			res = "Q2";
		} else if (Asc.c_oAscDynamicAutoFilter.q3 === val) {
			res = "Q3";
		} else if (Asc.c_oAscDynamicAutoFilter.q4 === val) {
			res = "Q4";
		} else if (Asc.c_oAscDynamicAutoFilter.m1 === val) {
			res = "M1";
		} else if (Asc.c_oAscDynamicAutoFilter.m2 === val) {
			res = "M2";
		} else if (Asc.c_oAscDynamicAutoFilter.m3 === val) {
			res = "M3";
		} else if (Asc.c_oAscDynamicAutoFilter.m4 === val) {
			res = "M4";
		} else if (Asc.c_oAscDynamicAutoFilter.m5 === val) {
			res = "M5";
		} else if (Asc.c_oAscDynamicAutoFilter.m6 === val) {
			res = "M6";
		} else if (Asc.c_oAscDynamicAutoFilter.m7 === val) {
			res = "M7";
		} else if (Asc.c_oAscDynamicAutoFilter.m8 === val) {
			res = "M8";
		} else if (Asc.c_oAscDynamicAutoFilter.m9 === val) {
			res = "M9";
		} else if (Asc.c_oAscDynamicAutoFilter.m10 === val) {
			res = "M10";
		} else if (Asc.c_oAscDynamicAutoFilter.m11 === val) {
			res = "M11";
		} else if (Asc.c_oAscDynamicAutoFilter.m12 === val) {
			res = "M12";
		}
		return res;
	}

	function FromXml_ST_DateTimeGrouping(val) {
		var res = -1;
		if ("year" === val) {
			res = Asc.EDateTimeGroup.datetimegroupYear;
		} else if ("month" === val) {
			res = Asc.EDateTimeGroup.datetimegroupMonth;
		} else if ("day" === val) {
			res = Asc.EDateTimeGroup.datetimegroupDay;
		} else if ("hour" === val) {
			res = Asc.EDateTimeGroup.datetimegroupHour;
		} else if ("minute" === val) {
			res = Asc.EDateTimeGroup.datetimegroupMinute;
		} else if ("second" === val) {
			res = Asc.EDateTimeGroup.datetimegroupSecond;
		}
		return res;
	}

	function ToXml_ST_DateTimeGrouping(val) {
		var res = "";
		if (Asc.EDateTimeGroup.datetimegroupYear === val) {
			res = "year";
		} else if (Asc.EDateTimeGroup.datetimegroupMonth === val) {
			res = "month";
		} else if (Asc.EDateTimeGroup.datetimegroupDay === val) {
			res = "day";
		} else if (Asc.EDateTimeGroup.datetimegroupHour === val) {
			res = "hour";
		} else if (Asc.EDateTimeGroup.datetimegroupMinute === val) {
			res = "minute";
		} else if (Asc.EDateTimeGroup.datetimegroupSecond === val) {
			res = "second";
		}
		return res;
	}

	function ToXml_ST_HorizontalAlignment(val, default_null) {
		var res = default_null ? null : -1;
		switch (val) {
			case -1:
				res = "general";
				break;
			case AscCommon.align_Left:
				res = "left";
				break;
			case AscCommon.align_Center:
				res = "center";
				break;
			case AscCommon.align_Right:
				res = "right";
				break;
			case AscCommon.align_Justify:
				res = "justify";
				break;
		}
		return res;
	}

	function FromXml_ST_HorizontalAlignment(val, default_null) {
		var res = default_null ? null: -1;//general == null
		if ("left" === val) {
			res = AscCommon.align_Left;
		} else if ("center" === val) {
			res = AscCommon.align_Center;
		} else if ("right" === val) {
			res = AscCommon.align_Right;
		} else if ("fill" === val) {
			res = AscCommon.align_Justify;
		} else if ("justify" === val) {
			res = AscCommon.align_Justify;
		} else if ("centerContinuous" === val) {
			res = AscCommon.align_Center;
		} else if ("distributed" === val) {
			res = AscCommon.align_Justify;
		}
		return res;
	}

	function ToXml_ST_VerticalAlignment(val, default_null) {
		var res = default_null ? null : -1;
		switch (val) {
			case Asc.c_oAscVAlign.Top:
				res = "top";
				break;
			case Asc.c_oAscVAlign.Center:
				res = "center";
				break;
			case Asc.c_oAscVAlign.Bottom:
				res = "bottom";
				break;
			case Asc.c_oAscVAlign.Just:
				res = "justify";
				break;
			case Asc.c_oAscVAlign.Dist:
				res = "distributed";
				break;
		}
		return res;
	}

	function FromXml_ST_VerticalAlignment(val) {
		var res = -1;
		if ("top" === val) {
			res = Asc.c_oAscVAlign.Top;
		} else if ("center" === val) {
			res = Asc.c_oAscVAlign.Center;
		} else if ("bottom" === val) {
			res = Asc.c_oAscVAlign.Bottom;
		} else if ("justify" === val) {
			res = Asc.c_oAscVAlign.Just;
		} else if ("distributed" === val) {
			res = Asc.c_oAscVAlign.Dist;
		}
		return res;
	}

	function ToXml_ST_CfvoType(nType) {
		var sType = "";
		switch (nType) {
			case AscCommonExcel.ECfvoType.Formula:
				sType = "formula";
				break;
			case AscCommonExcel.ECfvoType.Maximum:
				sType = "max";
				break;
			case AscCommonExcel.ECfvoType.Minimum:
				sType = "min";
				break;
			case AscCommonExcel.ECfvoType.Number:
				sType = "num";
				break;
			case AscCommonExcel.ECfvoType.Percent:
				sType = "percent";
				break;
			case AscCommonExcel.ECfvoType.Percentile:
				sType = "percentile";
				break;
			case AscCommonExcel.ECfvoType.AutoMin:
				sType = "autoMin";
				break;
			case AscCommonExcel.ECfvoType.AutoMax:
				sType = "autoMax";
				break;
		}

		return sType;
	}

	function FromXml_ST_CfvoType(sType) {
		var nType = -1;
		switch (sType) {
			case "formula":
				nType = AscCommonExcel.ECfvoType.Formula;
				break;
			case "max":
				nType = AscCommonExcel.ECfvoType.Maximum;
				break;
			case "min":
				nType = AscCommonExcel.ECfvoType.Minimum;
				break;
			case "num":
				nType = AscCommonExcel.ECfvoType.Number;
				break;
			case "percent":
				nType = AscCommonExcel.ECfvoType.Percent;
				break;
			case "percentile":
				nType = AscCommonExcel.ECfvoType.Percentile;
				break;
			case "autoMin":
				nType = AscCommonExcel.ECfvoType.AutoMin;
				break;
			case "autoMax":
				nType = AscCommonExcel.ECfvoType.AutoMax;
				break;
		}

		return nType;
	}

	function ToXml_ST_DataValidationErrorStyle(nType) {
		var sType = undefined;
		switch (nType) {
			case Asc.EDataValidationErrorStyle.Stop:
				sType = "stop";
				break;
			case Asc.EDataValidationErrorStyle.Warning:
				sType = "warning";
				break;
			case Asc.EDataValidationErrorStyle.Information:
				sType = "information";
				break;
		}

		return sType;
	}

	function FromXml_ST_DataValidationErrorStyle(sType) {
		var nType = undefined;
		switch (sType) {
			case "stop":
				nType = Asc.EDataValidationErrorStyle.Stop;
				break;
			case "warning":
				nType = Asc.EDataValidationErrorStyle.Warning;
				break;
			case "information":
				nType = Asc.EDataValidationErrorStyle.Information;
				break;
		}

		return nType;
	}

	function FromXml_ST_GradientType(val, default_null) {
		var res = default_null ? null : -1;
		if ("linear" === val) {
			res = Asc.c_oAscFillGradType.GRAD_LINEAR;
		} else if ("path" === val) {
			res = Asc.c_oAscFillGradType.GRAD_PATH;
		}
		return res;
	}

	function ToXml_ST_GradientType(val, default_null) {
		var res = default_null ? null : -1;
		if (Asc.c_oAscFillGradType.GRAD_LINEAR === val) {
			res = "linear";
		} else if (Asc.c_oAscFillGradType.GRAD_PATH === val) {
			res = "path";
		}
		return res;
	}

	function FromXml_ST_DataBarDirection(val) {
		var res = undefined;
		switch (val) {
			case "context":
				res = AscCommonExcel.EDataBarDirection.context;
				break;
			case "leftToRight":
				res = AscCommonExcel.EDataBarDirection.leftToRight;
				break;
			case "rightToLeft":
				res = AscCommonExcel.EDataBarDirection.rightToLeft;
				break;
		}
		return res;
	}

	function ToXml_ST_DataBarDirection(val) {
		var res = undefined;
		switch (val) {
			case AscCommonExcel.EDataBarDirection.context:
				res = "context";
				break;
			case AscCommonExcel.EDataBarDirection.leftToRight:
				res = "leftToRight";
				break;
			case AscCommonExcel.EDataBarDirection.rightToLeft:
				res = "rightToLeft";
				break;
		}
		return res;
	}

	function FromXml_ST_DataBarAxisPosition(val) {
		var res = undefined;
		switch (val) {
			case "automatic":
				res = AscCommonExcel.EDataBarAxisPosition.context;
				break;
			case "middle":
				res = AscCommonExcel.EDataBarAxisPosition.middle;
				break;
			case "none":
				res = AscCommonExcel.EDataBarAxisPosition.none;
				break;
		}
		return res;
	}

	function ToXml_ST_DataBarAxisPosition(val) {
		var res = undefined;
		switch (val) {
			case AscCommonExcel.EDataBarAxisPosition.automatic:
				res = "automatic";
				break;
			case AscCommonExcel.EDataBarAxisPosition.middle:
				res = "middle";
				break;
			case AscCommonExcel.EDataBarAxisPosition.none:
				res = "none";
				break;
		}
		return res;
	}

	function FromXml_ST_CellValueType(val) {
		var res = undefined;
		switch (val) {
			case "s":
				res = CellValueType.String;
				break;
			case "str":
				res = CellValueType.String;
				break;
			case "n":
				res = CellValueType.Number;
				break;
			case "e":
				res = CellValueType.Error;
				break;
			case "b":
				res =  CellValueType.Bool;
				break;
			case "inlineStr":
				res = CellValueType.String;
				break;
			case "d":
				res = CellValueType.String;
				break;
		}
		return res;
	}

	function ToXml_ST_CellValueType(val) {
		var res = undefined;
		switch (val) {
			case CellValueType.String:
				res = "s";
				break;
			/*case CellValueType.String:
				res = "str";
				break;*/
			case CellValueType.Number:
				res = "n";
				break;
			case CellValueType.Error:
				res = "e";
				break;
			case CellValueType.Bool:
				res =   "b";
				break;
			/*case "inlineStr":
				res = CellValueType.String;
				break;
			case "d":
				res = CellValueType.String;
				break;*/
		}
		return res;
	}



	//additional functions
	function prepareCommentsToWrite(m_mapComments, personList) {
		var mapByAuthors = [];
		var pComments = null;
		var pThreadedComments = null;
		var aComments;

		var getThreadedComment = function (oCommentData) {
			var res = new CT_CThreadedComment();

			var sOOTime = oCommentData.asc_getOnlyOfficeTime();
			if (sOOTime) {
				res.dT = new Date(sOOTime - 0).toISOString().slice(0, 22) + "Z";
			}
			var userId = oCommentData.asc_getUserId();
			var displayName = oCommentData.asc_getUserName();
			var providerId = oCommentData.asc_getProviderId();
			var person = personList.find(function isPrime(element) {
				return userId === element.userId && displayName === element.displayName && providerId === element.providerId;
			});

			if (!person) {
				person = {id: AscCommon.CreateGUID(), userId: userId, displayName: displayName, providerId: providerId};
				personList.push(person);
			}

			res.personId = person.id;
			var guid = oCommentData.asc_getGuid();
			if (guid) {
				res.id = guid;
			}
			var solved = oCommentData.asc_getSolved();
			if (null != solved) {
				res.done = solved;
			}

			var text = oCommentData.asc_getText();
			if (text) {
				res.text = text;
			}

			if (oCommentData.aReplies && oCommentData.aReplies.length > 0) {
				for (i = 0; i < oCommentData.aReplies.length; ++i) {
					res.m_arrReplies.push(getThreadedComment(oCommentData.aReplies[i]));
				}
			}

			return res;
		};

		for (var it = 0; it < m_mapComments.length; it++) {
			var pCommentItem = m_mapComments[it];
			if (/*pCommentItem.IsValid()*/pCommentItem) {
				var pNewComment = new CT_CComment();
				if (null != pCommentItem.nRow && null != pCommentItem.nCol) {
					pNewComment.ref = new Asc.Range(pCommentItem.nCol, pCommentItem.nRow, pCommentItem.nCol, pCommentItem.nRow).getName();
				}

				var needAddEmptyAuthor = false;
				var saveThreadedComments = true;
				if (saveThreadedComments && pCommentItem.isValidThreadComment()/*pCommentItem.pThreadedComment*/) {
					if (null === pThreadedComments) {
						pThreadedComments = new CT_CThreadedComments();
					}

					var pThreadedComment = getThreadedComment(pCommentItem)//pCommentItem->m_pThreadedComment;
					if (pNewComment.ref) {
						pThreadedComment.ref = pNewComment.ref;
					}

					if (!pThreadedComment.id) {
						pThreadedComment.id = AscCommon.CreateGUID();
					}

					pNewComment.uid = pThreadedComment.id;
					pCommentItem.m_sAuthor = "tc=" + pThreadedComment.id;

					//pThreadedComment.m_arrReplies = pCommentItem.aReplies;

					//pCommentItem->m_oText.Init();


					//var mapPersonList = this.InitSaveManager.personList;
					/*if (m_oWorkbook.m_pPersonList)
					{
						mapPersonList = m_oWorkbook.m_pPersonList->GetPersonList();
					}*/
					//BinaryCommentReader::addThreadedComment(pCommentItem.text, pThreadedComment, mapPersonList);

					pNewComment.generateText(pCommentItem, personList);


					pThreadedComments.arr.push(pThreadedComment);
					for (var i = 0; i < pThreadedComment.m_arrReplies.length; ++i) {
						pThreadedComment.m_arrReplies[i].parentId = pThreadedComment.id;
						pThreadedComment.m_arrReplies[i].ref = pThreadedComment.ref;
						if (null === pThreadedComment.m_arrReplies[i].id) {
							pThreadedComment.m_arrReplies[i].id = AscCommon.CreateGUID();
						}
						pThreadedComments.arr.push(pThreadedComment.m_arrReplies[i]);
					}
				} else {
					pNewComment.generateText(pCommentItem, personList);
					pNewComment.uid = AscCommon.CreateGUID();
					pNewComment.authorId = 0;

					needAddEmptyAuthor = true;
				}

				if (undefined !== pCommentItem.m_sAuthor) {
					var sAuthor = pCommentItem.m_sAuthor;
					var pFind;
					for (var j = 0; j < mapByAuthors.length; j++) {
						if (mapByAuthors[j] === sAuthor) {
							pFind = j;
							break;
						}
					}

					var nAuthorId;
					if (null != pFind) {
						nAuthorId = pFind;
					} else {
						nAuthorId = mapByAuthors.length;
						mapByAuthors.push(sAuthor);
						if (pComments === null) {
							pComments = new CT_CComments();
							pComments.commentList = new CT_CCommentList();
							aComments = pComments.commentList.arr;
							pComments.authors = new CT_CAuthors();
						}
						pComments.authors.arr.push(sAuthor);
					}

					pNewComment.authorId = nAuthorId;
				}


				//pNewComment.generateText(pCommentItem);

				if (pComments === null) {
					pComments = new CT_CComments();
					pComments.commentList = new CT_CCommentList();
					aComments = pComments.commentList.arr;
					pComments.authors = new CT_CAuthors();
					if (needAddEmptyAuthor) {
						pComments.authors.arr.push("");
					}
				}
				aComments.push(pNewComment);
			}
			/*else if (NULL != pCommentItem->m_pThreadedComment)
			{
				RELEASEOBJECT(pCommentItem->m_pThreadedComment);
				for (size_t i = 0; i < pCommentItem->m_pThreadedComment->m_arrReplies.size(); ++i)
				{
					RELEASEOBJECT(pCommentItem->m_pThreadedComment->m_arrReplies[i]);
				}
			}*/
		}

		/*NSCommon::smart_ptr<OOX::File> pCommentsFile(pComments);
		m_pCurWorksheet->Add(pCommentsFile);*/

		return pComments || pThreadedComments ? {comments: pComments, threadedComments: pThreadedComments} : null;
	}

	function getSimpleArrayFromXml(reader, childName, attrName, attrType) {
		var res = [];
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name2 = reader.GetNameNoNS();
			if (childName === name2) {
				while (reader.MoveToNextAttribute()) {
					if (attrName === reader.GetNameNoNS()) {
						res.push(reader.GetValue());
					}
				}
			}
		}
		return res;
	}

	function readOneAttr(reader, attrName, func) {
		var res = null;

		while (reader.MoveToNextAttribute()) {
			if (attrName === reader.GetNameNoNS()) {
				func();
			}
		}

		return res;
	}

	//возможно стоит перенести в writer, пока не расширяю интерфейс
	function toXML2(writer, sName, m_sText) {
		if (m_sText) {
			writer.WriteXmlString("<");
			writer.WriteXmlString(sName);
			writer.WriteXmlString(">");

			writer.WriteXmlStringEncode(m_sText);

			writer.WriteXmlString("</");
			writer.WriteXmlString(sName);
			writer.WriteXmlString(">");
		}
	}

	function boolToNumber(val) {
		return val ? 1 : 0;
	}

	function prepareTextFromXml(val) {
		//TODO поработать с текстом, проблемы с переносом строки
		if (!val) {
			return val;
		}
		val = val.replace(/&#xD;&#xA;/g, "\n");
		val = val.replace(/_x000D_\r\n/g, "\n");
		val = val.replace(/\r\n/g, "\n");

		val = unescape_ST_Xstring(val);

		return val;
	}

	function prepareTextFormatFromXml(val) {
		//TODO обрубаю &quot, чтобы при записи & внутри него не заменился на &apos; (в функции WriteXmlNullableAttributeStringEncode)
		if (!val) {
			return val;
		}
		val = prepareTextFromXml(val);
		val = val.replace(/&apos;/g, "'");
		val = val.replace(/&quot;/g, '"');
		return val;
	}

	function prepareTextToXml(val) {
		//x2t пишет &#xA; , но в WriteXmlValueStringEncode & заменяется
		//val = val.replace(/\n/g, "\&#xA;");
		//от мс приходит \r\n
		//val = val.replace(/\n/g, "\r\n");
		val = escape_ST_Xstring(val);
		return val;
	}

	function unescape_ST_Xstring(s) {
		var res = "";
		for (var i = 0; i < s.length; i++) {
			if (i + 6 < s.length && s[i] === '_' && s[i + 1] === 'x' && !isNaN(s[i + 2]) && !isNaN(s[i + 3]) && s[i + 6] === '_') {
				res += hex2Str(s.substring(i + 2, i + 6));
				i += 6;
			} else {
				res += s[i];
			}
		}
		return res;
	}

	var match_hex = "^_x[0-9A-F]{4}_";
	function escape_ST_Xstring(wstr) {
		var ret_val = "";

		for (var i = 0; i < wstr.length; i++) {
			var str = wstr[i];
			var strCode = str.charCodeAt(0);
			if (strCode < 0x0009 || strCode === 0x000B || strCode === 0x000C || strCode > 0x000D && strCode < 0x0020 || strCode > 0xD7FF && strCode < 0xE000 || strCode > 0xFFFD) {
				ret_val += "_x" + wchar_t2hex_str(strCode) + "_";
			} else if ('_' === str && wstr.substring(i, 7).match(match_hex)) {
				ret_val += "_x005F_";
			} else {
				ret_val += str;
			}
		}
		return ret_val;
	}

	function wchar_t2hex_str(val)
	{
		var hex = "0123456789ABCDEF";
		var code_string = "";
		code_string += hex[(val & 0xf000) >> 12];
		code_string += hex[(val & 0x0f00) >> 8];
		code_string += hex[(val & 0x00f0) >> 4];
		code_string += hex[(val & 0x000f)];

		if (code_string === "0000") {
			return val;
		}

		return code_string;
	}

	function hex2Str(hex) {
		var res = "";
		for (var i = 0; i < hex.length; i += 2) {
			var val = parseInt(hex.substr(i, 2), 16);
			if (val) {
				res += String.fromCharCode(val);
			}
		}
		return res;
	}

	//for uri/namespaces
	var extUri = {
		conditionalFormattings: "{78C0D931-6437-407d-A8EE-F0AAD7539E65}",
		dataValidations: "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}",
		sparklineGroups: "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}",
		slicerList: "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}",
		slicerListExt: "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}",
		protectedRanges: "{FC87AEE6-9EDD-4A0A-B7FB-166176984837}",
		ignoredErrors: "{01252117-D84E-4E92-8308-4BE1C098FCBB}",
		webExtensions: "{F7C9EE02-42E1-4005-9D12-6889AFFD525C}",
		timelineRefs: "{7E03D99C-DC04-49d9-9315-930204A7B6E9}",
		slicerStyles: "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}",
		dxfs: "{46F421CA-312F-682F-3DD2-61675219B42D}",
		slicerCaches: "{BBE1A952-AA13-448e-AADC-164F8A28A991}",//m_sAdditionalNamespace = L"xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"";
		slicerCachesExt: "{46BE6895-7355-4a93-B00E-2C351335B9C9}",//m_sAdditionalNamespace = L"xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"";
		tableSlicerCache: "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}"//m_sAdditionalNamespace = _T("xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"")
	};

	var additionalNamespace = {
		slicerCaches: " xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"",
		slicerCachesExt:" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"",
		tableSlicerCache: " xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"",
		slicerList: " xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"",
		slicerListExt: " xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\""
	};

	//START READ/WRITE functions
	AscCommonExcel.Workbook.prototype.toZip = function (zip, context) {
		context.wb = this;
		context.InitSaveManager = new AscCommonExcel.InitSaveManager(this);

		//функция дёргается в serialize перед записью ws
		context.InitSaveManager._prepeareStyles(context.stylesForWrite);

		var memory = new AscCommon.CMemory();
		memory.context = context;
		var filePart = new AscCommon.openXml.OpenXmlPackage(zip, memory);

		//core
		if (this.Core) {
			var corePart = filePart.addPart(AscCommon.openXml.Types.coreFileProperties);
			corePart.part.setDataXml(this.Core, memory);
			memory.Seek(0);
		}

		//app
		if (this.App) {
			var appPart = filePart.addPart(AscCommon.openXml.Types.extendedFileProperties);
			appPart.part.setDataXml(this.App, memory);
			memory.Seek(0);
		}

		//wb
		var wbXml = new CT_Workbook(this);
		var wbPart = filePart.addPart(AscCommon.openXml.Types.workbook);
		wbPart.part.setDataXml(wbXml, memory);
		memory.Seek(0);

		//SharedStrings
		if (context.oSharedStrings.index > 0) {
			var sharedString = new CT_SharedStrings();
			sharedString.initFromMap(this, context.oSharedStrings);
			var sharedStringPart = wbPart.part.addPart(AscCommon.openXml.Types.sharedStringTable);
			sharedStringPart.part.setDataXml(sharedString, memory);
			memory.Seek(0);
		}

		//styles
		//на чтение используется CT_Stylesheet, на запись StylesForWrite
		var stylesheetPart = wbPart.part.addPart(AscCommon.openXml.Types.workbookStyles);
		stylesheetPart.part.setDataXml(context.stylesForWrite, memory);
		memory.Seek(0);

		//theme
		var themePart = wbPart.part.addPart(AscCommon.openXml.Types.theme);
		themePart.part.setDataXml(this.theme, memory);
		memory.Seek(0);

		//jsaMacros
		var jsaMacros = this.oApi.macros.GetData();
		if (jsaMacros) {
			memory.WriteXmlString(jsaMacros);
			var jsaData = memory.GetDataUint8();
			var jsaPart = wbPart.part.addPart(AscCommon.openXml.Types.jsaProject);
			jsaPart.part.setData(jsaData);
			memory.Seek(0);
		}

		//vbaMacros
		var vbaMacros = this.oApi.vbaMacros;
		if (vbaMacros) {
			var vbaPart = wbPart.part.addPart(AscCommon.openXml.Types.vbaProject);
			vbaPart.part.setData(vbaMacros);
			memory.Seek(0);
		}

		//person list
		if (context.InitSaveManager.personList && context.InitSaveManager.personList.length) {
			var oPerson = new CT_PersonList();
			var personPart = wbPart.part.addPart(AscCommon.openXml.Types.person);
			oPerson.personList = context.InitSaveManager.personList;
			personPart.part.setDataXml(oPerson, memory);
			memory.Seek(0);
		}

		//connections
		if (this.connections && !Array.isArray(this.connections)) {
			memory.WriteXmlString(this.connections);
			var connectionsData = memory.GetDataUint8();
			var connectionsPart = wbPart.part.addPart(AscCommon.openXml.Types.connections);
			connectionsPart.part.setData(connectionsData);
			memory.Seek(0);
		}

		//customXmls
		if (this.customXmls) {
			for (var i = 0; i < this.customXmls.length; i++) {
				if (this.customXmls[i].item) {
					/*var customXmlPart = wbPart.part.addPart(AscCommon.openXml.Types.customXml);
					customXmlPart.part.setData(this.customXmls[i].item);
					memory.Seek(0);

					var customXmlPartProps = customXmlPart.part.addPart(AscCommon.openXml.Types.customXmlProps);
					customXmlPartProps.part.setData(this.customXmls[i].itemProps);
					memory.Seek(0);*/

					var customXmlPart = wbPart.part.addPart(AscCommon.openXml.Types.customXml);
					memory.WriteXmlString(this.customXmls[i].item);
					var customXmlData = memory.GetDataUint8();
					customXmlPart.part.setData(customXmlData);
					memory.Seek(0);

					var customXmlPartProps = customXmlPart.part.addPart(AscCommon.openXml.Types.customXmlProps);
					memory.WriteXmlString(this.customXmls[i].itemProps);
					var customXmlPropsData = memory.GetDataUint8();
					customXmlPartProps.part.setData(customXmlPropsData);
					memory.Seek(0);
				}
			}
		}

		//comments wb
		if (this.aComments) {
			var binaryComments = AscCommonExcel.WriteWbComments(this);
			if (binaryComments) {
				var wbCommentsPart = wbPart.part.addPart(AscCommon.openXml.Types.workbookComment);
				wbCommentsPart.part.setData(binaryComments);
				memory.Seek(0);
			}
		}
	};

	//****workbook****
	function CT_Workbook(wb) {
		//Members
		this.wb = wb;
		this.sheets = null;
		this.pivotCaches = null;
		this.externalReferences = [];
		this.extLst = null;
		this.slicerCachesIds = [];
		this.newDefinedNames = [];
	}

	CT_Workbook.prototype.fromXml = function (reader) {
		if (!reader.ReadNextNode()) {
			return;
		}
		if ("workbook" !== reader.GetNameNoNS()) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}

		var t = this, val;
		if ("workbook" === reader.GetNameNoNS()) {
			var depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				var name = reader.GetNameNoNS();
				if ("sheets" === name) {
					var sheets = new CT_Sheets(this.wb);
					sheets.fromXml(reader);
					this.sheets = sheets.sheets;
				} else if ("pivotCaches" === name) {
					var pivotCaches = new CT_PivotCaches();
					pivotCaches.fromXml(reader);
					this.pivotCaches = pivotCaches.pivotCaches;
				} else if ("workbookPr" === name) {
					val = new CT_WorkbookPr();
					val.fromXml(reader);
					this.wb.WorkbookPr = val.val;
				} else if ("bookViews" === name) {
					val = new CT_BookViews(t.wb);
					val.fromXml(reader)
					if (val.val && val.val.nActive != null) {
						t.wb.nActive = val.val.nActive;
					}
				} else if ("definedNames" === name) {
					if (reader.IsEmptyNode()) {
						continue;
					}
					var depthInside = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depthInside)) {
						var nameInside = reader.GetNameNoNS();
						if ("definedName" === nameInside || "NamedRange" === nameInside) {
							var oNewDefinedName = new Asc.asc_CDefName();
							oNewDefinedName.fromXml(reader);
							//PostLoadPrepareDefNames делается в api
							this.newDefinedNames.push(oNewDefinedName);
						}
					}

				} else if ("calcPr" === name) {
					val = new CT_CalcPr();
					val.fromXml(reader)
					this.wb.calcPr = val.val;
				} else if ("externalReferences" === name) {
					reader.readXmlArray("externalReference", function () {
						while (reader.MoveToNextAttribute()) {
							if ("id" === reader.GetNameNoNS()) {
								val = reader.GetValue();
								t.externalReferences.push(val);
							}
						}
					});
				} else if ("comments" === name) {
					//TODO разобраться как читаются/записываются комменты в бинарник


				} else if ("slicerCaches" === name) {
					reader.readXmlArray("slicerCache", function () {
						while (reader.MoveToNextAttribute()) {
							if ("id" === reader.GetNameNoNS()) {
								t.slicerCachesIds.push(reader.GetValue());
							}
						}
					});
				} else if ("workbookProtection" === name) {
					var workbookProtection = new Asc.CWorkbookProtection(this.wb);
					workbookProtection.fromXml(reader);
					this.wb.workbookProtection = workbookProtection;
				} else if ("extLst" === name) {
					//пока использую только slicerCaches
					var extLst = new COfficeArtExtensionList(this);
					extLst.fromXml(reader);
					this.extLst = extLst.arrExt;
				} else if ("fileVersion" === name) {
				}
				//c_oSerWorkbookTypes.VbaProject - read in api, save as array
				//c_oSerWorkbookTypes.JsaProject -> read in api
			}
		}
	};

	CT_Workbook.prototype.toXml = function (writer) {
		writer.WriteXmlString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
		writer.WriteXmlNodeStart("workbook");
		writer.WriteXmlString(' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
		writer.WriteXmlAttributesEnd();

		let pivotCaches = {};
		let pivotCacheIndex = this.wb.preparePivotForSerialization(pivotCaches, null);
		let ct_pivotCaches = null;
		if (pivotCacheIndex > 0) {
			ct_pivotCaches = new CT_PivotCaches();
			ct_pivotCaches.initFromPivotCaches(writer.context.part, writer, pivotCaches);
		}

		//WorkbookPr
		if (this.wb.WorkbookPr) {
			var workbookPr = new CT_WorkbookPr(this.wb.WorkbookPr);
			workbookPr.toXml(writer, "workbookPr");
		}
		//workbookProtection
		if (this.wb.workbookProtection) {
			//CWorkbookProtection
			this.wb.workbookProtection.toXml(writer, "workbookProtection");
		}

		//bookViews
		var bookViews = new CT_BookViews(this.wb);
		bookViews.toXml(writer, "bookViews");

		//sheets
		var sheetsXml = new CT_Sheets(this.wb);
		sheetsXml.toXml(writer, "sheets");

		//externalReferences
		var externalReferences = new CT_ExternalReferences(this.wb);
		externalReferences.toXml(writer, "externalReferences");

		//ссылки на slicerCache всегда пишеи через ext, но с разными префиксами slicerCachesIds/slicerCachesExtIds
		//сами slicerCache пишем в slicerCache[n].xml
		//cссылки на файлы slicerCache должны находиться в workbook->rels
		var i;
		var officeArtExtensionList, officeArtExtension;
		var slicerCaches = writer.context.InitSaveManager.getSlicersCache();
		var slicerCachesExt = writer.context.InitSaveManager.getSlicersCache(true);
		var doSlicerCaches = function (_slicerCaches, bExt) {
			if (_slicerCaches) {
				var slicerCachesIds = [];
				for (var j in _slicerCaches) {
					//тут пишем slicerCacheDefinition (slicerCache1.xml)
					var slicerCachePart = writer.context.part.addPart(AscCommon.openXml.Types.slicerCache);
					slicerCachePart.part.setDataXml(_slicerCaches[j], writer);
					slicerCachesIds.push(slicerCachePart.rId);
				}

				//ссылки на slicerCaches кладём в wb в extLst
				if (!officeArtExtensionList) {
					officeArtExtensionList = new COfficeArtExtensionList();
					officeArtExtension = new COfficeArtExtension();
					officeArtExtension.uri = bExt ? extUri.slicerCachesExt : extUri.slicerCaches;
					officeArtExtension.additionalNamespace = bExt ? additionalNamespace.slicerCachesExt : additionalNamespace.slicerCaches;
					if (bExt) {
						officeArtExtension.slicerCachesExtIds = slicerCachesIds;
					} else {
						officeArtExtension.slicerCachesIds = slicerCachesIds;
					}
					officeArtExtensionList.arrExt.push(officeArtExtension);
				}
			}
		}
		doSlicerCaches(slicerCaches);
		doSlicerCaches(slicerCachesExt, true);

		//asc_CDefName
		var defNameList = writer.context.InitSaveManager.defNameList;
		if (defNameList && defNameList.length) {
			writer.WriteXmlNodeStart("definedNames");
			writer.WriteXmlAttributesEnd();

			for (i = 0; i < defNameList.length; ++i) {
				if (defNameList[i]) {
					defNameList[i].toXml(writer, "definedName");
				}
			}

			writer.WriteXmlNodeEnd("definedNames");
		}

		if (null != this.wb.calcPr) {
			var calcPr = new CT_CalcPr(this.wb.calcPr);
			calcPr.toXml(writer, "calcPr");
		}

		if (ct_pivotCaches) {
			ct_pivotCaches.toXml(writer, "pivotCaches");
		}

		if (officeArtExtensionList) {
			officeArtExtensionList.toXml(writer);
		}

		writer.WriteXmlNodeEnd("workbook");
	};

	function CT_BookViews(wb, val) {
		this.wb = wb;
		this.val = val;
	}

	CT_BookViews.prototype.fromXml = function (reader) {
		var t = this;

		if (!this.val) {
			this.val = {};
		}

		reader.readXmlArray("workbookView", function () {
			while (reader.MoveToNextAttribute()) {
				if ("activeTab" === reader.GetNameNoNS()) {
					t.val.nActive = reader.GetValueInt();
				}
			}
		});
	};

	CT_BookViews.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		//TODO в x2t пишутся все остальные аттрибуты
		/*WritingStringNullableAttrInt(L"xWindow", m_oXWindow, m_oXWindow->GetValue());
		WritingStringNullableAttrInt(L"yWindow", m_oYWindow, m_oYWindow->GetValue());
		WritingStringNullableAttrInt(L"windowWidth", m_oWindowWidth, m_oWindowWidth->GetValue());
		WritingStringNullableAttrInt(L"windowHeight", m_oWindowHeight, m_oWindowHeight->GetValue());
		WritingStringNullableAttrInt(L"activeTab", m_oActiveTab, m_oActiveTab->GetValue());*/

		writer.WriteXmlNodeStart(ns + name/*'bookViews'*/);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNodeStart('workbookView');
		if (null != this.wb.nActive) {
			writer.WriteXmlAttributeNumber("activeTab", this.wb.nActive);
		}
		writer.WriteXmlAttributesEnd(true);
		writer.WriteXmlNodeEnd(ns + name);
	};

	function CT_CalcPr(val) {
		this.val = val ? val : {
			calcId: null,
			calcMode: null,
			fullCalcOnLoad: null,
			refMode: null,
			iterate: null,
			iterateCount: null,
			iterateDelta: null,
			fullPrecision: null,
			calcCompleted: null,
			calcOnSave: null,
			concurrentCalc: null,
			concurrentManualCount: null,
			forceFullCalc: null
		};
	}

	CT_CalcPr.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};
	CT_CalcPr.prototype.fromXml = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("calcId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.val.calcId = val;
			} else if ("calcMode" === reader.GetName()) {
				val = reader.GetValue();
				this.val.calcMode = val;
			} else if ("fullCalcOnLoad" === reader.GetName()) {
				val = reader.GetValueBool();
				this.val.fullCalcOnLoad = val;
			} else if ("refMode" === reader.GetName()) {
				val = reader.GetValue();
				this.val.refMode = val;
			} else if ("iterate" === reader.GetName()) {
				val = reader.GetValueBool();
				this.val.iterate = val;
			} else if ("iterateCount" === reader.GetName()) {
				val = reader.GetValueInt();
				this.val.iterateCount = val;
			} else if ("iterateDelta" === reader.GetName()) {
				val = reader.GetValue();
				this.val.iterateDelta = val;
			} else if ("fullPrecision" === reader.GetName()) {
				val = reader.GetValueBool();
				this.val.fullPrecision = val;
			} else if ("calcCompleted" === reader.GetName()) {
				val = reader.GetValueBool();
				this.val.calcCompleted = val;
			} else if ("calcOnSave" === reader.GetName()) {
				val = reader.GetValueBool();
				this.val.calcOnSave = val;
			} else if ("concurrentCalc" === reader.GetName()) {
				val = reader.GetValueBool();
				this.val.concurrentCalc = val;
			} else if ("concurrentManualCount" === reader.GetName()) {
				val = reader.GetValueInt();
				this.val.concurrentManualCount = val;
			} else if ("forceFullCalc" === reader.GetName()) {
				val = reader.GetValueBool();
				this.val.forceFullCalc = val;
			}
		}
	};

	CT_CalcPr.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}
		writer.WriteXmlNodeStart(ns + name/*'calcPr'*/);
		writer.WriteXmlNullableAttributeNumber("calcId", this.val.calcId);
		writer.WriteXmlNullableAttributeString("calcMode", this.val.calcMode);
		writer.WriteXmlNullableAttributeBool("fullCalcOnLoad", this.val.fullCalcOnLoad);
		writer.WriteXmlNullableAttributeString("refMode", this.val.refMode);
		writer.WriteXmlNullableAttributeBool("iterate", this.val.iterate);
		writer.WriteXmlNullableAttributeNumber("iterateCount", this.val.iterateCount);
		writer.WriteXmlNullableAttributeNumber("iterateDelta", this.val.iterateDelta);
		writer.WriteXmlNullableAttributeBool("fullPrecision", this.val.fullPrecision);
		writer.WriteXmlNullableAttributeBool("calcCompleted", this.val.calcCompleted);
		writer.WriteXmlNullableAttributeBool("calcOnSave", this.val.calcOnSave);
		writer.WriteXmlNullableAttributeBool("concurrentCalc", this.val.concurrentCalc);
		writer.WriteXmlNullableAttributeNumber("concurrentManualCount", this.val.concurrentManualCount);
		writer.WriteXmlNullableAttributeBool("forceFullCalc", this.val.forceFullCalc);
		writer.WriteXmlAttributesEnd(true);
	}

	function CT_WorkbookPr(val) {
		this.val = val ? val : {};
	}

	CT_WorkbookPr.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};
	CT_WorkbookPr.prototype.fromXml = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			/*if ("allowRefreshQuery" === reader.GetName()) {
				val = reader.GetValueBool();
				this.allowRefreshQuery = val;
			} else if ("autoCompressPictures" === reader.GetName()) {
				val = reader.GetValueBool();
				this.autoCompressPictures = val;
			} else if ("backupFile" === reader.GetName()) {
				val = reader.GetValueBool();
				this.backupFile = val;
			} else if ("checkCompatibility" === reader.GetName()) {
				val = reader.GetValueBool();
				this.checkCompatibility = val;
			} else if ("codeName" === reader.GetName()) {
				val = reader.GetValue();
				this.codeName = val;
			} else*/
			if ("date1904" === reader.GetName()) {
				val = reader.GetValueBool();
				this.val.Date1904 = val;
			} else if ("dateCompatibility" === reader.GetName()) {
				val = reader.GetValueBool();
				this.val.DateCompatibility = val;
			} /*else if ("defaultThemeVersion" === reader.GetName()) {
				val = reader.GetValueInt();
				this.defaultThemeVersion = val;
			} else if ("filterPrivacy" === reader.GetName()) {
				val = reader.GetValueBool();
				this.filterPrivacy = val;
			}*/ else if ("hidePivotFieldList" === reader.GetName()) {
				val = reader.GetValueBool();
				this.val.HidePivotFieldList = val;
			} /*else if ("promptedSolutions" === reader.GetName()) {
				val = reader.GetValueBool();
				this.promptedSolutions = val;
			} else if ("publishItems" === reader.GetName()) {
				val = reader.GetValueBool();
				this.publishItems = val;
			} else if ("refreshAllConnections" === reader.GetName()) {
				val = reader.GetValueBool();
				this.refreshAllConnections = val;
			} else if ("showBorderUnselectedTables" === reader.GetName()) {
				val = reader.GetValueBool();
				this.showBorderUnselectedTables = val;
			} else if ("showInkAnnotation" === reader.GetName()) {
				val = reader.GetValueBool();
				this.showInkAnnotation = val;
			} else if ("showObjects" === reader.GetName()) {
				val = reader.GetValue();
				this.showObjects = val;
			}*/ else if ("showPivotChartFilter" === reader.GetName()) {
				val = reader.GetValueBool();
				this.val.ShowPivotChartFilter = val;
			} /*else if ("updateLinks" === reader.GetName()) {
				val = reader.GetValue();
				this.updateLinks = val;
			}*/
		}
	};

	CT_WorkbookPr.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}
		writer.WriteXmlNodeStart(ns + name/*'workbookPr'*/);

		writer.WriteXmlNullableAttributeBool("date1904", this.val.Date1904);
		writer.WriteXmlNullableAttributeBool("dateCompatibility", this.val.DateCompatibility);
		writer.WriteXmlNullableAttributeBool("hidePivotFieldList", this.val.HidePivotFieldList);
		writer.WriteXmlNullableAttributeBool("showPivotChartFilter", this.val.ShowPivotChartFilter);
		writer.WriteXmlAttributesEnd();

		writer.WriteXmlNodeEnd(ns + name/*'workbookPr'*/);
	};

	Asc.CWorkbookProtection.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlString(ns + name/*workbookProtection*/);
		writer.WriteXmlNullableAttributeString("workbookAlgorithmName", AscCommonExcel.ToXml_ST_AlgorithmName(this.workbookAlgorithmName));
		writer.WriteXmlNullableAttributeString("workbookHashValue", this.workbookHashValue);
		writer.WriteXmlNullableAttributeString("workbookSaltValue", this.workbookSaltValue);
		writer.WriteXmlNullableAttributeNumber("workbookSpinCount", this.workbookSpinCount);
		writer.WriteXmlNullableAttributeNumber("lockStructure", boolToNumber(this.lockStructure));
		writer.WriteXmlNullableAttributeNumber("lockWindows", boolToNumber(this.lockWindows));
		writer.WriteXmlNullableAttributeString("workbookPassword", this.workbookPassword);
		writer.WriteXmlString("/>");
	};

	Asc.CWorkbookProtection.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	Asc.CWorkbookProtection.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("workbookAlgorithmName" === reader.GetName()) {
				val = reader.GetValue();
				this.workbookAlgorithmName = AscCommonExcel.FromXml_ST_AlgorithmName(val);
			} else if ("workbookHashValue" === reader.GetName()) {
				val = reader.GetValue();
				this.workbookHashValue = val;
			} else if ("workbookSaltValue" === reader.GetName()) {
				val = reader.GetValue();
				this.workbookSaltValue = val;
			} else if ("workbookSpinCount" === reader.GetName()) {
				val = reader.GetValueInt();
				this.workbookSpinCount = val;
			} else if ("revisionsAlgorithmName" === reader.GetName()) {
				val = reader.GetValue();
				this.revisionsAlgorithmName = AscCommonExcel.FromXml_ST_AlgorithmName(val);
			} else if ("revisionsHashValue" === reader.GetName()) {
				val = reader.GetValue();
				this.revisionsHashValue = val;
			} else if ("revisionsSaltValue" === reader.GetName()) {
				val = reader.GetValue();
				this.revisionsSaltValue = val;
			} else if ("revisionsSpinCount" === reader.GetName()) {
				val = reader.GetValueInt();
				this.revisionsSpinCount = val;
			} else if ("lockRevision" === reader.GetName()) {
				val = reader.GetValueBool();
				this.lockRevision = val;
			} else if ("lockStructure" === reader.GetName()) {
				val = reader.GetValueBool();
				this.lockStructure = val;
			} else if ("lockWindows" === reader.GetName()) {
				val = reader.GetValueBool();
				this.lockWindows = val;
			} else if ("workbookPassword" === reader.GetName()) {
				val = reader.GetValue();
				this.workbookPassword = val;
			}
		}
	};

	AscCommonExcel.Worksheet.prototype.fromXml = function (reader) {
		if (!reader.ReadNextNode()) {
			return;
		}
		if ("worksheet" !== reader.GetNameNoNS()) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}

		let t = this;
		let oVmlDrawingReader = null;
		if ("worksheet" === reader.GetNameNoNS()) {
			var context = reader.GetContext();
			context.aOleObjectsData = [];
			context.initFromWS(this);
			var depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				var name = reader.GetNameNoNS();
				if ("cols" === name) {
					var cols = new CT_Cols(this, context.InitOpenManager.aCellXfs);
					cols.fromXml(reader);
					var aTempCols = cols.val;

					context.InitOpenManager.prepareAfterReadCols(this, aTempCols);
				} else if ("dimension" === name) {
					//do not support serialize
				} else if ("drawing" === name) {
					var drawing = new AscCommonExcel.CT_DrawingWSRef();
					drawing.fromXml(reader);
					context.drawingId = drawing.id;
				} else if ("hyperlinks" === name) {
					var hyperlinks = [];
					var depth2 = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depth2)) {
						var name2 = reader.GetNameNoNS();
						if ("hyperlink" === name2) {
							var hyperlink = new AscCommonExcel.Hyperlink();
							hyperlink.fromXml(reader, this);
							hyperlinks.push(hyperlink);
						}
					}

					context.InitOpenManager.prepareAfterReadHyperlinks(hyperlinks);
				} else if ("mergeCells" === name) {
					var aMerged = getSimpleArrayFromXml(reader, "mergeCell", "ref");
					if (aMerged) {
						context.InitOpenManager.prepareAfterReadMergedCells(this, aMerged);
					}
				} else if ("pageMargins" === name) {
					//asc_CPageMargins
					if (this.PagePrintOptions && this.PagePrintOptions.pageMargins) {
						this.PagePrintOptions.pageMargins.fromXml(reader);
						this.PagePrintOptions.pageMargins.ws = this;
					}
				} else if ("pageSetup" === name) {
					//asc_CPageSetup
					if (this.PagePrintOptions && this.PagePrintOptions.pageSetup) {
						this.PagePrintOptions.pageSetup.fromXml(reader);
					}
				} else if ("printOptions" === name) {
					//asc_CPageOptions
					if (this.PagePrintOptions) {
						this.PagePrintOptions.fromXml(reader);
					}
				} else if ("sheetData" === name) {
					/*var sheetData = new AscCommonExcel.CT_SheetData(this);
					sheetData.fromXml(reader);*/
					context.InitOpenManager.oReadResult.sheetData.push({ws: this, reader: reader, state: reader.getState()});
				} else if ("WorksheetOptions" === name) {
					//do not support serialize
				} else if ("Names" === name) {
					//do not support serialize
					//def names
				} else if ("sheetPr" === name) {
					this.sheetPr = new AscCommonExcel.asc_CSheetPr();
					this.sheetPr.fromXml(reader);
				} else if ("autoFilter" === name) {
					var autoFilter = new AscCommonExcel.AutoFilter();
					autoFilter.fromXml(reader);
					this.AutoFilter = autoFilter;
				} else if ("legacyDrawing" === name) {
					//do not support serialize
					//need for comments
					readOneAttr(reader, "id", function () {
						context.InitOpenManager.legacyDrawingId = reader.GetValue();
					});
					let oRel = reader.rels.getRelationship(context.InitOpenManager.legacyDrawingId);
					let oRelPart = reader.rels.pkg.getPartByUri(oRel.targetFullName);
					let oContent = oRelPart.getDocumentContent();
					let oReader = new AscCommon.StaxParser(oContent, oRelPart, reader.context);
					let oElement = new AscFormat.CVMLDrawing();
					if (oElement) {
						oElement.fromXml(oReader, true);
						context.InitOpenManager.legacyDrawing = oElement;
					}
					oVmlDrawingReader = oReader;
				} else if ("legacyDrawingHF" === name) {
					//do not support serialize - commented
				} else if ("oleObjects" === name) {
					let oWS = this;
					let oDrawingData = null;
					let fReadOleDrawing = function (reader, name) {
						if(oDrawingData) {
							return true;
						}
						if(name === "AlternateContent") {
							fCallReadDrawing();
						} else if(name === "Choice") {
							fCallReadDrawing();
						} else if(name === "oleObject") {
							let oFrom = null, oTo = null, oPrNode = null;
							let oNode = new CT_XmlNode(function(reader, name) {
								if(name === "objectPr") {
									oPrNode = new CT_XmlNode(function(reader, name) {
										if(name === "anchor") {
											let oAnchorNode = new CT_XmlNode(function(reader, name){
												if(name === "from") {
													oFrom = new AscFormat.CCellObjectInfo();
													oFrom.fromXml(reader);
												}
												else if(name === "to") {
													oTo = new AscFormat.CCellObjectInfo();
													oTo.fromXml(reader);
												}
												return true;
											});
											oAnchorNode.fromXml(reader);
										}
										return true;
									});
									oPrNode.fromXml(reader);
								}
								return true;
							});
							oNode.fromXml(reader);
							oDrawingData = {drawingNode: oNode, from: oFrom, to: oTo, objectPr: oPrNode};
							context.aOleObjectsData.push(oDrawingData);
						}
					};
					let fCallReadDrawing = function() {
						let oNode = new CT_XmlNode(function (reader, name) {
							fReadOleDrawing(reader, name);
							return true;
						});
						oNode.fromXml(reader);
					};
					fCallReadDrawing();

				} else if ("controls" === name) {
					//do not support serialize
				} else if ("headerFooter" === name) {
					if (this.headerFooter) {
						this.headerFooter.fromXml(reader);
					}
				} else if ("dataValidations" === name) {
					var dataValidations = new AscCommonExcel.CDataValidations();
					dataValidations.fromXml(reader);
					this.dataValidations = dataValidations;
				} else if ("dataConsolidate" === name) {

				} else if ("conditionalFormatting" === name && typeof AscCommonExcel.CConditionalFormatting != "undefined") {
					var oConditionalFormatting = new AscCommonExcel.CConditionalFormatting();
					oConditionalFormatting.fromXml(reader);

					context.InitOpenManager.prepareConditionalFormatting(this, oConditionalFormatting);
				} else if ("sheetFormatPr" === name) {
					//SheetFormatPr
					if (this.oSheetFormatPr) {
						this.oSheetFormatPr.fromXml(reader, this);
					}
				} else if ("sheetViews" === name) {
					reader.readXmlArray("sheetView", function (index) {
						if (index === 0) {
							var sheetViewSettings = new AscCommonExcel.asc_CSheetViewSettings();
							sheetViewSettings.fromXml(reader, t);
							t.sheetViews.push(sheetViewSettings);
						}
					});
				} else if ("protectedRanges" === name) {
					reader.readXmlArray("protectedRange", function () {
						var oProtectedRange = Asc.CProtectedRange ? new Asc.CProtectedRange() : null;
						if (oProtectedRange) {
							oProtectedRange.fromXml(reader);
							t.aProtectedRanges.push(oProtectedRange);
						}
					});
				} else if ("sortState" === name) {
					var sortState = new AscCommonExcel.SortState();
					sortState.fromXml(reader);
					this.sortState = sortState;
				} else if ("extLst" === name) {
					var extLst = new COfficeArtExtensionList(this);
					extLst.fromXml(reader);
				} else if ("picture" === name) {
					//do not support serialize - commented
				} else if ("rowBreaks" === name) {
					//do not support serialize - commented
				} else if ("colBreaks" === name) {
					//do not support serialize - commented
				} else if ("sheetProtection" === name) {
					var sheetProtection = Asc.CSheetProtection ? new Asc.CSheetProtection(this) : null;
					if (sheetProtection) {
						sheetProtection.fromXml(reader);
						this.sheetProtection = sheetProtection;
					}
				}
			}
		}

		this.prepareExtLst(extLst, context.InitOpenManager);
		this.prepareLegacyDrawings(reader, oVmlDrawingReader)
	};
	AscCommonExcel.Worksheet.prototype.prepareLegacyDrawings = function(reader, oVmlDrawingReader) {
		let context = reader.context;
		let aOleObjectsData = context.aOleObjectsData;
		let oVmlDrawing = context.InitOpenManager.legacyDrawing;

		let objectRender = new AscFormat.DrawingObjects();
		let bVmlReader = false;


		let oOldReader;
		if(oVmlDrawingReader) {
			oOldReader = oVmlDrawingReader.context.reader;
			oVmlDrawingReader.context.reader = oVmlDrawingReader;
		}
		for(let nOle = 0; nOle < aOleObjectsData.length; ++nOle) {
			let oData = aOleObjectsData[nOle];
			let oFrom, oTo;
			let oDrawingNode = oData.drawingNode;
			let oPrNode = oData.objectPr;
			oFrom = oData.from;
			oTo = oData.to;
			let sShapeId = oDrawingNode.attributes["shapeId"];
			let oVMLSp = oVmlDrawing.getShape(sShapeId);
			let sImageRId = null;
			let sDataLinkId = oDrawingNode.attributes["id"];
			let sProgId = oDrawingNode.attributes["progId"];
			if(oVMLSp) {
				let oClientData = oVMLSp.getClientData();
				if(oClientData) {
					let sAnchor = oClientData.m_oAnchor;
					if(sAnchor) {
						let aCoords = sAnchor.split(",");
						if(aCoords.length === 8) {
							let nFromCol = parseInt(aCoords[0]);
							let dFromColOff = AscFormat.Px_To_Mm(parseInt(aCoords[1]));
							let nFromRow = parseInt(aCoords[2]);
							let dFromRowOff = AscFormat.Px_To_Mm(parseInt(aCoords[3]));
							let nToCol = parseInt(aCoords[4]);
							let dToColOff = AscFormat.Px_To_Mm(parseInt(aCoords[5]));
							let nToRow = parseInt(aCoords[6]);
							let dToRowOff = AscFormat.Px_To_Mm(parseInt(aCoords[7]));
							if(AscFormat.isRealNumber(nFromCol) && AscFormat.isRealNumber(dFromColOff) &&
								AscFormat.isRealNumber(nFromRow) && AscFormat.isRealNumber(dFromRowOff) &&
								AscFormat.isRealNumber(nToCol) && AscFormat.isRealNumber(dToColOff) &&
								AscFormat.isRealNumber(nToRow) && AscFormat.isRealNumber(dToRowOff)) {
								if(!oFrom) {
									oFrom = new AscFormat.CCellObjectInfo();
								}
								oFrom.col = nFromCol;
								oFrom.colOff = dFromColOff;
								oFrom.row = nFromRow;
								oFrom.rowOff = dFromRowOff;
								if(!oTo) {
									oTo = new AscFormat.CCellObjectInfo();
								}
								oTo.col = nToCol;
								oTo.colOff = dToColOff;
								oTo.row = nToRow;
								oTo.rowOff = dToRowOff;
							}
						}
					}
				}
				let oImageData = oVMLSp.getImageData();
				if(oImageData) {
					if (oImageData.m_oRelId) sImageRId = oImageData.m_oRelId;
					else if (oImageData.m_rId) sImageRId = oImageData.m_rId;
					else if (oImageData.m_rPict) sImageRId = oImageData.m_rPict;
					if(sImageRId) {
						bVmlReader = !!oVmlDrawingReader;
					}
				}
			}
			if(!sImageRId) {
				if(oPrNode) {
					sImageRId = oPrNode.attributes["id"];
					bVmlReader = false;
				}
			}
			if(oFrom && oTo) {
				let oOleObject = new AscFormat.COleObject();
				AscFormat.fillImage(oOleObject, "", 0, 0, 50, 50);
				AscFormat.fReadXmlRasterImageId(bVmlReader ? oVmlDrawingReader : reader, sImageRId, oOleObject.blipFill);
				oOleObject.fillDataLink(sDataLinkId, reader);
				oOleObject.setApplicationId(sProgId);
				let oDrawing = objectRender.createDrawingObject(AscCommon.c_oAscCellAnchorType.cellanchorTwoCell);
				oDrawing.graphicObject = oOleObject;
				oDrawing.from = oFrom;
				oDrawing.to = oTo;
				oDrawing.initAfterSerialize(this);
			}
		}
		//signature lines
		if(oVmlDrawing) {
			let aSL = oVmlDrawing.getSignatureLines();
			let aOOXMLSl = [];
			for(let nSL = 0; nSL < aSL.length; ++nSL) {
				let oSL = aSL[nSL];
				let oOOXMLSL = oSL.convertToOOXML(oVmlDrawing.items, null, oVmlDrawingReader.context);
				let oFrom, oTo;
				let oClientData = oSL.getClientData();
				if(oClientData) {
					let sAnchor = oClientData.m_oAnchor;
					if(sAnchor) {
						let aCoords = sAnchor.split(",");
						if(aCoords.length === 8) {
							let nFromCol = parseInt(aCoords[0]);
							let dFromColOff = AscFormat.Px_To_Mm(parseInt(aCoords[1]));
							let nFromRow = parseInt(aCoords[2]);
							let dFromRowOff = AscFormat.Px_To_Mm(parseInt(aCoords[3]));
							let nToCol = parseInt(aCoords[4]);
							let dToColOff = AscFormat.Px_To_Mm(parseInt(aCoords[5]));
							let nToRow = parseInt(aCoords[6]);
							let dToRowOff = AscFormat.Px_To_Mm(parseInt(aCoords[7]));
							if(AscFormat.isRealNumber(nFromCol) && AscFormat.isRealNumber(dFromColOff) &&
								AscFormat.isRealNumber(nFromRow) && AscFormat.isRealNumber(dFromRowOff) &&
								AscFormat.isRealNumber(nToCol) && AscFormat.isRealNumber(dToColOff) &&
								AscFormat.isRealNumber(nToRow) && AscFormat.isRealNumber(dToRowOff)) {
								oFrom = new AscFormat.CCellObjectInfo();
								oFrom.col = nFromCol;
								oFrom.colOff = dFromColOff;
								oFrom.row = nFromRow;
								oFrom.rowOff = dFromRowOff;
								oTo = new AscFormat.CCellObjectInfo();
								oTo.col = nToCol;
								oTo.colOff = dToColOff;
								oTo.row = nToRow;
								oTo.rowOff = dToRowOff;
							}
						}
					}
				}
				if(oOOXMLSL && oOOXMLSL.isSignatureLine() && oFrom && oTo) {
					let oDrawing = objectRender.createDrawingObject(AscCommon.c_oAscCellAnchorType.cellanchorTwoCell);
					let oSpPr = oOOXMLSL.spPr;
					oOOXMLSL = oOOXMLSL.convertToPPTX(this.getDrawingDocument(), this, false);
					oOOXMLSL.setSpPr(oSpPr);
					oDrawing.graphicObject = oOOXMLSL;
					oDrawing.from = oFrom;
					oDrawing.to = oTo;
					oDrawing.initAfterSerialize(this);
				}
			}
			return aOOXMLSl;
		}

		if(oVmlDrawingReader) {
			oVmlDrawingReader.context.reader = oOldReader;
		}
	};
	//TOoDO PrepareToWrite делается в x2t, здесь не вижу необходиимости, но проверить нужно
	AscCommonExcel.Worksheet.prototype.toXml = function (writer) {
		var t = this, i;
		var context = writer.context;
		context.ws = this;

		writer.WriteXmlString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
		writer.WriteXmlNodeStart("worksheet");

		writer.WriteXmlString(' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
		writer.WriteXmlString(' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
		writer.WriteXmlString(' xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"');
		writer.WriteXmlString(' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"');
		writer.WriteXmlString(' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"');
		writer.WriteXmlString(' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"');
		writer.WriteXmlString(' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"');
		writer.WriteXmlString(' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"');
		writer.WriteXmlString(' xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" mc:Ignorable="x14ac"');

		writer.WriteXmlAttributesEnd();

		//TOoDO dimension ?
		//writer.WriteXmlString('<dimension ref="A1"/>');

		//AscCommonExcel.asc_CSheetPr
		if (this.sheetPr) {
			this.sheetPr.toXml(writer, "sheetPr");
		}
		//AscCommonExcel.asc_CSheetViewSettings
		if (this.sheetViews) {
			writer.WriteXmlArray(this.sheetViews, "sheetView", "sheetViews");
		}
		//AscCommonExcel.SheetFormatPr
		if (this.oSheetFormatPr) {
			this.oSheetFormatPr.toXml(writer, "sheetFormatPr");
		}

		//WriteWorksheetCols
		var colsToWrite = new CT_Cols(this);
		context.InitSaveManager.writeCols(this, context.stylesForWrite, function (oColToWrite) {
			colsToWrite.val.push(oColToWrite);
		});
		if (colsToWrite.val.length) {
			colsToWrite.toXml(writer, "cols");
		}

		//sheetData
		writer.WriteXmlNodeStart('sheetData');
		writer.WriteXmlAttributesEnd();
		this.toXmlSheetData(writer);
		writer.WriteXmlNodeEnd("sheetData");

		//Asc.CSheetProtection
		if (this.sheetProtection) {
			this.sheetProtection.toXml(writer, "sheetProtection");
		}
		//Asc.CProtectedRange
		if (this.aProtectedRanges) {
			writer.WriteXmlArray(this.aProtectedRanges, "protectedRange", "protectedRanges");
		}
		//AscCommonExcel.AutoFilter
		if (this.AutoFilter) {
			this.AutoFilter.toXml(writer, "autoFilter");
		}
		//AscCommonExcel.SortState()
		if (this.sortState) {
			this.sortState.toXml(writer, "sortState");
		}

		/*if (this.DataConsolidate) {
			this.DataConsolidate.toXml(writer);
		}*/

		var mergedArr = [];
		context.InitSaveManager.WriteMergeCells(this, function (ref) {
			mergedArr.push(ref);
		});
		if (mergedArr && mergedArr.length) {
			writer.WriteXmlString("<mergeCells");
			writer.WriteXmlNullableAttributeNumber("count", mergedArr.length);
			writer.WriteXmlString(">");

			for (i = 0; i < mergedArr.length; ++i) {
				if (mergedArr[i]) {
					writer.WriteXmlString("<mergeCell");
					writer.WriteXmlNullableAttributeStringEncode("ref", mergedArr[i]);
					writer.WriteXmlString("/>");
				}
			}
			writer.WriteXmlString("</mergeCells>");
		}

		var aConditionalFormattingExt = [];
		if (this.aConditionalFormattingRules) {
			for (i = 0; i < this.aConditionalFormattingRules.length; i++) {
				var rule = this.aConditionalFormattingRules[i];
				if (rule.isExtended()) {
					aConditionalFormattingExt.push(rule);
				} else {
					var oConditionalFormatting = new AscCommonExcel.CConditionalFormatting();
					oConditionalFormatting.aRules = [rule];
					oConditionalFormatting.toXml(writer)
				}
			}
		}

		//AscCommonExcel.CDataValidations
		if (this.dataValidations) {
			this.dataValidations.toXml(writer);
		}

		var oHyperlinks = this.hyperlinkManager.getAll();
		if (oHyperlinks && oHyperlinks.length) {
			writer.WriteXmlString("<hyperlinks>");

			for (i in oHyperlinks) {
				var elem = oHyperlinks[i];

				//write only active hyperlink, if copy/paste
				if (!context.InitSaveManager.isCopyPaste ||
					(context.InitSaveManager.isCopyPaste && elem && elem.bbox && context.InitSaveManager.isCopyPaste.containsRange(elem.bbox))) {

					if (elem.data.Hyperlink) {
						elem.data.rId = context.part.addRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", "ya.ru", "External");
					}

					elem.data.toXml(writer, "hyperlink");
				}
			}

			writer.WriteXmlString("</hyperlinks>");
		}


		//asc_CPageOptions
		if (this.PagePrintOptions) {
			this.PagePrintOptions.toXml(writer, "printOptions");
		}
		//asc_CPageMargins
		if (this.PagePrintOptions && this.PagePrintOptions.pageMargins) {
			this.PagePrintOptions.pageMargins.toXml(writer, "pageMargins");
		}
		//asc_CPageSetup
		if (this.PagePrintOptions && this.PagePrintOptions.pageSetup) {
			this.PagePrintOptions.pageSetup.toXml(writer, "pageSetup");
		}
		//Asc.CHeaderFooter
		if (this.headerFooter) {
			this.headerFooter.toXml(writer, "headerFooter");
		}

		//skip m_oRowBreaks
		//skip m_oColBreaks

		//drawings
		if (this.Drawings.length > 0) {
			var drawing = new AscCommonExcel.CT_DrawingWS(t);
			var drawingPart = context.part.addPart(AscCommon.openXml.Types.drawings);
			drawingPart.part.setDataXml(drawing, writer);
			var drawingRef = new AscCommonExcel.CT_DrawingWSRef();
			drawingRef.id = drawingPart.rId;
			drawingRef.toXml(writer, "drawing");
		}

		let vmldrawingPart = null;
		let vmlDrawingMemory = null;
		let bVmlDrawing = (context.oleDrawings.length > 0 ||
			context.signatureDrawings.length > 0 ||
			this.aComments.length > 0);
		if(bVmlDrawing) {
			vmldrawingPart = context.part.addPart(AscCommon.openXml.Types.vmlDrawing);
			vmlDrawingMemory = new AscCommon.CMemory();
			vmlDrawingMemory.context = context;
			let vmldrawingRef = new AscCommonExcel.CT_DrawingWSRef();
			vmldrawingRef.id = vmldrawingPart.rId;
			vmldrawingRef.toXml(writer, "legacyDrawing");
			vmlDrawingMemory.WriteXmlString("<xml \
xmlns:v=\"urn:schemas-microsoft-com:vml\" \
xmlns:o=\"urn:schemas-microsoft-com:office:office\" \
xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
		}

		//vml drawings
		if (this.aComments.length > 0) {
			//TODO m_sGfxdata - не протаскивается
			var vmldrawing = new AscFormat.CVMLDrawing();
			vmldrawing.m_mapComments = this.aComments;
			vmlDrawingMemory.WriteXmlString(vmldrawing.getXmlString());

		}

		//skip m_oLegacyDrawingHF
		//skip m_oPicture
		//OLEObjects
		if(context.oleDrawings.length > 0) {
			writer.WriteXmlNodeStart("oleObjects");
			writer.WriteXmlAttributesEnd();
			let aDrawings = context.oleDrawings;
			for(let nDrawing = 0; nDrawing < aDrawings.length; ++nDrawing) {
				let oDrawing = aDrawings[nDrawing];
				oDrawing.toXmlOle(writer);
			}
			writer.WriteXmlNodeEnd("oleObjects");
		}

		if(vmlDrawingMemory) {
			vmlDrawingMemory.context.clearCurrentPartDataMaps();
		}
		if(context.oleDrawings.length > 0) {
			let oldPart = vmlDrawingMemory.context.part;
			vmlDrawingMemory.context.part = vmldrawingPart.part;
			let aDrawings = context.oleDrawings;
			for(let nDrawing = 0; nDrawing < aDrawings.length; ++nDrawing) {
				let oDrawing = aDrawings[nDrawing];
				oDrawing.graphicObject.toXmlVML(vmlDrawingMemory, "", "", "", "_x0000_s" + oDrawing.nShapeId)
			}
			vmlDrawingMemory.context.part = oldPart;
		}
		if(context.signatureDrawings.length > 0) {
			let oldPart = vmlDrawingMemory.context.part;
			vmlDrawingMemory.context.part = vmldrawingPart.part;
			let aDrawings = context.signatureDrawings;
			for(let nDrawing = 0; nDrawing < aDrawings.length; ++nDrawing) {
				let oDrawing = aDrawings[nDrawing];
				oDrawing.graphicObject.toXmlVML(vmlDrawingMemory, "", "", "")
			}
			vmlDrawingMemory.context.part = oldPart;
		}

		if(bVmlDrawing) {
			vmlDrawingMemory.WriteXmlString("</xml>");
			let vmlData = vmlDrawingMemory.GetDataUint8();
			vmldrawingPart.part.setData(vmlData);
		}
		//skip m_oControls

		if (this.TableParts && this.TableParts.length > 0) {

			writer.WriteXmlNodeStart("tableParts");
			writer.WriteXmlAttributeNumber("count", this.TableParts.length);
			writer.WriteXmlAttributesEnd();

			for (i = 0; i < this.TableParts.length; i++) {
				var tablePart = context.part.addPart(AscCommon.openXml.Types.tableDefinition);
				tablePart.part.setDataXml(this.TableParts[i], writer);
				//CT_DrawingWSRef - внутри только id, поэтому здесь использую. либо переимонвать, либо базовый сделать с такими аттрибутами
				var tableRef = new AscCommonExcel.CT_DrawingWSRef();
				tableRef.id = tablePart.rId;
				tableRef.toXml(writer, "tablePart");
			}

			writer.WriteXmlNodeEnd("tableParts");
		}

		var oComments = prepareCommentsToWrite(this.aComments, context.InitSaveManager.personList);
		if (oComments) {
			if (oComments.comments) {
				var commentsPart = context.part.addPart(AscCommon.openXml.Types.worksheetComments);
				commentsPart.part.setDataXml(oComments.comments, writer);
			}
			if (oComments.threadedComments) {
				var threadedCommentsPart = context.part.addPart(AscCommon.openXml.Types.threadedComment);
				threadedCommentsPart.part.setDataXml(oComments.threadedComments, writer);
			}
		}

		if (this.aNamedSheetViews && this.aNamedSheetViews.length > 0) {
			var namedSheetViews = new Asc.CT_NamedSheetViews();
			namedSheetViews.namedSheetView = this.aNamedSheetViews;
			var namedSheetViewsPart = context.part.addPart(AscCommon.openXml.Types.namedSheetViews);
			namedSheetViewsPart.part.setDataXml(namedSheetViews, writer);
		}

		var officeArtExtensionList = new COfficeArtExtensionList(this);
		var officeArtExtension;
		if (aConditionalFormattingExt && aConditionalFormattingExt.length) {
			officeArtExtension = new COfficeArtExtension(this);
			officeArtExtension.uri = extUri.conditionalFormattings;
			officeArtExtension.aConditionalFormattingRules = aConditionalFormattingExt;
			officeArtExtensionList.arrExt.push(officeArtExtension);
		}
		if (this.dataValidations) {
			officeArtExtension = new COfficeArtExtension(this);
			officeArtExtension.uri = extUri.dataValidations;
			officeArtExtension.dataValidations = this.dataValidations;
			officeArtExtensionList.arrExt.push(officeArtExtension);
		}
		if (this.aSparklineGroups && this.aSparklineGroups.length) {
			officeArtExtension = new COfficeArtExtension(this);
			officeArtExtension.uri = extUri.sparklineGroups;
			officeArtExtension.sparklineGroups = this.aSparklineGroups;
			officeArtExtensionList.arrExt.push(officeArtExtension);
		}


		var slicers = new Asc.CT_slicers();
		var slicerExt = new Asc.CT_slicers();
		for (i = 0; i < this.aSlicers.length; ++i) {
			/*if (this.isCopyPaste) {
				var _graphicObject = ws.workbook.getSlicerViewByName(ws.aSlicers[i].name);
				if (!_graphicObject || !_graphicObject.selected) {
					continue;
				}
			}*/

			if (this.aSlicers[i].isExt()) {
				slicerExt.slicer.push(this.aSlicers[i]);
			} else {
				slicers.slicer.push(this.aSlicers[i]);
			}
		}

		var doSlicer = function (_slicer, bExt) {
			if (_slicer && _slicer.slicer.length) {
				var sliceListIds = [];
				var slicerPart = writer.context.part.addPart(AscCommon.openXml.Types.slicers);
				slicerPart.part.setDataXml(_slicer, writer);
				sliceListIds.push(slicerPart.rId);

				//ссылки на slicer кладём в ws в extLst
				officeArtExtension = new COfficeArtExtension();
				officeArtExtension.uri = bExt ? extUri.slicerListExt : extUri.slicerList;
				officeArtExtension.additionalNamespace = bExt ? additionalNamespace.slicerListExt : additionalNamespace.slicerList;
				if (bExt) {
					officeArtExtension.slicerListExtIds = sliceListIds;
				} else {
					officeArtExtension.slicerListIds = sliceListIds;
				}
				officeArtExtensionList.arrExt.push(officeArtExtension);

			}
		}
		doSlicer(slicers);
		doSlicer(slicerExt, true);

		officeArtExtensionList.toXml(writer);

		writer.WriteXmlNodeEnd("worksheet");

		for (i = 0; i < this.pivotTables.length; ++i) {
			let pivotTable = this.pivotTables[i];
			let pivotTablePart = context.part.addPart(AscCommon.openXml.Types.pivotTable);
			pivotTablePart.part.setDataXml(this.pivotTables[i], writer);
			if (pivotTable.cacheDefinition) {
				let uri = writer.context.oUriMap[pivotTable.cacheDefinition.Get_Id()];
				if (uri) {
					let type = AscCommon.openXml.Types.pivotTableCacheDefinition;
					let name = uri.substring(uri.lastIndexOf('/') + 1);
					pivotTablePart.part.addRelationship(type.relationType, type.dir + "/" + name);
				}
			}
		}

		context.ws = null;
	};

	Asc.asc_CDefName.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name /*definedName*/);
		writer.WriteXmlNullableAttributeStringEncode("name", this.Name);
		writer.WriteXmlNullableAttributeNumber("localSheetId", this.LocalSheetId);
		writer.WriteXmlNullableAttributeBool("hidden", this.Hidden);
		writer.WriteXmlAttributesEnd();
		if (this.Ref) {
			writer.WriteXmlStringEncode(this.Ref);
		}
		writer.WriteXmlNodeEnd(ns + name);
	};

	Asc.asc_CDefName.prototype.fromXml = function (reader) {

		/*ReadAttributes( oReader );

						if ( oReader.IsEmptyNode() )
							return;

						m_oRef = oReader.GetText3();*/

		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		this.Ref = reader.GetText();
	};

	Asc.asc_CDefName.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			/*if ("comment" === reader.GetName()) {
				val = reader.GetValue();
				this.comment = val;
			} else if ("customMenu" === reader.GetName()) {
				val = reader.GetValue();
				this.customMenu = val;
			} else if ("description" === reader.GetName()) {
				val = reader.GetValue();
				this.description = val;
			} else if ("function" === reader.GetName()) {
				val = reader.GetValue();
				this.function = val;
			} else if ("functionGroupId" === reader.GetName()) {
				val = reader.GetValue();
				this.functionGroupId = val;
			} else if ("help" === reader.GetName()) {
				val = reader.GetValue();
				this.help = val;
			} else*/
			if ("hidden" === reader.GetName()) {
				val = reader.GetValueBool();
				this.Hidden = val;
			} else if ("localSheetId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.LocalSheetId = val;
			} else if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.Name = val;
			} /*else if ("publishToServer" === reader.GetName()) {
				val = reader.GetValue();
				this.PublishToServer = val;
			} else if ("shortcutKey " === reader.GetName()) {
				val = reader.GetValue();
				this.ShortcutKey  = val;
			} else if ("statusBar" === reader.GetName()) {
				val = reader.GetValue();
				this.StatusBar = val;
			} else if ("vbProcedure" === reader.GetName()) {
				val = reader.GetValue();
				this.VbProcedure = val;
			} else if ("workbookParameter" === reader.GetName()) {
				val = reader.GetValue();
				this.WorkbookParameter = val;
			} else if ("xlm" === reader.GetName()) {
				val = reader.GetValue();
				this.Xlm = val;
			} else if ("ss:Name" === reader.GetName()) {
				val = reader.GetValue();
				this.Ss:Name = val;
			} else if ("ss:RefersTo" === reader.GetName()) {
				val = reader.GetValue();
				this.Ss:RefersTo = val;
			}*/
		}
	};


	function CT_Cols(ws, aCellXfs) {
		this.val = [];
		this.ws = ws;
		this.aCellXfs = aCellXfs;
	}

	CT_Cols.prototype.fromXml = function (reader) {
		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("col" === name) {

				/*WritingStringNullableAttrBool(L"bestFit", m_oBestFit);
				WritingStringNullableAttrBool(L"collapsed", m_oCollapsed);
				WritingStringNullableAttrBool(L"customWidth", m_oCustomWidth);
				WritingStringNullableAttrBool(L"hidden", m_oHidden);
				WritingStringNullableAttrInt(L"min", m_oMin, m_oMin->GetValue());
				WritingStringNullableAttrInt(L"max", m_oMax, m_oMax->GetValue());
				WritingStringNullableAttrInt(L"outlineLevel", m_oOutlineLevel, m_oOutlineLevel->GetValue());
				WritingStringNullableAttrBool(L"phonetic", m_oPhonetic);
				WritingStringNullableAttrInt(L"style", m_oStyle, m_oStyle->GetValue());
				WritingStringNullableAttrDouble(L"width", m_oWidth, m_oWidth->GetValue());*/


				/*if ( c_oSerWorksheetColTypes.BestFit == type )
					oTempCol.col.BestFit = this.stream.GetBool();
				else if ( c_oSerWorksheetColTypes.Hidden == type )
					oTempCol.col.setHidden(this.stream.GetBool());
				else if ( c_oSerWorksheetColTypes.Max == type )
					oTempCol.Max = this.stream.GetULongLE();
				else if ( c_oSerWorksheetColTypes.Min == type )
					oTempCol.Min = this.stream.GetULongLE();
				else if (c_oSerWorksheetColTypes.Style == type) {
					var xfs = aCellXfs[this.stream.GetULongLE()];
					if (xfs) {
						oTempCol.col.setStyle(xfs);
					}
				} else if ( c_oSerWorksheetColTypes.Width == type )
					oTempCol.col.width = this.stream.GetDoubleLE();
				else if ( c_oSerWorksheetColTypes.CustomWidth == type )
					oTempCol.col.CustomWidth = this.stream.GetBool();
				else if ( c_oSerWorksheetColTypes.OutLevel == type )
					oTempCol.col.outlineLevel = this.stream.GetULongLE();
				else if ( c_oSerWorksheetColTypes.Collapsed == type )
					oTempCol.col.collapsed = this.stream.GetBool();*/


				var ptWidth, bAutoFit, sStyleID;
				var oTempCol = {Max: null, Min: null, col: new AscCommonExcel.Col(this.ws, 0)};
				var val;
				while (reader.MoveToNextAttribute()) {
					if ("bestFit" === reader.GetName()) {
						val = reader.GetValueBool();
						oTempCol.col.BestFit = val;
					} else if ("collapsed" === reader.GetName()) {
						val = reader.GetValueBool();
						oTempCol.col.collapsed = val;
					} else if ("customWidth" === reader.GetName()) {
						val = reader.GetValueBool();
						oTempCol.col.CustomWidth = val;
					} else if ("hidden" === reader.GetName()) {
						val = reader.GetValueBool();
						oTempCol.col.setHidden(val);
					} else if ("min" === reader.GetName()) {
						val = reader.GetValueInt();
						oTempCol.Min = val;
					} else if ("max" === reader.GetName()) {
						val = reader.GetValueInt();
						oTempCol.Max = val;
					} else if ("outlineLevel" === reader.GetName()) {
						val = reader.GetValueInt();
						oTempCol.col.outlineLevel = val;
					} /*else if ("phonetic" === reader.GetName()) {
						val = reader.GetValueBool();
						oTempCol.col.phonetic = val;
					}*/ else if ("style" === reader.GetName()) {
						var xfs = this.aCellXfs[reader.GetValueInt()];
						if (xfs) {
							oTempCol.col.setStyle(xfs);
						}
					} else if ("width" === reader.GetName()) {
						//val = reader.GetValueDouble();
						//мс себя так ведёт - всё что не число, зануляет
						val = reader.GetValue();
						if (AscCommon.isNumber(val)) {
							oTempCol.col.width = parseFloat(val);
						} else {
							oTempCol.col.width = 0;
						}
					} else if ("ss:Width" === reader.GetName()) {
						ptWidth = reader.GetValueInt();
					} else if ("ss:AutoFitWidth" === reader.GetName()) {
						bAutoFit = reader.GetValueBool();
					} else if ("ss:StyleID" === reader.GetName()) {
						sStyleID = reader.GetValue();
					}
				}

				//добавляю обработку, аналогичную той, которая есть в x2t
				if (ptWidth) {
					var pixDpi = ptWidth / 72 * 96;
					if (pixDpi < 5) {
						pixDpi = 7;
					}
					var maxDigitSize = 4.25;
					oTempCol.col.width = parseInt((pixDpi - 5) / maxDigitSize * 100. + 0.5) / 100 * 0.9;
					oTempCol.col.CustomWidth = true;
				}

				/*if (bAutoFit && bAutoFit === false) {

				} else if (xlsx_flat) {
					oTempCol.col.BestFit = true;

					if (false === ptWidth) {
						oTempCol.col.CustomWidth = true;
						oTempCol.col.width = 9;//?
					}
				}*/

				this.val.push(oTempCol);
			}
		}
	};


	CT_Cols.prototype.toXml = function (writer, name, ns) {
		if (!this.val.length) {
			return;
		}
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlAttributesEnd();

		for (var i = 0; i < this.val.length; i++) {
			var oTmpCol = this.val[i];
			var oCol = oTmpCol.col;

			writer.WriteXmlNodeStart("col");
			writer.WriteXmlNullableAttributeBool("bestFit", oCol.BestFit);
			writer.WriteXmlNullableAttributeBool("collapsed", oCol.collapsed ? oCol.collapsed : null);
			writer.WriteXmlNullableAttributeBool("customWidth", oCol.CustomWidth);
			writer.WriteXmlNullableAttributeBool("hidden", oCol.hd ? oCol.hd : null);
			writer.WriteXmlNullableAttributeNumber("min", oTmpCol.Min);
			writer.WriteXmlNullableAttributeNumber("max", oTmpCol.Max);
			writer.WriteXmlNullableAttributeNumber("outlineLevel", oCol.outlineLevel > 0 ? oCol.outlineLevel : null);
			//writer.WriteXmlNullableAttributeBool("phonetic", this.phonetic);
			writer.WriteXmlNullableAttributeNumber("style", oTmpCol.xfsid);
			writer.WriteXmlNullableAttributeDouble("width", oTmpCol.width);
			writer.WriteXmlAttributesEnd(true);
		}

		writer.WriteXmlNodeEnd(ns + name);
	};


	AscCommonExcel.Worksheet.prototype.toXmlSheetData = function (writer) {
		var ws = this;
		var context = writer.context;
		var isCopyPaste = context.isCopyPaste;
		var range;
		if (isCopyPaste) {
			range = ws.getRange3(isCopyPaste.r1, isCopyPaste.c1, isCopyPaste.r2, isCopyPaste.c2);
		} else {
			range = ws.getRange3(0, 0, AscCommon.gc_nMaxRow0, AscCommon.gc_nMaxCol0);
		}
		var isFirstRow = true;
		var writeEndRow = function () {
			if (isFirstRow) {
				isFirstRow = false;
			} else {
				writer.WriteXmlNodeEnd("row");
			}
		};


		var curRow = -1;
		var allRow = ws.getAllRowNoEmpty();
		var tempRow = new AscCommonExcel.Row(ws);
		if (allRow) {
			tempRow.copyFrom(allRow);
		}
		range._foreachRowNoEmpty(function (row, excludedCount) {
			writeEndRow();
			row.toXml(writer, "row");
			curRow = row.getIndex();
		}, function (cell, nRow0, nCol0, nRowStart0, nColStart0, excludedCount) {
			if (curRow !== nRow0) {
				tempRow.setIndex(nRow0);
				writeEndRow();
				tempRow.toXml(writer, "row");
				curRow = tempRow.getIndex();
			}
			//сохраняем как и Excel даже пустой стиль(нужно чтобы убрать стиль строки/колонки)
			if (null != cell.xfs || false === cell.isNullText()) {
				cell.toXml(writer, "c");
			}
		}, (ws.bExcludeHiddenRows && isCopyPaste));
		writeEndRow();
	};

	AscCommonExcel.Worksheet.prototype.prepareExtLst = function (extLst, InitOpenManager) {
		if (extLst) {
			this._prepareConditionalFormatting(extLst, InitOpenManager);
			this._prepareDataValidations(extLst);
			this._prepareSparklineGroups(extLst);
		}
	}
	AscCommonExcel.Worksheet.prototype._prepareConditionalFormatting = function (extLst, InitOpenManager) {
		//добавляем из ext
		if (extLst) {
			var sheetRules = this.aConditionalFormattingRules;
			var getSheetCf = function (openId) {
				for (var n = 0; n < sheetRules.length; n++) {
					if (sheetRules[n]._openId === openId) {
						return sheetRules[n];
					}
				}
			};

			for (var i = 0; i < extLst.arrExt.length; i++) {
				if (extLst.arrExt[i] && extLst.arrExt[i].aConditionalFormattingRules) {
					for (var j = 0; j < extLst.arrExt[i].aConditionalFormattingRules.length; j++) {
						var extConditionalFormatting = extLst.arrExt[i].aConditionalFormattingRules[j];
						var extRule = extConditionalFormatting.aRules[0];
						//далее смотрим по id, если такое правило на листе
						var sheetRule = getSheetCf(extRule ? extRule._openId : null);
						if (sheetRule) {
							//мержим
							sheetRule.merge(extRule);
						} else {
							//добавляем
							InitOpenManager.prepareConditionalFormatting(this, extConditionalFormatting);
						}
					}
				}
			}
		}
	}
	AscCommonExcel.Worksheet.prototype._prepareDataValidations = function (extLst) {
		if (extLst) {
			for (var i = 0; i < extLst.arrExt.length; i++) {
				if (extLst.arrExt[i] && extLst.arrExt[i].dataValidations) {
					if (extLst.arrExt[i].dataValidations.elems) {
						//for (var j = 0; j < extLst.arrExt[i].dataValidations.elems.length; j++) {
							//if (extLst.arrExt[i].dataValidations.elems[j]) {
								if (this.dataValidations) {
									if (!this.dataValidations.disablePrompts) {
										this.dataValidations.disablePrompts = extLst.arrExt[i].dataValidations.disablePrompts;
										this.dataValidations.xWindow = extLst.arrExt[i].dataValidations.xWindow;
										this.dataValidations.yWindow = extLst.arrExt[i].dataValidations.yWindow;
									}
									this.dataValidations.elems = this.dataValidations.elems.concat(extLst.arrExt[i].dataValidations.elems);
								} else {
									this.dataValidations = extLst.arrExt[i].dataValidations;
								}
							//}
						//}
					}
				}
			}
		}
	};

	AscCommonExcel.Worksheet.prototype._prepareSparklineGroups = function (extLst) {
		if (extLst) {
			for (var i = 0; i < extLst.arrExt.length; i++) {
				if (extLst.arrExt[i] && extLst.arrExt[i].sparklineGroups) {
					//обрабатываю только ситуацию, когад 1 элумент. несколько элементов не встречал, но нужно будет проверить и обработать.
					for (var j = 0; j < extLst.arrExt[i].sparklineGroups.length; j++) {
						var newSparklineGroup = extLst.arrExt[i].sparklineGroups[j];
						newSparklineGroup.setWorksheet(this);
						this.aSparklineGroups.push(newSparklineGroup);
					}
				}
			}
		}
	};

	AscCommonExcel.Row.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		var context = writer.context;
		var s = context.stylesForWrite.add(this.xfs) || null;
		var outlineLevel = this.outlineLevel > 0 ? this.outlineLevel : null;

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlAttributeNumber("r", this.index + 1);
		writer.WriteXmlNullableAttributeNumber("s", s);
		if (s != null) {
			writer.WriteXmlNullableAttributeNumber("customFormat", 1);
		}
		writer.WriteXmlNullableAttributeNumber("ht", this.h);
		writer.WriteXmlAttributeBoolIfTrue("hidden", this.getHidden());
		writer.WriteXmlAttributeBoolIfTrue("customHeight", this.getCustomHeight());
		writer.WriteXmlNullableAttributeNumber("outlineLevel", outlineLevel);
		writer.WriteXmlAttributeBoolIfTrue("collapsed", this.getCollapsed());
		writer.WriteXmlAttributesEnd();
	};

	AscCommonExcel.Row.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			if ("c" === reader.GetName()) {
				this._tempCell.clear();
				this._tempCell.fromXml(reader);
				this._tempCell.saveContent();
			}
		}
	};

	AscCommonExcel.Row.prototype.fromXml2 = function (reader) {
		this.readAttr(reader);

		var tmp = reader.GetContext().InitOpenManager.tmp;
		if (tmp) {
			if (null === this.index) {
				this.index = tmp.prevRow + 1;
			}
			this.saveContent();

			this.ws.cellsByColRowsCount = Math.max(this.ws.cellsByColRowsCount, this.index + 1);
			this.ws.nRowsCount = Math.max(this.ws.nRowsCount, this.ws.cellsByColRowsCount);
			tmp.prevRow = this.index;
			tmp.prevCol = -1;
		}


		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			if ("c" === reader.GetName()) {


				tmp.cell.clear();
				tmp.formula.clean();

				//AscCommonExcel.Cell -> fromXml
				tmp.cell.fromXml(reader);

				if (tmp.cell.isNullTextString()) {
					//set default value in case of empty cell value
					tmp.cell.setTypeInternal(CellValueType.Number);
				}
				if (tmp.cell.hasRowCol()) {
					tmp.prevCol = tmp.cell.nCol;
				} else {
					tmp.prevCol++;
					tmp.cell.setRowCol(tmp.prevRow, tmp.prevCol);
				}

				reader.GetContext().InitOpenManager.initCellAfterRead(tmp);
			}
		}
	};

	AscCommonExcel.Row.prototype.readAttr = function (reader) {

		var val;
		while (reader.MoveToNextAttribute()) {
			if ("r" === reader.GetName() || "ss:Index" === reader.GetName()) {
				val = reader.GetValueInt() - 1;
				this.setIndex(val);
			} else if ("s" === reader.GetName()) {
				val = reader.GetValueInt();
				var aCellXfs = reader.context.InitOpenManager && reader.context.InitOpenManager.aCellXfs;
				if (aCellXfs) {
					var xfs = aCellXfs[val];
					if (xfs) {
						this.setStyle(xfs);
					}
				}
			} else if ("customFormat" === reader.GetName()) {
			} else if ("ht" === reader.GetName()) {
				val = reader.GetValue();
				//если не число - мс пропускает
				if (AscCommon.isNumber(val)) {
					this.setHeight(val);
				}
			} else if ("ss:Height" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.setHeight(val);
				this.setCustomHeight(true);
			} else if ("hidden" === reader.GetName()) {
				val = reader.GetValueBool();
				if (val) {
					this.setHidden(true);
				}
			} else if ("customHeight" === reader.GetName()) {
				val = reader.GetValueBool();
				if (val) {
					this.setCustomHeight(true);
				}
			} else if ("outlineLevel" === reader.GetName()) {
				val = reader.GetValueInt();
				this.setOutlineLevel(val);
			} else if ("collapsed" === reader.GetName()) {
				val = reader.GetValueBool();
				this.setCollapsed(val);
			} else if ("x14ac:dyDescent" === reader.GetName()) {
			} else if ("thickBot" === reader.GetName()) {
			} else if ("thickTop" === reader.GetName()) {
			} else if ("ph" === reader.GetName()) {
			}
		}
	};

	AscCommonExcel.Cell.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		var ctx = reader.GetContext();
		var value = ctx.cellValue;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			if ("v" === reader.GetName()) {
				value.fromXml(reader);
				if (CellValueType.String === this.type) {
					var ss;
					if (ctx.originType === "str") {
						ss = prepareTextFromXml(value.val);
					} else {
						ss = reader.GetContext().sharedStrings[parseInt(value.val)];
					}

					if (undefined !== ss) {
						if (typeof ss === 'string') {
							this.setValueTextInternal(ss);
						} else {
							this.setValueMultiTextInternal(ss);
						}
					}
				} else if (CellValueType.Error === this.type) {
					this.setValueTextInternal(value.val);
				} else {
					this.setValueNumberInternal(parseFloat(value.val));
				}
			} else if ("f" === reader.GetName()) {
				var val = reader.GetContext().InitOpenManager.tmp && reader.GetContext().InitOpenManager.tmp.formula;
				val.fromXml(reader);
			}
		}
	};
	AscCommonExcel.Cell.prototype.readAttr = function (reader) {
		var val;
		var ctx = reader.GetContext();
		var cellBase = ctx.cellBase;
		ctx.originType = null;
		while (reader.MoveToNextAttribute()) {
			if ("r" === reader.GetName()) {
				val = reader.GetValue();
				cellBase.fromRefA1(val);
				this.setRowCol(cellBase.row, cellBase.col);
				this.ws.nRowsCount = Math.max(this.ws.nRowsCount, this.nRow);
				this.ws.nColsCount = Math.max(this.ws.nColsCount, this.nCol);
				this.ws.cellsByColRowsCount = Math.max(this.ws.cellsByColRowsCount, this.nCol);
			} else if ("s" === reader.GetName()) {
				var nStyleIndex = reader.GetValueInt();
				if (0 !== nStyleIndex) {
					var xfs = reader.GetContext().InitOpenManager.aCellXfs[nStyleIndex];
					if (null != xfs) {
						this.setStyle(xfs);
					}
				}
			} else if ("t" === reader.GetName()) {
				val = reader.GetValue();
				ctx.originType = val;
				var type = FromXml_ST_CellValueType(val);
				if (type != null) {
					this.type = type;
				}
			}
		}
	};
	AscCommonExcel.Cell.prototype.toXml = function (writer, name, ns) {
		var context = writer.context;
		var ws = this.ws;
		var ref = this.getName();
		var s = context.stylesForWrite.add(this.xfs) || null;
		var formulaToWrite = null;
		if (this.isFormula() && !(context.isCopyPaste && ws.bIgnoreWriteFormulas)) {
			formulaToWrite = writer.context.InitSaveManager.PrepareFormulaToWrite(this);
		}
		var text = null;
		var number = null;
		var type = null;
		//TODO ToXml_ST_CellValueType
		switch (this.type) {
			case CellValueType.String:
				type = "s";
				if (formulaToWrite) {
					text = this.text;
					type = "str";
				} else {
					var textIndex = this.getTextIndex();
					if (null !== textIndex) {
						var index = context.oSharedStrings.strings[textIndex];
						if (undefined === index) {
							index = context.oSharedStrings.index++;
							context.oSharedStrings.strings[textIndex] = index;
						}
						number = index;
					}
				}
				break;
			case CellValueType.Error:
				type = "e";
				text = this.text;
				break;
			case CellValueType.Bool:
				type = "b";
				number = this.number;
				break;
			default:
				number = this.number;
				break;
		}

		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlAttributeString("r", ref);
		writer.WriteXmlNullableAttributeNumber("s", s);
		writer.WriteXmlNullableAttributeString("t", type);
		if (!this.isNullText()) {
			writer.WriteXmlAttributesEnd();

			if (formulaToWrite && formulaToWrite.formula) {
				var formulaWrite = new AscCommonExcel.OpenFormula();
				formulaWrite.ca = formulaToWrite.ca;
				formulaWrite.t = formulaToWrite.type;
				formulaWrite.v = formulaToWrite.formula;
				formulaWrite.si = formulaToWrite.si;
				formulaWrite.ref = formulaToWrite.ref;
				formulaWrite.toXml(writer, "f");
			}

			if (null !== text) {
				writer.WriteXmlValueStringEncode2("v", prepareTextToXml(text));
			} else if (null !== number) {
				writer.WriteXmlValueNumber("v", number);
			}
			writer.WriteXmlNodeEnd(ns + name);

		} else {
			writer.WriteXmlAttributesEnd(true);
		}
	};
	AscCommonExcel.OpenFormula.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		this.v = prepareTextFromXml(reader.GetTextDecodeXml());
	};
	AscCommonExcel.OpenFormula.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("aca" === reader.GetName()) {
				this.aca = reader.GetValueBool();
			} else if ("bx" === reader.GetName()) {
				this.bx = reader.GetValueBool();
			} else if ("ca" === reader.GetName()) {
				this.ca = reader.GetValueBool();
			} else if ("del1" === reader.GetName()) {
				this.del1 = reader.GetValueBool();
			} else if ("del2" === reader.GetName()) {
				this.del2 = reader.GetValueBool();
			} else if ("dt2D" === reader.GetName()) {
				this.dt2d = reader.GetValueBool();
			} else if ("dtr" === reader.GetName()) {
				this.dtr = reader.GetValueBool();
			} else if ("r1" === reader.GetName()) {
				this.r1 = reader.GetValue();
			} else if ("r2" === reader.GetName()) {
				this.r2 = reader.GetValue();
			} else if ("ref" === reader.GetName()) {
				this.ref = reader.GetValue();
			} else if ("si" === reader.GetName()) {
				this.si = reader.GetValueInt();
			} else if ("t" === reader.GetName()) {
				val = reader.GetValue();
				this.t = FromXml_ST_CellFormulaType(val);
			}
		}
	};
	AscCommonExcel.OpenFormula.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);

		writer.WriteXmlNullableAttributeString("t", ToXml_ST_CellFormulaType(this.t));
		writer.WriteXmlNullableAttributeBool("aca", this.aca);
		writer.WriteXmlNullableAttributeString("ref", this.ref && this.ref.getName());
		writer.WriteXmlNullableAttributeBool("dt2D", this.dt2D);
		writer.WriteXmlNullableAttributeBool("dtr", this.dtr);
		writer.WriteXmlNullableAttributeBool("del1", this.del1);
		writer.WriteXmlNullableAttributeBool("del2", this.del2);
		writer.WriteXmlNullableAttributeString("r1", this.r1);
		writer.WriteXmlNullableAttributeString("r2", this.r2);
		writer.WriteXmlNullableAttributeBool("ca", this.ca);
		writer.WriteXmlNullableAttributeNumber("si", this.si);
		writer.WriteXmlNullableAttributeBool("bx", this.bx);

		writer.WriteXmlAttributesEnd();

		writer.WriteXmlStringEncode(this.v);

		writer.WriteXmlNodeEnd(ns + name);
	};


	function CT_DrawingWS(ws) {
		this.ws = ws;
		this.anchors = [];
	}

	CT_DrawingWS.prototype.fromXml = function (reader) {
		if (!reader.ReadNextNode()) {
			return;
		}
		if ("wsDr" !== reader.GetNameNoNS()) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		var objectRender = new AscFormat.DrawingObjects();
		if ("wsDr" === reader.GetNameNoNS()) {
			var depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				var name = reader.GetNameNoNS();
				if ("twoCellAnchor" === name) {
					let drawing = objectRender.createDrawingObject(AscCommon.c_oAscCellAnchorType.cellanchorTwoCell);
					drawing.fromXml(reader);
				} else if ("oneCellAnchor" === name) {
					let drawing = objectRender.createDrawingObject(AscCommon.c_oAscCellAnchorType.cellanchorOneCell);
					drawing.fromXml(reader);
				} else if ("absoluteAnchor" === name) {
					let drawing = objectRender.createDrawingObject(AscCommon.c_oAscCellAnchorType.cellanchorAbsolute);
					drawing.fromXml(reader);
				}
			}
		}
	};
	CT_DrawingWS.prototype.toXml = function (writer) {
		writer.WriteXmlString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
		writer.WriteXmlNodeStart("xdr:wsDr");
		writer.WriteXmlString(' xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"');
		writer.WriteXmlAttributesEnd();
		this.ws.Drawings.forEach(function (drawing) {
			drawing.toXml(writer);
		});
		writer.WriteXmlNodeEnd("xdr:wsDr");
	};

	function CT_SharedStrings() {
		this.sharedStrings = [];
	}

	CT_SharedStrings.prototype.initFromMap = function (wb, oSharedStrings) {
		for (var i in oSharedStrings.strings) {
			if (oSharedStrings.strings.hasOwnProperty(i)) {
				var from = parseInt(i);
				var to = oSharedStrings.strings[i];
				this.sharedStrings[to] = wb.sharedStrings.get(from);
			}
		}
	};
	CT_SharedStrings.prototype.fromXml = function (reader) {
		if (!reader.ReadNextNode()) {
			return;
		}
		if ("sst" !== reader.GetNameNoNS()) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		var si = new CT_Si();
		if ("sst" === reader.GetNameNoNS()) {
			var depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				var name = reader.GetNameNoNS();
				if ("si" === name) {
					si.clean();
					si.fromXml(reader);
					if (null !== si.text) {
						this.sharedStrings.push(si.text);
					} else if (null !== si.multiText) {
						this.sharedStrings.push(si.multiText);
					} else {
						this.sharedStrings.push("");
					}
				}
			}
		}
		reader.GetContext().sharedStrings = this.sharedStrings;
	};
	CT_SharedStrings.prototype.toXml = function (writer) {
		writer.WriteXmlString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
		writer.WriteXmlNodeStart("sst");
		writer.WriteXmlString(' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
		writer.WriteXmlAttributeNumber("count", this.sharedStrings.length);
		writer.WriteXmlAttributeNumber("uniqueCount", this.sharedStrings.length);
		writer.WriteXmlAttributesEnd();

		var si = new CT_Si();
		this.sharedStrings.forEach(function (elem) {
			si.clean();
			if (typeof elem === 'string') {
				si.text = elem;
			} else {
				si.multiText = elem;
			}
			si.toXml(writer, "si");
		});
		writer.WriteXmlNodeEnd("sst");
	};

	function CT_PersonList() {
		this.personList = [];
	}

	CT_PersonList.prototype.fromXml = function (reader) {
		if (!reader.ReadNextNode()) {
			return;
		}

		var sName = reader.GetNameNoNS();
		if ("personList" === sName) {
			if (reader.IsEmptyNode()) {
				return;
			}
			var depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				var name = reader.GetNameNoNS();
				if ("person" === name) {
					var val;
					var person = {providerId: "", userId: "", displayName: ""};
					while (reader.MoveToNextAttribute()) {
						if ("displayName" === reader.GetName()) {
							val = reader.GetValue();
							person.displayName = val;
						} else if ("userId" === reader.GetName()) {
							val = reader.GetValue();
							person.userId = val;
						} else if ("providerId" === reader.GetName()) {
							val = reader.GetValue();
							person.providerId = val;
						} else if ("id" === reader.GetName()) {
							val = reader.GetValue();
							person.id = val;
						}
					}

					this.personList.push(person);

					/*if (reader.IsEmptyNode()) {
						continue;
					}
					var depth2 = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depth2)) {
						var name2 = reader.GetNameNoNS();
						if ("extLst" === name2) {

						}
					}*/
				}
			}
		}
	};
	CT_PersonList.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}
		if (!name) {
			name = "personList";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlString(" xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\"");
		writer.WriteXmlString(" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
		writer.WriteXmlAttributesEnd();

		for (var i = 0; i < this.personList.length; ++i) {
			if (this.personList[i]) {
				writer.WriteXmlString("<person");
				writer.WriteXmlNullableAttributeStringEncode("displayName", this.personList[i].displayName);
				writer.WriteXmlNullableAttributeString("id", this.personList[i].id);
				if (this.personList[i].userId) {
					writer.WriteXmlNullableAttributeString("userId", this.personList[i].userId);
				}
				if (this.personList[i].providerId) {
					writer.WriteXmlNullableAttributeStringEncode("providerId", this.personList[i].providerId);
				}
				writer.WriteXmlString("/>");
			}
		}

		writer.WriteXmlNodeEnd(ns + name);
	};
	CT_PersonList.prototype.getByGuid = function (name) {
		for (var i = 0; i < this.personList.length; ++i) {
			if (this.personList[i]) {
				if (name === this.personList[i].id) {
					return this.personList[i];
				}
			}
		}
	};

	function CT_Si() {
		this.text = null;
		this.multiText = null;
	}

	CT_Si.prototype.clean = function () {
		this.text = null;
		this.multiText = null;
	};
	CT_Si.prototype.fromXml = function (reader) {
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			if ("t" === reader.GetName()) {
				this.text = prepareTextFromXml(reader.GetTextDecodeXml());
			} else if ("r" === reader.GetName()) {
				var oMultiText = new AscCommonExcel.CMultiTextElem();
				oMultiText.fromXml(reader);
				if (!this.multiText) {
					this.multiText = [];
				}
				this.multiText.push(oMultiText);
			}
		}
	};
	CT_Si.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}
		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlAttributesEnd();

		if (null !== this.text) {
			writer.WriteXmlValueStringEncode("t", prepareTextToXml(this.text));
		} else if (null !== this.multiText) {
			writer.WriteXmlArray(this.multiText, "r");
		}
		writer.WriteXmlNodeEnd(ns + name);
	};
	CT_Si.prototype.getText = function () {
		var res = "";

		if (this.multiText) {
			for (var i = 0; i < this.multiText.length; i++) {
				res += this.multiText[i].text;
			}
		} else if (this.text) {
			res += this.text;
		}

		return res;
	};

	AscCommonExcel.CMultiTextElem.prototype.fromXml = function (reader) {
		if (reader.IsEmptyNode()) {
			return;
		}

		var wb = reader.context.InitOpenManager.wb;
		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("rPr" === name) {
				val = new AscCommonExcel.Font();
				val.fromXml(reader);
				if (wb) {
					val.checkSchemeFont(wb.theme);
				}
				this.format = val;
			} else if ("t" === name) {
				if (null == this.text) {
					this.text = "";
				}

				//TODO поработать с текстом, проблемы с переносом строки
				this.text += prepareTextFromXml(reader.GetTextDecodeXml());
			}
		}
	};

	AscCommonExcel.CMultiTextElem.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlAttributesEnd();

		if (this.format) {
			this.format.toXml(writer, "rPr", childns);
		}
		if (this.text) {
			writer.WriteXmlValueStringEncode("t", prepareTextToXml(this.text));
		}
		writer.WriteXmlNodeEnd(ns + name);
	};


	function CT_PivotCaches() {
		this.pivotCaches = [];
	}

	CT_PivotCaches.prototype.fromXml = function (reader) {
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			if ("pivotCache" === reader.GetNameNoNS()) {
				var pivotCache = new CT_PivotCache();
				pivotCache.fromXml(reader);
				this.pivotCaches.push(pivotCache);
			}
		}
	};
	CT_PivotCaches.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.pivotCaches.forEach(function (pivotCache) {
			pivotCache.toXml(writer, "pivotCache");
		});
		writer.WriteXmlNodeEnd(name);
	};
	CT_PivotCaches.prototype.initFromPivotCaches = function(part, writer, pivotCaches) {
		for (let i in pivotCaches) {
			if (pivotCaches.hasOwnProperty(i)) {
				let pivotCache = pivotCaches[i];
				let pivotCacheDefinitionPart = part.addPart(AscCommon.openXml.Types.pivotTableCacheDefinitionWorkbook);
				pivotCacheDefinitionPart.part.setDataXml(pivotCache.cache, writer);

				if (pivotCache.cache.cacheRecords) {
					let cacheRecordsPart = pivotCacheDefinitionPart.part.addPart(AscCommon.openXml.Types.pivotTableCacheRecords);
					cacheRecordsPart.part.setDataXml(pivotCache.cache.cacheRecords, writer);
				}

				writer.context.oUriMap[pivotCache.cache.Get_Id()] = pivotCacheDefinitionPart.part.uri;

				let elem = new CT_PivotCache();
				elem.cacheId = pivotCache.id;
				elem.id = pivotCacheDefinitionPart.rId;
				this.pivotCaches.push(elem);
			}
		}
	};

	function CT_DrawingWSRef() {
		this.id = null;
	}

	CT_DrawingWSRef.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_DrawingWSRef.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("r:id", this.id);
		writer.WriteXmlAttributesEnd(true);
	};
	CT_DrawingWSRef.prototype.readAttr = function (reader) {
		while (reader.MoveToNextAttribute()) {
			if ("id" === reader.GetNameNoNS()) {
				this.id = reader.GetValueDecodeXml();
			}
		}
	};

	function CT_SheetData(ws) {
		this.ws = ws;
		this._openRow = new AscCommonExcel.Row(ws);
	}

	CT_SheetData.prototype.fromXml = function (reader) {
		var depth = reader.GetDepth();
		var row = reader.GetContext().row;
		while (reader.ReadNextSiblingNode(depth)) {
			if ("row" === reader.GetName()) {
				row.clear();
				row.fromXml(reader);
				row.saveContent();
			}
		}
	};
	CT_SheetData.prototype.fromXml2 = function (reader) {

		var depth = reader.GetDepth();
		var tmp = reader.GetContext().InitOpenManager.tmp;
		if (tmp) {
			while (reader.ReadNextSiblingNode(depth)) {
				if ("row" === reader.GetName()) {
					tmp.row.clear();
					tmp.row.fromXml2(reader);
				} /*else if (xlsb) {

			}*/
			}
		}

		/*var res = c_oSerConstants.ReadOk;
		var oThis = this;
		if ( c_oSerWorksheetsTypes.XlsbPos === type )
		{
			var oldPos = this.stream.GetCurPos();
			this.stream.Seek2(this.stream.GetULongLE());

			tmp.ws.fromXLSB(this.stream, this.stream.XlsbReadRecordType(), tmp, this.aCellXfs, this.aSharedStrings,
				function(tmp) {
					oThis.initCellAfterRead(tmp);
				});

			this.stream.Seek2(oldPos);
			res = c_oSerConstants.ReadUnknown;
		}
		else if ( c_oSerWorksheetsTypes.Row === type )
		{
			tmp.pos =  null;
			tmp.len = null;
			tmp.row.clear();
			res = this.bcr.Read2Spreadsheet(length, function(t,l){
				return oThis.ReadRow(t,l, tmp);
			});
			if(null === tmp.row.index) {
				tmp.row.index = tmp.prevRow + 1;
			}
			tmp.row.saveContent();
			tmp.ws.cellsByColRowsCount = Math.max(tmp.ws.cellsByColRowsCount, tmp.row.index + 1);
			tmp.ws.nRowsCount = Math.max(tmp.ws.nRowsCount, tmp.ws.cellsByColRowsCount);
			tmp.prevRow = tmp.row.index;
			tmp.prevCol = -1;
			//читаем ячейки
			if (null !== tmp.pos && null !== tmp.len) {
				var nOldPos = this.stream.GetCurPos();
				this.stream.Seek2(tmp.pos);
				res = this.bcr.Read1(tmp.len, function(t,l){
					return oThis.ReadCells(t,l, tmp);
				});
				this.stream.Seek2(nOldPos);
			}
		}
		else
			res = c_oSerConstants.ReadUnknown;
		return res;*/


	};

	AscCommonExcel.Hyperlink.prototype.fromXml = function (reader, ws) {
		this.readAttr(reader, ws);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	AscCommonExcel.Hyperlink.prototype.readAttr = function (reader, ws) {
		var val;
		while (reader.MoveToNextAttribute()) {
			/*if ("display" === reader.GetName()) {
				val = reader.GetValue();
				this.display = val;
			} else*/
			if ("r:id" === reader.GetName()) {
				//TODO пока обрабатываю здесь, возможно стоит это сдлелать после чтения всей книги
				val = reader.GetValue();
				var rIdDoc = reader.rels.getRelationshipById(val);
				if (rIdDoc) {
					this.Hyperlink = rIdDoc.targetFullName;
				}
				/*} else if ("relationships:id" === reader.GetName()) {
					val = reader.GetValue();
					this.relationships:id = val;
				*/
			} else if ("location" === reader.GetName()) {
				val = reader.GetValue();
				this.setLocation(val);
				//this.setLocation(this.stream.GetString2LE(length));
			} else if ("ref" === reader.GetName()) {
				val = reader.GetValue();
				this.Ref = ws.getRange2(val);
			} else if ("tooltip" === reader.GetName()) {
				val = reader.GetValue();
				this.Tooltip = val;
			} /*else if ("hyperlink" === reader.GetName()) {

			}*/
		}
	};

	AscCommonExcel.Hyperlink.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		//writer.WriteXmlNullableAttributeStringEncode("display", this.display);
		writer.WriteXmlNullableAttributeString("r:id", this.rId);
		writer.WriteXmlNullableAttributeStringEncode("location", this.Location);
		writer.WriteXmlNullableAttributeStringEncode("ref", this.Ref && this.Ref.bbox && this.Ref.bbox.getName());
		writer.WriteXmlNullableAttributeStringEncode("tooltip", this.Tooltip);
		writer.WriteXmlAttributesEnd(true);
	};

	function CT_Sheets(wb) {
		this.wb = wb;
		this.sheets = [];
	}

	CT_Sheets.prototype.fromXml = function (reader) {
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			if ("sheet" === reader.GetNameNoNS()) {
				var sheet = new CT_Sheet();
				sheet.fromXml(reader);
				this.sheets.push(sheet);
			}
		}
	};
	CT_Sheets.prototype.toXml = function (writer, name, ns) {
		var t = this;
		var context = writer.context;
		var index = 1;

		this.wb.forEach(function (ws) {
			var sheetXml = new CT_Sheet();
			var wsPart = context.part.addPart(AscCommon.openXml.Types.worksheet);
			wsPart.part.setDataXml(ws, writer);
			sheetXml.id = wsPart.rId;
			sheetXml.sheetId = index++;
			sheetXml.name = ws.getName();
			sheetXml.bHidden = ws.bHidden;
			t.sheets.push(sheetXml);
			context.sheetIds[ws.getId()] = sheetXml.sheetId;
		}, context.isCopyPaste);

		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"sheets"*/);
		writer.WriteXmlAttributesEnd();
		this.sheets.forEach(function (sheetXml) {
			sheetXml.toXml(writer);
		}, context.isCopyPaste);
		writer.WriteXmlNodeEnd(ns + name);
	};

	function CT_Sheet() {
		//Attributes
		this.name = null;
		this.sheetId = null;
		this.id = null;
		this.bHidden = null;
	}

	CT_Sheet.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_Sheet.prototype.toXml = function (writer) {
		writer.WriteXmlNodeStart("sheet");
		writer.WriteXmlNullableAttributeString("name", this.name);
		writer.WriteXmlNullableAttributeNumber("sheetId", this.sheetId);
		writer.WriteXmlNullableAttributeString("r:id", this.id);
		if (this.bHidden != null) {
			writer.WriteXmlAttributeString("state", this.bHidden ? "hidden" : "visible");
		}
		writer.WriteXmlAttributesEnd(true);
	};
	CT_Sheet.prototype.readAttributes = function (attr, uq) {
		if (attr()) {
			this.parseAttributes(attr());
		}
	};
	CT_Sheet.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			var name = reader.GetNameNoNS();
			if ("name" === name) {
				this.name = reader.GetValueDecodeXml();
			} else if ("sheetId" === name) {
				this.sheetId = reader.GetValueInt();
			} else if ("id" === name) {
				this.id = reader.GetValueDecodeXml();
			} else if ("state" === name) {
				val = reader.GetValue();
				if ("hidden" === val) {
					this.bHidden = true;
				} else if ("veryHidden" === val) {
					this.bHidden = true;
				} else if ("visible" === val) {
					this.bHidden = false;
				}
			}
		}
	};
	CT_Sheet.prototype.parseAttributes = function (vals, uq) {
		var val;
		val = vals["r:id"];
		if (undefined !== val) {
			this.id = AscCommon.unleakString(uq(val));
		}
	};

	function CT_Value() {
		//Attributes
		this.space = null;
		this.val = null;
	}

	CT_Value.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		this.val = reader.GetTextDecodeXml();
	};
	CT_Value.prototype.readAttributes = function (attr, uq) {
		if (attr()) {
			this.parseAttributes(attr());
		}
	};
	CT_Value.prototype.readAttr = function (reader) {
		//todo space
	};
	CT_Value.prototype.parseAttributes = function (vals, uq) {
		var val;
		val = vals["r:id"];
		if (undefined !== val) {
			this.id = AscCommon.unleakString(uq(val));
		}
	};

	function CT_PivotCache() {
		//Attributes
		this.cacheId = null;
		this.id = null;
	}

	CT_PivotCache.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_PivotCache.prototype.readAttributes = function (attr, uq) {
		if (attr()) {
			var vals = attr();
			this.parseAttributes(attr(), uq);
		}
	};
	CT_PivotCache.prototype.readAttr = function (reader) {
		while (reader.MoveToNextAttribute()) {
			var name = reader.GetNameNoNS();
			if ("id" === name) {
				this.id = reader.GetValueDecodeXml();
			} else if ("cacheId" === name) {
				this.cacheId = parseInt(reader.GetValue());
			}
		}
	};
	CT_PivotCache.prototype.parseAttributes = function (vals, uq) {
		var val;
		val = vals["cacheId"];
		if (undefined !== val) {
			this.cacheId = val - 0;
		}
		val = vals["r:id"];
		if (undefined !== val) {
			this.id = AscCommon.unleakString(uq(val));
		}
	};
	CT_PivotCache.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeUInt("cacheId", this.cacheId);
		writer.WriteXmlNullableAttributeStringEncode("r:id", this.id);
		writer.WriteXmlAttributesEnd(true);
	};


	//Tables & AutoFilter
	AscCommonExcel.TablePart.prototype.fromXml = function (reader) {
		if (!reader.ReadNextNode()) {
			return;
		}

		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("autoFilter" === name) {
				var autoFilter = new AscCommonExcel.AutoFilter();
				autoFilter.fromXml(reader);
				this.AutoFilter = autoFilter;
			} else if ("sortState" === name) {
				var sortState = new AscCommonExcel.SortState();
				sortState.fromXml(reader);
				this.SortState = sortState;
			} else if ("tableColumns" === name) {
				var _depth = reader.GetDepth();
				while (reader.ReadNextSiblingNode(_depth)) {
					var _name = reader.GetNameNoNS();
					if ("tableColumn" === _name) {
						var tableColumn = new AscCommonExcel.TableColumn();
						tableColumn.fromXml(reader);
						if (!this.TableColumns) {
							this.TableColumns = [];
						}
						this.TableColumns.push(tableColumn);
					}
				}
			} else if ("tableStyleInfo" === name) {
				var tableStyleInfo = new AscCommonExcel.TableStyleInfo();
				tableStyleInfo.fromXml(reader);
				this.TableStyleInfo = tableStyleInfo;
			} else if ("extLst" === name) {

			}
		}
	};
	AscCommonExcel.TablePart.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("ref" === reader.GetName()) {
				val = reader.GetValue();
				this.Ref = AscCommonExcel.g_oRangeCache.getAscRange(val);
			} else if ("displayName" === reader.GetName()) {
				val = reader.GetValue();
				this.DisplayName = val;
			} else if ("headerRowCount" === reader.GetName()) {
				val = reader.GetValueInt();
				this.HeaderRowCount = val;
			} else if ("totalsRowCount" === reader.GetName()) {
				val = reader.GetValueInt();
				this.TotalsRowCount = val;
			} else if ("id" === reader.GetName()) {
				reader.context.InitOpenManager.oReadResult.tableIds[reader.GetValue()] = this;
			} else if ("tableType" === reader.GetName()) {
				val = reader.GetValue();
				this.tableType = val;
			}
		}
	};
	AscCommonExcel.TablePart.prototype.toXml = function (writer) {

		writer.WriteXmlString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');
		writer.WriteXmlNodeStart("table");
		writer.WriteXmlString(
			' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"');
		var tableIds = writer.context.InitSaveManager.getTableIds();
		writer.WriteXmlNullableAttributeNumber("id", undefined != tableIds[this.DisplayName] ? tableIds[this.DisplayName].id : null);
		writer.WriteXmlNullableAttributeStringEncode("name", this.Name ? this.Name : null);
		writer.WriteXmlNullableAttributeStringEncode("displayName", this.DisplayName);
		writer.WriteXmlNullableAttributeString("ref", this.Ref.getName());
		writer.WriteXmlNullableAttributeStringEncode("tableType", this.tableType);
		writer.WriteXmlNullableAttributeNumber("headerRowCount", this.HeaderRowCount);
		writer.WriteXmlNullableAttributeNumber("totalsRowCount", this.TotalsRowCount);

		writer.WriteXmlAttributesEnd();

		//AscCommonExcel.AutoFilter
		if (this.AutoFilter) {
			this.AutoFilter.toXml(writer, "autoFilter");
		}
		//AscCommonExcel.SortState
		if (this.SortState) {
			this.SortState.toXml(writer, "sortState");
		}
		if (this.TableColumns) {
			writer.WriteXmlNodeStart("tableColumns");
			writer.WriteXmlAttributeNumber("count", this.TableColumns.length);
			writer.WriteXmlAttributesEnd();

			for (var i = 0; i < this.TableColumns.length; ++i) {
				//AscCommonExcel.TableColumn
				this.TableColumns[i].toXml(writer, "tableColumn", "", i);
			}
			writer.WriteXmlNodeEnd("tableColumns");
		}
		//AscCommonExcel.TableStyleInfo
		if (this.TableStyleInfo) {
			this.TableStyleInfo.toXml(writer, "tableStyleInfo");
		}

		var context = writer.context;
		if (this.QueryTable) {
			var queryTable = context.part.addPart(AscCommon.openXml.Types.queryTable);
			queryTable.part.setDataXml(this.QueryTable, writer);
		}

		/*if(m_oExtLst.IsInit())
		{
			writer.WriteString(m_oExtLst->toXMLWithNS(L""));
		}*/

		writer.WriteXmlNodeEnd("table");
	};
	AscCommonExcel.TablePart.prototype.setAttr = function (attr, val, oAttr) {
		//была идея делать так для удобства, но память нужно экономить
		/*var oAttr = {"ref": "Ref", "displayName" : "DisplayName", "headerRowCount" : "HeaderRowCount"};
		while (reader.MoveToNextAttribute()) {
			this.setAttr(reader.GetName(), reader.GetValue(), oAttr);
		}*/

		if (oAttr[attr]) {
			this[oAttr[attr]] = val;
		}
	};

	AscCommonExcel.AutoFilter.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			if ("filterColumn" === reader.GetName()) {
				var filterColumn = new AscCommonExcel.FilterColumn();
				filterColumn.fromXml(reader);
				if (!this.FilterColumns) {
					this.FilterColumns = [];
				}
				this.FilterColumns.push(filterColumn);
			} else if ("sortState" === reader.GetName()) {
				var sortState = new AscCommonExcel.SortState();
				sortState.fromXml(reader);
				this.SortState = sortState;
			}
		}
	};
	AscCommonExcel.AutoFilter.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("ref" === reader.GetName()) {
				val = reader.GetValue();
				this.Ref = AscCommonExcel.g_oRangeCache.getAscRange(val);
			}
		}
	};
	AscCommonExcel.AutoFilter.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}
		writer.WriteXmlNodeStart(ns + name/*"autoFilter"*/);
		if (null !== this.Ref) {
			writer.WriteXmlAttributeStringEncode("ref", this.Ref.getName());
		}
		writer.WriteXmlAttributesEnd();
		if (this.FilterColumns) {
			for (var i = 0; i < this.FilterColumns.length; ++i) {
				var elem = this.FilterColumns[i];
				elem.toXml(writer, "filterColumn");
			}
		}
		if (this.SortState) {
			this.SortState.toXml(writer, "sortState");
		}

		writer.WriteXmlNodeEnd(ns + name);
	};

	AscCommonExcel.TableStyleInfo.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (!reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}

	};
	AscCommonExcel.TableStyleInfo.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.Name = val;
			} else if ("showColumnStripes" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ShowColumnStripes = val;
			} else if ("showFirstColumn" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ShowFirstColumn = val;
			} else if ("showLastColumn" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ShowLastColumn = val;
			} else if ("showRowStripes" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ShowRowStripes = val;
			}
		}
	};

	AscCommonExcel.TableStyleInfo.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"tableStyleInfo"*/);

		writer.WriteXmlNullableAttributeStringEncode("name", this.Name);
		writer.WriteXmlAttributeNumber("showFirstColumn", boolToNumber(this.ShowFirstColumn));
		writer.WriteXmlAttributeNumber("showLastColumn", boolToNumber(this.ShowLastColumn));
		writer.WriteXmlAttributeNumber("showRowStripes", boolToNumber(this.ShowRowStripes));
		writer.WriteXmlAttributeNumber("showColumnStripes", boolToNumber(this.ShowColumnStripes));
		writer.WriteXmlAttributesEnd();

		writer.WriteXmlNodeEnd(ns + name);
	};

	AscCommonExcel.TableColumn.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("totalsRowFormula" === name) {
				/*var formula = this.stream.GetString2LE(length);
				this.oReadResult.tableCustomFunc.push({formula: formula, column: oTableColumn, ws: this.ws});*/
			} else if ("calculatedColumnFormula" === name) {
				//reader.context.InitOpenManager.Dxfs[
				/*var DxfId = this.stream.GetULongLE();
				oTableColumn.dxf = this.Dxfs[DxfId];*/
			}
		}

		/*int nCurDepth = oReader.GetDepth();
		while( oReader.ReadNextSiblingNode( nCurDepth ) )
		{
			std::wstring sName = XmlUtils::GetNameNoNS(oReader.GetName());

			if ( (L"totalsRowFormula") == sName )
			m_oTotalsRowFormula = oReader.GetText3();
		else if ( (L"calculatedColumnFormula") == sName )
			m_oCalculatedColumnFormula = oReader.GetText3();
		}*/
	};
	AscCommonExcel.TableColumn.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("name" === reader.GetName()) {
				val = reader.GetValueDecodeXml();
				this.Name = prepareTextFromXml(val);
			} else if ("totalsRowLabel" === reader.GetName()) {
				val = reader.GetValue();
				this.TotalsRowLabel = val;
			} else if ("totalsRowFunction" === reader.GetName()) {
				val = reader.GetValue();
				this.TotalsRowFunction = val;
			} else if ("dataDxfId" === reader.GetName()) {
				val = reader.GetValue();
				this.dxf = reader.context.InitOpenManager.Dxfs[val];
			} else if ("showRowStripes" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ShowRowStripes = val;
			} else if ("queryTableFieldId" === reader.GetName()) {
				val = reader.GetValue();
				this.queryTableFieldId = val;
				//oTableColumn.queryTableFieldId = this.stream.GetULongLE();
			} else if ("uniqueName" === reader.GetName()) {
				val = reader.GetValue();
				this.uniqueName = val;
			}
		}
	};


	AscCommonExcel.TableColumn.prototype.toXml = function (writer, name, ns, index) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"tableColumn"*/);

		writer.WriteXmlNullableAttributeNumber("id", this.Id ? this.Id : index + 1);
		writer.WriteXmlNullableAttributeStringEncode("name", this.Name);
		writer.WriteXmlNullableAttributeStringEncode("uniqueName", this.uniqueName ? this.uniqueName : null);
		writer.WriteXmlNullableAttributeStringEncode("totalsRowLabel", this.TotalsRowLabel);
		//есть такой баг: при сохранениии "sum" и названия таблицы "Table1" (русский excel), выдается ошибка в формулах
		writer.WriteXmlNullableAttributeString("totalsRowFunction", this.TotalsRowFunction);
		writer.WriteXmlNullableAttributeNumber("queryTableFieldId", this.queryTableFieldId ? this.queryTableFieldId : null);
		//writer.WriteXmlNullableAttributeString("dataCellStyle", this.DataCellStyle ? this.DataCellStyle : null);

		if (null != this.dxf) {
			writer.WriteXmlAttributeNumber("dataDxfId", writer.context.InitSaveManager.aDxfs.length);
			writer.context.InitSaveManager.aDxfs.push(this.dxf);
		}

		//writer.WriteXmlNullableAttributeString("headerRowCellStyle", this.HeaderRowCellStyle ? this.HeaderRowCellStyle : null);
		//writer.WriteXmlNullableAttributeNumber("headerRowDxfId", this.HeaderRowDxfId ? this.HeaderRowDxfId : null);
		//writer.WriteXmlNullableAttributeString("totalsRowCellStyle", this.TotalsRowCellStyle ? this.TotalsRowCellStyle : null);
		//writer.WriteXmlNullableAttributeNumber("totalsRowDxfId", this.TotalsRowDxfId ? this.TotalsRowDxfId : null);
		writer.WriteXmlAttributesEnd();

		writer.WriteXmlNodeEnd(ns + name);
	};

	AscCommonExcel.SortState.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("sortCondition" === name) {
				var sortCondition = new AscCommonExcel.SortCondition();
				sortCondition.fromXml(reader);
				if (!this.SortConditions) {
					this.SortConditions = [];
				}
				this.SortConditions.push(sortCondition);
			}
		}
	};

	AscCommonExcel.SortState.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("ref" === reader.GetName()) {
				val = reader.GetValue();
				this.Ref = AscCommonExcel.g_oRangeCache.getAscRange(val);
			} else if ("caseSensitive" === reader.GetName()) {
				val = reader.GetValueBool();
				this.CaseSensitive = val;
			} else if ("columnSort" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ColumnSort = val;
			} else if ("sortMethod" === reader.GetName()) {
				val = reader.GetValue();
				this.SortMethod = val;
			}
		}
	};

	AscCommonExcel.SortState.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"sortState"*/);
		writer.WriteXmlString(' xmlns:xlrd2="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"');

		if (null !== this.Ref) {
			writer.WriteXmlAttributeStringEncode("ref", this.Ref.getName());
		}
		if (null !== this.CaseSensitive) {
			writer.WriteXmlAttributeBool("caseSensitive", this.CaseSensitive);
		}
		if (null !== this.ColumnSort) {
			writer.WriteXmlAttributeBool("columnSort", this.ColumnSort);
		}
		if (null !== this.SortMethod) {
			writer.WriteXmlAttributeString("sortMethod", this.SortMethod);
		}
		writer.WriteXmlAttributesEnd();

		if (this.SortConditions) {
			for (var i = 0; i < this.SortConditions.length; ++i) {
				this.SortConditions[i].toXml(writer, "sortCondition");
			}
		}

		writer.WriteXmlNodeEnd(ns + name);
	};

	AscCommonExcel.FilterColumn.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("colorFilter" === name) {
				val = new Asc.ColorFilter();
				val.fromXml(reader);
				this.ColorFilter = val;
			} else if ("dynamicFilter" === name) {
				val = new Asc.DynamicFilter();
				val.fromXml(reader);
				this.DynamicFilter = val;
			} else if ("customFilters" === name) {
				val = new Asc.CustomFilters();
				val.fromXml(reader);
				this.CustomFiltersObj = val;
			} else if ("filters" === name) {
				val = new AscCommonExcel.Filters();
				val.fromXml(reader);
				this.Filters = val;
				this.Filters.sortDate();
				/*oFilterColumn.Filters = new AscCommonExcel.Filters();
				res = this.bcr.Read1(length, function(t,l){
					return oThis.ReadFilters(t,l, oFilterColumn.Filters);
				});
				oFilterColumn.Filters.sortDate();*/
			} else if ("top10" === name) {
				val = new AscCommonExcel.Top10();
				val.fromXml(reader);
				this.Filters = val;
			}
		}
	};

	AscCommonExcel.FilterColumn.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("colId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.ColId = val;
			} else if ("hiddenButton" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ShowButton = !val;
			} else if ("showButton" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ShowButton = val;
			}
		}
	};

	AscCommonExcel.FilterColumn.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"filterColumn"*/);

		if (null !== this.ColId) {
			writer.WriteXmlAttributeNumber("colId", this.ColId);
		}
		if (true !== this.ShowButton) {
			if (this.ShowButton) {
				writer.WriteXmlAttributeBool("showButton", this.ShowButton);
			} else {
				writer.WriteXmlAttributeBool("hiddenButton", !this.ShowButton);
			}
		}
		writer.WriteXmlAttributesEnd();

		if (null !== this.ColorFilter) {
			this.ColorFilter.toXml(writer, "colorFilter", childns, childns);
		}
		if (null !== this.CustomFiltersObj) {
			this.CustomFiltersObj.toXml(writer, "customFilters", childns, childns);
		}
		if (null !== this.DynamicFilter) {
			this.DynamicFilter.toXml(writer, "dynamicFilter", childns, childns);
		}
		if (null !== this.Filters) {
			this.Filters.toXml(writer, "filters", childns, childns);
		}
		if (null !== this.Top10) {
			this.Top10.toXml(writer, "top10", childns, childns);
		}
		writer.WriteXmlNodeEnd(ns + name);

	};

	AscCommonExcel.SortCondition.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (!reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	AscCommonExcel.SortCondition.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("descending" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ConditionDescending = val ? Asc.c_oAscSortOptions.Descending : Asc.c_oAscSortOptions.Ascending;
			} else if ("ref" === reader.GetName()) {
				val = AscCommonExcel.g_oRangeCache.getAscRange(reader.GetValue());
				this.Ref = val;
			} else if ("sortBy" === reader.GetName()) {
				val = reader.GetValue();
				this.ConditionSortBy = FromXml_ST_SortBy(val);
			} else if ("dxfId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.dxf = reader.context.InitOpenManager.Dxfs[val];
			}
		}
	};

	AscCommonExcel.SortCondition.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"sortCondition"*/);

		if (null !== this.ConditionSortBy) {
			writer.WriteXmlAttributeString("sortBy", ToXml_ST_SortBy(this.ConditionSortBy));
		}
		if (null !== this.ConditionDescending) {
			writer.WriteXmlAttributeBool("descending", this.ConditionDescending === Asc.c_oAscSortOptions.Descending);
		}
		if (null !== this.Ref) {
			writer.WriteXmlAttributeStringEncode("ref", this.Ref.getName());
		}
		if (null != this.dxf) {
			writer.WriteXmlAttributeNumber("dxfId", writer.context.InitSaveManager.aDxfs.length);
			writer.context.InitSaveManager.aDxfs.push(this.dxf);

		}
		writer.WriteXmlAttributesEnd();


		writer.WriteXmlNodeEnd(ns + name);
	};

	Asc.ColorFilter.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (!reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	Asc.ColorFilter.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("cellColor" === reader.GetName()) {
				val = reader.GetValueBool();
				this.CellColor = val;
			} else if ("dxfId" === reader.GetName()) {
				val = reader.GetValue();
				this.dxf = reader.context.InitOpenManager.Dxfs[val];
			}
		}
	};

	Asc.ColorFilter.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*colorFilter*/);
		if (this.dxf != null) {
			writer.WriteXmlNullableAttributeNumber("dxfId", writer.context.InitSaveManager.aDxfs.length);
			writer.context.InitSaveManager.aDxfs.push(this.dxf);
		}

		if (this.CellColor === false) {
			writer.WriteXmlNullableAttributeNumber("cellColor", 0);
		}

		writer.WriteXmlAttributesEnd(true);
	};

	Asc.DynamicFilter.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (!reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	Asc.DynamicFilter.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("type" === reader.GetName()) {
				val = reader.GetValue();
				this.Type = FromXml_ST_DynamicFilterType(val);
			} else if ("val" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.val = val;
			} else if ("maxVal" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.MaxVal = val;
			}
		}
	};

	Asc.DynamicFilter.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"dynamicFilter"*/);

		writer.WriteXmlAttributeString("type", ToXml_ST_DynamicFilterType(this.Type));
		writer.WriteXmlNullableAttributeNumber("val", this.val);
		writer.WriteXmlNullableAttributeNumber("maxVal", this.MaxVal);

		writer.WriteXmlAttributesEnd(true);
	};

	Asc.Top10.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	Asc.Top10.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("filterVal" === reader.GetName()) {
				val = reader.GetValue();
				this.FilterVal = val;
			} else if ("percent" === reader.GetName()) {
				val = reader.GetValue();
				this.Percent = val;
			} else if ("top" === reader.GetName()) {
				val = reader.GetValue();
				this.Top = val;
			} else if ("val" === reader.GetName()) {
				val = reader.GetValue();
				this.Val = val;
			}
		}
	};

	Asc.Top10.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}
		writer.WriteXmlNodeStart(ns + name);
		if (true !== this.Top) {
			writer.WriteXmlAttributeBool("top", this.Top);
		}
		if (false !== this.Percent) {
			writer.WriteXmlAttributeBool("percent", this.Percent);
		}
		if (null !== this.Val) {
			writer.WriteXmlAttributeNumber("val", this.Val);
		}
		if (null !== this.FilterVal) {
			writer.WriteXmlAttributeNumber("filterVal", this.FilterVal);
		}
		writer.WriteXmlAttributesEnd(true);
	};

	Asc.CustomFilters.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("customFilter" === name) {
				var val = new Asc.CustomFilter();
				val.fromXml(reader);
				if (!this.CustomFilters) {
					this.CustomFilters = [];
				}
				this.CustomFilters.push(val);
			}
		}
	};

	Asc.CustomFilters.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("And" === reader.GetName()) {
				val = reader.GetValueBool();
				this.And = val;
			}
		}
	};

	Asc.CustomFilters.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}

		writer.WriteXmlNodeStart(ns + name);

		writer.WriteXmlNullableAttributeString("and", this.And ? 1 : null);
		writer.WriteXmlAttributesEnd();

		for (var i = 0; i < this.CustomFilters.length; ++i) {
			this.CustomFilters[i].toXml(writer, "customFilter", childns, childns);
		}

		writer.WriteXmlNodeEnd(ns + name);
	};

	Asc.CustomFilter.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};

	Asc.CustomFilter.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("operator" === reader.GetName()) {
				val = reader.GetValue();
				this.Operator = AscCommonExcel.FromXml_ST_FilterOperator(val);
			} else if ("val" === reader.GetName()) {
				val = reader.GetValue();
				this.Val = val;
			}
		}
	};

	Asc.CustomFilter.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"сustomFilter"*/);
		writer.WriteXmlAttributeString("operator", AscCommonExcel.ToXml_ST_FilterOperator(this.Operator));
		writer.WriteXmlAttributeStringEncode("val", this.Val);
		writer.WriteXmlAttributesEnd();

		writer.WriteXmlNodeEnd(ns + name);
	};

	AscCommonExcel.Filters.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("dateGroupItem" === name) {
				val = new AscCommonExcel.DateGroupItem();
				val.fromXml(reader);
				var autoFilterDateElem = new AscCommonExcel.AutoFilterDateElem();
				autoFilterDateElem.convertDateGroupItemToRange(val);
				this.Dates.push(autoFilterDateElem);
			} else if ("filter" === name) {
				val = new AscCommonExcel.Filter();
				val.fromXml(reader);
				if (null != val.Val) {
					this.Values[val.Val] = 1;
				}
			}
		}
	};

	AscCommonExcel.Filters.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("blank" === reader.GetName()) {
				val = reader.GetValueBool();
				this.Blank = val;
			}
		}
	};

	AscCommonExcel.Filters.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		if (null !== this.Blank) {
			writer.WriteXmlAttributeBool("blank", this.Blank);
		}
		writer.WriteXmlAttributesEnd();
		for (var val in this.Values) {
			var filter = new AscCommonExcel.CT_Filter();
			filter.Val = val;
			filter.toXml(writer, "filter", childns, childns);
		}
		for (var i = 0; i < this.Dates.length; ++i) {
			var elem = this.Dates[i];
			var dateGroupItem = new AscCommonExcel.DateGroupItem();
			dateGroupItem.convertRangeToDateGroupItem(elem);
			dateGroupItem.toXml(writer, "dateGroupItem", childns, childns);
		}
		writer.WriteXmlNodeEnd(ns + name);
	};

	AscCommonExcel.CT_Filter.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}
		writer.WriteXmlNodeStart(ns + name);
		if (null !== this.Val) {
			writer.WriteXmlAttributeStringEncode("val", this.Val);
		}
		writer.WriteXmlAttributesEnd(true);
	};

	AscCommonExcel.DateGroupItem.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};

	AscCommonExcel.DateGroupItem.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("dateTimeGrouping" === reader.GetName()) {
				val = reader.GetValue();
				this.DateTimeGrouping = FromXml_ST_DateTimeGrouping(val);
			} else if ("day" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.Day = val;
			} else if ("hour" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.Hour = val;
			} else if ("minute" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.Minute = val;
			} else if ("month" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.Month = val;
			} else if ("second" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.Second = val;
			} else if ("year" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.Year = val;
			}
		}
	};

	AscCommonExcel.DateGroupItem.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"dateGroupItem"*/);

		writer.WriteXmlNullableAttributeNumber("year", this.Year);
		writer.WriteXmlNullableAttributeNumber("month", this.Month);
		writer.WriteXmlNullableAttributeNumber("day", this.Day);
		writer.WriteXmlNullableAttributeNumber("hour", this.Hour);
		writer.WriteXmlNullableAttributeNumber("minute", this.Minute);
		writer.WriteXmlNullableAttributeNumber("second", this.Second);
		writer.WriteXmlAttributeStringEncode("dateTimeGrouping", ToXml_ST_DateTimeGrouping(this.DateTimeGrouping));
		writer.WriteXmlAttributesEnd();

		writer.WriteXmlNodeEnd(ns + name);

	};

	AscCommonExcel.Filter.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};

	AscCommonExcel.Filter.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("val" === reader.GetName()) {
				val = reader.GetValue();
				this.Val = val;
			}
		}
	};

	AscCommonExcel.Filter.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"filter"*/);

		writer.WriteXmlNullableAttributeStringEncode("val", this.Val);
		writer.WriteXmlAttributesEnd();

		writer.WriteXmlNodeEnd(ns + name);

	};

	//AscCommonExcel.QueryTable
	AscCommonExcel.QueryTable.prototype.fromXml = function (reader) {
		if (!reader.ReadNextNode()) {
			return;
		}

		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("queryTableRefresh" === name) {
				this.queryTableRefresh = new AscCommonExcel.QueryTableRefresh();
				this.queryTableRefresh.fromXml(reader);
			} else if ("extLst" === name) {
			}
		}
	};

	AscCommonExcel.QueryTable.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("adjustColumnWidth" === reader.GetName()) {
				val = reader.GetValueBool();
				this.adjustColumnWidth = val;
			} else if ("applyAlignmentFormats" === reader.GetName()) {
				val = reader.GetValueBool();
				this.applyAlignmentFormats = val;
			} else if ("applyBorderFormats" === reader.GetName()) {
				val = reader.GetValueBool();
				this.applyBorderFormats = val;
			} else if ("applyFontFormats" === reader.GetName()) {
				val = reader.GetValueBool();
				this.applyFontFormats = val;
			} else if ("applyNumberFormats" === reader.GetName()) {
				val = reader.GetValueBool();
				this.applyNumberFormats = val;
			} else if ("applyPatternFormats" === reader.GetName()) {
				val = reader.GetValueBool();
				this.applyPatternFormats = val;
			} else if ("applyWidthHeightFormats" === reader.GetName()) {
				val = reader.GetValueBool();
				this.applyWidthHeightFormats = val;
			} else if ("autoFormatId" === reader.GetName()) {
				val = reader.GetValue();
				this.autoFormatId = val;
			} else if ("backgroundRefresh" === reader.GetName()) {
				val = reader.GetValueBool();
				this.backgroundRefresh = val;
			} else if ("connectionId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.connectionId = val;
			} else if ("disableEdit" === reader.GetName()) {
				val = reader.GetValueBool();
				this.disableEdit = val;
			} else if ("disableRefresh" === reader.GetName()) {
				val = reader.GetValueBool();
				this.disableRefresh = val;
			} else if ("fillFormulas" === reader.GetName()) {
				val = reader.GetValueBool();
				this.fillFormulas = val;
			} else if ("firstBackgroundRefresh" === reader.GetName()) {
				val = reader.GetValueBool();
				this.firstBackgroundRefresh = val;
			} else if ("growShrinkType" === reader.GetName()) {
				val = reader.GetValue();
				this.growShrinkType = val;
			} else if ("headers" === reader.GetName()) {
				val = reader.GetValueBool();
				this.headers = val;
			} else if ("intermediate" === reader.GetName()) {
				val = reader.GetValueBool();
				this.intermediate = val;
			} else if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.name = val;
			} else if ("preserveFormatting" === reader.GetName()) {
				val = reader.GetValueBool();
				this.preserveFormatting = val;
			} else if ("refreshOnLoad" === reader.GetName()) {
				val = reader.GetValueBool();
				this.refreshOnLoad = val;
			} else if ("removeDataOnSave" === reader.GetName()) {
				val = reader.GetValueBool();
				this.removeDataOnSave = val;
			} else if ("rowNumbers" === reader.GetName()) {
				val = reader.GetValueBool();
				this.rowNumbers = val;
			}
		}
	};

	AscCommonExcel.QueryTable.prototype.toXml = function (writer) {
		writer.WriteXmlString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');

		writer.WriteXmlString("<queryTable \
xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" \
mc:Ignorable=\"xr16\" \
xmlns:xr16=\"http://schemas.microsoft.com/office/spreadsheetml/2017/revision16\"");

		writer.WriteXmlNullableAttributeStringEncode("name", this.name);
		writer.WriteXmlNullableAttributeBool2("backgroundRefresh", this.backgroundRefresh);
		writer.WriteXmlNullableAttributeNumber("connectionId", this.connectionId);
		writer.WriteXmlNullableAttributeNumber("autoFormatId", this.autoFormatId);
		writer.WriteXmlNullableAttributeBool2("adjustColumnWidth", this.adjustColumnWidth);
		writer.WriteXmlNullableAttributeBool2("applyBorderFormats", this.applyBorderFormats);
		writer.WriteXmlNullableAttributeBool2("applyFontFormats", this.applyFontFormats);
		writer.WriteXmlNullableAttributeBool2("applyNumberFormats", this.applyNumberFormats);
		writer.WriteXmlNullableAttributeBool2("applyPatternFormats", this.applyPatternFormats);
		writer.WriteXmlNullableAttributeBool2("applyWidthHeightFormats", this.applyWidthHeightFormats);
		writer.WriteXmlNullableAttributeBool2("applyAlignmentFormats", this.applyAlignmentFormats);

		writer.WriteXmlNullableAttributeBool2("disableEdit", this.disableEdit);
		writer.WriteXmlNullableAttributeBool2("disableRefresh", this.disableRefresh);
		writer.WriteXmlNullableAttributeBool2("fillFormulas", this.fillFormulas);
		writer.WriteXmlNullableAttributeBool2("firstBackgroundRefresh", this.firstBackgroundRefresh);

		writer.WriteXmlNullableAttributeStringEncode("growShrinkType", this.growShrinkType);

		writer.WriteXmlNullableAttributeBool2("headers", this.headers);
		writer.WriteXmlNullableAttributeBool2("intermediate", this.intermediate);
		writer.WriteXmlNullableAttributeBool2("preserveFormatting", this.preserveFormatting);
		writer.WriteXmlNullableAttributeBool2("refreshOnLoad", this.refreshOnLoad);
		writer.WriteXmlNullableAttributeBool2("removeDataOnSave", this.removeDataOnSave);
		writer.WriteXmlNullableAttributeBool2("rowNumbers", this.rowNumbers);

		writer.WriteXmlString(">");

		if (this.queryTableRefresh) {
			this.queryTableRefresh.toXml(writer);
		}
		/*if (m_oExtLst.IsInit()) {
			writer.WriteXmlString(m_oExtLst.toXMLWithNS(("")));
		}*/

		writer.WriteXmlString("</queryTable>");
	};

	AscCommonExcel.QueryTableRefresh.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var t = this;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("queryTableFields" === name) {
				reader.readXmlArray("queryTableField", function () {
					var queryTableField = new AscCommonExcel.QueryTableField();
					queryTableField.fromXml(reader);
					if (!t.queryTableFields) {
						t.queryTableFields = [];
					}
					t.queryTableFields.push(queryTableField);
				});
			} else if ("queryTableDeletedFields" === name) {
				reader.readXmlArray("queryTableDeletedField", function () {
					var queryTableDeletedField = new AscCommonExcel.QueryTableDeletedField();
					queryTableDeletedField.fromXml(reader);
					if (!t.queryTableDeletedFields) {
						t.queryTableDeletedFields = [];
					}
					t.queryTableDeletedFields.push(queryTableDeletedField);
				});
			} else if ("sortState" === name) {
				var sortState = new AscCommonExcel.SortState();
				sortState.fromXml(reader);
				this.sortState = sortState;
			} else if ("extLst" === name) {
			}
		}
	};

	AscCommonExcel.QueryTableRefresh.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("nextId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.nextId = val;
			} else if ("minimumVersion" === reader.GetName()) {
				val = reader.GetValueInt();
				this.minimumVersion = val;
			} else if ("fieldIdWrapped" === reader.GetName()) {
				val = reader.GetValueBool();
				this.fieldIdWrapped = val;
			} else if ("headersInLastRefresh" === reader.GetName()) {
				val = reader.GetValueBool();
				this.headersInLastRefresh = val;
			} else if ("preserveSortFilterLayout" === reader.GetName()) {
				val = reader.GetValueBool();
				this.preserveSortFilterLayout = val;
			} else if ("unboundColumnsLeft" === reader.GetName()) {
				val = reader.GetValueInt();
				this.unboundColumnsLeft = val;
			} else if ("unboundColumnsRight" === reader.GetName()) {
				val = reader.GetValueInt();
				this.unboundColumnsRight = val;
			}
		}
	};

	AscCommonExcel.QueryTableRefresh.prototype.toXml = function (writer) {
		writer.WriteXmlString("<queryTableRefresh");
		writer.WriteXmlNullableAttributeNumber("nextId", this.nextId);
		writer.WriteXmlNullableAttributeNumber("unboundColumnsLeft", this.unboundColumnsLeft);
		writer.WriteXmlNullableAttributeNumber("unboundColumnsRight", this.unboundColumnsRight);

		writer.WriteXmlNullableAttributeBool2("fieldIdWrapped", this.fieldIdWrapped);
		writer.WriteXmlNullableAttributeBool2("headersInLastRefresh", this.headersInLastRefresh);
		writer.WriteXmlNullableAttributeBool2("preserveSortFilterLayout", this.preserveSortFilterLayout);
		writer.WriteXmlNullableAttributeNumber("minimumVersion", this.minimumVersion);
		writer.WriteXmlString(">");

		if (this.queryTableFields) {
			writer.WriteXmlArray(this.queryTableFields, "queryTableField", "queryTableFields", true);
		}
		if (this.queryTableDeletedFields) {
			writer.WriteXmlArray(this.queryTableDeletedFields, "deletedField", "queryTableDeletedFields", true);
		}

		if (this.sortState) {
			this.sortState.toXml(writer);
		}

		writer.WriteXmlString("</queryTableRefresh>");
	};

	AscCommonExcel.QueryTableField.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("extLst" === name) {
			}
		}
	};

	AscCommonExcel.QueryTableField.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("id" === reader.GetName()) {
				val = reader.GetValueInt();
				this.id = val;
			} else if ("tableColumnId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.tableColumnId = val;
			} else if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.name = val;
			} else if ("rowNumbers" === reader.GetName()) {
				val = reader.GetValueBool();
				this.rowNumbers = val;
			} else if ("fillFormulas" === reader.GetName()) {
				val = reader.GetValueBool();
				this.fillFormulas = val;
			} else if ("dataBound" === reader.GetName()) {
				val = reader.GetValueBool();
				this.dataBound = val;
			} else if ("clipped" === reader.GetName()) {
				val = reader.GetValueBool();
				this.clipped = val;
			}
		}
	};

	AscCommonExcel.QueryTableField.prototype.toXml = function (writer, name) {
		//queryTableField
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("name", this.name);
		writer.WriteXmlNullableAttributeNumber("id", this.id);
		writer.WriteXmlNullableAttributeNumber("tableColumnId", this.tableColumnId);

		writer.WriteXmlNullableAttributeBool2("rowNumbers", this.rowNumbers);
		writer.WriteXmlNullableAttributeBool2("fillFormulas", this.fillFormulas);
		writer.WriteXmlNullableAttributeBool2("dataBound", this.dataBound);
		writer.WriteXmlNullableAttributeBool2("clipped", this.clipped);
		writer.WriteXmlAttributesEnd(true);
	};

	AscCommonExcel.QueryTableDeletedField.prototype.fromXml = function (reader) {

		/*ReadAttributes( oReader );
				if ( !oReader.IsEmptyNode() )
					oReader.ReadTillEnd();*/

		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	AscCommonExcel.QueryTableDeletedField.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.name = val;
			}
		}
	};

	AscCommonExcel.QueryTableDeletedField.prototype.toXml = function (writer, name) {
		//deletedField
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("name", this.name);
		writer.WriteXmlAttributesEnd(true);
	};


	//****data validation****
	AscCommonExcel.CDataValidations.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("dataValidation" === name) {
				val = new AscCommonExcel.CDataValidation();
				val.fromXml(reader);
				this.elems.push(val);
			}
		}
	};

	AscCommonExcel.CDataValidations.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("disablePrompts" === reader.GetName()) {
				val = reader.GetValueBool();
				this.disablePrompts = val;
			} else if ("xWindow" === reader.GetName()) {
				val = reader.GetValueInt();
				this.xWindow = val;
			} else if ("yWindow" === reader.GetName()) {
				val = reader.GetValueInt();
				this.yWindow = val;
			}
		}
	};


	AscCommonExcel.CDataValidations.prototype.toXml = function (writer, bExtendedWrite) {
		if (!this.elems || !this.elems.length) {
			return;
		}

		//на данном этапе разделяю что записывать в ext, а что нет
		//если массив для записи пустой, пропускаем
		var writeMap = null;
		var i;
		for (i = 0; i < this.elems.length; ++i) {
			var isExtended = this.elems[i].isExtended();
			if ((bExtendedWrite && isExtended) || (!bExtendedWrite && !isExtended)) {
				if (!writeMap) {
					writeMap = {};
				}
				writeMap[i] = 1;
			}
		}

		if (!writeMap) {
			return;
		}

		var node_name = bExtendedWrite ? "x14:dataValidations" : "dataValidations";
		writer.WriteXmlString("<" + node_name);
		if (bExtendedWrite) {
			writer.WriteXmlAttributeString("xmlns:xm", "http://schemas.microsoft.com/office/excel/2006/main");
		}

		writer.WriteXmlNullableAttributeNumber("count", this.elems.length);
		writer.WriteXmlNullableAttributeNumber("disablePrompts", boolToNumber(this.disablePrompts));
		writer.WriteXmlNullableAttributeNumber("xWindow", this.xWindow);
		writer.WriteXmlNullableAttributeNumber("yWindow", this.yWindow);
		writer.WriteXmlAttributesEnd();

		for (i = 0; i < this.elems.length; ++i) {
			if (writeMap[i]) {
				this.elems[i].toXml(writer, bExtendedWrite);
			}
		}

		writer.WriteXmlString("</" + node_name + ">");
	};


	AscCommonExcel.CDataValidation.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("formula1" === name) {
				val = new Asc.CDataFormula(prepareTextFromXml(reader.GetTextDecodeXml()));

				this.formula1 = val;
			} else if ("formula2" === name) {
				val = new Asc.CDataFormula(prepareTextFromXml(reader.GetTextDecodeXml()));
				this.formula2 = val;
			} else if ("sqref" === name) {
				this.setSqRef(reader.GetText());
			}
			//--------------------------------------------------- xml spreadsheet 2002
			else if ("Range" === name) {
				/*r1c1_formula_convert::base_row = 1;
				r1c1_formula_convert::base_col = 1;

				r1c1_formula_convert convert;

				m_oSqRef = convert.convert(oReader.GetText2());*/
				//TODO r1c1? не могу найти в каком виде она должна быть здесь записана
				val = reader.GetText()
				this.setSqRef(val);
			} else if ("Type" === name) {

				val = reader.GetText();
				this.type = FromXml_ST_DataValidationType(val);
				this.allowBlank = true;
				this.showInputMessage = true;
			} else if ("Value" === name) {
				/*r1c1_formula_convert::base_row = 1;
				r1c1_formula_convert::base_col = 1;

				r1c1_formula_convert convert;

				m_oFormula1 = new CDataValidationFormula(m_pMainDocument);
				m_oFormula1->m_sText = convert.convert(oReader.GetText3());*/

				//TODO r1c1? не могу найти в каком виде она должна быть здесь записана
				val = new Asc.CDataFormula(prepareTextFromXml(reader.GetTextDecodeXml()));
				this.formula1 = val;
			}
		}

		/*void CDataValidation::fromXML(XmlUtils::CXmlLiteReader& oReader)
		{
			ReadAttributes( oReader );

			if ( oReader.IsEmptyNode() )
				return;

			int nCurDepth = oReader.GetDepth();
			while (oReader.ReadNextSiblingNode(nCurDepth))
			{
				std::wstring sName = XmlUtils::GetNameNoNS(oReader.GetName());
				if (L"formula1" == sName)
				{
					m_oFormula1 = oReader;
				}
			else if (L"formula2" == sName)
				{
					m_oFormula2 = oReader;
				}
			else if (L"sqref" == sName)
				{
					m_oSqRef = oReader.GetText2();
				}
//--------------------------------------------------- xml spreadsheet 2002
			else if (L"Range" == sName)
				{
					r1c1_formula_convert::base_row = 1;
					r1c1_formula_convert::base_col = 1;

					r1c1_formula_convert convert;

					m_oSqRef = convert.convert(oReader.GetText2());
				}
			else if (L"Type" == sName)
				{
					m_oType = oReader.GetText2();

					m_oAllowBlank.Init();
					m_oAllowBlank->FromBool(true);

					m_oShowInputMessage.Init();
					m_oShowInputMessage->FromBool(true);
				}
			else if (L"Value" == sName)
				{
					r1c1_formula_convert::base_row = 1;
					r1c1_formula_convert::base_col = 1;

					r1c1_formula_convert convert;

					m_oFormula1 = new CDataValidationFormula(m_pMainDocument);
					m_oFormula1->m_sText = convert.convert(oReader.GetText3());

					//if (m_oFormula1->m_sText.find(L"!") == std::wstring::npos)
					//{
					//	CXlsxFlat* xlsx_flat = dynamic_cast<CXlsxFlat*>(m_pMainDocument);
					//	if (xlsx_flat)
					//	{
					//		CSheet *pSheet = xlsx_flat->m_pWorkbook->m_oSheets->m_arrItems.back();
					//		if (pSheet->m_oName.IsInit())
					//		{
					//			m_oFormula1->m_sText = *pSheet->m_oName + L"!" + m_oFormula1->m_sText;
					//		}
					//	}
					//}
				}
			}
		}*/
	};

	AscCommonExcel.CDataValidation.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("allowBlank" === reader.GetName()) {
				val = reader.GetValueBool();
				this.allowBlank = val;
			} else if ("error" === reader.GetName()) {
				val = reader.GetValue();
				this.error = val;
			} else if ("errorStyle" === reader.GetName()) {
				val = reader.GetValue();
				this.errorStyle = FromXml_ST_DataValidationErrorStyle(val);
			} else if ("errorTitle" === reader.GetName()) {
				val = reader.GetValue();
				this.errorTitle = val;
			} else if ("imeMode" === reader.GetName()) {
				val = reader.GetValue();
				this.imeMode = FromXml_ST_DataValidationImeMode(val);
			} else if ("operator" === reader.GetName()) {
				val = reader.GetValue();
				this.operator = FromXml_ST_DataValidationOperator(val);
			} else if ("prompt" === reader.GetName()) {
				val = reader.GetValue();
				this.prompt = val;
			} else if ("promptTitle" === reader.GetName()) {
				val = reader.GetValue();
				this.promptTitle = val;
			} else if ("showDropDown" === reader.GetName()) {
				val = reader.GetValueBool();
				this.showDropDown = val;
			} else if ("showErrorMessage" === reader.GetName()) {
				val = reader.GetValueBool();
				this.showErrorMessage = val;
			} else if ("showInputMessage" === reader.GetName()) {
				val = reader.GetValueBool();
				this.showInputMessage = val;
			} else if ("sqref" === reader.GetName()) {
				val = reader.GetValue();
				this.setSqRef(val);
			} else if ("type" === reader.GetName()) {
				val = reader.GetValue();
				this.type = FromXml_ST_DataValidationType(val);
			}
		}
	};


	//COfficeArtExtensionList

	/*virtual void fromXML(XmlUtils::CXmlLiteReader& oReader)
	{
		if ( oReader.IsEmptyNode() )
			return;

		int nCurDepth = oReader.GetDepth();
		while( oReader.ReadNextSiblingNode( nCurDepth ) )
		{
			std::wstring sName = XmlUtils::GetNameNoNS(oReader.GetName());
			if ( _T("ext") == sName )
			{
				OOX::Drawing::COfficeArtExtension *oExt = new OOX::Drawing::COfficeArtExtension(oReader);
				if (oExt) m_arrExt.push_back( oExt );
			}
		}
	}*/

	AscCommonExcel.CDataValidation.prototype.isExtended = function () {
		var result1 = true, result2 = true;
		if (this.formula1 && this.formula1.text) {
			if (this.formula1.text.indexOf("!") !== -1 && this.formula1.text.toUpperCase() !== "#REF!") {
				result1 = false;
			}
		} else {
			result1 = false;
		}

		if (this.formula2 && this.formula2.text) {
			if (this.formula2.text.indexOf("!") !== -1 && this.formula2.text.toUpperCase() !== "#REF!") {
				result1 = false;
			}
		} else {
			result2 = false;
		}

		return result1 || result2;
	};

	AscCommonExcel.CDataValidation.prototype.toXml = function (writer, bExtendedWrite) {
		//TODO HHHH
		var node_name = bExtendedWrite ? "x14:dataValidation" : "dataValidation";

		writer.WriteXmlString("<" + node_name);
		if (bExtendedWrite) {
			/*if (false == m_oUuid.IsInit())
			{
				m_oUuid = L"{" + XmlUtils::GenerateGuid() + L"}";
			}
			WritingStringNullableAttrString	(L"xr:uid",	m_oUuid, m_oUuid.get());*/

			writer.WriteXmlNullableAttributeString("xr:uid", AscCommon.CreateGUID());
		} else {
			writer.WriteXmlNullableAttributeString("sqref", AscCommonExcel.getSqRefString(this.ranges));
		}

		writer.WriteXmlNullableAttributeString("type", ToXml_ST_DataValidationType(this.type));
		writer.WriteXmlNullableAttributeNumber("allowBlank", boolToNumber(this.allowBlank));

		//WritingStringNullableAttrEncodeXmlStringHHHH
		writer.WriteXmlNullableAttributeStringEncode("error", this.error);

		writer.WriteXmlNullableAttributeString("errorStyle", ToXml_ST_DataValidationErrorStyle(this.errorStyle));

		//WritingStringNullableAttrEncodeXmlStringHHHH
		writer.WriteXmlNullableAttributeStringEncode("errorTitle", this.errorTitle);

		writer.WriteXmlNullableAttributeString("imeMode", ToXml_ST_DataValidationImeMode(this.imeMode));
		writer.WriteXmlNullableAttributeString("operator", ToXml_ST_DataValidationOperator(this.operator));

		//WritingStringNullableAttrEncodeXmlStringHHHH
		writer.WriteXmlNullableAttributeStringEncode("prompt", this.prompt);
		writer.WriteXmlNullableAttributeStringEncode("promptTitle", this.promptTitle);

		writer.WriteXmlNullableAttributeNumber("showDropDown", boolToNumber(this.showDropDown));
		writer.WriteXmlNullableAttributeNumber("showErrorMessage", boolToNumber(this.showErrorMessage));
		writer.WriteXmlNullableAttributeNumber("showInputMessage", boolToNumber(this.showInputMessage));

		writer.WriteXmlAttributesEnd();


		if (bExtendedWrite) {
			var node_formula_name = bExtendedWrite ? "xm:f" : "formula";
			if (this.formula1) {
				writer.WriteXmlString("<x14:formula1>");

				writer.WriteXmlString("<" + node_formula_name + ">");
				writer.WriteXmlStringEncode(this.formula1.text);
				writer.WriteXmlString("</" + node_formula_name + ">");

				writer.WriteXmlString("</x14:formula1>");
			}
			if (this.formula2) {
				writer.WriteXmlString("<x14:formula2>");

				writer.WriteXmlString("<" + node_formula_name + ">");
				writer.WriteXmlStringEncode(this.formula2.text);
				writer.WriteXmlString("</" + node_formula_name + ">");

				writer.WriteXmlString("</x14:formula2>");
			}
			if (this.ranges) {
				writer.WriteXmlString("<xm:sqref>" + AscCommonExcel.getSqRefString(this.ranges) + "</xm:sqref>");
			}
		} else {
			if (this.formula1) {
				writer.WriteXmlString("<formula1>");
				writer.WriteXmlString(this.formula1.text);
				writer.WriteXmlString("</formula1>");
			}
			if (this.formula2) {
				writer.WriteXmlString("<formula2>");
				writer.WriteXmlString(this.formula2.text);
				writer.WriteXmlString("</formula2>");
			}
		}
		writer.WriteXmlString("</" + node_name + ">");
	};

	Asc.asc_CPageSetup.prototype.fromXml = function (reader) {

		/*ReadAttributes( oReader );

						if ( !oReader.IsEmptyNode() )
							oReader.ReadTillEnd();*/

		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	Asc.asc_CPageSetup.prototype.readAttr = function (reader) {
		var paperUnits, paperHeight, paperWidth;
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("blackAndWhite" === reader.GetName()) {
				val = reader.GetValueBool();
				this.blackAndWhite = val;
			} else if ("cellComments" === reader.GetName()) {
				val = reader.GetValue();
				this.cellComments = FromXml_ST_CellComments(val);
			} else if ("copies" === reader.GetName()) {
				val = reader.GetValueInt();
				this.copies = val;
			} else if ("draft" === reader.GetName()) {
				val = reader.GetValueBool();
				this.draft = val;
			} else if ("errors" === reader.GetName()) {
				val = reader.GetValue();
				this.errors = FromXml_ST_PrintError(val);
			} else if ("firstPageNumber" === reader.GetName()) {
				val = reader.GetValueInt();
				this.firstPageNumber = val;
			} else if ("fitToHeight" === reader.GetName()) {
				val = reader.GetValueInt();
				this.fitToHeight = val;
			} else if ("fitToWidth" === reader.GetName()) {
				val = reader.GetValueInt();
				this.fitToWidth = val;
			} else if ("horizontalDpi" === reader.GetName()) {
				val = reader.GetValueInt();
				this.horizontalDpi = val;
			} /*else if ("r:id" === reader.GetName()) {
				val = reader.GetValue();
				this.r:id = val;
			} else if ("relationships:id" === reader.GetName()) {
				val = reader.GetValue();
				this.relationships:id = val;
			}*/ else if ("orientation" === reader.GetName()) {
				var byteFormatOrientation = reader.GetValue();
				var byteOrientation = FromXml_ST_PageOrientation(byteFormatOrientation);
				if (null != byteOrientation) {
					this.asc_setOrientation(byteOrientation);
				}
			} else if ("pageOrder" === reader.GetName()) {
				//ST_PageOrder
				val = reader.GetValue();
				this.pageOrder = FromXml_ST_PageOrder(val);
			} else if ("paperHeight" === reader.GetName()) {
				paperHeight = reader.GetValue();
				this.height = paperHeight;
			} else if ("paperSize" === reader.GetName()) {
				var bytePaperSize = reader.GetValueInt();
				var item = AscCommonExcel.DocumentPageSize.getSizeById(bytePaperSize);
				this.asc_setWidth(item.w_mm);
				this.asc_setHeight(item.h_mm);
			} else if ("paperWidth" === reader.GetName()) {
				paperWidth = reader.GetValue();
				this.width = paperWidth;
			} else if ("paperUnits" === reader.GetName()) {
				val = reader.GetValue();
				paperUnits = val;
			} else if ("scale" === reader.GetName()) {
				val = reader.GetValueInt();
				this.scale = val;
			} else if ("useFirstPageNumber" === reader.GetName()) {
				val = reader.GetValueBool();
				this.useFirstPageNumber = val;
			} else if ("usePrinterDefaults" === reader.GetName()) {
				val = reader.GetValueBool();
				this.usePrinterDefaults = val;
			} else if ("verticalDpi" === reader.GetName()) {
				val = reader.GetValueInt();
				this.verticalDpi = val;
			}
		}

		//paperUnits - ST_PositiveUniversalMeasure. пробовал в мс задавать разные метрики, в тч и пример из документации
		//при открытии скидывает на дефолтовый первый тип - letter, поскольку мс никак не реагирует - не обрабатываю

		//When paperHeight, paperWidth, and paperUnits are specified, paperSize should be ignored.
		if (paperUnits && paperHeight && paperWidth) {
			//this.width = AscCommon.universalMeasureToMm(paperWidth + paperUnits, 1, 0);
			//this.height = AscCommon.universalMeasureToMm(paperHeight + paperUnits, 1, 0);
		}
	};

	Asc.asc_CPageSetup.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);

		/*<xsd:simpleType name="ST_PositiveUniversalMeasure">
			94 <xsd:restriction base="ST_UniversalMeasure">
			95 <xsd:pattern value="[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)"/>
			96 </xsd:restriction>
			97 </xsd:simpleType>*/

		var isWritePaperSize;
		var dWidth = this.asc_getWidth();
		var dHeight = this.asc_getHeight();
		if (null != dWidth && null != dHeight) {
			var item = AscCommonExcel.DocumentPageSize.getSizeByWH(dWidth, dHeight);
			writer.WriteXmlNullableAttributeString("paperSize", item.id + "");
			isWritePaperSize = true;
		}

		//не записываю повторно, если уже есть paperSize
		if (!isWritePaperSize) {
			writer.WriteXmlNullableAttributeDouble("paperHeight", dHeight);//ST_PositiveUniversalMeasure
			writer.WriteXmlNullableAttributeDouble("paperWidth", dWidth);//ST_PositiveUniversalMeasure
		}

		writer.WriteXmlNullableAttributeUInt("scale", this.scale);
		if (-1 !== this.firstPageNumber) {
			writer.WriteXmlNullableAttributeUInt("firstPageNumber", this.firstPageNumber);
		}
		writer.WriteXmlNullableAttributeNumber("fitToWidth", this.fitToWidth);
		writer.WriteXmlNullableAttributeNumber("fitToHeight", this.fitToHeight);

		//ST_PageOrder
		writer.WriteXmlNullableAttributeString("pageOrder", ToXml_ST_PageOrder(this.pageOrder));

		writer.WriteXmlNullableAttributeString("orientation", ToXml_ST_PageOrientation(this.asc_getOrientation()));

		writer.WriteXmlNullableAttributeBool("usePrinterDefaults", this.usePrinterDefaults);
		writer.WriteXmlNullableAttributeBool("blackAndWhite", this.blackAndWhite);
		writer.WriteXmlNullableAttributeBool("draft", this.draft);

		//ST_CellComments
		writer.WriteXmlNullableAttributeString("cellComments", ToXml_ST_CellComments(this.cellComments));

		writer.WriteXmlNullableAttributeBool("useFirstPageNumber", this.useFirstPageNumber);


		writer.WriteXmlNullableAttributeString("errors", ToXml_ST_PrintError(this.errors));

		writer.WriteXmlNullableAttributeUInt("horizontalDpi", this.horizontalDpi);
		writer.WriteXmlNullableAttributeUInt("verticalDpi", this.verticalDpi);
		writer.WriteXmlNullableAttributeUInt("copies", this.copies);
		//writer.WriteXmlNullableAttributeString("paperUnits", this.paperUnits);
		//writer.WriteXmlNullableAttributeString("r:id", this.r:id);
		writer.WriteXmlAttributesEnd(true);
	};

	Asc.asc_CPageMargins.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	Asc.asc_CPageMargins.prototype.readAttr = function (reader) {
		//TODO баг 21685 - при получении значений -  необходимо грамотно округлять

		var val;
		while (reader.MoveToNextAttribute()) {
			if ("left" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.asc_setLeft(val * AscCommonWord.g_dKoef_in_to_mm);
			} else if ("top" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.asc_setTop(val * AscCommonWord.g_dKoef_in_to_mm);
			} else if ("right" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.asc_setRight(val * AscCommonWord.g_dKoef_in_to_mm);
			} else if ("bottom" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.asc_setBottom(val * AscCommonWord.g_dKoef_in_to_mm);
			} else if ("header" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.asc_setHeader(val * AscCommonWord.g_dKoef_in_to_mm);
			} else if ("footer" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.asc_setFooter(val * AscCommonWord.g_dKoef_in_to_mm);
			}
		}
	};

	Asc.asc_CPageMargins.prototype.toXml = function (writer, name, ns) {
		//TODO баг 21685 - при получении значений -  необходимо грамотно округлять
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);

		writer.WriteXmlNullableAttributeDouble("left", this.left / AscCommonWord.g_dKoef_in_to_mm);
		writer.WriteXmlNullableAttributeDouble("right", this.right / AscCommonWord.g_dKoef_in_to_mm);
		writer.WriteXmlNullableAttributeDouble("top", this.top / AscCommonWord.g_dKoef_in_to_mm);
		writer.WriteXmlNullableAttributeDouble("bottom", this.bottom / AscCommonWord.g_dKoef_in_to_mm);
		writer.WriteXmlNullableAttributeDouble("header", this.header / AscCommonWord.g_dKoef_in_to_mm);
		writer.WriteXmlNullableAttributeDouble("footer", this.footer / AscCommonWord.g_dKoef_in_to_mm);

		writer.WriteXmlAttributesEnd(true);
	};


	Asc.asc_CPageOptions.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	Asc.asc_CPageOptions.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("gridLines" === reader.GetName()) {
				val = reader.GetValueBool();
				this.asc_setGridLines(val);
			} /*else if ("gridLinesSet" === reader.GetName()) {
				val = reader.GetValueBool();
				this.gridLinesSet = val;
			}*/ else if ("headings" === reader.GetName()) {
				val = reader.GetValueBool();
				this.asc_setHeadings(val);
			} /*else if ("horizontalCentered" === reader.GetName()) {
				val = reader.GetValueBool();
				this.horizontalCentered = val;
			} else if ("verticalCentered" === reader.GetName()) {
				val = reader.GetValueBool();
				this.verticalCentered = val;
			}*/
		}
	};

	Asc.asc_CPageOptions.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeBool("headings", this.headings);
		writer.WriteXmlNullableAttributeBool("gridLines", this.gridLines);
		//writer.WriteXmlNullableAttributeBool("gridLinesSet", this.gridLinesSet);
		writer.WriteXmlAttributesEnd(true);
	};

	AscCommonExcel.SheetFormatPr.prototype.fromXml = function (reader, oWorksheet) {
		this.readAttr(reader, oWorksheet);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	AscCommonExcel.SheetFormatPr.prototype.readAttr = function (reader, oWorksheet) {
		//TODO в x2t ветки с использованием xlsx_flat
		var val;
		var oAllRow;
		while (reader.MoveToNextAttribute()) {
			if ("baseColWidth" === reader.GetName()) {
				val = reader.GetValueInt();
				this.nBaseColWidth = val;
			} else if ("customHeight" === reader.GetName()) {
				oAllRow = oWorksheet.getAllRow();
				var CustomHeight = reader.GetValueBool();
				if (CustomHeight) {
					oAllRow.setCustomHeight(true);
				}
			} else if ("defaultColWidth" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.dDefaultColWidth = val;
			} else if ("defaultRowHeight" === reader.GetName()) {
				oAllRow = oWorksheet.getAllRow();
				oAllRow.setHeight(reader.GetValueDouble());
			} else if ("outlineLevelCol" === reader.GetName()) {
				val = reader.GetValueInt();
				this.nOutlineLevelCol = val;
			} else if ("outlineLevelRow" === reader.GetName()) {
				oAllRow = oWorksheet.getAllRow();
				oAllRow.setOutlineLevel(reader.GetValueInt());
			} /*else if ("thickBottom" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ThickBottom = val;
			} else if ("thickTop" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ThickTop = val;
			}*/ else if ("zeroHeight" === reader.GetName()) {
				oAllRow = oWorksheet.getAllRow();
				var hd = reader.GetValueBool();
				if (hd) {
					oAllRow.setHidden(true);
				}
			}
		}
	};

	AscCommonExcel.SheetFormatPr.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		//в x2t все функции с окончанием 2
		var oAllRow = this.oAllRow;
		writer.WriteXmlNodeStart(ns + name/*"sheetFormatPr"*/);
		writer.WriteXmlNullableAttributeNumber("baseColWidth", this.nBaseColWidth);
		writer.WriteXmlNullableAttributeBool("customHeight", oAllRow.getCustomHeight() ? true : null);
		writer.WriteXmlNullableAttributeDouble("defaultColWidth", this.dDefaultColWidth);
		writer.WriteXmlNullableAttributeDouble("defaultRowHeight", oAllRow.h ? oAllRow.h : null);
		writer.WriteXmlNullableAttributeNumber("outlineLevelCol", oAllRow.nOutlineLevelCol > 0 ? oAllRow.nOutlineLevelCol : null);
		writer.WriteXmlNullableAttributeNumber("outlineLevelRow", oAllRow.getOutlineLevel() > 0 ? true : null);
		//writer.WriteXmlNullableAttributeBool("thickBottom", this.thickBottom);
		//writer.WriteXmlNullableAttributeBool("thickTop", this.thickTop);
		writer.WriteXmlNullableAttributeBool("zeroHeight", oAllRow.getHidden() ? true : null);
		writer.WriteXmlAttributesEnd(true);
	};

	AscCommonExcel.asc_CSheetViewSettings.prototype.fromXml = function (reader, ws) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("pane" === name) {
				this.pane = new AscCommonExcel.asc_CPane();
				this.pane.fromXml(reader);

				this.pane.init();
			} else if ("selection" === name) {
				//AscCommonExcel.SelectionRange
				ws.selectionRange.clean();
				ws.selectionRange.fromXml(reader);
				ws.selectionRange.update();
			}
		}
	};

	AscCommonExcel.asc_CSheetViewSettings.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("colorId" === reader.GetName()) {
				//val = reader.GetValueInt();
				//this.colorId = val;
			} else if ("defaultGridColor" === reader.GetName()) {
				//val = reader.GetValueBool();
				//this.defaultGridColor = val;
			} else if ("rightToLeft" === reader.GetName()) {
				//val = reader.GetValueBool();
				//this.rightToLeft = val;
			} else if ("showFormulas" === reader.GetName()) {
				//val = reader.GetValueBool();
				//this.showFormulas = val;
			} else if ("showGridLines" === reader.GetName()) {
				val = reader.GetValueBool();
				this.showGridLines = val;
			} else if ("showOutlineSymbols" === reader.GetName()) {
				//val = reader.GetValueBool();
				//this.showOutlineSymbols = val;
			} else if ("showRowColHeaders" === reader.GetName()) {
				val = reader.GetValueBool();
				this.showRowColHeaders = val;
			} else if ("showRuler" === reader.GetName()) {
				//val = reader.GetValueBool();
				//this.showRuler = val;
			} else if ("showWhiteSpace" === reader.GetName()) {
				//val = reader.GetValueBool();
				//this.showWhiteSpace = val;
			} else if ("showZeros" === reader.GetName()) {
				val = reader.GetValueBool();
				this.showZeros = val;
			} else if ("tabSelected" === reader.GetName()) {
				//val = reader.GetValueBool();
				//this.tabSelected = val;
			} else if ("topLeftCell" === reader.GetName()) {
				var _topLeftCell = AscCommonExcel.g_oRangeCache.getAscRange(reader.GetValue());
				if (_topLeftCell) {
					this.topLeftCell = new Asc.Range(_topLeftCell.c1, _topLeftCell.r1, _topLeftCell.c1, _topLeftCell.r1);
				}
			} else if ("view" === reader.GetName()) {
				//val = reader.GetValue();
				//this.View = val;
			} else if ("windowProtection" === reader.GetName()) {
				//val = reader.GetValueBool();
				//this.WindowProtection = val;
			} else if ("workbookViewId" === reader.GetName()) {
				//val = reader.GetValueInt();
				//this.WorkbookViewId = val;
			} else if ("zoomScale" === reader.GetName()) {
				val = reader.GetValueInt();
				this.zoomScale = val;
			} else if ("zoomScaleNormal" === reader.GetName()) {
				//val = reader.GetValueInt();
				//this.ZoomScaleNormal = val;
			} else if ("zoomScalePageLayoutView" === reader.GetName()) {
				//val = reader.GetValueInt();
				//this.ZoomScalePageLayoutView = val;
			} else if ("zoomScaleSheetLayoutView" === reader.GetName()) {
				//val = reader.GetValueInt();
				//this.ZoomScaleSheetLayoutView = val;
			}
		}
	};

	AscCommonExcel.asc_CSheetViewSettings.prototype.toXml = function (writer, name, ns) {
		var ws = writer.context.ws;
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		//writer.WriteXmlNullableAttributeNumber("colorId", this.colorId);
		//writer.WriteXmlNullableAttributeBool("defaultGridColor", this.defaultGridColor);
		//writer.WriteXmlNullableAttributeBool("rightToLeft", this.rightToLeft);
		//writer.WriteXmlNullableAttributeBool("showFormulas", this.showFormulas);
		writer.WriteXmlNullableAttributeBool("showGridLines", this.showGridLines);
		//writer.WriteXmlNullableAttributeBool("showOutlineSymbols", this.showOutlineSymbols);
		writer.WriteXmlNullableAttributeBool("showRowColHeaders", this.showRowColHeaders);
		//writer.WriteXmlNullableAttributeBool("showRuler", this.showRuler);
		//writer.WriteXmlNullableAttributeBool("showWhiteSpace", this.showWhiteSpace);
		writer.WriteXmlNullableAttributeBool("showZeros", this.showZeros);
		//writer.WriteXmlNullableAttributeBool("tabSelected", this.tabSelected);
		writer.WriteXmlNullableAttributeString("topLeftCell", this.topLeftCell ? this.topLeftCell.getName() : null);
		//writer.WriteXmlNullableAttributeString("view", this.view);
		//writer.WriteXmlNullableAttributeBool("windowProtection", this.windowProtection);
		writer.WriteXmlNullableAttributeNumber("zoomScale", this.zoomScale);

		//в x2t предварительная обработка -> если не определено значение, присваиваем 0. в данном случае эта опция не поддерживается в редакторе, всегда ноль
		writer.WriteXmlNullableAttributeNumber("workbookViewId", 0);

		//writer.WriteXmlNullableAttributeNumber("zoomScaleNormal", this.zoomScaleNormal);
		//writer.WriteXmlNullableAttributeNumber("zoomScalePageLayoutView", this.zoomScalePageLayoutView);
		//writer.WriteXmlNullableAttributeNumber("zoomScaleSheetLayoutView", this.zoomScaleSheetLayoutView);
		writer.WriteXmlAttributesEnd();

		//AscCommonExcel.asc_CPane
		if (this.pane) {
			this.pane.toXml(writer, "pane");
		}
		//AscCommonExcel.SelectionRange
		if (ws.selectionRange) {
			ws.selectionRange.toXml(writer, "selection");
		}

		writer.WriteXmlNodeEnd(ns + name);
	};


	AscCommonExcel.asc_CPane.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	AscCommonExcel.asc_CPane.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("activePane" === reader.GetName()) {
				val = reader.GetValue();
				this.activePane = val;
			} else if ("state" === reader.GetName()) {
				val = reader.GetValue();
				this.state = val;
			} else if ("topLeftCell" === reader.GetName()) {
				val = reader.GetValue();
				this.topLeftCell = val;
			} else if ("xSplit" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.xSplit = val;
			} else if ("ySplit" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.ySplit = val;
			}
		}
	};

	AscCommonExcel.asc_CPane.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeString("activePane", this.activePane);
		writer.WriteXmlNullableAttributeString("state", this.state);
		writer.WriteXmlNullableAttributeString("topLeftCell", this.topLeftCell);
		writer.WriteXmlNullableAttributeNumber("xSplit", this.xSplit);
		writer.WriteXmlNullableAttributeNumber("ySplit", this.ySplit);
		writer.WriteXmlAttributesEnd(true);
	};

	AscCommonExcel.SelectionRange.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	AscCommonExcel.SelectionRange.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("activeCell" === reader.GetName()) {
				var activeCell = AscCommonExcel.g_oRangeCache.getAscRange(reader.GetValue());
				if (activeCell) {
					this.activeCell = new AscCommon.CellBase(activeCell.r1, activeCell.c1);
				}
			} else if ("activeCellId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.activeCellId = val;
			} else if ("sqref" === reader.GetName()) {
				var sqRef = reader.GetValue();
				var selectionNew = AscCommonExcel.g_oRangeCache.getRangesFromSqRef(sqRef);
				if (selectionNew.length > 0) {
					this.ranges = selectionNew;
				}
			} else if ("pane" === reader.GetName()) {
				//val = reader.GetValue();
				//this.pane = val;
			}
		}
	};

	AscCommonExcel.SelectionRange.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeString("activeCell", this.activeCell ? this.activeCell.getName() : null);
		writer.WriteXmlNullableAttributeNumber("activeCellId", this.activeCellId);
		//writer.WriteXmlNullableAttributeString("pane", this.pane);
		writer.WriteXmlNullableAttributeString("sqref", this.ranges ? AscCommonExcel.getSqRefString(this.ranges) : null);
		writer.WriteXmlAttributesEnd(true);
	};

	Asc.CProtectedRange.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("securityDescriptor" === name) {
				if (!this.securityDescriptors) {
					this.securityDescriptors = [];
				}
				this.securityDescriptors.push(prepareTextFromXml(reader.GetTextDecodeXml()));
			}
		}
	};

	Asc.CProtectedRange.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("algorithmName" === reader.GetName()) {
				val = reader.GetValue();
				this.algorithmName = AscCommonExcel.FromXml_ST_AlgorithmName(val);
			} else if ("hashValue" === reader.GetName()) {
				val = reader.GetValue();
				this.hashValue = val;
			} else if ("saltValue" === reader.GetName()) {
				val = reader.GetValue();
				this.saltValue = val;
			} else if ("spinCount" === reader.GetName()) {
				val = reader.GetValueInt();
				this.spinCount = val;
			} else if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.name = val;
			} else if ("sqref" === reader.GetName()) {
				var sqRef = reader.GetValue();
				var newSqRef = AscCommonExcel.g_oRangeCache.getRangesFromSqRef(sqRef);
				if (newSqRef.length > 0) {
					this.sqref = newSqRef;
				}
			} else if ("securityDescriptor" === reader.GetName()) {
				val = reader.GetValue();
				if (!this.securityDescriptors) {
					this.securityDescriptors = [];
				}
				this.securityDescriptors.push(val);
			}
		}
	};

	Asc.CProtectedRange.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeString("name", this.name);
		writer.WriteXmlNullableAttributeString("sqref", AscCommonExcel.getSqRefString(this.sqref));
		writer.WriteXmlNullableAttributeString("algorithmName", AscCommonExcel.ToXml_ST_AlgorithmName(this.algorithmName));
		writer.WriteXmlNullableAttributeString("hashValue", this.hashValue);
		writer.WriteXmlNullableAttributeString("saltValue", this.saltValue);
		writer.WriteXmlNullableAttributeNumber("spinCount", this.spinCount);

		if (this.securityDescriptors && this.securityDescriptors.length === 1) {
			writer.WriteXmlNullableAttributeString("securityDescriptor",);
			writer.WriteXmlAttributeStringEncode("securityDescriptor", this.securityDescriptors[0]);
		}
		if (this.securityDescriptors && this.securityDescriptors.length > 1) {
			writer.WriteXmlString(">");
			for (var i = 0; i < this.securityDescriptors.length; ++i) {
				writer.WriteXmlString("<securityDescriptor>");
				writer.WriteXmlStringEncode(this.securityDescriptors[i]);
				writer.WriteXmlString("</securityDescriptor>");
			}
			writer.WriteXmlNodeEnd(ns + name);
		} else {
			writer.WriteXmlAttributesEnd(true);
		}
	};

	Asc.CHeaderFooter.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		var val;
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("evenFooter" === name) {
				val = prepareTextFromXml(reader.GetTextDecodeXml());
				if (val) {
					this.setEvenFooter(val);
				}
			} else if ("evenHeader" === name) {
				val = prepareTextFromXml(reader.GetTextDecodeXml());
				if (val) {
					this.setEvenHeader(val);
				}
			} else if ("firstFooter" === name) {
				val = prepareTextFromXml(reader.GetTextDecodeXml());
				if (val) {
					this.setFirstFooter(val);
				}
			} else if ("firstHeader" === name) {
				val = prepareTextFromXml(reader.GetTextDecodeXml());
				if (val) {
					this.setFirstHeader(val);
				}
			} else if ("oddFooter" === name) {
				val = prepareTextFromXml(reader.GetTextDecodeXml());
				if (val) {
					this.setOddFooter(val);
				}
			} else if ("oddHeader" === name) {
				val = prepareTextFromXml(reader.GetTextDecodeXml());
				if (val) {
					this.setOddHeader(val);
				}
			}
		}
	};

	Asc.CHeaderFooter.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("alignWithMargins" === reader.GetName()) {
				val = reader.GetValueBool();
				this.setAlignWithMargins(val);
			} else if ("differentFirst" === reader.GetName()) {
				val = reader.GetValueBool();
				this.setDifferentFirst(val);
			} else if ("differentOddEven" === reader.GetName()) {
				val = reader.GetValueBool();
				this.setDifferentOddEven(val);
			} else if ("scaleWithDoc" === reader.GetName()) {
				val = reader.GetValueBool();
				this.setScaleWithDoc(val);
			}
		}
	};

	Asc.CHeaderFooter.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeBool("alignWithMargins", this.alignWithMargins);
		writer.WriteXmlNullableAttributeBool("differentFirst", this.differentFirst);
		writer.WriteXmlNullableAttributeBool("differentOddEven", this.differentOddEven);
		writer.WriteXmlNullableAttributeBool("scaleWithDoc", this.scaleWithDoc);
		writer.WriteXmlAttributesEnd();


		if (this.oddHeader) {
			toXML2(writer, "oddHeader", this.oddHeader.str);
		}
		if (this.oddFooter) {
			toXML2(writer, "oddFooter", this.oddFooter.str);
		}
		if (this.evenHeader) {
			toXML2(writer, "evenHeader", this.evenHeader.str);
		}
		if (this.evenFooter) {
			toXML2(writer, "evenFooter", this.evenFooter.str);
		}
		if (this.firstHeader) {
			toXML2(writer, "firstHeader", this.firstHeader.str);
		}
		if (this.firstFooter) {
			toXML2(writer, "firstFooter", this.firstFooter.str);
		}
		writer.WriteXmlNodeEnd(ns + name);
	};

	Asc.CSheetProtection.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	Asc.CSheetProtection.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("password" === reader.GetName()) {
				val = reader.GetValue();
				this.password = val;
			} else if ("algorithmName" === reader.GetName()) {
				val = reader.GetValue();
				this.algorithmName = AscCommonExcel.FromXml_ST_AlgorithmName(val);
			} else if ("hashValue" === reader.GetName()) {
				val = reader.GetValue();
				this.hashValue = val;
			} else if ("saltValue" === reader.GetName()) {
				val = reader.GetValue();
				this.saltValue = val;
			} else if ("spinCount" === reader.GetName()) {
				val = reader.GetValueInt();
				this.spinCount = val;
			} else if ("autoFilter" === reader.GetName()) {
				val = reader.GetValueBool();
				this.autoFilter = val;
			} else if ("content" === reader.GetName()) {
				val = reader.GetValueBool();
				this.content = val;
			} else if ("deleteColumns" === reader.GetName()) {
				val = reader.GetValueBool();
				this.deleteColumns = val;
			} else if ("deleteRows" === reader.GetName()) {
				val = reader.GetValueBool();
				this.deleteRows = val;
			} else if ("formatCells" === reader.GetName()) {
				val = reader.GetValueBool();
				this.formatCells = val;
			} else if ("formatColumns" === reader.GetName()) {
				val = reader.GetValueBool();
				this.formatColumns = val;
			} else if ("formatRows" === reader.GetName()) {
				val = reader.GetValueBool();
				this.formatRows = val;
			} else if ("insertColumns" === reader.GetName()) {
				val = reader.GetValueBool();
				this.insertColumns = val;
			} else if ("insertHyperlinks" === reader.GetName()) {
				val = reader.GetValueBool();
				this.insertHyperlinks = val;
			} else if ("insertRows" === reader.GetName()) {
				val = reader.GetValueBool();
				this.insertRows = val;
			} else if ("objects" === reader.GetName()) {
				val = reader.GetValueBool();
				this.objects = val;
			} else if ("pivotTables" === reader.GetName()) {
				val = reader.GetValueBool();
				this.pivotTables = val;
			} else if ("scenarios" === reader.GetName()) {
				val = reader.GetValueBool();
				this.scenarios = val;
			} else if ("selectLockedCells" === reader.GetName()) {
				val = reader.GetValueBool();
				this.selectLockedCells = val;
			} else if ("selectUnlockedCells" === reader.GetName()) {
				val = reader.GetValueBool();
				this.selectUnlockedCells = val;
			} else if ("sheet" === reader.GetName()) {
				val = reader.GetValueBool();
				this.sheet = val;
			} else if ("sort" === reader.GetName()) {
				val = reader.GetValueBool();
				this.sort = val;
			}
		}
	};

	Asc.CSheetProtection.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*sheetProtection*/);
		writer.WriteXmlNullableAttributeString("password", this.password);
		writer.WriteXmlNullableAttributeString("algorithmName", AscCommonExcel.ToXml_ST_AlgorithmName(this.algorithmName));
		writer.WriteXmlNullableAttributeString("hashValue", this.hashValue);
		writer.WriteXmlNullableAttributeString("saltValue", this.saltValue);
		writer.WriteXmlNullableAttributeNumber("spinCount", this.spinCount);

		writer.WriteXmlNullableAttributeNumber("autoFilter", boolToNumber(this.autoFilter));
		writer.WriteXmlNullableAttributeNumber("content", boolToNumber(this.content));
		writer.WriteXmlNullableAttributeNumber("deleteColumns", boolToNumber(this.deleteColumns));
		writer.WriteXmlNullableAttributeNumber("deleteRows", boolToNumber(this.deleteRows));
		writer.WriteXmlNullableAttributeNumber("formatCells", boolToNumber(this.formatCells));
		writer.WriteXmlNullableAttributeNumber("formatColumns", boolToNumber(this.formatColumns));
		writer.WriteXmlNullableAttributeNumber("formatRows", boolToNumber(this.formatRows));
		writer.WriteXmlNullableAttributeNumber("insertColumns", boolToNumber(this.insertColumns));
		writer.WriteXmlNullableAttributeNumber("insertHyperlinks", boolToNumber(this.insertHyperlinks));
		writer.WriteXmlNullableAttributeNumber("insertRows", boolToNumber(this.insertRows));
		writer.WriteXmlNullableAttributeNumber("objects", boolToNumber(this.objects));
		writer.WriteXmlNullableAttributeNumber("pivotTables", boolToNumber(this.pivotTables));
		writer.WriteXmlNullableAttributeNumber("scenarios", boolToNumber(this.scenarios));
		writer.WriteXmlNullableAttributeNumber("selectLockedCells", boolToNumber(this.selectLockedCells));
		writer.WriteXmlNullableAttributeNumber("selectUnlockedCells", boolToNumber(this.selectUnlockedCells));
		writer.WriteXmlNullableAttributeNumber("sheet", boolToNumber(this.sheet));
		writer.WriteXmlNullableAttributeNumber("sort", boolToNumber(this.sort));
		writer.WriteXmlAttributesEnd(true);
	};

	//****sparklines****
	AscCommonExcel.sparklineGroup.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("colorSeries" === name) {
				this.colorSeries = AscCommon.getColorFromXml2(reader);
			} else if ("colorNegative" === name) {
				this.colorNegative = AscCommon.getColorFromXml2(reader);
			} else if ("colorAxis" === name) {
				this.colorAxis = AscCommon.getColorFromXml2(reader);
			} else if ("colorMarkers" === name) {
				this.colorMarkers = AscCommon.getColorFromXml2(reader);
			} else if ("colorFirst" === name) {
				this.colorFirst = AscCommon.getColorFromXml2(reader);
			} else if ("colorLast" === name) {
				this.colorLast = AscCommon.getColorFromXml2(reader);
			} else if ("colorHigh" === name) {
				this.colorHigh = AscCommon.getColorFromXml2(reader);
			} else if ("colorLow" === name) {
				this.colorLow = AscCommon.getColorFromXml2(reader);
			} else if ("f" === name) {
				//TODO текст, возможно нужно использовать prepareTextToXml
				this.f = prepareTextFromXml(reader.GetTextDecodeXml());
			} else if ("sparklines" === name) {
				var depth2 = reader.GetDepth();
				while (reader.ReadNextSiblingNode(depth2)) {
					var name2 = reader.GetNameNoNS();
					if ("sparkline" === name2) {
						var newSparklineGroup = new AscCommonExcel.sparkline();
						//newSparklineGroup.setWorksheet(oWorksheet);
						newSparklineGroup.fromXml(reader);
						//oWorksheet.aSparklineGroups.push(newSparklineGroup);
						this.arrSparklines.push(newSparklineGroup);
					}
				}

			}
		}
	};

	AscCommonExcel.sparklineGroup.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("manualMax" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.manualMax = val;
			} else if ("manualMin" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.manualMin = val;
			} else if ("lineWeight" === reader.GetName()) {
				val = reader.GetValueDouble();
				this.lineWeight = val;
			} else if ("type" === reader.GetName()) {
				val = reader.GetValue();
				this.type = FromXml_ST_SparklineType(val);
			} else if ("dateAxis" === reader.GetName()) {
				val = reader.GetValueBool();
				this.dateAxis = val;
			} else if ("displayEmptyCellsAs" === reader.GetName()) {
				val = reader.GetValue();
				this.displayEmptyCellsAs = FromXML_ST_DispBlanksAs(val);
			} else if ("markers" === reader.GetName()) {
				val = reader.GetValueBool();
				this.markers = val;
			} else if ("high" === reader.GetName()) {
				val = reader.GetValueBool();
				this.high = val;
			} else if ("low" === reader.GetName()) {
				val = reader.GetValueBool();
				this.low = val;
			} else if ("first" === reader.GetName()) {
				val = reader.GetValueBool();
				this.first = val;
			} else if ("last" === reader.GetName()) {
				val = reader.GetValueBool();
				this.last = val;
			} else if ("negative" === reader.GetName()) {
				val = reader.GetValueBool();
				this.negative = val;
			} else if ("displayXAxis" === reader.GetName()) {
				val = reader.GetValueBool();
				this.displayXAxis = val;
			} else if ("displayHidden" === reader.GetName()) {
				val = reader.GetValueBool();
				this.displayHidden = val;
			} else if ("minAxisType" === reader.GetName()) {
				val = reader.GetValue();
				this.minAxisType = FromXml_ST_SparklineAxisMinMax(val);
			} else if ("maxAxisType" === reader.GetName()) {
				val = reader.GetValue();
				this.maxAxisType = FromXml_ST_SparklineAxisMinMax(val);
			} else if ("rightToLeft" === reader.GetName()) {
				val = reader.GetValueBool();
				this.rightToLeft = val;
			}
		}
	};

	AscCommonExcel.sparklineGroup.prototype.toXml = function (writer, name, ns, childns) {

		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeNumber("manualMax", this.manualMax);
		writer.WriteXmlNullableAttributeNumber("manualMax", this.manualMax);
		writer.WriteXmlNullableAttributeNumber("lineWeight", this.lineWeight);
		writer.WriteXmlNullableAttributeString("type", ToXml_ST_SparklineType(this.type));
		writer.WriteXmlNullableAttributeBool("dateAxis", this.dateAxis);
		writer.WriteXmlNullableAttributeString("displayEmptyCellsAs", ToXML_ST_DispBlanksAs(this.displayEmptyCellsAs));
		writer.WriteXmlNullableAttributeBool("markers", this.markers);
		writer.WriteXmlNullableAttributeBool("high", this.high);
		writer.WriteXmlNullableAttributeBool("low", this.low);
		writer.WriteXmlNullableAttributeBool("first", this.first);
		writer.WriteXmlNullableAttributeBool("last", this.last);
		writer.WriteXmlNullableAttributeBool("negative", this.negative);
		writer.WriteXmlNullableAttributeBool("displayXAxis", this.displayXAxis);
		writer.WriteXmlNullableAttributeBool("displayHidden", this.displayHidden);

		writer.WriteXmlNullableAttributeString("minAxisType", ToXml_ST_SparklineAxisMinMax(this.minAxisType));
		writer.WriteXmlNullableAttributeString("minAxisType", ToXml_ST_SparklineAxisMinMax(this.minAxisType));


		writer.WriteXmlNullableAttributeBool("rightToLeft", this.rightToLeft);
		writer.WriteXmlAttributesEnd();


		//DOCUMENTATION: The auto attribute of the CT_Color element MUST NOT exist
		var checkOnAuto = function (val) {
			return !val.rgb && !val.theme && !val.tint;
		};

		if (this.colorSeries && !checkOnAuto(this.colorSeries)) {
			AscCommon.writeColorToXml(writer, childns + "colorSeries", this.colorSeries);
		}
		if (this.colorNegative && !checkOnAuto(this.colorNegative)) {
			AscCommon.writeColorToXml(writer, childns + "colorNegative", this.colorNegative);
		}
		if (this.colorAxis && !checkOnAuto(this.colorAxis)) {
			AscCommon.writeColorToXml(writer, childns + "colorAxis", this.colorAxis);
		}
		if (this.colorMarkers && !checkOnAuto(this.colorMarkers)) {
			AscCommon.writeColorToXml(writer, childns + "colorMarkers", this.colorMarkers);
		}
		if (this.colorFirst && !checkOnAuto(this.colorFirst)) {
			AscCommon.writeColorToXml(writer, childns + "colorFirst", this.colorFirst);
		}
		if (this.colorLast && !checkOnAuto(this.colorLast)) {
			AscCommon.writeColorToXml(writer, childns + "colorLast", this.colorLast);
		}
		if (this.colorHigh && !checkOnAuto(this.colorHigh)) {
			AscCommon.writeColorToXml(writer, childns + "colorHigh", this.colorHigh);
		}
		if (this.colorLow && !checkOnAuto(this.colorLow)) {
			AscCommon.writeColorToXml(writer, childns + "colorLow", this.colorLow);
		}


		if (this.f) {
			writer.WriteXmlString("<" + "\"xm:f\"" + ">");
			writer.WriteXmlStringEncode(this.f);
			writer.WriteXmlString("</" + "\"xm:f\"" + ">");
		}

		if (this.arrSparklines) {
			writer.WriteXmlNodeStart(childns + "sparklines");
			writer.WriteXmlAttributesEnd();

			for (var i = 0; i < this.arrSparklines.length; ++i) {
				if (this.arrSparklines[i]) {
					this.arrSparklines[i].toXml(writer, "sparkline", childns);
				}
			}

			writer.WriteXmlNodeEnd(childns + "sparklines");
		}

		writer.WriteXmlNodeEnd(ns + name);
	};


	AscCommonExcel.sparkline.prototype.fromXml = function (reader) {
		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("f" === name) {
				this.setF(prepareTextFromXml(reader.GetTextDecodeXml()));
			} else if ("sqref" === name) {
				this.setSqRef(reader.GetText());
			}
		}
	};

	AscCommonExcel.sparkline.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlAttributesEnd();

		if (this.f) {
			writer.WriteXmlString("<xm:f>");
			writer.WriteXmlStringEncode(this.f);
			writer.WriteXmlString("</xm:f>");
		}
		if (this.sqRef) {
			writer.WriteXmlString("<xm:sqref>");
			writer.WriteXmlStringEncode(this.sqRef.getName());
			writer.WriteXmlString("</xm:sqref>");
		}

		writer.WriteXmlNodeEnd(ns + name);
	};

	function COfficeArtExtensionList(_ws) {
		this.arrExt = [];

		this._ws = _ws;
	}

	COfficeArtExtensionList.prototype.fromXml = function (reader) {
		if (reader.IsEmptyNode()) {
			return;
		}

		/*virtual void fromXML(XmlUtils::CXmlLiteReader& oReader)
		{
			if ( oReader.IsEmptyNode() )
				return;

			int nCurDepth = oReader.GetDepth();
			while( oReader.ReadNextSiblingNode( nCurDepth ) )
			{
				std::wstring sName = XmlUtils::GetNameNoNS(oReader.GetName());
				if ( _T("ext") == sName )
				{
					OOX::Drawing::COfficeArtExtension *oExt = new OOX::Drawing::COfficeArtExtension(oReader);
					if (oExt) m_arrExt.push_back( oExt );
				}
			}
		}*/

		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("ext" === name) {
				val = new COfficeArtExtension(this._ws);
				val.fromXml(reader);
				this.arrExt.push(val);
			}
		}
	};
	COfficeArtExtensionList.prototype.getConditionalFormattingId = function () {
		var res = null;

		this.arrExt.forEach(function (ext) {
			if (ext && ext.ids) {
				res = ext.ids[0];
			}
		});

		return res;
	};


	COfficeArtExtensionList.prototype.getSlicerStyles = function () {
		var res = null;
		this.arrExt.forEach(function (ext) {
			if (ext && ext.slicerStyles) {
				res = ext.slicerStyles;
			}
		});
		return res;
	};

	COfficeArtExtensionList.prototype.getDxfs = function () {
		var res = [];
		this.arrExt.forEach(function (ext) {
			if (ext && ext.dxfs && ext.dxfs.length) {
				res = res.concat(ext.dxfs);
			}
		});
		return res;
	};

	COfficeArtExtensionList.prototype.toXml = function (writer, ns) {
		if (this.arrExt && this.arrExt.length) {

			if (!ns) {
				ns = "";
			}

			writer.WriteXmlNodeStart(ns + "extLst");
			writer.WriteXmlAttributesEnd();

			this.arrExt.forEach(function (ext) {
				if (ext) {
					ext.toXml(writer, ns);
				}
			});

			writer.WriteXmlNodeEnd(ns + "extLst");
		}
	};

	function COfficeArtExtension(_ws) {
		this.uri = null;
		this.dataValidations = null;

		this.slicerCachesIds = [];
		this.slicerCachesExtIds = [];
		this.slicerListIds = [];
		this.slicerListExtIds = [];

		this.tableSlicerCache = null;
		this.aConditionalFormattingRules = [];
		this.sparklineGroups = [];
		this.dxfs = null;

		this.slicerStyles = null;

		this.ids = [];

		this._ws = _ws;
	}

	COfficeArtExtension.prototype.fromXml = function (reader, ws) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		if (this.uri === "{63B3BB69-23CF-44E3-9099-C40C66FF867C}" || this.uri === "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" || this.uri ===
			"{504A1905-F514-4f6f-8877-14C23A59335A}" || this.uri === "{78C0D931-6437-407d-A8EE-F0AAD7539E65}" || this.uri === "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" ||
			this.uri === "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" || this.uri === "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" || this.uri ===
			"{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}" || this.uri === "{BBE1A952-AA13-448e-AADC-164F8A28A991}" || this.uri === "{46BE6895-7355-4a93-B00E-2C351335B9C9}" ||
			this.uri === "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" || this.uri === "{03082B11-2C62-411c-B77F-237D8FCFBE4C}" || this.uri ===
			"{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" || this.uri === "{470722E0-AACD-4C17-9CDC-17EF765DBC7E}" || this.uri === "{46F421CA-312F-682f-3DD2-61675219B42D}" ||
			this.uri === "{DE250136-89BD-433C-8126-D09CA5730AF9}" || this.uri === "{19B8F6BF-5375-455C-9EA6-DF929625EA0E}" || this.uri ===
			"http://schemas.microsoft.com/office/drawing/2008/diagram") {
			var name2, depth2;
			var val;
			var depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				var name = reader.GetNameNoNS();
				if ("compatExt" === name) {

				} else if ("compatExt" === name) {

				} else if ("sparklineGroups" === name) {
					depth2 = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depth2)) {
						name2 = reader.GetNameNoNS();
						if ("sparklineGroup" === name2) {

							var newSparklineGroup = new AscCommonExcel.sparklineGroup(true);
							//newSparklineGroup.setWorksheet(oWorksheet);
							newSparklineGroup.fromXml(reader);
							//oWorksheet.aSparklineGroups.push(newSparklineGroup);
							this.sparklineGroups.push(newSparklineGroup);
						}
					}
				} else if ("dataModelExt" === name) {

				} else if ("table" === name) {

				} else if ("conditionalFormattings" === name) {

					depth2 = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depth2)) {
						name2 = reader.GetNameNoNS();
						if ("conditionalFormatting" === name2) {
							var oConditionalFormatting = new AscCommonExcel.CConditionalFormatting();
							oConditionalFormatting.fromXml(reader);
							this.aConditionalFormattingRules.push(oConditionalFormatting);
						}
					}
				} else if ("dataValidations" === name) {
					val = new AscCommonExcel.CDataValidations();
					val.fromXml(reader);
					this.dataValidations = val;
				} else if ("connection" === name) {

				} else if ("slicerList" === name) {

				} else if ("slicerCaches" === name) {
					depth2 = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depth2)) {
						name2 = reader.GetNameNoNS();
						if ("slicerCache" === name2) {

							while (reader.MoveToNextAttribute()) {
								if ("id" === reader.GetNameNoNS()) {
									this.slicerCachesIds.push(reader.GetValue());
								}
							}
						}
					}
				} else if ("dxfs" === name) {
					depth2 = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depth2)) {
						name2 = reader.GetNameNoNS();
						if ("dxf" === name2) {
							val = new AscCommonExcel.CellXfs();
							val.fromXml(reader);
							if (!this.dxfs) {
								this.dxfs = [];
							}
							this.dxfs.push(val);
						}
					}
				} else if ("slicerStyles" === name && typeof Asc.CT_slicerStyles != "undefined") {
					val = new Asc.CT_slicerStyles();
					val.fromXml(reader);
					this.slicerStyles = val;
				} else if ("slicerCachePivotTables" === name) {

				} else if ("tableSlicerCache" === name) {
					val = new Asc.CT_tableSlicerCache();
					val.fromXml(reader);
					this.tableSlicerCache = val;
				} else if ("slicerCacheHideItemsWithNoData" === name) {

				} else if ("id" === name) {
					val = prepareTextFromXml(reader.GetTextDecodeXml());
					this.ids.push(val);
				} else if ("presenceInfo" === name) {

				}
			}
		} else {
			if (!reader.IsEmptyNode()) {
				reader.ReadTillEnd();
			}
		}
	};

	COfficeArtExtension.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("uri" === reader.GetName()) {
				val = reader.GetValue();
				this.uri = val;
			}
		}
	};

	COfficeArtExtension.prototype.toXml = function (writer, ns) {
		if (!ns) {
			ns = "";
		}

		//проверкой xml была выявлена ошибка, данный ext(tableSlicerCache) пишется с префиксом x:
		if (this.tableSlicerCache) {
			ns = "x:";
		}

		var i;
		writer.WriteXmlNodeStart(ns + "ext");

		//attributes
		if (this.uri) {
			writer.WriteXmlAttributeString("uri", this.uri);
			if (this.additionalNamespace) {
				writer.WriteXmlString(this.additionalNamespace);
			} else {
				writer.WriteXmlAttributeString("xmlns:x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
			}
		}
		if (this.additionalNamespace) {
		}

		writer.WriteXmlAttributesEnd();


		if (this.CompatExt) {
		}

		if (this.sparklineGroups && this.sparklineGroups.length) {
			writer.WriteXmlString("<x14:sparklineGroups xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">");
			for (i = 0; i < this.sparklineGroups.length; ++i) {
				this.sparklineGroups[i].toXml(writer, "sparklineGroup", "x14:", "x14:");
			}
			writer.WriteXmlNodeEnd("x14:sparklineGroups");
		}

		if (this.AltTextTable) {
		}
		if (this.DataModelExt) {
		}
		if (this.aConditionalFormattingRules && this.aConditionalFormattingRules.length) {
			writer.WriteXmlNodeStart("x14:conditionalFormattings");
			writer.WriteXmlAttributesEnd();

			for (i = 0; i < this.aConditionalFormattingRules.length; ++i) {
				var oConditionalFormatting = new AscCommonExcel.CConditionalFormatting();
				oConditionalFormatting.aRules = [this.aConditionalFormattingRules[i]];
				oConditionalFormatting.toXml(writer, true)
			}

			writer.WriteXmlNodeEnd("x14:conditionalFormattings");
		}

		if (this.dataValidations) {
			this.dataValidations.toXml(writer, true);
		}

		var oSlicers;
		if (this.slicerListIds && this.slicerListIds.length) {
			oSlicers = new CSlicerRefs();
			oSlicers.arr = this.slicerListIds;
			oSlicers.toXml(writer, "x14:slicerList");
		}
		if (this.slicerListExtIds && this.slicerListExtIds.length) {
			oSlicers = new CSlicerRefs();
			oSlicers.arr = this.slicerListExtIds;
			oSlicers.toXml(writer, "x14:slicerList");
		}

		var oSlicerCaches;
		if (this.slicerCachesIds && this.slicerCachesIds.length) {
			oSlicerCaches = new CSlicerCaches();
			oSlicerCaches.arr = this.slicerCachesIds;
			oSlicerCaches.toXml(writer, "slicerCaches", "x14:");
		}
		if (this.slicerCachesExtIds && this.slicerCachesExtIds.length) {
			oSlicerCaches = new CSlicerCaches();
			oSlicerCaches.arr = this.slicerCachesExtIds;
			oSlicerCaches.toXml(writer, "slicerCaches", "x15:");
		}

		if (this.dxfs && this.dxfs.length) {
			writer.WriteXmlArray(this.dxfs, "dxf", "x14:dxfs", true);
		}
		if (this.slicerStyles) {
			this.slicerStyles.toXml(writer, "x14:slicerStyles");
		}

		if (this.m_oSlicerCachePivotTables) {
		}
		if (this.tableSlicerCache) {
			this.tableSlicerCache.toXml(writer, "x15:tableSlicerCache");
		}
		if (this.m_oSlicerCacheHideItemsWithNoData) {
		}
		if (this.m_oId) {
		}

		writer.WriteXmlNodeEnd(ns + "ext");
	};


	Asc.CT_slicerStyles.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("slicerStyle" === name) {
				var val = new Asc.CT_slicerStyle();
				val.fromXml(reader);
				this.slicerStyle.push(val);
			}
		}
	};

	Asc.CT_slicerStyles.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("defaultSlicerStyle" === reader.GetName()) {
				val = reader.GetValue();
				this.defaultSlicerStyle = val;
			}
		}
	};

	Asc.CT_slicerStyles.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name/*"x14:slicerStyles"*/);
		writer.WriteXmlNullableAttributeStringEncode("defaultSlicerStyle", this.defaultSlicerStyle);

		writer.WriteXmlAttributesEnd();
		if (this.slicerStyle.length > 0) {
			for (var i = 0; i < this.slicerStyle.length; ++i) {
				this.slicerStyle[i].toXml(writer, "x14:slicerStyle");
			}
		}
		writer.WriteXmlNodeEnd(name);
	};

	Asc.CT_slicerStyle.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("slicerStyleElement" === name) {
				var val = new Asc.CT_slicerStyleElement();
				val.fromXml(reader);
				this.slicerStyleElements.push(val);
			}
		}
	};

	Asc.CT_slicerStyle.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.name = val;
			}
		}
	};

	//один аттрибут + массив значений
	Asc.CT_slicerStyle.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name/*"x14:slicerStyle"*/);
		writer.WriteXmlAttributeStringEncode("name", this.name);
		writer.WriteXmlAttributesEnd();
		if (this.slicerStyleElements.length > 0) {
			for (var i = 0; i < this.slicerStyleElements.length; ++i) {
				this.slicerStyleElements[i].toXml(writer, "x14:slicerStyleElement");
			}
		}
		writer.WriteXmlNodeEnd(name);
	};

	Asc.CT_slicerStyleElement.prototype.fromXml = function (reader) {

		/*ReadAttributes(oReader);
			if (oReader.IsEmptyNode())
				return;
			oReader.ReadTillEnd();*/

		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	Asc.CT_slicerStyleElement.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("type" === reader.GetName()) {
				val = reader.GetValue();
				this.type = val;
			} else if ("dxfId" === reader.GetName()) {
				val = reader.GetValue();
				this.dxfId = val;
			}
		}
	};

	Asc.CT_slicerStyleElement.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributeString("type", this.type);
		writer.WriteXmlNullableAttributeNumber("dxfId", this.dxfId);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNodeEnd(name);
	};

	function CSlicerCaches () {
		this.arr = null;
	}

	CSlicerCaches.prototype.toXml = function (writer, name, ns) {
		if (!this.arr || !this.arr.length) {
			return;
		}
		if (!ns) {
			ns = "";
		}
		writer.WriteXmlNodeStart(ns + name);
		var sChildPrefix;
		if(ns.length > 0 && "x14:" !== ns) {
			sChildPrefix = "x14:";
			writer.WriteXmlAttributeString("xmlns:x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
		}
		writer.WriteXmlAttributesEnd();
		if (this.arr.length > 0) {
			for (var i = 0; i < this.arr.length; ++i) {
				writer.WriteXmlNodeStart("x14:slicerCache");
				writer.WriteXmlAttributeString("r:id", this.arr[i]);
				writer.WriteXmlAttributesEnd();
				writer.WriteXmlNodeEnd("x14:slicerCache");
			}
		}
		writer.WriteXmlNodeEnd(ns + name);
	};

	function CSlicerRefs () {
		this.arr = null;
	}

	CSlicerRefs.prototype.toXml = function (writer, name) {
		if (!this.arr || !this.arr.length) {
			return;
		}

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributeString("xmlns:x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
		writer.WriteXmlAttributesEnd();
		if (this.arr.length > 0) {
			for (var i = 0; i < this.arr.length; ++i) {
				writer.WriteXmlNodeStart("x14:slicer");
				writer.WriteXmlAttributeString("r:id", this.arr[i]);
				writer.WriteXmlAttributesEnd();
				writer.WriteXmlNodeEnd("x14:slicer");
			}
		}
		writer.WriteXmlNodeEnd(name);
	};

	AscCommonExcel.CConditionalFormatting.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("cfRule" === name) {
				//var ext = {isExt: false};
				val = new AscCommonExcel.CConditionalFormattingRule();
				val.fromXml(reader);
				this.aRules.push(val);
			} else if ("sqref" === name || "Range" === name) {
				val = reader.GetText();
				this.setSqRef(val);
			}
		}
	};

	AscCommonExcel.CConditionalFormatting.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("sqref" === reader.GetName()) {
				val = reader.GetValue();
				this.setSqRef(val);
			} else if ("pivot" === reader.GetName()) {
				val = reader.GetValueBool();
				this.pivot = val;
			}
		}
	};

	AscCommonExcel.CConditionalFormatting.prototype.toXml = function (writer, bExtendedWrite) {
		var node_name = bExtendedWrite ? "x14:conditionalFormatting" : "conditionalFormatting";

		writer.WriteXmlString("<" + node_name);
		if (bExtendedWrite) {
			writer.WriteXmlAttributeString("xmlns:xm", "http://schemas.microsoft.com/office/excel/2006/main");
		} else {
			writer.WriteXmlAttributeString("sqref", AscCommonExcel.getSqRefString(this.aRules[0].ranges));
		}
		if (this.pivot) {
			writer.WriteXmlString(" pivot=\"1\"");
		}
		writer.WriteXmlString(">");

		for (var i = 0; i < this.aRules.length; ++i) {
			if (this.aRules[i]) {
				this.aRules[i].toXml(writer, bExtendedWrite);
			}
		}
		//в случае extLst беру у первого элемента ranges
		if (bExtendedWrite && this.aRules[0]) {
			writer.WriteXmlString("<xm:sqref>" + AscCommonExcel.getSqRefString(this.aRules[0].ranges) + "</xm:sqref>");
		}
		writer.WriteXmlString("</" + node_name + ">");
	};

	AscCommonExcel.CConditionalFormattingRule.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}


		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("colorScale" === name) {
				val = new AscCommonExcel.CColorScale();
				val.fromXml(reader);
				this.aRuleElements.push(val);
			} else if ("dataBar" === name) {
				val = new AscCommonExcel.CDataBar();
				val.fromXml(reader);
				this.aRuleElements.push(val);
			} else if ("formula" === name || "f" === name) {
				val = new AscCommonExcel.CFormulaCF()
				val.Text = prepareTextFromXml(reader.GetTextDecodeXml());
				this.aRuleElements.push(val);
			} else if ("iconSet" === name) {
				val = new AscCommonExcel.CIconSet();
				val.fromXml(reader);
				this.aRuleElements.push(val);
			} else if ("dxf" === name) {
				val = new AscCommonExcel.CellXfs();
				val.fromXml(reader);
				this.dxf = val;
			} else if ("extLst" === name) {
				var extLst = new COfficeArtExtensionList(this);
				extLst.fromXml(reader);

				this._openId = extLst.getConditionalFormattingId();
			}
		}

	};

	AscCommonExcel.CConditionalFormattingRule.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("aboveAverage" === reader.GetName()) {
				val = reader.GetValueBool();
				this.aboveAverage = val;
			} else if ("bottom" === reader.GetName()) {
				val = reader.GetValueBool();
				this.bottom = val;
			} else if ("dxfId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.dxf = reader.context.InitOpenManager.Dxfs[val];
			} else if ("equalAverage" === reader.GetName()) {
				val = reader.GetValueBool();
				this.equalAverage = val;
			} else if ("operator" === reader.GetName()) {
				val = reader.GetValue();
				this.operator = FromXml_CFOperatorType(val);
			} else if ("percent" === reader.GetName()) {
				val = reader.GetValueBool();
				this.percent = val;
			} else if ("priority" === reader.GetName()) {
				val = reader.GetValueInt();
				this.priority = val;
			} else if ("rank" === reader.GetName()) {
				val = reader.GetValueInt();
				this.rank = val;
			} else if ("stdDev" === reader.GetName()) {
				val = reader.GetValueInt();
				this.stdDev = val;
			} else if ("stopIfTrue" === reader.GetName()) {
				val = reader.GetValueBool();
				this.stopIfTrue = val;
			} else if ("text" === reader.GetName()) {
				val = reader.GetValue();
				this.text = val;
			} else if ("timePeriod" === reader.GetName()) {
				val = reader.GetValue();
				this.timePeriod = FromXml_ST_TimePeriod(val);
			} else if ("type" === reader.GetName()) {
				val = reader.GetValue();
				this.type = FromXml_ST_CfType(val);
			} else if ("id" === reader.GetName()) {
				val = reader.GetValue();
				this._openId = val;
			}
		}
	};

	AscCommonExcel.CConditionalFormattingRule.prototype.toXml = function (writer, bExtendedWrite) {
		var node_name = bExtendedWrite ? "x14:cfRule" : "cfRule";

		writer.WriteXmlString("<" + node_name);

		/*<xsd:simpleType name="ST_CfType">
			2673 <xsd:restriction base="xsd:string">
			2674 <xsd:enumeration value="expression"/>
			2675 <xsd:enumeration value="cellIs"/>
			2676 <xsd:enumeration value="colorScale"/>
			2677 <xsd:enumeration value="dataBar"/>
			2678 <xsd:enumeration value="iconSet"/>
			2679 <xsd:enumeration value="top10"/>
			2680 <xsd:enumeration value="uniqueValues"/>
			2681 <xsd:enumeration value="duplicateValues"/>
			2682 <xsd:enumeration value="containsText"/>
			2683 <xsd:enumeration value="notContainsText"/>
			2684 <xsd:enumeration value="beginsWith"/>
			2685 <xsd:enumeration value="endsWith"/>
			2686 <xsd:enumeration value="containsBlanks"/>
			2687 <xsd:enumeration value="notContainsBlanks"/>
			2688 <xsd:enumeration value="containsErrors"/>
			2689 <xsd:enumeration value="notContainsErrors"/>
			2690 <xsd:enumeration value="timePeriod"/>
			2691 <xsd:enumeration value="aboveAverage"/>
			2692 </xsd:restriction>*/


		writer.WriteXmlAttributeString("type", ToXml_ST_CfType(this.type));
		writer.WriteXmlAttributeNumber("priority", this.priority);

		if (false === this.aboveAverage) {
			writer.WriteXmlString(" aboveAverage=\"0\"");
		}
		if (true === this.bottom) {
			writer.WriteXmlString(" bottom=\"1\"");
		}
		writer.WriteXmlNullableAttributeNumber("dxfId", this.dxfId);
		if (true === this.equalAverage) {
			writer.WriteXmlString(" equalAverage=\"1\"");
		}
		writer.WriteXmlNullableAttributeString("operator", ToXml_CFOperatorType(this.operator));
		if (true === this.percent) {
			writer.WriteXmlString(" percent=\"1\"");
		}
		writer.WriteXmlNullableAttributeNumber("rank", this.rank);
		writer.WriteXmlNullableAttributeNumber("stdDev", this.stdDev);
		if (true === this.stopIfTrue) {
			writer.WriteXmlString(" stopIfTrue=\"1\"");
		}
		writer.WriteXmlNullableAttributeStringEncode("text", this.text);
		writer.WriteXmlNullableAttributeString("timePeriod", ToXml_ST_TimePeriod(this.timePeriod));

		if (bExtendedWrite) {
			/*if (!this.id) {
				writer.WriteXmlAttributeString("id", "{" + AscCommon.GUID() + "}");
			} else {
				//проверить, пойдёт ли такой id
				writer.WriteXmlNullableAttributeString("id", this.id);
			}*/
			writer.WriteXmlAttributeString("id", "{" + AscCommon.GUID() + "}");
		}

		writer.WriteXmlString(">");

		for (var i = 0; i < this.aRuleElements.length; ++i) {
			var elem = this.aRuleElements[i];
			if (elem.toXml) {
				elem.toXml(writer, bExtendedWrite);
			} else {
				var node_formula_name = bExtendedWrite ? "xm:f" : "formula";
				if (elem) {
					writer.WriteXmlString("<" + node_formula_name + ">");
					writer.WriteXmlStringEncode(elem.Text);
					writer.WriteXmlString("</" + node_formula_name + ">");
				}
			}
		}

		if (this.dxf) {
			this.dxf.toXml(writer, "x14:dxf");
		}

		writer.WriteXmlString("</" + node_name + ">");
	};

	AscCommonExcel.CConditionalFormattingRule.prototype.isExtended = function () {
		if (this.dxf) {
			return true;
		}
		for (var i = 0; i < this.aRuleElements.length; i++) {
			if (this.aRuleElements[i].isExtended()) {
				return true;
			}
		}
	};


	AscCommonExcel.CColorScale.prototype.fromXml = function (reader) {
		if (reader.IsEmptyNode()) {
			return;
		}

		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("cfvo" === name) {
				val = new AscCommonExcel.CConditionalFormatValueObject();
				val.fromXml(reader);
				this.aCFVOs.push(val);
			} else if ("color" === name) {
				val = AscCommon.getColorFromXml2(reader);
				if (null != val) {
					this.aColors.push(val);
				}
			}
		}
	};

	AscCommonExcel.CColorScale.prototype.toXml = function (writer, bExtendedWrite) {
		if (1 < this.aCFVOs.length && 1 < this.aColors.length) // min 2 + 2
		{
			writer.WriteXmlString("<colorScale>");
			var i;
			for (i = 0; i < this.aCFVOs.length; ++i) {
				if (this.aCFVOs[i]) {
					this.aCFVOs[i].toXml(writer, bExtendedWrite);
				}
			}
			for (i = 0/*m_arrValues.length*/; i < this.aColors.length; ++i) {
				if (this.aColors[i]) {
					AscCommon.writeColorToXml(writer, "color", this.aColors[i]);
				}
			}
			writer.WriteXmlString("</colorScale>");
		}
	};

	AscCommonExcel.CColorScale.prototype.isExtended = function () {
		return false;
	};

	AscCommonExcel.CConditionalFormatValueObject.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("formula" === name || "f" === name) {
				//TODO prepareTextToXml?
				this.Val = prepareTextFromXml(reader.GetTextDecodeXml());
			}
		}
	};

	AscCommonExcel.CConditionalFormatValueObject.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("gte" === reader.GetName()) {
				val = reader.GetValueBool();
				this.Gte = val;
			} else if ("type" === reader.GetName()) {
				val = reader.GetValue();
				this.Type = FromXml_ST_CfvoType(val);
			} else if ("val" === reader.GetName()) {
				val = reader.GetValue();
				this.Val = val;
			}
		}
	};

	AscCommonExcel.CConditionalFormatValueObject.prototype.toXml = function (writer, bExtendedWrite) {
		if (bExtendedWrite === false) {
			//if (m_oType.GetValue() == SimpleTypes::Spreadsheet::autoMin) m_oType.SetValue(SimpleTypes::Spreadsheet::Minimum);
			//if (m_oType.GetValue() == SimpleTypes::Spreadsheet::autoMax) m_oType.SetValue(SimpleTypes::Spreadsheet::Maximum);
		}

		var node_name = bExtendedWrite ? "x14:cfvo" : "cfvo";

		writer.WriteXmlString("<" + node_name);
		writer.WriteXmlAttributeString("type", ToXml_ST_CfvoType(this.Type));
		if (false === this.Gte) {
			writer.WriteXmlString(" gte=\"0\"");
		}
		if (!bExtendedWrite) {
			if (this.Val) {
				writer.WriteXmlAttributeStringEncode("val", this.Val);
			} else if (this.formula) {
				writer.WriteXmlAttributeStringEncode("val", this.formula);
			}
		}
		writer.WriteXmlString(">");

		if (bExtendedWrite) {
			if (this.formula) {
				this.formula.toXml(writer, true);
			} else if (null != this.Val) {
				//TODO prepareTextToXml?
				var formula = new AscCommonExcel.CFormulaCF()
				formula.Text = this.Val;
				formula.toXml(writer, true);
			}
		}

		writer.WriteXmlString("</" + node_name + ">");
	};

	AscCommonExcel.CFormulaCF.prototype.toXml = function (writer, bExtendedWrite) {
		var node_name = bExtendedWrite ? "xm:f" : "formula";

		writer.WriteXmlString("<" + node_name + ">");
		writer.WriteXmlStringEncode(this.Text);
		writer.WriteXmlString("</" + node_name + ">");
	};

	AscCommonExcel.CDataBar.prototype.fromXml = function (reader) {
		/*if (!reader.ReadNextNode()) {
			return;
		}*/

		this.readAttr(reader);

		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("cfvo" === name) {
				val = new AscCommonExcel.CConditionalFormatValueObject();
				val.fromXml(reader);
				this.aCFVOs.push(val);
			} else if ("color" === name || "fillColor" === name) {
				val = AscCommon.getColorFromXml2(reader);
				if (null != val) {
					this.Color = val;
				}
			} else if ("axisColor" === name) {
				val = AscCommon.getColorFromXml2(reader);
				if (null != val) {
					this.AxisColor = val;
				}
			} else if ("borderColor" === name) {
				val = AscCommon.getColorFromXml2(reader);
				if (null != val) {
					this.BorderColor = val;
				}
			} else if ("negativeFillColor" === name) {
				val = AscCommon.getColorFromXml2(reader);
				if (null != val) {
					this.NegativeColor = val;
				}
			} else if ("negativeBorderColor" === name) {
				val = AscCommon.getColorFromXml2(reader);
				if (null != val) {
					this.NegativeBorderColor = val;
				}
			}
		}
	};

	AscCommonExcel.CDataBar.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("maxLength" === reader.GetName()) {
				val = reader.GetValueInt();
				this.MaxLength = val;
			} else if ("minLength" === reader.GetName()) {
				val = reader.GetValueInt();
				this.MinLength = val;
			} else if ("showValue" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ShowValue = val;
			} else if ("axisPosition" === reader.GetName()) {
				val = reader.GetValue();
				this.AxisPosition = FromXml_ST_DataBarAxisPosition(val);
			} else if ("border" === reader.GetName()) {
				val = reader.GetValue();
				this.Border = val;
			} else if ("gradient" === reader.GetName()) {
				val = reader.GetValueBool();
				this.Gradient = val;
			} else if ("direction" === reader.GetName()) {
				val = reader.GetValue();
				this.Direction = FromXml_ST_DataBarDirection(val);
			} else if ("negativeBarColorSameAsPositive" === reader.GetName()) {
				val = reader.GetValueBool();
				this.NegativeBarColorSameAsPositive = val;
			} else if ("negativeBarBorderColorSameAsPositive" === reader.GetName()) {
				val = reader.GetValueBool();
				this.NegativeBarBorderColorSameAsPositive = val;
			}
		}
	};

	AscCommonExcel.CDataBar.prototype.toXml = function (writer, bExtendedWrite) {
		if (2 !== this.aCFVOs.length || !this.Color) {
			return;
		}

		var node_name = bExtendedWrite ? "x14:dataBar" : "dataBar";

		writer.WriteXmlString("<" + node_name);
		writer.WriteXmlNullableAttributeNumber("maxLength", this.MaxLength);
		writer.WriteXmlNullableAttributeNumber("minLength", this.MinLength);
		if (false === this.ShowValue) {
			writer.WriteXmlString(" showValue=\"0\"");
		}
		if (bExtendedWrite) {
			if (this.Border) {
				writer.WriteXmlString(" border=\"1\"");
			}
			//todo мс не пишет деволтовые значение, x2t пишет, добавить проверку?
			writer.WriteXmlNullableAttributeString("axisPosition", ToXml_ST_DataBarAxisPosition(this.AxisPosition));
			writer.WriteXmlNullableAttributeString("direction", ToXml_ST_DataBarDirection(this.Direction));

			if (false === this.Gradient) {
				writer.WriteXmlString(" gradient=\"0\"");
			}
			if (true === this.NegativeBarColorSameAsPositive) {
				writer.WriteXmlString(" negativeBarColorSameAsPositive=\"1\"");
			}
			if (false === this.NegativeBarBorderColorSameAsPositive) {
				writer.WriteXmlString(" negativeBarBorderColorSameAsPositive=\"0\"");
			}
		}
		writer.WriteXmlString(">");

		for (var i = 0; i < this.aCFVOs.length; ++i) {
			if (this.aCFVOs[i]) {
				this.aCFVOs[i].toXml(writer, bExtendedWrite);
			}
		}

		AscCommon.writeColorToXml(writer, bExtendedWrite ? "x14:fillColor" : "color", this.Color);

		if (bExtendedWrite) {
			if (this.BorderColor) {
				AscCommon.writeColorToXml(writer, "x14:borderColor", this.BorderColor);
			}
			if (this.NegativeColor) {
				AscCommon.writeColorToXml(writer, "x14:negativeFillColor", this.NegativeColor);
			}
			if (this.NegativeBorderColor) {
				AscCommon.writeColorToXml(writer, "x14:negativeBorderColor", this.NegativeBorderColor);
			}
			if (this.AxisColor) {
				AscCommon.writeColorToXml(writer, "x14:axisColor", this.AxisColor);
			}
		}

		writer.WriteXmlString("</" + node_name + ">");
	};

	AscCommonExcel.CDataBar.prototype.isExtended = function () {
		if (this.AxisColor || this.AxisPosition !== AscCommonExcel.EDataBarAxisPosition.automatic || this.Direction !== AscCommonExcel.EDataBarDirection.context ||
			this.BorderColor || this.NegativeColor || this.NegativeBorderColor || this.NegativeBarColorSameAsPositive || this.NegativeBarBorderColorSameAsPositive) {
			return true;
		}
		return false;
	};

	AscCommonExcel.CIconSet.prototype.fromXml = function (reader) {
		if (this.readAttr) {
			this.readAttr(reader);
		}

		if (reader.IsEmptyNode()) {
			return;
		}

		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("cfvo" === name) {
				val = new AscCommonExcel.CConditionalFormatValueObject();
				val.fromXml(reader);
				this.aCFVOs.push(val);
			} else if ("cfIcon" === name) {
				val = new AscCommonExcel.CConditionalFormatIconSet();
				val.fromXml(reader);
				this.aIconSets.push(val);
			}
		}
	};

	AscCommonExcel.CIconSet.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("iconSet" === reader.GetName()) {
				val = reader.GetValue();
				this.IconSet = FromXml_ST_IconSetType(val);
			} else if ("percent" === reader.GetName()) {
				val = reader.GetValue();
				this.Percent = val;
			} else if ("showValue" === reader.GetName()) {
				val = reader.GetValue();
				this.ShowValue = val;
			} else if ("custom" === reader.GetName()) {
				val = reader.GetValueBool();
				this.Custom = val;
			} else if ("reverse" === reader.GetName()) {
				val = reader.GetValue();
				this.Reverse = val;
			}
		}
	};

	AscCommonExcel.CIconSet.prototype.toXml = function (writer, bExtendedWrite) {
		if (this.aCFVOs.length < 2) {
			return;
		}	// min value = 2

		var node_name = bExtendedWrite ? "x14:iconSet" : "iconSet";

		writer.WriteXmlString("<" + node_name);
		writer.WriteXmlNullableAttributeString("iconSet", ToXml_ST_IconSetType(this.IconSet));
		if (false === this.Percent) {
			writer.WriteXmlString(" percent=\"0\"");
		}
		if (true === this.Reverse) {
			writer.WriteXmlString(" reverse=\"1\"");
		}
		if (false === this.ShowValue) {
			writer.WriteXmlString(" showValue=\"0\"");
		}
		if (bExtendedWrite && this.aIconSets && this.aIconSets.length) {
			writer.WriteXmlString(" custom=\"1\"");
		}
		writer.WriteXmlString(">");

		var i;
		for (i = 0; i < this.aCFVOs.length; ++i) {
			if (this.aCFVOs[i]) {
				this.aCFVOs[i].toXml(writer, bExtendedWrite);
			}
		}
		for (i = 0; bExtendedWrite && i < this.aIconSets.length; ++i) {
			if (this.aIconSets[i]) {
				this.aIconSets[i].toXml(writer, bExtendedWrite);
			}
		}

		writer.WriteXmlString("</" + node_name + ">");
	};

	AscCommonExcel.CIconSet.prototype.isExtended = function () {
		return 0 !== this.aIconSets.length || (this.IconSet !== Asc.EIconSetType.Traffic3Lights1 && this.IconSet > 15);
	};

	AscCommonExcel.CConditionalFormatIconSet.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};

	AscCommonExcel.CConditionalFormatIconSet.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("iconSet" === reader.GetName()) {
				val = reader.GetValue();
				this.IconSet = FromXml_ST_IconSetType(val);
			} else if ("iconId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.IconId = val;
			}
		}
	};

	AscCommonExcel.CConditionalFormatIconSet.prototype.toXml = function (writer, bExtendedWrite) {
		if (!bExtendedWrite) {
			return;
		}

		writer.WriteXmlString("<x14:cfIcon");
		writer.WriteXmlNullableAttributeString("iconSet", ToXml_ST_IconSetType(this.IconSet));
		writer.WriteXmlNullableAttributeNumber("iconId", this.IconId);
		writer.WriteXmlString("/>");
	};


	//SheetPr
	AscCommonExcel.asc_CSheetPr.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("tabColor" === name) {
				var color = AscCommon.getColorFromXml2(reader);
				if (color) {
					this.TabColor = color;
				}
			} else if ("pageSetUpPr" === name) {
				val = new CPageSetUpPr(this);
				val.SheetPr = this;
				val.fromXml(reader);
			} else if ("outlinePr" === name) {
				val = new COutlinePr();
				val.outlinePr = this;
				val.fromXml(reader);
			}
		}
	};

	AscCommonExcel.asc_CSheetPr.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("codeName" === reader.GetName()) {
				val = reader.GetValue();
				this.CodeName = val;
			} else if ("enableFormatConditionsCalculation" === reader.GetName()) {
				val = reader.GetValueBool();
				this.EnableFormatConditionsCalculation = val;
			} else if ("filterMode" === reader.GetName()) {
				val = reader.GetValueBool();
				this.FilterMode = val;
			} else if ("published" === reader.GetName()) {
				val = reader.GetValueBool();
				this.Published = val;
			} else if ("syncHorizontal" === reader.GetName()) {
				val = reader.GetValueBool();
				this.SyncHorizontal = val;
			} else if ("syncRef" === reader.GetName()) {
				val = reader.GetValue();
				this.SyncRef = val;
			} else if ("syncVertical" === reader.GetName()) {
				val = reader.GetValueBool();
				this.SyncVertical = val;
			} else if ("transitionEntry" === reader.GetName()) {
				val = reader.GetValueBool();
				this.TransitionEntry = val;
			} else if ("transitionEvaluation" === reader.GetName()) {
				val = reader.GetValueBool();
				this.TransitionEvaluation = val;
			}
		}
	};

	AscCommonExcel.asc_CSheetPr.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*sheetPr*/);
		writer.WriteXmlNullableAttributeStringEncode("codeName", this.CodeName);
		writer.WriteXmlNullableAttributeBool("enableFormatConditionsCalculation", this.EnableFormatConditionsCalculation);
		writer.WriteXmlNullableAttributeBool("filterMode", this.FilterMode);
		writer.WriteXmlNullableAttributeBool("published", this.Published);
		writer.WriteXmlNullableAttributeBool("syncHorizontal", this.SyncHorizontal);
		writer.WriteXmlNullableAttributeStringEncode("syncRef", this.SyncRef);
		writer.WriteXmlNullableAttributeBool("syncVertical", this.SyncVertical);
		writer.WriteXmlNullableAttributeBool("transitionEntry", this.TransitionEntry);
		writer.WriteXmlNullableAttributeBool("transitionEvaluation", this.TransitionEvaluation);
		writer.WriteXmlAttributesEnd();


		var val;
		if (this.TabColor) {
			AscCommon.writeColorToXml(writer, "tabColor", this.TabColor);
		}
		//пока не делаю отедльных классов для модели. нужно сделать + изменить serialize + history
		val = new COutlinePr(this);
		val.toXml(writer, "outlinePr");

		val = new CPageSetUpPr(this);
		val.toXml(writer, "pageSetUpPr");

		writer.WriteXmlNodeEnd(ns + name);
	};

	function CPageSetUpPr(sheetPr) {
		this.SheetPr = sheetPr;
	}

	CPageSetUpPr.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};
	CPageSetUpPr.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("autoPageBreaks" === reader.GetName()) {
				val = reader.GetValueBool();
				this.SheetPr.AutoPageBreaks = val;
			} else if ("fitToPage" === reader.GetName()) {
				val = reader.GetValueBool();
				this.SheetPr.FitToPage = val;
			}
		}
	};

	CPageSetUpPr.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeBool("autoPageBreaks", this.SheetPr.AutoPageBreaks);
		writer.WriteXmlNullableAttributeBool("fitToPage", this.SheetPr.FitToPage);
		writer.WriteXmlAttributesEnd(true);
	};

	function COutlinePr(outlinePr) {
		this.outlinePr = outlinePr;
	}

	COutlinePr.prototype.fromXml = function (reader) {

		/*ReadAttributes( oReader );

						if ( !oReader.IsEmptyNode() )
							oReader.ReadTillEnd();*/

		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	COutlinePr.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("applyStyles" === reader.GetName()) {
				val = reader.GetValueBool();
				this.outlinePr.ApplyStyles = val;
			} else if ("showOutlineSymbols" === reader.GetName()) {
				val = reader.GetValueBool();
				this.outlinePr.ShowOutlineSymbols = val;
			} else if ("summaryBelow" === reader.GetName()) {
				val = reader.GetValueBool();
				this.outlinePr.SummaryBelow = val;
			} else if ("summaryRight" === reader.GetName()) {
				val = reader.GetValueBool();
				this.outlinePr.SummaryRight = val;
			}
		}
	};

	COutlinePr.prototype.toXml = function (writer, name, ns) {
		//outlinePr
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeBool("applyStyles", this.outlinePr.ApplyStyles);
		writer.WriteXmlNullableAttributeBool("summaryBelow", this.outlinePr.SummaryBelow);
		writer.WriteXmlNullableAttributeBool("summaryRight", this.outlinePr.SummaryRight);
		writer.WriteXmlNullableAttributeBool("showOutlineSymbols", this.outlinePr.ShowOutlineSymbols);
		writer.WriteXmlAttributesEnd(true);
	};


	//***STYLE****
	AscCommonExcel.CT_Stylesheet.prototype.fromXml = function (reader) {
		if (!reader.ReadNextNode()) {
			return;
		}

		var wb = reader.context.InitOpenManager.wb;
		var sName = reader.GetNameNoNS();
		var val, depth2, name2, t = this;
		if ("styleSheet" === sName) {
			if (reader.IsEmptyNode()) {
				return;
			}
			var depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				var name = reader.GetNameNoNS();
				if ("borders" === name) {
					if (reader.IsEmptyNode()) {
						continue;
					}
					depth2 = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depth2)) {
						name2 = reader.GetNameNoNS();
						if ("border" === name2 || "Border" === name2) {
							val = new AscCommonExcel.Border();
							val.fromXml(reader);
							this.borders.push(val);
						}
					}
				} else if ("cellStyles" === name) {
					reader.readXmlArray("cellStyle", function () {
						val = new AscCommonExcel.CCellStyle();
						val.fromXml(reader);
						t.cellStyles.push(val);
					});
				} else if ("cellStyleXfs" === name) {
					if (reader.IsEmptyNode()) {
						continue;
					}
					depth2 = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depth2)) {
						name2 = reader.GetNameNoNS();
						if ("xf" === name2) {
							val = new AscCommonExcel.OpenXf();
							val.fromXml(reader);
							this.cellStyleXfs.push(val);
						}
					}
				} else if ("cellXfs" === name) {
					if (reader.IsEmptyNode()) {
						continue;
					}
					depth2 = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depth2)) {
						name2 = reader.GetNameNoNS();
						if ("xf" === name2) {
							val = new AscCommonExcel.OpenXf();
							val.fromXml(reader);
							this.cellXfs.push(val);
						}
					}
				} else if ("colors" === name) {
					//не вижу в serialize
				} else if ("dxfs" === name) {
					if (reader.IsEmptyNode()) {
						continue;
					}
					depth2 = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depth2)) {
						name2 = reader.GetNameNoNS();
						if ("dxf" === name2) {
							val = new AscCommonExcel.CellXfs();
							val.fromXml(reader);
							this.dxfs.push(val);
						}
					}
				} else if ("fills" === name) {
					if (reader.IsEmptyNode()) {
						continue;
					}
					depth2 = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depth2)) {
						name2 = reader.GetNameNoNS();
						if ("fill" === name2) {
							val = new AscCommonExcel.Fill();
							val.fromXml(reader);
							this.fills.push(val);
						}
					}
				} else if ("fonts" === name) {
					if (reader.IsEmptyNode()) {
						continue;
					}
					depth2 = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depth2)) {
						name2 = reader.GetNameNoNS();
						if ("font" === name2) {
							val = new AscCommonExcel.Font();
							val.fromXml(reader);
							if (wb) {
								val.checkSchemeFont(wb.theme);
							}
							this.fonts.push(val);
						}
					}
				} else if ("numFmts" === name) {
					if (reader.IsEmptyNode()) {
						continue;
					}
					depth2 = reader.GetDepth();
					while (reader.ReadNextSiblingNode(depth2)) {
						name2 = reader.GetNameNoNS();
						if ("numFmt" === name2) {
							val = {f: null, id: null};
							while (reader.MoveToNextAttribute()) {
								if ("formatCode" === reader.GetName()) {
									val.f = reader.GetValueDecodeXml();
									val.f = prepareTextFormatFromXml(val.f);
								} else if ("numFmtId" === reader.GetName()) {
									val.id = reader.GetValueInt();
								}
							}

							//TODO parseNum
							if (null != val.id) {
								AscCommonExcel.InitOpenManager.prototype.ParseNum.call(this, val, this.numFmts/*, this.useNumId*/);
							}
						}
					}
				} else if ("tableStyles" === name) {
					//CTableStyles
					this.tableStyles.fromXml(reader);
				} else if ("Style" === name) {
					//TODO
				} else if ("extLst" === name) {
					var extLst = new COfficeArtExtensionList(this);
					extLst.fromXml(reader);
					this.oCustomSlicerStyles = extLst.getSlicerStyles();
					//не нашёл файла, где можно проверить
					this.aExtDxfs = extLst.getDxfs();
				}
			}
		}
	};

	AscCommonExcel.CCellStyle.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	AscCommonExcel.CCellStyle.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("builtinId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.BuiltinId = val;
			} else if ("customBuiltin" === reader.GetName()) {
				val = reader.GetValueBool();
				this.CustomBuiltin = val;
			} else if ("hidden" === reader.GetName()) {
				val = reader.GetValueBool();
				this.Hidden = val;
			} else if ("iLevel" === reader.GetName()) {
				val = reader.GetValueInt();
				this.ILevel = val;
			} else if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.Name = val;
			} else if ("xfId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.XfId = val;
			}
		}
	};

	AscCommonExcel.CCellStyle.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);

		/*<xsd:attribute name="name" type="s:ST_Xstring" use="optional"/>
		3640 <xsd:attribute name="xfId" type="ST_CellStyleXfId" use="required"/>
			3641 <xsd:attribute name="builtinId" type="xsd:unsignedInt" use="optional"/>
			3642 <xsd:attribute name="iLevel" type="xsd:unsignedInt" use="optional"/>
			3643 <xsd:attribute name="hidden" type="xsd:boolean" use="optional"/>
			3644 <xsd:attribute name="customBuiltin" type="xsd:boolean" use="optional"/>*/

		//в x2t пишется данных меньше, чем читается
		writer.WriteXmlNullableAttributeStringEncode("name", this.Name);
		writer.WriteXmlNullableAttributeNumber("xfId", this.XfId);
		writer.WriteXmlNullableAttributeNumber("builtinId", this.BuiltinId);

		//не пишеится в x2t
		writer.WriteXmlNullableAttributeNumber("iLevel", this.ILevel);
		//не пишеится в x2t
		writer.WriteXmlNullableAttributeBool("hidden", this.Hidden);
		//не пишеится в x2t
		writer.WriteXmlNullableAttributeBool("customBuiltin", this.CustomBuiltin);


		writer.WriteXmlAttributesEnd();

		writer.WriteXmlNodeEnd(ns + name);
	};

	AscCommonExcel.OpenXf.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("alignment" === name) {
				if (null == this.align) {
					this.align = new AscCommonExcel.Align();
				}
				this.align.fromXml(reader);
			} else if ("protection" === name) {
				while (reader.MoveToNextAttribute()) {
					if ("hidden" === reader.GetName()) {
						this.hidden = reader.GetValueBool();
					} else if ("locked" === reader.GetName()) {
						this.locked = reader.GetValueBool();
					}
				}
			}
		}
	};

	Asc.CTableStyles.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("tableStyle" === name) {
				var val = new Asc.CTableStyle();
				var aElements = [];
				val.fromXml(reader, aElements);
				if (null != val.name) {
					if (null === val.displayName) {
						val.displayName = val.name;
					}
					this.CustomStyles[val.name] = {style: val, elements: aElements};
				}

				/*if(null != oNewStyle.name) {
					if (null === oNewStyle.displayName)
						oNewStyle.displayName = oNewStyle.name;
					oCustomStyles[oNewStyle.name] = {style : oNewStyle, elements: aElements};
				}*/
			}
		}
	};

	Asc.CTableStyles.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("defaultPivotStyle" === reader.GetName()) {
				val = reader.GetValue();
				this.DefaultPivotStyle = val;
			} else if ("defaultTableStyle" === reader.GetName()) {
				val = reader.GetValue();
				this.DefaultTableStyle = val;
			}
		}
	};

	Asc.CTableStyles.prototype.toXml = function (writer, name, ns) {
		var length = 0, i;
		for (i in this.CustomStyles) {
			length++;
		}

		if (!ns) {
			ns = "";
		}

		//tableStyles
		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeNumber("count", length);
		writer.WriteXmlNullableAttributeStringEncode("defaultTableStyle", this.DefaultTableStyle);
		writer.WriteXmlNullableAttributeStringEncode("defaultPivotStyle", this.DefaultPivotStyle);


		var bEmptyCustom = true;
		for (var j in this.CustomStyles) {
			bEmptyCustom = false;
			break;
		}
		if (!bEmptyCustom) {
			writer.WriteXmlAttributesEnd();

			for (i in this.CustomStyles) {
				var style = this.CustomStyles[i];
				writer.context.InitSaveManager.WriteTableCustomStyleElements(this.CustomStyles[i], function (type, tableStyleElement) {
					style.toXml(writer, "tableStyle", "", type, [tableStyleElement]);
				});

			}

			writer.WriteXmlNodeEnd(ns + name);
		} else {
			writer.WriteXmlAttributesEnd(true);
		}
	};

	Asc.CTableStyle.prototype.fromXml = function (reader, aElements) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		var val;
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("tableStyleElement" === name) {
				var tableStyleElement = {Type: null, Size: null, DxfId: null};

				while (reader.MoveToNextAttribute()) {
					if ("dxfId" === reader.GetName()) {
						val = reader.GetValueInt();
						tableStyleElement.DxfId = val;
					} else if ("size" === reader.GetName()) {
						val = reader.GetValueInt();
						tableStyleElement.Size = val;
					} else if ("type" === reader.GetName()) {
						val = reader.GetValue();
						tableStyleElement.Type = FromXml_ST_TableStyleType(val);
					}
				}

				/*var oNewStyleElement = {Type: null, Size: null, DxfId: null};
				res = this.bcr.Read2Spreadsheet(length, function(t,l){
					return oThis.ReadTableCustomStyleElement(t,l, oNewStyleElement);
				});
				if(null != oNewStyleElement.Type && null != oNewStyleElement.DxfId)
					aElements.push(oNewStyleElement);*/

				/*if (c_oSer_TableStyleElement.Type === type)
					oNewStyleElement.Type = this.stream.GetUChar();
				else if (c_oSer_TableStyleElement.Size === type)
					oNewStyleElement.Size = this.stream.GetULongLE();
				else if (c_oSer_TableStyleElement.DxfId === type)
					oNewStyleElement.DxfId = this.stream.GetULongLE();
				else
					res = c_oSerConstants.ReadUnknown;*/

				aElements.push(tableStyleElement);
			}
		}
	};

	Asc.CTableStyle.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.name = val;
			} else if ("pivot" === reader.GetName()) {
				val = reader.GetValueBool();
				this.pivot = val;
			} else if ("table" === reader.GetName()) {
				val = reader.GetValueBool();
				this.table = val;
			} else if ("displayName" === reader.GetName()) {
				val = reader.GetValue();
				this.displayName = val;
			}
		}
	};

	Asc.CTableStyle.prototype.toXml = function (writer, name, ns, type, aElements) {
		if (this.name && aElements && aElements.length) {
			if (!ns) {
				ns = "";
			}

			//tableStyle
			writer.WriteXmlNodeStart(ns + name);
			writer.WriteXmlNullableAttributeStringEncode("name", this.name);
			writer.WriteXmlNullableAttributeBool("table", this.table);
			writer.WriteXmlNullableAttributeBool("pivot", this.pivot);
			writer.WriteXmlNullableAttributeNumber("count", aElements.length);
			writer.WriteXmlAttributesEnd();

			for (var i = 0; i < aElements.length; ++i) {
				if (aElements[i]) {
					//tableStyleElement
					writer.WriteXmlNodeStart("tableStyleElement");
					writer.WriteXmlAttributeString("type", ToXml_ST_TableStyleType(type));
					writer.WriteXmlNullableAttributeNumber("size", aElements[i].size);

					var dxfId = writer.context.tableStylesMap[this.displayName];
					if (null != aElements[i].dxf) {
						writer.WriteXmlAttributeNumber("dxfId", dxfId);
					}
					writer.WriteXmlAttributesEnd(true);
				}
			}

			writer.WriteXmlNodeEnd(ns + name);
		}
	};

	AscCommonExcel.OpenXf.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("applyAlignment" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ApplyAlignment = val;
			} else if ("applyBorder" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ApplyBorder = val;
			} else if ("applyFill" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ApplyFill = val;
			} else if ("applyFont" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ApplyFont = val;
			} else if ("applyNumberFormat" === reader.GetName()) {
				val = reader.GetValueBool();
				this.ApplyNumberFormat = val;
			} else if ("applyProtection" === reader.GetName()) {
				val = reader.GetValueBool();
				this.applyProtection = val;
			} else if ("borderId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.borderid = val;
			} else if ("fillId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.fillid = val;
			} else if ("fontId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.fontid = val;
			} else if ("numFmtId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.numid = val;
			} else if ("pivotButton" === reader.GetName()) {
				val = reader.GetValueBool();
				this.PivotButton = val;
			} else if ("quotePrefix" === reader.GetName()) {
				val = reader.GetValueBool();
				this.QuotePrefix = val;
			} else if ("xfId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.XfId = val;
			}
		}
	};

	AscCommonExcel.BorderProp.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("color" === name) {
				//ReadColorSpreadsheet2 - возможно стоит объединить
				this.c = AscCommon.getColorFromXml2(reader);
			}
		}
	};

	AscCommonExcel.BorderProp.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("style" === reader.GetName()) {
				val = reader.GetValue();
				this.setStyle(AscCommonExcel.FromXml_ST_BorderStyle(val));
			} /*else if ("ss:Color" === reader.GetName()) {
				val = reader.GetValue();

			} else if ("ss:LineStyle" === reader.GetName()) {

			}*/
		}
	};

	AscCommonExcel.BorderProp.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeString("style", AscCommonExcel.ToXml_ST_BorderStyle(this.s));
		writer.WriteXmlAttributesEnd();


		//в x2t используется toXMLWithNS

		if (this.c && this.s !== Asc.c_oAscBorderStyles.None) {
			AscCommon.writeColorToXml(writer, "color", this.c, childns);
		}

		writer.WriteXmlNodeEnd(ns + name);
	};


	AscCommonExcel.Border.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("bottom" === name) {
				this.b.fromXml(reader);
			} else if ("diagonal" === name) {
				this.d.fromXml(reader);
			} else if ("end" === name || "right" === name) {
				this.r.fromXml(reader);
			} else if ("horizontal" === name) {
				this.ih.fromXml(reader);
			} else if ("start" === name || "left" === name) {
				this.l.fromXml(reader);
			} else if ("top" === name) {
				this.t.fromXml(reader);
			} else if ("vertical" === name) {
				this.iv.fromXml(reader);
			}
		}
	};

	AscCommonExcel.Border.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("diagonalDown" === reader.GetName()) {
				val = reader.GetValueBool();
				this.dd = val;
			} else if ("diagonalUp" === reader.GetName()) {
				val = reader.GetValueBool();
				this.du = val;
			} else if ("outline" === reader.GetName()) {
				val = reader.GetValueBool();
				this.outline = val;
			}
		}
	};

	AscCommonExcel.Border.prototype.toXml = function (writer, name, ns, childns) {

		//убрал проверку на дефолтовый бордер, поскольку в файл записываются у нас всё
		var border = this;/*this.getDif(AscCommonExcel.g_oDefaultFormat.BorderAbs);*/

		if (!border) {
			return;
		}

		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		if (border.dd) {
			writer.WriteXmlAttributeString("diagonalDown", border.dd);
		}
		if (border.du) {
			writer.WriteXmlAttributeString("diagonalUp", border.du);
		}
		writer.WriteXmlAttributesEnd();


		if (null != border.l) {
			border.l.toXml(writer, "left", childns, childns);
		} else {
			writer.WriteXmlString("<left/>");
		}

		if (null != border.r) {
			border.r.toXml(writer, "right", childns, childns);
		} else {
			writer.WriteXmlString("<right/>");
		}

		if (null != border.t) {
			border.t.toXml(writer, "top", childns, childns);
		} else {
			writer.WriteXmlString("<top/>");
		}

		if (null != border.b) {
			border.b.toXml(writer, "bottom", childns, childns);
		} else {
			writer.WriteXmlString("<bottom/>");
		}

		if (null != border.d) {
			border.d.toXml(writer, "diagonal", childns, childns);
		} else {
			writer.WriteXmlString("<diagonal/>");
		}

		if (null != border.iv && Asc.c_oAscBorderStyles.None !== border.iv.s) {
			border.iv.toXml(writer, "vertical", childns, childns);
		}
		if (null != border.ih && Asc.c_oAscBorderStyles.None !== border.ih.s) {
			border.ih.toXml(writer, "horizontal", childns, childns);
		}

		writer.WriteXmlNodeEnd(ns + name);
	};

	AscCommonExcel.Align.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	AscCommonExcel.Align.prototype.readAttr = function (reader) {
		//TODO метрика
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("horizontal" === reader.GetName() || "ss:Horizontal" === reader.GetName()) {
				val = reader.GetValue();
				this.hor = FromXml_ST_HorizontalAlignment(val, true);
			} else if ("indent" === reader.GetName() || "ss:Indent" === reader.GetName()) {
				val = reader.GetValueInt();
				this.indent = val;
			} else if ("justifyLastLine" === reader.GetName()) {
				val = reader.GetValueBool();
				this.justifyLastLine = val;
			} else if ("readingOrder" === reader.GetName() || "ss:ReadingOrder" === reader.GetName()) {
				//val = reader.GetValueInt();
				//this.readingOrder = val;
			} else if ("relativeIndent" === reader.GetName()) {
				val = reader.GetValueInt();
				this.RelativeIndent = val;
			} else if ("shrinkToFit" === reader.GetName() || "ss:ShrinkToFit" === reader.GetName()) {
				val = reader.GetValueBool();
				this.shrink = val;
			} else if ("textRotation" === reader.GetName() || "ss:Rotate" === reader.GetName()) {
				val = reader.GetValueInt();
				this.angle = val;
			} else if ("vertical" === reader.GetName() || "ss:Vertical" === reader.GetName()) {
				val = reader.GetValue();
				this.ver = FromXml_ST_VerticalAlignment(val);
			} else if ("wrapText" === reader.GetName() || "ss:WrapText" === reader.GetName()) {
				val = reader.GetValueBool();
				this.wrap = val;
			}
		}
	};


	AscCommonExcel.Align.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeString("horizontal", ToXml_ST_HorizontalAlignment(this.hor, true));
		writer.WriteXmlNullableAttributeNumber("indent", this.indent !== 0 ? this.indent : null);
		writer.WriteXmlNullableAttributeBool("justifyLastLine", this.justifyLastLine);
		//writer.WriteXmlNullableAttributeNumber("readingOrder", this.readingOrder);
		writer.WriteXmlNullableAttributeNumber("relativeIndent", this.RelativeIndent !== 0 ? this.RelativeIndent : null);
		writer.WriteXmlNullableAttributeBool("shrinkToFit", this.shrink !== false ? this.angle : null);
		writer.WriteXmlNullableAttributeNumber("textRotation", this.angle !== 0 ? this.angle : null);
		writer.WriteXmlNullableAttributeString("vertical", this.ver !== 0 ? ToXml_ST_VerticalAlignment(this.ver, true) : null);
		writer.WriteXmlNullableAttributeBool("wrapText", this.wrap !== false ? this.wrap : null);
		writer.WriteXmlAttributesEnd(true);
	};

	AscCommonExcel.CellXfs.prototype.fromXml = function (reader) {
		if (reader.IsEmptyNode()) {
			return;
		}

		//TODO applyProtection
		var wb = reader.context.wb;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			var val;
			if ("alignment" === name) {
				val = new AscCommonExcel.Align();
				val.fromXml(reader);
				this.align = val;
			} else if ("border" === name) {
				val = new AscCommonExcel.Border();
				val.fromXml(reader);
				this.border = val;
			} else if ("fill" === name) {
				val = new AscCommonExcel.Fill();
				val.fromXml(reader);
				val.fixForDxf();
				this.fill = val;
			} else if ("font" === name) {
				val = new AscCommonExcel.Font();
				val.fromXml(reader);
				if (wb) {
					val.checkSchemeFont(wb.theme);
				}
				this.font = val;
			} else if ("numFmt" === name) {
				val = new AscCommonExcel.Num();
				val.fromXml(reader);
				this.num = val;
			} else if ("protection" === name) {
				while (reader.MoveToNextAttribute()) {
					if ("hidden" === reader.GetName()) {
						this.hidden = val;
					} else if ("locked" === reader.GetName()) {
						this.locked = val;
					}
				}
			}
		}
	};

	AscCommonExcel.CellXfs.prototype.toXml = function (writer, name, ns, childns, index) {
		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}
		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlAttributesEnd();


		if (this.font) {
			this.font.toXml(writer, "font", childns, childns);
		}
		if (this.num) {
			this.num.toXml(writer, "numFmt", null, childns, childns, writer.context.mapIndexNumId && writer.context.mapIndexNumId[index]);
		}
		if (this.fill) {
			this.fill.toXml(writer, "fill", childns, childns);
		}
		if (this.align) {
			this.align.toXml(writer, "alignment", childns, childns);
		}
		if (this.border) {
			this.border.toXml(writer, "border", childns, childns);
		}
		if (null != this.locked || null != this.hidden) {
			writer.WriteXmlString("<" + childns + "protection");
			writer.WriteXmlNullableAttributeBool("hidden", this.hidden);
			writer.WriteXmlNullableAttributeBool("locked", this.locked);
			writer.WriteXmlString("/>");
		}

		writer.WriteXmlNodeEnd(ns + name);
	};

	AscCommonExcel.Fill.prototype.fromXml = function (reader) {
		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		var val;
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("patternFill" === name) {
				val = new AscCommonExcel.PatternFill();
				val.fromXml(reader);
				this.patternFill = val;
			} else if ("gradientFill" === name) {
				val = new AscCommonExcel.GradientFill();
				val.fromXml(reader);
				this.gradientFill = val;
			}
		}
	};

	AscCommonExcel.Fill.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlAttributesEnd();


		//в x2t используется toXMLWithNS
		/*if(m_oPatternFill.IsInit())
			m_oPatternFill->toXMLWithNS(writer, child_ns, L"patternFill", child_ns);
		if(m_oGradientFill.IsInit())
			m_oGradientFill->toXMLWithNS(writer, child_ns, L"gradientFill", child_ns);*/

		if (this.patternFill) {
			this.patternFill.toXml(writer, "patternFill", childns, childns);
		}
		if (this.gradientFill) {
			this.gradientFill.toXml(writer, "gradientFill", childns, childns);
		}

		writer.WriteXmlNodeEnd(ns + name);
	};

	AscCommonExcel.PatternFill.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			var val;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("bgColor" === name) {
				this.bgColor = AscCommon.getColorFromXml2(reader);
			} else if ("fgColor" === name) {
				this.fgColor = AscCommon.getColorFromXml2(reader);
			}
			//PatternBgColor_deprecated ?
		}
	};

	AscCommonExcel.PatternFill.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("patternType" === reader.GetName()) {
				val = reader.GetValue();
				val = AscCommonExcel.FromXml_ST_PatternType(val);
				if (-1 !== val) {
					this.patternType = val;
				}
			}
		}
	};

	AscCommonExcel.PatternFill.prototype.toXml = function (writer, name, ns, childns) {

		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}

		/*writer.StartNodeWithNS(node_ns, node_name);
		writer.StartAttributes();
		WritingStringNullableAttrString(L"patternType", m_oPatternType, m_oPatternType->ToString());
		if(m_oBgColor.IsInit() || m_oFgColor.IsInit())
		{
			writer.EndAttributes();
			if(m_oBgColor.IsInit() && m_oFgColor.IsInit())
			{
				m_oFgColor->toXMLWithNS(writer, child_ns, L"fgColor", child_ns);
				m_oBgColor->toXMLWithNS(writer, child_ns, L"bgColor", child_ns);
			}
			else if(m_oFgColor.IsInit())
			{
				m_oFgColor->toXMLWithNS(writer, child_ns, L"fgColor", child_ns);
				m_oFgColor->toXMLWithNS(writer, child_ns, L"bgColor", child_ns);
			}
			else if(m_oBgColor.IsInit())
			{
				m_oBgColor->toXMLWithNS(writer, child_ns, L"fgColor", child_ns);
				m_oBgColor->toXMLWithNS(writer, child_ns, L"bgColor", child_ns);
			}

			writer.EndNodeWithNS(node_ns, node_name);*/

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeString("patternType", AscCommonExcel.ToXml_ST_PatternType(this.patternType));
		writer.WriteXmlAttributesEnd();


		//в x2t используется toXMLWithNS

		if (this.fgColor) {
			AscCommon.writeColorToXml(writer, "fgColor", this.fgColor, childns);
		}
		if (this.bgColor) {
			AscCommon.writeColorToXml(writer, "bgColor", this.bgColor, childns);
		}

		writer.WriteXmlNodeEnd(ns + name);
	};

	AscCommonExcel.GradientFill.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("stop" === name) {
				var val = new AscCommonExcel.GradientStop();
				val.fromXml(reader);
				this.stop.push(val);
			}
		}
	};

	AscCommonExcel.GradientFill.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("bottom" === reader.GetName()) {
				val = reader.GetValue();
				this.bottom = val;
			} else if ("degree" === reader.GetName()) {
				val = reader.GetValue();
				this.degree = val;
			} else if ("left" === reader.GetName()) {
				val = reader.GetValue();
				this.left = val;
			} else if ("right" === reader.GetName()) {
				val = reader.GetValue();
				this.right = val;
			} else if ("top" === reader.GetName()) {
				val = reader.GetValue();
				this.top = val;
			} else if ("type" === reader.GetName()) {
				val = reader.GetValue();
				this.type = FromXml_ST_GradientType(val, true);
			}
		}
	};

	AscCommonExcel.GradientFill.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeString("type", ToXml_ST_GradientType(this.type, true));
		writer.WriteXmlNullableAttributeDouble("left", this.left);
		writer.WriteXmlNullableAttributeDouble("right", this.right);
		writer.WriteXmlNullableAttributeDouble("top", this.top);
		writer.WriteXmlNullableAttributeDouble("bottom", this.bottom);
		writer.WriteXmlNullableAttributeDouble("degree", this.degree);
		writer.WriteXmlAttributesEnd();

		writer.WriteXmlArray(this.stop, "stop", null, null, childns, childns);

		writer.WriteXmlNodeEnd(ns + name);
	};

	AscCommonExcel.GradientStop.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("color" === name) {
				this.color = AscCommon.getColorFromXml2(reader);
			}
		}
	};

	AscCommonExcel.GradientStop.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("position" === reader.GetName()) {
				val = reader.GetValue();
				this.position = val;
			}
		}
	};

	AscCommonExcel.GradientStop.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}
		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeDouble("position", this.position);
		writer.WriteXmlAttributesEnd();

		if (this.color) {
			AscCommon.writeColorToXml(writer, "color", this.color, childns);
		}

		writer.WriteXmlNodeEnd(ns + name);
	};


	AscCommonExcel.Font.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var t = this;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("b" === name) {
				if (reader.GetValue() !== null) {
					this.b = reader.GetValueBool();
				} else {
					this.b = true;
				}
			} else if ("charset" === name) {
			} else if ("color" === name) {
				this.c = AscCommon.getColorFromXml2(reader);
			} else if ("condense" === name) {
			} else if ("extend" === name) {
			} else if ("family" === name) {
			} else if ("i" === name) {
				if (reader.GetValue() !== null) {
					this.i = reader.GetValueBool();
				} else {
					this.i = true;
				}
			} else if ("name" === name || "rFont" === name) {
				readOneAttr(reader, "val", function () {
					t.fn = reader.GetValue();
				});
			} else if ("outline" === name) {
			} else if ("scheme" === name) {
				readOneAttr(reader, "val", function () {
					t.scheme = FromXml_ST_FontScheme(reader.GetValue());
				});
			} else if ("shadow" === name) {
			} else if ("strike" === name) {
				if (reader.GetValue() !== null) {
					this.s = reader.GetValueBool();
				} else {
					this.s = true;
				}
			} else if ("sz" === name) {
				readOneAttr(reader, "val", function () {
					t.fs = reader.GetValueDouble();
				});
			} else if ("u" === name) {
				if (reader.GetValue() !== null) {
					this.u = FromXml_ST_UnderlineValues(reader.GetValue());
				} else {
					this.u = Asc.EUnderline.underlineSingle;
				}
			} else if ("vertAlign" === name) {
				readOneAttr(reader, "val", function () {
					switch (reader.GetValue()) {
						case "baseline":
							t.va = AscCommon.vertalign_Baseline;
							break;
						case "superscript":
							t.va = AscCommon.vertalign_SuperScript;
							break;
						case "subscript":
							t.va = AscCommon.vertalign_SubScript;
							break;
					}
				});


			}
			//TODO skip, repeate
		}
	};

	AscCommonExcel.Font.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("ss:FontName" === reader.GetName()) {
			} else if ("x:Family" === reader.GetName()) {
			} else if ("x:CharSet" === reader.GetName()) {
			} else if ("ss:Size" === reader.GetName()) {
			} else if ("ss:Color" === reader.GetName()) {
			} else if ("ss:Underline" === reader.GetName()) {
			} else if ("ss:Bold" === reader.GetName()) {
			} else if ("ss:Italic" === reader.GetName()) {
			} else if ("ss:Italic" === reader.GetName()) {
			} else if ("ss:VerticalAlign" === reader.GetName()) {
			}
		}
	};

	AscCommonExcel.Font.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlAttributesEnd();


		/*if(m_oCharset.IsInit() && m_oCharset->m_oCharset.IsInit())
		{
			WritingValNode("charset", m_oCharset->m_oCharset->ToString());
		}
		if(m_oFamily.IsInit() && m_oFamily->m_oFontFamily.IsInit())
		{
			WritingValNode("family", m_oFamily->m_oFontFamily->ToString());
		}*/


		if (this.b != null) {
			writer.WritingValNodeIf(childns, "b", !this.b, "0");
		}
		if (this.i != null) {
			writer.WritingValNodeIf(childns, "i", !this.i, "0");
		}
		if (this.s != null) {
			writer.WritingValNodeIf(childns, "strike", !this.s, "0");
		}

		/*if(m_oOutline.IsInit())
		{
			WritingValNodeIf("outline", !m_oOutline->ToBool(), L"0");
		}
		if(m_oShadow.IsInit())
		{
			WritingValNodeIf("shadow", !m_oShadow->ToBool(), L"0");
		}
		if(m_oCondense.IsInit())
		{
			WritingValNodeIf("condense", !m_oCondense->ToBool(), L"0");
		}
		if(m_oExtend.IsInit())
		{
			WritingValNodeIf("extend", !m_oExtend->ToBool(), L"0");
		}*/


		/*if(m_oColor.IsInit())
			m_oColor->toXMLWithNS(writer, "color", child_ns);*/


		if (this.fs != null) {
			writer.WritingValNode(childns, "sz", this.fs + "");
		}

		if (this.c) {
			AscCommon.writeColorToXml(writer, "color", this.c, childns);
		}

		if (this.fn) {
			var fN = name === "rPr" ? "rFont" : "name";
			if (this.fn.length <= 31) {
				writer.WritingValNodeEncodeXml(childns, fN, this.fn);
			} else {
				writer.WritingValNodeEncodeXml(childns, fN, this.fn.substr(0, 31));
			}
		}

		if (this.u != null) {
			writer.WritingValNodeIf(childns, "u", this.u !== Asc.EUnderline.underlineSingle, ToXml_ST_UnderlineValues(this.u));
		}

		if (this.va != null) {
			var val = null;
			if (this.va === AscCommon.vertalign_SuperScript) {
				val = "superscript";
			} else if (this.va === AscCommon.vertalign_SubScript) {
				val = "subscript";
			} else if (this.va === AscCommon.vertalign_Baseline) {
				val = "baseline";
			}
			if (val) {
				writer.WritingValNode(childns, "vertAlign", val);
			}
		}
		var scheme = ToXml_ST_FontScheme(this.scheme);
		if (scheme !== null) {
			writer.WritingValNode(childns, "scheme", scheme);
		}

		writer.WriteXmlNodeEnd(ns + name);
	};


	AscCommonExcel.StylesForWrite.prototype.toXml = function (writer) {
		var wb = writer.context.wb;

		/*//borders
		this.bs.WriteItem(c_oSerStylesTypes.Borders, function(){oThis.WriteBorders();});
		//fills
		this.bs.WriteItem(c_oSerStylesTypes.Fills, function(){oThis.WriteFills();});
		//fonts
		this.bs.WriteItem(c_oSerStylesTypes.Fonts, function(){oThis.WriteFonts();});
		//CellStyleXfs
		this.bs.WriteItem(c_oSerStylesTypes.CellStyleXfs, function(){oThis.WriteCellStyleXfs();});
		//cellxfs
		this.bs.WriteItem(c_oSerStylesTypes.CellXfs, function(){oThis.WriteCellXfs();});

		//CellStyles
		this.bs.WriteItem(c_oSerStylesTypes.CellStyles, function(){oThis.WriteCellStyles(wb.CellStyles.CustomStyles);});

		if(null != wb.TableStyles)
			this.bs.WriteItem(c_oSerStylesTypes.TableStyles, function(){oThis.WriteTableStyles(wb.TableStyles);});

		//Dxfs пишется после TableStyles, потому что Dxfs может пополниться при записи TableStyles
		var dxfs = this.InitSaveManager.getDxfs();
		if(null != dxfs && dxfs.length > 0) {
			this.bs.WriteItem(c_oSerStylesTypes.Dxfs, function(){oThis.WriteDxfs(dxfs);});
		}
		var aExtDxfs = [];
		var slicerStyles = this.PrepareSlicerStyles(wb.SlicerStyles, aExtDxfs);
		if(aExtDxfs.length > 0) {
			this.bs.WriteItem(c_oSerStylesTypes.ExtDxfs, function(){oThis.WriteDxfs(aExtDxfs);});
		}
		this.bs.WriteItem(c_oSerStylesTypes.SlicerStyles, function(){oThis.WriteSlicerStyles(slicerStyles);});
		//numfmts пишется в конце потому что они могут пополниться при записи Dxfs
		this.bs.WriteItem(c_oSerStylesTypes.NumFmts, function(){oThis.WriteNumFmts();});*/

		writer.WriteXmlString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
		writer.WriteXmlNodeStart("styleSheet");
		writer.WriteXmlString(' xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"');
		writer.WriteXmlString(' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"  mc:Ignorable="x14ac x16r2"');
		writer.WriteXmlString(
			' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"');
		writer.WriteXmlAttributesEnd();


		//делаю предаврительный map, потому что в serialize - numfmts пишется в конце потому что они могут пополниться при записи Dxfs"
		var dxfs = writer.context.InitSaveManager.getDxfs();
		var mapIndexNumId;
		for (var i = 0, length = dxfs.length; i < length; ++i) {
			if (null != dxfs[i].num) {
				var numId = writer.context.stylesForWrite.getNumIdByFormat(dxfs[i].num);
				if (null != numId) {
					if (!mapIndexNumId) {
						mapIndexNumId = [];
					}
					mapIndexNumId[i] = numId;
				}
			}
		}

		if (this.oNumMap.elems) {
			//AscCommonExcel.Num
			writer.WriteXmlArray(this.oNumMap.elems, "numFmt", "numFmts", true);
		}
		if (this.oFontMap.elems) {
			//AscCommonExcel.Font
			writer.WriteXmlArray(this.oFontMap.elems, "font", "fonts", true);
		}
		if (this.oFillMap.elems) {
			//AscCommonExcel.Fill
			writer.WriteXmlArray(this.oFillMap.elems, "fill", "fills", true);
		}
		if (this.oBorderMap.elems) {
			//AscCommonExcel.Border
			writer.WriteXmlArray(this.oBorderMap.elems, "border", "borders", true);
		}
		if (this.oXfsStylesMap) {
			//xfForWrite
			//добавляю в context по аналогии с записью в serialize, проверить!
			writer.context.isCellStyle = true;
			writer.WriteXmlArray(this.oXfsStylesMap, "xf", "cellStyleXfs", true);
			writer.context.isCellStyle = undefined;
		}
		if (this.oXfsMap.elems) {
			//xfForWrite
			writer.WriteXmlArray(this.oXfsMap.elems, "xf", "cellXfs", true);
		}

		if (wb.CellStyles.CustomStyles) {
			//AscCommonExcel.CCellStyle
			writer.WriteXmlArray(wb.CellStyles.CustomStyles, "cellStyle", "cellStyles", true);

		}

		//отсутвует в serialize, в x2t не пишется
		/*if(m_oColors.IsInit())
			m_oColors->toXML(writer);*/


		//делаю предаврительный map, потому что в serialize - Dxfs пишется после TableStyles, потому что Dxfs может пополниться при записи TableStyles
		var tableStylesMap = null;
		if (wb.TableStyles) {
			for (i in wb.TableStyles.CustomStyles) {
				var style = wb.TableStyles.CustomStyles[i];
				if (style) {
					var tableStyleElement = style.getTableStyleElement();
					if (tableStyleElement && tableStyleElement.dxf && null != dxfs) {
						if (!tableStylesMap) {
							tableStylesMap = {};
						}
						tableStylesMap[style.displayName] = dxfs.length;
						dxfs.push(tableStyleElement.dxf);
					}
				}
			}
		}


		//DXFS
		if (null != dxfs && dxfs.length > 0) {
			//AscCommonExcel.CellXfs. можно оберунть в dxf
			writer.context.mapIndexNumId = mapIndexNumId;
			writer.WriteXmlArray(dxfs, "dxf", "dxfs", true);
			writer.context.mapIndexNumId = null;
		}

		//пишется после TableStyles, потому что Dxfs может пополниться при записи TableStyles
		if (null != wb.TableStyles) {
			//Asc.CTableStyles
			writer.context.tableStylesMap = tableStylesMap;
			wb.TableStyles.toXml(writer, "tableStyles");
			writer.context.tableStylesMap = null;
		}

		var aExtDxfs = [];
		var slicerStyles = writer.context.InitSaveManager.PrepareSlicerStyles(wb.SlicerStyles, aExtDxfs);
		var officeArtExtensionList = new COfficeArtExtensionList(this);
		var officeArtExtension;
		if (aExtDxfs && aExtDxfs.length) {
			officeArtExtension = new COfficeArtExtension(this);
			officeArtExtension.uri = extUri.dxfs;
			officeArtExtension.dxfs = aExtDxfs;
			officeArtExtensionList.arrExt.push(officeArtExtension);
		}
		if (slicerStyles) {
			officeArtExtension = new COfficeArtExtension(this);
			officeArtExtension.uri = extUri.slicerStyles;
			officeArtExtension.slicerStyles = slicerStyles;
			officeArtExtensionList.arrExt.push(officeArtExtension);
		}
		officeArtExtensionList.toXml(writer);

		writer.WriteXmlString("</styleSheet>");
	};

	AscCommonExcel.Num.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	AscCommonExcel.Num.prototype.readAttr = function (reader) {
		var val;
		var sFormat = null;
		var id;
		while (reader.MoveToNextAttribute()) {
			if ("formatCode" === reader.GetName()) {
				val = reader.GetValue();
				if (undefined !== val) {
					sFormat = prepareTextFormatFromXml(val);
				}
			} else if ("numFmtId" === reader.GetName()) {
				val = reader.GetValue();
				if (undefined !== val) {
					id = val - 0;
				}
			} /*else if ("sourceLinked" === reader.GetName()) {
				val = reader.GetValue();
				this.sourceLinked = val;
			} else if ("ss:Format" === reader.GetName()) {
				val = reader.GetValue();
				this.ss:Format = val;
			}*/
		}

		this.f = null != sFormat ? sFormat : (AscCommonExcel.aStandartNumFormats[id] || "General");
		if ((5 <= id && id <= 8) || (14 <= id && id <= 17) || 22 === id || (27 <= id && id <= 31) || (36 <= id && id <= 44)) {
			this.id = id;
		}
	};

	AscCommonExcel.Num.prototype.toXml = function (writer, name, ns, childns, index, _id) {
		//запись и из dxfs и из stylesForWrite. в случае dxfs протаскиваем id через mapIndexNumId. для
		var id = _id ? _id : AscCommonExcel.g_nNumsMaxId + index;

		var format = this.getFormat();

		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeNumber("numFmtId", id);
		writer.WriteXmlNullableAttributeStringEncode("formatCode", format);
		//writer.WriteXmlNullableAttributeBool("sourceLinked", this.sourceLinked);
		writer.WriteXmlAttributesEnd(true);
	};

	AscCommonExcel.XfForWrite.prototype.toXml = function (writer, name, ns) {

		/*writer.WriteString(_T("<xf"));
						WritingStringNullableAttrInt(L"fontId", m_oFontId, m_oFontId->GetValue());
						WritingStringNullableAttrInt(L"fillId", m_oFillId, m_oFillId->GetValue());
						WritingStringNullableAttrInt(L"borderId", m_oBorderId, m_oBorderId->GetValue());
						WritingStringNullableAttrInt(L"numFmtId", m_oNumFmtId, m_oNumFmtId->GetValue());
						WritingStringNullableAttrInt(L"xfId", m_oXfId, m_oXfId->GetValue());
						WritingStringNullableAttrBool(L"applyNumberFormat", m_oApplyNumberFormat);
						WritingStringNullableAttrBool(L"applyFont", m_oApplyFont);
						WritingStringNullableAttrBool(L"applyFill", m_oApplyFill);
						WritingStringNullableAttrBool(L"applyBorder", m_oApplyBorder);
						WritingStringNullableAttrBool(L"applyAlignment", m_oApplyAlignment);
						WritingStringNullableAttrBool(L"applyProtection", m_oApplyProtection);
						WritingStringNullableAttrBool(L"quotePrefix", m_oQuotePrefix);
						WritingStringNullableAttrBool(L"pivotButton", m_oPivotButton);
						if (m_oAligment.IsInit() || m_oProtection.IsInit())
						{
							writer.WriteString(_T(">"));

							if (m_oAligment.IsInit())m_oAligment->toXML(writer);
							if (m_oProtection.IsInit())m_oProtection->toXML(writer);

							writer.WriteString(_T("</xf>"));
						}
						else
							writer.WriteString(_T("/>"));*/


		var isCellStyle = writer.context.isCellStyle;
		var xf = this.xf;

		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeNumber("fontId", this.fontid);
		writer.WriteXmlNullableAttributeNumber("fillId", this.fillid);
		writer.WriteXmlNullableAttributeNumber("borderId", this.borderid);
		writer.WriteXmlNullableAttributeNumber("numFmtId", this.numid);

		if (xf && !isCellStyle) {
			writer.WriteXmlNullableAttributeNumber("xfId", xf.XfId);
		}

		writer.WriteXmlNullableAttributeBool("applyNumberFormat", this.ApplyNumberFormat);
		writer.WriteXmlNullableAttributeBool("applyFont", this.ApplyFont);
		writer.WriteXmlNullableAttributeBool("applyFill", this.ApplyFill);
		writer.WriteXmlNullableAttributeBool("applyBorder", this.ApplyBorder);
		writer.WriteXmlNullableAttributeBool("applyAlignment", this.ApplyAlignment);

		if (xf) {
			writer.WriteXmlNullableAttributeBool("applyProtection", xf.applyProtection);
			writer.WriteXmlNullableAttributeBool("quotePrefix", xf.QuotePrefix);
			writer.WriteXmlNullableAttributeBool("pivotButton", xf.PivotButton);
		}


		var isProtection = xf && (null != xf.locked || null != xf.hidden);
		if (xf && (isProtection || xf.align)) {
			writer.WriteXmlAttributesEnd();
			if (xf.align) {
				xf.align.toXml(writer, "alignment");
			}
			if (isProtection) {
				writer.WriteXmlAttributesEnd();
				writer.WriteXmlNodeStart("protection");
				writer.WriteXmlNullableAttributeBool("hidden", xf.hidden);
				writer.WriteXmlNullableAttributeBool("locked", xf.locked);
				writer.WriteXmlAttributesEnd(true);
			}

			writer.WriteXmlNodeEnd(ns + name);
		} else {
			writer.WriteXmlAttributesEnd(true);
		}
	};


	//***External Reference****
	function CT_ExternalReferences(wb) {
		this.wb = wb;
		this.externalReferences = [];
	}

	CT_ExternalReferences.prototype.toXml = function (writer, name, ns) {
		var t = this;
		var context = writer.context;
		var ids = [];

		this.wb.externalReferences.forEach(function (externalReference) {
			//здесь пишем externalLink[]
			var oExternalReference = new CT_ExternalReference();
			oExternalReference.val = externalReference;

			var wsPart = context.part.addPart(AscCommon.openXml.Types.externalWorkbook);
			var rId = wsPart.part.addRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath", externalReference.Id, "External");
			var trueId = externalReference.Id;
			externalReference.Id = rId;

			//внутри дёргается toXml
			wsPart.part.setDataXml(oExternalReference, writer);

			externalReference.Id = trueId;
			ids.push(wsPart.rId);
		});

		if (!ns) {
			ns = "";
		}

		if (ids.length) {
			//externalReferences
			writer.WriteXmlNodeStart(ns + name);
			writer.WriteXmlAttributesEnd();
			ids.forEach(function (id) {
				//здесь пишем ссылку в externalReferences на externalLink[]
				writer.WriteXmlNodeStart("externalReference");
				writer.WriteXmlNullableAttributeString("r:id", id);
				writer.WriteXmlAttributesEnd(true);
			});
			writer.WriteXmlNodeEnd(ns + name);
		}
	};

	function CT_ExternalReference() {
		this.val = {};
	}

	CT_ExternalReference.prototype.fromXml = function (reader) {
		//не далаю отдельных классов для externalLink/externalBook
		if (reader.IsEmptyNode()) {
			return;
		}
		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("externalLink" === name) {

				if (reader.IsEmptyNode()) {
					return;
				}

				var depth2 = reader.GetDepth();
				while (reader.ReadNextSiblingNode(depth2)) {
					var name2 = reader.GetNameNoNS();
					if ("externalBook" === name2) {
						val = new CT_ExternalBook();
						val.fromXml(reader);
						this.val.externalBook = val.val;
					} else if ("oleLink" === name) {
						//TODO
						//хранится в бинарном виде
						//this.oWorkbook.externalReferences.push({Type: 1, Buffer: this.stream.GetBuffer(length)});
					} else if ("ddeLink" === name) {
						//TODO
						//хранится в бинарном виде
						//this.oWorkbook.externalReferences.push({Type: 2, Buffer: this.stream.GetBuffer(length)});
					}
				}
			}
		}
	};

	CT_ExternalReference.prototype.toXml = function (writer) {

		writer.WriteXmlString(("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"));
		writer.WriteXmlString(("<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"));

		if (this.val) {
			var externalBook = new CT_ExternalBook();
			externalBook.val = this.val;
			externalBook.toXml(writer);
		}
		/*if (m_oExternalBook.IsInit())
		{
			m_oExternalBook->toXML(sXml);
		}
		if (m_oOleLink.IsInit())
		{
			m_oOleLink->toXML(sXml);
		}
		if (m_oDdeLink.IsInit())
		{
			m_oDdeLink->toXML(sXml);
		}*/

		writer.WriteXmlString(("</externalLink>"));

	};

	function CT_ExternalBook() {
		this.val = {Type: 0, Id: null, SheetNames: [], DefinedNames: [], SheetDataSet: []};
	}

	CT_ExternalBook.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("sheetNames" === name) {
				val = new CT_ExternalSheetNames();
				val.fromXml(reader);
				if (val.val) {
					this.val.SheetNames = val.val;
				}
			} else if ("definedNames" === name) {
				val = new CT_ExternalDefinedNames();
				val.fromXml(reader);
				if (val.val) {
					this.val.DefinedNames = val.val;
				}
			} else if ("sheetDataSet" === name) {
				val = new CT_ExternalSheetDataSet();
				val.fromXml(reader);
				if (val.val) {
					this.val.SheetDataSet = val.val;
				}
			}
		}
	};
	CT_ExternalBook.prototype.readAttr = function (reader) {
		while (reader.MoveToNextAttribute()) {
			if ("id" === reader.GetNameNoNS()) {
				this.val.Id = reader.GetValue();
			}
		}
	};
	CT_ExternalBook.prototype.toXml = function (writer) {

		writer.WriteXmlString("<externalBook");
		if (this.val.Id) {
			writer.WriteXmlString(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
			writer.WriteXmlString(" r:id=\"");
			writer.WriteXmlString(this.val.Id);
			writer.WriteXmlString("\"");
		}
		writer.WriteXmlString(">");

		var val;
		if (this.val.SheetNames) {
			val = new CT_ExternalSheetNames();
			val.val = this.val.SheetNames;
			val.toXml(writer);
		}
		if (this.val.DefinedNames) {
			val = new CT_ExternalDefinedNames();
			val.val = this.val.DefinedNames;
			val.toXml(writer);
		}
		if (this.val.SheetDataSet) {
			val = new CT_ExternalSheetDataSet();
			val.val = this.val.SheetDataSet;
			val.toXml(writer);
		}

		writer.WriteXmlString("</externalBook>");
	};

	function CT_ExternalSheetNames() {
		this.val = null;
	}

	CT_ExternalSheetNames.prototype.fromXml = function (reader) {
		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("sheetName" === name) {
				while (reader.MoveToNextAttribute()) {
					if ("val" === reader.GetNameNoNS()) {
						if (!this.val) {
							this.val = [];
						}
						this.val.push(reader.GetValue());
					}
				}
			}
		}
	};
	CT_ExternalSheetNames.prototype.toXml = function (writer) {

		writer.WriteXmlString("<sheetNames>");
		for (var i = 0; i < this.val.length; ++i) {
			if (this.val[i]) {
				writer.WriteXmlString("<sheetName");
				writer.WriteXmlAttributeStringEncode("val", this.val[i]);
				writer.WriteXmlString("/>");
			}
		}
		writer.WriteXmlString("</sheetNames>");
	};

	function CT_ExternalDefinedNames() {
		this.val = null;
	}

	CT_ExternalDefinedNames.prototype.fromXml = function (reader) {
		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		var val;
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("definedName" === name) {
				var definedName = {Name: null, RefersTo: null, SheetId: null};
				while (reader.MoveToNextAttribute()) {
					if ("name" === reader.GetName()) {
						val = reader.GetValue();
						definedName.Name = val;
					} else if ("refersTo" === reader.GetName()) {
						val = reader.GetValue();
						definedName.RefersTo = val;
					} else if ("sheetId" === reader.GetName()) {
						val = reader.GetValue();
						definedName.SheetId = val;
					}
				}
				if (!this.val) {
					this.val = [];
				}
				this.val.push(definedName);
			}
		}
	};

	CT_ExternalDefinedNames.prototype.toXml = function (writer) {
		if (!this.val || !this.val.length) {
			return;
		}

		writer.WriteXmlString("<definedNames>");
		for (var i = 0; i < this.val.length; ++i) {
			if (this.val[i]) {
				writer.WriteXmlString("<sheetName>");

				/*writer.WriteString(L"<definedName");
				WritingStringNullableAttrEncodeXmlString(L"name", m_oName, m_oName.get());
				WritingStringNullableAttrEncodeXmlString(L"refersTo", m_oRefersTo, m_oRefersTo.get());
				WritingStringNullableAttrInt(L"sheetId", m_oSheetId, m_oSheetId->GetValue());
				writer.WriteString(L"/>");*/

				writer.WriteXmlString("<definedName");
				writer.WriteXmlNullableAttributeStringEncode("name", this.val[i].Name);
				writer.WriteXmlNullableAttributeStringEncode("refersTo", this.val[i].RefersTo);
				writer.WriteXmlNullableAttributeNumber("sheetId", this.val[i].SheetId);
				writer.WriteXmlString("/>");

				writer.WriteXmlString("/>");
			}
		}
		writer.WriteXmlString("</definedNames>");

	};

	function CT_ExternalSheetDataSet() {
		this.val = null;
	}

	CT_ExternalSheetDataSet.prototype.fromXml = function (reader) {
		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		var val;
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("sheetData" === name) {
				val = new CT_ExternalSheetData();
				val.fromXml(reader);
				if (val.val) {
					if (!this.val) {
						this.val = [];
					}

					this.val.push(val.val);
				}
			}
		}
	}

	CT_ExternalSheetDataSet.prototype.toXml = function (writer) {
		if (!this.val || !this.val.length) {
			return;
		}

		writer.WriteXmlString("<sheetDataSet>");
		for (var i = 0; i < this.val.length; ++i) {
			if (this.val[i]) {
				var val = new CT_ExternalSheetData();
				val.val = this.val[i];
				val.toXml(writer);
			}
		}
		writer.WriteXmlString("</sheetDataSet>");
	};

	function CT_ExternalSheetData() {
		this.val = {SheetId: null, RefreshError: null, Row: []};
	}

	CT_ExternalSheetData.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			if ("row" === reader.GetName()) {
				var val = new CT_ExternalRow();
				val.fromXml(reader);
				if (val.val) {
					this.val.Row.push(val.val);
				}
			}
		}
	}

	CT_ExternalSheetData.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {

			if ("refersTo" === reader.GetName()) {
				val = reader.GetValue();
				this.val.RefersTo = val;
			} else if ("sheetId" === reader.GetName()) {
				val = reader.GetValue();
				this.val.SheetId = val;
			}
		}
	};

	CT_ExternalSheetData.prototype.toXml = function (writer) {
		if (!this.val) {
			return;
		}
		writer.WriteXmlString("<sheetData");
		writer.WriteXmlNullableAttributeNumber("sheetId", this.val.SheetId);
		writer.WriteXmlNullableAttributeBool("refreshError", this.val.RefreshError);
		writer.WriteXmlString(">");

		if (this.val.Row) {
			for (var i = 0; i < this.val.Row.length; ++i) {
				if (this.val.Row[i]) {
					var val = new CT_ExternalRow();
					val.val = this.val.Row[i];
					val.toXml(writer);
				}
			}
		}

		writer.WriteXmlString("</sheetData>");
	};


	function CT_ExternalRow() {
		this.val = {R: null, Cell: []};
	}

	CT_ExternalRow.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			if ("cell" === reader.GetName()) {
				var externalCell = new CT_ExternalCell();
				externalCell.fromXml(reader);
				if (externalCell.val) {
					this.val.Cell.push(externalCell.val);
				}
			}
		}
	}
	CT_ExternalRow.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("r" === reader.GetName()) {
				val = reader.GetValue();
				this.val.R = val;
			}
		}
	};

	CT_ExternalRow.prototype.toXml = function (writer) {
		writer.WriteXmlString("<row");
		writer.WriteXmlNullableAttributeNumber("r", this.val.R);
		writer.WriteXmlString(">");

		if (this.val.Cell) {
			for (var i = 0; i < this.val.Cell.length; ++i) {
				if (this.val.Cell[i]) {
					var val = new CT_ExternalCell();
					val.val = this.val.Cell[i];
					val.toXml(writer);
				}
			}
		}

		writer.WriteXmlString("</row>");
	};

	function CT_ExternalCell() {
		this.val = {Ref: null, CellType: null, CellValue: null};
	}

	CT_ExternalCell.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		var depth = reader.GetDepth();
		var val;
		while (reader.ReadNextSiblingNode(depth)) {
			if ("v" === reader.GetName()) {
				val = prepareTextFromXml(reader.GetTextDecodeXml());
				this.val.CellValue = val;
			}
		}
	}
	CT_ExternalCell.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {

			if ("r" === reader.GetName()) {
				val = reader.GetValue();
				this.val.Ref = val;
			} else if ("t" === reader.GetName()) {
				val = reader.GetValue();
				this.val.CellType = FromXml_ST_CellValueType(val);
			} else if ("vm" === reader.GetName()) {
				/*val = reader.GetValue();
				this.val.CellType = val;*/
			}
		}
	};

	CT_ExternalCell.prototype.toXml = function (writer) {
		writer.WriteXmlString("<cell");
		writer.WriteXmlNullableAttributeString("r", this.val.Ref);
		writer.WriteXmlNullableAttributeString("t", ToXml_ST_CellValueType(this.val.CellType));
		//writer.WriteXmlNullableAttributeNumber("vm", this.vm);
		writer.WriteXmlString(">");

		if (this.val.CellValue) {
			writer.WriteXmlString("<v");
			writer.WriteXmlString(">");
			writer.WriteXmlString(this.val.CellValue);
			writer.WriteXmlString("</v>");
		}

		writer.WriteXmlString("</cell>");
	};


	//****Slicer cache****
	Asc.CT_slicerCacheDefinition.prototype.fromXml = function (reader) {
		if (!reader.ReadNextNode()) {
			return;
		}

		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		var val;
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("pivotTables" === name) {
				var depth2 = reader.GetDepth();
				while (reader.ReadNextSiblingNode(depth2)) {
					var name2 = reader.GetNameNoNS();
					if ("pivotTable" === name2) {
						val = new Asc.CT_slicerCachePivotTable();
						val.fromXml(reader);
						this.pivotTables.push(val);
					}
				}
			} else if ("data" === name) {
				val = new Asc.CT_slicerCacheData();
				val.fromXml(reader);
				this.data = val;
			} else if ("extLst" === name) {
				var extLst = new COfficeArtExtensionList(this);
				extLst.fromXml(reader);
				this.extLst = extLst.arrExt;
			}
		}

		//TODO tableSlicerCache from extLst
		//CT_tableSlicerCache.prototype.initPostOpen
		this.parseExtLst();
	};

	Asc.CT_slicerCacheDefinition.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.name = val;
			} /*else if ("uid" === reader.GetNameNoNS()) {
				val = reader.GetValue();
				this.uid = val;
			}*/ else if ("sourceName" === reader.GetName()) {
				val = reader.GetValue();
				this.sourceName = val;
			}
		}
	};

	Asc.CT_slicerCacheDefinition.prototype.toXml = function (writer, sName) {
		writer.WriteXmlString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");

		if (!sName) {
			sName = "slicerCacheDefinition";
		}

		writer.WriteXmlNodeStart(sName);
		writer.WriteXmlString(" xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x xr10\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\"");
		if (this.name) {
			writer.WriteXmlAttributeStringEncode("name", this.name);
		}
		if (this.uid) {
			writer.WriteXmlAttributeStringEncode("xr10:uid",this.uid);
		}
		if (this.sourceName) {
			writer.WriteXmlAttributeStringEncode("sourceName", this.sourceName);
		}
		writer.WriteXmlAttributesEnd();

		//SlicerCachePivotTable
		if (this.pivotTables.length > 0) {
			writer.WriteXmlNodeStart("pivotTables");
			writer.WriteXmlAttributesEnd();
			for (var i = 0; i < this.pivotTables.length; ++i) {
				this.pivotTables[i].toXml(writer, "pivotTable");
			}
			writer.WriteXmlNodeEnd("pivotTables");
		}
		//CSlicerCacheData
		if (this.data) {
			this.data.toXml(writer, "data");
		}

		if (this.tableSlicerCache) {
			var officeArtExtensionList = new COfficeArtExtensionList();
			var officeArtExtension = new COfficeArtExtension(this);
			officeArtExtension.uri = extUri.tableSlicerCache;
			officeArtExtension.additionalNamespace = additionalNamespace.tableSlicerCache;
			officeArtExtension.tableSlicerCache = this.tableSlicerCache;

			officeArtExtensionList.arrExt.push(officeArtExtension);
			officeArtExtensionList.toXml(writer);
		}
		//TODO EXTLst
		//writer.WriteXmlString(m_oExtLst.toXMLWithNS(""))

		writer.WriteXmlNodeEnd(sName);
	};

	Asc.CT_slicerCacheDefinition.prototype.parseExtLst = function () {
		if (this.extLst) {
			for (var i = 0; i < this.extLst.length; i++) {
				var tableSlicerCache = this.extLst[i].tableSlicerCache;
				if (tableSlicerCache) {
					this.tableSlicerCache = tableSlicerCache;
				}
			}
			this.extLst = null;
		}
	};

	Asc.CT_slicerCachePivotTable.prototype.fromXml = function (reader) {

		/*ReadAttributes(oReader);
			if (oReader.IsEmptyNode())
				return;
			oReader.ReadTillEnd();*/

		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	Asc.CT_slicerCachePivotTable.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("tabId" === reader.GetName()) {
				val = reader.GetValue();
				this.tabIdOpen = val;
			} else if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.name = val;
			}
		}
	};

	Asc.CT_slicerCachePivotTable.prototype.toXml = function (writer, sName) {
		if (!sName) {
			sName = "pivotTable";
		}

		writer.WriteXmlNodeStart(sName);

		var sheetIds = writer.context.InitSaveManager.getSheetIds();
		var tabId = sheetIds[this.sheetId];

		if (tabId != null) {
			writer.WriteXmlAttributeString("tabId", tabId);
		}
		if (this.name) {
			writer.WriteXmlAttributeStringEncode("name", this.name);
		}

		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNodeEnd(sName);
	};

	Asc.CT_slicerCacheData.prototype.fromXml = function (reader) {
		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		var val;
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			/*this.olap = null;//OlapSlicerCache
			this.tabular = null;//TabularSlicerCache*/

			if ("olap" === name) {
				val = new Asc.CT_olapSlicerCache();
				val.fromXml(reader);
				this.olap = val;
			} else if ("tabular" === name) {
				val = new Asc.CT_tabularSlicerCache();
				val.fromXml(reader);
				this.tabular = val;
			}
		}
	};

	Asc.CT_slicerCacheData.prototype.toXml = function (writer, sName) {

		if (!sName) {
			sName = "data";
		}

		writer.WriteXmlNodeStart(sName);
		writer.WriteXmlAttributesEnd();

		//OlapSlicerCache
		if (this.olap) {
			this.olap.toXml(writer, "olap");
		}
		//TabularSlicerCache
		if (this.tabular) {
			this.tabular.toXml(writer, "tabular");
		}

		writer.WriteXmlNodeEnd(sName);
	};

	Asc.CT_olapSlicerCache.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var depth2, name2, val;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			/*this.levels = [];//OlapSlicerCacheLevelData
			this.selections = [];//OlapSlicerCacheSelection
			this.pivotCacheId = null;
			this.pivotCacheDefinition = null;*/

			if ("levels" === name) {
				depth2 = reader.GetDepth();
				while (reader.ReadNextSiblingNode(depth2)) {
					name2 = reader.GetNameNoNS();
					if ("level" === name2) {
						val = Asc.CT_olapSlicerCacheLevelData();
						val.fromXml(reader);
						this.levels.push(val);
					}
				}
			} else if ("selections" === name) {
				depth2 = reader.GetDepth();
				while (reader.ReadNextSiblingNode(depth2)) {
					name2 = reader.GetNameNoNS();
					if ("level" === name2) {
						val = Asc.CT_olapSlicerCacheSelection();
						val.fromXml(reader);
						this.selections.push(val);
					}
				}
			}
		}
	};

	Asc.CT_olapSlicerCache.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("pivotCacheId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.pivotCacheId = val;
			}
		}
	};

	Asc.CT_olapSlicerCache.prototype.toXml = function (writer, sName) {
		if (!sName) {
			sName = "olap";
		}

		writer.WriteXmlNodeStart(sName);

		if (this.pivotCacheId) {
			writer.WriteXmlAttributeString("pivotCacheId", this.pivotCacheId);
		}
		writer.WriteXmlAttributesEnd();

		if (this.data) {
			this.data.toXml(writer, "data");
		}
		if (this.data) {
			this.data.toXml(writer, "data");
		}

		//OlapSlicerCacheLevelsData
		if (this.levels.length > 0) {
			writer.WriteXmlArray(this.levels, "level", "levels", true);
		}
		if (this.selections.length > 0) {
			writer.WriteXmlArray(this.selections, "selection", "selections", true);
		}

		//extLst
		//WritingNullable(m_oExtLst, writer.WriteString(m_oExtLst->toXMLWithNS(L"")););

		writer.WriteXmlNodeEnd(sName);
	};

	Asc.CT_olapSlicerCacheLevelData.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("ranges" === name) {
				var depth2 = reader.GetDepth();
				while (reader.ReadNextSiblingNode(depth2)) {
					var name2 = reader.GetNameNoNS();
					if ("range" === name2) {
						var val = Asc.CT_olapSlicerCacheRange();
						val.fromXml(reader);
						this.ranges.push(val);
					}
				}
			}
		}
	};

	Asc.CT_olapSlicerCacheLevelData.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("uniqueName" === reader.GetName()) {
				val = reader.GetValue();
				this.uniqueName = val;
			} else if ("sourceCaption" === reader.GetName()) {
				val = reader.GetValue();
				this.sourceCaption = val;
			} else if ("count" === reader.GetName()) {
				val = reader.GetValue();
				this.count = val;
			} else if ("sortOrder" === reader.GetName()) {
				val = reader.GetValue();
				this.sortOrder = FromXml_ST_TabularSlicerCacheSortOrder(val);
			} else if ("crossFilter" === reader.GetName()) {
				val = reader.GetValue();
				this.crossFilter = FromXml_ST_SlicerCacheCrossFilter(val);
			}
		}
	};

	Asc.CT_olapSlicerCacheLevelData.prototype.toXml = function (writer, sName) {
		if (!sName) {
			sName = "level";
		}
		writer.WriteXmlNodeStart(sName);

		if (this.uniqueName) {
			writer.WriteXmlAttributeStringEncode("uniqueName", this.uniqueName);
		}
		if (this.sourceCaption) {
			writer.WriteXmlAttributeStringEncode("sourceCaption", this.sourceCaption);
		}
		if (this.count) {
			writer.WriteXmlAttributeString("sourceCaption", this.count);
		}
		if (this.sortOrder) {
			writer.WriteXmlAttributeString("sortOrder", ToXml_ST_TabularSlicerCacheSortOrder(this.sortOrder));
		}
		if (this.crossFilter) {
			writer.WriteXmlAttributeString("crossFilter", ToXml_ST_SlicerCacheCrossFilter(this.crossFilter));
		}

		writer.WriteXmlAttributesEnd();

		//CT_olapSlicerCacheRange
		if (this.ranges.length > 0) {
			writer.WriteXmlArray(this.ranges, "range", "ranges");
		}

		writer.WriteXmlNodeEnd(sName);
	};

	Asc.CT_olapSlicerCacheRange.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			/*this.i = [];//OlapSlicerCacheItem
			this.startItem = null*/

			if ("i" === name) {
				var val = Asc.CT_olapSlicerCacheItem();
				val.fromXml(reader);
				this.i.push(val);
			}
		}
	};

	Asc.CT_olapSlicerCacheRange.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("startItem" === reader.GetName()) {
				val = reader.GetValue();
				this.startItem = val;
			}
		}
	};

	Asc.CT_olapSlicerCacheRange.prototype.toXml = function (writer, sName) {
		if (!sName) {
			sName = "range";
		}

		writer.WriteXmlNodeStart(sName);
		if (this.startItem != null) {
			writer.WriteXmlAttributeString("startItem", this.startItem);
		}
		writer.WriteXmlAttributesEnd();

		if (this.i.length > 0) {
			writer.WriteXmlArray(this.i, "i");
		}

		writer.WriteXmlNodeEnd(sName);
	};

	Asc.CT_olapSlicerCacheItem.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			/*this.p = [];//OlapSlicerCacheItemParent - состоит из одного поля, поэтому данную структуру не добавляю
			this.n = null;
			this.c = null;
			this.nd = false;*/

			if ("p" === name) {
				while (reader.MoveToNextAttribute()) {
					if ("n" === reader.GetName()) {
						this.p.push(reader.GetValue());
					}
				}
			}
		}
	};

	Asc.CT_olapSlicerCacheItem.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("n" === reader.GetName()) {
				val = reader.GetValue();
				this.n = val;
			} else if ("c" === reader.GetName()) {
				val = reader.GetValue();
				this.c = val;
			} else if ("nd" === reader.GetName()) {
				val = reader.GetValue();
				this.nd = val;
			}
		}
	};

	Asc.CT_olapSlicerCacheItem.prototype.toXml = function (writer, sName) {

		if (!sName) {
			sName = "i";
		}

		writer.WriteXmlNodeStart(sName);
		if (this.n != null) {
			writer.WriteXmlAttributeStringEncode("n", this.n);
		}
		if (this.c != null) {
			writer.WriteXmlAttributeStringEncode("c", this.c);
		}
		if (this.nd != null) {
			writer.WriteXmlAttributeString("nd", this.nd);
		}
		writer.WriteXmlAttributesEnd();

		if (this.p.length > 0) {
			for (var i = 0; i < this.p.length; i++) {
				writer.WriteXmlNodeStart("p");
				if (this.c != null) {
					writer.WriteXmlAttributeStringEncode("n", this.n);
				}
				writer.WriteXmlAttributesEnd();
				writer.WriteXmlNodeEnd("p");
			}
		}

		writer.WriteXmlNodeEnd(sName);
	};

	//CT_olapSlicerCacheSelection
	Asc.CT_olapSlicerCacheSelection.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			/*this.p = [];//OlapSlicerCacheItemParent - состоит из одного поля, поэтому данную структуру не добавляю
		this.n = null;*/

			if ("p" === name) {
				while (reader.MoveToNextAttribute()) {
					if ("n" === reader.GetNameNoNS()) {
						this.p.push(reader.getValue());
					}
				}
			}
		}
	};

	Asc.CT_olapSlicerCacheSelection.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("n" === reader.GetName()) {
				val = reader.GetValue();
				this.n = val;
			}
		}
	};


	Asc.CT_olapSlicerCacheSelection.prototype.toXml = function (writer, sName) {
		if (!sName) {
			sName = "selection";
		}

		writer.WriteXmlNodeStart(sName);
		if (this.n != null) {
			writer.WriteXmlAttributeStringEncode("n", this.n);
		}
		writer.WriteXmlAttributesEnd();

		if (this.p.length > 0) {
			for (var i = 0; i < this.p.length; i++) {
				writer.WriteXmlNodeStart("p");
				if (this.c != null) {
					writer.WriteXmlAttributeStringEncode("n", this.n);
				}
				writer.WriteXmlAttributesEnd();
				writer.WriteXmlNodeEnd("p");
			}
		}

		writer.WriteXmlNodeEnd(sName);
	};

	Asc.CT_tabularSlicerCache.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("items" === name) {
				var depth2 = reader.GetDepth();
				while (reader.ReadNextSiblingNode(depth2)) {
					var name2 = reader.GetNameNoNS();
					if ("i" === name2) {
						//CTabularSlicerCacheItem
						var val = new Asc.CT_tabularSlicerCacheItem();
						val.fromXml(reader);
						this.items.push(val);
					}
				}
			} else if ("extLst" === name) {
			}
		}
	};

	Asc.CT_tabularSlicerCache.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("pivotCacheId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.pivotCacheId = val;
			} else if ("sortOrder" === reader.GetName()) {
				val = reader.GetValue();
				this.sortOrder = FromXml_ST_TabularSlicerCacheSortOrder(val);
			} else if ("customListSort" === reader.GetName()) {
				val = reader.GetValue();
				this.customListSort = val;
			} else if ("showMissing" === reader.GetName()) {
				val = reader.GetValue();
				this.showMissing = val;
			} else if ("crossFilter" === reader.GetName()) {
				val = reader.GetValue();
				this.crossFilter = FromXml_ST_SlicerCacheCrossFilter(val);
			}
		}
	};

	Asc.CT_tabularSlicerCache.prototype.toXml = function (writer, sName) {

		if (!sName) {
			sName = "tabular";
		}

		writer.WriteXmlNodeStart(sName);
		var pivotCacheId = this.pivotCacheId;
		if (!pivotCacheId && this.pivotCacheDefinition) {
			pivotCacheId = this.pivotCacheDefinition.getPivotCacheId()
		}
		if (pivotCacheId != null) {
			writer.WriteXmlAttributeString("pivotCacheId", pivotCacheId);
		}
		if (this.sortOrder != null) {
			writer.WriteXmlAttributeString("sortOrder", ToXml_ST_TabularSlicerCacheSortOrder(this.sortOrder));
		}
		if (this.customListSort != null) {
			writer.WriteXmlAttributeBool("customListSort", this.customListSort);
		}
		if (this.showMissing != null) {
			writer.WriteXmlAttributeBool("showMissing", this.showMissing);
		}
		if (this.crossFilter != null) {
			writer.WriteXmlAttributeString("crossFilter", ToXml_ST_SlicerCacheCrossFilter(this.crossFilter));
		}
		writer.WriteXmlAttributesEnd();

		if (this.items.length > 0) {
			writer.WriteXmlArray(this.items, "i", "items", true);
		}

		writer.WriteXmlNodeEnd(sName);
	};

	Asc.CT_tabularSlicerCacheItem.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	Asc.CT_tabularSlicerCacheItem.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("x" === reader.GetName()) {
				val = reader.GetValueInt();
				this.x = val;
			} else if ("s" === reader.GetName()) {
				val = reader.GetValueBool();
				this.s = val;
			} else if ("nd" === reader.GetName()) {
				val = reader.GetValueBool();
				this.nd = val;
			}
		}
	};

	Asc.CT_tabularSlicerCacheItem.prototype.toXml = function (writer, sName) {

		if (!sName) {
			sName = "i";
		}

		writer.WriteXmlNodeStart(sName);
		if (this.x != null) {
			writer.WriteXmlAttributeString("x", this.x);
		}
		if (this.s != null) {
			writer.WriteXmlAttributeBool("s", this.s);
		}
		if (this.nd != null) {
			writer.WriteXmlAttributeBool("nd", this.nd);
		}

		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNodeEnd(sName);
	};

	Asc.CT_tableSlicerCache.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("extLst" === name) {
			}
		}
	};

	Asc.CT_tableSlicerCache.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("tableId" === reader.GetName()) {
				val = reader.GetValueInt();
				this.tableIdOpen = val;
			} else if ("column" === reader.GetName()) {
				val = reader.GetValueInt();
				this.columnOpen = val;
			} else if ("sortOrder" === reader.GetName()) {
				val = reader.GetValue();
				this.sortOrder = FromXml_ST_TabularSlicerCacheSortOrder(val);
			} else if ("customListSort" === reader.GetName()) {
				val = reader.GetValue();
				this.customListSort = val;
			} else if ("crossFilter" === reader.GetName()) {
				val = reader.GetValue();
				this.crossFilter = FromXml_ST_SlicerCacheCrossFilter(val);
			}
		}
	};

	Asc.CT_tableSlicerCache.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		var tableIds = writer.context.InitSaveManager.getTableIds();
		var tableIdOpen = null;
		var columnOpen = null;
		var elem = tableIds && tableIds[this.tableId];
		if (elem) {
			tableIdOpen = elem.id;
			columnOpen = (elem.table.getTableIndexColumnByName(this.column) + 1) || null;
		}

		writer.WriteXmlNodeStart(ns + name);
		writer.WriteXmlNullableAttributeString("tableId", tableIdOpen);
		writer.WriteXmlNullableAttributeString("column", columnOpen);
		writer.WriteXmlNullableAttributeString("sortOrder", ToXml_ST_TabularSlicerCacheSortOrder(this.sortOrder));
		writer.WriteXmlNullableAttributeString("customListSort", this.customListSort);
		writer.WriteXmlNullableAttributeString("crossFilter", ToXml_ST_SlicerCacheCrossFilter(this.crossFilter));
		writer.WriteXmlAttributesEnd();
		//WritingNullable(m_oExtLst, writer.WriteXmlString(m_oExtLst.toXMLWithNS("")););
		writer.WriteXmlNodeEnd(ns + name);
	};

	Asc.CT_slicers.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var t = this;
		reader.readXmlArray("slicers", function () {
			if (reader.IsEmptyNode()) {
				return;
			}
			var depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				var name = reader.GetNameNoNS();

				if ("slicer" === name) {
					var slicer = new Asc.CT_slicer(t._ws);
					slicer.fromXml(reader);
					if (!t.slicer) {
						t.slicer = [];
					}
					t.slicer.push(slicer);
				}
			}
		});
	};

	Asc.CT_slicers.prototype.readAttr = function (reader) {

//documentation
		/*<xsd:complexType name="CT_Slicers">
		<xsd:sequence>
		<xsd:element name="slicer" type="CT_Slicer" minOccurs="1" maxOccurs="unbounded"/>
		</xsd:sequence>
		</xsd:complexType>*/

//x2t
		/*WritingElement_ReadAttributes_StartChar_No_NS(oReader)
					WritingElement_ReadAttributes_EndChar_No_NS( oReader )*/

//serialize
		/*var res = c_oSerConstants.ReadOk;
					if(c_oSer_QueryTableDeletedField.Name == type)
					{
						pQueryTableDeletedField.name = this.stream.GetString2LE(length);
					}*/
	};

	Asc.CT_slicers.prototype.toXml = function (writer, name) {
		if (!name) {
			name = "slicers";
		}

		writer.WriteXmlString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlString(
			" xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x xr10\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\"");
		writer.WriteXmlAttributesEnd();

		if (this.slicer.length > 0) {
			for (var i = 0; i < this.slicer.length; ++i) {
				this.slicer[i].toXml(writer, "slicer");
			}
		}

		writer.WriteXmlNodeEnd(name);
	};

	Asc.CT_slicer.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			//extLst
		}
	};

	Asc.CT_slicer.prototype.readAttr = function (reader) {
		var slicerCaches = reader.context.InitOpenManager.oReadResult.slicerCaches;
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.name = val;
			} else if ("uid" === reader.GetName()) {
				val = reader.GetValue();
				this.uid = val;
			} else if ("cache" === reader.GetName()) {
				val = reader.GetValue();
				this.cacheDefinition = slicerCaches[val] || null;
			} else if ("caption" === reader.GetName()) {
				val = reader.GetValue();
				this.caption = val;
			} else if ("startItem" === reader.GetName()) {
				val = reader.GetValueInt();
				this.startItem = val;
			} else if ("columnCount" === reader.GetName()) {
				val = reader.GetValueInt();
				this.columnCount = val;
			} else if ("showCaption" === reader.GetName()) {
				val = reader.GetValueBool();
				this.showCaption = val;
			} else if ("level" === reader.GetName()) {
				val = reader.GetValueInt();
				this.level = val;
			} else if ("style" === reader.GetName()) {
				val = reader.GetValue();
				this.style = val;
			} else if ("lockedPosition" === reader.GetName()) {
				val = reader.GetValueBool();
				this.lockedPosition = val;
			} else if ("rowHeight" === reader.GetName()) {
				//TODO метрика
				val = reader.GetValueInt();
				this.rowHeight = val;
			}
		}
	};

	Asc.CT_slicer.prototype.toXml = function (writer, name) {
		if (!name) {
			name = "slicer";
		}

		writer.WriteXmlNodeStart(name);
		if (null != this.name) {
			writer.WriteXmlAttributeStringEncode("name", this.name);
		}
		if (null != this.uid) {
			writer.WriteXmlAttributeStringEncode("xr10:uid", this.uid);
		}
		if (this.cacheDefinition && null != this.cacheDefinition.name) {
			writer.WriteXmlAttributeStringEncode("cache", this.cacheDefinition.name);
		}
		if (null != this.caption) {
			writer.WriteXmlAttributeStringEncode("caption", this.caption);
		}
		if (null != this.startItem) {
			writer.WriteXmlAttributeInt("startItem", this.startItem);
		}
		if (null != this.columnCount) {
			writer.WriteXmlAttributeInt("columnCount", this.columnCount);
		}
		if (null != this.showCaption) {
			writer.WriteXmlAttributeBool("showCaption", this.showCaption);
		}
		if (null != this.level) {
			writer.WriteXmlAttributeInt("level", this.level);
		}
		if (null != this.style) {
			writer.WriteXmlAttributeString("style", this.style);
		}
		if (null != this.lockedPosition) {
			writer.WriteXmlAttributeBool("lockedPosition", this.lockedPosition);
		}
		if (null != this.rowHeight) {
			writer.WriteXmlAttributeInt("rowHeight", this.rowHeight);
		}

		writer.WriteXmlAttributesEnd();
		//WritingNullable(m_oExtLst, writer.WriteXmlString(m_oExtLst.toXMLWithNS("")););
		writer.WriteXmlNodeEnd(name);
	};


	//пока читаю в строку connections. в serialize сейчас аналогично не парсим структуру, а храним в виде массива байтов
	//connections
	function CT_Connections() {
		//Members
		this.val = [];
	}

	CT_Connections.prototype.fromXml = function (reader) {
		//count
		//this.readAttr(reader);

		if (!reader.ReadNextNode()) {
			return;
		}
		if (reader.IsEmptyNode()) {
			return;
		}

		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("connection" === name) {
				var connection = new AscCommonExcel.CT_Connection();
				connection.fromXml(reader);
				this.val.push(connection);
			}
		}
	};

	AscCommonExcel.CT_Connection.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("dbPr" === name) {

			} else if ("olapPr" === name) {

			} else if ("textPr" === name) {

			} else if ("webPr" === name) {

			} else if ("rangePr" === name) {

			} else if ("extLst" === name) {

			}
		}
	};

	AscCommonExcel.CT_Connection.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("type" === reader.GetName()) {
				val = reader.GetValueInt();
				this.type = val;
			} else if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.name = val;
			} else if ("id" === reader.GetName()) {
				val = reader.GetValueInt();
				this.id = val;
			} else if ("background" === reader.GetName()) {
				val = reader.GetValueBool();
				this.background = val;
			} else if ("credentials" === reader.GetName()) {
				val = reader.GetValue();
				this.credentials = val;
			} else if ("deleted" === reader.GetName()) {
				val = reader.GetValueBool();
				this.deleted = val;
			} else if ("description" === reader.GetName()) {
				val = reader.GetValue();
				this.description = val;
			} else if ("interval" === reader.GetName()) {
				val = reader.GetValueInt();
				this.interval = val;
			} else if ("keepAlive" === reader.GetName()) {
				val = reader.GetValueBool();
				this.keepAlive = val;
			} else if ("minRefreshableVersion" === reader.GetName()) {
				val = reader.GetValue();
				this.minRefreshableVersion = val;
			} else if ("new" === reader.GetName()) {
				val = reader.GetValueBool();
				this.new = val;
			} else if ("odcFile" === reader.GetName()) {
				val = reader.GetValue();
				this.odcFile = val;
			} else if ("onlyUseConnectionFile" === reader.GetName()) {
				val = reader.GetValueBool();
				this.onlyUseConnectionFile = val;
			} else if ("reconnectionMethod" === reader.GetName()) {
				val = reader.GetValueInt();
				this.reconnectionMethod = val;
			} else if ("refreshedVersion" === reader.GetName()) {
				val = reader.GetValue();
				this.refreshedVersion = val;
			} else if ("refreshOnLoad" === reader.GetName()) {
				val = reader.GetValueBool();
				this.refreshOnLoad = val;
			} else if ("saveData" === reader.GetName()) {
				val = reader.GetValueBool();
				this.saveData = val;
			} else if ("savePassword" === reader.GetName()) {
				val = reader.GetValueBool();
				this.savePassword = val;
			} else if ("singleSignOnId" === reader.GetName()) {
				val = reader.GetValue();
				this.singleSignOnId = val;
			} else if ("sourceFile" === reader.GetName()) {
				val = reader.GetValue();
				this.sourceFile = val;
			}
		}
	};


	Asc.CT_NamedSheetViews.prototype.fromXml = function (reader) {
		var t = this;
		reader.readXmlArray("namedSheetViews", function () {
			if (reader.IsEmptyNode()) {
				return;
			}
			var depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				var name = reader.GetNameNoNS();

				if ("namedSheetView" === name) {
					var namedSheetView = new Asc.CT_NamedSheetView();
					namedSheetView.fromXml(reader);
					t.namedSheetView.push(namedSheetView);
				} else if ("extLst" === name) {
				}
			}
		});

	};

	Asc.CT_NamedSheetViews.prototype.toXml = function (writer, name) {
		if (!this.namedSheetView || !this.namedSheetView.length) {
			return;
		}

		writer.WriteXmlString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');
		writer.WriteXmlNodeStart("namedSheetViews");
		writer.WriteXmlString(
			" xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
		writer.WriteXmlAttributesEnd();
		for (var i = 0; i < this.namedSheetView.length; ++i) {
			this.namedSheetView[i].toXml(writer, "namedSheetView");
		}
		//WritingNullable(m_oExtLst, writer.WriteXmlString(m_oExtLst.toXMLWithNS("")););
		writer.WriteXmlNodeEnd("namedSheetViews");
	};

	Asc.CT_NamedSheetView.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			var val;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("nsvFilter" === name) {
				var nsvFilter = new Asc.CT_NsvFilter();
				nsvFilter.fromXml(reader);
				this.nsvFilters.push(nsvFilter);
			} else if ("extLst" === name) {
			}
		}
	};

	Asc.CT_NamedSheetView.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("name" === reader.GetName()) {
				val = reader.GetValue();
				this.name = val;
			} else if ("id" === reader.GetName()) {
				val = reader.GetValue();
				this.id = val;
			}
		}
	};

	Asc.CT_NamedSheetView.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributeStringEncode("name", this.name);
		writer.WriteXmlAttributeStringEncode("id", this.id);
		writer.WriteXmlAttributesEnd();
		for (var i = 0; i < this.nsvFilters.length; ++i) {
			this.nsvFilters[i].toXml(writer, "nsvFilter");
		}
		//WritingNullable(m_oExtLst, writer.WriteXmlString(m_oExtLst.toXMLWithNS("")););
		writer.WriteXmlNodeEnd(name);
	};

	Asc.CT_NsvFilter.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("columnFilter" === name) {
				var columnFilter = new Asc.CT_ColumnFilter();
				columnFilter.fromXml(reader);
				this.columnsFilter.push(columnFilter);
			}
			if ("sortRules" === name) {
				var sortRules = new Asc.CT_SortRules();
				sortRules.fromXml(reader);
				this.sortRules = sortRules.sortRule;
			} else if ("extLst" === name) {
			}
		}
	};

	Asc.CT_NsvFilter.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("filterId" === reader.GetName()) {
				val = reader.GetValue();
				this.filterId = val;
			} else if ("ref" === reader.GetName()) {
				val = reader.GetValue();
				this.ref = val;
			} else if ("tableId" === reader.GetName()) {
				val = reader.GetValue();
				//this.tableId = val;
				this.tableIdOpen = val;
			}
		}
	};

	Asc.CT_NsvFilter.prototype.toXml = function (writer, name) {
		var tableIds = writer.context.InitSaveManager.getTableIds();
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributeStringEncode("filterId", this.filterId);
		writer.WriteXmlNullableAttributeStringEncode("ref", this.ref);
		if ("0" === this.tableId) {
			writer.WriteXmlAttributeString("tableId", "0");
		} else {
			var elem = tableIds && tableIds[this.tableId];
			if (elem) {
				writer.WriteXmlAttributeString("tableId", elem.id + "");
			}
		}
		writer.WriteXmlAttributesEnd();
		for (var i = 0; i < this.columnsFilter.length; ++i) {
			this.columnsFilter[i].toXml(writer, "columnFilter");
		}
		if (this.sortRules && this.sortRules.length) {
			var sortRules = new Asc.CT_SortRules();
			sortRules.sortRule = this.sortRules;
			sortRules.toXml(writer, "sortRules");
		}

		//WritingNullable(m_oExtLst, writer.WriteXmlString(m_oExtLst.toXMLWithNS("")););
		writer.WriteXmlNodeEnd(name);
	};

	Asc.CT_ColumnFilter.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		var val;
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("dxf" === name) {
				val = new AscCommonExcel.CellXfs();
				val.fromXml(reader);
				this.dxf = val;
			} else if ("filter" === name) {
				val = new window['AscCommonExcel'].FilterColumn();
				val.fromXml(reader);
				if (!this.filter) {
					this.filter = [];
				}
				this.filter = val;
			} else if ("extLst" === name) {
			}
		}
	};

	Asc.CT_ColumnFilter.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("colId" === reader.GetName()) {
				val = reader.GetValue();
				this.colId = val;
			} else if ("id" === reader.GetName()) {
				val = reader.GetValue();
				this.id = val;
			}
		}
	};

	Asc.CT_ColumnFilter.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name/*"columnFilter"*/);
		//WritingNullable(m_oDefaultSlicerStyle, writer.WriteAttributeEncodeXml("defaultSlicerStyle", *m_oDefaultSlicerStyle););
		//colId WriteAttribute
		writer.WriteXmlAttributeString("colId", this.colId);
		writer.WriteXmlAttributeStringEncode("id", this.id);
		writer.WriteXmlAttributesEnd();
		if (this.dxf) {
			//AscCommonExcel.CellXfs
			this.dxf.toXml(writer, "dxf", "x:", "", "x:");
		}
		if (this.filter) {
			this.filter.toXml(writer, "filter", "", "x:");
		}
		//WritingNullable(m_oExtLst, writer.WriteXmlString(m_oExtLst.toXMLWithNS("")););
		writer.WriteXmlNodeEnd(name);
	};

	Asc.CT_SortRules.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("sortRule" === name) {
				var val = new Asc.CT_SortRule();
				val.fromXml(reader);
				this.sortRule.push(val);
			} else if ("extLst" === name) {
			}
		}
	};

	Asc.CT_SortRules.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("sortMethod" === reader.GetName()) {
				val = reader.GetValue();
				this.sortMethod = FromXml_ST_SortMethod(val);
			} else if ("caseSensitive" === reader.GetName()) {
				val = reader.GetValue();
				this.caseSensitive = val;
			}
		}
	};

	Asc.CT_SortRules.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name/*"sortRules"*/);
		writer.WriteXmlNullableAttributeString("sortMethod", ToXml_ST_SortMethod(this.sortMethod));
		writer.WriteXmlNullableAttributeString("caseSensitive", this.caseSensitive);
		writer.WriteXmlAttributesEnd();
		for (var i = 0; i < this.sortRule.length; ++i) {
			this.sortRule[i].toXml(writer, "sortRule");
		}
		//WritingNullable(m_oExtLst, writer.WriteXmlString(m_oExtLst.toXMLWithNS("")););
		writer.WriteXmlNodeEnd(name);
	};


	Asc.CT_SortRule.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		var val;
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("dxf" === name) {
				val = new AscCommonExcel.CellXfs();
				val.fromXml(reader);
				if (!this.sortCondition) {
					this.sortCondition = new AscCommonExcel.SortCondition();
				}
				this.sortCondition.dxf = val;
			} else if ("richSortCondition" === name) {
				//в ser не читается, пока не поддерживаю и тут
			} else if ("sortCondition" === name) {
				if (!this.sortCondition) {
					this.sortCondition = new AscCommonExcel.SortCondition();
				}
				this.sortCondition.fromXml(reader);
			}
		}
	};

	Asc.CT_SortRule.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("colId" === reader.GetName()) {
				val = reader.GetValue();
				this.colId = val;
			} else if ("id" === reader.GetName()) {
				val = reader.GetValue();
				this.id = val;
			}
		}
	};

	Asc.CT_SortRule.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributeString("colId", this.colId);
		writer.WriteXmlAttributeStringEncode("id", this.id);
		writer.WriteXmlAttributesEnd();

		if (this.sortCondition && this.sortCondition.dxf) {
			this.sortCondition.dxf.toXml(writer, "dxf", "", "x:");
		}
		if (this.sortCondition) {
			this.sortCondition.toXml(writer, "sortCondition");
		}

		writer.WriteXmlNodeEnd(name);
	};

	function CT_CComments() {
		this.authors = null;
		this.commentList = null;
	}

	CT_CComments.prototype.fromXml = function (reader) {
		var t = this;
		reader.readXmlArray("comments", function () {
			if (reader.IsEmptyNode()) {
				return;
			}
			var depth = reader.GetDepth();
			var val;
			while (reader.ReadNextSiblingNode(depth)) {
				var name = reader.GetNameNoNS();

				if ("authors" === name) {
					val = new CT_CAuthors();
					val.fromXml(reader);
					t.authors = val;
				} else if ("commentList" === name) {
					val = new CT_CCommentList();
					val.fromXml(reader);
					t.commentList = val;
				}
			}
		});
	};

	CT_CComments.prototype.toXml = function (writer) {
		writer.WriteXmlString(
			"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" mc:Ignorable=\"xr\">");
		if (this.authors) {
			this.authors.toXml(writer);
		}
		if (this.commentList) {
			this.commentList.toXml(writer);
		}
		writer.WriteXmlString("</comments>");
	};


	function CT_CAuthors() {
		this.arr = [];
	}

	CT_CAuthors.prototype.fromXml = function (reader) {
		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("author" === name) {
				this.arr.push(prepareTextFromXml(reader.GetTextDecodeXml()));
			}
		}
	};

	CT_CAuthors.prototype.toXml = function (writer) {
		writer.WriteXmlString("<authors>");

		for (var i = 0; i < this.arr.length; ++i) {
			writer.WriteXmlString("<author>");
			writer.WriteXmlStringEncode(this.arr[i]);
			writer.WriteXmlString("</author>");
		}
		writer.WriteXmlString("</authors>");
	};

	function CT_CCommentList() {
		this.arr = [];
	}

	CT_CCommentList.prototype.fromXml = function (reader) {
		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("comment" === name) {
				var val = new CT_CComment();
				val.fromXml(reader);
				this.arr.push(val);
			}
		}
	};

	CT_CCommentList.prototype.toXml = function (writer) {
		writer.WriteXmlString("<commentList>");

		for (var i = 0; i < this.arr.length; ++i) {
			if (this.arr[i]) {

				/*var val = new CT_Si();
				if (typeof si === 'string') {
					val.text = elem;
				} else {
					val.multiText = elem;
				}*/

				this.arr[i].toXml(writer);
			}
		}

		writer.WriteXmlString("</commentList>");
	};

	function CT_CComment() {
		this.ref = null;
		this.authorId = null;
		this.uid = null;

		this.oText = null;
	}

	CT_CComment.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("text" === name) {
				var val = new CT_Si();
				//val.clean();
				val.fromXml(reader)
				this.oText = val;
			}
		}
	};

	CT_CComment.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("ref" === reader.GetName()) {
				val = reader.GetValue();
				this.ref = val;
			} else if ("authorId" === reader.GetName()) {
				val = reader.GetValue();
				this.authorId = val;
			} else if ("xr:uid" === reader.GetName()) {
				val = reader.GetValue();
				this.uid = val;
			}
		}
	};
	CT_CComment.prototype.toXml = function (writer) {
		if (null != this.ref && null != this.authorId && null != this.uid) {
			writer.WriteXmlString("<comment");
			writer.WriteXmlNullableAttributeStringEncode("ref", this.ref);
			writer.WriteXmlNullableAttributeString("authorId", this.authorId);
			writer.WriteXmlNullableAttributeString("xr:uid", this.uid);
			writer.WriteXmlString(">");

			if (this.oText) {
				this.oText.toXml(writer, "text");
			}
			writer.WriteXmlString("</comment>");
		}
	};
	CT_CComment.prototype.generateText = function (pCommentData, personList) {
		//parseCommentData
		var MAX_STRING_LEN = 0x7FFF;
		var addCommentRun = function (oSi, text, isBold, nLimit) {
			var pRun = new AscCommonExcel.CMultiTextElem();
			var pRPr = new AscCommonExcel.Font();
			if (isBold) {
				pRPr.b = true;
			}
			pRPr.fn = "Tahoma";
			pRPr.fs = 9;

			pRun.text = text;
			nLimit -= text.length;

			pRun.format = pRPr;

			if (!oSi.multiText) {
				oSi.multiText = [];
			}

			oSi.multiText.push(pRun);
			return nLimit;
		};

		var getThreadedCommentAuthor = function (mapPersonList, personId, sDefault) {
			if (mapPersonList && personId != null) {
				var person = mapPersonList.find(function isPrime(element) {
					return personId === element.userId;
				});

				if (person) {
					return person.displayName;
				}
			}
			return sDefault;
		};

		/*void BinaryCommentReader::addThreadedComment(OOX::Spreadsheet::CSi& oSi, OOX::Spreadsheet::CThreadedComment* pThreadedComment, nullable<std::unordered_map<std::wstring, OOX::Spreadsheet::CPerson*>>& mapPersonList)
		{
			int nLimit = OOX::Spreadsheet::SpreadsheetCommon::MAX_STRING_LEN;
			if(pThreadedComment->m_oText.IsInit())
			{
				std::wstring displayName = getThreadedCommentAuthor(mapPersonList, pThreadedComment->personId, L"Comment");
				nLimit = addCommentRun(oSi, displayName + L":", true, nLimit);
				if (nLimit <= 0)
					return;
				nLimit = addCommentRun(oSi, L"\n" + pThreadedComment->m_oText->ToString() + L"\n", false, nLimit);
				if (nLimit <= 0)
					return;
			}
			for(size_t i = 0; i < pThreadedComment->m_arrReplies.size(); ++i)
			{
				if(pThreadedComment->m_arrReplies[i]->m_oText.IsInit())
				{
					std::wstring displayName = getThreadedCommentAuthor(mapPersonList, pThreadedComment->m_arrReplies[i]->personId, L"Reply");
					nLimit = addCommentRun(oSi, displayName + L":", true, nLimit);
					if (nLimit <= 0)
						return;
					nLimit = addCommentRun(oSi, L"\n" + pThreadedComment->m_arrReplies[i]->m_oText->ToString() + L"\n", false, nLimit);
					if (nLimit <= 0)
						return;
				}
			}
		}*/

		var oSi;
		var nLimit = MAX_STRING_LEN;
		if (pCommentData.aReplies && pCommentData.aReplies.length) {
			oSi = new CT_Si();
			if (pCommentData.sText) {
				var displayName = getThreadedCommentAuthor(personList, pCommentData.sUserId, "Comment");
				nLimit = addCommentRun(oSi, displayName + ":", true, nLimit);
				if (nLimit <= 0) {
					return;
				}
				nLimit = addCommentRun(oSi, "\n" + pCommentData.sText + "\n", false, nLimit);
				if (nLimit <= 0) {
					return;
				}
			}
			for (var i = 0; i < pCommentData.aReplies.length; ++i) {
				if (pCommentData.aReplies[i].sText) {
					var displayName = getThreadedCommentAuthor(personList, pCommentData.aReplies[i].sUserId, "Reply");
					nLimit = addCommentRun(oSi, displayName + ":", true, nLimit);
					if (nLimit <= 0) {
						return;
					}
					nLimit = addCommentRun(oSi, "\n" + pCommentData.aReplies[i].sText + "\n", false, nLimit);
					if (nLimit <= 0) {
						return;
					}
				}
			}
			this.oText = oSi;
		} else if (pCommentData && pCommentData.sText) {
			oSi = new CT_Si();
			if (!pCommentData.sUserName) {
				addCommentRun(oSi, pCommentData.sText, false, nLimit);
			} else {
				nLimit = addCommentRun(oSi, pCommentData.sUserName + (":"), true, nLimit);
				if (nLimit <= 0) {
					return;
				}
				addCommentRun(oSi, "\n" + pCommentData.sText, false, nLimit);
			}
			this.oText = oSi;
		}
	};

	function CT_CThreadedComments() {
		this.arr = [];
		this.m_mapTopLevelThreadedComments = [];
	}

	CT_CThreadedComments.prototype.fromXml = function (reader) {
		if (!reader.ReadNextNode()) {
			return;
		}

		var name = reader.GetNameNoNS();
		if ("ThreadedComments" === name) {
			if (reader.IsEmptyNode()) {
				return;
			}
			var depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				var name2 = reader.GetNameNoNS();

				if ("threadedComment" === name2) {
					var val = new CT_CThreadedComment();
					val.fromXml(reader);
					this.arr.push(val);
				}
			}
		}

		this.PrepareTopLevelComments();
	};

	CT_CThreadedComments.prototype.PrepareTopLevelComments = function () {
		//find TopLevelComments
		var pThreadedComment, i;
		for (i = 0; i < this.arr.length; ++i) {
			pThreadedComment = this.arr[i];
			if (pThreadedComment.id && !pThreadedComment.parentId) {
				this.m_mapTopLevelThreadedComments[pThreadedComment.id] = pThreadedComment;
			}
		}

		//to remove reply duplicates
		var mapUniqueue = [];
		//add Replies
		for (i = 0; i < this.arr.length; ++i) {
			pThreadedComment = this.arr[i];
			if (pThreadedComment.parentId) {
				if (pThreadedComment.parentId === "{00000000-0000-0000-0000-000000000000}" && pThreadedComment.ref) {
					if (pThreadedComment.dT && !mapUniqueue[pThreadedComment.dT + ""]) {
						mapUniqueue[pThreadedComment.dT + ""] = true;
						//find parents by ref
						for (var it in this.m_mapTopLevelThreadedComments) {
							if (this.m_mapTopLevelThreadedComments[it].ref && pThreadedComment.ref === this.m_mapTopLevelThreadedComments[it].ref) {
								this.m_mapTopLevelThreadedComments[it].m_arrReplies.push(pThreadedComment);
								break;
							}
						}
					}
				} else {
					var oFind = this.m_mapTopLevelThreadedComments[pThreadedComment.parentId + ""];
					if (oFind) {
						oFind.m_arrReplies.push(pThreadedComment);
					}
				}
			}
		}
		//TODO sort Replies
		/*for (std::unordered_map<std::wstring, CThreadedComment*>::const_iterator it = m_mapTopLevelThreadedComments.begin(); it != m_mapTopLevelThreadedComments.end(); ++it)
		{
			std::sort (it->second->m_arrReplies.begin(), it->second->m_arrReplies.end(), CThreadedComment::Compare);
		}*/
	};


	CT_CThreadedComments.prototype.toXml = function (writer) {
		writer.WriteXmlString("<ThreadedComments");
		writer.WriteXmlString(" xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\"");
		writer.WriteXmlString(" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
		writer.WriteXmlString(">");

		for (var i = 0; i < this.arr.length; ++i) {
			if (this.arr[i]) {
				this.arr[i].toXml(writer);
			}
		}

		writer.WriteXmlString("</ThreadedComments>");
	};

	function CT_CThreadedComment() {
		this.ref = null;
		this.dT = null;
		this.personId = null;
		this.id = null;
		this.parentId = null;
		this.done = null;

		this.text = null;
		this.mentions = null;

		this.m_arrReplies = [];
		//extLst
	}

	CT_CThreadedComment.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			return;
		}

		var t = this;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();

			if ("text" === name) {
				this.text = prepareTextFromXml(reader.GetTextDecodeXml());
			} else if ("mentions" === name) {
				//m_oMentions = oReader;
				reader.readXmlArray("mention", function () {
					var commentMention = new CT_CThreadedCommentMention();
					if (commentMention) {
						commentMention.fromXml(reader);
						t.mentions.push(commentMention);
					}
				});
			}
		}

		//PrepareText
	};

	CT_CThreadedComment.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("ref" === reader.GetName()) {
				val = reader.GetValue();
				this.ref = val;
			} else if ("dT" === reader.GetName()) {
				val = reader.GetValue();
				this.dT = val;
			} else if ("personId" === reader.GetName()) {
				val = reader.GetValue();
				this.personId = val;
			} else if ("id" === reader.GetName()) {
				val = reader.GetValue();
				this.id = val;
			} else if ("parentId" === reader.GetName()) {
				val = reader.GetValue();
				this.parentId = val;
			} else if ("done" === reader.GetName()) {
				val = reader.GetValue();
				this.done = val;
			}
		}
	};

	CT_CThreadedComment.prototype.toXml = function (writer) {
		writer.WriteXmlString("<threadedComment");
		writer.WriteXmlNullableAttributeStringEncode("ref", this.ref);
		writer.WriteXmlNullableAttributeString("dT", this.dT);
		writer.WriteXmlNullableAttributeString("personId", this.personId);
		writer.WriteXmlNullableAttributeString("id", this.id);
		writer.WriteXmlNullableAttributeString("parentId", this.parentId);
		writer.WriteXmlNullableAttributeBool("done", boolToNumber(this.done));
		writer.WriteXmlString(">");

		if (this.text) {
			writer.WriteXmlString("<text xml:space=\"preserve\">");
			writer.WriteXmlStringEncode(this.text);
			//last '\n' not in format but excel add it
			writer.WriteXmlString("\n");//toodo \r?
			writer.WriteXmlString("</text>");
		}
		if (this.mentions) {
			writer.WriteXmlArray(this.mentions, "mention", "mentions");
		}
		writer.WriteXmlString("</threadedComment>");
	};

	function CThreadedCommentMention() {
		this.mentionpersonId = null;
		this.mentionId = null;
		this.startIndex = null;
		this.length = null;
	}

	CThreadedCommentMention.prototype.fromXml = function (reader) {
		this.readAttr(reader);

		if (reader.IsEmptyNode()) {
			reader.ReadTillEnd();
		}
	};

	CThreadedCommentMention.prototype.readAttr = function (reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			if ("mentionpersonId" === reader.GetName()) {
				val = reader.GetValue();
				this.mentionpersonId = val;
			} else if ("mentionId" === reader.GetName()) {
				val = reader.GetValue();
				this.mentionId = val;
			} else if ("startIndex" === reader.GetName()) {
				val = reader.GetValue();
				this.startIndex = val;
			} else if ("length" === reader.GetName()) {
				val = reader.GetValue();
				this.length = val;
			}
		}
	};

	CThreadedCommentMention.prototype.toXml = function (writer) {
		writer.WriteXmlNullableAttributeString("mentionpersonId", this.mentionpersonId);
		writer.WriteXmlNullableAttributeString("mentionId", this.mentionId);
		writer.WriteXmlNullableAttributeNumber("startIndex", this.startIndex);
		writer.WriteXmlNullableAttributeNumber("length", this.length);
	};


	 function PrepareComments (ws, xmlParserContext, comments, pThreadedComments, personList) {
		var m_mapComments = {};
		var pVmlDrawing = xmlParserContext.InitOpenManager.legacyDrawing;
		if (!pVmlDrawing || !comments) {
			return;
		}

		var wb = ws.workbook;
		var mapCheckCopyThreadedComments = [];
		var arAuthors = comments.authors && comments.authors.arr;

		var i, j, nRow, nCol, nAuthorId, pCommentItem;
		if (comments.commentList) {
			var aComments = comments.commentList.arr;

			for (i = 0; i < aComments.length; ++i) {
				var pComment = aComments[i];
				if (!pComment) {
					continue;
				}


				var bThreadedCommentCopy = false;
				var pThreadedComment = null;
				if (pThreadedComments) {
					var pFind;

					var isPlaceholder = false;
					if (pComment.authorId) {
						nAuthorId = parseInt(pComment.authorId);

						if (nAuthorId >= 0 && nAuthorId < arAuthors.length) {
							var sAuthor = arAuthors[nAuthorId];
							if ("tc=" === sAuthor.substring(0, 3)) {
								isPlaceholder = true;
								var sGUID = sAuthor.substr(3);
								//todo IsZero() is added to fix comments with zero ids(5.4.0)(bug 42947). Remove after few releases
								if ("{00000000-0000-0000-0000-000000000000}" === sGUID && pComment.ref) {
									for (j in pThreadedComments.m_mapTopLevelThreadedComments) {
										var it = pThreadedComments.m_mapTopLevelThreadedComments[j];
										if (it.ref && pComment.ref === it.ref) {
											pFind = it;
											break;
										}
									}
								} else {
									pFind = pThreadedComments.m_mapTopLevelThreadedComments[sGUID];
								}

							}
						}
					}

					if (pFind) {
						pThreadedComment = pFind;
						if (mapCheckCopyThreadedComments[pThreadedComment.id]) {
							bThreadedCommentCopy = true;
						} else {
							mapCheckCopyThreadedComments[pThreadedComment.id + ""] = 1;
						}
					} else if (isPlaceholder) {
						continue;
					}
				}


				if (pComment.ref && pComment.authorId) {
					var sRef = AscCommonExcel.g_oRangeCache.getAscRange(pComment.ref);
					if (sRef) {
						nRow = sRef.r1;
						nCol = sRef.c1;
						pCommentItem = new Asc.asc_CCommentData();
						pCommentItem.asc_putDocumentFlag(false);
						pCommentItem.nRow = nRow;
						pCommentItem.nCol = nCol;

						//TODO флаг всегда будет false
						/*if (pCommentItem.asc_getDocumentFlag()) {
							pCommentItem.nId = "doc_" + (this.wb.aComments.length + 1);
						} else {
							pCommentItem.wsId = ws.Id;
							pCommentItem.nId = "sheet" + pCommentItem.wsId + "_" + (ws.aComments.length + 1);
						}*/

						nAuthorId = parseInt(pComment.authorId);
						if (nAuthorId >= 0 && nAuthorId < arAuthors.length) {
							pCommentItem.asc_putUserName(arAuthors[nAuthorId]);
						}

						var pSi = pComment.oText;
						if (pSi) {
							pCommentItem.asc_putText(pSi.getText());
						}

						pCommentItem.threadedComment = pThreadedComment;//c_oSer_Comments.ThreadedComment
						pCommentItem.ThreadedCommentCopy = bThreadedCommentCopy;//bool m_bThreadedCommentCopy

						var sNewId = nRow + "-" + nCol;
						m_mapComments[sNewId] = pCommentItem;
					}
				}
			}
		}


		for (i = 0; i < pVmlDrawing.items.length; ++i) {
			var pShape = pVmlDrawing.items[i];
			if (!pShape || AscDFH.historyitem_type_VMLShape !== pShape.getObjectType()) {
				continue;
			}

			/*if (pShape.sId)
			{//mark shape as used
				boost::unordered_map<std::wstring, OOX::CVmlDrawing::_vml_shape>::iterator pFind = pVmlDrawing->m_mapShapes.find(pShape->m_sId.get());
				if (pFind != pVmlDrawing->m_mapShapes.end())
				{
					pFind->second.bUsed = true;
				}
			}*/

			for (j = 0; j < pShape.items.length; ++j) {
				var pElem = pShape.items[j];

				if (!pElem) {
					continue;
				}

				if (AscDFH.historyitem_type_VMLClientData === pElem.getObjectType()) {
					var pClientData = pElem;
					if (null != pClientData.m_oRow && null != pClientData.m_oColumn) {
						nRow = parseInt(pClientData.m_oRow);
						nCol = parseInt(pClientData.m_oColumn);
						var sId = nRow + "" + "-" + nCol + "";

						pCommentItem = m_mapComments[sId];
						if (pCommentItem) {
							/*if(pShape->m_sGfxData.IsInit())
								pCommentItem->m_sGfxdata = *pShape->m_sGfxData;*/

							var oCommentCoords = new AscCommonExcel.asc_CCommentCoords();
							var m_aAnchor = [];
							pClientData.getAnchorArray(m_aAnchor);
							if (8 <= m_aAnchor.length) {
								oCommentCoords.nLeft = Math.abs(m_aAnchor[0]);
								oCommentCoords.nLeftOffset = Math.abs(m_aAnchor[1]);
								oCommentCoords.nTop = Math.abs(m_aAnchor[2]);
								oCommentCoords.nTopOffset = Math.abs(m_aAnchor[3]);
								oCommentCoords.nRight = Math.abs(m_aAnchor[4]);
								oCommentCoords.nRightOffset = Math.abs(m_aAnchor[5]);
								oCommentCoords.nBottom = Math.abs(m_aAnchor[6]);
								oCommentCoords.nBottomOffset = Math.abs(m_aAnchor[7]);
							}
							oCommentCoords.bMoveWithCells = pClientData.m_oMoveWithCells;
							oCommentCoords.bSizeWithCells = pClientData.m_oSizeWithCells;

							oCommentCoords.nCol = nCol;
							oCommentCoords.nRow = nRow;

							pCommentItem.coords = oCommentCoords;

							//todo bHidden ?
							//oCommentCoords->m_bVisible = pClientData->m_oVisible;
							//pCommentItem.bHidden = !pClientData.m_oVisible;

							//pCommentItem.coords = oCommentCoords;

							if (pShape.m_oFillColor) {
								/*BYTE r = pShape->m_oFillColor->Get_R();
								BYTE g = pShape->m_oFillColor->Get_G();
								BYTE b = pShape->m_oFillColor->Get_B();

								std::wstringstream sstream;
								sstream << boost::wformat( L"%02X%02X%02X" ) % r % g % b;

								pCommentItem->m_sFillColorRgb = sstream.str();*/
							}

							var oPoint = new AscFormat.CPoint(), oUCssValue;
							for (var k = 0; k < pShape.m_oStyle.m_arrProperties.length; ++k) {
								if (!pShape.m_oStyle.m_arrProperties[i]) {
									continue;
								}

								var oProperty = pShape.m_oStyle.m_arrProperties[k];
								if (AscFormat.ECssPropertyType.cssptMarginLeft === oProperty.get_Type()) {
									oUCssValue = oProperty.m_oValue;
									if (AscFormat.ECssUnitsType.cssunitstypeUnits === oUCssValue.oValue.m_eType) {
										oPoint.FromPoints(oUCssValue.oValue.dValue);
										pCommentItem.coords.dLeftMM = oPoint.ToMm();
									}
								} else if (AscFormat.ECssPropertyType.cssptMarginTop === oProperty.get_Type()) {
									oUCssValue = oProperty.m_oValue;
									if (AscFormat.ECssUnitsType.cssunitstypeUnits === oUCssValue.oValue.m_eType) {
										oPoint.FromPoints(oUCssValue.oValue.dValue);
										pCommentItem.coords.dTopMM = oPoint.ToMm();
									}
								} else if (AscFormat.ECssPropertyType.cssptWidth === oProperty.get_Type()) {
									oUCssValue = oProperty.m_oValue;
									if (AscFormat.ECssUnitsType.cssunitstypeUnits === oUCssValue.oValue.m_eType) {
										oPoint.FromPoints(oUCssValue.oValue.dValue);
										pCommentItem.coords.dWidthMM = oPoint.ToMm();
									}
								} else if (AscFormat.ECssPropertyType.cssptHeight === oProperty.get_Type()) {
									oUCssValue = oProperty.m_oValue;
									if (AscFormat.ECssUnitsType.cssunitstypeUnits === oUCssValue.oValue.m_eType) {
										oPoint.FromPoints(oUCssValue.oValue.dValue);
										pCommentItem.coords.dHeightMM = oPoint.ToMm();
									}
								}
							}
						}
					}
				}
			}
		}


		var applyThreadedComment = function (_commentData, _threadedComment) {
			oAdditionalData.isThreadedComment = true;
			_commentData.asc_putSolved(false);
			_commentData.aReplies = [];

			if (_threadedComment.dT != null) {
				_commentData.asc_putTime("");
				var dateMs = AscCommon.getTimeISO8601(_threadedComment.dT);
				if (!isNaN(dateMs)) {
					_commentData.asc_putOnlyOfficeTime(dateMs + "");
				}
			}

			if (_threadedComment.personId != null) {
				var person = personList.getByGuid(_threadedComment.personId);
				if (person) {
					_commentData.asc_putUserName(person.displayName);
					_commentData.asc_putUserId(person.userId);
					_commentData.asc_putProviderId(person.providerId);
				}
			}

			if (_threadedComment.id != null) {
				_commentData.asc_putGuid(_threadedComment.id);
			}

			if (_threadedComment.done != null) {
				_commentData.asc_putSolved(_threadedComment.done === "1");
			}

			if (_threadedComment.text != null) {
				_commentData.asc_putText(_threadedComment.text);
			}

			if (_threadedComment.m_arrReplies && _threadedComment.m_arrReplies.length) {
				for (var j = 0; j < _threadedComment.m_arrReplies.length; j++) {
					var reply = new Asc.asc_CCommentData();
					applyThreadedComment(reply, _threadedComment.m_arrReplies[j]);
					_commentData.asc_addReply(reply);
				}
			}
		};

		for (i in m_mapComments) {

			if (m_mapComments[i].asc_getDocumentFlag()) {
				m_mapComments[i].nId = "doc_" + (wb.aComments.length + 1);
			} else {
				m_mapComments[i].wsId = ws.Id;
				m_mapComments[i].nId = "sheet" + m_mapComments[i].wsId + "_" + (ws.aComments.length + 1);
			}

			var oAdditionalData = {isThreadedComment: false};
			if (m_mapComments[i].threadedComment) {
				applyThreadedComment(m_mapComments[i], m_mapComments[i].threadedComment);

			}
			xmlParserContext.InitOpenManager.prepareComments(ws, m_mapComments[i].coords, [m_mapComments[i]], oAdditionalData);
		}
	};

	// var _x2tFromXml = ''
	// var _x2t = ''
	// var _documentation = ''
	// var _serialize = ''
	//
	// var _x2tToXml = ''
	//
	//
	// //by test automatic add function
	// analizeXmlFrom(_x2tFromXml);
	//
	// function analizeAttr(x2t, documentation, serialize) {
	// 	var isUpperCaseName = false;
	//
	// 	var _getAttrVal = function (from, _name) {
	// 		var _split = from.split(_name + "=\"")
	// 		var _res = null;
	// 		if (_split && _split[1]) {
	// 			for (var i = 0; i < _split[1].length; i++) {
	// 				if (!_res) {
	// 					_res = ""
	// 				}
	// 				if (_split[1][i] === '"') {
	// 					break;
	// 				}
	// 				_res += _split[1][i];
	// 			}
	// 		}
	// 		return _res;
	// 	};
	//
	// 	var docSplit = documentation && documentation.split("element");
	// 	var attributeMapType = [];
	// 	var documentationClassStr = "";
	// 	if (docSplit) {
	// 		for (var j = 0; j < docSplit.length; j++) {
	// 			var _name = _getAttrVal(docSplit[j], "name");
	// 			var _type = _getAttrVal(docSplit[j], "type");
	// 			if (_name && _type) {
	// 				attributeMapType[_name] = _type;
	// 				documentationClassStr += "\n" + "this." + _name + " = null;//" + _type
	// 			}
	// 		}
	// 	}
	//
	// 	var docSplit = documentation && documentation.split("attribute");
	// 	if (docSplit) {
	// 		for (var j = 0; j < docSplit.length; j++) {
	// 			var _name = _getAttrVal(docSplit[j], "name");
	// 			var _type = _getAttrVal(docSplit[j], "type");
	// 			if (_name && _type) {
	// 				attributeMapType[_name] = _type;
	// 				documentationClassStr += "\n" + "this." + _name + " = null;//" + _type
	// 			}
	// 		}
	// 	}
	// 	console.log(documentationClassStr)
	//
	//
	// 	var getFuncName = function (_attrName) {
	// 		var _res = null;
	// 		var _type = attributeMapType[_attrName] && attributeMapType[_attrName].toLowerCase();
	// 		if (_type) {
	// 			if (-1 !== _type.indexOf("bool")) {
	// 				_res = "GetValueBool()"
	// 			} else if (-1 !== _type.indexOf("int")) {
	// 				_res = "GetValueInt()"
	// 			} else if (-1 !== _type.indexOf("double")) {
	// 				_res = "GetValueDouble()"
	// 			}
	// 		}
	// 		return _res;
	// 	};
	//
	// 	var serializeSplit = serialize ? serialize.split("if (") : null;
	// 	var getFuncNameFromSerialize = function (_attrName, _upperAttrName) {
	// 		var _res = null;
	// 		if (!serializeSplit) {
	// 			return null;
	// 		}
	//
	// 		for (var n = 0; n < serializeSplit.length; n++) {
	// 			var isFound
	// 			if (-1 !== (serializeSplit[n].indexOf(_attrName + " = "))) {
	// 				isFound = true
	// 			} else if (-1 !== (serializeSplit[n].indexOf(_upperAttrName + " = "))) {
	// 				isUpperCaseName = true;
	// 				isFound = true
	// 			}
	//
	// 			if (isFound) {
	// 				var _split2 = serializeSplit[n].split(" = ")
	// 				if (_split2 && _split2[1]) {
	// 					var _type = _split2[1].toLowerCase()
	// 					if (-1 !== _type.indexOf("bool")) {
	// 						_res = "GetValueBool()"
	// 					} else if (-1 !== _type.indexOf("int") || -1 !== _type.indexOf("long")) {
	// 						_res = "GetValueInt()"
	// 					} else if (-1 !== _type.indexOf("double")) {
	// 						_res = "GetValueDouble()"
	// 					}
	// 				}
	// 				return _res;
	// 			}
	// 		}
	// 		return _res;
	// 	};
	//
	// 	var res = "TEST.prototype.readAttr = function (reader) {\n"
	//
	// 	res += '\n//documentation\n/*' + documentation + '*/\n';
	// 	res += '\n//x2t\n/*' + x2t + '*/\n';
	// 	res += '\n//serialize\n/*' + serialize + '*/\n\n';
	//
	// 	var initUpperCase = false;
	// 	var x2tSplit = x2t.indexOf("WritingElement_ReadAttributes_Read_else_if") !== -1 ? x2t.split("WritingElement_ReadAttributes_Read_else_if") :
	// 		x2t.split("WritingElement_ReadAttributes_Read_if")
	// 	res += "var val;\n" + "\t\twhile (reader.MoveToNextAttribute()) {\n"
	// 	for (var i = 0; i < x2tSplit.length; i++) {
	// 		var attr = x2tSplit[i].split('"');
	// 		if (attr[1]) {
	// 			if (i !== 0) {
	// 				res += "else "
	// 			}
	// 			res += 'if ("' + attr[1] + '" === reader.GetName()) {\n'
	//
	// 			var funcName = getFuncName(attr[1]);
	// 			if (funcName === null) {
	// 				funcName = getFuncNameFromSerialize(attr[1], attr[1][0].toUpperCase() + attr[1].slice(1));
	// 			} else if (!initUpperCase) {
	// 				getFuncNameFromSerialize(attr[1], attr[1][0].toUpperCase() + attr[1].slice(1));
	// 			}
	//
	// 			initUpperCase = true;
	//
	// 			if (funcName === null) {
	// 				funcName = "GetValue()";
	// 			}
	//
	// 			var attrName = attr[1];
	// 			if (isUpperCaseName) {
	// 				attrName = attr[1][0].toUpperCase() + attr[1].slice(1);
	// 			}
	//
	// 			res += "val = reader." + funcName + ";\n"
	// 			res += "this." + attrName + " = val;\n} ";
	// 		}
	// 	}
	// 	res += "}\n};";
	// 	return res;
	// }
	//
	// function analizeXmlFrom(x2tFromXml) {
	// 	var isneedGenerateAttr;
	// 	var res = "TEST.prototype.fromXml = function (reader) {\n"
	//
	// 	res += '\n/*' + x2tFromXml + '*/\n\n';
	//
	// 	if (-1 != x2tFromXml.indexOf("ReadAttributes(")) {
	// 		res += "this.readAttr(reader);\n\n";
	// 		isneedGenerateAttr = true;
	// 	}
	// 	if (-1 != x2tFromXml.indexOf("oReader.IsEmptyNode()")) {
	// 		res += "if (reader.IsEmptyNode())\n";
	// 		if (-1 != x2tFromXml.indexOf("oReader.ReadTillEnd();")) {
	// 			res += "reader.ReadTillEnd();";
	// 		}
	// 	} else if (-1 != x2tFromXml.indexOf("oReader.ReadTillEnd();")) {
	// 		res += "reader.ReadTillEnd();";
	// 	}
	// 	if (-1 != x2tFromXml.indexOf("oReader.ReadNextSiblingNode(")) {
	// 		res += "var val;\n" + "\tvar depth = reader.GetDepth();\n";
	// 		res += "while (reader.ReadNextSiblingNode(depth)) {\n"
	// 		res += "var name = reader.GetNameNoNS();\n"
	//
	// 		if (-1 != x2tFromXml.indexOf("GetNameNoNS(oReader.GetName());")) {
	// 			var splitXml = x2tFromXml.substring(x2tFromXml.indexOf("GetNameNoNS(oReader.GetName());")).split('L"')
	// 			for (var i = 1; i < splitXml.length; i++) {
	// 				var namePos = splitXml[i].indexOf("sName")
	// 				var test = splitXml[i].substring(0, namePos + 5)
	// 				test = test.replace("sName", "name")
	// 				test = test.replace(")", "")
	// 				test = test.replace("==", "===")
	// 				test = "\"" + test
	//
	// 				var ifst = "if"
	// 				if (i > 1) {
	// 					ifst = "else if"
	// 				}
	// 				res += ifst + " ( " + test + ") {\n\n} "
	// 			}
	// 		}
	//
	//
	// 		res += "}\n"
	// 	}
	// 	res += "};"
	//
	// 	if (isneedGenerateAttr) {
	// 		res += "\n\n" + analizeAttr(_x2t, _documentation, _serialize);
	// 	}
	// 	console.log(res);
	// }
	//
	//
	// analizeXmlTo(_x2tToXml, false);
	//
	// function analizeXmlTo(x2tToXml, isUpperCaseName) {
	//
	// 	var res = "TEST.prototype.toXml = function (writer) {\n"
	//
	// 	res += '\n/*' + x2tToXml + '*/\n\n';
	//
	// 	var toAttrMap = {
	// 		WritingStringNullableAttrInt: "WriteXmlNullableAttributeNumber",
	// 		WritingStringNullableAttrString: "WriteXmlNullableAttributeString",
	// 		WritingStringNullableAttrEncodeXmlString: "WriteXmlNullableAttributeStringEncode",
	// 		WritingStringAttrString: "WriteXmlAttributeString",
	// 		WritingStringAttrInt: "WriteXmlAttributeNumber",
	// 		WritingStringAttrEncodeXmlString: "WriteXmlAttributeStringEncode",
	// 		WritingStringNullableAttrBool: "WriteXmlNullableAttributeBool",
	// 		WriteEncodeXmlString: "WriteXmlStringEncode",
	// 		WritingStringNullableAttrDouble: "WriteXmlNullableAttributeDouble",
	// 		WritingStringAttrDouble: "WriteXmlAttributeDouble",
	// 		WritingStringNullableAttrInt2: "WriteXmlNullableAttributeNumber2",
	// 		WritingStringNullableAttrDouble2: "WriteXmlNullableAttributeDouble2",
	// 		WritingStringNullableAttrBool2: "WriteXmlNullableAttributeBool2"
	//
	// 	};
	//
	// 	var splitRows = x2tToXml.split("\n");
	// 	for (var i = 0; i < splitRows.length; i++) {
	// 		var isWriteAttr = false;
	// 		if (-1 != splitRows[i].indexOf("WritingStringNullableAttrInt")) {
	// 			isWriteAttr = "WritingStringNullableAttrInt";
	// 		} else if (-1 != splitRows[i].indexOf("WritingStringNullableAttrString")) {
	// 			isWriteAttr = "WritingStringNullableAttrString";
	// 		} else if (-1 != splitRows[i].indexOf("WritingStringNullableAttrEncodeXmlString")) {
	// 			isWriteAttr = "WritingStringNullableAttrEncodeXmlString";
	// 		} else if (-1 != splitRows[i].indexOf("WritingStringNullableAttrEncodeXmlString")) {
	// 			isWriteAttr = "WritingStringNullableAttrEncodeXmlString";
	// 		} else if (-1 != splitRows[i].indexOf("WritingStringAttrString")) {
	// 			isWriteAttr = "WritingStringAttrString";
	// 		} else if (-1 != splitRows[i].indexOf("WritingStringAttrInt")) {
	// 			isWriteAttr = "WritingStringAttrInt";
	// 		} else if (-1 != splitRows[i].indexOf("WritingStringAttrEncodeXmlString")) {
	// 			isWriteAttr = "WritingStringAttrEncodeXmlString"
	// 		} else if (-1 != splitRows[i].indexOf("WritingStringNullableAttrBool")) {
	// 			isWriteAttr = "WritingStringNullableAttrBool"
	// 		} else if (-1 != splitRows[i].indexOf("WritingStringNullableAttrDouble")) {
	// 			isWriteAttr = "WritingStringNullableAttrDouble"
	// 		} else if (-1 != splitRows[i].indexOf("WritingStringAttrDouble")) {
	// 			isWriteAttr = "WritingStringAttrDouble"
	// 		} else if (-1 != splitRows[i].indexOf("WritingStringNullableAttrInt2")) {
	// 			isWriteAttr = "WritingStringNullableAttrInt2"
	// 		} else if (-1 != splitRows[i].indexOf("WritingStringNullableAttrDouble2")) {
	// 			isWriteAttr = "WritingStringNullableAttrDouble2"
	// 		} else if (-1 != splitRows[i].indexOf("WritingStringNullableAttrBool2")) {
	// 			isWriteAttr = "WritingStringNullableAttrBool2"
	// 		} else if (-1 != splitRows[i].indexOf("WriteEncodeXmlString")) {
	// 			isWriteAttr = "WriteEncodeXmlString"
	// 		}
	//
	// 		if (isWriteAttr) {
	// 			var attrName = splitRows[i].indexOf('L"');
	// 			if (-1 !== attrName) {
	// 				attrName = splitRows[i].substring(attrName)
	// 				attrName = attrName.split('"')
	// 				if (attrName && attrName[1]) {
	// 					attrName = attrName[1]
	// 				}
	// 			}
	//
	// 			if (!attrName) {
	// 				res += "ERROR Do not found attr NAme"
	// 			} else {
	// 				var propName = attrName;
	// 				if (isUpperCaseName) {
	// 					propName = attrName[0].toUpperCase() + attrName.slice(1);
	// 				}
	// 				res += "writer." + toAttrMap[isWriteAttr] + '("' + attrName + '", ' + "this." + propName + ");\n"
	// 			}
	// 		} else {
	// 			if (-1 != splitRows[i].indexOf("WriteString")) {
	// 				splitRows[i] = splitRows[i].replaceAll("WriteString", "WriteXmlString");
	// 			}
	// 			if (-1 != splitRows[i].indexOf("std::wstring")) {
	// 				splitRows[i] = splitRows[i].replaceAll("std::wstring", "var");
	// 			}
	// 			if (-1 != splitRows[i].indexOf("size_t")) {
	// 				splitRows[i] = splitRows[i].replaceAll("size_t", "var");
	// 			}
	// 			if (-1 != splitRows[i].indexOf(' (L"')) {
	// 				splitRows[i] = splitRows[i].replaceAll(' (L"', '("');
	// 			}
	// 			if (-1 != splitRows[i].indexOf('(L"')) {
	// 				splitRows[i] = splitRows[i].replaceAll('(L"', '("');
	// 			}
	// 			if (-1 != splitRows[i].indexOf('(L')) {
	// 				splitRows[i] = splitRows[i].replaceAll('(L', '(');
	// 			}
	// 			if (-1 != splitRows[i].indexOf('L"')) {
	// 				splitRows[i] = splitRows[i].replaceAll('L"', '"');
	// 			}
	// 			if (-1 != splitRows[i].indexOf('.size()')) {
	// 				splitRows[i] = splitRows[i].replaceAll('.size()', '.length');
	// 			}
	// 			if (-1 != splitRows[i].indexOf('->toXML')) {
	// 				splitRows[i] = splitRows[i].replaceAll('->toXML', '.toXML');
	// 			}
	// 			if (-1 != splitRows[i].indexOf('->')) {
	// 				splitRows[i] = splitRows[i].replaceAll('->', '.');
	// 			}
	// 			if (-1 != splitRows[i].indexOf("WriteEncodeXmlString")) {
	// 				splitRows[i] = splitRows[i].replaceAll('WriteXmlStringEncode', '.');
	// 			}
	//
	//
	// 			res += splitRows[i] + "\n";
	// 		}
	//
	// 	}
	//
	// 	res += "};"
	//
	// 	console.log(res)
	// 	return res;
	// }


	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].CT_Workbook = CT_Workbook;
	window['AscCommonExcel'].CT_SharedStrings = CT_SharedStrings;
	window['AscCommonExcel'].CT_PersonList = CT_PersonList;
	window['AscCommonExcel'].CT_SheetData = CT_SheetData;
	window['AscCommonExcel'].CT_Value = CT_Value;
	window['AscCommonExcel'].CT_DrawingWS = CT_DrawingWS;
	window['AscCommonExcel'].CT_DrawingWSRef = CT_DrawingWSRef;
	window['AscCommonExcel'].CT_ExternalReference = CT_ExternalReference;
	window['AscCommonExcel'].CT_Connections = CT_Connections;
	window['AscCommonExcel'].CT_CComments = CT_CComments;
	window['AscCommonExcel'].CT_CThreadedComments = CT_CThreadedComments;

	window['AscCommonExcel'].ToXml_ST_DataValidationOperator = ToXml_ST_DataValidationOperator;
	window['AscCommonExcel'].FromXml_ST_DataValidationOperator = FromXml_ST_DataValidationOperator;
	window['AscCommonExcel'].ToXml_CFOperatorType = ToXml_CFOperatorType;
	window['AscCommonExcel'].FromXml_CFOperatorType = FromXml_CFOperatorType;
	window['AscCommonExcel'].ToXml_ST_TimePeriod = ToXml_ST_TimePeriod;
	window['AscCommonExcel'].FromXml_ST_TimePeriod = FromXml_ST_TimePeriod;
	window["AscCommonExcel"].FromXml_ST_FilterOperator = FromXml_ST_FilterOperator;
	window["AscCommonExcel"].ToXml_ST_FilterOperator = ToXml_ST_FilterOperator;
	window["AscCommonExcel"].ToXml_ST_DynamicFilterType = ToXml_ST_DynamicFilterType;
	window["AscCommonExcel"].FromXml_ST_DynamicFilterType = FromXml_ST_DynamicFilterType;
	window["AscCommonExcel"].ToXml_ST_DateTimeGrouping = ToXml_ST_DateTimeGrouping;
	window["AscCommonExcel"].FromXml_ST_DateTimeGrouping = FromXml_ST_DateTimeGrouping;
	window["AscCommonExcel"].ToXml_ST_HorizontalAlignment = ToXml_ST_HorizontalAlignment;
	window["AscCommonExcel"].FromXml_ST_HorizontalAlignment = FromXml_ST_HorizontalAlignment;
	window["AscCommonExcel"].ToXml_ST_VerticalAlignment = ToXml_ST_VerticalAlignment;
	window["AscCommonExcel"].FromXml_ST_VerticalAlignment = FromXml_ST_VerticalAlignment;
	window["AscCommonExcel"].ToXml_ST_CfvoType = ToXml_ST_CfvoType;
	window["AscCommonExcel"].FromXml_ST_CfvoType = FromXml_ST_CfvoType;
	window["AscCommonExcel"].ToXml_ST_IconSetType = ToXml_ST_IconSetType;
	window["AscCommonExcel"].FromXml_ST_IconSetType = FromXml_ST_IconSetType;
	window["AscCommonExcel"].ToXml_ST_CfType = ToXml_ST_CfType;
	window["AscCommonExcel"].FromXml_ST_CfType = FromXml_ST_CfType;
	window["AscCommonExcel"].ToXml_ST_DataValidationErrorStyle = ToXml_ST_DataValidationErrorStyle;
	window["AscCommonExcel"].FromXml_ST_DataValidationErrorStyle = FromXml_ST_DataValidationErrorStyle;
	window["AscCommonExcel"].ToXml_ST_DataValidationType = ToXml_ST_DataValidationType;
	window["AscCommonExcel"].FromXml_ST_DataValidationType = FromXml_ST_DataValidationType;
	window["AscCommonExcel"].ToXml_ST_DataValidationImeMode = ToXml_ST_DataValidationImeMode;
	window["AscCommonExcel"].FromXml_ST_DataValidationImeMode = FromXml_ST_DataValidationImeMode;
	window["AscCommonExcel"].ToXml_ST_GradientType = ToXml_ST_GradientType;
	window["AscCommonExcel"].FromXml_ST_GradientType = FromXml_ST_GradientType;

	window["AscCommonExcel"].PrepareComments = PrepareComments;


})(window);

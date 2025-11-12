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


(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {
	/*cFormulaFunctionGroup['_xlfn'] = [

	 cFILTERXML,//web not support in MS Office Online
	 cWEBSERVICE,//web not support in MS Office Online

	 cQUERYSTRING
	 ];*/

	var cBaseFunction = AscCommonExcel.cBaseFunction;
	var cFormulaFunctionGroup = AscCommonExcel.cFormulaFunctionGroup;
	var argType = Asc.c_oAscFormulaArgumentType;

	//xlfn functions list:

	//_xlfn.ACOT ; _xlfn.ACOTH ; _xlfn.AGGREGATE ; _xlfn.ARABIC ; _xlfn.BASE ; _xlfn.BETA.DIST ; _xlfn.BETA.INV ;
	// _xlfn.BINOM.DIST ; _xlfn.BINOM.DIST.RANGE ; _xlfn.BINOM.INV ; _xlfn.BITAND ; _xlfn.BITLSHIFT ; _xlfn.BITOR ;
	// _xlfn.BITRSHIFT ; _xlfn.BITXOR ; _xlfn.CEILING.MATH ; _xlfn.CEILING.PRECISE ; _xlfn.CHISQ.DIST ;
	// _xlfn.CHISQ.DIST.RT ; _xlfn.CHISQ.INV ; _xlfn.CHISQ.INV.RT ; _xlfn.CHISQ.TEST ; _xlfn.COMBINA ;
	// _xlfn.CONCAT ; _xlfn.CONFIDENCE.NORM ; _xlfn.CONFIDENCE.T ; _xlfn.COT ; _xlfn.COTH ; _xlfn.COVARIANCE.P ;
	// _xlfn.COVARIANCE.S ; _xlfn.CSC ; _xlfn.CSCH ; _xlfn.DAYS ; _xlfn.DECIMAL ; ECMA.CEILING ;
	// _xlfn.ERF.PRECISE ; _xlfn.ERFC.PRECISE ; _xlfn.EXPON.DIST ; _xlfn.F.DIST ; _xlfn.F.DIST.RT ;
	// _xlfn.F.INV ; _xlfn.F.INV.RT ; _xlfn.F.TEST ; _xlfn.FIELDVALUE ; _xlfn.FILTERXML ; _xlfn.FLOOR.MATH ;
	// _xlfn.FLOOR.PRECISE ; _xlfn.FORECAST.ETS ; _xlfn.FORECAST.ETS.CONFINT ; _xlfn.FORECAST.ETS.SEASONALITY ;
	// _xlfn.FORECAST.ETS.STAT ; _xlfn.FORECAST.LINEAR ; _xlfn.FORMULATEXT ; _xlfn.GAMMA ; _xlfn.GAMMA.DIST ;
	// _xlfn.GAMMA.INV ; _xlfn.GAMMALN.PRECISE ; _xlfn.GAUSS ; _xlfn.HYPGEOM.DIST ; _xlfn.IFNA ; _xlfn.IFS ;
	// _xlfn.IMCOSH ; _xlfn.IMCOT ; _xlfn.IMCSC ; _xlfn.IMCSCH ; _xlfn.IMSEC ; _xlfn.IMSECH ; _xlfn.IMSINH ;
	// _xlfn.IMTAN ; _xlfn.ISFORMULA ; ISO.CEILING ; _xlfn.ISOWEEKNUM ; _xlfn.LET ; _xlfn.LOGNORM.DIST ;
	// _xlfn.LOGNORM.INV ; _xlfn.MAXIFS ; _xlfn.MINIFS ; _xlfn.MODE.MULT ; _xlfn.MODE.SNGL ; _xlfn.MUNIT ;
	// _xlfn.NEGBINOM.DIST ; NETWORKDAYS.INTL ; _xlfn.NORM.DIST ; _xlfn.NORM.INV ; _xlfn.NORM.S.DIST ;
	// _xlfn.NORM.S.INV ; _xlfn.NUMBERVALUE ; _xlfn.PDURATION ; _xlfn.PERCENTILE.EXC ; _xlfn.PERCENTILE.INC ;
	// _xlfn.PERCENTRANK.EXC ; _xlfn.PERCENTRANK.INC ; _xlfn.PERMUTATIONA ; _xlfn.PHI ; _xlfn.POISSON.DIST ;
	// _xlfn.QUARTILE.EXC ; _xlfn.QUARTILE.INC ; _xlfn.QUERYSTRING ; _xlfn.RANDARRAY ; _xlfn.RANK.AVG ;
	// _xlfn.RANK.EQ ; _xlfn.RRI ; _xlfn.SEC ; _xlfn.SECH ; _xlfn.SEQUENCE ; _xlfn.SHEET ; _xlfn.SHEETS ;
	// _xlfn.SKEW.P ; _xlfn.SORTBY ; _xlfn.STDEV.P ; _xlfn.STDEV.S ; _xlfn.SWITCH ; _xlfn.T.DIST ; _xlfn.T.DIST.2T ;
	// _xlfn.T.DIST.RT ; _xlfn.T.INV ; _xlfn.T.INV.2T ; _xlfn.T.TEST ; _xlfn.TEXTJOIN ; _xlfn.UNICHAR ;
	// _xlfn.UNICODE ; _xlfn.UNIQUE ; _xlfn.VAR.P ; _xlfn.VAR.S ; _xlfn.WEBSERVICE ; _xlfn.WEIBULL.DIST ;
	// WORKDAY.INTL ; _xlfn.XLOOKUP ; _xlfn.XOR ; _xlfn.Z.TEST


	/*new funcions with _xlnf-prefix*/
	cFormulaFunctionGroup['TextAndData'] = cFormulaFunctionGroup['TextAndData'] || [];
	cFormulaFunctionGroup['TextAndData'].push(cDBCS);
	cFormulaFunctionGroup['Statistical'] = cFormulaFunctionGroup['Statistical'] || [];
	cFormulaFunctionGroup['Statistical'].push(cFORECAST_ETS, cFORECAST_ETS_CONFINT,
		cFORECAST_ETS_SEASONALITY, cFORECAST_ETS_STAT);

	cFormulaFunctionGroup['NotRealised'] = cFormulaFunctionGroup['NotRealised'] || [];
	cFormulaFunctionGroup['NotRealised'].push(cDBCS, cFORECAST_ETS,
		cFORECAST_ETS_CONFINT, cFORECAST_ETS_SEASONALITY, cFORECAST_ETS_STAT);

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cDBCS() {
	}

	cDBCS.prototype = Object.create(cBaseFunction.prototype);
	cDBCS.prototype.constructor = cDBCS;
	cDBCS.prototype.name = "DBCS";
	cDBCS.prototype.isXLFN = true;

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFILTERXML() {
	}

	cFILTERXML.prototype = Object.create(cBaseFunction.prototype);
	cFILTERXML.prototype.constructor = cFILTERXML;
	cFILTERXML.prototype.name = "FILTERXML";
	cFILTERXML.prototype.isXLFN = true;
	cFILTERXML.prototype.argumentsType = [argType.text, argType.text];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFORECAST_ETS() {
	}

	cFORECAST_ETS.prototype = Object.create(cBaseFunction.prototype);
	cFORECAST_ETS.prototype.constructor = cFORECAST_ETS;
	cFORECAST_ETS.prototype.name = "FORECAST.ETS";
	cFORECAST_ETS.prototype.isXLFN = true;
	cFORECAST_ETS.prototype.argumentsType = [argType.number, argType.reference, argType.reference, argType.number, argType.number, argType.number];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFORECAST_ETS_CONFINT() {
	}

	cFORECAST_ETS_CONFINT.prototype = Object.create(cBaseFunction.prototype);
	cFORECAST_ETS_CONFINT.prototype.constructor = cFORECAST_ETS_CONFINT;
	cFORECAST_ETS_CONFINT.prototype.name = "FORECAST.ETS.CONFINT";
	cFORECAST_ETS_CONFINT.prototype.isXLFN = true;
	cFORECAST_ETS_CONFINT.prototype.argumentsType = [argType.number, argType.reference, argType.reference, argType.number, argType.number,
		argType.number, argType.number];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFORECAST_ETS_SEASONALITY() {
	}

	cFORECAST_ETS_SEASONALITY.prototype = Object.create(cBaseFunction.prototype);
	cFORECAST_ETS_SEASONALITY.prototype.constructor = cFORECAST_ETS_SEASONALITY;
	cFORECAST_ETS_SEASONALITY.prototype.name = "FORECAST.ETS.SEASONALITY";
	cFORECAST_ETS_SEASONALITY.prototype.isXLFN = true;
	cFORECAST_ETS_SEASONALITY.prototype.argumentsType = [argType.reference, argType.reference, argType.number, argType.number];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFORECAST_ETS_STAT() {
	}

	cFORECAST_ETS_STAT.prototype = Object.create(cBaseFunction.prototype);
	cFORECAST_ETS_STAT.prototype.constructor = cFORECAST_ETS_STAT;
	cFORECAST_ETS_STAT.prototype.name = "FORECAST.ETS.STAT";
	cFORECAST_ETS_STAT.prototype.isXLFN = true;
	cFORECAST_ETS_STAT.prototype.argumentsType = [argType.reference, argType.reference, argType.number, argType.number,
		argType.number, argType.number];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cQUERYSTRING() {
	}

	cQUERYSTRING.prototype = Object.create(cBaseFunction.prototype);
	cQUERYSTRING.prototype.constructor = cQUERYSTRING;
	cQUERYSTRING.prototype.name = "QUERYSTRING";
	cQUERYSTRING.prototype.isXLFN = true;

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cWEBSERVICE() {
	}

	cWEBSERVICE.prototype = Object.create(cBaseFunction.prototype);
	cWEBSERVICE.prototype.constructor = cWEBSERVICE;
	cWEBSERVICE.prototype.name = "WEBSERVICE";
	cWEBSERVICE.prototype.isXLFN = true;
	cWEBSERVICE.prototype.argumentsType = [argType.text];
})(window);

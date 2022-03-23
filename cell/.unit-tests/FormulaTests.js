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

$( function () {

	var cDate = Asc.cDate;

    function toFixed( n ) {
        return n;//.toFixed( AscCommonExcel.cExcelSignificantDigits ) - 0;
    }

    function difBetween( a, b ) {
        return Math.abs( a - b ) < dif
    }

    function _getPMT( fZins, fZzr, fBw, fZw, nF ){
        var fRmz;
        if( fZins == 0.0 )
            fRmz = ( fBw + fZw ) / fZzr;
        else{
            var	fTerm = Math.pow( 1.0 + fZins, fZzr );
            if( nF > 0 )
                fRmz = ( fZw * fZins / ( fTerm - 1.0 ) + fBw * fZins / ( 1.0 - 1.0 / fTerm ) ) / ( 1.0 + fZins );
            else
                fRmz = fZw * fZins / ( fTerm - 1.0 ) + fBw * fZins / ( 1.0 - 1.0 / fTerm );
        }

        return -fRmz;
    }

    function _getFV( fZins, fZzr, fRmz, fBw, nF ){
        var fZw;
        if( fZins == 0.0 )
            fZw = fBw + fRmz * fZzr;
        else{
            var fTerm = Math.pow( 1.0 + fZins, fZzr );
            if( nF > 0 )
                fZw = fBw * fTerm + fRmz * ( 1.0 + fZins ) * ( fTerm - 1.0 ) / fZins;
            else
                fZw = fBw * fTerm + fRmz * ( fTerm - 1.0 ) / fZins;
        }

        return -fZw;
    }

    function _getDDB( cost, salvage, life, period, factor ) {
        var ddb, ipmt, oldCost, newCost;
        ipmt = factor / life;
        if ( ipmt >= 1 ) {
            ipmt = 1;
            if ( period == 1 )
                oldCost = cost;
            else
                oldCost = 0;
        }
        else
            oldCost = cost * Math.pow( 1 - ipmt, period - 1 );
        newCost = cost * Math.pow( 1 - ipmt, period );

        if ( newCost < salvage )
            ddb = oldCost - salvage;
        else
            ddb = oldCost - newCost;
        if ( ddb < 0 )
            ddb = 0;
        return ddb;
    }

    function _getIPMT(rate, per, pv, type, pmt) {
        var ipmt;

        if ( per == 1 ) {
            if ( type > 0 )
                ipmt = 0;
            else
                ipmt = -pv;
        }
        else {
            if ( type > 0 )
                ipmt = _getFV( rate, per - 2, pmt, pv, 1 ) - pmt;
            else
                ipmt = _getFV( rate, per - 1, pmt, pv, 0 );
        }
        return ipmt * rate
    }

    function _diffDate(d1, d2, mode){
        var date1 = d1.getDate(),
            month1 = d1.getMonth(),
            year1 = d1.getFullYear(),
            date2 = d2.getDate(),
            month2 = d2.getMonth(),
            year2 = d2.getFullYear();

        switch ( mode ) {
            case 0:
                return Math.abs( GetDiffDate360( date1, month1, year1, date2, month2, year2, true ) );
            case 1:
                var yc = Math.abs( year2 - year1 ),
                    sd = year1 > year2 ? d2 : d1,
                    yearAverage = sd.isLeapYear() ? 366 : 365, dayDiff = Math.abs( d2 - d1 );
                for ( var i = 0; i < yc; i++ ) {
                    sd.addYears( 1 );
                    yearAverage += sd.isLeapYear() ? 366 : 365;
                }
                yearAverage /= (yc + 1);
                dayDiff /= c_msPerDay;
                return dayDiff;
            case 2:
                var dayDiff = Math.abs( d2 - d1 );
                dayDiff /= c_msPerDay;
                return dayDiff;
            case 3:
                var dayDiff = Math.abs( d2 - d1 );
                dayDiff /= c_msPerDay;
                return dayDiff;
            case 4:
                return Math.abs( GetDiffDate360( date1, month1, year1, date2, month2, year2, false ) );
            default:
                return "#NUM!";
        }
    }

    function _yearFrac(d1, d2, mode) {
        var date1 = d1.getDate(),
            month1 = d1.getMonth()+1,
            year1 = d1.getFullYear(),
            date2 = d2.getDate(),
            month2 = d2.getMonth()+1,
            year2 = d2.getFullYear();

        switch ( mode ) {
            case 0:
                return Math.abs( GetDiffDate360( date1, month1, year1, date2, month2, year2, true ) ) / 360;
            case 1:
                var yc = /*Math.abs*/( year2 - year1 ),
                    sd = year1 > year2 ? new cDate(d2) : new cDate(d1),
                    yearAverage = sd.isLeapYear() ? 366 : 365, dayDiff = /*Math.abs*/( d2 - d1 );
                for ( var i = 0; i < yc; i++ ) {
                    sd.addYears( 1 );
                    yearAverage += sd.isLeapYear() ? 366 : 365;
                }
                yearAverage /= (yc + 1);
                dayDiff /= (yearAverage * c_msPerDay);
                return dayDiff;
            case 2:
                var dayDiff = Math.abs( d2 - d1 );
                dayDiff /= (360 * c_msPerDay);
                return dayDiff;
            case 3:
                var dayDiff = Math.abs( d2 - d1 );
                dayDiff /= (365 * c_msPerDay);
                return dayDiff;
            case 4:
                return Math.abs( GetDiffDate360( date1, month1, year1, date2, month2, year2, false ) ) / 360;
            default:
                return "#NUM!";
        }
    }

    function _lcl_GetCouppcd(settl, matur, freq){
        matur.setFullYear( settl.getFullYear() );
        if( matur < settl )
            matur.addYears( 1 );
        while( matur > settl ){
            matur.addMonths( -12 / freq );
        }
    }

    function _lcl_GetCoupncd( settl, matur, freq ){
        matur.setFullYear( settl.getFullYear() );
        if( matur > settl )
            matur.addYears( -1 );
        while( matur <= settl ){
            matur.addMonths( 12 / freq );
        }
    }

    function _getcoupdaybs( settl, matur, frequency, basis ) {
        _lcl_GetCouppcd( settl, matur, frequency );
        return _diffDate( settl, matur, basis );
    }

    function _getcoupdays( settl, matur, frequency, basis ) {
        _lcl_GetCouppcd( settl, matur, frequency );
        var n = new cDate( matur );
        n.addMonths( 12 / frequency );
        return _diffDate( matur, n, basis );
    }

    function _getdiffdate( d1,d2, nMode ){
        var bNeg = d1 > d2;

        if( bNeg )
        {
            var n = d2;
            d2 = d1;
            d1 = n;
        }

        var nRet,pOptDaysIn1stYear;

        var nD1 = d1.getDate(),
            nM1 = d1.getMonth(),
            nY1  = d1.getFullYear(),
            nD2 = d2.getDate(),
            nM2 = d2.getMonth(),
            nY2 = d2.getFullYear();

        switch( nMode )
        {
            case 0:			// 0=USA (NASD) 30/360
            case 4:			// 4=Europe 30/360
            {
                var bLeap = d1.isLeapYear();
                var nDays, nMonths/*, nYears*/;

                nMonths = nM2 - nM1;
                nDays = nD2 - nD1;

                nMonths += ( nY2 - nY1 ) * 12;

                nRet = nMonths * 30 + nDays;
                if( nMode == 0 && nM1 == 2 && nM2 != 2 && nY1 == nY2 )
                    nRet -= bLeap? 1 : 2;

                pOptDaysIn1stYear = 360;
            }
                break;
            case 1:			// 1=exact/exact
                pOptDaysIn1stYear = d1.isLeapYear() ? 366 : 365;
                nRet = d2 - d1;
                break;
            case 2:			// 2=exact/360
                nRet = d2 - d1;
                pOptDaysIn1stYear = 360;
                break;
            case 3:			//3=exact/365
                nRet = d2 - d1;
                pOptDaysIn1stYear = 365;
                break;
        }

        return (bNeg ? -nRet : nRet) / c_msPerDay / pOptDaysIn1stYear;
    }

    function _getprice( nSettle, nMat, fRate, fYield, fRedemp, nFreq, nBase ){

        var fdays = AscCommonExcel.getcoupdays( new cDate(nSettle), new cDate(nMat), nFreq, nBase ),
            fdaybs = AscCommonExcel.getcoupdaybs( new cDate(nSettle), new cDate(nMat), nFreq, nBase ),
            fnum = AscCommonExcel.getcoupnum( new cDate(nSettle), (nMat), nFreq, nBase ),
            fdaysnc = ( fdays - fdaybs ) / fdays,
            fT1 = 100 * fRate / nFreq,
            fT2 = 1 + fYield / nFreq,
            res = fRedemp / ( Math.pow( 1 + fYield / nFreq, fnum - 1 + fdaysnc ) );

        /*var fRet = fRedemp / ( Math.pow( 1.0 + fYield / nFreq, fnum - 1.0 + fdaysnc ) );
        fRet -= 100.0 * fRate / nFreq * fdaybs / fdays;

        var fT1 = 100.0 * fRate / nFreq;
        var fT2 = 1.0 + fYield / nFreq;

        for( var fK = 0.0 ; fK < fnum ; fK++ ){
            fRet += fT1 / Math.pow( fT2, fK + fdaysnc );
        }

        return fRet;*/

        if( fnum == 1){
            return (fRedemp + fT1) / (1 + fdaysnc * fYield / nFreq) - 100 * fRate / nFreq * fdaybs / fdays;
        }

        res -= 100 * fRate / nFreq * fdaybs / fdays;

        for ( var i = 0; i < fnum; i++ ) {
            res += fT1 / Math.pow( fT2, i + fdaysnc );
        }

        return res;
    }

    function _getYield( nSettle, nMat, fCoup, fPrice, fRedemp, nFreq, nBase ){
        var fRate = fCoup, fPriceN = 0.0, fYield1 = 0.0, fYield2 = 1.0;
        var fPrice1 = _getprice( nSettle, nMat, fRate, fYield1, fRedemp, nFreq, nBase );
        var fPrice2 = _getprice( nSettle, nMat, fRate, fYield2, fRedemp, nFreq, nBase );
        var fYieldN = ( fYield2 - fYield1 ) * 0.5;

        for( var nIter = 0 ; nIter < 100 && fPriceN != fPrice ; nIter++ )
        {
            fPriceN = _getprice( nSettle, nMat, fRate, fYieldN, fRedemp, nFreq, nBase );

            if( fPrice == fPrice1 )
                return fYield1;
            else if( fPrice == fPrice2 )
                return fYield2;
            else if( fPrice == fPriceN )
                return fYieldN;
            else if( fPrice < fPrice2 )
            {
                fYield2 *= 2.0;
                fPrice2 = _getprice( nSettle, nMat, fRate, fYield2, fRedemp, nFreq, nBase );

                fYieldN = ( fYield2 - fYield1 ) * 0.5;
            }
            else
            {
                if( fPrice < fPriceN )
                {
                    fYield1 = fYieldN;
                    fPrice1 = fPriceN;
                }
                else
                {
                    fYield2 = fYieldN;
                    fPrice2 = fPriceN;
                }

                fYieldN = fYield2 - ( fYield2 - fYield1 ) * ( ( fPrice - fPrice2 ) / ( fPrice1 - fPrice2 ) );
            }
        }

        if( Math.abs( fPrice - fPriceN ) > fPrice / 100.0 )
            return "#NUM!";		// result not precise enough

        return fYieldN;
    }

    function _getyieldmat( nSettle, nMat, nIssue, fRate, fPrice, nBase ){

        var fIssMat = _yearFrac( nIssue, nMat, nBase );
        var fIssSet = _yearFrac( nIssue, nSettle, nBase );
        var fSetMat = _yearFrac( nSettle, nMat, nBase );

        var y = 1.0 + fIssMat * fRate;
        y /= fPrice / 100.0 + fIssSet * fRate;
        y--;
        y /= fSetMat;

        return y;

    }

    function _coupnum( settlement, maturity, frequency, basis ) {

        basis = ( basis !== undefined ? basis : 0 );

        var n = new cDate(maturity);
        _lcl_GetCouppcd( settlement, n, frequency );
        var nMonths = (maturity.getFullYear() - n.getFullYear()) * 12 + maturity.getMonth() - n.getMonth();
        return nMonths * frequency / 12 ;

    }

    function _duration( settlement, maturity, coupon, yld, frequency, basis ){
        var dbc = AscCommonExcel.getcoupdaybs(new cDate( settlement ),new cDate( maturity ),frequency,basis),
            coupD = AscCommonExcel.getcoupdays(new cDate( settlement ),new cDate( maturity ),frequency,basis),
            numCoup = AscCommonExcel.getcoupnum(new cDate( settlement ),new cDate( maturity ),frequency);

        if ( settlement >= maturity || basis < 0 || basis > 4 || ( frequency != 1 && frequency != 2 && frequency != 4 ) || yld < 0 || coupon < 0 ){
            return "#NUM!";
        }

        var duration = 0, p = 0;

        var dsc = coupD - dbc;
        var diff = dsc / coupD - 1;
        yld = yld / frequency + 1;


        coupon *= 100/frequency;

        for(var index = 1; index <= numCoup; index++ ){
            var di = index + diff;

            var yldPOW = Math.pow( yld, di);

            duration += di * coupon / yldPOW;

            p += coupon / yldPOW;
        }

        duration += (diff + numCoup) * 100 / Math.pow( yld, diff + numCoup);
        p += 100 / Math.pow( yld, diff + numCoup);

        return duration / p / frequency ;
    }

	function numDivFact(num, fact){
		var res = num / Math.fact(fact);
		res = res.toString();
		return res;
	}

	function testArrayFormula(assert, func, dNotSupportAreaArg) {

		var getValue = function(ref) {
			oParser = new parserFormula( func + "(" + ref + ")", "A2", ws );
			assert.ok( oParser.parse() );
			return oParser.calculate().getValue();
		};

		//***array-formula***
		ws.getRange2( "A100" ).setValue( "1" );
		ws.getRange2( "B100" ).setValue( "3" );
		ws.getRange2( "C100" ).setValue( "-4" );
		ws.getRange2( "A101" ).setValue( "2" );
		ws.getRange2( "B101" ).setValue( "4" );
		ws.getRange2( "C101" ).setValue( "5" );


		oParser = new parserFormula( func + "(A100:C101)", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:H107").bbox);
		assert.ok( oParser.parse() );
		var array = oParser.calculate();
		if(AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual( array.getElementRowCol(0,0).getValue(), getValue("A100"));
			assert.strictEqual( array.getElementRowCol(0,1).getValue(), getValue("B100"));
			assert.strictEqual( array.getElementRowCol(0,2).getValue(), getValue("C100"));
			assert.strictEqual( array.getElementRowCol(1,0).getValue(), getValue("A101"));
			assert.strictEqual( array.getElementRowCol(1,1).getValue(), getValue("B101"));
			assert.strictEqual( array.getElementRowCol(1,2).getValue(), getValue("C101"));
		} else {
			if(!dNotSupportAreaArg) {
				assert.strictEqual( false, true);
			}
			consoleLog("func: " + func + " don't return area array");
		}

		oParser = new parserFormula( func + "({1,2,-3})", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:H107").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		assert.strictEqual( array.getElementRowCol(0,0).getValue(), getValue(1));
		assert.strictEqual( array.getElementRowCol(0,1).getValue(), getValue(2));
		assert.strictEqual( array.getElementRowCol(0,2).getValue(), getValue(-3));
	}

	//returnOnlyValue - те функции, на вход которых всегда должны подаваться массивы и которые возвращают единственное значение
	function testArrayFormula2(assert, func, minArgCount, maxArgCount, dNotSupportAreaArg, returnOnlyValue) {

		var getValue = function(ref, countArg) {
			var argStr = "(";
			for(var j = 1; j <= countArg; j++) {
				argStr += ref;
				if(i !== j) {
					argStr += ",";
				} else {
					argStr += ")";
				}
			}
			oParser = new parserFormula( func + argStr, "A2", ws );
			assert.ok( oParser.parse() );
			return oParser.calculate().getValue();
		};


		//***array-formula***
		ws.getRange2( "A100" ).setValue( "1" );
		ws.getRange2( "B100" ).setValue( "3" );
		ws.getRange2( "C100" ).setValue( "-4" );
		ws.getRange2( "A101" ).setValue( "2" );
		ws.getRange2( "B101" ).setValue( "4" );
		ws.getRange2( "C101" ).setValue( "5" );

		//формируем массив значений
		var randomArray = [];
		var randomStrArray = "{";
		var maxArg = 4;
		for(var i = 1; i <= maxArg; i++) {
			var randVal = Math.random();
			randomArray.push(randVal);
			randomStrArray += randVal;
			if(i !== maxArg) {
				randomStrArray += ",";
			} else {
				randomStrArray += "}";
			}
		}

		for(var i = minArgCount; i <= maxArgCount; i++) {
			var argStrArr = "(";
			var randomArgStrArr = "(";
			for(var j = 1; j <= i; j++) {
				argStrArr += "A100:C101";
				randomArgStrArr += randomStrArray;
				if(i !== j) {
					argStrArr += ",";
					randomArgStrArr += ",";
				} else {
					argStrArr += ")";
					randomArgStrArr += ")";
				}
			}

			oParser = new parserFormula( func + argStrArr, "A1", ws );
			oParser.setArrayFormulaRef(ws.getRange2("E106:H107").bbox);
			assert.ok( oParser.parse() );
			var array = oParser.calculate();
			if(AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual( array.getElementRowCol(0,0).getValue(), getValue("A100", i));
				assert.strictEqual( array.getElementRowCol(0,1).getValue(), getValue("B100", i));
				assert.strictEqual( array.getElementRowCol(0,2).getValue(), getValue("C100", i));
				assert.strictEqual( array.getElementRowCol(1,0).getValue(), getValue("A101", i));
				assert.strictEqual( array.getElementRowCol(1,1).getValue(), getValue("B101", i));
				assert.strictEqual( array.getElementRowCol(1,2).getValue(), getValue("C101", i));
			} else {
				if(!(dNotSupportAreaArg || returnOnlyValue)) {
					assert.strictEqual( false, true);
				}
				consoleLog("func: " + func + " don't return area array");
			}

			oParser = new parserFormula( func + randomArgStrArr, "A1", ws );
			oParser.setArrayFormulaRef(ws.getRange2("E106:H107").bbox);
			assert.ok( oParser.parse() );
			array = oParser.calculate();
			if(AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual( array.getElementRowCol(0,0).getValue(), getValue(randomArray[0], i));
				assert.strictEqual( array.getElementRowCol(0,1).getValue(), getValue(randomArray[1], i));
				assert.strictEqual( array.getElementRowCol(0,2).getValue(), getValue(randomArray[2], i));
			} else {
				if(!returnOnlyValue) {
					assert.strictEqual( false, true);
				}
				consoleLog("func: " + func + " don't return array");
			}
		}
	}

	function testArrayFormulaEqualsValues(assert, str, formula, isNotLowerCase) {
		//***array-formula***
		ws.getRange2( "A1" ).setValue( "1" );
		ws.getRange2( "B1" ).setValue( "3.123" );
		ws.getRange2( "C1" ).setValue( "-4" );
		ws.getRange2( "A2" ).setValue( "2" );
		ws.getRange2( "B2" ).setValue( "4" );
		ws.getRange2( "C2" ).setValue( "5" );

		oParser = new parserFormula( formula, "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E6:H8").bbox);
		assert.ok( oParser.parse() );
		var array = oParser.calculate();

		var splitStr = str.split(";");

		for(var i = 0; i < splitStr.length; i++) {
			var subSplitStr = splitStr[i].split(",");
			for(var j = 0; j < subSplitStr.length; j++) {
				var valMs = subSplitStr[j];
				var element;
				if(array.getElementRowCol) {
					var row = 1 === array.array.length ? 0 : i;
					var col = 1 === array.array[0].length ? 0 : j;
					if(array.array[row] && array.array[row][col]) {
						element = array.getElementRowCol(row, col);
					} else {
						element = new window['AscCommonExcel'].cError(window['AscCommonExcel'].cErrorType.not_available);
					}
				} else {
					element = array;
				}
				var ourVal = element && undefined != element.value ? element.value.toString() : "#N/A";
				if(!isNotLowerCase) {
					valMs = valMs.toLowerCase();
					ourVal = ourVal.toLowerCase();
				}
				assert.strictEqual(valMs, ourVal, "formula: " + formula + " i: " + i + " j: " + j)
			}
		}
	}

	function _getValue(from, row, col) {
		var res;
		if (from.type === AscCommonExcel.cElementType.array) {
			res = from.getElementRowCol(row !== undefined ? row : 0, col !== undefined ? col : 0).getValue();
		} else if (from.type === AscCommonExcel.cElementType.cellsRange || from.type === AscCommonExcel.cElementType.cellsRange3D) {
			res = from.getValueByRowCol(row !== undefined ? row : 0, col !== undefined ? col : 0).getValue();
		} else if (from.type === AscCommonExcel.cElementType.cell || from.type === AscCommonExcel.cElementType.cell3D) {
			res = from.getValue().getValue();
		} else {
			res = from.getValue();
		}
		return res;
	}

	function consoleLog(val) {
		//console.log(val);
	}

	var newFormulaParser = false;

    var c_msPerDay = AscCommonExcel.c_msPerDay;
    var parserFormula = AscCommonExcel.parserFormula;
    var GetDiffDate360 = AscCommonExcel.GetDiffDate360;
    var fSortAscending = AscCommon.fSortAscending;
    var g_oIdCounter = AscCommon.g_oIdCounter;

    var oParser, wb, ws, dif = 1e-9, sData = AscCommon.getEmpty(), tmp;
    if ( AscCommon.c_oSerFormat.Signature === sData.substring( 0, AscCommon.c_oSerFormat.Signature.length ) ) {
        wb = new AscCommonExcel.Workbook( new AscCommonExcel.asc_CHandlersList(), {wb:{getWorksheet:function(){}}} );
        AscCommon.History.init(wb);

        AscCommon.g_oTableId.init();
        if ( this.User )
            g_oIdCounter.Set_UserId(this.User.asc_getId());

        AscCommonExcel.g_oUndoRedoCell = new AscCommonExcel.UndoRedoCell(wb);
        AscCommonExcel.g_oUndoRedoWorksheet = new AscCommonExcel.UndoRedoWoorksheet(wb);
        AscCommonExcel.g_oUndoRedoWorkbook = new AscCommonExcel.UndoRedoWorkbook(wb);
        AscCommonExcel.g_oUndoRedoCol = new AscCommonExcel.UndoRedoRowCol(wb, false);
        AscCommonExcel.g_oUndoRedoRow = new AscCommonExcel.UndoRedoRowCol(wb, true);
        AscCommonExcel.g_oUndoRedoComment = new AscCommonExcel.UndoRedoComment(wb);
        AscCommonExcel.g_oUndoRedoAutoFilters = new AscCommonExcel.UndoRedoAutoFilters(wb);
		AscCommonExcel.g_DefNameWorksheet = new AscCommonExcel.Worksheet(wb, -1);
        g_oIdCounter.Set_Load(false);

        var oBinaryFileReader = new AscCommonExcel.BinaryFileReader();
        oBinaryFileReader.Read( sData, wb );
        ws = wb.getWorksheet( wb.getActive() );
        AscCommonExcel.getFormulasInfo();
    }

	wb.dependencyFormulas.lockRecal();

	QUnit.module("Formula");
	QUnit.test("Test: \"ABS\"", function (assert) {

		ws.getRange2( "A22" ).setValue( "-4" );

		oParser = new parserFormula( "ABS(2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );
		oParser = new parserFormula( "ABS(-2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );
		oParser = new parserFormula( "ABS(A22)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );

		testArrayFormula(assert, "ABS");
	} );

    QUnit.test("Test: \"Absolute reference\"", function (assert) {

        ws.getRange2( "A7" ).setValue( "1" );
        ws.getRange2( "A8" ).setValue( "2" );
        ws.getRange2( "A9" ).setValue( "3" );
        oParser = new parserFormula( 'A$7+A8', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3 );

        oParser = new parserFormula( 'A$7+A$8', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3 );

        oParser = new parserFormula( '$A$7+$A$8', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3 );

        oParser = new parserFormula( 'SUM($A$7:$A$9)', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 6 );
    } );

	QUnit.test("Test: \"Asc\"", function (assert) {
		oParser = new parserFormula( 'ASC("ｔｅＳｔ")', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "teSt" );

		oParser = new parserFormula( 'ASC("デジタル")', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "デジタル" );

		oParser = new parserFormula( 'ASC("￯")', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "" );
	} );

	QUnit.test("Test: \"Cross\"", function (assert) {

		ws.getRange2( "A7" ).setValue( "1" );
		ws.getRange2( "A8" ).setValue( "2" );
		ws.getRange2( "A9" ).setValue( "3" );
		oParser = new parserFormula( 'A7:A9', null, ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().cross(new Asc.Range(0, 5, 0, 5), ws.getId()).getValue(), "#VALUE!" );
		assert.strictEqual( oParser.calculate().cross(new Asc.Range(0, 6, 0, 6), ws.getId()).getValue(), 1 );
		assert.strictEqual( oParser.calculate().cross(new Asc.Range(0, 7, 0, 7), ws.getId()).getValue(), 2 );
		assert.strictEqual( oParser.calculate().cross(new Asc.Range(0, 8, 0, 8), ws.getId()).getValue(), 3 );
		assert.strictEqual( oParser.calculate().cross(new Asc.Range(0, 9, 0, 9), ws.getId()).getValue(), "#VALUE!" );

	} );

	QUnit.test("Test: \"Defined names cycle\"", function (assert) {

		var newNameQ = new Asc.asc_CDefName("q", "SUM('"+ws.getName()+"'!A2)");
		wb.editDefinesNames(null, newNameQ);
		ws.getRange2( "Q1" ).setValue( "=q" );
		ws.getRange2( "Q2" ).setValue( "=q" );
		ws.getRange2( "Q3" ).setValue( "1" );
		assert.strictEqual( ws.getRange2( "Q1" ).getValueWithFormat(), "1" );
		assert.strictEqual( ws.getRange2( "Q2" ).getValueWithFormat(), "1" );

		var newNameW = new Asc.asc_CDefName("w", "'"+ws.getName()+"'!A1");
		wb.editDefinesNames(null, newNameW);
		ws.getRange2( "Q4" ).setValue( "=w" );
		assert.strictEqual( ws.getRange2( "Q4" ).getValueWithFormat(), "#REF!" );
		//clean up
		ws.getRange2( "Q1:Q4" ).cleanAll();
		wb.delDefinesNames(newNameW);
		wb.delDefinesNames(newNameQ);
	});

	QUnit.test("Test: \"Parse intersection\"", function (assert) {

		ws.getRange2( "A7" ).setValue( "1" );
		ws.getRange2( "A8" ).setValue( "2" );
		ws.getRange2( "A9" ).setValue( "3" );
		oParser = new parserFormula( '1     +    (    A7   +A8   )   *   2', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.assemble(), "1+(A7+A8)*2" );
		assert.strictEqual( oParser.calculate().getValue(), 7 );

		oParser = new parserFormula( 'sum                    A1:A5', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.assemble(), "sum A1:A5" );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( 'sum(   A1:A5    ,        B1:B5     )     ', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.assemble(), "SUM(A1:A5,B1:B5)" );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( 'sum(   A1:A5    ,        B1:B5  , "    3 , 14 15 92 6 "   )     ', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.assemble(), 'SUM(A1:A5,B1:B5,"    3 , 14 15 92 6 ")' );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

	} );



    QUnit.test("Test: \"Arithmetical operations\"", function (assert) {
        oParser = new parserFormula( '1+3', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4 );

        oParser = new parserFormula( '(1+2)*4+3', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), (1 + 2) * 4 + 3 );

        oParser = new parserFormula( '2^52', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), Math.pow( 2, 52 ) );

        oParser = new parserFormula( '-10', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -10 );

        oParser = new parserFormula( '-10*2', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -20 );

        oParser = new parserFormula( '-10+10', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 0 );

        oParser = new parserFormula( '12%', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 0.12 );

        oParser = new parserFormula( "2<>\"3\"", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "TRUE", "2<>\"3\"" );

        oParser = new parserFormula( "2=\"3\"", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "FALSE", "2=\"3\"" );

        oParser = new parserFormula( "2>\"3\"", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "FALSE", "2>\"3\"" );

        oParser = new parserFormula( "\"f\">\"3\"", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

        oParser = new parserFormula( "\"f\"<\"3\"", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( "FALSE", oParser.calculate().getValue(), "FALSE" );

        oParser = new parserFormula( "FALSE>=FALSE", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

        oParser = new parserFormula( "\"TRUE\"&\"TRUE\"", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "TRUETRUE" );

        oParser = new parserFormula( "10*\"\"", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        oParser = new parserFormula( "-TRUE", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -1 );


		ws.getRange2( "M106" ).setValue( "1" );
		ws.getRange2( "M107" ).setValue( "2" );
		ws.getRange2( "M108" ).setValue( "2" );
		ws.getRange2( "M109" ).setValue( "4" );
		ws.getRange2( "M110" ).setValue( "5" );
		ws.getRange2( "M111" ).setValue( "-23" );
		ws.getRange2( "M112" ).setValue( "6" );
		ws.getRange2( "M113" ).setValue( "5" );

		ws.getRange2( "N106" ).setValue( "1" );
		ws.getRange2( "N107" ).setValue( "" );
		ws.getRange2( "N108" ).setValue( "" );
		ws.getRange2( "N109" ).setValue( "3" );
		ws.getRange2( "N110" ).setValue( "" );
		ws.getRange2( "N111" ).setValue( "2" );
		ws.getRange2( "N112" ).setValue( "" );
		ws.getRange2( "N113" ).setValue( "3" );

		ws.getRange2( "O106" ).setValue( "1" );
		ws.getRange2( "O107" ).setValue( "3" );
		ws.getRange2( "O108" ).setValue( "2" );
		ws.getRange2( "O109" ).setValue( "12" );
		ws.getRange2( "O110" ).setValue( "3" );
		ws.getRange2( "O111" ).setValue( "4" );
		ws.getRange2( "O112" ).setValue( "3" );
		ws.getRange2( "O113" ).setValue( "2" );

		ws.getRange2( "P106" ).setValue( "3" );
		ws.getRange2( "P107" ).setValue( "4" );
		ws.getRange2( "P108" ).setValue( "5" );
		ws.getRange2( "P109" ).setValue( "1" );
		ws.getRange2( "P110" ).setValue( "23" );
		ws.getRange2( "P111" ).setValue( "4" );
		ws.getRange2( "P112" ).setValue( "3" );
		ws.getRange2( "P113" ).setValue( "1" );

		oParser = new parserFormula( "M106:P113+M106:P113", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok( oParser.parse() );
		var array = oParser.calculate();
		if(AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual( array.getElementRowCol(0,0).getValue(), 2);
			assert.strictEqual( array.getElementRowCol(1,0).getValue(), 4);
			assert.strictEqual( array.getElementRowCol(2,0).getValue(), 4);
			assert.strictEqual( array.getElementRowCol(3,0).getValue(), 8);
			assert.strictEqual( array.getElementRowCol(4,0).getValue(), 10);
			assert.strictEqual( array.getElementRowCol(5,0).getValue(), -46);
			assert.strictEqual( array.getElementRowCol(6,0).getValue(), 12);
			assert.strictEqual( array.getElementRowCol(7,0).getValue(), 10);

			assert.strictEqual( array.getElementRowCol(0,1).getValue(), 2);
			assert.strictEqual( array.getElementRowCol(1,1).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(2,1).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(3,1).getValue(), 6);
			assert.strictEqual( array.getElementRowCol(4,1).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(5,1).getValue(), 4);
			assert.strictEqual( array.getElementRowCol(6,1).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(7,1).getValue(), 6);

			assert.strictEqual( array.getElementRowCol(0,2).getValue(), 2);
			assert.strictEqual( array.getElementRowCol(1,2).getValue(), 6);
			assert.strictEqual( array.getElementRowCol(2,2).getValue(), 4);
			assert.strictEqual( array.getElementRowCol(3,2).getValue(), 24);
			assert.strictEqual( array.getElementRowCol(4,2).getValue(), 6);
			assert.strictEqual( array.getElementRowCol(5,2).getValue(), 8);
			assert.strictEqual( array.getElementRowCol(6,2).getValue(), 6);
			assert.strictEqual( array.getElementRowCol(7,2).getValue(), 4);

			assert.strictEqual( array.getElementRowCol(0,3).getValue(), 6);
			assert.strictEqual( array.getElementRowCol(1,3).getValue(), 8);
			assert.strictEqual( array.getElementRowCol(2,3).getValue(), 10);
			assert.strictEqual( array.getElementRowCol(3,3).getValue(), 2);
			assert.strictEqual( array.getElementRowCol(4,3).getValue(), 46);
			assert.strictEqual( array.getElementRowCol(5,3).getValue(), 8);
			assert.strictEqual( array.getElementRowCol(6,3).getValue(), 6);
			assert.strictEqual( array.getElementRowCol(7,3).getValue(), 2);

		}

		oParser = new parserFormula( "M106:P113*M106:P113", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		if(AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual( array.getElementRowCol(0,0).getValue(), 1);
			assert.strictEqual( array.getElementRowCol(1,0).getValue(), 4);
			assert.strictEqual( array.getElementRowCol(2,0).getValue(), 4);
			assert.strictEqual( array.getElementRowCol(3,0).getValue(), 16);
			assert.strictEqual( array.getElementRowCol(4,0).getValue(), 25);
			assert.strictEqual( array.getElementRowCol(5,0).getValue(), 529);
			assert.strictEqual( array.getElementRowCol(6,0).getValue(), 36);
			assert.strictEqual( array.getElementRowCol(7,0).getValue(), 25);

			assert.strictEqual( array.getElementRowCol(0,1).getValue(), 1);
			assert.strictEqual( array.getElementRowCol(1,1).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(2,1).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(3,1).getValue(), 9);
			assert.strictEqual( array.getElementRowCol(4,1).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(5,1).getValue(), 4);
			assert.strictEqual( array.getElementRowCol(6,1).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(7,1).getValue(), 9);

			assert.strictEqual( array.getElementRowCol(0,2).getValue(), 1);
			assert.strictEqual( array.getElementRowCol(1,2).getValue(), 9);
			assert.strictEqual( array.getElementRowCol(2,2).getValue(), 4);
			assert.strictEqual( array.getElementRowCol(3,2).getValue(), 144);
			assert.strictEqual( array.getElementRowCol(4,2).getValue(), 9);
			assert.strictEqual( array.getElementRowCol(5,2).getValue(), 16);
			assert.strictEqual( array.getElementRowCol(6,2).getValue(), 9);
			assert.strictEqual( array.getElementRowCol(7,2).getValue(), 4);

			assert.strictEqual( array.getElementRowCol(0,3).getValue(), 9);
			assert.strictEqual( array.getElementRowCol(1,3).getValue(), 16);
			assert.strictEqual( array.getElementRowCol(2,3).getValue(), 25);
			assert.strictEqual( array.getElementRowCol(3,3).getValue(), 1);
			assert.strictEqual( array.getElementRowCol(4,3).getValue(), 529);
			assert.strictEqual( array.getElementRowCol(5,3).getValue(), 16);
			assert.strictEqual( array.getElementRowCol(6,3).getValue(), 9);
			assert.strictEqual( array.getElementRowCol(7,3).getValue(), 1);
		}

		oParser = new parserFormula( "M106:P113-M106:P113", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		if(AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual( array.getElementRowCol(0,0).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(1,0).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(2,0).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(3,0).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(4,0).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(5,0).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(6,0).getValue(), 0);
			assert.strictEqual( array.getElementRowCol(7,0).getValue(), 0);
		}

		oParser = new parserFormula( "M106:P113=M106:P113", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		if(AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual( array.getElementRowCol(0,0).getValue(), "TRUE");
			assert.strictEqual( array.getElementRowCol(1,0).getValue(), "TRUE");
			assert.strictEqual( array.getElementRowCol(2,0).getValue(), "TRUE");
			assert.strictEqual( array.getElementRowCol(3,0).getValue(), "TRUE");
			assert.strictEqual( array.getElementRowCol(4,0).getValue(), "TRUE");
			assert.strictEqual( array.getElementRowCol(5,0).getValue(), "TRUE");
			assert.strictEqual( array.getElementRowCol(6,0).getValue(), "TRUE");
			assert.strictEqual( array.getElementRowCol(7,0).getValue(), "TRUE");
		}

		oParser = new parserFormula( "M106:P113/M106:P113", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		if(AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual( array.getElementRowCol(0,0).getValue(), 1);
			assert.strictEqual( array.getElementRowCol(1,0).getValue(), 1);
			assert.strictEqual( array.getElementRowCol(2,1).getValue(), "#DIV/0!");
		}

		oParser = new parserFormula( "M106:P113<>M106:P113", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		if(AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual( array.getElementRowCol(0,0).getValue(), "FALSE");
			assert.strictEqual( array.getElementRowCol(1,0).getValue(), "FALSE");
			assert.strictEqual( array.getElementRowCol(2,0).getValue(), "FALSE");
		}

		oParser = new parserFormula( "M106:P113>M106:P113", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		if(AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual( array.getElementRowCol(0,0).getValue(), "FALSE");
			assert.strictEqual( array.getElementRowCol(1,0).getValue(), "FALSE");
			assert.strictEqual( array.getElementRowCol(2,0).getValue(), "FALSE");
		}

		oParser = new parserFormula( "M106:P113<M106:P113", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		if(AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual( array.getElementRowCol(0,0).getValue(), "FALSE");
			assert.strictEqual( array.getElementRowCol(1,0).getValue(), "FALSE");
			assert.strictEqual( array.getElementRowCol(2,0).getValue(), "FALSE");
		}

		oParser = new parserFormula( "M106:P113>=M106:P113", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		if(AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual( array.getElementRowCol(0,0).getValue(), "TRUE");
			assert.strictEqual( array.getElementRowCol(1,0).getValue(), "TRUE");
			assert.strictEqual( array.getElementRowCol(2,0).getValue(), "TRUE");
		}

		oParser = new parserFormula( "SUM(M:P*M:P)", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		assert.strictEqual( array.getValue(), 1465 );

		oParser = new parserFormula( "SUM(M:P+M:P)", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		assert.strictEqual( array.getValue(), 170 );

	} );

	QUnit.test("Test: \"ACOS\"", function (assert) {
		oParser = new parserFormula( 'ACOS(-0.5)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 2.094395102 );

		testArrayFormula(assert, "ACOS");
	} );

	QUnit.test("Test: \"ACOSH\"", function (assert) {
		oParser = new parserFormula( 'ACOSH(1)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( 'ACOSH(10)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 2.9932228 );

		testArrayFormula(assert, "ACOSH");
	} );

	QUnit.test("Test: \"ASIN\"", function (assert) {
		oParser = new parserFormula( 'ASIN(-0.5)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, -0.523598776 );

		testArrayFormula(assert, "ASIN");
	} );

	QUnit.test("Test: \"ASINH\"", function (assert) {
		oParser = new parserFormula( 'ASINH(-2.5)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, -1.647231146 );

		oParser = new parserFormula( 'ASINH(10)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 2.99822295 );

		testArrayFormula(assert, "ASINH");
	} );

    QUnit.test("Test: \"SIN have wrong arguments count\"", function (assert) {
        oParser = new parserFormula( 'SIN(3.1415926,3.1415926*2)', "A1", ws );
        assert.ok( !oParser.parse() );
    } );

    QUnit.test("Test: \"SIN(3.1415926)\"", function (assert) {
        oParser = new parserFormula( 'SIN(3.1415926)', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), Math.sin( 3.1415926 ) );

		testArrayFormula(assert, "SIN");
    } );

	QUnit.test("Test: \"SQRT\"", function (assert) {
		ws.getRange2( "A202" ).setValue( "-16" );

		oParser = new parserFormula( 'SQRT(16)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );

		oParser = new parserFormula( 'SQRT(A202)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		oParser = new parserFormula( 'SQRT(ABS(A202))', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );

		testArrayFormula(assert, "SQRT");
	} );

	QUnit.test("Test: \"SQRTPI\"", function (assert) {
		oParser = new parserFormula( 'SQRTPI(1)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 1.772454 );

		oParser = new parserFormula( 'SQRTPI(2)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 2.506628 );

		testArrayFormula(assert, "SQRTPI", true);
	} );

    QUnit.test("Test: \"COS(PI()/2)\"", function (assert) {
        oParser = new parserFormula( 'COS(PI()/2)', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), Math.cos( Math.PI / 2 ) );
    } );

	QUnit.test("Test: \"ACOT(2)\"", function (assert) {
		oParser = new parserFormula( 'ACOT(2)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), Math.PI / 2 - Math.atan(2) );
	} );

	QUnit.test("Test: \"ACOTH(6)\"", function (assert) {
		oParser = new parserFormula( 'ACOTH(6)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), Math.atanh(1 / 6) );

		testArrayFormula(assert, "ACOTH");
	} );



	QUnit.test("Test: \"COT\"", function (assert) {
		oParser = new parserFormula( 'COT(30)', "A1", ws );
		assert.ok( oParser.parse(), 'COT(30)' );
		assert.strictEqual( oParser.calculate().getValue().toFixed(3) - 0, -0.156, 'COT(30)' );

		oParser = new parserFormula( 'COT(0)', "A1", ws );
		assert.ok( oParser.parse(), 'COT(0)' );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!", 'COT(0)' );

		oParser = new parserFormula( 'COT(1000000000)', "A1", ws );
		assert.ok( oParser.parse(), 'COT(1000000000)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", 'COT(1000000000)' );

		oParser = new parserFormula( 'COT(-1000000000)', "A1", ws );
		assert.ok( oParser.parse(), 'COT(-1000000000)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", 'COT(-1000000000)' );

		oParser = new parserFormula( 'COT(test)', "A1", ws );
		assert.ok( oParser.parse(), 'COT(test)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NAME?", 'COT(test)' );

		oParser = new parserFormula( 'COT("test")', "A1", ws );
		assert.ok( oParser.parse(), 'COT("test")' );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!", 'COT("test")' );

		testArrayFormula(assert, "COT");
	} );

	QUnit.test("Test: \"COTH\"", function (assert) {
		oParser = new parserFormula( 'COTH(2)', "A1", ws );
		assert.ok( oParser.parse(), 'COTH(2)' );
		assert.strictEqual( oParser.calculate().getValue().toFixed(3) - 0, 1.037, 'COTH(2)' );

		oParser = new parserFormula( 'COTH(0)', "A1", ws );
		assert.ok( oParser.parse(), 'COTH(0)' );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!", 'COTH(0)' );

		oParser = new parserFormula( 'COTH(1000000000)', "A1", ws );
		assert.ok( oParser.parse(), 'COTH(1000000000)' );
		assert.strictEqual( oParser.calculate().getValue(), 1, 'COTH(1000000000)' );

		oParser = new parserFormula( 'COTH(-1000000000)', "A1", ws );
		assert.ok( oParser.parse(), 'COTH(-1000000000)' );
		assert.strictEqual( oParser.calculate().getValue(), -1, 'COTH(-1000000000)' );

		oParser = new parserFormula( 'COTH(test)', "A1", ws );
		assert.ok( oParser.parse(), 'COTH(test)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NAME?", 'COTH(test)' );

		oParser = new parserFormula( 'COTH("test")', "A1", ws );
		assert.ok( oParser.parse(), 'COTH("test")' );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!", 'COTH("test")' );

		testArrayFormula(assert, "COTH");
	} );

	QUnit.test("Test: \"CSC\"", function (assert) {
		oParser = new parserFormula( 'CSC(15)', "A1", ws );
		assert.ok( oParser.parse(), 'CSC(15)' );
		assert.strictEqual( oParser.calculate().getValue().toFixed(3) - 0, 1.538, 'CSC(15)' );

		oParser = new parserFormula( 'CSC(0)', "A1", ws );
		assert.ok( oParser.parse(), 'CSC(0)' );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!", 'CSC(0)' );

		oParser = new parserFormula( 'CSC(1000000000)', "A1", ws );
		assert.ok( oParser.parse(), 'CSC(1000000000)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", 'CSC(1000000000)' );

		oParser = new parserFormula( 'CSC(-1000000000)', "A1", ws );
		assert.ok( oParser.parse(), 'CSC(-1000000000)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", 'CSC(-1000000000)' );

		oParser = new parserFormula( 'CSC(test)', "A1", ws );
		assert.ok( oParser.parse(), 'CSC(test)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NAME?", 'CSC(test)' );

		oParser = new parserFormula( 'CSC("test")', "A1", ws );
		assert.ok( oParser.parse(), 'CSC("test")' );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!", 'CSC("test")' );

		testArrayFormula(assert, "CSC");
	} );

	QUnit.test("Test: \"CSCH\"", function (assert) {
		oParser = new parserFormula( 'CSCH(1.5)', "A1", ws );
		assert.ok( oParser.parse(), 'CSCH(1.5)' );
		assert.strictEqual( oParser.calculate().getValue().toFixed(4) - 0, 0.4696, 'CSCH(1.5)' );

		oParser = new parserFormula( 'CSCH(0)', "A1", ws );
		assert.ok( oParser.parse(), 'CSCH(0)' );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!", 'CSCH(0)' );

		oParser = new parserFormula( 'CSCH(1000000000)', "A1", ws );
		assert.ok( oParser.parse(), 'CSCH(1000000000)' );
		assert.strictEqual( oParser.calculate().getValue(), 0, 'CSCH(1000000000)' );

		oParser = new parserFormula( 'CSCH(-1000000000)', "A1", ws );
		assert.ok( oParser.parse(), 'CSCH(-1000000000)' );
		assert.strictEqual( oParser.calculate().getValue(), 0, 'CSCH(-1000000000)' );

		oParser = new parserFormula( 'CSCH(test)', "A1", ws );
		assert.ok( oParser.parse(), 'CSCH(test)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NAME?", 'CSCH(test)' );

		oParser = new parserFormula( 'CSCH("test")', "A1", ws );
		assert.ok( oParser.parse(), 'CSCH("test")' );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!", 'CSCH("test")' );

		testArrayFormula(assert, "CSCH");
	} );

	QUnit.test("Test: \"CLEAN\"", function (assert) {
		ws.getRange2( "A202" ).setValue( '=CHAR(9)&"Monthly report"&CHAR(10)' );

		oParser = new parserFormula( 'CLEAN(A202)', "A1", ws );
		assert.ok( oParser.parse());
		assert.strictEqual( oParser.calculate().getValue(), "Monthly report" );

		testArrayFormula(assert, "CLEAN");
	} );

	QUnit.test("Test: \"DEGREES\"", function (assert) {
		oParser = new parserFormula( 'DEGREES(PI())', "A1", ws );
		assert.ok( oParser.parse(), 'DEGREES(PI())' );
		assert.strictEqual( oParser.calculate().getValue(), 180, 'DEGREES(PI())' );

		testArrayFormula(assert, "DEGREES");
	} );

	QUnit.test("Test: \"SEC\"", function (assert) {
		oParser = new parserFormula( 'SEC(45)', "A1", ws );
		assert.ok( oParser.parse(), 'SEC(45)' );
		assert.strictEqual( oParser.calculate().getValue().toFixed(5) - 0, 1.90359, 'SEC(45)' );

		oParser = new parserFormula( 'SEC(30)', "A1", ws );
		assert.ok( oParser.parse(), 'SEC(30)' );
		assert.strictEqual( oParser.calculate().getValue().toFixed(5) - 0, 6.48292, 'SEC(30)' );

		oParser = new parserFormula( 'SEC(0)', "A1", ws );
		assert.ok( oParser.parse(), 'SEC(0)' );
		assert.strictEqual( oParser.calculate().getValue(), 1, 'SEC(0)' );

		oParser = new parserFormula( 'SEC(1000000000)', "A1", ws );
		assert.ok( oParser.parse(), 'SEC(1000000000)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", 'SEC(1000000000)' );

		oParser = new parserFormula( 'SEC(test)', "A1", ws );
		assert.ok( oParser.parse(), 'SEC(test)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NAME?", 'SEC(test)' );

		oParser = new parserFormula( 'SEC("test")', "A1", ws );
		assert.ok( oParser.parse(), 'SEC("test")' );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!", 'SEC("test")' );

		testArrayFormula(assert, "SEC");
	} );

	QUnit.test("Test: \"SECH\"", function (assert) {
		oParser = new parserFormula( 'SECH(5)', "A1", ws );
		assert.ok( oParser.parse(), 'SECH(5)' );
		assert.strictEqual( oParser.calculate().getValue().toFixed(3) - 0, 0.013, 'SECH(5)' );

		oParser = new parserFormula( 'SECH(0)', "A1", ws );
		assert.ok( oParser.parse(), 'SECH(0)' );
		assert.strictEqual( oParser.calculate().getValue(), 1, 'SECH(0)' );

		oParser = new parserFormula( 'SECH(1000000000)', "A1", ws );
		assert.ok( oParser.parse(), 'SECH(1000000000)' );
		assert.strictEqual( oParser.calculate().getValue(), 0, 'SECH(1000000000)' );

		oParser = new parserFormula( 'SECH(test)', "A1", ws );
		assert.ok( oParser.parse(), 'SECH(test)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NAME?", 'SECH(test)' );

		oParser = new parserFormula( 'SECH("test")', "A1", ws );
		assert.ok( oParser.parse(), 'SECH("test")' );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!", 'SECH("test")' );

		testArrayFormula(assert, "SECH");
	} );

	QUnit.test("Test: \"SECOND\"", function (assert) {

		ws.getRange2( "A202" ).setValue( "12:45:03 PM" );
		ws.getRange2( "A203" ).setValue( "4:48:18 PM" );
		ws.getRange2( "A204" ).setValue( "4:48 PM" );

		oParser = new parserFormula( "SECOND(A202)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "SECOND(A203)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 18 );

		oParser = new parserFormula( "SECOND(A204)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		testArrayFormula2(assert, "SECOND",1,1);
	} );

	QUnit.test("Test: \"FLOOR\"", function (assert) {
		oParser = new parserFormula( 'FLOOR(3.7,2)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR(3.7,2)' );
		assert.strictEqual( oParser.calculate().getValue(), 2, 'FLOOR(3.7,2)' );

		oParser = new parserFormula( 'FLOOR(-2.5,-2)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR(-2.5,-2)' );
		assert.strictEqual( oParser.calculate().getValue(), -2, 'FLOOR(-2.5,-2)' );

		oParser = new parserFormula( 'FLOOR(2.5,-2)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR(2.5,-2)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", 'FLOOR(2.5,-2)' );

		oParser = new parserFormula( 'FLOOR(1.58,0.1)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR(1.58,0.1)' );
		assert.strictEqual( oParser.calculate().getValue(), 1.5, 'FLOOR(1.58,0.1)' );

		oParser = new parserFormula( 'FLOOR(0.234,0.01)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR(0.234,0.01)' );
		assert.strictEqual( oParser.calculate().getValue(), 0.23, 'FLOOR(0.234,0.01)' );

		testArrayFormula2(assert, "FLOOR", 2, 2);
	} );

	QUnit.test("Test: \"FLOOR.PRECISE\"", function (assert) {
		oParser = new parserFormula( 'FLOOR.PRECISE(-3.2, -1)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR.PRECISE(-3.2, -1)' );
		assert.strictEqual( oParser.calculate().getValue(), -4, 'FLOOR.PRECISE(-3.2, -1)' );

		oParser = new parserFormula( 'FLOOR.PRECISE(3.2, 1)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR.PRECISE(3.2, 1)' );
		assert.strictEqual( oParser.calculate().getValue(), 3, 'FLOOR.PRECISE(3.2, 1)' );

		oParser = new parserFormula( 'FLOOR.PRECISE(-3.2, 1)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR.PRECISE(-3.2, 1)' );
		assert.strictEqual( oParser.calculate().getValue(), -4, 'FLOOR.PRECISE(-3.2, 1)' );

		oParser = new parserFormula( 'FLOOR.PRECISE(3.2, -1)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR.PRECISE(3.2, -1)' );
		assert.strictEqual( oParser.calculate().getValue(), 3, 'FLOOR.PRECISE(3.2, -1)' );

		oParser = new parserFormula( 'FLOOR.PRECISE(3.2)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR.PRECISE(3.2)' );
		assert.strictEqual( oParser.calculate().getValue(), 3, 'FLOOR.PRECISE(3.2)' );

		oParser = new parserFormula( 'FLOOR.PRECISE(test)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR.PRECISE(test)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NAME?", 'FLOOR.PRECISE(test)' );

		testArrayFormula2(assert, "FLOOR.PRECISE", 1, 2);
	} );

	QUnit.test("Test: \"FLOOR.MATH\"", function (assert) {
		oParser = new parserFormula( 'FLOOR.MATH(24.3, 5)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR.MATH(24.3, 5)' );
		assert.strictEqual( oParser.calculate().getValue(), 20, 'FLOOR.MATH(24.3, 5)' );

		oParser = new parserFormula( 'FLOOR.MATH(6.7)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR.MATH(6.7)' );
		assert.strictEqual( oParser.calculate().getValue(), 6, 'FLOOR.MATH(6.7)' );

		oParser = new parserFormula( 'FLOOR.MATH(-8.1, 5)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR.MATH(-8.1, 5)' );
		assert.strictEqual( oParser.calculate().getValue(), -10, 'FLOOR.MATH(-8.1, 5)' );

		oParser = new parserFormula( 'FLOOR.MATH(-5.5, 2, -1)', "A1", ws );
		assert.ok( oParser.parse(), 'FLOOR.MATH(-5.5, 2, -1)' );
		assert.strictEqual( oParser.calculate().getValue(), -4, 'FLOOR.MATH(-5.5, 2, -1)' );

		testArrayFormula2(assert, "FLOOR.MATH", 1, 3);
	} );

	QUnit.test("Test: \"CEILING.MATH\"", function (assert) {
		oParser = new parserFormula( 'CEILING.MATH(24.3, 5)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING.MATH(24.3, 5)' );
		assert.strictEqual( oParser.calculate().getValue(), 25, 'CEILING.MATH(24.3, 5)' );

		oParser = new parserFormula( 'CEILING.MATH(6.7)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING.MATH(6.7)' );
		assert.strictEqual( oParser.calculate().getValue(), 7, 'CEILING.MATH(6.7)' );

		oParser = new parserFormula( 'CEILING.MATH(-8.1, 2)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING.MATH(-8.1, 2)' );
		assert.strictEqual( oParser.calculate().getValue(), -8, 'CEILING.MATH(-8.1, 2)' );

		oParser = new parserFormula( 'CEILING.MATH(-5.5, 2, -1)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING.MATH(-5.5, 2, -1)' );
		assert.strictEqual( oParser.calculate().getValue(), -6, 'CEILING.MATH(-5.5, 2, -1)' );

		testArrayFormula2(assert, "CEILING.MATH", 1, 3);
	} );

	QUnit.test("Test: \"CEILING.PRECISE\"", function (assert) {
		oParser = new parserFormula( 'CEILING.PRECISE(4.3)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING.PRECISE(4.3)' );
		assert.strictEqual( oParser.calculate().getValue(), 5, 'CEILING.PRECISE(4.3)' );

		oParser = new parserFormula( 'CEILING.PRECISE(-4.3)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING.PRECISE(-4.3)' );
		assert.strictEqual( oParser.calculate().getValue(), -4, 'CEILING.PRECISE(-4.3)' );

		oParser = new parserFormula( 'CEILING.PRECISE(4.3, 2)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING.PRECISE(4.3, 2)' );
		assert.strictEqual( oParser.calculate().getValue(), 6, 'CEILING.PRECISE(4.3, 2)' );

		oParser = new parserFormula( 'CEILING.PRECISE(4.3,-2)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING.PRECISE(4.3,-2)' );
		assert.strictEqual( oParser.calculate().getValue(), 6, 'CEILING.PRECISE(4.3,-2)' );

		oParser = new parserFormula( 'CEILING.PRECISE(-4.3,2)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING.PRECISE(-4.3,2)' );
		assert.strictEqual( oParser.calculate().getValue(), -4, 'CEILING.PRECISE(-4.3,2)' );

		oParser = new parserFormula( 'CEILING.PRECISE(-4.3,-2)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING.PRECISE(-4.3,-2)' );
		assert.strictEqual( oParser.calculate().getValue(), -4, 'CEILING.PRECISE(-4.3,-2)' );

		oParser = new parserFormula( 'CEILING.PRECISE(test)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING.PRECISE(test)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NAME?", 'CEILING.PRECISE(test)' );

		testArrayFormula2(assert, "CEILING.PRECISE", 1, 2);
	} );

	QUnit.test("Test: \"ISO.CEILING\"", function (assert) {
		oParser = new parserFormula( 'ISO.CEILING(4.3)', "A1", ws );
		assert.ok( oParser.parse(), 'ISO.CEILING(4.3)' );
		assert.strictEqual( oParser.calculate().getValue(), 5, 'ISO.CEILING(4.3)' );

		oParser = new parserFormula( 'ISO.CEILING(-4.3)', "A1", ws );
		assert.ok( oParser.parse(), 'ISO.CEILING(-4.3)' );
		assert.strictEqual( oParser.calculate().getValue(), -4, 'ISO.CEILING(-4.3)' );

		oParser = new parserFormula( 'ISO.CEILING(4.3, 2)', "A1", ws );
		assert.ok( oParser.parse(), 'ISO.CEILING(4.3, 2)' );
		assert.strictEqual( oParser.calculate().getValue(), 6, 'ISO.CEILING(4.3, 2)' );

		oParser = new parserFormula( 'ISO.CEILING(4.3,-2)', "A1", ws );
		assert.ok( oParser.parse(), 'ISO.CEILING(4.3,-2)' );
		assert.strictEqual( oParser.calculate().getValue(), 6, 'ISO.CEILING(4.3,-2)' );

		oParser = new parserFormula( 'ISO.CEILING(-4.3,2)', "A1", ws );
		assert.ok( oParser.parse(), 'ISO.CEILING(-4.3,2)' );
		assert.strictEqual( oParser.calculate().getValue(), -4, 'ISO.CEILING(-4.3,2)' );

		oParser = new parserFormula( 'ISO.CEILING(-4.3,-2)', "A1", ws );
		assert.ok( oParser.parse(), 'ISO.CEILING(-4.3,-2)' );
		assert.strictEqual( oParser.calculate().getValue(), -4, 'ISO.CEILING(-4.3,-2)' );

		testArrayFormula2(assert, "ISO.CEILING", 1, 2);
	} );

	QUnit.test("Test: \"ISBLANK\"", function (assert) {

		ws.getRange2( "A202" ).setValue( "" );
		ws.getRange2( "A203" ).setValue( "test" );

		oParser = new parserFormula( 'ISBLANK(A202)', "A1", ws );
		assert.ok( oParser.parse(), 'ISBLANK(A202)' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", 'ISBLANK(A202)' );

		oParser = new parserFormula( 'ISBLANK(A203)', "A1", ws );
		assert.ok( oParser.parse(), 'ISBLANK(A203)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", 'ISBLANK(A203)' );

		testArrayFormula2(assert, "ISBLANK", 1, 1);
	} );

	QUnit.test("Test: \"ISERROR\"", function (assert) {

		ws.getRange2( "A202" ).setValue( "" );
		ws.getRange2( "A203" ).setValue( "#N/A" );

		oParser = new parserFormula( 'ISERROR(A202)', "A1", ws );
		assert.ok( oParser.parse(), 'ISERROR(A202)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", 'ISERROR(A202)' );

		oParser = new parserFormula( 'ISERROR(A203)', "A1", ws );
		assert.ok( oParser.parse(), 'ISERROR(A203)' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", 'ISERROR(A203)' );

		testArrayFormula2(assert, "ISERROR", 1, 1);
	} );

	QUnit.test("Test: \"ISERR\"", function (assert) {

		ws.getRange2( "A202" ).setValue( "" );
		ws.getRange2( "A203" ).setValue( "#N/A" );
		ws.getRange2( "A204" ).setValue( "#VALUE!" );

		oParser = new parserFormula( 'ISERR(A202)', "A1", ws );
		assert.ok( oParser.parse(), 'ISERR(A202)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", 'ISERR(A202)' );

		oParser = new parserFormula( 'ISERR(A203)', "A1", ws );
		assert.ok( oParser.parse(), 'ISERR(A203)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", 'ISERR(A203)' );

		oParser = new parserFormula( 'ISERR(A203)', "A1", ws );
		assert.ok( oParser.parse(), 'ISERR(A203)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", 'ISERR(A203)' );

		testArrayFormula2(assert, "ISERR", 1, 1);
	} );

	QUnit.test("Test: \"ISEVEN\"", function (assert) {

		oParser = new parserFormula( 'ISEVEN(-1)', "A1", ws );
		assert.ok( oParser.parse(), 'ISEVEN(-1)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", 'ISEVEN(-1)' );

		oParser = new parserFormula( 'ISEVEN(2.5)', "A1", ws );
		assert.ok( oParser.parse(), 'ISEVEN(2.5)' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", 'ISEVEN(2.5)' );

		oParser = new parserFormula( 'ISEVEN(5)', "A1", ws );
		assert.ok( oParser.parse(), 'ISEVEN(5)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", 'ISEVEN(5)' );

		oParser = new parserFormula( 'ISEVEN(0)', "A1", ws );
		assert.ok( oParser.parse(), 'ISEVEN(0)' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", 'ISEVEN(0)' );

		oParser = new parserFormula( 'ISEVEN(12/23/2011)', "A1", ws );
		assert.ok( oParser.parse(), 'ISEVEN(12/23/2011)' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", 'ISEVEN(12/23/2011)' );

		testArrayFormula2(assert, "ISEVEN", 1, 1, true, null);
	} );

	QUnit.test("Test: \"ISLOGICAL\"", function (assert) {

		oParser = new parserFormula( 'ISLOGICAL(TRUE)', "A1", ws );
		assert.ok( oParser.parse(), 'ISLOGICAL(TRUE)' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", 'ISLOGICAL(TRUE)' );

		oParser = new parserFormula( 'ISLOGICAL("TRUE")', "A1", ws );
		assert.ok( oParser.parse(), 'ISLOGICAL("TRUE")' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", 'ISLOGICAL("TRUE")' );

		testArrayFormula2(assert, "ISLOGICAL", 1, 1);
	} );

	QUnit.test("Test: \"CEILING\"", function (assert) {

		oParser = new parserFormula( 'CEILING(2.5, 1)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING(2.5, 1)' );
		assert.strictEqual( oParser.calculate().getValue(), 3, 'CEILING(2.5, 1)' );

		oParser = new parserFormula( 'CEILING(-2.5, -2)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING(-2.5, -2)' );
		assert.strictEqual( oParser.calculate().getValue(), -4, 'CEILING(-2.5, -2)' );

		oParser = new parserFormula( 'CEILING(-2.5, 2)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING(-2.5, 2)' );
		assert.strictEqual( oParser.calculate().getValue(), -2, 'CEILING(-2.5, 2)' );

		oParser = new parserFormula( 'CEILING(1.5, 0.1)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING(1.5, 0.1)' );
		assert.strictEqual( oParser.calculate().getValue(), 1.5, 'CEILING(1.5, 0.1)' );

		oParser = new parserFormula( 'CEILING(0.234, 0.01)', "A1", ws );
		assert.ok( oParser.parse(), 'CEILING(0.234, 0.01)' );
		assert.strictEqual( oParser.calculate().getValue(), 0.24, 'CEILING(0.234, 0.01)' );

		testArrayFormula2(assert, "CEILING", 2, 2);
	} );

	QUnit.test("Test: \"ECMA.CEILING\"", function (assert) {

		oParser = new parserFormula( 'ECMA.CEILING(2.5, 1)', "A1", ws );
		assert.ok( oParser.parse(), 'ECMA.CEILING(2.5, 1)' );
		assert.strictEqual( oParser.calculate().getValue(), 3, 'ECMA.CEILING(2.5, 1)' );

		oParser = new parserFormula( 'ECMA.CEILING(-2.5, -2)', "A1", ws );
		assert.ok( oParser.parse(), 'ECMA.CEILING(-2.5, -2)' );
		assert.strictEqual( oParser.calculate().getValue(), -4, 'ECMA.CEILING(-2.5, -2)' );

		oParser = new parserFormula( 'ECMA.CEILING(-2.5, 2)', "A1", ws );
		assert.ok( oParser.parse(), 'ECMA.CEILING(-2.5, 2)' );
		assert.strictEqual( oParser.calculate().getValue(), -2, 'ECMA.CEILING(-2.5, 2)' );

		oParser = new parserFormula( 'ECMA.CEILING(1.5, 0.1)', "A1", ws );
		assert.ok( oParser.parse(), 'ECMA.CEILING(1.5, 0.1)' );
		assert.strictEqual( oParser.calculate().getValue(), 1.5, 'ECMA.CEILING(1.5, 0.1)' );

		oParser = new parserFormula( 'ECMA.CEILING(0.234, 0.01)', "A1", ws );
		assert.ok( oParser.parse(), 'ECMA.CEILING(0.234, 0.01)' );
		assert.strictEqual( oParser.calculate().getValue(), 0.24, 'ECMA.CEILING(0.234, 0.01)' );

	} );

	QUnit.test("Test: \"COMBINA\"", function (assert) {
		oParser = new parserFormula( 'COMBINA(4,3)', "A1", ws );
		assert.ok( oParser.parse(), 'COMBINA(4,3)' );
		assert.strictEqual( oParser.calculate().getValue(), 20, 'COMBINA(4,3)' );

		oParser = new parserFormula( 'COMBINA(10,3)', "A1", ws );
		assert.ok( oParser.parse(), 'COMBINA(10,3)' );
		assert.strictEqual( oParser.calculate().getValue(), 220, 'COMBINA(10,3)' );

		oParser = new parserFormula( 'COMBINA(3,10)', "A1", ws );
		assert.ok( oParser.parse(), 'COMBINA(3,10)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", 'COMBINA(10,3)' );

		oParser = new parserFormula( 'COMBINA(10,-3)', "A1", ws );
		assert.ok( oParser.parse(), 'COMBINA(10,-3)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", 'COMBINA(10,-3)' );

		testArrayFormula2(assert, "COMBINA", 2, 2);
	} );

	QUnit.test("Test: \"DECIMAL\"", function (assert) {
		oParser = new parserFormula( 'DECIMAL("FF",16)', "A1", ws );
		assert.ok( oParser.parse(), 'DECIMAL("FF",16)' );
		assert.strictEqual( oParser.calculate().getValue(), 255, 'DECIMAL("FF",16)' );

		oParser = new parserFormula( 'DECIMAL(111,2)', "A1", ws );
		assert.ok( oParser.parse(), 'DECIMAL(111,2)' );
		assert.strictEqual( oParser.calculate().getValue(), 7, 'DECIMAL(111,2)' );

		oParser = new parserFormula( 'DECIMAL("zap",36)', "A1", ws );
		assert.ok( oParser.parse(), 'DECIMAL("zap",36)' );
		assert.strictEqual( oParser.calculate().getValue(), 45745, 'DECIMAL("zap",36)' );

		oParser = new parserFormula( 'DECIMAL("00FF",16)', "A1", ws );
		assert.ok( oParser.parse(), 'DECIMAL("00FF",16)' );
		assert.strictEqual( oParser.calculate().getValue(), 255, 'DECIMAL("00FF",16)' );

		oParser = new parserFormula( 'DECIMAL("101b",2)', "A1", ws );
		assert.ok( oParser.parse(), 'DECIMAL("101b",2)' );
		assert.strictEqual( oParser.calculate().getValue(), 5, 'DECIMAL("101b",2)' );

		testArrayFormula2(assert, "DECIMAL", 2, 2);
	} );

	QUnit.test("Test: \"BASE\"", function (assert) {
		oParser = new parserFormula( 'BASE(7,2)', "A1", ws );
		assert.ok( oParser.parse(), 'BASE(7,2)' );
		assert.strictEqual( oParser.calculate().getValue(), "111", 'BASE(7,2)' );

		oParser = new parserFormula( 'BASE(100,16)', "A1", ws );
		assert.ok( oParser.parse(), 'BASE(100,16)' );
		assert.strictEqual( oParser.calculate().getValue(), "64", 'BASE(100,16)' );

		oParser = new parserFormula( 'BASE(15,2,10)', "A1", ws );
		assert.ok( oParser.parse(), 'BASE(15,2,10)' );
		assert.strictEqual( oParser.calculate().getValue(), "0000001111", 'BASE(15,2,10)' );

		testArrayFormula2(assert, "BASE", 2, 3);
	} );

	QUnit.test("Test: \"ARABIC('LVII')\"", function (assert) {
		oParser = new parserFormula( 'ARABIC("LVII")', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 57 );
	} );

	QUnit.test("Test: \"TDIST\"", function (assert) {
		oParser = new parserFormula( "TDIST(60,1,2)", "A1", ws );
		assert.ok( oParser.parse(), "TDIST(60,1,2)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.010609347, "TDIST(60,1,2)" );

		oParser = new parserFormula( "TDIST(8,3,1)", "A1", ws );
		assert.ok( oParser.parse(), "TDIST(8,3,1)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.002038289, "TDIST(8,3,1)" );

		ws.getRange2( "A2" ).setValue( "1.959999998" );
		ws.getRange2( "A3" ).setValue( "60" );

		oParser = new parserFormula( "TDIST(A2,A3,2)", "A1", ws );
		assert.ok( oParser.parse(), "TDIST(A2,A3,2)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.054644930, "TDIST(A2,A3,2)" );

		oParser = new parserFormula( "TDIST(A2,A3,1)", "A1", ws );
		assert.ok( oParser.parse(), "TDIST(A2,A3,1)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.027322465, "TDIST(A2,A3,1)" );

		testArrayFormula2(assert, "TDIST", 3, 3);
	} );

	QUnit.test("Test: \"T.DIST\"", function (assert) {
		oParser = new parserFormula( "T.DIST(60,1,TRUE)", "A1", ws );
		assert.ok( oParser.parse(), "T.DIST(60,1,TRUE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 0.99469533, "T.DIST(60,1,TRUE)" );

		oParser = new parserFormula( "T.DIST(8,3,FALSE)", "A1", ws );
		assert.ok( oParser.parse(), "T.DIST(8,3,FALSE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 0.00073691, "T.DIST(8,3,FALSE)" );

		testArrayFormula2(assert, "T.DIST", 3, 3);
	} );

	QUnit.test("Test: \"T.DIST.2T\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "1.959999998" );
		ws.getRange2( "A3" ).setValue( "60" );

		oParser = new parserFormula( "T.DIST.2T(A2,A3)", "A1", ws );
		assert.ok( oParser.parse(), "T.DIST.2T(A2,A3)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.054644930, "T.DIST.2T(A2,A3)" );

		testArrayFormula2(assert, "T.DIST.2T", 2, 2);
	} );

	QUnit.test("Test: \"T.DIST.RT\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "1.959999998" );
		ws.getRange2( "A3" ).setValue( "60" );

		oParser = new parserFormula( "T.DIST.RT(A2,A3)", "A1", ws );
		assert.ok( oParser.parse(), "T.DIST.RT(A2,A3)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.027322, "T.DIST.RT(A2,A3)" );

		testArrayFormula2(assert, "T.DIST.RT", 2, 2);
	} );

	QUnit.test("Test: \"TTEST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "3" );
		ws.getRange2( "A3" ).setValue( "4" );
		ws.getRange2( "A4" ).setValue( "5" );
		ws.getRange2( "A5" ).setValue( "8" );
		ws.getRange2( "A6" ).setValue( "9" );
		ws.getRange2( "A7" ).setValue( "1" );
		ws.getRange2( "A8" ).setValue( "2" );
		ws.getRange2( "A9" ).setValue( "4" );
		ws.getRange2( "A10" ).setValue( "5" );

		ws.getRange2( "B2" ).setValue( "6" );
		ws.getRange2( "B3" ).setValue( "19" );
		ws.getRange2( "B4" ).setValue( "3" );
		ws.getRange2( "B5" ).setValue( "2" );
		ws.getRange2( "B6" ).setValue( "14" );
		ws.getRange2( "B7" ).setValue( "4" );
		ws.getRange2( "B8" ).setValue( "5" );
		ws.getRange2( "B9" ).setValue( "17" );
		ws.getRange2( "B10" ).setValue( "1" );

		oParser = new parserFormula( "TTEST(A2:A10,B2:B10,2,1)", "A1", ws );
		assert.ok( oParser.parse(), "TTEST(A2:A10,B2:B10,2,1)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.196016, "TTEST(A2:A10,B2:B10,2,1)" );

		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "TTEST", 4, 4, null, true);
	} );

	QUnit.test("Test: \"T.TEST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "3" );
		ws.getRange2( "A3" ).setValue( "4" );
		ws.getRange2( "A4" ).setValue( "5" );
		ws.getRange2( "A5" ).setValue( "8" );
		ws.getRange2( "A6" ).setValue( "9" );
		ws.getRange2( "A7" ).setValue( "1" );
		ws.getRange2( "A8" ).setValue( "2" );
		ws.getRange2( "A9" ).setValue( "4" );
		ws.getRange2( "A10" ).setValue( "5" );

		ws.getRange2( "B2" ).setValue( "6" );
		ws.getRange2( "B3" ).setValue( "19" );
		ws.getRange2( "B4" ).setValue( "3" );
		ws.getRange2( "B5" ).setValue( "2" );
		ws.getRange2( "B6" ).setValue( "14" );
		ws.getRange2( "B7" ).setValue( "4" );
		ws.getRange2( "B8" ).setValue( "5" );
		ws.getRange2( "B9" ).setValue( "17" );
		ws.getRange2( "B10" ).setValue( "1" );

		oParser = new parserFormula( "T.TEST(A2:A10,B2:B10,2,1)", "A1", ws );
		assert.ok( oParser.parse(), "T.TEST(A2:A10,B2:B10,2,1)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(5) - 0, 0.19602, "T.TEST(A2:A10,B2:B10,2,1)" );
	} );

	QUnit.test("Test: \"ZTEST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "3" );
		ws.getRange2( "A3" ).setValue( "6" );
		ws.getRange2( "A4" ).setValue( "7" );
		ws.getRange2( "A5" ).setValue( "8" );
		ws.getRange2( "A6" ).setValue( "6" );
		ws.getRange2( "A7" ).setValue( "5" );
		ws.getRange2( "A8" ).setValue( "4" );
		ws.getRange2( "A9" ).setValue( "2" );
		ws.getRange2( "A10" ).setValue( "1" );
		ws.getRange2( "A11" ).setValue( "9" );

		oParser = new parserFormula( "ZTEST(A2:A11,4)", "A1", ws );
		assert.ok( oParser.parse(), "ZTEST(A2:A11,4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.090574, "ZTEST(A2:A11,4)" );

		oParser = new parserFormula( "2 * MIN(ZTEST(A2:A11,4), 1 - ZTEST(A2:A11,4))", "A1", ws );
		assert.ok( oParser.parse(), "2 * MIN(ZTEST(A2:A11,4), 1 - ZTEST(A2:A11,4))" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.181148, "2 * MIN(ZTEST(A2:A11,4), 1 - ZTEST(A2:A11,4))" );

		oParser = new parserFormula( "ZTEST(A2:A11,6)", "A1", ws );
		assert.ok( oParser.parse(), "ZTEST(A2:A11,6)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.863043, "ZTEST(A2:A11,6)" );

		oParser = new parserFormula( "2 * MIN(ZTEST(A2:A11,6), 1 - ZTEST(A2:A11,6))", "A1", ws );
		assert.ok( oParser.parse(), "2 * MIN(ZTEST(A2:A11,6), 1 - ZTEST(A2:A11,6))" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.273913, "2 * MIN(ZTEST(A2:A11,6), 1 - ZTEST(A2:A11,6))" );

		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "Z.TEST", 2, 3, null, true);
	} );

	QUnit.test("Test: \"Z.TEST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "3" );
		ws.getRange2( "A3" ).setValue( "6" );
		ws.getRange2( "A4" ).setValue( "7" );
		ws.getRange2( "A5" ).setValue( "8" );
		ws.getRange2( "A6" ).setValue( "6" );
		ws.getRange2( "A7" ).setValue( "5" );
		ws.getRange2( "A8" ).setValue( "4" );
		ws.getRange2( "A9" ).setValue( "2" );
		ws.getRange2( "A10" ).setValue( "1" );
		ws.getRange2( "A11" ).setValue( "9" );

		oParser = new parserFormula( "Z.TEST(A2:A11,4)", "A1", ws );
		assert.ok( oParser.parse(), "Z.TEST(A2:A11,4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.090574, "Z.TEST(A2:A11,4)" );

		oParser = new parserFormula( "2 * MIN(Z.TEST(A2:A11,4), 1 - Z.TEST(A2:A11,4))", "A1", ws );
		assert.ok( oParser.parse(), "2 * MIN(Z.TEST(A2:A11,4), 1 - Z.TEST(A2:A11,4))" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.181148, "2 * MIN(Z.TEST(A2:A11,4), 1 - Z.TEST(A2:A11,4))" );

		oParser = new parserFormula( "Z.TEST(A2:A11,6)", "A1", ws );
		assert.ok( oParser.parse(), "Z.TEST(A2:A11,6)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.863043, "Z.TEST(A2:A11,6)" );

		oParser = new parserFormula( "2 * MIN(Z.TEST(A2:A11,6), 1 - Z.TEST(A2:A11,6))", "A1", ws );
		assert.ok( oParser.parse(), "2 * MIN(Z.TEST(A2:A11,6), 1 - Z.TEST(A2:A11,6))" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.273913, "2 * MIN(Z.TEST(A2:A11,6), 1 - Z.TEST(A2:A11,6))" );

		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "Z.TEST", 2, 3, null, true);
	} );



	QUnit.test("Test: \"F.DIST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "15.2069" );
		ws.getRange2( "A3" ).setValue( "6" );
		ws.getRange2( "A4" ).setValue( "4" );

		oParser = new parserFormula( "F.DIST(A2,A3,A4,TRUE)", "A1", ws );
		assert.ok( oParser.parse(), "F.DIST(A2,A3,A4,TRUE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.99, "F.DIST(A2,A3,A4,TRUE)" );

		oParser = new parserFormula( "F.DIST(A2,A3,A4,FALSE)", "A1", ws );
		assert.ok( oParser.parse(), "F.DIST(A2,A3,A4,FALSE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0012238, "F.DIST(A2,A3,A4,FALSE)" );

		testArrayFormula2(assert, "F.DIST", 4, 4);
	} );

	QUnit.test("Test: \"F.DIST.RT\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "15.2069" );
		ws.getRange2( "A3" ).setValue( "6" );
		ws.getRange2( "A4" ).setValue( "4" );

		oParser = new parserFormula( "F.DIST.RT(A2,A3,A4)", "A1", ws );
		assert.ok( oParser.parse(), "F.DIST.RT(A2,A3,A4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.01, "F.DIST.RT(A2,A3,A4)" );

		testArrayFormula2(assert, "F.DIST.RT", 3, 3);
	} );

	QUnit.test("Test: \"FDIST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "15.2069" );
		ws.getRange2( "A3" ).setValue( "6" );
		ws.getRange2( "A4" ).setValue( "4" );

		oParser = new parserFormula( "FDIST(A2,A3,A4)", "A1", ws );
		assert.ok( oParser.parse(), "FDIST(A2,A3,A4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.01, "FDIST(A2,A3,A4)" );
	} );

	QUnit.test("Test: \"FINV\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "0.01" );
		ws.getRange2( "A3" ).setValue( "6" );
		ws.getRange2( "A4" ).setValue( "4" );

		oParser = new parserFormula( "FINV(A2,A3,A4)", "A1", ws );
		assert.ok( oParser.parse(), "FINV(A2,A3,A4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 15.206865, "FINV(A2,A3,A4)" );

		testArrayFormula2(assert, "FINV", 3, 3);
	} );

	QUnit.test("Test: \"F.INV\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "0.01" );
		ws.getRange2( "A3" ).setValue( "6" );
		ws.getRange2( "A4" ).setValue( "4" );

		oParser = new parserFormula( "F.INV(A2,A3,A4)", "A1", ws );
		assert.ok( oParser.parse(), "F.INV(A2,A3,A4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 0.10930991, "F.INV(A2,A3,A4)" );

		testArrayFormula2(assert, "F.INV", 3, 3);
	} );

	QUnit.test("Test: \"F.INV.RT\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "0.01" );
		ws.getRange2( "A3" ).setValue( "6" );
		ws.getRange2( "A4" ).setValue( "4" );

		oParser = new parserFormula( "F.INV.RT(A2,A3,A4)", "A1", ws );
		assert.ok( oParser.parse(), "F.INV.RT(A2,A3,A4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(5) - 0, 15.20686, "F.INV.RT(A2,A3,A4)" );
	} );

	function fTestFormulaTest(assert){
		ws.getRange2( "A2" ).setValue( "6" );
		ws.getRange2( "A3" ).setValue( "7" );
		ws.getRange2( "A4" ).setValue( "9" );
		ws.getRange2( "A5" ).setValue( "15" );
		ws.getRange2( "A6" ).setValue( "21" );

		ws.getRange2( "B2" ).setValue( "20" );
		ws.getRange2( "B3" ).setValue( "28" );
		ws.getRange2( "B4" ).setValue( "31" );
		ws.getRange2( "B5" ).setValue( "38" );
		ws.getRange2( "B6" ).setValue( "40" );

		oParser = new parserFormula( "FTEST(A2:A6,B2:B6)", "A1", ws );
		assert.ok( oParser.parse(), "FTEST(A2:A6,B2:B6)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 0.64831785, "FTEST(A2:A6,B2:B6)" );

		oParser = new parserFormula( "FTEST(A2,B2:B6)", "A1", ws );
		assert.ok( oParser.parse(), "FTEST(A2,B2:B6)" );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!", "FTEST(A2,B2:B6)" );

		oParser = new parserFormula( "FTEST(1,B2:B6)", "A1", ws );
		assert.ok( oParser.parse(), "FTEST(1,B2:B6)" );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!", "FTEST(1,B2:B6)" );

		oParser = new parserFormula( "FTEST({1,2,3},{2,3,4,5})", "A1", ws );
		assert.ok( oParser.parse(), "FTEST({1,2,3},{2,3,4,5})" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.792636779, "FTEST({1,2,3},{2,3,4,5})" );

		oParser = new parserFormula( "FTEST({1,\"test\",\"test\"},{2,3,4,5})", "A1", ws );
		assert.ok( oParser.parse(), "FTEST({1,\"test\",\"test\"},{2,3,4,5})" );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!", "FTEST({1,\"test\",\"test\"},{2,3,4,5})" );
    }
	QUnit.test("Test: \"FTEST\"", function (assert) {
		fTestFormulaTest(assert);
		testArrayFormula2(assert, "FTEST", 2, 2, null, true);
	} );

	QUnit.test("Test: \"F.TEST\"", function (assert) {
		fTestFormulaTest(assert);
		testArrayFormula2(assert, "F.TEST", 2, 2, null, true);
	} );

	QUnit.test("Test: \"T.INV\"", function (assert) {
		oParser = new parserFormula( "T.INV(0.75,2)", "A1", ws );
		assert.ok( oParser.parse(), "T.INV(0.75,2)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.8164966, "T.INV(0.75,2)" );

		testArrayFormula2(assert, "T.INV", 2, 2);
	} );

	QUnit.test("Test: \"T.INV.2T\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "0.546449" );
		ws.getRange2( "A3" ).setValue( "60" );

		oParser = new parserFormula( "T.INV.2T(A2,A3)", "A1", ws );
		assert.ok( oParser.parse(), "T.INV.2T(A2,A3)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.606533, "T.INV.2T(A2,A3)" );

		testArrayFormula2(assert, "T.INV.2T", 2, 2);
	} );

	QUnit.test("Test: \"RANK\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "7" );
		ws.getRange2( "A3" ).setValue( "3.5" );
		ws.getRange2( "A4" ).setValue( "3.5" );
		ws.getRange2( "A5" ).setValue( "1" );
		ws.getRange2( "A6" ).setValue( "2" );

		oParser = new parserFormula( "RANK(A3,A2:A6,1)", "A1", ws );
		assert.ok( oParser.parse(), "RANK(A3,A2:A6,1)" );
		assert.strictEqual( oParser.calculate().getValue(), 3, "RANK(A3,A2:A6,1)" );

		oParser = new parserFormula( "RANK(A2,A2:A6,1)", "A1", ws );
		assert.ok( oParser.parse(), "RANK(A2,A2:A6,1)" );
		assert.strictEqual( oParser.calculate().getValue(), 5, "RANK(A2,A2:A6,1)" );
	} );

	QUnit.test("Test: \"RANK.EQ\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "7" );
		ws.getRange2( "A3" ).setValue( "3.5" );
		ws.getRange2( "A4" ).setValue( "3.5" );
		ws.getRange2( "A5" ).setValue( "1" );
		ws.getRange2( "A6" ).setValue( "2" );

		oParser = new parserFormula( "RANK.EQ(A2,A2:A6,1)", "A1", ws );
		assert.ok( oParser.parse(), "RANK.EQ(A2,A2:A6,1)" );
		assert.strictEqual( oParser.calculate().getValue(), 5, "RANK.EQ(A2,A2:A6,1)" );

		oParser = new parserFormula( "RANK.EQ(A6,A2:A6)", "A1", ws );
		assert.ok( oParser.parse(), "RANK.EQ(A6,A2:A6)" );
		assert.strictEqual( oParser.calculate().getValue(), 4, "RANK.EQ(A6,A2:A6)" );

		oParser = new parserFormula( "RANK.EQ(A3,A2:A6,1)", "A1", ws );
		assert.ok( oParser.parse(), "RANK.EQ(A3,A2:A6,1)" );
		assert.strictEqual( oParser.calculate().getValue(), 3, "RANK.EQ(A3,A2:A6,1)" );
	} );

	QUnit.test("Test: \"RANK.AVG\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "89" );
		ws.getRange2( "A3" ).setValue( "88" );
		ws.getRange2( "A4" ).setValue( "92" );
		ws.getRange2( "A5" ).setValue( "101" );
		ws.getRange2( "A6" ).setValue( "94" );
		ws.getRange2( "A7" ).setValue( "97" );
		ws.getRange2( "A8" ).setValue( "95" );

		oParser = new parserFormula( "RANK.AVG(94,A2:A8)", "A1", ws );
		assert.ok( oParser.parse(), "RANK.AVG(94,A2:A8)" );
		assert.strictEqual( oParser.calculate().getValue(), 4, "RANK.AVG(94,A2:A8)" );
	} );

	QUnit.test("Test: \"RADIANS\"", function (assert) {
		oParser = new parserFormula( "RADIANS(270)", "A1", ws );
		assert.ok( oParser.parse(), "RADIANS(270)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 4.712389 );

		testArrayFormula(assert, "RADIANS");
	} );

	QUnit.test("Test: \"LOG\"", function (assert) {
		oParser = new parserFormula( "LOG(10)", "A1", ws );
		assert.ok( oParser.parse(), "LOG(10)" );
		assert.strictEqual( oParser.calculate().getValue(), 1, "LOG(10)" );

		oParser = new parserFormula( "LOG(8,2)", "A1", ws );
		assert.ok( oParser.parse(), "LOG(8,2)" );
		assert.strictEqual( oParser.calculate().getValue(), 3, "LOG(8,2)" );

		oParser = new parserFormula( "LOG(86, 2.7182818)", "A1", ws );
		assert.ok( oParser.parse(), "LOG(86, 2.7182818)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 4.4543473, "LOG(86, 2.7182818)" );

		oParser = new parserFormula( "LOG(8,1)", "A1", ws );
		assert.ok( oParser.parse(), "LOG(8,1)" );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!", "LOG(8,1)" );

		testArrayFormula(assert, "LOG", 1, 2);
	} );

	QUnit.test("Test: \"LOGNORM.DIST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "4" );
		ws.getRange2( "A3" ).setValue( "3.5" );
		ws.getRange2( "A4" ).setValue( "1.2" );

		oParser = new parserFormula( "LOGNORM.DIST(A2,A3,A4,TRUE)", "A1", ws );
		assert.ok( oParser.parse(), "LOGNORM.DIST(A2,A3,A4,TRUE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0390836, "LOGNORM.DIST(A2,A3,A4,TRUE)" );

		oParser = new parserFormula( "LOGNORM.DIST(A2,A3,A4,FALSE)", "A1", ws );
		assert.ok( oParser.parse(), "LOGNORM.DIST(A2,A3,A4,FALSE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0176176, "LOGNORM.DIST(A2,A3,A4,FALSE)" );

		testArrayFormula2(assert, "LOGNORM.DIST", 4, 4);
	} );

	QUnit.test("Test: \"LOGNORM.INV\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "0.039084" );
		ws.getRange2( "A3" ).setValue( "3.5" );
		ws.getRange2( "A4" ).setValue( "1.2" );

		oParser = new parserFormula( "LOGNORM.INV(A2, A3, A4)", "A1", ws );
		assert.ok( oParser.parse(), "LOGNORM.INV(A2, A3, A4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 4.0000252, "LOGNORM.INV(A2, A3, A4)" );

		testArrayFormula2(assert, "LOGNORM.INV", 3, 3);
	} );

	QUnit.test("Test: \"LOGNORMDIST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "4" );
		ws.getRange2( "A3" ).setValue( "3.5" );
		ws.getRange2( "A4" ).setValue( "1.2" );

		oParser = new parserFormula( "LOGNORMDIST(A2, A3, A4)", "A1", ws );
		assert.ok( oParser.parse(), "LOGNORMDIST(A2, A3, A4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0390836, "LOGNORMDIST(A2, A3, A4)" );

		testArrayFormula2(assert, "LOGNORMDIST", 3, 3);
	} );

	QUnit.test("Test: \"LOWER\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "E. E. Cummings" );
		ws.getRange2( "A3" ).setValue( "Apt. 2B" );

		oParser = new parserFormula( "LOWER(A2)", "A1", ws );
		assert.ok( oParser.parse(), "LOWER(A2)" );
		assert.strictEqual( oParser.calculate().getValue(), "e. e. cummings", "LOWER(A2)" );

		oParser = new parserFormula( "LOWER(A3)", "A1", ws );
		assert.ok( oParser.parse(), "LOWER(A3)" );
		assert.strictEqual( oParser.calculate().getValue(), "apt. 2b", "LOWER(A3)" );

		testArrayFormula2(assert, "LOWER", 1, 1);
	} );

	QUnit.test("Test: \"EXPON.DIST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "0.2" );
		ws.getRange2( "A3" ).setValue( "10" );

		oParser = new parserFormula( "EXPON.DIST(A2,A3,TRUE)", "A1", ws );
		assert.ok( oParser.parse(), "EXPON.DIST(A2,A3,TRUE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 0.86466472, "EXPON.DIST(A2,A3,TRUE)" );

		oParser = new parserFormula( "EXPON.DIST(0.2,10,FALSE)", "A1", ws );
		assert.ok( oParser.parse(), "EXPON.DIST(0.2,10,FALSE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 1.35335283, "EXPON.DIST(0.2,10,FALSE)" );

		testArrayFormula2(assert, "EXPON.DIST", 3, 3);
	} );

	QUnit.test("Test: \"GAMMA.DIST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "10.00001131" );
		ws.getRange2( "A3" ).setValue( "9" );
		ws.getRange2( "A4" ).setValue( "2" );

	    oParser = new parserFormula( "GAMMA.DIST(A2,A3,A4,FALSE)", "A1", ws );
		assert.ok( oParser.parse(), "GAMMA.DIST(A2,A3,A4,FALSE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.032639, "GAMMA.DIST(A2,A3,A4,FALSE)" );

		oParser = new parserFormula( "GAMMA.DIST(A2,A3,A4,TRUE)", "A1", ws );
		assert.ok( oParser.parse(), "GAMMA.DIST(A2,A3,A4,TRUE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.068094, "GAMMA.DIST(A2,A3,A4,TRUE)" );

		testArrayFormula2(assert, "GAMMA.DIST", 4, 4);
	} );

	QUnit.test("Test: \"GAMMADIST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "10.00001131" );
		ws.getRange2( "A3" ).setValue( "9" );
		ws.getRange2( "A4" ).setValue( "2" );

		oParser = new parserFormula( "GAMMADIST(A2,A3,A4,FALSE)", "A1", ws );
		assert.ok( oParser.parse(), "GAMMADIST(A2,A3,A4,FALSE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.032639, "GAMMADIST(A2,A3,A4,FALSE)" );

		oParser = new parserFormula( "GAMMADIST(A2,A3,A4,TRUE)", "A1", ws );
		assert.ok( oParser.parse(), "GAMMADIST(A2,A3,A4,TRUE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.068094, "GAMMADIST(A2,A3,A4,TRUE)" );
	} );

	QUnit.test("Test: \"GAMMADIST\"", function (assert) {

		oParser = new parserFormula( "GAMMADIST(A2,A3,A4,FALSE)", "A1", ws );
		assert.ok( oParser.parse(), "GAMMADIST(A2,A3,A4,FALSE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.032639, "GAMMADIST(A2,A3,A4,FALSE)" );

		oParser = new parserFormula( "GAMMADIST(A2,A3,A4,TRUE)", "A1", ws );
		assert.ok( oParser.parse(), "GAMMADIST(A2,A3,A4,TRUE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.068094, "GAMMADIST(A2,A3,A4,TRUE)" );
	} );


	QUnit.test("Test: \"GAMMA\"", function (assert) {

		oParser = new parserFormula( "GAMMA(2.5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(3), "1.329" );

		oParser = new parserFormula( "GAMMA(-3.75)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(3), "0.268" );

		oParser = new parserFormula( "GAMMA(0)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		oParser = new parserFormula( "GAMMA(-2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );


		testArrayFormula2(assert, "GAMMA", 1, 1);
	} );

	QUnit.test("Test: \"CHITEST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "58" );
		ws.getRange2( "A3" ).setValue( "11" );
		ws.getRange2( "A4" ).setValue( "10" );
		ws.getRange2( "A5" ).setValue( "x" );
		ws.getRange2( "A6" ).setValue( "45.35" );
		ws.getRange2( "A7" ).setValue( "17.56" );
		ws.getRange2( "A8" ).setValue( "16.09" );

		ws.getRange2( "B2" ).setValue( "35" );
		ws.getRange2( "B3" ).setValue( "25" );
		ws.getRange2( "B4" ).setValue( "23" );
		ws.getRange2( "B5" ).setValue( "x" );
		ws.getRange2( "B6" ).setValue( "47.65" );
		ws.getRange2( "B7" ).setValue( "18.44" );
		ws.getRange2( "B8" ).setValue( "16.91" );

		oParser = new parserFormula( "CHITEST(A2:B4,A6:B8)", "A1", ws );
		assert.ok( oParser.parse(), "CHITEST(A2:B4,A6:B8)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0003082, "CHITEST(A2:B4,A6:B8)" );

		testArrayFormula2(assert, "CHITEST", 2, 2, null, true);
	} );

	QUnit.test("Test: \"CHISQ.TEST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "58" );
		ws.getRange2( "A3" ).setValue( "11" );
		ws.getRange2( "A4" ).setValue( "10" );
		ws.getRange2( "A5" ).setValue( "x" );
		ws.getRange2( "A6" ).setValue( "45.35" );
		ws.getRange2( "A7" ).setValue( "17.56" );
		ws.getRange2( "A8" ).setValue( "16.09" );

		ws.getRange2( "B2" ).setValue( "35" );
		ws.getRange2( "B3" ).setValue( "25" );
		ws.getRange2( "B4" ).setValue( "23" );
		ws.getRange2( "B5" ).setValue( "x" );
		ws.getRange2( "B6" ).setValue( "47.65" );
		ws.getRange2( "B7" ).setValue( "18.44" );
		ws.getRange2( "B8" ).setValue( "16.91" );

		oParser = new parserFormula( "CHISQ.TEST(A2:B4,A6:B8)", "A1", ws );
		assert.ok( oParser.parse(), "CHISQ.TEST(A2:B4,A6:B8)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0003082, "CHISQ.TEST(A2:B4,A6:B8)" );
	} );

	QUnit.test("Test: \"CHITEST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "18.307" );
		ws.getRange2( "A3" ).setValue( "10" );

		oParser = new parserFormula( "CHIDIST(A2,A3)", "A1", ws );
		assert.ok( oParser.parse(), "CHIDIST(A2,A3)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0500006, "CHIDIST(A2,A3)" );

		testArrayFormula2(assert, "CHIDIST", 2, 2);
	} );

	QUnit.test("Test: \"GAUSS\"", function (assert) {
		oParser = new parserFormula( "GAUSS(2)", "A1", ws );
		assert.ok( oParser.parse(), "GAUSS(2)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(5) - 0, 0.47725, "GAUSS(2)" );

		testArrayFormula2(assert, "GAUSS", 1, 1);
	} );

	QUnit.test("Test: \"CHISQ.DIST.RT\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "18.307" );
		ws.getRange2( "A3" ).setValue( "10" );

		oParser = new parserFormula( "CHISQ.DIST.RT(A2,A3)", "A1", ws );
		assert.ok( oParser.parse(), "CHISQ.DIST.RT(A2,A3)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0500006, "CHISQ.DIST.RT(A2,A3)" );

		testArrayFormula2(assert, "CHISQ.INV.RT", 2, 2);
	} );

	QUnit.test("Test: \"CHISQ.INV\"", function (assert) {
		oParser = new parserFormula( "CHISQ.INV(0.93,1)", "A1", ws );
		assert.ok( oParser.parse(), "CHISQ.INV(0.93,1)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 3.283020287, "CHISQ.INV(0.93,1)" );

		oParser = new parserFormula( "CHISQ.INV(0.6,2)", "A1", ws );
		assert.ok( oParser.parse(), "CHISQ.INV(0.6,2)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 1.832581464, "CHISQ.INV(0.6,2)" );

		testArrayFormula2(assert, "CHISQ.INV", 2, 2);
	} );

	QUnit.test("Test: \"CHISQ.DIST\"", function (assert) {
		oParser = new parserFormula( "CHISQ.DIST(0.5,1,TRUE)", "A1", ws );
		assert.ok( oParser.parse(), "CHISQ.DIST(0.5,1,TRUE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 0.52049988, "CHISQ.DIST(0.5,1,TRUE)" );

		oParser = new parserFormula( "CHISQ.DIST(2,3,FALSE)", "A1", ws );
		assert.ok( oParser.parse(), "CHISQ.DIST(2,3,FALSE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 0.20755375, "CHISQ.DIST(2,3,FALSE)" );

		testArrayFormula2(assert, "CHISQ.DIST", 3, 3);
	} );

	QUnit.test("Test: \"CHIINV\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "0.050001" );
		ws.getRange2( "A3" ).setValue( "10" );

		oParser = new parserFormula( "CHIINV(A2,A3)", "A1", ws );
		assert.ok( oParser.parse(), "CHIINV(A2,A3)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 18.306973, "CHIINV(A2,A3)" );

		testArrayFormula2(assert, "CHIINV", 2, 2);
	} );

	QUnit.test("Test: \"CHISQ.INV.RT\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "0.050001" );
		ws.getRange2( "A3" ).setValue( "10" );

		oParser = new parserFormula( "CHISQ.INV.RT(A2,A3)", "A1", ws );
		assert.ok( oParser.parse(), "CHISQ.INV.RT(A2,A3)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 18.306973, "CHISQ.INV.RT(A2,A3)" );

		testArrayFormula2(assert, "CHISQ.INV.RT", 2, 2);
	} );

	QUnit.test("Test: \"CHOOSE\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "st" );
		ws.getRange2( "A3" ).setValue( "2nd" );
		ws.getRange2( "A4" ).setValue( "3rd" );
		ws.getRange2( "A5" ).setValue( "Finished" );

		ws.getRange2( "B2" ).setValue( "Nails" );
		ws.getRange2( "B3" ).setValue( "Screws" );
		ws.getRange2( "B4" ).setValue( "Nuts" );
		ws.getRange2( "B5" ).setValue( "Bolts" );

		oParser = new parserFormula( "CHOOSE(2,A2,A3,A4,A5)", "A1", ws );
		assert.ok( oParser.parse(), "CHOOSE(2,A2,A3,A4,A5)" );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "2nd", "CHOOSE(2,A2,A3,A4,A5)" );

		oParser = new parserFormula( "CHOOSE(4,B2,B3,B4,B5)", "A1", ws );
		assert.ok( oParser.parse(), "CHOOSE(4,B2,B3,B4,B5)" );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "Bolts", "CHOOSE(4,B2,B3,B4,B5))" );

		oParser = new parserFormula( 'CHOOSE(3,"Wide",115,"world",8)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "world" );

		//функция возвращает ref
		//testArrayFormula2(assert, "CHOOSE", 2, 9);
	} );

	QUnit.test("Test: \"BETA.INV\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "0.685470581" );
		ws.getRange2( "A3" ).setValue( "8" );
		ws.getRange2( "A4" ).setValue( "10" );
		ws.getRange2( "A5" ).setValue( "1" );
		ws.getRange2( "A6" ).setValue( "3" );

		oParser = new parserFormula( "BETA.INV(A2,A3,A4,A5,A6)", "A1", ws );
		assert.ok( oParser.parse(), "BETA.INV(A2,A3,A4,A5,A6)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(1) - 0, 2, "BETA.INV(A2,A3,A4,A5,A6)" );

		testArrayFormula2(assert, "BETA.INV", 3, 5);
	} );

	QUnit.test("Test: \"BETAINV\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "0.685470581" );
		ws.getRange2( "A3" ).setValue( "8" );
		ws.getRange2( "A4" ).setValue( "10" );
		ws.getRange2( "A5" ).setValue( "1" );
		ws.getRange2( "A6" ).setValue( "3" );

		oParser = new parserFormula( "BETAINV(A2,A3,A4,A5,A6)", "A1", ws );
		assert.ok( oParser.parse(), "BETAINV(A2,A3,A4,A5,A6)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(1) - 0, 2, "BETAINV(A2,A3,A4,A5,A6)" );

		testArrayFormula2(assert, "BETAINV", 3, 5);
	} );

	QUnit.test("Test: \"BETA.DIST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "2" );
		ws.getRange2( "A3" ).setValue( "8" );
		ws.getRange2( "A4" ).setValue( "10" );
		ws.getRange2( "A5" ).setValue( "1" );
		ws.getRange2( "A6" ).setValue( "3" );

		oParser = new parserFormula( "BETA.DIST(A2,A3,A4,TRUE,A5,A6)", "A1", ws );
		assert.ok( oParser.parse(), "BETA.DIST(A2,A3,A4,TRUE,A5,A6)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.6854706, "BETA.DIST(A2,A3,A4,TRUE,A5,A6)" );

		oParser = new parserFormula( "BETA.DIST(A2,A3,A4,FALSE,A5,A6)", "A1", ws );
		assert.ok( oParser.parse(), "BETA.DIST(A2,A3,A4,FALSE,A5,A6)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 1.4837646, "BETA.DIST(A2,A3,A4,FALSE,A5,A6)" );

		testArrayFormula2(assert, "BETA.DIST", 4, 6);
	} );

	QUnit.test("Test: \"BETADIST\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "2" );
		ws.getRange2( "A3" ).setValue( "8" );
		ws.getRange2( "A4" ).setValue( "10" );
		ws.getRange2( "A5" ).setValue( "1" );
		ws.getRange2( "A6" ).setValue( "3" );

		oParser = new parserFormula( "BETADIST(A2,A3,A4,A5,A6)", "A1", ws );
		assert.ok( oParser.parse(), "BETADIST(A2,A3,A4,A5,A6)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.6854706, "BETADIST(A2,A3,A4,A5,A6)" );

		oParser = new parserFormula( "BETADIST(1,2,3,1,6)", "A1", ws );
		assert.ok( oParser.parse(), "BETADIST(1,2,3,1,6)" );
		assert.strictEqual( oParser.calculate().getValue(), 0, "BETADIST(1,2,3,1,6)" );

		oParser = new parserFormula( "BETADIST(6,2,3,1,6)", "A1", ws );
		assert.ok( oParser.parse(), "BETADIST(6,2,3,1,6)" );
		assert.strictEqual( oParser.calculate().getValue(), 1, "BETADIST(6,2,3,1,6)" );

		testArrayFormula2(assert, "BETADIST", 3, 5);
	} );

	QUnit.test("Test: \"BESSELJ\"", function (assert) {

		oParser = new parserFormula( "BESSELJ(1.9, 2)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELJ(1.9, 2)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.329925728, "BESSELJ(1.9, 2)" );

		oParser = new parserFormula( "BESSELJ(1.9, 2.4)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELJ(1.9, 2.4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.329925728, "BESSELJ(1.9, 2.4)" );

		oParser = new parserFormula( "BESSELJ(-1.9, 2.4)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELJ(-1.9, 2.4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.329925728, "BESSELJ(-1.9, 2.4)" );

		oParser = new parserFormula( "BESSELJ(-1.9, -2.4)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELJ(-1.9, -2.4)" );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		testArrayFormula2(assert, "BESSELJ", 2, 2, true, null);
	} );

	QUnit.test("Test: \"BESSELK\"", function (assert) {

		oParser = new parserFormula( "BESSELK(1.5, 1)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELK(1.5, 1)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.277387804, "BESSELK(1.5, 1)" );

		oParser = new parserFormula( "BESSELK(1, 3)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELK(1, 3)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 7.10126281, "BESSELK(1, 3)" );

		oParser = new parserFormula( "BESSELK(-1.123,2)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELK(-1.123,2)" );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		oParser = new parserFormula( "BESSELK(1,-2)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELK(1,-2)" );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		testArrayFormula2(assert, "BESSELK", 2, 2, true, null);

	} );

	QUnit.test("Test: \"BESSELY\"", function (assert) {

		oParser = new parserFormula( "BESSELY(2.5, 1)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELY(2.5, 1)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.1459181, "BESSELY(2.5, 1)" );

		oParser = new parserFormula( "BESSELY(1,-2)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELY(1,-2)" );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", "BESSELY(1,-2)" );

		oParser = new parserFormula( "BESSELY(-1,2)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELY(-1,2)" );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", "BESSELY(-1,2)" );

		testArrayFormula2(assert, "BESSELY", 2, 2, true, null);

	} );

	QUnit.test("Test: \"BESSELI\"", function (assert) {
		//есть различия excel в некоторых формулах(неточности в 7 цифре после точки)
		oParser = new parserFormula( "BESSELI(1.5, 1)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELI(1.5, 1)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.981666, "BESSELI(1.5, 1)" );

		oParser = new parserFormula( "BESSELI(1,2)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELI(1,2)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.135748, "BESSELI(1,2)" );

		oParser = new parserFormula( "BESSELI(1,-2)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELI(1,-2)" );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", "BESSELI(1,-2)" );

		oParser = new parserFormula( "BESSELI(-1,2)", "A1", ws );
		assert.ok( oParser.parse(), "BESSELI(-1,2)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.135748, "BESSELI(-1,2)" );

		testArrayFormula2(assert, "BESSELI", 2, 2, true, null);
	} );

	QUnit.test("Test: \"GAMMA.INV\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "0.068094" );
		ws.getRange2( "A3" ).setValue( "9" );
		ws.getRange2( "A4" ).setValue( "2" );

		oParser = new parserFormula( "GAMMA.INV(A2,A3,A4)", "A1", ws );
		assert.ok( oParser.parse(), "GAMMA.INV(A2,A3,A4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 10.0000112, "GAMMA.INV(A2,A3,A4)" );

		testArrayFormula2(assert, "GAMMA.INV", 3, 3);
	} );

	QUnit.test("Test: \"GAMMAINV\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "0.068094" );
		ws.getRange2( "A3" ).setValue( "9" );
		ws.getRange2( "A4" ).setValue( "2" );

		oParser = new parserFormula( "GAMMAINV(A2,A3,A4)", "A1", ws );
		assert.ok( oParser.parse(), "GAMMAINV(A2,A3,A4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 10.0000112, "GAMMAINV(A2,A3,A4)" );
	} );

	QUnit.test("Test: \"SUM(1,2,3)\"", function (assert) {
        oParser = new parserFormula( 'SUM(1,2,3)', "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 + 2 + 3 );

        testArrayFormula2(assert, "SUM", 1, 8, null, true);
    } );

    QUnit.test("Test: \"\"s\"&5\"", function (assert) {
        oParser = new parserFormula( "\"s\"&5", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "s5" );
    } );

    QUnit.test("Test: \"String+Number\"", function (assert) {
        oParser = new parserFormula( "1+\"099\"", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 100 );

        ws.getRange2( "A1469" ).setValue( "'099" );
        ws.getRange2( "A1470" ).setValue( "\"099\"" );

        oParser = new parserFormula( "1+A1469", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 100 );


        oParser = new parserFormula( "1+A1470", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

    } );

    QUnit.test("Test: \"POWER(2,8)\"", function (assert) {
        oParser = new parserFormula( "POWER(2,8)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), Math.pow( 2, 8 ) );
    } );

    QUnit.test("Test: \"POWER(0,-3)\"", function (assert) {
        oParser = new parserFormula( "POWER(0,-3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );

		testArrayFormula2(assert, "POWER", 2, 2);
    } );

	QUnit.test("Test: \"ISNA(A1)\"", function (assert) {
		ws.getRange2( "A1" ).setValue( "#N/A" );

		oParser = new parserFormula( "ISNA(A1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

		testArrayFormula2(assert, "ISNA",1,1);
	} );

	QUnit.test("Test: \"ISNONTEXT\"", function (assert) {
		oParser = new parserFormula( 'ISNONTEXT("123")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE" );

		testArrayFormula2(assert, "ISNONTEXT",1,1);
	} );

	QUnit.test("Test: \"ISNUMBER\"", function (assert) {
		ws.getRange2( "A1" ).setValue( "123" );

		oParser = new parserFormula( 'ISNUMBER(4)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

		oParser = new parserFormula( 'ISNUMBER(A1)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

		testArrayFormula2(assert, "ISNUMBER",1,1);
	} );

	QUnit.test("Test: \"ISODD\"", function (assert) {
		oParser = new parserFormula( 'ISODD(-1)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

		oParser = new parserFormula( 'ISODD(2.5)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE" );

		oParser = new parserFormula( 'ISODD(5)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

		testArrayFormula2(assert, "ISODD",1,1,true, null);
	} );

	QUnit.test("Test: \"ROUND\"", function (assert) {
		oParser = new parserFormula( "ROUND(2.15, 1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2.2 );

		oParser = new parserFormula( "ROUND(2.149, 1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2.1 );

		oParser = new parserFormula( "ROUND(-1.475, 2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -1.48 );

		oParser = new parserFormula( "ROUND(21.5, -1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 20 );

		oParser = new parserFormula( "ROUND(626.3,-3)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1000 );

		oParser = new parserFormula( "ROUND(1.98,-1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "ROUND(-50.55,-2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -100 );

		oParser = new parserFormula('ROUND("test",-2.1)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula('ROUND(123.431,"test")', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula('ROUND(123.431,#NUM!)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula('ROUND(#NUM!,123.431)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula( "ROUND(-50.55,-2.1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -100 );

		oParser = new parserFormula( "ROUND(-50.55,-2.9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -100 );

		oParser = new parserFormula( "ROUND(-50.55,0.9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -51 );

		oParser = new parserFormula( "ROUND(-50.55,0.1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -51 );


		testArrayFormula2(assert, "ROUND", 2, 2);
	} );

	QUnit.test("Test: \"ROUNDUP(31415.92654,-2)\"", function (assert) {
        oParser = new parserFormula( "ROUNDUP(31415.92654,-2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 31500 );
    } );

    QUnit.test("Test: \"ROUNDUP(3.2,0)\"", function (assert) {
        oParser = new parserFormula( "ROUNDUP(3.2,0)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4 );
    } );

    QUnit.test("Test: \"ROUNDUP(-3.14159,1)\"", function (assert) {
        oParser = new parserFormula( "ROUNDUP(-3.14159,1)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -3.2 );
    } );

    QUnit.test("Test: \"ROUNDUP(3.14159,3)\"", function (assert) {
        oParser = new parserFormula( "ROUNDUP(3.14159,3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3.142 );

		testArrayFormula2(assert, "ROUNDUP", 2, 2);
    } );

	QUnit.test("Test: \"ROUNDUP\"", function (assert) {
		oParser = new parserFormula( "ROUNDUP(2.1123,4)", "A1", ws );
		assert.ok(oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(4) - 0, 2.1123 );

		//TODO в хроме при расчёте разница, временно убираю
		oParser = new parserFormula( "ROUNDUP(2,4)", "A1", ws );
		assert.ok(oParser.parse() );
		//assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "ROUNDUP(2,0)", "A1", ws );
		assert.ok(oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "ROUNDUP(2.1123,-1)", "A1", ws );
		assert.ok(oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 10 );

		oParser = new parserFormula( "ROUNDUP(2.1123,0)", "A1", ws );
		assert.ok(oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula("ROUNDUP(123.431,0.1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue() - 0, 124);

		oParser = new parserFormula("ROUNDUP(123.431,0.9)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue() - 0, 124);

		oParser = new parserFormula("ROUNDUP(123.431,-0.9)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue() - 0, 124);

		oParser = new parserFormula("ROUNDUP(123.431,-0.1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue() - 0, 124);

		oParser = new parserFormula("ROUNDUP(123.431,-2.1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue() - 0, 200);

		oParser = new parserFormula('ROUNDUP("test",-2.1)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula('ROUNDUP(123.431,"test")', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula('ROUNDUP(123.431,#NUM!)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula('ROUNDUP(#NUM!,123.431)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula('ROUNDUP(123.431,-1.9)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue() - 0, 130);

		oParser = new parserFormula( "ROUNDUP(-50.55,0.9)", "A1", ws );
		assert.ok(oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -51 );

		oParser = new parserFormula( "ROUNDUP(-50.55,0.1)", "A1", ws );
		assert.ok(oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -51 );

	} );


    QUnit.test("Test: \"ROUNDDOWN(31415.92654,-2)\"", function (assert) {
        oParser = new parserFormula( "ROUNDDOWN(31415.92654,-2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 31400 );
    } );

    QUnit.test("Test: \"ROUNDDOWN(-3.14159,1)\"", function (assert) {
        oParser = new parserFormula( "ROUNDDOWN(-3.14159,1)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -3.1 );
    } );

    QUnit.test("Test: \"ROUNDDOWN(3.14159,3)\"", function (assert) {
        oParser = new parserFormula( "ROUNDDOWN(3.14159,3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3.141 );
    } );

    QUnit.test("Test: \"ROUNDDOWN(3.2,0)\"", function (assert) {
        oParser = new parserFormula( "ROUNDDOWN(3.2,0)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3 );

		testArrayFormula2(assert, "ROUNDDOWN", 2, 2);
    } );

	QUnit.test("Test: \"ROUNDDOWN\"", function (assert) {
		oParser = new parserFormula("ROUNDDOWN(123.431,0.1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue() - 0, 123);

		oParser = new parserFormula("ROUNDDOWN(123.431,0.9)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue() - 0, 123);

		oParser = new parserFormula("ROUNDDOWN(123.431,-0.9)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue() - 0, 123);

		oParser = new parserFormula("ROUNDDOWN(123.431,-0.1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue() - 0, 123);

		oParser = new parserFormula("ROUNDDOWN(123.431,-2.1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue() - 0, 100);

		oParser = new parserFormula('ROUNDDOWN("test",-2.1)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula('ROUNDDOWN(123.431,"test")', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula('ROUNDDOWN(123.431,#NUM!)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula('ROUNDDOWN(#NUM!,123.431)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula( "ROUNDDOWN(-50.55,0.9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -50 );

		oParser = new parserFormula( "ROUNDDOWN(-50.55,0.1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -50 );
	} );



    QUnit.test("Test: \"MROUND\"", function (assert) {
        var multiple;//должен равняться значению второго аргумента
        function mroundHelper( num ) {
            var multiplier = Math.pow( 10, Math.floor( Math.log( Math.abs( num ) ) / Math.log( 10 ) ) - AscCommonExcel.cExcelSignificantDigits + 1 );
            var nolpiat = 0.5 * (num > 0 ? 1 : num < 0 ? -1 : 0) * multiplier;
            var y = (num + nolpiat) / multiplier;
            y = y / Math.abs( y ) * Math.floor( Math.abs( y ) )
            var x = y * multiplier / multiple

            // var x = number / multiple;
            var nolpiat = 5 * (x / Math.abs( x )) * Math.pow( 10, Math.floor( Math.log( Math.abs( x ) ) / Math.log( 10 ) ) - AscCommonExcel.cExcelSignificantDigits );
            x = x + nolpiat;
            x = x | x;

            return x * multiple;
        }


        oParser = new parserFormula( "MROUND(10,3)", "A1", ws );
        assert.ok( oParser.parse() );
        multiple = 3;
        assert.strictEqual( oParser.calculate().getValue(), mroundHelper( 10 + 3 / 2 ) );

        oParser = new parserFormula( "MROUND(-10,-3)", "A1", ws );
        assert.ok( oParser.parse() );
        multiple = -3;
        assert.strictEqual( oParser.calculate().getValue(), mroundHelper( -10 + -3 / 2 ) );

        oParser = new parserFormula( "MROUND(1.3,0.2)", "A1", ws );
        assert.ok( oParser.parse() );
        multiple = 0.2;
        assert.strictEqual( oParser.calculate().getValue(), mroundHelper( 1.3 + 0.2 / 2 ) );

		testArrayFormula2(assert, "MROUND", 2, 2, true, null);
    } );

	QUnit.test("Test: \"T(\"HELLO\")\"", function (assert) {
		oParser = new parserFormula( "T(\"HELLO\")", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "HELLO" );
	} );

	QUnit.test("Test: \"MMULT\"", function (assert) {
		ws.getRange2( "AAA102" ).setValue( "4" );
		ws.getRange2( "AAA103" ).setValue( "5" );
		ws.getRange2( "AAA104" ).setValue( "6" );
		ws.getRange2( "AAA105" ).setValue( "7" );
		ws.getRange2( "AAB102" ).setValue( "1" );
		ws.getRange2( "AAB103" ).setValue( "2" );
		ws.getRange2( "AAB104" ).setValue( "3" );
		ws.getRange2( "AAB105" ).setValue( "2" );
		ws.getRange2( "AAC102" ).setValue( "4" );
		ws.getRange2( "AAC103" ).setValue( "5" );
		ws.getRange2( "AAC104" ).setValue( "6" );
		ws.getRange2( "AAC105" ).setValue( "3" );
		ws.getRange2( "AAD102" ).setValue( "7" );
		ws.getRange2( "AAD103" ).setValue( "8" );
		ws.getRange2( "AAD104" ).setValue( "9" );
		ws.getRange2( "AAD105" ).setValue( "4" );

		ws.getRange2( "AAF102" ).setValue( "1" );
		ws.getRange2( "AAF103" ).setValue( "2" );
		ws.getRange2( "AAF104" ).setValue( "3" );
		ws.getRange2( "AAF105" ).setValue( "6" );

		ws.getRange2( "AAG102" ).setValue( "2" );
		ws.getRange2( "AAG103" ).setValue( "3" );
		ws.getRange2( "AAG104" ).setValue( "4" );
		ws.getRange2( "AAG105" ).setValue( "5" );

		oParser = new parserFormula( "MMULT(AAC102,AAF104)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 12 );

		oParser = new parserFormula( "MMULT(AAA102:AAD105,AAF104)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "MMULT(AAC102,AAF104)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 12 );

		oParser = new parserFormula( "MMULT(AAA102:AAD105,AAF102:AAG105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 60 );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue(), 62 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,0).getValue(), 72 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,1).getValue(), 76 );

		oParser = new parserFormula( "MMULT(AAA102:AAD105,AAF102:AAF105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 60 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,0).getValue(), 72 );

		oParser = new parserFormula( "MMULT(AAA102:AAD105,AAF102:AAF105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 60 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,0).getValue(), 72 );

		oParser = new parserFormula( "MMULT(AAA102:AAD105,AAF102:AAF104)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "MMULT(AAA102:AAD105,AAK110:AAN110)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "MMULT(AAA102:AAD105,AAA102:AAD105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 94 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,0).getValue(), 116 );
		assert.strictEqual( oParser.calculate().getElementRowCol(2,0).getValue(), 138 );

		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue(), 32 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,1).getValue(), 40 );
		assert.strictEqual( oParser.calculate().getElementRowCol(2,1).getValue(), 48 );

		oParser = new parserFormula( "MMULT(AAF102:AAF105,AAG102:AAG105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "MMULT(AAF102:AAF105,AAA102:AAD102)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 4 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,0).getValue(), 8 );
		assert.strictEqual( oParser.calculate().getElementRowCol(2,0).getValue(), 12 );

	} );

    QUnit.test("Test: \"T(123)\"", function (assert) {
        oParser = new parserFormula( "T(123)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.ok( !oParser.calculate().getValue(), "123" );
    } );

    QUnit.test("Test: YEAR", function (assert) {
        oParser = new parserFormula( "YEAR(2013)", "A1", ws );
        assert.ok( oParser.parse() );
        if ( AscCommon.bDate1904 )
            assert.strictEqual( oParser.calculate().getValue(), 1909 );
        else
            assert.strictEqual( oParser.calculate().getValue(), 1905 );

        testArrayFormula2(assert, "YEAR",1,1);
    } );

    QUnit.test("Test: DAY", function (assert) {
        oParser = new parserFormula( "DAY(2013)", "A1", ws );
        assert.ok( oParser.parse() );
        if ( AscCommon.bDate1904 )
            assert.strictEqual( oParser.calculate().getValue(), 6 );
        else
            assert.strictEqual( oParser.calculate().getValue(), 5 );

        testArrayFormula2(assert, "DAY", 1, 1);
    } );

	QUnit.test("Test: DAYS", function (assert) {
		ws.getRange2( "A2" ).setValue( "12/31/2011" );
		ws.getRange2( "A3" ).setValue( "1/1/2011" );

	    oParser = new parserFormula( 'DAYS("3/15/11","2/1/11")', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 42 );

		oParser = new parserFormula( "DAYS(A2,A3)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 364 );

		oParser = new parserFormula( "DAYS(A2,A3)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 364 );

		oParser = new parserFormula( 'DAYS("2008-03-03","2008-03-01")', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( 'DAYS("2008-03-01","2008-03-03")', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -2 );


		testArrayFormula2(assert, "DAYS", 2, 2);
	} );

    QUnit.test("Test: DAY 2", function (assert) {
        oParser = new parserFormula( "DAY(\"20 may 2045\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 20 );
    } );

    QUnit.test("Test: MONTH #1", function (assert) {
        oParser = new parserFormula( "MONTH(2013)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 7 );
    } );

    QUnit.test("Test: MONTH #2", function (assert) {
        oParser = new parserFormula( "MONTH(DATE(2013,2,2))", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2 );
    } );

    QUnit.test("Test: MONTH #3", function (assert) {
        oParser = new parserFormula( "MONTH(NOW())", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), new cDate().getUTCMonth() + 1 );

        testArrayFormula2(assert, "MONTH",1,1);
    } );

    QUnit.test("Test: \"10-3\"", function (assert) {
        oParser = new parserFormula( "10-3", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 7 );
    } );

    QUnit.test("Test: \"SUM\"", function (assert) {

		ws.getRange2( "S5" ).setValue( "1" );
		ws.getRange2( "S6" ).setValue( numDivFact(-1, 2) );
		ws.getRange2( "S7" ).setValue( numDivFact(1, 4) );
		ws.getRange2( "S8" ).setValue( numDivFact(-1, 6) );

        oParser = new parserFormula( "SUM(S5:S8)", "A1", ws );
        assert.ok( oParser.parse() );
//        assert.strictEqual( oParser.calculate().getValue(), 1-1/Math.fact(2)+1/Math.fact(4)-1/Math.fact(6) );
        assert.ok( Math.abs( oParser.calculate().getValue() - (1 - 1 / Math.fact( 2 ) + 1 / Math.fact( 4 ) - 1 / Math.fact( 6 )) ) < dif );
    } );

    QUnit.test("Test: \"MAX\"", function (assert) {

		ws.getRange2( "S5" ).setValue( "1" );
		ws.getRange2( "S6" ).setValue( numDivFact(-1, 2) );
		ws.getRange2( "S7" ).setValue( numDivFact(1, 4) );
		ws.getRange2( "S8" ).setValue( numDivFact(-1, 6) );

        oParser = new parserFormula( "MAX(S5:S8)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

		ws.getRange2( "S5" ).setValue( "#DIV/0!" );
		ws.getRange2( "S6" ).setValue( "TRUE" );
		ws.getRange2( "S7" ).setValue( "qwe" );
		ws.getRange2( "S8" ).setValue( "" );
		ws.getRange2( "S9" ).setValue( "-1" );
		oParser = new parserFormula( "MAX(S5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );
		oParser = new parserFormula( "MAX(S6)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "MAX(S7)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "MAX(S8)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "MAX(S5:S9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );
		oParser = new parserFormula( "MAX(S6:S9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -1 );
		oParser = new parserFormula( "MAX(-1, TRUE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		testArrayFormula2(assert, "MAX", 1, 8, null, true);
    } );

    QUnit.test("Test: \"MAXA\"", function (assert) {

		ws.getRange2( "S5" ).setValue( "1" );
		ws.getRange2( "S6" ).setValue( numDivFact(-1, 2) );
		ws.getRange2( "S7" ).setValue( numDivFact(1, 4) );
		ws.getRange2( "S8" ).setValue( numDivFact(-1, 6) );

        oParser = new parserFormula( "MAXA(S5:S8)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

		ws.getRange2( "S5" ).setValue( "#DIV/0!" );
		ws.getRange2( "S6" ).setValue( "TRUE" );
		ws.getRange2( "S7" ).setValue( "qwe" );
		ws.getRange2( "S8" ).setValue( "" );
		ws.getRange2( "S9" ).setValue( "-1" );
		oParser = new parserFormula( "MAXA(S5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );
		oParser = new parserFormula( "MAXA(S6)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );
		oParser = new parserFormula( "MAXA(S7)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "MAXA(S8)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "MAXA(S5:S9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );
		oParser = new parserFormula( "MAXA(S6:S9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );
		oParser = new parserFormula( "MAXA(-1, TRUE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		testArrayFormula2(assert, "MAXA", 1, 8, null, true);
    } );

    QUnit.test("Test: \"MIN\"", function (assert) {

		ws.getRange2( "S5" ).setValue( "1" );
		ws.getRange2( "S6" ).setValue( numDivFact(-1, 2) );
		ws.getRange2( "S7" ).setValue( numDivFact(1, 4) );
		ws.getRange2( "S8" ).setValue( numDivFact(-1, 6) );

        oParser = new parserFormula( "MIN(S5:S8)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -1 / Math.fact( 2 ) );

		ws.getRange2( "S5" ).setValue( "#DIV/0!" );
		ws.getRange2( "S6" ).setValue( "TRUE" );
		ws.getRange2( "S7" ).setValue( "qwe" );
		ws.getRange2( "S8" ).setValue( "" );
		ws.getRange2( "S9" ).setValue( "2" );
		oParser = new parserFormula( "MIN(S5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );
		oParser = new parserFormula( "MIN(S6)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "MIN(S7)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "MIN(S8)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "MIN(S5:S9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );
		oParser = new parserFormula( "MIN(S6:S9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );
		oParser = new parserFormula( "MIN(2, TRUE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		testArrayFormula2(assert, "min", 1, 8, null, true);
    } );

    QUnit.test("Test: \"MINA\"", function (assert) {

		ws.getRange2( "S5" ).setValue( "1" );
		ws.getRange2( "S6" ).setValue( numDivFact(-1, 2) );
		ws.getRange2( "S7" ).setValue( numDivFact(1, 4) );
		ws.getRange2( "S8" ).setValue( numDivFact(-1, 6) );

        oParser = new parserFormula( "MINA(S5:S8)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -1 / Math.fact( 2 ) );

		ws.getRange2( "S5" ).setValue( "#DIV/0!" );
		ws.getRange2( "S6" ).setValue( "TRUE" );
		ws.getRange2( "S7" ).setValue( "qwe" );
		ws.getRange2( "S8" ).setValue( "" );
		ws.getRange2( "S9" ).setValue( "2" );
		oParser = new parserFormula( "MINA(S5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );
		oParser = new parserFormula( "MINA(S6)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );
		oParser = new parserFormula( "MINA(S7)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "MINA(S8)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "MINA(S5:S9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );
		oParser = new parserFormula( "MINA(S6:S9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "MINA(2, TRUE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		testArrayFormula2(assert, "mina", 1, 8, null, true);
    } );

    QUnit.test("Test: SUM(S7:S9,{1,2,3})", function (assert) {
        ws.getRange2( "S7" ).setValue( "1" );
        ws.getRange2( "S8" ).setValue( "2" );
        ws.getRange2( "S9" ).setValue( "3" );

		oParser = new parserFormula( "SUM(S7:S9,{1,2,3})", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 12 );
    } );

    QUnit.test("Test: ISREF", function (assert) {
        oParser = new parserFormula( "ISREF(G0)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "FALSE" );

        testArrayFormula2(assert, "ISREF",1,1,null,true);
    } );

	QUnit.test("Test: ISTEXT", function (assert) {
		ws.getRange2( "S7" ).setValue( "test" );

		oParser = new parserFormula( "ISTEXT(S7)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

		testArrayFormula2(assert, "ISTEXT",1,1);
	} );

    QUnit.test("Test: MOD", function (assert) {
        oParser = new parserFormula( "MOD(7,3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( "MOD(-10,5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "MOD(-9,5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( "MOD(-8,5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "MOD(-7,5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "MOD(-6,5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );

		oParser = new parserFormula( "MOD(-5,5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "MOD(10,5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "MOD(9,5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );

		oParser = new parserFormula( "MOD(8,5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "MOD(15,5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "MOD(15,0)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );

		testArrayFormula2(assert, "MOD", 2, 2);
    } );

    QUnit.test("Test: rename sheet #1", function (assert) {
		wb.dependencyFormulas.unlockRecal();
		ws.getRange2( "S95" ).setValue( "2" );
		ws.getRange2( "S100" ).setValue( "=" + wb.getWorksheet( 0 ).getName() + "!S95" );
		ws.setName( "SheetTmp" );
        assert.strictEqual( ws.getCell2( "S100" ).getFormula(), ws.getName() + "!S95" );
		wb.dependencyFormulas.lockRecal();
    } );

    QUnit.test("Test: wrong ref", function (assert) {
        oParser = new parserFormula( "1+XXX1", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NAME?" );
    } );

    QUnit.test("Test: \"CODE\"", function (assert) {
        oParser = new parserFormula( "CODE(\"abc\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 97 );

		oParser = new parserFormula( "CODE(TRUE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 84 );

        testArrayFormula2(assert, "CODE", 1, 1);
    } );

    QUnit.test("Test: \"CHAR\"", function (assert) {
        oParser = new parserFormula( "CHAR(97)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "a" );

        testArrayFormula2(assert, "CHAR", 1, 1);
    } );

    QUnit.test("Test: \"CHAR(CODE())\"", function (assert) {
        oParser = new parserFormula( "CHAR(CODE(\"A\"))", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "A" );
    } );

    QUnit.test("Test: \"PROPER\"", function (assert) {

        oParser = new parserFormula( "PROPER(\"2-cent's worth\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "2-Cent'S Worth" );

        oParser = new parserFormula( "PROPER(\"76BudGet\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "76Budget" );

        oParser = new parserFormula( "PROPER(\"this is a TITLE\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "This Is A Title" );

		oParser = new parserFormula( 'PROPER(TRUE)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "True" );

        testArrayFormula2(assert, "PROPER", 1, 1);
    } );

    QUnit.test("Test: \"GCD\"", function (assert) {
        oParser = new parserFormula( "GCD(10,100,50)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 10 );
        oParser = new parserFormula( "GCD(24.6,36.2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 12 );
        oParser = new parserFormula( "GCD(-1,39,52)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        testArrayFormula2(assert, "GCD", 1, 8, null, true);
    } );

    QUnit.test("Test: \"FIXED\"", function (assert) {
        oParser = new parserFormula( "FIXED(1234567,-3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "1,235,000" );
        oParser = new parserFormula( "FIXED(.555555,10)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "0.5555550000" );
        oParser = new parserFormula( "FIXED(1234567.555555,4,TRUE)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "1234567.5556" );
        oParser = new parserFormula( "FIXED(1234567)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "1,234,567.00" );

		testArrayFormula2(assert, "FIXED", 2, 3);
    } );

    QUnit.test("Test: \"REPLACE\"", function (assert) {

        oParser = new parserFormula( "REPLACE(\"abcdefghijk\",3,4,\"XY\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "abXYghijk" );

        oParser = new parserFormula( "REPLACE(\"abcdefghijk\",3,1,\"12345\")", "B2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "ab12345defghijk" );

        oParser = new parserFormula( "REPLACE(\"abcdefghijk\",15,4,\"XY\")", "C2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "abcdefghijkXY" );

        testArrayFormula2(assert, "REPLACE", 4, 4);
    } );

    QUnit.test("Test: \"SEARCH\"", function (assert) {

        oParser = new parserFormula( "SEARCH(\"~*\",\"abc*dEF\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4 );

        oParser = new parserFormula( "SEARCH(\"~\",\"abc~dEF\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4 );

        oParser = new parserFormula( "SEARCH(\"de\",\"abcdEF\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4 );

        oParser = new parserFormula( "SEARCH(\"?c*e\",\"abcdEF\")", "B2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2 );

        oParser = new parserFormula( "SEARCH(\"de\",\"dEFabcdEF\",3)", "C2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 7 );

        oParser = new parserFormula( "SEARCH(\"de\",\"dEFabcdEF\",30)", "C2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		testArrayFormula2(assert, "SEARCH", 2, 3);
    } );

    QUnit.test("Test: \"SUBSTITUTE\"", function (assert) {

        oParser = new parserFormula( "SUBSTITUTE(\"abcaAabca\",\"a\",\"xx\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "xxbcxxAxxbcxx" );

        oParser = new parserFormula( "SUBSTITUTE(\"abcaaabca\",\"a\",\"xx\")", "B2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "xxbcxxxxxxbcxx" );

        oParser = new parserFormula( "SUBSTITUTE(\"abcaaabca\",\"a\",\"\",10)", "C2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "bcbc" );

        oParser = new parserFormula( "SUBSTITUTE(\"abcaaabca\",\"a\",\"xx\",3)", "C2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "abcaxxabca" );

        testArrayFormula2(assert, "SUBSTITUTE", 3, 4);
    } );

	QUnit.test("Test: \"SHEET\"", function (assert) {

		oParser = new parserFormula( "SHEET(Hi_Temps)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NAME?" );

		testArrayFormula2(assert, "SHEET", 1, 1, null, true);
	} );

	QUnit.test("Test: \"SHEETS\"", function (assert) {

		oParser = new parserFormula( "SHEETS(Hi_Temps)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NAME?" );

		oParser = new parserFormula( "SHEETS()", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		testArrayFormula2(assert, "SHEETS", 1, 1, null, true);
	} );

    QUnit.test("Test: \"TRIM\"", function (assert) {

        oParser = new parserFormula( "TRIM(\"     abc         def      \")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "abc def" );

		oParser = new parserFormula( "TRIM(\" First Quarter Earnings \")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "First Quarter Earnings" );

		testArrayFormula2(assert, "TRIM", 1, 1);
    } );

	QUnit.test("Test: \"TRIMMEAN\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "4" );
		ws.getRange2( "A3" ).setValue( "5" );
		ws.getRange2( "A4" ).setValue( "6" );
		ws.getRange2( "A5" ).setValue( "7" );
		ws.getRange2( "A6" ).setValue( "2" );
		ws.getRange2( "A7" ).setValue( "3" );
		ws.getRange2( "A8" ).setValue( "4" );
		ws.getRange2( "A9" ).setValue( "5" );
		ws.getRange2( "A10" ).setValue( "1" );
		ws.getRange2( "A11" ).setValue( "2" );
		ws.getRange2( "A12" ).setValue( "3" );

		oParser = new parserFormula( "TRIMMEAN(A2:A12,0.2)", "A1", ws );
		assert.ok( oParser.parse(), "TRIMMEAN(A2:A12,0.2)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(3) - 0, 3.778, "TRIMMEAN(A2:A12,0.2)" );

		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "TRIMMEAN", 2, 2);
	} );

    QUnit.test("Test: \"DOLLAR\"", function (assert) {

        oParser = new parserFormula( "DOLLAR(1234.567)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "$1,234.57" );

        oParser = new parserFormula( "DOLLAR(1234.567,-2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "$1,200" );

        oParser = new parserFormula( "DOLLAR(-1234.567,4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "($1,234.5670)" );

        testArrayFormula2(assert, "DOLLAR", 2, 2);
    } );

	QUnit.test("Test: \"EXACT\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "word" );
		ws.getRange2( "A3" ).setValue( "Word" );
		ws.getRange2( "A4" ).setValue( "w ord" );
		ws.getRange2( "B2" ).setValue( "word" );
		ws.getRange2( "B3" ).setValue( "word" );
		ws.getRange2( "B4" ).setValue( "word" );

		oParser = new parserFormula( "EXACT(A2,B2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

		oParser = new parserFormula( "EXACT(A3,B3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE" );

		oParser = new parserFormula( "EXACT(A4,B4)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE" );

		oParser = new parserFormula( "EXACT(TRUE,TRUE)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

		oParser = new parserFormula( 'EXACT("TRUE",TRUE)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

		oParser = new parserFormula( 'EXACT("TRUE","TRUE")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

		oParser = new parserFormula( 'EXACT("true",TRUE)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE" );

		testArrayFormula2(assert, "EXACT", 2, 2);
	} );

	QUnit.test("Test: \"LEFT\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "Sale Price" );
		ws.getRange2( "A3" ).setValue( "Sweden" );


		oParser = new parserFormula( "LEFT(A2,4)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "Sale" );

		oParser = new parserFormula( "LEFT(A3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "S" );

		testArrayFormula2(assert, "LEFT", 1, 2);
	} );

	QUnit.test("Test: \"LEN\"", function (assert) {

		ws.getRange2( "A201" ).setValue( "Phoenix, AZ" );
		ws.getRange2( "A202" ).setValue( "" );
		ws.getRange2( "A203" ).setValue( "     One   " );

		oParser = new parserFormula( "LEN(A201)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 11 );

		oParser = new parserFormula( "LEN(A202)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "LEN(A203)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 11 );

		oParser = new parserFormula( 'LEN(TRUE)', "A2", ws );
		assert.ok( oParser.parse(), 'LEN(TRUE)' );
		assert.strictEqual( oParser.calculate().getValue(), 4, 'LEN(TRUE)');

		testArrayFormula2(assert, "LEN", 1, 1);
	} );

	QUnit.test("Test: \"REPT\"", function (assert) {

		oParser = new parserFormula( 'REPT("*-", 3)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "*-*-*-" );

		oParser = new parserFormula( 'REPT("-",10)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "----------" );

		testArrayFormula2(assert, "REPT", 2, 2);
	} );

	QUnit.test("Test: \"RIGHT\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "Sale Price" );
		ws.getRange2( "A3" ).setValue( "Stock Number" );

		oParser = new parserFormula( "RIGHT(A2,5)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "Price" );

		oParser = new parserFormula( "RIGHT(A3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "r" );

		testArrayFormula2(assert, "RIGHT", 1, 2);
	} );


	QUnit.test("Test: \"VALUE\"", function (assert) {

        oParser = new parserFormula( "VALUE(\"123.456\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 123.456 );

        oParser = new parserFormula( "VALUE(\"$1,000\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1000 );

        oParser = new parserFormula( "VALUE(\"23-Mar-2002\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 37338 );

        oParser = new parserFormula( "VALUE(\"03-26-2006\")", "A2", ws );
        assert.ok( oParser.parse() );

        if ( AscCommon.bDate1904 )
            assert.strictEqual( oParser.calculate().getValue(), 37340 );
        else
            assert.strictEqual( oParser.calculate().getValue(), 38802 );

        oParser = new parserFormula( "VALUE(\"16:48:00\")-VALUE(\"12:17:12\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), AscCommon.g_oFormatParser.parse( "16:48:00" ).value - AscCommon.g_oFormatParser.parse( "12:17:12" ).value );

        testArrayFormula2(assert, "value", 1, 1);
    } );

	QUnit.test("Test: \"DATE\"", function (assert) {

		testArrayFormula2(assert, "DATE", 3, 3);
	} );

    QUnit.test("Test: \"DATEVALUE\"", function (assert) {

        oParser = new parserFormula( "DATEVALUE(\"10-10-2010 10:26\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 40461 );

        oParser = new parserFormula( "DATEVALUE(\"10-10-2010 10:26\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 40461 );

        tmp = ws.getRange2( "A7" );
        tmp.setNumFormat('@');
        tmp.setValue( "3-Mar" );
        oParser = new parserFormula( "DATEVALUE(A7)", "A2", ws );
        assert.ok( oParser.parse() );
        var d = new cDate();
        d.setUTCMonth(2);
        d.setUTCDate(3);
        assert.strictEqual( oParser.calculate().getValue(), d.getExcelDate() );

        oParser = new parserFormula( "DATEVALUE(\"$1,000\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        oParser = new parserFormula( "DATEVALUE(\"23-Mar-2002\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 37338 );

        oParser = new parserFormula( "DATEVALUE(\"03-26-2006\")", "A2", ws );
        assert.ok( oParser.parse() );

        if ( AscCommon.bDate1904 )
            assert.strictEqual( oParser.calculate().getValue(), 37340 );
        else
            assert.strictEqual( oParser.calculate().getValue(), 38802 );

        testArrayFormula(assert, "DATEVALUE");
    } );

    QUnit.test("Test: \"EDATE\"", function (assert) {

        if ( !AscCommon.bDate1904 ) {
            oParser = new parserFormula( "EDATE(DATE(2006,1,31),5)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 38898 );

            oParser = new parserFormula( "EDATE(DATE(2004,2,29),12)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 38411 );

            ws.getRange2( "A7" ).setValue( "02-28-2004" );
            oParser = new parserFormula( "EDATE(A7,12)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 38411 );

            oParser = new parserFormula( "EDATE(DATE(2004,1,15),-23)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 37302 );
        }
        else {
            oParser = new parserFormula( "EDATE(DATE(2006,1,31),5)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 37436 );

            oParser = new parserFormula( "EDATE(DATE(2004,2,29),12)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 36949 );

            ws.getRange2( "A7" ).setValue( "02-28-2004" );
            oParser = new parserFormula( "EDATE(A7,12)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 36949 );

            oParser = new parserFormula( "EDATE(DATE(2004,1,15),-23)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 35840 );
        }

        testArrayFormula2(assert, "EDATE", 2, 2, true, null);
    } );

    QUnit.test("Test: \"EOMONTH\"", function (assert) {

        if ( !AscCommon.bDate1904 ) {
            oParser = new parserFormula( "EOMONTH(DATE(2006,1,31),5)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 38898 );

            oParser = new parserFormula( "EOMONTH(DATE(2004,2,29),12)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 38411 );

            ws.getRange2( "A7" ).setValue( "02-28-2004" );
            oParser = new parserFormula( "EOMONTH(A7,12)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 38411 );

            oParser = new parserFormula( "EOMONTH(DATE(2004,1,15),-23)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 37315 );
        }
        else {
            oParser = new parserFormula( "EOMONTH(DATE(2006,1,31),5)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 37436 );

            oParser = new parserFormula( "EOMONTH(DATE(2004,2,29),12)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 36949 );

            ws.getRange2( "A7" ).setValue( "02-28-2004" );
            oParser = new parserFormula( "EOMONTH(A7,12)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 36949 );

            oParser = new parserFormula( "EOMONTH(DATE(2004,1,15),-23)", "A2", ws );
            assert.ok( oParser.parse() );
            assert.strictEqual( oParser.calculate().getValue(), 35853 );
        }

		testArrayFormula2(assert, "EOMONTH", 2, 2, true, null);
    } );

	QUnit.test("Test: \"EVEN\"", function (assert) {

		oParser = new parserFormula( "EVEN(1.5)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "EVEN(3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );

		oParser = new parserFormula( "EVEN(2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "EVEN(-1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -2 );

		testArrayFormula(assert, "EVEN");

	} );


	QUnit.test("Test: \"NETWORKDAYS\"", function (assert) {

        oParser = new parserFormula( "NETWORKDAYS(DATE(2006,1,1),DATE(2006,1,31))", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 22 );

        oParser = new parserFormula( "NETWORKDAYS(DATE(2006,1,31),DATE(2006,1,1))", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -22 );

        oParser = new parserFormula( "NETWORKDAYS(DATE(2006,1,1),DATE(2006,2,1),{\"01-02-2006\",\"01-16-2006\"})", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 21 );

        testArrayFormula2(assert, "NETWORKDAYS", 2, 3, true, null);
    } );

	QUnit.test("Test: \"NETWORKDAYS.INTL\"", function (assert) {

		var formulaStr = "NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))";
	    oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 22, formulaStr );

		formulaStr = "NETWORKDAYS.INTL(DATE(2006,2,28),DATE(2006,1,31))";
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), -21, formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 22, formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),17,{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 26, formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"1111111",{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 0, formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"0010001",{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 20, formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"0000000",{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 30, formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"19",{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!", formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),19,{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(DATE(1901,1,1),DATE(2006,2,1),"0000000",{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 38381, formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(DATE(1901,1,1),DATE(2006,2,1),17,{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 32898, formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,11)';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 8490, formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,1)';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 7075, formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,2)';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 7075, formulaStr );

		//TODO посмотреть почему неверно считается
		//проблема повторяется с новым и со старым вариантом реализации NETWORKDAYS.INTL

		/*formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,5)';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 7074, formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,5,{123,1000})';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 7073, formulaStr );*/
	} );

	QUnit.test("Test: \"N\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "7" );
		ws.getRange2( "A3" ).setValue( "Even" );
		ws.getRange2( "A4" ).setValue( "TRUE" );
		ws.getRange2( "A5" ).setValue( "4/17/2011" );

		oParser = new parserFormula( "N(A2)", "A7", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 7 );

		oParser = new parserFormula( "N(A3)", "A7", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "N(A4)", "A7", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( "N(A5)", "A7", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 40650 );

		oParser = new parserFormula( 'N("7")', "A7", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "N", 1, 1);
	} );

    QUnit.test("Test: \"SUMIF\"", function (assert) {

        ws.getRange2( "A2" ).setValue( "100000" );
        ws.getRange2( "A3" ).setValue( "200000" );
        ws.getRange2( "A4" ).setValue( "300000" );
        ws.getRange2( "A5" ).setValue( "400000" );

        ws.getRange2( "B2" ).setValue( "7000" );
        ws.getRange2( "B3" ).setValue( "14000" );
        ws.getRange2( "B4" ).setValue( "21000" );
        ws.getRange2( "B5" ).setValue( "28000" );

        ws.getRange2( "C2" ).setValue( "250000" );

        oParser = new parserFormula( "SUMIF(A2:A5,\">160000\",B2:B5)", "A7", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 63000 );

        oParser = new parserFormula( "SUMIF(A2:A5,\">160000\")", "A8", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 900000 );

        oParser = new parserFormula( "SUMIF(A2:A5,300000,B2:B5)", "A9", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 21000 );

        oParser = new parserFormula( "SUMIF(A2:A5,\">\" & C2,B2:B5)", "A10", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 49000 );

		oParser = new parserFormula( "SUMIF(A2,\">160000\",B2:B5)", "A7", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "SUMIF(A3,\">160000\",B2:B5)", "A7", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 7000 );

		oParser = new parserFormula( "SUMIF(A4,\">160000\",B4:B5)", "A7", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 21000 );

		oParser = new parserFormula( "SUMIF(A4,\">160000\")", "A7", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 300000);


		ws.getRange2( "A12" ).setValue( "Vegetables" );
        ws.getRange2( "A13" ).setValue( "Vegetables" );
        ws.getRange2( "A14" ).setValue( "Fruits" );
        ws.getRange2( "A15" ).setValue( "" );
        ws.getRange2( "A16" ).setValue( "Vegetables" );
        ws.getRange2( "A17" ).setValue( "Fruits" );

        ws.getRange2( "B12" ).setValue( "Tomatoes" );
        ws.getRange2( "B13" ).setValue( "Celery" );
        ws.getRange2( "B14" ).setValue( "Oranges" );
        ws.getRange2( "B15" ).setValue( "Butter" );
        ws.getRange2( "B16" ).setValue( "Carrots" );
        ws.getRange2( "B17" ).setValue( "Apples" );

        ws.getRange2( "C12" ).setValue( "2300" );
        ws.getRange2( "C13" ).setValue( "5500" );
        ws.getRange2( "C14" ).setValue( "800" );
        ws.getRange2( "C15" ).setValue( "400" );
        ws.getRange2( "C16" ).setValue( "4200" );
        ws.getRange2( "C17" ).setValue( "1200" );

        oParser = new parserFormula( "SUMIF(A12:A17,\"Fruits\",C12:C17)", "A19", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2000 );

        oParser = new parserFormula( "SUMIF(A12:A17,\"Vegetables\",C12:C17)", "A20", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 12000 );

        oParser = new parserFormula( "SUMIF(B12:B17,\"*es\",C12:C17)", "A21", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4300 );

        oParser = new parserFormula( "SUMIF(A12:A17,\"\",C12:C17)", "A22", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 400 );

    } );

    QUnit.test("Test: \"SUMIFS\"", function (assert) {

        ws.getRange2( "A2" ).setValue( "5" );
        ws.getRange2( "A3" ).setValue( "4" );
        ws.getRange2( "A4" ).setValue( "15" );
        ws.getRange2( "A5" ).setValue( "3" );
        ws.getRange2( "A6" ).setValue( "22" );
        ws.getRange2( "A7" ).setValue( "12" );
        ws.getRange2( "A8" ).setValue( "10" );
        ws.getRange2( "A9" ).setValue( "33" );

        ws.getRange2( "B2" ).setValue( "Apples" );
        ws.getRange2( "B3" ).setValue( "Apples" );
        ws.getRange2( "B4" ).setValue( "Artichokes" );
        ws.getRange2( "B5" ).setValue( "Artichokes" );
        ws.getRange2( "B6" ).setValue( "Bananas" );
        ws.getRange2( "B7" ).setValue( "Bananas" );
        ws.getRange2( "B8" ).setValue( "Carrots" );
        ws.getRange2( "B9" ).setValue( "Carrots" );

        ws.getRange2( "C2" ).setValue( "Tom" );
        ws.getRange2( "C3" ).setValue( "Sarah" );
        ws.getRange2( "C4" ).setValue( "Tom" );
        ws.getRange2( "C5" ).setValue( "Sarah" );
        ws.getRange2( "C6" ).setValue( "Tom" );
        ws.getRange2( "C7" ).setValue( "Sarah" );
        ws.getRange2( "C8" ).setValue( "Tom" );
        ws.getRange2( "C9" ).setValue( "Sarah" );

        oParser = new parserFormula( "SUMIFS(A2:A9, B2:B9, \"=A*\", C2:C9, \"Tom\")", "A10", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 20 );

        oParser = new parserFormula( "SUMIFS(A2:A9, B2:B9, \"<>Bananas\", C2:C9, \"Tom\")", "A11", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 30 );

		oParser = new parserFormula( "SUMIFS(D:D,E:E,$H2)", "A11", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "SUMIFS(C:D,E:E,$H2)", "A11", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );
    } );

	QUnit.test("Test: \"MAXIFS\"", function (assert) {

		ws.getRange2( "AAA2" ).setValue( "10" );
		ws.getRange2( "AAA3" ).setValue( "1" );
		ws.getRange2( "AAA4" ).setValue( "100" );
		ws.getRange2( "AAA5" ).setValue( "1" );
		ws.getRange2( "AAA6" ).setValue( "1" );
		ws.getRange2( "AAA7" ).setValue( "50" );

		ws.getRange2( "BBB2" ).setValue( "b" );
		ws.getRange2( "BBB3" ).setValue( "a" );
		ws.getRange2( "BBB4" ).setValue( "a" );
		ws.getRange2( "BBB5" ).setValue( "b" );
		ws.getRange2( "BBB6" ).setValue( "a" );
		ws.getRange2( "BBB7" ).setValue( "b" );

		ws.getRange2( "DDD2" ).setValue( "100" );
		ws.getRange2( "DDD3" ).setValue( "100" );
		ws.getRange2( "DDD4" ).setValue( "200" );
		ws.getRange2( "DDD5" ).setValue( "300" );
		ws.getRange2( "DDD6" ).setValue( "100" );
		ws.getRange2( "DDD7" ).setValue( "400" );

		oParser = new parserFormula( 'MAXIFS(AAA2:AAA7,BBB2:BBB7,"b",DDD2:DDD7,">100")', "A22", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 50 );

		oParser = new parserFormula( 'MAXIFS(AAA2:AAA6,BBB2:BBB6,"a",DDD2:DDD6,">200")', "A22", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		testArrayFormulaEqualsValues(assert, "1,3.123,-4,#N/A;2,4,5,#N/A;#N/A,#N/A,#N/A,#N/A","MAXIFS(A1:C2,A1:C2,A1:C2,A1:C2, A1:C2,A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert, "1,0,0,#N/A;0,0,0,#N/A;#N/A,#N/A,#N/A,#N/A","MAXIFS(A1:C2,A1:C2,A1:A1,A1:C2,A1:C2,A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert, "1,0,0,#N/A;2,0,0,#N/A;#N/A,#N/A,#N/A,#N/A","MAXIFS(A1:C2,A1:C2,A1:A2,A1:C2,A1:C2,A1:C2,A1:C2)");
	} );

	QUnit.test("Test: \"MINIFS\"", function (assert) {

		ws.getRange2( "AAA2" ).setValue( "10" );
		ws.getRange2( "AAA3" ).setValue( "1" );
		ws.getRange2( "AAA4" ).setValue( "100" );
		ws.getRange2( "AAA5" ).setValue( "1" );
		ws.getRange2( "AAA6" ).setValue( "1" );
		ws.getRange2( "AAA7" ).setValue( "50" );

		ws.getRange2( "BBB2" ).setValue( "b" );
		ws.getRange2( "BBB3" ).setValue( "a" );
		ws.getRange2( "BBB4" ).setValue( "a" );
		ws.getRange2( "BBB5" ).setValue( "b" );
		ws.getRange2( "BBB6" ).setValue( "a" );
		ws.getRange2( "BBB7" ).setValue( "b" );

		ws.getRange2( "DDD2" ).setValue( "100" );
		ws.getRange2( "DDD3" ).setValue( "100" );
		ws.getRange2( "DDD4" ).setValue( "200" );
		ws.getRange2( "DDD5" ).setValue( "300" );
		ws.getRange2( "DDD6" ).setValue( "100" );
		ws.getRange2( "DDD7" ).setValue( "400" );

		oParser = new parserFormula( 'MINIFS(AAA2:AAA7,BBB2:BBB7,"b",DDD2:DDD7,">100")', "A22", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( 'MINIFS(AAA2:AAA6,BBB2:BBB6,"a",DDD2:DDD6,">200")', "A22", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		testArrayFormulaEqualsValues(assert, "1,3.123,-4,#N/A;2,4,5,#N/A;#N/A,#N/A,#N/A,#N/A","MINIFS(A1:C2,A1:C2,A1:C2,A1:C2, A1:C2,A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert, "1,0,0,#N/A;0,0,0,#N/A;#N/A,#N/A,#N/A,#N/A","MINIFS(A1:C2,A1:C2,A1:A1,A1:C2,A1:C2,A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert, "1,0,0,#N/A;2,0,0,#N/A;#N/A,#N/A,#N/A,#N/A","MINIFS(A1:C2,A1:C2,A1:A2,A1:C2,A1:C2,A1:C2,A1:C2)");
	} );

    QUnit.test("Test: \"TEXT\"", function (assert) {
        var culturelciddefault =  AscCommon.g_oDefaultCultureInfo.LCID;

        oParser = new parserFormula( "TEXT(1234.567,\"$0.00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "$1234.57" );

        oParser = new parserFormula( "TEXT(0.125,\"0.0%\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "12.5%" );

        oParser = new parserFormula( "TEXT(123123,\"hh:mmm:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:Feb:00" );

        oParser = new parserFormula( "TEXT(123123,\"hh:mmmm:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:February:00" );

        oParser = new parserFormula( "TEXT(123123,\"hh:mmmmm:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:F:00" );

        testArrayFormula2(assert, "TEXT", 2, 2);

        //____________________________________en_____________________________________________
        AscCommon.setCurrentCultureInfo(1025);
        oParser = new parserFormula( "TEXT(123,\"yy-mm-dd\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"yy-MM-dd\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"hh:MM:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"hh:mm:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"general\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123" );

        oParser = new parserFormula( "TEXT(123,\"0.00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123.00" );

		oParser = new parserFormula( "TEXT(123123123,\"#,###\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123,123,123" );

		oParser = new parserFormula( "TEXT(123123123,\"#,##0.00\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123,123,123.00" );

		//todo
		// oParser = new parserFormula( "TEXT(123123123,\"###0,.00\")", "A2", ws );
		// assert.ok( oParser.parse() );
		// assert.strictEqual( oParser.calculate().getValue(), "123123.12" );
		//
		// oParser = new parserFormula( "TEXT(123123123,\"###0,.00,\")", "A2", ws );
		// assert.ok( oParser.parse() );
		// assert.strictEqual( oParser.calculate().getValue(), "123.12" );

		oParser = new parserFormula( "TEXT(123123123,\"###0.,0,0\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123123123.00" );

		oParser = new parserFormula( "TEXT(123123123,\"#,,,#,#,0.,,0,0\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123,123,123.00" );

		oParser = new parserFormula( "TEXT(123123123,\"0,0,, q,w,,,w,,0,,0\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123,123,1 q,w,w,23" );

		oParser = new parserFormula( "TEXT(123123123,\"00,,q,,,q00\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "1231231q,q23" );

		oParser = new parserFormula( "TEXT(123123123,\"0,q,w,,\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123123q,w," );

		oParser = new parserFormula( "TEXT(123123123,\"#,,\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123" );

		oParser = new parserFormula( "TEXT(123123123,\",,,#\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), ",123123123" );

		oParser = new parserFormula( "TEXT(123123123,\"###0,.\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123123." );

		oParser = new parserFormula( "TEXT(123123123,\"###0.,\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123123." );

		oParser = new parserFormula( "TEXT(123123123,\"###0.,q\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123123.q" );

		oParser = new parserFormula( "TEXT(123123123,\"###0,.\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123123." );

		oParser = new parserFormula( "TEXT(123123123,\"###0 ,\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123123123 ," );

		oParser = new parserFormula( "TEXT(123,\"hh qq\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00 qq" );

        oParser = new parserFormula( "TEXT(123,\"dd hh\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "02 00" );

        AscCommon.setCurrentCultureInfo(culturelciddefault);

        //__________________________________fi________________________________________________
        AscCommon.setCurrentCultureInfo(1035);
        oParser = new parserFormula( "TEXT(123,\"vv-kk-pp\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"vv-mm-pp\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-00-02" );

        oParser = new parserFormula( "TEXT(123,\"tt.mm.ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00.00.00" );

        oParser = new parserFormula( "TEXT(123,\"tt.MM.ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00.00.00" );

        oParser = new parserFormula( "TEXT(123,\"tt.kk.ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00.05.00" );

        oParser = new parserFormula( "TEXT(125,\"yleinen\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "125" );

        oParser = new parserFormula( "TEXT(125,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "125,00" );

        oParser = new parserFormula( "TEXT(125125,\"# ###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "125 125" );

        oParser = new parserFormula( "TEXT(123,\"pp tt\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "02 00" );

        oParser = new parserFormula( "TEXT(123,\"p t\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "2 0" );

        oParser = new parserFormula( "TEXT(123,\"tt qq\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00 qq" );

        AscCommon.setCurrentCultureInfo(culturelciddefault);

        //__________________________________fy________________________________________________
        AscCommon.setCurrentCultureInfo(1043);
        oParser = new parserFormula( "TEXT(123,\"jj-mm-dd\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"uu:mm:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(126,\"standaard\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "126" );

        oParser = new parserFormula( "TEXT(126,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "126,00" );

        oParser = new parserFormula( "TEXT(126126,\"#.###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "126.126" );

        oParser = new parserFormula( "TEXT(123,\"dd uu\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "02 00" );

        oParser = new parserFormula( "TEXT(123,\"d u\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "2 0" );

        oParser = new parserFormula( "TEXT(123,\"uu qq\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00 qq" );

        AscCommon.setCurrentCultureInfo(culturelciddefault);

        //__________________________________es________________________________________________
        AscCommon.setCurrentCultureInfo(3082);
        oParser = new parserFormula( "TEXT(123,\"aa-mm-dd\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"estándar\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123" );

        oParser = new parserFormula( "TEXT(123,\"hh:mm:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123,00" );

        oParser = new parserFormula( "TEXT(123123123,\"#.###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123.123.123" );

        oParser = new parserFormula( "TEXT(123,\"dd hh\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "02 00" );

        oParser = new parserFormula( "TEXT(123,\"hh qq\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00 qq" );

        AscCommon.setCurrentCultureInfo(culturelciddefault);

        //___________________________________ru______________________________________________
        AscCommon.setCurrentCultureInfo(1049);
        oParser = new parserFormula( "TEXT(123,\"гг-ММ-дд\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"чч:ММ:сс\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

         oParser = new parserFormula( "TEXT(123,\"чч:мм:сс\")", "A2", ws );
         assert.ok( oParser.parse() );
         assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

         oParser = new parserFormula( "TEXT(123,\"гг-мм-дд\")", "A2", ws );
         assert.ok( oParser.parse() );
         assert.strictEqual( oParser.calculate().getValue(), "00-00-02" );

        oParser = new parserFormula( "TEXT(123,\"основной\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123" );

        oParser = new parserFormula( "TEXT(123,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123,00" );

        oParser = new parserFormula( "TEXT(123123123,\"# ###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123 123 123" );

		oParser = new parserFormula( "TEXT(123123123;\"# ##0,00\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123 123 123,00" );

		//todo
		// oParser = new parserFormula( "TEXT(123123123,\"###0 ,00\")", "A2", ws );
		// assert.ok( oParser.parse() );
		// assert.strictEqual( oParser.calculate().getValue(), "123123,12" );
		//
		// oParser = new parserFormula( "TEXT(123123123,\"###0 ,00 \")", "A2", ws );
		// assert.ok( oParser.parse() );
		// assert.strictEqual( oParser.calculate().getValue(), "123,12" );

		oParser = new parserFormula( "TEXT(123123123,\"###0, 0 0\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123123123,00" );

		oParser = new parserFormula( "TEXT(123123123,\"#   # # 0,  0 0\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123 123 123,00" );

		oParser = new parserFormula( "TEXT(123123123,\"0 0   q w   w  0  0\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123 123 1q w w 23" );

		oParser = new parserFormula( "TEXT(123123123,\"00  q   q00\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "1231231q q23" );

		oParser = new parserFormula( "TEXT(123123123,\"0 q w  \")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123123q w " );

		oParser = new parserFormula( "TEXT(123123123,\"#  \")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123" );

		oParser = new parserFormula( "TEXT(123123123,\"   #\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), " 123123123" );

		oParser = new parserFormula( "TEXT(123123123,\"###0 ,\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123123," );

		oParser = new parserFormula( "TEXT(123123123,\"###0, \")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123123," );

        oParser = new parserFormula( "TEXT(123,\"дд чч\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "02 00" );

        oParser = new parserFormula( "TEXT(123,\"чч qq\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00 qq" );

         oParser = new parserFormula( "TEXT(123,\"ММ мм\")", "A2", ws );
         assert.ok( oParser.parse() );
         assert.strictEqual( oParser.calculate().getValue(), "05 00" );

        AscCommon.setCurrentCultureInfo(culturelciddefault);

        //__________________________________fr________________________________________________
        AscCommon.setCurrentCultureInfo(1036);
        oParser = new parserFormula( "TEXT(123,\"aa-mm-jj\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"hh:mm:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"standard\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123" );

        oParser = new parserFormula( "TEXT(123,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123,00" );

        oParser = new parserFormula( "TEXT(123123123,\"# ###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123 123 123" );

        oParser = new parserFormula( "TEXT(123,\"jj hh\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "02 00" );

        oParser = new parserFormula( "TEXT(123,\"j h\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "2 0" );

        oParser = new parserFormula( "TEXT(123,\"hh qq\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00 qq" );

        AscCommon.setCurrentCultureInfo(culturelciddefault);

        //_______________________________de___________________________________________________
        AscCommon.setCurrentCultureInfo(1031);
        oParser = new parserFormula( "TEXT(123,\"jj-MM-tt\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"hh:mm:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"hh:MM:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"jj-mm-tt\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-00-02" );

        oParser = new parserFormula( "TEXT(123,\"standard\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123" );

        oParser = new parserFormula( "TEXT(123,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123,00" );

        oParser = new parserFormula( "TEXT(123123123,\"#.###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123.123.123" );

        oParser = new parserFormula( "TEXT(123,\"t h\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "2 0" );

        oParser = new parserFormula( "TEXT(123,\"h qq\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "0 qq" );

        oParser = new parserFormula( "TEXT(123,\"s t\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "0 2" );

        AscCommon.setCurrentCultureInfo(culturelciddefault);

        //_______________________________it___________________________________________________
        AscCommon.setCurrentCultureInfo(1040);
        oParser = new parserFormula( "TEXT(123,\"aa-MM-gg\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"aa-mm-gg\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"hh:MM:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"hh:mm:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"standard\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123" );

        oParser = new parserFormula( "TEXT(123,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123,00" );

        oParser = new parserFormula( "TEXT(123123123,\"#.###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123.123.123" );

        oParser = new parserFormula( "TEXT(123,\"gg hh\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "02 00" );

        oParser = new parserFormula( "TEXT(123,\"hh qq\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00 qq" );

        AscCommon.setCurrentCultureInfo(culturelciddefault);

        //_______________________________da____________________________________________________
        AscCommon.setCurrentCultureInfo(1053);
        oParser = new parserFormula( "TEXT(123,\"åå-MM-dd\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"åå-mm-dd\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-00-02" );

        oParser = new parserFormula( "TEXT(123,\"tt:mm:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"tt:MM:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"standard\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123" );

        oParser = new parserFormula( "TEXT(123,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123,00" );

        oParser = new parserFormula( "TEXT(123,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123,00" );

        oParser = new parserFormula( "TEXT(123123123,\"# ###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123 123 123" );

        oParser = new parserFormula( "TEXT(123,\"dd tt\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "02 00" );

        oParser = new parserFormula( "TEXT(123,\"tt qq\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00 qq" );

        AscCommon.setCurrentCultureInfo(culturelciddefault);
        //_____________________________ch________________________________________________
        AscCommon.setCurrentCultureInfo(1028);
        oParser = new parserFormula( "TEXT(123,\"g/通用格式\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123" );

        oParser = new parserFormula( "TEXT(123,\"0.00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123.00" );

        oParser = new parserFormula( "TEXT(123,\"dd hh\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "02 00" );

        AscCommon.setCurrentCultureInfo(culturelciddefault);
        //_____________________________el________________________________________________
        AscCommon.setCurrentCultureInfo(1032);
        oParser = new parserFormula( "TEXT(123,\"εε-μμ-ηη\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"ωω:λλ:δδ\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"γενικός τύπος\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123" );

        oParser = new parserFormula( "TEXT(123,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123,00" );

        oParser = new parserFormula( "TEXT(123123123,\"#.###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123.123.123" );

        oParser = new parserFormula( "TEXT(123,\"ηη ωω\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "02 00" );

        oParser = new parserFormula( "TEXT(123,\"ωω qq\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00 qq" );

        oParser = new parserFormula( "TEXT(123,\"ηη qq\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "02 qq" );

        oParser = new parserFormula( "TEXT(123,\"δδ ηη\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00 02" );

        AscCommon.setCurrentCultureInfo(culturelciddefault);

        //_____________________________hu________________________________________________
        AscCommon.setCurrentCultureInfo(1038);
        oParser = new parserFormula( "TEXT(123,\"éé-hh-nn\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"óó:pp:mm\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"éé-pp-nn\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-00-02" );

        oParser = new parserFormula( "TEXT(123,\"óó:hh:mm\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:05:00" );

        oParser = new parserFormula( "TEXT(125,\"normál\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "125" );

        oParser = new parserFormula( "TEXT(125,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "125,00" );

        oParser = new parserFormula( "TEXT(123123123,\"# ###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123 123 123" );
        AscCommon.setCurrentCultureInfo(culturelciddefault);
        //_____________________________tr________________________________________________
        AscCommon.setCurrentCultureInfo(1055);
        oParser = new parserFormula( "TEXT(123,\"yy-aa-gg\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"ss:dd:nn\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"yy-dd-gg\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-00-02" );

        oParser = new parserFormula( "TEXT(123,\"ss:aa:nn\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:05:00" );

        oParser = new parserFormula( "TEXT(125,\"genel\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "125" );

        oParser = new parserFormula( "TEXT(125,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "125,00" );

        oParser = new parserFormula( "TEXT(123123123,\"#.###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123.123.123" );
        AscCommon.setCurrentCultureInfo(culturelciddefault);
        //_____________________________pl________________________________________________
        AscCommon.setCurrentCultureInfo(1045);
        oParser = new parserFormula( "TEXT(123,\"rr-mm-dd\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"gg:mm:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(125,\"standardowy\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "125" );

        oParser = new parserFormula( "TEXT(125,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "125,00" );

        oParser = new parserFormula( "TEXT(123123123,\"# ###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123 123 123" );
        AscCommon.setCurrentCultureInfo(culturelciddefault);
        //_____________________________cs________________________________________________
        AscCommon.setCurrentCultureInfo(1029);
        oParser = new parserFormula( "TEXT(123,\"rr-mm-dd\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00-05-02" );

        oParser = new parserFormula( "TEXT(123,\"hh:mm:ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00:00:00" );

        oParser = new parserFormula( "TEXT(123,\"vęeobecný\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123" );

        oParser = new parserFormula( "TEXT(123,\"0,00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123,00" );

        oParser = new parserFormula( "TEXT(123123123,\"# ###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123 123 123" );

        oParser = new parserFormula( "TEXT(123,\"rr ss\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00 00" );

        oParser = new parserFormula( "TEXT(123,\"dd hh\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "02 00" );

        oParser = new parserFormula( "TEXT(123,\"hh qq\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "00 qq" );

        AscCommon.setCurrentCultureInfo(culturelciddefault);
        //_____________________________ja________________________________________________
        AscCommon.setCurrentCultureInfo(1041);
        oParser = new parserFormula( "TEXT(124,\"G/標準\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "124" );

        oParser = new parserFormula( "TEXT(123,\"0.00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123.00" );

        oParser = new parserFormula( "TEXT(123123123,\"#,###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123,123,123" );
        //_____________________________ko________________________________________________
        AscCommon.setCurrentCultureInfo(1042);
        oParser = new parserFormula( "TEXT(124,\"g/표준\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "124" );

        oParser = new parserFormula( "TEXT(123,\"0.00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123.00" );

        oParser = new parserFormula( "TEXT(123123123,\"#,###\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "123,123,123" );
        AscCommon.setCurrentCultureInfo(culturelciddefault);
    } );
	QUnit.test("Test: \"TEXTJOIN\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "Tulsa" );
		ws.getRange2( "A3" ).setValue( "Seattle" );
		ws.getRange2( "A4" ).setValue( "Iselin" );
		ws.getRange2( "A5" ).setValue( "Fort Lauderdale" );
		ws.getRange2( "A6" ).setValue( "Tempe" );
		ws.getRange2( "A7" ).setValue( "end" );

		ws.getRange2( "B2" ).setValue( "OK" );
		ws.getRange2( "B3" ).setValue( "WA" );
		ws.getRange2( "B4" ).setValue( "NJ" );
		ws.getRange2( "B5" ).setValue( "FL" );
		ws.getRange2( "B6" ).setValue( "AZ" );
		ws.getRange2( "B7" ).setValue( "" );

		ws.getRange2( "C2" ).setValue( "74133" );
		ws.getRange2( "C3" ).setValue( "98109" );
		ws.getRange2( "C4" ).setValue( "8830" );
		ws.getRange2( "C5" ).setValue( "33309" );
		ws.getRange2( "C6" ).setValue( "85285" );
		ws.getRange2( "C7" ).setValue( "" );

		ws.getRange2( "D2" ).setValue( "US" );
		ws.getRange2( "D3" ).setValue( "US" );
		ws.getRange2( "D4" ).setValue( "US" );
		ws.getRange2( "D5" ).setValue( "US" );
		ws.getRange2( "D6" ).setValue( "US" );
		ws.getRange2( "D7" ).setValue( "" );

		ws.getRange2( "A9" ).setValue( "," );
		ws.getRange2( "B9" ).setValue( "," );
		ws.getRange2( "C9" ).setValue( "," );
		ws.getRange2( "D9" ).setValue( ";" );


		oParser = new parserFormula( "TEXTJOIN(A9:D9, TRUE, A2:D7)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "Tulsa,OK,74133,US;Seattle,WA,98109,US;Iselin,NJ,8830,US;Fort Lauderdale,FL,33309,US;Tempe,AZ,85285,US;end" );

		oParser = new parserFormula( "TEXTJOIN(A9:D9, FALSE, A2:D7)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "Tulsa,OK,74133,US;Seattle,WA,98109,US;Iselin,NJ,8830,US;Fort Lauderdale,FL,33309,US;Tempe,AZ,85285,US;end,,," );

		oParser = new parserFormula( "TEXTJOIN(A2:D5, 1, B6:D6)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "AZTulsa85285OKUS" );

		testArrayFormulaEqualsValues(assert, "113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,#N/A;113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,#N/A;#N/A,#N/A,#N/A,#N/A", "TEXTJOIN(A1:C2,A1:C2,A1:C2,A1:C2, A1:C2)");
		testArrayFormulaEqualsValues(assert, "113.1232-41224152113.1232-4122415,113.1232-41224152113.1232-4122415,113.1232-41224152113.1232-4122415,#N/A;113.1232-41224152113.1232-4122415,113.1232-41224152113.1232-4122415,113.1232-41224152113.1232-4122415,#N/A;#N/A,#N/A,#N/A,#N/A", "TEXTJOIN(A1:A2,A1:C2,A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert, "113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445;113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445;#N/A,#N/A,#N/A,#N/A", "TEXTJOIN(A1:C2,A1:A2,A1:C2,A1:C2,A1:C2,A1:C2)");
	} );


	QUnit.test("Test: \"WORKDAY\"", function (assert) {

		oParser = new parserFormula("WORKDAY(DATE(2006,1,1),0)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 38718);

		oParser = new parserFormula("WORKDAY(DATE(2006,1,1),10)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 38730);

		oParser = new parserFormula("WORKDAY(DATE(2006,1,1),-10)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 38705);

		oParser = new parserFormula("WORKDAY(DATE(2006,1,1),20,{\"1-2-2006\",\"1-16-2006\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 38748);

		oParser = new parserFormula("WORKDAY(DATE(2017,10,6),1,DATE(2017,10,9))", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43018);

		oParser = new parserFormula("WORKDAY(DATE(2017,10,7),1,DATE(2017,10,9))", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43018);

		oParser = new parserFormula("WORKDAY(DATE(2017,9,25),-1,DATE(2017,9,10))", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43000);

		oParser = new parserFormula("WORKDAY(DATE(2017,9,25),-1,DATE(2017,9,10))", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43000);

		oParser = new parserFormula("WORKDAY(DATE(2017,9,20),-1,DATE(2017,9,10))", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 42997);

		oParser = new parserFormula("WORKDAY(DATE(2017,10,2),-1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43007);

		oParser = new parserFormula("WORKDAY(DATE(2017,10,2),-1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43007);

		oParser = new parserFormula("WORKDAY(DATE(2017,10,3),-3)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43006);

		oParser = new parserFormula("WORKDAY(DATE(2017,10,4),-2)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43010);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,30),1,{\"5-1-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43222);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,30),2,{\"5-1-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43224);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,30),3,{\"5-1-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43227);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,30),1,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43224);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,30),3,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43228);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,29),1,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43220);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,29),2,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43224);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,29),3,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43227);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,29),-1,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43217);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,29),-2,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43216);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,29),0,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43219);

		oParser = new parserFormula("WORKDAY({1,2,3},{1,2})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula("WORKDAY({1,2,3},1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula("WORKDAY(1,{1,2})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula("WORKDAY({1,2,3},1.123)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula("WORKDAY({1,2,3},-1.123)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula("WORKDAY({1,2,3},5)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 6);

		oParser = new parserFormula("WORKDAY(1,15)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 20);

		/*oParser = new parserFormula("WORKDAY(1,50)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 69);

		oParser = new parserFormula("WORKDAY(1,60)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 83);

		oParser = new parserFormula("WORKDAY(1,61)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 86);*/

		//todo ms выдаёт ошибки
		/*ws.getRange2( "A101" ).setValue( "1" );
		ws.getRange2( "B101" ).setValue( "3.123" );
		ws.getRange2( "C101" ).setValue( "-4" );

		oParser = new parserFormula("WORKDAY(A101:B101,A101:B101)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("WORKDAY(A101,A101:B101)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("WORKDAY(A101:B101,A101)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("WORKDAY(A101,A101)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);*/
	});

	QUnit.test("Test: \"WORKDAY.INTL\"", function (assert) {

		oParser = new parserFormula( "WORKDAY.INTL(DATE(2012,1,1),30,0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		oParser = new parserFormula( "WORKDAY.INTL(DATE(2012,1,1),90,11)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 41013 );

		oParser = new parserFormula( 'TEXT(WORKDAY.INTL(DATE(2012,1,1),30,17),"m/dd/yyyy")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "2/05/2012" );

		oParser = new parserFormula( 'WORKDAY.INTL(151,8,"0000000")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 159 );

		oParser = new parserFormula( 'WORKDAY.INTL(151,8,"0000000")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 159 );

		oParser = new parserFormula( 'WORKDAY.INTL(159,8,"0011100")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 171 );

		oParser = new parserFormula( 'WORKDAY.INTL(151,-18,"0000000")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 133 );

		oParser = new parserFormula( 'WORKDAY.INTL(151,8,"1111111")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( 'WORKDAY.INTL(DATE(2006,1,1),20,1,{"1/2/2006","1/16/2006"})', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 38748 );

		oParser = new parserFormula( 'WORKDAY.INTL(DATE(2006,1,1),20,{"1/2/2006","1/16/2006"})', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		oParser = new parserFormula( 'WORKDAY.INTL(DATE(2006,1,1),-20,1,{"1/2/2006",,"1/16/2006"})', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 38691 );

	} );

	QUnit.test("Test: \"TIME\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "12" );
		ws.getRange2( "A3" ).setValue( "16" );

		ws.getRange2( "B2" ).setValue( "0" );
		ws.getRange2( "B3" ).setValue( "48" );

		ws.getRange2( "C2" ).setValue( "0" );
		ws.getRange2( "C3" ).setValue( "10" );

		oParser = new parserFormula( "TIME(A2,B2,C2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0.5 );

		oParser = new parserFormula( "TIME(A3,B3,C3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.7001157 );

		oParser = new parserFormula( "TIME(1,1,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0423727 );

		oParser = new parserFormula( "TIME(1.34,1,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0423727 );

		oParser = new parserFormula( "TIME(1.34,1.456,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0423727 );

		oParser = new parserFormula( "TIME(1.34,1.456,1.9)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0423727 );

		oParser = new parserFormula( "TIME(-1.34,1.456,1.9)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		testArrayFormula2(assert, "TIME", 3, 3);
	} );

    QUnit.test("Test: \"TIMEVALUE\"", function (assert) {
        oParser = new parserFormula( "timevalue(\"10:02:34\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( Math.abs( oParser.calculate().getValue() - 0.4184490740740740 ) < dif );

        oParser = new parserFormula( "timevalue(\"02-01-2006 10:15:29 AM\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( Math.abs( oParser.calculate().getValue() - 0.4274189814823330 ) < dif );

        oParser = new parserFormula( "timevalue(\"22:02\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( Math.abs( oParser.calculate().getValue() - 0.9180555555555560 ) < dif );

		testArrayFormula(assert, "TIMEVALUE");
    } );

	QUnit.test("Test: \"TYPE\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "Smith" );

		oParser = new parserFormula( "TYPE(A2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( 'TYPE("Mr. "&A2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( 'TYPE(2+A2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 16 );

		oParser = new parserFormula( '(2+A2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( 'TYPE({1,2;3,4})', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 64 );

		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "TYPE", 1, 1);
	} );

    QUnit.test("Test: \"DAYS360\"", function (assert) {

        oParser = new parserFormula( "DAYS360(DATE(2002,2,3),DATE(2005,5,31))", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1198 );

        oParser = new parserFormula( "DAYS360(DATE(2005,5,31),DATE(2002,2,3))", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -1197 );

        oParser = new parserFormula( "DAYS360(DATE(2002,2,3),DATE(2005,5,31),FALSE)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1198 );

        oParser = new parserFormula( "DAYS360(DATE(2002,2,3),DATE(2005,5,31),TRUE)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1197 );

        testArrayFormula2(assert, "DAYS360", 2, 3);
    } );

    QUnit.test("Test: \"WEEKNUM\"", function (assert) {
        oParser = new parserFormula( "WEEKNUM(DATE(2006,1,1))", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "WEEKNUM(DATE(2006,1,1),17)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "WEEKNUM(DATE(2006,1,1),1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "WEEKNUM(DATE(2006,1,1),21)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 52 );

        oParser = new parserFormula( "WEEKNUM(DATE(2006,2,1),1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 5 );

        oParser = new parserFormula( "WEEKNUM(DATE(2006,2,1),2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 6 );

        oParser = new parserFormula( "WEEKNUM(DATE(2006,2,1),11)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 6 );

        oParser = new parserFormula( "WEEKNUM(DATE(2007,1,1),15)", "A2", ws );//понед
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "WEEKNUM(DATE(2008,1,1),15)", "A2", ws );//втор
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "WEEKNUM(DATE(2003,1,1),15)", "A2", ws );//сред
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "WEEKNUM(DATE(2009,1,1),15)", "A2", ws );//чет
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "WEEKNUM(DATE(2010,1,1),15)", "A2", ws );//пят
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "WEEKNUM(DATE(2011,1,1),15)", "A2", ws );//суб
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "WEEKNUM(DATE(2012,1,1),11)", "A2", ws );//вск
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "WEEKNUM(DATE(2008,1,4),11)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "WEEKNUM(DATE(2008,1,10),11)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2 );

        oParser = new parserFormula( "WEEKNUM(DATE(2008,1,11),11)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2 );

        oParser = new parserFormula( "WEEKNUM(DATE(2008,1,17),11)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3 );

        oParser = new parserFormula( "WEEKNUM(DATE(2008,1,18),11)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3 );

        oParser = new parserFormula( "WEEKNUM(DATE(2008,1,24),11)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4 );

        oParser = new parserFormula( "WEEKNUM(DATE(2013,1,1),21)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "WEEKNUM(DATE(2013,1,7))", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "WEEKNUM(0, 21)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 52 );

		oParser = new parserFormula( "WEEKNUM(1, 21)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 52 );

		oParser = new parserFormula( "WEEKNUM(2, 21)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( "WEEKNUM(0, 17)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "WEEKNUM(1, 17)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

        testArrayFormula2(assert, "WEEKNUM", 1, 2, true, null);
    } );

	QUnit.test("Test: \"ISOWEEKNUM\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "3/9/2012" );

	    oParser = new parserFormula( "ISOWEEKNUM(A2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 10 );

		oParser = new parserFormula( "ISOWEEKNUM(123)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 18 );

		oParser = new parserFormula( "ISOWEEKNUM(120003)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 30 );

		oParser = new parserFormula( "ISOWEEKNUM(120003)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 30 );

		oParser = new parserFormula( "ISOWEEKNUM(-100)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		oParser = new parserFormula( "ISOWEEKNUM(1203)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 16 );

		testArrayFormula2(assert, "ISOWEEKNUM",1,1);
	} );

	QUnit.test("Test: \"WEEKDAY\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "2/14/2008" );

		oParser = new parserFormula( "WEEKDAY(A2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 5 );

		oParser = new parserFormula( "WEEKDAY(A2, 2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );

		oParser = new parserFormula( "WEEKDAY(A2, 3)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		testArrayFormula2(assert, "WEEKDAY", 1, 2);
	} );


	QUnit.test("Test: \"WEIBULL\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "105" );
		ws.getRange2( "A3" ).setValue( "20" );
		ws.getRange2( "A4" ).setValue( "100" );

		oParser = new parserFormula( "WEIBULL(A2,A3,A4,TRUE)", "A20", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.929581 );

		oParser = new parserFormula( "WEIBULL(A2,A3,A4,FALSE)", "A20", ws );
		assert.ok( oParser.parse(), "WEIBULL(A2,A3,A4,FALSE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.035589 );

		testArrayFormula2(assert, "WEIBULL", 4, 4);
	} );

	QUnit.test("Test: \"WEIBULL.DIST\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "105" );
		ws.getRange2( "A3" ).setValue( "20" );
		ws.getRange2( "A4" ).setValue( "100" );

		oParser = new parserFormula( "WEIBULL.DIST(A2,A3,A4,TRUE)", "A20", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.929581 );

		oParser = new parserFormula( "WEIBULL.DIST(A2,A3,A4,FALSE)", "A20", ws );
		assert.ok( oParser.parse(), "WEIBULL.DIST(A2,A3,A4,FALSE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.035589 );

		testArrayFormula2(assert, "WEIBULL.DIST", 4, 4);
	} );

    QUnit.test("Test: \"YEARFRAC\"", function (assert) {
        function okWrapper( a, b ) {
            assert.ok( Math.abs( a - b ) < dif );
        }

        oParser = new parserFormula( "YEARFRAC(DATE(2006,1,1),DATE(2006,3,26))", "A2", ws );
        assert.ok( oParser.parse() );
        okWrapper( oParser.calculate().getValue(), 0.236111111 );

        oParser = new parserFormula( "YEARFRAC(DATE(2006,3,26),DATE(2006,1,1))", "A2", ws );
        assert.ok( oParser.parse() );
        okWrapper( oParser.calculate().getValue(), 0.236111111 );

        oParser = new parserFormula( "YEARFRAC(DATE(2006,1,1),DATE(2006,7,1))", "A2", ws );
        assert.ok( oParser.parse() );
        okWrapper( oParser.calculate().getValue(), 0.5 );

        oParser = new parserFormula( "YEARFRAC(DATE(2006,1,1),DATE(2007,9,1))", "A2", ws );
        assert.ok( oParser.parse() );
        okWrapper( oParser.calculate().getValue(), 1.666666667 );

        oParser = new parserFormula( "YEARFRAC(DATE(2006,1,1),DATE(2006,7,1),0)", "A2", ws );
        assert.ok( oParser.parse() );
        okWrapper( oParser.calculate().getValue(), 0.5 );

        oParser = new parserFormula( "YEARFRAC(DATE(2006,1,1),DATE(2006,7,1),1)", "A2", ws );
        assert.ok( oParser.parse() );
        okWrapper( oParser.calculate().getValue(), 0.495890411 );

        oParser = new parserFormula( "YEARFRAC(DATE(2006,1,1),DATE(2006,7,1),2)", "A2", ws );
        assert.ok( oParser.parse() );
        okWrapper( oParser.calculate().getValue(), 0.502777778 );

        oParser = new parserFormula( "YEARFRAC(DATE(2006,1,1),DATE(2006,7,1),3)", "A2", ws );
        assert.ok( oParser.parse() );
        okWrapper( oParser.calculate().getValue(), 0.495890411 );

        oParser = new parserFormula( "YEARFRAC(DATE(2006,1,1),DATE(2006,7,1),4)", "A2", ws );
        assert.ok( oParser.parse() );
        okWrapper( oParser.calculate().getValue(), 0.5 );

        oParser = new parserFormula( "YEARFRAC(DATE(2004,3,1),DATE(2006,3,1),1)", "A2", ws );
        assert.ok( oParser.parse() );
        okWrapper( oParser.calculate().getValue(), 1.998175182481752 );

        testArrayFormula2(assert, "YEARFRAC", 2, 3, true, null);
    } );

    QUnit.test("Test: \"DATEDIF\"", function (assert) {

        oParser = new parserFormula( "DATEDIF(DATE(2001,1,1),DATE(2003,1,1),\"Y\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2 );

        oParser = new parserFormula( "DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"D\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 440 );

        oParser = new parserFormula( "DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"YD\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 75 );

        oParser = new parserFormula( "DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"MD\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 14 );

        testArrayFormula2(assert, "DATEDIF", 3, 3);
    } );

	QUnit.test("Test: \"PRODUCT\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "5" );
		ws.getRange2( "A3" ).setValue( "15" );
		ws.getRange2( "A4" ).setValue( "30" );

		oParser = new parserFormula( "PRODUCT(A2:A4)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2250 );

		oParser = new parserFormula( "PRODUCT(A2:A4, 2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4500 );

		testArrayFormula2(assert, "PRODUCT", 1, 8, null, true);
	} );

    QUnit.test("Test: \"SUMPRODUCT\"", function (assert) {

        oParser = new parserFormula( "SUMPRODUCT({2,3})", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 5 );

        oParser = new parserFormula( "SUMPRODUCT({2,3},{4,5})", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 23 );

        oParser = new parserFormula( "SUMPRODUCT({2,3},{4,5},{2,2})", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 46 );

        oParser = new parserFormula( "SUMPRODUCT({2,3;4,5},{2,2;3,4})", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 42 );

        ws.getRange2( "N44" ).setValue( "1" );
        ws.getRange2( "N45" ).setValue( "2" );
        ws.getRange2( "N46" ).setValue( "3" );
        ws.getRange2( "N47" ).setValue( "4" );

        ws.getRange2( "O44" ).setValue( "5" );
        ws.getRange2( "O45" ).setValue( "6" );
        ws.getRange2( "O46" ).setValue( "7" );
        ws.getRange2( "O47" ).setValue( "8" );

        ws.getRange2( "P44" ).setValue( "9" );
        ws.getRange2( "P45" ).setValue( "10" );
        ws.getRange2( "P46" ).setValue( "11" );
        ws.getRange2( "P47" ).setValue( "12" );
		ws.getRange2( "P48" ).setValue( "" );
		ws.getRange2( "P49" ).setValue( "" );
		ws.getRange2( "N48" ).setValue( "0.456" );
		ws.getRange2( "O48" ).setValue( "0.123212" );

        oParser = new parserFormula( "SUMPRODUCT(N44:N47,O44:O47,P44:P47)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 780 );

		oParser = new parserFormula( "SUMPRODUCT(N44:N47*O44:O47)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 70 );

		oParser = new parserFormula( "SUMPRODUCT(SUM(N44:N47*O44:O47))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 70 );

		oParser = new parserFormula( "SUMPRODUCT({1,2,TRUE,3})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 6 );

		oParser = new parserFormula( "SUMPRODUCT({1,2,FALSE,3})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 6 );

		oParser = new parserFormula( "SUMPRODUCT({TRUE,TRUE,FALSE,3})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "SUMPRODUCT(P48)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "SUMPRODUCT(P48, P44:P47)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "SUMPRODUCT(P48:P49)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "SUM(SUMPRODUCT(N44:N47*O44:O47))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 70 );

		oParser = new parserFormula( "SUMPRODUCT(N44:O47*P44:P47)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 388 );

		oParser = new parserFormula( "SUM(SUMPRODUCT(N44:O47*P44:P47))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 388 );

		oParser = new parserFormula( "SUM(SUMPRODUCT(N44:O47))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUM(SUMPRODUCT(N44:O47))" );
		assert.strictEqual( oParser.calculate().getValue(), 36 );

		oParser = new parserFormula( "SUMPRODUCT(YEAR(N45:O47))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 11400 );

		oParser = new parserFormula( "SUMPRODUCT(MONTH(N45:O47))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 6 );

		oParser = new parserFormula( "SUMPRODUCT(DAY(N45:O47))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 30 );

		oParser = new parserFormula( "SUMPRODUCT(HOUR(N45:P48))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 12 );

		oParser = new parserFormula( "SUMPRODUCT(MINUTE(N45:P48))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 113 );

		oParser = new parserFormula( "SUMPRODUCT(SECOND(N45:P48))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 64 );

		oParser = new parserFormula( "SUMPRODUCT(DAY(N44:P49))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 78 );

		oParser = new parserFormula( "SUMPRODUCT(MONTH(N44:P49))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 18 );

		oParser = new parserFormula( "SUMPRODUCT(YEAR(N44:P49))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 34200 );

		oParser = new parserFormula( "SUMPRODUCT(({1,2,3})*({TRUE,TRUE,TRUE}))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 6 );

		/*oParser = new parserFormula( "SUMPRODUCT(({1,2,3})*({TRUE;TRUE;TRUE;TRUE}))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 24 );*/

		oParser = new parserFormula( "SUMPRODUCT({TRUE,TRUE,FALSE})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "SUMPRODUCT({1,2,3,3,TRUE})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 9 );

		oParser = new parserFormula( "SUMPRODUCT({1,2,3,3,TRUE})+SUMPRODUCT({1,2,3,3,TRUE})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUMPRODUCT({1,2,3,3,TRUE})+SUMPRODUCT({1,2,3,3,TRUE})" );
		assert.strictEqual( oParser.calculate().getValue(), 18 );

		oParser = new parserFormula( "SUMPRODUCT({1,2,3,3,TRUE})*SUMPRODUCT({1,2,3,3,TRUE})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUMPRODUCT({1,2,3,3,TRUE})*SUMPRODUCT({1,2,3,3,TRUE})" );
		assert.strictEqual( oParser.calculate().getValue(), 81 );

		oParser = new parserFormula( "SUMPRODUCT(SUMPRODUCT({1,2,3,3,TRUE})*SUMPRODUCT({1,2,3,3,TRUE}))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUMPRODUCT(SUMPRODUCT({1,2,3,3,TRUE})*SUMPRODUCT({1,2,3,3,TRUE}))" );
		assert.strictEqual( oParser.calculate().getValue(), 81 );

		oParser = new parserFormula( "SUM(SUMPRODUCT({1,2,3,3,TRUE})*SUMPRODUCT({1,2,3,3,TRUE}))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUM(SUMPRODUCT({1,2,3,3,TRUE})*SUMPRODUCT({1,2,3,3,TRUE}))" );
		assert.strictEqual( oParser.calculate().getValue(), 81 );

		oParser = new parserFormula( "SUM(SUMPRODUCT({1,2,3,3,TRUE})*SUMPRODUCT({1,2,3,3,TRUE}),1,2,3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUM(SUMPRODUCT({1,2,3,3,TRUE})*SUMPRODUCT({1,2,3,3,TRUE}),1,2,3)" );
		assert.strictEqual( oParser.calculate().getValue(), 87 );

		oParser = new parserFormula( "SUM(SUMPRODUCT(N44:O47))+SUM(SUMPRODUCT(N44:O47))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUM(SUMPRODUCT(N44:O47))+SUM(SUMPRODUCT(N44:O47))" );
		assert.strictEqual( oParser.calculate().getValue(), 72 );

		oParser = new parserFormula( "SUM(SUMPRODUCT({1,2,3,3,TRUE})*SUMPRODUCT({1,2,3,3,TRUE}),SUMPRODUCT({1,2,3,3,TRUE}),2,SUMPRODUCT({1,2,3,3}))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUM(SUMPRODUCT({1,2,3,3,TRUE})*SUMPRODUCT({1,2,3,3,TRUE}),SUMPRODUCT({1,2,3,3,TRUE}),2,SUMPRODUCT({1,2,3,3}))" );
		assert.strictEqual( oParser.calculate().getValue(), 101 );

		ws.getRange2( "A101" ).setValue( "5" );
		ws.getRange2( "A102" ).setValue( "6" );
		ws.getRange2( "A103" ).setValue( "7" );
		ws.getRange2( "A104" ).setValue( "8" );
		ws.getRange2( "A105" ).setValue( "9" );

		ws.getRange2( "B101" ).setValue( "1" );
		ws.getRange2( "B102" ).setValue( "1" );
		ws.getRange2( "B103" ).setValue( "0" );
		ws.getRange2( "B104" ).setValue( "1" );
		ws.getRange2( "B105" ).setValue( "1" );

		oParser = new parserFormula( "SUMPRODUCT((A101:A105)*((B101:B105)=1))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUMPRODUCT((A101:A105)*((B101:B105)=1))" );
		assert.strictEqual( oParser.calculate().getValue(), 28 );

		oParser = new parserFormula( "SUMPRODUCT((A101:A105)*((B101:B105)=1))+SUMPRODUCT((A101:A104)*((B101:B104)=1))+SUMPRODUCT((A101:A103)*((B101:B103)=1))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUMPRODUCT((A101:A105)*((B101:B105)=1))+SUMPRODUCT((A101:A104)*((B101:B104)=1))+SUMPRODUCT((A101:A103)*((B101:B103)=1))" );
		assert.strictEqual( oParser.calculate().getValue(), 58 );

		oParser = new parserFormula( "SUMPRODUCT(({3})*({TRUE,TRUE,TRUE,TRUE}))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUMPRODUCT(({3})*({TRUE,TRUE,TRUE,TRUE}))" );
		assert.strictEqual( oParser.calculate().getValue(), 12 );

		oParser = new parserFormula( "SUMPRODUCT(({3;2;2;2})*({TRUE;TRUE;TRUE;TRUE}))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUMPRODUCT(({3;2;2;2})*({TRUE;TRUE;TRUE;TRUE}))" );
		assert.strictEqual( oParser.calculate().getValue(), 9 );

		oParser = new parserFormula( "SUMPRODUCT(--ISNUMBER({5;6;7;1;2;3;4}))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUMPRODUCT(--ISNUMBER({5;6;7;1;2;3;4}))" );
		assert.strictEqual( oParser.calculate().getValue(), 7 );

		oParser = new parserFormula( "SUMPRODUCT(--ISNUMBER(SEARCH({5;6;7;1;2;3;4},123)))", "A2", ws );
		assert.ok( oParser.parse() );
		assert.ok( oParser.assemble() == "SUMPRODUCT(--ISNUMBER(SEARCH({5;6;7;1;2;3;4},123)))" );
		assert.strictEqual( oParser.calculate().getValue(), 3 );



		testArrayFormula2(assert, "SUMPRODUCT", 1, 8, null, true);
    } );

    QUnit.test("Test: \"SINH\"", function (assert) {

        oParser = new parserFormula( "SINH(0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 0 );

        oParser = new parserFormula( "SINH(1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), ((Math.E - 1 / Math.E) / 2) );

		testArrayFormula(assert, "SINH");
    } );

	QUnit.test("Test: \"SIGN\"", function (assert) {

		oParser = new parserFormula( "SIGN(10)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( "SIGN(4-4)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "SIGN(-0.00001)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -1 );

		testArrayFormula(assert, "SIGN");
	} );

    QUnit.test("Test: \"COSH\"", function (assert) {

        oParser = new parserFormula( "COSH(0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "COSH(1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), ((Math.E + 1 / Math.E) / 2) );
    } );

	QUnit.test("Test: \"IMCOSH\"", function (assert) {
		oParser = new parserFormula( 'IMCOSH("4+3i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMCOSH("4+3i")' );
		assert.strictEqual( oParser.calculate().getValue(), "-27.03494560307422+3.8511533348117766i", 'IMCOSH("4+3i")' );

		testArrayFormula(assert, "IMCOSH", true);
	} );

	QUnit.test("Test: \"IMCOS\"", function (assert) {
		oParser = new parserFormula( 'IMCOS("1+i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMCOS("1+i")' );
		assert.strictEqual( oParser.calculate().getValue(), "0.8337300251311491-0.9888977057628651i", 'IMCOS("1+i")' );

		testArrayFormula(assert, "IMCOS", true);
	} );

	QUnit.test("Test: \"IMCOT\"", function (assert) {
		oParser = new parserFormula( 'IMCOT("4+3i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMCOT("4+3i")' );
		assert.strictEqual( oParser.calculate().getValue(), "0.004901182394304475-0.9992669278059015i", 'IMCOT("4+3i")' );

		testArrayFormula(assert, "IMCOT", true);
	} );

	QUnit.test("Test: \"IMCSC\"", function (assert) {
		oParser = new parserFormula( 'IMCSC("4+3i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMCSC("4+3i")' );
		assert.strictEqual( oParser.calculate().getValue(), "-0.0754898329158637+0.06487747137063551i", 'IMCSC("4+3i")' );

		testArrayFormula(assert, "IMCSC", true);
	} );

	QUnit.test("Test: \"IMCSCH\"", function (assert) {

		oParser = new parserFormula( 'IMCSCH("4+3i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMCSCH("4+3i")' );
		//assert.strictEqual( oParser.calculate().getValue(), "-0.03627588962862601-0.0051744731840193976i", 'IMCSCH("4+3i")' );

		testArrayFormula(assert, "IMCSCH", true);
	} );

	QUnit.test("Test: \"IMSIN\"", function (assert) {
		oParser = new parserFormula( 'IMSIN("4+3i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMSIN("4+3i")' );
		assert.strictEqual( oParser.calculate().getValue(), "-7.619231720321408-6.548120040911002i", 'IMSIN("4+3i")' );

		testArrayFormula(assert, "IMSIN", true);
	} );

	QUnit.test("Test: \"IMSINH\"", function (assert) {
		oParser = new parserFormula( 'IMSINH("4+3i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMSINH("4+3i")' );
		assert.strictEqual( oParser.calculate().getValue(), "-27.01681325800393+3.8537380379193764i", 'IMSINH("4+3i")' );

		testArrayFormula(assert, "IMSINH", true);
	} );

	QUnit.test("Test: \"IMSEC\"", function (assert) {
		oParser = new parserFormula( 'IMSEC("4+3i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMSEC("4+3i")' );
		assert.strictEqual( oParser.calculate().getValue(), "-0.06529402785794705-0.07522496030277323i", 'IMSEC("4+3i")' );

		testArrayFormula(assert, "IMSEC", true);
	} );

	QUnit.test("Test: \"IMSECH\"", function (assert) {
		//TODO в хроме при расчёте разница, временно убираю
		oParser = new parserFormula( 'IMSECH("4+3i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMSECH("4+3i")' );
		//assert.strictEqual( oParser.calculate().getValue(), "-0.03625349691586888-0.00516434460775318i", 'IMSECH("4+3i")' );

		testArrayFormula(assert, "IMSECH", true);
	} );

	QUnit.test("Test: \"IMTAN\"", function (assert) {
		oParser = new parserFormula( 'IMTAN("4+3i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMTAN("4+3i")' );
		assert.strictEqual( oParser.calculate().getValue(), "0.004908258067496062+1.000709536067233i", 'IMTAN("4+3i")' );

		testArrayFormula(assert, "IMTAN", true);
	} );


	QUnit.test("Test: \"IMSQRT\"", function (assert) {
		oParser = new parserFormula( 'IMSQRT("1+i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMSQRT("1+i")' );
		//assert.strictEqual( oParser.calculate().getValue(), "1.0986841134678098+0.4550898605622274i", 'IMSQRT("1+i")' );

		testArrayFormula(assert, "IMSQRT", true);
	} );

	QUnit.test("Test: \"IMREAL\"", function (assert) {
		oParser = new parserFormula( 'IMREAL("6-9i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMREAL("6-9i")' );
		assert.strictEqual( oParser.calculate().getValue(), 6, 'IMREAL("6-9i")' );

		testArrayFormula(assert, "IMREAL", true);
	} );

	QUnit.test("Test: \"IMLOG2\"", function (assert) {
		//TODO в excel результат данной формулы - "2.32192809488736+1.33780421245098i"
		oParser = new parserFormula( 'IMLOG2("3+4i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMLOG2("3+4i")' );
		assert.strictEqual( oParser.calculate().getValue(), "2.321928094887362+1.3378042124509761i", 'IMLOG2("3+4i")' );

		testArrayFormula(assert, "IMLOG2", true);
	} );

	QUnit.test("Test: \"IMLOG10\"", function (assert) {
		//TODO в excel результат данной формулы - "0.698970004336019+0.402719196273373i"
		oParser = new parserFormula( 'IMLOG10("3+4i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMLOG10("3+4i")' );
		assert.strictEqual( oParser.calculate().getValue(), "0.6989700043360186+0.40271919627337305i", 'IMLOG10("3+4i")' );

		testArrayFormula(assert, "IMLOG10", true);
	} );

	QUnit.test("Test: \"IMLN\"", function (assert) {
		//TODO в excel результат данной формулы - "1.6094379124341+0.927295218001612i"
		oParser = new parserFormula( 'IMLN("3+4i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMLN("3+4i")' );
		assert.strictEqual( oParser.calculate().getValue(), "1.6094379124341003+0.9272952180016123i", 'IMLN("3+4i")' );

		testArrayFormula(assert, "IMLN", true);
	} );

	QUnit.test("Test: \"IMEXP\"", function (assert) {
		//TODO в excel результат данной формулы - "1.46869393991589+2.28735528717884i"
		oParser = new parserFormula( 'IMEXP("1+i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMEXP("1+i")' );
		assert.strictEqual( oParser.calculate().getValue(), "1.4686939399158851+2.2873552871788423i", 'IMEXP("1+i")' );

		testArrayFormula(assert, "IMEXP", true);
	} );

	QUnit.test("Test: \"IMCONJUGATE\"", function (assert) {
		oParser = new parserFormula( 'IMCONJUGATE("3+4i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMCONJUGATE("3+4i")' );
		assert.strictEqual( oParser.calculate().getValue(), "3-4i", 'IMCONJUGATE("3+4i")' );

		testArrayFormula(assert, "IMCONJUGATE", true);
	} );

	QUnit.test("Test: \"IMARGUMENT\"", function (assert) {
		oParser = new parserFormula( 'IMARGUMENT("3+4i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMARGUMENT("3+4i")' );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 0.92729522, 'IMARGUMENT("3+4i")' );

		testArrayFormula(assert, "IMARGUMENT", true);
	} );

	QUnit.test("Test: \"IMAGINARY\"", function (assert) {
		oParser = new parserFormula( 'IMAGINARY("3+4i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMAGINARY("3+4i")' );
		assert.strictEqual( oParser.calculate().getValue(), 4, 'IMAGINARY("3+4i")' );

		oParser = new parserFormula( 'IMAGINARY("0-j")', "A2", ws );
		assert.ok( oParser.parse(), 'IMAGINARY("0-j")' );
		assert.strictEqual( oParser.calculate().getValue(), -1, 'IMAGINARY("0-j")' );

		oParser = new parserFormula( 'IMAGINARY("4")', "A2", ws );
		assert.ok( oParser.parse(), 'IMAGINARY("4")' );
		assert.strictEqual( oParser.calculate().getValue(), 0, 'IMAGINARY("4")' );

		testArrayFormula(assert, "IMAGINARY", true);
	} );

	QUnit.test("Test: \"IMDIV\"", function (assert) {
		oParser = new parserFormula( 'IMDIV("-238+240i","10+24i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMDIV("-238+240i","10+24i")' );
		assert.strictEqual( oParser.calculate().getValue(), "5+12i", 'IMDIV("-238+240i","10+24i")' );

		testArrayFormula2(assert, "IMDIV", 2, 2, true, null);
	} );

	QUnit.test("Test: \"IMPOWER\"", function (assert) {

		testArrayFormula2(assert, "IMPOWER", 2, 2, true, null);
	} );

	QUnit.test("Test: \"IMABS\"", function (assert) {
		oParser = new parserFormula( 'IMABS("5+12i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMABS("5+12i"' );
		assert.strictEqual( oParser.calculate().getValue(), 13, 'IMABS("5+12i"' );

		testArrayFormula(assert, "IMABS", true);
	} );

	QUnit.test("Test: \"IMSUB\"", function (assert) {
		oParser = new parserFormula( 'IMSUB("13+4i","5+3i")', "A2", ws );
		assert.ok( oParser.parse(), 'IMSUB("13+4i","5+3i")' );
		assert.strictEqual( oParser.calculate().getValue(), "8+i", 'IMSUB("13+4i","5+3i")' );

		testArrayFormula2(assert, "IMSUB", 2, 2, true, null);
	} );

	QUnit.test("Test: \"TAN\"", function (assert) {

		oParser = new parserFormula( "TAN(0.785)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(5) - 0, 0.99920 );

		oParser = new parserFormula( "TAN(45*PI()/180)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(1) - 0, 1 );

		testArrayFormula(assert, "TAN");
	} );

    QUnit.test("Test: \"TANH\"", function (assert) {

        oParser = new parserFormula( "TANH(0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 0 );

        oParser = new parserFormula( "TANH(1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), ((Math.E * Math.E - 1) / (Math.E * Math.E + 1)) ), true );

		testArrayFormula(assert, "TANH");
    } );

	QUnit.test("Test: \"ATAN\"", function (assert) {
		oParser = new parserFormula( 'ATAN(1)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.785398163 );

		testArrayFormula(assert, "ATAN");
	} );

	QUnit.test("Test: \"ATAN2\"", function (assert) {
		oParser = new parserFormula( 'ATAN2(1, 1)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.785398163);

		oParser = new parserFormula( 'ATAN2(-1, -1)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, -2.35619449);

		oParser = new parserFormula( 'ATAN2(-1, -1)*180/PI()', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -135);

		oParser = new parserFormula( 'DEGREES(ATAN2(-1, -1))', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -135);

		testArrayFormula2(assert, "ATAN2", 2, 2);
	} );

	QUnit.test("Test: \"ATANH\"", function (assert) {
		oParser = new parserFormula( 'ATANH(0.76159416)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 1.00000001 );

		oParser = new parserFormula( 'ATANH(-0.1)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, -0.100335348 );

		testArrayFormula(assert, "ATANH");
	} );

	QUnit.test("Test: \"XOR\"", function (assert) {
		oParser = new parserFormula( 'XOR(3>0,2<9)', "A2", ws );
		assert.ok( oParser.parse(), 'XOR(3>0,2<9)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", 'XOR(3>0,2<9)' );

		oParser = new parserFormula( 'XOR(3>12,4>6)', "A2", ws );
		assert.ok( oParser.parse(), 'XOR(3>12,4>6)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", 'XOR(3>12,4>6)' );

		oParser = new parserFormula( 'XOR(3>12,4<6)', "A2", ws );
		assert.ok( oParser.parse(), 'XOR(3>12,4<6)' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", 'XOR(3>12,4<6)' );

		//area - specific for xor function
		//all empty - false result
		ws.getRange2( "A101" ).setValue( "5" );
		ws.getRange2( "A102" ).setValue( "6" );
		ws.getRange2( "A103" ).setValue( "test1" );
		ws.getRange2( "A104" ).setValue( "" );
		ws.getRange2( "A105" ).setValue( "false" );

		ws.getRange2( "B101" ).setValue( "1" );
		ws.getRange2( "B102" ).setValue( "1" );
		ws.getRange2( "B103" ).setValue( "test2" );
		ws.getRange2( "B104" ).setValue( "" );
		ws.getRange2( "B105" ).setValue( "false" );

		ws.getRange2( "B106" ).setValue( "#VALUE!" );

		oParser = new parserFormula( 'XOR(A101:B102)', "A2", ws );
		assert.ok( oParser.parse(), 'XOR(A101:B102)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE" );

		oParser = new parserFormula( 'XOR(A101:B103)', "A2", ws );
		assert.ok( oParser.parse(), 'XOR(A101:B103)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE" );

		oParser = new parserFormula( 'XOR(A101:A103)', "A2", ws );
		assert.ok( oParser.parse(), 'XOR(A101:A103)' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

		oParser = new parserFormula( 'XOR(A101:A104)', "A2", ws );
		assert.ok( oParser.parse(), 'XOR(A101:A104)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE" );

		oParser = new parserFormula( 'XOR(A104:B104)', "A2", ws );
		assert.ok( oParser.parse(), 'XOR(A104:B104)' );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( 'XOR(A101:B104)', "A2", ws );
		assert.ok( oParser.parse(), 'XOR(A101:B104)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE" );

		oParser = new parserFormula( 'XOR(A101:B105)', "A2", ws );
		assert.ok( oParser.parse(), 'XOR(A101:B105)' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE" );

		oParser = new parserFormula( 'XOR(A101:A105)', "A2", ws );
		assert.ok( oParser.parse(), 'XOR(A101:A105)' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE" );

		oParser = new parserFormula( 'XOR(B101:A106)', "A2", ws );
		assert.ok( oParser.parse(), 'XOR(B101:A106)' );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );


		testArrayFormula2(assert, "XOR", 1, 8, null, true);
	} );

    QUnit.test("Test: \"COMBIN\"", function (assert) {

        oParser = new parserFormula( "COMBIN(8,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 28 );

        oParser = new parserFormula( "COMBIN(10,4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 210 );

        oParser = new parserFormula( "COMBIN(6,5)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 6 );

        oParser = new parserFormula( "COMBIN(-6,5)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "COMBIN(3,5)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "COMBIN(6,-5)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );
    } );

    QUnit.test("Test: \"FACTDOUBLE\"", function (assert) {

        oParser = new parserFormula( "FACTDOUBLE(8)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2 * 4 * 6 * 8 );

        oParser = new parserFormula( "FACTDOUBLE(9)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 9 * 7 * 5 * 3 );

        oParser = new parserFormula( "FACTDOUBLE(6.5)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 6 * 4 * 2 );

        oParser = new parserFormula( "FACTDOUBLE(-6)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "FACTDOUBLE(600)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		testArrayFormula(assert, "FACTDOUBLE", true);
    } );

	QUnit.test("Test: \"FACT\"", function (assert) {
		oParser = new parserFormula( "FACT(5)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 120 );

		oParser = new parserFormula( "FACT(1.9)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( "FACT(0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( "FACT(-1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		oParser = new parserFormula( "FACT(1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		testArrayFormula(assert, "FACT");
	} );

    QUnit.test("Test: \"GCD\"", function (assert) {
        oParser = new parserFormula( "LCM(5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 5 );

        oParser = new parserFormula( "LCM(24.6,36.2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 72 );

        oParser = new parserFormula( "LCM(-1,39,52)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "LCM(0,39,52)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "LCM(24,36,15)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 360 );

		testArrayFormula2(assert, "LCM", 1, 8, null, true);
    } );


    QUnit.test("Test: \"RANDBETWEEN\"", function (assert) {
        var res;
        oParser = new parserFormula( "RANDBETWEEN(1,6)", "A1", ws );
        assert.ok( oParser.parse() );
        res = oParser.calculate().getValue();
        assert.ok( res >= 1 && res <= 6 );

        oParser = new parserFormula( "RANDBETWEEN(-10,10)", "A1", ws );
        assert.ok( oParser.parse() );
        res = oParser.calculate().getValue();
        assert.ok( res >= -10 && res <= 10 );

        oParser = new parserFormula( "RANDBETWEEN(-25,-3)", "A1", ws );
        assert.ok( oParser.parse() );
        res = oParser.calculate().getValue();
        assert.ok( res >= -25 && res <= -3 );

        testArrayFormula2(assert, "RANDBETWEEN", 2, 2, true)
    } );

	QUnit.test("Test: \"RANDARRAY\"", function (assert) {
		ws.getRange2( "A101" ).setValue( "2" );
		ws.getRange2( "A102" ).setValue( "2" );
		ws.getRange2( "A103" ).setValue( "test1" );
		ws.getRange2( "A104" ).setValue( "-4" );
		ws.getRange2( "A105" ).setValue( "-1" );

		var res;
		oParser = new parserFormula( "RANDARRAY(A101,A102)", "A1", ws );
		assert.ok( oParser.parse() );
		res = oParser.calculate().getElementRowCol(0,0).getValue();
		assert.ok( res >= 0 && res <= 1 );
		res = oParser.calculate().getElementRowCol(1,0).getValue();
		assert.ok( res >= 0 && res <= 1 );
		res = oParser.calculate().getElementRowCol(0,1).getValue();
		assert.ok( res >= 0 && res <= 1 );
		res = oParser.calculate().getElementRowCol(1,1).getValue();
		assert.ok( res >= 0 && res <= 1 );

		oParser = new parserFormula( "RANDARRAY(A101,A102,A104,A105)", "A1", ws );
		assert.ok( oParser.parse() );
		res = oParser.calculate().getElementRowCol(0,0).getValue();
		assert.ok( res >= -4 && res <= -1 );
		res = oParser.calculate().getElementRowCol(1,0).getValue();
		assert.ok( res >= -4 && res <= -1 );
		res = oParser.calculate().getElementRowCol(0,1).getValue();
		assert.ok( res >= -4 && res <= -1 );
		res = oParser.calculate().getElementRowCol(1,1).getValue();
		assert.ok( res >= -4 && res <= -1 );

		oParser = new parserFormula( "RANDARRAY()", "A1", ws );
		assert.ok( oParser.parse() );
		res = oParser.calculate().getElementRowCol(0,0).getValue();
		assert.ok( res >= 0 && res <= 1 );

		oParser = new parserFormula( "RANDARRAY(2,2)", "A1", ws );
		assert.ok( oParser.parse() );
		res = oParser.calculate().getElementRowCol(0,0).getValue();
		assert.ok( res >= 0 && res <= 1 );
		res = oParser.calculate().getElementRowCol(1,0).getValue();
		assert.ok( res >= 0 && res <= 1 );
		res = oParser.calculate().getElementRowCol(0,1).getValue();
		assert.ok( res >= 0 && res <= 1 );
		res = oParser.calculate().getElementRowCol(1,1).getValue();
		assert.ok( res >= 0 && res <= 1 );

		oParser = new parserFormula( "RANDARRAY(2,2,-10,-5)", "A1", ws );
		assert.ok( oParser.parse() );
		res = oParser.calculate().getElementRowCol(0,0).getValue();
		assert.ok( res >= -10 && res <= -5 );
		res = oParser.calculate().getElementRowCol(1,0).getValue();
		assert.ok( res >= -10 && res <= -5 );
		res = oParser.calculate().getElementRowCol(0,1).getValue();
		assert.ok( res >= -10 && res <= -5 );
		res = oParser.calculate().getElementRowCol(1,1).getValue();
		assert.ok( res >= -10 && res <= -5 );

		oParser = new parserFormula( "RANDARRAY(2,2,-10,-5, true)", "A1", ws );
		assert.ok( oParser.parse() );
		res = oParser.calculate().getElementRowCol(0,0).getValue();
		assert.ok( res >= -10 && res <= -5 );
		res = oParser.calculate().getElementRowCol(1,0).getValue();
		assert.ok( res >= -10 && res <= -5 );
		res = oParser.calculate().getElementRowCol(0,1).getValue();
		assert.ok( res >= -10 && res <= -5 );
		res = oParser.calculate().getElementRowCol(1,1).getValue();
		assert.ok( res >= -10 && res <= -5 );

		oParser = new parserFormula( "RANDARRAY(2,2,,, true)", "A1", ws );
		assert.ok( oParser.parse() );
		res = oParser.calculate().getElementRowCol(0,0).getValue();
		assert.ok( res >= 0 && res <= 1 );
		res = oParser.calculate().getElementRowCol(1,0).getValue();
		assert.ok( res >= 0 && res <= 1 );
		res = oParser.calculate().getElementRowCol(0,1).getValue();
		assert.ok( res >= 0 && res <= 1 );
		res = oParser.calculate().getElementRowCol(1,1).getValue();
		assert.ok( res >= 0 && res <= 1 );

		oParser = new parserFormula( "RANDARRAY(-1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "RANDARRAY(1, -1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "RANDARRAY(-1, -1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "RANDARRAY(1, 1, 5, 2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "RANDARRAY(0, 0, 5, 2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "RANDARRAY(0, 1, 5, 2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "RANDARRAY(1, 0, 5, 2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "RANDARRAY(1,1,,-2,FALSE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "RANDARRAY(1,1,2,,FALSE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "RANDARRAY(A103,1,2,,FALSE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		ws.getRange2( "A101" ).setValue( "#DIV/0!" );
		ws.getRange2( "A102" ).setValue( "2" );
		ws.getRange2( "A103" ).setValue( "3" );
		ws.getRange2( "A104" ).setValue( "4" );
		ws.getRange2( "A105" ).setValue( "5" );
		ws.getRange2( "A105" ).setValue( "6" );

		ws.getRange2( "B101" ).setValue( "-3" );
		ws.getRange2( "B102" ).setValue( "-5" );
		ws.getRange2( "B103" ).setValue( "2" );
		ws.getRange2( "B104" ).setValue( "1" );
		ws.getRange2( "B105" ).setValue( "12" );
		ws.getRange2( "B105" ).setValue( "13" );

		oParser = new parserFormula( "RANDARRAY({1,123,3},{1,2,3,4},,2,A101:C101)", "A1", ws );
		assert.ok( oParser.parse() );
		res = oParser.calculate().getElementRowCol(0,0).getValue();
		assert.strictEqual( res, "#DIV/0!" );
		res = oParser.calculate().getElementRowCol(0,1).getValue();
		assert.ok( res >= 0 && res <= 2 );
		res = oParser.calculate().getElementRowCol(0,2).getValue();
		assert.ok( res >= 0 && res <= 2 );

	} );

    QUnit.test("Test: \"QUOTIENT\"", function (assert) {
        oParser = new parserFormula( "QUOTIENT(1,6)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 0 );

        oParser = new parserFormula( "QUOTIENT(-10,3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -3 );

        oParser = new parserFormula( "QUOTIENT(5,3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "QUOTIENT(5,0)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );

        testArrayFormula2(assert, "QUOTIENT", 2 , 2, true)
    } );

    QUnit.test("Test: \"TRUNC\"", function (assert) {
        oParser = new parserFormula( "TRUNC(PI())", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3 );

        oParser = new parserFormula( "TRUNC(PI(),3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3.141 );

        oParser = new parserFormula( "TRUNC(PI(),-2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 0 );

        oParser = new parserFormula( "TRUNC(-PI(),2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -3.14 );

		oParser = new parserFormula( "TRUNC(8.9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 8 );

		oParser = new parserFormula( "TRUNC(-8.9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -8 );

		oParser = new parserFormula( "TRUNC(0.45)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );


		oParser = new parserFormula( "TRUNC(43214)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 43214 );

		oParser = new parserFormula( "TRUNC(43214, 10)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 43214 );

		oParser = new parserFormula( "TRUNC(43214, -2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 43200 );

		oParser = new parserFormula( "TRUNC(43214, -10)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "TRUNC(34123.123, -2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 34100 );

		oParser = new parserFormula( "TRUNC(123.23423,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 123.2 );

		testArrayFormula2(assert, "TRUNC", 1, 2);
	} );

    QUnit.test("Test: \"MULTINOMIAL\"", function (assert) {
        oParser = new parserFormula( "MULTINOMIAL(2,3,4)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), Math.fact( 2 + 3 + 4 ) / (Math.fact( 2 ) * Math.fact( 3 ) * Math.fact( 4 )) );

        oParser = new parserFormula( "MULTINOMIAL(2,3,\"r\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        oParser = new parserFormula( "MULTINOMIAL(150,50)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        testArrayFormula2(assert, "MULTINOMIAL", 1, 8, null, true);
    } );

	QUnit.test("Test: \"MUNIT\"", function (assert) {
		ws.getRange2( "A101" ).setValue( "5" );
		ws.getRange2( "B102" ).setValue( "6" );

		oParser = new parserFormula( "MUNIT(1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 1 );

		oParser = new parserFormula( "MUNIT(-1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "MUNIT(1.123)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 1 );

		oParser = new parserFormula( "MUNIT(2.123)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,1).getValue(), 1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,0).getValue(), 0 );

		oParser = new parserFormula( "MUNIT(A101)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,1).getValue(), 1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,0).getValue(), 0 );

		oParser = new parserFormula( "MUNIT(A101:B102)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "MUNIT({0,0;1,2;123,\"sdf\"})", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), "#VALUE!" );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,1).getValue(), 1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(2,0).getValue(), 1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(2,1).getValue(), "#VALUE!" );

		oParser = new parserFormula( "MUNIT({12,2})", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue(), 1 );
	} );

    QUnit.test("Test: \"SUMSQ\"", function (assert) {
        oParser = new parserFormula( "SUMSQ(2.5,-3.6,2.4)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2.5 * 2.5 + 3.6 * 3.6 + 2.4 * 2.4 );

        oParser = new parserFormula( "SUMSQ(2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4 );

        oParser = new parserFormula( "SUMSQ(150,50)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 150 * 150 + 50 * 50 );

        oParser = new parserFormula( "SUMSQ(150,\"f\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        testArrayFormula2(assert, "SUMSQ", 1, 8, null, true);
    } );

    QUnit.test("Test: \"ROMAN\"", function (assert) {

        oParser = new parserFormula( "ROMAN(499,0)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "CDXCIX" );

        oParser = new parserFormula( "ROMAN(499,1)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "LDVLIV" );

        oParser = new parserFormula( "ROMAN(499,2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "XDIX" );

        oParser = new parserFormula( "ROMAN(499,3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "VDIV" );

        oParser = new parserFormula( "ROMAN(499,4)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "ID" );

        oParser = new parserFormula( "ROMAN(2013,0)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "MMXIII" );

        oParser = new parserFormula( "ROMAN(2013,5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        oParser = new parserFormula( "ROMAN(-2013,1)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        oParser = new parserFormula( "ROMAN(2499,1)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "MMLDVLIV" );

		oParser = new parserFormula( "ROMAN(499)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "CDXCIX" );

		testArrayFormula2(assert, "ROMAN", 2, 2);
    } );

    QUnit.test("Test: \"SUMXMY2\"", function (assert) {

		ws.getRange2( "A101" ).setValue( "5" );
		ws.getRange2( "A102" ).setValue( "6" );
		ws.getRange2( "A103" ).setValue( "test1" );
		ws.getRange2( "A104" ).setValue( "" );
		ws.getRange2( "A105" ).setValue( "false" );

		ws.getRange2( "B101" ).setValue( "1" );
		ws.getRange2( "B102" ).setValue( "1" );
		ws.getRange2( "B103" ).setValue( "test2" );
		ws.getRange2( "B104" ).setValue( "" );
		ws.getRange2( "B105" ).setValue( "false" );

		oParser = new parserFormula( "SUMXMY2(A101,B101)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 16 );

		oParser = new parserFormula( "SUMXMY2(A103,B103)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "SUMXMY2(A101:A102,B101:B102)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 41 );

		oParser = new parserFormula( "SUMXMY2(A105,B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

    	oParser = new parserFormula( "SUMXMY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 79 );

        oParser = new parserFormula( "SUMXMY2({2,3,9;1,8,7},{6,5,11;7,5,4})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 78 );

        oParser = new parserFormula( "SUMXMY2(7,5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4 );

		testArrayFormula2(assert, "SUMXMY2", 2, 2, null, true)
    } );

    QUnit.test("Test: \"SUMX2MY2\"", function (assert) {

		ws.getRange2( "A101" ).setValue( "5" );
		ws.getRange2( "A102" ).setValue( "6" );
		ws.getRange2( "A103" ).setValue( "test1" );
		ws.getRange2( "A104" ).setValue( "" );
		ws.getRange2( "A105" ).setValue( "false" );

		ws.getRange2( "B101" ).setValue( "1" );
		ws.getRange2( "B102" ).setValue( "1" );
		ws.getRange2( "B103" ).setValue( "test2" );
		ws.getRange2( "B104" ).setValue( "" );
		ws.getRange2( "B105" ).setValue( "false" );

        oParser = new parserFormula( "SUMX2MY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -55 );

        oParser = new parserFormula( "SUMX2MY2({2,3,9;1,8,7},{6,5,11;7,5,4})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -64 );

        oParser = new parserFormula( "SUMX2MY2(7,5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 24 );

		oParser = new parserFormula( "SUMX2MY2(A101,B101)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 24 );

		oParser = new parserFormula( "SUMX2MY2(A103,B103)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "SUMX2MY2(A101:A102,B101:B102)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 59 );

		/*oParser = new parserFormula( "SUMX2MY2(A101:A105,B101:B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 59 );*/

		oParser = new parserFormula( "SUMX2MY2(A105,B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        testArrayFormula2(assert, "SUMX2MY2", 2, 2, null, true);
    } );

    QUnit.test("Test: \"SUMX2PY2\"", function (assert) {

		ws.getRange2( "A101" ).setValue( "5" );
		ws.getRange2( "A102" ).setValue( "6" );
		ws.getRange2( "A103" ).setValue( "test1" );
		ws.getRange2( "A104" ).setValue( "" );
		ws.getRange2( "A105" ).setValue( "false" );

		ws.getRange2( "B101" ).setValue( "1" );
		ws.getRange2( "B102" ).setValue( "1" );
		ws.getRange2( "B103" ).setValue( "test2" );
		ws.getRange2( "B104" ).setValue( "" );
		ws.getRange2( "B105" ).setValue( "false" );

        oParser = new parserFormula( "SUMX2PY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 521 );

        oParser = new parserFormula( "SUMX2PY2({2,3,9;1,8,7},{6,5,11;7,5,4})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 480 );

        oParser = new parserFormula( "SUMX2PY2(7,5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 74 );


		oParser = new parserFormula( "SUMX2PY2(A101,B101)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 26 );

		oParser = new parserFormula( "SUMX2PY2(A103,B103)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "SUMX2PY2(A101:A102,B101:B102)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 63 );

		oParser = new parserFormula( "SUMX2PY2(A105,B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		testArrayFormula2(assert, "SUMX2PY2", 2, 2, null, true);
    } );

    QUnit.test("Test: \"SERIESSUM\"", function (assert) {


		ws.getRange2( "A2" ).setValue( "1" );
		ws.getRange2( "A3" ).setValue( numDivFact(-1, 2) );
		ws.getRange2( "A4" ).setValue( numDivFact(1, 4) );
		ws.getRange2( "A5" ).setValue( numDivFact(-1, 6) );

        oParser = new parserFormula( "SERIESSUM(PI()/4,0,2,A2:A5)", "A7", ws );
        assert.ok( oParser.parse() );
        assert.ok( Math.abs( oParser.calculate().getValue() - (1 - 1 / 2 * Math.pow( Math.PI / 4, 2 ) + 1 / Math.fact( 4 ) * Math.pow( Math.PI / 4, 4 ) - 1 / Math.fact( 6 ) * Math.pow( Math.PI / 4, 6 )) ) < dif );

		ws.getRange2( "B2" ).setValue( "1" );
		ws.getRange2( "B3" ).setValue( numDivFact(-1, 3) );
		ws.getRange2( "B4" ).setValue( numDivFact(1, 5) );
		ws.getRange2( "B5" ).setValue( numDivFact(-1, 7) );

        oParser = new parserFormula( "SERIESSUM(PI()/4,1,2,B2:B5)", "B7", ws );
        assert.ok( oParser.parse() );
        assert.ok( Math.abs( oParser.calculate().getValue() - (Math.PI / 4 - 1 / Math.fact( 3 ) * Math.pow( Math.PI / 4, 3 ) + 1 / Math.fact( 5 ) * Math.pow( Math.PI / 4, 5 ) - 1 / Math.fact( 7 ) * Math.pow( Math.PI / 4, 7 )) ) < dif );

		//TODO нужна другая функция для тестирования
        //testArrayFormula2(assert, "SERIESSUM", 4, 4);
    } );

    /*
    * Mathematical Function
    * */
    QUnit.test("Test: \"CEILING\"", function (assert) {

        oParser = new parserFormula( "CEILING(2.5,1)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3 );

        oParser = new parserFormula( "CEILING(-2.5,-2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -4 );

        oParser = new parserFormula( "CEILING(-2.5,2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -2 );

        oParser = new parserFormula( "CEILING(1.5,0.1)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1.5 );

        oParser = new parserFormula( "CEILING(0.234,0.01)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 0.24 );


		testArrayFormula2(assert, "CEILING", 2, 2);
    } );

	QUnit.test("Test: \"CELL\"", function (assert) {

		ws.getRange2( "J2" ).setValue( "1" );
		ws.getRange2( "J3" ).setValue( "test" );
		ws.getRange2( "J4" ).setValue( "test2" );


		oParser = new parserFormula( 'CELL("address",J3)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "$J$3" );

		oParser = new parserFormula( 'CELL("address",J3:O12)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "$J$3" );

		oParser = new parserFormula( 'CELL("col",J3)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 10 );

		oParser = new parserFormula( 'CELL("col",J3:O12)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 10 );

		oParser = new parserFormula( 'CELL("row",J3)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( 'CELL("row",J3:O12)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( 'CELL("color",J3)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( 'CELL("color",J3:O12)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( 'CELL("contents",J3)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "test" );

		oParser = new parserFormula( 'CELL("contents",J3:O12)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "test" );

		oParser = new parserFormula( 'CELL("contents",J4:O12)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "test2" );

		oParser = new parserFormula( 'CELL("contents",J5:O12)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( 'CELL("prefix",J3)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "'" );

		/*oParser = new parserFormula( 'CELL("prefix",J2)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "" );*/

		oParser = new parserFormula( 'CELL("prefix",J6:O12)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "" );

	} );


    /*
     * Statistical Function
     * */
    QUnit.test("Test: \"AVEDEV\"", function (assert) {

        oParser = new parserFormula( "AVEDEV(-3.5,1.4,6.9,-4.5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4.075 );

        oParser = new parserFormula( "AVEDEV({-3.5,1.4,6.9,-4.5})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4.075 );

        oParser = new parserFormula( "AVEDEV(-3.5,1.4,6.9,-4.5,-0.3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), 3.32 ), true );

		testArrayFormula2(assert, "AVEDEV", 1, 8, null, true);
    } );

    QUnit.test("Test: \"AVERAGE\"", function (assert) {

        oParser = new parserFormula( "AVERAGE(1,2,3,4,5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3 );

        oParser = new parserFormula( "AVERAGE({1,2;3,4})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2.5 );

        oParser = new parserFormula( "AVERAGE({1,2,3,4,5},6,\"7\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4 );

        oParser = new parserFormula( "AVERAGE({1,\"2\",TRUE,4})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2.5 );

        testArrayFormula2(assert, "AVERAGE", 1, 8, null, true);
    } );

    QUnit.test("Test: \"AVERAGEA\"", function (assert) {

        ws.getRange2( "E2" ).setValue( "TRUE" );
        ws.getRange2( "E3" ).setValue( "FALSE" );

		ws.getRange2( "F2" ).setValue( "10" );
		ws.getRange2( "F3" ).setValue( "7" );
		ws.getRange2( "F4" ).setValue( "9" );
		ws.getRange2( "F5" ).setValue( "2" );
		ws.getRange2( "F6" ).setValue( "Not available" );
		ws.getRange2( "F7" ).setValue( "" );

        oParser = new parserFormula( "AVERAGEA(10,E1)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 10 );

        oParser = new parserFormula( "AVERAGEA(10,E2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 5.5 );

        oParser = new parserFormula( "AVERAGEA(10,E3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 5 );

		oParser = new parserFormula( "AVERAGEA(F2:F6)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 5.6 );

		oParser = new parserFormula( "AVERAGEA(F2:F5,F7)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 7 );

		testArrayFormula2(assert, "AVERAGEA", 1, 8, null, true);
    } );

    QUnit.test("Test: \"AVERAGEIF\"", function (assert) {

        ws.getRange2( "E2" ).setValue( "10" );
        ws.getRange2( "E3" ).setValue( "20" );
        ws.getRange2( "E4" ).setValue( "28" );
        ws.getRange2( "E5" ).setValue( "30" );

        oParser = new parserFormula( "AVERAGEIF(E2:E5,\">15\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 26 );

        testArrayFormula2(assert, "AVERAGEIF", 2, 3, null, true);
    } );

	QUnit.test("Test: \"AVERAGEIFS\"", function (assert) {

		ws.getRange2( "E2" ).setValue( "Quiz" );
		ws.getRange2( "E3" ).setValue( "Grade" );
		ws.getRange2( "E4" ).setValue( "75" );
		ws.getRange2( "E5" ).setValue( "94" );

		ws.getRange2( "F2" ).setValue( "Quiz" );
		ws.getRange2( "F3" ).setValue( "Grade" );
		ws.getRange2( "F4" ).setValue( "85" );
		ws.getRange2( "F5" ).setValue( "80" );

		ws.getRange2( "G2" ).setValue( "Exam" );
		ws.getRange2( "G3" ).setValue( "Grade" );
		ws.getRange2( "G4" ).setValue( "87" );
		ws.getRange2( "G5" ).setValue( "88" );

		oParser = new parserFormula( "AVERAGEIFS(E2:E5,E2:E5,\">70\",E2:E5,\"<90\")", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 75 );

		oParser = new parserFormula( "AVERAGEIFS(F2:F5,F2:F5,\">95\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );

		oParser = new parserFormula( "AVERAGEIFS(G2:G5,G2:G5,\"<>Incomplete\",G2:G5,\">80\")", "A3", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 87.5 );

		testArrayFormulaEqualsValues(assert, "1,3.123,-4,#N/A;2,4,5,#N/A;#N/A,#N/A,#N/A,#N/A", "AVERAGEIFS(A1:C2,A1:C2,A1:C2,A1:C2, A1:C2,A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert, "1,#DIV/0!,#DIV/0!,#N/A;#DIV/0!,#DIV/0!,#DIV/0!,#N/A;#N/A,#N/A,#N/A,#N/A", "AVERAGEIFS(A1:C2,A1:C2,A1:A1,A1:C2,A1:C2,A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert, "1,#DIV/0!,#DIV/0!,#N/A;2,#DIV/0!,#DIV/0!,#N/A;#N/A,#N/A,#N/A,#N/A", "AVERAGEIFS(A1:C2,A1:C2,A1:A2,A1:C2,A1:C2,A1:C2,A1:C2)");
	} );

	QUnit.test("Test: \"AGGREGATE\"", function (assert) {

		ws.getRange2( "A101" ).setValue( "TEST" );
		ws.getRange2( "A102" ).setValue( "72" );
		ws.getRange2( "A103" ).setValue( "30" );
		ws.getRange2( "A104" ).setValue( "TEST2" );
		ws.getRange2( "A105" ).setValue( "31" );
		ws.getRange2( "A106" ).setValue( "96" );
		ws.getRange2( "A107" ).setValue( "32" );
		ws.getRange2( "A108" ).setValue( "81" );
		ws.getRange2( "A109" ).setValue( "33" );
		ws.getRange2( "A110" ).setValue( "53" );
		ws.getRange2( "A111" ).setValue( "34" );

		ws.getRange2( "B101" ).setValue( "82" );
		ws.getRange2( "B102" ).setValue( "65" );
		ws.getRange2( "B103" ).setValue( "95" );
		ws.getRange2( "B104" ).setValue( "63" );
		ws.getRange2( "B105" ).setValue( "53" );
		ws.getRange2( "B106" ).setValue( "71" );
		ws.getRange2( "B107" ).setValue( "55" );
		ws.getRange2( "B108" ).setValue( "83" );
		ws.getRange2( "B109" ).setValue( "100" );
		ws.getRange2( "B110" ).setValue( "91" );
		ws.getRange2( "B111" ).setValue( "89" );


		oParser = new parserFormula( "AGGREGATE(4, 6, A101:A111)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 96 );

		oParser = new parserFormula( "AGGREGATE(14, 6, A101:A111, 3)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 72 );

		oParser = new parserFormula( "AGGREGATE(15, 6, A101:A111)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( "AGGREGATE(12, 6, A101:A111, B101:B111)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 68 );

		oParser = new parserFormula( "AGGREGATE(12, 6, A101:A111, B101:B111)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 68 );

		oParser = new parserFormula( "AGGREGATE(1,1,A101:B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 61.375);

		oParser = new parserFormula( "AGGREGATE(2,1,A101:B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 8);

		oParser = new parserFormula( "AGGREGATE(3,1,A101:B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 10);

		oParser = new parserFormula( "AGGREGATE(4,1,A101:B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 95);

		oParser = new parserFormula( "AGGREGATE(5,3,A101:B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 30);

		oParser = new parserFormula( "AGGREGATE(6,1,100)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula( "AGGREGATE(7,3,A101:B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 22.87192602);

		oParser = new parserFormula( "AGGREGATE(8,3,A101:B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 21.39472774);

		oParser = new parserFormula( "AGGREGATE(9,3,A101:B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 491);

		oParser = new parserFormula( "AGGREGATE(10,3,A101:B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 523.125);

		oParser = new parserFormula( "AGGREGATE(11,3,A101:B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 457.734375);

		oParser = new parserFormula( "AGGREGATE(12,3,A101:B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 64);

		oParser = new parserFormula( "AGGREGATE(13,3,A101:B105,A101:B105)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 30);

		oParser = new parserFormula( "AGGREGATE(14,3,A101:B105,2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 82);

		oParser = new parserFormula( "AGGREGATE(15,3,A101:B105,2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 31);

		oParser = new parserFormula( "AGGREGATE(16,3,A101:B105,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 95);

		oParser = new parserFormula( "AGGREGATE(17,3,A101:B105,3)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 74.5);

		oParser = new parserFormula( "AGGREGATE(18,3,A101:B105,0.2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 30.8);

		oParser = new parserFormula( "AGGREGATE(19,3,A101:B105,2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 64);

	} );

	QUnit.test("Test: \"AND\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "50" );
		ws.getRange2( "A3" ).setValue( "100" );

		oParser = new parserFormula( "AND(A2>1,A2<100)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula( 'AND(A2<A3,A2<100)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula( 'AND(A3>1,A3<100)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE");

		testArrayFormula2(assert, "AND", 1, 8, null, true);
	} );

	QUnit.test("Test: \"OR\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "50" );
		ws.getRange2( "A3" ).setValue( "100" );

		oParser = new parserFormula( "AND(A2>1,A2<100)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula( 'AND(A2<A3,A2<100)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula( 'AND(A3<1,A3>100)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE");

		testArrayFormula2(assert, "OR", 1, 8, null, true);
	} );

    QUnit.test("Test: \"BINOMDIST\"", function (assert) {

        function binomdist( x, n, p ) {
            x = parseInt( x );
            n = parseInt( n );
            return Math.binomCoeff( n, x ) * Math.pow( p, x ) * Math.pow( 1 - p, n - x );
        }

        oParser = new parserFormula( "BINOMDIST(6,10,0.5,FALSE)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.ok( Math.abs( oParser.calculate().getValue() - binomdist( 6, 10, 0.5 ) ) < dif );

        oParser = new parserFormula( "BINOMDIST(6,10,0.5,TRUE)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.ok( Math.abs( oParser.calculate().getValue() - (function () {
            var bm = 0;
            for ( var y = 0; y <= 6; y++ ) {
                bm += binomdist( y, 10, 0.5 )
            }
            return bm;
        })() ) < dif );

        oParser = new parserFormula( "BINOMDIST(11,10,0.5,FALSE)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        testArrayFormula2(assert, "BINOMDIST", 4, 4);
    } );

	QUnit.test("Test: \"BINOM.DIST\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "6" );
		ws.getRange2( "A3" ).setValue( "10" );
		ws.getRange2( "A4" ).setValue( "0.5" );

	    oParser = new parserFormula( "BINOM.DIST(A2,A3,A4,FALSE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 7 ) - 0, 0.2050781);
	} );

	QUnit.test("Test: \"BINOM.DIST.RANGE\"", function (assert) {

		oParser = new parserFormula( "BINOM.DIST.RANGE(60,0.75,48)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 3 ) - 0, 0.084);

		oParser = new parserFormula( "BINOM.DIST.RANGE(60,0.75,45,50)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 3 ) - 0, 0.524);

		testArrayFormula2(assert, "BINOM.DIST.RANGE", 3, 4);
	} );

    QUnit.test("Test: \"CONFIDENCE\"", function (assert) {

        oParser = new parserFormula( "CONFIDENCE(0.4,5,12)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), 1.214775614397568 ), true );

        oParser = new parserFormula( "CONFIDENCE(0.75,9,7)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), 1.083909233527114 ), true );

        testArrayFormula2(assert, "CONFIDENCE", 3, 3);
    } );

	QUnit.test("Test: \"CONFIDENCE.NORM\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "0.05" );
		ws.getRange2( "A3" ).setValue( "2.5" );
		ws.getRange2( "A4" ).setValue( "50" );

		oParser = new parserFormula( "CONFIDENCE.NORM(A2,A3,A4)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 6 ) - 0, 0.692952);
	} );

	QUnit.test("Test: \"CONFIDENCE.T\"", function (assert) {

		oParser = new parserFormula( "CONFIDENCE.T(0.05,1,50)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 9 ) - 0, 0.284196855);

		testArrayFormula2(assert, "CONFIDENCE.T", 3, 3);
	} );

    QUnit.test("Test: \"CORREL\"", function (assert) {

        oParser = new parserFormula( "CORREL({2.532,5.621;2.1,3.4},{5.32,2.765;5.2,\"f\"})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), -0.988112020032211 ), true );

        oParser = new parserFormula( "CORREL({1;2;3},{4;5;\"E\"})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), 1 ), true );

        oParser = new parserFormula( "CORREL({1,2},{1,\"e\"})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );

        testArrayFormula2(assert, "CORREL", 2, 2, null, true)
    } );

    QUnit.test("Test: \"COUNT\"", function (assert) {

        ws.getRange2( "E2" ).setValue( "TRUE" );

        oParser = new parserFormula( "COUNT({1,2,3,4,5})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 5 );

        oParser = new parserFormula( "COUNT(1,2,3,4,5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 5 );

        oParser = new parserFormula( "COUNT({1,2,3,4,5},6,\"7\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 7 );

        oParser = new parserFormula( "COUNT(10,E150)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "COUNT(10,E2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

		ws.getRange2( "S5" ).setValue( "#DIV/0!" );
		ws.getRange2( "S6" ).setValue( "TRUE" );
		ws.getRange2( "S7" ).setValue( "qwe" );
		ws.getRange2( "S8" ).setValue( "" );
		ws.getRange2( "S9" ).setValue( "2" );
		oParser = new parserFormula( "COUNT(S5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "COUNT(S6)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "COUNT(S7)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "COUNT(S8)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "COUNT(S5:S9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );
		oParser = new parserFormula( "COUNT(S6:S9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		testArrayFormula2(assert, "COUNT", 2, 2, null, true);
    } );

    QUnit.test("Test: \"COUNTA\"", function (assert) {

        ws.getRange2( "E2" ).setValue( "TRUE" );

        oParser = new parserFormula( "COUNTA({1,2,3,4,5})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 5 );

        oParser = new parserFormula( "COUNTA(1,2,3,4,5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 5 );

        oParser = new parserFormula( "COUNTA({1,2,3,4,5},6,\"7\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 7 );

        oParser = new parserFormula( "COUNTA(10,E150)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "COUNTA(10,E2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2 );

		ws.getRange2( "S5" ).setValue( "#DIV/0!" );
		ws.getRange2( "S6" ).setValue( "TRUE" );
		ws.getRange2( "S7" ).setValue( "qwe" );
		ws.getRange2( "S8" ).setValue( "" );
		ws.getRange2( "S9" ).setValue( "2" );
		oParser = new parserFormula( "COUNTA(S5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );
		oParser = new parserFormula( "COUNTA(S6)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );
		oParser = new parserFormula( "COUNTA(S7)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );
		oParser = new parserFormula( "COUNTA(S8)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
		oParser = new parserFormula( "COUNTA(S5:S9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );
		oParser = new parserFormula( "COUNTA(S6:S9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );
    } );

    QUnit.test("Test: \"COUNTIFS\"", function (assert) {

        ws.getRange2( "A15" ).setValue( "Yes" );
        ws.getRange2( "A16" ).setValue( "Yes" );
        ws.getRange2( "A17" ).setValue( "Yes" );
        ws.getRange2( "A18" ).setValue( "No" );

        ws.getRange2( "B15" ).setValue( "No" );
        ws.getRange2( "B16" ).setValue( "Yes" );
        ws.getRange2( "B17" ).setValue( "Yes" );
        ws.getRange2( "B18" ).setValue( "Yes" );

		ws.getRange2( "C15" ).setValue( "No" );
		ws.getRange2( "C16" ).setValue( "No" );
		ws.getRange2( "C17" ).setValue( "Yes" );
		ws.getRange2( "C18" ).setValue( "Yes" );

        oParser = new parserFormula( "COUNTIFS(A15:C15,\"=Yes\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "COUNTIFS(A15:A18,\"=Yes\",B15:B18,\"=Yes\")", "B1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "COUNTIFS(A18:C18,\"=Yes\",A16:C16,\"=Yes\")", "C1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

		ws.getRange2( "D15" ).setValue( "1" );
		ws.getRange2( "D16" ).setValue( "2" );
		ws.getRange2( "D17" ).setValue( "3" );
		ws.getRange2( "D18" ).setValue( "4" );
		ws.getRange2( "D19" ).setValue( "5" );
		ws.getRange2( "D20" ).setValue( "6" );

		ws.getRange2( "E15" ).setValue( "5/1/2011" );
		ws.getRange2( "E16" ).setValue( "5/2/2011" );
		ws.getRange2( "E17" ).setValue( "5/3/2011" );
		ws.getRange2( "E18" ).setValue( "5/4/2011" );
		ws.getRange2( "E19" ).setValue( "5/5/2011" );
		ws.getRange2( "E20" ).setValue( "5/6/2011" );

		oParser = new parserFormula( "COUNTIFS(D15:D20,\"<6\",D15:D20,\">1\")", "D1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );

		oParser = new parserFormula( "COUNTIFS(D15:D20,\"<5\",E15:E20,\"<5/3/2011\")", "E1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "COUNTIFS(D15:D20,\"<\" & D19,E15:E20,\"<\" & E17)", "E1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );


		testArrayFormulaEqualsValues(assert, "1,1,1,#N/A;1,1,1,#N/A;#N/A,#N/A,#N/A,#N/A", "COUNTIFS(A1:C2,A1:C2,A1:C2,A1:C2, A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert, "1,0,0,#N/A;1,0,0,#N/A;#N/A,#N/A,#N/A,#N/A", "COUNTIFS(A1:C2,A1:A2,A1:C2,A1:C2,A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert, "#VALUE!,#VALUE!,#VALUE!,#N/A;#VALUE!,#VALUE!,#VALUE!,#N/A;#N/A,#N/A,#N/A,#N/A", "COUNTIFS(A1:C2,A1:C2,A1:A2,A1:C2,A1:A2,A1:C2)");


		ws.getRange2( "DS2" ).setValue( "12" );
		ws.getRange2( "DS3" ).setValue( "2" );
		ws.getRange2( "DS4" ).setValue( "3" );
		ws.getRange2( "DS5" ).setValue( "4" );
		ws.getRange2( "DS6" ).setValue( "e" );
		ws.getRange2( "DS10" ).setValue( "12" );

		ws.getRange2( "DU2" ).setValue( "1" );
		ws.getRange2( "DU3" ).setValue( "2" );
		ws.getRange2( "DU4" ).setValue( "11" );
		ws.getRange2( "DU5" ).setValue( "11" );
		ws.getRange2( "DU6" ).setValue( "w" );
		ws.getRange2( "DU9" ).setValue( "12" );
		ws.getRange2( "DU10" ).setValue( "12" );

		oParser = new parserFormula( 'COUNTIFS(DS:DS,">3",DU:DU,">10")', "E1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( 'COUNTIFS(DS:DS,">3",DU1:DU2,">10")', "E1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( 'COUNTIFS(DS2:DS3,">3",DU:DU,">10")', "E1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( 'COUNTIFS(DS:DS,"e",DU:DU,"w")', "E1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( 'COUNTIFS(DS1:DS10,"<5",DU1:DU10,">2")', "E1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( 'COUNTIFS(DS1:DS11,"<5",DU1:DU10,">2")', "E1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		oParser = new parserFormula( 'COUNTIFS(DS1:DS10,"<5",DU1:DU11,">2")', "E1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

    } );

	QUnit.test("Test: \"COUNTIF\"", function (assert) {

		ws.getRange2( "A7" ).setValue( "3" );
		ws.getRange2( "B7" ).setValue( "10" );
		ws.getRange2( "C7" ).setValue( "7" );
		ws.getRange2( "D7" ).setValue( "10" );

		ws.getRange2( "A8" ).setValue( "apples" );
		ws.getRange2( "B8" ).setValue( "oranges" );
		ws.getRange2( "C8" ).setValue( "grapes" );
		ws.getRange2( "D8" ).setValue( "melons" );


		oParser = new parserFormula( "COUNTIF(A7:D7,\"=10\")", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "COUNTIF(A7:D7,\">5\")", "B1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "COUNTIF(A7:D7,\"<>10\")", "C1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "COUNTIF(A8:D8,\"*es\")", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "COUNTIF(A8:D8,\"??a*\")", "B2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "COUNTIF(A8:D8,\"*l*\")", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );


		wb.dependencyFormulas.unlockRecal();

		ws.getRange2( "CC1" ).setValue( "1" );
		ws.getRange2( "CC2" ).setValue( "0" );
		ws.getRange2( "CC3" ).setValue( "1" );
		ws.getRange2( "CC4" ).setValue( "true" );
		ws.getRange2( "CC5" ).setValue( "=true" );
		ws.getRange2( "CC6" ).setValue( "=true()" );
		ws.getRange2( "CC7" ).setValue( "'true'" );
		ws.getRange2( "CC8" ).setValue( "" );

		/*oParser = new parserFormula( "COUNTIF(CC1:CC8,\"<\"&\"F007\")", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );*/

		oParser = new parserFormula( "COUNTIF(CC1:CC7, TRUE())", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "COUNTIF(CC1:CC7, TRUE)", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "COUNTIF(CC1:CC7, 1)", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "COUNTIF(CC1:CC7, 0)", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		ws.getRange2( "CC8" ).setValue( ">3" );
		oParser = new parserFormula( "COUNTIF(CC8,\">3\")", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		ws.getRange2( "CC8" ).setValue( ">3" );
		oParser = new parserFormula( "COUNTIF(CC8,\"=>3\")", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		ws.getRange2( "CC9" ).setValue( "=NA()" );
		ws.getRange2( "CC10" ).setValue( "#N/A" );

		oParser = new parserFormula( "COUNTIF(CC9:CC10,\"#N/A\")", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "COUNTIF(CC9:CC10, NA())", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "COUNTIF(CC9:CC10,\"=NA()\")", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "COUNTIF(#REF!, 1)", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "COUNTIF(CC1:CC8,\">=1\")", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "COUNTIF(CC1:CC8,\"=1\")", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "COUNTIF(CC1:CC8,\"<1\")", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( "COUNTIF(CC1:CC8,\">1\")", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "COUNTIF(CC1:CC8,\"=\"&CC8)", "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		wb.dependencyFormulas.lockRecal();

		ws.getRange2( "A22" ).setValue( "apples" );
		ws.getRange2( "A23" ).setValue( "" );
		ws.getRange2( "A24" ).setValue( "oranges" );
		ws.getRange2( "A25" ).setValue( "peaches" );
		ws.getRange2( "A26" ).setValue( "" );
		ws.getRange2( "A27" ).setValue( "apples" );

		oParser = new parserFormula( 'COUNTIF(A22:A27,"*es")', "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );

		oParser = new parserFormula( 'COUNTIF(A22:A27,"?????es")', "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( 'COUNTIF(A22:A27,"*")', "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );

		oParser = new parserFormula( 'COUNTIF(A22:A27,"<>"&"***")', "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( 'COUNTIF(A22:A27,"<>"&"*")', "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( 'COUNTIF(A22:A27,"<>"&"?")', "C2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 6 );

		testArrayFormula2(assert, "COUNTIF", 2, 2)
	} );

	QUnit.test("Test: \"COUNTBLANK\"", function (assert) {

		ws.getRange2( "A22" ).setValue( "6" );
		ws.getRange2( "A23" ).setValue( "" );
		ws.getRange2( "A24" ).setValue( "4" );

		ws.getRange2( "B22" ).setValue( "" );
		ws.getRange2( "B23" ).setValue( "27" );
		ws.getRange2( "B24" ).setValue( "34" );

		oParser = new parserFormula( "COUNTBLANK(A22:B24)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "COUNTBLANK(A22)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "COUNTBLANK(A23)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );
	} );

    QUnit.test("Test: \"COVAR\"", function (assert) {

        oParser = new parserFormula( "COVAR({2.532,5.621;2.1,3.4},{5.32,2.765;5.2,6.7})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), -1.3753740625 ), true );

        oParser = new parserFormula( "COVAR({1,2},{4,5})", "B1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), 0.25 ), true );

		testArrayFormula2(assert, "COVAR", 2, 2, null, true)

    } );

	QUnit.test("Test: \"COVARIANCE.P\"", function (assert) {

		ws.getRange2( "AA1" ).setValue( "3" );
		ws.getRange2( "AA2" ).setValue( "2" );
		ws.getRange2( "AA4" ).setValue( "4" );
		ws.getRange2( "AA5" ).setValue( "5" );
		ws.getRange2( "AA6" ).setValue( "6" );

		ws.getRange2( "BB1" ).setValue( "9" );
		ws.getRange2( "BB2" ).setValue( "7" );
		ws.getRange2( "BB4" ).setValue( "12" );
		ws.getRange2( "BB5" ).setValue( "15" );
		ws.getRange2( "BB6" ).setValue( "17" );


		oParser = new parserFormula( "COVARIANCE.P(AA1:AA6, BB1:BB6)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual(oParser.calculate().getValue(), 5.2 );

		testArrayFormula2(assert, "COVARIANCE.P", 2, 2, null, true);
	} );

	QUnit.test("Test: \"COVARIANCE.S\"", function (assert) {

		ws.getRange2( "AAA1" ).setValue( "2" );
		ws.getRange2( "AAA2" ).setValue( "4" );
		ws.getRange2( "AAA3" ).setValue( "8" );

		ws.getRange2( "BBB1" ).setValue( "5" );
		ws.getRange2( "BBB2" ).setValue( "11" );
		ws.getRange2( "BBB3" ).setValue( "12" );

		oParser = new parserFormula( "COVARIANCE.S({2,4,8},{5,11,12})", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 9.666666667 );

		oParser = new parserFormula( "COVARIANCE.S(AAA1:AAA3,BBB1:BBB3)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 9.666666667 );

		testArrayFormula2(assert, "COVARIANCE.S", 2, 2, null, true);
	} );

    QUnit.test("Test: \"CRITBINOM\"", function (assert) {

        oParser = new parserFormula( "CRITBINOM(6,0.5,0.75)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4 );

        oParser = new parserFormula( "CRITBINOM(12,0.3,0.95)", "B1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 6 );

        oParser = new parserFormula( "CRITBINOM(-12,0.3,0.95)", "B1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "CRITBINOM(-12,1.3,0.95)", "B1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "CRITBINOM(-12,-1.3,0.95)", "B1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "CRITBINOM(-12,0,0.95)", "B1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "CRITBINOM(-12,0.3,1.95)", "B1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        testArrayFormula2(assert, "CRITBINOM", 3, 3);
    } );

	QUnit.test("Test: \"CONCAT\"", function (assert) {

		ws.getRange2( "AA1" ).setValue( "a1" );
		ws.getRange2( "AA2" ).setValue( "a2" );
		ws.getRange2( "AA4" ).setValue( "a4" );
		ws.getRange2( "AA5" ).setValue( "a5" );
		ws.getRange2( "AA6" ).setValue( "a6" );
		ws.getRange2( "AA7" ).setValue( "a7" );

		ws.getRange2( "BB1" ).setValue( "b1" );
		ws.getRange2( "BB2" ).setValue( "b2" );
		ws.getRange2( "BB4" ).setValue( "b4" );
		ws.getRange2( "BB5" ).setValue( "b5" );
		ws.getRange2( "BB6" ).setValue( "b6" );
		ws.getRange2( "BB7" ).setValue( "b7" );

		oParser = new parserFormula('CONCAT("The"," ","sun"," ","will"," ","come"," ","up"," ","tomorrow.")', "A3", ws);
		assert.ok(oParser.parse(), "CONCAT(AA:AA, BB:BB)");
		assert.strictEqual(oParser.calculate().getValue(), "The sun will come up tomorrow.", "CONCAT(AA:AA, BB:BB)");

	    oParser = new parserFormula("CONCAT(AA:AA, BB:BB)", "A3", ws);
		assert.ok(oParser.parse(), "CONCAT(AA:AA, BB:BB)");
		assert.strictEqual(oParser.calculate().getValue(), "a1a2a4a5a6a7b1b2b4b5b6b7", "CONCAT(AA:AA, BB:BB)");

		oParser = new parserFormula("CONCAT(AA1:BB7)", "A3", ws);
		assert.ok(oParser.parse(), "CONCAT(AA1:BB7)");
		assert.strictEqual(oParser.calculate().getValue(), "a1b1a2b2a4b4a5b5a6b6a7b7", "CONCAT(AA1:BB7)");

		oParser = new parserFormula( 'CONCAT(TRUE,"test")', "A2", ws );
		assert.ok( oParser.parse(), 'CONCAT(TRUE,"test")' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUEtest", 'CONCAT(TRUE,"test")');

		testArrayFormulaEqualsValues(assert, "13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245;13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245;13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245", "CONCAT(A1:C2,A1:C2,A1:C2)")
	});

	QUnit.test("Test: \"CONCATENATE\"", function (assert) {

		ws.getRange2( "AA2" ).setValue( "brook trout" );
		ws.getRange2( "AA3" ).setValue( "species" );
		ws.getRange2( "AA4" ).setValue( "32" );

		ws.getRange2( "AB2" ).setValue( "Andreas" );
		ws.getRange2( "AB3" ).setValue( "Fourth" );

		ws.getRange2( "AC2" ).setValue( "Hauser" );
		ws.getRange2( "AC3" ).setValue( "Pine" );

		oParser = new parserFormula( 'CONCATENATE("Stream population for ", AA2, " ", AA3, " is ", AA4, "/mile.")', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "Stream population for brook trout species is 32/mile." );

		oParser = new parserFormula( 'CONCATENATE(AB2, " ", AC2)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "Andreas Hauser" );

		oParser = new parserFormula( 'CONCATENATE(AC2, ", ", AB2)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "Hauser, Andreas" );

		oParser = new parserFormula( 'CONCATENATE(AB3, " & ", AC3)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "Fourth & Pine" );

		oParser = new parserFormula( 'CONCATENATE(TRUE,"test")', "A2", ws );
		assert.ok( oParser.parse(), 'CONCATENATE(TRUE,"test")' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUEtest", 'CONCATENATE(TRUE,"test")');

		testArrayFormula2(assert, "CONCATENATE", 1, 8);
	});

	QUnit.test("Test: \"&\"", function (assert) {
		ws.getRange2( "AAA101" ).setValue( "1" );
		ws.getRange2( "AAA102" ).setValue( "2" );
		ws.getRange2( "AAB101" ).setValue( "3" );
		ws.getRange2( "AAB102" ).setValue( "4" );

		ws.getRange2( "AAD101" ).setValue( "2" );
		ws.getRange2( "AAD102" ).setValue( "3" );
		ws.getRange2( "AAE101" ).setValue( "4" );
		ws.getRange2( "AAE102" ).setValue( "5" );
		ws.getRange2( "AAF101" ).setValue( "test" );
		ws.getRange2( "AAF102" ).setValue( "test" );

		var array;
		oParser = new parserFormula( "1&AAA101:AAB102", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		assert.strictEqual( array.getElementRowCol(0,0).getValue(), "11");
		assert.strictEqual( array.getElementRowCol(1,0).getValue(), "12");
		assert.strictEqual( array.getElementRowCol(0,1).getValue(), "13");
		assert.strictEqual( array.getElementRowCol(1,1).getValue(), "14");

		oParser = new parserFormula( "AAA101:AAB102&AAD101:AAE102", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		assert.strictEqual( array.getElementRowCol(0,0).getValue(), "12");
		assert.strictEqual( array.getElementRowCol(1,0).getValue(), "23");
		assert.strictEqual( array.getElementRowCol(0,1).getValue(), "34");
		assert.strictEqual( array.getElementRowCol(1,1).getValue(), "45");

		oParser = new parserFormula( "AAA101:AAB102&AAD101:AAF102", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		assert.strictEqual( array.getElementRowCol(0,0).getValue(), "12");
		assert.strictEqual( array.getElementRowCol(1,0).getValue(), "23");
		assert.strictEqual( array.getElementRowCol(0,1).getValue(), "34");
		assert.strictEqual( array.getElementRowCol(1,1).getValue(), "45");

		oParser = new parserFormula( "AJ2:AM5&AAA101:AAB102", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		assert.strictEqual( array.getElementRowCol(0,0).getValue(), "1");
		assert.strictEqual( array.getElementRowCol(1,0).getValue(), "2");
		assert.strictEqual( array.getElementRowCol(0,1).getValue(), "3");
		assert.strictEqual( array.getElementRowCol(1,1).getValue(), "4");

		oParser = new parserFormula( "AJ2:AM5&AAA101:AAA102", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		assert.strictEqual( array.getElementRowCol(0,0).getValue(), "1");
		assert.strictEqual( array.getElementRowCol(1,0).getValue(), "2");
		assert.strictEqual( array.getElementRowCol(0,1).getValue(), "1");
		assert.strictEqual( array.getElementRowCol(1,1).getValue(), "2");
		assert.strictEqual( array.getElementRowCol(0,2).getValue(), "1");
		assert.strictEqual( array.getElementRowCol(1,2).getValue(), "2");

		oParser = new parserFormula( "AAA101:AAB101&AAD101:AAF102", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		assert.strictEqual( array.getElementRowCol(0,0).getValue(), "12");
		assert.strictEqual( array.getElementRowCol(1,0).getValue(), "13");
		assert.strictEqual( array.getElementRowCol(0,1).getValue(), "34");
		assert.strictEqual( array.getElementRowCol(1,1).getValue(), "35");

		oParser = new parserFormula( "AAA101&AAD101:AAF102", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		assert.strictEqual( array.getElementRowCol(0,0).getValue(), "12");
		assert.strictEqual( array.getElementRowCol(1,0).getValue(), "13");
		assert.strictEqual( array.getElementRowCol(0,1).getValue(), "14");
		assert.strictEqual( array.getElementRowCol(1,1).getValue(), "15");
		assert.strictEqual( array.getElementRowCol(0,2).getValue(), "1test");
		assert.strictEqual( array.getElementRowCol(1,2).getValue(), "1test");

		oParser = new parserFormula( "1&AAD101:AAF102", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		assert.strictEqual( array.getElementRowCol(0,0).getValue(), "12");
		assert.strictEqual( array.getElementRowCol(1,0).getValue(), "13");
		assert.strictEqual( array.getElementRowCol(0,1).getValue(), "14");
		assert.strictEqual( array.getElementRowCol(1,1).getValue(), "15");
		assert.strictEqual( array.getElementRowCol(0,2).getValue(), "1test");
		assert.strictEqual( array.getElementRowCol(1,2).getValue(), "1test");

		oParser = new parserFormula( "AAD101:AAF102&\"test\"", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		assert.strictEqual( array.getElementRowCol(0,0).getValue(), "2test");
		assert.strictEqual( array.getElementRowCol(1,0).getValue(), "3test");
		assert.strictEqual( array.getElementRowCol(0,1).getValue(), "4test");
		assert.strictEqual( array.getElementRowCol(1,1).getValue(), "5test");
		assert.strictEqual( array.getElementRowCol(0,2).getValue(), "testtest");
		assert.strictEqual( array.getElementRowCol(1,2).getValue(), "testtest");
	});

    QUnit.test("Test: \"DEVSQ\"", function (assert) {
        ws.getRange2( "A1" ).setValue( "5.6" );
        ws.getRange2( "A2" ).setValue( "8.2" );
        ws.getRange2( "A3" ).setValue( "9.2" );

        oParser = new parserFormula( "DEVSQ(5.6,8.2,9.2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), 6.906666666666665 ), true );

        oParser = new parserFormula( "DEVSQ({5.6,8.2,9.2})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), 6.906666666666665 ), true );

        oParser = new parserFormula( "DEVSQ(5.6,8.2,\"9.2\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), 3.379999999999999 ), true );

        oParser = new parserFormula( "DEVSQ(" + ws.getName() + "!A1:A3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), 6.906666666666665 ), true );

        testArrayFormula2(assert, "DEVSQ", 1, 8, null, true);
    } );

    QUnit.test("Test: \"EXPONDIST\"", function (assert) {

        oParser = new parserFormula( "EXPONDIST(0.2,10,FALSE)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), 1.353352832366127 ), true );

        oParser = new parserFormula( "EXPONDIST(2.3,1.5,TRUE)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), 0.968254363621932 ), true );

        testArrayFormula2(assert, "EXPONDIST", 3, 3);
    } );


	QUnit.test("Test: \"SIN(3.1415926)\"", function (assert) {
		oParser = new parserFormula( 'SIN(3.1415926)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), Math.sin( 3.1415926 ) );

		testArrayFormula(assert, "SIN");
	} );

	QUnit.test("Test: \"EXP\"", function (assert) {

		oParser = new parserFormula( "EXP(1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 2.71828183 );

		oParser = new parserFormula( "EXP(2)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 7.3890561 );

		testArrayFormula(assert, "EXP");

	} );

    QUnit.test("Test: \"FISHER\"", function (assert) {

        function fisher( x ) {
            return toFixed( 0.5 * Math.ln( (1 + x) / (1 - x) ) );
        }

        oParser = new parserFormula( "FISHER(-0.43)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), fisher( -.43 ) );

        oParser = new parserFormula( "FISHER(0.578)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), fisher( 0.578 ) );

        oParser = new parserFormula( "FISHER(1)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "FISHER(-1)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        testArrayFormula(assert, "FISHER");

    } );

    QUnit.test("Test: \"FISHERINV\"", function (assert) {

        function fisherInv( x ) {
            return toFixed( ( Math.exp( 2 * x ) - 1 ) / ( Math.exp( 2 * x ) + 1 ) );
        }

        oParser = new parserFormula( "FISHERINV(-0.43)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), fisherInv( -.43 ) );

        oParser = new parserFormula( "FISHERINV(0.578)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), fisherInv( 0.578 ) );

        oParser = new parserFormula( "FISHERINV(1)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), fisherInv( 1 ) );

        oParser = new parserFormula( "FISHERINV(-1)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), fisherInv( -1 ) );

		testArrayFormula(assert, "FISHERINV");

    } );

    QUnit.test("Test: \"FORECAST\"", function (assert) {

        function forecast( fx, y, x ) {

            var fSumDeltaXDeltaY = 0, fSumSqrDeltaX = 0, _x = 0, _y = 0, xLength = 0;
            for ( var i = 0; i < x.length; i++ ) {
                _x += x[i];
                _y += y[i];
                xLength++;
            }

            _x /= xLength;
            _y /= xLength;

            for ( var i = 0; i < x.length; i++ ) {

                var fValX = x[i];
                var fValY = y[i];

                fSumDeltaXDeltaY += ( fValX - _x ) * ( fValY - _y );
                fSumSqrDeltaX += ( fValX - _x ) * ( fValX - _x );

            }

            return toFixed( _y + fSumDeltaXDeltaY / fSumSqrDeltaX * ( fx - _x ) );

        }

        oParser = new parserFormula( "FORECAST(30,{6,7,9,15,21},{20,28,31,38,40})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), forecast( 30, [6, 7, 9, 15, 21], [20, 28, 31, 38, 40] ) );

    } );

	function putDataForForecastEts(){
		ws.getRange2( 'A4' ).setValue( '39814' );
		ws.getRange2( 'A5' ).setValue( '39845' );
		ws.getRange2( 'A6' ).setValue( '39873' );
		ws.getRange2( 'A7' ).setValue( '39904' );
		ws.getRange2( 'A8' ).setValue( '39934' );
		ws.getRange2( 'A9' ).setValue( '39965' );
		ws.getRange2( 'A10' ).setValue( '39995' );
		ws.getRange2( 'A11' ).setValue( '40026' );
		ws.getRange2( 'A12' ).setValue( '40057' );
		ws.getRange2( 'A13' ).setValue( '40087' );
		ws.getRange2( 'A14' ).setValue( '40118' );
		ws.getRange2( 'A15' ).setValue( '40148' );
		ws.getRange2( 'A16' ).setValue( '40179' );
		ws.getRange2( 'A17' ).setValue( '40210' );
		ws.getRange2( 'A18' ).setValue( '40238' );
		ws.getRange2( 'A19' ).setValue( '40269' );
		ws.getRange2( 'A20' ).setValue( '40299' );
		ws.getRange2( 'A21' ).setValue( '40330' );
		ws.getRange2( 'A22' ).setValue( '40360' );
		ws.getRange2( 'A23' ).setValue( '40391' );
		ws.getRange2( 'A24' ).setValue( '40422' );
		ws.getRange2( 'A25' ).setValue( '40452' );
		ws.getRange2( 'A26' ).setValue( '40483' );
		ws.getRange2( 'A27' ).setValue( '40513' );
		ws.getRange2( 'A28' ).setValue( '40544' );
		ws.getRange2( 'A29' ).setValue( '40575' );
		ws.getRange2( 'A30' ).setValue( '40603' );
		ws.getRange2( 'A31' ).setValue( '40634' );
		ws.getRange2( 'A32' ).setValue( '40664' );
		ws.getRange2( 'A33' ).setValue( '40695' );
		ws.getRange2( 'A34' ).setValue( '40725' );
		ws.getRange2( 'A35' ).setValue( '40756' );
		ws.getRange2( 'A36' ).setValue( '40787' );
		ws.getRange2( 'A37' ).setValue( '40817' );
		ws.getRange2( 'A38' ).setValue( '40848' );
		ws.getRange2( 'A39' ).setValue( '40878' );
		ws.getRange2( 'A40' ).setValue( '40909' );
		ws.getRange2( 'A41' ).setValue( '40940' );
		ws.getRange2( 'A42' ).setValue( '40969' );
		ws.getRange2( 'A43' ).setValue( '41000' );
		ws.getRange2( 'A44' ).setValue( '41030' );
		ws.getRange2( 'A45' ).setValue( '41061' );
		ws.getRange2( 'A46' ).setValue( '41091' );
		ws.getRange2( 'A47' ).setValue( '41122' );
		ws.getRange2( 'A48' ).setValue( '41153' );
		ws.getRange2( 'A49' ).setValue( '41183' );
		ws.getRange2( 'A50' ).setValue( '41214' );
		ws.getRange2( 'A51' ).setValue( '41244' );
		ws.getRange2( 'A52' ).setValue( '41275' );
		ws.getRange2( 'A53' ).setValue( '41306' );
		ws.getRange2( 'A54' ).setValue( '41334' );
		ws.getRange2( 'A55' ).setValue( '41365' );
		ws.getRange2( 'A56' ).setValue( '41395' );
		ws.getRange2( 'A57' ).setValue( '41426' );
		ws.getRange2( 'A58' ).setValue( '41456' );
		ws.getRange2( 'A59' ).setValue( '41487' );
		ws.getRange2( 'A60' ).setValue( '41518' );

		ws.getRange2( 'B4' ).setValue( '2644539' );
		ws.getRange2( 'B5' ).setValue( '2359800' );
		ws.getRange2( 'B6' ).setValue( '2925918' );
		ws.getRange2( 'B7' ).setValue( '3024973' );
		ws.getRange2( 'B8' ).setValue( '3177100' );
		ws.getRange2( 'B9' ).setValue( '3419595' );
		ws.getRange2( 'B10' ).setValue( '3649702' );
		ws.getRange2( 'B11' ).setValue( '3650668' );
		ws.getRange2( 'B12' ).setValue( '3191526' );
		ws.getRange2( 'B13' ).setValue( '3249428' );
		ws.getRange2( 'B14' ).setValue( '2971484' );
		ws.getRange2( 'B15' ).setValue( '3074209' );
		ws.getRange2( 'B16' ).setValue( '2785466' );
		ws.getRange2( 'B17' ).setValue( '2515361' );
		ws.getRange2( 'B18' ).setValue( '3105958' );
		ws.getRange2( 'B19' ).setValue( '3139059' );
		ws.getRange2( 'B20' ).setValue( '3380355' );
		ws.getRange2( 'B21' ).setValue( '3612886' );
		ws.getRange2( 'B22' ).setValue( '3765824' );
		ws.getRange2( 'B23' ).setValue( '3771842' );
		ws.getRange2( 'B24' ).setValue( '3356365' );
		ws.getRange2( 'B25' ).setValue( '3490100' );
		ws.getRange2( 'B26' ).setValue( '3163659' );
		ws.getRange2( 'B27' ).setValue( '3167124' );
		ws.getRange2( 'B28' ).setValue( '2883810' );
		ws.getRange2( 'B29' ).setValue( '2610667' );
		ws.getRange2( 'B30' ).setValue( '3129205' );
		ws.getRange2( 'B31' ).setValue( '3200527' );
		ws.getRange2( 'B32' ).setValue( '3547804' );
		ws.getRange2( 'B33' ).setValue( '3766323' );
		ws.getRange2( 'B34' ).setValue( '3935589' );
		ws.getRange2( 'B35' ).setValue( '3917884' );
		ws.getRange2( 'B36' ).setValue( '3564970' );
		ws.getRange2( 'B37' ).setValue( '3602455' );
		ws.getRange2( 'B38' ).setValue( '3326859' );
		ws.getRange2( 'B39' ).setValue( '3441693' );
		ws.getRange2( 'B40' ).setValue( '3211600' );
		ws.getRange2( 'B41' ).setValue( '2998119' );
		ws.getRange2( 'B42' ).setValue( '3472440' );
		ws.getRange2( 'B43' ).setValue( '3563007' );
		ws.getRange2( 'B44' ).setValue( '3820570' );
		ws.getRange2( 'B45' ).setValue( '4107195' );
		ws.getRange2( 'B46' ).setValue( '4284443' );
		ws.getRange2( 'B47' ).setValue( '4356216' );
		ws.getRange2( 'B48' ).setValue( '3819379' );
		ws.getRange2( 'B49' ).setValue( '3844987' );
		ws.getRange2( 'B50' ).setValue( '3478890' );
		ws.getRange2( 'B51' ).setValue( '3443039' );
		ws.getRange2( 'B52' ).setValue( '3204637' );
		ws.getRange2( 'B53' ).setValue( '2966477' );
		ws.getRange2( 'B54' ).setValue( '3593364' );
		ws.getRange2( 'B55' ).setValue( '3604104' );
		ws.getRange2( 'B56' ).setValue( '3933016' );
		ws.getRange2( 'B57' ).setValue( '4146797' );
		ws.getRange2( 'B58' ).setValue( '4176486' );
		ws.getRange2( 'B59' ).setValue( '4347059' );
		ws.getRange2( 'B60' ).setValue( '3781168' );


		ws.getRange2( 'A61' ).setValue( '41548' );
		ws.getRange2( 'A62' ).setValue( '41579' );
		ws.getRange2( 'A63' ).setValue( '41609' );
		ws.getRange2( 'A64' ).setValue( '41640' );
		ws.getRange2( 'A65' ).setValue( '41671' );
		ws.getRange2( 'A66' ).setValue( '41699' );
		ws.getRange2( 'A67' ).setValue( '41730' );
		ws.getRange2( 'A68' ).setValue( '41760' );
		ws.getRange2( 'A69' ).setValue( '41791' );
		ws.getRange2( 'A70' ).setValue( '41821' );
		ws.getRange2( 'A71' ).setValue( '41852' );
		ws.getRange2( 'A72' ).setValue( '41883' );
		ws.getRange2( 'A73' ).setValue( '41913' );
		ws.getRange2( 'A74' ).setValue( '41944' );
		ws.getRange2( 'A75' ).setValue( '41974' );
		ws.getRange2( 'A76' ).setValue( '42005' );
		ws.getRange2( 'A77' ).setValue( '42036' );
		ws.getRange2( 'A78' ).setValue( '42064' );
		ws.getRange2( 'A79' ).setValue( '42095' );
		ws.getRange2( 'A80' ).setValue( '42125' );
		ws.getRange2( 'A81' ).setValue( '42156' );
		ws.getRange2( 'A82' ).setValue( '42186' );
		ws.getRange2( 'A83' ).setValue( '42217' );
		ws.getRange2( 'A84' ).setValue( '42248' );
	}

	QUnit.test("Test: \"FORECAST.ETS\"", function (assert) {
		//результаты данного теста соответсвуют результатам LO, но отличаются от MS!!!

		putDataForForecastEts();

		oParser = new parserFormula( "FORECAST.ETS(A61,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3868499.49723621);

		oParser = new parserFormula( "FORECAST.ETS(A62,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3560200.99816396);

		oParser = new parserFormula( "FORECAST.ETS(A63,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3619491.6524986);

		oParser = new parserFormula( "FORECAST.ETS(A64,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3397521.44972895);

		oParser = new parserFormula( "FORECAST.ETS(A65,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3152698.4854144);

		oParser = new parserFormula( "FORECAST.ETS(A66,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3704079.5812005);

		oParser = new parserFormula( "FORECAST.ETS(A67,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3747546.50043675);

		oParser = new parserFormula( "FORECAST.ETS(A68,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 4042011.75785885);

		oParser = new parserFormula( "FORECAST.ETS(A69,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 4250095.33429725);

		oParser = new parserFormula( "FORECAST.ETS(A70,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 4360538.1411926);

		oParser = new parserFormula( "FORECAST.ETS(A71,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 4463640.2710391);

		oParser = new parserFormula( "FORECAST.ETS(A72,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3963675.88150212);

		oParser = new parserFormula( "FORECAST.ETS(A73,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 4028087.58056954);

		oParser = new parserFormula( "FORECAST.ETS(A74,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3719789.0814973);

		oParser = new parserFormula( "FORECAST.ETS(A75,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3779079.73583193);

		oParser = new parserFormula( "FORECAST.ETS(A76,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3557109.53306228);

		oParser = new parserFormula( "FORECAST.ETS(A77,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3312286.56874774);

		oParser = new parserFormula( "FORECAST.ETS(A78,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3863667.66453383);

		oParser = new parserFormula( "FORECAST.ETS(A79,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 3907134.58377009);

		oParser = new parserFormula( "FORECAST.ETS(A80,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 4201599.84119218);

		oParser = new parserFormula( "FORECAST.ETS(A81,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 4409683.41763059);

		oParser = new parserFormula( "FORECAST.ETS(A82,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 4520126.22452593);

		oParser = new parserFormula( "FORECAST.ETS(A83,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 4623228.35437243);

		oParser = new parserFormula( "FORECAST.ETS(A84,B4:B60,A4:A60,1,1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 8 ) - 0, 4123263.96483545);

	} );

	QUnit.test("Test: \"FORECAST.ETS.SEASONALITY\"", function (assert) {
		//результаты данного теста соответсвуют результатам LO, но отличаются от MS!!!

		putDataForForecastEts();

		oParser = new parserFormula("FORECAST.ETS.SEASONALITY(B4:B60,A4:A60,1,1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 12);
	} );

	QUnit.test("Test: \"FORECAST.ETS.STAT\"", function (assert) {
		//результаты данного теста соответсвуют результатам LO, но отличаются от MS!!!

		putDataForForecastEts();

		oParser = new parserFormula("FORECAST.ETS.STAT(B4:B60,A4:A60,1,1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue().toFixed( 8 ) - 0, 0.65234375);
	} );

	QUnit.test("Test: \"FORECAST.LINEAR\"", function (assert) {
		oParser = new parserFormula( "FORECAST(30,{6,7,9,15,21},{20,28,31,38,40})", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 13 ) - 0, 10.6072530864198);
	} );

	QUnit.test("FORMULATEXT", function (assert) {
		wb.dependencyFormulas.unlockRecal();

		ws.getRange2( "S101" ).setValue( "=TODAY()" );
		ws.getRange2( "S102" ).setValue( "" );
		ws.getRange2( "S103" ).setValue( "=1+1" );

		oParser = new parserFormula( "FORMULATEXT(S101)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "=TODAY()" );

		oParser = new parserFormula( "FORMULATEXT(S101:S102)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "=TODAY()" );

		oParser = new parserFormula( "FORMULATEXT(S102)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

		oParser = new parserFormula( "FORMULATEXT(S100:105)", "A1", ws );
		assert.ok( oParser.parse() );
		//"#N/A" - в ms excel
		assert.strictEqual( oParser.calculate().getValue(), newFormulaParser ? "#N/A" : "#VALUE!");

		oParser = new parserFormula( "FORMULATEXT(S103)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "=1+1" );

		oParser = new parserFormula( "FORMULATEXT(#REF!)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		wb.dependencyFormulas.lockRecal();

		testArrayFormulaEqualsValues(assert, "#N/A,#N/A,#N/A,#N/A;#N/A,#N/A,#N/A,#N/A;#N/A,#N/A,#N/A,#N/A", "FORMULATEXT(A1:C2)");
	} );

    QUnit.test("Test: \"FREQUENCY\"", function (assert) {

        ws.getRange2( "A202" ).setValue( "79" );
        ws.getRange2( "A203" ).setValue( "85" );
        ws.getRange2( "A204" ).setValue( "78" );
        ws.getRange2( "A205" ).setValue( "85" );
        ws.getRange2( "A206" ).setValue( "50" );
        ws.getRange2( "A207" ).setValue( "81" );
        ws.getRange2( "A208" ).setValue( "95" );
        ws.getRange2( "A209" ).setValue( "88" );
        ws.getRange2( "A210" ).setValue( "97" );

        ws.getRange2( "B202" ).setValue( "70" );
        ws.getRange2( "B203" ).setValue( "89" );
        ws.getRange2( "B204" ).setValue( "79" );

        oParser = new parserFormula( "FREQUENCY(A202:A210,B202:B204)", "A201", ws );
        assert.ok( oParser.parse() );
        var a = oParser.calculate();
        assert.strictEqual( a.getElement( 0 ).getValue(), 1 );
        assert.strictEqual( a.getElement( 1 ).getValue(), 2 );
        assert.strictEqual( a.getElement( 2 ).getValue(), 4 );
        assert.strictEqual( a.getElement( 3 ).getValue(), 2 );
    } );

    QUnit.test("Test: \"GAMMALN\"", function (assert) {

        oParser = new parserFormula( "GAMMALN(4.5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue().toFixed( 14 ) - 0, 2.45373657084244 );

        oParser = new parserFormula( "GAMMALN(-4.5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        testArrayFormula(assert, "GAMMALN");

    } );

	QUnit.test("Test: \"GAMMALN.PRECISE\"", function (assert) {

		oParser = new parserFormula( "GAMMALN.PRECISE(4)", "A1", ws );
		assert.ok( oParser.parse(), "GAMMALN.PRECISE(4)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed( 7 ) - 0, 1.7917595, "GAMMALN.PRECISE(4)" );

		oParser = new parserFormula( "GAMMALN.PRECISE(-4.5)", "A1", ws );
		assert.ok( oParser.parse(), "GAMMALN.PRECISE(-4.5)" );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", "GAMMALN.PRECISE(-4.5)" );

		testArrayFormula2(assert, "GAMMALN.PRECISE", 1, 1);
	} );

    QUnit.test("Test: \"GEOMEAN\"", function (assert) {

        function geommean( x ) {

            var s1 = 0, _x = 1, xLength = 0, _tx;
            for ( var i = 0; i < x.length; i++ ) {
                _x *= x[i];
            }

            return  Math.pow( _x, 1 / x.length )
        }

        oParser = new parserFormula( "GEOMEAN(10.5,5.3,2.9)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), geommean( [10.5, 5.3, 2.9] ) );

        oParser = new parserFormula( "GEOMEAN(10.5,{5.3,2.9},\"12\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), geommean( [10.5, 5.3, 2.9, 12] ) );

        oParser = new parserFormula( "GEOMEAN(10.5,{5.3,2.9},\"12\",0)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

    } );

    QUnit.test("Test: \"HARMEAN\"", function (assert) {

        function harmmean( x ) {

            var _x = 0, xLength = 0;
            for ( var i = 0; i < x.length; i++ ) {
                _x += 1 / x[i];
                xLength++;
            }
            return xLength / _x;
        }

        oParser = new parserFormula( "HARMEAN(10.5,5.3,2.9)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), harmmean( [10.5, 5.3, 2.9] ) );

        oParser = new parserFormula( "HARMEAN(10.5,{5.3,2.9},\"12\")", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), harmmean( [10.5, 5.3, 2.9, 12] ) );

        oParser = new parserFormula( "HARMEAN(10.5,{5.3,2.9},\"12\",0)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        testArrayFormula2(assert, "HARMEAN", 1, 8, null, true);
    } );

    QUnit.test("Test: \"HYPGEOMDIST\"", function (assert) {

        function hypgeomdist( x, n, M, N ) {
            return toFixed( Math.binomCoeff( M, x ) * Math.binomCoeff( N - M, n - x ) / Math.binomCoeff( N, n ) );
        }

        oParser = new parserFormula( "HYPGEOMDIST(1,4,8,20)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), hypgeomdist( 1, 4, 8, 20 ) );

        oParser = new parserFormula( "HYPGEOMDIST(1,4,8,20)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), hypgeomdist( 1, 4, 8, 20 ) );

        oParser = new parserFormula( "HYPGEOMDIST(-1,4,8,20)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "HYPGEOMDIST(5,4,8,20)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        testArrayFormula2(assert, "HYPGEOMDIST", 4, 4);
    } );

	QUnit.test("Test: \"HYPGEOM.DIST\"", function (assert) {

		oParser = new parserFormula( "HYPGEOM.DIST(1,4,8,20,TRUE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(4) - 0, 0.4654 );

		oParser = new parserFormula( "HYPGEOM.DIST(1,4,8,20,FALSE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(4) - 0, 0.3633 );

		oParser = new parserFormula( "HYPGEOM.DIST(2,2,3,40,0)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.003846154);

		oParser = new parserFormula( "HYPGEOM.DIST(2,3,3,40,5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.999898785);

		oParser = new parserFormula( "HYPGEOM.DIST(1,2,3,4,5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue() - 0, 0.5);

		testArrayFormula2(assert, "HYPGEOM.DIST", 5, 5);
	} );

	QUnit.test("Test: \"HYPLINK\"", function (assert) {

		ws.getRange2( "D101" ).setValue( "" );
		ws.getRange2( "D102" ).setValue( "123" );

		oParser = new parserFormula( 'HYPERLINK("http://example.microsoft.com/report/budget report.xlsx", "Click for report")', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "Click for report" );
		assert.strictEqual( oParser.value.hyperlink, "http://example.microsoft.com/report/budget report.xlsx" );

		oParser = new parserFormula( 'HYPERLINK("[http://example.microsoft.com/report/budget report.xlsx]Annual!F10", D1)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue() - 0, 0 );
		assert.strictEqual( oParser.value.hyperlink, "[http://example.microsoft.com/report/budget report.xlsx]Annual!F10" );

		oParser = new parserFormula( 'HYPERLINK("http://example.microsoft.com/Annual Report.docx]QrtlyProfits", "Quarterly Profit Report")', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "Quarterly Profit Report" );
		assert.strictEqual( oParser.value.hyperlink, 'http://example.microsoft.com/Annual Report.docx]QrtlyProfits' );

		oParser = new parserFormula( 'HYPERLINK("\\FINANCE\Statements\1stqtr.xlsx",D101)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue() - 0, 0 );
		assert.strictEqual( oParser.value.hyperlink, '\\FINANCE\Statements\1stqtr.xlsx' );

		oParser = new parserFormula( 'HYPERLINK("http://test.com")', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "http://test.com" );
		assert.strictEqual( oParser.value.hyperlink, "http://test.com" );

		oParser = new parserFormula( 'HYPERLINK(D101,111)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 111 );
		assert.strictEqual( oParser.value.hyperlink - 0, 0 );

		oParser = new parserFormula( 'HYPERLINK(D102,111)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 111 );
		assert.strictEqual( oParser.value.hyperlink, "123" );

		oParser = new parserFormula( 'HYPERLINK(D102)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "123" );
		assert.strictEqual( oParser.value.hyperlink - 0, 123 );

		oParser = new parserFormula( 'HYPERLINK(D101,TRUE)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE" );
		assert.strictEqual( oParser.value.hyperlink - 0, 0 );
	} );

	QUnit.test("Test: \"HOUR\"", function (assert) {

		ws.getRange2( "A202" ).setValue( "0.75" );
		ws.getRange2( "A203" ).setValue( "7/18/2011 7:45" );
		ws.getRange2( "A204" ).setValue( "4/21/2012" );

		oParser = new parserFormula( "HOUR(A202)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 18 );

		oParser = new parserFormula( "HOUR(A203)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 7 );

		oParser = new parserFormula( "HOUR(A204)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		testArrayFormula2(assert, "HOUR", 1, 1);
	} );

    QUnit.test("Test: \"INTERCEPT\"", function (assert) {

        function intercept( y, x ) {

            var fSumDeltaXDeltaY = 0, fSumSqrDeltaX = 0, _x = 0, _y = 0, xLength = 0;
            for ( var i = 0; i < x.length; i++ ) {
                _x += x[i];
                _y += y[i];
                xLength++;
            }

            _x /= xLength;
            _y /= xLength;

            for ( var i = 0; i < x.length; i++ ) {

                var fValX = x[i];
                var fValY = y[i];

                fSumDeltaXDeltaY += ( fValX - _x ) * ( fValY - _y );
                fSumSqrDeltaX += ( fValX - _x ) * ( fValX - _x );

            }

            return toFixed( _y - fSumDeltaXDeltaY / fSumSqrDeltaX * _x );

        }

        oParser = new parserFormula( "INTERCEPT({6,7,9,15,21},{20,28,31,38,40})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), intercept( [6, 7, 9, 15, 21], [20, 28, 31, 38, 40] ) );

		testArrayFormula2(assert, "INTERCEPT", 2, 2, null, true);
    } );

	QUnit.test("Test: \"INT\"", function (assert) {

		ws.getRange2( "A202" ).setValue( "19.5" );

		oParser = new parserFormula( "INT(8.9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 8 );

		oParser = new parserFormula( "INT(-8.9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -9 );

		oParser = new parserFormula( "A202-INT(A202)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0.5 );

		testArrayFormula(assert, "INT");
	} );

    QUnit.test("Test: \"KURT\"", function (assert) {

        function kurt( x ) {

            var sumSQRDeltaX = 0, _x = 0, xLength = 0, standDev = 0, sumSQRDeltaXDivstandDev = 0;
            for ( var i = 0; i < x.length; i++ ) {
                _x += x[i];
                xLength++;
            }

            _x /= xLength;

            for ( var i = 0; i < x.length; i++ ) {
                sumSQRDeltaX += Math.pow( x[i] - _x, 2 );
            }

            standDev = Math.sqrt( sumSQRDeltaX / ( xLength - 1 ) );

            for ( var i = 0; i < x.length; i++ ) {
                sumSQRDeltaXDivstandDev += Math.pow( (x[i] - _x) / standDev, 4 );
            }

            return toFixed( xLength * (xLength + 1) / (xLength - 1) / (xLength - 2) / (xLength - 3) * sumSQRDeltaXDivstandDev - 3 * (xLength - 1) * (xLength - 1) / (xLength - 2) / (xLength - 3) )

        }

        oParser = new parserFormula( "KURT(10.5,12.4,19.4,23.2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), kurt( [10.5, 12.4, 19.4, 23.2] ) );

        oParser = new parserFormula( "KURT(10.5,{12.4,19.4},23.2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), kurt( [10.5, 12.4, 19.4, 23.2] ) );

        oParser = new parserFormula( "KURT(10.5,12.4,19.4)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        testArrayFormula2(assert, "KURT", 1, 8, null, true);
    } );

    QUnit.test("Test: \"LARGE\"", function (assert) {

        oParser = new parserFormula( "LARGE({3,5,3,5,4;4,2,4,6,7},3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 5 );

        oParser = new parserFormula( "LARGE({3,5,3,5,4;4,2,4,6,7},7)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4 );

		//TODO нужна другая функция для тестирования
        //testArrayFormula2(assert, "LARGE", 2, 2)
    } );

	QUnit.test("Test: \"LN\"", function (assert) {

		oParser = new parserFormula( "LN(86)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 4.4543473 );

		oParser = new parserFormula( "LN(2.7182818)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(1) - 0, 1 );

		oParser = new parserFormula( "LN(EXP(3))", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		testArrayFormula(assert, "LN");
	} );

	QUnit.test("Test: \"LOG10\"", function (assert) {

		oParser = new parserFormula( "LOG10(86)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(4) - 0, 1.9345 );

		oParser = new parserFormula( "LOG10(10)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( "LOG10(100000)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 5 );

		oParser = new parserFormula( "LOG10(10^5)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 5 );

		testArrayFormula(assert, "LOG10");
	} );

	QUnit.test("Test: \"LINEST\"", function (assert) {

		ws.getRange2( "A202" ).setValue( "1" );
		ws.getRange2( "A203" ).setValue( "2" );
		ws.getRange2( "A204" ).setValue( "3" );

		ws.getRange2( "B202" ).setValue( "2" );
		ws.getRange2( "B203" ).setValue( "3" );
		ws.getRange2( "B204" ).setValue( "4" );


		oParser = new parserFormula( "LINEST(A202:B204)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 0.54285714);

		oParser = new parserFormula( "LINEST(A202:B204, A202:B204)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 1);

		oParser = new parserFormula( "LINEST(A202:B204, A202:B204, 1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 1);

		oParser = new parserFormula( "LINEST(A202:B204, A202:B204, 1, 1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 1);

		ws.getRange2( "A202" ).setValue( "1" );
		ws.getRange2( "A203" ).setValue( "9" );
		ws.getRange2( "A204" ).setValue( "5" );
		ws.getRange2( "A205" ).setValue( "7" );

		ws.getRange2( "B202" ).setValue( "0" );
		ws.getRange2( "B203" ).setValue( "4" );
		ws.getRange2( "B204" ).setValue( "2" );
		ws.getRange2( "B205" ).setValue( "3" );

		oParser = new parserFormula( "LINEST(A202:A205,B202:B205,,FALSE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 2);

		oParser = new parserFormula( "LINEST(A202:A205,B202:B205,FALSE,FALSE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 2.31034483);

		oParser = new parserFormula( "LINEST(A202:A205,B202:B205,FALSE,TRUE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 2.31034483);

		ws.getRange2( "A102" ).setValue( "1" );
		ws.getRange2( "A103" ).setValue( "9" );
		ws.getRange2( "A104" ).setValue( "5" );
		ws.getRange2( "A105" ).setValue( "7" );

		ws.getRange2( "B102" ).setValue( "0" );
		ws.getRange2( "B103" ).setValue( "4" );
		ws.getRange2( "B104" ).setValue( "2" );
		ws.getRange2( "B105" ).setValue( "3" );

		oParser = new parserFormula( "LINEST(A102:A105,B102:B105,,FALSE)", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E106:F106").bbox);
		assert.ok( oParser.parse() );
		var array = oParser.calculate();
		if(AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual( array.getElementRowCol(0,0).getValue(), 2);
			assert.strictEqual( array.getElementRowCol(0,1).getValue(), 1);
		}


		ws.getRange2( "A102" ).setValue( "2310" );
		ws.getRange2( "A103" ).setValue( "2333" );
		ws.getRange2( "A104" ).setValue( "2356" );
		ws.getRange2( "A105" ).setValue( "2379" );
		ws.getRange2( "A106" ).setValue( "2402" );
		ws.getRange2( "A107" ).setValue( "2425" );
		ws.getRange2( "A108" ).setValue( "2448" );
		ws.getRange2( "A109" ).setValue( "2471" );
		ws.getRange2( "A110" ).setValue( "2494" );
		ws.getRange2( "A111" ).setValue( "2517" );
		ws.getRange2( "A112" ).setValue( "2540" );

		ws.getRange2('B102').setValue('2');
		ws.getRange2('B103').setValue('2');
		ws.getRange2('B104').setValue('3');
		ws.getRange2('B105').setValue('3');
		ws.getRange2('B106').setValue('2');
		ws.getRange2('B107').setValue('4');
		ws.getRange2('B108').setValue('2');
		ws.getRange2('B109').setValue('2');
		ws.getRange2('B110').setValue('3');
		ws.getRange2('B111').setValue('4');
		ws.getRange2('B112').setValue('2');

		ws.getRange2('C102').setValue('2');
		ws.getRange2('C103').setValue('2');
		ws.getRange2('C104').setValue('1.5');
		ws.getRange2('C105').setValue('2');
		ws.getRange2('C106').setValue('3');
		ws.getRange2('C107').setValue('2');
		ws.getRange2('C108').setValue('1.5');
		ws.getRange2('C109').setValue('2');
		ws.getRange2('C110').setValue('3');
		ws.getRange2('C111').setValue('4');
		ws.getRange2('C112').setValue('3');

		ws.getRange2('D102').setValue('20');
		ws.getRange2('D103').setValue('12');
		ws.getRange2('D104').setValue('33');
		ws.getRange2('D105').setValue('43');
		ws.getRange2('D106').setValue('53');
		ws.getRange2('D107').setValue('23');
		ws.getRange2('D108').setValue('99');
		ws.getRange2('D109').setValue('34');
		ws.getRange2('D110').setValue('23');
		ws.getRange2('D111').setValue('55');
		ws.getRange2('D112').setValue('22');

		ws.getRange2('E102').setValue('142000');
		ws.getRange2('E103').setValue('144000');
		ws.getRange2('E104').setValue('151000');
		ws.getRange2('E105').setValue('150000');
		ws.getRange2('E106').setValue('139000');
		ws.getRange2('E107').setValue('169000');
		ws.getRange2('E108').setValue('126000');
		ws.getRange2('E109').setValue('142900');
		ws.getRange2('E110').setValue('163000');
		ws.getRange2('E111').setValue('169000');
		ws.getRange2('E112').setValue('149000');

		oParser = new parserFormula( "LINEST(E102:E112,A102:D112,TRUE,TRUE)", "A1", ws );
		oParser.setArrayFormulaRef(ws.getRange2("E120:E123").bbox);
		assert.ok( oParser.parse() );
		array = oParser.calculate();
		if(AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual( array.getElementRowCol(0,0).getValue().toFixed(7) - 0, -234.2371645);
			assert.strictEqual( array.getElementRowCol(1,0).getValue().toFixed(8) - 0, 13.26801148);
			assert.strictEqual( array.getElementRowCol(2,0).getValue().toFixed(9) - 0, 0.996747993);
			assert.strictEqual( array.getElementRowCol(3,0).getValue().toFixed(7) - 0, 459.7536742);
		}

		ws.getRange2( "A120" ).setValue( "1" );
		ws.getRange2( "B120" ).setValue( "2" );

		oParser = new parserFormula( "LINEST(A120:B120)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1);

		oParser = new parserFormula( "LINEST({1;2})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1);

		oParser = new parserFormula( "LINEST({1,2})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1);

	} );

    QUnit.test("Test: \"MEDIAN\"", function (assert) {

        function median( x ) {

            x.sort(fSortAscending);

            if ( x.length % 2 )
                return x[(x.length - 1) / 2];
            else
                return (x[x.length / 2 - 1] + x[x.length / 2]) / 2;
        }

        oParser = new parserFormula( "MEDIAN(10.5,12.4,19.4,23.2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), median( [10.5, 12.4, 19.4, 23.2] ) );

        oParser = new parserFormula( "MEDIAN(10.5,{12.4,19.4},23.2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), median( [10.5, 12.4, 19.4, 23.2] ) );

        oParser = new parserFormula( "MEDIAN(-3.5,1.4,6.9,-4.5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), median( [-3.5, 1.4, 6.9, -4.5] ) );

        testArrayFormula2(assert, "MEDIAN", 1, 8, null, true);
    } );

    QUnit.test("Test: \"MODE\"", function (assert) {

        function mode( x ) {

            x.sort(AscCommon.fSortAscending);

            if ( x.length < 1 )
                return "#VALUE!";
            else {
                var nMaxIndex = 0, nMax = 1, nCount = 1, nOldVal = x[0], i;

                for ( i = 1; i < x.length; i++ ) {
                    if ( x[i] == nOldVal )
                        nCount++;
                    else {
                        nOldVal = x[i];
                        if ( nCount > nMax ) {
                            nMax = nCount;
                            nMaxIndex = i - 1;
                        }
                        nCount = 1;
                    }
                }
                if ( nCount > nMax ) {
                    nMax = nCount;
                    nMaxIndex = i - 1;
                }
                if ( nMax == 1 && nCount == 1 )
                    return "#VALUE!";
                else if ( nMax == 1 )
                    return nOldVal;
                else
                    return x[nMaxIndex];
            }
        }

        oParser = new parserFormula( "MODE(9,1,5,1,9,5,6,6)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), mode( [9, 1, 5, 1, 9, 5, 6, 6] ) );

        oParser = new parserFormula( "MODE(1,9,5,1,9,5,6,6)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), mode( [1, 9, 5, 1, 9, 5, 6, 6] ) );

        oParser = new parserFormula( "MODE(1,9,5,5,9,5,6,6)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), mode( [1, 9, 5, 5, 9, 5, 6, 6] ) );

        testArrayFormula2(assert, "mode", 1, 8, null, true);
    } );

	QUnit.test("Test: \"MODE.MULT \"", function (assert) {

		ws.getRange2( "F202" ).setValue( "1" );
		ws.getRange2( "F203" ).setValue( "2" );
		ws.getRange2( "F204" ).setValue( "3" );
		ws.getRange2( "F205" ).setValue( "4" );
		ws.getRange2( "F206" ).setValue( "3" );
		ws.getRange2( "F207" ).setValue( "2" );
		ws.getRange2( "F208" ).setValue( "1" );
		ws.getRange2( "F209" ).setValue( "2" );
		ws.getRange2( "F210" ).setValue( "3" );
		ws.getRange2( "F211" ).setValue( "5" );
		ws.getRange2( "F212" ).setValue( "6" );
		ws.getRange2( "F213" ).setValue( "1" );

		oParser = new parserFormula( "MODE.MULT(F202:F213)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

	} );

	QUnit.test("Test: \"MODE.SNGL \"", function (assert) {

		ws.getRange2( "F202" ).setValue( "5.6" );
		ws.getRange2( "F203" ).setValue( "4" );
		ws.getRange2( "F204" ).setValue( "4" );
		ws.getRange2( "F205" ).setValue( "3" );
		ws.getRange2( "F206" ).setValue( "2" );
		ws.getRange2( "F207" ).setValue( "4" );

		oParser = new parserFormula( "MODE.SNGL(F202:F207)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );

	} );

	QUnit.test("Test: \"NUMBERVALUE\"", function (assert) {
		oParser = new parserFormula( 'NUMBERVALUE("2.500,27",",",".")', "A1", ws );
		assert.ok( oParser.parse(), 'NUMBERVALUE("2.500,27",",",".")');
		assert.strictEqual( oParser.calculate().getValue(), 2500.27, 'NUMBERVALUE("2.500,27",",",".")');

		oParser = new parserFormula( 'NUMBERVALUE("3.5%")', "A1", ws );
		assert.ok( oParser.parse(), 'NUMBERVALUE("3.5%")');
		assert.strictEqual( oParser.calculate().getValue(), 0.035, 'NUMBERVALUE("3.5%")');

		oParser = new parserFormula( 'NUMBERVALUE("3.5%%%")', "A1", ws );
		assert.ok( oParser.parse(), 'NUMBERVALUE("3.5%%%")');
		assert.strictEqual( oParser.calculate().getValue(), 0.0000035, 'NUMBERVALUE("3.5%%%")');

		oParser = new parserFormula( 'NUMBERVALUE(123123,6,6)', "A1", ws );
		assert.ok( oParser.parse(), 'NUMBERVALUE(123123,6,6)');
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!", 'NUMBERVALUE(123123,6,6)');

		testArrayFormula2(assert, "NUMBERVALUE", 1, 3);
	});

    QUnit.test("Test: \"NORMDIST\"", function (assert) {

        function normdist( x, mue, sigma, kum ) {
            if ( sigma <= 0 )
                return "#NUM!";
            else if ( kum == false )
                return toFixed( AscCommonExcel.phi( (x - mue) / sigma ) / sigma );
            else
                return toFixed( 0.5 + AscCommonExcel.gauss( (x - mue) / sigma ) );

        }

        oParser = new parserFormula( "NORMDIST(42,40,1.5,TRUE)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), normdist( 42, 40, 1.5, true ) );

        oParser = new parserFormula( "NORMDIST(42,40,1.5,FALSE)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), normdist( 42, 40, 1.5, false ) );

        oParser = new parserFormula( "NORMDIST(42,40,-1.5,TRUE)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), normdist( 42, 40, -1.5, true ) );

        oParser = new parserFormula( "NORMDIST(1,40,-1.5,TRUE)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), normdist( 1, 40, -1.5, true ) );

        testArrayFormula2(assert, "NORMDIST", 4, 4);
    } );

	QUnit.test("Test: \"NORM.DIST \"", function (assert) {

		ws.getRange2( "F202" ).setValue( "42" );
		ws.getRange2( "F203" ).setValue( "40" );
		ws.getRange2( "F204" ).setValue( "1.5" );

		oParser = new parserFormula( "NORM.DIST(F202,F203,F204,TRUE)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.9087888 );

		oParser = new parserFormula( "NORM.DIST(F202,F203,F204,FALSE)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(5) - 0, 0.10934 );

	} );

    QUnit.test("Test: \"NORMSDIST\"", function (assert) {

        function normsdist( x ) {
            return toFixed( 0.5 + AscCommonExcel.gauss( x ) );
        }

        oParser = new parserFormula( "NORMSDIST(1.333333)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), normsdist( 1.333333 ) );

        oParser = new parserFormula( "NORMSDIST(-1.5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), normsdist( -1.5 ) );

        testArrayFormula(assert, "NORMSDIST");

    } );

	QUnit.test("Test: \"NORM.S.DIST\"", function (assert) {

		oParser = new parserFormula( "NORM.S.DIST(1.333333,TRUE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.908788726 );

		oParser = new parserFormula( "NORM.S.DIST(1.333333,FALSE)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.164010148 );

		testArrayFormula2(assert, "NORM.S.DIST", 2, 2)
	} );

	QUnit.test("Test: \"NEGBINOMDIST\"", function (assert) {

        function negbinomdist( x, r, p ) {
            x = parseInt( x );
            r = parseInt( r );
            if ( x < 0 || r < 1 || p < 0 || p > 1 )
                return "#NUM!";
            else
                return toFixed( Math.binomCoeff( x + r - 1, r - 1 ) * Math.pow( p, r ) * Math.pow( 1 - p, x ) );
        }

        oParser = new parserFormula( "NEGBINOMDIST(6,10,0.5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), negbinomdist( 6, 10, 0.5 ) );

        oParser = new parserFormula( "NEGBINOMDIST(6,10,1.5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), negbinomdist( 6, 10, 1.5 ) );

        oParser = new parserFormula( "NEGBINOMDIST(20,10,0.63)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), negbinomdist( 20, 10, 0.63 ) );

        testArrayFormula2(assert, "NEGBINOMDIST", 3, 3);
    } );

	QUnit.test("Test: \"NEGBINOM.DIST \"", function (assert) {

		ws.getRange2( "F202" ).setValue( "10" );
		ws.getRange2( "F203" ).setValue( "5" );
		ws.getRange2( "F204" ).setValue( "0.25" );

		oParser = new parserFormula( "NEGBINOM.DIST(F202,F203,F204,TRUE)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.3135141 );

		oParser = new parserFormula( "NEGBINOM.DIST(F202,F203,F204,FALSE)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0550487 );

		testArrayFormula2(assert, "NEGBINOM.DIST", 4, 4);
	} );

	QUnit.test("Test: \"NEGBINOMDIST \"", function (assert) {

		ws.getRange2( "F202" ).setValue( "10" );
		ws.getRange2( "F203" ).setValue( "5" );
		ws.getRange2( "F204" ).setValue( "0.25" );

		oParser = new parserFormula( "NEGBINOMDIST(F202,F203,F204)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 0.05504866 );
	} );

    QUnit.test("Test: \"NORMSINV\"", function (assert) {

        function normsinv( x ) {
            if ( x <= 0.0 || x >= 1.0 )
                return "#N/A";
            else
                return toFixed( AscCommonExcel.gaussinv( x ) );
        }

        oParser = new parserFormula( "NORMSINV(0.954)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), normsinv( 0.954 ) );

        oParser = new parserFormula( "NORMSINV(0.13)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), normsinv( 0.13 ) );

        oParser = new parserFormula( "NORMSINV(0.6782136)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), normsinv( 0.6782136 ) );

        oParser = new parserFormula( "NORMSINV(1.6782136)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), normsinv( 1.6782136 ) );

        oParser = new parserFormula( "NORMSINV(-1.6782136)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), normsinv( -1.6782136 ) );

        testArrayFormula(assert, "NORMSINV");
    } );

	QUnit.test("Test: \"NORM.S.INV \"", function (assert) {

		oParser = new parserFormula( "NORM.S.INV(0.908789)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 1.3333347 );

	} );

    QUnit.test("Test: \"LOGINV\"", function (assert) {

        function loginv( x, mue, sigma ) {
            if ( sigma <= 0 || x <= 0 || x >= 1 )
                return "#NUM!";
            else
                return toFixed( Math.exp( mue + sigma * ( AscCommonExcel.gaussinv( x ) ) ) );
        }

        oParser = new parserFormula( "LOGINV(0.039084,3.5,1.2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), loginv( 0.039084, 3.5, 1.2 ) );

        oParser = new parserFormula( "LOGINV(0,3.5,1.2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), loginv( 0, 3.5, 1.2 ) );

        oParser = new parserFormula( "LOGINV(0,3.5,1.2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), loginv( 10, 3.5, 1.2 ) );

        oParser = new parserFormula( "LOGINV(0,3.5,1.2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), loginv( -10, 3.5, 1.2 ) );

        testArrayFormula2(assert, "LOGINV", 3, 3);
    } );

    QUnit.test("Test: \"NORMINV\"", function (assert) {

        function norminv( x, mue, sigma ) {
            if ( sigma <= 0.0 || x <= 0.0 || x >= 1.0 )
                return "#NUM!";
            else
                return toFixed( AscCommonExcel.gaussinv( x ) * sigma + mue );
        }

        oParser = new parserFormula( "NORMINV(0.954,40,1.5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), norminv( 0.954, 40, 1.5 ) );

        oParser = new parserFormula( "NORMINV(0.13,100,0.5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), norminv( 0.13, 100, 0.5 ) );

        oParser = new parserFormula( "NORMINV(0.6782136,6,0.005)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), norminv( 0.6782136, 6, 0.005 ) );

        oParser = new parserFormula( "NORMINV(-1.6782136,7,0)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), norminv( -1.6782136, 7, 0 ) );

        testArrayFormula2(assert, "NORMINV", 3, 3);
    } );

	QUnit.test("Test: \"NORM.INV \"", function (assert) {

		ws.getRange2( "F202" ).setValue( "0.908789" );
		ws.getRange2( "F203" ).setValue( "40" );
		ws.getRange2( "F204" ).setValue( "1.5" );

		oParser = new parserFormula( "NORM.INV(F202,F203,F204)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 42.000002 );
	} );

    QUnit.test("Test: \"PEARSON\"", function (assert) {

        function pearson( x, y ) {

            var sumXDeltaYDelta = 0, sqrXDelta = 0, sqrYDelta = 0, _x = 0, _y = 0, xLength = 0;

            if ( x.length != y.length )
                return "#N/A"
            for ( var i = 0; i < x.length; i++ ) {

                _x += x[i]
                _y += y[i]
                xLength++;
            }

            _x /= xLength;
            _y /= xLength;

            for ( var i = 0; i < x.length; i++ ) {

                sumXDeltaYDelta += (x[i] - _x) * (y[i] - _y);
                sqrXDelta += (x[i] - _x) * (x[i] - _x);
                sqrYDelta += (y[i] - _y) * (y[i] - _y);

            }

            if ( sqrXDelta == 0 || sqrYDelta == 0 )
                return "#DIV/0!"
            else
                return toFixed( sumXDeltaYDelta / Math.sqrt( sqrXDelta * sqrYDelta ) );
        }

        oParser = new parserFormula( "PEARSON({9,7,5,3,1},{10,6,1,5,3})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), pearson( [9, 7, 5, 3, 1], [10, 6, 1, 5, 3] ) );

        testArrayFormula2(assert, "PEARSON", 2, 2, null, true)
    } );

    QUnit.test("Test: \"PERCENTILE\"", function (assert) {

        function percentile( A, k ) {

            A.sort(fSortAscending)

            var nSize = A.length;
            if ( A.length < 1 || nSize == 0 )
                return new AscCommonExcel.cError( AscCommonExcel.cErrorType.not_available ).toString();
            else {
                if ( nSize == 1 )
                    return toFixed( A[0] );
                else {
                    var nIndex = Math.floor( k * (nSize - 1) );
                    var fDiff = k * (nSize - 1) - Math.floor( k * (nSize - 1) );
                    if ( fDiff == 0.0 )
                        return toFixed( A[nIndex] );
                    else {
                        return toFixed( A[nIndex] +
                            fDiff * (A[nIndex + 1] - A[nIndex]) );
                    }
                }
            }

        }

        oParser = new parserFormula( "PERCENTILE({1,3,2,4},0.3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), percentile( [1, 3, 2, 4], 0.3 ) );

        oParser = new parserFormula( "PERCENTILE({1,3,2,4},0.75)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), percentile( [1, 3, 2, 4], 0.75 ) );

        //TODO нужна другая функция для тестирования
        //testArrayFormula2(assert, "PERCENTILE", 2, 2, null, true);
    } );

	QUnit.test("Test: \"PERCENTILE.INC\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "1" );
		ws.getRange2( "A3" ).setValue( "2" );
		ws.getRange2( "A4" ).setValue( "3" );
		ws.getRange2( "A5" ).setValue( "4" );

		oParser = new parserFormula( "PERCENTILE.INC(A2:A5,0.3)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1.9 );
	} );

	QUnit.test("Test: \"PERCENTILE.EXC\"", function (assert) {
		ws.getRange2( "A202" ).setValue( "1" );
		ws.getRange2( "A203" ).setValue( "2" );
		ws.getRange2( "A204" ).setValue( "3" );
		ws.getRange2( "A205" ).setValue( "6" );
		ws.getRange2( "A206" ).setValue( "6" );
		ws.getRange2( "A207" ).setValue( "6" );
		ws.getRange2( "A208" ).setValue( "7" );
		ws.getRange2( "A209" ).setValue( "8" );
		ws.getRange2( "A210" ).setValue( "9" );

		oParser = new parserFormula( "PERCENTILE.EXC(A202:A210, 0.25)", "A1", ws );
		assert.ok( oParser.parse(), "PERCENTILE.EXC(A202:A210, 0.25)" );
		assert.strictEqual( oParser.calculate().getValue(), 2.5, "PERCENTILE.EXC(A202:A210, 0.25)" );

		oParser = new parserFormula( "PERCENTILE.EXC(A202:A210, 0)", "A1", ws );
		assert.ok( oParser.parse(), "PERCENTILE.EXC(A202:A210, 0)" );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", "PERCENTILE.EXC(A202:A210, 0)" );

		oParser = new parserFormula( "PERCENTILE.EXC(A202:A210, 0.01)", "A1", ws );
		assert.ok( oParser.parse(), "PERCENTILE.EXC(A202:A210, 0.01)" );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", "PERCENTILE.EXC(A202:A210, 0.01)" );

		oParser = new parserFormula( "PERCENTILE.EXC(A202:A210, 2)", "A1", ws );
		assert.ok( oParser.parse(), "PERCENTILE.EXC(A202:A210, 2)" );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", "PERCENTILE.EXC(A202:A210, 2)" );

		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "PERCENTILE.EXC", 2, 2, null, true)
	} );

    QUnit.test("Test: \"PERCENTRANK\"", function (assert) {

        function percentrank( A, x, k ) {

            var tA = A, t, fNum = x;

            if ( !k ) k = 3;

            tA.sort(fSortAscending);

            var nSize = tA.length;
            if ( tA.length < 1 || nSize == 0 )
                return "#N/A";

            else {
                if ( fNum < tA[0] || fNum > tA[nSize - 1] )
                    return "#N/A";
                else if ( nSize == 1 )
                    return 1
                else {
                    var fRes, nOldCount = 0, fOldVal = tA[0], i;
                    for ( i = 1; i < nSize && tA[i] < fNum; i++ ) {
                        if ( tA[i] != fOldVal ) {
                            nOldCount = i;
                            fOldVal = tA[i];
                        }
                    }
                    if ( tA[i] != fOldVal )
                        nOldCount = i;
                    if ( fNum == tA[i] )
                        fRes = nOldCount / (nSize - 1);
                    else {
                        if ( nOldCount == 0 ) {
                            fRes = 0.0;
                        }
                        else {
                            var fFract = ( fNum - tA[nOldCount - 1] ) /
                                ( tA[nOldCount] - tA[nOldCount - 1] );
                            fRes = ( nOldCount - 1 + fFract ) / (nSize - 1);
                        }
                    }
                    return fRes.toString().substr( 0, fRes.toString().indexOf( "." ) + 1 + k ) - 0;
                }
            }
        }

        oParser = new parserFormula( "PERCENTRANK({12,6,7,9,3,8},4)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), percentrank( [12, 6, 7, 9, 3, 8], 4 ) );

        oParser = new parserFormula( "PERCENTRANK({12,6,7,9,3,8},5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), percentrank( [12, 6, 7, 9, 3, 8], 5 ) );

        //TODO нужен другой тест
        //testArrayFormula2(assert, "PERCENTRANK", 2, 3, null, true);
    } );

	QUnit.test("Test: \"PERCENTRANK.EXC\"", function (assert) {
		ws.getRange2( "A202" ).setValue( "1" );
		ws.getRange2( "A203" ).setValue( "2" );
		ws.getRange2( "A204" ).setValue( "3" );
		ws.getRange2( "A205" ).setValue( "6" );
		ws.getRange2( "A206" ).setValue( "6" );
		ws.getRange2( "A207" ).setValue( "6" );
		ws.getRange2( "A208" ).setValue( "7" );
		ws.getRange2( "A209" ).setValue( "8" );
		ws.getRange2( "A210" ).setValue( "9" );

		oParser = new parserFormula( "PERCENTRANK.EXC(A202:A210, 7)", "A1", ws );
		assert.ok( oParser.parse(), "PERCENTRANK.EXC(A202:A210, 7)" );
		assert.strictEqual( oParser.calculate().getValue(), 0.7, "PERCENTRANK.EXC(A202:A210, 7)" );

		oParser = new parserFormula( "PERCENTRANK.EXC(A202:A210, 5.43)", "A1", ws );
		assert.ok( oParser.parse(), "PERCENTRANK.EXC(A202:A210, 5.43)" );
		assert.strictEqual( oParser.calculate().getValue(), 0.381, "PERCENTRANK.EXC(A202:A210, 5.43)" );

		oParser = new parserFormula( "PERCENTRANK.EXC(A202:A210, 5.43, 1)", "A1", ws );
		assert.ok( oParser.parse(), "PERCENTRANK.EXC(A202:A210, 5.43, 1)" );
		assert.strictEqual( oParser.calculate().getValue(), 0.3, "PERCENTRANK.EXC(A202:A210, 5.43, 1)" );

		//TODO нужен другой тест
		//testArrayFormula2(assert, "PERCENTRANK.EXC", 2, 3, null, true);
	} );

	QUnit.test("Test: \"PERCENTRANK.INC\"", function (assert) {
		ws.getRange2( "A202" ).setValue( "13" );
		ws.getRange2( "A203" ).setValue( "12" );
		ws.getRange2( "A204" ).setValue( "11" );
		ws.getRange2( "A205" ).setValue( "8" );
		ws.getRange2( "A206" ).setValue( "4" );
		ws.getRange2( "A207" ).setValue( "3" );
		ws.getRange2( "A208" ).setValue( "2" );
		ws.getRange2( "A209" ).setValue( "1" );
		ws.getRange2( "A210" ).setValue( "1" );
		ws.getRange2( "A211" ).setValue( "1" );

		oParser = new parserFormula( "PERCENTRANK.INC(A202:A211, 2)", "A1", ws );
		assert.ok( oParser.parse(), "PERCENTRANK.INC(A202:A211, 2)" );
		assert.strictEqual( oParser.calculate().getValue(), 0.333, "PERCENTRANK.INC(A202:A211, 2)" );

		oParser = new parserFormula( "PERCENTRANK.INC(A202:A211, 4)", "A1", ws );
		assert.ok( oParser.parse(), "PERCENTRANK.INC(A202:A211, 4)" );
		assert.strictEqual( oParser.calculate().getValue(), 0.555, "PERCENTRANK.INC(A202:A211, 4)" );

		oParser = new parserFormula( "PERCENTRANK.INC(A202:A211, 8)", "A1", ws );
		assert.ok( oParser.parse(), "PERCENTRANK.INC(A202:A211, 8)" );
		assert.strictEqual( oParser.calculate().getValue(), 0.666, "PERCENTRANK.INC(A202:A211, 8)" );

		oParser = new parserFormula( "PERCENTRANK.INC(A202:A211, 5)", "A1", ws );
		assert.ok( oParser.parse(), "PERCENTRANK.INC(A202:A211, 5)" );
		assert.strictEqual( oParser.calculate().getValue(), 0.583, "PERCENTRANK.INC(A202:A211, 5)" );
	} );

	QUnit.test("Test: \"PERMUT\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "100" );
		ws.getRange2( "A3" ).setValue( "3" );

		oParser = new parserFormula( "PERMUT(A2,A3)", "A1", ws );
		assert.ok( oParser.parse(), "PERMUT(A2,A3)" );
		assert.strictEqual( oParser.calculate().getValue(), 970200, "PERMUT(A2,A3)" );

		oParser = new parserFormula( "PERMUT(3,2)", "A1", ws );
		assert.ok( oParser.parse(), "PERMUT(3,2)" );
		assert.strictEqual( oParser.calculate().getValue(), 6, "PERMUT(3,2)" );

		testArrayFormula2(assert, "PERMUT", 2, 2);
	} );

	QUnit.test("Test: \"PERMUTATIONA\"", function (assert) {
		oParser = new parserFormula( "PERMUTATIONA(3,2)", "A1", ws );
		assert.ok( oParser.parse(), "PERMUTATIONA(3,2)" );
		assert.strictEqual( oParser.calculate().getValue(), 9, "PERMUTATIONA(3,2)" );

		oParser = new parserFormula( "PERMUTATIONA(2,2)", "A1", ws );
		assert.ok( oParser.parse(), "PERMUTATIONA(2,2)" );
		assert.strictEqual( oParser.calculate().getValue(), 4, "PERMUTATIONA(2,2)" );

		testArrayFormula2(assert, "PERMUTATIONA", 2, 2);
	} );

	QUnit.test("Test: \"PHI\"", function (assert) {
		oParser = new parserFormula( "PHI(0.75)", "A1", ws );
		assert.ok( oParser.parse(), "PHI(0.75)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 0.301137432, "PHI(0.75)" );

		testArrayFormula2(assert, "PHI", 1, 1);
	} );

	QUnit.test("Test: \"POISSON\"", function (assert) {

        function poisson( x, l, cumulativeFlag ) {
            var _x = parseInt( x ), _l = l, f = cumulativeFlag;

            if ( f ) {
                var sum = 0;
                for ( var k = 0; k <= x; k++ ) {
                    sum += Math.pow( _l, k ) / Math.fact( k );
                }
                sum *= Math.exp( -_l );
                return toFixed( sum );
            }
            else {
                return toFixed( Math.exp( -_l ) * Math.pow( _l, _x ) / Math.fact( _x ) );
            }

        }

        oParser = new parserFormula( "POISSON(8,2,false)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), poisson( 8, 2, false ) );

        oParser = new parserFormula( "POISSON(8,2,true)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), poisson( 8, 2, true ) );

        oParser = new parserFormula( "POISSON(2.6,5,false)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), poisson( 2, 5, false ) );

        oParser = new parserFormula( "POISSON(2,5.7,true)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), poisson( 2, 5.7, true ) );

        oParser = new parserFormula( "POISSON(-6,5,true)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "POISSON(6,-5,false)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        testArrayFormula2(assert, "POISSON", 3, 3);
    } );

	QUnit.test("Test: \"POISSON.DIST\"", function (assert) {
		ws.getRange2( "A202" ).setValue( "2" );
		ws.getRange2( "A203" ).setValue( "5" );

		oParser = new parserFormula( "POISSON.DIST(A202,A203,TRUE)", "A1", ws );
		assert.ok( oParser.parse(), "POISSON.DIST(A202,A203,TRUE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.124652, "POISSON.DIST(A202,A203,TRUE)" );

		oParser = new parserFormula( "POISSON.DIST(A202,A203,FALSE)", "A1", ws );
		assert.ok( oParser.parse(), "POISSON.DIST(A202,A203,FALSE)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.084224, "POISSON.DIST(A202,A203,FALSE)" );

		testArrayFormula2(assert, "POISSON.DIST", 3, 3);
	} );

    QUnit.test("Test: \"PROB\"", function (assert) {

        oParser = new parserFormula( "PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 0.1 );

        oParser = new parserFormula( "PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 0.8 );

    } );

    QUnit.test("Test: \"PROB\"", function (assert) {

        function quartile( A, k ) {

            var fFlag = k;

            A.sort(fSortAscending);

            var nSize = A.length;
            if ( A.length < 1 || nSize == 0 )
                return "#N/A"
            else {
                if ( nSize == 1 )
                    return toFixed( A[0] );
                else {

                    if ( fFlag < 0.0 || fFlag > 4 )
                        return "#NUM!";
                    else if ( fFlag == 0.0 )
                        return toFixed( A[0] );
                    else if ( fFlag == 1.0 ) {
                        var nIndex = Math.floor( 0.25 * (nSize - 1) ),
                            fDiff = 0.25 * (nSize - 1) - Math.floor( 0.25 * (nSize - 1) );
                        if ( fDiff == 0.0 )
                            return toFixed( A[nIndex] );
                        else {
                            return toFixed( A[nIndex] +
                                fDiff * (A[nIndex + 1] - A[nIndex]) );
                        }
                    }
                    else if ( fFlag == 2.0 ) {
                        if ( nSize % 2 == 0 )
                            return toFixed( (A[nSize / 2 - 1] + A[nSize / 2]) / 2.0 );
                        else
                            return toFixed( A[(nSize - 1) / 2] );
                    }
                    else if ( fFlag == 3.0 ) {
                        var nIndex = Math.floor( 0.75 * (nSize - 1) ),
                            fDiff = 0.75 * (nSize - 1) - Math.floor( 0.75 * (nSize - 1) );
                        if ( fDiff == 0.0 )
                            return toFixed( A[nIndex] );
                        else {
                            return toFixed( A[nIndex] +
                                fDiff * (A[nIndex + 1] - A[nIndex]) );
                        }
                    }
                    else
                        return toFixed( A[nSize - 1] );

                }
            }

        }

        oParser = new parserFormula( "QUARTILE({1,2,4,7,8,9,10,12},-1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), quartile( [1, 2, 4, 7, 8, 9, 10, 12], -1 ) );

        oParser = new parserFormula( "QUARTILE({1,2,4,7,8,9,10,12},0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), quartile( [1, 2, 4, 7, 8, 9, 10, 12], 0 ) );

        oParser = new parserFormula( "QUARTILE({1,2,4,7,8,9,10,12},1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), quartile( [1, 2, 4, 7, 8, 9, 10, 12], 1 ) );

        oParser = new parserFormula( "QUARTILE({1,2,4,7,8,9,10,12},2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), quartile( [1, 2, 4, 7, 8, 9, 10, 12], 2 ) );

        oParser = new parserFormula( "QUARTILE({1,2,4,7,8,9,10,12},3)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), quartile( [1, 2, 4, 7, 8, 9, 10, 12], 3 ) );

        oParser = new parserFormula( "QUARTILE({1,2,4,7,8,9,10,12},4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), quartile( [1, 2, 4, 7, 8, 9, 10, 12], 4 ) );

        oParser = new parserFormula( "QUARTILE({1,2,4,7,8,9,10,12},5)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), quartile( [1, 2, 4, 7, 8, 9, 10, 12], 5 ) );

    } );

	QUnit.test("Test: \"QUARTILE\"", function (assert) {
		ws.getRange2( "A202" ).setValue( "1" );
		ws.getRange2( "A203" ).setValue( "2" );
		ws.getRange2( "A204" ).setValue( "4" );
		ws.getRange2( "A205" ).setValue( "7" );
		ws.getRange2( "A206" ).setValue( "8" );
		ws.getRange2( "A207" ).setValue( "9" );
		ws.getRange2( "A208" ).setValue( "10" );
		ws.getRange2( "A209" ).setValue( "12" );

		oParser = new parserFormula( "QUARTILE(A202:A209,1)", "A1", ws );
		assert.ok( oParser.parse(), "QUARTILE(A202:A209,1)" );
		assert.strictEqual( oParser.calculate().getValue(), 3.5, "QUARTILE(A202:A209,1)" );

		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "QUARTILE", 2, 2)
	} );

    QUnit.test("Test: \"QUARTILE.INC\"", function (assert) {
		ws.getRange2( "A202" ).setValue( "1" );
		ws.getRange2( "A203" ).setValue( "2" );
		ws.getRange2( "A204" ).setValue( "4" );
		ws.getRange2( "A205" ).setValue( "7" );
		ws.getRange2( "A206" ).setValue( "8" );
		ws.getRange2( "A207" ).setValue( "9" );
		ws.getRange2( "A208" ).setValue( "10" );
		ws.getRange2( "A209" ).setValue( "12" );

		oParser = new parserFormula( "QUARTILE.INC(A202:A209,1)", "A1", ws );
		assert.ok( oParser.parse(), "QUARTILE.INC(A202:A209,1)" );
		assert.strictEqual( oParser.calculate().getValue(), 3.5, "QUARTILE.INC(A202:A209,1)" );
	} );

	QUnit.test("Test: \"QUARTILE.EXC\"", function (assert) {
		ws.getRange2( "A202" ).setValue( "6" );
		ws.getRange2( "A203" ).setValue( "7" );
		ws.getRange2( "A204" ).setValue( "15" );
		ws.getRange2( "A205" ).setValue( "36" );
		ws.getRange2( "A206" ).setValue( "39" );
		ws.getRange2( "A207" ).setValue( "40" );
		ws.getRange2( "A208" ).setValue( "41" );
		ws.getRange2( "A209" ).setValue( "42" );
		ws.getRange2( "A210" ).setValue( "43" );
		ws.getRange2( "A211" ).setValue( "47" );
		ws.getRange2( "A212" ).setValue( "49" );

		oParser = new parserFormula( "QUARTILE.EXC(A202:A212,1)", "A1", ws );
		assert.ok( oParser.parse(), "QUARTILE.EXC(A202:A212,1)" );
		assert.strictEqual( oParser.calculate().getValue(), 15, "QUARTILE.EXC(A202:A212,1)" );

		oParser = new parserFormula( "QUARTILE.EXC(A202:A212,3)", "A1", ws );
		assert.ok( oParser.parse(), "QUARTILE.EXC(A202:A212,3)" );
		assert.strictEqual( oParser.calculate().getValue(), 43, "QUARTILE.EXC(A202:A212,3)" );

		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "QUARTILE.EXC", 2, 2)
	} );

    QUnit.test("Test: \"RSQ\"", function (assert) {

        function rsq( x, y ) {

            var sumXDeltaYDelta = 0, sqrXDelta = 0, sqrYDelta = 0, _x = 0, _y = 0, xLength = 0;

            if ( x.length != y.length )
                return "#N/A"
            for ( var i = 0; i < x.length; i++ ) {

                _x += x[i]
                _y += y[i]
                xLength++;
            }

            _x /= xLength;
            _y /= xLength;

            for ( var i = 0; i < x.length; i++ ) {

                sumXDeltaYDelta += (x[i] - _x) * (y[i] - _y);
                sqrXDelta += (x[i] - _x) * (x[i] - _x);
                sqrYDelta += (y[i] - _y) * (y[i] - _y);

            }

            if ( sqrXDelta == 0 || sqrYDelta == 0 )
                return "#DIV/0!"
            else
                return toFixed( Math.pow( sumXDeltaYDelta / Math.sqrt( sqrXDelta * sqrYDelta ), 2 ) );
        }

        oParser = new parserFormula( "RSQ({9,7,5,3,1},{10,6,1,5,3})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), rsq( [9, 7, 5, 3, 1], [10, 6, 1, 5, 3] ) );

        oParser = new parserFormula( "RSQ({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), rsq( [2, 3, 9, 1, 8, 7, 5], [6, 5, 11, 7, 5, 4, 4] ) );

		testArrayFormula2(assert, "RSQ", 2, 2, null, true)
    } );

    QUnit.test("Test: \"SKEW\"", function (assert) {

        function skew( x ) {

            var sumSQRDeltaX = 0, _x = 0, xLength = 0, standDev = 0, sumSQRDeltaXDivstandDev = 0;
            for ( var i = 0; i < x.length; i++ ) {

                _x += x[i];
                xLength++;

            }

            if ( xLength <= 2 )
                return "#N/A"

            _x /= xLength;

            for ( var i = 0; i < x.length; i++ ) {

                sumSQRDeltaX += Math.pow( x[i] - _x, 2 );

            }

            standDev = Math.sqrt( sumSQRDeltaX / ( xLength - 1 ) );

            for ( var i = 0; i < x.length; i++ ) {

                sumSQRDeltaXDivstandDev += Math.pow( (x[i] - _x) / standDev, 3 );

            }

            return toFixed( xLength / (xLength - 1) / (xLength - 2) * sumSQRDeltaXDivstandDev )

        }

        oParser = new parserFormula( "SKEW(3,4,5,2,3,4,5,6,4,7)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), skew( [3, 4, 5, 2, 3, 4, 5, 6, 4, 7] ) );

        oParser = new parserFormula( "SKEW({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), skew( [2, 3, 9, 1, 8, 7, 5, 6, 5, 11, 7, 5, 4, 4] ) );

        testArrayFormula2(assert, "SKEW", 1, 8, null, true);
    } );

	QUnit.test("Test: \"SKEW.P\"", function (assert) {
		ws.getRange2( "A202" ).setValue( "3" );
		ws.getRange2( "A203" ).setValue( "4" );
		ws.getRange2( "A204" ).setValue( "5" );
		ws.getRange2( "A205" ).setValue( "2" );
		ws.getRange2( "A206" ).setValue( "3" );
		ws.getRange2( "A207" ).setValue( "4" );
		ws.getRange2( "A208" ).setValue( "5" );
		ws.getRange2( "A209" ).setValue( "6" );
		ws.getRange2( "A210" ).setValue( "4" );
		ws.getRange2( "A211" ).setValue( "7" );

		oParser = new parserFormula( "SKEW.P(A202:A211)", "A1", ws );
		assert.ok( oParser.parse(), "SKEW.P(A202:A211)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.303193, "SKEW.P(A202:A211)" );
	} );

    QUnit.test("Test: \"SMALL\"", function (assert) {

        oParser = new parserFormula( "SMALL({3,5,3,5,4;4,2,4,6,7},3)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3 );

        oParser = new parserFormula( "SMALL({3,5,3,5,4;4,2,4,6,7},7)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 5 );

		oParser = new parserFormula( "SMALL({1,TRUE,FALSE,3,4,5,32,5,4,3},9)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		oParser = new parserFormula( "SMALL({1,TRUE,FALSE,3,4,5,32,5,4,3},8)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 32 );

		oParser = new parserFormula( "SMALL({1,TRUE,10,3,4,5,32,5,4,3},10)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		oParser = new parserFormula( "SMALL({1,TRUE,10,3,4,5,32,5,4,3},1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		//TODO нужна другая функция для тестирования
        //testArrayFormula2(assert, "SMALL", 2, 2)
    } );

    QUnit.test("Test: \"SLOPE\"", function (assert) {

        function slope( y, x ) {

            var sumXDeltaYDelta = 0, sqrXDelta = 0, _x = 0, _y = 0, xLength = 0;

            if ( x.length != y.length )
                return "#N/A"
            for ( var i = 0; i < x.length; i++ ) {

                _x += x[i]
                _y += y[i]
                xLength++;
            }

            _x /= xLength;
            _y /= xLength;

            for ( var i = 0; i < x.length; i++ ) {

                sumXDeltaYDelta += (x[i] - _x) * (y[i] - _y);
                sqrXDelta += (x[i] - _x) * (x[i] - _x);

            }

            if ( sqrXDelta == 0 )
                return "#DIV/0!"
            else
                return toFixed( sumXDeltaYDelta / sqrXDelta );
        }

        oParser = new parserFormula( "SLOPE({9,7,5,3,1},{10,6,1,5,3})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), slope( [9, 7, 5, 3, 1], [10, 6, 1, 5, 3] ) );

        oParser = new parserFormula( "SLOPE({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), slope( [2, 3, 9, 1, 8, 7, 5], [6, 5, 11, 7, 5, 4, 4] ) );

        testArrayFormula2(assert, "SLOPE", 2, 2, null, true);
    } );

	QUnit.test("Test: \"STEYX\"", function (assert) {
		ws.getRange2( "A203" ).setValue( "2" );
		ws.getRange2( "A204" ).setValue( "3" );
		ws.getRange2( "A205" ).setValue( "9" );
		ws.getRange2( "A206" ).setValue( "1" );
		ws.getRange2( "A207" ).setValue( "8" );
		ws.getRange2( "A208" ).setValue( "7" );
		ws.getRange2( "A209" ).setValue( "5" );

		ws.getRange2( "B203" ).setValue( "6" );
		ws.getRange2( "B204" ).setValue( "5" );
		ws.getRange2( "B205" ).setValue( "11" );
		ws.getRange2( "B206" ).setValue( "7" );
		ws.getRange2( "B207" ).setValue( "5" );
		ws.getRange2( "B208" ).setValue( "4" );
		ws.getRange2( "B209" ).setValue( "4" );

		oParser = new parserFormula( "STEYX(A203:A209,B203:B209)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 3.305719 );

		testArrayFormula2(assert, "STEYX", 2, 2, null, true);
	} );

    QUnit.test("Test: \"STANDARDIZE\"", function (assert) {

        function STANDARDIZE( x, mean, sigma ) {

            if ( sigma <= 0 )
                return "#NUM!"
            else
                return toFixed( (x - mean) / sigma );
        }

        oParser = new parserFormula( "STANDARDIZE(42,40,1.5)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), STANDARDIZE( 42, 40, 1.5 ) );

        oParser = new parserFormula( "STANDARDIZE(22,12,2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), STANDARDIZE( 22, 12, 2 ) );

        oParser = new parserFormula( "STANDARDIZE(22,12,-2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), STANDARDIZE( 22, 12, -2 ) );

        testArrayFormula2(assert, "STANDARDIZE", 3, 3);
    } );

    QUnit.test("Test: \"STDEV\"", function (assert) {

        function stdev() {
            var average = 0, res = 0;
            for ( var i = 0; i < arguments.length; i++ ) {
                average += arguments[i];
            }
            average /= arguments.length;
            for ( var i = 0; i < arguments.length; i++ ) {
                res += (arguments[i] - average) * (arguments[i] - average);
            }
            return toFixed( Math.sqrt( res / (arguments.length - 1) ) );
        }

        oParser = new parserFormula( "STDEV(123,134,143,173,112,109)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), stdev( 123, 134, 143, 173, 112, 109 ) );

        ws.getRange2( "E400" ).setValue( "\"123\"" );
        ws.getRange2( "E401" ).setValue( "134" );
        ws.getRange2( "E402" ).setValue( "143" );
        ws.getRange2( "E403" ).setValue( "173" );
        ws.getRange2( "E404" ).setValue( "112" );
        ws.getRange2( "E405" ).setValue( "109" );

        oParser = new parserFormula( "STDEV(E400:E405)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), stdev( 134, 143, 173, 112, 109 ) );

    } );

	QUnit.test("Test: \"STDEV.S\"", function (assert) {
		ws.getRange2( "A202" ).setValue( "1345" );
		ws.getRange2( "A203" ).setValue( "1301" );
		ws.getRange2( "A204" ).setValue( "1368" );
		ws.getRange2( "A205" ).setValue( "1322" );
		ws.getRange2( "A206" ).setValue( "1310" );
		ws.getRange2( "A207" ).setValue( "1370" );
		ws.getRange2( "A208" ).setValue( "1318" );
		ws.getRange2( "A209" ).setValue( "1350" );
		ws.getRange2( "A210" ).setValue( "1303" );
		ws.getRange2( "A211" ).setValue( "1299" );

		oParser = new parserFormula( "STDEV.S(A202:A211)", "A1", ws );
		assert.ok( oParser.parse(), "STDEV.S(A202:A211)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 27.46391572, "STDEV.S(A202:A211)" );
	} );

	QUnit.test("Test: \"STDEV.P\"", function (assert) {
		ws.getRange2( "A202" ).setValue( "1345" );
		ws.getRange2( "A203" ).setValue( "1301" );
		ws.getRange2( "A204" ).setValue( "1368" );
		ws.getRange2( "A205" ).setValue( "1322" );
		ws.getRange2( "A206" ).setValue( "1310" );
		ws.getRange2( "A207" ).setValue( "1370" );
		ws.getRange2( "A208" ).setValue( "1318" );
		ws.getRange2( "A209" ).setValue( "1350" );
		ws.getRange2( "A210" ).setValue( "1303" );
		ws.getRange2( "A211" ).setValue( "1299" );

		oParser = new parserFormula( "STDEV.P(A202:A211)", "A1", ws );
		assert.ok( oParser.parse(), "STDEV.P(A202:A211)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 26.05455814, "STDEV.P(A202:A211)" );
	} );


	QUnit.test("Test: \"STDEVA\"", function (assert) {

        ws.getRange2( "E400" ).setValue( "\"123\"" );
        ws.getRange2( "E401" ).setValue( "134" );
        ws.getRange2( "E402" ).setValue( "143" );
        ws.getRange2( "E403" ).setValue( "173" );
        ws.getRange2( "E404" ).setValue( "112" );
        ws.getRange2( "E405" ).setValue( "109" );


        function stdeva() {
            var average = 0, res = 0;
            for ( var i = 0; i < arguments.length; i++ ) {
                average += arguments[i];
            }
            average /= arguments.length;
            for ( var i = 0; i < arguments.length; i++ ) {
                res += (arguments[i] - average) * (arguments[i] - average);
            }
            return toFixed( Math.sqrt( res / (arguments.length - 1) ) );
        }

        oParser = new parserFormula( "STDEVA(123,134,143,173,112,109)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), stdeva( 123, 134, 143, 173, 112, 109 ) );

        oParser = new parserFormula( "STDEVA(123,134,143,173,112,109)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), stdeva( 123, 134, 143, 173, 112, 109 ) );

        testArrayFormula2(assert, "STDEVA", 1, 8, null, true);
    } );

	QUnit.test("Test: \"SWITCH\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "2" );
		ws.getRange2( "A3" ).setValue( "99" );
		ws.getRange2( "A4" ).setValue( "99" );
		ws.getRange2( "A5" ).setValue( "2" );
		ws.getRange2( "A6" ).setValue( "3" );

		oParser = new parserFormula( 'SWITCH(WEEKDAY(A2),1,"Sunday",2,"Monday",3,"Tuesday","No match")', "A1", ws );
		assert.ok( oParser.parse(), 'SWITCH(WEEKDAY(A2),1,"Sunday",2,"Monday",3,"Tuesday","No match")' );
		assert.strictEqual( oParser.calculate().getValue(), "Monday", 'SWITCH(WEEKDAY(A2),1,"Sunday",2,"Monday",3,"Tuesday","No match")' );

		oParser = new parserFormula( 'SWITCH(A3,1,"Sunday",2,"Monday",3,"Tuesday")', "A1", ws );
		assert.ok( oParser.parse(), 'SWITCH(A3,1,"Sunday",2,"Monday",3,"Tuesday")' );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A", 'SWITCH(A3,1,"Sunday",2,"Monday",3,"Tuesday")' );

		oParser = new parserFormula( 'SWITCH(A4,1,"Sunday",2,"Monday",3,"Tuesday","No match")', "A1", ws );
		assert.ok( oParser.parse(), 'SWITCH(A4,1,"Sunday",2,"Monday",3,"Tuesday","No match")' );
		assert.strictEqual( oParser.calculate().getValue(), "No match", 'SWITCH(A4,1,"Sunday",2,"Monday",3,"Tuesday","No match")' );

		oParser = new parserFormula( 'SWITCH(A5,1,"Sunday",7,"Saturday","weekday")', "A1", ws );
		assert.ok( oParser.parse(), 'SWITCH(A5,1,"Sunday",7,"Saturday","weekday")' );
		assert.strictEqual( oParser.calculate().getValue(), "weekday", 'SWITCH(A5,1,"Sunday",7,"Saturday","weekday")' );

		oParser = new parserFormula( 'SWITCH(A6,1,"Sunday",2,"Monday",3,"Tuesday","No match")', "A1", ws );
		assert.ok( oParser.parse(), 'SWITCH(A6,1,"Sunday",2,"Monday",3,"Tuesday","No match")' );
		assert.strictEqual( oParser.calculate().getValue(), "Tuesday", 'SWITCH(A6,1,"Sunday",2,"Monday",3,"Tuesday","No match")' );

		oParser = new parserFormula( 'SWITCH(122,1,"Sunday",2,"Monday",3,"Tuesday","No match")', "A1", ws );
		assert.ok( oParser.parse(), 'SWITCH(122,1,"Sunday",2,"Monday",3,"Tuesday","No match")' );
		assert.strictEqual( oParser.calculate().getValue(), "No match", 'SWITCH(122,1,"Sunday",2,"Monday",3,"Tuesday","No match")' );

		oParser = new parserFormula( 'SWITCH({1,"2asd",3},{12,2,3},{"asd",2,3,4})', "A1", ws );
		assert.ok( oParser.parse(), 'SWITCH({1,"2asd",3},{12,2,3},{"asd",2,3,4})' );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A", 'SWITCH({1,"2asd",3},{12,2,3},{"asd",2,3,4})' );

		oParser = new parserFormula( 'SWITCH({"asd1","2asd",3},{"asd1",1,3},"sdf")', "A1", ws );
		assert.ok( oParser.parse(), 'SWITCH({"asd1","2asd",3},{"asd1",1,3},"sdf")' );
		assert.strictEqual( oParser.calculate().getValue(), "sdf", 'SWITCH({"asd1","2asd",3},{"asd1",1,3},"sdf")' );

		testArrayFormulaEqualsValues(assert, "1,3.123,-4,#N/A;2,4,5,#N/A;#N/A,#N/A,#N/A,#N/A", "SWITCH(A1:C2,A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert, "1,1,1,#N/A;1,1,1,#N/A;#N/A,#N/A,#N/A,#N/A", "SWITCH(A1:C2,A1:C2,A1:A1,A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert, "1,1,1,#N/A;2,2,2,#N/A;#N/A,#N/A,#N/A,#N/A", "SWITCH(A1:C2,A1:C2,A1:A2,A1:C2,A1:A2,A1:C2)");
	} );

    QUnit.test("Test: \"VAR\"", function (assert) {

        function _var( x ) {

            var sumSQRDeltaX = 0, _x = 0, xLength = 0, standDev = 0, sumSQRDeltaXDivstandDev = 0;
            for ( var i = 0; i < x.length; i++ ) {
                _x += x[i];
                xLength++;
            }

            _x /= xLength;

            for ( var i = 0; i < x.length; i++ ) {
                sumSQRDeltaX += Math.pow( x[i] - _x, 2 );
            }

            return toFixed( sumSQRDeltaX / (xLength - 1) )

        }

        oParser = new parserFormula( "VAR(10.5,12.4,19.4,23.2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), _var( [10.5, 12.4, 19.4, 23.2] ) );

        oParser = new parserFormula( "VAR(10.5,{12.4,19.4},23.2)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), _var( [10.5, 12.4, 19.4, 23.2] ) );

        oParser = new parserFormula( "VAR(10.5,12.4,19.4)", "A1", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), _var( [10.5, 12.4, 19.4] ) );

		oParser = new parserFormula( "VAR(1)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );

		oParser = new parserFormula( "VAR({1})", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );

		ws.getRange2( "A202" ).setValue( "1345" );
		ws.getRange2( "A203" ).setValue( "" );
		ws.getRange2( "A204" ).setValue( "" );

		oParser = new parserFormula( "VAR(A202)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );

		oParser = new parserFormula( "VAR(A202:A204)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#DIV/0!" );

		oParser = new parserFormula( "VAR(#REF!)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

    } );

	QUnit.test("Test: \"VAR.P\"", function (assert) {

		ws.getRange2( "A202" ).setValue( "1345" );
		ws.getRange2( "A203" ).setValue( "1301" );
		ws.getRange2( "A204" ).setValue( "1368" );
		ws.getRange2( "A205" ).setValue( "1322" );

		ws.getRange2( "A206" ).setValue( "1310" );
		ws.getRange2( "A207" ).setValue( "1370" );
		ws.getRange2( "A208" ).setValue( "1318" );
		ws.getRange2( "A209" ).setValue( "1350" );

		ws.getRange2( "A210" ).setValue( "1303" );
		ws.getRange2( "A211" ).setValue( "1299" );

	    oParser = new parserFormula( "VAR.P(A202:A211)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(2) - 0, 678.84 );

		testArrayFormula2(assert, "VAR.P", 1, 8, null, true);
	} );

	QUnit.test("Test: \"VAR.S\"", function (assert) {

		ws.getRange2( "A202" ).setValue( "1345" );
		ws.getRange2( "A203" ).setValue( "1301" );
		ws.getRange2( "A204" ).setValue( "1368" );
		ws.getRange2( "A205" ).setValue( "1322" );

		ws.getRange2( "A206" ).setValue( "1310" );
		ws.getRange2( "A207" ).setValue( "1370" );
		ws.getRange2( "A208" ).setValue( "1318" );
		ws.getRange2( "A209" ).setValue( "1350" );

		ws.getRange2( "A210" ).setValue( "1303" );
		ws.getRange2( "A211" ).setValue( "1299" );

		oParser = new parserFormula( "VAR.S(A202:A211)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(2) - 0, 754.27 );

		testArrayFormula2(assert, "VAR.S", 1, 8, null, true);

	} );

	QUnit.test("Test: \"VARPA\"", function (assert) {

		ws.getRange2( "A202" ).setValue( "1345" );
		ws.getRange2( "A203" ).setValue( "1301" );
		ws.getRange2( "A204" ).setValue( "1368" );
		ws.getRange2( "A205" ).setValue( "1322" );

		ws.getRange2( "A206" ).setValue( "1310" );
		ws.getRange2( "A207" ).setValue( "1370" );
		ws.getRange2( "A208" ).setValue( "1318" );
		ws.getRange2( "A209" ).setValue( "1350" );

		ws.getRange2( "A210" ).setValue( "1303" );
		ws.getRange2( "A211" ).setValue( "1299" );

		oParser = new parserFormula( "VARPA(A202:A211)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(2) - 0, 678.84 );

		testArrayFormula2(assert, "VARPA", 1, 8, null, true);
	} );

    /*
    * Lookup and Reference
    */
    QUnit.test("Test: \"HLOOKUP\"", function (assert) {

        ws.getRange2( "A401" ).setValue( "Axles" );ws.getRange2( "B401" ).setValue( "Bearings" );ws.getRange2( "C401" ).setValue( "Bolts" );
        ws.getRange2( "A402" ).setValue( "4" );ws.getRange2( "B402" ).setValue( "6" );ws.getRange2( "C402" ).setValue( "9" );
        ws.getRange2( "A403" ).setValue( "5" );ws.getRange2( "B403" ).setValue( "7" );ws.getRange2( "C403" ).setValue( "10" );
        ws.getRange2( "A404" ).setValue( "6" );ws.getRange2( "B404" ).setValue( "8" );ws.getRange2( "C404" ).setValue( "11" );


        oParser = new parserFormula( "HLOOKUP(\"Axles\",A401:C404,2,TRUE)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 4 );

        oParser = new parserFormula( "HLOOKUP(\"Bearings\",A401:C404,3,FALSE)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 7 );

        oParser = new parserFormula( "HLOOKUP(\"B\",A401:C404,3,TRUE)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 5 );

        oParser = new parserFormula( "HLOOKUP(\"Bolts\",A401:C404,4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 11 );

        oParser = new parserFormula( "HLOOKUP(3,{1,2,3;\"a\",\"b\",\"c\";\"d\",\"e\",\"f\"},2,TRUE)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "c" );


		/*oParser = new parserFormula( "HLOOKUP(1,{1,2,3;2,3,4},2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "HLOOKUP(1,{1,2,3;2,3,4},3,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "HLOOKUP(1,{1,2,3;2,3,4},3,0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "HLOOKUP({2,3,4},{1,2,3;2,3,4},2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "HLOOKUP({2,3,4},{1,2,3;2,3,4},{4,5,6})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "HLOOKUP({2,3,4},{1,2,3;2,3,4},{1,5,6})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "HLOOKUP({2,3,4},{1,2,3;2,3,4;6,7,8},{1,5,6})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "HLOOKUP({5,3,4},{1,2,3;2,3,4;6,7,8},{1,5,6})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "HLOOKUP(4,{1,2,3;2,3,4;6,7,8},2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );

		oParser = new parserFormula( "HLOOKUP(4,{1,2,3;2,3,4;6,7,8},3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 8 );

		oParser = new parserFormula( "HLOOKUP(4,{1,2,3;2,3,4;6,7,8},5)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "HLOOKUP({2,3,4},{1,2,3;2,3,4;6,7,8},1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );*/

    } );

    QUnit.test("Test: \"VLOOKUP\"", function (assert) {

        ws.getRange2( "A501" ).setValue( "Density" );ws.getRange2( "B501" ).setValue( "Bearings" );ws.getRange2( "C501" ).setValue( "Bolts" );
		ws.getRange2( "D501" ).setValue( "" );

        ws.getRange2( "A502" ).setValue( "0.457" );ws.getRange2( "B502" ).setValue( "3.55" );ws.getRange2( "C502" ).setValue( "500" );
        ws.getRange2( "A503" ).setValue( "0.525" );ws.getRange2( "B503" ).setValue( "3.25" );ws.getRange2( "C503" ).setValue( "400" );
        ws.getRange2( "A504" ).setValue( "0.616" );ws.getRange2( "B504" ).setValue( "2.93" );ws.getRange2( "C504" ).setValue( "300" );
        ws.getRange2( "A505" ).setValue( "0.675" );ws.getRange2( "B505" ).setValue( "2.75" );ws.getRange2( "C505" ).setValue( "250" );
        ws.getRange2( "A506" ).setValue( "0.746" );ws.getRange2( "B506" ).setValue( "2.57" );ws.getRange2( "C506" ).setValue( "200" );
        ws.getRange2( "A507" ).setValue( "0.835" );ws.getRange2( "B507" ).setValue( "2.38" );ws.getRange2( "C507" ).setValue( "15" );
        ws.getRange2( "A508" ).setValue( "0.946" );ws.getRange2( "B508" ).setValue( "2.17" );ws.getRange2( "C508" ).setValue( "100" );
        ws.getRange2( "A509" ).setValue( "1.09" );ws.getRange2( "B509" ).setValue( "1.95" );ws.getRange2( "C509" ).setValue( "50" );
        ws.getRange2( "A510" ).setValue( "1.29" );ws.getRange2( "B510" ).setValue( "1.71" );ws.getRange2( "C510" ).setValue( "0" );


        oParser = new parserFormula( "VLOOKUP(1,A502:C510,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2.17 );

        oParser = new parserFormula( "VLOOKUP(1,A502:C510,3,TRUE)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 100.00 );

        oParser = new parserFormula( "VLOOKUP(2,A502:C510,2,TRUE)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1.71 );

		oParser = new parserFormula( "VLOOKUP(1,{1,2,3;2,3,4},2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "VLOOKUP(1,{1,2,3;2,3,4},3,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "VLOOKUP(1,{1,2,3;2,3,4},3,0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "VLOOKUP({2,3,4},{1,2,3;2,3,4},2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "VLOOKUP({2,3,4},{1,2,3;2,3,4},{4,5,6})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "VLOOKUP({2,3,4},{1,2,3;2,3,4},{1,5,6})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "VLOOKUP({2,3,4},{1,2,3;2,3,4;6,7,8},{1,5,6})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "VLOOKUP({5,3,4},{1,2,3;2,3,4;6,7,8},{1,5,6})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( "VLOOKUP(4,{1,2,3;2,3,4;6,7,8},2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "VLOOKUP(4,{1,2,3;2,3,4;6,7,8},3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 4 );

		oParser = new parserFormula( "VLOOKUP(4,{1,2,3;2,3,4;6,7,8},5)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "VLOOKUP({2,3,4},{1,2,3;2,3,4;6,7,8},1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2 );

		oParser = new parserFormula( 'VLOOKUP(,A502:C510,2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

		oParser = new parserFormula( 'VLOOKUP(D501,A502:C510,2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

    } );

	QUnit.test("Test: \"LOOKUP\"", function (assert) {

		ws.getRange2( "A102" ).setValue( "4.14" );
		ws.getRange2( "A103" ).setValue( "4.19" );
		ws.getRange2( "A104" ).setValue( "5.17" );
		ws.getRange2( "A105" ).setValue( "5.77" );
		ws.getRange2( "A106" ).setValue( "6.39" );

		ws.getRange2( "B102" ).setValue( "red" );
		ws.getRange2( "B103" ).setValue( "orange" );
		ws.getRange2( "B104" ).setValue( "yellow" );
		ws.getRange2( "B105" ).setValue( "green" );
		ws.getRange2( "B106" ).setValue( "blue" );


		oParser = new parserFormula( "LOOKUP(4.19, A102:A106, B102:B106)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "orange" );

		oParser = new parserFormula( "LOOKUP(5.75, A102:A106, B102:B106)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "yellow" );

		oParser = new parserFormula( "LOOKUP(7.66, A102:A106, B102:B106)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "blue" );

		oParser = new parserFormula( "LOOKUP(0, A102:A106, B102:B106)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

		ws.getRange2( "C101" ).setValue( "4.14" );
		ws.getRange2( "D101" ).setValue( "4.19" );
		ws.getRange2( "E101" ).setValue( "5.17" );
		ws.getRange2( "F101" ).setValue( "5.77" );
		ws.getRange2( "G101" ).setValue( "6.39" );
		ws.getRange2( "H101" ).setValue( "7.99" );

		ws.getRange2( "C102" ).setValue( "red" );
		ws.getRange2( "D102" ).setValue( "orange" );
		ws.getRange2( "E102" ).setValue( "yellow" );
		ws.getRange2( "F102" ).setValue( "green" );
		ws.getRange2( "G102" ).setValue( "blue" );
		ws.getRange2( "H102" ).setValue( "black" );

		oParser = new parserFormula( "LOOKUP(4.19,C101:H101,C102:H102)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "orange" );

		oParser = new parserFormula( "LOOKUP(5.75,C101:H101,C102:H102)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "yellow" );

		oParser = new parserFormula( "LOOKUP(7.66,C101:H101,C102:H102)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "blue" );

		oParser = new parserFormula( "LOOKUP(0,C101:H101,C102:H102)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

		oParser = new parserFormula( "LOOKUP(5.17,C101:H101,C102:H102)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "yellow" );

		oParser = new parserFormula( "LOOKUP(9,C101:H101,C102:H102)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "black" );

		ws.getRange2( "A102" ).setValue( "1" );
		ws.getRange2( "A103" ).setValue( "" );
		ws.getRange2( "A104" ).setValue( "3" );
		ws.getRange2( "A105" ).setValue( "3" );

		ws.getRange2( "B102" ).setValue( "a" );
		ws.getRange2( "B103" ).setValue( "b" );
		ws.getRange2( "B104" ).setValue( "c" );
		ws.getRange2( "B105" ).setValue( "d" );

		oParser = new parserFormula( "LOOKUP(1,A102:A105,B102:B105)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "a" );

		oParser = new parserFormula( "LOOKUP(,A102:A105,B102:B105)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

		ws.getRange2( "C101" ).setValue( "2" );
		ws.getRange2( "D101" ).setValue( "" );
		ws.getRange2( "E101" ).setValue( "4" );
		ws.getRange2( "F101" ).setValue( "5" );
		ws.getRange2( "G101" ).setValue( "" );
		ws.getRange2( "H101" ).setValue( "" );
		ws.getRange2( "I101" ).setValue( "6" );
		ws.getRange2( "J101" ).setValue( "" );
		ws.getRange2( "K101" ).setValue( "7" );

		ws.getRange2( "C102" ).setValue( "a" );
		ws.getRange2( "D102" ).setValue( "b" );
		ws.getRange2( "E102" ).setValue( "c" );
		ws.getRange2( "F102" ).setValue( "d" );
		ws.getRange2( "G102" ).setValue( "e" );
		ws.getRange2( "H102" ).setValue( "f" );
		ws.getRange2( "I102" ).setValue( "g" );
		ws.getRange2( "J102" ).setValue( "h" );
		ws.getRange2( "K102" ).setValue( "i" );

		oParser = new parserFormula( "LOOKUP(3,C101:K101,C102:K102)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "a" );

		oParser = new parserFormula( "LOOKUP(2,C101:K101,C102:K102)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "a" );

		oParser = new parserFormula( "LOOKUP(7,C101:K101,C102:K102)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "i" );

		oParser = new parserFormula( "LOOKUP(10,C101:K101,C102:K102)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), "i" );

		oParser = new parserFormula( "LOOKUP(1,C101:K101,C102:K102)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A" );
	});


	QUnit.test("Test: \"XLOOKUP\"", function (assert) {

		ws.getRange2( "A551" ).setValue( "6" );
		ws.getRange2( "A552" ).setValue( "2" );
		ws.getRange2( "A553" ).setValue( "test" );
		ws.getRange2( "A554" ).setValue( "5" );
		ws.getRange2( "A555" ).setValue( "4" );
		ws.getRange2( "A556" ).setValue( "12" );
		ws.getRange2( "A557" ).setValue( "16" );
		ws.getRange2( "A558" ).setValue( "14" );
		ws.getRange2( "A559" ).setValue( "99" );
		ws.getRange2( "A560" ).setValue( "tt" );
		ws.getRange2( "A561" ).setValue( "t4" );

		ws.getRange2( "B551" ).setValue( "a" );
		ws.getRange2( "B552" ).setValue( "b" );
		ws.getRange2( "B553" ).setValue( "c" );
		ws.getRange2( "B554" ).setValue( "d" );
		ws.getRange2( "B555" ).setValue( "e" );
		ws.getRange2( "B556" ).setValue( "f" );
		ws.getRange2( "B557" ).setValue( "g" );
		ws.getRange2( "B558" ).setValue( "l" );
		ws.getRange2( "B559" ).setValue( "n" );
		ws.getRange2( "B560" ).setValue( "u" );
		ws.getRange2( "B561" ).setValue( "p" );

		ws.getRange2( "C551" ).setValue( "a1" );
		ws.getRange2( "C552" ).setValue( "b1" );
		ws.getRange2( "C553" ).setValue( "test2" );
		ws.getRange2( "C554" ).setValue( "c2" );
		ws.getRange2( "C555" ).setValue( "d2" );
		ws.getRange2( "C556" ).setValue( "h3" );
		ws.getRange2( "C557" ).setValue( "j5" );
		ws.getRange2( "C558" ).setValue( "u2" );
		ws.getRange2( "C559" ).setValue( "m1" );
		ws.getRange2( "C560" ).setValue( "k" );
		ws.getRange2( "C561" ).setValue( "l" );

		ws.getRange2( "C565" ).setValue( "99" );

		oParser = new parserFormula( "XLOOKUP(14,A551:A561,C551:C561)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate()), "u2" );

		oParser = new parserFormula( "XLOOKUP(C565,A551:A561,B551:C561)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 0), "n" );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 1), "m1" );

		oParser = new parserFormula( 'XLOOKUP(1,A551:A561,B551:C561,"not found")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate()), "not found" );

		oParser = new parserFormula( 'XLOOKUP(10,A551:A561,B551:C561,0,1,1)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 0), "f" );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 1), "h3" );

		oParser = new parserFormula( 'XLOOKUP(10,A551:A561,B551:C561,0,-1,1)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 0), "a" );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 1), "a1" );

		oParser = new parserFormula( 'XLOOKUP(10,A551:A561,B551:C561,0,0)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate()), 0 );

		oParser = new parserFormula( 'XLOOKUP(10,A551:A561,B551:C561)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate()), "#N/A" );

		oParser = new parserFormula( 'XLOOKUP(10,A551:A561,B551:C561,,,2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate()), "#N/A" );

		oParser = new parserFormula( 'XLOOKUP(10,A551:A561,B551:C561,,,-2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate()), "#N/A" );

		oParser = new parserFormula( 'XLOOKUP(10,A551:A561,B551:C561,,-1,-2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate()), "#N/A" );

		oParser = new parserFormula( 'XLOOKUP("test",A551:A561,B551:C561,,-1,2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 0), "n" );

		oParser = new parserFormula( 'XLOOKUP("tt",A551:A561,B551:C561,,-1,2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 0), "u" );

		oParser = new parserFormula( 'XLOOKUP("t???",A551:A561,B551:C561,,2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 0), "c" );

		oParser = new parserFormula( 'XLOOKUP("t???",A551:A561,B551:C561,,2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 0), "c" );

		oParser = new parserFormula( 'XLOOKUP("t?",A551:A561,B551:C561,,2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 0), "u" );

		oParser = new parserFormula( 'XLOOKUP("t?",A551:A561,B551:C561,,0)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate()), "#N/A" );

		oParser = new parserFormula( 'XLOOKUP("t*",A551:A561,B551:C561,,0)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate()), "#N/A" );

		oParser = new parserFormula( 'XLOOKUP("t*",A551:A561,B551:C561,,2)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 0), "c" );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 1), "test2" );

		oParser = new parserFormula( 'XLOOKUP(1,{1,2,3},{2,2,3})', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate(), 0, 0), 2 );

		oParser = new parserFormula( 'XLOOKUP(1,{1,2,3,4},{2,2,3})', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate()), "#VALUE!" );

		oParser = new parserFormula( 'XLOOKUP(4,{1,2,6},{2,2,3},,1)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate()), 3 );

		oParser = new parserFormula( 'XLOOKUP(4,{1,2,6},{2,2,3},,-1)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate()), 2 );

		oParser = new parserFormula( 'XLOOKUP(4,{1,2,6},{2,2,3})', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( _getValue(oParser.calculate()), "#N/A" );
	} );


	QUnit.test("Test: \"MATCH\"", function (assert) {

        ws.getRange2( "A551" ).setValue( "28" );
        ws.getRange2( "A552" ).setValue( "29" );
        ws.getRange2( "A553" ).setValue( "31" );
        ws.getRange2( "A554" ).setValue( "45" );
        ws.getRange2( "A555" ).setValue( "89" );

        ws.getRange2( "B551" ).setValue( "89" );
        ws.getRange2( "B552" ).setValue( "45" );
        ws.getRange2( "B553" ).setValue( "31" );
        ws.getRange2( "B554" ).setValue( "29" );
        ws.getRange2( "B555" ).setValue( "28" );

        ws.getRange2( "C551" ).setValue( "89" );
        ws.getRange2( "C552" ).setValue( "45" );
        ws.getRange2( "C553" ).setValue( "31" );
        ws.getRange2( "C554" ).setValue( "29" );
        ws.getRange2( "C555" ).setValue( "28" );

        oParser = new parserFormula( "MATCH(30,A551:A555,-1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

        oParser = new parserFormula( "MATCH(30,A551:A555,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2 );

        oParser = new parserFormula( "MATCH(30,A551:A555,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

        oParser = new parserFormula( "MATCH(30,B551:B555)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

        oParser = new parserFormula( "MATCH(30,B551:B555,-1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3 );

        oParser = new parserFormula( "MATCH(30,B551:B555,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

        oParser = new parserFormula( "MATCH(31,C551:C555,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 3 );

        oParser = new parserFormula( "MATCH(\"b\",{\"a\";\"b\";\"c\"},0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 2 );

		ws.getRange2( "F3" ).setValue( "" );

		ws.getRange2( "F106" ).setValue( "1" );
		ws.getRange2( "F107" ).setValue( "" );
		ws.getRange2( "F108" ).setValue( "" );
		ws.getRange2( "F109" ).setValue( "" );
		ws.getRange2( "F110" ).setValue( "2" );
		ws.getRange2( "F111" ).setValue( "123" );
		ws.getRange2( "F112" ).setValue( "4" );
		ws.getRange2( "F113" ).setValue( "5" );
		ws.getRange2( "F114" ).setValue( "6" );
		ws.getRange2( "F115" ).setValue( "0" );
		ws.getRange2( "F116" ).setValue( "" );
		ws.getRange2( "F117" ).setValue( "0" );

		oParser = new parserFormula( "MATCH(F3,F106:F114,0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

		oParser = new parserFormula( "MATCH(F3,F106:F117,0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 10 );

		oParser = new parserFormula( "MATCH(0,F106:F114,0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

		oParser = new parserFormula( "MATCH(0,F106:F117,0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 10 );

		oParser = new parserFormula( "MATCH(6,F106:F117,0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 9 );

		oParser = new parserFormula( "MATCH(6,F106:F117,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 5 );

		oParser = new parserFormula( "MATCH(6,F106:F117,-1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

		oParser = new parserFormula( "MATCH({6,2,3},F106:F117,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 5 );

		//TODO excel по-другому работает
		/*oParser = new parserFormula( "MATCH(123,F106:F117,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 6 );*/

    } );

    QUnit.test("Test: \"INDEX\"", function (assert) {

        ws.getRange2( "A651" ).setValue( "1" );
        ws.getRange2( "A652" ).setValue( "2" );
        ws.getRange2( "A653" ).setValue( "3" );
        ws.getRange2( "A654" ).setValue( "4" );
        ws.getRange2( "A655" ).setValue( "5" );

        ws.getRange2( "B651" ).setValue( "6" );
        ws.getRange2( "B652" ).setValue( "7" );
        ws.getRange2( "B653" ).setValue( "8" );
        ws.getRange2( "B654" ).setValue( "9" );
        ws.getRange2( "B655" ).setValue( "10" );

        ws.getRange2( "C651" ).setValue( "11" );
        ws.getRange2( "C652" ).setValue( "12" );
        ws.getRange2( "C653" ).setValue( "13" );
        ws.getRange2( "C654" ).setValue( "14" );
        ws.getRange2( "C655" ).setValue( "15" );

        oParser = new parserFormula( "INDEX({\"Apples\",\"Lemons\";\"Bananas\",\"Pears\"},2,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "Pears" );

        oParser = new parserFormula( "INDEX({\"Apples\",\"Lemons\";\"Bananas\",\"Pears\"},1,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "Lemons" );

        oParser = new parserFormula( "INDEX(\"Apples\",2,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        oParser = new parserFormula( "INDEX({\"Apples\",\"Lemons\"},,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "Lemons" );

		//данная функция возвращает area а далее уже в функции simplifyRefType находится резальтат
		// - пересечение а ячейкой, где располагается формула
        oParser = new parserFormula( "INDEX(A651:C655,,2)", "A2", ws );
        assert.ok( oParser.parse() );
        var parent =  AscCommonExcel.g_oRangeCache.getAscRange(oParser.parent);
        assert.strictEqual( oParser.simplifyRefType(oParser.calculate(), ws, parent.r1, parent.c1).getValue(), "#VALUE!" );

		oParser = new parserFormula( "INDEX(A651:C655,,2)", "D651", ws );
		assert.ok( oParser.parse() );
		parent =  AscCommonExcel.g_oRangeCache.getAscRange(oParser.parent);
		assert.strictEqual( oParser.simplifyRefType(oParser.calculate(), ws, parent.r1, parent.c1).getValue(), 6 );

		oParser = new parserFormula( "INDEX(A651:C655,,2)", "D652", ws );
		assert.ok( oParser.parse() );
		parent =  AscCommonExcel.g_oRangeCache.getAscRange(oParser.parent);
		assert.strictEqual( oParser.simplifyRefType(oParser.calculate(), ws, parent.r1, parent.c1).getValue(), 7 );

		oParser = new parserFormula( "INDEX(A651:C655,,3)", "E652", ws );
		assert.ok( oParser.parse() );
		parent =  AscCommonExcel.g_oRangeCache.getAscRange(oParser.parent);
		assert.strictEqual( oParser.simplifyRefType(oParser.calculate(), ws, parent.r1, parent.c1).getValue(), 12 );

		oParser = new parserFormula( "INDEX(A651:C655,,4)", "E652", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "INDEX(A651:C655,,14)", "E652", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );


        oParser = new parserFormula( "INDEX(A651:C655,3,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue().getValue(), 8 );

        oParser = new parserFormula( "INDEX(A651:C655,10,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "INDEX(A651:C651,1,3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), 11 );

		oParser = new parserFormula( "INDEX(A651:C651,1,2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), 6 );

		oParser = new parserFormula( "INDEX(A651:C651,0,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), 1 );

		oParser = new parserFormula( "INDEX(A651:C651,1,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), 1 );

		oParser = new parserFormula( "INDEX(A651:C651,2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), 6 );

		oParser = new parserFormula( "INDEX(A651:C651,3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), 11 );

		oParser = new parserFormula( "INDEX(A651:C651,4)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "INDEX(A651:C652,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "INDEX(A651:C652,2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "INDEX(A651:C652,0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "INDEX(A651:C651,1,1,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), 1 );

		oParser = new parserFormula( "INDEX(A651:C651,1,1,2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#REF!" );

		oParser = new parserFormula( "INDEX(A651:A655,A651,0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), 1 );
	} );

	QUnit.test("Test: \"INDIRECT\"", function (assert) {

		ws.getRange2( "A22" ).setValue( "B22" );
		ws.getRange2( "B22" ).setValue( "1.333" );

		ws.getRange2( "A23" ).setValue( "B23" );
		ws.getRange2( "B23" ).setValue( "45" );

		ws.getRange2( "A24" ).setValue( "George" );
		ws.getRange2( "B24" ).setValue( "10" );

		ws.getRange2( "A25" ).setValue( "25" );
		ws.getRange2( "B25" ).setValue( "62" );


		oParser = new parserFormula( "INDIRECT(A22)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), 1.333 );

		oParser = new parserFormula( "INDIRECT(A23)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), 45 );

		/*oParser = new parserFormula( "INDIRECT(A24)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), 10 );*/

		oParser = new parserFormula( 'INDIRECT("B"&A25)', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().getValue(), 62 );
	} );


	QUnit.test("Test: \"OFFSET\"", function (assert) {

        ws.getRange2( "C150" ).setValue( "1" );
        ws.getRange2( "D150" ).setValue( "2" );
        ws.getRange2( "E150" ).setValue( "3" );

        ws.getRange2( "C151" ).setValue( "2" );
        ws.getRange2( "D151" ).setValue( "3" );
        ws.getRange2( "E151" ).setValue( "4" );

        ws.getRange2( "C152" ).setValue( "3" );
        ws.getRange2( "D152" ).setValue( "4" );
        ws.getRange2( "E152" ).setValue( "5" );

        oParser = new parserFormula( "OFFSET(C3,2,3,1,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().toString(), "F5" );

        oParser = new parserFormula( "SUM(OFFSET(C151:E155,-1,0,3,3))", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 27 );

        oParser = new parserFormula( "OFFSET(B3, -2, 0, 1, 1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().toString(), "B1" );

		oParser = new parserFormula( "OFFSET(B3, 0, 0, -1, 1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "B3" );

		oParser = new parserFormula( "OFFSET(B3, 0, 0, -1, -1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "B3" );

		oParser = new parserFormula( "OFFSET(B3, 0, 0,,)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "B3" );

		oParser = new parserFormula( "OFFSET(B3, 0, 0, 1,)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "B3" );

		oParser = new parserFormula( "OFFSET(B3, 0, 0, -2, -2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "A2:B3" );

		oParser = new parserFormula( "OFFSET(B3, 0, 0, -1, -2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "A3:B3" );

		oParser = new parserFormula( "OFFSET(B3, 0, 0, 0, -2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "#REF!" );

		oParser = new parserFormula( "OFFSET(B3, 0, 0, 2, 0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "#REF!" );

		oParser = new parserFormula( "OFFSET(C3:D4, 0, 0, 2, 2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "C3:D4" );

		oParser = new parserFormula( "OFFSET(C3:D4, 0, 0, 3, 3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "C3:E5" );

		oParser = new parserFormula( "OFFSET(C3:D4, 2, 2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "E5:F6" );

		oParser = new parserFormula( "OFFSET(C3:D4,2,2,3,3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "E5:G7" );

		oParser = new parserFormula( "OFFSET(C3:E6, 0, 0, 3, 3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "C3:E5" );


		oParser = new parserFormula( "OFFSET(C3:D4, 0, 0, -2, -2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "B2:C3" );

		oParser = new parserFormula( "OFFSET(C3:D4, 0, 0, -3, -3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "A1:C3" );

		oParser = new parserFormula( "OFFSET(C3:E6, 0, 0, -3, -3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "A1:C3" );

		oParser = new parserFormula( "OFFSET(F10:M17, 0, 0, -7,-5)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "B4:F10" );
    } );

    /*
    * Financial
    */
    QUnit.test("Test: \"FV\"", function (assert) {

        function fv( rate, nper, pmt, pv, type ) {
            var res;
            if ( type === undefined || type === null )
                type = 0;

            if ( pv === undefined || pv === null )
                pv = 0;

            if ( rate != 0 ) {
                res = -1 * ( pv * Math.pow( 1 + rate, nper ) + pmt * ( 1 + rate * type ) * ( Math.pow( 1 + rate, nper ) - 1) / rate );
            }
            else {
                res = -1 * ( pv + pmt * nper );
            }
            return res;
        }

        oParser = new parserFormula( "FV(0.06/12,10,-200,-500,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), fv( 0.06 / 12, 10, -200, -500, 1 ) );

        oParser = new parserFormula( "FV(0.12/12,12,-1000)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), fv( 0.12 / 12, 12, -1000 ) );

        oParser = new parserFormula( "FV(0.11/12,35,-2000,,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( Math.abs( oParser.calculate().getValue() - fv( 0.11 / 12, 35, -2000, null, 1 ) ) < dif );

        oParser = new parserFormula( "FV(0.06/12,12,-100,-1000,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( Math.abs( oParser.calculate().getValue() - fv( 0.06 / 12, 12, -100, -1000, 1 ) ) < dif );

        testArrayFormula2(assert, "FV", 3, 5);
    } );

    QUnit.test("Test: \"PMT\"", function (assert) {

        function pmt( rate, nper, pv, fv, type ) {
            var res;
            if ( type === undefined || type === null )
                type = 0;

            if ( fv === undefined || fv === null )
                fv = 0;

            if ( rate != 0 ) {
                res = -1 * ( pv * Math.pow( 1 + rate, nper ) + fv ) /
                    ( ( 1 + rate * type ) * ( Math.pow( 1 + rate, nper ) - 1 ) / rate );
            }
            else {
                res = -1 * ( pv + fv ) / nper;
            }
            return res;
        }

        oParser = new parserFormula( "PMT(0.08/12,10,10000)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( Math.abs( oParser.calculate().getValue() - pmt( 0.08 / 12, 10, 10000 ) ) < dif );

        oParser = new parserFormula( "PMT(0.08/12,10,10000,0,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( Math.abs( oParser.calculate().getValue() - pmt( 0.08 / 12, 10, 10000, 0, 1 ) ) < dif );


		testArrayFormula2(assert, "PMT", 3, 5);
    } );

    QUnit.test("Test: \"NPER\"", function (assert) {

        function nper(rate,pmt,pv,fv,type){

            if ( rate === undefined || rate === null )
                rate = 0;

            if ( pmt === undefined || pmt === null )
                pmt = 0;

            if ( pv === undefined || pv === null )
                pv = 0;

            if ( type === undefined || type === null )
                type = 0;

            if ( fv === undefined || fv === null )
                fv = 0;

            var res;
            if ( rate != 0 ) {
                res = (-fv * rate + pmt * (1 + rate * type)) / (rate * pv + pmt * (1 + rate * type))
                res = Math.log( res ) / Math.log( 1+rate )
            }
            else {
                res = (- pv - fv )/ pmt ;
            }
            return res;
        }

        oParser = new parserFormula( "NPER(0.12/12,-100,-1000,10000,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), nper(0.12/12,-100,-1000,10000,1) );

        oParser = new parserFormula( "NPER(0.12/12,-100,-1000)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), nper(0.12/12,-100,-1000) );


		testArrayFormula2(assert, "NPER", 3, 5);
    } );

    QUnit.test("Test: \"PV\"", function (assert) {

        function pv( rate, nper, pmt, fv, type ) {
            if ( rate != 0 ) {
                return -1 * ( fv + pmt * (1 + rate * type) * ( (Math.pow( (1 + rate), nper ) - 1) / rate ) ) / Math.pow( 1 + rate, nper )
            }
            else {
                return -1 * ( fv + pmt * nper );
            }

        }

        oParser = new parserFormula( "PV(0.08/12,12*20,500,,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), pv( 0.08 / 12, 12 * 20, 500, 0, 0 ) );

        oParser = new parserFormula( "PV(0,12*20,500,,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), pv( 0, 12 * 20, 500, 0, 0 ) );

        testArrayFormula2(assert, "PV", 3, 5);
    } );

    QUnit.test("Test: \"NPV\"", function (assert) {

		//TODO в хроме при расчёте разница, временно убираю
		oParser = new parserFormula( "NPV(0.1,-10000,3000,4200,6800)", "A2", ws );
        assert.ok( oParser.parse() );
        //assert.strictEqual( oParser.calculate().getValue(), 1188.4434123352216 );

    } );

    QUnit.test("Test: \"EFFECT\"", function (assert) {

        function effect(nr,np){

            if( nr <= 0 || np < 1 ) return "#NUM!";

            return Math.pow( ( 1 + nr/np ), np ) - 1;

        }

        oParser = new parserFormula( "EFFECT(0.0525,4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), effect(0.0525,4) );

        oParser = new parserFormula( "EFFECT(0.0525,-4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), effect(0.0525,-4) );

        oParser = new parserFormula( "EFFECT(0.0525,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), effect(0.0525,1) );

        oParser = new parserFormula( "EFFECT(-1,54)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), effect(-1,54) );

        testArrayFormula2(assert, "EFFECT", 2, 2, true)
    } );

    QUnit.test("Test: \"ISPMT\"", function (assert) {

        function ISPMT( rate, per, nper, pv ){

            return pv * rate * (per / nper - 1.0)

        }

        oParser = new parserFormula( "ISPMT(0.1/12,1,3*12,8000000)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), ISPMT(0.1/12,1,3*12,8000000) );

        oParser = new parserFormula( "ISPMT(0.1,1,3,8000000)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), ISPMT(0.1,1,3,8000000) );

        testArrayFormula2(assert, "ISPMT", 4, 4);
    } );

	QUnit.test("Test: \"ISFORMULA\"", function (assert) {

		ws.getRange2( "C150" ).setValue( "=TODAY()" );
		ws.getRange2( "C151" ).setValue( "7" );
		ws.getRange2( "C152" ).setValue( "Hello, world!" );
		ws.getRange2( "C153" ).setValue( "=3/0" );

		oParser = new parserFormula( "ISFORMULA(C150)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "TRUE" );

		oParser = new parserFormula( "ISFORMULA(C151)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "FALSE" );

		oParser = new parserFormula( "ISFORMULA(C152)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "FALSE" );

		oParser = new parserFormula( "ISFORMULA(C153)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().toString(), "TRUE" );

		testArrayFormulaEqualsValues(assert, "FALSE,FALSE,FALSE,#N/A;FALSE,FALSE,FALSE,#N/A;#N/A,#N/A,#N/A,#N/A", "ISFORMULA(A1:C2)");
		testArrayFormulaEqualsValues(assert, "FALSE,FALSE,#N/A,#N/A;FALSE,FALSE,#N/A,#N/A;FALSE,FALSE,#N/A,#N/A", "ISFORMULA(A1:B1)");
		testArrayFormulaEqualsValues(assert, "FALSE,FALSE,FALSE,FALSE;FALSE,FALSE,FALSE,FALSE;FALSE,FALSE,FALSE,FALSE", "ISFORMULA(A1)");
	} );


	QUnit.test("Test: \"IFNA\"", function (assert) {

		oParser = new parserFormula( 'IFNA(MATCH(30,B1:B5,0),"Not found")', "A2", ws );
		assert.ok( oParser.parse(), 'IFNA(MATCH(30,B1:B5,0),"Not found")' );
		assert.strictEqual( oParser.calculate().getValue(), "Not found", 'IFNA(MATCH(30,B1:B5,0),"Not found")' );

	} );

	QUnit.test("Test: \"IFERROR\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "210" );
		ws.getRange2( "A3" ).setValue( "55" );
		ws.getRange2( "A4" ).setValue( "" );

		ws.getRange2( "B2" ).setValue( "35" );
		ws.getRange2( "B3" ).setValue( "0" );
		ws.getRange2( "B4" ).setValue( "23" );


		oParser = new parserFormula( 'IFERROR(A2/B2,"Error in calculation")', "A22", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 6 );

		oParser = new parserFormula( 'IFERROR(A3/B3,"Error in calculation")', "A22", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 'Error in calculation');

		oParser = new parserFormula( 'IFERROR(A4/B4,"Error in calculation")', "A22", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0);

		//testArrayFormula2(assert, "IFERROR", 2, 2);
	} );


	QUnit.test("Test: \"XNPV\"", function (assert) {

        function xnpv( rate, valueArray, dateArray ){
            var res = 0, r = rate;

            var d1 = dateArray[0];

            for( var i = 0; i < dateArray.length; i++ ){

                res += valueArray[i] / ( Math.pow( ( 1 + r ), ( dateArray[i] - d1 ) / 365 ) )
            }

            return res;
        }

        ws.getRange2( "A701" ).setValue( "39448" );
        ws.getRange2( "A702" ).setValue( "39508" );
        ws.getRange2( "A703" ).setValue( "39751" );
        ws.getRange2( "A704" ).setValue( "39859" );
        ws.getRange2( "A705" ).setValue( "39904" );

        oParser = new parserFormula( "XNPV(0.09,{-10000,2750,4250,3250,2750},A701:A705)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), xnpv( 0.09, [-10000,2750,4250,3250,2750], [39448,39508,39751,39859,39904] ) );

        ws.getRange2( "A705" ).setValue( "43191" );

        oParser = new parserFormula( "XNPV(0.09,{-10000,2750,4250,3250,2750},A701:A705)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), xnpv( 0.09, [-10000,2750,4250,3250,2750], [39448,39508,39751,39859,43191] ) );

    } );

    QUnit.test("Test: \"IRR\"", function (assert) {

        function irr( costArr, x ){

            if (!x) x = 0.1

            var nC = 0, g_Eps = 1e-7, fEps = 1.0, fZ = 0, fN = 0, xN = 0, nIM = 100, nMC = 0,arr0 = costArr[0], arrI, wasNegative = false, wasPositive = false;

            if( arr0 < 0 )
                wasNegative = true;
            else if( arr0 > 0 )
                wasPositive = true;

            while(fEps > g_Eps && nMC < nIM ){
                nC = 0; fZ = 0; fN = 0;
                fZ += costArr[0]/Math.pow( 1.0 + x, nC );
                fN += -nC * costArr[0]/Math.pow( 1 + x, nC + 1 );
                nC++;
                for(var i = 1; i < costArr.length; i++){
                    arrI = costArr[i];
                    fZ += arrI/Math.pow( 1.0 + x, nC );
                    fN += -nC * arrI/Math.pow( 1 + x, nC + 1 );
                    if( arrI < 0 )
                        wasNegative = true;
                    else if( arrI > 0 )
                        wasPositive = true
                    nC++
                }
                xN = x - fZ / fN;
                nMC ++;
                fEps = Math.abs( xN - x );
                x = xN;
            }


            if( !(wasNegative && wasPositive)  )
                return "#NUM!";

            if (fEps < g_Eps)
                return x;
            else
                return "#NUM!";

        }


		//TODO в хроме при расчёте разница, временно убираю
        oParser = new parserFormula( "IRR({-70000,12000,15000,18000,21000})", "A2", ws );
        assert.ok( oParser.parse() );
        //assert.strictEqual( oParser.calculate().getValue(), -0.021244848273410923 );

        ws.getRange2( "A705" ).setValue( "43191" );

		//TODO в хроме при расчёте разница, временно убираю
		oParser = new parserFormula( "IRR({-70000,12000,15000,18000,21000,26000})", "A2", ws );
        assert.ok( oParser.parse() );
        //assert.strictEqual( oParser.calculate().getValue(), 0.08663094803653171 );

        oParser = new parserFormula( "IRR({-70000,12000,15000},-0.1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), -0.44350694133450463 );

        oParser = new parserFormula( "IRR({-70000},-0.1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        //TODO пересмотреть тест для этой функции
        //testArrayFormula2(assert, "IRR", 1, 2, true)
    } );

    QUnit.test("Test: \"ACCRINT\"", function (assert) {

        oParser = new parserFormula( "ACCRINT(DATE(2006,3,1),DATE(2006,9,1),DATE(2006,5,1),0.1,1100,2,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 18.333333333333332 );

        oParser = new parserFormula( "ACCRINT(DATE(2006,3,1),DATE(2006,9,1),DATE(2006,5,1),0.1,,2,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 16.666666666666664 );

        oParser = new parserFormula( "ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2010,5,1),0.1,1000,2,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 216.94444444444444 );

        oParser = new parserFormula( "ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2010,5,1),0.1,1000,2,0,TRUE)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 216.94444444444444 );

        oParser = new parserFormula( "ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2010,5,1),0.1,1000,2,0,FALSE)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 216.66666666666666 );

        testArrayFormula2(assert, "ACCRINT", 6, 8, true);
    } );

    QUnit.test("Test: \"ACCRINTM\"", function (assert) {

        oParser = new parserFormula( "ACCRINTM(DATE(2006,3,1),DATE(2006,5,1),0.1,1100,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 18.333333333333332 );

        oParser = new parserFormula( "ACCRINTM(DATE(2006,3,1),DATE(2006,5,1),0.1,,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 16.666666666666664 );

        oParser = new parserFormula( "ACCRINTM(DATE(2006,3,1),DATE(2006,5,1),0.1,)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 16.666666666666664 );

        testArrayFormula2(assert, "ACCRINTM", 4, 5, true)

    } );

    QUnit.test("Test: \"AMORDEGRC\"", function (assert) {

        oParser = new parserFormula( "AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 776 );

        oParser = new parserFormula( "AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.50,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.20,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 819 );

        oParser = new parserFormula( "AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.33,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 972 );

        testArrayFormula2(assert, "AMORDEGRC", 6, 7, true);
    } );

    QUnit.test("Test: \"AMORLINC\"", function (assert) {

        oParser = new parserFormula( "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 360 );

        oParser = new parserFormula( "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.70,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1484 );

		testArrayFormula2(assert, "AMORLINC", 6, 7, true);
    } );

    QUnit.test("Test: \"CUMIPMT\"", function (assert) {

        function cumipmt(fRate, nNumPeriods, fVal, nStartPer, nEndPer, nPayType){

            var fRmz, fZinsZ;

            if( nStartPer < 1 || nEndPer < nStartPer || fRate <= 0.0 || nEndPer > nNumPeriods  || nNumPeriods <= 0 ||
                fVal <= 0.0 || ( nPayType != 0 && nPayType != 1 ) )
                return "#NUM!"

            fRmz = _getPMT( fRate, nNumPeriods, fVal, 0.0, nPayType );

            fZinsZ = 0.0;

            if( nStartPer == 1 )
            {
                if( nPayType <= 0 )
                    fZinsZ = -fVal;

                nStartPer++;
            }

            for( var i = nStartPer ; i <= nEndPer ; i++ )
            {
                if( nPayType > 0 )
                    fZinsZ += _getFV( fRate, i - 2, fRmz, fVal, 1 ) - fRmz;
                else
                    fZinsZ += _getFV( fRate, i - 1, fRmz, fVal, 0 );
            }

            fZinsZ *= fRate;

            return fZinsZ;

        }

        oParser = new parserFormula( "CUMIPMT(0.09/12,30*12,125000,1,1,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), cumipmt(0.09/12,30*12,125000,1,1,0) );

        oParser = new parserFormula( "CUMIPMT(0.09/12,30*12,125000,13,24,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), cumipmt(0.09/12,30*12,125000,13,24,0) );

        testArrayFormula2(assert, "CUMIPMT", 6, 6, true);
    } );

    QUnit.test("Test: \"CUMPRINC\"", function (assert) {

        function cumpring(fRate, nNumPeriods, fVal, nStartPer, nEndPer, nPayType){

            var fRmz, fKapZ;

            if( nStartPer < 1 || nEndPer < nStartPer || nEndPer < 1  || fRate <= 0 || nNumPeriods <= 0 || fVal <= 0 || ( nPayType != 0 && nPayType != 1 ) )
                return "#NUM!"

            fRmz = _getPMT( fRate, nNumPeriods, fVal, 0.0, nPayType );

            fKapZ = 0.0;

            var nStart = nStartPer;
            var nEnd = nEndPer;

            if( nStart == 1 )
            {
                if( nPayType <= 0 )
                    fKapZ = fRmz + fVal * fRate;
                else
                    fKapZ = fRmz;

                nStart++;
            }

            for( var i = nStart ; i <= nEnd ; i++ )
            {
                if( nPayType > 0 )
                    fKapZ += fRmz - ( _getFV( fRate, i - 2, fRmz, fVal, 1 ) - fRmz ) * fRate;
                else
                    fKapZ += fRmz - _getFV( fRate, i - 1, fRmz, fVal, 0 ) * fRate;
            }

            return fKapZ

        }

        oParser = new parserFormula( "CUMPRINC(0.09/12,30*12,125000,1,1,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), cumpring(0.09/12,30*12,125000,1,1,0) );

        oParser = new parserFormula( "CUMPRINC(0.09/12,30*12,-125000,1,1,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), cumpring(0.09/12,30*12,-125000,1,1,0) );

        oParser = new parserFormula( "CUMPRINC(0.09/12,30*12,125000,13,24,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), cumpring(0.09/12,30*12,125000,13,24,0) );

		testArrayFormula2(assert, "CUMPRINC", 6, 6, true);
    } );

    QUnit.test("Test: \"NOMINAL\"", function (assert) {

        function nominal(rate,np){

            if( rate <= 0 || np < 1 )
                return "#NUM!"

            return ( Math.pow( rate + 1, 1 / np ) - 1 ) * np;

        }

        oParser = new parserFormula( "NOMINAL(0.053543,4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), nominal(0.053543,4) );

        oParser = new parserFormula( "NOMINAL(0.053543,-4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), nominal(0.053543,-4) );

        testArrayFormula2(assert, "NOMINAL", 2, 2, true);
    } );

	QUnit.test( "Test: \"NOT\"", function (assert) {
		testArrayFormula2(assert, "NOT", 1, 1);
	} );

    QUnit.test("Test: \"FVSCHEDULE\"", function (assert) {

        function fvschedule(rate,shedList){

            for( var i = 0; i < shedList.length; i++){
                rate *= 1 + shedList[i]
            }

            return rate;

        }

        oParser = new parserFormula( "FVSCHEDULE(1,{0.09,0.11,0.1})", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), fvschedule(1,[0.09,0.11,0.1]) );

        //testArrayFormula2(assert, "FVSCHEDULE", 2, 2, true, true);
    } );

    QUnit.test("Test: \"DISC\"", function (assert) {

        function disc( settlement, maturity, pr, redemption, basis ){

            if( settlement >= maturity || pr <= 0 || redemption <= 0 || basis < 0 || basis > 4 )
                return "#NUM!"

            return ( 1.0 - pr / redemption ) / AscCommonExcel.yearFrac( settlement, maturity, basis );

        }

        oParser = new parserFormula( "DISC(DATE(2007,1,25),DATE(2007,6,15),97.975,100,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), disc( new cDate(2007,0,25),new cDate(2007,5,15),97.975,100,1 ) );

        testArrayFormula2(assert, "DISC",4,5,true);
    } );

    QUnit.test("Test: \"DOLLARDE\"", function (assert) {

        function dollarde( fractionalDollar, fraction ){

            if( fraction < 0 )
                return "#NUM!";
            else if( fraction == 0 )
                return "#DIV/0!";

            var fInt = Math.floor( fractionalDollar ), res  = fractionalDollar - fInt;

            res /= fraction;

            res *= Math.pow( 10, Math.ceil( Math.log( fraction ) / Math.log( 10 ) ) );

            res += fInt;

            return res;

        }

        oParser = new parserFormula( "DOLLARDE(1.02,16)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), dollarde( 1.02,16 ) );

        oParser = new parserFormula( "DOLLARDE(1.1,32)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), dollarde( 1.1,32 ) );

        testArrayFormula2(assert, "DOLLARDE", 2, 2, true);
    } );

    QUnit.test("Test: \"DOLLARFR\"", function (assert) {

        function dollarde( fractionalDollar, fraction ){

            if( fraction < 0 )
                return "#NUM!";
            else if( fraction == 0 )
                return "#DIV/0!";

            var fInt = Math.floor( fractionalDollar ), res  = fractionalDollar - fInt;

            res *= fraction;

            res *= Math.pow( 10.0, -Math.ceil( Math.log( fraction ) / Math.log( 10 ) ) );

            res += fInt;

            return res;

        }

        oParser = new parserFormula( "DOLLARFR(1.125,16)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), dollarde( 1.125,16 ) );

        oParser = new parserFormula( "DOLLARFR(1.125,32)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), dollarde( 1.125,32 ) );

		testArrayFormula2(assert, "DOLLARFR", 2, 2, true);
    } );

    QUnit.test("Test: \"RECEIVED\"", function (assert) {

        function received( settlement, maturity, investment, discount, basis ){

            if( settlement >= maturity || investment <= 0 || discount <= 0 || basis < 0 || basis > 4 )
                return "#NUM!"

            return investment / ( 1 - ( discount * AscCommonExcel.yearFrac( settlement, maturity, basis) ) )

        }

        oParser = new parserFormula( "RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), received( new cDate(2008,1,15),new cDate(2008,4,15),1000000,0.0575,2 ) );

        testArrayFormula2(assert, "RECEIVED", 4, 5, true);
    } );

    QUnit.test("Test: \"RATE\"", function (assert) {

        function RateIteration( fNper, fPayment, fPv, fFv, fPayType, fGuess ) {
            function approxEqual( a, b ) {
                if ( a == b )
                    return true;
                var x = a - b;
                return (x < 0.0 ? -x : x)
                    < ((a < 0.0 ? -a : a) * (1.0 / (16777216.0 * 16777216.0)));
            }

            var bValid = true, bFound = false, fX, fXnew, fTerm, fTermDerivation, fGeoSeries, fGeoSeriesDerivation;
            var nIterationsMax = 150, nCount = 0, fEpsilonSmall = 1.0E-14, SCdEpsilon = 1.0E-7;
            fFv = fFv - fPayment * fPayType;
            fPv = fPv + fPayment * fPayType;
            if ( fNper == Math.round( fNper ) ) {
                fX = fGuess.fGuess;
                var fPowN, fPowNminus1;
                while ( !bFound && nCount < nIterationsMax ) {
                    fPowNminus1 = Math.pow( 1.0 + fX, fNper - 1.0 );
                    fPowN = fPowNminus1 * (1.0 + fX);
                    if ( approxEqual( Math.abs( fX ), 0.0 ) ) {
                        fGeoSeries = fNper;
                        fGeoSeriesDerivation = fNper * (fNper - 1.0) / 2.0;
                    }
                    else {
                        fGeoSeries = (fPowN - 1.0) / fX;
                        fGeoSeriesDerivation = fNper * fPowNminus1 / fX - fGeoSeries / fX;
                    }
                    fTerm = fFv + fPv * fPowN + fPayment * fGeoSeries;
                    fTermDerivation = fPv * fNper * fPowNminus1 + fPayment * fGeoSeriesDerivation;
                    if ( Math.abs( fTerm ) < fEpsilonSmall )
                        bFound = true;
                    else {
                        if ( approxEqual( Math.abs( fTermDerivation ), 0.0 ) )
                            fXnew = fX + 1.1 * SCdEpsilon;
                        else
                            fXnew = fX - fTerm / fTermDerivation;
                        nCount++;
                        bFound = (Math.abs( fXnew - fX ) < SCdEpsilon);
                        fX = fXnew;
                    }
                }
                bValid =(fX >=-1.0);
            }
            else {
                fX = (fGuess.fGuest < -1.0) ? -1.0 : fGuess.fGuest;
                while ( bValid && !bFound && nCount < nIterationsMax ) {
                    if ( approxEqual( Math.abs( fX ), 0.0 ) ) {
                        fGeoSeries = fNper;
                        fGeoSeriesDerivation = fNper * (fNper - 1.0) / 2.0;
                    }
                    else {
                        fGeoSeries = (Math.pow( 1.0 + fX, fNper ) - 1.0) / fX;
                        fGeoSeriesDerivation = fNper * Math.pow( 1.0 + fX, fNper - 1.0 ) / fX - fGeoSeries / fX;
                    }
                    fTerm = fFv + fPv * pow( 1.0 + fX, fNper ) + fPayment * fGeoSeries;
                    fTermDerivation = fPv * fNper * Math.pow( 1.0 + fX, fNper - 1.0 ) + fPayment * fGeoSeriesDerivation;
                    if ( Math.abs( fTerm ) < fEpsilonSmall )
                        bFound = true;
                    else {
                        if ( approxEqual( Math.abs( fTermDerivation ), 0.0 ) )
                            fXnew = fX + 1.1 * SCdEpsilon;
                        else
                            fXnew = fX - fTerm / fTermDerivation;
                        nCount++;
                        bFound = (Math.abs( fXnew - fX ) < SCdEpsilon);
                        fX = fXnew;
                        bValid = (fX >= -1.0);
                    }
                }
            }
            fGuess.fGuess = fX;
            return bValid && bFound;
        }

        function rate(nper, pmt, pv, fv, type, quess){

            if ( fv === undefined ) fv = 0;
            if ( type === undefined ) type = 0;
            if ( quess === undefined ) quess = 0.1;

            var res = {fGuess:0};

            if( RateIteration(nper, pmt, pv, fv, type, res) )
                return res.fGuess;

            return "#VALUE!"
        }

        oParser = new parserFormula( "RATE(4*12,-200,8000)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), rate(4*12,-200,8000) ), true );

        oParser = new parserFormula( "RATE(4*12,-200,8000)*12", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( difBetween( oParser.calculate().getValue(), rate(4*12,-200,8000)*12 ), true );

        testArrayFormula2(assert, "RATE", 3, 6, true);
    } );

	QUnit.test("Test: \"RRI\"", function (assert) {

		oParser = new parserFormula( "RRI(96, 10000, 11000)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0009933 );

		oParser = new parserFormula( "RRI(0, 10000, 11000)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		oParser = new parserFormula( "RRI(-10, 10000, 11000)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		oParser = new parserFormula( "RRI(10, 10000, -11000)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		oParser = new parserFormula( "RRI(1, 1, -1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -2 );

		testArrayFormula2(assert, "RRI", 3, 3);
	} );

    QUnit.test("Test: \"INTRATE\"", function (assert) {

        function intrate( settlement, maturity, investment, redemption, basis ){

            if( settlement >= maturity || investment <= 0 || redemption <= 0 || basis < 0 || basis > 4 )
                return "#NUM!"

            return ( ( redemption / investment ) - 1 ) / AscCommonExcel.yearFrac( settlement, maturity, basis )

        }

        oParser = new parserFormula( "INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), intrate( new cDate(2008,1,15),new cDate(2008,4,15),1000000,1014420,2 ) );

        testArrayFormula2(assert, "INTRATE", 4, 5, true);
    } );

    QUnit.test("Test: \"TBILLEQ\"", function (assert) {

        function tbilleq( settlement, maturity, discount ){

            maturity = cDate.prototype.getDateFromExcel(maturity.getExcelDate() + 1);

            var d1 = settlement, d2 = maturity;
            var date1 = d1.getDate(), month1 = d1.getMonth(), year1 = d1.getFullYear(),
                date2 = d2.getDate(), month2 = d2.getMonth(), year2 = d2.getFullYear();

            var nDiff = GetDiffDate360( date1, month1, year1, date2, month2, year2, true );

            if( settlement >= maturity || discount <= 0 || nDiff > 360 )
                return "#NUM!";

            return ( 365 * discount ) / ( 360 - discount * nDiff );

        }

        oParser = new parserFormula( "TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), tbilleq( new cDate(Date.UTC(2008,2,31)), new cDate(Date.UTC(2008,5,1)), 0.0914 ) );

        testArrayFormula2(assert, "TBILLEQ", 3, 3, true);
    } );

    QUnit.test("Test: \"TBILLPRICE\"", function (assert) {

        function tbillprice( settlement, maturity, discount ){

            maturity = cDate.prototype.getDateFromExcel(maturity.getExcelDate() + 1)

            var d1 = settlement
            var d2 = maturity

            var fFraction = AscCommonExcel.yearFrac(d1, d2, 0);

            if( fFraction - Math.floor( fFraction ) == 0 )
                return "#NUM!";

            return 100 * ( 1 - discount * fFraction );

        }

        oParser = new parserFormula( "TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), tbillprice( new cDate(Date.UTC(2008,2,31)), new cDate(Date.UTC(2008,5,1)), 0.09 ) );

        testArrayFormula2(assert, "TBILLPRICE", 3, 3, true);
    } );

    QUnit.test("Test: \"TBILLYIELD\"", function (assert) {

        function tbillyield( settlement, maturity, pr ){

            var d1 = settlement;
            var d2 = maturity;
            var date1 = d1.getDate(), month1 = d1.getMonth(), year1 = d1.getFullYear(),
                date2 = d2.getDate(), month2 = d2.getMonth(), year2 = d2.getFullYear();

            var nDiff = GetDiffDate360( date1, month1, year1, date2, month2, year2, true );
            nDiff++;
            if( settlement >= maturity || pr <= 0 || nDiff > 360 )
                return "#NUM!";

            return ( ( 100 - pr ) / pr) * (360 / nDiff);

        }

        oParser = new parserFormula( "TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), tbillyield( new cDate(2008,2,31), new cDate(2008,5,1), 98.45 ) );

    } );

    QUnit.test("Test: \"COUPDAYBS\"", function (assert) {

        function coupdaybs( settlement, maturity, frequency, basis ){

            basis = ( basis !== undefined ? basis : 0 );

            return _getcoupdaybs(settlement, maturity, frequency, basis)

        }

        oParser = new parserFormula( "COUPDAYBS(DATE(2007,1,25),DATE(2008,11,15),2,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 71 );

        oParser = new parserFormula( "COUPDAYBS(DATE(2007,1,25),DATE(2008,11,15),2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), coupdaybs( new cDate(2007,0,25), new cDate(2008,10,15), 2 ) );

        testArrayFormula2(assert, "COUPDAYBS", 3, 4, true);
    } );

    QUnit.test("Test: \"COUPDAYS\"", function (assert) {

        function coupdays( settlement, maturity, frequency, basis ){

            basis = ( basis !== undefined ? basis : 0 );

            return _getcoupdays(settlement, maturity, frequency, basis)

        }

        oParser = new parserFormula( "COUPDAYS(DATE(2007,1,25),DATE(2008,11,15),2,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), coupdays( new cDate(2007,0,25), new cDate(2008,10,15), 2, 1 ) );

        oParser = new parserFormula( "COUPDAYS(DATE(2007,1,25),DATE(2008,11,15),2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), coupdays( new cDate(2007,0,25), new cDate(2008,10,15), 2 ) );

		testArrayFormula2(assert, "COUPDAYS", 3, 4, true);
    } );

    QUnit.test("Test: \"COUPDAYSNC\"", function (assert) {

        function coupdaysnc( settlement, maturity, frequency, basis ) {

            basis = ( basis !== undefined ? basis : 0 );

            if ( (basis != 0) && (basis != 4) ) {

                _lcl_GetCoupncd( settlement, maturity, frequency );
                return _diffDate( settlement, maturity, basis );
            }

            return _getcoupdays( new cDate( settlement ), new cDate( maturity ), frequency, basis ) - _getcoupdaybs( new cDate( settlement ), new cDate( maturity ), frequency, basis );

        }

        oParser = new parserFormula( "COUPDAYSNC(DATE(2007,1,25),DATE(2008,11,15),2,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 110 );

        oParser = new parserFormula( "COUPDAYSNC(DATE(2007,1,25),DATE(2008,11,15),2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), coupdaysnc( new cDate(2007,0,25), new cDate(2008,10,15), 2 ) );

		testArrayFormula2(assert, "COUPDAYSNC", 3, 4, true);
    } );

    QUnit.test("Test: \"COUPNCD\"", function (assert) {

        function coupncd( settlement, maturity, frequency, basis ) {

            basis = ( basis !== undefined ? basis : 0 );

            _lcl_GetCoupncd( settlement, maturity, frequency );

            return maturity.getExcelDate();

        }

        oParser = new parserFormula( "COUPNCD(DATE(2007,1,25),DATE(2008,11,15),2,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), coupncd( new cDate(Date.UTC(2007,0,25)), new cDate(Date.UTC(2008,10,15)), 2, 1 ) );

		testArrayFormula2(assert, "COUPNCD", 3, 4, true);
    } );

    QUnit.test("Test: \"COUPNUM\"", function (assert) {

        oParser = new parserFormula( "COUPNUM(DATE(2007,1,25),DATE(2008,11,15),2,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), _coupnum( new cDate(2007,0,25), new cDate(2008,10,15), 2, 1 ) );

		testArrayFormula2(assert, "COUPNUM", 3, 4, true);
    } );

    QUnit.test("Test: \"COUPPCD\"", function (assert) {

        function couppcd( settlement, maturity, frequency, basis ) {

            basis = ( basis !== undefined ? basis : 0 );

            _lcl_GetCouppcd( settlement, maturity, frequency );
            return maturity.getExcelDate();

        }

        oParser = new parserFormula( "COUPPCD(DATE(2007,1,25),DATE(2008,11,15),2,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), couppcd( new cDate(Date.UTC(2007,0,25)), new cDate(Date.UTC(2008,10,15)), 2, 1 ) );

		testArrayFormula2(assert, "COUPPCD", 3, 4, true);
    } );

	QUnit.test("Test: \"CONVERT\"", function (assert) {

		oParser = new parserFormula( 'CONVERT(68, "F", "C")', "A2", ws );
		assert.ok( oParser.parse(), 'CONVERT(68, "F", "C")' );
		assert.strictEqual( oParser.calculate().getValue(), 20, 'CONVERT(68, "F", "C")' );

		oParser = new parserFormula( 'CONVERT(2.5, "ft", "sec")', "A2", ws );
		assert.ok( oParser.parse(), 'CONVERT(2.5, "ft", "sec")' );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A", 'CONVERT(2.5, "ft", "sec")' );

		oParser = new parserFormula( 'CONVERT(CONVERT(100,"ft","m"),"ft","m")', "A2", ws );
		assert.ok( oParser.parse(), 'CONVERT(CONVERT(100,"ft","m"),"ft","m")' );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 9.290304, 'CONVERT(CONVERT(100,"ft","m"),"ft","m")' );

		oParser = new parserFormula( 'CONVERT(7,"bit","byte")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(3) - 0, 0.875 );

		oParser = new parserFormula( 'CONVERT(7,"admkn","kn")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(14) - 0, 6.99999939524838 );

		oParser = new parserFormula( 'CONVERT(7,"admkn","m/s")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 3.6011108 );

		oParser = new parserFormula( 'CONVERT(7,"admkn","mph")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 8.0554554 );

		oParser = new parserFormula( 'CONVERT(7,"m/h","m/sec")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0019444 );

		oParser = new parserFormula( 'CONVERT(7,"m/hr","mph")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0043496 );

		oParser = new parserFormula( 'CONVERT(7,"m","mi")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0043496 );

		oParser = new parserFormula( 'CONVERT(7,"m","Pica")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 19842.5196850 );

		oParser = new parserFormula( 'CONVERT(7,"m","pica")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 1653.5433071 );

		oParser = new parserFormula( 'CONVERT(7,"Nmi","pica")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 3062362.2047251 );

		oParser = new parserFormula( 'CONVERT(7,"yr","day")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 2556.75 );

		oParser = new parserFormula( 'CONVERT(7,"yr","min")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3681720 );

		oParser = new parserFormula( 'CONVERT(7,"day","min")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 10080 );

		oParser = new parserFormula( 'CONVERT(7,"hr","sec")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 25200 );

		oParser = new parserFormula( 'CONVERT(7,"min","sec")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 420 );

		oParser = new parserFormula( 'CONVERT(7,"Pa","mmHg")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0525043 );

		oParser = new parserFormula( 'CONVERT(7,"Pa","psi")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0010153 );

		oParser = new parserFormula( 'CONVERT(7,"Pa","Torr")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0525045 );

		oParser = new parserFormula( 'CONVERT(7,"g","sg")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0004797 );

		oParser = new parserFormula( 'CONVERT(7,"g","lbm")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.0154324 );

		oParser = new parserFormula( 'CONVERT(1, "lbm", "kg")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(7) - 0, 0.4535924 );

		oParser = new parserFormula( 'CONVERT(1, "lbm", "mg")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(0) - 0, 453592 );

		oParser = new parserFormula( 'CONVERT(1, "klbm", "mg")', "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A" );

		testArrayFormula2(assert, "CONVERT", 3, 3, true);
	} );

    QUnit.test("Test: \"PRICE\"", function (assert) {

        oParser = new parserFormula( "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), _getprice( new cDate( Date.UTC(2008, 1, 15 )), new cDate( Date.UTC(2017, 10, 15 )), 0.0575, 0.065, 100, 2, 0 ) );

        testArrayFormula2(assert, "PRICE", 6, 7, true);
    } );

    QUnit.test("Test: \"PRICEDISC\"", function (assert) {

        function pricedisc(settl, matur, discount, redemption, basis){
            return redemption * ( 1.0 - discount * _getdiffdate( settl, matur, basis ) );
        }

        oParser = new parserFormula( "PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), pricedisc( new cDate(2008,1,16), new cDate(2008,2,1),0.0525,100,2 ) );

		testArrayFormula2(assert, "PMT", 4, 5, true);
    } );

    QUnit.test("Test: \"PRICEMAT\"", function (assert) {

        function pricemat( settl, matur, iss, rate, yld, basis ) {

            var fIssMat = _yearFrac( new cDate(iss), new cDate(matur), basis );
            var fIssSet = _yearFrac( new cDate(iss), new cDate(settl), basis );
            var fSetMat = _yearFrac( new cDate(settl), new cDate(matur), basis );

            var res = 1.0 + fIssMat * rate;
            res /= 1.0 + fSetMat * yld;
            res -= fIssSet * rate;
            res *= 100.0;

            return res;
        }

        oParser = new parserFormula( "PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), pricemat( new cDate(2008,1,15),new cDate(2008,3,13),new cDate(2007,10,11),0.061,0.061,0 ) );

        testArrayFormula2(assert, "PRICEMAT", 5, 6, true);
    } );

    QUnit.test("Test: \"YIELD\"", function (assert) {

        oParser = new parserFormula( "YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), _getYield( new cDate(Date.UTC(2008,1,15)), new cDate(Date.UTC(2016,10,15)),0.0575,95.04287,100,2,0 ) );

        testArrayFormula2(assert, "YIELD", 6, 7, true);
    } );

    QUnit.test("Test: \"YIELDDISC\"", function (assert) {

        function yielddisc( settlement, maturity, pr, redemption, basis ){

            var fRet = ( redemption / pr ) - 1.0;
            fRet /= _yearFrac( settlement, maturity, basis );
            return fRet;

        }

        oParser = new parserFormula( "YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), yielddisc( new cDate( 2008, 1, 16 ), new cDate( 2008, 2, 1 ), 99.795, 100, 2 ) );

		testArrayFormula2(assert, "YIELDDISC", 4, 5, true);
    } );

    QUnit.test("Test: \"YIELDMAT\"", function (assert) {

        oParser = new parserFormula( "YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), _getyieldmat( new cDate( 2008, 2, 15 ), new cDate( 2008, 10, 3 ), new cDate( 2007, 10, 8 ), 0.0625, 100.0123, 0 ) );

        testArrayFormula2(assert, "YIELDMAT", 5, 6, true);
    } );

	QUnit.test("Test: \"ODD\"", function (assert) {

		oParser = new parserFormula( "ODD(1.5)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "ODD(3)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "ODD(2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		oParser = new parserFormula( "ODD(-1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -1 );

		oParser = new parserFormula( "ODD(-2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -3 );

		testArrayFormula(assert, "ODD");

	} );

    QUnit.test("Test: \"ODDLPRICE\"", function (assert) {

        function oddlprice( settlement, maturity, last_interest, rate, yld, redemption, frequency, basis ){

            var fDCi = _yearFrac( last_interest, maturity, basis ) * frequency;
            var fDSCi = _yearFrac( settlement, maturity, basis ) * frequency;
            var fAi = _yearFrac( last_interest, settlement, basis ) * frequency;

            var res = redemption + fDCi * 100.0 * rate / frequency;
            res /= fDSCi * yld / frequency + 1.0;
            res -= fAi * 100.0 * rate / frequency;

            return res;
        }

        oParser = new parserFormula( "ODDLPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),0.0785,0.0625,100,2,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), oddlprice( new cDate(Date.UTC(2008,10,11)), new cDate(Date.UTC(2021,2,1)), new cDate(Date.UTC(2008,9,15)), 0.0785, 0.0625, 100, 2, 1 ) );

        testArrayFormula2(assert, "ODDLPRICE", 7, 8, true);
    } );

    QUnit.test("Test: \"ODDLYIELD\"", function (assert) {

        function oddlyield( settlement, maturity, last_interest, rate, pr, redemption, frequency, basis ){

            var fDCi = _yearFrac( last_interest, maturity, basis ) * frequency;
            var fDSCi = _yearFrac( settlement, maturity, basis ) * frequency;
            var fAi = _yearFrac( last_interest, settlement, basis ) * frequency;

            var res = redemption + fDCi * 100.0 * rate / frequency;
            res /= pr + fAi * 100.0 * rate / frequency;
            res--;
            res *= frequency / fDSCi;

            return res;
        }

        oParser = new parserFormula( "ODDLYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),0.0575,84.5,100,2,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), oddlyield( new cDate(2008,10,11), new cDate(2021,2,1), new cDate(2008,9,15), 0.0575, 84.5, 100, 2, 0 ) );

		testArrayFormula2(assert, "ODDLYIELD", 7, 8, true);
    } );

    QUnit.test("Test: \"DURATION\"", function (assert) {

        oParser = new parserFormula( "DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), _duration( new cDate(Date.UTC(2008,0,1)), new cDate(Date.UTC(2016,0,1)), 0.08, 0.09, 2, 1 ) );

        oParser = new parserFormula( "DURATION(DATE(2008,1,1),DATE(2016,1,1),-0.08,0.09,2,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), _duration( new cDate(Date.UTC(2008,0,1)), new cDate(Date.UTC(2016,0,1)), -0.08, 0.09, 2, 1 ) );

        oParser = new parserFormula( "DURATION(DATE(2008,1,1),DATE(2016,1,1),-0.08,0.09,5,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), _duration( new cDate(Date.UTC(2008,0,1)), new cDate(Date.UTC(2016,0,1)), -0.08, 0.09, 5, 1 ) );

        testArrayFormula2(assert, "DURATION", 5, 6, true);
    } );

    QUnit.test("Test: \"MDURATION\"", function (assert) {

        function mduration(settl, matur, coupon, yld, frequency, basis){

            return _duration( settl, matur, coupon, yld, frequency, basis ) / (1 + yld/frequency);

        }

        oParser = new parserFormula( "MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), mduration( new cDate(Date.UTC(2008,0,1)), new cDate(Date.UTC(2016,0,1)), 0.08, 0.09, 2, 1 ) );

        testArrayFormula2(assert, "MDURATION", 5, 6, true);
    } );

	QUnit.test("Test: \"MDETERM\"", function (assert) {

		ws.getRange2( "A2" ).setValue( "1" );
		ws.getRange2( "A3" ).setValue( "1" );
		ws.getRange2( "A4" ).setValue( "1" );
		ws.getRange2( "A5" ).setValue( "7" );

		ws.getRange2( "B2" ).setValue( "3" );
		ws.getRange2( "B3" ).setValue( "3" );
		ws.getRange2( "B4" ).setValue( "1" );
		ws.getRange2( "B5" ).setValue( "3" );

		ws.getRange2( "C2" ).setValue( "8" );
		ws.getRange2( "C3" ).setValue( "6" );
		ws.getRange2( "C4" ).setValue( "1" );
		ws.getRange2( "C5" ).setValue( "10" );

		ws.getRange2( "D2" ).setValue( "5" );
		ws.getRange2( "D3" ).setValue( "1" );
		ws.getRange2( "D4" ).setValue( "0" );
		ws.getRange2( "D5" ).setValue( "2" );


		oParser = new parserFormula( "MDETERM(A2:D5)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 88 );

		oParser = new parserFormula( "MDETERM({3,6,1;1,1,0;3,10,2})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( "MDETERM({3,6;1,1})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), -3 );

		oParser = new parserFormula( "MDETERM({1,3,8,5;1,3,6,1})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

	} );

    QUnit.test("Test: \"SYD\"", function (assert) {

        function syd( cost, salvage, life, per ){

            if( life == -1 || life == 0 )
                return "#NUM!";

            var res = 2;
            res *= cost - salvage;
            res *= life+1-per;
            res /= (life+1)*life;

            return res < 0 ? "#NUM!" : res;
        }

        oParser = new parserFormula( "SYD(30000,7500,10,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), syd( 30000,7500,10,1 ) );

        oParser = new parserFormula( "SYD(30000,7500,-1,10)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), syd( 30000,7500,-1,10 ) );

        oParser = new parserFormula( "SYD(30000,7500,-10,10)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), syd( 30000,7500,-10,10 ) );

        testArrayFormula2(assert, "SYD", 4, 4);
    } );

    QUnit.test("Test: \"PPMT\"", function (assert) {

        function ppmt( rate, per, nper, pv, fv, type ){

            if( fv == undefined ) fv = 0;
            if( type == undefined ) type = 0;

            var fRmz = _getPMT(rate, nper, pv, fv, type);

            return fRmz - _getIPMT(rate, per, pv, type, fRmz);

        }

        oParser = new parserFormula( "PPMT(0.1/12,1,2*12,2000)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), ppmt( 0.1/12,1,2*12,2000 ) );

        oParser = new parserFormula( "PPMT(0.08,10,10,200000)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), ppmt( 0.08,10,10,200000 ) );

        testArrayFormula2(assert, "PPMT", 4, 6);
    } );

    QUnit.test("Test: \"MIRR\"", function (assert) {

        function mirr( valueArray, fRate1_invest, fRate1_reinvest ){

            fRate1_invest = fRate1_invest + 1;
            fRate1_reinvest = fRate1_reinvest + 1;

            var fNPV_reinvest = 0, fPow_reinvest = 1, fNPV_invest = 0, fPow_invest = 1, fCellValue,
                wasNegative = false, wasPositive = false;

            for(var i = 0; i < valueArray.length; i++){
                fCellValue = valueArray[i];

                if( fCellValue > 0 ){
                    wasPositive = true;
                    fNPV_reinvest += fCellValue * fPow_reinvest;
                }
                else if( fCellValue < 0 ){
                    wasNegative = true;
                    fNPV_invest += fCellValue * fPow_invest;
                }
                fPow_reinvest /= fRate1_reinvest;
                fPow_invest /= fRate1_invest;

            }

            if( !( wasNegative && wasPositive ) )
                return "#DIV/0!";

            var fResult = -fNPV_reinvest / fNPV_invest;
            fResult *= Math.pow( fRate1_reinvest, valueArray.length - 1 );
            fResult = Math.pow( fResult, 1 / (valueArray.length - 1) );

            return fResult - 1;

        }

        oParser = new parserFormula( "MIRR({-120000,39000,30000,21000,37000,46000},0.1,0.12)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), mirr( [-120000,39000,30000,21000,37000,46000],0.1,0.12 ) );

        oParser = new parserFormula( "MIRR({-120000,39000,30000,21000},0.1,0.12)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), mirr( [-120000,39000,30000,21000],0.1,0.12 ) );

        oParser = new parserFormula( "MIRR({-120000,39000,30000,21000,37000,46000},0.1,0.14)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), mirr( [-120000,39000,30000,21000,37000,46000],0.1,0.14 ) );

        //testArrayFormula2(assert, "MIRR", 3, 3, null, true);
    } );

    QUnit.test("Test: \"IPMT\"", function (assert) {

        function ipmt( rate, per, nper, pv, fv, type ){

            if( fv == undefined ) fv = 0;
            if( type == undefined ) type = 0;

            var res = AscCommonExcel.getPMT(rate, nper, pv, fv, type);
            res = AscCommonExcel.getIPMT(rate, per, pv, type, res);

            return res;

        }

        oParser = new parserFormula( "IPMT(0.1/12,1*3,3,8000)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), ipmt( 0.1/12,1*3,3,8000 ) );

        oParser = new parserFormula( "IPMT(0.1,3,3,8000)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), ipmt( 0.1,3,3,8000 ) );

		testArrayFormula2(assert, "IPMT", 4, 6);
    } );

    QUnit.test("Test: \"DB\"", function (assert) {

        function db( cost, salvage, life, period, month ){

            if ( salvage >= cost ) {
                return this.value = new AscCommonExcel.cNumber( 0 );
            }

            if ( month < 1 || month > 12 || salvage < 0 || life <= 0 || period < 0 || life + 1 < period || cost < 0 ) {
                return "#NUM!";
            }

            var nAbRate = 1 - Math.pow( salvage / cost, 1 / life );
            nAbRate = Math.floor( (nAbRate * 1000) + 0.5 ) / 1000;
            var nErsteAbRate = cost * nAbRate * month / 12;

            var res = 0;
            if ( Math.floor( period ) == 1 )
                res = nErsteAbRate;
            else {
                var nSummAbRate = nErsteAbRate, nMin = life;
                if ( nMin > period ) nMin = period;
                var iMax = Math.floor( nMin );
                for ( var i = 2; i <= iMax; i++ ) {
                    res = (cost - nSummAbRate) * nAbRate;
                    nSummAbRate += res;
                }
                if ( period > life )
                    res = ((cost - nSummAbRate) * nAbRate * (12 - month)) / 12;
            }

            return res

        }

        oParser = new parserFormula( "DB(1000000,100000,6,1,7)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), db(1000000,100000,6,1,7) );

        oParser = new parserFormula( "DB(1000000,100000,6,2,7)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), db(1000000,100000,6,2,7) );

        oParser = new parserFormula( "DB(1000000,100000,6,3,7)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), db(1000000,100000,6,3,7) );

        oParser = new parserFormula( "DB(1000000,100000,6,4,7)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), db(1000000,100000,6,4,7) );

        oParser = new parserFormula( "DB(1000000,100000,6,5,7)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), db(1000000,100000,6,5,7) );

        oParser = new parserFormula( "DB(1000000,100000,6,6,7)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), db(1000000,100000,6,6,7) );

        oParser = new parserFormula( "DB(1000000,100000,6,7,7)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), db(1000000,100000,6,7,7) );

        testArrayFormula2(assert, "DB",4,5);

    } );

    QUnit.test("Test: \"DDB\"", function (assert) {

        function ddb( cost, salvage, life, period, factor ){

            if( factor === undefined || factor === null ) factor = 2;
            return _getDDB(cost, salvage, life, period, factor);
        }

        oParser = new parserFormula( "DDB(2400,300,10*365,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), ddb(2400,300,10*365,1) );

        oParser = new parserFormula( "DDB(2400,300,10*12,1,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), ddb(2400,300,10*12,1,2) );

        oParser = new parserFormula( "DDB(2400,300,10,1,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), ddb(2400,300,10,1,2) );

        oParser = new parserFormula( "DDB(2400,300,10,2,1.5)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), ddb(2400,300,10,2,1.5) );

        oParser = new parserFormula( "DDB(2400,300,10,10)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), ddb(2400,300,10,10) );


        //TODO format $
		ws.getRange2( "A102" ).setValue( "2400" );
		ws.getRange2( "A103" ).setValue( "300" );
		ws.getRange2( "A104" ).setValue( "10" );

		oParser = new parserFormula( "DDB(A102,A103,A104*365,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(2) - 0, 1.32 );

		oParser = new parserFormula( "DDB(A102,A103,A104*12,1,2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 40 );

		oParser = new parserFormula( "DDB(A102,A103,A104,1,2)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 480 );

		oParser = new parserFormula( "DDB(A102,A103,A104,2,1.5)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(0) - 0, 306 );

		oParser = new parserFormula( "DDB(A102,A103,A104,10)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(2) - 0, 22.12 );

		oParser = new parserFormula( "DDB(A102,A103,0,10)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		testArrayFormula2(assert, "DDB",4,5);

    } );

    QUnit.test("Test: \"SLN\"", function (assert) {

        function sln( cost, salvage, life ){

            if ( life == 0 ) return "#NUM!";

            return ( cost - salvage ) / life;
        }

        oParser = new parserFormula( "SLN(30000,7500,10)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), sln(30000,7500,10) );

		testArrayFormula2(assert, "SLN", 3, 3);
    } );

    QUnit.test("Test: \"XIRR\"", function (assert) {
        function lcl_sca_XirrResult( rValues, rDates, fRate ) {
            var D_0 = rDates[0];
            var r = fRate + 1;
            var fResult = rValues[0];
            for ( var i = 1, nCount = rValues.length; i < nCount; ++i )
                fResult += rValues[i] / Math.pow( r, (rDates[i] - D_0) / 365 );
            return fResult;
        }

        function lcl_sca_XirrResult_Deriv1( rValues, rDates, fRate ) {
            var D_0 = rDates[0];
            var r = fRate + 1;
            var fResult = 0;
            for ( var i = 1, nCount = rValues.length; i < nCount; ++i ) {
                var E_i = (rDates[i] - D_0) / 365;
                fResult -= E_i * rValues[i] / Math.pow( r, E_i + 1 );
            }
            return fResult;
        }

        function xirr( valueArray, dateArray, rate ) {

            var res = rate
            if ( res <= -1 )
                return "#NUM!"

            var fMaxEps = 1e-6, maxIter = 100;

            var newRate, eps, xirrRes, bContLoop;
            do
             {
                 xirrRes = lcl_sca_XirrResult( valueArray, dateArray, res );
                 newRate = res - xirrRes / lcl_sca_XirrResult_Deriv1( valueArray, dateArray, res );
                 eps = Math.abs( newRate - res );
                 res = newRate;
                 bContLoop = (eps > fMaxEps) && (Math.abs( xirrRes ) > fMaxEps);
             }
             while ( --maxIter && bContLoop );

            if ( bContLoop )
                return "#NUM!";

            return res;

        }

        ws.getRange2( "F100" ).setValue( "1/1/2008" );
        ws.getRange2( "G100" ).setValue( "3/1/2008" );
        ws.getRange2( "H100" ).setValue( "10/30/2008" );
        ws.getRange2( "I100" ).setValue( "2/15/2009" );
        ws.getRange2( "J100" ).setValue( "4/1/2009" );

        oParser = new parserFormula( "XIRR({-10000,2750,4250,3250,2750},F100:J100,0.1)", "A2", ws );
        assert.ok( oParser.parse() );
		assert.ok( difBetween( oParser.calculate().getValue(), 0.3733625335188316 ) );

        ws.getRange2( "F100" ).setValue( 0 );
        assert.ok( oParser.parse() );
		assert.ok( difBetween( oParser.calculate().getValue(), 0.0024114950175866895 ) );

    } );

	QUnit.test("Test: \"VDB\"", function (assert) {


		function _getVDB(cost, salvage, life, life1, startperiod, factor) {
			var fVdb = 0, nLoopEnd = end = Math.ceil(startperiod), fTerm, fLia = 0, fRestwert = cost -
				salvage, bNowLia = false, fGda;

			for (var i = 1; i <= nLoopEnd; i++) {
				if (!bNowLia) {

					fGda = _getDDB(cost, salvage, life, i, factor);
					fLia = fRestwert / (life1 - (i - 1));

					if (fLia > fGda) {
						fTerm = fLia;
						bNowLia = true;
					} else {
						fTerm = fGda;
						fRestwert -= fGda;
					}

				} else {
					fTerm = fLia;
				}

				if (i == nLoopEnd) {
					fTerm *= ( startperiod + 1.0 - end );
				}

				fVdb += fTerm;
			}
			return fVdb;
		}

		function vdb(cost, salvage, life, startPeriod, endPeriod, factor, flag) {

			if (factor === undefined || factor === null) {
				factor = 2;
			}
			if (flag === undefined || flag === null) {
				flag = false;
			}

			var start = Math.floor(startPeriod), end = Math.ceil(endPeriod), loopStart = start, loopEnd = end;

			var res = 0;
			if (flag) {
				for (var i = loopStart + 1; i <= loopEnd; i++) {
					var ddb = _getDDB(cost, salvage, life, i, factor);

					if (i == loopStart + 1) {
						ddb *= ( Math.min(endPeriod, start + 1) - startPeriod );
					} else if (i ==
						loopEnd) {
						ddb *= ( endPeriod + 1 - end );
					}

					res += ddb;
				}
			} else {

				var life1 = life;

				if (!Math.approxEqual(startPeriod, Math.floor(startPeriod))) {
					if (factor > 1) {
						if (startPeriod > life / 2 || Math.approxEqual(startPeriod, life / 2)) {
							var fPart = startPeriod - life / 2;
							startPeriod = life / 2;
							endPeriod -= fPart;
							life1 += 1;
						}
					}
				}

				cost -= _getVDB(cost, salvage, life, life1, startPeriod, factor);
				res = _getVDB(cost, salvage, life, life - startPeriod, endPeriod - startPeriod, factor);
			}

			return res;

		}

		oParser = new parserFormula("VDB(2400,300,10*365,0,1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), vdb(2400, 300, 10 * 365, 0, 1));

		oParser = new parserFormula("VDB(2400,300,10*12,0,1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), vdb(2400, 300, 10 * 12, 0, 1));

		oParser = new parserFormula("VDB(2400,300,10*12,6,18)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), vdb(2400, 300, 10 * 12, 6, 18));

		oParser = new parserFormula("VDB(0,0,0,0,0)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!");

		oParser = new parserFormula("VDB(100,0,5,3,4,2,{1,2,3})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue().toFixed(2) - 0, 8.64);

		oParser = new parserFormula('VDB(100,0,5,3,4,2,1)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue().toFixed(2) - 0, 8.64);

		oParser = new parserFormula('VDB(100,0,5,3,4,2,FALSE)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue().toFixed(2) - 0, 10.8);

		oParser = new parserFormula('VDB(100,0,5,3,4,2,"test")', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		testArrayFormula2(assert, "VDB", 5, 7);
	});

    QUnit.test("Test: \"ODDFPRICE\"", function (assert) {

        oParser = new parserFormula( "ODDFPRICE(DATE(1999,2,28),DATE(2016,1,1),DATE(1998,2,28),DATE(2015,1,1),7%,0,100,2,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( difBetween(oParser.calculate().getValue(), 217.878453038674) );

        oParser = new parserFormula( "ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( difBetween(oParser.calculate().getValue(), 113.597717474079) );

        oParser = new parserFormula( "ODDFPRICE(DATE(1990,6,1),DATE(1995,12,31),DATE(1990,1,1),DATE(1990,12,31),6%,5%,1000,1,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( difBetween(oParser.calculate().getValue(), 790.11323221867) );

        testArrayFormula2(assert, "ODDFPRICE", 8, 9, true);
    } );

    QUnit.test("Test: \"ODDFYIELD\"", function (assert) {

        oParser = new parserFormula( "ODDFYIELD(DATE(1990,6,1),DATE(1995,12,31),DATE(1990,1,1),DATE(1990,12,31),6%,790,100,1,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "ODDFYIELD(DATE(1990,6,1),DATE(1995,12,31),DATE(1990,1,1),DATE(1990,12,31),6%,790,100,1,1)" );
        assert.ok( difBetween(oParser.calculate().getValue(),-0.2889178784774840 ) );

        oParser = new parserFormula( "ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)" );
        assert.ok( difBetween(oParser.calculate().getValue(), 0.0772455415972989 ) );

        oParser = new parserFormula( "ODDFYIELD(DATE(2008,12,11),DATE(2021,4,1),DATE(2008,10,15),DATE(2009,4,1),6%,100,100,4,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "ODDFYIELD(DATE(2008,12,11),DATE(2021,4,1),DATE(2008,10,15),DATE(2009,4,1),6%,100,100,4,1)" );
        assert.ok( difBetween(oParser.calculate().getValue(), 0.0599769985558904 ) );

		testArrayFormula2(assert, "ODDFYIELD", 8, 9, true);
    } );

    /*
    * Engineering
    * */

    QUnit.test("Test: \"BIN2DEC\"", function (assert) {

        oParser = new parserFormula( "BIN2DEC(101010)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2DEC(101010)" );
        assert.strictEqual( oParser.calculate().getValue(), 42 );

        oParser = new parserFormula( "BIN2DEC(\"101010\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2DEC(\"101010\")" );
        assert.strictEqual( oParser.calculate().getValue(), 42 );

        oParser = new parserFormula( "BIN2DEC(111111111)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2DEC(111111111)" );
        assert.strictEqual( oParser.calculate().getValue(), 511 );

        oParser = new parserFormula( "BIN2DEC(1000000000)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2DEC(1000000000)" );
        assert.strictEqual( oParser.calculate().getValue(), -512 );

        oParser = new parserFormula( "BIN2DEC(1111111111)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2DEC(1111111111)" );
        assert.strictEqual( oParser.calculate().getValue(), -1 );

        oParser = new parserFormula( "BIN2DEC(1234567890)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2DEC(1234567890)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "BIN2DEC(\"Hello World!\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2DEC(\"Hello World!\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        testArrayFormula2(assert, "BIN2DEC",1,1,true);
    });

    QUnit.test("Test: \"BIN2HEX\"", function (assert) {

        oParser = new parserFormula( "BIN2HEX(101010)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(101010)" );
        assert.strictEqual( oParser.calculate().getValue(), "2A" );

        oParser = new parserFormula( "BIN2HEX(\"101010\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(\"101010\")" );
        assert.strictEqual( oParser.calculate().getValue(), "2A" );

        oParser = new parserFormula( "BIN2HEX(111111111)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(111111111)" );
        assert.strictEqual( oParser.calculate().getValue(), "1FF" );

        oParser = new parserFormula( "BIN2HEX(1000000000)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(1000000000)" );
        assert.strictEqual( oParser.calculate().getValue(), "FFFFFFFE00" );

        oParser = new parserFormula( "BIN2HEX(1111111111)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(1111111111)" );
        assert.strictEqual( oParser.calculate().getValue(), "FFFFFFFFFF" );

        oParser = new parserFormula( "BIN2HEX(101010,2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(101010,2)" );
        assert.strictEqual( oParser.calculate().getValue(), "2A" );

        oParser = new parserFormula( "BIN2HEX(101010,4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(101010,4)" );
        assert.strictEqual( oParser.calculate().getValue(), "002A" );

        oParser = new parserFormula( "BIN2HEX(101010,4.5)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(101010,4.5)" );
        assert.strictEqual( oParser.calculate().getValue(), "002A" );

        oParser = new parserFormula( "BIN2HEX(1234567890)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(1234567890)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "BIN2HEX(\"Hello World!\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(\"Hello World!\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "BIN2HEX(101010101010)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(101010101010)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "BIN2HEX(101010,1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(101010,1)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "BIN2HEX(101010,-4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(101010,-4)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "BIN2HEX(101010, \"Hello World!\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2HEX(101010,\"Hello World!\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        testArrayFormula2(assert, "BIN2HEX", 1, 2, true)

    });

    QUnit.test("Test: \"BIN2OCT\"", function (assert) {

        oParser = new parserFormula( "BIN2OCT(101010)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(101010)" );
        assert.strictEqual( oParser.calculate().getValue(), "52" );

        oParser = new parserFormula( "BIN2OCT(\"101010\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(\"101010\")" );
        assert.strictEqual( oParser.calculate().getValue(), "52" );

        oParser = new parserFormula( "BIN2OCT(111111111)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(111111111)" );
        assert.strictEqual( oParser.calculate().getValue(), "777" );

        oParser = new parserFormula( "BIN2OCT(1000000000)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(1000000000)" );
        assert.strictEqual( oParser.calculate().getValue(), "7777777000" );

        oParser = new parserFormula( "BIN2OCT(1111111111)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(1111111111)" );
        assert.strictEqual( oParser.calculate().getValue(), "7777777777" );

        oParser = new parserFormula( "BIN2OCT(101010, 2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(101010,2)" );
        assert.strictEqual( oParser.calculate().getValue(), "52" );

        oParser = new parserFormula( "BIN2OCT(101010, 4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(101010,4)" );
        assert.strictEqual( oParser.calculate().getValue(), "0052" );

        oParser = new parserFormula( "BIN2OCT(101010, 4.5)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(101010,4.5)" );
        assert.strictEqual( oParser.calculate().getValue(), "0052" );

        oParser = new parserFormula( "BIN2OCT(\"Hello World!\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(\"Hello World!\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "BIN2OCT(1234567890)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(1234567890)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "BIN2OCT(101010101010)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(101010101010)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "BIN2OCT(101010, 1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(101010,1)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "BIN2OCT(101010, -4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(101010,-4)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "BIN2OCT(101010, \"Hello World!\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "BIN2OCT(101010,\"Hello World!\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		testArrayFormula2(assert, "BIN2OCT", 1, 2, true);
    });

    QUnit.test("Test: \"DEC2BIN\"", function (assert) {

        oParser = new parserFormula( "DEC2BIN(42)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(42)" );
        assert.strictEqual( oParser.calculate().getValue(), "101010" );

        oParser = new parserFormula( "DEC2BIN(\"42\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(\"42\")" );
        assert.strictEqual( oParser.calculate().getValue(), "101010" );

        oParser = new parserFormula( "DEC2BIN(-512)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(-512)" );
        assert.strictEqual( oParser.calculate().getValue(), "1000000000" );

        oParser = new parserFormula( "DEC2BIN(-511)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(-511)" );
        assert.strictEqual( oParser.calculate().getValue(), "1000000001" );

        oParser = new parserFormula( "DEC2BIN(-1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(-1)" );
        assert.strictEqual( oParser.calculate().getValue(), "1111111111" );

        oParser = new parserFormula( "DEC2BIN(0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(0)" );
        assert.strictEqual( oParser.calculate().getValue(), "0" );

        oParser = new parserFormula( "DEC2BIN(1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(1)" );
        assert.strictEqual( oParser.calculate().getValue(), "1" );

        oParser = new parserFormula( "DEC2BIN(510)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(510)" );
        assert.strictEqual( oParser.calculate().getValue(), "111111110" );

        oParser = new parserFormula( "DEC2BIN(511)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(511)" );
        assert.strictEqual( oParser.calculate().getValue(), "111111111" );

        oParser = new parserFormula( "DEC2BIN(42, 6)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(42,6)" );
        assert.strictEqual( oParser.calculate().getValue(), "101010" );

        oParser = new parserFormula( "DEC2BIN(42, 8)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(42,8)" );
        assert.strictEqual( oParser.calculate().getValue(), "00101010" );

        oParser = new parserFormula( "DEC2BIN(\"Hello World!\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(\"Hello World!\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        oParser = new parserFormula( "DEC2BIN(\"2a\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(\"2a\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        oParser = new parserFormula( "DEC2BIN(-513)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(-513)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "DEC2BIN(512)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(512)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "DEC2BIN(42, -8)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2BIN(42,-8)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		testArrayFormula2(assert, "DEC2BIN", 1, 2, true)
    });

    QUnit.test("Test: \"DEC2HEX\"", function (assert) {

        oParser = new parserFormula( "DEC2HEX(42)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2HEX(42)" );
        assert.strictEqual( oParser.calculate().getValue(), "2A" );

        oParser = new parserFormula( "DEC2HEX(\"42\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2HEX(\"42\")" );
        assert.strictEqual( oParser.calculate().getValue(), "2A" );

        oParser = new parserFormula( "DEC2HEX(-549755813888)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2HEX(-549755813888)" );
        assert.strictEqual( oParser.calculate().getValue(), "8000000000" );

        oParser = new parserFormula( "DEC2HEX(-549755813887)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2HEX(-549755813887)" );
        assert.strictEqual( oParser.calculate().getValue(), "8000000001" );

        oParser = new parserFormula( "DEC2HEX(-1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2HEX(-1)" );
        assert.strictEqual( oParser.calculate().getValue(), "FFFFFFFFFF" );

        oParser = new parserFormula( "DEC2HEX(0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2HEX(0)" );
        assert.strictEqual( oParser.calculate().getValue(), "0" );

        oParser = new parserFormula( "DEC2HEX(1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2HEX(1)" );
        assert.strictEqual( oParser.calculate().getValue(), "1" );

        oParser = new parserFormula( "DEC2HEX(549755813886)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2HEX(549755813886)" );
        assert.strictEqual( oParser.calculate().getValue(), "7FFFFFFFFE" );

        oParser = new parserFormula( "DEC2HEX(549755813887)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2HEX(549755813887)" );
        assert.strictEqual( oParser.calculate().getValue(), "7FFFFFFFFF" );

        oParser = new parserFormula( "DEC2HEX(42, 2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2HEX(42,2)" );
        assert.strictEqual( oParser.calculate().getValue(), "2A" );

        oParser = new parserFormula( "DEC2HEX(42, 4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2HEX(42,4)" );
        assert.strictEqual( oParser.calculate().getValue(), "002A" );

        oParser = new parserFormula( "DEC2HEX(\"Hello World!\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2HEX(\"Hello World!\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        oParser = new parserFormula( "DEC2HEX(\"2a\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2HEX(\"2a\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		testArrayFormula2(assert, "DEC2HEX", 1, 2, true);
    });

    QUnit.test("Test: \"DEC2OCT\"", function (assert) {

        oParser = new parserFormula( "DEC2OCT(42)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(42)" );
        assert.strictEqual( oParser.calculate().getValue(), "52" );

        oParser = new parserFormula( "DEC2OCT(\"42\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(\"42\")" );
        assert.strictEqual( oParser.calculate().getValue(), "52" );

        oParser = new parserFormula( "DEC2OCT(-536870912)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(-536870912)" );
        assert.strictEqual( oParser.calculate().getValue(), "4000000000" );

        oParser = new parserFormula( "DEC2OCT(-536870911)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(-536870911)" );
        assert.strictEqual( oParser.calculate().getValue(), "4000000001" );

        oParser = new parserFormula( "DEC2OCT(-1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(-1)" );
        assert.strictEqual( oParser.calculate().getValue(), "7777777777" );

        oParser = new parserFormula( "DEC2OCT(0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(0)" );
        assert.strictEqual( oParser.calculate().getValue(), "0" );

        oParser = new parserFormula( "DEC2OCT(-0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(-0)" );
        assert.strictEqual( oParser.calculate().getValue(), "0" );

        oParser = new parserFormula( "DEC2OCT(1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(1)" );
        assert.strictEqual( oParser.calculate().getValue(), "1" );

        oParser = new parserFormula( "DEC2OCT(536870910)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(536870910)" );
        assert.strictEqual( oParser.calculate().getValue(), "3777777776" );

        oParser = new parserFormula( "DEC2OCT(536870911)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(536870911)" );
        assert.strictEqual( oParser.calculate().getValue(), "3777777777" );

        oParser = new parserFormula( "DEC2OCT(42, 2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(42,2)" );
        assert.strictEqual( oParser.calculate().getValue(), "52" );

        oParser = new parserFormula( "DEC2OCT(42, 4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(42,4)" );
        assert.strictEqual( oParser.calculate().getValue(), "0052" );

        oParser = new parserFormula( "DEC2OCT(\"Hello World!\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(\"Hello World!\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        oParser = new parserFormula( "DEC2OCT(\"2a\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(\"2a\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

        oParser = new parserFormula( "DEC2OCT(-536870913)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(-536870913)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "DEC2OCT(536870912)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(536870912)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "DEC2OCT(42, 1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "DEC2OCT(42,1)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

		testArrayFormula2(assert, "DEC2OCT", 1, 2, true);
    });

    QUnit.test("Test: \"HEX2BIN\"", function (assert) {

        oParser = new parserFormula( "HEX2BIN(\"2a\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"2a\")" );
        assert.strictEqual( oParser.calculate().getValue(), "101010" );

        oParser = new parserFormula( "HEX2BIN(\"fffffffe00\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"fffffffe00\")" );
        assert.strictEqual( oParser.calculate().getValue(), "1000000000" );

        oParser = new parserFormula( "HEX2BIN(\"fffffffe01\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"fffffffe01\")" );
        assert.strictEqual( oParser.calculate().getValue(), "1000000001" );

        oParser = new parserFormula( "HEX2BIN(\"ffffffffff\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"ffffffffff\")" );
        assert.strictEqual( oParser.calculate().getValue(), "1111111111" );

        oParser = new parserFormula( "HEX2BIN(0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(0)" );
        assert.strictEqual( oParser.calculate().getValue(), "0" );

        oParser = new parserFormula( "HEX2BIN(1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(1)" );
        assert.strictEqual( oParser.calculate().getValue(), "1" );

        oParser = new parserFormula( "HEX2BIN(\"1fe\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"1fe\")" );
        assert.strictEqual( oParser.calculate().getValue(), "111111110" );

        oParser = new parserFormula( "HEX2BIN(\"1ff\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"1ff\")" );
        assert.strictEqual( oParser.calculate().getValue(), "111111111" );

        oParser = new parserFormula( "HEX2BIN(\"2a\",6)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"2a\",6)" );
        assert.strictEqual( oParser.calculate().getValue(), "101010" );

        oParser = new parserFormula( "HEX2BIN(\"2a\",8)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"2a\",8)" );
        assert.strictEqual( oParser.calculate().getValue(), "00101010" );

        oParser = new parserFormula( "HEX2BIN(\"Hello World!\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"Hello World!\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "HEX2BIN(\"fffffffdff\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"fffffffdff\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "HEX2BIN(\"200\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"200\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "HEX2BIN(\"2a\", 5)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"2a\",5)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "HEX2BIN(\"2a\", -8)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"2a\",-8)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );

        oParser = new parserFormula( "HEX2BIN(\"2a\", \"Hello World!\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2BIN(\"2a\",\"Hello World!\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		testArrayFormula2(assert, "HEX2BIN", 1, 2, true);
    });

    QUnit.test("Test: \"HEX2DEC\"", function (assert) {

        oParser = new parserFormula( "HEX2DEC(\"2a\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2DEC(\"2a\")" );
        assert.strictEqual( oParser.calculate().getValue(), 42);

        oParser = new parserFormula( "HEX2DEC(\"8000000000\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2DEC(\"8000000000\")" );
        assert.strictEqual( oParser.calculate().getValue(), -549755813888);

        oParser = new parserFormula( "HEX2DEC(\"ffffffffff\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2DEC(\"ffffffffff\")" );
        assert.strictEqual( oParser.calculate().getValue(), -1);

        oParser = new parserFormula( "HEX2DEC(\"0\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2DEC(\"0\")" );
        assert.strictEqual( oParser.calculate().getValue(), 0);

        oParser = new parserFormula( "HEX2DEC(\"1\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2DEC(\"1\")" );
        assert.strictEqual( oParser.calculate().getValue(), 1);

        oParser = new parserFormula( "HEX2DEC(0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2DEC(0)" );
        assert.strictEqual( oParser.calculate().getValue(), 0);

        oParser = new parserFormula( "HEX2DEC(1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2DEC(1)" );
        assert.strictEqual( oParser.calculate().getValue(), 1);

        oParser = new parserFormula( "HEX2DEC(\"7fffffffff\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2DEC(\"7fffffffff\")" );
        assert.strictEqual( oParser.calculate().getValue(), 549755813887);

		testArrayFormula2(assert, "HEX2DEC", 1, 1, true);
    });

    QUnit.test("Test: \"HEX2OCT\"", function (assert) {

        oParser = new parserFormula( "HEX2OCT(\"2a\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2OCT(\"2a\")" );
        assert.strictEqual( oParser.calculate().getValue(), "52");

        oParser = new parserFormula( "HEX2OCT(\"ffe0000000\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2OCT(\"ffe0000000\")" );
        assert.strictEqual( oParser.calculate().getValue(), "4000000000");

        oParser = new parserFormula( "HEX2OCT(\"ffe0000001\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2OCT(\"ffe0000001\")" );
        assert.strictEqual( oParser.calculate().getValue(), "4000000001");

        oParser = new parserFormula( "HEX2OCT(0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2OCT(0)" );
        assert.strictEqual( oParser.calculate().getValue(), "0");

        oParser = new parserFormula( "HEX2OCT(1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2OCT(1)" );
        assert.strictEqual( oParser.calculate().getValue(), "1");

        oParser = new parserFormula( "HEX2OCT(\"1ffffffe\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2OCT(\"1ffffffe\")" );
        assert.strictEqual( oParser.calculate().getValue(), "3777777776");

        oParser = new parserFormula( "HEX2OCT(\"1fffffff\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2OCT(\"1fffffff\")" );
        assert.strictEqual( oParser.calculate().getValue(), "3777777777");

        oParser = new parserFormula( "HEX2OCT(\"2a\",2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2OCT(\"2a\",2)" );
        assert.strictEqual( oParser.calculate().getValue(), "52");

        oParser = new parserFormula( "HEX2OCT(\"2a\",4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2OCT(\"2a\",4)" );
        assert.strictEqual( oParser.calculate().getValue(), "0052");

        oParser = new parserFormula( "HEX2OCT(\"Hello World!\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2OCT(\"Hello World!\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!");

        oParser = new parserFormula( "HEX2OCT(\"ffdfffffff\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2OCT(\"ffdfffffff\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!");

        oParser = new parserFormula( "HEX2OCT(\"2a\", 1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "HEX2OCT(\"2a\",1)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!");

		testArrayFormula2(assert, "HEX2OCT", 1, 2, true);
    });

    QUnit.test("Test: \"OCT2BIN\"", function (assert) {

        oParser = new parserFormula( "OCT2BIN(\"52\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2BIN(\"52\")" );
        assert.strictEqual( oParser.calculate().getValue(), "101010");

        oParser = new parserFormula( "OCT2BIN(\"7777777000\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2BIN(\"7777777000\")" );
        assert.strictEqual( oParser.calculate().getValue(), "1000000000");

        oParser = new parserFormula( "OCT2BIN(\"7777777001\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2BIN(\"7777777001\")" );
        assert.strictEqual( oParser.calculate().getValue(), "1000000001");

        oParser = new parserFormula( "OCT2BIN(\"7777777777\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2BIN(\"7777777777\")" );
        assert.strictEqual( oParser.calculate().getValue(), "1111111111");

        oParser = new parserFormula( "OCT2BIN(\"0\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2BIN(\"0\")" );
        assert.strictEqual( oParser.calculate().getValue(), "0");

        oParser = new parserFormula( "OCT2BIN(\"1\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2BIN(\"1\")" );
        assert.strictEqual( oParser.calculate().getValue(), "1");

        oParser = new parserFormula( "OCT2BIN(\"776\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2BIN(\"776\")" );
        assert.strictEqual( oParser.calculate().getValue(), "111111110");

        oParser = new parserFormula( "OCT2BIN(\"777\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2BIN(\"777\")" );
        assert.strictEqual( oParser.calculate().getValue(), "111111111");

        oParser = new parserFormula( "OCT2BIN(\"52\", 6)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2BIN(\"52\",6)" );
        assert.strictEqual( oParser.calculate().getValue(), "101010");

        oParser = new parserFormula( "OCT2BIN(\"52\", 8)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2BIN(\"52\",8)" );
        assert.strictEqual( oParser.calculate().getValue(), "00101010");

        oParser = new parserFormula( "OCT2BIN(\"Hello World!\", 8)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2BIN(\"Hello World!\",8)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!");

        oParser = new parserFormula( "OCT2BIN(\"52\",\"Hello World!\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2BIN(\"52\",\"Hello World!\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!");

		testArrayFormula2(assert, "OCT2BIN", 1, 2, true)
    });

    QUnit.test("Test: \"OCT2DEC\"", function (assert) {

        oParser = new parserFormula( "OCT2DEC(\"52\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2DEC(\"52\")" );
        assert.strictEqual( oParser.calculate().getValue(), 42);

        oParser = new parserFormula( "OCT2DEC(\"4000000000\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2DEC(\"4000000000\")" );
        assert.strictEqual( oParser.calculate().getValue(), -536870912);

        oParser = new parserFormula( "OCT2DEC(\"7777777777\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2DEC(\"7777777777\")" );
        assert.strictEqual( oParser.calculate().getValue(), -1);

        oParser = new parserFormula( "OCT2DEC(\"0\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2DEC(\"0\")" );
        assert.strictEqual( oParser.calculate().getValue(), 0);

        oParser = new parserFormula( "OCT2DEC(\"1\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2DEC(\"1\")" );
        assert.strictEqual( oParser.calculate().getValue(), 1);

        oParser = new parserFormula( "OCT2DEC(\"3777777776\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2DEC(\"3777777776\")" );
        assert.strictEqual( oParser.calculate().getValue(), 536870910);

        oParser = new parserFormula( "OCT2DEC(\"3777777777\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2DEC(\"3777777777\")" );
        assert.strictEqual( oParser.calculate().getValue(), 536870911);

        oParser = new parserFormula( "OCT2DEC(\"3777777777\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2DEC(\"3777777777\")" );
        assert.strictEqual( oParser.calculate().getValue(), 536870911);

        testArrayFormula2(assert, "OCT2DEC",1,1,true);
    });

    QUnit.test("Test: \"OCT2HEX\"", function (assert) {

        oParser = new parserFormula( "OCT2HEX(\"52\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2HEX(\"52\")" );
        assert.strictEqual( oParser.calculate().getValue(), "2A");

        oParser = new parserFormula( "OCT2HEX(\"4000000000\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2HEX(\"4000000000\")" );
        assert.strictEqual( oParser.calculate().getValue(), "FFE0000000");

        oParser = new parserFormula( "OCT2HEX(\"4000000001\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2HEX(\"4000000001\")" );
        assert.strictEqual( oParser.calculate().getValue(), "FFE0000001");

        oParser = new parserFormula( "OCT2HEX(\"7777777777\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2HEX(\"7777777777\")" );
        assert.strictEqual( oParser.calculate().getValue(), "FFFFFFFFFF");

        oParser = new parserFormula( "OCT2HEX(\"0\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2HEX(\"0\")" );
        assert.strictEqual( oParser.calculate().getValue(), "0");

        oParser = new parserFormula( "OCT2HEX(\"1\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2HEX(\"1\")" );
        assert.strictEqual( oParser.calculate().getValue(), "1");

        oParser = new parserFormula( "OCT2HEX(\"3777777776\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2HEX(\"3777777776\")" );
        assert.strictEqual( oParser.calculate().getValue(), "1FFFFFFE");

        oParser = new parserFormula( "OCT2HEX(\"3777777777\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2HEX(\"3777777777\")" );
        assert.strictEqual( oParser.calculate().getValue(), "1FFFFFFF");

        oParser = new parserFormula( "OCT2HEX(\"52\", 2)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2HEX(\"52\",2)" );
        assert.strictEqual( oParser.calculate().getValue(), "2A");

        oParser = new parserFormula( "OCT2HEX(\"52\", 4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2HEX(\"52\",4)" );
        assert.strictEqual( oParser.calculate().getValue(), "002A");

        oParser = new parserFormula( "OCT2HEX(\"Hello World!\", 4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2HEX(\"Hello World!\",4)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!");

        oParser = new parserFormula( "OCT2HEX(\"52\", -4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2HEX(\"52\",-4)" );
        assert.strictEqual( oParser.calculate().getValue(), "#NUM!");

        oParser = new parserFormula( "OCT2HEX(\"52\", \"Hello World!\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "OCT2HEX(\"52\",\"Hello World!\")" );
        assert.strictEqual( oParser.calculate().getValue(), "#VALUE!");

		testArrayFormula2(assert, "OCT2HEX", 1, 2, true)

    });

    QUnit.test("Test: \"COMPLEX\"", function (assert) {

        oParser = new parserFormula( "COMPLEX(-3.5,19.6)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "COMPLEX(-3.5,19.6)" );
        assert.strictEqual( oParser.calculate().getValue(), "-3.5+19.6i");

        oParser = new parserFormula( "COMPLEX(3.5,-19.6,\"j\")", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "COMPLEX(3.5,-19.6,\"j\")" );
        assert.strictEqual( oParser.calculate().getValue(), "3.5-19.6j");

        oParser = new parserFormula( "COMPLEX(3.5,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "COMPLEX(3.5,0)" );
        assert.strictEqual( oParser.calculate().getValue(), "3.5");

        oParser = new parserFormula( "COMPLEX(0,2.4)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "COMPLEX(0,2.4)" );
        assert.strictEqual( oParser.calculate().getValue(), "2.4i");

        oParser = new parserFormula( "COMPLEX(0,0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.ok( oParser.assemble() == "COMPLEX(0,0)" );
        assert.strictEqual( oParser.calculate().getValue(), "0");

        testArrayFormula2(assert, "COMPLEX", 2, 3, true);
    });

    QUnit.test("Test: \"DELTA\"", function (assert) {

        oParser = new parserFormula( "DELTA(10.5,10.5)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1);

        oParser = new parserFormula( "DELTA(10.5,10.6)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 0);

        oParser = new parserFormula( "DELTA(10.5)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 0);

        oParser = new parserFormula( "DELTA(0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1);

		testArrayFormula2(assert, "DELTA", 1, 2, true);
    });

    QUnit.test("Test: \"ERF\"", function (assert) {

        oParser = new parserFormula( "ERF(1.234,4.5432)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue().toFixed(14)-0, 0.08096058291050978.toFixed(14)-0 );

        oParser = new parserFormula( "ERF(1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue().toFixed(14)-0, 0.8427007929497149.toFixed(14)-0 );

        oParser = new parserFormula( "ERF(0,1.345)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue().toFixed(14)-0, 0.9428441710878559.toFixed(14)-0 );

        oParser = new parserFormula( "ERF(1.234)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue().toFixed(14)-0, 0.9190394169576684.toFixed(14)-0 );

		testArrayFormula2(assert, "ERF", 1, 2, true);
    });

	QUnit.test("Test: \"GESTEP\"", function (assert) {

		oParser = new parserFormula( "GESTEP(5, 4)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1);

		oParser = new parserFormula( "GESTEP(5, 5)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1);

		oParser = new parserFormula( "GESTEP(-4, -5)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1);

		oParser = new parserFormula( "GESTEP(-1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0);

		testArrayFormula2(assert, "GESTEP", 1, 2, true);
	});

	QUnit.test("Test: \"ERF.PRECISE\"", function (assert) {

		oParser = new parserFormula( "ERF.PRECISE(1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(14)-0, 0.8427007929497149.toFixed(14)-0 );

		oParser = new parserFormula( "ERF.PRECISE(1.234)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(14)-0, 0.9190394169576684.toFixed(14)-0 );

		oParser = new parserFormula( "ERF.PRECISE(0.745)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 0.70792892 );

		oParser = new parserFormula( "ERF.PRECISE(1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 0.84270079 );

		testArrayFormula2(assert, "ERF.PRECISE",1,1,true);
	});

    QUnit.test("Test: \"ERFC\"", function (assert) {

        oParser = new parserFormula( "ERFC(1.234)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue().toFixed(14)-0, 0.08096058304233157.toFixed(14)-0 );

        oParser = new parserFormula( "ERFC(1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue().toFixed(14)-0, 0.15729920705028513.toFixed(14)-0 );

        oParser = new parserFormula( "ERFC(0)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue(), 1 );

        oParser = new parserFormula( "ERFC(-1)", "A2", ws );
        assert.ok( oParser.parse() );
        assert.strictEqual( oParser.calculate().getValue().toFixed(14)-0, 1.8427007929497148.toFixed(14)-0 );

        testArrayFormula2(assert, "ERFC",1,1,true);
    });

	QUnit.test("Test: \"ERFC.PRECISE\"", function (assert) {

		oParser = new parserFormula( "ERFC.PRECISE(1.234)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(14)-0, 0.08096058304233157.toFixed(14)-0 );

		oParser = new parserFormula( "ERFC.PRECISE(1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(14)-0, 0.15729920705028513.toFixed(14)-0 );

		oParser = new parserFormula( "ERFC.PRECISE(0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( "ERFC.PRECISE(-1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(14)-0, 1.8427007929497148.toFixed(14)-0 );

	});

	QUnit.test("Test: \"BITAND\"", function (assert) {

		oParser = new parserFormula( 'BITAND(1,5)', "AA2", ws );
		assert.ok( oParser.parse(), 'BITAND(1,5)' );
		assert.strictEqual( oParser.calculate().getValue(), 1, 'BITAND(1,5)' );

		oParser = new parserFormula( 'BITAND(13,25)', "AA2", ws );
		assert.ok( oParser.parse(), 'BITAND(13,25)' );
		assert.strictEqual( oParser.calculate().getValue(), 9, 'BITAND(13,25)' );

		testArrayFormula2(assert, "BITAND", 2, 2);
	});

	QUnit.test("Test: \"BITOR\"", function (assert) {

		oParser = new parserFormula( 'BITOR(23,10)', "AA2", ws );
		assert.ok( oParser.parse());
		assert.strictEqual( oParser.calculate().getValue(), 31 );

		testArrayFormula2(assert, "BITOR", 2, 2);
	});

	QUnit.test("Test: \"BITXOR\"", function (assert) {

		oParser = new parserFormula( 'BITXOR(5,3)', "AA2", ws );
		assert.ok( oParser.parse());
		assert.strictEqual( oParser.calculate().getValue(), 6 );

		testArrayFormula2(assert, "BITXOR", 2, 2);
	});

	QUnit.test("Test: \"BITRSHIFT\"", function (assert) {

		oParser = new parserFormula( 'BITRSHIFT(13,2)', "AA2", ws );
		assert.ok( oParser.parse());
		assert.strictEqual( oParser.calculate().getValue(), 3 );

		testArrayFormula2(assert, "BITRSHIFT", 2, 2);
	});

	QUnit.test("Test: \"BITLSHIFT\"", function (assert) {

		oParser = new parserFormula( 'BITLSHIFT(4,2)', "AA2", ws );
		assert.ok( oParser.parse());
		assert.strictEqual( oParser.calculate().getValue(), 16 );

		testArrayFormula2(assert, "BITLSHIFT", 2, 2);
	});

	function putDataForDatabase(){
		ws.getRange2( "A1" ).setValue( "Tree" );
		ws.getRange2( "A2" ).setValue( "Apple" );
		ws.getRange2( "A3" ).setValue( "Pear" );

		ws.getRange2( "A4" ).setValue( "Tree" );

		ws.getRange2( "A5" ).setValue( "Apple" );
		ws.getRange2( "A6" ).setValue( "Pear" );
		ws.getRange2( "A7" ).setValue( "Cherry" );
		ws.getRange2( "A8" ).setValue( "Apple" );
		ws.getRange2( "A9" ).setValue( "Pear" );
		ws.getRange2( "A10" ).setValue( "Apple" );


		ws.getRange2( "B1" ).setValue( "Height" );
		ws.getRange2( "B2" ).setValue( ">10" );
		ws.getRange2( "B3" ).setValue( "" );

		ws.getRange2( "B4" ).setValue( "Height" );

		ws.getRange2( "B5" ).setValue( "18" );
		ws.getRange2( "B6" ).setValue( "12" );
		ws.getRange2( "B7" ).setValue( "13" );
		ws.getRange2( "B8" ).setValue( "14" );
		ws.getRange2( "B9" ).setValue( "9" );
		ws.getRange2( "B10" ).setValue( "8" );


		ws.getRange2( "C1" ).setValue( "Age" );
		ws.getRange2( "C2" ).setValue( "" );
		ws.getRange2( "C3" ).setValue( "" );

		ws.getRange2( "C4" ).setValue( "Age" );

		ws.getRange2( "C5" ).setValue( "20" );
		ws.getRange2( "C6" ).setValue( "12" );
		ws.getRange2( "C7" ).setValue( "14" );
		ws.getRange2( "C8" ).setValue( "15" );
		ws.getRange2( "C9" ).setValue( "8" );
		ws.getRange2( "C10" ).setValue( "9" );


		ws.getRange2( "C1" ).setValue( "Age" );
		ws.getRange2( "C2" ).setValue( "" );
		ws.getRange2( "C3" ).setValue( "" );

		ws.getRange2( "C4" ).setValue( "Age" );

		ws.getRange2( "C5" ).setValue( "20" );
		ws.getRange2( "C6" ).setValue( "12" );
		ws.getRange2( "C7" ).setValue( "14" );
		ws.getRange2( "C8" ).setValue( "15" );
		ws.getRange2( "C9" ).setValue( "8" );
		ws.getRange2( "C10" ).setValue( "9" );


		ws.getRange2( "D1" ).setValue( "Yield" );
		ws.getRange2( "D2" ).setValue( "" );
		ws.getRange2( "D3" ).setValue( "" );

		ws.getRange2( "D4" ).setValue( "Yield" );

		ws.getRange2( "D5" ).setValue( "14" );
		ws.getRange2( "D6" ).setValue( "10" );
		ws.getRange2( "D7" ).setValue( "9" );
		ws.getRange2( "D8" ).setValue( "10" );
		ws.getRange2( "D9" ).setValue( "8" );
		ws.getRange2( "D10" ).setValue( "6" );


		ws.getRange2( "E1" ).setValue( "Profit" );
		ws.getRange2( "E2" ).setValue( "" );
		ws.getRange2( "E3" ).setValue( "" );

		ws.getRange2( "E4" ).setValue( "Profit" );

		ws.getRange2( "E5" ).setValue( "105" );
		ws.getRange2( "E6" ).setValue( "96" );
		ws.getRange2( "E7" ).setValue( "105" );
		ws.getRange2( "E8" ).setValue( "75" );
		ws.getRange2( "E9" ).setValue( "76.8" );
		ws.getRange2( "E10" ).setValue( "45" );

		ws.getRange2( "F1" ).setValue( "Height" );
		ws.getRange2( "F2" ).setValue( "<16" );
		ws.getRange2( "F3" ).setValue( "" );
    }

	//database formulas
	QUnit.test("Test: \"DAVERAGE\"", function (assert) {

		putDataForDatabase();

		oParser = new parserFormula( 'DAVERAGE(A4:E10, "Yield", A1:B2)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 12 );

		oParser = new parserFormula( 'DAVERAGE(A4:E10, 3, A4:E10)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 13 );

	});

	QUnit.test("Test: \"DCOUNT\"", function (assert) {

		putDataForDatabase();

		oParser = new parserFormula( 'DCOUNT(A4:E10, "Age", A1:F2)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( 'DCOUNT(A4:E10,, A1:F2)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( 'DCOUNT(A4:E10,"", A1:F2)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

	});

	QUnit.test("Test: \"DCOUNTA\"", function (assert) {

		putDataForDatabase();

		oParser = new parserFormula( 'DCOUNTA(A4:E10, "Profit", A1:F2)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( 'DCOUNTA(A4:E10,, A1:F2)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 1 );

		oParser = new parserFormula( 'DCOUNTA(A4:E10,"", A1:F2)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

	});

	QUnit.test("Test: \"DGET\"", function (assert) {

		putDataForDatabase();

		oParser = new parserFormula( 'DGET(A4:E10, "Yield", A1:A3)', "AA2", ws );
		assert.ok( oParser.parse(), 'DGET(A4:E10, "Yield", A1:A3)' );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!", 'DGET(A4:E10, "Yield", A1:A3)' );

		oParser = new parserFormula( 'DGET(A4:E10, "Yield", A1:F2)', "AA2", ws );
		assert.ok( oParser.parse(), 'DGET(A4:E10, "Yield", A1:F2)' );
		assert.strictEqual( oParser.calculate().getValue(), 10, 'DGET(A4:E10, "Yield", A1:F2)' );

	});

	QUnit.test("Test: \"DMAX\"", function (assert) {

		putDataForDatabase();

		oParser = new parserFormula( 'DMAX(A4:E10, "Profit", A1:F3)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 96 );

	});

	QUnit.test("Test: \"DMIN\"", function (assert) {

		putDataForDatabase();

		oParser = new parserFormula( 'DMIN(A4:E10, "Profit", A1:F3)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 75 );

	});

	QUnit.test("Test: \"DPRODUCT\"", function (assert) {

		putDataForDatabase();

		oParser = new parserFormula( 'DPRODUCT(A4:E10, "Yield", A1:F3)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 800 );

	});

	QUnit.test("Test: \"DSTDEV\"", function (assert) {

		putDataForDatabase();

		oParser = new parserFormula( 'DSTDEV(A4:E10, "Yield", A1:F3)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(4) - 0, 1.1547);

	});

	QUnit.test("Test: \"DSTDEVP\"", function (assert) {

		putDataForDatabase();

		oParser = new parserFormula( 'DSTDEVP(A4:E10, "Yield", A1:F3)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(6) - 0, 0.942809);

	});

	QUnit.test("Test: \"STDEVPA\"", function (assert) {

		ws.getRange2( "A103" ).setValue( "1345" );

		ws.getRange2( "A104" ).setValue( "1301" );

		ws.getRange2( "A105" ).setValue( "1368" );
		ws.getRange2( "A106" ).setValue( "1322" );
		ws.getRange2( "A107" ).setValue( "1310" );
		ws.getRange2( "A108" ).setValue( "1370" );
		ws.getRange2( "A109" ).setValue( "1318" );
		ws.getRange2( "A110" ).setValue( "1350" );

		ws.getRange2( "A111" ).setValue( "1303" );
		ws.getRange2( "A112" ).setValue( "1299" );

		oParser = new parserFormula( 'STDEVPA(A103:A112)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(5) - 0, 26.05456);

		testArrayFormula2(assert, "STDEVPA", 1, 8, null, true);
	});

	QUnit.test("Test: \"STDEVP\"", function (assert) {

		ws.getRange2( "A103" ).setValue( "1345" );

		ws.getRange2( "A104" ).setValue( "1301" );

		ws.getRange2( "A105" ).setValue( "1368" );
		ws.getRange2( "A106" ).setValue( "1322" );
		ws.getRange2( "A107" ).setValue( "1310" );
		ws.getRange2( "A108" ).setValue( "1370" );
		ws.getRange2( "A109" ).setValue( "1318" );
		ws.getRange2( "A110" ).setValue( "1350" );

		ws.getRange2( "A111" ).setValue( "1303" );
		ws.getRange2( "A112" ).setValue( "1299" );

		oParser = new parserFormula( 'STDEVP(A103:A112)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(5) - 0, 26.05456);

		testArrayFormula2(assert, "STDEVP", 1, 8, null, true);
	});

	QUnit.test("Test: \"STDEV\"", function (assert) {

		ws.getRange2( "A103" ).setValue( "1345" );

		ws.getRange2( "A104" ).setValue( "1301" );

		ws.getRange2( "A105" ).setValue( "1368" );
		ws.getRange2( "A106" ).setValue( "1322" );
		ws.getRange2( "A107" ).setValue( "1310" );
		ws.getRange2( "A108" ).setValue( "1370" );
		ws.getRange2( "A109" ).setValue( "1318" );
		ws.getRange2( "A110" ).setValue( "1350" );

		ws.getRange2( "A111" ).setValue( "1303" );
		ws.getRange2( "A112" ).setValue( "1299" );

		oParser = new parserFormula( 'STDEV(A103:A112)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(5) - 0, 27.46392);

		testArrayFormula2(assert, "STDEV", 1, 8, null, true);
	});


	QUnit.test("Test: \"DSUM\"", function (assert) {

		putDataForDatabase();

		oParser = new parserFormula( 'DSUM(A4:E10,"Profit",A1:A2)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 225);

		oParser = new parserFormula( 'DSUM(A4:E10,"Profit", A1:F3)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 247.8);

	});

	QUnit.test("Test: \"DVAR\"", function (assert) {

		putDataForDatabase();

		oParser = new parserFormula( 'DVAR(A4:E10, "Yield", A1:A3)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(1) - 0, 8.8);

	});

	QUnit.test("Test: \"DVARP\"", function (assert) {

		putDataForDatabase();

		oParser = new parserFormula( 'DVARP(A4:E10, "Yield", A1:A3)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(2) - 0, 7.04);

	});

	QUnit.test("Test: \"UNICODE\"", function (assert) {

		oParser = new parserFormula( 'UNICODE(" ")', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 32);

		oParser = new parserFormula( 'UNICODE("B")', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 66);

		oParser = new parserFormula( 'UNICODE(0)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 48);

		oParser = new parserFormula( 'UNICODE(1)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 49);

		oParser = new parserFormula( 'UNICODE("true")', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 116);

		oParser = new parserFormula( 'UNICODE(#N/A)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A");

	});

	QUnit.test("Test: \"UNICHAR\"", function (assert) {

		oParser = new parserFormula( 'UNICHAR(66)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "B");

		oParser = new parserFormula( 'UNICHAR(32)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), " ");

		oParser = new parserFormula( 'UNICHAR(0)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula( 'UNICHAR(48)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "0");

		oParser = new parserFormula( 'UNICHAR(49)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "1");

	});

	QUnit.test("Test: \"UPPER\"", function (assert) {
		ws.getRange2( "A2" ).setValue( "total" );
		ws.getRange2( "A3" ).setValue( "Yield" );

		oParser = new parserFormula( 'UPPER(A2)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TOTAL");

		oParser = new parserFormula( 'UPPER(A3)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "YIELD");

		testArrayFormula2(assert, "UPPER", 1, 1);
	});

	QUnit.test("Test: \"UNIQUE \"", function (assert) {

		ws.getRange2( "A101" ).setValue( "1" );
		ws.getRange2( "A102" ).setValue( "2" );
		ws.getRange2( "A103" ).setValue( "2" );
		ws.getRange2( "A104" ).setValue( "-1" );
		ws.getRange2( "A105" ).setValue( "-1" );
		ws.getRange2( "A106" ).setValue( "ds" );
		ws.getRange2( "A107" ).setValue( "ds" );
		ws.getRange2( "A108" ).setValue( "#NUM!" );
		ws.getRange2( "A109" ).setValue( "#NUM!" );

		ws.getRange2( "B101" ).setValue( "1" );
		ws.getRange2( "B102" ).setValue( "2" );
		ws.getRange2( "B103" ).setValue( "2" );
		ws.getRange2( "B104" ).setValue( "4" );
		ws.getRange2( "B105" ).setValue( "5" );
		ws.getRange2( "B106" ).setValue( "7" );
		ws.getRange2( "B107" ).setValue( "7" );
		ws.getRange2( "B108" ).setValue( "8" );
		ws.getRange2( "B109" ).setValue( "8" );

		ws.getRange2( "C101" ).setValue( "2" );
		ws.getRange2( "C102" ).setValue( "2" );
		ws.getRange2( "C103" ).setValue( "2" );
		ws.getRange2( "C104" ).setValue( "1" );
		ws.getRange2( "C105" ).setValue( "1" );
		ws.getRange2( "C106" ).setValue( "2" );
		ws.getRange2( "C107" ).setValue( "3" );
		ws.getRange2( "C108" ).setValue( "8" );
		ws.getRange2( "C109" ).setValue( "8" );

		ws.getRange2( "D101" ).setValue( "2" );
		ws.getRange2( "D102" ).setValue( "2" );
		ws.getRange2( "D103" ).setValue( "2" );
		ws.getRange2( "D104" ).setValue( "1" );
		ws.getRange2( "D105" ).setValue( "1" );
		ws.getRange2( "D106" ).setValue( "2" );
		ws.getRange2( "D107" ).setValue( "3" );
		ws.getRange2( "D108" ).setValue( "8" );
		ws.getRange2( "D109" ).setValue( "8" );

		oParser = new parserFormula( "UNIQUE(A101:A109)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,0).getValue(), 2 );
		assert.strictEqual( oParser.calculate().getElementRowCol(2,0).getValue(), -1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(3,0).getValue(), "ds" );
		assert.strictEqual( oParser.calculate().getElementRowCol(4,0).getValue(), "#NUM!" );

		oParser = new parserFormula( "UNIQUE(A101:A109)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,0).getValue(), 2 );
		assert.strictEqual( oParser.calculate().getElementRowCol(2,0).getValue(), -1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(3,0).getValue(), "ds" );
		assert.strictEqual( oParser.calculate().getElementRowCol(4,0).getValue(), "#NUM!" );

		oParser = new parserFormula( "UNIQUE(A101:D101)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue(), 1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,2).getValue(), 2 );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,3).getValue(), 2 );

		oParser = new parserFormula( "UNIQUE(A101:D101, true)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), 1 );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue(), 2 );

		oParser = new parserFormula( "UNIQUE(A101:D101, true, true)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!" );

		ws.getRange2( "F102" ).setValue( "test" );
		ws.getRange2( "F103" ).setValue( "#VALUE!" );
		ws.getRange2( "F104" ).setValue( "test" );
		ws.getRange2( "F105" ).setValue( "#VALUE!" );
		ws.getRange2( "F106" ).setValue( "2" );
		ws.getRange2( "F107" ).setValue( "-3" );

		ws.getRange2( "G102" ).setValue( "2" );
		ws.getRange2( "G103" ).setValue( "yyy" );
		ws.getRange2( "G104" ).setValue( "4" );
		ws.getRange2( "G105" ).setValue( "yyy" );
		ws.getRange2( "G106" ).setValue( "asd" );
		ws.getRange2( "G107" ).setValue( "7" );

		ws.getRange2( "H102" ).setValue( "test" );
		ws.getRange2( "H103" ).setValue( "#VALUE!" );
		ws.getRange2( "H104" ).setValue( "test" );
		ws.getRange2( "H105" ).setValue( "#VALUE!" );
		ws.getRange2( "H106" ).setValue( "2" );
		ws.getRange2( "H107" ).setValue( "-3" );

		ws.getRange2( "I102" ).setValue( "2" );
		ws.getRange2( "I103" ).setValue( "123" );
		ws.getRange2( "I104" ).setValue( "4" );
		ws.getRange2( "I105" ).setValue( "123" );
		ws.getRange2( "I106" ).setValue( "6" );
		ws.getRange2( "I107" ).setValue( "4" );

		oParser = new parserFormula( "UNIQUE(F102:I107)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,2).getValue(), "test" );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,2).getValue(), "#VALUE!" );
		assert.strictEqual( oParser.calculate().getElementRowCol(2,2).getValue(), "test" );
		assert.strictEqual( oParser.calculate().getElementRowCol(3,2).getValue(), 2 );
		assert.strictEqual( oParser.calculate().getElementRowCol(4,2).getValue(), -3 );

		oParser = new parserFormula( "UNIQUE(F102:I107, true)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,2).getValue(), 2 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,2).getValue(), 123 );
		assert.strictEqual( oParser.calculate().getElementRowCol(2,2).getValue(), 4 );
		assert.strictEqual( oParser.calculate().getElementRowCol(3,2).getValue(), 123 );
		assert.strictEqual( oParser.calculate().getElementRowCol(4,2).getValue(), 6 );
		assert.strictEqual( oParser.calculate().getElementRowCol(5,2).getValue(), 4 );

		oParser = new parserFormula( "UNIQUE(F102:I107, false, true)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue(), "test" );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,1).getValue(), 4 );
		assert.strictEqual( oParser.calculate().getElementRowCol(2,2).getValue(), 2 );
		assert.strictEqual( oParser.calculate().getElementRowCol(3,3).getValue(), 4 );

		oParser = new parserFormula( "UNIQUE(F102:I107, true, true)", "F1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue(), 2 );
		assert.strictEqual( oParser.calculate().getElementRowCol(1,1).getValue(), 123 );
		assert.strictEqual( oParser.calculate().getElementRowCol(2,1).getValue(), 4 );
		assert.strictEqual( oParser.calculate().getElementRowCol(3,1).getValue(), 123 );
		assert.strictEqual( oParser.calculate().getElementRowCol(4,1).getValue(), 6 );
		assert.strictEqual( oParser.calculate().getElementRowCol(5,1).getValue(), 4 );

		ws.getRange2( "A120" ).setValue( "1" );
		ws.getRange2( "B120" ).setValue( "2" );

		oParser = new parserFormula( "UNIQUE(A120:B120)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1);

		oParser = new parserFormula( "UNIQUE({1;2})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1);

		oParser = new parserFormula( "UNIQUE({1,2})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1);
	} );

	QUnit.test("Test: \"GROWTH\"", function (assert) {

		ws.getRange2( "A102" ).setValue( "11" );
		ws.getRange2( "A103" ).setValue( "12" );
		ws.getRange2( "A104" ).setValue( "13" );
		ws.getRange2( "A105" ).setValue( "14" );
		ws.getRange2( "A106" ).setValue( "15" );
		ws.getRange2( "A107" ).setValue( "16" );

		ws.getRange2( "B102" ).setValue( "33100" );
		ws.getRange2( "B103" ).setValue( "47300" );
		ws.getRange2( "B104" ).setValue( "69000" );
		ws.getRange2( "B105" ).setValue( "102000" );
		ws.getRange2( "B106" ).setValue( "150000" );
		ws.getRange2( "B107" ).setValue( "220000" );

		ws.getRange2( "C102" ).setValue( "32618" );
		ws.getRange2( "C103" ).setValue( "47729" );
		ws.getRange2( "C104" ).setValue( "69841" );
		ws.getRange2( "C105" ).setValue( "102197" );
		ws.getRange2( "C106" ).setValue( "149542" );
		ws.getRange2( "C107" ).setValue( "218822" );

		ws.getRange2( "A109" ).setValue( "17" );
		ws.getRange2( "A110" ).setValue( "18" );

		oParser = new parserFormula( "GROWTH(B102:B107,A102:A107,A109:A110)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(4) - 0, 320196.7184);

		oParser = new parserFormula( "GROWTH(B102:B107,A102:A107)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(5) - 0, 32618.20377);

		oParser = new parserFormula( "GROWTH(A102:C102,A103:C104,A105:C106,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 11.00782679);

		oParser = new parserFormula( "GROWTH(A102:C102,A103:C104,A105:C106,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 11.00782679);

		oParser = new parserFormula( "GROWTH(A103:C103,A104:C105,A106:C107,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 12.00187209);

		oParser = new parserFormula( "GROWTH(A103:C103,A104:C105,A106:C107,10)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 12.00187209);

		oParser = new parserFormula( "GROWTH(A103:C103,A104:C105,A106:C107,0)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1.0017632);
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue().toFixed(3) - 0, 12047829814.167);
		assert.strictEqual( oParser.calculate().getElementRowCol(0,2).getValue().toFixed(3) - 0, 10705900594.962);

		oParser = new parserFormula( "GROWTH({1,2,3},A104:C105,A106:C107,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1.00038318);

		oParser = new parserFormula( "GROWTH({1,2,3},A104:C105,A106:C107,A106:C107)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula( "GROWTH(A103:C103,A104:C105,A106:C107,1)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 12.00187209);
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue().toFixed(3) - 0, 676231620.297);
		assert.strictEqual( oParser.calculate().getElementRowCol(0,2).getValue().toFixed(3) - 0, 612512904.254)

		ws.getRange2( "A120" ).setValue( "1" );
		ws.getRange2( "B120" ).setValue( "2" );

		oParser = new parserFormula( "GROWTH(A120:B120)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1);

		oParser = new parserFormula( "GROWTH({1;2})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1);

		oParser = new parserFormula( "GROWTH({1,2})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1);
	} );


	QUnit.test("Test: \"LOGEST\"", function (assert) {

		ws.getRange2( "A101" ).setValue( "1" );
		ws.getRange2( "A102" ).setValue( "2" );
		ws.getRange2( "A103" ).setValue( "3" );
		ws.getRange2( "A104" ).setValue( "4" );
		ws.getRange2( "A105" ).setValue( "5" );
		ws.getRange2( "A106" ).setValue( "6" );
		ws.getRange2( "A107" ).setValue( "7" );
		ws.getRange2( "A108" ).setValue( "8" );
		ws.getRange2( "A109" ).setValue( "9" );
		ws.getRange2( "A110" ).setValue( "10" );
		ws.getRange2( "A111" ).setValue( "11" );
		ws.getRange2( "A112" ).setValue( "12" );


		ws.getRange2( "B101" ).setValue( "133890" );
		ws.getRange2( "B102" ).setValue( "135000" );
		ws.getRange2( "B103" ).setValue( "135790" );
		ws.getRange2( "B104" ).setValue( "137300" );
		ws.getRange2( "B105" ).setValue( "138130" );
		ws.getRange2( "B106" ).setValue( "139100" );
		ws.getRange2( "B107" ).setValue( "139900" );
		ws.getRange2( "B108" ).setValue( "141120" );
		ws.getRange2( "B109" ).setValue( "141890" );
		ws.getRange2( "B110" ).setValue( "143230" );
		ws.getRange2( "B111" ).setValue( "144000" );
		ws.getRange2( "B112" ).setValue( "145290" );

		ws.getRange2( "A115" ).setValue( "13" );
		ws.getRange2( "A116" ).setValue( "14" );
		ws.getRange2( "A117" ).setValue( "15" );
		ws.getRange2( "A118" ).setValue( "16" );
		ws.getRange2( "A119" ).setValue( "17" );

		oParser = new parserFormula( "LOGEST(B101:B112,A101:A112)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1.00732561);
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue().toFixed(4) - 0, 133044.8167);

		oParser = new parserFormula( "LOGEST(B101:B112,A101:A112,,false)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1.00732561);
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue().toFixed(4) - 0, 133044.8167);

		//todo необходимо перепроверить остальные значения в данном случае
		oParser = new parserFormula( "LOGEST(B101:B112,A101:A112,,true)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1.00732561);
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue().toFixed(4) - 0, 133044.8167);

		oParser = new parserFormula( "LOGEST(B101:B112,A101:A112,true,true)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1.00732561);
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue().toFixed(4) - 0, 133044.8167);

		//todo необходимо перепроверить остальные значения в данном случае
		oParser = new parserFormula( "LOGEST(B101:B112,A101:A112,false,true)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 4.15001464);
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue().toFixed(4) - 0, 1);

		oParser = new parserFormula( "LOGEST(A101:B105,A106:B110,FALSE,TRUE)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1.0000838);
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue().toFixed(4) - 0, 1);
		assert.strictEqual( oParser.calculate().getElementRowCol(1,0).getValue().toFixed(8) - 0, 0.00000264);
		assert.strictEqual( oParser.calculate().getElementRowCol(2,0).getValue().toFixed(4) - 0, 0.9911);
		assert.strictEqual( oParser.calculate().getElementRowCol(3,0).getValue().toFixed(4) - 0, 1005.3131);
		assert.strictEqual( oParser.calculate().getElementRowCol(4,0).getValue().toFixed(4) - 0, 698.5684);

		oParser = new parserFormula( "LOGEST(A101:B105,A106:B110,,)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1.00007701);
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue().toFixed(4) - 0, 2.6063);

		oParser = new parserFormula( "LOGEST(A101:B105,A106:B110,false,false)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1.0000838);
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue().toFixed(4) - 0, 1);

		//todo необходимо перепроверить остальные значения в данном случае
		oParser = new parserFormula( "LOGEST(A101:B105,A106:B110,true,true)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 1.00007701);
		assert.strictEqual( oParser.calculate().getElementRowCol(0,1).getValue().toFixed(4) - 0, 2.6063);
		assert.strictEqual( oParser.calculate().getElementRowCol(1,0).getValue().toFixed(8) - 0, 0.00000205);
		assert.strictEqual( oParser.calculate().getElementRowCol(2,0).getValue().toFixed(4) - 0, 0.9944);
		assert.strictEqual( oParser.calculate().getElementRowCol(3,0).getValue().toFixed(4) - 0, 1416.4887);
		assert.strictEqual( oParser.calculate().getElementRowCol(4,0).getValue().toFixed(4) - 0, 294.9627);

		ws.getRange2( "A120" ).setValue( "1" );
		ws.getRange2( "B120" ).setValue( "2" );

		oParser = new parserFormula( "LOGEST(A120:B120)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 2);

		oParser = new parserFormula( "LOGEST({1;2})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 2);

		oParser = new parserFormula( "LOGEST({1,2})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getElementRowCol(0,0).getValue().toFixed(8) - 0, 2);

	} );


	QUnit.test("Test: \"PDURATION\"", function (assert) {
		oParser = new parserFormula( "PDURATION(2.5%,2000,2200)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(2) - 0, 3.86);

		oParser = new parserFormula( "PDURATION(0.025/12,1000,1200)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(1) - 0, 87.6);

		oParser = new parserFormula( "PDURATION(0.025,1000,1200)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(2) - 0, 7.38);

		oParser = new parserFormula( "PDURATION(-0.025,1000,1200)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula( "PDURATION(0.025,-1000,1200)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula( "PDURATION(0.025,1000,-1200)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula( "PDURATION({0.025},{1000},{1200})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue().toFixed(2) - 0, 7.38);

		oParser = new parserFormula( "PDURATION(\"TEST\",1000,-1200)", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#VALUE!");

		testArrayFormula2(assert, "PDURATION", 3, 3);
	});

	QUnit.test("Test: \"IFS\"", function (assert) {

		oParser = new parserFormula( 'IFS(1,"TEST")', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TEST");

		oParser = new parserFormula( 'IFS(0,"TEST",1,"TEST2")', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TEST2");

		oParser = new parserFormula( 'IFS(2<1,">3")', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula( 'IFS(2<1,">3",2>1)', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula( 'IFS(2<1,"TEST",2<1,2,4>3,"TEST2")', "AA2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "TEST2");

		testArrayFormulaEqualsValues(assert, "1,3.123,-4,#N/A;2,4,5,#N/A;#N/A,#N/A,#N/A,#N/A","IFS(A1:C2,A1:C2,A1:C2,A1:C2, A1:C2,A1:C2)");
	});

	QUnit.test("Test: \"IF\"", function (assert) {

		oParser = new parserFormula('IF(1,"TEST")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TEST");

		oParser = new parserFormula('IF(0,"TEST")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		ws.getRange2( "A101" ).setValue( "1" );

		oParser = new parserFormula('IF(A101=1,"Yes","No")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Yes");

		oParser = new parserFormula('IF(A101=2,"Yes","No")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "No");

		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "IF", 2, 3);
	});

	QUnit.test("Test: \"COLUMN\"", function (assert) {

		oParser = new parserFormula('COLUMN(B6)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula('COLUMN(C16)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 3);

		oParser = new parserFormula('COLUMN()', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula('COLUMN()+COLUMN()', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		testArrayFormulaEqualsValues(assert, "5,6,7,8;5,6,7,8;5,6,7,8", "COLUMN()");
		testArrayFormulaEqualsValues(assert, "1,2,3,#N/A;1,2,3,#N/A;1,2,3,#N/A", "COLUMN(A1:C2)");
		testArrayFormulaEqualsValues(assert, "1,2,#N/A,#N/A;1,2,#N/A,#N/A;1,2,#N/A,#N/A", "COLUMN(A1:B1)");
		testArrayFormulaEqualsValues(assert, "1,1,1,1;1,1,1,1;1,1,1,1", "COLUMN(A1)");

	});

	QUnit.test("Test: \"COLUMNS\"", function (assert) {

		oParser = new parserFormula('COLUMNS(C1:E4)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 3);

		oParser = new parserFormula('COLUMNS({1,2,3;4,5,6})', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 3);

		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "COLUMNS", 1, 1);
	});

	QUnit.test("Test: \"ROW\"", function (assert) {

		oParser = new parserFormula('ROW(B6)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 6);

		oParser = new parserFormula('ROW(C16)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 16);

		oParser = new parserFormula('ROW()', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula('ROW()+ROW()', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		testArrayFormulaEqualsValues(assert, "6,6,6,6;7,7,7,7;8,8,8,8", "ROW()");
		testArrayFormulaEqualsValues(assert, "1,1,1,1;2,2,2,2;#N/A,#N/A,#N/A,#N/A", "ROW(A1:C2)");
		testArrayFormulaEqualsValues(assert, "1,1,1,1;1,1,1,1;1,1,1,1", "ROW(A1:B1)");
		testArrayFormulaEqualsValues(assert, "1,1,1,1;1,1,1,1;1,1,1,1", "ROW(A1)");
	});

	QUnit.test("Test: \"ROWS\"", function (assert) {

		oParser = new parserFormula('ROWS(C1:E4)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 4);

		oParser = new parserFormula('ROWS({1,2,3;4,5,6})', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "COLUMNS", 1, 1);
	});

	QUnit.test("Test: \"SUBTOTAL\"", function (assert) {
		ws.getRange2( "A102" ).setValue( "120" );
		ws.getRange2( "A103" ).setValue( "10" );
		ws.getRange2( "A104" ).setValue( "150" );
		ws.getRange2( "A105" ).setValue( "23" );

		oParser = new parserFormula( "SUBTOTAL(1,A102:A105)", "A2", ws );
		assert.ok( oParser.parse(), "SUBTOTAL(1,A102:A105)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(2) - 0, 75.75, "SUBTOTAL(1,A102:A105)");

		oParser = new parserFormula( "SUBTOTAL(2,A102:A105)", "A2", ws );
		assert.ok( oParser.parse(), "SUBTOTAL(2,A102:A105)" );
		assert.strictEqual( oParser.calculate().getValue(), 4, "SUBTOTAL(2,A102:A105)");

		oParser = new parserFormula( "SUBTOTAL(3,A102:A105)", "A2", ws );
		assert.ok( oParser.parse(), "SUBTOTAL(3,A102:A105)" );
		assert.strictEqual( oParser.calculate().getValue(), 4, "SUBTOTAL(3,A102:A105)");

		oParser = new parserFormula( "SUBTOTAL(4,A102:A105)", "A2", ws );
		assert.ok( oParser.parse(), "SUBTOTAL(4,A102:A105)" );
		assert.strictEqual( oParser.calculate().getValue(), 150, "SUBTOTAL(4,A102:A105)");

		oParser = new parserFormula( "SUBTOTAL(5,A102:A105)", "A2", ws );
		assert.ok( oParser.parse(), "SUBTOTAL(5,A102:A105)" );
		assert.strictEqual( oParser.calculate().getValue(), 10, "SUBTOTAL(5,A102:A105)");

		oParser = new parserFormula( "SUBTOTAL(6,A102:A105)", "A2", ws );
		assert.ok( oParser.parse(), "SUBTOTAL(6,A102:A105)" );
		assert.strictEqual( oParser.calculate().getValue(), 4140000, "SUBTOTAL(6,A102:A105)");

		oParser = new parserFormula( "SUBTOTAL(7,A102:A105)", "A2", ws );
		assert.ok( oParser.parse(), "SUBTOTAL(7,A102:A105)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 69.70592992, "SUBTOTAL(7,A102:A105)");

		oParser = new parserFormula( "SUBTOTAL(8,A102:A105)", "A2", ws );
		assert.ok( oParser.parse(), "SUBTOTAL(8,A102:A105)" );
		assert.strictEqual( oParser.calculate().getValue().toFixed(8) - 0, 60.36710611, "SUBTOTAL(8,A102:A105)");

		oParser = new parserFormula( "SUBTOTAL(9,A102:A105)", "A2", ws );
		assert.ok( oParser.parse(), "SUBTOTAL(9,A102:A105)" );
		assert.strictEqual( oParser.calculate().getValue(), 303, "SUBTOTAL(9,A102:A105)");
	} );

	QUnit.test("Test: \"MID\"", function (assert) {
		ws.getRange2( "A101" ).setValue( "Fluid Flow" );

		oParser = new parserFormula( "MID(A101,1,5)", "A2", ws );
		assert.ok( oParser.parse(), "MID(A101,1,5)" );
		assert.strictEqual( oParser.calculate().getValue(), "Fluid", "MID(A101,1,5)");

		oParser = new parserFormula( "MID(A101,7,20)", "A2", ws );
		assert.ok( oParser.parse(), "MID(A101,7,20)" );
		assert.strictEqual( oParser.calculate().getValue(), "Flow", "MID(A101,7,20)");

		oParser = new parserFormula( "MID(A101,20,5)", "A2", ws );
		assert.ok( oParser.parse(), "MID(A101,20,5)" );
		assert.strictEqual( oParser.calculate().getValue(), "", "MID(A101,20,5))");

		oParser = new parserFormula( "MID(TRUE,2,5)", "A2", ws );
		assert.ok( oParser.parse(), "MID(TRUE,2,5)" );
		assert.strictEqual( oParser.calculate().getValue(), "RUE", "MID(TRUE,2,5)");

		testArrayFormula2(assert, "MID", 3, 3);
	} );

	QUnit.test("Test: \"MIDB\"", function (assert) {
		ws.getRange2( "A101" ).setValue( "Fluid Flow" );

		oParser = new parserFormula( "MIDB(A101,1,5)", "A2", ws );
		assert.ok( oParser.parse(), "MIDB(A101,1,5)" );
		assert.strictEqual( oParser.calculate().getValue(), "Fluid", "MIDB(A101,1,5)");

		oParser = new parserFormula( "MIDB(A101,7,20)", "A2", ws );
		assert.ok( oParser.parse(), "MIDB(A101,7,20)" );
		assert.strictEqual( oParser.calculate().getValue(), "Flow", "MIDB(A101,7,20)");

		oParser = new parserFormula( "MIDB(A101,20,5)", "A2", ws );
		assert.ok( oParser.parse(), "MIDB(A101,20,5)" );
		assert.strictEqual( oParser.calculate().getValue(), "", "MIDB(A101,20,5))");

		oParser = new parserFormula( "MIDB(TRUE,2,5)", "A2", ws );
		assert.ok( oParser.parse(), "MIDB(TRUE,2,5)" );
		assert.strictEqual( oParser.calculate().getValue(), "RUE", "MIDB(TRUE,2,5)");
	} );

	QUnit.test("Test: \"MINUTE\"", function (assert) {

		ws.getRange2( "A202" ).setValue( "12:45:00 PM" );
		ws.getRange2( "A203" ).setValue( "7/18/2011 7:45" );
		ws.getRange2( "A204" ).setValue( "4/21/2012" );

		oParser = new parserFormula( "MINUTE(A202)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 45 );

		oParser = new parserFormula( "MINUTE(A203)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 45 );

		oParser = new parserFormula( "MINUTE(A204)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		ws.getRange2( "A205" ).setValue( "06/30/2020 20:00" );
		ws.getRange2( "A206" ).setValue( "06/30/2020 21:15" );

		ws.getRange2( "A207" ).setValue( "06/30/2020 23:15" );

		oParser = new parserFormula( "MINUTE(A206-A205)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 15 );

		oParser = new parserFormula( "MINUTE(A207-A205)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 15 );

		oParser = new parserFormula( "MINUTE(A207-A206)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );

		oParser = new parserFormula( "MINUTE(A207+A206)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 30 );

		oParser = new parserFormula( "MINUTE(123.1231231 - 1.12334343)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 59 );

		oParser = new parserFormula( "MINUTE(1.12334343 - 123.1231231)", "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "#NUM!" );
	} );

	/*QUnit.test("Test: \"MINVERSE\"", function (assert) {

		ws.getRange2( "A202" ).setValue( "4" );
		ws.getRange2( "A203" ).setValue( "2" );
		ws.getRange2( "B202" ).setValue( "-1" );
		ws.getRange2( "B203" ).setValue( "0" );

		oParser = new parserFormula( "MINVERSE({4,-1;2,0})", "A2", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), 0 );
	} );*/

	QUnit.test("Test: \"FIND\"", function (assert) {
		ws.getRange2( "A101" ).setValue( "Miriam McGovern" );

		oParser = new parserFormula( 'FIND("M",A101)', "A2", ws );
		assert.ok( oParser.parse(), 'FIND("M",A101)' );
		assert.strictEqual( oParser.calculate().getValue(), 1, 'FIND("M",A101)');

		oParser = new parserFormula( 'FIND("m",A101)', "A2", ws );
		assert.ok( oParser.parse(), 'FIND("m",A101)' );
		assert.strictEqual( oParser.calculate().getValue(), 6, 'FIND("m",A101)');

		oParser = new parserFormula( 'FIND("M",A101,3)', "A2", ws );
		assert.ok( oParser.parse(), 'FIND("M",A101,3)' );
		assert.strictEqual( oParser.calculate().getValue(), 8, 'FIND("M",A101,3)');

		oParser = new parserFormula( 'FIND("U",TRUE)', "A2", ws );
		assert.ok( oParser.parse(), 'FIND("T",TRUE)' );
		assert.strictEqual( oParser.calculate().getValue(), 3, 'FIND("T",TRUE)');

		testArrayFormula2(assert, "FIND", 2, 3);
	} );

	QUnit.test("Test: \"FINDB\"", function (assert) {
		ws.getRange2( "A101" ).setValue( "Miriam McGovern" );

		oParser = new parserFormula( 'FINDB("M",A101)', "A2", ws );
		assert.ok( oParser.parse(), 'FINDB("M",A101)' );
		assert.strictEqual( oParser.calculate().getValue(), 1, 'FINDB("M",A101)');

		oParser = new parserFormula( 'FINDB("m",A101)', "A2", ws );
		assert.ok( oParser.parse(), 'FINDB("m",A101)' );
		assert.strictEqual( oParser.calculate().getValue(), 6, 'FINDB("m",A101)');

		oParser = new parserFormula( 'FINDB("M",A101,3)', "A2", ws );
		assert.ok( oParser.parse(), 'FINDB("M",A101,3)' );
		assert.strictEqual( oParser.calculate().getValue(), 8, 'FINDB("M",A101,3)');

		oParser = new parserFormula( 'FINDB("U",TRUE)', "A2", ws );
		assert.ok( oParser.parse(), 'FINDB("T",TRUE)' );
		assert.strictEqual( oParser.calculate().getValue(), 3, 'FINDB("T",TRUE)');
	} );

	QUnit.test("Test: \">\"", function (assert) {
		oParser = new parserFormula( '1.123>1.5', "A2", ws );
		assert.ok( oParser.parse(), '1.123>1.5' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", '1.123>1.5');

		oParser = new parserFormula( '1.555>1.5', "A2", ws );
		assert.ok( oParser.parse(), '1.555>1.5' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", '1.555>1.5');
	} );

	QUnit.test("Test: \"<\"", function (assert) {
		oParser = new parserFormula( '1.123<1.5', "A2", ws );
		assert.ok( oParser.parse(), '1.123<1.5' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", '1.123<1.5');

		oParser = new parserFormula( '1.555<1.5', "A2", ws );
		assert.ok( oParser.parse(), '1.555<1.5' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", '1.555<1.5');
	} );

	QUnit.test("Test: \"=\"", function (assert) {
		oParser = new parserFormula( '1.123=1.5', "A2", ws );
		assert.ok( oParser.parse(), '1.123=1.5' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", '1.123=1.5');

		oParser = new parserFormula( '1.555=1.555', "A2", ws );
		assert.ok( oParser.parse(), '1.555=1.555' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", '1.555=1.555');
	} );

	QUnit.test("Test: \"<>\"", function (assert) {
		oParser = new parserFormula( '1.123<>1.5', "A2", ws );
		assert.ok( oParser.parse(), '1.123<>1.5' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", '1.123<>1.5');

		oParser = new parserFormula( '1.555<>1.555', "A2", ws );
		assert.ok( oParser.parse(), '1.555<>1.555' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", '1.555<>1.555');
	} );

	QUnit.test("Test: \">=\"", function (assert) {
		oParser = new parserFormula( '1.123>=1.5', "A2", ws );
		assert.ok( oParser.parse(), '1.123>=1.5' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", '1.123>=1.5');

		oParser = new parserFormula( '1.555>=1.555', "A2", ws );
		assert.ok( oParser.parse(), '1.555>=1.555' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", '1.555>=1.555');

		oParser = new parserFormula( '1.557>=1.555', "A2", ws );
		assert.ok( oParser.parse(), '1.557>=1.555' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", '1.557>=1.555');
	} );

	QUnit.test("Test: \"<=\"", function (assert) {
		oParser = new parserFormula( '1.123<=1.5', "A2", ws );
		assert.ok( oParser.parse(), '1.123<=1.5' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", '1.123<=1.5');

		oParser = new parserFormula( '1.555<=1.555', "A2", ws );
		assert.ok( oParser.parse(), '1.555<=1.555' );
		assert.strictEqual( oParser.calculate().getValue(), "TRUE", '1.555<=1.555');

		oParser = new parserFormula( '1.557<=1.555', "A2", ws );
		assert.ok( oParser.parse(), '1.557<=1.555' );
		assert.strictEqual( oParser.calculate().getValue(), "FALSE", '1.557<=1.555');
	} );


	QUnit.test("Test: \"ADDRESS\"", function (assert) {

		oParser = new parserFormula( "ADDRESS(2,3,2)", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(2,3,2)" );
		assert.strictEqual( oParser.calculate().getValue(), "C$2", "ADDRESS(2,3,2)");

		oParser = new parserFormula( "ADDRESS(2,3,2,FALSE)", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(2,3,2,FALSE)" );
		assert.strictEqual( oParser.calculate().getValue(), "R2C[3]", "ADDRESS(2,3,2,FALSE)");

		oParser = new parserFormula( 'ADDRESS(2,3,1,FALSE,"[Book1]Sheet1")', "A2", ws );
		assert.ok( oParser.parse(), 'ADDRESS(2,3,1,FALSE,"[Book1]Sheet1")' );
		assert.strictEqual( oParser.calculate().getValue(), "'[Book1]Sheet1'!R2C3", 'ADDRESS(2,3,1,FALSE,"[Book1]Sheet1")');

		oParser = new parserFormula( 'ADDRESS(2,3,1,FALSE,"EXCEL SHEET")', "A2", ws );
		assert.ok( oParser.parse(), 'ADDRESS(2,3,1,FALSE,"EXCEL SHEET")' );
		assert.strictEqual( oParser.calculate().getValue(), "'EXCEL SHEET'!R2C3", 'ADDRESS(2,3,1,FALSE,"EXCEL SHEET")');

		ws.getRange2( "A101" ).setValue( "" );

		oParser = new parserFormula( "ADDRESS(2,3,2,1,A101)", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(2,3,2,1,A101" );
		assert.strictEqual( oParser.calculate().getValue(), "!C$2", "ADDRESS(2,3,2,1,A101");

		ws.getRange2( "A101" ).setValue( "'" );

		oParser = new parserFormula( "ADDRESS(2,3,2,1,A101)", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(2,3,2,1,A101" );
		assert.strictEqual( oParser.calculate().getValue(), "!C$2", "ADDRESS(2,3,2,1,A101");

		oParser = new parserFormula( 'ADDRESS(2,3,2,1,"")', "A2", ws );
		assert.ok( oParser.parse(), 'ADDRESS(2,3,2,1,"")' );
		assert.strictEqual( oParser.calculate().getValue(), "!C$2", 'ADDRESS(2,3,2,1,"")');

		oParser = new parserFormula( "ADDRESS(2,3,2,1,\"'\")", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(2,3,2,1,\"'\")" );
		assert.strictEqual( oParser.calculate().getValue(), "''''!C$2", "ADDRESS(2,3,2,1,\"'\")");

		oParser = new parserFormula( "ADDRESS(2,3,,,1)", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(2,3,,,1)" );
		assert.strictEqual( oParser.calculate().getValue(), "'1'!$C$2", "ADDRESS(2,3,,,1)");

		oParser = new parserFormula( "ADDRESS(2,3,1,,1)", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(2,3,1,,1)" );
		assert.strictEqual( oParser.calculate().getValue(), "'1'!$C$2", "ADDRESS(2,3,1,,1)");

		oParser = new parserFormula( "ADDRESS(2,3,2,,1)", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(2,3,2,,1)" );
		assert.strictEqual( oParser.calculate().getValue(), "'1'!C$2", "ADDRESS(2,3,2,,1)");

		oParser = new parserFormula( "ADDRESS(2,3,,TRUE,1)", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(2,3,,TRUE,1)" );
		assert.strictEqual( oParser.calculate().getValue(), "'1'!$C$2", "ADDRESS(2,3,,TRUE,1)");

		oParser = new parserFormula( "ADDRESS(2,3,,FALSE,1)", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(2,3,,FALSE,1)" );
		assert.strictEqual( oParser.calculate().getValue(), "'1'!R2C3", "ADDRESS(2,3,,FALSE,1)");

		oParser = new parserFormula( "ADDRESS(2,3,,FALSE,1)", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(2,3,,FALSE,1)" );
		assert.strictEqual( oParser.calculate().getValue(), "'1'!R2C3", "ADDRESS(2,3,,FALSE,1)");

		oParser = new parserFormula( "ADDRESS(1,7,,)", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(1,7,,)" );
		assert.strictEqual( oParser.calculate().getValue(), "$G$1", "ADDRESS(1,7,,)");

		oParser = new parserFormula( "ADDRESS(1,7,,,)", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(1,7,,,)" );
		assert.strictEqual( oParser.calculate().getValue(), "$G$1", "ADDRESS(1,7,,,)");

		oParser = new parserFormula( "ADDRESS(2.123,3.3213,2)", "A2", ws );
		assert.ok( oParser.parse(), "ADDRESS(2.123,3.3213,2)" );
		assert.strictEqual( oParser.calculate().getValue(), "C$2", "ADDRESS(2.123,3.3213,2)");

		testArrayFormula2(assert, "ADDRESS", 2, 5);
	} );

	QUnit.test("Test: \"reference argument test\"", function (assert) {
		ws.getRange2( "A1" ).setValue( "1" );
		ws.getRange2( "A2" ).setValue( "2" );
		ws.getRange2( "A3" ).setValue( "3" );
		ws.getRange2( "A4" ).setValue( "4" );
		ws.getRange2( "A5" ).setValue( "5" );
		ws.getRange2( "A6" ).setValue( "6" );

		ws.getRange2( "B1" ).setValue( "2" );
		ws.getRange2( "B2" ).setValue( "" );
		ws.getRange2( "B3" ).setValue( "3" );
		ws.getRange2( "B4" ).setValue( "4" );
		ws.getRange2( "B5" ).setValue( "5" );
		ws.getRange2( "B6" ).setValue( "6" );

		oParser = new parserFormula( 'IRR(SIN(A1:B4))', 'A2', ws );
		assert.ok( oParser.parse(),'IRR(SIN(A1:B4))' );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, -0.123554096,'IRR(SIN(A1:B4))');

		oParser = new parserFormula( 'MIRR(SIN(A2:B4),1,1)', 'A2', ws );
		assert.ok( oParser.parse(),'MIRR(SIN(A2:B4),1,1)' );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0, 2.36894463,'MIRR(SIN(A2:B4),1,1)');

		oParser = new parserFormula( 'COLUMN(INDEX(A1:B3,1,1))', 'A2', ws );
		assert.ok( oParser.parse(),'COLUMN(INDEX(A1:B3,1,1))' );
		assert.strictEqual( oParser.calculate().getValue(),1,'COLUMN(INDEX(A1:B3,1,1))');

		oParser = new parserFormula( 'COLUMNS(SIN($A$1:$B$4))', 'A2', ws );
		assert.ok( oParser.parse(),'COLUMNS(SIN($A$1:$B$4))' );
		assert.strictEqual( oParser.calculate().getValue(),2,'COLUMNS(SIN($A$1:$B$4))');

		oParser = new parserFormula( 'INDEX(SIN(A1:B3),1,1)', 'A2', ws );
		assert.ok( oParser.parse(),'INDEX(SIN(A1:B3),1,1)' );
		assert.strictEqual( oParser.calculate().getValue().toFixed(9) - 0,0.841470985,'INDEX(SIN(A1:B3),1,1)');

		/*oParser = new parserFormula( 'OFFSET(INDEX(A1:B3,1,1),1,1)', 'A2', ws );
		assert.ok( oParser.parse(),'OFFSET(INDEX(A1:B3,1,1),1,1)' );
		assert.strictEqual( oParser.calculate().getValue(),0,'OFFSET(INDEX(A1:B3,1,1),1,1)');*/

		oParser = new parserFormula( 'ROW(INDEX(A1:B3,1,1))', 'A2', ws );
		assert.ok( oParser.parse(),'ROW(INDEX(A1:B3,1,1))' );
		assert.strictEqual( oParser.calculate().getValue(),1,'ROW(INDEX(A1:B3,1,1))');

		oParser = new parserFormula( 'ROWS(SIN(A1:B3))', 'A2', ws );
		assert.ok( oParser.parse(),'ROWS(SIN(A1:B3))' );
		assert.strictEqual( oParser.calculate().getValue(),3,'ROWS(SIN(A1:B3))');

		oParser = new parserFormula( 'SUBTOTAL(1,INDEX(A1:B3,1,1))', 'A2', ws );
		assert.ok( oParser.parse(),'SUBTOTAL(1,INDEX(A1:B3,1,1))' );
		assert.strictEqual( oParser.calculate().getValue(),1,'SUBTOTAL(1,INDEX(A1:B3,1,1))');

		oParser = new parserFormula( 'SUMIF(INDEX(A1:B3,1,1),1,INDEX(A1:B3,1,1))', 'A2', ws );
		assert.ok( oParser.parse(),'SUMIF(INDEX(A1:B3,1,1),1,INDEX(A1:B3,1,1))' );
		assert.strictEqual( oParser.calculate().getValue(),1,'SUMIF(INDEX(A1:B3,1,1),1,INDEX(A1:B3,1,1))');

		oParser = new parserFormula( 'SUMIFS(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))', 'A2', ws );
		assert.ok( oParser.parse(),'SUMIFS(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))' );
		assert.strictEqual( oParser.calculate().getValue(),1,'SUMIFS(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))');

		oParser = new parserFormula( 'AVERAGEIF(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))', 'A2', ws );
		assert.ok( oParser.parse(),'AVERAGEIF(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))' );
		assert.strictEqual( oParser.calculate().getValue(),1,'AVERAGEIF(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))');

		oParser = new parserFormula( 'COUNTBLANK(INDEX(A1:B3,1,1))', 'A2', ws );
		assert.ok( oParser.parse(),'COUNTBLANK(INDEX(A1:B3,1,1))' );
		assert.strictEqual( oParser.calculate().getValue(),0,'COUNTBLANK(INDEX(A1:B3,1,1))');

		oParser = new parserFormula( 'COUNTIF(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))', 'A2', ws );
		assert.ok( oParser.parse(),'COUNTIF(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))' );
		assert.strictEqual( oParser.calculate().getValue(),1,'COUNTIF(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))');

		oParser = new parserFormula( 'COUNTIFS(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))', 'A2', ws );
		assert.ok( oParser.parse(),'COUNTIFS(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))' );
		assert.strictEqual( oParser.calculate().getValue(),1,'COUNTIFS(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))');

		ws.getRange2( "A2" ).setValue( "qq" );
		ws.getRange2( "A3" ).setValue( "ww" );
		ws.getRange2( "A4" ).setValue( "ee" );
		ws.getRange2( "A5" ).setValue( "qq" );
		ws.getRange2( "A6" ).setValue( "qq" );
		ws.getRange2( "A7" ).setValue( "ww" );
		ws.getRange2( "A8" ).setValue( "ww" );
		ws.getRange2( "A9" ).setValue( "ww" );
		ws.getRange2( "A10" ).setValue( "eee" );

		ws.getRange2( "B1" ).setValue( "qqqq" );
		ws.getRange2( "B2" ).setValue( "ee" );

		var _f = 'IFERROR(INDEX($A$2:$A$10,MATCH(0,INDEX(COUNTIF($B$1:B1,$A$2:$A$10)+(COUNTIF($A$2:$A$10,$A$2:$A$10)<>1),0,0),0)),"")';
		oParser = new parserFormula( _f, 'A2', ws );
		assert.ok( oParser.parse(), _f );
		assert.strictEqual( oParser.calculate().getValue().getValue(),"ee",_f);

		_f = 'IFERROR(INDEX($A$2:$A$10,MATCH(0,INDEX(COUNTIF($B$1:B2,$A$2:$A$10)+(COUNTIF($A$2:$A$10,$A$2:$A$10)<>1),0,0),0)),"")';
		oParser = new parserFormula( _f, 'A2', ws );
		assert.ok( oParser.parse(), _f );
		assert.strictEqual( oParser.calculate().getValue().getValue(),"eee",_f);

		_f = 'INDEX($A$2:$A$10,MATCH(0,INDEX(COUNTIF($B$1:B1,$A$2:$A$10)+(COUNTIF($A$2:$A$10,$A$2:$A$10)<>1),0,0),0))';
		oParser = new parserFormula( _f, 'A2', ws );
		assert.ok( oParser.parse(), _f );
		assert.strictEqual( oParser.calculate().getValue().getValue(),"ee",_f);

		_f = 'MATCH(0,INDEX({1;1;0;1;1;1;1;1;0},0,0))';
		oParser = new parserFormula( _f, 'A2', ws );
		assert.ok( oParser.parse(), _f );
		assert.strictEqual( oParser.calculate().getValue(),"#N/A",_f);

		_f = 'INDEX($A$2:$A$10,MATCH(0,INDEX({1;1;0;1;1;1;1;1;0},0,0),0))';
		oParser = new parserFormula( _f, 'A2', ws );
		assert.ok( oParser.parse(), _f );
		assert.strictEqual( oParser.calculate().getValue().getValue(),"ee",_f);

		_f = 'INDEX($A$2:$A$10,3)';
		oParser = new parserFormula( _f, 'A2', ws );
		assert.ok( oParser.parse(), _f );
		assert.strictEqual( oParser.calculate().getValue().getValue(),"ee",_f);

		_f = 'INDEX($A$2:$A$10,MATCH(0,{1;1;0;1;1;1;1;1;0},0))';
		oParser = new parserFormula( _f, 'A2', ws );
		assert.ok( oParser.parse(), _f );
		assert.strictEqual( oParser.calculate().getValue().getValue(),"ee",_f);

		_f = 'MATCH(0,INDEX(COUNTIF($B$1:B1,$A$2:$A$10)+(COUNTIF($A$2:$A$10,$A$2:$A$10)<>1),0,0),0)';
		oParser = new parserFormula( _f, 'A2', ws );
		assert.ok( oParser.parse(), _f );
		assert.strictEqual( oParser.calculate().getValue(),3,_f);


	} );

	wb.dependencyFormulas.unlockRecal();
} );

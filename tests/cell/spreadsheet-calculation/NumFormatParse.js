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

$(function () {
    window["AscCommonExcel"] = window["AscCommonExcel"] || {};
    window["AscCommonExcel"].Font = function () {
    };
    window["AscCommonExcel"].RgbColor = function () {
    };

    QUnit.module('NumFomat parse');
    let eps = 1e-15;
    QUnit.test('parseDate', function (assert) {
        let data = [
            ["1/2/2000 11:34:56", "m/d/yyyy h:mm", 36527.482592592591],
            ["1/2/2000 11:34:5", "m/d/yyyy h:mm", 36527.482002314813],
            ["1/2/2000 11:34:", "m/d/yyyy h:mm", 36527.481944444444],
            ["1/2/2000 11:34", "m/d/yyyy h:mm", 36527.481944444444],
            ["1/2/2000 11:3", "m/d/yyyy h:mm", 36527.460416666669],
            ["1/2/2000 11:", "m/d/yyyy h:mm", 36527.458333333336],
            ["11:34:56", "h:mm:ss", 0.48259259259259263],
            ["11:34:5", "h:mm:ss", 0.48200231481481487],
            ["11:34:", "h:mm", 0.48194444444444445],
            ["11:34", "h:mm", 0.48194444444444445],
            ["11:3", "h:mm", 0.4604166666666667],
            ["11:", "h:mm", 0.45833333333333331],
            ["1/2/2000 11:34:56 AM", "m/d/yyyy h:mm", 36527.482592592591],
            ["1/2/2000 11:34:5 AM", "m/d/yyyy h:mm", 36527.482002314813],
            ["1/2/2000 11:34: AM", "m/d/yyyy h:mm", 36527.481944444444],
            ["1/2/2000 11:34 AM", "m/d/yyyy h:mm", 36527.481944444444],
            ["1/2/2000 11:3 AM", "m/d/yyyy h:mm", 36527.460416666669],
            ["1/2/2000 11: AM", "m/d/yyyy h:mm", 36527.458333333336],
            ["11:34:56 AM", "h:mm:ss AM/PM", 0.48259259259259263],
            ["11:34:5 AM", "h:mm:ss AM/PM", 0.48200231481481487],
            ["11:34: AM", "h:mm AM/PM", 0.48194444444444445],
            ["11:34 AM", "h:mm AM/PM", 0.48194444444444445],
            ["11:3 AM", "h:mm AM/PM", 0.4604166666666667],
            ["11: AM", "h:mm AM/PM", 0.45833333333333331],
            ["11:00:00", "h:mm:ss", 0.45833333333333331],
            ["11:00:0", "h:mm:ss", 0.45833333333333331],
            ["11:00:", "h:mm", 0.45833333333333331],
            ["11:0", "h:mm", 0.45833333333333331],
            ["11:", "h:mm", 0.45833333333333331],
            ["1/2/2000 55:34:56", "General", 36529.315925925926],
            ["1/2/2000 55:34:5", "General", 36529.315335648149],
            ["1/2/2000 55:34:", "General", 36529.31527777778],
            ["1/2/2000 55:34", "General", 36529.31527777778],
            ["1/2/2000 55:3", "General", 36529.293749999997],
            ["1/2/2000 55:", "General", 36529.291666666664],
            ["55:34:56", "[h]:mm:ss", 2.3159259259259257],
            ["55:34:5", "[h]:mm:ss", 2.3153356481481482],
            ["55:34:", "[h]:mm:ss", 2.3152777777777778],
            ["55:34", "[h]:mm:ss", 2.3152777777777778],
            ["55:3", "[h]:mm:ss", 2.2937499999999997],
            ["55:", "[h]:mm:ss", 2.2916666666666665],
        ]
        for(let i = 0; i < data.length; i++){
            let date = AscCommon.g_oFormatParser.parse(data[i][0]);
            assert.strictEqual(date.format, data[i][1], `Case format: ${data[i][0]}`);
            assert.strictEqual(Math.abs(date.value - data[i][2]) < eps, true, `Case value: ${data[i][0]}`);
        }
    });

    QUnit.test('formatNumber', function (assert) {
        let testCases = [
            // Thousand separators
            [1234, '#,##0', '1,234'],
            [1234567, '#,##0', '1,234,567'],
            [0, '#,##0', '0'],
            [-1234, '#,##0', '-1,234'],
            
            // Decimal places
            [1234.56, '#,##0.00', '1,234.56'],
            [1234.5, '#,##0.00', '1,234.50'],
            [0.5, '0.00', '0.50'],
            [1.234, '0.00', '1.23'],
            
            // Percentages
            [0.5, '0%', '50%'],
            [0.125, '0.00%', '12.50%'],
            [1, '0%', '100%'],
            [0.999, '0%', '100%'],
            
            // Currency with text literals
            [1234.56, '"$"#,##0.00', '$1,234.56'],
            [0, '"$"#,##0.00', '$0.00'],
            [-50, '"$"#,##0.00', '-$50.00'],
            [1000, '"USD "0.00', 'USD 1000.00'],
            
            // Negative numbers in parentheses
            [100, '0;(0)', '100'],
            [-100, '0;(0)', '(100)'],
            [0, '0;(0)', '0'],
            [-50.5, '0.00;(0.00)', '(50.50)'],
            
            // Optional digits with #
            [123, '###', '123'],
            [0, '###', ''],
            [12.3, '##.#', '12.3'],
            [12, '##.#', '12.'],
            
            // Mandatory zeros
            [5, '000', '005'],
            [123, '000', '123'],
            [5.5, '000.00', '005.50'],
            [0, '00', '00'],
            
            // Space alignment with ?
            [1, '??', '01'],
            [10, '??', '10'],
            [1.5, '?.??', '1.50'],
            [10.25, '?.??', '10.25'],
            
            // Escaped characters
            [100, '\\#0', '#100'],
            [50, '0\\%', '50%'],
            [10, '0\\-', '10-'],
            [25, '\\+0', '+25'],
            
            // Mixed format
            [1234.5, '#,##0.00;[Red](#,##0.00)', '1,234.50'],
            [-1234.5, '#,##0.00;[Red](#,##0.00)', '(1,234.50)'],
            
            // Additional important cases
            [0.75, '0.#', '0.8'],
            [100.123, '0.0', '100.1'],
            [1234, '"Total: "#,##0', 'Total: 1,234'],
            [0.5555, '0.00%', '55.55%'],
            [999999, '#,##0', '999,999'],
            [-0.25, '0.00;(0.00)', '(0.25)'],

            // Date format cases
            [0.684027777777778, 'mm', '01'],
            [0.684027777777778, '[mm]', '985'],
            [0.684027777777778, '[h] "hours"', '16 hours'],
            [0.684027777777778, '[h]:mm', '16:25'],
            [0.684027777777778, '[h]:mm" ""minutes"', '16:25 minutes'],
            [0.684027777777778, '[s]', '59100'],
            [0.684027777777778, '[s]" ""seconds"', '59100 seconds'],
            [0.684027777777778, '[ss].0', '59100.0'],
            [0.684027777777778, '[mm]:ss', '985:00'],
            [0.684027777777778, '[mm]:mm', '985:01'],
            [0.684027777777778, '[hh]', '16'],
            [0.684027777777778, '[h]:mm:ss.000', '16:25:00.000'],
            [0.684027777777778, 'dd"d "hh"h "mm"m "ss"s"" "AM/PM', '00d 04h 25m 00s PM'],
            [0.684027777777778, '[h]"h*"mm"m*"ss"s*"ss"ms"', '16h*25m*00s*00ms'],
            [0.684027777777778, 'yyyy"Y-"mm"M-"dd"D "hh"H:"mm"M:"ss"."s"S"" "AM/PM', '1900Y-01M-00D 04H:25M:00.0S PM'],
            [0.684027777777778, 'dd:mm:yyyy" "hh:mm:ss" "[hh]:[mm]" "AM/PM" ""minutes AM/PM"', '00:01:1900 04:25:00 04:985 PM minutes AM/PM'],

            [37753.6844097222, 'mm', '05'],
            [37753.6844097222, '[mm]', '54365305'],
            [37753.6844097222, '[h] "hours"', '906088 hours'],
            [37753.6844097222, '[h]:mm', '906088:25'],
            [37753.6844097222, '[h]:mm" ""minutes"', '906088:25 minutes'],
            [37753.6844097222, '[s]', '3261918333'],
            [37753.6844097222, '[s]" ""seconds"', '3261918333 seconds'],
            [37753.6844097222, '[ss].0', '3261918333.0'],
            [37753.6844097222, '[mm]:ss', '54365305:33'],
            [37753.6844097222, '[mm]:mm', '54365305:05'],
            [37753.6844097222, '[hh]', '906088'],
            [37753.6844097222, '[h]:mm:ss.000', '906088:25:33.000'],
            [37753.6844097222, 'dd"d "hh"h "mm"m "ss"s"" "AM/PM', '12d 04h 25m 33s PM'],
            [37753.6844097222, '[h]"h*"mm"m*"ss"s*"ss"ms"', '906088h*25m*33s*33ms'],
            [37753.6844097222, 'yyyy"Y-"mm"M-"dd"D "hh"H:"mm"M:"ss"."s"S"" "AM/PM', '2003Y-05M-12D 04H:25M:33.33S PM'],
            [37753.6844097222, 'dd:mm:yyyy" "hh:mm:ss" "[hh]:[mm]" "AM/PM" ""minutes AM/PM"', '12:05:2003 04:25:33 04:54365305 PM minutes AM/PM'],
        ];
        
        for (let i = 0; i < testCases.length; i++) {
            let value = testCases[i][0];
            let format = testCases[i][1];
            let expected = testCases[i][2];
            
            let expr = new AscCommon.CellFormat(format);
            let formatted = expr.format(value);
            let text = '';
            for (let j = 0, length = formatted.length; j < length; ++j) {
                text += formatted[j].text;
            }
            
            assert.strictEqual(text, expected, `format("${format}", ${value})`);
        }
    });


    QUnit.test('parseDate', function (assert) {
        // Same date and time separators
        const fnCultureInfo = {LCID: 11, Name: "fi", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["sunnuntai", "maanantai", "tiistai", "keskiviikko", "torstai", "perjantai", "lauantai"], AbbreviatedDayNames: ["su", "ma", "ti", "ke", "to", "pe", "la"], MonthNames: ["tammikuu", "helmikuu", "maaliskuu", "huhtikuu", "toukokuu", "kesäkuu", "heinäkuu", "elokuu", "syyskuu", "lokakuu", "marraskuu", "joulukuu", ""], AbbreviatedMonthNames: ["tammi", "helmi", "maalis", "huhti", "touko", "kesä", "heinä", "elo", "syys", "loka", "marras", "joulu", ""], MonthGenitiveNames: ["tammikuuta", "helmikuuta", "maaliskuuta", "huhtikuuta", "toukokuuta", "kesäkuuta", "heinäkuuta", "elokuuta", "syyskuuta", "lokakuuta", "marraskuuta", "joulukuuta", ""], AbbreviatedMonthGenitiveNames: ["tammik.", "helmik.", "maalisk.", "huhtik.", "toukok.", "kesäk.", "heinäk.", "elok.", "syysk.", "lokak.", "marrask.", "jouluk.", ""], AMDesignator: "ap.", PMDesignator: "ip.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ".", ShortDatePattern: "025", LongDatePattern: "dddd\\ d\\.\\ mmmm\\ yyyy"};
        // Short date pattern 205 (american)
        const enCultureInfo = {LCID: 9, Name: "en", CurrencyPositivePattern: 0, CurrencyNegativePattern: 0, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "205", LongDatePattern: "dddd\\,\\ mmmm\\ d\\,\\ yyyy"};
        // Short date pattern 135 (european)
        const ruCultureInfo = {LCID: 1049, Name: "ru-RU", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₽", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["воскресенье", "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"], AbbreviatedDayNames: ["Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"], MonthNames: ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь", ""], AbbreviatedMonthNames: ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек", ""], MonthGenitiveNames: ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря", ""], AbbreviatedMonthGenitiveNames: ["янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек", ""], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\\ \"г.\""};
	    // Short date pattern 531 (asian)
        const jpCultureInfo = {LCID: 17, Name: "ja", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"], AbbreviatedDayNames: ["日", "月", "火", "水", "木", "金", "土"], MonthNames: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", ""], AbbreviatedMonthNames: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "午前", PMDesignator: "午後", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""};

        let testCases = [

            //SAME DATE SEPARATORS

            // Date values
            // 2 - values

            // . separator
            ['0.0', null, null, fnCultureInfo],
            ['0.6', null, null, fnCultureInfo],
            ['6.0', null, null, fnCultureInfo],
            ['12.0', null, null, fnCultureInfo],
            ['20.0', null, null, fnCultureInfo],
            ['13.15', null, null, fnCultureInfo],
            ['15.13', null, null, fnCultureInfo],
            ['0.9999', null, null, fnCultureInfo],
            ['0.10000', null, null, fnCultureInfo],
            
            ['6.15', 'mmm-yy', 42156, fnCultureInfo],
            ['15.6', 'd-mmm', 45823, fnCultureInfo],
            ['15.12', 'd-mmm', 46006, fnCultureInfo],
            ['31.12', 'd-mmm', 46022, fnCultureInfo],
            ['6.31', 'mmm-yy', 11475, fnCultureInfo],
            ['6.9999', 'mmm-yy', 2958252, fnCultureInfo],
            ['12.50', 'mmm-yy', 18598, fnCultureInfo],

            // / separator
            ['0/0', null, null, fnCultureInfo],
            ['0/6', null, null, fnCultureInfo],
            ['6/0', null, null, fnCultureInfo],
            ['12/0', null, null, fnCultureInfo],
            ['20/0', null, null, fnCultureInfo],
            ['13/15', null, null, fnCultureInfo],
            ['15/13', null, null, fnCultureInfo],
            ['0/9999', null, null, fnCultureInfo],
            ['0/10000', null, null, fnCultureInfo],
            
            ['6/15', 'mmm-yy', 42156, fnCultureInfo],
            ['15/6', 'd-mmm', 45823, fnCultureInfo],
            ['15/12', 'd-mmm', 46006, fnCultureInfo],
            ['31/12', 'd-mmm', 46022, fnCultureInfo],
            ['6/31', 'mmm-yy', 11475, fnCultureInfo],
            ['6/9999', 'mmm-yy', 2958252, fnCultureInfo],
            ['12/50', 'mmm-yy', 18598, fnCultureInfo],

            // 3 - values

            // . separator
            ['0.0.0', null, null, fnCultureInfo],
            ['0.0.2000', null, null, fnCultureInfo],
            ['0.6.2000', null, null, fnCultureInfo],
            ['0.13.2000', null, null, fnCultureInfo],
            ['6.0.2000', null, null, fnCultureInfo],
            ['13.0.2000', null, null, fnCultureInfo],
            ['01.01.0000', null, null, fnCultureInfo],


            ['01.01.2000', 'd/m/yyyy', 36526, fnCultureInfo],
            ['15.6.2000', 'd/m/yyyy', 36692, fnCultureInfo],
            ['15.12.2000', 'd/m/yyyy', 36875, fnCultureInfo],
            ['31.12.9999', 'd/m/yyyy', 2958465, fnCultureInfo],
            
            ['6.15.2000', null, null, fnCultureInfo],
            ['12.15.2000', null, null, fnCultureInfo],
            ['50.12.2000', null, null, fnCultureInfo],
            ['15.6.10000', null, null, fnCultureInfo],
            ['6.15.10000', null, null, fnCultureInfo],

            // / separator
            ['0/0/0', null, null, fnCultureInfo],
            ['0/0/2000', null, null, fnCultureInfo],
            ['0/6/2000', null, null, fnCultureInfo],
            ['0/13/2000', null, null, fnCultureInfo],
            ['6/0/2000', null, null, fnCultureInfo],
            ['13/0/2000', null, null, fnCultureInfo],
            ['01/01/0000', null, null, fnCultureInfo],


            ['01/01/2000', 'd/m/yyyy', 36526, fnCultureInfo],
            ['15/6/2000', 'd/m/yyyy', 36692, fnCultureInfo],
            ['15/12/2000', 'd/m/yyyy', 36875, fnCultureInfo],
            ['31/12/9999', 'd/m/yyyy', 2958465, fnCultureInfo],
            
            ['6/15/2000', null, null, fnCultureInfo],
            ['12/15/2000', null, null, fnCultureInfo],
            ['50/12/2000', null, null, fnCultureInfo],
            ['15/6/10000', null, null, fnCultureInfo],
            ['6/15/10000', null, null, fnCultureInfo],


            // Time values
            // 2 - values
            ['0:0', 'h:mm', 0, fnCultureInfo],
            ['0:30', 'h:mm', 0.020833333333333332, fnCultureInfo],
            ['0:60', 'General', 0.041666666666666664, fnCultureInfo],
            ['0:120', 'General', 0.08333333333333333, fnCultureInfo],
            ['0:9999', 'General', 6.94375, fnCultureInfo],
            ['0:10000', null, null, fnCultureInfo],
            ['12:0', 'h:mm', 0.5, fnCultureInfo],
            ['12:30', 'h:mm', 0.5208333333333334, fnCultureInfo],
            ['12:60', 'General', 0.5416666666666666, fnCultureInfo],
            ['12:120', 'General', 0.5833333333333334, fnCultureInfo],
            ['24:0', 'h:mm', 1, fnCultureInfo],
            ['24:30', '[h]:mm:ss', 1.0208333333333333, fnCultureInfo],
            ['24:60', null, null, fnCultureInfo],
            ['24:120', null, null, fnCultureInfo],
            ['25:0', '[h]:mm:ss', 1.0416666666666667, fnCultureInfo],
            ['25:30', '[h]:mm:ss', 1.0625, fnCultureInfo],
            ['25:60', null, null, fnCultureInfo],

            // 3 - values
            ['0:0:0', 'h:mm:ss', 0, fnCultureInfo],
            ['0:0:30', 'h:mm:ss', 0.00034722222222222224, fnCultureInfo],
            ['0:0:60', 'General', 0.0006944444444444445, fnCultureInfo],
            ['0:0:9999', 'General', 0.11572916666666666, fnCultureInfo],
            ['0:0:10000', null, null, fnCultureInfo],

            ['0:30:30', 'h:mm:ss', 0.021180555555555557, fnCultureInfo],
            ['0:30:60', 'General', 0.021527777777777778, fnCultureInfo],
            ['0:30:9999', 'General', 0.1365625, fnCultureInfo],
            ['0:30:10000', null, null, fnCultureInfo],


            ['0:60:30', 'General', 0.04201388888888889, fnCultureInfo],
            ['0:60:60', null, null, fnCultureInfo],
            ['0:60:9999', null, null, fnCultureInfo],

            ['0:9999:30', 'General', 6.944097222222222, fnCultureInfo],
            ['0:9999:60', null, null, fnCultureInfo],
            ['0:9999:9999', null, null, fnCultureInfo],

            ['12:0:0', 'h:mm:ss', 0.5, fnCultureInfo],
            ['12:0:30', 'h:mm:ss', 0.5003472222222223, fnCultureInfo],
            ['12:0:60', 'General', 0.5006944444444444, fnCultureInfo],
            ['12:60:0', 'General', 0.5416666666666666, fnCultureInfo],
            ['12:60:30', 'General', 0.5420138888888889, fnCultureInfo],
            ['12:60:60', null, null, fnCultureInfo],

            ['25:0:0', '[h]:mm:ss', 1.0416666666666667, fnCultureInfo],
            ['25:0:30', '[h]:mm:ss', 1.0420138888888888, fnCultureInfo],
            ['25:30:0', '[h]:mm:ss', 1.0625, fnCultureInfo],
            ['25:0:60', null, null, fnCultureInfo],
            ['25:60:0', null, null, fnCultureInfo],
            ['23:60:59', 'General', 1.0006828703703703, fnCultureInfo],
            ['24:60:0 ', null, null, fnCultureInfo],


            // Short date pattern 205 (american)

            // Date values

            // 2 - values

            // . separator
            ['0.0', null, null, enCultureInfo],
            ['0.6', null, null, enCultureInfo],
            ['6.0', null, null, enCultureInfo],
            ['12.0', null, null, enCultureInfo],
            ['20.0', null, null, enCultureInfo],
            ['13.15', null,  null, enCultureInfo],
            ['15.13', null,  null, enCultureInfo],
            ['0.9999', null, null, enCultureInfo],
            ['0.10000', null, null, enCultureInfo],
            
            ['6.15', null, null, enCultureInfo],
            ['15.6', null, null, enCultureInfo],
            ['15.12', null, null, enCultureInfo],
            ['31.12', null, null, enCultureInfo],
            ['6.31', null, null, enCultureInfo],
            ['6.9999', null, null, enCultureInfo],
            ['12.50', null, null, enCultureInfo],
            // / separator
            ['0/0', null, null, enCultureInfo],
            ['0/6', null, null, enCultureInfo],
            ['6/0', 'm-yyyy', 123, enCultureInfo],
            ['12/0', 'm-yyyy', 123, enCultureInfo],
            ['20/0', null, null, enCultureInfo],
            ['13/15', null, null, enCultureInfo],
            ['15/13', null, null, enCultureInfo],
            ['0/9999', null, null, enCultureInfo],
            ['0/10000', null, null, enCultureInfo],
            
            ['6/15', 'd-mmm', 45823, enCultureInfo],
            ['15/6', null, null, enCultureInfo],
            ['15/12', null, null, enCultureInfo],
            ['31/12', null, null, enCultureInfo],
            ['6/31', 'mmm-yy', 11475, enCultureInfo],
            ['6/9999', 'mmm-yy', 2958252, enCultureInfo],
            ['12/50', 'mmm-yy', 18598, enCultureInfo],

            // 3 - values

            // . separator
            ['0.0.0', null, null, enCultureInfo],
            ['0.0.2000', null, null, enCultureInfo],
            ['0.6.2000', null, null, enCultureInfo],
            ['0.13.2000', null, null, enCultureInfo],
            ['6.0.2000', null, null, enCultureInfo],
            ['13.0.2000', null, null, enCultureInfo],
            ['01.01.0000', null, null, enCultureInfo],


            ['01.01.2000', null, null, enCultureInfo],
            ['15.6.2000', null, null, enCultureInfo],
            ['15.12.2000', null, null, enCultureInfo],
            ['31.12.9999', null, null, enCultureInfo],
            
            ['6.15.2000', null, null, enCultureInfo],
            ['12.15.2000', null, null, enCultureInfo],
            ['50.12.2000', null, null, enCultureInfo],
            ['15.6.10000', null, null, enCultureInfo],
            ['6.15.10000', null, null, enCultureInfo],

            // / separator
            ['0/0/0', null, null, enCultureInfo],
            ['0/0/2000', null, null, enCultureInfo],
            ['0/6/2000', null, null, enCultureInfo],
            ['0/13/2000', null, null, enCultureInfo],
            ['6/0/2000', null, null, enCultureInfo],
            ['13/0/2000', null, null, enCultureInfo],
            ['01/01/0000', null, null, enCultureInfo],

            ['01/01/2000', 'm/d/yyyy', 36526, enCultureInfo],
            ['15/6/2000', null, null, enCultureInfo],
            ['15/12/2000', null, null, enCultureInfo],
            ['31/12/9999', null, null, enCultureInfo],
            
            ['6/15/2000', 'm/d/yyyy', 36692, enCultureInfo],
            ['12/15/2000', 'm/d/yyyy', 36875, enCultureInfo],
            ['50/12/2000', null, null, enCultureInfo],
            ['15/6/10000', null, null, enCultureInfo],
            ['6/15/10000', null, null, enCultureInfo],


            // Time values

            // 2 - values
            ['0:0', 'h:mm', 0, enCultureInfo],
            ['0:30', 'h:mm', 0.020833333333333332, enCultureInfo],
            ['0:60', 'General', 0.041666666666666664, enCultureInfo],
            ['0:120', 'General', 0.08333333333333333, enCultureInfo],
            ['0:9999', 'General', 6.94375, enCultureInfo],
            ['0:10000', null, null, enCultureInfo],
            ['12:0', 'h:mm', 0.5, enCultureInfo],
            ['12:30', 'h:mm', 0.5208333333333334, enCultureInfo],
            ['12:60', 'General', 0.5416666666666666, enCultureInfo],
            ['12:120', 'General', 0.5833333333333334, enCultureInfo],
            ['24:0', 'h:mm', 1, enCultureInfo],
            ['24:30', '[h]:mm:ss', 1.0208333333333333, enCultureInfo],
            ['24:60', null, null, enCultureInfo],
            ['24:120', null, null, enCultureInfo],
            ['25:0', '[h]:mm:ss', 1.0416666666666667, enCultureInfo],
            ['25:30', '[h]:mm:ss', 1.0625, enCultureInfo],
            ['25:60', null, null, enCultureInfo],

            // 3 - values
            ['0:0:0', 'h:mm:ss', 0, enCultureInfo],
            ['0:0:30', 'h:mm:ss', 0.00034722222222222224, enCultureInfo],
            ['0:0:60', 'General', 0.0006944444444444445, enCultureInfo],
            ['0:0:9999', 'General', 0.11572916666666666, enCultureInfo],
            ['0:0:10000', null, null, enCultureInfo],
            ['0:30:30', 'h:mm:ss', 0.021180555555555557, enCultureInfo],
            ['0:30:60', 'General', 0.021527777777777778, enCultureInfo],
            ['0:30:9999', 'General', 0.1365625, enCultureInfo],
            ['0:30:10000', null, null, enCultureInfo],


            ['0:60:30', 'General', 0.04201388888888889, enCultureInfo],
            ['0:60:60', null, null, enCultureInfo],
            ['0:60:9999', null, null, enCultureInfo],

            ['0:9999:30', 'General', 6.944097222222222, enCultureInfo],
            ['0:9999:60', null, null, enCultureInfo],
            ['0:9999:9999', null, null, enCultureInfo],

            ['12:0:0', 'h:mm:ss', 0.5, enCultureInfo],
            ['12:0:30', 'h:mm:ss', 0.5003472222222223, enCultureInfo],
            ['12:0:60', 'General', 0.5006944444444444, enCultureInfo],
            ['12:60:0', 'General', 0.5416666666666666, enCultureInfo],
            ['12:60:30', 'General', 0.5420138888888889, enCultureInfo],
            ['12:60:60', null, null, enCultureInfo],

            ['25:0:0', '[h]:mm:ss', 1.0416666666666667, enCultureInfo],
            ['25:0:30', '[h]:mm:ss', 1.0420138888888888, enCultureInfo],
            ['25:30:0', '[h]:mm:ss', 1.0625, enCultureInfo],
            ['25:0:60', null, null, enCultureInfo],
            ['25:60:0', null, null, enCultureInfo],
            ['23:60:59', 'General', 1.0006828703703703, enCultureInfo],
            ['24:60:0 ', null, null, enCultureInfo],
            

            // Short date pattern 135 (european)
            
            // Date values

            // 2 - values

            // . separator
            ['0.0', null, null, ruCultureInfo],
            ['0.6', null, null, ruCultureInfo],
            ['6.0', null, null, ruCultureInfo],
            ['12.0', null, null, ruCultureInfo],
            ['20.0', null, null, ruCultureInfo],
            ['13.15', null, null, ruCultureInfo],
            ['15.13', null, null, ruCultureInfo],
            ['0.9999', null, null, ruCultureInfo],
            ['0.10000', null, null, ruCultureInfo],
            
            ['6.15', 'mmm-yy', 42156, ruCultureInfo],
            ['15.6', 'd-mmm', 45823, ruCultureInfo],
            ['15.12', 'd-mmm', 46006, ruCultureInfo],
            ['31.12', 'd-mmm', 46022, ruCultureInfo],
            ['6.31', 'mmm-yy', 11475, ruCultureInfo],
            ['6.9999', 'mmm-yy', 2958252, ruCultureInfo],
            ['12.50', 'mmm-yy', 18598, ruCultureInfo],

            // / separator
            ['0/0', null, null, ruCultureInfo],
            ['0/6', null, null, ruCultureInfo],
            ['6/0', null, null, ruCultureInfo],
            ['12/0', null, null, ruCultureInfo],
            ['20/0', null, null, ruCultureInfo],
            ['13/15', null, null, ruCultureInfo],
            ['15/13', null, null, ruCultureInfo],
            ['0/9999', null, null, ruCultureInfo],
            ['0/10000', null, null, ruCultureInfo],
            
            ['6/15', 'mmm-yy', 42156, ruCultureInfo],
            ['15/6', 'd-mmm', 45823, ruCultureInfo],
            ['15/12', 'd-mmm', 46006, ruCultureInfo],
            ['31/12', 'd-mmm', 46022, ruCultureInfo],
            ['6/31', 'mmm-yy', 11475, ruCultureInfo],
            ['6/9999', 'mmm-yy', 2958252, ruCultureInfo],
            ['12/50', 'mmm-yy', 18598, ruCultureInfo],

            // 3 - values

            // . separator
            ['0.0.0', null, null, ruCultureInfo],
            ['0.0.2000', null, null, ruCultureInfo],
            ['0.6.2000', null, null, ruCultureInfo],
            ['0.13.2000', null, null, ruCultureInfo],
            ['6.0.2000', null, null, ruCultureInfo],
            ['13.0.2000', null, null, ruCultureInfo],
            ['01.01.0000', null, null, ruCultureInfo],


            ['01.01.2000', 'dd/mm/yyyy', 36526, ruCultureInfo],
            ['15.6.2000', 'dd/mm/yyyy', 36692, ruCultureInfo],
            ['15.12.2000', 'dd/mm/yyyy', 36875, ruCultureInfo],
            ['31.12.9999', 'dd/mm/yyyy', 2958465, ruCultureInfo],
            
            ['6.15.2000', null, null, ruCultureInfo],
            ['12.15.2000', null, null, ruCultureInfo],
            ['50.12.2000', null, null, ruCultureInfo],
            ['15.6.10000', null, null, ruCultureInfo],
            ['6.15.10000', null, null, ruCultureInfo],
            // / separator
            ['0/0/0', null, null, ruCultureInfo],
            ['0/0/2000', null, null, ruCultureInfo],
            ['0/6/2000', null, null, ruCultureInfo],
            ['0/13/2000', null, null, ruCultureInfo],
            ['6/0/2000', null, null, ruCultureInfo],
            ['13/0/2000', null, null, ruCultureInfo],
            ['01/01/0000', null, null, ruCultureInfo],


            ['01/01/2000', 'dd/mm/yyyy', 36526, ruCultureInfo],
            ['15/6/2000', 'dd/mm/yyyy', 36692, ruCultureInfo],
            ['15/12/2000', 'dd/mm/yyyy', 36875, ruCultureInfo],
            ['31/12/9999', 'dd/mm/yyyy', 2958465, ruCultureInfo],
            
            ['6/15/2000', null, null, ruCultureInfo],
            ['12/15/2000', null, null, ruCultureInfo],
            ['50/12/2000', null, null, ruCultureInfo],
            ['15/6/10000', null, null, ruCultureInfo],
            ['6/15/10000', null, null, ruCultureInfo],


            // Time values

            // 2 - values
            ['0:0', 'h:mm', 0, ruCultureInfo],
            ['0:30', 'h:mm', 0.020833333333333332, ruCultureInfo],
            ['0:60', 'General', 0.041666666666666664, ruCultureInfo],
            ['0:120', 'General', 0.08333333333333333, ruCultureInfo],
            ['0:9999', 'General', 6.94375, ruCultureInfo],
            ['0:10000', null, null, ruCultureInfo],
            ['12:0', 'h:mm', 0.5, ruCultureInfo],
            ['12:30', 'h:mm', 0.5208333333333334, ruCultureInfo],
            ['12:60', 'General', 0.5416666666666666, ruCultureInfo],
            ['12:120', 'General', 0.5833333333333334, ruCultureInfo],
            ['24:0', 'h:mm', 1, ruCultureInfo],
            ['24:30', '[h]:mm:ss', 1.0208333333333333, ruCultureInfo],
            ['24:60', null, null, ruCultureInfo],
            ['24:120', null, null, ruCultureInfo],
            ['25:0', '[h]:mm:ss', 1.0416666666666667, ruCultureInfo],
            ['25:30', '[h]:mm:ss', 1.0625, ruCultureInfo],
            ['25:60', null, null, ruCultureInfo],

            // 3 - values
            ['0:0:0', 'h:mm:ss', 0, ruCultureInfo],
            ['0:0:30', 'h:mm:ss', 0.00034722222222222224, ruCultureInfo],
            ['0:0:60', 'General', 0.0006944444444444445, ruCultureInfo],
            ['0:0:9999', 'General', 0.11572916666666666, ruCultureInfo],
            ['0:0:10000', null, null, ruCultureInfo],

            ['0:30:30', 'h:mm:ss', 0.021180555555555557, ruCultureInfo],
            ['0:30:60', 'General', 0.021527777777777778, ruCultureInfo],
            ['0:30:9999', 'General', 0.1365625, ruCultureInfo],
            ['0:30:10000', null, null, ruCultureInfo],


            ['0:60:30', 'General', 0.04201388888888889, ruCultureInfo],
            ['0:60:60', null, null, ruCultureInfo],
            ['0:60:9999', null, null, ruCultureInfo],

            ['0:9999:30', 'General', 6.944097222222222, ruCultureInfo],
            ['0:9999:60', null, null, ruCultureInfo],
            ['0:9999:9999', null, null, ruCultureInfo],

            ['12:0:0', 'h:mm:ss', 0.5, ruCultureInfo],
            ['12:0:30', 'h:mm:ss', 0.5003472222222223, ruCultureInfo],
            ['12:0:60', 'General', 0.5006944444444444, ruCultureInfo],
            ['12:60:0', 'General', 0.5416666666666666, ruCultureInfo],
            ['12:60:30', 'General', 0.5420138888888889, ruCultureInfo],
            ['12:60:60', null, null, ruCultureInfo],

            ['25:0:0', '[h]:mm:ss', 1.0416666666666667, ruCultureInfo],
            ['25:0:30', '[h]:mm:ss', 1.0420138888888888, ruCultureInfo],
            ['25:30:0', '[h]:mm:ss', 1.0625, ruCultureInfo],
            ['25:0:60', null, null, ruCultureInfo],
            ['25:60:0', null, null, ruCultureInfo],
            ['23:60:59', 'General', 1.0006828703703703, ruCultureInfo],
            ['24:60:0 ', null, null, ruCultureInfo],


            // Short date pattern 531 (asian)

            // Date values

            // 2 - values
            // . separator
            ['0.0', null, null, jpCultureInfo],
            ['0.6', null, null, jpCultureInfo],
            ['6.0', null, null, jpCultureInfo],
            ['12.0', null, null, jpCultureInfo],
            ['20.0', null, null, jpCultureInfo],
            ['13.15', null, null, jpCultureInfo],
            ['15.13', null, null, jpCultureInfo],
            ['0.9999', null, null, jpCultureInfo],
            ['0.10000', null, null, jpCultureInfo],
            
            ['6.15', null, null, jpCultureInfo],
            ['15.6', null, null, jpCultureInfo],
            ['15.12', null, null, jpCultureInfo],
            ['31.12', null, null, jpCultureInfo],
            ['6.31', null, null, jpCultureInfo],
            ['6.9999', null, null, jpCultureInfo],
            ['12.50', null, null, jpCultureInfo],

            // / separator
            ['0/0', null, null, jpCultureInfo],
            ['0/6', null, null, jpCultureInfo],
            ['6/0', null, null, jpCultureInfo],
            ['12/0', null, null, jpCultureInfo],
            ['20/0', null, null, jpCultureInfo],
            ['13/15', null, null, jpCultureInfo],
            ['15/13', null, null, jpCultureInfo],
            ['0/9999', null, null, jpCultureInfo],
            ['0/10000', null, null, jpCultureInfo],
            
            ['6/15', 'd-mmm', 45823, jpCultureInfo],
            ['15/6', 'd-mmm', 45823, jpCultureInfo],
            ['15/12', 'd-mmm', 46006, jpCultureInfo],
            ['31/12', 'd-mmm', 46022, jpCultureInfo],
            ['6/31', null, null, jpCultureInfo],
            ['6/9999', null, null, jpCultureInfo],
            ['12/50', null, null, jpCultureInfo],

            // 3 - values

            // . separator
            ['0.0.0', null, null, jpCultureInfo],
            ['0.0.2000', null, null, jpCultureInfo],
            ['0.6.2000', null, null, jpCultureInfo],
            ['0.13.2000', null, null, jpCultureInfo],
            ['6.0.2000', null, null, jpCultureInfo],
            ['13.0.2000', null, null, jpCultureInfo],
            ['01.01.0000', null, null, jpCultureInfo],


            ['01.01.2000', null, null, jpCultureInfo],
            ['15.6.2000', null, null, jpCultureInfo],
            ['15.12.2000', null, null, jpCultureInfo],
            ['31.12.9999', null, null, jpCultureInfo],
            
            ['6.15.2000', null, null, jpCultureInfo],
            ['12.15.2000', null, null, jpCultureInfo],
            ['50.12.2000', null, null, jpCultureInfo],
            ['15.6.10000', null, null, jpCultureInfo],
            ['6.15.10000', null, null, jpCultureInfo],

            // / separator
            ['0/0/0', null, null, jpCultureInfo],
            ['0/0/2000', null, null, jpCultureInfo],
            ['0/6/2000', null, null, jpCultureInfo],
            ['0/13/2000', null, null, jpCultureInfo],
            ['6/0/2000', null, null, jpCultureInfo],
            ['13/0/2000', null, null, jpCultureInfo],
            ['01/01/0000', null, null, jpCultureInfo],


            ['01/01/2000', null, 123, jpCultureInfo],
            ['15/6/2000', null, 123, jpCultureInfo],
            ['15/12/2000', null, 123, jpCultureInfo],
            ['31/12/9999', null, 123, jpCultureInfo],
            
            ['6/15/2000', null, null, jpCultureInfo],
            ['12/15/2000', null, null, jpCultureInfo],
            ['50/12/2000', null, null, jpCultureInfo],
            ['15/6/10000', null, null, jpCultureInfo],
            ['6/15/10000', null, null, jpCultureInfo],


            // Time values
            // 2 - values
            ['0:0', 'h:mm', 0, jpCultureInfo],
            ['0:30', 'h:mm', 0.020833333333333332, jpCultureInfo],
            ['0:60', 'General', 0.041666666666666664, jpCultureInfo],
            ['0:120', 'General', 0.08333333333333333, jpCultureInfo],
            ['0:9999', 'General', 6.94375, jpCultureInfo],
            ['0:10000', null, null, jpCultureInfo],
            ['12:0', 'h:mm', 0.5, jpCultureInfo],
            ['12:30', 'h:mm', 0.5208333333333334, jpCultureInfo],
            ['12:60', 'General', 0.5416666666666666, jpCultureInfo],
            ['12:120', 'General', 0.5833333333333334, jpCultureInfo],
            ['24:0', 'h:mm', 1, jpCultureInfo],
            ['24:30', '[h]:mm:ss', 1.0208333333333333, jpCultureInfo],
            ['24:60', null, null, jpCultureInfo],
            ['24:120', null, null, jpCultureInfo],
            ['25:0', '[h]:mm:ss', 1.0416666666666667, jpCultureInfo],
            ['25:30', '[h]:mm:ss', 1.0625, jpCultureInfo],
            ['25:60', null, null, jpCultureInfo],

            // 3 - values
            ['0:0:0', 'h:mm:ss', 0, jpCultureInfo],
            ['0:0:30', 'h:mm:ss', 0.00034722222222222224, jpCultureInfo],
            ['0:0:60', 'General', 0.0006944444444444445, jpCultureInfo],
            ['0:0:9999', 'General', 0.11572916666666666, jpCultureInfo],
            ['0:0:10000', null, null, jpCultureInfo],

            ['0:30:30', 'h:mm:ss', 0.021180555555555557, jpCultureInfo],
            ['0:30:60', 'General', 0.021527777777777778, jpCultureInfo],
            ['0:30:9999', 'General', 0.1365625, jpCultureInfo],
            ['0:30:10000', null, null, jpCultureInfo],


            ['0:60:30', 'General', 0.04201388888888889, jpCultureInfo],
            ['0:60:60', null, null, jpCultureInfo],
            ['0:60:9999', null, null, jpCultureInfo],

            ['0:9999:30', 'General', 6.944097222222222, jpCultureInfo],
            ['0:9999:60', null, null, jpCultureInfo],
            ['0:9999:9999', null, null, jpCultureInfo],

            ['12:0:0', 'h:mm:ss', 0.5, jpCultureInfo],
            ['12:0:30', 'h:mm:ss', 0.5003472222222223, jpCultureInfo],
            ['12:0:60', 'General', 0.5006944444444444, jpCultureInfo],
            ['12:60:0', 'General', 0.5416666666666666, jpCultureInfo],
            ['12:60:30', 'General', 0.5420138888888889, jpCultureInfo],
            ['12:60:60', null, null, jpCultureInfo],

            ['25:0:0', '[h]:mm:ss', 1.0416666666666667, jpCultureInfo],
            ['25:0:30', '[h]:mm:ss', 1.0420138888888888, jpCultureInfo],
            ['25:30:0', '[h]:mm:ss', 1.0625, jpCultureInfo],
            ['25:0:60', null, null, jpCultureInfo],
            ['23:60:59', 'General', 1.0006828703703703, jpCultureInfo],
            ['24:60:0 ', null, null, jpCultureInfo],
            ['25:60:0', null, null, jpCultureInfo],
        ];

        for (let i = 0; i < testCases.length; i++) {
            let value = testCases[i][0];
            let expectedFormat = testCases[i][1];
            let expectedValue = testCases[i][2];
            let cultureInfo = testCases[i][3];
            
            let formatted = AscCommon.g_oFormatParser.parseDate(value, cultureInfo);

            if (formatted)
            {
                assert.strictEqual(formatted.format, expectedFormat, `Case format: ${value}. Expected format: ${expectedFormat}`);              
                assert.strictEqual(formatted.value, expectedValue, `Case format: ${value}. Expected value: ${expectedValue}`);               
            } else {
                assert.strictEqual(formatted, expectedFormat, `Case format: ${expectedFormat}`);              

            }

        }      
    });


});

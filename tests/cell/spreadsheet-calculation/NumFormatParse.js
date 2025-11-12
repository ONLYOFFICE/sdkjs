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

});

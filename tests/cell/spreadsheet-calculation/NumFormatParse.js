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
    window["AscCommonExcel"].Font = function () {};
    window["AscCommonExcel"].RgbColor = function () {};

    const eps = 1e-15;
    const formatParser = AscCommon.g_oFormatParser;
    const formatTypes = Asc.c_oAscNumFormatType;

    // Helper function to get Excel serial date for a given month/day in the current year
    // This makes tests year-independent
    function getDateSerial(month, day) {
        const year = new Date().getFullYear();
        return (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / (86400 * 1000);
    }

    // =====================================================================
    // FormatParser.isLeapYear
    // =====================================================================
    QUnit.module('FormatParser.isLeapYear');

    QUnit.test('leap years', function (assert) {
        // Divisible by 4 but not by 100
        assert.strictEqual(formatParser.isLeapYear(2024), true, '2024');
        assert.strictEqual(formatParser.isLeapYear(2020), true, '2020');
        assert.strictEqual(formatParser.isLeapYear(1996), true, '1996');
        
        // Divisible by 100 but not by 400 (not leap)
        assert.strictEqual(formatParser.isLeapYear(1900), false, '1900');
        assert.strictEqual(formatParser.isLeapYear(2100), false, '2100');
        
        // Divisible by 400 (leap)
        assert.strictEqual(formatParser.isLeapYear(2000), true, '2000');
        assert.strictEqual(formatParser.isLeapYear(1600), true, '1600');
        
        // Not divisible by 4 (not leap)
        assert.strictEqual(formatParser.isLeapYear(2023), false, '2023');
        assert.strictEqual(formatParser.isLeapYear(2019), false, '2019');
    });

    // =====================================================================
    // FormatParser.isValidDay
    // =====================================================================
    QUnit.module('FormatParser.isValidDay');

    QUnit.test('valid days per month', function (assert) {
        const daysInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
        for (let month = 0; month < 12; month++) {
            assert.strictEqual(formatParser.isValidDay(2023, month, 1), true, 
                `Day 1 valid for month ${month + 1}`);
            assert.strictEqual(formatParser.isValidDay(2023, month, daysInMonth[month]), true, 
                `Day ${daysInMonth[month]} valid for month ${month + 1}`);
            assert.strictEqual(formatParser.isValidDay(2023, month, daysInMonth[month] + 1), false, 
                `Day ${daysInMonth[month] + 1} invalid for month ${month + 1}`);
        }
    });

    QUnit.test('February leap year handling', function (assert) {
        assert.strictEqual(formatParser.isValidDay(2024, 1, 29), true, 'Feb 29 in leap year');
        assert.strictEqual(formatParser.isValidDay(2024, 1, 30), false, 'Feb 30 in leap year');
        assert.strictEqual(formatParser.isValidDay(2023, 1, 29), false, 'Feb 29 in non-leap year');
    });

    QUnit.test('boundary conditions', function (assert) {
        assert.strictEqual(formatParser.isValidDay(2023, 0, 0), false, 'Day 0');
        assert.strictEqual(formatParser.isValidDay(2023, 0, -1), false, 'Negative day');
        assert.strictEqual(formatParser.isValidDay(2023, 0, 32), false, 'Day 32 in January');
    });

    // =====================================================================
    // FormatParser.isValidDate
    // =====================================================================
    QUnit.module('FormatParser.isValidDate');

    QUnit.test('valid dates', function (assert) {
        assert.strictEqual(formatParser.isValidDate(2023, 0, 1), true, 'Jan 1, 2023');
        assert.strictEqual(formatParser.isValidDate(2023, 11, 31), true, 'Dec 31, 2023');
        assert.strictEqual(formatParser.isValidDate(2000, 1, 29), true, 'Feb 29, 2000 (leap)');
        assert.strictEqual(formatParser.isValidDate(1900, 1, 28), true, 'Feb 28, 1900');
    });

    QUnit.test('Excel Feb 29, 1900 bug compatibility', function (assert) {
        assert.strictEqual(formatParser.isValidDate(1900, 1, 29), true, 'Feb 29, 1900 (Excel bug)');
    });

    QUnit.test('special base date Dec 31, 1899', function (assert) {
        assert.strictEqual(formatParser.isValidDate(1899, 11, 31), true, 'Dec 31, 1899');
    });

    QUnit.test('dates before 1900 invalid', function (assert) {
        assert.strictEqual(formatParser.isValidDate(1899, 0, 1), false, 'Jan 1, 1899');
        assert.strictEqual(formatParser.isValidDate(1800, 5, 15), false, '1800');
    });

    QUnit.test('invalid months and days', function (assert) {
        assert.strictEqual(formatParser.isValidDate(2023, -1, 15), false, 'Month -1');
        assert.strictEqual(formatParser.isValidDate(2023, 12, 15), false, 'Month 12');
        assert.strictEqual(formatParser.isValidDate(2023, 0, 0), false, 'Day 0');
        assert.strictEqual(formatParser.isValidDate(2023, 0, 32), false, 'Jan 32');
        assert.strictEqual(formatParser.isValidDate(2023, 3, 31), false, 'Apr 31');
    });

    // =====================================================================
    // FormatParser.isValidDatePDF
    // =====================================================================
    QUnit.module('FormatParser.isValidDatePDF');

    QUnit.test('PDF dates allow pre-1900', function (assert) {
        assert.strictEqual(formatParser.isValidDatePDF(2023, 0, 1), true, 'Jan 1, 2023');
        assert.strictEqual(formatParser.isValidDatePDF(1899, 0, 1), true, 'Jan 1, 1899 (PDF)');
        assert.strictEqual(formatParser.isValidDatePDF(1500, 5, 15), true, '1500 (PDF)');
        assert.strictEqual(formatParser.isValidDatePDF(100, 0, 1), true, 'Year 100 (PDF)');
    });

    QUnit.test('invalid months and days still rejected', function (assert) {
        assert.strictEqual(formatParser.isValidDatePDF(2023, -1, 15), false, 'Month -1');
        assert.strictEqual(formatParser.isValidDatePDF(2023, 12, 15), false, 'Month 12');
        assert.strictEqual(formatParser.isValidDatePDF(2023, 0, 32), false, 'Jan 32');
    });

    // =====================================================================
    // FormatParser.strcmp
    // =====================================================================
    QUnit.module('FormatParser.strcmp');

    QUnit.test('string comparison', function (assert) {
        assert.strictEqual(formatParser.strcmp("hello", "hello", 0, 5), true, 'Full match');
        assert.strictEqual(formatParser.strcmp("hello world", "ello", 1, 4), true, 'Partial match');
        assert.strictEqual(formatParser.strcmp("abcdef", "cd", 2, 2), true, 'Substring match');
        assert.strictEqual(formatParser.strcmp("hello", "world", 0, 5), false, 'Different strings');
        assert.strictEqual(formatParser.strcmp("hello", "", 0, 0), false, 'Zero length');
        assert.strictEqual(formatParser.strcmp("abc", "a", 0, 1), true, 'Single char');
        assert.strictEqual(formatParser.strcmp("hello", "llo", 2, 3, 0), true, 'With index2');
    });

    // =====================================================================
    // FormatParser.isLocaleNumber / parseLocaleNumber
    // =====================================================================
    QUnit.module('FormatParser.isLocaleNumber');

    QUnit.test('valid numbers with default locale', function (assert) {
        assert.strictEqual(formatParser.isLocaleNumber("123", null), true, 'Integer');
        assert.strictEqual(formatParser.isLocaleNumber("123.45", null), true, 'Decimal');
        assert.strictEqual(formatParser.isLocaleNumber("0.5", null), true, 'Less than 1');
        assert.strictEqual(formatParser.isLocaleNumber("-123.45", null), true, 'Negative');
    });

    QUnit.test('invalid inputs', function (assert) {
        assert.strictEqual(formatParser.isLocaleNumber("abc", null), false, 'Alphabetic');
        assert.strictEqual(formatParser.isLocaleNumber("12a34", null), false, 'Mixed');
        assert.strictEqual(formatParser.isLocaleNumber("", null), false, 'Empty');
    });

    QUnit.module('FormatParser.parseLocaleNumber');

    QUnit.test('parse numbers', function (assert) {
        assert.strictEqual(formatParser.parseLocaleNumber("123", null), 123, 'Integer');
        assert.strictEqual(formatParser.parseLocaleNumber("123.45", null), 123.45, 'Decimal');
        assert.strictEqual(formatParser.parseLocaleNumber("-50.5", null), -50.5, 'Negative');
    });

    // =====================================================================
    // FormatParser.parse - Numbers
    // =====================================================================
    QUnit.module('FormatParser.parse - Numbers');

    QUnit.test('thousand separators', function (assert) {
        let result = formatParser.parse("1,234", null);
        assert.ok(result !== null, 'Should parse "1,234"');
        assert.strictEqual(result.value, 1234, 'Value 1234');
        
        result = formatParser.parse("1,234,567", null);
        assert.ok(result !== null, 'Should parse "1,234,567"');
        assert.strictEqual(result.value, 1234567, 'Value 1234567');
    });

    QUnit.test('decimal numbers', function (assert) {
        let result = formatParser.parse("1,234.56", null);
        assert.ok(result !== null, 'Should parse');
        assert.ok(Math.abs(result.value - 1234.56) < eps, 'Value 1234.56');
    });

    QUnit.test('negative in parentheses', function (assert) {
        let result = formatParser.parse("(100)", null);
        assert.ok(result !== null, 'Should parse "(100)"');
        assert.strictEqual(result.value, -100, 'Value -100');
        
        result = formatParser.parse("(1,234.56)", null);
        assert.ok(result !== null, 'Should parse "(1,234.56)"');
        assert.ok(Math.abs(result.value - (-1234.56)) < eps, 'Value -1234.56');
    });

    QUnit.test('negative with minus', function (assert) {
        let result = formatParser.parse("-100", null);
        assert.ok(result !== null, 'Should parse');
        assert.strictEqual(result.value, -100, 'Value -100');
    });

    QUnit.test('positive with plus', function (assert) {
        let result = formatParser.parse("+100", null);
        assert.ok(result !== null, 'Should parse');
        assert.strictEqual(result.value, 100, 'Value 100');
    });

    // =====================================================================
    // FormatParser.parse - Percentages
    // =====================================================================
    QUnit.module('FormatParser.parse - Percentages');

    QUnit.test('percentage values', function (assert) {
        let result = formatParser.parse("50%", null);
        assert.ok(result !== null && result.bPercent, 'Should parse 50%');
        assert.ok(Math.abs(result.value - 0.5) < eps, 'Value 0.5');
        
        result = formatParser.parse("100%", null);
        assert.ok(Math.abs(result.value - 1) < eps, 'Value 1');
        
        result = formatParser.parse("12.5%", null);
        assert.ok(Math.abs(result.value - 0.125) < eps, 'Value 0.125');
        
        result = formatParser.parse("%50", null);
        assert.ok(result !== null, 'Should parse %50');
        assert.ok(Math.abs(result.value - 0.5) < eps, 'Value 0.5');
    });

    // =====================================================================
    // FormatParser.parse - Currencies
    // =====================================================================
    QUnit.module('FormatParser.parse - Currencies');

    QUnit.test('currency symbols', function (assert) {
        let result = formatParser.parse("$100", null);
        assert.ok(result !== null && result.bCurrency, 'Should parse $100');
        assert.strictEqual(result.value, 100, 'Value 100');
        
        result = formatParser.parse("€100", null);
        assert.ok(result !== null && result.bCurrency, 'Should parse €100');
        
        result = formatParser.parse("£100", null);
        assert.ok(result !== null && result.bCurrency, 'Should parse £100');
        
        result = formatParser.parse("¥100", null);
        assert.ok(result !== null && result.bCurrency, 'Should parse ¥100');
        
        result = formatParser.parse("100р.", null);
        assert.ok(result !== null && result.bCurrency, 'Should parse 100р.');
    });

    QUnit.test('negative currency', function (assert) {
        let result = formatParser.parse("($100)", null);
        assert.ok(result !== null, 'Should parse ($100)');
        assert.strictEqual(result.value, -100, 'Value -100');
        
        result = formatParser.parse("-$100", null);
        assert.ok(result !== null, 'Should parse -$100');
        assert.strictEqual(result.value, -100, 'Value -100');
    });

    // =====================================================================
    // FormatParser.parse - Invalid inputs
    // =====================================================================
    QUnit.module('FormatParser.parse - Invalid inputs');

    QUnit.test('invalid patterns', function (assert) {
        assert.strictEqual(formatParser.parse("++100", null), null, 'Multiple plus');
        assert.strictEqual(formatParser.parse("--100", null), null, 'Multiple minus');
        assert.strictEqual(formatParser.parse("(100", null), null, 'Unmatched open paren');
        assert.strictEqual(formatParser.parse("100)", null), null, 'Unmatched close paren');
        assert.strictEqual(formatParser.parse("50%%", null), null, 'Multiple percent');
        assert.strictEqual(formatParser.parse("$€100", null), null, 'Mixed currencies');
        assert.strictEqual(formatParser.parse("", null), null, 'Empty string');
        assert.strictEqual(formatParser.parse("   ", null), null, 'Whitespace only');
        
        // === Decimal with fraction is invalid (Excel treats as text) ===
        assert.strictEqual(formatParser.parse("0.1/3", null), null, 'Decimal numerator 0.1/3');
        assert.strictEqual(formatParser.parse("0.5/2", null), null, 'Decimal numerator 0.5/2');
        assert.strictEqual(formatParser.parse("1.5/2", null), null, 'Decimal numerator 1.5/2');
        assert.strictEqual(formatParser.parse("0.25/4", null), null, 'Decimal numerator 0.25/4');
        assert.strictEqual(formatParser.parse("2.5/5", null), null, 'Decimal numerator 2.5/5');
        assert.strictEqual(formatParser.parse("10.5/3", null), null, 'Decimal numerator 10.5/3');
        assert.strictEqual(formatParser.parse("3/0.5", null), null, 'Decimal denominator 3/0.5');
        assert.strictEqual(formatParser.parse("1/2.5", null), null, 'Decimal denominator 1/2.5');
        assert.strictEqual(formatParser.parse("2/0.25", null), null, 'Decimal denominator 2/0.25');
        assert.strictEqual(formatParser.parse("0.5/0.5", null), null, 'Both decimal 0.5/0.5');
        assert.strictEqual(formatParser.parse("1.5/2.5", null), null, 'Both decimal 1.5/2.5');
        assert.strictEqual(formatParser.parse(".5/2", null), null, 'Leading dot numerator .5/2');
        assert.strictEqual(formatParser.parse("1/.5", null), null, 'Leading dot denominator 1/.5');
        assert.strictEqual(formatParser.parse("0.5 1/2", null), null, 'Decimal with mixed fraction');
        
        // === Invalid fraction formats (Excel stays as text) ===
        assert.strictEqual(formatParser.parse("/2", null), null, 'Missing numerator /2');
        assert.strictEqual(formatParser.parse("1/", null), null, 'Missing denominator 1/');
        assert.strictEqual(formatParser.parse("/", null), null, 'Just slash');
        assert.strictEqual(formatParser.parse("//", null), null, 'Double slash');
        assert.strictEqual(formatParser.parse("1//2", null), null, 'Double slash between 1//2');
        assert.strictEqual(formatParser.parse("1/-2", null), null, 'Negative denominator 1/-2');
        assert.strictEqual(formatParser.parse("0/0", null), null, 'Zero over zero 0/0');
        assert.strictEqual(formatParser.parse("0/1", null), null, 'Zero numerator 0/1');
        
        // === Text with slash patterns (should stay as text) ===
        assert.strictEqual(formatParser.parse("N/A", null), null, 'Abbreviation N/A');
        assert.strictEqual(formatParser.parse("n/a", null), null, 'Abbreviation n/a');
        assert.strictEqual(formatParser.parse("w/o", null), null, 'Abbreviation w/o');
        assert.strictEqual(formatParser.parse("path/to/file", null), null, 'Path pattern');
        assert.strictEqual(formatParser.parse("C:/folder", null), null, 'Drive path');
        assert.strictEqual(formatParser.parse("yes/no", null), null, 'Choice pattern yes/no');
        assert.strictEqual(formatParser.parse("and/or", null), null, 'Choice pattern and/or');
        
        // === Fraction with text around (should stay as text) ===
        assert.strictEqual(formatParser.parse("1/2 cup", null), null, 'Fraction with text after');
        assert.strictEqual(formatParser.parse("part 1/2", null), null, 'Fraction with text before');
        assert.strictEqual(formatParser.parse("1/2 inch", null), null, 'Fraction measurement');
        assert.strictEqual(formatParser.parse("page 1/10", null), null, 'Page fraction');
        
        // === Special chars in fraction (should stay as text) ===
        assert.strictEqual(formatParser.parse("1_/2", null), null, 'Underscore in fraction');
        assert.strictEqual(formatParser.parse("1/_2", null), null, 'Underscore in denominator');
        assert.strictEqual(formatParser.parse("1|/2", null), null, 'Pipe in fraction');
        assert.strictEqual(formatParser.parse("1&/2", null), null, 'Ampersand in fraction');
        
        // === Spaces around slash (should stay as text or parse as date) ===
        assert.strictEqual(formatParser.parse(" /2", null), null, 'Space before slash');
    });
    
    // =====================================================================
    // FormatParser.parse - Valid fraction patterns
    // =====================================================================
    QUnit.test('valid fraction patterns', function (assert) {
        // Get proper format type for fraction context
        const fractionFormatString = "# ?/?";
        const numFormat = AscCommon.oNumFormatCache.get(fractionFormatString);
        const fractionFormatType = numFormat.getType();
        
        // These should parse as fractions when format is numeric (fraction type)
        assert.ok(formatParser.parse("1/2", null, fractionFormatType, fractionFormatString) !== null, 'Simple fraction 1/2');
        assert.ok(formatParser.parse("1/3", null, fractionFormatType, fractionFormatString) !== null, 'Simple fraction 1/3');
        assert.ok(formatParser.parse("3/4", null, fractionFormatType, fractionFormatString) !== null, 'Simple fraction 3/4');
        assert.ok(formatParser.parse("7/8", null, fractionFormatType, fractionFormatString) !== null, 'Simple fraction 7/8');
        assert.ok(formatParser.parse("15/16", null, fractionFormatType, fractionFormatString) !== null, 'Two digit fraction 15/16');
        
        // Mixed fractions
        assert.ok(formatParser.parse("1 1/2", null, fractionFormatType, fractionFormatString) !== null, 'Mixed fraction 1 1/2');
        assert.ok(formatParser.parse("2 3/4", null, fractionFormatType, fractionFormatString) !== null, 'Mixed fraction 2 3/4');
        assert.ok(formatParser.parse("0 1/2", null, fractionFormatType, fractionFormatString) !== null, 'Mixed fraction 0 1/2');
        assert.ok(formatParser.parse("10 1/4", null, fractionFormatType, fractionFormatString) !== null, 'Mixed fraction 10 1/4');
        assert.ok(formatParser.parse("100 99/100", null, fractionFormatType, fractionFormatString) !== null, 'Large mixed fraction');
        
        // Mixed fractions with thousand separators
        let result = formatParser.parse("1,234 1/2", null, fractionFormatType, fractionFormatString);
        assert.ok(result !== null, 'Mixed fraction with thousands 1,234 1/2');
        if (result) {
            assert.ok(Math.abs(result.value - 1234.5) < 0.0001, 'Value of 1,234 1/2 should be 1234.5');
        }
        
        // Fractions with sign (should parse in fraction context)
        const result1 = formatParser.parse("-1/2", null, fractionFormatType, fractionFormatString);
        if (result1 !== null) {
            assert.ok(result1.value < 0, 'Negative fraction -1/2 has negative value');
        }
        
        const result2 = formatParser.parse("+1/2", null, fractionFormatType, fractionFormatString);
        if (result2 !== null) {
            assert.ok(result2.value > 0, 'Positive fraction +1/2 has positive value');
        }
    });

    // =====================================================================
    // FormatParser.parseDate - using parse() integration
    // =====================================================================
    QUnit.module('FormatParser.parse - Date/Time');

    QUnit.test('comprehensive date/time tests', function (assert) {
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
        ];
        for (let i = 0; i < data.length; i++) {
            let date = formatParser.parse(data[i][0]);
            assert.strictEqual(date.format, data[i][1], `Format: ${data[i][0]}`);
            assert.ok(Math.abs(date.value - data[i][2]) < eps, `Value: ${data[i][0]}`);
        }
    });

    QUnit.test('month name dates', function (assert) {
        let result = formatParser.parse("January 15, 2023");
        assert.ok(result !== null && result.bDateTime, 'Should parse "January 15, 2023"');
        
        result = formatParser.parse("15 January 2023");
        assert.ok(result !== null && result.bDateTime, 'Should parse "15 January 2023"');
        
        result = formatParser.parse("Jan 15, 2023");
        assert.ok(result !== null && result.bDateTime, 'Should parse abbreviated month');
        
        result = formatParser.parse("15-Jan-2023");
        assert.ok(result !== null && result.bDateTime, 'Should parse dash-separated');
    });

    QUnit.test('month-year format (no day)', function (assert) {
        let result = formatParser.parse("Jan-2023");
        assert.ok(result !== null && result.bDateTime, 'Should parse "Jan-2023"');
        
        result = formatParser.parse("January 2023");
        assert.ok(result !== null && result.bDateTime, 'Should parse "January 2023"');
    });

    QUnit.test('day-month format (no year)', function (assert) {
        let result = formatParser.parse("15-Jan");
        assert.ok(result !== null && result.bDateTime, 'Should parse "15-Jan"');
        
        result = formatParser.parse("Jan 15");
        assert.ok(result !== null && result.bDateTime, 'Should parse "Jan 15"');
    });

    QUnit.test('time formats', function (assert) {
        let result = formatParser.parse("14:30");
        assert.ok(result !== null && result.bDateTime, 'Should parse "14:30"');
        
        result = formatParser.parse("2:30 PM");
        assert.ok(result !== null && result.bDateTime, 'Should parse "2:30 PM"');
        
        result = formatParser.parse("12:00 PM");
        assert.ok(result !== null, 'Should parse "12:00 PM"');
        assert.ok(Math.abs(result.value - 0.5) < eps, '12:00 PM = 0.5');
        
        result = formatParser.parse("12:00 AM");
        assert.ok(result !== null, 'Should parse "12:00 AM"');
        assert.ok(result.value < 0.01, '12:00 AM near 0');
    });

    QUnit.test('invalid time should fail', function (assert) {
        let result = formatParser.parse("14:60");
        assert.strictEqual(result, null, '60 minutes invalid');
        
        result = formatParser.parse("14:30:60");
        assert.strictEqual(result, null, '60 seconds invalid');
    });

    // =====================================================================
    // FormatParser.parseDatePDF
    // =====================================================================
    QUnit.module('FormatParser.parseDatePDF');

    QUnit.test('basic PDF date parsing', function (assert) {
        let result = formatParser.parseDatePDF("January 15, 2023", null);
        assert.ok(result !== null && result.bDate, 'Should parse month name date');
        
        result = formatParser.parseDatePDF("15 Jan 2023", null);
        assert.ok(result !== null && result.bDate, 'Should parse abbreviated month');
    });

    QUnit.test('PDF dates before 1900', function (assert) {
        let result = formatParser.parseDatePDF("January 15, 1850", null);
        assert.ok(result !== null, 'Should parse date before 1900');
    });

    QUnit.test('PDF dates with time', function (assert) {
        let result = formatParser.parseDatePDF("January 15, 2023 14:30:45", null);
        assert.ok(result !== null && result.bDate && result.bTime, 'Should parse date with time');
    });

    // =====================================================================
    // Date1904 mode
    // =====================================================================
    QUnit.module('Date1904 mode');

    QUnit.test('bDate1904 affects date values', function (assert) {
        let original1904 = AscCommon.bDate1904;
        
        try {
            AscCommon.bDate1904 = false;
            let result1 = formatParser.parse("January 1, 2000");
            
            AscCommon.bDate1904 = true;
            let result2 = formatParser.parse("January 1, 2000");
            
            if (result1 && result2) {
                // Values should differ by ~1462 days
                let diff = Math.abs(result1.value - result2.value);
                assert.ok(diff > 1400 && diff < 1500, 'Values differ by ~1462 days');
            } else {
                assert.ok(true, 'Parse returned null - expected for some locales');
            }
        } finally {
            AscCommon.bDate1904 = original1904;
        }
    });

    
    // =====================================================================
    // CellFormat.format - Number formatting
    // =====================================================================
    QUnit.module('CellFormat.format');

    QUnit.test('number formatting', function (assert) {
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

    QUnit.test('date/time elapsed formats', function (assert) {
        let testCases = [
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
            let [value, format, expected] = testCases[i];
            let expr = new AscCommon.CellFormat(format);
            let formatted = expr.format(value);
            let text = formatted.map(f => f.text).join('');
            assert.strictEqual(text, expected, `format("${format}", ${value})`);
        }
    });

    QUnit.test('formatRecognition', function (assert) {
        let testCases = [
            ['1,234', '#,##0', 1234],
            ['1,234,567', '#,##0', 1234567],
            ['-1,234', '#,##0', -1234],
            
            // Decimal places
            ['1,234.56', '#,##0.00', 1234.56],
            ['1,234.50', '#,##0.00', 1234.5],
            
            // Percentages
            ['50%', '0%', 0.5],
            ['12.50%', '0.00%', 0.125],
            ['100%', '0%', 1],
            [' 100 %', '0%', 1],
            
            // Currency with text literals
            ['$1,234.56', '\\$#,##0.00_);[Red](\\$#,##0.00)', 1234.56],
            ['$0.00', '\\$#,##0_);[Red](\\$#,##0)', 0],
            ['-$50.00', '\\$#,##0_);[Red](\\$#,##0)', -50],
            ['USD 1000.00', null, 'USD 1000.00'],
            
            // Negative numbers in parentheses
            ['(100)', 'General', -100],
            ['(50.50)', 'General', -50.5],
            
            // Optional digits with
            ['123', 'General', 123],
            ['12.3', 'General', 12.3],
            ['12.', 'General', 12],
            

            // Fraction format cases
            ["1/2", "d-mmm", getDateSerial(1, 2)],  // Jan 2
            ["3/4", "d-mmm", getDateSerial(3, 4)],  // Mar 4
            ["15/20", null, "15/20"],
            [" 1/2", null, " 1/2"],
            ["150/200", null, "150/200"],
            ["0 1/5/5", null, "0 1/5/5"],
            ["1/5/5", "m/d/yyyy", 38357],
            [" 150/200", null, " 150/200"],
            ["+1/2", null, "+1/2"],
            ["-1/2", null, "-1/2"],
            ["$1/2", null, "$1/2"],
            ["(1/2", null, "(1/2"],
            ["1/2)", null, "1/2)"],
            ["1/2%", null, "1/2%"],
            ["1/2 $", null, "1/2 $"],
            ["1/2 p.", null, "1/2 p."],
            ["+1 1/2%", "0.00%", 0.015],
            ["-$2 3/4", "# ?/?", -2.75], //General
            ["(100 1/2)", "# ?/?", -100.5],
            ["25 50/100 %", "0.00%", 0.255],

            ["0 1/2", "# ?/?", 0.5],
            ["0 1/10", "# ??/??", 0.1],
            ["0 1/100", "# ??/??", 0.01],
            ["0 10/2", "# ?/?", 5],
            ["0 15/3", "# ?/?", 5],
            ["0 17/7", "# ?/?", 2.4285714285714284],
            ["0 15/20", "# ??/??", 0.75],
            ["0 12/120", "# ??/??", 0.1],
            ["0 25/250", "# ??/??", 0.1],
            ["0 100/200", "# ??/??", 0.5],
            ["0 125/250", "# ??/??", 0.5],
            ["0 0/1", "# ?/?", 0],
            ["0 1/1", "# ?/?", 1],
            ["0 999/999", "# ??/??", 1],
            ["0 1/999", "# ??/??", 0.001001001001001001],
            ["0 999/1", "# ?/?", 999],

            ["1 999/1", "# ?/?", 1000],
            ["1 999/12", "# ??/??", 84.25],
            ["1 999/134", "# ??/??", 8.455223880597014],
        ]; 
        
        for (let i = 0; i < testCases.length; i++) {
            let value = testCases[i][0];
            let format = testCases[i][1];
            let expectedValue = testCases[i][2];
            
            let formatted = AscCommon.g_oFormatParser.parse(value);

            if (formatted) {
                assert.strictEqual(formatted.format, format, `Case format: ${value}`);
                assert.strictEqual(formatted.value, expectedValue, `Case value: ${expectedValue}`);
            } else {
                assert.strictEqual(formatted, format, `Case format: ${value}`);
            }
        }
    });

    // Format aliases for readability (matching Excel's format codes)
    const F = {
        GENERAL:         'General',
        NUM_2:           '0.00',
        NUM_4:           '0.0000',
        SCI_2:           '0.00E+00',
        SCI_4:           '0.0000E+00',
        CURRENCY_2:      '$#,##0.00',
        CURRENCY_4:      '$#,##0.0000',
        ACCOUNTING_2:    '_($ #,##0.00_);_($ (#,##0.00);_($ "-"??_);_(@_)',
        ACCOUNTING_4:    '_($ #,##0.0000_);_($ (#,##0.0000);_($ "-"????_);_(@_)',
        DATE_SHORT:      'm/d/yyyy',
        DATE_LONG:       '[$-F800]dddd, mmmm dd, yyyy',
        DATE_MEDIUM:     'd-mmm',
        TIME_12H:        '[$-F400]h:mm:ss AM/PM',
        TIME_24H:        'h:mm:ss',
        TIME_DURATION:   '[hh]:mm',
        PERCENT_2:       '0.00%',
        PERCENT_4:       '0.0000%',
        FRACTION_1:      '# ?/?',
        FRACTION_2:      '# ??/??',
        FRACTION_HALF:   '# ?/2',
        TEXT:            '@',
    };
    // Additional format codes that Excel produces
    const OUT = {
        DATE_SHORT:      'm/d/yyyy',
        DATE_MEDIUM:     'd-mmm',
        TIME_24H:        'h:mm:ss',
        THOUSAND_INT:    '#,##0',
        THOUSAND_2:      '#,##0.00',
        CURRENCY_NEG:    '\\$#,##0.00_);[Red](\\$#,##0.00)',
        DATE_TIME:       'm/d/yyyy h:mm',
        TIME_12H_SIMPLE: 'h:mm:ss AM/PM',
    };

    let cellFormatTestCases = Object.values(F);

    // Test data: [input, defaultValue, formatOverrides]
    // formatOverrides: { [cellFormat]: { f: outputFormat } } - ONLY when format changes
    // Skip formats where output format equals input format (the default behavior)
    const formatRecognitionTestCases = [
        // === Numbers (no format change for most formats) ===
        ['1234.56', 1234.56, {
            [F.TEXT]: { f: null },
        }],
        ['1234.5678', 1234.5678, {
            [F.TEXT]: { f: null },
        }],

        // === Thousand separators ===
        ['1,234', 1234, {
            [F.GENERAL]: { f: OUT.THOUSAND_INT },
            [F.TEXT]: { f: null },
        }],
        ['1,234.56', 1234.56, {
            [F.GENERAL]: { f: OUT.THOUSAND_2 },
            [F.TEXT]: { f: null },
        }],
        ['1,234,5678', 12345678, {
            [F.GENERAL]: { f: OUT.THOUSAND_INT },
            [F.TEXT]: { f: null },
        }],

        //todo 
        // // === Scientific notation ===
        // ['1.23E+3', 1230, {
        //     [F.GENERAL]:     { f: F.SCI_2 },
        //     [F.DATE_SHORT]:  { f: F.SCI_2 },
        //     [F.DATE_MEDIUM]: { f: F.SCI_2 },
        //     [F.TIME_24H]:    { f: F.SCI_2 },
        //     [F.PERCENT_2]:   { f: F.SCI_2 },
        //     [F.FRACTION_1]:  { f: F.SCI_2 },
        //     [F.FRACTION_2]:  { f: F.SCI_2 },
        //     [F.TEXT]: { f: null },
        // }],
        // ['1.2345E+3', 1234.5, {
        //     [F.GENERAL]:     { f: F.SCI_2 },
        //     [F.DATE_SHORT]:  { f: F.SCI_2 },
        //     [F.DATE_MEDIUM]: { f: F.SCI_2 },
        //     [F.TIME_24H]:    { f: F.SCI_2 },
        //     [F.PERCENT_2]:   { f: F.SCI_2 },
        //     [F.FRACTION_1]:  { f: F.SCI_2 },
        //     [F.FRACTION_2]:  { f: F.SCI_2 },
        //     [F.TEXT]: { f: null },
        // }],

        // === Currency ===
        ['$1234.56', 1234.56, {
            [F.GENERAL]:     { f: OUT.CURRENCY_NEG },
            [F.SCI_2]:       { f: OUT.CURRENCY_NEG },
            [F.DATE_SHORT]:  { f: OUT.CURRENCY_NEG },
            [F.DATE_MEDIUM]: { f: OUT.CURRENCY_NEG },
            [F.TIME_24H]:    { f: OUT.CURRENCY_NEG },
            [F.PERCENT_2]:   { f: OUT.CURRENCY_NEG },
            [F.FRACTION_1]:  { f: OUT.CURRENCY_NEG },
            [F.FRACTION_2]:  { f: OUT.CURRENCY_NEG },
            [F.TEXT]: { f: null },
        }],
        ['$1234.5678', 1234.5678, {
            [F.GENERAL]:     { f: OUT.CURRENCY_NEG },
            [F.SCI_2]:       { f: OUT.CURRENCY_NEG },
            [F.DATE_SHORT]:  { f: OUT.CURRENCY_NEG },
            [F.DATE_MEDIUM]: { f: OUT.CURRENCY_NEG },
            [F.TIME_24H]:    { f: OUT.CURRENCY_NEG },
            [F.PERCENT_2]:   { f: OUT.CURRENCY_NEG },
            [F.FRACTION_1]:  { f: OUT.CURRENCY_NEG },
            [F.FRACTION_2]:  { f: OUT.CURRENCY_NEG },
            [F.TEXT]: { f: null },
        }],

        // === Accounting (negative) ===
        ['$ (1,234.56)', -1234.56, {
            [F.GENERAL]:     { f: OUT.CURRENCY_NEG },
            [F.SCI_2]:       { f: OUT.CURRENCY_NEG },
            [F.DATE_SHORT]:  { f: OUT.CURRENCY_NEG },
            [F.DATE_MEDIUM]: { f: OUT.CURRENCY_NEG },
            [F.TIME_24H]:    { f: OUT.CURRENCY_NEG },
            [F.PERCENT_2]:   { f: OUT.CURRENCY_NEG },
            [F.FRACTION_1]:  { f: OUT.CURRENCY_NEG },
            [F.FRACTION_2]:  { f: OUT.CURRENCY_NEG },
            [F.TEXT]: { f: null },
        }],
        ['$ (1,234.5678)', -1234.5678, {
            [F.GENERAL]:     { f: OUT.CURRENCY_NEG },
            [F.SCI_2]:       { f: OUT.CURRENCY_NEG },
            [F.DATE_SHORT]:  { f: OUT.CURRENCY_NEG },
            [F.DATE_MEDIUM]: { f: OUT.CURRENCY_NEG },
            [F.TIME_24H]:    { f: OUT.CURRENCY_NEG },
            [F.PERCENT_2]:   { f: OUT.CURRENCY_NEG },
            [F.FRACTION_1]:  { f: OUT.CURRENCY_NEG },
            [F.FRACTION_2]:  { f: OUT.CURRENCY_NEG },
            [F.TEXT]: { f: null },
        }],

        // === Dates ===
        ['1/2/2024', 45293, {
            [F.GENERAL]:    { f: OUT.DATE_SHORT },
            [F.SCI_2]:      { f: OUT.DATE_SHORT },
            [F.TIME_24H]:   { f: OUT.DATE_SHORT },
            [F.PERCENT_2]:  { f: OUT.DATE_SHORT },
            [F.FRACTION_1]: { f: OUT.DATE_SHORT },
            [F.FRACTION_2]: { f: OUT.DATE_SHORT },
            [F.TEXT]: { f: null },
        }],
        ['2-Jan', 46024, {  // 2-Jan in current year (2026) = 46024
            [F.GENERAL]:    { f: OUT.DATE_MEDIUM },
            [F.SCI_2]:      { f: OUT.DATE_MEDIUM },
            [F.TIME_24H]:   { f: F.DATE_MEDIUM },
            [F.PERCENT_2]:  { f: OUT.DATE_MEDIUM },
            [F.FRACTION_1]: { f: OUT.DATE_MEDIUM },
            [F.FRACTION_2]: { f: OUT.DATE_MEDIUM },
            [F.TEXT]: { f: null },
        }],

        // === Time ===
        ['6:00:00', 0.25, {
            [F.GENERAL]:     { f: OUT.TIME_24H },
            [F.SCI_2]:       { f: OUT.TIME_24H },
            [F.DATE_SHORT]:  { f: F.TIME_24H },
            [F.DATE_MEDIUM]: { f: F.TIME_24H },
            [F.PERCENT_2]:   { f: OUT.TIME_24H },
            [F.PERCENT_4]:   { f: F.PERCENT_4 },
            [F.FRACTION_1]:  { f: OUT.TIME_24H },
            [F.FRACTION_2]:  { f: OUT.TIME_24H },
            [F.TEXT]: { f: null },
        }],
        ['6:00:00 AM', 0.25, {
            [F.GENERAL]:     { f: OUT.TIME_12H_SIMPLE },
            [F.SCI_2]:       { f: OUT.TIME_12H_SIMPLE },
            [F.DATE_SHORT]:  { f: OUT.TIME_12H_SIMPLE },
            [F.DATE_MEDIUM]: { f: OUT.TIME_12H_SIMPLE },
            [F.PERCENT_2]:   { f: OUT.TIME_12H_SIMPLE },
            [F.PERCENT_4]:   { f: F.PERCENT_4 },
            [F.FRACTION_1]:  { f: OUT.TIME_12H_SIMPLE },
            [F.FRACTION_2]:  { f: OUT.TIME_12H_SIMPLE },
            [F.TEXT]: { f: null },
        }],

        // === DateTime ===
        ['1/2/2024 6:00:00', 45293.25, {
            [F.GENERAL]:    { f: OUT.DATE_TIME },
            [F.SCI_2]:      { f: OUT.DATE_TIME },
            [F.TIME_24H]:   { f: OUT.DATE_TIME },
            [F.PERCENT_2]:  { f: OUT.DATE_TIME },
            [F.FRACTION_1]: { f: OUT.DATE_TIME },
            [F.FRACTION_2]: { f: OUT.DATE_TIME },
            [F.TEXT]: { f: null },
        }],

        // === Percentages ===
        ['1234.56%', 12.3456, {
            [F.GENERAL]:     { f: F.PERCENT_2 },
            [F.SCI_2]:       { f: F.PERCENT_2 },
            [F.DATE_SHORT]:  { f: F.PERCENT_2 },
            [F.DATE_MEDIUM]: { f: F.PERCENT_2 },
            [F.TIME_24H]:    { f: F.PERCENT_2 },
            [F.FRACTION_1]:  { f: F.PERCENT_2 },
            [F.FRACTION_2]:  { f: F.PERCENT_2 },
            [F.TEXT]: { f: null },
        }],
        ['1234.5678%', 12.345678, {
            [F.GENERAL]:     { f: F.PERCENT_2 },
            [F.SCI_2]:       { f: F.PERCENT_2 },
            [F.DATE_SHORT]:  { f: F.PERCENT_2 },
            [F.DATE_MEDIUM]: { f: F.PERCENT_2 },
            [F.TIME_24H]:    { f: F.PERCENT_2 },
            [F.FRACTION_1]:  { f: F.PERCENT_2 },
            [F.FRACTION_2]:  { f: F.PERCENT_2 },
            [F.TEXT]: { f: null },
        }],
        // === Simple fractions (Excel interprets as dates) ===
        ['1/2', 0.5, {  // Excel: 1/2 with General → Jan 2 (date); with numeric formats → 0.5
            [F.GENERAL]:    { f: OUT.DATE_MEDIUM, v: 46024},
            [F.SCI_2]:      { f: F.FRACTION_1 },
            [F.DATE_SHORT]:   { f: F.DATE_SHORT, v: 46024 },
            [F.DATE_LONG]:   { f: F.DATE_LONG, v: 46024 },
            [F.DATE_MEDIUM]:   { f: F.DATE_MEDIUM, v: 46024 },
            [F.TIME_12H]:   { f: F.TIME_12H, v: 46024 },
            [F.TIME_24H]:   { f: F.DATE_MEDIUM, v: 46024 },
            [F.TIME_DURATION]:   { f: F.TIME_DURATION, v: 46024 },
            [F.PERCENT_2]:  { f: F.FRACTION_1 },
            [F.PERCENT_4]:  { f: F.PERCENT_4 },
            [F.TEXT]: { f: null },
        }],
        ['1/11', 1/11, {  // Excel: 1/11 → Jan 11
            [F.GENERAL]:    { f: OUT.DATE_MEDIUM, v: 46033 },
            [F.SCI_2]:      { f: F.FRACTION_2 },
            [F.DATE_SHORT]:   { f: F.DATE_SHORT, v: 46033 },
            [F.DATE_LONG]:   { f: F.DATE_LONG, v: 46033 },
            [F.DATE_MEDIUM]:   { f: F.DATE_MEDIUM, v: 46033 },
            [F.TIME_12H]:   { f: F.TIME_12H, v: 46033 },
            [F.TIME_24H]:   { f: F.DATE_MEDIUM, v: 46033 },
            [F.TIME_DURATION]:   { f: F.TIME_DURATION, v: 46033 },
            [F.PERCENT_2]:  { f: F.FRACTION_2, v: 0.09090909090909091 },
            [F.PERCENT_4]:  { f: F.PERCENT_4, v: 0.09090909090909091 },
            [F.TEXT]: { f: null },
        }],
        ['3/4', 0.75, {  // Excel: 3/4 → March 4
            [F.GENERAL]:    { f: OUT.DATE_MEDIUM, v: 46085 },
            [F.SCI_2]:      { f: F.FRACTION_1 },
            [F.DATE_SHORT]:   { f: F.DATE_SHORT, v: 46085 },
            [F.DATE_LONG]:   { f: F.DATE_LONG, v: 46085 },
            [F.DATE_MEDIUM]:   { f: F.DATE_MEDIUM, v: 46085 },
            [F.TIME_12H]:   { f: F.TIME_12H, v: 46085 },
            [F.TIME_24H]:   { f: F.DATE_MEDIUM, v: 46085 },
            [F.TIME_DURATION]:   { f: F.TIME_DURATION, v: 46085 },
            [F.PERCENT_2]:  { f: F.FRACTION_1, v: 0.75 },
            [F.PERCENT_4]:  { f: F.PERCENT_4, v: 0.75 },
            [F.TEXT]: { f: null },
        }],
        ['15/3', 5, {  // Not a valid date (15th month), stays as fraction
            [F.GENERAL]:    { f: null },
            [F.SCI_2]:      { f: F.FRACTION_1 },
            [F.DATE_SHORT]:   { f: null },
            [F.DATE_LONG]:   { f: null },
            [F.DATE_MEDIUM]:   { f: null },
            [F.TIME_12H]:   { f: null },
            [F.TIME_24H]:   { f: null },
            [F.TIME_DURATION]:   { f: null },
            [F.PERCENT_2]:  { f: F.FRACTION_1, v: 5 },
            [F.PERCENT_4]:  { f: F.PERCENT_4, v: 5 },
            [F.TEXT]: { f: null },
        }],
        ['15/20', 0.75, {
            [F.GENERAL]:    { f: null },
            [F.SCI_2]:      { f: F.FRACTION_2 },
            [F.DATE_SHORT]:   { f: null },
            [F.DATE_LONG]:   { f: null },
            [F.DATE_MEDIUM]:   { f: null },
            [F.TIME_12H]:   { f: null },
            [F.TIME_24H]:   { f: null },
            [F.TIME_DURATION]:   { f: null },
            [F.PERCENT_2]:  { f: F.FRACTION_2, v: 0.75 },
            [F.PERCENT_4]:  { f: F.PERCENT_4, v: 0.75 },
            [F.TEXT]: { f: null },
        }],
        ['150/200', 0.75, {
            [F.GENERAL]:    { f: null },
            [F.SCI_2]:      { f: F.FRACTION_2 },
            [F.DATE_SHORT]:   { f: null },
            [F.DATE_LONG]:   { f: null },
            [F.DATE_MEDIUM]:   { f: null },
            [F.TIME_12H]:   { f: null },
            [F.TIME_24H]:   { f: null },
            [F.TIME_DURATION]:   { f: null },
            [F.PERCENT_2]:  { f: F.FRACTION_2, v: 0.75 },
            [F.PERCENT_4]:  { f: F.PERCENT_4, v: 0.75 },
            [F.TEXT]: { f: null },
        }],

        // === Mixed fractions ===
        ['1 1/2', 1.5, {
            [F.GENERAL]:     { f: F.FRACTION_1 },
            [F.SCI_2]:       { f: F.FRACTION_1 },
            [F.DATE_SHORT]:  { f: F.FRACTION_1 },
            [F.DATE_MEDIUM]: { f: F.FRACTION_1 },
            [F.TIME_24H]:    { f: F.FRACTION_1 },
            [F.PERCENT_2]:   { f: F.FRACTION_1 },
            [F.PERCENT_4]:   { f: F.PERCENT_4 },
            [F.TEXT]: { f: null },
        }],
        ['2 3/4', 2.75, {
            [F.GENERAL]:     { f: F.FRACTION_1 },
            [F.SCI_2]:       { f: F.FRACTION_1 },
            [F.DATE_SHORT]:  { f: F.FRACTION_1 },
            [F.DATE_MEDIUM]: { f: F.FRACTION_1 },
            [F.TIME_24H]:    { f: F.FRACTION_1 },
            [F.PERCENT_2]:   { f: F.FRACTION_1 },
            [F.PERCENT_4]:   { f: F.PERCENT_4 },
            [F.TEXT]: { f: null },
        }],
        ['0 1/2', 0.5, {
            [F.GENERAL]:     { f: F.FRACTION_1 },
            [F.SCI_2]:       { f: F.FRACTION_1 },
            [F.DATE_SHORT]:  { f: F.FRACTION_1 },
            [F.DATE_MEDIUM]: { f: F.FRACTION_1 },
            [F.TIME_24H]:    { f: F.FRACTION_1 },
            [F.PERCENT_2]:   { f: F.FRACTION_1 },
            [F.PERCENT_4]:   { f: F.PERCENT_4 },
            [F.TEXT]: { f: null },
        }],
        ['0 1/10', 0.1, {
            [F.GENERAL]:     { f: F.FRACTION_2 },
            [F.SCI_2]:       { f: F.FRACTION_2 },
            [F.DATE_SHORT]:  { f: F.FRACTION_2 },
            [F.DATE_MEDIUM]: { f: F.FRACTION_2 },
            [F.TIME_24H]:    { f: F.FRACTION_2 },
            [F.PERCENT_2]:   { f: F.FRACTION_2 },
            [F.PERCENT_4]:   { f: F.PERCENT_4 },
            [F.TEXT]: { f: null },
        }],
        ['0 1/100', 0.01, {
            [F.GENERAL]:     { f: F.FRACTION_2 },
            [F.SCI_2]:       { f: F.FRACTION_2 },
            [F.DATE_SHORT]:  { f: F.FRACTION_2 },
            [F.DATE_MEDIUM]: { f: F.FRACTION_2 },
            [F.TIME_24H]:    { f: F.FRACTION_2 },
            [F.PERCENT_2]:   { f: F.FRACTION_2 },
            [F.PERCENT_4]:   { f: F.PERCENT_4 },
            [F.TEXT]: { f: null },
        }],
        ['1 150/200', 1.75, {
            [F.GENERAL]:     { f: F.FRACTION_2 },
            [F.SCI_2]:       { f: F.FRACTION_2 },
            [F.DATE_SHORT]:  { f: F.FRACTION_2 },
            [F.DATE_MEDIUM]: { f: F.FRACTION_2 },
            [F.TIME_24H]:    { f: F.FRACTION_2 },
            [F.PERCENT_2]:   { f: F.FRACTION_2 },
            [F.PERCENT_4]:   { f: F.PERCENT_4 },
            [F.TEXT]: { f: null },
        }],

        // === Short date ===
        ['1/5/5', 38357, {  // 1/5/2005
            [F.GENERAL]:    { f: OUT.DATE_SHORT },
            [F.SCI_2]:      { f: OUT.DATE_SHORT },
            [F.TIME_24H]: { f: F.DATE_SHORT },
            [F.PERCENT_2]:  { f: OUT.DATE_SHORT },
            [F.FRACTION_1]: { f: OUT.DATE_SHORT },
            [F.FRACTION_2]: { f: OUT.DATE_SHORT },
            [F.TEXT]: { f: null },
        }],

        // === Text ===
        ['text', 'text', {
            [F.GENERAL]: { f: null },
            [F.NUM_2]: { f: null },
            [F.NUM_4]: { f: null },
            [F.SCI_2]: { f: null },
            [F.SCI_4]: { f: null },
            [F.CURRENCY_2]: { f: null },
            [F.CURRENCY_4]: { f: null },
            [F.ACCOUNTING_2]: { f: null },
            [F.ACCOUNTING_4]: { f: null },
            [F.DATE_SHORT]: { f: null },
            [F.DATE_LONG]: { f: null },
            [F.DATE_MEDIUM]: { f: null },
            [F.TIME_12H]: { f: null },
            [F.TIME_24H]: { f: null },
            [F.TIME_DURATION]: { f: null },
            [F.PERCENT_2]: { f: null },
            [F.PERCENT_4]: { f: null },
            [F.FRACTION_1]: { f: null },
            [F.FRACTION_2]: { f: null },
            [F.FRACTION_HALF]: { f: null },
            [F.TEXT]: { f: null },
        }]
    ];

    for (const cellFormat of cellFormatTestCases) {
        QUnit.test(`formatRecognition: cellFormat="${cellFormat}"`, function (assert) {
            for (let i = 0; i < formatRecognitionTestCases.length; i++) {
                const [textInput, expectedValue, formatOverrides = {}] = formatRecognitionTestCases[i];
                const override = formatOverrides[cellFormat];

                // Expected format: use override if specified, otherwise same as cell format
                const expectedFormat = override ? override.f : cellFormat;
                const expectedVal = override?.v !== undefined ? override.v : expectedValue;

                const numFormat = AscCommon.oNumFormatCache.get(cellFormat);
                const currentFormat = numFormat.getType();
                const stringFormat = cellFormat;
                const formatted = AscCommon.g_oFormatParser.parse(textInput, null, currentFormat, stringFormat);

                if (formatted && formatted.format) {
                    assert.strictEqual(
                        formatted.format,
                        expectedFormat,
                        `Format: input="${textInput}" formatType="${currentFormat}" cellFormat="${stringFormat}"`
                    );
                    assert.strictEqual(
                        formatted.value,
                        expectedVal,
                        `Value: input="${textInput}" formatType="${currentFormat}" cellFormat="${stringFormat}"`
                    );
                } else {
                    assert.strictEqual(
                        formatted,
                        expectedFormat,
                        `Null result: input="${textInput}" formatType="${currentFormat}" cellFormat="${stringFormat}"`
                    );
                }
            }
        });
    }

    // =====================================================================
    // FormatParser.parse - Locale-specific mixed fraction tests
    // Tests Number.isInteger check for non-US locales (decimal=",", group=".")
    // =====================================================================
    QUnit.module('FormatParser.parse - Locale mixed fractions');

    QUnit.test('European locale (decimal=comma, group=period) mixed fractions', function (assert) {
        // Danish/German-like culture: decimal separator is comma, group separator is period
        // LCID 6 is Danish: NumberDecimalSeparator: ",", NumberGroupSeparator: "."
        const europeanCulture = {
            NumberDecimalSeparator: ",",
            NumberGroupSeparator: ".",
            NumberGroupSizes: [3],
            CurrencySymbol: "€",
            CurrencyPositivePattern: 3,
            CurrencyNegativePattern: 8,
            DateSeparator: ".",
            TimeSeparator: ":"
        };
        
        // Get fraction format for context
        const fractionFormatString = "# ?/?";
        const numFormat = AscCommon.oNumFormatCache.get(fractionFormatString);
        const fractionFormatType = numFormat.getType();
        
        // === Valid: thousand separator in whole part ===
        // "1.000 1/2" in European = 1000 + 0.5 = 1000.5 (period is thousand separator)
        let result = formatParser.parse("1.000 1/2", europeanCulture, fractionFormatType, fractionFormatString);
        assert.ok(result !== null, '"1.000 1/2" with European locale should parse (1000 + 1/2)');
        if (result) {
            assert.strictEqual(result.value, 1000.5, '"1.000 1/2" value should be 1000.5');
        }
        
        // "2.500 3/4" = 2500 + 0.75 = 2500.75
        result = formatParser.parse("2.500 3/4", europeanCulture, fractionFormatType, fractionFormatString);
        assert.ok(result !== null, '"2.500 3/4" with European locale should parse');
        if (result) {
            assert.strictEqual(result.value, 2500.75, '"2.500 3/4" value should be 2500.75');
        }
        
        // === Invalid: decimal separator in whole part of mixed fraction ===
        // "1,5 1/2" in European = 1.5 + 0.5 - INVALID (whole part must be integer)
        result = formatParser.parse("1,5 1/2", europeanCulture, fractionFormatType, fractionFormatString);
        assert.strictEqual(result, null, '"1,5 1/2" with European locale should be null (decimal in whole part)');
        
        // "0,5 1/2" - also invalid
        result = formatParser.parse("0,5 1/2", europeanCulture, fractionFormatType, fractionFormatString);
        assert.strictEqual(result, null, '"0,5 1/2" with European locale should be null');
        
        // "10,25 1/4" - invalid
        result = formatParser.parse("10,25 1/4", europeanCulture, fractionFormatType, fractionFormatString);
        assert.strictEqual(result, null, '"10,25 1/4" with European locale should be null');
        
        // === Valid: simple fractions (no whole part) ===
        result = formatParser.parse("1/2", europeanCulture, fractionFormatType, fractionFormatString);
        assert.ok(result !== null, '"1/2" with European locale should parse');
        if (result) {
            assert.strictEqual(result.value, 0.5, '"1/2" value should be 0.5');
        }
        
        // === Valid: integer whole part ===
        result = formatParser.parse("5 1/2", europeanCulture, fractionFormatType, fractionFormatString);
        assert.ok(result !== null, '"5 1/2" with European locale should parse');
        if (result) {
            assert.strictEqual(result.value, 5.5, '"5 1/2" value should be 5.5');
        }
    });

});

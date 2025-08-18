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
            assert.ok(date, `Parsed should not be null for: ${data[i][0]}`);
            if (!date) { continue; }
            assert.strictEqual(date.format, data[i][1], `Case format: ${data[i][0]}`);
            assert.strictEqual(Math.abs(date.value - data[i][2]) < eps, true, `Case value: ${data[i][0]}`);
        }
    });
    QUnit.test('parseDate_invalidCases', function (assert) {
        const invalids = [
            "11 AM extra",           // AM/PM not last
            "10:75",                 // minutes > 59
            "10:00:60",              // seconds >= 60
            "1-2-3 10",              // stray digits outside date/time
            "12345/1/1",             // digit token too long
            "Octember 5, 2010",      // invalid month name
            "abc",                    // non-date text
            "11AM",                   // AM/PM must be separated by space
            "1:2:3:4",               // too many time parts
            "25 PM",                  // hour > 23 with AM/PM
            "AM",                     // AM/PM without time
            "1/2/2000, 11:00",       // comma before time prevents ':' from being time
            "2/30/2000",             // invalid day for month
            "13/1/2000",             // invalid month > 12
            "0/1/2000",              // invalid month 0
            "1/0/2000",              // invalid day 0
            "1/2-2000",              // mixed separators
            "1/2/999",               // 3-digit year should be invalid
            "29/2/2000",             // invalid under default m/d culture (month 29)
            "1/2/2000.",             // stray punctuation
            "10-00",                 // invalid time separator
            "2000-01-02T11:00",      // ISO T separator not supported
            ", 1/2/2000",            // leading comma
            "PM 1:00",               // AM/PM at start
            "1//2000",               // empty token
            "",                      // empty string
            "1/2/2000 UTC"           // timezone suffix not supported
        ];
        for (let i = 0; i < invalids.length; i++) {
            const parsed = AscCommon.g_oFormatParser.parse(invalids[i]);
            assert.strictEqual(parsed, null, `Should be null for: ${invalids[i]}`);
        }
    });
    QUnit.test('parseDate_allBranches', function (assert) {
        // Cover Y-M-D path, month names, two-digit years, special 1900 rules, AM/PM, and long-hour times
        const cases = [
            // Y-M-D numeric first (year > 1000 triggers direct mapping using culture short format)
            ["2000/1/2", "m/d/yyyy", 36527],

            // Month-Year only (numeric) => mmm-yy, defaults day=1
            ["1/2000", "mmm-yy", 36526],
            ["1-2000", "mmm-yy", 36526],

            // Special 1900 leap-day compatibility and baseline
            ["2/29/1900", "m/d/yyyy", 60],
            ["12/31/1899", "m/d/yyyy", 0],

            // Month names (abbr and full), with comma and without; includes AM/PM on date-time
            ["Oct 11, 2008", "dd-mmm-yy", 39732],
            ["October 11, 2008", "dd-mmm-yy", 39732],
            ["Sep. 5, 2010", "dd-mmm-yy", 40426],
            ["Oct 11 2008", "dd-mmm-yy", 39732],
            ["Sep 5 2010", "dd-mmm-yy", 40426],
            ["Sep. 5 2010", "dd-mmm-yy", 40426],
            ["September 2008", "mmm-yy", 39692],
            ["Jan 2000", "mmm-yy", 36526],
            ["Oct 11, 2008 3:30 PM", "dd-mmm-yy h:mm", 39732 + 15.5 / 24],
            ["Oct 11, 2008 03:30:05 PM", "dd-mmm-yy h:mm:ss", 39732 + 15 / 24 + 30 / 1440 + 5 / 86400],

            // Day Month Year textual order (month in the middle)
            ["1 Oct 2008", "d-mmm-yy", 39722],
            ["Oct 2008", "mmm-yy", 39722],
            ["31 Dec 1999", "d-mmm-yy", 36525],

            // Two-digit years normalization
            ["1/2/00", "m/d/yyyy", 36527],   // 2000-01-02
            ["1/2/05", "m/d/yyyy", 38354],   // 2005-01-02
            ["1/2/75", "m/d/yyyy", 27396],   // 1975-01-02

            // AM/PM time-only variants
            ["12 PM", "h:mm AM/PM", 12 / 24],
            ["12 AM", "h:mm AM/PM", 0],
            ["1:00 PM", "h:mm AM/PM", 13 / 24],
            ["12:00 AM", "h:mm AM/PM", 0],
            ["1:02:03 PM", "h:mm:ss AM/PM", 13 / 24 + 2 / 1440 + 3 / 86400],

            // Long hours => [h]:mm and add :ss when dValue > 1
            ["48:00", "[h]:mm:ss", 2],

            // Fine-grained time-only
            ["0:0:5", "h:mm:ss", 5 / 86400],
            ["0:0", "h:mm", 0],
            ["1:2", "h:mm", 1 / 24 + 2 / 1440],
            ["11:00:05", "h:mm:ss", 11 / 24 + 5 / 86400],
            ["0:0:0", "h:mm:ss", 0],
            ["25:00:01", "[h]:mm:ss", 25 / 24 + 1 / 86400],

            // Leap-day valid in leap year
            ["Feb 29, 2000", "dd-mmm-yy", 36585]
        ];
        for (let i = 0; i < cases.length; i++) {
            const input = cases[i][0];
            const expectedFormat = cases[i][1];
            const expectedValue = cases[i][2];
            const parsed = AscCommon.g_oFormatParser.parse(input);
            assert.ok(parsed, `Parsed should not be null for: ${input}`);
            if (!parsed) { continue; }
            assert.strictEqual(parsed.format, expectedFormat, `Format for: ${input}`);
            assert.strictEqual(Math.abs(parsed.value - expectedValue) < eps, true, `Value for: ${input}`);
        }
    });
    QUnit.test('parseDate_moreCases', function (assert) {
        // Additional coverage to guide refactoring of FormatParser.parseDate
        let data = [
            // Date-only inputs
            ["1/2/2000", "m/d/yyyy", 36527],
            ["1-2-2000", "m/d/yyyy", 36527],

            // Date-time with AM/PM (format should remain without AM/PM for date-time)
            ["1/2/2000 11:00 PM", "m/d/yyyy h:mm", 36527 + 23 / 24],
            ["1/2/2000 12:00 AM", "m/d/yyyy h:mm", 36527 + 0 / 24],

            // Time-only additional variants
            ["10:00:00", "h:mm:ss", 10 / 24],
            ["0:00", "h:mm", 0],
            ["55:00", "[h]:mm:ss", 55 / 24],

            // 24-hour boundary should switch to [h]:mm:ss and equal 1 day
            ["24:00", "[h]:mm:ss", 1],
            ["24:00:00", "[h]:mm:ss", 1],

            // Whitespace tolerance
            [" 1/2/2000 ", "m/d/yyyy", 36527],
            ["1/2/2000   11:00", "m/d/yyyy h:mm", 36527 + 11 / 24]
        ];
        for (let i = 0; i < data.length; i++) {
            let date = AscCommon.g_oFormatParser.parse(data[i][0]);
            assert.ok(date, `Parsed should not be null for: ${data[i][0]}`);
            if (!date) { continue; }
            assert.strictEqual(date.format, data[i][1], `Case format: ${data[i][0]}`);
            assert.strictEqual(Math.abs(date.value - data[i][2]) < eps, true, `Case value: ${data[i][0]}`);
        }
    });
    QUnit.test('parseWithFormat_strict', function (assert) {
        const ci = AscCommon.g_oDefaultCultureInfo;

        /**
         * Compute default-mode Excel serial for a given UTC date/time (no timezone issues).
         * Mirrors default mode epoch in FormatParser._buildDateTimeResult.
         * @param {number} y
         * @param {number} m1 Month 1-12
         * @param {number} d1
         * @param {number} [hh]
         * @param {number} [mm]
         * @param {number} [ss]
         */
        const excelSerial = function (y, m1, d1, hh, mm, ss) {
            const ms = Date.UTC(y, (m1 - 1), d1, hh || 0, mm || 0, ss || 0) - Date.UTC(1899, 11, 30, 0, 0, 0);
            return ms / (86400 * 1000);
        };

        // Success cases (full input must be consumed)
        const okCases = [
            { inp: "02/03/2004", fmt: "dd/mm/yyyy", expFmt: "dd/mm/yyyy", expVal: excelSerial(2004, 3, 2, 0, 0, 0) },
            { inp: "2/3/2004 11:34", fmt: "d/m/yyyy h:mm", expFmt: "d/m/yyyy h:mm", expVal: excelSerial(2004, 3, 2, 11, 34, 0) },
            { inp: "11:34 AM", fmt: "h:mm AM/PM", expFmt: "h:mm AM/PM", expVal: (11 / 24) + (34 / 1440) },
            { inp: "55:34", fmt: "[h]:mm", expFmt: "[h]:mm:ss", expVal: (55 / 24) + (34 / 1440) },
            { inp: "Sep 5, 2010", fmt: "mmm d, yyyy", expFmt: "mmm d, yyyy", expVal: excelSerial(2010, 9, 5, 0, 0, 0) },
            { inp: "Sep. 5, 2010", fmt: "mmm. d, yyyy", expFmt: "mmm. d, yyyy", expVal: excelSerial(2010, 9, 5, 0, 0, 0) }
        ];
        for (let i = 0; i < okCases.length; i++) {
            const c = okCases[i];
            const r = AscCommon.g_oFormatParser.parseWithFormat(c.inp, c.fmt, ci, { mode: "default" });
            assert.ok(r, `parseWithFormat should succeed: ${c.inp} ~ ${c.fmt}`);
            if (!r) { continue; }
            assert.strictEqual(r.format, c.expFmt, `Format echo for: ${c.inp}`);
            assert.strictEqual(Math.abs(r.value - c.expVal) < eps, true, `Value for: ${c.inp}`);
        }

        // Two-digit year expansion with custom window
        const rYY = AscCommon.g_oFormatParser.parseWithFormat("02/03/04", "dd/mm/yy", ci, { mode: "default", yearWindowStart: 2000 });
        assert.ok(rYY, "Two-digit year parse should succeed");
        if (rYY) {
            assert.strictEqual(rYY.format, "dd/mm/yy", "Two-digit year format preserved");
            assert.strictEqual(Math.abs(rYY.value - excelSerial(2004, 3, 2, 0, 0, 0)) < eps, true, "Two-digit year expanded using windowStart");
        }

        // Failure cases (strict full-consumption)
        const bad = [
            { inp: "02/03/2004x", fmt: "dd/mm/yyyy" },      // trailing garbage
            { inp: " 02/03/2004", fmt: "dd/mm/yyyy" },      // leading space not in format
            { inp: "02/03/2004", fmt: "dd/mm/yyyy " },      // missing trailing literal space
            { inp: "1", fmt: "dd" },                        // dd requires 2 digits
            { inp: "11:34", fmt: "h:mm AM/PM" },            // AM/PM missing
            { inp: "11:34 amx", fmt: "h:mm AM/PM" }         // extra after AM/PM
        ];
        for (let i = 0; i < bad.length; i++) {
            const b = bad[i];
            const r = AscCommon.g_oFormatParser.parseWithFormat(b.inp, b.fmt, ci, { mode: "default" });
            assert.strictEqual(r, null, `parseWithFormat should fail: ${b.inp} ~ ${b.fmt}`);
        }
    });
});

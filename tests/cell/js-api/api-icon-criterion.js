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
    var ws = AscTest.JsApi.GetActiveSheet();

    // ======= TEST HELPERS =======
    const initializeTest = function () {
        const r = ws.GetRange("A1:Z100");
        r.Clear();
    };

    const createIconCondition = function (rangeAddress) {
        const range = ws.GetRange(rangeAddress);
        const formatConditions = range.GetFormatConditions();
        const iconCondition = formatConditions.AddIconSetCondition();

        return {
            range: range,
            formatConditions: formatConditions,
            iconCondition: iconCondition
        };
    };

    const getCriteria = function (rangeAddress) {
        const ctx = createIconCondition(rangeAddress);
        const criteria = ctx.iconCondition.GetIconCriteria();

        return {
            range: ctx.range,
            iconCondition: ctx.iconCondition,
            criteria: criteria
        };
    };

    // ======= TESTS =======

    QUnit.module("ApiIconCriterion — smoke");

    QUnit.test("GetIconCriteria returns criteria collection", function (assert) {
        initializeTest();

        const ctx = getCriteria("A1");
        assert.ok(ctx.iconCondition, "Icon set condition exists");
        assert.ok(Array.isArray(ctx.criteria), "Criteria is an array");
        assert.ok(ctx.criteria.length > 0, "Criteria array is not empty");
        assert.ok(ctx.criteria[0], "First criterion exists");
    });

    QUnit.test("GetIndex / Index return 1-based indexes", function (assert) {
        initializeTest();

        const ctx = getCriteria("B1");
        const criteria = ctx.criteria;

        assert.ok(criteria.length > 0, "Criteria exist");
        assert.strictEqual(criteria[0].GetIndex(), 1, "First criterion GetIndex is 1");
        assert.strictEqual(criteria[0].Index, 1, "First criterion Index property is 1");

        if (criteria.length > 1) {
            assert.strictEqual(criteria[1].GetIndex(), 2, "Second criterion GetIndex is 2");
            assert.strictEqual(criteria[1].Index, 2, "Second criterion Index property is 2");
        }
    });

    QUnit.module("ApiIconCriterion — Operator");

    QUnit.test("SetOperator / GetOperator roundtrip on first criterion", function (assert) {
        initializeTest();

        const ctx = getCriteria("C1");
        const firstCriterion = ctx.criteria[0];

        firstCriterion.SetOperator("xlGreater");
        assert.strictEqual(firstCriterion.GetOperator(), "xlGreater", "Operator changed to xlGreater");

        firstCriterion.SetOperator("xlGreaterEqual");
        assert.strictEqual(firstCriterion.GetOperator(), "xlGreaterEqual", "Operator changed back to xlGreaterEqual");
    });

    QUnit.test("Operator property setter/getter roundtrip", function (assert) {
        initializeTest();

        const ctx = getCriteria("D1");
        const firstCriterion = ctx.criteria[0];

        firstCriterion.Operator = "xlGreater";
        assert.strictEqual(firstCriterion.Operator, "xlGreater", "Operator property returns xlGreater");

        firstCriterion.Operator = "xlGreaterEqual";
        assert.strictEqual(firstCriterion.Operator, "xlGreaterEqual", "Operator property returns xlGreaterEqual");
    });

    QUnit.module("ApiIconCriterion — Type");

    QUnit.test("GetType returns a value for non-empty criterion", function (assert) {
        initializeTest();

        const ctx = getCriteria("E1");
        const criteria = ctx.criteria;
        const target = criteria.length > 1 ? criteria[1] : criteria[0];

        assert.ok(target.GetType() !== null, "GetType returns non-null");
    });

    QUnit.test("SetType / GetType roundtrip on second criterion", function (assert) {
        initializeTest();

        const ctx = getCriteria("F1");
        const criteria = ctx.criteria;

        assert.ok(criteria.length > 1, "Need at least two criteria for SetType test");

        const secondCriterion = criteria[1];
        secondCriterion.SetType("xlConditionValuePercent");
        assert.strictEqual(secondCriterion.GetType(), "xlConditionValuePercent", "Type changed to percent");

        secondCriterion.SetType("xlConditionValueNumber");
        assert.strictEqual(secondCriterion.GetType(), "xlConditionValueNumber", "Type changed to number");

        secondCriterion.SetType("xlConditionValueFormula");
        assert.strictEqual(secondCriterion.GetType(), "xlConditionValueFormula", "Type changed to formula");
    });

    QUnit.test("Type property setter/getter roundtrip on second criterion", function (assert) {
        initializeTest();

        const ctx = getCriteria("G1");
        const criteria = ctx.criteria;

        assert.ok(criteria.length > 1, "Need at least two criteria for Type property test");

        const secondCriterion = criteria[1];
        secondCriterion.Type = "xlConditionValuePercentile";
        assert.strictEqual(secondCriterion.Type, "xlConditionValuePercentile", "Type property changed to percentile");
    });

    QUnit.test("SetType does not change first criterion", function (assert) {
        initializeTest();

        const ctx = getCriteria("H1");
        const firstCriterion = ctx.criteria[0];
        const before = firstCriterion.GetType();

        firstCriterion.SetType("xlConditionValuePercent");
        assert.strictEqual(firstCriterion.GetType(), before, "First criterion type remains unchanged");
    });

    QUnit.module("ApiIconCriterion — Value");

    QUnit.test("GetValue returns a value for criterion", function (assert) {
        initializeTest();

        const ctx = getCriteria("I1");
        const criteria = ctx.criteria;
        const target = criteria.length > 1 ? criteria[1] : criteria[0];

        assert.ok(target.GetValue() !== null, "GetValue returns non-null");
    });

    QUnit.test("SetValue / GetValue roundtrip on second criterion", function (assert) {
        initializeTest();

        const ctx = getCriteria("J1");
        const criteria = ctx.criteria;

        assert.ok(criteria.length > 1, "Need at least two criteria for SetValue test");

        const secondCriterion = criteria[1];

        secondCriterion.SetType("xlConditionValueNumber");
        secondCriterion.SetValue("25");
        assert.strictEqual(String(secondCriterion.GetValue()), "25", "Value changed to 25");

        secondCriterion.SetValue("50");
        assert.strictEqual(String(secondCriterion.GetValue()), "50", "Value changed to 50");
    });

    QUnit.test("Value property setter/getter roundtrip on second criterion", function (assert) {
        initializeTest();

        const ctx = getCriteria("K1");
        const criteria = ctx.criteria;

        assert.ok(criteria.length > 1, "Need at least two criteria for Value property test");

        const secondCriterion = criteria[1];

        secondCriterion.Type = "xlConditionValueNumber";
        secondCriterion.Value = "77";
        assert.strictEqual(String(secondCriterion.Value), "77", "Value property changed to 77");
    });

    QUnit.module("ApiIconCriterion — Icon");

    QUnit.test("GetIcon returns a value for criterion", function (assert) {
        initializeTest();

        const ctx = getCriteria("L1");
        const criteria = ctx.criteria;
        const target = criteria.length > 1 ? criteria[1] : criteria[0];

        assert.ok(target.GetIcon() !== null, "GetIcon returns non-null");
    });

    QUnit.test("SetIcon / GetIcon roundtrip on first criterion", function (assert) {
        initializeTest();
        console.log("here");
        const ctx = getCriteria("M1");
        const firstCriterion = ctx.criteria[0];

        firstCriterion.SetIcon("xlIconGoldStar");
        assert.strictEqual(firstCriterion.GetIcon(), "xlIconGoldStar", "Icon changed to xlIconGoldStar");

        firstCriterion.SetIcon("xlIconGreenCheck");
        assert.strictEqual(firstCriterion.GetIcon(), "xlIconGreenCheck", "Icon changed to xlIconGreenCheck");
    });

    QUnit.test("Icon property setter/getter roundtrip", function (assert) {
        initializeTest();

        const ctx = getCriteria("N1");
        const firstCriterion = ctx.criteria[0];

        firstCriterion.Icon = "xlIconRedFlag";
        assert.strictEqual(firstCriterion.Icon, "xlIconRedFlag", "Icon property changed to xlIconRedFlag");
    });

    QUnit.module("ApiIconCriterion — combined live updates");

    QUnit.test("Sequential updates on same criterion are read from live rule data", function (assert) {
        initializeTest();

        const ctx = getCriteria("O1");
        const criteria = ctx.criteria;

        assert.ok(criteria.length > 1, "Need at least two criteria for combined test");

        const secondCriterion = criteria[1];

        secondCriterion.SetType("xlConditionValueNumber");
        secondCriterion.SetValue("33");
        secondCriterion.SetOperator("xlGreater");
        secondCriterion.SetIcon("xlIconGreenCheck");

        assert.strictEqual(secondCriterion.GetType(), "xlConditionValueNumber", "Type is current");
        assert.strictEqual(String(secondCriterion.GetValue()), "33", "Value is current");
        assert.strictEqual(secondCriterion.GetOperator(), "xlGreater", "Operator is current");
        assert.strictEqual(secondCriterion.GetIcon(), "xlIconGreenCheck", "Icon is current");
    });

    QUnit.module("ApiIconCriterion — invalid input safety");

    QUnit.test("Invalid SetType input does not break current type", function (assert) {
        initializeTest();

        const ctx = getCriteria("P1");
        const criteria = ctx.criteria;

        assert.ok(criteria.length > 1, "Need at least two criteria for invalid SetType test");

        const secondCriterion = criteria[1];
        secondCriterion.SetType("xlConditionValueNumber");
        const before = secondCriterion.GetType();

        secondCriterion.SetType("xlConditionValue__Bogus");
        assert.strictEqual(secondCriterion.GetType(), before, "Invalid type does not change current type");
    });

    QUnit.test("Invalid SetOperator input does not throw", function (assert) {
        initializeTest();

        const ctx = getCriteria("Q1");
        const firstCriterion = ctx.criteria[0];

        try {
            firstCriterion.SetOperator("xlOperator__Bogus");
            assert.ok(true, "Invalid operator did not throw");
        } catch (e) {
            assert.ok(false, "Invalid operator should not throw");
        }
    });

    QUnit.test("Invalid SetIcon input does not break current icon", function (assert) {
        initializeTest();

        const ctx = getCriteria("R1");
        const firstCriterion = ctx.criteria[0];
        const before = firstCriterion.GetIcon();

        firstCriterion.SetIcon("xlIcon__Bogus");
        assert.strictEqual(firstCriterion.GetIcon(), before, "Invalid icon does not change current icon");
    });
});
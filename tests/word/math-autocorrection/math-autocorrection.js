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

    let Root, MathContent, logicDocument;

    function Init() {
		logicDocument = AscTest.CreateLogicDocument();
        logicDocument.RemoveFromContent(0, logicDocument.GetElementsCount(), false);

        let p1 = new AscWord.Paragraph();
        logicDocument.AddToContent(0, p1);

        MathContent = new ParaMath();

        if (p1.Content.length > 0)
            p1.Content.splice(0, 1);

        p1.AddToContent(0, MathContent);
        Root = MathContent.Root;
	};
	Init();

    function Clear() {
        Root.Remove_FromContent(0, Root.Content.length);
        Root.Correct_Content();
	};
	function AddText(str)
	{
		let one = str.getUnicodeIterator();

		while (one.isInside()) {
			let oElement = new AscWord.CRunText(one.value());
			MathContent.Add(oElement);
			one.next();
		}
	};
	function Test(str, intCurPos, arrResult, isLaTeX, strNameOfTest, isConvertAfter, isGetIntDifferentForm)
	{
		let nameOfTest = strNameOfTest ? strNameOfTest + " \'" + str + "\'" : str;

		QUnit.test(nameOfTest, function (assert)
		{
			if (isLaTeX)
				logicDocument.SetMathInputType(1);
			else
				logicDocument.SetMathInputType(0);

            function AutoTest(isLaTeX, str, arrResultContent)
			{
				let CurPos = Root.CurPos;

				if (isConvertAfter === true)
					AscMath.SetAutoConvertation(false);

                AddText(str);

				if (isConvertAfter || isLaTeX)
				{
					MathContent.ConvertView(false, isLaTeX ? Asc.c_oAscMathInputType.LaTeX : Asc.c_oAscMathInputType.Unicode);
				}

                for (let i = CurPos; i < Root.Content.length; i++)
				{
                    let CurrentContent = Root.Content[i];
                    let CheckContent = arrResultContent[i];

					if (CheckContent === undefined)
						break;

                    assert.strictEqual(
                        CurrentContent.constructor.name,
                        CheckContent[0],
                        "Content[" + i + "] === " +
                        Root.Content[i].constructor.name
                    );

					let TextContent = CurrentContent.GetTextOfElement(isGetIntDifferentForm ? !logicDocument.MathInputType : logicDocument.MathInputType);
                    assert.strictEqual(TextContent.GetText(), CheckContent[1], "Text of Content[" + i + "]: '" + CheckContent[1] + "'");

                }

				if (isConvertAfter === true)
					AscMath.SetAutoConvertation(true);
            }

			Clear();
            AutoTest(isLaTeX, str, arrResult);
        })
	};
	function MultiLineTest(arrStr, arrCurPos, arrResult, arrCurPosMove)
	{
		QUnit.test("MultiLineTest \'" + arrStr.flat(2).join("") + "\'", function (assert) {

			Clear();
			for (let i = 0; i < arrStr.length; i++)
			{
				let str = arrStr[i];
				let intCurPos = arrCurPos[i];
				let arrCurResult = arrResult[i];
				let CurPosMove = arrCurPosMove[i];

				function AutoTest(str, intCurPos, arrResultContent, CurPosMove)
				{
					AddText(str);

					for (let i = 0; i < Root.Content.length; i++)
					{
						let CurrentContent = Root.Content[i];
						let ResultContent = arrResultContent[i];

						if (ResultContent === undefined) {
							ResultContent = [];
							ResultContent[0] = " " + Root.Content[i].constructor.name;
							ResultContent[1] = CurrentContent.GetTextOfElement();
						}

						assert.strictEqual(CurrentContent.constructor.name, ResultContent[0], "For: \'" + str + "\' block - " + "Content[" + i + "] === " + Root.Content[i].constructor.name);

						let TextContent = CurrentContent.GetTextOfElement();
						assert.strictEqual(TextContent, ResultContent[1], "For: \'" + str + "\' block - " + "Text of Content[" + i + "]: '" + ResultContent[1] + "'");

						if (CurrentContent.constructor.name === "ParaRun" && i === intCurPos)
							assert.strictEqual(CurrentContent.IsCursorAtEnd(), true, "For: \'" + str + "\' block - " + "Cursor at the end of ParaRun");
					}

					if (CurPosMove)
						Root.CurPos += CurPosMove;

					assert.strictEqual(Root.CurPos, intCurPos, "For: \'" + str + "\' block - " + "Check cursor position: " + intCurPos);
				}
				AutoTest(str, intCurPos, arrCurResult, CurPosMove);
			}
		})
	};


	QUnit.module( "Unicode", function ()
	{
		QUnit.module( "Auto-convert rules", function ()
		{
			Test("4^2^2^2 ", 2, [["ParaRun", "4^2^"], ["CDegree", "2^2"], ["ParaRun", ""]], false, "Check degree autocorrection rule");
			Test("4^2^2^2+", 2, [["ParaRun", ""], ["CDegree", "4^(2^(2^2))"], ["ParaRun", "+"]], false, "Check degree autocorrection rule");

			Test("4┴2┴2┴2 ", 2, [["ParaRun", "4┴2┴"], ["CLimit", "2┴2"], ["ParaRun", ""]], false, "Check Unicode AboveBelow");
			Test("4┴2┴2┴2+", 2, [["ParaRun", ""], ["CLimit", "4┴(2┴(2┴2))"], ["ParaRun", "+"]], false, "Check Unicode AboveBelow");
		})

		QUnit.module( "AboveBelow", function ()
		{
			QUnit.module( "auto-convert");
			Test("4┴2 +2", 2, [["ParaRun", ""], ["CLimit", "4┴2"], ["ParaRun", "+2"]], false, "Check Unicode AboveBelow");
			Test("base┴ex *xz", 2, [["ParaRun", ""], ["CLimit", "base┴ex"], ["ParaRun", "*xz"]], false, "Check Unicode AboveBelow");
			Test("2┴ex -p", 2, [["ParaRun", ""], ["CLimit", "2┴ex"], ["ParaRun", "-p"]], false, "Check Unicode AboveBelow");
			Test("base┬2 *x", 2, [["ParaRun", ""], ["CLimit", "base┬2"], ["ParaRun", "*x"]], false, "Check Unicode AboveBelow");
			Test("4┬2 +x/y ", 4, [["ParaRun", ""], ["CLimit", "4┬2"], ["ParaRun", "+"], ["CFraction", "x/y"]], false, "Check Unicode AboveBelow");
			Test("base┬x *y^2 ", 4, [["ParaRun", ""], ["CLimit", "base┬x"], ["ParaRun", "*"], ["CDegree", "y^2"]], false, "Check Unicode AboveBelow");
			Test("2┬ex -x_i ", 4, [["ParaRun", ""], ["CLimit", "2┬ex"], ["ParaRun", "-"], ["CDegree", "x_i"]], false, "Check Unicode AboveBelow");
			Test("2┬(ex+2) +(2+1) ", 4, [["ParaRun", ""], ["CLimit", "2┬(ex+2)"], ["ParaRun", "+"], ["CDelimiter", "(2+1)"]], false, "Check Unicode AboveBelow");
			Test("2┬(ex+2+x/2)^2 -1", 2, [["ParaRun", ""], ["CLimit", "2┬((ex+2+x/2)^2)"], ["ParaRun", "-1"]], false, "Check Unicode AboveBelow");
			Test("(2+x)┬ex ", 2, [["ParaRun", ""], ["CLimit", "(2+x)┬ex"]], false, "Check Unicode AboveBelow");
			Test("(2+y)┬(ex+2+x/2) ", 2, [["ParaRun", ""], ["CLimit", "(2+y)┬(ex+2+x/2)"]], false, "Check Unicode AboveBelow");
			Test("(2+y^2)┬(ex_3+2+x/2) ", 2, [["ParaRun", ""], ["CLimit", "(2+y^2)┬(ex_3+2+x/2)"]], false, "Check Unicode AboveBelow");

			QUnit.module("convert");
			Test("4┴2+2", 2, [["ParaRun", ""], ["CLimit", "4┴2"], ["ParaRun", "+2"]], false, "Check Unicode AboveBelow",true);
			Test("base┴ex*xz", 2, [["ParaRun", ""], ["CLimit", "base┴ex"], ["ParaRun", "*xz"]], false, "Check Unicode AboveBelow", true);
			Test("2┴ex-p", 2, [["ParaRun", ""], ["CLimit", "2┴ex"], ["ParaRun", "-p"]], false, "Check Unicode AboveBelow", true);
			Test("base┬2*x", 2, [["ParaRun", ""], ["CLimit", "base┬2"], ["ParaRun", "*x"]], false, "Check Unicode AboveBelow", true);
			Test("4┬2+x/y", 3, [["ParaRun", ""], ["CLimit", "4┬2"], ["ParaRun", "+"], ["CFraction", "x/y"]], false, "Check Unicode AboveBelow", true);
			Test("base┬x*y^2", 3, [["ParaRun", ""], ["CLimit", "base┬x"], ["ParaRun", "*"], ["CDegree", "y^2"]], false, "Check Unicode AboveBelow", true);
			Test("2┬ex-x_i", 3, [["ParaRun", ""], ["CLimit", "2┬ex"], ["ParaRun", "-"], ["CDegree", "x_i"]], false, "Check Unicode AboveBelow", true);
			Test("2┬(ex+2)+(2+1)", 3, [["ParaRun", ""], ["CLimit", "2┬(ex+2)"], ["ParaRun", "+"], ["CDelimiter", "(2+1)"]], false, "Check Unicode AboveBelow", true);
			Test("2┬(ex+2+x/2)^2-1", 2, [["ParaRun", ""], ["CLimit", "2┬((ex+2+x/2)^2)"], ["ParaRun", "-1"]], false, "Check Unicode AboveBelow", true);
			Test("(2+x)┬ex", 2, [["ParaRun", ""], ["CLimit", "(2+x)┬ex"]], false, "Check Unicode AboveBelow", true);
			Test("(2+y)┬(ex+2+x/2)", 2, [["ParaRun", ""], ["CLimit", "(2+y)┬(ex+2+x/2)"]], false, "Check Unicode AboveBelow", true);
			Test("(2+y^2)┬(ex_3+2+x/2)", 2, [["ParaRun", ""], ["CLimit", "(2+y^2)┬(ex_3+2+x/2)"]], false, "Check Unicode AboveBelow", true);

			Test("base┴2+2", 2, [["ParaRun", ""], ["CLimit", "base┴2"], ["ParaRun", "+2"]], false, "Check diacritics");
			Test("base┴2┴x+2", 2, [["ParaRun", ""], ["CLimit", "base┴(2┴x)"], ["ParaRun", "+2"]], false, "Check diacritics");
			Test("base┴2┴(x/y+6)+2", 2, [["ParaRun", ""], ["CLimit", "base┴(2┴(x/y+6))"], ["ParaRun", "+2"]], false, "Check diacritics");

			// for now, we don't trigger autocorrection by divide
			//Test("x^23┴2/y", 2, [["ParaRun", ""], ["CAccent", "e⃗"], ["ParaRun", ""]], false, "Check diacritics");
			//Test("(x^23)┴2/y", 2, [["ParaRun", ""], ["CAccent", "e⃗"], ["ParaRun", ""]], false, "Check diacritics");
		})

		QUnit.module( "Box and Rect", function ()
		{
			QUnit.module( "auto-convert");
			Test("□(1+2) ", 2, [["ParaRun", ""], ["CBox", "□(1+2)"]], false, "Check Unicode Box");
			Test("□1 ", 2, [["ParaRun", ""], ["CBox", "□1"]], false, "Check Unicode Box");
			Test("□1/2 ", 2, [["ParaRun", ""], ["CFraction", "□1/2"]], false, "Check Unicode Box");
			Test("▭(1+2) ", 2, [["ParaRun", ""], ["CBorderBox", "▭(1+2)"]], false, "Check Unicode Box");
			Test("▭1 ", 2, [["ParaRun", ""], ["CBorderBox", "▭1"]], false, "Check Unicode Box");
			Test("▭1/2 ", 2, [["ParaRun", ""], ["CFraction", "▭1/2"]], false, "Check Unicode Box");

			QUnit.module( "convert");
			Test("□(1+2)", 2, [["ParaRun", ""], ["CBox", "□(1+2)"]], false, "Check Unicode Box", true);
			Test("□1", 2, [["ParaRun", ""], ["CBox", "□1"]], false, "Check Unicode Box", true);
			Test("□1/2", 2, [["ParaRun", ""], ["CFraction", "□1/2"]], false, "Check Unicode Box", true);
			Test("▭(1+2)", 2, [["ParaRun", ""], ["CBorderBox", "▭(1+2)"]], false, "Check Unicode Box", true);
			Test("▭1", 2, [["ParaRun", ""], ["CBorderBox", "▭1"]], false, "Check Unicode Box", true);
			Test("▭1/2", 2, [["ParaRun", ""], ["CFraction", "▭1/2"]], false, "Check Unicode Box", true);

			Test("\\rect ", 0, [["ParaRun", "▭"]], false, "Check box literal");
			Test("\\rect 1/2 ", 2, [["ParaRun", ""], ["CFraction", "▭1/2"], ["ParaRun", ""]], false, "Check box");
			Test("\\rect (1/2) ", 2, [["ParaRun", ""], ["CBorderBox", "▭(1/2)"], ["ParaRun", ""]], false, "Check box");
			Test("\\rect (E=mc^2) ", 2, [["ParaRun", ""], ["CBorderBox", "▭(E=mc^2)"], ["ParaRun", ""]], false, "Check box");
		})

		QUnit.module( "Underbar", function ()
		{
			QUnit.module("auto-convert");
			Test("▁(1+2) ", 2, [["ParaRun", ""], ["CBar", "▁(1+2)"]], false, "Check Unicode underbar");
			Test("▁1 ", 2, [["ParaRun", ""], ["CBar", "▁1"]], false, "Check Unicode underbar");
			Test("▁1/2 ", 2, [["ParaRun", ""], ["CFraction", "▁1/2"]], false, "Check Unicode underbar");

			QUnit.module( "convert");
			Test("▁(1+2)", 2, [["ParaRun", ""], ["CBar", "▁(1+2)"]], false, "Check Unicode underbar", true);
			Test("▁1", 2, [["ParaRun", ""], ["CBar", "▁1"]], false, "Check Unicode underbar", true);
			Test("▁1/2", 2, [["ParaRun", ""], ["CFraction", "▁1/2"]], false, "Check Unicode underbar", true);
		})

		QUnit.module( "Brackets", function ()
		{
			QUnit.module( "auto-convert Brackets");
			Test(`(1+2) +2`, 2, [["ParaRun", ""], ["CDelimiter", "(1+2)"], ["ParaRun", "+2"]], false, "Check Unicode bracket", false);
			Test(`{1+2} -X`, 2, [["ParaRun", ""], ["CDelimiter", "{1+2}"], ["ParaRun", "-X"]], false, "Check Unicode bracket", false);
			Test(`[1+2] *i`, 2, [["ParaRun", ""], ["CDelimiter", "[1+2]"], ["ParaRun", "*i"]], false, "Check Unicode bracket", false);
			Test(`|1+2| -89/2 `, 3, [["ParaRun", ""], ["CDelimiter", "|1+2|"], ["ParaRun", "-"], ["CFraction", "89/2"]], false, "Check Unicode bracket", false);
			Test(`|1+2| -〖89/2〗 `, 3, [["ParaRun", ""], ["CDelimiter", "|1+2|"], ["ParaRun", "-"], ["CDelimiter", "〖89/2〗"]], false, "Check Unicode bracket", false);
			Test(`⌈1+2⌉ -〖89/2〗 `, 3, [["ParaRun", ""], ["CDelimiter", "⌈1+2⌉"], ["ParaRun", "-"], ["CDelimiter", "〖89/2〗"]], false, "Check Unicode bracket", false);
			Test(`⌊1+2⌋ -〖89/2〗 `, 3, [["ParaRun", ""], ["CDelimiter", "⌊1+2⌋"], ["ParaRun", "-"], ["CDelimiter", "〖89/2〗"]], false, "Check Unicode bracket", false);
			Test(`〖89/2〗/2 `, 2, [["ParaRun", ""], ["CFraction", "〖89/2〗/2"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`√〖89/2〗 `, 2, [["ParaRun", ""], ["CRadical", "√(89/2)"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`〖89/2〗_2 `, 2, [["ParaRun", ""], ["CDegree", "〖89/2〗_2"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`〖89/2〗^2 `, 2, [["ParaRun", ""], ["CDegree", "〖89/2〗^2"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2_〖89/2〗 `, 2, [["ParaRun", ""], ["CDegree", "2_〖89/2〗"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2^〖89/2〗 `, 2, [["ParaRun", ""], ["CDegree", "2^〖89/2〗"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2_〖89/2〗_2 `, 2, [["ParaRun", "2_"], ["CDegree", "〖89/2〗_2"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2^〖89/2〗^2 `, 2, [["ParaRun", "2^"], ["CDegree", "〖89/2〗^2"], ["ParaRun", ""]], false, "Check Unicode bracket",  false);
			Test(`2┴〖89/2〗 `, 2, [["ParaRun", ""], ["CLimit", "2┴〖89/2〗"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2┴〖89/2〗┴2 `, 2, [["ParaRun", "2┴"], ["CLimit", "〖89/2〗┴2"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2┬〖89/2〗 `, 2, [["ParaRun", ""], ["CLimit", "2┬〖89/2〗"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2┬〖89/2〗┬2 `, 2, [["ParaRun", "2┬"], ["CLimit", "〖89/2〗┬2"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`├]a+b┤[ `, 2, [["ParaRun", ""], ["CDelimiter", "├]a+b┤["], ["ParaRun", ""]], false, "Check Unicode bracket", false);

			Test("(", 0, [["ParaRun", "("]], false);
			Test("[", 0, [["ParaRun", "["]], false);
			Test("{", 0, [["ParaRun", "{"]], false);

			Test("( ", 0, [["ParaRun", "( "]], false);
			Test("[ ", 0, [["ParaRun", "[ "]], false);
			Test("{ ", 0, [["ParaRun", "{ "]], false);

			Test("(((", 0, [["ParaRun", "((("]], false);
			Test("[[[", 0, [["ParaRun", "[[["]], false);
			Test("{{{", 0, [["ParaRun", "{{{"]], false);

			Test("((( ", 0, [["ParaRun", "((( "]], false);
			Test("[[[ ", 0, [["ParaRun", "[[[ "]], false);
			Test("{{{ ", 0, [["ParaRun", "{{{ "]], false);

			Test("(((1", 0, [["ParaRun", "(((1"]], false);
			Test("[[[1", 0, [["ParaRun", "[[[1"]], false);
			Test("{{{1", 0, [["ParaRun", "{{{1"]], false);

			Test("(((1 ", 0, [["ParaRun", "(((1 "]], false);
			Test("[[[1 ", 0, [["ParaRun", "[[[1 "]], false);
			Test("{{{1 ", 0, [["ParaRun", "{{{1 "]], false);

			Test("1(((1", 0, [["ParaRun", "1(((1"]], false);
			Test("1[[[1", 0, [["ParaRun", "1[[[1"]], false);
			Test("1{{{1", 0, [["ParaRun", "1{{{1"]], false);

			Test("1(((1 ", 0, [["ParaRun", "1(((1 "]], false);
			Test("1[[[1 ", 0, [["ParaRun", "1[[[1 "]], false);
			Test("1{{{1 ", 0, [["ParaRun", "1{{{1 "]], false);

			Test("1(((1+", 0, [["ParaRun", "1(((1+"]], false);
			Test("1[[[1+", 0, [["ParaRun", "1[[[1+"]], false);
			Test("1{{{1+", 0, [["ParaRun", "1{{{1+"]], false);
			Test("1(((1+=", 0, [["ParaRun", "1(((1+="]], false);
			Test("1[[[1+=", 0, [["ParaRun", "1[[[1+="]], false);
			Test("1{{{1+=", 0, [["ParaRun", "1{{{1+="]], false);

			Test("1(((1+ ", 0, [["ParaRun", "1(((1+ "]], false);
			Test("1[[[1+ ", 0, [["ParaRun", "1[[[1+ "]], false);
			Test("1{{{1+ ", 0, [["ParaRun", "1{{{1+ "]], false);
			Test("1(((1+= ", 0, [["ParaRun", "1(((1+= "]], false);
			Test("1[[[1+= ", 0, [["ParaRun", "1[[[1+= "]], false);
			Test("1{{{1+= ", 0, [["ParaRun", "1{{{1+= "]], false);

			Test(")", 0, [["ParaRun", ")"]], false);
			Test("]", 0, [["ParaRun", "]"]], false);
			Test("}", 0, [["ParaRun", "}"]], false);

			Test(") ", 0, [["ParaRun", ") "]], false);
			Test("] ", 0, [["ParaRun", "] "]], false);
			Test("} ", 0, [["ParaRun", "} "]], false);

			Test(")))", 0, [["ParaRun", ")))"]], false);
			Test("]]]", 0, [["ParaRun", "]]]"]], false);
			Test("}}}", 0, [["ParaRun", "}}}"]], false);

			Test("))) ", 0, [["ParaRun", "))) "]], false);
			Test("]]] ", 0, [["ParaRun", "]]] "]], false);
			Test("}}} ", 0, [["ParaRun", "}}} "]], false);

			Test(")))1", 0, [["ParaRun", ")))1"]], false);
			Test("]]]1", 0, [["ParaRun", "]]]1"]], false);
			Test("}}}1", 0, [["ParaRun", "}}}1"]], false);

			Test(")))1 ", 0, [["ParaRun", ")))1 "]], false);
			Test("]]]1 ", 0, [["ParaRun", "]]]1 "]], false);
			Test("}}}1 ", 0, [["ParaRun", "}}}1 "]], false);

			Test("1)))1", 0, [["ParaRun", "1)))1"]], false);
			Test("1]]]1", 0, [["ParaRun", "1]]]1"]], false);
			Test("1}}}1", 0, [["ParaRun", "1}}}1"]], false);

			Test("1)))1 ", 0, [["ParaRun", "1)))1 "]], false);
			Test("1]]]1 ", 0, [["ParaRun", "1]]]1 "]], false);
			Test("1}}}1 ", 0, [["ParaRun", "1}}}1 "]], false);

			Test("1)))1+", 0, [["ParaRun", "1)))1+"]], false);
			Test("1]]]1+", 0, [["ParaRun", "1]]]1+"]], false);
			Test("1}}}1+", 0, [["ParaRun", "1}}}1+"]], false);
			Test("1)))1+=", 0, [["ParaRun", "1)))1+="]], false);
			Test("1]]]1+=", 0, [["ParaRun", "1]]]1+="]], false);
			Test("1}}}1+=", 0, [["ParaRun", "1}}}1+="]], false);

			Test("1)))1+ ", 0, [["ParaRun", "1)))1+ "]], false);
			Test("1]]]1+ ", 0, [["ParaRun", "1]]]1+ "]], false);
			Test("1}}}1+ ", 0, [["ParaRun", "1}}}1+ "]], false);
			Test("1)))1+= ", 0, [["ParaRun", "1)))1+= "]], false);
			Test("1]]]1+= ", 0, [["ParaRun", "1]]]1+= "]], false);
			Test("1}}}1+= ", 0, [["ParaRun", "1}}}1+= "]], false);

			Test("() ", 2, [["ParaRun", ""], ["CDelimiter", "()"], ["ParaRun", ""]], false);
			Test("{} ", 2, [["ParaRun", ""], ["CDelimiter", "{}"], ["ParaRun", ""]], false);
			Test("[] ", 2, [["ParaRun", ""], ["CDelimiter", "[]"], ["ParaRun", ""]], false);
			Test("|| ", 2, [["ParaRun", ""], ["CDelimiter", "||"], ["ParaRun", ""]], false);

			Test("()+", 2, [["ParaRun", ""], ["CDelimiter", "()"], ["ParaRun", "+"]], false);
			Test("{}+", 2, [["ParaRun", ""], ["CDelimiter", "{}"], ["ParaRun", "+"]], false);
			Test("[]+", 2, [["ParaRun", ""], ["CDelimiter", "[]"], ["ParaRun", "+"]], false);
			Test("||+", 2, [["ParaRun", ""], ["CDelimiter", "||"], ["ParaRun", "+"]], false);

			Test("(1+2)+", 2, [["ParaRun", ""], ["CDelimiter", "(1+2)"], ["ParaRun", "+"]], false);
			Test("{1+2}+", 2, [["ParaRun", ""], ["CDelimiter", "{1+2}"], ["ParaRun", "+"]], false);
			Test("[1+2]+", 2, [["ParaRun", ""], ["CDelimiter", "[1+2]"], ["ParaRun", "+"]], false);
			Test("|1+2|+", 2, [["ParaRun", ""], ["CDelimiter", "|1+2|"], ["ParaRun", "+"]], false);

			Test("(1/2 ", 0, [["ParaRun", "("], ["CFraction", "1/2"]], false);
			Test("{1/2 ", 0, [["ParaRun", "{"], ["CFraction", "1/2"]], false);
			Test("[1/2 ", 0, [["ParaRun", "["], ["CFraction", "1/2"]], false);
			Test("|1/2 ", 0, [["ParaRun", "|"], ["CFraction", "1/2"]], false);

			Test("(1/2)", 2, [["ParaRun", "("], ["CFraction", "1/2"], ["ParaRun", ")"]], false);

			Test("(1\\mid 2\\mid 3) ", 2, [["ParaRun", ""], ["CDelimiter", "(1∣2∣3)"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("[1\\mid 2\\mid 3) ", 2, [["ParaRun", ""], ["CDelimiter", "[1∣2∣3)"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("|1\\mid 2\\mid 3) ", 2, [["ParaRun", ""], ["CDelimiter", "|1∣2∣3)"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("{1\\mid 2\\mid 3) ", 2, [["ParaRun", ""], ["CDelimiter", "{1∣2∣3)"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("(1\\mid 2\\mid 3] ", 2, [["ParaRun", ""], ["CDelimiter", "(1∣2∣3]"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("(1\\mid 2\\mid 3} ", 2, [["ParaRun", ""], ["CDelimiter", "(1∣2∣3}"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("(1\\mid 2\\mid 3| ", 2, [["ParaRun", ""], ["CDelimiter", "(1∣2∣3|"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("|1\\mid 2\\mid 3| ", 2, [["ParaRun", ""], ["CDelimiter", "|1∣2∣3|"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("{1\\mid 2\\mid 3} ", 2, [["ParaRun", ""], ["CDelimiter", "{1∣2∣3}"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("[1\\mid 2\\mid 3] ", 2, [["ParaRun", ""], ["CDelimiter", "[1∣2∣3]"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");

			Test("(1+ ", 0, [["ParaRun", "(1+ "]], false, "Check brackets");
			Test("(1+2) ", 2, [["ParaRun", ""], ["CDelimiter", "(1+2)"], ["ParaRun", ""]], false, "Check brackets");
			Test("[1+2] ", 2, [["ParaRun", ""], ["CDelimiter", "[1+2]"], ["ParaRun", ""]], false, "Check brackets");
			Test("{1+2} ", 2, [["ParaRun", ""], ["CDelimiter", "{1+2}"], ["ParaRun", ""]], false, "Check brackets");

			Test(")123 ", 0, [["ParaRun", ")123 "]], false, "Check brackets");
			Test(")12) ", 0, [["ParaRun", ")12) "]], false, "Check brackets");
			Test(")12] ", 0, [["ParaRun", ")12] "]], false, "Check brackets");
			Test(")12} ", 0, [["ParaRun", ")12} "]], false, "Check brackets");

			Test("(1+2] ", 2, [["ParaRun", ""], ["CDelimiter", "(1+2]"], ["ParaRun", ""]], false, "Check brackets");
			Test("|1+2] ", 2, [["ParaRun", ""], ["CDelimiter", "|1+2]"], ["ParaRun", ""]], false, "Check brackets");
			Test("{1+2] ", 2, [["ParaRun", ""], ["CDelimiter", "{1+2]"], ["ParaRun", ""]], false, "Check brackets");

			QUnit.module( " convert Brackets");
			Test(`(1+2)+2`, 2, [["ParaRun", ""], ["CDelimiter", "(1+2)"], ["ParaRun", "+2"]], false, "Check Unicode bracket", true);
			Test(`{1+2}-X`, 2, [["ParaRun", ""], ["CDelimiter", "{1+2}"], ["ParaRun", "-X"]], false, "Check Unicode bracket", true);
			Test(`[1+2]*i`, 2, [["ParaRun", ""], ["CDelimiter", "[1+2]"], ["ParaRun", "*i"]], false, "Check Unicode bracket", true);
			Test(`|1+2|-89/2`, 3, [["ParaRun", ""], ["CDelimiter", "|1+2|"], ["ParaRun", "-"], ["CFraction", "89/2"]], false, "Check Unicode bracket", true);
			Test(`|1+2|-〖89/2〗`, 3, [["ParaRun", ""], ["CDelimiter", "|1+2|"], ["ParaRun", "-"], ["CDelimiter", "〖89/2〗"]], false, "Check Unicode bracket", true);
			Test(`⌈1+2⌉-〖89/2〗`, 3, [["ParaRun", ""], ["CDelimiter", "⌈1+2⌉"], ["ParaRun", "-"], ["CDelimiter", "〖89/2〗"]], false, "Check Unicode bracket", true);
			Test(`⌊1+2⌋-〖89/2〗`, 3, [["ParaRun", ""], ["CDelimiter", "⌊1+2⌋"], ["ParaRun", "-"], ["CDelimiter", "〖89/2〗"]], false, "Check Unicode bracket", true);
			Test(`〖89/2〗/2`, 2, [["ParaRun", ""], ["CFraction", "〖89/2〗/2"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`√〖89/2〗`, 2, [["ParaRun", ""], ["CRadical", "√(89/2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`〖89/2〗_2`, 2, [["ParaRun", ""], ["CDegree", "〖89/2〗_2"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`〖89/2〗^2`, 2, [["ParaRun", ""], ["CDegree", "〖89/2〗^2"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2_〖89/2〗`, 2, [["ParaRun", ""], ["CDegree", "2_〖89/2〗"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2^〖89/2〗`, 2, [["ParaRun", ""], ["CDegree", "2^〖89/2〗"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2_〖89/2〗_2`, 2, [["ParaRun", ""], ["CDegree", "2_(〖89/2〗_2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2^〖89/2〗^2`, 2, [["ParaRun", ""], ["CDegree", "2^(〖89/2〗^2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true );
			Test(`2┴〖89/2〗`, 2, [["ParaRun", ""], ["CLimit", "2┴〖89/2〗"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2┴〖89/2〗┴2`, 2, [["ParaRun", ""], ["CLimit", "2┴(〖89/2〗┴2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2┬〖89/2〗`, 2, [["ParaRun", ""], ["CLimit", "2┬〖89/2〗"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2┬〖89/2〗┬2`, 2, [["ParaRun", ""], ["CLimit", "2┬(〖89/2〗┬2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`├]a+b┤[`, 2, [["ParaRun", ""], ["CDelimiter", "├]a+b┤["], ["ParaRun", ""]], false, "Check Unicode bracket", true);
		})

		QUnit.module( "Complex", function ()
		{
			QUnit.module( " convert Complex");
			//Test(`(a + b)^n =∑_(k=0)^n▒(n¦k) a^k  b^(n-k)  `, 2, [["ParaRun", ""], ["CDegree", "(a + b)^n"], ["ParaRun", "="], ["CNary", "∑^n_(k=0)▒(n¦k)"],  ["ParaRun", ""], ["CDegree", "a^k"], ["CDegree", "b^(n-k)"]], false, "Check Complex content", true);
			Test(`∑_2^2▒(n/23)`, 2, [["ParaRun", ""], ["CNary", "∑_2^2▒(n/23)"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`(x+⋯+x)^(k "times")`, 2, [["ParaRun", ""], ["CDegree", "(x+⋯+x)^(k \"times\")"], ["ParaRun", ""]], false, "Check Complex content", true);
			Test(`𝐸 = 𝑚𝑐^2`, 2, [["ParaRun", "𝐸 ="], ["CDegree", " 𝑚𝑐^2"], ["ParaRun", ""]], false, "Check Complex content", true);
			Test(`∫_0^a▒xⅆx/(x^2+a^2)`, 2, [["ParaRun", ""], ["CNary", "∫_0^a▒〖xⅆx/(x^2+a^2)〗"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`lim┬(n→∞) a_n`, 2, [["ParaRun", ""], ["CLimit", "lim┬(n→∞)⁡a_n"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`ⅈ²=-1`, 2, [["ParaRun", ""], ["CDegree", "ⅈ²=-1"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`E = m⁢c²`, 2, [["ParaRun", "E ="], ["CDegree", "〖 m⁢c〗^2"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`a²⋅b²=c²`, 5, [["ParaRun", ""], ["CDegree", "a^2"], ["ParaRun", "⋅"],  ["CDegree", "b^2"], ["ParaRun", "="],  ["CDegree", "c^2"]], false, "Check Complex content", true);
			//Test(`f̂(ξ)=∫_-∞^∞▒f(x)ⅇ^-2πⅈxξ ⅆx`, 5, [["ParaRun", ""], ["CAccent", "f̂"], ["CDelimiter", "(ξ)"], ["ParaRun", "="], ["CNary", "∫▒〖ⅇ^ⅈxξ ⅆx〗"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`(𝑎 + 𝑏)┴→`, 2, [["ParaRun", ""], ["CLimit", "(𝑎 + 𝑏)┴→"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`𝑎┴→`, 2, [["ParaRun", ""], ["CLimit", "𝑎┴→"], ["ParaRun", ""]], false, "Check Complex content", true);
		})

		QUnit.module( "Fractions", function ()
		{
			QUnit.module( " convert fractions");
			Test(`1/2`, 2, [["ParaRun", ""], ["CFraction", "1/2"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`x+5/2`, 2, [["ParaRun", "x+"], ["CFraction", "5/2"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`x+5/x+2`, 2, [["ParaRun", "x+"], ["CFraction", "5/x"], ["ParaRun", "+2"]], false, "Check fraction content", true);
			Test(`1∕2`, 2, [["ParaRun", ""], ["CFraction", "1∕2"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`(x+5)/2`, 2, [["ParaRun", ""], ["CFraction", "(x+5)/2"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`x/(2+1)`, 2, [["ParaRun", ""], ["CFraction", "x/(2+1)"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`(x-5)/(2+1)`, 2, [["ParaRun", ""], ["CFraction", "(x-5)/(2+1)"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`1+3/2/3`, 2, [["ParaRun", "1+"], ["CFraction", "3/(2/3)"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`(𝛼_2^3)/(𝛽_2^3+𝛾_2^3)`, 2, [["ParaRun", ""], ["CFraction", "(𝛼_2^3)/(𝛽_2^3+𝛾_2^3)"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`(a/(b+c))/(d/e + f)`, 2, [["ParaRun", ""], ["CFraction", "(a/(b+c))/(d/e + f)"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`(a/(c/(z/x)))`, 2, [["ParaRun", ""], ["CDelimiter", "(a/(c/(z/x)))"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`1¦2`, 2, [["ParaRun", ""], ["CFraction", "1¦2"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`(1¦2)`, 2, [["ParaRun", ""], ["CDelimiter", "(1¦2)"], ["ParaRun", ""]], false, "Check fraction content", true);
		})

		QUnit.module( "Horizontal brackets", function ()
		{
			QUnit.module( " convert hbrackets");
			// Test(`⏞(x+⋯+x)`, 2, [["ParaRun", ""], ["CGroupCharacter", "⏞(x+⋯+x)"], ["ParaRun", ""]], false, "Check hbrack content", true);
			// Test(`⏞(x+⋯+x)^2`, 2, [["ParaRun", ""], ["CLimit", "⏞(x+⋯+x)┴2"], ["ParaRun", ""]], false, "Check hbrack content", true);
			// Test(`⏞(x+⋯+x)_2`, 2, [["ParaRun", ""], ["CLimit", "⏞(x+⋯+x)┬2"], ["ParaRun", ""]], false, "Check hbrack content", true);
			// Test(`⏞(x+⋯+x)_2^Y`, 2, [["ParaRun", ""], ["CLimit", "⏞(x+⋯+x)┬2^Y"], ["ParaRun", ""]], false, "Check hbrack content", true);
			// Test(`⏞(x+⋯+x)_2^2`, 2, [["ParaRun", ""], ["CLimit", "⏞(x+⋯+x)┬2^2"], ["ParaRun", ""]], false, "Check hbrack content", true);
		})

		QUnit.module( "Autocorrection", function ()
		{
			Test("\\above", 1, [["ParaRun", "\\above"]], false, "Check literal", true)
			Test("\\acute", 1, [["ParaRun", "\\acute"]], false, "Check literal", true)
			Test("\\aleph", 1, [["ParaRun", "\\aleph"]], false, "Check literal", true)
			Test("\\alpha", 0, [["ParaRun", "\\alpha"]], false, "Check literal", true)
			Test("\\amalg", 1, [["ParaRun", "\\amalg"]], false, "Check literal", true)
			Test("\\angle", 1, [["ParaRun", "\\angle"]], false, "Check literal", true)
			Test("\\aoint", 1, [["ParaRun", "\\aoint"]], false, "Check literal", true)
			Test("\\approx", 1, [["ParaRun", "\\approx"]], false, "Check literal", true)
			Test("\\asmash", 1, [["ParaRun", "\\asmash"]], false, "Check literal", true)
			Test("\\ast", 1, [["ParaRun", "\\ast"]], false, "Check literal", true)
			Test("\\asymp", 1, [["ParaRun", "\\asymp"]], false, "Check literal", true)
			Test("\\atop", 1, [["ParaRun", "\\atop"]], false, "Check literal", true)
			Test("\\Bar", 1, [["ParaRun", "\\Bar"]], false, "Check literal", true)
			Test("\\bar", 1, [["ParaRun", "\\bar"]], false, "Check literal", true)
			Test("\\because", 1, [["ParaRun", "\\because"]], false, "Check literal", true)
			Test("\\begin", 1, [["ParaRun", "\\begin"]], false, "Check literal", true)
			Test("\\below", 1, [["ParaRun", "\\below"]], false, "Check literal", true)
			Test("\\beta", 1, [["ParaRun", "\\beta"]], false, "Check literal", true)
			Test("\\beth", 1, [["ParaRun", "\\beth"]], false, "Check literal", true)
			Test("\\bot", 1, [["ParaRun", "\\bot"]], false, "Check literal", true)
			Test("\\bigcap", 1, [["ParaRun", "\\bigcap"]], false, "Check literal", true)
			Test("\\bigcup", 1, [["ParaRun", "\\bigcup"]], false, "Check literal", true)
			Test("\\bigodot", 1, [["ParaRun", "\\bigodot"]], false, "Check literal", true)
			Test("\\bigoplus", 1, [["ParaRun", "\\bigoplus"]], false, "Check literal", true)
			Test("\\bigotimes", 1, [["ParaRun", "\\bigotimes"]], false, "Check literal", true)
			Test("\\bigsqcup", 1, [["ParaRun", "\\bigsqcup"]], false, "Check literal", true)
			Test("\\biguplus", 1, [["ParaRun", "\\biguplus"]], false, "Check literal", true)
			Test("\\bigvee", 1, [["ParaRun", "\\bigvee"]], false, "Check literal", true)
			Test("\\bigwedge", 1, [["ParaRun", "\\bigwedge"]], false, "Check literal", true)
			Test("\\bowtie", 1, [["ParaRun", "\\bowtie"]], false, "Check literal", true)
			Test("\\box", 1, [["ParaRun", "\\box"]], false, "Check literal", true)
			Test("\\bra", 1, [["ParaRun", "\\bra"]], false, "Check literal", true)
			Test("\\breve", 1, [["ParaRun", "\\breve"]], false, "Check literal", true)
			Test("\\bullet", 1, [["ParaRun", "\\bullet"]], false, "Check literal", true)
			Test("\\boxdot", 1, [["ParaRun", "\\boxdot"]], false, "Check literal", true)
			Test("\\boxminus", 1, [["ParaRun", "\\boxminus"]], false, "Check literal", true)
			Test("\\boxplus", 1, [["ParaRun", "\\boxplus"]], false, "Check literal", true)
			Test("\\cap", 1, [["ParaRun", "\\cap"]], false, "Check literal", true)
			Test("\\cbrt", 1, [["ParaRun", "\\cbrt"]], false, "Check literal", true)
			Test("\\cdots", 1, [["ParaRun", "\\cdots"]], false, "Check literal", true)
			Test("\\cdot", 1, [["ParaRun", "\\cdot"]], false, "Check literal", true)
			Test("\\check", 1, [["ParaRun", "\\check"]], false, "Check literal", true)
			Test("\\chi", 1, [["ParaRun", "\\chi"]], false, "Check literal", true)
			Test("\\circ", 1, [["ParaRun", "\\circ"]], false, "Check literal", true)
			Test("\\close", 1, [["ParaRun", "\\close"]], false, "Check literal", true)
			Test("\\clubsuit", 1, [["ParaRun", "\\clubsuit"]], false, "Check literal", true)
			Test("\\coint", 1, [["ParaRun", "\\coint"]], false, "Check literal", true)
			Test("\\cong", 1, [["ParaRun", "\\cong"]], false, "Check literal", true)
			Test("\\contain", 1, [["ParaRun", "\\contain"]], false, "Check literal", true)
			Test("\\cup", 1, [["ParaRun", "\\cup"]], false, "Check literal", true)
			Test("\\daleth", 1, [["ParaRun", "\\daleth"]], false, "Check literal", true)
			Test("\\dashv", 1, [["ParaRun", "\\dashv"]], false, "Check literal", true)
			Test("\\dd", 1, [["ParaRun", "\\dd"]], false, "Check literal", true)
			Test("\\ddddot", 1, [["ParaRun", "\\ddddot"]], false, "Check literal", true)
			Test("\\dddot", 1, [["ParaRun", "\\dddot"]], false, "Check literal", true)
			Test("\\ddot", 1, [["ParaRun", "\\ddot"]], false, "Check literal", true)
			Test("\\ddots", 1, [["ParaRun", "\\ddots"]], false, "Check literal", true)
			Test("\\degree", 1, [["ParaRun", "\\degree"]], false, "Check literal", true)
			Test("\\Delta", 1, [["ParaRun", "\\Delta"]], false, "Check literal", true)
			Test("\\delta", 1, [["ParaRun", "\\delta"]], false, "Check literal", true)
			Test("\\diamond", 1, [["ParaRun", "\\diamond"]], false, "Check literal", true)
			Test("\\diamondsuit", 1, [["ParaRun", "\\diamondsuit"]], false, "Check literal", true)
			Test("\\div", 1, [["ParaRun", "\\div"]], false, "Check literal", true)
			Test("\\dot", 1, [["ParaRun", "\\dot"]], false, "Check literal", true)
			Test("\\doteq", 1, [["ParaRun", "\\doteq"]], false, "Check literal", true)
			Test("\\dots", 1, [["ParaRun", "\\dots"]], false, "Check literal", true)
			Test("\\downarrow", 1, [["ParaRun", "\\downarrow"]], false, "Check literal", true)
			Test("\\dsmash", 1, [["ParaRun", "\\dsmash"]], false, "Check literal", true)
			Test("\\degc", 1, [["ParaRun", "\\degc"]], false, "Check literal", true)
			Test("\\degf", 1, [["ParaRun", "\\degf"]], false, "Check literal", true)
			Test("\\ee", 1, [["ParaRun", "\\ee"]], false, "Check literal", true)
			Test("\\ell", 1, [["ParaRun", "\\ell"]], false, "Check literal", true)
			Test("\\emptyset", 1, [["ParaRun", "\\emptyset"]], false, "Check literal", true)
			Test("\\emsp", 1, [["ParaRun", "\\emsp"]], false, "Check literal", true)
			Test("\\end", 1, [["ParaRun", "\\end"]], false, "Check literal", true)
			Test("\\ensp", 1, [["ParaRun", "\\ensp"]], false, "Check literal", true)
			Test("\\epsilon", 1, [["ParaRun", "\\epsilon"]], false, "Check literal", true)
			Test("\\eqarray", 1, [["ParaRun", "\\eqarray"]], false, "Check literal", true)
			Test("\\eqno", 1, [["ParaRun", "\\eqno"]], false, "Check literal", true)
			Test("\\equiv", 1, [["ParaRun", "\\equiv"]], false, "Check literal", true)
			Test("\\eta", 1, [["ParaRun", "\\eta"]], false, "Check literal", true)
			Test("\\exists", 1, [["ParaRun", "\\exists"]], false, "Check literal", true)
			Test("\\forall", 1, [["ParaRun", "\\forall"]], false, "Check literal", true)
			Test("\\funcapply", 1, [["ParaRun", "\\funcapply"]], false, "Check literal", true)
			Test("\\frown", 1, [["ParaRun", "\\frown"]], false, "Check literal", true)
			Test("\\Gamma", 1, [["ParaRun", "\\Gamma"]], false, "Check literal", true)
			Test("\\gamma", 1, [["ParaRun", "\\gamma"]], false, "Check literal", true)
			Test("\\ge", 1, [["ParaRun", "\\ge"]], false, "Check literal", true)
			Test("\\geq", 1, [["ParaRun", "\\geq"]], false, "Check literal", true)
			Test("\\gets", 1, [["ParaRun", "\\gets"]], false, "Check literal", true)
			Test("\\gg", 1, [["ParaRun", "\\gg"]], false, "Check literal", true)
			Test("\\gimel", 1, [["ParaRun", "\\gimel"]], false, "Check literal", true)
			Test("\\grave", 1, [["ParaRun", "\\grave"]], false, "Check literal", true)
			Test("\\hairsp", 1, [["ParaRun", "\\hairsp"]], false, "Check literal", true)
			Test("\\hat", 1, [["ParaRun", "\\hat"]], false, "Check literal", true)
			Test("\\hbar", 1, [["ParaRun", "\\hbar"]], false, "Check literal", true)
			Test("\\heartsuit", 1, [["ParaRun", "\\heartsuit"]], false, "Check literal", true)
			Test("\\hookleftarrow", 1, [["ParaRun", "\\hookleftarrow"]], false, "Check literal", true)
			Test("\\hphantom", 1, [["ParaRun", "\\hphantom"]], false, "Check literal", true)
			Test("\\hsmash", 1, [["ParaRun", "\\hsmash"]], false, "Check literal", true)
			Test("\\hvec", 1, [["ParaRun", "\\hvec"]], false, "Check literal", true)
			Test("\\Im", 1, [["ParaRun", "\\Im"]], false, "Check literal", true)
			Test("\\iiiint", 1, [["ParaRun", "\\iiiint"]], false, "Check literal", true)
			Test("\\iiint", 1, [["ParaRun", "\\iiint"]], false, "Check literal", true)
			Test("\\iint", 1, [["ParaRun", "\\iint"]], false, "Check literal", true)
			Test("\\ii", 1, [["ParaRun", "\\ii"]], false, "Check literal", true)
			Test("\\int", 1, [["ParaRun", "\\int"]], false, "Check literal", true)
			Test("\\imath", 1, [["ParaRun", "\\imath"]], false, "Check literal", true)
			Test("\\inc", 1, [["ParaRun", "\\inc"]], false, "Check literal", true)
			Test("\\infty", 1, [["ParaRun", "\\infty"]], false, "Check literal", true)
			Test("\\in", 1, [["ParaRun", "\\in"]], false, "Check literal", true)
			Test("\\iota", 1, [["ParaRun", "\\iota"]], false, "Check literal", true)
			Test("\\jj", 1, [["ParaRun", "\\jj"]], false, "Check literal", true)
			Test("\\jmath", 1, [["ParaRun", "\\jmath"]], false, "Check literal", true)
			Test("\\kappa", 1, [["ParaRun", "\\kappa"]], false, "Check literal", true)
			Test("\\ket", 1, [["ParaRun", "\\ket"]], false, "Check literal", true)
			Test("\\Longleftrightarrow", 1, [["ParaRun", "\\Longleftrightarrow"]], false, "Check literal", true)
			Test("\\Longrightarrow", 1, [["ParaRun", "\\Longrightarrow"]], false, "Check literal", true)
			Test("\\Lambda", 1, [["ParaRun", "\\Lambda"]], false, "Check literal", true)
			Test("\\lambda", 1, [["ParaRun", "\\lambda"]], false, "Check literal", true)
			Test("\\langle", 1, [["ParaRun", "\\langle"]], false, "Check literal", true)
			Test("\\lbrack", 1, [["ParaRun", "\\lbrack"]], false, "Check literal", true)
			Test("\\ldiv", 1, [["ParaRun", "\\ldiv"]], false, "Check literal", true)
			Test("\\ldots", 1, [["ParaRun", "\\ldots"]], false, "Check literal", true)
			Test("\\le", 1, [["ParaRun", "\\le"]], false, "Check literal", true)
			Test("\\Leftarrow", 1, [["ParaRun", "\\Leftarrow"]], false, "Check literal", true)
			Test("\\leftarrow", 1, [["ParaRun", "\\leftarrow"]], false, "Check literal", true)
			Test("\\leftharpoondown", 1, [["ParaRun", "\\leftharpoondown"]], false, "Check literal", true)
			Test("\\leftharpoonup", 1, [["ParaRun", "\\leftharpoonup"]], false, "Check literal", true)
			Test("\\Leftrightarrow", 1, [["ParaRun", "\\Leftrightarrow"]], false, "Check literal", true)
			Test("\\leftrightarrow", 1, [["ParaRun", "\\leftrightarrow"]], false, "Check literal", true)
			Test("\\leq", 1, [["ParaRun", "\\leq"]], false, "Check literal", true)
			Test("\\lfloor", 1, [["ParaRun", "\\lfloor"]], false, "Check literal", true)
			Test("\\ll", 1, [["ParaRun", "\\ll"]], false, "Check literal", true)
			Test("\\Longleftarrow", 1, [["ParaRun", "\\Longleftarrow"]], false, "Check literal", true)
			Test("\\longleftarrow", 1, [["ParaRun", "\\longleftarrow"]], false, "Check literal", true)
			Test("\\longleftrightarrow", 1, [["ParaRun", "\\longleftrightarrow"]], false, "Check literal", true)
			Test("\\longrightarrow", 1, [["ParaRun", "\\longrightarrow"]], false, "Check literal", true)
			Test("\\lmoust", 1, [["ParaRun", "\\lmoust"]], false, "Check literal", true)
			Test("\\mapsto", 1, [["ParaRun", "\\mapsto"]], false, "Check literal", true)
			Test("\\matrix", 1, [["ParaRun", "\\matrix"]], false, "Check literal", true)
			Test("\\medsp", 1, [["ParaRun", "\\medsp"]], false, "Check literal", true)
			Test("\\mid", 1, [["ParaRun", "\\mid"]], false, "Check literal", true)
			Test("\\models", 1, [["ParaRun", "\\models"]], false, "Check literal", true)
			Test("\\mp", 1, [["ParaRun", "\\mp"]], false, "Check literal", true)
			Test("\\mu", 1, [["ParaRun", "\\mu"]], false, "Check literal", true)
			Test("\\nabla", 1, [["ParaRun", "\\nabla"]], false, "Check literal", true)
			Test("\\naryand", 1, [["ParaRun", "\\naryand"]], false, "Check literal", true)
			Test("\\nbsp", 1, [["ParaRun", "\\nbsp"]], false, "Check literal", true)
			Test("\\ndiv", 1, [["ParaRun", "\\ndiv"]], false, "Check literal", true)
			Test("\\ne", 1, [["ParaRun", "\\ne"]], false, "Check literal", true)
			Test("\\nearrow", 1, [["ParaRun", "\\nearrow"]], false, "Check literal", true)
			Test("\\neg", 1, [["ParaRun", "\\neg"]], false, "Check literal", true)
			Test("\\neq", 1, [["ParaRun", "\\neq"]], false, "Check literal", true)
			Test("\\ni", 1, [["ParaRun", "\\ni"]], false, "Check literal", true)
			Test("\\norm", 1, [["ParaRun", "\\norm"]], false, "Check literal", true)
			Test("\\nu", 1, [["ParaRun", "\\nu"]], false, "Check literal", true)
			Test("\\nwarrow", 1, [["ParaRun", "\\nwarrow"]], false, "Check literal", true)
			Test("\\Omega", 1, [["ParaRun", "\\Omega"]], false, "Check literal", true)
			Test("\\odot", 1, [["ParaRun", "\\odot"]], false, "Check literal", true)
			Test("\\of", 1, [["ParaRun", "\\of"]], false, "Check literal", true)
			Test("\\oiiint", 1, [["ParaRun", "\\oiiint"]], false, "Check literal", true)
			Test("\\oiint", 1, [["ParaRun", "\\oiint"]], false, "Check literal", true)
			Test("\\oint", 1, [["ParaRun", "\\oint"]], false, "Check literal", true)
			Test("\\omega", 1, [["ParaRun", "\\omega"]], false, "Check literal", true)
			Test("\\ominus", 1, [["ParaRun", "\\ominus"]], false, "Check literal", true)
			Test("\\open", 1, [["ParaRun", "\\open"]], false, "Check literal", true)
			Test("\\oplus", 1, [["ParaRun", "\\oplus"]], false, "Check literal", true)
			Test("\\oslash", 1, [["ParaRun", "\\oslash"]], false, "Check literal", true)
			Test("\\otimes", 1, [["ParaRun", "\\otimes"]], false, "Check literal", true)
			Test("\\over", 1, [["ParaRun", "\\over"]], false, "Check literal", true)
			Test("\\overbar", 1, [["ParaRun", "\\overbar"]], false, "Check literal", true)
			Test("\\overbrace", 1, [["ParaRun", "\\overbrace"]], false, "Check literal", true)
			Test("\\overbracket", 1, [["ParaRun", "\\overbracket"]], false, "Check literal", true)
			Test("\\overparen", 1, [["ParaRun", "\\overparen"]], false, "Check literal", true)
			Test("\\overshell", 1, [["ParaRun", "\\overshell"]], false, "Check literal", true)
			Test("\\over", 1, [["ParaRun", "\\over"]], false, "Check literal", true)
			Test("\\Pi", 1, [["ParaRun", "\\Pi"]], false, "Check literal", true)
			Test("\\Phi", 1, [["ParaRun", "\\Phi"]], false, "Check literal", true)
			Test("\\Psi", 1, [["ParaRun", "\\Psi"]], false, "Check literal", true)
			Test("\\parallel", 1, [["ParaRun", "\\parallel"]], false, "Check literal", true)
			Test("\\partial", 1, [["ParaRun", "\\partial"]], false, "Check literal", true)
			Test("\\perp", 1, [["ParaRun", "\\perp"]], false, "Check literal", true)
			Test("\\phantom", 1, [["ParaRun", "\\phantom"]], false, "Check literal", true)
			Test("\\phi", 1, [["ParaRun", "\\phi"]], false, "Check literal", true)
			Test("\\pi", 1, [["ParaRun", "\\pi"]], false, "Check literal", true)
			Test("\\pm", 1, [["ParaRun", "\\pm"]], false, "Check literal", true)
			Test("\\pppprime", 1, [["ParaRun", "\\pppprime"]], false, "Check literal", true)
			Test("\\ppprime", 1, [["ParaRun", "\\ppprime"]], false, "Check literal", true)
			Test("\\pprime", 1, [["ParaRun", "\\pprime"]], false, "Check literal", true)
			Test("\\prcue", 1, [["ParaRun", "\\prcue"]], false, "Check literal", true)
			Test("\\prec", 1, [["ParaRun", "\\prec"]], false, "Check literal", true)
			Test("\\preceq", 1, [["ParaRun", "\\preceq"]], false, "Check literal", true)
			Test("\\preccurlyeq", 1, [["ParaRun", "\\preccurlyeq"]], false, "Check literal", true)
			Test("\\prime", 1, [["ParaRun", "\\prime"]], false, "Check literal", true)
			Test("\\propto", 1, [["ParaRun", "\\propto"]], false, "Check literal", true)
			Test("\\psi", 1, [["ParaRun", "\\psi"]], false, "Check literal", true)
			Test("\\qdrt", 1, [["ParaRun", "\\qdrt"]], false, "Check literal", true)
			Test("\\Re", 1, [["ParaRun", "\\Re"]], false, "Check literal", true)
			Test("\\Rightarrow", 1, [["ParaRun", "\\Rightarrow"]], false, "Check literal", true)
			Test("\\rangle", 1, [["ParaRun", "\\rangle"]], false, "Check literal", true)
			Test("\\ratio", 1, [["ParaRun", "\\ratio"]], false, "Check literal", true)
			Test("\\rbrace", 1, [["ParaRun", "\\rbrace"]], false, "Check literal", true)
			Test("\\rbrack", 1, [["ParaRun", "\\rbrack"]], false, "Check literal", true)
			Test("\\rceil", 1, [["ParaRun", "\\rceil"]], false, "Check literal", true)
			Test("\\rddots", 1, [["ParaRun", "\\rddots"]], false, "Check literal", true)
			Test("\\rect", 1, [["ParaRun", "\\rect"]], false, "Check literal", true)
			Test("\\rfloor", 1, [["ParaRun", "\\rfloor"]], false, "Check literal", true)
			Test("\\rho", 1, [["ParaRun", "\\rho"]], false, "Check literal", true)
			Test("\\right", 1, [["ParaRun", "\\right"]], false, "Check literal", true)
			Test("\\rightarrow", 1, [["ParaRun", "\\rightarrow"]], false, "Check literal", true)
			Test("\\rightharpoondown", 1, [["ParaRun", "\\rightharpoondown"]], false, "Check literal", true)
			Test("\\rightharpoonup", 1, [["ParaRun", "\\rightharpoonup"]], false, "Check literal", true)
			Test("\\rmoust", 1, [["ParaRun", "\\rmoust"]], false, "Check literal", true)
			Test("\\rrect", 1, [["ParaRun", "\\rrect"]], false, "Check literal", true)
			Test("\\root", 1, [["ParaRun", "\\root"]], false, "Check literal", true)
			Test("\\Sigma", 1, [["ParaRun", "\\Sigma"]], false, "Check literal", true)
			Test("\\sdiv", 1, [["ParaRun", "\\sdiv"]], false, "Check literal", true)
			Test("\\searrow", 1, [["ParaRun", "\\searrow"]], false, "Check literal", true)
			Test("\\setminus", 1, [["ParaRun", "\\setminus"]], false, "Check literal", true)
			Test("\\sigma", 1, [["ParaRun", "\\sigma"]], false, "Check literal", true)
			Test("\\sim", 1, [["ParaRun", "\\sim"]], false, "Check literal", true)
			Test("\\simeq", 1, [["ParaRun", "\\simeq"]], false, "Check literal", true)
			Test("\\smash", 1, [["ParaRun", "\\smash"]], false, "Check literal", true)
			Test("\\smile", 1, [["ParaRun", "\\smile"]], false, "Check literal", true)
			Test("\\spadesuit", 1, [["ParaRun", "\\spadesuit"]], false, "Check literal", true)
			Test("\\sqcap", 1, [["ParaRun", "\\sqcap"]], false, "Check literal", true)
			Test("\\sqcup", 1, [["ParaRun", "\\sqcup"]], false, "Check literal", true)
			Test("\\sqrt", 1, [["ParaRun", "\\sqrt"]], false, "Check literal", true)
			Test("\\sqsubseteq", 1, [["ParaRun", "\\sqsubseteq"]], false, "Check literal", true)
			Test("\\sqsuperseteq", 1, [["ParaRun", "\\sqsuperseteq"]], false, "Check literal", true)
			Test("\\star", 1, [["ParaRun", "\\star"]], false, "Check literal", true)
			Test("\\subset", 1, [["ParaRun", "\\subset"]], false, "Check literal", true)
			Test("\\subseteq", 1, [["ParaRun", "\\subseteq"]], false, "Check literal", true)
			Test("\\succeq", 1, [["ParaRun", "\\succeq"]], false, "Check literal", true)
			Test("\\succ", 1, [["ParaRun", "\\succ"]], false, "Check literal", true)
			Test("\\sum", 1, [["ParaRun", "\\sum"]], false, "Check literal", true)
			Test("\\superset", 1, [["ParaRun", "\\superset"]], false, "Check literal", true)
			Test("\\superseteq", 1, [["ParaRun", "\\superseteq"]], false, "Check literal", true)
			Test("\\swarrow", 1, [["ParaRun", "\\swarrow"]], false, "Check literal", true)
		})

		QUnit.module( "Degree", function ()
		{
			QUnit.module( "degree");
			Test("2^2 + 2", 2, [["ParaRun", ""], ["CDegree", "2^2"], ["ParaRun", "+ 2"]], false, "Check scripts", true)
			Test("x^2+2", 2, [["ParaRun", ""], ["CDegree", "x^2"], ["ParaRun", "+2"]], false, "Check scripts", true)
			Test("x^(256+34)*y", 2, [["ParaRun", ""], ["CDegree", "x^(256+34)"], ["ParaRun", "*y"]], false, "Check scripts", true)
			Test("(x+34)^(256+34)-y/x", 3, [["ParaRun", ""], ["CDegree", "(x+34)^(256+34)"], ["ParaRun", "-"], ["CFraction", "y/x"]], false, "Check scripts", true)

			Test("2_1", 0, [["ParaRun", "2_1"]], false);
			Test("2_1 ", 2, [["ParaRun", ""], ["CDegree", "2_1"], ["ParaRun", ""]], false);
			Test("\\int", 0, [["ParaRun", "\\int"]], false);
			Test("\\int _x^y\\of 1/2 ", 2, [["ParaRun", "∫_x^y▒"], ["CFraction", "1/2"], ["ParaRun", ""]], false);
			Test("1/2 ", 2, [["ParaRun", ""], ["CFraction", "1/2"], ["ParaRun", ""]], false);
			Test("1/2 +", 2, [["ParaRun", ""], ["CFraction", "1/2"], ["ParaRun", "+"]], false);
			Test("1/2=", 2, [["ParaRun", ""], ["CFraction", "1/2"], ["ParaRun", "="]], false);
			Test("1/2+1/2=x/y ", 6, [["ParaRun", ""], ["CFraction", "1/2"], ["ParaRun", "+"], ["CFraction", "1/2"], ["ParaRun", "="], ["CFraction", "x/y"], ["ParaRun", ""]], false);

			Test("x_y ", 2, [["ParaRun", ""], ["CDegree", "x_y"], ["ParaRun", ""]], false, "Check degree");
			Test("_ ", 2, [["ParaRun", ""], ["CDegree", "_"]], false, "Check degree");
			Test("x_1 ", 2, [["ParaRun", ""], ["CDegree", "x_1"], ["ParaRun", ""]], false, "Check degree");
			Test("1_x ", 2, [["ParaRun", ""], ["CDegree", "1_x"], ["ParaRun", ""]], false, "Check degree");
			Test("x_(1+2) ", 2, [["ParaRun", ""], ["CDegree", "x_(1+2)"], ["ParaRun", ""]], false, "Check degree");
			Test("x_[1+2] ", 2, [["ParaRun", ""], ["CDegree", "x_[1+2]"], ["ParaRun", ""]], false, "Check degree");
			Test("x_[1+2} ", 2, [["ParaRun", ""], ["CDegree", "x_[1+2}"], ["ParaRun", ""]], false, "Check degree");
			Test("x_1/2", 2, [["ParaRun", ""], ["CDegree", "x_1"], ["ParaRun", "/2"]], false, "Check degree");
			// Test("x_1/2 ", 2, [["ParaRun", ""], ["CFraction", "x_1/2"], ["ParaRun", ""]], false, "Check degree");

			QUnit.module( "[U-AC] degree");
			Test("^ ", 2, [["ParaRun", ""], ["CDegree", "^"]], false, "Check index");
			Test("x^y ", 2, [["ParaRun", ""], ["CDegree", "x^y"], ["ParaRun", ""]], false, "Check index");
			Test("x^1 ", 2, [["ParaRun", ""], ["CDegree", "x^1"], ["ParaRun", ""]], false, "Check index");
			Test("1^x ", 2, [["ParaRun", ""], ["CDegree", "1^x"], ["ParaRun", ""]], false, "Check index");
			Test("x^(1+2) ", 2, [["ParaRun", ""], ["CDegree", "x^(1+2)"], ["ParaRun", ""]], false, "Check index");
			Test("x^[1+2] ", 2, [["ParaRun", ""], ["CDegree", "x^[1+2]"], ["ParaRun", ""]], false, "Check index");
			Test("x^[1+2} ", 2, [["ParaRun", ""], ["CDegree", "x^[1+2}"], ["ParaRun", ""]], false, "Check index");
			//Test("x^1/2", 2, [["ParaRun", ""], ["CFraction", "x^1"], ["ParaRun", "/2"]], false, "Check index");
			//Test("x^1/2 ", 2, [["ParaRun", ""], ["CFraction", "x^1/2"], ["ParaRun", ""]], false, "Check index");

			Test("x^y_1 ", 2, [["ParaRun", ""], ["CDegreeSubSup", "x_1^y"], ["ParaRun", ""]], false, "Check index degree");
			Test("x^1_i ", 2, [["ParaRun", ""], ["CDegreeSubSup", "x_i^1"], ["ParaRun", ""]], false, "Check index degree");
			Test("1^x_y ", 2, [["ParaRun", ""], ["CDegreeSubSup", "1_y^x"], ["ParaRun", ""]], false, "Check index degree");
			Test("x^[1+2]_[g_i] ", 2, [["ParaRun", ""], ["CDegreeSubSup", "x_[g_i]^[1+2]"], ["ParaRun", ""]], false, "Check index degree");
			Test("x^[1+2}_[6+1} ", 2, [["ParaRun", ""], ["CDegreeSubSup", "x_[6+1}^[1+2}"], ["ParaRun", ""]], false, "Check index degree");
			//Test("x^1/2_1/2 ", 2, [["ParaRun", ""], ["CFraction", "x^1/(2_1/2)"], ["ParaRun", ""]], false, "Check index degree");

			//Test("𝑊^3𝛽_𝛿1𝜌1𝜎2 ", 2, [["ParaRun", ""], ["CDegree", "𝑊^3𝛽_𝛿1𝜌1𝜎2"], ["ParaRun", ""]], false, "Check index degree with Unicode symbols");
			QUnit.module( "pre-script");
			// Test("(_1^f)f ", 2, [["ParaRun", ""], ["CDegreeSubSup", "(_1^f)f"], ["ParaRun", ""]], false, "Check prescript index degree");
			// Test("(_(1/2)^y)f ", 2, [["ParaRun", ""], ["CDegreeSubSup", "(_(1/2)^y)f"], ["ParaRun", ""]], false, "Check prescript index degree");
			// Test("(_(1/2)^[x_i])x/y  ", 2, [["ParaRun", ""], ["CDegreeSubSup", "(_(1/2)^[x_i])x/y"], ["ParaRun", ""]], false, "Check prescript index degree");
		})

		QUnit.module( "Radicals", function ()
		{
			QUnit.module( " convert radicals");
			Test("√5", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("√a", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("√a/2", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("√(2&a-4)", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("∛5", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("∛a", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("∛a/2", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("∛(a-4)", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("∜5", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("∜a", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("∜a/2", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("∜(a-4)", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("√(10&a/4)", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("√(10^2&a/4+2)", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("√5^2", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("√5_2", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("√5^2_x", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("√5_2^x", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("(_5^2)√5", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("√5┴exp1", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("√5┬exp1", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("(√5┬exp1]", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("□√5", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("▭√5", 2, [["ParaRun", ""]], false, "Check special", true)
			Test("▁√5", 2, [["ParaRun", ""]], false, "Check special", true)
			// Test(` ̄√5`.trim(), 2, [["ParaRun", ""]], false, "Check special", true)
			Test("∑_√5^√5", 2, [["ParaRun", ""]], false, "Check special", true)

			Test("\\sqrt ", 0, [["ParaRun", "√"]], false, "Check");
			Test("\\sqrt (2&1+2) ", 2, [["ParaRun", ""], ["CRadical", "√(2&1+2)"], ["ParaRun", ""]], false, "Check radical");
			Test("\\sqrt (1+2) ", 2, [["ParaRun", ""], ["CRadical", "√(1+2)"], ["ParaRun", ""]], false, "Check radical");
			Test("√1 ", 2, [["ParaRun", ""], ["CRadical", "√1"], ["ParaRun", ""]], false, "Check radical");

			Test("\\cbrt ", 0, [["ParaRun", "∛"]], false, "Check");
			Test("\\cbrt (1+2) ", 2, [["ParaRun", ""], ["CRadical", "∛(1+2)"], ["ParaRun", ""]], false, "Check radical");
			Test("\\cbrt 1/2 ", 2, [["ParaRun", ""], ["CFraction", "∛1/2"], ["ParaRun", ""]], false, "Check radical");
			Test("∛1 ", 2, [["ParaRun", ""], ["CRadical", "∛1"], ["ParaRun", ""]], false, "Check radical");
			Test("∛(1) ", 2, [["ParaRun", ""], ["CRadical", "∛1"], ["ParaRun", ""]], false, "Check radical");

			Test("\\qdrt ", 0, [["ParaRun", "∜"]], false, "Check");
			Test("\\qdrt (1+2) ", 2, [["ParaRun", ""], ["CRadical", "∜(1+2)"], ["ParaRun", ""]], false, "Check radical");
			Test("\\qdrt 1/2 ", 2, [["ParaRun", ""], ["CFraction", "∜1/2"], ["ParaRun", ""]], false, "Check radical");
			Test("∜1 ", 2, [["ParaRun", ""], ["CRadical", "∜1"], ["ParaRun", ""]], false, "Check radical");
			Test("∜(1) ", 2, [["ParaRun", ""], ["CRadical", "∜1"], ["ParaRun", ""]], false, "Check radical");
		})

		QUnit.module( "Other", function ()
		{
			QUnit.module( " convert operators");
			Test("×", 0, [["ParaRun", "×"]], false, "Check literal", true)
			Test("⋅", 0, [["ParaRun", "⋅"]], false, "Check literal")
			Test("∈", 0, [["ParaRun", "∈"]], false, "Check literal")
			Test("∋", 0, [["ParaRun", "∋"]], false, "Check literal")
			Test("∼", 0, [["ParaRun", "∼"]], false, "Check literal")
			Test("≃", 0, [["ParaRun", "≃"]], false, "Check literal")
			Test("≅", 0, [["ParaRun", "≅"]], false, "Check literal")
			Test("≈", 0, [["ParaRun", "≈"]], false, "Check literal")
			Test("≍", 0, [["ParaRun", "≍"]], false, "Check literal")
			Test("≡", 0, [["ParaRun", "≡"]], false, "Check literal")
			Test("≤", 0, [["ParaRun", "≤"]], false, "Check literal")
			Test("≥", 0, [["ParaRun", "≥"]], false, "Check literal")
			Test("≶", 0, [["ParaRun", "≶"]], false, "Check literal")
			Test("≷", 0, [["ParaRun", "≷"]], false, "Check literal")
			Test("≽", 0, [["ParaRun", "≽"]], false, "Check literal")
			Test("≺", 0, [["ParaRun", "≺"]], false, "Check literal")
			Test("≻", 0, [["ParaRun", "≻"]], false, "Check literal")
			Test("≼", 0, [["ParaRun", "≼"]], false, "Check literal")
			Test("⊂", 0, [["ParaRun", "⊂"]], false, "Check literal")
			Test("⊃", 0, [["ParaRun", "⊃"]], false, "Check literal")
			Test("⊆", 0, [["ParaRun", "⊆"]], false, "Check literal")
			Test("⊇", 0, [["ParaRun", "⊇"]], false, "Check literal")
			Test("⊑", 0, [["ParaRun", "⊑"]], false, "Check literal")
			Test("⊒", 0, [["ParaRun", "⊒"]], false, "Check literal")
			Test("+", 0, [["ParaRun", "+"]], false, "Check literal")
			Test("-", 0, [["ParaRun", "-"]], false, "Check literal")
			Test("=", 0, [["ParaRun", "="]], false, "Check literal")
			Test("*", 0, [["ParaRun", "*"]], false, "Check literal")
			Test("∃", 0, [["ParaRun", "∃"]], false, "Check literal")
			Test("∀", 0, [["ParaRun", "∀"]], false, "Check literal")
			Test("¬", 0, [["ParaRun", "¬"]], false, "Check literal")
			Test("∧", 0, [["ParaRun", "∧"]], false, "Check literal")
			Test("∨", 0, [["ParaRun", "∨"]], false, "Check literal")
			Test("⇒", 0, [["ParaRun", "⇒"]], false, "Check literal")
			Test("⇔", 0, [["ParaRun", "⇔"]], false, "Check literal")
			Test("⊕", 0, [["ParaRun", "⊕"]], false, "Check literal")
			Test("⊤", 0, [["ParaRun", "⊤"]], false, "Check literal")
			Test("⊥", 0, [["ParaRun", "⊥"]], false, "Check literal")
			Test("⊢", 0, [["ParaRun", "⊢"]], false, "Check literal")
			Test("⨯", 0, [["ParaRun", "⨯"]], false, "Check literal")
			Test("⟕", 0, [["ParaRun", "⟕"]], false, "Check literal")
			Test("⟖", 0, [["ParaRun", "⟖"]], false, "Check literal")
			Test("⟗", 0, [["ParaRun", "⟗"]], false, "Check literal")
			Test("⋉", 0, [["ParaRun", "⋉"]], false, "Check literal")
			Test("⋊", 0, [["ParaRun", "⋊"]], false, "Check literal")
			Test("▷", 0, [["ParaRun", "▷"]], false, "Check literal")
			Test("÷", 0, [["ParaRun", "÷"]], false, "Check literal")
			Test("⁡", 0, [["ParaRun", "⁡"]], false, "Check literal")
			Test("⁢", 0, [["ParaRun", "⁢"]], false, "Check literal")
			Test("⁣", 0, [["ParaRun", "⁣"]], false, "Check literal")
			Test("⁤", 0, [["ParaRun", "⁤"]], false, "Check literal")
			Test("​", 0, [["ParaRun", "​"]], false, "Check literal")
			Test(" ", 0, [["ParaRun", " "]], false, "Check literal")
			Test("  ", 0, [["ParaRun", "  "]], false, "Check literal")
			Test(" ", 0, [["ParaRun", " "]], false, "Check literal")
			Test(" ", 0, [["ParaRun", " "]], false, "Check literal")
			Test(" ", 0, [["ParaRun", " "]], false, "Check literal")
			Test("  ", 0, [["ParaRun", "  "]], false, "Check literal")
			Test(" ", 0, [["ParaRun", " "]], false, "Check literal")
			Test(" ", 0, [["ParaRun", " "]], false, "Check literal")
			Test(" ", 0, [["ParaRun", " "]], false, "Check literal")
			Test(" ", 0, [["ParaRun", " "]], false, "Check literal")
			Test(`a`, 0, [["ParaRun", "a"]], false, "Check literal")
			Test(`abcdef`, 0, [["ParaRun", "abcdef"]], false, "Check literal")
			Test(`1`, 0, [["ParaRun", "1"]], false, "Check literal")
			Test(`1234`, 0, [["ParaRun", "1234"]], false, "Check literal")
			Test(`1+2`, 0, [["ParaRun", "1+2"]], false, "Check literal")
			Test(`1+2+3`, 0, [["ParaRun", "1+2+3"]], false, "Check literal")
			Test(`ΑαΒβΓγΔδΕεΖζΗηΘθΙιΚκΛλΜμΝνΞξΟοΠπΡρΣσΤτΥυΦφΧχΨψΩω`, 0, [["ParaRun", "ΑαΒβΓγΔδΕεΖζΗηΘθΙιΚκΛλΜμΝνΞξΟοΠπΡρΣσΤτΥυΦφΧχΨψΩω"]], false, "Check literal")
			Test("abc123def", 0, [["ParaRun", "abc123def"]], false, "Check literal")
			Test("abc+123+def", 0, [["ParaRun", "abc+123+def"]], false, "Check literal")
			Test("𝐀𝐁𝐂𝐨𝐹", 0, [["ParaRun", "𝐀𝐁𝐂𝐨𝐹"]], false, "Check literal")
			Test("   𝐀𝐁𝐂𝐨𝐹   ", 0, [["ParaRun", "   𝐀𝐁𝐂𝐨𝐹   "]], false, "Check literal")
			Test(" 	𝐀𝐁𝐂𝐨𝐹  	 ", 0, [["ParaRun", " 	𝐀𝐁𝐂𝐨𝐹  	 "]], false, "Check literal")
			Test(`1+fbnd+(3+𝐀𝐁𝐂𝐨𝐹)+c+5`, 2, [["ParaRun", "1+fbnd+"], ["CDelimiter", "(3+𝐀𝐁𝐂𝐨𝐹)"], ["ParaRun", "+c+5"]], false, "Check literal")
			//Test(`1/3.1416`, 1, [["ParaRun", ""]], false, "Check literal")
			// Test("1\\above2", 2, [["ParaRun", ""], ["CLimit", "1┴2"]], false, "Check literal", true)
			// Test("1\\acute2", 2, [["ParaRun", ""], ["ParaRun", ""]], false, "Check literal", true)


			//QUnit.module( " convert operators");
			// Test("2⁰¹²³⁴⁵⁶⁷⁸⁹", 1, [], false, "Check special")
			// Test("2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹", 1, [], false, "Check special")
			// Test("2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45", 1, [], false, "Check special")
			// Test("x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45", 1, [], false, "Check special")
			// Test("2₂₃₄₊₍₆₇₋₀₌₆₇₎56", 1, [], false, "Check special")
			// Test("z₂₃₄₊₍₆₇₋₀₌₆₇₎56", 1, [], false, "Check special")
			// Test("2⁰¹²³⁴⁵⁶⁷⁸⁹₂₃₄₊₍₆₇₋₀₌₆₇₎", 1, [], false, "Check special")
			// Test("2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎", 1, [], false, "Check special")
			// Test("2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45", 1, [], false, "Check special")
			// Test("x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45", 1, [], false, "Check special")
			// Test("2₂₃₄₊₍₆₇₋₀₌₆₇₎⁰¹²³⁴⁵⁶⁷⁸⁹", 1, [], false, "Check special")
			// Test("2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹", 1, [], false, "Check special")
			// Test("2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45", 1, [], false, "Check special")
			// Test("x₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45", 1, [], false, "Check special")
		})

		QUnit.module( "Nary", function ()
		{
			Test("\\int ", 0, [["ParaRun", "∫"]], false, "Check large operators");
			Test("\\int  ", 2, [["ParaRun", ""], ["CNary", "∫"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int _x ", 2, [["ParaRun", ""], ["CNary", "∫_x"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int ^x ", 2, [["ParaRun", ""], ["CNary", "∫^x"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int ^(x+1) ", 2, [["ParaRun", ""], ["CNary", "∫^(x+1)"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int ^(x+1) ", 2, [["ParaRun", ""], ["CNary", "∫^(x+1)"], ["ParaRun", ""]],false, "Check large operators");
			Test("\\int ^(x+1)_(1_i) ", 2, [["ParaRun", ""], ["CNary", "∫_(1_i)^(x+1)"], ["ParaRun", ""]], false, "Check large operators");

			Test("\\int \\of x ", 2, [["ParaRun", ""], ["CNary", "∫▒x"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int _x\\of 1/2  ", 2, [["ParaRun", ""], ["CNary", "∫_x▒〖1/2〗"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int ^x\\of 1/2  ", 2, [["ParaRun", ""], ["CNary", "∫^x▒〖1/2〗"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int _(x+1)\\of 1/2  ", 2, [["ParaRun", ""], ["CNary", "∫_(x+1)▒〖1/2〗"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\prod ^(x+1)\\of 1/2  ", 2, [["ParaRun", ""], ["CNary", "∏^(x+1)▒〖1/2〗"], ["ParaRun", ""]],false, "Check large operators");
			Test("∫^(x+1)_(1_i)\\of 1/2  ", 2, [["ParaRun", ""], ["CNary", "∫_(1_i)^(x+1)▒〖1/2〗"], ["ParaRun", ""]], false, "Check large operators");
		})

		QUnit.module( "Functions", function ()
		{

			Test("sin ", 2, [["ParaRun", ""], ["CMathFunc", "sin⁡"]], false, "Check functions");
			Test("cos ", 2, [["ParaRun", ""], ["CMathFunc", "cos⁡"]], false, "Check functions");
			Test("tan ", 2, [["ParaRun", ""], ["CMathFunc", "tan⁡"]], false, "Check functions");
			Test("csc ", 2, [["ParaRun", ""], ["CMathFunc", "csc⁡"]], false, "Check functions");
			Test("sec ", 2, [["ParaRun", ""], ["CMathFunc", "sec⁡"]], false, "Check functions");
			Test("cot ", 2, [["ParaRun", ""], ["CMathFunc", "cot⁡"]], false, "Check functions");

			Test("sin", 2, [["ParaRun", "sin"]], false, "Check functions");
			Test("cos", 2, [["ParaRun", "cos"]], false, "Check functions");
			Test("tan", 2, [["ParaRun", "tan"]], false, "Check functions");
			Test("csc", 2, [["ParaRun", "csc"]], false, "Check functions");
			Test("sec", 2, [["ParaRun", "sec"]], false, "Check functions");
			Test("cot", 2, [["ParaRun", "cot"]], false, "Check functions");

			Test("sin a", 2, [["ParaRun", ""], ["CMathFunc", "sin⁡a"], ["ParaRun", ""]], false, "Check functions");
			Test("cos a", 2, [["ParaRun", ""], ["CMathFunc", "cos⁡a"], ["ParaRun", ""]], false, "Check functions");
			Test("tan a", 2, [["ParaRun", ""], ["CMathFunc", "tan⁡a"], ["ParaRun", ""]], false, "Check functions");
			Test("csc a", 2, [["ParaRun", ""], ["CMathFunc", "csc⁡a"], ["ParaRun", ""]], false, "Check functions");
			Test("sec a", 2, [["ParaRun", ""], ["CMathFunc", "sec⁡a"], ["ParaRun", ""]], false, "Check functions");
			Test("cot a", 2, [["ParaRun", ""], ["CMathFunc", "cot⁡a"], ["ParaRun", ""]], false, "Check functions");

			Test("sin (1+2_i) ", 2, [["ParaRun", ""], ["CMathFunc", "sin⁡(1+2_i)"], ["ParaRun", ""]], false, "Check functions");
			Test("cos (1+2_i) ", 2, [["ParaRun", ""], ["CMathFunc", "cos⁡(1+2_i)"], ["ParaRun", ""]], false, "Check functions");
			Test("tan (1+2_i) ", 2, [["ParaRun", ""], ["CMathFunc", "tan⁡(1+2_i)"], ["ParaRun", ""]], false, "Check functions");
			Test("csc (1+2_i) ", 2, [["ParaRun", ""], ["CMathFunc", "csc⁡(1+2_i)"], ["ParaRun", ""]], false, "Check functions");
			Test("sec (1+2_i) ", 2, [["ParaRun", ""], ["CMathFunc", "sec⁡(1+2_i)"], ["ParaRun", ""]], false, "Check functions");
			Test("cot (1+2_i) ", 2, [["ParaRun", ""], ["CMathFunc", "cot⁡(1+2_i)"], ["ParaRun", ""]], false, "Check functions");

			// Test("lim_a ", 2, [["ParaRun", ""], ["CMathFunc", "lim_a⁡"], ["ParaRun", ""]], false, "In one session we must save what type of token used for limit _ or ┬");
			// Test("lim┬a ", 2, [["ParaRun", ""], ["CMathFunc", "lim┬a⁡"], ["ParaRun", ""]], false, "In one session we must save what type of token used for limit _ or ┬");

			Test("log ", 2, [["ParaRun", ""], ["CMathFunc", "log⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("log⁡a ", 2, [["ParaRun", ""], ["CMathFunc", "log⁡a"], ["ParaRun", ""]], false, "Check functions");
			Test("log⁡(a+2) ", 2, [["ParaRun", ""], ["CMathFunc", "log⁡(a+2)"], ["ParaRun", ""]], false, "Check functions");

			Test("lim ", 2, [["ParaRun", ""], ["CMathFunc", "lim⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("lim_a ", 2, [["ParaRun", ""], ["CMathFunc", "lim┬a⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("lim^a ", 2, [["ParaRun", ""], ["CMathFunc", "lim┴a⁡"], ["ParaRun", ""]], false, "Check functions");

			Test("min ", 2, [["ParaRun", ""], ["CMathFunc", "min⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("min_a ", 2, [["ParaRun", ""], ["CMathFunc", "min┬a⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("min^a ", 2, [["ParaRun", ""], ["CMathFunc", "min┴a⁡"], ["ParaRun", ""]], false, "Check functions");

			Test("max ", 2, [["ParaRun", ""], ["CMathFunc", "max⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("max_a ", 2, [["ParaRun", ""], ["CMathFunc", "max┬a⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("max^a ", 2, [["ParaRun", ""], ["CMathFunc", "max┴a⁡"], ["ParaRun", ""]], false, "Check functions");

			Test("ln ", 2, [["ParaRun", ""], ["CMathFunc", "ln⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("ln_a ", 2, [["ParaRun", ""], ["CMathFunc", "ln┬a⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("ln^a ", 2, [["ParaRun", ""], ["CMathFunc", "ln┴a⁡"], ["ParaRun", ""]], false, "Check functions");
		})

		QUnit.module( "Matrix", function ()
		{
			// Test("■ ", 0, [["ParaRun", "■ "]], false, "Check matrix");
			// Test("■(1&2@3&4) ", 2, [["ParaRun", ""], ["CMathMatrix", "■(1&2@3&4)"], ["ParaRun", ""]], false, "Check matrix");
			Test("■(1&2) ", 2, [["ParaRun", ""], ["CMathMatrix", "■(1&2)"], ["ParaRun", ""]], false, "Check matrix");
			//Test("■(&1&2@3&4) ", 2, [["ParaRun", ""], ["CMathMatrix", "■(&1&2@3&4&)"], ["ParaRun", ""]], false, "Check matrix");
		})

		QUnit.module( "Accents", function ()
		{
			Test("e\\tilde  ", 2, [["ParaRun", ""], ["CAccent", "ẽ"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\hat  ", 2, [["ParaRun", ""], ["CAccent", "ê"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\breve  ", 2, [["ParaRun", ""], ["CAccent", "ĕ"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\dot  ", 2, [["ParaRun", ""], ["CAccent", "ė"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\ddot  ", 2, [["ParaRun", ""], ["CAccent", "ë"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\dddot  ", 2, [["ParaRun", ""], ["CAccent", "e⃛"], ["ParaRun", ""]], false, "Check diacritics");
			// Test("e\\prime  ", 2, [["ParaRun", ""], ["CAccent", "e′"], ["ParaRun", ""]], false, "Check diacritics");
			// Test("e\\pprime  ", 2, [["ParaRun", ""], ["CAccent", "e″"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\check  ", 2, [["ParaRun", ""], ["CAccent", "ě"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\acute  ", 2, [["ParaRun", ""], ["CAccent", "é"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\grave  ", 2, [["ParaRun", ""], ["CAccent", "è"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\bar  ", 2, [["ParaRun", ""], ["CAccent", "e̅"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\Bar  ", 2, [["ParaRun", ""], ["CAccent", "e̿"], ["ParaRun", ""]], false, "Check diacritics");
			// Test("e\\ubar  ", 2, [["ParaRun", ""], ["CAccent", "e̲"], ["ParaRun", ""]], false, "Check diacritics");
			// Test("e\\Ubar  ", 2, [["ParaRun", ""], ["CAccent", "e̳"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\vec  ", 2, [["ParaRun", ""], ["CAccent", "e⃗"], ["ParaRun", ""]], false, "Check diacritics");
		})

		QUnit.module( "Bugs", function ()
		{
			// QUnit.module( "Check bug #61007" );
			// Test("\\begin{matrix}1&2\\\\3&4\\\\\\end{matrix}", 2, [["ParaRun", ""], ["CMathMatrix", "\\begin{matrix}1&2\\\\3&4\\\\\\end{matrix}"]], true, "Check bug #61007 default matrix");
			// Test("\\begin{pmatrix}1&2\\\\3&4\\\\\\end{pmatrix}", 2, [["ParaRun", ""], ["CDelimiter", "\\begin{pmatrix}1&2\\\\3&4\\\\\\end{pmatrix}"]], true, "Check bug #61007 pmatrix");
			// Test("\\left[\\begin{matrix}1&2\\\\3&4\\\\\\end{matrix}\\right]", 2, [["ParaRun", ""], ["CDelimiter", "\\left[\\begin{matrix}1&2\\\\3&4\\\\\\end{matrix}\\right]"]], true, "Check bug #61007 pmatrix");

			// QUnit.module( "Check bug #67181" );
			// Test("\\mathbb{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "𝕢𝕨𝕖𝕣𝕥𝕪𝕦𝕚𝕠𝕡𝕒𝕤𝕕𝕗𝕘𝕙𝕛𝕜𝕝𝕫𝕩𝕔𝕧𝕓𝕟𝕞"]], true, "Check bug #67181", true);
			// Test("\\mathbb{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "\\mathbb{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");
			//
			// Test("\\mathcal{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "𝓆𝓌ℯ𝓇𝓉𝓎𝓊𝒾ℴ𝓅𝒶𝓈𝒹𝒻ℊ𝒽𝒿𝓀𝓁𝓏𝓍𝒸𝓋𝒷𝓃𝓂"]], true, "Check bug #67181", true);
			// Test("\\mathcal{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "\\mathcal{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");
			//
			// Test("\\mathsf{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "𝗊𝗐𝖾𝗋𝗍𝗒𝗎𝗂𝗈𝗉𝖺𝗌𝖽𝖿𝗀𝗁𝗃𝗄𝗅𝗓𝗑𝖼𝗏𝖻𝗇𝗆"]], true, "Check bug #67181", true);
			// Test("\\mathsf{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "\\mathsf{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");
			//
			// Test("\\mathrm{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "qwertyuiopasdfghjklzxcvbnm"]], true, "Check bug #67181", true);
			// Test("\\mathrm{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "qwertyuiopasdfghjklzxcvbnm"]], true, "Check bug #67181");
			//
			// Test("\\mathit{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "𝑞𝑤𝑒𝑟𝑡𝑦𝑢𝑖𝑜𝑝𝑎𝑠𝑑𝑓𝑔ℎ𝑗𝑘𝑙𝑧𝑥𝑐𝑣𝑏𝑛𝑚"]], true, "Check bug #67181", true);
			// Test("\\mathit{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "\\mathit{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");
			//
			// Test("\\mathfrak{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "𝔮𝔴𝔢𝔯𝔱𝔶𝔲𝔦𝔬𝔭𝔞𝔰𝔡𝔣𝔤𝔥𝔧𝔨𝔩𝔷𝔵𝔠𝔳𝔟𝔫𝔪"]], true, "Check bug #67181", true);
			// Test("\\mathfrak{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "\\mathfrak{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");
			//
			// Test("\\mathbfcal{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "𝓺𝔀𝓮𝓻𝓽𝔂𝓾𝓲𝓸𝓹𝓪𝓼𝓭𝓯𝓰𝓱𝓳𝓴𝓵𝔃𝔁𝓬𝓿𝓫𝓷𝓶"]], true, "Check bug #67181", true);
			// Test("\\mathbfcal{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "\\mathbfcal{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");
			//
			// Test("\\mathbf{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "𝐪𝐰𝐞𝐫𝐭𝐲𝐮𝐢𝐨𝐩𝐚𝐬𝐝𝐟𝐠𝐡𝐣𝐤𝐥𝐳𝐱𝐜𝐯𝐛𝐧𝐦"]], true, "Check bug #67181", true);
			// Test("\\mathbf{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "\\mathbf{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");
			//
			// Test("\\mathbb{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "𝕢𝕨𝕖𝕣𝕥𝕪𝕦𝕚𝕠𝕡𝕒𝕤𝕕𝕗𝕘𝕙𝕛𝕜𝕝𝕫𝕩𝕔𝕧𝕓𝕟𝕞"]], true, "Check bug #67181", true);
			// Test("\\mathbb{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "\\mathbb{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");
			//
			// // в будущем лучше объединять такие последовательности в один Run, когда будет готова работа со стилями перепроверить
			// Test("\\mathfrak{qwerty}\\mathfrak{uiopasdfghjklzxcvbnm}", 2, [["ParaRun", "𝔮𝔴𝔢𝔯𝔱𝔶"], ["ParaRun", ""],["ParaRun", "𝔲𝔦𝔬𝔭𝔞𝔰𝔡𝔣𝔤𝔥𝔧𝔨𝔩𝔷𝔵𝔠𝔳𝔟𝔫𝔪"]], true, "Check bug #67181", true);
			// Test("\\mathfrak{qwerty}\\mathfrak{uiopasdfghjklzxcvbnm}", 2, [["ParaRun", "\\mathfrak{qwerty}"], ["ParaRun", ""], ["ParaRun", "\\mathfrak{uiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");
			//
			// // non-standard for Word LaTeX operations
			// Test("\\fraktur{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "𝔮𝔴𝔢𝔯𝔱𝔶𝔲𝔦𝔬𝔭𝔞𝔰𝔡𝔣𝔤𝔥𝔧𝔨𝔩𝔷𝔵𝔠𝔳𝔟𝔫𝔪"]], true, "Check bug #67181 check non-standard", true);
			// Test("\\fraktur{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "\\mathfrak{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181 check non-standard");
			//
			// Test("\\sf{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "𝗊𝗐𝖾𝗋𝗍𝗒𝗎𝗂𝗈𝗉𝖺𝗌𝖽𝖿𝗀𝗁𝗃𝗄𝗅𝗓𝗑𝖼𝗏𝖻𝗇𝗆"]], true, "Check bug #67181 check non-standard", true);
			// Test("\\sf{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "\\mathsf{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181 check non-standard");
			//
			// Test("\\script{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "𝓆𝓌ℯ𝓇𝓉𝓎𝓊𝒾ℴ𝓅𝒶𝓈𝒹𝒻ℊ𝒽𝒿𝓀𝓁𝓏𝓍𝒸𝓋𝒷𝓃𝓂"]], true, "Check bug #67181 check non-standard", true);
			// Test("\\script{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "\\mathcal{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181 check non-standard");
			//
			// Test("\\double{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "𝕢𝕨𝕖𝕣𝕥𝕪𝕦𝕚𝕠𝕡𝕒𝕤𝕕𝕗𝕘𝕙𝕛𝕜𝕝𝕫𝕩𝕔𝕧𝕓𝕟𝕞"]], true, "Check bug #67181 check non-standard", true);
			// Test("\\double{qwertyuiopasdfghjklzxcvbnm}", 1, [["ParaRun", "\\mathbb{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181 check non-standard");
		})

	})

	QUnit.module( "LaTeX", function ()
	{
		// Test("\\dot{a}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\ddot{b}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\acute{c}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\grave{d}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\check{e}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		//
		// Test("\\breve{f}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\tilde{g}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\bar{h}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\widehat{j}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\vec{k}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\vec \\frac{k}{2}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("5''", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\frac{4}{5}''", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("(a)[b]\\{c\\}|d|\\|e\\|\\langlef\\rangle\\lfloorg\\rfloor\\lceilh\\rceil\\ulcorneri\\urcorner/j\\backslash", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("(2+1]", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\left.1+2\\right)", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("|2|+\\{1\\}+|2|", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("2^2", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("a^b", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("a^2", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("2^b", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("2_2", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("a_b", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("a_2", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("2_b", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test(`k_{n+1} = n^2 + k_n^2 - k_{n-1}`, 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test(`\\frac{1}{2}^{2}`, 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test(`\\frac{1}{2}_2`, 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\frac{1}{2}_2^y", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\frac{1}{2}_{2}^{y}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\frac{1}{2}_1_2_3_4_5_6_7", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\frac{1}{2}^1^2^3^4^5^6^7", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\frac{1}{2}^1^2^3^4^5^6^7_x", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\frac{1}{2}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\frac{1+\\frac{x}{y}}{2}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\frac{1^x}{2_y}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\sum^{2}_{x}4", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\int^2_x{4}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\binom{1}{2}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\sum_{i=1}^{10} t_i", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\exp_a b = a^b, \\exp b = e^b, 10^m, \\exp_{a}^x {b}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\ln c, \\lg d = \\log e, \\log_{10} f", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\sin a, \\cos b, \\tan c, \\cot d, \\sec e, \\csc f, \\cos^2_{y}{b}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\arcsin h, \\arccos_x i, \\arctan^y_{x} {j}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\sinhk, \\cosh {l}, \\tanh_x^y m, \\coth^{x}_y_1_2 {n}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\left\\vert s \\right\\vert", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\min(x,y), \\max(x,y)", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("0 \\leq \\lim_{n\\to \\infty}\\frac{n!}{(2n)!} \\leq \\lim_{n\\to \\infty} \\frac{n!}{(n!)^2} = \\lim_{k \\to \\infty, k = n!}\\frac{k}{k^2} = \\lim_{k \\to \\infty} \\frac{1}{k} = 0", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\sqrt5", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\sqrt\\frac{1}{2}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\sqrt[2^2]\\frac{1}{2}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\sqrt[2^2] {\\frac{1}{2}+3}", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\doubleAB", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\frakturAB", 1, [["ParaRun", ""]], true, "Check LaTeX words");
		// Test("\\dd", 1, [["ParaRun", "𝑑"]], true, "Check LaTeX words");
		// Test("\\Dd", 1, [["ParaRun", "𝐷"]], true, "Check LaTeX words");
		// Test("\\ee", 1, [["ParaRun", "𝑒"]], true, "Check LaTeX words");
		// Test("\\hbar", 1, [["ParaRun", "ℏ"]], true, "Check LaTeX words");
		// Test("\\ii", 1, [["ParaRun", "𝑖"]], true, "Check LaTeX words");
		// Test("\\Im", 1, [["ParaRun", "𝕴"]], true, "Check LaTeX words");
		// Test("\\imath", 1, [["ParaRun", "𝚤"]], true, "Check LaTeX words");
		// Test("\\j", 1, [["ParaRun", "𝐽𝑎𝑦"]], true, "Check LaTeX words");
		// Test("\\jj", 1, [["ParaRun", "𝑗"]], true, "Check LaTeX words");
		// Test("\\jmath", 1, [["ParaRun", "𝐽"]], true, "Check LaTeX words");
		// Test("\\partial", 1, [["ParaRun", "∂"]], true, "Check LaTeX words");
		// Test("\\Re", 1, [["ParaRun", "ℜ"]], true, "Check LaTeX words");
		// Test("\\wp", 1, [["ParaRun", "℘"]], true, "Check LaTeX words");
		// Test("\\aleph", 1, [["ParaRun", "ℵ"]], true, "Check LaTeX words");
		// Test("\\bet", 1, [["ParaRun", "ℶ"]], true, "Check LaTeX words");
		// Test("\\beth", 1, [["ParaRun", "ℶ"]], true, "Check LaTeX words");
		// Test("\\gimel", 1, [["ParaRun", "ℷ"]], true, "Check LaTeX words");
		//
		//
		// Test("\\alpha", 1, [["ParaRun", "\\alpha"]], true, "Check LaTeX words");
		// Test("\\Alpha", 1, [["ParaRun", "\\Alpha"]], true, "Check LaTeX words");
		// Test("\\beta", 1, [["ParaRun", "\\beta"]], true, "Check LaTeX words");
		// Test("\\Beta", 1, [["ParaRun", "\\Beta"]], true, "Check LaTeX words");
		// Test("\\gamma", 1, [["ParaRun", "\\gamma"]], true, "Check LaTeX words");
		// Test("\\Gamma", 1, [["ParaRun", "\\Gamma"]], true, "Check LaTeX words");
		// Test("\\pi", 1, [["ParaRun", "\\pi"]], true, "Check LaTeX words");
		// Test("\\Pi", 1, [["ParaRun", "\\Pi"]], true, "Check LaTeX words");
		// Test("\\phi", 1, [["ParaRun", "\\phi"]], true, "Check LaTeX words");
		// Test("\\varphi", 1, [["ParaRun", "\\varphi"]], true, "Check LaTeX words");
		// Test("\\mu", 1, [["ParaRun", "\\mu"]], true, "Check LaTeX words");
		// Test("\\Phi", 1, [["ParaRun", "\\Phi"]], true, "Check LaTeX words");
		//
		// Test("\\cos(2\\theta ) ", 2, [["ParaRun", ""], ["CMathFunc", "\\cos{\\left(2\\theta\\right)}"], ["ParaRun", ""]], true, "Check LaTeX function");
		// Test("\\lim_{x\\to \\infty }\\exp(x) ", 2, [["ParaRun", ""], ["CMathFunc", "\\lim_{x→∞} { \\exp { (x)}}"], ["ParaRun", ""]], true, "Check LaTeX function");
		//
		// Test("k^{n+1} ", 2, [["ParaRun", ""], ["CDegree", "k^{n+1}"], ["ParaRun", ""]], true, "Check LaTeX degree");
		// Test("n^2 ", 2, [["ParaRun", ""], ["CDegree", "n^2"], ["ParaRun", ""]], true, "Check LaTeX degree");
		// Test("n^{2} ", 2, [["ParaRun", ""], ["CDegree", "n^2"], ["ParaRun", ""]], true, "Check LaTeX degree");
		// Test("n^(2) ", 2, [["ParaRun", ""], ["CDegree", "n^{\\left(2\\right)}"], ["ParaRun", ""]], true, "Check LaTeX degree");
	})

 })


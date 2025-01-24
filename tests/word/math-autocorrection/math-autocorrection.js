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

    let Root, MathContent, logicDocument, p1;

    function Init() {
		logicDocument = AscTest.CreateLogicDocument();
		logicDocument.Start_SilentMode();
        logicDocument.RemoveFromContent(0, logicDocument.GetElementsCount(), false);

        p1 = new AscWord.Paragraph();
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
	function Test(str, arrResult, isLaTeX, strNameOfTest, isConvertAfter, isGetIntDifferentForm)
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

				if (isConvertAfter === true || isLaTeX)
					AscMath.SetAutoConvertation(false);

                AddText(str);

				if (isConvertAfter || isLaTeX)
					MathContent.ConvertView(false, isLaTeX ? Asc.c_oAscMathInputType.LaTeX : Asc.c_oAscMathInputType.Unicode);

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

	QUnit.testStart(function (){
		AscTest.ClearDocument();
		AscCommon.History.Clear();
		Clear();
		Init();
	})

	// TODO
	//  1. substack in LaTeX mode
	//
	//

	QUnit.module( "Unicode", function ()
	{
		QUnit.module( "Auto-convert rules", function ()
		{
			Test("4^2^2^2 ", [["ParaRun", "4^2^"], ["CDegree", "2^2"], ["ParaRun", ""]], false, "Check degree autocorrection rule");
			Test("4^2^2^2+", [["ParaRun", ""], ["CDegree", "4^(2^(2^2))"], ["ParaRun", "+"]], false, "Check degree autocorrection rule");

			Test("4┴2┴2┴2 ", [["ParaRun", "4┴2┴"], ["CLimit", "2┴2"], ["ParaRun", ""]], false, "Check Unicode AboveBelow");
			Test("4┴2┴2┴2+", [["ParaRun", ""], ["CLimit", "4┴(2┴(2┴2))"], ["ParaRun", "+"]], false, "Check Unicode AboveBelow");
		})

		QUnit.module( "AboveBelow", function ()
		{
			QUnit.module( "auto-convert");
			Test("4┴2 +2",  [["ParaRun", ""], ["CLimit", "4┴2"], ["ParaRun", "+2"]], false, "Check Unicode AboveBelow");
			Test("base┴ex *xz",  [["ParaRun", ""], ["CLimit", "base┴ex"], ["ParaRun", "*xz"]], false, "Check Unicode AboveBelow");
			Test("2┴ex -p", [["ParaRun", ""], ["CLimit", "2┴ex"], ["ParaRun", "-p"]], false, "Check Unicode AboveBelow");
			Test("base┬2 *x", [["ParaRun", ""], ["CLimit", "base┬2"], ["ParaRun", "*x"]], false, "Check Unicode AboveBelow");
			Test("4┬2 +x/y ", [["ParaRun", ""], ["CLimit", "4┬2"], ["ParaRun", "+"], ["CFraction", "x/y"]], false, "Check Unicode AboveBelow");
			Test("base┬x *y^2 ",  [["ParaRun", ""], ["CLimit", "base┬x"], ["ParaRun", "*"], ["CDegree", "y^2"]], false, "Check Unicode AboveBelow");
			Test("2┬ex -x_i ", [["ParaRun", ""], ["CLimit", "2┬ex"], ["ParaRun", "-"], ["CDegree", "x_i"]], false, "Check Unicode AboveBelow");
			Test("2┬(ex+2) +(2+1) ", [["ParaRun", ""], ["CLimit", "2┬(ex+2)"], ["ParaRun", "+"], ["CDelimiter", "(2+1)"]], false, "Check Unicode AboveBelow");
			Test("2┬(ex+2+x/2)^2 -1", [["ParaRun", ""], ["CLimit", "2┬((ex+2+x/2)^2)"], ["ParaRun", "-1"]], false, "Check Unicode AboveBelow");
			Test("(2+x)┬ex ", [["ParaRun", ""], ["CLimit", "(2+x)┬ex"]], false, "Check Unicode AboveBelow");
			Test("(2+y)┬(ex+2+x/2) ", [["ParaRun", ""], ["CLimit", "(2+y)┬(ex+2+x/2)"]], false, "Check Unicode AboveBelow");
			Test("(2+y^2)┬(ex_3+2+x/2) ", [["ParaRun", ""], ["CLimit", "(2+y^2)┬(ex_3+2+x/2)"]], false, "Check Unicode AboveBelow");

			QUnit.module("convert");
			Test("4┴2+2", [["ParaRun", ""], ["CLimit", "4┴2"], ["ParaRun", "+2"]], false, "Check Unicode AboveBelow",true);
			Test("base┴ex*xz", [["ParaRun", ""], ["CLimit", "base┴ex"], ["ParaRun", "*xz"]], false, "Check Unicode AboveBelow", true);
			Test("2┴ex-p", [["ParaRun", ""], ["CLimit", "2┴ex"], ["ParaRun", "-p"]], false, "Check Unicode AboveBelow", true);
			Test("base┬2*x", [["ParaRun", ""], ["CLimit", "base┬2"], ["ParaRun", "*x"]], false, "Check Unicode AboveBelow", true);
			Test("4┬2+x/y", [["ParaRun", ""], ["CLimit", "4┬2"], ["ParaRun", "+"], ["CFraction", "x/y"]], false, "Check Unicode AboveBelow", true);
			Test("base┬x*y^2", [["ParaRun", ""], ["CLimit", "base┬x"], ["ParaRun", "*"], ["CDegree", "y^2"]], false, "Check Unicode AboveBelow", true);
			Test("2┬ex-x_i", [["ParaRun", ""], ["CLimit", "2┬ex"], ["ParaRun", "-"], ["CDegree", "x_i"]], false, "Check Unicode AboveBelow", true);
			Test("2┬(ex+2)+(2+1)", [["ParaRun", ""], ["CLimit", "2┬(ex+2)"], ["ParaRun", "+"], ["CDelimiter", "(2+1)"]], false, "Check Unicode AboveBelow", true);
			Test("2┬(ex+2+x/2)^2-1", [["ParaRun", ""], ["CLimit", "2┬((ex+2+x/2)^2)"], ["ParaRun", "-1"]], false, "Check Unicode AboveBelow", true);
			Test("(2+x)┬ex", [["ParaRun", ""], ["CLimit", "(2+x)┬ex"]], false, "Check Unicode AboveBelow", true);
			Test("(2+y)┬(ex+2+x/2)", [["ParaRun", ""], ["CLimit", "(2+y)┬(ex+2+x/2)"]], false, "Check Unicode AboveBelow", true);
			Test("(2+y^2)┬(ex_3+2+x/2)", [["ParaRun", ""], ["CLimit", "(2+y^2)┬(ex_3+2+x/2)"]], false, "Check Unicode AboveBelow", true);

			Test("base┴2+2", [["ParaRun", ""], ["CLimit", "base┴2"], ["ParaRun", "+2"]], false, "Check diacritics");
			Test("base┴2┴x+2", [["ParaRun", ""], ["CLimit", "base┴(2┴x)"], ["ParaRun", "+2"]], false, "Check diacritics");
			Test("base┴2┴(x/y+6)+2", [["ParaRun", ""], ["CLimit", "base┴(2┴(x/y+6))"], ["ParaRun", "+2"]], false, "Check diacritics");

			Test("x^23┴2/y", [["ParaRun", ""], ["CLimit", "(x^23)┴2"], ["ParaRun", "/y"]], false, "Start autocorrection by divide");
			Test("(x^23)┴2/y", [["ParaRun", ""], ["CLimit", "(x^23)┴2"], ["ParaRun", "/y"]], false, "Start autocorrection by divide");
		})

		QUnit.module( "Box and Rect", function ()
		{
			QUnit.module("auto-convert");
			Test("□(1+2) ", [["ParaRun", ""], ["CBox", "□(1+2)"]], false, "Check Unicode Box");
			Test("□1 ", [["ParaRun", ""], ["CBox", "□1"]], false, "Check Unicode Box");
			Test("□ ", [["ParaRun", ""], ["CBox", "□"]], false, "Check Unicode Box");
			Test("□1/2 ", [["ParaRun", ""], ["CFraction", "□1/2"]], false, "Check Unicode Box");
			Test("▭(1+2) ", [["ParaRun", ""], ["CBorderBox", "▭(1+2)"]], false, "Check Unicode Box");
			Test("▭1 ", [["ParaRun", ""], ["CBorderBox", "▭1"]], false, "Check Unicode Box");
			Test("▭ ", [["ParaRun", ""], ["CBorderBox", "▭"]], false, "Check Unicode Box");
			Test("▭1/2 ", [["ParaRun", ""], ["CFraction", "▭1/2"]], false, "Check Unicode Box");

			QUnit.module("convert");
			Test("□(1+2)", [["ParaRun", ""], ["CBox", "□(1+2)"]], false, "Check Unicode Box", true);
			Test("□1", [["ParaRun", ""], ["CBox", "□1"]], false, "Check Unicode Box", true);
			Test("□", [["ParaRun", ""], ["CBox", "□"]], false, "Check Unicode Box", true);
			Test("□1/2", [["ParaRun", ""], ["CFraction", "□1/2"]], false, "Check Unicode Box", true);
			Test("▭(1+2)", [["ParaRun", ""], ["CBorderBox", "▭(1+2)"]], false, "Check Unicode Box", true);
			Test("▭1", [["ParaRun", ""], ["CBorderBox", "▭1"]], false, "Check Unicode Box", true);
			Test("▭", [["ParaRun", ""], ["CBorderBox", "▭"]], false, "Check Unicode Box", true);
			Test("▭1/2", [["ParaRun", ""], ["CFraction", "▭1/2"]], false, "Check Unicode Box", true);

			Test("\\rect ", [["ParaRun", "▭"]], false, "Check box literal");
			Test("\\rect 1/2 ", [["ParaRun", ""], ["CFraction", "▭1/2"], ["ParaRun", ""]], false, "Check box");
			Test("\\rect (1/2) ", [["ParaRun", ""], ["CBorderBox", "▭(1/2)"], ["ParaRun", ""]], false, "Check box");
			Test("\\rect (E=mc^2) ", [["ParaRun", ""], ["CBorderBox", "▭(E=mc^2)"], ["ParaRun", ""]], false, "Check box");
		})

		QUnit.module( "Underbar", function ()
		{
			QUnit.module("auto-convert");
			Test("▁(1+2) ", [["ParaRun", ""], ["CBar", "▁(1+2)"]], false, "Check Unicode underbar");
			Test("▁1 ", [["ParaRun", ""], ["CBar", "▁1"]], false, "Check Unicode underbar");
			Test("▁1/2 ", [["ParaRun", ""], ["CFraction", "▁1/2"]], false, "Check Unicode underbar");

			QUnit.module( "convert");
			Test("▁(1+2)", [["ParaRun", ""], ["CBar", "▁(1+2)"]], false, "Check Unicode underbar", true);
			Test("▁1", [["ParaRun", ""], ["CBar", "▁1"]], false, "Check Unicode underbar", true);
			Test("▁1/2", [["ParaRun", ""], ["CFraction", "▁1/2"]], false, "Check Unicode underbar", true);
		})

		QUnit.module( "Brackets", function ()
		{
			QUnit.module( "auto-convert Brackets");
			Test(`(1+2) +2`, [["ParaRun", ""], ["CDelimiter", "(1+2)"], ["ParaRun", "+2"]], false, "Check Unicode bracket", false);
			Test(`{1+2} -X`, [["ParaRun", ""], ["CDelimiter", "{1+2}"], ["ParaRun", "-X"]], false, "Check Unicode bracket", false);
			Test(`[1+2] *i`, [["ParaRun", ""], ["CDelimiter", "[1+2]"], ["ParaRun", "*i"]], false, "Check Unicode bracket", false);
			Test(`|1+2| -89/2 `, [["ParaRun", ""], ["CDelimiter", "|1+2|"], ["ParaRun", "-"], ["CFraction", "89/2"]], false, "Check Unicode bracket", false);
			Test(`|1+2| -〖89/2〗 `, [["ParaRun", ""], ["CDelimiter", "|1+2|"], ["ParaRun", "-"], ["CFraction", "89/2"]], false, "Check Unicode bracket", false);
			Test(`⌈1+2⌉ -〖89/2〗 `, [["ParaRun", ""], ["CDelimiter", "⌈1+2⌉"], ["ParaRun", "-"], ["CFraction", "89/2"]], false, "Check Unicode bracket", false);
			Test(`⌊1+2⌋ -〖89/2〗 `, [["ParaRun", ""], ["CDelimiter", "⌊1+2⌋"], ["ParaRun", "-"], ["CFraction", "89/2"]], false, "Check Unicode bracket", false);
			Test(`〖89/2〗/2 `, [["ParaRun", ""], ["CFraction", "(89/2)/2"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`√〖89/2〗 `, [["ParaRun", ""], ["CRadical", "√(89/2)"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`〖89/2〗_2 `, [["ParaRun", ""], ["CDegree", "(89/2)_2"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`〖89/2〗^2 `, [["ParaRun", ""], ["CDegree", "(89/2)^2"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2_〖89/2〗 `, [["ParaRun", ""], ["CDegree", "2_(89/2)"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2^〖89/2〗 `, [["ParaRun", ""], ["CDegree", "2^(89/2)"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2_〖89/2〗_2 `, [["ParaRun", "2_"], ["CDegree", "(89/2)_2"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2^〖89/2〗^2 `, [["ParaRun", "2^"], ["CDegree", "(89/2)^2"], ["ParaRun", ""]], false, "Check Unicode bracket",  false);
			Test(`2┴〖89/2〗 `, [["ParaRun", ""], ["CLimit", "2┴(89/2)"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2┴〖89/2〗┴2 `, [["ParaRun", "2┴"], ["CLimit", "(89/2)┴2"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2┬〖89/2〗 `, [["ParaRun", ""], ["CLimit", "2┬(89/2)"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`2┬〖89/2〗┬2 `, [["ParaRun", "2┬"], ["CLimit", "(89/2)┬2"], ["ParaRun", ""]], false, "Check Unicode bracket", false);
			Test(`├]a+b┤[ `, [["ParaRun", ""], ["CDelimiter", "├]a+b┤["], ["ParaRun", ""]], false, "Check Unicode bracket", false);

			Test("(", [["ParaRun", "("]], false);
			Test("[", [["ParaRun", "["]], false);
			Test("{", [["ParaRun", "{"]], false);

			Test("( ", [["ParaRun", "( "]], false);
			Test("[ ", [["ParaRun", "[ "]], false);
			Test("{ ", [["ParaRun", "{ "]], false);

			Test("(((", [["ParaRun", "((("]], false);
			Test("[[[", [["ParaRun", "[[["]], false);
			Test("{{{", [["ParaRun", "{{{"]], false);

			Test("((( ", [["ParaRun", "((( "]], false);
			Test("[[[ ", [["ParaRun", "[[[ "]], false);
			Test("{{{ ", [["ParaRun", "{{{ "]], false);

			Test("(((1", [["ParaRun", "(((1"]], false);
			Test("[[[1", [["ParaRun", "[[[1"]], false);
			Test("{{{1", [["ParaRun", "{{{1"]], false);

			Test("(((1 ", [["ParaRun", "(((1 "]], false);
			Test("[[[1 ", [["ParaRun", "[[[1 "]], false);
			Test("{{{1 ", [["ParaRun", "{{{1 "]], false);

			Test("1(((1", [["ParaRun", "1(((1"]], false);
			Test("1[[[1", [["ParaRun", "1[[[1"]], false);
			Test("1{{{1", [["ParaRun", "1{{{1"]], false);

			Test("1(((1 ", [["ParaRun", "1(((1 "]], false);
			Test("1[[[1 ", [["ParaRun", "1[[[1 "]], false);
			Test("1{{{1 ", [["ParaRun", "1{{{1 "]], false);

			Test("1(((1+", [["ParaRun", "1(((1+"]], false);
			Test("1[[[1+", [["ParaRun", "1[[[1+"]], false);
			Test("1{{{1+", [["ParaRun", "1{{{1+"]], false);
			Test("1(((1+=", [["ParaRun", "1(((1+="]], false);
			Test("1[[[1+=", [["ParaRun", "1[[[1+="]], false);
			Test("1{{{1+=", [["ParaRun", "1{{{1+="]], false);

			Test("1(((1+ ", [["ParaRun", "1(((1+ "]], false);
			Test("1[[[1+ ", [["ParaRun", "1[[[1+ "]], false);
			Test("1{{{1+ ", [["ParaRun", "1{{{1+ "]], false);
			Test("1(((1+= ", [["ParaRun", "1(((1+= "]], false);
			Test("1[[[1+= ", [["ParaRun", "1[[[1+= "]], false);
			Test("1{{{1+= ", [["ParaRun", "1{{{1+= "]], false);

			Test(")", [["ParaRun", ")"]], false);
			Test("]", [["ParaRun", "]"]], false);
			Test("}", [["ParaRun", "}"]], false);

			Test(") ", [["ParaRun", ") "]], false);
			Test("] ", [["ParaRun", "] "]], false);
			Test("} ", [["ParaRun", "} "]], false);

			Test(")))", [["ParaRun", ")))"]], false);
			Test("]]]", [["ParaRun", "]]]"]], false);
			Test("}}}", [["ParaRun", "}}}"]], false);

			Test("))) ", [["ParaRun", "))) "]], false);
			Test("]]] ", [["ParaRun", "]]] "]], false);
			Test("}}} ", [["ParaRun", "}}} "]], false);

			Test(")))1", [["ParaRun", ")))1"]], false);
			Test("]]]1", [["ParaRun", "]]]1"]], false);
			Test("}}}1", [["ParaRun", "}}}1"]], false);

			Test(")))1 ", [["ParaRun", ")))1 "]], false);
			Test("]]]1 ", [["ParaRun", "]]]1 "]], false);
			Test("}}}1 ", [["ParaRun", "}}}1 "]], false);

			Test("1)))1", [["ParaRun", "1)))1"]], false);
			Test("1]]]1", [["ParaRun", "1]]]1"]], false);
			Test("1}}}1", [["ParaRun", "1}}}1"]], false);

			Test("1)))1 ", [["ParaRun", "1)))1 "]], false);
			Test("1]]]1 ", [["ParaRun", "1]]]1 "]], false);
			Test("1}}}1 ", [["ParaRun", "1}}}1 "]], false);

			Test("1)))1+", [["ParaRun", "1)))1+"]], false);
			Test("1]]]1+", [["ParaRun", "1]]]1+"]], false);
			Test("1}}}1+", [["ParaRun", "1}}}1+"]], false);
			Test("1)))1+=", [["ParaRun", "1)))1+="]], false);
			Test("1]]]1+=", [["ParaRun", "1]]]1+="]], false);
			Test("1}}}1+=", [["ParaRun", "1}}}1+="]], false);

			Test("1)))1+ ", [["ParaRun", "1)))1+ "]], false);
			Test("1]]]1+ ", [["ParaRun", "1]]]1+ "]], false);
			Test("1}}}1+ ", [["ParaRun", "1}}}1+ "]], false);
			Test("1)))1+= ", [["ParaRun", "1)))1+= "]], false);
			Test("1]]]1+= ", [["ParaRun", "1]]]1+= "]], false);
			Test("1}}}1+= ", [["ParaRun", "1}}}1+= "]], false);

			Test("() ", [["ParaRun", ""], ["CDelimiter", "()"], ["ParaRun", ""]], false);
			Test("{} ", [["ParaRun", ""], ["CDelimiter", "{}"], ["ParaRun", ""]], false);
			Test("[] ", [["ParaRun", ""], ["CDelimiter", "[]"], ["ParaRun", ""]], false);
			Test("|| ", [["ParaRun", ""], ["CDelimiter", "||"], ["ParaRun", ""]], false);

			Test("()+", [["ParaRun", ""], ["CDelimiter", "()"], ["ParaRun", "+"]], false);
			Test("{}+", [["ParaRun", ""], ["CDelimiter", "{}"], ["ParaRun", "+"]], false);
			Test("[]+", [["ParaRun", ""], ["CDelimiter", "[]"], ["ParaRun", "+"]], false);
			Test("||+", [["ParaRun", ""], ["CDelimiter", "||"], ["ParaRun", "+"]], false);

			Test("|z}+", [["ParaRun", ""], ["CDelimiter", "|z}"], ["ParaRun", "+"]], false);
			Test("(c|+", [["ParaRun", ""], ["CDelimiter", "(c|"], ["ParaRun", "+"]], false);

			Test("(1+2)+", [["ParaRun", ""], ["CDelimiter", "(1+2)"], ["ParaRun", "+"]], false);
			Test("{1+2}+", [["ParaRun", ""], ["CDelimiter", "{1+2}"], ["ParaRun", "+"]], false);
			Test("[1+2]+", [["ParaRun", ""], ["CDelimiter", "[1+2]"], ["ParaRun", "+"]], false);
			Test("|1+2|+", [["ParaRun", ""], ["CDelimiter", "|1+2|"], ["ParaRun", "+"]], false);

			Test("(1/2 ", [["ParaRun", "("], ["CFraction", "1/2"]], false);
			Test("{1/2 ", [["ParaRun", "{"], ["CFraction", "1/2"]], false);
			Test("[1/2 ", [["ParaRun", "["], ["CFraction", "1/2"]], false);
			Test("|1/2 ", [["ParaRun", "|"], ["CFraction", "1/2"]], false);

			Test("(1/2)",[["ParaRun", "("], ["CFraction", "1/2"], ["ParaRun", ")"]], false);

			Test("(1\\mid 2\\mid 3) ", [["ParaRun", ""], ["CDelimiter", "(1∣2∣3)"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("[1\\mid 2\\mid 3) ", [["ParaRun", ""], ["CDelimiter", "[1∣2∣3)"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("|1\\mid 2\\mid 3) ", [["ParaRun", ""], ["CDelimiter", "|1∣2∣3)"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("{1\\mid 2\\mid 3) ", [["ParaRun", ""], ["CDelimiter", "{1∣2∣3)"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("(1\\mid 2\\mid 3] ", [["ParaRun", ""], ["CDelimiter", "(1∣2∣3]"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("(1\\mid 2\\mid 3} ", [["ParaRun", ""], ["CDelimiter", "(1∣2∣3}"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("(1\\mid 2\\mid 3| ", [["ParaRun", ""], ["CDelimiter", "(1∣2∣3|"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("|1\\mid 2\\mid 3| ", [["ParaRun", ""], ["CDelimiter", "|1∣2∣3|"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("{1\\mid 2\\mid 3} ", [["ParaRun", ""], ["CDelimiter", "{1∣2∣3}"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");
			Test("[1\\mid 2\\mid 3] ", [["ParaRun", ""], ["CDelimiter", "[1∣2∣3]"], ["ParaRun", ""]], false, "Check  Unicode bracket with mid");

			Test("(1+ ", [["ParaRun", "(1+ "]], false, "Check brackets");
			Test("(1+2) ", [["ParaRun", ""], ["CDelimiter", "(1+2)"], ["ParaRun", ""]], false, "Check brackets");
			Test("[1+2] ", [["ParaRun", ""], ["CDelimiter", "[1+2]"], ["ParaRun", ""]], false, "Check brackets");
			Test("{1+2} ", [["ParaRun", ""], ["CDelimiter", "{1+2}"], ["ParaRun", ""]], false, "Check brackets");

			Test(")123 ", [["ParaRun", ")123 "]], false, "Check brackets");
			Test(")12) ", [["ParaRun", ")12) "]], false, "Check brackets");
			Test(")12] ", [["ParaRun", ")12] "]], false, "Check brackets");
			Test(")12} ", [["ParaRun", ")12} "]], false, "Check brackets");

			Test("(1+2] ", [["ParaRun", ""], ["CDelimiter", "(1+2]"], ["ParaRun", ""]], false, "Check brackets");
			Test("|1+2] ", [["ParaRun", ""], ["CDelimiter", "|1+2]"], ["ParaRun", ""]], false, "Check brackets");
			Test("{1+2] ", [["ParaRun", ""], ["CDelimiter", "{1+2]"], ["ParaRun", ""]], false, "Check brackets");

			QUnit.module( " convert Brackets");
			Test(`(1+2)+2`, [["ParaRun", ""], ["CDelimiter", "(1+2)"], ["ParaRun", "+2"]], false, "Check Unicode bracket", true);
			Test(`{1+2}-X`, [["ParaRun", ""], ["CDelimiter", "{1+2}"], ["ParaRun", "-X"]], false, "Check Unicode bracket", true);
			Test(`[1+2]*i`, [["ParaRun", ""], ["CDelimiter", "[1+2]"], ["ParaRun", "*i"]], false, "Check Unicode bracket", true);
			Test(`|1+2|-89/2`, [["ParaRun", ""], ["CDelimiter", "|1+2|"], ["ParaRun", "-"], ["CFraction", "89/2"]], false, "Check Unicode bracket", true);
			Test(`|1+2|-〖89/2〗`, [["ParaRun", ""], ["CDelimiter", "|1+2|"], ["ParaRun", "-"], ["CFraction", "89/2"]], false, "Check Unicode bracket", true);
			Test(`⌈1+2⌉-〖89/2〗`, [["ParaRun", ""], ["CDelimiter", "⌈1+2⌉"], ["ParaRun", "-"], ["CFraction", "89/2"]], false, "Check Unicode bracket", true);
			Test(`⌊1+2⌋-〖89/2〗`, [["ParaRun", ""], ["CDelimiter", "⌊1+2⌋"], ["ParaRun", "-"], ["CFraction", "89/2"]], false, "Check Unicode bracket", true);
			Test(`〖89/2〗/2`, [["ParaRun", ""], ["CFraction", "(89/2)/2"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`√〖89/2〗`, [["ParaRun", ""], ["CRadical", "√(89/2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`〖89/2〗_2`, [["ParaRun", ""], ["CDegree", "(89/2)_2"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`〖89/2〗^2`, [["ParaRun", ""], ["CDegree", "(89/2)^2"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2_〖89/2〗`, [["ParaRun", ""], ["CDegree", "2_(89/2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2^〖89/2〗`, [["ParaRun", ""], ["CDegree", "2^(89/2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2_〖89/2〗_2`, [["ParaRun", ""], ["CDegree", "2_((89/2)_2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2^〖89/2〗^2`, [["ParaRun", ""], ["CDegree", "2^((89/2)^2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true );
			Test(`2┴〖89/2〗`, [["ParaRun", ""], ["CLimit", "2┴(89/2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2┴〖89/2〗┴2`, [["ParaRun", ""], ["CLimit", "2┴((89/2)┴2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2┬〖89/2〗`, [["ParaRun", ""], ["CLimit", "2┬(89/2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`2┬〖89/2〗┬2`, [["ParaRun", ""], ["CLimit", "2┬((89/2)┬2)"], ["ParaRun", ""]], false, "Check Unicode bracket", true);
			Test(`├]a+b┤[`, [["ParaRun", ""], ["CDelimiter", "├]a+b┤["], ["ParaRun", ""]], false, "Check Unicode bracket", true);

			Test(`〖1∣2〗`, [["ParaRun", ""], ["CDelimiter", "〖1∣2〗"], ["ParaRun", ""]], false, "Check special bracket with some contents", true);
			Test(`\\left(1\\right)`, [["ParaRun", ""], ["CDelimiter", "(1)"], ["ParaRun", ""]], false, "Is convert \\left and \\right without autocorrection", true);
			Test(`\\open(1\\close)`, [["ParaRun", ""], ["CDelimiter", "(1)"], ["ParaRun", ""]], false, "Is convert \\left and \\right without autocorrection", true);
		})

		QUnit.module( "Complex", function ()
		{
			QUnit.module( " convert Complex");
			//Test(`(a + b)^n =∑_(k=0)^n▒(n¦k) a^k  b^(n-k)  `, [["ParaRun", ""], ["CDegree", "(a + b)^n"], ["ParaRun", "="], ["CNary", "∑^n_(k=0)▒(n¦k)"],  ["ParaRun", ""], ["CDegree", "a^k"], ["CDegree", "b^(n-k)"]], false, "Check Complex content", true);
			Test(`∑_2^2▒(n/23)`, [["ParaRun", ""], ["CNary", "∑_2^2▒(n/23)"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`(x+⋯+x)^(k "times")`, [["ParaRun", ""], ["CDegree", "(x+⋯+x)^(k \"times\")"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`𝐸 = 𝑚𝑐^2`, [["ParaRun", "𝐸 = "], ["CDegree", "𝑚𝑐^2"], ["ParaRun", ""]], false, "Check Complex content", true);
			Test(`∫_0^a▒xⅆx/(x^2+a^2)`, [["ParaRun", ""], ["CNary", "∫_0^a▒〖xⅆx/(x^2+a^2)〗"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`lim┬(n→∞) a_n`, [["ParaRun", ""], ["CLimit", "lim┬(n→∞)⁡a_n"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`ⅈ²=-1`, [["ParaRun", ""], ["CDegree", "ⅈ²=-1"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`E = m⁢c²`, [["ParaRun", "E ="], ["CDegree", "〖 m⁢c〗^2"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`a²⋅b²=c²`, [["ParaRun", ""], ["CDegree", "a^2"], ["ParaRun", "⋅"],  ["CDegree", "b^2"], ["ParaRun", "="],  ["CDegree", "c^2"]], false, "Check Complex content", true);
			//Test(`f̂(ξ)=∫_-∞^∞▒f(x)ⅇ^-2πⅈxξ ⅆx`,[["ParaRun", ""], ["CAccent", "f̂"], ["CDelimiter", "(ξ)"], ["ParaRun", "="], ["CNary", "∫▒〖ⅇ^ⅈxξ ⅆx〗"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`(𝑎 + 𝑏)┴→`, [["ParaRun", ""], ["CLimit", "(𝑎 + 𝑏)┴→"], ["ParaRun", ""]], false, "Check Complex content", true);
			//Test(`𝑎┴→`, [["ParaRun", ""], ["CLimit", "𝑎┴→"], ["ParaRun", ""]], false, "Check Complex content", true);
		})

		QUnit.module( "Fractions", function ()
		{
			QUnit.module( " convert fractions");
			Test(`1/2`, [["ParaRun", ""], ["CFraction", "1/2"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`x+5/2`, [["ParaRun", "x+"], ["CFraction", "5/2"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`x+5/x+2`, [["ParaRun", "x+"], ["CFraction", "5/x"], ["ParaRun", "+2"]], false, "Check fraction content", true);
			Test(`1∕2`, [["ParaRun", ""], ["CFraction", "1∕2"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`(x+5)/2`, [["ParaRun", ""], ["CFraction", "(x+5)/2"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`x/(2+1)`, [["ParaRun", ""], ["CFraction", "x/(2+1)"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`(x-5)/(2+1)`, [["ParaRun", ""], ["CFraction", "(x-5)/(2+1)"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`1+3/2/3`, [["ParaRun", "1+"], ["CFraction", "3/(2/3)"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`(𝛼_2^3)/(𝛽_2^3+𝛾_2^3)`, [["ParaRun", ""], ["CFraction", "(𝛼_2^3)/(𝛽_2^3+𝛾_2^3)"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`(a/(b+c))/(d/e+f)`, [["ParaRun", ""], ["CFraction", "(a/(b+c))/(d/e+f)"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`(a/(c/(z/x)))`, [["ParaRun", ""], ["CDelimiter", "(a/(c/(z/x)))"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`1¦2`, [["ParaRun", ""], ["CFraction", "1¦2"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test(`(1¦2)`, [["ParaRun", ""], ["CDelimiter", "(1¦2)"], ["ParaRun", ""]], false, "Check fraction content", true);
			Test("(sin⁡θ)/(cos⁡θ) ", [["ParaRun", ""], ["CFraction", "(sin⁡θ)/(cos⁡θ)"], ["ParaRun", ""]], false, "Check functions");
			Test("(1/2)/", [["ParaRun", ""], ["CDelimiter", "(1/2)"], ["ParaRun", "/"]], false, "Check devide");
		})

		QUnit.module( "Horizontal brackets", function ()
		{
			QUnit.module( " convert hbrackets");
			Test(`⏞(x+⋯+x)`, [["ParaRun", ""], ["CGroupCharacter", "⏞(x+⋯+x)"], ["ParaRun", ""]], false, "Check hbrack content", true);
			Test(`⏞(x+⋯+x)^2`, [["ParaRun", ""], ["CLimit", "⏞(x+⋯+x)┴2"], ["ParaRun", ""]], false, "Check hbrack content", true);
			Test(`⏞(x+⋯+x)_2`, [["ParaRun", ""], ["CLimit", "⏞(x+⋯+x)┬2"], ["ParaRun", ""]], false, "Check hbrack content", true);
			Test(`⏞(x+⋯+x)_2^Y`, [["ParaRun", ""], ["CLimit", "⏞(x+⋯+x)┬(2^Y)"], ["ParaRun", ""]], false, "Check hbrack content", true);
			Test(`⏞(x+⋯+x)_2^2`, [["ParaRun", ""], ["CLimit", "⏞(x+⋯+x)┬(2^2)"], ["ParaRun", ""]], false, "Check hbrack content", true);
		})

		QUnit.module( "Autocorrection", function ()
		{
			Test("\\above", [["ParaRun", "\\above"]], false, "Check literal", true)
			Test("\\acute", [["ParaRun", "\\acute"]], false, "Check literal", true)
			Test("\\aleph", [["ParaRun", "\\aleph"]], false, "Check literal", true)
			Test("\\alpha", [["ParaRun", "\\alpha"]], false, "Check literal", true)
			Test("\\amalg", [["ParaRun", "\\amalg"]], false, "Check literal", true)
			Test("\\angle", [["ParaRun", "\\angle"]], false, "Check literal", true)
			Test("\\aoint", [["ParaRun", "\\aoint"]], false, "Check literal", true)
			Test("\\approx", [["ParaRun", "\\approx"]], false, "Check literal", true)
			Test("\\asmash", [["ParaRun", "\\asmash"]], false, "Check literal", true)
			Test("\\ast", [["ParaRun", "\\ast"]], false, "Check literal", true)
			Test("\\asymp", [["ParaRun", "\\asymp"]], false, "Check literal", true)
			Test("\\atop", [["ParaRun", "\\atop"]], false, "Check literal", true)
			Test("\\Bar", [["ParaRun", "\\Bar"]], false, "Check literal", true)
			Test("\\bar", [["ParaRun", "\\bar"]], false, "Check literal", true)
			Test("\\because",  [["ParaRun", "\\because"]], false, "Check literal", true)
			Test("\\begin", [["ParaRun", "\\begin"]], false, "Check literal", true)
			Test("\\below", [["ParaRun", "\\below"]], false, "Check literal", true)
			Test("\\beta", [["ParaRun", "\\beta"]], false, "Check literal", true)
			Test("\\beth", [["ParaRun", "\\beth"]], false, "Check literal", true)
			Test("\\bot", [["ParaRun", "\\bot"]], false, "Check literal", true)
			Test("\\bigcap", [["ParaRun", "\\bigcap"]], false, "Check literal", true)
			Test("\\bigcup", [["ParaRun", "\\bigcup"]], false, "Check literal", true)
			Test("\\bigodot", [["ParaRun", "\\bigodot"]], false, "Check literal", true)
			Test("\\bigoplus", [["ParaRun", "\\bigoplus"]], false, "Check literal", true)
			Test("\\bigotimes",  [["ParaRun", "\\bigotimes"]], false, "Check literal", true)
			Test("\\bigsqcup", [["ParaRun", "\\bigsqcup"]], false, "Check literal", true)
			Test("\\biguplus", [["ParaRun", "\\biguplus"]], false, "Check literal", true)
			Test("\\bigvee", [["ParaRun", "\\bigvee"]], false, "Check literal", true)
			Test("\\bigwedge", [["ParaRun", "\\bigwedge"]], false, "Check literal", true)
			Test("\\bowtie", [["ParaRun", "\\bowtie"]], false, "Check literal", true)
			Test("\\box", [["ParaRun", "\\box"]], false, "Check literal", true)
			Test("\\bra", [["ParaRun", "\\bra"]], false, "Check literal", true)
			Test("\\breve", [["ParaRun", "\\breve"]], false, "Check literal", true)
			Test("\\bullet", [["ParaRun", "\\bullet"]], false, "Check literal", true)
			Test("\\boxdot", [["ParaRun", "\\boxdot"]], false, "Check literal", true)
			Test("\\boxminus", [["ParaRun", "\\boxminus"]], false, "Check literal", true)
			Test("\\boxplus", [["ParaRun", "\\boxplus"]], false, "Check literal", true)
			Test("\\cap", [["ParaRun", "\\cap"]], false, "Check literal", true)
			Test("\\cbrt", [["ParaRun", "\\cbrt"]], false, "Check literal", true)
			Test("\\cdots", [["ParaRun", "\\cdots"]], false, "Check literal", true)
			Test("\\cdot", [["ParaRun", "\\cdot"]], false, "Check literal", true)
			Test("\\check", [["ParaRun", "\\check"]], false, "Check literal", true)
			Test("\\chi", [["ParaRun", "\\chi"]], false, "Check literal", true)
			Test("\\circ", [["ParaRun", "\\circ"]], false, "Check literal", true)
			Test("\\close", [["ParaRun", "\\close"]], false, "Check literal", true)
			Test("\\clubsuit", [["ParaRun", "\\clubsuit"]], false, "Check literal", true)
			Test("\\coint", [["ParaRun", "\\coint"]], false, "Check literal", true)
			Test("\\cong", [["ParaRun", "\\cong"]], false, "Check literal", true)
			Test("\\contain", [["ParaRun", "\\contain"]], false, "Check literal", true)
			Test("\\cup", [["ParaRun", "\\cup"]], false, "Check literal", true)
			Test("\\daleth", [["ParaRun", "\\daleth"]], false, "Check literal", true)
			Test("\\dashv", [["ParaRun", "\\dashv"]], false, "Check literal", true)
			Test("\\dd", [["ParaRun", "\\dd"]], false, "Check literal", true)
			Test("\\ddddot", [["ParaRun", "\\ddddot"]], false, "Check literal", true)
			Test("\\dddot", [["ParaRun", "\\dddot"]], false, "Check literal", true)
			Test("\\ddot", [["ParaRun", "\\ddot"]], false, "Check literal", true)
			Test("\\ddots", [["ParaRun", "\\ddots"]], false, "Check literal", true)
			Test("\\degree", [["ParaRun", "\\degree"]], false, "Check literal", true)
			Test("\\Delta", [["ParaRun", "\\Delta"]], false, "Check literal", true)
			Test("\\delta", [["ParaRun", "\\delta"]], false, "Check literal", true)
			Test("\\diamond", [["ParaRun", "\\diamond"]], false, "Check literal", true)
			Test("\\diamondsuit", [["ParaRun", "\\diamondsuit"]], false, "Check literal", true)
			Test("\\div", [["ParaRun", "\\div"]], false, "Check literal", true)
			Test("\\dot", [["ParaRun", "\\dot"]], false, "Check literal", true)
			Test("\\doteq", [["ParaRun", "\\doteq"]], false, "Check literal", true)
			Test("\\dots", [["ParaRun", "\\dots"]], false, "Check literal", true)
			Test("\\downarrow", [["ParaRun", "\\downarrow"]], false, "Check literal", true)
			Test("\\dsmash", [["ParaRun", "\\dsmash"]], false, "Check literal", true)
			Test("\\degc", [["ParaRun", "\\degc"]], false, "Check literal", true)
			Test("\\degf", [["ParaRun", "\\degf"]], false, "Check literal", true)
			Test("\\ee", [["ParaRun", "\\ee"]], false, "Check literal", true)
			Test("\\ell", [["ParaRun", "\\ell"]], false, "Check literal", true)
			Test("\\emptyset", [["ParaRun", "\\emptyset"]], false, "Check literal", true)
			Test("\\emsp", [["ParaRun", "\\emsp"]], false, "Check literal", true)
			Test("\\end", [["ParaRun", "\\end"]], false, "Check literal", true)
			Test("\\ensp", [["ParaRun", "\\ensp"]], false, "Check literal", true)
			Test("\\epsilon", [["ParaRun", "\\epsilon"]], false, "Check literal", true)
			Test("\\eqarray", [["ParaRun", "\\eqarray"]], false, "Check literal", true)
			Test("\\eqno", [["ParaRun", "\\eqno"]], false, "Check literal", true)
			Test("\\equiv", [["ParaRun", "\\equiv"]], false, "Check literal", true)
			Test("\\eta", [["ParaRun", "\\eta"]], false, "Check literal", true)
			Test("\\exists", [["ParaRun", "\\exists"]], false, "Check literal", true)
			Test("\\forall", [["ParaRun", "\\forall"]], false, "Check literal", true)
			Test("\\funcapply", [["ParaRun", "\\funcapply"]], false, "Check literal", true)
			Test("\\frown", [["ParaRun", "\\frown"]], false, "Check literal", true)
			Test("\\Gamma", [["ParaRun", "\\Gamma"]], false, "Check literal", true)
			Test("\\gamma", [["ParaRun", "\\gamma"]], false, "Check literal", true)
			Test("\\ge", [["ParaRun", "\\ge"]], false, "Check literal", true)
			Test("\\geq", [["ParaRun", "\\geq"]], false, "Check literal", true)
			Test("\\gets", [["ParaRun", "\\gets"]], false, "Check literal", true)
			Test("\\gg", [["ParaRun", "\\gg"]], false, "Check literal", true)
			Test("\\gimel", [["ParaRun", "\\gimel"]], false, "Check literal", true)
			Test("\\grave", [["ParaRun", "\\grave"]], false, "Check literal", true)
			Test("\\hairsp", [["ParaRun", "\\hairsp"]], false, "Check literal", true)
			Test("\\hat", [["ParaRun", "\\hat"]], false, "Check literal", true)
			Test("\\hbar", [["ParaRun", "\\hbar"]], false, "Check literal", true)
			Test("\\heartsuit", [["ParaRun", "\\heartsuit"]], false, "Check literal", true)
			Test("\\hookleftarrow", [["ParaRun", "\\hookleftarrow"]], false, "Check literal", true)
			Test("\\hphantom", [["ParaRun", "\\hphantom"]], false, "Check literal", true)
			Test("\\hsmash", [["ParaRun", "\\hsmash"]], false, "Check literal", true)
			Test("\\hvec", [["ParaRun", "\\hvec"]], false, "Check literal", true)
			Test("\\Im", [["ParaRun", "\\Im"]], false, "Check literal", true)
			Test("\\iiiint", [["ParaRun", "\\iiiint"]], false, "Check literal", true)
			Test("\\iiint", [["ParaRun", "\\iiint"]], false, "Check literal", true)
			Test("\\iint", [["ParaRun", "\\iint"]], false, "Check literal", true)
			Test("\\ii", [["ParaRun", "\\ii"]], false, "Check literal", true)
			Test("\\int", [["ParaRun", "\\int"]], false, "Check literal", true)
			Test("\\imath", [["ParaRun", "\\imath"]], false, "Check literal", true)
			Test("\\inc", [["ParaRun", "\\inc"]], false, "Check literal", true)
			Test("\\infty", [["ParaRun", "\\infty"]], false, "Check literal", true)
			Test("\\in", [["ParaRun", "\\in"]], false, "Check literal", true)
			Test("\\iota", [["ParaRun", "\\iota"]], false, "Check literal", true)
			Test("\\jj", [["ParaRun", "\\jj"]], false, "Check literal", true)
			Test("\\jmath", [["ParaRun", "\\jmath"]], false, "Check literal", true)
			Test("\\kappa", [["ParaRun", "\\kappa"]], false, "Check literal", true)
			Test("\\ket", [["ParaRun", "\\ket"]], false, "Check literal", true)
			Test("\\Longleftrightarrow", [["ParaRun", "\\Longleftrightarrow"]], false, "Check literal", true)
			Test("\\Longrightarrow", [["ParaRun", "\\Longrightarrow"]], false, "Check literal", true)
			Test("\\Lambda", [["ParaRun", "\\Lambda"]], false, "Check literal", true)
			Test("\\lambda", [["ParaRun", "\\lambda"]], false, "Check literal", true)
			Test("\\langle", [["ParaRun", "\\langle"]], false, "Check literal", true)
			Test("\\lbrack", [["ParaRun", "\\lbrack"]], false, "Check literal", true)
			Test("\\ldiv", [["ParaRun", "\\ldiv"]], false, "Check literal", true)
			Test("\\ldots", [["ParaRun", "\\ldots"]], false, "Check literal", true)
			Test("\\le", [["ParaRun", "\\le"]], false, "Check literal", true)
			Test("\\Leftarrow", [["ParaRun", "\\Leftarrow"]], false, "Check literal", true)
			Test("\\leftarrow", [["ParaRun", "\\leftarrow"]], false, "Check literal", true)
			Test("\\leftharpoondown", [["ParaRun", "\\leftharpoondown"]], false, "Check literal", true)
			Test("\\leftharpoonup", [["ParaRun", "\\leftharpoonup"]], false, "Check literal", true)
			Test("\\Leftrightarrow", [["ParaRun", "\\Leftrightarrow"]], false, "Check literal", true)
			Test("\\leftrightarrow", [["ParaRun", "\\leftrightarrow"]], false, "Check literal", true)
			Test("\\leq", [["ParaRun", "\\leq"]], false, "Check literal", true)
			Test("\\lfloor", [["ParaRun", "\\lfloor"]], false, "Check literal", true)
			Test("\\ll", [["ParaRun", "\\ll"]], false, "Check literal", true)
			Test("\\Longleftarrow", [["ParaRun", "\\Longleftarrow"]], false, "Check literal", true)
			Test("\\longleftarrow", [["ParaRun", "\\longleftarrow"]], false, "Check literal", true)
			Test("\\longleftrightarrow", [["ParaRun", "\\longleftrightarrow"]], false, "Check literal", true)
			Test("\\longrightarrow", [["ParaRun", "\\longrightarrow"]], false, "Check literal", true)
			Test("\\lmoust", [["ParaRun", "\\lmoust"]], false, "Check literal", true)
			Test("\\mapsto", [["ParaRun", "\\mapsto"]], false, "Check literal", true)
			Test("\\matrix", [["ParaRun", "\\matrix"]], false, "Check literal", true)
			Test("\\medsp", [["ParaRun", "\\medsp"]], false, "Check literal", true)
			Test("\\mid", [["ParaRun", "\\mid"]], false, "Check literal", true)
			Test("\\models", [["ParaRun", "\\models"]], false, "Check literal", true)
			Test("\\mp", [["ParaRun", "\\mp"]], false, "Check literal", true)
			Test("\\mu", [["ParaRun", "\\mu"]], false, "Check literal", true)
			Test("\\nabla", [["ParaRun", "\\nabla"]], false, "Check literal", true)
			Test("\\naryand", [["ParaRun", "\\naryand"]], false, "Check literal", true)
			Test("\\nbsp", [["ParaRun", "\\nbsp"]], false, "Check literal", true)
			Test("\\ndiv", [["ParaRun", "\\ndiv"]], false, "Check literal", true)
			Test("\\ne", [["ParaRun", "\\ne"]], false, "Check literal", true)
			Test("\\nearrow", [["ParaRun", "\\nearrow"]], false, "Check literal", true)
			Test("\\neg", [["ParaRun", "\\neg"]], false, "Check literal", true)
			Test("\\neq", [["ParaRun", "\\neq"]], false, "Check literal", true)
			Test("\\ni", [["ParaRun", "\\ni"]], false, "Check literal", true)
			Test("\\norm", [["ParaRun", "\\norm"]], false, "Check literal", true)
			Test("\\nu", [["ParaRun", "\\nu"]], false, "Check literal", true)
			Test("\\nwarrow", [["ParaRun", "\\nwarrow"]], false, "Check literal", true)
			Test("\\Omega", [["ParaRun", "\\Omega"]], false, "Check literal", true)
			Test("\\odot", [["ParaRun", "\\odot"]], false, "Check literal", true)
			Test("\\of", [["ParaRun", "\\of"]], false, "Check literal", true)
			Test("\\oiiint", [["ParaRun", "\\oiiint"]], false, "Check literal", true)
			Test("\\oiint", [["ParaRun", "\\oiint"]], false, "Check literal", true)
			Test("\\oint", [["ParaRun", "\\oint"]], false, "Check literal", true)
			Test("\\omega", [["ParaRun", "\\omega"]], false, "Check literal", true)
			Test("\\ominus", [["ParaRun", "\\ominus"]], false, "Check literal", true)
			Test("\\open", [["ParaRun", "\\open"]], false, "Check literal", true)
			Test("\\oplus", [["ParaRun", "\\oplus"]], false, "Check literal", true)
			Test("\\oslash", [["ParaRun", "\\oslash"]], false, "Check literal", true)
			Test("\\otimes", [["ParaRun", "\\otimes"]], false, "Check literal", true)
			Test("\\over", [["ParaRun", "\\over"]], false, "Check literal", true)
			Test("\\overbar", [["ParaRun", "\\overbar"]], false, "Check literal", true)
			Test("\\overbrace", [["ParaRun", "\\overbrace"]], false, "Check literal", true)
			Test("\\overbracket", [["ParaRun", "\\overbracket"]], false, "Check literal", true)
			Test("\\overparen", [["ParaRun", "\\overparen"]], false, "Check literal", true)
			Test("\\overshell", [["ParaRun", "\\overshell"]], false, "Check literal", true)
			Test("\\over", [["ParaRun", "\\over"]], false, "Check literal", true)
			Test("\\Pi", [["ParaRun", "\\Pi"]], false, "Check literal", true)
			Test("\\Phi", [["ParaRun", "\\Phi"]], false, "Check literal", true)
			Test("\\Psi", [["ParaRun", "\\Psi"]], false, "Check literal", true)
			Test("\\parallel", [["ParaRun", "\\parallel"]], false, "Check literal", true)
			Test("\\partial", [["ParaRun", "\\partial"]], false, "Check literal", true)
			Test("\\perp", [["ParaRun", "\\perp"]], false, "Check literal", true)
			Test("\\phantom", [["ParaRun", "\\phantom"]], false, "Check literal", true)
			Test("\\phi", [["ParaRun", "\\phi"]], false, "Check literal", true)
			Test("\\pi", [["ParaRun", "\\pi"]], false, "Check literal", true)
			Test("\\pm", [["ParaRun", "\\pm"]], false, "Check literal", true)
			Test("\\pppprime", [["ParaRun", "\\pppprime"]], false, "Check literal", true)
			Test("\\ppprime", [["ParaRun", "\\ppprime"]], false, "Check literal", true)
			Test("\\pprime", [["ParaRun", "\\pprime"]], false, "Check literal", true)
			Test("\\prcue", [["ParaRun", "\\prcue"]], false, "Check literal", true)
			Test("\\prec", [["ParaRun", "\\prec"]], false, "Check literal", true)
			Test("\\preceq", [["ParaRun", "\\preceq"]], false, "Check literal", true)
			Test("\\preccurlyeq", [["ParaRun", "\\preccurlyeq"]], false, "Check literal", true)
			Test("\\prime", [["ParaRun", "\\prime"]], false, "Check literal", true)
			Test("\\propto", [["ParaRun", "\\propto"]], false, "Check literal", true)
			Test("\\psi", [["ParaRun", "\\psi"]], false, "Check literal", true)
			Test("\\qdrt", [["ParaRun", "\\qdrt"]], false, "Check literal", true)
			Test("\\Re", [["ParaRun", "\\Re"]], false, "Check literal", true)
			Test("\\Rightarrow", [["ParaRun", "\\Rightarrow"]], false, "Check literal", true)
			Test("\\rangle", [["ParaRun", "\\rangle"]], false, "Check literal", true)
			Test("\\ratio", [["ParaRun", "\\ratio"]], false, "Check literal", true)
			Test("\\rbrace", [["ParaRun", "\\rbrace"]], false, "Check literal", true)
			Test("\\rbrack", [["ParaRun", "\\rbrack"]], false, "Check literal", true)
			Test("\\rceil", [["ParaRun", "\\rceil"]], false, "Check literal", true)
			Test("\\rddots", [["ParaRun", "\\rddots"]], false, "Check literal", true)
			Test("\\rect", [["ParaRun", "\\rect"]], false, "Check literal", true)
			Test("\\rfloor", [["ParaRun", "\\rfloor"]], false, "Check literal", true)
			Test("\\rho", [["ParaRun", "\\rho"]], false, "Check literal", true)
			Test("\\right", [["ParaRun", "\\right"]], false, "Check literal", true)
			Test("\\rightarrow", [["ParaRun", "\\rightarrow"]], false, "Check literal", true)
			Test("\\rightharpoondown", [["ParaRun", "\\rightharpoondown"]], false, "Check literal", true)
			Test("\\rightharpoonup", [["ParaRun", "\\rightharpoonup"]], false, "Check literal", true)
			Test("\\rmoust", [["ParaRun", "\\rmoust"]], false, "Check literal", true)
			Test("\\rrect", [["ParaRun", "\\rrect"]], false, "Check literal", true)
			Test("\\root", [["ParaRun", "\\root"]], false, "Check literal", true)
			Test("\\Sigma", [["ParaRun", "\\Sigma"]], false, "Check literal", true)
			Test("\\sdiv", [["ParaRun", "\\sdiv"]], false, "Check literal", true)
			Test("\\searrow", [["ParaRun", "\\searrow"]], false, "Check literal", true)
			Test("\\setminus", [["ParaRun", "\\setminus"]], false, "Check literal", true)
			Test("\\sigma", [["ParaRun", "\\sigma"]], false, "Check literal", true)
			Test("\\sim", [["ParaRun", "\\sim"]], false, "Check literal", true)
			Test("\\simeq", [["ParaRun", "\\simeq"]], false, "Check literal", true)
			Test("\\smash", [["ParaRun", "\\smash"]], false, "Check literal", true)
			Test("\\smile", [["ParaRun", "\\smile"]], false, "Check literal", true)
			Test("\\spadesuit", [["ParaRun", "\\spadesuit"]], false, "Check literal", true)
			Test("\\sqcap", [["ParaRun", "\\sqcap"]], false, "Check literal", true)
			Test("\\sqcup", [["ParaRun", "\\sqcup"]], false, "Check literal", true)
			Test("\\sqrt", [["ParaRun", "\\sqrt"]], false, "Check literal", true)
			Test("\\sqsubseteq", [["ParaRun", "\\sqsubseteq"]], false, "Check literal", true)
			Test("\\sqsuperseteq", [["ParaRun", "\\sqsuperseteq"]], false, "Check literal", true)
			Test("\\star", [["ParaRun", "\\star"]], false, "Check literal", true)
			Test("\\subset", [["ParaRun", "\\subset"]], false, "Check literal", true)
			Test("\\subseteq", [["ParaRun", "\\subseteq"]], false, "Check literal", true)
			Test("\\succeq", [["ParaRun", "\\succeq"]], false, "Check literal", true)
			Test("\\succ", [["ParaRun", "\\succ"]], false, "Check literal", true)
			Test("\\sum", [["ParaRun", "\\sum"]], false, "Check literal", true)
			Test("\\superset", [["ParaRun", "\\superset"]], false, "Check literal", true)
			Test("\\superseteq", [["ParaRun", "\\superseteq"]], false, "Check literal", true)
			Test("\\swarrow", [["ParaRun", "\\swarrow"]], false, "Check literal", true)
		})

		QUnit.module( "Degree", function ()
		{
			QUnit.module( "convert");
			Test("2^2 + 2", [["ParaRun", ""], ["CDegree", "2^2"], ["ParaRun", "+ 2"]], false, "Check scripts", true)
			Test("x^2+2", [["ParaRun", ""], ["CDegree", "x^2"], ["ParaRun", "+2"]], false, "Check scripts", true)
			Test("x^(256+34)*y", [["ParaRun", ""], ["CDegree", "x^(256+34)"], ["ParaRun", "*y"]], false, "Check scripts", true)
			Test("(x+34)^(256+34)-y/x",[["ParaRun", ""], ["CDegree", "(x+34)^(256+34)"], ["ParaRun", "-"], ["CFraction", "y/x"]], false, "Check scripts", true)

			Test("2_1", [["ParaRun", "2_1"]], false);
			Test("2_1 ", [["ParaRun", ""], ["CDegree", "2_1"], ["ParaRun", ""]], false);
			Test("\\int", [["ParaRun", "\\int"]], false);
			Test("\\int _x^y\\of 1/2 ", [["ParaRun", "∫_x^y▒"], ["CFraction", "1/2"], ["ParaRun", ""]], false);
			Test("1/2 ", [["ParaRun", ""], ["CFraction", "1/2"], ["ParaRun", ""]], false);
			Test("1/2 +", [["ParaRun", ""], ["CFraction", "1/2"], ["ParaRun", "+"]], false);
			Test("1/2=", [["ParaRun", ""], ["CFraction", "1/2"], ["ParaRun", "="]], false);
			Test("1/2+1/2=x/y ", [["ParaRun", ""], ["CFraction", "1/2"], ["ParaRun", "+"], ["CFraction", "1/2"], ["ParaRun", "="], ["CFraction", "x/y"], ["ParaRun", ""]], false);

			Test("x_y ", [["ParaRun", ""], ["CDegree", "x_y"], ["ParaRun", ""]], false, "Check degree");
			Test("x_ ", [["ParaRun", ""], ["CDegree", "x_"], ["ParaRun", ""]], false, "Check degree");
			Test("_ ", [["ParaRun", ""], ["CDegree", "_"]], false, "Check degree");
			Test("x_1 ", [["ParaRun", ""], ["CDegree", "x_1"], ["ParaRun", ""]], false, "Check degree");
			Test("1_x ", [["ParaRun", ""], ["CDegree", "1_x"], ["ParaRun", ""]], false, "Check degree");
			Test("x_(1+2) ", [["ParaRun", ""], ["CDegree", "x_(1+2)"], ["ParaRun", ""]], false, "Check degree");
			Test("x_[1+2] ", [["ParaRun", ""], ["CDegree", "x_[1+2]"], ["ParaRun", ""]], false, "Check degree");
			Test("x_[1+2} ", [["ParaRun", ""], ["CDegree", "x_[1+2}"], ["ParaRun", ""]], false, "Check degree");
			Test("x_1/2", [["ParaRun", ""], ["CDegree", "x_1"], ["ParaRun", "/2"]], false, "Check degree");
			Test("x_1/2 ", [["ParaRun", ""], ["CFraction", "(x_1)/2"], ["ParaRun", ""]], false, "Check degree");

			QUnit.module( "autocorrect");
			Test("^ ", [["ParaRun", ""], ["CDegree", "^"]], false, "Check index");
			Test("x^y ", [["ParaRun", ""], ["CDegree", "x^y"], ["ParaRun", ""]], false, "Check index");
			Test("x^' ", [["ParaRun", ""], ["CDegree", "x^'"], ["ParaRun", ""]], false, "Check index");
			Test("x^\" ", [["ParaRun", ""], ["CDegree", "x^\""], ["ParaRun", ""]], false, "Check index");
			Test("x^1 ", [["ParaRun", ""], ["CDegree", "x^1"], ["ParaRun", ""]], false, "Check index");
			Test("1^x ", [["ParaRun", ""], ["CDegree", "1^x"], ["ParaRun", ""]], false, "Check index");
			Test("x^(1+2) ", [["ParaRun", ""], ["CDegree", "x^(1+2)"], ["ParaRun", ""]], false, "Check index");
			Test("x^[1+2] ", [["ParaRun", ""], ["CDegree", "x^[1+2]"], ["ParaRun", ""]], false, "Check index");
			Test("x^[1+2} ", [["ParaRun", ""], ["CDegree", "x^[1+2}"], ["ParaRun", ""]], false, "Check index");
			Test("x^1/2", [["ParaRun", ""], ["CDegree", "x^1"], ["ParaRun", "/2"]], false, "Check index");
			Test("x^1/2 ", [["ParaRun", ""], ["CFraction", "(x^1)/2"], ["ParaRun", ""]], false, "Check index");
			Test("x^'=", [["ParaRun", ""], ["CDegree", "x^'"], ["ParaRun", "="]], false, "Check index");

			Test("x^y_1 ", [["ParaRun", ""], ["CDegreeSubSup", "x_1^y"], ["ParaRun", ""]], false, "Check index degree");
			Test("x^1_i ", [["ParaRun", ""], ["CDegreeSubSup", "x_i^1"], ["ParaRun", ""]], false, "Check index degree");
			Test("1^x_y ", [["ParaRun", ""], ["CDegreeSubSup", "1_y^x"], ["ParaRun", ""]], false, "Check index degree");
			Test("x^[1+2]_[g_i] ", [["ParaRun", ""], ["CDegreeSubSup", "x_[g_i]^[1+2]"], ["ParaRun", ""]], false, "Check index degree");
			Test("x^[1+2}_[6+1} ", [["ParaRun", ""], ["CDegreeSubSup", "x_[6+1}^[1+2}"], ["ParaRun", ""]], false, "Check index degree");
			//Test("x^1/2_1/2 ", [["ParaRun", ""], ["CFraction", "x^1/(2_1/2)"], ["ParaRun", ""]], false, "Check index degree");
			Test("𝑊^3𝛽_𝛿1𝜌1𝜎2 ", [["ParaRun", ""], ["CDegreeSubSup", "𝑊_𝛿1𝜌1𝜎2^3𝛽"], ["ParaRun", ""]], false, "Check index degree with Unicode symbols");

			QUnit.module( "pre-script");
			Test("(_1^f)f ", [["ParaRun", ""], ["CDegreeSubSup", "(_1^f)f"], ["ParaRun", ""]], false, "Check prescript index degree");
			Test("_1^f x ", [["ParaRun", ""], ["CDegreeSubSup", "(_1^f)x"], ["ParaRun", ""]], false, "Check prescript index degree");
			Test("(_(1/2)^y)f ", [["ParaRun", ""], ["CDegreeSubSup", "(_(1/2)^y)f"], ["ParaRun", ""]], false, "Check prescript index degree");
			Test("(_(1/2)^[x_i])x ", [["ParaRun", ""], ["CDegreeSubSup", "(_(1/2)^[x_i])x"], ["ParaRun", ""]], false, "Check prescript index degree");
		})

		QUnit.module( "Radicals", function ()
		{
			QUnit.module( " convert radicals");
			Test("√5", [["ParaRun", ""], ["CRadical", "√5"],["ParaRun", ""]], false, "Check special", true)
			Test("√a", [["ParaRun", ""], ["CRadical", "√a"],["ParaRun", ""]], false, "Check special", true)
			Test("√a/2", [["ParaRun", ""], ["CFraction", "√a/2"],["ParaRun", ""]], false, "Check special", true)
			Test("√(2&a-4)", [["ParaRun", ""], ["CRadical", "√(2&a-4)"],["ParaRun", ""]], false, "Check special", true)
			Test("∛5", [["ParaRun", ""], ["CRadical", "∛5"],["ParaRun", ""]], false, "Check special", true)
			Test("∛a", [["ParaRun", ""], ["CRadical", "∛a"],["ParaRun", ""]], false, "Check special", true)
			Test("∛a/2", [["ParaRun", ""], ["CFraction", "∛a/2"],["ParaRun", ""]], false, "Check special", true)
			Test("∛(a-4)", [["ParaRun", ""], ["CRadical", "∛(a-4)"],["ParaRun", ""]], false, "Check special", true)
			Test("∜5", [["ParaRun", ""], ["CRadical", "∜5"],["ParaRun", ""]], false, "Check special", true)
			Test("∜a", [["ParaRun", ""], ["CRadical", "∜a"],["ParaRun", ""]], false, "Check special", true)
			Test("∜a/2", [["ParaRun", ""], ["CFraction", "∜a/2"],["ParaRun", ""]], false, "Check special", true)
			Test("∜(a-4)", [["ParaRun", ""], ["CRadical", "∜(a-4)"],["ParaRun", ""]], false, "Check special", true)
			Test("√(10&a/4)", [["ParaRun", ""], ["CRadical", "√(10&a/4)"],["ParaRun", ""]], false, "Check special", true)
			Test("√(10^2&a/4+2)", [["ParaRun", ""], ["CRadical", "√(10^2&a/4+2)"],["ParaRun", ""]], false, "Check special", true)
			Test("√5^2", [["ParaRun", ""], ["CRadical", "√(5^2)"],["ParaRun", ""]], false, "Check special", true)
			Test("√5_2", [["ParaRun", ""], ["CRadical", "√(5_2)"],["ParaRun", ""]], false, "Check special", true)
			Test("√5^2_x", [["ParaRun", ""], ["CRadical", "√(5_x^2)"],["ParaRun", ""]], false, "Check special", true)
			Test("√5_2^x", [["ParaRun", ""], ["CRadical", "√(5_2^x)"],["ParaRun", ""]], false, "Check special", true)
			Test("(_5^2)√5", [["ParaRun", ""], ["CDegreeSubSup", "(_5^2)√5"],["ParaRun", ""]], false, "Check special", true)
			Test("√5┴exp1", [["ParaRun", ""], ["CRadical", "√(5┴exp1)"],["ParaRun", ""]], false, "Check special", true)
			Test("√5┬exp1", [["ParaRun", ""], ["CRadical", "√(5┬exp1)"],["ParaRun", ""]], false, "Check special", true)
			Test("(√5┬exp1]", [["ParaRun", ""], ["CDelimiter", "(√(5┬exp1)]"],["ParaRun", ""]], false, "Check special", true)
			//Test("□√5", [["ParaRun", ""], ["CBox", "□(√5)"],["ParaRun", ""]], false, "Check special", true)
			//Test("▭√5", [["ParaRun", ""], ["CRect", "▭(√5)"],["ParaRun", ""]], false, "Check special", true)
			//Test("▁√5", [["ParaRun", ""], ["CBar", "▁(√5)"],["ParaRun", ""]], false, "Check special", true)
			//Test(`¯√5`, [["ParaRun", ""], ["CBar", "¯(√5)"],["ParaRun", ""]], false, "Check special", true)
			Test("∑_√5^√5", [["ParaRun", ""], ["CNary", "∑_√(5^√5)"],["ParaRun", ""]], false, "Check special", true)

			Test("\\sqrt ", [["ParaRun", "√"]], false, "Check");
			Test("\\sqrt (2&1+2) ", [["ParaRun", ""], ["CRadical", "√(2&1+2)"], ["ParaRun", ""]], false, "Check radical");
			Test("\\sqrt (1+2) ", [["ParaRun", ""], ["CRadical", "√(1+2)"], ["ParaRun", ""]], false, "Check radical");
			Test("√1 ", [["ParaRun", ""], ["CRadical", "√1"], ["ParaRun", ""]], false, "Check radical");

			Test("\\cbrt ", [["ParaRun", "∛"]], false, "Check");
			Test("\\cbrt (1+2) ", [["ParaRun", ""], ["CRadical", "∛(1+2)"], ["ParaRun", ""]], false, "Check radical");
			Test("\\cbrt 1/2 ", [["ParaRun", ""], ["CFraction", "∛1/2"], ["ParaRun", ""]], false, "Check radical");
			Test("∛1 ", [["ParaRun", ""], ["CRadical", "∛1"], ["ParaRun", ""]], false, "Check radical");
			Test("∛(1) ", [["ParaRun", ""], ["CRadical", "∛1"], ["ParaRun", ""]], false, "Check radical");

			Test("\\qdrt ", [["ParaRun", "∜"]], false, "Check");
			Test("\\qdrt (1+2) ", [["ParaRun", ""], ["CRadical", "∜(1+2)"], ["ParaRun", ""]], false, "Check radical");
			Test("\\qdrt 1/2 ", [["ParaRun", ""], ["CFraction", "∜1/2"], ["ParaRun", ""]], false, "Check radical");
			Test("∜1 ", [["ParaRun", ""], ["CRadical", "∜1"], ["ParaRun", ""]], false, "Check radical");
			Test("∜(1) ", [["ParaRun", ""], ["CRadical", "∜1"], ["ParaRun", ""]], false, "Check radical");
		})

		QUnit.module( "Other", function ()
		{
			QUnit.module( " convert operators");
			Test("×", [["ParaRun", "×"]], false, "Check literal", true)
			Test("⋅", [["ParaRun", "⋅"]], false, "Check literal")
			Test("∈", [["ParaRun", "∈"]], false, "Check literal")
			Test("∋", [["ParaRun", "∋"]], false, "Check literal")
			Test("∼", [["ParaRun", "∼"]], false, "Check literal")
			Test("≃", [["ParaRun", "≃"]], false, "Check literal")
			Test("≅", [["ParaRun", "≅"]], false, "Check literal")
			Test("≈", [["ParaRun", "≈"]], false, "Check literal")
			Test("≍", [["ParaRun", "≍"]], false, "Check literal")
			Test("≡", [["ParaRun", "≡"]], false, "Check literal")
			Test("≤", [["ParaRun", "≤"]], false, "Check literal")
			Test("≥", [["ParaRun", "≥"]], false, "Check literal")
			Test("≶", [["ParaRun", "≶"]], false, "Check literal")
			Test("≷", [["ParaRun", "≷"]], false, "Check literal")
			Test("≽", [["ParaRun", "≽"]], false, "Check literal")
			Test("≺", [["ParaRun", "≺"]], false, "Check literal")
			Test("≻", [["ParaRun", "≻"]], false, "Check literal")
			Test("≼", [["ParaRun", "≼"]], false, "Check literal")
			Test("⊂", [["ParaRun", "⊂"]], false, "Check literal")
			Test("⊃", [["ParaRun", "⊃"]], false, "Check literal")
			Test("⊆", [["ParaRun", "⊆"]], false, "Check literal")
			Test("⊇", [["ParaRun", "⊇"]], false, "Check literal")
			Test("⊑", [["ParaRun", "⊑"]], false, "Check literal")
			Test("⊒", [["ParaRun", "⊒"]], false, "Check literal")
			Test("+", [["ParaRun", "+"]], false, "Check literal")
			Test("-", [["ParaRun", "-"]], false, "Check literal")
			Test("=", [["ParaRun", "="]], false, "Check literal")
			Test("*", [["ParaRun", "*"]], false, "Check literal")
			Test("∃", [["ParaRun", "∃"]], false, "Check literal")
			Test("∀", [["ParaRun", "∀"]], false, "Check literal")
			Test("¬", [["ParaRun", "¬"]], false, "Check literal")
			Test("∧", [["ParaRun", "∧"]], false, "Check literal")
			Test("∨", [["ParaRun", "∨"]], false, "Check literal")
			Test("⇒", [["ParaRun", "⇒"]], false, "Check literal")
			Test("⇔", [["ParaRun", "⇔"]], false, "Check literal")
			Test("⊕", [["ParaRun", "⊕"]], false, "Check literal")
			Test("⊤", [["ParaRun", "⊤"]], false, "Check literal")
			Test("⊥", [["ParaRun", "⊥"]], false, "Check literal")
			Test("⊢", [["ParaRun", "⊢"]], false, "Check literal")
			Test("⨯", [["ParaRun", "⨯"]], false, "Check literal")
			Test("⟕", [["ParaRun", "⟕"]], false, "Check literal")
			Test("⟖", [["ParaRun", "⟖"]], false, "Check literal")
			Test("⟗", [["ParaRun", "⟗"]], false, "Check literal")
			Test("⋉", [["ParaRun", "⋉"]], false, "Check literal")
			Test("⋊", [["ParaRun", "⋊"]], false, "Check literal")
			Test("▷", [["ParaRun", "▷"]], false, "Check literal")
			Test("÷", [["ParaRun", "÷"]], false, "Check literal")
			Test("⁡", [["ParaRun", "⁡"]], false, "Check literal")
			Test("⁢", [["ParaRun", "⁢"]], false, "Check literal")
			Test("⁣", [["ParaRun", "⁣"]], false, "Check literal")
			Test("⁤", [["ParaRun", "⁤"]], false, "Check literal")
			Test("​", [["ParaRun", "​"]], false, "Check literal")
			Test(" ", [["ParaRun", " "]], false, "Check literal")
			Test("  ", [["ParaRun", "  "]], false, "Check literal")
			Test(" ", [["ParaRun", " "]], false, "Check literal")
			Test(" ", [["ParaRun", " "]], false, "Check literal")
			Test(" ", [["ParaRun", " "]], false, "Check literal")
			Test("  ", [["ParaRun", "  "]], false, "Check literal")
			Test(" ", [["ParaRun", " "]], false, "Check literal")
			Test(" ", [["ParaRun", " "]], false, "Check literal")
			Test(" ", [["ParaRun", " "]], false, "Check literal")
			Test(" ", [["ParaRun", " "]], false, "Check literal")
			Test(`a`, [["ParaRun", "a"]], false, "Check literal")
			Test(`abcdef`, [["ParaRun", "abcdef"]], false, "Check literal")
			Test(`1`, [["ParaRun", "1"]], false, "Check literal")
			Test(`1234`, [["ParaRun", "1234"]], false, "Check literal")
			Test(`1+2`, [["ParaRun", "1+2"]], false, "Check literal")
			Test(`1+2+3`, [["ParaRun", "1+2+3"]], false, "Check literal")
			Test(`ΑαΒβΓγΔδΕεΖζΗηΘθΙιΚκΛλΜμΝνΞξΟοΠπΡρΣσΤτΥυΦφΧχΨψΩω`, [["ParaRun", "ΑαΒβΓγΔδΕεΖζΗηΘθΙιΚκΛλΜμΝνΞξΟοΠπΡρΣσΤτΥυΦφΧχΨψΩω"]], false, "Check literal")
			Test("abc123def", [["ParaRun", "abc123def"]], false, "Check literal")
			Test("abc+123+def", [["ParaRun", "abc+123+def"]], false, "Check literal")
			Test("𝐀𝐁𝐂𝐨𝐹", [["ParaRun", "𝐀𝐁𝐂𝐨𝐹"]], false, "Check literal")
			Test("   𝐀𝐁𝐂𝐨𝐹   ", [["ParaRun", "   𝐀𝐁𝐂𝐨𝐹   "]], false, "Check literal")
			Test(" 	𝐀𝐁𝐂𝐨𝐹  	 ", [["ParaRun", " 	𝐀𝐁𝐂𝐨𝐹  	 "]], false, "Check literal")
			Test(`1+fbnd+(3+𝐀𝐁𝐂𝐨𝐹)+c+5`, [["ParaRun", "1+fbnd+"], ["CDelimiter", "(3+𝐀𝐁𝐂𝐨𝐹)"], ["ParaRun", "+c+5"]], false, "Check literal")
			Test(`1/3.1416 `, [["ParaRun", ""], ["CFraction", "1/3.1416"],["ParaRun", ""]], false, "Check literal")
			Test("1\\above 2 ", [["ParaRun", ""], ["CLimit", "1┴2"]], false, "Check literal")
			Test("1\\acute 2 ", [["ParaRun", ""], ["CAccent", "1́"], ["ParaRun", "2"]], false, "Check literal")

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
			Test("\\int ", [["ParaRun", "∫"]], false, "Check large operators");
			Test("\\int  ", [["ParaRun", ""], ["CNary", "∫"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int _x ", [["ParaRun", ""], ["CNary", "∫_x"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int ^x ", [["ParaRun", ""], ["CNary", "∫^x"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int ^(x+1) ", [["ParaRun", ""], ["CNary", "∫^(x+1)"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int ^(x+1) ", [["ParaRun", ""], ["CNary", "∫^(x+1)"], ["ParaRun", ""]],false, "Check large operators");
			Test("\\int ^(x+1)_(1_i) ", [["ParaRun", ""], ["CNary", "∫_(1_i)^(x+1)"], ["ParaRun", ""]], false, "Check large operators");

			Test("\\int \\of x ", [["ParaRun", ""], ["CNary", "∫▒x"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int _x\\of 1/2  ", [["ParaRun", ""], ["CNary", "∫_x▒〖1/2〗"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int ^x\\of 1/2  ", [["ParaRun", ""], ["CNary", "∫^x▒〖1/2〗"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\int _(x+1)\\of 1/2  ", [["ParaRun", ""], ["CNary", "∫_(x+1)▒〖1/2〗"], ["ParaRun", ""]], false, "Check large operators");
			Test("\\prod ^(x+1)\\of 1/2  ", [["ParaRun", ""], ["CNary", "∏^(x+1)▒〖1/2〗"], ["ParaRun", ""]],false, "Check large operators");
			Test("∫^(x+1)_(1_i)\\of 1/2  ", [["ParaRun", ""], ["CNary", "∫_(1_i)^(x+1)▒〖1/2〗"], ["ParaRun", ""]], false, "Check large operators");
			Test("∑_(k=0)^n▒〖(n¦k) a^k b^(n-k)〗   ", [["ParaRun", ""], ["CNary", "∑_(k=0)^n▒〖(n¦k) a^k b^(n-k)〗"], ["ParaRun", " "]], false, "Check add space after nary");
		})

		QUnit.module( "Functions", function ()
		{
			Test("sin ", [["ParaRun", ""], ["CMathFunc", "sin⁡"]], false, "Check functions");
			Test("cos ", [["ParaRun", ""], ["CMathFunc", "cos⁡"]], false, "Check functions");
			Test("tan ", [["ParaRun", ""], ["CMathFunc", "tan⁡"]], false, "Check functions");
			Test("csc ", [["ParaRun", ""], ["CMathFunc", "csc⁡"]], false, "Check functions");
			Test("sec ", [["ParaRun", ""], ["CMathFunc", "sec⁡"]], false, "Check functions");
			Test("cot ", [["ParaRun", ""], ["CMathFunc", "cot⁡"]], false, "Check functions");

			Test("sin", [["ParaRun", "sin"]], false, "Check functions");
			Test("cos", [["ParaRun", "cos"]], false, "Check functions");
			Test("tan", [["ParaRun", "tan"]], false, "Check functions");
			Test("csc", [["ParaRun", "csc"]], false, "Check functions");
			Test("sec", [["ParaRun", "sec"]], false, "Check functions");
			Test("cot", [["ParaRun", "cot"]], false, "Check functions");

			Test("sin a", [["ParaRun", ""], ["CMathFunc", "sin⁡a"], ["ParaRun", ""]], false, "Check functions");
			Test("cos a", [["ParaRun", ""], ["CMathFunc", "cos⁡a"], ["ParaRun", ""]], false, "Check functions");
			Test("tan a", [["ParaRun", ""], ["CMathFunc", "tan⁡a"], ["ParaRun", ""]], false, "Check functions");
			Test("csc a", [["ParaRun", ""], ["CMathFunc", "csc⁡a"], ["ParaRun", ""]], false, "Check functions");
			Test("sec a", [["ParaRun", ""], ["CMathFunc", "sec⁡a"], ["ParaRun", ""]], false, "Check functions");
			Test("cot a", [["ParaRun", ""], ["CMathFunc", "cot⁡a"], ["ParaRun", ""]], false, "Check functions");

			Test("sin⁡(1+2_i)", [["ParaRun", ""], ["CMathFunc", "sin⁡(1+2_i)"], ["ParaRun", ""]], false, "Check functions", true);
			Test("cos⁡(1+2_i)", [["ParaRun", ""], ["CMathFunc", "cos⁡(1+2_i)"], ["ParaRun", ""]], false, "Check functions", true);
			Test("tan⁡(1+2_i)", [["ParaRun", ""], ["CMathFunc", "tan⁡(1+2_i)"], ["ParaRun", ""]], false, "Check functions", true);
			Test("csc⁡(1+2_i)", [["ParaRun", ""], ["CMathFunc", "csc⁡(1+2_i)"], ["ParaRun", ""]], false, "Check functions", true);
			Test("sec⁡(1+2_i)", [["ParaRun", ""], ["CMathFunc", "sec⁡(1+2_i)"], ["ParaRun", ""]], false, "Check functions", true);
			Test("cot⁡(1+2_i)", [["ParaRun", ""], ["CMathFunc", "cot⁡(1+2_i)"], ["ParaRun", ""]], false, "Check functions", true);

			// Test("lim_a ", [["ParaRun", ""], ["CMathFunc", "lim_a⁡"], ["ParaRun", ""]], false, "In one session we must save what type of token used for limit _ or ┬");
			// Test("lim┬a ", [["ParaRun", ""], ["CMathFunc", "lim┬a⁡"], ["ParaRun", ""]], false, "In one session we must save what type of token used for limit _ or ┬");

			Test("log ", [["ParaRun", ""], ["CMathFunc", "log⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("log⁡a ", [["ParaRun", ""], ["CMathFunc", "log⁡a"], ["ParaRun", ""]], false, "Check functions");
			Test("log⁡(a+2) ", [["ParaRun", ""], ["CMathFunc", "log⁡(a+2)"], ["ParaRun", ""]], false, "Check functions");

			Test("lim ", [["ParaRun", ""], ["CMathFunc", "lim⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("lim_a ", [["ParaRun", ""], ["CMathFunc", "lim┬a⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("lim^a ", [["ParaRun", ""], ["CMathFunc", "lim┴a⁡"], ["ParaRun", ""]], false, "Check functions");

			Test("min ", [["ParaRun", ""], ["CMathFunc", "min⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("min_a ", [["ParaRun", ""], ["CMathFunc", "min┬a⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("min^a ", [["ParaRun", ""], ["CMathFunc", "min┴a⁡"], ["ParaRun", ""]], false, "Check functions");

			Test("max ", [["ParaRun", ""], ["CMathFunc", "max⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("max_a ", [["ParaRun", ""], ["CMathFunc", "max┬a⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("max^a ", [["ParaRun", ""], ["CMathFunc", "max┴a⁡"], ["ParaRun", ""]], false, "Check functions");

			Test("ln ", [["ParaRun", ""], ["CMathFunc", "ln⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("ln_a ", [["ParaRun", ""], ["CMathFunc", "ln┬a⁡"], ["ParaRun", ""]], false, "Check functions");
			Test("ln^a ", [["ParaRun", ""], ["CMathFunc", "ln┴a⁡"], ["ParaRun", ""]], false, "Check functions");
		})

		QUnit.module( "Matrix", function ()
		{
			Test("■ ", [["ParaRun", "■ "]], false, "Check matrix");
			Test("■(1&2@3&4) ", [["ParaRun", ""], ["CMathMatrix", "■(1&2@3&4)"], ["ParaRun", ""]], false, "Check matrix");
			Test("■(1&2) ", [["ParaRun", ""], ["CMathMatrix", "■(1&2)"], ["ParaRun", ""]], false, "Check matrix");
			Test("■(&2&3@4&5) ", [["ParaRun", ""], ["CMathMatrix", "■(&2&3@4&5&)"], ["ParaRun", ""]], false, "Check matrix");
			Test("■(&&&@@@) ", [["ParaRun", ""], ["CMathMatrix", "■(&&&@&&&@&&&@&&&)"], ["ParaRun", ""]], false, "Check matrix");
		})

		QUnit.module( "Accents", function ()
		{
			Test("e\\tilde  ", [["ParaRun", ""], ["CAccent", "ẽ"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\hat  ", [["ParaRun", ""], ["CAccent", "ê"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\breve  ", [["ParaRun", ""], ["CAccent", "ĕ"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\dot  ", [["ParaRun", ""], ["CAccent", "ė"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\ddot  ", [["ParaRun", ""], ["CAccent", "ë"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\dddot  ", [["ParaRun", ""], ["CAccent", "e⃛"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\prime  ", [["ParaRun", ""], ["CDegree", "e^′"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\pprime  ", [["ParaRun", ""], ["CDegree", "e^″"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\check  ", [["ParaRun", ""], ["CAccent", "ě"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\acute  ", [["ParaRun", ""], ["CAccent", "é"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\grave  ", [["ParaRun", ""], ["CAccent", "è"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\bar  ", [["ParaRun", ""], ["CAccent", "e̅"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\Bar  ", [["ParaRun", ""], ["CAccent", "e̿"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\ubar  ", [["ParaRun", ""], ["CAccent", "e̲"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\Ubar  ", [["ParaRun", ""], ["CAccent", "e̳"], ["ParaRun", ""]], false, "Check diacritics");
			Test("e\\vec  ", [["ParaRun", ""], ["CAccent", "e⃗"], ["ParaRun", ""]], false, "Check diacritics");
		})

		QUnit.test('Binomial auto-correction', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);
			AddText('\\binomial ');
			assert.ok(true, "Add text '\\binomial'");
			
			let strBinomial = MathContent.GetTextOfElement(0).GetText();
			assert.strictEqual(strBinomial, '(a+b)^n=∑_(k=0)^n ▒(n¦k)a^k b^(n-k)', 'Check text of binomial');

			AddText(' ');
			assert.ok(true, "Add space and trigger auto-correction");

			let r = MathContent.Root;
			let arr = ['ParaRun', 'CDelimiter', 'ParaRun', 'CDegree', 'ParaRun', 'CDegree', 'ParaRun'];

			for (let i = 0; i < r.Content.length; i++)
			{
				assert.strictEqual(r.Content[i].constructor.name, arr[i], r.Content[i].constructor.name + "is " + arr[i]);
			}

			AddText(' ');
			assert.ok(true, "Add space and trigger auto-correction second time");

			let nary = r.Content[1];
			assert.strictEqual(nary.constructor.name, 'CNary', 'Result is one Nary');

			MathContent.ConvertView(true, Asc.c_oAscMathInputType.Unicode);
			assert.ok(true, "Convert to linear view");

			let str = MathContent.Root.GetTextOfElement().GetText();
			assert.strictEqual(str, '(a+b)^n=∑_(k=0)^n▒〖(n¦k) a^k b^(n-k)〗', 'Check linear form is "(a+b)^n=∑_(k=0)^n▒〖(n¦k) a^k b^(n-k)〗"');
		})

		QUnit.module( "Bugs", function ()
		{
			QUnit.test('Check correct word while input rBrackets', function (assert)
			{
				Clear();
				logicDocument.SetMathInputType(0);
				AddText('(1->\\infty)');
				assert.ok(true, "Add text '(1->\\infty)'");

				MathContent.ConvertView(true, Asc.c_oAscMathInputType.Unicode);
				assert.ok(true, "Convert to linear view");

				let strBinomial = MathContent.GetTextOfElement(0).GetText();
				assert.strictEqual(strBinomial, '(1→∞)', 'Check \\infty');
			})

			QUnit.test('Check absolute brackets inside normal brackets', function (assert)
			{
				Clear();
				logicDocument.SetMathInputType(0);
				AddText('(x|y|z)');
				assert.ok(true, "Add text '(x|y|z)'");

				MathContent.ConvertView(false, Asc.c_oAscMathInputType.Unicode);
				assert.ok(true, "Convert to professional view");

				let oWrapDelimiter	= MathContent.Root.Content[1];
				let oDelContent		= oWrapDelimiter.Content[0];
				let oFirstParaRun	= oDelContent.Content[0];
				let oInnerDel		= oDelContent.Content[1];
				let oSecondParaRun	= oDelContent.Content[2];

				assert.strictEqual(oFirstParaRun.GetTextOfElement(0).GetText(), 'x', 'Check first text is x');
				assert.strictEqual(oInnerDel.GetTextOfElement(0).GetText(), '|y|', 'Check del is |y|');
				assert.strictEqual(oSecondParaRun.GetTextOfElement(0).GetText(), 'z', 'Check second text is z');

				let strBinomial = MathContent.GetTextOfElement(0).GetText();
				assert.strictEqual(strBinomial, '(x|y|z)', 'Check');
			})

			QUnit.test('Check Dirac notion', function (assert)
			{
				// in future find better algorithm
				//
				// 		"...Another case where we treat | as a close delimiter is if it is followed by a space
				// 		(U+0020). This handles the important case of the bra vector ⟨ | in Dirac notation..."
				//
				// now we don't need to add spaces to handle these brackets correctly

				Clear();
				logicDocument.SetMathInputType(0);

				AscMath.SetAutoConvertation(false);
				AddText('p=∑_ψ▒〖P_ψ |ψ⟩ ⟨ψ|〗');
				assert.ok(true, "Add text 'p=∑_ψ▒〖P_ψ |ψ⟩ ⟨ψ|〗'");

				MathContent.ConvertView(false, Asc.c_oAscMathInputType.Unicode);
				assert.ok(true, "Convert to professional view");

				let strBinomial = MathContent.GetTextOfElement(0).GetText();
				assert.strictEqual(strBinomial, 'p=∑_ψ▒〖P_ψ |ψ⟩ ⟨ψ|〗', 'Check');

				MathContent.ConvertView(true, Asc.c_oAscMathInputType.Unicode);
				assert.ok(true, "Convert to professional view");

				strBinomial = MathContent.GetTextOfElement(0).GetText();
				assert.strictEqual(strBinomial, 'p=∑_ψ▒〖P_ψ |ψ⟩ ⟨ψ|〗', 'Check');

				AscMath.SetAutoConvertation(true);
			})

			QUnit.test('Check eqarray', function (assert)
			{
				Clear();
				logicDocument.SetMathInputType(0);
				AddText('{█(a, n odd@|a|, n even)┤');
				assert.ok(true, "Add text '{█(a, n odd@|a|, n even)┤'");

				MathContent.ConvertView(false, Asc.c_oAscMathInputType.Unicode);
				assert.ok(true, "Convert to linear view");

				let strBinomial = MathContent.GetTextOfElement(0).GetText();
				assert.strictEqual(strBinomial, '{█(a, n odd@|a|, n even)┤', 'Check');
			})

			QUnit.test('Check eqarray frac', function (assert)
			{
				Clear();
				logicDocument.SetMathInputType(0);
				AddText('█(1@█(@█(@█(@))))/2');
				assert.ok(true, "Add text '█(1@█(@█(@█(@))))/2'");

				MathContent.ConvertView(false, Asc.c_oAscMathInputType.Unicode);
				assert.ok(true, "Convert to linear view");

				let strBinomial = MathContent.GetTextOfElement(0).GetText();
				assert.strictEqual(strBinomial, '█(1@█(@█(@█(@))))/2', 'Check');
			})

			QUnit.test('Check review info convert math; bug #67505', function (assert)
			{
				Clear();
				logicDocument.SetMathInputType(0);
				AddText('(1+2)');
				assert.ok(true, "Add text '(1+2)'");

				let r = MathContent.Root.Content[0];										//	(1
				let r2 = r.Split2(2, MathContent.Root, 0);					//	+
				let r3 = r2.Split2(1, MathContent.Root, 1);					//	2)

				let reviewInfo = new AscWord.ReviewInfo();
				reviewInfo.UserId   = "this.UserId";
				reviewInfo.UserName = "this.UserName";
				reviewInfo.DateTime = new Date().toDateString();
				r2.SetReviewTypeWithInfo(reviewtype_Add, reviewInfo);

				assert.ok(true, "Split run and set ReviewType for '+' === reviewtype_Add");

				MathContent.ConvertView(false, Asc.c_oAscMathInputType.Unicode);
				assert.ok(true, "Convert to professional view");

				let rOne = MathContent.Root.Content[1].Content[0].Content[0];
				assert.strictEqual(rOne.GetReviewType(), reviewtype_Common, 'Is "1" is reviewtype_Common');

				let rPlus = MathContent.Root.Content[1].Content[0].Content[1];
				assert.strictEqual(rPlus.GetReviewType(), reviewtype_Add, 'Is "+" is reviewtype_Add');
				assert.strictEqual(rPlus.GetReviewInfo(), reviewInfo, 'reviewInfo');

				MathContent.ConvertView(true, Asc.c_oAscMathInputType.Unicode);
				assert.ok(true, "Convert to linear view");

				let nRPlus = MathContent.Root.Content[1];
				assert.strictEqual(nRPlus.GetReviewType(), reviewtype_Add, 'Is "+" is reviewtype_Add');
				assert.strictEqual(nRPlus.GetReviewInfo(), reviewInfo, 'Check reviewInfo');
			})

			QUnit.test('Bug 64357', function (assert)
			{
				Clear();

				logicDocument.SetMathInputType(0);
				let matrix = MathContent.Root.Add_Matrix(new CTextPr(), 1, 2, false, []);
				let oContent = matrix.getContentElement(0, 0);
				assert.ok(true, "Add matrix for example");

				let r = new ParaRun(matrix.Paragraph, true);
				r.AddMathPlaceholder();
				oContent.Add_Element(r);
				oContent.Select_Element(r);
				matrix.Paragraph.Select_Math(MathContent);
				assert.ok(true, "Select content inside matrix content");

				logicDocument.TurnOff_Recalculate();
				logicDocument.ConvertMathView(true);
				logicDocument.TurnOn_Recalculate();

				assert.ok(true, "Convert to linear view");

				let isNotMatrix = !(MathContent.Root.Content[1] instanceof CMathMatrix);
				assert.strictEqual(isNotMatrix, true, 'We must get not Matrix');
			})

			QUnit.test('Save manual break', function (assert)
			{
				Clear();
				assert.ok(true, "Set Unicode mode");

				logicDocument.SetMathInputType(0);
				AddText("1/2+x_2 ");
				assert.ok(true, "Add '1/2+x_2' and convert it");

				AscTest.MoveCursorLeft(false, false, 6);
				assert.ok(true, "Move cursor to start of '+'");

				let oBase = new CMathMenuBase();
				oBase.insert_ManualBreak();
				logicDocument.Set_MathProps(oBase);
				assert.ok(true, "Add manual break");

				assert.strictEqual(MathContent.Root.Content[2].MathPrp.brk !== undefined, true, 'Check brk in "+" ParaRun');

				logicDocument.ConvertMathView(true);
				assert.ok(true, "Convert to linear view");

				assert.strictEqual(MathContent.Root.Content[1].MathPrp.brk !== undefined, true, 'Check brk in "+" ParaRun');

				logicDocument.ConvertMathView(false);
				assert.ok(true, "Convert to professional view");
				assert.strictEqual(MathContent.Root.Content[2].MathPrp.brk !== undefined, true, 'Check brk in "+" ParaRun');
			})

			QUnit.test('Check complex content in math func', function (assert)
			{
				Clear();
				logicDocument.SetMathInputType(0);

				AddText('cos  \\theta ');

				assert.ok(true, "Convert to linear view");

				let strFunc = MathContent.GetTextOfElement(0).GetText();
				assert.strictEqual(strFunc, 'cos⁡〖 θ〗', 'Check complex math func content');
			})
		})

		QUnit.module( "Font", function ()
		{
			Test("\\doubleE ", [["ParaRun", "𝔼"]], false, "Check math font autocorrection");
			Test("\\frakturE ", [["ParaRun", "𝔈"]], false, "Check math font autocorrection");
			Test("\\scriptE ", [["ParaRun", "ℰ"]], false, "Check math font autocorrection");
		})
	})

	QUnit.module("Cursor", function ()
	{
		QUnit.test('Check cursor position after convert empty math func (cos, sin..)', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);

			AddText('cos ');

			let cont = MathContent.Root;
			let func = cont.Content[1];
			let arg = func.getArgument();

			assert.strictEqual(cont.CurPos, 1, 'Cursor inside function');
			assert.strictEqual(func.CurPos, 1, 'Cursor in func argument');
			assert.strictEqual(arg.CurPos, 0, 'Cursor selected first paraRun in func argument');
		})

		QUnit.test('Check cursor position after convert empty big nary', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);

			AddText('\\int  ');

			let cont = MathContent.Root;
			let func = cont.Content[1];
			let arg = func.getBase();

			assert.strictEqual(cont.CurPos, 1, 'Cursor inside nary');
			assert.strictEqual(func.CurPos, 2, 'Cursor in nary base');
			assert.strictEqual(arg.CurPos, 0, 'Cursor selected first paraRun in func argument');
		})

		QUnit.test('Check two spaces after nary (for example) not trigger auto-correction', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);

			AddText('∫ ');
			logicDocument.Document_Undo();
			AddText(' ');

			let cont = MathContent.Root
			assert.strictEqual(cont.Content.length, 1, 'check only one content');
			let run = cont.Content[0];
			assert.strictEqual(run instanceof ParaRun, true, 'check run');
		})
		QUnit.test('Check cursor position after convert math func with content (cos, sin..)', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);
			AddText('cos\\funcapply(1+2) ');

			let cont = MathContent.Root;

			assert.strictEqual(cont.CurPos, 2, 'Cursor after function');
		})

		QUnit.test('Add nary from menu', function (assert)
		{
			if (p1.Content.length > 0)
				p1.Content.splice(0, p1.Content.length);

			p1.AddToContentToEnd(AscTest.CreateRun())
			p1.AddToContentToEnd(AscTest.CreateRun())
			p1.CorrectContent(true);

			AscTest.MoveCursorToParagraph(p1, true);

			logicDocument.SetMathInputType(0);
			logicDocument.AddParaMath(67108864); // nary without content

			MathContent = logicDocument.GetCurrentMath();
			assert.ok(true, "Add nary without content ∫");

			let strNary = MathContent.GetTextOfElement(0).GetText();
			assert.strictEqual(strNary, '∫', 'Check');
		})

		QUnit.test('Add linear fraction and convert to linear to proff', function (assert)
		{
			if (p1.Content.length > 0)
				p1.Content.splice(0, p1.Content.length);

			p1.AddToContentToEnd(AscTest.CreateRun())
			p1.AddToContentToEnd(AscTest.CreateRun())
			p1.CorrectContent(true);

			AscTest.MoveCursorToParagraph(p1, true);

			logicDocument.SetMathInputType(0);
			logicDocument.AddParaMath(16777218);
			MathContent = logicDocument.GetCurrentMath();
			assert.ok(true, "Add linear fraction without content []/[]");
			AscTest.MoveCursorLeft(false, false, 1);
			AddText('y');
			AscTest.MoveCursorLeft(false, false, 3);
			AddText('x');
			assert.ok(true, "Add content to numerator and denominator");

			let cont = MathContent.Root
			let frac = cont.Content[1];
			let isLinear = frac.Pr.type === LINEAR_FRACTION;
			assert.strictEqual(isLinear, true, 'Check is linear fraction');

			logicDocument.ConvertMathView(true);
			assert.ok(true, "Convert in linear form");

			let strLinearFrac = MathContent.GetTextOfElement(0).GetText();
			assert.strictEqual(strLinearFrac, 'x∕y', 'Check');

			logicDocument.ConvertMathView(false);
			assert.ok(true, "Convert in professional form");

			isLinear = frac.Pr.type === LINEAR_FRACTION;
			assert.strictEqual(isLinear, true, 'Check is linear fraction');
		})

		QUnit.test('Check spaces degradation while convert math', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);
			AddText('1/2  1/2  1/2 ');
			assert.ok(true, "Add 1/2 1/2 1/2 fraction");
			logicDocument.ConvertMathView(true); assert.ok(true, "Convert to linear form");
			logicDocument.ConvertMathView(false); assert.ok(true, "Convert to professional form");
			logicDocument.ConvertMathView(true);assert.ok(true, "Convert to linear form");
			logicDocument.ConvertMathView(false); assert.ok(true, "Convert to professional form");
			logicDocument.ConvertMathView(true);assert.ok(true, "Convert to linear form");
			logicDocument.ConvertMathView(false); assert.ok(true, "Convert to professional form");

			let strNary = MathContent.GetTextOfElement(0).GetText();
			assert.strictEqual(strNary, '1/2 1/2 1/2', 'Check text');
		})

		QUnit.test('Function autocorrection with _ and ^', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);
			AddText('cos_');

			let strFunc = MathContent.GetTextOfElement(0).GetText();
			assert.strictEqual('⁡', strFunc[3], 'Check \\funcapply');
			assert.strictEqual(MathContent.Root.Content[0].MathPrp.sty === 3, true, 'Check normal style');
		})

		QUnit.test('Limit function autocorrection with _ and ^', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);
			AddText('lim_');

			let strFunc = MathContent.GetTextOfElement(0).GetText();
			assert.strictEqual('⁡', strFunc[3], 'Check \\funcapply');
		})

		QUnit.test('Check processing of fractions', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);
			AddText('1/2/3/4/5/6/7/8 ');
			assert.ok(true, "Add 1/2/3/4/5/6/7/8 fraction");
			let strFrac = MathContent.GetTextOfElement(0).GetText();
			assert.strictEqual(strFrac, '((((((1/2)/3)/4)/5)/6)/7)/8', 'Check');
		})

		QUnit.test('Check processing of fractions 2', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);
			AddText('1/2 //3 ');
			assert.ok(true, "Add 1/2 //3  fraction");
			let strFrac = MathContent.GetTextOfElement(0).GetText();
			assert.strictEqual(strFrac, '((1/2)/)/3', 'Check');
		})

		QUnit.test('Check space eating while auto-convert between-correction', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);
			AddText(' 1/2 ');
			assert.ok(true, "Add ' 1/2 ' fraction with space before");
			let strFrac = MathContent.GetTextOfElement(0).GetText();
			assert.strictEqual(strFrac, '1/2', 'Check is space ate');
		})

		QUnit.test('Check cursor pos after del content inside math func argument', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);

			AddText('cos ');

			let cont = MathContent.Root;
			let func = cont.Content[1];

			AddText('2_x ');

			AscTest.PressKey(8);
			AscTest.PressKey(8);

			assert.strictEqual(func.CurPos, 1, 'Cursor inside func argument');
		})

		QUnit.test('Check degree pos after convert inside math content', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);
			AscMath.SetAutoConvertation(false);
			AddText('1/ ');
			MathContent.ConvertView(false, Asc.c_oAscMathInputType.Unicode);
			AscMath.SetAutoConvertation(true);

			let cont = MathContent.Root;
			let frac = cont.Content[1];
			let den = frac.getDenominator();

			den.SelectThisElement(1);
			den.SelectAll(1);

			AddText('2_x ');

			assert.strictEqual(den.CurPos, 2, 'Cursor after degree');
		})

		QUnit.test('Undo for empty math content placeholder', function (assert)
		{
			Clear();
			logicDocument.SetMathInputType(0);

			let delimiter			= MathContent.Root.Add_DelimiterEx(new CTextPr(), 1, [null],  null,  null);
			let oDelimiterContent	= delimiter.getElementMathContent(0);
			assert.ok(true, "Add empty bracket block ()");

			MathContent.Root.Correct_Content(true);
			oDelimiterContent.SelectThisElement(1);
			oDelimiterContent.SelectAll(1);
			assert.ok(true, "Select empty placeholder inside bracket block");

			AscTest.EnterText("1");
			assert.ok(true, "Enter '1'");

			logicDocument.Document_Undo();
			assert.ok(true, "Undo");
			let r = oDelimiterContent.getElem(0);

			let strR = r.GetTextOfElement().GetText();
			assert.strictEqual(strR, '', 'Inside bracket block empty string: \"'+ strR + '\"');
			assert.strictEqual(r.IsPlaceholder(), true, 'only placeholder');
		})
	})

	QUnit.module( "LaTeX", function ()
	{
		Test("\\theta", [["ParaRun", "\\theta"]], true, "Check LaTeX words");
		Test("\\theta", [["ParaRun", "θ"]], true, "Check LaTeX words", undefined, true);
		Test("\\pm", [["ParaRun", "\\pm"]], true, "Check LaTeX words");
		Test("\\pm", [["ParaRun", "±"]], true, "Check LaTeX words", undefined, true);
		Test("\\infty", [["ParaRun", "∞"]], true, "Check LaTeX words", undefined, true);
		Test("\\infty", [["ParaRun", "\\infty"]], true, "Check LaTeX words");
		Test("\\sum_{\\begin{matrix}0\\lei\\lem\\\\0<j<n\\\\\\end{matrix}}{P\\left(i,j\\right)}", [["ParaRun", ""], ["CNary", "\\sum_{\\begin{matrix}0\\lei\\lem\\\\0<j<n\\\\\\end{matrix}}{P\\left(i,j\\right)}"]], true, "Check LaTeX words");
		Test("1\\ 2", [["ParaRun", "1\\ 2"]], true, "Check LaTeX words");
		Test("\\dot{}\\lim\\below{n\\rightarrow\\infty}{\\left(1+\\frac{1}{n}\\right)^n}", [["ParaRun", ""], ["CAccent", "\\dot{}"], ["ParaRun", ""], ["CMathFunc", "\\lim\\below{n\\to\\infty}{\\left(1+\\frac{1}{n}\\right)^n}"],], true, "Check LaTeX words");

		QUnit.module( "accent", function ()
		{
			Test("\\dot{}", [["ParaRun", ""], ["CAccent", "\\dot{}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\dot{a}", [["ParaRun", ""], ["CAccent", "\\dot{a}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\ddot{b}", [["ParaRun", ""], ["CAccent", "\\ddot{b}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\acute{c}", [["ParaRun", ""], ["CAccent", "\\acute{c}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\grave{d}", [["ParaRun", ""], ["CAccent", "\\grave{d}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\check{e}", [["ParaRun", ""], ["CAccent", "\\check{e}"],["ParaRun", ""]], true, "Check LaTeX words");

			Test("\\breve{f}", [["ParaRun", ""], ["CAccent", "\\breve{f}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\tilde{g}", [["ParaRun", ""], ["CAccent", "\\tilde{g}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\bar{h}", [["ParaRun", ""], ["CAccent", "\\bar{h}"],["ParaRun", ""]], true, "Check LaTeX words");

			Test("\\widehat{j}", [["ParaRun", ""], ["CAccent", "\\hat{j}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\vec{k}", [["ParaRun", ""], ["CAccent", "\\vec{k}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\vec{\\frac{k}{2}}", [["ParaRun", ""], ["CAccent", "\\vec{\\frac{k}{2}}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\overbrace{cpppppppp}", [["ParaRun", ""], ["CGroupCharacter", "\\overbrace{cpppppppp}"],["ParaRun", ""]], true, "Check LaTeX words");
			// Test("5''", [["ParaRun", ""], ["ParaRun", ""],["CAccent", ""]], true, "Check LaTeX words");
			// Test("\\frac{4}{5}''", [["ParaRun", ""], ["CAccent", ""],["ParaRun", ""]], true, "Check LaTeX words");
		})

		QUnit.module( "bar", function ()
		{
			Test("\\overbrace{cpppppppp}", [["ParaRun", ""], ["CGroupCharacter", "\\overbrace{cpppppppp}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\underline{1+2}", [["ParaRun", ""], ["CBar", "\\underline{1+2}"],["ParaRun", ""]], true, "Check LaTeX words");
		})

		QUnit.module( "horizontal arrows", function ()
		{
			Test("{\\gets\\below{xxxx}}", [["ParaRun", ""], ["CGroupCharacter", "{\\gets\\below{xxxx}}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("{\\rightarrow\\below{ssss}}", [["ParaRun", ""], ["CGroupCharacter", "{\\rightarrow\\below{ssss}}"],["ParaRun", ""]], true, "Check LaTeX words");

		})
		QUnit.module( "bar and limit", function ()
		{
			Test("\\overbrace{cpppppppp}_{2+1}", [["ParaRun", ""], ["CLimit", "\\overbrace{cpppppppp}\\below{2+1}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\underline{1+2}\\above{2+1}", [["ParaRun", ""], ["CLimit", "\\underline{1+2}\\above{2+1}"],["ParaRun", ""]], true, "Check LaTeX words");
		})

		QUnit.module( "brackets", function ()
		{
			Test("(2+1)", [["ParaRun", ""], ["CDelimiter", "\\left(2+1\\right)"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("()", [["ParaRun", ""], ["CDelimiter", "\\left(\\right)"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("(", [["ParaRun", "("]], true, "Check LaTeX words");
			Test(")", [["ParaRun", ")"]], true, "Check LaTeX words");

			Test("[2+1]", [["ParaRun", ""], ["CDelimiter", "\\left[2+1\\right]"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("[]", [["ParaRun", ""], ["CDelimiter", "\\left[\\right]"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("[", [["ParaRun", "["]], true, "Check LaTeX words");
			Test("]", [["ParaRun", "]"]], true, "Check LaTeX words");

			Test("\\{2+1\\}", [["ParaRun", ""], ["CDelimiter", "\\left\\{2+1\\right\\}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\{\\}", [["ParaRun", ""], ["CDelimiter", "\\left\\{\\right\\}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\{", [["ParaRun", "\\{"]], true, "Check LaTeX words");
			Test("\\}", [["ParaRun", "\\}"]], true, "Check LaTeX words");

			Test("|2+1|", [["ParaRun", ""], ["CDelimiter", "\\left|2+1\\right|"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("||", [["ParaRun", ""], ["CDelimiter", "\\left|\\right|"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("|", [["ParaRun", "|"]], true, "Check LaTeX words");

			Test("\\left]2+1\\right[", [["ParaRun", ""], ["CDelimiter", "\\left]2+1\\right["],["ParaRun", ""]], true, "Check LaTeX words");
			Test("(2+1]", [["ParaRun", ""], ["CDelimiter", "\\left(2+1\\right]"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\left.1+2\\right)", [["ParaRun", ""], ["CDelimiter", "\\left.1+2\\right)"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("|2|+\\{1\\}+|2|", [["ParaRun", ""], ["CDelimiter", "\\left|2\\right|"],["ParaRun", "+"], ["CDelimiter", "\\left\\{1\\right\\}"], ["ParaRun", "+"], ["CDelimiter", "\\left|2\\right|"]], true, "Check LaTeX words");
			Test("\\left\\vert s \\right\\vert", [["ParaRun", ""], ["CDelimiter", "\\left|s\\right|"],["ParaRun", ""]], true, "Check LaTeX words");
		})

		QUnit.module( "degree", function ()
		{
			Test("2^2", [["ParaRun", ""], ["CDegree", "2^2"], ["ParaRun", ""]], true, "Check LaTeX words");
			Test("a^b", [["ParaRun", ""], ["CDegree", "a^b"], ["ParaRun", ""]], true, "Check LaTeX words");
			Test("a^2", [["ParaRun", ""], ["CDegree", "a^2"], ["ParaRun", ""]], true, "Check LaTeX words");
			Test("2^b", [["ParaRun", ""], ["CDegree", "2^b"], ["ParaRun", ""]], true, "Check LaTeX words");
			Test("2_2", [["ParaRun", ""], ["CDegree", "2_2"], ["ParaRun", ""]], true, "Check LaTeX words");
			Test("a_b", [["ParaRun", ""], ["CDegree", "a_b"], ["ParaRun", ""]], true, "Check LaTeX words");
			Test("a_2", [["ParaRun", ""], ["CDegree", "a_2"], ["ParaRun", ""]], true, "Check LaTeX words");
			Test("2_b", [["ParaRun", ""], ["CDegree", "2_b"], ["ParaRun", ""]], true, "Check LaTeX words");
			Test(`k_{n+1} = n^2 + k_n^2 - k_{n-1}`, [["ParaRun", ""], ["CDegree", "k_{n+1}"], ["ParaRun", "="], ["CDegree", "n^2"], ["ParaRun", "+"], ["CDegreeSubSup", "k_n^2"], ["ParaRun", "-"], ["CDegree", "k_{n-1}"] ], true, "Check LaTeX words");
			Test("k^{n+1} ", [["ParaRun", ""], ["CDegree", "k^{n+1}"], ["ParaRun", ""]], true, "Check LaTeX degree");
			Test("n^2 ",[["ParaRun", ""], ["CDegree", "n^2"], ["ParaRun", ""]], true, "Check LaTeX degree");
			Test("n^{2} ", [["ParaRun", ""], ["CDegree", "n^2"], ["ParaRun", ""]], true, "Check LaTeX degree");
			Test("n^(2) ", [["ParaRun", ""], ["CDegree", "n^{\\left(2\\right)}"], ["ParaRun", ""]], true, "Check LaTeX degree");
			Test("n^{2+1}_y", [["ParaRun", ""], ["CDegreeSubSup", "n_y^{2+1}"], ["ParaRun", ""]], true, "Check LaTeX degree");
		})

		QUnit.module( "prescript", function ()
		{
			Test("{_1^n}Y", [["ParaRun", ""], ["CDegreeSubSup", "{_1^n}Y"], ["ParaRun", ""]], true, "Check LaTeX prescript");
		})

		QUnit.module( "frac", function ()
		{
			Test(`\\frac{1}{2}^{2}`, [["ParaRun", ""], ["CDegree", "{\\frac{1}{2}}^2"],["ParaRun", ""]], true, "Check LaTeX words");
			Test(`\\frac{1}{2}_2`, [["ParaRun", ""], ["CDegree", "{\\frac{1}{2}}_2"], ["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\frac{1}{2}_2^y", [["ParaRun", ""], ["CDegreeSubSup", "{\\frac{1}{2}}_2^y"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\frac{1}{2}_{2}^{y}", [["ParaRun", ""], ["CDegreeSubSup", "{\\frac{1}{2}}_2^y"],["ParaRun", ""]], true, "Check LaTeX words");
			// Test("\\frac{1}{2}_1_2_3_4_5_6_7", [["ParaRun", ""], ["CDegree", ""],["ParaRun", ""]], true, "Check LaTeX words");
			// Test("\\frac{1}{2}^1^2^3^4^5^6^7", [["ParaRun", ""], ["ParaRun", ""],["ParaRun", ""]], true, "Check LaTeX words");
			// Test("\\frac{1}{2}^1^2^3^4^5^6^7_x", [["ParaRun", ""], ["ParaRun", ""],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\frac{1}{2}", [["ParaRun", ""], ["CFraction", "\\frac{1}{2}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\frac{1+\\frac{x}{y}}{2}", [["ParaRun", ""], ["CFraction", "\\frac{1+\\frac{x}{y}}{2}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\frac{1^x}{2_y}", [["ParaRun", ""], ["CFraction", "\\frac{1^x}{2_y}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\binom{1}{2}", [["ParaRun", ""], ["CDelimiter", "\\binom{1}{2}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\sfrac{1}{2}", [["ParaRun", ""], ["CFraction", "\\sfrac{1}{2}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("a\\atop2", [["ParaRun", ""], ["CFraction", "{a\\atop2}"],["ParaRun", ""]], true, "Check LaTeX words");
		})

		QUnit.module( "nary", function ()
		{
			Test("\\sum^{2}_{x}4", [["ParaRun", ""], ["CNary", "\\sum_{x}^{2}4"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\int^2_x{4}", [["ParaRun", ""], ["CNary", "\\int_{x}^{2}4"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\sum_{i=1}^{10} t_i", [["ParaRun", ""], ["CNary", "\\sum_{i=1}^{10}{t_i}"],["ParaRun", ""]], true, "Check LaTeX words");
		})

		QUnit.module( "func", function ()
		{
			Test("\\exp_a b", [["ParaRun", ""], ["CMathFunc", "\\exp_a{b}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\exp b = e^b, 10^m", [["ParaRun", ""], ["CMathFunc", "\\exp{b}"], ["ParaRun", "="], ["CDegree", "e^b"], ["ParaRun", ","], ["CDegree", "10^m"]], true, "Check LaTeX words");
			Test("\\exp_{a}^x {b}", [["ParaRun", ""], ["CMathFunc", "\\exp_a^x{b}"],["ParaRun", ""]], true, "Check LaTeX words");

			Test("\\ln c", [["ParaRun", ""], ["CMathFunc", "\\ln{c}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\log e", [["ParaRun", ""], ["CMathFunc", "\\log{e}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\log_{10} f", [["ParaRun", ""], ["CMathFunc", "\\log_{10}{f}"],["ParaRun", ""]], true, "Check LaTeX words");

			Test("\\sin a", [["ParaRun", ""], ["CMathFunc", "\\sin{a}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\cos b", [["ParaRun", ""], ["CMathFunc", "\\cos{b}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\tan c", [["ParaRun", ""], ["CMathFunc", "\\tan{c}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\cot d", [["ParaRun", ""], ["CMathFunc", "\\cot{d}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\sec e", [["ParaRun", ""], ["CMathFunc", "\\sec{e}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\csc f", [["ParaRun", ""], ["CMathFunc", "\\csc{f}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\cos^2_{y}{b}", [["ParaRun", ""], ["CMathFunc", "\\cos_y^2{b}"],["ParaRun", ""]], true, "Check LaTeX words");

			Test("\\arcsin h", [["ParaRun", ""], ["CMathFunc", "\\arcsin{h}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\arccos_x i", [["ParaRun", ""], ["CMathFunc", "\\arccos_x{i}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\arctan^y_{x} {j}", [["ParaRun", ""], ["CMathFunc", "\\arctan_x^y{j}"],["ParaRun", ""]], true, "Check LaTeX words");

			//Test("\\sinhk", [["ParaRun", ""], ["CMathFunc", ""],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\cosh {l}", [["ParaRun", ""], ["CMathFunc", "\\cosh{l}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\tanh_x^y m", [["ParaRun", ""], ["CMathFunc", "\\tanh_x^y{m}"],["ParaRun", ""]], true, "Check LaTeX words");
			//Test("\\coth^{x}_y_1_2 {n}", [["ParaRun", ""], ["CDegree", ""],["ParaRun", ""]], true, "Check LaTeX words");

			Test("\\min(x,y)", [["ParaRun", ""], ["CMathFunc", "\\min{\\left(x,y\\right)}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\cos(2\\theta)", [["ParaRun", ""], ["CMathFunc", "\\cos{\\left(2\\theta\\right)}"], ["ParaRun", ""]], true, "Check LaTeX function");
			//Test("\\lim_{x\\to \\infty }\\exp(x) ", [["ParaRun", ""], ["CMathFunc", "\\lim_{x→∞} { \\exp { (x)}}"], ["ParaRun", ""]], true, "Check LaTeX function");
		})

		QUnit.module( "radical", function ()
		{
			Test("\\sqrt5", [["ParaRun", ""], ["CRadical", "\\sqrt5"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\sqrt\\frac{1}{2}", [["ParaRun", ""], ["CRadical", "\\sqrt{\\frac{1}{2}}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\sqrt[2^2]\\frac{1}{2}", [["ParaRun", ""], ["CRadical", "\\sqrt[2^2]{\\frac{1}{2}}"],["ParaRun", ""]], true, "Check LaTeX words");
			Test("\\sqrt[2^2] {\\frac{1}{2}+3}", [["ParaRun", ""], ["CRadical", "\\sqrt[2^2]{\\frac{1}{2}+3}"],["ParaRun", ""]], true, "Check LaTeX words");

			QUnit.test('Check pos for radical in LaTeX', function (assert)
			{
				Clear();
				logicDocument.SetMathInputType(1);

				AddText('\\pm\\sqrt');

				MathContent.ConvertView(true, Asc.c_oAscMathInputType.LaTeX);
				assert.ok(true, "Convert to proff. view");

				MathContent.ConvertView(true, Asc.c_oAscMathInputType.LaTeX);
				assert.ok(true, "Convert to linear view");

				let strFunc = MathContent.GetTextOfElement(0).GetText();
				assert.strictEqual(strFunc, '\\pm\\sqrt', 'Check complex math func content');
			})
		})

		QUnit.module( "bugs", function ()
		{
			QUnit.test('Check linear form for nonstandard name func', function (assert)
			{
				Clear();
				logicDocument.SetMathInputType(1);

				AddText('\\lim\\below{\\left(n\\to\\infty\\right){\\left(1+\\frac{1}{n}\\right)^n}}');

				MathContent.ConvertView(true, Asc.c_oAscMathInputType.LaTeX);
				assert.ok(true, "Convert to proff. view");

				MathContent.ConvertView(true, Asc.c_oAscMathInputType.LaTeX);
				assert.ok(true, "Convert to linear view");

				let strFunc = MathContent.GetTextOfElement(0).GetText();
				assert.strictEqual(strFunc, '\\lim\\below{\\left(n\\to\\infty\\right){\\left(1+\\frac{1}{n}\\right)^n}}', 'Check complex math func content');
			})

			// QUnit.todo('Check eqarray frac - Find case for LaTeX', function (assert)
			// {
			// 	// Find case for LaTeX
			// 	Clear();
			// 	logicDocument.SetMathInputType(1);
			// 	AddText('\\frac{\\substack{1\\\\\\substack{\\\\\\substack{\\\\\\substack{\\\\}}}}}{2}');
			// 	assert.ok(true, "Add text '\\frac{\\substack{1\\\\\\substack{\\\\\\substack{\\\\\\substack{\\\\}}}}}{2}'");
			//
			// 	MathContent.ConvertView(false, Asc.c_oAscMathInputType.LaTeX);
			// 	assert.ok(true, "Convert to linear view");
			//
			// 	let strBinomial = MathContent.GetTextOfElement(true).GetText();
			// 	assert.strictEqual(strBinomial, '\\frac{\\substack{1\\\\\\substack{\\\\\\substack{\\\\\\substack{\\\\}}}}}{2}', 'Check');
			// })

			QUnit.module( "Check LaTeX matrix", function ()
			{
				Test("\\begin{matrix}1&2\\\\3&4\\\\\\end{matrix}", [["ParaRun", ""], ["CMathMatrix", "\\begin{matrix}1&2\\\\3&4\\\\\\end{matrix}"]], true, "Check bug #61007 default matrix");
				Test("\\begin{pmatrix}1&2\\\\3&4\\\\\\end{pmatrix}", [["ParaRun", ""], ["CDelimiter", "\\left(\\begin{matrix}1&2\\\\3&4\\\\\\end{matrix}\\right)"]], true, "Check bug #61007 pmatrix");
				Test("\\left[\\begin{matrix}1&2\\\\3&4\\\\\\end{matrix}\\right]", [["ParaRun", ""], ["CDelimiter", "\\left[\\begin{matrix}1&2\\\\3&4\\\\\\end{matrix}\\right]"]], true, "Check bug #61007 pmatrix");
				Test("\\begin{matrix}&&\\\\&&\\\\&&\\\\&&\\end{matrix}", [["ParaRun", ""], ["CMathMatrix", "\\begin{matrix}&&\\\\&&\\\\&&\\\\&&\\\\\\end{matrix}"]], true, "Check matrix bug #71892");
				Test("\\begin{array}{l} n + 1\\end{array}", [["ParaRun", ""], ["CEqArray", "\\matrix{n+1}"]], true, "Check matrix bug #71892");
			})

			QUnit.module( "Check bug #67181", function ()
			{
				Test("\\mathcal{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "𝓆𝓌ℯ𝓇𝓉𝓎𝓊𝒾ℴ𝓅𝒶𝓈𝒹𝒻ℊ𝒽𝒿𝓀𝓁𝓏𝓍𝒸𝓋𝒷𝓃𝓂"]], true, "Check bug #67181", true, true);
				Test("\\mathcal{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "\\mathcal{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");

				Test("\\mathsf{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "𝗊𝗐𝖾𝗋𝗍𝗒𝗎𝗂𝗈𝗉𝖺𝗌𝖽𝖿𝗀𝗁𝗃𝗄𝗅𝗓𝗑𝖼𝗏𝖻𝗇𝗆"]], true, "Check bug #67181", true, true);
				Test("\\mathsf{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "\\mathsf{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");

				Test("\\mathrm{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "qwertyuiopasdfghjklzxcvbnm"]], true, "Check bug #67181", true, true);
				Test("\\mathrm{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "qwertyuiopasdfghjklzxcvbnm"]], true, "Check bug #67181");

				Test("\\mathit{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "𝑞𝑤𝑒𝑟𝑡𝑦𝑢𝑖𝑜𝑝𝑎𝑠𝑑𝑓𝑔ℎ𝑗𝑘𝑙𝑧𝑥𝑐𝑣𝑏𝑛𝑚"]], true, "Check bug #67181", true, true);
				Test("\\mathit{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "𝑞𝑤𝑒𝑟𝑡𝑦𝑢𝑖𝑜𝑝𝑎𝑠𝑑𝑓𝑔ℎ𝑗𝑘𝑙𝑧𝑥𝑐𝑣𝑏𝑛𝑚"]], true, "Check bug #67181"); // in word not convert

				Test("\\mathfrak{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "𝔮𝔴𝔢𝔯𝔱𝔶𝔲𝔦𝔬𝔭𝔞𝔰𝔡𝔣𝔤𝔥𝔧𝔨𝔩𝔷𝔵𝔠𝔳𝔟𝔫𝔪"]], true, "Check bug #67181", true, true);
				Test("\\mathfrak{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "\\mathfrak{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");

				Test("\\mathbf{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "𝐪𝐰𝐞𝐫𝐭𝐲𝐮𝐢𝐨𝐩𝐚𝐬𝐝𝐟𝐠𝐡𝐣𝐤𝐥𝐳𝐱𝐜𝐯𝐛𝐧𝐦"]], true, "Check bug #67181", true, true);
				Test("\\mathbf{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "\\mathbf{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");

				Test("\\mathbb{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "𝕢𝕨𝕖𝕣𝕥𝕪𝕦𝕚𝕠𝕡𝕒𝕤𝕕𝕗𝕘𝕙𝕛𝕜𝕝𝕫𝕩𝕔𝕧𝕓𝕟𝕞"]], true, "Check bug #67181", true, true);
				Test("\\mathbb{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "\\mathbb{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");

				Test("\\mathfrak{qwerty}\\mathfrak{uiopasdfghjklzxcvbnm}", [["ParaRun", "𝔮𝔴𝔢𝔯𝔱𝔶𝔲𝔦𝔬𝔭𝔞𝔰𝔡𝔣𝔤𝔥𝔧𝔨𝔩𝔷𝔵𝔠𝔳𝔟𝔫𝔪"]], true, "Check bug #67181", true, true);
				Test("\\mathfrak{qwerty}\\mathfrak{uiopasdfghjklzxcvbnm}", [["ParaRun", "\\mathfrak{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181");

				Test("\\sf{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "𝗊𝗐𝖾𝗋𝗍𝗒𝗎𝗂𝗈𝗉𝖺𝗌𝖽𝖿𝗀𝗁𝗃𝗄𝗅𝗓𝗑𝖼𝗏𝖻𝗇𝗆"]], true, "Check bug #67181 check non-standard", true, true);
				Test("\\sf{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "\\mathsf{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181 check non-standard");

				Test("\\double{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "𝕢𝕨𝕖𝕣𝕥𝕪𝕦𝕚𝕠𝕡𝕒𝕤𝕕𝕗𝕘𝕙𝕛𝕜𝕝𝕫𝕩𝕔𝕧𝕓𝕟𝕞"]], true, "Check bug #67181 check non-standard", true, true);
				Test("\\double{qwertyuiopasdfghjklzxcvbnm}", [["ParaRun", "\\mathbb{qwertyuiopasdfghjklzxcvbnm}"]], true, "Check bug #67181 check non-standard");
			})
		})
	})
 })


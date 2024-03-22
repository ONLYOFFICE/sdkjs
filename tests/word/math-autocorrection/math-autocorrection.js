/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
    }

	Init();

    function Clear() {
        Root.Remove_FromContent(0, Root.Content.length);
        Root.Correct_Content();
    }

    function AddText(str) {
		let one = str.getUnicodeIterator();

		while (one.isInside()) {
			let oElement = new AscWord.CRunText(one.value());
			MathContent.Add(oElement);
			one.next();
		}
    }

    function Test(str, intCurPos, arrResult, isLaTeX, strNameOfTest)
	{
		let nameOfTest = strNameOfTest ? strNameOfTest + " \'" + str + "\'" : str;
        QUnit.test(nameOfTest, function (assert)
		{
			if (isLaTeX) {
				logicDocument.SetMathInputType(1);
			} else {
				logicDocument.SetMathInputType(0);
			}
            function AutoTest(isLaTeX, str, intCurPos, arrResultContent)
			{

				let CurPos = Root.CurPos;
                AddText(str);

                for (let i = CurPos; i < Root.Content.length; i++)
				{
                    let CurrentContent = Root.Content[i];
                    let CheckContent = arrResultContent[i];

                    assert.strictEqual(
                        CurrentContent.constructor.name,
                        CheckContent[0],
                        "Content[" + i + "] === " +
                        Root.Content[i].constructor.name
                    );

                    let TextContent = CurrentContent.GetTextOfElement(logicDocument.MathInputType);
                    assert.strictEqual(TextContent, CheckContent[1], "Text of Content[" + i + "]: '" + CheckContent[1] + "'");

                    if (CurrentContent.constructor.name === "ParaRun" && i === intCurPos) {
                        assert.strictEqual(CurrentContent.IsCursorAtEnd(), true, "Cursor at the end of ParaRun");
                    }

                }

                assert.strictEqual(Root.CurPos, intCurPos, "Check cursor position: " + intCurPos);
            }

			Clear()
            AutoTest(isLaTeX, str, intCurPos, arrResult);
        })
    }

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
	}

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

	Test("()+", 2, [["ParaRun", ""], ["CDelimiter", "()"], ["ParaRun", "+"], ["ParaRun", ""]], false);
	Test("{}+", 2, [["ParaRun", ""], ["CDelimiter", "{}"], ["ParaRun", "+"], ["ParaRun", ""]], false);
	Test("[]+", 2, [["ParaRun", ""], ["CDelimiter", "[]"], ["ParaRun", "+"], ["ParaRun", ""]], false);
	Test("||+", 2, [["ParaRun", ""], ["CDelimiter", "||"], ["ParaRun", "+"], ["ParaRun", ""]], false);

	Test("(1+2)+", 3, [["ParaRun", ""], ["CDelimiter", "(1+2)"], ["ParaRun", "+"], ["ParaRun", ""]], false);
	Test("{1+2}+", 3, [["ParaRun", ""], ["CDelimiter", "{1+2}"], ["ParaRun", "+"], ["ParaRun", ""]], false);
	Test("[1+2]+", 3, [["ParaRun", ""], ["CDelimiter", "[1+2]"], ["ParaRun", "+"], ["ParaRun", ""]], false);
	Test("|1+2|+", 3, [["ParaRun", ""], ["CDelimiter", "|1+2|"], ["ParaRun", "+"], ["ParaRun", ""]], false);

	Test("1/2 )", 2, [["ParaRun", ""], ["CFraction", "〖1/2〗"], ["ParaRun", ")"]], false);
	Test("1/2 }", 2, [["ParaRun", ""], ["CFraction", "〖1/2〗"], ["ParaRun", "}"]], false);
	Test("1/2 ]", 2, [["ParaRun", ""], ["CFraction", "〖1/2〗"], ["ParaRun", "]"]], false);
	Test("1/2 |", 2, [["ParaRun", ""], ["CFraction", "〖1/2〗"], ["ParaRun", "|"]], false);

	Test("(1/2 ", 2, [["ParaRun", "("], ["CFraction", "〖1/2〗"], ["ParaRun", ""]], false);
	Test("{1/2 ", 2, [["ParaRun", "{"], ["CFraction", "〖1/2〗"], ["ParaRun", ""]], false);
	Test("[1/2 ", 2, [["ParaRun", "["], ["CFraction", "〖1/2〗"], ["ParaRun", ""]], false);
	Test("|1/2 ", 2, [["ParaRun", "|"], ["CFraction", "〖1/2〗"], ["ParaRun", ""]], false);

	Test("(1/2 )+", 2, [["ParaRun", ""], ["CDelimiter", "(〖1/2〗)"], ["ParaRun", "+"], ["ParaRun", ""]], false);
	Test("{1/2 }+", 2, [["ParaRun", ""], ["CDelimiter", "{〖1/2〗}"], ["ParaRun", "+"], ["ParaRun", ""]], false);
	Test("[1/2 ]+", 2, [["ParaRun", ""], ["CDelimiter", "[〖1/2〗]"], ["ParaRun", "+"], ["ParaRun", ""]], false);
	Test("|1/2 |+", 2, [["ParaRun", ""], ["CDelimiter", "|〖1/2〗|"], ["ParaRun", "+"], ["ParaRun", ""]], false);

	Test("(1/2)", 2, [["ParaRun", "("], ["CFraction", "〖1/2〗"], ["ParaRun", ")"]], false);

	Test("2_1", 0, [["ParaRun", "2_1"]], false);
	Test("2_1 ", 2, [["ParaRun", ""], ["CDegree", "2_(1)"], ["ParaRun", ""]], false);
    Test("\\int", 0, [["ParaRun", "\\int"]], false);
    Test("\\int _x^y\\of 1/2  ", 2, [["ParaRun", ""], ["CDelimiter", "〖∫^y_x▒〖〖1/2〗〗〗"], ["ParaRun", ""]], false);
    Test("1/2 ", 2, [["ParaRun", ""], ["CFraction", "〖1/2〗"], ["ParaRun", ""]], false);
    Test("1/2 +", 2, [["ParaRun", ""], ["CFraction", "〖1/2〗"], ["ParaRun", "+"]], false);
    Test("1/2=", 2, [["ParaRun", ""], ["CFraction", "〖1/2〗"], ["ParaRun", "="]], false);
	Test("1/2+1/2=x/y ", 6, [["ParaRun", ""], ["CFraction", "〖1/2〗"], ["ParaRun", "+"], ["CFraction", "〖1/2〗"], ["ParaRun", "="], ["CFraction", "〖x/y〗"], ["ParaRun", ""]], false);
	//
	// MultiLineTest(
	// 	["1/2", " "],
	// 	[0, 2],
	// 	[
	// 		[
	// 			["ParaRun", "1/2"]
	// 		],
	// 		[
	// 			["ParaRun", ""],
	// 			["CFraction", "〖1/2〗"],
	// 			["ParaRun", ""]
	// 		],
	// 	],
	// 	[]
	// 	);
	//
	// MultiLineTest(
	// 	["1/2 ", "+", "x/y", " "],
	// 	[2, 2, 2, 4],
	// 	[
	// 		[
	// 			["ParaRun", ""],
	// 			["CFraction", "〖1/2〗"],
	// 			["ParaRun", ""]
	// 		],
	// 		[
	// 			["ParaRun", ""],
	// 			["CDelimiter", "〖1/2〗"],
	// 			["ParaRun", "+"]
	// 		],
	// 		[
	// 			["ParaRun", ""],
	// 			["CDelimiter", "〖1/2〗"],
	// 			["ParaRun", "+x/y"]
	// 		],
	// 		[
	// 			["ParaRun", ""],
	// 			["CDelimiter", "〖1/2〗"],
	// 			["ParaRun", "+"],
	// 			["CFraction", "〖x/y〗"],
	// 			["ParaRun", ""],
	// 		],
	// 	],
	// 	[]
	// );
	//
	// Test("1/2 ", 2, [["ParaRun", ""], ["CFraction", "〖1/2〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("1/3.1416 ", 2, [["ParaRun", ""], ["CFraction", "〖1/(3.1416)〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("x/y ", 2, [["ParaRun", ""], ["CFraction", "〖x/y〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("x/2 ", 2, [["ParaRun", ""], ["CFraction", "〖x/2〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("x/(1+2) ", 2, [["ParaRun", ""], ["CFraction", "〖x/(1+2)〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("x/((1+2)) ", 2, [["ParaRun", ""], ["CFraction", "〖x/(1+2)〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("x/[1+2]  ", 2, [["ParaRun", ""], ["CDelimiter", "〖x/([1+2])〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("x/{1+2} ", 2, [["ParaRun", ""], ["CFraction", "〖x/({1+2})〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("x/[1+2} ", 2, [["ParaRun", ""], ["CFraction", "〖x/([1+2})〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("(1_i)/32 ", 2, [["ParaRun", ""], ["CFraction", "〖(1_(i))/(32)〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("(1_i)/32 ", 2, [["ParaRun", ""], ["CFraction", "〖(1_(i))/(32)〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("\\sdiv ", 0, [["ParaRun", "⁄"]], false, "Check fraction symbol");
	// Test("1\\sdiv 2 ", 2, [["ParaRun", ""], ["CFraction", "〖1∕2〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("x\\sdiv y ", 2, [["ParaRun", ""], ["CFraction", "〖x∕y〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("x\\sdiv (y+1_i) ", 2, [["ParaRun", ""], ["CFraction", "〖x∕(y+1_(i))〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("\\ndiv ", 0, [["ParaRun", "⊘"]], false, "Check fraction symbol");
	// Test("1\\ndiv 2 ", 2, [["ParaRun", ""], ["CFraction", "〖1⊘2〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("x\\ndiv y ", 2, [["ParaRun", ""], ["CFraction", "〖x⊘y〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("x\\ndiv (y+1_i) ", 2, [["ParaRun", ""], ["CFraction", "〖x⊘(y+1_(i))〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("\\atop ", 0, [["ParaRun", "¦"]], false, "Check fraction symbol");
	// Test("1\\atop 2 ", 2, [["ParaRun", ""], ["CFraction", "〖1¦2〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("x\\atop y ", 2, [["ParaRun", ""], ["CFraction", "〖x¦y〗"], ["ParaRun", ""]], false, "Check fraction");
	// Test("x\\atop (y+1_i) ", 2, [["ParaRun", ""], ["CFraction", "〖x¦(y+1_(i))〗"], ["ParaRun", ""]], false, "Check fraction");
	//
	Test("x_y ", 2, [["ParaRun", ""], ["CDegree", "x_(y)"], ["ParaRun", ""]], false, "Check degree");
	Test("_ ", 1, [["ParaRun", "_"], ["ParaRun", " "]], false, "Check degree");
	Test("x_1 ", 2, [["ParaRun", ""], ["CDegree", "x_(1)"], ["ParaRun", ""]], false, "Check degree");
	Test("1_x ", 2, [["ParaRun", ""], ["CDegree", "1_(x)"], ["ParaRun", ""]], false, "Check degree");
	Test("x_(1+2) ", 2, [["ParaRun", ""], ["CDegree", "x_(1+2)"], ["ParaRun", ""]], false, "Check degree");
	Test("x_[1+2] ", 2, [["ParaRun", ""], ["CDegree", "x_([1+2])"], ["ParaRun", ""]], false, "Check degree");
	Test("x_[1+2} ", 2, [["ParaRun", ""], ["CDegree", "x_([1+2})"], ["ParaRun", ""]], false, "Check degree");
	Test("x_1/2 ", 2, [["ParaRun", ""], ["CFraction", "〖(x_(1))/2〗"], ["ParaRun", ""]], false, "Check degree");

	Test("^ ", 1, [["ParaRun", "^"], ["ParaRun", " "]], false, "Check index");
	Test("x^y ", 2, [["ParaRun", ""], ["CDegree", "x^(y)"], ["ParaRun", ""]], false, "Check index");
	Test("x^1 ", 2, [["ParaRun", ""], ["CDegree", "x^(1)"], ["ParaRun", ""]], false, "Check index");
	Test("1^x ", 2, [["ParaRun", ""], ["CDegree", "1^(x)"], ["ParaRun", ""]], false, "Check index");
	Test("x^(1+2) ", 2, [["ParaRun", ""], ["CDegree", "x^(1+2)"], ["ParaRun", ""]], false, "Check index");
	Test("x^[1+2] ", 2, [["ParaRun", ""], ["CDegree", "x^([1+2])"], ["ParaRun", ""]], false, "Check index");
	Test("x^[1+2} ", 2, [["ParaRun", ""], ["CDegree", "x^([1+2})"], ["ParaRun", ""]], false, "Check index");
	Test("x^1/2 ", 2, [["ParaRun", ""], ["CFraction", "〖(x^(1))/2〗"], ["ParaRun", ""]], false, "Check index");

	Test("x^y_1 ", 2, [["ParaRun", ""], ["CDegreeSubSup", "x_(1)^(y)"], ["ParaRun", ""]], false, "Check index degree");
	Test("x^1_i ", 2, [["ParaRun", ""], ["CDegreeSubSup", "x_(i)^(1)"], ["ParaRun", ""]], false, "Check index degree");
	Test("1^x_y ", 2, [["ParaRun", ""], ["CDegreeSubSup", "1_(y)^(x)"], ["ParaRun", ""]], false, "Check index degree");
	Test("x^(1+2)_(g/2) ", 2, [["ParaRun", ""], ["CDegreeSubSup", "x_(〖g/2〗)^(1+2)"], ["ParaRun", ""]], false, "Check index degree");
	Test("x^[1+2]_[g_i] ", 2, [["ParaRun", ""], ["CDegreeSubSup", "x_([g_(i)])^([1+2])"], ["ParaRun", ""]], false, "Check index degree");
	Test("x^[1+2}_[6+1} ", 2, [["ParaRun", ""], ["CDegreeSubSup", "x_([6+1})^([1+2})"], ["ParaRun", ""]], false, "Check index degree");
	Test("x^1/2_1/2 ", 2, [["ParaRun", ""], ["CFraction", "〖(x^(1))/(〖(2_(1))/2〗)〗"], ["ParaRun", ""]], false, "Check index degree");

	 Test("𝑊^3𝛽_𝛿1𝜌1𝜎2 ", 2, [["ParaRun", ""], ["CDegreeSubSup", "𝑊_(𝛿1𝜌1𝜎2)^(3𝛽)"], ["ParaRun", ""]], false, "Check index degree with Unicode symbols");

	Test("(_1^f)f ", 2, [["ParaRun", ""], ["CDegreeSubSup", "(_(1)^(f))f"], ["ParaRun", ""]], false, "Check prescript index degree");
	Test("(_(1/2)^y)f ", 2, [["ParaRun", ""], ["CDegreeSubSup", "(_(〖1/2〗)^(y))f"], ["ParaRun", ""]], false, "Check prescript index degree");
	Test("(_(1/2)^[x_i])x/y  ", 2, [["ParaRun", ""], ["CDegreeSubSup", "(_(〖1/2〗)^([x_(i)]))〖x/y〗"], ["ParaRun", ""]], false, "Check prescript index degree");

	Test("\\sqrt ", 0, [["ParaRun", "√"]], false, "Check");
	Test("\\sqrt (2&1+2) ", 2, [["ParaRun", ""], ["CRadical", "〖√(2&1+2)〗"], ["ParaRun", ""]], false, "Check radical");
	Test("\\sqrt (1+2) ", 2, [["ParaRun", ""], ["CRadical", "〖√(1+2)〗"], ["ParaRun", ""]], false, "Check radical");
	Test("√1 ", 2, [["ParaRun", ""], ["CRadical", "〖√1〗"], ["ParaRun", ""]], false, "Check radical");

	Test("\\cbrt ", 0, [["ParaRun", "∛"]], false, "Check");
	Test("\\cbrt (1+2) ", 2, [["ParaRun", ""], ["CRadical", "〖∛(1+2)〗"], ["ParaRun", ""]], false, "Check radical");
	Test("\\cbrt 1/2 ", 2, [["ParaRun", ""], ["CFraction", "〖(〖∛1〗)/2〗"], ["ParaRun", ""]], false, "Check radical");
	Test("∛1 ", 2, [["ParaRun", ""], ["CRadical", "〖∛1〗"], ["ParaRun", ""]], false, "Check radical");
	Test("∛(1) ", 2, [["ParaRun", ""], ["CRadical", "〖∛1〗"], ["ParaRun", ""]], false, "Check radical");

	Test("\\qdrt ", 0, [["ParaRun", "∜"]], false, "Check");
	Test("\\qdrt (1+2) ", 2, [["ParaRun", ""], ["CRadical", "〖∜(1+2)〗"], ["ParaRun", ""]], false, "Check radical");
	Test("\\qdrt 1/2  ", 2, [["ParaRun", ""], ["CDelimiter", "〖(〖∜1〗)/2〗"], ["ParaRun", ""]], false, "Check radical");
	Test("∜1 ", 2, [["ParaRun", ""], ["CRadical", "〖∜1〗"], ["ParaRun", ""]], false, "Check radical");
	Test("∜(1) ", 2, [["ParaRun", ""], ["CRadical", "〖∜1〗"], ["ParaRun", ""]], false, "Check radical");

	Test("\\rect ", 0, [["ParaRun", "▭"]], false, "Check box literal");
	Test("\\rect 1/2 ", 2, [["ParaRun", ""], ["CFraction", "〖(▭1)/2〗"], ["ParaRun", ""]], false, "Check box");
	Test("\\rect (1/2) ", 2, [["ParaRun", ""], ["CBorderBox", "▭(〖1/2〗)"], ["ParaRun", ""]], false, "Check box");
	Test("▭(𝐸 = 𝑚𝑐^2) ", 2, [["ParaRun", ""], ["CBorderBox", "▭(𝐸=〖𝑚〗^(2))"], ["ParaRun", ""]], false, "Check box");

	Test("\\int ", 0, [["ParaRun", "∫"]], false, "Check large operators");
	Test("\\int  ", 2, [["ParaRun", ""], ["CNary", "〖∫〗"], ["ParaRun", ""]], false, "Check large operators");
	Test("\\int _x ", 2, [["ParaRun", ""], ["CNary", "〖∫_x〗"], ["ParaRun", ""]], false, "Check large operators");
	Test("\\int ^x ", 2, [["ParaRun", ""], ["CNary", "〖∫^x〗"], ["ParaRun", ""]], false, "Check large operators");
	Test("\\int ^(x+1) ", 2, [["ParaRun", ""], ["CNary", "〖∫^(x+1)〗"], ["ParaRun", ""]], false, "Check large operators");
	Test("\\int ^(x+1) ", 2, [["ParaRun", ""], ["CNary", "〖∫^(x+1)〗"], ["ParaRun", ""]],false, "Check large operators");
	Test("\\int ^(x+1)_(1_i) ", 2, [["ParaRun", ""], ["CNary", "〖∫^(x+1)_(1_(i))〗"], ["ParaRun", ""]], false, "Check large operators");

	Test("\\int \\of x ", 2, [["ParaRun", ""], ["CNary", "〖∫▒x〗"], ["ParaRun", ""]], false, "Check large operators");
	Test("\\int _x\\of 1/2  ", 2, [["ParaRun", ""], ["CDelimiter", "〖∫_x▒〖〖1/2〗〗〗"], ["ParaRun", ""]], false, "Check large operators");
	Test("\\int ^x\\of 1/2  ", 2, [["ParaRun", ""], ["CDelimiter", "〖∫^x▒〖〖1/2〗〗〗"], ["ParaRun", ""]], false, "Check large operators");
	Test("\\int _(x+1)\\of 1/2  ", 2, [["ParaRun", ""], ["CNary", "〖∫_(x+1)▒〖〖1/2〗〗〗"], ["ParaRun", ""]], false, "Check large operators");
	Test("\\prod ^(x+1)\\of 1/2  ", 2, [["ParaRun", ""], ["CNary", "〖∏^(x+1)▒〖〖1/2〗〗〗"], ["ParaRun", ""]],false, "Check large operators");
	Test("∫^(x+1)_(1_i)\\of 1/2  ", 2, [["ParaRun", ""], ["CNary", "〖∫^(x+1)_(1_(i))▒〖〖1/2〗〗〗"], ["ParaRun", ""]], false, "Check large operators");

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

	Test("|1+2| |1+2| ", 4, [["ParaRun", ""], ["CDelimiter", "|1+2|"], ["ParaRun", ""],  ["CDelimiter", "|1+2|"], ["ParaRun", ""]], false, "Check brackets");

	Test("sin ", 1, [["ParaRun", ""], ["CMathFunc", "sin "], ["ParaRun", ""]], false, "Check functions");
	Test("cos ", 1, [["ParaRun", ""], ["CMathFunc", "cos "], ["ParaRun", ""]], false, "Check functions");
	Test("tan ", 1, [["ParaRun", ""], ["CMathFunc", "tan "], ["ParaRun", ""]], false, "Check functions");
	Test("csc ", 1, [["ParaRun", ""], ["CMathFunc", "csc "], ["ParaRun", ""]], false, "Check functions");
	Test("sec ", 1, [["ParaRun", ""], ["CMathFunc", "sec "], ["ParaRun", ""]], false, "Check functions");
	Test("cot ", 1, [["ParaRun", ""], ["CMathFunc", "cot "], ["ParaRun", ""]], false, "Check functions");

	Test("sin a", 1, [["ParaRun", ""], ["CMathFunc", "〖sin a〗"], ["ParaRun", ""]], false, "Check functions");
	Test("cos a", 1, [["ParaRun", ""], ["CMathFunc", "〖cos a〗"], ["ParaRun", ""]], false, "Check functions");
	Test("tan a", 1, [["ParaRun", ""], ["CMathFunc", "〖tan a〗"], ["ParaRun", ""]], false, "Check functions");
	Test("csc a", 1, [["ParaRun", ""], ["CMathFunc", "〖csc a〗"], ["ParaRun", ""]], false, "Check functions");
	Test("sec a", 1, [["ParaRun", ""], ["CMathFunc", "〖sec a〗"], ["ParaRun", ""]], false, "Check functions");
	Test("cot a", 1, [["ParaRun", ""], ["CMathFunc", "〖cot a〗"], ["ParaRun", ""]], false, "Check functions");

	Test("sin(1+2_i) ", 2, [["ParaRun", ""], ["CMathFunc", "〖sin 〖 1+2_(i)〗〗"], ["ParaRun", ""]], false, "Check functions");
	Test("cos(1+2_i) ", 2, [["ParaRun", ""], ["CMathFunc", "〖cos 〖 1+2_(i)〗〗"], ["ParaRun", ""]], false, "Check functions");
	Test("tan(1+2_i) ", 2, [["ParaRun", ""], ["CMathFunc", "〖tan 〖 1+2_(i)〗〗"], ["ParaRun", ""]], false, "Check functions");
	Test("csc(1+2_i) ", 2, [["ParaRun", ""], ["CMathFunc", "〖csc 〖 1+2_(i)〗〗"], ["ParaRun", ""]], false, "Check functions");
	Test("sec(1+2_i) ", 2, [["ParaRun", ""], ["CMathFunc", "〖sec 〖 1+2_(i)〗〗"], ["ParaRun", ""]], false, "Check functions");
	Test("cot(1+2_i) ", 2, [["ParaRun", ""], ["CMathFunc", "〖cot 〖 1+2_(i)〗〗"], ["ParaRun", ""]], false, "Check functions");

	Test("log ", 1, [["ParaRun", ""], ["CMathFunc", "log "], ["ParaRun", ""]], false, "Check functions");
	Test("log a ", 1, [["ParaRun", ""], ["CMathFunc", "〖log a 〗"], ["ParaRun", ""]], false, "Check functions");
	Test("log(a+2)  ", 2, [["ParaRun", ""], ["CDelimiter", "〖log 〖 a+2〗〗"], ["ParaRun", ""]], false, "Check functions");
	Test("lim ", 1, [["ParaRun", ""], ["CMathFunc", "lim "], ["ParaRun", ""]], false, "Check functions");
	Test("lim_a ", 1, [["ParaRun", ""], ["CMathFunc", "lim┬a "], ["ParaRun", ""]], false, "Check functions");
	Test("lim^a ", 1, [["ParaRun", ""], ["CMathFunc", "lim┴a "], ["ParaRun", ""]], false, "Check functions");

	Test("min ", 1, [["ParaRun", ""], ["CMathFunc", "min "], ["ParaRun", ""]], false, "Check functions");
	Test("min_a ", 1, [["ParaRun", ""], ["CMathFunc", "min┬a "], ["ParaRun", ""]], false, "Check functions");
	Test("min^a ", 1, [["ParaRun", ""], ["CMathFunc", "min┴a "], ["ParaRun", ""]], false, "Check functions");

	Test("max ", 1, [["ParaRun", ""], ["CMathFunc", "max "], ["ParaRun", ""]], false, "Check functions");
	Test("max_a ", 1, [["ParaRun", ""], ["CMathFunc", "max┬a "], ["ParaRun", ""]], false, "Check functions");
	Test("max^a ", 1, [["ParaRun", ""], ["CMathFunc", "max┴a "], ["ParaRun", ""]], false, "Check functions");

	Test("ln ", 1, [["ParaRun", ""], ["CMathFunc", "ln "], ["ParaRun", ""]], false, "Check functions");
	Test("ln_a ", 1, [["ParaRun", ""], ["CMathFunc", "〖ln〗_(a) "], ["ParaRun", ""]], false, "Check functions");
	Test("ln^a ", 1, [["ParaRun", ""], ["CMathFunc", "〖ln〗^(a) "], ["ParaRun", ""]], false, "Check functions");

	Test("■ ", 0, [["ParaRun", "■ "]], false, "Check matrix");
	Test("■(1&2@3&4) ", 2, [["ParaRun", ""], ["CMathMatrix", "■(1&2@3&4)"], ["ParaRun", ""]], false, "Check matrix");
	Test("■(1&2) ", 2, [["ParaRun", ""], ["CMathMatrix", "■(1&2)"], ["ParaRun", ""]], false, "Check matrix");
	Test("■(&1&2@3&4) ", 2, [["ParaRun", ""], ["CMathMatrix", "■(&1&2@3&4&)"], ["ParaRun", ""]], false, "Check matrix");

	Test("(1\\mid 2\\mid 3) ", 2, [["ParaRun", ""], ["CDelimiter", "(1∣2∣3)"], ["ParaRun", ""]], false, "Check  fraction");
	Test("[1\\mid 2\\mid 3) ", 2, [["ParaRun", ""], ["CDelimiter", "[1∣2∣3)"], ["ParaRun", ""]], false, "Check  fraction");
	Test("|1\\mid 2\\mid 3) ", 2, [["ParaRun", ""], ["CDelimiter", "|1∣2∣3)"], ["ParaRun", ""]], false, "Check  fraction");
	Test("{1\\mid 2\\mid 3) ", 2, [["ParaRun", ""], ["CDelimiter", "{1∣2∣3)"], ["ParaRun", ""]], false, "Check  fraction");
	Test("(1\\mid 2\\mid 3] ", 2, [["ParaRun", ""], ["CDelimiter", "(1∣2∣3]"], ["ParaRun", ""]], false, "Check  fraction");
	Test("(1\\mid 2\\mid 3} ", 2, [["ParaRun", ""], ["CDelimiter", "(1∣2∣3}"], ["ParaRun", ""]], false, "Check  fraction");
	Test("(1\\mid 2\\mid 3| ", 2, [["ParaRun", ""], ["CDelimiter", "(1∣2∣3|"], ["ParaRun", ""]], false, "Check  fraction");
	Test("|1\\mid 2\\mid 3| ", 2, [["ParaRun", ""], ["CDelimiter", "|1∣2∣3|"], ["ParaRun", ""]], false, "Check  fraction");
	Test("{1\\mid 2\\mid 3} ", 2, [["ParaRun", ""], ["CDelimiter", "{1∣2∣3}"], ["ParaRun", ""]], false, "Check  fraction");
	Test("[1\\mid 2\\mid 3] ", 2, [["ParaRun", ""], ["CDelimiter", "[1∣2∣3]"], ["ParaRun", ""]], false, "Check  fraction");

	Test("e\\tilde  ", 2, [["ParaRun", ""], ["CAccent", "ẽ"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\hat  ", 2, [["ParaRun", ""], ["CAccent", "ê"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\breve  ", 2, [["ParaRun", ""], ["CAccent", "ĕ"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\dot  ", 2, [["ParaRun", ""], ["CAccent", "ė"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\ddot  ", 2, [["ParaRun", ""], ["CAccent", "ë"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\dddot  ", 2, [["ParaRun", ""], ["CAccent", "e⃛"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\prime  ", 2, [["ParaRun", ""], ["CAccent", "e′"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\pprime  ", 2, [["ParaRun", ""], ["CAccent", "e″"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\check  ", 2, [["ParaRun", ""], ["CAccent", "ě"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\acute  ", 2, [["ParaRun", ""], ["CAccent", "é"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\grave  ", 2, [["ParaRun", ""], ["CAccent", "è"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\bar  ", 2, [["ParaRun", ""], ["CAccent", "e̅"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\Bar  ", 2, [["ParaRun", ""], ["CAccent", "e̿"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\ubar  ", 2, [["ParaRun", ""], ["CAccent", "e̲"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\Ubar  ", 2, [["ParaRun", ""], ["CAccent", "e̳"], ["ParaRun", ""]], false, "Check diacritics");
	Test("e\\vec  ", 2, [["ParaRun", ""], ["CAccent", "e⃗"], ["ParaRun", ""]], false, "Check diacritics");

	Test("\\alpha ", 0, [["ParaRun", "α"]], true, "Check LaTeX words");
	Test("\\Alpha ", 0, [["ParaRun", "Α"]], true, "Check LaTeX words");
	Test("\\beta ", 0, [["ParaRun", "β"]], true, "Check LaTeX words");
	Test("\\Beta ", 0, [["ParaRun", "Β"]], true, "Check LaTeX words");
	Test("\\gamma ", 0, [["ParaRun", "γ"]], true, "Check LaTeX words");
	Test("\\Gamma ", 0, [["ParaRun", "Γ"]], true, "Check LaTeX words");
	Test("\\pi ", 0, [["ParaRun", "π"]], true, "Check LaTeX words");
	Test("\\Pi ", 0, [["ParaRun", "Π"]], true, "Check LaTeX words");
	Test("\\phi ", 0, [["ParaRun", "ϕ"]], true, "Check LaTeX words");
	Test("\\varphi ", 0, [["ParaRun", "φ"]], true, "Check LaTeX words");
	Test("\\mu ", 0, [["ParaRun", "μ"]], true, "Check LaTeX words");
	Test("\\Phi ", 0, [["ParaRun", "Φ"]], true, "Check LaTeX words");

	Test("\\cos(2\\theta ) ", 2, [["ParaRun", ""], ["CMathFunc", "\\cos { (2θ)}"], ["ParaRun", ""]], true, "Check LaTeX function");
	Test("\\lim_{x\\to \\infty }\\exp(x) ", 2, [["ParaRun", ""], ["CMathFunc", "\\lim_{x→∞} { \\exp { (x)}}"], ["ParaRun", ""]], true, "Check LaTeX function");

	Test("k^{n+1} ", 2, [["ParaRun", ""], ["CDegree", "k^{n+1}"], ["ParaRun", ""]], true, "Check LaTeX degree");
	Test("n^2 ", 2, [["ParaRun", ""], ["CDegree", "n^{2}"], ["ParaRun", ""]], true, "Check LaTeX degree");
	Test("n^{2} ", 2, [["ParaRun", ""], ["CDegree", "n^{2}"], ["ParaRun", ""]], true, "Check LaTeX degree");
	Test("n^(2) ", 2, [["ParaRun", ""], ["CDegree", "n^{(2)}"], ["ParaRun", ""]], true, "Check LaTeX degree");

	Test("k_{n+1} ", 2, [["ParaRun", ""], ["CDegree", "k_{n+1}"], ["ParaRun", ""]], true, "Check LaTeX degree");
	Test("n_2 ", 2, [["ParaRun", ""], ["CDegree", "n_{2}"], ["ParaRun", ""]], true, "Check LaTeX degree");
	Test("n_{2} ", 2, [["ParaRun", ""], ["CDegree", "n_{2}"], ["ParaRun", ""]], true, "Check LaTeX degree");
	Test("n_(2) ", 2, [["ParaRun", ""], ["CDegree", "n_{(2)}"], ["ParaRun", ""]], true, "Check LaTeX degree");

	Test("\\frac{12}{x} ", 2, [["ParaRun", ""], ["CFraction", "\\cos { (2θ)}"], ["ParaRun", ""]], true, "Check LaTeX fraction");
	Test("\\frac12b ", 2, [["ParaRun", ""], ["CFraction", "\\cos { (2θ)}"], ["ParaRun", ""]], true, "Check LaTeX fraction");
	Test("\\binom{12}{x} ", 2, [["ParaRun", ""], ["CFraction", "\\cos { (2θ)}"], ["ParaRun", ""]], true, "Check LaTeX fraction");


 })


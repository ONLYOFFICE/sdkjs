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

	function AddFromMathML(strMathMl)
	{
		return AscWord.ParaMath.fromMathML(strMathMl);
	}

	QUnit.testStart(function (){
		AscTest.ClearDocument();
		AscCommon.History.Clear();
		Clear();
		Init();
	})


	QUnit.module("Presentation MathML", function (assert)
	{
		QUnit.module("<math>", function (assert)
		{
			// attributes
		})

		QUnit.module("MathML Core Attributes", function (assert)
		{
			QUnit.test('mathcolor in <mi>', function (assert)
			{
				math = AddFromMathML(`<math><mi mathcolor="red">x</mi></math>`);
				let math_data = math.GetTextOfElement();
				let math_text = math_data.GetText();
				let style = math_data.arr[0].additionalMathData.style;
				let color = style.Color;
				let colorText =`${color.r}, ${color.g}, ${color.b}`;

				assert.strictEqual(
					colorText,
					"255, 0, 0",
					"Check attribute mathcolor"
				);
			})

			QUnit.test('mathbackground in <mi>', function (assert)
			{
				math = AddFromMathML(`<math><mi mathbackground="blue">y</mi></math>`);
				math_data = math.GetTextOfElement();
				math_text = math_data.GetText();
				style = math_data.arr[0].additionalMathData.style;
				let highLight = style.HighLight;
				let highLightText =`${highLight.r}, ${highLight.g}, ${highLight.b}`;

				assert.strictEqual(
					highLightText,
					"0, 0, 255",
					"Check attribute mathbackground"
				);
			})

			QUnit.test('mathcolor & mathbackground in <mi>', function (assert)
			{
				math = AddFromMathML(`<math><mi mathcolor="red" mathbackground="blue">z</mi></math>`);
				math_data = math.GetTextOfElement();
				math_text = math_data.GetText();
				style = math_data.arr[0].additionalMathData.style;
				highLight = style.HighLight;
				color = style.Color;
				highLightText =`${highLight.r}, ${highLight.g}, ${highLight.b}`;
				colorText =`${color.r}, ${color.g}, ${color.b}`;

				assert.strictEqual(
					colorText + '-' +  highLightText,
					"255, 0, 0-0, 0, 255" ,
					"Check attribute mathbackground and mathcolor"
				);
			})

			QUnit.test('mathcolor in <sqrt>', function (assert)
			{
				math = AddFromMathML(`<math><mi mathcolor="red">x</mi></math>`);
				let math_data = math.GetTextOfElement();
				let math_text = math_data.GetText();
				let style = math_data.arr[0].additionalMathData.style;
				let color = style.Color;
				let colorText =`${color.r}, ${color.g}, ${color.b}`;

				assert.strictEqual(
					colorText,
					"255, 0, 0",
					"Check attribute mathcolor"
				);
			})

			QUnit.test('mathbackground in <sqrt>', function (assert)
			{
				math = AddFromMathML(`<math><mi mathbackground="blue">y</mi></math>`);
				math_data = math.GetTextOfElement();
				math_text = math_data.GetText();
				style = math_data.arr[0].additionalMathData.style;
				let highLight = style.HighLight;
				let highLightText =`${highLight.r}, ${highLight.g}, ${highLight.b}`;

				assert.strictEqual(
					highLightText,
					"0, 0, 255",
					"Check attribute mathbackground"
				);
			})

			QUnit.test('mathcolor & mathbackground in <sqrt>', function (assert)
			{
				math = AddFromMathML(`<math><mi mathcolor="red" mathbackground="blue">z</mi></math>`);
				math_data = math.GetTextOfElement();
				math_text = math_data.GetText();
				style = math_data.arr[0].additionalMathData.style;
				highLight = style.HighLight;
				color = style.Color;
				highLightText =`${highLight.r}, ${highLight.g}, ${highLight.b}`;
				colorText =`${color.r}, ${color.g}, ${color.b}`;

				assert.strictEqual(
					colorText + '-' +  highLightText,
					"255, 0, 0-0, 0, 255" ,
					"Check attribute mathbackground and mathcolor"
				);
			})
		})

		QUnit.module("Token Elements", function ()
		{
			QUnit.module('Attributes', function ()
			{
				QUnit.test('mathvariant', function (assert)
				{
					math = AddFromMathML(`<math><mi>c</mi></math>`);
					let math_data = math.GetTextOfElement();
					let math_text = math_data.GetText();
					let style = math_data.arr[0].additionalMathData.style;
					
	
					assert.strictEqual(
						colorText,
						"255, 0, 0",
						"Check attribute mathcolor"
					);
				})

				// QUnit.test('mathsize', function (assert)
				// {
				
				// })

				// QUnit.test('dir', function (assert)
				// {
				
				// })
			})
			QUnit.module('mi', function ()
			{
				QUnit.module('Attributes', function ()
				{
					QUnit.test('Check default mathvariant is italic if content length = 1', function (assert)
					{
						let math = AddFromMathML(`<math><mi>c</mi></math>`);
						assert.strictEqual(
							math.GetTextOfElement().GetText(),
							"𝑐",
							"Add <mi> with mathvariant is italic"
						);
					})

					QUnit.test('Check default mathvariant is normal if content length > 1', function (assert)
					{
						let math = AddFromMathML(`<math><mi>cx</mi></math>`);
						assert.strictEqual(
							math.GetTextOfElement().GetText(),
							"cx",
							"Add <mi> with mathvariant is normal"
						);
					})

					QUnit.test('Add <mi> with normal mathvariant = normal', function (assert)
					{
						let math = AddFromMathML(`<math><mi mathvariant="normal">c</mi></math>`);
						assert.strictEqual(
							math.GetTextOfElement().GetText(),
							"c",
							"Add <mi> with 'c' text"
						);
					})

					QUnit.test('Add <mi> with custom mathvariant = fraktur', function (assert)
					{
						let math = AddFromMathML(`<math><mi mathvariant="fraktur">xy</mi></math>`);
						assert.strictEqual(
							math.GetTextOfElement().GetText(),
							"𝔵𝔶",
							"Add <mi> with '𝔵𝔶' text"
						);
					})
				})

				QUnit.test('Add mi as text', function (assert)
				{
					let math = AddFromMathML(`<math><mi>c</mi></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						"𝑐",
						"Add <mi> with '𝑐'"
					);
				})
			})

			QUnit.module('mn', function ()
			{
				QUnit.test('Add mn as text', function (assert)
				{
					let math = AddFromMathML(`<math><mn>2</mn></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						"2",
						"Add <mi> with '2'"
					);
				})
			})

			QUnit.module('mo', function ()
			{
				QUnit.test('Add mo', function (assert)
				{
					let math = AddFromMathML(`<math><mo>+</mo></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						"+",
						"Add <mo> with '+'"
					);
				})

				QUnit.module('Dictionary-based attributes', function ()
				{
					// form
					// lspace
					// rspace
					// stretchy
					// symmetric
					// maxsize
					// minsize
					// largeop
					//			Ooxml control all this attributes. For now unesessary

					QUnit.test('movablelimits default attribute', function (assert)
					{
						let math = AddFromMathML(`<math displaystyle="inline"><munderover><mo>∑</mo><mi>i</mi><mi>n</mi></munderover></math>`);
						assert.strictEqual(
							math.Root.Content[1].Pr.limLoc,
							NARY_UndOvr,
							"Limloc is nary"
						);
					})

					QUnit.test('movablelimits attribute', function (assert)
					{
						let math = AddFromMathML(`<math display="inline"><munderover><mo movablelimits="true">∑</mo><mi>i</mi><mi>n</mi></munderover></math>`);
						assert.strictEqual(
							math.Root.Content[1].Pr.limLoc,
							NARY_SubSup,
							"Limloc is supSub"
						);
					})


				})

				QUnit.module('Linebreaking attributes', function ()
				{
					// for now only one variant
					QUnit.test('linebreak attribute', function (assert)
					{
						let math = AddFromMathML(`<math displaystyle="inline"><munderover><mo>∑</mo><mi>i</mi><mi>n</mi></munderover></math>`);
						assert.strictEqual(
							math.Root.Content[1].Pr.limLoc,
							NARY_UndOvr,
							"Limloc is nary"
						);
					})

					// lineleading
					// linebreakstyle
					// linebreakmultchar
				});

				QUnit.module('Indentation attributes', function ()
				{
					// indentalign
					// indentshift
					// indenttarget
					// indentalignfirst
					// indentshiftfirst
					// indentalignlast
					// indentshiftlast
				});


			})

			QUnit.module('mtext', function (){
				QUnit.test('Add mtext #1', function (assert)
				{
					let math = AddFromMathML(`<math><mtext> Theorem 1: </mtext></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						"Theorem 1:",
						"Add <mtext> with 'Theorem 1:'"
					);
				})

				QUnit.test('Add mtext #2', function (assert)
				{
					let math = AddFromMathML(`<math><mtext> &#x2009; </mtext></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						" ",
						"Add <mtext> with ' '"
					);
				})

				QUnit.test('Add mtext #3', function (assert)
				{
					let math = AddFromMathML(`<math>
						<mi>y</mi>
						<mo>=</mo>
						<mrow>
							<msup>
							<mi>x</mi>
							<mn>2</mn>
							</msup>
							<mtext>&nbsp;if&nbsp;</mtext>
							<mrow>
							<mi>x</mi>
							<mo>≥</mo>
							<mn>1</mn>
							</mrow>
							<mtext>&nbsp;and&nbsp;</mtext>
							<mn>2</mn>
							<mtext>&nbsp;otherwise.</mtext>
						</mrow>
					</math>`);

					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						"𝑦=𝑥^2 if 𝑥≥1 and 2 otherwise.",
						"Add <mtext> with ' '"
					);
				})
			})

			QUnit.module('space', function (){
				// QUnit.test('Add mspace #1', function (assert)
				// {
				// 	let math = AddFromMathML(`<math><mtext> Theorem 1: </mtext></math>`);
				// 	assert.strictEqual(
				// 		math.GetTextOfElement().GetText(),
				// 		"Theorem 1:",
				// 		"Add <mtext> with 'Theorem 1:'"
				// 	);
				// })
			})

			QUnit.module('ms', function (){
				QUnit.test('Add ms #1', function (assert)
				{
					let math = AddFromMathML(`<math><ms> n </ms></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						"\"n\"",
						"Add <ms> with '\"n\"'"
					);
				})

				QUnit.test('Add ms with custom lquote attribute', function (assert)
				{
					let math = AddFromMathML(`<math><ms lquote="1"> n </ms></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						"1n\"",
						"Add <ms> with '1n\"'"
					);
				})

				QUnit.test('Add ms with custom rquote attribute', function (assert)
				{
					let math = AddFromMathML(`<math><ms rquote="2"> n </ms></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						"\"n2",
						"Add <ms> with '\"n2'"
					);
				})
			})

			QUnit.module('mfrac', function (){
				QUnit.test('Add mfrac', function (assert)
				{
					let math = AddFromMathML(`<math><mfrac><mi> a </mi><mi> b </mi></mfrac></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						"𝑎/𝑏",
						"Add <ms> with '𝑎/𝑏'"
					);
				})

				QUnit.test('Add complex mfrac', function (assert)
				{
					let math = AddFromMathML(`<math>
						<mfrac>
						<mrow>
							<mo> ( </mo>
								<mfrac>
									<mi> a </mi>
									<mi> b </mi>
								</mfrac>
							<mo> ) </mo>
							<mfrac>
								<mi> a </mi>
								<mi> b </mi>
							</mfrac>
						</mrow>
						<mfrac>
							<mi> c </mi>
							<mi> d </mi>
						</mfrac>
						</mfrac>
					</math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						"((𝑎/𝑏) 𝑎/𝑏)/(𝑐/𝑑)",
						"Add <ms> with '((𝑎/𝑏)𝑎/𝑏)/(𝑐/𝑑)'"
					);
				})

				QUnit.module('Attributes', function ()
				{
					// QUnit.test('linethickness', function (assert)
					// {})
					// QUnit.test('numalign', function (assert)
					// {})
					// QUnit.test('denomalign', function (assert)
					// {})

					QUnit.test('bevelled', function (assert)
					{
						let math = AddFromMathML(`<math><mfrac bevelled="true"><mi> a </mi><mi> b </mi></mfrac></math>`);
						assert.strictEqual(
							math.GetTextOfElement().GetText(),
							"𝑎⁄𝑏",
							"Add <ms> with '𝑎⁄𝑏'"
						);
					})


				})
			})

			QUnit.module('msqrt', function (){
				QUnit.test('Add msqrt', function (assert)
				{
					let math = AddFromMathML(`<math><msqrt><mi>x</mi></msqrt></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						"√(𝑥)",
						"Add <ms> with '√(𝑥)'"
					);
				})
			})

			QUnit.module('mroot', function (){
				QUnit.test('Add mroot', function (assert)
				{
					let math = AddFromMathML(`<math><mroot><mi>x</mi><mn>5</mn></mroot></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						"√(5&𝑥)",
						"Add <ms> with '√(5&𝑥)'"
					);
				})
			})

			QUnit.module('mstyle', function (){
			})

			QUnit.module('merror', function (){
				QUnit.test('Add merror', function (assert)
				{
					let math = AddFromMathML(`
						<math>
						<merror>
							<mtext> Unrecognized element: mfraction; arguments were:</mtext>
							<mrow> <mn> 1 </mn> <mo> + </mo> <msqrt> <mn> 5 </mn> </msqrt> </mrow>
							<mtext> and </mtext>
							<mn> 2 </mn>
						</merror>
						</math>
					`);
					
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						"Unrecognized element: mfraction; arguments were:1+√5and2",
						"Add <merror>"
					);
				})
			})

			QUnit.module('mpadded', function (){
				QUnit.test('Add mpadded', function (assert)
				{
					let math = AddFromMathML(`<math>
						<mpadded lspace="2em" voffset="-1em" height="1em" depth="3em" width="7em" style="background: lightblue;">
						<mfrac>
							<mn>23456</mn>
							<mn>78</mn>
						</mfrac>
						</mpadded>
					</math>`);

					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'23456/78',
						"Check text"
					);
				})
			})

			QUnit.module('mphantom', function (){
				QUnit.test('Add mphantom', function (assert)
				{
					let math = AddFromMathML(`<math><mphantom><mi>x</mi></mphantom></math>`);

					assert.strictEqual(
						math.Root.Content[1] instanceof CPhantom ,
						true,
						"Add <mphantom> with 'x'"
					);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'\\mphantom𝑥',
						"Check text"
					);
				})
			})

			QUnit.module('mfenced', function (){
				QUnit.test('Add mfenced #1', function (assert)
				{
					let math = AddFromMathML(`<math><mfenced><mi>x</mi></mfenced></math>`);

					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'(𝑥)',
						"Check text"
					);
				})

				QUnit.test('Add mfenced #2', function (assert)
				{
					let math = AddFromMathML(`<math><mfenced><mi>x</mi><mi>y</mi></mfenced></math>`);
					console.log(math)
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'(𝑥∣𝑦)',
						"Check text"
					);
					
					assert.strictEqual(
						math.Root.Content[1].Pr.sepChr,
						44,
						"Check separator"
					);
				})
			})

			QUnit.module('menclose', function (){
				QUnit.test('Add menclose #1', function (assert)
				{
					let math = AddFromMathML(`<math><menclose><mi> x </mi><mo> + </mo><mi> y </mi></menclose></math>`);

					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'▭(𝑥+𝑦)',
						"Menclose always convert to simple box"
					);
				})
			})

			QUnit.module('msub', function (){
				QUnit.test('Add msub', function (assert)
				{
					let math = AddFromMathML(`<math><msub> <mi> ⅇ </mi><mi> x </mi></msub></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'ⅇ_𝑥',
						"Check add msub"
					);
				})
			})

			QUnit.module('msup', function (){
				QUnit.test('Add msup', function (assert)
				{
					let math = AddFromMathML(`<math><msup> <mi> ⅇ </mi><mi> x </mi></msup></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'ⅇ^𝑥',
						"Check add msup"
					);
				})
			})

			QUnit.module('msubsup', function (){
				QUnit.test('Add msubsup', function (assert)
				{
					let math = AddFromMathML(`<math><msubsup><mi> 2 </mi><mi> ⅇ </mi><mi> x </mi></msubsup></math>`);
					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'2_ⅇ^𝑥',
						"Check add msubsup"
					);
				})
			})

			QUnit.module('munder', function (){
				QUnit.test('Add munder #1', function (assert)
				{
					let math = AddFromMathML(`<math>
						<munder accentunder="true">
							<mrow>
							<mi> x </mi>
							<mo> + </mo>
							<mi> y </mi>
							<mo> + </mo>
							<mi> z </mi>
							</mrow>
							<mo> ⏟ </mo>
						</munder>
					</math>`);

					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'⏟(𝑥+𝑦+𝑧)',
						"Check add munder"
					);
				})

				QUnit.test('Add munder #2', function (assert)
				{
					let math = AddFromMathML(`<math>
						<munder accentunder="false">
							<mrow>
							<mi> x </mi>
							<mo> + </mo>
							<mi> y </mi>
							<mo> + </mo>
							<mi> z </mi>
							</mrow>
							<mo> ⏟ </mo>
						</munder>
					</math>`);

					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'⏟(𝑥+𝑦+𝑧)',
						"Check add munder"
					);
				})

			})

			QUnit.module('mover', function (){
				QUnit.test('Add mover #1', function (assert)
				{
					let math = AddFromMathML(`<math>
						 <mover accent="true">
							<mi> x </mi>
							<mo> ^ </mo>
						</mover>
					</math>`);

					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'𝑥┴^',
						"Check add mover"
					);
				})

				QUnit.test('Add mover #2', function (assert)
				{
					let math = AddFromMathML(`<math>
						<mover>
							<mn>3</mn>
							<mn>4</mn>
						</mover>
					</math>`);

					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'3┴4',
						"Check add mover"
					);
				})

				QUnit.test('Add mover #3', function (assert)
				{
					let math = AddFromMathML(`<math>
						<mover accent="true">
							<mrow>
							<mi> x </mi>
							<mo> + </mo>
							<mi> y </mi>
							<mo> + </mo>
							<mi> z </mi>
							</mrow>
							<mo> ⏞ </mo>
						</mover>
						<mover accent="false">
							<mrow>
							<mi> x </mi>
							<mo> + </mo>
							<mi> y </mi>
							<mo> + </mo>
							<mi> z </mi>
							</mrow>
							<mo> ⏞ </mo>
						</mover>
					</math>`);

					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'⏞(𝑥+𝑦+𝑧) ⏞(𝑥+𝑦+𝑧)',
						"Check add mover"
					);
				})
			})

			QUnit.module('munderover', function (){
				QUnit.test('Add munderover #1', function (assert)
				{
					let math = AddFromMathML(`<math>
						 <munderover>
							<mn>5</mn>
							<mn>6</mn>
							<mn>7</mn>
						</munderover>
					</math>`);

					assert.strictEqual(
						math.GetTextOfElement().GetText(),
						'(5┴7)┬6',
						"Check add munderover"
					);
				})
			})

			// QUnit.module('mmultiscripts', function (){
			// 	QUnit.test('Add mmultiscripts', function (assert)
			// 	{
			// 		let math = AddFromMathML(`<math>
			// 			<mmultiscripts>
			// 				<mn>1</mn>
			// 				<mn>2</mn>
			// 				<mn>3</mn>
			// 				<mrow></mrow>
			// 				<mn>5</mn>
			// 				<mprescripts/>
			// 				<mn>6</mn>
			// 				<mrow></mrow>
			// 				<mn>8</mn>
			// 				<mn>9</mn>
			// 			</mmultiscripts>
			// 		</math>`);

			// 		assert.strictEqual(
			// 			math.GetTextOfElement().GetText(),
			// 			'(5┴7)┬6',
			// 			"Check add mmultiscripts"
			// 		);
			// 	})
			// })

			QUnit.module('mtable', function (){
				QUnit.test('Normalize nested mtr in mtable', function (assert)
				{
					let math = AddFromMathML(`<math display="block">
						<semantics>
							<mtable>
								<mtr>
									<mtr>
										<mtd><mi>a</mi></mtd>
										<mtd><mo>=</mo></mtd>
										<mtd><mn>0</mn></mtd>
									</mtr>
									<mtr>
										<mtd><mi>b</mi></mtd>
										<mtd><mo>=</mo></mtd>
										<mtd><mn>1</mn></mtd>
									</mtr>
									<mtr>
										<mtd><mi>c</mi></mtd>
										<mtd><mo>=</mo></mtd>
										<mtd><mn>2</mn></mtd>
									</mtr>
								</mtr>
							</mtable>
						</semantics>
					</math>`);

					let content = math.Root.Content[1];
					assert.ok(content instanceof CMathMatrix, "Check matrix created");

					let rows = content.getRowsCount ? content.getRowsCount() : 0;
					assert.strictEqual(rows, 3, "Check 3 rows created from nested mtr");
					
					let Row1 = content.getContentElement(0,0).GetTextOfElement().GetText();
					let Row2 = content.getContentElement(1,0).GetTextOfElement().GetText();
					let Row3 = content.getContentElement(2,0).GetTextOfElement().GetText();

					assert.ok(Row1 === "𝑎" && Row2 === "𝑏" && Row3 === "𝑐", "Check all rows present");
				})

				QUnit.test('Normalize nested mtd in mtable', function (assert)
				{
					let math = AddFromMathML(`<math display="block">
						
							<mtable>
								<mtd>
									<mtd><mi>x</mi></mtd>
									<mtd><mo>=</mo></mtd>
									<mtd><mn>1</mn></mtd>
								</mtd>
							</mtable>
						
					</math>`);

					let content = math.Root.Content[1];
					
					assert.ok(content instanceof CMathMatrix, "Check matrix created");

					let rows = content.getRowsCount ? content.getRowsCount() : 0;
					assert.strictEqual(rows, 1, "Check 1 row created from nested mtd");

					let cols = content.getColsCount ? content.getColsCount() : 0;
					assert.strictEqual(cols, 3, "Check 3 columns created from nested mtd");

					let cell1 = content.getContentElement(0,0).GetTextOfElement().GetText();
					let cell2 = content.getContentElement(0,1).GetTextOfElement().GetText();
					let cell3 = content.getContentElement(0,2).GetTextOfElement().GetText();
					assert.ok(cell1 === "𝑥" && cell2 === "=" && cell3 === "1", "Check all cells present");
				})

				QUnit.test('Normalize mixed nested mtr and mtd in mtable', function (assert)
				{
					let math = AddFromMathML(`<math display="block">
						<semantics>
							<mtable>
								<mtr>
									<mtr>
										<mtd><mi>a</mi></mtd>
										<mtd><mo>=</mo></mtd>
										<mtd><mn>0</mn></mtd>
									</mtr>
								</mtr>
								<mtd>
									<mtd><mi>b</mi></mtd>
									<mtd><mo>=</mo></mtd>
									<mtd><mn>1</mn></mtd>
								</mtd>
								<mtr>
									<mtd><mi>c</mi></mtd>
									<mtd><mo>=</mo></mtd>
									<mtd><mn>2</mn></mtd>
								</mtr>
							</mtable>
						</semantics>
					</math>`);

					let content = math.Root.Content[1];
					assert.ok(content instanceof CMathMatrix, "Check matrix created");

					let rows = content.getRowsCount ? content.getRowsCount() : 0;
					assert.strictEqual(rows, 3, "Check 3 rows created from mixed structure");

					let row1 = content.getContentElement(0,0).GetTextOfElement().GetText();
					let row2 = content.getContentElement(1,0).GetTextOfElement().GetText();
					let row3 = content.getContentElement(2,0).GetTextOfElement().GetText();

					assert.ok(row1 === "𝑎" && row2 === "𝑏" && row3 === "𝑐", "Check all rows present in mixed structure");
				})
			})
		})
	})
 })

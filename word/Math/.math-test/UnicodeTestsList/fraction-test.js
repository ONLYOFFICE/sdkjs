function fractionTests(test) {
	test(
		`1/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					NumberLiteral: "1",
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка работы базового деления: 1/2"
	);
	test(
		`x/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					CharLiteral: "x",
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка работы базового деления: x/2"
	);
	test(
		`x+5/2`,
		{
			type: "UnicodeEquation",
			body: [
				{
					CharLiteral: "x",
				},
				{
					Operator: "+",
				},
				{
					type: "fractionLiteral",
					up: {
						NumberLiteral: "5",
					},
					opOver: "/",
					down: {
						NumberLiteral: "2",
					},
				},
			],
		},
		"Проверка работы базового деления: x+5/2"
	);
	test(
		`x+5/x+2`,
		{
			type: "UnicodeEquation",
			body: [
				{
					CharLiteral: "x",
				},
				{
					Operator: "+",
				},
				{
					type: "fractionLiteral",
					up: {
						NumberLiteral: "5",
					},
					opOver: "/",
					down: {
						CharLiteral: "x",
					},
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "2",
				},
			],
		},
		"Проверка работы базового деления: x+5/x+2"
	);
	test(
		`1∕2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					NumberLiteral: "1",
				},
				opOver: "∕",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка работы базового деления: 1∕2"
	);
	test(
		`(x+5)/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expBracketLiteral",
					exp: [
						{
							CharLiteral: "x",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "5",
						},
					],
					open: "(",
					close: ")",
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка работы деления: (x+5)/2"
	);
	test(
		`x/(2+1)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					CharLiteral: "x",
				},
				opOver: "/",
				down: {
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "2",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "1",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка работы деления: x/(2+1)"
	);
	test(
		`(x-5)/(2+1)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expBracketLiteral",
					exp: [
						{
							CharLiteral: "x",
						},
						{
							Operator: "-",
						},
						{
							NumberLiteral: "5",
						},
					],
					open: "(",
					close: ")",
				},
				opOver: "/",
				down: {
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "2",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "1",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка работы деления: (x-5)/(2+1)"
	);
	test(
		`1+3/2/3`,
		{
			type: "UnicodeEquation",
			body: [
				{
					NumberLiteral: "1",
				},
				{
					Operator: "+",
				},
				{
					type: "fractionLiteral",
					up: {
						NumberLiteral: "3",
					},
					opOver: "/",
					down: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "2",
						},
						opOver: "/",
						down: {
							NumberLiteral: "3",
						},
					},
				},
			],
		},
		"Проверка работы цепи деления: 1+3/2/3"
	);
	test(
		`(𝛼_2^3)/(𝛽_2^3+𝛾_2^3)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expBracketLiteral",
					exp: {
						type: "expSubsup",
						base: {
							type: "anOther",
							value: "𝛼",
						},
						down: {
							type: "soperandLiteral",
							operand: {
								NumberLiteral: "2",
							},
						},
						up: {
							type: "soperandLiteral",
							operand: {
								NumberLiteral: "3",
							},
						},
					},
					open: "(",
					close: ")",
				},
				opOver: "/",
				down: {
					type: "expBracketLiteral",
					exp: [
						{
							type: "expSubsup",
							base: {
								type: "anOther",
								value: "𝛽",
							},
							down: {
								type: "soperandLiteral",
								operand: {
									NumberLiteral: "2",
								},
							},
							up: {
								type: "soperandLiteral",
								operand: {
									NumberLiteral: "3",
								},
							},
						},
						{
							Operator: "+",
						},
						{
							type: "expSubsup",
							base: {
								type: "anOther",
								value: "𝛾",
							},
							down: {
								type: "soperandLiteral",
								operand: {
									NumberLiteral: "2",
								},
							},
							up: {
								type: "soperandLiteral",
								operand: {
									NumberLiteral: "3",
								},
							},
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка работы деления: (𝛼_2^3)/(𝛽_2^3+𝛾_2^3)"
	);

	test(
		`(a/(b+c))/(d/e + f)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							CharLiteral: "a",
						},
						opOver: "/",
						down: {
							type: "expBracketLiteral",
							exp: [
								{
									CharLiteral: "b",
								},
								{
									Operator: "+",
								},
								{
									CharLiteral: "c",
								},
							],
							open: "(",
							close: ")",
						},
					},
					open: "(",
					close: ")",
				},
				opOver: "/",
				down: {
					type: "expBracketLiteral",
					exp: [
						{
							type: "fractionLiteral",
							up: {
								CharLiteral: "d",
							},
							opOver: "/",
							down: {
								CharLiteral: "e",
							},
						},
						{
							type: "SpaceLiteral",
							value: " ",
						},
						{
							Operator: "+",
						},
						{
							type: "SpaceLiteral",
							value: " ",
						},
						{
							CharLiteral: "f",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка работы вложенного деления: (a/(b+c))/(d/e + f)"
	);

	test(
		`(a/(c/(z/x)))`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expBracketLiteral",
				exp: {
					type: "fractionLiteral",
					up: {
						CharLiteral: "a",
					},
					opOver: "/",
					down: {
						type: "expBracketLiteral",
						exp: {
							type: "fractionLiteral",
							up: {
								CharLiteral: "c",
							},
							opOver: "/",
							down: {
								type: "expBracketLiteral",
								exp: {
									type: "fractionLiteral",
									up: {
										CharLiteral: "z",
									},
									opOver: "/",
									down: {
										CharLiteral: "x",
									},
								},
								open: "(",
								close: ")",
							},
						},
						open: "(",
						close: ")",
					},
				},
				open: "(",
				close: ")",
			},
		},
		"Проверка работы деления: (a/(c/(z/x)))"
	);

	test(
		`1¦2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "binomLiteral",
				numerator: {
					NumberLiteral: "1",
				},
				operand: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка работы деления: 1¦2"
	);
	test(
		`(1¦2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expBracketLiteral",
				exp: {
					type: "binomLiteral",
					numerator: {
						NumberLiteral: "1",
					},
					operand: {
						NumberLiteral: "2",
					},
				},
				open: "(",
				close: ")",
			},
		},
		"Проверка работы деления: (1¦2)"
	);
}
window["AscMath"].fraction = fractionTests;

function sqrtTests(test) {
	test(
		`√5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					NumberLiteral: "5",
				},
			},
		},
		"Check √5"
	);
	test(
		`√a`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					CharLiteral: "a",
				},
			},
		},
		"Check √a"
	);
	test(
		`√a/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "fractionLiteral",
					up: {
						CharLiteral: "a",
					},
					opOver: "/",
					down: {
						NumberLiteral: "2",
					},
				},
			},
		},
		"Check √a/2"
	);
	test(
		`√(2&a-4)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "nthrtLiteral",
				index: {
					NumberLiteral: "2",
				},
				content: [
					{
						CharLiteral: "a",
					},
					{
						Operator: "-",
					},
					{
						NumberLiteral: "4",
					},
				],
			},
		},
		"Check √(2&a-4)"
	);
	test(
		`∛5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "cubertLiteral",
				value: {
					NumberLiteral: "5",
				},
			},
		},
		"Check ∛5"
	);
	test(
		`∛a`,
		{
			type: "UnicodeEquation",
			body: {
				type: "cubertLiteral",
				value: {
					CharLiteral: "a",
				},
			},
		},
		"Check ∛a"
	);
	test(
		`∛a/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "cubertLiteral",
					value: {
						CharLiteral: "a",
					},
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Check ∛a/2"
	);
	test(
		`∛(a-4)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "cubertLiteral",
				value: {
					type: "expBracketLiteral",
					exp: [
						{
							CharLiteral: "a",
						},
						{
							Operator: "-",
						},
						{
							NumberLiteral: "4",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Check ∛(a-4)"
	);
	test(
		`∜5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fourthrtLiteral",
				value: {
					NumberLiteral: "5",
				},
			},
		},
		"Check ∜5"
	);
	test(
		`∜a`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fourthrtLiteral",
				value: {
					CharLiteral: "a",
				},
			},
		},
		"Check ∜a"
	);
	test(
		`∜a/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "fourthrtLiteral",
					value: {
						CharLiteral: "a",
					},
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Check ∜a/2"
	);
	test(
		`∜(a-4)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fourthrtLiteral",
				value: {
					type: "expBracketLiteral",
					exp: [
						{
							CharLiteral: "a",
						},
						{
							Operator: "-",
						},
						{
							NumberLiteral: "4",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Check ∜(a-4)"
	);
	test(
		`√(10&a/4)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "nthrtLiteral",
				index: {
					NumberLiteral: "10",
				},
				content: {
					type: "fractionLiteral",
					up: {
						CharLiteral: "a",
					},
					opOver: "/",
					down: {
						NumberLiteral: "4",
					},
				},
			},
		},
		"Check √(10&a/4)"
	);
	test(
		`√(10^2&a/4+2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "nthrtLiteral",
				index: {
					type: "expSuperscript",
					base: [
						{
							NumberLiteral: "10",
						},
					],
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
				content: [
					{
						type: "fractionLiteral",
						up: {
							CharLiteral: "a",
						},
						opOver: "/",
						down: {
							NumberLiteral: "4",
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
		},
		"Check √(10^2&a/4+2)"
	);
	test(
		`√5^2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expSuperscript",
					base: [
						{
							NumberLiteral: "5",
						},
					],
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
			},
		},
		"Check √5^2"
	);
	test(
		`√5_2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expSubscript",
					base: [
						{
							NumberLiteral: "5",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
			},
		},
		"Check √5_2"
	);
	test(
		`√5^2_x`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expSubsup",
					base: [
						{
							NumberLiteral: "5",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "x",
						},
					},
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
			},
		},
		"Check √5^2_x"
	);
	test(
		`√5_2^x`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expSubsup",
					base: [
						{
							NumberLiteral: "5",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
					up: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "x",
						},
					},
				},
			},
		},
		"Check √5_2^x"
	);
	test(
		`(_5^2)√5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "prescriptSubsup",
				base: {
					type: "sqrtLiteral",
					value: {
						NumberLiteral: "5",
					},
				},
				down: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "5",
					},
				},
				up: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "2",
					},
				},
			},
		},
		"Check (_5^2)√5"
	);
	test(
		`√5┴exp1`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expAbove",
					base: [
						{
							NumberLiteral: "5",
						},
					],
					up: {
						type: "soperandLiteral",
						operand: [
							{
								CharLiteral: "exp",
							},
							{
								NumberLiteral: "1",
							},
						],
					},
				},
			},
		},
		"Check √5┴exp1"
	);
	test(
		`√5┬exp1`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expBelow",
					base: [
						{
							NumberLiteral: "5",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: [
							{
								CharLiteral: "exp",
							},
							{
								NumberLiteral: "1",
							},
						],
					},
				},
			},
		},
		"Check √5┬exp1"
	);
	test(
		`(√5┬exp1]`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expBracketLiteral",
				exp: {
					type: "sqrtLiteral",
					value: {
						type: "expBelow",
						base: [
							{
								NumberLiteral: "5",
							},
						],
						down: {
							type: "soperandLiteral",
							operand: [
								{
									CharLiteral: "exp",
								},
								{
									NumberLiteral: "1",
								},
							],
						},
					},
				},
				open: "(",
				close: "]",
			},
		},
		"Check (√5┬exp1]"
	);
	test(
		`□√5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "boxLiteral",
				value: {
					type: "sqrtLiteral",
					value: {
						NumberLiteral: "5",
					},
				},
			},
		},
		"Check □√5"
	);
	test(
		`▭√5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "rectLiteral",
				value: {
					type: "sqrtLiteral",
					value: {
						NumberLiteral: "5",
					},
				},
			},
		},
		"Check ▭√5"
	);
	test(
		`▁√5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "underbarLiteral",
				value: {
					type: "sqrtLiteral",
					value: {
						NumberLiteral: "5",
					},
				},
			},
		},
		"Check ▁√5"
	);
	test(
		` ̄√5`.trim(),
		{
			type: "UnicodeEquation",
			body: {
				type: "overbarLiteral",
				value: {
					type: "sqrtLiteral",
					value: {
						NumberLiteral: "5",
					},
				},
			},
		},
		"Check ̄√5"
	);
	test(
		`∑_√5^√5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: {
					type: "opNary",
					value: "∑",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						type: "sqrtLiteral",
						value: {
							type: "expSuperscript",
							base: [
								{
									NumberLiteral: "5",
								},
							],
							up: {
								type: "soperandLiteral",
								operand: {
									type: "sqrtLiteral",
									value: {
										NumberLiteral: "5",
									},
								},
							},
						},
					},
				},
			},
		},
		"Check ∑_√5^√5"
	);
	// test(
	// 	`\\root n+1\\of(b+c)+x`,
	// 	{},
	// 	"Check \\root n+1\\of(b+c)+x"
	// );
}
window["AscMath"].sqrt = sqrtTests;

function bracketsTests(test) {
	test(
		`(1+2)+2`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "1",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "2",
						},
					],
					open: "(",
					close: ")",
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "2",
				},
			],
		},
		"Проверка скобок: (1+2)+2"
	);
	test(
		`{1+2}-X`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "1",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "2",
						},
					],
					open: "{",
					close: "}",
				},
				{
					Operator: "-",
				},
				{
					CharLiteral: "X",
				},
			],
		},
		"Проверка скобок: {1+2}-X"
	);
	test(
		`[1+2]*i`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "1",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "2",
						},
					],
					open: "[",
					close: "]",
				},
				{
					Operator: "*",
				},
				{
					CharLiteral: "i",
				},
			],
		},
		"Проверка скобок: [1+2]*i"
	);
	test(
		`|1+2|-89/2`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "1",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "2",
						},
					],
					open: "|",
					close: "|",
				},
				{
					Operator: "-",
				},
				{
					type: "fractionLiteral",
					up: {
						NumberLiteral: "89",
					},
					opOver: "/",
					down: {
						NumberLiteral: "2",
					},
				},
			],
		},
		"Проверка скобок: |1+2|-89/2"
	);
	test(
		`|1+2|-〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "1",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "2",
						},
					],
					open: "|",
					close: "|",
				},
				{
					Operator: "-",
				},
				{
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
			],
		},
		"Проверка скобок: |1+2|-〖89/2〗"
	);
	test(
		`⌈1+2⌉-〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "1",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "2",
						},
					],
					open: "⌈",
					close: "⌉",
				},
				{
					Operator: "-",
				},
				{
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
			],
		},
		"Проверка скобок: ⌈1+2⌉-〖89/2〗"
	);
	test(
		`⌊1+2⌋-〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "1",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "2",
						},
					],
					open: "⌊",
					close: "⌋",
				},
				{
					Operator: "-",
				},
				{
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
			],
		},
		"Проверка скобок: ⌊1+2⌋-〖89/2〗"
	);
	test(
		`〖89/2〗/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка скобок:〖89/2〗/2"
	);
	test(
		`√〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
			},
		},
		"Проверка скобок:√〖89/2〗"
	);
	test(
		`〖89/2〗_2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: {
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "2",
					},
				},
			},
		},
		"Проверка скобок:〖89/2〗_2"
	);
	test(
		`〖89/2〗^2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSuperscript",
				base: {
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
				up: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "2",
					},
				},
			},
		},
		"Проверка скобок:〖89/2〗^2"
	);
	test(
		`2_〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: {
							type: "fractionLiteral",
							up: {
								NumberLiteral: "89",
							},
							opOver: "/",
							down: {
								NumberLiteral: "2",
							},
						},
						open: "〖",
						close: "〗",
					},
				},
			},
		},
		"Проверка скобок: 2_〖89/2〗"
	);
	test(
		`2^〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSuperscript",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				up: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: {
							type: "fractionLiteral",
							up: {
								NumberLiteral: "89",
							},
							opOver: "/",
							down: {
								NumberLiteral: "2",
							},
						},
						open: "〖",
						close: "〗",
					},
				},
			},
		},
		"Проверка скобок: 2^〖89/2〗"
	);

	test(
		`2_〖89/2〗_2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expSubscript",
						base: {
							type: "expBracketLiteral",
							exp: {
								type: "fractionLiteral",
								up: {
									NumberLiteral: "89",
								},
								opOver: "/",
								down: {
									NumberLiteral: "2",
								},
							},
							open: "〖",
							close: "〗",
						},
						down: {
							type: "soperandLiteral",
							operand: {
								NumberLiteral: "2",
							},
						},
					},
				},
			},
		},
		"Проверка скобок: 2_〖89/2〗_2"
	);
	test(
		`2^〖89/2〗^2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSuperscript",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				up: {
					type: "soperandLiteral",
					operand: {
						type: "expSuperscript",
						base: {
							type: "expBracketLiteral",
							exp: {
								type: "fractionLiteral",
								up: {
									NumberLiteral: "89",
								},
								opOver: "/",
								down: {
									NumberLiteral: "2",
								},
							},
							open: "〖",
							close: "〗",
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
		},
		"Проверка скобок: 2^〖89/2〗^2"
	);

	test(
		`2┴〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expAbove",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				up: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: {
							type: "fractionLiteral",
							up: {
								NumberLiteral: "89",
							},
							opOver: "/",
							down: {
								NumberLiteral: "2",
							},
						},
						open: "〖",
						close: "〗",
					},
				},
			},
		},
		"Проверка скобок: 2┴〖89/2〗"
	);
	test(
		`2┴〖89/2〗┴2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expAbove",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				up: {
					type: "soperandLiteral",
					operand: {
						type: "expAbove",
						base: {
							type: "expBracketLiteral",
							exp: {
								type: "fractionLiteral",
								up: {
									NumberLiteral: "89",
								},
								opOver: "/",
								down: {
									NumberLiteral: "2",
								},
							},
							open: "〖",
							close: "〗",
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
		},
		"Проверка скобок: 2┴〖89/2〗┴2"
	);
	test(
		`2┬〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expBelow",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: {
							type: "fractionLiteral",
							up: {
								NumberLiteral: "89",
							},
							opOver: "/",
							down: {
								NumberLiteral: "2",
							},
						},
						open: "〖",
						close: "〗",
					},
				},
			},
		},
		"Проверка скобок: 2┬〖89/2〗"
	);
	test(
		`2┬〖89/2〗┬2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expBelow",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expBelow",
						base: {
							type: "expBracketLiteral",
							exp: {
								type: "fractionLiteral",
								up: {
									NumberLiteral: "89",
								},
								opOver: "/",
								down: {
									NumberLiteral: "2",
								},
							},
							open: "〖",
							close: "〗",
						},
						down: {
							type: "soperandLiteral",
							operand: {
								NumberLiteral: "2",
							},
						},
					},
				},
			},
		},
		"Проверка скобок: 2┬〖89/2〗┬2"
	);
	test(
		"├]a+b┤[",
		{
			type: "UnicodeEquation",
			body: {
				type: "expBracketLiteral",
				open: "]",
				close: "[",
				value: [
					{
						type: "CharLiteral",
						value: "a"
					},
					{
						type: "OperatorLiteral",
						value: "+"
					},
					{
						type: "CharLiteral",
						value: "b"
					}
				]
			}
		},
		"Check: ├]a+b┤["
	)
}

window["AscMath"].bracket = bracketsTests;

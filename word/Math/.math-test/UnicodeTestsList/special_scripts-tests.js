function specialTest(test) {
	test(
		`2⁰¹²³⁴⁵⁶⁷⁸⁹`,
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
						NumberLiteral: "0123456789",
					},
				},
			},
		},
		"Проверка `2⁰¹²³⁴⁵⁶⁷⁸⁹`"
	);
	test(
		`2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹`,
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
					operand: [
						[
							{
								NumberLiteral: "4",
							},
							{
								CharLiteral: "in",
							},
						],
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "5",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "6",
								},
								{
									Operator: "+",
								},
								{
									NumberLiteral: "7",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "8",
								},
							],
							open: "(",
							close: ")",
						},
						{
							NumberLiteral: "9",
						},
					],
				},
			},
		},
		"Проверка `2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹`"
	);
	test(
		`2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSuperscript",
					base: [
						{
							NumberLiteral: "2",
						},
					],
					up: {
						type: "soperandLiteral",
						operand: [
							[
								{
									NumberLiteral: "4",
								},
								{
									CharLiteral: "in",
								},
							],
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "5",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "6",
									},
									{
										Operator: "+",
									},
									{
										NumberLiteral: "7",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "8",
									},
								],
								open: "(",
								close: ")",
							},
							{
								NumberLiteral: "9",
							},
						],
					},
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "45",
				},
			],
		},
		"Проверка `2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`"
	);
	test(
		`x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "x",
					},
					up: {
						type: "soperandLiteral",
						operand: [
							[
								{
									NumberLiteral: "4",
								},
								{
									CharLiteral: "in",
								},
							],
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "5",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "6",
									},
									{
										Operator: "+",
									},
									{
										NumberLiteral: "7",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "8",
									},
								],
								open: "(",
								close: ")",
							},
							{
								NumberLiteral: "9",
							},
						],
					},
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "45",
				},
			],
		},
		"Проверка `x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`"
	);

	test(
		`2₂₃₄₊₍₆₇₋₀₌₆₇₎56`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSubscript",
					base: [
						{
							NumberLiteral: "2",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: [
							{
								NumberLiteral: "234",
							},
							{
								Operator: "+",
							},
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "67",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "0",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "67",
									},
								],
								open: "(",
								close: ")",
							},
						],
					},
				},
				{
					NumberLiteral: "56",
				},
			],
		},
		"Проверка `2₂₃₄₊₍₆₇₋₀₌₆₇₎56`"
	);
	test(
		`z₂₃₄₊₍₆₇₋₀₌₆₇₎56`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSubscript",
					base: {
						CharLiteral: "z",
					},
					down: {
						type: "soperandLiteral",
						operand: [
							{
								NumberLiteral: "234",
							},
							{
								Operator: "+",
							},
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "67",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "0",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "67",
									},
								],
								open: "(",
								close: ")",
							},
						],
					},
				},
				{
					NumberLiteral: "56",
				},
			],
		},
		"Проверка `z₂₃₄₊₍₆₇₋₀₌₆₇₎56`"
	);

	test(
		`2⁰¹²³⁴⁵⁶⁷⁸⁹₂₃₄₊₍₆₇₋₀₌₆₇₎`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubsup",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: [
						{
							NumberLiteral: "234",
						},
						{
							Operator: "+",
						},
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "67",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "0",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "67",
								},
							],
							open: "(",
							close: ")",
						},
					],
				},
				up: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "0123456789",
					},
				},
			},
		},
		"Проверка `2⁰¹²³⁴⁵⁶⁷⁸⁹₂₃₄₊₍₆₇₋₀₌₆₇₎`"
	);
	test(
		`2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubsup",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: [
						{
							NumberLiteral: "234",
						},
						{
							Operator: "+",
						},
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "67",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "0",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "67",
								},
							],
							open: "(",
							close: ")",
						},
					],
				},
				up: {
					type: "soperandLiteral",
					operand: [
						[
							{
								NumberLiteral: "4",
							},
							{
								CharLiteral: "in",
							},
						],
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "5",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "6",
								},
								{
									Operator: "+",
								},
								{
									NumberLiteral: "7",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "8",
								},
							],
							open: "(",
							close: ")",
						},
						{
							NumberLiteral: "9",
						},
					],
				},
			},
		},
		"Проверка `2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎`"
	);
	test(
		`2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSubsup",
					base: [
						{
							NumberLiteral: "2",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: [
							{
								NumberLiteral: "234",
							},
							{
								Operator: "+",
							},
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "67",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "0",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "67",
									},
								],
								open: "(",
								close: ")",
							},
						],
					},
					up: {
						type: "soperandLiteral",
						operand: [
							[
								{
									NumberLiteral: "4",
								},
								{
									CharLiteral: "in",
								},
							],
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "5",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "6",
									},
									{
										Operator: "+",
									},
									{
										NumberLiteral: "7",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "8",
									},
								],
								open: "(",
								close: ")",
							},
							{
								NumberLiteral: "9",
							},
						],
					},
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "45",
				},
			],
		},
		"Проверка `2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45`"
	);
	test(
		`x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSubsup",
					base: {
						CharLiteral: "x",
					},
					down: {
						type: "soperandLiteral",
						operand: [
							{
								NumberLiteral: "234",
							},
							{
								Operator: "+",
							},
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "67",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "0",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "67",
									},
								],
								open: "(",
								close: ")",
							},
						],
					},
					up: {
						type: "soperandLiteral",
						operand: [
							[
								{
									NumberLiteral: "4",
								},
								{
									CharLiteral: "in",
								},
							],
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "5",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "6",
									},
									{
										Operator: "+",
									},
									{
										NumberLiteral: "7",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "8",
									},
								],
								open: "(",
								close: ")",
							},
							{
								NumberLiteral: "9",
							},
						],
					},
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "45",
				},
			],
		},
		"Проверка `x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45`"
	);

	test(
		`2₂₃₄₊₍₆₇₋₀₌₆₇₎⁰¹²³⁴⁵⁶⁷⁸⁹`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubsup",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: [
						{
							NumberLiteral: "234",
						},
						{
							Operator: "+",
						},
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "67",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "0",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "67",
								},
							],
							open: "(",
							close: ")",
						},
					],
				},
				up: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "0123456789",
					},
				},
			},
		},
		"Проверка `2₂₃₄₊₍₆₇₋₀₌₆₇₎⁰¹²³⁴⁵⁶⁷⁸⁹`"
	);
	test(
		`2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubsup",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: [
						{
							NumberLiteral: "234",
						},
						{
							Operator: "+",
						},
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "67",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "0",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "67",
								},
							],
							open: "(",
							close: ")",
						},
					],
				},
				up: {
					type: "soperandLiteral",
					operand: [
						[
							{
								NumberLiteral: "4",
							},
							{
								CharLiteral: "in",
							},
						],
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "5",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "6",
								},
								{
									Operator: "+",
								},
								{
									NumberLiteral: "7",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "8",
								},
							],
							open: "(",
							close: ")",
						},
						{
							NumberLiteral: "9",
						},
					],
				},
			},
		},
		"Проверка `2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹`"
	);
	test(
		`2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSubsup",
					base: [
						{
							NumberLiteral: "2",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: [
							{
								NumberLiteral: "234",
							},
							{
								Operator: "+",
							},
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "67",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "0",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "67",
									},
								],
								open: "(",
								close: ")",
							},
						],
					},
					up: {
						type: "soperandLiteral",
						operand: [
							[
								{
									NumberLiteral: "4",
								},
								{
									CharLiteral: "in",
								},
							],
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "5",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "6",
									},
									{
										Operator: "+",
									},
									{
										NumberLiteral: "7",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "8",
									},
								],
								open: "(",
								close: ")",
							},
							{
								NumberLiteral: "9",
							},
						],
					},
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "45",
				},
			],
		},
		"Проверка `2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`"
	);
	test(
		`x₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSubsup",
					base: {
						CharLiteral: "x",
					},
					down: {
						type: "soperandLiteral",
						operand: [
							{
								NumberLiteral: "234",
							},
							{
								Operator: "+",
							},
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "67",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "0",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "67",
									},
								],
								open: "(",
								close: ")",
							},
						],
					},
					up: {
						type: "soperandLiteral",
						operand: [
							[
								{
									NumberLiteral: "4",
								},
								{
									CharLiteral: "in",
								},
							],
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "5",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "6",
									},
									{
										Operator: "+",
									},
									{
										NumberLiteral: "7",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "8",
									},
								],
								open: "(",
								close: ")",
							},
							{
								NumberLiteral: "9",
							},
						],
					},
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "45",
				},
			],
		},
		"Проверка `x₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`"
	);
}
window["AscMath"].special = specialTest;

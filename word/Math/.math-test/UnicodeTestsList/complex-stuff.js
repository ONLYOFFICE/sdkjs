function complexTest(test) {
	test(
		`(a + b)^n = ∑_(k=0)^n▒(n¦k) a^k b^(n-k),`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSuperscript",
					base: {
						type: "expBracketLiteral",
						exp: [
							{
								CharLiteral: "a",
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
								CharLiteral: "b",
							},
						],
						open: "(",
						close: ")",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "n",
						},
					},
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					Operator: "=",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "expSubsup",
					base: {
						type: "opNary",
						value: "∑",
					},
					down: {
						type: "soperandLiteral",
						operand: {
							type: "expBracketLiteral",
							exp: [
								{
									CharLiteral: "k",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "0",
								},
							],
							open: "(",
							close: ")",
						},
					},
					up: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "n",
						},
					},
					thirdSoperand: {
						type: "soperandLiteral",
						operand: {
							type: "expBracketLiteral",
							exp: {
								type: "binomLiteral",
								numerator: {
									CharLiteral: "n",
								},
								operand: {
									CharLiteral: "k",
								},
							},
							open: "(",
							close: ")",
						},
					},
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "a",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "k",
						},
					},
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "b",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							type: "expBracketLiteral",
							exp: [
								{
									CharLiteral: "n",
								},
								{
									Operator: "-",
								},
								{
									CharLiteral: "k",
								},
							],
							open: "(",
							close: ")",
						},
					},
				},
			],
		},
		"Проверка простого литерала: (a + b)^n = ∑_(k=0)^n▒(n¦k) a^k b^(n-k),"
	);
	test(
		`∑_2^2▒(n/23)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubsup",
				base: {
					type: "opNary",
					value: "∑",
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
						NumberLiteral: "2",
					},
				},
				thirdSoperand: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: {
							type: "fractionLiteral",
							up: {
								CharLiteral: "n",
							},
							opOver: "/",
							down: {
								NumberLiteral: "23",
							},
						},
						open: "(",
						close: ")",
					},
				},
			},
		},
		"Проверка простого литерала: ∑_2^2▒(n/23)"
	);
	test(
		`⏞(x+⋯+x)^(k "times")`,
		{
			type: "UnicodeEquation",
			body: {
				type: "hbrackLiteral",
				operand: {
					type: "expBracketLiteral",
					exp: [
						{
							CharLiteral: "x",
						},
						{
							Operator: "+",
						},
						{
							CharLiteral: "⋯",
						},
						{
							Operator: "+",
						},
						{
							CharLiteral: "x",
						},
					],
					open: "(",
					close: ")",
				},
				up: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: [
							{
								CharLiteral: "k",
							},
							{
								type: "SpaceLiteral",
								value: " ",
							},
							{
								CharLiteral: '"times"',
							},
						],
						open: "(",
						close: ")",
					},
				},
			},
		},
		"Проверка простого литерала: ⏞(x+⋯+x)^(k 'times')"
	);
	test(
		`𝐸 = 𝑚𝑐^2`,
		{
			type: "UnicodeEquation",
			body: [
				{
					CharLiteral: "𝐸",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					Operator: "=",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "𝑚𝑐",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
			],
		},
		"Проверка простого литерала: 𝐸 = 𝑚𝑐^2"
	);
	test(
		`∫_0^a▒xⅆx/(x^2+a^2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expSubsup",
					base: {
						type: "opNary",
						value: "∫",
					},
					down: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "0",
						},
					},
					up: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "a",
						},
					},
					thirdSoperand: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "xⅆx",
						},
					},
				},
				opOver: "/",
				down: {
					type: "expBracketLiteral",
					exp: [
						{
							type: "expSuperscript",
							base: {
								CharLiteral: "x",
							},
							up: {
								type: "soperandLiteral",
								operand: {
									NumberLiteral: "2",
								},
							},
						},
						{
							Operator: "+",
						},
						{
							type: "expSuperscript",
							base: {
								CharLiteral: "a",
							},
							up: {
								type: "soperandLiteral",
								operand: {
									NumberLiteral: "2",
								},
							},
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка простого литерала: ∫_0^a▒xⅆx/(x^2+a^2)"
	);
	test(
		`lim┬(n→∞) a_n`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBelow",
					base: {
						CharLiteral: "lim",
					},
					down: {
						type: "soperandLiteral",
						operand: {
							type: "expBracketLiteral",
							exp: [
								{
									CharLiteral: "n",
								},
								{
									Operator: "→",
								},
								{
									Operator: "∞",
								},
							],
							open: "(",
							close: ")",
						},
					},
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "expSubscript",
					base: {
						CharLiteral: "a",
					},
					down: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "n",
						},
					},
				},
			],
		},
		"Проверка простого литерала: lim┬(n→∞) a_n"
	);
	test(
		`ⅈ²=-1`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "ⅈ",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
				{
					Operator: "=",
				},
				{
					Operator: "-",
				},
				{
					NumberLiteral: "1",
				},
			],
		},
		"Проверка простого литерала: ⅈ²=-1"
	);
	test(
		`E = m⁢c²`,
		{
			type: "UnicodeEquation",
			body: [
				{
					CharLiteral: "E",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					Operator: "=",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "m⁢c",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
			],
		},
		"Проверка простого литерала: E = m⁢c²"
	);
	test(
		`a²⋅b²=c²`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "a",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
				{
					Operator: "⋅",
				},
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "b",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
				{
					Operator: "=",
				},
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "c",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
			],
		},
		"Проверка простого литерала: a²⋅b²=c²"
	);
	test(
		`f̂(ξ)=∫_-∞^∞▒f(x)ⅇ^-2πⅈxξ ⅆx`,
		{
			type: "UnicodeEquation",
			body: [
				[
					{
						CharLiteral: "f̂",
					},
					{
						type: "expBracketLiteral",
						exp: {
							type: "anOther",
							value: "ξ",
						},
						open: "(",
						close: ")",
					},
				],
				{
					Operator: "=",
				},
				{
					type: "expSubsup",
					base: {
						type: "opNary",
						value: "∫",
					},
					down: {
						type: "soperandLiteral",
						operand: "-∞",
					},
					up: {
						type: "soperandLiteral",
						operand: "∞",
					},
					thirdSoperand: {
						type: "soperandLiteral",
						operand: [
							{
								CharLiteral: "f",
							},
							{
								type: "expBracketLiteral",
								exp: {
									CharLiteral: "x",
								},
								open: "(",
								close: ")",
							},
							{
								type: "expSuperscript",
								base: {
									CharLiteral: "ⅇ",
								},
								up: {
									type: "soperandLiteral",
									operand: [
										{
											NumberLiteral: "2",
										},
										{
											type: "anOther",
											value: "π",
										},
										{
											CharLiteral: "ⅈxξ",
										},
									],
									minus: true,
								},
							},
						],
					},
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					CharLiteral: "ⅆx",
				},
			],
		},
		"Проверка простого литерала: f̂(ξ)=∫_-∞^∞▒f(x)ⅇ^-2πⅈxξ ⅆx"
	);
	test(
		`(𝑎 + 𝑏)┴→`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expAbove",
				base: {
					type: "expBracketLiteral",
					exp: [
						{
							CharLiteral: "𝑎",
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
							CharLiteral: "𝑏",
						},
					],
					open: "(",
					close: ")",
				},
				up: {
					Operator: "→",
				},
			},
		},
		"Проверка простого литерала: (𝑎 + 𝑏)┴→"
	);
	test(
		`𝑎┴→`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expAbove",
				base: {
					CharLiteral: "𝑎",
				},
				up: {
					Operator: "→",
				},
			},
		},
		"Проверка простого литерала: 𝑎┴→"
	);
}
window["AscMath"].complex = complexTest;

function hbrackTests(test) {
	test(
		`⏞(x+⋯+x)`,
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
			},
		},
		"Проверка работы hBrack: ⏞(x+⋯+x)"
	);

	test(
		`⏞(x+⋯+x)^2`,
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
						NumberLiteral: "2",
					},
				},
			},
		},
		"Проверка работы hBrack: ⏞(x+⋯+x)^2"
	);
	test(
		`⏞(x+⋯+x)_2`,
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
				down: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "2",
					},
				},
			},
		},
		"Проверка работы hBrack: ⏞(x+⋯+x)_2"
	);
	test(
		`⏞(x+⋯+x)_2^Y`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSuperscript",
				base: {
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
					down: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
				up: {
					type: "soperandLiteral",
					operand: {
						CharLiteral: "Y",
					},
				},
			},
		},
		"Проверка работы hBrack: ⏞(x+⋯+x)_2^Y"
	);

	test(
		`⏞(x+⋯+x)_2^2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSuperscript",
				base: {
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
					down: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
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
		"Проверка работы hBrack: ⏞(x+⋯+x)_2^2" // WORD воспринимает hBrack как обычный элемент!
	);
}
window["AscMath"].hbrack = hbrackTests;

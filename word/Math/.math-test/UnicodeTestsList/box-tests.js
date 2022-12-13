function boxTests(test) {
	test(
		`□(1+2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "boxLiteral",
				value: {
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
			},
		},
		"Проверка box: □(1+2)"
	);
	test(
		`□(1+2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "boxLiteral",
				value: {
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
			},
		},
		"Проверка box: □(1+2)"
	);
	test(
		`□1`,
		{
			type: "UnicodeEquation",
			body: {
				type: "boxLiteral",
				value: {
					NumberLiteral: "1",
				},
			},
		},
		"Проверка box: □1"
	);
	test(
		`□1/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "boxLiteral",
					value: {
						NumberLiteral: "1",
					},
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка box: □1/2"
	);
	test(
		`▭(1+2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "rectLiteral",
				value: {
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
			},
		},
		"Проверка rect: ▭(1+2)"
	);
	test(
		`▭1`,
		{
			type: "UnicodeEquation",
			body: {
				type: "rectLiteral",
				value: {
					NumberLiteral: "1",
				},
			},
		},
		"Проверка rect: ▭1"
	);
	test(
		`▭1/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "rectLiteral",
					value: {
						NumberLiteral: "1",
					},
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка rect: ▭1/2"
	);
	test(
		`▁(1+2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "underbarLiteral",
				value: {
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
			},
		},
		"Проверка underbar: ▁(1+2)"
	);

	test(
		`▁1`,
		{
			type: "UnicodeEquation",
			body: {
				type: "underbarLiteral",
				value: {
					NumberLiteral: "1",
				},
			},
		},
		"Проверка underbar: ▁1"
	);
	test(
		`▁1/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "underbarLiteral",
					value: {
						NumberLiteral: "1",
					},
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка underbar: ▁1/2"
	);
	test(
		` ̄(1+2)`.trim(),
		{
			type: "UnicodeEquation",
			body: {
				type: "overbarLiteral",
				value: {
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
			},
		},
		"Проверка underbar:  ̄(1+2)"
	);

	test(
		` ̄1`.trim(),
		{
			type: "UnicodeEquation",
			body: {
				type: "overbarLiteral",
				value: {
					NumberLiteral: "1",
				},
			},
		},
		"Проверка underbar:  ̄1"
	);
	test(
		`(1+2)̂`.trim(),
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "DiacriticLiteral",
				"base": {
					"type": "expBracketLiteral",
					"exp": [
						[
							{
								"type": "NumberLiteral",
								"value": "1"
							}
						],
						{
							"type": "OperatorLiteral",
							"value": "+"
						},
						[
							{
								"type": "NumberLiteral",
								"value": "2"
							}
						]
					],
					"open": "(",
					"close": ")"
				},
				"value": "̂"
			}
		},
		"Проверка underbar: (1+2)̂"
	);
}

window["AscMath"].box = boxTests;


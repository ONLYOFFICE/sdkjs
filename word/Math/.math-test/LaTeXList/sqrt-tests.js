function sqrt(test) {
	test(
		"\\sqrt5",
		{
			"body": {
				"type": "SqrtLiteral",
				"value": {
					"type": "NumberLiteral",
					"value": "5"
				}
			},
			"type": "LaTeXEquation"
		},
		"Check \\sqrt5"
	);
	test(
		"\\sqrt\\frac{1}{2}",
		{
			"body": {
				"type": "SqrtLiteral",
				"value": {
					"down": {
						"type": "NumberLiteral",
						"value": "2"
					},
					"type": "FractionLiteral",
					"up": {
						"type": "NumberLiteral",
						"value": "1"
					}
				}
			},
			"type": "LaTeXEquation"
		},
		"Check \\sqrt\\frac{1}{2}"
	);
	test(
		"\\sqrt[2^2]\\frac{1}{2}",
		{
			"body": {
				"index": {
					"type": "SubSupLiteral",
					"up": {
						"type": "NumberLiteral",
						"value": "2"
					},
					"value": {
						"type": "NumberLiteral",
						"value": "2"
					}
				},
				"type": "SqrtLiteral",
				"value": {
					"down": {
						"type": "NumberLiteral",
						"value": "2"
					},
					"type": "FractionLiteral",
					"up": {
						"type": "NumberLiteral",
						"value": "1"
					}
				}
			},
			"type": "LaTeXEquation"
		},
		"Check \\sqrt[2^2]\\frac{1}{2}"
	);
	test(
		"\\sqrt[2^2] {\\frac{1}{2}+3}",
		{
			"body": {
				"index": {
					"type": "SubSupLiteral",
					"up": {
						"type": "NumberLiteral",
						"value": "2"
					},
					"value": {
						"type": "NumberLiteral",
						"value": "2"
					}
				},
				"type": "SqrtLiteral",
				"value": [
					{
						"down": {
							"type": "NumberLiteral",
							"value": "2"
						},
						"type": "FractionLiteral",
						"up": {
							"type": "NumberLiteral",
							"value": "1"
						}
					},
					{
						"type": "OperatorLiteral",
						"value": "+"
					},
					{
						"type": "NumberLiteral",
						"value": "3"
					}
				]
			},
			"type": "LaTeXEquation"
		},
		"Check \\sqrt[2^2] {\\frac{1}{2}+3}"
	);
}
window["AscMath"].sqrt = sqrt;

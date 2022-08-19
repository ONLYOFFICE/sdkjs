function fraction(test) {
	test(
		`\\frac{1}{2}`,
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "FractionLiteral",
				"up": {
					"type": "NumberLiteral",
					"value": "1"
				},
				"down": {
					"type": "NumberLiteral",
					"value": "2"
				}
			}
		},
		"Check \\frac{1}{2}"
	);
	test(
		`\\frac{1+\\frac{x}{y}}{2}`,
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "FractionLiteral",
				"up": [
					{
						"type": "NumberLiteral",
						"value": "1"
					},
					{
						"type": "OperatorLiteral",
						"value": "+"
					},
					{
						"type": "FractionLiteral",
						"up": {
							"type": "CharLiteral",
							"value": "x"
						},
						"down": {
							"type": "CharLiteral",
							"value": "y"
						}
					}
				],
				"down": {
					"type": "NumberLiteral",
					"value": "2"
				}
			}
		},
		"Check \\frac{1+\\frac{x}{y}}{2}"
	);
	test(
		`\\frac{1^x}{2_y}`,
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "FractionLiteral",
				"up": {
					"type": "SubSupLiteral",
					"value": {
						"type": "NumberLiteral",
						"value": "1"
					},
					"up": {
						"type": "CharLiteral",
						"value": "x"
					}
				},
				"down": {
					"type": "SubSupLiteral",
					"value": {
						"type": "NumberLiteral",
						"value": "2"
					},
					"down": {
						"type": "CharLiteral",
						"value": "y"
					}
				}
			}
		},
		"Check \\frac{1^x}{2_y}"
	);
	test(
		`\\sum^{2}_{x}4`,
		{
			"body": {
				"down": {
					"type": "CharLiteral",
					"value": "x"
				},
				"third": {
					"type": "NumberLiteral",
					"value": "4"
				},
				"type": "SubSupLiteral",
				"up": {
					"type": "NumberLiteral",
					"value": "2"
				},
				"value": {
					"type": "opNaryLiteral",
					"value": "∑"
				}
			},
			"type": "LaTeXEquation"
		},
		"Check \\sum^{2}_{x}4"
	);
	test(
		`\\int^2_x{4}`,
		{
			"body": {
				"down": {
					"type": "CharLiteral",
					"value": "x"
				},
				"third": {
					"type": "NumberLiteral",
					"value": "4"
				},
				"type": "SubSupLiteral",
				"up": {
					"type": "NumberLiteral",
					"value": "2"
				},
				"value": {
					"type": "opNaryLiteral",
					"value": "∫"
				}
			},
			"type": "LaTeXEquation"
		},
		"Check \\int^2_x{4}"
	);
	test(
		`\\binom{1}{2}`,
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "BinomLiteral",
				"up": {
					"type": "NumberLiteral",
					"value": "1"
				},
				"down": {
					"type": "NumberLiteral",
					"value": "2"
				}
			}
		},
		"Check \\binom{1}{2}"
	);
	test(
		`\\sum_{i=1}^{10} t_i`,
		{
			"body": {
				"down": [
					{
						"type": "CharLiteral",
						"value": "i"
					},
					{
						"type": "OperatorLiteral",
						"value": "="
					},
					{
						"type": "NumberLiteral",
						"value": "1"
					}
				],
				"third": {
					"down": {
						"type": "CharLiteral",
						"value": "i"
					},
					"type": "SubSupLiteral",
					"value": {
						"type": "CharLiteral",
						"value": "t"
					}
				},
				"type": "SubSupLiteral",
				"up": {
					"type": "NumberLiteral",
					"value": "10"
				},
				"value": {
					"type": "opNaryLiteral",
					"value": "∑"
				}
			},
			"type": "LaTeXEquation"
		},
		"Check \\sum_{i=1}^{10} t_i"
	);
}

window["AscMath"].fraction = fraction;

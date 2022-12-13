function accents(test) {
	test(
		"\\dot{a}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "a"
				},
				"value": "̇"
			}
		},
		"Check \\dot{a}"
	);
	test(
		"\\ddot{b}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "b"
				},
				"value": "̈"
			}
		},
		"Check \\ddot{b}"
	);
	test(
		"\\acute{c}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "c"
				},
				"value": "́"
			}
		},
		"Check \\acute{c}"
	);
	test(
		"\\grave{d}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "d"
				},
				"value": "̀"
			}
		},
		"Check \\grave{d}"
	);
	test(
		"\\check{e}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "e"
				},
				"value": "̌"
			}
		},
		"Check \\check{e}"
	);
	test(
		"\\breve{f}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "f"
				},
				"value": "̆"
			}
		},
		"Check \\breve{f}"
	);
	test(
		"\\tilde{g}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "g"
				},
				"value": "̃"
			}
		},
		"Check \\tilde{g}"
	);
	test(
		"\\bar{h}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "h"
				},
				"value": "̅"
			}
		},
		"Check \\bar{h}"
	);
	test(
		"\\widehat{j}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "j"
				},
				"value": "̂"
			}
		},
		"Check \\widehat{j}"
	);
	test(
		"\\vec{k}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "k"
				},
				"value": "⃗"
			}
		}
		,
		"Check \\vec{k}"
	);
	//doesn't implement in word
	// test(
	// 	"\\not{l}",
	// 	{
	// 		"type": "LaTeXEquation",
	// 		"body": {
	// 			"type": "AccentLiteral",
	// 			"base": {
	// 				"type": "CharLiteral",
	// 				"value": "l"
	// 			},
	// 			"value": 824
	// 		}
	// 	},
	// 	"Check \\not{l}"
	// );
	test(
		"\\vec \\frac{k}{2}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "FractionLiteral",
					"up": {
						"type": "CharLiteral",
						"value": "k"
					},
					"down": {
						"type": "NumberLiteral",
						"value": "2"
					}
				},
				"value": "⃗"
			}
		},
		"Check \\vec \\frac{k}{2}"
	);
	// test(
	// 	"\\not\\notl2",
	// 	{
	// 		"type": "LaTeXEquation",
	// 		"body": [
	// 			{
	// 				"type": "AccentLiteral",
	// 				"base": {
	// 					"type": "AccentLiteral",
	// 					"base": {
	// 						"type": "CharLiteral",
	// 						"value": "l"
	// 					},
	// 					"value": 824
	// 				},
	// 				"value": 824
	// 			},
	// 			{
	// 				"type": "NumberLiteral",
	// 				"value": "2"
	// 			}
	// 		]
	// 	},
	// 	"Check \\notl"
	// );
	test(
		"5''",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "NumberLiteral",
					"value": "5"
				},
				"value": "''"
			}
		},
		"Check 5''"
	);
	test(
		"\\frac{4}{5}''",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "FractionLiteral",
					"up": {
						"type": "NumberLiteral",
						"value": "4"
					},
					"down": {
						"type": "NumberLiteral",
						"value": "5"
					}
				},
				"value": "''"
			}
		},
		"Check \\frac{4}{5}''"
	);
}
window["AscMath"].accents = accents;

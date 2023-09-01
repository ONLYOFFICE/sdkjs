/*
 * (c) Copyright Ascensio System SIA 2010-2023
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */
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

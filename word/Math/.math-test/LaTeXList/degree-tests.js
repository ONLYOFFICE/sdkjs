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
function degree(test) {
	test(
		"2^2",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "NumberLiteral",
					"value": "2"
				},
				"up": {
					"type": "NumberLiteral",
					"value": "2"
				}
			}
		},
		"Check 2^2"
	);
	test(
		"a^b",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "CharLiteral",
					"value": "a"
				},
				"up": {
					"type": "CharLiteral",
					"value": "b"
				}
			}
		},
		"Check a^b"
	);
	test(
		"a^2",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "CharLiteral",
					"value": "a"
				},
				"up": {
					"type": "NumberLiteral",
					"value": "2"
				}
			}
		},
		"Check a^2"
	);
	test(
		"2^b",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "NumberLiteral",
					"value": "2"
				},
				"up": {
					"type": "CharLiteral",
					"value": "b"
				}
			}
		},
		"Check 2^b"
	);
	test(
		"2_2",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "NumberLiteral",
					"value": "2"
				},
				"down": {
					"type": "NumberLiteral",
					"value": "2"
				}
			}
		},
		"Check 2_2"
	);
	test(
		"a_b",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "CharLiteral",
					"value": "a"
				},
				"down": {
					"type": "CharLiteral",
					"value": "b"
				}
			}
		},
		"Check a_b"
	);
	test(
		"a_2",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "CharLiteral",
					"value": "a"
				},
				"down": {
					"type": "NumberLiteral",
					"value": "2"
				}
			}
		},
		"Check a_2"
	);
	test(
		"2_b",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "NumberLiteral",
					"value": "2"
				},
				"down": {
					"type": "CharLiteral",
					"value": "b"
				}
			}
		},
		"Check 2_b"
	);
	test(
		`k_{n+1} = n^2 + k_n^2 - k_{n-1}`,
		{
			"type": "LaTeXEquation",
			"body": [
				{
					"type": "SubSupLiteral",
					"value": {
						"type": "CharLiteral",
						"value": "k"
					},
					"down": [
						{
							"type": "CharLiteral",
							"value": "n"
						},
						{
							"type": "OperatorLiteral",
							"value": "+"
						},
						{
							"type": "NumberLiteral",
							"value": "1"
						}
					]
				},
				{
					"type": "SpaceLiteral",
					"value": " "
				},
				{
					"type": "OperatorLiteral",
					"value": "="
				},
				{
					"type": "SpaceLiteral",
					"value": " "
				},
				{
					"type": "SubSupLiteral",
					"value": {
						"type": "CharLiteral",
						"value": "n"
					},
					"up": {
						"type": "NumberLiteral",
						"value": "2"
					}
				},
				{
					"type": "SpaceLiteral",
					"value": " "
				},
				{
					"type": "OperatorLiteral",
					"value": "+"
				},
				{
					"type": "SpaceLiteral",
					"value": " "
				},
				{
					"type": "SubSupLiteral",
					"value": {
						"type": "CharLiteral",
						"value": "k"
					},
					"up": {
						"type": "NumberLiteral",
						"value": "2"
					},
					"down": {
						"type": "CharLiteral",
						"value": "n"
					}
				},
				{
					"type": "SpaceLiteral",
					"value": " "
				},
				{
					"type": "OperatorLiteral",
					"value": "-"
				},
				{
					"type": "SpaceLiteral",
					"value": " "
				},
				{
					"type": "SubSupLiteral",
					"value": {
						"type": "CharLiteral",
						"value": "k"
					},
					"down": [
						{
							"type": "CharLiteral",
							"value": "n"
						},
						{
							"type": "OperatorLiteral",
							"value": "-"
						},
						{
							"type": "NumberLiteral",
							"value": "1"
						}
					]
				}
			]
		},
		"Check k_{n+1} = n^2 + k_n^2 - k_{n-1}"
	);
	test(
		`\\frac{1}{2}^{2}`,
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "FractionLiteral",
					"up": {
						"type": "NumberLiteral",
						"value": "1"
					},
					"down": {
						"type": "NumberLiteral",
						"value": "2"
					}
				},
				"up": {
					"type": "NumberLiteral",
					"value": "2"
				}
			}
		},
		"Check \\frac{1}{2}^{2}"
	);
	test(
		`\\frac{1}{2}_2`,
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "FractionLiteral",
					"up": {
						"type": "NumberLiteral",
						"value": "1"
					},
					"down": {
						"type": "NumberLiteral",
						"value": "2"
					}
				},
				"down": {
					"type": "NumberLiteral",
					"value": "2"
				}
			}
		},
		"Check \\frac{1}{2}_2"
	);
	test(
		`\\frac{1}{2}_2^y`,
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "FractionLiteral",
					"up": {
						"type": "NumberLiteral",
						"value": "1"
					},
					"down": {
						"type": "NumberLiteral",
						"value": "2"
					}
				},
				"up": {
					"type": "CharLiteral",
					"value": "y"
				},
				"down": {
					"type": "NumberLiteral",
					"value": "2"
				}
			}
		},
		"Check \\frac{1}{2}_2^y"
	);
	test(
		`\\frac{1}{2}_{2}^{y}`,
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "FractionLiteral",
					"up": {
						"type": "NumberLiteral",
						"value": "1"
					},
					"down": {
						"type": "NumberLiteral",
						"value": "2"
					}
				},
				"up": {
					"type": "CharLiteral",
					"value": "y"
				},
				"down": {
					"type": "NumberLiteral",
					"value": "2"
				}
			}
		},
		"Check \\frac{1}{2}_{2}^{y}"
	);
	test(
		`\\frac{1}{2}_1_2_3_4_5_6_7`,
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "FractionLiteral",
					"up": {
						"type": "NumberLiteral",
						"value": "1"
					},
					"down": {
						"type": "NumberLiteral",
						"value": "2"
					}
				},
				"down": {
					"type": "SubSupLiteral",
					"value": {
						"type": "NumberLiteral",
						"value": "1"
					},
					"down": {
						"type": "SubSupLiteral",
						"value": {
							"type": "NumberLiteral",
							"value": "2"
						},
						"down": {
							"type": "SubSupLiteral",
							"value": {
								"type": "NumberLiteral",
								"value": "3"
							},
							"down": {
								"type": "SubSupLiteral",
								"value": {
									"type": "NumberLiteral",
									"value": "4"
								},
								"down": {
									"type": "SubSupLiteral",
									"value": {
										"type": "NumberLiteral",
										"value": "5"
									},
									"down": {
										"type": "SubSupLiteral",
										"value": {
											"type": "NumberLiteral",
											"value": "6"
										},
										"down": {
											"type": "NumberLiteral",
											"value": "7"
										}
									}
								}
							}
						}
					}
				}
			}
		},
		"Check \\frac{1}{2}_1_2_3_4_5_6_7"
	);
	test(
		`\\frac{1}{2}^1^2^3^4^5^6^7`,
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "FractionLiteral",
					"up": {
						"type": "NumberLiteral",
						"value": "1"
					},
					"down": {
						"type": "NumberLiteral",
						"value": "2"
					}
				},
				"up": {
					"type": "SubSupLiteral",
					"value": {
						"type": "NumberLiteral",
						"value": "1"
					},
					"up": {
						"type": "SubSupLiteral",
						"value": {
							"type": "NumberLiteral",
							"value": "2"
						},
						"up": {
							"type": "SubSupLiteral",
							"value": {
								"type": "NumberLiteral",
								"value": "3"
							},
							"up": {
								"type": "SubSupLiteral",
								"value": {
									"type": "NumberLiteral",
									"value": "4"
								},
								"up": {
									"type": "SubSupLiteral",
									"value": {
										"type": "NumberLiteral",
										"value": "5"
									},
									"up": {
										"type": "SubSupLiteral",
										"value": {
											"type": "NumberLiteral",
											"value": "6"
										},
										"up": {
											"type": "NumberLiteral",
											"value": "7"
										}
									}
								}
							}
						}
					}
				}
			}
		},
		"Check \\frac{1}{2}^1^2^3^4^5^6^7"
	);
	test(
		`\\frac{1}{2}^1^2^3^4^5^6^7_x`,
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "SubSupLiteral",
				"value": {
					"type": "FractionLiteral",
					"up": {
						"type": "NumberLiteral",
						"value": "1"
					},
					"down": {
						"type": "NumberLiteral",
						"value": "2"
					}
				},
				"up": {
					"type": "SubSupLiteral",
					"value": {
						"type": "NumberLiteral",
						"value": "1"
					},
					"up": {
						"type": "SubSupLiteral",
						"value": {
							"type": "NumberLiteral",
							"value": "2"
						},
						"up": {
							"type": "SubSupLiteral",
							"value": {
								"type": "NumberLiteral",
								"value": "3"
							},
							"up": {
								"type": "SubSupLiteral",
								"value": {
									"type": "NumberLiteral",
									"value": "4"
								},
								"up": {
									"type": "SubSupLiteral",
									"value": {
										"type": "NumberLiteral",
										"value": "5"
									},
									"up": {
										"type": "SubSupLiteral",
										"value": {
											"type": "NumberLiteral",
											"value": "6"
										},
										"up": {
											"type": "NumberLiteral",
											"value": "7"
										}
									}
								}
							}
						}
					}
				},
				"down": {
					"type": "CharLiteral",
					"value": "x"
				}
			}
		},
		"Check \\frac{1}{2}^1^2^3^4^5^6^7_x"
	);
}

window["AscMath"].degree = degree;

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

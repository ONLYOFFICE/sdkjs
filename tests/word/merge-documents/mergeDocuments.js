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

// attention: There is a difference in merge in our editors and microsoft editors.
// Within one paragraph, pieces of text that are missing in the largest common document, we always add with a review like adding
// When merging, first we add the missing text from the second document, then from the first

QUnit.dump.maxDepth = 7;
const arrTestObjectsInfo = [
	///////////////////////// -> 1 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Привет')
			]
		]
	},
	///////////////////////// -> 2 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('')
			]
		]
	},
	///////////////////////// -> 3 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Привет')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Приветище')
			]
		]
	},
	///////////////////////// -> 4 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Привет')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Привет', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000})
			]
		]
	},
	///////////////////////// -> 5 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Привет')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Приветище', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000})
			]
		]
	},
	///////////////////////// -> 6 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Привет, как дела?')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('При', {
					reviewType: reviewtype_Add,
					userName  : 'John Smith',
					dateTime  : 1000000
				}), createParagraphInfo('вет, как дела?')
			]
		]
	},
	///////////////////////// -> 7 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Привет Привет Привет')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Привет'), createParagraphInfo(' Привет', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' Привет')
			]
		]
	},
	///////////////////////// -> 8 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Привет'), createParagraphInfo(' ой', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' Привет'), createParagraphInfo(' Привет')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Привет'), createParagraphInfo(' Привет', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' Привет')
			]
		]
	},
	///////////////////////// -> 9 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('как дела?')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Привет, ', {
					reviewType: reviewtype_Add,
					userName  : 'John Smith',
					dateTime  : 1000000
				}), createParagraphInfo('как дела?')
			]
		]
	},
	///////////////////////// -> 10 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Привет, как дела?')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Приветик, ', {
					reviewType: reviewtype_Add,
					userName  : 'John Smith',
					dateTime  : 1000000
				}), createParagraphInfo('как дела?')
			]
		]
	},
	///////////////////////// -> 11 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Привет, как дела?'), createParagraphInfo(' Нормально, а у тебя как?', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			})
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Привет, как дела?'), createParagraphInfo(' Хорошо, а у тебя как?', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 2000000
			})
			]
		]
	},
	///////////////////////// -> 12 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Привет, как дела? Нормально, а у тебя как?')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Привет, как дела?'), createParagraphInfo(' Хорошо, а у тебя как?', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			})
			]
		]
	},
	///////////////////////// -> 13 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Привет, как дела?   Хорошо, а у тебя как?')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Привет, как дела?    '), createParagraphInfo(' Нормально, а у тебя как?', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			})
			]
		]
	},
	///////////////////////// -> 14 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('При', {
					reviewType: reviewtype_Remove,
					userName  : 'John Smith',
					dateTime  : 1000000
				}), createParagraphInfo('вет, как д'), createParagraphInfo('ел', {
				reviewType: reviewtype_Remove,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('а?')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Пр'), createParagraphInfo('и', {
				reviewType: reviewtype_Add,
				userName  : 'John Smoth',
				dateTime  : 2000000
			}), createParagraphInfo('в'), createParagraphInfo('е', {
				reviewType: reviewtype_Add,
				userName  : 'John Smoth',
				dateTime  : 2000000
			}), createParagraphInfo('т'), createParagraphInfo(',', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' к'), createParagraphInfo('а', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('к', {
				reviewType: reviewtype_Remove,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' '), createParagraphInfo('д', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('ела?')
			]
		]
	},
	///////////////////////// -> 15 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Привет, как уюю у '), createParagraphInfo('тебя', {
				reviewType: reviewtype_Remove,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' дела    ', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' дела?')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Привет,'), createParagraphInfo(' ну ты даешь,', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' как у '), createParagraphInfo(' опо', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' дела?')
			]
		]
	},
	///////////////////////// -> 16 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Хрусть, '), createParagraphInfo('Хрусть', {
				reviewType: reviewtype_Remove,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(', пок ')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Хрусть, ', {
					reviewType: reviewtype_Remove,
					userName  : 'Mark Pottato',
					dateTime  : 1000000
				}), createParagraphInfo('Хрусть', {
				reviewType: reviewtype_Remove,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(', ок п')
			]
		]
	},
	///////////////////////// -> 17 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Hello Hello Hello')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Hello', null, null, {
					start: [{id: '1', name: 's1'}, {id: '2', name: 's2'}],
					end  : [{id: '1'}, {id: '2'}]
				}), createParagraphInfo(' Hello Hello')
			]
		]
	},
	///////////////////////// -> 18 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Hello', null, null, {
					start: [{id: '1', name: 's1'}],
					end  : [{id: '1'}]
				}), createParagraphInfo(' Hello Hello')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Hel', null, null, {
					start: [{id: '1', name: 's1'}],
					end  : [{id: '1'}]
				}), createParagraphInfo('lo Hello Hello')
			]
		]
	},
	///////////////////////// -> 19 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Привет привет привет')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('П', null, null, {
					start: [{id: '1', name: 's1'}, {id: '2', name: 's8'}, {
						id   : '3',
						name : 's3',
						start: true
					}]
				}),
				createParagraphInfo('ри', null, null, {start: [{id: '4', name: 's11'}]}),
				createParagraphInfo('ве', null, null, {start: [{id: '5', name: 's2'}], end: [{id: '4'}]}),
				createParagraphInfo('т '),
				createParagraphInfo('пр', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {end: [{id: '1'}]}),
				createParagraphInfo('и', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
					end: [{
						id  : '7',
						name: 's7'
					}, {id: '5'}]
				}),
				createParagraphInfo('в', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null),
				createParagraphInfo('ет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
					start: [{
						id  : '8',
						name: 's4'
					}]
				}),
				createParagraphInfo(' '),
				createParagraphInfo('п', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
					start: [{
						id  : '9',
						name: 's6'
					}]
				}),
				createParagraphInfo('ри', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
					start: [{
						id  : '11',
						name: 's5'
					}],
					end  : [{id: '11'}]
				}),
				createParagraphInfo('в', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {end: [{id: '7'}]}),
				createParagraphInfo('ет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
					end: [{
						id  : '12',
						name: 's9'
					}, {id: '13', name: 's10'}, {id: '2'}]
				}),
				createParagraphInfo(' ', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {end: [{id: '12'}]}),
				createParagraphInfo('прив', null, null, {end: [{id: '14', name: 's12'}, {id: '9'}, {id: '14'}]}),
				createParagraphInfo('ет', null, null, {end: [{id: '3'}, {id: '8'}, {id: '13'}]})
			]
		]
	},
	///////////////////////// -> 20 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Hello Hello')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Hello He'),
				createParagraphInfo('llo Hel', null, null, {start: [{id: '1', name: 's1'}], end: [{id: '1'}]}),
				createParagraphInfo('lo')
			]
		]
	},
	///////////////////////// -> 21 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Hello Hello Hello')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Hello', null, null, {start: [{id: '1', name: 's1'}], end: [{id: '1'}]}),
				createParagraphInfo(' Hello')
			]
		]
	},
	///////////////////////// -> 22 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Hello hello hello')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Hello hello '),
				createParagraphInfo('hello hello', null, null, {start: [{id: '1', name: 's1'}, {id: '1'}]})
			]
		]
	},
	///////////////////////// -> 23 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('hello', undefined, undefined, undefined, {
					comments: {
						start: [{start: true, id: 1}],
						end  : [{id: 1, data: {text: '123'}}]
					}
				})
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('hello', undefined, undefined, undefined, {
					comments: {
						start: [{start: true, id: 1}],
						end  : [{id: 1, data: {text: '1234'}}]
					}
				})
			]
		]
	},
	///////////////////////// -> 24 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(' ', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {text: '123'}
					}]
				}
			}), createParagraphInfo('привет привет')]],
		revisedDocument : [
			[createParagraphInfo('привет привет '), createParagraphInfo('привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {text: '123'}
					}]
				}
			})]]
	},
	///////////////////////// -> 25 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('Привет '), createParagraphInfo('привет ', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo('при', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 1
					}]
				}
			}), createParagraphInfo('вет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 2
					}]
				}
			}), createParagraphInfo(' прив', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text     : '1',
							quoteText: 'привет привет'
						}
					}]
				}
			}), createParagraphInfo('ет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 2,
						data : {
							text     : '1',
							quoteText: 'вет прив'
						}
					}]
				}
			}), createParagraphInfo(' привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 1,
						data : {
							text     : '1',
							quoteText: 'привет привет'
						}
					}]
				}
			})]],
		revisedDocument : [
			[createParagraphInfo('Привет привет '), createParagraphInfo('при', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo('вет ', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 1
					}]
				}
			}), createParagraphInfo('прив', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 2
					}]
				}
			}), createParagraphInfo('ет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 1,
						data : {
							text     : '1',
							quoteText: 'вет прив'
						}
					}]
				}
			}), createParagraphInfo(' привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 2,
						data : {
							text     : '1',
							quoteText: 'привет'
						}
					}, {start: false, id: 0, data: {text: '1', quoteText: 'привет привет'}}]
				}
			})]
		]
	}
	,
	///////////////////////// -> 26 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('арварвар', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text     : '111',
							quoteText: 'арварвар'
						}
					}]
				}
			})]
		],
		revisedDocument : [
			[createParagraphInfo('арварвар', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text     : '222',
							quoteText: 'арварвар'
						}
					}]
				}
			})]
		]
	},
	///////////////////////// -> 27 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('привет '), createParagraphInfo('привет привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text     : '123',
							quoteText: 'привет привет'
						}
					}]
				}
			})]
		],
		revisedDocument : [
			[createParagraphInfo('привет ', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo('привет привет'), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text     : '123',
							quoteText: 'привет привет привет'
						}
					}]
				}
			})]
		]
	},
	///////////////////////// -> 28 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('Привет привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(' привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text     : '123',
							quoteText: 'Привет привет'
						}
					}]
				}
			})]],
		revisedDocument : [
			[createParagraphInfo('Привет привет привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text     : '123',
							quoteText: 'Привет привет привет'
						}
					}]
				}
			})]]
	},
	///////////////////////// -> 29 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('Кусь '), createParagraphInfo('пусь', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}}), createParagraphInfo(' тусь', undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'пусь'}}]}}), createParagraphInfo(' привет привет')]
			],
		revisedDocument : [
			[createParagraphInfo('Привет '), createParagraphInfo('привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}}), createParagraphInfo(' привет привет привет', undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'привет'}}]}})]
			]
	},
	///////////////////////// -> 30 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('привет '), createParagraphInfo('привет привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(' привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text     : '123',
							quoteText: 'привет привет'
						}
					}]
				}
			})]],
		revisedDocument : [
			[createParagraphInfo('привет '), createParagraphInfo('привет ой', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(' привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text     : '123',
							quoteText: 'привет ой'
						}
					}]
				}
			})]]
	},
	///////////////////////// -> 31 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('привет '), createParagraphInfo('привет привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}, {start: true, id: 1}]}}), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 1, data:{text: '1', quoteText: 'привет привет', arrAnswers: ['123', '12']}}, {start: false, id: 0, data:{text: '1', quoteText: 'привет привет', arrAnswers: ['123', '12', '1', '2']}}]}})]
		],
		revisedDocument : [
			[createParagraphInfo('привет '), createParagraphInfo('привет привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}, {start: true, id: 1}]}}), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '1', quoteText: 'привет привет', arrAnswers: ['123', '12', '1', '2']}}, {start: false, id: 1, data:{text: '1', quoteText: 'привет привет', arrAnswers: ['123', '12', '1']}}]}})]
		]
	},
	///////////////////////// -> 32 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('привет '), createParagraphInfo('привет привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}}), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'привет привет', arrAnswers: ['1234', '12', '13', '14']}}]}})]
		],
		revisedDocument : [
			[createParagraphInfo('привет '), createParagraphInfo('привет привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}, {start: true, id: 1}]}}), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 1, data:{text: '123', quoteText: 'привет привет', arrAnswers: ['1234', '12']}}, {start: false, id: 0, data:{text: '123', quoteText: 'привет привет', arrAnswers: ['1234', '12', '13']}}]}})]
		]
	},
	///////////////////////// -> 33 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('привет привет '), createParagraphInfo('привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}}), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'привет', arrAnswers: ['1']}}]}})]
		],
		revisedDocument : [
			[createParagraphInfo('привет '), createParagraphInfo('привет ', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}}), createParagraphInfo('привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 1}]}}), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 1, data:{text: '123', quoteText: 'привет', arrAnswers: null}}, {start: false, id: 0, data:{text: '123', quoteText: 'привет привет', arrAnswers: ['1']}}]}})]
		]
	},
	///////////////////////// -> 34<- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('Привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}}), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'Привет', arrAnswers: ['1', '2', '3']}}]}})]
		],
		revisedDocument : [
			[createParagraphInfo('Привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}}), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'Привет', arrAnswers: ['4', '2', '3']}}]}})]
		]
	},
	///////////////////////// -> 35<- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('Привет привет привет привет')]
		],
		revisedDocument : [
			[createParagraphInfo('Привет ', undefined, undefined, {start: [{id: '6', name: 's6'}, {id: '1', name: 's1'}]}), createParagraphInfo('п', undefined, undefined, {start: [{id: '3', name: 's3'}]}, {comments:{start:[{start: true, id: 0}, {start: true, id: 1}]}}), createParagraphInfo('ри', undefined, undefined, {start: [{id: '7', name: 's7'}]}), createParagraphInfo('в', undefined, undefined, {start: [{id: '2', name: 's2'}]}), createParagraphInfo('е', undefined, undefined, {start: [{id: '3'}, {id: '5', name: 's5'}]}), createParagraphInfo('т', undefined, undefined, {start: [{id: '1'}]}), createParagraphInfo(' ', undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '4', quoteText: 'привет', arrAnswers: ['53']}}]}}), createParagraphInfo('пр', undefined, undefined, {start: [{id: '4', name: 's4'}]}), createParagraphInfo('ивет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 2}]}}), createParagraphInfo(' п', undefined, undefined, {start: [{id: '7'}, {id: '4'}, {id: '2'}]}), createParagraphInfo('рив', undefined, undefined, {start: [{id: '5'}]}), createParagraphInfo('ет', undefined, undefined, undefined, {comments:{start:[{start: false, id: 2, data:{text: '43212', quoteText: 'ивет прив', arrAnswers: null}}]}}), createParagraphInfo(undefined, undefined, undefined, {start: [{id: '6'}]}, {comments:{start:[{start: false, id: 1, data:{text: '123', quoteText: 'привет привет привет', arrAnswers: ['432']}}]}})]
		]
	},
	///////////////////////// -> 36 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('Привет привет привет')]
		],
		revisedDocument : [
			[createParagraphInfo('Привет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: '1', name: 's1'}]}, {comments:{start:[{start: true, id: 0}]}}), createParagraphInfo(' ', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: '3', name: 's3'}]}), createParagraphInfo('при', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: '2', name: 's2'}]}, {comments:{start:[{start: true, id: 1}]}}), createParagraphInfo('вет', new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000), undefined, {start: [{id: '3'}]}), createParagraphInfo(' приве', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {comments:{start:[{start: false, id: 1, data:{text: '123', quoteText: 'привет', arrAnswers: null}}]}}), createParagraphInfo('т', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: '2'}]}), createParagraphInfo(undefined, undefined, undefined, {start: [{id: '1'}]}, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'Привет привет привет', arrAnswers: null}}]}})]
		]
	},
	///////////////////////// -> 37 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('Привет привет привет')]
		],
		revisedDocument : [
			[createParagraphInfo('При', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}, {start: true, id: 1}]}}), createParagraphInfo('в', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined), createParagraphInfo('е', new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000), undefined, undefined, {comments:{start:[{start: true, id: 2}]}}), createParagraphInfo('т', undefined, undefined), createParagraphInfo(' ', new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000), undefined, undefined, {comments:{start:[{start: false, id: 1, data:{text: '432', quoteText: 'Привет', arrAnswers: null}}]}}), createParagraphInfo('п', new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000), undefined, undefined, {comments:{start:[{start: true, id: 3}]}}), createParagraphInfo('р', new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000), undefined, undefined, {comments:{start:[{start: true, id: 4}]}}), createParagraphInfo('иве', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined), createParagraphInfo('т', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {comments:{start:[{start: false, id: 2, data:{text: '432', quoteText: 'ет приве', arrAnswers: null}}, {start: false, id: 4, data:{text: '432', quoteText: 'риве', arrAnswers: null}}]}}), createParagraphInfo(' привет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'Привет привет', arrAnswers: null}}]}}), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 3, data:{text: '432', quoteText: 'привет привет', arrAnswers: null}}]}})]
		]
	},
	///////////////////////// -> 38 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('Привет привет ', undefined, undefined), createParagraphInfo('привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}})], [createParagraphInfo('Привет привет привет', undefined, undefined), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'приветПривет привет привет', arrAnswers: null}}]}})]		],
		revisedDocument : [
			[createParagraphInfo('Привет ', undefined, undefined), createParagraphInfo('привет привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}})], [createParagraphInfo('Привет привет привет', undefined, undefined), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'привет приветПривет привет привет', arrAnswers: null}}]}})]
		]
	}
];

const arrAnswers = [
	/////////////////////////////////// -> 1 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo(undefined, {reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000})],
			[
				createParagraphInfo('Привет', {
					reviewType: reviewtype_Add,
					userName  : 'Valdemar',
					dateTime  : 3000000
				}), createParagraphInfo(undefined, {reviewType: reviewtype_Add, userName: 'Valdemar', dateTime: 3000000})
			]
		]
	},
	/////////////////////////////////// -> 2 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo()]
		]
	},
	/////////////////////////////////// -> 3 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('Привет', {
				reviewType: reviewtype_Remove,
				userName  : 'Valdemar',
				dateTime  : 3000000
			}), createParagraphInfo(undefined, {reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000})],
			[createParagraphInfo('Приветище', {
				reviewType: reviewtype_Add,
				userName  : 'Valdemar',
				dateTime  : 3000000
			}), createParagraphInfo(undefined, {reviewType: reviewtype_Add, userName: 'Valdemar', dateTime: 3000000})]
		]
	},
	/////////////////////////////////// -> 4 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('Привет', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(undefined)]
		]
	},
	/////////////////////////////////// -> 5 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('Привет', {
				reviewType: reviewtype_Remove,
				userName  : 'Valdemar',
				dateTime  : 3000000
			}), createParagraphInfo(undefined, {reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000})],
			[createParagraphInfo('Приветище', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(undefined, {reviewType: reviewtype_Add, userName: 'Valdemar', dateTime: 3000000})]
		]
	},
	/////////////////////////////////// -> 6 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('При', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('вет, как дела?')]
		]
	},
	/////////////////////////////////// -> 7 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('Привет'), createParagraphInfo(' Привет', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' Привет')]
		]
	},
	/////////////////////////////////// -> 8 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Привет'), createParagraphInfo(' ой', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' ', {
				reviewType: reviewtype_Remove,
				userName  : 'Valdemar',
				dateTime  : 3000000
			}), createParagraphInfo('Привет', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' Привет')
			]
		]
	},
	/////////////////////////////////// -> 9 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Привет, ', {
					reviewType: reviewtype_Add,
					userName  : 'John Smith',
					dateTime  : 1000000
				}), createParagraphInfo('как дела?')
			]
		]
	},
	/////////////////////////////////// -> 10 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Приветик', {
					reviewType: reviewtype_Add,
					userName  : 'John Smith',
					dateTime  : 1000000
				}), createParagraphInfo('Привет', {
				reviewType: reviewtype_Remove,
				userName  : 'Valdemar',
				dateTime  : 3000000
			}), createParagraphInfo(', ', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('как дела?')
			]
		]
	},
	/////////////////////////////////// -> 11 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Привет, как дела?'), createParagraphInfo(' ', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('Хорошо', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 2000000
			}), createParagraphInfo('Нормально', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(', а у тебя как?', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			})
			]
		]
	},
	/////////////////////////////////// -> 12 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Привет, как дела?'), createParagraphInfo(' ', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('Хорошо', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('Нормально', {
				reviewType: reviewtype_Remove,
				userName  : 'Valdemar',
				dateTime  : 3000000
			}), createParagraphInfo(', а у тебя как?', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			})
			]
		]
	},
	/////////////////////////////////// -> 13 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Привет, как дела?   '), createParagraphInfo(' ', {
				reviewType: reviewtype_Add,
				userName  : 'Valdemar',
				dateTime  : 3000000
			}), createParagraphInfo(' Нормально', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('Хорошо', {
				reviewType: reviewtype_Remove,
				userName  : 'Valdemar',
				dateTime  : 3000000
			}), createParagraphInfo(', а у тебя как?', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			})
			]
		]
	},
	/////////////////////////////////// -> 14 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Пр', {
					reviewType: reviewtype_Remove,
					userName  : 'John Smith',
					dateTime  : 1000000
				}), createParagraphInfo('и', {
				reviewType: reviewtype_Remove,
				userName  : 'John Smith',
				dateTime  : 1000000
			}, {
				reviewType: reviewtype_Add,
				userName  : 'John Smoth',
				dateTime  : 2000000
			}), createParagraphInfo('в'), createParagraphInfo('е', {
				reviewType: reviewtype_Add,
				userName  : 'John Smoth',
				dateTime  : 2000000
			}), createParagraphInfo('т'), createParagraphInfo(',', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' к'), createParagraphInfo('а', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('к', {
				reviewType: reviewtype_Remove,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' '), createParagraphInfo('д', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('ел', {
				reviewType: reviewtype_Remove,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('а?')
			]
		]
	},
	/////////////////////////////////// -> 15 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Привет,'), createParagraphInfo(' ну ты даешь,', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' как'), createParagraphInfo(' уюю', {
				reviewType: reviewtype_Remove,
				userName  : 'Valdemar',
				dateTime  : 3000000
			}), createParagraphInfo(' у '), createParagraphInfo(' опо', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('тебя', {
				reviewType: reviewtype_Remove,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' дела    ', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(' дела?')
			]
		]
	},
	/////////////////////////////////// -> 16 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Хрусть, ', {
					reviewType: reviewtype_Remove,
					userName  : 'Mark Pottato',
					dateTime  : 1000000
				}), createParagraphInfo('Хрусть', {
				reviewType: reviewtype_Remove,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo(', '), createParagraphInfo('ок п', {
				reviewType: reviewtype_Add,
				userName  : 'Valdemar',
				dateTime  : 3000000
			}), createParagraphInfo('пок ', {reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000})
			]
		]
	},
	/////////////////////////////////// -> 17 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Hello', null, null, {
					start: [{id: '1', name: 's1'}, {id: '2', name: 's2'}],
					end  : [{id: '1'}, {id: '2'}]
				}), createParagraphInfo(' Hello Hello')
			]
		]
	},
	/////////////////////////////////// -> 18 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Hello', null, null, {
					start: [{id: '1', name: 's1'}],
					end  : [{id: '1'}]
				}), createParagraphInfo(' Hello Hello')
			]
		]
	},
	/////////////////////////////////// -> 19 <- ////////////////////////////////////////////
	{
		finalDocument: [[
			createParagraphInfo('П', null, null, {
				start: [{id: '2', name: 's1'}, {id: '6', name: 's8'}, {
					id: '10', name: 's3'
				}]
			}),
			createParagraphInfo('ри', null, null, {
				start: [{
					id: '1', name: 's11'
				}]
			}),
			createParagraphInfo('ве', null, null, {
				start: [{id: '3', name: 's2'}], end: [{id: '1'}]
			}),
			createParagraphInfo('т '),
			createParagraphInfo('пр', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {end: [{id: '2'}]}), createParagraphInfo('и', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
				end: [{
					id: '5', name: 's7'
				}, {id: '3'}]
			}),
			createParagraphInfo('в', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null),
			createParagraphInfo('ет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
				start: [{
					id: '11', name: 's4'
				}]
			}),
			createParagraphInfo(' ', new CCreatingReviewInfo('Valdemar', reviewtype_Add, 3000000)),
			createParagraphInfo('п', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
				start: [{
					id: '8', name: 's6'
				}]
			}),
			createParagraphInfo('ри', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
				start: [{id: '4', name: 's5'}], end: [{id: '4'}]
			}),
			createParagraphInfo('в', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {end: [{id: '5'}]}),
			createParagraphInfo('ет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
				end: [{id: '6'}, {id: '7', name: 's9'}, {id: '12', name: 's10'}]
			}), createParagraphInfo(' ', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {end: [{id: '7'}]}),
			createParagraphInfo('прив', null, null, {
				end: [{
					id: '9', name: 's12'
				}, {id: '8'}, {id: '9'}]
			}),
			createParagraphInfo('ет', null, null, {end: [{id: '10'}, {id: '11'}, {id: '12'}]})]]
	},
	/////////////////////////////////// -> 20 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Hello ', createFindingReviewInfo(reviewtype_Add)),
				createParagraphInfo('He', null, null, {end: [{id: '1', name: 's1'}]}),
				createParagraphInfo('llo Hel', null, null, {end: [{id: '1'}]}),
				createParagraphInfo('lo')
			]
		]
	},
	/////////////////////////////////// -> 21 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Hello ', createFindingReviewInfo(reviewtype_Remove)),
				createParagraphInfo('Hello', null, null, {start: [{id: '1', name: 's1'}], end: [{id: '1'}]}),
				createParagraphInfo(' Hello')
			]
		]
	},
	/////////////////////////////////// -> 22 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Hello'),
				createParagraphInfo(' hello', createFindingReviewInfo(reviewtype_Add)),
				createParagraphInfo(' '),
				createParagraphInfo('hello hello', null, null, {start: [{id: '1', name: 's1'}, {id: '1'}]})
			]
		]
	},
	/////////////////////////////////// -> 23 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('hello', undefined, undefined, undefined, {
					comments: {
						start: [{
							start: true,
							id   : 1
						}, {start: true, id: 2}],
						end  : [{
							id  : 2,
							data: {text: '1234'}
						}, {id: 1, data: {text: '123'}}]
					}
				})

			]
		]
	},
	/////////////////////////////////// -> 24 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(' привет ', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {text: '123'}
					}]
				}
			}), createParagraphInfo('привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 1
					}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 1,
						data : {text: '123'}
					}]
				}
			})]
		]
	},
	/////////////////////////////////// -> 25 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('Привет '), createParagraphInfo('привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(' '), createParagraphInfo('при', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 1
					}]
				}
			}), createParagraphInfo('вет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 2
					}]
				}
			}), createParagraphInfo(' ', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text     : '1',
							quoteText: 'привет привет'
						}
					}]
				}
			}), createParagraphInfo('прив', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 3
					}]
				}
			}), createParagraphInfo('ет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 2,
						data : {
							text     : '1',
							quoteText: 'вет прив'
						}
					}]
				}
			}), createParagraphInfo(' привет', undefined, undefined, undefined, {
				comments: {
					start: [{start: false, id: 1, data: {text: '1', quoteText: 'привет привет'}}, {
						start: false,
						id   : 3,
						data : {
							text     : '1',
							quoteText: 'привет'
						}
					}]
				}
			})]
		]
	},
	/////////////////////////////////// -> 26 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('арварвар', undefined, undefined, undefined, {
				comments: {
					start: [{start: true, id: 1}, {
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text     : '222',
							quoteText: 'арварвар'
						}
					}, {start: false, id: 1, data: {text: '111', quoteText: 'арварвар'}}]
				}
			})]
		]
	},
	/////////////////////////////////// -> 27 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('привет привет привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text     : '123',
							quoteText: 'привет привет привет'
						}
					}]
				}
			})]
		]
	}
	,
	/////////////////////////////////// -> 28 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('Привет привет', undefined, undefined, undefined, {
				comments: {
					start: [{start: true, id: 1}, {
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(' привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 1,
						data : {
							text     : '123',
							quoteText: 'Привет привет'
						}
					}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text     : '123',
							quoteText: 'Привет привет привет'
						}
					}]
				}
			})]]
	},
	/////////////////////////////////// -> 29 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('Привет ', new CCreatingReviewInfo('Valdemar', reviewtype_Add, 3000000), undefined), createParagraphInfo('привет', new CCreatingReviewInfo('Valdemar', reviewtype_Add, 3000000), undefined, undefined, {comments:{start:[{start: true, id: 0}]}}), createParagraphInfo(' привет', new CCreatingReviewInfo('Valdemar', reviewtype_Add, 3000000), undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'привет'}}]}}), createParagraphInfo('Кусь ', new CCreatingReviewInfo('Valdemar', reviewtype_Remove, 3000000), undefined), createParagraphInfo('пусь', new CCreatingReviewInfo('Valdemar', reviewtype_Remove, 3000000), undefined, undefined, {comments:{start:[{start: true, id: 1}]}}), createParagraphInfo(' тусь', new CCreatingReviewInfo('Valdemar', reviewtype_Remove, 3000000), undefined, undefined, {comments:{start:[{start: false, id: 1, data:{text: '123', quoteText: 'пусь'}}]}}), createParagraphInfo(' привет привет')]
			]
	}
	,
	/////////////////////////////////// -> 30 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('привет '), createParagraphInfo('привет ', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo('ой', new CCreatingReviewInfo('Valdemar', reviewtype_Add, 3000000), undefined), createParagraphInfo('привет', new CCreatingReviewInfo('Valdemar', reviewtype_Remove, 3000000), undefined),  createParagraphInfo(' привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text: '123',
							quoteText: 'привет ой'
						}
					}]
				}
			})]]
	},
	/////////////////////////////////// -> 31 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('привет '), createParagraphInfo('привет привет', undefined, undefined, undefined, {
				comments: {
					start: [{start: true, id: 1}, {
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{start: false, id: 1, data: {text: '1', quoteText: 'привет привет', arrAnswers: ['123', '12', '1']}}, {
						start: false,
						id   : 0,
						data : {
							text      : '1',
							quoteText : 'привет привет',
							arrAnswers: ['123', '12', '1', '2']
						}
					}]
				}
			})]
		]
	},
	/////////////////////////////////// -> 32 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('привет '), createParagraphInfo('привет привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}, {start: true, id: 1}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{start: false, id: 1, data: {text: '123', quoteText: 'привет привет', arrAnswers: ['1234', '12', '13', '14']}}, {
						start: false,
						id   : 0,
						data : {
							text      : '123',
							quoteText : 'привет привет',
							arrAnswers: ['1234', '12']
						}
					}]
				}
			})]
		]
	},
	/////////////////////////////////// -> 33 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('привет '), createParagraphInfo('привет ', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				}
			}), createParagraphInfo('привет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 1
					}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [ {
						start: false,
						id   : 0,
						data : {
							text      : '123',
							quoteText : 'привет привет',
							arrAnswers: ['1']
						}
					}, {
						start: false,
						id   : 1,
						data : {
							text      : '123',
							quoteText : 'привет',
							arrAnswers: ['1']
						}
					}]
				}
			})]
		]
	},
	/////////////////////////////////// -> 34 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('Привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}, {start: true, id: 1}]}}), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'Привет', arrAnswers: ['4', '2', '3']}}, {start: false, id: 1, data:{text: '1', quoteText: 'Привет', arrAnswers: null}}]}})]
		]
	},
	/////////////////////////////////// -> 35 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('Привет ', undefined, undefined, {
				start: [{
					id  : '7',
					name: 's6'
				},{id: '2', name: 's1'}]
			}), createParagraphInfo('п', undefined, undefined, {
				start: [{
					id  : '1',
					name: 's3'
				}]
			}, {
				comments: {
					start: [{
						start: true,
						id   : 1
					}, {start: true, id: 0} ]
				}
			}), createParagraphInfo('ри', undefined, undefined, {
				start: [{
					id  : '3',
					name: 's7'
				}]
			}), createParagraphInfo('в', undefined, undefined, {
				start: [{
					id  : '5',
					name: 's2'
				}]
			}), createParagraphInfo('е', undefined, undefined, {
				start: [{id: '1'}, {
					id  : '6',
					name: 's5'
				}]
			}), createParagraphInfo('т', undefined, undefined, {start: [{id: '2'}]}), createParagraphInfo(' ', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text      : '4',
							quoteText : 'привет',
							arrAnswers: ['53']
						}
					}]
				}
			}), createParagraphInfo('пр', undefined, undefined, {
				start: [{
					id  : '4',
					name: 's4'
				}]
			}), createParagraphInfo('ивет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 2
					}]
				}
			}), createParagraphInfo(' ', undefined, undefined, {start: [{id: '3'}, {id: '4'}, {id: '5'}]}), createParagraphInfo('п', undefined, undefined), createParagraphInfo('рив', undefined, undefined, {start: [{id: '6'}]}), createParagraphInfo('ет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 2,
						data : {
							text      : '43212',
							quoteText : 'ивет прив',
							arrAnswers: null
						}
					}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, {start: [{id: '7'}]}, {
				comments: {
					start: [{
						start: false,
						id   : 1,
						data : {
							text      : '123',
							quoteText : 'привет привет привет',
							arrAnswers: ['432']
						}
					}]
				}
			})]
		]
	},
	/////////////////////////////////// -> 36 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('Привет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: '3', name: 's1'}]}, {comments:{start:[{start: true, id: 0}]}}), createParagraphInfo(' ', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: '1', name: 's3'}]}), createParagraphInfo('при', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: '2', name: 's2'}]}, {comments:{start:[{start: true, id: 1}]}}), createParagraphInfo('вет', new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000), undefined, {start: [{id: '1'}]}), createParagraphInfo(' приве', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {comments:{start:[{start: false, id: 1, data:{text: '123', quoteText: 'привет', arrAnswers: null}}]}}), createParagraphInfo('т', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: '2'}]}), createParagraphInfo(undefined, undefined, undefined, {start: [{id: '3'}]}, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'Привет привет привет', arrAnswers: null}}]}})]
		]
	},
	/////////////////////////////////// -> 37 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('При', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}, {start: true, id: 1}]
				}
			}), createParagraphInfo('в', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined), createParagraphInfo('е', new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000), undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 2
					}]
				}
			}), createParagraphInfo('т', undefined, undefined), createParagraphInfo(' ', new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000), undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 1,
						data : {
							text      : '432',
							quoteText : 'Привет',
							arrAnswers: null
						}
					}]
				}
			}), createParagraphInfo('п', new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000), undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 3
					}]
				}
			}), createParagraphInfo('р', new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000), undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 4
					}]
				}
			}), createParagraphInfo('иве', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined), createParagraphInfo('т', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {
				comments: {
					start: [{start: false, id: 4, data: {text: '432', quoteText: 'риве', arrAnswers: null}}, {
						start: false,
						id   : 2,
						data : {
							text      : '432',
							quoteText : 'ет приве',
							arrAnswers: null
						}
					}]
				}
			}), createParagraphInfo(' привет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text      : '123',
							quoteText : 'Привет привет',
							arrAnswers: null
						}
					}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 3,
						data : {
							text      : '432',
							quoteText : 'привет привет',
							arrAnswers: null
						}
					}]
				}
			})]
		]
	},
	/////////////////////////////////// -> 38 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo('Привет ', undefined, undefined), createParagraphInfo('привет привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}})], [createParagraphInfo('Привет привет привет', undefined, undefined), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'привет приветПривет привет привет', arrAnswers: null}}]}})]
		]
	}
];

const comments = [
	'Merging an empty document and a document with a non-reviewed paragraph',
	'Merging empty documents',
	'Merging documents with different paragraphs without review',
	'Merging two documents with the same content in paragraphs, but different in review',
	'Merging two documents with different paragraphs, the first without review, the second with review',
	'Merging two documents with the same content, where part of the word has a review',
	'Merging identical documents, in the middle of the document the word has another review',
	'Merging documents with insertion and review',
	'Merging to start',
	'Merging documents with different origins',
	'Merging documents with differences in text with the same review',
	'Merging documents with differences in text with different reviews',
	'Merging documents with differences in text with different reviews and requiring additional reviews',
	'Merging identical documents with different types of reviews in letters',
	'Merging two documents with changes in common text',
	'Merging two documents with complicated check of review types and spaces',
	'Merging two documents with bookmarks',
	'Blocking merging duplicate bookmarks',
	'Merging some bookmarks',
	'Merging in start of document and add bookmark',
	'Remove from start of document and add bookmark',
	'Merging start and end bookmarks to start of word',
	'Merging different comments for equal word',
	'Merging equal comments in different places in document',
	'Merging equal comments in different places in document',
	'Merging different comments for equal word',
	'Merging equal comments with different quote texts in end',
	'Merging equal comments with different quote texts in start',
	'Merging insertions and deletions with comments',
	'Merging comments with part in insertions and deletions',
	'Collapse comments with comparable answers',
	'Collapse comments with comparable answers',
	'Collapse comments with comparable answers',
	'Adding a separate answer as a comment',
	'Merging bookmarks and comments',
	'Merging bookmarks, comments and review',
	'Merging comments and review',
	'Merging two paragraph with different starts of comment'
];

function merge(oMainDocument, oRevisedDocument, fCallback)
{
	const oMerge = new AscCommonWord.CDocumentMerge(oMainDocument, oRevisedDocument, new AscCommonWord.ComparisonOptions());
	const fOldMergeCallback = oMerge.applyLastMergeCallback;
	oMerge.applyLastMergeCallback = function ()
	{
		fOldMergeCallback.call(this);
		fCallback();
	}
	oMerge.merge();
}

function getTestObject(oDocument)
{
	return oDocument.getTestObject();
}


$(function ()
{

	QUnit.module("Unit-tests for merge documents feature");

	QUnit.test("Test", function (assert)
	{
		AscFormat.ExecuteNoHistory(function ()
		{
			for (let i = 0; i < arrTestObjectsInfo.length; i += 1)
			{
				const oTestInformation = arrTestObjectsInfo[i];
				merge(readMainDocument(oTestInformation.originalDocument), readRevisedDocument(oTestInformation.revisedDocument), function ()
				{
					const oResultDocument = mockEditor.WordControl.m_oLogicDocument;
					oMainComments = oResultDocument.Comments;
					const oResultObject = getTestObject(oResultDocument);
					assert.deepEqual(oResultObject, getTestObject(readMainDocument(arrAnswers[i].finalDocument)), comments[i]);
				});
			}
		}, this, []);
	});
});

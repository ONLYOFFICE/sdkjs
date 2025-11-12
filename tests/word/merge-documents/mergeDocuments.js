/*
 * (c) Copyright Ascensio System SIA 2010-2024
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
const arrWordTestDocumentInfo = [
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
					start: [{id: 1, name: 's1'}, {id: 2, name: 's2'}],
					end  : [{id: 1}, {id: 2}]
				}), createParagraphInfo(' Hello Hello')
			]
		]
	},
	///////////////////////// -> 18 <- /////////////////////////////
	{
		originalDocument: [
			[
				createParagraphInfo('Hello', null, null, {
					start: [{id: 1, name: 's1'}],
					end  : [{id: 1}]
				}), createParagraphInfo(' Hello Hello')
			]
		],
		revisedDocument : [
			[
				createParagraphInfo('Hel', null, null, {
					start: [{id: 1, name: 's1'}],
					end  : [{id: 1}]
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
					start: [{id: 1, name: 's1'}, {id: 2, name: 's8'}, {
						id   : '3',
						name : 's3',
						start: true
					}]
				}),
				createParagraphInfo('ри', null, null, {start: [{id: 4, name: 's11'}]}),
				createParagraphInfo('ве', null, null, {start: [{id: 5, name: 's2'}], end: [{id: 4}]}),
				createParagraphInfo('т '),
				createParagraphInfo('пр', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {end: [{id: 1}]}),
				createParagraphInfo('и', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
					end: [{
						id  : '7',
						name: 's7'
					}, {id: 5}]
				}),
				createParagraphInfo('в', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null),
				createParagraphInfo('ет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
					start: [{
						id  : 8,
						name: 's4'
					}]
				}),
				createParagraphInfo(' '),
				createParagraphInfo('п', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
					start: [{
						id  : 9,
						name: 's6'
					}]
				}),
				createParagraphInfo('ри', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
					start: [{
						id  : 11,
						name: 's5'
					}],
					end  : [{id: 11}]
				}),
				createParagraphInfo('в', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {end: [{id: 7}]}),
				createParagraphInfo('ет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
					end: [{
						id  : 12,
						name: 's9'
					}, {id: 13, name: 's10'}, {id: 2}]
				}),
				createParagraphInfo(' ', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {end: [{id: 12}]}),
				createParagraphInfo('прив', null, null, {end: [{id: 14, name: 's12'}, {id: 9}, {id: 14}]}),
				createParagraphInfo('ет', null, null, {end: [{id: 3}, {id: 8}, {id: 13}]})
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
				createParagraphInfo('llo Hel', null, null, {start: [{id: 1, name: 's1'}], end: [{id: 1}]}),
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
				createParagraphInfo('Hello', null, null, {start: [{id: 1, name: 's1'}], end: [{id: 1}]}),
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
				createParagraphInfo('hello hello', null, null, {start: [{id: 1, name: 's1'}, {id: 1}]})
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
			[createParagraphInfo('привет '), createParagraphInfo('привет ', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}}), createParagraphInfo('привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 1}]}}), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'привет привет', arrAnswers: ['1']}}, {start: false, id: 1, data:{text: '123', quoteText: 'привет', arrAnswers: null}}]}})]
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
			[createParagraphInfo('Привет ', undefined, undefined, {start: [{id: 6, name: 's6'}, {id: 1, name: 's1'}]}), createParagraphInfo('п', undefined, undefined, {start: [{id: 3, name: 's3'}]}, {comments:{start:[{start: true, id: 0}, {start: true, id: 1}]}}), createParagraphInfo('ри', undefined, undefined, {start: [{id: 7, name: 's7'}]}), createParagraphInfo('в', undefined, undefined, {start: [{id: 2, name: 's2'}]}), createParagraphInfo('е', undefined, undefined, {start: [{id: 3}, {id: 5, name: 's5'}]}), createParagraphInfo('т', undefined, undefined, {start: [{id: 1}]}), createParagraphInfo(' ', undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '4', quoteText: 'привет', arrAnswers: ['53']}}]}}), createParagraphInfo('пр', undefined, undefined, {start: [{id: 4, name: 's4'}]}), createParagraphInfo('ивет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 2}]}}), createParagraphInfo(' п', undefined, undefined, {start: [{id: 7}, {id: 4}, {id: 2}]}), createParagraphInfo('рив', undefined, undefined, {start: [{id: 5}]}), createParagraphInfo('ет', undefined, undefined, undefined, {comments:{start:[{start: false, id: 2, data:{text: '43212', quoteText: 'ивет прив', arrAnswers: null}}]}}), createParagraphInfo(undefined, undefined, undefined, {start: [{id: 6}]}, {comments:{start:[{start: false, id: 1, data:{text: '123', quoteText: 'привет привет привет', arrAnswers: ['432']}}]}})]
		]
	},
	///////////////////////// -> 36 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo('Привет привет привет')]
		],
		revisedDocument : [
			[createParagraphInfo('Привет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: 1, name: 's1'}]}, {comments:{start:[{start: true, id: 0}]}}), createParagraphInfo(' ', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: 3, name: 's3'}]}), createParagraphInfo('при', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: 2, name: 's2'}]}, {comments:{start:[{start: true, id: 1}]}}), createParagraphInfo('вет', new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000), undefined, {start: [{id: 3}]}), createParagraphInfo(' приве', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {comments:{start:[{start: false, id: 1, data:{text: '123', quoteText: 'привет', arrAnswers: null}}]}}), createParagraphInfo('т', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: 2}]}), createParagraphInfo(undefined, undefined, undefined, {start: [{id: 1}]}, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'Привет привет привет', arrAnswers: null}}]}})]
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
			[createParagraphInfo('Привет привет ', undefined, undefined), createParagraphInfo('привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}})], [createParagraphInfo('Привет привет привет', undefined, undefined), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'привет\r\nПривет привет привет', arrAnswers: null}}]}})]		],
		revisedDocument : [
			[createParagraphInfo('Привет ', undefined, undefined), createParagraphInfo('привет привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}})], [createParagraphInfo('Привет привет привет', undefined, undefined), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'привет привет\r\nПривет привет привет', arrAnswers: null}}]}})]
		]
	},
	///////////////////////// -> 39 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo("在真实、充分地表达各自意愿的基础上，根据《中华人民共和国合同法》的规定，达成如下协议，并由双方共同恪守。")]
		],
		revisedDocument: [
			[createParagraphInfo("在真实、充分地表达各自意愿的基础上，根据《中华人民共和国"),
			createParagraphInfo("合同法", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000)),
			createParagraphInfo("民法典", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
			createParagraphInfo("》的规定，达成如下协议，并由双方共同恪守。")]
		]
	},
	///////////////////////// -> 40 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo("1.这是一个测试段落，这是一个测试段落，这是一个测试段落，这是一个测试段落")],
			[createParagraphInfo("2. 这是一个测试段落，这是一个测试段落"), createParagraphInfo("，这是一个测试段落", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000)), createParagraphInfo("，这是一个测试段落")],
			[createParagraphInfo("3.这是一个测试段落，这是一个测试段落，这是一个测试段落，这是一个测试段落")]
		],
		revisedDocument: [
			[createParagraphInfo("1.这是一个测试段落，这是一个测试段落，这是一个测试段落，这是一个测试段落")],
			[createParagraphInfo("2. 这是一个测试段落，这是一个测试段落"), createParagraphInfo("，这是一个测试段落", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000), undefined, undefined, {comments: {
					start: [{
						id   : 1,
						start: true
					}],
					end: [
						{id: 1,
						data: {
							text: "comment"
						}
						}
					]
				}}), createParagraphInfo("，这是一个测试段落")],
			[createParagraphInfo("3.这是一个测试段落，这是一个测试段落，这是一个测试段落，这是一个测试段落")]
		],
	},
	///////////////////////// -> 41 <- /////////////////////////////
	{
		originalDocument: [
			[createParagraphInfo("1.这是一个测试段落，这是一个测试段落，这是一个测试段落，这是一个测试段落")],
				[createParagraphInfo("2. 这是一个测试段落，这是一个测试段落"), createParagraphInfo("，这是一个测试段落", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000), undefined, undefined, {comments: {
						start: [{
							id   : 1,
							start: true
						}],
						end: [
							{id: 1,
								data: {
									text: "comment"
								}
							}
						]
					}}), createParagraphInfo("，这是一个测试段落")],
				[createParagraphInfo("3.这是一个测试段落，这是一个测试段落，这是一个测试段落，这是一个测试段落")]
			],
		revisedDocument: [
			[createParagraphInfo("1.这是一个测试段落，这是一个测试段落，这是一个测试段落，这是一个测试段落")],
			[createParagraphInfo("2. 这是一个测试段落，这是一个测试段落"), createParagraphInfo("，这是一个测试段落", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000)), createParagraphInfo("，这是一个测试段落")],
			[createParagraphInfo("3.这是一个测试段落，这是一个测试段落，这是一个测试段落，这是一个测试段落")]
		]
	},
	///////////////////////// -> 42 <- /////////////////////////////
	{
		originalDocument: [
			[
			createParagraphInfo("Приве", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript}}),
			createParagraphInfo("т "),
			createParagraphInfo("прив", undefined, undefined, undefined, {textPr: {Underline: true}}),
			createParagraphInfo("ет пр"),
			createParagraphInfo("иве", undefined, undefined, undefined, {textPr: {Bold: true}}),
			createParagraphInfo("т")
			]
		],
		revisedDocument: [
			[
			createParagraphInfo("П", undefined, undefined,  undefined,{textPr: {VertAlign: AscCommon.vertalign_SubScript}}),
			createParagraphInfo("риве", undefined, undefined,  undefined,{textPr: {VertAlign: AscCommon.vertalign_SuperScript}}),
			createParagraphInfo("т "),
			createParagraphInfo("прив", undefined, undefined,  undefined,{textPr: {Underline: true}}),
			createParagraphInfo("ет "),
			createParagraphInfo("пр", undefined, undefined,  undefined,{textPr: {Underline: true}}),
			createParagraphInfo("и", undefined, undefined,  undefined,{textPr: {Bold: true}}),
			createParagraphInfo("ве", undefined, undefined,  undefined,{textPr: {Bold: true, Italic: true}}),
			createParagraphInfo("т", undefined, undefined,  undefined,{textPr: {Italic: true}})
			]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("пр"),
				createParagraphInfo("иве", undefined, undefined, undefined, {textPr: {VertAlign: AscFormat.vertalign_SuperScript}}),
				createParagraphInfo("т п"),
				createParagraphInfo("риве",undefined, undefined, undefined,  {textPr: {Strikeout: true}}),
				createParagraphInfo("т п"),
				createParagraphInfo("ривет",undefined, undefined, undefined,  {textPr: {Bold: true}}),
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("пр"),
				createParagraphInfo("иве", undefined, undefined, undefined, {textPr: {VertAlign: AscFormat.vertalign_SuperScript}}),
				createParagraphInfo("т п"),
				createParagraphInfo("риве", undefined, undefined, undefined, {textPr: {Strikeout: true}}),
				createParagraphInfo("т привет п"),
				createParagraphInfo("ривет", undefined, undefined, undefined, {textPr: {Bold: true}})
			]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("привет "),
				createParagraphInfo("при", undefined, undefined, undefined, {textPr: {Bold: true}}),
				createParagraphInfo("вет привет")
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("привет привет привет")
			]
		]
	},
	{
		originalDocument: [
			[createParagraphInfo("прив"), createParagraphInfo("ет ", undefined, undefined, {
				start: [{
					id  : 0,
					name: "s2"
				}]
			}, {textPr: {VertAlign: 1,}}), createParagraphInfo("при", undefined, undefined, undefined, {
				textPr: {
					Bold     : true,
					VertAlign: AscCommon.vertalign_SubScript,
					BoldCS   : true,
				}
			}), createParagraphInfo("в", undefined, undefined, {start: [{id: 0}]}, {
				textPr: {
					Bold     : true,
					Italic   : true,
					VertAlign: AscCommon.vertalign_SubScript,
					BoldCS   : true,
					ItalicCS : true,
				}
			}), createParagraphInfo("е", undefined, undefined, {
				start: [{
					id  : 1,
					name: "s1"
				}]
			}, {
				comments: {start: [{start: true, id: 0}]},
				textPr  : {Bold: true, Italic: true, BoldCS: true, ItalicCS: true,}
			}), createParagraphInfo("т", undefined, undefined, undefined, {
				textPr: {
					Bold  : true,
					BoldCS: true,
				}
			}), createParagraphInfo(" ", undefined, undefined, undefined, {
				textPr: {
					Bold      : true,
					FontSize  : 12,
					BoldCS    : true,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("п", undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text      : "1",
							quoteText : "ет ",
							arrAnswers: null
						}
					}]
				},
				textPr  : {
					Bold      : true,
					Underline : true,
					FontSize  : 12,
					BoldCS    : true,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("рив", undefined, undefined, undefined, {
				textPr: {
					Bold     : true,
					Underline: true,
					BoldCS   : true,
				}
			}), createParagraphInfo("ет привет", undefined, undefined, {start: [{id: 1}]}, {textPr: {Underline: true,}}), createParagraphInfo()]],
		revisedDocument: [
			[createParagraphInfo("пр", undefined, undefined, {start: [{id: 0, name: "s4"}]}, {
				comments : {
					start: [{
						start: true,
						id   : 0
					}]
				}, textPr: {VertAlign: 2,}
			}), createParagraphInfo("и", undefined, undefined, {
				start: [{
					id  : 2,
					name: "s7"
				}]
			}, {textPr: {VertAlign: 2,}}), createParagraphInfo("в"), createParagraphInfo("ет ", undefined, undefined, {
				start: [{
					id  : 3,
					name: "s2"
				}, {id: 4, name: "s3"}, {id: 5, name: "s5"}, {id: 0}]
			}, {textPr: {VertAlign: 1,}}), createParagraphInfo("при", undefined, undefined, undefined, {
				textPr: {
					Bold      : true,
					FontSize  : 14,
					VertAlign : 1,
					BoldCS    : true,
					FontSizeCS: 14,
				}
			}), createParagraphInfo("в", undefined, undefined, {start: [{id: 3}]}, {
				textPr: {
					Bold      : true,
					Italic    : true,
					FontSize  : 14,
					VertAlign : 1,
					BoldCS    : true,
					ItalicCS  : true,
					FontSizeCS: 14,
				}
			}), createParagraphInfo("е", undefined, undefined, {
				start: [{
					id  : 6,
					name: "s1"
				}, {id: 2}, {id: 4}]
			}, {
				comments: {start: [{start: true, id: 1}]},
				textPr  : {Bold: true, Italic: true, BoldCS: true, ItalicCS: true,}
			}), createParagraphInfo("т", undefined, undefined, undefined, {
				textPr: {
					Bold  : true,
					BoldCS: true,
				}
			}), createParagraphInfo(" ", undefined, undefined, undefined, {
				textPr: {
					Bold      : true,
					FontSize  : 12,
					BoldCS    : true,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("п", undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 1,
						data : {
							text      : "1",
							quoteText : "ет ",
							arrAnswers: null
						}
					}, {start: true, id: 2}]
				},
				textPr  : {
					Bold      : true,
					FontSize  : 12,
					BoldCS    : true,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("рив", undefined, undefined, {start: [{id: 9, name: "s6"}]}, {
				textPr: {
					Bold  : true,
					BoldCS: true,
				}
			}), createParagraphInfo("ет ", undefined, undefined, {start: [{id: 5}, {id: 6}]}, {textPr: {Underline: true,}}), createParagraphInfo("п", undefined, undefined, {start: [{id: 9}]}, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text      : "3",
							quoteText : "привет привет привет ",
							arrAnswers: null
						}
					}]
				},
				textPr  : {Underline: true,}
			}), createParagraphInfo("рив", undefined, undefined, undefined, {
				textPr: {
					Underline : true,
					FontSize  : 12,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("ет", undefined, undefined, undefined, {textPr: {Underline: true,}}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 2,
						data : {
							text      : "2",
							quoteText : "привет привет",
							arrAnswers: null
						}
					}]
				},
			})]
		]
	},
	{
		originalDocument: [
			[createParagraphInfo("пр", undefined, undefined, undefined, {}), createParagraphInfo("и", undefined, undefined, undefined, {
				textPr: {
					Bold     : true,
					Underline: true,
					BoldCS   : true,
				}
			}), createParagraphInfo("в", undefined, undefined, {start: [{id: 0, name: "s2"}]}, {
				textPr: {
					Bold     : true,
					Underline: true,
					BoldCS   : true,
				}
			}), createParagraphInfo("ет ", undefined, undefined, undefined, {
				textPr: {
					Bold  : true,
					BoldCS: true,
				}
			}), createParagraphInfo("пр", undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				},
			}), createParagraphInfo("и", undefined, undefined, undefined, {textPr: {VertAlign: 2,}}), createParagraphInfo("вет", undefined, undefined, {
				start: [{
					id  : 2,
					name: "s1"
				}, {id: 0}]
			}, {
				comments: {start: [{start: true, id: 1}]},
				textPr  : {VertAlign: 1,}
			}), createParagraphInfo(" ", undefined, undefined, {start: [{id: 2}]}, {}), createParagraphInfo("прив", undefined, undefined, undefined, {
				textPr: {
					Italic   : true,
					Strikeout: true,
					ItalicCS : true,
				}
			}), createParagraphInfo("е", undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 1,
						data : {
							text      : "1",
							quoteText : "вет прив",
							arrAnswers: null
						}
					}]
				},
				textPr  : {
					Italic   : true,
					Strikeout: true,
					ItalicCS : true,
				}
			}), createParagraphInfo("т ", undefined, undefined, undefined, {textPr: {Strikeout: true,}}), createParagraphInfo("привет", undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text      : "2",
							quoteText : "привет привет ",
							arrAnswers: null
						}
					}]
				},
				textPr  : {
					Strikeout: true,
					VertAlign: 1,
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {})]
		],
		revisedDocument: [
			[createParagraphInfo("пр", undefined, undefined, {
				start: [{id: 0, name: "s6"}, {
					id  : 1,
					name: "s3"
				}]
			}, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				},
			}), createParagraphInfo("и", undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 1
					}, {start: true, id: 2}]
				},
				textPr  : {
					Bold     : true,
					Underline: true,
					BoldCS   : true,
				}
			}), createParagraphInfo("в", undefined, undefined, {start: [{id: 5, name: "s2"}]}, {
				textPr: {
					Bold     : true,
					Underline: true,
					BoldCS   : true,
				}
			}), createParagraphInfo("ет", undefined, undefined, undefined, {
				textPr: {
					Bold  : true,
					BoldCS: true,
				}
			}), createParagraphInfo(" ", undefined, undefined, undefined, {
				textPr: {
					Bold      : true,
					FontSize  : 12,
					BoldCS    : true,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("пр", undefined, undefined, {
				start: [{
					id  : 6,
					name: "s7"
				}, {id: 0}]
			}, {
				comments: {start: [{start: true, id: 3}]},
				textPr  : {Strikeout: true, VertAlign: 1,}
			}), createParagraphInfo("и", undefined, undefined, undefined, {
				comments : {
					start: [{
						start: false,
						id   : 1,
						data : {
							text      : "5",
							quoteText : "ивет пр",
							arrAnswers: null
						}
					}]
				}, textPr: {Strikeout: true, VertAlign: 2,}
			}), createParagraphInfo("вет", undefined, undefined, {
				start: [{
					id  : 8,
					name: "s1"
				}, {id: 1}, {id: 5}, {id: 6}]
			}, {
				comments: {start: [{start: true, id: 4}]},
				textPr  : {VertAlign: 1,}
			}), createParagraphInfo(" ", undefined, undefined, {start: [{id: 8}]}, {}), createParagraphInfo("пр", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {
				textPr: {
					Italic   : true,
					Strikeout: true,
					VertAlign: 1,
					ItalicCS : true,
				}
			}), createParagraphInfo("и", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {
				start: [{
					id  : 11,
					name: "s5"
				}]
			}, {
				textPr: {
					Italic    : true,
					Strikeout : true,
					FontSize  : 12,
					VertAlign : 1,
					ItalicCS  : true,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("в", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 2,
						data : {
							text      : "4",
							quoteText : "ивет привет при",
							arrAnswers: null
						}
					}]
				},
				textPr  : {
					Italic    : true,
					Strikeout : true,
					FontSize  : 12,
					ItalicCS  : true,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("е", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 4,
						data : {
							text      : "1",
							quoteText : "вет прив",
							arrAnswers: null
						}
					}]
				},
				textPr  : {
					Italic   : true,
					Strikeout: true,
					ItalicCS : true,
				}
			}), createParagraphInfo("т ", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {textPr: {Strikeout: true,}}), createParagraphInfo("при", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {
				start: [{
					id  : 12,
					name: "s4"
				}]
			}, {
				comments : {
					start: [{
						start: false,
						id   : 3,
						data : {text: "2", quoteText: "привет привет ", arrAnswers: null}
					}, {start: false, id: 0, data: {text: "3", quoteText: "привет привет привет ", arrAnswers: null}}]
				}, textPr: {Strikeout: true, VertAlign: 1,}
			}), createParagraphInfo("вет", undefined, undefined, undefined, {
				textPr: {
					Strikeout: true,
					VertAlign: 1,
				}
			}), createParagraphInfo(undefined, undefined, undefined, {start: [{id: 11}, {id: 12}]}, {})]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("привет привет")
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("привет hello"), 
				createParagraphInfo(" при", undefined, undefined, undefined, {textPr: {Bold: true}}), 
				createParagraphInfo("вет")
			]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("归乙方所有，甲方"),
				createParagraphInfo("【有权无偿】", undefined, undefined, undefined, {textPr: {Bold: true}}),
				createParagraphInfo("使用。")
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("归乙方所有，"),
				createParagraphInfo("测试", undefined, undefined, undefined, {
					comments: {
						start: [{id: 0, start: true}],
						end  : [{id: 0,data: {text: '4325', quoteText: "测试"}
						}]
					}
				}),
				createParagraphInfo("甲方", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000)),
				createParagraphInfo("【有权无偿】", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {textPr: {Bold: true}}),
				createParagraphInfo("使用", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000)),
				createParagraphInfo("。")
			]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("培养更多优秀人才")
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("培养更多优秀人才gdfgfdgdfgfdgfdgfdgdg gdfg dfgdfgfdgfdgfdg")
			]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("培养更培养更培养更培养更培养更培养更培养更培养更培养更")
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("培养更培养更培养更培养更培养更培养更培养更培养更培养更"),
				createParagraphInfo("培养更", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000))
			]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("培养更培养更培养更培养更培养更培养更培养更培养更培养更")
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("培养更", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("培养更培养更培养更培养更培养更培养更培养更培养更培养更")

			]
		]
	},
];

const arrWordAnswers = [
	/////////////////////////////////// -> 1 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Привет', {
					reviewType: reviewtype_Add,
					userName  : 'Valdemar',
					dateTime  : 3000000
				})
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
			[createParagraphInfo('Приветище', {
				reviewType: reviewtype_Add,
				userName  : 'Valdemar',
				dateTime  : 3000000
			}), createParagraphInfo('Привет', {
				reviewType: reviewtype_Remove,
				userName  : 'Valdemar',
				dateTime  : 3000000
			})]
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
			[createParagraphInfo('Приветище', {
				reviewType: reviewtype_Add,
				userName  : 'John Smith',
				dateTime  : 1000000
			}), createParagraphInfo('Привет', {
				reviewType: reviewtype_Remove,
				userName  : 'Valdemar',
				dateTime  : 3000000
			})]
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
					start: [{id: 1, name: 's1'}, {id: 2, name: 's2'}],
					end  : [{id: 1}, {id: 2}]
				}), createParagraphInfo(' Hello Hello')
			]
		]
	},
	/////////////////////////////////// -> 18 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Hello', null, null, {
					start: [{id: 1, name: 's1'}],
					end  : [{id: 1}]
				}), createParagraphInfo(' Hello Hello')
			]
		]
	},
	/////////////////////////////////// -> 19 <- ////////////////////////////////////////////
	{
		finalDocument: [[
			createParagraphInfo('П', null, null, {
				start: [{id: 2, name: 's1'}, {id: 6, name: 's8'}, {
					id: 10, name: 's3'
				}]
			}),
			createParagraphInfo('ри', null, null, {
				start: [{
					id: 1, name: 's11'
				}]
			}),
			createParagraphInfo('ве', null, null, {
				start: [{id: 3, name: 's2'}], end: [{id: 1}]
			}),
			createParagraphInfo('т '),
			createParagraphInfo('пр', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {end: [{id: 2}]}), createParagraphInfo('и', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
				end: [{
					id: 5, name: 's7'
				}, {id: 3}]
			}),
			createParagraphInfo('в', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null),
			createParagraphInfo('ет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
				start: [{
					id: 11, name: 's4'
				}]
			}),
			createParagraphInfo(' ', new CCreatingReviewInfo('Valdemar', reviewtype_Add, 3000000)),
			createParagraphInfo('п', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
				start: [{
					id: 8, name: 's6'
				}]
			}),
			createParagraphInfo('ри', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
				start: [{id: 4, name: 's5'}], end: [{id: 4}]
			}),
			createParagraphInfo('в', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {end: [{id: 5}]}),
			createParagraphInfo('ет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {
				end: [{id: 6}, {id: 7, name: 's9'}, {id: 12, name: 's10'}]
			}), createParagraphInfo(' ', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), null, {end: [{id: 7}]}),
			createParagraphInfo('прив', null, null, {
				end: [{
					id: 9, name: 's12'
				}, {id: 8}, {id: 9}]
			}),
			createParagraphInfo('ет', null, null, {end: [{id: 10}, {id: 11}, {id: 12}]})]]
	},
	/////////////////////////////////// -> 20 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Hello ', createFindingReviewInfo(reviewtype_Add)),
				createParagraphInfo('He', null, null, {end: [{id: 1, name: 's1'}]}),
				createParagraphInfo('llo Hel', null, null, {end: [{id: 1}]}),
				createParagraphInfo('lo')
			]
		]
	},
	/////////////////////////////////// -> 21 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo('Hello ', createFindingReviewInfo(reviewtype_Remove)),
				createParagraphInfo('Hello', null, null, {start: [{id: 1, name: 's1'}], end: [{id: 1}]}),
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
				createParagraphInfo('hello hello', null, null, {start: [{id: 1, name: 's1'}, {id: 1}]})
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
							data: {text: '1234', quoteText: "hello"}
						}, {id: 1, data: {text: '123', quoteText: "hello"}}]
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
						data : {text: '123', quoteText: "привет"}
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
						data : {text: '123', quoteText: "привет"}
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
					start: [{
						start: false,
						id   : 3,
						data : {
							text     : '1',
							quoteText: 'привет'
						}
					}, {start: false, id: 1, data: {text: '1', quoteText: 'привет привет'}}]
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
							quoteText: 'привет ойпривет'
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
					start: [{
						start: true,
						id   : 0
					}, {start: true, id: 1}]
				}
			}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text      : '1',
							quoteText : 'привет привет',
							arrAnswers: ['123', '12', '1', '2']
						}
					}, {start: false, id: 1, data: {text: '1', quoteText: 'привет привет', arrAnswers: ['123', '12', '1']}}]
				}
			})]
		]
	},
	/////////////////////////////////// -> 32 <- ////////////////////////////////////////////
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
					start: [{
						start: false,
						id   : 0,
						data : {
							text      : '123',
							quoteText : 'привет привет',
							arrAnswers: ['1234', '12']
						}
					}, {start: false, id: 1, data: {text: '123', quoteText: 'привет привет', arrAnswers: ['1234', '12', '13', '14']}}]
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
					id  : 7,
					name: 's6'
				},{id: 2, name: 's1'}]
			}), createParagraphInfo('п', undefined, undefined, {
				start: [{
					id  : 1,
					name: 's3'
				}]
			}, {
				comments: {
					start: [{start: true, id: 0}, {
						start: true,
						id   : 1
					} ]
				}
			}), createParagraphInfo('ри', undefined, undefined, {
				start: [{
					id  : 3,
					name: 's7'
				}]
			}), createParagraphInfo('в', undefined, undefined, {
				start: [{
					id  : 5,
					name: 's2'
				}]
			}), createParagraphInfo('е', undefined, undefined, {
				start: [{id: 1}, {
					id  : 6,
					name: 's5'
				}]
			}), createParagraphInfo('т', undefined, undefined, {start: [{id: 2}]}), createParagraphInfo(' ', undefined, undefined, undefined, {
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
					id  : 4,
					name: 's4'
				}]
			}), createParagraphInfo('ивет', undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 2
					}]
				}
			}), createParagraphInfo(' ', undefined, undefined, {start: [{id: 3}, {id: 4}, {id: 5}]}), createParagraphInfo('п', undefined, undefined), createParagraphInfo('рив', undefined, undefined, {start: [{id: 6}]}), createParagraphInfo('ет', undefined, undefined, undefined, {
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
			}), createParagraphInfo(undefined, undefined, undefined, {start: [{id: 7}]}, {
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
			[createParagraphInfo('Привет', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: 3, name: 's1'}]}, {comments:{start:[{start: true, id: 0}]}}), createParagraphInfo(' ', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: 1, name: 's3'}]}), createParagraphInfo('при', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: 2, name: 's2'}]}, {comments:{start:[{start: true, id: 1}]}}), createParagraphInfo('вет', new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000), undefined, {start: [{id: 1}]}), createParagraphInfo(' приве', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {comments:{start:[{start: false, id: 1, data:{text: '123', quoteText: 'привет', arrAnswers: null}}]}}), createParagraphInfo('т', new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {start: [{id: 2}]}), createParagraphInfo(undefined, undefined, undefined, {start: [{id: 3}]}, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'Привет привет привет', arrAnswers: null}}]}})]
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
					start: [{
						start: false,
						id   : 2,
						data : {
							text      : '432',
							quoteText : 'ет приве',
							arrAnswers: null
						}
					},
					{start: false, id: 4, data: {text: '432', quoteText: 'риве', arrAnswers: null}
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
			[createParagraphInfo('Привет ', undefined, undefined), createParagraphInfo('привет привет', undefined, undefined, undefined, {comments:{start:[{start: true, id: 0}]}})], [createParagraphInfo('Привет привет привет', undefined, undefined), createParagraphInfo(undefined, undefined, undefined, undefined, {comments:{start:[{start: false, id: 0, data:{text: '123', quoteText: 'привет привет\r\nПривет привет привет', arrAnswers: null}}]}})]
		]
	},
	/////////////////////////////////// -> 39 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo("在真实、充分地表达各自意愿的基础上，根据《中华人民共和国"),
			createParagraphInfo("合同法", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000)),
			createParagraphInfo("民法典", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
			createParagraphInfo("》的规定，达成如下协议，并由双方共同恪守。")]
		]
	},
	/////////////////////////////////// -> 40 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo("1.这是一个测试段落，这是一个测试段落，这是一个测试段落，这是一个测试段落")],
			[createParagraphInfo("2. 这是一个测试段落，这是一个测试段落"), createParagraphInfo("，这是一个测试段落", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000), undefined, undefined, {comments: {
					start: [{
						id   : 1,
						start: true
					}],
					end: [
						{id: 1,
							data: {
								text: "comment",
								quoteText: "，这是一个测试段落"
							}
						}
					]
				}}), createParagraphInfo("，这是一个测试段落")],
			[createParagraphInfo("3.这是一个测试段落，这是一个测试段落，这是一个测试段落，这是一个测试段落")]
		]
	},
	/////////////////////////////////// -> 41 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[createParagraphInfo("1.这是一个测试段落，这是一个测试段落，这是一个测试段落，这是一个测试段落")],
			[createParagraphInfo("2. 这是一个测试段落，这是一个测试段落"), createParagraphInfo("，这是一个测试段落", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000), undefined, undefined, {comments: {
					start: [{
						id   : 1,
						start: true
					}],
					end: [
						{id: 1,
							data: {
								text: "comment",
								quoteText: "，这是一个测试段落"
							}
						}
					]
				}}), createParagraphInfo("，这是一个测试段落")],
			[createParagraphInfo("3.这是一个测试段落，这是一个测试段落，这是一个测试段落，这是一个测试段落")]
		]
	},
	/////////////////////////////////// -> 42 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
			createParagraphInfo("П", undefined, undefined,  undefined,{textPr: {VertAlign: AscCommon.vertalign_SubScript}}),
			createParagraphInfo("риве", undefined, undefined,  undefined,{textPr: {VertAlign: AscCommon.vertalign_SuperScript}}),
			createParagraphInfo("т "),
			createParagraphInfo("прив", undefined, undefined,  undefined,{textPr: {Underline: true}}),
			createParagraphInfo("ет "),
			createParagraphInfo("пр", undefined, undefined,  undefined,{textPr: {Underline: true}}),
			createParagraphInfo("и", undefined, undefined,  undefined,{textPr: {Bold: true}}),
			createParagraphInfo("ве", undefined, undefined,  undefined,{textPr: {Bold: true, Italic: true}}),
			createParagraphInfo("т", undefined, undefined,  undefined,{textPr: {Italic: true}})
			]
		]
	},
	/////////////////////////////////// -> 43 <- ////////////////////////////////////////////
	{
		finalDocument: [
			[
				createParagraphInfo("пр", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000)),
				createParagraphInfo("иве", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000), undefined, undefined, {textPr: {VertAlign: AscFormat.vertalign_SuperScript}}),
				createParagraphInfo("т ", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000)),
				createParagraphInfo("п"),
				createParagraphInfo("риве", undefined, undefined, undefined, {textPr: {Strikeout: true}}),
				createParagraphInfo("т привет п"),
				createParagraphInfo("ривет", undefined, undefined, undefined, {textPr: {Bold: true}})
			]
			]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("привет привет привет")
			]
		]
	},
	{
		finalDocument: [
			[createParagraphInfo("пр", undefined, undefined, {start: [{id: 0, name: "s4"}]}, {
				comments : {
					start: [{
						start: true,
						id   : 0
					}]
				}, textPr: {VertAlign: 2,}
			}), createParagraphInfo("и", undefined, undefined, {
				start: [{
					id  : 2,
					name: "s7"
				}]
			}, {textPr: {VertAlign: 2,}}), createParagraphInfo("в"), createParagraphInfo("ет ", undefined, undefined, {
				start: [{
					id  : 3,
					name: "s2"
				}, {id: 4, name: "s3"}, {id: 5, name: "s5"}, {id: 0}]
			}, {textPr: {VertAlign: 1,}}), createParagraphInfo("при", undefined, undefined, undefined, {
				textPr: {
					Bold      : true,
					FontSize  : 14,
					VertAlign : 1,
					BoldCS    : true,
					FontSizeCS: 14,
				}
			}), createParagraphInfo("в", undefined, undefined, {start: [{id: 3}], end: [{
					id  : 6,
					name: "s1"
				}]}, {
				textPr: {
					Bold      : true,
					Italic    : true,
					FontSize  : 14,
					VertAlign : 1,
					BoldCS    : true,
					ItalicCS  : true,
					FontSizeCS: 14,
				}
			}), createParagraphInfo("е", undefined, undefined, {
				start: [{id: 2}, {id: 4}]
			}, {
				comments: {start: [{start: true, id: 1}]},
				textPr  : {Bold: true, Italic: true, BoldCS: true, ItalicCS: true,}
			}), createParagraphInfo("т", undefined, undefined, undefined, {
				textPr: {
					Bold  : true,
					BoldCS: true,
				}
			}), createParagraphInfo(" ", undefined, undefined, undefined, {
				textPr: {
					Bold      : true,
					FontSize  : 12,
					BoldCS    : true,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("п", undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 1,
						data : {
							text      : "1",
							quoteText : "ет ",
							arrAnswers: null
						}
					}, {start: true, id: 2}]
				},
				textPr  : {
					Bold      : true,
					FontSize  : 12,
					BoldCS    : true,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("рив", undefined, undefined, {start: [{id: 9, name: "s6"}]}, {
				textPr: {
					Bold  : true,
					BoldCS: true,
				}
			}), createParagraphInfo("ет ", undefined, undefined, {start: [{id: 6}, {id: 5}]}, {textPr: {Underline: true,}}), createParagraphInfo("п", undefined, undefined, {start: [{id: 9}]}, {
				comments: {
					start: [{
						start: false,
						id   : 0,
						data : {
							text      : "3",
							quoteText : "привет привет привет ",
							arrAnswers: null
						}
					}]
				},
				textPr  : {Underline: true,}
			}), createParagraphInfo("рив", undefined, undefined, undefined, {
				textPr: {
					Underline : true,
					FontSize  : 12,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("ет", undefined, undefined, undefined, {textPr: {Underline: true,}}), createParagraphInfo(undefined, undefined, undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 2,
						data : {
							text      : "2",
							quoteText : "привет привет",
							arrAnswers: null
						}
					}]
				},
			})]
		]
	},
	{
		finalDocument: [
			[createParagraphInfo("пр", undefined, undefined, {
				start: [{id: 0, name: "s6"}, {
					id  : 1,
					name: "s3"
				}]
			}, {
				comments: {
					start: [{
						start: true,
						id   : 0
					}]
				},
			}), createParagraphInfo("и", undefined, undefined, undefined, {
				comments: {
					start: [{
						start: true,
						id   : 1
					}, {start: true, id: 2}]
				},
				textPr  : {
					Bold     : true,
					Underline: true,
					BoldCS   : true,
				}
			}), createParagraphInfo("в", undefined, undefined, {start: [{id: 5, name: "s2"}]}, {
				textPr: {
					Bold     : true,
					Underline: true,
					BoldCS   : true,
				}
			}), createParagraphInfo("ет", undefined, undefined, undefined, {
				textPr: {
					Bold  : true,
					BoldCS: true,
				}
			}), createParagraphInfo(" ", undefined, undefined, {
				end: [{id: 0}]
			}, {
				textPr: {
					Bold      : true,
					FontSize  : 12,
					BoldCS    : true,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("пр", undefined, undefined, {
				start: [{
					id  : 6,
					name: "s7"
				}]
			}, {
				comments: {start: [{start: true, id: 3}]},
				textPr  : {Strikeout: true, VertAlign: 1,}
			}), createParagraphInfo("и", undefined, undefined, {
				end: [{
					id  : 8,
					name: "s1"
				}, {id: 5}]
			}, {
				comments : {
					start: [{
						start: false,
						id   : 1,
						data : {
							text      : "5",
							quoteText : "ивет пр",
							arrAnswers: null
						}
					}]
				}, textPr: {Strikeout: true, VertAlign: 2,}
			}), createParagraphInfo("вет", undefined, undefined, {
				start: [{id: 1}, {id: 6}]
			}, {
				comments: {start: [{start: true, id: 4}]},
				textPr  : {VertAlign: 1,}
			}), createParagraphInfo(" ", undefined, undefined, {start: [{id: 8}]}, {}), createParagraphInfo("пр", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {
				textPr: {
					Italic   : true,
					Strikeout: true,
					VertAlign: 1,
					ItalicCS : true,
				}
			}), createParagraphInfo("и", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {
				start: [{
					id  : 11,
					name: "s5"
				}]
			}, {
				textPr: {
					Italic    : true,
					Strikeout : true,
					FontSize  : 12,
					VertAlign : 1,
					ItalicCS  : true,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("в", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 2,
						data : {
							text      : "4",
							quoteText : "ивет привет при",
							arrAnswers: null
						}
					}]
				},
				textPr  : {
					Italic    : true,
					Strikeout : true,
					FontSize  : 12,
					ItalicCS  : true,
					FontSizeCS: 12,
				}
			}), createParagraphInfo("е", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {
				comments: {
					start: [{
						start: false,
						id   : 4,
						data : {
							text      : "1",
							quoteText : "вет прив",
							arrAnswers: null
						}
					}]
				},
				textPr  : {
					Italic   : true,
					Strikeout: true,
					ItalicCS : true,
				}
			}), createParagraphInfo("т ", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {textPr: {Strikeout: true,}}), createParagraphInfo("при", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, {
				start: [{
					id  : 12,
					name: "s4"
				}]
			}, {
				comments : {
					start: [{
						start: false,
						id   : 3,
						data : {text: "2", quoteText: "привет привет ", arrAnswers: null}
					}, {start: false, id: 0, data: {text: "3", quoteText: "привет привет привет ", arrAnswers: null}}]
				}, textPr: {Strikeout: true, VertAlign: 1,}
			}), createParagraphInfo("вет", undefined, undefined, undefined, {
				textPr: {
					Strikeout: true,
					VertAlign: 1,
				}
			}), createParagraphInfo(undefined, undefined, undefined, {start: [{id: 11}, {id: 12}]}, {})]
		]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("привет"),
				createParagraphInfo(" hello", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000)),
				createParagraphInfo(" при", undefined, undefined, undefined, {textPr: {Bold: true}}),
				createParagraphInfo("вет")
			]
		]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("归乙方所有，"),
				createParagraphInfo("测试", new CCreatingReviewInfo('Valdemar', reviewtype_Add, 3000000), undefined, undefined, {
					comments: {
						start: [{id: 0, start: true}],
						end  : [{id: 0, data: {text: '4325', quoteText: "测试"}}]
					}
				}),
				createParagraphInfo("甲方", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000)),
				createParagraphInfo("【有权无偿】", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000), undefined, undefined, {textPr: {Bold: true}}),
				createParagraphInfo("使用", new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000)),
				createParagraphInfo("。")
			]
		]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("培养更多优秀人才"),
				createParagraphInfo("gdfgfdgdfgfdgfdgfdgdg gdfg dfgdfgfdgfdgfdg", new CCreatingReviewInfo('Valdemar', reviewtype_Add, 3000000)),
			]
		]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("培养更培养更培养更培养更培养更培养更培养更培养更培养更"),
				createParagraphInfo("培养更", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000))
			]
		]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("培养更", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("培养更培养更培养更培养更培养更培养更培养更培养更培养更")
			]
		]
	},
];

const arrWordComments = [
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
	'Merging two paragraph with different starts of comment',
	'Merging two documents with changes in the second',
	'Merging two documents with comment'
];

const arrSymbolDocumentTestInfo = [
	{
		originalDocument: [
			[
				createParagraphInfo("培养更培养更培养更培养更培养更培养更培养更培养更培养更cxz")
			]
		],
		revisedDocument: [
			[
			createParagraphInfo("培养更培养更培养更培养更培养更培养更培养更培养更培养更"),
			createParagraphInfo("培养更", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000))]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("依法向", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("乙方", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000)),
				createParagraphInfo("乙乙方乙方", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000), new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("原告所在地具有管辖权的人民法院起诉", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("依法向乙方", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000)),
				createParagraphInfo("甲方", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000), new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("所在地具有管辖权的人民法院起诉", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000)),
				createParagraphInfo("。")
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("依法向", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("乙方", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000)),
				createParagraphInfo("乙乙方乙方", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000),  new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("原告", new CCreatingReviewInfo("User1", reviewtype_Remove, 1000), new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("乙方", new CCreatingReviewInfo("User1", reviewtype_Add, 1000)),
				createParagraphInfo("所在地具有管辖权的人民法院起诉", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("依法向乙方", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000)),
				createParagraphInfo("甲方", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000), new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("所在地具有管辖权的人民法院起诉", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000)),
				createParagraphInfo("。"),
				]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("421421423241234124 214234314321432 241124123423141234")
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("421421423241234124 21423431214321432 241124123423141234")
			]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("421421423241234124 214234314321432 241124123423141234")
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("421421423241234124 21423431"),
				createParagraphInfo("21", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("4321432 241124123423141234"),
			]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("Hello hello hello")
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("Hello hello hellok k")
			]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("Привет привет "),
				createParagraphInfo("прив", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript}}),
				createParagraphInfo("ет", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript}}),
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("Привет привет п", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("ри", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000), undefined, undefined, {textPr: {Bold: true}}),
				createParagraphInfo("вет", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000), undefined, undefined, {textPr: {Bold: true, Italic: true}}),
				createParagraphInfo("ик", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000), undefined, undefined, {textPr: {Italic: true}}),
			]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("Привет привет "),
				createParagraphInfo("прив", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript}}),
				createParagraphInfo("ет", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript}}),
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("Привет привет п"),
				createParagraphInfo("ри", undefined, undefined, undefined, {textPr: {Bold: true}}),
				createParagraphInfo("вет", undefined, undefined, undefined, {textPr: {Bold: true, Italic: true}}),
				createParagraphInfo("ик", undefined, undefined, undefined, {textPr: {Italic: true}}),
			]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("Привет привет привет")
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("Привет пр", undefined, undefined, {start: [{id: 0, name: "s1"}], end: [{id: 1, name: "s2"}]}),
				createParagraphInfo("и", undefined, undefined, {end: [{id: 0}]}),
				createParagraphInfo("в", undefined, undefined, {end: [{id: 3, name: "s4"}]}),
				createParagraphInfo("ет", undefined, undefined, {end: [{id: 1}]}),
				createParagraphInfo("и", undefined, undefined, {end: [{id: 2, name: "s3"}, {id: 2}]}),
				createParagraphInfo("к"),
				createParagraphInfo(" при", undefined, undefined, {end: [{id: 3}]}),
				createParagraphInfo("вет"),
			]
		]
	},
	{
		originalDocument: [
			[
				createParagraphInfo("Q"),
				createParagraphInfo("wer", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript}}),
				createParagraphInfo("tyuiop", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript, Bold: true}}),
				createParagraphInfo("asdfghjklzxcvb", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript, Bold: true, Italic: true}}),
				createParagraphInfo("nmqwertyuiopasdfghjklzxcvbnmqwert", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true, Italic: true}}),
				createParagraphInfo("yuioas", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true}}),
				createParagraphInfo("dfghj", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript}}),
				createParagraphInfo("kl"),
			]
		],
		revisedDocument: [
			[
				createParagraphInfo("Q"),
				createParagraphInfo("wereeeee", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript}}),
				createParagraphInfo("tyuiop", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript, Bold: true}}),
				createParagraphInfo("asdfghjklzqwefffffdsg", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript, Bold: true, Italic: true}}),
				createParagraphInfo("fghjklzxcvbnmqwerte", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true, Italic: true}}),
				createParagraphInfo("yeeuioeeeeeas", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true}}),
				createParagraphInfo("dfghj", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript}}),
				createParagraphInfo("kl"),
			]
		]
	},
	// todo
	// {
	// 	originalDocument: [
	// 		[
	// 			createParagraphInfo("Привет привет", undefined, undefined, {start:[{id: 0, name: "s1"}], end: [{id: 0}]}),
	// 			createParagraphInfo(" привет"),
	// 		]
	// 	],
	// 	revisedDocument: [
	// 		[
	// 			createParagraphInfo("Привет приве", undefined, undefined, {start: [{id: 0, name: "s1"}, {id:1, name: "s2"}], end: [{id: 2, name: "s3"}]}),
	// 			createParagraphInfo("т", undefined, undefined, {end: [{id: 0}]}),
	// 			createParagraphInfo("ик", undefined, undefined, {end: [{id: 1}]}),
	// 			createParagraphInfo("и приве", undefined, undefined, {end:[{id: 2}]}),
	// 			createParagraphInfo("т"),
	// 		]
	// 	]
	// },
];
const arrSymbolAnswers = [
	{
		finalDocument: [
			[
			createParagraphInfo("培养更培养更培养更培养更培养更培养更培养更培养更培养更"),
			createParagraphInfo("培养更", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
			createParagraphInfo("cxz", new CCreatingReviewInfo("Valdemar", reviewtype_Remove, 3000000))]
		]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("依法向", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("乙方", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000)),
				createParagraphInfo("乙乙方乙方", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000),  new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("原告", new CCreatingReviewInfo("User1", reviewtype_Remove, 1000), new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("乙方", new CCreatingReviewInfo("User1", reviewtype_Add, 1000)),
				createParagraphInfo("所在地具有管辖权的人民法院起诉", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("依法向乙方", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000)),
				createParagraphInfo("甲方", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000), new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("所在地具有管辖权的人民法院起诉", new CCreatingReviewInfo("Mark Potato", reviewtype_Remove, 1000)),
				createParagraphInfo("。"),
			]
		]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("421421423241234124 2142343"),
				createParagraphInfo("12", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000)),
				createParagraphInfo("14321432 241124123423141234"),
			]
		]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("421421423241234124 21423431"),
				createParagraphInfo("21", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("4321432 241124123423141234"),
			]
		]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("Hello hello "),
				createParagraphInfo("hellok k", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000)),
				createParagraphInfo("hello", new CCreatingReviewInfo("Valdemar", reviewtype_Remove, 3000000)),
			]
		]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("Привет привет п", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000)),
				createParagraphInfo("ри", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000), undefined, undefined, {textPr: {Bold: true}}),
				createParagraphInfo("вет", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000), undefined, undefined, {textPr: {Bold: true, Italic: true}}),
				createParagraphInfo("ик", new CCreatingReviewInfo("Mark Potato", reviewtype_Add, 1000), undefined, undefined, {textPr: {Italic: true}}),
			]
		]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("Привет привет п"),
				createParagraphInfo("ри", undefined, undefined, undefined, {textPr: {Bold: true}}),
				createParagraphInfo("вет", undefined, undefined, undefined, {textPr: {Bold: true, Italic: true}}),
				createParagraphInfo("ик", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000), undefined, undefined, {textPr: {Italic: true}}),
			]
		]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("Привет пр", undefined, undefined, {start: [{id: 0, name: "s1"}], end: [{id: 1, name: "s2"}]}),
				createParagraphInfo("и", undefined, undefined, {end: [{id: 0}]}),
				createParagraphInfo("в", undefined, undefined, {end: [{id: 3, name: "s4"}]}),
				createParagraphInfo("ет", undefined, undefined, {end: [{id: 1}]}),
				createParagraphInfo("и", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000), undefined, {end: [{id: 2, name: "s3"}, {id: 2}]}),
				createParagraphInfo("к", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000)),
				createParagraphInfo(" при", undefined, undefined, {end: [{id: 3}]}),
				createParagraphInfo("вет"),
			]
		]
	},
	{
		finalDocument: [
			[
				createParagraphInfo("Q"),
				createParagraphInfo("wer", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript}}),
				createParagraphInfo("eeeee", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000), undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript}}),
				createParagraphInfo("tyuiop", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript, Bold: true}}),
				createParagraphInfo("asdfghjklz", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript, Bold: true, Italic: true}}),
				createParagraphInfo("xcvb", new CCreatingReviewInfo("Valdemar", reviewtype_Remove, 3000000), undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript, Bold: true, Italic: true}}),
				createParagraphInfo("nm", new CCreatingReviewInfo("Valdemar", reviewtype_Remove, 3000000), undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true, Italic: true}}),
				createParagraphInfo("qwe", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript, Bold: true, Italic: true}}),
				createParagraphInfo("fffffd", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000), undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript, Bold: true, Italic: true}}),
				createParagraphInfo("rtyuiopa", new CCreatingReviewInfo("Valdemar", reviewtype_Remove, 3000000), undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true, Italic: true}}),
				createParagraphInfo("s", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript, Bold: true, Italic: true}}),
				createParagraphInfo("g", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000), undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SuperScript, Bold: true, Italic: true}}),
				createParagraphInfo("d", new CCreatingReviewInfo("Valdemar", reviewtype_Remove,3000000), undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true, Italic: true}}),
				createParagraphInfo("fghjklzxcvbnmqwert", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true, Italic: true}}),
				createParagraphInfo("e", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000), undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true, Italic: true}}),
				createParagraphInfo("y", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true}}),
				createParagraphInfo("ee", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000), undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true}}),
				createParagraphInfo("uio", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true}}),
				createParagraphInfo("eeeee", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000), undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true}}),
				createParagraphInfo("as", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript, Bold: true}}),
				createParagraphInfo("dfghj", undefined, undefined, undefined, {textPr: {VertAlign: AscCommon.vertalign_SubScript}}),
				createParagraphInfo("kl"),
			]
		]
	},
	// todo
	// {
	// 	finalDocument: [
	// 		[
	// 			createParagraphInfo("Привет приве", undefined, undefined, {start: [{id: 0, name: "s1"}, {id:1, name: "s2"}], end: [{id: 2, name: "s3"}]}),
	// 			createParagraphInfo("т", undefined, undefined, {end: [{id: 0}]}),
	// 			createParagraphInfo("ик", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000), undefined, {end: [{id: 1}]}),
	// 			createParagraphInfo("и", new CCreatingReviewInfo("Valdemar", reviewtype_Add, 3000000)),
	// 			createParagraphInfo(" приве", undefined, undefined, {end:[{id: 2}]}),
	// 			createParagraphInfo("т"),
	// 		]
	// 	]
	// },
];
const arrSymbolComments = [];
function merge(oMainDocument, oRevisedDocument, options, fCallback)
{
	const oMerge = new AscCommonWord.CDocumentMerge(oMainDocument, oRevisedDocument, options);
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

	QUnit.test("Test word document combine", function (assert)
	{
		AscFormat.ExecuteNoHistory(function ()
		{
			for (let i = 0; i < arrWordTestDocumentInfo.length; i += 1)
			{
				const oTestInformation = arrWordTestDocumentInfo[i];
				merge(readMainDocument(oTestInformation.originalDocument), readRevisedDocument(oTestInformation.revisedDocument), new AscCommonWord.ComparisonOptions(), function ()
				{
					const oResultDocument = mockEditor.WordControl.m_oLogicDocument;
					const oResultObject = getTestObject(oResultDocument);
					assert.deepEqual(oResultObject, getTestObject(readMainDocument(arrWordAnswers[i].finalDocument)), arrWordComments[i]);
				});
			}
		}, this, []);
	});

	QUnit.test("Test symbol document combine", function (assert)
	{
		AscFormat.ExecuteNoHistory(function ()
		{
			for (let i = 0; i < arrSymbolDocumentTestInfo.length; i += 1)
			{
				const oOptions = new AscCommonWord.ComparisonOptions();
				oOptions.putWords(false);
				const oTestInformation = arrSymbolDocumentTestInfo[i];
				merge(readMainDocument(oTestInformation.originalDocument), readRevisedDocument(oTestInformation.revisedDocument), oOptions, function ()
				{
					const oResultDocument = mockEditor.WordControl.m_oLogicDocument;
					const oResultObject = getTestObject(oResultDocument);
					assert.deepEqual(oResultObject, getTestObject(readMainDocument(arrSymbolAnswers[i].finalDocument)), arrSymbolComments[i]);
				});
			}
		}, this, []);
	});
});

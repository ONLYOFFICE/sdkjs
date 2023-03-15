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

const arrTestObjectsInfo = [
    ///////////////////////// -> 1 <- /////////////////////////////
    {
        originalDocument: [
            [
                createParagraphInfo(''),
            ]
        ],
        revisedDocument: [
            [
                createParagraphInfo('Привет')
            ]
        ]
    },
    ///////////////////////// -> 2 <- /////////////////////////////
    {
        originalDocument: [
            [
                createParagraphInfo(''),
            ]
        ],
        revisedDocument: [
            [
                createParagraphInfo('')
            ]
        ]
    },
    ///////////////////////// -> 3 <- /////////////////////////////
    {
        originalDocument: [
            [
                createParagraphInfo('Привет'),
            ]
        ],
        revisedDocument: [
            [
                createParagraphInfo('Приветище')
            ]
        ]
    },
    ///////////////////////// -> 4 <- /////////////////////////////
    {
        originalDocument: [
            [
                createParagraphInfo('Привет'),
            ]
        ],
        revisedDocument: [
            [
                createParagraphInfo('Привет', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000})
            ]
        ]
    },
    ///////////////////////// -> 5 <- /////////////////////////////
    {
        originalDocument: [
            [
                createParagraphInfo('Привет'),
            ]
        ],
        revisedDocument: [
            [
                createParagraphInfo('Приветище', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000})
            ]
        ]
    },
    ///////////////////////// -> 6 <- /////////////////////////////
    {
        originalDocument: [
            [
                createParagraphInfo('Привет, как дела?'),
            ]
        ],
        revisedDocument: [
            [
                createParagraphInfo('При', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('вет, как дела?')
            ]
        ]
    },
    ///////////////////////// -> 7 <- /////////////////////////////
    {
        originalDocument: [
            [
                createParagraphInfo('Привет Привет Привет'),
            ]
        ],
        revisedDocument: [
            [
                createParagraphInfo('Привет'), createParagraphInfo(' Привет', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' Привет')
            ]
        ]
    },
    ///////////////////////// -> 8 <- /////////////////////////////
    {
        originalDocument: [
            [
                createParagraphInfo('Привет'), createParagraphInfo(' ой', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' Привет'), createParagraphInfo(' Привет'),
            ]
        ],
        revisedDocument: [
            [
                createParagraphInfo('Привет'), createParagraphInfo(' Привет', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' Привет')
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
        revisedDocument: [
            [
                createParagraphInfo('Привет, ', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('как дела?')
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
        revisedDocument: [
            [
                createParagraphInfo('Приветик, ', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('как дела?')
            ]
        ]
    },
    ///////////////////////// -> 11 <- /////////////////////////////
    {
        originalDocument: [
            [
                createParagraphInfo('Привет, как дела?'), createParagraphInfo(' Нормально, а у тебя как?', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000})
            ]
        ],
        revisedDocument: [
            [
                createParagraphInfo('Привет, как дела?'), createParagraphInfo(' Хорошо, а у тебя как?', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 2000000})
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
        revisedDocument: [
            [
                createParagraphInfo('Привет, как дела?'), createParagraphInfo(' Хорошо, а у тебя как?', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000})
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
        revisedDocument: [
            [
                createParagraphInfo('Привет, как дела?    '), createParagraphInfo(' Нормально, а у тебя как?', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000})
            ]
        ]
    },
    ///////////////////////// -> 14 <- /////////////////////////////
    {
        originalDocument: [
            [
                createParagraphInfo('При', {reviewType: reviewtype_Remove, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('вет, как д'),createParagraphInfo('ел', {reviewType: reviewtype_Remove, userName: 'John Smith', dateTime: 1000000}),createParagraphInfo('а?'),
            ]
        ],
        revisedDocument: [
            [
                createParagraphInfo('Пр'),createParagraphInfo('и', {reviewType: reviewtype_Add, userName: 'John Smoth', dateTime: 2000000}),createParagraphInfo('в'),createParagraphInfo('е', {reviewType: reviewtype_Add, userName: 'John Smoth', dateTime: 2000000}),createParagraphInfo('т'),createParagraphInfo(',', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}),createParagraphInfo(' к'),createParagraphInfo('а', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}),createParagraphInfo('к', {reviewType: reviewtype_Remove, userName: 'John Smith', dateTime: 1000000}),createParagraphInfo(' '),createParagraphInfo('д', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}),createParagraphInfo('ела?'),
            ]
        ]
    },
    ///////////////////////// -> 15 <- /////////////////////////////
    {
        originalDocument: [
            [
                createParagraphInfo('Привет, как уюю у '), createParagraphInfo('тебя', {reviewType: reviewtype_Remove, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' дела    ', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' дела?')
            ]
        ],
        revisedDocument: [
            [
                createParagraphInfo('Привет,'), createParagraphInfo(' ну ты даешь,', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' как у '), createParagraphInfo(' опо', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' дела?')
            ]
        ]
    },

];
const arrAnswers = [
    /////////////////////////////////// -> 1 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [createParagraphInfo(undefined, {reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000})],
            [
                createParagraphInfo('Привет', {reviewType: reviewtype_Add, userName: 'Valdemar', dateTime: 3000000})
            ]
        ]
    },
    /////////////////////////////////// -> 2 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [createParagraphInfo()],
        ]
    },
    /////////////////////////////////// -> 3 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [createParagraphInfo('Привет', {reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000}), createParagraphInfo(undefined, {reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000})],
            [createParagraphInfo('Приветище', {reviewType: reviewtype_Add, userName: 'Valdemar', dateTime: 3000000}), createParagraphInfo(undefined)]
        ]
    },
    /////////////////////////////////// -> 4 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [createParagraphInfo('Привет', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(undefined)]
        ]
    },
    /////////////////////////////////// -> 5 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [createParagraphInfo('Привет', {reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000}), createParagraphInfo(undefined, {reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000})],
            [createParagraphInfo('Приветище', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(undefined)]
        ]
    },
    /////////////////////////////////// -> 6 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [createParagraphInfo('При', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('вет, как дела?')],
        ]
    },
    /////////////////////////////////// -> 7 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [createParagraphInfo('Привет'), createParagraphInfo(' Привет', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' Привет')]
        ]
    },
    /////////////////////////////////// -> 8 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [
                createParagraphInfo('Привет'), createParagraphInfo(' ой', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}),createParagraphInfo(' ', {reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000}),createParagraphInfo('Привет', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' Привет')
            ]
        ]
    },
    /////////////////////////////////// -> 9 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [
                createParagraphInfo('Привет, ', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('как дела?')
            ]
        ]
    },
    /////////////////////////////////// -> 10 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [
                createParagraphInfo('Приветик',{reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('Привет',{reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000}),createParagraphInfo(', ',{reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('как дела?')
            ]
        ]
    },
    /////////////////////////////////// -> 11 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [
                createParagraphInfo('Привет, как дела?'), createParagraphInfo(' ',{reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('Хорошо', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 2000000}), createParagraphInfo('Нормально',{reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}),createParagraphInfo(', а у тебя как?', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}),
            ]
        ]
    },
    /////////////////////////////////// -> 12 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [
                createParagraphInfo('Привет, как дела?'), createParagraphInfo(' ',{reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('Хорошо', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('Нормально', {reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000}), createParagraphInfo(', а у тебя как?', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000})
            ]
        ]
    },
    /////////////////////////////////// -> 13 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [
                createParagraphInfo('Привет, как дела?   '), createParagraphInfo(' ', {reviewType: reviewtype_Add, userName: 'Valdemar', dateTime: 3000000}), createParagraphInfo(' Нормально', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}),createParagraphInfo('Хорошо', {reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000}), createParagraphInfo(', а у тебя как?', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}),
            ]
        ]
    },
    /////////////////////////////////// -> 14 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [
                createParagraphInfo('Пр', {reviewType: reviewtype_Remove, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('и', {reviewType: reviewtype_Remove, userName: 'John Smith', dateTime: 1000000}, {reviewType: reviewtype_Add, userName: 'John Smoth', dateTime: 2000000}), createParagraphInfo('в'), createParagraphInfo('е', {reviewType: reviewtype_Add, userName: 'John Smoth', dateTime: 2000000}), createParagraphInfo('т'), createParagraphInfo(',', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' к'), createParagraphInfo('а', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('к', {reviewType: reviewtype_Remove, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' '), createParagraphInfo('д', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('ел', {reviewType: reviewtype_Remove, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('а?')
            ]
        ]
    },
    /////////////////////////////////// -> 15 <- ////////////////////////////////////////////
    {
        finalDocument: [
            [
                createParagraphInfo('Привет,'),createParagraphInfo(' ну ты даешь,', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' как'), createParagraphInfo(' уюю', {reviewType: reviewtype_Remove, userName: 'Valdemar', dateTime: 3000000}),createParagraphInfo(' у '),createParagraphInfo(' опо', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo('тебя', {reviewType: reviewtype_Remove, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' дела    ', {reviewType: reviewtype_Add, userName: 'John Smith', dateTime: 1000000}), createParagraphInfo(' дела?')
            ]
        ]
    },
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
];

function merge(oMainDocument, oRevisedDocument, fCallback) {
    const oMerge = new AscCommonWord.CDocumentMerge(oMainDocument, oRevisedDocument, new AscCommonWord.ComparisonOptions());
    const fOldMergeCallback = oMerge.applyLastMergeCallback;
    oMerge.applyLastMergeCallback = function () {
        fOldMergeCallback.call(this);
        fCallback();
    }
    oMerge.merge();
}

function getTestObject(oDocument) {
    return oDocument.getTestObject();
}


$(function () {

    QUnit.module("Unit-tests for merge documents feature");

    QUnit.test("Test", function(assert)
    {
        AscFormat.ExecuteNoHistory(function () {
            for (let i = 0; i < arrTestObjectsInfo.length; i += 1) {
                const oTestInformation = arrTestObjectsInfo[i];
                merge(readMainDocument(oTestInformation.originalDocument), readRevisedDocument(oTestInformation.revisedDocument), function () {
                    const oResultDocument = mockEditor.WordControl.m_oLogicDocument;
                    const oResultObject = getTestObject(oResultDocument);
                    assert.deepEqual(oResultObject, getTestObject(readMainDocument(arrAnswers[i].finalDocument)), comments[i]);
                });
            }
        }, this, []);
    });
});

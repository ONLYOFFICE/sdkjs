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

"use strict";
(function(window, document)
{
	const CBinaryFileWriter = window['AscCommon'].CBinaryFileWriter;

	function BinaryVSDYWriter()
	{
		this.memory = null;
		this.nRealTableCount = 0;
		this.nStart = 0;
		this.nLastFilePos = 0;

		this.Write = function(document)
		{
			let t = this;
			const writer = new AscCommon.CBinaryFileWriter();
			return writer.WriteDocument3(document, false, "VSDY;v1;", function () {
				t.WriteContent(writer, document);
			});
		};
		this.WriteContent = function(binaryFileWriter, document) {
			// Define table types
			const TABLE_TYPES = {
				DOCUMENT: 1,
				APP: 2,
				CORE: 3,
				CUSTOM_PROPERTIES: 4
			};
			this.memory = binaryFileWriter

			const nTableCount = 128;//Специально ставим большое число, чтобы не увеличивать его при добавлении очередной таблицы.
			this.nRealTableCount = 0;
			this.nStart = this.memory.GetCurPosition();
			//вычисляем с какой позиции можно писать таблицы
			const nmtItemSize = 5;//5 byte
			this.nLastFilePos = this.nStart + nTableCount * nmtItemSize;
			//Write mtLen
			this.memory.WriteUChar(0);

			const t = this;
			//Write SignatureTable
			if (document.app) {
				this.WriteTable(TABLE_TYPES.APP, {Write: function(){
						document.app.toStream(binaryFileWriter);
					}});
			}
			if (document.core) {
				this.WriteTable(TABLE_TYPES.CORE, {Write: function(){
						document.core.toStream(binaryFileWriter);
					}});
			}
			if (document.customProperties && document.customProperties.hasProperties()) {
				this.WriteTable(TABLE_TYPES.CUSTOM_PROPERTIES, {Write: function(){
						document.customProperties.toStream(binaryFileWriter);
					}});
			}
			this.WriteTable(TABLE_TYPES.DOCUMENT, {Write: function(){
					binaryFileWriter.WriteRecordPPTY(0, document);
				}});

			//Пишем количество таблиц
			this.memory.Seek(this.nStart);
			this.memory.WriteUChar(this.nRealTableCount);

			//seek в конец, потому что GetBase64Memory заканчивает запись на текущей позиции.
			this.memory.Seek(this.nLastFilePos);
		}
		//todo remove coping
		this.WriteTable = function(type, oTableSer)
		{
			const nCurPos = this.WriteTableStart(type);
			oTableSer.Write();
			this.WriteTableEnd(nCurPos);
		}
		this.WriteTableStart = function(type)
		{
			//Write mtItem
			//Write mtiType
			this.memory.WriteUChar(type);
			//Write mtiOffBits
			this.memory.WriteULong(this.nLastFilePos);

			//Write table
			//Запоминаем позицию в MainTable
			const nCurPos = this.memory.GetCurPosition();
			//Seek в свободную область
			this.memory.Seek(this.nLastFilePos);
			return nCurPos;
		}
		this.WriteTableEnd = function(nCurPos)
		{
			//сдвигаем позицию куда можно следующую таблицу
			this.nLastFilePos = this.memory.GetCurPosition();
			//Seek вобратно в MainTable
			this.memory.Seek(nCurPos);

			this.nRealTableCount++;
		}
	}

	AscVisio.CVisioDocument.prototype.privateWriteAttributes = undefined;
	/**
	 * Write children to stream for CVisioDocument
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CVisioDocument.prototype.writeChildren = function (pWriter) {
		// Write documentSettings
		pWriter.WriteRecordPPTY(0, this.documentSettings);
		
		// Write colors
		if (this.colors && this.colors.length > 0) {
			pWriter.StartRecord(1);
			for (let i = 0; i < this.colors.length; i++) {
				pWriter.WriteRecordPPTY(0, this.colors[i]);
			}
			pWriter.EndRecord();
		}
		
		// Write faceNames
		if (this.faceNames && this.faceNames.length > 0) {
			pWriter.StartRecord(2);
			for (let i = 0; i < this.faceNames.length; i++) {
				pWriter.WriteRecordPPTY(0, this.faceNames[i]);
			}
			pWriter.EndRecord();
		}
		
		// Write styleSheets
		if (this.styleSheets && this.styleSheets.length > 0) {
			pWriter.StartRecord(3);
			for (let i = 0; i < this.styleSheets.length; i++) {
				pWriter.WriteRecordPPTY(0, this.styleSheets[i]);
			}
			pWriter.EndRecord();
		}
		
		// Write documentSheet
		pWriter.WriteRecordPPTY(4, this.documentSheet);
		
		// Write eventList
		if (this.eventList && this.eventList.length > 0) {
			pWriter.StartRecord(5);
			for (let i = 0; i < this.eventList.length; i++) {
				pWriter.WriteRecordPPTY(0, this.eventList[i]);
			}
			pWriter.EndRecord();
		}
		
		// Write headerFooter
		pWriter.WriteRecordPPTY(6, this.headerFooter);
		
		// Write masters
		pWriter.WriteRecordPPTY(7, this.masters);
		
		// Write pages
		pWriter.WriteRecordPPTY(8, this.pages);
		
		// Write dataConnections
		if (this.dataConnections) {
			pWriter.StartRecord(9);
			pWriter.WriteRecordPPTY(0, this.dataConnections);
			pWriter.EndRecord();
		}
		
		// Write dataRecordSets
		if (this.dataRecordSets) {
			pWriter.StartRecord(10);
			pWriter.WriteRecordPPTY(0, this.dataRecordSets);
			pWriter.EndRecord();
		}
		
		// Write solutions
		if (this.solutions) {
			pWriter.StartRecord(11);
			pWriter.WriteRecordPPTY(0, this.solutions);
			pWriter.EndRecord();
		}
		
		// Write validation
		pWriter.WriteRecordPPTY(12, this.validation);
		
		// Write commentsPart
		pWriter.StartRecord(13);
		pWriter.WriteRecordPPTY(0, this.commentsPart);
		pWriter.EndRecord();
		
		// Write windows
		pWriter.StartRecord(14);
		pWriter.WriteRecordPPTY(0, this.windows);
		pWriter.EndRecord();

		if (this.themes) {
			for (let i = 0; i < this.themes.length; i++) {
				pWriter.StartRecord(15);
				pWriter.WriteTheme(this.themes[i]);
				pWriter.EndRecord();
			}
		}
		//todo VbaProject
		//todo JsaProject
	};

	/**
	 * Write attributes to stream for CWindows
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CWindows.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.clientWidth);
		pWriter._WriteUInt2(1, this.clientHeight);
	};

	/**
	 * Write children to stream for CWindows
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CWindows.prototype.writeChildren = function (pWriter) {
		if (this.window) {
			for (let i = 0; i < this.window.length; i++) {
				pWriter.WriteRecordPPTY(0, this.window[i]);
			}
		}
	};

	AscVisio.CMasters.prototype.privateWriteAttributes = undefined;
	/**
	 * Write children to stream for CMasters
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CMasters.prototype.writeChildren = function (pWriter) {
		if (this.master) {
			for (let i = 0; i < this.master.length; i++) {
				pWriter.WriteRecordPPTY(0, this.master[i]);
			}
		}
	};

	AscVisio.CPageContents.prototype.privateWriteAttributes = undefined;
	/**
	 * Write children to stream for CPageContents
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CPageContents.prototype.writeChildren = function (pWriter) {
		// Write shapes
		if (this.shapes && this.shapes.length > 0) {
			pWriter.StartRecord(0);
			for (let i = 0; i < this.shapes.length; i++) {
				pWriter.WriteRecordPPTY(0, this.shapes[i]);
			}
			pWriter.EndRecord();
		}
		
		// Write connects
		if (this.connects && this.connects.length > 0) {
			pWriter.StartRecord(1);
			for (let i = 0; i < this.connects.length; i++) {
				pWriter.WriteRecordPPTY(0, this.connects[i]);
			}
			pWriter.EndRecord();
		}
	};

	AscVisio.CMasterContents.prototype.privateWriteAttributes = undefined;
	/**
	 * Write children to stream for CMasterContents
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CMasterContents.prototype.writeChildren = function (pWriter) {
		// Write shapes
		if (this.shapes && this.shapes.length > 0) {
			pWriter.StartRecord(0);
			for (let i = 0; i < this.shapes.length; i++) {
				pWriter.WriteRecordPPTY(0, this.shapes[i]);
			}
			pWriter.EndRecord();
		}
		
		// Write connects
		if (this.connects && this.connects.length > 0) {
			pWriter.StartRecord(1);
			for (let i = 0; i < this.connects.length; i++) {
				pWriter.WriteRecordPPTY(0, this.connects[i]);
			}
			pWriter.EndRecord();
		}
	};

	AscVisio.CPages.prototype.privateWriteAttributes = undefined;
	/**
	 * Write children to stream for CPages
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CPages.prototype.writeChildren = function (pWriter) {
		if (this.page) {
			for (let i = 0; i < this.page.length; i++) {
				pWriter.WriteRecordPPTY(0, this.page[i]);
			}
		}
	};

	/**
	 * Write attributes to stream for StyleSheet_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.StyleSheet_Type.prototype.privateWriteAttributes = function (pWriter) {
        pWriter._WriteUInt2(0, this.id);
        pWriter._WriteString2(1, this.nameU);
        pWriter._WriteString2(2, this.name);
		pWriter._WriteBool2(3, this.isCustomName);
		pWriter._WriteBool2(4, this.isCustomNameU);
		pWriter._WriteUInt2(5, this.lineStyle);
		pWriter._WriteUInt2(6, this.fillStyle);
		pWriter._WriteUInt2(7, this.textStyle);
    };
	/**
	 * Write children to stream for StyleSheet_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.StyleSheet_Type.prototype.writeChildren = function (pWriter) {
		for (let i in this.elements) {
			const elem = this.elements[i];
			switch (elem.kind) {
				case AscVisio.c_oVsdxSheetStorageKind.Cell_Type:
					pWriter.WriteRecordPPTY(0, elem);
					break;
				case AscVisio.c_oVsdxSheetStorageKind.Trigger_Type:
					pWriter.WriteRecordPPTY(1, elem);
					break;
				case AscVisio.c_oVsdxSheetStorageKind.Section_Type:
					pWriter.WriteRecordPPTY(2, elem);
					break;
			}
		}
	};

	/**
	 * Write attributes to stream for CComments
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CComments.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteBool2(0, this.showCommentTags);
	};

	/**
	 * Write children to stream for CComments
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CComments.prototype.writeChildren = function (pWriter) {
		// Write authorList
		if (this.authorList && this.authorList.length > 0) {
			pWriter.StartRecord(0);
			for (let i = 0; i < this.authorList.length; i++) {
				pWriter.WriteRecordPPTY(0, this.authorList[i]);
			}
			pWriter.EndRecord();
		}
		
		// Write commentList
		if (this.commentList && this.commentList.length > 0) {
			pWriter.StartRecord(1);
			for (let i = 0; i < this.commentList.length; i++) {
				pWriter.WriteRecordPPTY(0, this.commentList[i]);
			}
			pWriter.EndRecord();
		}
	};

	/**
	 * Write attributes to stream for CDataConnections
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CDataConnections.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.nextID);
	};

	/**
	 * Write children to stream for CDataConnections
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CDataConnections.prototype.writeChildren = function (pWriter) {
		if (this.dataConnection) {
			for (let i = 0; i < this.dataConnection.length; i++) {
				pWriter.WriteRecordPPTY(0, this.dataConnection[i]);
			}
		}
	};

	/**
	 * Write attributes to stream for CDataRecordSets
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CDataRecordSets.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.nextID);
		pWriter._WriteUInt2(1, this.activeRecordsetID);
		pWriter._WriteString2(2, this.dataWindowOrder);
	};

	/**
	 * Write children to stream for CDataRecordSets
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CDataRecordSets.prototype.writeChildren = function (pWriter) {
		if (this.dataRecordSet) {
			for (let i = 0; i < this.dataRecordSet.length; i++) {
				pWriter.WriteRecordPPTY(0, this.dataRecordSet[i]);
			}
		}
	};

	AscVisio.CValidation.prototype.privateWriteAttributes = undefined;
	/**
	 * Write children to stream for CValidation
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CValidation.prototype.writeChildren = function (pWriter) {
		// Write validationProperties
		pWriter.WriteRecordPPTY(0, this.validationProperties);
		
		// Write ruleSets
		if (this.ruleSets && this.ruleSets.length > 0) {
			pWriter.StartRecord(1);
			for (let i = 0; i < this.ruleSets.length; i++) {
				pWriter.WriteRecordPPTY(0, this.ruleSets[i]);
			}
			pWriter.EndRecord();
		}
		
		// Write issues
		if (this.issues && this.issues.length > 0) {
			pWriter.StartRecord(2);
			for (let i = 0; i < this.issues.length; i++) {
				pWriter.WriteRecordPPTY(0, this.issues[i]);
			}
			pWriter.EndRecord();
		}
	};

	/**
	 * Write attributes to stream for RuleTest_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.RuleTest_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteString2(0, this.value);
	};

	/**
	 * Write attributes to stream for RuleFilter_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.RuleFilter_Type.prototype.privateWriteAttributes = AscVisio.RuleTest_Type.prototype.privateWriteAttributes;

	/**
	 * Write attributes to stream for RowKeyValue_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.RowKeyValue_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.rowID);
		pWriter._WriteString2(1, this.value);
	};

	/**
	 * Write attributes to stream for DataColumn_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.DataColumn_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteString2(0, this.columnNameID);
		pWriter._WriteString2(1, this.name);
		pWriter._WriteString2(2, this.label);
		pWriter._WriteString2(3, this.origLabel);
		// Note: langID is skipped in the reader as a TODO
		pWriter._WriteUInt2(5, this.calendar);
		pWriter._WriteUInt2(6, this.dataType);
		pWriter._WriteString2(7, this.unitType);
		pWriter._WriteUInt2(8, this.currency);
		pWriter._WriteUInt2(9, this.degree);
		pWriter._WriteUInt2(10, this.displayWidth);
		pWriter._WriteUInt2(11, this.displayOrder);
		pWriter._WriteBool2(12, this.mapped);
		pWriter._WriteBool2(13, this.hyperlink);
	};

	/**
	 * Write attributes to stream for RuleInfo_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.RuleInfo_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.ruleID);
		pWriter._WriteUInt2(1, this.ruleSetID);
	};

	/**
	 * Write attributes to stream for IssueTarget_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.IssueTarget_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.pageID);
		pWriter._WriteUInt2(1, this.shapeID);
	};

	/**
	 * Write attributes to stream for Rule_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Rule_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.id);
		pWriter._WriteString2(1, this.category);
		pWriter._WriteString2(2, this.nameU);
		pWriter._WriteBool2(3, this.ignored);
		pWriter._WriteString2(4, this.description);
		pWriter._WriteInt2(5, this.ruleTarget);
	};

	/**
	 * Write children to stream for Rule_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Rule_Type.prototype.writeChildren = function (pWriter) {
		// Write ruleFilter
		pWriter.WriteRecordPPTY(0, this.ruleFilter);
		
		// Write ruleTest
		pWriter.WriteRecordPPTY(1, this.ruleTest);
	};

	/**
	 * Write attributes to stream for RuleSetFlags_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.RuleSetFlags_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteBool2(0, this.hidden);
	};

	/**
	 * Write attributes to stream for RowMap_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.RowMap_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.rowID);
		pWriter._WriteUInt2(1, this.pageID);
		pWriter._WriteUInt2(2, this.shapeID);
	};

	/**
	 * Write attributes to stream for PrimaryKey_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.PrimaryKey_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteString2(0, this.columnNameID);
	};

	/**
	 * Write children to stream for PrimaryKey_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.PrimaryKey_Type.prototype.writeChildren = function (pWriter) {
		// Write rowKeyValue elements
		if (this.rowKeyValue) {
			for (let i = 0; i < this.rowKeyValue.length; i++) {
				pWriter.WriteRecordPPTY(0, this.rowKeyValue[i]);
			}
		}
	};

	/**
	 * Write attributes to stream for DataColumns_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.DataColumns_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteString2(0, this.sortColumn);
		pWriter._WriteBool2(1, this.sortAsc);
	};

	/**
	 * Write children to stream for DataColumns_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.DataColumns_Type.prototype.writeChildren = function (pWriter) {
		// Write dataColumn elements
		if (this.dataColumn) {
			for (let i = 0; i < this.dataColumn.length; i++) {
				pWriter.WriteRecordPPTY(0, this.dataColumn[i]);
			}
		}
	};

	/**
	 * Write to binary format for Icon_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Icon_Type.prototype.toPPTY = function (pWriter) {
		pWriter.WriteString2(this.value);
	};

	/**
	 * Write attributes to stream for PageSheet_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.PageSheet_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteString2(0, this.uniqueID);
		pWriter._WriteUInt2(1, this.lineStyle);
		pWriter._WriteUInt2(2, this.fillStyle);
		pWriter._WriteUInt2(3, this.textStyle);
	};

	/**
	 * Write children to stream for PageSheet_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.PageSheet_Type.prototype.writeChildren = function (pWriter) {
		// Write elements (cells, triggers, sections)
		for (let i in this.elements) {
			const elem = this.elements[i];
			switch (elem.kind) {
				case AscVisio.c_oVsdxSheetStorageKind.Cell_Type:
					pWriter.WriteRecordPPTY(0, elem);
					break;
				case AscVisio.c_oVsdxSheetStorageKind.Trigger_Type:
					pWriter.WriteRecordPPTY(1, elem);
					break;
				case AscVisio.c_oVsdxSheetStorageKind.Section_Type:
					pWriter.WriteRecordPPTY(2, elem);
					break;
			}
		}
	};

	/**
	 * Write attributes to stream for tp_Type (Text Properties Type)
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.tp_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.ix);
	};

	/**
	 * Write attributes to stream for pp_Type (Text Paragraph Properties Type)
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.pp_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.ix);
	};

	/**
	 * Write attributes to stream for fld_Type (Text Field Type)
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.fld_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.ix);
		pWriter._WriteString2(1, this.value);
	};

	/**
	 * Write attributes to stream for cp_Type (Character Properties Type)
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.cp_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.ix);
	};

	/**
	 * Write attributes to stream for CommentEntry_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CommentEntry_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.authorID);
		pWriter._WriteUInt2(1, this.pageID);
		pWriter._WriteUInt2(2, this.shapeID);
		pWriter._WriteString2(3, this.date);
		pWriter._WriteString2(4, this.editDate);
		pWriter._WriteUInt2(5, this.commentID);
		pWriter._WriteUInt2(6, this.autoCommentType);
		pWriter._WriteString2(7, this.value);
	};

	/**
	 * Write attributes to stream for AuthorEntry_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.AuthorEntry_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.id);
		pWriter._WriteString2(1, this.name);
		pWriter._WriteString2(2, this.initials);
		pWriter._WriteString2(3, this.resolutionID);
	};

	/**
	 * Write attributes to stream for Issue_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Issue_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.id);
		pWriter._WriteBool2(1, this.ignored);
	};

	/**
	 * Write children to stream for Issue_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Issue_Type.prototype.writeChildren = function (pWriter) {
		// Write issueTarget
		pWriter.WriteRecordPPTY(0, this.issueTarget);
		
		// Write ruleInfo
		pWriter.WriteRecordPPTY(1, this.ruleInfo);
	};

	/**
	 * Write attributes to stream for Issue_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Issue_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteString2(0, this.id);
		pWriter._WriteString2(1, this.ignored);
		pWriter._WriteString2(2, this.ruleID);
	};

	/**
	 * Write children to stream for Issue_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Issue_Type.prototype.writeChildren = function (pWriter) {
		// Write targets if present
		if (this.targets && this.targets.length > 0) {
			pWriter.StartRecord(0);
			for (let i = 0; i < this.targets.length; i++) {
				pWriter.WriteRecordPPTY(0, this.targets[i]);
			}
			pWriter.EndRecord();
		}
	};

	/**
	 * Write attributes to stream for RuleSet_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.RuleSet_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.id);
		pWriter._WriteString2(1, this.nameU);
		pWriter._WriteString2(2, this.name);
		pWriter._WriteString2(3, this.description);
		pWriter._WriteBool2(4, this.enabled);
	};

	/**
	 * Write children to stream for RuleSet_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.RuleSet_Type.prototype.writeChildren = function (pWriter) {
		// Write ruleSetFlags if present
		pWriter.WriteRecordPPTY(0, this.ruleSetFlags);
		
		// Write rules if present
		if (this.rule) {
			for (let i = 0; i < this.rule.length; i++) {
				pWriter.WriteRecordPPTY(1, this.rule[i]);
			}
		}
	};

	/**
	 * Write attributes to stream for ValidationProperties_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.ValidationProperties_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteBool2(0, this.showIgnored);
		pWriter._WriteString2(1, this.lastValidated);
	};

	/**
	 * Write attributes to stream for DataRecordSet_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.DataRecordSet_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.id);
		pWriter._WriteUInt2(1, this.connectionID);
		pWriter._WriteString2(2, this.command);
		pWriter._WriteUInt2(3, this.options);
		pWriter._WriteString2(4, this.timeRefreshed);
		pWriter._WriteUInt2(5, this.nextRowID);
		pWriter._WriteString2(6, this.name);
		pWriter._WriteString2(7, this.rowOrder);
		pWriter._WriteBool2(8, this.refreshOverwriteAll);
		pWriter._WriteBool2(9, this.refreshNoReconciliationUI);
		pWriter._WriteUInt2(10, this.refreshInterval);
		pWriter._WriteBool2(11, this.replaceLinks);
		pWriter._WriteUInt2(12, this.checksum);
	};

	/**
	 * Write children to stream for DataRecordSet_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.DataRecordSet_Type.prototype.writeChildren = function (pWriter) {
		// Write dataColumns if present
		pWriter.WriteRecordPPTY(0, this.dataColumns);
		
		// Write primaryKey if present
		if (this.primaryKey) {
			for (let i = 0; i < this.primaryKey.length; i++) {
				pWriter.StartRecord(1);
				pWriter.WriteRecordPPTY(0, this.primaryKey[i]);
				pWriter.EndRecord();
			}
		}
		
		// Write rowMaps if present
		if (this.rowMaps) {
			for (let i = 0; i < this.rowMaps.length; i++) {
				pWriter.StartRecord(2);
				pWriter.WriteRecordPPTY(0, this.rowMaps[i]);
				pWriter.EndRecord();
			}
		}
		// Write content if present
		if (this.content) {
			pWriter.StartRecord(5);
			pWriter.StartRecord(0);
			pWriter.WriteBuffer(this.content, 0, this.content.length);
			pWriter.EndRecord();
			pWriter.EndRecord();
		}
	};

	/**
	 * Write attributes to stream for DataConnection_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.DataConnection_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.id);
		pWriter._WriteString2(1, this.fileName);
		pWriter._WriteString2(2, this.connectionString);
		pWriter._WriteString2(3, this.command);
		pWriter._WriteString2(4, this.friendlyName);
		pWriter._WriteUInt2(5, this.timeout);
		pWriter._WriteBool2(6, this.alwaysUseConnectionFile);
	};

	/**
	 * Write attributes to stream for Solution_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Solution_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteString2(0, this.name);
		pWriter._WriteString2(1, this.nameU);
		pWriter._WriteString2(2, this.xml);
	};

	/**
	 * Write children to stream for Solution_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Solution_Type.prototype.writeChildren = function (pWriter) {
		// Write content if present
		if (this.content) {
			pWriter.StartRecord(0);
			pWriter.StartRecord(0);
			pWriter.WriteBuffer(this.content, 0, this.content.length);
			pWriter.EndRecord();
			pWriter.EndRecord();
		}
	};

	/**
	 * Write attributes to stream for Window_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Window_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.id);
		pWriter._WriteUChar2(1, this.windowType);
		pWriter._WriteUInt2(2, this.windowState);
		pWriter._WriteInt2(3, this.windowLeft);
		pWriter._WriteInt2(4, this.windowTop);
		pWriter._WriteUInt2(5, this.windowWidth);
		pWriter._WriteUInt2(6, this.windowHeight);
		pWriter._WriteUChar2(7, this.containerType);
		pWriter._WriteUInt2(8, this.container);
		pWriter._WriteUInt2(9, this.page);
		pWriter._WriteUInt2(10, this.sheet);
		pWriter._WriteDoubleReal2(11, this.viewScale);
		pWriter._WriteDoubleReal2(12, this.viewCenterX);
		pWriter._WriteDoubleReal2(13, this.viewCenterY);
		pWriter._WriteString2(14, this.document);
		pWriter._WriteUInt2(15, this.parentWindow);
		pWriter._WriteBool2(16, this.readOnly);
		pWriter._WriteBool2(17, this.showRulers);
		pWriter._WriteBool2(18, this.showGrid);
		pWriter._WriteBool2(19, this.showPageBreaks);
		pWriter._WriteBool2(20, this.showGuides);
		pWriter._WriteBool2(21, this.showConnectionPoints);
		pWriter._WriteUInt2(22, this.glueSettings);
		pWriter._WriteUInt2(23, this.snapSettings);
		pWriter._WriteUInt2(24, this.snapExtensions);
		pWriter._WriteBool2(25, this.snapAngles);
		pWriter._WriteBool2(26, this.dynamicGridEnabled);
		pWriter._WriteDoubleReal2(27, this.tabSplitterPos);
		pWriter._WriteUInt2(28, this.stencilGroup);
		pWriter._WriteUInt2(29, this.stencilGroupPos);
	};

	/**
	 * Write attributes to stream for Page_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Page_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.id);
		pWriter._WriteString2(1, this.name);
		pWriter._WriteString2(2, this.nameU);
		pWriter._WriteBool2(3, this.isCustomName);
		pWriter._WriteBool2(4, this.isCustomNameU);
		pWriter._WriteBool2(5, this.background);
		pWriter._WriteUInt2(6, this.backPage);
		pWriter._WriteDoubleReal2(7, this.viewScale);
		pWriter._WriteDoubleReal2(8, this.viewCenterX);
		pWriter._WriteDoubleReal2(9, this.viewCenterY);
		pWriter._WriteUInt2(10, this.reviewerID);
		pWriter._WriteUInt2(11, this.associatedPage);
	};

	/**
	 * Write children to stream for Page_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Page_Type.prototype.writeChildren = function (pWriter) {
		// Write pageSheet if present
		pWriter.WriteRecordPPTY(0, this.pageSheet);
		
		// Write rel if present
		pWriter.WriteRecordPPTY(1, this.content);
	};

	/**
	 * Write attributes to stream for Connect_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Connect_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.fromSheet);
		pWriter._WriteString2Utf8(1, this.fromCell);
		pWriter._WriteInt2(2, this.fromPart);
		pWriter._WriteUInt2(3, this.toSheet);
		pWriter._WriteString2Utf8(4, this.toCell);
		pWriter._WriteInt2(5, this.toPart);
	};

		/**
	 * Write attributes to stream for Shape_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Shape_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.id);
		pWriter._WriteUChar2(1, this.type);
		pWriter._WriteUInt2(2, this.originalID);
		pWriter._WriteBool2(3, this.del);
		pWriter._WriteUInt2(4, this.masterShape);
		pWriter._WriteString2(5, this.uniqueID);
		pWriter._WriteString2(6, this.nameU);
		pWriter._WriteString2(7, this.name);
		pWriter._WriteBool2(8, this.isCustomName);
		pWriter._WriteBool2(9, this.isCustomNameU);
		pWriter._WriteUInt2(10, this.master);
		pWriter._WriteUInt2(11, this.lineStyle);
		pWriter._WriteUInt2(12, this.fillStyle);
		pWriter._WriteUInt2(13, this.textStyle);
	};

	/**
	 * Write children to stream for Shape_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Shape_Type.prototype.writeChildren = function (pWriter) {
		// Write elements
		for (let i in this.elements) {
			const elem = this.elements[i];
			switch (elem.kind) {
				case AscVisio.c_oVsdxSheetStorageKind.Cell_Type:
					pWriter.WriteRecordPPTY(0, elem);
					break;
				case AscVisio.c_oVsdxSheetStorageKind.Trigger_Type:
					pWriter.WriteRecordPPTY(1, elem);
					break;
				case AscVisio.c_oVsdxSheetStorageKind.Section_Type:
					pWriter.WriteRecordPPTY(2, elem);
					break;
				case AscVisio.c_oVsdxSheetStorageKind.Text_Type:
					pWriter.WriteRecordPPTY(3, elem);
					break;
				case AscVisio.c_oVsdxSheetStorageKind.ForeignData_Type:
					pWriter.WriteRecordPPTY(4, elem);
					break;
			}
		}
		// Write shapes
		if (this.shapes && this.shapes.length > 0) {
			pWriter.StartRecord(5);
			for (let i = 0; i < this.shapes.length; i++) {
				pWriter.WriteRecordPPTY(0, this.shapes[i]);
			}
			pWriter.EndRecord();
		}
	};

	AscVisio.Text_Type.prototype.privateWriteAttributes = undefined;
	/**
	 * Write children to stream for Text_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Text_Type.prototype.writeChildren = function (pWriter) {
		this.elements.forEach(function(elem) {
			switch (elem.kind) {
				case AscVisio.c_oVsdxTextKind.CP:
					pWriter.WriteRecordPPTY(0, elem);
					break;
				case AscVisio.c_oVsdxTextKind.PP:
					pWriter.WriteRecordPPTY(1, elem);
					break;
				case AscVisio.c_oVsdxTextKind.TP:
					pWriter.WriteRecordPPTY(2, elem);
					break;
				case AscVisio.c_oVsdxTextKind.FLD:
					pWriter.WriteRecordPPTY(3, elem);
					break;
				default:
					pWriter.StartRecord(4);
					pWriter.WriteUChar(AscCommon.g_nodeAttributeStart);
					pWriter._WriteString2(0, elem);
					pWriter.WriteUChar(AscCommon.g_nodeAttributeEnd);
					pWriter.EndRecord();
					break;
			}
		});
	};

	/**
	 * Write attributes to stream for RefBy_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.RefBy_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.id);
		pWriter._WriteString2(1, this.t);
	};

	/**
	 * Write attributes to stream for HeaderFooter_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.HeaderFooter_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteDoubleReal2(0, this.headerMargin);
		pWriter._WriteDoubleReal2(1, this.footerMargin);
		pWriter._WriteString2(2, this.headerLeft);
		pWriter._WriteString2(3, this.headerCenter);
		pWriter._WriteString2(4, this.headerRight);
		pWriter._WriteString2(5, this.footerLeft);
		pWriter._WriteString2(6, this.footerCenter);
		pWriter._WriteString2(7, this.footerRight);
		pWriter._WriteString2(8, this.headerFooterFont);
	};

	/**
	 * Write attributes to stream for EventItem_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.EventItem_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.id);
		pWriter._WriteUInt2(1, this.action);
		pWriter._WriteUInt2(2, this.eventCode);
		pWriter._WriteBool2(3, this.enabled);
		pWriter._WriteString2(4, this.target);
		pWriter._WriteString2(5, this.targetArgs);
};
	/**
	 * Write attributes to stream for DocumentSheet_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.DocumentSheet_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteString2(0, this.uniqueID);
		pWriter._WriteString2(1, this.nameU);
		pWriter._WriteString2(2, this.name);
		pWriter._WriteBool2(3, this.isCustomName);
		pWriter._WriteBool2(4, this.isCustomNameU);
		pWriter._WriteUInt2(5, this.lineStyle);
		pWriter._WriteUInt2(6, this.fillStyle);
		pWriter._WriteUInt2(7, this.textStyle);
	};

	/**
	 * Write children to stream for DocumentSheet_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.DocumentSheet_Type.prototype.writeChildren = function (pWriter) {
		// Write elements
		for (let i in this.elements) {
			const elem = this.elements[i];
			switch (elem.kind) {
				case AscVisio.c_oVsdxSheetStorageKind.Cell_Type:
					pWriter.WriteRecordPPTY(0, elem);
					break;
				case AscVisio.c_oVsdxSheetStorageKind.Trigger_Type:
					pWriter.WriteRecordPPTY(1, elem);
					break;
				case AscVisio.c_oVsdxSheetStorageKind.Section_Type:
					pWriter.WriteRecordPPTY(2, elem);
					break;
			}
		}
	};

	/**
	 * Write attributes to stream for FaceName_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.FaceName_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteString2(0, this.nameU);
		pWriter._WriteString2(1, this.unicodeRanges);
		pWriter._WriteString2(2, this.charSets);
		pWriter._WriteString2(3, this.panos);
		pWriter._WriteString2(4, this.panose);
		pWriter._WriteUInt2(5, this.flags);
	};

	/**
	 * Write attributes to stream for ColorEntry_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.ColorEntry_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.ix);
		pWriter._WriteString2(1, this.rgb);
	};

	/**
	 * Write attributes to stream for DocumentSettings_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.DocumentSettings_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.topPage);
		pWriter._WriteUInt2(1, this.defaultTextStyle);
		pWriter._WriteUInt2(2, this.defaultLineStyle);
		pWriter._WriteUInt2(3, this.defaultFillStyle);
		pWriter._WriteUInt2(4, this.defaultGuideStyle);
		pWriter._WriteInt2(5, this.glueSettings);
		pWriter._WriteInt2(6, this.snapSettings);
		pWriter._WriteInt2(7, this.snapExtensions);
		pWriter._WriteInt2(8, this.snapAngles);
		pWriter._WriteBool2(9, this.dynamicGridEnabled);
		pWriter._WriteBool2(10, this.protectStyles);
		pWriter._WriteBool2(11, this.protectShapes);
		pWriter._WriteBool2(12, this.protectBkgnds);
		pWriter._WriteBool2(13, this.protectMasters);
		pWriter._WriteString2(14, this.customMenusFile);
		pWriter._WriteString2(15, this.customToolbarsFile);
		pWriter._WriteString2(16, this.attachedToolbars);
	};

	/**
	 * Write attributes to stream for Section_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Section_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.ix);
		pWriter._WriteString2Utf8(1, this.n);
		pWriter._WriteBool2(2, this.del);
	};

	/**
	 * Write children to stream for Section_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Section_Type.prototype.writeChildren = function (pWriter) {
		// Write elements
		for (let i in this.elements) {
			const elem = this.elements[i];
			switch (elem.kind) {
				case AscVisio.c_oVsdxSheetStorageKind.Cell_Type:
					pWriter.WriteRecordPPTY(0, elem);
					break;
				case AscVisio.c_oVsdxSheetStorageKind.Trigger_Type:
					pWriter.WriteRecordPPTY(1, elem);
					break;
				case AscVisio.c_oVsdxSheetStorageKind.Row_Type:
					pWriter.WriteRecordPPTY(6, elem);
					break;
			}
		}
	};

	/**
	 * Write attributes to stream for Trigger_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Trigger_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteString2Utf8(0, this.n);
	};

	/**
	 * Write children to stream for Trigger_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Trigger_Type.prototype.writeChildren = function (pWriter) {
		// Write refBy
		for (let i = 0; i < this.refBy.length; i++) {
			pWriter.WriteRecordPPTY(0, this.refBy[i]);
		}
	};

	/**
	 * Write attributes to stream for Row_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Row_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.ix);
		pWriter._WriteString2Utf8(1, this.n);
		pWriter._WriteString2(2, this.localName);
		pWriter._WriteString2Utf8(3, this.t);
		pWriter._WriteBool2(4, this.del);
	};

	/**
	 * Write children to stream for Row_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Row_Type.prototype.writeChildren = function (pWriter) {
		// Write elements
		for (let i in this.elements) {
			const elem = this.elements[i];
			switch (elem.kind) {
				case AscVisio.c_oVsdxSheetStorageKind.Cell_Type:
					pWriter.WriteRecordPPTY(0, elem);
					break;
				case AscVisio.c_oVsdxSheetStorageKind.Trigger_Type:
					pWriter.WriteRecordPPTY(1, elem);
					break;
			}
		}
	};

	/**
	 * Write attributes to stream for Cell_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Cell_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteString2Utf8(0, this.n);
		pWriter._WriteString2Utf8(1, this.u);
		pWriter._WriteString2Utf8(2, this.e);
		pWriter._WriteString2Utf8(3, this.f);
		pWriter._WriteString2Utf8(4, this.v);
	};

	/**
	 * Write children to stream for Cell_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Cell_Type.prototype.writeChildren = function (pWriter) {
		// Write refBy
		for (let i = 0; i < this.refBy.length; i++) {
			pWriter.WriteRecordPPTY(0, this.refBy[i]);
		}
	};
	/**
	 * Write attributes to stream for ForeignData_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.ForeignData_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUChar2(0, this.foreignType);
		pWriter._WriteUInt2(1, this.objectType);
		pWriter._WriteBool2(2, this.showAsIcon);
		pWriter._WriteDoubleReal2(3, this.objectWidth);
		pWriter._WriteDoubleReal2(4, this.objectHeight);
		pWriter._WriteDoubleReal2(5, this.extentX);
		pWriter._WriteDoubleReal2(6, this.extentY);
		pWriter._WriteUChar2(7, this.compressionType);
		pWriter._WriteDoubleReal2(8, this.compressionLevel);
	};

	/**
	 * Write children to stream for ForeignData_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.ForeignData_Type.prototype.writeChildren = function (pWriter) {
		// Write image data
		if (null !== this.mediaFilename) {
			pWriter.StartRecord(1);
			pWriter.WriteString2(this.mediaFilename);
			pWriter.EndRecord();
		}
		if (null !== this.oleFilename) {
			pWriter.StartRecord(2);
			pWriter.WriteString2(this.oleFilename);
			pWriter.EndRecord();
		}
	};

	AscVisio.CSolutions.prototype.privateWriteAttributes = undefined;
	/**
	 * Write attributes to stream for CSolutions
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CSolutions.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteString2(0, this.current);
	};

	/**
	 * Write children to stream for CSolutions
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.CSolutions.prototype.writeChildren = function (pWriter) {
		// Write solution objects
		if (this.solution && this.solution.length > 0) {
			for (let i = 0; i < this.solution.length; i++) {
				pWriter.WriteRecordPPTY(0, this.solution[i]);
			}
		}
	};
	/**
	 * Write attributes to stream for Issue_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Issue_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.id);
		pWriter._WriteBool2(1, this.ignored);
	};

	/**
	 * Write children to stream for Issue_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Issue_Type.prototype.writeChildren = function (pWriter) {
		// Write issueTargets
		if (this.issueTargets && this.issueTargets.length > 0) {
			pWriter.StartRecord(0);
			for (let i = 0; i < this.issueTargets.length; i++) {
				pWriter.WriteRecordPPTY(0, this.issueTargets[i]);
			}
			pWriter.EndRecord();
		}
		
		// Write ruleInfo
		pWriter.WriteRecordPPTY(1, this.ruleInfo);
	};


	/**
	 * Write attributes to stream for Master_Type
	 * 
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Master_Type.prototype.privateWriteAttributes = function (pWriter) {
		pWriter._WriteUInt2(0, this.id);
		pWriter._WriteString2(1, this.name);
		pWriter._WriteString2(2, this.nameU);
		pWriter._WriteString2(3, this.baseID);
		pWriter._WriteString2(4, this.uniqueID);
		pWriter._WriteBool2(5, this.matchByName);
		pWriter._WriteBool2(6, this.isCustomName);
		pWriter._WriteBool2(7, this.isCustomNameU);
		pWriter._WriteUInt2(8, this.iconSize);
		pWriter._WriteUInt2(9, this.patternFlags);
		pWriter._WriteString2(10, this.prompt);
		pWriter._WriteBool2(11, this.hidden);
		pWriter._WriteBool2(12, this.iconUpdate);
		pWriter._WriteUInt2(13, this.alignName);
		pWriter._WriteUInt2(14, this.masterType);
	};

	/**
	 * Write children to stream for Master_Type
	 *
	 * @param {CBinaryFileWriter} pWriter - The binary writer
	 */
	AscVisio.Master_Type.prototype.writeChildren = function (pWriter) {
		// Write pageSheet
		pWriter.WriteRecordPPTY(0, this.pageSheet);
		// Write icon
		pWriter.WriteRecordPPTY(1, this.icon);
		// Write rel
		pWriter.WriteRecordPPTY(2, this.content);
	};

	window['AscVisio']  = window['AscVisio'] || {};

	window['AscVisio'].BinaryVSDYWriter = BinaryVSDYWriter;

})(window, window.document);

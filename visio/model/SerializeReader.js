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
	function BinaryVSDYLoader()
	{
		this.stream = null;
		this.document = null;

		this.ImageMapChecker = {};
		this.binaryPPTYLoader = new AscCommon.BinaryPPTYLoader();

		this.Load = function(base64_ppty, document)
		{
			this.document = document;
			this.DrawingDocument = null;
			if(document)
			{
				this.DrawingDocument = document.DrawingDocument;
			}
			else
			{
				this.DrawingDocument = null;
			}
			this.ImageMapChecker = {};

			var isBase64 = typeof base64_ppty === 'string';
			var srcLen = isBase64 ? base64_ppty.length : base64_ppty.length;
			var nWritten = 0;

			var index = 0;
			var read_main_prop = "";
			while (true)
			{
				var _c = isBase64 ? base64_ppty.charCodeAt(index) : base64_ppty[index];
				if (_c == ";".charCodeAt(0))
					break;

				read_main_prop += String.fromCharCode(_c);
				index++;
			}
			index++;

			if ("VSDY" != read_main_prop)
				return false;

			read_main_prop = "";
			while (true)
			{
				var _c = isBase64 ? base64_ppty.charCodeAt(index) : base64_ppty[index];
				if (_c == ";".charCodeAt(0))
					break;

				read_main_prop += String.fromCharCode(_c);
				index++;
			}
			index++;

			var _version_num_str = read_main_prop.substring(1);
			var version = 1;
			if(_version_num_str.length > 0)
			{
				version = _version_num_str - 0;
			}
			read_main_prop = "";
			while (true)
			{
				var _c = isBase64 ? base64_ppty.charCodeAt(index) : base64_ppty[index];
				if (_c == ";".charCodeAt(0))
					break;

				read_main_prop += String.fromCharCode(_c);
				index++;
			}
			index++;

			const c_nVersionNoBase64 = 1;//todo Asc.c_nVersionNoBase64
			if (c_nVersionNoBase64 !== version) {
				var dstLen_str = read_main_prop;

				var dstLen = parseInt(dstLen_str);
				var memoryData = AscCommon.Base64.decode(base64_ppty, false, dstLen, index);
				this.stream = new AscCommon.FileStream(memoryData, memoryData.length);
			} else {
				this.stream = new AscCommon.FileStream();
				this.stream.obj    = null;
				this.stream.data   = base64_ppty;
				this.stream.size   = base64_ppty.length;
				//skip header
				// this.stream.EnterFrame(index);
				this.stream.Seek2(index);
			}

			this.binaryPPTYLoader.stream = this.stream;
			this.binaryPPTYLoader.presentation = this.document;

			this.LoadDocument();
			// if(AscFonts.IsCheckSymbols)
			// {
			// 	var bLoad = AscCommon.g_oIdCounter.m_bLoad;
			// 	AscCommon.g_oIdCounter.Set_Load(false);
			// 	for(var nField = 0; nField < this.fields.length; ++nField)
			// 	{
			// 		var oField = this.fields[nField];
			// 		var sValue = oField.private_GetString();
			// 		if(typeof sValue === "string" && sValue.length > 0)
			// 		{
			// 			AscFonts.FontPickerByCharacter.getFontsByString(sValue);
			// 		}
			// 	}
			// 	AscCommon.g_oIdCounter.Set_Load(bLoad);
			// }
			// this.fields.length = 0;
			// AscFormat.checkPlaceholdersText();

			this.ImageMapChecker = null;
		};

		/**
		 * Loads a Visio document from binary stream and processes its components.
		 * 
		 * @param  {BinaryVSDYLoader} pReader - The binary reader
		 */
		this.LoadDocument = function()
		{
			let res = c_oSerConstants.ReadOk;
			const stream = this.stream;
			
			// Define table types
			const TABLE_TYPES = {
				DOCUMENT: 1,
				APP: 2,
				CORE: 3,
				CUSTOM_PROPERTIES: 4
			};
			
			// Read the main table length
			const mtLen = stream.GetUChar();
			
			// Read and process sections immediately
			for (let i = 0; i < mtLen; ++i) {
				const sectionType = stream.GetUChar();
				const sectionOffset = stream.GetULong();
				
				// Store current position to return after processing section
				const currentPos = stream.GetCurPos();
				
				// Seek to section and process based on type
				res = stream.Seek2(sectionOffset);
				if (c_oSerConstants.ReadOk !== res) {
					return res;
				}
				
				switch (sectionType) {
					case TABLE_TYPES.APP:
						this.document.app = new AscCommon.CApp();
						this.document.app.fromStream(stream);
						break;
					
					case TABLE_TYPES.CORE:
						this.document.core = new AscCommon.CCore();
						this.document.core.fromStream(stream);
						break;
					
					case TABLE_TYPES.CUSTOM_PROPERTIES:
						if (this.document.customProperties) {
							this.document.customProperties.fromStream(stream);
						}
						break;
					
					case TABLE_TYPES.DOCUMENT:
						stream.Skip2(1);//type
						this.document.fromPPTY(this);
						break;
				}
				
				// Return to position after section type and offset
				res = stream.Seek2(currentPos);
				if (c_oSerConstants.ReadOk !== res) {
					return res;
				}
			}
			
			return res;
		};
	}

	// /**
	//  * Loads a Visio document from a binary stream
	//  *
	//  * @param  {BinaryVSDYLoader} pReader - The binary reader
	//  */
	// AscVisio.CVisioDocument.prototype.fromPPTY = function(pReader) {
	// 	const stream = pReader.stream;
	// 	stream.Skip2(1);
	//
	// 	const startPos = stream.GetCurPos();
	// 	const recordSize = stream.GetLong();
	// 	const endPos = startPos + recordSize + 4;
	//
	// 	while (stream.GetCurPos() < endPos) {
	// 		const recordType = stream.GetUChar();
	//
	// 		// this.readChild(recordType, pReader);
	// 		const res = this.readChild(recordType, pReader);
	// 			if (false === res) {
	// 				stream.SkipRecord();
	// 			}
	// 	}
	//
	// 	stream.Seek2(endPos);
	// };

	AscVisio.CVisioDocument.prototype.readAttribute = undefined;
	/**
	 * Read child elements from stream for CVisioDocument
	 *
	 * @param  {BinaryVSDYLoader} pReader - The binary reader
	 * @param {number} elementType - The type of child element
	 */
	AscVisio.CVisioDocument.prototype.readChild = function(elementType, pReader) {
		const stream = pReader.stream;
		let t = this;
		switch (elementType) {
			case 0:
				if (!this.documentSettings) {
					this.documentSettings = new AscVisio.DocumentSettings_Type();
				}
				this.documentSettings.fromPPTY(pReader);
				break;
				
			case 1:
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						const stream = pReader.stream;
						if (elementType === 0) {
							const colorEntry = new AscVisio.ColorEntry_Type();
							colorEntry.fromPPTY(pReader);
							t.colors.push(colorEntry);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
				
			case 2:
				//CFaceNames
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						const stream = pReader.stream;
						if (elementType === 0) {
							const faceName = new AscVisio.FaceName_Type();
							faceName.fromPPTY(pReader);
							t.faceNames.push(faceName);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
				
			case 3:
				//CStyleSheets
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						const stream = pReader.stream;
						if (elementType === 0) {
							const styleSheet = new AscVisio.StyleSheet_Type();
							styleSheet.fromPPTY(pReader);
							t.styleSheets.push(styleSheet);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
				
			case 4:
				this.documentSheet = new AscVisio.DocumentSheet_Type();
				this.documentSheet.fromPPTY(pReader);
				break;
				
			case 5:
				//CEventList
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						const stream = pReader.stream;
						if (elementType === 0) {
							const styleSheet = new AscVisio.EventItem_Type();
							styleSheet.fromPPTY(pReader);
							t.eventList.push(styleSheet);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
				
			case 6:
				if (!this.headerFooter) {
					this.headerFooter = new AscVisio.HeaderFooter_Type();
				}
				this.headerFooter.fromPPTY(pReader);
				break;
				
			case 7:
				if (!this.masters) {
					this.masters = new AscVisio.CMasters();
				}
				this.masters.fromPPTY(pReader);
				break;
				
			case 8:
				if (!this.pages) {
					this.pages = new AscVisio.CPages();
				}
				this.pages.fromPPTY(pReader);
				break;
				
			case 9:
				//DataConnections
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						const stream = pReader.stream;
						if (elementType === 0) {
							t.dataConnections = new AscVisio.CDataConnections();
							t.dataConnections.fromPPTY(pReader);
							return true;
						}
						return false;
					}
				}, pReader);
				break;

			case 10:
				//Recordsets
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						const stream = pReader.stream;
						if (elementType === 0) {
							t.dataRecordSets = new AscVisio.CDataRecordSets();
							t.dataRecordSets.fromPPTY(pReader);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
				
			case 11:
				//Solutions
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						const stream = pReader.stream;
						if (elementType === 0) {
							if (!t.solutions) {
								t.solutions = new AscVisio.CSolutions();
							}
							t.solutions.fromPPTY(pReader);
							return true;
						}
						return false;
					}
				}, pReader);

				break;
				
			case 12:
				if (!this.validation) {
					this.validation = new AscVisio.CValidation();
				}
				this.validation.fromPPTY(pReader);
				break;

			case 13:
				//Comments
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						const stream = pReader.stream;
						if (elementType === 0) {
							t.commentsPart = new AscVisio.CComments();
							t.commentsPart.fromPPTY(pReader);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
				
			case 14:
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						if (elementType === 0 && !t.windows) {
							t.windows = new AscVisio.CWindows();
							t.windows.fromPPTY(pReader);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
				
			case 15:
				const len =stream.GetULong(4);//len
				if (len > 0) {
					//Themes
					let theme = pReader.binaryPPTYLoader.ReadTheme()
					this.themes.push(theme);
				}
				break;

			// 	//todo VbaProject
			// case 16:
			// 	break;
			//
			// //todo JsaProject
			// case 17:
			// 	break;
			default:
				return false;
		}
	};

	/**
	 * Read attributes from stream for DocumentSettings_Type
	 * 
	 * @param  {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.DocumentSettings_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.topPage = stream.GetULong();
				break;
			case 1:
				this.defaultTextStyle = stream.GetULong();
				break;
			case 2:
				this.defaultLineStyle = stream.GetULong();
				break;
			case 3:
				this.defaultFillStyle = stream.GetULong();
				break;
			case 4:
				this.defaultGuideStyle = stream.GetULong();
				break;
			case 5:
				this.glueSettings = stream.GetLong();
				break;
			case 6:
				this.snapSettings = stream.GetLong();
				break;
			case 7:
				this.snapExtensions = stream.GetLong();
				break;
			case 8:
				this.snapAngles = stream.GetLong();
				break;
			case 9:
				this.dynamicGridEnabled = stream.GetBool();
				break;
			case 10:
				this.protectStyles = stream.GetBool();
				break;
			case 11:
				this.protectShapes = stream.GetBool();
				break;
			case 12:
				this.protectBkgnds = stream.GetBool();
				break;
			case 13:
				this.protectMasters = stream.GetBool();
				break;
			case 14:
				this.customMenusFile = stream.GetString2();
				break;
			case 15:
				this.customToolbarsFile = stream.GetString2();
				break;
			case 16:
				this.attachedToolbars = stream.GetString2();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read attributes from stream for ColorEntry_Type
	 * 
	 * @param  {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.ColorEntry_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.ix = stream.GetULong();
				break;
			case 1:
				this.rgb = stream.GetString2();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read attributes from stream for FaceName_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.FaceName_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.nameU = stream.GetString2();
				break;
			case 1:
				this.unicodeRanges = stream.GetString2();
				break;
			case 2:
				this.charSets = stream.GetString2();
				break;
			case 3:
				this.panos = stream.GetString2();
				break;
			case 4:
				this.panose = stream.GetString2();
				break;
			case 5:
				this.flags = stream.GetULong();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read attributes from stream for StyleSheet_Type
	 * 
	 * @param  {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.StyleSheet_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.id = stream.GetULong();
				break;
			case 1:
				this.nameU = stream.GetString2();
				break;
			case 2:
				this.name = stream.GetString2();
				break;
			case 3:
				this.isCustomName = stream.GetBool();
				break;
			case 4:
				this.isCustomNameU = stream.GetBool();
				break;
			case 5:
				this.lineStyle = stream.GetULong();
				break;
			case 6:
				this.fillStyle = stream.GetULong();
				break;
			case 7:
				this.textStyle = stream.GetULong();
				break;
			default:
				return false;
		}
		
		return true;
	};
	
	/**
	 * Read child elements from stream for StyleSheet_Type
	 * 
	 * @param  {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.StyleSheet_Type.prototype.readChild = function(elementType, pReader) {
		const stream = pReader.stream;
		switch (elementType) {
			case 0: {
				const cell = new AscVisio.Cell_Type();
				cell.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(cell);
				this.elements[key] = cell;
				break;
			}
			case 1: {
				const trigger = new AscVisio.Trigger_Type();
				trigger.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(trigger);
				this.elements[key] = trigger;
				break;
			}
			case 2: {
				const section = new AscVisio.Section_Type();
				section.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(section);
				this.elements[key] = section;
				break;
			}
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read attributes from stream for Cell_Type
	 * 
	 * @param  {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.Cell_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.n = stream.GetString2Utf8();
				break;
			case 1:
				this.u = stream.GetString2Utf8();
				break;
			case 2:
				this.e = stream.GetString2Utf8();
				break;
			case 3:
				this.f = stream.GetString2Utf8();
				break;
			case 4:
				this.v = stream.GetString2Utf8();
				break;
			default:
				return false;
		}
		
		return true;
	};
	/**
	 * Read child elements from stream for Cell_Type
	 *
	 * @param  {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.Cell_Type.prototype.readChild = function(elementType, pReader) {
		const stream = pReader.stream;
		if (elementType === 0) {
			const refBy = new AscVisio.RefBy_Type();
			refBy.fromPPTY(pReader);
			this.refBy.push(refBy);
			return true;
		}

		return false;
	};

	/**
	 * Read attributes from stream for EventItem_Type
	 * 
	 * @param  {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.EventItem_Type.prototype.readAttribute = function(attrType, pReader) {
		switch (attrType) {
			case 0:
				this.id = pReader.stream.GetULong();
				return true;
			case 1:
				this.action = pReader.stream.GetULong();
				return true;
			case 2:
				this.eventCode = pReader.stream.GetULong();
				return true;
			case 3:
				this.enabled = pReader.stream.GetBool();
				return true;
			case 4:
				this.target = pReader.stream.GetString2();
				return true;
			case 5:
				this.targetArgs = pReader.stream.GetString2();
				return true;
		}
		return false;
	};

	/**
	 * Read attributes from stream for HeaderFooter_Type
	 * 
	 * @param  {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.HeaderFooter_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.headerMargin = stream.GetDouble();
				break;
			case 1:
				this.footerMargin = stream.GetDouble();
				break;
			case 2:
				this.headerLeft = stream.GetString2();
				break;
			case 3:
				this.headerCenter = stream.GetString2();
				break;
			case 4:
				this.headerRight = stream.GetString2();
				break;
			case 5:
				this.footerLeft = stream.GetString2();
				break;
			case 6:
				this.footerCenter = stream.GetString2();
				break;
			case 7:
				this.footerRight = stream.GetString2();
				break;
			case 8:
				this.headerFooterFont = stream.GetString2();
				break;
			default:
				return false;
		}
		
		return true;
	};

	AscVisio.CMasters.prototype.readAttribute = undefined;
	/**
	 * Read child elements from stream for CMasters
	 * 
	 * @param  {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.CMasters.prototype.readChild = function(elementType, pReader) {
		const stream = pReader.stream;
		if (elementType === 0) {
			const master = new AscVisio.Master_Type();
			master.fromPPTY(pReader);
			this.master.push(master);
			return true;
		}
		
		return false;
	};

	/**
	 * Read attributes from stream for Master_Type
	 * 
	 * @param  {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.Master_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.id = stream.GetULong();
				break;
			case 1:
				this.name = stream.GetString2();
				break;
			case 2:
				this.nameU = stream.GetString2();
				break;
			case 3:
				this.baseID = stream.GetString2();
				break;
			case 4:
				this.uniqueID = stream.GetString2();
				break;
			case 5:
				this.matchByName = stream.GetBool();
				break;
			case 6:
				this.isCustomName = stream.GetBool();
				break;
			case 7:
				this.isCustomNameU = stream.GetBool();
				break;
			case 8:
				this.iconSize = stream.GetULong();
				break;
			case 9:
				this.patternFlags = stream.GetULong();
				break;
			case 10:
				this.prompt = stream.GetString2();
				break;
			case 11:
				this.hidden = stream.GetBool();
				break;
			case 12:
				this.iconUpdate = stream.GetBool();
				break;
			case 13:
				this.alignName = stream.GetULong();
				break;
			case 14:
				this.masterType = stream.GetULong();
				break;
			default:
				return false;
		}
		
		return true;
	};
	
	/**
	 * Read child elements from stream for Master_Type
	 * 
	 * @param  {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.Master_Type.prototype.readChild = function(elementType, pReader) {
		const stream = pReader.stream;
		switch (elementType) {
			case 0: {
				this.pageSheet = new AscVisio.PageSheet_Type();
				this.pageSheet.fromPPTY(pReader);
				break;
			}
			case 1: {
				this.icon = new AscVisio.Icon_Type();
				this.icon.fromPPTY(pReader);
				break;
			}
			case 2: {
				const masterContents = new AscVisio.CMasterContents();
				masterContents.fromPPTY(pReader);
				this.content = masterContents;
				break;
			}
			default:
				return false;
		}

		return true;
	};

	AscVisio.CMasterContents.prototype.readAttribute = undefined;
	/**
	 * Read child elements from stream for CMasterContents
	 *
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.CMasterContents.prototype.readChild = function(elementType, pReader) {
		let t = this;
		switch (elementType) {
			case 0: {
				// Read Shapes
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						if (elementType === 0) {
							const shape = new AscVisio.Shape_Type();
							shape.fromPPTY(pReader);
							t.shapes.push(shape);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
			}
			case 1: {
				// Read Connects
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						if (elementType === 0) {
							const connect = new AscVisio.Connect_Type();
							connect.fromPPTY(pReader);
							t.connects.push(connect);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
			}
			default:
				return false;
		}
		
		return true;
	};
	
	/**
	 * Read attributes from stream for Connect_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.Connect_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.fromSheet = stream.GetULong();
				break;
			case 1:
				this.fromCell = stream.GetString2Utf8();
				break;
			case 2:
				this.fromPart = stream.GetLong();
				break;
			case 3:
				this.toSheet = stream.GetULong();
				break;
			case 4:
				this.toCell = stream.GetString2Utf8();
				break;
			case 5:
				this.toPart = stream.GetLong();
				break;
			default:
				return false;
		}
		
		return true;
	};

	AscVisio.CPages.prototype.readAttribute = undefined;
	/**
	 * Read child elements from stream for CPages
	 * 
	 * @param  {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.CPages.prototype.readChild = function(elementType, pReader) {
		const stream = pReader.stream;
		if (elementType === 0) {
			const page = new AscVisio.Page_Type();
			page.fromPPTY(pReader);
			this.page.push(page);
			return true;
		}
		
		return false;
	};

	/**
	 * Read attributes from stream for Page_Type
	 * 
	 * @param  {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.Page_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.id = stream.GetULong();
				break;
			case 1:
				this.name = stream.GetString2();
				break;
			case 2:
				this.nameU = stream.GetString2();
				break;
			case 3:
				this.isCustomName = stream.GetBool();
				break;
			case 4:
				this.isCustomNameU = stream.GetBool();
				break;
			case 5:
				this.background = stream.GetBool();
				break;
			case 6:
				this.backPage = stream.GetULong();
				break;
			case 7:
				this.viewScale = stream.GetDouble();
				break;
			case 8:
				this.viewCenterX = stream.GetDouble();
				break;
			case 9:
				this.viewCenterY = stream.GetDouble();
				break;
			case 10:
				this.reviewerID = stream.GetULong();
				break;
			case 11:
				this.associatedPage = stream.GetULong();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read child elements from stream for Page_Type
	 * 
	 * @param  {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.Page_Type.prototype.readChild = function(elementType, pReader) {
		switch (elementType) {
			case 0: {
				this.pageSheet = new AscVisio.PageSheet_Type();
				this.pageSheet.fromPPTY(pReader);
				break;
			}
			case 1: {
				const pageContents = new AscVisio.CPageContents();
				pageContents.fromPPTY(pReader);
				this.content = pageContents;
				break;
			}
			default:
				return false;
		}

		return true;
	};

	/**
	 * Read attributes from stream for PageSheet_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.PageSheet_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.uniqueID = stream.GetString2();
				break;
			case 1:
				this.lineStyle = stream.GetULong();
				break;
			case 2:
				this.fillStyle = stream.GetULong();
				break;
			case 3:
				this.textStyle = stream.GetULong();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read child elements from stream for PageSheet_Type
	 *
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.PageSheet_Type.prototype.readChild = function(elementType, pReader) {
		switch (elementType) {
			case 0: {
				const cell = new AscVisio.Cell_Type();
				cell.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(cell);
				this.elements[key] = cell;
				break;
			}
			case 1: {
				const trigger = new AscVisio.Trigger_Type();
				trigger.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(trigger);
				this.elements[key] = trigger;
				break;
			}
			case 2: {
				const section = new AscVisio.Section_Type();
				section.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(section);
				this.elements[key] = section;
				break;
			}
			default:
				return false;
		}

		return true;
	};

	// /**
	//  * Read attributes from stream for Comment_Type
	//  *
	//  * @param  {number} attrType - The type of attribute
	//  * @param {BinaryVSDYLoader} pReader - The binary reader
	//  * @returns {boolean} - True if attribute was handled, false otherwise
	//  */
	// AscVisio.Comment_Type.prototype.readAttribute = function(attrType, pReader) {
	// 	const stream = pReader.stream;
	// 	switch (attrType) {
	// 		case 0:
	// 			this.id = stream.GetULong();
	// 			break;
	// 		case 1:
	// 			this.authorName = stream.GetString2();
	// 			break;
	// 		case 2:
	// 			this.authorInitials = stream.GetString2();
	// 			break;
	// 		case 3:
	// 			this.text = stream.GetString2();
	// 			break;
	// 		case 4:
	// 			this.date = stream.GetString2();
	// 			break;
	// 		case 5:
	// 			this.reviewerID = stream.GetULong();
	// 			break;
	// 		case 6:
	// 			this.shapeId = stream.GetULong();
	// 			break;
	// 		case 7:
	// 			this.pageId = stream.GetULong();
	// 			break;
	// 		default:
	// 			return false;
	// 	}
	//
	// 	return true;
	// };

	// /**
	//  * Read attributes from stream for Recordset_Type
	//  *
	//  * @param  {number} attrType - The type of attribute
	//  * @param {BinaryVSDYLoader} pReader - The binary reader
	//  * @returns {boolean} - True if attribute was handled, false otherwise
	//  */
	// AscVisio.Recordset_Type.prototype.readAttribute = function(attrType, pReader) {
	// 	const stream = pReader.stream;
	// 	switch (attrType) {
	// 		case 0:
	// 			this.id = stream.GetULong();
	// 			break;
	// 		case 1:
	// 			this.name = stream.GetString2();
	// 			break;
	// 		case 2:
	// 			this.dataSource = stream.GetString2();
	// 			break;
	// 		case 3:
	// 			this.command = stream.GetString2();
	// 			break;
	// 		case 4:
	// 			this.options = stream.GetULong();
	// 			break;
	// 		case 5:
	// 			this.timeRefreshed = stream.GetString2();
	// 			break;
	// 		case 6:
	// 			this.nextRowID = stream.GetULong();
	// 			break;
	// 		default:
	// 			return false;
	// 	}
	//
	// 	return true;
	// };

	/**
	 * Read attributes from stream for CComments
	 * 
	 * @param  {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.CComments.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.showCommentTags = pReader.stream.GetBool();
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read child elements from stream for CComments
	 * 
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.CComments.prototype.readChild = function(elementType, pReader)
	{
		let handled = true;
		let t = this;
		switch (elementType)
		{
			case 0:
				// Read author list elements
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						const stream = pReader.stream;
						if (elementType === 0) {
							const authorEntry = new AscVisio.AuthorEntry_Type();
							authorEntry.fromPPTY(pReader);
							t.authorList.push(authorEntry);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
			case 1:
				// Read comment list elements
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						const stream = pReader.stream;
						if (elementType === 0) {
							const commentEntry = new AscVisio.CommentEntry_Type();
							commentEntry.fromPPTY(pReader);
							t.commentList.push(commentEntry);
							return true;
						}
						return false;
					}
				}, pReader);

				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	AscVisio.CSolutions.prototype.readAttribute = undefined;
	/**
	 * Read child elements from stream for CSolutions
	 * 
	 * @param  {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.CSolutions.prototype.readChild = function(elementType, pReader) {
		const stream = pReader.stream;
		if (elementType === 0) {
			const solution = new AscVisio.Solution_Type();
			solution.fromPPTY(pReader);
			this.solution.push(solution);
			return true;
		}
		
		return false;
	};

	/**
	 * Read attributes from stream for Solution_Type
	 * 
	 * @param  {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.Solution_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.name = stream.GetString2();
				break;
			default:
				return false;
		}
		
		return true;
	};
	/**
	 * Read child elements from stream for Solution_Type
	 *
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.Solution_Type.prototype.readChild = function(elementType, pReader) {
		let res = true;
		let t = this;
		switch (elementType) {
			case 0: {
				pReader.stream.Skip2(4);//len
				pReader.stream.Skip2(1);//type
				const len = pReader.stream.GetULong();
				this.content = pReader.stream.GetBufferUint8(len);
				break;
			}
			default: {
				res = false;
				break;
			}
		}

		return res;
	}

	AscVisio.CValidation.prototype.readAttribute = undefined;
	/**
	 * Read child elements from stream for CValidation
	 * 
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.CValidation.prototype.readChild = function(elementType, pReader) {
		let res = true;
		let t = this;
		switch (elementType) {
			case 0: {
				this.validationProperties = new AscVisio.ValidationProperties_Type();
				this.validationProperties.fromPPTY(pReader);
				break;
			}
			case 1: {
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						const stream = pReader.stream;
						if (elementType === 0) {
							let elem = new AscVisio.RuleSet_Type();
							elem.fromPPTY(pReader);
							t.ruleSets.push(elem);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
			}
			case 2: {
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						const stream = pReader.stream;
						if (elementType === 0) {
							let elem = new AscVisio.Issue_Type();
							elem.fromPPTY(pReader);
							t.issues.push(elem);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
			}
			default: {
				res = false;
				break;
			}
		}
		
		return res;
	};

	/**
	 * Read attributes from stream for CWindows
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.CWindows.prototype.readAttribute = function(attrType, pReader) {
		let stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.clientWidth = stream.GetULong();
				return true;
			case 1:
				this.clientHeight = stream.GetULong();
				return true;
		}
		
		return false;
	};

	/**
	 * Read child elements from stream for CWindows
	 *
	 * @param  {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.CWindows.prototype.readChild = function(elementType, pReader) {
		let stream = pReader.stream;
		if (elementType === 0) {
			let window = new AscVisio.Window_Type();
			window.fromPPTY(pReader);
			this.window.push(window);
			return true;
		}
		
		return false;
	};

	/**
	 * Read attributes from stream for Window_Type
	 * 
	 * @param  {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.Window_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.id = stream.GetULong();
				return true;
			case 1:
				this.windowType = stream.GetUChar();
				return true;
			case 2:
				this.windowState = stream.GetULong();
				return true;
			case 3:
				this.windowLeft = stream.GetLong();
				return true;
			case 4:
				this.windowTop = stream.GetLong();
				return true;
			case 5:
				this.windowWidth = stream.GetULong();
				return true;
			case 6:
				this.windowHeight = stream.GetULong();
				return true;
			case 7:
				this.containerType = stream.GetUChar();
				return true;
			case 8:
				this.container = stream.GetULong();
				return true;
			case 9:
				this.page = stream.GetULong();
				return true;
			case 10:
				this.sheet = stream.GetULong();
				return true;
			case 11:
				this.viewScale = stream.GetDouble();
				return true;
			case 12:
				this.viewCenterX = stream.GetDouble();
				return true;
			case 13:
				this.viewCenterY = stream.GetDouble();
				return true;
			case 14:
				this.document = stream.GetString2();
				return true;
			case 15:
				this.parentWindow = stream.GetULong();
				return true;
			case 16:
				this.readOnly = stream.GetBool();
				return true;
			case 17:
				this.showRulers = stream.GetBool();
				return true;
			case 18:
				this.showGrid = stream.GetBool();
				return true;
			case 19:
				this.showPageBreaks = stream.GetBool();
				return true;
			case 20:
				this.showGuides = stream.GetBool();
				return true;
			case 21:
				this.showConnectionPoints = stream.GetBool();
				return true;
			case 22:
				this.glueSettings = stream.GetULong();
				return true;
			case 23:
				this.snapSettings = stream.GetULong();
				return true;
			case 24:
				this.snapExtensions = stream.GetULong();
				return true;
			case 25:
				this.snapAngles = stream.GetBool();
				return true;
			case 26:
				this.dynamicGridEnabled = stream.GetBool();
				return true;
			case 27:
				this.tabSplitterPos = stream.GetDouble();
				return true;
			case 28:
				this.stencilGroup = stream.GetULong();
				return true;
			case 29:
				this.stencilGroupPos = stream.GetULong();
				return true;
		}
		
		return false;
	};

	/**
	 * Read attributes from stream for DocumentSheet_Type
	 * 
	 * @param  {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.DocumentSheet_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.uniqueID = stream.GetString2();
				break;
			case 1:
				this.nameU = stream.GetString2();
				break;
			case 2:
				this.name = stream.GetString2();
				break;
			case 3:
				this.isCustomName = stream.GetBool();
				break;
			case 4:
				this.isCustomNameU = stream.GetBool();
				break;
			case 5:
				this.lineStyle = stream.GetULong();
				break;
			case 6:
				this.fillStyle = stream.GetULong();
				break;
			case 7:
				this.textStyle = stream.GetULong();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read child elements from stream for DocumentSheet_Type
	 *
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.DocumentSheet_Type.prototype.readChild = function(elementType, pReader) {
		const stream = pReader.stream;
		switch (elementType) {
			case 0: {
				const cell = new AscVisio.Cell_Type();
				cell.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(cell);
				this.elements[key] = cell;
				break;
			}
			case 1: {
				const trigger = new AscVisio.Trigger_Type();
				trigger.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(trigger);
				this.elements[key] = trigger;
				break;
			}
			case 2: {
				const selection = new AscVisio.Section_Type();
				selection.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(selection);
				this.elements[key] = selection;
				break;
			}
			default:
				return false;
		}

		return true;
	};

	/**
	 * Read attributes from stream for Trigger_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.Trigger_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.n = stream.GetString2Utf8();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read child elements from stream for Trigger_Type
	 *
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.Trigger_Type.prototype.readChild = function(elementType, pReader) {
		if (elementType === 0) {
			const refBy = new AscVisio.RefBy_Type();
			refBy.fromPPTY(pReader);
			this.refBy.push(refBy);
			return true;
		}

		return false;
	};

	/**
	 * Read attributes from stream for Section_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.Section_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.ix = stream.GetULong();
				break;
			case 1:
				this.n = stream.GetString2Utf8();
				break;
			case 2:
				this.del = stream.GetBool();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read child elements from stream for Section_Type
	 *
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.Section_Type.prototype.readChild = function(elementType, pReader) {
		switch (elementType) {
			case 0: {
				const cell = new AscVisio.Cell_Type();
				cell.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(cell);
				this.elements[key] = cell;
				break;
			}
			case 1: {
				const trigger = new AscVisio.Trigger_Type();
				trigger.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(trigger);
				this.elements[key] = trigger;
				break;
			}
			case 6: {
				const row = new AscVisio.Row_Type();
				row.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(row);
				this.elements[key] = row;
				break;
			}
			default:
				return false;
		}

		return true;
	};

	/**
	 * Read attributes from stream for Row_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.Row_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.ix = stream.GetULong();
				break;
			case 1:
				this.n = stream.GetString2Utf8();
				break;
			case 2:
				this.localName = stream.GetString2();
				break;
			case 3:
				this.t = stream.GetString2Utf8();
				break;
			case 4:
				this.del = stream.GetBool();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read child elements from stream for Row_Type
	 *
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.Row_Type.prototype.readChild = function(elementType, pReader) {
		switch (elementType) {
			case 0: {
				const cell = new AscVisio.Cell_Type();
				cell.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(cell);
				this.elements[key] = cell;
				break;
			}
			case 1: {
				const trigger = new AscVisio.Trigger_Type();
				trigger.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(trigger);
				this.elements[key] = trigger;
				break;
			}
			default:
				return false;
		}

		return true;
	};

	AscVisio.CPageContents.prototype.readAttribute = undefined;
	/**
	 * Read child elements from stream for CPageContents
	 *
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.CPageContents.prototype.readChild = function(elementType, pReader) {
		let t = this;
		switch (elementType) {
			case 0: {
				// Read Shapes
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						const stream = pReader.stream;
						if (elementType === 0) {
							const shape = new AscVisio.Shape_Type();
							shape.fromPPTY(pReader);
							t.shapes.push(shape);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
			}
			case 1: {
				// Read Connects
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						if (elementType === 0) {
							const connect = new AscVisio.Connect_Type();
							connect.fromPPTY(pReader);
							t.connects.push(connect);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
			}
			// case 2: {
			// 	// Handle Master relationships
			// 	const stream = pReader.stream;
			// 	const endPos = stream.GetPos() + stream.GetULong() + 4;
			//
			// 	// Skip start attributes byte
			// 	stream.Skip2(1);
			//
			// 	let id = null;
			// 	let uniqueID = null;
			//
			// 	// Read attributes
			// 	while (true) {
			// 		const attrType = stream.GetUChar();
			// 		if (attrType === AscVisio.g_nodeAttributeEnd)
			// 			break;
			//
			// 		switch (attrType) {
			// 			case 0:
			// 				id = stream.GetULong();
			// 				break;
			// 			case 1:
			// 				uniqueID = stream.GetString2();
			// 				break;
			// 		}
			// 	}
			//
			// 	// Store master relationship info
			// 	if (id !== null || uniqueID !== null) {
			// 		if (!this.masterRels) {
			// 			this.masterRels = [];
			// 		}
			// 		this.masterRels.push({ id: id, uniqueID: uniqueID });
			// 	}
			//
			// 	// Seek to end of record
			// 	stream.Seek2(endPos);
			// 	break;
			// }
			// case 3: {
			// 	// Handle Page relationships
			// 	const stream = pReader.stream;
			// 	const endPos = stream.GetPos() + stream.GetULong() + 4;
			//
			// 	// Skip start attributes byte
			// 	stream.Skip2(1);
			//
			// 	let id = null;
			//
			// 	// Read attributes
			// 	while (true) {
			// 		const attrType = stream.GetUChar();
			// 		if (attrType === AscVisio.g_nodeAttributeEnd)
			// 			break;
			//
			// 		switch (attrType) {
			// 			case 0:
			// 				id = stream.GetULong();
			// 				break;
			// 		}
			// 	}
			//
			// 	// Store page relationship info
			// 	if (id !== null) {
			// 		if (!this.pageRels) {
			// 			this.pageRels = [];
			// 		}
			// 		this.pageRels.push(id);
			// 	}
			//
			// 	// Seek to end of record
			// 	stream.Seek2(endPos);
			// 	break;
			// }
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Loads a Visio document from a binary stream
	 *
	 * @param  {BinaryVSDYLoader} pReader - The binary reader
	 */
	AscVisio.Shape_Type.prototype.fromPPTY = function(pReader) {
		AscVisio.SheetStorageAndStyles.prototype.fromPPTY.call(this, pReader);
		//todo make common with xml
		if (this.type === AscVisio.SHAPE_TYPES_FOREIGN) {
			let foreignDataObject = this.getForeignDataObject();
			if (foreignDataObject && foreignDataObject.mediaFilename) {
				// if image init image
				// imitate presentation parse workflow

				// CBlipFill stores data about blib fill (stretch, effects, rasterImageId) and calls CBlip.fromXml
				// CBlip has this.blipFill and this.link
				// CBlip.prototype.readAttrXml reads image rId and calls
				// fReadXmlRasterImageId which inits reader.context.imageMap, imageMap stores blipFills array.
				// CBlipFill is part of CUniFill which is part of CImageShape

				// pic tag read
				let oSp = new AscFormat.CImageShape();

				// nvPicPr tag read
				let prop = new AscFormat.UniNvPr();
				oSp.setNvPicPr(prop);
				oSp.setLocks(prop.getLocks());


				// blipFill tag read
				let uni_fill = new AscFormat.CUniFill();
				uni_fill.fill = new AscFormat.CBlipFill();
				if (uni_fill.fill) {
					let oImageShape = oSp;
					var sReadPath = foreignDataObject.mediaFilename;
					if (this.IsUseFullUrl && this.insertDocumentUrlsData && this.insertDocumentUrlsData.imageMap) {
						var sReadPathNew = this.insertDocumentUrlsData.imageMap[AscCommon.g_oDocumentUrls.mediaPrefix + sReadPath];
						if(sReadPathNew){
							sReadPath = sReadPathNew;
						}
					}
					if(this.IsUseFullUrl) {
						if(window["native"] && window["native"]["CopyTmpToMedia"]){
							if(!(window.documentInfo && window.documentInfo["iscoauthoring"])){
								var sMedia = window["native"]["CopyTmpToMedia"](sReadPath);
								if(typeof sMedia === "string" && sMedia.length > 0){
									sReadPath = sMedia;
								}
							}
						}
					}
					uni_fill.fill.setRasterImageId(sReadPath);

					// TEST version ---------------
					var _s = sReadPath;
					var indS = _s.lastIndexOf("emf");
					if (indS == -1)
						indS = _s.lastIndexOf("wmf");

					if (indS != -1 && (indS == (_s.length - 3)))
					{
						_s = _s.substring(0, indS);
						_s += "svg";
						sReadPath = _s;
						uni_fill.fill.setRasterImageId(_s);
					}
					// ----------------------------

					if (this.IsThemeLoader)
					{
						sReadPath = "theme" + (this.Api.ThemeLoader.CurrentLoadThemeIndex + 1) + "/media/" + sReadPath;
						uni_fill.fill.setRasterImageId(sReadPath);
					}

					if (pReader.ImageMapChecker != null)
					{
						let bAddToMap = true;
						if(oImageShape && oImageShape instanceof AscFormat.COleObject)
						{
							if(sReadPath.indexOf(".") === -1)
							{
								bAddToMap = false;
							}
						}
						if(bAddToMap)
						{
							if (!pReader.document.ImageMap) {
								pReader.document.ImageMap = {};
							}
							pReader.document.ImageMap[Object.keys(pReader.document.ImageMap).length] = sReadPath;
						}
					}

					if (this.IsUseFullUrl) {
						//this.RebuildImages.push(new CBuilderImages(uni_fill.fill, sReadPath, oImageShape, oSpPr, oLn, undefined, undefined, undefined, oParagraph, oBullet));
					}


					if (uni_fill.checkTransparent) {
						uni_fill.checkTransparent();
					}
				}
				oSp.setBlipFill(uni_fill.fill);

				this.cImageShape = oSp;
			}
		}
	}
	/**
	 * Read attributes from stream for Shape_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.Shape_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.id = stream.GetULong();
				break;
			case 1:
				this.type = stream.GetUChar();
				break;
			case 2:
				this.originalID = stream.GetULong();
				break;
			case 3:
				this.del = stream.GetBool();
				break;
			case 4:
				this.masterShape = stream.GetULong();
				break;
			case 5:
				this.uniqueID = stream.GetString2();
				break;
			case 6:
				this.nameU = stream.GetString2();
				break;
			case 7:
				this.name = stream.GetString2();
				break;
			case 8:
				this.isCustomName = stream.GetBool();
				break;
			case 9:
				this.isCustomNameU = stream.GetBool();
				break;
			case 10:
				this.master = stream.GetULong();
				break;
			case 11:
				this.lineStyle = stream.GetULong();
				break;
			case 12:
				this.fillStyle = stream.GetULong();
				break;
			case 13:
				this.textStyle = stream.GetULong();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read child elements from stream for Shape_Type
	 *
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.Shape_Type.prototype.readChild = function(elementType, pReader) {
		let t = this;
		switch (elementType) {
			case 0: {
				const cell = new AscVisio.Cell_Type();
				cell.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(cell);
				this.elements[key] = cell;
				break;
			}
			case 1: {
				const trigger = new AscVisio.Trigger_Type();
				trigger.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(trigger);
				this.elements[key] = trigger;
				break;
			}
			case 2: {
				const section = new AscVisio.Section_Type();
				section.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(section);
				this.elements[key] = section;
				break;
			}
			case 3: {
				const text = new AscVisio.Text_Type();
				text.fromPPTY(pReader);
				const key = AscVisio.createKeyFromSheetObject(text);
				this.elements[key] = text;
				break;
			}
			case 4: {
				const foreignData = new AscVisio.ForeignData_Type();
				foreignData.fromPPTY(pReader);
				let key = AscVisio.createKeyFromSheetObject(foreignData);
				this.elements[key] = foreignData;
				break;
			}
			case 5: {
				// Read Shapes
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readChild: function(elementType, pReader) {
						if (elementType === 0) {
							const shape = new AscVisio.Shape_Type();
							shape.fromPPTY(pReader);
							t.shapes.push(shape);
							return true;
						}
						return false;
					}
				}, pReader);
				break;
			}
			default:
				return false;
		}
		
		return true;
	};

	AscVisio.Text_Type.prototype.readAttribute = undefined;
	/**
	 * Read child elements from stream for Text_Type
	 *
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.Text_Type.prototype.readChild = function(elementType, pReader) {
		let t = this;
		switch (elementType) {
			case 0: {
				// Text_cp (Character properties)
				const textCp = new AscVisio.cp_Type();
				textCp.fromPPTY(pReader);
				this.elements.push(textCp);
				break;
			}
			case 1: {
				// Text_pp (Paragraph properties)
				const textPp = new AscVisio.pp_Type();
				textPp.fromPPTY(pReader);
				this.elements.push(textPp);
				break;
			}
			case 2: {
				// Text_tp (Tab properties)
				const textTp = new AscVisio.tp_Type();
				textTp.fromPPTY(pReader);
				this.elements.push(textTp);
				break;
			}
			case 3: {
				// Text_fld (Field)
				const textFld = new AscVisio.fld_Type();
				textFld.fromPPTY(pReader);
				this.elements.push(textFld);
				break;
			}
			case 4: {
				// Text_text (Text content)
				// Read Shapes
				AscFormat.CBaseFormatNoIdObject.prototype.fromPPTY.call({
					readChildren: AscFormat.CBaseFormatNoIdObject.prototype.readChildren,
					readAttributes: AscFormat.CBaseFormatNoIdObject.prototype.readAttributes,
					readAttribute: function(attrType, pReader) {
						if (attrType === 0) {
							t.elements.push(pReader.stream.GetString2());
							return true;
						}
						return false;
					}
				}, pReader);
				break;
			}
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read attributes from stream for Text_fld_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.fld_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.ix = stream.GetULong();
				break;
			case 1:
				this.value = stream.GetString2();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read attributes from stream for Text_tp_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.tp_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.ix = stream.GetULong();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read attributes from stream for Text_pp_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.pp_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.ix = stream.GetULong();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read attributes from stream for Text_cp_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.cp_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.ix = stream.GetULong();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read attributes from stream for CDataConnections
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.CDataConnections.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.nextID = pReader.stream.GetULong();
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read child elements from stream for CDataConnections
	 * 
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.CDataConnections.prototype.readChild = function(elementType, pReader)
	{
		let handled = true;
		
		switch (elementType)
		{
			case 0:
				let dataConnection = new AscVisio.DataConnection_Type();
				dataConnection.fromPPTY(pReader);
				this.dataConnection.push(dataConnection);
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for CDataRecordSets
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.CDataRecordSets.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.nextID = pReader.stream.GetULong();
				break;
			case 1:
				this.activeRecordsetID = pReader.stream.GetULong();
				break;
			case 2:
				this.dataWindowOrder = pReader.stream.GetString2();
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read child elements from stream for CDataRecordSets
	 * 
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.CDataRecordSets.prototype.readChild = function(elementType, pReader)
	{
		let handled = true;
		
		switch (elementType)
		{
			case 0:
				let dataRecordSet = new AscVisio.DataRecordSet_Type();
				dataRecordSet.fromPPTY(pReader);
				this.dataRecordSet.push(dataRecordSet);
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for CommentEntry_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.CommentEntry_Type.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.authorID = pReader.stream.GetULong();
				break;
			case 1:
				this.pageID = pReader.stream.GetULong();
				break;
			case 2:
				this.shapeID = pReader.stream.GetULong();
				break;
			case 3:
				this.date = pReader.stream.GetString2();
				break;
			case 4:
				this.editDate = pReader.stream.GetString2();
				break;
			case 5:
				this.commentID = pReader.stream.GetULong();
				break;
			case 6:
				this.autoCommentType = pReader.stream.GetULong();
				break;
			case 7:
				this.value = pReader.stream.GetString2(); // Called 'content' in C++
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for AuthorEntry_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.AuthorEntry_Type.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.id = pReader.stream.GetULong();
				break;
			case 1:
				this.name = pReader.stream.GetString2();
				break;
			case 2:
				this.initials = pReader.stream.GetString2();
				break;
			case 3:
				this.resolutionID = pReader.stream.GetString2();
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for RuleTest_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.RuleTest_Type.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.value = pReader.stream.GetString2(); // Formula in C++
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};
	AscVisio.RuleFilter_Type.prototype.readAttribute = AscVisio.RuleTest_Type.prototype.readAttribute

	/**
	 * Read attributes from stream for RowKeyValue_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.RowKeyValue_Type.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.rowID = pReader.stream.GetULong();
				break;
			case 1:
				this.value = pReader.stream.GetString2();
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for DataColumn_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.DataColumn_Type.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.columnNameID = pReader.stream.GetString2();
				break;
			case 1:
				this.name = pReader.stream.GetString2();
				break;
			case 2:
				this.label = pReader.stream.GetString2();
				break;
			case 3:
				this.origLabel = pReader.stream.GetString2();
				break;
			case 4:
				//todo int or string
				pReader.stream.GetString2();
				//this.langID ;
				break;
			case 5:
				this.calendar = pReader.stream.GetULong();
				break;
			case 6:
				this.dataType = pReader.stream.GetULong();
				break;
			case 7:
				this.unitType = pReader.stream.GetString2();
				break;
			case 8:
				this.currency = pReader.stream.GetULong();
				break;
			case 9:
				this.degree = pReader.stream.GetULong();
				break;
			case 10:
				this.displayWidth = pReader.stream.GetULong();
				break;
			case 11:
				this.displayOrder = pReader.stream.GetULong();
				break;
			case 12:
				this.mapped = pReader.stream.GetBool();
				break;
			case 13:
				this.hyperlink = pReader.stream.GetBool();
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for RuleInfo_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.RuleInfo_Type.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.ruleID = pReader.stream.GetULong();
				break;
			case 1:
				this.ruleSetID = pReader.stream.GetULong();
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for IssueTarget_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.IssueTarget_Type.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.pageID = pReader.stream.GetULong();
				break;
			case 1:
				this.shapeID = pReader.stream.GetULong();
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for Rule_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.Rule_Type.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.id = pReader.stream.GetULong();
				break;
			case 1:
				this.category = pReader.stream.GetString2();
				break;
			case 2:
				this.nameU = pReader.stream.GetString2();
				break;
			case 3:
				this.ignored = pReader.stream.GetBool();
				break;
			case 4:
				this.description = pReader.stream.GetString2();
				break;
			case 5:
				this.ruleTarget = pReader.stream.GetLong();
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read child elements from stream for Rule_Type
	 * 
	 * @param {number} recType - The type of record
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if child was handled, false otherwise
	 */
	AscVisio.Rule_Type.prototype.readChild = function(recType, pReader)
	{
		let handled = true;
		
		switch (recType)
		{
			case 0:
				this.ruleFilter = new AscVisio.RuleFilter_Type();
				this.ruleFilter.fromPPTY(pReader);
				break;
			case 1:
				this.ruleTest = new AscVisio.RuleTest_Type();
				this.ruleTest.fromPPTY(pReader);
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for RuleSetFlags_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.RuleSetFlags_Type.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.hidden = pReader.stream.GetBool();
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for RowMap_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.RowMap_Type.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.rowID = pReader.stream.GetULong();
				break;
			case 1:
				this.pageID = pReader.stream.GetULong();
				break;
			case 2:
				this.shapeID = pReader.stream.GetULong();
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for PrimaryKey_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.PrimaryKey_Type.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.columnNameID = pReader.stream.GetString2();
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read child elements from stream for PrimaryKey_Type
	 * 
	 * @param {number} recType - The type of record
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if child was handled, false otherwise
	 */
	AscVisio.PrimaryKey_Type.prototype.readChild = function(recType, pReader)
	{
		let handled = true;
		
		switch (recType)
		{
			case 0:
				const rowKeyValue = new AscVisio.RowKeyValue_Type();
				rowKeyValue.fromPPTY(pReader);
				this.rowKeyValue.push(rowKeyValue);
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for DataColumns_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.DataColumns_Type.prototype.readAttribute = function(attrType, pReader)
	{
		let handled = true;
		
		switch (attrType)
		{
			case 0:
				this.sortColumn = pReader.stream.GetString2();
				break;
			case 1:
				this.sortAsc = pReader.stream.GetBool();
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read child elements from stream for DataColumns_Type
	 * 
	 * @param {number} recType - The type of record
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if child was handled, false otherwise
	 */
	AscVisio.DataColumns_Type.prototype.readChild = function(recType, pReader)
	{
		let handled = true;
		
		switch (recType)
		{
			case 0:
				const dataColumn = new AscVisio.DataColumn_Type();
				dataColumn.fromPPTY(pReader);
				this.dataColumn.push(dataColumn);
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};


	/**
	 * Read attributes from stream for ForeignData_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.ForeignData_Type.prototype.readAttribute = function(attrType, pReader) {
		let res = true;
		
		switch (attrType) {
			case 0:
				this.foreignType = pReader.stream.GetUChar();
				break;
			case 1:
				this.objectType = pReader.stream.GetULong();
				break;
			case 2:
				this.showAsIcon = pReader.stream.GetBool();
				break;
			case 3:
				this.objectWidth = pReader.stream.GetDouble();
				break;
			case 4:
				this.objectHeight = pReader.stream.GetDouble();
				break;
			case 5:
				this.extentX = pReader.stream.GetDouble();
				break;
			case 6:
				this.extentY = pReader.stream.GetDouble();
				break;
			case 7:
				this.compressionType = pReader.stream.GetUChar();
				break;
			case 8:
				this.compressionLevel = pReader.stream.GetDouble();
				break;
			default:
				res = false;
				break;
		}
		
		return res;
	};

	/**
	 * Read child elements from stream for ForeignData_Type
	 *
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.ForeignData_Type.prototype.readChild = function(elementType, pReader) {
		let res = true;
		
		switch (elementType) {
			case 1:
				pReader.stream.Skip2(4);//len
				// media filename
				this.mediaFilename = pReader.stream.GetString2();
				break;
			case 2:
				pReader.stream.Skip2(4);//len
				// ole filename
				this.oleFilename = pReader.stream.GetString2();
				break;
			default:
				res = false;
				break;
		}
		
		return res;
	};

	/**
	 * Read attributes from stream for DocumentSheet_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.DocumentSheet_Type.prototype.readAttribute = function(attrType, pReader) {
		let res = true;
		
		switch (attrType) {
			case 0:
				this.uniqueID = pReader.stream.GetString2();
				break;
			case 1:
				this.nameU = pReader.stream.GetString2();
				break;
			case 2:
				this.name = pReader.stream.GetString2();
				break;
			case 3:
				this.isCustomName = pReader.stream.GetBool();
				break;
			case 4:
				this.isCustomNameU = pReader.stream.GetBool();
				break;
			case 5:
				this.lineStyle = pReader.stream.GetULong();
				break;
			case 6:
				this.fillStyle = pReader.stream.GetULong();
				break;
			case 7:
				this.textStyle = pReader.stream.GetULong();
				break;
			default:
				res = false;
				break;
		}
		
		return res;
	};

	/**
	 * Read child elements from stream for DocumentSheet_Type
	 *
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.DocumentSheet_Type.prototype.readChild = function(elementType, pReader) {
		let res = true;
		let element, key;
		
		switch (elementType) {
			case 0:
				element = new AscVisio.Cell_Type();
				element.fromPPTY(pReader);
				key = AscVisio.createKeyFromSheetObject(element);
				this.elements[key] = element;
				break;
			case 1:
				element = new AscVisio.Trigger_Type();
				element.fromPPTY(pReader);
				key = AscVisio.createKeyFromSheetObject(element);
				this.elements[key] = element;
				break;
			case 2:
				element = new AscVisio.Section_Type();
				element.fromPPTY(pReader);
				key = AscVisio.createKeyFromSheetObject(element);
				this.elements[key] = element;
				break;
			default:
				res = false;
				break;
		}
		
		return res;
	};

	/**
	 * Read attributes from stream for RefBy_Type
	 *
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.RefBy_Type.prototype.readAttribute = function(attrType, pReader) {
		let res = true;
		
		switch (attrType) {
			case 0:
				this.id = pReader.stream.GetULong();
				break;
			case 1:
				this.t = pReader.stream.GetString2();
				break;
			default:
				res = false;
				break;
		}
		
		return res;
	};

	/**
	 * Read from binary stream for Icon_Type
	 *
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 */
	AscVisio.Icon_Type.prototype.fromPPTY = function(pReader) {
		const _end_rec = pReader.stream.cur + pReader.stream.GetULong() + 4;
		
		this.value = pReader.stream.GetString2();
		
		pReader.stream.Seek2(_end_rec);
	};

	/**
	 * Read attributes from stream for Issue_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.Issue_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.id = stream.GetULong();
				break;
			case 1:
				this.ignored = stream.GetBool();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read child elements from stream for Issue_Type
	 * 
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.Issue_Type.prototype.readChild = function(elementType, pReader) {
		let handled = true;
		
		switch (elementType) {
			case 0:
				this.issueTarget = new AscVisio.IssueTarget_Type();
				this.issueTarget.fromPPTY(pReader);
				break;
			case 1:
				this.ruleInfo = new AscVisio.RuleInfo_Type();
				this.ruleInfo.fromPPTY(pReader);
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for RuleSet_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.RuleSet_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.id = stream.GetULong();
				break;
			case 1:
				this.nameU = stream.GetString2();
				break;
			case 2:
				this.name = stream.GetString2();
				break;
			case 3:
				this.description = stream.GetString2();
				break;
			case 4:
				this.enabled = stream.GetBool();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read child elements from stream for RuleSet_Type
	 * 
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.RuleSet_Type.prototype.readChild = function(elementType, pReader) {
		let handled = true;
		
		switch (elementType) {
			case 0:
				this.ruleSetFlags = new AscVisio.RuleSetFlags_Type();
				this.ruleSetFlags.fromPPTY(pReader);
				break;
			case 1:
				let rule = new AscVisio.Rule_Type();
				rule.fromPPTY(pReader);
				this.rule.push(rule);
				break;
			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for ValidationProperties_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.ValidationProperties_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.showIgnored = stream.GetBool();
				break;
			case 1:
				this.lastValidated = stream.GetString2();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read attributes from stream for DataRecordSet_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.DataRecordSet_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.id = stream.GetULong();
				break;
			case 1:
				this.connectionID = stream.GetULong();
				break;
			case 2:
				this.command = stream.GetString2();
				break;
			case 3:
				this.options = stream.GetULong();
				break;
			case 4:
				this.timeRefreshed = stream.GetString2();
				break;
			case 5:
				this.nextRowID = stream.GetULong();
				break;
			case 6:
				this.name = stream.GetString2();
				break;
			case 7:
				this.rowOrder = stream.GetString2();
				break;
			case 8:
				this.refreshOverwriteAll = stream.GetBool();
				break;
			case 9:
				this.refreshNoReconciliationUI = stream.GetBool();
				break;
			case 10:
				this.refreshInterval = stream.GetULong();
				break;
			case 11:
				this.replaceLinks = stream.GetBool();
				break;
			case 12:
				this.checksum = stream.GetULong();
				break;
			default:
				return false;
		}
		
		return true;
	};

	/**
	 * Read child elements from stream for DataRecordSet_Type
	 * 
	 * @param {number} elementType - The type of child element
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if element was handled, false otherwise
	 */
	AscVisio.DataRecordSet_Type.prototype.readChild = function(elementType, pReader) {
		let handled = true;
		
		switch (elementType) {
			//todo
			// case 0:
			// 	this.aDOData = new AscVisio.ADOData_Type();
			// 	this.aDOData.fromPPTY(pReader);
			// 	break;
			case 0:
				this.dataColumns = new AscVisio.DataColumns_Type();
				this.dataColumns.fromPPTY(pReader);
				break;
			case 1:
				let primaryKey = new AscVisio.PrimaryKey_Type();
				primaryKey.fromPPTY(pReader);
				this.primaryKey.push(primaryKey);
				break;
			case 2:
				let rowMap = new AscVisio.RowMap_Type();
				rowMap.fromPPTY(pReader);
				this.rowMap.push(rowMap);
				break;
			case 5:
				pReader.stream.Skip2(4);//len
				pReader.stream.Skip2(1);//type
				const len = pReader.stream.GetULong();
				this.content = pReader.stream.GetBufferUint8(len);
				break;
			//todo
			// case 5:
			// 	let refreshConflict = new AscVisio.RefreshConflict_Type();
			// 	refreshConflict.fromPPTY(pReader);
			// 	this.refreshConflict.push(refreshConflict);
			// 	break;
			// case 6:
			// 	let autoLinkComparison = new AscVisio.AutoLinkComparison_Type();
			// 	autoLinkComparison.fromPPTY(pReader);
			// 	this.autoLinkComparison.push(autoLinkComparison);
			// 	break;


			default:
				handled = false;
				break;
		}
		
		return handled;
	};

	/**
	 * Read attributes from stream for DataConnection_Type
	 * 
	 * @param {number} attrType - The type of attribute
	 * @param {BinaryVSDYLoader} pReader - The binary reader
	 * @returns {boolean} - True if attribute was handled, false otherwise
	 */
	AscVisio.DataConnection_Type.prototype.readAttribute = function(attrType, pReader) {
		const stream = pReader.stream;
		switch (attrType) {
			case 0:
				this.id = stream.GetULong();
				break;
			case 1:
				this.fileName = stream.GetString2();
				break;
			case 2:
				this.connectionString = stream.GetString2();
				break;
			case 3:
				this.command = stream.GetString2();
				break;
			case 4:
				this.friendlyName = stream.GetString2();
				break;
			case 5:
				this.timeout = stream.GetULong();
				break;
			case 6:
				this.alwaysUseConnectionFile = stream.GetBool();
				break;
			default:
				return false;
		}
		
		return true;
	};

	window['AscVisio']  = window['AscVisio'] || {};

	window['AscVisio'].BinaryVSDYLoader = BinaryVSDYLoader;

})(window, window.document);

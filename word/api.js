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

"use strict";

(function(window, document)
{

	// Import
	var g_fontApplication = null;

	var c_oAscAdvancedOptionsAction      = AscCommon.c_oAscAdvancedOptionsAction;
	var DownloadType                     = AscCommon.DownloadType;
	var c_oAscFormatPainterState         = AscCommon.c_oAscFormatPainterState;
	var locktype_None                    = AscCommon.c_oAscLockTypes.kLockTypeNone;
	var locktype_Mine                    = AscCommon.c_oAscLockTypes.kLockTypeMine;
	var locktype_Other                   = AscCommon.c_oAscLockTypes.kLockTypeOther;
	var locktype_Other2                  = AscCommon.c_oAscLockTypes.kLockTypeOther2;
	var locktype_Other3                  = AscCommon.c_oAscLockTypes.kLockTypeOther3;
	var changestype_None                 = AscCommon.changestype_None;
	var changestype_Paragraph_Content    = AscCommon.changestype_Paragraph_Content;
	var changestype_Paragraph_Properties = AscCommon.changestype_Paragraph_Properties;
	var changestype_Table_Properties     = AscCommon.changestype_Table_Properties;
	var changestype_Table_RemoveCells    = AscCommon.changestype_Table_RemoveCells;
	var changestype_HdrFtr               = AscCommon.changestype_HdrFtr;
	var asc_CTextFontFamily              = AscCommon.asc_CTextFontFamily;
	var asc_CSelectedObject              = AscCommon.asc_CSelectedObject;
	var g_oDocumentUrls                  = AscCommon.g_oDocumentUrls;
	var sendCommand                      = AscCommon.sendCommand;
	var g_oIdCounter                     = AscCommon.g_oIdCounter;
	var g_oTableId                       = AscCommon.g_oTableId;
	var PasteElementsId                  = null;
	var global_mouseEvent                = null;
	var History                          = null;

	var c_oAscError                 = Asc.c_oAscError;
	var c_oAscFileType              = Asc.c_oAscFileType;
	var c_oAscAsyncAction           = Asc.c_oAscAsyncAction;
	var c_oAscAdvancedOptionsID     = Asc.c_oAscAdvancedOptionsID;
	var c_oAscFontRenderingModeType = Asc.c_oAscFontRenderingModeType;
	var c_oAscAsyncActionType       = Asc.c_oAscAsyncActionType;
	var c_oAscTypeSelectElement     = Asc.c_oAscTypeSelectElement;
	var c_oAscFill                  = Asc.c_oAscFill;
	var asc_CImgProperty            = Asc.asc_CImgProperty;
	var asc_CShapeFill              = Asc.asc_CShapeFill;
	var asc_CFillBlip               = Asc.asc_CFillBlip;
	var c_oAscSdtLockType           = Asc.c_oAscSdtLockType;

	function CAscSection()
	{
		this.PageWidth  = 0;
		this.PageHeight = 0;

		this.MarginLeft   = 0;
		this.MarginRight  = 0;
		this.MarginTop    = 0;
		this.MarginBottom = 0;
	}

	CAscSection.prototype.get_PageWidth    = function()
	{
		return this.PageWidth;
	};
	CAscSection.prototype.get_PageHeight   = function()
	{
		return this.PageHeight;
	};
	CAscSection.prototype.get_MarginLeft   = function()
	{
		return this.MarginLeft;
	};
	CAscSection.prototype.get_MarginRight  = function()
	{
		return this.MarginRight;
	};
	CAscSection.prototype.get_MarginTop    = function()
	{
		return this.MarginTop;
	};
	CAscSection.prototype.get_MarginBottom = function()
	{
		return this.MarginBottom;
	};

	function CHeaderProp(obj)
	{
		/*{
		 Type : hdrftr_Footer (hdrftr_Header),
		 Position : 12.5,
		 DifferentFirst : true/false,
		 DifferentEvenOdd : true/false,
		 }*/
		if (obj)
		{
			this.Type             = (undefined != obj.Type) ? obj.Type : null;
			this.Position         = (undefined != obj.Position) ? obj.Position : null;
			this.DifferentFirst   = (undefined != obj.DifferentFirst) ? obj.DifferentFirst : null;
			this.DifferentEvenOdd = (undefined != obj.DifferentEvenOdd) ? obj.DifferentEvenOdd : null;
			this.LinkToPrevious   = (undefined != obj.LinkToPrevious) ? obj.LinkToPrevious : null;
			this.Locked           = (undefined != obj.Locked) ? obj.Locked : false;
			this.StartPageNumber  = (undefined != obj.StartPageNumber) ? obj.StartPageNumber : -1;
			this.NumFormat        = (undefined !== obj.NumFormat) ? obj.NumFormat : -1;
		}
		else
		{
			this.Type             = AscCommon.hdrftr_Footer;
			this.Position         = 12.5;
			this.DifferentFirst   = false;
			this.DifferentEvenOdd = false;
			this.LinkToPrevious   = null;
			this.Locked           = false;
			this.StartPageNumber  = -1;
			this.NumFormat        = -1;
		}
	}

	CHeaderProp.prototype.get_Type             = function()
	{
		return this.Type;
	};
	CHeaderProp.prototype.put_Type             = function(v)
	{
		this.Type = v;
	};
	CHeaderProp.prototype.get_Position         = function()
	{
		return this.Position;
	};
	CHeaderProp.prototype.put_Position         = function(v)
	{
		this.Position = v;
	};
	CHeaderProp.prototype.get_DifferentFirst   = function()
	{
		return this.DifferentFirst;
	};
	CHeaderProp.prototype.put_DifferentFirst   = function(v)
	{
		this.DifferentFirst = v;
	};
	CHeaderProp.prototype.get_DifferentEvenOdd = function()
	{
		return this.DifferentEvenOdd;
	};
	CHeaderProp.prototype.put_DifferentEvenOdd = function(v)
	{
		this.DifferentEvenOdd = v;
	};
	CHeaderProp.prototype.get_LinkToPrevious   = function()
	{
		return this.LinkToPrevious;
	};
	CHeaderProp.prototype.get_Locked           = function()
	{
		return this.Locked;
	};
	CHeaderProp.prototype.get_StartPageNumber = function()
	{
		return this.StartPageNumber;
	};
	CHeaderProp.prototype.put_StartPageNumber = function(nStartPage)
	{
		this.StartPageNumber = nStartPage;
	};
	CHeaderProp.prototype.get_NumFormat = function()
	{
		return this.NumFormat;
	};
	CHeaderProp.prototype.put_NumFormat = function(format)
	{
		this.NumFormat = format;
	};

	var DocumentPageSize = new function()
	{
		this.oSizes    = [{name : "US Letter", w_mm : 215.9, h_mm : 279.4, w_tw : 12240, h_tw : 15840},
			{name : "US Legal", w_mm : 215.9, h_mm : 355.6, w_tw : 12240, h_tw : 20160},
			{name : "A4", w_mm : 210, h_mm : 297, w_tw : 11907, h_tw : 16839},
			{name : "A5", w_mm : 148.1, h_mm : 209.9, w_tw : 8391, h_tw : 11907},
			{name : "B5", w_mm : 176, h_mm : 250.1, w_tw : 9979, h_tw : 14175},
			{name : "Envelope #10", w_mm : 104.8, h_mm : 241.3, w_tw : 5940, h_tw : 13680},
			{name : "Envelope DL", w_mm : 110.1, h_mm : 220.1, w_tw : 6237, h_tw : 12474},
			{name : "Tabloid", w_mm : 279.4, h_mm : 431.7, w_tw : 15842, h_tw : 24477},
			{name : "A3", w_mm : 297, h_mm : 420.1, w_tw : 16840, h_tw : 23820},
			{name : "Tabloid Oversize", w_mm : 304.8, h_mm : 457.1, w_tw : 17282, h_tw : 25918},
			{name : "ROC 16K", w_mm : 196.8, h_mm : 273, w_tw : 11164, h_tw : 15485},
			{name : "Envelope Coukei 3", w_mm : 119.9, h_mm : 234.9, w_tw : 6798, h_tw : 13319},
			{name : "Super B/A3", w_mm : 330.2, h_mm : 482.5, w_tw : 18722, h_tw : 27358}
		];
		this.sizeEpsMM = 0.5;
		this.getSize   = function(widthMm, heightMm)
		{
			for (var index in this.oSizes)
			{
				var item = this.oSizes[index];
				if (Math.abs(widthMm - item.w_mm) < this.sizeEpsMM && Math.abs(heightMm - item.h_mm) < this.sizeEpsMM)
					return item;
			}
			return {w_mm : widthMm, h_mm : heightMm};
		};
	};

	function CMailMergeSendData(obj)
	{
		if (obj)
		{
			if (typeof obj.from != 'undefined')
			{
				this["from"] = obj.from;
			}
			if (typeof obj.to != 'undefined')
			{
				this["to"] = obj.to;
			}
			if (typeof obj.subject != 'undefined')
			{
				this["subject"] = obj.subject;
			}
			if (typeof obj.mailFormat != 'undefined')
			{
				this["mailFormat"] = obj.mailFormat;
			}
			if (typeof obj.fileName != 'undefined')
			{
				this["fileName"] = obj.fileName;
			}
			if (typeof obj.message != 'undefined')
			{
				this["message"] = obj.message;
			}
			if (typeof obj.recordFrom != 'undefined')
			{
				this["recordFrom"] = obj.recordFrom;
			}
			if (typeof obj.recordTo != 'undefined')
			{
				this["recordTo"] = obj.recordTo;
			}
			if (typeof obj.isJson != 'undefined')
			{
				this["isJson"] = obj.isJson;
			}
		}
		else
		{
			this["from"]        = null;
			this["to"]          = null;
			this["subject"]     = null;
			this["mailFormat"]  = null;
			this["fileName"]    = null;
			this["message"]     = null;
			this["recordFrom"]  = null;
			this["recordTo"]    = null;
			this["recordCount"] = null;
			this["userId"]      = null;
			this["isJson"]      = null;
		}
	}

	CMailMergeSendData.prototype.get_From        = function()
	{
		return this["from"]
	};
	CMailMergeSendData.prototype.put_From        = function(v)
	{
		this["from"] = v;
	};
	CMailMergeSendData.prototype.get_To          = function()
	{
		return this["to"]
	};
	CMailMergeSendData.prototype.put_To          = function(v)
	{
		this["to"] = v;
	};
	CMailMergeSendData.prototype.get_Subject     = function()
	{
		return this["subject"]
	};
	CMailMergeSendData.prototype.put_Subject     = function(v)
	{
		this["subject"] = v;
	};
	CMailMergeSendData.prototype.get_MailFormat  = function()
	{
		return this["mailFormat"]
	};
	CMailMergeSendData.prototype.put_MailFormat  = function(v)
	{
		this["mailFormat"] = v;
	};
	CMailMergeSendData.prototype.get_FileName    = function()
	{
		return this["fileName"]
	};
	CMailMergeSendData.prototype.put_FileName    = function(v)
	{
		this["fileName"] = v;
	};
	CMailMergeSendData.prototype.get_Message     = function()
	{
		return this["message"]
	};
	CMailMergeSendData.prototype.put_Message     = function(v)
	{
		this["message"] = v;
	};
	CMailMergeSendData.prototype.get_RecordFrom  = function()
	{
		return this["recordFrom"]
	};
	CMailMergeSendData.prototype.put_RecordFrom  = function(v)
	{
		this["recordFrom"] = v;
	};
	CMailMergeSendData.prototype.get_RecordTo    = function()
	{
		return this["recordTo"]
	};
	CMailMergeSendData.prototype.put_RecordTo    = function(v)
	{
		this["recordTo"] = v;
	};
	CMailMergeSendData.prototype.get_RecordCount = function()
	{
		return this["recordCount"]
	};
	CMailMergeSendData.prototype.put_RecordCount = function(v)
	{
		this["recordCount"] = v;
	};
	CMailMergeSendData.prototype.get_UserId      = function()
	{
		return this["userId"]
	};
	CMailMergeSendData.prototype.put_UserId      = function(v)
	{
		this["userId"] = v;
	};
	CMailMergeSendData.prototype.get_IsJson      = function()
	{
		return this["isJson"]
	};
	CMailMergeSendData.prototype.put_IsJson      = function(v)
	{
		this["isJson"] = v;
	};

	function CAscFootnotePr(obj)
	{
		this.NumRestart = undefined;
		this.NumFormat  = undefined;
		this.NumStart   = undefined;
		this.Pos        = undefined;

		if (obj)
		{
			this.NumRestart = obj.NumRestart;
			this.NumFormat  = obj.NumFormat;
			this.NumStart   = obj.NumStart;
			this.Pos        = obj.Pos;
		}
	}
	CAscFootnotePr.prototype.get_Pos = function()
	{
		return this.Pos;
	};
	CAscFootnotePr.prototype.put_Pos = function(v)
	{
		this.Pos = v;
	};
	CAscFootnotePr.prototype.get_NumStart = function()
	{
		return this.NumStart;
	};
	CAscFootnotePr.prototype.put_NumStart = function(v)
	{
		this.NumStart = v;
	};
	CAscFootnotePr.prototype.get_NumFormat = function()
	{
		return this.NumFormat;
	};
	CAscFootnotePr.prototype.put_NumFormat = function(v)
	{
		this.NumFormat = v;
	};
	CAscFootnotePr.prototype.get_NumRestart = function()
	{
		return this.NumRestart;
	};
	CAscFootnotePr.prototype.put_NumRestart = function(v)
	{
		this.NumRestart = v;
	};

	function CContentControlPluginWorker(_api, _docs)
	{
		this.api = _api;
		this.documents = _docs;
		this.returnDocuments = [];
		this.current = -1;

		this.start = function()
		{
			// save worker in api
			this.api.__content_control_worker = this;
			this.api.incrementCounterLongAction();

			if (window.g_asc_plugins)
				window.g_asc_plugins.setPluginMethodReturnAsync();

			this.run();
		};
		this.end = function()
		{
			if (window.g_asc_plugins)
				window.g_asc_plugins.onPluginMethodReturn(this.returnDocuments);

			delete this.api.__content_control_worker;
			this.api.decrementCounterLongAction();

			this.api.WordControl.m_oLogicDocument.UpdateSelection();
			this.api.WordControl.m_oLogicDocument.FinalizeAction();
		};

		/**
		 * функция-коллбэк с информацией о локе и о типе лока
		 * @param isLocked заблокировано или нет
		 * @param type 0 - вставка/исполнение скрипта, 1 - изменение настроек
		 * @returns {boolean}
		*/
		this.checkCurrentSelectionLockCallback = function(isLocked, type)
		{
			function private_nextFile() {
				// избавляемся от рекурсии
				setTimeout(function() { window.g_asc_plugins.api.__content_control_worker.run(); }, 1);
			}

			if (isLocked && type === 0)
			{
				// залочена вставка - проверяем настройки
				var LogicDocument = this.api.WordControl.m_oLogicDocument;
				LogicDocument.IsSelectionLocked(AscCommon.changestype_ContentControl_Properties, undefined, false, false, function(isLocked) {
					window.g_asc_plugins.api.__content_control_worker.checkCurrentSelectionLockCallback(isLocked, 1);
				});
				return;
			}
			if (isLocked)
			{
				// залочена и вставка и настройки - т.е. просто пропускаем этот документ
				return private_nextFile();
			}

			var _obj = null;
			var LogicDocument = this.api.WordControl.m_oLogicDocument;
			var _current = this.documents[this.current];

			if (0 === type)
			{
				var _content_control_pr;
				var _blockStd;
				var _isReplaced = false;

				if (_current["Url"] !== undefined || _current["Script"] !== undefined)
				{
					_blockStd = null;
					if (undefined !== _current["Props"]["InternalId"])
					{
						_blockStd = LogicDocument.GetContentControl(_current["Props"]["InternalId"]);
						if (_blockStd && !_blockStd.IsPlaceHolder())
							_blockStd.ReplaceContentWithPlaceHolder();
						
						_isReplaced = true;
					}

					_content_control_pr = new AscCommon.CContentControlPr();
					AscCommon.readContentControlCommonPr(_content_control_pr, _current["Props"]);

					_content_control_pr.Lock       = c_oAscSdtLockType.Unlocked;
					_content_control_pr.InternalId = _current["Props"]["InternalId"];

					// Page break
					if (undefined !== _current["Props"]["SectionBreak"])
					{
						var oCurPara = LogicDocument.GetCurrentParagraph();
						LogicDocument.Content[oCurPara.Index].Document_SetThisElementCurrent();
						LogicDocument.Add_SectionBreak(_current["Props"]["SectionBreak"]);
						LogicDocument.Add_SectionBreak(_current["Props"]["SectionBreak"]);
						LogicDocument.Content[oCurPara.Index + 1].Document_SetThisElementCurrent();
					}

					// Page Size
					if (undefined !== _current["Props"]["PageSizeW"] || undefined !== _current["Props"]["PageSizeH"])
					{
						var Width  = _current["Props"]["PageSizeW"];
						if (Width === undefined)
						{
							Width = LogicDocument.Get_DocumentPageSize().W;
						}
						if (Width < 58)
						{
							Width = 58;
						}

						var Height = _current["Props"]["PageSizeH"];
						if (Height === undefined)
						{
							Height = LogicDocument.Get_DocumentPageSize().H;
						}
						if (Height < 43)
						{
							Height = 43;
						}

						LogicDocument.Set_DocumentPageSize(Width, Height);
					}

					// Orient
					if (undefined !== _current["Props"]["Orient"])
					{
						if (_current["Props"]["Orient"] === 0 || _current["Props"]["Orient"] === 1)
						{
							LogicDocument.Set_DocumentOrientation(_current["Props"]["Orient"]);
						}
					}

					// Margins
					if (undefined !== _current["Props"]["MarginT"] || undefined !== _current["Props"]["MarginL"] || undefined !== _current["Props"]["MarginR"] || undefined !== _current["Props"]["MarginB"])
					{
						var oMargins = {
							Left: _current["Props"]["MarginL"],
							Top: _current["Props"]["MarginT"],
							Right: _current["Props"]["MarginR"],
							Bottom: _current["Props"]["MarginB"]
						};
						LogicDocument.SetDocumentMargin(oMargins);
					}
					
					if (null === _blockStd)
					{
						let curPara = LogicDocument.GetCurrentParagraph();
						if (curPara)
						{
							let paraParent = curPara.GetParent();
							if (paraParent && paraParent.IsBlockLevelSdtContent())
							{
								let parentBlockSdt = paraParent.GetParent();
								if (parentBlockSdt.IsPlaceHolder())
									parentBlockSdt.ReplacePlaceHolderWithContent();
								else
									paraParent.AddNewParagraph(true);
							}
							else if (!curPara.IsCursorAtBegin())
							{
								LogicDocument.AddNewParagraph(false, true);
							}
						}
						
						
						_blockStd = LogicDocument.AddContentControl(c_oAscSdtLevelType.Block);
					}

					_blockStd.SetContentControlPr(_content_control_pr);
					_blockStd.SelectContentControl();

					_obj = _blockStd.GetContentControlPr();
					this.returnDocuments.push({"Tag" : _obj.Tag, "Id" : _obj.Id, "Lock" : _obj.Lock, "InternalId" : _obj.InternalId, "Alias" : _obj.Alias, "Appearance" : _obj.Appearance });
				}

				if (_current["Url"] !== undefined)
				{
					// insert/replace document
					this.api.insertDocumentUrlsData = {imageMap: null, documents: [{url : _current["Url"], format: _current["Format"], token: _current["Token"]}], convertCallback: function(_api, url) {
						_api.insertDocumentUrlsData.imageMap = url;
						AscCommon.loadFileContent(url['output.bin'], function(httpRequest) {
							var stream;
							if (null === httpRequest || !(stream = AscCommon.initStreamFromResponse(httpRequest)))
							{
								_api.endInsertDocumentUrls();
								_api.sendEvent("asc_onError", c_oAscError.ID.DirectUrl,
									c_oAscError.Level.NoCritical);
								return;
							}
							_api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Internal, stream, undefined, undefined, true, function() {
								_api.WordControl.m_oLogicDocument.MoveCursorRight(false, false, true);
								_api.WordControl.m_oLogicDocument.Recalculate();

								if (_api.insertDocumentUrlsData.documents.length > 0) {
									var options = new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.CANVAS_WORD);
									options.isNaturalDownload = true;
									_api.downloadAs(Asc.c_oAscAsyncAction.DownloadAs, options);
								} else {
									_api.endInsertDocumentUrls();
								}
							}, false);
						}, "arraybuffer");
					}, endCallback : function(_api) {
						if (_isReplaced)
							_blockStd.MoveCursorToContentControl(false);
						else
							_blockStd.MoveCursorOutsideForm(false);
						
						var _worker = _api.__content_control_worker;
						if (_worker.documents[_worker.current]["Props"])
							_blockStd.SetContentControlPr({ Lock : _worker.documents[_worker.current]["Props"]["Lock"] });
						_worker = null;
						_blockStd = null;

						window.g_asc_plugins.api.asc_Recalculate(true);
						setTimeout(function() { window.g_asc_plugins.api.__content_control_worker.run(); }, 1);
					}};
					var options = new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.CANVAS_WORD);
					options.isNaturalDownload = true;
					if (_current["Url"]) {
						options.errorDirect = Asc.c_oAscError.ID.DirectUrl;
					}
					this.api.downloadAs(Asc.c_oAscAsyncAction.DownloadAs, options);
					return;
				}
				else if (_current["Script"] !== undefined)
				{
					if ( !AscCommon.isValidJs( _current["Script"] ) )
					{
						console.error('Invalid JS.');
						return;	
					}
					
					AscCommon.safePluginEval(_current["Script"]);
					
					if (_isReplaced)
						_blockStd.MoveCursorToContentControl(false);
					else
						_blockStd.MoveCursorOutsideForm(false);
					
					var _worker = _api.__content_control_worker;
					if (_worker.documents[_worker.current]["Props"])
						_blockStd.SetContentControlPr({ Lock : _worker.documents[_worker.current]["Props"]["Lock"] });
					_worker = null;

					var _fonts         = LogicDocument.Document_Get_AllFontNames();
					var _imagesArray   = LogicDocument.Get_AllImageUrls();
					var _images        = {};
					for (var i = 0; i < _imagesArray.length; i++)
					{
						_images[_imagesArray[i]] = _imagesArray[i];
					}

					window.g_asc_plugins.images_rename = _images;
					AscCommon.Check_LoadingDataBeforePrepaste(window.g_asc_plugins.api, _fonts, _images,
						function()
						{
							var _api = window.g_asc_plugins.api;
							_api.WordControl.m_oLogicDocument.Reassign_ImageUrls(window.g_asc_plugins.images_rename);
							delete window.g_asc_plugins.images_rename;
							_api.asc_Recalculate(true);
							_api.WordControl.m_oLogicDocument.UnlockPanelStyles(true);

							setTimeout(function() { window.g_asc_plugins.api.__content_control_worker.run(); }, 1);
						});

					return;
				}
				else if (_current["Props"])
				{
					// не дублируем код!!!
					// меняем тип - тогда настройки изменятся дальше
					type = 1;
				}
			}

			if (1 === type)
			{
				if (_current["Props"] && _current["Url"] === undefined && _current["Script"] === undefined)
				{
					// change properties
					var _blockStd = LogicDocument.GetContentControl(_current["Props"]["InternalId"]);

					if (_blockStd)
					{
						_content_control_pr = new AscCommon.CContentControlPr();
						AscCommon.readContentControlCommonPr(_content_control_pr, _current["Props"]);
						_content_control_pr.InternalId = _current["Props"]["InternalId"];
						
						_blockStd.SetContentControlPr(_content_control_pr);
						LogicDocument.Recalculate();

						_obj = _blockStd.GetContentControlPr();
						this.returnDocuments.push({
							"Tag":        _obj.Tag,
							"Id":         _obj.Id,
							"Lock":       _obj.Lock,
							"InternalId": _obj.InternalId,
							"Alias": _obj.Alias,
							"Appearance": _obj.Appearance
						});
					}
				}
			}

			private_nextFile();
		};

		this.run = function()
		{
			++this.current;

			var LogicDocument = this.api.WordControl.m_oLogicDocument;
			if (0 === this.current)
				LogicDocument.StartAction(AscDFH.historydescription_Document_InsertDocumentsByUrls);

			if (this.current >= this.documents.length)
			{
				this.end();
				return;
			}

			// функция LogicDocument.IsSelectionLocked при указаном callback(isLock)
			// возвращает ВСЕГДА true. реальное значение - ТОЛЬКО в callback
			// callback должен вызываться ВСЕГДА

			function private_CheckLockCallback(isLocked) {
				window.g_asc_plugins.api.__content_control_worker.checkCurrentSelectionLockCallback(isLocked, 0);
			}

			// проверяем лок на вставку - все действия - в коллбэке
			var _current = this.documents[this.current];
			if (undefined === _current["Props"])
				_current["Props"] = {};
			
			var _internalId     = _current["Props"]["InternalId"];
			var _contentControl = _internalId ? g_oTableId.Get_ById(_internalId) : null;
			if ((_current["Url"] !== undefined || _current["Script"] !== undefined) && _contentControl)
			{
				LogicDocument.IsSelectionLocked(AscCommon.changestype_None, {
					Type      : AscCommon.changestype_2_ElementsArray_and_Type,
					Elements  : [_contentControl],
					CheckType : AscCommon.changestype_Document_Content_Add
				}, false, false, private_CheckLockCallback);
			}
			else
			{
				LogicDocument.IsSelectionLocked(AscCommon.changestype_Document_Content_Add,
					undefined, false, false, private_CheckLockCallback);
			}
		};

		this.delete = function()
		{
			var LogicDocument = this.api.WordControl.m_oLogicDocument;

			var arrContentControl = [];
			for (var i = 0; i < this.documents.length; i++)
			{
				var oContentControl = g_oTableId.Get_ById(this.documents[i]["InternalId"]);
				if (oContentControl
					&& (oContentControl instanceof AscCommonWord.CBlockLevelSdt
					|| oContentControl instanceof AscCommonWord.CInlineLevelSdt))
					arrContentControl.push(g_oTableId.Get_ById(this.documents[i]["InternalId"]));
			}

			LogicDocument.SetCheckContentControlsLock(false);
			if (false === LogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
					Type      : AscCommon.changestype_2_ElementsArray_and_Type,
					Elements  : arrContentControl,
					CheckType : AscCommon.changestype_Remove
				}))
			{
				LogicDocument.StartAction(AscDFH.historydescription_Document_InsertDocumentsByUrls);
				for (var i = 0; i < this.documents.length; i++)
				{
					LogicDocument.RemoveContentControl(this.documents[i]["InternalId"]);
				}
				LogicDocument.FinalizeAction();
			}
			LogicDocument.SetCheckContentControlsLock(true);
			this.api.asc_Recalculate(true);
			delete this.api.__content_control_worker;
		};
	}

    AscCommon.CContentControlPluginWorker = CContentControlPluginWorker;

	// пользоваться так:
	// подрубить его последним из скриптов к страничке
	// и вызвать, после подгрузки (конец метода OnInit <- Drawing/HtmlPage.js)
	// var _api = new asc_docs_api();
	// _api.init(oWordControl);

	/**
	 *
	 * @param config
	 * @constructor
	 * @extends {AscCommon.baseEditorsApi}
	 */
	function asc_docs_api(config)
	{
		AscCommon.baseEditorsApi.call(this, config, AscCommon.c_oEditorId.Word);

		/************ private!!! **************/
		this.WordControl = null;

		this.documentFormatSave = c_oAscFileType.DOCX;

		//todo убрать из native, copypaste, chart, loadfont
		this.InterfaceLocale = null;

		this.ShowParaMarks        = false;
		this.ShowSmartGuides        = true;
		this.isAddSpaceBetweenPrg = false;
		this.isPageBreakBefore    = false;
		this.isKeepLinesTogether  = false;

		this.isMarkerFormat             = false;
		this.isStartAddShape            = false;
		this.isDrawTablePen				= false;
        this.isDrawTableErase			= false;
		this.addShapePreset             = "";
		this.isShowTableEmptyLine       = true;
		this.isShowTableEmptyLineAttack = false;

		this.isApplyChangesOnOpen        = false;

		this.IsSpellCheckCurrentWord = false;

		this.mailMergeFileData = null;
		this.insertDocumentUrlsData = null;

		this.isCoMarksDraw  = false;
		this.tmpCoMarksDraw = false;
		this.tmpViewRulers  = null;
		this.tmpZoomType    = null;
		this.tmpDocumentUnits = null;

		// это чтобы сразу показать ридер, без возможности вернуться в редактор/вьюер
		this.isOnlyReaderMode = false;

		/**************************************/

		this.bInit_word_control = false;
		this.isDocumentModify   = false;

		this.tmpFontRenderingMode = null;
		this.FontAsyncLoadType    = 0;
		this.FontAsyncLoadParam   = null;

		this.isPasteFonts_Images = false;

		this.pasteCallback       = null;
		this.pasteImageMap       = null;
		this.EndActionLoadImages = 0;

		this.isSaveFonts_Images = false;
		this.saveImageMap       = null;

		this.isLoadImagesCustom = false;
		this.loadCustomImageMap = null;

		this.ServerImagesWaitComplete = false;

		this.DocumentOrientation = false;

		this.SelectedObjectsStack = [];

		this.nCurPointItemsLength = -1;
		this.isDocumentEditor     = true;

		this.CurrentTranslate = null;

		this.CollaborativeMarksShowType = c_oAscCollaborativeMarksShowType.All;
		
		this.stylePainter = null;
		
		this.forceSaveDisconnectRequest = false;

		// объекты, нужные для отправки в тулбар (шрифты, стили)
		this._gui_control_colors = null;

		this.DocumentReaderMode = null;

		if (window.editor == undefined)
		{
			window.editor = this;
			window['editor'] = window.editor;
			Asc['editor'] = Asc.editor = this;

			if (window["NATIVE_EDITOR_ENJINE"])
				editor = window.editor;
		}

		this.RevisionChangesStack = [];

		this.isDarkMode = false;

		this.TableStylesPreviewGenerator = null;
		
		this.isHandMode = false;

		//g_clipboardBase.Init(this);
		this.headingsColor = config["headings-color"] ? config["headings-color"] : null;
		
		this._init();
	}

	asc_docs_api.prototype = Object.create(AscCommon.baseEditorsApi.prototype);
	asc_docs_api.prototype.constructor = asc_docs_api;

	asc_docs_api.prototype.sendEvent           = function()
	{
		this.sendInternalEvent.apply(this, arguments);
		var name = arguments[0];
		if (_callbacks.hasOwnProperty(name))
		{
			for (var i = 0; i < _callbacks[name].length; ++i)
			{
				_callbacks[name][i].apply(this || window, Array.prototype.slice.call(arguments, 1));
			}
			return true;
		}
		return false;
	};

	asc_docs_api.prototype.getEditorErrorInfo = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return "DocumentInfo: no logic document";

		let result = "DocumentInfo:";

		if (logicDocument.Action.Start)
			result += "\n Action: " + logicDocument.Action.Description;

		result += "\n Selection: " + logicDocument.IsSelectionUse();

		return result;
	};

	asc_docs_api.prototype.SetCollaborativeMarksShowType = function(Type)
	{
		if (c_oAscCollaborativeMarksShowType.None !== this.CollaborativeMarksShowType && c_oAscCollaborativeMarksShowType.None === Type && this.WordControl && this.WordControl.m_oLogicDocument)
		{
			this.CollaborativeMarksShowType = Type;
			AscCommon.CollaborativeEditing.Clear_CollaborativeMarks(true);
		}
		else
		{
			this.CollaborativeMarksShowType = Type;
		}
	};

	asc_docs_api.prototype.GetCollaborativeMarksShowType = function(Type)
	{
		return this.CollaborativeMarksShowType;
	};
	/**
	 * @returns {?CDocument}
	 */
	asc_docs_api.prototype.private_GetLogicDocument = function()
	{
		if (!this.WordControl || !this.WordControl.m_oLogicDocument)
			return null;

		return this.WordControl.m_oLogicDocument;
	};
	asc_docs_api.prototype.private_GetFormsManager = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return null;

		return oLogicDocument.GetFormsManager();
	};

	asc_docs_api.prototype.isLongAction = function()
	{
		if (this.WordControl.m_oLogicDocument)
			return (0 !== this.IsLongActionCurrent || this.WordControl.m_oLogicDocument.IsActionStarted());

		return (0 !== this.IsLongActionCurrent);
	};

	asc_docs_api.prototype.Clear_CollaborativeMarks = function()
	{
		AscCommon.CollaborativeEditing.Clear_CollaborativeMarks(true);
	};

	asc_docs_api.prototype.SetLanguage = function(langId)
	{
		langId = langId.toLowerCase();
		if (undefined !== AscCommonWord.translations_map[langId])
			this.CurrentTranslate = AscCommonWord.translations_map[langId];
	};

	asc_docs_api.prototype.TranslateStyleName   = function(style_name)
	{
		var ret = this.CurrentTranslate.DefaultStyles[style_name];

		if (ret !== undefined)
			return ret;

		return style_name;
	};
	asc_docs_api.prototype.CheckChangedDocument = function()
	{
		if (true === AscCommon.History.Have_Changes())
		{
			// дублирование евента. когда будет undo-redo - тогда
			// эти евенты начнут отличаться
			this.SetDocumentModified(true);
		}
		else
		{
			this.SetDocumentModified(false);
		}

		this._onUpdateDocumentCanSave();
	};
	asc_docs_api.prototype.SetUnchangedDocument = function()
	{
		this.SetDocumentModified(false);
		this._onUpdateDocumentCanSave();
	};

	asc_docs_api.prototype.SetDocumentModified = function(bValue)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (logicDocument && logicDocument.getCompositeInput().isInProgress())
			return;

		this.isDocumentModify = bValue;
		this.sendEvent("asc_onDocumentModifiedChanged");

		if (undefined !== window["AscDesktopEditor"])
		{
			window["AscDesktopEditor"]["onDocumentModifiedChanged"](bValue);
		}
	};

	asc_docs_api.prototype.isDocumentModified = function()
	{
		if (!this.canSave)
		{
			// Пока идет сохранение, мы не закрываем документ
			return true;
		}
		return this.isDocumentModify;
	};
	
	asc_docs_api.prototype.asc_getCanUndo = function()
	{
		if (!this.canUndoRedoByRestrictions())
			return false;
		
		if (AscCommon.History.Can_Undo())
			return true;
		
		if (AscCommon.CollaborativeEditing && AscCommon.CollaborativeEditing.Is_Fast() && !AscCommon.CollaborativeEditing.Is_SingleUser())
			return AscCommon.CollaborativeEditing.CanUndo();
		
		return false;
	};
	asc_docs_api.prototype.asc_getCanRedo = function()
	{
		if (!this.canUndoRedoByRestrictions())
			return false;
		
		return AscCommon.History.Can_Redo();
	};

	asc_docs_api.prototype.sync_BeginCatchSelectedElements = function()
	{
		if (0 != this.SelectedObjectsStack.length)
			this.SelectedObjectsStack.splice(0, this.SelectedObjectsStack.length);

		if (this.WordControl && this.WordControl.m_oDrawingDocument)
			this.WordControl.m_oDrawingDocument.StartTableStylesCheck();
	};
	asc_docs_api.prototype.sync_EndCatchSelectedElements = function(isExternalTrigger)
	{
		if (this.WordControl && this.WordControl.m_oDrawingDocument)
			this.WordControl.m_oDrawingDocument.EndTableStylesCheck();

		this.sendEvent("asc_onFocusObject", this.SelectedObjectsStack, !isExternalTrigger);
	};
	asc_docs_api.prototype.getSelectedElements             = function(bUpdate)
	{
		if (true === bUpdate)
			this.WordControl.m_oLogicDocument.Document_UpdateInterfaceState(true);

		return this.SelectedObjectsStack;
	};
	asc_docs_api.prototype.sync_ChangeLastSelectedElement  = function(type, obj)
	{
		var oUnkTypeObj = null;

		switch (type)
		{
			case c_oAscTypeSelectElement.Paragraph:
				oUnkTypeObj = new Asc.asc_CParagraphProperty(obj);
				break;
			case c_oAscTypeSelectElement.Image:
				oUnkTypeObj = new asc_CImgProperty(obj);
				break;
			case c_oAscTypeSelectElement.Table:
				oUnkTypeObj = new Asc.CTableProp(obj);
				break;
			case c_oAscTypeSelectElement.Header:
				oUnkTypeObj = new CHeaderProp(obj);
				break;
			case Asc.c_oAscTypeSelectElement.UnProtectedRegion:
				oUnkTypeObj = new Asc.RangePermProp(obj);
				break;
		}

		var _i       = this.SelectedObjectsStack.length - 1;
		var bIsFound = false;
		while (_i >= 0)
		{
			if (this.SelectedObjectsStack[_i].Type == type)
			{

				this.SelectedObjectsStack[_i].Value = oUnkTypeObj;
				bIsFound                            = true;
				break;
			}
			_i--;
		}

		if (!bIsFound)
		{
			this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(type, oUnkTypeObj);
		}
	};

	asc_docs_api.prototype.Init = function()
	{
		this.WordControl.Init();
	};

	asc_docs_api.prototype.asc_setLocale = function(val)
	{
		this.InterfaceLocale = val;
	};

	asc_docs_api.prototype.asc_getLocale = function()
	{
		return this.InterfaceLocale;
	};

	asc_docs_api.prototype.ChangeReaderMode  = function()
	{
		return this.WordControl.ChangeReaderMode();
	};
	asc_docs_api.prototype.SetReaderModeOnly = function()
	{
		this.isOnlyReaderMode                       = true;
		if (this.ImageLoader)
			this.ImageLoader.bIsAsyncLoadDocumentImages = false;
	};
	asc_docs_api.prototype.IncreaseReaderFontSize = function()
	{
		return this.WordControl.IncreaseReaderFontSize();
	};
	asc_docs_api.prototype.DecreaseReaderFontSize = function()
	{
		return this.WordControl.DecreaseReaderFontSize();
	};

	asc_docs_api.prototype.SetMobileTopOffset = function(offset, offsetScrollTop)
	{
		if (!this.WordControl)
		{
			this.startMobileOffset = { offset : offset, offsetScrollTop : offsetScrollTop };
		}
		else
		{
			if (!this.WordControl.IsInitControl && !this.isDocumentRenderer())
			{
				this.startMobileOffset = { offset : offset, offsetScrollTop : offsetScrollTop };
				return;
			}

			this.WordControl.setOffsetTop(offset, offsetScrollTop);
		}
	};

	asc_docs_api.prototype.CreateCSS = function()
	{
		var _head = document.getElementsByTagName('head')[0];

		var style0       = document.createElement('style');
		style0.type      = 'text/css';
		style0.innerHTML = ".block_elem { position:absolute;padding:0;margin:0; }";
		_head.appendChild(style0);

		var style2       = document.createElement('style');
		style2.type      = 'text/css';
		style2.innerHTML = ".buttonRuler {\
background-image: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAwCAYAAAAYX/pXAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwwAADsMBx2+oZAAAABp0RVh0U29mdHdhcmUAUGFpbnQuTkVUIHYzLjUuMTAw9HKhAAABhElEQVRIS62Uwa6CMBBF/VQNQcOCBS5caOICApEt3+Wv+AcmfQ7pbdreqY+CJifTdjpng727aZrMFmbB+/3erYEE+/3egMPhMPP57QR/EJCgKAoTs1hQlqURjsdjAESyPp1O7pwEVVWZ1+s1VyB7DemRoK5rN+CvNaRPgqZpgqHz+UwSnEklweVyCQbivX8mlQTX65UGfG63m+vLXRLc7/ekQHoAexK0bWs0uq5TKwli8Afq+94Mw+CQPe78K5D6eDzMOI4GVcCdr4IlOMEWfiP4fJpVkEDLA38ghgR+DgB/ICYQ5OYBCez7d1mAvQZ6gcBmAK010A8ENg8c9u2rZ6iBwL51R7z3z1ADgc2DJDYPZnA3ENi3rhLlgauBAO8/JpUHJEih5QF6iwRaHqC3SPANJ9jCbwTP53MVJNDywB+IIYGfA8AfiAkEqTyQDEAO+HlAgtw8IEFuHpAgNw9IkJsHJMjNAxLk5gEJ8P5jUnlAghRaHqC3SKDlAXqLBN9wgvVM5g/dFuEU6U2wnAAAAABJRU5ErkJggg==);\
background-position: 0px 0px;\
background-repeat: no-repeat;\
}";
		_head.appendChild(style2);

		var style3       = document.createElement('style');
		style3.type      = 'text/css';
		style3.innerHTML = ".buttonPrevPage {\
background-image: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAABgBAMAAADm/++TAAAABGdBTUEAALGPC/xhBQAAABJQTFRFAAAA////UVNVu77Cenp62Nrc3x8hMQAAAAF0Uk5TAEDm2GYAAABySURBVCjPY2AgETDBGEoKUAElJcJSxANjKGAwDQWDYAKMIBhDSRXCCFJSIixF0GS4M+AMExcwcCbAcIQxBEUgDEdBQcJSBE2GO4PU6IJHASxS4NGER4p28YWIAlikwKMJjxTt4gsRBbBIgUcTHini4wsAwMmIvYZODL0AAAAASUVORK5CYII=);\
background-position: 0px 0px;\
background-repeat: no-repeat;\
}";
		_head.appendChild(style3);

		var style4       = document.createElement('style');
		style4.type      = 'text/css';
		style4.innerHTML = ".buttonNextPage {\
background-image: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAABgBAMAAADm/++TAAAABGdBTUEAALGPC/xhBQAAABJQTFRFAAAA////UVNVu77Cenp62Nrc3x8hMQAAAAF0Uk5TAEDm2GYAAABySURBVCjPY2AgETDBGEoKUAElJcJSxANjKGAwDQWDYAKMIBhDSRXCCFJSIixF0GS4M+AMExcwcCbAcIQxBEUgDEdBQcJSBE2GO4PU6IJHASxS4NGER4p28YWIAlikwKMJjxTt4gsRBbBIgUcTHini4wsAwMmIvYZODL0AAAAASUVORK5CYII=);\
background-position: 0px -48px;\
background-repeat: no-repeat;\
}";
		_head.appendChild(style4);
	};

	asc_docs_api.prototype.CreateComponents = function()
	{
		this.asc_setSkin(this.skinObject);
		delete this.skinObject;

		this.CreateCSS();

		if (this.HtmlElement != null)
			this.HtmlElement.innerHTML = ("<div id=\"id_main\" class=\"block_elem\" style=\"touch-action:none;-ms-touch-action: none;-moz-user-select:none;-khtml-user-select:none;user-select:none;background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";overflow:hidden;\" UNSELECTABLE=\"on\">\
								<div id=\"id_panel_left\" class=\"block_elem\">\
									<canvas id=\"id_buttonTabs\" class=\"block_elem\"></canvas>\
									<canvas id=\"id_vert_ruler\" class=\"block_elem\"></canvas>\
								</div>\
									<div id=\"id_panel_top\" class=\"block_elem\">\
									<canvas id=\"id_hor_ruler\" class=\"block_elem\"></canvas>\
									</div>\
                                    <div id=\"id_main_view\" class=\"block_elem\" style=\"touch-action:none;overflow:hidden\">\
                                        <canvas id=\"id_viewer\" class=\"block_elem\" style=\"touch-action:none;-ms-touch-action: none;-webkit-user-select: none; background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";z-index:1\"></canvas>\
									    <canvas id=\"id_viewer_overlay\" class=\"block_elem\" style=\"touch-action:none;-ms-touch-action: none;-webkit-user-select: none; z-index:2\"></canvas>\
									    <div id=\"id_target_cursor\" class=\"block_elem\" width=\"1\" height=\"1\" style=\"touch-action:none;-ms-touch-action: none;-webkit-user-select: none;width:2px;height:13px;z-index:4;\"></div>\
                                    </div>\
								</div>\
									<div id=\"id_panel_right\" class=\"block_elem\" style=\"touch-action:none;margin-right:1px;background-color:" + AscCommon.GlobalSkin.ScrollBackgroundColor + ";\">\
									<div id=\"id_buttonRulers\" class=\"block_elem buttonRuler\"></div>\
									<div id=\"id_vertical_scroll\" style=\"left:0;top:0px;width:14px;overflow:hidden;position:absolute;\">\
									</div>\
									<div id=\"id_buttonPrevPage\" class=\"block_elem buttonPrevPage\"></div>\
									<div id=\"id_buttonNextPage\" class=\"block_elem buttonNextPage\"></div>\
								</div>\
									<div id=\"id_horscrollpanel\" class=\"block_elem\" style=\"touch-action:none;margin-bottom:1px;background-color:" + AscCommon.GlobalSkin.ScrollBackgroundColor + ";\">\
									<div id=\"id_horizontal_scroll\" style=\"left:0px;top:0;height:14px;overflow:hidden;position:absolute;width:100%;\">\
									</div>\
									</div>" + this.HtmlElement.innerHTML);
	};

	asc_docs_api.prototype.asc_setSkin = function(theme)
	{
		AscCommon.updateGlobalSkin(theme);

		var obj_id_main = document.getElementById("id_main");
		if (obj_id_main)
		{
			obj_id_main.style.backgroundColor = AscCommon.GlobalSkin.BackgroundColor;
			document.getElementById("id_viewer").style.backgroundColor = AscCommon.GlobalSkin.BackgroundColor;
			document.getElementById("id_panel_right").style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor;
			document.getElementById("id_horscrollpanel").style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor;
		}

		if (this.WordControl && this.WordControl.m_oBody)
		{
			this.WordControl.OnResize(true);
			if (this.WordControl.m_oEditor && this.WordControl.m_oEditor.HtmlElement)
			{
				this.WordControl.m_oEditor.HtmlElement.fullRepaint = true;
				this.WordControl.OnScroll();
			}
		}
	};

	asc_docs_api.prototype.GetCopyPasteDivId = function()
	{
		if (this.isMobileVersion)
			return this.WordControl.Name;
		return "";
	};
	
	asc_docs_api.prototype.ContentToHTML = function()
	{
		if (!this.WordControl || !this.WordControl.m_oLogicDocument)
			return "";
		
		this.DocumentReaderMode = new AscCommon.CDocumentReaderMode();
		
		this.WordControl.m_oLogicDocument.SelectAll();
		
		var text_data = {
			data     : "",
			pushData : function(format, value)
			{
				this.data = value;
			}
		};
		
		this.asc_CheckCopy(text_data, 2);
		
		this.WordControl.m_oLogicDocument.RemoveSelection();
		
		return text_data.data;
	};
	asc_docs_api.prototype.executeShortcut = function(type)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;
		
		return logicDocument.executeShortcut(type);
	};
	asc_docs_api.prototype._InitCommonShortcuts = function () 
	{
		// ActionType, Key, Ctrl, Shift, Alt
		this.initShortcuts([
			[c_oAscDocumentShortcutType.InsertPageBreak, 13, true, false, false],
			[c_oAscDocumentShortcutType.InsertLineBreak, 13, false, true, false],
			[c_oAscDocumentShortcutType.InsertColumnBreak, 13, true, true, false],
			[c_oAscDocumentShortcutType.ResetChar, 32, true, false, false],
			[c_oAscDocumentShortcutType.NonBreakingSpace, 32, true, true, false],
			[c_oAscDocumentShortcutType.ShowAll, 56, true, true, false],
			[c_oAscDocumentShortcutType.EditSelectAll, 65, true, false, false],
			[c_oAscDocumentShortcutType.Bold, 66, true, false, false],
			[c_oAscDocumentShortcutType.CopyFormat, 67, true, false, true],
			[c_oAscDocumentShortcutType.InsertEndnoteNow, 68, true, false, true],
			[c_oAscDocumentShortcutType.CenterPara, 69, true, false, false],
			[c_oAscDocumentShortcutType.EuroSign, 69, true, false, true],
			[c_oAscDocumentShortcutType.CopyrightSign, 71, true, false, true],
			[c_oAscDocumentShortcutType.Italic, 73, true, false, false],
			[c_oAscDocumentShortcutType.JustifyPara, 74, true, false, false],
			[c_oAscDocumentShortcutType.InsertHyperlink, 75, true, false, false],
			[c_oAscDocumentShortcutType.ApplyListBullet, 76, true, true, false],
			[c_oAscDocumentShortcutType.LeftPara, 76, true, false, false],
			[c_oAscDocumentShortcutType.Indent, 77, true, false, false],
			[c_oAscDocumentShortcutType.UnIndent, 77, true, true, false],
			[c_oAscDocumentShortcutType.PrintPreviewAndPrint, 80, true, false, false],
			[c_oAscDocumentShortcutType.InsertPageNumber, 80, true, true, false],
			[c_oAscDocumentShortcutType.RightPara, 82, true, false, false],
			[c_oAscDocumentShortcutType.RegisteredSign, 82, true, false, true],
			[c_oAscDocumentShortcutType.Save, 83, true, false, false],
			[c_oAscDocumentShortcutType.TrademarkSign, 84, true, false, true],
			[c_oAscDocumentShortcutType.Underline, 85, true, false, false],
			[c_oAscDocumentShortcutType.PasteFormat, 86, true, false, true],
			[c_oAscDocumentShortcutType.EditRedo, 89, true, false, false],
			[c_oAscDocumentShortcutType.EditUndo, 90, true, false, false],
			[c_oAscDocumentShortcutType.EnDash, 109, true, false, false],
			[c_oAscDocumentShortcutType.EmDash, 109, true, false, true],
			[c_oAscDocumentShortcutType.UpdateFields, 120, false, false, false],
			[c_oAscDocumentShortcutType.Superscript, 188, true, false, false],
			[c_oAscDocumentShortcutType.NonBreakingHyphen, 189, true, true, false],
			[c_oAscDocumentShortcutType.NonBreakingHyphen, 173, true, true, false],
			[c_oAscDocumentShortcutType.Subscript, 190, true, false, false],
			[c_oAscDocumentShortcutType.IncreaseFontSize, 221, true, false, false],
			[c_oAscDocumentShortcutType.DecreaseFontSize, 219, true, false, false],
			[c_oAscDocumentShortcutType.SpeechWorker, 90, true, false, true]
		]);
	}

	asc_docs_api.prototype._InitWindowsShortcuts = function ()
	{
		// ActionType, Key, Ctrl, Shift, Alt
		this.initShortcuts([
			[c_oAscDocumentShortcutType.ApplyHeading1, 49, false, false, true],
			[c_oAscDocumentShortcutType.ApplyHeading2, 50, false, false, true],
			[c_oAscDocumentShortcutType.ApplyHeading3, 51, false, false, true],
			[c_oAscDocumentShortcutType.Strikeout, 53, true, false, false],
			[c_oAscDocumentShortcutType.InsertFootnoteNow, 70, true, false, true],
			[c_oAscDocumentShortcutType.InsertEquation, 187, false, false, true],
			[c_oAscDocumentShortcutType.InsertEquation, 61, false, false, true],
			[c_oAscDocumentShortcutType.SoftHyphen, 189, false, false, true],
			[c_oAscDocumentShortcutType.SoftHyphen, 173, false, false, true],
			[c_oAscDocumentShortcutType.HorizontalEllipsis, 190, true, false, true]
		]);
	}

	asc_docs_api.prototype._InitMacOsShortcuts = function () {
		// ActionType, Key, Ctrl, Shift, Alt
		this.initShortcuts([
			[c_oAscDocumentShortcutType.ApplyHeading1, 49, true, false, true],
			[c_oAscDocumentShortcutType.ApplyHeading2, 50, true, false, true],
			[c_oAscDocumentShortcutType.ApplyHeading3, 51, true, false, true],
			[c_oAscDocumentShortcutType.Strikeout, 88, true, true, false],
			[c_oAscDocumentShortcutType.InsertEquation, 187, true, false, true],
			[c_oAscDocumentShortcutType.InsertEquation, 61, true, false, true],
			[c_oAscDocumentShortcutType.SoftHyphen, 189, true, false, true],
			[c_oAscDocumentShortcutType.SoftHyphen, 173, true, false, true]
		]);
	}

	asc_docs_api.prototype.initDefaultShortcuts = function()
	{
		this._InitCommonShortcuts();
		if (AscCommon.AscBrowser.isMacOs) 
		{
			this._InitMacOsShortcuts();
		} 
		else 
		{
			this._InitWindowsShortcuts();
		}
	};

	asc_docs_api.prototype.InitEditor = function()
	{
		this.WordControl.m_oLogicDocument                    = new AscCommonWord.CDocument(this.WordControl.m_oDrawingDocument);
		this.WordControl.m_oDrawingDocument.m_oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!this.isSpellCheckEnable || this.isRestrictionComments() || this.isRestrictionForms())
			this.WordControl.m_oLogicDocument.TurnOffSpellCheck();

		if (this.WordControl.MobileTouchManager)
			this.WordControl.MobileTouchManager.delegate.LogicDocument = this.WordControl.m_oLogicDocument;

		if (this.isRestrictionForms() || this.isRestrictionComments())
		{
			this.ShowParaMarks = false;
			this.WordControl.HideRulers();
		}
	};

	asc_docs_api.prototype.isDocumentRenderer = function()
	{
		return !!this.WordControl.m_oDrawingDocument.m_oDocumentRenderer;
	};
	
	asc_docs_api.prototype.canUndoRedoByRestrictions = function()
	{
		// В режиме вью можно редактировать в разрешенных областях
		return (this.canEdit() || this.isRestrictionComments() || this.isRestrictionForms() || this.isRestrictionView());

	};

	asc_docs_api.prototype["asc_setViewerThumbnailsZoom"] = function(value) {
		if (this.WordControl.m_oDrawingDocument.m_oDocumentRenderer &&
			this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.Thumbnails)
			this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.Thumbnails.setZoom(value);
	};
	asc_docs_api.prototype["asc_setViewerThumbnailsUsePageRect"] = function(value) {
		if (this.WordControl.m_oDrawingDocument.m_oDocumentRenderer &&
			this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.Thumbnails)
			this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.Thumbnails.setIsDrawCurrentRect(value);
	};
	asc_docs_api.prototype["asc_viewerThumbnailsResize"] = function() {
		if (this.WordControl.m_oDrawingDocument.m_oDocumentRenderer &&
			this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.Thumbnails)
			this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.Thumbnails.resize();
	};
	asc_docs_api.prototype["asc_viewerNavigateTo"] = function(value) {
		if (this.WordControl.m_oDrawingDocument.m_oDocumentRenderer)
			this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.navigate(value);
	};
	asc_docs_api.prototype["asc_setViewerTargetType"] = asc_docs_api.prototype.asc_setViewerTargetType = function(type) {
		this.isHandMode = ("hand" === type);
		this.WordControl.checkMouseHandMode();
		this.WordControl.onMouseMove();
		this.sendEvent("asc_onChangeViewerTargetType", this.isHandMode);
	};

	function BeforeOpenDocument()
	{
		this.InitEditor();
		this.DocumentType   = 2;
		this.LoadedObjectDS = this.WordControl.m_oLogicDocument.CopyStyle();

		g_oIdCounter.Set_Load(true);
		AscFonts.IsCheckSymbols = true;
	};
	function AfterOpenDocument(data, size)
	{
		if (History && History.Update_FileDescription)
			History.Update_FileDescription(data, size);

		g_oIdCounter.Set_Load(false);
		this.LoadedObject = 1;

		// проверяем какие шрифты нужны
		var StylesPainter = new AscCommonWord.CStylesPainter();
		StylesPainter.CheckStylesNames(this, this.LoadedObjectDS);

		this.WordControl.m_oDrawingDocument.CheckFontNeeds();
		AscCommon.pptx_content_loader.CheckImagesNeeds(this.WordControl.m_oLogicDocument);

		this.FontLoader.LoadDocumentFonts(this.WordControl.m_oLogicDocument.Fonts, false);

		AscFonts.IsCheckSymbols = false;

		//callback
		editor.DocumentOrientation = (null == editor.WordControl.m_oLogicDocument) ? true : !editor.WordControl.m_oLogicDocument.Orientation;
		var sizeMM;
		if (editor.DocumentOrientation)
			sizeMM = DocumentPageSize.getSize(AscCommon.Page_Width, AscCommon.Page_Height);
		else
			sizeMM = DocumentPageSize.getSize(AscCommon.Page_Height, AscCommon.Page_Width);
		editor.sync_DocSizeCallback(sizeMM.w_mm, sizeMM.h_mm);
		editor.sync_PageOrientCallback(editor.get_DocumentOrientation());

		if (this.isMobileVersion)
		{
			AscCommon.AscBrowser.isSafariMacOs   = false;
			PasteElementsId.PASTE_ELEMENT_ID     = "wrd_pastebin";
			PasteElementsId.ELEMENT_DISPAY_STYLE = "none";
		}
	};
	asc_docs_api.prototype.OpenDocumentFromBin = function(url, gObject)
	{
		BeforeOpenDocument.call(this);

		var openParams        = {};
		var oBinaryFileReader = new AscCommonWord.BinaryFileReader(this.WordControl.m_oLogicDocument, openParams);
		if (!oBinaryFileReader.Read(gObject))
			editor.sendEvent("asc_onError", c_oAscError.ID.MobileUnexpectedCharCount, c_oAscError.Level.Critical);

		AfterOpenDocument.call(this, oBinaryFileReader.stream.data, oBinaryFileReader.stream.size);
		return true;
	};
	asc_docs_api.prototype.OpenDocumentFromZip = function(data)
	{
		BeforeOpenDocument.call(this);

		let res = this.OpenDocumentFromZipNoInit(data);

		AfterOpenDocument.call(this, data, data.length);
		return res;
	};
	asc_docs_api.prototype.OpenDocumentFromZipNoInit = function(data)
	{
		if (!data) {
			return false;
		}
		let xmlParserContext = new AscCommon.XmlParserContext();
		xmlParserContext.DrawingDocument = this.WordControl.m_oDrawingDocument;

		let jsZlib = new AscCommon.ZLib();
		if (!jsZlib.open(data)) {
			return false;
		}

		let reader, openParams = {};
		let oBinaryFileReader = new AscCommonWord.BinaryFileReader(this.WordControl.m_oLogicDocument, openParams);
		//очищать pptx_content_loader не надо, т.к. открываем zip
		oBinaryFileReader.PreLoadPrepare(undefined, false);

		this.WordControl.m_oLogicDocument.fromZip(jsZlib, xmlParserContext, oBinaryFileReader.oReadResult);

		oBinaryFileReader.PostLoadPrepare(xmlParserContext, false);
		jsZlib.close();
		return true;
	};
	// Callbacks
	/* все имена callback'оф начинаются с On. Пока сделаны:
	 OnBold,
	 OnItalic,
	 OnUnderline,
	 OnTextPrBaseline(возвращается расположение строки - supstring, superstring, baseline),
	 OnPrAlign(выравнивание по ширине, правому краю, левому краю, по центру),
	 OnListType( возвращается AscCommon.asc_CListType )

	 фейк-функции ожидающие TODO:
	 Print,Undo,Redo,Copy,Cut,Paste,Share,Save,Download & callbacks
	 OnFontName, OnFontSize, OnLineSpacing

	 OnFocusObject( возвращается массив asc_CSelectedObject )
	 OnInitEditorStyles( возвращается CStylesPainter )
	 OnSearchFound( возвращается CSearchResult );
	 OnParaSpacingLine( возвращается AscCommon.asc_CParagraphSpacing )
	 OnLineSpacing( не используется? )
	 OnTextColor( возвращается AscCommon.CColor )
	 OnTextHighLight( возвращается AscCommon.CColor )
	 OnInitEditorFonts( возвращается массив объектов СFont )
	 OnFontFamily( возвращается asc_CTextFontFamily )
	 */
	var _callbacks = {};

	asc_docs_api.prototype.asc_registerCallback = function(name, callback)
	{
		if (!_callbacks.hasOwnProperty(name))
			_callbacks[name] = [];
		_callbacks[name].push(callback);
	};

	asc_docs_api.prototype.asc_unregisterCallback = function(name, callback)
	{
		if (_callbacks.hasOwnProperty(name))
		{
			for (var i = _callbacks[name].length - 1; i >= 0; --i)
			{
				if (_callbacks[name][i] == callback)
					_callbacks[name].splice(i, 1);
			}
		}
	};

	asc_docs_api.prototype.asc_checkNeedCallback = function(name)
	{
		return _callbacks.hasOwnProperty(name);
	};

	// тут методы, замены евентов
	asc_docs_api.prototype.get_PropertyThemeColors       = function()
	{
		return [this._gui_control_colors.Colors, this._gui_control_colors.StandartColors];
	};
	// -------

	// GoTo
	asc_docs_api.prototype._goToComment = function(data)
	{
		if (this.WordControl && this.WordControl.m_oLogicDocument && data)
		{
			var commentId = this.WordControl.m_oLogicDocument.Comments.GetCommentIdByGuid(data);
			if (commentId) {
				this.asc_selectComment(commentId);
				this.asc_showComment(commentId);
			}
		}
	};
	asc_docs_api.prototype._goToBookmark = function(data)
	{
		if (this.WordControl && this.WordControl.m_oLogicDocument && data)
		{
			this.WordControl.m_oLogicDocument.GoToBookmark(data, true);
		}
	};

	/////////////////////////////////////////////////////////////////////////
	///////////////////CoAuthoring and Chat api//////////////////////////////
	/////////////////////////////////////////////////////////////////////////
	// Init CoAuthoring
	asc_docs_api.prototype._coAuthoringSetChange = function(change, oColor)
	{
		var oChange = new AscCommon.CCollaborativeChanges();
		oChange.Set_Data(change);
		oChange.Set_Color(oColor);
		AscCommon.CollaborativeEditing.Add_Changes(oChange);
	};

	asc_docs_api.prototype._coAuthoringSetChanges = function(e, oColor)
	{
		var Count = e.length;
		for (var Index = 0; Index < Count; ++Index)
			this._coAuthoringSetChange(e[Index], oColor);
	};

	asc_docs_api.prototype._coAuthoringInitEnd = function()
	{
		var t                                        = this;
		this.CoAuthoringApi.onCursor                 = function(e)
		{
			if (true === AscCommon.CollaborativeEditing.Is_Fast())
			{
				t.WordControl.m_oLogicDocument.Update_ForeignCursor(e[e.length - 1]['cursor'], e[e.length - 1]['user'], true, e[e.length - 1]['useridoriginal']);
			}
		};
		this.CoAuthoringApi.onConnectionStateChanged = function(e)
		{
			if (true === AscCommon.CollaborativeEditing.Is_Fast() && false === e['state'])
			{
				t.WordControl.m_oLogicDocument.Remove_ForeignCursor(e['id']);
			}
			t.sendEvent("asc_onConnectionStateChanged", e);
		};
		this.CoAuthoringApi.onLocksAcquired          = function(e)
		{
			if (t._coAuthoringCheckEndOpenDocument(t.CoAuthoringApi.onLocksAcquired, e))
			{
				return;
			}

			if (2 != e["state"])
			{
				var Id    = e["block"];
				var Class = g_oTableId.Get_ById(Id);
				if (null != Class)
				{
					var Lock = Class.Lock;

					// Выставляем ID пользователя, залочившего данный элемент
					Lock.Set_UserId(e["user"]);

					var OldType = Class.Lock.Get_Type();
					if (locktype_Other2 === OldType || locktype_Other3 === OldType)
					{
						Lock.Set_Type(locktype_Other3, true);
					}
					else
					{
						Lock.Set_Type(locktype_Other, true);
					}

					if (Class instanceof AscCommonWord.CHeaderFooterController)
					{
						t.sync_LockHeaderFooters();
					}
					else if (Class instanceof AscCommonWord.CDocument)
					{
						t.sync_LockDocumentProps();
					}
					else if (Class instanceof AscCommon.CComment)
					{
						t.sync_LockComment(Class.Get_Id(), e["user"]);
					}
					else if (Class instanceof AscCommonWord.CGraphicObjects)
					{
						t.sync_LockDocumentSchema();
					}
					else if(Class instanceof AscCommon.CCore)
                    {
                        editor.sendEvent("asc_onLockCore", true);
                    }
					else if(Class instanceof AscCommonWord.CDocProtect)
					{
						editor.sendEvent("asc_onLockDocumentProtection", true);
					}
					// Теперь обновлять состояние необходимо, чтобы обновить локи в режиме рецензирования.
					t.WordControl.m_oLogicDocument.UpdateInterface(undefined, true);
				}
				else
				{
					AscCommon.CollaborativeEditing.Add_NeedLock(Id, e["user"]);
				}
			}
		};
		this.CoAuthoringApi.onLocksReleased          = function(e, bChanges)
		{
			if (t._coAuthoringCheckEndOpenDocument(t.CoAuthoringApi.onLocksReleased, e, bChanges))
			{
				return;
			}

			var Id    = e["block"];
			var Class = g_oTableId.Get_ById(Id);
			if (null != Class)
			{
				var Lock = Class.Lock;
				if ("undefined" != typeof(Lock))
				{
					var CurType = Lock.Get_Type();

					var NewType = locktype_None;

					if (CurType === locktype_Other)
					{
						if (true != bChanges)
						{
							NewType = locktype_None;
						}
						else
						{
							NewType = locktype_Other2;
							AscCommon.CollaborativeEditing.Add_Unlock(Class);
						}
					}
					else if (CurType === locktype_Mine)
					{
						// Такого быть не должно
						NewType = locktype_Mine;
					}
					else if (CurType === locktype_Other2 || CurType === locktype_Other3)
					{
						NewType = locktype_Other2;
					}

					Lock.Set_Type(NewType, true);

					// Теперь обновлять состояние необходимо, чтобы обновить локи в режиме рецензирования.
					t.WordControl.m_oLogicDocument.UpdateInterface(undefined, true);

					if (Class instanceof AscCommonWord.CHeaderFooterController)
					{
						if (NewType !== locktype_Mine && NewType !== locktype_None)
						{
							t.sync_LockHeaderFooters();
						}
						else
						{
							t.sync_UnLockHeaderFooters();
						}
					}
					else if (Class instanceof AscCommonWord.CDocument)
					{
						if (NewType !== locktype_Mine && NewType !== locktype_None)
						{
							t.sync_LockDocumentProps();
						}
						else
						{
							t.sync_UnLockDocumentProps();
						}
					}
					else if (Class instanceof AscCommon.CComment)
					{
						if (NewType !== locktype_Mine && NewType !== locktype_None)
						{
							t.sync_LockComment(Class.Get_Id(), e["user"]);
						}
						else
						{
							t.sync_UnLockComment(Class.Get_Id());
						}
					}
					else if (Class instanceof AscCommonWord.CGraphicObjects)
					{
						if (NewType !== locktype_Mine && NewType !== locktype_None)
						{
							t.sync_LockDocumentSchema();
						}
						else
						{
							t.sync_UnLockDocumentSchema();
						}
					}
					else if(Class instanceof AscCommon.CCore)
					{
						if (NewType !== locktype_Mine && NewType !== locktype_None)
						{
							editor.sendEvent("asc_onLockCore", true);
						}
						else
						{
							editor.sendEvent("asc_onLockCore", false);
						}
					}
					else if(Class instanceof AscCommonWord.CDocProtect)
					{
						if (NewType !== locktype_Mine && NewType !== locktype_None)
						{
							editor.sendEvent("asc_onLockDocumentProtection", true);
						}
						else
						{
							editor.sendEvent("asc_onLockDocumentProtection", false);
						}
					}
				}
			}
			else
			{
				AscCommon.CollaborativeEditing.Remove_NeedLock(Id);
			}
		};
		this.CoAuthoringApi.onSaveChanges            = function(e, userId, bFirstLoad)
		{
			var bUseColor;
			if (bFirstLoad)
			{
				bUseColor = -1 === AscCommon.CollaborativeEditing.m_nUseType;
			}
			if (t.CollaborativeMarksShowType === c_oAscCollaborativeMarksShowType.None)
			{
				bUseColor = false;
			}

			var oCommonColor = AscCommon.getUserColorById(userId, null, false, false);
			var oColor       = false === bUseColor ? null : oCommonColor;
			t._coAuthoringSetChange(e, oColor);
			// т.е. если bSendEvent не задан, то посылаем  сообщение + когда загрузился документ
			if (t._canSyncCollaborativeChanges(bFirstLoad))
				t.sync_CollaborativeChanges();
		};
		this.CoAuthoringApi.onChangesIndex = function(changesIndex)
		{
			let count = AscCommon.CollaborativeEditing.GetAllChangesCount() - changesIndex;
			if (count < 0)
			{
				// This should never happen
				return;
			}
			
			if (!t.isViewMode && count > 0)
			{
				// TODO: This shouldn't happen too with the current behaviour, but it actually can be handled
				return;
			}
			
			AscCommon.CollaborativeEditing.UndoGlobal(count);
		};
		this.CoAuthoringApi.onRecalcLocks            = function(e)
		{
			if (e && true === AscCommon.CollaborativeEditing.Is_Fast())
			{
				AscCommon.CollaborativeEditing.UpdateForeignCursorByAdditionalInfo(JSON.parse(e));
			}
		};
	};
	
	asc_docs_api.prototype._canSyncCollaborativeChanges = function(isFirstLoad)
	{
		return (!isFirstLoad && this.bInit_word_control);
	};

	asc_docs_api.prototype.startCollaborationEditing = function()
		{
		AscCommon.CollaborativeEditing.Start_CollaborationEditing();
		if (this.WordControl && this.WordControl.m_oLogicDocument)
		{
			this.WordControl.m_oLogicDocument.StartCollaborationEditing();
			}
		this.asc_setDrawCollaborationMarks(true);
		if(window['AscCommon'].g_specialPasteHelper && AscCommon.CollaborativeEditing.m_bFast)
		{
						window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide();
					}
		};
	asc_docs_api.prototype.endCollaborationEditing = function()
		{
		AscCommon.CollaborativeEditing.End_CollaborationEditing();
		this.asc_setDrawCollaborationMarks(false);
		};

	//----------------------------------------------------------------------------------------------------------------------
	// SpellCheck_CallBack
	//          Функция ответа от сервера.
	//----------------------------------------------------------------------------------------------------------------------
	asc_docs_api.prototype.SpellCheck_CallBack = function(Obj)
	{
		if (undefined != Obj && undefined != Obj["ParagraphId"])
		{
			var ParaId    = Obj["ParagraphId"];
			var Paragraph = g_oTableId.Get_ById(ParaId);
			var Type      = Obj["type"];
			if (null != Paragraph)
			{
				if ("spell" === Type)
				{
					Paragraph.SpellChecker.SpellCheckResponse(Obj["RecalcId"], Obj["usrCorrect"]);
					Paragraph.ReDraw();
				}
				else if ("suggest" === Type)
				{
					Paragraph.SpellChecker.SuggestResponse(Obj["RecalcId"], Obj["ElementId"], Obj["usrSuggest"][0]);
					this.sync_SpellCheckVariantsFound();
				}
			}
			else
			{
				this.WordControl.m_oLogicDocument.Spelling.RemoveWaitingParagraphById(ParaId);
			}
		}
	};

	asc_docs_api.prototype._spellCheckDisconnect   = function()
	{
		if (this.WordControl.m_oLogicDocument)
			this.WordControl.m_oLogicDocument.TurnOffSpellCheck();
	};
	asc_docs_api.prototype._onUpdateDocumentCanSave   = function()
	{
		var CollEditing = AscCommon.CollaborativeEditing;

		// Можно модифицировать это условие на более быстрое (менять самим состояние в аргументах, а не запрашивать каждый раз)
		var isCanSave = this.isDocumentModified() || (true !== CollEditing.Is_SingleUser() && 0 !== CollEditing.getOwnLocksLength());

		if (true === CollEditing.Is_Fast() && true !== CollEditing.Is_SingleUser())
			isCanSave = false;

		if (isCanSave !== this.isDocumentCanSave)
		{
			this.isDocumentCanSave = isCanSave;
			this.sendEvent('asc_onDocumentCanSaveChanged', this.isDocumentCanSave);
		}
	};
	asc_docs_api.prototype._onUpdateDocumentCanUndoRedo = function ()
	{
		if (this.WordControl && this.WordControl.m_oLogicDocument)
			this.WordControl.m_oLogicDocument.Document_UpdateUndoRedoState();
	};

	// get functions
	// Возвращает
	//{
	// ParaPr :
	// {
	//    ContextualSpacing : false,            // Удалять ли интервал между параграфами одинакового стиля
	//
	//    Ind :
	//    {
	//        Left      : 0,                    // Левый отступ
	//        Right     : 0,                    // Правый отступ
	//        FirstLine : 0                     // Первая строка
	//    },
	//
	//    Jc : align_Left,                      // Прилегание параграфа
	//
	//    KeepLines : false,                    // переносить параграф на новую страницу,
	//                                          // если на текущей он целиком не убирается
	//    KeepNext  : false,                    // переносить параграф вместе со следующим параграфом
	//
	//    PageBreakBefore : false,              // начинать параграф с новой страницы
	//
	//    Spacing :
	//    {
	//        Line     : 1.15,                  // Расстояние между строками внутри абзаца
	//        LineRule : linerule_Auto,         // Тип расстрояния между строками
	//        Before   : 0,                     // Дополнительное расстояние до абзаца
	//        After    : 10 * g_dKoef_pt_to_mm  // Дополнительное расстояние после абзаца
	//    },
	//
	//    Shd :
	//    {
	//        Value : shd_Nil,
	//        Color :
	//        {
	//            r : 255,
	//            g : 255,
	//            b : 255
	//        }
	//    },
	//
	//    WidowControl : true,                  // Запрет висячих строк
	//
	//    Tabs : []
	// },
	//
	// TextPr :
	// {
	//    Bold       : false,
	//    Italic     : false,
	//    Underline  : false,
	//    Strikeout  : false,
	//    FontFamily :
	//    {
	//        Name  : "Times New Roman",
	//        Index : -1
	//    },
	//    FontSize   : 12,
	//    Color      :
	//    {
	//        r : 0,
	//        g : 0,
	//        b : 0
	//    },
	//    VertAlign : vertalign_Baseline,
	//    HighLight : highlight_None
	// }
	//}


	asc_docs_api.prototype.put_FramePr = function(Obj)
	{
		if (undefined != Obj.FontFamily)
		{
			var loader     = AscCommon.g_font_loader;
			var fontinfo   = g_fontApplication.GetFontInfo(Obj.FontFamily);
			var isasync    = loader.LoadFont(fontinfo, editor.asyncFontEndLoaded_DropCap, Obj);
			Obj.FontFamily = new asc_CTextFontFamily({Name : fontinfo.Name, Index : -1});

			if (false === isasync)
			{
				if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
				{
					this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetFramePrWithFontFamily);
					this.WordControl.m_oLogicDocument.SetParagraphFramePr(Obj);
					this.WordControl.m_oLogicDocument.FinalizeAction();
				}
			}
		}
		else
		{
			if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
			{
				this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetFramePr);
				this.WordControl.m_oLogicDocument.SetParagraphFramePr(Obj);
				this.WordControl.m_oLogicDocument.FinalizeAction();
			}
		}
	};



	asc_docs_api.prototype.asyncFontEndLoaded_DropCap = function(Obj)
	{
		this.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadFont);
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetFramePrWithFontFamilyLong);
			this.WordControl.m_oLogicDocument.SetParagraphFramePr(Obj);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
		// отжать заморозку меню
	};

	asc_docs_api.prototype.asc_addDropCap = function(bInText)
	{
		this.WordControl.m_oLogicDocument.AddDropCap(bInText);
	};

	asc_docs_api.prototype.removeDropcap = function(bDropCap)
	{
		this.WordControl.m_oLogicDocument.RemoveDropCap(bDropCap);
	};

	asc_docs_api.prototype.get_TextProps = function()
	{
		var Doc    = this.WordControl.m_oLogicDocument;
		var ParaPr = Doc.GetCalculatedParaPr();
		var TextPr = Doc.GetCalculatedTextPr();

		// return { ParaPr: ParaPr, TextPr : TextPr };
		return new Asc.CParagraphAndTextProp(ParaPr, TextPr);	// uncomment if this method will be used externally. 20/03/2012 uncommented for testers
	};

	// -------
	asc_docs_api.prototype.GetJSONLogicDocument = function()
	{
		return JSON.stringify(this.WordControl.m_oLogicDocument);
	};

	asc_docs_api.prototype.get_ContentCount = function()
	{
		return this.WordControl.m_oLogicDocument.Content.length;
	};

	asc_docs_api.prototype.select_Element = function(Index)
	{
		var Document = this.WordControl.m_oLogicDocument;

		if (true === Document.Selection.Use)
			Document.RemoveSelection();

		Document.DrawingDocument.SelectEnabled(true);
		Document.DrawingDocument.TargetEnd();

		Document.Selection.Use   = true;
		Document.Selection.Start = false;
		Document.Selection.Flag  = AscCommon.selectionflag_Common;

		Document.Selection.StartPos = Index;
		Document.Selection.EndPos   = Index;

		Document.Content[Index].Selection.Use      = true;
		Document.Content[Index].Selection.StartPos = Document.Content[Index].Internal_GetStartPos();
		Document.Content[Index].Selection.EndPos   = Document.Content[Index].Content.length - 1;

		Document.Selection_Draw();
	};

	asc_docs_api.prototype.UpdateTextPr        = function(TextPr)
	{
		if ("undefined" != typeof(TextPr))
		{
			this.sync_BoldCallBack(TextPr.Bold);
			this.sync_ItalicCallBack(TextPr.Italic);
			this.sync_UnderlineCallBack(TextPr.Underline);
			this.sync_StrikeoutCallBack(TextPr.Strikeout);
			this.sync_TextPrFontSizeCallBack(TextPr.FontSize);
			this.sync_TextPrFontFamilyCallBack(TextPr.FontFamily);
			this.sync_VerticalAlign(TextPr.VertAlign);
			this.sync_TextHighLight(TextPr.HighLight);
			this.sync_TextSpacing(TextPr.Spacing);
			this.sync_TextDStrikeout(TextPr.DStrikeout);
			this.sync_TextCaps(TextPr.Caps);
			this.sync_TextSmallCaps(TextPr.SmallCaps);
			this.sync_TextPosition(TextPr.Position);
			this.sync_TextLangCallBack(TextPr.Lang);
			this.sync_TextColor(TextPr);

			if (this.isMobileVersion)
				this.sendEvent("asc_onTextShd", new Asc.asc_CParagraphShd(TextPr.Shd));
		}
	};
	asc_docs_api.prototype.UpdateParagraphProp = function(ParaPr)
	{
		//if ( true === CollaborativeEditing.Get_GlobalLock() )
		//{
		//    ParaPr.Locked      = true;
		//    ParaPr.CanAddTable = false;
		//}

		// var prgrhPr = this.get_TextProps();
		// var prProp = {};
		// prProp.Ind = prgrhPr.ParaPr.Ind;
		// prProp.ContextualSpacing = prgrhPr.ParaPr.ContextualSpacing;
		// prProp.Spacing = prgrhPr.ParaPr.Spacing;
		// prProp.PageBreakBefore = prgrhPr.ParaPr.PageBreakBefore;
		// prProp.KeepLines = prgrhPr.ParaPr.KeepLines;

		// {
		//    ContextualSpacing : false,            // Удалять ли интервал между параграфами одинакового стиля
		//
		//    Ind :
		//    {
		//        Left      : 0,                    // Левый отступ
		//        Right     : 0,                    // Правый отступ
		//        FirstLine : 0                     // Первая строка
		//    },
		//    Jc : align_Left,                      // Прилегание параграфа
		//    KeepLines : false,                    // переносить параграф на новую страницу,
		//                                          // если на текущей он целиком не убирается
		//    PageBreakBefore : false,              // начинать параграф с новой страницы
		//
		//    Spacing :
		//    {
		//        Line     : 1.15,                  // Расстояние между строками внутри абзаца
		//        LineRule : linerule_Auto,         // Тип расстрояния между строками
		//        Before   : 0,                     // Дополнительное расстояние до абзаца
		//        After    : 10 * g_dKoef_pt_to_mm  // Дополнительное расстояние после абзаца
		//    }
		//	}
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;

		// TODO: как только разъединят настройки параграфа и текста переделать тут
		var TextPr         = logicDocument.GetCalculatedTextPr();
		ParaPr.Subscript   = TextPr.VertAlign === AscCommon.vertalign_SubScript;
		ParaPr.Superscript = TextPr.VertAlign === AscCommon.vertalign_SuperScript;
		ParaPr.Strikeout   = TextPr.Strikeout;
		ParaPr.DStrikeout  = TextPr.DStrikeout;
		ParaPr.AllCaps     = TextPr.Caps;
		ParaPr.SmallCaps   = TextPr.SmallCaps;
		ParaPr.TextSpacing = TextPr.Spacing;
		ParaPr.Position    = TextPr.Position;
		ParaPr.Ligatures   = TextPr.Ligatures;
		//-----------------------------------------------------------------------------

		if (true === ParaPr.Spacing.AfterAutoSpacing)
			ParaPr.Spacing.After = AscCommonWord.spacing_Auto;
		else if (undefined === ParaPr.Spacing.AfterAutoSpacing)
			ParaPr.Spacing.After = AscCommonWord.UnknownValue;

		if (true === ParaPr.Spacing.BeforeAutoSpacing)
			ParaPr.Spacing.Before = AscCommonWord.spacing_Auto;
		else if (undefined === ParaPr.Spacing.BeforeAutoSpacing)
			ParaPr.Spacing.Before = AscCommonWord.UnknownValue;

		if (undefined !== ParaPr.FramePr && undefined !== ParaPr.FramePr.Wrap)
		{
			if (AscCommonWord.wrap_NotBeside === ParaPr.FramePr.Wrap)
				ParaPr.FramePr.Wrap = false;
			else if (AscCommonWord.wrap_Around === ParaPr.FramePr.Wrap)
				ParaPr.FramePr.Wrap = true;
			else
				ParaPr.FramePr.Wrap = undefined;
		}

		this.sync_ParaSpacingLine(ParaPr.Spacing);
		this.Update_ParaInd(ParaPr.Ind);
		this.sync_PrAlignCallBack(ParaPr.Jc);
		
		let bidi = ParaPr.Bidi;
		if (undefined === bidi)
		{
			let paragraph = logicDocument.GetCurrentParagraph(false, false, {FirstInSelection : true});
			bidi = paragraph ? paragraph.GetParagraphBidi() : undefined;
		}
		
		this.sendEvent("asc_onTextDirection", bidi);
		
		this.sync_PrPropCallback(ParaPr);
	};
	/*----------------------------------------------------------------*/
	/*functions for working with clipboard, document*/
	asc_docs_api.prototype._printDesktop = function (options)
	{
		let desktopOptions = {};
		if (options && options.advancedOptions)
			desktopOptions["nativeOptions"] = options.advancedOptions.asc_getNativeOptions();

		if (!this.isPdfEditor())
		{
			if (options && options.advancedOptions && (Asc.c_oAscPrintType.Selection === options.advancedOptions.asc_getPrintType()))
			{
				desktopOptions["printOptions"] = { "selection" : 1 };
			}
			desktopOptions["documentLayout"] = {"drawPlaceHolders":false,"drawFormHighlight":false,"isPrint":true};
			window["AscDesktopEditor"]["Print"](JSON.stringify(desktopOptions));
		}
		return true;
	};
	asc_docs_api.prototype.Undo           = function()
	{
		if (!this.canUndoRedoByRestrictions())
			return;
		
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		oLogicDocument.Document_Undo();
	};
	asc_docs_api.prototype.Redo           = function()
	{
		if (!this.canUndoRedoByRestrictions())
			return;
		
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		oLogicDocument.Document_Redo();
	};
	asc_docs_api.prototype.Copy           = function()
	{
		return AscCommon.g_clipboardBase.Button_Copy();
	};
	asc_docs_api.prototype.Update_ParaTab = function(Default_Tab, ParaTabs)
	{
		this.WordControl.m_oDrawingDocument.Update_ParaTab(Default_Tab, ParaTabs);
	};
	asc_docs_api.prototype.Cut            = function()
	{
		return AscCommon.g_clipboardBase.Button_Cut();
	};
	asc_docs_api.prototype.Paste          = function()
	{
		if (!this.WordControl.m_oLogicDocument)
			return false;

		if (AscCommon.g_clipboardBase.IsWorking())
			return false;

		return AscCommon.g_clipboardBase.Button_Paste();
	};

	asc_docs_api.prototype.Share = function()
	{

	};

	asc_docs_api.prototype.asc_CheckCopy = function(_clipboard /* CClipboardData */, _formats)
	{
		if (!this.WordControl.m_oLogicDocument)
		{
			var _text_object = (AscCommon.c_oAscClipboardDataFormat.Text & _formats) ? {Text : ""} : null;
			var _html_data   = this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.Copy(_text_object);

			//TEXT
			if (AscCommon.c_oAscClipboardDataFormat.Text & _formats)
			{
				_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Text, _text_object.Text);
			}
			//HTML
			if (AscCommon.c_oAscClipboardDataFormat.Html & _formats)
			{
				_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Html, _html_data);
			}
			return;
		}

		var sBase64 = null, _data;
		var checkData = function(data) {
			return 	"" === data ? null : data;
		};

		//TEXT
		if (AscCommon.c_oAscClipboardDataFormat.Text & _formats)
		{
			_data = this.WordControl.m_oLogicDocument.GetSelectedText(false);
			_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Text, checkData(_data))
		}
		//HTML
		if (AscCommon.c_oAscClipboardDataFormat.Html & _formats)
		{
			var oCopyProcessor = new AscCommon.CopyProcessor(this);
			sBase64            = oCopyProcessor.Start();
			_data              = oCopyProcessor.getInnerHtml();

			_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Html, checkData(_data))
		}
		//INTERNAL
		if (AscCommon.c_oAscClipboardDataFormat.Internal & _formats)
		{
			if (sBase64 === null)
			{
				if(window["NATIVE_EDITOR_ENJINE"])
				{
					var oCopyProcessor = new AscCommon.CopyProcessor(this, true);
					sBase64 = oCopyProcessor.getSelectedBinary();
				}
				else
				{
					var oCopyProcessor = new AscCommon.CopyProcessor(this);
					sBase64 = oCopyProcessor.Start();
				}
			}

			_data = sBase64;
			_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Internal, checkData(_data))
		}
	};

	asc_docs_api.prototype.asc_SelectionCut = function()
	{
	    if (AscCommon.CollaborativeEditing.Get_GlobalLock())
    	    return;

		var _logicDoc = this.WordControl.m_oLogicDocument;
		if (!_logicDoc || _logicDoc.IsSelectionEmpty(true))
			return;

		if (!_logicDoc.IsSelectionLocked(AscCommon.changestype_Remove, null, true, _logicDoc.IsFormFieldEditing()))
		{
			_logicDoc.StartAction(AscDFH.historydescription_Cut);
			_logicDoc.Remove(-1, true, true, false, false, true); // -1 - нормальное удаление  (например, для таблиц)
			_logicDoc.Recalculate();
			_logicDoc.UpdateSelection();
			_logicDoc.FinalizeAction();
		}
	};

	asc_docs_api.prototype.asc_PasteData = function(_format, data1, data2, text_data, useCurrentPoint, callback, checkLocks, rejectCallback)
	{
		if (AscCommon.CollaborativeEditing.Get_GlobalLock())
			return;

		var _logicDoc = this.WordControl.m_oLogicDocument;
		if (!_logicDoc)
			return;

		// TODO: isPasteImage заменить на проверку того, что вставляется просто картинка
		var isPasteImage = AscCommon.checkOnlyOneImage(data1);

		var isLocked = false;
		var oCC      = null;

		if (isPasteImage)
			oCC = _logicDoc.GetContentControl();

		if (false !== checkLocks)
		{
			if (oCC && oCC.IsPicture())
				isLocked = _logicDoc.IsSelectionLocked(AscCommon.changestype_Image_Properties, null, true, _logicDoc.IsFormFieldEditing());
			else
				isLocked = _logicDoc.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, _logicDoc.IsFormFieldEditing());
		}

		if (isLocked)
		{
			rejectCallback && rejectCallback();
		}
		else
		{
			window['AscCommon'].g_specialPasteHelper.Paste_Process_Start(arguments[5]);

			if (!useCurrentPoint)
				_logicDoc.StartAction(AscDFH.historydescription_Document_PasteHotKey);

			AscCommon.Editor_Paste_Exec(this, _format, data1, data2, text_data, undefined, callback, rejectCallback);
		}
	};

	asc_docs_api.prototype.GetContentFromHtml = function(html, callback)
	{
		AscCommon.GetContentFromHtml(this, html, callback);
	};

	asc_docs_api.prototype._finalizeAction = function()
	{
		var _logicDoc = this.WordControl.m_oLogicDocument;
		if (!_logicDoc){
			return;
		}
		_logicDoc.FinalizeAction();
	};
	
	asc_docs_api.prototype.asc_SpecialPaste = function(props) 
	{
		return AscCommon.g_specialPasteHelper.Special_Paste(props);
	};
	
	asc_docs_api.prototype.asc_SpecialPasteData = function(props) 
	{
		if (AscCommon.CollaborativeEditing.Get_GlobalLock())
	        return;

		var _logicDoc = this.WordControl.m_oLogicDocument;
		if (!_logicDoc)
			return;
		
		//TODO пересмотреть проверку лока и добавление новой точки(AscDFH.historydescription_Document_PasteHotKey)
		if (false === _logicDoc.Document_Is_SelectionLocked(changestype_Paragraph_Content, null, true, false))
		{
			window['AscCommon'].g_specialPasteHelper.Paste_Process_Start();
			window['AscCommon'].g_specialPasteHelper.Special_Paste_Start();
			
			//undo previous action
			this.WordControl.m_oLogicDocument.Document_Undo();
			
			_logicDoc.StartAction(AscDFH.historydescription_Document_PasteHotKey);
			AscCommon.Editor_Paste_Exec(this, null, null, null, null, props, function () {
				_logicDoc.FinalizeAction();
			});
		}
	};
	
	asc_docs_api.prototype.asc_specialPasteShowButton = function()
	{
		var specialPasteHelper = window['AscCommon'].g_specialPasteHelper;
		//при быстром совместном редактировании отключаем возможность специальной вставки
		if(!specialPasteHelper || specialPasteHelper.CheckFastCoEditing())
		{
			return;
		}

		var specialPasteShowOptions = specialPasteHelper.buttonInfo ? specialPasteHelper.buttonInfo : null;
		if(specialPasteShowOptions && null !== specialPasteHelper.showButtonIdParagraph)
		{
			var isUpdate = specialPasteShowOptions.cellCoord;
			var id       = specialPasteHelper.showButtonIdParagraph;
			var elem     = g_oTableId.Get_ById(id);

			var _X       = elem.X;
			var _Y       = elem.Y;
			var _PageNum = elem.GetCurrentPageAbsolute();

			if (elem.IsTable())
			{
				var nPagesCount = elem.GetPagesCount();
				if (nPagesCount > 0)
				{
					var oBounds = elem.GetPageBounds(nPagesCount - 1);
					_X          = oBounds.Right;
					_Y          = oBounds.Bottom;
					_PageNum    = elem.GetAbsolutePage(nPagesCount - 1);
				}
			}
			else
			{
				var oBounds = elem.GetSelectionBounds(true);
				if (oBounds && oBounds.End)
				{
					_X       = oBounds.End.X + oBounds.End.W;
					_Y       = oBounds.End.Y + oBounds.End.H;
					_PageNum = oBounds.End.Page;
				}
			}

			var oTransform = elem.Get_ParentTextTransform();
			if (oTransform)
			{
				var __X = oTransform.TransformPointX(_X, _Y);
				var __Y = oTransform.TransformPointY(_X, _Y);

				_X = __X;
				_Y = __Y;
			}

			specialPasteHelper.buttonInfo.fixPosition = {x : _X, y : _Y, pageNum : _PageNum};

			var _coord   = this.WordControl.m_oLogicDocument.DrawingDocument.ConvertCoordsToCursorWR(_X, _Y, _PageNum);
			var curCoord = new AscCommon.asc_CRect(_coord.X, _coord.Y, 0, 0);
			specialPasteShowOptions.asc_setCellCoord(curCoord);

			if (isUpdate)
			{
				specialPasteShowOptions.options = [];
				this.asc_UpdateSpecialPasteButton(specialPasteShowOptions);
			}
			else
			{
				this.asc_ShowSpecialPasteButton(specialPasteShowOptions);
			}
		}

		specialPasteHelper.showButtonIdParagraph = null;
		return true;
	};

	asc_docs_api.prototype.asc_ShowSpecialPasteButton = function(props)
	{
		this.sendEvent("asc_onShowSpecialPasteOptions", props);
	};
	
	asc_docs_api.prototype.asc_HideSpecialPasteButton = function() 
	{
		this.sendEvent("asc_onHideSpecialPasteOptions");
	};
	
	asc_docs_api.prototype.asc_UpdateSpecialPasteButton = function()
	{
		var pasteHepler = AscCommon.g_specialPasteHelper;
		if(!pasteHepler)
		{
			return;
		}

		var props;
		if(pasteHepler.buttonInfo && pasteHepler.buttonInfo.fixPosition)
		{
			props = pasteHepler.buttonInfo;

			var _Y = props.fixPosition.y;
			var _X = props.fixPosition.x;
			var _PageNum = props.fixPosition.pageNum;

			var _coord = this.WordControl.m_oLogicDocument.DrawingDocument.ConvertCoordsToCursorWR(_X, _Y, _PageNum);
			var curCoord = new AscCommon.asc_CRect( _coord.X, _coord.Y, 0, 0 );
			props.asc_setCellCoord(curCoord);
		}

		if(props)
		{
			this.sendEvent("asc_onShowSpecialPasteOptions", props);
		}
	};

    asc_docs_api.prototype.beginInlineDropTarget = function(e)
    {
    	if (this.WordControl.m_oLogicDocument && this.WordControl.m_oDrawingDocument)
		{
			this.WordControl.m_oDrawingDocument.StartTrackText();
            this.WordControl.StartUpdateOverlay();
			this.WordControl.onMouseMove(e);
            this.WordControl.EndUpdateOverlay();
		}
    };
    asc_docs_api.prototype.endInlineDropTarget = function(e)
    {
        if (this.WordControl.m_oLogicDocument && this.WordControl.m_oDrawingDocument)
        {
            this.WordControl.m_oDrawingDocument.EndTrackText(true);
        }
    };
	
	asc_docs_api.prototype._onSaveCallbackInner = function()
	{
		var t = this;
		AscCommon.History.CheckUnionLastPoints();
		if (c_oAscCollaborativeMarksShowType.LastChanges === this.CollaborativeMarksShowType)
		{
			AscCommon.CollaborativeEditing.Clear_CollaborativeMarks();
		}

		// Принимаем чужие изменения
		var HaveOtherChanges = AscCommon.CollaborativeEditing.Have_OtherChanges();
		AscCommon.CollaborativeEditing.Apply_Changes();

		this.CoAuthoringApi.onUnSaveLock = function()
		{
			t.CoAuthoringApi.onUnSaveLock = null;
			if (t.isForceSaveOnUserSave && t.IsUserSave) {
				t.forceSaveButtonContinue = t.forceSave();
			}
			if (t.forceSaveForm) {
				t.forceSaveForm();
			}

			// Выставляем, что документ не модифицирован
			t.CheckChangedDocument();
			t.canSave    = true;
			t.IsUserSave = false;
			if (!t.forceSaveButtonContinue) {
				t.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.Save);
			}

			// Обновляем состояние возможности сохранения документа
			t._onUpdateDocumentCanSave();

			if (undefined !== window["AscDesktopEditor"])
			{
				window["AscDesktopEditor"]["OnSave"]();
			}
			if(AscCommon.g_specialPasteHelper && !AscCommon.CollaborativeEditing.Is_SingleUser()){
				AscCommon.g_specialPasteHelper.SpecialPasteButton_Hide();
			}

			if (t.canUnlockDocument) {
				t._unlockDocument();
			}
		};

		let cursorInfo = null;
		if (true === AscCommon.CollaborativeEditing.Is_Fast())
			cursorInfo = AscCommon.History.Get_DocumentPositionBinary();
		
		function sendChanges()
		{
			// Пересылаем свои изменения
			AscCommon.CollaborativeEditing.Send_Changes(t.IsUserSave, {
				"UserId"      : t.CoAuthoringApi.getUserConnectionId(),
				"UserShortId" : t.DocInfo.get_UserId(),
				"CursorInfo"  : cursorInfo
			}, HaveOtherChanges, true);
		}

		if (this.forceSaveUndoRequest)
		{
			AscCommon.CollaborativeEditing.Set_GlobalLock(false);
			AscCommon.CollaborativeEditing.Undo();
			this.forceSaveUndoRequest = false;
		}
		else if (this.forceSaveSendFormRequest)
		{
			this.forceSaveSendFormRequest = false;
			this.setCurrentRoleFilled();
			sendChanges();
			this._sendForm();
		}
		else
		{
			sendChanges();
			
			if (this.forceSaveDisconnectRequest)
			{
				this.forceSaveDisconnectRequest = false;
				AscCommon.CollaborativeEditing.Set_GlobalLock(false);
				this.setViewModeDisconnect(true);
				this.asc_coAuthoringDisconnect();
				this.sendEvent("asc_onDisconnectEveryone");
			}
		}
	};
	asc_docs_api.prototype._autoSaveInner = function () {
		var _curTime = new Date();
		if (null === this.lastSaveTime) {
			this.lastSaveTime = _curTime;
		}


		if (AscCommon.CollaborativeEditing.Is_Fast() && !AscCommon.CollaborativeEditing.Is_SingleUser()) {
			this.WordControl.m_oLogicDocument.Continue_FastCollaborativeEditing();
		} else if (this.isLiveViewer()) {
			if (AscCommon.CollaborativeEditing.Have_OtherChanges()) {
				AscCommon.CollaborativeEditing.Apply_Changes();
			}
		} else {
			var _bIsWaitScheme = false;
			if (AscCommon.History.Points && AscCommon.History.Index >= 0 && AscCommon.History.Index < AscCommon.History.Points.length) {
				if ((_curTime - AscCommon.History.Points[AscCommon.History.Index].Time) < this.intervalWaitAutoSave) {
					_bIsWaitScheme = true;
				}
			}

			if (!_bIsWaitScheme) {
				var _interval = (AscCommon.CollaborativeEditing.m_nUseType <= 0) ? this.autoSaveGapSlow :
					this.autoSaveGapFast;

				if ((_curTime - this.lastSaveTime) > _interval) {
					if (AscCommon.History.Have_Changes(true) == true) {
						this.asc_Save(true);
					}
					this.lastSaveTime = _curTime;
				}
			}
		}
	};
	asc_docs_api.prototype._saveCheck = function() {
		return !this.isLongAction() &&
			!(this.WordControl.m_oLogicDocument && this.WordControl.m_oLogicDocument.IsViewModeInReview());
	};
	asc_docs_api.prototype._haveOtherChanges = function () {
		return AscCommon.CollaborativeEditing.Have_OtherChanges();
	};
	asc_docs_api.prototype.asc_DownloadOrigin = function(options)
	{
		if (!options) {
			options = new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.UNKNOWN);
		}
		let bIsDownloadEvent = options.isDownloadEvent;
		let changes = null;
		let isCloudLocal = false;
		if (this.isUseNativeViewer && this.isDocumentRenderer()) {
			isCloudLocal = this.isCloudSaveAsLocalToDrawingFormat(Asc.c_oAscAsyncAction.DownloadAs, Asc.c_oAscFileType.PDF);
			changes = this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.Save();
			if (!changes && isCloudLocal)
				window["AscDesktopEditor"]["emulateCloudPrinting"](false);
		}

		if (changes || this.watermarkDraw) {
			if (changes && isCloudLocal) {
				this.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.DownloadAs);
				this.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Save);

				window["AscDesktopEditor"]["localSaveToDrawingFormat2"](this.documentTitle, this.DocumentUrl || "", "",
					AscCommon.Base64.encode(changes), AscCommon.Base64.encode(this.DocumentRenderer.getFileNativeBinary()), this.currentPassword || "", Asc.c_oAscFileType.PDF);

				return true;
			}

			options.asc_setFileType(Asc.c_oAscFileType.PDF);
			options.pdfChanges = changes;
			this.asc_DownloadAs(options);
			return;
		} else if((this.documentIsWopi && (options.isSaveAs && options.wopiSaveAsPath)) || !(this.isUseNativeViewer && this.isDocumentRenderer())) {
			//todo add new server request wopi_putRelativeOrigin
			options.asc_setFileType(AscCommon.getFormatByExtention(this.documentFormat));
			this.asc_DownloadAs(options);
			return;
		}
		//скачивание оригинального pdf, djvu, xps
		var downloadType = bIsDownloadEvent ? DownloadType.Download : DownloadType.None;
		var rData        = {
			"c"     : 'pathurl',
			"title" : AscCommon.changeFileExtention(this.documentTitle, this.documentFormat, Asc.c_nMaxDownloadTitleLen),
			"data"  : 'origin.' + this.documentFormat
		};
		var t            = this;
		t.fCurCallback   = function(input)
		{
			if (null != input && "pathurl" == input["type"])
			{
				if ('ok' == input["status"])
				{
					var url = input["data"];
					if (url)
					{
						t.processSavedFile(url, downloadType, input["filetype"]);
					}
					else
					{
						t.handlers.trigger("asc_onError", c_oAscError.ID.Unknown, c_oAscError.Level.NoCritical);
					}
				}
				else
				{
					t.handlers.trigger("asc_onError", AscCommon.mapAscServerErrorToAscError(parseInt(input["data"])),
						c_oAscError.Level.NoCritical);
				}
			}
			else
			{
				t.handlers.trigger("asc_onError", c_oAscError.ID.Unknown, c_oAscError.Level.NoCritical);
			}
		};
		sendCommand(this, null, rData);
	};
	asc_docs_api.prototype.asc_DownloadAs     = function(options)
	{
		if (this.isLongAction()) {
			return;
		}
		this.downloadAs(c_oAscAsyncAction.DownloadAs, options);
	};
	asc_docs_api.prototype.asc_DownloadAsMailMerge         = function(typeFile, StartIndex, EndIndex, bIsDownload)
	{
		if (this.isLongAction()) {
			return false;
		}
		var oDocumentMailMerge = this.WordControl.m_oLogicDocument.Get_MailMergedDocument(StartIndex, EndIndex);
		if (null != oDocumentMailMerge)
		{
			var actionType = null;
			var options = new Asc.asc_CDownloadOptions(typeFile, true);
			options.oDocumentMailMerge = oDocumentMailMerge;
			options.errorDirect = c_oAscError.ID.MailMergeSaveFile;
			if (bIsDownload) {
				actionType = Asc.c_oAscAsyncAction.DownloadMerge;
				options.isDownloadEvent = false;
			}
			this.downloadAs(actionType, options);
		}
		return null != oDocumentMailMerge;
	};
	asc_docs_api.prototype.Resize             = function()
	{
		if (false === this.bInit_word_control)
			return;
		this.WordControl.OnResize(false);
	};
	asc_docs_api.prototype.AddURL             = function(url)
	{

	};
	asc_docs_api.prototype.Help               = function()
	{

	};
	/*
	 idOption идентификатор дополнительного параметра, c_oAscAdvancedOptionsID.TXT.
	 option - какие свойства применить, пока массив. для TXT объект asc_CTextOptions(codepage)
	 exp:	asc_setAdvancedOptions(c_oAscAdvancedOptionsID.TXT, new Asc.asc_CTextOptions(1200) );
	 */
	asc_docs_api.prototype.asc_setAdvancedOptions       = function(idOption, option)
	{
		// Проверяем тип состояния в данный момент
		if (this.advancedOptionsAction !== c_oAscAdvancedOptionsAction.Open) {
			return;
		}
        if (AscCommon.EncryptionWorker.asc_setAdvancedOptions(this, idOption, option))
            return;

		switch (idOption)
		{
			case c_oAscAdvancedOptionsID.TXT:
				var rData = {
					"id"            : this.documentId,
					"userid"        : this.documentUserId,
					"format"        : this.documentFormat,
					"c"             : "reopen",
					"title"         : this.documentTitle,
					"codepage"      : option.asc_getCodePage(),
					"nobase64"      : true
				};
				sendCommand(this, null, rData);
				break;
			case c_oAscAdvancedOptionsID.DRM:
				this.currentPassword = option.asc_getPassword();
				var v = {
					"id": this.documentId,
					"userid": this.documentUserId,
					"format": this.documentFormat,
					"c": "reopen",
					"title": this.documentTitle,
					"password": option.asc_getPassword(),
					"nobase64": true
				};
				sendCommand(this, null, v);
				break;
		}
	};
	asc_docs_api.prototype.SetFontRenderingMode         = function(mode)
	{
		if (!this.isLoadFullApi)
		{
			this.tmpFontRenderingMode = mode;
			return;
		}

		if (c_oAscFontRenderingModeType.noHinting === mode)
            AscCommon.g_fontManager.SetHintsProps(false, false);
		else if (c_oAscFontRenderingModeType.hinting === mode)
            AscCommon.g_fontManager.SetHintsProps(true, false);
		else if (c_oAscFontRenderingModeType.hintingAndSubpixeling === mode)
            AscCommon.g_fontManager.SetHintsProps(true, true);

		this.WordControl.m_oDrawingDocument.ClearCachePages();

		if (AscCommon.g_fontManager2 !== undefined && AscCommon.g_fontManager2 !== null)
            AscCommon.g_fontManager2.ClearFontsRasterCache();

		if (this.bInit_word_control)
			this.WordControl.OnScroll();
	};
	asc_docs_api.prototype.processSavedFile             = function(url, downloadType, filetype)
	{
		var t = this;
		if (this.mailMergeFileData)
		{
			this.mailMergeFileData = null;
			AscCommon.loadFileContent(url, function(httpRequest)
			{
				if (null === httpRequest)
				{
					t.sendEvent("asc_onError", c_oAscError.ID.MailMergeLoadFile, c_oAscError.Level.NoCritical);
					return;
				}
				try
				{
					t.asc_StartMailMergeByList(JSON.parse(httpRequest.responseText));
				} catch (e)
				{
					t.sendEvent("asc_onError", c_oAscError.ID.MailMergeLoadFile, c_oAscError.Level.NoCritical);
				}
			});
		}
		else if (this.insertDocumentUrlsData && this.insertDocumentUrlsData.convertCallback)
		{
			this.insertDocumentUrlsData.convertCallback(this, url);
		}
		else
		{
			AscCommon.baseEditorsApi.prototype.processSavedFile.call(this, url, downloadType, filetype);
		}
	};
	asc_docs_api.prototype.startGetDocInfo              = function()
	{
		/*
		 Возвращаем объект следующего вида:
		 {
		 PageCount: 12,
		 WordsCount: 2321,
		 ParagraphCount: 45,
		 SymbolsCount: 232345,
		 SymbolsWSCount: 34356
		 }
		 */
		this.sync_GetDocInfoStartCallback();

		if (this.WordControl.m_oLogicDocument)
			this.WordControl.m_oLogicDocument.Statistics_Start();
	};
	asc_docs_api.prototype.stopGetDocInfo               = function()
	{
		this.sync_GetDocInfoStopCallback();
		
		if (this.WordControl.m_oLogicDocument)
			this.WordControl.m_oLogicDocument.Statistics_Stop();
	};
	asc_docs_api.prototype.sync_DocInfoCallback         = function(obj)
	{
		this.sendEvent("asc_onDocInfo", new CDocInfoProp(obj));
	};
	asc_docs_api.prototype.sync_GetDocInfoStartCallback = function()
	{
		this.sendEvent("asc_onGetDocInfoStart");
	};
	asc_docs_api.prototype.sync_GetDocInfoStopCallback  = function()
	{
		this.sendEvent("asc_onGetDocInfoStop");
	};
	asc_docs_api.prototype.sync_GetDocInfoEndCallback   = function()
	{
		this.sendEvent("asc_onGetDocInfoEnd");
	};
	asc_docs_api.prototype.sync_CanUndoCallback         = function(canUndo)
	{
		if (!this.canUndoRedoByRestrictions())
			canUndo = false;
		
		this.sendEvent("asc_onCanUndo", canUndo);
	};
	asc_docs_api.prototype.sync_CanRedoCallback         = function(canRedo)
	{
		if (!this.canUndoRedoByRestrictions()
			|| (true === AscCommon.CollaborativeEditing.Is_Fast() && true !== AscCommon.CollaborativeEditing.Is_SingleUser()))
			canRedo = false;

		this.sendEvent("asc_onCanRedo", canRedo);
	};

	asc_docs_api.prototype.can_CopyCut = function()
	{
		return this.WordControl.m_oLogicDocument.Can_CopyCut();
	};

	asc_docs_api.prototype.setStartPointHistory = function()
	{
		this.noCreatePoint  = true;
		this.exucuteHistory = true;
		this.incrementCounterLongAction();
		//this.WordControl.m_oLogicDocument.TurnOff_InterfaceEvents();
	};
	asc_docs_api.prototype.setEndPointHistory   = function()
	{
		this.noCreatePoint     = false;
		this.exucuteHistoryEnd = true;
		this.decrementCounterLongAction();
		//this.WordControl.m_oLogicDocument.TurnOn_InterfaceEvents();
	};

	function CDocInfoProp(obj)
	{
		if (obj)
		{
			this.PageCount      = obj.PageCount;
			this.WordsCount     = obj.WordsCount;
			this.ParagraphCount = obj.ParagraphCount;
			this.SymbolsCount   = obj.SymbolsCount;
			this.SymbolsWSCount = obj.SymbolsWSCount;
		}
		else
		{
			this.PageCount      = -1;
			this.WordsCount     = -1;
			this.ParagraphCount = -1;
			this.SymbolsCount   = -1;
			this.SymbolsWSCount = -1;
		}
	}

	CDocInfoProp.prototype.get_PageCount      = function()
	{
		return this.PageCount;
	};
	CDocInfoProp.prototype.put_PageCount      = function(v)
	{
		this.PageCount = v;
	};
	CDocInfoProp.prototype.get_WordsCount     = function()
	{
		return this.WordsCount;
	};
	CDocInfoProp.prototype.put_WordsCount     = function(v)
	{
		this.WordsCount = v;
	};
	CDocInfoProp.prototype.get_ParagraphCount = function()
	{
		return this.ParagraphCount;
	};
	CDocInfoProp.prototype.put_ParagraphCount = function(v)
	{
		this.ParagraphCount = v;
	};
	CDocInfoProp.prototype.get_SymbolsCount   = function()
	{
		return this.SymbolsCount;
	};
	CDocInfoProp.prototype.put_SymbolsCount   = function(v)
	{
		this.SymbolsCount = v;
	};
	CDocInfoProp.prototype.get_SymbolsWSCount = function()
	{
		return this.SymbolsWSCount;
	};
	CDocInfoProp.prototype.put_SymbolsWSCount = function(v)
	{
		this.SymbolsWSCount = v;
	};

	/*callbacks*/
	/*asc_docs_api.prototype.sync_CursorLockCallBack = function(isLock){
	 this.sendEvent("asc_onCursorLock",isLock);
	 }*/
	asc_docs_api.prototype.sync_UndoCallBack       = function()
	{
		this.sendEvent("asc_onUndo");
	};
	asc_docs_api.prototype.sync_RedoCallBack       = function()
	{
		this.sendEvent("asc_onRedo");
	};
	asc_docs_api.prototype.sync_CopyCallBack       = function()
	{
		this.sendEvent("asc_onCopy");
	};
	asc_docs_api.prototype.sync_CutCallBack        = function()
	{
		this.sendEvent("asc_onCut");
	};
	asc_docs_api.prototype.sync_PasteCallBack      = function()
	{
		this.sendEvent("asc_onPaste");
	};
	asc_docs_api.prototype.sync_ShareCallBack      = function()
	{
		this.sendEvent("asc_onShare");
	};
	asc_docs_api.prototype.sync_SaveCallBack       = function()
	{
		this.sendEvent("asc_onSave");
	};
	asc_docs_api.prototype.sync_DownloadAsCallBack = function()
	{
		this.sendEvent("asc_onDownload");
	};

	asc_docs_api.prototype.sync_AddURLCallback  = function()
	{
		this.sendEvent("asc_onAddURL");
	};
	asc_docs_api.prototype.sync_ErrorCallback   = function(errorID, errorLevel)
	{
		this.sendEvent("asc_onError", errorID, errorLevel);
	};
	asc_docs_api.prototype.sync_HelpCallback    = function(url)
	{
		this.sendEvent("asc_onHelp", url);
	};
	asc_docs_api.prototype.sync_UpdateZoom      = function(zoom)
	{
		this.sendEvent("asc_onZoom", zoom);
	};
	asc_docs_api.prototype.ClearPropObjCallback = function(prop)
	{//колбэк предшествующий приходу свойств объекта, prop а всякий случай
		this.sendEvent("asc_onClearPropObj", prop);
	};

	// mobile version methods:
	asc_docs_api.prototype.asc_GetDefaultTableStyles = function()
	{
		if (!this.WordControl.m_oLogicDocument)
			return;

		this.WordControl.m_oDrawingDocument.StartTableStylesCheck();
		this.WordControl.m_oDrawingDocument.TableStylesCheckLook = new AscCommon.CTableLook(true, true, false, false, true, false);
		this.WordControl.m_oDrawingDocument.EndTableStylesCheck();
	};

	asc_docs_api.prototype.CollectHeaders                  = function()
	{
		this.sync_ReturnHeadersCallback([]);
	};
	asc_docs_api.prototype.GetActiveHeader                 = function()
	{

	};
	asc_docs_api.prototype.gotoHeader                      = function(page, X, Y)
	{
		this.goToPage(page);
	};
	asc_docs_api.prototype.sync_ChangeActiveHeaderCallback = function(position, header)
	{
		this.sendEvent("asc_onChangeActiveHeader", position, new Asc.CHeader(header));
	};
	asc_docs_api.prototype.sync_ReturnHeadersCallback      = function(headers)
	{
		var _headers = [];
		for (var i = 0; i < headers.length; i++)
		{
			_headers[i] = new Asc.CHeader(headers[i]);
		}

		this.sendEvent("asc_onReturnHeaders", _headers);
	};
	/*----------------------------------------------------------------*/
	/*functions for working with search*/
	/*
	 структура поиска, предварительно, выглядит так
	 {
	 text: "...<b>слово поиска</b>...",
	 pageNumber: 0, //содержит номер страницы, где находится искомая последовательность
	 X: 0,//координаты по OX начала последовательности на данной страницы
	 Y: 0//координаты по OY начала последовательности на данной страницы
	 }
	 */

	asc_docs_api.prototype.asc_searchEnabled = function(bIsEnabled)
	{
		if (null != this.WordControl.m_oDrawingDocument.m_oDocumentRenderer)
		{
			this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.SearchResults.IsSearch = bIsEnabled;
			this.WordControl.OnUpdateOverlay();
		}
	};
	asc_docs_api.prototype.asc_findText = function(oProps, isNext, callback)
	{
		var result = 0;
		var CurMatchIdx = 0;
		var isAsync = false;
		var oViewer;

		if (null != this.WordControl.m_oDrawingDocument.m_oDocumentRenderer)
		{
			this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.SearchResults.IsSearch = true;
			oProps.SetText(oProps.GetText().trim());
			oViewer = this.WordControl.m_oDrawingDocument.m_oDocumentRenderer;
			isAsync = (true === oViewer.findText(oProps.GetText(), oProps.IsMatchCase(), oProps.IsWholeWords(), isNext, this.sync_setSearchCurrent)) ? true : false;
			result = oViewer.SearchResults.Count;

			if (oViewer.SearchResults.CurrentPage === 0)
				CurMatchIdx = oViewer.SearchResults.Current;
			else
			{
				// чтобы узнать, под каким номером в списке текущее совпадение
				// нужно посчитать сколько совпадений было до текущего на текущей странице
				for (var nPage = 0; nPage <= oViewer.SearchResults.CurrentPage; nPage++)
				{
					for (var nMatch = 0; nMatch < oViewer.SearchResults.Pages[nPage].length; nMatch++)
					{
						if (nPage == oViewer.SearchResults.CurrentPage && nMatch == oViewer.SearchResults.Current)
							break;
						CurMatchIdx++;
					}
				}
			}

			oViewer.SearchResults.CurMatchIdx = CurMatchIdx;
			this.sync_setSearchCurrent(CurMatchIdx, result);
		}
		else
		{
			var SearchEngine = editor.WordControl.m_oLogicDocument.Search(oProps);

			var Id = this.WordControl.m_oLogicDocument.GetSearchElementId(isNext);

			if (null != Id)
				this.WordControl.m_oLogicDocument.SelectSearchElement(Id);

			result = SearchEngine.Count;
		}

		if (!isAsync && callback)
			callback(result);

		return result;
	};
	asc_docs_api.prototype.asc_endFindText = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (oLogicDocument && oLogicDocument.SearchEngine)
		{
			return oLogicDocument.ClearSearch();
		}
	};
	asc_docs_api.prototype.asc_replaceText = function(oProps, replaceWith, isReplaceAll)
	{
		if (this.asc_GetErrorForReplaceString(replaceWith))
			return false;

		if (null == this.WordControl.m_oLogicDocument)
			return false;

		this.WordControl.m_oLogicDocument.Search(oProps);

		if (true === isReplaceAll)
			this.WordControl.m_oLogicDocument.ReplaceSearchElement(replaceWith, true, -1);
		else
		{
			var CurId      = this.WordControl.m_oLogicDocument.SearchEngine.CurId;
			var bDirection = this.WordControl.m_oLogicDocument.SearchEngine.Direction;
			if (-1 != CurId)
				this.WordControl.m_oLogicDocument.ReplaceSearchElement(replaceWith, false, CurId);

			var Id = this.WordControl.m_oLogicDocument.GetSearchElementId(bDirection);

			if (null != Id)
			{
				this.WordControl.m_oLogicDocument.SelectSearchElement(Id);
				return true;
			}

			return false;
		}
	};
	asc_docs_api.prototype.asc_GetErrorForReplaceString = function(sString)
	{
		return (new AscCommonWord.CSearchPatternEngine()).GetErrorForReplaceString(sString);
	};

	asc_docs_api.prototype._selectSearchingResults = function(bShow)
	{
		this.WordControl.m_oLogicDocument.HighlightSearchResults(bShow);
	};
	asc_docs_api.prototype.asc_isSelectSearchingResults = function()
	{
		if (null != this.WordControl.m_oDrawingDocument.m_oDocumentRenderer)
		{
			return this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.SearchResults.Show;
		}
		return this.WordControl.m_oLogicDocument.IsHighlightSearchResults();
	};
	asc_docs_api.prototype.sync_ReplaceAllCallback = function(ReplaceCount, OverallCount)
	{
		this.sendEvent("asc_onReplaceAll", OverallCount, ReplaceCount);
	};
	asc_docs_api.prototype.sync_SearchEndCallback = function()
	{
		this.sendEvent("asc_onSearchEnd");
	};
	asc_docs_api.prototype.sync_setSearchCurrent = function(nCurrent, nOverallCount)
	{
		this.sendEvent("asc_onSetSearchCurrent", nCurrent, nOverallCount);
	};
	asc_docs_api.prototype.asc_StartTextAroundSearch = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (oLogicDocument && oLogicDocument.SearchEngine)
		{
			oLogicDocument.SearchEngine.StartTextAround();
		}
	};
	asc_docs_api.prototype.asc_SelectSearchElement = function(sId)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (oLogicDocument && oLogicDocument.SearchEngine)
		{
			oLogicDocument.SelectSearchElement(sId);
		}
	};
	asc_docs_api.prototype.sync_startTextAroundSearch = function()
	{
		this.sendEvent("asc_onStartTextAroundSearch");
	};
	asc_docs_api.prototype.sync_endTextAroundSearch = function()
	{
		this.sendEvent("asc_onEndTextAroundSearch");
	};
	asc_docs_api.prototype.sync_getTextAroundSearchPack = function(arrElements)
	{
		this.sendEvent("asc_onGetTextAroundSearchPack", arrElements);
	};
	asc_docs_api.prototype.sync_removeTextAroundSearch = function(sId)
	{
		this.sendEvent("asc_onRemoveTextAroundSearch", [sId]);
	};
	/*----------------------------------------------------------------*/
	/*functions for working with font*/
	/*setters*/
	asc_docs_api.prototype.put_TextPrFontName = function(name)
	{
		var loader   = AscCommon.g_font_loader;
		var fontinfo = g_fontApplication.GetFontInfo(name);
		var isasync  = loader.LoadFont(fontinfo);
		if (false === isasync)
		{
			if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
			{
				this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextFontName);
				this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({
					FontFamily : {
						Name  : name,
						Index : -1
					}
				}));

				this.WordControl.m_oLogicDocument.Recalculate();
				this.WordControl.m_oLogicDocument.UpdateInterface();
				this.WordControl.m_oLogicDocument.FinalizeAction();
			}
		}
	};
	asc_docs_api.prototype.put_TextPrFontSize = function(size)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextFontSize);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({FontSize : Math.min(size, 300)}));
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.put_TextPrBold       = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextBold);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Bold : value}));
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_TextPrItalic     = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextItalic);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Italic : value}));
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_TextPrUnderline  = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextUnderline);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Underline : value}));
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_TextPrStrikeout  = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextStrikeout);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({
				Strikeout  : value,
				DStrikeout : false
			}));
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_TextPrDStrikeout = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextDStrikeout);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({
				DStrikeout : value,
				Strikeout  : false
			}));
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_TextPrSpacing    = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextSpacing);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Spacing : value}));
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.put_TextPrCaps = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextCaps);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({
				Caps      : value,
				SmallCaps : false
			}));
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.put_TextPrSmallCaps = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextSmallCaps);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({
				SmallCaps : value,
				Caps      : false
			}));
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};


	asc_docs_api.prototype.put_TextPrPosition = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextPosition);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Position : value}));
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.put_TextPrLang = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextLang);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Lang : {Val : value}}));
			this.WordControl.m_oLogicDocument.Spelling.CheckCurrentParagraph();
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	
	
	asc_docs_api.prototype.asc_setRtlTextDirection = function(isRtl)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		if (logicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties))
			return;
		
		logicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphBidi);
		logicDocument.SetParagraphBidi(isRtl);
		logicDocument.FinalizeAction();
	};
	asc_docs_api.prototype.asc_isRtlTextDirection = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;
		
		let paragraph = logicDocument.GetCurrentParagraph(false, false, {FirstInSelection : true});
		return paragraph ? paragraph.GetParagraphBidi() : false;
	};
	asc_docs_api.prototype._addRemoveSpaceBeforeAfterParagraph = function(event)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		let paragraphs = logicDocument.GetSelectedParagraphs();
		if (!paragraphs.length)
			return;
		
		let p = paragraphs[0];
		let finalPr = p.Get_CompiledPr2().ParaPr;
		if (!event.check(finalPr))
			return;
		
		if (logicDocument.IsSelectionLocked(AscCommon.changestype_None, {
			Type : AscCommon.changestype_2_ElementsArray_and_Type,
			Elements : paragraphs,
			CheckType : AscCommon.changestype_Paragraph_Properties
		}))
		{
			return;
		}
		
		let stylePr = p.getCompiledPrFromStyle().ParaPr;
		let newVal = event.getNewValue(stylePr);
		
		logicDocument.StartAction(AscDFH.historydescription_Document_AddRemoveBeforeAfterParagraph);
		for (let i = 0; i < paragraphs.length; ++i)
		{
			event.applyValue(paragraphs[i], newVal);
		}
		
		logicDocument.Recalculate();
		logicDocument.UpdateInterface();
		logicDocument.FinalizeAction()
	};
	asc_docs_api.prototype.asc_addSpaceBeforeParagraph = function()
	{
		this._addRemoveSpaceBeforeAfterParagraph({
			check : function(finalPr)
			{
				return finalPr.Spacing.Before < 0.001;
			},
			
			getNewValue : function(stylePr)
			{
				let val = 12 * AscCommonWord.g_dKoef_pt_to_mm;
				if (stylePr.Spacing.Before > 0.001)
					val = stylePr.Spacing.Before;
				
				return val;
			},
			
			applyValue : function(p, val)
			{
				let stylePr = p.getCompiledPrFromStyle().ParaPr;
				if (Math.abs(stylePr.Spacing.Before - val) > 0.001)
				{
					p.Set_Spacing({Before : val}, false);
				}
				else
				{
					let sp = p.Pr.Spacing.Copy();
					sp.Before = undefined;
					p.Set_Spacing(sp, true);
				}
			}
		});
	};
	asc_docs_api.prototype.asc_addSpaceAfterParagraph = function()
	{
		this._addRemoveSpaceBeforeAfterParagraph({
			check : function(finalPr)
			{
				return finalPr.Spacing.After < 0.001;
			},
			
			getNewValue : function(stylePr)
			{
				let val = 12 * AscCommonWord.g_dKoef_pt_to_mm;
				if (stylePr.Spacing.After > 0.001)
					val = stylePr.Spacing.After;
				
				return val;
			},
			
			applyValue : function(p, val) {
				let stylePr = p.getCompiledPrFromStyle().ParaPr;
				if (Math.abs(stylePr.Spacing.After - val) > 0.001)
				{
					p.Set_Spacing({After : val}, false);
				}
				else
				{
					let sp = p.Pr.Spacing.Copy();
					sp.After = undefined;
					p.Set_Spacing(sp, true);
				}
			}
		});
	};
	asc_docs_api.prototype.asc_removeSpaceBeforeParagraph = function()
	{
		this._addRemoveSpaceBeforeAfterParagraph({
			check : function(finalPr)
			{
				return finalPr.Spacing.Before > 0.001;
			},
			
			getNewValue : function(stylePr)
			{
				return 0;
			},
			
			applyValue : function(p, val)
			{
				let stylePr = p.getCompiledPrFromStyle().ParaPr;
				if (Math.abs(stylePr.Spacing.Before - val) > 0.001)
				{
					p.Set_Spacing({Before : val}, false);
				}
				else
				{
					let sp = p.Pr.Spacing.Copy();
					sp.Before = undefined;
					p.Set_Spacing(sp, true);
				}
			}
		});
	};
	asc_docs_api.prototype.asc_removeSpaceAfterParagraph = function()
	{
		this._addRemoveSpaceBeforeAfterParagraph({
			check : function(finalPr)
			{
				return finalPr.Spacing.After > 0.001;
			},
			
			getNewValue : function(stylePr)
			{
				return 0;
			},
			
			applyValue : function(p, val)
			{
				let stylePr = p.getCompiledPrFromStyle().ParaPr;
				if (Math.abs(stylePr.Spacing.After - val) > 0.001)
				{
					p.Set_Spacing({After : val}, false);
				}
				else
				{
					let sp = p.Pr.Spacing.Copy();
					sp.After = undefined;
					p.Set_Spacing(sp, true);
				}
			}
		});
	};
	asc_docs_api.prototype.asc_haveSpaceBeforeParagraph = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;
		
		let paragraphs = logicDocument.GetSelectedParagraphs();
		if (!paragraphs.length)
			return false;
		
		let p = paragraphs[0];
		let paraPr = p.Get_CompiledPr2().ParaPr;
		return paraPr.Spacing.Before > 0.001;
	};
	asc_docs_api.prototype.asc_haveSpaceAfterParagraph = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;
		
		let paragraphs = logicDocument.GetSelectedParagraphs();
		if (!paragraphs.length)
			return false;
		
		let p = paragraphs[0];
		let paraPr = p.Get_CompiledPr2().ParaPr;
		return paraPr.Spacing.After > 0.001;
	};
	
	asc_docs_api.prototype.put_PrLineSpacing          = function(Type, Value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphLineSpacing);
			this.WordControl.m_oLogicDocument.SetParagraphSpacing({LineRule : Type, Line : Value});
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_LineSpacingBeforeAfter = function(type, value)//"type == 0" means "Before", "type == 1" means "After"
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphLineSpacingBeforeAfter);
			switch (type)
			{
				case 0:
				{
					if (AscCommonWord.spacing_Auto === value)
						this.WordControl.m_oLogicDocument.SetParagraphSpacing({BeforeAutoSpacing : true});
					else
						this.WordControl.m_oLogicDocument.SetParagraphSpacing({
							Before            : value,
							BeforeAutoSpacing : false
						});

					break;
				}
				case 1:
				{
					if (AscCommonWord.spacing_Auto === value)
						this.WordControl.m_oLogicDocument.SetParagraphSpacing({AfterAutoSpacing : true});
					else
						this.WordControl.m_oLogicDocument.SetParagraphSpacing({
							After            : value,
							AfterAutoSpacing : false
						});

					break;
				}
			}
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.FontSizeIn                 = function()
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_IncFontSize);
			this.WordControl.m_oLogicDocument.IncreaseDecreaseFontSize(true);
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.FontSizeOut                = function()
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_DecFontSize);
			this.WordControl.m_oLogicDocument.IncreaseDecreaseFontSize(false);
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	// Object:
	// {
	//    Bottom :
	//    {
	//        Color : { r : 0, g : 0, b : 0 },
	//        Value : border_Single,
	//        Size  : 0.5 * g_dKoef_pt_to_mm
	//        Space : 0
	//    },
	//    Left :
	//    {
	//        ....
	//    }
	//    Right :
	//    {
	//        ....
	//    }
	//    Top :
	//    {
	//        ....
	//    }
	//    },
	//    Between :
	//    {
	//        ....
	//    }
	// }


	asc_docs_api.prototype.put_Borders = function(Obj)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphBorders);
			this.WordControl.m_oLogicDocument.SetParagraphBorders(Obj);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	/*callbacks*/
	asc_docs_api.prototype.sync_BoldCallBack             = function(isBold)
	{
		this.sendEvent("asc_onBold", isBold);
	};
	asc_docs_api.prototype.sync_ItalicCallBack           = function(isItalic)
	{
		this.sendEvent("asc_onItalic", isItalic);
	};
	asc_docs_api.prototype.sync_UnderlineCallBack        = function(isUnderline)
	{
		this.sendEvent("asc_onUnderline", isUnderline);
	};
	asc_docs_api.prototype.sync_StrikeoutCallBack        = function(isStrikeout)
	{
		this.sendEvent("asc_onStrikeout", isStrikeout);
	};
	asc_docs_api.prototype.sync_TextPrFontFamilyCallBack = function(FontFamily)
	{
		if (undefined != FontFamily)
			this.sendEvent("asc_onFontFamily", new asc_CTextFontFamily(FontFamily));
		else
			this.sendEvent("asc_onFontFamily", new asc_CTextFontFamily({Name : "", Index : -1}));
	};
	asc_docs_api.prototype.sync_TextPrFontSizeCallBack   = function(FontSize)
	{
		this.sendEvent("asc_onFontSize", FontSize);
	};
	asc_docs_api.prototype.sync_PrLineSpacingCallBack    = function(LineSpacing)
	{
		this.sendEvent("asc_onLineSpacing", new Asc.asc_CParagraphInd(LineSpacing));
	};
	asc_docs_api.prototype.sync_InitEditorStyles         = function(styles_painter)
	{
		if (!this.isViewMode) {
			this.sendEvent("asc_onInitEditorStyles", styles_painter);
		}
	};
	asc_docs_api.prototype.sync_InitEditorTableStyles    = function()
	{
		if (!this.isViewMode) {
			this.sendEvent("asc_onInitTableTemplates");
		}
	};
	asc_docs_api.prototype.asc_getTableStylesPreviews    = function(bUseDefault, arrIds)
	{
		let oDrawingDocument = this.WordControl ? this.WordControl.m_oDrawingDocument : null;
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oDrawingDocument || !oLogicDocument)
			return [];

		this.private_CheckTableStylesPreviewGenerator();

		let arrPreviews = [];

		if (arrIds && arrIds.length)
		{
			let oStylesList = oLogicDocument.GetStyles();
			for (let nIndex = 0, nCount = arrIds.length; nIndex < nCount; ++nIndex)
			{
				let oStyle   = oStylesList.Get(arrIds[nIndex]);
				let oPreview = this.TableStylesPreviewGenerator.GetPreview(oStyle);
				if (oPreview)
					arrPreviews.push(oPreview);
			}
		}
		else
		{
			arrPreviews = this.TableStylesPreviewGenerator.GetAllPreviews(bUseDefault);
		}

		return arrPreviews;
	};
	asc_docs_api.prototype.asc_generateTableStylesPreviews = function(bUseDefault)
	{
		this.private_CheckTableStylesPreviewGenerator();
		this.TableStylesPreviewGenerator.Begin(bUseDefault);
	};
	asc_docs_api.prototype.private_CheckTableStylesPreviewGenerator = function()
	{
		if (!this.TableStylesPreviewGenerator)
			this.TableStylesPreviewGenerator = new AscCommon.CTableStylesPreviewGenerator(this.private_GetLogicDocument())
	};


	/*----------------------------------------------------------------*/
	/*functions for working with paragraph*/
	/*setters*/
	// Right = 0; Left = 1; Center = 2; Justify = 3; or using enum that written above

	/* структура для параграфа
	 Ind :
	 {
	 Left      : 0,                    // Левый отступ
	 Right     : 0,                    // Правый отступ
	 FirstLine : 0                     // Первая строка
	 }
	 Spacing :
	 {
	 Line     : 1.15,                  // Расстояние между строками внутри абзаца
	 LineRule : linerule_Auto,         // Тип расстрояния между строками
	 Before   : 0,                     // Дополнительное расстояние до абзаца
	 After    : 10 * g_dKoef_pt_to_mm  // Дополнительное расстояние после абзаца
	 },
	 KeepLines : false,                    // переносить параграф на новую страницу,
	 // если на текущей он целиком не убирается
	 PageBreakBefore : false
	 */

	asc_docs_api.prototype.paraApply = function(Props)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var arrAdditional = [];
		if (undefined != Props.DefaultTab)
		{
			arrAdditional.push({
				Type      : AscCommon.changestype_2_Element_and_Type,
				Element   : oLogicDocument,
				CheckType : AscCommon.changestype_Document_SectPr
			});
		}

		if (undefined !== Props.Subscript
			|| undefined !== Props.Strikeout
			|| undefined !== Props.DStrikeout
			|| undefined !== Props.SmallCaps
			|| undefined !== Props.AllCaps
			|| undefined !== Props.TextSpacing
			|| undefined !== Props.Position)
		{
			arrAdditional.push({
				Type  : AscCommon.changestype_2_AdditionalTypes,
				Types : [AscCommon.changestype_Paragraph_TextProperties]
			});
		}

		if (!oLogicDocument.IsSelectionLocked(changestype_Paragraph_Properties, arrAdditional))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphPr);

			if ("undefined" != typeof(Props.ContextualSpacing) && null != Props.ContextualSpacing)
				this.WordControl.m_oLogicDocument.SetParagraphContextualSpacing(Props.ContextualSpacing);

			if ("undefined" != typeof(Props.Ind) && null != Props.Ind)
				this.WordControl.m_oLogicDocument.SetParagraphIndent(Props.Ind);

			if ("undefined" != typeof(Props.Jc) && null != Props.Jc)
				this.WordControl.m_oLogicDocument.SetParagraphAlign(Props.Jc);

			if ("undefined" != typeof(Props.KeepLines) && null != Props.KeepLines)
				this.WordControl.m_oLogicDocument.SetParagraphKeepLines(Props.KeepLines);

			if (undefined != Props.KeepNext && null != Props.KeepNext)
				this.WordControl.m_oLogicDocument.SetParagraphKeepNext(Props.KeepNext);

			if (undefined != Props.WidowControl && null != Props.WidowControl)
				this.WordControl.m_oLogicDocument.SetParagraphWidowControl(Props.WidowControl);

			if ("undefined" != typeof(Props.PageBreakBefore) && null != Props.PageBreakBefore)
				this.WordControl.m_oLogicDocument.SetParagraphPageBreakBefore(Props.PageBreakBefore);

			if ("undefined" != typeof(Props.Spacing) && null != Props.Spacing)
				this.WordControl.m_oLogicDocument.SetParagraphSpacing(Props.Spacing);

			if (undefined !== Props.OutlineLvl)
				this.WordControl.m_oLogicDocument.SetParagraphOutlineLvl(Props.OutlineLvl);

			if (Props.Shd)
			{
				var Unifill        = new AscFormat.CUniFill();
				Unifill.fill       = new AscFormat.CSolidFill();
				Unifill.fill.color = AscFormat.CorrectUniColor(Props.Shd.Color, Unifill.fill.color, 1);
				
				let shd = {
					Value   : Asc.c_oAscShd.Clear,
					Color   : {
						r    : Props.Shd.Color.asc_getR(),
						g    : Props.Shd.Color.asc_getG(),
						b    : Props.Shd.Color.asc_getB(),
						Auto : false
					},
					Fill    : {
						r    : Props.Shd.Color.asc_getR(),
						g    : Props.Shd.Color.asc_getG(),
						b    : Props.Shd.Color.asc_getB(),
						Auto : false
					},
					Unifill : Unifill
				};

				if (Asc.c_oAscShd.Nil === Props.Shd.Value)
				{
					shd.Color.Auto = true;
					shd.Fill.Auto  = true;
					shd.Unifill    = undefined;
					shd.ThemeFill  = undefined;
				}
				
				this.WordControl.m_oLogicDocument.SetParagraphShd(shd);
			}

			if ("undefined" != typeof(Props.Brd) && null != Props.Brd)
			{
				if (Props.Brd.Left && Props.Brd.Left.Color)
				{
					Props.Brd.Left.Unifill = AscFormat.CreateUnifillFromAscColor(Props.Brd.Left.Color, 1);
				}
				if (Props.Brd.Top && Props.Brd.Top.Color)
				{
					Props.Brd.Top.Unifill = AscFormat.CreateUnifillFromAscColor(Props.Brd.Top.Color, 1);
				}
				if (Props.Brd.Right && Props.Brd.Right.Color)
				{
					Props.Brd.Right.Unifill = AscFormat.CreateUnifillFromAscColor(Props.Brd.Right.Color, 1);
				}
				if (Props.Brd.Bottom && Props.Brd.Bottom.Color)
				{
					Props.Brd.Bottom.Unifill = AscFormat.CreateUnifillFromAscColor(Props.Brd.Bottom.Color, 1);
				}
				if (Props.Brd.InsideH && Props.Brd.InsideH.Color)
				{
					Props.Brd.InsideH.Unifill = AscFormat.CreateUnifillFromAscColor(Props.Brd.InsideH.Color, 1);
				}
				if (Props.Brd.InsideV && Props.Brd.InsideV.Color)
				{
					Props.Brd.InsideV.Unifill = AscFormat.CreateUnifillFromAscColor(Props.Brd.InsideV.Color, 1);
				}

				this.WordControl.m_oLogicDocument.SetParagraphBorders(Props.Brd);
			}

			if (undefined != Props.Tabs)
			{
				var Tabs = new AscCommonWord.CParaTabs();
				Tabs.Set_FromObject(Props.Tabs.Tabs);
				this.WordControl.m_oLogicDocument.SetParagraphTabs(Tabs);
			}

			if (undefined != Props.DefaultTab)
			{
				this.WordControl.m_oLogicDocument.Set_DocumentDefaultTab(Props.DefaultTab);
			}

			if (undefined !== Props.SuppressLineNumbers)
				oLogicDocument.SetParagraphSuppressLineNumbers(Props.SuppressLineNumbers);
			
			if (undefined !== Props.Bidi)
				oLogicDocument.SetParagraphBidi(Props.Bidi);

			// TODO: как только разъединят настройки параграфа и текста переделать тут
			var TextPr = new AscCommonWord.CTextPr();

			if (true === Props.Subscript)
				TextPr.VertAlign = AscCommon.vertalign_SubScript;
			else if (true === Props.Superscript)
				TextPr.VertAlign = AscCommon.vertalign_SuperScript;
			else if (false === Props.Superscript || false === Props.Subscript)
				TextPr.VertAlign = AscCommon.vertalign_Baseline;

			if (undefined != Props.Strikeout)
			{
				TextPr.Strikeout  = Props.Strikeout;
				TextPr.DStrikeout = false;
			}

			if (undefined != Props.DStrikeout)
			{
				TextPr.DStrikeout = Props.DStrikeout;
				if (true === TextPr.DStrikeout)
					TextPr.Strikeout = false;
			}

			if (undefined != Props.SmallCaps)
			{
				TextPr.SmallCaps = Props.SmallCaps;
				TextPr.AllCaps   = false;
			}

			if (undefined != Props.AllCaps)
			{
				TextPr.Caps = Props.AllCaps;
				if (true === TextPr.AllCaps)
					TextPr.SmallCaps = false;
			}

			if (undefined != Props.TextSpacing)
				TextPr.Spacing = Props.TextSpacing;

			if (undefined != Props.Position)
				TextPr.Position = Props.Position;

			if (undefined !== Props.Ligatures)
				TextPr.Ligatures = Props.Ligatures;
			
			oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr(TextPr));
			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.asc_IsShowListIndentsSettings = function ()
	{
		const oLogicDocument = this.private_GetLogicDocument();

		if (!oLogicDocument)
			return false;

		const oNumPr = oLogicDocument.GetSelectedNum(true);
		if (!oNumPr)
			return false;

		const oNum = oLogicDocument.GetNumbering().GetNum(oNumPr.NumId);
		if (!oNum)
			return false;

		const oAbstractNum = oNum.GetAbstractNum();
		if (oAbstractNum)
		{
			const nMultiLvlType = oAbstractNum.GetMultiLvlType();
			return nMultiLvlType !== Asc.c_oAbstractNumMultiLvlTypes.MultiLevel;
		}
		return false;
	};
	
	asc_docs_api.prototype.put_PrAlign = function(value)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (logicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties))
			return;
		
		logicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphAlign);
		logicDocument.SetParagraphAlign(value);
		logicDocument.UpdateInterface();
		logicDocument.Recalculate();
		logicDocument.FinalizeAction();
	};
	// 0- baseline, 2-subscript, 1-superscript
	asc_docs_api.prototype.put_TextPrBaseline = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextVertAlign);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({VertAlign : value}));
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_ListType = function(type, subtype)
	{
		let numObject = AscWord.GetNumberingObjectByDeprecatedTypes(type, subtype);
		if (!numObject)
			return;

		this.put_ListTypeCustom(numObject);
	};
	asc_docs_api.prototype.put_ListTypeCustom = function(value)
	{
		let numInfo = AscWord.CNumInfo.Parse(value);
		if (!numInfo)
			return;

		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;

		return new Promise(function(resolve)
		{
			let symbols = AscWord.GetNumberingSymbols(numInfo);
			if (symbols && symbols.length)
				AscFonts.FontPickerByCharacter.checkText(symbols, this, resolve);
			else
				resolve();
		}).then(function()
		{
			let oRes = null;
			if (!logicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties))
			{
				logicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphNumbering);
				let numPr = logicDocument.SetParagraphNumbering(numInfo);
				logicDocument.FinalizeAction();
				
				let num = numPr ? logicDocument.GetNumbering().GetNum(numPr.NumId) : null;
				if (!num)
					oRes = null;
				else if ((!numInfo.HaveLvl() ||numInfo.IsSingleLevel()) && undefined !== numPr.Lvl)
					oRes = AscWord.CNumInfo.FromNum(num, numPr.Lvl).ToJson();
				else
					oRes = AscWord.CNumInfo.FromNum(num, null, logicDocument.GetStyles()).ToJson();
			}
			return oRes;
		});
	};
	asc_docs_api.prototype.asc_CompareNumberingPresets = function(value1, value2)
	{
		let numInfo1 = AscWord.CNumInfo.Parse(value1);
		let numInfo2 = AscWord.CNumInfo.Parse(value2);
		if (!numInfo1 || !numInfo2)
			return false;
		
		return numInfo1.IsEqual(numInfo2);
	};
	asc_docs_api.prototype.asc_IsCurrentNumberingPreset = function(value, singleLevel)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;
		
		let numInfo = AscWord.CNumInfo.Parse(value);
		if (!numInfo)
			return false;
		
		let numPr;
		if (singleLevel)
		{
			// MSWord показывает нумерацию у первого параграфа в выделении (т.е. кнопка с нумерация может не быть зажата
			// вообще, но при этом в списке нумераций выбрана нумерация первого параграфа)
			let paragraph = logicDocument.GetCurrentParagraph(false, false, {FirstInSelection : true});
			numPr = paragraph ? paragraph.GetNumPr() : null;
		}
		else
		{
			numPr = logicDocument.GetSelectedNum(true);
		}
		
		if (!numPr || !numPr.IsValid())
			return numInfo.IsNone();
		
		let num = logicDocument.GetNumbering().GetNum(numPr.NumId);
		return numInfo.CompareWithNum(num, singleLevel ? numPr.Lvl : undefined);
	};
	asc_docs_api.prototype.asc_GetCurrentNumberingJson = function(singleLevel)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return null;
		
		let numPr     = logicDocument.GetSelectedNum();
		let numbering = logicDocument.GetNumbering();
		let num       = numPr && numPr.IsValid() ? numbering.GetNum(numPr.NumId) : null;
		if (!num)
			return null;

		return AscWord.CNumInfo.FromNum(num, singleLevel ? 0 : null, logicDocument.GetStyles()).ToJson();
	};
	asc_docs_api.prototype.asc_ContinueNumbering = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		oLogicDocument.ContinueNumbering();
	};
	asc_docs_api.prototype.asc_RestartNumbering = function(nRestartValue)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		oLogicDocument.RestartNumbering(nRestartValue);
	};
	asc_docs_api.prototype.asc_GetCurrentNumberingId = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return null;

		var oNumPr = oLogicDocument.GetSelectedNum(true);
		if (!oNumPr)
			return null;

		var oNum = oLogicDocument.GetNumbering().GetNum(oNumPr.NumId);
		if (!oNum)
			return null;

		return oNumPr.NumId;
	};
	asc_docs_api.prototype.asc_GetCurrentNumberingLvl = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return -1;

		var oNumPr = oLogicDocument.GetSelectedNum(true);
		if (!oNumPr || undefined === oNumPr.Lvl || null === oNumPr.Lvl)
			return -1;

		return oNumPr.Lvl;
	};
	asc_docs_api.prototype.asc_GetCalculatedNumberingValue = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return -1;

		var oParagraph = oLogicDocument.GetCurrentParagraph(true);
		if (!oParagraph)
			return -1;

		return oParagraph.GetNumberingCalculatedValue();
	};
	asc_docs_api.prototype.asc_GetNumberingPr = function(sNumId)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return null;

		var oNum = oLogicDocument.GetNumbering().GetNum(sNumId);
		if (!oNum)
			return null;

		var oAscNumbering = new Asc.CAscNumbering();
		oNum.FillToAscNum(oAscNumbering);
		return oAscNumbering;
	};
	asc_docs_api.prototype.asc_AddNewNumbering = function(oAscNumbering, isApply)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return null;

		var arrParagraphs;
		if (isApply)
		{
			arrParagraphs = oLogicDocument.GetSelectedParagraphs();

			if (oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
					Type      : AscCommon.changestype_2_ElementsArray_and_Type,
					Elements  : arrParagraphs,
					CheckType : AscCommon.changestype_Paragraph_Properties
				}))
			{
				return null;
			}
		}

		oLogicDocument.StartAction(AscDFH.historydescription_Document_CreateNum);
		var oNumbering = oLogicDocument.GetNumbering();
		var oNum = oNumbering.CreateNum();
		oNum.FillFromAscNum(oAscNumbering);
		var sNumId = oNum.GetId();

		if (arrParagraphs)
		{
			for (var nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
			{
				var oPara = arrParagraphs[nIndex];
				var oNumPr = oPara.GetNumPr();
				if (!oNumPr)
					oPara.ApplyNumPr(sNumId, 0);
				else
					oPara.ApplyNumPr(sNumId, oNumPr.Lvl);
			}

			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.UpdateSelection();
		}

		oLogicDocument.FinalizeAction();
	};
	asc_docs_api.prototype.asc_ChangeNumberingLvl = function(sNumId, oAscNumberingLvl, nLvl)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		var oNum = oLogicDocument.GetNumbering().GetNum(sNumId);
		if (!oNum)
			return;

		if (!oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
				Type      : AscCommon.changestype_2_ElementsArray_and_Type,
				Elements  : [oNum],
				CheckType : AscCommon.changestype_Paragraph_Properties
			}))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_ChangeNumLvl);

			if (oAscNumberingLvl instanceof Asc.CAscNumberingLvl)
			{
				oNum.SetAscLvl(oAscNumberingLvl, nLvl);
			}
			else if (undefined !== oAscNumberingLvl.length && oAscNumberingLvl.length === nLvl.length)
			{
				for (var nIndex = 0, nCount = oAscNumberingLvl.length; nIndex < nCount; ++nIndex)
				{
					oNum.SetAscLvl(oAscNumberingLvl[nIndex], nLvl[nIndex]);
				}
			}

			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.UpdateSelection();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_SetNumberingLvl = function(nLvl)
	{
		var _nLvl = Math.max(-1, Math.min(8, nLvl));

		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var arrSelectedParagraphs = oLogicDocument.GetSelectedParagraphs();
		if (arrSelectedParagraphs.length <= 0)
			return;

		if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_None, {
			Type : AscCommon.changestype_2_ElementsArray_and_Type,
			Elements : arrSelectedParagraphs,
			CheckType : AscCommon.changestype_Paragraph_Properties
		}))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_SetNumberingLvl);

			for (var nIndex = 0, nCount = arrSelectedParagraphs.length; nIndex < nCount; ++nIndex)
			{
				var oParagraph = arrSelectedParagraphs[nIndex];
				var oNumPr = oParagraph.GetNumPr();
				if (!oNumPr)
					continue;

				if (-1 === _nLvl)
					oParagraph.RemoveNumPr();
				else
					oParagraph.SetNumPr(oNumPr.NumId, _nLvl);
			}

			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.UpdateSelection();
			oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.put_Style    = function(sName)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphStyle);
			this.WordControl.m_oLogicDocument.SetParagraphStyle(sName, true);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_SetParagraphSuppressLineNumbers = function(isSuppress)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphSuppressLineNumbers);
			oLogicDocument.SetParagraphSuppressLineNumbers(isSuppress);
			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.SetDeviceInputHelperId = function(idKeyboard)
	{
		if (window.ID_KEYBOARD_AREA === undefined && this.WordControl.m_oMainView != null)
		{
			window.ID_KEYBOARD_AREA = document.getElementById(idKeyboard);

			window.ID_KEYBOARD_AREA.onkeypress = function(e)
			{
				if (false === editor.WordControl.IsFocus)
				{
					editor.WordControl.IsFocus = true;
					var ret                    = editor.WordControl.onKeyPress(e);
					editor.WordControl.IsFocus = false;
					return ret;
				}
			}
			window.ID_KEYBOARD_AREA.onkeydown  = function(e)
			{
				if (false === editor.WordControl.IsFocus)
				{
					editor.WordControl.IsFocus = true;
					var ret                    = editor.WordControl.onKeyDown(e);
					editor.WordControl.IsFocus = false;
					return ret;
				}
			}
		}
	};

	asc_docs_api.prototype.put_ShowSnapLines = function(isShow)
	{
		this.ShowSmartGuides = isShow;
	};
	asc_docs_api.prototype.get_ShowSnapLines = function()
	{
		return this.ShowSmartGuides;
	};

	asc_docs_api.prototype.put_ShowSmartGuides = function(isShow)
	{
		this.ShowSmartGuides = isShow;
	};
	asc_docs_api.prototype.get_ShowSmartGuides = function()
	{
		return this.ShowSmartGuides;
	};

	asc_docs_api.prototype.put_ShowParaMarks      = function(isShow)
	{
		/*
		 if (window.IsAddDiv === undefined && this.WordControl.m_oMainView != null)
		 {
		 window.IsAddDiv = true;

		 var _div = this.WordControl.m_oMainView.HtmlElement;

		 var test = document.createElement('textarea');
		 test.id = "area_id";

		 test.setAttribute("style", "font-family:arial;font-size:12pt;position:absolute;resize:none;padding:2px;margin:0px;font-weight:normal;box-sizing:content-box;-moz-box-sizing:content-box;-webkit-box-sizing:content-box;z-index:1000");
		 test.style.border = "2px solid #4363A4";

		 test.style.width = "100px";
		 //this.TextBoxInput.style.height = "40px";
		 test.rows = 1;

		 _div.appendChild(test);

		 test.onkeypress = function(e){
		 return editor.WordControl.onKeyPress(e);
		 }
		 test.onkeydown = function(e){
		 return editor.WordControl.onKeyDown(e);
		 }
		 }
		 */

		if (this.isRestrictionForms() || this.isRestrictionComments())
			isShow = false;

		this.ShowParaMarks = isShow;
		this.WordControl.OnRePaintAttack();

		if (true === this.isMarkerFormat)
			this.sync_MarkerFormatCallback(false);

		return this.ShowParaMarks;
	};
	asc_docs_api.prototype.get_ShowParaMarks      = function()
	{
		return this.ShowParaMarks;
	};
	asc_docs_api.prototype.sync_ShowParaMarks     = function()
	{
		this.sendEvent("asc_onShowParaMarks", this.get_ShowParaMarks());
	};
	asc_docs_api.prototype.put_ShowTableEmptyLine = function(isShow)
	{
		this.isShowTableEmptyLine = isShow;
		this.WordControl.OnRePaintAttack();

		if (true === this.isMarkerFormat)
			this.sync_MarkerFormatCallback(false);

		return this.isShowTableEmptyLine;
	};
	asc_docs_api.prototype.get_ShowTableEmptyLine = function()
	{
		return this.isShowTableEmptyLine;
	};
	asc_docs_api.prototype.put_PageBreak          = function(isBreak)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			this.isPageBreakBefore = isBreak;
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphPageBreakBefore);
			this.WordControl.m_oLogicDocument.SetParagraphPageBreakBefore(isBreak);
			this.sync_PageBreakCallback(isBreak);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.put_WidowControl = function(bValue)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphWidowControl);
			this.WordControl.m_oLogicDocument.SetParagraphWidowControl(bValue);
			this.sync_WidowControlCallback(bValue);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.put_KeepLines = function(isKeepLines)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			this.isKeepLinesTogether = isKeepLines;
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphKeepLines);
			this.WordControl.m_oLogicDocument.SetParagraphKeepLines(isKeepLines);
			this.sync_KeepLinesCallback(isKeepLines);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.put_KeepNext = function(isKeepNext)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphKeepNext);
			this.WordControl.m_oLogicDocument.SetParagraphKeepNext(isKeepNext);
			this.sync_KeepNextCallback(isKeepNext);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.put_AddSpaceBetweenPrg = function(isSpacePrg)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			this.isAddSpaceBetweenPrg = isSpacePrg;
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphContextualSpacing);
			this.WordControl.m_oLogicDocument.SetParagraphContextualSpacing(isSpacePrg);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_LineHighLight      = function(is_flag, r, g, b)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			if (false === is_flag)
			{
				this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextHighlightNone);
				this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({HighLight : AscCommonWord.highlight_None}));
				this.WordControl.m_oLogicDocument.FinalizeAction();
			}
			else
			{
				this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextHighlightColor);
				this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({
					HighLight : {
						r : r,
						g : g,
						b : b
					}
				}));
				this.WordControl.m_oLogicDocument.FinalizeAction();
			}
		}
	};
	asc_docs_api.prototype.put_TextColor          = function(color)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextColor);

			if (true === color.Auto)
			{
				oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({
					Color      : {
						Auto : true,
						r    : 0,
						g    : 0,
						b    : 0
					}, Unifill : undefined
				}));
			}
			else
			{
				if(color && color.asc_getType && color.asc_getType() === Asc.c_oAscColor.COLOR_TYPE_SRGB)
				{
					var oColor = new AscCommonWord.CDocumentColor(color.r, color.g, color.b, false);
					oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Color: oColor, Unifill: undefined}));
				}
				else
				{
					var Unifill        = new AscFormat.CUniFill();
					Unifill.fill       = new AscFormat.CSolidFill();
					Unifill.fill.color = AscFormat.CorrectUniColor(color, Unifill.fill.color, 1);
					oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Unifill : Unifill}));
				}
			}

			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_ParagraphShade     = function(is_flag, color, isOnlyPara)
	{
		if (!this.WordControl.m_oLogicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties, {
			Type  : AscCommon.changestype_2_AdditionalTypes,
			Types : [AscCommon.changestype_Paragraph_TextProperties]
		}))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphShd);

			if (true === isOnlyPara)
				this.WordControl.m_oLogicDocument.SetUseTextShd(false);

			if (false === is_flag)
			{
				this.WordControl.m_oLogicDocument.SetParagraphShd({
					Value : Asc.c_oAscShd.Clear,
					Color : {
						r    : 0,
						g    : 0,
						b    : 0,
						Auto : true
					},
					Fill  : {
						r    : 255,
						g    : 255,
						b    : 255,
						Auto : true
					}
				});
			}
			else
			{
				var Unifill        = new AscFormat.CUniFill();
				Unifill.fill       = new AscFormat.CSolidFill();
				Unifill.fill.color = AscFormat.CorrectUniColor(color, Unifill.fill.color, 1);

				this.WordControl.m_oLogicDocument.SetParagraphShd({
					Value     : Asc.c_oAscShd.Clear,
					Color     : {
						r    : color.asc_getR(),
						g    : color.asc_getG(),
						b    : color.asc_getB(),
						Auto : false
					},
					Fill      : {
						r    : color.asc_getR(),
						g    : color.asc_getG(),
						b    : color.asc_getB(),
						Auto : false
					},
					Unifill   : Unifill,
					ThemeFill : Unifill
				});
			}

			this.WordControl.m_oLogicDocument.SetUseTextShd(true);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_PrIndent           = function(value, levelValue)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphIndent);
			this.WordControl.m_oLogicDocument.SetParagraphIndent({Left : value, ChangeLevel : levelValue});
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_ParagraphOutlineLvl = function(nLvl)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		if (false === oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphOutlineLvl);
			oLogicDocument.SetParagraphOutlineLvl(nLvl);
			oLogicDocument.FinalizeAction();
			oLogicDocument.UpdateDocumentOutlinePosition();
		}
	};

	// signatures
	asc_docs_api.prototype.asc_addSignatureLine = function (oPr, Width, Height, sImgUrl) {
		var oLogicDocument = this.WordControl.m_oLogicDocument;
        if (false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Document_Content_Add))
        {
			oLogicDocument.StartAction(AscDFH.historydescription_Document_InsertSignatureLine);
			var oSpToEdit = null;
			var sGuid = oPr.asc_getGuid();
			if(sGuid)
			{
				var oDrawingObjects = oLogicDocument.DrawingObjects;
				var ret = [], allSpr = [];
				allSpr = allSpr.concat(allSpr.concat(oDrawingObjects.getAllSignatures2(ret, oDrawingObjects.getDrawingArray())));
				for(var i = 0; i < allSpr.length; ++i)
				{
					if(allSpr[i].getSignatureLineGuid() === sGuid)
					{
						oSpToEdit = allSpr[i];
						break;
					}
				}
			}
			if(!oSpToEdit)
			{
				var oSignature = AscFormat.fCreateSignatureShape(oPr, true, null, Width, Height, sImgUrl);
				var Drawing   = new AscCommonWord.ParaDrawing(oSignature.spPr.xfrm.extX, oSignature.spPr.xfrm.extY, null, this.WordControl.m_oDrawingDocument, null, null);
				oSignature.setParent(Drawing);
				Drawing.Set_GraphicObject(oSignature);
				oLogicDocument.AddSignatureLine(Drawing);
				oSignature.setSignature(oSignature.signatureLine);
				this.sendEvent("asc_onAddSignature", oSignature.signatureLine.id);
			}
			else
			{
				oSpToEdit.setSignaturePr(oPr, sImgUrl);
				oLogicDocument.Recalculate();
				this.sendEvent("asc_onAddSignature", sGuid);
			}
			oLogicDocument.UpdateInterface();
			oLogicDocument.FinalizeAction();
        }
    };

    asc_docs_api.prototype.asc_getAllSignatures = function(){
    	if (!this.WordControl.m_oLogicDocument)
    		return [];
    	return this.WordControl.m_oLogicDocument.GetAllSignatures();
	};


    asc_docs_api.prototype.asc_CallSignatureDblClickEvent = function(sGuid){
        return this.WordControl.m_oLogicDocument.CallSignatureDblClickEvent(sGuid);
    };
    asc_docs_api.prototype.gotoSignatureInternal = function(sGuid){
        return this.WordControl.m_oLogicDocument.GoToSignature(sGuid);
    };
    //////////////////////////////////////////////////////////////////////////

	asc_docs_api.prototype.IncreaseIndent         = function()
	{
		this.WordControl.m_oLogicDocument.IncreaseIndent();
	};
	asc_docs_api.prototype.DecreaseIndent         = function()
	{
		this.WordControl.m_oLogicDocument.DecreaseIndent();
	};
	asc_docs_api.prototype.put_PrIndentRight      = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphIndentRight);
			this.WordControl.m_oLogicDocument.SetParagraphIndent({Right : value});
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_PrFirstLineIndent  = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphIndentFirstLine);
			this.WordControl.m_oLogicDocument.SetParagraphIndent({FirstLine : value});
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_Margins            = function(left, top, right, bottom)
	{
		this.WordControl.m_oLogicDocument.SetDocumentMargin({Left : left, Top : top, Right : right, Bottom : bottom});
	};
	asc_docs_api.prototype.getFocusObject         = function()
	{//возвратит тип элемента - параграф c_oAscTypeSelectElement.Paragraph, изображение c_oAscTypeSelectElement.Image, таблица c_oAscTypeSelectElement.Table, колонтитул c_oAscTypeSelectElement.Header.

	};

	/*callbacks*/
	asc_docs_api.prototype.sync_VerticalAlign     = function(typeBaseline)
	{
		this.sendEvent("asc_onVerticalAlign", typeBaseline);
	};
	asc_docs_api.prototype.sync_PrAlignCallBack   = function(value)
	{
		this.sendEvent("asc_onPrAlign", value);
	};
	asc_docs_api.prototype.sync_TextColor         = function(TextPr)
	{
		if (TextPr.Unifill && TextPr.Unifill.fill && TextPr.Unifill.fill.type === c_oAscFill.FILL_TYPE_SOLID && TextPr.Unifill.fill.color)
		{
			this.sendEvent("asc_onTextColor", AscCommon.CreateAscColor(TextPr.Unifill.fill.color));
		}
		else if (undefined != TextPr.Color)
		{
			this.sendEvent("asc_onTextColor", AscCommon.CreateAscColorCustom(TextPr.Color.r, TextPr.Color.g, TextPr.Color.b, TextPr.Color.Auto));
		}
	};
	asc_docs_api.prototype.sync_TextHighLight     = function(HighLight)
	{
		if (undefined != HighLight)
			this.sendEvent("asc_onTextHighLight", new AscCommon.CColor(HighLight.r, HighLight.g, HighLight.b));
	};
	asc_docs_api.prototype.sync_TextSpacing       = function(Spacing)
	{
		this.sendEvent("asc_onTextSpacing", Spacing);
	};
	asc_docs_api.prototype.sync_TextDStrikeout    = function(Value)
	{
		this.sendEvent("asc_onTextDStrikeout", Value);
	};
	asc_docs_api.prototype.sync_TextCaps          = function(Value)
	{
		this.sendEvent("asc_onTextCaps", Value);
	};
	asc_docs_api.prototype.sync_TextSmallCaps     = function(Value)
	{
		this.sendEvent("asc_onTextSmallCaps", Value);
	};
	asc_docs_api.prototype.sync_TextPosition      = function(Value)
	{
		this.sendEvent("asc_onTextPosition", Value);
	};
	asc_docs_api.prototype.sync_TextLangCallBack  = function(Lang)
	{
		// TODO: По-хорошему, надо сюда уже присылать lcid, а решать какой выше
		let lcid = Lang.Val;
		if (this.asc_isRtlTextDirection() && undefined !== Lang.Bidi)
			lcid = Lang.Bidi;
		
		this.sendEvent("asc_onTextLanguage", lcid);
	};
	asc_docs_api.prototype.sync_ParaStyleName     = function(Name)
	{
		this.sendEvent("asc_onParaStyleName", Name);
	};
	asc_docs_api.prototype.sync_ParaSpacingLine   = function(SpacingLine)
	{
		if (true === SpacingLine.AfterAutoSpacing)
			SpacingLine.After = AscCommonWord.spacing_Auto;
		else if (undefined === SpacingLine.AfterAutoSpacing)
			SpacingLine.After = AscCommonWord.UnknownValue;

		if (true === SpacingLine.BeforeAutoSpacing)
			SpacingLine.Before = AscCommonWord.spacing_Auto;
		else if (undefined === SpacingLine.BeforeAutoSpacing)
			SpacingLine.Before = AscCommonWord.UnknownValue;

		this.sendEvent("asc_onParaSpacingLine", new AscCommon.asc_CParagraphSpacing(SpacingLine));
	};
	asc_docs_api.prototype.sync_PageBreakCallback = function(isBreak)
	{
		this.sendEvent("asc_onPageBreak", isBreak);
	};

	asc_docs_api.prototype.sync_WidowControlCallback = function(bValue)
	{
		this.sendEvent("asc_onWidowControl", bValue);
	};

	asc_docs_api.prototype.sync_KeepNextCallback = function(bValue)
	{
		this.sendEvent("asc_onKeepNext", bValue);
	};

	asc_docs_api.prototype.sync_KeepLinesCallback       = function(isKeepLines)
	{
		this.sendEvent("asc_onKeepLines", isKeepLines);
	};
	asc_docs_api.prototype.sync_ShowParaMarksCallback   = function()
	{
		this.sendEvent("asc_onShowParaMarks");
	};
	asc_docs_api.prototype.sync_SpaceBetweenPrgCallback = function()
	{
		this.sendEvent("asc_onSpaceBetweenPrg");
	};
	asc_docs_api.prototype.sync_RangePermPropCallback = function(pr)
	{
		this.sync_ChangeLastSelectedElement(Asc.c_oAscTypeSelectElement.UnProtectedRegion, pr);
	};
	asc_docs_api.prototype.sync_PrPropCallback          = function(prProp)
	{
		var _len = this.SelectedObjectsStack.length;
		if (_len > 0)
		{
			if (this.SelectedObjectsStack[_len - 1].Type == c_oAscTypeSelectElement.Paragraph)
			{
				this.SelectedObjectsStack[_len - 1].Value = new Asc.asc_CParagraphProperty(prProp);
				return;
			}
		}

		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.Paragraph, new Asc.asc_CParagraphProperty(prProp));
	};

	asc_docs_api.prototype.sync_MathPropCallback = function(MathProp)
	{
		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.Math, MathProp);
	};

	asc_docs_api.prototype.sync_EndAddShape = function()
	{
		editor.sendEvent("asc_onEndAddShape");
		if (this.WordControl.m_oDrawingDocument.m_sLockedCursorType == "crosshair")
		{
			this.WordControl.m_oDrawingDocument.UnlockCursorType();
		}
	};

	asc_docs_api.prototype.SetDrawingFreeze = function(bIsFreeze)
	{
		if (!this.WordControl)
			return;

		this.WordControl.DrawingFreeze = bIsFreeze;

		var _elem1 = document.getElementById("id_main");
		if (_elem1)
		{
			var _elem2 = document.getElementById("id_horscrollpanel");
			var _elem3 = document.getElementById("id_panel_right");
			if (bIsFreeze)
			{
				_elem1.style.display = "none";
				if (_elem2) _elem2.style.display = "none";
				if (_elem3) _elem3.style.display = "none";
			}
			else
			{
				_elem1.style.display = "block";
				if (_elem2) _elem2.style.display = "block";
				if (_elem3) _elem3.style.display = "block";
			}
		}

		if (!bIsFreeze)
			this.WordControl.OnScroll();
	};

	//----------------------------------------------------------------------------------------------------------------------
	// Работаем с формулами
	//----------------------------------------------------------------------------------------------------------------------
	asc_docs_api.prototype.asc_SetMathProps = function(MathProps)
	{
		this.WordControl.m_oLogicDocument.Set_MathProps(MathProps);
	};

	asc_docs_api.prototype["asc_SetMathProps"] = asc_docs_api.prototype.asc_SetMathProps;
	//----------------------------------------------------------------------------------------------------------------------
	// Работаем с настройками секции
	//----------------------------------------------------------------------------------------------------------------------
	asc_docs_api.prototype.change_PageOrient       = function(isPortrait)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Document_SectPr))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetPageOrientation);
			if (isPortrait)
			{
				this.WordControl.m_oLogicDocument.Set_DocumentOrientation(Asc.c_oAscPageOrientation.PagePortrait);
				this.DocumentOrientation = isPortrait;
			}
			else
			{
				this.WordControl.m_oLogicDocument.Set_DocumentOrientation(Asc.c_oAscPageOrientation.PageLandscape);
				this.DocumentOrientation = isPortrait;
			}
			this.sync_PageOrientCallback(editor.get_DocumentOrientation());
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.get_DocumentOrientation = function()
	{
		return this.DocumentOrientation;
	};
	asc_docs_api.prototype.change_DocSize          = function(width, height)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Document_SectPr))
		{
			if (this.isMobileVersion && this.WordControl.MobileTouchManager)
				this.WordControl.MobileTouchManager.BeginZoomCheck();

			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetPageSize);
			if (this.DocumentOrientation)
				this.WordControl.m_oLogicDocument.Set_DocumentPageSize(width, height);
			else
				this.WordControl.m_oLogicDocument.Set_DocumentPageSize(height, width);

			this.WordControl.m_oLogicDocument.FinalizeAction();

			if (this.isMobileVersion && this.WordControl.MobileTouchManager)
				this.WordControl.MobileTouchManager.EndZoomCheck();
		}
	};

	asc_docs_api.prototype.get_DocumentWidth = function()
	{
		return AscCommon.Page_Width;
	};

	asc_docs_api.prototype.get_DocumentHeight = function()
	{
		return AscCommon.Page_Height;
	};

	asc_docs_api.prototype.asc_SetSectionProps       = function(Props)
	{
		this.WordControl.m_oLogicDocument.Set_SectionProps(Props);
	};
	asc_docs_api.prototype.asc_GetSectionProps       = function()
	{
		return this.WordControl.m_oLogicDocument.Get_SectionProps();
	};
	asc_docs_api.prototype.asc_GetCurrentColumnWidth = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return 0;

		return oLogicDocument.GetCurrentColumnWidth();
	};
	asc_docs_api.prototype.sync_SectionPropsCallback = function(Props)
	{
		this.sendEvent("asc_onSectionProps", Props);
	};
	asc_docs_api.prototype["asc_SetSectionProps"]       = asc_docs_api.prototype.asc_SetSectionProps;
	asc_docs_api.prototype["asc_GetSectionProps"]       = asc_docs_api.prototype.asc_GetSectionProps;
	asc_docs_api.prototype["asc_GetCurrentColumnWidth"] = asc_docs_api.prototype.asc_GetCurrentColumnWidth;

	asc_docs_api.prototype.asc_SetColumnsProps       = function(ColumnsProps)
	{
		this.WordControl.m_oLogicDocument.Set_ColumnsProps(ColumnsProps);
	};
	asc_docs_api.prototype.asc_GetColumnsProps       = function()
	{
		return this.WordControl.m_oLogicDocument.Get_ColumnsProps();
	};
	asc_docs_api.prototype["asc_SetColumnsProps"]    = asc_docs_api.prototype.asc_SetColumnsProps;
	asc_docs_api.prototype["asc_GetColumnsProps"]    = asc_docs_api.prototype.asc_GetColumnsProps;

	asc_docs_api.prototype.asc_SetLineNumbersProps = function(nApplyType, oLineNumbersProps)
	{
		this.WordControl.m_oLogicDocument.SetLineNumbersProps(nApplyType, oLineNumbersProps);
	};
	asc_docs_api.prototype.asc_GetLineNumbersProps = function()
	{
		return this.WordControl.m_oLogicDocument.GetLineNumbersProps();
	};
	asc_docs_api.prototype["asc_SetLineNumbersProps"] = asc_docs_api.prototype.asc_SetLineNumbersProps;
	asc_docs_api.prototype["asc_GetLineNumbersProps"] = asc_docs_api.prototype.asc_GetLineNumbersProps;

	asc_docs_api.prototype.asc_GetWatermarkProps = function()
	{
		return this.WordControl.m_oLogicDocument.GetWatermarkProps();
	};

	asc_docs_api.prototype.asc_SetWatermarkProps = function(oProps)
	{
		let oCurrentProps = this.asc_GetWatermarkProps();
		if(oCurrentProps.isEqual(oProps))
		{
			return;
		}
		let oTextPr = oProps.get_TextPr();
		let oApi = this;
		if(oTextPr)
		{
			let oFontFamily = oTextPr.get_FontFamily();
			if(oFontFamily && typeof oFontFamily.get_Name() === "string")
			{
				if(!g_fontApplication)
				{
					return;
				}
				let oLoader     = AscCommon.g_font_loader;
				let oFontInfo   = g_fontApplication.GetFontInfo(oFontFamily.get_Name());
				oFontFamily.put_Name(oFontInfo.Name);
				let bAsync    = oLoader.LoadFont(oFontInfo, function () {
					this.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadFont);
					oApi.asc_SetWatermarkProps(oProps);
				}, null);
				if(bAsync)
				{
					return;
				}
			}
		}
		let sText = oProps.get_Text();
		if(!(typeof sText === "string")) {
			sText = "";
		}
		AscFonts.FontPickerByCharacter.checkText(sText, this, function () {
			return oApi.WordControl.m_oLogicDocument.SetWatermarkProps(oProps);
		});
	};
	asc_docs_api.prototype.asc_WatermarkRemove = function(oProps)
	{
		var oProps = new Asc.CAscWatermarkProperties();
		oProps.put_Type(Asc.c_oAscWatermarkType.None);
		return this.WordControl.m_oLogicDocument.SetWatermarkProps(oProps);
	};

	asc_docs_api.prototype.sync_ColumnsPropsCallback = function(ColumnsProps)
	{
		this.sendEvent("asc_onColumnsProps", ColumnsProps);
	};
	asc_docs_api.prototype.sync_LineNumbersPropsCollback = function(oLineNumbersProps)
	{
		this.sendEvent("asc_onLineNumbersProps", oLineNumbersProps);
	};
	asc_docs_api.prototype.asc_SetFootnoteProps = function(oFootnotePr, bApplyToAll)
	{
		this.WordControl.m_oLogicDocument.SetFootnotePr(oFootnotePr, bApplyToAll);
	};
	asc_docs_api.prototype.asc_GetFootnoteProps = function()
	{
		return this.WordControl.m_oLogicDocument.GetFootnotePr();
	};
	asc_docs_api.prototype.asc_AddFootnote = function(sText)
	{
		return this.WordControl.m_oLogicDocument.AddFootnote(sText);
	};
	asc_docs_api.prototype.asc_RemoveAllFootnotes = function(bRemoveFootnotes, bRemoveEndnotes)
	{
		this.WordControl.m_oLogicDocument.RemoveAllFootnotes(bRemoveFootnotes, bRemoveEndnotes);
	};
	asc_docs_api.prototype.asc_GotoFootnote = function(isNext)
	{
		this.WordControl.m_oLogicDocument.GotoFootnote(isNext);
	};
	asc_docs_api.prototype.asc_IsCursorInFootnote = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (oLogicDocument && true === oLogicDocument.IsCursorInFootnote())
			return true;

		return false;
	};
	asc_docs_api.prototype.asc_AddEndnote = function(sText)
	{
		this.WordControl.m_oLogicDocument.AddEndnote(sText);
	};
	asc_docs_api.prototype.asc_GotoEndnote = function(isNext)
	{
		this.WordControl.m_oLogicDocument.GotoEndnote(isNext);
	};
	asc_docs_api.prototype.asc_IsCursorInEndnote = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		return (oLogicDocument ? oLogicDocument.IsCursorInEndnote() : false);
	};
	asc_docs_api.prototype.asc_SetEndnoteProps = function(oEndnotePr, bApplyToAll)
	{
		this.WordControl.m_oLogicDocument.SetEndnotePr(oEndnotePr, bApplyToAll);
	};
	asc_docs_api.prototype.asc_GetEndnoteProps = function()
	{
		return this.WordControl.m_oLogicDocument.GetEndnotePr();
	};
	asc_docs_api.prototype.asc_ConvertFootnoteType = function(isCurrent, isFootnotes, isEndnotes)
	{
		this.WordControl.m_oLogicDocument.ConvertFootnoteType(isCurrent, isFootnotes, isEndnotes);
	};
	asc_docs_api.prototype["asc_AddFootnote"]         = asc_docs_api.prototype.asc_AddFootnote;
	asc_docs_api.prototype["asc_RemoveAllFootnotes"]  = asc_docs_api.prototype.asc_RemoveAllFootnotes;
	asc_docs_api.prototype["asc_GetFootnoteProps"]    = asc_docs_api.prototype.asc_GetFootnoteProps;
	asc_docs_api.prototype["asc_SetFootnoteProps"]    = asc_docs_api.prototype.asc_SetFootnoteProps;
	asc_docs_api.prototype["asc_GotoFootnote"]        = asc_docs_api.prototype.asc_GotoFootnote;
	asc_docs_api.prototype["asc_IsCursorInFootnote"]  = asc_docs_api.prototype.asc_IsCursorInFootnote;
	asc_docs_api.prototype["asc_AddEndnote"]          = asc_docs_api.prototype.asc_AddEndnote;
	asc_docs_api.prototype["asc_GetEndnoteProps"]     = asc_docs_api.prototype.asc_GetEndnoteProps;
	asc_docs_api.prototype["asc_SetEndnoteProps"]     = asc_docs_api.prototype.asc_SetEndnoteProps;
	asc_docs_api.prototype["asc_GotoEndnote"]         = asc_docs_api.prototype.asc_GotoEndnote;
	asc_docs_api.prototype["asc_IsCursorInEndnote"]   = asc_docs_api.prototype.asc_IsCursorInEndnote;
	asc_docs_api.prototype["asc_ConvertFootnoteType"] = asc_docs_api.prototype.asc_ConvertFootnoteType;


	asc_docs_api.prototype.put_AddPageBreak              = function()
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
		{
			var Document = this.WordControl.m_oLogicDocument;

			if (null === Document.IsCursorInHyperlink(false))
			{
				this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddPageBreak);
				this.WordControl.m_oLogicDocument.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Page));
				this.WordControl.m_oLogicDocument.FinalizeAction();
			}
		}
	};
	asc_docs_api.prototype.put_AddColumnBreak            = function()
	{
		var Document = this.WordControl.m_oLogicDocument;
		if (false === Document.Document_Is_SelectionLocked(changestype_Paragraph_Content))
		{
			if (null === Document.IsCursorInHyperlink(false))
			{
				this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddPageBreak);
				this.WordControl.m_oLogicDocument.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Column));
				this.WordControl.m_oLogicDocument.FinalizeAction();
			}
		}
	};
	asc_docs_api.prototype.Update_ParaInd                = function(Ind)
	{
		var FirstLine = 0,
			Left      = 0,
			Right     = 0;
		if ("undefined" != typeof(Ind))
		{
			if ("undefined" != typeof(Ind.FirstLine))
			{
				FirstLine = Ind.FirstLine;
			}
			if ("undefined" != typeof(Ind.Left))
			{
				Left = Ind.Left;
			}
			if ("undefined" != typeof(Ind.Right))
			{
				Right = Ind.Right;
			}
		}

		var bIsUpdate = false;
		var _ruler    = this.WordControl.m_oHorRuler;
		if (_ruler.m_dIndentLeft != Left)
		{
			_ruler.m_dIndentLeft = Left;
			bIsUpdate            = true;
		}
		if (_ruler != (FirstLine + Left))
		{
			_ruler.m_dIndentLeftFirst = (FirstLine + Left);
			bIsUpdate                 = true;
		}
		if (_ruler.m_dIndentRight != Right)
		{
			_ruler.m_dIndentRight = Right;
			bIsUpdate             = true;
		}
		if (bIsUpdate)
			this.WordControl.UpdateHorRuler();
	};
	asc_docs_api.prototype.Internal_Update_Ind_FirstLine = function(FirstLine, Left)
	{
		if (this.WordControl.m_oHorRuler.m_dIndentLeftFirst != (FirstLine + Left))
		{
			this.WordControl.m_oHorRuler.m_dIndentLeftFirst = (FirstLine + Left);
			this.WordControl.UpdateHorRuler();
		}
	};
	asc_docs_api.prototype.Internal_Update_Ind_Left      = function(Left)
	{
		if (this.WordControl.m_oHorRuler.m_dIndentLeft != Left)
		{
			this.WordControl.m_oHorRuler.m_dIndentLeft = Left;
			this.WordControl.UpdateHorRuler();
		}
	};
	asc_docs_api.prototype.Internal_Update_Ind_Right     = function(Right)
	{
		if (this.WordControl.m_oHorRuler.m_dIndentRight != Right)
		{
			this.WordControl.m_oHorRuler.m_dIndentRight = Right;
			this.WordControl.UpdateHorRuler();
		}
	};

	// "where" где нижний или верхний, align выравнивание
	asc_docs_api.prototype.put_PageNum = function(where, align)
	{
		if (where >= 0)
		{
			if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_None, {Type : AscCommon.changestype_2_HdrFtr}))
			{
				this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddPageNumToHdrFtr);
				this.WordControl.m_oLogicDocument.Document_AddPageNum(where, align);
				this.WordControl.m_oLogicDocument.FinalizeAction();
			}
		}
		else
		{
			if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
			{
				this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddPageNumToCurrentPos);
				this.WordControl.m_oLogicDocument.Document_AddPageNum(where, align);
				this.WordControl.m_oLogicDocument.FinalizeAction();
			}
		}
	};

	asc_docs_api.prototype.put_HeadersAndFootersDistance = function(value)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_HdrFtr))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetHdrFtrDistance);
			this.WordControl.m_oLogicDocument.Document_SetHdrFtrDistance(value);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.HeadersAndFooters_DifferentFirstPage = function(isOn)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_HdrFtr))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetHdrFtrFirstPage);
			this.WordControl.m_oLogicDocument.Document_SetHdrFtrFirstPage(isOn);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.HeadersAndFooters_DifferentOddandEvenPage = function(isOn)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_HdrFtr))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetHdrFtrEvenAndOdd);
			this.WordControl.m_oLogicDocument.Document_SetHdrFtrEvenAndOddHeaders(isOn);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.HeadersAndFooters_LinkToPrevious = function(isOn)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_HdrFtr))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetHdrFtrLink);
			this.WordControl.m_oLogicDocument.Document_SetHdrFtrLink(isOn);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.asc_SetSectionStartPage = function(nStartPage)
	{
		if (isNaN(nStartPage))
			return;

		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		if (false === oLogicDocument.Document_Is_SelectionLocked())
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_SectionStartPage);
			oLogicDocument.SetSectionStartPage(nStartPage);
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_SetSectionPageNumFormat = function(format)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		return new Promise(function(resolve)
		{
			let symbols = AscWord.GetNumberingSymbolsByFormat(format);
			if (symbols && symbols.length)
				AscFonts.FontPickerByCharacter.checkText(symbols, this, resolve);
			else
				resolve();
		}).then(function()
		{
			if (logicDocument.IsSelectionLocked(AscCommon.changestype_HdrFtr))
				return;
			
			logicDocument.StartAction(AscDFH.historydescription_Document_SectionPageNumFormat);
			logicDocument.SetSectionPageNumFormat(format);
			logicDocument.FinalizeAction();
		});
	};

	/*структура для передачи настроек колонтитулов
	 {
	 Type : hdrftr_Footer (hdrftr_Header),
	 Position : 12.5,
	 DifferentFirst : true/false,
	 DifferentEvenOdd : true/false,
	 }
	 */
	/*callback*/
	asc_docs_api.prototype.sync_DocSizeCallback               = function(width, height)
	{
		this.sendEvent("asc_onDocSize", width, height);
	};
	asc_docs_api.prototype.sync_PageOrientCallback            = function(isPortrait)
	{
		this.sendEvent("asc_onPageOrient", isPortrait);
	};
	asc_docs_api.prototype.sync_HeadersAndFootersPropCallback = function(hafProp)
	{
		if (true === hafProp)
			hafProp.Locked = true;

		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.Header, new CHeaderProp(hafProp));
	};

	/*----------------------------------------------------------------*/
	/*functions for working with table*/
	asc_docs_api.prototype.put_Table               = function(col, row, sStyleId)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Document_Content_Add))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddTable);
			var oTable = this.WordControl.m_oLogicDocument.AddInlineTable(col, row);
			if(oTable && sStyleId)
			{
				oTable.Set_TableStyle(sStyleId);
			}
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.addRowAbove             = function(nCount)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Table_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_TableAddRowAbove);
			this.WordControl.m_oLogicDocument.AddTableRow(true, nCount);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.addRowBelow             = function(nCount)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Table_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_TableAddRowBelow);
			this.WordControl.m_oLogicDocument.AddTableRow(false, nCount);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.addColumnLeft           = function(nCount)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Table_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_TableAddColumnLeft);
			this.WordControl.m_oLogicDocument.AddTableColumn(true, nCount);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.addColumnRight          = function(nCount)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Table_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_TableAddColumnRight);
			this.WordControl.m_oLogicDocument.AddTableColumn(false, nCount);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.remRow                  = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return false;

		oLogicDocument.SelectTable(c_oAscTableSelectionType.Row);

		if (!oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Remove))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_TableRemoveRow);
			oLogicDocument.RemoveTableRow();
			oLogicDocument.FinalizeAction();

			return true;
		}
		else
		{
			return false;
		}
	};
	asc_docs_api.prototype.remColumn               = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return false;

		oLogicDocument.SelectTable(c_oAscTableSelectionType.Column);

		if (!oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Remove))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_TableRemoveColumn);
			oLogicDocument.RemoveTableCells();
			oLogicDocument.FinalizeAction();
			return true;
		}
		else
		{
			return false;
		}
	};
	asc_docs_api.prototype.remTable                = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return false;

		oLogicDocument.SelectTable(c_oAscTableSelectionType.Table);

		if (!oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Remove))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_RemoveTable);
			oLogicDocument.RemoveTable();
			oLogicDocument.FinalizeAction();

			return true;
		}
		else
		{
			return false;
		}
	};
	asc_docs_api.prototype.selectRow               = function()
	{
		this.WordControl.m_oLogicDocument.SelectTable(c_oAscTableSelectionType.Row);
	};
	asc_docs_api.prototype.selectColumn            = function()
	{
		this.WordControl.m_oLogicDocument.SelectTable(c_oAscTableSelectionType.Column);
	};
	asc_docs_api.prototype.selectCell              = function()
	{
		this.WordControl.m_oLogicDocument.SelectTable(c_oAscTableSelectionType.Cell);
	};
	asc_docs_api.prototype.selectTable             = function()
	{
		this.WordControl.m_oLogicDocument.SelectTable(c_oAscTableSelectionType.Table);
	};
	asc_docs_api.prototype.setColumnWidth          = function(width)
	{

	};
	asc_docs_api.prototype.setRowHeight            = function(height)
	{

	};
	asc_docs_api.prototype.set_TblDistanceFromText = function(left, top, right, bottom)
	{

	};
	asc_docs_api.prototype.CheckBeforeMergeCells   = function()
	{
		return this.WordControl.m_oLogicDocument.CanMergeTableCells();
	};
	asc_docs_api.prototype.CheckBeforeSplitCells   = function()
	{
		return this.WordControl.m_oLogicDocument.CanSplitTableCells();
	};
	asc_docs_api.prototype.MergeCells              = function()
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Table_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_MergeTableCells);
			this.WordControl.m_oLogicDocument.MergeTableCells();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.SplitCell               = function(Cols, Rows)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Table_Properties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SplitTableCells);
			this.WordControl.m_oLogicDocument.SplitTableCells(Cols, Rows);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_DistributeTableCells = function(isHorizontally)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		var bResult = false;

		if (false === oLogicDocument.Document_Is_SelectionLocked(changestype_Table_Properties))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_DistributeTableCells);
			bResult = oLogicDocument.DistributeTableCells(isHorizontally);
			oLogicDocument.FinalizeAction();
		}

		return bResult;
	};
	asc_docs_api.prototype.asc_RemoveTableCells    = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return false;

		if (!oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Remove))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_RemoveTableCells);
			oLogicDocument.RemoveTableCells();
			oLogicDocument.FinalizeAction();

			return true;
		}
		else
		{
			return false;
		}
	};
	asc_docs_api.prototype.widthTable              = function(width)
	{

	};
	asc_docs_api.prototype.put_CellsMargin         = function(left, top, right, bottom)
	{

	};
	asc_docs_api.prototype.set_TblWrap             = function(type)
	{

	};
	asc_docs_api.prototype.set_TblIndentLeft       = function(spacing)
	{

	};
	asc_docs_api.prototype.set_Borders             = function(typeBorders, size, Color)
	{//если size == 0 то границы нет.

	};
	asc_docs_api.prototype.set_TableBackground     = function(Color)
	{

	};
	asc_docs_api.prototype.set_AlignCell           = function(align)
	{// c_oAscAlignType.RIGHT, c_oAscAlignType.LEFT, c_oAscAlignType.CENTER
		switch (align)
		{
			case c_oAscAlignType.LEFT :
				break;
			case c_oAscAlignType.CENTER :
				break;
			case c_oAscAlignType.RIGHT :
				break;
		}
	};
	asc_docs_api.prototype.set_TblAlign            = function(align)
	{// c_oAscAlignType.RIGHT, c_oAscAlignType.LEFT, c_oAscAlignType.CENTER
		switch (align)
		{
			case c_oAscAlignType.LEFT :
				break;
			case c_oAscAlignType.CENTER :
				break;
			case c_oAscAlignType.RIGHT :
				break;
		}
	};
	asc_docs_api.prototype.set_SpacingBetweenCells = function(isOn, spacing)
	{// c_oAscAlignType.RIGHT, c_oAscAlignType.LEFT, c_oAscAlignType.CENTER
		if (isOn)
		{

		}
	};
	asc_docs_api.prototype.tblApply = function(obj)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Table_Properties))
		{
			if (obj.CellBorders)
			{
				if (obj.CellBorders.Left && obj.CellBorders.Left.Color)
				{
					obj.CellBorders.Left.Unifill = AscFormat.CreateUnifillFromAscColor(obj.CellBorders.Left.Color, 1);
				}
				if (obj.CellBorders.Top && obj.CellBorders.Top.Color)
				{
					obj.CellBorders.Top.Unifill = AscFormat.CreateUnifillFromAscColor(obj.CellBorders.Top.Color, 1);
				}
				if (obj.CellBorders.Right && obj.CellBorders.Right.Color)
				{
					obj.CellBorders.Right.Unifill = AscFormat.CreateUnifillFromAscColor(obj.CellBorders.Right.Color, 1);
				}
				if (obj.CellBorders.Bottom && obj.CellBorders.Bottom.Color)
				{
					obj.CellBorders.Bottom.Unifill = AscFormat.CreateUnifillFromAscColor(obj.CellBorders.Bottom.Color, 1);
				}
				if (obj.CellBorders.InsideH && obj.CellBorders.InsideH.Color)
				{
					obj.CellBorders.InsideH.Unifill = AscFormat.CreateUnifillFromAscColor(obj.CellBorders.InsideH.Color, 1);
				}
				if (obj.CellBorders.InsideV && obj.CellBorders.InsideV.Color)
				{
					obj.CellBorders.InsideV.Unifill = AscFormat.CreateUnifillFromAscColor(obj.CellBorders.InsideV.Color, 1);
				}
			}
			if (obj.CellsBackground && obj.CellsBackground.Color)
			{
				obj.CellsBackground.Unifill = AscFormat.CreateUnifillFromAscColor(obj.CellsBackground.Color, 1);
			}

			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_ApplyTablePr);
			this.WordControl.m_oLogicDocument.SetTableProps(obj);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	/*callbacks*/
	asc_docs_api.prototype.sync_AddTableCallback            = function()
	{
		this.sendEvent("asc_onAddTable");
	};
	asc_docs_api.prototype.sync_AlignCellCallback           = function(align)
	{
		this.sendEvent("asc_onAlignCell", align);
	};
	asc_docs_api.prototype.sync_TblPropCallback             = function(tblProp)
	{
		//if ( true === CollaborativeEditing.Get_GlobalLock() )
		//    tblProp.Locked = true;

		// TODO: вызвать функцию asc_onInitTableTemplatesв зависимости от TableLook
		if (tblProp.CellsBackground && tblProp.CellsBackground.Unifill)
		{
			var LogicDocument = this.WordControl.m_oLogicDocument;
			tblProp.CellsBackground.Unifill.check(LogicDocument.Get_Theme(), LogicDocument.Get_ColorMap());
			var RGBA                      = tblProp.CellsBackground.Unifill.getRGBAColor();
			tblProp.CellsBackground.Color = new AscCommonWord.CDocumentColor(RGBA.R, RGBA.G, RGBA.B, false);
		}
		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.Table, new Asc.CTableProp(tblProp));
	};
	asc_docs_api.prototype.sync_TblWrapStyleChangedCallback = function(style)
	{
		this.sendEvent("asc_onTblWrapStyleChanged", style);
	};
	asc_docs_api.prototype.sync_TblAlignChangedCallback     = function(style)
	{
		this.sendEvent("asc_onTblAlignChanged", style);
	};

	/*----------------------------------------------------------------*/
	/*functions for working with images*/
	asc_docs_api.prototype.ChangeImageFromFile      = function()
	{
		this.asc_addImage({isImageChangeUrl: true});
	};
	asc_docs_api.prototype.ChangeShapeImageFromFile = function(type)
	{
		this.asc_addImage({isShapeImageChangeUrl: true, textureType: type});
	};

	asc_docs_api.prototype.AddImage     = function()
	{
		this.asc_addImage();
	};

	asc_docs_api.prototype._addImageUrl      = function(arrUrls, oOptionObject)
	{
		if (oOptionObject)
		{
			if (oOptionObject.sendUrlsToFrameEditor && this.isOpenedChartFrame)
			{
				this.addImageUrlsFromGeneralToFrameEditor(arrUrls);
				return;
			}
			else if (oOptionObject.isImageChangeUrl || oOptionObject.isShapeImageChangeUrl || oOptionObject["obj"] || (oOptionObject instanceof AscCommon.CContentControlPr && oOptionObject.GetInternalId()) || oOptionObject.fAfterUploadOleObjectImage)
			{
				this.AddImageUrlAction(arrUrls[0], undefined, oOptionObject);
				return;
			}
		}

		if (this.ImageLoader)
		{
			const oApi = this;
			this.ImageLoader.LoadImagesWithCallback(arrUrls, function()
			{
				const sContentControlId = oApi.asc_GetCurrentContentControl();
				const oContentControl   = AscCommon.g_oTableId.Get_ById(sContentControlId);
				if (oContentControl && oContentControl.IsPicture() && arrUrls && arrUrls.length)
				{
					oApi.asc_SetContentControlPictureUrl(arrUrls[0], sContentControlId);
				}
				else if (oOptionObject instanceof AscCommon.DrawingPlaceholder)
				{
					const arrImages = [];
					for(let i = 0; i < arrUrls.length; ++i)
					{
						const oImage = oApi.ImageLoader.LoadImage(arrUrls[i], 1);
						if(oImage)
						{
							arrImages.push(oImage);
						}
					}
					if(arrImages.length)
					{
						oApi.WordControl.m_oLogicDocument.AddPlaceholderImages(arrImages, oOptionObject);
					}
				}
				else if (this.isPdfEditor() && oOptionObject && oOptionObject.GetType() === AscPDF.FIELD_TYPES.button)
				{
					const oImage = oApi.ImageLoader.LoadImage(arrUrls[0], 1);
					if(oImage && oImage.Image)
					{
						oOptionObject.AddImage(oImage);
					}
				}
				else if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
				{
					const arrImages = [];
					for(let i = 0; i < arrUrls.length; ++i)
					{
						const oImage = oApi.ImageLoader.LoadImage(arrUrls[i], 1);
						if(oImage)
						{
							arrImages.push(oImage);
						}
					}
					if(arrImages.length)
					{
						oApi.WordControl.m_oLogicDocument.StartAction();
						oApi.WordControl.m_oLogicDocument.AddImages(arrImages);
						oApi.WordControl.m_oLogicDocument.FinalizeAction();
					}
				}
			}, []);
		}
	};

	asc_docs_api.prototype.AddImageUrlAction = function(url, imgProp, obj)
	{
		var _image = this.ImageLoader.LoadImage(url, 1);
		if (null != _image)
		{
			var ColumnSize = this.WordControl.m_oLogicDocument.GetColumnSize();

			var _w = Math.max(1, ColumnSize.W);
			var _h = Math.max(1, ColumnSize.H);
			if (_image.Image != null)
			{
				var __w = Math.max((_image.Image.width * AscCommon.g_dKoef_pix_to_mm), 1);
				var __h = Math.max((_image.Image.height * AscCommon.g_dKoef_pix_to_mm), 1);
				_w      = Math.max(5, Math.min(_w, __w));
				_h      = Math.max(5, Math.min((_w * __h / __w)));
			}

			var src = _image.src;
			if (obj && obj.isShapeImageChangeUrl)
			{
				var AscShapeProp       = new Asc.asc_CShapeProperty();
				AscShapeProp.fill      = new asc_CShapeFill();
				AscShapeProp.fill.type = c_oAscFill.FILL_TYPE_BLIP;
				AscShapeProp.fill.fill = new asc_CFillBlip();
				AscShapeProp.fill.fill.asc_putUrl(src);
				if(obj.textureType !== null && obj.textureType !== undefined){
                    AscShapeProp.fill.fill.asc_putType(obj.textureType);
				}
				this.ImgApply(new asc_CImgProperty({ShapeProperties : AscShapeProp}));
			}
			else if (obj && obj.isImageChangeUrl)
			{
				var AscImageProp      = new asc_CImgProperty();
				AscImageProp.ImageUrl = src;
				this.ImgApply(AscImageProp);
			}
			else if (obj && obj.fAfterUploadOleObjectImage)
			{
				obj.fAfterUploadOleObjectImage(src);
			}
			else if (obj && obj["obj"] && obj["obj"].Get_Id)
			{
				this.asc_SetContentControlPictureUrl(src, obj["obj"].Get_Id());
			}
			else
			{
				if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
				{
					var imageLocal = g_oDocumentUrls.getImageLocal(src);
					if (imageLocal)
					{
						src = imageLocal;
					}

					this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddImageUrl);
					if (undefined === imgProp || undefined === imgProp.WrappingStyle || 0 == imgProp.WrappingStyle)
						this.WordControl.m_oLogicDocument.AddInlineImage(_w, _h, src);
					else
						this.WordControl.m_oLogicDocument.AddInlineImage(_w, _h, src, null, true);
					this.WordControl.m_oLogicDocument.FinalizeAction();
				}
			}
		}
		else
		{
			this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadImage);
			this.asyncImageEndLoaded2 = function(_image)
			{
				var ColumnSize = this.WordControl.m_oLogicDocument.GetColumnSize();

				var _w = Math.max(1, ColumnSize.W);
				var _h = Math.max(1, ColumnSize.H);
				if (_image.Image != null)
				{
					var __w = Math.max((_image.Image.width * AscCommon.g_dKoef_pix_to_mm), 1);
					var __h = Math.max((_image.Image.height * AscCommon.g_dKoef_pix_to_mm), 1);
					_w      = Math.max(5, Math.min(_w, __w));
					_h      = Math.max(5, Math.min((_w * __h / __w)));
				}
				var src = _image.src;

				if (obj && obj.isShapeImageChangeUrl)
				{
					var AscShapeProp       = new Asc.asc_CShapeProperty();
					AscShapeProp.fill      = new asc_CShapeFill();
					AscShapeProp.fill.type = c_oAscFill.FILL_TYPE_BLIP;
					AscShapeProp.fill.fill = new asc_CFillBlip();
					AscShapeProp.fill.fill.asc_putUrl(src);

                    if(obj.textureType !== null && obj.textureType !== undefined){
                        AscShapeProp.fill.fill.asc_putType(obj.textureType);
                    }
					this.ImgApply(new asc_CImgProperty({ShapeProperties : AscShapeProp}));
				}
				else if (obj && obj.isImageChangeUrl)
				{
					var AscImageProp      = new asc_CImgProperty();
					AscImageProp.ImageUrl = src;
					this.ImgApply(AscImageProp);
				}
				else if (obj && obj.fAfterUploadOleObjectImage)
				{
					obj.fAfterUploadOleObjectImage(src);
				}
				else if (obj && obj["obj"] && obj["obj"].Get_Id)
				{
					this.asc_SetContentControlPictureUrl(src, obj["obj"].Get_Id());
				}
				else if (obj instanceof AscCommon.CContentControlPr && obj.GetInternalId())
				{
					this.asc_SetContentControlPictureUrl(src, obj.GetInternalId());
				}
				else
				{

					if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
					{
						var imageLocal = g_oDocumentUrls.getImageLocal(src);
						if (imageLocal)
						{
							src = imageLocal;
						}
						this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddImageUrlLong);
						if (undefined === imgProp || undefined === imgProp.WrappingStyle || 0 == imgProp.WrappingStyle)
							this.WordControl.m_oLogicDocument.AddInlineImage(_w, _h, src);
						else
							this.WordControl.m_oLogicDocument.AddInlineImage(_w, _h, src, null, true);
						this.WordControl.m_oLogicDocument.FinalizeAction();
					}
				}
				this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadImage);

				this.asyncImageEndLoaded2 = null;
			}
		}
	};
	/*
	 Добавляем картинку на заданную страницу. Преполагаем, что картинка уже доступна по ссылке.
	 */
	asc_docs_api.prototype.AddImageToPage = function(sUrl, nPageIndex, dX, dY, dW, dH, nWrappingStyle)
	{
		var LogicDocument = this.WordControl.m_oLogicDocument;

		var oldClickCount            = global_mouseEvent.ClickCount;
		global_mouseEvent.Button     = 0;
		global_mouseEvent.ClickCount = 1;
		LogicDocument.OnMouseDown(global_mouseEvent, dX, dY, nPageIndex);
		LogicDocument.OnMouseUp(global_mouseEvent, dX, dY, nPageIndex);
		LogicDocument.OnMouseMove(global_mouseEvent, dX, dY, nPageIndex);
		global_mouseEvent.ClickCount = oldClickCount;

		if (false === LogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
		{
			var oPosH = new Asc.CImagePositionH();
			oPosH.put_RelativeFrom(Asc.c_oAscRelativeFromH.Column);
			oPosH.put_Align(false);
			oPosH.put_Value(0);
			var oPosV = new Asc.CImagePositionV();
			oPosV.put_RelativeFrom(Asc.c_oAscRelativeFromV.Paragraph);
			oPosV.put_Align(false);
			oPosV.put_Value(0);
			var oImageProps = new asc_CImgProperty();
			let nWrappingStyleToSet;
			if(nWrappingStyle !== undefined && nWrappingStyle !== null)
			{
				nWrappingStyleToSet = nWrappingStyle;
			}
			else
			{
				nWrappingStyleToSet = c_oAscWrapStyle2.Square;
			}
			oImageProps.asc_putWrappingStyle(nWrappingStyleToSet);
			oImageProps.asc_putPositionH(oPosH);
			oImageProps.asc_putPositionV(oPosV);

			// TODO: Убрать SilentMode, когда переделаются функции AddInlineImage, SetImageProps (в них нужно передалать пересчет и обновляения согласно схеме Actions)
			LogicDocument.StartAction(AscDFH.historydescription_Document_AddImageToPage);
			LogicDocument.Start_SilentMode();

			let oDrawingObjects = LogicDocument.DrawingObjects;
			let oImage = oDrawingObjects.createImage(sUrl, 0, 0, dW, dH);
			let oDrawing   = new AscCommonWord.ParaDrawing(dW, dH, oImage, LogicDocument.DrawingDocument, null, null);
			oImage.setParent(oDrawing);
			oDrawing.Set_GraphicObject(oImage);
			let oNearestPos = LogicDocument.Get_NearestPos(nPageIndex, dX, dY, true, oDrawing);
			if(oNearestPos)
			{
				oDrawing.Set_Props(oImageProps)
				oDrawing.Set_Distance( 3.2,  0,  3.2, 0 )
				oNearestPos.Paragraph.Check_NearestPos(oNearestPos);
				oNearestPos.Page = nPageIndex;
				oDrawing.Set_XYForAdd(dX, dY, oNearestPos, nPageIndex);
				oDrawing.AddToDocument(oNearestPos);
				oDrawing.CheckWH();
				oDrawingObjects.resetSelection();
				oImage.select(oDrawingObjects, nPageIndex);
			}
			LogicDocument.End_SilentMode(true);
			LogicDocument.FinalizeAction();
		}
	};
	/* В качестве параметра  передается объект класса Asc.asc_CImgProperty, он же приходит на OnImgProp
	 Asc.asc_CImgProperty заменяет пережнюю структуру:
	 если параметр не имеет значения то передвать следует null, напримере inline-картинок: в качестве left,top,bottom,right,X,Y,ImageUrl необходимо передавать null.
	 {
	 Width: 0,
	 Height: 0,
	 WrappingStyle: 0,
	 Paddings: { Left : 0, Top : 0, Bottom: 0, Right: 0 },
	 Position : {X : 0, Y : 0},
	 ImageUrl : ""
	 }
	 */

	asc_docs_api.prototype.asc_getSelectedDrawingObjectsCount = function()
	{
		if(!this.WordControl)
		{
			return 0;
		}
		if(!this.WordControl.m_oLogicDocument)
		{
			return 0;
		}
		return this.WordControl.m_oLogicDocument.GetSelectedDrawingObjectsCount();
	};


	asc_docs_api.prototype.put_ShapesAlign = function(type, alignType)
	{
		if(!this.WordControl)
		{
			return;
		}
		if(!this.WordControl.m_oLogicDocument)
		{
			return;
		}
		if(!AscFormat.isRealNumber(alignType))
		{
			alignType = Asc.c_oAscObjectsAlignType.Slide;
		}
		this.WordControl.m_oLogicDocument.PutShapesAlign(type, alignType);

	};
	asc_docs_api.prototype.DistributeHorizontally = function(alignType)
	{
		if(!AscFormat.isRealNumber(alignType))
		{
			alignType = Asc.c_oAscObjectsAlignType.Margin;
		}
		this.WordControl.m_oLogicDocument.DistributeDrawingsHorizontally(alignType);
	};
	asc_docs_api.prototype.DistributeVertically   = function(alignType)
	{
		if(!AscFormat.isRealNumber(alignType))
		{
			alignType = Asc.c_oAscObjectsAlignType.Margin;
		}
		this.WordControl.m_oLogicDocument.DistributeDrawingsVertically(alignType);
	};

	asc_docs_api.prototype.ImgApply                = function(obj)
	{

		if (!AscCommon.isRealObject(obj))
			return;
		var ImagePr = obj, AdditionalData, LogicDocument = this.WordControl.m_oLogicDocument;

		/*проверка корректности данных для биржевой диаграммы*/
		if (obj.ChartProperties && obj.ChartProperties.type === Asc.c_oAscChartTypeSettings.stock)
		{
			if (!AscFormat.CheckStockChart(this.WordControl.m_oLogicDocument.DrawingObjects, this))
			{
				return;
			}
		}

		/*изменение z-индекса*/
		if (AscFormat.isRealNumber(ImagePr.ChangeLevel))
		{
			switch (ImagePr.ChangeLevel)
			{
				case 0:
				{
					this.WordControl.m_oLogicDocument.DrawingObjects.bringToFront();
					break;
				}
				case 1:
				{
					this.WordControl.m_oLogicDocument.DrawingObjects.bringForward();
					break;
				}
				case 2:
				{
					this.WordControl.m_oLogicDocument.DrawingObjects.sendToBack();
					break;
				}
				case 3:
				{
					this.WordControl.m_oLogicDocument.DrawingObjects.bringBackward();
				}
			}
			return;
		}

		/*параграфы в которых лежат выделенные ParaDrawing*/
		var aParagraphs = [], aSelectedObjects = this.WordControl.m_oLogicDocument.DrawingObjects.selectedObjects, i, j, oParentParagraph;
		for (i = 0; i < aSelectedObjects.length; ++i)
		{
			oParentParagraph = aSelectedObjects[i].parent.Get_ParentParagraph();
			AscFormat.checkObjectInArray(aParagraphs, oParentParagraph);
		}


		AdditionalData = {
			Type      : AscCommon.changestype_2_ElementsArray_and_Type,
			Elements  : aParagraphs,
			CheckType : changestype_Paragraph_Content
		};
		/*группировка и разгруппировка*/
		if (ImagePr.Group === 1 || ImagePr.Group === -1)
		{
			if (false == this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props, AdditionalData))
			{
				this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_GroupUnGroup);
				if (ImagePr.Group === 1)
				{
					this.WordControl.m_oLogicDocument.DrawingObjects.groupSelectedObjects();
				}
				else
				{
					this.WordControl.m_oLogicDocument.DrawingObjects.unGroupSelectedObjects();
				}
				this.WordControl.m_oLogicDocument.FinalizeAction();
			}
			return;
		}


		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props))
		{
			if (ImagePr.ShapeProperties)
				ImagePr.ImageUrl = "";


			var sImageUrl = null, fReplaceCallback = null, bImageUrl = false, sImageToDownLoad = "", sToken = undefined;
			if (!AscCommon.isNullOrEmptyString(ImagePr.ImageUrl))
			{
				if (!g_oDocumentUrls.getImageLocal(ImagePr.ImageUrl))
				{
					sImageUrl        = ImagePr.ImageUrl;
					sToken           = ImagePr.Token;
					fReplaceCallback = function(sUrl)
					{
						ImagePr.ImageUrl = sUrl;
						sImageToDownLoad = sUrl;
					}
				}
				sImageToDownLoad = ImagePr.ImageUrl;
			}
			else if (ImagePr.ShapeProperties && ImagePr.ShapeProperties.fill &&
				ImagePr.ShapeProperties.fill.fill && !AscCommon.isNullOrEmptyString(ImagePr.ShapeProperties.fill.fill.url))
			{
				if (!g_oDocumentUrls.getImageLocal(ImagePr.ShapeProperties.fill.fill.url))
				{
					sImageUrl        = ImagePr.ShapeProperties.fill.fill.url;
					sToken           = ImagePr.ShapeProperties.fill.fill.token;
					fReplaceCallback = function(sUrl)
					{
						ImagePr.ShapeProperties.fill.fill.url = sUrl;
						sImageToDownLoad                      = sUrl;
					}
				}
				sImageToDownLoad = ImagePr.ShapeProperties.fill.fill.url;
			}

			var oApi = this;

			if (!AscCommon.isNullOrEmptyString(sImageToDownLoad))
			{

				var fApplyCallback = function()
				{
					var _img = oApi.ImageLoader.LoadImage(sImageToDownLoad, 1);
					if (null != _img)
					{
						oApi.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_ApplyImagePrWithUrl);
						oApi.WordControl.m_oLogicDocument.SetImageProps(ImagePr);
						oApi.WordControl.m_oLogicDocument.UpdateInterface();
						oApi.WordControl.m_oLogicDocument.UpdateSelection();
						oApi.WordControl.m_oLogicDocument.FinalizeAction();
					}
					else
					{
						oApi.asyncImageEndLoaded2 = function(_image)
						{
							oApi.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_ApplyImagePrWithUrlLong);
							oApi.WordControl.m_oLogicDocument.SetImageProps(ImagePr);
							oApi.WordControl.m_oLogicDocument.UpdateInterface();
							oApi.WordControl.m_oLogicDocument.UpdateSelection();
							oApi.WordControl.m_oLogicDocument.FinalizeAction();
						}
					}
				};

				if (sImageUrl)
				{

					if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]())
					{
						var _url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](sImageToDownLoad);
						_url     = g_oDocumentUrls.getImageUrl(_url);
						fReplaceCallback(_url);
						fApplyCallback();
						return;
					}

                    AscCommon.sendImgUrls(this, [sImageToDownLoad], function(data) {

                        if (data && data[0] && data[0].url !== "error")
                        {
                            fReplaceCallback(data[0].url);
                            fApplyCallback();
                        }

                    }, undefined, sToken);
				}
				else
				{
					fApplyCallback();
				}
			}
			else
			{
				ImagePr.ImageUrl = null;
				if (!this.noCreatePoint || this.exucuteHistory)
				{
					if (!this.noCreatePoint && !this.exucuteHistory && this.exucuteHistoryEnd)
					{
						if (-1 !== this.nCurPointItemsLength)
						{
							AscCommon.History.UndoLastPoint();
						}
						else
						{
							this.WordControl.m_oLogicDocument.Create_NewHistoryPoint(AscDFH.historydescription_Document_ApplyImagePr);
						}
						this.WordControl.m_oLogicDocument.SetImageProps(ImagePr);
						this.exucuteHistoryEnd    = false;
						this.nCurPointItemsLength = -1;
					}
					else
					{
						this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_ApplyImagePr);
						this.WordControl.m_oLogicDocument.SetImageProps(ImagePr);
						this.WordControl.m_oLogicDocument.UpdateInterface();
						this.WordControl.m_oLogicDocument.UpdateSelection();
						this.WordControl.m_oLogicDocument.FinalizeAction();
					}
					if (this.exucuteHistory)
					{
						this.exucuteHistory = false;
						var oPoint          = AscCommon.History.Points[AscCommon.History.Index];
						if (oPoint)
						{
							this.nCurPointItemsLength = oPoint.Items.length;
						}
					}
					if(this.exucuteHistoryEnd)
					{
						this.exucuteHistoryEnd = false;
					}
				}
				else
				{
					var bNeedCheckChangesCount = false;
					if (-1 !== this.nCurPointItemsLength)
					{
						AscCommon.History.UndoLastPoint();
					}
					else
					{
						bNeedCheckChangesCount = true;
						this.WordControl.m_oLogicDocument.Create_NewHistoryPoint(AscDFH.historydescription_Document_ApplyImagePr);
					}
					this.WordControl.m_oLogicDocument.SetImageProps(ImagePr);
					if (bNeedCheckChangesCount)
					{
						var oPoint = AscCommon.History.Points[AscCommon.History.Index];
						if (oPoint)
						{
							this.nCurPointItemsLength = oPoint.Items.length;
						}
					}
				}
                this.exucuteHistoryEnd = false;
			}
		}
	};
	asc_docs_api.prototype.set_Size                = function(width, height)
	{

	};
	asc_docs_api.prototype.set_ConstProportions    = function(isOn)
	{
		if (isOn)
		{

		}
		else
		{

		}
	};
	asc_docs_api.prototype.set_WrapStyle           = function(type)
	{

	};
	asc_docs_api.prototype.deleteImage             = function()
	{

	};
	asc_docs_api.prototype.set_ImgDistanceFromText = function(left, top, right, bottom)
	{

	};
	asc_docs_api.prototype.set_PositionOnPage      = function(X, Y)
	{//расположение от начала страницы

	};
	asc_docs_api.prototype.get_OriginalSizeImage   = function()
	{
        for(var i = 0; i < this.SelectedObjectsStack.length; ++i){
            if(this.SelectedObjectsStack[i].Type == c_oAscTypeSelectElement.Image && this.SelectedObjectsStack[i].Value && this.SelectedObjectsStack[i].Value.ImageUrl){
                return this.SelectedObjectsStack[i].Value.asc_getOriginSize(this);
            }
        }
        return null;
	};

	asc_docs_api.prototype.ShapeApply = function(shapeProps)
	{
		// нужно определить, картинка это или нет
		var image_url = "";
		if (shapeProps.fill != null)
		{
			if (shapeProps.fill.fill != null && shapeProps.fill.type == c_oAscFill.FILL_TYPE_BLIP)
			{
				image_url = shapeProps.fill.fill.asc_getUrl();

				var _tx_id = shapeProps.fill.fill.asc_getTextureId();
				if (null != _tx_id && 0 <= _tx_id && _tx_id < AscCommon.g_oUserTexturePresets.length)
				{
					image_url = AscCommon.g_oUserTexturePresets[_tx_id];
				}
			}
		}
		if (image_url != "")
		{
			var _image = this.ImageLoader.LoadImage(image_url, 1);

			var imageLocal = g_oDocumentUrls.getImageLocal(image_url);
			if (imageLocal)
			{
				shapeProps.fill.fill.asc_putUrl(imageLocal); // erase documentUrl
			}

			if (null != _image)
			{
				this.WordControl.m_oLogicDocument.ShapeApply(shapeProps);
				this.WordControl.m_oDrawingDocument.DrawImageTextureFillShape(image_url);
			}
			else
			{
				this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadImage);

				var oProp                 = shapeProps;
				this.asyncImageEndLoaded2 = function(_image)
				{
					this.WordControl.m_oLogicDocument.ShapeApply(oProp);
					this.WordControl.m_oDrawingDocument.DrawImageTextureFillShape(image_url);

					this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadImage);
					this.asyncImageEndLoaded2 = null;
				}
			}
		}
		else
		{
			this.WordControl.m_oLogicDocument.ShapeApply(shapeProps);
		}
	};
	/*callbacks*/
	asc_docs_api.prototype.sync_AddImageCallback            = function()
	{
		this.sendEvent("asc_onAddImage");
	};
	asc_docs_api.prototype.sync_ImgPropCallback             = function(imgProp)
	{
		//if ( true === CollaborativeEditing.Get_GlobalLock() )
		//    imgProp.Locked = true;

		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.Image, new asc_CImgProperty(imgProp));
	};
	asc_docs_api.prototype.sync_ImgWrapStyleChangedCallback = function(style)
	{
		this.sendEvent("asc_onImgWrapStyleChanged", style);
	};


	asc_docs_api.prototype.asc_addOleObjectAction = function(sLocalUrl, sData, sApplicationId, fWidth, fHeight, nWidthPix, nHeightPix, bSelect, arrImagesForAddToHistory)
	{
		var _image = this.ImageLoader.LoadImage(AscCommon.getFullImageSrc2(sLocalUrl), 1);
		if (null != _image)//картинка уже должна быть загружена
		{
            this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_PasteHotKey);
			let oDrawing = this.WordControl.m_oLogicDocument.AddOleObject(fWidth, fHeight, nWidthPix, nHeightPix, sLocalUrl, sData, sApplicationId, bSelect, arrImagesForAddToHistory);
            this.WordControl.m_oLogicDocument.FinalizeAction();
			if(window.g_asc_plugins)
			{
				if(oDrawing)
				{
					let oData = oDrawing.GraphicObj.getDataObject();
					window.g_asc_plugins.onPluginEvent("onInsertOleObjects", [oData]);
				}
			}
		}
	};

	asc_docs_api.prototype.asc_editOleObjectAction = function(oOleObject, sImageUrl, sData, fWidthMM, fHeightMM, nPixWidth, nPixHeight, arrImagesForAddToHistory)
	{
		if (oOleObject)
		{
            this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_PasteHotKey);
			this.WordControl.m_oLogicDocument.EditOleObject(oOleObject, sData, sImageUrl, fWidthMM, fHeightMM, nPixWidth, nPixHeight, arrImagesForAddToHistory);
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
            this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

    asc_docs_api.prototype.asc_startEditCurrentOleObject = function(){
		this.WordControl.m_oLogicDocument.DrawingObjects.startEditCurrentOleObject();
    };
	//-----------------------------------------------------------------
	// События контекстного меню
	//-----------------------------------------------------------------

	function CContextMenuData(obj)
	{
		if (obj)
		{
			this.Type  = ( undefined != obj.Type ) ? obj.Type : Asc.c_oAscContextMenuTypes.Common;
			this.X_abs = ( undefined != obj.X_abs ) ? obj.X_abs : 0;
			this.Y_abs = ( undefined != obj.Y_abs ) ? obj.Y_abs : 0;

			switch (this.Type)
			{
				case Asc.c_oAscContextMenuTypes.ChangeHdrFtr :
				{
					this.PageNum = ( undefined != obj.PageNum ) ? obj.PageNum : 0;
					this.Header  = ( undefined != obj.Header  ) ? obj.Header : true;

					break;
				}
			}
		}
		else
		{
			this.Type  = Asc.c_oAscContextMenuTypes.Common;
			this.X_abs = 0;
			this.Y_abs = 0;
		}
	}

	CContextMenuData.prototype.get_Type    = function()
	{
		return this.Type;
	};
	CContextMenuData.prototype.get_X       = function()
	{
		return this.X_abs;
	};
	CContextMenuData.prototype.get_Y       = function()
	{
		return this.Y_abs;
	};
	CContextMenuData.prototype.get_PageNum = function()
	{
		return this.PageNum;
	};
	CContextMenuData.prototype.is_Header   = function()
	{
		return this.Header;
	};

	asc_docs_api.prototype.sync_ContextMenuCallback = function(Data)
	{
		this.sendEvent("asc_onContextMenu", new CContextMenuData(Data));
	};


	asc_docs_api.prototype.sync_MouseMoveStartCallback = function()
	{
		this.sendEvent("asc_onMouseMoveStart");
	};

	asc_docs_api.prototype.sync_MouseMoveEndCallback = function()
	{
		this.sendEvent("asc_onMouseMoveEnd");
	};

	asc_docs_api.prototype.sync_MouseMoveCallback = function(Data)
	{
		this.sendEvent("asc_onMouseMove", Data);
	};

	asc_docs_api.prototype.sync_ShowForeignCursorLabel = function(UserId, X, Y, Color)
	{
		this.sendEvent("asc_onShowForeignCursorLabel", UserId, X, Y, new AscCommon.CColor(Color.r, Color.g, Color.b, 255));
	};
	asc_docs_api.prototype.sync_HideForeignCursorLabel = function(UserId)
	{
		this.sendEvent("asc_onHideForeignCursorLabel", UserId);
	};

	//-----------------------------------------------------------------
	// Функции для работы с гиперссылками
	//-----------------------------------------------------------------
	asc_docs_api.prototype.can_AddHyperlink = function()
	{
		if (!this.WordControl.m_oLogicDocument.CanAddHyperlink(true))
			return false;

		return this.WordControl.m_oLogicDocument.GetSelectedText(true, {Numbering : false});
	};
	/**
	 * Добавляем гиперссылку
	 * @param {CHyperlinkProperty} oHyperProps
	 */
	asc_docs_api.prototype.add_Hyperlink = function(oHyperProps)
	{
		var sBookmarkName = this.private_CheckHeadingHyperlinkProps(oHyperProps);

		if (null !== oHyperProps.Text && undefined !== oHyperProps.Text)
		{
			AscFonts.FontPickerByCharacter.checkText(oHyperProps.Text, this,
				function()
				{
					this.private_AddHyperlink(oHyperProps, sBookmarkName);
				});
		}
		else
		{
			this.private_AddHyperlink(oHyperProps, sBookmarkName);
		}
	};
	asc_docs_api.prototype.private_CheckHeadingHyperlinkProps = function(oHyperProps)
	{
		var oHeading      = oHyperProps.get_Heading();
		var sBookmarkName = null;

		if (oHeading)
		{
			var oBookmarksManager = this.WordControl.m_oLogicDocument.GetBookmarksManager();
			sBookmarkName         = oBookmarksManager.GetNameForHeadingBookmark(oHeading);
			var oBookmark         = oBookmarksManager.GetBookmarkByName(sBookmarkName);

			if (oBookmark && oBookmark[0].GetParagraph() === oHeading)
			{
				oHyperProps.put_Heading(null);
				oHyperProps.put_Bookmark(sBookmarkName);
				sBookmarkName = null;
			}
			else
			{
				if (oBookmark)
				{
					var nCounter = 1;
					while (oBookmarksManager.GetBookmarkByName(sBookmarkName + "_" + nCounter))
					{
						nCounter++;
					}

					sBookmarkName += "_" + nCounter;
				}

				oHyperProps.put_Bookmark(sBookmarkName);
			}
		}

		return sBookmarkName;
	};
	asc_docs_api.prototype.private_AddHyperlink = function(oHyperProps, sBookmarkName, oParagraph)
	{
		var isLocked = false;
		if (sBookmarkName)
		{
			isLocked = this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content, {
				Type      : AscCommon.changestype_2_Element_and_Type,
				Element   : oHyperProps.get_Heading(),
				CheckType : changestype_Paragraph_Content
			});
		}
		else
		{
			isLocked = this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content);
		}

		if (!isLocked)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddHyperlink);

			if (sBookmarkName && oHyperProps.get_Heading())
				oHyperProps.get_Heading().AddBookmarkAtBegin(sBookmarkName);

			this.WordControl.m_oLogicDocument.AddHyperlink(oHyperProps);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}

	};
	/**
	 * Изменяем гиперссылку
	 * @param {CHyperlinkProperty} oHyperProps
	 */
	asc_docs_api.prototype.change_Hyperlink = function(oHyperProps)
	{
		if (!oHyperProps || !oHyperProps.get_InternalHyperlink())
			return;

		var sBookmarkName = this.private_CheckHeadingHyperlinkProps(oHyperProps);

		if (null !== oHyperProps.Text && undefined !== oHyperProps.Text)
		{
			AscFonts.FontPickerByCharacter.checkText(oHyperProps.Text, this,
				function()
				{
					this.private_ChangeHyperlink(oHyperProps, sBookmarkName);
				});
		}
		else
		{
			this.private_ChangeHyperlink(oHyperProps, sBookmarkName);
		}
	};
	asc_docs_api.prototype.private_ChangeHyperlink = function(hyperProps, bookmarkName)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument
			|| !hyperProps
			|| !hyperProps.get_InternalHyperlink()
			|| !hyperProps.get_InternalHyperlink().IsUseInDocument())
			return;
		
		let additionalCheck = bookmarkName ? {
			Type      : AscCommon.changestype_2_Element_and_Type,
			Element   : hyperProps.get_Heading(),
			CheckType : changestype_Paragraph_Content
		} : undefined;
		
		if (logicDocument.IsSelectionLocked(changestype_Paragraph_Content, additionalCheck))
			return;
		
		logicDocument.StartAction(AscDFH.historydescription_Document_ChangeHyperlink);
		
		if (bookmarkName && hyperProps.get_Heading())
			hyperProps.get_Heading().AddBookmarkAtBegin(bookmarkName);
		
		logicDocument.ModifyHyperlink(hyperProps);
		logicDocument.FinalizeAction();
	};
	/**
	 * Удаляем гиперссылку
	 * @param {CHyperlinkProperty} hyperProps
	 */
	asc_docs_api.prototype.remove_Hyperlink = function(hyperProps)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument
			|| !hyperProps
			|| !hyperProps.get_InternalHyperlink()
			|| !hyperProps.get_InternalHyperlink().IsUseInDocument())
			return;
		
		if (logicDocument.IsSelectionLocked(changestype_Paragraph_Content))
			return;
		
		logicDocument.StartAction(AscDFH.historydescription_Document_RemoveHyperlink);
		logicDocument.RemoveHyperlink(hyperProps);
		logicDocument.FinalizeAction();
	};
	asc_docs_api.prototype.asc_GetHyperlinkAnchors = function()
	{
		if (!this.WordControl || !this.WordControl.m_oLogicDocument)
			return [];

		return this.WordControl.m_oLogicDocument.GetHyperlinkAnchors();
	};



	asc_docs_api.prototype.sync_HyperlinkPropCallback = function(hyperProp)
	{
		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.Hyperlink, new Asc.CHyperlinkProperty(hyperProp));
	};

	asc_docs_api.prototype.sync_HyperlinkClickCallback = function(Url)
	{
		this.sendEvent("asc_onHyperlinkClick", Url);
	};

	asc_docs_api.prototype.sync_CanAddHyperlinkCallback = function(bCanAdd)
	{
		//if ( true === CollaborativeEditing.Get_GlobalLock() )
		//    this.sendEvent("asc_onCanAddHyperlink", false);
		//else
		this.sendEvent("asc_onCanAddHyperlink", bCanAdd);
	};

	asc_docs_api.prototype.sync_DialogAddHyperlink = function()
	{
		this.sendEvent("asc_onDialogAddHyperlink");
	};

	asc_docs_api.prototype.sync_DialogAddHyperlink = function()
	{
		this.sendEvent("asc_onDialogAddHyperlink");
	};

	//-----------------------------------------------------------------
	// Функции для работы с орфографией
	//-----------------------------------------------------------------
	asc_docs_api.prototype.sync_SpellCheckCallback = function(Word, Checked, Variants, ParaId, Element)
	{
		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.SpellCheck, new AscCommon.asc_CSpellCheckProperty(Word, Checked, Variants, ParaId, Element));
	};

	asc_docs_api.prototype.sync_SpellCheckVariantsFound = function()
	{
		this.sendEvent("asc_onSpellCheckVariantsFound");
	};

	asc_docs_api.prototype.asc_replaceMisspelledWord = function(Word, SpellCheckProperty)
	{
		var ParaId = SpellCheckProperty.ParaId;

		var Paragraph = g_oTableId.Get_ById(ParaId);
		if (null != Paragraph && false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_None, {
				Type      : AscCommon.changestype_2_Element_and_Type,
				Element   : Paragraph,
				CheckType : changestype_Paragraph_Content
			}))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_ReplaceMisspelledWord);
			Paragraph.ReplaceMisspelledWord(Word, SpellCheckProperty.Element);
			Paragraph.Document_SetThisElementCurrent(true);
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.UpdateSelection();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.asc_ignoreMisspelledWord = function(SpellCheckProperty, bAll)
	{
		if (false === bAll)
		{
			var ParaId = SpellCheckProperty.ParaId;

			var Paragraph = g_oTableId.Get_ById(ParaId);
			if (null != Paragraph)
			{
				Paragraph.IgnoreMisspelledWord(SpellCheckProperty.Element);
			}
		}
		else
		{
			var LogicDocument = editor.WordControl.m_oLogicDocument;
			LogicDocument.Spelling.AddToIgnore(SpellCheckProperty.Word);
			LogicDocument.DrawingDocument.ClearCachePages();
			LogicDocument.DrawingDocument.FirePaint();
		}
	};

    asc_docs_api.prototype._spellCheckRestart = function(word)
    {
		var LogicDocument = this.WordControl.m_oLogicDocument;
		if (LogicDocument)
		{
			// TODO: сделать нормальный сброс слова
			var oldWordStatus = LogicDocument.Spelling.IsIgnored(word);
			if (true !== oldWordStatus)
			{
				LogicDocument.Spelling.AddToIgnore(word);
				LogicDocument.DrawingDocument.ClearCachePages();
				LogicDocument.DrawingDocument.FirePaint();
				delete LogicDocument.Spelling.Words[word];
			}
		}
    };
    asc_docs_api.prototype.asc_spellCheckClearDictionary = function()
    {
        if (window["AscDesktopEditor"])
            window["AscDesktopEditor"]["SpellCheck"]("{\"type\":\"clear\"}");
    };

	asc_docs_api.prototype.asc_setDefaultLanguage = function(Lang)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Document_SectPr))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetDefaultLanguage);
			this.WordControl.m_oLogicDocument.SetDefaultLanguage(Lang);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.asc_getDefaultLanguage = function()
	{
		return editor.WordControl.m_oLogicDocument.GetDefaultLanguage();
	};

	asc_docs_api.prototype.asc_getKeyboardLanguage = function()
	{
		if (undefined !== window["asc_current_keyboard_layout"])
			return window["asc_current_keyboard_layout"];
		return -1;
	};
	
	asc_docs_api.prototype.asc_getInputLanguage = function()
	{
		let keyboardLang = this.asc_getKeyboardLanguage();
		if (-1 !== keyboardLang)
			return keyboardLang;
		
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return lcid_enUS;
		
		let run = logicDocument.GetCurrentRun();
		if (!run)
			return lcid_enUS;
		
		return run.Get_CompiledPr(false).Lang.Val;
	};

	asc_docs_api.prototype.asc_setSpellCheck = function(isOn)
	{
		if (editor.WordControl.m_oLogicDocument)
		{
			editor.WordControl.m_oLogicDocument.Spelling.Use = isOn;
			editor.WordControl.m_oDrawingDocument.ClearCachePages();
			editor.WordControl.m_oDrawingDocument.FirePaint();
		}
	};

	asc_docs_api.prototype.asc_restartCheckSpelling = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (oLogicDocument)
			oLogicDocument.RestartSpellCheck();
	};

	asc_docs_api.prototype.asc_setSpellCheckSettings = function(oSettings)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		let oSpellCheck = oLogicDocument.GetSpellCheckManager();
		if (oSpellCheck.SetSettings(oSettings))
			oLogicDocument.RestartSpellCheck();
	};
	asc_docs_api.prototype.asc_getSpellCheckSettings = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return new AscCommon.CSpellCheckSettings();

		return oLogicDocument.GetSpellCheckManager().GetSettings();
	};

	//-----------------------------------------------------------------
	// Функции для работы с комментариями
	//-----------------------------------------------------------------
	function asc_CCommentDataWord(obj)
	{
		if (obj)
		{
			this.m_bGlobal    = (undefined != obj.m_bGlobal   ) ? obj.m_bGlobal : false;
			this.m_sText      = (undefined != obj.m_sText     ) ? obj.m_sText : "";
			this.m_sTime      = (undefined != obj.m_sTime     ) ? obj.m_sTime : "";
			this.m_sOOTime    = (undefined != obj.m_sOOTime   ) ? obj.m_sOOTime : "";
			this.m_sUserId    = (undefined != obj.m_sUserId   ) ? obj.m_sUserId : "";
			this.m_sProviderId= (undefined != obj.m_sProviderId) ? obj.m_sProviderId : "";
			this.m_sQuoteText = (undefined != obj.m_sQuoteText) ? obj.m_sQuoteText : null;
			this.m_bSolved    = (undefined != obj.m_bSolved   ) ? obj.m_bSolved : false;
			this.m_sUserName  = (undefined != obj.m_sUserName ) ? obj.m_sUserName : "";
			this.m_sInitials  = (undefined != obj.m_sInitials ) ? obj.m_sInitials : this.asc_makeInitials(this.m_sUserName);
			this.m_nDurableId = (undefined != obj.m_nDurableId) ? obj.m_nDurableId : AscCommon.CreateDurableId();
			this.m_sUserData  = (undefined !== obj.m_sUserData) ? obj.m_sUserData : "";
			this.m_aReplies   = [];
			if (undefined != obj.m_aReplies)
			{
				var Count = obj.m_aReplies.length;
				for (var Index = 0; Index < Count; Index++)
				{
					var Reply = new asc_CCommentDataWord(obj.m_aReplies[Index]);
					this.m_aReplies.push(Reply);
				}
			}
		}
		else
		{
			this.m_bGlobal    = false;
			this.m_sText      = "";
			this.m_sTime      = "";
			this.m_sOOTime    = "";
			this.m_sUserId    = "";
			this.m_sProviderId= "";
			this.m_sQuoteText = null;
			this.m_bSolved    = false;
			this.m_sUserName  = "";
			this.m_sInitials  = "";
			this.m_nDurableId = AscCommon.CreateDurableId();
			this.m_sUserData  = "";
			this.m_aReplies   = [];
		}
	}

	asc_CCommentDataWord.prototype.asc_getText         = function()
	{
		return this.m_sText;
	};
	asc_CCommentDataWord.prototype.asc_putText         = function(v)
	{
		this.m_sText = v ? v.slice(0, Asc.c_oAscMaxCellOrCommentLength) : v;
	};
	asc_CCommentDataWord.prototype.asc_getTime         = function()
	{
		return this.m_sTime;
	};
	asc_CCommentDataWord.prototype.asc_putTime         = function(v)
	{
		this.m_sTime = v;
	};
	asc_CCommentDataWord.prototype.asc_getOnlyOfficeTime         = function()
	{
		return this.m_sOOTime;
	};
	asc_CCommentDataWord.prototype.asc_putOnlyOfficeTime         = function(v)
	{
		this.m_sOOTime = v;
	};
	asc_CCommentDataWord.prototype.asc_getUserId       = function()
	{
		return this.m_sUserId;
	};
	asc_CCommentDataWord.prototype.asc_putUserId       = function(v)
	{
		this.m_sUserId = v;
		this.m_sProviderId = "Teamlab";
	};
	asc_CCommentDataWord.prototype.asc_getProviderId       = function()
	{
		return this.m_sProviderId;
	};
	asc_CCommentDataWord.prototype.asc_putProviderId       = function(v)
	{
		this.m_sProviderId = v;
	};
	asc_CCommentDataWord.prototype.asc_getUserName     = function()
	{
		return this.m_sUserName;
	};
	asc_CCommentDataWord.prototype.asc_putUserName     = function(v)
	{
		this.m_sUserName = v;
		this.m_sInitials = this.asc_makeInitials(this.m_sUserName);
	};
	asc_CCommentDataWord.prototype.asc_getInitials       = function()
	{
		return this.m_sInitials;
	};
	asc_CCommentDataWord.prototype.asc_putInitials       = function(v)
	{
		this.m_sInitials = v;
	};
	asc_CCommentDataWord.prototype.asc_getQuoteText    = function()
	{
		return this.m_sQuoteText;
	};
	asc_CCommentDataWord.prototype.asc_putQuoteText    = function(v)
	{
		this.m_sQuoteText = v;
	};
	asc_CCommentDataWord.prototype.asc_getSolved       = function()
	{
		return this.m_bSolved;
	};
	asc_CCommentDataWord.prototype.asc_putSolved       = function(v)
	{
		this.m_bSolved = v;
	};
	asc_CCommentDataWord.prototype.asc_getGuid       = function()
	{
		return this.m_nDurableId.toString(16).padStart(8, "0");
	};
	asc_CCommentDataWord.prototype.asc_putGuid       = function(v)
	{
		this.m_nDurableId = parseInt(v, 16);
	};
	asc_CCommentDataWord.prototype.asc_getDurableId       = function()
	{
		return this.m_nDurableId;
	};
	asc_CCommentDataWord.prototype.asc_getReply        = function(i)
	{
		return this.m_aReplies[i];
	};
	asc_CCommentDataWord.prototype.asc_addReply        = function(v)
	{
		this.m_aReplies.push(v);
	};
	asc_CCommentDataWord.prototype.asc_getRepliesCount = function(v)
	{
		return this.m_aReplies.length;
	};
	asc_CCommentDataWord.prototype.asc_putDocumentFlag = function(val)
	{
		this.m_bGlobal = val;
	};
	asc_CCommentDataWord.prototype.asc_getDocumentFlag = function()
	{
		return this.m_bGlobal;
	};
	asc_CCommentDataWord.prototype.asc_makeInitials = function(name)
	{
		var initials = "";
		if(name){
			name.split(" ").forEach(function(elem) {
				if (elem.length > 0) {
					initials += elem[0];
				}
			});
		}
		return initials;
	};
	asc_CCommentDataWord.prototype.asc_getUserData = function()
	{
		return this.m_sUserData;
	};
	asc_CCommentDataWord.prototype.asc_putUserData = function(v)
	{
		this.m_sUserData = v;
	};
	asc_CCommentDataWord.prototype.fromCValue = function (value) {
		if (!value) {
			return value;
		}
		let pr = new AscWord.CSdtPictureFormPr();
		pr.asc_putText(value["Text"]);
		pr.asc_putTime(value["Time"]);
		pr.asc_putOnlyOfficeTime(value["OnlyOfficeTime"]);
		pr.asc_putUserId(value["UserId"]);
		pr.asc_putProviderId(value["ProviderId"]);
		pr.asc_putUserName(value["UserName"]);
		pr.asc_putInitials(value["Initials"]);
		pr.asc_putGuid(value["Guid"]);
		pr.asc_putTimeZoneBias(value["TimeZoneBias"]);
		pr.asc_putQuoteText(value["QuoteText"]);
		pr.asc_putSolved(value["Solved"]);

		if (Array.isArray(value["asc_getReplies"])) {
			for (let nIdx = 0; nIdx < value["asc_getReplies"].length; ++nIdx) {
				let reply = Asc.asc_CCommentData.prototype.fromCValue(value["asc_getReplies"][nIdx]);
				if (reply) {
					this.asc_addReply(reply);
				}
			}
		}
		pr.asc_putDocumentFlag(value["DocumentFlag"]);
		pr.asc_putUserData(value["UserData"]);
		return pr;
	};
	asc_CCommentDataWord.prototype.toCValue = function () {
		let value = {};
		value["Text"] = this.asc_getText();
		value["Time"] = this.asc_getTime();
		value["OnlyOfficeTime"] = this.asc_getOnlyOfficeTime();
		value["UserId"] = this.asc_getUserId();
		value["ProviderId"] = this.asc_getProviderId();
		value["UserName"] = this.asc_getUserName();
		value["Initials"] = this.asc_getInitials();
		value["Guid"] = this.asc_getGuid();
		value["TimeZoneBias"] = this.asc_getTimeZoneBias();
		value["QuoteText"] = this.asc_getQuoteText();
		value["Solved"] = this.asc_getSolved();
		value["asc_getReplies"] = [];
		let Count = this.m_aReplies.length;
		for (let nIdx = 0; nIdx < Count; nIdx++) {
			let val = this.m_aReplies[nIdx].toCValue();
			if (val) {
				value["asc_getReplies"].push(val);
			}
		}
		value["DocumentFlag"] = this.asc_getDocumentFlag();
		value["UserData"] = this.asc_getUserData();
		return value;
	};

	asc_docs_api.prototype.asc_showComments = function(isShowSolved)
	{
		if (null == this.WordControl.m_oLogicDocument)
			return;

		this.WordControl.m_oLogicDocument.ShowComments(isShowSolved);
	};

	asc_docs_api.prototype.asc_hideComments = function()
	{
		if (null == this.WordControl.m_oLogicDocument)
			return;

		this.WordControl.m_oLogicDocument.HideComments();
		editor.sync_HideComment();
	};

	asc_docs_api.prototype.asc_addComment = function(AscCommentData)
	{
		if (true === AscCommon.CollaborativeEditing.Get_GlobalLock())
			return null;

		var oLogicDocument = this.WordControl.m_oLogicDocument;

		if (!oLogicDocument)
			return null;

		// Комментарий без цитаты позволяем добавить всегда
		if (true !== this.can_AddQuotedComment() || false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, oLogicDocument.IsEditCommentsMode()))
		{
			var CommentData = new AscCommon.CCommentData();
			CommentData.Read_FromAscCommentData(AscCommentData);

			if ("" === CommentData.m_sTime)
				CommentData.m_sTime = ((new Date()).getTime() - (new Date()).getTimezoneOffset() * 60000).toString();

			if ("" === CommentData.m_sOOTime)
				CommentData.m_sOOTime = ((new Date()).getTime()).toString();

			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddComment);
			var oComment = this.WordControl.m_oLogicDocument.AddComment(CommentData, AscCommentData.asc_getDocumentFlag());
			if (oComment)
				this.sync_AddComment(oComment.GetId(), CommentData);

			this.WordControl.m_oLogicDocument.FinalizeAction();

			if (oComment)
				return oComment.GetId();
		}

		return null;
	};

	asc_docs_api.prototype.asc_removeComment = function(Id)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		if (false === oLogicDocument.Document_Is_SelectionLocked(changestype_None, {
				Type : AscCommon.changestype_2_Comment,
				Id   : Id
			}, false, oLogicDocument.IsEditCommentsMode()))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_RemoveComment);
			oLogicDocument.RemoveComment(Id, true, true);
			oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.asc_changeComment = function(Id, AscCommentData)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		if (false === oLogicDocument.Document_Is_SelectionLocked(changestype_None, {
				Type : AscCommon.changestype_2_Comment,
				Id   : Id
			}, false, oLogicDocument.IsEditCommentsMode()))
		{
			var CommentData = new AscCommon.CCommentData();
			CommentData.Read_FromAscCommentData(AscCommentData);

			oLogicDocument.StartAction(AscDFH.historydescription_Document_ChangeComment);
			oLogicDocument.EditComment(Id, CommentData);
			oLogicDocument.FinalizeAction();

			this.sync_ChangeCommentData(Id, CommentData);
		}
	};

	asc_docs_api.prototype.asc_selectComment = function(Id)
	{
		if (null == this.WordControl.m_oLogicDocument)
			return;

		this.WordControl.m_oLogicDocument.SelectComment(Id, true);
	};

	asc_docs_api.prototype.asc_showComment = function(Id)
	{
		if (Id instanceof Array)
			this.WordControl.m_oLogicDocument.ShowComment(Id);
		else
			this.WordControl.m_oLogicDocument.ShowComment([Id]);
	};

	asc_docs_api.prototype.asc_GetCommentsReportByAuthors = function()
	{
		var oReport = {};

		function privateProcessCommentData(isTopComment, oCommentData)
		{
			var sUserName = oCommentData.GetUserName();
			var nDateTime = oCommentData.GetDateTime();

			if (!oReport[sUserName])
				oReport[sUserName] = [];

			var arrUserComments = oReport[sUserName];

			var nPos = 0;
			var nLen = arrUserComments.length;
			while (nPos < nLen)
			{
				if (nDateTime < arrUserComments[nPos].Data.GetDateTime())
					break;

				nPos++;
			}

			arrUserComments.splice(nPos, 0, {Top : isTopComment, Data : oCommentData});

			for (var nIndex = 0, nCount = oCommentData.GetRepliesCount(); nIndex < nCount; ++nIndex)
			{
				privateProcessCommentData(false, oCommentData.GetReply(nIndex))
			}
		}

		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return oReport;

		var oAllComments = oLogicDocument.Comments.GetAllComments();
		for (var sId in oAllComments)
		{
			var oComment = oAllComments[sId];
			privateProcessCommentData(true, oComment.GetData());
		}

		return oReport;
	};

	asc_docs_api.prototype.can_AddQuotedComment = function()
	{
		//if ( true === CollaborativeEditing.Get_GlobalLock() )
		//    return false;

		return this.WordControl.m_oLogicDocument.CanAddComment();
	};

	asc_docs_api.prototype.sync_RemoveComment = function(Id)
	{
		this.sendEvent("asc_onRemoveComment", Id);

		if (window.g_asc_plugins)
			window.g_asc_plugins.onPluginEvent("onRemoveComment", {"Id" : Id});
	};

	asc_docs_api.prototype.sync_AddComment = function(Id, CommentData)
	{
		var AscCommentData = new asc_CCommentDataWord(CommentData);
		this.sendEvent("asc_onAddComment", Id, AscCommentData);

		if (window.g_asc_plugins)
			window.g_asc_plugins.onPluginEvent("onAddComment", {"Id" : Id, "Data" : CommentData.ConvertToSimpleObject()});
	};

	asc_docs_api.prototype.sync_ShowComment = function(arrId, X, Y)
	{
		this.sendEvent("asc_onShowComment", arrId, X, Y);
	};

	asc_docs_api.prototype.sync_HideComment = function()
	{
		this.sendEvent("asc_onHideComment");
	};

	asc_docs_api.prototype.sync_UpdateCommentPosition = function(Id, X, Y)
	{
		// TODO: Переделать на нормальный массив
		this.sendEvent("asc_onUpdateCommentPosition", [Id], X, Y);
	};

	asc_docs_api.prototype.sync_ChangeCommentData = function(Id, CommentData)
	{
		var AscCommentData = new asc_CCommentDataWord(CommentData);
		this.sendEvent("asc_onChangeCommentData", Id, AscCommentData);

		if (window.g_asc_plugins)
			window.g_asc_plugins.onPluginEvent("onChangeCommentData", {"Id" : Id, "Data" : CommentData.ConvertToSimpleObject()});
	};

	asc_docs_api.prototype.sync_LockComment = function(Id, UserId)
	{
		this.sendEvent("asc_onLockComment", Id, UserId);
	};

	asc_docs_api.prototype.sync_UnLockComment = function(Id)
	{
		this.sendEvent("asc_onUnLockComment", Id);
	};
	asc_docs_api.prototype.asc_RemoveAllComments = function(isMine, isCurrent, arrIds)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var arrCommentsId = undefined !== arrIds ? arrIds : oLogicDocument.GetAllComments(isMine, isCurrent);

		if (!oLogicDocument.IsSelectionLocked(changestype_None, {
				Type : AscCommon.changestype_2_Comment,
				Id   : arrCommentsId
			}, false, oLogicDocument.IsEditCommentsMode()))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_RemoveAllComments);

			for (var nIndex = 0, nCount = arrCommentsId.length; nIndex < nCount; ++nIndex)
			{
				var oComment = oLogicDocument.Comments.Get_ById(arrCommentsId[nIndex]);
				if (oComment && AscCommon.UserInfoParser.canDeleteComment(oComment.GetUserName()))
					oLogicDocument.RemoveComment(arrCommentsId[nIndex], true, false);
			}

			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_ResolveAllComments = function(isMine, isCurrent, arrIds)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var arrCommentsId = undefined !== arrIds ? arrIds : oLogicDocument.GetAllComments(isMine, isCurrent);

		if (!oLogicDocument.IsSelectionLocked(changestype_None, {
			Type : AscCommon.changestype_2_Comment,
			Id   : arrCommentsId
		}, false, oLogicDocument.IsEditCommentsMode()))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_ResolveAllComments);

			for (var nIndex = 0, nCount = arrCommentsId.length; nIndex < nCount; ++nIndex)
			{
				var oComment = oLogicDocument.Comments.Get_ById(arrCommentsId[nIndex]);
				if (oComment && AscCommon.UserInfoParser.canEditComment(oComment.GetUserName()) && !oComment.IsSolved() && oComment.GetData())
				{
					var oCommentData = oComment.GetData().Copy();
					oCommentData.SetSolved(true);
					oComment.SetData(oCommentData);
					this.sync_ChangeCommentData(arrCommentsId[nIndex], oCommentData);
				}
			}

			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_GetCommentLogicPosition = function(sId)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return -1;

		var oCommentsManager = oLogicDocument.GetCommentsManager();
		var oComment = oCommentsManager.GetById(sId);
		if (!oComment)
			return -1;

		return oComment.GetPosition();
	};
	asc_docs_api.prototype.sync_ChangeCommentLogicalPosition = function(oChangedComments, nOverallCount)
	{
		var isChanged = false;
		for (var sId in oChangedComments)
		{
			isChanged = true;
			break;
		}

		if (!isChanged)
			return;

		this.sendEvent("asc_onChangeCommentLogicalPosition", oChangedComments, nOverallCount);
	};

	//-----------------------------------------------------------------
	asc_docs_api.prototype.sync_LockHeaderFooters = function()
	{
		this.sendEvent("asc_onLockHeaderFooters");
	};

	asc_docs_api.prototype.sync_LockDocumentProps = function()
	{
		this.sendEvent("asc_onLockDocumentProps");
	};

	asc_docs_api.prototype.sync_UnLockHeaderFooters = function()
	{
		this.sendEvent("asc_onUnLockHeaderFooters");
	};

	asc_docs_api.prototype.sync_UnLockDocumentProps = function()
	{
		this.sendEvent("asc_onUnLockDocumentProps");
	};

	asc_docs_api.prototype.sync_CollaborativeChanges = function()
	{
		if (true !== AscCommon.CollaborativeEditing.Is_Fast() && (true !== this.WordControl.m_oLogicDocument.IsViewModeInReview() || true !== this.WordControl.m_oLogicDocument.IsFastCollaborationBeforeViewModeInReview()))
			this.sendEvent("asc_onCollaborativeChanges");
	};

	asc_docs_api.prototype.sync_LockDocumentSchema = function()
	{
		this.sendEvent("asc_onLockDocumentSchema");
	};

	asc_docs_api.prototype.sync_UnLockDocumentSchema = function()
	{
		this.sendEvent("asc_onUnLockDocumentSchema");
	};


	/*----------------------------------------------------------------*/
	/*functions for working with zoom & navigation*/
	asc_docs_api.prototype.zoomIn         = function()
	{
		this.WordControl.zoom_In();
	};
	asc_docs_api.prototype.zoomOut        = function()
	{
		this.WordControl.zoom_Out();
	};
	asc_docs_api.prototype.zoomFitToPage  = function()
	{
		if (!this.isLoadFullApi)
		{
			this.tmpZoomType = AscCommon.c_oZoomType.FitToPage;
			return;
		}
		this.WordControl.zoom_FitToPage();
	};
	asc_docs_api.prototype.zoomFitToWidth = function()
	{
		if (!this.isLoadFullApi)
		{
			this.tmpZoomType = AscCommon.c_oZoomType.FitToWidth;
			return;
		}
		this.WordControl.zoom_FitToWidth();
	};
	asc_docs_api.prototype.zoomCustomMode = function()
	{
		if (!this.isLoadFullApi)
		{
			this.tmpZoomType = AscCommon.c_oZoomType.CustomMode;
			return;
		}
		this.WordControl.m_nZoomType = 0;
		this.WordControl.zoom_Fire(0, this.WordControl.m_nZoomValue);
	};
	asc_docs_api.prototype.zoom100        = function()
	{
		this.zoom(100);
	};
	asc_docs_api.prototype.zoom           = function(percent)
	{
		var _old_val                  = this.WordControl.m_nZoomValue;
		this.WordControl.m_nZoomValue = percent;
		this.WordControl.m_nZoomType  = 0;
		this.WordControl.zoom_Fire(0, _old_val);
	};
	asc_docs_api.prototype.goToPage       = function(number)
	{
		this.WordControl.GoToPage(number);
	};
	asc_docs_api.prototype.getCountPages  = function()
	{
		return this.WordControl.m_oDrawingDocument.m_lPagesCount;
	};
	asc_docs_api.prototype.getCurrentPage = function()
	{
		return this.WordControl.m_oDrawingDocument.m_lCurrentPage;
	};
	/*callbacks*/
	asc_docs_api.prototype.sync_zoomChangeCallback  = function(percent, type)
	{	//c_oAscZoomType.Current, c_oAscZoomType.FitWidth, c_oAscZoomType.FitPage
		if (this.WordControl)
		{
			if (percent > this.WordControl.m_nZoomValueMax || this.WordControl.m_nZoomValueMax === -1)
				this.WordControl.m_nZoomValueMax = percent;
			if (percent < this.WordControl.m_nZoomValueMin || this.WordControl.m_nZoomValueMin === -1)
				this.WordControl.m_nZoomValueMin = percent;
		}
		this.sendEvent("asc_onZoomChange", percent, type);
	};
	asc_docs_api.prototype.sync_countPagesCallback  = function(count)
	{
		this.sendEvent("asc_onCountPages", count);
	};
	asc_docs_api.prototype.sync_currentPageCallback = function(number)
	{
		this.sendEvent("asc_onCurrentPage", number);
	};

	/*----------------------------------------------------------------*/
	asc_docs_api.prototype.asc_enableKeyEvents = function(value, isFromInput)
	{
		if (!this.isLoadFullApi)
		{
			this.tmpFocus = value;
			return;
		}

		if (this.WordControl && this.WordControl.IsFocus != value)
		{
			this.WordControl.IsFocus = value;

			if (this.WordControl.IsFocus && null != this.WordControl.TextBoxInput)
				this.WordControl.TextBoxInput.focus();

			this.sendEvent("asc_onEnableKeyEventsChanged", value);
		}

		if (isFromInput !== true && AscCommon.g_inputContext)
			AscCommon.g_inputContext.setInterfaceEnableKeyEvents(value);
	};
	asc_docs_api.prototype.asc_IsFocus         = function(bIsNaturalFocus)
	{
		var _ret = false;
		if (this.WordControl.IsFocus)
			_ret = true;
		if (_ret && bIsNaturalFocus && this.WordControl.TextBoxInputFocus)
			_ret = false;
		return _ret;
	};

	// работа с шрифтами
	asc_docs_api.prototype.asyncFontsDocumentStartLoaded = function(blockType)
	{
		// здесь прокинуть евент о заморозке меню
		// и нужно вывести информацию в статус бар
		if (this.isPasteFonts_Images)
			this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadFont);
		else if (this.isSaveFonts_Images)
			this.sync_StartAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadFont);
		else
		{
			this.sync_StartAction(undefined === blockType ? c_oAscAsyncActionType.BlockInteraction : blockType, c_oAscAsyncAction.LoadDocumentFonts);

            if (undefined !== this.asyncMethodCallback)
            	return;

			// заполним прогресс
			var _progress         = this.OpenDocumentProgress;
			_progress.Type        = c_oAscAsyncAction.LoadDocumentFonts;
			_progress.FontsCount  = this.FontLoader.fonts_loading.length;
			_progress.CurrentFont = 0;

			var _loader_object = this.WordControl.m_oLogicDocument;
			var _count         = 0;
			if (_loader_object !== undefined && _loader_object != null)
			{
				for (var i in _loader_object.ImageMap)
				{
					if (this.DocInfo.get_OfflineApp())
					{
						var localUrl = _loader_object.ImageMap[i];
						g_oDocumentUrls.addImageUrl(localUrl, this.documentUrl + 'media/' + localUrl);
					}
					++_count;
				}
			}

			_progress.ImagesCount  = _count;
			_progress.CurrentImage = 0;
		}
	};
	asc_docs_api.prototype.GenerateStyles                = function()
	{
		if (window["NATIVE_EDITOR_ENJINE"] === true)
		{
			if (!this.asc_checkNeedCallback("asc_onInitEditorStyles"))
				return;
		}
		
		if (!this.stylePainter)
			this.stylePainter = new AscCommonWord.CStylesPainter();
		
		this.stylePainter.GenerateStyles(this, (null == this.LoadedObject) ? this.WordControl.m_oLogicDocument.Get_Styles().Style : this.LoadedObjectDS);
	};
	asc_docs_api.prototype.asyncFontsDocumentEndLoaded   = function(blockType)
	{
		// все, шрифты загружены. Теперь нужно подгрузить картинки
		if (this.isPasteFonts_Images)
			this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadFont);
		else if (this.isSaveFonts_Images)
			this.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadFont);
		else
			this.sync_EndAction(undefined === blockType ? c_oAscAsyncActionType.BlockInteraction : blockType, c_oAscAsyncAction.LoadDocumentFonts);

        if (undefined !== this.asyncMethodCallback)
        {
            this.asyncMethodCallback();
            this.asyncMethodCallback = undefined;
            return;
        }

		this.EndActionLoadImages = 0;
		if (this.isPasteFonts_Images)
		{
			var _count = 0;
			for (var i in this.pasteImageMap)
				++_count;

			if (_count > 0)
			{
				this.EndActionLoadImages = 2;
				this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadImage);
			}

			var _oldAsyncLoadImages                     = this.ImageLoader.bIsAsyncLoadDocumentImages;
			this.ImageLoader.bIsAsyncLoadDocumentImages = false;
			this.ImageLoader.LoadDocumentImages(this.pasteImageMap);
			this.ImageLoader.bIsAsyncLoadDocumentImages = true;
			return;
		}
		else if (this.isSaveFonts_Images)
		{
			var _count = 0;
			for (var i in this.saveImageMap)
				++_count;

			if (_count > 0)
			{
				this.EndActionLoadImages = 2;
				this.sync_StartAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadImage);
			}

			this.ImageLoader.LoadDocumentImages(this.saveImageMap);
			return;
		}

		this.GenerateStyles();

		if (null != this.WordControl.m_oLogicDocument)
		{
			this.WordControl.m_oDrawingDocument.CheckGuiControlColors();
			this.sendColorThemes(this.WordControl.m_oLogicDocument.theme);
			this.sendEvent("asc_onUpdateChartStyles");
		}

		// открытие после загрузки документа

		var _loader_object = this.WordControl.m_oLogicDocument;
		if (null == _loader_object)
			_loader_object = this.WordControl.m_oDrawingDocument.m_oDocumentRenderer;

		var _count = 0;
		for (var i in _loader_object.ImageMap)
			++_count;

		if (!this.isOnlyReaderMode && !this.isPdfEditor())
		{
			// add const textures
			var _st_count = AscCommon.g_oUserTexturePresets.length;
			for (var i = 0; i < _st_count; i++)
				_loader_object.ImageMap[_count + i] = AscCommon.g_oUserTexturePresets[i];

			if (this.OpenDocumentProgress && !this.ImageLoader.bIsAsyncLoadDocumentImages)
			{
				this.OpenDocumentProgress.ImagesCount += _st_count;
			}
		}

		if (_count > 0)
		{
			this.EndActionLoadImages = 1;
			this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadDocumentImages);
		}

		this.ImageLoader.bIsLoadDocumentFirst = true;
		this.ImageLoader.LoadDocumentImages(_loader_object.ImageMap);
	};

	asc_docs_api.prototype.CreateFontsCharMap = function()
	{
		var _info = new CFontsCharMap();
		_info.StartWork();

		this.WordControl.m_oLogicDocument.Document_CreateFontCharMap(_info);

		return _info.EndWork();
	};

	asc_docs_api.prototype.sync_SendThemeColors       = function(colors, standart_colors)
	{
		this._gui_control_colors = {Colors : colors, StandartColors : standart_colors};
		this.sendEvent("asc_onSendThemeColors", colors, standart_colors);
	};


	asc_docs_api.prototype.getCurrentTheme = function ()
	{
		if (null == this.WordControl.m_oLogicDocument)
			return null;

		return this.WordControl.m_oLogicDocument.theme;
	};

	asc_docs_api.prototype.getGraphicController = function () {
		var document = this.private_GetLogicDocument();
		return document && document.DrawingsController && document.DrawingsController.DrawingObjects;
	};

	asc_docs_api.prototype.ChangeColorScheme            = function(sSchemeName)
	{
		if (null == this.WordControl.m_oLogicDocument)
			return;

		var _changer = this.WordControl.m_oLogicDocument.DrawingObjects;
		if (null == _changer)
			return;

		var theme = this.WordControl.m_oLogicDocument.theme;
		var scheme = AscCommon.getColorSchemeByName(sSchemeName);
		if (!scheme)
		{
			scheme = theme.getExtraClrScheme(sSchemeName);
		}
		if(!scheme)
		{
			return;
		}
		if (this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_ColorScheme) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_ChangeColorScheme);
			theme.changeColorScheme(scheme);
			this.WordControl.m_oDrawingDocument.CheckGuiControlColors();
			this.chartPreviewManager.clearPreviews();
			this.textArtPreviewManager.clear();
			this.sendEvent("asc_onUpdateChartStyles");
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.FinalizeAction();


			// TODO:
			this.WordControl.m_oDrawingDocument.ClearCachePages();
			this.WordControl.OnScroll();

			this.WordControl.m_oDrawingDocument.CheckGuiControlColors();
			this.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
		}

	};
	asc_docs_api.prototype.asc_ChangeColorSchemeByIdx            = function(nIdx)
	{
		if (null == this.WordControl.m_oLogicDocument)
			return;

		var _changer = this.WordControl.m_oLogicDocument.DrawingObjects;
		if (null == _changer)
			return;

		var theme = this.getCurrentTheme();
		if(!theme)
		{
			return;
		}
		var scheme = this.getColorSchemeByIdx(nIdx);
		if(!scheme)
		{
			return;
		}
		if (this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_ColorScheme) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_ChangeColorScheme);
			theme.changeColorScheme(scheme);
			this.WordControl.m_oDrawingDocument.CheckGuiControlColors();
			this.chartPreviewManager.clearPreviews();
			this.textArtPreviewManager.clear();
			this.sendEvent("asc_onUpdateChartStyles");
			this.WordControl.m_oLogicDocument.Recalculate();
			this.WordControl.m_oLogicDocument.FinalizeAction();


			// TODO:
			this.WordControl.m_oDrawingDocument.ClearCachePages();
			this.WordControl.OnScroll();

			this.WordControl.m_oDrawingDocument.CheckGuiControlColors();
			this.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
		}

	};
	asc_docs_api.prototype.asyncImagesDocumentEndLoaded = function()
	{
		this.ImageLoader.bIsLoadDocumentFirst = false;
		var _bIsOldPaste                      = this.isPasteFonts_Images;

		if (null != this.WordControl.m_oDrawingDocument.m_oDocumentRenderer)
		{
			if (this.EndActionLoadImages == 1)
			{
				this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadDocumentImages);
			}
			else if (this.EndActionLoadImages == 2)
			{
				if (this.isPasteFonts_Images)
					this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadImage);
				else
					this.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadImage);
			}
			this.EndActionLoadImages = 0;

			this.WordControl.m_oDrawingDocument.OpenDocument();

			this.LoadedObject = null;

			this.bInit_word_control = true;

			if (false === this.isPasteFonts_Images)
				this.onDocumentContentReady();

			this.WordControl.InitControl();

			if (this.isViewMode)
				this.asc_setViewMode(true);
			return;
		}

		// на методе _openDocumentEndCallback может поменяться this.EndActionLoadImages
		if (this.EndActionLoadImages == 1)
		{
			this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadDocumentImages);
		}
		else if (this.EndActionLoadImages == 2)
		{
			if (_bIsOldPaste)
				this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadImage);
			else
				this.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadImage);
		}
		this.EndActionLoadImages = 0;

		// размораживаем меню... и начинаем считать документ
		if (false === this.isPasteFonts_Images && false === this.isSaveFonts_Images && false === this.isLoadImagesCustom)
		{
			this.ServerImagesWaitComplete = true;
			this._openDocumentEndCallback();
		}
		else
		{
			if (this.isPasteFonts_Images)
			{
				this.isPasteFonts_Images = false;
				this.pasteImageMap       = null;
				this.pasteCallback();
				this.pasteCallback            = null;
				this.decrementCounterLongAction();
			}
			else if (this.isSaveFonts_Images)
			{
				this.isSaveFonts_Images = false;
				this.saveImageMap       = null;
				this.pre_SaveCallback();

				if (this.bInit_word_control === false)
				{
					this.bInit_word_control = true;
					this.onDocumentContentReady();
				}
			}
			else if (this.isLoadImagesCustom)
			{
				this.isLoadImagesCustom = false;
				this.loadCustomImageMap = null;

				if (!this.ImageLoader.bIsAsyncLoadDocumentImages)
					this.SyncLoadImages_callback();
			}
		}
	};

	asc_docs_api.prototype._openDocumentEndCallback = function()
	{
		if (this.isDocumentLoadComplete || !this.ServerImagesWaitComplete || !this.ServerIdWaitComplete || !this.WordControl || !this.WordControl.m_oLogicDocument)
			return;

		var isSendOnReady = false;

		if (0 == this.DocumentType)
			this.WordControl.m_oLogicDocument.LoadEmptyDocument();
		else if (1 == this.DocumentType)
		{
			this.WordControl.m_oLogicDocument.LoadTestDocument();
		}
		else
		{
			if (this.LoadedObject)
			{
				var Document = this.WordControl.m_oLogicDocument;

				if (this.isApplyChangesOnOpenEnabled)
				{
					if (AscCommon.EncryptionWorker)
					{
						AscCommon.EncryptionWorker.init();
						if (!AscCommon.EncryptionWorker.isChangesHandled)
							return AscCommon.EncryptionWorker.handleChanges(AscCommon.CollaborativeEditing.m_aChanges, this, this._openDocumentEndCallback);
					}

                    if (false === this.isSaveFonts_Images && !isSendOnReady)
                    {
                        isSendOnReady = true;
                        this.bInit_word_control = true;
                        Document.Start_SilentMode();
                    }

					this.isApplyChangesOnOpenEnabled = false;
					this._applyPreOpenLocks();

					// TODO: onDocumentContentReady вызываем в конце загрузки всех изменений (и объектов для этих изменений)
					let oThis = this;

					let perfStart = performance.now();
					let OtherChanges = AscCommon.CollaborativeEditing.Have_OtherChanges();
					AscCommon.CollaborativeEditing.Apply_Changes(function()
					{
						let perfEnd = performance.now();
						if (OtherChanges) {
							AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onApplyChanges", perfEnd - perfStart), oThis);
						}
						Document.MoveCursorToStartOfDocument();

						if (isSendOnReady)
						{
							oThis.onDocumentContentReady();
							Document.End_SilentMode(false);
						}
					});
					AscCommon.CollaborativeEditing.Release_Locks();

					this.isApplyChangesOnOpen = true;
				}
				else
				{
					if (false === this.isSaveFonts_Images && !isSendOnReady)
					{
						isSendOnReady           = true;
						this.bInit_word_control = true;
						Document.Start_SilentMode();
					}

					Document.MoveCursorToStartOfDocument();

					if (isSendOnReady)
					{
						this.onDocumentContentReady();
						Document.End_SilentMode(false);
					}
				}

				if (!this.isOnlyReaderMode)
				{
					if (false === this.isSaveFonts_Images)
						Document.RecalculateFromStart();

					this.WordControl.m_oDrawingDocument.TargetStart();
				}
				else
				{
					Document.RecalculateAllTables();
					Document.DrawingObjects.recalculate();

					if (!this.WordControl.IsReaderMode())
						this.ChangeReaderMode();
					else
						this.WordControl.UpdateReaderContent();
				}
			}
		}

		if (false === this.isSaveFonts_Images && !isSendOnReady)
		{
            isSendOnReady = true;
			this.bInit_word_control = true;
			this.onDocumentContentReady();
		}

		this.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
		//this.WordControl.m_oLogicDocument.Document_UpdateRulersState();
		this.WordControl.m_oLogicDocument.Document_UpdateSelectionState();
		this.LoadedObject = null;

		this.WordControl.InitControl();

		if (!this.isViewMode)
		{
			this.sendStandartTextures();
			this.sendMathToMenu();

			if (this.shapeElementId)
			{
				this.WordControl.m_oDrawingDocument.InitGuiCanvasShape(this.shapeElementId);
			}
		}

		if (this.isViewMode)
			this.asc_setViewMode(true);

		// Меняем тип состояния (на никакое)
		this.advancedOptionsAction = c_oAscAdvancedOptionsAction.None;
	};

	asc_docs_api.prototype.UpdateInterfaceState = function()
	{
		if (this.WordControl.m_oLogicDocument != null)
			this.WordControl.m_oLogicDocument.UpdateInterface();
	};

	asc_docs_api.prototype.asyncFontEndLoaded = function(fontinfo)
	{
		this.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadFont);

		if (undefined !== this.asyncMethodCallback)
		{
			this.asyncMethodCallback();
			this.asyncMethodCallback = undefined;
			return;
		}

		var _fontSelections = g_fontApplication.g_fontSelections;
		if (_fontSelections.CurrentLoadedObj != null)
		{
			var _rfonts = _fontSelections.getSetupRFonts(_fontSelections.CurrentLoadedObj);
			this.WordControl.m_oLogicDocument.TextBox_Put(_fontSelections.CurrentLoadedObj.text, _rfonts);
			this.WordControl.ReinitTB();

			_fontSelections.CurrentLoadedObj = null;
			this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadFont);
			return;
		}

		if (this.FontAsyncLoadType == 1)
		{
			this.FontAsyncLoadType = 0;
			this.asc_AddMath2(this.FontAsyncLoadParam);
			this.FontAsyncLoadParam = null;
			return;
		}

		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextFontNameLong);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({
				FontFamily : {
					Name  : fontinfo.Name,
					Index : -1
				}
			}));
			this.WordControl.m_oLogicDocument.UpdateInterface();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
		// отжать заморозку меню
	};

	asc_docs_api.prototype.asc_replaceLoadImageCallback = function(fCallback)
	{
		this.asyncImageEndLoaded2 = fCallback;
	};

	asc_docs_api.prototype.asyncImageEndLoaded = function(_image)
	{
		// отжать заморозку меню
		if (this.asyncImageEndLoaded2)
			this.asyncImageEndLoaded2(_image);
		else
		{
			if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
			{
				this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddImage);
				this.WordControl.m_oLogicDocument.AddInlineImage(50, 50, _image.src);
				this.WordControl.m_oLogicDocument.FinalizeAction();
			}
		}
	};

	asc_docs_api.prototype.openDocument = function(file)
	{
		let perfStart = performance.now();
		if (file.changes && this.VersionHistory)
		{
			this.VersionHistory.changes = file.changes;
			this.VersionHistory.applyChanges(this);
		}

		if (file.bSerFormat)
		{
			this.OpenDocumentFromBin(file.url, file.data);
		}
		else
		{
			this.isOpenOOXInBrowser = this["asc_isSupportFeature"]("ooxml") && AscCommon.checkOOXMLSignature(file.data);
			if (this.isOpenOOXInBrowser)
			{
				this.openOOXInBrowserZip = file.data;
				this.OpenDocumentFromZip(file.data);
			}
			else
				this.OpenDocument(file.url, file.data);
		}
		let perfEnd = performance.now();
		AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onOpenDocument", perfEnd - perfStart), this);
	};

	asc_docs_api.prototype.asyncImageEndLoadedBackground = function(_image)
	{
		this.WordControl.m_oDrawingDocument.CheckRasterImageOnScreen(_image.src);
		if (this.WordControl.m_oDrawingDocument.LastDrawingUrl == _image.src)
		{
			this.WordControl.m_oDrawingDocument.LastDrawingUrl = "";
			this.UpdateInterfaceState();
		}
	};
	asc_docs_api.prototype.IsAsyncOpenDocumentImages     = function()
	{
		return true;
	};

	asc_docs_api.prototype.pre_Paste = function(_fonts, _images, callback)
	{
		if (undefined !== window["native"] && undefined !== window["native"]["GetImageUrl"])
		{
			callback();
			return;
		}
		if(window['IS_NATIVE_EDITOR'])
		{
			callback();
			return;
		}
		this.pasteCallback = callback;
		this.pasteImageMap = _images;

		var _count = 0;
		for (var i in this.pasteImageMap)
			++_count;

        AscFonts.FontPickerByCharacter.extendFonts(_fonts);
		if (0 == _count && false === this.FontLoader.CheckFontsNeedLoading(_fonts))
		{
			// никаких евентов. ничего грузить не нужно. сделано для сафари под макОс.
			// там при LongActions теряется фокус и вставляются пробелы
			this.pasteCallback();
			this.pasteCallback            = null;

			return;
		}

		this.incrementCounterLongAction();
		this.isPasteFonts_Images = true;
		this.FontLoader.LoadDocumentFonts2(_fonts);
	};

	asc_docs_api.prototype.pre_Save = function(_images)
	{
		this.isSaveFonts_Images = true;
		this.saveImageMap       = _images;
		this.WordControl.m_oDrawingDocument.CheckFontNeeds();
		this.FontLoader.LoadDocumentFonts2(this.WordControl.m_oLogicDocument.Fonts);
	};

	asc_docs_api.prototype.SyncLoadImages          = function(_images)
	{
		this.isLoadImagesCustom = true;
		this.loadCustomImageMap = _images;

		var _count  = 0;
		var _loaded = this.ImageLoader.map_image_index;

		var _new_len = this.loadCustomImageMap.length;
		for (var i = 0; i < _new_len; i++)
		{
			if (undefined !== _loaded[this.loadCustomImageMap[i]])
			{
				this.loadCustomImageMap.splice(i, 1);
				i--;
				_new_len--;
				continue;
			}
			++_count;
		}

		if (_count > 0)
		{
			this.EndActionLoadImages = 2;
			this.sync_StartAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadImage);
		}

		this.ImageLoader.LoadDocumentImages(this.loadCustomImageMap);
	};
	asc_docs_api.prototype.SyncLoadImages_callback = function()
	{
		this.WordControl.OnRePaintAttack();
	};

	asc_docs_api.prototype.pre_SaveCallback = function()
	{
		AscCommon.CollaborativeEditing.OnEnd_Load_Objects();

		if (this.isApplyChangesOnOpen)
		{
			this.isApplyChangesOnOpen = false;
			this._openDocumentEndCallback();
		}
		if (this.isApplyChangesOnVersionHistory)
		{
			this.isApplyChangesOnVersionHistory = false;
			this._openVersionHistoryEndCallback();
		}
	};

	asc_docs_api.prototype.ViewScrollToX             = function(x)
	{
		this.WordControl.m_oScrollHorApi.scrollToX(x);
	};
	asc_docs_api.prototype.ViewScrollToY             = function(y)
	{
		this.WordControl.m_oScrollVerApi.scrollToY(y);
	};
	asc_docs_api.prototype.GetDocWidthPx             = function()
	{
		return this.WordControl.m_dDocumentWidth;
	};
	asc_docs_api.prototype.GetDocHeightPx            = function()
	{
		return this.WordControl.m_dDocumentHeight;
	};
	asc_docs_api.prototype.ClearSearch               = function()
	{
		return this.WordControl.m_oDrawingDocument.EndSearch(true);
	};
	asc_docs_api.prototype.GetCurrentVisiblePage     = function()
	{
		if (window["IS_NATIVE_EDITOR"])
			return window["native"]["GetDrawingEndPage"]();

		var lPage1 = this.WordControl.m_oDrawingDocument.m_lDrawingFirst;
		var lPage2 = lPage1 + 1;

		if (lPage2 > this.WordControl.m_oDrawingDocument.m_lDrawingEnd)
			return lPage1;

		var lWindHeight = this.WordControl.m_oEditor.HtmlElement.height;
		var arPages     = this.WordControl.m_oDrawingDocument.m_arrPages;

		var dist1 = arPages[lPage1].drawingPage.bottom;
		var dist2 = lWindHeight - arPages[lPage2].drawingPage.top;

		if (dist1 > dist2)
			return lPage1;

		return lPage2;
	};
	asc_docs_api.prototype.GetCurrentVisiblePages = function() {
		let start, end;
		
		if (window["IS_NATIVE_EDITOR"]) {
			start = window["native"]["GetDrawingStartPage"]();
			end = window["native"]["GetDrawingEndPage"]();
		} else {
			start = this.WordControl.m_oDrawingDocument.m_lDrawingFirst;
			end = this.WordControl.m_oDrawingDocument.m_lDrawingEnd;
		}
		
		if (start == -1 || end == -1) {
			return [];
		}
		
		const pages = [];
		for (let i = start; i <= end; i++) {
			pages.push(i);
		}
	
		return pages;
	};

	asc_docs_api.prototype.asc_SetDocumentPlaceChangedEnabled = function(bEnabled)
	{
		if (this.WordControl)
			this.WordControl.m_bDocumentPlaceChangedEnabled = bEnabled;
	};

	asc_docs_api.prototype.asc_SetViewRulers       = function(bRulers)
	{
		//if (false === this.bInit_word_control || true === this.isViewMode)
		//    return;

		if (!this.isLoadFullApi)
		{
			this.tmpViewRulers = bRulers;
			return;
		}

		if (this.WordControl.m_bIsRuler != bRulers)
		{
			this.WordControl.m_bIsRuler = bRulers;
			this.WordControl.checkNeedRules();
			this.WordControl.OnResize(true);
		}
	};
	asc_docs_api.prototype.asc_SetViewRulersChange = function()
	{
		//if (false === this.bInit_word_control || true === this.isViewMode)
		//    return;

		this.WordControl.m_bIsRuler = !this.WordControl.m_bIsRuler;
		this.WordControl.checkNeedRules();
		this.WordControl.OnResize(true);
		return this.WordControl.m_bIsRuler;
	};
	asc_docs_api.prototype.asc_GetViewRulers       = function()
	{
		return this.WordControl.m_bIsRuler;
	};

	asc_docs_api.prototype.asc_SetDocumentUnits = function(_units)
	{
		if (this.WordControl && this.WordControl.m_oHorRuler && this.WordControl.m_oVerRuler)
		{
			this.WordControl.m_oHorRuler.Units = _units;
			this.WordControl.m_oVerRuler.Units = _units;
			this.WordControl.UpdateHorRulerBack(true);
			this.WordControl.UpdateVerRulerBack(true);
		}
		else
		{
            this.tmpDocumentUnits = _units;
		}
	};

	asc_docs_api.prototype.GoToHeader = function(pageNumber)
	{
		if (this.WordControl.m_oDrawingDocument.IsFreezePage(pageNumber))
			return;

		var bForceRedraw  = false;
		var LogicDocument = this.WordControl.m_oLogicDocument;
		if (AscCommonWord.docpostype_HdrFtr !== LogicDocument.GetDocPosType())
		{
			LogicDocument.SetDocPosType(AscCommonWord.docpostype_HdrFtr);
			bForceRedraw = true;
		}

		var oldClickCount            = global_mouseEvent.ClickCount;
		global_mouseEvent.Button     = 0;
		global_mouseEvent.ClickCount = 1;

		LogicDocument.OnMouseDown(global_mouseEvent, 0, 0, pageNumber);
		LogicDocument.OnMouseUp(global_mouseEvent, 0, 0, pageNumber);
		LogicDocument.OnMouseMove(global_mouseEvent, 0, 0, pageNumber);
		LogicDocument.MoveCursorLeft();
		LogicDocument.UpdateInterface();
		LogicDocument.UpdateSelection();

		global_mouseEvent.ClickCount = oldClickCount;

		if (true === bForceRedraw)
		{
			this.WordControl.m_oDrawingDocument.ClearCachePages();
			this.WordControl.m_oDrawingDocument.FirePaint();
		}
	};

	asc_docs_api.prototype.GoToFooter = function(pageNumber)
	{
		if (this.WordControl.m_oDrawingDocument.IsFreezePage(pageNumber))
			return;

		var bForceRedraw  = false;
		var LogicDocument = this.WordControl.m_oLogicDocument;
		if (AscCommonWord.docpostype_HdrFtr !== LogicDocument.GetDocPosType())
		{
			LogicDocument.SetDocPosType(AscCommonWord.docpostype_HdrFtr);
			bForceRedraw = true;
		}

		var oldClickCount            = global_mouseEvent.ClickCount;
		global_mouseEvent.Button     = 0;
		global_mouseEvent.ClickCount = 1;

		LogicDocument.OnMouseDown(global_mouseEvent, 0, AscCommon.Page_Height, pageNumber);
		LogicDocument.OnMouseUp(global_mouseEvent, 0, AscCommon.Page_Height, pageNumber);
		LogicDocument.OnMouseMove(global_mouseEvent, 0, 0, pageNumber);
		LogicDocument.MoveCursorLeft();
		LogicDocument.UpdateInterface();
		LogicDocument.UpdateSelection();

		global_mouseEvent.ClickCount = oldClickCount;

		if (true === bForceRedraw)
		{
			this.WordControl.m_oDrawingDocument.ClearCachePages();
			this.WordControl.m_oDrawingDocument.FirePaint();
		}
	};

	asc_docs_api.prototype.ExitHeader_Footer = function(pageNumber)
	{
		if (this.WordControl.m_oDrawingDocument.IsFreezePage(pageNumber))
			return;

		var oldClickCount            = global_mouseEvent.ClickCount;
		global_mouseEvent.ClickCount = 2;
		this.WordControl.m_oLogicDocument.OnMouseDown(global_mouseEvent, 0, AscCommon.Page_Height / 2, pageNumber);
		this.WordControl.m_oLogicDocument.OnMouseUp(global_mouseEvent, 0, AscCommon.Page_Height / 2, pageNumber);

		this.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();

		global_mouseEvent.ClickCount = oldClickCount;
	};

	asc_docs_api.prototype.asc_RemoveHeader = function(pageNumber)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		oLogicDocument.RemoveHdrFtr(pageNumber, true);
	};

	asc_docs_api.prototype.asc_RemoveFooter = function(pageNumber)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		oLogicDocument.RemoveHdrFtr(pageNumber, false);
	};

	asc_docs_api.prototype.GetCurrentPixOffsetY = function()
	{
		return this.WordControl.m_dScrollY;
	};

	asc_docs_api.prototype.retrieveFormatPainterData = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
		{
			return null;
		}
		return oLogicDocument.GetFormattingPasteData();
	};

	asc_docs_api.prototype.SetPaintFormat = function(_value)
	{
		const oLogicDocument = this.WordControl.m_oLogicDocument;
		if(!oLogicDocument)
			return;
		if (this.isDrawTablePen || this.isDrawTableErase)
		{
			this.isDrawTablePen && this.sync_TableDrawModeCallback(false);
			this.isDrawTableErase && this.sync_TableEraseModeCallback(false);
			oLogicDocument.UpdateCursorType(oLogicDocument.CurPos.RealX, oLogicDocument.CurPos.RealY, oLogicDocument.CurPage, new AscCommon.CMouseEventHandler());
		}
		else if (true === this.isMarkerFormat)
		{
			this.sync_MarkerFormatCallback(false);
			oLogicDocument.UpdateCursorType(oLogicDocument.CurPos.RealX, oLogicDocument.CurPos.RealY, oLogicDocument.CurPage, new AscCommon.CMouseEventHandler());
		}
		else if (this.isStartAddShape)
		{
			this.sync_StartAddShapeCallback(false);
			this.sync_EndAddShape();
			oLogicDocument.DrawingObjects.endTrackNewShape();
			oLogicDocument.UpdateCursorType(oLogicDocument.CurPos.RealX, oLogicDocument.CurPos.RealY, oLogicDocument.CurPage, new AscCommon.CMouseEventHandler());
		}
		this.stopInkDrawer();
		this.cancelEyedropper();


		var value = ( true === _value ? c_oAscFormatPainterState.kOn : ( false === _value ? c_oAscFormatPainterState.kOff : _value ) );

		this.formatPainter.putState(value);

		if (c_oAscFormatPainterState.kOff !== value)
			oLogicDocument.Document_Format_Copy();
	};

	asc_docs_api.prototype.ChangeShapeType = function(value)
	{
		this.ImgApply(new asc_CImgProperty({ShapeProperties : {type : value}}));
	};

	asc_docs_api.prototype.sync_PaintFormatCallback = function(_value)
	{
		var value = ( true === _value ? c_oAscFormatPainterState.kOn : ( false === _value ? c_oAscFormatPainterState.kOff : _value ) );

		this.formatPainter.putState(value);
		return this.sendEvent("asc_onPaintFormatChanged", value);
	};
	asc_docs_api.prototype.SetMarkerFormat          = function(value, is_flag, r, g, b)
	{
		this.isMarkerFormat = value;

		if (this.isMarkerFormat)
		{
			this.WordControl.m_oLogicDocument.SetParagraphHighlight(is_flag, r, g, b);
			this.WordControl.m_oLogicDocument.Document_Format_Copy();
		}
	};

    asc_docs_api.prototype.sync_MarkerFormatCallback = function(value)
    {
        this.isMarkerFormat = value;
        return this.sendEvent("asc_onMarkerFormatChanged", value);
    };
	
	asc_docs_api.prototype.asc_SendForm = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		let oform = logicDocument ? logicDocument.GetOFormDocument() : null;
		if (!logicDocument || !oform)
			return false;
		
		if (!oform.getCurrentRole())
			return this._sendForm();
		
		if (!oform.canFillRole(oform.getCurrentRole()))
			return false;
		
		this.forceSaveSendFormRequest = true;
		this.asc_Save();
		
		return true;
	};
	asc_docs_api.prototype.setCurrentRoleFilled = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		let oform = logicDocument.GetOFormDocument();
		
		logicDocument.StartAction(AscDFH.historydescription_OForm_RoleFilled);
		oform.setRoleFilled(oform.getCurrentRole(), true);
		logicDocument.FinalizeAction();
	};
	asc_docs_api.prototype._sendForm = function()
	{
		var t = this;
		this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.Submit);
		var data = {'type': 'sendForm', 'userconnectionid': this.CoAuthoringApi.getUserConnectionId(), 'formdata': {
				'formsdata': this.asc_GetAllFormsData(true)
		}};
		this.saveFromChanges(data, Asc.c_nMaxConversionTime, function(isTimeout, response) {
			if (!(response && AscCommon.c_oAscServerCommandErrors.NoError === response.code)) {
				t.sendEvent('asc_onError', Asc.c_oAscError.ID.Submit, c_oAscError.Level.NoCritical);
			}
			t.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.Submit);
		});
		
		this.sendEvent("onSubmitForm");
		
		if (window.g_asc_plugins)
			window.g_asc_plugins.onPluginEvent("onSubmitForm");
		
		return true;
	};

    asc_docs_api.prototype.SetTableDrawMode = function(value)
    {
    	if (!this.WordControl || !this.WordControl.m_oLogicDocument)
    		return;

        this.isDrawTablePen = value;
        this.WordControl.m_oLogicDocument.DrawTableMode.Draw = value;

        if (this.isDrawTablePen && this.isDrawTableErase)
			this.SetTableEraseMode(false);

        this.WordControl.m_oDrawingDocument.UnlockCursorType();
        if (this.isDrawTablePen)
            this.WordControl.m_oDrawingDocument.LockCursorType(AscCommon.Cursors.TablePen);
    };
    asc_docs_api.prototype.sync_TableDrawModeCallback = function(value)
    {
        this.isDrawTablePen = value;
        this.WordControl.m_oLogicDocument.DrawTableMode.Draw = value;
        if (!this.isDrawTablePen)
		{
			this.WordControl.m_oLogicDocument.DrawTableMode.Start = false;
			this.WordControl.m_oDrawingDocument.UnlockCursorType();
		}

        return this.sendEvent("asc_onTableDrawModeChanged", value);
    };
    asc_docs_api.prototype.SetTableEraseMode = function(value)
    {
        if (!this.WordControl || !this.WordControl.m_oLogicDocument)
            return;

        this.isDrawTableErase = value;
        this.WordControl.m_oLogicDocument.DrawTableMode.Erase = value;

        if (this.isDrawTableErase && this.isDrawTablePen)
            this.SetTableDrawMode(false);

        this.WordControl.m_oDrawingDocument.UnlockCursorType();
        if (this.isDrawTableErase)
            this.WordControl.m_oDrawingDocument.LockCursorType(AscCommon.Cursors.TableEraser);
    };
    asc_docs_api.prototype.sync_TableEraseModeCallback = function(value)
    {
        this.isDrawTableErase = value;
        this.WordControl.m_oLogicDocument.DrawTableMode.Erase = value;
        if (!this.isDrawTableErase)
		{
			this.WordControl.m_oLogicDocument.DrawTableMode.Start = false;
			this.WordControl.m_oDrawingDocument.UnlockCursorType();
		}

        return this.sendEvent("asc_onTableEraseModeChanged", value);
    };

	asc_docs_api.prototype.StartAddShape = function(sPreset, is_apply)
	{
		if (this.isDrawTablePen)
		{
			this.sync_TableDrawModeCallback(false);
        }
        if (this.isDrawTableErase)
		{
            this.sync_TableEraseModeCallback(false);
        }

		this.stopInkDrawer();
		this.cancelEyedropper();

		this.isStartAddShape = true;
		this.addShapePreset  = sPreset;
		if (is_apply)
		{
			this.WordControl.m_oDrawingDocument.LockCursorType("crosshair");
		}
		else
		{
			editor.sync_EndAddShape();
			editor.sync_StartAddShapeCallback(false);
		}
	};

	asc_docs_api.prototype.AddShapeOnCurrentPage = function(_type)
	{
		if (!this.WordControl.m_oLogicDocument)
			return;
		this.WordControl.m_oLogicDocument.DrawingObjects.addShapeOnPage(_type, this.GetCurrentVisiblePage());
	};


	asc_docs_api.prototype.asc_canEditCrop = function()
	{
		return this.WordControl.m_oLogicDocument.DrawingObjects.canStartImageCrop();
	};

	asc_docs_api.prototype.asc_startEditCrop = function()
	{
		return this.WordControl.m_oLogicDocument.DrawingObjects.startImageCrop();
	};

	asc_docs_api.prototype.asc_endEditCrop = function()
	{
		return this.WordControl.m_oLogicDocument.DrawingObjects.endImageCrop();
	};

	asc_docs_api.prototype.asc_cropFit = function()
	{
		return this.WordControl.m_oLogicDocument.DrawingObjects.cropFit();
	};

	asc_docs_api.prototype.asc_cropFill = function()
	{
		return this.WordControl.m_oLogicDocument.DrawingObjects.cropFill();
	};

	asc_docs_api.prototype.AddTextArt = function(nStyle)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddTextArt);
			this.WordControl.m_oLogicDocument.AddTextArt(nStyle);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.onInkDrawerChangeState = function() {
		const oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;
		if(this.isDrawTablePen)
		{
			this.sync_TableDrawModeCallback(false);
		}
		if (this.isDrawTableErase)
		{
			this.sync_TableEraseModeCallback(false);
		}
		if (this.isMarkerFormat)
		{
			this.sync_MarkerFormatCallback(false);
		}
		if (this.isFormatPainterOn())
		{
			this.sync_PaintFormatCallback(c_oAscFormatPainterState.kOff);
		}
		if (this.isStartAddShape)
		{
			this.sync_StartAddShapeCallback(false);
			this.sync_EndAddShape();
			oLogicDocument.DrawingObjects.endTrackNewShape();
		}
		if(this.isEyedropperStarted())
		{
			this.cancelEyedropper();
		}
		oLogicDocument.DrawingObjects.onInkDrawerChangeState();
		oLogicDocument.UpdateCursorType(oLogicDocument.CurPos.RealX, oLogicDocument.CurPos.RealY, oLogicDocument.CurPage, new AscCommon.CMouseEventHandler());

	};

	asc_docs_api.prototype.sync_StartAddShapeCallback = function(value)
	{
		this.isStartAddShape = value;
		return this.sendEvent("asc_onStartAddShapeChanged", value);
	};

	asc_docs_api.prototype.CanGroup = function()
	{
		return this.WordControl.m_oLogicDocument.CanGroup();
	};

	asc_docs_api.prototype.CanUnGroup = function()
	{
		return this.WordControl.m_oLogicDocument.CanUnGroup();
	};

	asc_docs_api.prototype.asc_canMergeSelectedShapes = function (operation) {
		return AscFormat.canMergeSelectedShapes(operation);
	};
	asc_docs_api.prototype.asc_mergeSelectedShapesAction = function (operation) {
		const isSelectionLocked = this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props);
		const canMerge = this.asc_canMergeSelectedShapes(operation);
		if (!isSelectionLocked && canMerge) {
			this.WordControl.m_oLogicDocument.StartAction(0);
			AscFormat.mergeSelectedShapes(operation);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.CanChangeWrapPolygon = function()
	{
		return this.WordControl.m_oLogicDocument.CanChangeWrapPolygon();
	};

	asc_docs_api.prototype.StartChangeWrapPolygon = function()
	{
		return this.WordControl.m_oLogicDocument.StartChangeWrapPolygon();
	};


	asc_docs_api.prototype.ClearFormating = function()
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_ClearFormatting);
			this.WordControl.m_oLogicDocument.ClearParagraphFormatting();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.GetSectionInfo = function()
	{
		var obj = new CAscSection();

		// TODO: Переделать данную функцию, если она вообще нужна
		obj.PageWidth  = 297;
		obj.PageHeight = 210;

		obj.MarginLeft   = 30;
		obj.MarginRight  = 15;
		obj.MarginTop    = 20;
		obj.MarginBottom = 20;

		return obj;
	};

	asc_docs_api.prototype.add_SectionBreak = function(_Type)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddSectionBreak);
			this.WordControl.m_oLogicDocument.Add_SectionBreak(_Type);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_GetSectionsCount = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return 1;

		return oLogicDocument.GetSectionsCount();
	};

	asc_docs_api.prototype.asc_setViewMode = function(isViewMode)
	{
		this.isViewMode = !!isViewMode;
		if (!this.isLoadFullApi)
		{
			return;
		}

		if (isViewMode)
		{
			this.asc_SpellCheckDisconnect();

			this.ShowParaMarks                           = false;
			AscCommon.CollaborativeEditing.Set_GlobalLock(true);
			//this.isShowTableEmptyLine = false;
			//this.WordControl.m_bIsRuler = true;

			if (null == this.WordControl.m_oDrawingDocument.m_oDocumentRenderer)
			{
				this.WordControl.m_oDrawingDocument.ClearCachePages();
				this.WordControl.m_oDrawingDocument.FirePaint();
				this.WordControl.HideRulers();
			}
			else
			{
				this.WordControl.HideRulers();
				this.WordControl.OnScroll();
			}
			this.turnOffSpecialModes();
		}
		else
		{
			//this.WordControl.m_bIsRuler = true;
			this.WordControl.checkNeedRules();
			this.WordControl.m_oDrawingDocument.ClearCachePages();
			this.WordControl.OnResize(true);
		}

		this.checkInputMode();
	};

	asc_docs_api.prototype.OnMouseUp = function(x, y)
	{
		this.WordControl.onMouseUpExternal(x, y);
	};

	asc_docs_api.prototype.asyncImageEndLoaded2       = null;

	asc_docs_api.prototype.SetDrawImagePlaceParagraph = function(element_id, props)
	{
		this.WordControl.m_oDrawingDocument.InitGuiCanvasTextProps(element_id);
		this.WordControl.m_oDrawingDocument.DrawGuiCanvasTextProps(props);
	};

	asc_docs_api.prototype.asc_getMasterCommentId = function()
	{
		return -1;
	};

	asc_docs_api.prototype.asc_getAnchorPosition = function()
	{
		var AnchorPos = this.WordControl.m_oLogicDocument.GetSelectionAnchorPos();
		return new AscCommon.asc_CRect(AnchorPos.X0, AnchorPos.Y, AnchorPos.X1 - AnchorPos.X0, 0);
	};

	asc_docs_api.prototype._onNeedParams  = function(data, opt_isPassword)
	{
		var t = this;
		if (opt_isPassword) {
			if (this.asc_checkNeedCallback("asc_onAdvancedOptions")) {
				t.sendEvent("asc_onAdvancedOptions", c_oAscAdvancedOptionsID.DRM, data ? data.message : undefined);
			} else {
				t.sendEvent("asc_onError", c_oAscError.ID.ConvertationPassword, c_oAscError.Level.Critical);
			}
		} else {
			if (this.asc_checkNeedCallback("asc_onAdvancedOptions")) {
				var cp = {'codepage': AscCommon.c_oAscCodePageNone, 'encodings': AscCommon.getEncodingParams()};
				if (data && typeof Blob !== 'undefined' && typeof FileReader !== 'undefined') {
					AscCommon.loadFileContent(data, function(httpRequest) {
						if (httpRequest && httpRequest.response) {
							cp['data'] = httpRequest.response;
							t.sendEvent("asc_onAdvancedOptions", c_oAscAdvancedOptionsID.TXT, new AscCommon.asc_CAdvancedOptions(cp));
						} else {
							t.sendEvent("asc_onError", c_oAscError.ID.Unknown, c_oAscError.Level.Critical);
						}
					}, "arraybuffer");
				} else {
					t.sendEvent("asc_onAdvancedOptions", c_oAscAdvancedOptionsID.TXT, new AscCommon.asc_CAdvancedOptions(cp));
				}
			} else {
				this.asc_setAdvancedOptions(c_oAscAdvancedOptionsID.TXT, new Asc.asc_CTextOptions(AscCommon.c_oAscCodePageUtf8));
			}
		}
	};
	asc_docs_api.prototype._waitPrint    = function(actionType, options)
	{
		if (this.WordControl && this.WordControl.m_oDrawingDocument &&
			(c_oAscFileType.PDF === options.fileType || c_oAscFileType.PDFA === options.fileType))
		{
			return this.WordControl.m_oDrawingDocument.CheckPrint([actionType, options]);
		}
		return false;
	};
	asc_docs_api.prototype._downloadAs    = function(actionType, options, oAdditionalData, dataContainer, downloadType)
	{
		var t = this;
		var fileType = options.fileType;

		if (this.WordControl.m_oLogicDocument && this.isCloudSaveAsLocalToDrawingFormat(actionType, fileType))
		{
			if (this.isUseNativeViewer && this.isDocumentRenderer())
			{
				let changes = this.getDocumentRenderer().Save();
				window["AscDesktopEditor"]["localSaveToDrawingFormat2"](this.documentTitle, this.DocumentUrl || "", "",
					changes ? AscCommon.Base64.encode(changes) : "", AscCommon.Base64.encode(this.getDocumentRenderer().getFileNativeBinary()), this.currentPassword || "", fileType);
			}
			else
			{
				this.localSaveToDrawingFormat(this.WordControl.m_oDrawingDocument.ToRendererPart(false, options.isPdfPrint || ((fileType & 0x0400) === 0x0400)), fileType);
			}
			return true;
		}

		if (c_oAscAsyncAction.SendMailMerge === actionType)
		{
			oAdditionalData["c"] = 'sendmm';
			oAdditionalData["userindex"] = this.CoAuthoringApi.get_indexUser();
		}
		else if (this.isUseNativeViewer && this.isDocumentRenderer())
		{
			oAdditionalData["c"] = 'savefromorigin';
		}

		let canDownloadFromBytes = null == options.oDocumentMailMerge && null == options.oMailMergeSendData
			&& !window.isCloudCryptoDownloadAs && DownloadType.None === downloadType
			&& !(AscCommon.AscBrowser.isAppleDevices && AscCommon.AscBrowser.isChrome)
			&& !AscCommon.AscBrowser.isAndroidNativeApp;
		if ('savefromorigin' === oAdditionalData["c"])
		{
			oAdditionalData["format"] = this.documentFormat;
			if (options.pdfChanges) {
				dataContainer.data = options.pdfChanges;
			} else if (this.isUseNativeViewer && this.isDocumentRenderer()) {
				dataContainer.data = this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.Save();
			}
		}
		else if ((c_oAscFileType.PDF === fileType || c_oAscFileType.PDFA === fileType) &&
			null == options.oDocumentMailMerge && (options.isPdfPrint || !(this.DocInfo && this.DocInfo.isFormatWithForms())))
		{
            var isSelection = false;
            if (options.advancedOptions && options.advancedOptions && (Asc.c_oAscPrintType.Selection === options.advancedOptions.asc_getPrintType()))
                isSelection = true;

			var dd             = this.WordControl.m_oDrawingDocument;
			if (isSelection)
				dd.GenerateSelectionPrint();

			dataContainer.data = dd.ToRendererPart(oAdditionalData["nobase64"], options.isPdfPrint);
			//console.log(oAdditionalData["data"]);
		}
		else if (c_oAscFileType.JSON === fileType)
		{
			oAdditionalData['url'] = this.mailMergeFileData['url'];
			oAdditionalData['format'] = this.mailMergeFileData['fileType'];
			if (this.mailMergeFileData['token']) {
				oAdditionalData['tokenDownload'] = this.mailMergeFileData['token'];
				//remove to reduce message size
				oAdditionalData['tokenSession'] = undefined;
			}
			// ToDo select csv params
			oAdditionalData['codepage']  = AscCommon.c_oAscCodePageUtf8;
			oAdditionalData['delimiter'] = AscCommon.c_oAscCsvDelimiter.Comma;
			dataContainer.data = this.mailMergeFileData['data'];
		}
		else if (this.insertDocumentUrlsData)
		{
			var last = this.insertDocumentUrlsData.documents.shift();
			oAdditionalData['url'] = last.url;
			oAdditionalData['format'] = last.format;
			if (last.token) {
				oAdditionalData['tokenDownload'] = last.token;
				//remove to reduce message size
				oAdditionalData['tokenSession'] = undefined;
			}
			oAdditionalData['outputurls']= true;
			// ToDo select txt params
			oAdditionalData["codepage"] = AscCommon.c_oAscCodePageUtf8;
			dataContainer.data = last.data;
		}
		else if (c_oAscFileType.HTML === fileType && canDownloadFromBytes)
		{
			//DownloadFileFromBytes has bug on Chrome on iOS https://github.com/kennethjiang/js-file-download/issues/72
			//в asc_nativeGetHtml будет вызван select all, чтобы выделился документ должны выйти из колонтитулов и автофигур
			var _e     = new AscCommon.CKeyboardEvent();
			_e.CtrlKey = false;
			_e.KeyCode = 27;
			this.WordControl.m_oLogicDocument.OnKeyDown(_e);

			var data =  window["asc_docs_api"].prototype["asc_nativeGetHtml"].call(this);
			AscCommon.DownloadFileFromBytes(data, oAdditionalData["title"], "text/html", '\ufeff');

			if (actionType)
			{
				this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, actionType);
			}
			return true;
		}
		else if(this.isOpenOOXInBrowser && this["asc_isSupportFeature"]("ooxml"))
		{
			var title = this.documentTitle;
			this.saveLogicDocumentToZip(fileType, options, function(data) {
				if (data) {
					if (c_oAscFileType.DOCX === fileType && canDownloadFromBytes) {
						AscCommon.DownloadFileFromBytes(data, title, AscCommon.openXml.GetMimeType("docx"));
					} else {
						dataContainer.data = data;
						if (window.isCloudCryptoDownloadAs)
						{
							var sParamXml = ("<m_nCsvTxtEncoding>" + oAdditionalData["codepage"] + "</m_nCsvTxtEncoding>");
							window["AscDesktopEditor"]["CryptoDownloadAs"](dataContainer.data, fileType, sParamXml);
							return true;
						}
						t._downloadAsUsingServer(actionType, options, oAdditionalData, dataContainer, downloadType);
						return;
					}
				} else {
					t.sendEvent("asc_onError", Asc.c_oAscError.ID.Unknown, Asc.c_oAscError.Level.NoCritical);
				}
				if (actionType) {
					t.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, actionType);
				}
			});
			return true;
		} else {
			if (options.advancedOptions instanceof Asc.asc_CTextOptions)
			{
				oAdditionalData["codepage"] = options.advancedOptions.asc_getCodePage();
			}
			var oLogicDocument;
			if (null != options.oDocumentMailMerge)
				oLogicDocument = options.oDocumentMailMerge;
			else
				oLogicDocument = this.WordControl.m_oLogicDocument;
			var oBinaryFileWriter;
			if (null != options.oMailMergeSendData && c_oAscFileType.HTML === options.oMailMergeSendData.get_MailFormat())
				oBinaryFileWriter = new AscCommonWord.BinaryFileWriter(oLogicDocument, false, true, options.compatible);
			else
				oBinaryFileWriter = new AscCommonWord.BinaryFileWriter(oLogicDocument, undefined, undefined, options.compatible);
			dataContainer.data = oBinaryFileWriter.Write(oAdditionalData["nobase64"]);
		}
		if (null != options.oMailMergeSendData)
		{
			oAdditionalData["mailmergesend"] = options.oMailMergeSendData;
			var MailMergeMap                 = this.WordControl.m_oLogicDocument.MailMergeMap;
			var aJsonOut                     = [];
			if (MailMergeMap)
			{
				if (MailMergeMap.length > 0)
				{
					var oFirstRow = MailMergeMap[0];
					var aRowOut   = [];
					for (var i in oFirstRow)
						aRowOut.push(i);
					aJsonOut.push(aRowOut);
				}
				//todo может надо запоминать порядок for in в первом столбце, если for in будет по-разному обходить строки
				for (var i = 0; i < MailMergeMap.length; ++i)
				{
					var oRow    = MailMergeMap[i];
					var aRowOut = [];
					for (var j in oRow)
						aRowOut.push(oRow[j]);
					aJsonOut.push(aRowOut);
				}
			}
			var editorData = dataContainer.data;
			dataContainer.data = JSON.stringify(aJsonOut);
			options.oMailMergeSendData.put_IsJson(true);
			//save Editor.bin after json
			var callback = options.callback;
			options.callback = function (incomeObject) {
				oAdditionalData["savekey"] = incomeObject["data"];
				var _dataContainer = {data : editorData, part : null, index : 0, count : 0};
				options.oMailMergeSendData.put_IsJson(false);

				AscCommon.saveWithParts(function(fCallback1, oAdditionalData1, dataContainer1) {
					sendCommand(t, fCallback1, oAdditionalData1, dataContainer1);
				}, t.fCurCallback, callback, oAdditionalData, _dataContainer);
			}
		}

        if (window.isCloudCryptoDownloadAs)
        {
            var sParamXml = ("<m_nCsvTxtEncoding>" + oAdditionalData["codepage"] + "</m_nCsvTxtEncoding>");
            window["AscDesktopEditor"]["CryptoDownloadAs"](dataContainer.data, fileType, sParamXml);
			return true;
        }
	};

	// Вставка диаграмм
	asc_docs_api.prototype.asc_getChartObject = function(type)
	{
		if (!AscFormat.isRealNumber(type))
		{
			this.asc_onOpenChartFrame();
			this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props);
		}

		return this.WordControl.m_oLogicDocument.GetChartObject(type);
	};

	asc_docs_api.prototype.asc_addChartDrawingObject = function(options)
	{
		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddChart);
			AscFonts.IsCheckSymbols = true;
			this.asc_SetSilentMode(true);
			this.WordControl.m_oLogicDocument.AddInlineImage(null, null, null, options);
			AscFonts.IsCheckSymbols = false;

			var oThis = this;
			AscFonts.FontPickerByCharacter.checkText("", this, function() {
				this.asc_SetSilentMode(false, true);
				oThis.WordControl.m_oLogicDocument.FinalizeAction();
			}, false, false, false);
		}
	};
	asc_docs_api.prototype.asc_doubleClickOnChart    = function(obj)
	{
		this.asc_onOpenChartFrame();
		
		if(!window['IS_NATIVE_EDITOR'])
		{
			this.WordControl.onMouseUpMainSimple();
		}
		
		this.sendEvent("asc_doubleClickOnChart", obj);
	};

	asc_docs_api.prototype.asc_doubleClickOnTableOleObject    = function(obj)
	{
		this.isOleEditor = true;	// Для совместного редактирования
		this.asc_onOpenChartFrame();

		if(!window['IS_NATIVE_EDITOR'])
		{
			this.WordControl.onMouseUpMainSimple();
		}
		if(this.asc_checkNeedCallback("asc_doubleClickOnTableOleObject"))
		{
			this.sendEvent("asc_doubleClickOnTableOleObject", obj);
		}
		else
		{
			this.sendEvent("asc_doubleClickOnChart", obj); // TODO: change event type
		}
	};


	asc_docs_api.prototype.asc_onCloseChartFrame               = function()
	{
		AscCommon.baseEditorsApi.prototype.asc_onCloseChartFrame.call(this);
		this.WordControl.m_bIsMouseLock = false;
	};

	asc_docs_api.prototype.asc_editChartDrawingObject = function(chartBinary)
	{
		/**/

		// Находим выделенную диаграмму и накатываем бинарник
		if (AscFormat.isObject(chartBinary))
		{
			if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
			{
				this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_EditChart);
				this.WordControl.m_oLogicDocument.EditChart(chartBinary);
				this.WordControl.m_oLogicDocument.FinalizeAction();
			}
		}
	};

	asc_docs_api.prototype.sync_closeChartEditor = function()
	{
		this.sendEvent("asc_onCloseChartEditor");
	};
	asc_docs_api.prototype.sync_closeOleEditor = function()
	{
		this.sendEvent("asc_onCloseOleEditor");
	};
	asc_docs_api.prototype.asc_setDrawCollaborationMarks = function(bDraw)
	{
		this.tmpCoMarksDraw = bDraw;
		if (!this.isLoadFullApi)
		{
			return;
		}

		if (bDraw !== this.isCoMarksDraw)
		{
			this.isCoMarksDraw = bDraw;
			this.WordControl.m_oDrawingDocument.ClearCachePages();
			this.WordControl.m_oDrawingDocument.FirePaint();
		}
	};

	asc_docs_api.prototype.asc_AddMath = function(Type)
	{
		var loader   = AscCommon.g_font_loader;
		var fontinfo = g_fontApplication.GetFontInfo("Cambria Math");
		var isasync  = loader.LoadFont(fontinfo);
		if (false === isasync)
		{
			return this.asc_AddMath2(Type);
		}
		else
		{
			this.FontAsyncLoadType  = 1;
			this.FontAsyncLoadParam = Type;
		}
	};
	asc_docs_api.prototype.asc_AddMath2 = function(nType)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_AddMath);
			oLogicDocument.AddParaMath(nType);
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_ConvertMathView = function(isToLinear, isAll)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		if (isAll)
			logicDocument.ConvertAllMathView(isToLinear);
		else
			logicDocument.ConvertMathView(isToLinear);
	};
	asc_docs_api.prototype.asc_AddPageCount = function()
	{
		this.WordControl.m_oLogicDocument.AddPageCount();
	};
	//----------------------------------------------------------------------------------------------------------------------
	// Функции для работы с MailMerge
	//----------------------------------------------------------------------------------------------------------------------
	asc_docs_api.prototype.asc_StartMailMerge              = function(oData)
	{
		var t = this;
		if (oData) {
			this.mailMergeFileData = oData;
			this.downloadAs(Asc.c_oAscAsyncAction.MailMergeLoadFile, new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.JSON));
		} else {
			AscCommon.ShowSpreadsheetFileDialog(function (error, files) {
				if (Asc.c_oAscError.ID.No !== error) {
					t.sendEvent("asc_onError", error, Asc.c_oAscError.Level.NoCritical);
					return;
				}
				var format = AscCommon.GetFileExtension(files[0].name);
				var reader = new FileReader();
				reader.onload = function () {
					t.mailMergeFileData = {'url': undefined, 'fileType': format, 'data': new Uint8Array(reader.result)};
					t.downloadAs(Asc.c_oAscAsyncAction.MailMergeLoadFile, new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.JSON));
				};
				reader.onerror = function () {
					t.sendEvent("asc_onError", Asc.c_oAscError.ID.Unknown, Asc.c_oAscError.Level.NoCritical);
				};
				reader.readAsArrayBuffer(files[0]);
			});
		}
	};
	asc_docs_api.prototype.asc_StartMailMergeByList        = function(aList)
	{
		if (!aList || !aList.length || aList.length <= 0)
			aList = [[]];

		var aFields = aList[0];
		if (!aFields || !aFields.length || aFields.length <= 0)
			aFields = [];

		// Пробегаемся по названиям полей и делаем следующее:
		// Если название пустой, тогда задем ему имя "F<номер столбца>"
		// Если название совпадает, тогда добавляем ему число, чтобы имя стало уникальным.

		var UsedNames = {};
		for (var Pos = 0, Count = aFields.length; Pos < Count; Pos++)
		{
			if ("" === aFields[Pos])
				aFields[Pos] = "F" + (Pos + 1);

			if (undefined !== UsedNames[aFields[Pos]])
			{
				var Add = 1;
				var NewName = aFields[Pos] + Add;
				while (undefined !== UsedNames[NewName])
				{
					Add++;
					NewName = aFields[Pos] + Add;
				}
				aFields[Pos] = NewName;
			}

			UsedNames[aFields[Pos]] = 1;
		}

		var DstList = [];
		var FieldsCount = aFields.length;
		for (var Index = 1, Count = aList.length; Index < Count; Index++)
		{
			var oSrcElement = aList[Index];
			var oDstElement = {};
			for (var FieldIndex = 0; FieldIndex < FieldsCount; FieldIndex++)
			{
				var sFieldName = aFields[FieldIndex];
				oDstElement[sFieldName] = oSrcElement[FieldIndex];
			}

			DstList.push(oDstElement);
		}

		this.WordControl.m_oLogicDocument.Start_MailMerge(DstList, aFields);
	};
	asc_docs_api.prototype.asc_GetReceptionsCount          = function()
	{
		return this.WordControl.m_oLogicDocument.Get_MailMergeReceptionsCount();
	};
	asc_docs_api.prototype.asc_GetMailMergeFieldsNameList  = function()
	{
		return this.WordControl.m_oLogicDocument.Get_MailMergeFieldsNameList();
	};
	asc_docs_api.prototype.asc_AddMailMergeField           = function(Name)
	{
		this.WordControl.m_oLogicDocument.Add_MailMergeField(Name);
	};
	asc_docs_api.prototype.asc_SetHighlightMailMergeFields = function(Value)
	{
		this.WordControl.m_oLogicDocument.Set_HightlighMailMergeFields(Value);
	};
	asc_docs_api.prototype.asc_PreviewMailMergeResult      = function(Index)
	{
		this.WordControl.m_oLogicDocument.Preview_MailMergeResult(Index);
	};
	asc_docs_api.prototype.asc_EndPreviewMailMergeResult   = function()
	{
		this.WordControl.m_oLogicDocument.EndPreview_MailMergeResult();
	};
	asc_docs_api.prototype.sync_StartMailMerge             = function()
	{
		this.sendEvent("asc_onStartMailMerge");
	};
	asc_docs_api.prototype.sync_PreviewMailMergeResult     = function(Index)
	{
		this.sendEvent("asc_onPreviewMailMergeResult", Index);
	};
	asc_docs_api.prototype.sync_EndPreviewMailMergeResult  = function()
	{
		this.sendEvent("asc_onEndPreviewMailMergeResult");
	};
	asc_docs_api.prototype.sync_HighlightMailMergeFields   = function(Value)
	{
		this.sendEvent("asc_onHighlightMailMergeFields", Value);
	};
	asc_docs_api.prototype.asc_getMailMergeData            = function()
	{
		return this.WordControl.m_oLogicDocument.Get_MailMergeReceptionsList();
	};
	asc_docs_api.prototype.asc_setMailMergeData            = function(aList)
	{
		this.asc_StartMailMergeByList(aList);
	};
	asc_docs_api.prototype.asc_sendMailMergeData           = function(oData)
	{
		if (this.isLongAction()) {
			return;
		}
		var t = this;
		var actionType = Asc.c_oAscAsyncAction.SendMailMerge;
		oData.put_UserId(this.documentUserId);
		oData.put_RecordCount(oData.get_RecordTo() - oData.get_RecordFrom() + 1);
		var options = new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.TXT);
		options.oMailMergeSendData = oData;
		options.callback = function(input) {
			if (null != input && "sendmm" === input["type"])
			{
				if ("ok" != input["status"])
				{
					t.sendEvent("asc_onError", AscCommon.mapAscServerErrorToAscError(parseInt(input["data"])),
						c_oAscError.Level.NoCritical);
				}
			}
			else
			{
				t.sendEvent("asc_onError", c_oAscError.ID.Unknown, c_oAscError.Level.NoCritical);
			}
			t.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, actionType);
		};
		this.downloadAs(actionType, options);
	};
	asc_docs_api.prototype.asc_GetMailMergeFiledValue      = function(nIndex, sName)
	{
		return this.WordControl.m_oLogicDocument.Get_MailMergeFieldValue(nIndex, sName);
	};
	//----------------------------------------------------------------------------------------------------------------------
	// Работаем со стилями
	//----------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns object which contains direct text and paragraph properties of the first run in selection.
	 * @returns {asc_CStyle}
	 */
	asc_docs_api.prototype.asc_GetStyleFromFormatting = function()
	{
		return this.WordControl.m_oLogicDocument.GetStyleFromFormatting();
	};
	asc_docs_api.prototype.asc_AddNewStyle            = function(oStyle)
	{
		this.WordControl.m_oLogicDocument.Add_NewStyle(oStyle);
	};
	asc_docs_api.prototype.asc_RemoveStyle            = function(sName)
	{
		this.WordControl.m_oLogicDocument.Remove_Style(sName);
	};
	asc_docs_api.prototype.asc_RemoveAllCustomStyles  = function()
	{
		this.WordControl.m_oLogicDocument.Remove_AllCustomStyles();
	};
	asc_docs_api.prototype.asc_IsStyleDefault         = function(sName)
	{
		return this.WordControl.m_oLogicDocument.IsStyleDefaultByName(sName);
	};
	asc_docs_api.prototype.asc_IsDefaultStyleChanged  = function(sName)
	{
		return this.WordControl.m_oLogicDocument.Is_DefaultStyleChanged(sName);
	};
	asc_docs_api.prototype.asc_GetStyleNameById       = function(StyleId)
	{
		return this.WordControl.m_oLogicDocument.Get_StyleNameById(StyleId);
	};
	//----------------------------------------------------------------------------------------------------------------------
	// Работаем с рецензированием
	//----------------------------------------------------------------------------------------------------------------------
	asc_docs_api.prototype.asc_SetTrackRevisions               = function(bTrack)
	{
		return this.asc_SetLocalTrackRevisions(bTrack);
	};
	asc_docs_api.prototype.asc_IsTrackRevisions                = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return false;

		return oLogicDocument.IsTrackRevisions();
	};
	asc_docs_api.prototype.asc_SetLocalTrackRevisions          = function(bTrack)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetLocalTrackRevisions(bTrack, true);
	};
	asc_docs_api.prototype.asc_GetLocalTrackRevisions          = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.GetLocalTrackRevisions();
	};
	asc_docs_api.prototype.asc_SetGlobalTrackRevisions         = function(bTrack)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_Document_Settings))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTrackRevisions);
			oLogicDocument.SetGlobalTrackRevisions(bTrack, true);
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_GetGlobalTrackRevisions         = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.GetGlobalTrackRevisions();
	};
	asc_docs_api.prototype.sync_OnTrackRevisionsChange         = function(localFlag, globalFlag, userId)
	{
		this.sendEvent("asc_onOnTrackRevisionsChange", localFlag, globalFlag, userId);
	};
	asc_docs_api.prototype.sync_BeginCatchRevisionsChanges     = function()
	{
		this.RevisionChangesStack = [];
	};
	asc_docs_api.prototype.sync_EndCatchRevisionsChanges       = function(isShow)
	{
		this.sendEvent("asc_onShowRevisionsChange", this.RevisionChangesStack, isShow);
	};
	asc_docs_api.prototype.asc_GetRevisionsChangesStack        = function()
	{
		return this.RevisionChangesStack;
	};
	asc_docs_api.prototype.sync_AddRevisionsChange             = function(Change)
	{
		this.RevisionChangesStack.push(Change);
	};
	asc_docs_api.prototype.asc_AcceptChanges                   = function(oChange)
	{
		if (oChange)
			this.WordControl.m_oLogicDocument.AcceptRevisionChange(oChange);
		else
			this.WordControl.m_oLogicDocument.AcceptRevisionChangesBySelection();
	};
	asc_docs_api.prototype.asc_RejectChanges                   = function(oChange)
	{
		if (oChange)
			this.WordControl.m_oLogicDocument.RejectRevisionChange(oChange);
		else
			this.WordControl.m_oLogicDocument.RejectRevisionChangesBySelection();
	};
	asc_docs_api.prototype.asc_AcceptChangesBySelection        = function(moveToNext)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;

		logicDocument.AcceptRevisionChangesBySelection();

		if (moveToNext)
			logicDocument.GetNextRevisionChange();
	};
	asc_docs_api.prototype.asc_RejectChangesBySelection        = function(moveToNext)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;

		logicDocument.RejectRevisionChangesBySelection();

		if (moveToNext)
			logicDocument.GetNextRevisionChange();
	};
	asc_docs_api.prototype.asc_HaveRevisionsChanges            = function(isCheckOwnChanges)
	{
		if (!this.WordControl.m_oLogicDocument)
			return false;
		return this.WordControl.m_oLogicDocument.HaveRevisionChanges(isCheckOwnChanges);
	};
	asc_docs_api.prototype.asc_HaveNewRevisionsChanges         = function()
	{
		return this.asc_HaveRevisionsChanges();
	};
	asc_docs_api.prototype.asc_GetNextRevisionsChange          = function()
	{
		return this.WordControl.m_oLogicDocument.GetNextRevisionChange();
	};
	asc_docs_api.prototype.asc_GetPrevRevisionsChange          = function()
	{
		return this.WordControl.m_oLogicDocument.GetPrevRevisionChange();
	};
	asc_docs_api.prototype.sync_UpdateRevisionsChangesPosition = function(X, Y)
	{
		this.sendEvent("asc_onUpdateRevisionsChangesPosition", X, Y);
	};
	asc_docs_api.prototype.asc_AcceptAllChanges                = function()
	{
		this.WordControl.m_oLogicDocument.AcceptAllRevisionChanges();
	};
	asc_docs_api.prototype.asc_RejectAllChanges                = function()
	{
		this.WordControl.m_oLogicDocument.RejectAllRevisionChanges();
	};
	asc_docs_api.prototype.asc_GetTrackRevisionsReportByAuthors= function()
	{
		var oResult = {};
		var oAllChanges = this.WordControl.m_oLogicDocument.TrackRevisionsManager.Get_AllChanges();
		for (var ParaId in oAllChanges)
		{
			var arrChanges = oAllChanges[ParaId];
			for (var nIndex = 0, nCount = arrChanges.length; nIndex < nCount; ++nIndex)
			{
				var oChange   = arrChanges[nIndex];
				var sUserName = oChange.get_UserName();
				var nDateTime = oChange.get_DateTime();

				if (!oResult[sUserName])
					oResult[sUserName] = [];

				var arrUserChanges = oResult[sUserName];

				var nPos = 0;
				var nLen = arrUserChanges.length;
				while (nPos < nLen)
				{
					if (nDateTime < arrUserChanges[nPos].get_DateTime())
						break;

					nPos++;
				}

				arrUserChanges.splice(nPos, 0, oChange);
			}
		}

		return oResult;
	};
	asc_docs_api.prototype.asc_FollowRevisionMove = function(oChange)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var sMoveId = oChange.GetMoveId();
		var isFrom  = oChange.GetMoveType() === Asc.c_oAscRevisionsMove.MoveTo;
		oLogicDocument.SelectTrackMove(sMoveId, isFrom, true, true);
	};

	asc_docs_api.prototype.asc_undoAllChanges       = function()
	{
		this.WordControl.m_oLogicDocument.Document_Undo({All : true});
	};
	asc_docs_api.prototype.asc_CloseFile            = function()
	{
		AscCommon.History.Clear();
		g_oIdCounter.Clear();
		g_oTableId.Clear();
		AscCommon.CollaborativeEditing.Clear();
		this.isApplyChangesOnOpenEnabled = true;
		this.isDocumentLoadComplete = false;
		this.ServerImagesWaitComplete = false;
		this.turnOffSpecialModes();

		var oLogicDocument = this.WordControl.m_oLogicDocument;
		oLogicDocument.StopRecalculate();
		oLogicDocument.StopSpellCheck();
		AscCommon.pptx_content_loader.ImageMapChecker = {};

		this.WordControl.m_oDrawingDocument.CloseFile();
		
		this.sendEvent("asc_onCloseFile");
	};
	asc_docs_api.prototype.asc_SetFastCollaborative = function(isOn)
	{
		if (!this.WordControl || !this.WordControl.m_oLogicDocument)
			return;
		if (AscCommon.CollaborativeEditing){
			AscCommon.CollaborativeEditing.Set_Fast(isOn);
			if(window['AscCommon'].g_specialPasteHelper && isOn && !AscCommon.CollaborativeEditing.Is_SingleUser()){
				window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide();
			}
		}
	};

	asc_docs_api.prototype._onEndLoadSdk = function()
	{
		AscCommon.baseEditorsApi.prototype._onEndLoadSdk.call(this);
		
		if (this.headingsColor)
		{
			AscWord.setDefaultHeadingColorStr(this.headingsColor);
		}
		
		History           = AscCommon.History;
		g_fontApplication = AscFonts.g_fontApplication;
		PasteElementsId   = AscCommon.PasteElementsId;
		global_mouseEvent = AscCommon.global_mouseEvent;

		this.WordControl      = new AscCommonWord.CEditorPage(this);
		this.WordControl.Name = this.HtmlElementName;

		this.CurrentTranslate = AscCommonWord.translations_map["en"];

		//выставляем тип copypaste
		PasteElementsId.g_bIsDocumentCopyPaste = true;

		this.CreateComponents();
		this.WordControl.Init();

		if (this.tmpFontRenderingMode)
		{
			this.SetFontRenderingMode(this.tmpFontRenderingMode);
		}
		if (null !== this.tmpViewRulers)
		{
			this.asc_SetViewRulers(this.tmpViewRulers);
		}
		if (null !== this.tmpZoomType)
		{
			switch (this.tmpZoomType)
			{
				case AscCommon.c_oZoomType.FitToPage:
					this.zoomFitToPage();
					break;
				case AscCommon.c_oZoomType.FitToWidth:
					this.zoomFitToWidth();
					break;
				case AscCommon.c_oZoomType.CustomMode:
					this.zoomCustomMode();
					break;
			}
		}
		if (null != this.tmpDocumentUnits)
		{
			this.asc_SetDocumentUnits(this.tmpDocumentUnits);
			this.tmpDocumentUnits = null;
		}

		this.asc_setViewMode(this.isViewMode);
		this.asc_setDrawCollaborationMarks(this.tmpCoMarksDraw);

		if (this.isOnlyReaderMode)
			this.ImageLoader.bIsAsyncLoadDocumentImages = false;

        if (this.openFileCryptBinary)
        {
            this.openFileCryptCallback(this.openFileCryptBinary);
        }
	};
	asc_docs_api.prototype.initCollaborativeEditing = function()
	{
		if (AscCommon.CollaborativeEditing)
			return;
		
		AscCommon.CollaborativeEditing = new AscCommon.CWordCollaborativeEditing();
	};

	asc_docs_api.prototype.asc_Recalculate = function(bIsUpdateInterface)
	{
		if (!this.WordControl.m_oLogicDocument)
			return;

		return this.WordControl.m_oLogicDocument.RecalculateFromStart(bIsUpdateInterface);
	};

	asc_docs_api.prototype.asc_canPaste = function()
	{
		if (!this.WordControl ||
			!this.WordControl.m_oLogicDocument ||
			this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
			return false;

		this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_PasteHotKey);
		return true;
	};
	asc_docs_api.prototype.canRunBuilderScript = function()
	{
		let logicDocument = this.getLogicDocument();
		if (!logicDocument)
			return false;
		
		// Разрешаем всегда выполнять скрипт билдера, даже если это вьювер, а скрипт меняет содержимое
		// В конце действия по выполненным изменениям проверяем можно ли оставлять данные изменения
		logicDocument.StartAction(AscDFH.historydescription_BuilderScript);
		logicDocument.CheckActionLock();
		return true;
	};
	asc_docs_api.prototype._onEndBuilderScript = function(callback)
	{
		let logicDocument = this.getLogicDocument();
		logicDocument.Recalculate();
		logicDocument.UpdateSelection();
		logicDocument.UpdateInterface();
		let result = logicDocument.FinalizeAction();
		if (callback)
			callback(result);
		
		return result;
	};
	//----------------------------------------------------------------------------------------------------------------------
	// Работаем с полями
	//----------------------------------------------------------------------------------------------------------------------
	asc_docs_api.prototype.asc_GetBlockChainData = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return [];

		var arrFields = oLogicDocument.GetAllFormTextFields();
		var arrResult = [];
		for (var nIndex = 0, nCount = arrFields.length; nIndex < nCount; ++nIndex)
		{
			var oField = arrFields[nIndex];
			arrResult.push(oField.GetValue());
		}
		return arrResult;
	};
	asc_docs_api.prototype.asc_SetBlockChainData = function(arrData)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return [];

		var arrFields = oLogicDocument.GetAllFormTextFields();
		var arrResult = [];
		for (var nIndex = 0, nCount = arrFields.length; nIndex < nCount; ++nIndex)
		{
			var oField = arrFields[nIndex];
			oField.SetValue(arrData[nIndex] ? arrData[nIndex] : "");
		}

		if (!this.isLongAction())
			oLogicDocument.RecalculateFromStart();
	};
	//----------------------------------------------------------------------------------------------------------------------
	// Работаем с ContentControl
	//----------------------------------------------------------------------------------------------------------------------
	asc_docs_api.prototype.asc_AddContentControl = function(nType, oContentControlPr)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return null;

		var oResult = null;
		if (c_oAscSdtLevelType.Block === nType)
		{
			if (false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_ContentControl_Add, null, true, false, null, AscDFH.historydescription_Document_AddBlockLevelContentControl))
			{
				oLogicDocument.StartAction(AscDFH.historydescription_Document_AddBlockLevelContentControl);

				var oContentControl = oLogicDocument.AddContentControl(c_oAscSdtLevelType.Block);
				if (oContentControl)
				{
					if (oContentControlPr)
						oContentControl.SetContentControlPr(oContentControlPr);

					oLogicDocument.Recalculate();
					oLogicDocument.UpdateInterface();
					oLogicDocument.UpdateSelection();

					oResult = oContentControl.GetContentControlPr();
				}

				oLogicDocument.FinalizeAction();
			}
		}
		else if (c_oAscSdtLevelType.Inline === nType)
		{
			if (false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_ContentControl_Add, null, true, false, null, AscDFH.historydescription_Document_AddInlineLevelContentControl))
			{
				oLogicDocument.StartAction(AscDFH.historydescription_Document_AddInlineLevelContentControl);

				var oContentControl = oLogicDocument.AddContentControl(c_oAscSdtLevelType.Inline);
				if (oContentControl)
				{
					if (oContentControlPr)
						oContentControl.SetContentControlPr(oContentControlPr);

					oLogicDocument.Recalculate();
					oLogicDocument.UpdateInterface();
					oLogicDocument.UpdateSelection();

					oResult = oContentControl.GetContentControlPr();
				}

				oLogicDocument.FinalizeAction();
			}
		}
		return oResult;
	};
	asc_docs_api.prototype.asc_AddContentControlCheckBox = function(oPr, oInternalPr, oCommonPr)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var nCheckedSymbol   = oPr && oPr.CheckedSymbol ? oPr.CheckedSymbol : Asc.c_oAscSdtCheckBoxDefaults.CheckedSymbol;
		var nUncheckedSymbol = oPr && oPr.UncheckedSymbol ? oPr.UncheckedSymbol : Asc.c_oAscSdtCheckBoxDefaults.UncheckedSymbol;
		var sCheckedFont     = oPr && oPr.CheckedFont ? oPr.CheckedFont : Asc.c_oAscSdtCheckBoxDefaults.CheckedFont;
		var sUncheckedFont   = oPr && oPr.UncheckedFont ? oPr && oPr.UncheckedFont : Asc.c_oAscSdtCheckBoxDefaults.UncheckedFont;

		var isLoadFonts = false;
		if (!AscCommon.IsAscFontSupport(sCheckedFont, nCheckedSymbol))
		{
			isLoadFonts = true;
			AscFonts.FontPickerByCharacter.getFontBySymbol(nCheckedSymbol);
		}

		if (!AscCommon.IsAscFontSupport(sUncheckedFont, nUncheckedSymbol))
		{
			isLoadFonts = true;
			AscFonts.FontPickerByCharacter.getFontBySymbol(nUncheckedSymbol);
		}

		function private_PerformAddCheckBox()
		{
			oLogicDocument.RemoveTextSelection();
			if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
			{
				oLogicDocument.StartAction(AscDFH.historydescription_Document_AddContentControlCheckBox);

				var oCC = oLogicDocument.AddContentControlCheckBox(oPr);

				if (oCC && oCommonPr)
					oCC.SetContentControlPr(oCommonPr);

				oLogicDocument.UpdateInterface();
				oLogicDocument.Recalculate();
				oLogicDocument.FinalizeAction();
			}
		}

		if (isLoadFonts)
		{
			var oFonts = {};
			oFonts[sCheckedFont]   = true;
			oFonts[sUncheckedFont] = true;

			AscCommon.Check_LoadingDataBeforePrepaste(this, oFonts, {}, private_PerformAddCheckBox);
		}
		else
		{
			private_PerformAddCheckBox();
		}
	};
	asc_docs_api.prototype.asc_AddContentControlPicture = function(oInternalPr, oCommonPr)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		oLogicDocument.RemoveTextSelection();
		if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_AddContentControlPicture);

			var oCC = oLogicDocument.AddContentControlPicture();

			if (oCC && oCommonPr)
				oCC.SetContentControlPr(oCommonPr);

			oLogicDocument.UpdateInterface();
			oLogicDocument.Recalculate();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_AddContentControlList = function(isComboBox, oPr, oInternalPr, oCommonPr)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		oLogicDocument.RemoveTextSelection();
		if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_AddContentControlList);

			var oCC;
			if (isComboBox)
				oCC = oLogicDocument.AddContentControlComboBox(oPr);
			else
				oCC = oLogicDocument.AddContentControlDropDownList(oPr);

			if (oCC && oCommonPr)
				oCC.SetContentControlPr(oCommonPr);

			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.UpdateSelection();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_AddContentControlDatePicker = function(oPr, oCommonPr)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		oLogicDocument.RemoveTextSelection();
		if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_AddContentControlList);
			var oCC = oLogicDocument.AddContentControlDatePicker(oPr, oCommonPr);

			if (oCC && oCommonPr)
				oCC.SetContentControlPr(oCommonPr);

			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.UpdateSelection();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_RemoveContentControl = function(Id)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		var isLocked        = true;
		var oContentControl = null;
		if (undefined === Id)
		{
			var oInfo          = oLogicDocument.GetSelectedElementsInfo({SkipTOC : true});
			var oInlineControl = oInfo.GetInlineLevelSdt();
			var oBlockControl  = oInfo.GetBlockLevelSdt();

			if (oInlineControl)
				oContentControl = oInlineControl;
			else if (oBlockControl)
				oContentControl = oBlockControl;
		}
		else
		{
			oContentControl = AscCommon.g_oTableId.Get_ById(Id);
		}

		if (oContentControl && oContentControl.GetContentControlType)
		{
			oLogicDocument.CheckFormPlaceHolder = false;
			oLogicDocument.RemoveSelection();
			oContentControl.SelectContentControl();
			oLogicDocument.CheckFormPlaceHolder = true;

			if (oContentControl.IsFixedForm())
			{
				var oShape = oContentControl.GetFixedFormWrapperShape();
				if (!oShape || !(oShape.parent instanceof AscCommonWord.ParaDrawing))
					return;

				oLogicDocument.Select_DrawingObject(oShape.parent.GetId());
				if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_Remove))
				{
					oLogicDocument.StartAction(AscDFH.historydescription_Document_RemoveContentControl);
					oLogicDocument.Remove(-1, true, false, false, false);
					oLogicDocument.Recalculate();
					oLogicDocument.UpdateInterface();
					oLogicDocument.UpdateSelection();
					oLogicDocument.FinalizeAction();
				}

				return;
			}

			if (c_oAscSdtLevelType.Block === oContentControl.GetContentControlType())
			{
				isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
					Type      : AscCommon.changestype_2_ElementsArray_and_Type,
					Elements  : [oContentControl],
					CheckType : AscCommon.changestype_ContentControl_Remove
				});
			}
			else if (c_oAscSdtLevelType.Inline === oContentControl.GetContentControlType())
			{
				var oParagraph = oContentControl.GetParagraph();
				if (oParagraph)
				{
					isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
						Type      : AscCommon.changestype_2_ElementsArray_and_Type,
						Elements  : [oParagraph],
						CheckType : AscCommon.changestype_Remove
					});
				}
			}

			Id = oContentControl.GetId();
		}

		if (false === isLocked)
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_RemoveContentControl);
			oLogicDocument.RemoveContentControl(Id);
			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.UpdateSelection();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_RemoveContentControlWrapper = function(Id)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		var isLocked        = true;
		var oContentControl = oLogicDocument.GetContentControl(Id);

		if (oContentControl && oContentControl.GetContentControlType)
		{
			oContentControl.SkipSpecialContentControlLock(true);
			if (c_oAscSdtLevelType.Block === oContentControl.GetContentControlType())
			{
				isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
					Type      : AscCommon.changestype_2_ElementsArray_and_Type,
					Elements  : [oContentControl],
					CheckType : AscCommon.changestype_ContentControl_Remove
				});
			}
			else if (c_oAscSdtLevelType.Inline === oContentControl.GetContentControlType())
			{
				var oParagraph = oContentControl.GetParagraph();
				if (oParagraph)
				{
					var oState = oLogicDocument.SaveDocumentState();
					oContentControl.SelectContentControl();

					isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
						Type      : AscCommon.changestype_2_ElementsArray_and_Type,
						Elements  : [oParagraph],
						CheckType : AscCommon.changestype_Remove
					});

					oLogicDocument.LoadDocumentState(oState);
				}
			}

			Id = oContentControl.GetId();
			oContentControl.SkipSpecialContentControlLock(false);
		}
		
		if (isLocked)
			return false;
		
		oLogicDocument.StartAction(AscDFH.historydescription_Document_RemoveContentControlWrapper);
		oLogicDocument.RemoveContentControlWrapper(Id);
		oLogicDocument.Recalculate();
		oLogicDocument.UpdateInterface();
		oLogicDocument.UpdateSelection();
		oLogicDocument.FinalizeAction();
		return true;
	};
	asc_docs_api.prototype.asc_SetContentControlProperties = function(oContentControlPr, Id, isApplyToAll)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		function preSetup() {
			oLogicDocument.StartAction(AscDFH.historydescription_Document_ChangeContentControlProperties);
			AscFonts.IsCheckSymbols = true;
		}
		function postSetup() {
			AscFonts.IsCheckSymbols = false;
			AscFonts.FontPickerByCharacter.checkText("", oLogicDocument, function() {
				this.Recalculate();
				this.UpdateInterface();
				this.UpdateSelection();
				this.FinalizeAction();
			}, false, false, false);
		}

		if (true === isApplyToAll)
		{
			var arrContentControls = oLogicDocument.GetAllContentControls();

			var arrCheckElements = [], arrCheckTypes = [];
			for (var nIndex = 0, nCount = arrContentControls.length; nIndex < nCount; ++nIndex)
			{
				var oContentControl = arrContentControls[nIndex];
				if (c_oAscSdtLevelType.Block === oContentControl.GetContentControlType())
				{
					arrCheckElements.push(oContentControl);
					arrCheckTypes.push(AscCommon.changestype_ContentControl_Properties);
				}
				else if (c_oAscSdtLevelType.Inline === oContentControl.GetContentControlType())
				{
					var oParagraph = oContentControl.GetParagraph();
					if (oParagraph)
					{
						arrCheckElements.push(oParagraph);
						arrCheckTypes.push(AscCommon.changestype_Paragraph_Properties);
					}
				}
			}

			if (false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
				Type       : AscCommon.changestype_2_Element_and_Type_Array,
				Elements   : arrCheckElements,
				CheckTypes : arrCheckTypes
			}))
			{
				preSetup();

				for (var nIndex = 0, nCount = arrContentControls.length; nIndex < nCount; ++nIndex)
				{
					arrContentControls[nIndex].SetContentControlPr(oContentControlPr);
				}

				postSetup();
			}
		}
		else
		{
			var isLocked        = true;
			var oContentControl = oLogicDocument.GetContentControl(Id);

			if (oContentControl && oContentControl.GetContentControlType)
			{
				if ((Asc.c_oAscSdtLockType.ContentLocked === oContentControl.GetContentControlLock()
					|| Asc.c_oAscSdtLockType.SdtContentLocked === oContentControl.GetContentControlLock())
					&& oContentControlPr
					&& Asc.c_oAscSdtLockType.Unlocked !== oContentControlPr.GetLock()
					&& Asc.c_oAscSdtLockType.SdtLocked !== oContentControlPr.GetLock())
				{
					if (oContentControl.IsDatePicker() && !oContentControl.GetDatePickerPr().IsEqual(oContentControlPr.DateTimePr))
					{
						oContentControlPr.DateTimePr = oContentControl.GetDatePickerPr().Copy();
					}
				}

				if (c_oAscSdtLevelType.Block === oContentControl.GetContentControlType())
				{
					isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
						Type      : AscCommon.changestype_2_ElementsArray_and_Type,
						Elements  : [oContentControl],
						CheckType : AscCommon.changestype_ContentControl_Properties
					});
				}
				else if (c_oAscSdtLevelType.Inline === oContentControl.GetContentControlType())
				{
					var oParagraph = oContentControl.GetParagraph();
					if (oParagraph)
					{
						isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
							Type      : AscCommon.changestype_2_ElementsArray_and_Type,
							Elements  : [oParagraph],
							CheckType : AscCommon.changestype_Paragraph_Properties
						});
					}
				}
			}

			if (false === isLocked)
			{
				preSetup();
				oContentControl.SetContentControlPr(oContentControlPr);
				postSetup();
			}
		}
	};
	asc_docs_api.prototype.asc_IsContentControl = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return false;

		var oCC = oLogicDocument.GetContentControl();
		return (oCC && !oCC.IsContentControlEquation());
	};
	asc_docs_api.prototype.asc_GetContentControlProperties = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return null;

		var oContentControl = oLogicDocument.GetContentControl();

		return oContentControl && !oContentControl.IsContentControlEquation() ? oContentControl.GetContentControlPr() : null;
	};
	asc_docs_api.prototype.asc_GetCurrentContentControl = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return null;

		var oContentControl = oLogicDocument.GetContentControl();
		return oContentControl && !oContentControl.IsContentControlEquation() ? oContentControl.GetId() : null;
	};
	asc_docs_api.prototype.sync_ContentControlCallback = function(oContentControlPr)
	{
		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.ContentControl, oContentControlPr);
	};
	asc_docs_api.prototype.asc_ClearContentControl = function(sId)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl)
			return;

		oLogicDocument.RemoveSelection();
		oContentControl.SelectContentControl();
		oContentControl.SkipSpecialContentControlLock(true);

		if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, undefined, false, oLogicDocument.IsFormFieldEditing()))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_ClearContentControl);
			oContentControl.ClearContentControlExt();

			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.UpdateSelection();

			oLogicDocument.FinalizeAction();
		}
		oContentControl.SkipSpecialContentControlLock(false);
	};
	asc_docs_api.prototype.asc_GetContentControlRightAnchorPosition = function(sId)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return null;

		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl)
			return null;

		var oAnchor = oContentControl.GetBoundingRect();

		if (!oAnchor || oAnchor.Page < 0)
			return null;

		var nX = oAnchor.X + oAnchor.W;
		var nY = oAnchor.Y + oAnchor.H / 2;

		if (oAnchor.Transform)
		{
			nX = oAnchor.Transform.TransformPointX(oAnchor.X + oAnchor.W, oAnchor.Y + oAnchor.H / 2);
			nY = oAnchor.Transform.TransformPointY(oAnchor.X + oAnchor.W, oAnchor.Y + oAnchor.H / 2);
		}

		var oRealCoords = oLogicDocument.GetDrawingDocument().ConvertCoordsToCursorWR(nX, nY, oAnchor.Page);
		return [oRealCoords.X, oRealCoords.Y];
	};
	asc_docs_api.prototype.asc_GetContentControlBoundingRect = function(sId)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return null;

		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl)
			return null;

		var oRect = oContentControl.GetBoundingRect();

		if (!oRect || oRect.Page < 0)
			return null;

		var nX, nY, nW, nH;
		var oTransform = oRect.Transform;
		if (oTransform)
		{
			var nX0 = oTransform.TransformPointX(oRect.X, oRect.Y);
			var nY0 = oTransform.TransformPointY(oRect.X, oRect.Y);
			var nX1 = oTransform.TransformPointX(oRect.X + oRect.W, oRect.Y);
			var nY1 = oTransform.TransformPointY(oRect.X + oRect.W, oRect.Y);
			var nX2 = oTransform.TransformPointX(oRect.X + oRect.W, oRect.Y + oRect.H);
			var nY2 = oTransform.TransformPointY(oRect.X + oRect.W, oRect.Y + oRect.H);
			var nX3 = oTransform.TransformPointX(oRect.X, oRect.Y + oRect.H);
			var nY3 = oTransform.TransformPointY(oRect.X, oRect.Y + oRect.H);

			nX = Math.min(nX0, nX1, nX2, nX3);
			nY = Math.min(nY0, nY1, nY2, nY3);
			nW = Math.max(nX0, nX1, nX2, nX3) - nX;
			nH = Math.max(nY0, nY1, nY2, nY3) - nY;
		}
		else
		{
			nX = oRect.X;
			nY = oRect.Y;
			nW = oRect.W;
			nH = oRect.H;
		}

		var oRealCoords0 = oLogicDocument.GetDrawingDocument().ConvertCoordsToCursorWR(nX, nY, oRect.Page);
		var oRealCoords1 = oLogicDocument.GetDrawingDocument().ConvertCoordsToCursorWR(nX + nW, nY + nH, oRect.Page);
		return [oRealCoords0.X, oRealCoords0.Y, oRealCoords1.X, oRealCoords1.Y];
	};
	asc_docs_api.prototype.asc_SetGlobalContentControlHighlightColor = function(r, g, b)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument || !oLogicDocument.CanPerformAction())
			return;

		// Если цвет не задан
		if (undefined === r || null == r)
		{
			r = 220;
			g = 220;
			b = 220;
		}

		// Лок можно не проверять, такие изменения нормально мержаться
		oLogicDocument.StartAction(AscDFH.historydescription_Document_SetGlobalSdtHighlightColor);
		oLogicDocument.SetSdtGlobalColor(r, g, b);

		oLogicDocument.GetDrawingDocument().ClearCachePages();
		oLogicDocument.GetDrawingDocument().FirePaint();
		oLogicDocument.FinalizeAction();
	};
	asc_docs_api.prototype.sync_OnChangeSdtGlobalSettings = function()
	{
		this.sendEvent("asc_onChangeSdtGlobalSettings");
	};
	asc_docs_api.prototype.asc_GetGlobalContentControlHighlightColor = function(isDefault)
	{
		if (true === isDefault)
			return new Asc.asc_CColor(220, 220, 220);

		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return new Asc.asc_CColor(0, 0, 0);

		var oColor = oLogicDocument.GetSdtGlobalColor();
		return new Asc.asc_CColor(oColor.r, oColor.g, oColor.b);
	};
	asc_docs_api.prototype.asc_SetGlobalContentControlShowHighlight = function(isShow, r, g, b)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument || !oLogicDocument.CanPerformAction())
			return;

		// Лок можно не проверять, таким изменения нормально мержаться
		oLogicDocument.StartAction(AscDFH.historydescription_Document_SetGlobalSdtShowHighlight);
		oLogicDocument.SetSdtGlobalShowHighlight(isShow);

		if (undefined !== r && undefined !== g && undefined !== b)
			oLogicDocument.SetSdtGlobalColor(r, g, b);

		oLogicDocument.Redraw();
		oLogicDocument.FinalizeAction();
	};
	asc_docs_api.prototype.asc_GetGlobalContentControlShowHighlight = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return false;

		return oLogicDocument.GetSdtGlobalShowHighlight();
	};
	asc_docs_api.prototype.asc_SetContentControlCheckBoxPr = function(oPr)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument || !oPr)
			return;

		var oContentControl = oLogicDocument.GetContentControl();
		if (!oContentControl || !oContentControl.IsCheckBox())
			return;

		if (oContentControl.IsForm() && oPr)
		{
			if (oPr.GroupKey)
			{
				oPr.CheckedSymbol   = 0x25C9;
				oPr.UncheckedSymbol = 0x25CB;
			}
			else
			{
				oPr.CheckedSymbol   = 0x2611;
				oPr.UncheckedSymbol = 0x2610;
			}

			oPr.CheckedFont   = "Segoe UI Symbol";
			oPr.UncheckedFont = "Segoe UI Symbol";
		}

		var nCheckedSymbol   = oPr.CheckedSymbol ? oPr.CheckedSymbol : Asc.c_oAscSdtCheckBoxDefaults.CheckedSymbol;
		var nUncheckedSymbol = oPr.UncheckedSymbol ? oPr.UncheckedSymbol : Asc.c_oAscSdtCheckBoxDefaults.UncheckedSymbol;
		var sCheckedFont     = oPr.CheckedFont ? oPr.CheckedFont : Asc.c_oAscSdtCheckBoxDefaults.CheckedFont;
		var sUncheckedFont   = oPr.UncheckedFont ? oPr.UncheckedFont : Asc.c_oAscSdtCheckBoxDefaults.UncheckedFont;

		var isLoadFonts = false;
		if (!AscCommon.IsAscFontSupport(sCheckedFont, nCheckedSymbol))
		{
			isLoadFonts = true;
			AscFonts.FontPickerByCharacter.getFontBySymbol(nCheckedSymbol);
		}

		if (!AscCommon.IsAscFontSupport(sUncheckedFont, nUncheckedSymbol))
		{
			isLoadFonts = true;
			AscFonts.FontPickerByCharacter.getFontBySymbol(nUncheckedSymbol);
		}

		function private_PeformSetCheckBoxPr()
		{
			var oParagraph = oContentControl.GetParagraph();
			if (oParagraph && !oLogicDocument.IsSelectionLocked(AscCommon.changestype_None, {
				Type      : AscCommon.changestype_2_ElementsArray_and_Type,
				Elements  : [oParagraph],
				CheckType : AscCommon.changestype_Paragraph_Properties
			}))
			{
				oLogicDocument.StartAction(AscDFH.historydescription_Document_SetContentControlCheckBoxPr);
				oContentControl.ApplyCheckBoxPr(oPr);
				oLogicDocument.Recalculate();
				oLogicDocument.UpdateTracks();
				oLogicDocument.FinalizeAction();
			}
		}

		if (isLoadFonts)
		{
			var oFonts = {};
			oFonts[oPr.CheckedFont]   = true;
			oFonts[oPr.UncheckedFont] = true;

			AscCommon.Check_LoadingDataBeforePrepaste(this, oFonts, {}, private_PeformSetCheckBoxPr);
		}
		else
		{
			private_PeformSetCheckBoxPr();
		}
	};
	asc_docs_api.prototype.asc_SetContentControlCheckBoxChecked = function(isChecked, sId)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl || !oContentControl.IsCheckBox())
			return;

		oContentControl.SkipSpecialContentControlLock(true);

		if (!oContentControl.CanBeEdited())
		{
			oContentControl.SkipSpecialContentControlLock(false);
			return;
		}

		var isLocked = false;
		if (c_oAscSdtLevelType.Block === oContentControl.GetContentControlType())
		{
			isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
				Type      : AscCommon.changestype_2_ElementsArray_and_Type,
				Elements  : [oContentControl],
				CheckType : AscCommon.changestype_Paragraph_AddText
			}, false, oLogicDocument.IsFormFieldEditing());
		}
		else if (c_oAscSdtLevelType.Inline === oContentControl.GetContentControlType())
		{
			var oParagraph = oContentControl.GetParagraph();
			if (oParagraph)
			{
				var oState = oLogicDocument.SaveDocumentState();
				oContentControl.SelectContentControl();

				isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
					Type      : AscCommon.changestype_2_ElementsArray_and_Type,
					Elements  : [oParagraph],
					CheckType : AscCommon.changestype_Paragraph_AddText
				}, false, oLogicDocument.IsFormFieldEditing());

				oLogicDocument.LoadDocumentState(oState);
			}
		}
		oContentControl.SkipSpecialContentControlLock(false);

		if (!isLocked)
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_SelectContentControlListItem);
			oContentControl.ToggleCheckBox(isChecked);
			oLogicDocument.RemoveSelection();
			oContentControl.MoveCursorToContentControl(true);
			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.UpdateTracks();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_IsContentControlCheckBoxChecked = function(sId)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return false;

		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl || !oContentControl.IsCheckBox())
			return false;

		return oContentControl.IsCheckBoxChecked();
	};
	asc_docs_api.prototype.asc_SetContentControlPictureUrl = function(sUrl, sId, sToken)
	{
		if (this.WordControl && this.WordControl.m_oDrawingDocument)
		{
			this.WordControl.m_oDrawingDocument.UnlockCursorType();
		}
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument || AscCommon.isNullOrEmptyString(sUrl))
			return;

		var oCC = oLogicDocument.GetContentControl(sId);
		oCC.SkipSpecialContentControlLock(true);
		oCC.SkipFillingFormModeCheck(true);
		if (!oCC || !oCC.IsPicture() || !oCC.SelectPicture() || !oCC.CanBeEdited())
		{
			oCC.SkipFillingFormModeCheck(false);
			oCC.SkipSpecialContentControlLock(false);
			return;
		}

		if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_Image_Properties, undefined, false, oLogicDocument.IsFormFieldEditing()))
		{
			oCC.SkipFillingFormModeCheck(false);
			oCC.SkipSpecialContentControlLock(false);

			var oImagePr = {
				ImageUrl : sUrl
			};

			var sImageUrl = null, fReplaceCallback = null, sImageToDownLoad = "";

			if (!g_oDocumentUrls.getImageLocal(sUrl))
			{
				sImageUrl        = sUrl;
				fReplaceCallback = function(sUrl)
				{
					oImagePr.ImageUrl = sUrl;
					sImageToDownLoad  = sUrl;
				}
			}

			sImageToDownLoad = sUrl;

			var oApi = this;
			var fApplyCallback = function()
			{
				var fPropsCallback = function(_img)
				{
					if(_img && _img.Image && oImagePr)
					{
						var oDrawingObjects = oApi.WordControl.m_oLogicDocument.DrawingObjects;
						if(oDrawingObjects && oDrawingObjects.selectedObjects[0])
						{
							var dWidth = oDrawingObjects.selectedObjects[0].extX;
							var dHeight = oDrawingObjects.selectedObjects[0].extY;
							var __w = Math.max((_img.Image.width * AscCommon.g_dKoef_pix_to_mm), 1);
							var __h = Math.max((_img.Image.height * AscCommon.g_dKoef_pix_to_mm), 1);
							var fKoeff = 1.0/Math.max(__w/dWidth, __h/dHeight);
							var _w      = Math.max(5, __w*fKoeff);
							var _h      = Math.max(5, __h*fKoeff);
							oImagePr.Width = _w;
							oImagePr.Height = _h;
						}
					}
					oApi.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_ApplyImagePrWithUrl);
					oApi.WordControl.m_oLogicDocument.SetImageProps(oImagePr);
					oCC.SetShowingPlcHdr(false);
					
					if (oCC.IsPictureForm())
					{
						oCC.UpdatePictureFormLayout();

						var oShape = oCC.GetFixedFormWrapperShape();
						if (oShape && oShape.parent instanceof AscCommonWord.ParaDrawing)
							oApi.WordControl.m_oLogicDocument.Select_DrawingObject(oShape.parent.GetId());
					}

					oApi.WordControl.m_oLogicDocument.UpdateTracks();
					oApi.WordControl.m_oLogicDocument.FinalizeAction();
				};
				var _img = oApi.ImageLoader.LoadImage(sImageToDownLoad, 1);
				if (null != _img)
				{
					fPropsCallback(_img);
				}
				else
				{
					oApi.asyncImageEndLoaded2 = function(_img)
					{
						fPropsCallback(_img);
					}
				}
			};

			if (sImageUrl)
			{
				if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]())
				{
					var _url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](sImageToDownLoad);
					_url     = g_oDocumentUrls.getImageUrl(_url);
					fReplaceCallback(_url);
					fApplyCallback();
					return;
				}

				AscCommon.sendImgUrls(this, [sImageToDownLoad], function(data)
				{
					if (data && data[0] && data[0].url !== "error")
					{
						fReplaceCallback(data[0].url);
						fApplyCallback();
					}
				}, undefined, sToken);
			}
			else
			{
				fApplyCallback();
			}
		}
		else
		{
			oCC.SkipFillingFormModeCheck(false);
			oCC.SkipSpecialContentControlLock(false);
		}
	};
	asc_docs_api.prototype.asc_SetContentControlListPr = function(oPr, sId)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var isLocked        = true;
		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl || (!oContentControl.IsComboBox() && !oContentControl.DropDownList()))
			return;

		if (c_oAscSdtLevelType.Block === oContentControl.GetContentControlType())
		{
			isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
				Type      : AscCommon.changestype_2_ElementsArray_and_Type,
				Elements  : [oContentControl],
				CheckType : AscCommon.changestype_ContentControl_Properties
			});
		}
		else if (c_oAscSdtLevelType.Inline === oContentControl.GetContentControlType())
		{
			var oParagraph = oContentControl.GetParagraph();
			if (oParagraph)
			{
				isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
					Type      : AscCommon.changestype_2_ElementsArray_and_Type,
					Elements  : [oParagraph],
					CheckType : AscCommon.changestype_Paragraph_Properties
				});
			}
		}

		if (false === isLocked)
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_SetContentControlListPr);

			if (oContentControl.IsComboBox())
				oContentControl.SetComboBoxPr(oPr);
			else
				oContentControl.SetDropDownListPr(oPr);

			oLogicDocument.UpdateInterface();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_SelectContentControlListItem = function(sValue, sId)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl || (!oContentControl.IsComboBox() && !oContentControl.IsDropDownList()))
			return;

		oContentControl.SkipSpecialContentControlLock(true);

		if (!oContentControl.CanBeEdited())
		{
			oContentControl.SkipSpecialContentControlLock(false);
			return;
		}

		var isLocked = false;
		if (c_oAscSdtLevelType.Block === oContentControl.GetContentControlType())
		{
			isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
				Type      : AscCommon.changestype_2_ElementsArray_and_Type,
				Elements  : [oContentControl],
				CheckType : AscCommon.changestype_Paragraph_AddText
			}, false, oLogicDocument.IsFormFieldEditing());
		}
		else if (c_oAscSdtLevelType.Inline === oContentControl.GetContentControlType())
		{
			var oParagraph = oContentControl.GetParagraph();
			if (oParagraph)
			{
				var oState = oLogicDocument.SaveDocumentState();
				oContentControl.SelectContentControl();

				isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
					Type      : AscCommon.changestype_2_ElementsArray_and_Type,
					Elements  : [oParagraph],
					CheckType : AscCommon.changestype_Paragraph_AddText
				}, false, oLogicDocument.IsFormFieldEditing());

				oLogicDocument.LoadDocumentState(oState);
			}
		}
		oContentControl.SkipSpecialContentControlLock(false);

		if (!isLocked)
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_SelectContentControlListItem);
			oContentControl.SelectListItem(sValue);
			oLogicDocument.RemoveSelection();
			oContentControl.MoveCursorToContentControl(true);
			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.UpdateTracks();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_GetContentControlListCurrentValue = function(sId)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return null;

		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl || (!oContentControl.IsComboBox() && !oContentControl.IsDropDownList()) || oContentControl.IsPlaceHolder())
			return null;

		var oState = oLogicDocument.SaveDocumentState();

		oLogicDocument.RemoveSelection();
		oContentControl.SelectContentControl();

		var oList  = oContentControl.IsDropDownList() ? oContentControl.GetDropDownListPr() : oContentControl.GetComboBoxPr();
		var nIndex = oList.FindByText(oContentControl.GetSelectedText());

		if (-1 === nIndex)
			return undefined;

		var sValue = oList.GetItemValue(nIndex);

		oLogicDocument.LoadDocumentState(oState);
		return sValue;
	};
	asc_docs_api.prototype.asc_SetContentControlDatePickerPr = function(oPr, sId, updateDate)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var isLocked        = true;
		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl || !oContentControl.IsDatePicker())
			return;
		
		oContentControl.SkipSpecialContentControlLock(true);
		oContentControl.SkipFillingFormModeCheck(true);
		if (c_oAscSdtLevelType.Block === oContentControl.GetContentControlType())
		{
			isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
				Type      : AscCommon.changestype_2_ElementsArray_and_Type,
				Elements  : [oContentControl],
				CheckType : AscCommon.changestype_ContentControl_Properties
			}, false, oLogicDocument.IsFormFieldEditing());
		}
		else if (c_oAscSdtLevelType.Inline === oContentControl.GetContentControlType())
		{
			var oParagraph = oContentControl.GetParagraph();
			if (oParagraph)
			{
				isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
					Type      : AscCommon.changestype_2_ElementsArray_and_Type,
					Elements  : [oParagraph],
					CheckType : AscCommon.changestype_Paragraph_Properties
				}, false, oLogicDocument.IsFormFieldEditing());
			}
		}
		
		oContentControl.SkipSpecialContentControlLock(false);
		oContentControl.SkipFillingFormModeCheck(false);
		
		if (isLocked)
			return;
		
		oLogicDocument.StartAction(AscDFH.historydescription_Document_SetContentControlListPr);
		oContentControl.ApplyDatePickerPr(oPr, updateDate);
		oContentControl.MoveCursorToContentControl(false);
		oLogicDocument.Recalculate();
		oLogicDocument.UpdateInterface();
		oLogicDocument.UpdateTracks();
		oLogicDocument.UpdateSelection();
		oLogicDocument.FinalizeAction();
	};
	asc_docs_api.prototype.asc_SetContentControlDatePickerDate = function(oPr, sId)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl || !oContentControl.IsDatePicker() || !oContentControl.CanBeEdited())
			return;
		
		this.asc_SetContentControlDatePickerPr(oPr, sId, true);
	};
	asc_docs_api.prototype.asc_SetContentControlTextPlaceholder = function(sText, sId)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl)
			return;

		oLogicDocument.SetContentControlTextPlaceholder(sText, oContentControl);
	};
	asc_docs_api.prototype.asc_SetContentControlText = function(sText, sId)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl)
			return;

		oLogicDocument.SetContentControlText(sText, oContentControl);
	};
	asc_docs_api.prototype.asc_GetTextFormKeys = function()
	{
		let oManager = this.private_GetFormsManager();
		return oManager ? oManager.GetAllKeys({Text : true, ComboBox : true, DropDownList : true}) : [];
	};
	asc_docs_api.prototype.asc_GetPictureFormKeys = function()
	{
		let oManager = this.private_GetFormsManager();
		return oManager ? oManager.GetAllKeys({Picture : true}) : [];
	};
	asc_docs_api.prototype.asc_GetCheckBoxFormKeys = function()
	{
		let oManager = this.private_GetFormsManager();
		return oManager ? oManager.GetAllKeys({CheckBox : true}) : [];
	};
	asc_docs_api.prototype.asc_GetRadioButtonGroupKeys = function()
	{
		let oManager = this.private_GetFormsManager();
		return oManager ? oManager.GetAllKeys({RadioGroup : true}) : [];
	};
	asc_docs_api.prototype.asc_GetFormKeysByType = function(type)
	{
		let formManager = this.private_GetFormsManager();
		if (!formManager)
			return [];

		let pr = {};
		switch (type)
		{
			case Asc.c_oAscContentControlSpecificType.Picture:
				pr.Picture = true;
				break;
			case Asc.c_oAscContentControlSpecificType.CheckBox:
				pr.CheckBox = true;
				break;
			case Asc.c_oAscContentControlSpecificType.Complex:
				pr.Complex = true;
				break;
			default:
				pr.Text         = true;
				pr.ComboBox     = true;
				pr.DropDownList = true;
				break;
		}

		return formManager.GetAllKeys(pr);
	};
	asc_docs_api.prototype.asc_ClearAllSpecialForms = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		oLogicDocument.ClearAllSpecialForms(true);
	};
	asc_docs_api.prototype.asc_SetSpecialFormsHighlightColor = function(r, g, b)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument || !oLogicDocument.CanPerformAction())
			return;

		// Лок можно не проверять, такие изменения нормально мержаться
		oLogicDocument.StartAction(AscDFH.historydescription_Document_SetGlobalSdtHighlightColor);

		// Если цвет не задан
		if (undefined === r || null == r)
			oLogicDocument.SetSpecialFormsHighlight(null);
		else
			oLogicDocument.SetSpecialFormsHighlight(r, g, b);

		oLogicDocument.GetDrawingDocument().ClearCachePages();
		oLogicDocument.GetDrawingDocument().FirePaint();
		oLogicDocument.FinalizeAction();
	};
	asc_docs_api.prototype.asc_GetSpecialFormsHighlightColor = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return null;

		var oColor = oLogicDocument.GetSpecialFormsHighlight();
		return oColor && !oColor.IsAuto() ? new Asc.asc_CColor(oColor.r, oColor.g, oColor.b) : null;
	};
	asc_docs_api.prototype.sync_OnChangeSpecialFormsGlobalSettings = function()
	{
		this.sendEvent("asc_onChangeSpecialFormsGlobalSettings");
	};
	asc_docs_api.prototype.asc_SetPerformContentControlActionByClick = function(isPerform)
	{
		// obsolete function
	};
	asc_docs_api.prototype.asc_GetTextFormAutoWidth = function(sId)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl || !oContentControl.IsTextForm() || !oContentControl.IsInlineLevel())
			return null;

		var oRun = oContentControl.GetFirstRun();
		if (!oRun)
			return null;

		return AscCommon.MMToTwips(oRun.GetTextFormAutoWidth());
	};
	asc_docs_api.prototype.asc_GetFormsCountByKey = function(sKey)
	{
		let formManager = this.private_GetFormsManager();
		if (!formManager)
			return 0;
		
		return formManager.GetAllFormsByKey(sKey).length + formManager.GetRadioButtons(sKey).length;
	};
	asc_docs_api.prototype.asc_MoveToFillingForm = function(isNext, isRequired, isNotFilled)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		if (true === isRequired)
		{
			var oStartCC  = oLogicDocument.GetContentControl();
			var oDocState = oLogicDocument.SaveDocumentState(false);

			// Защита от зависания, по логике сравнение с предыдущим не нужно
			var oPrevCC = null, oCurCC;
			while (1)
			{
				oLogicDocument.MoveToFillingForm(isNext);
				oCurCC = oLogicDocument.GetContentControl();

				if (!oCurCC || oCurCC === oStartCC || oCurCC === oPrevCC)
				{
					oLogicDocument.LoadDocumentState(oDocState);
					break;
				}

				if (oCurCC.IsForm()
					&& oCurCC.IsFormRequired()
					&& !oCurCC.IsCheckBox()
					&& (true !== isNotFilled || !oCurCC.IsFormFilled()))
					break;

				if (!oStartCC)
					oStartCC = oCurCC;

				oPrevCC = oCurCC;
			}
		}
		else
		{
			oLogicDocument.MoveToFillingForm(isNext);
		}

		oLogicDocument.UpdateSelection();
		oLogicDocument.UpdateInterface();
	};
	asc_docs_api.prototype.asc_IsAllRequiredFormsFilled = function(checkAll)
	{
		let oFormsManager = this.private_GetFormsManager();
		let oform = this.asc_GetOForm();
		if (!oFormsManager)
			return true;
		
		let roleName = !checkAll && oform ? oform.getCurrentRole() : null;
		return oFormsManager.IsAllRequiredFormsFilled(roleName);
	};
	asc_docs_api.prototype.sync_OnAllRequiredFormsFilled = function(isFilled)
	{
		this.sendEvent("sync_onAllRequiredFormsFilled", isFilled);
	};
	asc_docs_api.prototype.asc_SetFixedForm = function(sId, isFixed)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		oLogicDocument.ConvertFormFixedType(sId, isFixed);
	};
	asc_docs_api.prototype.asc_IsHighlightRequiredFields = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return false;

		return oLogicDocument.IsHighlightRequiredFields();
	};
	asc_docs_api.prototype.asc_SetHighlightRequiredFields = function(isHighlight)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		oLogicDocument.SetHighlightRequiredFields(isHighlight);
		oLogicDocument.DrawingDocument.ClearCachePages();
		oLogicDocument.DrawingDocument.FirePaint();
	}
	asc_docs_api.prototype.asc_OnChangeContentControl = function(oControl)
	{
		if (!oControl)
			return;

		this.sendEvent("onChangeContentControl", oControl.GetContentControlPr().GetEventObject());

		if (window.g_asc_plugins)
			window.g_asc_plugins.onPluginEvent("onChangeContentControl", oControl.GetContentControlPr().GetEventObject());
	};
	asc_docs_api.prototype.asc_OnFocusContentControl = function(oControl)
	{
		if (!oControl)
			return;

		this.sendEvent("onFocusContentControl", oControl.GetContentControlPr().GetEventObject());

		if (window.g_asc_plugins)
			window.g_asc_plugins.onPluginEvent("onFocusContentControl", oControl.GetContentControlPr().GetEventObject());
	};
	asc_docs_api.prototype.asc_OnBlurContentControl = function(oControl)
	{
		if (!oControl)
			return;

		this.sendEvent("onBlurContentControl", oControl.GetContentControlPr().GetEventObject());

		if (window.g_asc_plugins)
			window.g_asc_plugins.onPluginEvent("onBlurContentControl", oControl.GetContentControlPr().GetEventObject());
	};
	asc_docs_api.prototype.asc_GetAllFormsData = function()
	{
		let formManager = this.private_GetFormsManager();
		if (!formManager)
			return [];

		return formManager.GetAllFormsData();
	};
	asc_docs_api.prototype.asc_SetAllFormsData = function(data)
	{
		let logicDocument = this.private_GetLogicDocument();
		let formManager   = this.private_GetFormsManager();
		if (!formManager || !logicDocument)
			return;
		
		// TODO: For now we assume that there is no need to check for document selection lock
		logicDocument.StartAction(AscDFH.historydescription_Document_SetAllFormsData);
		formManager.SetAllFormsData(data);
		logicDocument.Recalculate();
		logicDocument.FinalizeAction();
	};
	asc_docs_api.prototype.asc_GetOForm = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return null;
		
		return logicDocument.GetOFormDocument();
	};
	asc_docs_api.prototype.asc_DisconnectEveryone = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;
		
		if (logicDocument.IsSelectionLocked(AscCommon.changestype_Document_Settings))
			return false;
		
		logicDocument.StartAction(AscDFH.historydescription_DisconnectEveryone);
		logicDocument.GetHistory().Add(new CChangesDocumentDisconnectEveryone(logicDocument));
		logicDocument.FinalizeAction();
		
		this.forceSaveDisconnectRequest = true;
		AscCommon.CollaborativeEditing.Set_GlobalLock(true);
		this.asc_Save(false);
	};
	asc_docs_api.prototype.asc_CompletePreparingOForm = function(disconnect)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;
		
		if (logicDocument.IsSelectionLocked(AscCommon.changestype_Document_Settings))
			return false;
		
		logicDocument.StartAction(AscDFH.historydescription_OForm_CompletePreparation);
		logicDocument.GetOFormDocument().setAllRolesNotFilled();
		if (disconnect)
			logicDocument.GetHistory().Add(new CChangesDocumentDisconnectEveryone(logicDocument));
		
		logicDocument.FinalizeAction();
		
		if (disconnect)
		{
			this.forceSaveDisconnectRequest = true;
			AscCommon.CollaborativeEditing.Set_GlobalLock(true);
			this.asc_Save(false);
		}
	};
	

	asc_docs_api.prototype.asc_UncheckContentControlButtons = function()
	{
		var _controls = (this.WordControl && this.WordControl.m_oDrawingDocument && this.WordControl.m_oDrawingDocument.contentControls) ? this.WordControl.m_oDrawingDocument.contentControls.ContentControlObjects : [];
		if (_controls.length === 0)
			return;

		if (!window['IS_NATIVE_EDITOR'] && AscCommon.global_mouseEvent.IsLocked)
		{
			this.WordControl.onMouseUpMainSimple();
		}

		for (var i = 0; i < _controls.length; i++)
		{
			_controls[i].ActiveButtonIndex = -2;
			_controls[i].HoverButtonIndex = -2;
		}

		this.WordControl.ShowOverlay();
		this.WordControl.StartUpdateOverlay();
		this.WordControl.OnUpdateOverlay();
		this.WordControl.EndUpdateOverlay();
	};

	asc_docs_api.prototype.asc_BeginViewModeInReview = function(isFinal)
	{
		this.asc_SetDisplayModeInReview(isFinal ? Asc.c_oAscDisplayModeInReview.Final : Asc.c_oAscDisplayModeInReview.Original);
	};
	asc_docs_api.prototype.asc_EndViewModeInReview = function()
	{
		this.asc_SetDisplayModeInReview(Asc.c_oAscDisplayModeInReview.Edit);
	};
	asc_docs_api.prototype.asc_SetDisplayModeInReview = function(nMode)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (oLogicDocument)
			oLogicDocument.SetDisplayModeInReview(nMode);
	};
	asc_docs_api.prototype.asc_GetDisplayModeInReview = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (oLogicDocument)
			return oLogicDocument.GetDisplayModeInReview();

		return Asc.c_oAscDisplayModeInReview.Edit;
	};

	asc_docs_api.prototype.asc_GetDocumentOutlineManager = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return null;

		return oLogicDocument.DocumentOutline;
	};
	asc_docs_api.prototype.asc_ShowDocumentOutline = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return null;

		oLogicDocument.DocumentOutline.SetUse(true);
		return oLogicDocument.DocumentOutline;
	};
	asc_docs_api.prototype.asc_HideDocumentOutline = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		oLogicDocument.DocumentOutline.SetUse(false);
	};
	asc_docs_api.prototype.sync_OnDocumentOutlineUpdate = function()
	{
		this.sendEvent("asc_onDocumentOutlineUpdate");
	};
	asc_docs_api.prototype.sync_OnDocumentOutlineCurrentPosition = function(nIndex)
	{
		this.sendEvent("asc_onDocumentOutlineCurrentPosition", nIndex);
	};
	asc_docs_api.prototype.sync_OnDocumentOutlineUpdateAdd = function(nIndex)
	{
		this.sendEvent("asc_onDocumentOutlineUpdateAdd", nIndex);
	};
	asc_docs_api.prototype.sync_OnDocumentOutlineUpdateChange = function(nIndex)
	{
		this.sendEvent("asc_onDocumentOutlineUpdateChange", nIndex);
	};
	asc_docs_api.prototype.sync_OnDocumentOutlineUpdateRemove = function(nIndex)
	{
		this.sendEvent("asc_onDocumentOutlineUpdateRemove", nIndex);
	};

	asc_docs_api.prototype.asc_AddTableOfContents = function(sHeading, oPr)
	{
        var sReplacementText = AscCommon.translateManager.getValue("No table of contents entries found.");
		AscFonts.FontPickerByCharacter.checkText(sHeading + sReplacementText, this, function() {

			var oLogicDocument = this.WordControl.m_oLogicDocument;
			if (!oLogicDocument)
				return;

			var oTOC = oLogicDocument.GetTableOfContents();
			if (oTOC instanceof AscCommonWord.CBlockLevelSdt && oTOC.IsBuiltInUnique())
			{
				if (oPr)
					this.asc_SetTableOfContentsPr(oPr);

				return;
			}

			this.WordControl.m_oLogicDocument.AddTableOfContents(sHeading, oPr);

		});
	};
	asc_docs_api.prototype.asc_RemoveTableOfContents = function(oTOC)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		if (!oTOC)
		{
			oTOC = oLogicDocument.GetTableOfContents();
			if (!oTOC)
				return;
		}

		if (oTOC instanceof AscCommonWord.CBlockLevelSdt)
		{
			this.asc_RemoveContentControl(oTOC.GetId());
		}
		else if (oTOC instanceof AscCommonWord.CComplexField)
		{
			if (!oTOC.IsValid())
				return;

			var oCF = oTOC;

			oTOC = null;

			var oBeginChar = oCF.GetBeginChar();
			if (oBeginChar && oBeginChar.GetRun())
			{
				var oParentCCs = oBeginChar.GetRun().GetParentContentControls();
				for (var nIndex = 0, nCount = oParentCCs.length; nIndex < nCount; ++nIndex)
				{
					if (oParentCCs[nIndex].IsBlockLevel() && oParentCCs[nIndex].IsBuiltInTableOfContents())
					{
						oTOC = oParentCCs[nIndex];
						break;
					}
				}
			}

			if (oTOC)
				this.asc_RemoveContentControl(oTOC.GetId());
			else
				this.asc_RemoveComplexField(oCF);
		}
	};
	asc_docs_api.prototype.asc_GetTableOfContentsPr = function(isCurrent)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return null;

		var oTOC = oLogicDocument.GetTableOfContents(isCurrent);
		if (!oTOC)
			return null;

		if (oTOC instanceof AscCommonWord.CBlockLevelSdt)
		{
			var oInnerTOC = oTOC.GetInnerTableOfContents();

			if (!(oInnerTOC instanceof AscCommonWord.CComplexField))
			{
				// Специальнный случай, когда у нас есть контейнер помеченный как TOC, но самого TOC внутри нет
				// посылаем в интерфейс класс так, чтобы с ним все равно можно было работать (например, удалить его)
				var oPr = new Asc.CTableOfContentsPr();
				oPr.InitFromSdtTOC(oTOC);
				return oPr;
			}

			oTOC = oInnerTOC;
		}

		if (oTOC instanceof AscCommonWord.CComplexField)
		{
			var oPr = new Asc.CTableOfContentsPr();
			oPr.InitFromTOCInstruction(oTOC);
			oPr.CheckStylesType(oLogicDocument.GetStyles());
			return oPr;
		}

		return null;
	};
	asc_docs_api.prototype.asc_SetTableOfContentsPr = function(oPr)
	{
		if (!(oPr instanceof Asc.CTableOfContentsPr))
			return;

		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		var oTOC = oPr.ComplexField;
		if (!oTOC
			|| !(oTOC instanceof AscCommonWord.CComplexField)
			|| !oTOC.IsValid())
		{
			oTOC = oLogicDocument.GetTableOfContents();
			if (!oTOC)
				return;
		}

		if (oTOC instanceof AscCommonWord.CBlockLevelSdt)
		{
			var oInnerTOC = oTOC.GetInnerTableOfContents();
			if (!oInnerTOC)
			{
				oLogicDocument.AddTableOfContents(null, oPr, oTOC);
				return;
			}

			oTOC = oInnerTOC;
		}

		if (!oTOC)
			return;

		var oStyles     = oLogicDocument.GetStyles();
		var nStylesType = oPr.get_StylesType();

		var isTOF = oPr.get_Caption() !== undefined;

		var isNeedChangeStyles = false;

		if(isTOF)
		{
			isNeedChangeStyles = (Asc.c_oAscTOFStylesType.Current !== nStylesType && nStylesType !== oStyles.GetTOFStyleType());
		}
		else
		{
			isNeedChangeStyles = (Asc.c_oAscTOCStylesType.Current !== nStylesType && nStylesType !== oStyles.GetTOCStylesType());
		}

		oTOC.SelectField();

		var isLocked = true;
		if (isNeedChangeStyles)
		{
			isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content, {
				Type : AscCommon.changestype_2_AdditionalTypes, Types : [AscCommon.changestype_Document_Styles]
			});
		}
		else
		{
			isLocked = oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content)
		}

		if (!isLocked)
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_SetComplexFieldPr);

			if (isNeedChangeStyles)
			{
				if(isTOF)
				{
					oStyles.SetTOFStyleType(nStylesType);
				}
				else
				{
					oStyles.SetTOCStylesType(nStylesType);
				}
			}

			oTOC.SetPr(oPr);
			oTOC.Update();

			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.UpdateSelection();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_UpdateTableOfContents = function(isUpdatePageNumbers, oTOC)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return null;

		if (!oTOC)
			oTOC = oLogicDocument.GetTableOfContents();

		if (oTOC instanceof AscCommonWord.CBlockLevelSdt)
			oTOC = oTOC.GetInnerTableOfContents();

		if (!oTOC
			|| !(oTOC instanceof AscCommonWord.CComplexField)
			|| !oTOC.IsValid())
		{
			this.sendEvent("asc_onError", c_oAscError.ID.ComplexFieldNoTOC, c_oAscError.Level.NoCritical);
			return;
		}

		var oState = oLogicDocument.SaveDocumentState();

		oTOC.SelectField();
		if (isUpdatePageNumbers)
		{
			if (false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content))
			{
				oLogicDocument.StartAction(AscDFH.historydescription_Document_UpdateTableOfContents);

				var arrParagraphs = oLogicDocument.GetCurrentParagraph(false, true);

				for (var nParaIndex = 0, nParasCount = arrParagraphs.length; nParaIndex < nParasCount; ++nParaIndex)
				{
					var arrPageNumbers = arrParagraphs[nParaIndex].GetComplexFieldsArrayByType(AscWord.fieldtype_PAGEREF);
					for (var nRefIndex = 0, nRefsCount = arrPageNumbers.length; nRefIndex < nRefsCount; ++nRefIndex)
					{
						arrPageNumbers[nRefIndex].Update();
					}
				}

				oLogicDocument.LoadDocumentState(oState);
				oLogicDocument.Recalculate();
				oLogicDocument.UpdateInterface();
				oLogicDocument.UpdateSelection();
				oLogicDocument.FinalizeAction();
			}
		}
		else
		{
			if (false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Document_Content))
			{
				oLogicDocument.StartAction(AscDFH.historydescription_Document_UpdateTableOfContents);

				oTOC.Update();

				oLogicDocument.LoadDocumentState(oState);
				oLogicDocument.Recalculate();
				oLogicDocument.UpdateInterface();
				oLogicDocument.UpdateSelection();
				oLogicDocument.FinalizeAction();
			}
		}

	};
	asc_docs_api.prototype.asc_AddParagraphToTOC = function(nLvl)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		let arrParagraphs = oLogicDocument.GetSelectedParagraphs();
		if (arrParagraphs.length < 0)
			return;

		if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_None, {
			Type      : AscCommon.changestype_2_ElementsArray_and_Type,
			Elements  : arrParagraphs,
			CheckType : AscCommon.changestype_Paragraph_Properties
		}, false, false))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_AddParagraphToTOC);

			let oStyles = oLogicDocument.GetStyles();
			if (-1 === nLvl || undefined === nLvl || null === nLvl)
			{
				for (let nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
				{
					let oParagraph = arrParagraphs[nIndex];
					oParagraph.SetOutlineLvl(null);

					let sCurStyleId = oParagraph.GetParagraphStyle();
					if (-1 !== oStyles.GetHeadingLevelById(sCurStyleId))
						oParagraph.SetParagraphStyleById(null);
				}
			}
			else
			{
				let sStyleId = oStyles.GetDefaultHeading((nLvl | 0));
				for (let nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
				{
					arrParagraphs[nIndex].SetParagraphStyleById(sStyleId);
				}
			}

			oLogicDocument.Recalculate();
			oLogicDocument.UpdateSelection();
			oLogicDocument.UpdateInterface();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_GetCurrentLevelTOC = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return -1;

		let arrParagraphs = oLogicDocument.GetSelectedParagraphs();
		let nResult = -1;

		for (let nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
		{
			let nCurrentLvl = arrParagraphs[nIndex].GetOutlineLvl();

			if (undefined !== nCurrentLvl
				&& null !== nCurrentLvl
				&& nCurrentLvl >= 0
				&& (-1 === nResult || nResult > nCurrentLvl))
				nResult = nCurrentLvl;
		}

		return nResult;
	};

	asc_docs_api.prototype.asc_CanUpdateTablesOfFigures = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		var aTOF = oLogicDocument.GetAllTablesOfFigures(true);
		return aTOF.length > 0;
	};
	asc_docs_api.prototype.asc_UpdateTablesOfFigures = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		var aTOF = oLogicDocument.GetAllTablesOfFigures(true);
		if(aTOF.length === 0)
		{
			return;
		}
		var oApi = this;
		var fCallback = function (isUpdatePageNumbers)
		{
			var oTOF = aTOF.pop();
			if(oTOF)
			{
				oApi.asc_UpdateTableOfContents(isUpdatePageNumbers, oTOF);
			}
			if(aTOF.length > 0)
			{
				aTOF[aTOF.length - 1].SelectField();
				oApi.sendEvent("asc_onAscTOFUpdate", fCallback);
			}
		};
		aTOF[aTOF.length - 1].SelectField();
		oApi.sendEvent("asc_onAscTOFUpdate", fCallback);
	};

	asc_docs_api.prototype.asc_AddTableOfFigures = function(oPr)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		var oApi = this;
		if(!oLogicDocument)
		{
			return;
		}
		var sCaption = oPr.get_CaptionForInstruction();
		var aTOF, oTOF, nIndex, sInstrCaption;
		var oTOFToReplace = null;
		var oInstruction;
		aTOF = oLogicDocument.GetAllTablesOfFigures(true);
		if(aTOF.length > 0)
		{
			oTOFToReplace = aTOF[0];
		}
		if(!oTOFToReplace)
		{
			aTOF = oLogicDocument.GetAllTablesOfFigures();
			if(aTOF.length > 0)
			{
				if(typeof sCaption === "string" && sCaption.length > 0)
				{
					for(nIndex = 0; nIndex < aTOF.length; ++nIndex)
					{
						oTOF = aTOF[nIndex];
						oInstruction = oTOF.GetInstruction();
						if(!oInstruction)
						{
							continue;
						}
						sInstrCaption = oInstruction.GetCaptionOnlyText() || oInstruction.GetCaption();
						if(sInstrCaption === sCaption)
						{
							oTOFToReplace = oTOF;
							break;
						}
					}
				}
			}
		}
		if(oTOFToReplace)
		{
			var oState = oLogicDocument.SaveDocumentState();
			oTOFToReplace.SelectFieldValue();
			this.sendEvent("asc_onAscReplaceCurrentTOF", function(bReplace){
				if(!bReplace)
				{
					oLogicDocument.MoveCursorRight(false, false, false);
					oLogicDocument.AddTableOfFigures(oPr);
				}
				else
				{
					oPr.ComplexField = oTOFToReplace;
					oLogicDocument.LoadDocumentState(oState);
					oApi.asc_SetTableOfContentsPr(oPr);
				}
			});
		}
		else
		{
			oLogicDocument.AddTableOfFigures(oPr);
		}
	};
	asc_docs_api.prototype.asc_GetTableOfFiguresPr = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
		{
			return null;
		}
		var aTOF = oLogicDocument.GetAllTablesOfFigures(true);
		if(aTOF.length > 0)
		{
			var oPr = new Asc.CTableOfContentsPr();
			oPr.InitFromTOCInstruction(aTOF[0]);
			oPr.CheckStylesType(oLogicDocument.GetStyles());
			return oPr;
		}
		return null;
	};

	asc_docs_api.prototype.asc_AddComplexFieldWithInstruction = function(instruction)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		return logicDocument.AddComplexField(instruction);
	};
	asc_docs_api.prototype.asc_EditComplexFieldInstruction = function(instruction)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;
		
		let complexField = logicDocument.GetCurrentComplexField();
		return logicDocument.EditComplexFieldInstruction(complexField, instruction);
	};
	asc_docs_api.prototype.asc_GetComplexFieldInstruction = function()
	{
		let complexField = this.asc_GetCurrentComplexField();
		if (!complexField || !(complexField instanceof AscWord.CComplexField))
			return "";
		
		return complexField.GetInstructionLine();
	};
	asc_docs_api.prototype.asc_GetCurrentComplexField = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return null;

		return oLogicDocument.GetCurrentComplexField();
	};
	asc_docs_api.prototype.asc_UpdateComplexField = function(oComplexField)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument
			|| !(oComplexField instanceof AscCommonWord.CComplexField
				|| oComplexField instanceof AscWord.CSimpleField
				|| oComplexField instanceof AscWord.CRunPageNum
				|| oComplexField instanceof AscWord.CRunPagesCount))
			return;

		if (oComplexField instanceof AscWord.CRunPageNum || oComplexField instanceof AscWord.CRunPagesCount)
		{
			var oRun = oComplexField.GetParent();
			if (!oRun)
				return;

			var nInRunPos = oRun.GetElementPosition(oComplexField);
			if (-1 === nInRunPos)
				return;

			var oParagraph = oRun.GetParagraph();
			if (!oParagraph)
				return;

			if (false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
					Type      : AscCommon.changestype_2_ElementsArray_and_Type,
					Elements  : [oParagraph],
					CheckType : AscCommon.changestype_Paragraph_Content
				}))
			{
				oLogicDocument.StartAction(AscDFH.historydescription_Document_SetComplexFieldPr);

				oRun.RemoveFromContent(nInRunPos, 1);
				oRun.AddToContent(nInRunPos, oComplexField instanceof AscWord.CRunPageNum ? new AscWord.CRunPageNum() : new AscWord.CRunPagesCount(oLogicDocument.GetPagesCount()));

				oLogicDocument.Recalculate();
				oLogicDocument.UpdateInterface();
				oLogicDocument.UpdateSelection();
				oLogicDocument.FinalizeAction();
			}
		}
		else
		{
			oComplexField.SelectField();
			if (false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content))
			{
				oLogicDocument.StartAction(AscDFH.historydescription_Document_SetComplexFieldPr);

				oComplexField.Update();

				oLogicDocument.Recalculate();
				oLogicDocument.UpdateInterface();
				oLogicDocument.UpdateSelection();
				oLogicDocument.FinalizeAction();
			}
		}
	};
	asc_docs_api.prototype.asc_RemoveComplexField = function(oComplexField)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oComplexField || !oLogicDocument)
			return;

		if (oComplexField instanceof AscWord.CRunPageNum || oComplexField instanceof AscWord.CRunPagesCount)
		{
			var oRun = oComplexField.GetParent();
			if (!oRun)
				return;

			var nInRunPos = oRun.GetElementPosition(oComplexField);
			if (-1 === nInRunPos)
				return;

			var oParagraph = oRun.GetParagraph();
			if (!oParagraph)
				return;

			if (false === oLogicDocument.Document_Is_SelectionLocked({
					Type      : AscCommon.changestype_2_ElementsArray_and_Type,
					Elements  : [oParagraph],
					CheckType : AscCommon.changestype_Paragraph_Content
				}))
			{
				oLogicDocument.StartAction(AscDFH.historydescription_Document_SetComplexFieldPr);

				oRun.RemoveFromContent(nInRunPos, 1);

				oLogicDocument.Recalculate();
				oLogicDocument.UpdateInterface();
				oLogicDocument.UpdateSelection();
				oLogicDocument.FinalizeAction();
			}
		}
		else
		{
			oComplexField.SelectField();
			if (false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Remove))
			{
				oLogicDocument.StartAction(AscDFH.historydescription_Document_RemoveComplexField);

				oComplexField.RemoveField();

				oLogicDocument.Recalculate();
				oLogicDocument.UpdateInterface();
				oLogicDocument.UpdateSelection();
				oLogicDocument.FinalizeAction();
			}
		}
	};
	asc_docs_api.prototype.asc_SetComplexFieldPr = function(oComplexField, oPr, isUpdate)
	{

		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument || !oPr || !(oComplexField instanceof AscCommonWord.CComplexField || oComplexField instanceof AscWord.CRunPageNum || oComplexField instanceof AscWord.CRunPagesCount))
			return;

		if (oComplexField instanceof AscWord.CRunPageNum || oComplexField instanceof AscWord.CRunPagesCount)
			return;

		oComplexField.SelectField();
		if (false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_SetComplexFieldPr);

			oComplexField.SetPr(oPr);
			if (isUpdate)
				oComplexField.Update();

			oLogicDocument.Recalculate();
			oLogicDocument.UpdateInterface();
			oLogicDocument.UpdateSelection();
			oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.asc_AddTableFormula = function(sFormula)
	{
		var oLogicDocument = this.private_GetLogicDocument();

		if (!oLogicDocument)
			return;

		oLogicDocument.AddTableCellFormula(sFormula);
	};
	asc_docs_api.prototype.asc_GetTableFormula = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return "=";

		return oLogicDocument.GetTableCellFormula();
	};
	asc_docs_api.prototype.asc_GetTableFormulaFormats = function()
	{
		return ["#,##0", "#,##0.00", "$#,##0.00;($#,##0.00)", "0", "0%", "0.00", "0.00%"];
	};
	asc_docs_api.prototype.asc_HaveFields = function(isInSelection)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;

		return (logicDocument.GetAllFields(isInSelection).length > 0);
	};
	asc_docs_api.prototype.asc_UpdateFields = function(isInSelection)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;

		return logicDocument.UpdateFields(isInSelection);
	};
	asc_docs_api.prototype.asc_ToggleComplexFieldCodes = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		logicDocument && logicDocument.ToggleComplexFieldCodes()
	};

	asc_docs_api.prototype.asc_ParseTableFormulaInstrLine = function(sInstrLine)
	{
		return this.WordControl.m_oLogicDocument.ParseTableFormulaInstrLine(sInstrLine);
	};

	asc_docs_api.prototype.asc_CreateInstructionLine = function(sFormula, sFormat)
	{
		var sRet = sFormula;
		if(typeof sFormat === "string" && sFormat.length > 0){
			sRet += " \\# \"" + sFormat + "\"";
		}
		return sRet;
	};

	asc_docs_api.prototype.asc_addDateTime = function(oPr)
	{
		var sTextForCheck = null;
		if (!oPr.get_Update())
			sTextForCheck = oPr.get_String();

		var oLogicDocument = this.private_GetLogicDocument();
		if (sTextForCheck)
		{
			AscFonts.FontPickerByCharacter.checkText(sTextForCheck, this, function()
			{
				oLogicDocument.AddDateTime(oPr);
			});
		}
		else
		{
			oLogicDocument.AddDateTime(oPr);
		}
	};

	asc_docs_api.prototype.asc_AddObjectCaption = function(oPr)
	{
		var oLogicDocument = this.private_GetLogicDocument();

		if (!oLogicDocument)
			return;

		oLogicDocument.AddCaption(oPr);
	};

	asc_docs_api.prototype.asc_GetBookmarksManager = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return null;

		var oManager = oLogicDocument.GetBookmarksManager();
		oManager.Update();
		return oManager;
	};
	asc_docs_api.prototype.asc_OnBookmarksUpdate = function()
	{
		this.sendEvent("asc_onBookmarksUpdate");
	};

	asc_docs_api.prototype.asc_OnProtectionUpdate = function(sUserId)
	{
		this.sendEvent("asc_onChangeDocumentProtection", sUserId);
	};

	asc_docs_api.prototype.asc_GetHeadingLevel = function(sStyleName)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument || !sStyleName)
			return -1;

		return oLogicDocument.GetStyles().GetHeadingLevelByName(sStyleName);
	};
	asc_docs_api.prototype.asc_GetStylesArray = function()
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return [];

		return oLogicDocument.GetStyles().GetAscStylesArray();
	};
	asc_docs_api.prototype.asc_getAllUsedParagraphStyles = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return [];
		var aStyles = oLogicDocument.GetAllUsedParagraphStyles();
		var aAscStyles = [];
		for(var nStyle = 0; nStyle < aStyles.length; ++nStyle)
		{
			aAscStyles.push(aStyles[nStyle].ToAscStyle());
		}
		return aAscStyles;
	};

	asc_docs_api.prototype.asc_SetAutomaticBulletedLists = function(isAuto)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutomaticBulletedLists(isAuto);
	};
	asc_docs_api.prototype.asc_SetAutomaticNumberedLists = function(isAuto)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutomaticNumberedLists(isAuto);
	};
	asc_docs_api.prototype.asc_SetAutoCorrectSmartQuotes = function(isSmartQuotes)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutoCorrectSmartQuotes(isSmartQuotes);
	};
	asc_docs_api.prototype.asc_SetAutoCorrectHyphensWithDash = function(isReplace)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutoCorrectHyphensWithDash(isReplace);
	};
	asc_docs_api.prototype.asc_SetAutoCorrectFirstLetterOfSentences = function(isCorrect)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutoCorrectFirstLetterOfSentences(isCorrect);
	};
	asc_docs_api.prototype.asc_SetAutoCorrectHyperlinks = function(isCorrect)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutoCorrectHyperlinks(isCorrect);
	};
	asc_docs_api.prototype.asc_SetAutoCorrectFirstLetterOfCells = function(isCorrect)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutoCorrectFirstLetterOfCells(isCorrect);
	};
	asc_docs_api.prototype.asc_SetAutoCorrectDoubleSpaceWithPeriod = function(isCorrect)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutoCorrectDoubleSpaceWithPeriod(isCorrect);
	};

	asc_docs_api.prototype.asc_SetFirstLetterAutoCorrectExceptions = function(exceptions, lang)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;

		let exceptionManager = logicDocument.GetAutoCorrectSettings().GetFirstLetterExceptionManager();
		return exceptionManager.SetExceptions(exceptions, lang);
	};
	asc_docs_api.prototype.asc_GetFirstLetterAutoCorrectExceptions = function(lang)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return [];

		let exceptionManager = logicDocument.GetAutoCorrectSettings().GetFirstLetterExceptionManager();
		return exceptionManager.GetExceptions(lang);
	};
	asc_docs_api.prototype.asc_GetAutoCorrectSettings = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return null;

		return logicDocument.GetAutoCorrectSettings();
	};

	asc_docs_api.prototype.asc_RemoveSelection = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return "";

		return oLogicDocument.RemoveSelection();
	};
	asc_docs_api.prototype.asc_AddText = function(sText, oSettings)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return "";

		return oLogicDocument.AddTextWithPr(sText, oSettings);
	};
	asc_docs_api.prototype.asc_GetCurrentWord = function(nDirection)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return "";

		return oLogicDocument.GetCurrentWord(nDirection);
	};
	asc_docs_api.prototype.asc_ReplaceCurrentWord = function(nDirection, sReplace)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return "";

		return oLogicDocument.ReplaceCurrentWord(nDirection, sReplace);
	};
	asc_docs_api.prototype.asc_GetSelectedText = function(bClearText, select_Pr)
	{
		bClearText = (bClearText === true);
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return "";
		
		return logicDocument.GetSelectedText(bClearText, select_Pr);
	};
	asc_docs_api.prototype.asc_AddBlankPage = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return null;

		oLogicDocument.AddBlankPage();
	};
	asc_docs_api.prototype.asc_EditSelectAll = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return null;
		
		if (this.isTargetHandMode() && !logicDocument.IsInFormField())
			this.asc_setViewerTargetType("select");
		
		logicDocument.SelectAll();
	};
	asc_docs_api.prototype.asc_enterText = function(value)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;

		return logicDocument.EnterText(value);
	};
	asc_docs_api.prototype.asc_correctEnterText = function(oldValue, newValue)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;

		return logicDocument.CorrectEnterText(oldValue, newValue);
	};

	// input
	asc_docs_api.prototype.Begin_CompositeInput = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;
		
		return logicDocument.getCompositeInput().begin();
	};
	asc_docs_api.prototype.Add_CompositeText = function(nCharCode)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		return logicDocument.getCompositeInput().add(nCharCode);
	};
	asc_docs_api.prototype.Remove_CompositeText = function(nCount)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		return logicDocument.getCompositeInput().remove(nCount);
	};
	asc_docs_api.prototype.Replace_CompositeText = function(arrCharCodes)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		return logicDocument.getCompositeInput().replace(arrCharCodes);
	};
	asc_docs_api.prototype.Set_CursorPosInCompositeText = function(nPos)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		return logicDocument.getCompositeInput().setPos(nPos);
	};
	asc_docs_api.prototype.Get_CursorPosInCompositeText = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return 0;
		
		return logicDocument.getCompositeInput().getPos();
	};
	asc_docs_api.prototype.End_CompositeInput = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;

		return logicDocument.getCompositeInput().end();
	};
	asc_docs_api.prototype.Get_MaxCursorPosInCompositeText = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return 0;
		
		return logicDocument.getCompositeInput().getMaxPos();
	};
	asc_docs_api.prototype.Input_UpdatePos = function()
	{
		if (this.WordControl.m_oLogicDocument)
			this.WordControl.m_oDrawingDocument.MoveTargetInInputContext();
	};

	asc_docs_api.prototype.onKeyDown = function(e)
	{
		return this.WordControl.onKeyDown(e);
	};
	asc_docs_api.prototype.onKeyPress = function(e)
	{
		return this.WordControl.onKeyPress(e);
	};
	asc_docs_api.prototype.onKeyUp = function(e)
	{
		return this.WordControl.onKeyUp(e);
	};
	asc_docs_api.prototype.getAddedTextOnKeyDown = function(e)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return [];

		return oLogicDocument.GetAddedTextOnKeyDown(e);
	};

	asc_docs_api.prototype.sync_OnConvertEquationToMath = function(oEquation)
	{
		this.sendEvent("asc_onConvertEquationToMath", oEquation);
	};
	asc_docs_api.prototype.asc_ConvertEquationToMath = function(oEquation, isAll)
	{
		var oLogicDocument = this.private_GetLogicDocument();

		// TODO: Вообще здесь нужно запрашивать шрифты, которые использовались в старой формуле,
		//      но пока это только 1 шрифт "Cambria Math".
		var loader   = AscCommon.g_font_loader;
		var fontinfo = g_fontApplication.GetFontInfo("Cambria Math");
		var isasync  = loader.LoadFont(fontinfo, function()
		{
			oLogicDocument.ConvertEquationToMath(oEquation, isAll);
		}, this);

		if (false === isasync)
		{
			oLogicDocument.ConvertEquationToMath(oEquation, isAll);
		}
	};
	asc_docs_api.prototype.asc_ConvertMathDisplayMode = function(isInline)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		oLogicDocument.ConvertMathDisplayMode(isInline);
	};
	asc_docs_api.prototype.asc_IsInlineMath = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return false;

		let oMath = oLogicDocument.GetCurrentMath();
		if (!oMath)
			return false;

		return oMath.IsInlineMode();
	};
	asc_docs_api.prototype.asc_GetAllNumberedParagraphs = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		return oLogicDocument.GetAllNumberedParagraphs();
	};
	asc_docs_api.prototype.asc_AddCrossRefToParagraph = function(oParagraph, nType, bHyperlink, bAboveBelow, sSeparator)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		oLogicDocument.AddRefToParagraph(oParagraph, nType, bHyperlink, bAboveBelow, sSeparator);
	};
	asc_docs_api.prototype.asc_GetAllHeadingParagraphs = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		var arrOutline = [];
		oLogicDocument.GetOutlineParagraphs(arrOutline, {SkipEmptyParagraphs : true, OutlineStart : 1, OutlineEnd : 9});
		var aParagraphs = [];
		for(var nIndex = 0; nIndex < arrOutline.length; ++nIndex)
		{
			aParagraphs.push(arrOutline[nIndex].Paragraph);
		}
		return aParagraphs;
	};
	asc_docs_api.prototype.asc_AddCrossRefToBookmark = function(sName, nType, bHyperlink, bAboveBelow, sSeparator)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		oLogicDocument.AddRefToBookmark(sName, nType, bHyperlink, bAboveBelow, sSeparator);
	};
	asc_docs_api.prototype.asc_GetAllFootNoteParagraphs = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		return oLogicDocument.GetFootNotesFirstParagraphs();
	};
	asc_docs_api.prototype.asc_AddCrossRefToNote = function(oParagraph, nType, bHyperlink, bAboveBelow)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		oLogicDocument.AddNoteRefToParagraph(oParagraph, nType, bHyperlink, bAboveBelow);
	};
	asc_docs_api.prototype.asc_GetAllEndNoteParagraphs = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		return oLogicDocument.GetEndNotesFirstParagraphs();
	};
	asc_docs_api.prototype.asc_GetAllCaptionParagraphs = function(sCaption)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		return oLogicDocument.GetAllCaptionParagraphs(sCaption);
	};
	asc_docs_api.prototype.asc_AddCrossRefToCaption = function(sCaption, oParagraph, nType, bHyperlink, bAboveBelow)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		oLogicDocument.AddRefToCaption(sCaption, oParagraph, nType, bHyperlink, bAboveBelow);
	};
	asc_docs_api.prototype.asc_ChangeTextCase = function(nType)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		oLogicDocument.ChangeTextCase(nType);
	};
	asc_docs_api.prototype.asc_SetSmartParagraphSelection = function(isSmart)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		logicDocument.SetSmartParagraphSelection(isSmart);
	};
	asc_docs_api.prototype.asc_IsSmartParagraphSelection = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;
		
		return logicDocument.IsSmartParagraphSelection();
	};

	//comparison
	asc_docs_api.prototype.asc_CompareDocumentUrl = function (sUrl, oOptions, token) {
		if (this.isLocalMode())
			return this.asc_CompareDocumentUrl_local(sUrl, oOptions, token);
		this._CompareDocument({url: sUrl, format: "docx", token: token}, oOptions);
	};
	asc_docs_api.prototype.asc_CompareDocumentFile = function (oOptions) {
		if (this.isLocalMode())
			return this.asc_CompareDocumentFile_local(oOptions);
		var t = this;
		AscCommon.ShowDocumentFileDialog(function (error, files) {
			if (Asc.c_oAscError.ID.No !== error) {
				t.sendEvent("asc_onError", error, Asc.c_oAscError.Level.NoCritical);
				return;
			}
			var format = AscCommon.GetFileExtension(files[0].name);
			var reader = new FileReader();
			reader.onload = function () {
				t._CompareDocument({data: new Uint8Array(reader.result), format: format}, oOptions);
			};
			reader.onerror = function () {
				t.sendEvent("asc_onError", Asc.c_oAscError.ID.Unknown, Asc.c_oAscError.Level.NoCritical);
			};
			reader.readAsArrayBuffer(files[0]);
		});
	};
	asc_docs_api.prototype.asc_CompareDocumentUrl_local = function (sUrl, oOptions, token) {
		var t = this;
		t.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
		window["AscDesktopEditor"]["DownloadFiles"]([sUrl], [], function(files) {
			t.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
			var _files= [];
			for (var _elem in files) {
				_files.push(files[_elem]);
			}
			if (_files.length == 1)
			{
				if (t.isCloudModeCrypt())
				{
					var fileContent = t.getFileAsFromChanges();
					window["AscDesktopEditor"]["CompareDocumentUrl"](_files[0], oOptions, fileContent.data, fileContent.header);
				}
				else
				{
					window["AscDesktopEditor"]["CompareDocumentUrl"](_files[0], oOptions);
				}
			}
		});
	};

	asc_docs_api.prototype.getDrawingObjects = function () {
/*		const oController = this.getGraphicController();
		if (oController) {
			return oController.drawingObjects;
		}*/
		return null;
	};

	asc_docs_api.prototype.getDrawingDocument = function () {
		const oLogicDocument = this.private_GetLogicDocument();
		return oLogicDocument.GetDrawingDocument();
	}
	asc_docs_api.prototype.getLogicDocument = asc_docs_api.prototype.private_GetLogicDocument;
	asc_docs_api.prototype.asc_createSmartArt = function (nSmartArtType, oPlaceholder)
	{
		const oThis = this;
		AscCommon.g_oBinarySmartArts.checkLoadDrawing().then(function()
		{
			return AscCommon.g_oBinarySmartArts.checkLoadData(nSmartArtType);
		}).then(function()
		{
			if (false === oThis.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
			{
				oThis.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_AddSmartArt);
				oThis.asc_SetSilentMode(true);
				AscFonts.IsCheckSymbols = true;
				const oSmartArt = new AscFormat.SmartArt();
				oSmartArt.fillByPreset(nSmartArtType);
				oSmartArt.fitToPageSize();
				oThis.WordControl.m_oLogicDocument.AddInlineImage(null, null, null, oSmartArt);
				oSmartArt.generateDrawingPart();
				AscFonts.IsCheckSymbols = false;

				AscFonts.FontPickerByCharacter.checkText("", oThis, function() {
					oThis.asc_SetSilentMode(false, true);
					oThis.WordControl.m_oLogicDocument.FinalizeAction();
				}, false, false, false);
			}
		});
	};
	asc_docs_api.prototype.asc_CompareDocumentFile_local = function (oOptions) {
		var t = this;
		window["AscDesktopEditor"]["OpenFilenameDialog"]("word", false, function(_file) {
			var file = _file;
			if (Array.isArray(file))
				file = file[0];

			if (file && "" != file)
			{
				if (t.isCloudModeCrypt())
				{
					var fileContent = t.getFileAsFromChanges();
					window["AscDesktopEditor"]["CompareDocumentFile"](file, oOptions, fileContent.data, fileContent.header);
				}
				else
				{
					window["AscDesktopEditor"]["CompareDocumentFile"](file, oOptions);
				}
			}
		});
	};

	asc_docs_api.prototype.asc_MergeDocumentUrl = function (sUrl, oOptions, token) {
		if (this.isLocalMode())
			return this.asc_MergeDocumentUrl_local(sUrl, oOptions, token);
		
		this._MergeDocument({url: sUrl, format: "docx", token: token}, oOptions);
	};
	asc_docs_api.prototype.asc_MergeDocumentFile = function (oOptions) {
		if (this.isLocalMode())
			return this.asc_MergeDocumentFile_local(oOptions);
		var t = this;
		AscCommon.ShowDocumentFileDialog(function (error, files) {
			if (Asc.c_oAscError.ID.No !== error) {
				t.sendEvent("asc_onError", error, Asc.c_oAscError.Level.NoCritical);
				return;
			}
			var format = AscCommon.GetFileExtension(files[0].name);
			var reader = new FileReader();
			reader.onload = function () {
				t._MergeDocument({data: new Uint8Array(reader.result), format: format}, oOptions);
			};
			reader.onerror = function () {
				t.sendEvent("asc_onError", Asc.c_oAscError.ID.Unknown, Asc.c_oAscError.Level.NoCritical);
			};
			reader.readAsArrayBuffer(files[0]);
		});
	};
	asc_docs_api.prototype.asc_MergeDocumentUrl_local = function (sUrl, oOptions, token) {
		var t = this;
		t.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
		window["AscDesktopEditor"]["DownloadFiles"]([sUrl], [], function(files) {
			t.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
			var _files= [];
			for (var _elem in files) {
				_files.push(files[_elem]);
			}
			if (_files.length == 1)
			{
				if (t.isCloudModeCrypt())
				{
					var fileContent = t.getFileAsFromChanges();
					window["AscDesktopEditor"]["MergeDocumentUrl"](_files[0], oOptions, fileContent.data, fileContent.header);
				}
				else
				{
					window["AscDesktopEditor"]["MergeDocumentUrl"](_files[0], oOptions);
				}
			}
		});
	};
	asc_docs_api.prototype.asc_MergeDocumentFile_local = function (oOptions) {
		var t = this;
		window["AscDesktopEditor"]["OpenFilenameDialog"]("word", false, function(_file) {
			var file = _file;
			if (Array.isArray(file))
				file = file[0];

			if (file && "" != file)
			{
				if (t.isCloudModeCrypt())
				{
					var fileContent = t.getFileAsFromChanges();
					window["AscDesktopEditor"]["MergeDocumentFile"](file, oOptions, fileContent.data, fileContent.header);
				}
				else
				{
					window["AscDesktopEditor"]["MergeDocumentFile"](file, oOptions);
				}
			}
		});
	};

	asc_docs_api.prototype._ConvertDocuments = function (documents, isUseDirectUrlError, fForEachCallback, fEndCallback) {
		fEndCallback = fEndCallback || function (api) {};
		this.insertDocumentUrlsData = {
			imageMap: null, documents: documents, convertCallback: function (_api, url) {
				if (!url['output.bin']) {
					_api.endInsertDocumentUrls();
					_api.sendEvent("asc_onError", Asc.c_oAscError.ID.DirectUrl,
						Asc.c_oAscError.Level.NoCritical);
					return;
				}
				AscCommon.loadFileContent(url['output.bin'], function (httpRequest) {
					const stream = httpRequest && AscCommon.initStreamFromResponse(httpRequest);
					if (stream) {
						fForEachCallback(stream, url);
					} else {
						_api.sendEvent("asc_onError", Asc.c_oAscError.ID.DirectUrl,
							Asc.c_oAscError.Level.NoCritical);
						return;
					}
					if (_api.insertDocumentUrlsData.documents.length > 0) {
						const options = new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.CANVAS_WORD);
						options.isNaturalDownload = true;
						options.isGetTextFromUrl = true;
						if (isUseDirectUrlError) {
							options.errorDirect = Asc.c_oAscError.ID.DirectUrl;
						}
						_api.downloadAs(Asc.c_oAscAsyncAction.DownloadAs, options);
					} else {
						_api.endInsertDocumentUrls();
					}
				}, "arraybuffer");
			}, endCallback: fEndCallback
		};

		const options = new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.CANVAS_WORD);
		options.isNaturalDownload = true;
		options.isGetTextFromUrl = true;
		if (isUseDirectUrlError) {
			options.errorDirect = Asc.c_oAscError.ID.DirectUrl;
		}
		this.downloadAs(Asc.c_oAscAsyncAction.DownloadAs, options);
	};
	
	asc_docs_api.prototype._CompareDocument = function (document, oOptions) {
		const oThis = this;
		this._ConvertDocuments([document], !!document.url, function (stream, imageMap) {
			oThis.insertDocumentUrlsData.imageMap = imageMap;
			AscCommonWord.CompareBinary(oThis, stream, oOptions);
		});
	};

	asc_docs_api.prototype._MergeDocument = function (document, oOptions) {
		const oThis = this;
		this._ConvertDocuments([document], !!document.url, function (stream, imageMap) {
			oThis.insertDocumentUrlsData.imageMap = imageMap;
			AscCommonWord.mergeBinary(oThis, stream, oOptions);
		});
	};
	
	// callback for local mode
	window["onDocumentCompare"] = function(folder, file_content, file_content_len, image_map, options) {
		var api = window.editor;
		if (file_content == "")
		{
			api.sendEvent("asc_onError", Asc.c_oAscError.ID.ConvertationOpenError, Asc.c_oAscError.Level.NoCritical);
			return;
		}

		var file = {};
		file["IsValid"] = function() { return true; };
		file["GetBinary"] = function() { return AscCommon.getBinaryArray(file_content, file_content_len); };
		file["GetImageMap"] = function() { return image_map; };

		AscCommonWord["CompareDocuments"](api, file);
	};

	window["onDocumentMerge"] = function(folder, file_content, file_content_len, image_map, options) {
		var api = window.editor;
		if (file_content == "")
		{
			api.sendEvent("asc_onError", Asc.c_oAscError.ID.ConvertationOpenError, Asc.c_oAscError.Level.NoCritical);
			return;
		}

		var file = {};
		file["IsValid"] = function() { return true; };
		file["GetBinary"] = function() { return AscCommon.getBinaryArray(file_content, file_content_len); };
		file["GetImageMap"] = function() { return image_map; };

		AscCommonWord["mergeDocuments"](api, file);
	};

	window["asc_docs_api"]                                      = asc_docs_api;
	window["asc_docs_api"].prototype["asc_nativeOpenFile"]      = function(base64File, version)
	{
		this.SpellCheckUrl = '';

		this.User = new AscCommon.asc_CUser();
		this.User.setId("TM");
		this.User.setUserName("native");

		this.WordControl.m_bIsRuler = false;
		this.WordControl.Init();

		this.InitEditor();
		this.DocumentType   = 2;
		this.LoadedObjectDS = this.WordControl.m_oLogicDocument.CopyStyle();

		g_oIdCounter.Set_Load(true);

		if (undefined !== version)
			AscCommon.CurFileVersion = version;

		this.isOpenOOXInBrowser = this["asc_isSupportFeature"]("ooxml") && AscCommon.checkOOXMLSignature(base64File);
		if (this.isOpenOOXInBrowser) {
			//slice because array contains garbage after end of function
			this.openOOXInBrowserZip = base64File.slice();
			if (!this.OpenDocumentFromZipNoInit(base64File))
				this.sendEvent("asc_onError", c_oAscError.ID.MobileUnexpectedCharCount, c_oAscError.Level.Critical);
		} else {
			let openParams        = {};
			let oBinaryFileReader = new AscCommonWord.BinaryFileReader(this.WordControl.m_oLogicDocument, openParams);
			if (!oBinaryFileReader.Read(base64File))
				this.sendEvent("asc_onError", c_oAscError.ID.MobileUnexpectedCharCount, c_oAscError.Level.Critical);
		}

		g_oIdCounter.Set_Load(false);
		this.LoadedObject = 1;

		if (window["NATIVE_EDITOR_ENJINE"] === true && undefined != window["native"])
		{
			AscCommon.CDocsCoApi.prototype.askSaveChanges = function(callback)
			{
				callback({"saveLock" : false});
			};
			AscCommon.CDocsCoApi.prototype.saveChanges    = function(arrayChanges, deleteIndex, excelAdditionalInfo)
			{
				if (window["native"]["SaveChanges"])
					window["native"]["SaveChanges"](arrayChanges.join("\",\""), deleteIndex, arrayChanges.length);
			};
		}

		//callback
		this.DocumentOrientation = (null == editor.WordControl.m_oLogicDocument) ? true : !editor.WordControl.m_oLogicDocument.Orientation;
		var sizeMM;
		if (this.DocumentOrientation)
			sizeMM = DocumentPageSize.getSize(AscCommon.Page_Width, AscCommon.Page_Height);
		else
			sizeMM = DocumentPageSize.getSize(AscCommon.Page_Height, AscCommon.Page_Width);
		this.sync_DocSizeCallback(sizeMM.w_mm, sizeMM.h_mm);
		this.sync_PageOrientCallback(editor.get_DocumentOrientation());

		if (this.GenerateNativeStyles !== undefined)
		{
			this.GenerateNativeStyles();

			if (this.WordControl.m_oDrawingDocument.CheckTableStylesOne !== undefined)
				this.WordControl.m_oDrawingDocument.CheckTableStylesOne();
		}
	};
	window["asc_docs_api"].prototype["asc_nativeCalculateFile"] = function(options)
	{
		if (null == this.WordControl.m_oLogicDocument)
			return;

		var Document = this.WordControl.m_oLogicDocument;

		if ((window["NATIVE_EDITOR_ENJINE"] === undefined) && this.isApplyChangesOnOpenEnabled)
		{
			this.isApplyChangesOnOpenEnabled = false;
			if (1 === AscCommon.CollaborativeEditing.m_nUseType)
			{
				this.isApplyChangesOnOpen = true;
				AscCommon.CollaborativeEditing.Apply_Changes();
				AscCommon.CollaborativeEditing.Release_Locks();
				return;
			}
		}

		Document.MoveCursorToStartPos();

		if (window['IS_NATIVE_EDITOR']) {
			Document.RecalculateFromStart();
		} else {
			var pagescount = undefined;
			if (options && options["printOptions"] && options["printOptions"]["onlyFirstPage"])
				pagescount = 1;
			Document.RecalculateAllAtOnce(true, pagescount);
		}

		Document.Document_UpdateInterfaceState();
		//Document.Document_UpdateRulersState();
		Document.Document_UpdateSelectionState();

		this.ShowParaMarks = false;
	};

	window["asc_docs_api"].prototype["asc_nativeApplyChanges"] = function(changes)
	{
		this._coAuthoringSetChanges(changes, new AscCommonWord.CDocumentColor(191, 255, 199));
		AscCommon.CollaborativeEditing.Apply_OtherChanges();
	};

	window["asc_docs_api"].prototype.asc_SetSilentMode = function(bEnabled, bAddition)
	{
		if (!this.WordControl.m_oLogicDocument)
			return;
		if (bEnabled)
			this.WordControl.m_oLogicDocument.Start_SilentMode();
		else
			this.WordControl.m_oLogicDocument.End_SilentMode(bAddition);
	};

	window["asc_docs_api"].prototype["asc_nativeApplyChanges2"] = function(data, isFull)
	{
		// Чтобы заново созданные параграфы не отображались залоченными
		g_oIdCounter.Set_Load(true);

		var stream = new AscCommon.FT_Stream2(data, data.length);
		stream.obj = null;
		var _color = new AscCommonWord.CDocumentColor(191, 255, 199);

		// Применяем изменения, пока они есть
		var _count = stream.GetLong();

		var _pos = 4;
		for (var i = 0; i < _count; i++)
		{
			if (window["NATIVE_EDITOR_ENJINE"] === true && window["native"]["CheckNextChange"])
			{
				if (!window["native"]["CheckNextChange"]())
					break;
			}

			var nChangeLen = stream.GetLong();
			_pos += 4;
			stream.size = _pos + nChangeLen;

			var ClassId = stream.GetString2();
			var Class   = AscCommon.g_oTableId.Get_ById(ClassId);

			var nReaderPos  = stream.GetCurPos();
			var nChangeType = stream.GetLong();

			if (Class)
			{
				var fChangesClass = AscDFH.changesFactory[nChangeType];
				if (fChangesClass)
				{
					var oChange = new fChangesClass(Class);
					oChange.ReadFromBinary(stream);

					if (true === AscCommon.CollaborativeEditing.private_AddOverallChange(oChange, false))
						oChange.Load(_color);
				}
				else
				{
					AscCommon.CollaborativeEditing.private_AddOverallChange(data, false);

					stream.Seek(nReaderPos);
					stream.Seek2(nReaderPos);

					Class.Load_Changes(stream, null, _color);
				}
			}

			_pos += nChangeLen;
			stream.Seek2(_pos);
			stream.size = data.length;
		}

		if (isFull)
		{
			AscCommon.CollaborativeEditing.m_aChanges = [];

			// Делаем проверки корректности новых изменений
			AscCommon.CollaborativeEditing.Check_MergeData();

			AscCommon.CollaborativeEditing.OnEnd_ReadForeignChanges();

			if (window["NATIVE_EDITOR_ENJINE"] === true && window["native"]["AddImageInChanges"])
			{
				var _new_images     = AscCommon.CollaborativeEditing.m_aNewImages;
				var _new_images_len = _new_images.length;

				for (var nImage = 0; nImage < _new_images_len; nImage++)
					window["native"]["AddImageInChanges"](_new_images[nImage]);
			}
		}

		g_oIdCounter.Set_Load(false);
	};

	window["asc_docs_api"].prototype["asc_nativeGetFile"] = function()
	{
		var oBinaryFileWriter = new AscCommonWord.BinaryFileWriter(this.WordControl.m_oLogicDocument);
		return oBinaryFileWriter.Write();
	};
	window["asc_docs_api"].prototype["asc_nativeGetFile2"] = function()
	{
		var oBinaryFileWriter = new AscCommonWord.BinaryFileWriter(this.WordControl.m_oLogicDocument);
		return oBinaryFileWriter.Write(true, true);

	};
    window["asc_docs_api"].prototype.asc_nativeGetFile3 = function()
    {
        var oBinaryFileWriter = new AscCommonWord.BinaryFileWriter(this.WordControl.m_oLogicDocument);
        return { data: oBinaryFileWriter.Write(true, true), header: (AscCommon.c_oSerFormat.Signature + ";v" + Asc.c_nVersionNoBase64 + ";" + oBinaryFileWriter.memory.GetCurPosition() + ";") };
    };

	window["asc_docs_api"].prototype["asc_nativeGetFileData"] = function()
	{
		if (this.isOpenOOXInBrowser && this["asc_isSupportFeature"]("ooxml")) {
			let res;
			this.saveLogicDocumentToZip(undefined, undefined, function(data) {
				res = data;
			});
			if (res) {
				window["native"]["Save_End"](";v10;", res.length);
				return res;
			}
			return new Uint8Array(0);
		} else {
			var oBinaryFileWriter = new AscCommonWord.BinaryFileWriter(this.WordControl.m_oLogicDocument, undefined, undefined, undefined, undefined, undefined);
			var _memory           = oBinaryFileWriter.memory;

			oBinaryFileWriter.Write(true);

			var _header = AscCommon.c_oSerFormat.Signature + ";v" + Asc.c_nVersionNoBase64 + ";" + _memory.GetCurPosition() + ";";
			window["native"]["Save_End"](_header, _memory.GetCurPosition());

			return _memory.data;
		}
	};
	
	window["asc_docs_api"].prototype["asc_nativeGetHtml"] = function()
	{
		var _old                           = PasteElementsId.copyPasteUseBinary;
		PasteElementsId.copyPasteUseBinary = false;
		var oCore = this.asc_getCoreProps();
		this.WordControl.m_oLogicDocument.SelectAll();
		var oCopyProcessor = new AscCommon.CopyProcessor(this);
		oCopyProcessor.Start();
		var _ret = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />";
		if (oCore)
		{
			if (oCore.asc_getTitle())
			{
				var title = AscCommon.escapeHtmlCharacters(oCore.asc_getTitle());
				_ret += ("<title>" + title + "</title>");
			}
			if (oCore.asc_getCreator())
			{
				var creator = AscCommon.escapeHtmlCharacters(oCore.asc_getCreator());
				_ret += ("<meta name=\"creator\" content=\"" + creator + "\" />");
			}
			if (oCore.asc_getDescription())
			{
				var description = AscCommon.escapeHtmlCharacters(oCore.asc_getDescription());
				_ret += ("<meta name=\"description\" content=\"" + description + "\" />");
			}
			if (oCore.asc_getKeywords())
			{
				var keywords = AscCommon.escapeHtmlCharacters(oCore.asc_getKeywords());
				_ret += ("<meta name=\"keywords\" content=\"" + keywords + "\" />");
			}
			if (oCore.asc_getSubject())
			{
				var subject = AscCommon.escapeHtmlCharacters(oCore.asc_getSubject());
				_ret += ("<meta name=\"subject\" content=\"" + subject + "\" />");
			}
			if (oCore.asc_getIdentifier())
			{
				var identifier = AscCommon.escapeHtmlCharacters(oCore.asc_getIdentifier());
				_ret += ("<meta name=\"identifier\" content=\"" + identifier + "\" />");
			}
			if (oCore.asc_getLanguage())
			{
				var language = AscCommon.escapeHtmlCharacters(oCore.asc_getLanguage());
				_ret += ("<meta name=\"language\" content=\"" + language + "\" />");
			}
		}
		_ret += ("</head><body>" + oCopyProcessor.getInnerHtml() + "</body></html>");

		this.WordControl.m_oLogicDocument.RemoveSelection();
		PasteElementsId.copyPasteUseBinary = _old;
		return _ret;
	};

	window["asc_docs_api"].prototype["asc_AddHtml"] = function(_iframeId)
	{
		var ifr = document.getElementById(_iframeId);

		var frameWindow = window.frames[_iframeId];
		if (frameWindow)
		{
			if (null != frameWindow.document && null != frameWindow.document.body && this.WordControl.m_oLogicDocument.CanPerformAction())
			{
				ifr.style.display = "block";
				this.WordControl.m_oLogicDocument.StartAction();
				this.asc_SetSilentMode(true);
				AscCommon.Editor_Paste_Exec(this, AscCommon.c_oAscClipboardDataFormat.HtmlElement, frameWindow.document.body, ifr);
				this.WordControl.m_oLogicDocument.FinalizeAction();
				this.asc_SetSilentMode(false);
			}
		}

		if (ifr)
			document.body.removeChild(ifr);
	};

	window["asc_docs_api"].prototype["asc_nativeCalculate"] = function()
	{
	};

	window["asc_docs_api"].prototype["asc_nativePrint"] = function(_printer, _page, _options)
	{
		let layoutOptions = _options ? _options["documentLayout"] : null;
		if (undefined === _printer && _page === undefined)
		{
			if (undefined !== window["AscDesktopEditor"])
			{
                var isSelection = (_options && _options["printOptions"] && _options["printOptions"]["selection"]) ? true : false;
				var _drawing_document = this.WordControl.m_oDrawingDocument;
                if (isSelection)
                	_drawing_document.GenerateSelectionPrint();

                var _drawing_document_print = _drawing_document.printedDocument ? _drawing_document.printedDocument.DrawingDocument : _drawing_document;
				var pagescount        = Math.min(_drawing_document_print.m_lPagesCount, _drawing_document_print.m_lCountCalculatePages);

				window["AscDesktopEditor"]["Print_Start"](this.DocumentUrl, pagescount, "", _drawing_document.printedDocument ? 0 : this.getCurrentPage());

				var oDocRenderer                  = new AscCommon.CDocumentRenderer();
				oDocRenderer.isPrintMode          = true;
                oDocRenderer.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);
				oDocRenderer.VectorMemoryForPrint = new AscCommon.CMemory();
				var bOldShowMarks                 = this.ShowParaMarks;
				this.ShowParaMarks                = false;

				let oLogicDocument = _drawing_document_print.m_oLogicDocument;
				oLogicDocument.SetupBeforeNativePrint(layoutOptions, oDocRenderer);

				for (var i = 0; i < pagescount; i++)
				{
					oDocRenderer.Memory.Seek(0);
					oDocRenderer.VectorMemoryForPrint.ClearNoAttack();

					var page = _drawing_document_print.m_arrPages[i];
					oDocRenderer.BeginPage(page.width_mm, page.height_mm);
                    _drawing_document_print.m_oLogicDocument.DrawPage(i, oDocRenderer);
					oDocRenderer.EndPage();

					window["AscDesktopEditor"]["Print_Page"](oDocRenderer.Memory.GetBase64Memory(), page.width_mm, page.height_mm);
				}

				this.ShowParaMarks = bOldShowMarks;
				oLogicDocument.RestoreAfterNativePrint();

                _drawing_document.printedDocument = null;
				window["AscDesktopEditor"]["Print_End"]();
			}
		}
		else
		{
			let oLogicDocument = this.WordControl.m_oLogicDocument;
			oLogicDocument.SetupBeforeNativePrint(layoutOptions, _printer);
			var page = this.WordControl.m_oDrawingDocument.m_arrPages[_page];
			_printer.BeginPage(page.width_mm, page.height_mm);
			oLogicDocument.DrawPage(_page, _printer);

			if (this.watermarkDraw)
				this.watermarkDraw.DrawOnRenderer(_printer, page.width_mm, page.height_mm);

			_printer.EndPage();
			oLogicDocument.RestoreAfterNativePrint();
		}
	};

	window["asc_docs_api"].prototype["asc_nativePrintPagesCount"] = function()
	{
		return this.WordControl.m_oDrawingDocument.m_lPagesCount;
	};

	window["asc_docs_api"].prototype["asc_nativeGetPDF"] = function(options)
	{
		if (options && options["watermark"])
		{
			this.watermarkDraw = new AscCommon.CWatermarkOnDraw(options["watermark"], this);
			this.watermarkDraw.generateNative();
		}
		else
		{
			this.watermarkDraw = null;
		}

		var pagescount = this["asc_nativePrintPagesCount"]();
		if (options && options["printOptions"] && options["printOptions"]["onlyFirstPage"])
            pagescount = 1;

		var _renderer                  = new AscCommon.CDocumentRenderer();
		if (options && ((options["saveFormat"] === "image") || (options["isPrint"] === true)))
			_renderer.isPrintMode = true;

        _renderer.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);
		_renderer.VectorMemoryForPrint = new AscCommon.CMemory();
		_renderer.DocInfo(this.asc_getCoreProps());
		var _bOldShowMarks             = this.ShowParaMarks;
		this.ShowParaMarks             = false;

		let nativeOptions = options ? options["nativeOptions"] : undefined;
		let pages = nativeOptions ? AscCommon.getNativePrintRanges(nativeOptions["pages"], nativeOptions["currentPage"], pagescount) : undefined;

		for (var i = 0; i < pagescount; i++)
		{
			if (pages !== undefined && !pages[i])
				continue;
			this["asc_nativePrint"](_renderer, i, options);
		}

		this.ShowParaMarks = _bOldShowMarks;

		window["native"]["Save_End"]("", _renderer.Memory.GetCurPosition());

		return _renderer.Memory.data;
	};

	window["asc_docs_api"].prototype["asc_nativeGetWatermark"] = function(options)
	{
		if (options && options["watermark"])
		{
			if (!this.User)
			{
				this.User = new AscCommon.asc_CUser();
				this.User.setId("TM");
				this.User.setUserName("native");
			}

			this.watermarkDraw = new AscCommon.CWatermarkOnDraw(options["watermark"], this);
			this.watermarkDraw.isNativeGlobalAlpha = true;
			this.watermarkDraw.generateNative();
			return this.watermarkDraw.imageBase64;
		}
		return "";
	};

	// cool api (autotests)
	window["asc_docs_api"].prototype["Add_Text"]                     = function(_text)
	{
		this.WordControl.m_oLogicDocument.TextBox_Put(_text);
	};
	window["asc_docs_api"].prototype["Add_NewParagraph"]             = function()
	{
		var LogicDocument = this.WordControl.m_oLogicDocument;
		if (false === LogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Document_Content_Add))
		{
			LogicDocument.StartAction(AscDFH.historydescription_Document_EnterButton);
			LogicDocument.AddNewParagraph(true);
			LogicDocument.FinalizeAction();
		}
	};
	window["asc_docs_api"].prototype["Cursor_MoveLeft"]              = function()
	{
		this.WordControl.m_oLogicDocument.MoveCursorLeft();
	};
	window["asc_docs_api"].prototype["Cursor_MoveRight"]             = function()
	{
		this.WordControl.m_oLogicDocument.MoveCursorRight();
	};
	window["asc_docs_api"].prototype["Cursor_MoveUp"]                = function()
	{
		this.WordControl.m_oLogicDocument.MoveCursorUp();
	};
	window["asc_docs_api"].prototype["Cursor_MoveDown"]              = function()
	{
		this.WordControl.m_oLogicDocument.MoveCursorDown();
	};
	window["asc_docs_api"].prototype["Get_DocumentRecalcId"]         = function()
	{
		return this.WordControl.m_oLogicDocument.RecalcId;
	};
	window["asc_docs_api"].prototype["asc_IsSpellCheckCurrentWord"]  = function()
	{
		return this.IsSpellCheckCurrentWord;
	};
	window["asc_docs_api"].prototype["asc_putSpellCheckCurrentWord"] = function(value)
	{
		this.IsSpellCheckCurrentWord = value;
	};
	window["asc_docs_api"].prototype["asc_setParagraphStylesSizes"] = function(width, height)
	{
		if (window["AscCommon"] && window["AscCommon"].GlobalSkin)
		{
			AscCommon.GlobalSkin.STYLE_THUMBNAIL_WIDTH = width;
			AscCommon.GlobalSkin.STYLE_THUMBNAIL_HEIGHT = height;
		}
		else
		{
			AscCommon.TEMP_STYLE_THUMBNAIL_WIDTH = width;
			AscCommon.TEMP_STYLE_THUMBNAIL_HEIGHT = height;
		}
	};

	asc_docs_api.prototype.SetDrawImagePlaceContents = function(id, props)
	{
		if (this.WordControl.m_oDrawingDocument)
			this.WordControl.m_oDrawingDocument.SetDrawImagePlaceContents(id, props);
	};
	asc_docs_api.prototype.SetDrawImagePlaceTableOfFigures = function(id, props)
	{
		if (this.WordControl.m_oDrawingDocument)
			this.WordControl.m_oDrawingDocument.SetDrawImagePlaceTableOfFigures(id, props);
	};

    asc_docs_api.prototype.SetDrawImagePreviewMargins = function(id, props)
    {
        if (this.WordControl.m_oDrawingDocument)
            this.WordControl.m_oDrawingDocument.SetDrawImagePreviewMargins(id, props);
    };

    asc_docs_api.prototype.SetDrawImagePreviewBullet = function(sDivId, oProps, nCurrentLvl, bIsMultiLvl)
    {
		const oDrawer = new AscCommon.CBulletPreviewDrawerAdvancedOptions(sDivId, oProps, nCurrentLvl, bIsMultiLvl);
        oDrawer.checkFontsAndDraw();
	};
	
	asc_docs_api.prototype.SetDrawImagePreviewBulletChangeListLevel = function(arrId, oProps)
    {
		const oDrawer = new AscCommon.CBulletPreviewDrawerChangeList(arrId, oProps);
		oDrawer.checkFontsAndDraw();
    };

	asc_docs_api.prototype.asc_PreConvertTextToTable = function(oProps)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		return oLogicDocument.GetConvertTextToTableProps(oProps);
	};

	asc_docs_api.prototype.asc_ConvertTextToTable = function(oProps)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		return oLogicDocument.ConvertTextToTable(oProps);
	};

	asc_docs_api.prototype.asc_ConvertTableToText = function(sep_type, sep, nested)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		return oLogicDocument.ConvertTableToText({type: sep_type, separator: sep, nested: nested});
	};
	asc_docs_api.prototype.asc_GetPossibleNumberingLanguage = function ()
	{
		const oLogicDocument = this.private_GetLogicDocument();
		if (oLogicDocument)
		{
			const oFirstParagraph = oLogicDocument.GetCurrentParagraph(false);
			if (oFirstParagraph)
			{
				const oTextPr = oFirstParagraph.Get_CompiledPr2(false).TextPr.Copy();
				const oTempTextPr = oFirstParagraph.TextPr && oFirstParagraph.TextPr.Value.Copy();
				if (oTempTextPr)
				{
					oTextPr.Merge(oTempTextPr);
				}
				return oTextPr.Lang;
			}

		}
		return null;
	};
	asc_docs_api.prototype.CheckDeprecatedBulletPreviewInfo = function (arrDrawingInfo, nTypeOfPreview)
	{
		return arrDrawingInfo;
	};
	asc_docs_api.prototype.ParseBulletPreviewInformation = function (arrDrawingInfo)
	{
		const arrNumberingLvls = [];
		AscFormat.ExecuteNoHistory(function ()
		{
			for (let i = 0; i < arrDrawingInfo.length; i += 1)
			{
				const oDrawInfo = arrDrawingInfo[i];
				const numInfo = AscWord.CNumInfo.Parse(oDrawInfo["numberingInfo"]);
				
				if (!numInfo)
					continue;
				
				const sDivId = oDrawInfo["divId"];
				if (numInfo.IsRemove())
				{
					const oLvl = new AscCommonWord.CNumberingLvl();
					const oTextPr = oLvl.GetTextPr();
					const sNoneText = AscCommon.translateManager.getValue("None");
					const sFontName = "Arial";
					oLvl.FillLvlText([sNoneText]);
					oTextPr.RFonts.SetAll(sFontName);
					arrNumberingLvls.push({divId: sDivId, arrLvls:[oLvl], isRemoving: true});
				}
				else
				{
					const arrNumberingLvl = [];
					const arrLvls = numInfo.Lvl;
					for (let j = 0; j < arrLvls.length; j += 1)
					{
						const oLvl = AscCommonWord.CNumberingLvl.FromJson(arrLvls[j]);
						if (oLvl.TextPr.RFonts.IsEmpty())
						{
							oLvl.TextPr.RFonts.SetAll("Arial");
						}
						if (numInfo.IsHeadings())
						{
							const oLogicDocument = this.private_GetLogicDocument();
							const oStyles = oLogicDocument.GetStyles();
							if (oStyles)
							{
								const nStyleId = oStyles.GetDefaultHeading(j);
								oLvl.SetPStyle(nStyleId);
							}
						}
						arrNumberingLvl.push(oLvl);
					}
					arrNumberingLvls.push({divId: sDivId, arrLvls: arrNumberingLvl});
				}
			}
		}, this);
		return arrNumberingLvls;
	};

	asc_docs_api.prototype.asc_OnHideContextMenu = function()
	{
		if (this.WordControl.MobileTouchManager)
		{
			this.WordControl.checkBodyOffset();
			this.WordControl.MobileTouchManager.showKeyboard();
		}
	};
	asc_docs_api.prototype.asc_OnShowContextMenu = function()
	{
		if (this.WordControl.MobileTouchManager)
		{
			this.WordControl.checkBodyOffset();
		}
	};

	asc_docs_api.prototype["asc_getButtonsTOC"] = function(id1, id2, styleWidth)
	{
		if (this.WordControl && this.WordControl.m_oDrawingDocument)
		{
			this.WordControl.m_oDrawingDocument.GetTOC_Buttons(id1, id2, styleWidth);
		}
	};

	asc_docs_api.prototype.getDefaultFontFamily = function () {
		//TODO переделать и отдавать дефолтовый шрифт
		var defaultFont = "Arial";
		return defaultFont;
	};

	asc_docs_api.prototype.getDefaultFontSize = function () {
		//TODO переделать и отдавать дефолтовый шрифт
		var defaultSize = 11;
		return defaultSize;
	};

	asc_docs_api.prototype.asc_getPdfProps = function()
	{
		return null;
	};

	asc_docs_api.prototype.asc_getAppProps = function()
	{
		return this.WordControl && this.WordControl.m_oLogicDocument && this.WordControl.m_oLogicDocument.App || null;
	};

	asc_docs_api.prototype.getInternalCoreProps = function()
	{
		return this.WordControl && this.WordControl.m_oLogicDocument && this.WordControl.m_oLogicDocument.Core;
	};

	asc_docs_api.prototype.asc_setCoreProps = function(oProps)
	{
		var oCore = this.getInternalCoreProps();
		if(!oCore)
		{
			return;
		}
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if(false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_CorePr, null))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_SetCoreproperties);
			oCore.setProps(oProps);
			this.UpdateInterfaceState();
			oLogicDocument.FinalizeAction(true);
		}
	};
	asc_docs_api.prototype.asc_isCompatibilityMode = function()
	{
		if (null !== this.WordControl.m_oLogicDocument)
		{
			return AscCommon.document_compatibility_mode_Word15 !== this.WordControl.m_oLogicDocument.GetCompatibilityMode();
		}
		return false;
	};
	asc_docs_api.prototype.asc_Remove = function()
	{
		var oDocument = this.private_GetLogicDocument();
		if(oDocument)
		{
			var oController = oDocument.DrawingObjects;
			if(oController)
			{
				oController.resetChartElementsSelection();
				AscCommon.baseEditorsApi.prototype.asc_Remove.call(this);
			}
		}
	};

	asc_docs_api.prototype.turnOffSpecialModes = function()
	{
		const oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
			return;
		let bReset = false;
		if(this.isDrawTablePen)
		{
			this.sync_TableDrawModeCallback(false);
			bReset = true;
		}
		if (this.isDrawTableErase)
		{
			this.sync_TableEraseModeCallback(false);
			bReset = true;
		}
		if (this.isMarkerFormat)
		{
			this.sync_MarkerFormatCallback(false);
			bReset = true;
		}
		if (this.isFormatPainterOn())
		{
			this.sync_PaintFormatCallback(c_oAscFormatPainterState.kOff);
			bReset = true;
		}
		if (this.isStartAddShape)
		{
			this.sync_StartAddShapeCallback(false);
			this.sync_EndAddShape();
			oLogicDocument.DrawingObjects.endTrackNewShape();
			bReset = true;
		}
		if(this.isInkDrawerOn())
		{
			this.stopInkDrawer();
			bReset = true;
		}
		if(this.isEyedropperStarted())
		{
			this.cancelEyedropper();
			bReset = true;
		}
		return bReset;
	};

	asc_docs_api.prototype.removeAllInks = function()
	{
		let oLogicDocument = this.getLogicDocument();
		if(!oLogicDocument) return;
		return oLogicDocument.RemoveAllInks();
	};


	asc_docs_api.prototype.haveInks = function() {
		let oLogicDocument = this.getLogicDocument();
		if(!oLogicDocument) return;
		return oLogicDocument.HaveInks();
	};

	asc_docs_api.prototype.onUpdateRestrictions = function(restrictionSettings)
	{
		if (this.WordControl)
			this.WordControl.checkMouseHandMode();

		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		if (this.WordControl && this.WordControl.m_oDrawingDocument)
		{
			var contentControls = this.WordControl.m_oDrawingDocument.contentControls;

			if (contentControls && contentControls.ContentControlObjects.length !== 0)
			{
				contentControls.clearAttack();
				oLogicDocument.UpdateSelection();
			}
		}

		const bTurnOffSpecialModes = this.turnOffSpecialModes();
		if(bTurnOffSpecialModes)
		{
			oLogicDocument.UpdateCursorType(oLogicDocument.CurPos.RealX, oLogicDocument.CurPos.RealY, oLogicDocument.CurPage, new AscCommon.CMouseEventHandler());
		}
		
		if (!oLogicDocument.IsFillingFormMode() && oLogicDocument.FocusCC && oLogicDocument.FocusCC.IsForm())
			oLogicDocument.FocusCC.SelectContentControl();

		oLogicDocument.private_CheckCursorPosInFillingFormMode();
		oLogicDocument.UpdateSelection();
		oLogicDocument.UpdateInterface();
		
		let oform = oLogicDocument.GetOFormDocument();
		if (oform)
		{
			let roleName = restrictionSettings ? restrictionSettings.GetOFormRole() : null;
			
			if (oLogicDocument.IsFillingFormMode() && restrictionSettings && restrictionSettings.IsOFormNoRole())
				oform.setCurrentNoRole();
			else if (oLogicDocument.IsFillingFormMode() && roleName)
				oform.setCurrentRole(roleName);
			else
				oform.clearCurrentRole();
			
			// Need to redraw when we change role
			let drawingDocument = oLogicDocument.GetDrawingDocument();
			if (drawingDocument)
			{
				drawingDocument.ClearCachePages();
				drawingDocument.FirePaint();
			}
		}
	};
	asc_docs_api.prototype.isShowShapeAdjustments = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (oLogicDocument)
			return oLogicDocument.IsShowShapeAdjustments();

		return true;
	};
	asc_docs_api.prototype.isShowTableAdjustments = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (oLogicDocument)
			return oLogicDocument.IsShowTableAdjustments();

		return true;
	};
	asc_docs_api.prototype.isShowEquationTrack = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (oLogicDocument)
			return oLogicDocument.IsShowEquationTrack();

		return true;
	};
	asc_docs_api.prototype.canEnterText = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		return logicDocument ? logicDocument.canEnterText() : false;
	};
	asc_docs_api.prototype.updateDarkMode = function()
	{
		if (!this.WordControl || !this.WordControl.m_oDrawingDocument)
			return;
		
		this.WordControl.m_oDrawingDocument.ClearCachePages();
		this.WordControl.m_oDrawingDocument.FirePaint();
		this.WordControl.m_oDrawingDocument.UpdateTargetNoAttack();
	};
    asc_docs_api.prototype.getPageBackgroundColor = function()
    {
    	// TODO: get color from theme
		if (this.isDarkMode)
			return { R: 0x3A, G : 0x3A, B : 0x3A };
		return { R: 0xFF, G : 0xFF, B : 0xFF };
    };
	asc_docs_api.prototype.getPageStrokeColor = function()
	{
		// TODO: get color from theme
		if (this.isDarkMode)
			return "#616161";
		return GlobalSkin.PageOutline;
	};

	asc_docs_api.prototype.asc_GetSelectionBounds = function()
	{
		var oLogicDocument   = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return [[0,0], [0,0], [0,0], [0,0]];

		var oDrawingDocument = oLogicDocument.GetDrawingDocument();
		if (!oDrawingDocument)
			return [[0,0], [0,0], [0,0], [0,0]];

		var oSelectBounds = oLogicDocument.GetSelectionBounds();
		if (!oSelectBounds || !oSelectBounds.Start || !oSelectBounds.End)
			return [[0,0], [0,0], [0,0], [0,0]];

		var x0 = oSelectBounds.Start.X;
		var y0 = oSelectBounds.Start.Y;
		var x1 = oSelectBounds.Start.X;
		var y1 = oSelectBounds.Start.Y + oSelectBounds.Start.H;
		var x2 = oSelectBounds.End.X + oSelectBounds.End.W;
		var y2 = oSelectBounds.End.Y
		var x3 = oSelectBounds.End.X + oSelectBounds.End.W;
		var y3 = oSelectBounds.End.Y + oSelectBounds.End.H;

		var oCurrentParagraph = oLogicDocument.GetCurrentParagraph();

		var oMatrix = oCurrentParagraph ? oCurrentParagraph.Get_ParentTextTransform() : null;
		if (oMatrix && !AscCommon.global_MatrixTransformer.IsIdentity(oMatrix))
		{
			x0 = oMatrix.TransformPointX(oSelectBounds.Start.X, oSelectBounds.Start.Y);
			y0 = oMatrix.TransformPointY(oSelectBounds.Start.X, oSelectBounds.Start.Y);
			x1 = oMatrix.TransformPointX(oSelectBounds.Start.X, oSelectBounds.Start.Y + oSelectBounds.Start.H);
			y1 = oMatrix.TransformPointY(oSelectBounds.Start.X, oSelectBounds.Start.Y + oSelectBounds.Start.H);
			x2 = oMatrix.TransformPointX(oSelectBounds.End.X + oSelectBounds.End.W, oSelectBounds.End.Y);
			y2 = oMatrix.TransformPointY(oSelectBounds.End.X + oSelectBounds.End.W, oSelectBounds.End.Y);
			x3 = oMatrix.TransformPointX(oSelectBounds.End.X + oSelectBounds.End.W, oSelectBounds.End.Y + oSelectBounds.End.H);
			y3 = oMatrix.TransformPointY(oSelectBounds.End.X + oSelectBounds.End.W, oSelectBounds.End.Y + oSelectBounds.End.H);
		}

		var oPos0 = oDrawingDocument.ConvertCoordsToCursorWR(x0, y0, oSelectBounds.Start.Page, undefined, false);
		var oPos1 = oDrawingDocument.ConvertCoordsToCursorWR(x1, y1, oSelectBounds.Start.Page, undefined, false);
		var oPos2 = oDrawingDocument.ConvertCoordsToCursorWR(x2, y2, oSelectBounds.End.Page, undefined, false);
		var oPos3 = oDrawingDocument.ConvertCoordsToCursorWR(x3, y3, oSelectBounds.End.Page, undefined, false);

		return [[oPos0.X, oPos0.Y], [oPos1.X, oPos1.Y], [oPos2.X, oPos2.Y], [oPos3.X, oPos3.Y]];
	};

	asc_docs_api.prototype.asc_canEditGeometry = function ()
	{
		var document = this.private_GetLogicDocument();
		if (document)
		{
			return document.DrawingObjects.canEditGeometry();
		}
		return false;
	};

	asc_docs_api.prototype.asc_editPointsGeometry = function()
	{
		var drawingObjects = this.private_GetLogicDocument().DrawingObjects;
		return drawingObjects.startEditGeometry();
	};

	asc_docs_api.prototype.getImageDataFromSelection = function ()
	{
		let oDocument = this.private_GetLogicDocument();
		if(!oDocument)
		{
			return null;
		}
		return oDocument.GetImageDataFromSelection();
	};

	asc_docs_api.prototype.putImageToSelection = function (sImageSrc, nWidth, nHeight, replaceMode)
	{
		let oDocument = this.private_GetLogicDocument();
		if(!oDocument)
		{
			return;
		}
		return oDocument.PutImageToSelection(sImageSrc, nWidth, nHeight, replaceMode);
	};
	asc_docs_api.prototype.getPluginContextMenuInfo = function(e)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return new AscCommon.CPluginCtxMenuInfo();
		
		let menuInfo;
		if (!logicDocument.IsSelectionUse())
		{
			menuInfo = new AscCommon.CPluginCtxMenuInfo(Asc.c_oPluginContextMenuTypes.Target);
		}
		else
		{
			let drawingObjects = logicDocument.getDrawingObjects();
			if (drawingObjects.selectedObjects && drawingObjects.selectedObjects.length > 0)
				menuInfo = drawingObjects.getPluginSelectionInfo();
			else
				menuInfo = new AscCommon.CPluginCtxMenuInfo(Asc.c_oPluginContextMenuTypes.Selection)
		}
		
		if (logicDocument.isHeaderEditing())
			menuInfo.setHdrFtr(true);
		else if (logicDocument.isFooterEditing())
			menuInfo.setHdrFtr(false);
		
		if (e && e instanceof CContextMenuData && undefined !== e.is_Header())
			menuInfo.setHdrFtrArea(e.is_Header());
		
		return menuInfo;
	};
	
	asc_docs_api.prototype.asc_getDocumentProtection = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument || !logicDocument.IsDocumentEditor())
			return null;
		
		let docProtection = logicDocument.GetDocumentSettings().DocumentProtection;
		if (!docProtection)
			return null;
		
		return docProtection.getAscDocProtect();
	};
	asc_docs_api.prototype.asc_setDocumentProtection = function (props) {
		//props -> CDocProtect

		// Проверка глобального лока
		if (AscCommon.CollaborativeEditing.Get_GlobalLock()) {
			return;
		}

		let oDocument = this.private_GetLogicDocument();
		if (!oDocument) {
			return;
		}

		let curDocProtection = oDocument.Settings && oDocument.Settings.DocumentProtection;
		if (curDocProtection) {
			//пытаемся выставить такие же настройки
			let curIsProtect = curDocProtection.edit != null && curDocProtection.edit !== Asc.c_oAscEDocProtect.None && curDocProtection.enforcement !== false;
			let isPropsProtect = props.edit != null && props.edit !== Asc.c_oAscEDocProtect.None && props.enforcement !== false;
			if (curIsProtect === isPropsProtect) {
				return;
			}
		}

		let t = this;

		let calculatedHashValue;
		let callback = function (res) {
			t.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction);

			if (res) {
				if(false === oDocument.Document_Is_SelectionLocked(AscCommon.changestype_DocumentProtection, null, null, true))
				{
					oDocument.StartAction(AscDFH.historydescription_Document_DocumentProtection);

					props.saltValue = salt;
					props.spinCount = spinCount;
					props.cryptAlgorithmSid = alg;
					props.hashValue = calculatedHashValue;
					props.enforcement = props.edit !== Asc.c_oAscEDocProtect.None;
					props.cryptProviderType = cryptProviderType;

					oDocument.SetProtection(props);
					oDocument.UpdateInterface();
					oDocument.FinalizeAction();
					t.asc_OnProtectionUpdate();
				}
			} else {
				t.sendEvent("asc_onError", c_oAscError.ID.PasswordIsNotCorrect, c_oAscError.Level.NoCritical);
			}
		};

		let prepareHash = function (_val) {
			//todo check end of base64
			//"L6VzBw==d==  d" and "L6VzBw==" in ms equal
			if (_val && _val.length) {
				return _val.replace(/\s/g, "");
			}
			return _val;
		};

		let password = props.temporaryPassword;
		props.temporaryPassword = null;
		let documentProtection = oDocument.Settings.DocumentProtection;
		let salt, alg, spinCount;
		let cryptProviderType = AscCommonWord.ECryptAlgType.TypeAny;
		if (password !== "" && password != null) {
			if (documentProtection) {
				salt = prepareHash(documentProtection.saltValue);
				spinCount =  documentProtection.spinCount;
				alg = documentProtection.cryptAlgorithmSid;
			}

			if (!salt || !spinCount) {
				let params = AscCommon.generateHashParams();
				salt = params.saltValue;
				spinCount = params.spinCount;
			}

			if (!alg) {
				alg = AscCommon.c_oSerCryptAlgorithmSid.SHA_512;
			}
		}

		let checkPassword = function (hash, doNotCheckPassword) {
			if (doNotCheckPassword) {
				callback(true);
			} else {
				if (props != null && props.edit != null && props.edit !== Asc.c_oAscEDocProtect.None) {
					//устанавливаем защиту
					calculatedHashValue = hash && hash[0] ? hash[0] : null;
					callback(true);
				} else {
					//пробуем снять защиту
					let documentHashValue = prepareHash(documentProtection.hashValue);
					if (documentProtection && hash && (hash[0] === documentHashValue || hash[1] === documentHashValue)) {
						salt = null;
						alg = null;
						spinCount = null;
						cryptProviderType = null;
						callback(true);
					} else {
						//неверный пароль
						t.sendEvent("asc_onError", c_oAscError.ID.PasswordIsNotCorrect, c_oAscError.Level.NoCritical);
						t.asc_OnProtectionUpdate();
						t.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction);
					}
				}
			}
		};

		this.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction);
		if (password != null) {
			if (password === "") {
				checkPassword([""]);
			} else {
				//перед тем, как сгенерировать хэш, мс предварительно преобразовывает пароль
				//в мс подходит как преобразованные пароль, так и не преобразованный
				//т.е. если сгенерировать вручную хэш из непреобразованного пароля и положить в xml, то документ можно будет разблокировать по первоначальному паролю
				let hashPassword = AscCommon.prepareWordPassword(password);
				let hashArr = [];
				if (hashPassword) {
					hashArr.push({password: hashPassword, salt: salt, spinCount: spinCount, alg: AscCommon.fromModelCryptAlgorithmSid(alg)});
				}
				hashArr.push({password: password, salt: salt, spinCount: spinCount, alg: AscCommon.fromModelCryptAlgorithmSid(alg)});
				AscCommon.calculateProtectHash(hashArr, checkPassword);
			}
		} else {
			checkPassword(null, true);
		}

		return true;
	};


	// print-preview
	asc_docs_api.prototype.asc_initPrintPreview = function(containerId, options)
	{
		if (this.printPreview)
			return;
		this.printPreview = new AscCommon.CPrintPreview(this, containerId);
	};
	asc_docs_api.prototype.asc_drawPrintPreview = function(index)
	{
		if (this.printPreview)
		{
			this.printPreview.page = index;
			this.printPreview.update();
		}
	};
	asc_docs_api.prototype.asc_closePrintPreview = function()
	{
		if (this.printPreview)
		{
			this.printPreview.close();
			delete this.printPreview;
		}
	};
	asc_docs_api.prototype.asc_getPageSize = function(pageIndex)
	{
		if (!this.WordControl
			|| !this.WordControl.m_oDrawingDocument
			|| this.WordControl.m_oDrawingDocument.IsFreezePage(pageIndex))
			return null;
		
		return {
			"W" : this.WordControl.m_oDrawingDocument.m_arrPages[pageIndex].width_mm,
			"H" : this.WordControl.m_oDrawingDocument.m_arrPages[pageIndex].height_mm
		};
	};
	asc_docs_api.prototype.getEyedropperImgData = function()
	{
		const oViewerCanvas = document.getElementById("id_viewer");
		const oOverlayCanvas = document.getElementById("id_viewer_overlay");
		if(!oViewerCanvas || !oOverlayCanvas)
		{
			return null;
		}
		let oCanvas = document.createElement("canvas");
		oCanvas.width = oViewerCanvas.width;
		oCanvas.height = oViewerCanvas.height;
		const oCtx = oCanvas.getContext("2d");
		oCtx.fillStyle = oViewerCanvas.style.backgroundColor;
		oCtx.fillRect(0, 0, oCanvas.width, oCanvas.height);
		oCtx.drawImage(oViewerCanvas, 0, 0);
		oCtx.drawImage(oOverlayCanvas, 0, 0);
		return oCtx.getImageData(0, 0, oCanvas.width, oCanvas.height);
	};
	
	// auto-hyphenation
	asc_docs_api.prototype.asc_setAutoHyphenation = function(isAuto)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		logicDocument.setAutoHyphenation(isAuto);
	};
	asc_docs_api.prototype.asc_isAutoHyphenation = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		return logicDocument ? logicDocument.getDocumentSettings().isAutoHyphenation() : false;
	};
	asc_docs_api.prototype.asc_setHyphenateCaps = function(isHyphenateCaps)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		logicDocument.setHyphenateCaps(isHyphenateCaps);
	};
	asc_docs_api.prototype.asc_isHyphenateCaps = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		return logicDocument ? logicDocument.getDocumentSettings().isHyphenateCaps() : true;
	};
	asc_docs_api.prototype.asc_setConsecutiveHyphenLimit = function(limit)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		logicDocument.setConsecutiveHyphenLimit(limit);
	};
	asc_docs_api.prototype.asc_getConsecutiveHyphenLimit = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		return logicDocument ? logicDocument.getDocumentSettings().getConsecutiveHyphenLimit() : 0;
	};
	asc_docs_api.prototype.asc_getAutoHyphenationSettings = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		return logicDocument ? logicDocument.getAutoHyphenationSettings() : new AscCommon.AutoHyphenationSettings();
	};
	asc_docs_api.prototype.asc_setAutoHyphenationSettings = function(settings)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		logicDocument.setAutoHyphenationSettings(settings);
	};
	
	asc_docs_api.prototype.getSelectionState = function()
	{
		return this.private_GetLogicDocument().GetSelectionState();
	};
	asc_docs_api.prototype.getSpeechDescription = function(prevState, action)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return null;
		
		return logicDocument.getSpeechDescription(prevState, action);
	};
	
	asc_docs_api.prototype.isTargetHandMode = function()
	{
		return this.isHandMode && this.isRestrictionForms();
	};


	asc_docs_api.prototype.asc_getCoHistory = function()
	{
		return AscCommon.CollaborativeEditing.getCoHistory();
	};
	asc_docs_api.prototype.asc_putPageColor = function(color)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		if (logicDocument.IsSelectionLocked(AscCommon.changestype_Document_Settings))
			return;
		
		logicDocument.StartAction(AscDFH.historydescription_Document_SetPageColor);
		
		if (!color)
		{
			logicDocument.setBackgroundColor(null, null);
		}
		else
		{
			if (color && color.asc_getType && color.asc_getType() === Asc.c_oAscColor.COLOR_TYPE_SRGB)
				logicDocument.setBackgroundColor(new AscCommonWord.CDocumentColor(color.r, color.g, color.b, false), null);
			else
				logicDocument.setBackgroundColor(null, AscFormat.CreateUnifillFromAscColor(color));
		}
		
		logicDocument.UpdateInterface();
		logicDocument.Redraw();
		logicDocument.FinalizeAction();
	};
	asc_docs_api.prototype.asc_getPageColor = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return null;
		
		return logicDocument.Background.getAscColor();
	};

	asc_docs_api.prototype.asc_insertTextFromFile = function () {
		const insertDocumentManager = new AscCommonWord.CInsertDocumentManager(this);
		insertDocumentManager.insertTextFromFile();
	};
	asc_docs_api.prototype.asc_insertTextFromUrl = function (url, token) {
		const insertDocumentManager = new AscCommonWord.CInsertDocumentManager(this);
		insertDocumentManager.insertTextFromUrl(url, token);
	};
	
	asc_docs_api.prototype.asc_showDeletedTextInVersionHistory = function()
	{
		if (!this.getVersionHistory())
			return false;
		
		return AscCommon.CollaborativeEditing.CoHistory.RecoverDeletedText()
	};
	asc_docs_api.prototype.asc_isShowedDeletedTextInVersionHistory = function()
	{
		if (!this.getVersionHistory())
			return false;
		
		return AscCommon.CollaborativeEditing.CoHistory.HaveDeletedTextRecovery()
	};
	asc_docs_api.prototype.asc_hideDeletedTextInVersionHistory = function()
	{
		if (!this.getVersionHistory())
			return;
		
		AscCommon.CollaborativeEditing.CoHistory.UndoDeletedTextRecovery();
	};
	asc_docs_api.prototype.getCustomProperties = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
			return null;
		return oLogicDocument.CustomProperties;
	};

	asc_docs_api.prototype.addCustomProperty = function(name, type, value)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
			return;
		oLogicDocument.AddCustomProperty(name, type, value);
	};

	asc_docs_api.prototype.modifyCustomProperty = function(idx, name, type, value)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
			return;
		oLogicDocument.ModifyCustomProperty(idx, name, type, value);
	};

	asc_docs_api.prototype.removeCustomProperty = function(idx)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
			return;
		oLogicDocument.RemoveCustomProperty(idx);
	};
	
	//-------------------------------------------------------------export---------------------------------------------------
	window['Asc']                                                       = window['Asc'] || {};
	CAscSection.prototype['get_PageWidth']                              = CAscSection.prototype.get_PageWidth;
	CAscSection.prototype['get_PageHeight']                             = CAscSection.prototype.get_PageHeight;
	CAscSection.prototype['get_MarginLeft']                             = CAscSection.prototype.get_MarginLeft;
	CAscSection.prototype['get_MarginRight']                            = CAscSection.prototype.get_MarginRight;
	CAscSection.prototype['get_MarginTop']                              = CAscSection.prototype.get_MarginTop;
	CAscSection.prototype['get_MarginBottom']                           = CAscSection.prototype.get_MarginBottom;
	Asc.CHeaderProp = CHeaderProp;
	CHeaderProp.prototype['get_Type']                                   = CHeaderProp.prototype.get_Type;
	CHeaderProp.prototype['put_Type']                                   = CHeaderProp.prototype.put_Type;
	CHeaderProp.prototype['get_Position']                               = CHeaderProp.prototype.get_Position;
	CHeaderProp.prototype['put_Position']                               = CHeaderProp.prototype.put_Position;
	CHeaderProp.prototype['get_DifferentFirst']                         = CHeaderProp.prototype.get_DifferentFirst;
	CHeaderProp.prototype['put_DifferentFirst']                         = CHeaderProp.prototype.put_DifferentFirst;
	CHeaderProp.prototype['get_DifferentEvenOdd']                       = CHeaderProp.prototype.get_DifferentEvenOdd;
	CHeaderProp.prototype['put_DifferentEvenOdd']                       = CHeaderProp.prototype.put_DifferentEvenOdd;
	CHeaderProp.prototype['get_LinkToPrevious']                         = CHeaderProp.prototype.get_LinkToPrevious;
	CHeaderProp.prototype['get_Locked']                                 = CHeaderProp.prototype.get_Locked;
	CHeaderProp.prototype['get_StartPageNumber']                        = CHeaderProp.prototype.get_StartPageNumber;
	CHeaderProp.prototype['put_StartPageNumber']                        = CHeaderProp.prototype.put_StartPageNumber;
	CHeaderProp.prototype['get_NumFormat']                              = CHeaderProp.prototype.get_NumFormat;
	CHeaderProp.prototype['put_NumFormat']                              = CHeaderProp.prototype.put_NumFormat;
	window['Asc']['CMailMergeSendData'] = window['Asc'].CMailMergeSendData = CMailMergeSendData;
	CMailMergeSendData.prototype['get_From']                            = CMailMergeSendData.prototype.get_From;
	CMailMergeSendData.prototype['put_From']                            = CMailMergeSendData.prototype.put_From;
	CMailMergeSendData.prototype['get_To']                              = CMailMergeSendData.prototype.get_To;
	CMailMergeSendData.prototype['put_To']                              = CMailMergeSendData.prototype.put_To;
	CMailMergeSendData.prototype['get_Subject']                         = CMailMergeSendData.prototype.get_Subject;
	CMailMergeSendData.prototype['put_Subject']                         = CMailMergeSendData.prototype.put_Subject;
	CMailMergeSendData.prototype['get_MailFormat']                      = CMailMergeSendData.prototype.get_MailFormat;
	CMailMergeSendData.prototype['put_MailFormat']                      = CMailMergeSendData.prototype.put_MailFormat;
	CMailMergeSendData.prototype['get_FileName']                        = CMailMergeSendData.prototype.get_FileName;
	CMailMergeSendData.prototype['put_FileName']                        = CMailMergeSendData.prototype.put_FileName;
	CMailMergeSendData.prototype['get_Message']                         = CMailMergeSendData.prototype.get_Message;
	CMailMergeSendData.prototype['put_Message']                         = CMailMergeSendData.prototype.put_Message;
	CMailMergeSendData.prototype['get_RecordFrom']                      = CMailMergeSendData.prototype.get_RecordFrom;
	CMailMergeSendData.prototype['put_RecordFrom']                      = CMailMergeSendData.prototype.put_RecordFrom;
	CMailMergeSendData.prototype['get_RecordTo']                        = CMailMergeSendData.prototype.get_RecordTo;
	CMailMergeSendData.prototype['put_RecordTo']                        = CMailMergeSendData.prototype.put_RecordTo;
	CMailMergeSendData.prototype['get_RecordCount']                     = CMailMergeSendData.prototype.get_RecordCount;
	CMailMergeSendData.prototype['put_RecordCount']                     = CMailMergeSendData.prototype.put_RecordCount;
	CMailMergeSendData.prototype['get_UserId']                          = CMailMergeSendData.prototype.get_UserId;
	CMailMergeSendData.prototype['put_UserId']                          = CMailMergeSendData.prototype.put_UserId;
	window['Asc']['CAscFootnotePr'] = window['Asc'].CAscFootnotePr = CAscFootnotePr;
	CAscFootnotePr.prototype['get_Pos']                                 = CAscFootnotePr.prototype.get_Pos;
	CAscFootnotePr.prototype['put_Pos']                                 = CAscFootnotePr.prototype.put_Pos;
	CAscFootnotePr.prototype['get_NumStart']                            = CAscFootnotePr.prototype.get_NumStart;
	CAscFootnotePr.prototype['put_NumStart']                            = CAscFootnotePr.prototype.put_NumStart;
	CAscFootnotePr.prototype['get_NumFormat']                           = CAscFootnotePr.prototype.get_NumFormat;
	CAscFootnotePr.prototype['put_NumFormat']                           = CAscFootnotePr.prototype.put_NumFormat;
	CAscFootnotePr.prototype['get_NumRestart']                          = CAscFootnotePr.prototype.get_NumRestart;
	CAscFootnotePr.prototype['put_NumRestart']                          = CAscFootnotePr.prototype.put_NumRestart;
	window['Asc']['asc_docs_api'] = window['Asc'].asc_docs_api          = asc_docs_api;
	asc_docs_api.prototype['SetCollaborativeMarksShowType']             = asc_docs_api.prototype.SetCollaborativeMarksShowType;
	asc_docs_api.prototype['GetCollaborativeMarksShowType']             = asc_docs_api.prototype.GetCollaborativeMarksShowType;
	asc_docs_api.prototype['Clear_CollaborativeMarks']                  = asc_docs_api.prototype.Clear_CollaborativeMarks;
	asc_docs_api.prototype['SetLanguage']                               = asc_docs_api.prototype.SetLanguage;
	asc_docs_api.prototype['asc_GetFontThumbnailsPath']                 = asc_docs_api.prototype.asc_GetFontThumbnailsPath;
	asc_docs_api.prototype['TranslateStyleName']                        = asc_docs_api.prototype.TranslateStyleName;
	asc_docs_api.prototype['CheckChangedDocument']                      = asc_docs_api.prototype.CheckChangedDocument;
	asc_docs_api.prototype['SetUnchangedDocument']                      = asc_docs_api.prototype.SetUnchangedDocument;
	asc_docs_api.prototype['SetDocumentModified']                       = asc_docs_api.prototype.SetDocumentModified;
	asc_docs_api.prototype['isDocumentModified']                        = asc_docs_api.prototype.isDocumentModified;
	asc_docs_api.prototype['asc_isDocumentCanSave']                     = asc_docs_api.prototype.asc_isDocumentCanSave;
	asc_docs_api.prototype['asc_getCanUndo']                            = asc_docs_api.prototype.asc_getCanUndo;
	asc_docs_api.prototype['asc_getCanRedo']                            = asc_docs_api.prototype.asc_getCanRedo;
	asc_docs_api.prototype['sync_BeginCatchSelectedElements']           = asc_docs_api.prototype.sync_BeginCatchSelectedElements;
	asc_docs_api.prototype['sync_EndCatchSelectedElements']             = asc_docs_api.prototype.sync_EndCatchSelectedElements;
	asc_docs_api.prototype['getSelectedElements']                       = asc_docs_api.prototype.getSelectedElements;
	asc_docs_api.prototype['sync_ChangeLastSelectedElement']            = asc_docs_api.prototype.sync_ChangeLastSelectedElement;
	asc_docs_api.prototype['asc_getEditorPermissions']                  = asc_docs_api.prototype.asc_getEditorPermissions;
	asc_docs_api.prototype['asc_setDocInfo']                            = asc_docs_api.prototype.asc_setDocInfo;
	asc_docs_api.prototype['asc_changeDocInfo']                         = asc_docs_api.prototype.asc_changeDocInfo;
	asc_docs_api.prototype['asc_setLocale']                             = asc_docs_api.prototype.asc_setLocale;
	asc_docs_api.prototype['asc_getLocale']                             = asc_docs_api.prototype.asc_getLocale;
	asc_docs_api.prototype['asc_LoadDocument']                          = asc_docs_api.prototype.asc_LoadDocument;
	asc_docs_api.prototype['SetTextBoxInputMode']                       = asc_docs_api.prototype.SetTextBoxInputMode;
	asc_docs_api.prototype['GetTextBoxInputMode']                       = asc_docs_api.prototype.GetTextBoxInputMode;
	asc_docs_api.prototype['ChangeReaderMode']                          = asc_docs_api.prototype.ChangeReaderMode;
	asc_docs_api.prototype['SetReaderModeOnly']                         = asc_docs_api.prototype.SetReaderModeOnly;
	asc_docs_api.prototype['IncreaseReaderFontSize']                    = asc_docs_api.prototype.IncreaseReaderFontSize;
	asc_docs_api.prototype['DecreaseReaderFontSize']                    = asc_docs_api.prototype.DecreaseReaderFontSize;
	asc_docs_api.prototype['SetMobileTopOffset']                    	= asc_docs_api.prototype.SetMobileTopOffset;
	asc_docs_api.prototype['CreateCSS']                                 = asc_docs_api.prototype.CreateCSS;
	asc_docs_api.prototype['GetCopyPasteDivId']                         = asc_docs_api.prototype.GetCopyPasteDivId;
	asc_docs_api.prototype['ContentToHTML']                             = asc_docs_api.prototype.ContentToHTML;
	asc_docs_api.prototype['InitEditor']                                = asc_docs_api.prototype.InitEditor;
	asc_docs_api.prototype['OpenDocument']                              = asc_docs_api.prototype.OpenDocument;
	asc_docs_api.prototype['OpenDocumentFromBin']                       = asc_docs_api.prototype.OpenDocumentFromBin;
	asc_docs_api.prototype['OpenDocumentFromZip']                       = asc_docs_api.prototype.OpenDocumentFromZip;
	asc_docs_api.prototype['asc_getDocumentName']                       = asc_docs_api.prototype.asc_getDocumentName;
	asc_docs_api.prototype['asc_getAppProps']                           = asc_docs_api.prototype.asc_getAppProps;
	asc_docs_api.prototype['asc_getCoreProps']                          = asc_docs_api.prototype.asc_getCoreProps;
	asc_docs_api.prototype['asc_setCoreProps']                          = asc_docs_api.prototype.asc_setCoreProps;
	asc_docs_api.prototype['asc_getPdfProps']                           = asc_docs_api.prototype.asc_getPdfProps;
	asc_docs_api.prototype['asc_isCompatibilityMode']                   = asc_docs_api.prototype.asc_isCompatibilityMode;
	asc_docs_api.prototype['asc_registerCallback']                      = asc_docs_api.prototype.asc_registerCallback;
	asc_docs_api.prototype['asc_unregisterCallback']                    = asc_docs_api.prototype.asc_unregisterCallback;
	asc_docs_api.prototype['asc_checkNeedCallback']                     = asc_docs_api.prototype.asc_checkNeedCallback;
	asc_docs_api.prototype['asc_getPropertyEditorShapes']               = asc_docs_api.prototype.asc_getPropertyEditorShapes;
	asc_docs_api.prototype['asc_getPropertyEditorTextArts']             = asc_docs_api.prototype.asc_getPropertyEditorTextArts;
	asc_docs_api.prototype['get_PropertyThemeColors']                   = asc_docs_api.prototype.get_PropertyThemeColors;
	asc_docs_api.prototype['_coAuthoringSetChange']                     = asc_docs_api.prototype._coAuthoringSetChange;
	asc_docs_api.prototype['_coAuthoringSetChanges']                    = asc_docs_api.prototype._coAuthoringSetChanges;
	asc_docs_api.prototype['asc_coAuthoringChatSendMessage']            = asc_docs_api.prototype.asc_coAuthoringChatSendMessage;
	asc_docs_api.prototype['asc_coAuthoringChatGetMessages']            = asc_docs_api.prototype.asc_coAuthoringChatGetMessages;
	asc_docs_api.prototype['asc_coAuthoringGetUsers']                   = asc_docs_api.prototype.asc_coAuthoringGetUsers;
	asc_docs_api.prototype['asc_coAuthoringDisconnect']                 = asc_docs_api.prototype.asc_coAuthoringDisconnect;
	asc_docs_api.prototype['asc_SpellCheckDisconnect']                  = asc_docs_api.prototype.asc_SpellCheckDisconnect;
	asc_docs_api.prototype['_onUpdateDocumentCanSave']                  = asc_docs_api.prototype._onUpdateDocumentCanSave;
	asc_docs_api.prototype['put_FramePr']                               = asc_docs_api.prototype.put_FramePr;
	asc_docs_api.prototype['asyncFontEndLoaded_MathDraw']               = asc_docs_api.prototype.asyncFontEndLoaded_MathDraw;
	asc_docs_api.prototype['sendMathTypesToMenu']                       = asc_docs_api.prototype.sendMathTypesToMenu;
	asc_docs_api.prototype['asyncFontEndLoaded_DropCap']                = asc_docs_api.prototype.asyncFontEndLoaded_DropCap;
	asc_docs_api.prototype['asc_addDropCap']                            = asc_docs_api.prototype.asc_addDropCap;
	asc_docs_api.prototype['removeDropcap']                             = asc_docs_api.prototype.removeDropcap;
	asc_docs_api.prototype['get_TextProps']                             = asc_docs_api.prototype.get_TextProps;
	asc_docs_api.prototype['GetJSONLogicDocument']                      = asc_docs_api.prototype.GetJSONLogicDocument;
	asc_docs_api.prototype['get_ContentCount']                          = asc_docs_api.prototype.get_ContentCount;
	asc_docs_api.prototype['select_Element']                            = asc_docs_api.prototype.select_Element;
	asc_docs_api.prototype['UpdateTextPr']                              = asc_docs_api.prototype.UpdateTextPr;
	asc_docs_api.prototype['UpdateParagraphProp']                       = asc_docs_api.prototype.UpdateParagraphProp;
	asc_docs_api.prototype['Undo']                                      = asc_docs_api.prototype.Undo;
	asc_docs_api.prototype['Redo']                                      = asc_docs_api.prototype.Redo;
	asc_docs_api.prototype['Copy']                                      = asc_docs_api.prototype.Copy;
	asc_docs_api.prototype['Update_ParaTab']                            = asc_docs_api.prototype.Update_ParaTab;
	asc_docs_api.prototype['Cut']                                       = asc_docs_api.prototype.Cut;
	asc_docs_api.prototype['Paste']                                     = asc_docs_api.prototype.Paste;
	asc_docs_api.prototype['Share']                                     = asc_docs_api.prototype.Share;
	asc_docs_api.prototype['asc_Save']                                  = asc_docs_api.prototype.asc_Save;
	asc_docs_api.prototype['forceSave']                                 = asc_docs_api.prototype.forceSave;
	asc_docs_api.prototype['asc_setIsForceSaveOnUserSave']              = asc_docs_api.prototype.asc_setIsForceSaveOnUserSave;
	asc_docs_api.prototype['asc_DownloadAs']                            = asc_docs_api.prototype.asc_DownloadAs;
	asc_docs_api.prototype['asc_DownloadAsMailMerge']                   = asc_docs_api.prototype.asc_DownloadAsMailMerge;
	asc_docs_api.prototype['asc_DownloadOrigin']                        = asc_docs_api.prototype.asc_DownloadOrigin;
	asc_docs_api.prototype['Resize']                                    = asc_docs_api.prototype.Resize;
	asc_docs_api.prototype['AddURL']                                    = asc_docs_api.prototype.AddURL;
	asc_docs_api.prototype['Help']                                      = asc_docs_api.prototype.Help;
	asc_docs_api.prototype['asc_setAdvancedOptions']                    = asc_docs_api.prototype.asc_setAdvancedOptions;
	asc_docs_api.prototype['asc_decodeBuffer']                    		= asc_docs_api.prototype.asc_decodeBuffer;
	asc_docs_api.prototype['SetFontRenderingMode']                      = asc_docs_api.prototype.SetFontRenderingMode;
	asc_docs_api.prototype['startGetDocInfo']                           = asc_docs_api.prototype.startGetDocInfo;
	asc_docs_api.prototype['stopGetDocInfo']                            = asc_docs_api.prototype.stopGetDocInfo;
	asc_docs_api.prototype['sync_DocInfoCallback']                      = asc_docs_api.prototype.sync_DocInfoCallback;
	asc_docs_api.prototype['sync_GetDocInfoStartCallback']              = asc_docs_api.prototype.sync_GetDocInfoStartCallback;
	asc_docs_api.prototype['sync_GetDocInfoStopCallback']               = asc_docs_api.prototype.sync_GetDocInfoStopCallback;
	asc_docs_api.prototype['sync_GetDocInfoEndCallback']                = asc_docs_api.prototype.sync_GetDocInfoEndCallback;
	asc_docs_api.prototype['sync_CanUndoCallback']                      = asc_docs_api.prototype.sync_CanUndoCallback;
	asc_docs_api.prototype['sync_CanRedoCallback']                      = asc_docs_api.prototype.sync_CanRedoCallback;
	asc_docs_api.prototype['can_CopyCut']                               = asc_docs_api.prototype.can_CopyCut;
	asc_docs_api.prototype['sync_CanCopyCutCallback']                   = asc_docs_api.prototype.sync_CanCopyCutCallback;
	asc_docs_api.prototype['setStartPointHistory']                      = asc_docs_api.prototype.setStartPointHistory;
	asc_docs_api.prototype['setEndPointHistory']                        = asc_docs_api.prototype.setEndPointHistory;
	asc_docs_api.prototype['sync_CursorLockCallBack']                   = asc_docs_api.prototype.sync_CursorLockCallBack;
	asc_docs_api.prototype['sync_UndoCallBack']                         = asc_docs_api.prototype.sync_UndoCallBack;
	asc_docs_api.prototype['sync_RedoCallBack']                         = asc_docs_api.prototype.sync_RedoCallBack;
	asc_docs_api.prototype['sync_CopyCallBack']                         = asc_docs_api.prototype.sync_CopyCallBack;
	asc_docs_api.prototype['sync_CutCallBack']                          = asc_docs_api.prototype.sync_CutCallBack;
	asc_docs_api.prototype['sync_PasteCallBack']                        = asc_docs_api.prototype.sync_PasteCallBack;
	asc_docs_api.prototype['sync_ShareCallBack']                        = asc_docs_api.prototype.sync_ShareCallBack;
	asc_docs_api.prototype['sync_SaveCallBack']                         = asc_docs_api.prototype.sync_SaveCallBack;
	asc_docs_api.prototype['sync_DownloadAsCallBack']                   = asc_docs_api.prototype.sync_DownloadAsCallBack;
	asc_docs_api.prototype['sync_StartAction']                          = asc_docs_api.prototype.sync_StartAction;
	asc_docs_api.prototype['sync_EndAction']                            = asc_docs_api.prototype.sync_EndAction;
	asc_docs_api.prototype['sync_AddURLCallback']                       = asc_docs_api.prototype.sync_AddURLCallback;
	asc_docs_api.prototype['sync_ErrorCallback']                        = asc_docs_api.prototype.sync_ErrorCallback;
	asc_docs_api.prototype['sync_HelpCallback']                         = asc_docs_api.prototype.sync_HelpCallback;
	asc_docs_api.prototype['sync_UpdateZoom']                           = asc_docs_api.prototype.sync_UpdateZoom;
	asc_docs_api.prototype['ClearPropObjCallback']                      = asc_docs_api.prototype.ClearPropObjCallback;
	asc_docs_api.prototype['CollectHeaders']                            = asc_docs_api.prototype.CollectHeaders;
	asc_docs_api.prototype['GetActiveHeader']                           = asc_docs_api.prototype.GetActiveHeader;
	asc_docs_api.prototype['gotoHeader']                                = asc_docs_api.prototype.gotoHeader;
	asc_docs_api.prototype['sync_ChangeActiveHeaderCallback']           = asc_docs_api.prototype.sync_ChangeActiveHeaderCallback;
	asc_docs_api.prototype['sync_ReturnHeadersCallback']                = asc_docs_api.prototype.sync_ReturnHeadersCallback;
	asc_docs_api.prototype['asc_searchEnabled']                         = asc_docs_api.prototype.asc_searchEnabled;
	asc_docs_api.prototype['asc_findText']                              = asc_docs_api.prototype.asc_findText;
	asc_docs_api.prototype['asc_endFindText']                           = asc_docs_api.prototype.asc_endFindText;
	asc_docs_api.prototype['asc_replaceText']                           = asc_docs_api.prototype.asc_replaceText;
	asc_docs_api.prototype['asc_GetErrorForReplaceString']              = asc_docs_api.prototype.asc_GetErrorForReplaceString;
	asc_docs_api.prototype['asc_isSelectSearchingResults']              = asc_docs_api.prototype.asc_isSelectSearchingResults;
	asc_docs_api.prototype['asc_StartTextAroundSearch']                 = asc_docs_api.prototype.asc_StartTextAroundSearch;
	asc_docs_api.prototype['asc_SelectSearchElement']                   = asc_docs_api.prototype.asc_SelectSearchElement;
	asc_docs_api.prototype['sync_ReplaceAllCallback']                   = asc_docs_api.prototype.sync_ReplaceAllCallback;
	asc_docs_api.prototype['sync_SearchEndCallback']                    = asc_docs_api.prototype.sync_SearchEndCallback;
	asc_docs_api.prototype['put_TextPrFontName']                        = asc_docs_api.prototype.put_TextPrFontName;
	asc_docs_api.prototype['put_TextPrFontSize']                        = asc_docs_api.prototype.put_TextPrFontSize;
	asc_docs_api.prototype['put_TextPrBold']                            = asc_docs_api.prototype.put_TextPrBold;
	asc_docs_api.prototype['put_TextPrItalic']                          = asc_docs_api.prototype.put_TextPrItalic;
	asc_docs_api.prototype['put_TextPrUnderline']                       = asc_docs_api.prototype.put_TextPrUnderline;
	asc_docs_api.prototype['put_TextPrStrikeout']                       = asc_docs_api.prototype.put_TextPrStrikeout;
	asc_docs_api.prototype['put_TextPrDStrikeout']                      = asc_docs_api.prototype.put_TextPrDStrikeout;
	asc_docs_api.prototype['put_TextPrSpacing']                         = asc_docs_api.prototype.put_TextPrSpacing;
	asc_docs_api.prototype['put_TextPrCaps']                            = asc_docs_api.prototype.put_TextPrCaps;
	asc_docs_api.prototype['put_TextPrSmallCaps']                       = asc_docs_api.prototype.put_TextPrSmallCaps;
	asc_docs_api.prototype['put_TextPrPosition']                        = asc_docs_api.prototype.put_TextPrPosition;
	asc_docs_api.prototype['put_TextPrLang']                            = asc_docs_api.prototype.put_TextPrLang;
	asc_docs_api.prototype['asc_setRtlTextDirection']                   = asc_docs_api.prototype.asc_setRtlTextDirection;
	asc_docs_api.prototype['asc_isRtlTextDirection']                    = asc_docs_api.prototype.asc_isRtlTextDirection;
	asc_docs_api.prototype['asc_addSpaceBeforeParagraph']               = asc_docs_api.prototype.asc_addSpaceBeforeParagraph;
	asc_docs_api.prototype['asc_addSpaceAfterParagraph']                = asc_docs_api.prototype.asc_addSpaceAfterParagraph;
	asc_docs_api.prototype['asc_removeSpaceBeforeParagraph']            = asc_docs_api.prototype.asc_removeSpaceBeforeParagraph;
	asc_docs_api.prototype['asc_removeSpaceAfterParagraph']             = asc_docs_api.prototype.asc_removeSpaceAfterParagraph;
	asc_docs_api.prototype['asc_haveSpaceBeforeParagraph']              = asc_docs_api.prototype.asc_haveSpaceBeforeParagraph;
	asc_docs_api.prototype['asc_haveSpaceAfterParagraph']               = asc_docs_api.prototype.asc_haveSpaceAfterParagraph;
	asc_docs_api.prototype['put_PrLineSpacing']                         = asc_docs_api.prototype.put_PrLineSpacing;
	asc_docs_api.prototype['put_LineSpacingBeforeAfter']                = asc_docs_api.prototype.put_LineSpacingBeforeAfter;
	asc_docs_api.prototype['FontSizeIn']                                = asc_docs_api.prototype.FontSizeIn;
	asc_docs_api.prototype['FontSizeOut']                               = asc_docs_api.prototype.FontSizeOut;
	asc_docs_api.prototype['put_Borders']                               = asc_docs_api.prototype.put_Borders;
	asc_docs_api.prototype['sync_BoldCallBack']                         = asc_docs_api.prototype.sync_BoldCallBack;
	asc_docs_api.prototype['sync_ItalicCallBack']                       = asc_docs_api.prototype.sync_ItalicCallBack;
	asc_docs_api.prototype['sync_UnderlineCallBack']                    = asc_docs_api.prototype.sync_UnderlineCallBack;
	asc_docs_api.prototype['sync_StrikeoutCallBack']                    = asc_docs_api.prototype.sync_StrikeoutCallBack;
	asc_docs_api.prototype['sync_TextPrFontFamilyCallBack']             = asc_docs_api.prototype.sync_TextPrFontFamilyCallBack;
	asc_docs_api.prototype['sync_TextPrFontSizeCallBack']               = asc_docs_api.prototype.sync_TextPrFontSizeCallBack;
	asc_docs_api.prototype['sync_PrLineSpacingCallBack']                = asc_docs_api.prototype.sync_PrLineSpacingCallBack;
	asc_docs_api.prototype['paraApply']                                 = asc_docs_api.prototype.paraApply;
	asc_docs_api.prototype['put_PrAlign']                               = asc_docs_api.prototype.put_PrAlign;
	asc_docs_api.prototype['put_TextPrBaseline']                        = asc_docs_api.prototype.put_TextPrBaseline;
	asc_docs_api.prototype['asc_IsShowListIndentsSettings']             = asc_docs_api.prototype.asc_IsShowListIndentsSettings;
	asc_docs_api.prototype['put_ListType']                              = asc_docs_api.prototype.put_ListType;
	asc_docs_api.prototype['put_ListTypeCustom']                        = asc_docs_api.prototype.put_ListTypeCustom;
	asc_docs_api.prototype['asc_GetCurrentNumberingJson']               = asc_docs_api.prototype.asc_GetCurrentNumberingJson;
	asc_docs_api.prototype['asc_IsCurrentNumberingPreset']              = asc_docs_api.prototype.asc_IsCurrentNumberingPreset;
	asc_docs_api.prototype['asc_CompareNumberingPresets']               = asc_docs_api.prototype.asc_CompareNumberingPresets;
	asc_docs_api.prototype['asc_ContinueNumbering']                     = asc_docs_api.prototype.asc_ContinueNumbering;
	asc_docs_api.prototype['asc_RestartNumbering']                      = asc_docs_api.prototype.asc_RestartNumbering;
	asc_docs_api.prototype['asc_GetCurrentNumberingId']                 = asc_docs_api.prototype.asc_GetCurrentNumberingId;
	asc_docs_api.prototype['asc_GetCurrentNumberingLvl']                = asc_docs_api.prototype.asc_GetCurrentNumberingLvl;
	asc_docs_api.prototype['asc_GetCalculatedNumberingValue']           = asc_docs_api.prototype.asc_GetCalculatedNumberingValue;
	asc_docs_api.prototype['asc_GetNumberingPr']                        = asc_docs_api.prototype.asc_GetNumberingPr;
	asc_docs_api.prototype['asc_AddNewNumbering']                       = asc_docs_api.prototype.asc_AddNewNumbering;
	asc_docs_api.prototype['asc_ChangeNumberingLvl']                    = asc_docs_api.prototype.asc_ChangeNumberingLvl;
	asc_docs_api.prototype['asc_SetNumberingLvl']                       = asc_docs_api.prototype.asc_SetNumberingLvl;
	asc_docs_api.prototype['put_Style']                                 = asc_docs_api.prototype.put_Style;
	asc_docs_api.prototype['asc_SetParagraphSuppressLineNumbers']       = asc_docs_api.prototype.asc_SetParagraphSuppressLineNumbers;
	asc_docs_api.prototype['SetDeviceInputHelperId']                    = asc_docs_api.prototype.SetDeviceInputHelperId;
	asc_docs_api.prototype['put_ShowSnapLines']                         = asc_docs_api.prototype.put_ShowSnapLines;
	asc_docs_api.prototype['get_ShowSnapLines']                         = asc_docs_api.prototype.get_ShowSnapLines;
	asc_docs_api.prototype['put_ShowSmartGuides']                       = asc_docs_api.prototype.put_ShowSmartGuides;
	asc_docs_api.prototype['get_ShowSmartGuides']                       = asc_docs_api.prototype.get_ShowSmartGuides;
	asc_docs_api.prototype['put_ShowParaMarks']                         = asc_docs_api.prototype.put_ShowParaMarks;
	asc_docs_api.prototype['get_ShowParaMarks']                         = asc_docs_api.prototype.get_ShowParaMarks;
	asc_docs_api.prototype['put_ShowTableEmptyLine']                    = asc_docs_api.prototype.put_ShowTableEmptyLine;
	asc_docs_api.prototype['get_ShowTableEmptyLine']                    = asc_docs_api.prototype.get_ShowTableEmptyLine;
	asc_docs_api.prototype['put_PageBreak']                             = asc_docs_api.prototype.put_PageBreak;
	asc_docs_api.prototype['put_WidowControl']                          = asc_docs_api.prototype.put_WidowControl;
	asc_docs_api.prototype['put_KeepLines']                             = asc_docs_api.prototype.put_KeepLines;
	asc_docs_api.prototype['put_KeepNext']                              = asc_docs_api.prototype.put_KeepNext;
	asc_docs_api.prototype['put_AddSpaceBetweenPrg']                    = asc_docs_api.prototype.put_AddSpaceBetweenPrg;
	asc_docs_api.prototype['put_LineHighLight']                         = asc_docs_api.prototype.put_LineHighLight;
	asc_docs_api.prototype['put_TextColor']                             = asc_docs_api.prototype.put_TextColor;
	asc_docs_api.prototype['put_ParagraphShade']                        = asc_docs_api.prototype.put_ParagraphShade;
	asc_docs_api.prototype['put_PrIndent']                              = asc_docs_api.prototype.put_PrIndent;
	asc_docs_api.prototype['put_ParagraphOutlineLvl']                   = asc_docs_api.prototype.put_ParagraphOutlineLvl;
	asc_docs_api.prototype['IncreaseIndent']                            = asc_docs_api.prototype.IncreaseIndent;
	asc_docs_api.prototype['DecreaseIndent']                            = asc_docs_api.prototype.DecreaseIndent;
	asc_docs_api.prototype['put_PrIndentRight']                         = asc_docs_api.prototype.put_PrIndentRight;
	asc_docs_api.prototype['put_PrFirstLineIndent']                     = asc_docs_api.prototype.put_PrFirstLineIndent;
	asc_docs_api.prototype['put_Margins']                               = asc_docs_api.prototype.put_Margins;
	asc_docs_api.prototype['getFocusObject']                            = asc_docs_api.prototype.getFocusObject;
	asc_docs_api.prototype['sync_VerticalAlign']                        = asc_docs_api.prototype.sync_VerticalAlign;
	asc_docs_api.prototype['sync_PrAlignCallBack']                      = asc_docs_api.prototype.sync_PrAlignCallBack;
	asc_docs_api.prototype['sync_TextColor']                            = asc_docs_api.prototype.sync_TextColor;
	asc_docs_api.prototype['sync_TextHighLight']                        = asc_docs_api.prototype.sync_TextHighLight;
	asc_docs_api.prototype['sync_TextSpacing']                          = asc_docs_api.prototype.sync_TextSpacing;
	asc_docs_api.prototype['sync_TextDStrikeout']                       = asc_docs_api.prototype.sync_TextDStrikeout;
	asc_docs_api.prototype['sync_TextCaps']                             = asc_docs_api.prototype.sync_TextCaps;
	asc_docs_api.prototype['sync_TextSmallCaps']                        = asc_docs_api.prototype.sync_TextSmallCaps;
	asc_docs_api.prototype['sync_TextPosition']                         = asc_docs_api.prototype.sync_TextPosition;
	asc_docs_api.prototype['sync_TextLangCallBack']                     = asc_docs_api.prototype.sync_TextLangCallBack;
	asc_docs_api.prototype['sync_ParaStyleName']                        = asc_docs_api.prototype.sync_ParaStyleName;
	asc_docs_api.prototype['sync_ParaSpacingLine']                      = asc_docs_api.prototype.sync_ParaSpacingLine;
	asc_docs_api.prototype['sync_PageBreakCallback']                    = asc_docs_api.prototype.sync_PageBreakCallback;
	asc_docs_api.prototype['sync_WidowControlCallback']                 = asc_docs_api.prototype.sync_WidowControlCallback;
	asc_docs_api.prototype['sync_KeepNextCallback']                     = asc_docs_api.prototype.sync_KeepNextCallback;
	asc_docs_api.prototype['sync_KeepLinesCallback']                    = asc_docs_api.prototype.sync_KeepLinesCallback;
	asc_docs_api.prototype['sync_ShowParaMarksCallback']                = asc_docs_api.prototype.sync_ShowParaMarksCallback;
	asc_docs_api.prototype['sync_SpaceBetweenPrgCallback']              = asc_docs_api.prototype.sync_SpaceBetweenPrgCallback;
	asc_docs_api.prototype['sync_RangePermPropCallback']                = asc_docs_api.prototype.sync_RangePermPropCallback;
	asc_docs_api.prototype['sync_PrPropCallback']                       = asc_docs_api.prototype.sync_PrPropCallback;
	asc_docs_api.prototype['sync_MathPropCallback']                     = asc_docs_api.prototype.sync_MathPropCallback;
	asc_docs_api.prototype['sync_EndAddShape']                          = asc_docs_api.prototype.sync_EndAddShape;
	asc_docs_api.prototype['SetDrawingFreeze']                          = asc_docs_api.prototype.SetDrawingFreeze;
	asc_docs_api.prototype['change_PageOrient']                         = asc_docs_api.prototype.change_PageOrient;
	asc_docs_api.prototype['get_DocumentOrientation']                   = asc_docs_api.prototype.get_DocumentOrientation;
	asc_docs_api.prototype['change_DocSize']                            = asc_docs_api.prototype.change_DocSize;
	asc_docs_api.prototype['get_DocumentWidth']                         = asc_docs_api.prototype.get_DocumentWidth;
	asc_docs_api.prototype['get_DocumentHeight']                        = asc_docs_api.prototype.get_DocumentHeight;
	asc_docs_api.prototype['put_AddPageBreak']                          = asc_docs_api.prototype.put_AddPageBreak;
	asc_docs_api.prototype['put_AddColumnBreak']                        = asc_docs_api.prototype.put_AddColumnBreak;
	asc_docs_api.prototype['Update_ParaInd']                            = asc_docs_api.prototype.Update_ParaInd;
	asc_docs_api.prototype['Internal_Update_Ind_FirstLine']             = asc_docs_api.prototype.Internal_Update_Ind_FirstLine;
	asc_docs_api.prototype['Internal_Update_Ind_Left']                  = asc_docs_api.prototype.Internal_Update_Ind_Left;
	asc_docs_api.prototype['Internal_Update_Ind_Right']                 = asc_docs_api.prototype.Internal_Update_Ind_Right;
	asc_docs_api.prototype['put_PageNum']                               = asc_docs_api.prototype.put_PageNum;
	asc_docs_api.prototype['put_HeadersAndFootersDistance']             = asc_docs_api.prototype.put_HeadersAndFootersDistance;
	asc_docs_api.prototype['HeadersAndFooters_DifferentFirstPage']      = asc_docs_api.prototype.HeadersAndFooters_DifferentFirstPage;
	asc_docs_api.prototype['HeadersAndFooters_DifferentOddandEvenPage'] = asc_docs_api.prototype.HeadersAndFooters_DifferentOddandEvenPage;
	asc_docs_api.prototype['HeadersAndFooters_LinkToPrevious']          = asc_docs_api.prototype.HeadersAndFooters_LinkToPrevious;
	asc_docs_api.prototype['asc_SetSectionStartPage']                   = asc_docs_api.prototype.asc_SetSectionStartPage;
	asc_docs_api.prototype['asc_SetSectionPageNumFormat']               = asc_docs_api.prototype.asc_SetSectionPageNumFormat;
	asc_docs_api.prototype['sync_DocSizeCallback']                      = asc_docs_api.prototype.sync_DocSizeCallback;
	asc_docs_api.prototype['sync_PageOrientCallback']                   = asc_docs_api.prototype.sync_PageOrientCallback;
	asc_docs_api.prototype['sync_HeadersAndFootersPropCallback']        = asc_docs_api.prototype.sync_HeadersAndFootersPropCallback;
	asc_docs_api.prototype['put_Table']                                 = asc_docs_api.prototype.put_Table;
	asc_docs_api.prototype['addRowAbove']                               = asc_docs_api.prototype.addRowAbove;
	asc_docs_api.prototype['addRowBelow']                               = asc_docs_api.prototype.addRowBelow;
	asc_docs_api.prototype['addColumnLeft']                             = asc_docs_api.prototype.addColumnLeft;
	asc_docs_api.prototype['addColumnRight']                            = asc_docs_api.prototype.addColumnRight;
	asc_docs_api.prototype['remRow']                                    = asc_docs_api.prototype.remRow;
	asc_docs_api.prototype['remColumn']                                 = asc_docs_api.prototype.remColumn;
	asc_docs_api.prototype['remTable']                                  = asc_docs_api.prototype.remTable;
	asc_docs_api.prototype['selectRow']                                 = asc_docs_api.prototype.selectRow;
	asc_docs_api.prototype['selectColumn']                              = asc_docs_api.prototype.selectColumn;
	asc_docs_api.prototype['selectCell']                                = asc_docs_api.prototype.selectCell;
	asc_docs_api.prototype['selectTable']                               = asc_docs_api.prototype.selectTable;
	asc_docs_api.prototype['setColumnWidth']                            = asc_docs_api.prototype.setColumnWidth;
	asc_docs_api.prototype['setRowHeight']                              = asc_docs_api.prototype.setRowHeight;
	asc_docs_api.prototype['set_TblDistanceFromText']                   = asc_docs_api.prototype.set_TblDistanceFromText;
	asc_docs_api.prototype['CheckBeforeMergeCells']                     = asc_docs_api.prototype.CheckBeforeMergeCells;
	asc_docs_api.prototype['CheckBeforeSplitCells']                     = asc_docs_api.prototype.CheckBeforeSplitCells;
	asc_docs_api.prototype['MergeCells']                                = asc_docs_api.prototype.MergeCells;
	asc_docs_api.prototype['SplitCell']                                 = asc_docs_api.prototype.SplitCell;
	asc_docs_api.prototype['asc_DistributeTableCells']                  = asc_docs_api.prototype.asc_DistributeTableCells;
	asc_docs_api.prototype['asc_RemoveTableCells']                      = asc_docs_api.prototype.asc_RemoveTableCells;
	asc_docs_api.prototype['widthTable']                                = asc_docs_api.prototype.widthTable;
	asc_docs_api.prototype['put_CellsMargin']                           = asc_docs_api.prototype.put_CellsMargin;
	asc_docs_api.prototype['set_TblWrap']                               = asc_docs_api.prototype.set_TblWrap;
	asc_docs_api.prototype['set_TblIndentLeft']                         = asc_docs_api.prototype.set_TblIndentLeft;
	asc_docs_api.prototype['set_Borders']                               = asc_docs_api.prototype.set_Borders;
	asc_docs_api.prototype['set_TableBackground']                       = asc_docs_api.prototype.set_TableBackground;
	asc_docs_api.prototype['set_AlignCell']                             = asc_docs_api.prototype.set_AlignCell;
	asc_docs_api.prototype['set_TblAlign']                              = asc_docs_api.prototype.set_TblAlign;
	asc_docs_api.prototype['set_SpacingBetweenCells']                   = asc_docs_api.prototype.set_SpacingBetweenCells;
	asc_docs_api.prototype['tblApply']                                  = asc_docs_api.prototype.tblApply;
	asc_docs_api.prototype['sync_AddTableCallback']                     = asc_docs_api.prototype.sync_AddTableCallback;
	asc_docs_api.prototype['sync_AlignCellCallback']                    = asc_docs_api.prototype.sync_AlignCellCallback;
	asc_docs_api.prototype['sync_TblPropCallback']                      = asc_docs_api.prototype.sync_TblPropCallback;
	asc_docs_api.prototype['sync_TblWrapStyleChangedCallback']          = asc_docs_api.prototype.sync_TblWrapStyleChangedCallback;
	asc_docs_api.prototype['sync_TblAlignChangedCallback']              = asc_docs_api.prototype.sync_TblAlignChangedCallback;
	asc_docs_api.prototype['ChangeImageFromFile']                       = asc_docs_api.prototype.ChangeImageFromFile;
	asc_docs_api.prototype['ChangeShapeImageFromFile']                  = asc_docs_api.prototype.ChangeShapeImageFromFile;
	asc_docs_api.prototype['AddImage']                                  = asc_docs_api.prototype.AddImage;
	asc_docs_api.prototype['asc_addImage']                              = asc_docs_api.prototype.asc_addImage;
	asc_docs_api.prototype['AddImageUrl']                               = asc_docs_api.prototype.AddImageUrl;
	asc_docs_api.prototype['AddImageUrlAction']                         = asc_docs_api.prototype.AddImageUrlAction;
	asc_docs_api.prototype['AddImageToPage']                            = asc_docs_api.prototype.AddImageToPage;
	asc_docs_api.prototype['asc_getSelectedDrawingObjectsCount']        = asc_docs_api.prototype.asc_getSelectedDrawingObjectsCount;
	asc_docs_api.prototype['put_ShapesAlign']                           = asc_docs_api.prototype.put_ShapesAlign;
	asc_docs_api.prototype['DistributeHorizontally']                    = asc_docs_api.prototype.DistributeHorizontally;
	asc_docs_api.prototype['DistributeVertically']                      = asc_docs_api.prototype.DistributeVertically;
	asc_docs_api.prototype['ImgApply']                                  = asc_docs_api.prototype.ImgApply;
	asc_docs_api.prototype['set_Size']                                  = asc_docs_api.prototype.set_Size;
	asc_docs_api.prototype['set_ConstProportions']                      = asc_docs_api.prototype.set_ConstProportions;
	asc_docs_api.prototype['set_WrapStyle']                             = asc_docs_api.prototype.set_WrapStyle;
	asc_docs_api.prototype['deleteImage']                               = asc_docs_api.prototype.deleteImage;
	asc_docs_api.prototype['set_ImgDistanceFromText']                   = asc_docs_api.prototype.set_ImgDistanceFromText;
	asc_docs_api.prototype['set_PositionOnPage']                        = asc_docs_api.prototype.set_PositionOnPage;
	asc_docs_api.prototype['get_OriginalSizeImage']                     = asc_docs_api.prototype.get_OriginalSizeImage;
	asc_docs_api.prototype['ShapeApply']                                = asc_docs_api.prototype.ShapeApply;
	asc_docs_api.prototype['sync_AddImageCallback']                     = asc_docs_api.prototype.sync_AddImageCallback;
	asc_docs_api.prototype['sync_ImgPropCallback']                      = asc_docs_api.prototype.sync_ImgPropCallback;
	asc_docs_api.prototype['sync_ImgWrapStyleChangedCallback']          = asc_docs_api.prototype.sync_ImgWrapStyleChangedCallback;
	asc_docs_api.prototype['sync_ContextMenuCallback']                  = asc_docs_api.prototype.sync_ContextMenuCallback;
	asc_docs_api.prototype['sync_MouseMoveStartCallback']               = asc_docs_api.prototype.sync_MouseMoveStartCallback;
	asc_docs_api.prototype['sync_MouseMoveEndCallback']                 = asc_docs_api.prototype.sync_MouseMoveEndCallback;
	asc_docs_api.prototype['sync_MouseMoveCallback']                    = asc_docs_api.prototype.sync_MouseMoveCallback;
	asc_docs_api.prototype['can_AddHyperlink']                          = asc_docs_api.prototype.can_AddHyperlink;
	asc_docs_api.prototype['add_Hyperlink']                             = asc_docs_api.prototype.add_Hyperlink;
	asc_docs_api.prototype['change_Hyperlink']                          = asc_docs_api.prototype.change_Hyperlink;
	asc_docs_api.prototype['remove_Hyperlink']                          = asc_docs_api.prototype.remove_Hyperlink;
	asc_docs_api.prototype['asc_GetHyperlinkAnchors']                   = asc_docs_api.prototype.asc_GetHyperlinkAnchors;
	asc_docs_api.prototype['sync_HyperlinkPropCallback']                = asc_docs_api.prototype.sync_HyperlinkPropCallback;
	asc_docs_api.prototype['sync_HyperlinkClickCallback']               = asc_docs_api.prototype.sync_HyperlinkClickCallback;
	asc_docs_api.prototype['sync_CanAddHyperlinkCallback']              = asc_docs_api.prototype.sync_CanAddHyperlinkCallback;
	asc_docs_api.prototype['sync_DialogAddHyperlink']                   = asc_docs_api.prototype.sync_DialogAddHyperlink;
	asc_docs_api.prototype['sync_DialogAddHyperlink']                   = asc_docs_api.prototype.sync_DialogAddHyperlink;
	asc_docs_api.prototype['sync_SpellCheckCallback']                   = asc_docs_api.prototype.sync_SpellCheckCallback;
	asc_docs_api.prototype['sync_SpellCheckVariantsFound']              = asc_docs_api.prototype.sync_SpellCheckVariantsFound;
	asc_docs_api.prototype['asc_replaceMisspelledWord']                 = asc_docs_api.prototype.asc_replaceMisspelledWord;
	asc_docs_api.prototype['asc_ignoreMisspelledWord']                  = asc_docs_api.prototype.asc_ignoreMisspelledWord;
    asc_docs_api.prototype['asc_spellCheckAddToDictionary']       		= asc_docs_api.prototype.asc_spellCheckAddToDictionary;
    asc_docs_api.prototype['asc_spellCheckClearDictionary']       		= asc_docs_api.prototype.asc_spellCheckClearDictionary;
	asc_docs_api.prototype['asc_setDefaultLanguage']                    = asc_docs_api.prototype.asc_setDefaultLanguage;
	asc_docs_api.prototype['asc_getDefaultLanguage']                    = asc_docs_api.prototype.asc_getDefaultLanguage;
	asc_docs_api.prototype['asc_getKeyboardLanguage']                   = asc_docs_api.prototype.asc_getKeyboardLanguage;
	asc_docs_api.prototype['asc_getInputLanguage']                      = asc_docs_api.prototype.asc_getInputLanguage;
	asc_docs_api.prototype['asc_setSpellCheckSettings']                 = asc_docs_api.prototype.asc_setSpellCheckSettings;
	asc_docs_api.prototype['asc_getSpellCheckSettings']                 = asc_docs_api.prototype.asc_getSpellCheckSettings;
	asc_docs_api.prototype['asc_setSpellCheck']                         = asc_docs_api.prototype.asc_setSpellCheck;
	asc_docs_api.prototype['asc_showComments']                          = asc_docs_api.prototype.asc_showComments;
	asc_docs_api.prototype['asc_hideComments']                          = asc_docs_api.prototype.asc_hideComments;
	asc_docs_api.prototype['asc_addComment']                            = asc_docs_api.prototype.asc_addComment;
	asc_docs_api.prototype['asc_removeComment']                         = asc_docs_api.prototype.asc_removeComment;
	asc_docs_api.prototype['asc_changeComment']                         = asc_docs_api.prototype.asc_changeComment;
	asc_docs_api.prototype['asc_selectComment']                         = asc_docs_api.prototype.asc_selectComment;
	asc_docs_api.prototype['asc_showComment']                           = asc_docs_api.prototype.asc_showComment;
	asc_docs_api.prototype['asc_GetCommentsReportByAuthors']            = asc_docs_api.prototype.asc_GetCommentsReportByAuthors;
	asc_docs_api.prototype['can_AddQuotedComment']                      = asc_docs_api.prototype.can_AddQuotedComment;
	asc_docs_api.prototype['sync_RemoveComment']                        = asc_docs_api.prototype.sync_RemoveComment;
	asc_docs_api.prototype['sync_AddComment']                           = asc_docs_api.prototype.sync_AddComment;
	asc_docs_api.prototype['sync_ShowComment']                          = asc_docs_api.prototype.sync_ShowComment;
	asc_docs_api.prototype['sync_HideComment']                          = asc_docs_api.prototype.sync_HideComment;
	asc_docs_api.prototype['sync_UpdateCommentPosition']                = asc_docs_api.prototype.sync_UpdateCommentPosition;
	asc_docs_api.prototype['sync_ChangeCommentData']                    = asc_docs_api.prototype.sync_ChangeCommentData;
	asc_docs_api.prototype['sync_LockComment']                          = asc_docs_api.prototype.sync_LockComment;
	asc_docs_api.prototype['sync_UnLockComment']                        = asc_docs_api.prototype.sync_UnLockComment;
	asc_docs_api.prototype['asc_RemoveAllComments']                     = asc_docs_api.prototype.asc_RemoveAllComments;
	asc_docs_api.prototype['asc_ResolveAllComments']                    = asc_docs_api.prototype.asc_ResolveAllComments;
	asc_docs_api.prototype['asc_GetCommentLogicPosition']               = asc_docs_api.prototype.asc_GetCommentLogicPosition;
	asc_docs_api.prototype['sync_LockHeaderFooters']                    = asc_docs_api.prototype.sync_LockHeaderFooters;
	asc_docs_api.prototype['sync_LockDocumentProps']                    = asc_docs_api.prototype.sync_LockDocumentProps;
	asc_docs_api.prototype['sync_UnLockHeaderFooters']                  = asc_docs_api.prototype.sync_UnLockHeaderFooters;
	asc_docs_api.prototype['sync_UnLockDocumentProps']                  = asc_docs_api.prototype.sync_UnLockDocumentProps;
	asc_docs_api.prototype['sync_CollaborativeChanges']                 = asc_docs_api.prototype.sync_CollaborativeChanges;
	asc_docs_api.prototype['sync_LockDocumentSchema']                   = asc_docs_api.prototype.sync_LockDocumentSchema;
	asc_docs_api.prototype['sync_UnLockDocumentSchema']                 = asc_docs_api.prototype.sync_UnLockDocumentSchema;
	asc_docs_api.prototype['zoomIn']                                    = asc_docs_api.prototype.zoomIn;
	asc_docs_api.prototype['zoomOut']                                   = asc_docs_api.prototype.zoomOut;
	asc_docs_api.prototype['zoomFitToPage']                             = asc_docs_api.prototype.zoomFitToPage;
	asc_docs_api.prototype['zoomFitToWidth']                            = asc_docs_api.prototype.zoomFitToWidth;
	asc_docs_api.prototype['zoomCustomMode']                            = asc_docs_api.prototype.zoomCustomMode;
	asc_docs_api.prototype['zoom100']                                   = asc_docs_api.prototype.zoom100;
	asc_docs_api.prototype['zoom']                                      = asc_docs_api.prototype.zoom;
	asc_docs_api.prototype['goToPage']                                  = asc_docs_api.prototype.goToPage;
	asc_docs_api.prototype['getCountPages']                             = asc_docs_api.prototype.getCountPages;
	asc_docs_api.prototype['getCurrentPage']                            = asc_docs_api.prototype.getCurrentPage;
	asc_docs_api.prototype['sync_countPagesCallback']                   = asc_docs_api.prototype.sync_countPagesCallback;
	asc_docs_api.prototype['sync_currentPageCallback']                  = asc_docs_api.prototype.sync_currentPageCallback;
	asc_docs_api.prototype['asc_enableKeyEvents']                       = asc_docs_api.prototype.asc_enableKeyEvents;
	asc_docs_api.prototype['GenerateStyles']                            = asc_docs_api.prototype.GenerateStyles;
	asc_docs_api.prototype['asyncFontsDocumentEndLoaded']               = asc_docs_api.prototype.asyncFontsDocumentEndLoaded;
	asc_docs_api.prototype['CreateFontsCharMap']                        = asc_docs_api.prototype.CreateFontsCharMap;
	asc_docs_api.prototype['sync_SendThemeColors']                      = asc_docs_api.prototype.sync_SendThemeColors;
	asc_docs_api.prototype['ChangeColorScheme']                         = asc_docs_api.prototype.ChangeColorScheme;
	asc_docs_api.prototype['asc_ChangeColorSchemeByIdx']                = asc_docs_api.prototype.asc_ChangeColorSchemeByIdx;
	asc_docs_api.prototype['UpdateInterfaceState']                      = asc_docs_api.prototype.UpdateInterfaceState;
	asc_docs_api.prototype['asyncFontEndLoaded']                        = asc_docs_api.prototype.asyncFontEndLoaded;
	asc_docs_api.prototype['asyncImageEndLoaded']                       = asc_docs_api.prototype.asyncImageEndLoaded;
	asc_docs_api.prototype['asyncImageEndLoadedBackground']             = asc_docs_api.prototype.asyncImageEndLoadedBackground;
	asc_docs_api.prototype['IsAsyncOpenDocumentImages']                 = asc_docs_api.prototype.IsAsyncOpenDocumentImages;
	asc_docs_api.prototype['pre_Paste']                                 = asc_docs_api.prototype.pre_Paste;
	asc_docs_api.prototype['pre_Save']                                  = asc_docs_api.prototype.pre_Save;
	asc_docs_api.prototype['SyncLoadImages']                            = asc_docs_api.prototype.SyncLoadImages;
	asc_docs_api.prototype['SyncLoadImages_callback']                   = asc_docs_api.prototype.SyncLoadImages_callback;
	asc_docs_api.prototype['ViewScrollToX']                             = asc_docs_api.prototype.ViewScrollToX;
	asc_docs_api.prototype['ViewScrollToY']                             = asc_docs_api.prototype.ViewScrollToY;
	asc_docs_api.prototype['GetDocWidthPx']                             = asc_docs_api.prototype.GetDocWidthPx;
	asc_docs_api.prototype['GetDocHeightPx']                            = asc_docs_api.prototype.GetDocHeightPx;
	asc_docs_api.prototype['ClearSearch']                               = asc_docs_api.prototype.ClearSearch;
	asc_docs_api.prototype['GetCurrentVisiblePage']                     = asc_docs_api.prototype.GetCurrentVisiblePage;
	asc_docs_api.prototype['GetCurrentVisiblePages']                    = asc_docs_api.prototype.GetCurrentVisiblePages;
	asc_docs_api.prototype['asc_setAutoSaveGap']                        = asc_docs_api.prototype.asc_setAutoSaveGap;
	asc_docs_api.prototype['asc_SetDocumentPlaceChangedEnabled']        = asc_docs_api.prototype.asc_SetDocumentPlaceChangedEnabled;
	asc_docs_api.prototype['asc_SetViewRulers']                         = asc_docs_api.prototype.asc_SetViewRulers;
	asc_docs_api.prototype['asc_SetViewRulersChange']                   = asc_docs_api.prototype.asc_SetViewRulersChange;
	asc_docs_api.prototype['asc_GetViewRulers']                         = asc_docs_api.prototype.asc_GetViewRulers;
	asc_docs_api.prototype['asc_SetDocumentUnits']                      = asc_docs_api.prototype.asc_SetDocumentUnits;
	asc_docs_api.prototype['GoToHeader']                                = asc_docs_api.prototype.GoToHeader;
	asc_docs_api.prototype['GoToFooter']                                = asc_docs_api.prototype.GoToFooter;
	asc_docs_api.prototype['ExitHeader_Footer']                         = asc_docs_api.prototype.ExitHeader_Footer;
	asc_docs_api.prototype['asc_RemoveHeader']                          = asc_docs_api.prototype.asc_RemoveHeader;
	asc_docs_api.prototype['asc_RemoveFooter']                          = asc_docs_api.prototype.asc_RemoveFooter;
	asc_docs_api.prototype['GetCurrentPixOffsetY']                      = asc_docs_api.prototype.GetCurrentPixOffsetY;
	asc_docs_api.prototype['SetPaintFormat']                            = asc_docs_api.prototype.SetPaintFormat;
	asc_docs_api.prototype['ChangeShapeType']                           = asc_docs_api.prototype.ChangeShapeType;
	asc_docs_api.prototype['sync_PaintFormatCallback']                  = asc_docs_api.prototype.sync_PaintFormatCallback;
	asc_docs_api.prototype['SetMarkerFormat']                           = asc_docs_api.prototype.SetMarkerFormat;
	asc_docs_api.prototype['sync_MarkerFormatCallback']                 = asc_docs_api.prototype.sync_MarkerFormatCallback;
	asc_docs_api.prototype['StartAddShape']                             = asc_docs_api.prototype.StartAddShape;
	asc_docs_api.prototype['AddShapeOnCurrentPage']                     = asc_docs_api.prototype.AddShapeOnCurrentPage;
	asc_docs_api.prototype['AddTextArt']                                = asc_docs_api.prototype.AddTextArt;
	asc_docs_api.prototype['asc_canEditCrop']                           = asc_docs_api.prototype.asc_canEditCrop;
	asc_docs_api.prototype['asc_startEditCrop']                         = asc_docs_api.prototype.asc_startEditCrop;
	asc_docs_api.prototype['asc_endEditCrop']                           = asc_docs_api.prototype.asc_endEditCrop;
	asc_docs_api.prototype['asc_cropFit']                               = asc_docs_api.prototype.asc_cropFit;
	asc_docs_api.prototype['asc_cropFill']                              = asc_docs_api.prototype.asc_cropFill;
	asc_docs_api.prototype["asc_GetWatermarkProps"]                     = asc_docs_api.prototype.asc_GetWatermarkProps;
	asc_docs_api.prototype["asc_SetWatermarkProps"]                     = asc_docs_api.prototype.asc_SetWatermarkProps;
	asc_docs_api.prototype["asc_WatermarkRemove"]                       = asc_docs_api.prototype.asc_WatermarkRemove;
	asc_docs_api.prototype['asc_canMergeSelectedShapes']                = asc_docs_api.prototype.asc_canMergeSelectedShapes;
	asc_docs_api.prototype['sync_StartAddShapeCallback']                = asc_docs_api.prototype.sync_StartAddShapeCallback;
	asc_docs_api.prototype['CanGroup']                                  = asc_docs_api.prototype.CanGroup;
	asc_docs_api.prototype['CanUnGroup']                                = asc_docs_api.prototype.CanUnGroup;
	asc_docs_api.prototype['CanChangeWrapPolygon']                      = asc_docs_api.prototype.CanChangeWrapPolygon;
	asc_docs_api.prototype['StartChangeWrapPolygon']                    = asc_docs_api.prototype.StartChangeWrapPolygon;
	asc_docs_api.prototype['ClearFormating']                            = asc_docs_api.prototype.ClearFormating;
	asc_docs_api.prototype['GetSectionInfo']                            = asc_docs_api.prototype.GetSectionInfo;
	asc_docs_api.prototype['add_SectionBreak']                          = asc_docs_api.prototype.add_SectionBreak;
	asc_docs_api.prototype['asc_GetSectionsCount']                      = asc_docs_api.prototype.asc_GetSectionsCount;
	asc_docs_api.prototype['asc_setViewMode']                           = asc_docs_api.prototype.asc_setViewMode;
	asc_docs_api.prototype['OnMouseUp']                                 = asc_docs_api.prototype.OnMouseUp;
	asc_docs_api.prototype['asyncImageEndLoaded2']                      = asc_docs_api.prototype.asyncImageEndLoaded2;
	asc_docs_api.prototype['SetDrawImagePlaceParagraph']                = asc_docs_api.prototype.SetDrawImagePlaceParagraph;
	asc_docs_api.prototype['asc_getMasterCommentId']                    = asc_docs_api.prototype.asc_getMasterCommentId;
	asc_docs_api.prototype['asc_getAnchorPosition']                     = asc_docs_api.prototype.asc_getAnchorPosition;
	asc_docs_api.prototype['asc_createSmartArt']                        = asc_docs_api.prototype.asc_createSmartArt;
	asc_docs_api.prototype['asc_getChartObject']                        = asc_docs_api.prototype.asc_getChartObject;
	asc_docs_api.prototype['asc_addChartDrawingObject']                 = asc_docs_api.prototype.asc_addChartDrawingObject;
	asc_docs_api.prototype['asc_doubleClickOnChart']                    = asc_docs_api.prototype.asc_doubleClickOnChart;
	asc_docs_api.prototype['asc_onCloseChartFrame']                     = asc_docs_api.prototype.asc_onCloseChartFrame;
	asc_docs_api.prototype['asc_editChartDrawingObject']                = asc_docs_api.prototype.asc_editChartDrawingObject;
	asc_docs_api.prototype['asc_getChartPreviews']                      = asc_docs_api.prototype.asc_getChartPreviews;
	asc_docs_api.prototype['asc_getTextArtPreviews']                    = asc_docs_api.prototype.asc_getTextArtPreviews;
	asc_docs_api.prototype['sync_closeOleEditor']                       = asc_docs_api.prototype.sync_closeOleEditor;
	asc_docs_api.prototype['asc_setDrawCollaborationMarks']             = asc_docs_api.prototype.asc_setDrawCollaborationMarks;
	asc_docs_api.prototype['asc_AddMath']                               = asc_docs_api.prototype.asc_AddMath;
	asc_docs_api.prototype['asc_AddMath2']                              = asc_docs_api.prototype.asc_AddMath2;
	asc_docs_api.prototype['asc_ConvertMathView']                       = asc_docs_api.prototype.asc_ConvertMathView;
	asc_docs_api.prototype['asc_AddPageCount']                          = asc_docs_api.prototype.asc_AddPageCount;
	asc_docs_api.prototype['asc_StartMailMerge']                        = asc_docs_api.prototype.asc_StartMailMerge;
	asc_docs_api.prototype['asc_StartMailMergeByList']                  = asc_docs_api.prototype.asc_StartMailMergeByList;
	asc_docs_api.prototype['asc_GetReceptionsCount']                    = asc_docs_api.prototype.asc_GetReceptionsCount;
	asc_docs_api.prototype['asc_GetMailMergeFieldsNameList']            = asc_docs_api.prototype.asc_GetMailMergeFieldsNameList;
	asc_docs_api.prototype['asc_AddMailMergeField']                     = asc_docs_api.prototype.asc_AddMailMergeField;
	asc_docs_api.prototype['asc_SetHighlightMailMergeFields']           = asc_docs_api.prototype.asc_SetHighlightMailMergeFields;
	asc_docs_api.prototype['asc_PreviewMailMergeResult']                = asc_docs_api.prototype.asc_PreviewMailMergeResult;
	asc_docs_api.prototype['asc_EndPreviewMailMergeResult']             = asc_docs_api.prototype.asc_EndPreviewMailMergeResult;
	asc_docs_api.prototype['sync_StartMailMerge']                       = asc_docs_api.prototype.sync_StartMailMerge;
	asc_docs_api.prototype['sync_PreviewMailMergeResult']               = asc_docs_api.prototype.sync_PreviewMailMergeResult;
	asc_docs_api.prototype['sync_EndPreviewMailMergeResult']            = asc_docs_api.prototype.sync_EndPreviewMailMergeResult;
	asc_docs_api.prototype['sync_HighlightMailMergeFields']             = asc_docs_api.prototype.sync_HighlightMailMergeFields;
	asc_docs_api.prototype['asc_getMailMergeData']                      = asc_docs_api.prototype.asc_getMailMergeData;
	asc_docs_api.prototype['asc_setMailMergeData']                      = asc_docs_api.prototype.asc_setMailMergeData;
	asc_docs_api.prototype['asc_sendMailMergeData']                     = asc_docs_api.prototype.asc_sendMailMergeData;
	asc_docs_api.prototype['asc_GetMailMergeFiledValue']                = asc_docs_api.prototype.asc_GetMailMergeFiledValue;
	asc_docs_api.prototype['asc_GetStyleFromFormatting']                = asc_docs_api.prototype.asc_GetStyleFromFormatting;
	asc_docs_api.prototype['asc_AddNewStyle']                           = asc_docs_api.prototype.asc_AddNewStyle;
	asc_docs_api.prototype['asc_RemoveStyle']                           = asc_docs_api.prototype.asc_RemoveStyle;
	asc_docs_api.prototype['asc_RemoveAllCustomStyles']                 = asc_docs_api.prototype.asc_RemoveAllCustomStyles;
	asc_docs_api.prototype['asc_IsStyleDefault']                        = asc_docs_api.prototype.asc_IsStyleDefault;
	asc_docs_api.prototype['asc_IsDefaultStyleChanged']                 = asc_docs_api.prototype.asc_IsDefaultStyleChanged;
	asc_docs_api.prototype['asc_GetStyleNameById']                      = asc_docs_api.prototype.asc_GetStyleNameById;
	asc_docs_api.prototype['asc_SetTrackRevisions']                     = asc_docs_api.prototype.asc_SetTrackRevisions;
	asc_docs_api.prototype['asc_IsTrackRevisions']                      = asc_docs_api.prototype.asc_IsTrackRevisions;
	asc_docs_api.prototype['asc_SetLocalTrackRevisions']                = asc_docs_api.prototype.asc_SetLocalTrackRevisions;
	asc_docs_api.prototype['asc_GetLocalTrackRevisions']                = asc_docs_api.prototype.asc_GetLocalTrackRevisions;
	asc_docs_api.prototype['asc_SetGlobalTrackRevisions']               = asc_docs_api.prototype.asc_SetGlobalTrackRevisions;
	asc_docs_api.prototype['asc_GetGlobalTrackRevisions']               = asc_docs_api.prototype.asc_GetGlobalTrackRevisions;
	asc_docs_api.prototype['sync_BeginCatchRevisionsChanges']           = asc_docs_api.prototype.sync_BeginCatchRevisionsChanges;
	asc_docs_api.prototype['sync_EndCatchRevisionsChanges']             = asc_docs_api.prototype.sync_EndCatchRevisionsChanges;
	asc_docs_api.prototype['asc_GetRevisionsChangesStack']              = asc_docs_api.prototype.asc_GetRevisionsChangesStack;
	asc_docs_api.prototype['sync_AddRevisionsChange']                   = asc_docs_api.prototype.sync_AddRevisionsChange;
	asc_docs_api.prototype['asc_AcceptChanges']                         = asc_docs_api.prototype.asc_AcceptChanges;
	asc_docs_api.prototype['asc_RejectChanges']                         = asc_docs_api.prototype.asc_RejectChanges;
	asc_docs_api.prototype['asc_AcceptChangesBySelection']              = asc_docs_api.prototype.asc_AcceptChangesBySelection;
	asc_docs_api.prototype['asc_RejectChangesBySelection']              = asc_docs_api.prototype.asc_RejectChangesBySelection;
	asc_docs_api.prototype['asc_HaveRevisionsChanges']                  = asc_docs_api.prototype.asc_HaveRevisionsChanges;
	asc_docs_api.prototype['asc_HaveNewRevisionsChanges']               = asc_docs_api.prototype.asc_HaveNewRevisionsChanges;
	asc_docs_api.prototype['asc_GetNextRevisionsChange']                = asc_docs_api.prototype.asc_GetNextRevisionsChange;
	asc_docs_api.prototype['asc_GetPrevRevisionsChange']                = asc_docs_api.prototype.asc_GetPrevRevisionsChange;
	asc_docs_api.prototype['sync_UpdateRevisionsChangesPosition']       = asc_docs_api.prototype.sync_UpdateRevisionsChangesPosition;
	asc_docs_api.prototype['asc_AcceptAllChanges']                      = asc_docs_api.prototype.asc_AcceptAllChanges;
	asc_docs_api.prototype['asc_RejectAllChanges']                      = asc_docs_api.prototype.asc_RejectAllChanges;
	asc_docs_api.prototype['asc_GetTrackRevisionsReportByAuthors']      = asc_docs_api.prototype.asc_GetTrackRevisionsReportByAuthors;
	asc_docs_api.prototype['asc_FollowRevisionMove']                    = asc_docs_api.prototype.asc_FollowRevisionMove;
	asc_docs_api.prototype['asc_stopSaving']                            = asc_docs_api.prototype.asc_stopSaving;
	asc_docs_api.prototype['asc_continueSaving']                        = asc_docs_api.prototype.asc_continueSaving;
	asc_docs_api.prototype['asc_undoAllChanges']                        = asc_docs_api.prototype.asc_undoAllChanges;
	asc_docs_api.prototype['asc_CloseFile']                             = asc_docs_api.prototype.asc_CloseFile;
	asc_docs_api.prototype['asc_SetFastCollaborative']                  = asc_docs_api.prototype.asc_SetFastCollaborative;
	asc_docs_api.prototype['asc_isOffline']                             = asc_docs_api.prototype.asc_isOffline;
	asc_docs_api.prototype['asc_getUrlType']                            = asc_docs_api.prototype.asc_getUrlType;
	asc_docs_api.prototype['asc_prepareUrl']                            = asc_docs_api.prototype.asc_prepareUrl;
	asc_docs_api.prototype['asc_getSessionToken']                 		= asc_docs_api.prototype.asc_getSessionToken;
	asc_docs_api.prototype["asc_setInterfaceDrawImagePlaceShape"]       = asc_docs_api.prototype.asc_setInterfaceDrawImagePlaceShape;
	asc_docs_api.prototype["asc_pluginsRegister"]                       = asc_docs_api.prototype.asc_pluginsRegister;
	asc_docs_api.prototype["asc_pluginRun"]                             = asc_docs_api.prototype.asc_pluginRun;
	asc_docs_api.prototype["asc_pluginStop"]                            = asc_docs_api.prototype.asc_pluginStop;
	asc_docs_api.prototype["asc_pluginResize"]                          = asc_docs_api.prototype.asc_pluginResize;
	asc_docs_api.prototype["asc_pluginButtonClick"]                     = asc_docs_api.prototype.asc_pluginButtonClick;
	asc_docs_api.prototype["asc_pluginEnableMouseEvents"]         		= asc_docs_api.prototype.asc_pluginEnableMouseEvents;

	asc_docs_api.prototype["asc_nativeInitBuilder"]                     = asc_docs_api.prototype.asc_nativeInitBuilder;
	asc_docs_api.prototype["asc_SetSilentMode"]                         = asc_docs_api.prototype.asc_SetSilentMode;
	asc_docs_api.prototype["asc_startEditCurrentOleObject"]             = asc_docs_api.prototype.asc_startEditCurrentOleObject;
	asc_docs_api.prototype["asc_doubleClickOnTableOleObject"]           = asc_docs_api.prototype.asc_doubleClickOnTableOleObject;
	asc_docs_api.prototype["asc_editOleObjectAction"]                   = asc_docs_api.prototype.asc_editOleObjectAction;
	asc_docs_api.prototype["asc_addOleObjectAction"]                    = asc_docs_api.prototype.asc_addOleObjectAction;
	asc_docs_api.prototype["asc_InputClearKeyboardElement"]             = asc_docs_api.prototype.asc_InputClearKeyboardElement;
	asc_docs_api.prototype["asc_SpecialPaste"]                          = asc_docs_api.prototype.asc_SpecialPaste;

	asc_docs_api.prototype["SetDrawImagePlaceContents"]					= asc_docs_api.prototype.SetDrawImagePlaceContents;
	asc_docs_api.prototype["SetDrawImagePlaceTableOfFigures"]			= asc_docs_api.prototype.SetDrawImagePlaceTableOfFigures;
    asc_docs_api.prototype["SetDrawImagePreviewMargins"]				= asc_docs_api.prototype.SetDrawImagePreviewMargins;
    asc_docs_api.prototype["SetDrawImagePreviewBullet"]					= asc_docs_api.prototype.SetDrawImagePreviewBullet;
    asc_docs_api.prototype["ConvertTextToTable"]						= asc_docs_api.prototype.ConvertTextToTable;
    asc_docs_api.prototype["ConvertTableToText"]						= asc_docs_api.prototype.ConvertTableToText;
    asc_docs_api.prototype["SetDrawImagePreviewBulletChangeListLevel"]	= asc_docs_api.prototype.SetDrawImagePreviewBulletChangeListLevel;
    asc_docs_api.prototype["SetDrawImagePreviewBulletForMenu"]			= asc_docs_api.prototype.SetDrawImagePreviewBulletForMenu;
    asc_docs_api.prototype["asc_GetPossibleNumberingLanguage"]			= asc_docs_api.prototype.asc_GetPossibleNumberingLanguage;

	asc_docs_api.prototype["asc_AddContentControl"]                     = asc_docs_api.prototype.asc_AddContentControl;
	asc_docs_api.prototype["asc_AddContentControlCheckBox"]             = asc_docs_api.prototype.asc_AddContentControlCheckBox;
	asc_docs_api.prototype["asc_AddContentControlPicture"]              = asc_docs_api.prototype.asc_AddContentControlPicture;
	asc_docs_api.prototype["asc_AddContentControlList"]                 = asc_docs_api.prototype.asc_AddContentControlList;
	asc_docs_api.prototype["asc_AddContentControlDatePicker"]           = asc_docs_api.prototype.asc_AddContentControlDatePicker;
	asc_docs_api.prototype["asc_RemoveContentControl"]                  = asc_docs_api.prototype.asc_RemoveContentControl;
	asc_docs_api.prototype["asc_RemoveContentControlWrapper"]           = asc_docs_api.prototype.asc_RemoveContentControlWrapper;
	asc_docs_api.prototype["asc_SetContentControlProperties"]           = asc_docs_api.prototype.asc_SetContentControlProperties;
	asc_docs_api.prototype["asc_IsContentControl"]                      = asc_docs_api.prototype.asc_IsContentControl;
	asc_docs_api.prototype["asc_GetContentControlProperties"]           = asc_docs_api.prototype.asc_GetContentControlProperties;
	asc_docs_api.prototype["asc_GetCurrentContentControl"]              = asc_docs_api.prototype.asc_GetCurrentContentControl;
	asc_docs_api.prototype["asc_ClearContentControl"]                   = asc_docs_api.prototype.asc_ClearContentControl;
	asc_docs_api.prototype["asc_GetContentControlRightAnchorPosition"]  = asc_docs_api.prototype.asc_GetContentControlRightAnchorPosition;
	asc_docs_api.prototype["asc_GetContentControlBoundingRect"]         = asc_docs_api.prototype.asc_GetContentControlBoundingRect;
	asc_docs_api.prototype["asc_UncheckContentControlButtons"]          = asc_docs_api.prototype.asc_UncheckContentControlButtons;
	asc_docs_api.prototype['asc_SetGlobalContentControlHighlightColor'] = asc_docs_api.prototype.asc_SetGlobalContentControlHighlightColor;
	asc_docs_api.prototype['asc_GetGlobalContentControlHighlightColor'] = asc_docs_api.prototype.asc_GetGlobalContentControlHighlightColor;
	asc_docs_api.prototype['asc_SetGlobalContentControlShowHighlight']  = asc_docs_api.prototype.asc_SetGlobalContentControlShowHighlight;
	asc_docs_api.prototype['asc_GetGlobalContentControlShowHighlight']  = asc_docs_api.prototype.asc_GetGlobalContentControlShowHighlight;
	asc_docs_api.prototype['asc_SetContentControlCheckBoxPr']           = asc_docs_api.prototype.asc_SetContentControlCheckBoxPr;
	asc_docs_api.prototype['asc_SetContentControlCheckBoxChecked']      = asc_docs_api.prototype.asc_SetContentControlCheckBoxChecked;
	asc_docs_api.prototype['asc_IsContentControlCheckBoxChecked']       = asc_docs_api.prototype.asc_IsContentControlCheckBoxChecked;
	asc_docs_api.prototype['asc_SetContentControlPictureUrl']           = asc_docs_api.prototype.asc_SetContentControlPictureUrl;
	asc_docs_api.prototype['asc_SetContentControlListPr']               = asc_docs_api.prototype.asc_SetContentControlListPr;
	asc_docs_api.prototype['asc_SelectContentControlListItem']          = asc_docs_api.prototype.asc_SelectContentControlListItem;
	asc_docs_api.prototype['asc_GetContentControlListCurrentValue']     = asc_docs_api.prototype.asc_GetContentControlListCurrentValue;
	asc_docs_api.prototype['asc_SetContentControlDatePickerPr']         = asc_docs_api.prototype.asc_SetContentControlDatePickerPr;
	asc_docs_api.prototype['asc_SetContentControlDatePickerDate']       = asc_docs_api.prototype.asc_SetContentControlDatePickerDate;
	asc_docs_api.prototype['asc_SetContentControlTextPlaceholder']      = asc_docs_api.prototype.asc_SetContentControlTextPlaceholder;
	asc_docs_api.prototype['asc_SetContentControlText']                 = asc_docs_api.prototype.asc_SetContentControlText;
	asc_docs_api.prototype['asc_GetTextFormKeys']                       = asc_docs_api.prototype.asc_GetTextFormKeys;
	asc_docs_api.prototype['asc_GetPictureFormKeys']                    = asc_docs_api.prototype.asc_GetPictureFormKeys;
	asc_docs_api.prototype['asc_GetCheckBoxFormKeys']                   = asc_docs_api.prototype.asc_GetCheckBoxFormKeys;
	asc_docs_api.prototype['asc_GetRadioButtonGroupKeys']               = asc_docs_api.prototype.asc_GetRadioButtonGroupKeys;
	asc_docs_api.prototype['asc_GetFormKeysByType']                     = asc_docs_api.prototype.asc_GetFormKeysByType;
	asc_docs_api.prototype['asc_ClearAllSpecialForms']                  = asc_docs_api.prototype.asc_ClearAllSpecialForms;
	asc_docs_api.prototype['asc_SetSpecialFormsHighlightColor']         = asc_docs_api.prototype.asc_SetSpecialFormsHighlightColor;
	asc_docs_api.prototype['asc_GetSpecialFormsHighlightColor']         = asc_docs_api.prototype.asc_GetSpecialFormsHighlightColor;
	asc_docs_api.prototype['asc_SetPerformContentControlActionByClick'] = asc_docs_api.prototype.asc_SetPerformContentControlActionByClick;
	asc_docs_api.prototype['asc_GetTextFormAutoWidth']                  = asc_docs_api.prototype.asc_GetTextFormAutoWidth;
	asc_docs_api.prototype['asc_GetFormsCountByKey']                    = asc_docs_api.prototype.asc_GetFormsCountByKey;
	asc_docs_api.prototype['asc_MoveToFillingForm']                     = asc_docs_api.prototype.asc_MoveToFillingForm;
	asc_docs_api.prototype['asc_IsAllRequiredFormsFilled']              = asc_docs_api.prototype.asc_IsAllRequiredFormsFilled;
	asc_docs_api.prototype['asc_SendForm']                    			= asc_docs_api.prototype.asc_SendForm;
	asc_docs_api.prototype['asc_SetFixedForm']                          = asc_docs_api.prototype.asc_SetFixedForm;
	asc_docs_api.prototype['asc_IsHighlightRequiredFields']             = asc_docs_api.prototype.asc_IsHighlightRequiredFields;
	asc_docs_api.prototype['asc_SetHighlightRequiredFields']            = asc_docs_api.prototype.asc_SetHighlightRequiredFields;
	asc_docs_api.prototype['asc_GetAllFormsData']                       = asc_docs_api.prototype.asc_GetAllFormsData;
	asc_docs_api.prototype['asc_GetOForm']                              = asc_docs_api.prototype.asc_GetOForm;
	asc_docs_api.prototype['asc_DisconnectEveryone']                    = asc_docs_api.prototype.asc_DisconnectEveryone;
	asc_docs_api.prototype['asc_CompletePreparingOForm']                = asc_docs_api.prototype.asc_CompletePreparingOForm;

	asc_docs_api.prototype['asc_BeginViewModeInReview']                 = asc_docs_api.prototype.asc_BeginViewModeInReview;
	asc_docs_api.prototype['asc_EndViewModeInReview']                   = asc_docs_api.prototype.asc_EndViewModeInReview;
	asc_docs_api.prototype['asc_SetDisplayModeInReview']                = asc_docs_api.prototype.asc_SetDisplayModeInReview;
	asc_docs_api.prototype['asc_GetDisplayModeInReview']                = asc_docs_api.prototype.asc_GetDisplayModeInReview;

	asc_docs_api.prototype['asc_GetDocumentOutlineManager']             = asc_docs_api.prototype.asc_GetDocumentOutlineManager;
	asc_docs_api.prototype['asc_ShowDocumentOutline']                   = asc_docs_api.prototype.asc_ShowDocumentOutline;
	asc_docs_api.prototype['asc_HideDocumentOutline']                   = asc_docs_api.prototype.asc_HideDocumentOutline;
	asc_docs_api.prototype['sync_OnDocumentOutlineUpdate']              = asc_docs_api.prototype.sync_OnDocumentOutlineUpdate;
	asc_docs_api.prototype['sync_OnDocumentOutlineCurrentPosition']     = asc_docs_api.prototype.sync_OnDocumentOutlineCurrentPosition;

	asc_docs_api.prototype['asc_AddTableOfContents']                    = asc_docs_api.prototype.asc_AddTableOfContents;
	asc_docs_api.prototype['asc_RemoveTableOfContents']                 = asc_docs_api.prototype.asc_RemoveTableOfContents;
	asc_docs_api.prototype['asc_GetTableOfContentsPr']                  = asc_docs_api.prototype.asc_GetTableOfContentsPr;
	asc_docs_api.prototype['asc_SetTableOfContentsPr']                  = asc_docs_api.prototype.asc_SetTableOfContentsPr;
	asc_docs_api.prototype['asc_UpdateTableOfContents']                 = asc_docs_api.prototype.asc_UpdateTableOfContents;
	asc_docs_api.prototype['asc_AddParagraphToTOC']                     = asc_docs_api.prototype.asc_AddParagraphToTOC;
	asc_docs_api.prototype['asc_GetCurrentLevelTOC']                    = asc_docs_api.prototype.asc_GetCurrentLevelTOC;

	asc_docs_api.prototype['asc_getAllUsedParagraphStyles']             = asc_docs_api.prototype.asc_getAllUsedParagraphStyles;
	asc_docs_api.prototype['asc_CanUpdateTablesOfFigures']              = asc_docs_api.prototype.asc_CanUpdateTablesOfFigures;
	asc_docs_api.prototype['asc_UpdateTablesOfFigures']                 = asc_docs_api.prototype.asc_UpdateTablesOfFigures;
	asc_docs_api.prototype['asc_AddTableOfFigures']                     = asc_docs_api.prototype.asc_AddTableOfFigures;
	asc_docs_api.prototype['asc_GetTableOfFiguresPr']                   = asc_docs_api.prototype.asc_GetTableOfFiguresPr;
	
	asc_docs_api.prototype['asc_AddComplexFieldWithInstruction']        = asc_docs_api.prototype.asc_AddComplexFieldWithInstruction;
	asc_docs_api.prototype['asc_GetComplexFieldInstruction']            = asc_docs_api.prototype.asc_GetComplexFieldInstruction;
	asc_docs_api.prototype['asc_EditComplexFieldInstruction']           = asc_docs_api.prototype.asc_EditComplexFieldInstruction;
	asc_docs_api.prototype['asc_GetCurrentComplexField']                = asc_docs_api.prototype.asc_GetCurrentComplexField;
	asc_docs_api.prototype['asc_UpdateComplexField']                    = asc_docs_api.prototype.asc_UpdateComplexField;
	asc_docs_api.prototype['asc_RemoveComplexField']                    = asc_docs_api.prototype.asc_RemoveComplexField;
	asc_docs_api.prototype['asc_SetComplexFieldPr']                     = asc_docs_api.prototype.asc_SetComplexFieldPr;
	asc_docs_api.prototype['asc_AddTableFormula']                       = asc_docs_api.prototype.asc_AddTableFormula;
	asc_docs_api.prototype['asc_GetTableFormula']                       = asc_docs_api.prototype.asc_GetTableFormula;
	asc_docs_api.prototype['asc_GetTableFormulaFormats']                = asc_docs_api.prototype.asc_GetTableFormulaFormats;
	asc_docs_api.prototype['asc_ParseTableFormulaInstrLine']            = asc_docs_api.prototype.asc_ParseTableFormulaInstrLine;
	asc_docs_api.prototype['asc_CreateInstructionLine']                 = asc_docs_api.prototype.asc_CreateInstructionLine;
	asc_docs_api.prototype['asc_HaveFields']                            = asc_docs_api.prototype.asc_HaveFields;
	asc_docs_api.prototype['asc_UpdateFields']                          = asc_docs_api.prototype.asc_UpdateFields;
	asc_docs_api.prototype['asc_ToggleComplexFieldCodes']               = asc_docs_api.prototype.asc_ToggleComplexFieldCodes;

	asc_docs_api.prototype["asc_addDateTime"]                           = asc_docs_api.prototype.asc_addDateTime;

	asc_docs_api.prototype['asc_AddObjectCaption']                      = asc_docs_api.prototype.asc_AddObjectCaption;

	asc_docs_api.prototype['asc_GetBookmarksManager']                   = asc_docs_api.prototype.asc_GetBookmarksManager;
	asc_docs_api.prototype['asc_OnBookmarksUpdate']                     = asc_docs_api.prototype.asc_OnBookmarksUpdate;

	asc_docs_api.prototype['asc_GetHeadingLevel']                       = asc_docs_api.prototype.asc_GetHeadingLevel;
	asc_docs_api.prototype['asc_GetStylesArray']                        = asc_docs_api.prototype.asc_GetStylesArray;

	asc_docs_api.prototype['asc_SetAutomaticBulletedLists']             = asc_docs_api.prototype.asc_SetAutomaticBulletedLists;
	asc_docs_api.prototype['asc_SetAutomaticNumberedLists']             = asc_docs_api.prototype.asc_SetAutomaticNumberedLists;
	asc_docs_api.prototype['asc_SetAutoCorrectSmartQuotes']             = asc_docs_api.prototype.asc_SetAutoCorrectSmartQuotes;
	asc_docs_api.prototype['asc_SetAutoCorrectHyphensWithDash']         = asc_docs_api.prototype.asc_SetAutoCorrectHyphensWithDash;
	asc_docs_api.prototype['asc_SetAutoCorrectFirstLetterOfSentences']  = asc_docs_api.prototype.asc_SetAutoCorrectFirstLetterOfSentences;
	asc_docs_api.prototype['asc_SetAutoCorrectHyperlinks']              = asc_docs_api.prototype.asc_SetAutoCorrectHyperlinks;
	asc_docs_api.prototype['asc_SetAutoCorrectFirstLetterOfCells']      = asc_docs_api.prototype.asc_SetAutoCorrectFirstLetterOfCells;
	asc_docs_api.prototype['asc_SetAutoCorrectDoubleSpaceWithPeriod']   = asc_docs_api.prototype.asc_SetAutoCorrectDoubleSpaceWithPeriod;

	asc_docs_api.prototype['asc_SetFirstLetterAutoCorrectExceptions']   = asc_docs_api.prototype.asc_SetFirstLetterAutoCorrectExceptions;
	asc_docs_api.prototype['asc_GetFirstLetterAutoCorrectExceptions']   = asc_docs_api.prototype.asc_GetFirstLetterAutoCorrectExceptions;
	asc_docs_api.prototype['asc_GetAutoCorrectSettings']                = asc_docs_api.prototype.asc_GetAutoCorrectSettings;

	asc_docs_api.prototype['asc_RemoveSelection']                       = asc_docs_api.prototype.asc_RemoveSelection;
	asc_docs_api.prototype['asc_AddText']                               = asc_docs_api.prototype.asc_AddText;
	asc_docs_api.prototype['asc_GetCurrentWord']                        = asc_docs_api.prototype.asc_GetCurrentWord;
	asc_docs_api.prototype['asc_ReplaceCurrentWord']                    = asc_docs_api.prototype.asc_ReplaceCurrentWord;
	asc_docs_api.prototype['asc_GetSelectedText']                       = asc_docs_api.prototype.asc_GetSelectedText;
	asc_docs_api.prototype['asc_AddBlankPage']                          = asc_docs_api.prototype.asc_AddBlankPage;
	asc_docs_api.prototype['asc_EditSelectAll']                         = asc_docs_api.prototype.asc_EditSelectAll;
    asc_docs_api.prototype['sendEvent']         						= asc_docs_api.prototype.sendEvent;
	asc_docs_api.prototype['asc_setSkin']								= asc_docs_api.prototype.asc_setSkin;

	asc_docs_api.prototype['asc_ConvertEquationToMath']                 = asc_docs_api.prototype.asc_ConvertEquationToMath;
	asc_docs_api.prototype['asc_ConvertMathDisplayMode']                = asc_docs_api.prototype.asc_ConvertMathDisplayMode;
	asc_docs_api.prototype['asc_IsInlineMath']                          = asc_docs_api.prototype.asc_IsInlineMath;


	//cross-references
	asc_docs_api.prototype['asc_GetAllNumberedParagraphs']              = asc_docs_api.prototype.asc_GetAllNumberedParagraphs;
	asc_docs_api.prototype['asc_AddCrossRefToParagraph']                = asc_docs_api.prototype.asc_AddCrossRefToParagraph;
	asc_docs_api.prototype['asc_GetAllHeadingParagraphs']               = asc_docs_api.prototype.asc_GetAllHeadingParagraphs;
	asc_docs_api.prototype['asc_AddCrossRefToBookmark']                 = asc_docs_api.prototype.asc_AddCrossRefToBookmark;
	asc_docs_api.prototype['asc_GetAllFootNoteParagraphs']              = asc_docs_api.prototype.asc_GetAllFootNoteParagraphs;
	asc_docs_api.prototype['asc_AddCrossRefToNote']                     = asc_docs_api.prototype.asc_AddCrossRefToNote;
	asc_docs_api.prototype['asc_GetAllEndNoteParagraphs']               = asc_docs_api.prototype.asc_GetAllEndNoteParagraphs;
	asc_docs_api.prototype['asc_GetAllCaptionParagraphs']               = asc_docs_api.prototype.asc_GetAllCaptionParagraphs;
	asc_docs_api.prototype['asc_AddCrossRefToCaption']                  = asc_docs_api.prototype.asc_AddCrossRefToCaption;

	asc_docs_api.prototype['asc_ChangeTextCase']                        = asc_docs_api.prototype.asc_ChangeTextCase;
	asc_docs_api.prototype['asc_putSmartParagraphSelection']            = asc_docs_api.prototype.asc_SetSmartParagraphSelection;
	asc_docs_api.prototype['asc_isSmartParagraphSelection']             = asc_docs_api.prototype.asc_IsSmartParagraphSelection;

    asc_docs_api.prototype['SetTableDrawMode']         					= asc_docs_api.prototype.SetTableDrawMode;
    asc_docs_api.prototype['SetTableEraseMode']         				= asc_docs_api.prototype.SetTableEraseMode;

	asc_docs_api.prototype['asc_PreConvertTextToTable']                 = asc_docs_api.prototype.asc_PreConvertTextToTable;
	asc_docs_api.prototype['asc_ConvertTextToTable']                    = asc_docs_api.prototype.asc_ConvertTextToTable;
	asc_docs_api.prototype['asc_ConvertTableToText']                    = asc_docs_api.prototype.asc_ConvertTableToText;

	asc_docs_api.prototype['asc_GetSelectionBounds']                    = asc_docs_api.prototype.asc_GetSelectionBounds;


	// mobile
	asc_docs_api.prototype["asc_GetDefaultTableStyles"]             	= asc_docs_api.prototype.asc_GetDefaultTableStyles;
	asc_docs_api.prototype["asc_Remove"]             					= asc_docs_api.prototype.asc_Remove;
	asc_docs_api.prototype["asc_OnHideContextMenu"] 					= asc_docs_api.prototype.asc_OnHideContextMenu;
	asc_docs_api.prototype["asc_OnShowContextMenu"] 					= asc_docs_api.prototype.asc_OnShowContextMenu;

	// signatures
	asc_docs_api.prototype["asc_addSignatureLine"] 						= asc_docs_api.prototype.asc_addSignatureLine;
	asc_docs_api.prototype["asc_CallSignatureDblClickEvent"]			= asc_docs_api.prototype.asc_CallSignatureDblClickEvent;
	asc_docs_api.prototype["asc_getRequestSignatures"] 					= asc_docs_api.prototype.asc_getRequestSignatures;
	asc_docs_api.prototype["asc_AddSignatureLine2"]             		= asc_docs_api.prototype.asc_AddSignatureLine2;
	asc_docs_api.prototype["asc_Sign"]             						= asc_docs_api.prototype.asc_Sign;
	asc_docs_api.prototype["asc_RequestSign"]             				= asc_docs_api.prototype.asc_RequestSign;
	asc_docs_api.prototype["asc_ViewCertificate"] 						= asc_docs_api.prototype.asc_ViewCertificate;
	asc_docs_api.prototype["asc_SelectCertificate"] 					= asc_docs_api.prototype.asc_SelectCertificate;
	asc_docs_api.prototype["asc_GetDefaultCertificate"] 				= asc_docs_api.prototype.asc_GetDefaultCertificate;
	asc_docs_api.prototype["asc_getSignatures"] 						= asc_docs_api.prototype.asc_getSignatures;
	asc_docs_api.prototype["asc_isSignaturesSupport"] 					= asc_docs_api.prototype.asc_isSignaturesSupport;
    asc_docs_api.prototype["asc_isProtectionSupport"] 					= asc_docs_api.prototype.asc_isProtectionSupport;
	asc_docs_api.prototype["asc_isAnonymousSupport"] 					= asc_docs_api.prototype.asc_isAnonymousSupport;
	asc_docs_api.prototype["asc_RemoveSignature"] 						= asc_docs_api.prototype.asc_RemoveSignature;
	asc_docs_api.prototype["asc_RemoveAllSignatures"] 					= asc_docs_api.prototype.asc_RemoveAllSignatures;
	asc_docs_api.prototype["asc_gotoSignature"] 						= asc_docs_api.prototype.asc_gotoSignature;
	asc_docs_api.prototype["asc_getSignatureSetup"] 					= asc_docs_api.prototype.asc_getSignatureSetup;

	// view modes
    asc_docs_api.prototype["asc_setContentDarkMode"]					= asc_docs_api.prototype.asc_setContentDarkMode;

	//comparison
	asc_docs_api.prototype["asc_CompareDocumentUrl"]                    = asc_docs_api.prototype.asc_CompareDocumentUrl;
	asc_docs_api.prototype["asc_CompareDocumentFile"]                   = asc_docs_api.prototype.asc_CompareDocumentFile;
	asc_docs_api.prototype["asc_CompareDocumentUrl_local"]              = asc_docs_api.prototype.asc_CompareDocumentUrl_local;
	asc_docs_api.prototype["asc_CompareDocumentFile_local"]             = asc_docs_api.prototype.asc_CompareDocumentFile_local;

	//merge
	asc_docs_api.prototype["asc_MergeDocumentUrl"]                    = asc_docs_api.prototype.asc_MergeDocumentUrl;
	asc_docs_api.prototype["asc_MergeDocumentFile"]                   = asc_docs_api.prototype.asc_MergeDocumentFile;
	asc_docs_api.prototype["asc_MergeDocumentUrl_local"]              = asc_docs_api.prototype.asc_MergeDocumentUrl_local;
	asc_docs_api.prototype["asc_MergeDocumentFile_local"]             = asc_docs_api.prototype.asc_MergeDocumentFile_local;


	asc_docs_api.prototype["asc_canEditGeometry"] 					    = asc_docs_api.prototype.asc_canEditGeometry;
	asc_docs_api.prototype["asc_editPointsGeometry"] 					= asc_docs_api.prototype.asc_editPointsGeometry;
	asc_docs_api.prototype["asc_getTableStylesPreviews"] 				= asc_docs_api.prototype.asc_getTableStylesPreviews;
	asc_docs_api.prototype["asc_generateTableStylesPreviews"] 		    = asc_docs_api.prototype.asc_generateTableStylesPreviews;
	asc_docs_api.prototype["asc_getDocumentProtection"] 		        = asc_docs_api.prototype.asc_getDocumentProtection;
	asc_docs_api.prototype["asc_setDocumentProtection"] 		        = asc_docs_api.prototype.asc_setDocumentProtection;

	// print-preview
	asc_docs_api.prototype["asc_initPrintPreview"] 	= asc_docs_api.prototype.asc_initPrintPreview;
	asc_docs_api.prototype["asc_drawPrintPreview"] 	= asc_docs_api.prototype.asc_drawPrintPreview;
	asc_docs_api.prototype["asc_closePrintPreview"] = asc_docs_api.prototype.asc_closePrintPreview;
	asc_docs_api.prototype["asc_getPageSize"] 		= asc_docs_api.prototype.asc_getPageSize;
	
	// auto hyphenation
	asc_docs_api.prototype["asc_setAutoHyphenation"]         = asc_docs_api.prototype.asc_setAutoHyphenation;
	asc_docs_api.prototype["asc_isAutoHyphenation"]          = asc_docs_api.prototype.asc_isAutoHyphenation;
	asc_docs_api.prototype["asc_setHyphenateCaps"]           = asc_docs_api.prototype.asc_setHyphenateCaps;
	asc_docs_api.prototype["asc_isHyphenateCaps"]            = asc_docs_api.prototype.asc_isHyphenateCaps;
	asc_docs_api.prototype["asc_setConsecutiveHyphenLimit"]  = asc_docs_api.prototype.asc_setConsecutiveHyphenLimit;
	asc_docs_api.prototype["asc_getConsecutiveHyphenLimit"]  = asc_docs_api.prototype.asc_getConsecutiveHyphenLimit;
	asc_docs_api.prototype["asc_setAutoHyphenationSettings"] = asc_docs_api.prototype.asc_setAutoHyphenationSettings;
	asc_docs_api.prototype["asc_getAutoHyphenationSettings"] = asc_docs_api.prototype.asc_getAutoHyphenationSettings;
	
	asc_docs_api.prototype["asc_putPageColor"] = asc_docs_api.prototype.asc_putPageColor;
	asc_docs_api.prototype["asc_getPageColor"] = asc_docs_api.prototype.asc_getPageColor;

	asc_docs_api.prototype["asc_insertTextFromFile"] = asc_docs_api.prototype.asc_insertTextFromFile;
	asc_docs_api.prototype["asc_insertTextFromUrl"] = asc_docs_api.prototype.asc_insertTextFromUrl;
	
	asc_docs_api.prototype["asc_showDeletedTextInVersionHistory"]     = asc_docs_api.prototype.asc_showDeletedTextInVersionHistory;
	asc_docs_api.prototype["asc_isShowedDeletedTextInVersionHistory"] = asc_docs_api.prototype.asc_isShowedDeletedTextInVersionHistory;
	asc_docs_api.prototype["asc_hideDeletedTextInVersionHistory"]     = asc_docs_api.prototype.asc_hideDeletedTextInVersionHistory;

	asc_docs_api.prototype["asc_getCoHistory"] = asc_docs_api.prototype.asc_getCoHistory;

	CDocInfoProp.prototype['get_PageCount']             = CDocInfoProp.prototype.get_PageCount;
	CDocInfoProp.prototype['put_PageCount']             = CDocInfoProp.prototype.put_PageCount;
	CDocInfoProp.prototype['get_WordsCount']            = CDocInfoProp.prototype.get_WordsCount;
	CDocInfoProp.prototype['put_WordsCount']            = CDocInfoProp.prototype.put_WordsCount;
	CDocInfoProp.prototype['get_ParagraphCount']        = CDocInfoProp.prototype.get_ParagraphCount;
	CDocInfoProp.prototype['put_ParagraphCount']        = CDocInfoProp.prototype.put_ParagraphCount;
	CDocInfoProp.prototype['get_SymbolsCount']          = CDocInfoProp.prototype.get_SymbolsCount;
	CDocInfoProp.prototype['put_SymbolsCount']          = CDocInfoProp.prototype.put_SymbolsCount;
	CDocInfoProp.prototype['get_SymbolsWSCount']        = CDocInfoProp.prototype.get_SymbolsWSCount;
	CDocInfoProp.prototype['put_SymbolsWSCount']        = CDocInfoProp.prototype.put_SymbolsWSCount;
	CContextMenuData.prototype['get_Type']    = CContextMenuData.prototype.get_Type;
	CContextMenuData.prototype['get_X']       = CContextMenuData.prototype.get_X;
	CContextMenuData.prototype['get_Y']       = CContextMenuData.prototype.get_Y;
	CContextMenuData.prototype['get_PageNum'] = CContextMenuData.prototype.get_PageNum;
	CContextMenuData.prototype['is_Header']   = CContextMenuData.prototype.is_Header;
	window['Asc']['asc_CCommentDataWord'] = window['Asc'].asc_CCommentDataWord = asc_CCommentDataWord;
	asc_CCommentDataWord.prototype['asc_getText']         = asc_CCommentDataWord.prototype.asc_getText;
	asc_CCommentDataWord.prototype['asc_putText']         = asc_CCommentDataWord.prototype.asc_putText;
	asc_CCommentDataWord.prototype['asc_getTime']         = asc_CCommentDataWord.prototype.asc_getTime;
	asc_CCommentDataWord.prototype['asc_putTime']         = asc_CCommentDataWord.prototype.asc_putTime;
	asc_CCommentDataWord.prototype['asc_getOnlyOfficeTime']         = asc_CCommentDataWord.prototype.asc_getOnlyOfficeTime;
	asc_CCommentDataWord.prototype['asc_putOnlyOfficeTime']         = asc_CCommentDataWord.prototype.asc_putOnlyOfficeTime;
	asc_CCommentDataWord.prototype['asc_getUserId']       = asc_CCommentDataWord.prototype.asc_getUserId;
	asc_CCommentDataWord.prototype['asc_putUserId']       = asc_CCommentDataWord.prototype.asc_putUserId;
	asc_CCommentDataWord.prototype['asc_getProviderId']   = asc_CCommentDataWord.prototype.asc_getProviderId;
	asc_CCommentDataWord.prototype['asc_putProviderId']   = asc_CCommentDataWord.prototype.asc_putProviderId;
	asc_CCommentDataWord.prototype['asc_getUserName']     = asc_CCommentDataWord.prototype.asc_getUserName;
	asc_CCommentDataWord.prototype['asc_putUserName']     = asc_CCommentDataWord.prototype.asc_putUserName;
	asc_CCommentDataWord.prototype['asc_getInitials']     = asc_CCommentDataWord.prototype.asc_getInitials;
	asc_CCommentDataWord.prototype['asc_putInitials']     = asc_CCommentDataWord.prototype.asc_putInitials;
	asc_CCommentDataWord.prototype['asc_getQuoteText']    = asc_CCommentDataWord.prototype.asc_getQuoteText;
	asc_CCommentDataWord.prototype['asc_putQuoteText']    = asc_CCommentDataWord.prototype.asc_putQuoteText;
	asc_CCommentDataWord.prototype['asc_getSolved']       = asc_CCommentDataWord.prototype.asc_getSolved;
	asc_CCommentDataWord.prototype['asc_putSolved']       = asc_CCommentDataWord.prototype.asc_putSolved;
	asc_CCommentDataWord.prototype['asc_getGuid']         = asc_CCommentDataWord.prototype.asc_getGuid;
	asc_CCommentDataWord.prototype['asc_putGuid']         = asc_CCommentDataWord.prototype.asc_putGuid;
	asc_CCommentDataWord.prototype['asc_getReply']        = asc_CCommentDataWord.prototype.asc_getReply;
	asc_CCommentDataWord.prototype['asc_addReply']        = asc_CCommentDataWord.prototype.asc_addReply;
	asc_CCommentDataWord.prototype['asc_getRepliesCount'] = asc_CCommentDataWord.prototype.asc_getRepliesCount;
	asc_CCommentDataWord.prototype['asc_getDocumentFlag'] = asc_CCommentDataWord.prototype.asc_getDocumentFlag;
	asc_CCommentDataWord.prototype['asc_putDocumentFlag'] = asc_CCommentDataWord.prototype.asc_putDocumentFlag;
	asc_CCommentDataWord.prototype['asc_getUserData']     = asc_CCommentDataWord.prototype.asc_getUserData;
	asc_CCommentDataWord.prototype['asc_putUserData']     = asc_CCommentDataWord.prototype.asc_putUserData;

	AscCommon.setUpAllFonts = function()
	{
		var testFontCurrent = 0;
		var testFontsInterval = 0;

		var _fonts = [];
		for (var i = 0; i < AscFonts.g_font_infos.length; i++)
			_fonts.push(new AscFonts.CFont(AscFonts.g_font_infos[i].Name));

		console.log("start...");

		function logFont(font, name)
		{
			var face = font.m_pFace;
			if (!face.os2 || face.os2.version == 0xFFFF)
				return;

			var isTypo = ((face.os2.fsSelection & 128) != 0);
			//if (!isTypo)
			//	return;

			if (isTypo && (face.height != (face.os2.sTypoAscender - face.os2.sTypoDescender + face.os2.sTypoLineGap)))
			{
			    console.log("[" + face.family_name + "] typo");
				console.log(face.ascender + ", " + face.descender + ", " + face.height);
                console.log(face.os2.sTypoAscender + ", " + face.os2.sTypoDescender + ", " + (face.os2.sTypoAscender - face.os2.sTypoDescender + face.os2.sTypoLineGap));
			}

            if (!isTypo && (face.height != (face.os2.usWinAscent + face.os2.usWinDescent)))
            {
                console.log("[" + face.family_name + "] win");
                console.log(face.ascender + ", " + face.descender + ", " + face.height);
                console.log(face.os2.usWinAscent + ", " + face.os2.usWinDescent + ", " + (face.os2.usWinAscent + face.os2.usWinDescent));
            }
		}

        editor.asyncMethodCallback = function() {

            testFontsInterval = setInterval(function(){
                if (testFontCurrent >= AscFonts.g_font_infos.length)
                {
                    clearInterval(testFontsInterval);
                    console.log("end");
                    return;
                }

                var _info = AscFonts.g_font_infos[testFontCurrent++];

                if (_info.indexR != -1)
				{
					var fontfile = AscCommon.g_font_loader.fontFiles[_info.indexR];
                    var pFontFile = AscCommon.g_fontManager.LoadFont(fontfile, _info.faceIndexR, 12, false, false, false, false, true);

					logFont(pFontFile, "regular");
                }

                if (_info.indexB != -1)
                {
                    var fontfile = AscCommon.g_font_loader.fontFiles[_info.indexB];
                    var pFontFile = AscCommon.g_fontManager.LoadFont(fontfile, _info.faceIndexB, 12, true, false, false, false, true);

                    logFont(pFontFile, "bold");
                }

                if (_info.indexI != -1)
                {
                    var fontfile = AscCommon.g_font_loader.fontFiles[_info.indexI];
                    var pFontFile = AscCommon.g_fontManager.LoadFont(fontfile, _info.faceIndexI, 12, false, true, false, false, true);

                    logFont(pFontFile, "italic");
                }

                if (_info.indexBI != -1)
                {
                    var fontfile = AscCommon.g_font_loader.fontFiles[_info.indexBI];
                    var pFontFile = AscCommon.g_fontManager.LoadFont(fontfile, _info.faceIndexBI, 12, true, true, false, false, true);

                    logFont(pFontFile, "bold italic");
                }

            }, 10);

        };

        AscCommon.g_font_loader.LoadDocumentFonts2(_fonts);
	};

	AscCommon['DocumentEditorApi'] = AscCommon.DocumentEditorApi = asc_docs_api;
	
})(window, window.document);

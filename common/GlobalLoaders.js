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

(function(window, document){
    // Import
    var g_fontApplication = AscFonts.g_fontApplication;
    var ImageLoadStatus   = AscFonts.ImageLoadStatus;
    var CImage            = AscFonts.CImage;

    function CGlobalFontLoader()
    {
        // сначала хотел писать "вытеснение" из этого мапа.
        // но тогда нужно хранить base64 строки. Это не круто. По памяти - даже
        // выигрыш будет. Не особо то шрифты жмутся lzw или deflate
        // поэтому лучше из памяти будем удалять base64 строки
        this.fonts_streams = [];

        // теперь вся информация о всех возможных шрифтах. Они во всех редакторах должны быть одни и те же
        this.fontFilesPath = "../../../../fonts/";
        this.fontFiles = AscFonts.g_font_files;
        this.fontInfos = AscFonts.g_font_infos;
        this.map_font_index = AscFonts.g_map_font_index;

        // динамическая подгрузка шрифтов
        this.ThemeLoader = null;
        this.Api = null;
        this.fonts_loading = [];
        this.bIsLoadDocumentFirst = false;

        // информация для загрузки по одному шрифту
        this.currentInfoLoaded = null;
        this.loadFontCallBack     = null;
        this.loadFontCallBackArgs = null;

        // при переоткрытиях файла - заменить на LoadDocumentFonts2
        this.IsLoadDocumentFonts2 = false;

        this.check_loaded_timer_id = -1;
        this.endLoadingCallback = null;
		
		// Счетчик загрузки шрифтов через метод LoadFonts
		this.loadFontsCounter = 0;

        this.perfStart = 0;

        this.put_Api = function(api)
        {
            this.Api = api;
        };

        // добавляем шрифт в список для загрузки
        this.AddLoadFonts = function(name, need_styles)
        {
            var fontinfo = g_fontApplication.GetFontInfo(name);
            this.fonts_loading[this.fonts_loading.length] = fontinfo;
            this.fonts_loading[this.fonts_loading.length - 1].NeedStyles = (need_styles === undefined) ? 0x0F : need_styles;
			return fontinfo;
        };
        this.AddLoadFontsNotPick = function(info, need_styles)
        {
            this.fonts_loading[this.fonts_loading.length] = info;
            this.fonts_loading[this.fonts_loading.length - 1].NeedStyles = (need_styles === undefined) ? 0x0F : need_styles;
        };

        // проверить все fontinfo из fonts_loading на нужность загрузки, и вернуть есть ли хоть один заново запущенный
        this.CheckFontsNeedLoadingLoad = function()
        {
            let fonts = this.fonts_loading;
            let isNeed = false;
            for (let i = 0, len = fonts.length; i < len; i++)
            {
                if (true === fonts[i].CheckFontLoadStyles(this))
                    isNeed = true;
            }
            return isNeed;
        };

        // нужно ли грузить хоть один из списка (без запуска загрузки)
        this.CheckFontsNeedLoading = function(fonts)
        {
            for (let i in fonts)
            {
                let info = g_fontApplication.GetFontInfo(fonts[i].name);
                if (true === info.CheckFontLoadStylesNoLoad(this))
                    return true;
            }
            return false;
        };

        this.isWorking = function()
        {
            return (this.check_loaded_timer_id !== -1) ? true : false;
        };

        this.LoadDocumentFonts = function(fonts)
        {
            if (this.IsLoadDocumentFonts2)
                return this.LoadDocumentFonts2(fonts);

            let gui_fonts = [];
            let gui_count = 0;
            for (let i = 0; i < this.fontInfos.length; i++)
            {
                let info = this.fontInfos[i];
                if (info.Name !== "ASCW3")
                    gui_fonts[gui_count++] = new AscFonts.CFont(info.Name, "", info.Thumbnail);
            }

            // сначала заполняем массив this.fonts_loading объекстами fontinfo
            for (let i in fonts)
            {
                this.AddLoadFonts(fonts[i].name, fonts[i].NeedStyles);
            }

            this.Api.sync_InitEditorFonts(gui_fonts);

            // но только если редактор!!!
            if (this.Api.IsNeedDefaultFonts())
            {
                // теперь добавим шрифты, без которых редактор как без рук (спецсимволы + дефолтовые стили документа)
                this.AddLoadFonts("Arial", 0x0F);
                this.AddLoadFonts("Symbol", 0x0F);
                this.AddLoadFonts("Wingdings", 0x0F);
                this.AddLoadFonts("Courier New", 0x0F);
                this.AddLoadFonts("Times New Roman", 0x0F);
            }

            this.Api.asyncFontsDocumentStartLoaded();

            this.bIsLoadDocumentFirst = true;

            this.CheckFontsNeedLoadingLoad();
            this._LoadFonts();
        };

        this.LoadDocumentFonts2 = function(fonts, blockType, callback)
        {
            if (this.isWorking())
            {
                // такого быть не должно
                return;
            }

            this.endLoadingCallback = (undefined !== callback) ? callback : null;
            this.BlockOperationType = blockType;

            // сначала заполняем массив this.fonts_loading объекстами fontinfo
            for (var i in fonts)
                this.AddLoadFonts(fonts[i].name, 0x0F);

            if (null == this.ThemeLoader)
                this.Api.asyncFontsDocumentStartLoaded(this.BlockOperationType);
            else
                this.ThemeLoader.asyncFontsStartLoaded();

            this.CheckFontsNeedLoadingLoad();
            this._LoadFonts();
        };

        this._LoadFonts = function()
        {
            if (this.bIsLoadDocumentFirst === true && 0 === this.perfStart && this.fonts_loading.length > 0)
                this.perfStart = performance.now();

            if (0 === this.fonts_loading.length)
            {
                if (this.perfStart > 0)
                {
                    let perfEnd = performance.now();
                    AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onLoadFonts", perfEnd - this.perfStart), this.Api);
                    this.perfStart = 0;
                }

                if (null != this.endLoadingCallback)
                {
                    this.endLoadingCallback.call(this.Api);
                    this.endLoadingCallback = null;
                }
                else if (null == this.ThemeLoader)
                    this.Api.asyncFontsDocumentEndLoaded(this.BlockOperationType);
                else
                    this.ThemeLoader.asyncFontsEndLoaded();

                this.BlockOperationType = undefined;
                this.bIsLoadDocumentFirst = false;
                return;
            }

            if (this.fonts_loading[0].CheckFontLoadStyles(this))
            {
                let _t = this;
                this.check_loaded_timer_id = setTimeout(function(){
                    _t.check_loaded_list();
                }, 50);
            }
            else
            {
                if (this.bIsLoadDocumentFirst === true)
                {
                    this.Api.OpenDocumentProgress.CurrentFont++;
                    this.Api.SendOpenProgress();
                }

                this.fonts_loading.shift();
                this._LoadFonts();
            }
        };

        this.check_loaded_list = function()
        {
            this.check_loaded_timer_id = -1;
            if (0 === this.fonts_loading.length)
            {
                // значит асинхронно удалилось
                this._LoadFonts();
                return;
            }

            let current = this.fonts_loading[0];
            let isNeed = current.CheckFontLoadStyles(this);
            if (true === isNeed)
            {
                let _t = this;
                this.check_loaded_timer_id = setTimeout(function(){
                    _t.check_loaded_list();
                }, 50);
            }
            else
            {
                if (this.bIsLoadDocumentFirst === true)
                {
                    this.Api.OpenDocumentProgress.CurrentFont++;
                    this.Api.SendOpenProgress();
                }

                this.fonts_loading.shift();
                this._LoadFonts();
            }
        };

        // одиночная загрузка шрифта
        this.LoadFont = function(fontinfo, loadFontCallBack, loadFontCallBackArgs)
        {
            this.currentInfoLoaded = fontinfo;
            this.currentInfoLoaded.NeedStyles = 15; // все стили

            let isNeed = this.currentInfoLoaded.CheckFontLoadStyles(this);

            if ( undefined === loadFontCallBack )
            {
                this.loadFontCallBack     = this.Api.asyncFontEndLoaded;
                this.loadFontCallBackArgs = this.currentInfoLoaded;
            }
            else
            {
                this.loadFontCallBack     = loadFontCallBack;
                this.loadFontCallBackArgs = loadFontCallBackArgs;
            }

            if (isNeed)
            {
                this.Api.asyncFontStartLoaded();
                let _t = this;
                setTimeout(function() {
                    _t.check_loaded();
                }, 20);
                return true;
            }
            else
            {
                this.currentInfoLoaded = null;
                return false;
            }
        };
        this.check_loaded = function()
        {
            if (!this.currentInfoLoaded)
                return;

            let isNeed = this.currentInfoLoaded.CheckFontLoadStyles(this);
            if (isNeed)
            {
                let _t = this;
                setTimeout(function() {
                    _t.check_loaded();
                }, 50);
            }
            else
            {
                this.loadFontCallBack.call( this.Api, this.loadFontCallBackArgs );
                this.currentInfoLoaded = null;
            }
        };

        // используется только в тестовом примере (предзагрузка в кэш браузера)
        this.LoadFontsFromServer = function(fonts)
        {
            let count = fonts.length;
            for (let i = 0; i < count; i++)
            {
                let info = g_fontApplication.GetFontInfo(fonts[i]);
                info && info.LoadFontsFromServer(this);
            }
        };
		
		this.isFontLoadInProgress = function()
		{
			return (this.isWorking() || this.loadFontsCounter > 0);
		};
		
		this.LoadFonts = function(fonts, callback)
		{
			++this.loadFontsCounter;
			
			let fontMap = {}
			
			if (fonts && Array.isArray(fonts))
			{
				for (let i = 0; i < fonts.length; ++i)
				{
					let name = fonts[i];
					fontMap[name] = AscFonts.g_fontApplication.GetFontInfo(name);
					fontMap[name].NeedStyles = 15;
				}
			}
			else
			{
				for (let name in fonts)
				{
					fontMap[name] = AscFonts.g_fontApplication.GetFontInfo(name);
					fontMap[name].NeedStyles = 15;
				}
			}
			
			
			let globalLoader = this;
			
			let checkLoaded = function()
			{
				let needLoad = 0;
				for (let name in fontMap)
				{
					if (!fontMap[name].CheckFontLoadStyles(globalLoader))
						delete fontMap[name];
					else
						++needLoad;
				}
				
				if (needLoad)
					return setTimeout(checkLoaded, 50);
				
				if (callback)
					callback();
				
				--globalLoader.loadFontsCounter;
			};
			checkLoaded();
		}
    }

    function CGlobalImageLoader()
    {
        this.map_image_index = {};

        // loading
        this.Api = null;
        this.ThemeLoader = null;
        this.images_loading = null;

        this.bIsLoadDocumentFirst = false;
        this.bIsAsyncLoadDocumentImages = false;

        this.nNoByOrderCounter = 0;

        this.isBlockchainSupport = false;
        var oThis = this;

        if (window["AscDesktopEditor"] &&
            window["AscDesktopEditor"]["IsLocalFile"] &&
            window["AscDesktopEditor"]["isBlockchainSupport"])
        {
            this.isBlockchainSupport = (window["AscDesktopEditor"]["isBlockchainSupport"]() && !window["AscDesktopEditor"]["IsLocalFile"]());

            if (this.isBlockchainSupport)
            {
                Image.prototype.preload_crypto = function(_url)
                {
                    window["crypto_images_map"] = window["crypto_images_map"] || {};
                    if (!window["crypto_images_map"][_url])
                        window["crypto_images_map"][_url] = [];
                    window["crypto_images_map"][_url].push(this);

                    window["AscDesktopEditor"]["PreloadCryptoImage"](_url, AscCommon.g_oDocumentUrls.getLocal(_url));

                    oThis.Api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadImage);
                };

                Image.prototype["onload_crypto"] = function(_src, _crypto_data)
                {
                    if (_crypto_data && AscCommon.EncryptionWorker && AscCommon.EncryptionWorker.isCryptoImages())
                    {
                        AscCommon.EncryptionWorker.decryptImage(_src, this, _crypto_data);
                        return;
                    }
					this.crossOrigin = "";
                    this.src = _src;
                    oThis.Api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadImage);
                };
            }
        }

        this.put_Api = function(api)
        {
            this.Api = api;

            if (this.Api.IsAsyncOpenDocumentImages !== undefined)
            {
                this.bIsAsyncLoadDocumentImages = this.Api.IsAsyncOpenDocumentImages();
                if (this.bIsAsyncLoadDocumentImages)
                {
                    if (undefined === this.Api.asyncImageEndLoadedBackground)
                        this.bIsAsyncLoadDocumentImages = false;
                }
            }
        };

        this.loadImageByUrl = function(image, url, isDisableCrypto)
        {
            if (this.isBlockchainSupport && (true !== isDisableCrypto))
                image.preload_crypto(url);
            else
                image.src = url;
        };
        
        this.LoadDocumentImages = function(images, isCheckExists)
        {
            if (isCheckExists)
            {
                for (let i = images.length - 1; i >= 0; i--)
                {
                    let id = AscCommon.getFullImageSrc2(images[i]);
                    if (this.map_image_index[id] && (this.map_image_index[id].Status === ImageLoadStatus.Complete))
                    {
                        images.splice(i, 1);
                    }
                }

                if (0 === images.length)
                    return;
            }

            // сначала заполним массив

            this.images_loading = [];
            for (let id in images)
            {
                this.images_loading[this.images_loading.length] = AscCommon.getFullImageSrc2(images[id]);
            }
            if (this.ThemeLoader == null)
                this.Api.asyncImagesDocumentStartLoaded(this.images_loading);
            else
                this.ThemeLoader.asyncImagesStartLoaded(this.images_loading);

            if (!this.bIsAsyncLoadDocumentImages)
            {
				this.nNoByOrderCounter = 0;
                this._LoadImages();
            }
            else
            {
                let len = this.images_loading.length;
                for (let i = 0; i < len; i++)
                    this.LoadImageAsync(i);

                this.images_loading.splice(0, len);

                if (this.ThemeLoader == null)
                    this.Api.asyncImagesDocumentEndLoaded();
                else
                    this.ThemeLoader.asyncImagesEndLoaded();
            }
        };

        this._LoadImages = function()
        {
			for (let i = 0; i < this.images_loading.length; i++)
            {
				let id = this.images_loading[i];
				if (this.map_image_index[id] && (this.map_image_index[id].Status === ImageLoadStatus.Complete))
                {
                    this.images_loading.splice(i, 1);
                }
            }
			let count_images = this.images_loading.length;

            if (0 === count_images)
            {
				this.nNoByOrderCounter = 0;

                if (this.ThemeLoader == null)
                    this.Api.asyncImagesDocumentEndLoaded();
                else
                    this.ThemeLoader.asyncImagesEndLoaded();

                return;
            }

            for (let i = 0; i < count_images; i++)
			{
				var _id = this.images_loading[i];
				var oImage = new CImage(_id);
				oImage.Status = ImageLoadStatus.Loading;
				oImage.Image = new Image();
				oThis.map_image_index[oImage.src] = oImage;
				oImage.Image.parentImage = oImage;
				oImage.Image.onload = function ()
				{
					this.parentImage.Status = ImageLoadStatus.Complete;
					oThis.nNoByOrderCounter++;

					if (oThis.bIsLoadDocumentFirst === true)
					{
						oThis.Api.OpenDocumentProgress.CurrentImage++;
						oThis.Api.SendOpenProgress();
					}

					if (oThis.nNoByOrderCounter === oThis.images_loading.length)
                    {
						oThis.images_loading = [];
						oThis._LoadImages();
                    }
				};
				oImage.Image.onerror = function ()
				{
					this.parentImage.Status = ImageLoadStatus.Complete;
					this.parentImage.Image = null;
					oThis.nNoByOrderCounter++;

					if (oThis.bIsLoadDocumentFirst === true)
					{
						oThis.Api.OpenDocumentProgress.CurrentImage++;
						oThis.Api.SendOpenProgress();
					}

					if (oThis.nNoByOrderCounter === oThis.images_loading.length)
					{
						oThis.images_loading = [];
						oThis._LoadImages();
					}
				};
				AscCommon.backoffOnErrorImg(oImage.Image, function(img) {
					oThis.loadImageByUrl(img, img.src);
				});
				//oImage.Image.crossOrigin = 'anonymous';
				oThis.loadImageByUrl(oImage.Image, oImage.src);
			}
        };

        this.LoadImage = function(src, type)
        {
            var image = this.map_image_index[src];
            if (undefined != image)
                return image;

            this.Api.asyncImageStartLoaded();

            var oImage = new CImage(src);

            // просто прокидываем параметр
            oImage.Type = type;

            oImage.Image = new Image();
            oImage.Status = ImageLoadStatus.Loading;
            oThis.map_image_index[oImage.src] = oImage;

            oImage.Image.onload = function() {
                oImage.Status = ImageLoadStatus.Complete;
                oThis.Api.asyncImageEndLoaded(oImage);
            };
            oImage.Image.onerror = function() {
                oImage.Image = null;
                oImage.Status = ImageLoadStatus.Complete;
                oThis.Api.asyncImageEndLoaded(oImage);
            };
            AscCommon.backoffOnErrorImg(oImage.Image, function(img) {
                oThis.loadImageByUrl(img, img.src);
            });
            //oImage.Image.crossOrigin = 'anonymous';
            this.loadImageByUrl(oImage.Image, oImage.src);
            return null;
        };

        this.LoadImageAsync = function(i)
        {
            var oImage = new CImage(this.images_loading[i]);

            oImage.Status = ImageLoadStatus.Loading;
            oImage.Image = new Image();
            oThis.map_image_index[oImage.src] = oImage;

            oImage.Image.onload = function() {
                oImage.Status = ImageLoadStatus.Complete;
                oThis.Api.asyncImageEndLoadedBackground(oImage);
            };
            oImage.Image.onerror = function() {
                oImage.Status = ImageLoadStatus.Complete;
                oImage.Image = null;
                oThis.Api.asyncImageEndLoadedBackground(oImage);
            };
            AscCommon.backoffOnErrorImg(oImage.Image, function(img) {
                oThis.loadImageByUrl(img, img.src);
            });
            //oImage.Image.crossOrigin = 'anonymous';
            oThis.loadImageByUrl(oImage.Image, oImage.src);
        };

        this.LoadImagesWithCallback = function(arr, loadImageCallBack, loadImageCallBackArgs, isDisableCrypto)
        {
            let arrAsync = [];
            for (let i = 0; i < arr.length; i++)
            {
                if (this.map_image_index[arr[i]] === undefined)
                    arrAsync.push(arr[i]);
            }

            if (arrAsync.length == 0)
            {
				loadImageCallBack.call(this.Api, loadImageCallBackArgs);
                return;
            }

            let asyncImageCounter = arrAsync.length;
            const callback = loadImageCallBack.bind(this.Api, loadImageCallBackArgs);

			for (let i = 0; i < arrAsync.length; i++)
			{
				var oImage = new CImage(arrAsync[i]);
				oImage.Image = new Image();
				oImage.Image.parentImage = oImage;
				oImage.Status = ImageLoadStatus.Loading;
				this.map_image_index[oImage.src] = oImage;

				oImage.Image.onload = function ()
				{
					this.parentImage.Status = ImageLoadStatus.Complete;
                    asyncImageCounter--;

					if (asyncImageCounter === 0)
					    callback();
				};
				oImage.Image.onerror = function ()
				{
					this.parentImage.Image = null;
					this.parentImage.Status = ImageLoadStatus.Complete;
                    asyncImageCounter--;

					if (asyncImageCounter === 0)
						callback();
				};
				AscCommon.backoffOnErrorImg(oImage.Image, function(img) {
					oThis.loadImageByUrl(img, img.src);
				});
				//oImage.Image.crossOrigin = 'anonymous';
                this.loadImageByUrl(oImage.Image, oImage.src, isDisableCrypto);
			}
        };
    }

    //---------------------------------------------------------export---------------------------------------------------
    window['AscCommon'] = window['AscCommon'] || {};
    window['AscCommon'].CGlobalFontLoader = CGlobalFontLoader;
    window['AscCommon'].g_font_loader = new CGlobalFontLoader();
    window['AscCommon'].g_image_loader = new CGlobalImageLoader();
})(window, window.document);

/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function(window, document){
    // Import
    var g_fontApplication = AscFonts.g_fontApplication;
    var ImageLoadStatus   = AscFonts.ImageLoadStatus;
    var CImage            = AscFonts.CImage;

    function CGlobalFontLoader()
    {
        // Initially wanted to implement "eviction" from this map.
        // But then we would need to store base64 strings. That's not cool. Memory-wise -
        // there would even be a gain. Fonts don't compress well with lzw or deflate anyway.
        // So it's better to delete base64 strings from memory.
        this.fonts_streams = [];

        // Now all information about all available fonts. They should be the same in all editors.
        this.fontFilesPath = "../../../../fonts/";
        this.fontFiles = AscFonts.g_font_files;
        this.fontInfos = AscFonts.g_font_infos;
        this.map_font_index = AscFonts.g_map_font_index;

        // Dynamic font loading
        this.ThemeLoader = null;
        this.Api = null;
        this.fonts_loading = [];
        this.bIsLoadDocumentFirst = false;

        // Information for loading a single font
        this.currentInfoLoaded = null;
        this.loadFontCallBack     = null;
        this.loadFontCallBackArgs = null;

        // When reopening files - replace with LoadDocumentFonts2
        this.IsLoadDocumentFonts2 = false;

        this.check_loaded_timer_id = -1;
        this.endLoadingCallback = null;
		
		// Counter for font loading via the LoadFonts method
		this.loadFontsCounter = 0;

        this.perfStart = 0;

        this.put_Api = function(api)
        {
            this.Api = api;
        };

        // Add font to the loading list
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

        // Check all fontinfo from fonts_loading for loading necessity, and return whether at least one was started again
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

        // Check if at least one from the list needs loading (without starting the load)
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

            // First, fill the this.fonts_loading array with fontinfo objects
            for (let i in fonts)
            {
                this.AddLoadFonts(fonts[i].name, fonts[i].NeedStyles);
            }

            this.Api.sync_InitEditorFonts(gui_fonts);

            // But only if it's an editor!!!
            if (this.Api.IsNeedDefaultFonts())
            {
                // Now add fonts that the editor can't function without (special symbols + default document styles)
                this.AddLoadFonts("Arial", 0x0F);
                this.AddLoadFonts("Symbol", 0x0F);
                this.AddLoadFonts("Wingdings", 0x0F);
                this.AddLoadFonts("Courier New", 0x0F);
                this.AddLoadFonts("Times New Roman", 0x0F);
            }

			this.BlockOperationType = undefined;
            this.Api.asyncFontsDocumentStartLoaded();

            this.bIsLoadDocumentFirst = true;

            this.CheckFontsNeedLoadingLoad();
            this._LoadFonts();
        };

        this.LoadDocumentFonts2 = function(fonts, blockType, callback)
        {
            if (this.isWorking())
            {
                // This shouldn't happen
                return;
            }

            this.endLoadingCallback = (undefined !== callback) ? callback : null;
            this.BlockOperationType = blockType;

            // First, fill the this.fonts_loading array with fontinfo objects
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
                    this.Api.sync_EndAction(undefined === this.BlockOperationType ? Asc.c_oAscAsyncActionType.BlockInteraction : this.BlockOperationType, Asc.c_oAscAsyncAction.LoadDocumentFonts);
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
                // Means it was deleted asynchronously
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

        // Single font loading
        this.LoadFont = function(fontinfo, loadFontCallBack, loadFontCallBackArgs)
        {
            this.currentInfoLoaded = fontinfo;
            this.currentInfoLoaded.NeedStyles = 15; // all styles

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

        // Used only in test example (preloading into browser cache)
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

        this.bIsLoadDocumentFirst = false;
        this.bIsAsyncLoadDocumentImages = false;

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

	    this.LoadDocumentImages = function (images, isCheckExists, syncImages) {
		    if (isCheckExists) {
			    for (let i = images.length - 1; i >= 0; i--) {
				    let id = AscCommon.getFullImageSrc2(images[i]);
				    if (this.map_image_index[id] && (this.map_image_index[id].Status === ImageLoadStatus.Complete)) {
					    images.splice(i, 1);
				    }
			    }

			    if (0 === images.length)
				    return;
		    }

		    // First, fill the array

		    const oRequiredSyncImages = {};
		    const arrImagesLoading = [];
		    for (let id in images) {
			    const sFullImageSrc = AscCommon.getFullImageSrc2(images[id]);
			    arrImagesLoading.push(sFullImageSrc);
			    if (syncImages && syncImages[images[id]]) {
				    oRequiredSyncImages[sFullImageSrc] = true;
			    }
		    }
		    if (syncImages) {
			    for (let id in syncImages) {
				    const sFullImageSrc = AscCommon.getFullImageSrc2(id);
				    if (!oRequiredSyncImages[sFullImageSrc]) {
					    arrImagesLoading.push(sFullImageSrc);
					    oRequiredSyncImages[sFullImageSrc] = true;
				    }
			    }
		    }
		    if (this.ThemeLoader == null)
			    this.Api.asyncImagesDocumentStartLoaded(arrImagesLoading);
		    else
			    this.ThemeLoader.asyncImagesStartLoaded(arrImagesLoading);

		    if (!this.bIsAsyncLoadDocumentImages) {
			    this._LoadImages(arrImagesLoading);
		    } else {
			    this._LoadImagesAsync(arrImagesLoading, oRequiredSyncImages);
		    }
	    };

	    this._LoadImages = function (arrImages) {
		    let fOnEachImageLoadCallback;
		    if (oThis.bIsLoadDocumentFirst === true) {
			    fOnEachImageLoadCallback = function () {
				    oThis.Api.OpenDocumentProgress.CurrentImage++;
				    oThis.Api.SendOpenProgress();
			    };
		    }
		    this.LoadImagesWithCallback(arrImages, function () {
			    if (oThis.ThemeLoader == null)
				    oThis.Api.asyncImagesDocumentEndLoaded();
			    else
				    oThis.ThemeLoader.asyncImagesEndLoaded();
		    }, [], false, fOnEachImageLoadCallback);
	    };

	    this._LoadImagesAsync = function (arrImages, oRequiredSyncImages) {
		    const arrAsyncImages = [];
		    const arrSyncImages = [];
		    for (let i = 0; i < arrImages.length; i += 1) {
			    if (oRequiredSyncImages[arrImages[i]]) {
				    arrSyncImages.push(arrImages[i]);
			    } else {
				    arrAsyncImages.push(arrImages[i]);
			    }
		    }
		    let fOnEachImageLoadCallback;
		    if (oThis.bIsLoadDocumentFirst === true) {
			    fOnEachImageLoadCallback = function () {
				    oThis.Api.OpenDocumentProgress.CurrentImage++;
				    oThis.Api.SendOpenProgress();
			    };
		    }
		    this.LoadImagesWithCallback(arrSyncImages, function () {
			    for (let i = 0; i < arrAsyncImages.length; i += 1) {
				    oThis.LoadImageAsync(arrAsyncImages[i]);
			    }

			    if (oThis.ThemeLoader == null)
				    oThis.Api.asyncImagesDocumentEndLoaded();
			    else
				    oThis.ThemeLoader.asyncImagesEndLoaded();
		    }, [], false, fOnEachImageLoadCallback);
	    };

        this.LoadImage = function(src, type)
        {
            var image = this.map_image_index[src];
            if (undefined != image)
                return image;

            this.Api.asyncImageStartLoaded();

            var oImage = new CImage(src);

            // Just passing through the parameter
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
                // Remove crossOrigin on retry to maximize display success
                img.crossOrigin = null;
                oThis.loadImageByUrl(img, img.src);
            });
            // Enable CORS for cross-origin images to allow canvas manipulation
            oImage.Image.crossOrigin = 'anonymous';
            this.loadImageByUrl(oImage.Image, oImage.src);
            return null;
        };

        this.LoadImageAsync = function(imgSrc)
        {
            if (oThis.map_image_index[imgSrc])
                return;

            var oImage = new CImage(imgSrc);

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
                // Remove crossOrigin on retry to maximize display success
                img.crossOrigin = null;
                oThis.loadImageByUrl(img, img.src);
            });
            // Enable CORS for cross-origin images to allow canvas manipulation
            oImage.Image.crossOrigin = 'anonymous';
            oThis.loadImageByUrl(oImage.Image, oImage.src);
        };

        this.LoadImagesWithCallback = function(arr, loadImageCallBack, loadImageCallBackArgs, isDisableCrypto, onEachImageLoadCallback)
        {
            let arrAsync = [];
            for (let i = 0; i < arr.length; i++)
            {
                if (this.map_image_index[arr[i]] && (this.map_image_index[arr[i]].Status === ImageLoadStatus.Complete))
									continue;

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
					onEachImageLoadCallback && onEachImageLoadCallback();
					if (asyncImageCounter === 0)
					    callback();
				};
				oImage.Image.onerror = function ()
				{
					this.parentImage.Image = null;
					this.parentImage.Status = ImageLoadStatus.Complete;
                    asyncImageCounter--;
					onEachImageLoadCallback && onEachImageLoadCallback();
					if (asyncImageCounter === 0)
						callback();
				};
				AscCommon.backoffOnErrorImg(oImage.Image, function(img) {
					// Remove crossOrigin on retry to maximize display success
					img.crossOrigin = null;
					oThis.loadImageByUrl(img, img.src);
				});
				// Enable CORS for cross-origin images to allow canvas manipulation
				oImage.Image.crossOrigin = 'anonymous';
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

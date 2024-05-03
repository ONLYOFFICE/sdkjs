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

function CCacheImage()
{
	this.image = null;
	this.image_locked = 0;
	this.image_unusedCount = 0;
}

function CCacheManager(isSlides)
{
	this.isSlides = !!isSlides;
	this.arrayImages = [];
	this.arrayCount = 0;
	this.countValidImage = 1;
}

CCacheManager.prototype.CheckImagesForNeed = function()
{
	for (var i = 0; i < this.arrayCount; ++i)
	{
		if ((this.arrayImages[i].image_locked == 0) && (this.arrayImages[i].image_unusedCount >= this.countValidImage))
		{
			delete this.arrayImages[i].image;
			this.arrayImages.splice(i, 1);
			--i;
			--this.arrayCount;
		}
	}
};

CCacheManager.prototype.UnLock = function(_cache_image)
{
	if (null == _cache_image)
		return;

	_cache_image.image_locked = 0;
	_cache_image.image_unusedCount = 0;
	// затем нужно сбросить ссылку в ноль (_cache_image = null) <- это обязательно !!!!!!!
};

CCacheManager.prototype.Lock = function (_w, _h, _drDocument)
{
	if (_w <= 0) _w = 1;
	if (_h <= 0) _h = 1;
	var backgroundColor = this.isSlides ? "#B0B0B0" : "#FFFFFF";
	if (_drDocument)
	{
		var backColor = _drDocument.m_oWordControl.m_oApi.getPageBackgroundColor();
		backgroundColor = "#" + backColor.R.toString(16) + backColor.G.toString(16) + backColor.B.toString(16);
	}

	for (var i = 0; i < this.arrayCount; ++i)
	{
		if (this.arrayImages[i].image_locked)
			continue;
		var _wI = this.arrayImages[i].image.width;
		var _hI = this.arrayImages[i].image.height;
		if ((_wI == _w) && (_hI == _h))
		{
			this.arrayImages[i].image_locked = 1;
			this.arrayImages[i].image_unusedCount = 0;
			this.arrayImages[i].image.ctx.globalAlpha = 1.0;
			this.arrayImages[i].image.ctx.setTransform(1, 0, 0, 1, 0, 0);
			this.arrayImages[i].image.ctx.fillStyle = backgroundColor;
			this.arrayImages[i].image.ctx.fillRect(0, 0, _w, _h);
			return this.arrayImages[i];
		}
		this.arrayImages[i].image_unusedCount++;
	}
	this.CheckImagesForNeed();
	var index = this.arrayCount;
	this.arrayCount++;

	this.arrayImages[index] = new CCacheImage();
	this.arrayImages[index].image = document.createElement('canvas');
	this.arrayImages[index].image.width = _w;
	this.arrayImages[index].image.height = _h;
	this.arrayImages[index].image.ctx = this.arrayImages[index].image.getContext('2d');
	this.arrayImages[index].image.ctx.globalAlpha = 1.0;
	this.arrayImages[index].image.ctx.setTransform(1, 0, 0, 1, 0, 0);
	this.arrayImages[index].image.ctx.fillStyle = backgroundColor;
	this.arrayImages[index].image.ctx.fillRect(0, 0, _w, _h);
	this.arrayImages[index].image_locked = 1;
	this.arrayImages[index].image_unusedCount = 0;
	return this.arrayImages[index];
};

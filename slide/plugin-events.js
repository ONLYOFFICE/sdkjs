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

/**
 * Event: onChangeCurrentSlide
 * @event Plugin#onChangeCurrentSlide
 * @memberof Plugin
 * @typeofeditors ["CPE"]
 * @alias onChangeCurrentSlide
 * @description The function called when the current slide has changed.
 * @param {number} index - The index of the newly activated slide.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onChangeCurrentSlide.js
 */

/**
 * Event: onSlideShowBegin
 * @event Plugin#onSlideShowBegin
 * @memberof Plugin
 * @typeofeditors ["CPE"]
 * @alias onSlideShowBegin
 * @description The function called when a slide show presentation starts.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onSlideShowBegin.js
 */

/**
 * Event: onSlideShowEnd
 * @event Plugin#onSlideShowEnd
 * @memberof Plugin
 * @typeofeditors ["CPE"]
 * @alias onSlideShowEnd
 * @description The function called when a slide show presentation ends.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onSlideShowEnd.js
 */

/**
 * Event: onSlideShowNextSlide
 * @event Plugin#onSlideShowNextSlide
 * @memberof Plugin
 * @typeofeditors ["CPE"]
 * @alias onSlideShowNextSlide
 * @description The function called after a slide has been changed and displayed during a slide show presentation.
 * Fires before the slide content is actually displayed.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onSlideShowNextSlide.js
 */

/**
 * Event: onSlideShowSlideChanged
 * @event Plugin#onSlideShowSlideChanged
 * @memberof Plugin
 * @typeofeditors ["CPE"]
 * @alias onSlideShowSlideChanged
 * @description The function called when slide changes during a slide show presentation. Provides information about both the current and previous slide.
 * @param {Object} data - The object containing slide information.
 * @param {number} data.slideIndex - The index of the new slide.
 * @param {number} data.previousSlideIndex - The index of the previous slide (-1 if it's the first slide).
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onSlideShowSlideChanged.js
 */

"use strict";

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

/**
 * Event: onClick
 * @event Plugin#onClick
 * @memberof Plugin
 * @alias onClick
 * @description The function called when the user clicks the element.
 * @param {boolean} isSelectionUse - Defines if the selection is used or not.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onClick.js
 */

/**
 * Event: onTargetPositionChanged
 * @event Plugin#onTargetPositionChanged
 * @memberof Plugin
 * @alias onTargetPositionChanged
 * @description The function called when the target position in the editor is changed.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onTargetPositionChanged.js
 */

/**
 * Event: onDocumentContentReady
 * @event Plugin#onDocumentContentReady
 * @memberof Plugin
 * @alias onDocumentContentReady
 * @description The function called when the document is completely loaded.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onDocumentContentReady.js
 */

/**
 * Event: onEnableMouseEvent
 * @event Plugin#onEnableMouseEvent
 * @memberof Plugin
 * @alias onEnableMouseEvent
 * @description The function called to turn the mouse or touchpad events on/off.
 * @param {boolean} isEnabled - Indicates whether the mouse or touchpad is enabled (**true**) or not (**false**).
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onEnableMouseEvent.js
 */

/**
 * Event: onExternalMouseUp
 * @event Plugin#onExternalMouseUp
 * @memberof Plugin
 * @alias onExternalMouseUp
 * @description The function called when the mouse button is released outside the plugin iframe.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onExternalMouseUp.js
 */

/**
 * Event: onChangeRestrictions
 * @event Plugin#onChangeRestrictions
 * @memberof Plugin
 * @alias onChangeRestrictions
 * @description The function called when the restrictions in the editor are changed.
 * @param {number} value - The restrictions value.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onChangeRestrictions.js
 */


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

/**
 * Event: onClick
 * @event Plugin#onClick
 * @memberof Plugin
 * @alias onClick
 * @description The function called when the user clicks on the element.
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
 * @param {boolean} isEnabled - Defines if the mouse or touchpad is enabled (**true**) or not (**false**).
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
 * @description The function called when the restrictions in the editor is changed.
 * @param {number} value - restrictions value.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onChangeRestrictions.js
 */


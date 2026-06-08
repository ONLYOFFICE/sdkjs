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

window.AscInterface = window.AscInterface || {};

AscInterface.updateStructure = function(structure)
{
    var elem = document.getElementById("bookmarksTree");
    elem.innerHTML = "";
    
    if (!structure)
        return;

    var jsonStructure = { children: [], parent: jsonStructure };
    var currentLevel = 0;
    var currentElement = jsonStructure;
    var parent, newElem, item, level;

    function makeFolder(obj) {
        obj.open = false;
        obj.type = Tree.FOLDER;
        obj.selected = true;
    }

    for (var len = structure.length, index = 0; index < len; index++)
    {
        item = structure[index];
        level = item.level;
        if (currentLevel == level)
        {
            // same level - common parent
            parent = currentElement.parent;
        }
        else if ((currentLevel + 1) == level)
        {
            // next level
            parent = currentElement;
        }
        else
        {
            // return to required level
            parent = currentElement;
            while (level < parent.level)
                parent = parent.parent;
            parent = parent.parent;
        }

        newElem = { name: item.description, id: index, children : [], parent: parent, level : level };
        parent.children.push(newElem);
        makeFolder(parent);
        currentLevel = item.level;
        currentElement = newElem;
    }

    var treeElem = new Tree(elem, { navigate: true });
    treeElem.json(jsonStructure.children);

    // subscribe afterwards
    treeElem.on('select', function(node) {
        window.Viewer.navigate(parseInt(node.getAttribute("nodeId")));
    });
};

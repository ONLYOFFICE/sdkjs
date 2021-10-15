/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
            // такой же уровень - общий родитель
            parent = currentElement.parent;
        }
        else if ((currentLevel + 1) == level)
        {
            // следующий уровень
            parent = currentElement;
        }
        else
        {
            // возврат на нужный уровень
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

    // подписываемся после
    treeElem.on('select', function(node) {
        window.Viewer.navigate(parseInt(node.getAttribute("nodeId")));
    });
};

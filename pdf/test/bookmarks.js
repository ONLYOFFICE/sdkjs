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

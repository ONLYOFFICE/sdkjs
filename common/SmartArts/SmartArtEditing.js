/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 * SPDX-License-Identifier: AGPL-3.0-only
 */
(function (window) {
    "use strict";

    const SmartArt = AscFormat.SmartArt;

    SmartArt.prototype.getEditableOutline = function () {
        if (!this.isCorretDataModel() || this.isEmptyLayout()) {
            return null;
        }
        // Read the serialized relations, rather than a cached layout tree after undo.
        const tree = new AscFormat.SmartArtAlgorithm(this);
        const outline = [];
        function visit(node, depth) {
            for (let i = 0; i < node.childs.length; ++i) {
                const child = node.childs[i];
                const body = child.point.getT();
                outline.push({
                    "id": child.getModelId(),
                    "text": body && body.content ? body.content.GetText({}).replace(/\r\n?/g, "\n").replace(/\n$/, "") : "",
                    "depth": depth,
                    "assistant": child.isAsst()
                });
                visit(child, depth + 1);
            }
        }
        visit(tree.dataRoot, 0);
        return outline;
    };

    function displayedNodes(smartart) {
        const displayed = new Map();
        const shapes = smartart.getShapeMap();
        const relations = smartart.getRelationOfContent() || {};
        Object.keys(relations).forEach(function (id) {
            const content = shapes[id].getDocContent && shapes[id].getDocContent();
            const text = content ? content.GetText({}).replace(/\r\n?/g, "\n") : "";
            relations[id].forEach(function (relation) {
                const pointId = relation.point.getModelId();
                displayed.set(pointId, (displayed.get(pointId) || "") + text);
            });
        });
        return displayed;
    }

    SmartArt.prototype.prepareEditableOutline = function (outline, recordHistory) {
        const current = this.getEditableOutline();
        if (!current || !Array.isArray(outline) || !outline.length) {
            return null;
        }
        const known = new Map(current.map(function (node) { return [node["id"], node]; }));
        const seen = new Set();
        for (let i = 0; i < outline.length; ++i) {
            const node = outline[i];
            if (!node || typeof node["text"] !== "string" || !Number.isInteger(node["depth"]) ||
                node["depth"] < 0 || node["depth"] > (i ? outline[i - 1]["depth"] + 1 : 0) ||
                (node["assistant"] !== undefined && typeof node["assistant"] !== "boolean") ||
                (node["id"] !== undefined && (!known.has(node["id"]) || seen.has(node["id"])))) {
                return null;
            }
            if (node["id"]) seen.add(node["id"]);
        }
        const prepare = function () {
            // Prepare a complete replacement before touching the live document.
            const copy = this.copy({cacheImage: false});
            const drawingDocument = (Asc.editor || window.editor).getDrawingDocument();
            copy.parent = this.parent;
            copy.worksheet = this.worksheet;
            copy.drawingObjects = this.drawingObjects;
            const model = copy.getDataModelFromData();
            const points = model.getPtLst();
            const connections = model.getCxnLst();
            const pointMap = points.getPtMap();
            const oldConnections = new Map();
            connections.list.forEach(function (cxn) {
                if (cxn.type === AscFormat.Cxn_type_parOf) oldConnections.set(cxn.destId, cxn);
            });
            const generator = new AscFormat.SmartArtAlgorithm(copy);
            const required = new Map();
            const parents = [model.getMainPoint()];
            const orders = new Map();
            for (let i = connections.list.length - 1; i >= 0; --i) {
                if (connections.list[i].type === AscFormat.Cxn_type_parOf) connections.removeFromLst(i);
            }
            function newPoint(type) {
                const point = new AscFormat.Point();
                point.setModelId(generator.getNewGUID());
                point.setType(type);
                point.setPrSet(new AscFormat.PrSet());
                point.setT(AscFormat.CreateTextBodyFromString("", drawingDocument, point));
                points.addToLst(points.list.length, point);
                return point;
            }
            outline.forEach(function (node) {
                const point = node["id"] ? pointMap[node["id"]] : newPoint(AscFormat.Point_type_node);
                if (!node["id"] || node["text"]) required.set(point.getModelId(), node["text"]);
                point.setType(node["assistant"] ? AscFormat.Point_type_asst : AscFormat.Point_type_node);
                const oldText = known.get(node["id"]);
                if (!oldText || oldText["text"] !== node["text"]) {
                    let body = point.getT();
                    if (!body) {
                        body = AscFormat.CreateTextBodyFromString(node["text"], drawingDocument, point);
                        point.setT(body);
                    } else {
                        const paragraph = body.content.Content[0];
                        const content = AscFormat.CreateDocContentFromString(node["text"], drawingDocument, body);
                        if (paragraph) {
                            content.Content[0].SetPr(paragraph.Pr.Copy());
                            const run = paragraph.Content.find(function (item) { return item.Pr && item.Pr.Copy; });
                            if (run) content.Content[0].Content[0].SetPr(run.Pr.Copy());
                        }
                        body.setContent(content);
                    }
                    if (point.getPrSet()) point.getPrSet().setPhldr(false);
                }
                const parentId = parents[node["depth"]].getModelId();
                let cxn = oldConnections.get(point.getModelId());
                if (!cxn) {
                    cxn = new AscFormat.Cxn();
                    cxn.setModelId(generator.getNewGUID());
                    cxn.setSibTransId(newPoint(AscFormat.Point_type_sibTrans).getModelId());
                    cxn.setParTransId(newPoint(AscFormat.Point_type_parTrans).getModelId());
                }
                cxn.setType(AscFormat.Cxn_type_parOf);
                cxn.setSrcId(parentId);
                cxn.setDestId(point.getModelId());
                cxn.setSrcOrd(orders.get(parentId) || 0);
                cxn.setDestOrd(0);
                orders.set(parentId, cxn.srcOrd + 1);
                connections.addToLst(connections.list.length, cxn);
                parents[node["depth"] + 1] = point;
            });
            copy.smartArtTree = null;
            copy.initSmartArtAlgorithm();
            copy.smartArtTree.checkDataModel();
            copy.generateDrawingPart(true);
            const drawing = copy.getDrawing();
            const originalShapes = this.getShapeMap();
            const replacementShapes = copy.getShapeMap();
            Object.keys(replacementShapes).forEach(function (id) {
                if (originalShapes[id]) replacementShapes[id].setFLocksText(originalShapes[id].fLocksText);
            });
            const displayed = displayedNodes(copy);
            // Some OOXML layouts intentionally cap their shapes or require children.
            // Reject an outline that would silently hide added nodes or existing text.
            if (Array.from(required).some(function (entry) {
                return !displayed.has(entry[0]) || !displayed.get(entry[0]).includes(entry[1]);
            })) return null;
            if (!drawing.spTree.length || drawing.spTree.some(function (shape) {
                const xfrm = shape.spPr && shape.spPr.xfrm;
                return !xfrm || ![xfrm.offX, xfrm.offY, xfrm.extX, xfrm.extY].every(Number.isFinite);
            })) {
                return null;
            }
            return {data: copy.getDataModel(), drawing: drawing};
        };
        if (recordHistory) return prepare.call(this);
        const history = AscCommon.History;
        const table = AscCommon.g_oTableId;
        const historyState = history.TurnOffHistory;
        const tableState = table.m_bTurnOff;
        history.TurnOff();
        table.m_bTurnOff = true;
        try {
            return prepare.call(this);
        } finally {
            history.TurnOffHistory = historyState;
            table.m_bTurnOff = tableState;
        }
    };

    SmartArt.prototype.applyEditableOutline = function (prepared) {
        this.setDataModel(prepared.data);
        this.removeFromSpTreeByPos(0);
        this.addToSpTree(0, prepared.drawing);
        this.setDrawing(prepared.drawing);
        prepared.drawing.setGroup(this);
        prepared.drawing.setWorksheet(this.worksheet);
        prepared.drawing.setDrawingObjects(this.drawingObjects);
        this.smartArtTree = null;
        this.recalcSmartArtConnections();
        this.recalcFitFontSize();
        this.addToRecalculate();
        if (this.SetNeedRecalc) this.SetNeedRecalc(true);
    };

    function selectedSmartArt(api) {
        const controller = api.getGraphicController();
        if (!controller) return null;
        const selection = controller.selection && controller.selection.groupSelection || controller;
        const selected = selection.selectedObjects;
        if (!selected || selected.length !== 1) return null;
        let object = selected[0];
        while (object) {
            if (object instanceof SmartArt) return object;
            object = object.group;
        }
        return null;
    }

    function protectedText(smartart, controller) {
        if (controller.checkSelectedObjectsProtectionText && controller.checkSelectedObjectsProtectionText()) return true;
        function visit(object) {
            return object.isProtectedText && object.isProtectedText() || object.spTree && object.spTree.some(visit);
        }
        return visit(smartart);
    }

    const api = AscCommon.baseEditorsApi.prototype;
    api["asc_getSmartArtOutline"] = function () {
        const smartart = selectedSmartArt(this);
        const nodes = smartart && smartart.getEditableOutline();
        return nodes ? {"id": smartart.GetId(), "nodes": nodes} : null;
    };
    api["asc_setSmartArtOutline"] = async function (id, nodes, expectedNodes) {
        const smartart = selectedSmartArt(this);
        const controller = this.getGraphicController();
        if (!this.canEdit() || this.isPdfViewer || !smartart || smartart.GetId() !== id ||
            (this.collaborativeEditing && this.collaborativeEditing.getGlobalLock()) ||
            (controller.checkSelectedObjectsProtection && controller.checkSelectedObjectsProtection()) ||
            protectedText(smartart, controller)) return false;
        const before = JSON.stringify(smartart.getEditableOutline());
        if (expectedNodes !== undefined && before !== JSON.stringify(expectedNodes)) return false;
        if (!smartart.prepareEditableOutline(nodes)) return false;
        await new Promise((resolve) => {
            AscFonts.FontPickerByCharacter.checkText(nodes.map(function (node) { return node["text"]; }).join(""), this, resolve);
        });
        const apply = (skipLock) => {
            if (!this.canEdit() || this.isPdfViewer || selectedSmartArt(this) !== smartart ||
                before !== JSON.stringify(smartart.getEditableOutline()) ||
                (this.collaborativeEditing && this.collaborativeEditing.getGlobalLock()) ||
                (controller.checkSelectedObjectsProtection && controller.checkSelectedObjectsProtection()) ||
                protectedText(smartart, controller)) return false;
            let applied = false;
            controller.checkSelectedObjectsAndCallback(function () {
                // Build again with history enabled so every new object and relation
                // is serialized for collaborative editing, as well as local undo.
                const prepared = smartart.prepareEditableOutline(nodes, true);
                if (!prepared) return;
                const page = smartart.selectStartPage;
                controller.resetSelection();
                const main = smartart.getMainGroup() || smartart;
                controller.selectObject(main, page);
                if (main !== smartart) {
                    controller.selection.groupSelection = main;
                    main.selectObject(smartart, page);
                }
                smartart.applyEditableOutline(prepared);
                applied = true;
            }, [], false, undefined, undefined, skipLock);
            return applied;
        };
        if (this.getEditorId() === AscCommon.c_oEditorId.Spreadsheet) {
            return new Promise((resolve) => {
                const main = smartart.getMainGroup() || smartart;
                this.checkObjectsLock([main.GetId()], function (unlocked) { resolve(unlocked ? apply(true) : false); });
            });
        }
        return apply(false);
    };
})(window);

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

(function(window, undefined) {
	var openXml = {};

	function SaxParserBase() {
		this.depth = 0;
		this.depthSkip = null;
		this.context = null;
		this.contextStack = [];
	}

	SaxParserBase.prototype.onError = function(msg) {
		throw new Error(msg);
	};
	SaxParserBase.prototype.onStartNode = function(elem, getAttrs, isTagEnd, getStringNode) {
		this.depth++;
		if (!this.isSkip()) {
			var newContext;
			if (this.context.onStartNode) {
				newContext = this.context.onStartNode.call(this.context, elem, getAttrs, EasySAXParser.entityDecode, isTagEnd, getStringNode);
				if (!newContext) {
					this.skip();
				}
			}
			if (!this.isSkip() && !isTagEnd) {
				this.context = newContext ? newContext : this.context;
				this.contextStack.push(this.context);
			}
		}
	};
	SaxParserBase.prototype.onTextNode = function(text) {
		if (this.context && this.context.onTextNode) {
			this.context.onTextNode.call(this.context, text, EasySAXParser.entityDecode);
		}
	};
	SaxParserBase.prototype.onEndNode = function(elem, isTagStart, getStringNode) {
		this.depth--;
		var isSkip = this.isSkip();
		if (isSkip && this.depth <= this.depthSkip) {
			this.depthSkip = null;
		}
		if (!isSkip){
			var prevContext = this.context;
			if(!isTagStart){
				this.contextStack.pop();
				this.context = this.contextStack[this.contextStack.length - 1];
			}
			if (this.context && this.context.onEndNode) {
				this.context.onEndNode.call(this.context, prevContext, elem, EasySAXParser.entityDecode, isTagStart, getStringNode);
			}
		}
	};
	SaxParserBase.prototype.skip = function() {
		this.depthSkip = this.depth - 1;
	};
	SaxParserBase.prototype.isSkip = function() {
		return null !== this.depthSkip
	};
	SaxParserBase.prototype.parse = function(xml, context) {
		var t = this;
		this.context = context;
		var parser = new EasySAXParser({'autoEntity': false});
		parser.on('error', function() {
			t.onError.apply(t, arguments);
		});
		parser.on('startNode', function() {
			t.onStartNode.apply(t, arguments);
		});
		parser.on('textNode', function() {
			t.onTextNode.apply(t, arguments);
		});
		parser.on('endNode', function() {
			t.onEndNode.apply(t, arguments);
		});
		parser.parse(xml);
	};

	openXml.SaxParserBase = SaxParserBase;
	openXml.SaxParserDataTransfer = {};

	function ContentTypes(){
		this.Defaults = {};
		this.Overrides = {};
	}
	ContentTypes.prototype.onStartNode = function(elem, attr, uq, tagend, getStrNode) {
		var attrVals;
		if ('Default' === elem) {
			if (attr()) {
				attrVals = attr();
				this.Defaults[attrVals['Extension']] = attrVals['ContentType'];
			}
		} else if ('Override' === elem) {
			if (attr()) {
				attrVals = attr();
				this.Overrides[attrVals['PartName']] = attrVals['ContentType'];
			}
		}
		return this;
	};
	ContentTypes.prototype.toXml = function(writer) {
		writer.WriteXmlString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
		writer.WriteXmlNodeStart("Types");
		writer.WriteXmlString(" xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"");
		writer.WriteXmlAttributesEnd();
		for (var ext in this.Defaults) {
			if (this.Defaults.hasOwnProperty(ext)) {
				writer.WriteXmlNodeStart("Default");
				writer.WriteXmlAttributeStringEncode("Extension", ext);
				writer.WriteXmlAttributeStringEncode("ContentType", this.Defaults[ext]);
				writer.WriteXmlAttributesEnd(true);
			}
		}
		for (var partName in this.Overrides) {
			if (this.Overrides.hasOwnProperty(partName)) {
				writer.WriteXmlNodeStart("Override");
				writer.WriteXmlAttributeStringEncode("PartName", partName);
				writer.WriteXmlAttributeStringEncode("ContentType", this.Overrides[partName]);
				writer.WriteXmlAttributesEnd(true);
			}
		}
		writer.WriteXmlNodeEnd("Types");
	};
	ContentTypes.prototype.add = function(partName, contentType) {
		var exti = partName.lastIndexOf(".");
		var ext = partName.substring(exti + 1);
		var res = !(this.Overrides[partName] && this.Defaults[ext]);
		if (contentType) {
			this.Overrides[partName] = contentType;
		}
		if (!this.Defaults[ext]) {
			var mime = openXml.GetMimeType(ext);
			this.Defaults[ext] = mime;
		}
		return res;
	};
	function Rels(pkg, part){
		this.pkg = pkg;
		this.part = part;
		this.rels = [];
		this.nextRId = 1;
	}

	Rels.prototype.onStartNode = function(elem, attr, uq, tagend, getStrNode) {
		var attrVals;
		if ('Relationships' === elem) {
		} else if ('Relationship' === elem) {
			if (attr()) {
				attrVals = attr();
				var rId = attrVals["Id"] || "";
				var targetMode = attrVals["TargetMode"] || null;
				var theRel = new openXml.OpenXmlRelationship(this.pkg, this.part, rId, attrVals["Type"],
					attrVals["Target"], targetMode);
				this.rels.push(theRel);
				if (rId.startsWith("rId")) {
					this.nextRId = Math.max(this.nextRId, parseInt(rId.substring("rId".length)) + 1 || 1);
				}
			}
		}
		return this;
	};
	Rels.prototype.toXml = function(writer) {
		writer.WriteXmlString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
		writer.WriteXmlNodeStart("Relationships");
		writer.WriteXmlString(" xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"");
		writer.WriteXmlAttributesEnd();
		this.rels.forEach(function(elem){
			elem.toXml(writer, "Relationship");
		});
		writer.WriteXmlNodeEnd("Relationships");
	};
	Rels.prototype.getNextRId = function() {
		return "rId" + (this.nextRId++);
	};


	/******************************** OpenXmlPackage ********************************/
	function openFromZip(zip, pkg) {
		let ctfBytes = zip.getFile("[Content_Types].xml");
		if (ctfBytes) {
			let ctfText = AscCommon.UTF8ArrayToString(ctfBytes, 0, ctfBytes.length);
			new SaxParserBase().parse(ctfText, pkg.cntTypes);
		}
		zip.files.forEach(function(path){
			if (!path.endsWith("/")) {
				var f2 = path;
				var contentType = null;
				if (path !== "[Content_Types].xml") {
					f2 = "/" + path;
					contentType = pkg.getContentType(f2);
				}
				pkg.parts[f2] = new openXml.OpenXmlPart(pkg, f2, contentType);
			}
		});
	}

	openXml.OpenXmlPackage = function(zip, xmlWriter) {
		this.zip = zip;
		this.xmlWriter = xmlWriter;
		this.parts = {};
		this.cntTypes = new ContentTypes();
		this.fileNameIndexes = {};

		this.openFromZip();
	};
	
	openXml.OpenXmlPackage.prototype.openFromZip = function(){
		openFromZip(this.zip, this);
	};
	openXml.OpenXmlPackage.prototype.removePart = function (uri) {
		var removePart = this.parts[uri];
		if(removePart) {
			delete this.parts[uri];
			this.zip.removeFile(removePart.getUriRelative());
		}
		return removePart;
	};
	openXml.OpenXmlPackage.prototype.generateNextFilename = function (type) {
		if (-1 === type.filename.indexOf("[N]")) {
			return type.filename;
		} else {
			let sIndexKey = type.enumerateType || type.relationType;
			var nextIndex = 1;
			if (!this.fileNameIndexes[sIndexKey]) {
				this.fileNameIndexes[sIndexKey] = nextIndex + 1;
			} else {
				nextIndex = this.fileNameIndexes[sIndexKey]++;
			}
			return type.filename.replace(/\[N\]/g, nextIndex.toString());
		}
	};
	openXml.OpenXmlPackage.prototype.generateTargetByType = function (type) {
		if (type.dir) {
			return type.dir + "/" + this.generateNextFilename(type);
		} else {
			return this.generateNextFilename(type);
		}
	};
	openXml.OpenXmlPackage.prototype.generateUriByType = function (target, base) {
		var baseSplit = base.split('/');
		baseSplit.pop();
		var targetSplit = target.split('/');
		for (var i = 0; i < targetSplit.length; ++i) {
			if ('..' === targetSplit[i]) {
				baseSplit.pop();
			} else {
				baseSplit.push(targetSplit[i]);
			}
		}
		return baseSplit.join('/');
	};
	openXml.OpenXmlPackage.prototype.addPartWithoutRels = function (uri, contentType) {
		//add part
		var newPart = new openXml.OpenXmlPart(this, uri, contentType);
		this.parts[uri] = newPart;
		// update [Content_Types].xml
		var changed = this.cntTypes.add(uri, contentType);
		if(changed) {
			this.zip.removeFile("[Content_Types].xml");
			this.zip.addFile("[Content_Types].xml", this.getXmlBytes(this.getRootPart(), this.cntTypes, this.xmlWriter));
		}
		return newPart;
	};
	openXml.OpenXmlPackage.prototype.addPart = function (type) {
		return this.getRootPart().addPart(type);
	};
	openXml.OpenXmlPackage.prototype.addRelationship = function (relationshipType, target, targetMode) {
		return this.getRootPart().addRelationship(relationshipType, target, targetMode);
	};

	openXml.OpenXmlPackage.prototype.getParts = function() {
		var parts = [];
		for (var part in this.parts) {
			if (this.parts[part].contentType !== openXml.Types.relationships.contentType && part !== "[Content_Types].xml") {
				parts.push(this.parts[part]);
			}
		}
		return parts;
	};

	openXml.OpenXmlPackage.prototype.getRootPart = function() {
		return new openXml.OpenXmlPart(this, "/", openXml.Types.relationships.contentType);
	};
	openXml.OpenXmlPackage.prototype.getRels = function() {
		return this.getRootPart().getRels();
	};
	openXml.OpenXmlPackage.prototype.getRelationships = function() {
		return this.getRootPart().getRelationships();
	};

	openXml.OpenXmlPackage.prototype.getRelationship = function(rId) {
		return this.getRootPart().getRelationship(rId);
	};

	openXml.OpenXmlPackage.prototype.getRelationshipsByRelationshipType = function(relationshipType) {
		return this.getRootPart().getRelationshipsByRelationshipType(relationshipType);
	};

	openXml.OpenXmlPackage.prototype.getPartsByRelationshipType = function(relationshipType) {
		return this.getRootPart().getPartsByRelationshipType(relationshipType);
	};

	openXml.OpenXmlPackage.prototype.getPartByRelationshipType = function(relationshipType) {
		return this.getRootPart().getPartByRelationshipType(relationshipType);
	};

	openXml.OpenXmlPackage.prototype.getRelationshipsByContentType = function(contentType) {
		return this.getRootPart().getRelationshipsByContentType(contentType);
	};

	openXml.OpenXmlPackage.prototype.getPartsByContentType = function(contentType) {
		return this.getRootPart().getPartsByContentType(contentType);
	};

	openXml.OpenXmlPackage.prototype.getRelationshipById = function(rId) {
		return this.getRootPart().getRelationshipById(rId);
	};

	openXml.OpenXmlPackage.prototype.getPartById = function(rId) {
		return this.getRootPart().getPartById(rId);
	};

	openXml.OpenXmlPackage.prototype.getPartByUri = function(uri) {
		var part = this.parts[uri];
		return part;
	};

	openXml.OpenXmlPackage.prototype.getContentType = function(uri) {
		var ct = this.cntTypes.Overrides[uri];
		if (!ct) {
			var exti = uri.lastIndexOf(".");
			var ext = uri.substring(exti + 1);
			ct = this.cntTypes.Defaults[ext];
		}
		return ct;
	};
	openXml.OpenXmlPackage.prototype.getXmlBytes = function(part, data, writer) {
		var oldPart = writer.context.part;
		writer.context.part = part;
		var oldPos = writer.GetCurPosition();
		data.toXml(writer);
		var pos = writer.GetCurPosition();
		var res = writer.GetDataUint8(oldPos, pos - oldPos);

		writer.Seek(oldPos);
		writer.context.part = oldPart;
		return res;
	};

	/*********** OpenXmlPart ***********/

	openXml.OpenXmlPart = function(pkg, uri, contentType) {
		this.pkg = pkg;      // reference to the parent package
		this.uri = uri;      // the part is also indexed by uri in the package
		this.contentType = contentType;
	};

	openXml.OpenXmlPart.prototype.getUriRelative = function() {
		return this.uri.substring(1);
	};
	openXml.OpenXmlPart.prototype.getDocumentContent = function(type) {
		type = type || "string";
		var data = this.pkg.zip.getFile(this.getUriRelative());
		if (!data) {
			data = new Uint8Array(0);
		}
		if ("string" === type) {
			return AscCommon.UTF8ArrayToString(data, 0, data.length);
		} else {
			return data;
		}
	};
	openXml.OpenXmlPart.prototype.addPart = function (type) {
		var target = this.pkg.generateTargetByType(type);
		var uri = this.pkg.generateUriByType(target, this.uri);
		var newPart = this.pkg.addPartWithoutRels(uri, type.contentType);
		//update rels
		var rId = this.addRelationship(type.relationType, target);
		return {part: newPart, rId: rId};
	};
	openXml.OpenXmlPart.prototype.addPartWithoutRels = function (type) {
		var target = this.pkg.generateTargetByType(type);
		var uri = this.pkg.generateUriByType(target, this.uri);
		return this.pkg.addPartWithoutRels(uri, type.contentType);
	};
	openXml.OpenXmlPart.prototype.setData = function (data) {
		this.pkg.zip.addFile(this.getUriRelative(), data);
	};
	openXml.OpenXmlPart.prototype.setDataXml = function (xmlObj, writer) {
		writer.context.clearCurrentPartDataMaps();
		var data = this.pkg.getXmlBytes(this, xmlObj, writer);
		this.pkg.zip.addFile(this.getUriRelative(), data);
	};
	openXml.OpenXmlPart.prototype.addRelationship = function (relationshipType, target, targetMode) {
		var relsFilename = getRelsPartUriOfPart(this);
		var rels = this.getRels();
		var rId = rels.getNextRId();
		var newRel = new openXml.OpenXmlRelationship(rels.pkg, rels.part, rId, relationshipType, target, targetMode);
		rels.rels.push(newRel);
		this.pkg.removePart(relsFilename);
		var relsPart = this.pkg.addPartWithoutRels(relsFilename, null);
		relsPart.setData(this.pkg.getXmlBytes(relsPart, rels, this.pkg.xmlWriter));
		return rId;
	};

	function getRelsPartUriOfPart(part) {
		var uri = part.uri;
		var lastSlash = uri.lastIndexOf('/');
		var partFileName = uri.substring(lastSlash + 1);
		var relsFileName = uri.substring(0, lastSlash) + "/_rels/" + partFileName + ".rels";
		return relsFileName;
	}

	function getPartUriOfRelsPart(part) {
		var uri = part.uri;
		var lastSlash = uri.lastIndexOf('/');
		var partFileName = uri.substring(lastSlash + 1, uri.length - '.rels'.length);
		var relsFileName = uri.substring(0, uri.lastIndexOf('/', lastSlash - 1) + 1) + partFileName;
		return relsFileName;
	}

	function getRelsPartOfPart(part) {
		var relsFileName = getRelsPartUriOfPart(part);
		var relsPart = part.pkg.getPartByUri(relsFileName);
		return relsPart;
	}

	openXml.OpenXmlPart.prototype.getRels = function() {
		var relsPackage = new Rels(null, this);
		var relsPart = getRelsPartOfPart(this);
		if(relsPart) {
			new SaxParserBase().parse(relsPart.getDocumentContent(), relsPackage);
		}
		return relsPackage;
	}
	openXml.OpenXmlPart.prototype.getRelationships = function() {
		return this.getRels().rels;
	}

	openXml.OpenXmlPart.prototype.getRelationship = function(rId) {
		var rels = this.getRelationships();
		for (var i = 0; i < rels.length; ++i) {
			var rel = rels[i];
			if (rel.relationshipId == rId) {
				return rel;
			}
		}
		return null;
	}

	// returns all related parts of the source part
	openXml.OpenXmlPart.prototype.getParts = function() {
		var parts = [];
		var rels = this.getRelationships();
		for (var i = 0; i < rels.length; ++i) {
			var part = this.pkg.getPartByUri(rels[i].targetFullName);
			parts.push(part);
		}
		return parts;
	}

	openXml.OpenXmlPart.prototype.getRelationshipsByRelationshipType = function(relationshipType) {
		var rels = this.getRelationships();
		return rels.filter(function (rel) {
			return openXml.IsEqualRelationshipType(rel.relationshipType, relationshipType);
		});
	}

	// returns all related parts of the source part with the given relationship type
	openXml.OpenXmlPart.prototype.getPartsByRelationshipType = function(relationshipType) {
		var parts = [];
		var rels = this.getRelationshipsByRelationshipType(relationshipType);
		for (var i = 0; i < rels.length; ++i) {
			var part = this.pkg.getPartByUri(rels[i].targetFullName);
			parts.push(part);
		}
		return parts;
	}

	openXml.OpenXmlPart.prototype.getPartByRelationshipType = function(relationshipType) {
		var parts = this.getPartsByRelationshipType(relationshipType);
		if (parts.length < 1) {
			return null;
		}
		return parts[0];
	}

	openXml.OpenXmlPart.prototype.getRelationshipsByContentType = function(contentType) {
		var rels = this.getRelationships();
		return rels.filter(function (rel) {
			return this.getContentType(rel.targetFullName) === contentType;
		});
	}

	openXml.OpenXmlPart.prototype.getPartsByContentType = function(contentType) {
		var parts = [];
		var rels = this.getRelationshipsByContentType(contentType);
		for (var i = 0; i < rels.length; ++i) {
			var part = this.pkg.getPartByUri(rels[i].targetFullName);
			parts.push(part);
		}
		return parts;
	}

	openXml.OpenXmlPart.prototype.getRelationshipById = function(relationshipId) {
		return this.getRelationship(relationshipId);
	}

	openXml.OpenXmlPart.prototype.getPartById = function(relationshipId) {
		var rel = this.getRelationshipById(relationshipId);
		if (rel) {
			var part = this.pkg.getPartByUri(rel.targetFullName);
			return part;
		}
		return null;
	}

	/******************************** OpenXmlRelationship ********************************/

	openXml.OpenXmlRelationship = function(pkg, part, relationshipId, relationshipType, target, targetMode) {
		this.fromPkg = pkg;        // if from a part, this will be null
		this.fromPart = part;      // if from a package, this will be null;
		this.relationshipId = relationshipId;
		this.relationshipType = relationshipType;
		this.target = target;
		this.targetMode = targetMode;
		if (!targetMode) {
			this.targetMode = "Internal";
		}

		var workingTarget = target;
		var workingCurrentPath;
		if (this.fromPkg) {
			workingCurrentPath = "/";
		}
		if (this.fromPart) {
			var slashIndex = this.fromPart.uri.lastIndexOf('/');
			if (slashIndex === -1) {
				workingCurrentPath = "/";
			} else {
				workingCurrentPath = this.fromPart.uri.substring(0, slashIndex) + "/";
			}
		}
		if (targetMode === openXml.TargetMode.external) {
			this.targetFullName = this.target;
			return;
		}
		while (workingTarget.startsWith('../')) {
			if (workingCurrentPath.endsWith('/')) {
				workingCurrentPath = workingCurrentPath.substring(0, workingCurrentPath.length - 1);
			}
			var indexOfLastSlash = workingCurrentPath.lastIndexOf('/');
			if (indexOfLastSlash === -1) {
				throw "internal error when processing relationships";
			}
			workingCurrentPath = workingCurrentPath.substring(0, indexOfLastSlash + 1);
			workingTarget = workingTarget.substring(3);
		}

		if (workingTarget.startsWith("/")) {
			this.targetFullName = workingTarget;
		} else {
			this.targetFullName = workingCurrentPath + workingTarget;
		}
	}
	openXml.OpenXmlRelationship.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		if (this.relationshipId) {
			writer.WriteXmlAttributeStringEncode("Id", this.relationshipId);
		}
		if (this.relationshipType) {
			writer.WriteXmlAttributeString("Type", this.relationshipType);
		}
		if (this.target) {
			writer.WriteXmlAttributeStringEncode("Target", this.target);
		}
		if (this.targetMode && this.targetMode !== openXml.TargetMode.internal) {
			writer.WriteXmlAttributeString("TargetMode", this.targetMode);
		}
		writer.WriteXmlAttributesEnd(true);
	};
	openXml.OpenXmlRelationship.prototype.getFullPath = function() {
		return this.targetFullName;
	};

	openXml.MimeTypes = {
		"bmp": "image/bmp",
		"gif": "image/gif",
		"png": "image/png",
		"tif": "image/tiff",
		"tiff": "image/tiff",
		"jpeg": "image/jpeg",
		"jpg": "image/jpeg",
		"jpe": "image/jpeg",
		"jfif": "image/jpeg",
		"rels": "application/vnd.openxmlformats-package.relationships+xml",
		"bin": "application/vnd.openxmlformats-officedocument.oleObject",
		"xml": "application/xml",
		"emf": "image/x-emf",
		"emz": "image/x-emz",
		"wmf": "image/x-wmf",
		"svg": "image/svg+xml",
		"svm": "image/svm",
		"wdp": "image/vnd.ms-photo",
		"wav": "audio/wav",
		"wma": "audio/x-wma",
		"m4a": "audio/unknown",
		"mp3": "audio/mpeg",
		"mp4": "video/unknown",
		"mov": "video/unknown",
		"m4v": "video/unknown",
		"mkv": "video/unknown",
		"avi": "video/avi",
		"flv": "video/x-flv",
		"wmv": "video/x-wmv",
		"webm": "video/webm",
		"xls": "application/vnd.ms-excel",
		"xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
		"xlsb": "application/vnd.ms-excel.sheet.binary.macroEnabled.12",
		"xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		"ppt": "application/vnd.ms-powerpoint",
		"pptm": "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
		"pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
		"sldm": "application/vnd.ms-powerpoint.slide.macroEnabled.12",
		"sldx": "application/vnd.openxmlformats-officedocument.presentationml.slide",
		"doc": "application/msword",
		"docm": "application/vnd.ms-word.document.macroEnabled.12",
		"docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
		"vml": "application/vnd.openxmlformats-officedocument.vmlDrawing",
		"vsd": "application/vnd.visio",
		"vsdx": "application/vnd.ms-visio.drawing"
	};
	openXml.GetMimeType = function(ext) {
		return openXml.MimeTypes[ext] || "application/octet-stream";
	};
	/******************************** OpenXmlRelationship ********************************/
	openXml.Types = {
		calculationChain: {dir: "", filename: "calcChain.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"},
		cellMetadata: {dir: "", filename: "cellMetadata.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"},
		chart: {dir: "../charts", filename: "chart[N].xml", contentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"},
		chartEx: {dir: "../charts", filename: "chartEx[N].xml", contentType: "application/vnd.ms-office.chartex+xml", relationType: "http://schemas.microsoft.com/office/2014/relationships/chartEx"},
		chartWord: {dir: "charts", filename: "chart[N].xml", contentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"},
		chartColorStyle: {dir: "", filename: "color[N].xml", contentType: "application/vnd.ms-office.chartcolorstyle+xml", relationType: "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"},
		chartExColorStyle: {dir: "", filename: "colors[N].xml", contentType: "application/vnd.ms-office.chartcolorstyle+xml", relationType: "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"},
		chartDrawing: {dir: "../drawings", filename: "drawing[N].xml", contentType: "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes", enumerateType: "drawings/drawing"},
		chartsheet: {dir: "chartsheets", filename: "sheet[N].xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"},
		chartStyle: {dir: "", filename: "style[N].xml", contentType: "application/vnd.ms-office.chartstyle+xml", relationType: "http://schemas.microsoft.com/office/2011/relationships/chartStyle"},
		commentAuthors: {dir: "", filename: "commentAuthors.xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"},
		connections: {dir: "", filename: "connections.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections"},
		coreFileProperties: {dir: "docProps", filename: "core.xml", contentType: "application/vnd.openxmlformats-package.core-properties+xml", relationType: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"},
		customFileProperties: {dir: "docProps", filename: "custom.xml", contentType: "application/vnd.openxmlformats-officedocument.custom-properties+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"},
		customization: {dir: "", filename: "customization.xml", contentType: "application/vnd.ms-word.keyMapCustomizations+xml", relationType: "http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations"},
		customProperty: {dir: "", filename: "customProperty.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty"},
		customXmlProperties: {dir: "", filename: "customXmlProperties.xml", contentType: "application/vnd.openxmlformats-officedocument.customXmlProperties+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"},
		diagramColors: {dir: "diagrams", filename: "colors[N].xml", contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors"},
		diagramData: {dir: "diagrams", filename: "data[N].xml", contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"},
		diagramLayoutDefinition: {dir: "diagrams", filename: "layout[N].xml", contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"},
		diagramPersistLayout: {dir: "diagrams", filename: "drawing[N].xml", contentType: "application/vnd.ms-office.drawingml.diagramDrawing+xml", relationType: "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"},
		diagramStyle: {dir: "diagrams", filename: "quickStyle[N].xml", contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle"},
		dialogsheet: {dir: "", filename: "dialogsheet.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet"},
		digitalSignatureOrigin: {dir: "", filename: "digitalSignatureOrigin.xml", contentType: "application/vnd.openxmlformats-package.digital-signature-origin", relationType: "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"},
		documentSettings: {dir: "", filename: "settings.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"},
		drawings: {dir: "../drawings", filename: "drawing[N].xml", contentType: "application/vnd.openxmlformats-officedocument.drawing+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", enumerateType: "drawings/drawing"},
		endnotes: {dir: "", filename: "endnotes.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"},
		excelAttachedToolbars: {dir: "", filename: "excelAttachedToolbars.xml", contentType: "application/vnd.ms-excel.attachedToolbars", relationType: "http://schemas.microsoft.com/office/2006/relationships/attachedToolbars"},
		extendedFileProperties: {dir: "docProps", filename: "app.xml", contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"},
		externalWorkbook: {dir: "externalLinks", filename: "externalLink[N].xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"},
		fontData: {dir: "", filename: "fontData.xml", contentType: "application/x-fontdata"},
		fontTable: {dir: "", filename: "fontTable.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"},
		footer: {dir: "", filename: "footer[N].xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"},
		footnotes: {dir: "", filename: "footnotes.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"},
		gif: {dir: "", filename: "gif.xml", contentType: "image/gif"},
		glossaryDocument: {dir: "glossary", filename: "document.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument"},
		handoutMaster: {dir: "", filename: "handoutMaster.xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster"},
		header: {dir: "", filename: "header[N].xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"},
		jpeg: {dir: "", filename: "jpeg.xml", contentType: "image/jpeg"},
		mainDocument: {dir: "word", filename: "document.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"},
		notesMaster: {dir: "notesMasters", filename: "notesMaster[N].xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"},
		notesSlide: {dir: "notesSlides", filename: "notesSlide[N].xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"},
		numbering: {dir: "", filename: "numbering.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"},
		pict: {dir: "", filename: "pict.xml", contentType: "image/pict"},
		pivotTable: {dir: "../pivotTables", filename: "pivotTable[N].xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"},
		pivotTableCacheDefinition: {dir: "../pivotCache", filename: "pivotCacheDefinition[N].xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"},
		pivotTableCacheDefinitionWorkbook: {dir: "pivotCache", filename: "pivotCacheDefinition[N].xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"},
		pivotTableCacheRecords: {dir: "", filename: "pivotCacheRecords[N].xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords"},
		package: {dir: "../embeddings", filename: "Embedding[N].xlsx", contentType: "image/png", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"},
		png: {dir: "", filename: "png.xml", contentType: "image/png"},
		presentation: {dir: "ppt", filename: "presentation.xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"},
		presentationProperties: {dir: "", filename: "presProps.xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"},
		presentationTemplate: {dir: "", filename: "presentationTemplate.xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"},
		queryTable: {dir: "../queryTables", filename: "queryTable[N].xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable"},
		relationships: {dir: "_rels", filename: ".rels", contentType: "application/vnd.openxmlformats-package.relationships+xml"},
		ribbonAndBackstageCustomizations: {dir: "", filename: "ribbonAndBackstageCustomizations.xml", contentType: "http://schemas.microsoft.com/office/2009/07/customui", relationType: "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"},
		sharedStringTable: {dir: "", filename: "sharedStrings.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"},
		singleCellTable: {dir: "", filename: "singleCellTable.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells"},
		slicerCache: {dir: "slicerCaches", filename: "slicerCache[N].xml", contentType: "application/vnd.ms-excel.slicerCache+xml", relationType: "http://schemas.microsoft.com/office/2007/relationships/slicerCache"},
		slicers: {dir: "../slicers", filename: "slicer[N].xml", contentType: "application/vnd.ms-excel.slicer+xml", relationType: "http://schemas.microsoft.com/office/2007/relationships/slicer"},
		slide: {dir: "slides", filename: "slide[N].xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"},
		slideComments: {dir: "comments", filename: "slideComments.xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.comments+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"},
		slideLayout: {dir: "slideLayouts", filename: "slideLayout[N].xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"},
		slideMaster: {dir: "slideMasters", filename: "slideMaster[N].xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"},
		slideShow: {dir: "", filename: "slideShow.xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"},
		slideSyncData: {dir: "", filename: "slideSyncData.xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.slideUpdateInfo+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideUpdateInfo"},
		styles: {dir: "", filename: "styles.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"},
		tableDefinition: {dir: "../tables", filename: "table[N].xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"},
		tableStyles: {dir: "", filename: "tableStyles.xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"},
		theme: {dir: "theme", filename: "theme[N].xml", contentType: "application/vnd.openxmlformats-officedocument.theme+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"},
		themeOverride: {dir: "../theme", filename: "themeOverride[N].xml", contentType: "application/vnd.openxmlformats-officedocument.themeOverride+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride"},
		tiff: {dir: "", filename: "tiff.xml", contentType: "image/tiff"},
		trueTypeFont: {dir: "", filename: "trueTypeFont.xml", contentType: "application/x-font-ttf"},
		userDefinedTags: {dir: "", filename: "userDefinedTags.xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.tags+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"},
		viewProperties: {dir: "", filename: "viewProps.xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"},
		vmlDrawing: {dir: "../drawings", filename: "vmlDrawing[N].vml", contentType: "application/vnd.openxmlformats-officedocument.vmlDrawing", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"},
		volatileDependencies: {dir: "", filename: "volatileDependencies.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies"},
		webSettings: {dir: "", filename: "webSettings.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"},
		wordAttachedToolbars: {dir: "", filename: "wordAttachedToolbars.xml", contentType: "application/vnd.ms-word.attachedToolbars", relationType: "http://schemas.microsoft.com/office/2006/relationships/attachedToolbars"},
		wordComments: {dir: "", filename: "comments.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"},
		wordCommentsExtended: {dir: "", filename: "commentsExtended.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml", relationType: "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"},
		wordCommentsExtensible: {dir: "", filename: "commentsExtensible.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml", relationType: "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"},
		wordCommentsIds: {dir: "", filename: "commentsIds.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml", relationType: "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"},
		wordPackage: {dir: "embeddings", filename: "Embedding[N].xlsx", contentType: "image/png", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"},
		wordPeople: {dir: "", filename: "people.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml", relationType: "http://schemas.microsoft.com/office/2011/relationships/people"},
		wordprocessingTemplate: {dir: "", filename: "wordprocessingTemplate.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"},
		workbook: {dir: "xl", filename: "workbook.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"},
		workbookRevisionHeader: {dir: "", filename: "workbookRevisionHeader.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders"},
		workbookRevisionLog: {dir: "", filename: "workbookRevisionLog.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog"},
		workbookStyles: {dir: "", filename: "styles.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"},
		workbookTemplate: {dir: "", filename: "workbookTemplate.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"},
		workbookUserData: {dir: "", filename: "workbookUserData.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.userNames+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames"},
		worksheet: {dir: "worksheets", filename: "sheet[N].xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"},
		worksheetComments: {dir: "..", filename: "comments[N].xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"},
		worksheetSortMap: {dir: "", filename: "worksheetSortMap.xml", contentType: "application/vnd.ms-excel.wsSortMap+xml", relationType: "http://schemas.microsoft.com/office/2006/relationships/wsSortMap"},
		xmlSignature: {dir: "", filename: "xmlSignature.xml", contentType: "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml", relationType: "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"},
		hyperlink: {dir: "", filename: "", contentType: "", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"},
		metadata: {dir: "xl", filename: "metadata.xml", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"},

		threadedComment: {dir: "../threadedComments", filename: "threadedComment[N].xml", contentType: "application/vnd.ms-excel.threadedcomments+xml", relationType: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"},
		person: {dir: "../persons", filename: "person.xml", contentType: "application/vnd.ms-excel.person+xml", relationType: "http://schemas.microsoft.com/office/2017/10/relationships/person"},
		ctrlProp: {dir: "../ctrlProps", filename: "ctrlProp[N].xml", contentType: "application/vnd.ms-excel.controlproperties+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp"},
		namedSheetViews: {dir: "../namedSheetViews", filename: "namedSheetView[N].xml", contentType: "application/vnd.ms-excel.namedsheetviews+xml", relationType: "http://schemas.microsoft.com/office/2019/04/relationships/namedSheetView"},
		workbookComment: {dir: "", filename: "workbookComments.bin", contentType: "application/octet-stream", relationType: "http://schemas.onlyoffice.com/workbookComments"},
		timelines: {dir: "../timelines", filename: "timeline[N].xml", contentType: "application/vnd.ms-excel.timeline+xml", relationType: "http://schemas.microsoft.com/office/2011/relationships/timeline"},
		timelineCaches: {dir: "timelineCaches", filename: "timelineCache[N].xml", contentType: "application/vnd.ms-excel.timelineCache+xml", relationType: "http://schemas.microsoft.com/office/2011/relationships/timelineCache"},

		jsaProject: {dir: "", filename: "jsaProject.bin", contentType: "application/octet-stream", relationType: "http://schemas.onlyoffice.com/jsaProject"},
		vbaProject: {dir: "", filename: "vbaProject.bin", contentType: "application/octet-stream", relationType: "http://schemas.microsoft.com/office/2006/relationships/vbaProject"},


		customXml: {dir: "../customXml", filename: "item[N].xml", /*contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",*/ relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"},
		customXmlProps: {dir: "", filename: "itemProps[N].xml", contentType: "application/vnd.openxmlformats-officedocument.customXmlProperties+xml", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"},

		//todo
		image: {dir: "../media", filename: "image[N].", contentType: "image/jpeg", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"},
		imageWord: {dir: "media", filename: "image[N].", contentType: "image/jpeg", relationType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"},


		//onlyf
		oformMain: {dir: "oform", filename: "main.xml", contentType: "application/vnd.openxmlformats-package.onlyf+xml", relationType: "https://schemas.onlyoffice.com/relationships/oform-main"},
		oformDefaultUserMaster: {dir: "oform/userMasters", filename: "default.xml", contentType: "application/vnd.openxmlformats-package.onlyf-default-userMaster+xml", relationType: "https://schemas.onlyoffice.com/relationships/oform-default-userMaster"},
		oformUserMaster: {dir: "oform/userMasters", filename: "userMaster[N].xml", contentType: "application/vnd.openxmlformats-package.onlyf-userMaster+xml", relationType: "https://schemas.onlyoffice.com/relationships/oform-userMaster"},
		oformUser: {dir: "oform/users", filename: "user[N].xml", /*contentType: "application/vnd.openxmlformats-package.onlyf-user+xml",*/ relationType: "https://schemas.onlyoffice.com/relationships/oform-user"},
		oformField: {dir: "oform/fields", filename: "field[N].xml", /*contentType: "application/vnd.openxmlformats-package.onlyf-field+xml",*/ relationType: "https://schemas.onlyoffice.com/relationships/oform-field"},
		oformFieldMaster: {dir: "oform/fieldMasters", filename: "fieldMaster[N].xml", contentType: "application/vnd.openxmlformats-package.onlyf-fieldMaster+xml", relationType: "https://schemas.onlyoffice.com/relationships/oform-fieldMaster"},

		//visio
		visioDocument: {dir: "visio", filename: "document.xml", contentType: "application/vnd.ms-visio.drawing.main+xml", relationType: "http://schemas.microsoft.com/visio/2010/relationships/document"},
		visioDocumentWindows: {dir: "", filename: "windows.xml", contentType: "application/vnd.ms-visio.windows+xml", relationType: "http://schemas.microsoft.com/visio/2010/relationships/windows"},
		pages: {dir: "pages", filename: "pages.xml", contentType: "application/vnd.ms-visio.pages+xml", relationType: "http://schemas.microsoft.com/visio/2010/relationships/pages"},
		page: {dir: "", filename: "page[N].xml", contentType: "application/vnd.ms-visio.page+xml", relationType: "http://schemas.microsoft.com/visio/2010/relationships/page"},
		masters: {dir: "masters", filename: "masters.xml", contentType: "application/vnd.ms-visio.masters+xml", relationType: "http://schemas.microsoft.com/visio/2010/relationships/masters"},
		master: {dir: "", filename: "master[N].xml", contentType: "application/vnd.ms-visio.master+xml", relationType: "http://schemas.microsoft.com/visio/2010/relationships/master"},
		masterFromPage: {dir: "../masters", filename: "master[N].xml", contentType: "application/vnd.ms-visio.master+xml", relationType: "http://schemas.microsoft.com/visio/2010/relationships/master"},
		thumbnail: {dir: "docProps", filename: "thumbnail.emf", contentType: "image/x-emf", relationType: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"},
		visioComments: {dir: "", filename: "comments.xml", contentType: "application/vnd.ms-visio.comments+xml", relationType: "http://schemas.microsoft.com/visio/2010/relationships/comments"},
		visioExtensions: {dir: "", filename: "extensions.xml", contentType: "application/vnd.ms-visio.extensions+xml", relationType: "http://schemas.microsoft.com/visio/2010/relationships/extensions"},
		visioDataConnections: {dir: "data", filename: "connections.xml", contentType: "application/vnd.ms-visio.connections+xml", relationType : "http://schemas.microsoft.com/visio/2010/relationships/connections"},
		visioDataRecordSets : {dir: "data", filename: "recordsets.xml", contentType: "application/vnd.ms-visio.recordsets+xml", relationType: "http://schemas.microsoft.com/visio/2010/relationships/recordsets"},
		solutions: {dir: "solutions", filename: "solutions.xml", contentType: "application/vnd.ms-visio.solutions+xml", relationType: "http://schemas.microsoft.com/visio/2010/relationships/solutions"},
		solution: {dir: "", filename: "solution[N].xml", contentType: "application/vnd.ms-visio.solution+xml", relationType: "http://schemas.microsoft.com/visio/2010/relationships/solution"},
		validation : {dir: "", filename: "validation.xml", contentType: "application/vnd.ms-visio.validation+xml", relationType: "http://schemas.microsoft.com/visio/2010/relationships/validation"}
	};
	openXml.TargetMode = {
		internal: "Internal",
		external: "External"
	};
	openXml.IsEqualRelationshipType = function(relationshipType1, relationshipType2) {
		//https://github.com/ONLYOFFICE/core/blob/7a822494aabb1edce441a12e44aa05c3a6501766/OOXML/DocxFormat/FileType.h#L95
		//RelationType
		//http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument
		//http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument
		//is valid and equal so compare tail

		//docs: If either or both of the arguments are negative or NaN, the substring() method treats them as if they were 0.
		const tail1 = relationshipType1.substring(relationshipType1.lastIndexOf("/") + 1);
		const tail2 = relationshipType2.substring(relationshipType2.lastIndexOf("/") + 1);
		return tail1 === tail2;
	};
	//----------------------------------------------------------export----------------------------------------------------
	var prot;
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon']['openXml'] = window['AscCommon'].openXml = openXml;

	prot = openXml;
	prot['GetMimeType'] = prot.GetMimeType;
}(window));

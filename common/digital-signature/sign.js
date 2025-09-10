/*
 * (c) Copyright Ascensio System SIA 2010-2025
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
(function () {
	const CRYPTO_DIGEST_ALGORITHM_TYPE = {
		SHA256: "SHA-256"
	};
	function getCryptoDigestFromOOXMLDigest(ooxmlType) {
		switch (ooxmlType) {
			case AscFormat.OOXML_DIGEST_ALGORITHM_TYPE.SHA1:
				return "SHA-1";
			case AscFormat.OOXML_DIGEST_ALGORITHM_TYPE.SHA256:
				return "SHA-256";
			case AscFormat.OOXML_DIGEST_ALGORITHM_TYPE.SHA384:
				return "SHA-384";
			case AscFormat.OOXML_DIGEST_ALGORITHM_TYPE.SHA512:
				return "SHA-512";
				default:
					return null;
		}
	}
	function CX509Cert() {
	}

	CX509Cert.prototype.getDigestAlgorithm = function () {
		return CRYPTO_DIGEST_ALGORITHM_TYPE.SHA256;
	};



	function getBase64FromBuffer(arrayBuffer) {
		const intArray = new Uint8Array(arrayBuffer);
		let arrayResult = [];
		for (let i = 0; i < intArray.length; i += 1) {
			arrayResult.push(String.fromCharCode(intArray[i]));
		}
		return window.btoa(arrayResult.join(""));
	}
	function isNeedSign(filePath) {
		return !(filePath.indexOf("_xmlsignatures") === 0 ||
			filePath.indexOf("docProps") === 0 ||
			filePath.indexOf("[Content_Types].xml") === 0 ||
			filePath.indexOf("[trash]") === 0);
	}
	function isRels(filePath) {
		const lastIndex = filePath.lastIndexOf(".");
		if (lastIndex !== -1) {
			return filePath.slice(lastIndex + 1) === "rels";
		}
		return false;
	}

	function CDigitalSigner(crypto, cert, zip) {
		this.crypto = crypto;
		this.openXml = new AscCommon.openXml.OpenXmlPackage(zip, null);
		this.cert = cert || new CX509Cert()/*todo temp*/;
		this.date = new Date();
		this.isInit = false;
		this.guid = "";
		this.validImg = "";
		this.invalidImg = "";
	}
	CDigitalSigner.prototype.init = function () {
		const oThis = this;
		return new Promise(function (resolve) {
			if (oThis.isInit) {
				resolve();
			} else {
				oThis.getCertInfo().then(function (cert) {
					oThis.cert = cert;
					oThis.isInit = true;
					resolve();
				});
			}
		});
	};
	CDigitalSigner.prototype.setGuid = function (pr) {
		this.guid = pr;
	};
	CDigitalSigner.prototype.getGuid = function () {
		return this.guid;
	};
	CDigitalSigner.prototype.setValidImg = function (img) {
		this.validImg = this.getBase64Img(img);
	};
	CDigitalSigner.prototype.setInvalidImg = function (img) {
		this.invalidImg = this.getBase64Img(img);
	};
	CDigitalSigner.prototype.getBase64Img = function (img) {
		return "";
	};
	CDigitalSigner.prototype.getValidImage = function () {
		return this.validImg;
	};
	CDigitalSigner.prototype.getInvalidImage = function () {
		return this.invalidImg;
	};
	CDigitalSigner.prototype.getCertInfo = function () {
		return new Promise(function (resolve) {
			resolve();
		});
	};
	CDigitalSigner.prototype.getZip = function () {
		return this.openXml.zip;
	};
	CDigitalSigner.prototype.getDigestAlgorithm = function () {

	};
	CDigitalSigner.prototype.digest = function (file) {
		const digestAlgorithm = this.getFormatDigestAlgorithm();
		return this.crypto.digest(getCryptoDigestFromOOXMLDigest(digestAlgorithm), file);
	};
	CDigitalSigner.prototype.getDigestsFromFiles = function (files) {
		const promises = [];
		for (let i = 0; i < files.length; i += 1) {
			const file = files[i];
			promises.push(this.digest(file));
		}
		return promises;
	}
	CDigitalSigner.prototype.getContentType = function (filePath) {
		return this.openXml.getContentType("/" + filePath);
	};
	CDigitalSigner.prototype.getFormatDigestAlgorithm = function () {
		//todo
		return AscFormat.OOXML_DIGEST_ALGORITHM_TYPE.SHA256;
	};
	CDigitalSigner.prototype.getFormatSignatureAlgorithm = function () {
		//todo
		return AscFormat.OOXML_SIGNATURE_ALGORITHM_TYPE.RSASHA256;
	};
	CDigitalSigner.prototype.getIssuerName = function () {
		//todo
		return "";
	};
	CDigitalSigner.prototype.getSerialNumber = function () {
		//todo
		return "";
	};
	CDigitalSigner.prototype.getBase64Cert = function () {
		//todo
		return "";
	};
	CDigitalSigner.prototype.getRelsReferences = function (filePaths) {
		const openXml = this.openXml;
		const writer = new AscCommon.CMemory();
		writer.context = new AscCommon.XmlWriterContext();
		const oThis= this;
		const files = [];
		const relsStructures = [];
		for (let i = 0; i < filePaths.length; i += 1) {
			const filePath = filePaths[i];
			const uri = "/" + filePath;
			const xmlPart = new AscCommon.openXml.OpenXmlPart(openXml, uri, openXml.getContentType(uri));
			const rels = new AscCommon.openXml.Rels(null, xmlPart);
			new AscCommon.openXml.SaxParserBase().parse(xmlPart.getDocumentContent(), rels);
			const relationships = rels.rels;
			for (let j = relationships.length - 1; j >= 0; j -= 1) {
				const relationship = relationships[j];
				relationship.checkTargetMode();
				if (!isNeedSign(relationship.target)) {
					relationships.splice(j, 1);
				}
			}
			relationships.sort(function (a, b) {
				if (a.relationshipId < b.relationshipId) {
					return -1;
				}
				if (a.relationshipId > b.relationshipId) {
					return 1;
				}
				return 0;
			});
			const relXml = openXml.getXmlBytes(xmlPart, rels, writer);
			files.push(relXml);
			relsStructures.push(rels);
		}
		return Promise.all(this.getDigestsFromFiles(files)).then(function (digests) {
			const references = [];
			for (let i = 0; i < relsStructures.length; i++) {
				const rels = relsStructures[i];
				const filePath = filePaths[i];
				const reference = new AscFormat.CReference();
				reference.transforms = rels.getTransforms();
				reference.URI = "/" + filePath + "?ContentType=" + oThis.getContentType(filePath);
				reference.digestMethod = oThis.getFormatDigestAlgorithm();
				reference.digestValue = digests[i];
				references.push(reference);
			}
			return references;
		});
	};
	CDigitalSigner.prototype.getContentReferences = function (filePaths) {
		const zip = this.getZip();
		const digestPromises = this.getDigestsFromFiles(filePaths.map(function (filePath) {
			return zip.getFile(filePath);
		}));
		const oThis = this;
		return Promise.all(digestPromises).then(function (digests) {
			const references = [];
			for (let i = 0; i < filePaths.length; i += 1) {
				const filePath = filePaths[i];
				const digest = digests[i];
				const contentType =  oThis.getContentType(filePath);
				const reference = new AscFormat.CReference();
				reference.URI = "/" + filePath + "?ContentType=" + contentType;
				reference.digestMethod = oThis.getFormatDigestAlgorithm();
				reference.digestValue = digest;
				references.push(reference);
			}
			return references;
		});
	};
	CDigitalSigner.prototype.getManifestObjectPromise = function () {
		const oThis = this;
		const zip = this.getZip();
		const files = zip.files;
		const relsFiles = [];
		const contentFiles = [];
		files.forEach(function (filePath) {
			if (isNeedSign(filePath)) {
				if (isRels(filePath)) {
					relsFiles.push(filePath);
				} else {
					contentFiles.push(filePath);
				}
			}
		});
		const relsReferencePromises = this.getRelsReferences(relsFiles);
		const contentReferencePromises = this.getContentReferences(contentFiles);
		return Promise.all([relsReferencePromises, contentReferencePromises]).then(function (fileReferences) {
			const relsReferences = fileReferences[0];
			const contentReferences = fileReferences[1];
			const manifestObject = new AscFormat.CManifestObject();
			manifestObject.references = [].concat(relsReferences, contentReferences);
			const signatureTime = new AscFormat.CSignatureTime();
			signatureTime.format = "YYYY-MM-DDThh:mm:ssTZD";
			signatureTime.date = oThis.date.toISOString();
			manifestObject.signatureTime = signatureTime;
			return manifestObject;
		});
	};
	CDigitalSigner.prototype.sign = function () {
		const oThis = this;
		const signature = new AscFormat.CSignature();
		return this.init().then(function () {
			const manifestPromise = oThis.getManifestObjectPromise();
			const signedPropertiesPromise = oThis.getSignedPropertiesPromise();
			return Promise.all([manifestPromise, signedPropertiesPromise]);
		}).then(function (arrPromiseStructures) {
			const manifestObject = arrPromiseStructures[0];
			const signedProperties = arrPromiseStructures[1];
			signature.manifestObject = manifestObject;
			signature.signedProperties = signedProperties;

			signature.officeObject = new AscFormat.COfficeObject();
			signature.officeObject.initDefault(oThis);

			if (oThis.validImg) {
				signature.validSignLnImg = new AscFormat.CValidImageObject();
				signature.validSignLnImg.value = oThis.validImg;
			}
			if (oThis.invalidImg) {
				signature.invalidSignLnImg = new AscFormat.CInvalidImageObject();
				signature.invalidSignLnImg.value = oThis.invalidImg;
			}
			signature.keyInfo = oThis.getKeyInfo();
			return oThis.getSignedInfoPromise(signature);
		}).then(function(signedInfo) {
			signature.signedInfo = signedInfo;
			return this.getSignatureValuePromise();
		}).then(function(signatureValue) {
			signature.signatureValue = signatureValue;

		});
	};
	CDigitalSigner.prototype.getSignatureValuePromise = function() {
		return new Promise(function(resolve, reject) {
			reject("todo");
		});
	};
	CDigitalSigner.prototype.getKeyInfo = function () {
		const keyInfo = new AscFormat.CKeyInfo();
		keyInfo.x509certificate = this.getBase64Cert();
		return keyInfo;
	};
	CDigitalSigner.prototype.getSignedInfoPromise = function(signature) {
		const signedInfo = new AscFormat.CSignedInfo();
		signedInfo.canonicalizationMethod = AscFormat.TRANSFORM_ALGORITHM.CANONIZATION10;
		signedInfo.signatureMethod = this.getFormatSignatureAlgorithm();
		const referencePromises =[];
		referencePromises.push(this.getReferencePromiseFromObject(signature.manifestObject, "#" + AscFormat.OBJECT_ID.MANIFEST));
		referencePromises.push(this.getReferencePromiseFromObject(signature.officeObject, "#" + AscFormat.OBJECT_ID.OFFICE));
		referencePromises.push(this.getReferencePromiseFromObject(signature.signedPropertiesObject, "#" + AscFormat.OBJECT_ID.SIGNED_PROPERTIES));
		if (signature.validSignLnImg) {
			referencePromises.push(this.getReferencePromiseFromObject(signature.validSignLnImg, "#" + AscFormat.OBJECT_ID.VALID_IMG));
		}
		if (signature.invalidSignLnImg) {
			referencePromises.push(this.getReferencePromiseFromObject(signature.invalidSignLnImg, "#" + AscFormat.OBJECT_ID.INVALID_IMG));
		}
		return Promise.all(referencePromises).then(function(references) {
			signedInfo.references = references;
			return signedInfo;
		});
	};
	CDigitalSigner.prototype.getReferencePromiseFromObject = function(object, uri) {
		const reference = new AscFormat.CReference();
		const digestMethod = this.getFormatDigestAlgorithm();
		reference.type = AscFormat.REFERENCE_TYPE.OBJECT;
		reference.URI = uri;
		reference.digestMethod = digestMethod;
		const digestPromise = this.getDigestFromObject(object);
		return digestPromise.then(function(digestValue) {
			reference.digestValue = digestValue;
			return reference;
		});
	};
	CDigitalSigner.prototype.getDigestFromObject = function(object) {
		const writer = new AscCommon.CMemory();
		writer.context = new AscCommon.XmlWriterContext();
		object.toXml(writer);
		const length = writer.GetCurPosition();
		const xmlBytes = writer.GetDataUint8(0, length);
		return this.digest(xmlBytes);
	};

	CDigitalSigner.prototype.getSignedPropertiesPromise = function () {
		const oThis = this;
		return this.getDigestCertValue().then(function (digestValue) {
			const signedProperties = new AscFormat.CSignedPropertiesObject();
			signedProperties.signingTime = oThis.date.toISOString();

			const cert = new AscFormat.CCert();
			signedProperties.cert = cert;

			const certDigest = new AscFormat.CCertDigest();
			cert.certDigest = certDigest;
			certDigest.digestValue = digestValue;
			certDigest.digestMethod = oThis.getFormatDigestAlgorithm();

			const issuerSerial = new AscFormat.CIssuerSerial()
			cert.issuerSerial = issuerSerial;
			issuerSerial.x509IssuerName = oThis.getIssuerName();
			issuerSerial.x509SerialNumber = oThis.getSerialNumber();

			signedProperties.signaturePolicyIdentifier = new AscFormat.CSignaturePolicyIdentifier();
			signedProperties.signaturePolicyIdentifier.implied = true;

			return signedProperties;
		});
	};

	window["AscCommon"] = window["AscCommon"] || {};
	window["AscCommon"].CDigitalSigner = CDigitalSigner;
})();

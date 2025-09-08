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

const CRYPTO_DIGEST_ALGORITHM_TYPE = {
	SHA256: "SHA-256"
}
const OOXML_DIGEST_ALGORITHM_TYPE = {
	SHA1: 1,
	SHA224: 2,
	SHA256: 3,
	SHA384: 4,
	SHA512: 5
}

const TRANSFORM_ALGORITHM = {
	RELATIONSHIP: 1,
	CANONIZATION10: 2,
};

const SIGNATURE_TYPE = {
	INVISIBLE: 1,
	SIGNATURE_LINE: 2
};

function digest() {

}
function CX509Cert() {

}
CX509Cert.prototype.getDigestAlgorithm = function () {
	return CRYPTO_DIGEST_ALGORITHM_TYPE.SHA256;
};
function getEnumFromStrDigestAlgorithm(str) {
	switch (str) {
		case "http://www.w3.org/2000/09/xmldsig#rsa-sha1":
		case "http://www.w3.org/2000/09/xmldsig#sha1": {
			return OOXML_DIGEST_ALGORITHM_TYPE.SHA1;
		}
		case "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256":
		case "http://www.w3.org/2001/04/xmldsig-more#sha256":
		case "http://www.w3.org/2001/04/xmlenc#sha256": {
			return OOXML_DIGEST_ALGORITHM_TYPE.SHA256;
		}
		case "http://www.w3.org/2001/04/xmldsig-more#rsa-sha384":
		case "http://www.w3.org/2001/04/xmldsig-more#sha384": {
			return OOXML_DIGEST_ALGORITHM_TYPE.SHA384;
		}
		case "http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha512":
		case "http://www.w3.org/2001/04/xmldsig-more#rsa-sha512":
		case "http://www.w3.org/2001/04/xmldsig-more#sha512":
		case "http://www.w3.org/2001/04/xmlenc#sha512": {
			return OOXML_DIGEST_ALGORITHM_TYPE.SHA512;
		}
		default: {
			return null;
		}
	}
}
function getFormatDigestAlgorithm(cryptoAlgorithm) {
	switch (cryptoAlgorithm) {
		case CRYPTO_DIGEST_ALGORITHM_TYPE.SHA256:
			return FORMAT_DIGEST_ALGORITHM_TYPE.SHA256;
	}
	return null;
}

(function () {
	function CDigitalSigner(crypto, cert, zip,) {
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
		if (this.isInit) {
			return new Promise(function (resolve) {
				resolve();
			});
		}
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
	CDigitalSigner.prototype.getDigestMethod = function () {
		return new Promise(function (resolve, reject) {
			reject("todo");
		});
	};
	CDigitalSigner.prototype.digest = function (file) {
		return this.crypto.digest("SHA-256", file);
		// return new Promise(function (resolve, reject) {
		// 	reject("todo");
		// });
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
		return 1;
	};
	CDigitalSigner.prototype.getIssuerName = function () {
		return "";
	};
	CDigitalSigner.prototype.getSerialNumber = function () {
		return "";
	};
	CDigitalSigner.prototype.getBase64Cert = function () {
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
				const reference = new CReference();
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
				const reference = new CReference();
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
		const signFiles = [];
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
			const manifestObject = new CManifestObject();
			manifestObject.references = [].concat(relsReferences, contentReferences);
			const signatureTime = new CSignatureTime();
			signatureTime.format = "YYYY-MM-DDThh:mm:ssTZD";
			signatureTime.date = oThis.date.toISOString();
			manifestObject.signatureTime = signatureTime;
			return manifestObject;
		});
	};
	CDigitalSigner.prototype.sign = function () {
		const oThis = this;
		const signature = new CSignature();
		return this.init().then(function () {
			const manifestPromise = oThis.getManifestObjectPromise();
			const signedPropertiesPromise = oThis.getSignedPropertiesPromise();
			return Promise.all([manifestPromise, signedPropertiesPromise]);
		}).then(function (arrPromiseStructures) {
			const manifestObject = arrPromiseStructures[0];
			const signedProperties = arrPromiseStructures[1];
			signature.manifestObject = manifestObject;
			signature.signedProperties = signedProperties;

			signature.officeObject = new COfficeObject();
			signature.officeObject.initDefault(oThis);

			if (oThis.validImg) {
				signature.validSignLnImg = new CImageObject();
				signature.validSignLnImg.value = oThis.validImg;
			}
			if (oThis.invalidImg) {
				signature.invalidSignLnImg = new CImageObject();
				signature.invalidSignLnImg.value = oThis.invalidImg;
			}
			const signedInfo = new CSignedInfo();
			signedInfo.canonicalizationMethod

			signature.keyInfo = oThis.getKeyInfo();
		});
	};
	CDigitalSigner.prototype.getKeyInfo = function () {
		const keyInfo = new CKeyInfo();
		keyInfo.x509certificate = this.getBase64Cert();
		return keyInfo;
	};

	CDigitalSigner.prototype.getSignedPropertiesPromise = function () {
		const oThis = this;
		return this.getDigestCertValue().then(function (digestValue) {
			const signedProperties = new CSignedPropertiesObject();
			signedProperties.signingTime = oThis.date.toISOString();

			const cert = new CCert();
			signedProperties.cert = cert;

			const certDigest = new CCertDigest();
			cert.certDigest = certDigest;
			certDigest.digestValue = digestValue;
			certDigest.digestMethod = oThis.getFormatDigestAlgorithm();

			const issuerSerial = new CIssuerSerial()
			cert.issuerSerial = issuerSerial;
			issuerSerial.x509IssuerName = oThis.getIssuerName();
			issuerSerial.x509SerialNumber = oThis.getSerialNumber();

			signedProperties.signaturePolicyIdentifier = new CSignaturePolicyIdentifier();
			signedProperties.signaturePolicyIdentifier.implied = true;

			return signedProperties;
		});
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


	function CSignature() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.signedInfo = null;
		this.signatureValue = null;
		this.keyInfo = null;
		this.manifestObject = null;
		this.officeObject = null;
		this.signedPropertiesObject = null;
		this.validSignLnImg = null;
		this.invalidSignLnImg = null;
	}
	// AscCommon.InitClassWithoutType(CSignature, AscFormat.CBaseNoIdObject);

	function CSignedInfo() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.canonicalizationMethod = null;
		this.signatureMethod = null;
		this.references = [];
	}
	// AscCommon.InitClassWithoutType(CSignedInfo, AscFormat.CBaseNoIdObject);

	function CReference() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.type = null;
		this.URI = null;
		this.digestMethod = null;
		this.digestValue = null;
		this.transforms = [];
	}
	// AscCommon.InitClassWithoutType(CReference, AscFormat.CBaseNoIdObject);

	function CTransform() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.algorithm = null;
		this.relationshipReferences = [];
	}
	// AscCommon.InitClassWithoutType(CTransform, AscFormat.CBaseNoIdObject);

	function CRelationshipReference() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.sourceId = null;
	}
	// AscCommon.InitClassWithoutType(CRelationshipReference, AscFormat.CBaseNoIdObject);

	function CKeyInfo() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.x509certificate = null;
	}
	// AscCommon.InitClassWithoutType(CKeyInfo, AscFormat.CBaseNoIdObject);

	function CManifestObject() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.references = [];
		this.signatureTime = null;
	}
	// AscCommon.InitClassWithoutType(CManifestObject, AscFormat.CBaseNoIdObject);

	function CSignatureTime() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.format = null;
		this.value = null;
	}
	// AscCommon.InitClassWithoutType(CSignatureTime, AscFormat.CBaseNoIdObject);

	function COfficeObject() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.signatureInfo = null;
	}
	// AscCommon.InitClassWithoutType(COfficeObject, AscFormat.CBaseNoIdObject);
	COfficeObject.prototype.initDefault = function (signer) {
		this.signatureInfo = new CSignatureInfo();
		this.signatureInfo.initDefault(signer);
	};

	function CSignatureInfo() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.applicationVersion = null;
		this.colorDepth = null;
		this.delegateSuggestedSigner = null;
		this.delegateSuggestedSigner2 = null;
		this.delegateSuggestedSignerEmail = null;
		this.horizontalResolution = null;
		this.manifestHashAlgorithm = null;
		this.monitors = null;
		this.officeVersion = null;
		this.setupID = null;
		this.signatureComments = null;
		this.signatureImage = null;
		this.signatureProviderDetails = null;
		this.signatureProviderId = null;
		this.signatureProviderUrl = null;
		this.signatureText = null;
		this.signatureType = null;
		this.verticalResolution = null;
		this.windowsVersion = null;
	}
	// AscCommon.InitClassWithoutType(CSignatureInfo, AscFormat.CBaseNoIdObject);
	CSignatureInfo.prototype.initDefault = function (signer) {
		this.applicationVersion = "16.0";
		this.colorDepth = 32;
		this.delegateSuggestedSigner = null;
		this.delegateSuggestedSigner2 = null;
		this.delegateSuggestedSignerEmail = null;
		this.horizontalResolution = 1680;
		this.manifestHashAlgorithm = null;
		this.monitors = 2;
		this.officeVersion = "16.0";
		this.setupID = signer.getGuid();
		this.signatureComments = null;
		this.signatureImage = signer.getValidImage();
		this.signatureProviderDetails = 9;
		this.signatureProviderId = "{00000000-0000-0000-0000-000000000000}";
		this.signatureProviderUrl = "";
		this.signatureText = "";
		this.signatureType = SIGNATURE_TYPE.SIGNATURE_LINE;
		this.verticalResolution = 1050;
		this.windowsVersion = "10.0";
	};

	function CSignedPropertiesObject() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.signingTime = null;
		this.cert = null;
		this.signaturePolicyIdentifier = null;
	}
	// AscCommon.InitClassWithoutType(CSignedPropertiesObject, AscFormat.CBaseNoIdObject);

	function CSignaturePolicyIdentifier() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.implied = false;
	}
	// AscCommon.InitClassWithoutType(CSignaturePolicyIdentifier, AscFormat.CBaseNoIdObject);

	function CCert() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.certDigest = null;
		this.issuerSerial = null;
	}
	// AscCommon.InitClassWithoutType(CCert, AscFormat.CBaseNoIdObject);
	function CCertDigest() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.digestMethod = null;
		this.digestValue = null;
	}
	// AscCommon.InitClassWithoutType(CCertDigest, AscFormat.CBaseNoIdObject);

	function CIssuerSerial() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.x509IssuerName = null;
		this.x509SerialNumber = null;
	}
	// AscCommon.InitClassWithoutType(CIssuerSerial, AscFormat.CBaseNoIdObject);
	function CImageObject() {
		// AscFormat.CBaseNoIdObject.call(this);
		this.value = null;
	}
	// AscCommon.InitClassWithoutType(CImageObject, AscFormat.CBaseNoIdObject);

	window["AscCommon"] = window["AscCommon"] || {};
	window["AscCommon"].CDigitalSigner = CDigitalSigner;

	window["AscFormat"] = window["AscFormat"] || {};
	window["AscFormat"].CRelationshipReference = CRelationshipReference;
	window["AscFormat"].CTransform = CTransform;
	window["AscFormat"].TRANSFORM_ALGORITHM = TRANSFORM_ALGORITHM;
})();

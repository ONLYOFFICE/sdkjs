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
	function CDigitalSigner(crypto, cert, zip) {
		this.crypto = crypto;
		this.zip = zip;
		this.cert = cert || new CX509Cert()/*todo temp*/;
		this.date = new Date();
		this.isInit = false;
		this.certGuid = cert;
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
				});
			}
		});
	};
	CDigitalSigner.prototype.getDigestAlgorithm = function () {

	};
	CDigitalSigner.prototype.getDigestMethod = function () {
		return new Promise(function (resolve, reject) {
			reject("todo");
		});
	};
	CDigitalSigner.prototype.digest = function (file) {
		return new Promise(function (resolve, reject) {
			reject("todo");
		});
	};
	CDigitalSigner.prototype.getDigestsFromFiles = function (filePaths) {
		const promises = [];
		for (let i = 0; i < filePaths.length; i += 1) {
			const filePath = filePaths[i];
			const file = this.zip.getFile(filePath);
			promises.push(this.digest(file));
		}
		return promises;
	}
	CDigitalSigner.prototype.getContentType = function (filePath) {

	};
	CDigitalSigner.prototype.getRelsReferences = function (filePaths, digests) {
		for (let i = 0; i < filePaths.length; i += 1) {
			const reference = new CReference();

		}
	};
	CDigitalSigner.prototype.getContentReferences = function (filePaths, digests) {
		const references = [];
		for (let i = 0; i < filePaths.length; i += 1) {
			const filePath = filePaths[i];
			const digest = digests[i];
			const contentType =  this.getContentType(filePath);
			const reference = new CReference();
			reference.URI = "/" + filePath + "?ContentType=" + contentType;
			reference.digestMethod = this.getFormatDigestAlgorithm();
			reference.digestValue = digest;
			references.push(reference);
		}
		return references;
	};
	CDigitalSigner.prototype.getManifestObjectPromise = function () {
		const oThis = this;
		const zip = this.zip;
		const files = zip.files;
		const promises = [];
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

		const relsPromises = this.getDigestsFromFiles(relsFiles);
		const contentPromises = this.getDigestsFromFiles(contentFiles);
		return Promise.all([Promise.all(relsPromises), Promise.all(contentPromises)]).then(function (fileDigests) {
			const relsDigests = fileDigests[0];
			const contentDigests = fileDigests[1];
			const relsReferences = oThis.getRelsReferences(relsFiles, relsDigests);
			const contentReferences = oThis.getContentReferences(contentFiles, contentDigests);
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
			return oThis.getManifestObjectPromise();
		}).then(function (manifest) {

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
})();

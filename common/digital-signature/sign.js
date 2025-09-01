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

function CX509Cert() {

}
CX509Cert.prototype.getDigestAlgorithm = function () {
	return "SHA-256";
};

(function () {
	function CDigitalSigner(crypto, cert, zip) {
		this.crypto = crypto;
		this.zip = zip;
		this.cert = cert || new CX509Cert()/*todo temp*/;
	}
	CDigitalSigner.prototype.getDigestAlgorithm = function () {

	};
	CDigitalSigner.prototype.sign = function () {
		const oThis = this;
		const zip = this.zip;
		const files = zip.files;
		const promises = [];
		const digestAlgorithm = this.cert.getDigestAlgorithm();
		files.forEach(function (filePath) {
			if (oThis.isNeedSign(filePath)) {
				const file = zip.getFile(filePath);
				promises.push(oThis.crypto.digest(digestAlgorithm, file));
			}
		});
		Promise.all(promises).then(function (digestArray) {
			for (let i = 0; i < digestArray.length; i += 1) {
				const intArray = new Uint8Array(digestArray[i]);
				console.log(getBase64FromBuffer(digestArray[i]))
			}
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
	CDigitalSigner.prototype.isNeedSign = function () {
		return true;
	};


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

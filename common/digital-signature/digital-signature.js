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
	const OOXML_SIGNATURE_ALGORITHM_TYPE = {
		RSASHA1: 1,
		RSASHA224: 2,
		RSASHA256: 3,
		RSASHA384: 4,
		RSASHA512: 5,
		ECDSASHA1: 6,
		ECDSASHA224: 7,
		ECDSASHA256: 8,
		ECDSASHA384: 9,
		ECDSASHA512: 10,
		DSASHA1: 11,
		DSASHA256: 12
	};
	const OOXML_DIGEST_ALGORITHM_TYPE = {
		SHA1: 1,
		SHA224: 2,
		SHA256: 3,
		SHA384: 4,
		SHA512: 5
	};
	const SIGNATURE_TYPE = {
		INVISIBLE: 1,
		SIGNATURE_LINE: 2
	};
	const OBJECT_ID = {
		SIGNATURE: "idPackageSignature",
		MANIFEST: "idPackageObject",
		OFFICE: "idOfficeObject",
		SIGNED_PROPERTIES: "idSignedProperties",
		VALID_IMG: "idValidSigLnImg",
		INVALID_IMG: "idInvalidSigLnImg",
		SIGNATURE_TIME: "idSignatureTime",
		OFFICE_DETAILS: "idOfficeV1Details"
	};

	const REFERENCE_TYPE = {
		OBJECT: 1
	};

	const TRANSFORM_ALGORITHM = {
		RELATIONSHIP: 1,
		CANONIZATION10: 2,
	};

	function CSignature() {
		AscFormat.CBaseNoIdObject.call(this);
		this.signedInfo = null;
		this.signatureValue = null;
		this.keyInfo = null;
		this.manifestObject = null;
		this.officeObject = null;
		this.signedPropertiesObject = null;
		this.validSignLnImg = null;
		this.invalidSignLnImg = null;
	}
	AscCommon.InitClassWithoutType(CSignature, AscFormat.CBaseNoIdObject);

	function CSignedInfo() {
		AscFormat.CBaseNoIdObject.call(this);
		this.canonicalizationMethod = null;
		this.signatureMethod = null;
		this.references = [];
	}
	AscCommon.InitClassWithoutType(CSignedInfo, AscFormat.CBaseNoIdObject);

	function CReference() {
		AscFormat.CBaseNoIdObject.call(this);
		this.type = null;
		this.URI = null;
		this.digestMethod = null;
		this.digestValue = null;
		this.transforms = [];
	}
	AscCommon.InitClassWithoutType(CReference, AscFormat.CBaseNoIdObject);

	function CTransform() {
		AscFormat.CBaseNoIdObject.call(this);
		this.algorithm = null;
		this.relationshipReferences = [];
	}
	AscCommon.InitClassWithoutType(CTransform, AscFormat.CBaseNoIdObject);

	function CRelationshipReference() {
		AscFormat.CBaseNoIdObject.call(this);
		this.sourceId = null;
	}
	AscCommon.InitClassWithoutType(CRelationshipReference, AscFormat.CBaseNoIdObject);

	function CKeyInfo() {
		AscFormat.CBaseNoIdObject.call(this);
		this.x509certificate = null;
	}
	AscCommon.InitClassWithoutType(CKeyInfo, AscFormat.CBaseNoIdObject);

	function CManifestObject() {
		AscFormat.CBaseNoIdObject.call(this);
		this.references = [];
		this.signatureTime = null;
	}
	AscCommon.InitClassWithoutType(CManifestObject, AscFormat.CBaseNoIdObject);

	function CSignatureTime() {
		AscFormat.CBaseNoIdObject.call(this);
		this.format = null;
		this.value = null;
	}
	AscCommon.InitClassWithoutType(CSignatureTime, AscFormat.CBaseNoIdObject);

	function COfficeObject() {
		AscFormat.CBaseNoIdObject.call(this);
		this.signatureInfo = null;
	}
	AscCommon.InitClassWithoutType(COfficeObject, AscFormat.CBaseNoIdObject);
	COfficeObject.prototype.initDefault = function (signer) {
		this.signatureInfo = new CSignatureInfo();
		this.signatureInfo.initDefault(signer);
	};

	function CSignatureInfo() {
		AscFormat.CBaseNoIdObject.call(this);
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
	AscCommon.InitClassWithoutType(CSignatureInfo, AscFormat.CBaseNoIdObject);
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
		AscFormat.CBaseNoIdObject.call(this);
		this.signingTime = null;
		this.cert = null;
		this.signaturePolicyIdentifier = null;
	}
	AscCommon.InitClassWithoutType(CSignedPropertiesObject, AscFormat.CBaseNoIdObject);

	function CSignaturePolicyIdentifier() {
		AscFormat.CBaseNoIdObject.call(this);
		this.implied = false;
	}
	AscCommon.InitClassWithoutType(CSignaturePolicyIdentifier, AscFormat.CBaseNoIdObject);

	function CCert() {
		AscFormat.CBaseNoIdObject.call(this);
		this.certDigest = null;
		this.issuerSerial = null;
	}
	AscCommon.InitClassWithoutType(CCert, AscFormat.CBaseNoIdObject);
	function CCertDigest() {
		AscFormat.CBaseNoIdObject.call(this);
		this.digestMethod = null;
		this.digestValue = null;
	}
	AscCommon.InitClassWithoutType(CCertDigest, AscFormat.CBaseNoIdObject);

	function CIssuerSerial() {
		AscFormat.CBaseNoIdObject.call(this);
		this.x509IssuerName = null;
		this.x509SerialNumber = null;
	}
	AscCommon.InitClassWithoutType(CIssuerSerial, AscFormat.CBaseNoIdObject);
	function CImageObject() {
		AscFormat.CBaseNoIdObject.call(this);
		this.value = null;
	}
	AscCommon.InitClassWithoutType(CImageObject, AscFormat.CBaseNoIdObject);

	function CValidImageObject() {
		CImageObject.call(this);
	}
	AscCommon.InitClassWithoutType(CValidImageObject, CImageObject);

	function CInvalidImageObject() {
		CImageObject.call(this);
	}
	AscCommon.InitClassWithoutType(CInvalidImageObject, CImageObject);

	window["AscFormat"] = window["AscFormat"] || {};
	window["AscFormat"].CSignature = CSignature;
	window["AscFormat"].CSignedInfo = CSignedInfo;
	window["AscFormat"].CReference = CReference;
	window["AscFormat"].CTransform = CTransform;
	window["AscFormat"].CRelationshipReference = CRelationshipReference;
	window["AscFormat"].CKeyInfo = CKeyInfo;
	window["AscFormat"].CManifestObject = CManifestObject;
	window["AscFormat"].CSignatureTime = CSignatureTime;
	window["AscFormat"].COfficeObject = COfficeObject;
	window["AscFormat"].CSignatureInfo = CSignatureInfo;
	window["AscFormat"].CSignedPropertiesObject = CSignedPropertiesObject;
	window["AscFormat"].CSignaturePolicyIdentifier = CSignaturePolicyIdentifier;
	window["AscFormat"].CCert = CCert;
	window["AscFormat"].CCertDigest = CCertDigest;
	window["AscFormat"].CIssuerSerial = CIssuerSerial;
	window["AscFormat"].CImageObject = CImageObject;
	window["AscFormat"].CValidImageObject = CValidImageObject;
	window["AscFormat"].CInvalidImageObject = CInvalidImageObject;
	window["AscFormat"].TRANSFORM_ALGORITHM = TRANSFORM_ALGORITHM;
	window["AscFormat"].OBJECT_ID = OBJECT_ID;
	window["AscFormat"].REFERENCE_TYPE = REFERENCE_TYPE;
	window["AscFormat"].OOXML_SIGNATURE_ALGORITHM_TYPE = OOXML_SIGNATURE_ALGORITHM_TYPE;
	window["AscFormat"].OOXML_DIGEST_ALGORITHM_TYPE = OOXML_DIGEST_ALGORITHM_TYPE;
})();

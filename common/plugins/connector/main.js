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

var connector = new Asc.EditorConnector({frame : "frameEditor"});

document.getElementById("buttonConnect").onclick = function()
{
	connector.connect();
};
document.getElementById("buttonDisconnect").onclick = function()
{
	connector.disconnect();
};
document.getElementById("buttonTest1").onclick = function()
{
	connector.executeMethod("GetVersion", [], function(version)
	{
		console.log(version);
	});

	connector.attachEvent("onTargetPositionChanged", function()
	{
		console.log("event: onTargetPositionChanged");
	});

	connector.attachEvent("onChangeContentControl", function()
	{
		console.log("event: onChangeContentControl");
	});

	connector.attachEvent("onFocusContentControl", function()
	{
		console.log("event: onFocusContentControl");
	});


	connector.attachEvent("onBlurContentControl", function(oPr)
	{
		if (oPr && "BankBIC" === oPr["Tag"])
		{
			connector.executeMethod("GetFormValue", [oPr["InternalId"]], function(value)
			{
				if ("12345678" !== value)
					return;

				connector.executeMethod("GetFormsByTag", ["BankAccount"], function(forms)
				{
					for (let i = 0; i < forms.length; ++i)
					{
						connector.executeMethod("SetFormValue", [forms[i]["InternalId"], "10101110100000000123"], null);
					}
				});

				connector.executeMethod("GetFormsByTag", ["BankName"], function(forms)
				{
					for (let i = 0; i < forms.length; ++i)
					{
						connector.executeMethod("SetFormValue", [forms[i]["InternalId"], "OnlyOffice BANK"], null);
					}
				});

				connector.executeMethod("GetFormsByTag", ["BankPlace"], function(forms)
				{
					for (let i = 0; i < forms.length; ++i)
					{
						connector.executeMethod("SetFormValue", [forms[i]["InternalId"], "Himalayas"], null);
					}
				});
			});
		}
		console.log("event: onBlurContentControl");
	});

};
document.getElementById("buttonTest2").onclick = function()
{
	connector.callCommand(function() {

		var oDocument = Api.GetDocument();
		var oParagraph = Api.CreateParagraph();
		oParagraph.AddText("Hello world!");
		oDocument.InsertContent([oParagraph]);

	}, function() { console.log("callback command"); });
};

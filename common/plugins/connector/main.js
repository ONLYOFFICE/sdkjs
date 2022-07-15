/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

var connector = new Asc.EditorConnector("frameEditor");

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
	connector.callMethod("GetVersion", [], function(version)
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
		let sTag = oPr["Tag"];
		console.log(oPr);
		console.log(sTag);
		if ("BankBIC" === sTag)
		{
			connector.callCommand(function() {

				let oDocument = Api.GetDocument();

				let arrForms  = oDocument.GetFormsByTag("BankBIC");
				if (1 === arrForms.length &&  arrForms[0].GetText() === "12345678")
				{
					arrForms = oDocument.GetFormsByTag("BankAccount");
					for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
					{
						let oForm = arrForms[nIndex];
						if (oForm.GetFormType() === "textForm")
							oForm.SetText("10101110100000000123");
					}

					arrForms = oDocument.GetFormsByTag("BankName");
					for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
					{
						let oForm = arrForms[nIndex];
						if (oForm.GetFormType() === "textForm")
							oForm.SetText("OnlyOffice BANK");
					}

					arrForms = oDocument.GetFormsByTag("BankPlace");
					for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
					{
						let oForm = arrForms[nIndex];
						if (oForm.GetFormType() === "textForm")
							oForm.SetText("Himalayas");
					}
				}
			}, function() { console.log("callback command"); });


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

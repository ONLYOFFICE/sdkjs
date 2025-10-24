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

(function(window)
{
	/**
	 * Class to work with MathML
	 * @constructor
	 */
	function MathML()
	{
		/**
		 *
		 * @param {string} xml
		 * @param {AscWord.CTextPr} [textPr=undefined]
		 * @returns {AscWord.ParaMath}
		 */
		this.toParaMath = function(xml, textPr)
		{
			xml = xml ? xml : "";

			// Replace <br> tags (with any attributes) with nothing, as they're not valid MathML
			xml = xml.replace(/<br[^>]*>/gi, '');

			let paraMath = new AscWord.ParaMath();
			
			this.mathMLData = {
				'mo-linebreak-todo' : [],
				'mo-id' : {}
			};

			let reader = new StaxParser(xml);
			if (!reader.ReadNextNode() || "math" !== reader.GetNameNoNS())
			{
				paraMath.Root.Correct_Content(true);
				return paraMath;
			}
			
			this.handleMathContent(reader, paraMath, paraMath.Root);
			paraMath.Root.Correct_Content(true);
			paraMath.SetParagraph(null);
			
			//paraMath.applyMathMLGlobalAttributes();
			
			if (textPr)
			{
				textPr.RFonts.SetAll("Cambria Math");
				paraMath.ApplyTextPr(textPr, undefined, true);
			}
			return paraMath;
		};
        this.addFlatToMathContent = function(mathContent, element)
		{
			// CMathContent cannot be nested inside another CMathContent
			if (element instanceof AscCommonWord.CMathContent)
			{
				for (let i = 0; i < element.Content.length; i++)
				{
					mathContent.addElementToContent(element.Content[i]);
				}
			}
			else
			{
				mathContent.addElementToContent(element);
			}
		};
        this.GetAttributes = function(reader)
        {
            let attributes = reader.GetAttributes();

            function valuesToLower(obj)
            {
                let out = {};
                for (let k in obj) {
                    if (Object.prototype.hasOwnProperty.call(obj, k))
                    {
                        let v = obj[k];
                        out[k] = (typeof v === 'string')
                            ? v.toLowerCase()
                            : v;
                    }
                }
                return out;
            }
            return valuesToLower(attributes);
        };
        this.proceedMathMLDefaultAttributes = function(attributes, elements, el, name)
        {
            let keysAttributes = Object.keys(attributes);
            for (let nAttribute = 0; nAttribute < keysAttributes.length; nAttribute++)
            {
                switch (keysAttributes[nAttribute])
                {
                    case 'mathcolor':
                    {
                        let rgb = AscFormat.mapPrstColor[attributes['mathcolor']];
                        if (rgb) {
                            let color = new AscWord.CDocumentColor((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
                            el.Set_Color(color);
                        }
                        break;
                    }
                    case 'mathbackground':
                    {
                        let rgb = AscFormat.mapPrstColor[attributes['mathbackground']];
                        if (rgb)
                        {
                            let color = new AscWord.CDocumentColor((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
                            el.Set_HighLight(color);
                        }
                        break;
                    }
                    case 'mathsize':
                    {
                        //default size if 16px or 1em
                        if (attributes['mathsize'] === 'normal')
                        {
                            el.SetFontSize(16);
                        }
                        else if (attributes['mathsize'] === 'small')
                        {
                            el.SetFontSize(12);
                        }
                        else if (attributes['mathsize'] === 'big')
                        {
                            el.SetFontSize(20);
                        }
                        else if (attributes['mathsize'].slice(-1) === "%" && Number(attributes['mathsize'].slice(0, -1)))
                        {
                            let percentage		= Number(attributes['mathsize'].slice(0, -1));
                            let defaultvalue	= 16;
                            let value			= defaultvalue * (percentage / 100);
                            el.SetFontSize(value);
                        }
                    }
                    case 'stretchy': {
                        if (attributes['stretchy'] === 'false')
                        {
                            if (!el.mathml_metadata)
                                el.mathml_metadata = {};
                            el.mathml_metadata.stretchy = false;
                        }
                    }
                    default: break;
                }
            }
            elements.push(el);
        };
        this.proceedMathMLImplicitDelimiter = function(result, start, end)
        {
            let props			= new CMathDelimiterPr();
            props.begChr		= start ? start.charCodeAt(0) : -1;
            props.endChr		= end ? end.charCodeAt(0) : -1;
            props.sepChr		= ','.charCodeAt(0);
            props.shp			= DELIMITER_SHAPE_MATCH;

            let mathContent = new CMathContent();
            for (let i = 0; i < result.length; i++)
            {
                mathContent.addElementToContent(result[i]);
            }

            mathContent.Correct_Content(true);
            props.content = [mathContent];

            return [this.handleDelimiter(undefined, props)];
        };
        this.readMathMLNode = function(reader, parentMathContent)
        {
            function decodeHexEntities(str)
            {
                var result = '';
                for (var i = 0; i < str.length;)
                {
                    if (str.charAt(i) === '&' && str.charAt(i + 1) === '#' && str.charAt(i + 2).toLowerCase() === 'x')
                    {
                        var j = i + 3;
                        var hex = '';
                        while (j < str.length && str.charAt(j) !== ';') {
                            hex += str.charAt(j);
                            j++;
                        }
                        if (str.charAt(j) === ';') {
                            var code = parseInt(hex, 16);
                            if (!isNaN(code)) {
                                result += String.fromCharCode(code);
                                i = j + 1;
                                continue;
                            }
                        }
                    }
    
                    result += str.charAt(i);
                    i++;
                }
    
                return result;
            }

            function proceedAttributes(name, text, attributes)
            {
                const GetMathFontChar = AscMath.GetMathFontChar;
                switch (name)
                {
                    case 'mi':
                    {
                        let type = -1;
                        switch (attributes['mathvariant'])
                        {
                            case 'bold':					type = 0; break;
                            case 'italic':					type = 1; break;
                            case 'bold-italic':				type = 2; break;
                            case 'double-struck':			type = 12; break;
                            case 'bold-fraktur':			type = 10; break;
                            case 'script':					type = 7; break;
                            case 'bold-script':				type = 8; break;
                            case 'fraktur':					type = 9; break;
                            case 'sans-serif':				type = 3; break;
                            case 'bold-sans-serif':			type = 4; break;
                            case 'sans-serif-italic':		type = 5; break;
                            case 'sans-serif-bold-italic':	type = 6; break;
                            case 'monospace':				type = 11; break;
                            default:						type = -1; break;
                        }
    
                        // single character must be italic
                        if (text.length === 1 && type === -1 && !attributes['mathvariant'])
                            type = 1;
    
                        let convertedText = "";
                        for (let oIter = text.getUnicodeIterator(); oIter.check(); oIter.next())
                        {
                            let currentChar = String.fromCodePoint(oIter.value());
    
                            if (GetMathFontChar[currentChar] && GetMathFontChar[currentChar][type])
                                convertedText += GetMathFontChar[currentChar][type];
                            else
                                convertedText += currentChar;
                        }
                        text = convertedText;
                        break;
                    }
                    case 'mo':
                    {
                        if (attributes['id'])
                            this.mathMLData['mo-id'][attributes['id']] = elements[0];
    
                        if (attributes['linebreak'])
                        {
                            this.mathMLData['mo-linebreak-todo'].push({
                                element: elements[0],
                                type: attributes['linebreak'],
                                indentalign: attributes['indentalign'] === "id",
                                indenttarget: attributes['indenttarget']
                            });
                        }
    
                        // indent
                        if (attributes['movablelimits'])
                        {
                            parentMathContent.mathml_metadata['movablelimits'] = attributes['movablelimits'];
                        }
    
                        break;
                    }
                    case 'mspace':
                    {
                        text = "";
                        break;
                    }
                    case 'ms':
                    {
                        if (attributes['lquote'])
                        {
                            text = attributes['lquote'] + text;
                        }
                        else
                        {
                            text = "\"" + text;
                        }
    
                        if (attributes['rquote'])
                        {
                            text = text + attributes['rquote'];
                        }
                        else
                        {
                            text = text + "\"";
                        }
                        break;
                    }
                }
                return text;
            }

            function updateBaseText(text)
            {
                text = text.trim();
                text = text.replaceAll(String.fromCharCode(8290), ""); // invisible *
                text = text.replaceAll(String.fromCharCode(8292), ""); // invisible +
                text = text.replaceAll("\n", "");
                text = text.replaceAll("\r", "");
                return text;
            }
            let elements = [];
            let name = reader.GetNameNoNS();
            switch (name)
            {
                case 'mi':
                case 'mo':
                case 'mn':
                case 'mspace':
                case 'mtext':
                case 'ms':
                    let attributes = this.GetAttributes(reader);
                    let text = updateBaseText(reader.GetText());

                    if (text.length === 0)
                        break;

                    text = decodeHexEntities(text);
                    text = proceedAttributes(name, text, attributes);
                    text = text.replaceAll("&nbsp;", " ");
    
                    if (text)
                    {
                        let run = new AscWord.Run(null, true);
                        run.AddText(text);
                        this.proceedMathMLDefaultAttributes(attributes, elements, run, name);
                    }
                    break;
                case 'menclose':
                    this.proceedMathMLDefaultAttributes(
                        this.GetAttributes(reader),
                        elements,
                        this.handleBorderBox(reader)
                    );
                    break;
                case 'mphantom':
                    elements.push(this.handlePhantom(reader));
                    break;
                case 'msub':
                    elements.push(this.handleDegree(reader, DEGREE_SUBSCRIPT));
                    break;
                case 'msup':
                    elements.push(this.handleDegree(reader, DEGREE_SUPERSCRIPT));
                    break;
                case 'msubsup':
                    elements.push(this.handleSubSup(reader, DEGREE_SubSup));
                    break;
                // case 'mmultiscripts':
                // {
                //     let content = []
                //     let depth = reader.GetDepth();
                //     debugger
                //     while (reader.ReadNextSiblingNode(depth))
                //     {
                //         let current = AscWord.ParaMath.readMathMLContent(reader);
                //         if (reader.GetNameNoNS() === 'mprescripts')
                //             break;
                //         content.push(current);
                //         let post = AscWord.ParaMath.readMathMLContent(reader);
                //         console.log(current, post)
                //     }
                //     console.log(content);
                //     break;
                // }
                case 'munderover':
                    let content = []
                    let depth = reader.GetDepth();
                    while (reader.ReadNextSiblingNode(depth))
                    {
                        let current = this.readMathMLContent(reader)
                        content.push(current);
                    }
    
                    let textFirstContent = content[0].GetTextOfElement().GetText().trim()
                    if (AscMath.MathLiterals.nary.SearchU(textFirstContent))
                    {
                        elements.push(this.handleSubSup(reader, DEGREE_SubSup, content));
                    }
                    else
                    {
                        let Pr =  new CMathLimitPr();
                        Pr.type	= LIMIT_UP
                        Pr.content = [
                            content[0],
                            content[2],
                        ];
    
                        var Limit = new AscMath.Limit(Pr);
    
                        var MathContent = new AscCommonWord.CMathContent();
                        MathContent.addElementToContent(Limit);
                        MathContent.Correct_Content(true)
    
                        Pr.type	= LIMIT_LOW
                        Pr.content = [
                            MathContent,
                            content[1]
                        ];
                        let Limit2 = new CLimit(Pr);
    
                        // if (munderoverAttributes['accent'] === 'true')
                        // {
                        //     content[2].Apply_MenuProps({Action: c_oMathMenuAction.IncreaseArgumentSize}, 0);
                        // }
                        // else if (munderoverAttributes['accentunder'] === 'true')
                        // {
                        //     content[1].Apply_MenuProps({Action: c_oMathMenuAction.IncreaseArgumentSize}, 0);
                        // }
                        elements.push(Limit2);
                    }
                    break;
                case 'mfrac':
                    elements.push(this.handleFraction(reader));
                    break;
                case 'msqrt':
                    this.proceedMathMLDefaultAttributes(
                        this.GetAttributes(reader),
                        elements,
                        this.handleRadical(reader)
                    );
                    break;
                case 'mroot':
                    elements.push(this.handleRadical(reader, true));
                    break;
                case 'mpadded':
                case 'mstyle':
                case 'mrow':
                case 'merror':
                case 'semantics':
                    elements = this.readMathMLMRow(reader);
                    break;
                case 'munder':
                {
                    let attributes = reader.GetAttributes(reader);
    
                    if (attributes['displaystyle'] === 'true')
                    {
                        elements.push(this.handleNary(reader));
                        break;
                    }
    
                    if (attributes['accentunder'] === 'true')
                        elements.push(this.handleAccent(reader, VJUST_TOP));
                    else
                        elements.push(this.handleGroupChr(reader));
    
                    break;
                }
                case 'mover':
                {
                    let attributes = this.GetAttributes(reader);
    
                    if (attributes['accent'] === 'true')
                        elements.push(this.handleAccent(reader, VJUST_BOT));
                    else
                        elements.push(this.handleGroupChr(reader, VJUST_TOP));
                    break;
                }
                case 'mtable':
                    elements.push(this.handleMatrix(reader));
                    break;
                // case 'mstack':
                // 	elements.push(AscMath.EqArray.fromMathML(reader, true));
                // 	break;
                case 'mtr':
                case 'mlabeledtr':
                    elements.push(this.handleMtr(reader));
                    break;
                case 'mtd':
                    elements.push(this.handleMtd(reader))
                    break;
                case 'mfenced':
                    elements.push(this.handleDelimiter(reader));
                    break;
                default:
                    return elements;
            }

            return elements;
        };
        this.readMathMLMRow = function(reader)
        {
            let result = []
            let depth = reader.GetDepth();
            while (reader.ReadNextSiblingNode(depth))
            {
                result = result.concat(this.readMathMLNode(reader));
            }
    
            let processedResult = [];
            let i = 0;
            while (i < result.length)
            {
                let current = result[i];
                let currentText = current instanceof ParaRun ? current.GetTextOfElement().GetText() : null;
                let isStretchy = !(current instanceof ParaRun && current.mathml_metadata && current.mathml_metadata.stretchy === false);

                if (currentText && isStretchy && (AscMath.MathLiterals.lBrackets.SearchU(currentText) || AscMath.MathLiterals.lrBrackets.SearchU(currentText)))
                {
                    let openBracket = currentText;
                    let bracketContent = [];
                    let j = i + 1;
                    let closeBracket = null;
                    
                    while (j < result.length)
                    {
                        let nextElement = result[j];
                        let nextText = nextElement instanceof ParaRun ? nextElement.GetTextOfElement().GetText() : null;
                        let nextIsStretchy = !(nextElement instanceof ParaRun && nextElement.mathml_metadata && nextElement.mathml_metadata.stretchy === false);

                        if (nextText && nextIsStretchy&& (AscMath.MathLiterals.rBrackets.SearchU(nextText) || AscMath.MathLiterals.lrBrackets.SearchU(nextText)))
                        {
                            closeBracket = nextText;
                            break;
                        }
                        else
                        {
                            bracketContent.push(nextElement);
                        }
                        j++;
                    }
                    
                    if (closeBracket)
                    {
                        processedResult.push(this.proceedMathMLImplicitDelimiter(bracketContent, openBracket, closeBracket)[0]);
                        i = j + 1;
                    }
                    else
                    {
                        let remainingContent = result.slice(i + 1);
                        processedResult.push(this.proceedMathMLImplicitDelimiter(remainingContent, openBracket, null)[0]);
                        break;
                    }
                }
                else if (currentText && isStretchy && AscMath.MathLiterals.rBrackets.SearchU(currentText))
                {
                    let precedingContent = [];
                    let k = processedResult.length - 1;

                    while (k >= 0)
                    {
                        let prevElement = processedResult[k];
                        if (prevElement instanceof AscMath.Delimiter)
                        {
                            break;
                        }
                        precedingContent.unshift(processedResult.pop());
                        k--;
                    }

                    if (precedingContent.length > 0)
                    {
                        processedResult.push(this.proceedMathMLImplicitDelimiter(precedingContent, null, currentText)[0]);
                    }
                    else
                    {
                        processedResult.push(current);
                    }
                    i++;
                }
                else if (current instanceof AscMath.Nary)
                {
                    // N-ary operator found - check if it has empty content
                    // If so, collect following inline elements as its argument
                    let naryContent = current.getBase();
                    let isEmpty = !naryContent || naryContent.Content.length === 0;

                    if (isEmpty && i + 1 < result.length)
                    {
                        let argContent = new AscCommonWord.CMathContent();
                        let j = i + 1;

                        while (j < result.length)
                        {
                            let nextElement = result[j];

                            if (nextElement instanceof AscMath.Nary ||
                                nextElement instanceof AscMath.Delimiter ||
                                nextElement instanceof AscMath.Fraction ||
                                nextElement instanceof AscMath.Radical ||
                                nextElement instanceof AscMath.Matrix)
                            {
                                this.addFlatToMathContent(naryContent, nextElement);
                                j++;
                                break;
                            }

                            this.addFlatToMathContent(naryContent, nextElement);
                            j++;
                        }

                        argContent.Correct_Content(true);
                        processedResult.push(current);
                        i = j;
                    }
                    else
                    {
                        processedResult.push(current);
                        i++;
                    }
                }
                else
                {
                    processedResult.push(current);
                    i++;
                }
            }

            if (processedResult.length >= 3)
            {
                for (let j = processedResult.length; j >= 0; j--)
                {
                    let curent = processedResult[j-1];
                    let funcapply = processedResult[j-2];
                    let funcName = processedResult[j-3];

                    if (curent 
                        && funcName 
                        && funcapply 
                        && AscMath.MathLiterals.invisible.SearchU(funcapply.GetTextOfElement().GetText())
                        && funcName instanceof ParaRun)
                    {
                        let MathFunc = new CMathFunc({});

                        let MathContent = MathFunc.getFName();
                        MathContent.Add_Text(funcName.GetTextOfElement().GetText());

                        MathContent = MathFunc.getArgument();
                        MathContent.Add_Element(curent);

                        processedResult.splice(j-3, j, MathFunc);
                        break;
                    }
                }
            }
            return processedResult;
        };
        this.checkLinebreak = function (element)
        {
            for (let i = 0; i < this.mathMLData['mo-linebreak-todo'].length; i++)
            {
                let current = this.mathMLData['mo-linebreak-todo'][i];
                let el = current.element;
                if (element === el)
                {
                    this.mathMLData['mo-linebreak-todo'].splice(i, 1);
                    let element =  this.mathMLData['mo-id'][current.indenttarget];
                    return element.private_GetPosInParent();
                    // todo -> to operators
                }
            }
            return null;
        };
        this.createContentFromText = function(text)
        {
            let mathContent = new CMathContent();
            let run = new AscWord.Run(null, true);
            run.AddText(text);
            mathContent.addElementToContent(run);
            return mathContent;
        };
        this.readMathMLContent = function(reader)
        {
            let mathContent = new CMathContent();
            mathContent.mathml_metadata = {};
            let elements = this.readMathMLNode(reader, mathContent);
            
            for (let i = 0; i < elements.length; ++i)
            {
                let current = elements[i];
                this.addFlatToMathContent(mathContent, current);

                let breakPos = this.checkLinebreak(current)
                if (breakPos !== null)
                {
                    current.Set_MathForcedBreak(true, undefined);
                }
            }
            
            mathContent.Correct_Content(true);
            return mathContent;
        };
        this.readMathMLContentOnLevel = function (reader)
        {
            let content = [];
            let mathContent = new CMathContent();
            let depth = reader.GetDepth();
    
            while (reader.ReadNextSiblingNode(depth))
            {
                content = content.concat(this.readMathMLNode(reader));
            }
    
            for (let i = 0; i < content.length; i++)
            {
                this.addFlatToMathContent(mathContent, content[i]);
            }
    
            if (content.length === 0)
                return;

            mathContent.Correct_Content(true);
            return mathContent;
        };
        this.handleAccent = function(reader, type)
        {
            let props = new CMathAccentPr();
            props.content = [];
            let isNary = false;

            function check(str)
            {
                for (let oIter = str.getUnicodeIterator(); oIter.check(); oIter.next())
                {
                    if (AscMath.MathLiterals.nary.SearchU(String.fromCodePoint(oIter.value())))
                        return true;
                }
                return false;
            }

            let mContents = [];
            let depth = reader.GetDepth();
            while (reader.ReadNextSiblingNode(depth))
            {
                let current = this.readMathMLContent(reader);
                let base = current.GetTextOfElement().GetText();

                if (mContents.length === 0 && check(base))
                    isNary = true;

                mContents.push(current);
            }

            if (isNary)
            {
                if (mContents.length > 2)
                {
                    return this.handleSubSup(reader, mContents);
                }
                else
                {
                    return this.handleDegree(
                        reader,
                        type ===  VJUST_TOP ? DEGREE_SUBSCRIPT : DEGREE_SUPERSCRIPT,
                        mContents
                    );
                }
            }

            if (mContents.length >= 2)
            {
                props.content[0] = mContents[0];
                if (mContents[1])
                {
                    let chrText = mContents[1].GetTextOfElement().GetText().trim();
                    if (chrText.length > 1 || !AscMath.MathLiterals.accent.SearchU(chrText))
                    {
                        return this.handleGroupChr(
                            reader,
                            type === VJUST_BOT ? VJUST_TOP: VJUST_BOT,
                            mContents
                        )
                    }
                    props.chr = chrText.charCodeAt(0);
                }
            }
            else
            {
                props.content[0] = mContents[0];
            }

            return new CAccent(props);
        };
        this.handleBorderBox = function(reader)
        {
            let props = new CMathBorderBoxPr();
            props.content = [];
            let mathContent = new CMathContent();
            props.content.push(mathContent);

            let depth = reader.GetDepth();
            while (reader.ReadNextSiblingNode(depth))
            {
                let elements = this.readMathMLNode(reader);
                for (let i = 0; i < elements.length; i++)
                {
                    this.addFlatToMathContent(props.content[0], elements[i]);
                }
            }

            return new CBorderBox(props);
        };
        this.handlePhantom = function (reader)
        {
            let props       = new CMathPhantomPr();
            props.show      = false;
            props.transp    = true;
            props.content   = [this.readMathMLContentOnLevel(reader)];

            return new CPhantom(props);
        };
        this.handleDegree = function (reader, type, content)
        {
            let depth = reader.GetDepth();

            let props = new CMathDegreePr();
            props.type = undefined !== type ? type : DEGREE_SUPERSCRIPT;
            props.content = content ? content : [];

            while (reader.ReadNextSiblingNode(depth))
            {
                let current = this.readMathMLContent(reader);
                props.content.push(current);
            }

            return new AscMath.Degree(props);
        };
        this.handleSubSup = function (reader, type, content)
        {
            let depth = reader.GetDepth();

            let props = new CMathDegreeSubSupPr();
            props.type = undefined !== type ? type : DEGREE_PreSubSup;
            props.content = content ? content : [];

            if (!content)
            {
                while (reader.ReadNextSiblingNode(depth))
                {
                    let current = this.readMathMLContent(reader);
                    props.content.push(current);
                }
            }

            let textOfBase = props.content[0].GetTextOfElement().GetText();
            if (AscMath.MathLiterals.nary.SearchU(textOfBase))
            {
                return this.handleNaryFromSubSup(props, false, false);
            }

            if (props.content.length === 3)
            {
                let temp = props.content[1];
                props.content[1] = props.content[2];
                props.content[2] = temp;
            }

            return new AscMath.DegreeSubSup(props);
        };
        this.handleFraction = function (reader)
        {
            let attributes = this.GetAttributes(reader);
            let props = new CMathFractionPr();
            props.content = [];

            if (attributes['bevelled'])
                props.type = SKEWED_FRACTION;

            let depth = reader.GetDepth();
            while (reader.ReadNextSiblingNode(depth))
            {
                props.content.push(this.readMathMLContent(reader));
            }
            
            return new CFraction(props);
        };
        this.handleLimit = function(reader, type, content) {
            let props = new CMathLimitPr();
            props.content = content ? content : content;
            props.type = type;

            let mContents = [];
            let depth = reader.GetDepth();
            while (reader.ReadNextSiblingNode(depth))
            {
                mContents.push(this.readMathMLContent(reader));
            }

            if (mContents.length)
            {
                if (mContents.length >= 2)
                {
                    props.content[0] = mContents[0];
                    props.content[1] = mContents[1];
                }
                else
                {
                    props.content[0] = mContents[0];
                }
            }

            return new CLimit(props);
        };
        this.getGlobalAttributes = function(attributes, paraMath)
        {
            if (!paraMath)
                return;

            paraMath.mathml_metadata = {};

            if (!attributes['display'])
                attributes['display'] = 'inline'; // default for mathMl

            if (attributes['display'] === 'block')
                paraMath.mathml_metadata['display'] = true;
            else
                paraMath.mathml_metadata['display'] = false;
        };
        this.handleMathContent = function(reader, paraMath, mathContent)
        {
            let depth = reader.GetDepth();
            this.getGlobalAttributes(this.GetAttributes(reader), paraMath);

            while (reader.ReadNextSiblingNode(depth))
            {
                let elements = this.readMathMLNode(reader);
                for (let i = 0; i < elements.length; ++i)
                {
                    let current = elements[i];
                    if (typeof current === 'string')
                        continue;

                    this.addFlatToMathContent(mathContent, current);

                    let breakPos = this.checkLinebreak(current);
                    if (breakPos !== null)
                        current.Set_MathForcedBreak(true, undefined);
                }
            }
        };
        this.normalizeMtableStructure = function(reader)
        {
            // Collect all rows, normalizing the structure mtable → mtr → mtd
            let normalizedRows = [];
            let depth = reader.GetDepth();

            while (reader.ReadNextSiblingNode(depth))
            {
                let name = reader.GetNameNoNS();

                if (name === "mtr" || name === "mlabeledtr")
                {
                    // Normalize mtr content
                    let nestedRows = [];
                    let cells = this.normalizeMtrContent(reader, nestedRows);

                    // If there are nested rows, add only them (parent row is empty)
                    if (nestedRows.length > 0)
                    {
                        normalizedRows = normalizedRows.concat(nestedRows);
                    }
                    else
                    {
                        // No nested rows - add current row
                        normalizedRows.push(cells);
                    }
                }
                else if (name === "mtd")
                {
                    // mtd directly in mtable - check for nested mtd
                    let nestedCells = this.checkForNestedMtd(reader);
                    if (nestedCells.length > 0)
                    {
                        // Has nested mtd - create row from them
                        normalizedRows.push(nestedCells);
                    }
                    else
                    {
                        // No nested mtd - process as regular cell
                        let cellContent = this.handleMtd(reader);
                        if (cellContent)
                            normalizedRows.push([cellContent]);
                    }
                }
                else
                {
                    // Any other element - wrap in mtr → mtd
                    let elements = this.readMathMLNode(reader);
                    if (elements.length > 0)
                    {
                        let mathContent = new CMathContent();
                        for (let i = 0; i < elements.length; i++)
                        {
                            mathContent.addElementToContent(elements[i]);
                        }
                        mathContent.Correct_Content(true);
                        normalizedRows.push([mathContent]);
                    }
                }
            }

            return normalizedRows;
        };

        this.checkForNestedMtd = function(reader)
        {
            // Check if mtd contains nested mtd elements
            let cells = [];
            let depth = reader.GetDepth();
            let hasNestedMtd = false;

            while (reader.ReadNextSiblingNode(depth))
            {
                let name = reader.GetNameNoNS();

                if (name === "mtd")
                {
                    hasNestedMtd = true;
                    let cellContent = this.handleMtd(reader);

                    if (cellContent)
                        cells.push(cellContent);
                }
            }

            // Return cells only if nested mtd were found
            return hasNestedMtd ? cells : [];
        };

        this.normalizeMtrContent = function(reader, collectNestedRows)
        {
            // Process mtr content, ensuring all children are mtd
            let cells = [];
            let depth = reader.GetDepth();
            let nonMtdElements = [];
            let nestedRows = collectNestedRows || [];

            while (reader.ReadNextSiblingNode(depth))
            {
                let name = reader.GetNameNoNS();

                if (name === "mtr" || name === "mlabeledtr")
                {
                    // Nested mtr - incorrect structure
                    // Collect nested rows to add at table level
                    let nestedCells = this.normalizeMtrContent(reader, nestedRows);
                    if (nestedCells.length > 0)
                    {
                        nestedRows.push(nestedCells);
                    }
                }
                else if (name === "mtd")
                {
                    // Valid mtd
                    let cellContent = this.handleMtd(reader);

                    // If non-mtd elements accumulated, wrap them in mtd
                    if (nonMtdElements.length > 0)
                    {
                        let mathContent = new CMathContent();
                        for (let i = 0; i < nonMtdElements.length; i++)
                        {
                            this.addFlatToMathContent(mathContent, nonMtdElements[i]);
                        }
                        mathContent.Correct_Content(true);
                        cells.push(mathContent);
                        nonMtdElements = [];
                    }

                    if (cellContent)
                        cells.push(cellContent);
                }
                else
                {
                    // Non-mtd element inside mtr - accumulate for wrapping
                    let elements = this.readMathMLNode(reader);
                    nonMtdElements = nonMtdElements.concat(elements);
                }
            }

            // If non-mtd elements remain, wrap them in final cell
            if (nonMtdElements.length > 0)
            {
                let mathContent = new CMathContent();
                for (let i = 0; i < nonMtdElements.length; i++)
                {
                    this.addFlatToMathContent(mathContent, nonMtdElements[i]);
                }
                mathContent.Correct_Content(true);
                cells.push(mathContent);
            }

            // If no cells at all, create empty one
            if (cells.length === 0)
            {
                cells.push(new CMathContent());
            }

            return cells;
        };

        this.handleMatrix = function(reader)
        {
            let props = new CMathMatrixPr();
            props.mrs = this.normalizeMtableStructure(reader);

            return new CMathMatrix(props);
        };
        this.handleMtr = function(reader)
        {
            let depth = reader.GetDepth();
            while (reader.ReadNextSiblingNode(depth))
            {
                let name = reader.GetNameNoNS();
                if (name === "mtd")
                    return this.readMathMLContent(reader);
            }
        };
        this.handleMtd = function(reader)
        {
            return this.readMathMLContentOnLevel(reader);
        };
        this.handleEqArray = function(reader) {
            let props = new CMathEqArrPr();
            props.content = [];

            let depth = reader.GetDepth();
            while (reader.ReadNextSiblingNode(depth))
            {
                let name = reader.GetNameNoNS();
                if (name !== "msline") //skip msline for now
                    props.content.push(this.readMathMLContent(reader));
            }

            return new CEqArray(props);
        };
        this.handleNary = function(reader, content, type)
        {
            let props = new CMathNaryPr();
            props.content = content ? content : [];
            props.type = NARY_SubSup;

            props.supHide = type ? type.supHide : false;
            props.subHide = type ? type.subHide : false;

            let depth = reader.GetDepth();
            while (reader.ReadNextSiblingNode(depth))
            {
                props.content.push(this.readMathMLContent(reader));
            }

            let firstContent = props.content.shift();
            let charText = firstContent.GetTextOfElement().GetText();
            props.chr = charText.charCodeAt(0);

            return this.private_fromMathML(props, true);
        };
        this.handleNary_private = function(props, isAlredyChr)
        {
            if (!isAlredyChr)
            {
                let firstContent = props.content.shift();
                let charText = firstContent.GetTextOfElement().GetText();
                props.chr = charText.charCodeAt(0);
            }
            return new CNary(props);
        };
        this.handleNaryFromSubSup = function(props, isSub, isSup)
        {
            let propsNary = new CMathNaryPr();
            propsNary.limLoc = NARY_UndOvr;
            propsNary.supHide = isSub;
            propsNary.subHide = isSup;
            propsNary.content = props.content;

            if (props.content[0])
            {
                if (props.content[0].mathml_metadata['movablelimits'] === 'true')
                    propsNary.limLoc = NARY_SubSup;
            }

            let nary = this.handleNary_private(propsNary);
            return nary;
        };
        this.handleDelimiter = function(reader, props)
        {
            if (!props)
            {
                let attributes = this.GetAttributes(reader);
                let open = attributes['open'] || "(";
                let close = attributes['close'] || ")";
                let separator = attributes['separators'] || ",";

                props = new CMathDelimiterPr();
                props.begChr		= open.trim().charCodeAt(0);
                props.endChr		= close.trim().charCodeAt(0);
                props.sepChr		= separator[0].trim().charCodeAt(0);
                props.content		= [];
                props.shp			= DELIMITER_SHAPE_MATCH;

                let depth = reader.GetDepth();
                while (reader.ReadNextSiblingNode(depth))
                {
                    props.content.push(this.readMathMLContent(reader));
                }
            }

            return new CDelimiter(props);
        };
        this.handleGroupChr = function(reader, type, content)
        {
            let props = new CMathGroupChrPr();
            props.content = content ? content : [];
            props.pos = type;
            props.vertJc = (type === VJUST_TOP ) ? VJUST_BOT : undefined;
        
            let mContents = [];
            let depth = reader.GetDepth();
        
            if (!content)
            {
                while (reader.ReadNextSiblingNode(depth))
                {
                    mContents.push(this.readMathMLContent(reader));
                }
            }
            else
            {
                mContents = content;
            }
        
            if (mContents.length >= 2 || !content)
            {
                props.content.push(mContents[0]);
        
                if (mContents[1])
                {
                    let chrText = mContents[1].GetTextOfElement().GetText().trim();
                    if (chrText.length > 1 || !AscMath.MathLiterals.hbrack.SearchU(chrText))
                        return this.handleLimit(reader, type === 0 ? 1 : type, mContents);
        
                    props.chr = chrText.charCodeAt(0);
                }
            }
            else if (!content)
            {
                props.content[0] = mContents[0];
            }
        
            return new CGroupCharacter(props);
        }
        this.handleRadical = function(reader, isRoot)
        {
            let props = new CMathRadicalPr();
            props.content = [];

            if (isRoot)
            {
                let mContents = [];
                let depth = reader.GetDepth();
                while (reader.ReadNextSiblingNode(depth))
                {
                    mContents.push(this.readMathMLContent(reader));
                }
                
                if (mContents.length === 2)
                {
                    props.content[0] = mContents[1];
                    props.content[1] = mContents[0];
                }
                else
                {
                    let mathContent = new CMathContent();
                    for (let i = 0; i < mContents.length; i++)
                    {
                        let currentMathContent = mContents[i];
                        for (let j = 0; j < currentMathContent.Content.length; j++)
                        {
                            let curData = mContents
                        }
                        this.addFlatToMathContent(mathContent, mContents[i]);
                    }
                    mathContent.Correct_Content(true);
                    props.content[1] = mathContent;
                }
            }
            else
            {
                let depth = reader.GetDepth();
                let mathContent = new CMathContent();
                while (reader.ReadNextSiblingNode(depth))
                {
                    let arrData = this.readMathMLNode(reader);
                    for (let j = 0; j < arrData.length; j++)
                    {
                        this.addFlatToMathContent(mathContent, arrData[j])
                    }
                }
                mathContent.Correct_Content(true);
                props.content[1] = mathContent
                props.degHide = true;
            }
            return new CRadical(props);
        };
    }
	//--------------------------------------------------------export----------------------------------------------------
	AscMath.MathML = new MathML();
})();

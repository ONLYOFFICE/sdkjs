'use strict';

/*
new function() {
    var parser = new EasySAXParser();

    parser.ns('rss', { // or false
        'http://search.yahoo.com/mrss/': 'media',
        'http://www.w3.org/1999/xhtml': 'xhtml',
        'http://www.w3.org/2005/Atom': 'atom',
        'http://purl.org/rss/1.0/': 'rss',
    });

    parser.on('error', function(msgError) {
    });

    parser.on('startNode', function(nodeName, getAttr, isTagEnd, getStrNode) {
        var attr = getAttr();
    });

    parser.on('endNode', function(nodeName, isTagStart, getStrNode) {
    });

    parser.on('textNode', function(text) {
    });

    parser.on('cdata', function(data) {
    });


    parser.on('comment', function(text) {
        //console.log('--'+text+'--')
    });

    //parser.on('unknownNS', function(key) {console.log('unknownNS: ' + key)});
    //parser.on('question', function() {}); // <? ... ?>
    //parser.on('attention', function() {}); // <!XXXXX zzzz="eeee">


    parser.write(stringChunk);
    parser.write(stringChunk);
    ...
    parser.end(stringChunk);
};

*/

// << ------------------------------------------------------------------------ >> //

EasySAXParser.entityDecode = xmlEntityDecode;
if (typeof module === 'object') {
    module.exports = EasySAXParser;
};


var stringFromCharCode = String.fromCharCode;
var objectCreate = Object.create;
function NULL_FUNC() {};


function entity2char(x) {
    if (x === 'amp') {
        return '&';
    };

    switch(x.toLocaleLowerCase()) {
        case 'quot': return '"';
        case 'amp': return '&'
        case 'lt': return '<'
        case 'gt': return '>'

        case 'plusmn': return '\u00B1';
        case 'laquo': return '\u00AB';
        case 'raquo': return '\u00BB';
        case 'micro': return '\u00B5';
        case 'nbsp': return '\u00A0';
        case 'copy': return '\u00A9';
        case 'sup2': return '\u00B2';
        case 'sup3': return '\u00B3';
        case 'para': return '\u00B6';
        case 'reg': return '\u00AE';
        case 'deg': return '\u00B0';
        case 'apos': return '\'';
    };

    return '&' + x + ';';
};

function replaceEntities(s, d, x, z) {
    if (z) {
        return entity2char(z);
    };

    if (d) {
        return stringFromCharCode(d);
    };

    return stringFromCharCode(parseInt(x, 16));
};

function xmlEntityDecode(s) {
    var s = ('' + s);

    if (s.length > 3 && s.indexOf('&') !== -1) {
        if (s.indexOf('&lt;') !== -1) {s = s.replace(/&lt;/g, '<');}
        if (s.indexOf('&gt;') !== -1) {s = s.replace(/&gt;/g, '>');}
        if (s.indexOf('&quot;') !== -1) {s = s.replace(/&quot;/g, '"');}

        if (s.indexOf('&') !== -1) {
            s = s.replace(/&#(\d+);|&#x([0123456789abcdef]+);|&(\w+);/ig, replaceEntities);
        };
    };

    return s;
};

function cloneMatrixNS(nsmatrix) {
    var nn = objectCreate(null);
    for (var n in nsmatrix) {
        nn[n] = nsmatrix[n];
    };
    return nn;
};

var EasySAXEvent = {
    Unknown: 0,
    START_ELEMENT: 1,
    CHARACTERS: 2,
    END_ELEMENT: 3,
    CDATA: 4,
    Comment: 5,
    Attention: 6,
    Question: 7
}

function EasySAXParser(config) {
    if (!this) {
        return null;
    };

    var onTextNode = NULL_FUNC, onStartNode = NULL_FUNC, onEndNode = NULL_FUNC, onCDATA = NULL_FUNC, onError = NULL_FUNC, onComment, onQuestion, onAttention, onUnknownNS;
    var is_onComment = false, is_onQuestion = false, is_onAttention = false, is_onUnknownNS = false;

    var isAutoEntity = true; // делать "EntityDecode" всегда
    var indexStartXML; // позиция на которой закончен разбор xml
    var entityDecode = xmlEntityDecode;
    var isNamespace = false;
    var returnError;
    var isParseStop; // прервать парсер
    var defaultNS;
    var nsmatrix = null;
    var useNS;
    var init = false;
    var xml; // string

    var stringNodePosStart; // number. для получения исходной строки узла
    var stringNodePosEnd; // number. для получения исходной строки узла
    var attrStartPos; // number начало позиции атрибутов в строке attrString <(div^ class="xxxx" title="sssss")/>
    var attrString; // строка атрибутов <(div class="xxxx" title="sssss")/>
    var attrRes; // закешированный результат разбора атрибутов , null - разбор не проводился, object - хеш атрибутов, true - нет атрибутов, false - невалидный xml

    function reset() {
        if (isNamespace) {
            nsmatrix = objectCreate(null);
            nsmatrix.xmlns = defaultNS;
        };

        indexStartXML = 0;
        returnError = '';
        isParseStop = false;
        xml = '';
    };

    this.setup = function (op) {
        for (var name in op) {
            switch(name) {
                case 'entityDecode': entityDecode = op.entityDecode || entityDecode; break;
                case 'autoEntity': isAutoEntity = !!op.autoEntity; break;
                case 'defaultNS': defaultNS = op.defaultNS || null; break;
                case 'ns': useNS = op.ns || null; break;
                case 'on':
                    var listeners = op.on;
                    for (var ev in listeners) {
                        this.on(ev, listeners[ev]);
                    };
                break;
            };
        };

        isNamespace = !!defaultNS && !!useNS;
    };

    this.on = function(name, cb) {
        // if (typeof cb !== 'function') {
        //     if (cb !== null) {
        //         throw error('required args on(string, function||null)');
        //     };
        // };

        switch(name) {
            case 'startNode': onStartNode = cb || NULL_FUNC; break;
            case 'endNode': onEndNode = cb || NULL_FUNC; break;
            case 'text': case 'textNode': onTextNode = cb || NULL_FUNC; break;
            case 'error': onError = cb || NULL_FUNC; break;
            case 'cdata': onCDATA = cb || NULL_FUNC; break;

            case 'unknownNS': onUnknownNS = cb; is_onUnknownNS = !!cb; break;
            case 'attention': onAttention = cb; is_onAttention = !!cb; break; // <!XXXXX zzzz="eeee">
            case 'question': onQuestion = cb; is_onQuestion = !!cb; break; // <? ....  ?>
            case 'comment': onComment = cb; is_onComment = !!cb; break;
        };
    };

    this.ns = function(root, ns) {
        if (!root) {
            isNamespace = false;
            defaultNS = null;
            useNS = null;
            return this;
        };

        if (!ns || typeof root !== 'string') {
            throw Error('required args ns(string, object)');
        };

        isNamespace = !!(useNS = ns || null);
        defaultNS = root || null;

        return this;
    };

    this.write = function(chunk) {
        if (typeof chunk !== 'string' || isParseStop) {
            return;
        };

        if (!init) {
            init = true;
            reset();
        };

        xml = xml ? xml + chunk : chunk;
        parse();

        if (isParseStop && returnError) {
            if (returnError) {
                onError(returnError);
                returnError = '';
            };
        };

        if (indexStartXML > 0) {
            xml = xml.substring(indexStartXML);
            indexStartXML = 0;
        };

        return this;
    };

    this.end = function() {
        if (returnError) {
            onError(returnError);
            returnError = '';
        };

        attrString = '';
        init = false;
        xml = '';
    };

    this.parse = function(xml) {
        this.write(xml);
        this.end();
    };
    this.staxParseXml = function(xml) {
        staxParseWithEvents.call(this, xml);
    };

    this.stop = function() {
        isParseStop = true;
    };

    if (config) {
        this.setup(config);
    };

    // -----------------------------------------------------

    var nodeParseAttrResult; // null - кеш пустой, true - атрибутов нет, {...} - карта атрибутов
    var nodeParseAttrSize = 0; // число элементов nodeParseAttrMap
    var nodeParseAttrMap = ['','','','','','','','','','']; // карта атрибутов. четные индексы "имя", не четные "значение"
    var nodeParseHasNS = false;
    var nodeParseName; // имя ноды

    // разбор ноды <nodeName ...> или <nodeName .../>
    // на вход indexStart = xml.indexOf('<');
    // return xml.indexOf('>', ixNameStart);
    function parseNode(indexStart) {
        var ixNameStart = +indexStart + 1; // позиция первого сивола имени
        var ixNameEnd; // позиция последнего + 1 сивола имени
        var attrName;

        var i = ixNameStart;
        var l = xml.length;
        var w;

        var iE = xml.indexOf('>', ixNameStart);
        var iR;

        if (iE === -1) { // не полный xml. дальнейший парсинг бессмыслен
            returnError = '#1901 invalid node'; // не полный xml
            return -1;
        };

        nodeParseAttrResult = null;
        nodeParseAttrSize = 0;
        nodeParseHasNS = false;
        nodeParseName = '';

        if (i >= l) {
            returnError = '#4952 invalid node'; // не полный xml
            return -1;
        };

        w = xml.charCodeAt(i);
        if (!(w > 96  && w < 123 || w > 64 && w < 91 || w === 95 || w === 58)) { // char 95"_" 58":"
            returnError = '#4940 first char <nodeName .../>';
            isParseStop = true; // дальнейший разбор невозможен
            return -1;
        };

        while(true) {
            if (++i >= l) {
                returnError = '#4950 invalid node'; // не полный xml
                return -1; // errorParse
            };

            w = xml.charCodeAt(i);

            if (w > 96 && w < 123 || w > 64 && w < 91 || w > 47 && w < 59 || w === 45 || w === 46 || w === 95) {
                continue; // символы имени тега только латиница
            };

            if (w === 32 || w === 9 || w === 10 || w === 11 || w === 12 || w === 13) { // \f\n\r\t\v пробел
                nodeParseName = xml.substring(ixNameStart, ixNameEnd = i);
                break;
            };

            if (w === 62 /* ">" */) { // тег закрылся, атрибутов нет
                nodeParseName = xml.substring(ixNameStart, ixNameEnd = i);
                return i;
            };

            if (w === 47 /* "/" */) {
                ixNameEnd = i;
                w = xml.charCodeAt(++i);

                if (w === 62) {
                    nodeParseName = xml.substring(ixNameStart, ixNameEnd);
                    return i;
                };
                returnError = '#0320 invalid node .../>?';
                isParseStop = true; // дальнейший разбор невозможен
                return -1;
            };

            returnError = '#5347 invalid nodeName';
            isParseStop = true; // дальнейший разбор невозможен
            return -1;
        };

        i += 1; // первый сивол пробел его пропускаем

        while(true) {
            iR = xml.indexOf('=', i);

            if (iR > iE || iR === -1) {
                break;
            };

            attrName = xml.substring(i, iR);

            if (isNamespace) {
                attrName = attrName.trim();
                if (attrName.charCodeAt(0) === 120 && attrName.substr(0, 6) === 'xmlns') {
                    nodeParseHasNS = true;
                };
            };

            nodeParseAttrMap[nodeParseAttrSize++] = attrName; // имя атрибута

            w = xml.charCodeAt(++iR);

            while(w === 32 || w === 9 || w === 10 || w === 11 || w === 12 || w === 13) { // \f\n\r\t\v
                w = xml.charCodeAt(++iR);
            };

            if (w === 34) {
                i = xml.indexOf('"', iR + 1);
            } else {
                i = xml.indexOf('\'', iR + 1);
            };

            if (i === -1) {
                returnError = '#5858 invalid node'; // не полный xml
                return -1;
            };

            nodeParseAttrMap[nodeParseAttrSize++] = xml.substring(iR + 1, i); // значение атрибута
            i += 1;

            if (i === iE) {
                break;
            };

            if (i > iE) {
                iE = xml.indexOf('>', i);
                if (iE === -1)  {
                    returnError = '#0901 invalid node'; // не полный xml
                    return -1;
                };
            };

            if (iE - i < 4) {
                break;
            };
        };

        return iE;
    };

    function upNSMATRIX() {
        var hasNewMatrix;
        var newalias;
        var alias;
        var value;
        var name;
        var j;

        if (!nodeParseAttrSize) {
            return;
        };

        for (j = 0; j < nodeParseAttrSize; j += 2) {
            name = nodeParseAttrMap[j];

            if (name !== 'xmlns') {
                if (name.charCodeAt(0) !== 120 || name.substr(0, 6) !== 'xmlns:') {
                    continue;
                };
                newalias = name.substr(6);
            } else {
                newalias = 'xmlns';
            };


            value = nodeParseAttrMap[j + 1];
            //alias = useNS[isAutoEntity ? value : entityDecode(value)];
            alias = useNS[entityDecode(value)];

            if (is_onUnknownNS && !alias) {
                alias = onUnknownNS(value);
            };

            if (alias) {
                if (nsmatrix[newalias] !== alias) {
                    if (!hasNewMatrix) {
                        nsmatrix = cloneMatrixNS(nsmatrix);
                        hasNewMatrix = true;
                    };
                    nsmatrix[newalias] = alias;
                };

                continue;
            };

            if (nsmatrix[newalias]) {
                if (!hasNewMatrix) {
                    nsmatrix = cloneMatrixNS(nsmatrix);
                    hasNewMatrix = true;
                };
                nsmatrix[newalias] = false;
            };
        };
    };

    function getAttrs() {
        if (nodeParseAttrResult !== null) {
            return nodeParseAttrResult;
        };

        if (nodeParseAttrSize === 0) {
            return nodeParseAttrResult = true;
        };

        var xmlnsAlias;
        var nsName;
        var iQ;

        var attrs = {};
        var value;
        var name;
        var j;


        if (isNamespace) {
            xmlnsAlias = nsmatrix.xmlns;
        };

        for (j = 0; j < nodeParseAttrSize; j++) {
            name = isNamespace ? nodeParseAttrMap[j] : nodeParseAttrMap[j].trim();

            if (isNamespace) {
                iQ = name.indexOf(':');
                if (iQ !== -1) {
                    nsName = nsmatrix[name.substring(0, iQ)];
                    if (!nsName || nsName === 'xmlns') {
                        continue;
                    };
                    name = xmlnsAlias !== nsName ? nsName + name.substr(iQ) : name.substr(iQ + 1);
                };
            };

            value = nodeParseAttrMap[++j];
            if (isAutoEntity) {
                value = entityDecode(value);
            };

            attrs[name] = value;
        };

        return nodeParseAttrResult = attrs;
    };

    function getStringNode() { // вернет исходную строку узла
        return xml.substring(stringNodePosStart, stringNodePosEnd);
    };


    var parseStackMatrixNS = [];
    var parseStackNodes = [];
    var stopIndexNS = 0;


    function parse() {
        // разбор идет по элементам (тег, текст cdata, ...).
        // элемент должен быть целиком в памяти

        var _nsmatrix;
        var isTagStart = false;
        var isTagEnd = false;
        //var nodeBody;
        var stopEmit; // используется при разборе "namespace" . если встретился неизвестное пространство то события не генерируются
        var nodeName;
        var xmlns;
        var iD;
        var iQ;
        var w;
        var i; // number

        returnError = null; // сброс ошибки неудачного разбора

        while(indexStartXML !== -1) {
            stopEmit = stopIndexNS > 0;

            // поиск начала тега
            if (xml.charCodeAt(indexStartXML) === 60) { // "<"
                i = indexStartXML;
            } else {
                i = xml.indexOf('<', indexStartXML);
            };

            if (i === -1) { // узел не найден. повторим попытку на след. write
                if (parseStackNodes.length) {
                    returnError = 'unexpected end parse';
                    return;
                };

                // --- нужно подумать как обрабатывать начало файла ---
                // if (indexStartXML === 0) { // разбор еше не начат. возможно это начало файла. мусор до первого тега игнор
                //     returnError = 'missing first tag';
                //     return;
                // };

                return;
            };

            if (indexStartXML !== i && !stopEmit) { // все что до тега это текст
                var text = xml.substring(indexStartXML, i);
                indexStartXML = i; // до этой позиции разбор завершен

                onTextNode(isAutoEntity ? entityDecode(text) : text);
                if (isParseStop) {
                    return;
                };
            };

            // ELEMENT
            // ---------------------------------------------

            w = xml.charCodeAt(i + 1);

            if (w === 33) { // 33 == "!"
                var wNext = xml.charCodeAt(i + 2);

                // CDATA
                // ---------------------------------------------
                if (wNext === 91 && xml.substr(i + 3, 6) === 'CDATA[') { // 91 == "["
                    var indexStartCDATA = i + 9;
                    var indexEndCDATA = xml.indexOf(']]>', indexStartCDATA);
                    if (indexEndCDATA === -1) {
                        returnError = 'cdata, not found ...]]>'; // не закрыт CDATA. повторим попытку на след. write
                        return;
                    };

                    indexStartXML = indexEndCDATA + 3;

                    if (!stopEmit) {
                        onCDATA(xml.substring(indexStartCDATA, indexEndCDATA));
                        if (isParseStop) {
                            return;
                        };
                    };
                    continue;
                };


                // COMMENT
                // ---------------------------------------------
                if (wNext === 45 && xml.charCodeAt(i + 3) === 45) { // 45 == "-"
                    var indexStartComment = i + 4;
                    var indexEndComment = xml.indexOf('-->', indexStartComment);
                    if (indexEndComment === -1) {
                        returnError = 'expected -->'; // не закрыт комментарий. повторим попытку на след. write
                        return;
                    };

                    indexStartXML = indexEndComment + 3;

                    if (is_onComment && !stopEmit) {
                        var commentText = xml.substring(indexStartComment, indexEndComment);
                        onComment(isAutoEntity ? entityDecode(commentText) : commentText);
                        if (isParseStop) {
                            return;
                        };
                    };
                    continue;
                };

                // ATTENTION
                // ---------------------------------------------
                {
                    var indexStartAttention = i + 1;
                    var indexEndAttention = xml.indexOf('>', indexStartAttention);
                    if (indexEndAttention === -1) {
                        returnError = 'expected attention ...>'; // повторим попытку на след. write
                        return;
                    };

                    indexStartXML = indexEndAttention + 1;

                    if (is_onAttention && !stopEmit) {
                        onAttention(xml.substring(i, indexStartXML)); // весь тег, так как не придумал api
                        if (isParseStop) {
                            return;
                        };
                    };
                };

                continue;
            };

            // QUESTION
            // ---------------------------------------------
            if (w === 63) { // "?"
                var indexEndQuestion = xml.indexOf('?>', i);
                if (indexEndQuestion === -1) { // error
                    returnError = 'expected question ...?>'; // повторим попытку на след. write
                    return;
                };

                indexStartXML = indexEndQuestion + 2;

                if (is_onQuestion) {
                    onQuestion(xml.substring(i, indexStartXML)); // весь тег, так как не придумал api
                    if (isParseStop) {
                        return;
                    };
                };
                continue;
            };


            // NODE ELEMENT
            // ---------------------------------------------

            if (w === 47) { // </...
                var indexEndNode = xml.indexOf('>', i + 1);
                if (indexEndNode === -1) { // error  ...> // не нашел знак закрытия тега
                    returnError = 'unclosed tag'; // повторим попытку на след. write
                    return;
                };

                isTagStart = false;
                isTagEnd = true;

                // проверяем что тег должен быть закрыт тот-же что и открывался
                if (!parseStackNodes.length) {
                    returnError = 'close tag, requires open tag';
                    isParseStop = true; // дальнейший разбор невозможен
                    return;
                };

                nodeName = parseStackNodes.pop();
                iQ = i + 2 + nodeName.length;

                if (nodeName !== xml.substring(i + 2, iQ)) {
                    returnError = 'close tag, not equal to the open tag';
                    isParseStop = true; // дальнейший разбор невозможен
                    return;
                };

                // проверим что в закрываюшем теге нет лишнего
                for(; iQ < indexEndNode; iQ++) {
                    var wNext = xml.charCodeAt(iQ);
                    if (wNext === 32 || wNext === 9 || wNext === 10 || wNext === 11 || wNext === 12 || wNext === 13) { // \f\n\r\t\v
                        continue;
                    };

                    returnError = 'close tag, unallowable char';
                    isParseStop = true; // дальнейший разбор невозможен
                    return;
                };

                indexStartXML = indexEndNode + 1;

            } else {
                var indexEndNode = parseNode(i);
                if (indexEndNode === -1) { // error  ...> // не нашел знак закрытия тега
                    returnError = returnError || 'unclosed tag'; // повторим попытку на след. write
                    return;
                };

                isTagStart = true;
                isTagEnd = xml.charCodeAt(indexEndNode - 1) === 47;
                nodeName = nodeParseName;

                if (!isTagEnd) {
                    parseStackNodes.push(nodeName);
                };

                indexStartXML = indexEndNode + 1;
            };


            if (isNamespace) {
                if (stopEmit) { // потомки неизвестного пространства имен
                    if (isTagEnd) {
                        if (!isTagStart) {
                            if (--stopIndexNS === 0) {
                                nsmatrix = parseStackMatrixNS.pop();
                            };
                        };

                    } else {
                        stopIndexNS += 1;
                    };
                    continue;
                };

                // добавляем в parseStackMatrixNS только если !isTagEnd, иначе сохраняем контекст пространств в переменной
                _nsmatrix = nsmatrix;
                if (!isTagEnd) {
                    parseStackMatrixNS.push(nsmatrix);
                };

                if (isTagStart && nodeParseHasNS) {  // есть подозрение на xmlns //  && (nodeParseAttrResult === null)
                    upNSMATRIX();
                };

                iD = nodeName.indexOf(':');
                if (iD !== -1) {
                    xmlns = nsmatrix[nodeName.substring(0, iD)];
                    nodeName = nodeName.substr(iD + 1);

                } else {
                    xmlns = nsmatrix.xmlns;
                };

                if (!xmlns) {
                    // элемент неизвестного пространства имен
                    if (isTagEnd) {
                        nsmatrix = _nsmatrix; // так как тут всегда isTagStart
                    } else {
                        stopIndexNS = 1; // первый элемент для которого не определено пространство имен
                    };
                    continue;
                };

                nodeName = xmlns + ':' + nodeName;
            };

            stringNodePosStart = i; // stringNodePosStart, stringNodePosEnd - для ручного разбора getStringNode()
            stringNodePosEnd = indexStartXML;

            if (isTagStart) {
                onStartNode(nodeName, getAttrs, isTagEnd, getStringNode);
                if (isParseStop) {
                    return;
                };
            };

            if (isTagEnd) {
                onEndNode(nodeName, isTagStart, getStringNode);
                if (isParseStop) {
                    return;
                };

                if (isNamespace) {
                    if (isTagStart) {
                        nsmatrix = _nsmatrix;
                    } else {
                        nsmatrix = parseStackMatrixNS.pop();
                    };
                };
            };
        };
    };

    var staxStream, staxStateisTagStart, staxStateisTagEnd, staxStatenodeName, staxStateeventType, staxStatetext, staxStatedepth;
    function staxInit(_xml) {
        init = true;
        reset();
        staxCleanState();

        var xmls = [_xml];
        staxStream = {read: function(){return xmls.pop();}};
        xml = staxStream.read();
    }
    function staxCleanState() {
        staxStateisTagStart = false;
        staxStateisTagEnd = false;
        staxStatenodeName = null;

        staxStateeventType = null;
        staxStatetext = null;
        staxStatedepth = 0;

        returnError = null; // сброс ошибки неудачного разбора
    }
    function staxHasNext() {
        return !isParseStop;
    }
    function staxGetEventType() {
        return staxStateeventType;
    }
    function staxGetName() {
        return staxStatenodeName;
    }
    function staxGetText() {
        return staxStatetext;
    }
    function staxGetDepth() {
        return staxStatedepth;
    }
    function staxIsEmptyNode() {
        return staxStateisTagStart && staxStateisTagEnd;
    }
    function staxGetError() {
        return returnError;
    }

    function staxParseWithEvents(xml) {
        this.staxInit(xml);

        while (staxHasNext()) {
            var eventType = staxNext();
            switch (eventType) {
                case EasySAXEvent.START_ELEMENT:
                    onStartNode(staxGetName(), getAttrs, staxStateisTagEnd, getStringNode);
                    break;
                case EasySAXEvent.CHARACTERS:
                    onTextNode(staxGetText());
                    break;
                case EasySAXEvent.END_ELEMENT:
                    onEndNode(staxGetName(), staxStateisTagStart, getStringNode);
                    break;
                case EasySAXEvent.CDATA:
                    onCDATA(staxGetText());
                    break;
                case EasySAXEvent.Comment:
                    if (is_onComment) {
                        onComment(staxGetText());
                    }
                    break;
                case EasySAXEvent.Attention:
                    if (is_onAttention) {
                        onAttention(staxGetText());
                    }
                    break;
                case EasySAXEvent.Question:
                    if (is_onQuestion) {
                        onQuestion(staxGetText());
                    }
                    break;
            }
        }
        this.staxEnd();
    }

    function staxNext() {
        // разбор идет по элементам (тег, текст cdata, ...).
        // элемент должен быть целиком в памяти

        // var _nsmatrix;
        // var isTagStart = false;
        // var isTagEnd = false;
        // //var nodeBody;
        // var stopEmit; // используется при разборе "namespace" . если встретился неизвестное пространство то события не генерируются
        // var nodeName;
        // var iD;
        var iQ;
        var w;
        var i; // number

        if (staxStateisTagEnd) {
            if (staxStateeventType === EasySAXEvent.START_ELEMENT) {
                staxStateeventType = EasySAXEvent.END_ELEMENT;
                return staxStateeventType;
            }
            staxStatedepth--;
        };

        staxStateeventType = EasySAXEvent.Unknown;

        // поиск начала тега
        if (xml.charCodeAt(indexStartXML) === 60) { // "<"
            i = indexStartXML;
        } else {
            i = xml.indexOf('<', indexStartXML);
        };

        if (i === -1) { // узел не найден. повторим попытку на след. write
            // --- нужно подумать как обрабатывать начало файла ---
            // if (indexStartXML === 0) { // разбор еше не начат. возможно это начало файла. мусор до первого тега игнор
            //     returnError = 'missing first tag';
            //     return;
            // };

            if (indexStartXML > 0) {
                xml = xml.substring(indexStartXML);
                indexStartXML = 0;
            };
            var chunk = staxStream.read();
            if(chunk) {
                xml = xml + chunk;

                // поиск начала тега
                if (xml.charCodeAt(indexStartXML) === 60) { // "<"
                    i = indexStartXML;
                } else {
                    i = xml.indexOf('<', indexStartXML);
                };
            }

            if (i === -1) {
                if (parseStackNodes.length) {
                    returnError = 'unexpected end parse';
                };
                if (returnError) {
                    onError(returnError);
                    returnError = '';
                };
                isParseStop = true;
                return staxStateeventType;
            }
        };

        if (indexStartXML !== i) { // все что до тега это текст
            var text = xml.substring(indexStartXML, i);
            indexStartXML = i; // до этой позиции разбор завершен

            staxStateeventType = EasySAXEvent.CHARACTERS;
            staxStatetext = isAutoEntity ? entityDecode(text) : text;
            return staxStateeventType;
        };

        // ELEMENT
        // ---------------------------------------------

        w = xml.charCodeAt(i + 1);

        if (w === 33) { // 33 == "!"
            var wNext = xml.charCodeAt(i + 2);

            // CDATA
            // ---------------------------------------------
            if (wNext === 91 && xml.substr(i + 3, 6) === 'CDATA[') { // 91 == "["
                var indexStartCDATA = i + 9;
                var indexEndCDATA = xml.indexOf(']]>', indexStartCDATA);
                if (indexEndCDATA === -1) {
                    returnError = 'cdata, not found ...]]>'; // не закрыт CDATA. повторим попытку на след. write
                    isParseStop = true;
                    return staxStateeventType;
                };

                indexStartXML = indexEndCDATA + 3;
                staxStateeventType = EasySAXEvent.CDATA;
                staxStatetext = xml.substring(indexStartCDATA, indexEndCDATA);
                return staxStateeventType;
            };


            // COMMENT
            // ---------------------------------------------
            if (wNext === 45 && xml.charCodeAt(i + 3) === 45) { // 45 == "-"
                var indexStartComment = i + 4;
                var indexEndComment = xml.indexOf('-->', indexStartComment);
                if (indexEndComment === -1) {
                    returnError = 'expected -->'; // не закрыт комментарий. повторим попытку на след. write
                    isParseStop = true;
                    return staxStateeventType;
                };

                indexStartXML = indexEndComment + 3;
                staxStateeventType = EasySAXEvent.Comment;
                staxStatetext = xml.substring(indexStartComment, indexEndComment);
                return staxStateeventType;
            };

            // ATTENTION
            // ---------------------------------------------
            {
                var indexStartAttention = i + 1;
                var indexEndAttention = xml.indexOf('>', indexStartAttention);
                if (indexEndAttention === -1) {
                    returnError = 'expected attention ...>'; // повторим попытку на след. write
                    isParseStop = true;
                    return staxStateeventType;
                };

                indexStartXML = indexEndAttention + 1;
                staxStateeventType = EasySAXEvent.Attention;
                staxStatetext = xml.substring(i, indexStartXML);
                return staxStateeventType;
            };

            return staxStateeventType;
        };

        // QUESTION
        // ---------------------------------------------
        if (w === 63) { // "?"
            var indexEndQuestion = xml.indexOf('?>', i);
            if (indexEndQuestion === -1) { // error
                returnError = 'expected question ...?>'; // повторим попытку на след. write
                isParseStop = true;
                return staxStateeventType;
            };

            indexStartXML = indexEndQuestion + 2;
            staxStateeventType = EasySAXEvent.Question;
            staxStatetext = xml.substring(i, indexStartXML);
            return staxStateeventType;
        };


        // NODE ELEMENT
        // ---------------------------------------------

        if (w === 47) { // </...
            var indexEndNode = xml.indexOf('>', i + 1);
            if (indexEndNode === -1) { // error  ...> // не нашел знак закрытия тега
                returnError = 'unclosed tag'; // повторим попытку на след. write
                isParseStop = true;
                return staxStateeventType;
            };

            staxStateisTagStart = false;
            staxStateisTagEnd = true;

            // проверяем что тег должен быть закрыт тот-же что и открывался
            if (!parseStackNodes.length) {
                returnError = 'close tag, requires open tag';
                isParseStop = true; // дальнейший разбор невозможен
                return staxStateeventType;
            };

            staxStatenodeName = parseStackNodes.pop();
            iQ = i + 2 + staxStatenodeName.length;

            if (staxStatenodeName !== xml.substring(i + 2, iQ)) {
                returnError = 'close tag, not equal to the open tag';
                isParseStop = true; // дальнейший разбор невозможен
                return staxStateeventType;
            };

            // проверим что в закрываюшем теге нет лишнего
            for(; iQ < indexEndNode; iQ++) {
                var wNext = xml.charCodeAt(iQ);
                if (wNext === 32 || wNext === 9 || wNext === 10 || wNext === 11 || wNext === 12 || wNext === 13) { // \f\n\r\t\v
                    continue;
                };

                returnError = 'close tag, unallowable char';
                isParseStop = true; // дальнейший разбор невозможен
                return staxStateeventType;
            };

            indexStartXML = indexEndNode + 1;

        } else {
            var indexEndNode = parseNode(i);
            if (indexEndNode === -1) { // error  ...> // не нашел знак закрытия тега
                returnError = returnError || 'unclosed tag'; // повторим попытку на след. write
                isParseStop = true;
                return staxStateeventType;
            };

            staxStateisTagStart = true;
            staxStateisTagEnd = xml.charCodeAt(indexEndNode - 1) === 47;
            staxStatenodeName = nodeParseName;

            if (!staxStateisTagEnd) {
                parseStackNodes.push(staxStatenodeName);
            };

            indexStartXML = indexEndNode + 1;
        };

        if (staxStateisTagStart) {
            staxStatedepth++;
            staxStateeventType = EasySAXEvent.START_ELEMENT;
            return staxStateeventType;
        };
        if (staxStateisTagEnd) {
            staxStateeventType = EasySAXEvent.END_ELEMENT;
            return staxStateeventType;
        };
    }
    
    this.staxInit = staxInit;
    this.staxHasNext = staxHasNext;
    this.staxNext = staxNext;
    this.staxGetEventType = staxGetEventType;
    this.staxGetName = staxGetName;
    this.staxGetAttrs = getAttrs;
    this.staxGetText = staxGetText;
    this.staxGetDepth = staxGetDepth;
    this.staxIsEmptyNode = staxIsEmptyNode;
    this.staxGetError = staxGetError;
    this.staxEnd = this.end;
};

function StaxParser(xml, rels, context) {
    this.xml = xml;
    this.index = 0;
    this.length = xml.length;
    this.rels = rels;
    this.context = context;

    this.isTagStart = null;
    this.isInAttr = false;
    this.isTagEnd = null;
    this.eventType = null;
    this.name = null;
    this.text = null;
    this.value = null;
    this.stop = false;
    this.depth = 0;
}
StaxParser.prototype.hasNext = function() {
    return !this.stop;
};
StaxParser.prototype.next = function() {
    if (this.isInAttr) {
        this.SkipAttributes();
    }
    if (this.isTagEnd) {
        if (this.eventType === EasySAXEvent.START_ELEMENT) {
            this.eventType = EasySAXEvent.END_ELEMENT;
            return this.eventType;
        }
        this.depth--;
    }

    this.eventType = EasySAXEvent.Unknown;
    this.isTagStart = null;
    this.isInAttr = false;
    this.isTagEnd = null;
    this.eventType = null;
    this.name = null;
    this.text = null;
    this.value = null;

    var i = this.index;
    if (this.xml.charCodeAt(i) !== 60) {//'<'
        i = this.xml.indexOf('<', i);
    }

    if (i === -1) {
        this.stop = true;
        return;
    }

    // Text
    if (this.index !== i) {
        this.eventType = EasySAXEvent.CHARACTERS;
        this.text = this.xml.substring(this.index, i);
        this.index = i;
        return this.eventType;
    }

    // ELEMENT
    // ---------------------------------------------
    var w = this.xml.charCodeAt(i + 1);
    if (w === 33) { // 33 == "!"
        this.index = this.xml.indexOf('>', i);
        return this.eventType;
    }
    // QUESTION
    // ---------------------------------------------
    if (w === 63) { // "?"
        this.index = this.xml.indexOf('>', i);
        return this.eventType;
    }
    // NODE ELEMENT
    // ---------------------------------------------
    var indexEndNode;
    if (w !== 47) { // </...
        indexEndNode = this.parseNode(i);
        if (indexEndNode === -1) { // error  ...> // не нашел знак закрытия тега
            this.stop = true;
            return this.eventType;
        }
        this.isTagStart = true;
        this.isInAttr = true;
        this.isTagEnd = false;
        // this.isTagEnd = this.xml.charCodeAt(indexEndNode - 1) === 47;
        this.eventType = EasySAXEvent.START_ELEMENT;
        this.depth++;
        this.index = indexEndNode;
    } else {
        //todo close tag name not used. don't parse for performance reason
        indexEndNode = this.xml.indexOf('>', i + 1);
        if (indexEndNode === -1) { // error  ...> // не нашел знак закрытия тега
            this.stop = true;
            return this.eventType;
        }
        this.isTagStart = false;
        this.isTagEnd = true;
        this.eventType = EasySAXEvent.END_ELEMENT;
        this.index = indexEndNode + 1;
    }
    return this.eventType;
};
StaxParser.prototype.parseNode = function(indexStart) {
    var ixNameStart = +indexStart + 1; // позиция первого сивола имени

    var i = ixNameStart;
    var w = this.xml.charCodeAt(i);
    while (i < this.length) {
        if (w > 96 && w < 123 || w > 64 && w < 91 || w > 47 && w < 59 || w === 45 || w === 46 || w === 95) {
            w = this.xml.charCodeAt(++i);
            continue;
        }
        if (w === 32 || w === 9 || w === 10 || w === 11 || w === 12 || w === 13) { // \f\n\r\t\v space
            this.name = this.xml.substring(ixNameStart, i);
            return i;
        }
        if (w === 62 /* ">" */ || w === 47 /* "/" */) {
            this.name = this.xml.substring(ixNameStart, i);
            return i;
        }
        return -1;
    }
    return -1;
};
StaxParser.prototype.MoveToNextAttribute = function() {
    var i = this.index;
    var w = this.xml.charCodeAt(i);
    while (i < this.length) {
        if (w === 61 /* "=" */ && i + 1 < this.length) {
            this.name = this.xml.substring(this.index, i);
            this.name = this.name.trim();
            var textStart = i + 2;
            if (this.xml.charCodeAt(textStart - 1) === 34/* "\"" */) {
                i = this.xml.indexOf("\"", textStart);
            } else {
                i = this.xml.slice(textStart - 1).search(/["']/g);
                if (-1 !== i) {
                    i += textStart - 1;//slice compensation
                    textStart = i + 1;
                    i = this.xml.indexOf(this.xml[i], textStart);
                }
            }
            if (-1 !== i) {
                this.text = this.xml.substring(textStart, i);
                this.index = i + 1;
                return true;
            } else {
                break;
            }
        }
        if (w === 32 || w === 9 || w === 10 || w === 11 || w === 12 || w === 13) { // \f\n\r\t\v space
            //todo spaces between name and =
            w = this.xml.charCodeAt(++i);
            continue;
        }
        if (w === 62 /* ">" */) {
            this.index = i + 1;
            this.isInAttr = false;
            return false;
        }
        if (w === 47 /* "/" */ && i + 1 < this.length && this.xml.charCodeAt(i + 1) === 62) {
            this.index = i + 2;
            this.isInAttr = false;
            this.isTagEnd = true;
            return false;
        }
        w = this.xml.charCodeAt(++i);
    }
    this.stop = true;
    this.index = i;
    this.isInAttr = false;
    return false;
};
StaxParser.prototype.SkipAttributes = function() {
    var i = this.xml.indexOf('>', this.index);
    if (i === -1) { // error  ...> // не нашел знак закрытия тега
        this.stop = true;
    }
    this.isTagEnd = this.xml.charCodeAt(i - 1) === 47;
    this.isInAttr = false;
    this.index = i + 1;
};
StaxParser.prototype.Read = function() {
    var hasNext = this.hasNext();
    this.next();
    return hasNext;
};
StaxParser.prototype.ReadNextNode = function() {
    var type = EasySAXEvent.Unknown;
    while (EasySAXEvent.START_ELEMENT !== type && this.hasNext()) {
        type = this.next();
    }
    return EasySAXEvent.START_ELEMENT === type;
};
StaxParser.prototype.ReadNextSiblingNode = function(depth) {
    var type = EasySAXEvent.Unknown;
    while (this.hasNext()) {
        type = this.next();
        var curDepth = this.depth;
        if (curDepth < depth)
            break;
        if (EasySAXEvent.START_ELEMENT == type && curDepth == depth + 1)
            return true;
        else if (EasySAXEvent.END_ELEMENT == type && curDepth == depth)
            return false;

    }
    return false;
};
StaxParser.prototype.ReadTillEnd = function (opt_depth) {
    var depth = opt_depth;
    if (!depth) {
        depth = this.depth;
    }
    var type = EasySAXEvent.Unknown;
    while (this.hasNext()) {
        type = this.next();
        var curDepth = this.GetDepth();
        if (curDepth < depth) {
            break;
        }
        if (EasySAXEvent.END_ELEMENT === type && curDepth === depth)
            break;
    }
    return true;
};
StaxParser.prototype.IsEmptyNode = function () {
    return this.isTagStart && this.isTagEnd;
};
StaxParser.prototype.GetDepth = function() {
    return this.depth;
};
StaxParser.prototype.GetName = function () {
    return this.name;
    // return this.ConvertToString(this.xml, this.nameStart, this.nameEnd);
};
StaxParser.prototype.GetNameNoNS = function () {
    return this.GetNameFromNodeName(this.GetName());
};
StaxParser.prototype.GetValue = function () {
    return this.text;
    // return this.ConvertToString(this.xml, this.textStart, this.textEnd);
};
StaxParser.prototype.GetBool = function (val) {
    return "1" === val || "true" === val || "t" === val || "on" === val;
};
StaxParser.prototype.GetInt = function (val, def, radix) {
    var num = parseInt(val, radix);
    return !isNaN(num) ? num : def;
};
StaxParser.prototype.GetUInt = function (val, def, radix) {
    var num = parseInt(val, radix);
    return !isNaN(num) && num >= 0 ? num : def;
};
StaxParser.prototype.GetDouble = function (val, def) {
    var num = parseFloat(val);
    return !isNaN(num) ? num : def;
};
StaxParser.prototype.GetDoubleOrNaN = function (val, def) {
    if(val === "NaN") {
        return NaN;
    }
		if (val === "INF") {
			return Infinity;
		}
    return this.GetDouble(val, def);
};
StaxParser.prototype.GetValueBool = function () {
    return this.GetBool(this.GetValue());
};
StaxParser.prototype.GetValueByte = function (def, radix) {
    return this.GetValueUInt(def, radix);
};
StaxParser.prototype.GetValueSByte = function (def, radix) {
    return this.GetValueInt(def, radix);
};
StaxParser.prototype.GetValueInt = function (def, radix) {
    return this.GetInt(this.GetValue(), def, radix);
};
StaxParser.prototype.GetValueUInt = function (def, radix) {
    return this.GetUInt(this.GetValue(), def, radix);
};
StaxParser.prototype.GetValueInt64 = function (def, radix) {
    return this.GetValueInt(def, radix);
};
StaxParser.prototype.GetValueUInt64 = function (def, radix) {
    return this.GetValueUInt(def, radix);
};
StaxParser.prototype.GetValueDouble = function (def) {
    return this.GetDouble(this.GetValue(), def);
};
StaxParser.prototype.GetValueDoubleOrNaN = function (def) {
    return this.GetDoubleOrNaN(this.GetValue(), def);
};
StaxParser.prototype.GetValueDecodeXml = function () {
    return this.DecodeXml(this.text);
};
StaxParser.prototype.GetValueDecodeXmlExt = function () {
    return this.GetValueDecodeXml();
};
StaxParser.prototype.DecodeXml = function (text) {
    if (-1 !== text.indexOf('&')) {
        var res = "";
        for (var i = 0; i < text.length; ++i) {
            if(text[i] === '&') {
                if(i + 3 < text.length) {
                    if(text[i + 1] == 'l' && text[i + 2] == 't' && text[i + 3] == ';') {
                        res += '<';
                        i+=3;
                        continue;
                    } else if(text[i + 1] == 'g' && text[i + 2] == 't' && text[i + 3] == ';') {
                        res += '>';
                        i+=3;
                        continue;
                    }
                }
                if(i + 4 < text.length && text[i + 1] == 'a' && text[i + 2] == 'm' && text[i + 3] == 'p' && text[i + 4] == ';') {
                    res += '&';
                    i+=4;
                    continue;
                }
                if(i + 5 < text.length) {
                    if(text[i + 1] == 'q' && text[i + 2] == 'u' && text[i + 3] == 'o' && text[i + 4] == 't' && text[i + 5] == ';') {
                        res += '\"';
                        i+=5;
                        continue;
                    } else if(text[i + 1] == 'a' && text[i + 2] == 'p' && text[i + 3] == 'o' && text[i + 4] == 's' && text[i + 5] == ';') {
                        res += '\'';
                        i+=5;
                        continue;
                    }
                }
            }
            res += text[i]
        }
        return res;
    } else {
        return (' ' + text).substr(1);
    }
};
StaxParser.prototype.GetText = function () {
    var text = "";
    var depth = this.depth;
    var type = EasySAXEvent.Unknown;
    while (this.hasNext()) {
        type = this.next();
        var curDepth = this.GetDepth();
        if (curDepth < depth) {
            break;
        }
        if (EasySAXEvent.END_ELEMENT === type && curDepth === depth)
            break;
        if (EasySAXEvent.CHARACTERS === type) {
            text += this.GetValue();
        }
    }
    return text;
};
StaxParser.prototype.GetTextDecodeXml = function () {
    return this.DecodeXml(this.GetText());
};
StaxParser.prototype.GetTextBool = function () {
    return this.GetBool(this.GetText());
};
StaxParser.prototype.GetTextByte = function (def, radix) {
    return this.GetTextInt(def, radix);
};
StaxParser.prototype.GetTextSByte = function (def, radix) {
    return this.GetTextUInt(def, radix);
};
StaxParser.prototype.GetTextInt = function (def, radix) {
    return this.GetInt(this.GetText(), def, radix);
};
StaxParser.prototype.GetTextUInt = function (def, radix) {
    return this.GetUInt(this.GetText(), def, radix);
};
StaxParser.prototype.GetTextInt64 = function (def, radix) {
    return this.GetTextInt(def, radix);
};
StaxParser.prototype.GetTextUInt64 = function (def, radix) {
    return this.GetTextUInt(def, radix);
};
StaxParser.prototype.GetTextDouble = function (def) {
    return this.GetDouble(this.GetText(), def);
};
StaxParser.prototype.ConvertToString = function(xml, start, end) {
    return xml.substring(start, end);
    // return "";
    // return name ? String.prototype.fromUtf8(buffer, start, end - start + 1) : "";
    // return String.prototype.fromUtf8(buffer, start, end - start + 1);
    // return name ? decoder.decode(name) : "";
    // return name ? new TextDecoder("utf-8").decode(name) : "";
    // return String.fromCharCode.apply(null, name);
};
StaxParser.prototype.GetNSFromNodeName = function(name) {
    var index = name.indexOf(':');
    if (-1 !== index) {
        return name.substring(0, index + 1);
    }
    return "";
};
StaxParser.prototype.GetNameFromNodeName = function(name) {
    var index = name.indexOf(':');
    if (-1 !== index) {
        return name.substring(index + 1);
    }
    return name;
};
StaxParser.prototype.GetEventType = function() {
    return this.eventType;
};
StaxParser.prototype.GetContext = function() {
    return this.context;
};
StaxParser.prototype.GetOformContext = function() {
	if (!this.context)
		return null;
	
	return this.context.getOformContext();
};
StaxParser.prototype.getState = function() {
    return {
        depth: this.depth,
        eventType: this.eventType,
        index: this.index,
        isInAttr: this.isInAttr,
        isTagEnd: this.isTagEnd,
        isTagStart: this.isTagStart,
        length: this.length,
        name: this.name,
        stop: this.stop,
        text: this.text,
        value: this.value
    };
};
StaxParser.prototype.setState = function(state) {
    this.depth = state.depth;
    this.eventType = state.eventType;
    this.index = state.index;
    this.isInAttr = state.isInAttr;
    this.isTagEnd = state.isTagEnd;
    this.isTagStart = state.isTagStart;
    this.length = state.length;
    this.name = state.name;
    this.stop = state.stop;
    this.text = state.text;
    this.value = state.value;
};
StaxParser.prototype.readXmlArray = function(childName, func) {
    if (this.IsEmptyNode()) {
        return;
    }

    var depth = this.GetDepth();
    var indexChild = 0;
    while (this.ReadNextSiblingNode(depth)) {
        if (childName === this.GetNameNoNS()) {
            func(indexChild);
            indexChild++;
        }
    }
};
StaxParser.prototype.AddConnectedObject = function(object) {
    this.context.addConnectorsPr(object);
};

function XmlParserContext(){
    //common
    this.zip = null;
    this.DrawingDocument = null;
    this.imageMap = {};
    this.curChart = null;
		this.smartarts = [];
    //docx
    this.commentDataById = {};
    this.oReadResult = AscCommonWord.DocReadResult && new AscCommonWord.DocReadResult();
    this.maxZIndex = 0;

    this.oformContext = null;
    this.sdtPrWithFieldPath = [];
    this.fieldMasterMap = {};

    this.fieldMastersWithUserMasterPath = [];
    this.userMasterMap = {};

    this.fieldGroupsWithFieldMasterPath = [];
    this.fieldWithFieldMastersPath = [];
    this.userWithUserMastersPath = [];

    //xlsx
    this.sharedStrings = [];
    this.row = null;
    this.cellValue = null;
    this.cellBase = null;
    this.drawingId = null;
    //pptx
    this.layoutsMap = {};
    this.notesMastersMap = {};
    this.TablesMap = {};
    this.TableStylesMap = {};
    this.ConnectorsPr = [];
}
XmlParserContext.prototype.initFromWS = function(ws) {
    this.ws = ws;
    this.row = new AscCommonExcel.Row(ws);
    this.cellValue = new AscCommonExcel.CT_Value();
    this.cellBase = new AscCommon.CellBase(0,0);
    this.drawingId = null;
};
XmlParserContext.prototype.clearSlideRelations = function() {
    this.layoutsMap = {};
    this.notesMastersMap = {};
};
XmlParserContext.prototype.clearFieldRelations = function() {
    this.sdtPrWithFieldPath = [];
    this.fieldMasterMap = {};
};
XmlParserContext.prototype.addSdtPrRelation = function(oSdtPr, sTarget) {
    this.sdtPrWithFieldPath.push({sdtPr: oSdtPr, target: sTarget});
};
XmlParserContext.prototype.addFieldGroupRelation = function(oFieldGroup, sTarget) {
    this.fieldGroupsWithFieldMasterPath.push({fieldGroup: oFieldGroup, target: sTarget});
};
XmlParserContext.prototype.addFieldMasterPath = function(oFieldMaster, sTarget) {
    this.fieldMasterMap[sTarget] = oFieldMaster;
};
XmlParserContext.prototype.addFieldWithFieldMaterPath = function(oField, sTarget) {
    this.fieldWithFieldMastersPath.push({field: oField, target: sTarget});
};
XmlParserContext.prototype.addUserWithUserMastersPath = function(oUser, sTarget) {
    this.userWithUserMastersPath.push({user: oUser, target: sTarget});
};
XmlParserContext.prototype.assignFieldsToFieldMasters = function() {
    for(let nFld = 0; nFld < this.fieldWithFieldMastersPath.length; ++nFld) {
        let oPair = this.fieldWithFieldMastersPath[nFld];
        let sTarget = oPair.target;
        for(let sKey in this.fieldMasterMap) {
            if(this.fieldMasterMap.hasOwnProperty(sKey)) {
                if(sKey.indexOf(sTarget) > -1) {
                    let oFieldMaster = this.fieldMasterMap[sKey];
                    if(oFieldMaster) {
                        oFieldMaster.setLogicField(oPair.field);
                        break;
                    }
                }
            }
        }
    }
};
XmlParserContext.prototype.assignFieldsToSdt = function() {
    for(let nSdt = 0; nSdt < this.sdtPrWithFieldPath.length; ++nSdt) {
        let oPair = this.sdtPrWithFieldPath[nSdt];
        let oFieldMaster = this.fieldMasterMap[oPair.target];
        if(oFieldMaster) {
            oPair.sdtPr.OForm = oFieldMaster;
        }
    }
};

XmlParserContext.prototype.addFieldMasterRelation = function(oFieldMaster, sTarget) {
    this.fieldMastersWithUserMasterPath.push({fieldMaster: oFieldMaster, target: sTarget});
};
XmlParserContext.prototype.addUserMasterPath = function(oUserMaster, sTarget) {
    this.userMasterMap[sTarget] = oUserMaster;
};
XmlParserContext.prototype.assignFieldsToFieldGroup = function() {
    for(let nFldGrp = 0; nFldGrp < this.fieldGroupsWithFieldMasterPath.length; ++nFldGrp) {
        let oPair = this.fieldGroupsWithFieldMasterPath[nFldGrp];
        let sTarget = oPair.target;
        for(let sKey in this.fieldMasterMap) {
            if(this.fieldMasterMap.hasOwnProperty(sKey)) {
                if(sKey.indexOf(sTarget) > -1) {
                    let oFieldMaster = this.fieldMasterMap[sKey];
                    if(oFieldMaster) {
                        oPair.fieldGroup.addField(oFieldMaster);
                        break;
                    }
                }
            }
        }
    }
};
XmlParserContext.prototype.assignUsersToFieldMasters = function() {
    for(let nPair = 0; nPair < this.fieldMastersWithUserMasterPath.length; ++nPair) {
        let oPair = this.fieldMastersWithUserMasterPath[nPair];
        let oUserMasterMaster = this.userMasterMap[oPair.target];
        if(oUserMasterMaster) {
            oPair.fieldMaster.addUser(oUserMasterMaster);
        }
    }
};
XmlParserContext.prototype.assignUsersToUserMasters = function() {
    for(let nPair = 0; nPair < this.userWithUserMastersPath.length; ++nPair) {
        let oPair = this.userWithUserMastersPath[nPair];
        let sTarget = oPair.target;
        for(let sKey in this.userMasterMap) {
            if(this.userMasterMap.hasOwnProperty(sKey)) {
                if(sKey.indexOf(sTarget) > -1) {
                    let oUserMaster = this.userMasterMap[sKey];
                    if(oUserMaster) {
                        oUserMaster.setUserId(oPair.user.Id);
                        break;
                    }
                }
            }
        }
    }
};
XmlParserContext.prototype.assignFormLinks = function() {
    this.assignFieldsToSdt();
    this.assignUsersToFieldMasters();
    this.assignFieldsToFieldGroup();
    this.assignFieldsToFieldMasters();
    this.assignUsersToUserMasters();

    this.sdtPrWithFieldPath = [];
    this.fieldMasterMap = {};
    this.fieldMastersWithUserMasterPath = [];
    this.userMasterMap = {};
    this.fieldGroupsWithFieldMasterPath = [];
    this.fieldWithFieldMastersPath = [];
    this.userWithUserMastersPath = [];
};
XmlParserContext.prototype.getOformContext = function() {
	return this.oformContext;
};
XmlParserContext.prototype.setOformContext = function(context) {
	this.oformContext = context;
};

XmlParserContext.prototype.addTableStyle = function(sGuid, oStyle) {
    this.TableStylesMap[sGuid] = oStyle;
};
XmlParserContext.prototype.getTableStyle = function(sGuid) {
    return this.TableStylesMap[sGuid] || null;
};
XmlParserContext.prototype.addConnectorsPr = function(oPr) {
    for(let nIdx = 0; nIdx < this.ConnectorsPr.length; ++nIdx) {
        if(oPr === this.ConnectorsPr[nIdx]) {
            return;
        }
    }
    this.ConnectorsPr.push(oPr);
};
XmlParserContext.prototype.assignConnectors = function(aSpTree) {
    for(let nIdx = 0; nIdx < this.ConnectorsPr.length; ++nIdx) {
        this.ConnectorsPr[nIdx].assignConnectors(aSpTree);
    }
    this.ConnectorsPr.length = 0;
};
XmlParserContext.prototype.checkZIndex = function(nZIndex) {
    if(AscFormat.isRealNumber(nZIndex)) {
        this.maxZIndex = Math.max(this.maxZIndex, nZIndex);
    }
};
XmlParserContext.prototype.loadDataLinks = function() {
    let _cur_ind = 0;
    let oImageMap = {};
    for (let path in this.imageMap) {
        if (this.imageMap.hasOwnProperty(path)) {
            oImageMap[_cur_ind++] = path;
            let data = this.zip.getFile(path);
            if (data) {
                let blobUrl = AscCommon.g_oDocumentBlobUrls.getBlobUrl(path, this.zip);
                AscCommon.g_oDocumentUrls.addImageUrl(path, blobUrl);

                this.imageMap[path].forEach(function(blipFill) {
                    AscCommon.pptx_content_loader.Reader.initAfterBlipFill(path, blipFill);
                });
            }
        }
    }
    return oImageMap;
};
XmlParserContext.prototype.GenerateSmartArts = function() {
	while (this.smartarts.length) {
		const smartart = this.smartarts.pop();
		smartart.generateDrawingPart();
	}
};
XmlParserContext.prototype.ClearSmartArts = function() {
	this.smartarts.length = 0;
};
function XmlWriterContext(editorId){
    //common
    this.editorId = editorId;
    this.zip = null;
    this.part = null;
    this.imageMap = {};
    this.currentPartImageMap = {};
    this.dataMap = {};
    this.currentPartDataMap = {};
    this.m_lObjectIdVML = 1025;

    this.oUriMap = {};
    this.objectId = 1;
    this.groupIndex = 0;
    this.flag = 0;
    switch (editorId) {
        case AscCommon.c_oEditorId.Word: {
            this.docType = AscFormat.XMLWRITER_DOC_TYPE_DOCX;
            break;
        }
        case AscCommon.c_oEditorId.Spreadsheet: {
            this.docType = AscFormat.XMLWRITER_DOC_TYPE_XLSX;
            break;
        }
        case AscCommon.c_oEditorId.Presentation: {
            this.docType = AscFormat.XMLWRITER_DOC_TYPE_PPTX;
            break;
        }
    }
    //docx
    this.document = null;
    this.oNumIdMap = {};
    this.commentIdIndex = 1;
    this.paraIdIndex = 1;
    this.commentDataById = {};
    this.docSaveParams = null;

    this.fieldMastersPartMap = {};
    this.userMasterPartMap = {};

    //xlsx
    this.wb = null;
    this.sheetIds = [];
    this.sharedStrings = null;
    this.isCopyPaste = null;
    this.stylesForWrite = AscCommonExcel.StylesForWrite && new AscCommonExcel.StylesForWrite(); //no StylesForWrite in vsdx now
    this.oSharedStrings = {index: 0, strings: {}};
    this.oleDrawings = [];
    this.signatureDrawings = [];
    //pptx
    this.presentation = null;
    this.sldMasterIdLst = [];
    this.sldLayoutIdLst = [];
    this.sldLayoutsCount = 0;
    this.notesMasterIdLst = [];
    this.handoutMasterIdLst = [];
    this.sldIdLst = [];
    this.tableStylesIdToGuid = {};
}
XmlWriterContext.prototype.initFromWS = function(ws) {
    this.ws = ws;
    this.row = new AscCommonExcel.Row(ws);
    this.cellValue = new AscCommonExcel.CT_Value();
    this.cellBase = new AscCommon.CellBase(0,0);
    this.drawingId = null;
};

XmlWriterContext.prototype.addSlideRel = function(sRel) {
    this.sldIdLst.push(sRel);
};
XmlWriterContext.prototype.addSlideLayoutRel = function(sRel) {
    this.sldLayoutIdLst.push(sRel);
};
XmlWriterContext.prototype.addSlideMasterRel = function(sRel) {
    this.sldMasterIdLst.push(sRel);
};
XmlWriterContext.prototype.addNotesMasterRel = function(sRel) {
    this.notesMasterIdLst.push(sRel);
};
XmlWriterContext.prototype.addFieldMasterPart = function(oFieldMaster, oPart) {
    this.fieldMastersPartMap[oFieldMaster.Id] = oPart;
};
XmlWriterContext.prototype.addUserMasterPart = function(oUserMaster, oPart) {
    this.userMasterPartMap[oUserMaster.Id] = oPart;
};
XmlWriterContext.prototype.clearSlideLayoutRels = function() {
    this.sldLayoutIdLst.length = 0;
};
XmlWriterContext.prototype.getSlideMastersCount = function() {
    return this.sldMasterIdLst.length;
};
XmlWriterContext.prototype.getSlidesCount = function() {
    return this.sldIdLst.length;
};
XmlWriterContext.prototype.clearCurrentPartDataMaps = function() {
    this.currentPartImageMap = {};
    this.currentPartDataMap = {};
};
XmlWriterContext.prototype.getImageRId = function(sRasterImageId) {
    let imagePart = this.imageMap[sRasterImageId];
    let type = this.editorId === AscCommon.c_oEditorId.Word ? AscCommon.openXml.Types.imageWord : AscCommon.openXml.Types.image;
    if (!imagePart) {
        if (this.part) {
            let ext = AscCommon.GetFileExtension(sRasterImageId);
            type = Object.assign({}, type);
            type.filename += ext;
            type.contentType = AscCommon.openXml.GetMimeType(ext);
            imagePart = this.part.addPart(type);
            if (imagePart) {
                this.imageMap[sRasterImageId] = imagePart;
                this.currentPartImageMap[sRasterImageId] = imagePart.rId;
            }
        }
    }
    else {
        if(!this.currentPartImageMap[sRasterImageId]) {
            if(this.part) {
                this.currentPartImageMap[sRasterImageId] = this.part.addRelationship(type.relationType, imagePart.part.uri);
            }
        }
    }
    return this.currentPartImageMap[sRasterImageId] ? this.currentPartImageMap[sRasterImageId] : "";
};
XmlWriterContext.prototype.getDataRId = function(sDataLink) {
    let dataPart = this.dataMap[sDataLink];
    let type = this.editorId === AscCommon.c_oEditorId.Word ? AscCommon.openXml.Types.wordPackage : AscCommon.openXml.Types.package;
    if (!dataPart) {
        if (this.part) {
            let ext = AscCommon.GetFileExtension(sDataLink);
            type = Object.assign({}, type);
            type.filename = AscCommon.changeFileExtention(type.filename, ext, null);
            type.contentType = AscCommon.openXml.GetMimeType(ext);
            dataPart = this.part.addPart(type);
            if (dataPart) {
                this.dataMap[sDataLink] = dataPart;
                this.currentPartDataMap[sDataLink] = dataPart.rId;
            }
        }
    }
    else {
        if(!this.currentPartDataMap[sDataLink]) {
            if(this.part) {
                this.currentPartDataMap[sDataLink] = this.part.addRelationship(type.relationType, dataPart.part.uri);
            }
        }
    }
    return this.currentPartDataMap[sDataLink] ? this.currentPartDataMap[sDataLink] : "";
};
XmlWriterContext.prototype.getSpIdxId = function(sEditorId){
    if(typeof sEditorId === "string" && sEditorId.length > 0) {
        var oDrawing = AscCommon.g_oTableId.Get_ById(sEditorId);
        if(oDrawing && oDrawing.getFormatId) {
            return oDrawing.getFormatId();
        }
    }
    return null;
};
function CT_XmlNode(opt_elemReader) {
    this.attributes = {};
    this.members = {};
    this.text = null;

    this.elemReader = opt_elemReader || function(){};
}
CT_XmlNode.fromReader = function(reader, opt_elemReader) {
	let node = new CT_XmlNode(opt_elemReader);
	node.fromXml(reader);
	return node;
};
CT_XmlNode.prototype.readAttr = function(reader) {
    while (reader.MoveToNextAttribute()) {
        this.attributes[reader.GetNameNoNS()] = reader.GetValue();
    }
};
CT_XmlNode.prototype.fromXml = function(reader) {
    this.readAttr(reader);
    var elem, depth = reader.GetDepth();
    while (reader.Read()) {
        switch(reader.GetEventType()) {
            case EasySAXEvent.START_ELEMENT:
                var name = reader.GetNameNoNS();
                elem = this.elemReader.call(this, reader, name);
                if (!elem) {
                    elem = new CT_XmlNode();
                    elem.fromXml(reader);
                }
                this.members[name] = elem;
                break;
            case EasySAXEvent.CHARACTERS:
                if(!this.text) {
                    this.text = "";
                }
                this.text += reader.GetValue();
                break;
            case EasySAXEvent.END_ELEMENT:
                if (reader.GetDepth() === depth)
                    return;
                break;
        }
    }
};
CT_XmlNode.prototype.toXml = function(writer, name) {
    var i;
    writer.WriteXmlNodeStart(name);
    for (i in this.attributes) {
        if (this.attributes.hasOwnProperty(i)) {
            writer.WriteXmlNullableAttributeString(i, this.attributes[i]);
        }
    }
    writer.WriteXmlAttributesEnd();
    for (i in this.members) {
        if (this.members.hasOwnProperty(i)) {
            var member = this.members[i];
            if(Array.isArray(member)) {
                for(var nIdx = 0; nIdx < member.length; ++nIdx) {
                    member[nIdx].toXml(writer, i);
                }
            }
            else {
                member.toXml(writer, i);
            }
        }
    }
    if (null !== this.text && undefined !== this.text) {
        writer.WriteXmlStringEncode(this.text.toString());
    }
    writer.WriteXmlNodeEnd(name);
};

window['AscCommon'] = window['AscCommon'] || {};
window['AscCommon'].XmlParserContext = XmlParserContext;
window["AscCommon"].XmlWriterContext = XmlWriterContext;
window['AscCommon'].StaxParser = StaxParser;

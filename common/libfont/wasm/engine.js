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

(function(window, undefined) {

    var AscFonts = window['AscFonts'];
    var printErr = undefined;
    var FS = undefined;
    var print = undefined;

    var fetch = window.fetch;
    var getBinaryPromise = null;
    if (window["AscDesktopEditor"] && document.currentScript && 0 == document.currentScript.src.indexOf("file:///"))
    {
        fetch = undefined; // fetch not support file:/// scheme
        getBinaryPromise = function() {

            var wasmPath = "ascdesktop://fonts/" + wasmBinaryFile.substr(8);
            return new Promise(function (resolve, reject) {

                var xhr = new XMLHttpRequest();
                xhr.open('GET', wasmPath, true);
                xhr.responseType = 'arraybuffer';

                if (xhr.overrideMimeType)
                    xhr.overrideMimeType('text/plain; charset=x-user-defined');
                else
                    xhr.setRequestHeader('Accept-Charset', 'x-user-defined');

                xhr.onload = function () {
                    if (this.status == 200) {
                        resolve(new Uint8Array(this.response));
                    }
                };

                xhr.send(null);

            });
        }
    }
    else
    {
        getBinaryPromise = function() {
            return getBinaryPromise2();
        }
    }

    var Module=typeof Module!=="undefined"?Module:{};var moduleOverrides={};var key;for(key in Module){if(Module.hasOwnProperty(key)){moduleOverrides[key]=Module[key]}}Module["arguments"]=[];Module["thisProgram"]="./this.program";Module["quit"]=function(status,toThrow){throw toThrow};Module["preRun"]=[];Module["postRun"]=[];var ENVIRONMENT_IS_WEB=true;var ENVIRONMENT_IS_WORKER=false;var scriptDirectory="";function locateFile(path){if(Module["locateFile"]){return Module["locateFile"](path,scriptDirectory)}else{return scriptDirectory+path}}if(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER){if(ENVIRONMENT_IS_WORKER){scriptDirectory=self.location.href}else if(document.currentScript){scriptDirectory=document.currentScript.src}if(scriptDirectory.indexOf("blob:")!==0){scriptDirectory=scriptDirectory.substr(0,scriptDirectory.lastIndexOf("/")+1)}else{scriptDirectory=""}Module["read"]=function shell_read(url){var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.send(null);return xhr.responseText};if(ENVIRONMENT_IS_WORKER){Module["readBinary"]=function readBinary(url){var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.responseType="arraybuffer";xhr.send(null);return new Uint8Array(xhr.response)}}Module["readAsync"]=function readAsync(url,onload,onerror){var xhr=new XMLHttpRequest;xhr.open("GET",url,true);xhr.responseType="arraybuffer";xhr.onload=function xhr_onload(){if(xhr.status==200||xhr.status==0&&xhr.response){onload(xhr.response);return}onerror()};xhr.onerror=onerror;xhr.send(null)};Module["setWindowTitle"]=function(title){document.title=title}}else{}var out=Module["print"]||(typeof console!=="undefined"?console.log.bind(console):typeof print!=="undefined"?print:null);var err=Module["printErr"]||(typeof printErr!=="undefined"?printErr:typeof console!=="undefined"&&console.warn.bind(console)||out);for(key in moduleOverrides){if(moduleOverrides.hasOwnProperty(key)){Module[key]=moduleOverrides[key]}}moduleOverrides=undefined;function dynamicAlloc(size){var ret=HEAP32[DYNAMICTOP_PTR>>2];var end=ret+size+15&-16;if(end<=_emscripten_get_heap_size()){HEAP32[DYNAMICTOP_PTR>>2]=end}else{var success=_emscripten_resize_heap(end);if(!success)return 0}return ret}var asm2wasmImports={"f64-rem":function(x,y){return x%y},"debugger":function(){debugger}};var functionPointers=new Array(0);var tempRet0=0;var setTempRet0=function(value){tempRet0=value};var getTempRet0=function(){return tempRet0};if(typeof WebAssembly!=="object"){err("no native wasm support detected")}var wasmMemory;var wasmTable;var ABORT=false;var EXITSTATUS=0;function getMemory(size){if(!runtimeInitialized)return dynamicAlloc(size);return _malloc(size)}var UTF8Decoder=typeof TextDecoder!=="undefined"?new TextDecoder("utf8"):undefined;function UTF8ArrayToString(u8Array,idx,maxBytesToRead){var endIdx=idx+maxBytesToRead;var endPtr=idx;while(u8Array[endPtr]&&!(endPtr>=endIdx))++endPtr;if(endPtr-idx>16&&u8Array.subarray&&UTF8Decoder){return UTF8Decoder.decode(u8Array.subarray(idx,endPtr))}else{var str="";while(idx<endPtr){var u0=u8Array[idx++];if(!(u0&128)){str+=String.fromCharCode(u0);continue}var u1=u8Array[idx++]&63;if((u0&224)==192){str+=String.fromCharCode((u0&31)<<6|u1);continue}var u2=u8Array[idx++]&63;if((u0&240)==224){u0=(u0&15)<<12|u1<<6|u2}else{u0=(u0&7)<<18|u1<<12|u2<<6|u8Array[idx++]&63}if(u0<65536){str+=String.fromCharCode(u0)}else{var ch=u0-65536;str+=String.fromCharCode(55296|ch>>10,56320|ch&1023)}}}return str}function UTF8ToString(ptr,maxBytesToRead){return ptr?UTF8ArrayToString(HEAPU8,ptr,maxBytesToRead):""}function stringToUTF8Array(str,outU8Array,outIdx,maxBytesToWrite){if(!(maxBytesToWrite>0))return 0;var startIdx=outIdx;var endIdx=outIdx+maxBytesToWrite-1;for(var i=0;i<str.length;++i){var u=str.charCodeAt(i);if(u>=55296&&u<=57343){var u1=str.charCodeAt(++i);u=65536+((u&1023)<<10)|u1&1023}if(u<=127){if(outIdx>=endIdx)break;outU8Array[outIdx++]=u}else if(u<=2047){if(outIdx+1>=endIdx)break;outU8Array[outIdx++]=192|u>>6;outU8Array[outIdx++]=128|u&63}else if(u<=65535){if(outIdx+2>=endIdx)break;outU8Array[outIdx++]=224|u>>12;outU8Array[outIdx++]=128|u>>6&63;outU8Array[outIdx++]=128|u&63}else{if(outIdx+3>=endIdx)break;outU8Array[outIdx++]=240|u>>18;outU8Array[outIdx++]=128|u>>12&63;outU8Array[outIdx++]=128|u>>6&63;outU8Array[outIdx++]=128|u&63}}outU8Array[outIdx]=0;return outIdx-startIdx}function lengthBytesUTF8(str){var len=0;for(var i=0;i<str.length;++i){var u=str.charCodeAt(i);if(u>=55296&&u<=57343)u=65536+((u&1023)<<10)|str.charCodeAt(++i)&1023;if(u<=127)++len;else if(u<=2047)len+=2;else if(u<=65535)len+=3;else len+=4}return len}var UTF16Decoder=typeof TextDecoder!=="undefined"?new TextDecoder("utf-16le"):undefined;function allocateUTF8(str){var size=lengthBytesUTF8(str)+1;var ret=_malloc(size);if(ret)stringToUTF8Array(str,HEAP8,ret,size);return ret}function writeAsciiToMemory(str,buffer,dontAddNull){for(var i=0;i<str.length;++i){HEAP8[buffer++>>0]=str.charCodeAt(i)}if(!dontAddNull)HEAP8[buffer>>0]=0}var WASM_PAGE_SIZE=65536;function alignUp(x,multiple){if(x%multiple>0){x+=multiple-x%multiple}return x}var buffer,HEAP8,HEAPU8,HEAP16,HEAPU16,HEAP32,HEAPU32,HEAPF32,HEAPF64;function updateGlobalBuffer(buf){Module["buffer"]=buffer=buf}function updateGlobalBufferViews(){Module["HEAP8"]=HEAP8=new Int8Array(buffer);Module["HEAP16"]=HEAP16=new Int16Array(buffer);Module["HEAP32"]=HEAP32=new Int32Array(buffer);Module["HEAPU8"]=HEAPU8=new Uint8Array(buffer);Module["HEAPU16"]=HEAPU16=new Uint16Array(buffer);Module["HEAPU32"]=HEAPU32=new Uint32Array(buffer);Module["HEAPF32"]=HEAPF32=new Float32Array(buffer);Module["HEAPF64"]=HEAPF64=new Float64Array(buffer)}var DYNAMIC_BASE=5352496,DYNAMICTOP_PTR=109360;var TOTAL_STACK=5242880;var TOTAL_MEMORY=Module["TOTAL_MEMORY"]||16777216;if(TOTAL_MEMORY<TOTAL_STACK)err("TOTAL_MEMORY should be larger than TOTAL_STACK, was "+TOTAL_MEMORY+"! (TOTAL_STACK="+TOTAL_STACK+")");if(Module["buffer"]){buffer=Module["buffer"]}else{if(typeof WebAssembly==="object"&&typeof WebAssembly.Memory==="function"){wasmMemory=new WebAssembly.Memory({"initial":TOTAL_MEMORY/WASM_PAGE_SIZE});buffer=wasmMemory.buffer}else{buffer=new ArrayBuffer(TOTAL_MEMORY)}Module["buffer"]=buffer}updateGlobalBufferViews();HEAP32[DYNAMICTOP_PTR>>2]=DYNAMIC_BASE;function callRuntimeCallbacks(callbacks){while(callbacks.length>0){var callback=callbacks.shift();if(typeof callback=="function"){callback();continue}var func=callback.func;if(typeof func==="number"){if(callback.arg===undefined){Module["dynCall_v"](func)}else{Module["dynCall_vi"](func,callback.arg)}}else{func(callback.arg===undefined?null:callback.arg)}}}var __ATPRERUN__=[];var __ATINIT__=[];var __ATMAIN__=[];var __ATPOSTRUN__=[function(){window["AscFonts"].onLoadModule();}];var runtimeInitialized=false;function preRun(){if(Module["preRun"]){if(typeof Module["preRun"]=="function")Module["preRun"]=[Module["preRun"]];while(Module["preRun"].length){addOnPreRun(Module["preRun"].shift())}}callRuntimeCallbacks(__ATPRERUN__)}function ensureInitRuntime(){if(runtimeInitialized)return;runtimeInitialized=true;callRuntimeCallbacks(__ATINIT__)}function preMain(){callRuntimeCallbacks(__ATMAIN__)}function postRun(){if(Module["postRun"]){if(typeof Module["postRun"]=="function")Module["postRun"]=[Module["postRun"]];while(Module["postRun"].length){addOnPostRun(Module["postRun"].shift())}}callRuntimeCallbacks(__ATPOSTRUN__)}function addOnPreRun(cb){__ATPRERUN__.unshift(cb)}function addOnPostRun(cb){__ATPOSTRUN__.unshift(cb)}var runDependencies=0;var runDependencyWatcher=null;var dependenciesFulfilled=null;function addRunDependency(id){runDependencies++;if(Module["monitorRunDependencies"]){Module["monitorRunDependencies"](runDependencies)}}function removeRunDependency(id){runDependencies--;if(Module["monitorRunDependencies"]){Module["monitorRunDependencies"](runDependencies)}if(runDependencies==0){if(runDependencyWatcher!==null){clearInterval(runDependencyWatcher);runDependencyWatcher=null}if(dependenciesFulfilled){var callback=dependenciesFulfilled;dependenciesFulfilled=null;callback()}}}Module["preloadedImages"]={};Module["preloadedAudios"]={};var dataURIPrefix="data:application/octet-stream;base64,";function isDataURI(filename){return String.prototype.startsWith?filename.startsWith(dataURIPrefix):filename.indexOf(dataURIPrefix)===0}var wasmBinaryFile="fonts.wasm";if(!isDataURI(wasmBinaryFile)){wasmBinaryFile=locateFile(wasmBinaryFile)}function getBinary(){try{if(Module["wasmBinary"]){return new Uint8Array(Module["wasmBinary"])}if(Module["readBinary"]){return Module["readBinary"](wasmBinaryFile)}else{throw"both async and sync fetching of the wasm failed"}}catch(err){abort(err)}}function getBinaryPromise2(){if(!Module["wasmBinary"]&&(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER)&&typeof fetch==="function"){return fetch(wasmBinaryFile,{credentials:"same-origin"}).then(function(response){if(!response["ok"]){throw"failed to load wasm binary file at '"+wasmBinaryFile+"'"}return response["arrayBuffer"]()}).catch(function(){return getBinary()})}return new Promise(function(resolve,reject){resolve(getBinary())})}function createWasm(env){var info={"env":env,"global":{"NaN":NaN,Infinity:Infinity},"global.Math":Math,"asm2wasm":asm2wasmImports};function receiveInstance(instance,module){var exports=instance.exports;Module["asm"]=exports;removeRunDependency("wasm-instantiate")}addRunDependency("wasm-instantiate");if(Module["instantiateWasm"]){try{return Module["instantiateWasm"](info,receiveInstance)}catch(e){err("Module.instantiateWasm callback failed with error: "+e);return false}}function receiveInstantiatedSource(output){receiveInstance(output["instance"])}function instantiateArrayBuffer(receiver){getBinaryPromise().then(function(binary){return WebAssembly.instantiate(binary,info)}).then(receiver,function(reason){err("failed to asynchronously prepare wasm: "+reason);abort(reason)})}if(!Module["wasmBinary"]&&typeof WebAssembly.instantiateStreaming==="function"&&!isDataURI(wasmBinaryFile)&&typeof fetch==="function"){WebAssembly.instantiateStreaming(fetch(wasmBinaryFile,{credentials:"same-origin"}),info).then(receiveInstantiatedSource,function(reason){err("wasm streaming compile failed: "+reason);err("falling back to ArrayBuffer instantiation");instantiateArrayBuffer(receiveInstantiatedSource)})}else{instantiateArrayBuffer(receiveInstantiatedSource)}return{}}Module["asm"]=function(global,env,providedBuffer){env["memory"]=wasmMemory;env["table"]=wasmTable=new WebAssembly.Table({"initial":814,"maximum":814,"element":"anyfunc"});env["__memory_base"]=1024;env["__table_base"]=0;var exports=createWasm(env);return exports};__ATINIT__.push({func:function(){___emscripten_environ_constructor()}});var ENV={};function ___buildEnvironment(environ){var MAX_ENV_VALUES=64;var TOTAL_ENV_SIZE=1024;var poolPtr;var envPtr;if(!___buildEnvironment.called){___buildEnvironment.called=true;ENV["USER"]=ENV["LOGNAME"]="web_user";ENV["PATH"]="/";ENV["PWD"]="/";ENV["HOME"]="/home/web_user";ENV["LANG"]="C.UTF-8";ENV["_"]=Module["thisProgram"];poolPtr=getMemory(TOTAL_ENV_SIZE);envPtr=getMemory(MAX_ENV_VALUES*4);HEAP32[envPtr>>2]=poolPtr;HEAP32[environ>>2]=envPtr}else{envPtr=HEAP32[environ>>2];poolPtr=HEAP32[envPtr>>2]}var strings=[];var totalSize=0;for(var key in ENV){if(typeof ENV[key]==="string"){var line=key+"="+ENV[key];strings.push(line);totalSize+=line.length}}if(totalSize>TOTAL_ENV_SIZE){throw new Error("Environment size exceeded TOTAL_ENV_SIZE!")}var ptrSize=4;for(var i=0;i<strings.length;i++){var line=strings[i];writeAsciiToMemory(line,poolPtr);HEAP32[envPtr+i*ptrSize>>2]=poolPtr;poolPtr+=line.length+1}HEAP32[envPtr+strings.length*ptrSize>>2]=0}function ___lock(){}var SYSCALLS={buffers:[null,[],[]],printChar:function(stream,curr){var buffer=SYSCALLS.buffers[stream];if(curr===0||curr===10){(stream===1?out:err)(UTF8ArrayToString(buffer,0));buffer.length=0}else{buffer.push(curr)}},varargs:0,get:function(varargs){SYSCALLS.varargs+=4;var ret=HEAP32[SYSCALLS.varargs-4>>2];return ret},getStr:function(){var ret=UTF8ToString(SYSCALLS.get());return ret},get64:function(){var low=SYSCALLS.get(),high=SYSCALLS.get();return low},getZero:function(){SYSCALLS.get()}};function ___syscall140(which,varargs){SYSCALLS.varargs=varargs;try{var stream=SYSCALLS.getStreamFromFD(),offset_high=SYSCALLS.get(),offset_low=SYSCALLS.get(),result=SYSCALLS.get(),whence=SYSCALLS.get();var offset=offset_low;FS.llseek(stream,offset,whence);HEAP32[result>>2]=stream.position;if(stream.getdents&&offset===0&&whence===0)stream.getdents=null;return 0}catch(e){if(typeof FS==="undefined"||!(e instanceof FS.ErrnoError))abort(e);return-e.errno}}function ___syscall145(which,varargs){SYSCALLS.varargs=varargs;try{var stream=SYSCALLS.getStreamFromFD(),iov=SYSCALLS.get(),iovcnt=SYSCALLS.get();return SYSCALLS.doReadv(stream,iov,iovcnt)}catch(e){if(typeof FS==="undefined"||!(e instanceof FS.ErrnoError))abort(e);return-e.errno}}function ___syscall146(which,varargs){SYSCALLS.varargs=varargs;try{var stream=SYSCALLS.get(),iov=SYSCALLS.get(),iovcnt=SYSCALLS.get();var ret=0;for(var i=0;i<iovcnt;i++){var ptr=HEAP32[iov+i*8>>2];var len=HEAP32[iov+(i*8+4)>>2];for(var j=0;j<len;j++){SYSCALLS.printChar(stream,HEAPU8[ptr+j])}ret+=len}return ret}catch(e){if(typeof FS==="undefined"||!(e instanceof FS.ErrnoError))abort(e);return-e.errno}}function ___setErrNo(value){if(Module["___errno_location"])HEAP32[Module["___errno_location"]()>>2]=value;return value}function ___syscall221(which,varargs){SYSCALLS.varargs=varargs;try{return 0}catch(e){if(typeof FS==="undefined"||!(e instanceof FS.ErrnoError))abort(e);return-e.errno}}function ___syscall5(which,varargs){SYSCALLS.varargs=varargs;try{var pathname=SYSCALLS.getStr(),flags=SYSCALLS.get(),mode=SYSCALLS.get();var stream=FS.open(pathname,flags,mode);return stream.fd}catch(e){if(typeof FS==="undefined"||!(e instanceof FS.ErrnoError))abort(e);return-e.errno}}function ___syscall54(which,varargs){SYSCALLS.varargs=varargs;try{return 0}catch(e){if(typeof FS==="undefined"||!(e instanceof FS.ErrnoError))abort(e);return-e.errno}}function ___syscall6(which,varargs){SYSCALLS.varargs=varargs;try{var stream=SYSCALLS.getStreamFromFD();FS.close(stream);return 0}catch(e){if(typeof FS==="undefined"||!(e instanceof FS.ErrnoError))abort(e);return-e.errno}}function ___unlock(){}function _emscripten_get_heap_size(){return TOTAL_MEMORY}function _longjmp(env,value){_setThrew(env,value||1);throw"longjmp"}function _emscripten_longjmp(env,value){_longjmp(env,value)}function abortOnCannotGrowMemory(requestedSize){abort("OOM")}function emscripten_realloc_buffer(size){var PAGE_MULTIPLE=65536;size=alignUp(size,PAGE_MULTIPLE);var old=Module["buffer"];var oldSize=old.byteLength;try{var result=wasmMemory.grow((size-oldSize)/65536);if(result!==(-1|0)){return Module["buffer"]=wasmMemory.buffer}else{return null}}catch(e){return null}}function _emscripten_resize_heap(requestedSize){var oldSize=_emscripten_get_heap_size();var PAGE_MULTIPLE=65536;var LIMIT=2147483648-PAGE_MULTIPLE;if(requestedSize>LIMIT){return false}var MIN_TOTAL_MEMORY=16777216;var newSize=Math.max(oldSize,MIN_TOTAL_MEMORY);while(newSize<requestedSize){if(newSize<=536870912){newSize=alignUp(2*newSize,PAGE_MULTIPLE)}else{newSize=Math.min(alignUp((3*newSize+2147483648)/4,PAGE_MULTIPLE),LIMIT)}}var replacement=emscripten_realloc_buffer(newSize);if(!replacement||replacement.byteLength!=newSize){return false}updateGlobalBuffer(replacement);updateGlobalBufferViews();TOTAL_MEMORY=newSize;return true}function _getenv(name){if(name===0)return 0;name=UTF8ToString(name);if(!ENV.hasOwnProperty(name))return 0;if(_getenv.ret)_free(_getenv.ret);_getenv.ret=allocateUTF8(ENV[name]);return _getenv.ret}function _emscripten_memcpy_big(dest,src,num){HEAPU8.set(HEAPU8.subarray(src,src+num),dest)}function invoke_iii(index,a1,a2){var sp=stackSave();try{return dynCall_iii(index,a1,a2)}catch(e){stackRestore(sp);if(e!==e+0&&e!=="longjmp")throw e;_setThrew(1,0)}}function invoke_iiii(index,a1,a2,a3){var sp=stackSave();try{return dynCall_iiii(index,a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0&&e!=="longjmp")throw e;_setThrew(1,0)}}function invoke_iiiii(index,a1,a2,a3,a4){var sp=stackSave();try{return dynCall_iiiii(index,a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0&&e!=="longjmp")throw e;_setThrew(1,0)}}function invoke_vii(index,a1,a2){var sp=stackSave();try{dynCall_vii(index,a1,a2)}catch(e){stackRestore(sp);if(e!==e+0&&e!=="longjmp")throw e;_setThrew(1,0)}}function invoke_viiii(index,a1,a2,a3,a4){var sp=stackSave();try{dynCall_viiii(index,a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0&&e!=="longjmp")throw e;_setThrew(1,0)}}var asmGlobalArg={};var asmLibraryArg={"b":abort,"e":setTempRet0,"d":getTempRet0,"r":invoke_iii,"q":invoke_iiii,"p":invoke_iiiii,"o":invoke_vii,"n":invoke_viiii,"m":___buildEnvironment,"l":___lock,"k":___setErrNo,"A":___syscall140,"z":___syscall145,"j":___syscall146,"f":___syscall221,"y":___syscall5,"i":___syscall54,"h":___syscall6,"g":___unlock,"x":_emscripten_get_heap_size,"w":_emscripten_longjmp,"v":_emscripten_memcpy_big,"u":_emscripten_resize_heap,"t":_getenv,"c":_longjmp,"s":abortOnCannotGrowMemory,"a":DYNAMICTOP_PTR};var asm=Module["asm"](asmGlobalArg,asmLibraryArg,buffer);Module["asm"]=asm;var _ASC_FT_Free=Module["_ASC_FT_Free"]=function(){return Module["asm"]["B"].apply(null,arguments)};var _ASC_FT_GetFaceInfo=Module["_ASC_FT_GetFaceInfo"]=function(){return Module["asm"]["C"].apply(null,arguments)};var _ASC_FT_GetFaceMaxAdvanceX=Module["_ASC_FT_GetFaceMaxAdvanceX"]=function(){return Module["asm"]["D"].apply(null,arguments)};var _ASC_FT_GetKerningX=Module["_ASC_FT_GetKerningX"]=function(){return Module["asm"]["E"].apply(null,arguments)};var _ASC_FT_Get_Glyph_Measure_Params=Module["_ASC_FT_Get_Glyph_Measure_Params"]=function(){return Module["asm"]["F"].apply(null,arguments)};var _ASC_FT_Get_Glyph_Render_Buffer=Module["_ASC_FT_Get_Glyph_Render_Buffer"]=function(){return Module["asm"]["G"].apply(null,arguments)};var _ASC_FT_Get_Glyph_Render_Params=Module["_ASC_FT_Get_Glyph_Render_Params"]=function(){return Module["asm"]["H"].apply(null,arguments)};var _ASC_FT_Glyph_Get_CBox=Module["_ASC_FT_Glyph_Get_CBox"]=function(){return Module["asm"]["I"].apply(null,arguments)};var _ASC_FT_Init=Module["_ASC_FT_Init"]=function(){return Module["asm"]["J"].apply(null,arguments)};var _ASC_FT_Malloc=Module["_ASC_FT_Malloc"]=function(){return Module["asm"]["K"].apply(null,arguments)};var _ASC_FT_Open_Face=Module["_ASC_FT_Open_Face"]=function(){return Module["asm"]["L"].apply(null,arguments)};var _ASC_FT_SetCMapForCharCode=Module["_ASC_FT_SetCMapForCharCode"]=function(){return Module["asm"]["M"].apply(null,arguments)};var _ASC_FT_Set_Transform=Module["_ASC_FT_Set_Transform"]=function(){return Module["asm"]["N"].apply(null,arguments)};var _ASC_FT_Set_TrueType_HintProp=Module["_ASC_FT_Set_TrueType_HintProp"]=function(){return Module["asm"]["O"].apply(null,arguments)};var _FT_Done_Face=Module["_FT_Done_Face"]=function(){return Module["asm"]["P"].apply(null,arguments)};var _FT_Done_FreeType=Module["_FT_Done_FreeType"]=function(){return Module["asm"]["Q"].apply(null,arguments)};var _FT_Get_Glyph=Module["_FT_Get_Glyph"]=function(){return Module["asm"]["R"].apply(null,arguments)};var _FT_Load_Glyph=Module["_FT_Load_Glyph"]=function(){return Module["asm"]["S"].apply(null,arguments)};var _FT_Set_Char_Size=Module["_FT_Set_Char_Size"]=function(){return Module["asm"]["T"].apply(null,arguments)};var _FT_Set_Transform=Module["_FT_Set_Transform"]=function(){return Module["asm"]["U"].apply(null,arguments)};var ___emscripten_environ_constructor=Module["___emscripten_environ_constructor"]=function(){return Module["asm"]["V"].apply(null,arguments)};var _free=Module["_free"]=function(){return Module["asm"]["W"].apply(null,arguments)};var _malloc=Module["_malloc"]=function(){return Module["asm"]["X"].apply(null,arguments)};var _setThrew=Module["_setThrew"]=function(){return Module["asm"]["Y"].apply(null,arguments)};var stackRestore=Module["stackRestore"]=function(){return Module["asm"]["da"].apply(null,arguments)};var stackSave=Module["stackSave"]=function(){return Module["asm"]["ea"].apply(null,arguments)};var dynCall_iii=Module["dynCall_iii"]=function(){return Module["asm"]["Z"].apply(null,arguments)};var dynCall_iiii=Module["dynCall_iiii"]=function(){return Module["asm"]["_"].apply(null,arguments)};var dynCall_iiiii=Module["dynCall_iiiii"]=function(){return Module["asm"]["$"].apply(null,arguments)};var dynCall_vi=Module["dynCall_vi"]=function(){return Module["asm"]["aa"].apply(null,arguments)};var dynCall_vii=Module["dynCall_vii"]=function(){return Module["asm"]["ba"].apply(null,arguments)};var dynCall_viiii=Module["dynCall_viiii"]=function(){return Module["asm"]["ca"].apply(null,arguments)};Module["asm"]=asm;function ExitStatus(status){this.name="ExitStatus";this.message="Program terminated with exit("+status+")";this.status=status}ExitStatus.prototype=new Error;ExitStatus.prototype.constructor=ExitStatus;dependenciesFulfilled=function runCaller(){if(!Module["calledRun"])run();if(!Module["calledRun"])dependenciesFulfilled=runCaller};function run(args){args=args||Module["arguments"];if(runDependencies>0){return}preRun();if(runDependencies>0)return;if(Module["calledRun"])return;function doRun(){if(Module["calledRun"])return;Module["calledRun"]=true;if(ABORT)return;ensureInitRuntime();preMain();if(Module["onRuntimeInitialized"])Module["onRuntimeInitialized"]();postRun()}if(Module["setStatus"]){Module["setStatus"]("Running...");setTimeout(function(){setTimeout(function(){Module["setStatus"]("")},1);doRun()},1)}else{doRun()}}Module["run"]=run;function abort(what){if(Module["onAbort"]){Module["onAbort"](what)}if(what!==undefined){out(what);err(what);what=JSON.stringify(what)}else{what=""}ABORT=true;EXITSTATUS=1;throw"abort("+what+"). Build with -s ASSERTIONS=1 for more info."}Module["abort"]=abort;if(Module["preInit"]){if(typeof Module["preInit"]=="function")Module["preInit"]=[Module["preInit"]];while(Module["preInit"].length>0){Module["preInit"].pop()()}}Module["noExitRuntime"]=true;run();

    window['AscFonts'] = window['AscFonts'] || {};
    var AscFonts = window['AscFonts'];

    AscFonts.CreateLibrary = function()
    {
        return Module["_ASC_FT_Init"]();
    };

    AscFonts.TT_INTERPRETER_VERSION_35 = 35;
    AscFonts.TT_INTERPRETER_VERSION_38 = 38;
    AscFonts.TT_INTERPRETER_VERSION_40 = 40;

    AscFonts.FT_Set_TrueType_HintProp = function(library, tt_interpreter)
    {
        return Module["_ASC_FT_Set_TrueType_HintProp"](library, tt_interpreter);
    };

    AscFonts.CreateNativeStream = function(_typed_array)
    {
        var _fontStreamPointer = Module["_ASC_FT_Malloc"](_typed_array.size);
        Module["HEAP8"].set(_typed_array.data, _fontStreamPointer);
        return { asc_marker: true, data: _fontStreamPointer, len: _typed_array.size};
    };

    AscFonts.CreateNativeStreamByIndex = function(stream_index)
    {
        var _stream_pos = AscFonts.g_fonts_streams[stream_index];
        if (_stream_pos && true !== _stream_pos.asc_marker)
        {
            var _native_stream = AscFonts.CreateNativeStream(AscFonts.g_fonts_streams[stream_index]);
            AscFonts.g_fonts_streams[stream_index] = null;
            AscFonts.g_fonts_streams[stream_index] = _native_stream;
        }
    };

    function CFaceInfo()
    {
        this.units_per_EM = 0;
        this.ascender = 0;
        this.descender = 0;
        this.height = 0;
        this.face_flags = 0;
        this.num_faces = 0;
        this.num_glyphs = 0;
        this.num_charmaps = 0;
        this.style_flags = 0;
        this.face_index = 0;

        this.family_name = "";

        this.style_name = "";

        this.os2_version = 0;
        this.os2_usWeightClass = 0;
        this.os2_fsSelection = 0;
        this.os2_usWinAscent = 0;
        this.os2_usWinDescent = 0;
        this.os2_usDefaultChar = 0;
        this.os2_sTypoAscender = 0;
        this.os2_sTypoDescender = 0;
        this.os2_sTypoLineGap = 0;

        this.os2_ulUnicodeRange1 = 0;
        this.os2_ulUnicodeRange2 = 0;
        this.os2_ulUnicodeRange3 = 0;
        this.os2_ulUnicodeRange4 = 0;
        this.os2_ulCodePageRange1 = 0;
        this.os2_ulCodePageRange2 = 0;

        this.os2_nSymbolic = 0;

        this.header_yMin = 0;
        this.header_yMax = 0;

        this.monochromeSizes = [];
    };

    CFaceInfo.prototype.load = function(face)
    {
        var _bufferPtr = Module["_ASC_FT_GetFaceInfo"](face);
        if (!_bufferPtr)
            return;

        var _len_buffer = Math.min((Module["HEAP8"].length - _bufferPtr) >> 2, 250); //max 230 symbols on name & style
        var _buffer = new Int32Array(Module["HEAP8"].buffer, _bufferPtr, _len_buffer);
        var _index = 0;

        this.units_per_EM 	= Math.abs(_buffer[_index++]);
        this.ascender 		= _buffer[_index++];
        this.descender 		= _buffer[_index++];
        this.height 		= _buffer[_index++];
        this.face_flags 	= _buffer[_index++];
        this.num_faces 		= _buffer[_index++];
        this.num_glyphs 	= _buffer[_index++];
        this.num_charmaps 	= _buffer[_index++];
        this.style_flags 	= _buffer[_index++];
        this.face_index 	= _buffer[_index++];

        var c = _buffer[_index++];
        while (c)
        {
            this.family_name += String.fromCharCode(c);
            c = _buffer[_index++];
        }

        c = _buffer[_index++];
        while (c)
        {
            this.style_name += String.fromCharCode(c);
            c = _buffer[_index++];
        }

        this.os2_version 		= _buffer[_index++];
        this.os2_usWeightClass 	= _buffer[_index++];
        this.os2_fsSelection 	= _buffer[_index++];
        this.os2_usWinAscent 	= _buffer[_index++];
        this.os2_usWinDescent 	= _buffer[_index++];
        this.os2_usDefaultChar 	= _buffer[_index++];
        this.os2_sTypoAscender 	= _buffer[_index++];
        this.os2_sTypoDescender = _buffer[_index++];
        this.os2_sTypoLineGap 	= _buffer[_index++];

        this.os2_ulUnicodeRange1 	= AscFonts.FT_Common.IntToUInt(_buffer[_index++]);
        this.os2_ulUnicodeRange2 	= AscFonts.FT_Common.IntToUInt(_buffer[_index++]);
        this.os2_ulUnicodeRange3 	= AscFonts.FT_Common.IntToUInt(_buffer[_index++]);
        this.os2_ulUnicodeRange4 	= AscFonts.FT_Common.IntToUInt(_buffer[_index++]);
        this.os2_ulCodePageRange1 	= AscFonts.FT_Common.IntToUInt(_buffer[_index++]);
        this.os2_ulCodePageRange2 	= AscFonts.FT_Common.IntToUInt(_buffer[_index++]);

        this.os2_nSymbolic 			= _buffer[_index++];
        this.header_yMin 			= _buffer[_index++];
        this.header_yMax 			= _buffer[_index++];

        var fixedSizesCount = _buffer[_index++];
        for (var i = 0; i < fixedSizesCount; i++)
            this.monochromeSizes.push(_buffer[_index++]);

        Module["_ASC_FT_Free"](_bufferPtr);
    };

    function CGlyphMetrics()
    {
        this.bbox_xMin = 0;
        this.bbox_yMin = 0;
        this.bbox_xMax = 0;
        this.bbox_yMax = 0;

        this.width          = 0;
        this.height         = 0;

        this.horiAdvance    = 0;
        this.horiBearingX   = 0;
        this.horiBearingY   = 0;

        this.vertAdvance    = 0;
        this.vertBearingX   = 0;
        this.vertBearingY   = 0;

        this.linearHoriAdvance = 0;
        this.linearVertAdvance = 0;
    }

    function CGlyphBitmapImage()
    {
        this.left   = 0;
        this.top    = 0;
        this.width  = 0;
        this.rows   = 0;
        this.pitch  = 0;
        this.mode   = 0;
    }

    AscFonts.CFaceInfo = CFaceInfo;
    AscFonts.CGlyphMetrics = CGlyphMetrics;
    AscFonts.CGlyphBitmapImage = CGlyphBitmapImage;

    AscFonts.FT_Open_Face = function(library, stream, face_index)
    {
        return Module["_ASC_FT_Open_Face"](library, stream.data, stream.len, face_index);
    };

    AscFonts.FT_Glyph_Get_Measure = function(face, vector_worker, painter)
    {
        var _bufferPtr = Module["_ASC_FT_Get_Glyph_Measure_Params"](face, vector_worker ? 1 : 0);
        if (!_bufferPtr)
            return null;

        var _len = 15;
        if (vector_worker)
            _len = Module["HEAP32"][_bufferPtr >> 2];

        var _buffer = new Int32Array(Module["HEAP8"].buffer, _bufferPtr, 4 * _len);

        var _info = new CGlyphMetrics();
        _info.bbox_xMin     = _buffer[1];
        _info.bbox_yMin     = _buffer[2];
        _info.bbox_xMax     = _buffer[3];
        _info.bbox_yMax     = _buffer[4];

        _info.width         = _buffer[5];
        _info.height        = _buffer[6];

        _info.horiAdvance   = _buffer[7];
        _info.horiBearingX  = _buffer[8];
        _info.horiBearingY  = _buffer[9];

        _info.vertAdvance   = _buffer[10];
        _info.vertBearingX  = _buffer[11];
        _info.vertBearingY  = _buffer[12];

        _info.linearHoriAdvance     = _buffer[13];
        _info.linearVertAdvance     = _buffer[14];

        if (vector_worker)
        {
            painter.start(vector_worker);

            var _pos = 15;
            while (_pos < _len)
            {
                switch (_buffer[_pos++])
                {
                    case 0:
                    {
                        painter._move_to(_buffer[_pos++], _buffer[_pos++], vector_worker);
                        break;
                    }
                    case 1:
                    {
                        painter._line_to(_buffer[_pos++], _buffer[_pos++], vector_worker);
                        break;
                    }
                    case 2:
                    {
                        painter._conic_to(_buffer[_pos++], _buffer[_pos++], _buffer[_pos++], _buffer[_pos++], vector_worker);
                        break;
                    }
                    case 3:
                    {
                        painter._cubic_to(_buffer[_pos++], _buffer[_pos++], _buffer[_pos++], _buffer[_pos++], _buffer[_pos++], _buffer[_pos++], vector_worker);
                        break;
                    }
                    default:
                        break;
                }
            }

            painter.end(vector_worker);
        }

        Module["_ASC_FT_Free"](_bufferPtr);
        _buffer = null;

        return _info;
    };

    AscFonts.FT_Glyph_Get_Raster = function(face, render_mode)
    {
        var _bufferPtr = Module["_ASC_FT_Get_Glyph_Render_Params"](face, render_mode);
        if (!_bufferPtr)
            return null;

        var _buffer = new Int32Array(Module["HEAP8"].buffer, _bufferPtr, 24);

        var _info = new CGlyphBitmapImage();
        _info.left    = _buffer[0];
        _info.top     = _buffer[1];
        _info.width   = _buffer[2];
        _info.rows    = _buffer[3];
        _info.pitch   = _buffer[4];
        _info.mode    = _buffer[5];

        Module["_ASC_FT_Free"](_bufferPtr);
        return _info;
    };

    AscFonts.FT_Load_Glyph = Module["_FT_Load_Glyph"];
    AscFonts.FT_Set_Transform = Module["_ASC_FT_Set_Transform"];
    AscFonts.FT_Set_Char_Size = Module["_FT_Set_Char_Size"];

    AscFonts.FT_SetCMapForCharCode = Module["_ASC_FT_SetCMapForCharCode"];
    AscFonts.FT_GetKerningX = Module["_ASC_FT_GetKerningX"];
    AscFonts.FT_GetFaceMaxAdvanceX = Module["_ASC_FT_GetFaceMaxAdvanceX"];
    AscFonts.FT_Get_Glyph_Render_Buffer = function(face, rasterInfo, isCopyToRasterMemory)
    {
        var _bufferPtr = Module["_ASC_FT_Get_Glyph_Render_Buffer"](face);
        var tmp = new Uint8Array(Module["HEAP8"].buffer, _bufferPtr, rasterInfo.pitch * rasterInfo.rows);

        if (!isCopyToRasterMemory)
            return tmp;

        AscFonts.raster_memory.CheckSize(rasterInfo.width, rasterInfo.rows);

        var offsetSrc = 0;
        var offsetDst = 3;
        var dstData = AscFonts.raster_memory.m_oBuffer.data;

        if (rasterInfo.pitch >= rasterInfo.width)
		{
			for (var j = 0; j < rasterInfo.rows; ++j, offsetSrc += rasterInfo.pitch)
			{
				offsetDst = 3 + j * AscFonts.raster_memory.pitch;
				for (var i = 0; i < rasterInfo.width; i++, offsetDst += 4)
				{
					dstData[offsetDst] = tmp[offsetSrc + i];
				}
			}
		}
		else
        {
            var bitNumber = 0;
            var byteNumber = 0;
			for (var j = 0; j < rasterInfo.rows; ++j, offsetSrc += rasterInfo.pitch)
			{
				offsetDst = 3 + j * AscFonts.raster_memory.pitch;
				bitNumber = 0;
				byteNumber = 0;
				for (var i = 0; i < rasterInfo.width; i++, offsetDst += 4, bitNumber++)
				{
				    if (8 == bitNumber)
                    {
                        bitNumber = 0;
                        byteNumber++;
                    }
					dstData[offsetDst] = (tmp[offsetSrc + byteNumber] & (1 << (7 - bitNumber))) ? 255 : 0;
				}
			}
        }

        tmp = null;
    };

    AscFonts.onLoadModule();

})(window, undefined);
